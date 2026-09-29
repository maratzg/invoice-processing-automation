from __future__ import annotations

import argparse
import json
import sys
from dataclasses import asdict
from pathlib import Path
from time import perf_counter

if __package__ in {None, ""}:  # Allows: python src/main.py
    sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src.config import AppConfig, ensure_runtime_directories, load_config
from src.exceptions import InvoiceAutomationError
from src.file_utils import validate_pdf_input_path
from src.invoice_models import Invoice, ValidationOutcome
from src.logging_setup import configure_logging, get_logger
from src.mapping_loader import create_mapping_templates, load_mappings
from src.moneo_writer import (
    build_moneo_rows,
    update_processed_log,
    write_error_report,
    write_extracted_data,
    write_moneo_import,
    write_no_import_workbook,
    write_validation_report,
)
from src.pdf_parser import parse_pdf
from src.validator import ensure_processed_log, load_processed_invoice_numbers, validate_invoice


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Extract PDF invoices and create MONEO-compatible Excel import files.")
    parser.add_argument("--config", default="config.json", help="Path to config.json.")
    parser.add_argument("--single-pdf", help="Process one PDF file instead of all files in input_pdfs/.")
    parser.add_argument("--init-mappings", action="store_true", help="Create editable mapping workbooks if missing.")
    parser.add_argument("--debug-json", action="store_true", help="Print parsed invoice objects as JSON-friendly dictionaries.")
    args = parser.parse_args(argv)

    try:
        config = load_config(args.config)
        ensure_runtime_directories(config)
        logger = configure_logging(config.run_log)
        logger.info("Starting MONEO invoice automation CLI")
    except InvoiceAutomationError as exc:
        print(f"Configuration error: {exc}")
        return 1
    except Exception as exc:
        print(f"Unexpected startup error: {type(exc).__name__}: {exc}")
        return 1

    try:
        if args.init_mappings:
            created = create_mapping_templates(config.mappings_dir, overwrite=False)
            ensure_processed_log(config.processed_log)
            print("Mapping/log templates are ready:")
            for path in created:
                print(f"- {path}")
            logger.info("Mapping/log templates initialized")
            return 0

        return run_pipeline(config, single_pdf=args.single_pdf, debug_json=args.debug_json)
    except InvoiceAutomationError as exc:
        get_logger().exception("Run failed with user-actionable error: %s", exc)
        print(f"Error: {exc}")
        return 1
    except Exception as exc:
        get_logger().exception("Run failed with unexpected error: %s", exc)
        print(f"Unexpected error: {type(exc).__name__}: {exc}")
        return 1


def run_pipeline(config: AppConfig, single_pdf: str | None = None, debug_json: bool = False) -> int:
    logger = get_logger()
    started = perf_counter()
    pdf_paths = _collect_pdf_paths(config, single_pdf)
    invoices: list[Invoice] = []
    outcomes: list[ValidationOutcome] = []
    ready_rows = []

    mappings = load_mappings(config.mappings_dir)
    processed_invoice_numbers = load_processed_invoice_numbers(config.processed_log)
    template_available = config.template_file.exists()
    batch_invoice_numbers_seen: set[str] = set()
    logger.info(
        "Run started | pdf_count=%s | input_dir=%s | template_available=%s",
        len(pdf_paths),
        config.input_pdfs_dir,
        template_available,
    )

    for pdf_path in pdf_paths:
        invoice_started = perf_counter()
        logger.info("Processing PDF | file=%s | size_bytes=%s", pdf_path.name, pdf_path.stat().st_size)
        try:
            invoice = parse_pdf(
                pdf_path,
                preferred_engine=config.parser.preferred_engine,
                min_text_chars=config.parser.min_extracted_text_chars,
                max_pdf_size_mb=config.parser.max_pdf_size_mb,
            )
            invoices.append(invoice)
            outcome = validate_invoice(
                invoice,
                mappings=mappings,
                processed_invoice_numbers=processed_invoice_numbers,
                amount_tolerance=config.validation.amount_tolerance,
                template_available=template_available,
                batch_invoice_numbers_seen=batch_invoice_numbers_seen,
            )
            if outcome.valid:
                try:
                    moneo_rows, rounding_warnings = build_moneo_rows(invoice, outcome, config.moneo)
                    outcome.warnings.extend(rounding_warnings)
                    ready_rows.extend(moneo_rows)
                except InvoiceAutomationError as exc:
                    outcome.valid = False
                    outcome.errors.append(f"MONEO conversion error: {exc}")
                    logger.exception("MONEO conversion failed | file=%s | invoice=%s", pdf_path.name, outcome.invoice_number)
            if invoice.header.invoice_number:
                batch_invoice_numbers_seen.add(invoice.header.invoice_number)
            outcomes.append(outcome)
            logger.info(
                "Processed PDF | file=%s | invoice=%s | status=%s | errors=%s | warnings=%s | seconds=%.3f",
                pdf_path.name,
                outcome.invoice_number,
                "READY" if outcome.valid else "ERROR",
                len(outcome.errors),
                len(outcome.warnings),
                perf_counter() - invoice_started,
            )
        except Exception as exc:
            logger.exception("Parser/runtime error | file=%s", pdf_path.name)
            outcomes.append(
                ValidationOutcome(
                    source_pdf=pdf_path.name,
                    invoice_number=None,
                    valid=False,
                    errors=[f"Parser/runtime error: {exc}"],
                )
            )

    if debug_json:
        print(json.dumps([_invoice_debug_dict(invoice) for invoice in invoices], default=str, ensure_ascii=False, indent=2))

    _write_outputs(config, invoices, outcomes, ready_rows)
    logger.info(
        "Run finished | processed=%s | ready=%s | errors=%s | seconds=%.3f",
        len(pdf_paths),
        sum(1 for outcome in outcomes if outcome.valid),
        sum(1 for outcome in outcomes if not outcome.valid),
        perf_counter() - started,
    )
    _print_summary(config, pdf_paths, outcomes)
    return 0 if not any(not outcome.valid for outcome in outcomes) else 2


def _collect_pdf_paths(config: AppConfig, single_pdf: str | None) -> list[Path]:
    if single_pdf:
        path = Path(single_pdf)
        if not path.is_absolute():
            path = config.project_root / path
        return [validate_pdf_input_path(path)]
    return [validate_pdf_input_path(path) for path in sorted(config.input_pdfs_dir.glob("*.pdf"))]


def _write_outputs(config: AppConfig, invoices: list[Invoice], outcomes: list[ValidationOutcome], ready_rows) -> None:
    output_dir = config.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    write_extracted_data(output_dir / "extracted_data.xlsx", invoices)
    write_validation_report(output_dir / "validation_report.xlsx", outcomes)
    write_error_report(output_dir / "error_report.xlsx", outcomes)

    if ready_rows:
        write_moneo_import(config.template_file, output_dir / "moneo_import_ready.xlsx", ready_rows, config.moneo)
    else:
        reason = "No invoices passed validation."
        if not config.template_file.exists():
            reason += " The official MONEO template is also missing."
        write_no_import_workbook(output_dir / "moneo_import_ready.xlsx", reason)

    update_processed_log(config.processed_log, invoices, outcomes)


def _print_summary(config: AppConfig, pdf_paths: list[Path], outcomes: list[ValidationOutcome]) -> None:
    ready_count = sum(1 for outcome in outcomes if outcome.valid)
    error_count = sum(1 for outcome in outcomes if not outcome.valid)
    print(f"Processed {len(pdf_paths)} PDFs")
    print(f"Ready for import: {ready_count}")
    print(f"Errors: {error_count}")
    print(f"Output written to: {config.output_dir}")


def _invoice_debug_dict(invoice: Invoice) -> dict[str, object]:
    return {
        "source_pdf": invoice.source_filename,
        "header": asdict(invoice.header),
        "lines": [asdict(line) for line in invoice.lines],
        "containers": [asdict(container) for container in invoice.containers],
        "parser_warnings": invoice.parser_warnings,
    }


if __name__ == "__main__":
    raise SystemExit(main())
