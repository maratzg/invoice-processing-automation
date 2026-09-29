from __future__ import annotations

import json
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
from typing import Any

from .exceptions import ConfigurationError
from .file_utils import resolve_path


@dataclass(frozen=True, slots=True)
class MoneoTemplateConfig:
    template_sheet: str | None
    header_row: int | None
    data_start_row: int | None
    invoice_type: str
    invoice_type_value: str
    currency: str
    discount_percent: Decimal | None


@dataclass(frozen=True, slots=True)
class ParserConfig:
    preferred_engine: str = "pdfplumber"
    max_pdf_size_mb: int = 25
    min_extracted_text_chars: int = 50


@dataclass(frozen=True, slots=True)
class ValidationConfig:
    amount_tolerance: Decimal = Decimal("0.02")


@dataclass(frozen=True, slots=True)
class AppConfig:
    project_root: Path
    input_pdfs_dir: Path
    output_dir: Path
    template_file: Path
    mappings_dir: Path
    logs_dir: Path
    processed_log: Path
    run_log: Path
    moneo: MoneoTemplateConfig
    parser: ParserConfig
    validation: ValidationConfig


def _decimal_or_none(value: Any) -> Decimal | None:
    if value is None or value == "":
        return None
    try:
        return Decimal(str(value))
    except Exception as exc:
        raise ConfigurationError(f"Invalid decimal value in config: {value!r}") from exc


def load_config(config_path: str | Path = "config.json") -> AppConfig:
    path = Path(config_path)
    if not path.is_absolute():
        path = Path.cwd() / path
    path = path.resolve()
    root = path.parent

    try:
        with path.open("r", encoding="utf-8") as handle:
            raw = json.load(handle)
    except json.JSONDecodeError as exc:
        raise ConfigurationError(f"Invalid JSON in config file {path}: {exc}") from exc

    moneo_raw = raw.get("moneo", {})
    parser_raw = raw.get("parser", {})
    validation_raw = raw.get("validation", {})

    moneo = MoneoTemplateConfig(
        template_sheet=moneo_raw.get("template_sheet"),
        header_row=moneo_raw.get("header_row"),
        data_start_row=moneo_raw.get("data_start_row"),
        invoice_type=moneo_raw.get("invoice_type", "Rēķins"),
        invoice_type_value=moneo_raw.get("invoice_type_value", "REK"),
        currency=moneo_raw.get("currency", "EUR"),
        discount_percent=_decimal_or_none(moneo_raw.get("discount_percent")),
    )

    return AppConfig(
        project_root=root,
        input_pdfs_dir=resolve_path(root, raw.get("input_pdfs_dir", "input_pdfs")),
        output_dir=resolve_path(root, raw.get("output_dir", "output")),
        template_file=resolve_path(root, raw.get("template_file", "templates/Rekinu_Imp_2024.xlsx")),
        mappings_dir=resolve_path(root, raw.get("mappings_dir", "mappings")),
        logs_dir=resolve_path(root, raw.get("logs_dir", "logs")),
        processed_log=resolve_path(root, raw.get("processed_log", "logs/processed_log.xlsx")),
        run_log=resolve_path(root, raw.get("run_log", "logs/run.log")),
        moneo=moneo,
        parser=ParserConfig(
            preferred_engine=parser_raw.get("preferred_engine", "pdfplumber"),
            max_pdf_size_mb=int(parser_raw.get("max_pdf_size_mb", 25)),
            min_extracted_text_chars=int(parser_raw.get("min_extracted_text_chars", 50)),
        ),
        validation=ValidationConfig(
            amount_tolerance=Decimal(str(validation_raw.get("amount_tolerance", "0.02")))
        ),
    )


def ensure_runtime_directories(config: AppConfig) -> None:
    for directory in [
        config.input_pdfs_dir,
        config.output_dir,
        config.template_file.parent,
        config.mappings_dir,
        config.logs_dir,
    ]:
        directory.mkdir(parents=True, exist_ok=True)
