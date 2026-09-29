from __future__ import annotations

import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

from .exceptions import PdfExtractionError
from .invoice_models import ContainerInfo, Invoice, InvoiceHeader, InvoiceLine, Money


SERVICE_CODES = ("O.F.", "D/F", "THC/D", "ISPS/D", "MERCH")
CONTAINER_NUMBER_RE = re.compile(r"\b[A-Z]{4}\d{7}\b")
DATE_FORMATS = ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y")
KNOWN_VAT_RATES = {Decimal("0"), Decimal("5"), Decimal("12"), Decimal("21")}


def extract_text_from_pdf(
    pdf_path: Path,
    preferred_engine: str = "pdfplumber",
    min_text_chars: int = 50,
) -> str:
    """Extract text from a text-based PDF.

    This first backend intentionally targets text PDFs. Scanned/image-only PDFs
    should be routed to OCR later, not silently guessed here.
    """

    engines = [preferred_engine]
    if preferred_engine != "pdfplumber":
        engines.append("pdfplumber")
    if preferred_engine != "pymupdf":
        engines.append("pymupdf")

    attempts: list[str] = []
    for engine in engines:
        try:
            if engine.lower() == "pdfplumber":
                text = _extract_with_pdfplumber(pdf_path)
            elif engine.lower() in {"pymupdf", "fitz"}:
                text = _extract_with_pymupdf(pdf_path)
            else:
                attempts.append(f"{engine}: unsupported engine")
                continue
            if len(text.strip()) >= min_text_chars:
                return text
            attempts.append(f"{engine}: extracted only {len(text.strip())} text characters")
        except Exception as exc:  # pragma: no cover - depends on optional PDF libs/files
            attempts.append(f"{engine}: {type(exc).__name__}: {exc}")

    detail = "; ".join(attempts) or "no extraction engine was attempted"
    raise PdfExtractionError(
        f"Could not extract usable text from {pdf_path.name}. "
        f"Scanned or malformed PDFs require OCR/manual handling. Details: {detail}"
    )


def _extract_with_pdfplumber(pdf_path: Path) -> str:
    import pdfplumber

    page_text: list[str] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            page_text.append(page.extract_text() or "")
    return "\n".join(page_text)


def _extract_with_pymupdf(pdf_path: Path) -> str:
    import fitz

    text_parts: list[str] = []
    with fitz.open(str(pdf_path)) as doc:
        for page in doc:
            text_parts.append(page.get_text())
    return "\n".join(text_parts)


def parse_pdf(
    pdf_path: str | Path,
    preferred_engine: str = "pdfplumber",
    min_text_chars: int = 50,
    max_pdf_size_mb: int = 25,
) -> Invoice:
    path = Path(pdf_path)
    if path.stat().st_size > max_pdf_size_mb * 1024 * 1024:
        raise PdfExtractionError(f"PDF exceeds configured size limit of {max_pdf_size_mb} MB: {path.name}")
    text = extract_text_from_pdf(path, preferred_engine=preferred_engine, min_text_chars=min_text_chars)
    return parse_invoice_text(text, source_path=path)


def parse_invoice_text(text: str, source_path: str | Path) -> Invoice:
    path = Path(source_path)
    normalized_text = _normalize_text(text)
    lines = normalized_text.splitlines()
    shipping_table_values = _extract_shipping_table_values(lines)
    invoice_number, invoice_date = _extract_invoice_number_and_date(lines)
    client_vat, due_date = _extract_client_vat_and_due_date(lines)

    header = InvoiceHeader(
        invoice_number=_find_text(
            normalized_text,
            [
                r"(?im)^\s*(?:invoice|rēķins|rēķina)\s*(?:no\.?|nr\.?|number)?\s*[:#-]?\s*([A-Z0-9][A-Z0-9./_-]+)\s*$",
                r"(?im)^\s*(?:pavadzīmes\s*/\s*rēķina\s*nr\.?)\s*[:#-]?\s*([A-Z0-9][A-Z0-9./_-]+)\s*$",
            ],
        )
        or invoice_number,
        invoice_date=_find_date(
            normalized_text,
            [
                r"(?im)^\s*(?:invoice|rēķina)\s*(?:date|datums)\s*[:#-]?\s*([0-9./-]{8,10})\s*$",
            ],
        )
        or invoice_date,
        client_name=_find_text(
            normalized_text,
            [
                r"(?im)^\s*(?:client|customer|klients|bill to|buyer)\s*[:#-]\s*(.+?)\s*$",
                r"(?im)^\s*(?:klienta nosaukums)\s*[:#-]\s*(.+?)\s*$",
            ],
        ),
        client_vat_reg_number=_find_text(
            normalized_text,
            [
                r"(?im)^\s*(?:client\s*)?(?:vat|pvn|registration|reg\.?|reģ\.?)\s*(?:number|no\.?|nr\.?)?\s*[:#-]\s*([A-Z0-9 ._-]+)\s*$",
                r"(?im)^\s*(?:klienta\s*)?(?:reģ\.?\s*nr\.?)\s*[:#-]\s*([A-Z0-9 ._-]+)\s*$",
            ],
        )
        or client_vat,
        due_date=_find_date(
            normalized_text,
            [
                r"(?is)(?:due date|payment due|samaksas termiņš)\s*[:#-]?\s*([0-9./-]{8,10})",
            ],
        )
        or due_date,
        payment_terms=_find_text(
            normalized_text,
            [
                r"(?im)^\s*(?:payment terms|samaksas termiņš|terms)\s*[:#-]\s*(.+?)\s*$",
            ],
        )
        or shipping_table_values.get("payment_terms"),
        project_code=_find_text(
            normalized_text,
            [r"(?im)^\s*(?:project code|projekta kods)\s*[:#-]\s*(.+?)\s*$"],
        )
        or shipping_table_values.get("project_code"),
        bl_number=_find_text(
            normalized_text,
            [r"(?im)^\s*(?:b/l|bl|bill of lading)\s*(?:number|no\.?|nr\.?)?\s*[:#-]\s*(.+?)\s*$"],
        )
        or shipping_table_values.get("bl_number"),
        place_of_receipt=_find_text(
            normalized_text,
            [r"(?im)^\s*(?:place of receipt)\s*[:#-]\s*(.+?)\s*$"],
        )
        or shipping_table_values.get("place_of_receipt"),
        port_of_loading=_find_text(
            normalized_text,
            [r"(?im)^\s*(?:port of loading|pol)\s*[:#-]\s*(.+?)\s*$"],
        )
        or shipping_table_values.get("port_of_loading"),
        port_of_discharge=_find_text(
            normalized_text,
            [r"(?im)^\s*(?:port of discharge|pod)\s*[:#-]\s*(.+?)\s*$"],
        )
        or shipping_table_values.get("port_of_discharge"),
        place_of_delivery=_find_text(
            normalized_text,
            [r"(?im)^\s*(?:place of delivery)\s*[:#-]\s*(.+?)\s*$"],
        )
        or shipping_table_values.get("place_of_delivery"),
        total_amount_eur_incl_vat=_find_money(
            normalized_text,
            [
                r"(?is)(?:total\s*(?:amount)?\s*(?:eur)?\s*(?:including|incl\.?)\s*vat|total\s*incl\.?\s*vat|kopā\s*ar\s*pvn)\s*[:#-]?\s*(?:EUR)?\s*([-+]?[0-9][0-9 .,\u00a0]*)",
                r"(?im)^\s*(?:total|kopā)\s*[:#-]?\s*(?:EUR)?\s*([-+]?[0-9][0-9 .,\u00a0]*)\s*(?:EUR)?\s*$",
            ],
        ),
        subtotal=_find_money(
            normalized_text,
            [
                r"(?is)(?:subtotal|amount\s*(?:excl\.?|excluding)\s*vat|summa\s*bez\s*pvn)\s*[:#-]?\s*(?:EUR)?\s*([-+]?[0-9][0-9 .,\u00a0]*)",
            ],
        ),
        vat_amount=_find_money(
            normalized_text,
            [
                r"(?is)(?:vat\s*amount|pvn\s*summa|pvn)\s*[:#-]?\s*(?:EUR)?\s*([-+]?[0-9][0-9 .,\u00a0]*)",
            ],
        ),
        vat_clause=_find_text(
            normalized_text,
            [r"(?im)^\s*(?:vat clause|pvn atruna|pvn pants)\s*[:#-]\s*(.+?)\s*$"],
        ),
    )

    invoice = Invoice(
        source_path=path,
        header=header,
        lines=parse_invoice_lines(normalized_text),
        containers=parse_containers(normalized_text),
        raw_text=normalized_text,
    )

    if not invoice.lines:
        invoice.parser_warnings.append("No invoice service lines were detected by the generic parser.")
    if not invoice.containers:
        invoice.parser_warnings.append("No container rows were detected by the generic parser.")
    return invoice


def parse_invoice_lines(text: str) -> list[InvoiceLine]:
    lines: list[InvoiceLine] = []
    # Service codes contain punctuation such as "." and "/", so word-boundary
    # matching is not reliable at the end of codes like "O.F.".
    service_pattern = re.compile(rf"(?<!\w)({'|'.join(re.escape(code) for code in SERVICE_CODES)})(?!\w)", re.IGNORECASE)
    for raw_line in text.splitlines():
        line = raw_line.strip()
        service_match = service_pattern.search(line)
        if not service_match:
            continue
        parsed = _parse_service_line(line, service_match.group(1).upper())
        if parsed is not None:
            lines.append(parsed)
    return lines


def parse_containers(text: str) -> list[ContainerInfo]:
    containers: dict[str, ContainerInfo] = {}
    for raw_line in text.splitlines():
        line = raw_line.strip()
        matches = list(CONTAINER_NUMBER_RE.finditer(line))
        for match in matches:
            number = match.group(0)
            after = line[match.end() :]
            slash_parts = [part.strip() for part in after.split("/") if part.strip()]
            unit_type = slash_parts[0].upper() if slash_parts else _find_unit_type(after)
            if len(slash_parts) >= 3:
                weight = parse_money(slash_parts[1])
                cbm = parse_money(slash_parts[2])
                containers[number] = ContainerInfo(number, unit_type=unit_type, weight=weight, cbm=cbm)
                continue
            numeric_values = _numbers_from_text(after)
            weight = numeric_values[0] if numeric_values else None
            cbm = numeric_values[1] if len(numeric_values) > 1 else None
            containers[number] = ContainerInfo(number, unit_type=unit_type, weight=weight, cbm=cbm)
    return list(containers.values())


def parse_date(value: str | None) -> date | None:
    if not value:
        return None
    cleaned = value.strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(cleaned, fmt).date()
        except ValueError:
            continue
    return None


def parse_money(value: str | int | float | Decimal | None) -> Money | None:
    if value is None or value == "":
        return None
    if isinstance(value, Decimal):
        return value

    text = str(value).strip()
    text = text.replace("\u00a0", " ").replace(" ", "")
    text = re.sub(r"(?i)(EUR|USD|GBP|PVN|VAT|%)", "", text)
    text = re.sub(r"[^0-9,.\-+]", "", text)
    if not text:
        return None

    if "," in text and "." in text:
        if text.rfind(",") > text.rfind("."):
            text = text.replace(".", "").replace(",", ".")
        else:
            text = text.replace(",", "")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def _parse_service_line(line: str, service_code: str) -> InvoiceLine | None:
    vat_rate = _find_vat_rate(line)
    numbers, description = _numeric_tail_from_service_line(line, service_code)
    if not numbers:
        return None

    currency_match = re.search(r"\b(EUR|USD|GBP)\b", line, flags=re.IGNORECASE)
    currency = currency_match.group(1).upper() if currency_match else "EUR"
    amount_ex_vat = numbers[-1]
    if vat_rate is None and len(numbers) >= 2 and numbers[-2] in KNOWN_VAT_RATES:
        vat_rate = numbers[-2]
    unit_price = numbers[-3] if len(numbers) >= 3 and numbers[-2] in KNOWN_VAT_RATES else numbers[-2] if len(numbers) >= 2 else amount_ex_vat
    quantity = numbers[0] if len(numbers) >= 3 else Decimal("1")
    exchange_rate = numbers[1] if len(numbers) >= 6 else None
    price_in_currency = numbers[2] if len(numbers) >= 6 else None

    return InvoiceLine(
        service_code=service_code,
        description=description or service_code,
        quantity=quantity,
        pricing_currency=currency,
        exchange_rate=exchange_rate,
        price_in_currency=price_in_currency,
        currency_code="EUR",
        unit_price_ex_vat=unit_price,
        vat_rate=vat_rate,
        line_amount_ex_vat=amount_ex_vat,
    )


def _numeric_tail_from_service_line(line: str, service_code: str) -> tuple[list[Money], str]:
    tokens = line.split()
    numeric_tokens_reversed: list[str] = []
    tail_start = len(tokens)
    for index in range(len(tokens) - 1, -1, -1):
        token = tokens[index].rstrip("%")
        if parse_money(token) is None:
            break
        numeric_tokens_reversed.append(token)
        tail_start = index
    numeric_tokens = list(reversed(numeric_tokens_reversed))
    numbers = [value for token in numeric_tokens if (value := parse_money(token)) is not None]

    service_index = 0
    for index, token in enumerate(tokens):
        if token.upper() == service_code.upper():
            service_index = index
            break
    description_tokens = tokens[service_index + 1 : tail_start]
    description = " ".join(description_tokens).strip(" -:\t") or service_code
    return numbers, description


def _numbers_from_text(text: str) -> list[Money]:
    values: list[Money] = []
    for match in re.finditer(r"[-+]?\d{1,3}(?:[ .]\d{3})*(?:[,.]\d+)?|[-+]?\d+(?:[,.]\d+)?", text):
        value = parse_money(match.group(0))
        if value is not None:
            values.append(value)
    return values


def _find_vat_rate(text: str) -> Money | None:
    match = re.search(r"\b(0|5|12|21)(?:[,.]0+)?\s*%", text)
    if match:
        return Decimal(match.group(1))
    return None


def _extract_invoice_number_and_date(lines: list[str]) -> tuple[str | None, date | None]:
    for index, line in enumerate(lines):
        match = re.match(r"(?i)^invoice\s+([A-Z0-9][A-Z0-9./_-]+)$", line.strip())
        if not match:
            continue
        invoice_number = match.group(1)
        invoice_date = parse_date(lines[index + 1]) if index + 1 < len(lines) else None
        return invoice_number, invoice_date
    return None, None


def _extract_client_vat_and_due_date(lines: list[str]) -> tuple[str | None, date | None]:
    for index, line in enumerate(lines):
        normalized = line.lower()
        if "vat registration no" not in normalized or "due date" not in normalized:
            continue
        if index + 1 >= len(lines):
            return None, None
        next_line = lines[index + 1]
        vat_match = re.search(r"\b[A-Z]{2}\d{8,14}\b", next_line)
        dates = [parse_date(match.group(0)) for match in re.finditer(r"\b\d{2}[./-]\d{2}[./-]\d{4}\b", next_line)]
        dates = [value for value in dates if value is not None]
        return (vat_match.group(0) if vat_match else None), (dates[-1] if dates else None)
    return None, None


def _extract_shipping_table_values(lines: list[str]) -> dict[str, str]:
    values: dict[str, str] = {}
    for index, line in enumerate(lines):
        normalized = line.lower()
        if "project code" in normalized and "b/l" in normalized and "payment terms" in normalized:
            if index + 1 < len(lines):
                parts = lines[index + 1].split()
                if len(parts) >= 3:
                    values["project_code"] = parts[0]
                    values["bl_number"] = parts[1]
                    values["payment_terms"] = " ".join(parts[2:])
        if (
            "place of receipt" in normalized
            and "port of loading" in normalized
            and "port of discharge" in normalized
            and "place of delivery" in normalized
        ):
            if index + 1 < len(lines):
                parts = lines[index + 1].split()
                if len(parts) >= 4:
                    values["place_of_receipt"] = parts[0]
                    values["port_of_loading"] = parts[1]
                    values["port_of_discharge"] = parts[2]
                    values["place_of_delivery"] = " ".join(parts[3:])
    return values


def _find_unit_type(text: str) -> str | None:
    match = re.search(r"\b(20DC|20DV|40DC|40DV|40HC|45HC|20'|40'|LCL|FCL)\b", text, flags=re.IGNORECASE)
    return match.group(1).upper() if match else None


def _find_text(text: str, patterns: list[str]) -> str | None:
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            value = " ".join(match.group(1).split())
            return value or None
    return None


def _find_date(text: str, patterns: list[str]) -> date | None:
    value = _find_text(text, patterns)
    return parse_date(value)


def _find_money(text: str, patterns: list[str]) -> Money | None:
    value = _find_text(text, patterns)
    return parse_money(value)


def _normalize_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = text.replace("\u00a0", " ")
    return "\n".join(line.strip() for line in text.splitlines() if line.strip())
