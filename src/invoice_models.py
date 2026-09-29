from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
from typing import Any


Money = Decimal


@dataclass(slots=True)
class InvoiceHeader:
    invoice_number: str | None = None
    invoice_date: date | None = None
    client_name: str | None = None
    client_vat_reg_number: str | None = None
    due_date: date | None = None
    payment_terms: str | None = None
    project_code: str | None = None
    bl_number: str | None = None
    place_of_receipt: str | None = None
    port_of_loading: str | None = None
    port_of_discharge: str | None = None
    place_of_delivery: str | None = None
    total_amount_eur_incl_vat: Money | None = None
    subtotal: Money | None = None
    vat_amount: Money | None = None
    vat_clause: str | None = None


@dataclass(slots=True)
class InvoiceLine:
    service_code: str | None = None
    description: str | None = None
    quantity: Money | None = None
    pricing_currency: str | None = None
    exchange_rate: Money | None = None
    price_in_currency: Money | None = None
    currency_code: str | None = None
    unit_price_ex_vat: Money | None = None
    vat_rate: Money | None = None
    line_amount_ex_vat: Money | None = None


@dataclass(slots=True)
class ContainerInfo:
    container_number: str
    unit_type: str | None = None
    weight: Money | None = None
    cbm: Money | None = None


@dataclass(slots=True)
class Invoice:
    source_path: Path
    header: InvoiceHeader = field(default_factory=InvoiceHeader)
    lines: list[InvoiceLine] = field(default_factory=list)
    containers: list[ContainerInfo] = field(default_factory=list)
    raw_text: str = ""
    parser_warnings: list[str] = field(default_factory=list)

    @property
    def source_filename(self) -> str:
        return self.source_path.name

    @property
    def container_count(self) -> int:
        unique_numbers = {container.container_number for container in self.containers}
        return len(unique_numbers)


@dataclass(slots=True)
class ClientRecord:
    pdf_client_name: str
    pdf_vat_reg_number: str | None
    moneo_client_name: str
    moneo_client_code: str | None
    country: str | None
    client_type: str | None = None


@dataclass(slots=True)
class VATRecord:
    country: str | None
    client_type: str | None
    vat_rate: Money
    pvn_likme: str | None
    pvn_kods: str | None
    manual_review_required: bool = False


@dataclass(slots=True)
class ArticleRecord:
    vat_rate: Money
    artikula_nosaukums: str
    artikula_kods: str | None = None


@dataclass(slots=True)
class MoneoRow:
    invoice_date: date
    client_name: str
    client_reg_number: str
    invoice_type: str
    invoice_type_value: str
    currency: str
    payment_terms: str
    invoice_number: str
    article_code: str | None
    article_name: str
    unit_price_ex_vat: Money
    quantity: int
    discount_percent: Money | None
    line_sum_ex_vat: Money
    pvn_likme: str
    pvn_kods: str
    source_pdf: str
    vat_rate: Money

    def to_template_values(self) -> dict[str, Any]:
        return {
            "Rēķina datums": self.invoice_date,
            "Klienta nosaukums": self.client_name,
            "Klienta Reģ.nr.": self.client_reg_number,
            "Rēķina veids": self.invoice_type,
            "Rēķina veids value": self.invoice_type_value,
            "Valūta": self.currency,
            "Samaksas termiņš": self.payment_terms,
            "Pavadzīmes / Rēķina nr": self.invoice_number,
            "Artikula kods": self.article_code,
            "Artikula nosaukums": self.article_name,
            "Cena BEZ PVN": self.unit_price_ex_vat,
            "Daudzums": self.quantity,
            "Atlaides %": self.discount_percent,
            "Rindas summa BEZ PVN": self.line_sum_ex_vat,
            "PVN likme": self.pvn_likme,
            "PVN kods": self.pvn_kods,
        }


@dataclass(slots=True)
class ValidationOutcome:
    source_pdf: str
    invoice_number: str | None
    valid: bool
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    client: ClientRecord | None = None
    vat_records_by_rate: dict[Money, VATRecord] = field(default_factory=dict)
    article_records_by_rate: dict[Money, ArticleRecord] = field(default_factory=dict)


def iso_datetime_now() -> str:
    return datetime.now().replace(microsecond=0).isoformat(sep=" ")

