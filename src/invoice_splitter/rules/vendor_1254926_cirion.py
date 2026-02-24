from __future__ import annotations

from decimal import Decimal
from typing import List

from invoice_splitter.models import InvoiceInput, LineItem
from invoice_splitter.rules.common import calc_iva_and_total, q2, validate_and_compute_allocations

CIRION_TABLE = "Cirion_table"
DEFAULT_CONCEPT = "Internet"

# Standard split
CC1 = 7475036
GL1 = 7418000000
PCT1 = Decimal("0.60")

CC2 = 3941036
GL2 = 3526400000
PCT2 = Decimal("0.40")

DEFAULT_BANDWIDTH = 40  # MBPS


def build_lines_for_cirion(invoice: InvoiceInput) -> List[LineItem]:
    """
    CIRION (vendor_id=1254926) -> Cirion_table

    - Default concept: "Internet" -> split estándar 60/40 con CC/GL fijos
    - Bandwidth (MBPS): default 40, editable por usuario (invoice.extras['bandwidth_mbps'])
    - Concepto custom (OTRO):
        - Si hay split custom: N líneas (CC/GL por línea) + bandwidth en todas
        - Si no hay split: 1 línea 100% con CC/GL del usuario + bandwidth
    """
    concept = (invoice.service_concept or "").strip() or DEFAULT_CONCEPT
    iva_rate = invoice.iva_rate

    bandwidth = invoice.extras.get("bandwidth_mbps", DEFAULT_BANDWIDTH)
    try:
        bandwidth = int(bandwidth)
    except Exception:
        raise ValueError("Bandwidth (MBPS) debe ser un número entero.")

    # Default concept -> standard split
    if concept == DEFAULT_CONCEPT:
        part1 = q2(invoice.subtotal * PCT1)
        part2 = q2(invoice.subtotal * PCT2)
        return [
            _make_line(invoice, concept, CC1, GL1, part1, iva_rate, bandwidth),
            _make_line(invoice, concept, CC2, GL2, part2, iva_rate, bandwidth),
        ]

    # Custom concept with split
    if invoice.alloc_mode and invoice.allocations:
        pairs = validate_and_compute_allocations(
            invoice.subtotal, invoice.alloc_mode, invoice.allocations
        )
        return [
            _make_line(
                invoice,
                (alloc.concept or concept).strip(),
                alloc.cc,
                alloc.gl_account,
                amount,
                iva_rate,
                bandwidth,
            )
            for alloc, amount in pairs
        ]

    # Custom concept without split
    cc = invoice.extras.get("cc")
    gl = invoice.extras.get("gl_account")
    if cc is None or gl is None:
        raise ValueError("Para concepto personalizado en CIRION debes ingresar CC y GL account.")
    return [_make_line(invoice, concept, int(cc), int(gl), invoice.subtotal, iva_rate, bandwidth)]


def _make_line(
    invoice: InvoiceInput,
    concept: str,
    cc: int,
    gl: int,
    subtotal_assigned,
    iva_rate,
    bandwidth_mbps: int,
) -> LineItem:
    iva, total = calc_iva_and_total(subtotal_assigned, iva_rate)
    return LineItem(
        table_name=CIRION_TABLE,
        values={
            "Date": invoice.invoice_date,
            "Bill number": invoice.bill_number,
            "ID": invoice.vendor_id,
            "Vendor": invoice.vendor_name,
            "Service/ concept": concept,
            "CC": int(cc),
            "GL account": int(gl),
            "Subtotal assigned by CC": subtotal_assigned,
            "% IVA": iva_rate,
            "IVA assigned by CC": iva,
            "Total assigned by CC": total,
            "Bandwidth (MBPS)": bandwidth_mbps,
        },
    )
