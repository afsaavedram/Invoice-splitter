from __future__ import annotations

from typing import List

from invoice_splitter.models import InvoiceInput, LineItem
from invoice_splitter.rules.common import calc_iva_and_total, validate_and_compute_allocations

PUNTONET_TABLE = "Puntonet_table"
DEFAULT_CONCEPT = "40 MBPS"

DEFAULT_CC = 1100036
DEFAULT_GL = 7418000000


def build_lines_for_puntonet(invoice: InvoiceInput) -> List[LineItem]:
    """
    PUNTONET (vendor_id=1261182) -> Puntonet_table

    - Concepto default: "40 MBPS" -> 1 línea CC/GL default.
    - Concepto custom (OTRO):
        - Si hay split custom (alloc_mode + allocations): N líneas (CC/GL por línea).
        - Si no hay split: 1 línea 100% con CC/GL del usuario (invoice.extras['cc'], ['gl_account']).
    """
    concept = (invoice.service_concept or "").strip() or DEFAULT_CONCEPT
    iva_rate = invoice.iva_rate

    # Default concept
    if concept == DEFAULT_CONCEPT:
        return [_make_line(invoice, concept, DEFAULT_CC, DEFAULT_GL, invoice.subtotal, iva_rate)]

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
            )
            for alloc, amount in pairs
        ]

    # Custom concept without split
    cc = invoice.extras.get("cc")
    gl = invoice.extras.get("gl_account")
    if cc is None or gl is None:
        raise ValueError("Para concepto personalizado en PUNTONET debes ingresar CC y GL account.")
    return [_make_line(invoice, concept, int(cc), int(gl), invoice.subtotal, iva_rate)]


def _make_line(
    invoice: InvoiceInput,
    concept: str,
    cc: int,
    gl: int,
    subtotal_assigned,
    iva_rate,
) -> LineItem:
    iva, total = calc_iva_and_total(subtotal_assigned, iva_rate)
    return LineItem(
        table_name=PUNTONET_TABLE,
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
        },
    )
