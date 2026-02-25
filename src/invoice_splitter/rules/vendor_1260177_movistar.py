from __future__ import annotations

from decimal import Decimal
from typing import List

from invoice_splitter.models import InvoiceInput, LineItem
from invoice_splitter.rules.common import calc_iva_and_total, q2, validate_and_compute_allocations

MOVISTAR_TABLE = "Movistar_table"
DEFAULT_CONCEPT = "10 lines DRP (4 lines 35 GB + 6 lines 53 GB)"

# Standard split
CC1 = 7475036
GL1 = 4649000000
PCT1 = Decimal("0.60")

CC2 = 3941036
GL2 = 3649000000
PCT2 = Decimal("0.40")

DEFAULT_LINES = 10


VENDOR_ID = 1260177


def build_lines_for_vendor(invoice: InvoiceInput) -> List[LineItem]:
    return build_lines_for_movistar(invoice)


def build_lines_for_movistar(invoice: InvoiceInput) -> List[LineItem]:
    """
    MOVISTAR (vendor_id=1260177) -> Movistar_table

    - Default concept -> split estándar 60/40 con CC/GL fijos
    - Phone lines quantity: default 10, editable por usuario (invoice.extras['phone_lines_qty'])
    - Concepto custom:
        - split custom: N líneas
        - sin split: 1 línea con CC/GL del usuario
    """
    concept = (invoice.service_concept or "").strip() or DEFAULT_CONCEPT
    iva_rate = invoice.iva_rate

    phone_lines = invoice.extras.get("phone_lines_qty", DEFAULT_LINES)
    try:
        phone_lines = int(phone_lines)
    except Exception:
        raise ValueError("Phone lines quantity debe ser un número entero.")

    # Default concept -> standard split
    if concept == DEFAULT_CONCEPT:
        part1 = q2(invoice.subtotal * PCT1)
        part2 = q2(invoice.subtotal * PCT2)
        return [
            _make_line(invoice, concept, CC1, GL1, part1, iva_rate, phone_lines),
            _make_line(invoice, concept, CC2, GL2, part2, iva_rate, phone_lines),
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
                phone_lines,
            )
            for alloc, amount in pairs
        ]

    # Custom concept without split
    cc = invoice.extras.get("cc")
    gl = invoice.extras.get("gl_account")
    if cc is None or gl is None:
        raise ValueError("Para concepto personalizado en MOVISTAR debes ingresar CC y GL account.")
    return [_make_line(invoice, concept, int(cc), int(gl), invoice.subtotal, iva_rate, phone_lines)]


def _make_line(
    invoice: InvoiceInput,
    concept: str,
    cc: int,
    gl: int,
    subtotal_assigned,
    iva_rate,
    phone_lines_qty: int,
) -> LineItem:
    iva, total = calc_iva_and_total(subtotal_assigned, iva_rate)
    return LineItem(
        table_name=MOVISTAR_TABLE,
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
            "Phone lines quantity": phone_lines_qty,
        },
    )
