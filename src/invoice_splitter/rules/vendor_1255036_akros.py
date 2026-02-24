from __future__ import annotations

from decimal import Decimal
from typing import List

from invoice_splitter.models import InvoiceInput, LineItem
from invoice_splitter.rules.common import (
    calc_iva_and_total,
    q2,
    validate_and_compute_allocations,
)

AKROS_TABLE = "Akros_bills_table"
DEFAULT_CONCEPT = "Printers & Copiers"

# Regla estándar:
# - 2 líneas
# - CC 3941036 -> 40% subtotal, GL 3526200000
# - CC 7475036 -> 60% subtotal, GL 7427000000
CC1 = 3941036
GL1 = 3526200000
PCT1 = Decimal("0.40")

CC2 = 7475036
GL2 = 7427000000
PCT2 = Decimal("0.60")


def build_lines_for_akros(invoice: InvoiceInput) -> List[LineItem]:
    """
    AKROS (vendor_id=1255036) -> Akros_bills_table

    Reglas:
    1) Concepto default ("Printers & Copiers"):
       - Split estándar 40/60:
         * CC 3941036 / GL 3526200000 -> 40%
         * CC 7475036 / GL 7427000000 -> 60%

    2) Concepto custom (concept != default):
       a) Si hay split custom (invoice.alloc_mode + invoice.allocations):
          - Genera N líneas con CC/GL por allocation
          - Split puede ser por porcentaje o por valor
          - Valida suma vs subtotal con tolerancia ±0.01 y ajusta última línea si aplica
          - Cada allocation puede traer concept propio; si no, usa el concepto general

       b) Si NO hay split custom:
          - Requiere CC y GL del usuario vía invoice.extras['cc'], invoice.extras['gl_account']
          - 1 línea 100%
    """

    concept = (invoice.service_concept or "").strip() or DEFAULT_CONCEPT
    iva_rate = invoice.iva_rate

    # --- Caso 1: default concept -> split estándar ---
    if concept == DEFAULT_CONCEPT:
        part1 = q2(invoice.subtotal * PCT1)
        part2 = q2(invoice.subtotal * PCT2)
        return [
            _make_line(invoice, concept, CC1, GL1, part1, iva_rate),
            _make_line(invoice, concept, CC2, GL2, part2, iva_rate),
        ]

    # --- Caso 2: concept custom con split configurado ---
    if invoice.alloc_mode and invoice.allocations:
        pairs = validate_and_compute_allocations(
            invoice.subtotal, invoice.alloc_mode, invoice.allocations
        )

        lines: List[LineItem] = []
        for alloc, amount in pairs:
            line_concept = (alloc.concept or concept).strip()
            lines.append(
                _make_line(invoice, line_concept, alloc.cc, alloc.gl_account, amount, iva_rate)
            )
        return lines

    # --- Caso 3: concept custom sin split -> 1 línea 100% a CC/GL del usuario ---
    cc = invoice.extras.get("cc")
    gl = invoice.extras.get("gl_account")
    if cc is None or gl is None:
        raise ValueError("Para concepto personalizado en AKROS debes ingresar CC y GL account.")

    return [_make_line(invoice, concept, int(cc), int(gl), invoice.subtotal, iva_rate)]


def _make_line(
    invoice: InvoiceInput,
    concept: str,
    cc: int,
    gl: int,
    subtotal_assigned: Decimal,
    iva_rate: Decimal,
) -> LineItem:
    """
    Construye una fila para Akros_bills_table con IVA/Total calculados.
    cc y gl numéricos (int), bill number texto (9 dígitos), fecha date.
    """
    iva, total = calc_iva_and_total(subtotal_assigned, iva_rate)

    return LineItem(
        table_name=AKROS_TABLE,
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
