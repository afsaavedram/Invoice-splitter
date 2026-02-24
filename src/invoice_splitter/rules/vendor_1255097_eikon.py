from __future__ import annotations

from decimal import Decimal
from typing import List

from invoice_splitter.models import InvoiceInput, LineItem
from invoice_splitter.rules.common import calc_iva_and_total, q2, validate_and_compute_allocations


EIKON_TABLE = "Eikon_table"
GL_DEFAULT = 7980100000

CONCEPTS = {
    "Infrastructure cloud (Monthly)",
    "Azure Consumptions (biannual)",
    "Maintenance and support (annual)",
    "Domains (annual)",
}


def build_lines_for_eikon(invoice: InvoiceInput) -> List[LineItem]:
    """
    Reglas (Modo 1):
    - Conceptos permitidos (o custom).
    - GL = 7980100000 para los conceptos estándar.
    - Si concept = Infrastructure cloud (Monthly):
        2 líneas: CC 7457036 (60%), CC 7475036 (40%)  <-- OJO: aquí usamos tu regla actual.
      (Si más adelante confirmas otra estructura, se cambia aquí).
    - Azure Consumptions (biannual): 1 línea CC 7475036 (100%)
    - Maintenance and support (annual): 1 línea CC 1100036 (100%)
    - Domains (annual): 1 línea CC 7475036 (100%)  <-- Confirmado por ti
    - Si concepto custom: pedir CC y GL al usuario (los vendrán en invoice.extras)
    """
    concept = (invoice.service_concept or "").strip()
    if not concept:
        concept = "Infrastructure cloud (Monthly)"

    iva_rate = invoice.iva_rate

    # Caso custom
    if concept not in CONCEPTS:
        # Si el usuario configuró splits:
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

        # Si NO hay split, comportamiento actual: 1 línea 100% a CC/GL del usuario
        cc = invoice.extras.get("cc")
        gl = invoice.extras.get("gl_account")
        if cc is None or gl is None:
            raise ValueError("Para concepto personalizado en EIKON debes ingresar CC y GL account.")
        return [_make_line(invoice, concept, int(cc), int(gl), invoice.subtotal, iva_rate)]

    # Casos estándar
    if concept == "Infrastructure cloud (Monthly)":
        part1 = q2(invoice.subtotal * Decimal("0.60"))
        part2 = q2(invoice.subtotal * Decimal("0.40"))
        return [
            _make_line(invoice, concept, 7457036, GL_DEFAULT, part1, iva_rate),
            _make_line(invoice, concept, 7475036, GL_DEFAULT, part2, iva_rate),
        ]

    if concept == "Azure Consumptions (biannual)":
        return [_make_line(invoice, concept, 7475036, GL_DEFAULT, invoice.subtotal, iva_rate)]

    if concept == "Maintenance and support (annual)":
        return [_make_line(invoice, concept, 1100036, GL_DEFAULT, invoice.subtotal, iva_rate)]

    if concept == "Domains (annual)":
        return [_make_line(invoice, concept, 7475036, GL_DEFAULT, invoice.subtotal, iva_rate)]

    # fallback defensivo
    raise ValueError(f"Concepto EIKON no manejado: {concept}")


def _make_line(
    invoice: InvoiceInput,
    concept: str,
    cc: int,
    gl: int,
    subtotal_assigned: Decimal,
    iva_rate: Decimal,
) -> LineItem:
    iva, total = calc_iva_and_total(subtotal_assigned, iva_rate)
    return LineItem(
        table_name=EIKON_TABLE,
        values={
            "Date": invoice.invoice_date,  # escribiremos como fecha real
            "Bill number": invoice.bill_number,  # texto 9 dígitos
            "ID": invoice.vendor_id,
            "Vendor": invoice.vendor_name,
            "Service/ concept": concept,
            "CC": cc,
            "GL account": gl,
            "Subtotal assigned by CC": subtotal_assigned,
            "% IVA": iva_rate,
            "IVA assigned by CC": iva,
            "Total assigned by CC": total,
        },
    )
