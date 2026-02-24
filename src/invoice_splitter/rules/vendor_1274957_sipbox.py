from __future__ import annotations

from typing import List

from invoice_splitter.models import InvoiceInput, LineItem
from invoice_splitter.rules.common import calc_iva_and_total, validate_and_compute_allocations


SIPBOX_TABLE = "Sipbox_table"

DEFAULT_CONCEPT = "Lenovo ThinkSmartHub + Stem speaker + POE switch"

DEFAULT_CC = 7475036
DEFAULT_GL = 7648100000


def build_lines_for_sipbox(invoice: InvoiceInput) -> List[LineItem]:
    """
    SIPBOX (vendor_id=1274957) -> Sipbox_table

    Reglas:
    - Por defecto: 1 línea con:
        concept = DEFAULT_CONCEPT
        CC = DEFAULT_CC
        GL = DEFAULT_GL
        subtotal assigned = subtotal factura (100%)
    - Si el usuario cambia el concepto (concept != DEFAULT_CONCEPT):
        - Si NO hay split (invoice.allocations vacío):
            - requiere CC y GL del usuario (invoice.extras['cc'], invoice.extras['gl_account'])
            - 1 línea 100%
        - Si hay split custom (invoice.alloc_mode + invoice.allocations):
            - se generan N líneas (una por allocation)
            - split puede ser por porcentaje o por valor
            - se valida suma vs subtotal con tolerancia ±0.01 y se ajusta última línea si aplica
            - cada allocation puede traer concept propio; si no, usa el concepto general
    """

    concept = (invoice.service_concept or "").strip() or DEFAULT_CONCEPT
    iva_rate = invoice.iva_rate

    # --- Caso 1: Concepto default -> comportamiento estándar (1 línea) ---
    if concept == DEFAULT_CONCEPT:
        return [_make_line(invoice, concept, DEFAULT_CC, DEFAULT_GL, invoice.subtotal, iva_rate)]

    # --- Caso 2: Concepto custom -> si hay split custom, usarlo ---
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

    # --- Caso 3: Concepto custom sin split -> 1 línea 100% con CC/GL del usuario ---
    cc = invoice.extras.get("cc")
    gl = invoice.extras.get("gl_account")
    if cc is None or gl is None:
        raise ValueError("Para concepto personalizado en SIPBOX debes ingresar CC y GL account.")

    return [_make_line(invoice, concept, int(cc), int(gl), invoice.subtotal, iva_rate)]


def _make_line(
    invoice: InvoiceInput,
    concept: str,
    cc: int,
    gl: int,
    subtotal_assigned: object,
    iva_rate: object,
) -> LineItem:
    """
    Construye una fila para Sipbox_table con IVA/Total calculados.
    cc y gl numéricos (int), bill number texto, fecha date.
    """
    iva, total = calc_iva_and_total(subtotal_assigned, iva_rate)

    return LineItem(
        table_name=SIPBOX_TABLE,
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
