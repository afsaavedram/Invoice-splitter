from __future__ import annotations

import re
from decimal import Decimal
from typing import List

from invoice_splitter.models import InvoiceInput, LineItem
from invoice_splitter.rules.common import calc_iva_and_total, q2, validate_and_compute_allocations


def _slug_table_name(vendor_name: str, vendor_id: int) -> str:
    """
    Construye un nombre de tabla único y estable para vendors genéricos.

    Regla robusta (anti-colisión):
      - <VendorNameNormalizado>_<VendorID>_table

    Notas:
      - Se normaliza a A-Z0-9 y '_' (sin espacios) para que sea válido como nombre de Excel Table.
      - Se trunca el nombre base si queda demasiado largo, para mantener un nombre razonable.
    """
    base = (vendor_name or "").strip()
    base = re.sub(r"\s+", "_", base)
    base = re.sub(r"[^0-9A-Za-z_]", "_", base)
    base = re.sub(r"_+", "_", base).strip("_")

    if not base:
        base = "Vendor"

    # Sufijo fijo: _<id>_table
    suffix = f"_{vendor_id}_table"

    # Truncamos base para evitar nombres excesivamente largos
    # (Excel Table displayName tolera más que 31, pero mantenemos algo razonable)
    max_base_len = 50
    if len(base) > max_base_len:
        base = base[:max_base_len].rstrip("_")

    return f"{base}{suffix}"


def build_lines_generic(invoice: InvoiceInput) -> List[LineItem]:
    """
    Fallback genérico si no existe regla específica por vendor_id.
    - Destino: tabla '<VendorName>_table' (normalizada)
    - Split:
        a) Si invoice.alloc_mode + invoice.allocations: N líneas (por allocation)
        b) Si no hay split: 1 línea con CC/GL de invoice.extras['cc'], invoice.extras['gl_account']
    """
    table_name = _slug_table_name(invoice.vendor_name, invoice.vendor_id)

    concept_general = (invoice.service_concept or "").strip() or "Concepto personalizado"
    iva_rate = invoice.iva_rate

    # Caso split personalizado
    if invoice.alloc_mode and invoice.allocations:
        pairs = validate_and_compute_allocations(
            invoice.subtotal, invoice.alloc_mode, invoice.allocations
        )
        lines: List[LineItem] = []
        for alloc, amount in pairs:
            iva, total = calc_iva_and_total(amount, iva_rate)
            lines.append(
                LineItem(
                    table_name=table_name,
                    values={
                        "Date": invoice.invoice_date,
                        "Bill number": invoice.bill_number,
                        "ID": invoice.vendor_id,
                        "Vendor": invoice.vendor_name,
                        "Service/ concept": (alloc.concept or concept_general).strip()
                        or concept_general,
                        "CC": int(alloc.cc),
                        "GL account": int(alloc.gl_account),
                        "Subtotal assigned by CC": q2(amount),
                        "% IVA": iva_rate,
                        "IVA assigned by CC": iva,
                        "Total assigned by CC": total,
                    },
                )
            )
        return lines

    # Caso sin split: requiere CC/GL en extras
    cc = invoice.extras.get("cc")
    gl = invoice.extras.get("gl_account")
    if cc is None or gl is None:
        raise ValueError(
            "Para vendors sin regla específica debes ingresar CC y GL (o configurar split personalizado)."
        )

    iva, total = calc_iva_and_total(invoice.subtotal, iva_rate)
    return [
        LineItem(
            table_name=table_name,
            values={
                "Date": invoice.invoice_date,
                "Bill number": invoice.bill_number,
                "ID": invoice.vendor_id,
                "Vendor": invoice.vendor_name,
                "Service/ concept": concept_general,
                "CC": int(cc),
                "GL account": int(gl),
                "Subtotal assigned by CC": q2(invoice.subtotal),
                "% IVA": iva_rate,
                "IVA assigned by CC": iva,
                "Total assigned by CC": total,
            },
        )
    ]
