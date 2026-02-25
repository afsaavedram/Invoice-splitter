from __future__ import annotations
from datetime import date
from decimal import Decimal
from typing import List

from invoice_splitter.models import InvoiceInput, LineItem
from invoice_splitter.rules.common import calc_iva_and_total, validate_and_compute_allocations

VENDOR_ID = 9999999
TABLE_NAME = "Dummy_table"  # o el nombre real que quieras que cree el writer


def build_lines_for_vendor(invoice: InvoiceInput) -> List[LineItem]:
    concept_general = (invoice.service_concept or "Dummy").strip() or "Dummy"

    # ✅ Caso 1: split personalizado
    if invoice.alloc_mode and invoice.allocations:
        pairs = validate_and_compute_allocations(
            invoice.subtotal, invoice.alloc_mode, invoice.allocations
        )  # calcula montos por línea y ajusta tolerancia [1](https://exceladept.com/invalid-names-when-opening-a-workbook-in-excel/)[1](https://exceladept.com/invalid-names-when-opening-a-workbook-in-excel/)

        lines: List[LineItem] = []
        for alloc, amount in pairs:
            iva, total = calc_iva_and_total(amount, invoice.iva_rate)
            lines.append(
                LineItem(
                    table_name=TABLE_NAME,
                    values={
                        "Date": invoice.invoice_date,
                        "Bill number": invoice.bill_number,
                        "ID": invoice.vendor_id,
                        "Vendor": invoice.vendor_name,
                        "Service/ concept": (alloc.concept or concept_general).strip(),
                        "CC": int(alloc.cc),
                        "GL account": int(alloc.gl_account),
                        "Subtotal assigned by CC": amount,
                        "% IVA": invoice.iva_rate,
                        "IVA assigned by CC": iva,
                        "Total assigned by CC": total,
                    },
                )
            )
        return lines

    # ✅ Caso 2: sin split (1 línea)
    cc = invoice.extras.get("cc", 1100036)
    gl = invoice.extras.get("gl_account", 7418000000)
    iva, total = calc_iva_and_total(invoice.subtotal, invoice.iva_rate)

    return [
        LineItem(
            table_name=TABLE_NAME,
            values={
                "Date": invoice.invoice_date,
                "Bill number": invoice.bill_number,
                "ID": invoice.vendor_id,
                "Vendor": invoice.vendor_name,
                "Service/ concept": concept_general,
                "CC": int(cc),
                "GL account": int(gl),
                "Subtotal assigned by CC": invoice.subtotal,
                "% IVA": invoice.iva_rate,
                "IVA assigned by CC": iva,
                "Total assigned by CC": total,
            },
        )
    ]
