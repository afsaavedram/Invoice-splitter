from __future__ import annotations

from typing import Callable, Dict, List

from invoice_splitter.models import InvoiceInput, LineItem

from invoice_splitter.rules.vendor_1255097_eikon import build_lines_for_eikon
from invoice_splitter.rules.vendor_1255036_akros import build_lines_for_akros
from invoice_splitter.rules.vendor_1274957_sipbox import build_lines_for_sipbox
from invoice_splitter.rules.vendor_1261182_puntonet import build_lines_for_puntonet
from invoice_splitter.rules.vendor_1254926_cirion import build_lines_for_cirion
from invoice_splitter.rules.vendor_1260177_movistar import build_lines_for_movistar
from invoice_splitter.rules.vendor_1254902_claro import build_lines_for_claro

RuleFn = Callable[[InvoiceInput], List[LineItem]]

RULES: Dict[int, RuleFn] = {
    1255097: build_lines_for_eikon,
    1255036: build_lines_for_akros,
    1274957: build_lines_for_sipbox,
    1261182: build_lines_for_puntonet,
    1254926: build_lines_for_cirion,
    1260177: build_lines_for_movistar,
    1254902: build_lines_for_claro,
}


def build_lines(invoice: InvoiceInput) -> List[LineItem]:
    fn = RULES.get(invoice.vendor_id)
    if not fn:
        raise ValueError(f"No hay regla implementada aún para vendor_id={invoice.vendor_id}")
    return fn(invoice)
