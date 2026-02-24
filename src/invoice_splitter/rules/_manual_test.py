from datetime import date
from decimal import Decimal

from invoice_splitter.models import InvoiceInput
from invoice_splitter.rules.registry import build_lines

inv = InvoiceInput(
    invoice_date=date.today(),
    vendor_id=1255097,
    vendor_name="EIKON SA",
    bill_number="000000472",
    subtotal=Decimal("100.00"),
    iva_rate=Decimal("0.15"),
    service_concept="Domains (annual)",
)

lines = build_lines(inv)
print(lines)
