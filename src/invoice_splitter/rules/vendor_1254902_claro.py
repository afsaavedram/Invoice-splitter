from __future__ import annotations

from decimal import Decimal
from typing import List

from invoice_splitter.models import InvoiceInput, LineItem
from invoice_splitter.rules.common import calc_iva_and_total, q2, validate_and_compute_allocations

CLARO_ID = 1254902

SIPTRUNK_TABLE = "Claro_siptrunk_table"
SBC_TABLE = "Claro_SBC_table"
MOBILE_TABLE = "Claro_mobile_table"

OTRO = "Otro (personalizado)"

# ---------------------------
# Defaults por servicio
# ---------------------------
SIPTRUNK_DEFAULT_BW = 50
SIPTRUNK_DEFAULT_CHANNELS = 10

SBC_DEFAULT_CONCEPT = "SBC in cloud"
SBC_DEFAULT_SIPTRUNK_MBPS = 2
SBC_DEFAULT_LIC_QTY = 14
SBC_DEFAULT_SIPTRUNK_PRICE = 80
SBC_DEFAULT_LIC_PRICE = 266

MOBILE_DEFAULT_CONCEPT = "2 lines 50 GB + 3 lines 20 GB"
MOBILE_DEFAULT_LINES_QTY = 5

# ---------------------------
# Claro Mobile: porcentajes confirmados (suman 100%)
# ---------------------------
MOBILE_SPLIT = [
    (7300036, 4649000000, Decimal("40.91")),
    (7410036, 4649000000, Decimal("15.63")),
    (3941036, 3649000000, Decimal("28.41")),
    (7475036, 4649000000, Decimal("15.05")),
]

# ---------------------------
# Claro SBC: split estándar 60/40
# ---------------------------
SBC_SPLIT = [
    (7475036, 7648100000, Decimal("60.00")),
    (3941036, 3648000000, Decimal("40.00")),
]

# ---------------------------
# Claro Siptrunk: 7 líneas
# 2 fijas (300 y 200) y 5 variables sobre (subtotal - fixed_total)
# 5 variables con porcentajes CORREGIDOS (suman 100%):
# ---------------------------
# OJO: guardamos las fijas como montos ABSOLUTOS y aplicamos signo según subtotal.
SIPTRUNK_FIXED_LINES_ABS = [
    ("CONECEL (Internet 50 Mbps) - SD WAN", 7475036, 7418000000, Decimal("300.00")),
    ("CONECEL (Internet 50 Mbps) - SD WAN", 3941036, 3526400000, Decimal("200.00")),
]

SIPTRUNK_VARIABLE_LINES = [
    ("Consumos SIP Trunk - Claro ECUADOR", 7000036, 7648100000, Decimal("37.00")),
    ("Consumos SIP Trunk - Claro ECUADOR", 7100036, 7648100000, Decimal("11.00")),
    ("Consumos SIP Trunk - Claro ECUADOR", 7300036, 7648100000, Decimal("32.00")),
    ("Consumos SIP Trunk - Claro ECUADOR", 3941036, 3648000000, Decimal("15.00")),
    ("Consumos SIP Trunk - Claro ECUADOR", 7475036, 7648100000, Decimal("5.00")),
]


def build_lines_for_claro(invoice: InvoiceInput) -> List[LineItem]:
    """
    CLARO (vendor_id=1254902)
    invoice.service_type: 'siptrunk' | 'sbc' | 'mobile'
    """
    service_type = (invoice.service_type or "").strip().lower()
    if service_type not in {"siptrunk", "sbc", "mobile"}:
        raise ValueError(
            "Para CLARO debes seleccionar el tipo de servicio: Siptrunk, SBC o Mobile."
        )
    if service_type == "siptrunk":
        return _build_siptrunk(invoice)
    if service_type == "sbc":
        return _build_sbc(invoice)
    return _build_mobile(invoice)


# ---------------------------
# Siptrunk
# ---------------------------
def _build_siptrunk(invoice: InvoiceInput) -> List[LineItem]:
    iva_rate = invoice.iva_rate

    bw = int(invoice.extras.get("bandwidth_mbps", SIPTRUNK_DEFAULT_BW))
    channels = int(invoice.extras.get("sip_channels", SIPTRUNK_DEFAULT_CHANNELS))

    concept = (invoice.service_concept or "").strip() or "Claro Siptrunk"

    # Custom (OTRO)
    if concept == OTRO:
        return _build_custom_into_siptrunk_table(invoice, bw, channels)

    # ✅ Signo correcto: positivo si subtotal >= 0, negativo si subtotal < 0
    sign = Decimal("-1") if invoice.subtotal < 0 else Decimal("1")

    lines: List[LineItem] = []

    # 2 fijas con signo (si subtotal negativo => -300 y -200)
    fixed_amounts: List[Decimal] = []
    for cpt, cc, gl, abs_amt in SIPTRUNK_FIXED_LINES_ABS:
        amt = q2(abs(abs_amt) * sign)
        fixed_amounts.append(amt)
        lines.append(_make_siptrunk_line(invoice, cpt, cc, gl, amt, iva_rate, bw, channels))

    fixed_total = q2(sum(fixed_amounts))

    # Base restante para las 5 variables
    base = q2(invoice.subtotal - fixed_total)

    # 5 variables por % (suman 100)
    variable_amounts: List[Decimal] = []
    for _cpt, _cc, _gl, pct in SIPTRUNK_VARIABLE_LINES:
        variable_amounts.append(q2(base * (pct / Decimal("100"))))

    # Ajuste por redondeo para que sum(variable_amounts) == base
    diff = q2(base - sum(variable_amounts))
    if diff != 0:
        variable_amounts[-1] = q2(variable_amounts[-1] + diff)

    for idx, (cpt, cc, gl, _pct) in enumerate(SIPTRUNK_VARIABLE_LINES):
        lines.append(
            _make_siptrunk_line(invoice, cpt, cc, gl, variable_amounts[idx], iva_rate, bw, channels)
        )

    # ✅ Garantía final (cierre exacto subtotal)
    total_assigned = q2(sum(Decimal(str(li.values["Subtotal assigned by CC"])) for li in lines))
    final_diff = q2(invoice.subtotal - total_assigned)
    if final_diff != 0:
        last = lines[-1]
        last.values["Subtotal assigned by CC"] = q2(
            Decimal(str(last.values["Subtotal assigned by CC"])) + final_diff
        )
        iva, total = calc_iva_and_total(last.values["Subtotal assigned by CC"], iva_rate)
        last.values["IVA assigned by CC"] = iva
        last.values["Total assigned by CC"] = total

    return lines


def _build_custom_into_siptrunk_table(
    invoice: InvoiceInput, bw: int, channels: int
) -> List[LineItem]:
    iva_rate = invoice.iva_rate
    concept_general = (
        invoice.extras.get("custom_concept") or ""
    ).strip() or "Claro Siptrunk - Custom"

    if invoice.alloc_mode and invoice.allocations:
        pairs = validate_and_compute_allocations(
            invoice.subtotal, invoice.alloc_mode, invoice.allocations
        )
        return [
            _make_siptrunk_line(
                invoice,
                (alloc.concept or concept_general).strip(),
                alloc.cc,
                alloc.gl_account,
                amount,
                iva_rate,
                bw,
                channels,
            )
            for alloc, amount in pairs
        ]

    cc = invoice.extras.get("cc")
    gl = invoice.extras.get("gl_account")
    if cc is None or gl is None:
        raise ValueError("Para Siptrunk personalizado sin split debes ingresar CC y GL account.")
    return [
        _make_siptrunk_line(
            invoice, concept_general, int(cc), int(gl), invoice.subtotal, iva_rate, bw, channels
        )
    ]


def _make_siptrunk_line(
    invoice: InvoiceInput,
    concept: str,
    cc: int,
    gl: int,
    subtotal_assigned,
    iva_rate,
    bandwidth_mbps: int,
    sip_channels: int,
) -> LineItem:
    iva, total = calc_iva_and_total(subtotal_assigned, iva_rate)
    return LineItem(
        table_name=SIPTRUNK_TABLE,
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
            "Bandwidth (MBPS)": int(bandwidth_mbps),
            "Troncal SIP (channels)": int(sip_channels),
        },
    )


# ---------------------------
# SBC
# ---------------------------
def _build_sbc(invoice: InvoiceInput) -> List[LineItem]:
    iva_rate = invoice.iva_rate
    concept = (invoice.service_concept or "").strip() or SBC_DEFAULT_CONCEPT

    subtotal_negative = invoice.subtotal < 0

    sip_mbps = int(invoice.extras.get("sbc_siptrunk_mbps", SBC_DEFAULT_SIPTRUNK_MBPS))
    lic_qty = int(invoice.extras.get("sbc_lic_qty", SBC_DEFAULT_LIC_QTY))

    sip_price = Decimal(str(invoice.extras.get("sbc_siptrunk_price", SBC_DEFAULT_SIPTRUNK_PRICE)))
    lic_price = Decimal(str(invoice.extras.get("sbc_lic_price", SBC_DEFAULT_LIC_PRICE)))

    # ✅ Requisito: si subtotal es negativo, los prices deben ser negativos
    if subtotal_negative:
        sip_price = -abs(sip_price)
        lic_price = -abs(lic_price)

    if concept == OTRO:
        return _build_custom_into_sbc_table(invoice, sip_mbps, lic_qty, sip_price, lic_price)

    lines: List[LineItem] = []
    for cc, gl, pct in SBC_SPLIT:
        amt = q2(invoice.subtotal * (pct / Decimal("100")))
        lines.append(
            _make_sbc_line(
                invoice,
                SBC_DEFAULT_CONCEPT,
                cc,
                gl,
                amt,
                iva_rate,
                sip_mbps,
                lic_qty,
                sip_price,
                lic_price,
            )
        )

    # Ajuste por redondeo para cerrar subtotal
    total_assigned = q2(sum(Decimal(str(li.values["Subtotal assigned by CC"])) for li in lines))
    diff = q2(invoice.subtotal - total_assigned)
    if diff != 0:
        lines[-1].values["Subtotal assigned by CC"] = q2(
            Decimal(str(lines[-1].values["Subtotal assigned by CC"])) + diff
        )
        iva, total = calc_iva_and_total(lines[-1].values["Subtotal assigned by CC"], iva_rate)
        lines[-1].values["IVA assigned by CC"] = iva
        lines[-1].values["Total assigned by CC"] = total

    return lines


def _build_custom_into_sbc_table(
    invoice: InvoiceInput, sip_mbps: int, lic_qty: int, sip_price: Decimal, lic_price: Decimal
) -> List[LineItem]:
    iva_rate = invoice.iva_rate
    concept_general = (invoice.extras.get("custom_concept") or "").strip() or "SBC - Custom"

    if invoice.alloc_mode and invoice.allocations:
        pairs = validate_and_compute_allocations(
            invoice.subtotal, invoice.alloc_mode, invoice.allocations
        )
        return [
            _make_sbc_line(
                invoice,
                (alloc.concept or concept_general).strip(),
                alloc.cc,
                alloc.gl_account,
                amount,
                iva_rate,
                sip_mbps,
                lic_qty,
                sip_price,
                lic_price,
            )
            for alloc, amount in pairs
        ]

    cc = invoice.extras.get("cc")
    gl = invoice.extras.get("gl_account")
    if cc is None or gl is None:
        raise ValueError("Para SBC personalizado sin split debes ingresar CC y GL account.")
    return [
        _make_sbc_line(
            invoice,
            concept_general,
            int(cc),
            int(gl),
            invoice.subtotal,
            iva_rate,
            sip_mbps,
            lic_qty,
            sip_price,
            lic_price,
        )
    ]


def _make_sbc_line(
    invoice: InvoiceInput,
    concept: str,
    cc: int,
    gl: int,
    subtotal_assigned,
    iva_rate,
    sip_mbps: int,
    lic_qty: int,
    sip_price: Decimal,
    lic_price: Decimal,
) -> LineItem:
    iva, total = calc_iva_and_total(subtotal_assigned, iva_rate)
    return LineItem(
        table_name=SBC_TABLE,
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
            "Siptrunk (MBPS)": int(sip_mbps),
            "Licences (Quantity)": int(lic_qty),
            "Siptrunk price": float(sip_price),
            "Licences price": float(lic_price),
        },
    )


# ---------------------------
# Mobile
# ---------------------------
def _build_mobile(invoice: InvoiceInput) -> List[LineItem]:
    iva_rate = invoice.iva_rate
    concept = (invoice.service_concept or "").strip() or MOBILE_DEFAULT_CONCEPT
    phone_lines = int(invoice.extras.get("mobile_phone_lines_qty", MOBILE_DEFAULT_LINES_QTY))

    if concept == OTRO:
        return _build_custom_into_mobile_table(invoice, phone_lines)

    lines: List[LineItem] = []
    for cc, gl, pct in MOBILE_SPLIT:
        amt = q2(invoice.subtotal * (pct / Decimal("100")))
        lines.append(
            _make_mobile_line(invoice, MOBILE_DEFAULT_CONCEPT, cc, gl, amt, iva_rate, phone_lines)
        )

    total_assigned = q2(sum(Decimal(str(li.values["Subtotal assigned by CC"])) for li in lines))
    diff = q2(invoice.subtotal - total_assigned)
    if diff != 0:
        lines[-1].values["Subtotal assigned by CC"] = q2(
            Decimal(str(lines[-1].values["Subtotal assigned by CC"])) + diff
        )
        iva, total = calc_iva_and_total(lines[-1].values["Subtotal assigned by CC"], iva_rate)
        lines[-1].values["IVA assigned by CC"] = iva
        lines[-1].values["Total assigned by CC"] = total

    return lines


def _build_custom_into_mobile_table(invoice: InvoiceInput, phone_lines: int) -> List[LineItem]:
    iva_rate = invoice.iva_rate
    concept_general = (
        invoice.extras.get("custom_concept") or ""
    ).strip() or "Claro Mobile - Custom"

    if invoice.alloc_mode and invoice.allocations:
        pairs = validate_and_compute_allocations(
            invoice.subtotal, invoice.alloc_mode, invoice.allocations
        )
        return [
            _make_mobile_line(
                invoice,
                (alloc.concept or concept_general).strip(),
                alloc.cc,
                alloc.gl_account,
                amount,
                iva_rate,
                phone_lines,
            )
            for alloc, amount in pairs
        ]

    cc = invoice.extras.get("cc")
    gl = invoice.extras.get("gl_account")
    if cc is None or gl is None:
        raise ValueError("Para Mobile personalizado sin split debes ingresar CC y GL account.")
    return [
        _make_mobile_line(
            invoice, concept_general, int(cc), int(gl), invoice.subtotal, iva_rate, phone_lines
        )
    ]


def _make_mobile_line(
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
        table_name=MOBILE_TABLE,
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
            "Phone lines quantity": int(phone_lines_qty),
        },
    )
