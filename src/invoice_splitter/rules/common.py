from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP
from typing import List, Tuple

from invoice_splitter.models import Allocation

TWOPLACES = Decimal("0.01")
TOLERANCE = Decimal("0.01")


def q2(value: Decimal) -> Decimal:
    """Redondea a 2 decimales con HALF_UP."""
    return value.quantize(TWOPLACES, rounding=ROUND_HALF_UP)


def calc_iva_and_total(subtotal: Decimal, iva_rate: Decimal) -> tuple[Decimal, Decimal]:
    """Calcula IVA y Total con redondeo a 2 decimales."""
    iva = q2(subtotal * iva_rate)
    total = q2(subtotal + iva)
    return iva, total


def parse_percent_value(raw: str | None) -> Decimal:
    """
    Acepta '40.91' o '40.91%' o vacío.
    Devuelve porcentaje como número 0..100 (Decimal).
    Vacío -> 0.00
    """
    if raw is None:
        return Decimal("0")
    s = str(raw).strip()
    if not s:
        return Decimal("0")
    s = s.replace("%", "").strip()
    # permitimos coma/punto en UI; aquí asumimos que ya viene normalizado o bien con '.'
    s = s.replace(",", ".")
    return q2(Decimal(s))


def validate_and_compute_allocations(
    subtotal: Decimal,
    mode: str,
    allocations: List[Allocation],
    tolerance: Decimal = TOLERANCE,
) -> List[Tuple[Allocation, Decimal]]:
    """
    Convierte allocations (porcentaje o valor) en montos (Decimal) por línea.
    - valida sumas
    - si la diferencia está dentro de ±0.01, ajusta la ÚLTIMA línea con la diferencia

    Reglas:
    - mode == 'percent': usa allocation.percent (0..100)
    - mode == 'amount' : usa allocation.amount
    - permite percent = 0 o amount = 0
    - NO exige que % sume 100 (pero si no suma, la diff probablemente fallará)
    """
    if mode not in {"percent", "amount"}:
        raise ValueError("alloc_mode debe ser 'percent' o 'amount'.")

    if not allocations:
        raise ValueError("No hay líneas de split para calcular.")

    amounts: List[Decimal] = []

    if mode == "percent":
        for a in allocations:
            pct = a.percent if a.percent is not None else Decimal("0")
            # pct es 0..100 (si el usuario pone 120, lo dejamos pasar pero la validación fallará normalmente)
            amt = q2(subtotal * pct / Decimal("100"))
            amounts.append(amt)

    else:  # amount
        for a in allocations:
            amt = a.amount if a.amount is not None else Decimal("0")
            amounts.append(q2(amt))

    total = q2(sum(amounts))
    diff = q2(subtotal - total)

    # Validación fuerte: si supera tolerancia, error.
    if diff.copy_abs() > tolerance:
        raise ValueError(
            f"La suma de líneas ({total}) no coincide con el subtotal ({subtotal}). "
            f"Diferencia={diff}. Debes corregir valores/porcentajes."
        )

    # Ajuste permitido dentro de tolerancia
    if diff != Decimal("0"):
        amounts[-1] = q2(amounts[-1] + diff)

    return list(zip(allocations, amounts))
