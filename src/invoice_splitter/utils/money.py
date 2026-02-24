from __future__ import annotations

from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
import re

TWOPLACES = Decimal("0.01")


def parse_decimal_user_input(value: str, field_name: str = "El valor") -> Decimal:
    """
    Convierte un string ingresado por el usuario a Decimal redondeado a 2 decimales.
    Acepta separador decimal "," o ".", y separadores de miles.

    field_name se usa para mensajes de error más claros (ej: 'Subtotal').
    """
    if value is None:
        raise ValueError(f"{field_name} no puede estar vacío.")

    s = str(value).strip()
    if not s:
        raise ValueError(f"{field_name} no puede estar vacío.")

    s = s.replace(" ", "")

    if "," in s and "." in s:
        last_comma = s.rfind(",")
        last_dot = s.rfind(".")
        if last_comma > last_dot:
            s = s.replace(".", "")
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    else:
        if "," in s and "." not in s:
            s = s.replace(".", "")
            s = s.replace(",", ".")
        elif "." in s and "," not in s:
            s = s.replace(",", "")

    if not re.fullmatch(r"[+-]?\d+(\.\d+)?", s):
        raise ValueError(f"{field_name} tiene un formato numérico inválido: '{value}'")

    try:
        dec = Decimal(s)
    except InvalidOperation as e:
        raise ValueError(f"{field_name} no se pudo convertir a número: '{value}'") from e

    return dec.quantize(TWOPLACES, rounding=ROUND_HALF_UP)


def parse_iva(value: str, default: Decimal = Decimal("0.15")) -> Decimal:
    """
    Permite ingresar IVA como:
      - '0.15'
      - '15'  (se interpreta como 15%)
      - '15%' (se interpreta como 15%)
    """
    if value is None or str(value).strip() == "":
        return default

    s = str(value).strip().replace("%", "")
    dec = parse_decimal_user_input(s, field_name="IVA")

    if dec > 1:
        dec = dec / Decimal("100")

    return dec.quantize(Decimal("0.0001"), rounding=ROUND_HALF_UP)


def normalize_bill_number(raw: str, field_name: str = "Número de factura") -> str:
    """
    Normaliza el bill number a 9 dígitos con ceros a la izquierda.
    """
    if raw is None:
        raise ValueError(f"{field_name} no puede estar vacío.")

    s = str(raw).strip()
    if not s:
        raise ValueError(f"{field_name} no puede estar vacío.")

    if not s.isdigit():
        raise ValueError(f"{field_name} debe contener solo dígitos.")

    if len(s) > 9:
        raise ValueError(f"{field_name} no puede tener más de 9 dígitos.")

    return s.zfill(9)
