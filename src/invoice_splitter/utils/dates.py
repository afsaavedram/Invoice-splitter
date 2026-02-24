from __future__ import annotations

from datetime import date, datetime


UI_DATE_FORMAT = "%Y-%m-%d"  # Formato estable para el DateEntry (entrada/salida)


def today() -> date:
    """Devuelve la fecha actual."""
    return date.today()


def parse_ui_date(value: str) -> date:
    """
    Convierte el string del DateEntry a date usando UI_DATE_FORMAT.
    """
    return datetime.strptime(value, UI_DATE_FORMAT).date()
