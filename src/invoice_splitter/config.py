from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import os

from dotenv import load_dotenv


def _project_root() -> Path:
    """
    Devuelve la raíz del proyecto (carpeta que contiene /src).
    Estructura esperada:
      project_root/
        src/
          invoice_splitter/
            config.py   <-- aquí
    """
    return Path(__file__).resolve().parents[2]


# Cargamos .env desde la raíz del proyecto.
load_dotenv(_project_root() / ".env")


@dataclass(frozen=True)
class Settings:
    """
    Configuración centralizada del proyecto.
    - EXCEL_PATH: ruta del archivo Excel principal.
    - DEFAULT_IVA: IVA por defecto (ej: 0.15).
    - DATE_DISPLAY_FORMAT: formato visual en Excel (ej: dd-mmm-yy).
    - VENDORS_SHEET / VENDORS_TABLE: ubicación de la tabla Vendors_table.
    """
    excel_path: Path
    default_iva: str
    date_display_format: str
    vendors_sheet: str
    vendors_table: str


def get_settings() -> Settings:
    """
    Lee variables de entorno (.env) y devuelve un Settings validado.
    Lanza errores claros si falta algo crítico.
    """
    excel_path_raw = os.getenv("EXCEL_PATH", "").strip()
    if not excel_path_raw:
        raise ValueError(
            "Falta EXCEL_PATH en el archivo .env. "
            "Ejemplo: EXCEL_PATH=C:\\ruta\\archivo.xlsx"
        )

    excel_path = Path(excel_path_raw)
    if not excel_path.exists():
        raise FileNotFoundError(f"No se encontró el Excel en: {excel_path}")

    default_iva = os.getenv("DEFAULT_IVA", "0.15").strip()
    date_display_format = os.getenv("DATE_DISPLAY_FORMAT", "dd-mmm-yy").strip()

    vendors_sheet = os.getenv("VENDORS_SHEET", "Vendors").strip()
    vendors_table = os.getenv("VENDORS_TABLE", "Vendors_table").strip()

    return Settings(
        excel_path=excel_path,
        default_iva=default_iva,
        date_display_format=date_display_format,
        vendors_sheet=vendors_sheet,
        vendors_table=vendors_table,
    )