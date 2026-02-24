from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import os

from dotenv import load_dotenv

import json


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


CONFIG_DIR_NAME = "InvoiceSplitter"
CONFIG_FILE_NAME = "config.json"


def get_user_config_path() -> Path:
    """
    Ruta del config por usuario en Windows:
    %APPDATA%\\InvoiceSplitter\\config.json
    """
    appdata = os.getenv("APPDATA")
    if appdata:
        base = Path(appdata)
    else:
        # Fallback por si APPDATA no está (raro en Win, pero por seguridad)
        base = Path.home() / "AppData" / "Roaming"
    return base / CONFIG_DIR_NAME / CONFIG_FILE_NAME


def load_user_config() -> dict:
    path = get_user_config_path()
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        # Si está corrupto, no tumbamos la app: volvemos a config vacío
        return {}


def save_user_config(cfg: dict) -> None:
    path = get_user_config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(cfg, indent=2, ensure_ascii=False), encoding="utf-8")


def set_excel_path_user_config(excel_path: Path) -> None:
    cfg = load_user_config()
    cfg["excel_path"] = str(excel_path)
    save_user_config(cfg)


def get_excel_path_from_sources() -> Path | None:
    """
    Prioridad:
    1) EXCEL_PATH en .env / variables de entorno
    2) excel_path en config.json del usuario
    """
    raw_env = os.getenv("EXCEL_PATH", "").strip()
    if raw_env:
        return Path(raw_env)

    cfg = load_user_config()
    raw_cfg = str(cfg.get("excel_path", "")).strip()

    # ✅ Defensa: si quedó guardada la repr de un objeto (ej. "<click.types.Path object ...>"), ignorar
    if raw_cfg.startswith("<") and "Path object" in raw_cfg:
        # opcional: borrar el valor corrupto para que no vuelva a molestar
        cfg.pop("excel_path", None)
        save_user_config(cfg)
        raw_cfg = ""

    if raw_cfg:
        return Path(raw_cfg)

    return None


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
    Lee EXCEL_PATH desde:
    - .env / variables de entorno
    - o config.json por usuario (si no hay env)
    """
    excel_path = get_excel_path_from_sources()
    if excel_path is None:
        raise ValueError(
            "No se ha configurado el Excel. "
            "Selecciona el archivo desde la UI (File Picker) o define EXCEL_PATH en .env."
        )
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
    """
    Lee variables de entorno (.env) y devuelve un Settings validado.
    Lanza errores claros si falta algo crítico.
    """
    excel_path_raw = os.getenv("EXCEL_PATH", "").strip()
    if not excel_path_raw:
        raise ValueError(
            "Falta EXCEL_PATH en el archivo .env. Ejemplo: EXCEL_PATH=C:\\ruta\\archivo.xlsx"
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
