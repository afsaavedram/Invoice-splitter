from __future__ import annotations
from pathlib import Path

from invoice_splitter.ui.main_window import MainWindow
from invoice_splitter.utils.logging import setup_logging


def main() -> None:

    # Logging en carpeta del proyecto (no depende de EXCEL_PATH)
    project_root = Path(__file__).resolve().parents[2]

    # Guardar app.log dentro de la carpeta del proyecto o cerca del excel, tú decides.
    # Opción recomendada: en la carpeta del proyecto
    # (si prefieres junto al excel, lo cambiamos a settings.excel_path.parent)
    setup_logging(log_dir=(project_root / "invoice_splitter_logs"))

    app = MainWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
