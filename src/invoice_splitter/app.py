from __future__ import annotations

from invoice_splitter.ui.main_window import MainWindow
from invoice_splitter.utils.logging import setup_logging
from invoice_splitter.config import get_settings


def main() -> None:
    settings = get_settings()
    # Guardar app.log dentro de la carpeta del proyecto o cerca del excel, tú decides.
    # Opción recomendada: en la carpeta del proyecto
    # (si prefieres junto al excel, lo cambiamos a settings.excel_path.parent)
    setup_logging(log_dir=(settings.excel_path.parent / "invoice_splitter_logs"))

    app = MainWindow()
    app.mainloop()


if __name__ == "__main__":
    main()
