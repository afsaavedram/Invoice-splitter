from __future__ import annotations

import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path


def setup_logging(log_dir: Path) -> None:
    """
    Configura logging:
    - app.log rotativo (evita crecer infinito)
    - nivel INFO
    """
    log_dir.mkdir(parents=True, exist_ok=True)
    log_path = log_dir / "app.log"

    logger = logging.getLogger("invoice_splitter")
    logger.setLevel(logging.INFO)

    # Evitar duplicación si reinicias en caliente
    if logger.handlers:
        return

    fmt = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")

    # Rotación: 5 MB por archivo, conserva 5 copias
    fh = RotatingFileHandler(log_path, maxBytes=5_000_000, backupCount=5, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)

    # También imprime a consola (útil en desarrollo)
    sh = logging.StreamHandler()
    sh.setFormatter(fmt)
    logger.addHandler(sh)
