"""Caso de uso de aplicación: configuración, trazabilidad y motor."""
from __future__ import annotations
from datetime import datetime
from pathlib import Path
import logging
from . import __version__
from .logging_config import configure
from .settings import Settings, validate
from .batch import run_batch

def candidates(directory: Path) -> list[Path]: return sorted(p for p in directory.iterdir() if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm"} and not p.name.startswith("~$"))

def run(settings: Settings) -> tuple[dict, Path, Path, Path]:
    validate(settings)
    if not settings.auxiliary(settings.retentions_file).is_file():
        raise ValueError(f"No se encuentra la tabla de retenciones: {settings.auxiliary(settings.retentions_file)}")
    if not settings.auxiliary(settings.classes_file).is_file():
        raise ValueError(f"No se encuentra la tabla de clases Más/Menos: {settings.auxiliary(settings.classes_file)}")
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    logger = configure(settings.log_dir, run_id)
    logger.info("Inicio de ejecución consolidada; versión %s", __version__)
    try:
        stats, output, report, processed_count = run_batch(settings, logger)
        logger.info("Fin correcto: partidas=%s grupos=%s salida=%s informe=%s", stats["proposed_rows"], stats["groups"], output, report)
        return stats, output, report, settings.log_dir / f"bps_{run_id}.log"
    except Exception:
        logger.exception("Error de ejecución"); raise
