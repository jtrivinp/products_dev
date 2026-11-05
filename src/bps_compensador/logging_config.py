"""Registro técnico por ejecución, separado de los mensajes de negocio."""
import logging
from pathlib import Path

def configure(log_dir: Path, run_id: str) -> logging.Logger:
    logger = logging.getLogger(f"bps.{run_id}"); logger.setLevel(logging.DEBUG); logger.handlers.clear()
    handler = logging.FileHandler(log_dir / f"bps_{run_id}.log", encoding="utf-8")
    handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
    logger.addHandler(handler); return logger
