"""Caso de uso de aplicación: configuración, trazabilidad y motor."""
from __future__ import annotations
from datetime import datetime
from pathlib import Path
import logging
from . import __version__
from .engine import process
from .logging_config import configure
from .settings import Settings, validate
from .html_report import generate

MODE_LABELS = {"retencion": "0 - Retención", "domiciliadas": "1 - Domiciliadas", "masmenos": "2 - Más/Menos"}

def candidates(directory: Path) -> list[Path]: return sorted(p for p in directory.iterdir() if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm"} and not p.name.startswith("~$"))

def run(settings: Settings, mode: str, source: Path) -> tuple[dict, Path, Path, Path]:
    validate(settings)
    if not source.is_file(): raise ValueError("No se encuentra el fichero Excel seleccionado.")
    if mode == "retencion" and not settings.auxiliary(settings.retentions_file).is_file():
        raise ValueError(f"No se encuentra la tabla de retenciones: {settings.auxiliary(settings.retentions_file)}")
    if mode == "masmenos" and not settings.auxiliary(settings.classes_file).is_file():
        raise ValueError(f"No se encuentra la tabla de clases Más/Menos: {settings.auxiliary(settings.classes_file)}")
    run_id = datetime.now().strftime("%Y%m%d_%H%M%S")
    logger = configure(settings.log_dir, run_id); output_dir = settings.output_dir / run_id; output_dir.mkdir(parents=True, exist_ok=False)
    output = output_dir / "resultado.xlsx"; logger.info("Inicio %s; versión %s; entrada=%s", mode, __version__, source.name)
    try:
        if mode == "retencion":
            stats = process("abiertas", source, output, retenciones=settings.auxiliary(settings.retentions_file), ventana_dias=settings.day_window, tolerancia_retencion_centimos=settings.retention_tolerance_cents)
        elif mode == "domiciliadas":
            payment_class = settings.gpc_class if "Cuenta" in _headers(source) else settings.vip_class
            stats = process("domiciliados", source, output, clase_pago_domiciliado=payment_class, tolerancia_centimos=settings.tolerance_cents, max_facturas_domiciliado=settings.max_invoices)
        elif mode == "masmenos": stats = process("masmenos", source, output, clases=settings.auxiliary(settings.classes_file))
        else: raise ValueError(f"Modo no reconocido: {mode}")
        report = generate(output, output_dir / "resultado.html", MODE_LABELS[mode], source.name, stats)
        logger.info("Fin correcto: partidas=%s grupos=%s salida=%s informe=%s", stats["proposed_rows"], stats["groups"], output, report)
        return stats, output, report, settings.log_dir / f"bps_{run_id}.log"
    except Exception:
        logger.exception("Error de ejecución"); raise

def _headers(path: Path) -> set[str]:
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True); return {str(cell.value).strip() for cell in wb[wb.sheetnames[0]][1] if cell.value is not None}
