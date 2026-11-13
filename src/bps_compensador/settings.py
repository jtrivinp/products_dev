"""Configuración externa y validación previa a una ejecución."""
from __future__ import annotations
from configparser import ConfigParser
from dataclasses import dataclass
from pathlib import Path
import sys

@dataclass(frozen=True)
class Settings:
    root: Path; input_dir: Path; output_dir: Path; log_dir: Path; aux_dir: Path
    tolerance_cents: int; retention_tolerance_cents: int; max_invoices: int; day_window: int; gpc_class: str; vip_class: str
    retentions_vip_file: str; retentions_gpc_file: str; classes_vip_file: str; classes_gpc_file: str

    def auxiliary(self, filename: str) -> Path: return self.aux_dir / filename

def load(path: Path | None = None) -> Settings:
    root = path.parent.parent if path else (Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parents[2])
    filename = path or root / "config" / "config.ini"
    parser = ConfigParser(); loaded = parser.read(filename, encoding="utf-8")
    if not loaded: raise ValueError(f"No se encuentra el archivo de configuración: {filename}")
    def directory(key):
        value = Path(parser.get("paths", key)); return value if value.is_absolute() else root / value
    try:
        tolerance = round(parser.getfloat("matching", "tolerancia_domiciliadas") * 100)
        return Settings(root, directory("input_dir"), directory("output_dir"), directory("log_dir"), directory("aux_dir"), tolerance, round(parser.getfloat("matching", "tolerancia_retencion") * 100), parser.getint("matching", "max_facturas_domiciliado"), parser.getint("matching", "ventana_dias_retencion"), parser.get("matching", "clase_pago_gpc").upper(), parser.get("matching", "clase_pago_vip").upper(), parser.get("matching", "retenciones_vip_file"), parser.get("matching", "retenciones_gpc_file"), parser.get("matching", "clases_masmenos_vip_file"), parser.get("matching", "clases_masmenos_gpc_file"))
    except Exception as exc: raise ValueError(f"Configuración inválida en {filename}: {exc}") from exc

def validate(settings: Settings) -> None:
    if not settings.input_dir.is_dir(): raise ValueError(f"No existe el directorio de entrada: {settings.input_dir}")
    if not settings.aux_dir.is_dir(): raise ValueError(f"No existe el directorio de auxiliares: {settings.aux_dir}")
    for directory in (settings.output_dir, settings.log_dir):
        directory.mkdir(parents=True, exist_ok=True)
        if not directory.is_dir(): raise ValueError(f"La ruta no es un directorio: {directory}")
