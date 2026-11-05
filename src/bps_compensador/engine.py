"""Orquestador sin interfaz de usuario para los tres modos funcionales."""

from pathlib import Path
from .modes import domiciliadas, masmenos, retencion


def process(modo: str, partidas: Path, output: Path, retenciones: Path | None = None, clases: Path | None = None, ventana_dias: int = 31, clase_pago_domiciliado: str | None = None, tolerancia_centimos: int = 3, max_facturas_domiciliado: int = 4) -> dict:
    if modo == "abiertas":
        if retenciones is None: raise ValueError("El modo 'abiertas' requiere --retenciones.")
        return retencion.process(partidas, retenciones, output, ventana_dias)
    if modo == "domiciliados": return domiciliadas.process(partidas, output, clase_pago_domiciliado, tolerancia_centimos, max_facturas_domiciliado)
    if modo == "masmenos":
        if clases is None: raise ValueError("El modo 'masmenos' requiere --clases.")
        return masmenos.process(partidas, clases, output)
    raise ValueError(f"Modo no soportado: {modo}")
