"""Regresión de integración del paquete extraído en la Fase 3."""

from __future__ import annotations

import os
import sys
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))
from bps_compensador.engine import process

_data_dir = os.environ.get("BPS_TEST_DATA_DIR")
DATA_DIR = Path(_data_dir) if _data_dir else None


@unittest.skipUnless(DATA_DIR and DATA_DIR.is_dir(), "Defina BPS_TEST_DATA_DIR con los Excel de baseline")
class RefactoredBaselineTests(unittest.TestCase):
    def test_all_supported_datasets_keep_legacy_counts(self):
        cases = [
            ("abiertas", "ABIERTAS VIP 23.09 1.XLSX", {"retenciones": DATA_DIR / "RETENCIONES_EJEMPLO.xlsx"}, 299, 144),
            ("domiciliados", "Domiciliados GPC.XLSX", {}, 222, 108),
            ("domiciliados", "Domiciliados VIP.XLSX", {}, 8, 3),
            ("masmenos", "Mas Menos GPC.XLSX", {"clases": DATA_DIR / "CLASES_MASMENOS_PRUEBA.xlsx"}, 48, 24),
            ("masmenos", "Mas Menos VIP.XLSX", {"clases": DATA_DIR / "CLASES_MASMENOS_PRUEBA.xlsx"}, 30, 15),
        ]
        with tempfile.TemporaryDirectory(prefix="bps-refactor-") as temp:
            for mode, filename, arguments, proposed, groups in cases:
                with self.subTest(mode=mode, filename=filename):
                    result = process(mode, DATA_DIR / filename, Path(temp) / f"{mode}-{filename}.xlsx", **arguments)
                    self.assertEqual(result["proposed_rows"], proposed)
                    self.assertEqual(result["groups"], groups)
