"""Pruebas de caracterización del prototipo multimodo.

Los Excel de SAP no forman parte del repositorio. Para ejecutar estas pruebas
de integración, defina BPS_TEST_DATA_DIR con la carpeta que los contiene.
"""

from __future__ import annotations

import importlib.util
import os
import tempfile
import unittest
from pathlib import Path


DATA_DIR = Path(os.environ.get("BPS_TEST_DATA_DIR", ""))
SOURCE = (
    Path(__file__).resolve().parents[1]
    / "src"
    / "legacy"
    / "compensador_partidas_multimodo.py"
)


def load_prototype():
    spec = importlib.util.spec_from_file_location("bps_legacy", SOURCE)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@unittest.skipUnless(DATA_DIR.is_dir(), "Defina BPS_TEST_DATA_DIR con los Excel de baseline")
class BaselineRegressionTests(unittest.TestCase):
    """Congela el comportamiento observado antes del refactor funcional."""

    @classmethod
    def setUpClass(cls):
        cls.engine = load_prototype()
        cls.tempdir = tempfile.TemporaryDirectory(prefix="bps-baseline-")
        cls.output_dir = Path(cls.tempdir.name)

    @classmethod
    def tearDownClass(cls):
        cls.tempdir.cleanup()

    def assert_stats(self, stats, expected):
        self.assertEqual(stats["proposed_rows"], expected["proposed_rows"])
        self.assertEqual(stats["groups"], expected["groups"])
        self.assertEqual(stats["by_type"], expected["by_type"])

    def test_retencion_vip(self):
        stats = self.engine.process_abiertas(
            DATA_DIR / "ABIERTAS VIP 23.09 1.XLSX",
            DATA_DIR / "RETENCIONES_EJEMPLO.xlsx",
            self.output_dir / "retencion_vip.xlsx",
        )
        self.assert_stats(
            stats,
            {
                "proposed_rows": 299,
                "groups": 144,
                "by_type": {
                    "R01_GRUPO_EXACTO": 57,
                    "R02_1A1_EXACTO": 83,
                    "R02_1A1_RETENCION": 3,
                    "R03_1A2": 1,
                },
            },
        )
        self.assertEqual(stats["unknown_retention_codes"], ["G0", "G1", "G5", "G7", "V", "Z0", "ZX"])

    def test_domiciliadas_gpc(self):
        stats = self.engine.process_domiciliados(
            DATA_DIR / "Domiciliados GPC.XLSX",
            self.output_dir / "domiciliadas_gpc.xlsx",
        )
        self.assert_stats(stats, {"proposed_rows": 222, "groups": 108, "by_type": {"DOM01_1A1": 103, "DOM02_1AN": 5}})
        self.assertEqual(stats["payment_class"], "TU")

    def test_domiciliadas_vip(self):
        stats = self.engine.process_domiciliados(
            DATA_DIR / "Domiciliados VIP.XLSX",
            self.output_dir / "domiciliadas_vip.xlsx",
        )
        self.assert_stats(stats, {"proposed_rows": 8, "groups": 3, "by_type": {"DOM01_1A1": 2, "DOM02_1AN": 1}})
        self.assertEqual(stats["payment_class"], "TP")
        self.assertIn("SUPUESTO", stats["payment_class_source"])

    def test_masmenos_gpc(self):
        stats = self.engine.process_masmenos(
            DATA_DIR / "Mas Menos GPC.XLSX",
            DATA_DIR / "CLASES_MASMENOS_PRUEBA.xlsx",
            self.output_dir / "masmenos_gpc.xlsx",
        )
        self.assert_stats(stats, {"proposed_rows": 48, "groups": 24, "by_type": {"MM01_1A1_EXACTO": 24}})

    def test_masmenos_vip(self):
        stats = self.engine.process_masmenos(
            DATA_DIR / "Mas Menos VIP.XLSX",
            DATA_DIR / "CLASES_MASMENOS_PRUEBA.xlsx",
            self.output_dir / "masmenos_vip.xlsx",
        )
        self.assert_stats(stats, {"proposed_rows": 30, "groups": 15, "by_type": {"MM01_1A1_EXACTO": 15}})


if __name__ == "__main__":
    unittest.main()
