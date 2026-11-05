# Baseline de regresión

Estas pruebas caracterizan el comportamiento del prototipo antes de extraer el
motor y modificar su arquitectura. Los Excel reales no se versionan.

Para ejecutarlas, sitúe los ficheros de prueba en una carpeta local y ejecute:

```bash
BPS_TEST_DATA_DIR="/ruta/a/los/datos" python -m unittest -v tests.test_baseline_regression
```

Ficheros requeridos:

- `ABIERTAS VIP 23.09 1.XLSX`
- `RETENCIONES_EJEMPLO.xlsx`
- `Domiciliados GPC.XLSX`
- `Domiciliados VIP.XLSX`
- `Mas Menos GPC.XLSX`
- `Mas Menos VIP.XLSX`
- `CLASES_MASMENOS_PRUEBA.xlsx`

Los resultados reflejan la parametrización de prueba disponible: indicador G
al 15 %, clase KR en Más/Menos y TP como equivalente provisional de TU en VIP.
