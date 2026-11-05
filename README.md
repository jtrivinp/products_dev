# BPS Compensador

Aplicación local para identificar propuestas de compensación en Excel de SAP.
No se conecta a SAP ni realiza compensaciones; la revisión y ejecución final es
siempre manual.

## Uso para negocio

1. Edite `config/config.ini` con las carpetas corporativas deseadas.
2. Coloque un Excel de SAP en `input/` (o en la carpeta configurada).
3. Copie los auxiliares aprobados por BPS a `config/auxiliares/`:
   `RETENCIONES.xlsx` para Retención y `CLASES_MASMENOS.xlsx` para Más/Menos.
4. Abra `BPS_Compensador.exe`, seleccione modo y fichero, y pulse **Procesar**.
5. Revise `output/<RUN_ID>/resultado.xlsx` y `resultado.html`. El log técnico
   queda en `logs/bps_<RUN_ID>.log`.

La equivalencia VIP `TP`/GPC `TU`, la exclusión por `Doc.compensación` y las
tablas auxiliares definitivas siguen pendientes de confirmación funcional BPS.

## Desarrollo y pruebas

Requiere Python 3.12+ y `pip install -r requirements.txt`.

Las pruebas de baseline no incluyen Excel reales. Para ejecutarlas:

```bash
BPS_TEST_DATA_DIR="C:\ruta\a\datos" python -m unittest discover -s tests -v
```

## Build Windows

En Windows, desde un entorno virtual con los requisitos instalados, ejecute
`build_windows.bat`. El ejecutable se crea en `dist/BPS_Compensador/`. Mantenga
externos `config.ini`, auxiliares, entradas, salidas y logs; no se empaquetan
datos contables ni configuración productiva.
