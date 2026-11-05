# Baseline del prototipo 0.0 legado

Ejecución de caracterización realizada el 5 de noviembre de 2026 con los
ficheros de prueba entregados por BPS y el script legado sin cambios funcionales.

| Modo | Sistema | Filas analizadas | Partidas propuestas | Grupos | Desglose |
| --- | --- | ---: | ---: | ---: | --- |
| Retención | VIP | 8.313 | 299 | 144 | R01: 57; R02 exacto: 83; R02 retención: 3; R03 1:2: 1 |
| Domiciliadas | GPC | 785 | 222 | 108 | DOM01: 103; DOM02: 5 |
| Domiciliadas | VIP | 470 | 8 | 3 | DOM01: 2; DOM02: 1 |
| Más/Menos | GPC | 3.022 | 48 | 24 | MM01: 24 |
| Más/Menos | VIP | 855 | 30 | 15 | MM01: 15 |

Observaciones de la ejecución:

- Retención utilizó `G = 15 %`; los indicadores no parametrizados fueron
  `G0`, `G1`, `G5`, `G7`, `V`, `Z0` y `ZX`.
- Domiciliadas GPC utilizó `TU`; VIP utilizó `TP`, que sigue siendo un supuesto
  de prueba pendiente de confirmación funcional.
- Más/Menos utilizó exclusivamente la clase `KR` de la tabla de prueba.

Estas cifras son una caracterización de regresión, no criterios funcionales
definitivos. Cualquier cambio intencional en las reglas debe explicar la
variación antes de actualizar este documento y los tests asociados.
