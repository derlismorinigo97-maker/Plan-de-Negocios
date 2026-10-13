# ¿El PPT «Plan Original vs Proyección Actualizada a 10 años» usa los datos del Excel?

Revisión del 13/10/2026. Archivos comparados:

| Archivo | Descripción |
|---|---|
| `Informe_Ejecutivo_Plan_vs_Actualizado_ALIANZA_DEL_ACERO_10_AÑOS.pptx` | 14 diapositivas, fechado «Junio 2026» |
| `Informe_Financiero_ALIANZA_DEL_ACERO_ORIGINAL.xlsx` | ORIGINAL, guardado el 11/06/2026 |
| `Informe_Financiero_ALIANZA_DEL_ACERO_INFORME2_09062026_AJUSTADO_A_10_AÑOS.xlsx` | INFORME2, guardado el 09/06/2026 |

Método:

1. Cada cifra del PPT se rastreó hasta su celda de origen. En total fueron 243 cifras: tablas, textos y cada punto de los 6 gráficos.
2. Un segundo revisor independiente intentó refutar cada «sin fuente» y cada «distinto».

El detalle completo está en `Trazabilidad_PPT_Plan10_vs_Excel.xlsx`.

## Conclusión

**Sí para la proyección actualizada; no para el plan original.**

**Los dos Excel son el mismo modelo.**

- Tienen las mismas 27 hojas y los mismos valores, salvo 64 celdas que no tocan la proyección.
- La diferencia visible está en el payback del Dashboard, que es un texto tipeado: «6 Años» en ORIGINAL y «4 Años» en INFORME2. El PPT dice 6 años, como ORIGINAL.
- ORIGINAL tiene `#¡REF!` en la hoja Payback y depende de INFORME2 por un vínculo externo.

**Lo «actualizado» del PPT sale de los Excel.** Coinciden:

- VAN USD 2.844.948, TIR 14,9% y WACC 9,87%;
- inversión USD 5.074.594;
- las series anuales de producción, ingresos, costos de producción, EBITDA, ventas y punto de equilibrio;
- el capital de trabajo y los préstamos CP.

Los gráficos del PPT tienen los valores pegados, no vinculados al Excel.

**Lo del «Plan Original» casi no está en los Excel.**

- No tienen fuente en ningún libro: VAN USD 2.843.528, TIR 15,19%, payback de 3 años, las series del plan de los años 3 a 10, los costos y el EBITDA del plan, y los acumulados de 180M, 21,1M y 17,8M.
- Solo producción e ingresos del plan de los años 1 y 2 coinciden con el libro.
- El EBITDA y la utilidad neta del plan de los años 1 y 2 no coinciden con el libro:
  - EBITDA del año 1: 733 mil en el PPT y 841 mil en Ej_Resultados!P13;
  - utilidad neta del año 1: 253 mil en el PPT y 361 mil en Ej_Resultados!P20.

**Hay errores en el propio PPT**, porque varias cifras ni siquiera coinciden con sus gráficos:

| Dato | PPT | Excel / suma de los gráficos del PPT |
|---|---|---|
| Ingresos acumulados 10 años | USD 217M (+21%) | USD 215,6M (+20,3% vs plan 179,2M) |
| EBITDA acumulado | USD 26,4M (+25%) | USD 25,4M (+17,5% vs plan 21,6M) |
| Utilidad neta acumulada | USD 19,6M (+10%) | USD 18,0M (+1,3% vs 17,8M) |
| Producción acumulada | +7% | +7,6% |
| MP acumulada | USD 183M | USD 181,1M |
| TIR vs plan | −0,29 pp | −0,27 pp |
| Margen de seguridad | 52-69% | 50,4-69,3% |
| Préstamos CP | USD 5-12M/año | USD 4,0-11,9M/año |
| Desvío del costo de MP | +6% a +16% | +2,7% a +15,8% (año 2) |
| Productos activos en el año 1 | 11 | 9 |

**La sensibilidad de la diapositiva 12 no tiene respaldo.**

- Ningún libro calcula sensibilidades.
- Una réplica del modelo da magnitudes mucho mayores: ±5% de precio mueve el VAN unos ±USD 5,9M (no ±1,2M), y ±5% de materia prima unos ∓USD 5,2M (no ±1,0M).
- Con un precio 5% más bajo, el VAN sería negativo. Es un punto relevante para la Gerencia.

## Recuento

| Estado | Cifras |
|---|---:|
| Coinciden exactamente con el Excel | 70 |
| Se calculan desde el Excel | 24 |
| Afirmación coherente | 54 |
| Distintas al Excel | 42 |
| Afirmación contradicha por el Excel | 8 |
| Sin fuente en ninguno de los dos Excel | 45 |
| **Total** | **243** |

Las cifras «sin fuente» son, casi todas, del plan original. Probablemente vienen del plan de negocio de enero de 2026 (PDF), que no se adjuntó.
