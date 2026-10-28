# Ecostar v7 (cuadros con la estructura del plan de negocio): pruebas paso a paso

Fecha: 28/10/2026. Repite todas las pruebas de las versiones anteriores sobre la v7 y agrega las de los cuadros.

- **Libro probado:** `Modelo_Maquila_ECOSTAR_v7.xlsx`: la v6 más la hoja «Cuadros» y 15 hojas con los cuadros del plan (C01 a C20 y RRHH). Los contadores de Detalle_24m pasaron a Calculo.
- **Resultados:** `Pruebas_Escenarios_ECOSTAR_v7.xlsx`. Contiene el resumen, la evolución anual por escenario, las verificaciones y la hoja «Cuadros_pruebas».
- **Recálculo:** LibreOffice 24.2. En Excel lo valida Claude en Excel con `Prompt_Claude_Excel_ECOSTAR_v7.md`; la ruta corta alcanza.
- **Unidades:** USD, supuestos del plan de octubre de 2025, sin datos reales.

## Resumen

| Bloque de pruebas | Casos | Comprobaciones | Resultado |
|---|---|---|---|
| Regresión frente a la v6, celda por celda: Calculo, Resultados, Proyeccion_24m, Guia_Carga, Control y hojas de entrada | 1 | 50.484 celdas, 0 distintas | OK |
| Batería de escenarios (financiamiento, cobranza, costos, ventas, plan desfavorable, plazos, corto vs largo plazo, TC y 25 productos) | 39 | 495 | OK |
| Regresión de la batería frente a la v6: mismos indicadores y misma caja, deuda, EBITDA y FCFF mes a mes | 39 | incluidas arriba | OK |
| Pruebas de operación (reales congelados, TC, días de cobro, selector, implantación) | 7 | 291.760 valores iguales a la v6 | OK |
| Proyección rodante de 24 meses (Detalle_24m y Proyeccion_24m) | 11 | 60 | OK |
| Cuadros con la estructura del plan (nuevo en la v7) | 7 | 147 | OK |
| Plantilla vacía | 1 | 0 errores de fórmula | OK |

**Conclusión: la v7 funciona correctamente.** Los cuadros nuevos cuadran con el motor y con Resultados en todos los casos probados, y el motor da exactamente lo mismo que la v6, ya validada en Excel. Las conclusiones financieras de las baterías anteriores siguen vigentes.

## Paso 1 — Regresión frente a la v6

| Hoja | Celdas comparadas | Distintas |
|---|---|---|
| Calculo | 43.492 | 0 |
| Resultados | 1.866 | 0 |
| Proyeccion_24m | 1.032 | 0 |
| Guia_Carga | 249 | 0 |
| Control | 110 | 0 |
| Supuestos, Productos, Costos, Inversion y Reales | 3.735 | 0 |

- **Detalle_24m:** las únicas diferencias son las 9 filas «Celdas con dato en el mes» (1.305 celdas), que ahora están vacías porque los contadores pasaron a Calculo (filas det.TC a det.FC, al final del motor).
- **Inicio:** cambió solo la lista de hojas de la sección 5 (se agregaron los cuadros).
- **Guía y Proyeccion_24m:** leen los contadores del motor con los mismos nombres (`Det_nQ`, `Det_nMP`…), así que dan lo mismo que antes. Las 60 pruebas del detalle lo confirman (paso 4).

## Paso 2 — Batería de 39 escenarios

- **Resultado:** 495 de 495 comprobaciones OK y 0 errores de fórmula en los 39 escenarios (incluye la regresión escenario por escenario frente a la v6).
- **Comparación directa con la corrida de la v6:** 1.625.364 valores (indicadores y filas mensuales del motor de los 39 escenarios) idénticos.

El detalle por escenario y su lectura financiera están en `Pruebas_Escenarios_v6.md` (pasos 2 a 10), que sigue vigente porque los resultados son idénticos.

## Paso 3 — Operación

Los 7 casos (base, 6 meses reales con escenario Vigente y Conservador, TC +10%, días de cobro 150, Conservador e implantación desde el mes 7) dan los mismos valores que en la v6: 291.760 valores comparados, 0 distintos.

## Paso 4 — Proyección rodante de 24 meses

60 de 60 comprobaciones OK, con los mismos 11 casos de la v6: volumen, precio, materia prima, personal, costo fijo, TC, ventana móvil con cierre, escenario Conservador y ventana de 12 meses. La guía (paso 29) sigue marcando «✔ 24 meses detallados», «⚠ 23 de 24 meses» y «○ Usa supuestos anuales» con los contadores ya movidos al motor.

## Paso 5 — Cuadros con la estructura del plan (nuevo)

`probar_cuadros.py` recalcula 7 casos y hace 147 comprobaciones:
- **Casos:**
  - base, en las dos versiones del selector;
  - Conservador, con atraso de 3 meses;
  - 6 meses reales con Conservador;
  - materia prima detallada en Detalle_24m;
  - implantación de 6 meses, en las dos versiones.
- **Resultado:** 147 de 147 OK, sin errores de fórmula.

| Comprobación (en cada caso) | Qué demuestra |
|---|---|
| Cuadro 18: fuentes − usos − caja del motor = 0 en todos los años, y caja al cierre = motor | El origen y aplicación de fondos cierra con la caja |
| Cuadro 15: ventas, EBITDA y utilidad = Resultados; EBITDA año por año = motor | El estado de resultados por año de operación es el del motor |
| Cuadro 10a: ventas − costos totales = utilidad + IRE (cada año); cuadro 11 = cuadro 10a | Los costos están completos y sin duplicaciones |
| Cuadro 19: FCFF descontado = VAN del proyecto; flujo de los socios descontado = VAN de los socios | Los VAN anuales suman los VAN mensuales de Resultados |
| Cuadro 16: capital de trabajo neto al cierre = motor; bruto − proveedores = neto | Saldos al cierre del año correctos |
| Cuadro 20: depreciación por ítem + ajuste = motor; cuadro 13: deuda cancelada en el año 10 | Depreciación y deuda consistentes |
| Cuadros 8 y 10b: los productos suman el total | Apertura por producto completa |
| Filas de ajuste en 0 sin reales ni detalle mensual (9 filas) | Los cuadros reconstruyen el motor con los supuestos anuales |

**Comprobaciones específicas:**

| Caso | Comprobación | Resultado |
|---|---|---|
| Base | Método del plan con los flujos publicados: VAN 4.588.737 y TIR 22,61% (iguales a los publicados) | OK |
| Base | Utilización de gabinetes del año 1 = 80%; precio del año 3 = 275,43; 65 personas en el año 1 (43 + 22) | OK |
| Base | Estructura de costos suma 100%; inversión fija inicial = 655.044; diferencia de EBITDA frente al plan = −1.044.506 | OK |
| Presupuesto original | El selector cambia de versión y, sin reales, da lo mismo que el vigente | OK |
| Conservador | Año 0 sin ventas y con costos de implantación (283.020); año 1 de operación desde el mes 4 (ventas 4.146.750); presupuesto sin cambios | OK |
| 6 meses reales + Conservador | Estado del año 1 = «Real+Proy.» y del año 2 = «Proyección»; fuentes = usos con el ajuste de conciliación de caja | OK |
| Materia prima detallada | El ajuste del cuadro 9 muestra el detalle: 88.200 en el año 1 (12.000 × 7,35), 97.852,50 en el año 2 y 0 en el año 3 | OK |
| Implantación (inicio en el mes 7) | Año 0 = 6 meses sin ventas, con costos e inversión; los años 1-10 suman las ventas del horizonte | OK |

## Paso 6 — Lo que muestran los cuadros frente al plan publicado

Presupuesto original del modelo frente al plan de octubre de 2025 (hoja «Cuadros», sección 3):

| Concepto | Modelo | Plan | Diferencia | Causa principal |
|---|---|---|---|---|
| Inversión fija inicial | 655.044 | 655.044 | 0 | — |
| Capital operativo inicial | 2.388.064 | 2.336.543 | 51.521 | Clientes a precio de venta (plan: al costo) |
| Ventas de los 10 años | 81.222.448 | 81.222.448 | 0 | — |
| EBITDA de los 10 años | 16.309.504 | 17.354.010 | −1.044.506 | C1 nómina al TC 7.300 y C2 materia prima de cajas |
| Utilidad neta de los 10 años | 14.751.246 | 15.945.168 | −1.193.922 | EBITDA y depreciación |
| Caja al cierre del año 10 | 1.637.884 | 1.909.351 | −271.467 | Menor EBITDA, mayor capital de trabajo y dividendos |
| Punto de equilibrio del año 1 (unidades) | 73.373 | 66.042 | 7.331 | Energía y mantenimiento como fijos; C1 |
| VAN — método del plan | 4.177.576 | 4.588.736 | −411.160 | Cifras (EBITDA, capital de trabajo, calendario), no método |
| TIR — método del plan | 23,99% | 22,61% | +1,38 p.p. | La inversión cae en el año 1 y no en el año 0 |
| Payback — método del plan | 8 años | 7 años | +1 | Sigue al VAN |

**Hallazgos nuevos que los cuadros dejan a la vista:**
- **El método del plan está bien replicado.** Aplicado a los flujos publicados, el cuadro 19 da exactamente el VAN y la TIR del plan. Con las cifras del modelo da VAN 4,18 M: la diferencia con el plan es de cifras, no de método.
- **El VAN de 6,99 M del modelo (FCFF) y el de 4,59 M del plan miden cosas distintas.** No hay que compararlos entre sí.
- **Tasa de corte:** el plan pondera solo el aporte inicial (200.000) y el préstamo, y llega a 9,49%. Con todos los aportes previstos (1.139.710) la tasa ponderada sería 11,12%. Conviene actualizar la tasa de corte antes de presentar el VAN (cuadro 19, sección E).
- **Clasificación del personal:**
  - el plan registra a los 22 indirectos (jefaturas, ingeniería, calidad, depósito, compras y administración) como gastos de administración;
  - el modelo separa 17 indirectos de producción y 5 de administración;
  - por eso «costos de producción» y «administración» difieren mucho (+290 mil y −227 mil en el año 1), aunque el total de personal solo difiere por C1.
- **Inconsistencias internas del plan** (la sección 4 de «Cuadros» las detalla):
  - el cuadro 10b asigna 398.170 de flete, frente a 424.000 en el 10a;
  - el «margen bruto %» del 10b es en realidad costo ÷ precio (57,2%; el margen sobre el precio es 42,8%);
  - la depreciación del año 1 es 0 en el 10a y 71.592 en el cuadro 20;
  - en la pág. 28, el año 0 figura con −2.791.587 y saldo final 4.788.736, cuando el cuadro 19 usa −2.991.587 y llega al VAN de 4.588.736.
- **Dividendos:** el modelo paga solo con caja sobre el mínimo, así que en el presupuesto los dividendos sobre la utilidad del año 3 no se pagan. El plan los paga en el año 4 (371.625).

## Conclusión

La v7 está validada técnicamente en LibreOffice. El motor es el mismo que la v6, ya validada en Excel; por eso en Excel alcanza con la ruta corta (pasos 0, 1, 3e, 3f y 11 del prompt v7).

Las cifras siguen siendo las del plan de 2025. Antes de usar los cuadros para decidir, hay que:
- confirmar con los autores del plan el TC de la nómina (C1) y el 1% de la materia prima de cajas (C2);
- cargar los reales desde octubre de 2025;
- actualizar WACC, Ke, TC y tasas.

**Para repetir las pruebas:**

```bash
python3 modelo/plantilla/armar_ecostar.py Modelo_Maquila_ECOSTAR_v7.xlsx versiones/Modelo_Maquila_ECOSTAR_v6_validacion_Excel.xlsx [fila_log.json]
python3 modelo/plantilla/probar_escenarios.py Modelo_Maquila_ECOSTAR_v7.xlsx recalc.py carpeta
MODELO="Ecostar v7" LIBRO=Modelo_Maquila_ECOSTAR_v7.xlsx FECHA=28/10/2026 \
  python3 modelo/plantilla/informe_escenarios.py carpeta/escenarios.pkl Pruebas_Escenarios_ECOSTAR_v7.xlsx [referencia.pkl]
python3 modelo/plantilla/probar_plantilla.py Modelo_Maquila_ECOSTAR_v7.xlsx recalc.py carpeta_operacion
python3 modelo/plantilla/probar_detalle.py   Modelo_Maquila_ECOSTAR_v7.xlsx recalc.py carpeta_detalle
python3 modelo/plantilla/probar_cuadros.py   Modelo_Maquila_ECOSTAR_v7.xlsx recalc.py carpeta_cuadros
```
