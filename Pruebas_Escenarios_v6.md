# Ecostar v6 (25 productos y proyección rodante de 24 meses): pruebas paso a paso

Fecha: 27/10/2026. Repite todas las pruebas de las versiones anteriores sobre la v6 y agrega las de la proyección de 24 meses.

- **Libro probado:** `Modelo_Maquila_ECOSTAR_v6.xlsx`: 25 productos, «Guia_Carga», «Detalle_24m» y «Proyeccion_24m», más «Claude Log» y «Esp_25Productos».
- **Resultados:** `Pruebas_Escenarios_ECOSTAR_v6.xlsx`. Contiene el resumen, la evolución anual por escenario y las verificaciones.
- **Recálculo:** LibreOffice 24.2. En Excel lo valida Claude en Excel con `Prompt_Claude_Excel_ECOSTAR_v6.md`.
- **Unidades:** USD, supuestos del plan de octubre de 2025, sin datos reales.

## Resumen

| Bloque de pruebas | Casos | Comprobaciones | Resultado |
|---|---|---|---|
| Batería de escenarios (financiamiento, cobranza, costos, ventas, plan desfavorable, plazos, corto vs largo plazo, TC y estructura de 25 productos) | 39 | 495 | OK |
| Regresión frente a la v4: mismos indicadores y misma caja, deuda, EBITDA y FCFF mes a mes | 39 | incluidas arriba | OK |
| Pruebas de operación (reales congelados, TC, días de cobro, selector, implantación) | 7 | todas | OK |
| Proyección rodante de 24 meses (Detalle_24m y Proyeccion_24m) | 11 | 60 | OK |
| Regresión celda por celda de Calculo frente a la v5, con el detalle vacío | 1 | 0 celdas distintas | OK |

**Conclusión: la v6 funciona correctamente.** En ningún caso hubo errores de fórmula y el balance cuadra en todos los meses. La proyección de 24 meses no altera ningún resultado cuando no se usa. Cuando se usa, cambia exactamente los meses de la ventana y por el monto esperado. Las conclusiones financieras de las baterías anteriores siguen vigentes.

## Paso 1 — Integridad del libro base

| Verificación | Resultado |
|---|---|
| Indicadores de Resultados (C6:D30) frente a la v5, la v4 y la versión de 12 productos | Idénticos (25 de 25) |
| Ventas / EBITDA / VAN proyecto / VAN socios | 81.222.448 / 16.309.504 / 6.989.922 / 3.695.678 |
| Aporte adicional / brecha máxima / meses con brecha | 311.740 / 373.701 / 12 |
| Balance de control (presupuesto / vigente) | 0 / 0 |
| Controles en REVISAR | 2: C13 (brecha de caja) y C14 (sin reales) |
| Criterios de texto con decimales (`">0.5"`) en todo el libro | 0 |
| Nombres definidos y filas frente a la hoja Esp_25Productos | Coinciden |

## Paso 2 — Estructura de 25 productos

| Id | Escenario | Ventas 10 años | EBITDA 10 años | VAN proyecto | VAN socios | Aporte adicional | Meses bajo caja mín. |
|---|---|---|---|---|---|---|---|
| E00 | Plan base (Oct-2025) | 81,22 M | 16,31 M | 6,99 M | 3,70 M | 0,31 M | 12 |
| X1 | Partición en los 25 espacios | 81,22 M | 16,31 M | 6,99 M | 3,70 M | 0,31 M | 12 |
| X2 | Productos nuevos 3, 13 y 25 | 83,42 M | 17,40 M | 7,64 M | 4,13 M | 0,20 M | 6 |
| X3 | Producto con volumen y sin precio | 81,22 M | 16,25 M | 6,95 M | 3,67 M | 0,32 M | 12 |

- **X1, partición:**
  - gabinetes repartidos en los productos 1-13 y cajas en los 14-25, con el mismo precio, costos y volumen total;
  - todos los indicadores y la caja mes a mes son iguales a la base, sin alertas de capacidad ni de precio;
  - prueba que los 25 espacios entran en todos los cálculos.
- **X2, productos nuevos 3, 13 y 25, cargados solo en vigente** (datos de prueba, no comerciales):
  - las ventas suben 2.197.064,63 y los materiales 933.279,30, exactamente lo esperado;
  - el presupuesto original no cambia;
  - el producto 25, último del rango, sí consume materiales.
- **X3, producto con volumen y sin precio:** se activa la alerta C11. No suma ventas y sí suma costos: por eso el VAN baja.
- **V4, volumen +10%:** la alerta C10 marca 16 años-producto sobre la capacidad.

## Paso 3 — Operación: reales, tipo de cambio, días, selector e implantación

| Prueba | Resultado |
|---|---|
| 6 meses reales cargados, cambio de Vigente a Conservador | Los meses 1-6 no cambian: ventas 254.000 por mes, la caja del mes 6 = saldo real 1.500.000 (ajuste de conciliación −68.081). Desde el mes 7: 404.167 (Vigente) frente a 345.562 (Conservador). El presupuesto no cambia. |
| TC vigente 8.030 (+10%) | Personal × 0,9091 (= 7.300 / 8.030). Materiales y ventas sin cambio. |
| Días de cobro 150 | EBITDA igual. El capital de trabajo máximo pasa de 4,50 M a 5,26 M. |
| Selector Conservador | La operación arranca en el mes 4. El CAPEX pendiente sube a 830.548 (+10%). |
| Implantación (inicio en el mes 7) | Sin ventas en los meses 1-6, personal parcial (36.744 frente a 53.055) y depreciación desde el mes 7. |

## Paso 4 — Financiamiento y cobranza

| Id | Escenario | VAN socios | Necesidad de fondos | Aporte adicional | Meses bajo caja mín. | Intereses | VAN proyecto |
|---|---|---|---|---|---|---|---|
| E00 | Plan base (Oct-2025) | 3,70 M | 3,21 M | 0,31 M | 12 | 0,81 M | 6,99 M |
| F1 | Más financiamiento | 3,59 M | 3,21 M | 0,00 M | 1 | 1,11 M | 6,99 M |
| F2 | Menos financiamiento | 3,68 M | 3,21 M | 0,82 M | 41 | 0,52 M | 6,99 M |
| F3 | Cobro a 180 días | 2,93 M | 4,22 M | 1,63 M | 57 | 0,81 M | 6,23 M |
| F4 | Cobro al contado | 3,90 M | 1,41 M | 0,00 M | 0 | 0,81 M | 8,51 M |
| F5 | Más financiamiento + cobro 180 días | 2,95 M | 4,22 M | 1,40 M | 48 | 1,11 M | 6,23 M |
| F6 | Menos financiamiento + contado | 3,98 M | 1,41 M | 0,00 M | 0 | 0,52 M | 8,51 M |
| F7 | Más financiamiento + contado | 3,81 M | 1,41 M | 0,00 M | 0 | 1,11 M | 8,51 M |
| F8 | Menos financiamiento + cobro 180 días | 2,81 M | 4,22 M | 2,06 M | 55 | 0,52 M | 6,23 M |

- El plazo de cobro define cuánta plata necesita el negocio: de 1,41 M (contado) a 4,22 M (180 días).
- Más deuda tapa el bache, pero cuesta intereses.
- La combinación más eficiente es menos deuda con cobro al contado (F6).
- El financiamiento no cambia el VAN del proyecto, pero sí la caja y el VAN de los socios.

## Paso 5 — Costos

| Id | Escenario | EBITDA 10 años | VAN proyecto | TIR proyecto | VAN socios | Aporte adicional | Meses bajo caja mín. |
|---|---|---|---|---|---|---|---|
| E00 | Plan base (Oct-2025) | 16,31 M | 6,99 M | 34,9% | 3,70 M | 0,31 M | 12 |
| C1 | Costos variables +10% | 11,13 M | 3,62 M | 21,8% | 1,33 M | 1,60 M | 55 |
| C2 | Costos variables +20% | 5,94 M | 0,25 M | 10,3% | −1,13 M | 3,19 M | 104 |
| C3 | Costos variables +30% | 0,76 M | −3,12 M | 0,0% | −3,73 M | 5,69 M | 117 |
| C4 | Costos que no se cubren con las ventas | −5,25 M | −7,02 M | −10,7% | −6,75 M | 10,60 M | 118 |

- Cada 1% de costo variable resta unos 337 mil de VAN, y con +20,7% el VAN del proyecto se anula.
- En C4 el EBITDA es negativo y faltan 10,6 M. El modelo lo muestra sin inventar fondos.

## Paso 6 — Ventas

| Id | Escenario | Ventas 10 años | VAN proyecto | VAN socios | Necesidad de fondos | Aporte adicional | Recupero (años) |
|---|---|---|---|---|---|---|---|
| V1 | Volumen −10% | 73,10 M | 5,42 M | 2,69 M | 3,08 M | 0,43 M | 4,50 |
| V2 | Volumen −20% | 64,98 M | 3,86 M | 1,66 M | 2,96 M | 0,66 M | 5,08 |
| V3 | Volumen −30% | 56,86 M | 2,29 M | 0,63 M | 2,94 M | 0,88 M | 6,02 |
| V4 | Vendo más: volumen +10% | 89,34 M | 8,56 M | 4,71 M | 3,33 M | 0,29 M | 3,84 |
| V5 | Vendo más dentro de la capacidad | 83,55 M | 7,60 M | 4,11 M | 3,59 M | 0,38 M | 3,67 |

- El volumen solo resiste más que el precio y el costo: con −30% el VAN sigue siendo positivo.
- Vender más crea valor, pero primero consume caja.
- V4 supera la capacidad y no incluye el CAPEX de ampliación, así que su resultado está sobrestimado.

## Paso 7 — Si el plan no sale como se esperaba

| Id | Escenario | Ventas 10 años | VAN proyecto | TIR proyecto | VAN socios | Aporte adicional | Recupero (años) |
|---|---|---|---|---|---|---|---|
| P1 | Desfavorable combinado | 50,95 M | −0,11 M | 9,0% | −1,16 M | 2,64 M | 9,59 |
| P2 | Selector «Conservador» del modelo | 67,47 M | −0,49 M | 7,9% | −1,72 M | 3,80 M | 9,85 |
| P3 | Selector «Optimista» del modelo | 86,99 M | 10,47 M | 57,7% | 5,39 M | 0,00 M | 2,77 |

- La combinación (volumen −30%, precio −5%, atraso de 6 meses y CAPEX +15%) deja la TIR debajo del WACC (9,49%).
- El selector Conservador lleva el VAN a negativo.

## Paso 8 — ¿En cuánto tiempo devuelvo el préstamo?

| Id | Escenario | Cancela (años) | DSCR mín. (año 2+) | Aporte adicional | Intereses | VAN socios |
|---|---|---|---|---|---|---|
| E00 | Plan base (Oct-2025) | 5,00 | 0,40 | 0,31 M | 0,81 M | 3,70 M |
| L1 | LP a 3 años (6 meses de gracia + 30 cuotas) | 3,00 | 0,25 | 1,36 M | 0,51 M | 3,64 M |
| L2 | LP a 7 años (6 + 78 cuotas) | 7,00 | 0,53 | 0,02 M | 1,13 M | 3,70 M |
| L3 | LP a 8 años (12 + 84 cuotas) | 8,00 | 0,55 | 0,00 M | 1,36 M | 3,65 M |
| L4 | Repago acelerado con todo el excedente | 4,00 | 1,56 | 0,00 M | 0,64 M | 3,76 M |

- **Capacidad real:** unos 4 años, si todo el excedente va a deuda (L4).
- **Cronograma fijo 6 + 54:** choca con los años 2 y 3, con DSCR de 0,40 y 0,96.
- **Plazos más largos:** a 7-8 años desaparece el faltante, con 0,32 a 0,54 M más de intereses.
- **A 3 años:** no es viable.

## Paso 9 — Corto plazo, largo plazo o ambos

Tasa del largo plazo 9% (plan); tasa del corto plazo 10% ilustrativa, con sensibilidad a 7,99% y 11,24%.

| Id | Escenario | VAN socios | Intereses | Uso máx. línea | Aporte adicional | Cancela (años) |
|---|---|---|---|---|---|---|
| E00 | Plan base (Oct-2025) | 3,70 M | 0,81 M | 0,00 M | 0,31 M | 5,00 |
| K1 | Solo LP, ampliado para cubrir la brecha | 3,65 M | 0,95 M | 0,00 M | 0,10 M | 5,00 |
| K2 | LP + préstamo CP de 12 meses | 3,69 M | 0,84 M | 0,00 M | 0,23 M | 5,00 |
| K3 | LP + línea rotativa a la par | 3,70 M | 0,83 M | 0,37 M | 0,00 M | 5,00 |
| K4 | Solo CP: línea rotativa renovada | 3,71 M | 0,73 M | 2,36 M | 0,00 M | 4,08 |
| K5 | Solo CP sin renovación | 3,65 M | 0,46 M | 2,36 M | 2,19 M | 2,08 |
| K6 | LP para CAPEX + CP para capital de trabajo | 3,74 M | 0,73 M | 1,72 M | 0,00 M | 5,00 |
| K7 | Solo CP renovado con tasa 7,99% | 3,81 M | 0,55 M | 2,31 M | 0,00 M | 3,92 |
| K8 | Solo CP renovado con tasa 11,24% | 3,65 M | 0,86 M | 2,39 M | 0,00 M | 4,17 |

- **Mejor relación entre costo y riesgo:** largo plazo para el CAPEX y el capital de trabajo permanente, más una línea rotativa a la par. Con K3, la línea se usa hasta 374 mil y elimina el faltante con 21 mil más de intereses.
- **Solo corto plazo:** depende de que el banco renueve la línea. En K5, si no la renueva en el mes 24, faltan 2,19 M de golpe.

## Paso 10 — Tipo de cambio

| Id | Escenario | EBITDA 10 años | VAN proyecto | VAN socios | Aporte adicional | Meses bajo caja mín. |
|---|---|---|---|---|---|---|
| T1 | Sube el TC: Gs 8.760 (+20%) | 17,84 M | 7,99 M | 4,36 M | 0,01 M | 2 |
| E00 | Plan base (Oct-2025) | 16,31 M | 6,99 M | 3,70 M | 0,31 M | 12 |
| T2 | Baja el TC: Gs 6.145 (−16%) | 14,59 M | 5,86 M | 2,93 M | 0,71 M | 35 |
| T3 | Baja el TC: Gs 5.840 (−20%) | 14,02 M | 5,49 M | 2,66 M | 0,87 M | 39 |

- Ventas y materiales están en USD; la nómina y parte de los fijos, en guaraníes.
- Un guaraní más fuerte resta unos 71 a 75 mil de VAN por cada 1%.

## Paso 11 — Regresión

| Verificación | Escenarios | Resultado |
|---|---|---|
| Mismos indicadores que la corrida de la v4 (tolerancia USD 0,01) | 39 | OK |
| Caja, deuda, EBITDA y FCFF mes a mes iguales a la v4 | 39 | OK |
| Calculo celda por celda igual a la v5 (detalle vacío) | todas las celdas | OK (0 distintas) |

La v4 a su vez reproducía exactamente la corrida de 12 productos, así que la cadena v2 → v4 → v5 → v6 no cambió ningún resultado.

## Paso 12 — Proyección rodante de 24 meses (nuevo en la v6)

Pruebas con datos de prueba (no comerciales), con `probar_detalle.py`:

| Prueba | Qué se cargó en Detalle_24m | Resultado esperado y obtenido |
|---|---|---|
| Volumen | Gabinetes 1.100 y cajas 6.500 por mes, meses 1-24 | Unidades de los meses 1-24 = detalle. Mes 25 = supuesto anual (1.500). Ventas del mes 13 = 1.100 × 272,70 (el precio sigue su índice anual) |
| Precio | Gabinetes a USD 280 en los meses 1-12 | Ventas = 1.000 × 280 en los meses 1-12. En el mes 13 vuelve al anual (1.250 × 272,70) |
| Materia prima | Gabinetes a USD 90 por unidad, meses 1-24 | La materia prima sube exactamente 1.000 × (90 − 82,65) = 7.350 en el mes 1. Sin cambio en el mes 25 |
| Personal | Operador láser: 6 personas a Gs 5.000.000, meses 1-24 | El personal directo sube lo esperado (salario y alimentación, al TC y con cargas). Personas directas = 45 |
| Costo fijo | Energía a Gs 20.000.000 por mes | Los fijos de producción suben Gs 6.000.000 / 7.300 por mes |
| TC | Gs 8.000, meses 1-24 | TC = 8.000 en los meses 1-24 y 7.300 en el mes 25. Personal administrativo × 7.300 / 8.000 |
| Ventana móvil | 6 meses reales cerrados y volumen en los meses 1-36 | Meses 1-6 = reales (600). Meses 7-30 = detalle (1.100). Meses 31-36 = anual: el detalle fuera de la ventana no se usa |
| Escenario | Conservador y volumen detallado | Meses 1-3 sin ventas (atraso). Meses 4-24 = 1.100 × 0,90 |
| Ventana de 12 meses | Inicio!C14 = 12 y volumen en los meses 1-24 | Meses 1-12 = detalle; meses 13-24 = anual |
| Guía de carga | Ventana completa, incompleta y vacía | «✔ 24 meses detallados», «⚠ 23 de 24 meses: completar hasta el mes 24», «○ Usa supuestos anuales» |
| Proyeccion_24m | Con 6 meses cerrados | Arranca en el mes 7 y termina en el 30. Origen de volúmenes = «Detalle». Ventas, totales por año y caja coinciden con Calculo |

En todos los casos el presupuesto original no cambió y el balance quedó en 0.

## Conclusión

La v6 está validada técnicamente y las recomendaciones de la batería anterior siguen vigentes:

- Largo plazo para la inversión y el capital de trabajo permanente, más una línea rotativa comprometida de 0,5 a 1,0 M para los años 2 a 4.
- Cuidar, en este orden, el costo unitario, los días de cobro y el tipo de cambio.

Las cifras siguen siendo del plan de 2025 y no describen la situación actual. Falta cargar los reales desde octubre de 2025 y actualizar WACC, Ke, TC y tasas con ofertas bancarias.

Supuestos, fuentes de tasas y limitaciones: ver `Pruebas_Escenarios_v2.md` (sin cambios) y la hoja «Supuestos» de `Pruebas_Escenarios_ECOSTAR_v6.xlsx`.

**Para repetir las pruebas:**

```bash
python3 modelo/plantilla/probar_escenarios.py Modelo_Maquila_ECOSTAR_v6.xlsx recalc.py carpeta
MODELO="Ecostar v6" LIBRO=Modelo_Maquila_ECOSTAR_v6.xlsx FECHA=27/10/2026 \
  python3 modelo/plantilla/informe_escenarios.py carpeta/escenarios.pkl Pruebas_Escenarios_ECOSTAR_v6.xlsx [referencia.pkl]
python3 modelo/plantilla/probar_plantilla.py Modelo_Maquila_ECOSTAR_v6.xlsx recalc.py carpeta_operacion
python3 modelo/plantilla/probar_detalle.py   Modelo_Maquila_ECOSTAR_v6.xlsx recalc.py carpeta_detalle
```
