# Revisión del modelo «Plan de Negocios – Ecostar 05-11-2026»

Fecha de revisión: 13/11/2026 · Archivo revisado: `Plan_de_Negocios_-_Ecostar_05-11-2026.xlsx` (27 hojas, ≈ 31.000 fórmulas, 88 turnos en «Claude Log»).

---

## 1. Conclusión

1. **La aritmética del libro es correcta.** No hay errores de fórmula. Los cuadres internos cierran (flujo, estado de resultados, caja, préstamos). VAN, TIR, payback, valor residual, cobranzas y cronogramas dan lo mismo al recalcularlos de forma independiente. Las cifras del plan aprobado también están bien cargadas: reproduje exactamente su VAN (4.588.737) y su TIR (22,61%).
2. **El resultado depende de supuestos que todavía no tienen respaldo.** El principal es el **precio de la caja: USD 38,65 en la proyección contra USD 23,69 facturado en octubre-26**. Con el precio facturado, la caja no cubre ni su costo variable (USD 23,79), y el VAN pasa de **+6,1 M a ≈ −1,9 M**. El VAN es cero con un precio de la caja de **≈ USD 27,2**.
3. **Tres puntos sobrestiman el resultado o lo presentan mejor de lo que es:**
   - los insumos de planta (≈ USD 264 mil/año) se cargan como costo, pero en la caja quedan como stock que se recupera en 2035;
   - el DSCR deja afuera el préstamo de corto plazo (1,41x informado contra 0,37x real en 2028);
   - el VAN se compara contra el del plan con financiamiento (4,59 M), cuando en la misma base el plan da **7,32 M**: la proyección actual es **−1,23 M peor que el plan**, no +1,5 M mejor.
4. **El financiamiento no alcanza.** Con la estructura cargada (préstamo de largo plazo de 1,1 M y préstamos de corto plazo de 4,5 M renovados cada año), la caja es negativa en ene-28 (−407 mil), nov-28 (−345 mil), cierre 2029 (−111 mil) y cierre 2031 (−324 mil). Además, depende de que el banco renueve el corto plazo todos los años.

---

## 2. Cómo se validó (10 capas)

| # | Validación | Resultado |
|---|---|---|
| 1 | Errores de fórmula (#REF!, #DIV/0!, #N/A, #VALUE!) en las 27 hojas | ✔ 0 errores |
| 2 | Recálculo completo del libro en un motor independiente | ⚠ No concluyente: PRÉSTAMOS usa `LET`, que ese motor no resuelve. Por eso los puntos 3 a 10 se recalcularon por separado |
| 3 | Cuadres del FLUJO DE CAJA en las 44 columnas: suma de operación, inversión y financiamiento, variación de caja, arrastre mes a mes y totales anuales = suma de los meses | ✔ 0 diferencias |
| 4 | Cuadres del ESTADO DE RESULTADOS (EBITDA, EBIT, resultado antes de impuestos, resultado neto) y cruce de intereses e IVA con el flujo | ✔ 0 diferencias |
| 5 | Préstamos: cuota francesa del préstamo de largo plazo (PMT 9%/12, 60 cuotas = 22.834,19), intereses de los préstamos de corto plazo (1,2 M × 9% × 10/12 = 90.000; 1,81 M × 3/12 = 40.725; 3,0 M × 2/12 = 45.000; 4,5 M × 12/12 = 405.000), IVA del 10% | ✔ |
| 6 | Cobranzas con 120 días de plazo (desfase de 4 meses mes a mes) | ✔ |
| 7 | VAN, TIR, VAN y TIR del accionista, payback simple (5,19 años) y descontado (6,03 años), tasa de corte (10,13%) | ✔ iguales al libro |
| 8 | Valor residual 2035 (Gs 5.930 M depreciables − Gs 5.360 M depreciados = Gs 570 M ÷ TC 7.685 = 74.143) | ✔ coincide con una anualidad de depreciación porque los activos a 10 años llevan 9 años depreciados; no es un error |
| 9 | Plan aprobado: ventas, EBITDA, resultado neto, volúmenes, dotación, CAPEX 655.044, préstamo, intereses + IVA 813 mil, dividendos, depreciación, tasa del 9,49%, flujos y VAN/TIR (4.588.737 y 22,61%; sin financiamiento: 7.324.740 y 33,82%) | ✔ |
| 10 | Coherencia entre hojas: compras contra consumo y stock, costos de ESTRUCTURA contra pagos del FLUJO, supuestos contra datos reales del Informe de CIE RV2 | ⚠ Ver hallazgos |

---

## 3. Hallazgos críticos (cambian la decisión)

### 3.1 Precio de la caja sin respaldo — **impacto VAN ≈ −8 M**

| | Caja | Armario |
|---|---|---|
| Precio proyectado (COSTOS VARIABLES) | **38,65** | 260,75 |
| Precio del plan, año 3 | 23,46 | 275,43 |
| Precio facturado real | **22,76** (ago-26, 392 cajas) · **23,69** (oct-26, 1.241 cajas) | sin ventas |
| Costo variable unitario | 23,79 | 154,11 |
| Margen de contribución a precio proyectado | 38% | 41% |
| Margen a precio facturado | **−0,4%** | — |

| Precio de la caja | VAN del proyecto (aprox.) |
|---|---|
| 38,65 (actual) | 6,10 M |
| 34,00 | 3,6 M |
| 30,00 | 1,5 M |
| **≈ 27,2** | **0 (equilibrio)** |
| 23,69 (facturado oct-26) | ≈ −1,9 M |

Las cajas son el 46% de la facturación a capacidad plena (USD 4,64 M de 10,11 M).

**Qué hacer:** documentar el origen de USD 38,65 (contrato, lista de precios o nueva especificación: la caja pasó de 13 kg a 19,89 kg de chapa). Si no hay respaldo firmado, usar el precio facturado y presentar 38,65 como escenario.

### 3.2 Insumos de planta: costo en resultados sin salida de caja propia — **impacto VAN −0,8 a −1,2 M**

- ESTRUCTURA DE COSTOS suma, además de los insumos por unidad de COSTOS VARIABLES, «Insumos» y «Consumibles» de RE_PROJEÇÃO. Se calculan a USD 1,87 por unidad producida: **≈ USD 264 mil/año** desde 2030.
- El FLUJO no los incluye en los costos fijos. El consumo con el que se calcula el stock (fila 112) es el de COSTOS VARIABLES, sin estos insumos.
- Resultado: las compras cargadas (6,36 M/año) superan al consumo (6,09 M) en ≈ 272 mil/año, y el **stock sube de 1,67 M (2030) a 3,03 M (2035)**: 182 días de consumo contra 120 de objetivo. Ese stock vuelve como «recuperación de capital de trabajo» en 2035 (6,40 M en total).
- Hay dos lecturas, y las dos sobrestiman el flujo:
  - si las compras incluyen esos insumos, el stock está inflado en ≈ 2,2 M al 2035 (efecto en el VAN ≈ −0,84 M);
  - si no los incluyen, su pago no figura en ningún lado (≈ −1,2 M).
- Además, el plan de compras está desbalanceado por categoría. La chapa importada (5,95 M/año) supera el consumo de chapa (4,30 M), mientras que los insumos comprados (0,36 M) y las tintas y packaging (0) quedan por debajo de su consumo (1,02 M y 0,77 M).

**Qué hacer:**
- confirmar con Producción si los USD 1,87 por unidad duplican los insumos por unidad;
- incluirlos en el consumo que descuenta el stock;
- conciliar el plan de compras por categoría en kg de chapa: la necesidad a capacidad plena es ≈ 5,2 millones de kg/año (120.000 × 19,89 + 21.000 × 134).

### 3.3 DSCR (cobertura del servicio de deuda) mal definido

INDICADORES fila 41 toma solo el préstamo de largo plazo más el IVA de **todos** los intereses. Deja afuera los intereses del préstamo de corto plazo (405 mil/año) y la reducción neta de ese préstamo (de 4,5 a 4,0, 3,0, 2,0 y 0 M).

| Año | Flujo operativo | DSCR informado | DSCR con intereses del corto plazo | **DSCR con servicio total** (lo que verá el banco) |
|---|---|---|---|---|
| 2027 | (1.078.947) | (3,33) | negativo | negativo |
| 2028 | 452.767 | 1,41 | 0,62 | **0,37** |
| 2029 | 1.522.844 | 4,83 | 2,25 | **0,91** |
| 2030 | 1.720.916 | 5,65 | 3,00 | **1,09** |
| 2031 | 2.090.810 | 7,74 | 4,64 | **0,85** |

El servicio total coincide con el «flujo neto de financiamiento» del FLUJO, sin dividendos. Por eso la caja se vuelve negativa justamente en 2029 y 2031.

### 3.4 Comparación del VAN con el plan en bases distintas

- INDICADORES (C29 contra D29) y COMPARACIÓN PLAN (fila 56) ponen el **VAN del proyecto sin financiamiento (6,10 M)** junto al **VAN del plan con financiamiento (4,59 M)**, lo que muestra «+1,51 M (+33%)».
- En la misma base, el plan da **7,32 M y TIR 33,8%** (ya está calculado en COMPARACIÓN DETALLE, filas 153-154). **La proyección actual es 1,23 M peor (−17%) y su TIR es 3,5 puntos menor.** El puente de COMPARACIÓN DETALLE (sección 9) está bien hecho y concilia.
- La **TIR del accionista de 65%** no es comparable con la TIR de 22,6% del plan, que no es la del accionista. Además, sale de un aporte de USD 200 mil contra USD 5,6 M de deuda. Es un efecto del apalancamiento: no la presentaría como indicador de decisión.

**Qué hacer:** en INDICADORES y COMPARACIÓN PLAN, comparar contra 7.324.740 y 33,82%. Dejar 4,59 M y 22,6% solo como dato reportado por el plan.

### 3.5 Financiamiento: déficit de caja y riesgo de renovación

- Máximo déficit acumulado antes del financiamiento: **USD 5,42 M (ago-27)**, contra 2,99 M en el plan.
- Préstamos 5 a 9: se llaman «LP capital de trabajo», pero están cargados como **CORTO PLAZO a 12 meses con pago único** de 4,5 / 4,5 / 4,0 / 3,0 / 2,0 M. La viabilidad depende de que el banco renueve cada noviembre.
- Caja negativa: ene-28 (−407 mil), feb-28 (−59 mil), nov-28 (−345 mil), dic-29 (−111 mil), dic-31 (−324 mil). Debajo del mínimo: 2028 a 2031.
- Desde 2029 el flujo es anual. Los meses del vencimiento y la renovación de noviembre no se ven, y puede haber faltantes dentro del año que hoy no aparecen.

**Qué hacer:**
- reemplazar los tramos renovables por un préstamo de mediano plazo comprometido (por ejemplo, 5 a 7 años con gracia hasta 2028) o una línea de capital de trabajo comprometida;
- mantener el flujo mensual hasta que se cancele la deuda (2031).

---

## 4. Errores y diferencias de cálculo (corregir)

| # | Hoja / celda | Hallazgo | Impacto | Corrección |
|---|---|---|---|---|
| 1 | FLUJO DE CAJA **C52** | Desembolso de corto plazo del año 0 **escrito a mano en 1.250.000**; el préstamo 1 es de 1.200.000 (PRÉSTAMOS E7). Los 50 mil de más nunca se devuelven | +USD 50 mil de caja en todos los períodos | Usar la misma fórmula SUMIFS que el resto de la fila |
| 2 | PRÉSTAMOS 1 a 3 | Intereses de la deuda de CIE calculados sobre el capital total y el plazo completo: 175.725 + IVA = 193.298. El Informe de CIE RV2 cobra **120.759** (implantación 36.279 sobre el capital efectivamente usado + préstamos 84.480) | +USD 72,5 mil de costo en 2026 | Alinear con lo que CIE va a facturar (o con lo que apruebe Energy Capital) |
| 3 | PRÉSTAMOS 1 | El préstamo de 1,2 M entra en ago-25, pero en 2025 se gastan solo 191 mil: queda caja ociosa de 1,06 M a dic-25 que en la realidad no existió | Distorsiona la caja de 2025-26 y los intereses | Modelar el fondeo de CIE como aportes mensuales iguales al déficit |
| 4 | Aportes (FLUJO fila 77) | Aporte en efectivo de USD 200 mil. El capital real es Gs 1.311 M: CIE Gs 643 M (**Gs 642 M en bienes**, que ya están dentro de la implantación, más Gs 1 M en efectivo) y Energy Capital Gs 668 M (≈ USD 113 mil, pendiente) | La caja está ≈ USD 86 mil por encima | Cargar como efectivo solo lo de Energy Capital, en el mes real de integración |
| 5 | PARÁMETROS C38-C40; ESTRUCTURA filas 15, 19, 40, 52 | Flete, despachos, seguro de caución y alimentación en **0**. El plan los incluye: alimentación ≈ USD 49 mil/año con 43 personas y seguro de caución 0,1% ≈ USD 10 mil/año | Costos posiblemente omitidos ≈ USD 60 mil/año, más flete y despacho si la chapa no está valorizada puesta en planta | Confirmar si la chapa ya incluye CIF y despacho; cargar alimentación y seguro de caución (obligatorio en maquila) |
| 6 | INDICADORES fila 8 | Tasa de corte ponderada con el aporte (200 mil) y solo el préstamo de largo plazo (1,1 M). El corto plazo (4,5 M), que en la práctica es deuda permanente, queda afuera | 10,13% contra ≈ 9,3% con toda la deuda | Definir una estructura objetivo o usar la tasa del plan (9,49%) para comparar |
| 7 | COMPARACIÓN DETALLE 1, fila 11 | En el plan, la chapa de la caja se ajusta 1 año (7,48) y el precio 2 años (23,46): criterio distinto entre productos | Menor (USD 0,08/u) | Ya estaba observado en la conciliación anterior; dejar la nota |
| 8 | RE_DEPRECIACION | La depreciación se calcula en Gs y se convierte con un TC creciente, así que baja en USD aunque los equipos son importados en USD | Menor: solo resultado y valor residual | Opcional: depreciar en USD los bienes pagados en USD |

---

## 5. Supuestos a validar (no son errores, pero mueven el resultado)

1. **Mano de obra directa fija en 24 personas de 2026 a 2035**, mientras el volumen pasa de 4.241 a 141.000 unidades. El plan prevé 43 a 52 personas. Implica 5.875 unidades por operario al año contra 2.712 en el plan: **2,2 veces la productividad del plan**. Hay que validarlo con Producción (capacidad por puesto y turnos). Con la dotación del plan, el costo sube ≈ USD 125 mil/año.
2. **Costo medio por persona de USD 14,1 mil contra 9,6 mil en el plan (+47%).** Una parte se explica por el TC (6.067 contra 7.300); el resto, por los salarios reales.
3. **IVA no recuperable de USD 63 a 87 mil/año.** Es un criterio prudente. Si Impuestos confirma que se recupera como exportador, el VAN sube ≈ USD 0,4 a 0,5 M.
4. **Precios constantes en USD en 10 años**, mientras los salarios suben 3,5% y el alquiler 5%. Es prudente, pero conviene explicitarlo (el plan suponía +1% anual).
5. **Gastos preoperativos capitalizados** (alquiler del galpón 124 mil, sueldos 48 mil, insumos de puesta a punto 40 mil, gastos administrativos 17 mil, EPIs 4 mil: ≈ USD 232 mil) dentro de «Adecuaciones de infraestructura» y depreciados a 10 años. El tratamiento contable y tributario tiene que validarlo el contador.
6. **Comparación de la inversión con el plan.** Los 1.107.496 (con IVA y preoperativos) contra 655.044 no son comparables. El CAPEX real comparable es ≈ USD 776 mil (maquinaria 707 mil, adecuaciones reales 44 mil, mobiliario 25 mil): **+18% contra el plan**. Los preoperativos (232 mil) y el IVA (101 mil) van aparte.
7. **Gastos de funcionamiento de CIE (7%) e intereses al 9%** incluidos: siguen pendientes de aprobación por Energy Capital.

---

## 6. Comparación con el plan: lo que está bien

- Ventas, EBITDA, resultado neto, dividendos y volúmenes del plan: correctos y coherentes entre sí (ventas − EBITDA − gastos de administración = costos de producción).
- Flujo del plan con y sin financiamiento: reproduce el VAN y la TIR publicados.
- Dotación (43/22, 46, 49, 52), CAPEX 655.044 y depreciación (50.717 + 8.700 + 12.176 = 71.592): correctos.
- Servicio de deuda del plan (2,79 M, 6 meses de gracia + 54 cuotas, intereses + IVA 813 mil) y tasa de corte del 9,49%: correctos.
- Puente del VAN de 7,32 M a 6,10 M (tasa, ingresos, materia prima, mano de obra, otros costos, capital de trabajo, inversión, valor residual): concilia.

**Lo que hay que corregir en la comparación:**
- el VAN y la TIR de referencia en INDICADORES y COMPARACIÓN PLAN (sección 3.4);
- la inversión comparable (sección 5.6);
- el DSCR (sección 3.3).

---

## 7. Mejoras de estructura sugeridas

1. **Hoja de escenarios** con selectores de precio de la caja, volumen, MOD, TC y días de cobro. Hoy el VAN de 6,1 M es un único escenario.
2. **Flujo mensual hasta 2031**, mientras haya deuda.
3. **Control de stock por categoría** (chapa en kg, insumos, tintas, packaging) con días de cobertura.
4. **Indicadores bancarios**: DSCR con el servicio total, deuda / EBITDA, cobertura de intereses y caja mínima del año (no solo la de diciembre).
5. Quitar del libro de trabajo las hojas sin uso (Hoja1, Hoja2, CONTROL SIN USO) o pasarlas a un anexo.
6. Separar el **balance de apertura de Ecostar al 31/10/2026** (activo fijo, stock, deuda con CIE, patrimonio) del flujo histórico de CIE.

---

## 8. Acciones prioritarias

| Prioridad | Acción | Responsable sugerido |
|---|---|---|
| 1 | Respaldar el precio de la caja (38,65) o volver al facturado y recalcular | Comercial / Gerencia |
| 2 | Resolver la doble lectura de los insumos de planta y conciliar el plan de compras por categoría | Producción / Compras / Planificación |
| 3 | Corregir el DSCR y la comparación del VAN con el plan | Planificación |
| 4 | Corregir C52 (50 mil), los intereses de CIE y el aporte en efectivo | Planificación |
| 5 | Redefinir el financiamiento (mediano plazo comprometido en lugar de renovaciones anuales) y extender el flujo mensual a 2031 | Finanzas |
| 6 | Validar MOD (24 personas), alimentación, seguro de caución, flete y despacho | Producción / RR. HH. / Comex |
| 7 | Confirmar el tratamiento del IVA y de los preoperativos | Impuestos / Contabilidad |
