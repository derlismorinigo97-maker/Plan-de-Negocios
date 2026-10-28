# Prompt para Claude en Excel — Evaluación del modelo Maquila ECOSTAR v7 (cuadros con la estructura del plan de negocio)

**Cómo usarlo**

1. Abrir `Modelo_Maquila_ECOSTAR_v7.xlsx` en Excel. Para correr escenarios, trabajar sobre una copia: Guardar como `…_v7_pruebas.xlsx`.
2. Abrir el panel de Claude, de preferencia en un chat nuevo, y pegar todo el texto que está debajo de la línea «COPIAR DESDE AQUÍ».
3. Claude va a trabajar paso por paso: en cada turno ejecuta un paso, informa el resultado y espera tu confirmación para seguir. `Paso_a_Paso_Validacion_ECOSTAR_v7.md` indica qué revisar en cada turno.
4. Opcional: después de la evaluación, pegar `Prompt_Guia_Uso_ECOSTAR_v2.md` para aprender a usar el libro.

---- COPIAR DESDE AQUÍ ----

## 1. Contexto

Soy Derlis Morinigo, Especialista en Planificación Financiera de CIE S.A. (Paraguay). Reporto al Gerente Financiero.

El libro abierto es el modelo financiero de **Ecostar**, la maquila CIE–PERG de gabinetes metálicos (153 kg) y cajas metálicas (13 kg). Está construido sobre la plantilla común (la misma que usaré para Alianza del Acero y para una tercera maquila).

**Versión v7** (la v4 amplió a 25 productos; la v5 agregó la guía de carga; la v6, la proyección rodante de 24 meses; la v7, los cuadros con la estructura del plan):
- Generado desde Python (`modelo/plantilla/generar_plantilla.py`): 25 productos individuales y corrección regional de los criterios `">0.5"`.
- «Guia_Carga» (paso a paso con estado automático), «Detalle_24m» (proyección mes a mes de la ventana móvil) y «Proyeccion_24m» (próximos 24 meses).
- **Nuevo en la v7:**
  - **Cuadros con la estructura del plan de negocio:** la hoja «Cuadros» (índice, selector de versión en C5 y resumen modelo frente a plan) y una hoja por cuadro (C01 Inversión, C07 Producción, C08 Ingresos, C09 Materiales, C10a Costos, C10b Costo unitario, C11 Rubros, C13 Deuda e IVA, C15 Resultados, C16 Capital trabajo, C17 Equilibrio, C18 Fondos, C19 VAN-TIR, C20 Depreciación, RRHH). Cada una tiene tres bloques con las mismas filas: A = modelo, B = plan publicado (entrada, Ecostar precargado) y C = diferencia con su explicación.
  - Los contadores de Detalle_24m («celdas con dato en el mes») pasaron al final de Calculo (filas det.TC a det.FC). En Detalle_24m ya no hay fórmulas dentro de los bloques.
  - El motor, Resultados, Proyeccion_24m, la guía y Control no cambiaron: dan exactamente lo mismo que la v6.
- Conserva «Claude Log» (hasta el turno 21, incluida tu validación de la v6) y «Esp_25Productos».

**Datos cargados:** los del plan de octubre de 2025, sin reales.
- Ventas USD 81,2 M en 10 años.
- CAPEX USD 755 mil.
- Préstamo de largo plazo USD 2,79 M al 9% (6 meses de gracia + 54 cuotas).
- Aportes USD 1,14 M.
- TC base Gs 7.300.

**Validación previa en LibreOffice (28/10/2026, sobre este mismo libro v7):** 39 escenarios con **495 de 495 comprobaciones OK**, idénticos a la v6 (1.625.364 valores); 7 pruebas de operación iguales a la v6; proyección de 24 meses **60 de 60 OK**; cuadros con la estructura del plan: 7 casos y **147 de 147 OK**; motor, Resultados, Proyeccion_24m, guía y Control idénticos a la v6 celda por celda.

**Validación de la v6 en Excel (27/10/2026, Excel 16 en Windows, configuración es-PY):** 439 de 439 comprobaciones OK, sin diferencias. Como el motor de la v7 es idéntico, en esta validación lo nuevo son los pasos 3e y 3f (cuadros).

Tu tarea es **confirmar que Excel reproduce esos resultados** y avisarme de cualquier diferencia.

## 2. Reglas de trabajo

- **Un paso por turno.** Ejecutá el paso, presentá una tabla (esperado / Excel / estado) y esperá mi «seguir».
- **Respuestas en español,** con lenguaje profesional y directo. La conclusión va primero.
- **Copia de trabajo.** Trabajá sobre la copia. Antes de cambiar una celda, anotá su contenido original en el Claude Log.
- **Restauración.** Las celdas vigentes contienen una fórmula que hereda del presupuesto, por ejemplo `=G8` o `=D36`. Al terminar cada escenario, restaurá exactamente la fórmula original, no un valor. Después confirmá que el caso base volvió: VAN del proyecto 6.989.922 y balance 0.
- **Celdas que no se modifican:**
  - columnas «Presupuesto original» (salvo la prueba X1, que es opcional);
  - hojas Calculo, Resultados, Control y las de cuadros (C01 a RRHH); en los cuadros, el bloque B «Plan publicado» tampoco se toca en esta validación;
  - filas, columnas o nombres definidos.
- **Diferencias.** Si algo no coincide con lo esperado, **no lo corrijas**: informá la celda, el valor esperado, el valor de Excel y tu hipótesis de la causa.
- **Tolerancia:** ±1 USD en montos, ±0,01 en ratios y ±0,01 puntos en tasas.
- **Criterio regional:** no uses decimales dentro de criterios de texto de COUNTIF/SUMIF/COUNTIFS/SUMIFS/MAXIFS/MINIFS. Usá comparaciones numéricas, por ejemplo `SUMPRODUCT(--(rango>0.5))`.
- **Limpieza de Detalle_24m:** en la v7 los contadores están en Calculo, así que Detalle_24m no tiene fórmulas dentro de los bloques. Igual conviene borrar solo los rangos cargados.
- **Producto nuevo** (que no estaba en el plan aprobado): se carga solo en las columnas vigentes (H, J, L, N y el volumen vigente) y deja el presupuesto en blanco.
- **Registro.** Anotá cada paso en el Claude Log: turno, fecha, qué se probó, resultado y restauración.
- **Cambios de estructura.** Si hace falta uno, proponelo y no lo apliques: se hace en el generador de Python para que Alianza y la tercera maquila lo hereden.

## 3. Estructura del libro v7 y celdas clave

| Hoja | Contenido | ¿Se edita? |
|---|---|---|
| Claude Log | Historial de turnos | Sí (registro) |
| Inicio | Nombre, etapa, fecha del mes 1 (C10), último mes cerrado (C11), mes real de inicio (C12), escenario (C13), meses de la ventana de detalle (C14 = 24), versión del plan (C15), fecha de la última revisión (C16), frecuencia Anual/Bienal (C17) | Sí |
| Guia_Carga | 37 pasos con estado automático (sección F = proyección de 24 meses y revisión del plan); C4 = mes en carga, C5 = pasos con alerta | Columnas G-J (Hecho, Fecha, Responsable, Notas) |
| Detalle_24m | Misma grilla que Reales (mes 1 = columna F; fila 4 = n° de mes; fila 7 = estado Real/Ventana/Anual). Bloques: TC (fila 10), volúmenes Q1-Q25, precios P1-P25, materia prima MP1-MP25, insumos IN1-IN25, otros variables OV1-OV25, dotación DO1-DO30, salarios SA1-SA30, costos fijos FC1-FC25 (código en la columna A). Celda vacía = supuesto anual; solo se usa dentro de la ventana | Sí (meses en amarillo) |
| Supuestos | Parámetros (D = presupuesto, E = vigente), supuestos por año (TC presupuesto en la fila 39, vigente en la 51), escenarios (E64:F71) | Sí |
| Productos | 25 productos en las filas 6-30. Columnas: B nombre, C unidad, D-F capacidad, G/H precio, I/J materia prima, K/L insumos, M/N otros variables (presupuesto/vigente). Volumen presupuesto en D34:M58 y vigente en D63:M87. Control de capacidad en N92:N116 | Sí |
| Costos | 30 puestos (salario vigente F6:F35) y 25 costos fijos (vigente F75:F99) | Sí |
| Inversion | CAPEX (filas 6-25; «por contratar» en la columna J), préstamo 1 (E32:E36), préstamo 2 (G32:G36), línea rotativa (E41:E45), aportes | Sí |
| Reales | Mes 1 = columna F. TC fila 10; unidades 12-36; ventas 38-62; costos 64-70; movimientos de caja 72-77; saldos 79-84 | Sí |
| Calculo | Motor de 144 meses. Filas clave: P.vtas 71, P.BRECHA 156, P.CHK 165; V.u1 177, V.vtas 227, V.BRECHA 312, V.CHK 321. Contadores de Detalle_24m: filas 324-332 (det.TC a det.FC) | **No** |
| Proyeccion_24m | Próximos 24 meses (columnas F a AC = meses Corte+1 a Corte+24): vigente, presupuesto y desvíos; C/D/E = año 1, año 2 y total de la ventana | **No** |
| Resultados | Indicadores en las filas 6-30 (C = presupuesto, D = vigente); C30/D30 = balance | **No** |
| Control | 16 controles (filas 5-20). En la fila 22, «Controles en REVISAR» | **No** |
| Cuadros | Índice de los cuadros. C5 = versión que muestran (Vigente / Presupuesto original). Resumen modelo frente a plan en las filas 29-43 (C = modelo, D = plan, E = diferencia, G = explicación) | Solo C5 |
| C01 a RRHH | Un cuadro del plan por hoja. Bloque A (modelo), B (plan publicado) y C (diferencia), con las mismas filas y el código de línea en la columna A. Columnas: D = año 0 (implantación), E a N = años de operación 1 a 10, O = total, P = explicación (bloque C) | **No** |
| Esp_25Productos | Especificación de la ampliación (sección 5: lista de verificación) | No |

Para un producto i (de 1 a 25):
- fila de datos = 5 + i;
- fila de volumen del presupuesto = 33 + i;
- fila de volumen vigente = 62 + i.

## 4. Lógica del motor

- **Tres capas:** presupuesto original (bloque P, que no cambia), reales hasta el último mes cerrado y proyección vigente.
- **Escenarios:** solo afectan los meses posteriores al corte.
- **Horizonte:** fijo, 10 años desde el inicio presupuestado.
- **Capital de trabajo:** por días (cobro 120, inventario 120, producto terminado 10, proveedores 0/30).
- **Tipo de cambio:** ventas, materiales y CAPEX en USD; nómina y parte de los fijos en guaraníes, convertidos con el TC del año.
- **Línea rotativa:**
  - se usa sola cuando la caja cae bajo el mínimo y se devuelve con el excedente;
  - mientras tiene saldo no se pagan dividendos;
  - al vencer se paga entera.
- **Caja negativa:** se informa como «Aporte adicional necesario». En el VAN de los socios se trata como un aporte implícito a costo Ke.
- **VAN del proyecto (FCFF a WACC 9,49%):** no cambia con el financiamiento. El VAN de los socios se calcula a Ke 16,32%.
- **DSCR:** el mínimo se mide desde el segundo año de operación.

## 5. Plan de evaluación paso a paso

**Ruta corta.** Si al pegar este prompt te pido «Ruta corta», ejecutá solo los pasos 0, 1, 3e, 3f y 11: el motor es idéntico al de la v6, que ya validaste en Excel.

**Paso 0 — Preparación.** Confirmá:
- que el libro es la v7 (30 hojas: Claude Log, Inicio, Guia_Carga, Supuestos, Productos, Costos, Inversion, Detalle_24m, Reales, Calculo, Proyeccion_24m, Resultados, Control, Cuadros, C01 Inversión, C07 Producción, C08 Ingresos, C09 Materiales, C10a Costos, C10b Costo unitario, C11 Rubros, C13 Deuda e IVA, C15 Resultados, C16 Capital trabajo, C17 Equilibrio, C18 Fondos, C19 VAN-TIR, C20 Depreciación, RRHH, Esp_25Productos) y que «Claude Log» llega al turno 21;
- que el cálculo está en automático.

Recalculá con Ctrl+Alt+F9.

**Paso 1 — Integridad del libro base.** Compará Resultados!D6:D30 con estos valores:

| Indicador | Esperado |
|---|---|
| Ventas totales / EBITDA total / EBITDA año 1 | 81.222.448 / 16.309.504 / 509.970 |
| Resultado neto / CAPEX | 14.751.246 / 755.044 |
| Necesidad máxima de fondos (mes) | 3.207.787 (mes 13) |
| Aporte adicional / deuda máxima / intereses | 311.740 / 2.791.587 / 813.214 |
| Cancelación de la deuda / brecha máxima / meses con brecha | 5,00 años / 373.701 / 12 |
| DSCR mínimo desde el año 2 | 0,40 |
| VAN / TIR del proyecto | 6.989.922 / 34,93% |
| VAN / TIR de los socios | 3.695.678 / 43,63% |
| Recupero / balance de control | 4,13 años / 0 |

Revisá además:
- Control: en REVISAR deben estar solo C13 y C14.
- No debe haber ningún criterio de texto con decimales en todo el libro.
- Los nombres `Prod_Nombre`, `Vol_O`, `Vol_V` y `c_CapExceso` deben llegar a 25 productos (B6:B30, D34:M58, D63:M87, N92:N116).

**Paso 2 — Estructura de 25 productos.** Corré X2, X3 y V4 del catálogo. X1 es opcional porque es muy laboriosa a mano y ya se validó en LibreOffice:
- **X2:** las ventas vigentes suben exactamente 2.197.064,63 y los materiales (materia prima + insumos) 933.279,30: materia prima +753.279,30 y insumos +180.000. Los otros variables por unidad suben 60.000 aparte. El presupuesto no cambia.
- **X3:** C11 en REVISAR.
- **V4:** C10 en REVISAR con 16.

**Paso 3 — Meses reales congelados.** Cargá estos datos en Reales, meses 1 a 6 (columnas F a K):

| Fila | Concepto | Valor por mes |
|---|---|---|
| 10 | TC | 7.400 |
| 12 / 13 | Unidades de gabinetes / cajas | 600 / 4.000 |
| 38 / 39 | Ventas de gabinetes / cajas | 162.000 / 92.000 |
| 64 | Materiales | 120.000 |
| 65 | Otros variables | 30.000 |
| 66 | Personal de producción (PYG) | 300.000.000 |
| 67 | Personal administrativo (PYG) | 60.000.000 |
| 68 | Fijos de producción | 15.000 |
| 69 | Administración | 25.000 |
| 70 | Tributo | 2.540 |
| 74 | Intereses | 23.026 |
| 77 | Dividendos | 0 |
| 75 | Amortización | 0 |

Movimientos que van en un solo mes:
- CAPEX (fila 72): mes 1 = 600.000 y mes 4 = 30.000; el resto, 0.
- Desembolso (fila 73): mes 1 = 2.791.587.
- Aporte (fila 76): mes 1 = 767.782.

Saldos del mes 6 (columna K):
- caja 1.500.000;
- clientes 900.000;
- inventarios 500.000;
- IVA 0;
- proveedores 100.000;
- deuda 2.791.587.

En Inversion, cargá ejecutado y pagado (columnas G y H):
- ítems 1-11 (filas 6-16) = presupuesto de la columna D;
- ítem 12 (fila 17) = 35.833,91.

Por último, Inicio!C11 = 6.

Valores esperados con escenario Vigente:
- ventas de los meses 1 a 6 = 254.000 y del mes 7 = 404.167;
- caja del mes 6 = 1.500.000 y ajuste de conciliación del mes 6 = −68.081;
- ventas del año 1 = 3.949.000 y EBITDA del año 1 = 331.853;
- aporte adicional = 557.911 y meses con brecha = 30;
- VAN del proyecto = 6.858.766 y VAN de los socios = 3.546.903;
- balance = 0.

Valores esperados al cambiar Inicio!C13 a «Conservador»:
- los meses 1 a 6 **no cambian** (ventas, EBITDA, caja);
- ventas del mes 7 = 345.562;
- VAN del proyecto = 31.573 y VAN de los socios = −1.330.308;
- aporte adicional = 3.286.964;
- CAPEX = 767.549;
- balance = 0.

C14 sigue en REVISAR mientras Inicio!C12 (mes real de inicio) esté vacío. Restaurá todo al terminar.

**Referencias para los pasos 3b a 3d.**
- **Libro de partida:** el libro base (sin reales), con Inicio!C14 = 24 y escenario Vigente.
- **Detalle_24m:** el mes m está en la columna 5 + m (mes 1 = F, mes 12 = Q, mes 24 = AC, mes 36 = AO). Filas: TC 10, Q1 14, Q2 15, P1 42, MP1 70, DO1 154, SA1 187, FC1 220.
- **Calculo:** filas V.tc 174, V.u1 177, V.u2 178, V.s1 202, V.mp_f 229, V.dir_n 233, V.dir_f 234, V.fp_f 242, V.PERA 254 (mes 1 = columna F, mes 25 = AD).
- **Proyeccion_24m:**
  - ventas en la fila 11 y EBITDA en la 19;
  - C = año 1 de la ventana, D = año 2, E = total y F = primer mes;
  - fila 7 = origen de volúmenes, fila 8 = origen de costos.
- **Valores de partida:**
  - ventas de la ventana: 4.850.000 / 6.181.200 / 11.031.200;
  - EBITDA del año 1: 509.970;
  - Guia_Carga paso 29 = «○ Usa supuestos anuales».
- **Restauración y controles:** cada caso empieza desde el libro base, así que hay que restaurar entre casos. En todos, el balance (Resultados!D30) = 0 y el presupuesto (Resultados columna C) no cambia.

**Paso 3b — Volumen y guía.**
- **Volumen:** cargá Q1 = 1.100 y Q2 = 6.500 en los meses 1-24 (F14:AC15). Esperado:
  - V.u1 = 1.100 en los meses 1-24 y 1.500 en el mes 25, que es el supuesto anual;
  - V.u2 = 6.500 en los meses 1-24;
  - Proyeccion_24m:
    - ventas del mes 1 = 446.500;
    - ventas año 1 / año 2 / total = 5.358.000 / 5.411.580 / 10.769.580;
    - EBITDA del año 1 = 679.293;
    - la fila 7 dice «Detalle» en los 24 meses;
  - Guia_Carga paso 29 = «✔ 24 meses detallados».
- **Mes faltante:** borrá Q1 **y** Q2 del mes 24 (AC14:AC15). El paso 29 cuenta los meses que tienen al menos un volumen cargado. Esperado: «⚠ 23 de 24 meses: completar hasta el mes 24». Restaurá.

**Paso 3c — Precio, costos y TC.** Corré un caso por vez y restaurá entre casos:

| Caso | Carga en Detalle_24m | Esperado |
|---|---|---|
| Precio | P1 = 280 en los meses 1-12 (F42:Q42) | V.s1 = 280.000 en los meses 1-12 y 340.875 en el mes 13 (1.250 × 272,70). Ventas del año 1 de la ventana = 4.970.000. |
| Materia prima | MP1 = 90 en los meses 1-24 | V.mp_f sube 7.350 en el mes 1 (1.000 × (90 − 82,65)) y 8.154,38 en el mes 13. El mes 25 no cambia. EBITDA del año 1 = 413.656. |
| Personal | DO1 = 6 y SA1 = 5.000.000 (Gs) en los meses 1-24 | V.dir_n del mes 1 = 45 (base 43). V.dir_f del mes 1 sube 2.377,55. EBITDA del año 1 = 479.713. |
| Costo fijo | FC1 = 20.000.000 (Gs) en los meses 1-24 (base Gs 14.000.000) | V.fp_f del mes 1 sube 821,92 (6.000.000 / 7.300). EBITDA del año 1 = 499.614. |
| TC | TC = 8.000 en los meses 1-24 (F10:AC10) | V.tc = 8.000 en los meses 1-24 y 7.300 en el mes 25. V.PERA del mes 1 = 8.175,99 (= 8.959,99 × 7.300 / 8.000). Ventas sin cambio. EBITDA del año 1 = 584.427. |

**Paso 3d — Escenario, ventana y ventana móvil.**
- **Escenario sobre el detalle:** Q1 = 1.100 en los meses 1-24 e Inicio!C13 = «Conservador».
  - Esperado: V.u1 = 0 en los meses 1-3 (atraso de 3 meses), 990 en los meses 4-24 (1.100 × 0,90) y 1.125 en el mes 25.
  - Restaurá C13 = «Vigente».
- **Ventana de 12 meses:** Q1 = 1.100 en los meses 1-24 e Inicio!C14 = 12.
  - Esperado: V.u1 = 1.100 en los meses 1-12 y 1.250 (anual) en los meses 13-24.
  - Restaurá C14 = 24.
- **Ventana móvil:** con los reales del paso 3 cargados (Inicio!C11 = 6) y Q1 = 1.100 en los meses 1-36 (F14:AO14). Esperado:
  - V.u1 = 600 en los meses 1-6 (reales), 1.100 en los meses 7-30 (ventana) y 1.500 en los meses 31-36 (el detalle fuera de la ventana no se usa);
  - Proyeccion_24m empieza en el mes 7, con ventas del primer mes = 431.166,67;
  - ventas año 1 / año 2 / total = 5.432.170 / 5.953.414,70 / 11.385.584,70;
  - balance = 0.
- **Restauración final:** Detalle_24m vacío, Inicio!C11 vacío, Reales e Inversion G:H como estaban. Confirmá que el VAN del proyecto vuelve a 6.989.922.

**Paso 3e — Cuadros con la estructura del plan: lectura y controles.** Sobre el libro base (sin reales), con Cuadros!C5 = «Vigente». Solo lectura, salvo el selector del punto 6.

1. **Resumen de la hoja Cuadros (filas 29-41; C = modelo, D = plan, E = diferencia):**

| Concepto | Modelo | Plan | Diferencia |
|---|---|---|---|
| Inversión fija inicial (fila 29) | 655.044 | 655.044 | 0 |
| Capital operativo inicial (fila 30) | 2.388.064 | 2.336.543 | 51.521 |
| Ventas de los 10 años (fila 31) | 81.222.448 | 81.222.448 | 0 |
| Costos totales de los 10 años (fila 32) | 66.471.202 | 65.277.283 | 1.193.919 |
| EBITDA de los 10 años (fila 33) | 16.309.504 | 17.354.010 | −1.044.506 |
| Utilidad neta de los 10 años (fila 34) | 14.751.246 | 15.945.168 | −1.193.922 |
| Caja al cierre del año 10 (fila 36) | 1.637.884 | 1.909.351 | −271.467 |
| Punto de equilibrio del año 1, unidades (fila 37) | 73.373 | 66.042 | 7.331 |
| VAN — método del plan (fila 39) | 4.177.576 | 4.588.736 | −411.160 |
| TIR — método del plan (fila 40) | 23,99% | 22,61% | 1,38 p.p. |
| Payback — método del plan (fila 41) | 8 | 7 | 1 |

2. **C18 Fondos:** la fila 27 («Control: saldo − caja del motor») = 0 en los años 0 a 10 (D27:N27). La fila 26 (caja al cierre) del año 10 (N26) = 1.637.884.
3. **C15 Resultados:** totales de la columna O iguales a Resultados!D: ventas (O8) 81.222.448, EBITDA (O11) 16.309.504, utilidad (O16) 14.751.246. EBITDA del año 1 (E11) = 509.970.
4. **C19 VAN-TIR:**
   - FCFF descontado, total (O15) = 6.989.922 = VAN del proyecto de Resultados; flujo de los socios descontado, total (O19) = 3.695.678 = VAN de los socios;
   - bloque B (plan publicado): flujo actualizado total (O61) = 4.588.737 (±1), y la TIR recalculada con los flujos publicados (E104) = 22,61%. Prueba que el método del plan está bien replicado.
5. **Filas de ajuste en 0** (sin reales ni detalle mensual): C09 Materiales fila 85; C10a Costos filas 23, 31 y 65; C11 Rubros fila 16; C17 Equilibrio fila 22; C20 Depreciación fila 34 (años 0 a 10).
6. **Selector:** Cuadros!C5 = «Presupuesto original». Los valores del punto 1 no cambian, porque no hay reales y el vigente hereda el presupuesto. Volvé a «Vigente».

**Paso 3f — Cuadros con escenario y con detalle mensual.** Un caso por vez; restaurá entre casos.

| Caso | Cambio | Esperado |
|---|---|---|
| Conservador | Inicio!C13 = «Conservador» | La operación empieza en el mes 4. C08 Ingresos: año 0 (D8) = 0 y año 1 (E8) = 4.146.750. C10a Costos: total del año 0 (D36) = 283.020 (3 meses de implantación). C18 Fondos: fila 27 = 0 en todos los años. |
| Detalle de materia prima | Detalle_24m MP1 = 90 en los meses 1-24 (F70:AC70) | C09 Materiales, fila 85 (ajuste de materia prima): año 1 (E85) = 88.200 (12.000 × 7,35), año 2 (F85) = 97.852,50 y año 3 (G85) = 0. C10a Costos: total del año 1 (E36) = 4.779.039. |

Restaurá todo al terminar: Inicio!C13 = «Vigente», Detalle_24m vacío. Confirmá que el VAN del proyecto vuelve a 6.989.922.

**Pasos 4 a 10 — Catálogo de escenarios.** Corré un grupo por turno, en este orden:
- financiamiento y cobranza (F);
- costos (C);
- ventas (V);
- plan desfavorable (P);
- plazo del préstamo (L);
- corto vs largo plazo (K);
- tipo de cambio (T).

En cada escenario, restaurá antes del siguiente.

**Paso 11 — Informe final.** Presentá una tabla con todos los pasos (incluidos 3b a 3f), escenarios probados, OK y diferencias. Agregá la conclusión sobre si el libro reproduce en Excel los resultados de LibreOffice y registrá el cierre en el Claude Log.

## 6. Catálogo de escenarios con valores esperados

«Línea X» significa: Inversion!E41 = X (límite), E42 = tasa, E43 = 1 y E44 vacío (renovada), salvo que se indique otra cosa. Las columnas de resultado son de Resultados, columna D.

| Id | Cambios | Ventas | VAN proyecto | VAN socios | Aporte adicional | Meses bajo caja mín. | Intereses | Cancela (años) | DSCR mín. |
|---|---|---|---|---|---|---|---|---|---|
| E00 | Base (sin cambios) | 81.222.448 | 6.989.922 | 3.695.678 | 311.740 | 12 | 813.214 | 5,0 | 0,40 |
| F1 | Inversion!E32 = 3.800.000 | 81.222.448 | 6.989.922 | 3.593.001 | 0 | 1 | 1.106.974 | 5,0 | 0,29 |
| F2 | Inversion!E32 = 1.800.000 | 81.222.448 | 6.989.922 | 3.680.996 | 816.529 | 41 | 524.356 | 5,0 | 0,62 |
| F3 | Supuestos!E21 = 180 | 81.222.448 | 6.230.650 | 2.932.499 | 1.632.135 | 57 | 813.214 | 5,0 | 0,12 |
| F4 | Supuestos!E21 = 0 | 81.222.448 | 8.508.466 | 3.895.390 | 0 | 0 | 813.214 | 5,0 | 0,96 |
| F5 | F1 + F3 | 81.222.448 | 6.230.650 | 2.954.155 | 1.397.534 | 48 | 1.106.974 | 5,0 | 0,09 |
| F6 | F2 + F4 | 81.222.448 | 8.508.466 | 3.975.861 | 0 | 0 | 524.356 | 5,0 | 1,49 |
| F7 | F1 + F4 | 81.222.448 | 8.508.466 | 3.813.553 | 0 | 0 | 1.106.974 | 5,0 | 0,71 |
| F8 | F2 + F3 | 81.222.448 | 6.230.650 | 2.810.935 | 2.055.741 | 55 | 524.356 | 5,0 | 0,18 |
| C1 | Productos!J6:J7, L6:L7, N6:N7 = I, K, M × 1,10 | 81.222.448 | 3.619.713 | 1.330.719 | 1.601.514 | 55 | 813.214 | 5,0 | −0,15 |
| C2 | Ídem × 1,20 | 81.222.448 | 249.504 | −1.129.352 | 3.186.573 | 104 | 813.214 | 5,0 | −0,69 |
| C3 | Ídem × 1,30 | 81.222.448 | −3.120.704 | −3.725.162 | 5.685.501 | 117 | 813.214 | 5,0 | −1,24 |
| C4 | C3 + Costos!F6:F31 = E × 1,30 + Costos!F75:F83 = E × 2 | 81.222.448 | −7.020.642 | −6.753.823 | 10.601.180 | 118 | 813.214 | 5,0 | −1,95 |
| V1 | Productos!D63:M64 = D34:M35 × 0,90 | 73.100.203 | 5.422.991 | 2.690.586 | 430.519 | 21 | 813.214 | 5,0 | 0,22 |
| V2 | Ídem × 0,80 | 64.977.958 | 3.856.060 | 1.656.795 | 656.645 | 33 | 813.214 | 5,0 | 0,03 |
| V3 | Ídem × 0,70 | 56.855.713 | 2.289.130 | 626.407 | 882.772 | 45 | 813.214 | 5,0 | −0,15 |
| V4 | Ídem × 1,10 (C10 = 16) | 89.344.692 | 8.556.853 | 4.713.191 | 290.460 | 8 | 813.214 | 5,0 | 0,58 |
| V5 | Productos!D63:F63 = 15.000 / 15.000 / 21.000; D64:F64 = 90.000 / 90.000 / 120.000 | 83.553.352 | 7.601.969 | 4.106.060 | 378.600 | 8 | 813.214 | 5,0 | 0,76 |
| P1 | D63:M64 = D34:M35 × 0,70; Productos!H6:H7 = G × 0,95; Supuestos!E6 = 7; Inversion!J6:J19 = D × 1,15 | 50.947.355 | −105.270 | −1.158.116 | 2.635.012 | 94 | 813.214 | 5,0 | −0,58 |
| P2 | Inicio!C13 = «Conservador» | 67.474.468 | −486.490 | −1.723.354 | 3.802.828 | 117 | 813.214 | 5,0 | −0,81 |
| P3 | Inicio!C13 = «Optimista» | 86.989.242 | 10.466.057 | 5.389.993 | 0 | 0 | 813.214 | 5,0 | 1,04 |
| L1 | Inversion!E36 = 30 | 81.222.448 | 6.989.922 | 3.636.649 | 1.361.878 | 34 | 508.039 | 3,0 | 0,25 |
| L2 | Inversion!E36 = 78 | 81.222.448 | 6.989.922 | 3.696.158 | 16.353 | 2 | 1.134.635 | 7,0 | 0,53 |
| L3 | Inversion!E35 = 12 y E36 = 84 | 81.222.448 | 6.989.922 | 3.645.606 | 0 | 0 | 1.355.681 | 8,0 | 0,55 |
| L4 | E32 = 0 + línea 4.000.000 al 9% | 81.222.448 | 6.989.922 | 3.760.141 | 0 | 0 | 642.648 | 4,0 | 1,56 |
| K1 | Inversion!E32 = 3.250.000 | 81.222.448 | 6.989.922 | 3.654.336 | 97.745 | 7 | 946.754 | 5,0 | 0,34 |
| K2 | Préstamo 2 (G32:G36): 500.000, mes 25, 10%, gracia 0, 12 cuotas | 81.222.448 | 6.989.922 | 3.691.088 | 234.637 | 4 | 843.459 | 5,0 | 0,40 |
| K3 | Línea 1.000.000 al 10% | 81.222.448 | 6.989.922 | 3.698.760 | 0 | 0 | 833.792 | 5,0 | 0,40 |
| K4 | E32 = 0 + línea 4.000.000 al 10% | 81.222.448 | 6.989.922 | 3.711.577 | 0 | 0 | 734.934 | 4,1 | 1,38 |
| K5 | K4 con E44 = 24 (vence en el mes 24) | 81.222.448 | 6.989.922 | 3.645.096 | 2.191.679 | 22 | 458.820 | 2,1 | 1,38 |
| K6 | E32 = 760.000 + línea 3.000.000 al 10% | 81.222.448 | 6.989.922 | 3.739.865 | 0 | 0 | 726.417 | 5,0 | 0,84 |
| K7 | K4 con E42 = 7,99% | 81.222.448 | 6.989.922 | 3.806.552 | 0 | 0 | 554.410 | 3,9 | 1,79 |
| K8 | K4 con E42 = 11,24% | 81.222.448 | 6.989.922 | 3.646.770 | 0 | 0 | 857.501 | 4,2 | 1,20 |
| T1 | Supuestos!D51:M51 = 8.760 | 81.222.448 | 7.989.367 | 4.356.428 | 13.960 | 2 | 813.214 | 5,0 | 0,58 |
| T2 | Ídem = 6.145 | 81.222.448 | 5.862.801 | 2.934.837 | 707.018 | 35 | 813.214 | 5,0 | 0,19 |
| T3 | Ídem = 5.840 | 81.222.448 | 5.490.753 | 2.660.071 | 872.928 | 39 | 813.214 | 5,0 | 0,13 |
| X1 | Opcional: partición en los 25 espacios (ver nota abajo) | 81.222.448 | 6.989.922 | 3.695.678 | 311.740 | 12 | 813.214 | 5,0 | 0,40 |
| X2 | Solo vigente. Nombres en B8, B18, B30 y «u» en la columna C. H/J/L/N: fila 8 = 100/30/10/5; fila 18 = 50/20/0/0; fila 30 = 200/80/20/0. Volumen vigente D65:M65 = 1.200; D75:M75 = 600; D87:M87 = 300 | 83.419.512 | 7.641.309 | 4.132.976 | 197.575 | 6 | 813.214 | 5,0 | 0,53 |
| X3 | Solo vigente. B30 = «Sin precio», C30 = «u», J30 = 10, D87:M87 = 500 (H30 sin cambiar). C11 → REVISAR | 81.222.448 | 6.951.415 | 3.672.044 | 324.970 | 12 | 813.214 | 5,0 | 0,39 |

**X1 (opcional): partición.** Cargá en las columnas de presupuesto de Productos los gabinetes repartidos en los productos 1-13 y las cajas en los 14-25:
- precio, costos y año de ampliación iguales al producto original;
- capacidad y volumen divididos por 13 (gabinetes) o por 12 (cajas).

Todos los indicadores deben quedar iguales a la base. Hacelo solo sobre una copia descartable, porque modifica el presupuesto.

## 7. Conclusiones vigentes (no cambiaron con la v4, v5, v6 ni v7)

- **Variables más sensibles:**
  - costo unitario: +20,7% en costos variables anula el VAN;
  - días de cobro: entre USD 15 y 17 mil de caja por día;
  - tipo de cambio: USD 71 a 75 mil de VAN por cada 1% de apreciación del guaraní.
- **Repago:** la capacidad es de unos 4 años si todo el excedente va a deuda. El cronograma 6 + 54 choca con los años 2 y 3 (DSCR 0,40 y 0,96).
- **Financiamiento:**
  - largo plazo para el CAPEX y el capital de trabajo permanente;
  - una línea rotativa comprometida de USD 0,5 a 1,0 M;
  - no financiar el capital de trabajo permanente con corto plazo, por el riesgo de no renovación.

## 8. Supuestos y limitaciones

- Supuestos del plan de 2025, sin reales. WACC 9,49% y Ke 16,32% históricos.
- Tasa del largo plazo 9% (plan) y del corto plazo 10% ilustrativa.
- Referencias del BCP: tasa activa promedio en ME 7,99% (marzo de 2026) y tope en ME 11,24% (julio de 2026).
- Los tamaños de los shocks y los datos de los productos de prueba (X2, X3) son ilustrativos.
- La caja positiva no genera intereses.
- El volumen sobre la capacidad solo genera una alerta; no se limita.

Empezá por el paso 0.
