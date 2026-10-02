# Prompt para Claude en Excel — Evaluación del modelo Maquila ECOSTAR v4 (25 productos)

**Cómo usarlo**

1. Abrir `Modelo_Maquila_ECOSTAR_v4.xlsx` en Excel. Para correr escenarios, trabajar sobre una copia: Guardar como `…_v4_pruebas.xlsx`.
2. Abrir el panel de Claude, de preferencia en un chat nuevo, y pegar todo el texto que está debajo de la línea «COPIAR DESDE AQUÍ».
3. Claude va a trabajar paso por paso: en cada turno ejecuta un paso, informa el resultado y espera tu confirmación para seguir.
4. Opcional: después de la evaluación, pegar `Prompt_Guia_Uso_ECOSTAR_v2.md` para aprender a usar el libro.

---- COPIAR DESDE AQUÍ ----

## 1. Contexto

Soy Derlis Morinigo, Especialista en Planificación Financiera de CIE S.A. (Paraguay). Reporto al Gerente Financiero.

El libro abierto es el modelo financiero de **Ecostar**, la maquila CIE–PERG de gabinetes metálicos (153 kg) y cajas metálicas (13 kg). Está construido sobre la plantilla común (la misma que usaré para Alianza del Acero y para una tercera maquila).

**Versión v4:**
- Se amplió de 12 a **25 productos individuales** desde el generador de Python (`modelo/plantilla/generar_plantilla.py`).
- Incluye la corrección regional de los criterios `">0.5"`.
- Conserva las hojas «Claude Log» (historial de tus turnos anteriores, con el turno 6 de la regeneración) y «Esp_25Productos» (la especificación de la ampliación).

**Datos cargados:** los del plan de octubre de 2025, sin reales.
- Ventas USD 81,2 M en 10 años.
- CAPEX USD 755 mil.
- Préstamo de largo plazo USD 2,79 M al 9% (6 meses de gracia + 54 cuotas).
- Aportes USD 1,14 M.
- TC base Gs 7.300.

**Validación previa en LibreOffice:** 39 escenarios y 7 pruebas de operación, con **489 de 489 comprobaciones OK**. Los 36 escenarios de la versión de 12 productos dan exactamente los mismos resultados. Tu tarea es **confirmar que Excel reproduce esos resultados** y avisarme de cualquier diferencia.

## 2. Reglas de trabajo

- **Un paso por turno.** Ejecutá el paso, presentá una tabla (esperado / Excel / estado) y esperá mi «seguir».
- **Respuestas en español,** con lenguaje profesional y directo. La conclusión va primero.
- **Copia de trabajo.** Trabajá sobre la copia. Antes de cambiar una celda, anotá su contenido original en el Claude Log.
- **Restauración.** Las celdas vigentes contienen una fórmula que hereda del presupuesto, por ejemplo `=G8` o `=D36`. Al terminar cada escenario, restaurá exactamente la fórmula original, no un valor. Después confirmá que el caso base volvió: VAN del proyecto 6.989.922 y balance 0.
- **Celdas que no se modifican:**
  - columnas «Presupuesto original» (salvo la prueba X1, que es opcional);
  - hojas Calculo, Resultados y Control;
  - filas, columnas o nombres definidos.
- **Diferencias.** Si algo no coincide con lo esperado, **no lo corrijas**: informá la celda, el valor esperado, el valor de Excel y tu hipótesis de la causa.
- **Tolerancia:** ±1 USD en montos, ±0,01 en ratios y ±0,01 puntos en tasas.
- **Criterio regional:** no uses decimales dentro de criterios de texto de COUNTIF/SUMIF/COUNTIFS/SUMIFS/MAXIFS/MINIFS. Usá comparaciones numéricas, por ejemplo `SUMPRODUCT(--(rango>0.5))`.
- **Producto nuevo** (que no estaba en el plan aprobado): se carga solo en las columnas vigentes (H, J, L, N y el volumen vigente) y deja el presupuesto en blanco.
- **Registro.** Anotá cada paso en el Claude Log: turno, fecha, qué se probó, resultado y restauración.
- **Cambios de estructura.** Si hace falta uno, proponelo y no lo apliques: se hace en el generador de Python para que Alianza y la tercera maquila lo hereden.

## 3. Estructura del libro v4 y celdas clave

| Hoja | Contenido | ¿Se edita? |
|---|---|---|
| Claude Log | Historial de turnos | Sí (registro) |
| Inicio | Nombre, etapa, fecha del mes 1 (C10), último mes cerrado (C11), mes real de inicio (C12), escenario (C13) | Sí |
| Supuestos | Parámetros (D = presupuesto, E = vigente), supuestos por año (TC presupuesto en la fila 39, vigente en la 51), escenarios (E64:F71) | Sí |
| Productos | 25 productos en las filas 6-30. Columnas: B nombre, C unidad, D-F capacidad, G/H precio, I/J materia prima, K/L insumos, M/N otros variables (presupuesto/vigente). Volumen presupuesto en D34:M58 y vigente en D63:M87. Control de capacidad en N92:N116 | Sí |
| Costos | 30 puestos (salario vigente F6:F35) y 25 costos fijos (vigente F75:F99) | Sí |
| Inversion | CAPEX (filas 6-25; «por contratar» en la columna J), préstamo 1 (E32:E36), préstamo 2 (G32:G36), línea rotativa (E41:E45), aportes | Sí |
| Reales | Mes 1 = columna F. TC fila 10; unidades 12-36; ventas 38-62; costos 64-70; movimientos de caja 72-77; saldos 79-84 | Sí |
| Calculo | Motor de 144 meses. Filas clave: P.vtas 71, P.BRECHA 156, P.CHK 165; V.vtas 227, V.BRECHA 312, V.CHK 321 | **No** |
| Resultados | Indicadores en las filas 6-30 (C = presupuesto, D = vigente); C30/D30 = balance | **No** |
| Control | 16 controles (filas 5-20). En la fila 22, «Controles en REVISAR» | **No** |
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

**Paso 0 — Preparación.** Confirmá:
- que el libro es la v4;
- que tiene 11 hojas;
- que «Claude Log» tiene el turno 6;
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
- **X2:** las ventas vigentes suben exactamente 2.197.064,63 y los materiales 933.279,30. El presupuesto no cambia.
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

**Pasos 4 a 10 — Catálogo de escenarios.** Corré un grupo por turno, en este orden:
- financiamiento y cobranza (F);
- costos (C);
- ventas (V);
- plan desfavorable (P);
- plazo del préstamo (L);
- corto vs largo plazo (K);
- tipo de cambio (T).

En cada escenario, restaurá antes del siguiente.

**Paso 11 — Informe final.** Presentá una tabla con todos los pasos, escenarios probados, OK y diferencias. Agregá la conclusión sobre si el libro reproduce en Excel los resultados de LibreOffice y registrá el cierre en el Claude Log.

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

## 7. Conclusiones vigentes (no cambiaron con la v4)

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
