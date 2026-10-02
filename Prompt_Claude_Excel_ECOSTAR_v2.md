# Prompt para Claude en Excel — Modelo Maquila ECOSTAR v2

**Cómo usarlo**

1. Abrir `Modelo_Maquila_ECOSTAR_v2.xlsx` en Excel. De preferencia, trabajar sobre una copia: Guardar como `…_pruebas.xlsx`.
2. Abrir el panel de Claude y pegar todo el texto que está debajo de la línea «COPIAR DESDE AQUÍ».
3. Empezar por la tarea 1, que valida que Excel da los mismos resultados que las pruebas hechas en LibreOffice.

---- COPIAR DESDE AQUÍ ----

## 1. Contexto

Soy Derlis Morinigo, Especialista en Planificación Financiera de CIE S.A. (Paraguay). Reporto al Gerente Financiero.

El libro abierto es el modelo financiero de **Ecostar**, la maquila CIE–PERG de gabinetes metálicos (153 kg) y cajas metálicas (13 kg). Está construido sobre una **plantilla común v2**, que después usaré también para Alianza del Acero y para una tercera maquila en implantación.

- Los datos cargados son los del plan de octubre de 2025: ventas USD 81,2 M en 10 años, CAPEX USD 755 mil, préstamo de largo plazo USD 2,79 M al 9% (6 meses de gracia + 54 cuotas) y aportes USD 1,14 M.
- **Todavía no hay datos reales cargados.**
- La moneda es USD, con TC base de Gs 7.300.
- Ya se corrió una batería de 36 escenarios en LibreOffice. Sus resultados de referencia están en la sección 6.

## 2. Reglas de trabajo

- Respondeme en español, con lenguaje profesional y directo. Empezá por la conclusión y después mostrá supuestos, cálculo, interpretación, riesgos y recomendación.
- **No inventes datos.** Si falta información, decilo y marcá el supuesto como provisional.
- **Solo modificá celdas de entrada:**
  - celdas amarillas con texto azul;
  - columnas «Vigente», que en gris heredan el presupuesto.
- **Nunca modifiques:**
  - las columnas «Presupuesto original»;
  - las hojas Calculo, Resultados y Control;
  - filas, columnas o nombres definidos.
- **Restauración.** Antes de cambiar una celda, anotá su contenido original. Las celdas vigentes contienen una fórmula que hereda del presupuesto, por ejemplo `=D39`. Al terminar una prueba, restaurá esa fórmula exacta; no un valor fijo.
- **Después de cada cambio:**
  - revisá la hoja Control y avisame de los controles en «REVISAR»;
  - confirmá que el «Balance de control» (Resultados!C30 y D30) sea 0.
- **Cambios de estructura.** Si hace falta un cambio estructural (filas, fórmulas del motor, nuevas hojas), proponelo y esperá mi confirmación. El libro se genera con un script de Python (`modelo/plantilla/generar_plantilla.py`), y esos cambios deben replicarse ahí para que Alianza y la tercera maquila los hereden.
- **Fórmulas que me des:** completas, para Excel en español y con una validación para comprobar el resultado.
- **Conceptos que no deben mezclarse:** rentabilidad (VAN/TIR), liquidez (caja), capital de trabajo y necesidad de financiamiento.

## 3. Estructura del libro (9 hojas)

| Hoja | Contenido | ¿Se edita? |
|---|---|---|
| Inicio | Nombre, etapa, fecha del mes 1, último mes cerrado (`n_Corte`, C11), mes real de inicio (`n_InicioReal`, C12), escenario (`n_Escenario`, C13: Vigente/Conservador/Optimista) | Sí |
| Supuestos | A: parámetros (columna D = presupuesto, E = vigente). B: supuestos por año (TC, variaciones, dividendos). C: escenarios Conservador/Optimista (E64:F71). D: plan publicado. E: valores aplicados (no editar). | Sí (D, E y C) |
| Productos | Precio y costos por unidad (presupuesto/vigente), capacidad y volúmenes por año | Sí |
| Costos | 30 puestos (salario y dotación) y 25 costos fijos (moneda PYG/USD, crecimiento, desde qué año) | Sí |
| Inversion | 20 ítems de CAPEX, 3 préstamos, línea rotativa (B2) y aportes | Sí |
| Reales | Carga mensual de la ejecución y saldos al cierre (mes 1 = columna F) | Sí (mensual) |
| Calculo | Motor de 144 meses: código en la columna A, total en la E, mes 1 en la F. Bloque P = presupuesto; bloque V = vigente. | **No** |
| Resultados | Indicadores (filas 6-30: C = presupuesto, D = vigente), resúmenes anuales, desvíos y gráficos | **No** |
| Control | 16 validaciones con su acción sugerida | **No** |

### Celdas clave para escenarios (nombres definidos)

| Nombre | Celda | Qué es |
|---|---|---|
| `vg_Inicio` | Supuestos!E6 | Mes de inicio de operación (vigente) |
| `vg_DSO` / `vg_DInv` | Supuestos!E21 / E22 | Días de cobro / días de inventario (vigente) |
| `vg_CajaMin` | Supuestos!E26 | Caja mínima en días de ventas |
| `vg_WACC` / `vg_Ke` | Supuestos!E30 / E31 | Tasas de descuento (vigente) |
| `TC_V` | Supuestos!D51:M51 | TC vigente, años 1 a 10 (presupuesto en la fila 39) |
| `Precio_V`, `MP_V`, `INS_V`, `OV_V` | Productos!H, J, L, N filas 6-7 | Precio y costos por unidad vigentes (presupuesto en G, I, K, M) |
| `Vol_V` | Productos!D37:M38 | Unidades vigentes (presupuesto en D21:M22) |
| `Sal_V` / `FC_V` | Costos!F6:F31 / F75:F83 | Salarios / costos fijos vigentes (presupuesto en la columna E) |
| CAPEX «Por contratar (si difiere)» | Inversion!J6:J19 | Costo final pendiente por ítem |
| Préstamo 1 vigente | Inversion!E32:E36 | Monto, mes, tasa, gracia, cuotas |
| Préstamo 2 vigente | Inversion!G32:G36 | Ídem |
| Línea rotativa vigente | Inversion!E41:E45 | Límite, tasa, desde, vencimiento, saldo al corte |

## 4. Lógica del motor (lo que debés respetar al interpretar)

- **Tres capas:**
  - presupuesto original (bloque P), que no cambia nunca;
  - reales, hasta el último mes cerrado;
  - proyección vigente, desde el mes siguiente.
- **Escenarios.** Solo afectan la proyección, nunca los meses reales.
- **Horizonte.** Es fijo: inicio presupuestado + 10 años (mes 120). Un atraso reduce los meses operativos; no extiende el horizonte.
- **Capital de trabajo.** Se calcula por días sobre el mes corriente: clientes 120, inventario de materiales 120, producto terminado 10, proveedores 0/30.
- **Tipo de cambio.** Ventas, materiales y CAPEX están en USD; la nómina y parte de los costos fijos en guaraníes, convertidos con el TC del año.
- **Línea rotativa:**
  - se usa sola cuando la caja cae bajo el mínimo y se devuelve con el excedente;
  - mientras tiene saldo no se pagan dividendos;
  - al vencer se paga entera.
- **Dividendos.** En el mes 4 de cada año del modelo se paga un porcentaje del resultado del año anterior (0%, 0%, 30%, 50% y luego 100%), solo con caja sobre el mínimo.
- **Caja negativa.** No se cubre con fondos ficticios. Se informa como «Aporte adicional necesario». En el VAN de los socios se trata como un aporte implícito, a costo Ke, que se devuelve cuando la caja se recupera.
- **VAN del proyecto.** Usa el flujo libre (FCFF) a WACC 9,49% y **no cambia con el financiamiento**. El VAN de los socios (aportes, dividendos y valor final) se calcula a Ke 16,32%.
- **Indicadores de deuda:**
  - El DSCR mínimo se mide desde el segundo año de operación, como flujo operativo / (intereses + cuotas).
  - El valor final de liquidación incluye la recuperación del capital de trabajo y del activo fijo.

## 5. Tarea 1: validar el libro en Excel

Recalculá el libro (Ctrl+Alt+F9) y compará la columna D de Resultados con estos valores, obtenidos en LibreOffice. Tolerancia: ±1 USD en montos y ±0,01 puntos en tasas.

| Indicador (Resultados, columna D) | Valor esperado |
|---|---|
| Ventas totales del horizonte | 81.222.448 |
| EBITDA total / EBITDA del año 1 | 16.309.504 / 509.970 |
| Resultado neto total | 14.751.246 |
| Inversión total (CAPEX) | 755.044 |
| Necesidad máxima de fondos antes de financiamiento | 3.207.787 (mes 13) |
| Aporte adicional necesario | 311.740 |
| Deuda máxima / intereses totales | 2.791.587 / 813.214 |
| Cancelación de la deuda | 5,00 años |
| Brecha máxima bajo caja mínima / meses con brecha | 373.701 / 12 |
| DSCR mínimo desde el año 2 | 0,40 |
| VAN / TIR del proyecto | 6.989.922 / 34,93% |
| VAN / TIR de los socios | 3.695.678 / 43,63% |
| Recupero de la inversión | 4,13 años |
| Balance de control | 0 |

En Control deben aparecer solo dos alertas:

- C13: brecha de caja, 12 meses;
- C14: negocio en operación sin reales cargados.

Informame cualquier diferencia antes de seguir y no la corrijas sin consultarme.

## 6. Catálogo de escenarios probados (para reproducir o ampliar)

Cada escenario parte de la base y modifica solo celdas vigentes. «Línea X» significa: Inversion!E41 = límite, E42 = tasa, E43 = 1, E44 vacío (renovada hasta el final), salvo que se indique otra cosa.

| Id | Cambios | VAN proyecto | VAN socios | Aporte adicional | Meses bajo caja mínima | Intereses | Cancela (años) | DSCR mín. |
|---|---|---|---|---|---|---|---|---|
| E00 | Base | 6.989.922 | 3.695.678 | 311.740 | 12 | 813.214 | 5,0 | 0,40 |
| F1 | Inversion!E32 = 3.800.000 | 6.989.922 | 3.593.001 | 0 | 1 | 1.106.974 | 5,0 | 0,29 |
| F2 | Inversion!E32 = 1.800.000 | 6.989.922 | 3.680.996 | 816.529 | 41 | 524.356 | 5,0 | 0,62 |
| F3 | Supuestos!E21 = 180 | 6.230.650 | 2.932.499 | 1.632.135 | 57 | 813.214 | 5,0 | 0,12 |
| F4 | Supuestos!E21 = 0 | 8.508.466 | 3.895.390 | 0 | 0 | 813.214 | 5,0 | 0,96 |
| F5 | F1 + F3 | 6.230.650 | 2.954.155 | 1.397.534 | 48 | 1.106.974 | 5,0 | 0,09 |
| F6 | F2 + F4 | 8.508.466 | 3.975.861 | 0 | 0 | 524.356 | 5,0 | 1,49 |
| F7 | F1 + F4 | 8.508.466 | 3.813.553 | 0 | 0 | 1.106.974 | 5,0 | 0,71 |
| F8 | F2 + F3 | 6.230.650 | 2.810.935 | 2.055.741 | 55 | 524.356 | 5,0 | 0,18 |
| C1 | Productos J, L, N (filas 6-7) = I, K, M × 1,10 | 3.619.713 | 1.330.719 | 1.601.514 | 55 | 813.214 | 5,0 | −0,15 |
| C2 | Ídem × 1,20 | 249.504 | −1.129.352 | 3.186.573 | 104 | 813.214 | 5,0 | −0,69 |
| C3 | Ídem × 1,30 | −3.120.704 | −3.725.162 | 5.685.501 | 117 | 813.214 | 5,0 | −1,24 |
| C4 | C3 + Costos!F6:F31 = E × 1,30 + Costos!F75:F83 = E × 2 | −7.020.642 | −6.753.823 | 10.601.180 | 118 | 813.214 | 5,0 | −1,95 |
| V1 | Productos!D37:M38 = D21:M22 × 0,90 | 5.422.991 | 2.690.586 | 430.519 | 21 | 813.214 | 5,0 | 0,22 |
| V2 | Ídem × 0,80 | 3.856.060 | 1.656.795 | 656.645 | 33 | 813.214 | 5,0 | 0,03 |
| V3 | Ídem × 0,70 | 2.289.130 | 626.407 | 882.772 | 45 | 813.214 | 5,0 | −0,15 |
| V4 | Ídem × 1,10 (activa C10: supera la capacidad) | 8.556.853 | 4.713.191 | 290.460 | 8 | 813.214 | 5,0 | 0,58 |
| V5 | Años 1-3: gabinetes 15.000/15.000/21.000; cajas 90.000/90.000/120.000 | 7.601.969 | 4.106.060 | 378.600 | 8 | 813.214 | 5,0 | 0,76 |
| P1 | Volumen × 0,70; Productos!H6:H7 = G × 0,95; Supuestos!E6 = 7; Inversion!J6:J19 = D × 1,15 | −105.270 | −1.158.116 | 2.635.012 | 94 | 813.214 | 5,0 | −0,58 |
| P2 | Inicio!C13 = «Conservador» | −486.490 | −1.723.354 | 3.802.828 | 117 | 813.214 | 5,0 | −0,81 |
| P3 | Inicio!C13 = «Optimista» | 10.466.057 | 5.389.993 | 0 | 0 | 813.214 | 5,0 | 1,04 |
| L1 | Inversion!E36 = 30 | 6.989.922 | 3.636.649 | 1.361.878 | 34 | 508.039 | 3,0 | 0,25 |
| L2 | Inversion!E36 = 78 | 6.989.922 | 3.696.158 | 16.353 | 2 | 1.134.635 | 7,0 | 0,53 |
| L3 | Inversion!E35 = 12 y E36 = 84 | 6.989.922 | 3.645.606 | 0 | 0 | 1.355.681 | 8,0 | 0,55 |
| L4 | E32 = 0 + línea 4.000.000 al 9% | 6.989.922 | 3.760.141 | 0 | 0 | 642.648 | 4,0 | 1,56 |
| K1 | E32 = 3.250.000 | 6.989.922 | 3.654.336 | 97.745 | 7 | 946.754 | 5,0 | 0,34 |
| K2 | Préstamo 2 (G32:G36): 500.000, mes 25, 10%, gracia 0, 12 cuotas | 6.989.922 | 3.691.088 | 234.637 | 4 | 843.459 | 5,0 | 0,40 |
| K3 | Línea 1.000.000 al 10% (uso máximo ≈ 374 mil) | 6.989.922 | 3.698.760 | 0 | 0 | 833.792 | 5,0 | 0,40 |
| K4 | E32 = 0 + línea 4.000.000 al 10% | 6.989.922 | 3.711.577 | 0 | 0 | 734.934 | 4,1 | 1,38 |
| K5 | K4 con vencimiento en el mes 24 (E44 = 24) | 6.989.922 | 3.645.096 | 2.191.679 | 22 | 458.820 | 2,1 | 1,38 |
| K6 | E32 = 760.000 + línea 3.000.000 al 10% | 6.989.922 | 3.739.865 | 0 | 0 | 726.417 | 5,0 | 0,84 |
| K7 | K4 con tasa de la línea 7,99% | 6.989.922 | 3.806.552 | 0 | 0 | 554.410 | 3,9 | 1,79 |
| K8 | K4 con tasa de la línea 11,24% | 6.989.922 | 3.646.770 | 0 | 0 | 857.501 | 4,2 | 1,20 |
| T1 | Supuestos!D51:M51 = 8.760 | 7.989.367 | 4.356.428 | 13.960 | 2 | 813.214 | 5,0 | 0,58 |
| T2 | Ídem = 6.145 | 5.862.801 | 2.934.837 | 707.018 | 35 | 813.214 | 5,0 | 0,19 |
| T3 | Ídem = 5.840 | 5.490.753 | 2.660.071 | 872.928 | 39 | 813.214 | 5,0 | 0,13 |

Verificaciones que ya pasaron en los 36 escenarios (385 comprobaciones):

- balance = 0;
- caja final igual a la suma de flujos;
- el presupuesto original no cambia;
- el financiamiento no altera el EBITDA, el FCFF ni el VAN del proyecto;
- cuotas pagadas = monto prestado;
- interés = saldo × tasa / 12;
- la línea respeta su límite y se cancela al vencer;
- los días de cobro no alteran el EBITDA;
- el TC no altera ventas ni materiales;
- se activan las alertas C10 y C13.

## 7. Conclusiones ya obtenidas (no las contradigas sin evidencia nueva)

- **Variables más sensibles:**
  - costo unitario: +10% en costos variables resta 3,37 M de VAN, y con +20,7% el VAN se anula;
  - plazo de cobro: cada día equivale a USD 15 a 17 mil de caja;
  - tipo de cambio: cada 1% de apreciación del guaraní resta USD 71 a 75 mil de VAN.
- **El volumen solo resiste más:** con −30% de volumen el VAN sigue siendo positivo. El escenario combinado P1 deja la TIR (9,0%) por debajo del WACC.
- **Repago:** la capacidad real es de unos 4 años si todo el excedente va a deuda. El cronograma 6 + 54 choca con los años 2 y 3 (DSCR 0,40 y 0,96): el problema es el perfil de pagos, no la capacidad.
- **Financiamiento:**
  - largo plazo para el CAPEX y el capital de trabajo permanente, que pasa de 2,4 M a 4,5 M;
  - una línea rotativa comprometida de USD 0,5 a 1,0 M para el bache de los años 2 a 4;
  - no financiar el capital de trabajo permanente con corto plazo, por el riesgo de no renovación (K5: faltan 2,2 M en el mes 25).
- **Vender más crea valor, pero primero consume caja** (+0,39 M de necesidad de fondos en V5). Superar la capacidad requiere un CAPEX que no está modelado.

## 8. Supuestos y limitaciones

- Supuestos del plan de 2025, sin reales. WACC 9,49% y Ke 16,32% son históricos.
- Tasa de largo plazo 9% (plan). Tasa de corto plazo 10% ilustrativa.
- Referencias del BCP (tomadas de un buscador; no pude abrir el sitio del BCP para confirmarlas): tasa activa promedio en ME 7,99% (marzo de 2026) y tope legal en ME 11,24% (julio de 2026).
- Los tamaños de los shocks son ilustrativos, sin base histórica.
- El TC de Gs 6.145 proviene del archivo de Alianza del Acero (junio de 2026), no de una fuente oficial.
- La caja positiva no genera intereses.
- No hay comisiones bancarias ni garantías.
- El volumen sobre la capacidad solo genera una alerta; no se limita.
- Pendiente de verificar: base del tributo de maquila (Ley 7547/2025), naturaleza del alquiler, nómina y cargas sociales actuales.

## 9. Próximos pasos en los que quiero ayuda

1. Cargar los reales desde octubre de 2025 en la hoja Reales y actualizar `n_Corte` y `n_InicioReal`.
2. Actualizar WACC, Ke, TC y tasas con ofertas bancarias concretas y repetir los escenarios clave (E00, F3, L2, K3, K5, T2).
3. Preparar una narrativa ejecutiva para la Gerencia Financiera: qué ocurre, por qué, impacto, alternativas, recomendación, riesgos y próximos pasos.
