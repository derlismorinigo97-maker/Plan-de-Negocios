# Plantilla común de maquilas v2: evaluación de Alianza del Acero y simplificación de Ecostar

Fecha: 01/10/2026. Archivos:

- `Modelo_Maquila_ECOSTAR_v2.xlsx`: Ecostar cargado en la plantilla común.
- `Plantilla_Maquila_Comun.xlsx`: plantilla vacía, para Alianza del Acero y la tercera maquila.
- `versiones/v1_Modelo_Financiero_Maquila_CIE_PERG_detallado.xlsx`: versión previa (15 hojas), sin cambios, para comparación.
- `modelo/plantilla/`: generador y datos de carga (`generar_plantilla.py`, `datos_ecostar.py`).

## 1. Conclusión

El informe de Alianza del Acero sirve como **referencia funcional**: muestra qué necesita ver la gerencia (ejecutado vs plan, desvíos, flujo sin financiamiento, payback). Sin embargo, **no sirve como base de cálculo**, por cuatro motivos:

- depende de archivos externos;
- tiene errores `#REF!` y textos de payback tipeados;
- tiene montos y factores embebidos en fórmulas;
- el presupuesto original quedó sobrescrito por la ejecución.

Además, su VAN mezcla el flujo del proyecto con el del financiamiento.

La plantilla v2 toma esas salidas y las construye sobre un motor único, con estas características:

- 9 hojas;
- cada dato se carga una sola vez;
- presupuesto original, reales y proyección vigente separados;
- los escenarios solo afectan los meses posteriores al último cierre.

Ecostar v2 reproduce las ventas del plan sin diferencias. El EBITDA total difiere −0,05% de la v1 depurada. El VAN del proyecto es 3,8% mayor que en la v1 por la simplificación del capital de trabajo (explicada en la sección 5).

## 2. Evaluación del informe de Alianza del Acero

### Estructura y lógica

INFORME2 tiene 27 hojas: 6 de ejecución (Ej_) y 21 del plan.

- **Cadena del plan:** Supuestos → Productos → Producción → Ingresos/Costos → Resultados → Flujo de Fondos → VAN-TIR/Payback.
- **Hojas Ej_:** comparan lo ejecutado con el plan mes a mes.

La lógica general es razonable, pero la información está dispersa y repetida, y no existe un mes de corte que separe lo ejecutado de lo proyectado.

### Verificación de cálculos

| Punto | Resultado de la verificación |
|---|---|
| VAN USD 2.844.948 y TIR 14,92% | Reproducidos exactamente. El flujo incluye préstamos (flujo con financiamiento) y se descuenta al WACC de 9,87%, lo que mezcla dos criterios. |
| Recupero del capital de trabajo en el año 10 (USD 10.928.041) | Aporta USD 4,26 M de valor presente. **Sin él, el VAN sería negativo.** |
| Flujo sin financiamiento (FCFF, fila 28 de INFORME2) al 9,87% | VAN USD 3,46 M sin valor terminal (TIR 19,99%). Con el recupero del capital de trabajo: USD 7,73 M (TIR 25,6%). |
| Capital de trabajo inicial USD 4.024.908 | Es un valor de cierre: préstamo + aporte − inversión fija. El capital de trabajo calculado para el año 1 es USD 1.541.094. La caja del año 0 queda en −1.984.780 porque los préstamos se desembolsan en el año 1. |
| Payback | ORIGINAL e INFORME1 tienen 11 celdas `#REF!`. El Dashboard muestra textos tipeados («6 Años», «4 Años», «7 Años») que no coinciden con las tablas. INFORME2 descuenta el flujo del proyecto a Ke en lugar del WACC. |
| Error de costos en INFORME1 | `Ej_Costos G37:N42` apunta a una celda vacía (P2). Por eso anula insumos menores, packaging y movimiento de carga de mayo a diciembre: EBITDA año 1 de 33.264 en lugar de −23.860 y VAN 2,897 M en lugar de 2,845 M. Está corregido en INFORME2. |
| PDF (enero 2026) vs Excel | TIR con financiamiento 15,19% y VAN al 9,87% de 2.843.528. TIR sin financiamiento 28,16% y VAN al 16,32% de 3.926.408. Las tasas están invertidas respecto de la práctica habitual: el flujo de los socios debe descontarse a Ke y el del proyecto al WACC. |
| Días de capital de trabajo y TC | Días del PDF (stock / producto terminado / clientes / proveedores menores): 90 / 60 / 0 / 34. Días del Excel actual: 45 / 0 / 90 / 15. El TC pasó de 7.800 a 6.145 sin documentar el cambio. |

### Ejecución frente al plan (año 1)

- **Ventas:** ejecutado más proyectado USD 4,21 M, frente a 10,82 M del plan (−61%).
- **Producción:** 39% de lo planificado.
- **EBITDA:** −24 mil frente a +841 mil del plan.
- **Proyección del año 2:** 28% por encima del plan original, optimista frente a la ejecución observada.
- **Flujo real (Ej_Flujo):** incluye préstamos de corto plazo para acero (USD 3 a 10,9 M por año), devoluciones a CIE S.A. y costos de implantación que el plan no contempla.

### Qué conviene aprovechar

1. Las columnas «originales vs actuales» de Supuestos. En la plantilla v2 pasan a ser «presupuesto original / vigente» en todas las entradas.
2. El seguimiento mensual ejecutado vs plan con desvío %. En v2 lo cubren la hoja Reales y la sección D de Resultados.
3. El dashboard de ejecución acumulada. En v2, el estado de cada año se marca como Real, Real+Proy. o Proyección.
4. La tabla de productos con precio y costo por unidad (toneladas, en el caso de Alianza) y la merma entre producción y venta.
5. El crecimiento de precios por etapa, que en v2 es un vector por año.
6. El flujo sin financiamiento (FCFF) agregado en INFORME2 y el payback.

### Qué requiere corrección (antes de migrar Alianza)

1. **Vínculos externos.** Hay 4 vínculos a archivos del PC del planificador anterior, entre ellos `CALCULADORA PRESTAMO.xlsx`. El préstamo 2 (`Servicio Deuda!D18:I20`) y el nombre `SCS` dependen de ellos.
2. **Supuestos sin uso.** Hay 15 supuestos que no se usan, entre ellos el horizonte, la tasa y plazo del préstamo 2, el aporte de CIE, el alquiler, el agua, el seguro de caución, los IVA y la tasa del crédito de corto plazo. El TC de Supuestos (6.145) solo se usa para mostrar valores.
3. **Datos embebidos en fórmulas.** Hay 76 fórmulas con factores embebidos (×1,1, ×1,05, /1,42, 20%×1,8%). Además, 31 montos en guaraníes están convertidos con un único TC (6.028, cierre de mayo) y hay 2.797 números tipeados en las hojas Ej_.
4. **Inconsistencias internas:**
   - la mano de obra se divide por 1,42, mientras que el parámetro de cargas es 43,09%;
   - el seguro de caución se calcula como 20% × 1,8% de la materia prima, mientras que el parámetro es 1% de las ventas;
   - la depreciación del año 1 sale de Ej_, mientras que la nota dice que empieza en el año 2;
   - el resultado incluye intereses del crédito rotativo, pero el flujo del VAN no;
   - los dividendos se corrigen a mano;
   - hay una celda suelta en `Ingresos!O26`.
5. **Presupuesto original sobrescrito.** En los años 1 y 2 el plan fue reemplazado por la ejecución. El original solo sobrevive en la columna E de Supuestos y en las columnas P/AE de las hojas Ej_.

### Información faltante para Alianza

- Mes de corte y saldos reales al cierre: caja, clientes, inventarios, proveedores, deuda.
- Condiciones de los préstamos 2 y de las líneas de corto plazo para acero (hoy están en el archivo externo).
- Detalle de las devoluciones a CIE S.A. y de los costos de implantación no presupuestados.
- Respaldo del cambio de días de capital de trabajo y de TC entre el PDF y el Excel.
- Tasas vigentes (WACC, Ke), dividendos efectivamente pagados y nómina actual.

## 3. Diseño de la plantilla común v2

| Hoja | Contenido | ¿Se carga? |
|---|---|---|
| Inicio | Nombre, etapa (Implantación/Operación), fecha del mes 1, último mes cerrado, mes real de inicio, escenario. Estado, instrucciones y pendientes. | Sí |
| Supuestos | Parámetros generales, supuestos por año (TC, variaciones de precio/costos/salarios, dividendos), escenarios y referencia del plan publicado. | Sí |
| Productos | Hasta 25 productos individuales: precio, materia prima, insumos y otros variables por unidad; capacidad y volúmenes por año. | Sí |
| Costos | Hasta 30 puestos (clase, moneda, salario, dotación por año) y 25 costos fijos (moneda, crecimiento, año de inicio, IVA, base de imprevistos). | Sí |
| Inversion | Hasta 20 ítems de CAPEX (presupuesto, ejecutado, pagado, comprometido, por contratar), 3 préstamos, una línea rotativa y 10 aportes. | Sí |
| Reales | Un dato por celda y por mes: TC, unidades y ventas por producto, 7 rubros de costo, CAPEX, financiamiento y saldos al cierre. | Sí (mensual) |
| Calculo | Motor mensual de 144 meses. El bloque P (presupuesto) y el bloque V (vigente) usan las mismas fórmulas. | No |
| Resultados | Indicadores, resumen anual vigente y presupuesto, desvíos, comparación con el plan publicado y gráficos. | No |
| Control | 16 validaciones con acción sugerida. | No |

Principios de diseño:

- **Un dato, una celda.** La columna «Vigente» hereda el presupuesto con una fórmula (en gris). Se sobrescribe solo lo que cambió; la celda se resalta y Control la cuenta.
- **Tres capas separadas:**
  - Presupuesto original: nunca se modifica.
  - Real: hoja Reales, hasta el último mes cerrado.
  - Proyección vigente: supuestos vigentes y escenario, desde el mes siguiente al corte.
- **Escenarios que no alteran la historia.** Ajustan volumen, precio, materiales, TC, atraso del inicio, sobrecosto del CAPEX pendiente, días de cobro e inventario. Solo se aplican después del corte. Se verificó que los meses reales, la depreciación y los saldos no cambian al pasar de Vigente a Conservador.
- **Multimoneda.** Salarios, costos fijos y reales pueden cargarse en PYG o USD. Se convierten con el TC de cada año, o con el TC real del mes en los meses cerrados.
- **Operación o implantación.** El mes de inicio define los meses de implantación. En ellos hay personal parcial (porcentaje configurable), solo corren los costos fijos marcados «desde año 0» y no hay ventas. El CAPEX se paga en el mes indicado y se deprecia desde la puesta en servicio.
- **Sin fondos ficticios.** Si la caja cae por debajo del mínimo, la brecha queda visible y Control la marca.

## 4. Cambios respecto de la versión 1 (Ecostar detallado)

| Versión 1 (15 hojas) | Versión 2 (9 hojas) |
|---|---|
| Interruptores C1 a C6 de correcciones metodológicas | Correcciones incorporadas como método: TC único en la nómina, 1% acumulativo de materia prima, depreciación desde la puesta en servicio, CAPEX del año 2 depreciable, indirectos de fábrica como costo de producción. |
| Base_PDF con la reconstrucción de los cuadros 1-20 | Queda en la v1 conservada. La v2 compara contra el plan publicado (ventas, EBITDA, resultado, VAN, TIR) en Supuestos D y Resultados E. |
| Capital de trabajo por cohortes (compras anticipadas, producto terminado en unidades, recuperos) | Capital de trabajo por días sobre el mes corriente. En el mes de corte se reemplaza por los saldos reales. |
| 5 escenarios + 15 pruebas de sensibilidad + capturas | 3 escenarios (Vigente, Conservador, Optimista) con 8 ajustes editables. |
| Línea de crédito hipotética y aportes automáticos | Sin aportes automáticos: la brecha queda visible. Línea rotativa opcional con límite, tasa y vencimiento (agregada tras las pruebas de escenarios). |
| Valor terminal por liquidación o perpetuidad | Solo liquidación (capital de trabajo y activo fijo realizables, en %). |
| Datos_Reales en tablas por tipo | Reales en una sola hoja, un dato por celda, con las mismas columnas de mes que el motor. |
| 34 controles + 25 pruebas del motor | 16 controles de integridad, carga y consistencia. |

## 5. Resultados de Ecostar v2 frente a la v1 y al plan publicado

Supuestos de octubre de 2025, sin datos reales cargados (presupuesto = vigente).

| Indicador (USD) | Plan publicado | v1 depurada | v2 | Comentario |
|---|---|---|---|---|
| Ventas 10 años | 81.222.448 | 81.222.448 | 81.222.448 | Igual (diferencias < USD 1). |
| EBITDA 10 años | 17.354.010 | 16.318.100 | 16.309.504 | La diferencia con el PDF se debe a las correcciones C1 (TC de la nómina) y C2 (materia prima). Frente a la v1: −0,05%. |
| EBITDA año 1 | 573.024 | 526.818 | 509.970 | Diferencia de período: en la v1 la producción del año 1 supera las ventas (12.411 vs 12.000 gabinetes), así que parte del costo queda en inventario hasta el año siguiente. En la v2 el costo se reconoce en el mes. |
| VAN del proyecto (FCFF, 9,49%) | 4.588.736 (*) | 6.732.541 | 6.989.922 | Capital de trabajo máximo: v2 4,50 M frente a v1 4,94 M (por días y no por cohortes). |
| TIR del proyecto | 22,61% (*) | 32,3% | 34,9% | |
| VAN de los socios (Ke 16,32%) | — | 3.748.260 | 3.695.678 | La v2 cobra a Ke la caja negativa como aporte implícito de los socios. |
| Brecha de caja máxima | — | 589.308 | 373.701 | Menor capital de trabajo en la v2. Los dividendos solo se pagan con caja disponible. |
| Recupero de la inversión | — | 4,4 años | 4,1 años | |

(*) El VAN y la TIR publicados mezclan el flujo del proyecto con el del accionista, así que no son comparables.

Lectura financiera:

- **Brecha de caja.** El financiamiento del plan cubre el total, pero no el momento: entre el año 3 y el 4 la caja cae por debajo del mínimo hasta en USD 374 mil y llega a ser negativa en USD 312 mil. Lo explica el capital de trabajo de 120 días de cobro y 120 de inventario.
- **DSCR.** El DSCR del año 1 es negativo porque el flujo operativo absorbe USD 2,39 M de capital de trabajo. Desde el segundo año, el mínimo es 0,40: en los años 2 y 3 la operación no alcanza a pagar las cuotas.
- **Peso del valor terminal.** Cerca de USD 1,8 M del VAN (26%) proviene de recuperar el capital de trabajo al final del horizonte.
- **Vigencia de las cifras.** Usan supuestos de 2025 y no acreditan la situación actual.

## 6. Cómo cargar y actualizar

**Alta de un negocio** (Alianza del Acero o la tercera maquila):

1. Copiar `Plantilla_Maquila_Comun.xlsx` y completar la hoja Inicio.
2. Cargar las columnas de presupuesto en Supuestos, Productos, Costos e Inversion. El TC es obligatorio en todos los años.
3. Revisar Control y Resultados.

**Cierre mensual:**

1. En Reales, cargar el mes cerrado: unidades y ventas por producto, 7 rubros de costo, CAPEX, desembolsos, intereses, amortizaciones, aportes y dividendos.
2. En el último mes cerrado, cargar los saldos de caja, clientes, inventarios, IVA, proveedores y deuda.
3. Actualizar «Último mes cerrado» en Inicio.
4. En Inversion, actualizar por ítem lo ejecutado, pagado, comprometido y por contratar. El total pagado debe coincidir con el CAPEX de Reales (control C08).
5. Revisar Control y ajustar los supuestos vigentes que hayan cambiado.

**Escenarios:** elegir el escenario en Inicio. Los ajustes de Conservador y Optimista están en Supuestos C. Son **ilustrativos**: deben reemplazarse por premisas aprobadas.

**Excel:** el libro recalcula al abrir. Se validó con LibreOffice: 0 errores y balance cuadrado en todos los casos de prueba. Se recomienda abrirlo en Excel y guardarlo una vez.

## 7. Criterios y limitaciones

- **Horizonte.** Es común para ambas versiones: desde el inicio presupuestado más los años evaluados. Un atraso reduce los meses operativos dentro de ese horizonte; no lo extiende.
- **Crecimientos y TC.** Los precios, costos y salarios crecen por año de operación. El TC se aplica por año del modelo.
- **Dividendos.** Se pagan en el mes indicado del año siguiente, como porcentaje del resultado. Si no hay caja sobre el mínimo, no se pagan y no se acumulan.
- **Capital de trabajo.** Se calcula por días sobre el mes corriente. El inventario de materiales no incluye flete ni despacho (la v1 sí los incluía: corrección C5).
- **Activo fijo.** El activo fijo neto sigue los pagos. Si hay CAPEX devengado y no pagado, se refleja al pagarse.
- **Préstamos.** Hay tres préstamos con sistema francés y gracia, y una línea rotativa opcional. La línea se usa solo si la caja cae bajo el mínimo y se devuelve con el excedente; al vencer se paga entera. Los usos reales de la línea se cargan en Reales, y su saldo al corte en Inversion.
- **Caja negativa.** No se cubre con fondos ficticios: queda visible como brecha y como «aporte adicional necesario». En el VAN de los socios se trata como aporte implícito, que se devuelve cuando la caja se recupera. La caja positiva no genera intereses.
- **Supuestos a verificar en Ecostar:**
  - tributo de maquila del 1% sobre ventas (según fuentes secundarias, la Ley 7547/2025 fijaría la base en el mayor entre el valor agregado nacional y la factura; pendiente de verificación en fuente oficial);
  - IVA de costos fijos no recuperable;
  - cargas sociales del 43,09% (coeficiente del plan, no validado como tasa legal).

## 8. Próximos pasos

1. **Ecostar.** Completar la fecha del mes 1, el mes real de inicio y los reales desde octubre de 2025, el avance del CAPEX por ítem y las tasas vigentes. Hoy Control marca este punto (C14).
2. **Alianza del Acero.** Cargar como presupuesto original el plan de enero de 2026 (columna E de Supuestos y columnas P/AE de las hojas Ej_):
   - productos en toneladas;
   - reales mensuales desde las hojas Ej_;
   - préstamos 1 y 2 y líneas de corto plazo;
   - sin vínculos externos.
3. **Tercera maquila.** Usar la etapa «Implantación»:
   - mes de inicio según el cronograma;
   - CAPEX por ítem con mes de pago;
   - personal previo al inicio (PreDir y PreInd);
   - costos fijos «desde año 0»;
   - aportes y préstamos.
