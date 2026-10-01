# Ecostar v2: pruebas de escenarios

Fecha: 01/10/2026.

- **Archivo:** `Modelo_Maquila_ECOSTAR_v2.xlsx`, con supuestos del plan de octubre de 2025 y sin datos reales.
- **Resultados:** `Pruebas_Escenarios_ECOSTAR_v2.xlsx`, con un resumen, la evolución anual por escenario y las verificaciones.
- **Unidad:** USD, salvo indicación. TC base Gs 7.300.

## 1. Resumen ejecutivo

- **El modelo funciona correctamente en los 36 escenarios probados.** No hay errores de fórmula y el balance cuadra mes a mes. La caja final concilia con los flujos. El presupuesto original no se altera, y las 385 comprobaciones automáticas (23 tipos) dieron OK.
- **El financiamiento no cambia la rentabilidad del proyecto:** el VAN es USD 6,99 M en todos los casos de financiamiento. Lo que cambia es la caja, el costo financiero y el retorno de los socios.
- **Las variables que más mueven el resultado son el costo unitario y el plazo de cobro.**
  - Un 10% más de costos variables resta USD 3,37 M de VAN. El VAN se anula con +20,7%.
  - Cobrar a 180 días en lugar de 120 resta USD 0,76 M de VAN y agrega USD 1,0 M de necesidad de fondos.
  - Cobrar al contado suma USD 1,52 M de VAN y elimina la brecha de caja.
- **Repago.** Con el cronograma actual (6 meses de gracia + 54 cuotas) el préstamo se cancela en 5 años. Pero en los años 2 y 3 la operación no alcanza a pagar las cuotas (DSCR 0,40 y 0,96), y queda un faltante de USD 312 mil. Si todo el excedente se destina a deuda, la capacidad real de repago es de **4 años**. El problema es el perfil de pagos, no la capacidad total.
- **Recomendación de financiamiento:**
  - Financiar la inversión y el capital de trabajo permanente con largo plazo.
  - Agregar una línea rotativa comprometida de USD 0,5 a 1,0 M para cubrir el bache de los años 2 a 4.
  - No financiar el capital de trabajo permanente con corto plazo. Si el banco no renueva la línea, aparece un faltante de USD 2,2 M en el mes 25.

## 2. Verificación del modelo

| Verificación | Escenarios | Resultado |
|---|---|---|
| Sin errores de fórmula después de recalcular | 36 | OK |
| Balance mensual (activo − pasivo − patrimonio = 0) en presupuesto y vigente | 36 | OK |
| Caja final = suma de flujos operativos, de inversión y de financiamiento | 36 | OK |
| El presupuesto original no cambia con ningún escenario | 36 | OK |
| El financiamiento no altera ventas, EBITDA, flujo operativo, CAPEX, FCFF ni el VAN del proyecto | 14 | OK |
| Cuotas pagadas = monto prestado; interés = saldo anterior × tasa / 12 | 14 | OK |
| La línea respeta su límite, se usa solo con caja bajo el mínimo y se cancela al vencer | 36 | OK |
| Los días de cobro no cambian el EBITDA; cuentas por cobrar = ventas × días / (365/12) | 2 | OK |
| El TC no cambia ventas ni materiales (USD); la nómina en Gs escala con 7.300 / TC | 3 | OK |
| Alertas: capacidad excedida (C10) y caja bajo el mínimo (C13) | 26 | OK |
| Atraso del inicio: sin ventas antes del nuevo mes de inicio; el selector Conservador aplica su atraso | 2 | OK |

Además se repitieron las pruebas anteriores con 6 meses reales cargados. Cambiar de escenario no altera los meses cerrados, y el arranque con implantación sigue funcionando.

### Ajustes al modelo que surgieron de las pruebas

1. **Línea de crédito rotativa opcional** (hoja Inversion, sección B2). Se carga con límite, tasa, desde, vencimiento y saldo al corte.
   - Se usa sola cuando la caja cae bajo el mínimo y se devuelve con el excedente.
   - Mientras tenga saldo no se pagan dividendos.
   - Al vencer se paga entera.
   - Sin esta línea no se podía comparar corto con largo plazo.
2. **El VAN de los socios ahora cobra la caja negativa.**
   - Antes, un faltante de caja se financiaba gratis y «menos financiamiento» parecía mejor para los socios (3,83 M frente a 3,71 M).
   - Ahora el faltante se trata como un aporte implícito de los socios, a costo Ke, que se devuelve cuando la caja se recupera. Queda en 3,68 M frente a 3,70 M.
   - El VAN de los socios de la base pasa de 3,708 M a 3,696 M.
3. **Nuevos indicadores:** aporte adicional necesario, intereses totales, años hasta cancelar la deuda, uso máximo de la línea y DSCR mínimo desde el segundo año de operación. Antes, el DSCR quedaba dominado por el año 1, que es negativo por la formación del capital de trabajo.
4. **Dividendos calculados sin referencias a filas completas.** Así se evita una referencia circular con la línea.

## 3. Resultados por pregunta

### 3.1 Más o menos financiamiento, con distintos plazos de cobro

| Escenario | VAN socios | Necesidad de fondos antes de financiar | Aporte adicional necesario | Meses bajo caja mínima | Intereses |
|---|---|---|---|---|---|
| Base: LP 2,79 M, cobro 120 días | 3,70 M | 3,21 M | 0,31 M | 12 | 0,81 M |
| Más financiamiento: LP 3,8 M | 3,59 M | 3,21 M | 0 | 1 | 1,11 M |
| Menos financiamiento: LP 1,8 M | 3,68 M | 3,21 M | 0,82 M | 41 | 0,52 M |
| Cobro a 180 días | 2,93 M | 4,22 M | 1,63 M | 57 | 0,81 M |
| Cobro al contado | 3,90 M | 1,41 M | 0 | 0 | 0,81 M |
| Más financiamiento + 180 días | 2,95 M | 4,22 M | 1,40 M | 48 | 1,11 M |
| Menos financiamiento + contado | **3,98 M** | 1,41 M | 0 | 0 | 0,52 M |
| Más financiamiento + contado | 3,81 M | 1,41 M | 0 | 0 | 1,11 M |
| Menos financiamiento + 180 días | 2,81 M | 4,22 M | **2,06 M** | 55 | 0,52 M |

Qué muestran las cifras:

- **El plazo de cobro define cuánta plata necesita el negocio.** La necesidad de fondos va de 1,41 M (contado) a 4,22 M (180 días). Cada día de cobro inmoviliza entre USD 15 y 17 mil.
- **Más deuda tapa el bache, pero cuesta.** Un millón adicional de préstamo elimina casi toda la brecha, pero agrega USD 294 mil de intereses y baja el VAN de los socios en 0,10 M. El modelo no paga intereses por la caja ociosa.
- **Combinación más eficiente: menos deuda y cobro al contado.** Si el cobro se estira a 180 días, ni siquiera un millón más de deuda alcanza: faltan USD 1,40 M.
- **Rentabilidad y liquidez son cosas distintas.** Con 180 días el proyecto sigue siendo rentable (VAN 6,23 M, TIR 27,2%), pero pasa 57 meses con caja bajo el mínimo.

### 3.2 Costos que no se cubren con las ventas

| Escenario | EBITDA 10 años | VAN proyecto | TIR proyecto | VAN socios | Aporte adicional necesario |
|---|---|---|---|---|---|
| Base | 16,31 M | 6,99 M | 34,9% | 3,70 M | 0,31 M |
| Costos variables +10% | 11,13 M | 3,62 M | 21,8% | 1,33 M | 1,60 M |
| Costos variables +20% | 5,94 M | 0,25 M | 10,3% | −1,13 M | 3,19 M |
| Costos variables +30% | 0,76 M | −3,12 M | 0,0% | −3,73 M | 5,69 M |
| Variables +30%, salarios +30% y fijos ×2 | −5,25 M | −7,02 M | −10,7% | −6,75 M | 10,60 M |

- **Muy sensible al costo unitario.** Los costos variables equivalen al 61% de las ventas. Por eso cada 1% de aumento resta unos USD 337 mil de VAN, y con +20,7% el VAN del proyecto se anula. Frente al costo de capital de los socios el margen es aún menor: con +20% su VAN ya es negativo.
- **El modelo responde bien cuando los costos superan las ventas.** En el caso extremo el EBITDA es negativo todos los años, la TIR es negativa, no hay recupero y se necesitan USD 10,6 M adicionales. El modelo no inventa fondos: lo muestra como caja negativa, y Control lo marca en 118 meses.

### 3.3 Si el plan no sale como se esperaba

| Escenario | Ventas 10 años | VAN proyecto | TIR proyecto | VAN socios | Aporte adicional | Recupero |
|---|---|---|---|---|---|---|
| Volumen −10% | 73,1 M | 5,42 M | 30,6% | 2,69 M | 0,43 M | 4,5 años |
| Volumen −20% | 65,0 M | 3,86 M | 25,7% | 1,66 M | 0,66 M | 5,1 años |
| Volumen −30% | 56,9 M | 2,29 M | 20,0% | 0,63 M | 0,88 M | 6,0 años |
| Combinado: volumen −30%, precio −5%, inicio +6 meses, CAPEX +15% | 50,9 M | −0,11 M | 9,0% | −1,16 M | 2,64 M | 9,6 años |
| Selector «Conservador» del modelo | 67,5 M | −0,49 M | 7,9% | −1,72 M | 3,80 M | 9,8 años |
| Selector «Optimista» del modelo | 87,0 M | 10,47 M | 57,7% | 5,39 M | 0 | 2,8 años |

- **El volumen solo resiste más que el precio y el costo.** Con −30% de volumen el proyecto sigue creando valor. Extrapolando, el VAN se anularía cerca de −45%.
- **La combinación es la que destruye valor.** Con menor volumen, menor precio, atraso y sobrecosto a la vez, la TIR (9,0%) queda debajo del WACC (9,49%) y los socios deben aportar USD 2,6 M más.

### 3.4 Si vendo más

| Escenario | VAN proyecto | VAN socios | Necesidad de fondos | Aporte adicional | Observación |
|---|---|---|---|---|---|
| Volumen +10% | 8,56 M | 4,71 M | 3,33 M | 0,29 M | Desde el año 4 supera la capacidad de gabinetes. Control lo marca (C10). El CAPEX de ampliación **no** está incluido, así que el resultado está sobrestimado. |
| Rampa acelerada dentro de la capacidad (años 1-2 a capacidad inicial) | 7,60 M | 4,11 M | 3,59 M | 0,38 M | EBITDA del año 1: 0,93 M frente a 0,51 M. |

Vender más crea valor, pero primero consume caja. Con 120 días de cobro y 120 de inventario, crecer más rápido aumenta la necesidad de fondos en USD 0,39 M. Crecer también requiere financiamiento.

### 3.5 ¿En cuánto tiempo puedo devolver el préstamo?

| Estructura (USD 2,79 M al 9%) | Cancelación | DSCR mínimo desde el año 2 | Aporte adicional necesario | Intereses |
|---|---|---|---|---|
| 6 meses de gracia + 30 cuotas | 3 años | 0,25 | 1,36 M | 0,51 M |
| **6 + 54 cuotas (plan)** | **5 años** | **0,40** | **0,31 M** | **0,81 M** |
| 6 + 78 cuotas | 7 años | 0,53 | 0,02 M | 1,14 M |
| 12 + 84 cuotas | 8 años | 0,55 | 0 | 1,36 M |
| Repago con todo el excedente (sin cuota fija) | **4 años** | — | 0 | 0,64 M |

- **Capacidad real: unos 4 años.** El negocio puede cancelar la deuda en 4 años si todo el excedente de caja va a pagarla y no se pagan dividendos hasta entonces.
- **El cronograma fijo choca con los años de menor caja.** Las cuotas de los años 2 y 3 llegan cuando el capital de trabajo todavía absorbe caja: el DSCR es 0,40 en el año 2 y 0,96 en el año 3. Un DSCR menor a 1 significa que la operación no alcanza a pagar las cuotas de ese año.
- **Cómo cerrar el bache:**
  - alargar el plazo a 7-8 años, con USD 0,32 a 0,54 M más de intereses;
  - negociar cuotas crecientes o más gracia;
  - mantener 5 años y sumar una línea rotativa (sección 3.6).
- **El plazo a 3 años no es viable** sin USD 1,36 M adicionales.

### 3.6 Corto plazo, largo plazo o ambos a la par

Tasas: LP al 9% (plan) y CP al 10% (ilustrativa). Como referencia, el BCP informa una tasa activa promedio ponderada en moneda extranjera de 7,99% (marzo de 2026) y un tope legal en ME de 11,24% (julio de 2026).

| Estructura | VAN socios | Intereses | Uso máximo de la línea | Aporte adicional | Cancelación |
|---|---|---|---|---|---|
| Base: solo LP 2,79 M | 3,70 M | 0,81 M | — | 0,31 M | 5 años |
| Solo LP, ampliado a 3,25 M | 3,65 M | 0,95 M | — | 0,10 M | 5 años |
| LP + préstamo CP de USD 500 mil a 12 meses | 3,69 M | 0,84 M | — | 0,23 M | 5 años |
| **LP + línea rotativa USD 1 M a la par** | **3,70 M** | **0,83 M** | **0,37 M** | **0** | 5 años |
| LP solo para CAPEX (0,76 M) + línea USD 3 M | 3,74 M | 0,73 M | 1,72 M | 0 | 5 años |
| Solo CP: línea USD 4 M renovada al 10% | 3,71 M | 0,74 M | 2,36 M | 0 | 4,1 años |
| Solo CP al 7,99% / al 11,24% | 3,81 M / 3,65 M | 0,55 M / 0,86 M | 2,31 M / 2,39 M | 0 | 3,9 / 4,2 años |
| **Solo CP sin renovación (vence en el mes 24)** | 3,64 M | 0,46 M | 2,36 M | **2,19 M** | 2,1 años |

- **¿Conviene el corto plazo para la operación?** Solo para la parte transitoria de la necesidad.
  - Un préstamo CP con cuotas fijas no sirve: su repago cae justo en el bache y la brecha sigue (USD 0,23 M).
  - Una línea rotativa sí sirve: se usa cuando falta caja y se devuelve con el excedente.
- **¿Conviene el largo plazo?** Sí, para el CAPEX y el capital de trabajo permanente, que crece de 2,4 M a 4,5 M y no se recupera hasta el final.
  - Financiar todo con corto plazo cuesta casi lo mismo (±0,1 M de VAN según la tasa).
  - Pero depende de que el banco renueve una línea de 2,4 M durante 4 años. Si no la renueva en el mes 24, faltan USD 2,2 M de golpe.
- **¿LP y CP a la par?** Sí, y es la mejor relación costo/riesgo en este caso. Con el LP del plan más una línea de USD 1 M se elimina la brecha con solo USD 21 mil más de intereses: la línea llega a usarse USD 374 mil.

### 3.7 Tipo de cambio

| TC (Gs/USD) | EBITDA 10 años | VAN proyecto | VAN socios | Aporte adicional | Meses bajo caja mínima |
|---|---|---|---|---|---|
| 8.760 (+20%) | 17,84 M | 7,99 M | 4,36 M | 0,01 M | 2 |
| 7.300 (base) | 16,31 M | 6,99 M | 3,70 M | 0,31 M | 12 |
| 6.145 (−16%, nivel del archivo de Alianza, jun-2026) | 14,59 M | 5,86 M | 2,93 M | 0,71 M | 35 |
| 5.840 (−20%) | 14,02 M | 5,49 M | 2,66 M | 0,87 M | 39 |

- **Un TC más alto favorece a Ecostar.** Ventas y materiales están en USD, mientras que la nómina y parte de los costos fijos están en guaraníes.
- **Cada 1% de apreciación del guaraní resta unos USD 71 a 75 mil de VAN.** El efecto no es simétrico: al subir el TC el beneficio es menor que la pérdida al bajar.
- **Es un riesgo actual, no hipotético.** El archivo de Alianza ya usa Gs 6.145 (−16% frente al plan). Con ese TC la rentabilidad sigue siendo positiva, pero el faltante de caja sube a USD 0,71 M.

## 4. Supuestos

| Supuesto | Valor | Tipo |
|---|---|---|
| Plan base | Ecostar, octubre de 2025: ventas 81,2 M, CAPEX 0,76 M, LP 2,79 M al 9% (6 + 54), aportes 1,14 M | Plan (histórico) |
| WACC / Ke | 9,49% / 16,32% | Plan (histórico, no actualizado) |
| Tasa del LP | 9% nominal anual | Plan |
| Tasa del CP | 10%, con sensibilidad a 7,99% y 11,24% | Ilustrativa. Las referencias son del BCP, pero no son ofertas bancarias. |
| Tamaño de los shocks | ±1 M de deuda; 180 días o contado; ±10-30% de costos y volumen; ±20% de TC | Ilustrativos, sin base histórica de Ecostar |
| TC 6.145 | Nivel del archivo de Alianza del Acero (junio de 2026) | Dato de un archivo interno, no oficial |
| Caja mínima | 3 días de ventas | Plan |
| Dividendos | 0%, 0%, 30%, 50%, 100% del resultado del año anterior; solo con caja sobre el mínimo y sin saldo en la línea | Plan + criterio del modelo |

## 5. Riesgos y limitaciones

- Los resultados usan supuestos de 2025 sin datos reales. No describen la situación actual de Ecostar.
- El modelo no paga intereses por la caja positiva, así que el exceso de deuda solo suma costo. Tampoco incluye comisiones de apertura, garantías ni covenants.
- La línea rotativa supone renovación automática hasta su vencimiento. El riesgo de no renovación se mide con el escenario sin renovación.
- El volumen sobre la capacidad no se limita, solo se alerta: «vender más +10%» no incluye el CAPEX de ampliación.
- WACC y Ke son históricos. Para decidir el financiamiento hay que actualizarlos y pedir ofertas bancarias concretas (tasa, plazo, gracia, garantías, comisiones).

## 6. Recomendación y próximos pasos

1. **Estructura de financiamiento.** Mantener el LP para CAPEX y capital de trabajo permanente. Negociar además una de estas dos opciones:
   - gracia de 12 meses o cuotas crecientes;
   - una línea rotativa comprometida de USD 0,5 a 1,0 M para los años 2 a 4.

   No financiar el capital de trabajo permanente con corto plazo.
2. **Cobranza.** Cada día de cobro vale entre USD 15 y 17 mil de caja. Hay que confirmar en el contrato CIE–PERG si existen anticipos, factoring o descuento de facturas.
3. **Costo unitario.** Con un margen de solo +20% antes de destruir valor, conviene fijar precios de materiales o indexar la tarifa.
4. **TC.** Medir la exposición real de la nómina y los costos en guaraníes con el TC actual, y evaluar coberturas o una cláusula de ajuste.
5. **Pendientes.** Cargar los reales desde octubre de 2025, las tasas vigentes y las ofertas de bancos, y repetir esta batería con `probar_escenarios.py`.

Fuentes de tasas:

- [BCP — Indicadores Financieros, marzo de 2026](https://www.bcp.gov.py/documents/20117/2322846/IF+informe+Marzo-26.pdf/f6399681-1915-a78b-99fc-7f5c0013373d?t=1778003766675)
- [BCP — Indicadores Financieros, enero de 2026](https://www.bcp.gov.py/documents/20117/2322846/IF+informe+Ene-26.pdf/9d37c85e-2364-beeb-41d9-12119c847b4a?t=1772718775176)
- [La Nación — límites de tasas, julio de 2026](https://www.lanacion.com.py/negocios/2026/07/02/limite-de-intereses-para-creditos-y-tarjetas-se-mantiene-casi-invariable/)
- [Mentu — BCP fija nuevos límites de tasas, julio de 2026](https://mentu.com.py/2026/07/01/bcp-fija-nuevos-limites-de-tasas-para-prestamos-y-tarjetas-de-credito/)

Las cifras se tomaron del resumen del buscador; el sitio del BCP no fue accesible desde este entorno y conviene confirmarlas.
