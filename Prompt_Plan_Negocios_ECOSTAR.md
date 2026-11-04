# Prompt de contexto — Plan de negocios ECOSTAR (maquila CIE–PERG)

> **Cómo usar este documento:** pegalo como primer mensaje en Claude en Excel (o adjuntalo en PDF) junto con el plan de negocio original («Procesamiento de chapas para gabinetes metálicos – Empresas CIE–PERG», Mujica & Saldivar, octubre 2025, 37 páginas). Resume todos los supuestos del plan, la situación real al 31/10/2026 y las decisiones abiertas.
>
> Fecha de referencia: 04/11/2026. Elaborado por: Derlis Morinigo (Planificación Financiera, CIE S.A.).

---

## 0. Instrucciones para Claude

Vas a ayudarme a armar en este libro de Excel el plan de negocios actualizado de **ECOSTAR S.A.** con los datos que tenemos. Trabajá así:

1. **Respondé en español**, con lenguaje profesional y directo. Usá Excel en español (separador `;`, funciones `SI`, `BUSCARV`, `SUMAPRODUCTO`, etc.).
2. **Separá tres capas de información** y no las mezcles:
   - **Presupuesto original**: el plan de octubre 2025 (sección 2). Es la línea base: no se modifica.
   - **Real**: lo ejecutado hasta el 31/10/2026 (sección 4).
   - **Vigente / proyección**: lo que esperamos desde noviembre 2026 (secciones 5 y 6).
3. **No inventes datos.** Si un dato falta, dejá la celda vacía o marcada como «pendiente», y avisame. Cada supuesto debe tener **fuente** (plan, informe CIE, contrato, supuesto propio) y **fecha**.
4. **Celdas de entrada separadas de las fórmulas.** Entradas en una hoja de supuestos (formato distintivo); el resto, fórmulas que las referencian. Nada de valores pegados dentro de fórmulas.
5. **Porcentajes como fracción** (0,05 = 5%). Moneda de presentación: **USD**. Los costos en guaraníes se convierten con el tipo de cambio de cada período.
6. **Diferenciá** resultado contable, EBITDA, flujo de caja, capital de trabajo, necesidad de financiamiento y creación de valor (VAN). Una TIR alta no implica que el negocio sea financiable.
7. **Incluí controles**: balance que cuadra, caja que arrastra, deuda que se cancela al final del plazo, inventario no negativo, conciliación con el plan publicado.
8. **Antes de construir**, proponé la estructura de hojas y esperá mi confirmación. Después de cada bloque, mostrá qué verificaste.

---

## 1. El negocio

| Concepto | Detalle |
|---|---|
| Sociedad | **ECOSTAR S.A.** (maquiladora; régimen de maquila en trámite) |
| Socios y capital | Capital a aportar **Gs 1.311.000.000**: CIE Maquiladora EAS 49% (Gs 643.000.000, integrado en bienes y efectivo) y **Energy Capital Participações Ltda 51%** (Gs 668.000.000, integración pendiente de confirmar). En el plan los socios figuran como «CIE–PERG». |
| Actividad | Procesamiento de chapas de acero: **gabinetes metálicos (153 kg)** y **cajas metálicas (13 kg)** para exportación. |
| Cliente | Grupo CIE en Brasil (las exportaciones a la fecha se facturaron de CIE S.A. a CIE do Brasil). |
| Planta | Nave alquilada de 3.865 m² en Zona Franca Global del Paraguay. |
| Cronología | Inicio de implantación 18/08/2025 · Fin de implantación 24/04/2026 · Inauguración 27/04/2026 · Inicio de producción 28/04/2026 · 1.ª exportación 26/08/2026 (392 cajas) · 2.ª exportación 07/09/2026 (323 cajas) · Autonomía operativa prevista desde **noviembre 2026**. |
| Financiamiento hasta hoy | **CIE S.A. financió el 100%** de la implantación y la producción (con fondos propios y préstamos bancarios a su nombre). |

---

## 2. Plan de negocio original (octubre 2025) — supuestos

Fuente: plan Mujica & Saldivar, oct-2025. «p.» = página impresa del plan (en el archivo PDF, sumar 2). Horizonte: **10 años de operación**; mes 1 = inicio de la operación.

### 2.1 Macroeconomía e índices por año

| Concepto | Año 1 | Año 2 | Años 3 a 10 | Fuente |
|---|---|---|---|---|
| Tipo de cambio (Gs/USD) | 7.300 | 7.300 | 7.300 | p. 4-5 |
| Variación de precios de venta | 0% | +1% | +1% anual acumulativo | p. 2 |
| Variación del costo de materia prima | 0% | +1% | +1% anual acumulativo | p. 2 |
| Variación de insumos y otros variables | 0% | 0% | 0% | — |
| Ajuste salarial | 0% | 0% | 0% | — |
| Dividendos (% del resultado del año) | 0% | 0% | Año 3: 30% · Año 4: 50% · Años 5-10: 100% | Plan |

### 2.2 Productos, precios y costos unitarios (USD por unidad)

| Producto | Precio año 1 | Materia prima | Insumos | Otros variables (consumibles + flete) | Capacidad inicial (u/año) | Capacidad ampliada desde año 3 |
|---|---|---|---|---|---|---|
| Gabinetes metálicos 153 kg | 270,00 | 82,65 | 45,00 | 27,91 | 15.000 | 22.000 |
| Cajas metálicas 13 kg | 23,00 | 7,41 | 4,851 | 2,73 | 90.000 | 121.000 |

Fuentes: precios p. 16; materia prima p. 17; insumos p. 18 (cajas: 339.570 / 70.000); otros variables p. 18-19 (cuadro 10a).

### 2.3 Volumen de ventas (unidades por año)

| Producto | Año 1 | Año 2 | Año 3 | Años 4 a 10 |
|---|---|---|---|---|
| Gabinetes | 12.000 | 15.000 | 18.000 | 21.000 |
| Cajas | 70.000 | 90.000 | 110.000 | 120.000 |

Merma entre producción y venta: 0%.

### 2.4 Parámetros de costos y gastos

| Parámetro | Valor | Base |
|---|---|---|
| Cargas sociales (aguinaldo, IPS, vacaciones) | 43,09% | Sobre salarios |
| Alimentación del personal | Gs 576.000 por persona por mes | — |
| Personal directo activo antes del inicio | 50% de la dotación del año 1 | — |
| Personal indirecto y administrativo antes del inicio | 100% | — |
| Despacho y gastos de importación | 3% | Materia prima + insumos |
| Packaging | 1% | Materiales, otros variables y personal directo |
| Imprevistos de producción | 5% | Costos de producción |
| Imprevistos de administración | 5% | Gastos de administración marcados como base |
| Tributo único de maquila | 1% | Ventas facturadas y exportadas (p. 7) |
| Seguro de caución y otros cargos | 0,1% | Ventas |
| Comisiones y logística de venta | 0% | Ventas |
| Impuesto a la renta | 0% | (régimen de maquila: tributo único) |
| IVA sobre intereses | 10% | Intereses |
| IVA recuperable de costos fijos | 0% (se cargan con IVA cuando no se recupera) | — |

### 2.5 Personal (26 puestos; salario mensual en Gs, sin cargas)

| Puesto | Clase | Salario (Gs) | Dot. año 1 | Año 2 | Año 3 | Año 4 | Años 5-10 |
|---|---|---|---|---|---|---|---|
| Operador láser | Directo | 4.668.900 | 4 | 4 | 5 | 5 | 6 |
| Doblador | Directo | 3.650.655 | 8 | 8 | 9 | 9 | 10 |
| Soldador | Directo | 3.968.900 | 5 | 5 | 6 | 6 | 6 |
| Baño | Directo | 3.650.655 | 2 | 2 | 2 | 2 | 3 |
| Pintor | Directo | 3.650.655 | 4 | 4 | 4 | 5 | 5 |
| Montador | Directo | 3.150.655 | 10 | 10 | 10 | 11 | 11 |
| Ayudante | Directo | 2.798.309 | 10 | 10 | 10 | 11 | 11 |
| Gerente General | Administración | 18.000.000 | 1 | 1 | 1 | 1 | 1 |
| Gerente Industrial | Indirecto | 9.800.000 | 1 | 1 | 1 | 1 | 1 |
| Coord. Prod. Corte y Deformación | Indirecto | 6.550.000 | 1 | 1 | 1 | 1 | 1 |
| Coord. Prod. Soldadura | Indirecto | 6.550.000 | 1 | 1 | 1 | 1 | 1 |
| Coord. Prod. Trat. Superficial y Pintura | Indirecto | 6.550.000 | 1 | 1 | 1 | 1 | 1 |
| Coord. Prod. Montaje | Indirecto | 6.550.000 | 1 | 1 | 1 | 1 | 1 |
| Ingeniería | Indirecto | 6.550.000 | 3 | 3 | 3 | 3 | 3 |
| Mantenimiento | Indirecto | 3.650.655 | 2 | 2 | 2 | 2 | 2 |
| PCP | Indirecto | 9.800.000 | 1 | 1 | 1 | 1 | 1 |
| Analista PCP | Indirecto | 2.798.309 | 1 | 1 | 1 | 1 | 1 |
| Depósito | Indirecto | 3.150.655 | 1 | 1 | 1 | 1 | 1 |
| Compras | Administración | 6.550.000 | 1 | 1 | 1 | 1 | 1 |
| Logística / Comex | Administración | 6.550.000 | 1 | 1 | 1 | 1 | 1 |
| Coord. Control de Calidad | Indirecto | 6.550.000 | 1 | 1 | 1 | 1 | 1 |
| Asistente Control de Calidad | Indirecto | 2.798.309 | 1 | 1 | 1 | 1 | 1 |
| Coord. Procesos | Indirecto | 6.550.000 | 1 | 1 | 1 | 1 | 1 |
| Asistente Procesos | Indirecto | 2.798.309 | 1 | 1 | 1 | 1 | 1 |
| Administrador | Administración | 9.800.000 | 1 | 1 | 1 | 1 | 1 |
| Asistente Administrativo | Administración | 2.798.309 | 1 | 1 | 1 | 1 | 1 |
| **Total dotación** | | | **65** | **65** | **68** | **71** | **74** |

Fuente: p. 30-34 (anexo de RR. HH.).

### 2.6 Costos fijos mensuales

| Concepto | Clase | Moneda | Monto mensual | Crecimiento anual | Nota |
|---|---|---|---|---|---|
| Energía eléctrica planta | Producción | Gs | 14.000.000 | 0% | Gs 168 M/año (inferido del plan) |
| Mantenimiento y reparaciones planta | Producción | USD | 2.475,67 | 2% | USD 29.708/año |
| Alquiler de nave (3.865 m² × USD 4) | Administración | USD | 17.006 (incluye IVA 10% no recuperable) | 3% | p. 6 y p. 19; el plan lo trata como costo de oportunidad |
| Expensas (USD 0,20/m²) | Administración | USD | 773 | 3% | p. 19 |
| Agua, comunicaciones y electricidad adm. | Administración | Gs | 750.000 | 0% | Inferido |
| Mantenimiento y reparaciones adm. | Administración | Gs | 3.000.000 | 0% | Inferido |
| Seguros sobre activos fijos | Administración | Gs | 5.000.000 | 0% | Inferido |
| Gastos generales | Administración | USD | 1.838,33 | 0% | USD 22.060/año |
| Honorarios y otros fijos | Administración | Gs | 15.166.667 | 0% | Gs 182 M/año (inferido) |

Los montos en Gs «inferidos» se derivaron de los totales anuales del cuadro 10a (p. 19); los totales coinciden con el plan.

### 2.7 Capital de trabajo

| Concepto | Días |
|---|---|
| Cobro a clientes | 120 |
| Inventario de materiales | 120 |
| Producto terminado | 10 |
| Pago a proveedores de materia prima | 0 (contado / anticipado) |
| Pago a otros proveedores | 30 |
| Caja mínima operativa | 3 días de ventas |
| Sueldos y anticipos (plan, p. 8) | 15 (no incluido en el modelo v7) |

Recuperación del capital de trabajo al final del horizonte: 100%. Meses para recuperar el IVA: 6.

### 2.8 Inversión (CAPEX, USD)

| Ítem | Clase | Monto | Vida útil |
|---|---|---|---|
| Adecuación de infraestructura | Infraestructura | 87.000 | 10 |
| Conjunto línea de pintura | Maquinaria | 183.073,60 | 10 |
| Cortadoras láser (2) | Maquinaria | 84.295,23 | 10 |
| Manipulador | Maquinaria | 10.197,00 | 10 |
| Dobladoras (4) | Maquinaria | 151.967,08 | 10 |
| Soldadoras (6) | Maquinaria | 3.399,00 | 10 |
| Lijadoras (4) | Maquinaria | 589,16 | 10 |
| Conjunto baño | Maquinaria | 28.325,01 | 10 |
| Línea de montaje | Maquinaria | 9.857,10 | 10 |
| Montacargas | Maquinaria | 29.458,01 | 10 |
| Soldadoras capacitivas (2) | Maquinaria | 6.004,90 | 10 |
| Software de gestión | Intangible | 43.478,00 | 5 |
| Mobiliario | Mobiliario | 17.400,00 | 5 |
| **Inversión fija inicial** | | **655.044** | |
| Inversión adicional año 2 (no desagregada, mes 13) | Pendiente | 100.000 | 10 |

Maquinaria = CIF + 3% de despacho (p. 12, cuadro 2). Valor de realización del activo fijo al final del horizonte: 100%.

### 2.9 Financiamiento y valoración

| Concepto | Valor |
|---|---|
| Préstamo bancario | USD 2.791.587 · 9% nominal anual + IVA sobre intereses · 6 meses de gracia + 54 cuotas (sistema francés) · desembolso mes 1 (p. 4 y p. 22) |
| Aportes de socios | USD 200.000 inicial + 567.782 (año 1) + 371.928 (año 2) = **USD 1.139.710** (p. 9 y p. 26) |
| Tasa de descuento del proyecto (WACC) | 9,49% |
| Costo del capital propio (Ke) | 16,32% |
| Mes de pago de dividendos | Abril del año siguiente |

### 2.10 Resultados publicados en el plan (USD)

| Año | Ventas | EBITDA | Resultado neto |
|---|---|---|---|
| 1 | 4.850.000 | 573.024 | 301.922 |
| 2 | 6.181.200 | 1.029.011 | 731.144 |
| 3 | 7.538.539 | 1.479.759 | 1.238.751 |
| 4 | 8.685.437 | 1.886.306 | 1.707.489 |
| 5 | 8.772.292 | 1.914.796 | 1.804.006 |
| 6 | 8.860.015 | 1.973.999 | 1.902.407 |
| 7 | 8.948.615 | 2.033.657 | 1.974.240 |
| 8 | 9.038.101 | 2.093.768 | 2.034.352 |
| 9 | 9.128.482 | 2.154.335 | 2.094.918 |
| 10 | 9.219.767 | 2.215.355 | 2.155.939 |

**VAN publicado USD 4.588.736 · TIR publicada 22,61%.**

### 2.11 Diferencias conocidas entre el plan y su recálculo (modelo v7)

- EBITDA a 10 años: 16,31 M en el recálculo contra 17,35 M en el plan (−6%), por dos correcciones de método: nómina convertida al TC de 7.300 (−682 mil) y materia prima con +1% acumulativo (−361 mil).
- El plan mezcla flujos del proyecto y del accionista en el VAN. Con el método del plan, el recálculo da VAN 4.177.576 y TIR 23,99%.
- A confirmar con los autores: TC usado en la nómina y si el 1% de materia prima es acumulativo también para las cajas.

---

## 3. Cambios de contexto respecto del plan (lo que hoy sabemos que es distinto)

1. **Financiamiento:** el plan suponía un préstamo bancario de USD 2,79 M y aportes por USD 1,14 M. En la realidad **no se tomó el préstamo ni se integraron esos aportes**: CIE S.A. financió todo, y el capital de Ecostar es de solo Gs 1.311 M (≈ USD 223 mil).
2. **Régimen tributario:** el programa de maquila **todavía no está aprobado** (resolución bicameral al 98%, cierre estimado 30/11/2026). Las exportaciones se hacen bajo **régimen general**: sin tributo de maquila, con IVA de importación pagado y recuperado por devolución al exportador, y potencialmente con IRE (10%) sobre las utilidades del período.
3. **Compras:** **todas las compras se hicieron a nombre de CIE S.A., con IVA**. Hay que trasladarlas a Ecostar por factura.
4. **Tipo de cambio:** el guaraní se apreció de 7.327 (ago-25) a ≈ 5.890 Gs/USD (oct-26). El plan usa 7.300. Como Ecostar vende en USD y paga salarios y parte de los fijos en Gs, esto **reduce el EBITDA** en USD (≈ USD 200 mil por año).
5. **Volumen real:** en ~5 meses de producción se exportaron 715 cajas, muy por debajo del ritmo del plan (70.000 cajas y 12.000 gabinetes el año 1).
6. **Inventario:** Ecostar recibe de CIE ≈ USD 1,14 M de materia prima (stock + en tránsito), equivalente a ~7 meses de consumo del plan (contra 120 días).

---

## 4. Situación real al 31/10/2026 — inversión de CIE S.A. en Ecostar

Fuente: «Informe de inversión de CIE S.A. en Ecostar RV2» (cierre de datos 31/08/2026, proyección a 31/10/2026, fecha de informe 29/09/2026) e «Informe ejecutivo RV3». TC de conversión: 5.890 Gs/USD (estimado BCP oct-26).

### 4.1 Total incurrido por CIE

| Concepto | Guaraníes | USD |
|---|---|---|
| (+) Etapa de implantación | 6.502.812.974 | 1.104.043 |
| (+) Etapa de producción, compromisos y estimaciones | 9.392.590.459 | 1.594.667 |
| (+) Gastos de funcionamiento (overhead 7%) | 1.162.467.279 | 197.363 |
| (+) Intereses de implantación (fondos propios CIE al 9%) | 213.684.783 | 36.279 |
| (+) Intereses de préstamos bancarios de CIE (al 9% del plan) | 497.587.200 | 84.480 |
| **(=) Total incurrido** | **17.769.142.695** | **3.016.832** |
| (−) Exportación a cobrar (factura 001-002-0000005) | (52.558.728) | (8.923) |
| (−) Aporte de capital de CIE en bienes y efectivo | (643.000.000) | (109.168) |
| **(=) Cuenta a cobrar a Ecostar (monto a facturar)** | **17.073.583.967** | **2.898.741** |

Valorizado al TC de cada mes del flujo, el recupero sería de USD 2.817.721 (diferencia de USD 81.020 por tipo de cambio).

### 4.2 Composición relevante

- **Implantación:** adecuaciones de infraestructura Gs 1.625 M (incluye alquiler de galpón Gs 728 M, sueldos de operadores Gs 280 M, insumos de puesta a punto Gs 237 M); maquinarias y equipos importados Gs 4.165 M (equipos de producción Gs 3.752 M); muebles y equipos de oficina Gs 147 M; IVA por compras Gs 566 M.
- **Producción (abr-oct 2026):** materia prima Gs 6.615 M, de la cual **stock Gs 2.071 M y en tránsito Gs 4.619 M** (≈ USD 1,14 M de inventario); insumos Gs 376 M; mano de obra de fábrica Gs 339 M; gastos de administración Gs 1.109 M (alquiler Gs 572 M, mano de obra administrativa Gs 399 M); IVA por compras Gs 787 M. Incluye proyección sep-oct de USD 938.894 (mercadería en tránsito USD 862.642).
- **Activo fijo trasladable** (partidas capitalizables): ≈ Gs 4.500 M (≈ USD 764 mil), incluido el aporte en bienes.
- **Aporte de CIE en bienes:** 13 activos por Gs 642.000.000 (dos dobladoras CNHAWE HPC-135/3100 de Gs 293 M cada una, remachadoras, aparejo, taladros, PCs, etc.) + Gs 1.000.000 en efectivo.
- **Préstamos tomados por CIE:** USD 1.810.000 (jun-2026, 120 días, vto. sep-2026) renovado por **USD 3.000.000 (60 días, vto. noviembre 2026)**. Tasa real de CIE 5%-6%; el informe carga 9% (tasa del plan).
- **Avance de autonomía a oct-26:** 86% (registro importador 50%, gestiones bancarias 51%, registro exportador 88%, patente 90%, documentos societarios 75%).

### 4.3 Observaciones abiertas sobre el monto a facturar (revisión de Planificación Financiera)

| # | Partida | Impacto aprox. | Estado |
|---|---|---|---|
| 1 | IVA incluido en el monto (compras + intereses), posiblemente ya usado por CIE como crédito fiscal | USD 240.600 (Gs 1.417 M) | Confirmar con Impuestos |
| 2 | Intereses «bancarios» al 9% del plan contra 5%-6% reales | USD 35.500 | Requiere acuerdo del socio |
| 3 | Overhead 7% calculado también sobre intereses e IVA | USD 24.500 | Revisar base |
| 4 | Diferencia de cambio trasladada a Ecostar al facturar en Gs | USD 81.000 | Decisión de socios |
| 5 | Segunda exportación (323 cajas, 07/09) no descontada | ≈ USD 7.400 | Verificar |
| 6 | Depreciación incluida sobre activos ya cargados; Gs 1 M de efectivo descontado como bien | < USD 3.000 | Menor |

**Estado de aprobación:** el detalle de gastos ya se presentó a Energy Capital; **todavía no está aprobado**.

### 4.4 Posición estimada de apertura de Ecostar (31/10/2026)

| Concepto | USD aprox. |
|---|---|
| Activo fijo (incluye aporte en bienes) | 764.000 |
| Inventario de materia prima (sin IVA) | 1.136.000 |
| Crédito fiscal IVA (si Ecostar puede usarlo) | 241.000 |
| Caja (aporte de Energy Capital, si se integra) | 113.000 |
| **Total activo** | **2.254.000** |
| Deuda con CIE | 2.899.000 |
| **Patrimonio neto** | **≈ (645.000)** |

El patrimonio negativo resulta de que los gastos preoperativos y de arranque (≈ USD 0,87 M) se van a pérdida y el capital integrado es bajo. Es un obstáculo para el crédito bancario.

---

## 5. Objetivo actual: primer crédito bancario de Ecostar (noviembre 2026)

El préstamo de USD 3,0 M está a nombre de CIE y vence en noviembre. El objetivo es que **Ecostar obtenga su propio crédito** para cancelar la deuda con CIE, y el banco pide un **flujo de caja a 12 y 24 meses**.

### 5.1 Supuestos del flujo bancario (nov-2026 a oct-2028)

| Supuesto | Valor | Fuente / criterio |
|---|---|---|
| Mes 1 | Noviembre 2026 | Inicio de la operación autónoma |
| TC año 1 y año 2 | 5.890 Gs/USD | BCP estimado oct-26; sin devaluación (prudente) |
| Volúmenes | Plan año 1 y año 2 | Con ramp-up mensual: 25%, 35%, 45%, 55%, 65%, 75%, 85%, 95% y 100% desde el mes 9 (supuesto propio, a validar con pedidos) |
| Personal | Dotación del plan; directos con piso del 50% durante el ramp-up | Plan |
| Inventario inicial de materia prima | Gs 6.690 M (≈ USD 1,14 M) | Informe CIE; se consume antes de recomprar hasta 120 días de objetivo |
| Régimen | General hasta el mes 3 (enero 2027); maquila desde entonces | IVA de importación 10% recuperado a 6 meses; tributo de maquila 1% desde la aprobación; arancel 0% (a confirmar) |
| Deuda con CIE | Base sin IVA Gs 15.656 M + IVA 10%, facturada en Gs (≈ USD 2,92 M) | Ecostar recupera el IVA como exportador a 6 meses |
| Préstamo bancario | USD 2,9 M · 9% · 6 meses de gracia + 54 cuotas | Condiciones del plan; reemplazar por la oferta del banco |
| Saldo con CIE no cubierto | 9% · 6 meses de gracia · 12 cuotas | Supuesto |
| Aporte de Energy Capital | Gs 668 M en el mes 1 | Capital social |
| CAPEX año 2 | USD 100.000 en el mes 13 | Plan |
| Caja mínima | USD 50.000 | Supuesto propio |

### 5.2 Resultados del flujo base (USD)

| Indicador | Año 1 | Año 2 |
|---|---|---|
| Ventas | 3,56 M | 6,18 M |
| EBITDA | (0,01 M) | 0,76 M (12%) |
| Flujo operativo | (0,73 M) | 0,18 M |
| DSCR | negativo | 0,10x |
| Financiamiento adicional requerido | | 1,90 M |

**Conclusión:** con los supuestos del plan, Ecostar no puede atender un préstamo de USD 2,9 M. Las causas son el cobro a 120 días, el tipo de cambio y la falta de capital.

### 5.3 Escenarios evaluados

| Escenario | DSCR año 2 | Financiamiento adicional |
|---|---|---|
| A. Base | 0,10x | 1,90 M |
| B. Cobro a 30 días | 0,50x | 0,35 M |
| C. B + 12 meses de gracia + 84 cuotas | 0,72x | 0,17 M |
| D. C + capitalización de Gs 6.000 M de la deuda con CIE (préstamo de USD 1,8 M) | **1,18x** | 0,14 M |
| E. D con cobro a 60 días | 0,86x | 0,50 M |

Estructura recomendada para el banco: la D (cobro a 30 días acordado con CIE do Brasil, préstamo de largo plazo con 12 meses de gracia, capitalización parcial de la deuda con CIE con aporte proporcional de ambos socios) más una línea de capital de trabajo o prefinanciación de exportaciones.

---

## 6. Decisiones y datos pendientes

1. **Aprobación de Energy Capital** del monto a recuperar (overhead 7%, tasa del 9%, facturación en Gs).
2. **Forma de recupero de CIE:** préstamo, capitalización o mixto, y tasa.
3. **Impuestos:** IVA de la factura de CIE y riesgo de doble recupero; plazo real de devolución del IVA al exportador; IRE durante el régimen general; base del tributo de maquila (Ley 7547/2025 y Decreto 5714/2026).
4. **Fecha de aprobación del programa de maquila** y tratamiento de la mercadería en tránsito importada a nombre de CIE (admisión temporaria).
5. **Volumen 2027-2028** respaldado con pedidos u órdenes de compra de CIE do Brasil.
6. **Plazo de cobro** a CIE do Brasil (el plan supone 120 días).
7. **Oferta del banco**: monto, tasa, gracia, plazo, garantías y covenants.
8. **Dotación y salarios reales** vigentes desde noviembre 2026.
9. **Integración del aporte** de Energy Capital (Gs 668 M) y fecha.
10. **Detalle mensual real** ago-2025 a oct-2026 (disponible en el informe de CIE, hoja «FLUJO»).

---

## 7. Estructura sugerida del libro

1. **Inicio**: datos del negocio, fecha de corte, escenario activo, estado de controles.
2. **Supuestos**: secciones 2 y 5 de este documento, con columnas «Plan oct-2025», «Vigente» y «Fuente».
3. **Productos y volumen**: precios, costos unitarios, capacidad, volumen anual y ramp-up mensual.
4. **Personal** y **Costos fijos**.
5. **Inversión y financiamiento**: CAPEX, préstamo bancario, deuda con CIE, aportes.
6. **Apertura**: balance de Ecostar al 31/10/2026 (sección 4.4).
7. **Flujo mensual**: 24 meses (resultado devengado, capital de trabajo, flujo de caja por método directo, deuda, DSCR).
8. **Proyección anual**: 10 años (estado de resultados, flujo, balance, VAN y TIR del proyecto y del accionista).
9. **Escenarios y sensibilidad**: volumen, TC, días de cobro, tasa y plazo del préstamo.
10. **Comparación con el plan**: plan publicado frente a vigente, con la diferencia explicada.
11. **Controles**.
