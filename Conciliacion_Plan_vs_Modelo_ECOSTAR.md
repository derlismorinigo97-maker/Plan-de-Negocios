# Conciliación: modelo Ecostar v5 frente al plan de negocio CIE–PERG (octubre 2025)

Fecha: 27/10/2026.

- **Plan:** «Procesamiento de chapas para gabinetes metálicos – Empresas CIE–PERG», Mujica & Saldivar, octubre 2025 (37 páginas).
- **Modelo:** `Modelo_Maquila_ECOSTAR_v5.xlsx`, bloque de presupuesto original. Unidad: USD.

## Conclusión

1. **Los datos de entrada coinciden con el plan.** Verifiqué en el PDF cada dato cargado y la página citada. Las referencias del libro usan la **numeración impresa** del documento (la del índice). Para ubicarla en el archivo PDF hay que sumar 2, por la portada y el índice: la p. 19 del plan es la página 21 del archivo.
2. **Coinciden con el plan** (diferencias menores a USD 10, por redondeo):
   - ventas, insumos, consumibles y fletes;
   - seguro de caución, tributo de maquila;
   - gastos de administración, intereses;
   - inversión, préstamo y aportes.
3. **El EBITDA a 10 años es USD 1,04 M menor que en el plan** (16,31 M frente a 17,35 M; −6,0%). La diferencia está explicada en su totalidad por dos correcciones de método:
   - nómina convertida al TC declarado de 7.300 (−682 mil);
   - materia prima con el 1% acumulativo que indica el texto del plan (−361 mil).
4. **El resultado neto es USD 1,19 M menor** (−7,5%): la diferencia del EBITDA, más 149 mil de depreciación adicional (desde el año 1 y sobre el CAPEX del año 2).
5. **El capital de trabajo del modelo es mayor**: 52 mil más en el año 1 y 519 mil más en el año 10. El modelo valoriza las cuentas por cobrar a precio de venta; el plan, a costo. Esto afecta la necesidad de fondos, no el EBITDA.
6. **VAN y TIR no son comparables directamente.** El plan mezcla flujos del proyecto y del accionista. La versión anterior del modelo (v1) reproduce exactamente el VAN publicado (4.588.736), lo que confirma que la base de datos es la misma.
7. **Con el método del plan, el modelo da VAN 4.177.576 y TIR 23,99%** (actualización v7, cuadro 19): 411 mil menos que el VAN publicado. El mismo cálculo aplicado a los flujos publicados reproduce el VAN y la TIR del plan (4.588.737 y 22,61%), así que la diferencia es de cifras (EBITDA, capital de trabajo y calendario), no de método. La TIR sube aunque el VAN baja, principalmente porque en el modelo la inversión cae en el año 1 y no en el año 0 (el horizonte entre la inversión y el valor final es un año más corto).

**Desde la v7, el libro muestra esta conciliación cuadro por cuadro:** hoja «Cuadros» (resumen y causas) y una hoja por cuadro del plan (C01 a C20 y RRHH), con el modelo, el plan publicado y la diferencia explicada en cada línea.

**A confirmar con los autores del plan:** el TC usado en la nómina y si el 1% de materia prima es acumulativo también para las cajas. Son las dos diferencias materiales.

## 1. Trazabilidad de los datos de entrada

«p.» = página impresa del plan; en el archivo PDF es la misma página + 2.

| Dato | Valor en el modelo | Dónde en el libro | Plan | Verificado |
|---|---|---|---|---|
| TC de referencia | Gs 7.300 | Supuestos fila 39 | p. 4-5 | ✔ |
| Precios año 1 y crecimiento | 270 y 23 USD/u; +1% anual desde el año 2 | Productos G6:G7, Supuestos fila 40 | p. 16 y p. 2 | ✔ |
| Capacidad instalada | 15.000→22.000 gabinetes; 90.000→121.000 cajas (año 3) | Productos D6:F7 | p. 2 | ✔ |
| Volumen de ventas | 12.000 / 15.000 / 18.000 / 21.000 gabinetes; 70.000 / 90.000 / 110.000 / 120.000 cajas | Productos D34:M35 | p. 2 y p. 16 | ✔ |
| Materia prima por unidad | 82,65 y 7,41 USD | Productos I6:I7 | p. 17 | ✔ |
| Incremento de materia prima | 1% acumulativo desde el año 2 | Supuestos fila 41 | p. 2 | ✔ (ver diferencia C2) |
| Insumos por unidad | 45,00 gabinetes; 4,851 cajas (339.570 / 70.000) | Productos K6:K7 | p. 18 | ✔ |
| Consumibles y flete por unidad | 27,91 y 2,73 USD | Productos M6:M7 | p. 18-19 (cuadro 10a) | ✔ inferido; total igual al plan |
| Nómina: 26 puestos, salario en Gs, dotación años 1-5 | Gerente General Gs 18.000.000, etc. | Costos filas 6-31 | p. 30-34 | ✔ |
| Cargas sociales | 43,09% | Supuestos G04 | p. 19 | ✔ |
| Alimentación | Gs 576.000 por persona por mes | Supuestos G05 | p. 19 | ✔ inferido; total igual al plan |
| Despacho / packaging / imprevistos | 3% / 1% / 5% | Supuestos G08-G11 | p. 19 | ✔ |
| Tributo de maquila | 1% de lo facturado y exportado | Supuestos G12 | p. 7 | ✔ |
| Seguro de caución | 0,1% de ventas | Supuestos G13 | p. 19 | ✔ inferido; total igual al plan |
| Energía, mantenimiento, agua, seguros, gastos generales, honorarios | Montos mensuales (Gs o USD) | Costos filas 75-83 | p. 19 (cuadro 10a) | ✔ montos en Gs inferidos; totales iguales |
| Alquiler | 3.865 m² × USD 4 + IVA (+3% anual) | Costos fila 77 | p. 6 y p. 19 | ✔ |
| Expensas | USD 0,20 por m² (+3% anual) | Costos fila 78 | p. 19 | ✔ |
| Días de capital de trabajo | Clientes 120, stock 120, producto terminado 10, proveedores menores 30 | Supuestos G16-G20 | p. 8 | ✔ |
| Proveedores de materia prima | 0 días | Supuestos G19 | p. 24 (cuadro 16) | ✔ |
| Caja mínima | 3 días de ventas | Supuestos G21 | p. 24 (cuadro 16) | ✔ |
| Sueldos y anticipos | 15 días | — | p. 8 y p. 24 | **No incluido** (ver sección 4) |
| CAPEX: infraestructura, software, mobiliario | 87.000 / 43.478 / 17.400 | Inversion filas 6, 17, 18 | p. 13 (cuadros 5 y 6) | ✔ |
| CAPEX: maquinaria | CIF + 3% de despacho | Inversion filas 7-16 | p. 12 (cuadro 2) | ✔ inversión fija 655.044 = p. 11 |
| CAPEX adicional del año 2 | 100.000 | Inversion fila 19 | p. 26 (cuadro 18) | ✔ |
| Préstamo | 2.791.587 al 9% + IVA; 6 meses de gracia + 54 cuotas | Inversion E32:E36 | p. 4 y p. 22 (cuadro 13) | ✔ |
| Aportes | 200.000 inicial; 567.782 y 371.928 adicionales | Inversion filas 50-52 | p. 9 y p. 26 (cuadro 18) | ✔ |
| WACC / Ke | 9,49% / 16,32% | Supuestos G25-G26 | p. 9 | ✔ |
| Política de dividendos | 0 / 0 / 30% / 50% / 100% | Supuestos fila 44 | p. 23 (cuadro 15) | ✔ |
| Plan publicado: ventas, EBITDA, utilidad | Referencia en Supuestos D y Resultados E | Supuestos filas 76-78 | p. 23 (cuadro 15) | ✔ |
| Plan publicado: VAN y TIR | 4.588.736 / 22,61% | Supuestos | p. 27 (cuadro 19) | ✔ |

## 2. Conciliación de resultados, línea por línea (presupuesto del modelo frente al plan)

| Línea (cuadro 10a, p. 19) | Plan 10 años | Modelo 10 años | Diferencia | Año 1 | Año 10 | Explicación |
|---|---|---|---|---|---|---|
| Ventas | 81.222.448 | 81.222.448 | 0 | 0 | 0 | Igual |
| Materia prima | 24.963.261 | 25.293.483 | +330.222 | 0 | +74.413 | C2: 1% acumulativo también en cajas |
| Insumos | 14.024.610 | 14.024.610 | 0 | 0 | 0 | Igual |
| Consumibles + fletes | 8.388.884 | 8.388.892 | +8 | 0 | +1 | Redondeo |
| Personal (fábrica + administración, con cargas y alimentación) | 7.413.350 | 8.059.647 | +646.297 | +59.744 | +66.714 | C1: nómina a TC 7.300 |
| Energía y mantenimiento de planta | 553.960 | 555.432 | +1.472 | 0 | +448 | El plan crece 2% lineal; el modelo, 2% compuesto |
| Despachos 3% | 1.169.636 | 1.179.543 | +9.907 | 0 | +2.232 | Derivado de C2 |
| Seguro de caución | 81.222 | 81.222 | 0 | 0 | 0 | Igual |
| Packaging 1% | 515.535 | 522.395 | +6.860 | +306 | +1.121 | Derivado de C1 y C2 |
| Imprevistos 5% (producción + administración) | 2.886.212 | 2.935.949 | +49.737 | +3.002 | +7.247 | Derivado de C1 y C2 |
| Tributo de maquila 1% | 812.224 | 812.224 | 0 | 0 | 0 | Igual |
| Gastos fijos de administración | 3.059.555 | 3.059.546 | −9 | −1 | −1 | Redondeo |
| **EBITDA** | **17.354.010** | **16.309.504** | **−1.044.506** | **−63.054** | **−152.175** | Ver puente (sección 3) |
| Depreciación | 595.628 | 745.044 | +149.416 | +71.592 | +10.000 | C3 y C4 |
| Gastos financieros (con IVA) | 813.214 | 813.214 | 0 | 0 | 0 | Igual |
| **Resultado neto** | **15.945.168** | **14.751.246** | **−1.193.922** | **−134.646** | **−162.176** | EBITDA + depreciación |
| Inversión (CAPEX) | 755.044 | 755.044 | 0 | — | — | Igual |
| Préstamo y aportes | 2.791.587 / 1.139.710 | Igual | 0 | — | — | Igual |

Diferencia de EBITDA por año (modelo − plan), años 1 a 10:

| Año 1 | Año 2 | Año 3 | Año 4 | Año 5 | Año 6 | Año 7 | Año 8 | Año 9 | Año 10 |
|---|---|---|---|---|---|---|---|---|---|
| −63.054 | −63.054 | −74.739 | −87.583 | −100.238 | −110.394 | −120.665 | −131.050 | −141.554 | −152.175 |

## 3. Puentes del plan al modelo

**EBITDA a 10 años:** 17.354.010 (plan) → 16.309.504 (modelo)

| Concepto | Efecto | Detalle |
|---|---|---|
| C1. Nómina al TC declarado de 7.300 | −682.348 | El plan convierte los salarios en Gs a un TC implícito de 8.000 (Gerente General: Gs 18.000.000 = USD 2.250, p. 30-34), aunque declara 7.300 (p. 4-5). Son unos 59.700 por año, más su efecto en packaging e imprevistos. |
| C2. Materia prima con 1% acumulativo | −360.603 | El texto (p. 2) dice «incremento de 1% acumulativo». El cálculo del plan no acumula el 1% en las cajas. Va de 0 en los años 1-2 a 74.413 en el año 10, más despachos, packaging e imprevistos. |
| Mantenimiento de planta 2% compuesto | −1.545 | El plan suma un 2% fijo sobre la base (lineal). |
| Redondeos | −10 | |
| **Total** | **−1.044.506** | Explicado al 100% |

**Resultado neto:** −1.193.922 = −1.044.506 (EBITDA) − 149.416 (depreciación).
- **C3. Depreciación desde la puesta en servicio (año 1).** El plan empieza en el año 2. Efecto neto +59.417, por un año más de depreciación de los activos a 10 años dentro del horizonte.
- **C4. CAPEX adicional del año 2 depreciado.** El plan no lo deprecia. Efecto +10.000 por año, +90.000 en total.

## 4. Capital de trabajo (cuadro 16, p. 24)

| Componente al cierre | Plan año 1 | Modelo año 1 | Plan año 10 | Modelo año 10 |
|---|---|---|---|---|
| Clientes (120 días) | 1.406.129 (a costo) | 1.594.521 (a precio de venta) | 2.302.820 | 3.031.156 |
| Stock de materiales y producto terminado | 936.492 | 892.649 | 1.671.473 | 1.626.938 |
| Caja y bancos (3 días) | 39.863 | — (se trata como caja mínima) | 75.779 | — |
| Sueldos y anticipos (15 días) | 17.890 | — (no incluido) | 19.977 | — |
| Proveedores | −63.831 | −99.105 | −89.232 | −158.332 |
| **Capital de trabajo neto** | **2.336.543** | **2.388.064** | **3.980.818** | **4.499.763** |

- **Clientes.** El plan calcula las «ventas a crédito» sobre los costos (producción + administración). El modelo las calcula sobre las ventas, que es lo correcto: lo que se cobra es lo facturado. Por eso el capital de trabajo del modelo es mayor y crece con el margen (+728 mil en el año 10).
- **Caja mínima.** El plan la incluye dentro del capital de trabajo. El modelo la usa como piso de caja; el efecto en la necesidad de fondos es equivalente.
- **Sueldos y anticipos (15 días).** Están en el plan y no en el modelo. Son unos 18.000 a 20.000 USD. Se pueden agregar al generador si se confirma que los anticipos de sueldos son una práctica real.

## 5. VAN y TIR (cuadro 19, p. 27)

| | Plan | Modelo |
|---|---|---|
| Flujo | «Beneficios netos» anuales: inversión total al año 0, financiamiento mezclado | FCFF mensual (sin financiamiento) para el proyecto; aportes y dividendos para los socios |
| Tasa | 9,49% (ponderada) | 9,49% para el proyecto; 16,32% para los socios |
| Resultado | VAN 4.588.736; TIR 22,61% | VAN del proyecto 6.989.922 (TIR 34,9%); VAN de los socios 3.695.678 (TIR 43,6%) |

No son comparables por método. La v1 (`versiones/v1_…detallado.xlsx`, hoja Base_PDF) reproduce el VAN publicado al centavo (4.588.736,60) y todos los cuadros con diferencias menores a USD 10. Esa reproducción respalda que los datos del modelo son los del plan.

## 6. Qué conviene decidir

1. **TC de la nómina (C1, −682 mil en 10 años).** Confirmar con los autores si la nómina en Gs se convirtió a 8.000. Con salarios reales en guaraníes y el TC actual (el archivo de Alianza usa 6.145), el costo en USD sería todavía mayor que en el modelo.
2. **1% acumulativo en cajas (C2, −361 mil).** El texto lo dice y el cálculo no lo aplica. Hay que definir cuál vale; el modelo sigue el texto.
3. **Depreciación (C3, C4).** Criterio contable estándar: se mantiene.
4. **Capital de trabajo.**
   - Mantener clientes a precio de venta.
   - Decidir si se agregan los anticipos de sueldos (15 días).
