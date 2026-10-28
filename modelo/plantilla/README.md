# Plantilla común de maquilas (v2)

Este generador produce el libro de cada negocio con la misma estructura de 28 hojas:

- **Modelo (12):** Inicio, Guia_Carga (guía de carga con estado automático), Supuestos, Productos, Costos, Inversion, Detalle_24m (proyección mes a mes de la ventana móvil), Reales, Calculo, Proyeccion_24m (vista de los próximos 24 meses), Resultados y Control.
- **Cuadros con la estructura del plan de negocio (16):** «Cuadros» (índice, versión que muestran y resumen modelo vs plan) y una hoja por cuadro: C01 Inversión, C07 Producción, C08 Ingresos, C09 Materiales, C10a Costos, C10b Costo unitario, C11 Rubros, C13 Deuda e IVA, C15 Resultados, C16 Capital trabajo, C17 Equilibrio, C18 Fondos, C19 VAN-TIR, C20 Depreciación y RRHH. Cada hoja tiene tres bloques con las mismas filas: A = modelo (presupuesto original o vigente), B = plan publicado (entrada; Ecostar precargado desde `datos_ecostar.py`) y C = diferencia con su explicación. Columnas: año 0 (implantación) y años de operación 1 a 10.

 El documento `../../Informe_Plantilla_Comun_v2.md` explica el diseño, los criterios y la carga.

## Generar

```bash
python3 generar_plantilla.py ecostar ../../Modelo_Maquila_ECOSTAR_v2.xlsx   # usa datos_ecostar.py
python3 generar_plantilla.py vacia   ../../Plantilla_Maquila_Comun.xlsx     # plantilla en blanco
python3 armar_ecostar.py ../../Modelo_Maquila_ECOSTAR_v7.xlsx libro_con_registro.xlsx [fila_log.json]   # libro de trabajo con «Claude Log» y «Esp_25Productos»
```

Después hay que recalcular, abriendo el libro en Excel o con LibreOffice headless. El libro no tiene macros ni vínculos externos.

Para generar el libro de otro negocio desde código, hay que crear `datos_<negocio>.py` con la misma estructura que `datos_ecostar.py` y cambiar la importación en `generar_plantilla.py`. Como alternativa, se puede copiar la plantilla vacía y cargarla a mano.

## Capacidad

| Elemento | Capacidad |
|---|---|
| Meses | 144 (12 años) |
| Productos | 25 (individuales) |
| Puestos | 30 |
| Costos fijos | 25 |
| Ítems de CAPEX | 20 |
| Préstamos | 3, más una línea rotativa opcional |
| Aportes | 10 |
| Supuestos anuales | 10 años |

## Convenciones del motor (hoja Calculo)

- Cada fila tiene un código en la columna A (`P.vtas`, `V.CAJA`…) y su total en la columna E. La columna F es el mes 1.
- **Bloque P:** presupuesto original. Usa solo los valores `o_*` y `*_O`.
- **Bloque V:** vigente. En los meses `<= Corte` toma Reales. Después del corte usa los supuestos vigentes (`vg_*`, `*_V`) con los ajustes del escenario (`e_*`).
- Los nombres de rango se cuidan de no repetirse: Excel no distingue mayúsculas.
  - `o_X`: presupuesto;
  - `vg_X`: vigente cargado;
  - `P_X` / `V_X`: valor aplicado por versión.
- **Contadores de Detalle_24m (filas `det.*`, al final de Calculo):** celdas con dato por mes y bloque; alimentan la guía y Proyeccion_24m. Están en el motor para que limpiar el detalle no los borre.
- **Cuadros:** suman las filas del motor por año de operación (`P.aop` / `V.aop`; año 0 = `pre`) y toman los saldos del último mes de cada año. El selector `cu_Ver` (hoja Cuadros) elige el bloque P o V.
- **Control de integridad (fila `CHK`):** activo − pasivo − patrimonio = 0 en todos los meses.
  - Activo: caja, clientes, inventarios, IVA y activo fijo.
  - Pasivo: proveedores y deuda.
  - Patrimonio: aportes, resultados y dividendos, más el ajuste de conciliación con la caja real.

## Scripts de prueba

```bash
python3 probar_plantilla.py   libro.xlsx recalc.py carpeta   # reales congelados, TC, días, atraso, implantación
python3 probar_escenarios.py  libro.xlsx recalc.py carpeta   # 39 escenarios (genera escenarios.pkl)
python3 informe_escenarios.py carpeta/escenarios.pkl ../../Pruebas_Escenarios_ECOSTAR_v6.xlsx [referencia.pkl]   # resumen, verificaciones y regresión
python3 probar_detalle.py     libro.xlsx recalc.py carpeta   # proyección rodante de 24 meses (60 comprobaciones)
python3 comparar_excel.py     referencia.xlsx libro_excel.xlsx [escenarios.pkl]   # libro devuelto por Excel frente a LibreOffice
python3 probar_cuadros.py     libro.xlsx recalc.py carpeta   # cuadros con la estructura del plan (7 casos)
```

`libro.xlsx` es el libro generado sin recalcular. `recalc.py` es el script de recálculo con LibreOffice headless.

## Pruebas realizadas (LibreOffice 24.2, 01/10/2026)

| Caso | Resultado |
|---|---|
| Ecostar base y plantilla vacía | 0 errores, balance = 0. |
| 6 meses reales + cambio Vigente → Conservador | Los meses 1-6 no cambian: ventas, EBITDA, depreciación, resultado, CAPEX, saldos. El bloque P no cambia. |
| TC +10% vigente | Materiales y ventas en USD sin cambio. Nómina en PYG ×1/1,1. Costos fijos solo en su parte en PYG. |
| Días de cobro 150 | EBITDA igual, capital de trabajo mayor, VAN menor. |
| Conservador | Inicio atrasado 3 meses. CAPEX pendiente +10% (755.044 → 830.548). |
| Implantación (inicio en el mes 7) | Personal parcial y costos «desde año 0» en los meses 1-6, ventas desde el mes 7, depreciación desde la puesta en servicio. |
| Batería de 36 escenarios (12 productos) | 385/385 verificaciones OK. Detalle en `../../Pruebas_Escenarios_v2.md`. |
| Batería v4 (25 productos): 39 escenarios + regresión | 489/489 verificaciones OK, mismos resultados que con 12 productos. Detalle en `../../Pruebas_Escenarios_v4.md`. |
| v7 (28/10/2026): cuadros con la estructura del plan | 147/147 OK en 7 casos: fuentes = usos y caja, totales = Resultados, VAN anual = VAN mensual, saldos al cierre, ajustes en 0, selector de versión, Conservador, reales, detalle mensual e implantación. El método del plan aplicado a los flujos publicados da VAN 4.588.737 y TIR 22,61% (los publicados). |
| v7: regresión frente a la v6 | Motor, Resultados, Proyeccion_24m, guía y Control idénticos (50.484 celdas). 39 escenarios 495/495 e idénticos a la v6 (1.625.364 valores); 7 de operación y 60 de detalle sin cambios. Plantilla vacía sin errores. Detalle en `../../Pruebas_Escenarios_v7.md`. |
| v6 en Excel 16 (Windows, es-PY) con Claude en Excel | 439/439 OK. Verificación cruzada: 342 indicadores de 38 escenarios y 53.197 celdas del caso base iguales a LibreOffice; fórmulas del libro devuelto idénticas a la v6. |
| v6 completa (27/10/2026): 39 escenarios + 7 de operación + 60 de detalle | 495/495 y 60/60 OK; regresión frente a la v4 sin diferencias. Detalle en `../../Pruebas_Escenarios_v6.md`. |
| v6: proyección rodante de 24 meses (27/10/2026) | Con Detalle_24m vacío, Calculo idéntico a la v5 (0 celdas distintas). 25 pruebas de detalle OK: volumen, precio, materiales, personal, fijos, TC, ventana móvil con cierre, escenario y ventana de 12 meses; balance 0. |
