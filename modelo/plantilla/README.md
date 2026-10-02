# Plantilla común de maquilas (v2)

Este generador produce el libro de cada negocio con la misma estructura de 9 hojas: Inicio, Supuestos, Productos, Costos, Inversion, Reales, Calculo, Resultados y Control. El documento `../../Informe_Plantilla_Comun_v2.md` explica el diseño, los criterios y la carga.

## Generar

```bash
python3 generar_plantilla.py ecostar ../../Modelo_Maquila_ECOSTAR_v2.xlsx   # usa datos_ecostar.py
python3 generar_plantilla.py vacia   ../../Plantilla_Maquila_Comun.xlsx     # plantilla en blanco
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
- **Control de integridad (fila `CHK`):** activo − pasivo − patrimonio = 0 en todos los meses.
  - Activo: caja, clientes, inventarios, IVA y activo fijo.
  - Pasivo: proveedores y deuda.
  - Patrimonio: aportes, resultados y dividendos, más el ajuste de conciliación con la caja real.

## Scripts de prueba

```bash
python3 probar_plantilla.py   libro.xlsx recalc.py carpeta   # reales congelados, TC, días, atraso, implantación
python3 probar_escenarios.py  libro.xlsx recalc.py carpeta   # 36 escenarios (genera escenarios.pkl)
python3 informe_escenarios.py carpeta/escenarios.pkl ../../Pruebas_Escenarios_ECOSTAR_v2.xlsx   # resumen y 385 verificaciones
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
