# Modelo financiero de maquila CIE–PERG

Libro principal: `../Modelo_Financiero_Maquila_CIE_PERG.xlsx` (editable, con fórmulas vinculadas, sin macros ni vínculos externos).

Base documental: plan «Procesamiento de chapas para gabinetes metálicos – Empresas CIE-PERG» (Mujica & Saldivar, actualizado a octubre de 2025) e informe de validación del 29/09/2026. No se suministraron datos reales posteriores al plan: el libro **no** representa la situación de 2026.

## Estructura del libro

| Hoja | Función |
|---|---|
| Guia | Uso, convenciones, métricas y registro de avance |
| Resumen | PDF publicado vs escenario activo, alertas y gráficos |
| Supuestos | Selector único de escenario, prueba de sensibilidad, parámetros, correcciones C1-C6, drivers por escenario |
| Calendario | Meses del modelo, años, banderas de operación/horizonte/reales, TC y descuento |
| Produccion_Ventas / Costos_Personal / Inversion_Activos / Capital_Trabajo / Deuda_Tributos / Resultados_Caja | Motor mensual único (157 meses) |
| Resultados_Anuales | Agregación anual, comparación con el PDF e indicadores |
| Escenarios_Sensib | Capturas identificadas de escenarios, 15 pruebas y puente de correcciones, con verificación en vivo |
| Datos_Reales | Tablas vacías para fechas, saldos, CAPEX, movimientos mensuales y documentación |
| Base_PDF | Datos publicados y reconstrucción de los cuadros 1-20 (referencia histórica inalterada) |
| Validaciones | 34 controles con tolerancia, 27 inconsistencias del PDF y 25 pruebas del motor |

## Regenerar

```bash
python3 generar_modelo.py libro_base.xlsx                      # libro sin capturas
python3 capturar_escenarios.py libro_base.xlsx capturas.json <ruta>/recalc.py
CAPTURAS=capturas.json python3 generar_modelo.py Modelo_Financiero_Maquila_CIE_PERG.xlsx
```

Luego recalcular (LibreOffice headless o abrir en Excel). Requiere `openpyxl`; las capturas requieren LibreOffice Calc.

## Resultados de la base depurada (supuestos Oct-2025, sin datos reales)

- Reproducción del PDF: diferencias ≤ USD 9,3 (redondeo). VAN publicado reproducido: USD 4.588.736,60; TIR 22,61%.
- VAN del proyecto (FCFF, tasa histórica 9,49%): USD 6,73 M; TIR 32,3%. VAN de socios (Ke 16,32%): USD 3,75 M; TIR 45,0%.
- Pico de necesidad de fondos antes de aportes adicionales: USD 1,53 M (el PDF planificó USD 939.710). Brecha no cubierta: USD 589.308 (mes 28).
- Estas cifras usan supuestos históricos y no acreditan la viabilidad actual.
