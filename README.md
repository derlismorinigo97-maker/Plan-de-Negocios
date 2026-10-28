# Plan de negocios — modelos financieros de maquilas

| Archivo | Uso |
|---|---|
| `Modelo_Maquila_ECOSTAR_v7.xlsx` | **Versión de trabajo actual**: la v6 más los cuadros con la estructura del plan de negocio (hoja «Cuadros» e índice, y una hoja por cuadro: C01 a C20 y RRHH, con modelo, plan publicado y diferencia explicada). Los contadores de Detalle_24m pasaron al motor. Incluye «Claude Log» (turnos 1 a 21) y «Esp_25Productos». |
| `Modelo_Maquila_ECOSTAR_v2.xlsx` | Salida directa del generador (misma estructura que la v7, sin las hojas de registro). |
| `Plantilla_Maquila_Comun.xlsx` | Plantilla vacía para Alianza del Acero y la tercera maquila. |
| `Informe_Plantilla_Comun_v2.md` | Evaluación de Alianza del Acero, cambios v1 → v2, resultados y guía de carga. |
| `Conciliacion_Plan_vs_Modelo_ECOSTAR.md` | Conciliación del modelo con el plan de negocio (Oct-2025): trazabilidad de cada dato a su página, diferencias línea por línea y puente del EBITDA. |
| `Pruebas_Escenarios_v7.md` / `Pruebas_Escenarios_ECOSTAR_v7.xlsx` | Todas las pruebas sobre la v7: escenarios, operación, proyección de 24 meses y cuadros con la estructura del plan, más la regresión frente a la v6. |
| `Pruebas_Escenarios_v6.md` / `Pruebas_Escenarios_ECOSTAR_v6.xlsx` | Todas las pruebas sobre la v6: 39 escenarios (495 comprobaciones, regresión frente a la v4), 7 de operación y 60 de la proyección de 24 meses. |
| `Pruebas_Escenarios_v4.md` / `Pruebas_Escenarios_ECOSTAR_v4.xlsx` | Pruebas paso a paso de la v4: 39 escenarios, 7 pruebas de operación y regresión frente a 12 productos (489/489 OK). |
| `Pruebas_Escenarios_v2.md` / `Pruebas_Escenarios_ECOSTAR_v2.xlsx` | Batería de 36 escenarios (financiamiento, cobranza, costos, ventas, plazos, corto vs largo plazo, TC) y verificaciones del motor. |
| `Prompt_Claude_Excel_ECOSTAR_v7.md` | Prompt para que Claude en Excel evalúe la v7 paso a paso (incluye las pruebas de los cuadros con la estructura del plan). |
| `Paso_a_Paso_Validacion_ECOSTAR_v7.md` | Recorrido de la validación de la v7 en Claude en Excel: qué revisar en cada turno y lista de control. |
| `Prompt_Claude_Excel_ECOSTAR_v6.md` | Prompt para que Claude en Excel evalúe la v6 paso a paso (incluye las pruebas de la proyección de 24 meses). |
| `Paso_a_Paso_Validacion_ECOSTAR_v6.md` | Recorrido de la validación de la v6 en Claude en Excel (ya realizada: 439/439 OK). |
| `Prompt_Claude_Excel_ECOSTAR_v4.md` | Prompt para que Claude en Excel evalúe la v4 paso a paso (celdas v4, 39 escenarios con valores esperados, prueba de reales). |
| `Prompt_Claude_Excel_ECOSTAR_v2.md` | Prompt de contexto para trabajar el modelo Ecostar v2 con Claude en Excel (validación, celdas clave, 36 escenarios de referencia). |
| `Prompt_Guia_Uso_ECOSTAR_v2.md` | Prompt para que Claude en Excel guíe paso a paso el uso del modelo (carga, cierre mensual, escenarios, lectura de resultados, proyección de 24 meses, cuadros con la estructura del plan, criterios y errores frecuentes). |
| `versiones/Modelo_Maquila_ECOSTAR_v6.xlsx` | Versión anterior (sin los cuadros con la estructura del plan). |
| `versiones/Modelo_Maquila_ECOSTAR_v6_validacion_Excel.xlsx` | Copia de pruebas devuelta por Claude en Excel (27/10/2026): «Claude Log» de la validación e «Informe_Cierre_v6» (439/439 OK). Es evidencia; no es para trabajar. |
| `versiones/Modelo_Maquila_ECOSTAR_v5.xlsx` | Versión anterior (sin la proyección de 24 meses). |
| `versiones/Modelo_Maquila_ECOSTAR_v4.xlsx` | Versión anterior (igual a la v5 sin la guía de carga). |
| `versiones/v1_Modelo_Financiero_Maquila_CIE_PERG_detallado.xlsx` | Versión previa detallada (15 hojas), conservada para comparación. |
| `modelo/plantilla/` | Generador de la plantilla v2 y datos de Ecostar. |
| `modelo/v1/` | Generador y pruebas de la versión 1. |
