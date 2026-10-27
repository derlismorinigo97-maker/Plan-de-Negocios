# Plan de negocios — modelos financieros de maquilas

| Archivo | Uso |
|---|---|
| `Modelo_Maquila_ECOSTAR_v6.xlsx` | **Versión de trabajo actual**: 25 productos, «Guia_Carga», proyección rodante de 24 meses («Detalle_24m» y «Proyeccion_24m») con vista anual a 10 años y datos de revisión del plan, más «Claude Log» y «Esp_25Productos». |
| `Modelo_Maquila_ECOSTAR_v2.xlsx` | Salida directa del generador (misma estructura que la v6, sin las hojas de registro). |
| `Plantilla_Maquila_Comun.xlsx` | Plantilla vacía para Alianza del Acero y la tercera maquila. |
| `Informe_Plantilla_Comun_v2.md` | Evaluación de Alianza del Acero, cambios v1 → v2, resultados y guía de carga. |
| `Conciliacion_Plan_vs_Modelo_ECOSTAR.md` | Conciliación del modelo con el plan de negocio (Oct-2025): trazabilidad de cada dato a su página, diferencias línea por línea y puente del EBITDA. |
| `Pruebas_Escenarios_v6.md` / `Pruebas_Escenarios_ECOSTAR_v6.xlsx` | Todas las pruebas sobre la v6: 39 escenarios (495 comprobaciones, regresión frente a la v4), 7 de operación y 60 de la proyección de 24 meses. |
| `Pruebas_Escenarios_v4.md` / `Pruebas_Escenarios_ECOSTAR_v4.xlsx` | Pruebas paso a paso de la v4: 39 escenarios, 7 pruebas de operación y regresión frente a 12 productos (489/489 OK). |
| `Pruebas_Escenarios_v2.md` / `Pruebas_Escenarios_ECOSTAR_v2.xlsx` | Batería de 36 escenarios (financiamiento, cobranza, costos, ventas, plazos, corto vs largo plazo, TC) y verificaciones del motor. |
| `Prompt_Claude_Excel_ECOSTAR_v6.md` | Prompt para que Claude en Excel evalúe la v6 paso a paso (incluye las pruebas de la proyección de 24 meses). |
| `Prompt_Claude_Excel_ECOSTAR_v4.md` | Prompt para que Claude en Excel evalúe la v4 paso a paso (celdas v4, 39 escenarios con valores esperados, prueba de reales). |
| `Prompt_Claude_Excel_ECOSTAR_v2.md` | Prompt de contexto para trabajar el modelo Ecostar v2 con Claude en Excel (validación, celdas clave, 36 escenarios de referencia). |
| `Prompt_Guia_Uso_ECOSTAR_v2.md` | Prompt para que Claude en Excel guíe paso a paso el uso del modelo (carga, cierre mensual, escenarios, lectura de resultados, criterios y errores frecuentes). |
| `versiones/Modelo_Maquila_ECOSTAR_v5.xlsx` | Versión anterior (sin la proyección de 24 meses). |
| `versiones/Modelo_Maquila_ECOSTAR_v4.xlsx` | Versión anterior (igual a la v5 sin la guía de carga). |
| `versiones/v1_Modelo_Financiero_Maquila_CIE_PERG_detallado.xlsx` | Versión previa detallada (15 hojas), conservada para comparación. |
| `modelo/plantilla/` | Generador de la plantilla v2 y datos de Ecostar. |
| `modelo/v1/` | Generador y pruebas de la versión 1. |
