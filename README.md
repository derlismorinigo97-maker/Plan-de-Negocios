# Plan de negocios — modelos financieros de maquilas

| Archivo | Uso |
|---|---|
| `Modelo_Maquila_ECOSTAR_v4.xlsx` | **Versión de trabajo actual**: Ecostar con 25 productos, más las hojas «Claude Log» y «Esp_25Productos» de Claude en Excel. |
| `Modelo_Maquila_ECOSTAR_v2.xlsx` | Salida directa del generador (misma estructura que la v4, sin las hojas de registro). |
| `Plantilla_Maquila_Comun.xlsx` | Plantilla vacía para Alianza del Acero y la tercera maquila. |
| `Informe_Plantilla_Comun_v2.md` | Evaluación de Alianza del Acero, cambios v1 → v2, resultados y guía de carga. |
| `alianza/Revision_Informe_Alianza_RV00.md` | Revisión del informe de cuentas con Alianza del Acero (Excel y PPT RV00): verificaciones, observaciones, sensibilidad del saldo y cambios sugeridos. |
| `alianza/ALIANZA_Informe_Ejecutivo_RV01.pptx` | Presentación ejecutiva con las correcciones de cifras y textos de la revisión; los puntos de decisión quedan abiertos. |
| `alianza/Informe_Financiero_ALIANZA_DEL_ACERO_INFORME2_09062026_AJUSTADO_A_10_AÑOS_CON_GUIA.xlsx` | Modelo a 10 años de Alianza (INFORME2, sin cambios en el cálculo) con dos hojas nuevas: «PASO A PASO» para actualizarlo con nuevas proyecciones de ventas y «CONTROLES Y PPT» con 26 controles vivos y los datos para el PowerPoint. La fase B se validó aplicándola sobre una copia: quedan 25 de 26 controles en OK y el VAN no cambia; el que sigue en REVISAR es el saldo de caja, que se resuelve en la fase E. |
| `alianza/Guia_Paso_a_Paso_Actualizacion_Plan_ALIANZA_DEL_ACERO.docx` | La misma guía paso a paso en Word para circularla: fases A-G, criterios, 26 controles, mapa de cada diapositiva del PPT a su celda de origen, decisiones pendientes, registro de actualizaciones y efecto de la fase B en la caja (anexo A). |
| `alianza/Comparacion_PPT_Plan10_vs_Excel.md` / `alianza/Trazabilidad_PPT_Plan10_vs_Excel.xlsx` | Comparación del PPT «Plan Original vs Proyección Actualizada a 10 años» contra los Excel ORIGINAL e INFORME2: 243 cifras rastreadas hasta su celda. |
| `Pruebas_Escenarios_v4.md` / `Pruebas_Escenarios_ECOSTAR_v4.xlsx` | Pruebas paso a paso de la v4: 39 escenarios, 7 pruebas de operación y regresión frente a 12 productos (489/489 OK). |
| `Pruebas_Escenarios_v2.md` / `Pruebas_Escenarios_ECOSTAR_v2.xlsx` | Batería de 36 escenarios (financiamiento, cobranza, costos, ventas, plazos, corto vs largo plazo, TC) y verificaciones del motor. |
| `Prompt_Claude_Excel_ECOSTAR_v4.md` | Prompt para que Claude en Excel evalúe la v4 paso a paso (celdas v4, 39 escenarios con valores esperados, prueba de reales). |
| `Prompt_Claude_Excel_ECOSTAR_v2.md` | Prompt de contexto para trabajar el modelo Ecostar v2 con Claude en Excel (validación, celdas clave, 36 escenarios de referencia). |
| `Prompt_Guia_Uso_ECOSTAR_v2.md` | Prompt para que Claude en Excel guíe paso a paso el uso del modelo (carga, cierre mensual, escenarios, lectura de resultados, criterios y errores frecuentes). |
| `versiones/v1_Modelo_Financiero_Maquila_CIE_PERG_detallado.xlsx` | Versión previa detallada (15 hojas), conservada para comparación. |
| `modelo/plantilla/` | Generador de la plantilla v2 y datos de Ecostar. |
| `modelo/v1/` | Generador y pruebas de la versión 1. |
