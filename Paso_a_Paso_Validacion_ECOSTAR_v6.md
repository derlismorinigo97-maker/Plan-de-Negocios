# Validación del modelo Ecostar v6 en Claude en Excel — paso a paso

**Objetivo:** confirmar que Excel calcula lo mismo que LibreOffice. En LibreOffice, la v6 pasó:
- 39 escenarios con 495 de 495 comprobaciones;
- 7 pruebas de operación;
- 60 de 60 comprobaciones de la proyección de 24 meses.

Al final vas a tener una tabla con OK o diferencias por paso y todo registrado en la hoja «Claude Log».

## Qué se necesita

| Archivo | Para qué |
|---|---|
| `Modelo_Maquila_ECOSTAR_v6.xlsx` | Libro que se valida |
| `Prompt_Claude_Excel_ECOSTAR_v6.md` | Instrucciones y valores esperados que se pegan en Claude |
| Este documento | Qué revisar en cada turno antes de escribir «seguir» |
| `Prompt_Guia_Uso_ECOSTAR_v2.md` (opcional, después) | Aprender a usar el libro: carga, cierre mensual, escenarios, proyección de 24 meses |

## Antes de empezar

1. **Guardar una copia.** Conservar el original sin tocar y hacer «Guardar como» `Modelo_Maquila_ECOSTAR_v6_pruebas.xlsx`. Las pruebas cargan y borran datos: se trabaja solo sobre la copia.
2. **Cálculo automático.** Verificar que esté activado: Fórmulas > Opciones para el cálculo > Automático.
3. **Panel de Claude.** Abrirlo, de preferencia en un chat nuevo.
4. **Pegar el prompt.** Abrir `Prompt_Claude_Excel_ECOSTAR_v6.md`, copiar **todo** lo que está debajo de la línea «COPIAR DESDE AQUÍ» y pegarlo en el chat. Claude empieza por el paso 0.

## Recorrido turno por turno

Claude hace un paso por turno y espera tu «seguir». Antes de escribirlo, controlá la columna «Qué revisar».

| Turno | Paso | Qué hace Claude | Qué revisar |
|---|---|---|---|
| 1 | 0. Preparación | Verifica las hojas y el modo de cálculo, y recalcula | 14 hojas; «Claude Log» hasta el turno 8 |
| 2 | 1. Integridad | Compara los 25 indicadores de Resultados | VAN del proyecto 6.989.922; VAN de los socios 3.695.678; balance 0; en REVISAR solo C13 y C14; 0 criterios con decimales |
| 3 | 2. 25 productos | Corre X2, X3 y V4 | X2: ventas +2.197.064,63 y presupuesto sin cambio; X3: C11 en REVISAR; V4: C10 = 16 |
| 4 | 3. Reales congelados | Carga 6 meses reales y cambia a Conservador | Meses 1-6 iguales en los dos escenarios; caja del mes 6 = 1.500.000; VAN del proyecto 6.858.766 (Vigente) y 31.573 (Conservador) |
| 5 | 3b. Volumen y guía | Carga volúmenes en Detalle_24m | Ventas de la ventana 10.769.580; paso 29 «✔ 24 meses detallados» y, sin el mes 24, «⚠ 23 de 24 meses» |
| 6 | 3c. Precio, costos y TC | 5 casos, uno por vez | EBITDA del año 1: materia prima 413.656; personal 479.713; fijo 499.614; TC 584.427; ventas con precio 280: 4.970.000 |
| 7 | 3d. Escenario y ventanas | Conservador, ventana de 12 meses y ventana móvil | Gabinetes por mes: 0/990 (Conservador); 1.100/1.250 (ventana de 12); 600/1.100/1.500 (ventana móvil) |
| 8 | 4. Financiamiento y cobranza (F) | 8 escenarios | F3: VAN del proyecto 6.230.650 |
| 9 | 5. Costos (C) | 4 escenarios | C2: VAN del proyecto 249.504 |
| 10 | 6. Ventas (V) | 5 escenarios | V3: ventas 56.855.713 |
| 11 | 7. Plan desfavorable (P) | 3 escenarios | P2: VAN del proyecto −486.490 |
| 12 | 8. Plazo del préstamo (L) | 4 escenarios | L1: cancela en 3,0 años |
| 13 | 9. Corto vs largo plazo (K) | 8 escenarios | K5: aporte adicional 2.191.679 |
| 14 | 10. Tipo de cambio (T) | 3 escenarios | T1: VAN del proyecto 7.989.367 |
| 15 | 11. Informe final | Tabla de todos los pasos y conclusión | Todo OK o lista de diferencias; cierre registrado en «Claude Log» |

**Después de cada turno:**
- Confirmar que Claude restauró el caso base: VAN del proyecto 6.989.922 y balance 0.
- Guardar la copia con Ctrl+G (Ctrl+S) para no perder el «Claude Log».

## Cómo leer el resultado

- **OK:** coincide dentro de la tolerancia (±1 USD en montos, ±0,01 en ratios).
- **Diferencia:** Claude no la corrige. Informa la celda, el valor esperado, el de Excel y su hipótesis.
  - Anotala y traé la tabla a la conversación de Claude Code.
  - La corrección se hace en el generador para que Alianza del Acero y la tercera maquila la hereden.
- **Causas más probables de una diferencia:**
  - **Restauración incompleta:** una celda vigente quedó con un valor en lugar de su fórmula (las grises heredan del presupuesto, por ejemplo `=G8`).
  - **Cálculo en manual:** Excel no recalculó.
  - **Configuración regional:** un criterio de texto con decimales dentro de un SUMIF o COUNTIF. El libro no tiene ninguno; si Claude agrega uno, cuenta mal.
  - **Referencia circular:** no debería aparecer. Si Excel avisa, hay que informarlo.

## Criterios a cuidar durante la validación

- No se tocan Calculo, Resultados, Control ni Proyeccion_24m, ni filas, columnas o nombres definidos.
- **Filas con fórmula en Detalle_24m:** las filas «Celdas con dato en el mes» (11, 39, 67, 95, 123, 151, 184, 217 y 245) no se borran. Para limpiar el detalle, Claude debe borrar solo las filas de datos de cada bloque. En la primera validación se borraron con un rango amplio y hubo que restaurarlas.
- **Un paso por turno.** Si Claude intenta hacer varios juntos, pedile que se detenga y siga de a uno.
- **Si el chat se alarga y Claude pierde el contexto:** abrir un chat nuevo, pegar de nuevo el prompt y escribir «Seguimos desde el paso N; lo hecho está en Claude Log».
- La copia de pruebas no se usa como modelo de trabajo. Para cargar datos reales se parte del original v6.

## Al terminar

1. **Libro restaurado:**
   - VAN del proyecto 6.989.922 y balance 0;
   - Detalle_24m vacío;
   - Inicio!C11 vacío y Reales sin datos.
2. **Copia de pruebas.** Guardarla con el «Claude Log» completo.
3. **Diferencias.** Si las hubo, traer la tabla del paso 11 para revisarlas.
4. **Siguiente etapa.** Pegar `Prompt_Guia_Uso_ECOSTAR_v2.md` para aprender a usar el libro. El módulo 13 cubre la proyección de 24 meses y la revisión del plan.

## Lista de control

- [ ] Copia `…_v6_pruebas.xlsx` creada; cálculo automático
- [ ] Paso 0 — Preparación
- [ ] Paso 1 — Integridad del libro base
- [ ] Paso 2 — Estructura de 25 productos (X2, X3, V4)
- [ ] Paso 3 — Meses reales congelados
- [ ] Paso 3b — Volumen y guía
- [ ] Paso 3c — Precio, costos y TC
- [ ] Paso 3d — Escenario, ventana de 12 meses y ventana móvil
- [ ] Pasos 4 a 10 — Catálogo F, C, V, P, L, K, T
- [ ] Paso 11 — Informe final y cierre en «Claude Log»
- [ ] Libro restaurado y diferencias (si las hay) enviadas para corregir
