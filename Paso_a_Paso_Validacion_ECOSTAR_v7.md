# Validación del modelo Ecostar v7 en Claude en Excel — paso a paso

**Objetivo:** confirmar que Excel calcula lo mismo que LibreOffice, sobre todo en lo nuevo de la v7: los cuadros con la estructura del plan de negocio.

**Qué se sabe antes de empezar:**
- El motor de la v7 es idéntico al de la v6, celda por celda, y la v6 ya pasó la validación en Excel (439 de 439 OK).
- En LibreOffice, la v7 pasó todas las pruebas: 39 escenarios con **495 de 495 comprobaciones OK**, idénticos a la v6 (1.625.364 valores); 7 pruebas de operación iguales a la v6; proyección de 24 meses **60 de 60 OK**; cuadros con la estructura del plan: 7 casos y **147 de 147 OK**; motor, Resultados, Proyeccion_24m, guía y Control idénticos a la v6 celda por celda.

**Recomendación: ruta corta.** Pasos 0, 1, 3e, 3f y 11, en 5 turnos. Valida lo nuevo y confirma que el libro base no cambió. La ruta completa (15 + 2 turnos) repite el catálogo de escenarios ya validado en la v6: hacela solo si querés la evidencia completa sobre la v7.

Al final vas a tener una tabla con OK o diferencias por paso y todo registrado en la hoja «Claude Log».

## Qué se necesita

| Archivo | Para qué |
|---|---|
| `Modelo_Maquila_ECOSTAR_v7.xlsx` | Libro que se valida |
| `Prompt_Claude_Excel_ECOSTAR_v7.md` | Instrucciones y valores esperados que se pegan en Claude |
| Este documento | Qué revisar en cada turno antes de escribir «seguir» |
| `Prompt_Guia_Uso_ECOSTAR_v2.md` (opcional, después) | Aprender a usar el libro: carga, cierre mensual, escenarios, proyección de 24 meses y cuadros (módulo 14) |

## Antes de empezar

1. **Guardar una copia.** Conservar el original sin tocar y hacer «Guardar como» `Modelo_Maquila_ECOSTAR_v7_pruebas.xlsx`. Las pruebas cargan y borran datos: se trabaja solo sobre la copia.
2. **Cálculo automático.** Verificar que esté activado: Fórmulas > Opciones para el cálculo > Automático.
3. **Panel de Claude.** Abrirlo, de preferencia en un chat nuevo.
4. **Pegar el prompt.** Abrir `Prompt_Claude_Excel_ECOSTAR_v7.md`, copiar **todo** lo que está debajo de la línea «COPIAR DESDE AQUÍ» y pegarlo en el chat. Para la ruta corta, agregar al final: «Ruta corta: pasos 0, 1, 3e, 3f y 11». Claude empieza por el paso 0.

## Recorrido turno por turno

Claude hace un paso por turno y espera tu «seguir». Antes de escribirlo, controlá la columna «Qué revisar».

**Ruta corta (recomendada):**

| Turno | Paso | Qué hace Claude | Qué revisar |
|---|---|---|---|
| 1 | 0. Preparación | Verifica las hojas y el modo de cálculo, y recalcula | 30 hojas; «Claude Log» hasta el turno 21 |
| 2 | 1. Integridad | Compara los 25 indicadores de Resultados | VAN del proyecto 6.989.922; VAN de los socios 3.695.678; balance 0; en REVISAR solo C13 y C14 |
| 3 | 3e. Cuadros: lectura y controles | Lee la hoja Cuadros y los controles de C15, C18 y C19; prueba el selector de versión | EBITDA −1.044.506 frente al plan; VAN método del plan 4.177.576 frente a 4.588.736; control del cuadro 18 = 0; VAN recalculado con los flujos del plan = 4.588.737 |
| 4 | 3f. Cuadros con escenario y detalle | Conservador y materia prima detallada | Conservador: año 0 sin ventas y año 1 = 4.146.750; ajuste de materia prima del año 1 = 88.200 |
| 5 | 11. Informe final | Tabla de los pasos y conclusión | Todo OK o lista de diferencias; cierre en «Claude Log» |

**Ruta completa:** los pasos de la ruta corta más los de la tabla siguiente, en el orden del prompt (3e y 3f van después de 3d).

| Turno | Paso | Qué hace Claude | Qué revisar |
|---|---|---|---|
| 1 | 0. Preparación | Verifica las hojas y el modo de cálculo, y recalcula | 30 hojas; «Claude Log» hasta el turno 21 |
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
| 15-17 | 3e, 3f y 11 | Cuadros e informe final | Como en la ruta corta |

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

- No se tocan Calculo, Resultados, Control, Proyeccion_24m ni las hojas de cuadros (salvo Cuadros!C5 en el paso 3e), ni filas, columnas o nombres definidos.
- **Detalle_24m:** en la v7 los contadores están en Calculo (filas det.*), así que limpiar el detalle ya no borra fórmulas. Igual conviene borrar solo los rangos cargados.
- **Un paso por turno.** Si Claude intenta hacer varios juntos, pedile que se detenga y siga de a uno.
- **Si el chat se alarga y Claude pierde el contexto:** abrir un chat nuevo, pegar de nuevo el prompt y escribir «Seguimos desde el paso N; lo hecho está en Claude Log».
- La copia de pruebas no se usa como modelo de trabajo. Para cargar datos reales se parte del original v7.

## Al terminar

1. **Libro restaurado:**
   - VAN del proyecto 6.989.922 y balance 0;
   - Detalle_24m vacío;
   - Inicio!C11 vacío, Inicio!C13 = «Vigente» y Reales sin datos;
   - Cuadros!C5 = «Vigente».
2. **Copia de pruebas.** Guardarla con el «Claude Log» completo.
3. **Diferencias.** Si las hubo, traer la tabla del paso 11 para revisarlas.
4. **Siguiente etapa.** Pegar `Prompt_Guia_Uso_ECOSTAR_v2.md` para aprender a usar el libro. El módulo 13 cubre la proyección de 24 meses y el módulo 14, los cuadros con la estructura del plan.

## Lista de control

- [ ] Copia `…_v7_pruebas.xlsx` creada; cálculo automático
- [ ] Paso 0 — Preparación
- [ ] Paso 1 — Integridad del libro base
- [ ] Paso 3e — Cuadros: lectura y controles
- [ ] Paso 3f — Cuadros con escenario y detalle mensual
- [ ] Paso 11 — Informe final y cierre en «Claude Log» (ruta corta: termina acá)
- [ ] Ruta completa: Paso 2 — Estructura de 25 productos (X2, X3, V4)
- [ ] Paso 3 — Meses reales congelados
- [ ] Paso 3b — Volumen y guía
- [ ] Paso 3c — Precio, costos y TC
- [ ] Paso 3d — Escenario, ventana de 12 meses y ventana móvil
- [ ] Pasos 4 a 10 — Catálogo F, C, V, P, L, K, T
- [ ] Libro restaurado y diferencias (si las hay) enviadas para corregir
