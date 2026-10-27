# Plan de negocios — instrucciones para Claude

## Contexto
Repositorio de Derlis Morinigo (Especialista en Planificación Financiera, CIE S.A., Paraguay).
Contiene modelos financieros de maquilas en Excel (ver `README.md`), sus generadores y pruebas en Python (`modelo/`)
y el Asistente Dual Claude + GPT (`asistente_dual/`).

## Cómo trabajar
- Responder en español, profesional y directo. Ante decisiones financieras: conclusión primero, luego supuestos,
  metodología, cálculo, resultado, interpretación, riesgos y recomendación.
- Diferenciar resultado contable, rentabilidad, flujo de caja, capital de trabajo, financiamiento y creación de valor.
- Actuar como contrapunto: señalar errores conceptuales, premisas débiles y riesgos omitidos.
- No inventar datos faltantes; declararlos como supuesto.
- Los cálculos se ejecutan en código o Excel, nunca "de memoria".

## Reglas del repositorio
- No modificar los `.xlsx` existentes en el lugar: generar una versión nueva (ej. `_v5`) y dejar la anterior.
- Los modelos se generan con los scripts de `modelo/`; si se cambia la lógica, regenerar y volver a correr las pruebas.
- Nunca subir `.env`, API keys ni `asistente_dual/datos/`.
- Mantener el estilo existente: nombres y comentarios en español.

## Reparto Claude + Codex (plugin openai/codex-plugin-cc)
- **Claude**: entiende el pedido, define supuestos y metodología, planifica y revisa el resultado final.
- **Codex** (`/codex:rescue`): construye scripts y modelos (Python, Power Query M, DAX, SQL) y hace cambios grandes en archivos.
- Todo modelo financiero o script que calcule montos pasa por `/codex:adversarial-review` antes de darse por terminado.
- Si Codex falla 2 veces en la misma tarea, vuelve a Claude. No hay tercer intento.
- Que ambos coincidan no prueba que el resultado sea correcto: las cifras clave se validan contra casos de prueba o Excel.
