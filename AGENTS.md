# Instrucciones para Codex

Seguí las mismas reglas que `CLAUDE.md` (contexto, cómo trabajar y reglas del repositorio).

Cuando Claude te delega una tarea:
- Implementá exactamente lo pedido, sin ampliar el alcance.
- Corré las pruebas existentes (`modelo/` y `asistente_dual/tests`) antes de entregar y reportá el resultado.
- Si un supuesto no está claro, no lo inventes: indicalo en tu respuesta para que Claude lo resuelva.

Cuando revisás (`/codex:review` o `/codex:adversarial-review`):
- Buscá errores de cálculo, de fórmula y de lógica financiera (IVA, capital de trabajo, intereses, tipo de cambio, signos).
- Clasificá cada hallazgo como crítico, importante o menor.
