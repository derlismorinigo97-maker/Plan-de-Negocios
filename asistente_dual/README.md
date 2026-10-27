# Asistente Dual · Claude + GPT

App web local (Streamlit) que hace trabajar juntos a Claude y GPT sobre análisis financieros.
Usa las **APIs** de Anthropic y OpenAI (pago por uso, aparte de las suscripciones Pro).

## Modos de trabajo

| Modo | Flujo | Llamadas |
|---|---|---|
| **Comparar** | La misma consulta a ambos; respuestas lado a lado | 2 |
| **Revisión cruzada** | A elabora → B critica como Gerente Financiero exigente → A acepta o rechaza cada observación y entrega la versión final | 3 |
| **Abogado del diablo** | A defiende la propuesta → B la ataca → A sintetiza una recomendación | 3 |
| **Juez** | Ambos responden por separado → A compara, detecta errores y entrega una respuesta integrada | 3 |

En los modos de tres pasos elegís quién es A (Claude o GPT); B es siempre el otro.

## Otras funciones

- **Perfil fijo** (`perfil.md`): tu contexto y estilo de respuesta se envían en cada consulta. Editalo libremente.
- **Plantillas** (`plantillas/`): evaluación de proyecto, informe ejecutivo, desviaciones, revisión de modelo y correo. Para agregar una, creá un `.md` cuya primera línea sea `# Título`.
- **Archivos de soporte**: Excel (valores calculados, no fórmulas), PDF, Word, CSV y texto.
- **Anonimizador**: en la barra lateral, `Término real = CÓDIGO` por línea. Se reemplaza antes de enviar y se restituye al mostrar.
- **Control de gasto**: costo por paso y por consulta en USD y Gs; bloqueo al alcanzar el tope mensual (`config.toml`).
- **Historial local** en `datos/historial.db` (no se sube a GitHub) y exportación a Markdown y Word.

## Instalación (Windows, una sola vez)

1. Instalar Python 3.11 o superior desde python.org (marcar «Add Python to PATH»).
2. Crear las API keys y cargar saldo:
   - Claude: console.anthropic.com → API Keys. En *Billing* fijá un límite de gasto mensual.
   - GPT: platform.openai.com → API keys. En *Limits* fijá un límite de gasto mensual.
3. En una terminal, dentro de la carpeta `asistente_dual`:
   ```
   python -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   copy .env.example .env
   ```
4. Abrir `.env` con el Bloc de notas y pegar las dos API keys.
5. Revisar en `config.toml` el tope mensual, el tipo de cambio y los IDs de modelo de OpenAI.

## Uso diario

```
cd asistente_dual
.venv\Scripts\activate
streamlit run app.py
```
Se abre en el navegador (http://localhost:8501). Para cerrarla, Ctrl+C en la terminal.

## Pruebas

```
pip install pytest
python -m pytest tests
```
Las pruebas simulan los proveedores, así que no consumen saldo.

## Precauciones

- **Confidencialidad**: no cargues datos de la empresa sin confirmar la política interna de uso de IA. Por defecto, las APIs no usan tus datos para entrenar, pero la información sale de la empresa.
- **Falso consenso**: que ambos modelos coincidan no garantiza que el resultado sea correcto.
- **Cálculos**: verificá en Excel los números que vayas a presentar.
- **API keys**: nunca compartas ni subas el archivo `.env`. Si una key se filtra, revocala en la consola correspondiente.
- **Costos**: en los modos de tres pasos, la consulta y los archivos se reenvían en cada paso. Con archivos grandes, el costo se multiplica.
