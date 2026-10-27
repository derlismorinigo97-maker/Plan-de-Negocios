"""Conexión con las APIs de Anthropic (Claude) y OpenAI (GPT).
Las API keys se leen de variables de entorno (archivo .env): ANTHROPIC_API_KEY y OPENAI_API_KEY."""
from dataclasses import dataclass

import anthropic
import openai


@dataclass
class Respuesta:
    texto: str
    proveedor: str        # "claude" | "gpt"
    modelo: str           # nombre legible del modelo
    tokens_entrada: int
    tokens_salida: int
    costo_usd: float
    aviso: str = ""       # rechazo, fallback u otra observación para el usuario


class ErrorProveedor(Exception):
    """Error de API traducido a un mensaje entendible."""


def costo(tokens_entrada, tokens_salida, precios):
    return tokens_entrada * precios["entrada"] / 1e6 + tokens_salida * precios["salida"] / 1e6


_clientes = {}


def _cliente(proveedor):
    if proveedor not in _clientes:
        _clientes[proveedor] = anthropic.Anthropic() if proveedor == "claude" else openai.OpenAI()
    return _clientes[proveedor]


def llamar_claude(nombre, cfg_modelo, sistema, mensaje, esfuerzo="high"):
    parametros = dict(
        model=cfg_modelo["id"],
        max_tokens=64000,
        system=sistema,
        messages=[{"role": "user", "content": mensaje}],
        output_config={"effort": esfuerzo},
    )
    if cfg_modelo.get("fallback"):
        # Si los filtros de seguridad rechazan la consulta, la API la reintenta con otro modelo de Claude.
        parametros.update(betas=["server-side-fallback-2026-07-01"], fallbacks="default")
    try:
        with _cliente("claude").beta.messages.stream(**parametros) as flujo:
            r = flujo.get_final_message()
    except anthropic.AuthenticationError:
        raise ErrorProveedor("Claude: API key inválida o ausente (revisá ANTHROPIC_API_KEY en .env).")
    except anthropic.RateLimitError:
        raise ErrorProveedor("Claude: límite de uso alcanzado. Esperá un momento y reintentá.")
    except anthropic.APIStatusError as e:
        raise ErrorProveedor(f"Claude: error {e.status_code} de la API: {e.message}")
    except anthropic.APIConnectionError:
        raise ErrorProveedor("Claude: no se pudo conectar con la API (revisá la conexión a internet).")

    aviso = ""
    if r.stop_reason == "refusal":
        aviso = "Claude rechazó la consulta (filtro de seguridad). La respuesta puede estar vacía o incompleta."
    elif r.stop_reason == "max_tokens":
        aviso = "La respuesta de Claude se cortó por longitud."
    if any(b.type == "fallback" for b in r.content):
        aviso = (aviso + " " if aviso else "") + f"Respondió el modelo de respaldo {r.model}."
    texto = "".join(b.text for b in r.content if b.type == "text")

    # Con fallback, usage.iterations detalla cada intento; sin él, usage cubre la única llamada.
    intentos = getattr(r.usage, "iterations", None) or [r.usage]
    t_in = sum(getattr(i, "input_tokens", 0) or 0 for i in intentos)
    t_out = sum(getattr(i, "output_tokens", 0) or 0 for i in intentos)
    return Respuesta(texto, "claude", nombre, t_in, t_out, costo(t_in, t_out, cfg_modelo), aviso)


def llamar_gpt(nombre, cfg_modelo, sistema, mensaje):
    try:
        r = _cliente("gpt").responses.create(
            model=cfg_modelo["id"],
            instructions=sistema,
            input=mensaje,
            max_output_tokens=32000,
        )
    except openai.AuthenticationError:
        raise ErrorProveedor("GPT: API key inválida o ausente (revisá OPENAI_API_KEY en .env).")
    except openai.RateLimitError:
        raise ErrorProveedor("GPT: límite de uso o saldo insuficiente. Revisá tu saldo en platform.openai.com.")
    except openai.NotFoundError:
        raise ErrorProveedor(f"GPT: el modelo '{cfg_modelo['id']}' no existe. Corregí el ID en config.toml.")
    except openai.APIStatusError as e:
        raise ErrorProveedor(f"GPT: error {e.status_code} de la API: {e.message}")
    except openai.APIConnectionError:
        raise ErrorProveedor("GPT: no se pudo conectar con la API (revisá la conexión a internet).")

    aviso = "La respuesta de GPT quedó incompleta." if r.status == "incomplete" else ""
    t_in, t_out = r.usage.input_tokens, r.usage.output_tokens
    return Respuesta(r.output_text, "gpt", nombre, t_in, t_out, costo(t_in, t_out, cfg_modelo), aviso)
