"""Reemplaza términos sensibles por códigos antes de enviar a las APIs y los restituye en la respuesta.
Formato del diccionario (uno por línea):  Término real = CÓDIGO   (ej.: ANDE = CLIENTE_A)"""
import re


def leer_diccionario(texto):
    pares = {}
    for linea in (texto or "").splitlines():
        if "=" in linea and not linea.strip().startswith("#"):
            real, _, codigo = linea.partition("=")
            if real.strip() and codigo.strip():
                pares[real.strip()] = codigo.strip()
    return pares


def _reemplazar(texto, pares):
    # Primero los términos más largos, para que "Proyecto Norte II" no quede a medias por "Proyecto Norte".
    for origen in sorted(pares, key=len, reverse=True):
        texto = re.sub(re.escape(origen), lambda _m, d=pares[origen]: d, texto, flags=re.IGNORECASE)
    return texto


def anonimizar(texto, pares):
    return _reemplazar(texto, pares)


def restaurar(texto, pares):
    return _reemplazar(texto, {codigo: real for real, codigo in pares.items()})
