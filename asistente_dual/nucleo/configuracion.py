"""Carga de config.toml y perfil.md."""
import tomllib
from pathlib import Path

BASE = Path(__file__).resolve().parent.parent


def cargar_config(ruta=BASE / "config.toml"):
    with open(ruta, "rb") as f:
        return tomllib.load(f)


def cargar_perfil(ruta=BASE / "perfil.md"):
    return Path(ruta).read_text(encoding="utf-8").strip()


def cargar_plantillas(carpeta=BASE / "plantillas"):
    """Devuelve {nombre legible: texto} para cada plantilla .md."""
    plantillas = {}
    for archivo in sorted(Path(carpeta).glob("*.md")):
        texto = archivo.read_text(encoding="utf-8")
        titulo, _, cuerpo = texto.partition("\n")
        plantillas[titulo.lstrip("# ").strip()] = cuerpo.strip()
    return plantillas
