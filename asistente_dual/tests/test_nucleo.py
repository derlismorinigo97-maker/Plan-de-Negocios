"""Pruebas sin conexión: los proveedores se simulan, no se llama a ninguna API.
Uso: python -m pytest tests  (desde la carpeta asistente_dual)"""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from nucleo import anonimizador, archivos, modos
from nucleo.historial import Historial, TopeExcedido
from nucleo.proveedores import Respuesta, costo


def proveedor_simulado(registro):
    def llamar(proveedor, rol, mensaje):
        registro.append((proveedor, rol, mensaje))
        return Respuesta(f"[{proveedor}] respuesta {len(registro)}", proveedor, proveedor, 100, 50, 0.01)
    return llamar


def test_comparar_llama_a_ambos():
    reg = []
    pasos = modos.comparar(proveedor_simulado(reg), "¿VAN?")
    assert sorted(p for p, _, _ in reg) == ["claude", "gpt"]
    assert [p.respuesta.proveedor for p in pasos] == ["claude", "gpt"]


@pytest.mark.parametrize("principal", ["claude", "gpt"])
def test_revision_cruzada_alterna_proveedores(principal):
    reg = []
    pasos = modos.revision_cruzada(proveedor_simulado(reg), "Evaluá el proyecto", principal=principal)
    revisor = modos.otro(principal)
    assert [p for p, _, _ in reg] == [principal, revisor, principal]
    assert reg[1][1] == modos.ROL_REVISOR
    assert pasos[0].respuesta.texto in reg[1][2]      # el revisor recibe el borrador
    assert pasos[1].respuesta.texto in reg[2][2]      # el consolidador recibe la revisión
    assert len(pasos) == 3


def test_abogado_del_diablo_y_juez():
    reg = []
    modos.abogado_del_diablo(proveedor_simulado(reg), "¿Comprar la grúa?")
    assert [p for p, _, _ in reg] == ["claude", "gpt", "claude"]
    reg = []
    pasos = modos.juez(proveedor_simulado(reg), "¿Leasing o crédito?", principal="gpt")
    assert reg[-1][0] == "gpt" and reg[-1][1] == modos.ROL_JUEZ
    assert pasos[0].respuesta.texto in reg[-1][2] and pasos[1].respuesta.texto in reg[-1][2]


def test_armar_consulta_incluye_documentos():
    c = modos.armar_consulta("Analizá", [("flujo.xlsx", "Caja 100")])
    assert '<documento nombre="flujo.xlsx">' in c and "Caja 100" in c


def test_anonimizador_ida_y_vuelta():
    pares = anonimizador.leer_diccionario("ANDE = CLIENTE_A\nProyecto Norte II = PROYECTO_02\nProyecto Norte = PROYECTO_01\n# comentario")
    original = "El contrato con ANDE y el Proyecto Norte II; también proyecto norte."
    anon = anonimizador.anonimizar(original, pares)
    assert "ANDE" not in anon and "Norte" not in anon.replace("PROYECTO", "")
    assert "PROYECTO_02" in anon and anon.count("PROYECTO_01") == 1
    assert anonimizador.restaurar(anon, pares) == "El contrato con ANDE y el Proyecto Norte II; también Proyecto Norte."


def test_costo_y_tope(tmp_path):
    assert costo(1_000_000, 1_000_000, {"entrada": 2, "salida": 10}) == pytest.approx(12)
    h = Historial(tmp_path / "h.db")
    h.registrar_llamada(Respuesta("x", "claude", "Claude", 0, 0, 9.5))
    assert h.gasto_mes() == pytest.approx(9.5)
    h.verificar_tope(10, 0.4)
    with pytest.raises(TopeExcedido):
        h.verificar_tope(10, 0.6)


def test_extraer_excel_y_truncar():
    import io
    from openpyxl import Workbook
    wb = Workbook(); ws = wb.active; ws.title = "Flujo"
    ws.append(["Mes", "Caja"]); ws.append([1, 1500000])
    buf = io.BytesIO(); wb.save(buf)
    texto, aviso = archivos.extraer_texto("f.xlsx", buf.getvalue(), 10_000)
    assert "## Hoja: Flujo" in texto and "1500000" in texto and not aviso
    texto, aviso = archivos.extraer_texto("n.txt", b"a" * 50, 10)
    assert len(texto) == 10 and "truncado" in aviso
