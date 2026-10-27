"""Asistente Dual: Claude + GPT trabajando juntos sobre análisis financieros.
Uso: streamlit run app.py   (desde la carpeta asistente_dual, con el archivo .env cargado)"""
import os

import streamlit as st
from dotenv import load_dotenv

from nucleo import anonimizador, archivos, modos
from nucleo.configuracion import BASE, cargar_config, cargar_perfil, cargar_plantillas
from nucleo.historial import Historial, TopeExcedido
from nucleo.proveedores import ErrorProveedor, llamar_claude, llamar_gpt

load_dotenv(BASE / ".env")
st.set_page_config(page_title="Asistente Dual", page_icon="⚖️", layout="wide")

cfg = cargar_config()
perfil = cargar_perfil()
plantillas = cargar_plantillas()
TOPE = cfg["general"]["tope_mensual_usd"]
TC = cfg["general"]["tipo_cambio_pyg"]


@st.cache_resource
def historial():
    return Historial()


hist = historial()


def fmt_costo(usd):
    return f"USD {usd:,.3f} (≈ Gs {usd * TC:,.0f})".replace(",", "X").replace(".", ",").replace("X", ".")


# ---------------- Barra lateral ----------------
with st.sidebar:
    st.header("Configuración")
    gastado = hist.gasto_mes()
    st.metric("Gasto del mes", f"USD {gastado:,.2f}", help=f"Tope mensual: USD {TOPE:,.2f} (config.toml)")
    st.progress(min(gastado / TOPE, 1.0) if TOPE else 1.0)

    modelo_claude = st.selectbox("Modelo Claude", list(cfg["modelos"]["claude"]))
    modelo_gpt = st.selectbox("Modelo GPT", list(cfg["modelos"]["gpt"]))
    esfuerzo = st.select_slider("Profundidad de razonamiento (Claude)", ["low", "medium", "high", "xhigh", "max"],
                                value=cfg["general"]["esfuerzo_claude"],
                                help="Mayor profundidad = mejor análisis en problemas complejos, pero más tokens y más costo.")

    st.subheader("Anonimizador")
    diccionario = st.text_area("Términos a reemplazar antes de enviar (uno por línea)",
                               placeholder="ANDE = CLIENTE_A\nProyecto Línea Norte = PROYECTO_01",
                               help="Se reemplazan antes de enviar a las APIs y se restituyen al mostrar el resultado.")
    pares = anonimizador.leer_diccionario(diccionario)

    faltan = [k for k in ("ANTHROPIC_API_KEY", "OPENAI_API_KEY") if not os.getenv(k)]
    if faltan:
        st.error("Faltan API keys en .env: " + ", ".join(faltan))
    st.caption("No cargues datos confidenciales de la empresa sin confirmar la política interna de uso de IA.")

# ---------------- Contenido principal ----------------
st.title("Asistente Dual · Claude + GPT")
tab_consulta, tab_historial = st.tabs(["Consulta", "Historial"])

with tab_consulta:
    modo = st.radio("Modo de trabajo", list(modos.MODOS), horizontal=True)
    st.caption(modos.MODOS[modo])
    principal = "claude"
    if modo != "Comparar":
        etiqueta = {"Revisión cruzada": "¿Quién elabora y consolida?", "Abogado del diablo": "¿Quién defiende y sintetiza?",
                    "Juez": "¿Quién actúa como juez?"}[modo]
        principal = st.radio(etiqueta, ["claude", "gpt"], horizontal=True, format_func={"claude": "Claude", "gpt": "GPT"}.get)

    col_p, col_b = st.columns([4, 1])
    plantilla = col_p.selectbox("Plantilla", ["(ninguna)"] + list(plantillas))
    if col_b.button("Insertar", use_container_width=True) and plantilla != "(ninguna)":
        st.session_state["pregunta"] = plantillas[plantilla]

    pregunta = st.text_area("Consulta", key="pregunta", height=260)
    cargados = st.file_uploader("Archivos de soporte (opcional)", type=archivos.EXTENSIONES, accept_multiple_files=True)

    documentos = []
    for f in cargados or []:
        texto, aviso = archivos.extraer_texto(f.name, f.getvalue(), cfg["general"]["max_caracteres_archivo"])
        if aviso:
            st.warning(aviso)
        documentos.append((f.name, texto))

    consulta = anonimizador.anonimizar(modos.armar_consulta(pregunta, documentos), pares)
    tokens = archivos.estimar_tokens(perfil + consulta)
    precio_in = max(cfg["modelos"]["claude"][modelo_claude]["entrada"], cfg["modelos"]["gpt"][modelo_gpt]["entrada"])
    llamadas = {"Comparar": 2, "Revisión cruzada": 3, "Abogado del diablo": 3, "Juez": 3}[modo]
    st.caption(f"Consulta: ~{tokens:,} tokens · Entrada estimada para {llamadas} llamadas: {fmt_costo(tokens * llamadas * precio_in / 1e6)} "
               "(más la salida y el texto que se reenvía entre pasos)")

    if pares:
        with st.expander("Ver el texto tal como se enviará (anonimizado)"):
            st.text(consulta[:5000] + ("…" if len(consulta) > 5000 else ""))

    if st.button("Ejecutar", type="primary", disabled=not pregunta.strip() or bool(faltan)):
        def llamar(proveedor, rol, mensaje):
            sistema = perfil + "\n\n" + rol
            if proveedor == "claude":
                c = cfg["modelos"]["claude"][modelo_claude]
                hist.verificar_tope(TOPE, archivos.estimar_tokens(sistema + mensaje) * c["entrada"] / 1e6)
                r = llamar_claude(modelo_claude, c, sistema, mensaje, esfuerzo)
            else:
                c = cfg["modelos"]["gpt"][modelo_gpt]
                hist.verificar_tope(TOPE, archivos.estimar_tokens(sistema + mensaje) * c["entrada"] / 1e6)
                r = llamar_gpt(modelo_gpt, c, sistema, mensaje)
            hist.registrar_llamada(r)
            return r

        try:
            with st.spinner(f"Ejecutando «{modo}»… (puede tardar unos minutos)"):
                pasos = modos.EJECUTAR[modo](llamar, consulta, principal=principal)
        except (ErrorProveedor, TopeExcedido) as e:
            st.error(str(e))
        else:
            # Los pasos se encadenan con el texto anonimizado; se restituye solo para mostrar y exportar.
            secciones = [(p.titulo, anonimizador.restaurar(p.respuesta.texto, pares)) for p in pasos]
            total = sum(p.respuesta.costo_usd for p in pasos)
            md = "\n\n".join(f"## {t}\n\n{x}" for t, x in secciones)
            hist.registrar_consulta(modo, pregunta, md, total)
            st.session_state["resultado"] = dict(modo=modo, pasos=pasos, secciones=secciones, total=total, md=md)

    res = st.session_state.get("resultado")
    if res:
        st.divider()
        st.subheader(f"Resultado · {res['modo']} · {fmt_costo(res['total'])}")

        def mostrar(paso, texto):
            r = paso.respuesta
            st.caption(f"{r.modelo} · {r.tokens_entrada:,} tokens de entrada / {r.tokens_salida:,} de salida · {fmt_costo(r.costo_usd)}")
            if r.aviso:
                st.warning(r.aviso)
            st.markdown(texto)

        if res["modo"] == "Comparar":
            for col, paso, (titulo, texto) in zip(st.columns(2), res["pasos"], res["secciones"]):
                with col:
                    st.markdown(f"#### {titulo}")
                    mostrar(paso, texto)
        else:
            *intermedios, final = list(zip(res["pasos"], res["secciones"]))
            st.markdown(f"#### {final[1][0]}")
            mostrar(final[0], final[1][1])
            for paso, (titulo, texto) in intermedios:
                with st.expander(titulo):
                    mostrar(paso, texto)

        st.info("Que ambos modelos coincidan no garantiza que el resultado sea correcto. Verificá los cálculos en Excel antes de usarlos.")
        c1, c2 = st.columns(2)
        c1.download_button("Descargar Markdown", res["md"], file_name="resultado.md", use_container_width=True)
        c2.download_button("Descargar Word", archivos.a_docx(f"Asistente Dual · {res['modo']}", res["secciones"]),
                           file_name="resultado.docx", use_container_width=True)

with tab_historial:
    for c in hist.consultas():
        with st.expander(f"{c['fecha']} · {c['modo']} · USD {c['costo_usd']:.3f} · {c['pregunta'][:80]}"):
            st.markdown(c["resultado_md"])
