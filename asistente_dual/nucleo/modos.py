"""Modos de trabajo. Cada modo recibe `llamar(proveedor, rol, mensaje) -> Respuesta`,
donde proveedor es "claude" o "gpt" y rol es una instrucción adicional al perfil base.
Devuelve una lista de Paso; el último paso es el resultado final (salvo en Comparar)."""
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass

from .proveedores import Respuesta

ROL_AUTOR = "Elaborá un análisis completo y riguroso de la consulta."

ROL_REVISOR = """Actuás como un Gerente Financiero exigente que revisa el trabajo de un analista antes de presentarlo al Directorio.
Tu tarea es encontrar problemas, no validar. Revisá:
1. Supuestos débiles, no documentados o poco realistas.
2. Errores de cálculo o de fórmula: recalculá todo lo que puedas y mostrá la diferencia.
3. Confusiones entre resultado contable, rentabilidad, flujo de caja, liquidez y creación de valor.
4. Riesgos omitidos y sensibilidades no analizadas.
5. Conclusiones o recomendaciones que no se desprenden del análisis.
Clasificá cada observación como CRÍTICA, IMPORTANTE o MENOR, con su fundamento. No elogies lo que está bien.
Si no encontrás observaciones críticas, decilo explícitamente."""

ROL_CONSOLIDADOR = """Recibiste observaciones de un revisor sobre tu análisis. Evaluá cada una con criterio propio:
aceptala si es correcta y rechazala con fundamento si no lo es (el revisor también puede equivocarse).
Entregá:
1. La versión final corregida del análisis, completa.
2. Al final, una tabla de observaciones: observación | aceptada/rechazada | motivo."""

ROL_DEFENSOR = """Construí el mejor caso posible A FAVOR de la propuesta o decisión planteada:
beneficios, creación de valor, condiciones bajo las cuales conviene. Sé riguroso; no inventes datos."""

ROL_CRITICO = """Actuás como abogado del diablo. Construí el mejor caso posible EN CONTRA de la propuesta:
riesgos, supuestos frágiles, costos ocultos, impacto en caja y capital de trabajo, alternativas mejores
y escenarios en los que la decisión destruye valor. Respondé también a los argumentos a favor que recibís."""

ROL_SINTESIS = """Recibiste los argumentos a favor y en contra de una propuesta. Redactá una recomendación equilibrada para la Gerencia:
qué está en juego, qué argumentos pesan más y por qué, la decisión recomendada (aprobar, aprobar con condiciones o rechazar),
las condiciones o mitigantes necesarios y qué información falta para decidir con más certeza."""

ROL_JUEZ = """Recibiste dos respuestas independientes a la misma consulta (Respuesta 1 y Respuesta 2).
1. Señalá coincidencias y diferencias relevantes. Recordá que una coincidencia NO prueba que algo sea correcto.
2. Identificá errores, omisiones o supuestos débiles en cada una.
3. Indicá cuál es más sólida en cada aspecto y por qué.
4. Entregá una respuesta integrada final.
5. Cerrá con una lista de lo que el usuario debe verificar por su cuenta antes de usar el resultado."""

MODOS = {
    "Comparar": "La misma consulta a ambos modelos; respuestas lado a lado.",
    "Revisión cruzada": "Un modelo elabora, el otro critica como Gerente Financiero exigente y el primero consolida la versión final.",
    "Abogado del diablo": "Un modelo defiende la propuesta, el otro la ataca y el primero sintetiza una recomendación.",
    "Juez": "Ambos responden por separado y un tercer paso compara, detecta diferencias y entrega una respuesta integrada.",
}


@dataclass
class Paso:
    titulo: str
    respuesta: Respuesta


def otro(proveedor):
    return "gpt" if proveedor == "claude" else "claude"


def bloque(etiqueta, texto):
    return f"<{etiqueta}>\n{texto}\n</{etiqueta}>"


def comparar(llamar, consulta, **_):
    with ThreadPoolExecutor(max_workers=2) as ex:
        fc = ex.submit(llamar, "claude", ROL_AUTOR, consulta)
        fg = ex.submit(llamar, "gpt", ROL_AUTOR, consulta)
        return [Paso("Respuesta de Claude", fc.result()), Paso("Respuesta de GPT", fg.result())]


def revision_cruzada(llamar, consulta, principal="claude"):
    revisor = otro(principal)
    borrador = llamar(principal, ROL_AUTOR, consulta)
    revision = llamar(revisor, ROL_REVISOR, bloque("consulta", consulta) + "\n\n" + bloque("analisis_a_revisar", borrador.texto))
    final = llamar(principal, ROL_CONSOLIDADOR,
                   bloque("consulta", consulta) + "\n\n" + bloque("tu_analisis", borrador.texto)
                   + "\n\n" + bloque("observaciones_del_revisor", revision.texto))
    return [Paso("1. Borrador", borrador), Paso("2. Revisión crítica", revision), Paso("3. Versión final consolidada", final)]


def abogado_del_diablo(llamar, consulta, principal="claude"):
    critico = otro(principal)
    a_favor = llamar(principal, ROL_DEFENSOR, consulta)
    en_contra = llamar(critico, ROL_CRITICO, bloque("consulta", consulta) + "\n\n" + bloque("argumentos_a_favor", a_favor.texto))
    sintesis = llamar(principal, ROL_SINTESIS,
                      bloque("consulta", consulta) + "\n\n" + bloque("argumentos_a_favor", a_favor.texto)
                      + "\n\n" + bloque("argumentos_en_contra", en_contra.texto))
    return [Paso("1. Argumentos a favor", a_favor), Paso("2. Argumentos en contra", en_contra), Paso("3. Recomendación", sintesis)]


def juez(llamar, consulta, principal="claude"):
    r1, r2 = (p.respuesta for p in comparar(llamar, consulta))
    veredicto = llamar(principal, ROL_JUEZ,
                       bloque("consulta", consulta) + "\n\n" + bloque("respuesta_1", r1.texto) + "\n\n" + bloque("respuesta_2", r2.texto))
    return [Paso("1. Respuesta de Claude", r1), Paso("2. Respuesta de GPT", r2), Paso("3. Comparación y respuesta integrada", veredicto)]


EJECUTAR = {"Comparar": comparar, "Revisión cruzada": revision_cruzada, "Abogado del diablo": abogado_del_diablo, "Juez": juez}


def armar_consulta(pregunta, documentos):
    """Une la pregunta con el contenido de los archivos cargados. documentos: lista de (nombre, texto)."""
    partes = [pregunta.strip()]
    for nombre, texto in documentos:
        partes.append(f'<documento nombre="{nombre}">\n{texto}\n</documento>')
    return "\n\n".join(partes)
