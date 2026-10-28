# -*- coding: utf-8 -*-
"""
Plantilla común de modelo financiero para maquilas (v2 simplificada).

Un archivo independiente por negocio, misma estructura:
  Inicio · Supuestos · Productos · Costos · Inversion · Reales · Calculo · Resultados · Control
  más las hojas de cuadros (Cuadros, C01 … RRHH) con la misma estructura que el plan de negocio.

- Entradas solo en Inicio, Supuestos, Productos, Costos, Inversion y Reales (celdas amarillas, texto azul).
- Cada dato se carga una vez: la columna «Vigente» hereda el «Presupuesto original» (gris) hasta que se sobrescribe.
- Calculo es un motor mensual único que se ejecuta dos veces con las mismas fórmulas:
  P = presupuesto original (sin reales) y V = vigente (reales hasta el mes de corte + proyección con escenario).
- Resultados y Control se calculan automáticamente.

Uso:
  python3 generar_plantilla.py ecostar  salida.xlsx    # carga de Ecostar
  python3 generar_plantilla.py vacia    salida.xlsx    # plantilla en blanco
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL, column_index_from_string
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.comments import Comment

MODO = sys.argv[1] if len(sys.argv) > 1 else "ecostar"
OUT = sys.argv[2] if len(sys.argv) > 2 else "Modelo_Maquila_v2.xlsx"
if MODO == "ecostar":
    from datos_ecostar import DATOS as D
else:
    D = {"negocio": None, "socios": None, "etapa": None, "producto_desc": None, "fecha_mes1": None, "corte": None,
         "inicio_real": None, "escenario": "Vigente", "params": {"Anios": (10, "")}, "anual": {}, "productos": [],
         "volumen": [], "personal": [], "fijos": [], "capex": [], "capex_nota": "", "prestamos": [], "aportes": [],
         "publicado": {}, "pendientes": []}

# ---------------------------------------------------------------------------- dimensiones
N = 144                     # meses de la rejilla (12 años: implantación + 10 años de operación)
FCOL = 6                    # columna del mes 1 en Reales y Calculo (F)
NP, NPER, NFIJ, NCAP, NAP, NL = 25, 30, 25, 20, 10, 3   # productos individuales, puestos, costos fijos, CAPEX, aportes, préstamos
NY = 10                     # años de supuestos
NYR = 12                    # años del resumen


def mc(m):
    return CL(FCOL + m - 1)


M1, MN = mc(1), mc(N)

# ---------------------------------------------------------------------------- estilos
FONT = "Arial"
F = Font(name=FONT, size=9)
F_IN = Font(name=FONT, size=9, color="0000FF")
F_INH = Font(name=FONT, size=9, color="808080", italic=True)
F_LINK = Font(name=FONT, size=9, color="008000")
F_B = Font(name=FONT, size=9, bold=True)
F_H = Font(name=FONT, size=9, bold=True, color="FFFFFF")
F_T = Font(name=FONT, size=13, bold=True, color="1F3864")
F_S = Font(name=FONT, size=9, italic=True, color="595959")
F_SEC = Font(name=FONT, size=10, bold=True, color="1F3864")
F_AL = Font(name=FONT, size=9, bold=True, color="C00000")
FL_IN = PatternFill("solid", fgColor="FFF2CC")
FL_INH = PatternFill("solid", fgColor="F2F2F2")
FL_H = PatternFill("solid", fgColor="1F3864")
FL_SEC = PatternFill("solid", fgColor="D9E1F2")
FL_TOT = PatternFill("solid", fgColor="EDEDED")
FL_P = PatternFill("solid", fgColor="FCE4D6")     # presupuesto original
FL_V = PatternFill("solid", fgColor="E2EFDA")     # vigente
FL_OK = PatternFill("solid", fgColor="C6EFCE")
FL_BAD = PatternFill("solid", fgColor="FFC7CE")
NF_USD = '#,##0;(#,##0);"-"'
NF_USD2 = '#,##0.00;(#,##0.00);"-"'
NF_USD4 = '#,##0.0000;(#,##0.0000);"-"'
NF_PCT = '0.0%;(0.0%);"-"'
NF_PCT2 = '0.00%;(0.00%);"-"'
NF_INT = '#,##0;(#,##0);"-"'
NF_N = '0'
NF_X = '0.00"x";(0.00"x");"-"'
NF_DATE = 'dd/mm/yyyy'
NF_MES = 'mmm-yy'
WRAP = Alignment(wrap_text=True, vertical="top")
CEN = Alignment(horizontal="center", vertical="center", wrap_text=True)

wb = Workbook()
wb.remove(wb.active)
SHEETS = ["Inicio", "Guia_Carga", "Supuestos", "Productos", "Costos", "Inversion", "Detalle_24m", "Reales", "Calculo", "Proyeccion_24m", "Resultados", "Control",
          "Cuadros", "C01 Inversión", "C07 Producción", "C08 Ingresos", "C09 Materiales", "C10a Costos", "C10b Costo unitario", "C11 Rubros",
          "C13 Deuda e IVA", "C15 Resultados", "C16 Capital trabajo", "C17 Equilibrio", "C18 Fondos", "C19 VAN-TIR", "C20 Depreciación", "RRHH"]
WS = {s: wb.create_sheet(s) for s in SHEETS}
CUR = [None]


# Criterio: no usar decimales dentro de criterios de texto (COUNTIF(..., ">0.5")): Excel los interpreta según la
# configuración regional (coma decimal en es-PY) y devuelve 0. Usar comparaciones numéricas (SUMPRODUCT(--(rango>0.5))).


def nm(name, sheet, ref):
    """Nombre de libro: ref puede ser celda (D5) o rango (D5:M5)."""
    parts = ref.split(":")
    absref = ":".join("$" + "".join(c for c in p if c.isalpha()) + "$" + "".join(c for c in p if c.isdigit()) for p in parts)
    wb.defined_names[name] = DefinedName(name, attr_text=f"'{sheet}'!{absref}")


def put(ws, r, c, v, font=None, fill=None, fmt=None, al=None, note=None):
    cell = ws.cell(row=r, column=c, value=v)
    cell.font = font or F
    if fill:
        cell.fill = fill
    if fmt:
        cell.number_format = fmt
    if al:
        cell.alignment = al
    if note:
        cell.comment = Comment(note, "Modelo")
    return cell


def inp(ws, r, c, v=None, fmt=None):
    return put(ws, r, c, v, font=F_IN, fill=FL_IN, fmt=fmt)


def inh(ws, r, c, formula, fmt=None):
    """Celda vigente que hereda del presupuesto: gris hasta que se sobrescribe."""
    return put(ws, r, c, formula, font=F_INH, fill=FL_INH, fmt=fmt)


def title(ws, t, sub=None):
    put(ws, 1, 1, t, font=F_T)
    if sub:
        put(ws, 2, 1, sub, font=F_S)


def section(ws, r, t, ncol=16):
    for c in range(1, ncol + 1):
        ws.cell(row=r, column=c).fill = FL_SEC
    put(ws, r, 1, t, font=F_SEC, fill=FL_SEC)


def header(ws, r, labels, c0=1, h=None):
    for i, l in enumerate(labels):
        put(ws, r, c0 + i, l, font=F_H, fill=FL_H, al=CEN)
    if h:
        ws.row_dimensions[r].height = h


def dv(ws, rng, opts):
    d = DataValidation(type="list", formula1='"' + ",".join(opts) + '"', allow_blank=True)
    ws.add_data_validation(d)
    d.add(rng)


def widths(ws, w):
    for k, v in w.items():
        ws.column_dimensions[k].width = v


def P(key, i=None):
    """Valor de un parámetro del diccionario de datos."""
    v = D.get("params", {}).get(key)
    return v[0] if v else None


# ============================================================================ INICIO
def build_inicio():
    ws = WS["Inicio"]
    widths(ws, {"A": 3, "B": 46, "C": 34, "D": 90})
    put(ws, 1, 1, '="Modelo financiero de maquila — "&IF(n_Negocio="","(completar nombre)",n_Negocio)', font=F_T)
    put(ws, 2, 1, "Plantilla común v2: un archivo por negocio, misma estructura. Presupuesto original, ejecución real y proyección vigente separados.", font=F_S)
    r = 4
    section(ws, r, "1. Configuración del negocio (completar)", 4); r += 1
    header(ws, r, ["", "Dato", "Valor", "Instrucción"], 1); r += 1
    rows = [
        ("n_Negocio", "Nombre del negocio", D.get("negocio"), None, "Texto libre."),
        ("n_Socios", "Socios / empresas", D.get("socios"), None, "Texto libre."),
        ("n_Producto", "Actividad / productos", D.get("producto_desc"), None, "Texto libre."),
        ("n_Etapa", "Etapa del negocio", D.get("etapa"), None, "Implantación (aún sin operar) u Operación."),
        ("n_Fecha1", "Fecha del mes 1 del modelo", D.get("fecha_mes1"), NF_DATE,
         "Primer mes de implantación u operación. Se recomienda que sea enero para que los años coincidan con años calendario. Vacío = meses relativos."),
        ("n_Corte", "Último mes cerrado con datos reales (n° de mes)", D.get("corte"), NF_N,
         "Los meses hasta este número usan solo la hoja Reales y no cambian con los escenarios. Vacío = sin datos reales."),
        ("n_InicioReal", "Mes real de inicio de operación (n° de mes)", D.get("inicio_real"), NF_N,
         "Cargar cuando la operación ya comenzó. Prevalece sobre el supuesto vigente y el atraso del escenario."),
        ("n_Escenario", "Escenario para la proyección", D.get("escenario") or "Vigente", None,
         "Vigente (sin ajustes), Conservador u Optimista. Solo afecta los meses posteriores al corte."),
        ("n_Ventana", "Meses de proyección detallada (ventana móvil)", D.get("ventana", 24), NF_N,
         "Meses siguientes al último cierre que usan el detalle mensual de Detalle_24m (recomendado 24). Después, supuestos anuales."),
        ("n_Version", "Versión del plan aprobado (presupuesto original)", D.get("version_plan"), None, "Ej.: Plan Oct-2025; Revisión 2027."),
        ("n_FechaRev", "Fecha de aprobación de la última revisión del plan", D.get("fecha_revision"), NF_DATE,
         "Con la frecuencia, define la próxima revisión del plan a 10 años."),
        ("n_Frecuencia", "Frecuencia de revisión del plan", D.get("frecuencia", "Anual"), None,
         "Anual (año 1 firme + año 2 actualizado) o Bienal. El detalle de 24 meses se actualiza en cada cierre."),
    ]
    for name, lab, val, fmt, ins in rows:
        put(ws, r, 2, lab)
        inp(ws, r, 3, val, fmt)
        put(ws, r, 4, ins, font=F_S)
        nm(name, "Inicio", f"C{r}")
        if name == "n_Etapa":
            dv(ws, f"C{r}", ["Implantación", "Operación"])
        if name == "n_Escenario":
            dv(ws, f"C{r}", ["Vigente", "Conservador", "Optimista"])
        if name == "n_Frecuencia":
            dv(ws, f"C{r}", ["Anual", "Bienal"])
        r += 1
    put(ws, r, 2, "Moneda de presentación")
    put(ws, r, 3, "USD (costos en Gs convertidos con el TC de cada año)")
    r += 2
    section(ws, r, "2. Estado del modelo (automático)", 4); r += 1
    est = [
        ("Fecha del último mes cerrado", '=IF(Corte=0,"Sin datos reales",IF(n_Fecha1="","Mes "&Corte,"Mes "&Corte&" ("&TEXT(MONTH(EOMONTH(n_Fecha1,Corte-1)),"00")&"/"&YEAR(EOMONTH(n_Fecha1,Corte-1))&")"))'),
        ("Inicio de operación (presupuesto / vigente)", '="Mes "&P_Inicio&" / mes "&V_Inicio'),
        ("Fin del horizonte evaluado (fijo)", '="Mes "&Fin&" ("&P_Anios&" años de operación desde el inicio presupuestado)"'),
        ("Supuestos vigentes modificados respecto del presupuesto", "=c_Modificados"),
        ("Controles con alerta (hoja Control)", "=c_Alertas"),
        ("Brecha de caja máxima vigente (USD)", "=r_BrechaV"),
        ("Pasos de la Guía de carga con alerta (⚠)", "=g_Pendientes"),
        ("Ventana de proyección detallada", '=IF(N(n_Ventana)=0,"Sin detalle mensual","Meses "&(Corte+1)&" a "&MIN(144,Corte+n_Ventana))'),
        ("Próxima revisión del plan", '=IF(ISNUMBER(n_FechaRev),DAY(EDATE(n_FechaRev,IF(n_Frecuencia="Bienal",24,12)))&"/"&MONTH(EDATE(n_FechaRev,IF(n_Frecuencia="Bienal",24,12)))&"/"&YEAR(EDATE(n_FechaRev,IF(n_Frecuencia="Bienal",24,12))),"Sin fecha de revisión")'),
    ]
    for lab, f in est:
        put(ws, r, 2, lab)
        put(ws, r, 3, f, font=F_B, fmt=NF_USD)
        r += 1
    r += 1
    section(ws, r, "3. Cómo cargar y actualizar", 4); r += 1
    pasos = [
        ("Paso a paso", "Seguir la hoja Guia_Carga: cada paso tiene vínculo a la celda, fuente, criterio y estado automático."),
        ("Proyección 24 meses", "En Detalle_24m cargar mes a mes lo asegurado para los próximos meses (volúmenes, precios, costos, personal, fijos, TC). "
                                "Celda vacía = supuesto anual. Con cada cierre la ventana avanza un mes: completar el nuevo último mes."),
        ("Revisión del plan", "Anual o bienal: el plan aprobado (presupuesto original, 10 años) se actualiza con los supuestos vigentes aprobados; guardar la versión anterior."),
        ("Alta de un negocio", "Copiar el archivo vacío, completar la sección 1 y luego las hojas en este orden: Supuestos → Productos → Costos → Inversion."),
        ("Presupuesto original", "Cargar en las columnas «Presupuesto original» los datos del plan aprobado. No se modifican después: son la referencia de comparación."),
        ("Proyección vigente", "Las columnas «Vigente» heredan el presupuesto (gris). Sobrescribir solo lo que cambió; la celda pasa a azul/amarillo y Control lo cuenta."),
        ("Cierre mensual", "1) Cargar en Reales el mes cerrado (unidades, ventas, costos, CAPEX, financiamiento). 2) Cargar los saldos al cierre en el mes de corte. "
                           "3) Actualizar «Último mes cerrado» arriba. 4) Revisar Control. 5) Ajustar la proyección vigente si cambió."),
        ("Inversión", "En Inversion registrar por ítem lo ejecutado, pagado, comprometido y por contratar. El costo final estimado se calcula solo."),
        ("Escenarios", "Elegir el escenario arriba. Los ajustes están en Supuestos sección C y solo afectan la proyección (nunca los meses reales)."),
        ("Lectura", "Resultados compara presupuesto original vs vigente (real + proyección). Control muestra los faltantes y las inconsistencias."),
    ]
    for a, b in pasos:
        put(ws, r, 2, a, font=F_B)
        put(ws, r, 3, b, al=WRAP)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        ws.row_dimensions[r].height = 26
        r += 1
    r += 1
    section(ws, r, "4. Convenciones", 4); r += 1
    conv = [(FL_IN, F_IN, "Celda para completar (dato de entrada)."),
            (FL_INH, F_INH, "Vigente heredado del presupuesto (fórmula). Sobrescribir si el dato cambió."),
            (None, F, "Fórmula: no editar."),
            (FL_P, F, "Presupuesto original (Resultados)."),
            (FL_V, F, "Vigente: real + proyección (Resultados).")]
    for fill, font, txt in conv:
        put(ws, r, 2, "Ejemplo 1.234", font=font, fill=fill)
        put(ws, r, 3, txt)
        r += 1
    put(ws, r, 2, "Porcentajes como fracción (0,05 = 5%). Montos en USD salvo que la columna Moneda indique PYG.", font=F_S)
    r += 2
    section(ws, r, "5. Hojas", 4); r += 1
    hojas = [("Guia_Carga", "Guía de carga paso a paso con estado automático (alta, presupuesto, vigente, cierre mensual y revisión)."),
             ("Supuestos", "Parámetros generales, supuestos por año, escenarios y referencia del plan publicado."),
             ("Productos", f"Hasta {NP} productos: precio, costos unitarios, capacidad y volumen anual (presupuesto y vigente)."),
             ("Costos", "Personal por puesto (dotación por año) y costos fijos (moneda, crecimiento, inicio)."),
             ("Inversion", "CAPEX por ítem (presupuesto, ejecutado, comprometido, por contratar), préstamos y aportes."),
             ("Detalle_24m", "Proyección mes a mes de la ventana móvil (por defecto 24 meses): volúmenes, precios, costos, personal, fijos y TC."),
             ("Reales", "Carga mensual de la ejecución y saldos al cierre (formato fijo: un dato por celda)."),
             ("Calculo", "Motor mensual automático (no editar). Bloque P = presupuesto; bloque V = vigente."),
             ("Proyeccion_24m", "Vista mensual de los próximos 24 meses (vigente vs presupuesto) con totales por año de ventana."),
             ("Resultados", "Indicadores, resúmenes anuales (10 años), desvíos y gráficos."),
             ("Control", "Validaciones y datos faltantes."),
             ("Cuadros", "Índice de los cuadros financieros con la estructura del plan de negocio, versión que muestran y resumen modelo vs plan."),
             ("C01 … RRHH", "Un cuadro por hoja, como el plan (inversión, producción, ingresos, materiales, costos, deuda, resultados, capital de trabajo, "
                            "equilibrio, fondos, VAN-TIR, depreciación y personal): modelo, plan publicado y diferencia.")]
    for a, b in hojas:
        put(ws, r, 2, a, font=F_B)
        put(ws, r, 3, b)
        r += 1
    r += 1
    section(ws, r, "6. Datos pendientes de este negocio (editable)", 4); r += 1
    pend = D.get("pendientes", [])
    for i in range(10):
        put(ws, r, 2, f"{i + 1}.")
        inp(ws, r, 3, pend[i] if i < len(pend) else None)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
        r += 1
    ws.freeze_panes = "A4"


# ============================================================================ SUPUESTOS
PARAMS = [
    # clave, etiqueta, unidad, formato, aplica escenario
    ("Inicio", "Mes de inicio de la operación (n° de mes del modelo)", "mes", NF_N),
    ("Anios", "Años de operación evaluados", "años", NF_N),
    ("Merma", "Merma entre producción y venta", "%", NF_PCT2),
    ("Cargas", "Cargas sociales sobre salarios (aguinaldo, IPS, vacaciones)", "%", NF_PCT2),
    ("Alim", "Alimentación del personal", "Gs/persona/mes", NF_INT),
    ("PreDir", "Personal directo activo antes del inicio de operación", "% dotación año 1", NF_PCT),
    ("PreInd", "Personal indirecto y administrativo activo antes del inicio", "% dotación año 1", NF_PCT),
    ("Desp", "Despacho y gastos de importación (s/ materia prima + insumos)", "%", NF_PCT2),
    ("Pkg", "Packaging (s/ materiales, otros variables y personal directo)", "%", NF_PCT2),
    ("ImpP", "Imprevistos de producción (s/ costos de producción)", "%", NF_PCT2),
    ("ImpA", "Imprevistos de administración (s/ gastos marcados como base)", "%", NF_PCT2),
    ("Trib", "Tributo de maquila (s/ ventas)", "%", NF_PCT2),
    ("SegV", "Seguro de caución y otros cargos (s/ ventas)", "%", NF_PCT2),
    ("ComV", "Comisiones y logística de venta (s/ ventas)", "%", NF_PCT2),
    ("IRE", "Impuesto a la renta (s/ resultado)", "%", NF_PCT2),
    ("DSO", "Días de cobro a clientes", "días", NF_N),
    ("DInv", "Días de inventario de materiales", "días", NF_N),
    ("DPT", "Días de producto terminado", "días", NF_N),
    ("DPOmp", "Días de pago a proveedores de materiales", "días", NF_N),
    ("DPOot", "Días de pago a otros proveedores", "días", NF_N),
    ("CajaMin", "Caja mínima operativa", "días de ventas", NF_N),
    ("IVArec", "IVA recuperable de los costos fijos gravados", "%", NF_PCT),
    ("IVAmeses", "Meses para recuperar el IVA", "meses", NF_N),
    ("IVAint", "IVA sobre intereses", "%", NF_PCT),
    ("WACC", "Tasa de descuento del proyecto (WACC, anual)", "%", NF_PCT2),
    ("Ke", "Costo del capital propio (Ke, anual)", "%", NF_PCT2),
    ("RealCT", "Recuperación del capital de trabajo al final del horizonte", "%", NF_PCT),
    ("RealAF", "Valor de realización del activo fijo al final (s/ valor contable)", "%", NF_PCT),
    ("MesDiv", "Mes de pago de dividendos del año siguiente (1-12)", "mes", NF_N),
]
ANUAL = [("TC", "Tipo de cambio (Gs/USD) — por año del modelo", "Gs/USD", NF_INT),
         ("dPrecio", "Variación de precios de venta (año de operación)", "%", NF_PCT),
         ("dMP", "Variación del costo de materia prima", "%", NF_PCT),
         ("dOV", "Variación de insumos y otros costos variables", "%", NF_PCT),
         ("dSal", "Ajuste salarial", "%", NF_PCT),
         ("Payout", "Dividendos (% del resultado del año)", "%", NF_PCT)]
IDX = [("Precio", "dPrecio"), ("MP", "dMP"), ("OV", "dOV"), ("Sal", "dSal")]
ESC = [("dVol", "Ajuste de volumen", "%", NF_PCT, [0, -0.10, 0.05]),
       ("dPrecio", "Ajuste de precio", "%", NF_PCT, [0, -0.05, 0.02]),
       ("dMat", "Ajuste de costo de materiales y variables", "%", NF_PCT, [0, 0.10, -0.03]),
       ("dTC", "Ajuste del tipo de cambio (+ = más Gs por USD)", "%", NF_PCT, [0, 0, 0]),
       ("Atraso", "Atraso adicional del inicio de operación", "meses", NF_N, [0, 3, 0]),
       ("dCapex", "Sobrecosto sobre CAPEX pendiente", "%", NF_PCT, [0, 0.10, 0]),
       ("DSOesc", "Días de cobro (vacío = vigente)", "días", NF_N, [None, 150, 90]),
       ("DInvEsc", "Días de inventario (vacío = vigente)", "días", NF_N, [None, None, 90])]
SUPR = {}


def build_supuestos():
    ws = WS["Supuestos"]
    title(ws, "Supuestos — presupuesto original y vigente",
          "Presupuesto original = plan aprobado (no se modifica). Vigente = hereda el presupuesto; sobrescribir solo lo que cambió.")
    widths(ws, {"A": 7, "B": 58, "C": 15, "D": 14, "E": 14, "F": 14, "G": 14, "H": 13, "I": 13, "J": 13, "K": 13, "L": 13, "M": 13, "N": 60})
    r = 4
    section(ws, r, "A. Parámetros generales", 14); r += 1
    header(ws, r, ["Cód.", "Parámetro", "Unidad", "Presupuesto original", "Vigente", "Nota / fuente"], 1, 30)
    ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=14)
    r += 1
    for i, (k, lab, unit, fmt) in enumerate(PARAMS):
        SUPR[k] = r
        put(ws, r, 1, f"G{i + 1:02d}")
        put(ws, r, 2, lab)
        put(ws, r, 3, unit)
        inp(ws, r, 4, P(k), fmt)
        note = D.get("params", {}).get(k, (None, ""))[1] or ""
        if k == "Anios":
            put(ws, r, 5, f"=D{r}", fmt=fmt)
            note = (note + " " if note else "") + "Horizonte común a ambas versiones: se define solo en el presupuesto."
        else:
            inh(ws, r, 5, f"=D{r}", fmt)
        put(ws, r, 6, note, font=F_S)
        nm(f"o_{k}", "Supuestos", f"D{r}")
        nm(f"vg_{k}", "Supuestos", f"E{r}")
        r += 1
    SUPR["p_last"] = r - 1
    r += 1
    section(ws, r, "B. Supuestos por año (año 1 = primer año de operación; el TC se aplica por año del modelo)", 14); r += 1
    header(ws, r, ["Cód.", "Concepto", "Unidad"] + [f"Año {y}" for y in range(1, NY + 1)] + ["Nota"], 1)
    r += 1
    for ver, lab_ver, fill in [("O", "PRESUPUESTO ORIGINAL", FL_P), ("V", "VIGENTE (hereda el presupuesto)", FL_V)]:
        put(ws, r, 2, lab_ver, font=F_B, fill=fill)
        r += 1
        for k, lab, unit, fmt in ANUAL:
            SUPR[f"{k}_{ver}"] = r
            put(ws, r, 1, k)
            put(ws, r, 2, lab)
            put(ws, r, 3, unit)
            vals = D.get("anual", {}).get(k, [None] * NY)
            for y in range(1, NY + 1):
                c = 3 + y
                if ver == "O":
                    inp(ws, r, c, vals[y - 1] if vals else None, fmt)
                else:
                    inh(ws, r, c, f"={CL(c)}{SUPR[k + '_O']}", fmt)
            nm(f"{k}_{ver}", "Supuestos", f"D{r}:M{r}")
            r += 1
        for k, src in IDX:
            SUPR[f"Idx{k}_{ver}"] = r
            put(ws, r, 2, f"Índice acumulado de {dict(Precio='precios', MP='materia prima', OV='otros variables', Sal='salarios')[k]} (cálculo)", font=F_S)
            for y in range(1, NY + 1):
                c = 3 + y
                f = "=1" if y == 1 else f"={CL(c - 1)}{r}*(1+N({CL(c)}{SUPR[src + '_' + ver]}))"
                put(ws, r, c, f, fmt=NF_USD4)
            nm(f"Idx{k}_{ver}", "Supuestos", f"D{r}:M{r}")
            r += 1
        r += 1
    section(ws, r, "C. Escenarios (se aplican solo a la proyección vigente, después del último mes cerrado)", 14); r += 1
    header(ws, r, ["Cód.", "Ajuste", "Unidad", "Vigente", "Conservador", "Optimista", "Aplicado", "Nota"], 1, 30)
    ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=14)
    r += 1
    SUPR["esc_first"] = r
    for i, (k, lab, unit, fmt, vals) in enumerate(ESC):
        put(ws, r, 1, f"E{i + 1}")
        put(ws, r, 2, lab)
        put(ws, r, 3, unit)
        put(ws, r, 4, vals[0], fmt=fmt, fill=FL_TOT)
        inp(ws, r, 5, vals[1], fmt)
        inp(ws, r, 6, vals[2], fmt)
        put(ws, r, 7, f'=IF(INDEX(D{r}:F{r},1,c_EscIdx)="","",INDEX(D{r}:F{r},1,c_EscIdx))', fmt=fmt, font=F_B)
        nm(f"e_{k}", "Supuestos", f"G{r}")
        r += 1
    put(ws, r, 2, "Índice del escenario activo")
    put(ws, r, 4, '=IF(n_Escenario="Conservador",2,IF(n_Escenario="Optimista",3,1))', fmt=NF_N)
    nm("c_EscIdx", "Supuestos", f"D{r}")
    put(ws, r, 8, "Conservador y Optimista son ajustes ILUSTRATIVOS editables, no premisas aprobadas.", font=F_S)
    r += 2
    section(ws, r, "D. Referencia del plan publicado (opcional; solo para comparar)", 14); r += 1
    header(ws, r, ["Cód.", "Concepto", "Unidad"] + [f"Año {y}" for y in range(1, NY + 1)] + ["Nota"], 1)
    r += 1
    pub = D.get("publicado", {})
    for k, lab in [("Ventas", "Ventas publicadas"), ("EBITDA", "EBITDA publicado"), ("Utilidad", "Resultado neto publicado")]:
        SUPR[f"pub_{k}"] = r
        put(ws, r, 1, k)
        put(ws, r, 2, lab)
        put(ws, r, 3, "USD")
        vals = pub.get(k, [None] * NY)
        for y in range(1, NY + 1):
            inp(ws, r, 3 + y, vals[y - 1] if vals else None, NF_USD)
        nm(f"pub_{k}", "Supuestos", f"D{r}:M{r}")
        r += 1
    for k, lab, fmt in [("VAN", "VAN publicado", NF_USD), ("TIR", "TIR publicada", NF_PCT2)]:
        put(ws, r, 1, k)
        put(ws, r, 2, lab)
        inp(ws, r, 4, pub.get(k), fmt)
        nm(f"pub_{k}", "Supuestos", f"D{r}")
        r += 1
    put(ws, r, 2, pub.get("nota", ""), font=F_S)
    r += 2
    section(ws, r, "E. Parámetros aplicados por versión (cálculo automático; no editar)", 14); r += 1
    header(ws, r, ["", "Parámetro", "", "Presupuesto (P)", "Vigente (V)"], 1)
    r += 1
    for k, lab, unit, fmt in PARAMS:
        put(ws, r, 2, lab, font=F_S)
        put(ws, r, 4, "=MAX(1,N(o_Inicio))" if k == "Inicio" else f"=N(o_{k})", fmt=fmt)
        if k == "DSO":
            fv = '=IF(e_DSOesc="",N(vg_DSO),e_DSOesc)'
        elif k == "DInv":
            fv = '=IF(e_DInvEsc="",N(vg_DInv),e_DInvEsc)'
        elif k == "Inicio":
            fv = '=MAX(1,IF(ISNUMBER(n_InicioReal),n_InicioReal,IF(N(vg_Inicio)<=Corte,N(vg_Inicio),N(vg_Inicio)+ROUND(N(e_Atraso),0))))'
        else:
            fv = f"=N(vg_{k})"
        put(ws, r, 5, fv, fmt=fmt)
        nm(f"P_{k}", "Supuestos", f"D{r}")
        nm(f"V_{k}", "Supuestos", f"E{r}")
        r += 1
    for k, lab in [("dVol", "Ajuste de volumen"), ("dPrecio", "Ajuste de precio"), ("dMat", "Ajuste de materiales"), ("dTC", "Ajuste de TC"),
                   ("dCapex", "Sobrecosto CAPEX pendiente")]:
        put(ws, r, 2, lab, font=F_S)
        put(ws, r, 4, 0, fmt=NF_PCT)
        put(ws, r, 5, f"=N(e_{k})", fmt=NF_PCT)
        nm(f"P_{k}", "Supuestos", f"D{r}")
        nm(f"V_{k}", "Supuestos", f"E{r}")
        r += 1
    for name, lab, f, fmt in [("Corte", "Último mes cerrado (0 = sin reales)", "=IF(ISNUMBER(n_Corte),n_Corte,0)", NF_N),
                              ("Fin", "Último mes del horizonte (fijo, según presupuesto)", "=P_Inicio+12*P_Anios-1", NF_N),
                              ("DiasMes", "Días por mes", "=365/12", NF_USD2)]:
        put(ws, r, 2, lab, font=F_S)
        put(ws, r, 4, f, fmt=fmt)
        nm(name, "Supuestos", f"D{r}")
        r += 1
    ws.freeze_panes = "D6"


# ============================================================================ PRODUCTOS
PR = {}


def build_productos():
    ws = WS["Productos"]
    title(ws, "Productos — precios, costos unitarios, capacidad y volúmenes (USD por unidad de venta)",
          f"Hasta {NP} productos. Las columnas «vigente» heredan el presupuesto; sobrescribir lo que cambió.")
    widths(ws, {"A": 4, "B": 34, "C": 8})
    for c in range(4, 16):
        ws.column_dimensions[CL(c)].width = 12.5
    ws.column_dimensions["O"].width = 50
    r = 4
    section(ws, r, "A. Datos por producto", 15); r += 1
    header(ws, r, ["#", "Producto", "Unidad", "Capacidad inicial (u/año)", "Capacidad ampliada (u/año)", "Desde año de operación",
                   "Precio año 1 — presupuesto", "Precio año 1 — vigente", "Materia prima /u — presupuesto", "Materia prima /u — vigente",
                   "Insumos /u — presupuesto", "Insumos /u — vigente", "Otros variables /u — presupuesto", "Otros variables /u — vigente", "Nota"], 1, 54)
    r += 1
    PR["first"] = r
    prods = D.get("productos", [])
    for i in range(NP):
        p = prods[i] if i < len(prods) else None
        put(ws, r, 1, i + 1)
        inp(ws, r, 2, p[0] if p else None)
        inp(ws, r, 3, p[1] if p else None)
        inp(ws, r, 4, p[2] if p else None, NF_INT)
        inp(ws, r, 5, p[3] if p else None, NF_INT)
        inp(ws, r, 6, p[4] if p else None, NF_N)
        for j, (co, cv) in enumerate([(7, 8), (9, 10), (11, 12), (13, 14)]):
            inp(ws, r, co, p[5 + j] if p else None, NF_USD4)
            inh(ws, r, cv, f"={CL(co)}{r}", NF_USD4)
        put(ws, r, 15, p[9] if p else None, font=F_S)
        r += 1
    PR["last"] = r - 1
    f, l = PR["first"], PR["last"]
    for name, col in [("Prod_Nombre", "B"), ("Cap_Ini", "D"), ("Cap_Amp", "E"), ("Cap_Anio", "F"), ("Precio_O", "G"), ("Precio_V", "H"),
                      ("MP_O", "I"), ("MP_V", "J"), ("INS_O", "K"), ("INS_V", "L"), ("OV_O", "M"), ("OV_V", "N")]:
        nm(name, "Productos", f"{col}{f}:{col}{l}")
    r += 1
    vols = D.get("volumen", [])
    for ver, lab, fill in [("O", "B. Ventas en unidades por año de operación — PRESUPUESTO ORIGINAL", FL_P),
                           ("V", "C. Ventas en unidades por año de operación — VIGENTE (hereda; sobrescribir lo que cambió)", FL_V)]:
        section(ws, r, lab, 15); r += 1
        header(ws, r, ["#", "Producto", "Unidad"] + [f"Año {y}" for y in range(1, NY + 1)] + ["Total"], 1)
        r += 1
        PR[f"vol_{ver}"] = r
        for i in range(NP):
            put(ws, r, 1, i + 1)
            put(ws, r, 2, f'=IF({CL(2)}{f + i}="","",{CL(2)}{f + i})', font=F_LINK)
            put(ws, r, 3, f'=IF({CL(3)}{f + i}="","",{CL(3)}{f + i})', font=F_LINK)
            for y in range(1, NY + 1):
                c = 3 + y
                if ver == "O":
                    v = vols[i][y - 1] if i < len(vols) else None
                    inp(ws, r, c, v, NF_INT)
                else:
                    inh(ws, r, c, f"={CL(c)}{PR['vol_O'] + i}", NF_INT)
            put(ws, r, 14, f"=SUM(D{r}:M{r})", fmt=NF_INT, font=F_B)
            r += 1
        nm(f"VolTot_{ver}", "Productos", f"N{PR['vol_' + ver]}:N{PR['vol_' + ver] + NP - 1}")
        put(ws, r, 2, "Total", font=F_B)
        for y in range(1, NY + 2):
            c = 3 + y
            put(ws, r, c, f"=SUM({CL(c)}{r - NP}:{CL(c)}{r - 1})", fmt=NF_INT, font=F_B, fill=FL_TOT)
        nm(f"Vol_{ver}", "Productos", f"D{PR['vol_' + ver]}:M{PR['vol_' + ver] + NP - 1}")
        r += 2
    section(ws, r, "D. Control de capacidad — unidades vigentes por encima de la capacidad (incluye merma)", 15); r += 1
    header(ws, r, ["#", "Producto", ""] + [f"Año {y}" for y in range(1, NY + 1)] + ["Años con exceso"], 1)
    r += 1
    PR["cap_first"] = r
    for i in range(NP):
        put(ws, r, 1, i + 1)
        put(ws, r, 2, f'=IF(B{f + i}="","",B{f + i})', font=F_LINK)
        for y in range(1, NY + 1):
            c = 3 + y
            vol = f"{CL(c)}{PR['vol_V'] + i}"
            cap = f"IF(AND(N($F{f + i})>0,{y}>=$F{f + i}),N($E{f + i}),N($D{f + i}))"
            put(ws, r, c, f"=IF({cap}=0,0,MAX(0,{vol}/(1-V_Merma)-{cap}))", fmt=NF_INT)
        put(ws, r, 14, f'=SUMPRODUCT(--(D{r}:M{r}>0.5))', fmt=NF_N, font=F_B)
        r += 1
    nm("c_CapExceso", "Productos", f"N{PR['cap_first']}:N{r - 1}")
    put(ws, r, 2, "Capacidad 0 = no informada (no se controla).", font=F_S)
    ws.freeze_panes = "C6"


# ============================================================================ COSTOS
CO = {}


def build_costos():
    ws = WS["Costos"]
    title(ws, "Costos — personal por puesto y costos fijos",
          "Salarios y costos fijos pueden cargarse en PYG o USD (columna Moneda). Se convierten con el TC de cada año.")
    widths(ws, {"A": 4, "B": 40, "C": 16, "D": 9, "E": 14, "F": 14})
    for c in range(7, 17):
        ws.column_dimensions[CL(c)].width = 10
    ws.column_dimensions["Q"].width = 40
    r = 4
    section(ws, r, "A. Personal por puesto — salario mensual y dotación por año de operación (PRESUPUESTO ORIGINAL; salario vigente al lado)", 17); r += 1
    header(ws, r, ["#", "Puesto", "Clase", "Moneda", "Salario mensual — presupuesto", "Salario mensual — vigente"] +
           [f"Dot. año {y}" for y in range(1, NY + 1)] + ["Nota"], 1, 40)
    r += 1
    CO["per_first"] = r
    pers = D.get("personal", [])
    for i in range(NPER):
        p = pers[i] if i < len(pers) else None
        put(ws, r, 1, i + 1)
        inp(ws, r, 2, p[0] if p else None)
        inp(ws, r, 3, p[1] if p else None)
        inp(ws, r, 4, p[2] if p else None)
        inp(ws, r, 5, p[3] if p else None, NF_INT)
        inh(ws, r, 6, f"=E{r}", NF_INT)
        for y in range(1, NY + 1):
            v = (p[4][min(y, len(p[4])) - 1]) if p else None
            inp(ws, r, 6 + y, v, NF_INT)
        r += 1
    CO["per_last"] = r - 1
    f, l = CO["per_first"], CO["per_last"]
    dv(ws, f"C{f}:C{l}", ["Directo", "Indirecto", "Administración"])
    dv(ws, f"D{f}:D{l}", ["PYG", "USD"])
    nm("Per_Clase", "Costos", f"C{f}:C{l}")
    nm("Per_Mon", "Costos", f"D{f}:D{l}")
    nm("Sal_O", "Costos", f"E{f}:E{l}")
    nm("Sal_V", "Costos", f"F{f}:F{l}")
    nm("Dot_O", "Costos", f"G{f}:P{l}")
    put(ws, r, 2, "Totales de dotación", font=F_B)
    for y in range(1, NY + 1):
        put(ws, r, 6 + y, f"=SUM({CL(6 + y)}{f}:{CL(6 + y)}{l})", fmt=NF_INT, font=F_B, fill=FL_TOT)
    r += 1
    put(ws, r, 2, "Clase: Directo = mano de obra de producción; Indirecto = supervisión, calidad, mantenimiento, depósito (costo de producción); "
                  "Administración = gerencia general y administración.", font=F_S)
    r += 2
    section(ws, r, "B. Dotación VIGENTE por año de operación (hereda; sobrescribir lo que cambió)", 17); r += 1
    header(ws, r, ["#", "Puesto", "Clase", "", "", ""] + [f"Dot. año {y}" for y in range(1, NY + 1)], 1)
    r += 1
    CO["dotv_first"] = r
    for i in range(NPER):
        put(ws, r, 1, i + 1)
        put(ws, r, 2, f'=IF(B{f + i}="","",B{f + i})', font=F_LINK)
        put(ws, r, 3, f'=IF(C{f + i}="","",C{f + i})', font=F_LINK)
        for y in range(1, NY + 1):
            inh(ws, r, 6 + y, f"={CL(6 + y)}{f + i}", NF_INT)
        r += 1
    nm("Dot_V", "Costos", f"G{CO['dotv_first']}:P{r - 1}")
    put(ws, r, 2, "Totales de dotación vigente", font=F_B)
    for y in range(1, NY + 1):
        put(ws, r, 6 + y, f"=SUM({CL(6 + y)}{CO['dotv_first']}:{CL(6 + y)}{r - 1})", fmt=NF_INT, font=F_B, fill=FL_TOT)
    r += 2
    section(ws, r, "C. Costos fijos y gastos (monto mensual del año 1)", 17); r += 1
    header(ws, r, ["#", "Concepto", "Clase", "Moneda", "Monto mensual — presupuesto", "Monto mensual — vigente", "Crecimiento anual",
                   "Desde año de operación (0 = implantación)", "IVA incluido", "Base de imprevistos (1/0)", "Nota"], 1, 54)
    ws.merge_cells(start_row=r, start_column=11, end_row=r, end_column=17)
    r += 1
    CO["fc_first"] = r
    fij = D.get("fijos", [])
    for i in range(NFIJ):
        p = fij[i] if i < len(fij) else None
        put(ws, r, 1, i + 1)
        inp(ws, r, 2, p[0] if p else None)
        inp(ws, r, 3, p[1] if p else None)
        inp(ws, r, 4, p[2] if p else None)
        inp(ws, r, 5, p[3] if p else None, NF_USD2)
        inh(ws, r, 6, f"=E{r}", NF_USD2)
        inp(ws, r, 7, p[4] if p else None, NF_PCT)
        inp(ws, r, 8, p[5] if p else None, NF_N)
        inp(ws, r, 9, p[6] if p else None, NF_PCT)
        inp(ws, r, 10, p[7] if p else None, NF_N)
        put(ws, r, 11, p[8] if p else None, font=F_S)
        r += 1
    f2, l2 = CO["fc_first"], r - 1
    dv(ws, f"C{f2}:C{l2}", ["Producción", "Administración"])
    dv(ws, f"D{f2}:D{l2}", ["PYG", "USD"])
    for name, col in [("FC_Clase", "C"), ("FC_Mon", "D"), ("FC_O", "E"), ("FC_V", "F"), ("FC_g", "G"), ("FC_Desde", "H"), ("FC_IVA", "I"), ("FC_Base", "J")]:
        nm(name, "Costos", f"{col}{f2}:{col}{l2}")
    put(ws, r, 2, "Montos con IVA incluido cuando no es recuperable (criterio prudencial). Los imprevistos de producción se calculan sobre todos los costos de producción.",
        font=F_S)
    ws.freeze_panes = "C6"


# ============================================================================ INVERSION
IV = {}


def build_inversion():
    ws = WS["Inversion"]
    title(ws, "Inversión y financiamiento — CAPEX por ítem, préstamos y aportes (USD)",
          "Costo final estimado = ejecutado + comprometido pendiente + por contratar. Los pendientes se pagan en el «mes de pago pendiente».")
    widths(ws, {"A": 4, "B": 40, "C": 14})
    for c in range(4, 22):
        ws.column_dimensions[CL(c)].width = 12.5
    r = 4
    section(ws, r, "A. Inversiones (CAPEX)", 21); r += 1
    header(ws, r, ["#", "Ítem", "Clase", "Presupuesto original", "Mes de pago (presup.)", "Vida útil (años)",
                   "Ejecutado (devengado real)", "Pagado (real)", "Comprometido pendiente", "Por contratar (si difiere)", "Mes de pago pendiente (si difiere)",
                   "Mes de puesta en servicio real", "Por contratar usado", "Costo final estimado", "Desvío vs presupuesto", "Avance financiero",
                   "Mes de pago pendiente usado", "Pendiente devengado + comprometido", "Servicio (presup.)", "Servicio (vigente)",
                   "Base depreciable vigente"], 1, 66)
    r += 1
    IV["first"] = r
    cap = D.get("capex", [])
    for i in range(NCAP):
        p = cap[i] if i < len(cap) else None
        put(ws, r, 1, i + 1)
        inp(ws, r, 2, p[0] if p else None)
        inp(ws, r, 3, p[1] if p else None)
        inp(ws, r, 4, p[2] if p else None, NF_USD)
        inp(ws, r, 5, p[3] if p else None, NF_N)
        inp(ws, r, 6, p[4] if p else None, NF_N)
        for c in range(7, 13):
            inp(ws, r, c, None, NF_USD if c < 11 else NF_N)
        put(ws, r, 13, f'=IF(J{r}<>"",N(J{r}),MAX(0,N(D{r})-N(G{r})-N(I{r})))', fmt=NF_USD)
        put(ws, r, 14, f"=N(G{r})+N(I{r})+M{r}", fmt=NF_USD, font=F_B)
        put(ws, r, 15, f"=N{r}-N(D{r})", fmt=NF_USD)
        put(ws, r, 16, f'=IF(N{r}>0,N(H{r})/N{r},"")', fmt=NF_PCT)
        put(ws, r, 17, f'=IF(K{r}<>"",K{r},MAX(N(E{r}),Corte+1))', fmt=NF_N)
        put(ws, r, 18, f"=N(G{r})-N(H{r})+N(I{r})", fmt=NF_USD)
        put(ws, r, 19, f"=MAX(N(E{r}),P_Inicio)", fmt=NF_N)
        put(ws, r, 20, f'=IF(L{r}<>"",L{r},MAX(N(E{r}),V_Inicio))', fmt=NF_N)
        put(ws, r, 21, f"=N(G{r})+N(I{r})+M{r}*(1+V_dCapex)", fmt=NF_USD)
        r += 1
    IV["last"] = r - 1
    f, l = IV["first"], IV["last"]
    for name, col in [("Inv_D", "D"), ("Inv_E", "E"), ("Inv_F", "F"), ("Inv_H", "H"), ("Inv_M", "M"), ("Inv_Q", "Q"), ("Inv_R", "R"),
                      ("Inv_SvcP", "S"), ("Inv_SvcV", "T"), ("Inv_BaseV", "U"), ("Inv_CF", "N")]:
        nm(name, "Inversion", f"{col}{f}:{col}{l}")
    put(ws, r, 2, "TOTAL", font=F_B)
    for c in [4, 7, 8, 9, 13, 14, 15, 18, 21]:
        put(ws, r, c, f"=SUM({CL(c)}{f}:{CL(c)}{l})", fmt=NF_USD, font=F_B, fill=FL_TOT)
    put(ws, r, 16, f'=IF(N{r}>0,H{r}/N{r},"")', fmt=NF_PCT, font=F_B)
    IV["tot"] = r
    nm("i_CapexP", "Inversion", f"D{r}")
    nm("i_CapexV", "Inversion", f"N{r}")
    nm("i_Pagado", "Inversion", f"H{r}")
    r += 1
    put(ws, r, 2, D.get("capex_nota") or "Aportes en especie (máquinas): cargar el ítem aquí y el mismo monto como aporte en el mismo mes.", font=F_S)
    r += 2
    IV["loans"] = r
    section(ws, r, "B. Préstamos (sistema francés; desembolso al inicio del mes; tasa mensual = nominal anual / 12)", 21); r += 1
    header(ws, r, ["", "Condición", "Unidad", "Préstamo 1 — presupuesto", "Préstamo 1 — vigente", "Préstamo 2 — presupuesto", "Préstamo 2 — vigente",
                   "Préstamo 3 — presupuesto", "Préstamo 3 — vigente"], 1, 42)
    r += 1
    loans = D.get("prestamos", [])
    for k, lab, unit, fmt, idx in [("Nombre", "Nombre / banco", "", None, 0), ("Monto", "Monto", "USD", NF_USD, 1),
                                   ("Mes", "Mes de desembolso", "mes", NF_N, 2), ("Tasa", "Tasa nominal anual", "%", NF_PCT2, 3),
                                   ("Gracia", "Meses de gracia (solo intereses)", "meses", NF_N, 4), ("Cuotas", "Cuotas de capital e interés", "cuotas", NF_N, 5)]:
        put(ws, r, 2, lab)
        put(ws, r, 3, unit)
        for j in range(NL):
            lo = loans[j] if j < len(loans) else None
            co, cv = 4 + 2 * j, 5 + 2 * j
            inp(ws, r, co, lo[idx] if lo else None, fmt)
            inh(ws, r, cv, f"={CL(co)}{r}", fmt)
            if k != "Nombre":
                nm(f"P_L{j + 1}_{k}", "Inversion", f"{CL(co)}{r}")
                nm(f"V_L{j + 1}_{k}", "Inversion", f"{CL(cv)}{r}")
        r += 1
    put(ws, r, 2, "Préstamo de corto plazo con cronograma (p. ej. 12 cuotas o pago único): cargarlo como préstamo 2 o 3.", font=F_S)
    r += 2
    section(ws, r, "B2. Línea de crédito rotativa de corto plazo (capital de trabajo): se usa solo si la caja cae bajo el mínimo y se devuelve con el excedente", 21); r += 1
    header(ws, r, ["", "Condición", "Unidad", "Presupuesto", "Vigente", "Nota"], 1, 30)
    ws.merge_cells(start_row=r, start_column=6, end_row=r, end_column=14)
    r += 1
    lc = D.get("linea", {})
    for k, lab, unit, fmt, note in [("Lim", "Límite aprobado", "USD", NF_USD, "0 o vacío = sin línea."),
                                    ("Tasa", "Tasa nominal anual (interés mensual sobre el saldo usado)", "%", NF_PCT2, "Más IVA sobre intereses (Supuestos G24)."),
                                    ("Desde", "Disponible desde el mes", "mes", NF_N, "Vacío = desde el mes 1."),
                                    ("Hasta", "Vencimiento: último mes de uso (luego se cancela el saldo)", "mes", NF_N,
                                     "Vacío = fin del horizonte. Al vencer se paga todo el saldo aunque falte caja (riesgo de renovación).")]:
        put(ws, r, 2, lab)
        put(ws, r, 3, unit)
        inp(ws, r, 4, lc.get(k), fmt)
        inh(ws, r, 5, f"=D{r}", fmt)
        put(ws, r, 6, note, font=F_S)
        nm(f"P_LC_{k}", "Inversion", f"D{r}")
        nm(f"V_LC_{k}", "Inversion", f"E{r}")
        r += 1
    put(ws, r, 2, "Saldo utilizado de la línea al último mes cerrado (solo vigente)")
    put(ws, r, 3, "USD")
    inp(ws, r, 5, None, NF_USD)
    put(ws, r, 6, "Los usos y devoluciones reales se cargan en Reales (desembolsos / amortización); aquí solo el saldo al corte.", font=F_S)
    nm("V_LC_Saldo0", "Inversion", f"E{r}")
    r += 1
    put(ws, r, 2, "Los dividendos solo se pagan con caja sobre el mínimo después de devolver la línea.", font=F_S)
    r += 2
    section(ws, r, "C. Aportes de socios (efectivo o especie)", 21); r += 1
    header(ws, r, ["#", "Concepto", "Mes — presupuesto", "Monto — presupuesto", "Mes — vigente", "Monto — vigente", "Nota"], 1, 30)
    r += 1
    IV["ap_first"] = r
    aps = D.get("aportes", [])
    for i in range(NAP):
        a = aps[i] if i < len(aps) else None
        put(ws, r, 1, i + 1)
        inp(ws, r, 2, a[0] if a else None)
        inp(ws, r, 3, a[1] if a else None, NF_N)
        inp(ws, r, 4, a[2] if a else None, NF_USD)
        inh(ws, r, 5, f"=C{r}", NF_N)
        inh(ws, r, 6, f"=D{r}", NF_USD)
        r += 1
    fa, la = IV["ap_first"], r - 1
    nm("Ap_MesO", "Inversion", f"C{fa}:C{la}")
    nm("Ap_MontoO", "Inversion", f"D{fa}:D{la}")
    nm("Ap_MesV", "Inversion", f"E{fa}:E{la}")
    nm("Ap_MontoV", "Inversion", f"F{fa}:F{la}")
    ws.freeze_panes = "C6"


# ============================================================================ REALES
RR = {}
REAL_ROWS = [
    ("sec", "Tipo de cambio"),
    ("TC", "Tipo de cambio real del mes (Gs/USD)", "Gs/USD"),
    ("sec", "Unidades vendidas por producto"),
] + [(f"U{i + 1}", None, "u") for i in range(NP)] + [
    ("sec", "Ventas por producto (facturación neta)"),
] + [(f"S{i + 1}", None, "mon") for i in range(NP)] + [
    ("sec", "Costos y gastos devengados (sin IVA recuperable)"),
    ("MAT", "Materia prima e insumos consumidos", "mon"),
    ("OVAR", "Otros costos variables (flete, despacho, packaging, seguros, comisiones)", "mon"),
    ("PERP", "Personal de producción (directo e indirecto, con cargas)", "mon"),
    ("PERA", "Personal administrativo (con cargas)", "mon"),
    ("FIJP", "Costos fijos de producción (energía, mantenimiento, otros)", "mon"),
    ("ADM", "Gastos de administración (alquiler, servicios, honorarios, otros)", "mon"),
    ("TRIB", "Tributo de maquila", "mon"),
    ("sec", "Inversión y financiamiento (movimientos de caja)"),
    ("CAPEX", "CAPEX pagado", "mon"),
    ("DESEMB", "Desembolsos de préstamos recibidos (incluye línea rotativa)", "mon"),
    ("INT", "Intereses, IVA y comisiones pagados", "mon"),
    ("AMORT", "Amortización de préstamos pagada (incluye devoluciones de la línea)", "mon"),
    ("APORTE", "Aportes de socios recibidos", "mon"),
    ("DIV", "Dividendos pagados", "mon"),
    ("sec", "Saldos al cierre del mes (obligatorio en el último mes cerrado)"),
    ("S_CAJA", "Caja y bancos", "mon"),
    ("S_CXC", "Cuentas por cobrar a clientes", "mon"),
    ("S_INV", "Inventarios (materiales y producto terminado)", "mon"),
    ("S_IVA", "Crédito fiscal IVA por recuperar", "mon"),
    ("S_PROV", "Proveedores y cuentas por pagar", "mon"),
    ("S_DEUDA", "Deuda financiera (control)", "mon"),
]


def build_reales():
    ws = WS["Reales"]
    title(ws, "Reales — carga mensual de la ejecución (un dato por celda; solo meses cerrados)",
          "Montos en la moneda de la columna C (USD o PYG). Un mes cerrado sin datos se informa en Control (un faltante no equivale a cero).")
    widths(ws, {"A": 9, "B": 52, "C": 8, "D": 14, "E": 4})
    for m in range(1, N + 1):
        ws.column_dimensions[mc(m)].width = 11
    r = 4
    labels = [("Mes n°", None), ("Fecha", NF_MES), ("Año del modelo", NF_N), ("Estado", None)]
    for i, (lab, fmt) in enumerate(labels):
        put(ws, r + i, 2, lab, font=F_H, fill=FL_H)
        for c in (1, 3, 4, 5):
            ws.cell(row=r + i, column=c).fill = FL_H
        for m in range(1, N + 1):
            if i == 0:
                v = m
            elif i == 1:
                v = f'=IF(n_Fecha1="","",EOMONTH(n_Fecha1,{m}-1))'
            elif i == 2:
                v = f"=ROUNDUP({m}/12,0)"
            else:
                v = f'=IF({m}<=Corte,"Cerrado","Abierto")'
            put(ws, r + i, FCOL + m - 1, v, font=F_H, fill=FL_H, fmt=fmt, al=CEN)
    put(ws, r, 4, "Total cargado", font=F_H, fill=FL_H)
    RR["hdr"] = r
    r += 5
    for it in REAL_ROWS:
        if it[0] == "sec":
            section(ws, r, it[1], 5)
            r += 1
            continue
        code, lab, unit = it
        RR[code] = r
        put(ws, r, 1, code, font=F_S)
        if code.startswith("U"):
            i = int(code[1:]) - 1
            lab = f'=IF(INDEX(Prod_Nombre,{i + 1})="","Producto {i + 1} (unidades)",INDEX(Prod_Nombre,{i + 1})&" (unidades)")'
            put(ws, r, 2, lab, font=F_LINK)
            put(ws, r, 3, "u")
        elif code.startswith("S") and code[1:].isdigit():
            i = int(code[1:]) - 1
            lab = f'=IF(INDEX(Prod_Nombre,{i + 1})="","Producto {i + 1} (ventas)",INDEX(Prod_Nombre,{i + 1})&" (ventas)")'
            put(ws, r, 2, lab, font=F_LINK)
            inp(ws, r, 3, "USD")
        else:
            put(ws, r, 2, lab)
            if unit == "mon":
                inp(ws, r, 3, "PYG" if code in ("PERP", "PERA") else "USD")
            else:
                put(ws, r, 3, unit)
        put(ws, r, 4, f"=SUM({M1}{r}:{MN}{r})" if not code.startswith("S_") and code != "TC" else None, fmt=NF_USD, font=F_B)
        for m in range(1, N + 1):
            inp(ws, r, FCOL + m - 1, None, NF_USD)
        r += 1
    dv(ws, f"C{RR['S1']}:C{RR['S_DEUDA']}", ["USD", "PYG"])
    ws.conditional_formatting.add(f"{M1}{RR['hdr'] + 3}:{MN}{RR['hdr'] + 3}",
                                  CellIsRule(operator="equal", formula=['"Cerrado"'], fill=FL_V, font=Font(name=FONT, size=9, bold=True, color="000000")))
    ws.freeze_panes = f"{M1}{RR['hdr'] + 4}"


# ============================================================================ DETALLE 24 MESES
DT = {}
DET_BLOQUES = [  # código, título, filas, unidad, referencia anual (fórmula para el índice i, 1..n)
    ("TC", "A. Tipo de cambio del mes (Gs/USD)", 1, "Gs/USD", lambda i: "=INDEX(TC_V,1,1)"),
    ("Q", "B. Volumen de ventas por producto (unidades del mes)", NP, "u", lambda i: f"=INDEX(Vol_V,{i},1)/12"),
    ("P", "C. Precio de venta por producto (USD por unidad)", NP, "USD/u", lambda i: f"=INDEX(Precio_V,{i})"),
    ("MP", "D. Materia prima por unidad (USD)", NP, "USD/u", lambda i: f"=INDEX(MP_V,{i})"),
    ("IN", "E. Insumos por unidad (USD)", NP, "USD/u", lambda i: f"=INDEX(INS_V,{i})"),
    ("OV", "F. Otros costos variables por unidad (USD)", NP, "USD/u", lambda i: f"=INDEX(OV_V,{i})"),
    ("DO", "G. Dotación por puesto (personas en el mes)", NPER, "pers.", lambda i: f"=INDEX(Dot_V,{i},1)"),
    ("SA", "H. Salario mensual por puesto (moneda del puesto, sin cargas)", NPER, "mon", lambda i: f"=INDEX(Sal_V,{i})"),
    ("FC", "I. Costos fijos: monto del mes por concepto (moneda del concepto)", NFIJ, "mon", lambda i: f"=INDEX(FC_V,{i})"),
]


def build_detalle():
    ws = WS["Detalle_24m"]
    title(ws, "Detalle_24m — proyección mes a mes de la ventana móvil (vigente)",
          "Celda vacía = usa el supuesto anual vigente. Solo se usan los meses de la ventana (amarillo): del mes siguiente al último cierre hasta el mes Corte + n_Ventana.")
    widths(ws, {"A": 7, "B": 46, "C": 9, "D": 14, "E": 4})
    for m in range(1, N + 1):
        ws.column_dimensions[mc(m)].width = 10
    r = 4
    labels = [("Mes n°", None), ("Fecha", NF_MES), ("Año del modelo", NF_N), ("Estado", None)]
    for i, (lab, fmt) in enumerate(labels):
        put(ws, r + i, 2, lab, font=F_H, fill=FL_H)
        for c in (1, 3, 4, 5):
            ws.cell(row=r + i, column=c).fill = FL_H
        for m in range(1, N + 1):
            v = [m, f'=IF(n_Fecha1="","",EOMONTH(n_Fecha1,{m}-1))', f"=ROUNDUP({m}/12,0)",
                 f'=IF({m}<=Corte,"Real",IF({m}<=Corte+n_Ventana,"Ventana","Anual"))'][i]
            put(ws, r + i, FCOL + m - 1, v, font=F_H, fill=FL_H, fmt=fmt, al=CEN)
    put(ws, r, 4, "Ref. anual", font=F_H, fill=FL_H)
    DT["hdr"] = r
    put(ws, 3, 2, '=HYPERLINK("#Detalle_24m!"&ADDRESS(9,MIN(149,5+Corte+1),4),"→ Ir al primer mes de la ventana: mes "&(Corte+1))',
        font=Font(name=FONT, size=9, color="0563C1", underline="single"))
    r += 5
    per_rows = lambda j: f"Costos!$B${CO['per_first'] + j - 1}"
    for code, tit, n, unit, ref in DET_BLOQUES:
        section(ws, r, tit, 5)
        r += 1
        DT[code + "_first"] = r
        for i in range(1, n + 1):
            rr = r
            DT[f"{code}{i}"] = rr
            put(ws, rr, 1, f"{code}{i}" if n > 1 else code, font=F_S)
            if code == "TC":
                put(ws, rr, 2, "Tipo de cambio del mes")
                put(ws, rr, 3, unit)
            elif code in ("Q", "P", "MP", "IN", "OV"):
                put(ws, rr, 2, f'=IF(INDEX(Prod_Nombre,{i})="","Producto {i}",INDEX(Prod_Nombre,{i}))', font=F_LINK)
                put(ws, rr, 3, unit)
            elif code in ("DO", "SA"):
                put(ws, rr, 2, f'=IF({per_rows(i)}="","Puesto {i}",{per_rows(i)})', font=F_LINK)
                put(ws, rr, 3, unit if code == "DO" else f'=IF(Costos!$D${CO["per_first"] + i - 1}="","",Costos!$D${CO["per_first"] + i - 1})', font=F_LINK)
            else:
                put(ws, rr, 2, f'=IF(Costos!$B${CO["fc_first"] + i - 1}="","Concepto {i}",Costos!$B${CO["fc_first"] + i - 1})', font=F_LINK)
                put(ws, rr, 3, f'=IF(Costos!$D${CO["fc_first"] + i - 1}="","",Costos!$D${CO["fc_first"] + i - 1})', font=F_LINK)
            put(ws, rr, 4, ref(i), fmt=NF_USD2, font=F_S)
            r += 1
        DT[code + "_last"] = r - 1
        r += 2   # fila libre: los contadores «celdas con dato en el mes» están en Calculo (no se borran al limpiar el detalle)
    hdr = DT["hdr"]
    area = f"{M1}{DT['TC_first']}:{MN}{r}"
    ws.conditional_formatting.add(area, FormulaRule(formula=[f"AND({M1}${hdr}>Corte,{M1}${hdr}<=Corte+n_Ventana)"], fill=FL_IN,
                                                     font=Font(name=FONT, size=9, color="0000FF")))
    ws.conditional_formatting.add(area, FormulaRule(formula=[f"{M1}${hdr}<=Corte"], fill=FL_INH))
    ws.conditional_formatting.add(f"{M1}{hdr + 3}:{MN}{hdr + 3}", CellIsRule(operator="equal", formula=['"Ventana"'], fill=FL_V,
                                                                         font=Font(name=FONT, size=9, bold=True, color="000000")))
    put(ws, r, 2, "Montos de salarios y costos fijos en la moneda definida en Costos (PYG o USD). Volumen: se usa desde el inicio de operación.", font=F_S)
    ws.freeze_panes = f"{M1}{hdr + 4}"


def dcol(code, m, n):
    """Columna del mes m de un bloque de Detalle_24m (n filas)."""
    return f"Detalle_24m!{mc(m)}${DT[code + '1']}:{mc(m)}${DT[code + str(n)]}"


def dcell(code, m):
    return f"Detalle_24m!{mc(m)}${DT[code]}"


# ============================================================================ CALCULO
ROW = {}
CUR = ["Calculo"]


def A(key, m):
    return f"{mc(m)}${ROW[key]}"


def RG(key, sheet=None):
    pre = f"{sheet}!" if sheet else ""
    return f"{pre}${M1}${ROW[key]}:${MN}${ROW[key]}"


def prv(key, m, d="0"):
    return A(key, m - 1) if m > 1 else d


def rv(code, m):
    """Valor real convertido a USD (Reales), con el TC real del mes o, si falta, el vigente."""
    c, r = mc(m), RR[code]
    if code.startswith("U"):
        return f"N(Reales!{c}{r})"
    return f'IF(Reales!$C${r}="PYG",IF({A("tcr", m)}>0,N(Reales!{c}{r})/{A("tcr", m)},0),N(Reales!{c}{r}))'


def rgiven(code, m):
    return f"ISNUMBER(Reales!{mc(m)}{RR[code]})"


def version_rows(X):
    """Filas del motor para la versión X ('P' presupuesto, 'V' vigente). Mismas fórmulas."""
    O = "O" if X == "P" else "V"
    real = X == "V"
    K = lambda k: f"{X}.{k}"
    a = lambda k, m: A(K(k), m)
    hor = lambda m: A("hor", m)

    def R(code, m, fc):
        return f"=IF({A('real', m)}=1,{rv(code, m)},{fc})" if real else f"={fc}"

    win = lambda m: f"AND({m}>Corte,{m}<=Corte+n_Ventana)"          # ventana de detalle mensual (solo vigente)
    wn = lambda m: f"({m}>Corte)*({m}<=Corte+n_Ventana)"
    rows = []
    rows.append(("sec", "Calendario de la versión"))
    rows += [
        ("aop", "Año de operación (0 = fuera de operación)", "n°", lambda m: f"=IF(AND({m}>={X}_Inicio,{m}<=Fin),INT(({m}-{X}_Inicio)/12)+1,0)", NF_N, None),
        ("op", "Operación (1/0)", "1/0", lambda m: f"=IF({a('aop', m)}>0,1,0)", NF_N, "sum"),
        ("pre", "Implantación: antes del inicio de operación (1/0)", "1/0", lambda m: f"=IF(AND({m}<{X}_Inicio,{m}<=Fin),1,0)", NF_N, "sum"),
        ("ya", "Índice de año para crecimientos (mín. 1)", "n°", lambda m: f"=MAX(1,{a('aop', m)})", NF_N, None),
        ("tc", "Tipo de cambio aplicado (Gs/USD)", "Gs/USD",
         (lambda m: (f"=IF(AND({win(m)},ISNUMBER({dcell('TC1', m)})),{dcell('TC1', m)},INDEX(TC_{O},1,MIN({NY},{A('anio', m)})))*(1+{X}_dTC)" if real
                     else f"=INDEX(TC_{O},1,MIN({NY},{A('anio', m)}))*(1+{X}_dTC)")), NF_INT, None),
        ("itc", "Inverso del TC (0 si falta el TC)", "", lambda m: f"=IF({a('tc', m)}>0,1/{a('tc', m)},0)", NF_USD4, None),
    ]
    rows.append(("sec", "Ventas"))
    for i in range(NP):
        rows.append((f"u{i + 1}", f'=IF(INDEX(Prod_Nombre,{i + 1})="","Producto {i + 1}",INDEX(Prod_Nombre,{i + 1}))&" — unidades"', "u",
                     (lambda m, i=i: (f"=IF({A('real', m)}=1,{rv('U' + str(i + 1), m)},IF({a('op', m)}=1,IF(AND({win(m)},ISNUMBER({dcell('Q' + str(i + 1), m)})),"
                                      f"{dcell('Q' + str(i + 1), m)},INDEX(Vol_{O},{i + 1},{a('aop', m)})/12)*(1+{X}_dVol),0))"
                                      if real else f"=IF({a('op', m)}=1,INDEX(Vol_{O},{i + 1},{a('aop', m)})/12*(1+{X}_dVol),0)")), NF_INT, "sum"))
    for i in range(NP):
        if real:
            fc = lambda m, i=i: (f"{a('u' + str(i + 1), m)}*IF(AND({win(m)},ISNUMBER({dcell('P' + str(i + 1), m)})),{dcell('P' + str(i + 1), m)},"
                                 f"INDEX(Precio_{O},{i + 1})*INDEX(IdxPrecio_{O},1,MIN({NY},{a('ya', m)})))*(1+{X}_dPrecio)")
        else:
            fc = lambda m, i=i: f"{a('u' + str(i + 1), m)}*INDEX(Precio_{O},{i + 1})*INDEX(IdxPrecio_{O},1,MIN({NY},{a('ya', m)}))*(1+{X}_dPrecio)"
        rows.append((f"s{i + 1}", f'=IF(INDEX(Prod_Nombre,{i + 1})="","Producto {i + 1}",INDEX(Prod_Nombre,{i + 1}))&" — ventas USD"', "USD",
                     (lambda m, i=i, fc=fc: R("S" + str(i + 1), m, fc(m))), NF_USD, "sum"))
    rows.append(("vtas", "VENTAS TOTALES", "USD", lambda m: f"=SUM({a('s1', m)}:{a('s' + str(NP), m)})", NF_USD, "sum", True))
    urng = lambda m: f"{mc(m)}{ROW[K('u1')]}:{mc(m)}{ROW[K('u' + str(NP))]}"
    yi = lambda nmx, m: f"INDEX({nmx}_{O},1,MIN({NY},{a('ya', m)}))"
    rows.append(("sec", "Costos de la proyección (en meses reales se reemplazan por Reales)"))
    def matv(code, nmx, idx):
        if not real:
            return lambda m: f"=SUMPRODUCT({urng(m)},{nmx}_{O})*{yi(idx, m)}*(1+{X}_dMat)/(1-{X}_Merma)"
        return lambda m: (f"=SUMPRODUCT({urng(m)},{wn(m)}*ISNUMBER({dcol(code, m, NP)})*{dcol(code, m, NP)}"
                          f"+(1-{wn(m)}*ISNUMBER({dcol(code, m, NP)}))*{nmx}_{O}*{yi(idx, m)})*(1+{X}_dMat)/(1-{X}_Merma)")
    rows += [
        ("mp_f", "Materia prima", "USD", matv("MP", "MP", "IdxMP"), NF_USD, "sum"),
        ("ins_f", "Insumos", "USD", matv("IN", "INS", "IdxOV"), NF_USD, "sum"),
        ("ov_f", "Otros costos variables por unidad", "USD", matv("OV", "OV", "IdxOV"), NF_USD, "sum"),
        ("desp_f", "Despacho de importación", "USD", lambda m: f"={X}_Desp*({a('mp_f', m)}+{a('ins_f', m)})", NF_USD, "sum"),
    ]
    act = lambda m, coef: f"({a('op', m)}+{a('pre', m)}*{coef})"
    for cls, k, coef in [("Directo", "dir", f"{X}_PreDir"), ("Indirecto", "ind", f"{X}_PreInd"), ("Administración", "adm", f"{X}_PreInd")]:
        if real:
            dot = lambda m, coef=coef: (f"({wn(m)}*ISNUMBER({dcol('DO', m, NPER)})*{dcol('DO', m, NPER)}+(1-{wn(m)}*ISNUMBER({dcol('DO', m, NPER)}))"
                                        f"*INDEX(Dot_{O},0,MIN({NY},{a('ya', m)}))*{act(m, coef)})")
            sal = lambda m: (f"({wn(m)}*ISNUMBER({dcol('SA', m, NPER)})*{dcol('SA', m, NPER)}+(1-{wn(m)}*ISNUMBER({dcol('SA', m, NPER)}))"
                             f"*Sal_{O}*{yi('IdxSal', m)})")
            rows.append((f"{k}_n", f"Personas activas — {cls}", "pers.",
                         (lambda m, cls=cls, dot=dot: f'=SUMPRODUCT((Per_Clase="{cls}")*{dot(m)})'), NF_USD2, None))
            rows.append((f"{k}_f", f"Personal — {cls} (salarios, cargas y alimentación)", "USD",
                         (lambda m, cls=cls, k=k, dot=dot, sal=sal: (
                             f'=SUMPRODUCT((Per_Clase="{cls}")*{sal(m)}*{dot(m)}*((Per_Mon="PYG")*{a("itc", m)}+(Per_Mon<>"PYG")))'
                             f"*(1+{X}_Cargas)+{a(k + '_n', m)}*{X}_Alim*{a('itc', m)}")), NF_USD, "sum"))
            continue
        rows.append((f"{k}_n", f"Personas activas — {cls}", "pers.",
                     (lambda m, cls=cls, coef=coef: f'=SUMPRODUCT((Per_Clase="{cls}")*INDEX(Dot_{O},0,MIN({NY},{a("ya", m)})))*{act(m, coef)}'), NF_USD2, None))
        rows.append((f"{k}_f", f"Personal — {cls} (salarios, cargas y alimentación)", "USD",
                     (lambda m, cls=cls, k=k, coef=coef: (
                         f'=SUMPRODUCT((Per_Clase="{cls}")*Sal_{O}*INDEX(Dot_{O},0,MIN({NY},{a("ya", m)}))*((Per_Mon="PYG")*{a("itc", m)}+(Per_Mon<>"PYG")))'
                         f"*{act(m, coef)}*{yi('IdxSal', m)}*(1+{X}_Cargas)+{a(k + '_n', m)}*{X}_Alim*{a('itc', m)}")), NF_USD, "sum"))
    fcact = lambda m: f"((FC_Desde=0)*({a('pre', m)}+{a('op', m)})+(FC_Desde>=1)*({a('aop', m)}>=FC_Desde)*{a('op', m)})"
    if real:
        fcusd = lambda m: (f"((FC_Mon=\"PYG\")*{a('itc', m)}+(FC_Mon<>\"PYG\"))*({wn(m)}*ISNUMBER({dcol('FC', m, NFIJ)})*{dcol('FC', m, NFIJ)}"
                           f"+(1-{wn(m)}*ISNUMBER({dcol('FC', m, NFIJ)}))*FC_{O}*(1+FC_g)^({a('ya', m)}-1)*{fcact(m)})")
    else:
        fcusd = lambda m: f"FC_{O}*((FC_Mon=\"PYG\")*{a('itc', m)}+(FC_Mon<>\"PYG\"))*(1+FC_g)^({a('ya', m)}-1)*{fcact(m)}"
    rows += [
        ("pkg_f", "Packaging", "USD", lambda m: f"={X}_Pkg*({a('mp_f', m)}+{a('ins_f', m)}+{a('ov_f', m)}+{a('dir_f', m)})", NF_USD, "sum"),
        ("vs_f", "Seguro de caución, comisiones y logística de venta", "USD", lambda m: f"=({X}_SegV+{X}_ComV)*{a('vtas', m)}", NF_USD, "sum"),
        ("trib_f", "Tributo de maquila", "USD", lambda m: f"={X}_Trib*{a('vtas', m)}", NF_USD, "sum"),
        ("fp_f", "Costos fijos de producción", "USD", lambda m: f'=SUMPRODUCT((FC_Clase="Producción")*{fcusd(m)})', NF_USD, "sum"),
        ("fa_f", "Costos fijos de administración", "USD", lambda m: f'=SUMPRODUCT((FC_Clase="Administración")*{fcusd(m)})', NF_USD, "sum"),
        ("fab_f", "Costos fijos de administración en base de imprevistos", "USD", lambda m: f'=SUMPRODUCT((FC_Clase="Administración")*FC_Base*{fcusd(m)})', NF_USD, "sum"),
        ("impp_f", "Imprevistos de producción", "USD",
         lambda m: (f"={X}_ImpP*({a('mp_f', m)}+{a('ins_f', m)}+{a('ov_f', m)}+{a('desp_f', m)}+{a('pkg_f', m)}+{a('vs_f', m)}"
                    f"+{a('dir_f', m)}+{a('ind_f', m)}+{a('fp_f', m)})"), NF_USD, "sum"),
        ("impa_f", "Imprevistos de administración", "USD", lambda m: f"={X}_ImpA*({a('adm_f', m)}+{a('fab_f', m)})", NF_USD, "sum"),
        ("ivac_f", "IVA contenido en costos fijos", "USD", lambda m: f"=SUMPRODUCT({fcusd(m)}*FC_IVA/(1+FC_IVA))", NF_USD, "sum"),
    ]
    rows.append(("sec", "Estado de resultados"))
    rows += [
        ("MAT", "Materia prima e insumos", "USD", lambda m: R("MAT", m, f"{a('mp_f', m)}+{a('ins_f', m)}"), NF_USD, "sum"),
        ("OVAR", "Otros costos variables", "USD", lambda m: R("OVAR", m, f"{a('ov_f', m)}+{a('desp_f', m)}+{a('pkg_f', m)}+{a('vs_f', m)}"), NF_USD, "sum"),
        ("PERP", "Personal de producción", "USD", lambda m: R("PERP", m, f"{a('dir_f', m)}+{a('ind_f', m)}"), NF_USD, "sum"),
        ("FIJP", "Costos fijos de producción e imprevistos", "USD", lambda m: R("FIJP", m, f"{a('fp_f', m)}+{a('impp_f', m)}"), NF_USD, "sum"),
        ("MB", "MARGEN BRUTO", "USD", lambda m: f"={a('vtas', m)}-{a('MAT', m)}-{a('OVAR', m)}-{a('PERP', m)}-{a('FIJP', m)}", NF_USD, "sum", True),
        ("PERA", "Personal administrativo", "USD", lambda m: R("PERA", m, f"{a('adm_f', m)}"), NF_USD, "sum"),
        ("ADM", "Gastos de administración e imprevistos", "USD", lambda m: R("ADM", m, f"{a('fa_f', m)}+{a('impa_f', m)}"), NF_USD, "sum"),
        ("TRIB", "Tributo de maquila", "USD", lambda m: R("TRIB", m, f"{a('trib_f', m)}"), NF_USD, "sum"),
        ("IVAR", "(+) IVA recuperable (reduce costos)", "USD",
         lambda m: (f"=IF({A('real', m)}=1,0,{a('ivac_f', m)}*{X}_IVArec)" if real else f"={a('ivac_f', m)}*{X}_IVArec"), NF_USD, "sum"),
        ("EBITDA", "EBITDA", "USD", lambda m: f"={a('MB', m)}-{a('PERA', m)}-{a('ADM', m)}-{a('TRIB', m)}+{a('IVAR', m)}", NF_USD, "sum", True),
        ("DEP", "Depreciación y amortización", "USD",
         lambda m: (f"=SUMPRODUCT({'Inv_D' if X == 'P' else '(Inv_CF*' + A('real', m) + '+Inv_BaseV*(1-' + A('real', m) + '))'}"
                    f"*(Inv_F>0)*({m}>=Inv_Svc{X})*({m}<Inv_Svc{X}+Inv_F*12)/(Inv_F*12+(Inv_F=0)))*{hor(m)}"),
         NF_USD, "sum"),
        ("EBIT", "EBIT", "USD", lambda m: f"={a('EBITDA', m)}-{a('DEP', m)}", NF_USD, "sum"),
        ("INT", "Intereses, IVA y comisiones", "USD",
         lambda m: R("INT", m, f"({a('L1_int', m)}+{a('L2_int', m)}+{a('L3_int', m)}+{a('LCI', m)})*(1+{X}_IVAint)"), NF_USD, "sum"),
        ("IRE", "Impuesto a la renta", "USD", lambda m: f"={X}_IRE*({a('EBIT', m)}-{a('INT', m)})", NF_USD, "sum"),
        ("NI", "RESULTADO NETO", "USD", lambda m: f"={a('EBIT', m)}-{a('INT', m)}-{a('IRE', m)}", NF_USD, "sum", True),
    ]

    def bal(code, calc):
        def f(m):
            base = f"IF(AND({A('real', m)}=1,{rgiven(code, m)}),{rv(code, m)},{calc(m)})" if real else calc(m)
            return f"=IF({m}>Fin,{prv(K(code_key[code]), m)},{base})"
        return f
    code_key = {"S_CXC": "cxc", "S_INV": "inv", "S_IVA": "ivab", "S_PROV": "prov"}

    def dias(k, m):
        """Días de cobro/inventario: en meses reales sin saldo cargado no se aplica el escenario."""
        return f"IF({A('real', m)}=1,N(vg_{k}),{X}_{k})" if real else f"{X}_{k}"
    rows.append(("sec", "Capital de trabajo (saldos por días; en el mes de corte se usan los saldos reales)"))
    rows += [
        ("cxc", "Cuentas por cobrar", "USD", bal("S_CXC", lambda m: f"{a('vtas', m)}*{dias('DSO', m)}/DiasMes"), NF_USD, "last"),
        ("inv", "Inventarios (materiales y producto terminado)", "USD",
         bal("S_INV", lambda m: f"{a('MAT', m)}*{dias('DInv', m)}/DiasMes+({a('MAT', m)}+{a('OVAR', m)}+{a('PERP', m)}+{a('FIJP', m)})*{X}_DPT/DiasMes"), NF_USD, "last"),
        ("ivab", "Crédito fiscal IVA", "USD", bal("S_IVA", lambda m: f"{a('IVAR', m)}*{X}_IVAmeses"), NF_USD, "last"),
        ("prov", "Proveedores", "USD",
         bal("S_PROV", lambda m: f"{a('MAT', m)}*{X}_DPOmp/DiasMes+({a('OVAR', m)}+{a('FIJP', m)}+{a('ADM', m)}+{a('TRIB', m)})*{X}_DPOot/DiasMes"),
         NF_USD, "last"),
        ("ctn", "CAPITAL DE TRABAJO NETO", "USD", lambda m: f"={a('cxc', m)}+{a('inv', m)}+{a('ivab', m)}-{a('prov', m)}", NF_USD, "max", True),
        ("dct", "Variación del capital de trabajo", "USD", lambda m: f"={a('ctn', m)}-{prv(K('ctn'), m)}", NF_USD, "sum"),
    ]
    rows.append(("sec", "Flujo de caja e inversión"))
    capex_f = (lambda m: f"(SUMIF(Inv_Q,{m},Inv_R)+SUMIF(Inv_Q,{m},Inv_M)*(1+{X}_dCapex))*{hor(m)}") if real else (lambda m: f"SUMIF(Inv_E,{m},Inv_D)*{hor(m)}")
    rows += [
        ("CFO", "FLUJO OPERATIVO (EBITDA - impuesto - Δ capital de trabajo)", "USD", lambda m: f"=({a('EBITDA', m)}-{a('IRE', m)}-{a('dct', m)})*{hor(m)}", NF_USD, "sum", True),
        ("CAPEX", "CAPEX pagado", "USD", lambda m: R("CAPEX", m, capex_f(m)), NF_USD, "sum"),
        ("AF", "Activo fijo neto (CAPEX acumulado - depreciación acumulada)", "USD",
         lambda m: f"={prv(K('AF'), m)}+{a('CAPEX', m)}-{a('DEP', m)}", NF_USD, "last"),
        ("VT", "Valor terminal (capital de trabajo y activo fijo realizables)", "USD",
         lambda m: f"=IF({m}=Fin,{a('ctn', m)}*{X}_RealCT+{a('AF', m)}*{X}_RealAF,0)", NF_USD, "sum"),
        ("FCFF", "FLUJO LIBRE DEL PROYECTO (FCFF)", "USD", lambda m: f"={a('CFO', m)}-{a('CAPEX', m)}+{a('VT', m)}", NF_USD, "sum", True),
        ("CUML", "Flujo acumulado antes de financiamiento (CFO - CAPEX)", "USD", lambda m: f"={prv(K('CUML'), m)}+{a('CFO', m)}-{a('CAPEX', m)}", NF_USD, "min"),
    ]
    rows.append(("sec", "Préstamos (cronograma contractual)"))
    for j in range(1, NL + 1):
        L = f"{X}_L{j}"
        kk = lambda m, L=L: f"({m}-N({L}_Mes)+1)"
        rows += [
            (f"L{j}_des", f"Préstamo {j}: desembolso", "USD", lambda m, L=L, kk=kk: f"=IF({kk(m)}=1,N({L}_Monto),0)*{hor(m)}", NF_USD, "sum"),
            (f"L{j}_int", f"Préstamo {j}: interés", "USD",
             lambda m, L=L, kk=kk, j=j: f"=IF({kk(m)}>=1,IF({kk(m)}=1,N({L}_Monto),{prv(K(f'L{j}_sal'), m)})*N({L}_Tasa)/12,0)*{hor(m)}", NF_USD, "sum"),
            (f"L{j}_am", f"Préstamo {j}: amortización", "USD",
             lambda m, L=L, kk=kk, j=j: (f"=IF(AND({kk(m)}>N({L}_Gracia),{kk(m)}<=N({L}_Gracia)+N({L}_Cuotas),N({L}_Cuotas)>0),"
                                         f"IF({kk(m)}=N({L}_Gracia)+N({L}_Cuotas),IF({kk(m)}=1,N({L}_Monto),{prv(K(f'L{j}_sal'), m)}),"
                                         f"PMT(N({L}_Tasa)/12,N({L}_Cuotas),-N({L}_Monto))-{a(f'L{j}_int', m)}),0)*{hor(m)}"), NF_USD, "sum"),
            (f"L{j}_sal", f"Préstamo {j}: saldo", "USD",
             lambda m, L=L, kk=kk, j=j: f"=IF({kk(m)}<1,0,IF({kk(m)}=1,N({L}_Monto),{prv(K(f'L{j}_sal'), m)})-{a(f'L{j}_am', m)})", NF_USD, "last"),
        ]
    rows.append(("sec", "Financiamiento y tesorería"))
    LC = f"{X}_LC"
    zr = (lambda m, f: f"=IF({A('real', m)}=1,0,{f})") if real else (lambda m, f: f"={f}")
    rows += [
        ("DESL", "Desembolsos de préstamos", "USD", lambda m: R("DESEMB", m, f"{a('L1_des', m)}+{a('L2_des', m)}+{a('L3_des', m)}"), NF_USD, "sum"),
        ("AML", "Amortización de préstamos (cronograma)", "USD",
         lambda m: R("AMORT", m, f"MAX(0,MIN({prv(K('DEUDA'), m)}-{prv(K('LCS'), m)}+{a('DESL', m)},{a('L1_am', m)}+{a('L2_am', m)}+{a('L3_am', m)}))"),
         NF_USD, "sum"),
        ("APORTE", "Aportes de socios", "USD", lambda m: R("APORTE", m, f"SUMIF(Ap_Mes{O},{m},Ap_Monto{O})*{hor(m)}"), NF_USD, "sum"),
        ("LCA", "Línea rotativa disponible (1/0)", "1/0",
         lambda m: f"=IF(AND(N({LC}_Lim)>0,{m}>=MAX(1,N({LC}_Desde)),{m}<=IF(N({LC}_Hasta)=0,Fin,N({LC}_Hasta))),1,0)", NF_N, None),
        ("LCI", "Línea rotativa: interés del mes (sobre el saldo anterior)", "USD",
         lambda m: zr(m, f"{prv(K('LCS'), m)}*N({LC}_Tasa)/12*{hor(m)}"), NF_USD, "sum"),
        ("CAJA0", "Caja antes de línea rotativa y dividendos", "USD",
         lambda m: (f"={prv(K('CAJA'), m)}+{a('CFO', m)}-{a('CAPEX', m)}+{a('DESL', m)}-{a('AML', m)}-{a('INT', m)}*{hor(m)}+{a('APORTE', m)}"), NF_USD, "min"),
        ("CMIN", "Caja mínima operativa", "USD", lambda m: f"={a('vtas', m)}*{X}_CajaMin/DiasMes*{hor(m)}", NF_USD, None),
        ("LCD", "Línea rotativa: uso", "USD",
         lambda m: zr(m, f"IF({a('LCA', m)}=1,MIN(MAX(0,N({LC}_Lim)-{prv(K('LCS'), m)}),MAX(0,{a('CMIN', m)}-{a('CAJA0', m)})),0)*{hor(m)}"), NF_USD, "sum"),
        ("LCR", "Línea rotativa: devolución (con excedente; total al vencer)", "USD",
         lambda m: zr(m, f"IF({a('LCA', m)}=1,MIN({prv(K('LCS'), m)},MAX(0,{a('CAJA0', m)}-{a('CMIN', m)})),{prv(K('LCS'), m)})*{hor(m)}"), NF_USD, "sum"),
        ("LCS", "Línea rotativa: saldo", "USD",
         lambda m: (f"=IF({A('real', m)}=1,IF({m}=Corte,N(V_LC_Saldo0),0),{prv(K('LCS'), m)}+{a('LCD', m)}-{a('LCR', m)})" if real
                    else f"={prv(K('LCS'), m)}+{a('LCD', m)}-{a('LCR', m)}"), NF_USD, "max"),
        ("DESEMB", "DESEMBOLSOS TOTALES (préstamos + línea)", "USD", lambda m: f"={a('DESL', m)}+{a('LCD', m)}", NF_USD, "sum"),
        ("AMORT", "AMORTIZACIONES TOTALES (préstamos + línea)", "USD", lambda m: f"={a('AML', m)}+{a('LCR', m)}", NF_USD, "sum"),
        ("DEUDA", "DEUDA FINANCIERA (saldo)", "USD", lambda m: f"={prv(K('DEUDA'), m)}+{a('DESEMB', m)}-{a('AMORT', m)}", NF_USD, "max", True),
        ("CAJA1", "Caja después de la línea rotativa", "USD", lambda m: f"={a('CAJA0', m)}+{a('LCD', m)}-{a('LCR', m)}", NF_USD, "min"),
        ("NIY", "Resultado neto acumulado del año", "USD",
         lambda m: f"=IF({A('mes', m)}=1,{a('NI', m)},{prv(K('NIY'), m)}+{a('NI', m)})", NF_USD, None),
        ("NIP", "Resultado neto del año anterior", "USD",
         lambda m: "=0" if m == 1 else f"=IF({A('mes', m)}=1,{prv(K('NIY'), m)},{prv(K('NIP'), m)})", NF_USD, None),
        ("DIVD", "Dividendos según política (resultado del año anterior)", "USD",
         lambda m: (f"=IF(AND({A('mes', m)}=MAX(1,{X}_MesDiv),{A('anio', m)}>=2),INDEX(Payout_{O},1,MIN({NY},{A('anio', m)}-1))"
                    f"*MAX(0,{a('NIP', m)}),0)*{hor(m)}"), NF_USD, "sum"),
        ("DIV", "Dividendos pagados (solo con caja sobre el mínimo)", "USD",
         lambda m: R("DIV", m, f"MIN({a('DIVD', m)},MAX(0,{a('CAJA1', m)}-{a('CMIN', m)}))"), NF_USD, "sum"),
        ("CAJA", "CAJA FINAL", "USD",
         lambda m: (f"=IF(AND({A('real', m)}=1,{rgiven('S_CAJA', m)}),{rv('S_CAJA', m)},{a('CAJA1', m)}-{a('DIV', m)})" if real
                    else f"={a('CAJA1', m)}-{a('DIV', m)}"), NF_USD, "min", True),
        ("AJ", "Ajuste de conciliación con la caja real", "USD", lambda m: f"={a('CAJA', m)}-({a('CAJA1', m)}-{a('DIV', m)})", NF_USD, "sum"),
        ("BRECHA", "Brecha: caja por debajo de la caja mínima", "USD", lambda m: f"=MAX(0,{a('CMIN', m)}-{a('CAJA', m)})*{hor(m)}", NF_USD, "max", True),
        ("DEF", "Caja negativa a cubrir por los socios (aporte implícito acumulado)", "USD", lambda m: f"=MAX(0,-{a('CAJA', m)})", NF_USD, "max"),
    ]
    rows.append(("sec", "Valuación y control"))
    rows += [
        ("DFP", "Factor de descuento del proyecto", "", lambda m: f"=1/(1+{X}_WACC)^(({m}-1)/12)", NF_USD4, None),
        ("DFE", "Factor de descuento de los socios", "", lambda m: f"=1/(1+{X}_Ke)^(({m}-1)/12)", NF_USD4, None),
        ("EQ", "Flujo de los socios (-aportes - cobertura de caja negativa + dividendos + valor final)", "USD",
         lambda m: (f"=-{a('APORTE', m)}-({a('DEF', m)}-{prv(K('DEF'), m)})+{a('DIV', m)}"
                    f"+IF({m}=Fin,{a('CAJA', m)}+{a('DEF', m)}+{a('VT', m)}-{a('DEUDA', m)},0)"), NF_USD, "sum"),
        ("CUMF", "FCFF acumulado", "USD", lambda m: f"={prv(K('CUMF'), m)}+{a('FCFF', m)}", NF_USD, "last"),
        ("PB", "Recupero de la inversión (mes con interpolación)", "mes",
         lambda m: "=0" if m == 1 else f"=IF(AND({prv(K('CUMF'), m)}<0,{a('CUMF', m)}>=0),{m - 2}+(-{prv(K('CUMF'), m)})/{a('FCFF', m)},0)", NF_USD2, "max"),
        ("PAT", "Patrimonio (aportes + resultados - dividendos + ajustes)", "USD",
         lambda m: f"={prv(K('PAT'), m)}+{a('APORTE', m)}+{a('NI', m)}-{a('DIV', m)}+{a('AJ', m)}", NF_USD, "last"),
        ("CHK", "CONTROL: activo - pasivo - patrimonio (0 = OK)", "USD",
         lambda m: (f"=ROUND({a('CAJA', m)}+{a('cxc', m)}+{a('inv', m)}+{a('ivab', m)}+{a('AF', m)}-{a('prov', m)}-{a('DEUDA', m)}-{a('PAT', m)},2)"),
         NF_USD2, "maxabs", True),
    ]
    return rows


def build_calculo():
    ws = WS["Calculo"]
    title(ws, "Cálculo — motor mensual (no editar)",
          "Bloque P = presupuesto original (solo supuestos originales). Bloque V = vigente: Reales hasta el último mes cerrado + proyección con escenario. Mismas fórmulas.")
    widths(ws, {"A": 9, "B": 60, "C": 7, "D": 3, "E": 15})
    for m in range(1, N + 1):
        ws.column_dimensions[mc(m)].width = 11
    # asignación de filas
    r = 4
    cal = [("m", "Mes n°", lambda m: m, NF_N), ("fecha", "Fecha", lambda m: f'=IF(n_Fecha1="","",EOMONTH(n_Fecha1,{m}-1))', NF_MES),
           ("anio", "Año del modelo", lambda m: f"=ROUNDUP({m}/12,0)", NF_N), ("mes", "Mes del año del modelo", lambda m: f"=MOD({m}-1,12)+1", NF_N),
           ("real", "Mes real cerrado (1/0)", lambda m: f"=IF({m}<=Corte,1,0)", NF_N), ("hor", "Dentro del horizonte (1/0)", lambda m: f"=IF({m}<=Fin,1,0)", NF_N),
           ("tcr", "TC para convertir reales (real del mes o vigente)", None, NF_INT)]
    for k, *_ in cal:
        ROW[k] = r
        r += 1
    r += 1
    blocks = {}
    for X in ("P", "V"):
        rows = version_rows(X)
        blocks[X] = (r, rows)
        r += 1   # título del bloque
        for it in rows:
            if it[0] == "sec":
                r += 1
            else:
                ROW[f"{X}.{it[0]}"] = r
                r += 1
        r += 2
    # escritura del calendario
    for k, lab, fn, fmt in cal:
        rr = ROW[k]
        put(ws, rr, 2, lab, font=F_H if k == "m" else F_B, fill=FL_H if k == "m" else None)
        put(ws, rr, 1, k, font=F_S)
        for m in range(1, N + 1):
            if k == "tcr":
                v = f"=IF(N(Reales!{mc(m)}{RR['TC']})>0,Reales!{mc(m)}{RR['TC']},INDEX(TC_V,1,MIN({NY},{A('anio', m)})))"
            else:
                v = fn(m)
            put(ws, rr, FCOL + m - 1, v, fmt=fmt, font=F_H if k == "m" else F, fill=FL_H if k == "m" else None)
    # bloques
    for X in ("P", "V"):
        start, rows = blocks[X]
        lab = ("BLOQUE P — PRESUPUESTO ORIGINAL" if X == "P" else "BLOQUE V — VIGENTE: REAL HASTA EL CORTE + PROYECCIÓN")
        for c in range(1, 6):
            ws.cell(row=start, column=c).fill = FL_P if X == "P" else FL_V
        put(ws, start, 2, lab, font=F_SEC, fill=FL_P if X == "P" else FL_V)
        r = start + 1
        for it in rows:
            if it[0] == "sec":
                put(ws, r, 2, it[1], font=F_SEC, fill=FL_SEC)
                for c in range(1, 6):
                    ws.cell(row=r, column=c).fill = FL_SEC
                r += 1
                continue
            key, lab, unit, fn, fmt, tot = it[:6]
            bold = len(it) > 6 and it[6]
            rr = ROW[f"{X}.{key}"]
            assert rr == r, (X, key, rr, r)
            put(ws, r, 1, f"{X}.{key}", font=F_S)
            put(ws, r, 2, lab, font=F_B if bold else (F_LINK if str(lab).startswith("=") else F), fill=FL_TOT if bold else None)
            put(ws, r, 3, unit)
            rng = f"F{r}:{MN}{r}"
            totf = {"sum": f"=SUM({rng})", "last": f"=INDEX({rng},1,MIN({N},Fin))", "max": f"=MAX({rng})", "min": f"=MIN({rng})",
                    "maxabs": f"=MAX(ABS(MAX({rng})),ABS(MIN({rng})))"}.get(tot)
            if totf:
                put(ws, r, 5, totf, fmt=fmt, font=F_B, fill=FL_TOT)
            for m in range(1, N + 1):
                put(ws, r, FCOL + m - 1, fn(m), fmt=fmt, font=F_B if bold else F)
            r += 1
    # contadores de Detalle_24m: celdas con dato por mes y bloque (alimentan la guía y Proyeccion_24m)
    r += 1
    put(ws, r, 2, "CONTADORES DE DETALLE_24M — celdas con dato en el mes (la guía y Proyeccion_24m los usan)", font=F_SEC, fill=FL_SEC)
    for c in range(1, 6):
        ws.cell(row=r, column=c).fill = FL_SEC
    r += 1
    for code, tit, *_ in DET_BLOQUES:
        ROW["det." + code] = r
        put(ws, r, 1, "det." + code, font=F_S)
        put(ws, r, 2, f"Celdas con dato — {tit}")
        put(ws, r, 5, f"=SUM(F{r}:{MN}{r})", fmt=NF_N, font=F_B, fill=FL_TOT)
        for m in range(1, N + 1):
            put(ws, r, FCOL + m - 1, f"=COUNT(Detalle_24m!{mc(m)}{DT[code + '_first']}:{mc(m)}{DT[code + '_last']})", fmt=NF_N)
        nm(f"Det_n{code}", "Calculo", f"{M1}{r}:{MN}{r}")
        r += 1
    ws.freeze_panes = f"F{ROW['real'] + 1}"


# ============================================================================ RESULTADOS
RS = {}


def build_resultados():
    ws = WS["Resultados"]
    CUR[0] = "Resultados"
    title(ws, "Resultados — presupuesto original vs vigente (real + proyección)")
    put(ws, 2, 1, '="Escenario de proyección: "&n_Escenario&"   |   Último mes cerrado: "&IF(Corte=0,"sin datos reales","mes "&Corte)', font=F_AL)
    widths(ws, {"A": 6, "B": 56, "C": 16, "D": 16, "E": 16, "F": 60})
    for c in range(4, 16):
        ws.column_dimensions[CL(c)].width = 13
    YR = f"Calculo!{RG('anio')}"
    r = 4
    section(ws, r, "A. Indicadores clave", 15); r += 1
    header(ws, r, ["", "Indicador", "Presupuesto original", "Vigente (real + proyección)", "Desvío", "Lectura"], 1, 30)
    r += 1

    def g(X, k):
        return f"Calculo!{RG(X + '.' + k)}"

    def tot(X, k):
        return f"Calculo!$E${ROW[X + '.' + k]}"

    def irr(X, k):
        rg = g(X, k)
        return (f'IF(COUNTIF({rg},"<0")=0,"No existe",IF(COUNTIF({rg},">0")=0,"No existe",'
                f'IFERROR((1+IRR({rg},0.01))^12-1,IFERROR((1+IRR({rg},-0.01))^12-1,"No converge"))))')

    kpis = [
        ("v1", "Ventas del año 1 del modelo", lambda X: f"=SUMIFS({g(X, 'vtas')},{YR},1)", NF_USD, ""),
        ("vt", "Ventas totales del horizonte", lambda X: f"={tot(X, 'vtas')}", NF_USD, ""),
        ("e1", "EBITDA del año 1 del modelo", lambda X: f"=SUMIFS({g(X, 'EBITDA')},{YR},1)", NF_USD, ""),
        ("et", "EBITDA total del horizonte", lambda X: f"={tot(X, 'EBITDA')}", NF_USD, ""),
        ("em", "Margen EBITDA del horizonte", lambda X: f"=IF({tot(X, 'vtas')}=0,\"\",{tot(X, 'EBITDA')}/{tot(X, 'vtas')})", NF_PCT, ""),
        ("nt", "Resultado neto total", lambda X: f"={tot(X, 'NI')}", NF_USD, ""),
        ("capex", "Inversión total (CAPEX)", lambda X: f"={tot(X, 'CAPEX')}", NF_USD, "Vigente: real pagado + pendiente con escenario."),
        ("nec", "Necesidad máxima de fondos antes de financiamiento", lambda X: f"=MAX(0,-MIN({g(X, 'CUML')}))", NF_USD,
         "Inversión + capital de trabajo + pérdidas iniciales que deben financiar socios y bancos."),
        ("necm", "Mes de la necesidad máxima", lambda X: f'=IF(MIN({g(X, "CUML")})>=0,"",MATCH(MIN({g(X, "CUML")}),{g(X, "CUML")},0))', NF_N, ""),
        ("ap", "Aportes de socios", lambda X: f"={tot(X, 'APORTE')}", NF_USD, ""),
        ("apimp", "Aporte adicional necesario para no tener caja negativa", lambda X: f"={tot(X, 'DEF')}", NF_USD,
         "Fondos no previstos que deberían aportar socios o bancos (máximo saldo de caja negativa)."),
        ("deuda", "Deuda financiera máxima", lambda X: f"={tot(X, 'DEUDA')}", NF_USD, "Préstamos + línea rotativa."),
        ("lcmax", "Línea rotativa: uso máximo", lambda X: f"={tot(X, 'LCS')}", NF_USD, "0 = no se usó o no hay línea."),
        ("intt", "Intereses, IVA y comisiones totales", lambda X: f"={tot(X, 'INT')}", NF_USD, ""),
        ("canc", "Cancelación total de la deuda (años desde el mes 1)",
         lambda X: (f'=IF(MAX({g(X, "DEUDA")})<1,"Sin deuda",IF(INDEX({g(X, "DEUDA")},1,MIN({N},Fin))>1,"No cancela en el horizonte",'
                    f'(_xlfn.MAXIFS(Calculo!{RG("m")},{g(X, "DEUDA")},">1")+1)/12))'), NF_USD2, "Último mes con saldo de deuda + 1, en años."),
        ("cajamin", "Caja final mínima del horizonte", lambda X: f"=MIN({g(X, 'CAJA')})", NF_USD, "Negativa = financiamiento no previsto."),
        ("brecha", "Brecha máxima bajo la caja mínima", lambda X: f"={tot(X, 'BRECHA')}", NF_USD, "No se insertan fondos ficticios: la brecha queda visible."),
        ("brechan", "Meses con brecha", lambda X: f'=SUMPRODUCT(--({g(X, "BRECHA")}>0.5))', NF_N, ""),
        ("dscr", "DSCR mínimo desde el 2.º año de operación (flujo operativo / servicio de deuda)", None, NF_X,
         "Servicio = intereses + cuotas de préstamos (sin devoluciones de la línea). < 1 = la operación no alcanza a pagar la deuda ese año."),
        ("van", "VAN del proyecto (FCFF a WACC)", lambda X: f"=SUMPRODUCT({g(X, 'FCFF')},{g(X, 'DFP')})", NF_USD, "Excluye préstamos, intereses, aportes y dividendos."),
        ("tir", "TIR del proyecto (anual)", lambda X: "=" + irr(X, "FCFF"), NF_PCT2, ""),
        ("vane", "VAN de los socios (a Ke)", lambda X: f"=SUMPRODUCT({g(X, 'EQ')},{g(X, 'DFE')})", NF_USD,
         "Aportes, dividendos pagados y valor final. La caja negativa se considera cubierta por los socios hasta que se recupera."),
        ("tire", "TIR de los socios (anual)", lambda X: "=" + irr(X, "EQ"), NF_PCT2, ""),
        ("pb", "Recupero de la inversión del proyecto (años)", lambda X: f'=IF(MAX({g(X, "PB")})=0,"No recupera",MAX({g(X, "PB")})/12)', NF_USD2, ""),
        ("chk", "Balance de control (0 = OK)", lambda X: f"={tot(X, 'CHK')}", NF_USD2, ""),
    ]
    for key, lab, fn, fmt, note in kpis:
        RS["k_" + key] = r
        put(ws, r, 2, lab)
        for col, X in [(3, "P"), (4, "V")]:
            if fn:
                put(ws, r, col, fn(X), fmt=fmt, fill=FL_P if X == "P" else FL_V, font=F_B)
        put(ws, r, 5, f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),D{r}-C{r},"")', fmt=fmt)
        put(ws, r, 6, note, font=F_S)
        r += 1
    nm("r_BrechaV", "Resultados", f"D{RS['k_brecha']}")
    nm("r_ChkP", "Resultados", f"C{RS['k_chk']}")
    nm("r_ChkV", "Resultados", f"D{RS['k_chk']}")
    r += 1
    put(ws, r, 2, "Plan publicado (referencia): VAN / TIR", font=F_S)
    put(ws, r, 3, '=IF(pub_VAN="","",pub_VAN)', fmt=NF_USD)
    put(ws, r, 4, '=IF(pub_TIR="","",pub_TIR)', fmt=NF_PCT2)
    put(ws, r, 6, "Cifras del documento original, calculadas con su propia metodología.", font=F_S)
    r += 2

    # ---------------------------------------------------------- tablas anuales
    def yhdr(rr, lab):
        header(ws, rr, ["", lab, "Total"] + [f"Año {y}" for y in range(1, NYR + 1)], 1)
        put(ws, rr + 1, 2, "Estado del año", font=F_S)
        put(ws, rr + 2, 2, "Año calendario", font=F_S)
        for y in range(1, NYR + 1):
            c = 3 + y
            put(ws, rr + 1, c, f'=IF(Corte>={12 * y},"Real",IF(Corte>{12 * (y - 1)},"Real+Proy.","Proyección"))', font=F_S, al=CEN)
            put(ws, rr + 2, c, f'=IF(n_Fecha1="","",YEAR(EOMONTH(n_Fecha1,{12 * (y - 1)})))', font=F_S, fmt=NF_N, al=CEN)
        return rr + 3

    def arow(rr, key, lab, X, src, kind="sum", fmt=NF_USD, bold=False):
        RS[key] = rr
        put(ws, rr, 1, key, font=F_S)
        put(ws, rr, 2, lab, font=F_B if bold else F, fill=FL_TOT if bold else None)
        for y in range(1, NYR + 1):
            c = 3 + y
            if kind == "sum":
                f = f"=SUMIFS({g(X, src)},{YR},{y})"
            elif kind == "end":
                f = f"=INDEX({g(X, src)},1,MIN({N},{12 * y}))"
            elif kind == "max":
                f = f"=_xlfn.MAXIFS({g(X, src)},{YR},{y})"
            else:
                f = src(y, CL(c))
            put(ws, rr, c, f, fmt=fmt, font=F_B if bold else F)
        if kind == "sum":
            put(ws, rr, 3, f"=SUM(D{rr}:O{rr})", fmt=fmt, font=F_B, fill=FL_TOT)
        elif kind in ("end",):
            put(ws, rr, 3, f"=INDEX({g(X, src)},1,MIN({N},Fin))", fmt=fmt, font=F_B, fill=FL_TOT)
        elif kind == "max":
            put(ws, rr, 3, f"=MAX(D{rr}:O{rr})", fmt=fmt, font=F_B, fill=FL_TOT)
        return rr + 1

    section(ws, r, "B. Vigente (real hasta el corte + proyección) — resumen anual", 15); r += 1
    r = yhdr(r, "Concepto (USD)")
    lines = [("V_vtas", "Ventas", "vtas", "sum", True), ("V_MAT", "Materia prima e insumos", "MAT", "sum", False),
             ("V_OVAR", "Otros costos variables", "OVAR", "sum", False), ("V_PERP", "Personal de producción", "PERP", "sum", False),
             ("V_FIJP", "Costos fijos de producción e imprevistos", "FIJP", "sum", False), ("V_MB", "Margen bruto", "MB", "sum", True),
             ("V_PERA", "Personal administrativo", "PERA", "sum", False), ("V_ADM", "Gastos de administración", "ADM", "sum", False),
             ("V_TRIB", "Tributo de maquila", "TRIB", "sum", False), ("V_IVAR", "(+) IVA recuperable", "IVAR", "sum", False),
             ("V_EBITDA", "EBITDA", "EBITDA", "sum", True), ("V_DEP", "Depreciación", "DEP", "sum", False), ("V_INT", "Intereses, IVA y comisiones", "INT", "sum", False),
             ("V_IRE", "Impuesto a la renta", "IRE", "sum", False), ("V_NI", "Resultado neto", "NI", "sum", True),
             ("V_CFO", "Flujo operativo", "CFO", "sum", True), ("V_CAPEX", "CAPEX pagado", "CAPEX", "sum", False), ("V_FCFF", "FCFF (incluye valor terminal)", "FCFF", "sum", True),
             ("V_APORTE", "Aportes de socios", "APORTE", "sum", False), ("V_DESEMB", "Desembolsos de préstamos", "DESEMB", "sum", False),
             ("V_AMORT", "Amortizaciones (préstamos + línea)", "AMORT", "sum", False), ("V_AML", "   de las cuales: cuotas de préstamos", "AML", "sum", False),
             ("V_DIV", "Dividendos pagados", "DIV", "sum", False),
             ("V_CAJA", "Caja al cierre", "CAJA", "end", True), ("V_DEUDA", "Deuda al cierre", "DEUDA", "end", False), ("V_LCS", "   de la cual: línea rotativa", "LCS", "end", False),
             ("V_CTN", "Capital de trabajo neto al cierre", "ctn", "end", False), ("V_BRECHA", "Brecha máxima del año", "BRECHA", "max", False)]
    for key, lab, src, kind, bold in lines:
        r = arow(r, key, lab, "V", src, kind, bold=bold)
    r = arow(r, "V_mg", "Margen EBITDA", "V", lambda y, c: f'=IF({c}{RS["V_vtas"]}=0,"",{c}{RS["V_EBITDA"]}/{c}{RS["V_vtas"]})', "f", NF_PCT)
    r = arow(r, "V_DS", "Servicio de deuda (intereses + amortización)", "V", lambda y, c: f"={c}{RS['V_INT']}+{c}{RS['V_AML']}", "f")
    r = arow(r, "V_DSCR", "DSCR (flujo operativo / servicio de deuda)", "V",
             lambda y, c: f'=IF({c}{RS["V_DS"]}>1,{c}{RS["V_CFO"]}/{c}{RS["V_DS"]},"")', "f", NF_X)
    r = arow(r, "YN", "Año n° (auxiliar)", "V", lambda y, c: f"={y}", "f", NF_N)
    put(ws, r, 2, "Unidades vendidas por producto", font=F_B)
    r += 1
    for i in range(NP):
        r = arow(r, f"V_u{i + 1}", f'=IF(INDEX(Prod_Nombre,{i + 1})="","Producto {i + 1}",INDEX(Prod_Nombre,{i + 1}))', "V", f"u{i + 1}", "sum", NF_INT)
    r += 1
    section(ws, r, "C. Presupuesto original — resumen anual", 15); r += 1
    r = yhdr(r, "Concepto (USD)")
    for key, lab, src, kind, bold in [("P_vtas", "Ventas", "vtas", "sum", True), ("P_EBITDA", "EBITDA", "EBITDA", "sum", True), ("P_NI", "Resultado neto", "NI", "sum", True),
                                      ("P_CFO", "Flujo operativo", "CFO", "sum", False), ("P_CAPEX", "CAPEX", "CAPEX", "sum", False),
                                      ("P_FCFF", "FCFF (incluye valor terminal)", "FCFF", "sum", True), ("P_APORTE", "Aportes de socios", "APORTE", "sum", False),
                                      ("P_INT", "Intereses, IVA y comisiones", "INT", "sum", False), ("P_AMORT", "Amortizaciones (préstamos + línea)", "AMORT", "sum", False),
                                      ("P_AML", "   de las cuales: cuotas de préstamos", "AML", "sum", False),
                                      ("P_CAJA", "Caja al cierre", "CAJA", "end", True), ("P_DEUDA", "Deuda al cierre", "DEUDA", "end", False),
                                      ("P_CTN", "Capital de trabajo neto al cierre", "ctn", "end", False)]:
        r = arow(r, key, lab, "P", src, kind, bold=bold)
    r = arow(r, "P_DS", "Servicio de deuda", "P", lambda y, c: f"={c}{RS['P_INT']}+{c}{RS['P_AML']}", "f")
    r = arow(r, "P_DSCR", "DSCR", "P", lambda y, c: f'=IF({c}{RS["P_DS"]}>1,{c}{RS["P_CFO"]}/{c}{RS["P_DS"]},"")', "f", NF_X)
    put(ws, r, 2, "Unidades vendidas por producto", font=F_B)
    r += 1
    for i in range(NP):
        r = arow(r, f"P_u{i + 1}", f'=IF(INDEX(Prod_Nombre,{i + 1})="","Producto {i + 1}",INDEX(Prod_Nombre,{i + 1}))', "P", f"u{i + 1}", "sum", NF_INT)
    r += 1
    section(ws, r, "D. Desvíos: vigente - presupuesto original", 15); r += 1
    r = yhdr(r, "Concepto (USD)")
    for key, lab, kv, kp in [("D_vtas", "Ventas", "V_vtas", "P_vtas"), ("D_EBITDA", "EBITDA", "V_EBITDA", "P_EBITDA"), ("D_NI", "Resultado neto", "V_NI", "P_NI"),
                             ("D_CAPEX", "CAPEX", "V_CAPEX", "P_CAPEX"), ("D_FCFF", "FCFF", "V_FCFF", "P_FCFF"), ("D_CAJA", "Caja al cierre", "V_CAJA", "P_CAJA")]:
        RS[key] = r
        put(ws, r, 2, lab)
        for y in range(0, NYR + 1):
            c = CL(3 + y)
            put(ws, r, 3 + y, f"={c}{RS[kv]}-{c}{RS[kp]}", fmt=NF_USD)
        r += 1
    RS["D_u"] = r
    put(ws, r, 2, "Unidades totales")
    for y in range(0, NYR + 1):
        c = CL(3 + y)
        put(ws, r, 3 + y, f"=SUM({c}{RS['V_u1']}:{c}{RS['V_u' + str(NP)]})-SUM({c}{RS['P_u1']}:{c}{RS['P_u' + str(NP)]})", fmt=NF_INT)
    r += 1
    put(ws, r, 2, "% de desvío en ventas")
    for y in range(0, NYR + 1):
        c = CL(3 + y)
        put(ws, r, 3 + y, f'=IF({c}{RS["P_vtas"]}=0,"",{c}{RS["D_vtas"]}/{c}{RS["P_vtas"]})', fmt=NF_PCT)
    r += 2
    section(ws, r, "E. Plan publicado (referencia) vs presupuesto original del modelo — explica diferencias de metodología", 15); r += 1
    r = yhdr(r, "Concepto (USD)")
    for k, lab, kp in [("Ventas", "Ventas publicadas", "P_vtas"), ("EBITDA", "EBITDA publicado", "P_EBITDA"), ("Utilidad", "Resultado neto publicado", "P_NI")]:
        RS["pub_" + k] = r
        put(ws, r, 2, lab)
        for y in range(1, NY + 1):
            put(ws, r, 3 + y, f'=IF(INDEX(pub_{k},1,{y})="","",INDEX(pub_{k},1,{y}))', fmt=NF_USD, font=F_LINK)
        put(ws, r, 3, f"=SUM(D{r}:M{r})", fmt=NF_USD, font=F_B, fill=FL_TOT)
        r += 1
        put(ws, r, 2, f"   Diferencia presupuesto del modelo - publicado")
        for y in range(1, NY + 1):
            c = CL(3 + y)
            put(ws, r, 3 + y, f'=IF({c}{r - 1}="","",{c}{RS[kp]}-{c}{r - 1})', fmt=NF_USD)
        r += 1
    # DSCR mínimo
    for col, k in [(3, "P_DSCR"), (4, "V_DSCR")]:
        crit = f'D{RS[k[0] + "_DS"]}:O{RS[k[0] + "_DS"]},">1",D{RS["YN"]}:O{RS["YN"]},">="&(ROUNDUP({k[0]}_Inicio/12,0)+1)'
        put(ws, RS["k_dscr"], col, f'=IF(COUNTIFS({crit})=0,"Sin deuda",_xlfn.MINIFS(D{RS[k]}:O{RS[k]},{crit}))',
            fmt=NF_X, fill=FL_P if col == 3 else FL_V, font=F_B)
    add_charts(ws, r + 2)
    ws.freeze_panes = "C6"


def add_charts(ws, anchor):
    from openpyxl.chart import BarChart, LineChart, Reference
    from openpyxl.chart.series import SeriesLabel
    lab_row = anchor
    for y in range(1, NYR + 1):
        ws.cell(row=lab_row, column=3 + y, value=f"Año {y}")
    put(ws, lab_row, 2, "Etiquetas de gráficos", font=F_S)
    cats = Reference(ws, min_col=4, max_col=3 + NYR, min_row=lab_row)

    def ref(key):
        return Reference(ws, min_col=4, max_col=3 + NYR, min_row=RS[key])

    ch = BarChart()
    ch.title = "Ventas: presupuesto vs vigente"
    ch.height, ch.width = 7.5, 17
    for k, n in [("P_vtas", "Presupuesto"), ("V_vtas", "Vigente")]:
        ch.add_data(ref(k), from_rows=True, titles_from_data=False)
        ch.series[-1].tx = SeriesLabel(v=n)
    ch.set_categories(cats)
    ws.add_chart(ch, f"B{anchor + 2}")
    ch2 = LineChart()
    ch2.title = "EBITDA y caja al cierre"
    ch2.height, ch2.width = 7.5, 17
    for k, n in [("P_EBITDA", "EBITDA presupuesto"), ("V_EBITDA", "EBITDA vigente"), ("P_CAJA", "Caja presupuesto"), ("V_CAJA", "Caja vigente")]:
        ch2.add_data(ref(k), from_rows=True, titles_from_data=False)
        ch2.series[-1].tx = SeriesLabel(v=n)
        ch2.series[-1].smooth = False
    ch2.set_categories(cats)
    ch2.x_axis.tickLblPos = "low"
    ws.add_chart(ch2, f"F{anchor + 2}")


# ============================================================================ PROYECCION 24 MESES
NV = 24


def build_proy():
    ws = WS["Proyeccion_24m"]
    title(ws, "Proyección de los próximos 24 meses — vigente (real hasta el corte + detalle mensual + supuestos anuales)")
    put(ws, 2, 1, '="Meses "&(Corte+1)&" a "&MIN(144,Corte+24)&" del modelo   |   Escenario: "&n_Escenario&"   |   Ventana de detalle: "&n_Ventana&" meses"', font=F_AL)
    widths(ws, {"A": 8, "B": 44, "C": 13, "D": 13, "E": 13})
    for k in range(1, NV + 1):
        ws.column_dimensions[CL(5 + k)].width = 11
    r = 4
    heads = [("Mes de la proyección", lambda k: k, NF_N), ("Mes del modelo", lambda k: f"=Corte+{k}", NF_N),
             ("Fecha", lambda k: f'=IF(OR(n_Fecha1="",Corte+{k}>{N}),"",EOMONTH(n_Fecha1,Corte+{k}-1))', NF_MES),
             ("Origen de volúmenes", lambda k: f'=IF(Corte+{k}>{N},"",IF(AND({k}<=n_Ventana,INDEX(Det_nQ,1,Corte+{k})>0),"Detalle","Anual"))', None),
             ("Origen de costos", lambda k: (f'=IF(Corte+{k}>{N},"",IF(AND({k}<=n_Ventana,INDEX(Det_nMP,1,Corte+{k})+INDEX(Det_nIN,1,Corte+{k})+INDEX(Det_nOV,1,Corte+{k})'
                                              f'+INDEX(Det_nDO,1,Corte+{k})+INDEX(Det_nSA,1,Corte+{k})+INDEX(Det_nFC,1,Corte+{k})>0),"Detalle","Anual"))'), None)]
    for j, (lab, fn, fmt) in enumerate(heads):
        put(ws, r + j, 2, lab, font=F_H, fill=FL_H)
        for c in (1, 3, 4, 5):
            ws.cell(row=r + j, column=c).fill = FL_H
        for k in range(1, NV + 1):
            put(ws, r + j, 5 + k, fn(k), font=F_H, fill=FL_H, fmt=fmt, al=CEN)
    put(ws, r, 3, "Año 1 ventana", font=F_H, fill=FL_H, al=CEN)
    put(ws, r, 4, "Año 2 ventana", font=F_H, fill=FL_H, al=CEN)
    put(ws, r, 5, "Total 24 meses", font=F_H, fill=FL_H, al=CEN)
    r += len(heads) + 1
    c1, c12, c13, c24 = CL(6), CL(17), CL(18), CL(29)

    def linea(rr, key, lab, src, kind="sum", fmt=NF_USD, bold=False):
        put(ws, rr, 1, key, font=F_S)
        put(ws, rr, 2, lab, font=F_B if bold else F, fill=FL_TOT if bold else None)
        for k in range(1, NV + 1):
            put(ws, rr, 5 + k, f'=IF(Corte+{k}>{N},"",INDEX(Calculo!{RG(src)},1,Corte+{k}))', fmt=fmt, font=F_B if bold else F)
        if kind == "sum":
            for col, a1, a2 in [(3, c1, c12), (4, c13, c24), (5, c1, c24)]:
                put(ws, rr, col, f"=SUM({a1}{rr}:{a2}{rr})", fmt=fmt, font=F_B, fill=FL_TOT)
        elif kind == "end":
            for col, cc in [(3, c12), (4, c24), (5, c24)]:
                put(ws, rr, col, f"={cc}{rr}", fmt=fmt, font=F_B, fill=FL_TOT)
        elif kind == "max":
            for col, a1, a2 in [(3, c1, c12), (4, c13, c24), (5, c1, c24)]:
                put(ws, rr, col, f"=MAX({a1}{rr}:{a2}{rr})", fmt=fmt, font=F_B, fill=FL_TOT)

    section(ws, r, "A. Vigente: resultado, caja y financiamiento", 29); r += 1
    lineas = [("V.vtas", "Ventas", "sum", True), ("V.MAT", "Materia prima e insumos", "sum", False), ("V.OVAR", "Otros costos variables", "sum", False),
              ("V.PERP", "Personal de producción", "sum", False), ("V.FIJP", "Costos fijos de producción e imprevistos", "sum", False),
              ("V.PERA", "Personal administrativo", "sum", False), ("V.ADM", "Gastos de administración", "sum", False), ("V.TRIB", "Tributo de maquila", "sum", False),
              ("V.EBITDA", "EBITDA", "sum", True), ("V.DEP", "Depreciación", "sum", False), ("V.INT", "Intereses, IVA y comisiones", "sum", False),
              ("V.NI", "Resultado neto", "sum", True), ("V.CFO", "Flujo operativo", "sum", True), ("V.CAPEX", "CAPEX pagado", "sum", False),
              ("V.DESEMB", "Desembolsos (préstamos + línea)", "sum", False), ("V.AMORT", "Amortizaciones (préstamos + línea)", "sum", False),
              ("V.APORTE", "Aportes de socios", "sum", False), ("V.DIV", "Dividendos pagados", "sum", False),
              ("V.CAJA", "Caja al cierre", "end", True), ("V.DEUDA", "Deuda al cierre", "end", False), ("V.LCS", "Línea rotativa: saldo", "end", False),
              ("V.ctn", "Capital de trabajo neto", "end", False), ("V.BRECHA", "Brecha bajo la caja mínima (máx.)", "max", False)]
    PY = {}
    for src, lab, kind, bold in lineas:
        PY[src] = r
        linea(r, src, lab, src, kind, bold=bold)
        r += 1
    PY["mg"] = r
    put(ws, r, 2, "Margen EBITDA")
    for col in range(3, 6 + NV):
        cc = CL(col)
        put(ws, r, col, f'=IF(N({cc}{PY["V.vtas"]})=0,"",{cc}{PY["V.EBITDA"]}/{cc}{PY["V.vtas"]})', fmt=NF_PCT)
    r += 2
    section(ws, r, "B. Presupuesto original en los mismos meses y desvío", 29); r += 1
    for src, lab, kind, bold in [("P.vtas", "Ventas — presupuesto", "sum", False), ("P.EBITDA", "EBITDA — presupuesto", "sum", False),
                                 ("P.NI", "Resultado neto — presupuesto", "sum", False), ("P.CAJA", "Caja al cierre — presupuesto", "end", False)]:
        PY[src] = r
        linea(r, src, lab, src, kind, bold=bold)
        r += 1
    for lab, v, pp, kind in [("Desvío ventas (vigente − presupuesto)", "V.vtas", "P.vtas", "sum"), ("Desvío EBITDA", "V.EBITDA", "P.EBITDA", "sum"),
                             ("Desvío resultado neto", "V.NI", "P.NI", "sum"), ("Desvío caja al cierre", "V.CAJA", "P.CAJA", "end")]:
        put(ws, r, 2, lab, font=F_B)
        for col in range(3, 6 + NV):
            cc = CL(col)
            put(ws, r, col, f'=IF(OR({cc}{PY[v]}="",{cc}{PY[pp]}=""),"",{cc}{PY[v]}-{cc}{PY[pp]})', fmt=NF_USD, font=F_B)
        r += 1
    r += 1
    put(ws, r, 2, "Fuente: hoja Calculo (bloques V y P). Meses reales = Reales; ventana = Detalle_24m donde hay dato; resto = supuestos anuales.", font=F_S)
    nm("py_Brecha", "Proyeccion_24m", f"E{PY['V.BRECHA']}")
    from openpyxl.chart import LineChart, Reference
    from openpyxl.chart.series import SeriesLabel
    ch = LineChart()
    ch.title = "Próximos 24 meses: EBITDA y caja (vigente)"
    ch.height, ch.width = 7.5, 22
    for key, n in [("V.EBITDA", "EBITDA"), ("V.CAJA", "Caja al cierre"), ("P.CAJA", "Caja presupuesto")]:
        ch.add_data(Reference(ws, min_col=6, max_col=5 + NV, min_row=PY[key]), from_rows=True, titles_from_data=False)
        ch.series[-1].tx = SeriesLabel(v=n)
        ch.series[-1].smooth = False
    ch.set_categories(Reference(ws, min_col=6, max_col=5 + NV, min_row=4))
    ch.x_axis.tickLblPos = "low"
    ws.add_chart(ch, f"B{r + 2}")
    ws.freeze_panes = "F10"


# ============================================================================ CONTROL
def build_control():
    ws = WS["Control"]
    title(ws, "Control — validaciones y datos faltantes", "Estado OK cuando el valor está dentro de lo esperado. Revisar cada cierre.")
    widths(ws, {"A": 6, "B": 70, "C": 16, "D": 14, "E": 12, "F": 80})
    r = 4
    header(ws, r, ["Cód.", "Control", "Valor", "Esperado", "Estado", "Qué hacer si no está OK"], 1)
    r += 1
    first = r
    real_m = f"(Calculo!{RG('real')}=1)"
    checks = [
        ("Balance del presupuesto cuadra (activo - pasivo - patrimonio)", "=r_ChkP", "= 0", "abs0", "Error de fórmula: no editar Calculo; revisar entradas en blanco."),
        ("Balance vigente cuadra", "=r_ChkV", "= 0", "abs0", "Idem."),
        ("Meses cerrados sin ventas cargadas", f"=SUMPRODUCT({real_m}*(Calculo!{RG('V.vtas')}=0))", "= 0", "zero",
         "Cargar ventas reales del mes en Reales (si realmente no hubo ventas, cargar 0 explícito y documentarlo)."),
        ("Meses cerrados sin costos de materiales ni personal cargados",
         f"=SUMPRODUCT({real_m}*(Calculo!{RG('V.MAT')}+Calculo!{RG('V.PERP')}+Calculo!{RG('V.PERA')}=0))", "= 0", "zero", "Cargar costos reales del mes."),
        ("Saldo de caja real cargado en el mes de corte", f'=IF(Corte=0,1,IF(ISNUMBER(INDEX(Reales!$F${RR["S_CAJA"]}:${MN}${RR["S_CAJA"]},1,Corte)),1,0))', "= 1", "one",
         "Cargar el saldo de caja y bancos al cierre del último mes cerrado."),
        ("Saldos de clientes, inventarios y proveedores cargados en el mes de corte",
         f'=IF(Corte=0,1,IF(AND(ISNUMBER(INDEX(Reales!$F${RR["S_CXC"]}:${MN}${RR["S_CXC"]},1,Corte)),ISNUMBER(INDEX(Reales!$F${RR["S_INV"]}:${MN}${RR["S_INV"]},1,Corte)),'
         f'ISNUMBER(INDEX(Reales!$F${RR["S_PROV"]}:${MN}${RR["S_PROV"]},1,Corte))),1,0))', "= 1", "one", "Sin estos saldos el capital de trabajo se estima por días."),
        ("Diferencia de deuda: real al corte - calculada",
         f'=IF(Corte=0,0,IF(ISNUMBER(INDEX(Reales!$F${RR["S_DEUDA"]}:${MN}${RR["S_DEUDA"]},1,Corte)),INDEX(Reales!$F${RR["S_DEUDA"]}:${MN}${RR["S_DEUDA"]},1,Corte)'
         f'-INDEX(Calculo!{RG("V.DEUDA")},1,Corte),0))', "= 0", "abs1", "Cargar desembolsos y amortizaciones reales o ajustar las condiciones vigentes del préstamo."),
        ("CAPEX: pagado por ítem (Inversion) - CAPEX real mensual (Reales)", f"=i_Pagado-SUM(Reales!$F${RR['CAPEX']}:${MN}${RR['CAPEX']})", "= 0", "abs1",
         "Ambos registros deben conciliar."),
        ("Tipo de cambio cargado en todos los años (presupuesto y vigente)", "=COUNTIF(TC_O,\">0\")+COUNTIF(TC_V,\">0\")", f"= {2 * NY}", "tc",
         "Completar el TC por año en Supuestos B (los costos en PYG se convierten con él)."),
        ("Productos con volumen vigente por encima de la capacidad (años)", "=SUM(c_CapExceso)", "= 0", "zero",
         "Revisar volumen o registrar la ampliación de capacidad (Productos A)."),
        ("Productos con volumen y sin precio vigente", "=SUMPRODUCT((Precio_V=0)*(VolTot_V>0))", "= 0", "zero", "Completar el precio."),
        ("Préstamos del presupuesto cancelados dentro del horizonte", "=r_DeudaFinP", "= 0", "abs1", "Revisar plazo o el horizonte."),
        ("Brecha de caja vigente (meses bajo la caja mínima)", f'=SUMPRODUCT(--(Calculo!{RG("V.BRECHA")}>0.5))', "= 0", "zero",
         "Definir aportes, financiamiento o mejoras de capital de trabajo (alerta, no error de fórmula)."),
        ("Negocio en operación: mes real de inicio y último mes cerrado cargados", '=IF(n_Etapa="Operación",IF(AND(ISNUMBER(n_InicioReal),Corte>0),1,0),1)',
         "= 1", "one", "Completar en Inicio el mes real de inicio de operación y cargar los meses cerrados en Reales."),
        ("Horizonte dentro de la rejilla de 144 meses (mes de inicio ≤ 25 con 10 años)", f"=IF(MAX(Fin,V_Inicio)<={N},1,0)", "= 1", "one",
         "Reducir los años evaluados o el mes de inicio (la rejilla tiene 144 meses)."),
        ("Supuestos vigentes modificados respecto del presupuesto (informativo)", "=c_Modificados", "≥ 0", "info", "Documentar en la nota de cada supuesto el motivo del cambio."),
    ]
    for i, (lab, f, exp, kind, todo) in enumerate(checks):
        put(ws, r, 1, f"C{i + 1:02d}")
        put(ws, r, 2, lab)
        put(ws, r, 3, f, fmt=NF_USD2)
        put(ws, r, 4, exp, al=CEN)
        cond = {"abs0": f"ABS(C{r})<0.01", "abs1": f"ABS(C{r})<1", "zero": f"C{r}=0", "one": f"C{r}=1", "tc": f"C{r}={2 * NY}", "info": "TRUE"}[kind]
        put(ws, r, 5, f'=IF({cond},"OK","REVISAR")', font=F_B, al=CEN)
        nm(f"c_C{i + 1:02d}", "Control", f"E{r}")
        put(ws, r, 6, todo, font=F_S)
        r += 1
    last = r - 1
    ws.conditional_formatting.add(f"E{first}:E{last}", CellIsRule(operator="equal", formula=['"OK"'], fill=FL_OK))
    ws.conditional_formatting.add(f"E{first}:E{last}", CellIsRule(operator="equal", formula=['"REVISAR"'], fill=FL_BAD))
    r += 1
    put(ws, r, 2, "Controles en REVISAR", font=F_B)
    put(ws, r, 3, f'=COUNTIF(E{first}:E{last},"REVISAR")', fmt=NF_N, font=F_B)
    nm("c_Alertas", "Control", f"C{r}")
    r += 1
    put(ws, r, 2, "Celdas vigentes sobrescritas (difieren del presupuesto)")
    put(ws, r, 3, (f"=SUMPRODUCT(--(Supuestos!$E$6:$E${SUPR['p_last']}<>Supuestos!$D$6:$D${SUPR['p_last']}))"
                   f"+SUMPRODUCT(--(Precio_V<>Precio_O))+SUMPRODUCT(--(MP_V<>MP_O))+SUMPRODUCT(--(INS_V<>INS_O))+SUMPRODUCT(--(OV_V<>OV_O))"
                   f"+SUMPRODUCT(--(Vol_V<>Vol_O))+SUMPRODUCT(--(Sal_V<>Sal_O))+SUMPRODUCT(--(Dot_V<>Dot_O))+SUMPRODUCT(--(FC_V<>FC_O))"
                   f"+SUMPRODUCT(--(TC_V<>TC_O))+SUMPRODUCT(--(dPrecio_V<>dPrecio_O))+SUMPRODUCT(--(dMP_V<>dMP_O))+SUMPRODUCT(--(dOV_V<>dOV_O))"
                   f"+SUMPRODUCT(--(dSal_V<>dSal_O))+SUMPRODUCT(--(Payout_V<>Payout_O))+SUMPRODUCT(--(Ap_MesV<>Ap_MesO))+SUMPRODUCT(--(Ap_MontoV<>Ap_MontoO))"),
        fmt=NF_N, font=F_B)
    nm("c_Modificados", "Control", f"C{r}")
    r += 1
    put(ws, r, 2, "Deuda del presupuesto al final del horizonte")
    put(ws, r, 3, f"=INDEX(Calculo!{RG('P.DEUDA')},1,MIN({N},Fin))", fmt=NF_USD)
    nm("r_DeudaFinP", "Control", f"C{r}")
    ws.freeze_panes = "A5"


# ============================================================================ GUIA DE CARGA
def build_guia():
    ws = WS["Guia_Carga"]
    title(ws, "Guía de carga paso a paso",
          "Seguir en orden. El estado se calcula con los datos del libro: confirma que el dato está cargado y concilia, no que sea correcto.")
    widths(ws, {"A": 5, "B": 46, "C": 26, "D": 26, "E": 46, "F": 30, "G": 8, "H": 11, "I": 14, "J": 30})
    nm("g_Mes", "Guia_Carga", "C4")
    put(ws, 4, 2, "Mes en carga (próximo mes a cerrar)", font=F_B)
    put(ws, 4, 3, f"=MIN({N},Corte+1)", font=F_B, fmt=NF_N, fill=FL_V)
    put(ws, 4, 4, '=IF(ISNUMBER(n_Fecha1),"Fecha: "&TEXT(MONTH(EOMONTH(n_Fecha1,g_Mes-1)),"00")&"/"&YEAR(EOMONTH(n_Fecha1,g_Mes-1)),"Sin fecha del mes 1")', font=F_S)
    put(ws, 5, 2, "Pasos con alerta (⚠)", font=F_B)
    put(ws, 6, 2, "Leyenda: ✔ listo · ⚠ pendiente o a revisar · ○ informativo o manual · — no aplica", font=F_S)
    rr = lambda a, b: f"Reales!$F${RR[a]}:${MN}${RR[b]}"
    cnt = lambda a, b: f"COUNT(INDEX({rr(a, b)},0,g_Mes))"
    named = "SUMPRODUCT(--(Prod_Nombre<>\"\"))"
    def link(sheet, cell, label):
        return f'=HYPERLINK("#{sheet}!{cell}","{label}")'
    def rlink(code, label):
        return f'=HYPERLINK("#Reales!"&ADDRESS({RR[code]},{FCOL - 1}+g_Mes,4),"{label}, mes "&g_Mes)'
    p_last = SUPR["p_last"]
    steps = [
        ("A. Alta del negocio (una vez)",),
        ("Nombre, socios, actividad y etapa", link("Inicio", "C6", "Inicio!C6:C9"), "Plan de negocio / acta", "Etapa «Operación» si ya factura; «Implantación» si todavía no.",
         '=IF(AND(n_Negocio<>"",n_Etapa<>""),"✔ Listo","⚠ Pendiente")'),
        ("Fecha del mes 1 del modelo", link("Inicio", "C10", "Inicio!C10"), "Plan aprobado", "Preferentemente enero. Solo cambia las etiquetas de fecha.",
         '=IF(ISNUMBER(n_Fecha1),"✔ Listo","⚠ Pendiente (recomendado)")'),
        ("Mes real de inicio de operación", link("Inicio", "C12", "Inicio!C12"), "Primera factura / acta de inicio",
         "No cambiar el inicio del presupuesto: el real se carga aquí.",
         '=IF(n_Etapa="Operación",IF(ISNUMBER(n_InicioReal),"✔ Listo","⚠ Pendiente"),IF(n_Etapa="","⚠ Definir primero la etapa (paso 1)","— No aplica (implantación)"))'),
        ("B. Presupuesto original (una vez, plan aprobado; no se modifica después)",),
        ("Parámetros generales (columna «Presupuesto original»)", link("Supuestos", "D6", f"Supuestos!D6:D{p_last}"), "Plan aprobado",
         "Porcentajes como fracción (0,05 = 5%). Anotar la fuente en la nota.",
         f'=IF(COUNTBLANK(Supuestos!$D$6:$D${p_last})=0,"✔ Listo","⚠ Faltan "&COUNTBLANK(Supuestos!$D$6:$D${p_last})&" parámetros")'),
        ("Tipo de cambio por año", link("Supuestos", f"D{SUPR['TC_O']}", f"Supuestos!D{SUPR['TC_O']}:M{SUPR['TC_O']}"), "Plan / BCP",
         "Obligatorio: convierte la nómina y los costos en guaraníes.",
         f'=IF(COUNTIF(TC_O,">0")={NY},"✔ Listo","⚠ Faltan "&({NY}-COUNTIF(TC_O,">0"))&" años")'),
        ("Variaciones de precio, costos y salarios, y dividendos por año", link("Supuestos", f"D{SUPR['dPrecio_O']}", f"Supuestos!D{SUPR['dPrecio_O']}:M{SUPR['Payout_O']}"),
         "Plan aprobado", "Vacío no es cero: cargar 0 si no hay variación.",
         f'=IF(COUNTBLANK(Supuestos!$D${SUPR["dPrecio_O"]}:$M${SUPR["Payout_O"]})=0,"✔ Listo","⚠ Faltan "&COUNTBLANK(Supuestos!$D${SUPR["dPrecio_O"]}:$M${SUPR["Payout_O"]})&" celdas")'),
        ("Productos: nombre, unidad, precio, costos por unidad y capacidad", link("Productos", f"B{PR['first']}", f"Productos!B{PR['first']}:N{PR['last']}"),
         "Plan / costeo de producto", f"Hasta {NP} productos individuales. Capacidad 0 = no se controla.",
         f'=IF({named}=0,"⚠ Sin productos",IF(SUMPRODUCT((VolTot_O>0)*(Precio_O=0))>0,"⚠ "&SUMPRODUCT((VolTot_O>0)*(Precio_O=0))&" producto(s) con volumen y sin precio","✔ "&{named}&" producto(s)"))'),
        ("Volumen de ventas por año (presupuesto)", link("Productos", f"D{PR['vol_O']}", f"Productos!D{PR['vol_O']}:M{PR['vol_O'] + NP - 1}"), "Plan comercial",
         "Unidades vendidas por año de operación.",
         '=IF(SUM(VolTot_O)=0,"⚠ Sin volumen",IF(SUMPRODUCT((Prod_Nombre<>"")*(VolTot_O=0)*(VolTot_V=0))>0,"⚠ Hay productos sin volumen","✔ Listo"))'),
        ("Personal: puesto, clase, moneda, salario mensual y dotación por año", link("Costos", f"B{CO['per_first']}", f"Costos!B{CO['per_first']}:P{CO['per_last']}"),
         "Nómina / RR. HH.", "Clase: Directo, Indirecto o Administración. Salario sin cargas sociales.",
         '=IF(COUNT(Sal_O)=0,"⚠ Sin personal","✔ "&COUNT(Sal_O)&" puesto(s)")'),
        ("Costos fijos: concepto, clase, moneda, monto, crecimiento y año de inicio", link("Costos", f"B{CO['fc_first']}", f"Costos!B{CO['fc_first']}:J{CO['fc_first'] + NFIJ - 1}"),
         "Contratos / presupuesto de gastos", "IVA incluido solo si no se recupera.",
         '=IF(COUNT(FC_O)=0,"⚠ Sin costos fijos","✔ "&COUNT(FC_O)&" concepto(s)")'),
        ("Inversión (CAPEX): monto, mes de pago y vida útil por ítem", link("Inversion", f"B{IV['first']}", f"Inversion!B{IV['first']}:F{IV['last']}"),
         "Presupuesto de inversión / cotizaciones", "Aporte en especie: cargar el ítem y el aporte en el mismo mes.",
         '=IF(COUNT(Inv_D)=0,"⚠ Sin CAPEX",IF(SUMPRODUCT((Inv_D>0)*((Inv_E="")+(Inv_F="")))>0,"⚠ Ítems sin mes o sin vida útil","✔ "&COUNT(Inv_D)&" ítem(s)"))'),
        ("Financiamiento: préstamos, línea rotativa y aportes", link("Inversion", f"B{IV['loans']}", "Inversion, secciones B, B2 y C"),
         "Contratos bancarios / acta de socios", "Tasa nominal anual; gracia y cuotas según contrato.",
         '=IF(N(P_L1_Monto)+N(P_L2_Monto)+N(P_L3_Monto)+N(P_LC_Lim)+SUM(Ap_MontoO)=0,"○ Sin financiamiento cargado","✔ Listo")'),
        ("Revisión del presupuesto: balance y plan publicado", link("Resultados", "B4", "Resultados"), "—",
         "Balance = 0. Explicar las diferencias con el plan publicado (sección E).",
         '=IF(c_C01="OK","✔ Balance OK","⚠ Revisar el balance")'),
        ("C. Proyección vigente (cuando cambia un supuesto)",),
        ("Sobrescribir en «Vigente» solo lo que cambió, con evidencia", link("Supuestos", "E6", "Columnas «Vigente»"),
         "Contrato, cotización, nómina, oferta bancaria", "Anotar fuente y fecha. Para volver al presupuesto, reponer la fórmula que hereda.",
         '="○ "&c_Modificados&" celda(s) vigente(s) modificada(s)"'),
        ("Escenario de trabajo", link("Inicio", "C13", "Inicio!C13"), "—", "Vigente para el trabajo normal; Conservador y Optimista solo para análisis.",
         '=IF(n_Escenario="Vigente","✔ Vigente","⚠ Activo: "&n_Escenario&" (volver a Vigente)")'),
        ("D. Cierre mensual (cada mes): los vínculos llevan a la columna del mes en carga",),
        ("TC real del mes", rlink("TC", "Reales, fila " + str(RR["TC"])), "BCP / política de CIE", "Siempre el mismo criterio (promedio o cierre).",
         f'=IF(ISNUMBER(INDEX(Reales!$F${RR["TC"]}:${MN}${RR["TC"]},1,g_Mes)),"✔ Cargado","⚠ Pendiente")', "tc"),
        ("Unidades vendidas por producto", rlink("U1", f"Reales, filas {RR['U1']}-{RR['U' + str(NP)]}"), "SAP SD (facturación)", "Una fila por producto con nombre.",
         f'=IF({named}=0,"— Sin productos",IF(SUMPRODUCT((Prod_Nombre<>"")*ISNUMBER(INDEX({rr("U1", "U" + str(NP))},0,g_Mes)))={named},"✔ Cargado","⚠ Faltan "&({named}-SUMPRODUCT((Prod_Nombre<>"")*ISNUMBER(INDEX({rr("U1", "U" + str(NP))},0,g_Mes))))&" producto(s)"))', "u"),
        ("Ventas netas por producto (moneda en la columna C)", rlink("S1", f"Reales, filas {RR['S1']}-{RR['S' + str(NP)]}"), "SAP SD / FI", "Facturación neta, sin IVA.",
         f'=IF({named}=0,"— Sin productos",IF(SUMPRODUCT((Prod_Nombre<>"")*ISNUMBER(INDEX({rr("S1", "S" + str(NP))},0,g_Mes)))={named},"✔ Cargado","⚠ Faltan "&({named}-SUMPRODUCT((Prod_Nombre<>"")*ISNUMBER(INDEX({rr("S1", "S" + str(NP))},0,g_Mes))))&" producto(s)"))', "s"),
        ("Costos devengados: 7 rubros sin IVA recuperable", rlink("MAT", f"Reales, filas {RR['MAT']}-{RR['TRIB']}"), "SAP MM / CO, nómina",
         "En un mes cerrado, vacío vale cero: cargar 0 si no hubo.",
         f'=IF({cnt("MAT", "TRIB")}=7,"✔ 7 de 7","⚠ "&{cnt("MAT", "TRIB")}&" de 7 cargados")', "c"),
        ("Movimientos de caja: CAPEX, desembolsos, intereses, amortizaciones, aportes, dividendos", rlink("CAPEX", f"Reales, filas {RR['CAPEX']}-{RR['DIV']}"),
         "Tesorería / SAP PS", "Cargar 0 si no hubo movimiento.",
         f'=IF({cnt("CAPEX", "DIV")}=6,"✔ 6 de 6","⚠ "&{cnt("CAPEX", "DIV")}&" de 6 cargados")', "m"),
        ("Saldos al cierre: caja, clientes, inventarios, IVA, proveedores, deuda", rlink("S_CAJA", f"Reales, filas {RR['S_CAJA']}-{RR['S_DEUDA']}"),
         "Balance de comprobación / bancos", "Obligatorios en el último mes cerrado.",
         f'=IF({cnt("S_CAJA", "S_DEUDA")}=6,"✔ 6 de 6","⚠ "&{cnt("S_CAJA", "S_DEUDA")}&" de 6 cargados")', "b"),
        ("CAPEX por ítem: ejecutado, pagado, comprometido, por contratar", link("Inversion", f"G{IV['first']}", f"Inversion!G{IV['first']}:L{IV['last']}"),
         "SAP PS (elementos PEP)", "El pagado total debe igualar el CAPEX de Reales (control C08).",
         f'=IF(AND(N(i_Pagado)=0,SUM(Reales!$F${RR["CAPEX"]}:${MN}${RR["CAPEX"]})=0),"○ Sin CAPEX ejecutado todavía",IF(c_C08="OK","✔ Concilia","⚠ No concilia con Reales"))'),
        ("Actualizar el último mes cerrado", link("Inicio", "C11", "Inicio!C11"), "—", "Recién cuando los seis pasos anteriores de datos estén en ✔.",
         None, "cierre"),
        ("Revisar Control después de cerrar", link("Control", "B4", "Control"), "—", "C03 a C08 en OK; explicar el ajuste de conciliación de caja si es grande.",
         '=IF(Corte=0,"○ Sin meses cerrados",IF(AND(c_C03="OK",c_C04="OK",c_C05="OK",c_C06="OK",c_C07="OK",c_C08="OK"),"✔ Controles de carga OK","⚠ "&c_Alertas&" control(es) en REVISAR"))'),
        ("E. Revisión y reporte",),
        ("Brecha de caja y financiamiento", link("Resultados", "B4", "Resultados, indicadores"), "—",
         "Si hay brecha: definir aporte, línea rotativa o mejora del capital de trabajo.",
         '=IF(c_C13="OK","✔ Sin brecha","⚠ Brecha máxima USD "&FIXED(r_BrechaV,0))'),
        ("Desvíos vigente vs presupuesto (ventas, EBITDA, caja)", link("Resultados", f"B{RS['D_vtas']}", "Resultados, sección D"), "—",
         "Explicar por volumen, precio, costo y atraso.", '="○ Revisar"'),
        ("Guardar una versión fechada del cierre y registrar los cambios", link("Inicio", "B4", "Inicio"), "—",
         "Por ejemplo: NEGOCIO_cierre_2026-10.xlsx.", '="○ Manual"'),
        ("Trimestral: revisar supuestos vigentes y escenarios", link("Supuestos", "E6", "Supuestos"), "—",
         "Actualizar tasas, TC, precios y costos con evidencia.", '="○ Manual"'),
        ("F. Proyección rodante de 24 meses (Detalle_24m) y revisión del plan",),
    ] + [(lab, f'=HYPERLINK("#Detalle_24m!"&ADDRESS({DT[code + "_first"]},MIN(149,5+Corte+1),4),"Detalle_24m, {ref}")', fuente, crit,
          (f'=IF(N(n_Ventana)=0,"— Sin ventana",IF(SUMPRODUCT((Detalle_24m!${M1}${DT["hdr"]}:${MN}${DT["hdr"]}>Corte)*(Detalle_24m!${M1}${DT["hdr"]}:${MN}${DT["hdr"]}<=Corte+n_Ventana)*(Det_n{code}>0))=0,'
           f'"○ Usa supuestos anuales",IF(SUMPRODUCT((Detalle_24m!${M1}${DT["hdr"]}:${MN}${DT["hdr"]}>Corte)*(Detalle_24m!${M1}${DT["hdr"]}:${MN}${DT["hdr"]}<=Corte+n_Ventana)*(Det_n{code}>0))'
           f'>=MIN(n_Ventana,{N}-Corte),"✔ "&MIN(n_Ventana,{N}-Corte)&" meses detallados","⚠ "&SUMPRODUCT((Detalle_24m!${M1}${DT["hdr"]}:${MN}${DT["hdr"]}>Corte)*(Detalle_24m!${M1}${DT["hdr"]}:${MN}${DT["hdr"]}<=Corte+n_Ventana)*(Det_n{code}>0))'
           f'&" de "&MIN(n_Ventana,{N}-Corte)&" meses: completar hasta el mes "&MIN({N},Corte+n_Ventana))))'))
         for code, ref, lab, fuente, crit in [
             ("Q", "bloque B", "Volúmenes por producto en la ventana", "Pedidos, plan comercial", "Con cada cierre, agregar el nuevo último mes de la ventana."),
             ("P", "bloque C", "Precios de venta por producto", "Contratos / listas de precios", "Solo si cambian respecto del supuesto anual."),
             ("MP", "bloque D", "Materia prima por unidad (asegurada)", "Cotizaciones y contratos de compra", "Costo asegurado de los próximos meses."),
             ("DO", "bloque G", "Dotación por puesto", "Plan de personal / RR. HH.", "Ingresos y egresos previstos por mes."),
             ("SA", "bloque H", "Salarios por puesto", "Nómina / ajustes salariales", "Incluye el mes de cada ajuste salarial."),
             ("FC", "bloque I", "Costos fijos del mes", "Contratos (alquiler, energía, mantenimiento)", "Montos en la moneda de cada concepto."),
             ("TC", "bloque A", "Tipo de cambio mensual", "BCP / presupuesto de TC", "Opcional: si no, se usa el TC anual.")]] + [
        ("Revisar la proyección de 24 meses", link("Proyeccion_24m", "B4", "Proyeccion_24m"), "—",
         "Caja, deuda y desvíos frente al presupuesto mes a mes.", '=IF(N(py_Brecha)>0.5,"⚠ Caja bajo el mínimo en los próximos 24 meses","✔ Sin brecha en 24 meses")'),
        ("Revisión del plan a 10 años (presupuesto original)", link("Inicio", "C15", "Inicio!C15:C17"), "Gerencia / Directorio",
         "Anual o bienal: guardar el archivo anterior, pasar los supuestos vigentes aprobados al presupuesto y actualizar versión y fecha.",
         '=IF(NOT(ISNUMBER(n_FechaRev)),"⚠ Cargar fecha de la última revisión",IF(TODAY()>EDATE(n_FechaRev,IF(n_Frecuencia="Bienal",24,12)),"⚠ Revisión vencida","✔ Próxima: "&DAY(EDATE(n_FechaRev,IF(n_Frecuencia="Bienal",24,12)))&"/"&MONTH(EDATE(n_FechaRev,IF(n_Frecuencia="Bienal",24,12)))&"/"&YEAR(EDATE(n_FechaRev,IF(n_Frecuencia="Bienal",24,12)))))'),
    ]
    r = 8
    header(ws, r, ["N°", "Paso", "Dónde (clic para ir)", "Fuente del dato", "Criterio", "Estado (automático)", "Hecho", "Fecha", "Responsable", "Notas"], 1, 30)
    r += 1
    first, n, filas = r, 0, {}
    for st in steps:
        if len(st) == 1:
            section(ws, r, st[0], 10)
            r += 1
            continue
        n += 1
        filas[st[5] if len(st) > 5 else n] = r
        r += 1
    r = first
    n = 0
    for st in steps:
        if len(st) == 1:
            r += 1
            continue
        n += 1
        paso, donde, fuente, crit, est = st[:5]
        if est is None:
            ok = ",".join(f'LEFT(F{filas[k]},1)="✔"' for k in ("tc", "u", "s", "c", "m", "b"))
            est = f'=IF(AND({ok}),"✔ Listo para cerrar: poner "&g_Mes,"⚠ Esperar: faltan datos del mes "&g_Mes)'
        put(ws, r, 1, n, al=CEN)
        put(ws, r, 2, paso, al=WRAP)
        put(ws, r, 3, donde, font=Font(name=FONT, size=9, color="0563C1", underline="single"))
        put(ws, r, 4, fuente, font=F_S, al=WRAP)
        put(ws, r, 5, crit, font=F_S, al=WRAP)
        put(ws, r, 6, est, font=F_B)
        inp(ws, r, 7, None)
        inp(ws, r, 8, None, NF_DATE)
        inp(ws, r, 9, None)
        inp(ws, r, 10, None)
        ws.row_dimensions[r].height = 26
        r += 1
    last = r - 1
    dv(ws, f"G{first}:G{last}", ["Sí", "No"])
    put(ws, 5, 3, f'=COUNTIF(F{first}:F{last},"⚠*")', font=F_B, fmt=NF_N)
    nm("g_Pendientes", "Guia_Carga", "C5")
    rng = f"F{first}:F{last}"
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(F{first},1)="✔"'], fill=FL_OK))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT(F{first},1)="⚠"'], fill=FL_BAD))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'OR(LEFT(F{first},1)="○",LEFT(F{first},1)="—")'], fill=FL_INH))
    r += 1
    notas = ["El estado automático confirma que el dato existe y concilia; la revisión de que sea correcto es del responsable.",
             "Cierre mensual: cargar todo el mes en carga, verificar que los pasos de datos estén en ✔ y recién entonces actualizar Inicio!C11.",
             "Al actualizar el último mes cerrado, el mes en carga pasa al siguiente y los pasos del cierre vuelven a «Pendiente»: es el ciclo normal."]
    for t in notas:
        put(ws, r, 2, t, font=F_S)
        r += 1
    ws.freeze_panes = f"C{first}"


# ============================================================================ CUADROS (estructura del plan de negocio)
# Una hoja por cuadro del plan, con el mismo orden y nombre de líneas, para identificar y comparar cada cifra.
# Cada hoja tiene tres bloques con las mismas filas: A = modelo (versión elegida en «Cuadros»), B = plan publicado
# (celdas de entrada; Ecostar precargado) y C = diferencia modelo − plan con su explicación.
# Columnas: D = año 0 (implantación, antes del inicio de operación), E a N = años de operación 1 a 10, O = total.
CU_HOJAS = [  # hoja, cuadros del plan, contenido, páginas del plan
    ("C01 Inversión", "1, 2, 5 y 6", "Inversión y financiamiento; detalle de máquinas, infraestructura y otras inversiones", "11-13"),
    ("C07 Producción", "7 y 7a", "Programa de producción: capacidad instalada, producción, venta y utilización", "14-15"),
    ("C08 Ingresos", "8", "Ingresos anuales proyectados: venta, cantidad y precio por producto", "16"),
    ("C09 Materiales", "9a y 9b", "Materia prima, insumos y consumibles por producto", "17-18"),
    ("C10a Costos", "10a", "Costos anuales proyectados (producción, administración, depreciación, financieros) y costos fijos por concepto", "19"),
    ("C10b Costo unitario", "10b", "Costos unitarios por producto y margen bruto", "20"),
    ("C11 Rubros", "11 y 12", "Costos por rubros y estructura de costos (%)", "21"),
    ("C13 Deuda e IVA", "13 y 14", "Servicio de la deuda y cálculo del IVA", "22"),
    ("C15 Resultados", "15", "Resultados operacionales proyectados y distribución de utilidades", "23"),
    ("C16 Capital trabajo", "16", "Capital de trabajo", "24"),
    ("C17 Equilibrio", "17", "Punto de equilibrio", "25"),
    ("C18 Fondos", "18", "Origen y aplicación de fondos", "26"),
    ("C19 VAN-TIR", "19", "VAN, TIR, payback y ponderación de la tasa de corte", "27-28"),
    ("C20 Depreciación", "20", "Depreciación del activo fijo y valor residual", "29"),
    ("RRHH", "—", "Recursos humanos: dotación y costo por puesto", "30-34"),
]
CUY = 10                       # años de operación en columnas
YC = lambda y: CL(4 + y)       # D = año 0, E..N = años 1..10
TC_ = CL(5 + CUY)              # O = total
NC_ = CL(6 + CUY)              # P = comentario
CUR_ROWS = {}                  # hoja -> {"A": {clave: fila}, "B": {...}, "C": {...}}
PLAN = D.get("plan_cuadros", {})
NOTA = D.get("plan_notas_lineas", {})
VIS_P = len(D.get("productos", [])) or 5        # productos visibles (el resto queda agrupado)
VIS_PER = len(D.get("personal", [])) or 10
VIS_FC = len(D.get("fijos", [])) or 10
VIS_CAP = len(D.get("capex", [])) or 10
VIS_L = len(D.get("prestamos", [])) or 1


def CR(X, k):
    return f"Calculo!{RG(X + '.' + k)}"


def CSEL(X, y):
    return f"({CR(X, 'pre')}=1)" if y == 0 else f"({CR(X, 'aop')}={y})"


def CS(X, k, y):
    """Suma del año de operación y (0 = implantación) de una fila del motor."""
    sel = f"{CR(X, 'pre')},1" if y == 0 else f"{CR(X, 'aop')},{y}"
    return f"SUMIFS({CR(X, k)},{sel})"


def CE(X, k, y):
    """Saldo al cierre del año de operación y (último mes del año)."""
    if y == 0:
        return f"IF(SUM({CR(X, 'pre')})=0,0,INDEX({CR(X, k)},1,SUM({CR(X, 'pre')})))"
    return f"IF(COUNTIFS({CR(X, 'aop')},{y})=0,0,INDEX({CR(X, k)},1,_xlfn.MAXIFS(Calculo!{RG('m')},{CR(X, 'aop')},{y})))"


def CN(X, y):
    """Meses del año de operación y."""
    return f"SUM({CR(X, 'pre')})" if y == 0 else f"COUNTIFS({CR(X, 'aop')},{y})"


def CSUM(X, y, keys):
    return "(" + "+".join(("-" if k.startswith("-") else "") + CS(X, k.lstrip("-"), y) for k in keys).replace("+-", "-") + ")"


def OV_(X):
    return "O" if X == "P" else "V"


def TCY(X, y):
    return f"INDEX(TC_{OV_(X)},1,{min(NY, max(1, y))})*(1+{X}_dTC)"


def UNITS(X, y):
    return f"SUMPRODUCT({CSEL(X, y)}*Calculo!${M1}${ROW[X + '.u1']}:${MN}${ROW[X + '.u' + str(NP)]})"


def ALIM(X, k, y):
    return f"SUMPRODUCT({CSEL(X, y)}*{CR(X, k + '_n')}*{CR(X, 'itc')})*{X}_Alim"


def BASE(X, k, y):
    """Salarios sin cargas ni alimentación de una clase de personal (dir, ind, adm)."""
    return f"(({CS(X, k + '_f', y)})-{ALIM(X, k, y)})/(1+{X}_Cargas)"


def MATP(X, y, i, nmx, idx):
    """Costo anual por producto con los supuestos anuales: unidades × costo unitario × índice."""
    return (f"{CS(X, 'u' + str(i), y)}*N(INDEX({nmx}_{OV_(X)},{i}))*INDEX(Idx{idx}_{OV_(X)},1,{min(NY, max(1, y))})"
            f"*(1+{X}_dMat)/(1-{X}_Merma)")


TOTCOST = ["MAT", "OVAR", "PERP", "FIJP", "TRIB", "PERA", "ADM", "-IVAR", "DEP", "INT"]


def pname(i, suf=""):
    return f'=IF(INDEX(Prod_Nombre,{i})="","Producto {i}",INDEX(Prod_Nombre,{i}))' + (f'&"{suf}"' if suf else "")


def pc(c):
    """Columna anterior."""
    return CL(column_index_from_string(c) - 1)


def L(k, t, m, f=NF_USD, b=False, u="USD", tot="sum", y0=True, pub=True, grp=False, nota=None, inB=False):
    """Línea de un cuadro. m: ("s", fila del motor: suma del año), ("e", fila: saldo al cierre), ("x", fn(X, y)),
    ("c", fn(col, R): fórmula con filas del mismo bloque) o ("t", constante). pub = se compara con el plan;
    inB = en el bloque del plan es dato de entrada aunque en el modelo sea una fórmula del bloque."""
    return dict(k=k, t=t, m=m, f=f, b=b, u=u, tot=tot, y0=y0, pub=pub, grp=grp, nota=nota, inB=inB)


TOL = {NF_USD: 2, NF_USD2: 0.005, NF_PCT: 0.0005, NF_PCT2: 0.00005, NF_USD4: 0.00005}   # redondeo del plan publicado


def S(t):
    return dict(sec=t)


def cu_head(ws, hoja, cuadro, sub):
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 7, "B": 54, "C": 11, TC_: 13, NC_: 70})
    for y in range(0, CUY + 1):
        ws.column_dimensions[YC(y)].width = 11.5
    put(ws, 1, 1, cuadro, font=F_T)
    put(ws, 2, 1, '="Información en dólares   |   Versión: "&cu_Ver&"   |   Escenario: "&n_Escenario&IF(Corte>0,"   |   Real hasta el mes "&Corte,"")',
        font=F_AL)
    put(ws, 3, 1, f'="Plan de referencia: "&IF(n_Version="","(sin versión)",n_Version)&" — {sub}"', font=F_S)
    put(ws, 1, 15, '=HYPERLINK("#Cuadros!A1","← Índice de cuadros")', font=Font(name=FONT, size=9, color="0563C1", underline="single"))
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.sheet_properties.outlinePr.summaryBelow = False


def cu_years(ws, r, lab, blk):
    header(ws, r, ["", lab, "Unidad"] + ["Año 0"] + [f"Año {y}" for y in range(1, CUY + 1)] + ["Total"] + (["Explicación de la diferencia"] if blk == "C" else []), 1, 24)
    r += 1
    put(ws, r, 2, "Estado del año" if blk == "A" else "Año calendario (modelo)", font=F_S)
    for y in range(0, CUY + 1):
        if blk == "A":
            f = ('=IF(cu_P=1,"Presupuesto",IF(V_Inicio=1,"—",IF(Corte>=V_Inicio-1,"Real",IF(Corte>0,"Real+Proy.","Proyección"))))' if y == 0 else
                 f'=IF(cu_P=1,"Presupuesto",IF(Corte>=V_Inicio+{12 * y - 1},"Real",IF(Corte>=V_Inicio+{12 * (y - 1)},"Real+Proy.","Proyección")))')
        else:
            f = (f'=IF(OR(n_Fecha1="",IF(cu_P=1,P_Inicio,V_Inicio)=1),"",YEAR(n_Fecha1))' if y == 0 else
                 f'=IF(n_Fecha1="","",YEAR(EOMONTH(n_Fecha1,IF(cu_P=1,P_Inicio,V_Inicio)-1+{12 * (y - 1)})))')
        put(ws, r, 4 + y, f, font=F_S, al=CEN, fmt=NF_N)
    return r + 1


def cu_block(ws, r, lines, blk, R):
    """Escribe un bloque (A modelo, B plan, C diferencia) con las filas ya asignadas en R[blk]."""
    RB = R[blk]
    for ln in lines:
        if "sec" in ln:
            put(ws, r, 2, ln["sec"], font=F_B, fill=FL_SEC)
            for c in range(1, 6 + CUY + (1 if blk == "C" else 0)):
                ws.cell(row=r, column=c).fill = FL_SEC
            r += 1
            continue
        k, m = ln["k"], ln["m"]
        assert RB[k] == r, (k, RB[k], r)
        bold = ln["b"]
        put(ws, r, 1, k, font=F_S)
        put(ws, r, 2, ln["t"], font=F_B if bold else (F_LINK if str(ln["t"]).startswith("=") else F), fill=FL_TOT if bold else None)
        put(ws, r, 3, ln["u"], font=F_S)
        for y in range(0, CUY + 1):
            c = YC(y)
            if y == 0 and not ln["y0"]:
                continue
            fill = FL_TOT if bold else None
            if blk == "C":
                a, p = f"{c}{R['A'][k]}", f"{c}{R['B'][k]}"
                v = f'=IF(OR(NOT(ISNUMBER({a})),NOT(ISNUMBER({p}))),"",{a}-{p})' if ln["pub"] else None
                put(ws, r, 4 + y, v, fmt=ln["f"], font=F_B if bold else F, fill=fill)
                continue
            if blk == "B" and (m[0] != "c" or ln["inB"]):
                if not ln["pub"]:
                    continue
                vals = PLAN.get(k)
                val = vals[y] if vals and y < len(vals) else None
                inp(ws, r, 4 + y, val, ln["f"])
                continue
            elif m[0] == "c":
                v = f"=IFERROR({m[1](c, RB)},\"\")"
            elif m[0] == "s":
                v = f"=IF(cu_P=1,{CS('P', m[1], y)},{CS('V', m[1], y)})"
            elif m[0] == "e":
                v = f"=IF(cu_P=1,{CE('P', m[1], y)},{CE('V', m[1], y)})"
            elif m[0] == "x":
                v = f"=IF(cu_P=1,{m[1]('P', y)},{m[1]('V', y)})"
            else:
                v = m[1]
            put(ws, r, 4 + y, v, fmt=ln["f"], font=F_B if bold else F, fill=fill)
        # total
        t = ln["tot"]
        tf = None
        if blk == "C":
            if ln["pub"] and t:
                a, p = f"{TC_}{R['A'][k]}", f"{TC_}{R['B'][k]}"
                tf = f'=IF(OR(NOT(ISNUMBER({a})),NOT(ISNUMBER({p}))),"",{a}-{p})'
            if ln["pub"]:
                rng, tol = f"D{r}:{YC(CUY)}{r}", TOL.get(ln["f"], 1)
                txt = (NOTA.get(k) or ln["nota"] or "Diferencia a explicar").replace('"', "'")
                put(ws, r, 6 + CUY, f'=IF(COUNT({rng})=0,"",IF(AND(MAX({rng})<{tol},MIN({rng})>-{tol}),"Sin diferencia","{txt}"))', font=F_S)
        elif t == "sum":
            tf = f"=SUM(D{r}:{YC(CUY)}{r})" if blk == "A" else f'=IF(COUNT(D{r}:{YC(CUY)}{r})=0,"",SUM(D{r}:{YC(CUY)}{r}))'
        elif t == "c" and m[0] == "c":
            tf = f"=IFERROR({m[1](TC_, RB)},\"\")"
        elif t == "last":
            tf = f"={YC(CUY)}{r}"
        if tf:
            put(ws, r, 5 + CUY, tf, fmt=ln["f"], font=F_B, fill=FL_TOT)
        if ln["grp"]:
            ws.row_dimensions[r].outlineLevel = 1
            ws.row_dimensions[r].hidden = True
            ws.sheet_format.outlineLevelRow = 1
        r += 1
    return r


def cu_sheet(hoja, cuadro, sub, lines, notas=None):
    ws = WS[hoja]
    cu_head(ws, hoja, cuadro, sub)
    r = 5
    R = {"A": {}, "B": {}, "C": {}}
    labels = {"A": "A. MODELO — versión seleccionada en la hoja «Cuadros» (presupuesto original o vigente)",
              "B": "B. PLAN PUBLICADO — cifras del documento (celdas de entrada; vacío = el plan no informa la línea)",
              "C": "C. DIFERENCIA: MODELO − PLAN PUBLICADO"}
    fills = {"A": FL_V, "B": FL_P, "C": FL_TOT}
    for blk in ("A", "B", "C"):
        for c in range(1, 6 + CUY + (1 if blk == "C" else 0)):
            ws.cell(row=r, column=c).fill = fills[blk]
        put(ws, r, 1, labels[blk], font=F_SEC, fill=fills[blk])
        r += 1
        if blk == "A":
            ws.freeze_panes = f"D{r + 1}"
        r = cu_years(ws, r, "Concepto (USD)", blk)
        rr = r
        for ln in lines:
            if "sec" in ln:
                rr += 1
                continue
            R[blk][ln["k"]] = rr
            rr += 1
        r = cu_block(ws, r, lines, blk, R)
        r += 1
    if notas:
        section(ws, r, "Notas", 5 + CUY)
        r += 1
        for t in notas:
            put(ws, r, 2, t, font=F_S)
            r += 1
    CUR_ROWS[hoja] = R
    return ws, R


def cR(k):
    """Referencia a la fila de la clave k en el bloque actual (para fórmulas «c»)."""
    return lambda c, R: f"{c}{R[k]}"


def cSum(keys):
    """Suma de filas del bloque; vacío si ninguna tiene dato (plan sin cifras)."""
    return lambda c, R: f'IF(COUNT({",".join(f"{c}{R[k]}" for k in keys)})=0,"",' + "+".join(f"N({c}{R[k]})" for k in keys) + ")"


def cDiv(a, b):
    return lambda c, R: f"IF(OR(NOT(ISNUMBER({c}{R[a]})),N({c}{R[b]})=0),\"\",{c}{R[a]}/{c}{R[b]})"


def cExpr(tpl):
    """Plantilla con claves entre llaves: «{c09.mp}-{c09.mp.1}» → referencias del bloque; vacío si no hay datos."""
    import re
    keys = re.findall(r"\{([^}]+)\}", tpl)
    return lambda c, R: (f'IF(COUNT({",".join(f"{c}{R[k]}" for k in keys)})=0,"",'
                         + re.sub(r"\{([^}]+)\}", lambda mm: f"N({c}{R[mm.group(1)]})", tpl) + ")")


# ---------------------------------------------------------------------------- C07 producción
def build_c07():
    lines = []
    for i in range(1, NP + 1):
        g = i > VIS_P
        lines += [
            L(f"c07.cap.{i}", pname(i, " — capacidad instalada"), ("x", lambda X, y, i=i: (
                f'IF(INDEX(Prod_Nombre,{i})="",0,IF(AND(N(INDEX(Cap_Anio,{i}))>0,{y}>=N(INDEX(Cap_Anio,{i}))),N(INDEX(Cap_Amp,{i})),N(INDEX(Cap_Ini,{i}))))')),
              NF_INT, u="u", tot=None, y0=False, grp=g),
            L(f"c07.p.{i}", "   producción (venta ÷ (1 − merma))", ("x", lambda X, y, i=i: f"{CS(X, 'u' + str(i), y)}/(1-{X}_Merma)"), NF_INT, u="u", grp=g),
            L(f"c07.u.{i}", "   venta", ("s", f"u{i}"), NF_INT, u="u", grp=g),
            L(f"c07.ut.{i}", "   producción / capacidad instalada", ("c", cDiv(f"c07.p.{i}", f"c07.cap.{i}")), NF_PCT, u="%", tot=None, y0=False, grp=g),
            L(f"c07.vp.{i}", "   venta / producción", ("c", cDiv(f"c07.u.{i}", f"c07.p.{i}")), NF_PCT, u="%", tot="c", y0=False, grp=g),
        ]
    lines += [S("Totales"),
              L("c07.ptot", "TOTAL PRODUCCIÓN DE UNIDADES", ("c", cSum([f"c07.p.{i}" for i in range(1, NP + 1)])), NF_INT, b=True, u="u"),
              L("c07.vtot", "TOTAL UNIDADES PRODUCIDAS Y VENDIDAS", ("c", cSum([f"c07.u.{i}" for i in range(1, NP + 1)])), NF_INT, b=True, u="u")]
    cu_sheet("C07 Producción", "Cuadros 7 y 7a — Programa de producción proyectada", "cuadros 7 y 7a, pág. 14-15", lines,
             [f"Productos {VIS_P + 1} a {NP} agrupados (botón + a la izquierda para verlos).",
              "Capacidad: hoja Productos (inicial y ampliada desde el año indicado). Producción = venta ÷ (1 − merma)."])


# ---------------------------------------------------------------------------- C08 ingresos
def build_c08():
    lines = [L("c08.vtas", "INGRESOS TOTALES — ingresos netos (USD)", ("s", "vtas"), b=True),
             L("c08.iva", "IVA sobre ventas (régimen de maquila: exportación sin IVA)", ("t", 0))]
    for i in range(1, NP + 1):
        g = i > VIS_P
        lines += [L(f"c08.s.{i}", pname(i, " — venta"), ("s", f"s{i}"), grp=g),
                  L(f"c08.q.{i}", "   cantidad", ("s", f"u{i}"), NF_INT, u="u", grp=g),
                  L(f"c08.p.{i}", "   precio promedio", ("c", cDiv(f"c08.s.{i}", f"c08.q.{i}")), NF_USD2, u="USD/u", tot="c", grp=g)]
    cu_sheet("C08 Ingresos", "Cuadro 8 — Ingresos anuales proyectados", "cuadro 8, pág. 16", lines,
             [f"Productos {VIS_P + 1} a {NP} agrupados. Precio promedio = venta ÷ cantidad (incluye ajustes de escenario y precios reales).",
              "Régimen de maquila: ventas de exportación sin IVA."])


# ---------------------------------------------------------------------------- C09 materiales
def build_c09():
    lines = []
    for code, nmx, idx, tit, tot_k in [("mp", "MP", "MP", "MATERIA PRIMA", "mp_f"), ("ins", "INS", "OV", "INSUMOS", "ins_f"),
                                       ("ov", "OV", "OV", "INSUMOS CONSUMIBLES Y FLETES DE MATERIA PRIMA (otros variables por unidad)", "ov_f")]:
        lines.append(S(f"Cuadro 9{'a' if code == 'mp' else 'b'} — {tit.lower()}"))
        lines.append(L(f"c09.{code}", f"TOTAL {tit}", ("s", tot_k), b=True))
        for i in range(1, NP + 1):
            g = i > VIS_P
            lines += [L(f"c09.{code}.{i}", pname(i), ("x", lambda X, y, i=i, nmx=nmx, idx=idx: MATP(X, y, i, nmx, idx)), grp=g),
                      L(f"c09.{code}u.{i}", "   costo por pieza", ("c", cDiv(f"c09.{code}.{i}", f"c09.u.{i}")), NF_USD2, u="USD/u", tot="c", y0=False, grp=g)]
            if code == "mp":
                lines.append(L(f"c09.u.{i}", "   piezas", ("s", f"u{i}"), NF_INT, u="u", grp=g))
        lines.append(L(f"c09.{code}aj", "Ajuste por detalle mensual o datos reales (total del motor − suma por producto)",
                       ("c", cExpr("{c09." + code + "}-(" + "+".join("{c09." + code + "." + str(i) + "}" for i in range(1, NP + 1)) + ")")), pub=False))
    cu_sheet("C09 Materiales", "Cuadros 9a y 9b — Materia prima e insumos", "cuadros 9a y 9b, pág. 17-18", lines,
             ["Costo por producto = unidades × costo unitario del año (índice de la hoja Supuestos). El ajuste muestra lo que aporta Detalle_24m o los meses reales.",
              "En el modelo, «otros variables por unidad» reúne los insumos consumibles y el flete de la materia prima (el plan los separa en los cuadros 9b y 10a)."])


# ---------------------------------------------------------------------------- C10a costos anuales
def build_c10a():
    PROD = ["MAT", "OVAR", "PERP", "FIJP", "TRIB"]
    lines = [
        S("1. Costos de producción"),
        L("c10.prod", "COSTOS DE PRODUCCIÓN", ("x", lambda X, y: CSUM(X, y, PROD)), b=True),
        L("c10.mp", "Materia prima", ("s", "mp_f")),
        L("c10.ins", "Insumos", ("s", "ins_f")),
        L("c10.ov", "Insumos consumibles y fletes de materia prima", ("s", "ov_f")),
        L("c10.mod", "Mano de obra directa (salarios)", ("x", lambda X, y: BASE(X, "dir", y))),
        L("c10.moi", "Mano de obra indirecta de producción (salarios)", ("x", lambda X, y: BASE(X, "ind", y))),
        L("c10.carp", "Prestaciones sociales del personal de producción", ("x", lambda X, y: f"({BASE(X, 'dir', y)}+{BASE(X, 'ind', y)})*{X}_Cargas")),
        L("c10.alip", "Alimentación del personal de producción", ("x", lambda X, y: f"{ALIM(X, 'dir', y)}+{ALIM(X, 'ind', y)}")),
        L("c10.fp", "Costos fijos de producción (energía, mantenimiento y otros)", ("s", "fp_f")),
        L("c10.desp", "Despachos de importación", ("s", "desp_f")),
        L("c10.vs", "Seguro de caución, comisiones y logística de venta", ("s", "vs_f")),
        L("c10.pkg", "Packaging", ("s", "pkg_f")),
        L("c10.impp", "Imprevistos de producción", ("s", "impp_f")),
        L("c10.trib", "Tasas y tributos — maquila", ("s", "trib_f")),
        L("c10.ajp", "Ajuste por datos reales (total del motor − suma de líneas)",
          ("c", cExpr("{c10.prod}-({c10.mp}+{c10.ins}+{c10.ov}+{c10.mod}+{c10.moi}+{c10.carp}+{c10.alip}+{c10.fp}+{c10.desp}+{c10.vs}+{c10.pkg}+{c10.impp}+{c10.trib})")),
          pub=False),
        S("2. Gastos de administración"),
        L("c10.adm", "GASTOS DE ADMINISTRACIÓN", ("x", lambda X, y: CSUM(X, y, ["PERA", "ADM"])), b=True),
        L("c10.moa", "Mano de obra administrativa (salarios)", ("x", lambda X, y: BASE(X, "adm", y))),
        L("c10.cara", "Prestaciones sociales", ("x", lambda X, y: f"{BASE(X, 'adm', y)}*{X}_Cargas")),
        L("c10.alia", "Alimentación del personal", ("x", lambda X, y: ALIM(X, "adm", y))),
        L("c10.fa", "Costos fijos de administración (detalle por concepto más abajo)", ("s", "fa_f")),
        L("c10.impa", "Imprevistos de administración", ("s", "impa_f")),
        L("c10.aja", "Ajuste por datos reales (total del motor − suma de líneas)",
          ("c", cExpr("{c10.adm}-({c10.moa}+{c10.cara}+{c10.alia}+{c10.fa}+{c10.impa})")), pub=False),
        S("3. IVA, gastos no desembolsados y financieros"),
        L("c10.ivar", "(−) IVA recuperable de costos fijos", ("x", lambda X, y: f"-{CS(X, 'IVAR', y)}")),
        L("c10.dep", "GASTOS NO DESEMBOLSADOS — depreciación", ("s", "DEP"), b=True),
        L("c10.fin", "GASTOS FINANCIEROS — intereses, IVA y comisiones", ("s", "INT"), b=True),
        L("c10.tot", "TOTAL COSTOS ANUALES", ("c", cSum(["c10.prod", "c10.adm", "c10.ivar", "c10.dep", "c10.fin"])), b=True),
        L("c10.uni", "Piezas vendidas (todos los productos)", ("x", UNITS), NF_INT, u="u"),
        L("c10.cpz", "Costo total por pieza vendida", ("c", cDiv("c10.tot", "c10.uni")), NF_USD2, u="USD/u", tot="c"),
        S("Detalle de costos fijos por concepto (hoja Costos, sección C)"),
    ]
    fcr = lambda j: CO["fc_first"] + j - 1
    for j in range(1, NFIJ + 1):
        def fcx(X, y, j=j):
            amt = f"N(INDEX(FC_{OV_(X)},{j}))"
            conv = f'IF(INDEX(FC_Mon,{j})="PYG",1/{TCY(X, y)},1)'
            if y == 0:
                return f"{amt}*{CN(X, 0)}*{conv}*IF(N(INDEX(FC_Desde,{j}))=0,1,0)"
            return (f"{amt}*(1+N(INDEX(FC_g,{j})))^{y - 1}*{CN(X, y)}*{conv}"
                    f"*IF(OR(N(INDEX(FC_Desde,{j}))=0,{y}>=N(INDEX(FC_Desde,{j}))),1,0)")
        lines.append(L(f"c10.fc.{j}", f'=IF(Costos!$B${fcr(j)}="","Concepto {j}",Costos!$B${fcr(j)}&" ("&Costos!$C${fcr(j)}&")")', ("x", fcx), grp=j > VIS_FC))
    lines.append(L("c10.fcaj", "Ajuste (detalle mensual, datos reales o TC por año calendario)",
                   ("c", cExpr("{c10.fp}+{c10.fa}-(" + "+".join("{c10.fc." + str(j) + "}" for j in range(1, NFIJ + 1)) + ")")), pub=False))
    cu_sheet("C10a Costos", "Cuadro 10a — Costos anuales proyectados", "cuadro 10a, pág. 19", lines,
             ["Salarios, prestaciones y alimentación se separan del costo de personal del motor (cargas sociales y alimentación de la hoja Supuestos).",
              "El modelo clasifica al personal en directo, indirecto de producción y administración (hoja Costos). Los costos fijos incluyen el IVA no recuperable.",
              f"Conceptos de costos fijos {VIS_FC + 1} a {NFIJ} agrupados."])


# ---------------------------------------------------------------------------- C10b costo unitario
def build_c10b():
    cd = lambda X, y, i: "(" + "+".join(MATP(X, y, i, n, x) for n, x in [("MP", "MP"), ("INS", "OV"), ("OV", "OV")]) + ")"
    lines = [L("c10b.tot", "COSTOS TOTALES (cuadro 10a)", ("x", lambda X, y: CSUM(X, y, TOTCOST)), b=True),
             L("c10b.cdt", "Costo directo (materia prima, insumos, consumibles y flete) de todos los productos",
               ("c", cSum([f"c10b.cd.{i}" for i in range(1, NP + 1)]))),
             L("c10b.ot", "Otros costos a asignar (total − costo directo)", ("c", cExpr("{c10b.tot}-{c10b.cdt}")))]
    for i in range(1, NP + 1):
        g = i > VIS_P
        lines += [
            L(f"c10b.cd.{i}", pname(i, " — costo directo"), ("x", lambda X, y, i=i: cd(X, y, i)), grp=g),
            L(f"c10b.u.{i}", "   unidades vendidas", ("s", f"u{i}"), NF_INT, u="u", grp=g),
            L(f"c10b.s.{i}", "   ventas", ("s", f"s{i}"), grp=g),
            L(f"c10b.cdu.{i}", "   costo directo unitario", ("c", cDiv(f"c10b.cd.{i}", f"c10b.u.{i}")), NF_USD2, u="USD/u", tot="c", grp=g),
            L(f"c10b.mbu.{i}", "   margen bruto unitario (precio − costo directo)",
              ("c", lambda c, R, i=i: f"IF(N({c}{R[f'c10b.u.{i}']})=0,\"\",({c}{R[f'c10b.s.{i}']}-{c}{R[f'c10b.cd.{i}']})/{c}{R[f'c10b.u.{i}']})"),
              NF_USD2, u="USD/u", tot="c", grp=g),
            L(f"c10b.mbp.{i}", "   margen bruto % (s/ precio)", ("c", lambda c, R, i=i: f"IF(N({c}{R[f'c10b.s.{i}']})=0,\"\",1-{c}{R[f'c10b.cd.{i}']}/{c}{R[f'c10b.s.{i}']})"),
              NF_PCT, u="%", tot="c", grp=g),
            L(f"c10b.oa.{i}", "   otros costos asignados (proporcional al costo directo)",
              ("c", lambda c, R, i=i: f"IF(OR(NOT(ISNUMBER({c}{R[f'c10b.cd.{i}']})),N({c}{R['c10b.cdt']})=0),\"\",{c}{R['c10b.ot']}*{c}{R[f'c10b.cd.{i}']}/{c}{R['c10b.cdt']})"), grp=g),
            L(f"c10b.ct.{i}", "   costo total asignado", ("c", cSum([f"c10b.cd.{i}", f"c10b.oa.{i}"])), grp=g),
            L(f"c10b.ctu.{i}", "   costo total por unidad", ("c", cDiv(f"c10b.ct.{i}", f"c10b.u.{i}")), NF_USD2, u="USD/u", tot="c", grp=g),
        ]
    cu_sheet("C10b Costo unitario", "Cuadro 10b — Costos unitarios por producto", "cuadro 10b, pág. 20", lines,
             ["Costo directo por producto con los supuestos anuales. Otros costos asignados en proporción al costo directo (mismo criterio que el plan).",
              "Margen bruto % = margen ÷ precio. Ver la explicación en el bloque C sobre el porcentaje publicado en el plan.",
              f"Productos {VIS_P + 1} a {NP} agrupados."])


# ---------------------------------------------------------------------------- C11 rubros
def build_c11():
    rub = [("mo", "Mano de obra y cargas sociales (incluye alimentación)", lambda X, y: CSUM(X, y, ["dir_f", "ind_f", "adm_f"])),
           ("mp", "Materia prima, insumos, consumibles y fletes", lambda X, y: CSUM(X, y, ["mp_f", "ins_f", "ov_f"])),
           ("ot", "Mantenimiento, servicios, alquiler y otros costos", lambda X, y: CSUM(X, y, ["fp_f", "fa_f", "pkg_f", "vs_f", "impp_f", "impa_f", "-IVAR"])),
           ("imp", "Impuestos, tasas y despachos (tributo de maquila y despachos)", lambda X, y: CSUM(X, y, ["trib_f", "desp_f"])),
           ("fin", "Gastos financieros", lambda X, y: CS(X, "INT", y)),
           ("dep", "Gastos no desembolsados (depreciación)", lambda X, y: CS(X, "DEP", y))]
    lines = [S("Cuadro 11 — Costos por rubros"), L("c11.tot", "COSTOS TOTALES", ("x", lambda X, y: CSUM(X, y, TOTCOST)), b=True)]
    lines += [L(f"c11.{k}", t, ("x", fn)) for k, t, fn in rub]
    lines.append(L("c11.aj", "Ajuste por datos reales (total del motor − suma de rubros)",
                   ("c", cExpr("{c11.tot}-(" + "+".join("{c11." + k + "}" for k, *_ in rub) + ")")), pub=False))
    lines += [S("Cuadro 12 — Estructura de costos (% del costo total)"), L("c12.tot", "COSTOS TOTALES", ("c", lambda c, R: "1"), NF_PCT, b=True, u="%", tot="c", pub=False)]
    lines += [L(f"c12.{k}", t, ("c", cDiv(f"c11.{k}", "c11.tot")), NF_PCT, u="%", tot="c") for k, t, _ in rub]
    cu_sheet("C11 Rubros", "Cuadros 11 y 12 — Costos por rubros y estructura de costos", "cuadros 11 y 12, pág. 21", lines,
             ["Rubros agrupados para que la comparación sea exacta: el plan separa «reparación y mantenimiento», «electricidad, combustibles y servicios» y «otros costos»; "
              "aquí se comparan juntos."])


# ---------------------------------------------------------------------------- C13 deuda e IVA
def build_c13():
    ints = lambda X, y: "(" + "+".join(CS(X, f"L{j}_int", y) for j in range(1, NL + 1)) + "+" + CS(X, "LCI", y) + ")"
    lines = [
        S("Cuadro 13 — Servicio de la deuda (préstamos)"),
        L("c13.des", "Préstamos: desembolsos", ("s", "DESL")),
        L("c13.sa", "Saldo antes de amortización", ("c", cSum(["c13.sal", "c13.am"])), tot=None),
        L("c13.am", "Amortizaciones (cuotas de capital)", ("s", "AML"), b=True),
        L("c13.sal", "Saldo después de amortización", ("x", lambda X, y: "(" + "+".join(CE(X, f"L{j}_sal", y) for j in range(1, NL + 1)) + ")"), tot=None),
        L("c13.int", "Intereses (sin IVA)", ("x", ints)),
        L("c13.iva", "IVA sobre intereses", ("x", lambda X, y: f"{CS(X, 'INT', y)}-{ints(X, y)}")),
        L("c13.fin", "Total gastos financieros", ("s", "INT"), b=True),
        L("c13.ds", "Servicio de la deuda (amortización + gastos financieros)", ("c", cSum(["c13.am", "c13.fin"])), b=True),
        S("Línea de crédito rotativa"),
        L("c13.lcd", "Uso de la línea", ("s", "LCD"), pub=False),
        L("c13.lcr", "Devolución de la línea", ("s", "LCR"), pub=False),
        L("c13.lcs", "Saldo de la línea al cierre", ("e", "LCS"), tot=None, pub=False),
        S("Detalle por préstamo"),
    ]
    for j in range(1, NL + 1):
        g = j > VIS_L
        lines += [L(f"c13.l{j}d", f"Préstamo {j}: desembolso", ("s", f"L{j}_des"), grp=g, pub=False),
                  L(f"c13.l{j}i", f"Préstamo {j}: interés (sin IVA)", ("s", f"L{j}_int"), grp=g, pub=False),
                  L(f"c13.l{j}a", f"Préstamo {j}: amortización", ("s", f"L{j}_am"), grp=g, pub=False),
                  L(f"c13.l{j}s", f"Préstamo {j}: saldo al cierre", ("e", f"L{j}_sal"), tot=None, grp=g, pub=False)]
    lines += [
        S("Cuadro 14 — Cálculo del IVA"),
        L("c14.deb", "IVA débito por ventas (exportación de maquila: 0)", ("t", 0)),
        L("c14.cf", "IVA contenido en costos fijos gravados", ("s", "ivac_f")),
        L("c14.rec", "IVA recuperable aplicado (reduce costos)", ("s", "IVAR")),
        L("c14.sal", "Saldo de crédito fiscal al cierre", ("e", "ivab"), tot=None),
        L("c14.int", "IVA sobre intereses (se registra como costo financiero)", ("c", cR("c13.iva"))),
    ]
    cu_sheet("C13 Deuda e IVA", "Cuadros 13 y 14 — Servicio de la deuda y cálculo del IVA", "cuadros 13 y 14, pág. 22", lines,
             ["Cronograma mensual (sistema francés) sumado por año de operación. El año 0 solo tiene datos si hay meses de implantación."])


# ---------------------------------------------------------------------------- C15 resultados
def build_c15():
    lines = [
        L("c15.vtas", "Ingresos netos", ("s", "vtas"), b=True),
        L("c15.prod", "Costos de producción", ("x", lambda X, y: CSUM(X, y, ["MAT", "OVAR", "PERP", "FIJP", "TRIB"]))),
        L("c15.adm", "Gastos de administración y ventas (neto de IVA recuperable)", ("x", lambda X, y: CSUM(X, y, ["PERA", "ADM", "-IVAR"]))),
        L("c15.ebitda", "FONDOS GENERADOS EN OPERACIONES — EBITDA", ("s", "EBITDA"), b=True),
        L("c15.dep", "Gastos no desembolsados (depreciación)", ("s", "DEP")),
        L("c15.fin", "Gastos financieros", ("s", "INT")),
        L("c15.uai", "UTILIDAD ANTES DEL IMPUESTO A LA RENTA", ("c", cExpr("{c15.ebitda}-{c15.dep}-{c15.fin}")), b=True),
        L("c15.ire", "Impuesto a la renta — IRE", ("s", "IRE")),
        L("c15.ni", "UTILIDAD DESPUÉS DE IMPUESTO", ("s", "NI"), b=True),
        L("c15.dec", "Dividendos sobre la utilidad del año (política × utilidad)",
          ("x", lambda X, y: f"INDEX(Payout_{OV_(X)},1,{max(1, y)})*MAX(0,{CS(X, 'NI', y)})" if y else "0")),
        L("c15.ret", "Retención anual de utilidades", ("c", cExpr("{c15.ni}-{c15.dec}"))),
        L("c15.div", "Dividendos pagados en el año (resultado del año anterior, con caja disponible)", ("s", "DIV"), pub=False),
        L("c15.mn", "Utilidad neta sobre ventas", ("c", cDiv("c15.ni", "c15.vtas")), NF_PCT, u="%", tot="c"),
        L("c15.pol", "Política de dividendos (% de la utilidad del año)", ("x", lambda X, y: f"INDEX(Payout_{OV_(X)},1,{max(1, y)})"), NF_PCT, u="%", tot=None, y0=False),
    ]
    cu_sheet("C15 Resultados", "Cuadro 15 — Resultados operacionales proyectados", "cuadro 15, pág. 23", lines,
             ["Costos de producción incluyen el tributo de maquila (como el plan). El modelo calcula la política de dividendos por año del modelo; "
              "los paga en el año siguiente, en el mes definido en Supuestos, solo si hay caja sobre el mínimo."])


# ---------------------------------------------------------------------------- C16 capital de trabajo
def build_c16():
    stk = lambda X, y: f"{CE(X, 'MAT', y)}*{X}_DInv/DiasMes"
    pmp = lambda X, y: f"{CE(X, 'MAT', y)}*{X}_DPOmp/DiasMes"
    dd = lambda k: f"=IF(cu_P=1,P_{k},V_{k})"
    lines = [
        L("c16.bruto", "CAPITAL DE TRABAJO BRUTO (activo corriente operativo)", ("c", cSum(["c16.stk", "c16.pt", "c16.cxc", "c16.iva"])), b=True, tot=None, inB=True),
        L("c16.caja", "Caja y bancos (caja mínima operativa; en el modelo no forma parte del capital de trabajo)", ("e", "CMIN"), u=dd("CajaMin"), tot=None),
        L("c16.sue", "Sueldos y anticipos (no incluido en el modelo)", ("t", 0), tot=None),
        L("c16.stk", "Stock de materiales y reservas", ("x", stk), u=dd("DInv"), tot=None),
        L("c16.pt", "Productos terminados", ("x", lambda X, y: f"{CE(X, 'inv', y)}-{stk(X, y)}"), u=dd("DPT"), tot=None),
        L("c16.cxc", "Ventas a crédito (cuentas por cobrar)", ("e", "cxc"), u=dd("DSO"), tot=None),
        L("c16.iva", "Crédito fiscal IVA", ("e", "ivab"), tot=None),
        L("c16.pas", "PASIVO CORRIENTE (proveedores)", ("e", "prov"), b=True, tot=None),
        L("c16.pmp", "Créditos de proveedores de materia prima", ("x", pmp), u=dd("DPOmp"), tot=None),
        L("c16.pot", "Créditos de proveedores menores y otros", ("x", lambda X, y: f"{CE(X, 'prov', y)}-{pmp(X, y)}"), u=dd("DPOot"), tot=None),
        L("c16.ctn", "CAPITAL DE TRABAJO NETO", ("e", "ctn"), b=True, tot=None),
        L("c16.var", "Variación del capital de trabajo en el año (necesidad de financiamiento)", ("s", "dct")),
    ]
    cu_sheet("C16 Capital trabajo", "Cuadro 16 — Capital de trabajo (saldos al cierre de cada año)", "cuadro 16, pág. 24", lines,
             ["La columna «Unidad» muestra los días aplicados (versión seleccionada). Saldos al último mes de cada año de operación.",
              "Stock de materiales = materiales del último mes × días de inventario; producto terminado = inventario total − stock de materiales."])


# ---------------------------------------------------------------------------- C17 punto de equilibrio
def build_c17():
    var = [("mp", "Materia prima, insumos, consumibles y fletes", ["mp_f", "ins_f", "ov_f"]), ("desp", "Despachos de importación", ["desp_f"]),
           ("pkg", "Packaging", ["pkg_f"]), ("vs", "Seguro de caución, comisiones y logística de venta", ["vs_f"]),
           ("trib", "Tributo de maquila", ["trib_f"]), ("imp", "Imprevistos de producción", ["impp_f"])]
    fij = [("per", "Personal de producción y administración (salarios, cargas y alimentación)", ["dir_f", "ind_f", "adm_f"]),
           ("fij", "Costos fijos de producción y administración, e imprevistos de administración (neto de IVA recuperable)", ["fp_f", "fa_f", "impa_f", "-IVAR"]),
           ("fin", "Gastos financieros", ["INT"]), ("dep", "Gastos no desembolsados (depreciación)", ["DEP"])]
    lines = [L("c17.vtas", "Ingresos netos", ("s", "vtas"), b=True),
             L("c17.tot", "COSTOS TOTALES", ("x", lambda X, y: CSUM(X, y, TOTCOST)), b=True),
             L("c17.var", "COSTOS VARIABLES", ("c", cSum([f"c17.v{k}" for k, *_ in var])), b=True, inB=True)]
    lines += [L(f"c17.v{k}", t, ("x", lambda X, y, ks=ks: CSUM(X, y, ks))) for k, t, ks in var]
    lines.append(L("c17.fij", "COSTOS FIJOS (total − variables)", ("c", cExpr("{c17.tot}-{c17.var}")), b=True))
    lines += [L(f"c17.f{k}", t, ("x", lambda X, y, ks=ks: CSUM(X, y, ks))) for k, t, ks in fij]
    lines += [
        L("c17.faj", "Ajuste por datos reales", ("c", cExpr("{c17.fij}-(" + "+".join("{c17.f" + k + "}" for k, *_ in fij) + ")")), pub=False),
        L("c17.mc", "Margen de contribución (ingresos − variables)", ("c", cExpr("{c17.vtas}-{c17.var}"))),
        L("c17.mcp", "Margen de contribución %", ("c", cDiv("c17.mc", "c17.vtas")), NF_PCT, u="%", tot="c"),
        L("c17.pev", "PUNTO DE EQUILIBRIO en ventas", ("c", lambda c, R: f"IF(N({c}{R['c17.mcp']})<=0,\"\",{c}{R['c17.fij']}/{c}{R['c17.mcp']})"), b=True, tot=None),
        L("c17.uni", "Unidades vendidas (todos los productos)", ("x", UNITS), NF_INT, u="u"),
        L("c17.peu", "PUNTO DE EQUILIBRIO en unidades (mezcla de productos del año)",
          ("c", lambda c, R: f"IF(N({c}{R['c17.vtas']})=0,\"\",{c}{R['c17.pev']}*{c}{R['c17.uni']}/{c}{R['c17.vtas']})"), NF_INT, b=True, u="u", tot=None),
        L("c17.ms", "Margen de seguridad (ventas sobre el equilibrio)", ("c", lambda c, R: f"IF(N({c}{R['c17.vtas']})=0,\"\",1-{c}{R['c17.pev']}/{c}{R['c17.vtas']})"),
          NF_PCT, u="%", tot=None),
    ]
    cu_sheet("C17 Equilibrio", "Cuadro 17 — Punto de equilibrio", "cuadro 17, pág. 25", lines,
             ["Punto de equilibrio en ventas = costos fijos ÷ margen de contribución %. En unidades, con la mezcla de productos de cada año.",
              "Variables: materiales, fletes, despachos, packaging, seguros de venta, tributo e imprevistos de producción. Fijos: personal, costos fijos, financieros y depreciación."])


# ---------------------------------------------------------------------------- C18 origen y aplicación de fondos
def build_c18():
    lines = [
        S("Fuentes"),
        L("c18.fue", "TOTAL DE FUENTES", ("c", cSum(["c18.ebitda", "c18.sant", "c18.prest", "c18.lcd", "c18.ap"])), b=True),
        L("c18.ebitda", "Fondos generados en operaciones (EBITDA)", ("s", "EBITDA")),
        L("c18.sant", "Saldo de caja del año anterior", ("x", lambda X, y: "0" if y == 0 else CE(X, "CAJA", y - 1)), tot=None),
        L("c18.prest", "Préstamos bancarios", ("s", "DESL")),
        L("c18.lcd", "Línea de crédito rotativa: usos", ("s", "LCD")),
        L("c18.ap", "Aportes de socios", ("s", "APORTE")),
        S("Usos"),
        L("c18.uso", "TOTAL DE USOS", ("c", cSum(["c18.inv", "c18.ct", "c18.am", "c18.lcr", "c18.fin", "c18.ire", "c18.div"])), b=True),
        L("c18.inv", "Inversiones fijas", ("s", "CAPEX")),
        L("c18.ct", "Necesidad de capital de trabajo (variación)", ("s", "dct")),
        L("c18.am", "Amortización de créditos", ("s", "AML")),
        L("c18.lcr", "Línea de crédito rotativa: devoluciones", ("s", "LCR")),
        L("c18.fin", "Gastos financieros", ("s", "INT")),
        L("c18.ire", "Impuesto a la renta — IRE", ("s", "IRE")),
        L("c18.div", "Dividendos", ("s", "DIV")),
        S("Saldo"),
        L("c18.aj", "Ajuste de conciliación con la caja real (solo con reales)", ("s", "AJ"), pub=False),
        L("c18.sal", "SALDO ANUAL — caja al cierre", ("c", cExpr("{c18.fue}-{c18.uso}+{c18.aj}")), b=True, tot=None),
        L("c18.chk", "Control: saldo − caja del motor (0 = OK)",
          ("x", lambda X, y: f"ROUND({CS(X, 'EBITDA', y)}+{'0' if y == 0 else CE(X, 'CAJA', y - 1)}+{CSUM(X, y, ['DESL', 'LCD', 'APORTE'])}"
                             f"-{CSUM(X, y, ['CAPEX', 'dct', 'AML', 'LCR', 'INT', 'IRE', 'DIV'])}+{CS(X, 'AJ', y)}-{CE(X, 'CAJA', y)},2)"),
          NF_USD2, tot=None, pub=False),
    ]
    cu_sheet("C18 Fondos", "Cuadro 18 — Origen y aplicación de fondos", "cuadro 18, pág. 26", lines,
             ["Fuentes − usos = caja al cierre del año. La fila de control debe ser 0 en todos los años."])


# ---------------------------------------------------------------------------- C19 VAN, TIR y payback
def build_c19():
    rs = lambda k: f"IF(cu_P=1,Resultados!$C${RS['k_' + k]},Resultados!$D${RS['k_' + k]})"
    acc = lambda k: k.replace('p_act', 'p_acu').replace('fcffd', 'acp').replace('eqd', 'ace')
    run = lambda k: (lambda c, R: f'IF(NOT(ISNUMBER({c}{R[k]})),"",' + (f"{c}{R[k]}" if c == "D" else f"N({pc(c)}{R[acc(k)]})+{c}{R[k]}") + ")")
    lines = [
        S("A. Flujo del proyecto (FCFF, sin financiamiento) — criterio del modelo"),
        L("c19.ebitda", "Fondos generados en operaciones (EBITDA)", ("s", "EBITDA"), pub=False),
        L("c19.ire", "(−) Impuesto a la renta", ("x", lambda X, y: f"-{CS(X, 'IRE', y)}"), pub=False),
        L("c19.dct", "(−) Formación de capital de trabajo", ("x", lambda X, y: f"-{CS(X, 'dct', y)}"), pub=False),
        L("c19.capex", "(−) Inversiones", ("x", lambda X, y: f"-{CS(X, 'CAPEX', y)}"), pub=False),
        L("c19.vt", "(+) Valor residual: capital de trabajo y activo fijo realizables", ("s", "VT"), pub=False),
        L("c19.fcff", "FLUJO LIBRE DEL PROYECTO (FCFF)", ("s", "FCFF"), b=True, pub=False),
        L("c19.fcffd", "FCFF descontado al WACC (mensual)", ("x", lambda X, y: f"SUMPRODUCT({CSEL(X, y)}*{CR(X, 'FCFF')}*{CR(X, 'DFP')})"), pub=False),
        L("c19.acp", "FCFF descontado acumulado", ("c", run("c19.fcffd")), tot=None, pub=False),
        S("B. Flujo de los socios — criterio del modelo"),
        L("c19.eq", "Flujo de los socios (− aportes − cobertura de caja negativa + dividendos + valor final)", ("s", "EQ"), pub=False),
        L("c19.eqd", "Flujo de los socios descontado al Ke (mensual)", ("x", lambda X, y: f"SUMPRODUCT({CSEL(X, y)}*{CR(X, 'EQ')}*{CR(X, 'DFE')})"), pub=False),
        L("c19.ace", "Flujo de los socios descontado acumulado", ("c", run("c19.eqd")), tot=None, pub=False),
        S("C. Flujo de beneficios netos — método del plan (cuadro 19), calculado con las cifras del modelo"),
        L("c19.p_ing", "INGRESOS TOTALES (EBITDA + valor residual)", ("x", lambda X, y: f"{CS(X, 'EBITDA', y)}+{CS(X, 'VT', y)}"), b=True),
        L("c19.p_inv", "Inversiones", ("s", "CAPEX")),
        L("c19.p_ct", "Formación de capital operativo", ("s", "dct")),
        L("c19.p_am", "Amortización de créditos", ("s", "AML")),
        L("c19.p_fin", "Gastos financieros", ("s", "INT")),
        L("c19.p_ire", "Impuesto a la renta — IRE", ("s", "IRE")),
        L("c19.p_egr", "EGRESOS TOTALES", ("c", cSum(["c19.p_inv", "c19.p_ct", "c19.p_am", "c19.p_fin", "c19.p_ire"])), b=True),
        L("c19.p_flu", "FLUJO DE BENEFICIOS NETOS", ("c", cExpr("{c19.p_ing}-{c19.p_egr}")), b=True),
        L("c19.p_fac", "Factor de descuento anual al WACC", ("c", lambda c, R: f"1/(1+IF(cu_P=1,P_WACC,V_WACC))^{c}{R['c19.yr']}"), NF_USD4, u="", tot=None, pub=False),
        L("c19.p_act", "Flujo actualizado", ("c", lambda c, R: f'IF(NOT(ISNUMBER({c}{R["c19.p_flu"]})),"",{c}{R["c19.p_flu"]}*{c}{R["c19.p_fac"]})')),
        L("c19.p_acu", "Saldo acumulado actualizado", ("c", run("c19.p_act")), tot=None),
        L("c19.yr", "Año n° (auxiliar)", ("c", lambda c, R: str(column_index_from_string(c) - 4)), NF_N, u="", tot=None, pub=False, grp=True),
    ]
    ws, R = cu_sheet("C19 VAN-TIR", "Cuadro 19 — Cálculo de VAN, TIR y payback", "cuadros 19 y siguientes, pág. 27-28", lines,
                     ["A y B: criterio del modelo (flujos mensuales descontados). C: el mismo flujo anual que usa el plan, recalculado con las cifras del modelo, "
                      "para separar la diferencia de cifras de la diferencia de método.",
                      "El método del plan descuenta los pagos de la deuda pero no suma el desembolso del préstamo: mezcla el flujo del proyecto y el del accionista."])
    # resumen de indicadores y ponderación de la tasa (a la derecha no: debajo de las notas)
    r = ws.max_row + 2
    A, B = R["A"], R["B"]
    section(ws, r, "D. Indicadores: modelo frente al plan publicado", 5 + CUY); r += 1
    header(ws, r, ["", "Indicador", "", "Modelo", "Plan", "Diferencia", "Explicación"], 1, 24); r += 1
    rng = lambda RB, k: f"D{RB[k]}:{YC(CUY)}{RB[k]}"
    ind = [
        ("i_van", "VAN — método del plan (flujo de beneficios netos al WACC); plan: VAN publicado", f"={TC_}{A['c19.p_act']}", ("inp", "c19.van"), NF_USD),
        ("i_vanb", "VAN — método del plan; plan: recalculado con los flujos publicados (bloque B)", f"={TC_}{A['c19.p_act']}",
         f'=IF(COUNT({rng(B, "c19.p_flu")})=0,"",{TC_}{B["c19.p_act"]})', NF_USD),
        ("i_tir", "TIR — método del plan; plan: TIR publicada", f'=IFERROR(IRR({rng(A, "c19.p_flu")}),"No converge")', ("inp", "c19.tir"), NF_PCT2),
        ("i_tirb", "TIR — método del plan; plan: recalculada con los flujos publicados", f'=IFERROR(IRR({rng(A, "c19.p_flu")}),"No converge")',
         f'=IF(COUNT({rng(B, "c19.p_flu")})=0,"",IFERROR(IRR({rng(B, "c19.p_flu")}),"No converge"))', NF_PCT2),
        ("i_pb", "Payback — método del plan (año en que el acumulado actualizado es positivo)",
         f'=IF(_xlfn.MINIFS({rng(A, "c19.yr")},{rng(A, "c19.p_acu")},">0",{rng(A, "c19.yr")},">=1")=0,"No recupera",'
         f'_xlfn.MINIFS({rng(A, "c19.yr")},{rng(A, "c19.p_acu")},">0",{rng(A, "c19.yr")},">=1"))', ("inp", "c19.pb"), NF_N),
        ("i_vanp", "VAN del proyecto — FCFF al WACC (criterio del modelo)", f"={rs('van')}", None, NF_USD),
        ("i_tirp", "TIR del proyecto (criterio del modelo)", f"={rs('tir')}", None, NF_PCT2),
        ("i_pbp", "Recupero del proyecto en años (FCFF sin descontar, criterio del modelo)", f"={rs('pb')}", None, NF_USD2),
        ("i_vane", "VAN de los socios — al Ke (criterio del modelo)", f"={rs('vane')}", None, NF_USD),
        ("i_tire", "TIR de los socios (criterio del modelo)", f"={rs('tire')}", None, NF_PCT2),
    ]
    CUR_ROWS["C19 VAN-TIR"]["I"] = {}
    for k, lab, fm, fp, fmt in ind:
        CUR_ROWS["C19 VAN-TIR"]["I"][k] = r
        put(ws, r, 2, lab)
        put(ws, r, 4, fm, fmt=fmt, font=F_B, fill=FL_V)
        if fp:
            if isinstance(fp, tuple):
                inp(ws, r, 5, PLAN.get(fp[1]), fmt)
            else:
                put(ws, r, 5, fp, fmt=fmt, fill=FL_P)
            put(ws, r, 6, f'=IF(AND(ISNUMBER(D{r}),ISNUMBER(E{r})),D{r}-E{r},"")', fmt=fmt)
            put(ws, r, 7, f'=IF(NOT(ISNUMBER(F{r})),"",IF(ABS(F{r})<{TOL.get(fmt, 1)},"Sin diferencia","'
                + (NOTA.get(k) or "Diferencia a explicar").replace('"', "'") + '"))', font=F_S)
        else:
            put(ws, r, 7, NOTA.get(k, ""), font=F_S)
        r += 1
    r += 1
    section(ws, r, "E. Ponderación de la tasa de corte (costo de oportunidad)", 5 + CUY); r += 1
    header(ws, r, ["", "Fuente de capital", "", "USD (modelo)", "Participación", "Costo", "Tasa ponderada"], 1, 24); r += 1
    tot = lambda k: f"IF(cu_P=1,Calculo!$E${ROW['P.' + k]},Calculo!$E${ROW['V.' + k]})"
    r0 = r
    for lab, usd, cost in [("Aportes de socios (total del horizonte)", tot("APORTE"), "IF(cu_P=1,P_Ke,V_Ke)"),
                           ("Préstamos (total desembolsado)", tot("DESL"), "IF(cu_P=1,P_L1_Tasa,V_L1_Tasa)")]:
        put(ws, r, 2, lab)
        put(ws, r, 4, f"={usd}", fmt=NF_USD)
        put(ws, r, 5, f'=IF(N(D{r0 + 2})=0,"",D{r}/D{r0 + 2})', fmt=NF_PCT)
        put(ws, r, 6, f"={cost}", fmt=NF_PCT2)
        put(ws, r, 7, f'=IF(E{r}="","",E{r}*F{r})', fmt=NF_PCT2)
        r += 1
    put(ws, r, 2, "Total — tasa ponderada con la estructura de fondos del modelo", font=F_B)
    put(ws, r, 4, f"=D{r0}+D{r0 + 1}", fmt=NF_USD, font=F_B)
    put(ws, r, 7, f"=SUM(G{r0}:G{r0 + 1})", fmt=NF_PCT2, font=F_B)
    r += 1
    put(ws, r, 2, "Tasa de descuento usada en el modelo (Supuestos: WACC)")
    put(ws, r, 7, "=IF(cu_P=1,P_WACC,V_WACC)", fmt=NF_PCT2, font=F_B)
    r += 1
    put(ws, r, 2, NOTA.get("i_wacc", "El plan pondera solo el aporte inicial y el préstamo; con todos los aportes la tasa ponderada cambia."), font=F_S)


# ---------------------------------------------------------------------------- C20 depreciación
def build_c20():
    ivr = lambda j: IV["first"] + j - 1
    lines = []
    for j in range(1, NCAP + 1):
        def dep(X, y, j=j):
            amt = f"N(INDEX(Inv_D,{j}))" if X == "P" else f"N(INDEX(Inv_BaseV,{j}))"
            svc = f"INDEX(Inv_Svc{X},{j})"
            vida = f"N(INDEX(Inv_F,{j}))"
            s, e = (("1", f"({X}_Inicio-1)") if y == 0 else (f"({X}_Inicio+{12 * (y - 1)})", f"({X}_Inicio+{12 * y - 1})"))
            return f"IF({vida}=0,0,{amt}/({vida}*12)*MAX(0,MIN({e},{svc}+{vida}*12-1,Fin)-MAX({s},{svc})+1))"
        lab = (f'=IF(Inversion!$B${ivr(j)}="","Ítem {j}",Inversion!$B${ivr(j)}&" — USD "&FIXED(IF(cu_P=1,N(Inversion!$D${ivr(j)}),N(Inversion!$U${ivr(j)})),0)'
               f'&" — "&N(Inversion!$F${ivr(j)})&" años")')
        ln = L(f"c20.i.{j}", lab, ("x", dep), grp=j > VIS_CAP, pub=False)
        ln["u"] = f'=IF(Inversion!$C${ivr(j)}="","",Inversion!$C${ivr(j)})'
        lines.append(ln)
    first, last = "c20.i.1", f"c20.i.{NCAP}"
    cls = lambda name: (lambda c, R: f'SUMIF($C${R[first]}:$C${R[last]},"{name}",{c}{R[first]}:{c}{R[last]})')
    lines += [S("Por rubro (como el plan)"),
              L("c20.maq", "Maquinarias y equipos (incluye despacho)", ("c", cls("Maquinaria")), inB=True),
              L("c20.inf", "Adecuación de infraestructura", ("c", cls("Infraestructura")), inB=True),
              L("c20.otr", "Otras inversiones (software, mobiliario y demás)", ("c", cExpr("{c20.tot}-{c20.maq}-{c20.inf}")), inB=True),
              L("c20.tot", "TOTAL DEPRECIACIÓN", ("c", lambda c, R: f"SUM({c}{R[first]}:{c}{R[last]})"), b=True, inB=True),
              L("c20.eng", "Depreciación del motor (hoja Calculo)", ("s", "DEP"), pub=False),
              L("c20.aj", "Ajuste (inversión ejecutada en meses reales)", ("c", cExpr("{c20.eng}-{c20.tot}")), pub=False),
              L("c20.af", "Activo fijo neto al cierre (valor residual contable)", ("e", "AF"), b=True, tot="last")]
    cu_sheet("C20 Depreciación", "Cuadro 20 — Depreciación de activo fijo y valor residual", "cuadro 20, pág. 29", lines,
             ["Depreciación lineal desde el mes de puesta en servicio, según la vida útil de cada ítem (hoja Inversion). Rubro = clase del ítem.",
              f"Ítems {VIS_CAP + 1} a {NCAP} agrupados."])


# ---------------------------------------------------------------------------- RRHH
def build_rrhh():
    prr = lambda j: CO["per_first"] + j - 1
    lab = lambda j: f'=IF(Costos!$B${prr(j)}="","Puesto {j}",Costos!$B${prr(j)}&" ("&Costos!$C${prr(j)}&")")'
    lines = [S("A. Dotación por puesto (personas; supuestos anuales de la hoja Costos)")]
    for j in range(1, NPER + 1):
        lines.append(L(f"rh.d.{j}", lab(j), ("x", lambda X, y, j=j: f"N(INDEX(Dot_{OV_(X)},{j},{y}))"), NF_INT, u="pers.", tot=None, y0=False,
                       grp=j > VIS_PER, pub=False))
    lines.append(S("B. Costo anual base por puesto (salario × 12 × dotación ÷ TC; sin cargas ni alimentación)"))
    for j in range(1, NPER + 1):
        def cost(X, y, j=j):
            return (f"N(INDEX(Sal_{OV_(X)},{j}))*INDEX(IdxSal_{OV_(X)},1,{y})*12*N(INDEX(Dot_{OV_(X)},{j},{y}))"
                    f'*IF(INDEX(Per_Mon,{j})="PYG",1/{TCY(X, y)},1)')
        lines.append(L(f"rh.c.{j}", lab(j), ("x", cost), y0=False, grp=j > VIS_PER, pub=False))
    avg = lambda k: (lambda X, y: f"IF({CN(X, y)}=0,0,{CS(X, k + '_n', y)}/{CN(X, y)})")
    lines += [
        S("C. Totales por clase (motor: incluye detalle mensual y datos reales)"),
        L("rh.ndir", "Personal directo (personas, promedio del año)", ("x", avg("dir")), NF_USD2, u="pers.", tot=None),
        L("rh.nind", "Personal indirecto de producción", ("x", avg("ind")), NF_USD2, u="pers.", tot=None, pub=False),
        L("rh.nadm", "Personal de administración", ("x", avg("adm")), NF_USD2, u="pers.", tot=None, pub=False),
        L("rh.nia", "Indirectos y administración (como el plan)", ("c", cSum(["rh.nind", "rh.nadm"])), NF_USD2, u="pers.", tot=None, inB=True),
        L("rh.ntot", "TOTAL DE PERSONAS", ("c", cSum(["rh.ndir", "rh.nind", "rh.nadm"])), NF_USD2, b=True, u="pers.", tot=None, inB=True),
        L("rh.cdir", "Salarios del personal directo", ("x", lambda X, y: BASE(X, "dir", y))),
        L("rh.cind", "Salarios del personal indirecto de producción", ("x", lambda X, y: BASE(X, "ind", y)), pub=False),
        L("rh.cadm", "Salarios del personal de administración", ("x", lambda X, y: BASE(X, "adm", y)), pub=False),
        L("rh.cia", "Salarios de indirectos y administración (como el plan)", ("c", cSum(["rh.cind", "rh.cadm"])), inB=True),
        L("rh.ctot", "TOTAL SALARIOS (costo base anual)", ("c", cSum(["rh.cdir", "rh.cind", "rh.cadm"])), b=True, inB=True),
        L("rh.car", "Prestaciones sociales", ("x", lambda X, y: f"({BASE(X, 'dir', y)}+{BASE(X, 'ind', y)}+{BASE(X, 'adm', y)})*{X}_Cargas"), pub=False),
        L("rh.ali", "Alimentación", ("x", lambda X, y: f"{ALIM(X, 'dir', y)}+{ALIM(X, 'ind', y)}+{ALIM(X, 'adm', y)}"), pub=False),
        L("rh.tot", "COSTO TOTAL DE PERSONAL", ("x", lambda X, y: CSUM(X, y, ["dir_f", "ind_f", "adm_f"])), b=True, pub=False),
    ]
    cu_sheet("RRHH", "Recursos humanos involucrados en el proyecto", "pág. 30-34", lines,
             ["A y B: por puesto, con los supuestos anuales. C: totales del motor (pueden diferir de B si se cargó detalle mensual o datos reales).",
              f"Puestos {VIS_PER + 1} a {NPER} agrupados. Salarios en Gs convertidos con el TC de cada año."])


# ---------------------------------------------------------------------------- C01 inversión (no anual)
def build_c01():
    ws = WS["C01 Inversión"]
    cu_head(ws, "C01 Inversión", "Cuadro 1 — Inversión y financiamiento; cuadros 2, 5 y 6 — detalle de inversiones", "cuadros 1, 2, 5 y 6, pág. 11-13")
    widths(ws, {"B": 50, "C": 15, "D": 15, "E": 12, "F": 10, "G": 15, "H": 15, "I": 80})
    ivf, ivl = IV["first"], IV["last"]
    cls_rng = f"Inversion!$C${ivf}:$C${ivl}"
    amt = lambda X: "Inv_D" if X == "P" else "Inv_BaseV"
    capt = lambda X: f"Calculo!$E${ROW[X + '.CAPEX']}"
    vv = lambda fn: f"=IF(cu_P=1,{fn('P')},{fn('V')})"
    r = 5
    section(ws, r, "Cuadro 1 — Inversión y financiamiento (USD)", 9); r += 1
    header(ws, r, ["", "Concepto", "", "Modelo (versión)", "", "", "Plan publicado", "Diferencia", "Explicación de la diferencia"], 1, 24); r += 1
    R = {}
    rows = [
        ("c01.inf", "Adecuaciones de infraestructura", vv(lambda X: f'SUMIF({cls_rng},"Infraestructura",{amt(X)})'), False),
        ("c01.maq", "Maquinaria y equipos (CIF + gastos de despacho)", vv(lambda X: f'SUMIF({cls_rng},"Maquinaria",{amt(X)})'), False),
        ("c01.otr", "Otras inversiones (software, mobiliario y demás)", None, False),
        ("c01.fija", "TOTAL INVERSIÓN FIJA", vv(capt), True),
        ("c01.ini", "   pagada hasta el año 1 de operación (inversión inicial)", vv(lambda X: f"{CS(X, 'CAPEX', 0)}+{CS(X, 'CAPEX', 1)}"), False),
        ("c01.post", "   pagada desde el año 2", None, False),
        ("c01.cop", "CAPITAL OPERATIVO INICIAL (capital de trabajo neto al cierre del año 1)", vv(lambda X: CE(X, "ctn", 1)), True),
        ("c01.tini", "TOTAL DE LA INVERSIÓN INICIAL (inversión inicial + capital operativo)", None, True),
        ("c01.ap", "Aportes de socios (total)", vv(lambda X: f"Calculo!$E${ROW[X + '.APORTE']}"), False),
        ("c01.pr", "Préstamos bancarios (total desembolsado)", vv(lambda X: f"Calculo!$E${ROW[X + '.DESL']}"), False),
        ("c01.fue", "TOTAL FUENTES", None, True),
        ("c01.pap", "   % aportes de socios", None, False),
        ("c01.ppr", "   % préstamos", None, False),
    ]
    for k, *_ in rows:
        R[k] = r
        r += 1
    r = R["c01.inf"]
    derived = {"c01.otr": lambda c: f"={c}{R['c01.fija']}-{c}{R['c01.inf']}-{c}{R['c01.maq']}",
               "c01.post": lambda c: f"={c}{R['c01.fija']}-{c}{R['c01.ini']}",
               "c01.tini": lambda c: f"={c}{R['c01.ini']}+{c}{R['c01.cop']}",
               "c01.fue": lambda c: f"={c}{R['c01.ap']}+{c}{R['c01.pr']}",
               "c01.pap": lambda c: f'=IF(N({c}{R["c01.fue"]})=0,"",{c}{R["c01.ap"]}/{c}{R["c01.fue"]})',
               "c01.ppr": lambda c: f'=IF(N({c}{R["c01.fue"]})=0,"",{c}{R["c01.pr"]}/{c}{R["c01.fue"]})'}
    for k, lab, f, bold in rows:
        fmt = NF_PCT if k in ("c01.pap", "c01.ppr") else NF_USD
        put(ws, r, 1, k, font=F_S)
        put(ws, r, 2, lab, font=F_B if bold else F, fill=FL_TOT if bold else None)
        put(ws, r, 4, f if f else derived[k]("D"), fmt=fmt, font=F_B, fill=FL_V)
        if k in derived:
            first = {"c01.otr": "c01.fija", "c01.post": "c01.fija", "c01.tini": "c01.ini", "c01.fue": "c01.ap", "c01.pap": "c01.fue", "c01.ppr": "c01.fue"}[k]
            put(ws, r, 7, f'=IF(G{R[first]}="","",{derived[k]("G")[1:]})', fmt=fmt, fill=FL_P)
        else:
            inp(ws, r, 7, PLAN.get(k), fmt)
        put(ws, r, 8, f'=IF(AND(ISNUMBER(D{r}),ISNUMBER(G{r})),D{r}-G{r},"")', fmt=fmt)
        put(ws, r, 9, f'=IF(NOT(ISNUMBER(H{r})),"",IF(ABS(H{r})<{TOL.get(fmt, 1)},"Sin diferencia","'
            + (NOTA.get(k) or "Diferencia a explicar").replace('"', "'") + '"))', font=F_S)
        r += 1
    r += 1
    section(ws, r, "Cuadros 2, 5 y 6 — Detalle de inversiones por ítem (hoja Inversion)", 9); r += 1
    header(ws, r, ["#", "Ítem", "Clase", "Modelo (versión)", "Mes de pago", "Vida útil", "Plan publicado", "Diferencia", "Explicación de la diferencia"], 1, 24)
    r += 1
    fi = r
    items = PLAN.get("c01.items", [])
    for j in range(1, NCAP + 1):
        q = ivf + j - 1
        put(ws, r, 1, j, font=F_S)
        put(ws, r, 2, f'=IF(Inversion!$B${q}="","",Inversion!$B${q})', font=F_LINK)
        put(ws, r, 3, f'=IF(Inversion!$C${q}="","",Inversion!$C${q})', font=F_LINK)
        put(ws, r, 4, f"=IF(cu_P=1,N(Inversion!$D${q}),N(Inversion!$U${q}))", fmt=NF_USD, fill=FL_V)
        put(ws, r, 5, f'=IF(Inversion!$B${q}="","",IF(cu_P=1,Inversion!$S${q},Inversion!$Q${q}))', fmt=NF_N)
        put(ws, r, 6, f'=IF(Inversion!$B${q}="","",N(Inversion!$F${q}))', fmt=NF_N)
        inp(ws, r, 7, items[j - 1] if j - 1 < len(items) else None, NF_USD)
        put(ws, r, 8, f'=IF(ISNUMBER(G{r}),D{r}-G{r},"")', fmt=NF_USD)
        put(ws, r, 9, f'=IF(NOT(ISNUMBER(H{r})),"",IF(ABS(H{r})<2,"Sin diferencia","'
            + ((D.get("plan_notas_items") or {}).get(j) or NOTA.get("c01.item") or "Diferencia a explicar").replace('"', "'") + '"))', font=F_S)
        if j > VIS_CAP:
            ws.row_dimensions[r].outlineLevel = 1
            ws.row_dimensions[r].hidden = True
        r += 1
    put(ws, r, 2, "TOTAL", font=F_B, fill=FL_TOT)
    for c in "DGH":
        put(ws, r, ord(c) - 64, f"=SUM({c}{fi}:{c}{r - 1})", fmt=NF_USD, font=F_B, fill=FL_TOT)
    r += 2
    section(ws, r, "Notas", 9); r += 1
    for t in ["Modelo: presupuesto = monto del ítem; vigente = base depreciable vigente (ejecutado + comprometido + por contratar con el escenario).",
              "Mes de pago: presupuesto = mes de pago; vigente = mes de pago pendiente usado. El plan registra la inversión inicial en el año 0; "
              "si el modelo empieza a operar en el mes 1, la paga en el año 1.",
              f"Ítems {VIS_CAP + 1} a {NCAP} agrupados."]:
        put(ws, r, 2, t, font=F_S)
        r += 1
    ws.freeze_panes = "C7"
    CUR_ROWS["C01 Inversión"] = {"A": R}


# ---------------------------------------------------------------------------- índice
def build_cuadros_indice():
    ws = WS["Cuadros"]
    ws.sheet_view.showGridLines = False
    widths(ws, {"A": 3, "B": 58, "C": 18, "D": 18, "E": 16, "F": 12, "G": 90})
    put(ws, 1, 1, "Cuadros financieros — misma estructura que el plan de negocio", font=F_T)
    put(ws, 2, 1, '="Plan de referencia: "&IF(n_Version="","(sin versión)",n_Version)&"   |   Cada hoja tiene el modelo (A), el plan publicado (B) y la diferencia (C)."', font=F_S)
    r = 4
    section(ws, r, "1. Versión que muestran los cuadros", 7); r += 1
    put(ws, r, 2, "Versión", font=F_B)
    inp(ws, r, 3, "Vigente")
    dv(ws, f"C{r}", ["Vigente", "Presupuesto original"])
    nm("cu_Ver", "Cuadros", f"C{r}")
    put(ws, r, 4, "Vigente = real hasta el último cierre + proyección con el escenario elegido. Presupuesto original = plan aprobado con la metodología del modelo.",
        font=F_S)
    r += 1
    put(ws, r, 2, "Indicador de versión (1 = presupuesto original)", font=F_S)
    put(ws, r, 3, '=IF(cu_Ver="Presupuesto original",1,0)', fmt=NF_N, font=F_S)
    nm("cu_P", "Cuadros", f"C{r}")
    r += 1
    put(ws, r, 2, "Escenario / último mes cerrado")
    put(ws, r, 3, '=n_Escenario&" / "&IF(Corte=0,"sin reales","mes "&Corte)', font=F_B)
    r += 2
    section(ws, r, "2. Índice de cuadros", 7); r += 1
    header(ws, r, ["", "Hoja", "Cuadros del plan", "Páginas del plan", "Ir", "", "Contenido"], 1, 24); r += 1
    for hoja, cuad, cont, pag in CU_HOJAS:
        put(ws, r, 2, hoja, font=F_B)
        put(ws, r, 3, cuad, al=CEN)
        put(ws, r, 4, pag, al=CEN)
        put(ws, r, 5, f'=HYPERLINK("#\'{hoja}\'!A1","Abrir")', font=Font(name=FONT, size=9, color="0563C1", underline="single"), al=CEN)
        put(ws, r, 7, cont)
        r += 1
    r += 1
    section(ws, r, "3. Resumen: modelo frente al plan publicado", 7); r += 1
    header(ws, r, ["", "Concepto", "Modelo", "Plan publicado", "Diferencia", "Diferencia %", "Explicación"], 1, 24); r += 1
    ref = lambda hoja, blk, k, col: f"'{hoja}'!${col}${CUR_ROWS[hoja][blk][k]}"
    expl = lambda rr, fmt, nk: (f'=IF(NOT(ISNUMBER(E{rr})),"",IF(ABS(E{rr})<{TOL.get(fmt, 1)},"Sin diferencia","'
                                + (NOTA.get(nk) or "Diferencia a explicar").replace('"', "'") + '"))')
    res = [
        ("Inversión fija inicial (pagada hasta el año 1)", ("C01 Inversión", "c01.ini", "D", "G"), NF_USD, "s_ini"),
        ("Capital operativo inicial (capital de trabajo neto del año 1)", ("C01 Inversión", "c01.cop", "D", "G"), NF_USD, "s_cop"),
        ("Ventas de los 10 años", ("C08 Ingresos", "c08.vtas", TC_, TC_), NF_USD, "s_vtas"),
        ("Costos totales de los 10 años", ("C10a Costos", "c10.tot", TC_, TC_), NF_USD, "s_cost"),
        ("EBITDA de los 10 años", ("C15 Resultados", "c15.ebitda", TC_, TC_), NF_USD, "s_ebitda"),
        ("Utilidad neta de los 10 años", ("C15 Resultados", "c15.ni", TC_, TC_), NF_USD, "s_ni"),
        ("Capital de trabajo neto al cierre del año 10", ("C16 Capital trabajo", "c16.ctn", YC(CUY), YC(CUY)), NF_USD, "s_ctn"),
        ("Caja al cierre del año 10", ("C18 Fondos", "c18.sal", YC(CUY), YC(CUY)), NF_USD, "s_caja"),
        ("Punto de equilibrio del año 1 (unidades)", ("C17 Equilibrio", "c17.peu", YC(1), YC(1)), NF_INT, "s_pe"),
        ("Personal total del año 1 (personas)", ("RRHH", "rh.ntot", YC(1), YC(1)), NF_USD2, "s_pers"),
    ]
    for lab, (hoja, k, cm, cp), fmt, nk in res:
        put(ws, r, 2, lab)
        if hoja == "C01 Inversión":
            mref, pref = f"'{hoja}'!$D${CUR_ROWS[hoja]['A'][k]}", f"'{hoja}'!$G${CUR_ROWS[hoja]['A'][k]}"
        else:
            mref, pref = ref(hoja, "A", k, cm), ref(hoja, "B", k, cp)
        put(ws, r, 3, f"={mref}", fmt=fmt, font=F_B, fill=FL_V)
        put(ws, r, 4, f'=IF(ISNUMBER({pref}),{pref},"")', fmt=fmt, fill=FL_P)
        put(ws, r, 5, f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),C{r}-D{r},"")', fmt=fmt)
        put(ws, r, 6, f'=IF(AND(ISNUMBER(E{r}),N(D{r})<>0),E{r}/D{r},"")', fmt=NF_PCT)
        put(ws, r, 7, expl(r, fmt, nk), font=F_S)
        r += 1
    I = CUR_ROWS["C19 VAN-TIR"]["I"]
    for lab, k, fmt, nk in [("VAN — método del plan (flujo de beneficios netos)", "i_van", NF_USD, "s_van"), ("TIR — método del plan", "i_tir", NF_PCT2, "s_tir"),
                            ("Payback — método del plan (años)", "i_pb", NF_N, "s_pb"), ("VAN del proyecto — criterio del modelo (FCFF)", "i_vanp", NF_USD, "s_vanp"),
                            ("VAN de los socios — criterio del modelo", "i_vane", NF_USD, "s_vane")]:
        put(ws, r, 2, lab)
        put(ws, r, 3, f"='C19 VAN-TIR'!$D${I[k]}", fmt=fmt, font=F_B, fill=FL_V)
        put(ws, r, 4, f"=IF(ISNUMBER('C19 VAN-TIR'!$E${I[k]}),'C19 VAN-TIR'!$E${I[k]},\"\")", fmt=fmt, fill=FL_P)
        put(ws, r, 5, f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),C{r}-D{r},"")', fmt=fmt)
        put(ws, r, 6, f'=IF(AND(ISNUMBER(E{r}),N(D{r})<>0),E{r}/D{r},"")', fmt=NF_PCT)
        put(ws, r, 7, expl(r, fmt, nk) if nk not in ("s_vanp", "s_vane") else NOTA.get(nk, ""), font=F_S)
        r += 1
    r += 1
    section(ws, r, "4. Diferencias de metodología conocidas (editable)", 7); r += 1
    notas = D.get("plan_notas", [])
    for i in range(12):
        put(ws, r, 2, f"{i + 1}.", font=F_S)
        inp(ws, r, 3, notas[i] if i < len(notas) else None)
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
        r += 1
    r += 1
    section(ws, r, "5. Cómo leer los cuadros", 7); r += 1
    for t in ["Años: año 0 = meses de implantación antes del inicio de operación; años 1 a 10 = años de operación (como el plan). Total = suma de los años.",
              "Bloque A (verde): modelo en la versión elegida arriba. Bloque B (naranja): cifras del plan publicado, para cargar a mano (Ecostar ya está cargado).",
              "Bloque C: diferencia modelo − plan, con la explicación de cada línea. Una diferencia sin explicación es un punto a revisar.",
              "Las filas de productos, puestos, costos fijos e ítems sin uso están agrupadas: el botón + de la izquierda las muestra.",
              "Las filas «Ajuste por datos reales» separan lo que viene de la hoja Reales o de Detalle_24m; sin reales ni detalle mensual deben ser 0."]:
        put(ws, r, 2, t, font=F_S)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=7)
        r += 1
    ws.freeze_panes = "A4"


def build_cuadros():
    build_c07()
    build_c08()
    build_c09()
    build_c10a()
    build_c10b()
    build_c11()
    build_c13()
    build_c15()
    build_c16()
    build_c17()
    build_c18()
    build_c19()
    build_c20()
    build_rrhh()
    build_c01()
    build_cuadros_indice()


# ============================================================================ MAIN
def marcar_modificados():
    """Formato condicional: celdas vigentes distintas del presupuesto resaltadas en amarillo/azul."""
    pairs = [("Supuestos", f"E6:E{SUPR['p_last']}", "E6<>D6"),
             ("Productos", f"H{PR['first']}:H{PR['last']}", f"H{PR['first']}<>G{PR['first']}"),
             ("Productos", f"J{PR['first']}:J{PR['last']}", f"J{PR['first']}<>I{PR['first']}"),
             ("Productos", f"L{PR['first']}:L{PR['last']}", f"L{PR['first']}<>K{PR['first']}"),
             ("Productos", f"N{PR['first']}:N{PR['last']}", f"N{PR['first']}<>M{PR['first']}"),
             ("Productos", f"D{PR['vol_V']}:M{PR['vol_V'] + NP - 1}", f"D{PR['vol_V']}<>D{PR['vol_O']}"),
             ("Costos", f"F{CO['per_first']}:F{CO['per_last']}", f"F{CO['per_first']}<>E{CO['per_first']}"),
             ("Costos", f"G{CO['dotv_first']}:P{CO['dotv_first'] + NPER - 1}", f"G{CO['dotv_first']}<>G{CO['per_first']}"),
             ("Costos", f"F{CO['fc_first']}:F{CO['fc_first'] + NFIJ - 1}", f"F{CO['fc_first']}<>E{CO['fc_first']}")]
    for sh, rng, cond in pairs:
        WS[sh].conditional_formatting.add(rng, FormulaRule(formula=[cond], fill=FL_IN, font=Font(name=FONT, size=9, color="0000FF", bold=True)))


def main():
    build_inicio()
    build_supuestos()
    build_productos()
    build_costos()
    build_inversion()
    build_reales()
    build_detalle()
    build_calculo()
    build_resultados()
    build_proy()
    build_control()
    build_guia()
    build_cuadros()
    marcar_modificados()
    tabs = {"Inicio": "1F3864", "Guia_Carga": "00B0F0", "Detalle_24m": "FFC000", "Proyeccion_24m": "548235", "Supuestos": "FFC000", "Productos": "FFC000", "Costos": "FFC000", "Inversion": "FFC000", "Reales": "FFC000",
            "Calculo": "7F7F7F", "Resultados": "548235", "Control": "C00000"}
    for s in SHEETS:
        WS[s].sheet_properties.tabColor = tabs.get(s, "2F5597")
        WS[s].sheet_view.zoomScale = 90
    wb.calculation.fullCalcOnLoad = True
    wb.active = 0
    wb.save(OUT)
    print("Guardado:", OUT)


if __name__ == "__main__":
    main()
