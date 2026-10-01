# -*- coding: utf-8 -*-
"""
Generador del Modelo_Financiero_Maquila_CIE_PERG.xlsx

Construye con openpyxl un libro editable, con fórmulas vinculadas, a partir del
plan de negocios «Procesamiento de chapas para gabinetes metálicos - CIE/PERG»
(Mujica & Saldivar, actualizado a octubre de 2025) y del informe de validación
del 29/09/2026.

El libro resultante NO depende de este script: todas las cifras de negocio
se calculan con fórmulas de Excel. El script solo documenta y reproduce la
construcción. Ejecutar:  python3 generar_modelo.py [ruta_salida.xlsx]
"""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule

OUT = sys.argv[1] if len(sys.argv) > 1 else "Modelo_Financiero_Maquila_CIE_PERG.xlsx"

# ---------------------------------------------------------------------------
# Rejilla mensual: mes 0 (inicio de implantación / desembolso) a mes N
# ---------------------------------------------------------------------------
N = 156               # 13 años de rejilla: admite implantación + atraso + 10 años
NC = N + 1            # cantidad de columnas mensuales
FC = 6                # primera columna mensual (F = mes 0)


def mc(m):
    return CL(FC + m)


MC0, MCN = mc(0), mc(N)
NYEARS = 13           # años de proyecto 1..13 en el resumen anual

# ---------------------------------------------------------------------------
# Estilos
# ---------------------------------------------------------------------------
FONT = "Arial"
F_BASE = Font(name=FONT, size=9)
F_IN = Font(name=FONT, size=9, color="0000FF")
F_LINK = Font(name=FONT, size=9, color="008000")
F_BOLD = Font(name=FONT, size=9, bold=True)
F_HDR = Font(name=FONT, size=9, bold=True, color="FFFFFF")
F_TITLE = Font(name=FONT, size=13, bold=True, color="1F3864")
F_SUB = Font(name=FONT, size=9, italic=True, color="595959")
F_SEC = Font(name=FONT, size=10, bold=True, color="1F3864")
F_ALERT = Font(name=FONT, size=9, bold=True, color="C00000")
F_INH = Font(name=FONT, size=9, color="7F7F7F", italic=True)

FILL_IN = PatternFill("solid", fgColor="FFF2CC")      # entrada editable
FILL_PEND = PatternFill("solid", fgColor="FFFF00")    # pendiente de cargar
FILL_HDR = PatternFill("solid", fgColor="1F3864")
FILL_SEC = PatternFill("solid", fgColor="D9E1F2")
FILL_TOT = PatternFill("solid", fgColor="F2F2F2")
FILL_INH = PatternFill("solid", fgColor="EDEDED")     # hereda de la base
FILL_ALERT = PatternFill("solid", fgColor="F8CBAD")
FILL_OK = PatternFill("solid", fgColor="E2EFDA")
FILL_HIST = PatternFill("solid", fgColor="FCE4D6")    # dato histórico PDF

THIN = Side(style="thin", color="BFBFBF")
B_TOP = Border(top=Side(style="thin", color="7F7F7F"))
B_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

NF_USD = '#,##0;(#,##0);"-"'
NF_USD2 = '#,##0.00;(#,##0.00);"-"'
NF_USD4 = '#,##0.0000;(#,##0.0000);"-"'
NF_PCT = '0.0%;(0.0%);"-"'
NF_PCT2 = '0.00%;(0.00%);"-"'
NF_PCT4 = '0.0000%;(0.0000%);"-"'
NF_INT = '#,##0;(#,##0);"-"'
NF_INT0 = '0'
NF_X = '0.00"x";(0.00"x");"-"'
NF_DATE = 'dd/mm/yyyy'
NF_MES = 'mmm-yy'
NF_GS = '#,##0;(#,##0);"-"'

WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

wb = Workbook()
wb.remove(wb.active)

SHEET_ORDER = [
    ("GUIA", "Guia"),
    ("RES", "Resumen"),
    ("SUP", "Supuestos"),
    ("CAL", "Calendario"),
    ("PV", "Produccion_Ventas"),
    ("CP", "Costos_Personal"),
    ("IA", "Inversion_Activos"),
    ("CT", "Capital_Trabajo"),
    ("DT", "Deuda_Tributos"),
    ("RC", "Resultados_Caja"),
    ("RA", "Resultados_Anuales"),
    ("ES", "Escenarios_Sensib"),
    ("DR", "Datos_Reales"),
    ("BP", "Base_PDF"),
    ("VAL", "Validaciones"),
]
SN = dict(SHEET_ORDER)
WS = {}
for code, name in SHEET_ORDER:
    WS[code] = wb.create_sheet(name)

ROWS = {}       # 'PV.clave' -> fila del motor
CUR = [None]    # hoja activa durante la escritura (código)


def add_name(nm, sheet_code, cell):
    """Nombre definido de libro apuntando a una celda absoluta."""
    col = "".join(ch for ch in cell if ch.isalpha())
    row = "".join(ch for ch in cell if ch.isdigit())
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"'{SN[sheet_code]}'!${col}${row}")


def add_range_name(nm, sheet_code, rng):
    wb.defined_names[nm] = DefinedName(nm, attr_text=f"'{SN[sheet_code]}'!{rng}")


def pref(sheet_code):
    return "" if sheet_code == CUR[0] else f"{SN[sheet_code]}!"


def A(key, m):
    """Referencia a la celda del mes m de la fila de motor 'HOJA.clave'."""
    sh = key.split(".", 1)[0]
    r = ROWS[key]
    return f"{pref(sh)}{mc(m)}${r}"


def RG(key):
    """Rango completo (meses 0..N) de una fila de motor."""
    sh = key.split(".", 1)[0]
    r = ROWS[key]
    return f"{pref(sh)}${MC0}${r}:${MCN}${r}"


def RGC(key, m1, m2):
    sh = key.split(".", 1)[0]
    r = ROWS[key]
    return f"{pref(sh)}{mc(m1)}${r}:{mc(m2)}${r}"


def setc(ws, row, col, value, font=None, fill=None, fmt=None, align=None, border=None, comment=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font = font or F_BASE
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    if align:
        c.alignment = align
    if border:
        c.border = border
    if comment:
        c.comment = Comment(comment, "Modelo CIE-PERG")
        c.comment.width = 320
        c.comment.height = 140
    return c


def title(ws, text, sub=None):
    setc(ws, 1, 1, text, font=F_TITLE)
    if sub:
        setc(ws, 2, 1, sub, font=F_SUB)


def section(ws, row, text, ncols=17):
    for col in range(1, ncols + 1):
        ws.cell(row=row, column=col).fill = FILL_SEC
    setc(ws, row, 1, text, font=F_SEC, fill=FILL_SEC)


def header(ws, row, labels, start_col=1, widths=None):
    for i, lab in enumerate(labels):
        setc(ws, row, start_col + i, lab, font=F_HDR, fill=FILL_HDR, align=CENTER)


def dv_list(ws, cells, options, prompt=None):
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True)
    if prompt:
        dv.prompt = prompt
        dv.showInputMessage = True
    ws.add_data_validation(dv)
    dv.add(cells)
    return dv


def dv_range(ws, cells, rng):
    dv = DataValidation(type="list", formula1=rng, allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(cells)
    return dv


# ===========================================================================
# 1) BASE_PDF: datos publicados (sin alterar) + reconstrucción + controles
# ===========================================================================
BP = {}          # clave -> fila en Base_PDF
BPY0 = 5         # columna E = año 0 ; F..O = años 1..10


def bpc(key, y):
    """Celda de Base_PDF para el año y (0..10). Para escalares usar y=0 (col E)."""
    return f"{pref('BP')}${CL(BPY0 + y)}${BP[key]}"


def bpr(key):
    """Rango años 1..10 de una fila de Base_PDF."""
    return f"{pref('BP')}$F${BP[key]}:$O${BP[key]}"


def build_base_pdf():
    ws = WS["BP"]
    CUR[0] = "BP"
    title(ws, "Base_PDF — Plan publicado «PDF publicado - Oct-2025» (referencia histórica inalterada)",
          "Datos en azul = transcritos del PDF (páginas internas; página física = interna + 2). "
          "Filas «Recalc.» y «Dif.» = comprobaciones. ESTA BASE ES UN PRESUPUESTO HISTÓRICO, NO UN REGISTRO DE EJECUCIÓN REAL.")
    widths = {1: 10, 2: 46, 3: 11, 4: 9}
    for k, v in widths.items():
        ws.column_dimensions[CL(k)].width = v
    for c in range(5, 16):
        ws.column_dimensions[CL(c)].width = 12.5
    ws.column_dimensions["P"].width = 14
    ws.column_dimensions["Q"].width = 70
    r = 4
    header(ws, r, ["Cód.", "Concepto", "Unidad", "Pág.", "Año 0 / Valor"] +
           [f"Año {y}" for y in range(1, 11)] + ["Total / Ref.", "Observación / criterio"])
    ws.freeze_panes = "E5"
    r = 6

    def row_vals(key, code, label, unit, page, vals, fmt=NF_USD, kind="pub", note=None, total=False, y0=None):
        """vals: lista de 10 valores (años 1..10) o fórmulas (función y->str)."""
        nonlocal r
        BP[key] = r
        font = F_IN if kind == "pub" else F_BASE
        setc(ws, r, 1, code)
        setc(ws, r, 2, label, font=F_BOLD if kind == "dif" else F_BASE)
        setc(ws, r, 3, unit)
        setc(ws, r, 4, page)
        if y0 is not None:
            v = y0(0) if callable(y0) else y0
            setc(ws, r, BPY0, v, font=font, fmt=fmt)
        for y in range(1, 11):
            if vals is None:
                continue
            v = vals(y) if callable(vals) else vals[y - 1]
            if v is None:
                continue
            setc(ws, r, BPY0 + y, v, font=font, fmt=fmt)
        if total:
            setc(ws, r, 16, f"=SUM(F{r}:O{r})", fmt=fmt)
        if note:
            setc(ws, r, 17, note, align=Alignment(wrap_text=False))
        if kind == "dif":
            for col in range(5, 17):
                ws.cell(row=r, column=col).fill = FILL_TOT
        r += 1
        return r - 1

    def scalar(key, code, label, unit, page, value, fmt=NF_USD, kind="pub", note=None):
        nonlocal r
        BP[key] = r
        setc(ws, r, 1, code)
        setc(ws, r, 2, label)
        setc(ws, r, 3, unit)
        setc(ws, r, 4, page)
        setc(ws, r, BPY0, value, font=F_IN if kind == "pub" else F_BASE, fmt=fmt)
        if note:
            setc(ws, r, 17, note)
        r += 1
        return r - 1

    def col(y):
        return CL(BPY0 + y)

    # ------------------------------------------------------------------ S1
    section(ws, r, "S1. Parámetros publicados (valor en columna E)"); r += 1
    scalar("tc_ref", "P01", "Tipo de cambio de referencia declarado", "Gs/USD", "4-5", 7300, NF_INT)
    scalar("gg_gs", "P02", "Salario Gerente General (nómina)", "Gs/mes", "30", 18000000, NF_GS)
    scalar("gg_usd", "P03", "Salario Gerente General convertido en la nómina", "USD/mes", "30", 2250, NF_USD)
    scalar("tc_nom", "P04", "TC implícito en la nómina = P02 / P03", "Gs/USD", "30-34",
           f"=E{BP['gg_gs']}/E{BP['gg_usd']}", NF_INT, kind="calc",
           note="Cálculo inferido: la nómina convierte a Gs 8.000/USD, no a Gs 7.300 declarados.")
    scalar("ke", "P05", "Costo de capital propio (Ke) histórico", "% anual", "9, 28", 0.1632, NF_PCT2)
    scalar("kd", "P06", "Tasa préstamo bancario (nominal anual)", "% anual", "4", 0.09, NF_PCT2)
    scalar("w_pub", "P07", "Tasa ponderada publicada (redondeada)", "% anual", "9, 28", 0.0949, NF_PCT2)
    scalar("aporte_ini", "P08", "Aporte propio inicial", "USD", "5, 11", 200000)
    scalar("prestamo", "P09", "Préstamo bancario", "USD", "4, 11", 2791587)
    scalar("plazo", "P10", "Plazo total préstamo", "meses", "4", 60, NF_INT)
    scalar("gracia", "P11", "Gracia (solo intereses)", "meses", "4", 6, NF_INT)
    scalar("cuotas", "P12", "Cuotas francesas", "meses", "4", 54, NF_INT)
    scalar("iva_int", "P13", "IVA sobre intereses", "%", "4, 22", 0.10, NF_PCT)
    scalar("prest", "P14", "Prestaciones sociales (coeficiente del plan)", "% s/ remun.", "19", 0.4309, NF_PCT2,
           note="Coeficiente del plan, no tasa legal validada. Ratio implícito en cuadro 10a ≈ 43,087% (diferencia de redondeo).")
    scalar("m2", "P15", "Superficie nave alquilada", "m2", "6", 3865, NF_INT)
    scalar("alq_m2", "P16", "Alquiler", "USD/m2/mes", "6", 4, NF_USD2)
    scalar("iva_alq", "P17", "IVA sobre alquiler", "%", "6", 0.10, NF_PCT)
    scalar("exp_m2", "P18", "Expensas", "USD/m2/mes", "19", 0.20, NF_USD2)
    scalar("desp", "P19", "Despachos de importación", "%", "19", 0.03, NF_PCT)
    scalar("pkg", "P20", "Packaging", "%", "19", 0.01, NF_PCT)
    scalar("imp", "P21", "Imprevistos (producción y administración)", "%", "19", 0.05, NF_PCT)
    scalar("trib", "P22", "Tributo maquila (s/ valor facturado y exportado)", "%", "7, 19", 0.01, NF_PCT)
    scalar("d_caja", "P23", "Caja y bancos", "días", "24", 3, NF_INT)
    scalar("d_sue", "P24", "Sueldos y anticipos", "días", "8, 24", 15, NF_INT)
    scalar("d_stock", "P25", "Stock de materiales", "días", "8, 24", 120, NF_INT)
    scalar("d_cxc", "P26", "Financiación al cliente (ventas a crédito)", "días", "8, 24", 120, NF_INT)
    scalar("d_pt", "P27", "Productos terminados", "días", "8, 24", 10, NF_INT)
    scalar("d_pmp", "P28", "Proveedores de materia prima", "días", "24", 0, NF_INT)
    scalar("d_pmen", "P29", "Proveedores menores", "días", "8, 24", 30, NF_INT)
    scalar("precio_g", "P30", "Precio gabinete año 1", "USD/u", "16", 270, NF_USD2)
    scalar("precio_c", "P31", "Precio caja año 1", "USD/u", "16", 23, NF_USD2)
    scalar("g_precio", "P32", "Incremento anual de precios desde año 2", "%", "2", 0.01, NF_PCT)
    scalar("mp_g", "P33", "Materia prima gabinete año 1", "USD/u", "17", 82.65, NF_USD2)
    scalar("mp_c", "P34", "Materia prima caja año 1", "USD/u", "17", 7.41, NF_USD2)
    scalar("g_mp", "P35", "Incremento anual materia prima desde año 2 (texto)", "%", "2", 0.01, NF_PCT)
    scalar("ins_g", "P36", "Insumos gabinete", "USD/u", "18", 45, NF_USD2)
    scalar("ins_c_tot", "P37", "Insumos cajas año 1 (total)", "USD", "18", 339570)
    scalar("ins_c", "P38", "Insumos caja = P37 / 70.000 (cálculo)", "USD/u", "18",
           f"=E{BP['ins_c_tot']}/F{r + 200}", NF_USD4, kind="calc",
           note="Se muestra 4,85 en el PDF; la precisión surge del total publicado.")
    scalar("con_g_tot", "P39", "Consumibles gabinetes año 1 (total)", "USD", "18", 56512)
    scalar("con_g", "P40", "Consumibles gabinete = P39 / 12.000 (inferido)", "USD/u", "18",
           None, NF_USD4, kind="calc",
           note="Precisión inferida de cifras redondeadas; no es el dato exacto del archivo fuente.")
    scalar("con_c", "P41", "Consumibles caja", "USD/u", "18", 0.65, NF_USD2)
    scalar("kg_g", "P42", "Peso gabinete", "kg/u", "1", 153, NF_INT)
    scalar("kg_c", "P43", "Peso caja", "kg/u", "1", 13, NF_INT)
    scalar("inv_fija", "P44", "Inversión fija inicial", "USD", "5, 11", 655044)
    scalar("ct_ini", "P45", "Capital operativo inicial", "USD", "5, 11", 2336543)
    scalar("inv_total", "P46", "Inversión inicial total", "USD", "5, 11", 2991587)
    scalar("capex2", "P47", "Inversión fija adicional año 2 (cuadros 18-19)", "USD", "26-27", 100000,
           note="Naturaleza y vida útil no desagregadas en el PDF.")
    scalar("aporte_a1", "P48", "Aporte adicional socios año 1", "USD", "26", 567782)
    scalar("aporte_a2", "P49", "Aporte adicional socios año 2", "USD", "26", 371928)
    scalar("rec_ct", "P50", "Recuperación de capital operativo año 10 (cuadro 19)", "USD", "27", 3974293)
    scalar("vr_af", "P51", "Valor residual inversión año 10 (cuadro 19)", "USD", "27", 59417)
    scalar("van_pub", "P52", "VAN publicado (al 9,49%)", "USD", "27, 35", 4588736)
    scalar("tir_pub", "P53", "TIR publicada «con financiamiento»", "%", "27, 35", 0.2261, NF_PCT2)
    scalar("pb_pub", "P54", "Payback publicado («durante el 7mo año»)", "años", "28", 7, NF_INT)
    scalar("alim_gs", "P55", "Alimentación inferida por persona", "Gs/mes", "19",
           576000, NF_GS, kind="calc",
           note="Inferido: USD 40.715/43 = USD 946,85/persona/año = Gs 576.000/mes a Gs 7.300 (p.ej. 24.000 Gs x 24 días). Confirmar.")
    r += 1

    # ------------------------------------------------------------------ S2
    section(ws, r, "S2. Inversión (cuadros 1, 2, 5 y 6)"); r += 1
    header(ws, r, ["Cód.", "Ítem (cuadro 2)", "Cant.", "Pág.", "FOB USD", "Seg/flete", "CIF USD", "CIF publ.", "Dif."]); r += 1
    maq = [("M01", "Conjunto línea de pintura", "Conj.", 161583, 177741),
           ("M02", "Cortadoras láser", 2, 74400, 81840),
           ("M03", "Manipulador", 1, 9000, 9900),
           ("M04", "Dobladoras", 4, 134128, 147541),
           ("M05", "Soldadoras", 6, 3000, 3300),
           ("M06", "Lijadoras", 4, 520, 572),
           ("M07", "Conjunto baño", "Conj.", 25000, 27500),
           ("M08", "Línea de montaje", 1, 8700, 9570),
           ("M09", "Montacargas", 1, 26000, 28600),
           ("M10", "Soldadoras capacitivas", 2, 5300, 5830)]
    BP["maq_first"] = r
    for code, name, cant, fob, cifp in maq:
        BP["maq_" + code] = r
        setc(ws, r, 1, code); setc(ws, r, 2, name); setc(ws, r, 3, cant, font=F_IN); setc(ws, r, 4, "12")
        setc(ws, r, 5, fob, font=F_IN, fmt=NF_USD2)
        setc(ws, r, 6, f"=E{r}*10%", fmt=NF_USD2)
        setc(ws, r, 7, f"=E{r}+F{r}", fmt=NF_USD2)
        setc(ws, r, 8, cifp, font=F_IN, fmt=NF_USD)
        setc(ws, r, 9, f"=G{r}-H{r}", fmt=NF_USD2)
        setc(ws, r, 17, "Valor FOB = total del renglón (no multiplicar por cantidad).")
        r += 1
    BP["maq_last"] = r - 1
    fr, lr = BP["maq_first"], BP["maq_last"]
    BP["maq_tot"] = r
    setc(ws, r, 2, "Subtotal maquinaria", font=F_BOLD)
    setc(ws, r, 5, f"=SUM(E{fr}:E{lr})", fmt=NF_USD2, font=F_BOLD)
    setc(ws, r, 6, f"=SUM(F{fr}:F{lr})", fmt=NF_USD2, font=F_BOLD)
    setc(ws, r, 7, f"=SUM(G{fr}:G{lr})", fmt=NF_USD2, font=F_BOLD)
    setc(ws, r, 8, 492394, font=F_IN, fmt=NF_USD)
    setc(ws, r, 9, f"=G{r}-H{r}", fmt=NF_USD2)
    setc(ws, r, 17, "Publicado: FOB 447.631; seguro/flete 44.763,10; CIF 492.394 (redondeado).")
    r += 1
    BP["desp_maq"] = r
    setc(ws, r, 2, "Gastos de despacho maquinaria (publicado)"); setc(ws, r, 4, "12")
    setc(ws, r, 5, 14772, font=F_IN, fmt=NF_USD)
    setc(ws, r, 6, f"=E{r}/G{BP['maq_tot']}", fmt=NF_PCT4)
    setc(ws, r, 17, "Col. F = despacho / CIF (≈3%). Se usa esa proporción exacta para asignar el despacho por ítem en Inversion_Activos.")
    r += 1
    scalar("inf", "K01", "Adecuación de infraestructura (cuadro 5)", "USD", "13", 87000)
    scalar("sof", "K02", "Software de gestión (cuadro 6)", "USD", "13", 43478)
    scalar("mob", "K03", "Mobiliario (cuadro 6)", "USD", "13", 17400)
    scalar("inv_fija_calc", "K04", "Inversión fija recalculada = infra + CIF + despacho + software + mobiliario", "USD", "11",
           f"=E{BP['inf']}+G{BP['maq_tot']}+E{BP['desp_maq']}+E{BP['sof']}+E{BP['mob']}", kind="calc")
    scalar("inv_fija_dif", "K05", "Dif. vs inversión fija publicada", "USD", "11",
           f"=E{BP['inv_fija_calc']}-E{BP['inv_fija']}", NF_USD2, kind="calc",
           note="Diferencia por redondeo del CIF (0,10).")
    scalar("inv_tot_calc", "K06", "Inversión inicial = fija + capital operativo", "USD", "11",
           f"=E{BP['inv_fija']}+E{BP['ct_ini']}", kind="calc")
    scalar("fuentes_calc", "K07", "Fuentes iniciales = aporte + préstamo", "USD", "11",
           f"=E{BP['aporte_ini']}+E{BP['prestamo']}", kind="calc")
    r += 1

    # ------------------------------------------------------------------ S3
    section(ws, r, "S3. Producción, capacidad y ventas (cuadros 7, 7a y 8)"); r += 1
    row_vals("cap_g", "Q01", "Gabinetes - capacidad instalada", "u/año", "14", [15000, 15000] + [22000] * 8, NF_INT)
    row_vals("prod_g", "Q02", "Gabinetes - producción", "u/año", "14", [12000, 15000, 18000] + [21000] * 7, NF_INT)
    row_vals("vta_g", "Q03", "Gabinetes - venta", "u/año", "14", [12000, 15000, 18000] + [21000] * 7, NF_INT)
    row_vals("cap_c", "Q04", "Cajas - capacidad instalada", "u/año", "14", [90000, 90000] + [121000] * 8, NF_INT)
    row_vals("prod_c", "Q05", "Cajas - producción", "u/año", "14", [70000, 90000, 110000] + [120000] * 7, NF_INT)
    row_vals("vta_c", "Q06", "Cajas - venta", "u/año", "14", [70000, 90000, 110000] + [120000] * 7, NF_INT)
    row_vals("util_g", "Q07", "Recalc. utilización gabinetes = producción / capacidad", "%", "15",
             lambda y: f"={col(y)}{BP['prod_g']}/{col(y)}{BP['cap_g']}", NF_PCT, kind="calc")
    row_vals("util_c", "Q08", "Recalc. utilización cajas", "%", "15",
             lambda y: f"={col(y)}{BP['prod_c']}/{col(y)}{BP['cap_c']}", NF_PCT, kind="calc")
    row_vals("capinc_g", "Q09", "Aumento de capacidad gabinetes vs año 2", "%", "3",
             lambda y: f"={col(y)}{BP['cap_g']}/$F{BP['cap_g']}-1", NF_PCT, kind="calc",
             note="El texto (p. 3) menciona +10%; la tabla sube 46,7% (gabinetes) y 34,4% (cajas) desde el año 3.")
    row_vals("capinc_c", "Q10", "Aumento de capacidad cajas vs año 2", "%", "3",
             lambda y: f"={col(y)}{BP['cap_c']}/$F{BP['cap_c']}-1", NF_PCT, kind="calc")
    row_vals("precio_g_pub", "Q11", "Precio gabinete publicado (2 decimales)", "USD/u", "16",
             [270.00, 272.70, 275.43, 278.18, 280.96, 283.77, 286.61, 289.48, 292.37, 295.30], NF_USD2)
    row_vals("precio_c_pub", "Q12", "Precio caja publicado (2 decimales)", "USD/u", "16",
             [23.00, 23.23, 23.46, 23.70, 23.93, 24.17, 24.41, 24.66, 24.91, 25.15], NF_USD2)
    row_vals("precio_g_calc", "Q13", "Recalc. precio gabinete = P30 x (1+P32)^(año-1)", "USD/u", "2",
             lambda y: f"=$E${BP['precio_g']}*(1+$E${BP['g_precio']})^({y}-1)", NF_USD4, kind="calc")
    row_vals("precio_c_calc", "Q14", "Recalc. precio caja", "USD/u", "2",
             lambda y: f"=$E${BP['precio_c']}*(1+$E${BP['g_precio']})^({y}-1)", NF_USD4, kind="calc")
    row_vals("ing_g_pub", "Q15", "Ventas gabinetes publicadas", "USD", "16",
             [3240000, 4090500, 4957686, 5841807, 5900225, 5959227, 6018819, 6079007, 6139798, 6201195], total=True)
    row_vals("ing_c_pub", "Q16", "Ventas cajas publicadas", "USD", "16",
             [1610000, 2090700, 2580853, 2843631, 2872067, 2900788, 2929796, 2959094, 2988685, 3018571], total=True)
    row_vals("ing_pub", "Q17", "Ingresos totales publicados", "USD", "16, 23",
             [4850000, 6181200, 7538539, 8685437, 8772292, 8860015, 8948615, 9038101, 9128482, 9219767], total=True)
    row_vals("ing_g_calc", "Q18", "Recalc. ventas gabinetes (precio sin redondear)", "USD", "16",
             lambda y: f"={col(y)}{BP['vta_g']}*{col(y)}{BP['precio_g_calc']}", kind="calc", total=True)
    row_vals("ing_c_calc", "Q19", "Recalc. ventas cajas", "USD", "16",
             lambda y: f"={col(y)}{BP['vta_c']}*{col(y)}{BP['precio_c_calc']}", kind="calc", total=True)
    row_vals("ing_calc", "Q20", "Recalc. ingresos totales", "USD", "16",
             lambda y: f"={col(y)}{BP['ing_g_calc']}+{col(y)}{BP['ing_c_calc']}", kind="calc", total=True)
    row_vals("ing_dif", "Q21", "Dif. ingresos recalculados - publicados", "USD", "",
             lambda y: f"={col(y)}{BP['ing_calc']}-{col(y)}{BP['ing_pub']}", NF_USD2, kind="dif", total=True)
    row_vals("ing_sum_dif", "Q22", "Dif. suma productos publicados - total publicado", "USD", "",
             lambda y: f"={col(y)}{BP['ing_g_pub']}+{col(y)}{BP['ing_c_pub']}-{col(y)}{BP['ing_pub']}", NF_USD2, kind="dif")
    r += 1
    # fix de referencias diferidas (P38 y P40 dependían de filas de volumen)
    ws.cell(row=BP["ins_c"], column=BPY0).value = f"=E{BP['ins_c_tot']}/F{BP['vta_c']}"
    ws.cell(row=BP["con_g"], column=BPY0).value = f"=E{BP['con_g_tot']}/F{BP['vta_g']}"

    # ------------------------------------------------------------------ S4
    section(ws, r, "S4. Materia prima e insumos (cuadros 9a y 9b)"); r += 1
    row_vals("mp_g_pub", "M01", "MP gabinetes publicada", "USD", "17",
             [991800, 1252148, 1517603, 1788242, 1806124, 1824186, 1842427, 1860852, 1879460, 1898255], total=True)
    row_vals("mp_c_pub", "M02", "MP cajas publicada", "USD", "17",
             [518700, 673569, 823251] + [898092] * 7, total=True)
    row_vals("mpu_c_pub", "M03", "MP caja por pieza publicada", "USD/u", "17",
             [7.41] + [7.48] * 9, NF_USD2)
    row_vals("mp_g_calc", "M04", "Recalc. MP gabinetes = u x 82,65 x 1,01^(año-1)", "USD", "2, 17",
             lambda y: f"={col(y)}{BP['prod_g']}*$E${BP['mp_g']}*(1+$E${BP['g_mp']})^({y}-1)", kind="calc", total=True)
    row_vals("mp_c_calc_pdf", "M05", "Recalc. MP cajas criterio PDF (1% solo una vez)", "USD", "17",
             lambda y: f"={col(y)}{BP['prod_c']}*$E${BP['mp_c']}*IF({y}=1,1,1+$E${BP['g_mp']})", kind="calc", total=True)
    row_vals("mp_c_calc_acum", "M06", "Recalc. MP cajas con 1% acumulativo (texto p. 2)", "USD", "2",
             lambda y: f"={col(y)}{BP['prod_c']}*$E${BP['mp_c']}*(1+$E${BP['g_mp']})^({y}-1)", kind="calc", total=True)
    row_vals("mp_g_dif", "M07", "Dif. MP gabinetes recalc. - publicada", "USD", "",
             lambda y: f"={col(y)}{BP['mp_g_calc']}-{col(y)}{BP['mp_g_pub']}", NF_USD2, kind="dif")
    row_vals("mp_c_dif", "M08", "Dif. MP cajas criterio PDF - publicada", "USD", "",
             lambda y: f"={col(y)}{BP['mp_c_calc_pdf']}-{col(y)}{BP['mp_c_pub']}", NF_USD2, kind="dif")
    row_vals("mp_c_dif_acum", "M09", "Efecto corrección 1% acumulativo cajas (acum. - publicada)", "USD", "",
             lambda y: f"={col(y)}{BP['mp_c_calc_acum']}-{col(y)}{BP['mp_c_pub']}", NF_USD2, kind="dif", total=True,
             note="Año 10 ≈ USD 74.413 antes de efectos asociados (despacho, packaging, imprevistos, CT).")
    row_vals("ins_g_pub", "M10", "Insumos gabinetes publicados", "USD", "18",
             [540000, 675000, 810000] + [945000] * 7)
    row_vals("ins_c_pub", "M11", "Insumos cajas publicados", "USD", "18",
             [339570, 436590, 533610] + [582120] * 7)
    row_vals("con_g_pub", "M12", "Consumibles gabinetes publicados", "USD", "18",
             [56512, 70640, 84767] + [98895] * 7)
    row_vals("con_c_pub", "M13", "Consumibles cajas publicados", "USD", "18",
             [45500, 58500, 71500] + [78000] * 7)
    row_vals("ins_dif", "M14", "Dif. insumos+consumibles recalc. (u x costo unitario) - publicados", "USD", "",
             lambda y: (f"={col(y)}{BP['prod_g']}*($E${BP['ins_g']}+$E${BP['con_g']})+{col(y)}{BP['prod_c']}*($E${BP['ins_c']}+$E${BP['con_c']})"
                        f"-({col(y)}{BP['ins_g_pub']}+{col(y)}{BP['ins_c_pub']}+{col(y)}{BP['con_g_pub']}+{col(y)}{BP['con_c_pub']})"),
             NF_USD2, kind="dif", note="Diferencias por redondeo del consumible inferido de gabinetes.")
    r += 1

    # ------------------------------------------------------------------ S5 nómina
    section(ws, r, "S5. Nómina publicada (pp. 30-34). Salario mensual en Gs; dotación años 1-5 (6-10 = año 5)"); r += 1
    header(ws, r, ["Cód.", "Puesto", "Clase PDF", "Pág.", "Salario Gs/mes", "Dot. año 1", "Dot. año 2", "Dot. año 3",
                   "Dot. año 4", "Dot. año 5", "USD/mes publ.", "USD año 1 publ."]); r += 1
    nomina = [
        ("D01", "Operador láser", "Directo", 4668900, [4, 4, 5, 5, 6], 584, 28013),
        ("D02", "Doblador («Dobrador» en el PDF)", "Directo", 3650655, [8, 8, 9, 9, 10], 456, 43808),
        ("D03", "Soldador («Soladador» en el PDF)", "Directo", 3968900, [5, 5, 6, 6, 6], 496, 29767),
        ("D04", "Baño", "Directo", 3650655, [2, 2, 2, 2, 3], 456, 10952),
        ("D05", "Pintor", "Directo", 3650655, [4, 4, 4, 5, 5], 456, 21904),
        ("D06", "Montador", "Directo", 3150655, [10, 10, 10, 11, 11], 394, 47260),
        ("D07", "Ayudante", "Directo", 2798309, [10, 10, 10, 11, 11], 350, 41975),
        ("I01", "Gerente General", "Indirecto", 18000000, [1] * 5, 2250, 27000),
        ("I02", "Gerente Industrial", "Indirecto", 9800000, [1] * 5, 1225, 14700),
        ("I03", "Coord. Prod. Corte y Deformación", "Indirecto", 6550000, [1] * 5, 819, 9825),
        ("I04", "Coord. Prod. Soldadura", "Indirecto", 6550000, [1] * 5, 819, 9825),
        ("I05", "Coord. Prod. Trat. Sup. y Pintura", "Indirecto", 6550000, [1] * 5, 819, 9825),
        ("I06", "Coord. Prod. Montaje", "Indirecto", 6550000, [1] * 5, 819, 9825),
        ("I07", "Ingeniería", "Indirecto", 6550000, [3] * 5, 819, 29475),
        ("I08", "Mantenimiento", "Indirecto", 3650655, [2] * 5, 456, 10952),
        ("I09", "PCP", "Indirecto", 9800000, [1] * 5, 1225, 14700),
        ("I10", "Analista PCP", "Indirecto", 2798309, [1] * 5, 350, 4197),
        ("I11", "Depósito", "Indirecto", 3150655, [1] * 5, 394, 4726),
        ("I12", "Compras", "Indirecto", 6550000, [1] * 5, 819, 9825),
        ("I13", "Logística/Comex", "Indirecto", 6550000, [1] * 5, 819, 9825),
        ("I14", "Coord. Control Calidad", "Indirecto", 6550000, [1] * 5, 819, 9825),
        ("I15", "Asistente Control Calidad", "Indirecto", 2798309, [1] * 5, 350, 4197),
        ("I16", "Coord. Procesos", "Indirecto", 6550000, [1] * 5, 819, 9825),
        ("I17", "Asistente Procesos", "Indirecto", 2798309, [1] * 5, 350, 4197),
        ("I18", "Administrador", "Indirecto", 9800000, [1] * 5, 1225, 14700),
        ("I19", "Asistente Administrativo", "Indirecto", 2798309, [1] * 5, 350, 4197),
    ]
    BP["nom_first"] = r
    for code, name, cls, gs, dots, usdm, usda in nomina:
        setc(ws, r, 1, code); setc(ws, r, 2, name); setc(ws, r, 3, cls, font=F_IN); setc(ws, r, 4, "30-34")
        setc(ws, r, 5, gs, font=F_IN, fmt=NF_GS)
        for i, d in enumerate(dots):
            setc(ws, r, 6 + i, d, font=F_IN, fmt=NF_INT)
        setc(ws, r, 11, usdm, font=F_IN, fmt=NF_USD)
        setc(ws, r, 12, usda, font=F_IN, fmt=NF_USD)
        r += 1
    BP["nom_last"] = r - 1
    BP["nom_data"] = nomina
    nf, nl = BP["nom_first"], BP["nom_last"]
    # dotación y masa por año (años 6-10 = año 5)
    def dotcol(y):
        return CL(5 + min(y, 5))
    row_vals("dot_dir", "N01", "Dotación directa", "personas", "30-34",
             lambda y: f'=SUMIF($C${nf}:$C${nl},"Directo",${dotcol(y)}${nf}:${dotcol(y)}${nl})', NF_INT, kind="calc")
    row_vals("dot_ind", "N02", "Dotación indirecta", "personas", "30-34",
             lambda y: f'=SUMIF($C${nf}:$C${nl},"Indirecto",${dotcol(y)}${nf}:${dotcol(y)}${nl})', NF_INT, kind="calc")
    row_vals("masa_dir_gs", "N03", "Remuneración base anual directos", "Gs", "30-34",
             lambda y: f'=SUMPRODUCT(($C${nf}:$C${nl}="Directo")*$E${nf}:$E${nl}*${dotcol(y)}${nf}:${dotcol(y)}${nl})*12', NF_GS, kind="calc")
    row_vals("masa_ind_gs", "N04", "Remuneración base anual indirectos", "Gs", "30-34",
             lambda y: f'=SUMPRODUCT(($C${nf}:$C${nl}="Indirecto")*$E${nf}:$E${nl}*${dotcol(y)}${nf}:${dotcol(y)}${nl})*12', NF_GS, kind="calc")
    row_vals("mod_8000", "N05", "Recalc. MO fábrica a TC nómina (8.000)", "USD", "19",
             lambda y: f"={col(y)}{BP['masa_dir_gs']}/$E${BP['tc_nom']}", NF_USD2, kind="calc")
    row_vals("moa_8000", "N06", "Recalc. MO administrativa a TC nómina (8.000)", "USD", "19",
             lambda y: f"={col(y)}{BP['masa_ind_gs']}/$E${BP['tc_nom']}", NF_USD2, kind="calc")
    row_vals("mod_7300", "N07", "Sensibilidad: MO fábrica a TC declarado (7.300)", "USD", "4",
             lambda y: f"={col(y)}{BP['masa_dir_gs']}/$E${BP['tc_ref']}", NF_USD2, kind="calc")
    row_vals("moa_7300", "N08", "Sensibilidad: MO administrativa a TC declarado (7.300)", "USD", "4",
             lambda y: f"={col(y)}{BP['masa_ind_gs']}/$E${BP['tc_ref']}", NF_USD2, kind="calc")
    row_vals("nom_fx_ef", "N09", "Efecto coherencia cambiaria (remuneración + prestaciones)", "USD", "",
             lambda y: (f"=({col(y)}{BP['mod_7300']}+{col(y)}{BP['moa_7300']}-{col(y)}{BP['mod_8000']}-{col(y)}{BP['moa_8000']})"
                        f"*(1+$E${BP['prest']})"), NF_USD2, kind="dif", total=True,
             note="Año 1 ≈ USD 59.730. Sensibilidad por coherencia cambiaria, no dato salarial real de 2026.")
    r += 1

    # ------------------------------------------------------------------ S6 cuadro 10a
    section(ws, r, "S6. Costos anuales (cuadro 10a) — publicado"); r += 1
    c10 = [
        ("MP", "a Materia prima", [1510500, 1925717, 2340854, 2686334, 2704216, 2722278, 2740519, 2758944, 2777552, 2796347]),
        ("INS", "b Insumos", [879570, 1111590, 1343610] + [1527120] * 7),
        ("CON", "c Insumos consumibles", [102012, 129140, 156267] + [176895] * 7),
        ("MOD", "d Mano de obra fábrica", [223678, 223678, 242111, 256510] + [274466] * 6),
        ("PRD", "e Prestaciones sociales 43,09%", [96376, 96376, 104318, 110522] + [118258] * 6),
        ("ALD", "h Alimentación de personal", [40715, 40715, 43555, 46396] + [49236] * 6),
        ("FLE", "i Flete marítimo + terrestre MP", [424000, 535200, 646400] + [736800] * 7),
        ("ENE", "j Energía eléctrica planta", [23014] * 10),
        ("MAN", "k Mantenimiento y reparaciones", [29708, 30302, 30897, 31491, 32085, 32679, 33273, 33867, 34462, 35056]),
        ("DES", "l Despachos de importación 3%", [71702, 91119, 110534, 126404, 126940, 127482, 128029, 128582, 129140, 129704]),
        ("SEG", "m Seguros caución maquila", [4850, 6181, 7539, 8685, 8772, 8860, 8949, 9038, 9128, 9220]),
        ("PKG", "o Packaging 1%", [32769, 40624, 48771, 55406, 55870, 56051, 56233, 56417, 56603, 56791]),
        ("IMP", "q Imprevistos 5%", [171945, 212683, 254893, 289279, 291684, 292657, 293640, 294632, 295634, 296645]),
        ("TRI", "r Tasas y tributos - maquila 1%", [48500, 61812, 75385, 86854, 87723, 88600, 89486, 90381, 91285, 92198]),
        ("PROD", "1 TOTAL COSTOS DE PRODUCCIÓN", [3659338, 4528150, 5428148, 6161710, 6213080, 6234396, 6255919, 6277651, 6299594, 6321750]),
        ("MOA", "a Mano de obra administrativa", [211643] * 10),
        ("PRA", "b Prestaciones sociales 43,09%", [91190] * 10),
        ("ALA", "c Alimentación de personal", [20831] * 10),
        ("SRV", "d Agua, comunicaciones y electricidad", [1233] * 10),
        ("ALQ", "e Alquiler nave industrial (USD 4,4/m2 con IVA)", [204072, 210194, 216500, 222995, 229685, 236575, 243673, 250983, 258512, 266268]),
        ("MAA", "f Mantenimiento y reparaciones", [4932] * 10),
        ("SGA", "g Seguros sobre activos fijos", [8219] * 10),
        ("GG", "h Gastos generales", [22060] * 10),
        ("HON", "j Honorarios y otros fijos", [24932] * 10),
        ("EXP", "l Expensas USD 0,2/m2", [9276, 9554, 9841, 10136, 10440, 10753, 11076, 11408, 11751, 12103]),
        ("IMA", "k Imprevistos 5%", [19252] * 10),
        ("ADM", "2 TOTAL GASTOS DE ADMINISTRACIÓN", [617639, 624039, 630632, 637422, 644416, 651620, 659039, 666682, 674554, 682662]),
        ("GV", "3 Gastos de ventas", [0] * 10),
        ("DEP", "4 Depreciación (gastos no desembolsados)", [0] + [71592] * 5 + [59417] * 4),
        ("FIN", "5 Gastos financieros (incluye IVA)", [271102, 226274, 169416, 107224, 39198] + [0] * 5),
        ("TOT", "TOTAL COSTOS ANUALES", [4548078, 5450056, 6299788, 6977948, 6968286, 6957607, 6974375, 7003749, 7033564, 7063828]),
    ]
    for code, label, vals in c10:
        row_vals("p_" + code, code, label, "USD", "19", vals, total=True)
    r += 1
    section(ws, r, "S6b. Cuadro 10a — reconstrucción con bases identificadas (cálculo inferido; ver columna Q)"); r += 1

    def P(code, y):
        return f"{col(y)}{BP['p_' + code]}"

    def Rc(code, y):
        return f"{col(y)}{BP['r_' + code]}"

    # flete: solución exacta de 2 ecuaciones (años 1 y 2)
    scalar("fle_b", "F01", "Flete MP por caja inferido (resuelve años 1-2)", "USD/u", "19",
           None, NF_USD4, kind="calc",
           note="b = (F2·G1 - F1·G2)/(C2·G1 - C1·G2). Ajuste exacto también en años 3 y 4. Inferido; confirmar con el origen.")
    scalar("fle_a", "F02", "Flete MP por gabinete inferido", "USD/u", "19", None, NF_USD4, kind="calc")
    fb, fa = BP["fle_b"], BP["fle_a"]
    ws.cell(row=fb, column=BPY0).value = (
        f"=(G{BP['p_FLE']}*F{BP['prod_g']}-F{BP['p_FLE']}*G{BP['prod_g']})/(G{BP['prod_c']}*F{BP['prod_g']}-F{BP['prod_c']}*G{BP['prod_g']})")
    ws.cell(row=fa, column=BPY0).value = f"=(F{BP['p_FLE']}-F{BP['prod_c']}*E{fb})/F{BP['prod_g']}"
    scalar("ene_gs", "F03", "Energía planta inferida en Gs", "Gs/año", "19", 168000000, NF_GS, kind="calc",
           note="USD 23.014 x 7.300 ≈ Gs 168.000.000 (monto redondo). Inferencia de moneda de origen; confirmar.")
    scalar("srv_gs", "F04", "Agua, comunicaciones y electricidad inferida en Gs", "Gs/año", "19", 9000000, NF_GS, kind="calc")
    scalar("maa_gs", "F05", "Mantenimiento administración inferido en Gs", "Gs/año", "19", 36000000, NF_GS, kind="calc")
    scalar("sga_gs", "F06", "Seguros activos fijos inferidos en Gs", "Gs/año", "19", 60000000, NF_GS, kind="calc")
    scalar("hon_gs", "F07", "Honorarios y otros fijos inferidos en Gs", "Gs/año", "19", 182000000, NF_GS, kind="calc")
    scalar("man_inc", "F08", "Mantenimiento planta: incremento lineal anual (% del año 1)", "%", "19", 0.02, NF_PCT, kind="calc",
           note="Inferido: 29.708 + 594,16 por año (= 2% de 29.708; también = 5% de maquinaria+infraestructura 594.166, +0,1% anual).")
    scalar("segc_pct", "F09", "Seguro de caución: % sobre ventas", "%", "19", 0.001, NF_PCT2, kind="calc",
           note="Inferido exacto: 4.850 / 4.850.000.")
    r += 1
    recon = [
        ("MP", lambda y: f"={col(y)}{BP['mp_g_pub']}+{col(y)}{BP['mp_c_pub']}", "Suma cuadro 9a (gabinetes + cajas)."),
        ("INS", lambda y: f"={col(y)}{BP['ins_g_pub']}+{col(y)}{BP['ins_c_pub']}", "Suma cuadro 9b."),
        ("CON", lambda y: f"={col(y)}{BP['con_g_pub']}+{col(y)}{BP['con_c_pub']}", "Suma cuadro 9b."),
        ("MOD", lambda y: f"={col(y)}{BP['mod_8000']}", "Nómina directa a Gs 8.000 (TC implícito)."),
        ("PRD", lambda y: f"={Rc('MOD', y)}*$E${BP['prest']}", "43,09% x MO fábrica."),
        ("ALD", lambda y: f"={col(y)}{BP['dot_dir']}*$E${BP['alim_gs']}*12/$E${BP['tc_ref']}", "Gs 576.000/persona/mes a Gs 7.300."),
        ("FLE", lambda y: f"={col(y)}{BP['prod_g']}*$E${fa}+{col(y)}{BP['prod_c']}*$E${fb}", "USD/u inferidos (F01-F02)."),
        ("ENE", lambda y: f"=$E${BP['ene_gs']}/$E${BP['tc_ref']}", "Gs 168 M / 7.300."),
        ("MAN", lambda y: f"=$F${BP['p_MAN']}*(1+$E${BP['man_inc']}*({y}-1))", "Año 1 publicado + 2% lineal."),
        ("DES", lambda y: f"=$E${BP['desp']}*({Rc('MP', y)}+{Rc('INS', y)})", "Base identificada: 3% x (MP + insumos)."),
        ("SEG", lambda y: f"=$E${BP['segc_pct']}*{col(y)}{BP['ing_pub']}", "0,1% x ventas."),
        ("PKG", lambda y: (f"=$E${BP['pkg']}*({Rc('MP', y)}+{Rc('INS', y)}+{Rc('CON', y)}+{Rc('FLE', y)}+{Rc('MOD', y)}+{Rc('PRD', y)}+{Rc('ALD', y)})"),
         "Base identificada: 1% x (MP+insumos+consumibles+flete+MO fábrica+prestaciones+alimentación)."),
        ("IMP", lambda y: (f"=$E${BP['imp']}*({Rc('MP', y)}+{Rc('INS', y)}+{Rc('CON', y)}+{Rc('MOD', y)}+{Rc('PRD', y)}+{Rc('ALD', y)}+{Rc('FLE', y)}"
                           f"+{Rc('ENE', y)}+{Rc('MAN', y)}+{Rc('DES', y)}+{Rc('SEG', y)}+{Rc('PKG', y)})"),
         "Base identificada: 5% x costos de producción excepto imprevistos y tributo."),
        ("TRI", lambda y: f"=$E${BP['trib']}*{col(y)}{BP['ing_pub']}", "1% x ventas exportadas."),
        ("PROD", lambda y: "=" + "+".join(Rc(k, y) for k in ["MP", "INS", "CON", "MOD", "PRD", "ALD", "FLE", "ENE", "MAN", "DES", "SEG", "PKG", "IMP", "TRI"]), "Suma."),
        ("MOA", lambda y: f"={col(y)}{BP['moa_8000']}", "Nómina indirecta (22) a Gs 8.000."),
        ("PRA", lambda y: f"={Rc('MOA', y)}*$E${BP['prest']}", "43,09%."),
        ("ALA", lambda y: f"={col(y)}{BP['dot_ind']}*$E${BP['alim_gs']}*12/$E${BP['tc_ref']}", "Gs 576.000/persona/mes."),
        ("SRV", lambda y: f"=$E${BP['srv_gs']}/$E${BP['tc_ref']}", "Gs 9 M / 7.300."),
        ("ALQ", lambda y: f"=$E${BP['m2']}*$E${BP['alq_m2']}*12*(1+$E${BP['iva_alq']})*(1+0.03)^({y}-1)", "3.865 m2 x USD 4 x 12 x 1,10 x 1,03^(año-1)."),
        ("MAA", lambda y: f"=$E${BP['maa_gs']}/$E${BP['tc_ref']}", "Gs 36 M / 7.300."),
        ("SGA", lambda y: f"=$E${BP['sga_gs']}/$E${BP['tc_ref']}", "Gs 60 M / 7.300."),
        ("GG", lambda y: f"=$F${BP['p_GG']}", "Publicado (moneda de origen no identificable)."),
        ("HON", lambda y: f"=$E${BP['hon_gs']}/$E${BP['tc_ref']}", "Gs 182 M / 7.300."),
        ("EXP", lambda y: f"=$E${BP['m2']}*$E${BP['exp_m2']}*12*(1+0.03)^({y}-1)", "3.865 m2 x USD 0,20 x 12 x 1,03^(año-1)."),
        ("IMA", lambda y: f"=$E${BP['imp']}*({Rc('MOA', y)}+{Rc('PRA', y)}+{Rc('ALA', y)}+{Rc('SRV', y)}+{Rc('MAA', y)}+{Rc('SGA', y)}+{Rc('GG', y)}+{Rc('HON', y)})",
         "Base identificada: 5% x administración excepto alquiler y expensas."),
        ("ADM", lambda y: "=" + "+".join(Rc(k, y) for k in ["MOA", "PRA", "ALA", "SRV", "ALQ", "MAA", "SGA", "GG", "HON", "EXP", "IMA"]), "Suma."),
    ]
    for code, fn, note in recon:
        row_vals("r_" + code, code, "Recalc. " + [l for c, l, v in c10 if c == code][0], "USD", "19", fn, NF_USD,
                 kind="calc", total=True, note=note)
    r += 1
    section(ws, r, "S6c. Cuadro 10a — diferencias reconstrucción - publicado (tolerancia: redondeo)"); r += 1
    for code, fn, note in recon:
        row_vals("d_" + code, code, "Dif. " + code, "USD", "",
                 lambda y, code=code: f"={Rc(code, y)}-{P(code, y)}", NF_USD2, kind="dif")
    row_vals("d_absmax", "DMX", "Máxima diferencia absoluta del bloque (control)", "USD", "",
             lambda y: f"=MAX(ABS(MIN({col(y)}{BP['d_MP']}:{col(y)}{BP['d_ADM']})),MAX({col(y)}{BP['d_MP']}:{col(y)}{BP['d_ADM']}))",
             NF_USD2, kind="dif")
    r += 1

    # ------------------------------------------------------------------ S7 cuadro 10b
    section(ws, r, "S7. Costos unitarios por producto (cuadro 10b) y márgenes"); r += 1
    row_vals("cd_g_pub", "U01", "Costo MP e insumos gabinetes (base directa 10b)", "USD", "20",
             [1854532, 2330562, 2811700, 3298022, 3315905, 3333966, 3352208, 3370632, 3389241, 3408035])
    row_vals("cd_c_pub", "U02", "Costo MP e insumos cajas (base directa 10b)", "USD", "20",
             [1035720, 1338309, 1635711] + [1784412] * 7)
    row_vals("mbp_g_pub", "U03", "«Margen bruto %» gabinetes publicado", "%", "20",
             [0.572, 0.570, 0.567, 0.565, 0.562, 0.559, 0.557, 0.554, 0.552, 0.550], NF_PCT)
    row_vals("mbp_c_pub", "U04", "«Margen bruto %» cajas publicado", "%", "20",
             [0.643, 0.640, 0.634, 0.628, 0.621, 0.615, 0.609, 0.603, 0.597, 0.591], NF_PCT)
    row_vals("cdp_g", "U05", "Recalc. costo directo / precio gabinetes", "%", "20",
             lambda y: f"=({col(y)}{BP['cd_g_pub']}/{col(y)}{BP['vta_g']})/{col(y)}{BP['precio_g_calc']}", NF_PCT, kind="calc",
             note="Coincide con el «margen» publicado: el PDF rotula como margen el cociente costo/precio.")
    row_vals("mg_g", "U06", "Margen sobre costo parcial incluido gabinetes = (precio - costo)/precio", "%", "20",
             lambda y: f"=1-{col(y)}{BP['cdp_g']}", NF_PCT, kind="calc")
    row_vals("cdp_c", "U07", "Recalc. costo directo / precio cajas", "%", "20",
             lambda y: f"=({col(y)}{BP['cd_c_pub']}/{col(y)}{BP['vta_c']})/{col(y)}{BP['precio_c_calc']}", NF_PCT, kind="calc")
    row_vals("mg_c", "U08", "Margen sobre costo parcial incluido cajas", "%", "20",
             lambda y: f"=1-{col(y)}{BP['cdp_c']}", NF_PCT, kind="calc")
    row_vals("base10b_dif", "U09", "Dif. base directa 10b vs MP+insumos+consumibles+flete (10a)", "USD", "19-21",
             lambda y: f"={col(y)}{BP['cd_g_pub']}+{col(y)}{BP['cd_c_pub']}-({P('MP', y)}+{P('INS', y)}+{P('CON', y)}+{P('FLE', y)})",
             NF_USD, kind="dif", note="Año 1: -25.830. El 10b asigna flete a USD 0,145/kg; el 10a implica 23,20 y 2,08 USD/u.")
    row_vals("fle10b_kg", "U10", "Flete implícito 10b por kg (gabinetes)", "USD/kg", "20",
             lambda y: (f"=({col(y)}{BP['cd_g_pub']}/{col(y)}{BP['vta_g']}-$E${BP['mp_g']}*(1+$E${BP['g_mp']})^({y}-1)"
                        f"-$E${BP['ins_g']}-$E${BP['con_g']})/$E${BP['kg_g']}"), NF_USD4, kind="calc")
    row_vals("ctot_c_pub", "U11", "Costo total asignado cajas (10b)", "USD", "20",
             [1629801, 1988039, 2316996, 2449915, 2437945, 2425620, 2422830, 2424339, 2425902, 2427522])
    row_vals("ctot_g_pub", "U12", "Costo total asignado gabinetes (10b)", "USD", "20",
             [2918277, 3462017, 3982792, 4528033, 4530340, 4531988, 4551544, 4579410, 4607662, 4636306])
    row_vals("res_c_asig", "U13", "Resultado asignado cajas = ventas - costo total asignado", "USD", "20",
             lambda y: f"={col(y)}{BP['ing_c_pub']}-{col(y)}{BP['ctot_c_pub']}", kind="calc",
             note="Año 1: -19.801 por prorrateo de costos comunes según costo directo. La contribución directa es positiva.")
    row_vals("contr_c", "U14", "Contribución directa cajas = ventas - base directa 10b", "USD", "20",
             lambda y: f"={col(y)}{BP['ing_c_pub']}-{col(y)}{BP['cd_c_pub']}", kind="calc")
    row_vals("asig_share", "U15", "Participación gabinetes en otros costos (10b) vs costo directo", "%", "20",
             lambda y: f"={col(y)}{BP['cd_g_pub']}/({col(y)}{BP['cd_g_pub']}+{col(y)}{BP['cd_c_pub']})", NF_PCT2, kind="calc",
             note="Reparto de costos comunes proporcional al costo directo (criterio provisional del PDF).")
    r += 1

    # ------------------------------------------------------------------ S8 cuadro 13 (deuda) - cronograma
    section(ws, r, "S8. Servicio de la deuda (cuadro 13) — publicado y reproducción mensual"); r += 1
    row_vals("amort_pub", "D01", "Amortizaciones publicadas", "USD", "22", [257526, 551019, 602708, 659246, 721088] + [0] * 5, total=True)
    row_vals("int_pub", "D02", "Intereses publicados", "USD", "22", [246456, 205704, 154015, 97476, 35635] + [0] * 5, total=True)
    row_vals("ivaint_pub", "D03", "IVA sobre intereses publicado", "USD", "22", [24646, 20570, 15401, 9748, 3563] + [0] * 5, total=True)
    row_vals("pctam_pub", "D04", "«% amortizaciones» publicado (fila inconsistente)", "%", "22",
             [0, 0, 0.10, 0.10, 0.10, 0.10, 0.10, 0.15, 0.15, 0.20], NF_PCT,
             note="No corresponde a los montos: se recalcula en D08.")
    ams = r + 6  # fila inicio cronograma (se completa abajo)
    row_vals("amort_calc", "D05", "Recalc. amortización (cronograma mensual abajo)", "USD", "22",
             lambda y: f"=SUMIF($F${ams}:$F${ams + 59},{y},$I${ams}:$I${ams + 59})", NF_USD2, kind="calc", total=True)
    row_vals("int_calc", "D06", "Recalc. intereses", "USD", "22",
             lambda y: f"=SUMIF($F${ams}:$F${ams + 59},{y},$H${ams}:$H${ams + 59})", NF_USD2, kind="calc", total=True)
    row_vals("ivaint_calc", "D07", "Recalc. IVA sobre intereses", "USD", "22",
             lambda y: f"=SUMIF($F${ams}:$F${ams + 59},{y},$K${ams}:$K${ams + 59})", NF_USD2, kind="calc", total=True)
    row_vals("pctam_calc", "D08", "Recalc. % amortización sobre préstamo", "%", "22",
             lambda y: f"={col(y)}{BP['amort_calc']}/$E${BP['prestamo']}", NF_PCT, kind="calc")
    row_vals("deuda_dif", "D09", "Dif. (amortización + intereses) recalc. - publicado", "USD", "",
             lambda y: f"={col(y)}{BP['amort_calc']}+{col(y)}{BP['int_calc']}-{col(y)}{BP['amort_pub']}-{col(y)}{BP['int_pub']}",
             NF_USD2, kind="dif")
    assert r == ams - 1, (r, ams)
    header(ws, r, ["", "Cronograma mensual préstamo original", "", "", "Mes", "Año", "Saldo inicial", "Interés", "Amortización",
                   "Cuota", "IVA interés", "Saldo final"]); r += 1
    BP["sch_first"] = r
    i_cell = f"$E${BP['kd']}"
    for k in range(1, 61):
        setc(ws, r, 5, k, fmt=NF_INT0)
        setc(ws, r, 6, f"=ROUNDUP(E{r}/12,0)", fmt=NF_INT0)
        setc(ws, r, 7, f"=$E${BP['prestamo']}" if k == 1 else f"=L{r - 1}", fmt=NF_USD2)
        setc(ws, r, 8, f"=G{r}*{i_cell}/12", fmt=NF_USD2)
        setc(ws, r, 10, f"=IF(E{r}<=$E${BP['gracia']},H{r},PMT({i_cell}/12,$E${BP['cuotas']},-$E${BP['prestamo']}))", fmt=NF_USD4)
        setc(ws, r, 9, f"=J{r}-H{r}", fmt=NF_USD2)
        setc(ws, r, 11, f"=H{r}*$E${BP['iva_int']}", fmt=NF_USD2)
        setc(ws, r, 12, f"=G{r}-I{r}", fmt=NF_USD2)
        r += 1
    BP["sch_last"] = r - 1
    scalar("int_gracia", "D10", "Interés mensual en gracia (control 20.936,9025)", "USD", "", f"=H{BP['sch_first']}", NF_USD4, kind="calc")
    scalar("cuota_calc", "D11", "Cuota francesa sin IVA (control 63.060,2066)", "USD", "", f"=J{BP['sch_first'] + 6}", NF_USD4, kind="calc")
    scalar("saldo_fin_sch", "D12", "Saldo final mes 60 (debe ser 0)", "USD", "", f"=L{BP['sch_last']}", NF_USD4, kind="calc")
    r += 1

    # ------------------------------------------------------------------ S9 cuadro 15 / 16
    section(ws, r, "S9. Resultados (cuadro 15) y capital de trabajo (cuadro 16)"); r += 1
    row_vals("ebitda_pub", "R01", "EBITDA publicado («fondos generados en operaciones»)", "USD", "23",
             [573024, 1029011, 1479759, 1886306, 1914796, 1973999, 2033657, 2093768, 2154335, 2215355], total=True)
    row_vals("ut_pub", "R02", "Utilidad neta publicada", "USD", "23",
             [301922, 731144, 1238751, 1707489, 1804006, 1902407, 1974240, 2034352, 2094918, 2155939], total=True)
    row_vals("payout_pub", "R03", "Dividendos devengados (% de utilidad)", "%", "23",
             [0, 0, 0.30, 0.50] + [1.0] * 6, NF_PCT)
    row_vals("ebitda_calc", "R04", "Recalc. EBITDA = ventas - producción - administración (publicados)", "USD", "23",
             lambda y: f"={col(y)}{BP['ing_pub']}-{P('PROD', y)}-{P('ADM', y)}", kind="calc", total=True)
    row_vals("ut_calc", "R05", "Recalc. utilidad = EBITDA - depreciación - financieros", "USD", "23",
             lambda y: f"={col(y)}{BP['ebitda_pub']}-{P('DEP', y)}-{P('FIN', y)}", kind="calc", total=True)
    row_vals("ebitda_dif", "R06", "Dif. EBITDA recalc. - publicado", "USD", "", lambda y: f"={col(y)}{BP['ebitda_calc']}-{col(y)}{BP['ebitda_pub']}", NF_USD2, kind="dif")
    row_vals("ut_dif", "R07", "Dif. utilidad recalc. - publicada", "USD", "", lambda y: f"={col(y)}{BP['ut_calc']}-{col(y)}{BP['ut_pub']}", NF_USD2, kind="dif")
    row_vals("fin_dif", "R08", "Dif. gastos financieros (int. + IVA) recalc. - publicados", "USD", "",
             lambda y: f"={col(y)}{BP['int_calc']}+{col(y)}{BP['ivaint_calc']}-{P('FIN', y)}", NF_USD2, kind="dif")
    row_vals("div_decl_pub", "R09", "Dividendos declarados publicados", "USD", "23",
             [0, 0, 371625, 853745, 1804006, 1902407, 1974240, 2034352, 2094918, 2155939], total=True)
    r += 1
    row_vals("ct_caja_pub", "W01", "Caja y bancos (3 días)", "USD", "24", [39863, 50804, 61961, 71387, 72101, 72822, 73550, 74286, 75029, 75779])
    row_vals("ct_sue_pub", "W02", "Sueldos y anticipos (15 días)", "USD", "24", [17890, 17890, 18647, 19239] + [19977] * 6)
    row_vals("ct_stock_pub", "W03", "Stock de materiales (120 días)", "USD", "24", [819315, 1041023, 1262706, 1443403, 1449282, 1455220, 1461217, 1467274, 1473392, 1479571])
    row_vals("ct_cxc_pub", "W04", "«Ventas a crédito» (120 días)", "USD", "24", [1406129, 1693870, 1991928, 2235331, 2254519, 2263895, 2273411, 2283068, 2292870, 2302820])
    row_vals("ct_pt_pub", "W05", "Productos terminados (10 días)", "USD", "24", [117177, 141156, 165994, 186278, 187877, 188658, 189451, 190256, 191073, 191902])
    row_vals("ct_pmen_pub", "W06", "Proveedores menores (30 días)", "USD", "24", [63831, 70082, 77382, 83478, 85233, 85995, 86775, 87574, 88393, 89232])
    row_vals("ctn_pub", "W07", "Capital de trabajo neto publicado", "USD", "24",
             [2336543, 2874662, 3423854, 3872159, 3898522, 3914577, 3930831, 3947287, 3963948, 3980818])
    row_vals("ct_caja_calc", "W08", "Recalc. caja = ventas x 3/365", "USD", "24", lambda y: f"={col(y)}{BP['ing_pub']}*$E${BP['d_caja']}/365", kind="calc")
    row_vals("ct_sue_calc", "W09", "Recalc. sueldos = (MO fábrica + MO adm.) x 15/365", "USD", "24",
             lambda y: f"=({P('MOD', y)}+{P('MOA', y)})*$E${BP['d_sue']}/365", kind="calc")
    row_vals("ct_stock_calc", "W10", "Recalc. stock = (MP + insumos + consumibles) x 120/365", "USD", "24",
             lambda y: f"=({P('MP', y)}+{P('INS', y)}+{P('CON', y)})*$E${BP['d_stock']}/365", kind="calc",
             note="Excluye flete y despacho: base de valuación a revisar (costo de adquisición).")
    row_vals("ct_cxc_calc", "W11", "Recalc. «ventas a crédito» = (producción + adm.) x 120/365", "USD", "24",
             lambda y: f"=({P('PROD', y)}+{P('ADM', y)})*$E${BP['d_cxc']}/365", kind="calc",
             note="Hallazgo: el PDF calcula la cuenta por cobrar sobre COSTOS, no sobre ventas.")
    row_vals("ct_cxc_vtas", "W12", "Alternativa: cuentas por cobrar = ventas x 120/365", "USD", "24",
             lambda y: f"={col(y)}{BP['ing_pub']}*$E${BP['d_cxc']}/365", kind="calc")
    row_vals("ct_cxc_dif", "W13", "Dif. CxC sobre ventas - publicado", "USD", "",
             lambda y: f"={col(y)}{BP['ct_cxc_vtas']}-{col(y)}{BP['ct_cxc_pub']}", NF_USD2, kind="dif",
             note="Año 1 ≈ USD 188.391,55 (condicionado a ventas uniformes y 100% a crédito).")
    row_vals("ct_pt_calc", "W14", "Recalc. PT = (producción + adm.) x 10/365", "USD", "24",
             lambda y: f"=({P('PROD', y)}+{P('ADM', y)})*$E${BP['d_pt']}/365", kind="calc",
             note="Incluye gastos administrativos en la valuación del inventario: criterio a corregir.")
    row_vals("ctn_calc", "W15", "Recalc. CT neto (con proveedores menores publicados)", "USD", "24",
             lambda y: (f"={col(y)}{BP['ct_caja_calc']}+{col(y)}{BP['ct_sue_calc']}+{col(y)}{BP['ct_stock_calc']}+{col(y)}{BP['ct_cxc_calc']}"
                        f"+{col(y)}{BP['ct_pt_calc']}-{col(y)}{BP['ct_pmen_pub']}"), kind="calc")
    row_vals("ctn_dif", "W16", "Dif. CT neto recalc. - publicado", "USD", "", lambda y: f"={col(y)}{BP['ctn_calc']}-{col(y)}{BP['ctn_pub']}", NF_USD2, kind="dif",
             note="Proveedores menores: base de 30 días no identificada; se toma el publicado.")
    row_vals("dct_pub", "W17", "Variación del CT (año siguiente - año actual), convención PDF", "USD", "26",
             lambda y: (f"=IF({y}<10,{col(y + 1)}{BP['ctn_pub']}-{col(y)}{BP['ctn_pub']},0)" if y < 10 else "=0"), kind="calc",
             y0=lambda _: f"=F{BP['ctn_pub']}")
    scalar("dct_a23", "W18", "Variación del CT para años 2 y 3 (control 1.087.311)", "USD", "26",
           f"=F{BP['dct_pub']}+G{BP['dct_pub']}", kind="calc")
    scalar("aportes_adic", "W19", "Aportes adicionales años 1-2 (control 939.710)", "USD", "26",
           f"=E{BP['aporte_a1']}+E{BP['aporte_a2']}", kind="calc")
    scalar("ct10_vs_rec", "W20", "Dif. CT año 10 (cuadro 16) - recuperación cuadro 19", "USD", "24, 27",
           f"=O{BP['ctn_pub']}-E{BP['rec_ct']}", kind="calc", note="USD 6.525 sin explicación en el PDF.")
    r += 1

    # ------------------------------------------------------------------ S10 cuadros 17, 18, 19, 20
    section(ws, r, "S10. Punto de equilibrio (17), fuentes y usos (18), VAN-TIR (19), payback y depreciación (20)"); r += 1
    row_vals("cv_pub", "E01", "Costos variables publicados (cuadro 17)", "USD", "25",
             [3298569, 4167382, 5038164, 5748282, 5771119, 5792435, 5813958, 5835690, 5857633, 5879790])
    row_vals("cf_pub", "E02", "Costos fijos publicados (incl. financieros y depreciación)", "USD", "25",
             [1249509, 1282674, 1261624, 1229666, 1197166, 1165172, 1160416, 1168059, 1175931, 1184039])
    row_vals("pe_pub", "E03", "Punto de equilibrio publicado (unidades mezcladas)", "u", "25",
             [66042, 66878, 64585, 59031, 56245, 53557, 52197, 51429, 50692, 49985], NF_INT)
    row_vals("pe_vtas", "E04", "Recalc. equilibrio en ventas = CF / (1 - CV/ventas)", "USD", "25",
             lambda y: f"={col(y)}{BP['cf_pub']}/(1-{col(y)}{BP['cv_pub']}/{col(y)}{BP['ing_pub']})", kind="calc")
    row_vals("pe_u", "E05", "Recalc. unidades = equilibrio en ventas / ventas x unidades totales", "u", "25",
             lambda y: f"={col(y)}{BP['pe_vtas']}/{col(y)}{BP['ing_pub']}*({col(y)}{BP['vta_g']}+{col(y)}{BP['vta_c']})", NF_INT, kind="calc",
             note="Solo válido con el mix del año. Ver equilibrio por producto en Resultados_Anuales.")
    row_vals("pe_g", "E06", "Gabinetes en equilibrio (manteniendo mix)", "u", "", lambda y: f"={col(y)}{BP['pe_vtas']}/{col(y)}{BP['ing_pub']}*{col(y)}{BP['vta_g']}", NF_INT, kind="calc")
    row_vals("pe_c", "E07", "Cajas en equilibrio (manteniendo mix)", "u", "", lambda y: f"={col(y)}{BP['pe_vtas']}/{col(y)}{BP['ing_pub']}*{col(y)}{BP['vta_c']}", NF_INT, kind="calc")
    r += 1
    row_vals("u_inv", "S01", "Usos: inversiones fijas (cuadro 18)", "USD", "26", [0, 100000] + [0] * 8, y0=lambda _: f"=E{BP['inv_fija']}")
    row_vals("u_nct", "S02", "Usos: necesidad de capital de trabajo", "USD", "26",
             [538119, 549192, 448306, 26363, 16055, 16254, 16456, 16661, 16870, 0], y0=lambda _: f"=E{BP['ct_ini']}")
    row_vals("u_div", "S03", "Usos: dividendos pagados (desfase de un año)", "USD", "26",
             [0, 0, 0, 371625, 853745, 1804006, 1902407, 1974240, 2034352, 2094918])
    row_vals("f_aporte", "S04", "Fuentes: aportes propios", "USD", "26", [567782, 371928] + [0] * 8, y0=lambda _: f"=E{BP['aporte_ini']}")
    row_vals("saldo18_pub", "S05", "Saldo anual acumulado publicado (cuadro 18)", "USD", "26",
             [74058, 48512, 307842, 1029689, 1314400, 1468140, 1582933, 1685801, 1788914, 1909351])
    row_vals("saldo18_calc", "S06", "Recalc. saldo = anterior + EBITDA + aportes - usos - servicio deuda", "USD", "26",
             lambda y: (f"={'0' if y == 1 else col(y - 1) + str(BP['saldo18_pub'])}+{col(y)}{BP['ebitda_pub']}+{col(y)}{BP['f_aporte']}"
                        f"-{col(y)}{BP['u_inv']}-{col(y)}{BP['u_nct']}-{col(y)}{BP['amort_pub']}-{P('FIN', y)}-{col(y)}{BP['u_div']}"), kind="calc")
    row_vals("saldo18_dif", "S07", "Dif. saldo recalc. - publicado", "USD", "", lambda y: f"={col(y)}{BP['saldo18_calc']}-{col(y)}{BP['saldo18_pub']}", NF_USD2, kind="dif")
    r += 1
    row_vals("fl19_pub", "V01", "Flujo de beneficios netos publicado (cuadro 19)", "USD", "27",
             [-493723, -397474, 259330, 1093473, 1138455, 1957746, 2017201, 2077107, 2137465, 6249065], y0=-2991587)
    row_vals("fl19_calc", "V02", "Recalc. = EBITDA + VR + recup. CT - inversión - CT - amortización - financieros", "USD", "27",
             lambda y: (f"={col(y)}{BP['ebitda_pub']}+IF({y}=10,$E${BP['vr_af']}+$E${BP['rec_ct']},0)-{col(y)}{BP['u_inv']}"
                        f"-{col(y)}{BP['u_nct']}-{col(y)}{BP['amort_pub']}-{P('FIN', y)}"),
             kind="calc", y0=lambda _: f"=-E{BP['inv_total']}")
    row_vals("fl19_dif", "V03", "Dif. flujo recalc. - publicado", "USD", "", lambda y: f"={col(y)}{BP['fl19_calc']}-{col(y)}{BP['fl19_pub']}",
             NF_USD2, kind="dif", y0=lambda _: f"=E{BP['fl19_calc']}-E{BP['fl19_pub']}")
    scalar("tasa_exacta", "V04", "Tasa ponderada exacta = (200.000 x 16,32% + 2.791.587 x 9%) / 2.991.587", "%", "28",
           f"=(E{BP['aporte_ini']}*E{BP['ke']}+E{BP['prestamo']}*E{BP['kd']})/(E{BP['aporte_ini']}+E{BP['prestamo']})", NF_PCT4, kind="calc")
    fl = BP["fl19_pub"]
    scalar("van_calc", "V05", "VAN reproducido (tasa exacta) — PUBLICADO, METODOLOGÍA POR DEPURAR", "USD", "27",
           f"=E{fl}+NPV(E{BP['tasa_exacta']},F{fl}:O{fl})", NF_USD2, kind="calc")
    scalar("van_949", "V06", "VAN con 9,49% exacto", "USD", "27", f"=E{fl}+NPV(E{BP['w_pub']},F{fl}:O{fl})", NF_USD2, kind="calc",
           note="≈ USD 357 menos: efecto de precisión de la tasa, no error del VAN publicado.")
    scalar("tir_calc", "V07", "TIR reproducida — PUBLICADA, METODOLOGÍA POR DEPURAR", "%", "27", f"=IRR(E{fl}:O{fl},0.1)", NF_PCT4, kind="calc",
           note="El cuadro 19 descuenta toda la inversión y resta servicio de deuda sin registrar el préstamo: mezcla perspectivas.")
    scalar("van_dif", "V08", "Dif. VAN reproducido - publicado", "USD", "", f"=E{BP['van_calc']}-E{BP['van_pub']}", NF_USD2, kind="calc")
    scalar("tir_dif", "V09", "Dif. TIR reproducida - publicada", "p.p.", "", f"=E{BP['tir_calc']}-E{BP['tir_pub']}", NF_PCT4, kind="calc")
    row_vals("fl19_acum", "V10", "Flujo publicado acumulado (payback simple, misma serie del VAN)", "USD", "",
             lambda y: f"={col(y - 1)}{BP['fl19_pub'] + 0}+0" if False else f"={col(y - 1)}{r}+{col(y)}{fl}",
             kind="calc", y0=lambda _: f"=E{fl}")
    row_vals("fl19_desc", "V11", "Flujo publicado descontado (tasa exacta)", "USD", "",
             lambda y: f"={col(y)}{fl}/(1+$E${BP['tasa_exacta']})^{y}", kind="calc", y0=lambda _: f"=E{fl}")
    row_vals("fl19_dacum", "V12", "Descontado acumulado (payback descontado)", "USD", "",
             lambda y: f"={col(y - 1)}{r}+{col(y)}{BP['fl19_desc']}", kind="calc", y0=lambda _: f"=E{BP['fl19_desc']}")

    def payback_formula(rowacc, rowfl):
        # último cruce de negativo a no negativo con interpolación lineal anual
        parts = []
        for y in range(1, 11):
            parts.append(f"IF(AND({col(y - 1)}{rowacc}<0,{col(y)}{rowacc}>=0),{y - 1}+(-{col(y - 1)}{rowacc})/{col(y)}{rowfl},0)")
        return "=IF(O{0}<0,\"No recupera en el horizonte\",MAX({1}))".format(rowacc, ",".join(parts))
    scalar("pb_simple", "V13", "Payback simple (misma serie del VAN, interpolación uniforme)", "años", "",
           payback_formula(BP["fl19_acum"], fl), NF_USD2, kind="calc", note="≈ 5,71 años.")
    scalar("pb_desc", "V14", "Payback descontado (misma serie, tasa exacta)", "años", "",
           payback_formula(BP["fl19_dacum"], BP["fl19_desc"]), NF_USD2, kind="calc", note="≈ 6,89 años (cruce durante el año 7).")
    row_vals("pb_pdf_ser", "V15", "Serie de payback publicada (parte de -2.791.587)", "USD", "28",
             [-3242520, -3574082, -3376505, -2615618, -1892088, -755706, 313707, 1319441, 2264701, 4788736], y0=-2791587,
             note="Inicia en -2.791.587 (préstamo), no en -2.991.587 como el VAN: diferencia USD 200.000.")
    r += 1
    section(ws, r, "S10b. Reexpresión de perspectivas con datos del PDF (anual, fin de año) — conciliación metodológica"); r += 1
    row_vals("fcff_pdf", "X01", "FCFF con datos PDF = EBITDA - inversión - ΔCT + (año 10: recup. CT + VR)", "USD", "",
             lambda y: (f"={col(y)}{BP['ebitda_pub']}-{col(y)}{BP['u_inv']}-{col(y)}{BP['u_nct']}"
                        f"+IF({y}=10,$E${BP['rec_ct']}+$E${BP['vr_af']},0)"),
             kind="calc", y0=lambda _: f"=-E{BP['inv_total']}",
             note="Excluye préstamo, intereses y amortizaciones. Mantiene los supuestos (y errores) de CT del PDF.")
    scalar("van_fcff_pdf", "X02", "VAN FCFF datos PDF a la tasa ponderada exacta", "USD", "",
           f"=E{BP['fcff_pdf']}+NPV(E{BP['tasa_exacta']},F{BP['fcff_pdf']}:O{BP['fcff_pdf']})", NF_USD, kind="calc")
    scalar("tir_fcff_pdf", "X03", "TIR FCFF datos PDF (TIR del proyecto)", "%", "", f"=IRR(E{BP['fcff_pdf']}:O{BP['fcff_pdf']},0.1)", NF_PCT2, kind="calc")
    row_vals("eq_pdf", "X04", "Flujo socios datos PDF = -aportes + dividendos pagados + (año 10: saldo caja + recup. CT + VR)", "USD", "",
             lambda y: (f"=-{col(y)}{BP['f_aporte']}+{col(y)}{BP['u_div']}"
                        f"+IF({y}=10,{col(y)}{BP['saldo18_pub']}+$E${BP['rec_ct']}+$E${BP['vr_af']},0)"),
             kind="calc", y0=lambda _: f"=-E{BP['aporte_ini']}",
             note="Valor final atribuible: caja acumulada no distribuida + CT + VR (deuda ya cancelada). No se suma FCFE no distribuido.")
    scalar("van_eq_pdf", "X05", "VAN socios datos PDF a Ke 16,32%", "USD", "",
           f"=E{BP['eq_pdf']}+NPV(E{BP['ke']},F{BP['eq_pdf']}:O{BP['eq_pdf']})", NF_USD, kind="calc")
    scalar("tir_eq_pdf", "X06", "TIR socios datos PDF", "%", "", f"=IRR(E{BP['eq_pdf']}:O{BP['eq_pdf']},0.2)", NF_PCT2, kind="calc")
    r += 1
    row_vals("dep_maq_pub", "Z01", "Depreciación maquinaria y equipos (507.166 / 10)", "USD", "29", [0] + [50717] * 9)
    row_vals("dep_inf_pub", "Z02", "Depreciación infraestructura (87.000 / 10)", "USD", "29", [0] + [8700] * 9)
    row_vals("dep_otr_pub", "Z03", "Depreciación otras inversiones (60.878 / 5)", "USD", "29", [0] + [12176] * 5 + [0] * 4)
    row_vals("dep_tot_calc", "Z04", "Recalc. depreciación total", "USD", "29",
             lambda y: f"={col(y)}{BP['dep_maq_pub']}+{col(y)}{BP['dep_inf_pub']}+{col(y)}{BP['dep_otr_pub']}", kind="calc", total=True,
             note="Empieza en año 2 (nota del cuadro dice año 3). No incluye el CAPEX de USD 100.000 del año 2.")
    row_vals("dep_dif", "Z05", "Dif. vs cuadro 10a", "USD", "", lambda y: f"={col(y)}{BP['dep_tot_calc']}-{P('DEP', y)}", NF_USD2, kind="dif")
    scalar("vnc_10", "Z06", "Valor neto contable al cierre del año 10 (activos del cuadro 20)", "USD", "29",
           f"=E{BP['inv_fija']}-P{BP['dep_tot_calc']}", kind="calc",
           note="≈ valor residual 59.417 publicado: es valor contable, no precio de venta realizable.")
    BP["_last"] = r
    return ws


build_base_pdf()


# ===========================================================================
# 2) SUPUESTOS
# ===========================================================================
SUP = {}                 # clave -> fila
SCEN = ["Base depurada (supuestos Oct-2025)", "Actualizado", "Conservador", "Recuperación operativa", "Personalizado"]
SHOCKS = [
    # nombre, capex, atraso, vol, precio, mat, tc, dso
    ("Ninguna", 0, 0, 0, 0, 0, 0, None),
    ("CAPEX +10%", 0.10, 0, 0, 0, 0, 0, None),
    ("CAPEX +20%", 0.20, 0, 0, 0, 0, 0, None),
    ("Atraso +3 meses", 0, 3, 0, 0, 0, 0, None),
    ("Atraso +6 meses", 0, 6, 0, 0, 0, 0, None),
    ("Volumen -10%", 0, 0, -0.10, 0, 0, 0, None),
    ("Volumen -20%", 0, 0, -0.20, 0, 0, 0, None),
    ("Precio -5%", 0, 0, 0, -0.05, 0, 0, None),
    ("Precio -10%", 0, 0, 0, -0.10, 0, 0, None),
    ("Materiales +10%", 0, 0, 0, 0, 0.10, 0, None),
    ("Materiales +20%", 0, 0, 0, 0, 0.20, 0, None),
    ("PYG/USD +10% (más Gs por USD)", 0, 0, 0, 0, 0, 0.10, None),
    ("PYG/USD -10% (menos Gs por USD)", 0, 0, 0, 0, 0, -0.10, None),
    ("DSO 90 días", 0, 0, 0, 0, 0, 0, 90),
    ("DSO 120 días", 0, 0, 0, 0, 0, 0, 120),
    ("DSO 150 días", 0, 0, 0, 0, 0, 0, 150),
]
# Datos_Reales: celdas de fechas y corte (fijas)
DR_T1 = {"r_FechaImpl": 6, "r_Arranque": 7, "r_Corte": 8, "r_FechaCorte": 9}


def build_supuestos():
    ws = WS["SUP"]
    CUR[0] = "SUP"
    title(ws, "Supuestos — selector de escenario, parámetros y drivers",
          "Celdas amarillas con texto azul = entradas editables. Gris itálica = hereda de la base (sobrescribir para cambiar). "
          "Los parámetros tributarios y económicos no están embebidos en fórmulas.")
    for k, v in {1: 8, 2: 50, 3: 13, 4: 16, 5: 16, 6: 16, 7: 16, 8: 16, 9: 14, 10: 14, 11: 16, 12: 34, 13: 60}.items():
        ws.column_dimensions[CL(k)].width = v
    r = 4
    section(ws, r, "A. Selector de escenario y prueba de sensibilidad", 13); r += 1
    setc(ws, r, 2, "Escenario activo (único selector de supuestos futuros)", font=F_BOLD)
    setc(ws, r, 4, SCEN[0], font=F_IN, fill=FILL_PEND)
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
    add_name("p_Escenario", "SUP", f"D{r}")
    SUP["esc"] = r
    setc(ws, r, 8, "Cambia solo el futuro: los meses reales (hasta el último mes cerrado) no se modifican.", font=F_SUB)
    r += 1
    setc(ws, r, 2, "Índice del escenario activo")
    SUP["esc_idx"] = r
    r += 1
    setc(ws, r, 2, "Estado del escenario activo", font=F_BOLD)
    SUP["esc_status"] = r
    r += 1
    setc(ws, r, 2, "Prueba de sensibilidad activa (choque ilustrativo sobre el escenario)", font=F_BOLD)
    setc(ws, r, 4, SHOCKS[0][0], font=F_IN, fill=FILL_PEND)
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=6)
    add_name("p_Shock", "SUP", f"D{r}")
    SUP["shock"] = r
    setc(ws, r, 8, "Choques de prueba del modelo, no premisas de negocio. Volver a «Ninguna» al terminar.", font=F_SUB)
    r += 1
    setc(ws, r, 2, "Índice de la prueba activa")
    SUP["shock_idx"] = r
    r += 2

    # ---------------------------------------------------------------- B generales
    section(ws, r, "B. Parámetros generales (no dependen del escenario)", 13); r += 1
    header(ws, r, ["Cód.", "Parámetro", "Unidad / moneda", "Valor", "Vigencia", "Fuente / pág.", "Estado", "Nota"])
    r += 1

    def gp(key, code, label, unit, value, vig, src, status, name=None, fmt=None, note=None, inp=True, pend=False):
        nonlocal r
        SUP[key] = r
        setc(ws, r, 1, code)
        setc(ws, r, 2, label)
        setc(ws, r, 3, unit)
        is_formula = isinstance(value, str) and value.startswith("=")
        link = is_formula and ("Base_PDF!" in value)
        font = F_LINK if link else (F_IN if (inp and not is_formula) else F_BASE)
        fill = FILL_PEND if pend else (FILL_IN if (inp and not is_formula) else None)
        setc(ws, r, 4, value, font=font, fill=fill, fmt=fmt)
        setc(ws, r, 5, vig)
        setc(ws, r, 6, src)
        setc(ws, r, 7, status, font=F_ALERT if status.startswith("Pendiente") else F_BASE)
        if note:
            setc(ws, r, 8, note)
        if name:
            add_name(name, "SUP", f"D{r}")
        r += 1

    BPs = lambda k: f"=Base_PDF!$E${BP[k]}"
    gp("tc_plan", "G01", "Tipo de cambio de referencia del plan", "Gs/USD", BPs("tc_ref"), "Histórico Oct-2025", "PDF p. 4-5", "Publicado (PDF)", "p_TC_Plan", NF_INT)
    gp("tc_nom", "G02", "TC implícito en la nómina del PDF (solo para reproducir)", "Gs/USD", BPs("tc_nom"), "Histórico", "PDF p. 30-34", "Inferido", "p_TC_NomPDF", NF_INT,
       "Se usa únicamente si la corrección C1 está desactivada.")
    gp("impl", "G03", "Meses de implantación previstos antes del primer mes operativo", "meses", 0, "Plan", "PDF: año 0 puntual", "Supuesto provisional", "p_ImplMeses", NF_INT0,
       "El PDF concentra la inversión en el año 0 y opera desde el año 1. Cargar el cronograma real si difiere.")
    gp("fecha_ini", "G04", "Fecha calendario del mes 0 (inicio de implantación / desembolso)", "fecha", None, "—", "No informada en el PDF", "Pendiente de confirmar",
       "p_FechaInicio", NF_DATE, "Sin fecha el modelo usa meses relativos. No se infiere que año 1 = 2025 o 2026.", pend=True)
    gp("modo_hor", "G05", "Modo de horizonte (1 = fecha final fija; 2 = 10 años desde el arranque efectivo)", "1/2", 1, "Todo", "Prompt 5.6", "Supuesto provisional", "p_ModoHor", NF_INT0,
       "Con 1 un atraso acorta la operación (no prolonga la vida). Con 2 se compara 10 años desde la puesta en marcha efectiva.")
    gp("anios_op", "G06", "Años operativos del plan", "años", 10, "Todo", "PDF", "Publicado (PDF)", "p_AniosOp", NF_INT0)
    gp("meses_op", "G07", "Meses operativos del plan", "meses", "=p_AniosOp*12", "Todo", "Cálculo", "Cálculo", "p_MesesOp", NF_INT0, inp=False)
    gp("dias_anio", "G08", "Días por año (convención de días de CT)", "días", 365, "Todo", "PDF cuadro 16", "Publicado (PDF)", "p_DiasAnio", NF_INT0)
    gp("dias_mes", "G09", "Días promedio por mes", "días", "=p_DiasAnio/12", "Todo", "Cálculo", "Cálculo", "p_DiasMes", NF_USD2, inp=False)
    gp("ke", "G10", "Costo del capital propio Ke (anual efectiva, USD nominal)", "%", BPs("ke"), "Histórico", "PDF p. 9", "Publicado (PDF) - histórico no vigente confirmado", "p_Ke", NF_PCT2,
       "Hipótesis histórica. Actualizar con fecha y fundamento (tasa libre de riesgo, prima, riesgo país, tamaño).")
    gp("kd", "G11", "Tasa bancaria de referencia (nominal anual)", "%", BPs("kd"), "Histórico", "PDF p. 4", "Publicado (PDF) - histórico", "p_Kd", NF_PCT2)
    gp("metodo_tasa", "G12", "Tasa de descuento del proyecto: 1 = ponderada inicial PDF; 2 = WACC con estructura objetivo; 3 = manual", "1/2/3", 1, "Todo",
       "Prompt 8.5", "Supuesto provisional", "p_MetodoTasa", NF_INT0,
       "La ponderada inicial usa pesos del momento 0 (6,7%/93,3%), que no representan la estructura durante 10 años.")
    gp("dv_obj", "G13", "Peso objetivo de la deuda D/(D+E) para el método 2", "%", None, "Todo", "—", "Pendiente de definir", "p_DV", NF_PCT,
       "Definir estructura objetivo justificada o usar APV si el apalancamiento cambia materialmente.", pend=True)
    gp("kd_ef", "G14", "Costo efectivo de la deuda = tasa x (1 + IVA intereses no recuperable)", "%", "=p_Kd*(1+p_IVAint)", "Todo", "Cálculo", "Cálculo",
       "p_KdEf", NF_PCT2, "Sin escudo de IRE: el plan asume IRE no aplicable bajo maquila (verificar).", inp=False)
    gp("tasa_man", "G15", "Tasa manual del proyecto (método 3)", "%", None, "Todo", "—", "Pendiente de definir", "p_TasaManual", NF_PCT2, pend=True)
    gp("tasa_pdf", "G16", "Tasa ponderada inicial exacta (método 1)", "%", f"=Base_PDF!$E${BP['tasa_exacta']}", "Histórico", "PDF p. 28", "Publicado (PDF) - cálculo exacto",
       "p_TasaPDF", NF_PCT4)
    gp("tasa_proy", "G17", "Tasa de descuento del proyecto aplicada (anual efectiva)", "%",
       '=IF(p_MetodoTasa=2,IF(p_DV="",p_TasaPDF,p_DV*p_KdEf+(1-p_DV)*p_Ke),IF(p_MetodoTasa=3,IF(p_TasaManual="",p_TasaPDF,p_TasaManual),p_TasaPDF))',
       "Todo", "Cálculo", "Cálculo", "p_TasaProy", NF_PCT4, "Si faltan datos del método elegido, vuelve al método 1 y lo informa en el estado siguiente.", inp=False)
    gp("tasa_estado", "G18", "Estado de la tasa del proyecto", "texto",
       '=IF(AND(p_MetodoTasa=2,p_DV=""),"Método 2 sin peso objetivo: se usa ponderada inicial",IF(AND(p_MetodoTasa=3,p_TasaManual=""),"Método 3 sin tasa: se usa ponderada inicial",'
       'IF(p_MetodoTasa=1,"Ponderada inicial PDF (histórica)","OK")))', "Todo", "Cálculo", "Cálculo", "p_TasaEstado", None, inp=False)
    gp("ire", "G19", "IRE sobre resultado (régimen de maquila)", "%", 0, "Todo", "PDF p. 7-8", "Publicado (PDF) - verificar Ley 7547/2025", "p_IRE", NF_PCT,
       "El plan considera IRE no aplicable. No se aplica escudo fiscal de intereses. Impuesto a dividendos: pendiente de verificar.")
    gp("tipo_vt", "G20", "Valor terminal: 1 = liquidación realizable; 2 = negocio en marcha (perpetuidad)", "1/2", 1, "Fin de horizonte", "Prompt 8.9", "Supuesto provisional", "p_TipoVT", NF_INT0,
       "En negocio en marcha no se recupera el CT ni se vende el activo (se necesitan para operar).")
    gp("g", "G21", "Crecimiento perpetuo nominal en USD (solo método 2)", "%", 0, "Perpetuidad", "—", "Supuesto provisional", "p_g", NF_PCT)
    gp("real_cxc", "G22", "Realización de cuentas por cobrar al cierre (liquidación)", "%", 1, "Fin de horizonte", "—", "Supuesto provisional", "p_RealCxC", NF_PCT)
    gp("real_inv", "G23", "Realización de inventarios al cierre (liquidación)", "%", 1, "Fin de horizonte", "—", "Supuesto provisional", "p_RealInv", NF_PCT)
    gp("real_af", "G24", "Realización del activo fijo (% del valor neto contable)", "%", 1, "Fin de horizonte", "—", "Supuesto provisional", "p_RealAF", NF_PCT,
       "Un valor contable residual no es automáticamente precio de venta. Reemplazar por tasación.")
    gp("costo_cierre", "G25", "Costos e impuestos de cierre / liquidación", "USD", 0, "Fin de horizonte", "—", "Pendiente de estimar", "p_CostoCierre", NF_USD, pend=True)
    gp("modalidad", "G26", "Modalidad de facturación: 1 = venta de producto (PDF); 2 = servicio de maquila", "1/2", 1, "Todo", "PDF p. 6; contrato pendiente", "Pendiente de confirmar",
       "p_Modalidad", NF_INT0, "Con 2 los ingresos son tarifa de servicio y los materiales del cliente no son costo ni inventario propio.")
    gp("tarifa_g", "G27", "Tarifa de servicio de maquila por gabinete (modalidad 2)", "USD/u", None, "Año 1", "Contrato", "Pendiente de cargar", "p_TarifaG", NF_USD2, pend=True)
    gp("tarifa_c", "G28", "Tarifa de servicio de maquila por caja (modalidad 2)", "USD/u", None, "Año 1", "Contrato", "Pendiente de cargar", "p_TarifaC", NF_USD2, pend=True)
    gp("prop_cli", "G29", "% de materiales (MP, insumos, flete MP) propiedad del cliente en modalidad 2", "%", 1, "Todo", "Contrato", "Pendiente de confirmar", "p_PropCliente", NF_PCT)
    gp("driver", "G30", "Driver de costos comunes: 1 = kg procesado; 2 = costo directo (criterio PDF); 3 = horas estándar", "1/2/3", 1, "Todo", "Prompt 6.6",
       "Supuesto provisional", "p_Driver", NF_INT0, "Reparto provisional. Reemplazar por horas máquina/hombre cuando se disponga.")
    gp("kg_g", "G31", "Peso por gabinete", "kg/u", BPs("kg_g"), "Todo", "PDF p. 1", "Publicado (PDF)", "p_KgG", NF_INT)
    gp("kg_c", "G32", "Peso por caja", "kg/u", BPs("kg_c"), "Todo", "PDF p. 1", "Publicado (PDF)", "p_KgC", NF_INT)
    gp("hs_g", "G33", "Horas estándar por gabinete (driver 3)", "h/u", None, "Todo", "Ingeniería", "Pendiente de cargar", "p_HsG", NF_USD2, pend=True)
    gp("hs_c", "G34", "Horas estándar por caja (driver 3)", "h/u", None, "Todo", "Ingeniería", "Pendiente de cargar", "p_HsC", NF_USD2, pend=True)
    gp("m2", "G35", "Superficie de nave", "m2", BPs("m2"), "Todo", "PDF p. 6", "Publicado (PDF)", "p_M2", NF_INT)
    gp("alq_m2", "G36", "Alquiler neto de IVA", "USD/m2/mes", BPs("alq_m2"), "Año 1", "PDF p. 6", "Publicado (PDF)", "p_AlqM2", NF_USD2)
    gp("iva_alq", "G37", "IVA incluido en el alquiler", "%", BPs("iva_alq"), "Todo", "PDF p. 6", "Publicado (PDF)", "p_IVAalq", NF_PCT)
    gp("alq_g", "G38", "Incremento anual de alquiler y expensas", "%", 0.03, "Desde año 2", "PDF cuadro 10a (inferido)", "Inferido", "p_AlqG", NF_PCT)
    gp("exp_m2", "G39", "Expensas", "USD/m2/mes", BPs("exp_m2"), "Año 1", "PDF p. 19", "Publicado (PDF)", "p_ExpM2", NF_USD2)
    gp("alq_caja", "G40", "Alquiler: 1 = pago real a tercero; 2 = pago a relacionada (CIE); 3 = costo económico sin desembolso", "1/2/3", 1, "Todo",
       "PDF p. 6 («costo de oportunidad»)", "Pendiente de confirmar", "p_AlqCaja", NF_INT0,
       "Con 3 el alquiler sigue en resultados y en el FCFF (costo económico) pero no sale de tesorería; se registra como aporte en especie.")
    gp("iva_int", "G41", "IVA sobre intereses (costo, no recuperable)", "%", BPs("iva_int"), "Todo", "PDF p. 4", "Publicado (PDF)", "p_IVAint", NF_PCT)
    gp("trib", "G42", "Tributo único de maquila", "%", BPs("trib"), "Todo", "PDF p. 7; Ley 7547/2025 (fuente secundaria)", "Publicado (PDF) - verificar", "p_Trib", NF_PCT)
    gp("base_trib", "G43", "Base del tributo: 1 = factura de exportación (PDF); 2 = mayor entre factura y valor agregado nacional", "1/2", 1, "Todo",
       "Ley 7547/2025 y Dec. 5714/2026 según fuentes secundarias", "Pendiente de verificar en texto oficial", "p_BaseTrib", NF_INT0,
       "No se pudo acceder a mic.gov.py ni bacn.gov.py desde este entorno (29/09/2026). Verificar texto vigente y transición.")
    gp("tope_iva", "G44", "Tope de recuperación de IVA para maquila de servicios (% de facturación)", "%", 0.005, "Todo",
       "Ley 7547/2025 según fuentes secundarias", "Pendiente de verificar en texto oficial", "p_TopeIVAServ", NF_PCT2, "Se aplica solo en modalidad 2.")
    gp("segc", "G45", "Seguro de caución maquila (% de ventas)", "%", f"=Base_PDF!$E${BP['segc_pct']}", "Todo", "PDF cuadro 10a (inferido exacto)", "Inferido", "p_SegCau", NF_PCT2)
    gp("desp", "G46", "Despacho de importación (% de MP + insumos)", "%", BPs("desp"), "Todo", "PDF p. 19; base inferida", "Publicado (PDF) / base inferida", "p_Desp", NF_PCT)
    gp("pkg", "G47", "Packaging (% de materiales + flete + MO directa con cargas y alimentación)", "%", BPs("pkg"), "Todo", "PDF p. 19; base inferida", "Publicado (PDF) / base inferida",
       "p_Pkg", NF_PCT, "Base reproducida del PDF. Reemplazar por USD/unidad si se dispone del costo real de embalaje.")
    gp("impp", "G48", "Imprevistos de producción (% costos producción excl. imprevistos y tributo)", "%", BPs("imp"), "Todo", "PDF p. 19; base inferida", "Publicado (PDF) / base inferida", "p_ImpP", NF_PCT)
    gp("impa", "G49", "Imprevistos de administración (% adm. excl. alquiler y expensas)", "%", BPs("imp"), "Todo", "PDF p. 19; base inferida", "Publicado (PDF) / base inferida", "p_ImpA", NF_PCT)
    gp("prest", "G50", "Prestaciones sociales (coeficiente del plan)", "% s/ remun.", BPs("prest"), "Todo", "PDF p. 19", "Publicado (PDF) - no validado legalmente", "p_Prest", NF_PCT2,
       "Desagregar IPS, aguinaldo, vacaciones y provisiones cuando se aporten, evitando duplicaciones.")
    gp("alim", "G51", "Alimentación por persona", "Gs/mes", BPs("alim_gs"), "Todo", "PDF cuadro 10a (inferido)", "Inferido", "p_AlimGs", NF_GS)
    gp("man_inc", "G52", "Mantenimiento de planta: incremento lineal anual (% del año 1)", "%", f"=Base_PDF!$E${BP['man_inc']}", "Desde año 2", "PDF cuadro 10a (inferido)", "Inferido", "p_ManInc", NF_PCT)
    gp("div_lag", "G53", "Meses entre el cierre del año y el pago de dividendos", "meses", 4, "Todo", "PDF: pago al año siguiente", "Supuesto provisional", "p_DivLag", NF_INT0)
    gp("div_restr", "G54", "Pagar dividendos solo con caja sobre el mínimo (1 = sí; 0 = no)", "1/0", 1, "Todo", "Prompt 8.2", "Supuesto provisional", "p_DivRestr", NF_INT0,
       "Los dividendos no se pagan solo porque existe utilidad.")
    gp("cmin_d", "G55", "Caja mínima operativa (días de ventas del mes)", "días", BPs("d_caja"), "Todo", "PDF cuadro 16 (3 días)", "Publicado (PDF)", "p_CajaMinDias", NF_INT0,
       "Se modela en tesorería, separada del capital de trabajo operativo.")
    gp("cmin_f", "G56", "Caja mínima fija adicional", "USD", 0, "Todo", "—", "Supuesto provisional", "p_CajaMinFija", NF_USD)
    gp("esp_ind", "G57", "Costos de espera: % del personal indirecto y administrativo activo antes del arranque", "%", 1, "Meses pre-operativos", "—",
       "Supuesto provisional (ilustrativo)", "p_EspInd", NF_PCT, "Aplica a meses de implantación/atraso. Reemplazar por dotación real.")
    gp("esp_dir", "G58", "Costos de espera: % del personal directo activo antes del arranque", "%", 0.5, "Meses pre-operativos", "—", "Supuesto provisional (ilustrativo)", "p_EspDir", NF_PCT)
    gp("esp_ene", "G59", "Costos de espera: % de energía y mantenimiento de planta", "%", 0.25, "Meses pre-operativos", "—", "Supuesto provisional (ilustrativo)", "p_EspEne", NF_PCT)
    gp("esp_fij", "G60", "Costos de espera: % de alquiler, expensas y costos fijos administrativos", "%", 1, "Meses pre-operativos", "—", "Supuesto provisional (ilustrativo)", "p_EspFijos", NF_PCT)
    gp("mes_rev", "G61", "Mes de revisión para la evaluación incremental (continuar vs. salir; -1 = desde el inicio)", "mes modelo", "=p_MesCorte", "—", "Prompt 8.8", "Cálculo (editable)",
       "p_MesRev", NF_INT0, "Por defecto = último mes cerrado. Los costos anteriores son hundidos y no se cargan.", inp=False)
    gp("val_salida", "G62", "Alternativa de salida: valor neto recuperable en el mes de revisión", "USD", None, "Mes de revisión", "—", "Pendiente de estimar", "p_ValorSalida", NF_USD,
       "Venta de activos + recuperación de CT - pasivos a cancelar - costos de salida.", pend=True)
    gp("ap_auto", "G63", "Aportes hipotéticos de socios para cubrir la brecha de caja (1 = activar) — NO APROBADOS", "1/0", 0, "Horizonte", "Prompt 7.8",
       "Supuesto provisional", "p_ApAuto", NF_INT0,
       "Apagado: la brecha queda visible y el VAN/TIR de socios supone cobertura implícita sin costo. Encendido: mide el esfuerzo de los socios.")
    r += 1

    # ---------------------------------------------------------------- préstamo original
    section(ws, r, "B2. Préstamo original (condiciones contractuales; modificar solo si el contrato lo establece)", 13); r += 1
    gp("l1_monto", "L01", "Préstamo 1: monto", "USD", BPs("prestamo"), "Contrato", "PDF p. 4", "Publicado (PDF) - confirmar contrato", "p_L1_Monto", NF_USD)
    gp("l1_mes", "L02", "Préstamo 1: mes de desembolso", "mes modelo", 0, "Contrato", "PDF: año 0", "Supuesto provisional", "p_L1_Mes", NF_INT0)
    gp("l1_tasa", "L03", "Préstamo 1: tasa nominal anual (mensual = tasa/12)", "%", BPs("kd"), "Contrato", "PDF p. 4", "Publicado (PDF)", "p_L1_Tasa", NF_PCT2,
       "Tasa nominal bancaria con capitalización mensual: tasa mensual = nominal/12 (no aplica a tasas efectivas anuales).")
    gp("l1_gracia", "L04", "Préstamo 1: meses de gracia (solo intereses)", "meses", BPs("gracia"), "Contrato", "PDF p. 4", "Publicado (PDF)", "p_L1_Gracia", NF_INT0)
    gp("l1_cuotas", "L05", "Préstamo 1: cuotas francesas", "meses", BPs("cuotas"), "Contrato", "PDF p. 4", "Publicado (PDF)", "p_L1_Cuotas", NF_INT0)
    gp("l1_com", "L06", "Préstamo 1: comisión inicial (% del monto)", "%", 0, "Contrato", "No informada", "Pendiente de confirmar", "p_L1_Com", NF_PCT2)
    r += 1

    # ---------------------------------------------------------------- C correcciones
    section(ws, r, "C. Correcciones metodológicas (1 = aplicada; 0 = criterio del PDF). Base depurada = todas en 1", 13); r += 1
    header(ws, r, ["Cód.", "Corrección", "Valor", "", "Efecto esperado", "", "Fuente", "Detalle"]); r += 1
    corr = [
        ("C1", "Tipo de cambio único para la nómina (Gs/USD del escenario en lugar de 8.000 implícito)", "Mayor costo USD de nómina (≈ USD 59.730/año a 7.300).", "Informe validación 1"),
        ("C2", "Materia prima de cajas con incremento acumulativo de 1% anual (texto del PDF)", "Mayor costo desde año 3 (≈ USD 74.413 en año 10).", "Informe validación 4"),
        ("C3", "Depreciación desde la puesta en servicio (año 1) y no desde el año 2", "Resultado contable; en FCFF solo vía valor terminal si la realización del activo se aproxima con su valor contable (G24).", "Informe validación 9"),
        ("C4", "CAPEX adicional del año 2 incorporado al activo depreciable (vida provisional)", "Resultado contable; menor valor contable final (afecta el valor terminal vía G24).", "Informe validación 9"),
        ("C5", "Inventario de materiales valuado a costo de adquisición (incluye flete y despacho)", "Más capital de trabajo inmovilizado en stock.", "Prompt 2.5 / 7.4"),
        ("C6", "Supervisión e indirectos de fábrica clasificados como costo de producción", "Cambia costo de ventas/adm. y valuación de PT; EBITDA casi igual.", "Prompt 6.5"),
    ]
    SUP["corr_first"] = r
    for code, label, eff, src in corr:
        SUP[code] = r
        setc(ws, r, 1, code, font=F_BOLD)
        setc(ws, r, 2, label)
        setc(ws, r, 3, 1, font=F_IN, fill=FILL_IN, fmt=NF_INT0)
        add_name(f"p_{code}", "SUP", f"C{r}")
        setc(ws, r, 5, eff)
        setc(ws, r, 7, src)
        r += 1
    SUP["corr_last"] = r - 1
    setc(ws, r, 2, "Cantidad de correcciones activas (6 = base depurada completa)")
    setc(ws, r, 3, f"=SUM(C{SUP['corr_first']}:C{SUP['corr_last']})", fmt=NF_INT0)
    add_name("p_NCorr", "SUP", f"C{r}")
    r += 1
    setc(ws, r, 2, "Correcciones inherentes al motor (no desactivables): CxC por calendario de ventas; CT mensual; FCFF y flujo de socios separados; "
                   "payback sobre la misma serie; márgenes rotulados; valor terminal explícito; dividendos pagados con caja disponible.", font=F_SUB)
    r += 2

    # ---------------------------------------------------------------- D escenarios
    section(ws, r, "D. Drivers por escenario (columna J = valor efectivo que usa el motor, incluye la prueba de sensibilidad)", 13); r += 1
    header(ws, r, ["Cód.", "Driver", "Unidad"] + SCEN + ["Activo", "Efectivo", "Vigencia", "Fuente / estado", "Nota"])
    SUP["scen_hdr"] = r
    r += 1
    setc(ws, r, 2, "Estado del escenario", font=F_BOLD)
    status = ["Supuestos históricos Oct-2025 con correcciones metodológicas. NO es ejecución real.",
              "=IF(p_MesCorte<0,\"PENDIENTE: sin datos reales cargados; hereda supuestos históricos. No presentar como actualizado.\","
              "\"Reales hasta mes \"&p_MesCorte&\"; forecast con supuestos heredados salvo los modificados. Verificar campos mínimos.\")",
              "ILUSTRATIVO: choques de prueba, no premisas aprobadas.",
              "ILUSTRATIVO: atraso con recuperación parcial y mejoras de CT, no premisas aprobadas.",
              "Combinación definida por el usuario (hereda la base hasta que se modifique)."]
    for i, s in enumerate(status):
        setc(ws, r, 4 + i, s, font=F_SUB, align=WRAP)
    ws.row_dimensions[r].height = 48
    SUP["scen_status"] = r
    r += 1

    def drv(key, code, label, unit, vals, vig, src, name, fmt, shock=None, note=None, rng=False):
        """vals: [base, act, cons, rec, pers]; '=' hereda la base; str con '=' fórmula."""
        nonlocal r
        SUP[key] = r
        setc(ws, r, 1, code)
        setc(ws, r, 2, label)
        setc(ws, r, 3, unit)
        for i, v in enumerate(vals):
            colx = 4 + i
            if i > 0 and v == "=":
                setc(ws, r, colx, f"=$D{r}", font=F_INH, fill=FILL_INH, fmt=fmt)
            elif isinstance(v, str) and v.startswith("="):
                setc(ws, r, colx, v, font=F_LINK if "Base_PDF" in v else F_BASE, fmt=fmt, fill=FILL_IN if i > 0 else None)
            else:
                setc(ws, r, colx, v, font=F_IN, fill=FILL_IN, fmt=fmt)
        setc(ws, r, 9, f"=INDEX($D{r}:$H{r},1,p_EscIdx)", fmt=fmt)
        eff = shock.replace("@", f"$I{r}") if shock else f"=$I{r}"
        setc(ws, r, 10, eff, fmt=fmt, font=F_BOLD)
        setc(ws, r, 11, vig)
        setc(ws, r, 12, src)
        if note:
            setc(ws, r, 13, note)
        if name and not rng:
            add_name(name, "SUP", f"J{r}")
        r += 1

    E = "="
    drv("atraso", "E01", "Atraso de arranque respecto del plan", "meses", [0, E, 3, 6, E], "Arranque", "Ilustrativo en escenarios 3-4", "e_Atraso", NF_INT0,
        "=@+s_Atraso", "Si se carga el mes real de arranque en Datos_Reales, prevalece el real.")
    drv("rampa", "E02", "Meses de rampa de aprendizaje", "meses", [0, E, 3, 6, E], "Desde arranque", "Ilustrativo", "e_Rampa", NF_INT0)
    drv("eff0", "E03", "Eficiencia del primer mes de rampa (% del ritmo del plan)", "%", [1, E, 0.6, 0.5, E], "Rampa", "Ilustrativo", "e_Eff0", NF_PCT)
    drv("recup", "E04", "% del volumen no producido (atraso/rampa) recuperable luego", "%", [0, E, 0, 0.5, E], "Horizonte", "Ilustrativo",
        "e_Recup", NF_PCT, note="0% = pérdida definitiva; >0% = desplazamiento recuperable sujeto a holgura de capacidad.")
    SUP["volg_first"] = r
    for y in range(1, 11):
        drv(f"volg{y}", f"E{4 + y:02d}", f"Volumen gabinetes año operativo {y} (% del plan PDF)", "%",
            [1, E, 0.9, 0.8 if y == 1 else (0.9 if y == 2 else 1), E], f"Año op. {y}", "Plan PDF = 100%", None, NF_PCT, "=@*(1+s_Vol)")
    SUP["volg_last"] = r - 1
    SUP["volc_first"] = r
    for y in range(1, 11):
        drv(f"volc{y}", f"E{14 + y:02d}", f"Volumen cajas año operativo {y} (% del plan PDF)", "%",
            [1, E, 0.9, 0.8 if y == 1 else (0.9 if y == 2 else 1), E], f"Año op. {y}", "Plan PDF = 100%", None, NF_PCT, "=@*(1+s_Vol)")
    SUP["volc_last"] = r - 1
    add_range_name("e_VolG", "SUP", f"$J${SUP['volg_first']}:$J${SUP['volg_last']}")
    add_range_name("e_VolC", "SUP", f"$J${SUP['volc_first']}:$J${SUP['volc_last']}")
    drv("capadic", "E25", "Capacidad adicional sobre el plan (turnos / horas extra)", "%", [0, E, 0, 0, E], "Horizonte", "—", "e_CapAdic", NF_PCT,
        note="Requiere respaldo técnico (equipos, turnos, horas y cuellos de botella comunes).")
    drv("precio_g", "E26", "Precio gabinete año 1", "USD/u", [BPs("precio_g"), E, E, E, E], "Año 1", "PDF p. 16", "e_PrecioG1", NF_USD2)
    drv("precio_c", "E27", "Precio caja año 1", "USD/u", [BPs("precio_c"), E, E, E, E], "Año 1", "PDF p. 16", "e_PrecioC1", NF_USD2)
    drv("precio_var", "E28", "Variación de precio sobre la base", "%", [0, E, -0.05, 0, E], "Horizonte", "Ilustrativo", "e_PrecioVar", NF_PCT, "=(1+@)*(1+s_Precio)-1")
    drv("precio_gr", "E29", "Incremento anual de precios desde año 2 (año de proyecto)", "%", [BPs("g_precio"), E, E, E, E], "Desde año 2", "PDF p. 2", "e_PrecioGr", NF_PCT)
    drv("mp_g", "E30", "Materia prima gabinete año 1", "USD/u", [BPs("mp_g"), E, E, E, E], "Año 1", "PDF p. 17", "e_MPg", NF_USD2)
    drv("mp_c", "E31", "Materia prima caja año 1", "USD/u", [BPs("mp_c"), E, E, E, E], "Año 1", "PDF p. 17", "e_MPc", NF_USD2)
    drv("mp_gr", "E32", "Incremento anual de materia prima desde año 2", "%", [BPs("g_mp"), E, E, E, E], "Desde año 2", "PDF p. 2", "e_MPgr", NF_PCT)
    drv("mp_var", "E33", "Variación del costo de materia prima (acero)", "%", [0, E, 0.10, 0, E], "Horizonte", "Ilustrativo", "e_MPvar", NF_PCT, "=(1+@)*(1+s_Mat)-1")
    drv("ins_g", "E34", "Insumos por gabinete", "USD/u", [BPs("ins_g"), E, E, E, E], "Año 1", "PDF p. 18", "e_INSg", NF_USD4)
    drv("ins_c", "E35", "Insumos por caja", "USD/u", [BPs("ins_c"), E, E, E, E], "Año 1", "PDF p. 18 (4,851)", "e_INSc", NF_USD4)
    drv("con_g", "E36", "Consumibles por gabinete (inferido)", "USD/u", [BPs("con_g"), E, E, E, E], "Año 1", "PDF p. 18 (inferido)", "e_CONg", NF_USD4)
    drv("con_c", "E37", "Consumibles por caja", "USD/u", [BPs("con_c"), E, E, E, E], "Año 1", "PDF p. 18", "e_CONc", NF_USD4)
    drv("ins_gr", "E38", "Incremento anual de insumos y consumibles", "%", [0, E, E, E, E], "Desde año 2", "PDF: constantes", "e_INSgr", NF_PCT)
    drv("ins_var", "E39", "Variación del costo de insumos y consumibles", "%", [0, E, 0.10, 0, E], "Horizonte", "Ilustrativo", "e_INSvar", NF_PCT, "=(1+@)*(1+s_Mat)-1")
    drv("fle_g", "E40", "Flete MP por gabinete (inferido)", "USD/u", [f"=Base_PDF!$E${BP['fle_a']}", E, E, E, E], "Horizonte", "PDF cuadro 10a (inferido exacto)", "e_FLEg", NF_USD4)
    drv("fle_c", "E41", "Flete MP por caja (inferido)", "USD/u", [f"=Base_PDF!$E${BP['fle_b']}", E, E, E, E], "Horizonte", "PDF cuadro 10a (inferido exacto)", "e_FLEc", NF_USD4)
    drv("fle_var", "E42", "Variación del flete de MP", "%", [0, E, 0.10, 0, E], "Horizonte", "Ilustrativo", "e_FLEvar", NF_PCT)
    drv("rech", "E43", "Rechazo (% de unidades iniciadas)", "%", [0, E, 0.02, 0.02, E], "Horizonte", "PDF implícito 0%; ilustrativo", "e_Rech", NF_PCT,
        note="Si el precio por pieza de MP ya incluye merma, no aplicar rechazo sobre la misma pérdida.")
    drv("chat", "E44", "Valor de chatarra por unidad rechazada", "USD/u", [0, E, E, E, E], "Horizonte", "No informado", "e_Chat", NF_USD2)
    drv("fact_dot", "E45", "Factor de dotación directa (productividad; 100% = plan)", "%", [1, E, 1.1, 1, E], "Horizonte", "Ilustrativo", "e_FactDot", NF_PCT)
    drv("aj_sal", "E46", "Ajuste salarial anual en Gs (desde año 2)", "%", [0, E, 0.03, 0, E], "Desde año 2", "PDF: constantes; ilustrativo (no dato de mercado)", "e_AjSal", NF_PCT)
    drv("infl_fij", "E47", "Inflación anual de costos fijos (desde año 2)", "%", [0, E, 0.02, 0, E], "Desde año 2", "PDF: constantes; ilustrativo", "e_InflFij", NF_PCT)
    SUP["tc_first"] = r
    for y in range(0, 14):
        drv(f"tc{y}", f"E{48 + y}", f"Tipo de cambio año de proyecto {y}", "Gs/USD", ["=p_TC_Plan", E, E, E, E], f"Año {y}",
            "PDF 7.300 (histórico); cargar TC real/proyectado", None, NF_INT, "=@*(1+s_TC)")
    SUP["tc_last"] = r - 1
    add_range_name("e_TC", "SUP", f"$J${SUP['tc_first']}:$J${SUP['tc_last']}")
    drv("cred", "E62", "Ventas a crédito", "% ventas", [1, E, E, E, E], "Horizonte", "PDF p. 3", "e_Cred", NF_PCT)
    drv("dso", "E63", "Plazo de cobro de ventas a crédito (DSO)", "días", [BPs("d_cxc"), E, 150, 90, E], "Horizonte", "PDF p. 8; ilustrativo", "e_DSO", NF_INT,
        '=IF(s_DSO="",@,s_DSO)')
    drv("cont", "E64", "Ventas al contado", "% ventas", [0, E, E, E, E], "Horizonte", "—", "e_Cont", NF_PCT)
    drv("ant", "E65", "Anticipos de clientes", "% ventas", [0, E, E, E, E], "Horizonte", "—", "e_Ant", NF_PCT)
    drv("ant_m", "E66", "Meses de anticipación del anticipo de clientes", "meses", [1, E, E, E, E], "Horizonte", "—", "e_AntMeses", NF_INT0)
    drv("incob", "E67", "Incobrables", "% ventas a crédito", [0, E, E, E, E], "Horizonte", "—", "e_Incob", NF_PCT)
    drv("d_stock", "E68", "Stock de materiales (días de consumo futuro)", "días", [BPs("d_stock"), E, E, 90, E], "Horizonte", "PDF p. 8", "e_DiasStock", NF_INT)
    drv("d_pt", "E69", "Stock de producto terminado (días de ventas siguientes)", "días", [BPs("d_pt"), E, E, E, E], "Horizonte", "PDF p. 8", "e_DiasPT", NF_INT)
    drv("dpo_mp", "E70", "Plazo de pago a proveedores de materiales (DPO)", "días", [BPs("d_pmp"), E, E, 30, E], "Horizonte", "PDF p. 24", "e_DPOmp", NF_INT)
    drv("ant_prov", "E71", "Anticipo a proveedores de materiales", "% compras", [0, E, E, E, E], "Horizonte", "—", "e_AntProv", NF_PCT)
    drv("ant_prov_m", "E72", "Meses de anticipación del pago a proveedores", "meses", [2, E, E, E, E], "Horizonte", "—", "e_AntProvMeses", NF_INT0)
    drv("dpo_men", "E73", "Plazo de pago a proveedores menores (DPO)", "días", [BPs("d_pmen"), E, E, E, E], "Horizonte", "PDF p. 8", "e_DPOmen", NF_INT)
    drv("d_antp", "E74", "Anticipos al personal (días de remuneración)", "días", [BPs("d_sue"), E, E, E, E], "Horizonte", "PDF p. 8", "e_DiasAntPers", NF_INT)
    drv("sobrecosto", "E75", "Sobrecosto sobre CAPEX pendiente (por contratar)", "%", [0, E, 0.10, 0.10, E], "Implantación", "Ilustrativo", "e_Sobrecosto", NF_PCT, "=(1+@)*(1+s_Capex)-1")
    drv("capex2", "E76", "CAPEX adicional del año 2", "USD", [BPs("capex2"), E, E, E, E], "Año 2", "PDF cuadros 18-19", "e_Capex2", NF_USD)
    drv("capex2_mes", "E77", "Mes de pago del CAPEX adicional", "mes modelo", ["=p_ImplMeses+13", E, E, E, E], "Año 2", "Supuesto provisional (inicio año 2)", "e_Capex2Mes", NF_INT0)
    drv("iva_rec", "E78", "IVA soportado recuperable", "%", [0, E, E, E, E], "Horizonte", "PDF: IVA como costo (prudencial)", "e_IVArec", NF_PCT,
        note="Escenario conservador de no recuperación = criterio histórico, no certeza jurídica.")
    drv("iva_lag", "E79", "Plazo de recuperación del IVA", "meses", [6, E, E, E, E], "Horizonte", "Supuesto provisional", "e_IVAlag", NF_INT0)
    drv("ap_pdf", "E80", "Incluir aportes adicionales del plan (567.782 y 371.928): 1 = sí", "1/0", [1, E, E, E, E], "Años 1-2", "PDF cuadro 18", "e_AportesPDF", NF_INT0)
    drv("l2_monto", "E81", "Nuevo préstamo (tramo 2): monto", "USD", [0, E, E, E, E], "—", "Sin contrato", "e_L2_Monto", NF_USD)
    drv("l2_mes", "E82", "Tramo 2: mes de desembolso", "mes modelo", [12, E, E, E, E], "—", "—", "e_L2_Mes", NF_INT0)
    drv("l2_tasa", "E83", "Tramo 2: tasa nominal anual", "%", [BPs("kd"), E, E, E, E], "—", "Histórica PDF", "e_L2_Tasa", NF_PCT2)
    drv("l2_gracia", "E84", "Tramo 2: meses de gracia", "meses", [0, E, E, E, E], "—", "—", "e_L2_Gracia", NF_INT0)
    drv("l2_cuotas", "E85", "Tramo 2: cuotas francesas", "meses", [36, E, E, E, E], "—", "—", "e_L2_Cuotas", NF_INT0)
    drv("linea", "E86", "Línea de crédito hipotética automática (1 = activar) — NO APROBADA", "1/0", [0, E, E, E, E], "—", "Hipotética", "e_Linea", NF_INT0,
        note="Sin circularidad: interés sobre saldo inicial del mes. Estado: no aprobada.")
    drv("cupo", "E87", "Cupo de la línea hipotética", "USD", [0, E, E, E, E], "—", "Hipotética", "e_Cupo", NF_USD)
    drv("tasa_linea", "E88", "Tasa nominal anual de la línea hipotética", "%", [BPs("kd"), E, E, E, E], "—", "Hipotética", "e_TasaLinea", NF_PCT2)
    drv("gv", "E89", "Gastos de venta y logística de exportación (flete, seguros, comisiones, garantías)", "% ventas", [0, E, E, E, E], "Horizonte",
        "PDF: cero; pendiente confirmar quién los asume", "e_GV", NF_PCT)
    drv("ene_var", "E90", "Energía: porción variable con kg procesados", "%", [0, E, E, E, E], "Horizonte", "PDF: constante", "e_EneVar", NF_PCT)
    SUP["scen_last"] = r - 1
    r += 1

    # ---------------------------------------------------------------- E choques
    section(ws, r, "E. Pruebas de sensibilidad (choques ilustrativos aplicados sobre el escenario activo)", 13); r += 1
    header(ws, r, ["N°", "Prueba", "", "CAPEX pendiente", "Atraso (meses)", "Volumen", "Precio", "Materiales", "TC Gs/USD", "DSO (días)"]); r += 1
    SUP["shk_first"] = r
    for i, (nm, cx, at, vo, pr, ma, tc, dso) in enumerate(SHOCKS):
        setc(ws, r, 1, i, fmt=NF_INT0)
        setc(ws, r, 2, nm)
        for j, v in enumerate([cx, at, vo, pr, ma, tc]):
            setc(ws, r, 4 + j, v, font=F_IN, fill=FILL_IN, fmt=NF_INT0 if j == 1 else NF_PCT)
        setc(ws, r, 10, dso, font=F_IN, fill=FILL_IN, fmt=NF_INT0)
        r += 1
    SUP["shk_last"] = r - 1
    f1, l1 = SUP["shk_first"], SUP["shk_last"]
    setc(ws, r, 2, "Choque activo", font=F_BOLD)
    names = ["s_Capex", "s_Atraso", "s_Vol", "s_Precio", "s_Mat", "s_TC"]
    for j, nm in enumerate(names):
        cl = CL(4 + j)
        setc(ws, r, 4 + j, f"=INDEX({cl}{f1}:{cl}{l1},p_ShockIdx)", fmt=NF_INT0 if j == 1 else NF_PCT, font=F_BOLD)
        add_name(nm, "SUP", f"{cl}{r}")
    setc(ws, r, 10, f'=IF(INDEX(J{f1}:J{l1},p_ShockIdx)="","",INDEX(J{f1}:J{l1},p_ShockIdx))', fmt=NF_INT0, font=F_BOLD)
    add_name("s_DSO", "SUP", f"J{r}")
    SUP["shk_active"] = r
    r += 2

    # ---------------------------------------------------------------- F dividendos / aportes planificados
    section(ws, r, "F. Política de dividendos (% de utilidad del año de proyecto) y aportes planificados", 13); r += 1
    header(ws, r, ["Año", "Payout (% de la utilidad del año)", "", "Valor", "Fuente"]); r += 1
    SUP["pay_first"] = r
    for y in range(1, NYEARS + 1):
        setc(ws, r, 1, y, fmt=NF_INT0)
        setc(ws, r, 2, f"Año de proyecto {y}")
        v = f"=Base_PDF!{CL(BPY0 + y)}${BP['payout_pub']}" if y <= 10 else 1
        setc(ws, r, 4, v, font=F_LINK if y <= 10 else F_IN, fill=None if y <= 10 else FILL_IN, fmt=NF_PCT)
        setc(ws, r, 5, "PDF cuadro 15" if y <= 10 else "Supuesto: igual al año 10")
        r += 1
    SUP["pay_last"] = r - 1
    add_range_name("p_Payout", "SUP", f"$D${SUP['pay_first']}:$D${SUP['pay_last']}")
    r += 1
    header(ws, r, ["N°", "Aporte planificado adicional (no confirmado)", "", "Mes modelo", "Monto USD", "Estado"]); r += 1
    SUP["ap_first"] = r
    for i in range(5):
        setc(ws, r, 1, i + 1, fmt=NF_INT0)
        setc(ws, r, 2, "Aporte de socios planificado")
        setc(ws, r, 4, None, font=F_IN, fill=FILL_IN, fmt=NF_INT0)
        setc(ws, r, 5, None, font=F_IN, fill=FILL_IN, fmt=NF_USD)
        setc(ws, r, 6, "Vacío = no hay compromiso")
        r += 1
    SUP["ap_last"] = r - 1
    add_range_name("p_ApMes", "SUP", f"$D${SUP['ap_first']}:$D${SUP['ap_last']}")
    add_range_name("p_ApMonto", "SUP", f"$E${SUP['ap_first']}:$E${SUP['ap_last']}")
    r += 1

    # ---------------------------------------------------------------- G derivados
    section(ws, r, "G. Parámetros derivados (cálculo; no editar)", 13); r += 1

    def dp(key, label, formula, name, fmt, note=None):
        nonlocal r
        SUP[key] = r
        setc(ws, r, 2, label)
        setc(ws, r, 4, formula, fmt=fmt)
        if note:
            setc(ws, r, 8, note)
        add_name(name, "SUP", f"D{r}")
        r += 1
    dp("corte", "Último mes cerrado con datos reales (-1 = ninguno)", f'=IF(Datos_Reales!$D${DR_T1["r_Corte"]}="",-1,Datos_Reales!$D${DR_T1["r_Corte"]})', "p_MesCorte", NF_INT0)
    dp("arr_real", "Mes real de arranque cargado (vacío = no informado)", f'=IF(Datos_Reales!$D${DR_T1["r_Arranque"]}="","",Datos_Reales!$D${DR_T1["r_Arranque"]})', "p_ArranqueReal", NF_INT0)
    dp("arranque", "Primer mes operativo efectivo (mes modelo)", '=IF(p_ArranqueReal<>"",p_ArranqueReal,p_ImplMeses+MAX(0,ROUND(e_Atraso,0))+1)', "p_MesArranque", NF_INT0)
    dp("mesfin", "Último mes del horizonte evaluado", f"=MIN({N},IF(p_ModoHor=1,p_ImplMeses+p_MesesOp,p_MesArranque+p_MesesOp-1))", "p_MesFin", NF_INT0,
       f"Tope técnico de la rejilla: mes {N}.")
    dp("mesfin_teor", "Último mes teórico sin tope de rejilla", "=IF(p_ModoHor=1,p_ImplMeses+p_MesesOp,p_MesArranque+p_MesesOp-1)", "p_MesFinTeor", NF_INT0)
    dp("lcob", "Plazo de cobro en meses", "=e_DSO/p_DiasMes", "p_Lcob", NF_USD4)
    dp("lcob_i", "Plazo de cobro: meses enteros", "=INT(p_Lcob)", "p_LcobI", NF_INT0)
    dp("lcob_f", "Plazo de cobro: fracción", "=p_Lcob-p_LcobI", "p_LcobF", NF_USD4)
    dp("lmp", "Plazo de pago materiales en meses", "=e_DPOmp/p_DiasMes", "p_Lmp", NF_USD4)
    dp("lmp_i", "Pago materiales: meses enteros", "=INT(p_Lmp)", "p_LmpI", NF_INT0)
    dp("lmp_f", "Pago materiales: fracción", "=p_Lmp-p_LmpI", "p_LmpF", NF_USD4)
    dp("lmen", "Plazo de pago proveedores menores en meses", "=e_DPOmen/p_DiasMes", "p_Lmen", NF_USD4)
    dp("lmen_i", "Pago menores: meses enteros", "=INT(p_Lmen)", "p_LmenI", NF_INT0)
    dp("lmen_f", "Pago menores: fracción", "=p_Lmen-p_LmenI", "p_LmenF", NF_USD4)
    dp("sk", "Stock de materiales: meses enteros de consumo futuro", "=INT(e_DiasStock/p_DiasMes)", "p_Sk", NF_INT0)
    dp("sf", "Stock de materiales: fracción del mes siguiente", "=e_DiasStock/p_DiasMes-p_Sk", "p_Sf", NF_USD4)
    dp("ptd", "Stock de PT: fracción de las ventas del mes siguiente", "=e_DiasPT/p_DiasMes", "p_PTf", NF_USD4)
    dp("kgbase", "Kg procesados por mes en el plan del año 1 (base de energía variable)",
       f"=(Base_PDF!$F${BP['prod_g']}*p_KgG+Base_PDF!$F${BP['prod_c']}*p_KgC)/12", "p_KgBaseMes", NF_USD)
    dp("wg", "Peso del driver de costos comunes: gabinete", "=IF(p_Driver=3,IF(p_HsG=\"\",p_KgG,p_HsG),IF(p_Driver=2,e_MPg+e_INSg+e_CONg+e_FLEg,p_KgG))", "p_WG", NF_USD4,
       "Driver 3 sin horas cargadas vuelve a kg. Driver 2 = costo directo unitario año 1 (criterio PDF).")
    dp("wc", "Peso del driver de costos comunes: caja", "=IF(p_Driver=3,IF(p_HsC=\"\",p_KgC,p_HsC),IF(p_Driver=2,e_MPc+e_INSc+e_CONc+e_FLEc,p_KgC))", "p_WC", NF_USD4)
    dp("cobro_sum", "Control: crédito + contado + anticipo (debe ser 100%)", "=e_Cred+e_Cont+e_Ant", "p_MixCobro", NF_PCT)
    SUP["_last"] = r

    # selector / estado
    ws.cell(row=SUP["esc_idx"], column=4).value = f"=MATCH(p_Escenario,$D${SUP['scen_hdr']}:$H${SUP['scen_hdr']},0)"
    add_name("p_EscIdx", "SUP", f"D{SUP['esc_idx']}")
    ws.cell(row=SUP["esc_status"], column=4).value = f"=INDEX($D${SUP['scen_status']}:$H${SUP['scen_status']},1,p_EscIdx)"
    ws.cell(row=SUP["esc_status"], column=4).font = F_ALERT
    ws.merge_cells(start_row=SUP["esc_status"], start_column=4, end_row=SUP["esc_status"], end_column=12)
    ws.cell(row=SUP["shock_idx"], column=4).value = f"=MATCH(p_Shock,$B${SUP['shk_first']}:$B${SUP['shk_last']},0)"
    add_name("p_ShockIdx", "SUP", f"D{SUP['shock_idx']}")
    dv_range(ws, f"D{SUP['esc']}", f"=$D${SUP['scen_hdr']}:$H${SUP['scen_hdr']}")
    dv_range(ws, f"D{SUP['shock']}", f"=$B${SUP['shk_first']}:$B${SUP['shk_last']}")
    dv_list(ws, f"C{SUP['corr_first']}:C{SUP['corr_last']}", ["0", "1"])
    ws.freeze_panes = "C4"


build_supuestos()


# ===========================================================================
# 3) TABLAS DE DATOS: nómina y maestro de costos (Costos_Personal),
#    ítems de CAPEX (Inversion_Activos) y Datos_Reales
# ===========================================================================
CPT = {}   # referencias de tablas en Costos_Personal
IAT = {}   # referencias de la tabla de CAPEX
DRT = {}   # referencias de Datos_Reales

CLASS_DEP = {"I01": "Administración", "I12": "Administración", "I13": "Administración",
             "I18": "Administración", "I19": "Administración"}

COST_MASTER = [
    # código, concepto, clasificación, comportamiento, moneda, importe anual origen (fórmula), IVA %, local, fuente, estado
    ("ENE", "Energía eléctrica planta", "Producción", "Fijo (porción variable editable)", "PYG", "ene_gs", 0, 1,
     "PDF 10a: USD 23.014 ≈ Gs 168 M a 7.300", "Inferido - confirmar moneda y tarifa"),
    ("MAN", "Mantenimiento y reparaciones planta", "Producción", "Fijo creciente (lineal)", "USD", "p_MAN", 0, 1,
     "PDF 10a año 1", "Publicado (PDF)"),
    ("SRV", "Agua, comunicaciones y electricidad (adm.)", "Administración", "Fijo", "PYG", "srv_gs", 0, 1,
     "PDF 10a: USD 1.233 ≈ Gs 9 M", "Inferido - confirmar"),
    ("MAA", "Mantenimiento y reparaciones (adm.)", "Administración", "Fijo", "PYG", "maa_gs", 0, 1,
     "PDF 10a: USD 4.932 ≈ Gs 36 M", "Inferido - confirmar"),
    ("SGA", "Seguros sobre activos fijos", "Administración", "Fijo", "PYG", "sga_gs", 0, 1,
     "PDF 10a: USD 8.219 ≈ Gs 60 M", "Inferido - confirmar"),
    ("GG", "Gastos generales", "Administración", "Fijo", "USD", "p_GG", 0, 1, "PDF 10a año 1", "Publicado (PDF)"),
    ("HON", "Honorarios y otros fijos", "Administración", "Fijo", "PYG", "hon_gs", 0, 1,
     "PDF 10a: USD 24.932 ≈ Gs 182 M", "Inferido - confirmar"),
]
LOCAL_FLAGS = [
    ("CON", "Insumos consumibles", 0), ("DES", "Despachos de importación (servicio local)", 1),
    ("PKG", "Packaging", 1), ("SEGC", "Seguro de caución", 1), ("ALQ", "Alquiler de nave", 1), ("EXP", "Expensas", 1),
]


def build_cp_tables():
    ws = WS["CP"]
    CUR[0] = "CP"
    title(ws, "Costos_Personal — nómina por puesto, costos fijos, materiales y motor mensual de costos (USD salvo indicación)",
          "Tablas superiores: plan de dotación y maestro de costos (editables). Bloque inferior: motor mensual (no editar fórmulas).")
    for k, v in {1: 8, 2: 40, 3: 14, 4: 18, 5: 16}.items():
        ws.column_dimensions[CL(k)].width = v
    r = 4
    section(ws, r, "A. Dotación y salarios por puesto (plan vigente; por defecto = PDF pp. 30-34)", 18); r += 1
    header(ws, r, ["Cód.", "Puesto", "Clase PDF", "Clase depurada (editable)", "Clase efectiva", "Salario Gs/mes"] +
           [f"Dot. año {y}" for y in range(1, 11)] + ["Fuente"]); r += 1
    CPT["pos_first"] = r
    for i, (code, name, cls, gs, dots, usdm, usda) in enumerate(BP["nom_data"]):
        br = BP["nom_first"] + i
        setc(ws, r, 1, code)
        setc(ws, r, 2, f"=Base_PDF!$B${br}", font=F_LINK)
        setc(ws, r, 3, f"=Base_PDF!$C${br}", font=F_LINK)
        dep = "Directo" if cls == "Directo" else CLASS_DEP.get(code, "Indirecto fábrica")
        setc(ws, r, 4, dep, font=F_IN, fill=FILL_IN)
        setc(ws, r, 5, f'=IF(p_C6=1,D{r},IF(C{r}="Directo","Directo","Administración"))')
        setc(ws, r, 6, f"=Base_PDF!$E${br}", font=F_LINK, fmt=NF_GS)
        for y in range(1, 11):
            src = f"=Base_PDF!{CL(5 + min(y, 5))}${br}" if y <= 5 else f"={CL(6 + 5)}{r}"
            setc(ws, r, 6 + y, src, font=F_LINK if y <= 5 else F_BASE, fmt=NF_INT)
        setc(ws, r, 17, "PDF pp. 30-34; años 6-10 = año 5 (plantilla invariable)")
        r += 1
    CPT["pos_last"] = r - 1
    pf, pl = CPT["pos_first"], CPT["pos_last"]
    dv_list(ws, f"D{pf}:D{pl}", ["Directo", "Indirecto fábrica", "Administración"])
    setc(ws, r, 2, "Clasificación depurada propuesta: supervisión, ingeniería, calidad, PCP, depósito y mantenimiento = indirectos de fábrica; "
                   "gerencia general, administración, compras y logística/comex = administración. Confirmar.", font=F_SUB)
    r += 2
    header(ws, r, ["", "Resumen por clase efectiva y año operativo", "", "", "", ""] + [f"Año {y}" for y in range(1, 11)]); r += 1
    for key, label, cls, kind in [("dot_d", "Dotación directa (plan)", "Directo", "dot"),
                                  ("dot_i", "Dotación indirecta de fábrica", "Indirecto fábrica", "dot"),
                                  ("dot_a", "Dotación administrativa", "Administración", "dot"),
                                  ("masa_d", "Masa salarial directa (Gs/mes, plan)", "Directo", "masa"),
                                  ("masa_i", "Masa salarial indirecta fábrica (Gs/mes)", "Indirecto fábrica", "masa"),
                                  ("masa_a", "Masa salarial administrativa (Gs/mes)", "Administración", "masa")]:
        CPT[key] = r
        setc(ws, r, 2, label)
        for y in range(1, 11):
            c = CL(6 + y)
            if kind == "dot":
                f = f'=SUMIF($E${pf}:$E${pl},"{cls}",{c}${pf}:{c}${pl})'
                fmt = NF_INT
            else:
                f = f'=SUMPRODUCT(($E${pf}:$E${pl}="{cls}")*$F${pf}:$F${pl}*{c}${pf}:{c}${pl})'
                fmt = NF_GS
            setc(ws, r, 6 + y, f, fmt=fmt)
        r += 1
    CPT["dot_d_aj"] = r
    setc(ws, r, 2, "Dotación directa ajustada por productividad (factor del escenario)")
    for y in range(1, 11):
        c = CL(6 + y)
        setc(ws, r, 6 + y, f"=ROUND({c}{CPT['dot_d']}*e_FactDot,0)", fmt=NF_INT)
    r += 1
    CPT["masa_d_aj"] = r
    setc(ws, r, 2, "Masa salarial directa ajustada (Gs/mes; mantiene el mix de puestos)")
    for y in range(1, 11):
        c = CL(6 + y)
        setc(ws, r, 6 + y, f"=IF({c}{CPT['dot_d']}>0,{c}{CPT['masa_d']}*{c}{CPT['dot_d_aj']}/{c}{CPT['dot_d']},0)", fmt=NF_GS)
    r += 2

    section(ws, r, "B. Maestro de costos fijos por rubro (importe anual del año 1 en moneda de origen)", 18); r += 1
    header(ws, r, ["Cód.", "Concepto", "Clasificación", "Comportamiento", "Moneda", "Importe anual origen", "IVA incluido %",
                   "Origen local (VA nac.)", "USD año 1 a TC plan", "Fuente", "", "", "Estado"]); r += 1
    CPT["cm_first"] = r
    for code, name, cls, beh, cur, src, iva, loc, fsrc, st in COST_MASTER:
        CPT["cm_" + code] = r
        setc(ws, r, 1, code)
        setc(ws, r, 2, name)
        setc(ws, r, 3, cls)
        setc(ws, r, 4, beh)
        setc(ws, r, 5, cur, font=F_IN, fill=FILL_IN)
        if src.startswith("p_"):
            v = f"=Base_PDF!$F${BP[src]}"
        else:
            v = f"=Base_PDF!$E${BP[src]}"
        setc(ws, r, 6, v, font=F_LINK, fmt=NF_GS)
        setc(ws, r, 7, iva, font=F_IN, fill=FILL_IN, fmt=NF_PCT)
        setc(ws, r, 8, loc, font=F_IN, fill=FILL_IN, fmt=NF_INT0)
        setc(ws, r, 9, f'=IF(E{r}="PYG",F{r}/p_TC_Plan,F{r})', fmt=NF_USD)
        setc(ws, r, 10, fsrc)
        setc(ws, r, 13, st)
        r += 1
    CPT["cm_last"] = r - 1
    dv_list(ws, f"E{CPT['cm_first']}:E{CPT['cm_last']}", ["PYG", "USD"])
    setc(ws, r, 2, "IVA incluido %: 0 = no identificado en el PDF (pendiente). Solo el alquiler declara IVA de 10% (parámetro G37).", font=F_SUB)
    r += 2
    section(ws, r, "C. Clasificación de origen local para el valor agregado nacional (base alternativa del tributo)", 18); r += 1
    header(ws, r, ["Cód.", "Concepto", "Local (1/0)", "Nota"]); r += 1
    for code, name, flag in LOCAL_FLAGS:
        CPT["lf_" + code] = r
        setc(ws, r, 1, code)
        setc(ws, r, 2, name)
        setc(ws, r, 3, flag, font=F_IN, fill=FILL_IN, fmt=NF_INT0)
        setc(ws, r, 4, "Clasificación provisional. Nómina siempre local; MP, insumos y flete MP se consideran importados.")
        r += 1
    r += 1
    CPT["engine_start"] = r
    return r


def cm(code, col):
    return f"Costos_Personal!${col}${CPT['cm_' + code]}" if CUR[0] != "CP" else f"${col}${CPT['cm_' + code]}"


def build_ia_table():
    ws = WS["IA"]
    CUR[0] = "IA"
    title(ws, "Inversion_Activos — presupuesto, ejecutado, comprometido, por contratar, pagos, altas y depreciación (USD)",
          "Costo final estimado = ejecutado devengado + comprometido pendiente + estimado por contratar (estados excluyentes). "
          "Los datos reales por ítem se cargan en Datos_Reales (tabla T3).")
    widths = {1: 7, 2: 40, 3: 18, 4: 20}
    for k, v in widths.items():
        ws.column_dimensions[CL(k)].width = v
    r = 4
    section(ws, r, "A. Ítems de inversión", 26); r += 1
    hdrs = ["Cód.", "Ítem", "Clase", "Tipo", "Base / FOB USD", "Seguro + flete", "Despacho asignado", "Presupuesto original",
            "Ejecutado devengado (real)", "Pagado (real)", "Devengado no pagado", "Comprometido pendiente (real)",
            "Por contratar revisado (real)", "Por contratar estimado", "Costo final estimado", "Desviación USD", "Desviación %",
            "Avance financiero %", "Mes de pago previsto", "Mes de pago efectivo", "Mes de puesta en servicio", "Vida útil (años)",
            "Valor residual %", "Pendiente de pago", "Pendiente de devengar", "Fuente / estado"]
    header(ws, r, hdrs)
    ws.row_dimensions[r].height = 42
    r += 1
    items = [("INF", "Adecuación de infraestructura", "Infraestructura", "Inicial", f"=Base_PDF!$E${BP['inf']}", None, 10,
              "PDF cuadro 5 - Publicado")]
    for code, name, cant, fob, cifp in [(c[0], c[1], c[2], c[3], c[4]) for c in
                                        [("M01", "Conjunto línea de pintura", 0, 0, 0), ("M02", "Cortadoras láser (2)", 0, 0, 0),
                                         ("M03", "Manipulador", 0, 0, 0), ("M04", "Dobladoras (4)", 0, 0, 0),
                                         ("M05", "Soldadoras (6)", 0, 0, 0), ("M06", "Lijadoras (4)", 0, 0, 0),
                                         ("M07", "Conjunto baño", 0, 0, 0), ("M08", "Línea de montaje", 0, 0, 0),
                                         ("M09", "Montacargas", 0, 0, 0), ("M10", "Soldadoras capacitivas (2)", 0, 0, 0)]]:
        items.append((code, name, "Maquinaria y equipos", "Inicial", f"=Base_PDF!$E${BP['maq_' + code]}", "maq", 10,
                      "PDF cuadro 2 - Publicado (FOB total del renglón)"))
    items += [("SOF", "Software de gestión del negocio", "Software (intangible)", "Inicial", f"=Base_PDF!$E${BP['sof']}", None, 5, "PDF cuadro 6 - Publicado"),
              ("MOB", "Mobiliario", "Mobiliario", "Inicial", f"=Base_PDF!$E${BP['mob']}", None, 5, "PDF cuadro 6 - Publicado"),
              ("ADD", "CAPEX adicional año 2 (naturaleza no desagregada)", "Pendiente de definir", "Expansión", "=e_Capex2", None, 10,
               "PDF cuadros 18-19. Vida útil 10 = SUPUESTO PROVISIONAL"),
              ("N01", "Nuevo ítem / reposición 1", "", "Mantenimiento/Reposición", None, None, None, "Vacío"),
              ("N02", "Nuevo ítem / reposición 2", "", "Mantenimiento/Reposición", None, None, None, "Vacío"),
              ("N03", "Nuevo ítem / expansión 3", "", "Expansión", None, None, None, "Vacío")]
    IAT["first"] = r
    IAT["codes"] = []
    for i, (code, name, cls, tipo, base, kind, life, src) in enumerate(items):
        IAT["codes"].append(code)
        IAT[code] = r
        drr = 31 + i   # fila en Datos_Reales T3
        setc(ws, r, 1, code)
        setc(ws, r, 2, name)
        setc(ws, r, 3, cls, font=F_IN if code.startswith("N") else F_BASE, fill=FILL_IN if code.startswith("N") else None)
        setc(ws, r, 4, tipo, font=F_IN, fill=FILL_IN)
        setc(ws, r, 5, base, font=F_LINK if (base and "Base_PDF" in str(base)) else (F_IN if base is None else F_BASE),
             fill=FILL_IN if base is None else None, fmt=NF_USD)
        if kind == "maq":
            setc(ws, r, 6, f"=Base_PDF!$F${BP['maq_' + code]}", font=F_LINK, fmt=NF_USD)
            setc(ws, r, 7, f"=(E{r}+F{r})*Base_PDF!$F${BP['desp_maq']}", fmt=NF_USD)
        else:
            setc(ws, r, 6, 0, fmt=NF_USD)
            setc(ws, r, 7, 0, fmt=NF_USD)
        setc(ws, r, 8, f"=N(E{r})+F{r}+G{r}", fmt=NF_USD, font=F_BOLD)
        setc(ws, r, 9, f'=IF(Datos_Reales!$C${drr}="",0,Datos_Reales!$C${drr})', font=F_LINK, fmt=NF_USD)
        setc(ws, r, 10, f'=IF(Datos_Reales!$D${drr}="",0,Datos_Reales!$D${drr})', font=F_LINK, fmt=NF_USD)
        setc(ws, r, 11, f"=I{r}-J{r}", fmt=NF_USD)
        setc(ws, r, 12, f'=IF(Datos_Reales!$E${drr}="",0,Datos_Reales!$E${drr})', font=F_LINK, fmt=NF_USD)
        setc(ws, r, 13, f'=IF(Datos_Reales!$F${drr}="","",Datos_Reales!$F${drr})', font=F_LINK, fmt=NF_USD)
        setc(ws, r, 14, f'=IF(M{r}<>"",N(M{r}),MAX(0,H{r}-I{r}-L{r})*(1+e_Sobrecosto))', fmt=NF_USD)
        setc(ws, r, 15, f"=I{r}+L{r}+N{r}", fmt=NF_USD, font=F_BOLD)
        setc(ws, r, 16, f"=O{r}-H{r}", fmt=NF_USD)
        setc(ws, r, 17, f'=IF(H{r}=0,"",P{r}/H{r})', fmt=NF_PCT)
        setc(ws, r, 18, f'=IF(O{r}=0,"",J{r}/O{r})', fmt=NF_PCT)
        pay = "=e_Capex2Mes" if code == "ADD" else (0 if not code.startswith("N") else None)
        setc(ws, r, 19, pay, font=F_IN if code != "ADD" else F_BASE, fill=FILL_IN if code != "ADD" else None, fmt=NF_INT0)
        setc(ws, r, 20, f"=MAX(N(S{r}),p_MesCorte+1)", fmt=NF_INT0)
        if code.startswith("N"):
            svc = f'=IF(Datos_Reales!$G${drr}<>"",Datos_Reales!$G${drr},T{r})'
        elif code == "ADD":
            svc = f'=IF(Datos_Reales!$G${drr}<>"",Datos_Reales!$G${drr},IF(p_C4=1,T{r},9999))'
        else:
            svc = f'=IF(Datos_Reales!$G${drr}<>"",Datos_Reales!$G${drr},IF(p_C3=1,p_MesArranque,p_MesArranque+12))'
        setc(ws, r, 21, svc, fmt=NF_INT0)
        setc(ws, r, 22, life, font=F_IN, fill=FILL_IN, fmt=NF_INT0)
        setc(ws, r, 23, 0, font=F_IN, fill=FILL_IN, fmt=NF_PCT)
        setc(ws, r, 24, f"=K{r}+L{r}+N{r}", fmt=NF_USD)
        setc(ws, r, 25, f"=L{r}+N{r}", fmt=NF_USD)
        setc(ws, r, 26, src)
        r += 1
    IAT["last"] = r - 1
    f, l = IAT["first"], IAT["last"]
    IAT["tot"] = r
    setc(ws, r, 2, "TOTAL", font=F_BOLD)
    for c in [5, 6, 7, 8, 9, 10, 11, 12, 14, 15, 16, 24, 25]:
        setc(ws, r, c, f"=SUM({CL(c)}{f}:{CL(c)}{l})", fmt=NF_USD, font=F_BOLD, fill=FILL_TOT)
    setc(ws, r, 17, f'=IF(H{r}=0,"",P{r}/H{r})', fmt=NF_PCT, font=F_BOLD)
    setc(ws, r, 18, f'=IF(O{r}=0,"",J{r}/O{r})', fmt=NF_PCT, font=F_BOLD)
    r += 1
    IAT["tot_ini"] = r
    setc(ws, r, 2, "Subtotal inversión inicial (control vs USD 655.044)")
    setc(ws, r, 8, f'=SUMIF($D${f}:$D${l},"Inicial",H{f}:H{l})', fmt=NF_USD)
    setc(ws, r, 15, f'=SUMIF($D${f}:$D${l},"Inicial",O{f}:O{l})', fmt=NF_USD)
    r += 1
    dv_list(ws, f"D{f}:D{l}", ["Inicial", "Expansión", "Mantenimiento/Reposición"])
    setc(ws, r, 2, "Mes de pago/servicio: mes del modelo. Si el mes previsto ya es real, el pendiente se reprograma al mes siguiente al corte. "
                   "La vida útil y el valor residual son editables; la del ítem ADD es provisional (no informada en el PDF).", font=F_SUB)
    r += 2
    IAT["engine_start"] = r
    return r


def build_datos_reales():
    ws = WS["DR"]
    CUR[0] = "DR"
    title(ws, "Datos_Reales — ejecución real, saldos y documentación (inicialmente vacías: no se suministraron datos posteriores al plan)",
          "Cada registro debe tener fecha/mes, identificador, moneda, importe o cantidad, tipo de dato, fuente y estado. "
          "Un faltante NO equivale a cero: el modelo solo usa reales en meses ≤ último mes cerrado.")
    for k, v in {1: 11, 2: 44, 3: 14, 4: 16, 5: 14, 6: 16, 7: 16, 8: 16, 9: 14, 10: 12, 11: 14, 12: 26, 13: 14, 14: 36}.items():
        ws.column_dimensions[CL(k)].width = v
    r = 4
    section(ws, r, "T1. Fechas del proyecto y corte", 14); r += 1
    header(ws, r, ["Cód.", "Dato", "", "Valor", "", "Instrucción"]); r += 1
    t1 = [("r_FechaImpl", "Fecha real de inicio de implantación", NF_DATE, "Informativo."),
          ("r_Arranque", "Mes del modelo del inicio real de operación", NF_INT0, "Si se carga, prevalece sobre el atraso del escenario."),
          ("r_Corte", "Último mes cerrado con datos reales (mes del modelo)", NF_INT0,
           "No tomar automáticamente el mes en curso. Vacío = sin reales. Los meses ≤ corte quedan congelados."),
          ("r_FechaCorte", "Fecha de corte", NF_DATE, "Informativo; debe coincidir con el mes anterior.")]
    for nm, label, fmt, ins in t1:
        assert DR_T1[nm] == r
        setc(ws, r, 1, nm)
        setc(ws, r, 2, label)
        setc(ws, r, 4, None, font=F_IN, fill=FILL_PEND, fmt=fmt)
        setc(ws, r, 6, ins)
        add_name(nm, "DR", f"D{r}")
        r += 1
    setc(ws, r, 2, "Responsable y fecha de la última carga")
    setc(ws, r, 4, None, font=F_IN, fill=FILL_IN)
    r += 2
    section(ws, r, "T2. Saldos reales al último mes cerrado (se usan solo en el mes de corte; la diferencia con el cálculo va a «ajustes de conciliación»)", 14); r += 1
    header(ws, r, ["Cód.", "Saldo", "Moneda", "Importe origen", "TC Gs/USD", "Importe USD", "Fuente", "Estado revisión"]); r += 1
    saldos = [("S_CAJA", "Caja y bancos"), ("S_CXC", "Cuentas por cobrar a clientes"), ("S_INVMAT", "Inventario de materiales (costo de adquisición)"),
              ("S_INVPT", "Inventario de productos terminados"), ("S_PROVMAT", "Proveedores de materiales"),
              ("S_PROVMEN", "Proveedores menores y obligaciones"), ("S_IVA", "Crédito fiscal IVA recuperable"),
              ("S_DEUDA1", "Saldo préstamo original"), ("S_DEUDA2", "Saldo préstamo tramo 2")]
    for code, label in saldos:
        setc(ws, r, 1, code)
        setc(ws, r, 2, label)
        setc(ws, r, 3, "USD", font=F_IN, fill=FILL_IN)
        setc(ws, r, 4, None, font=F_IN, fill=FILL_IN, fmt=NF_USD)
        setc(ws, r, 5, None, font=F_IN, fill=FILL_IN, fmt=NF_INT)
        setc(ws, r, 6, f'=IF(D{r}="","",IF(C{r}="PYG",IF(N(E{r})=0,"FALTA TC",D{r}/E{r}),D{r}))', fmt=NF_USD)
        setc(ws, r, 7, None, font=F_IN, fill=FILL_IN)
        setc(ws, r, 8, None, font=F_IN, fill=FILL_IN)
        add_name("r_" + code, "DR", f"F{r}")
        r += 1
    dv_list(ws, f"C{r - len(saldos)}:C{r - 1}", ["USD", "PYG"])
    r += 1
    assert r <= 29
    r = 29
    section(ws, r, "T3. CAPEX real por ítem (acumulado al corte, USD al TC de cada operación)", 14); r += 1
    header(ws, r, ["Cód. ítem", "Ítem", "Ejecutado devengado", "Pagado", "Comprometido pendiente de devengar", "Por contratar (estimación revisada)",
                   "Mes real de puesta en uso", "Fuente", "Estado revisión"])
    ws.row_dimensions[r].height = 36
    r += 1
    assert r == 31
    for i, code in enumerate(IAT["codes"]):
        setc(ws, r, 1, f"=Inversion_Activos!$A${IAT[code]}", font=F_LINK)
        setc(ws, r, 2, f"=Inversion_Activos!$B${IAT[code]}", font=F_LINK)
        for c in range(3, 8):
            setc(ws, r, c, None, font=F_IN, fill=FILL_IN, fmt=NF_USD if c < 7 else NF_INT0)
        setc(ws, r, 8, None, font=F_IN, fill=FILL_IN)
        setc(ws, r, 9, None, font=F_IN, fill=FILL_IN)
        r += 1
    DRT["capex_last"] = r - 1
    r += 1
    section(ws, r, "T4. Catálogo de códigos para la tabla de movimientos (T7)", 14); r += 1
    header(ws, r, ["Código", "Descripción", "Unidad", "Tipo de dato", "Destino en el motor"]); r += 1
    DRT["cat_first"] = r
    for code, desc, unit, tipo, dest in CODES:
        setc(ws, r, 1, code, font=F_BOLD)
        setc(ws, r, 2, desc)
        setc(ws, r, 3, unit)
        setc(ws, r, 4, tipo)
        setc(ws, r, 5, dest)
        r += 1
    DRT["cat_last"] = r - 1
    r += 1
    section(ws, r, "T5. Documentación contractual y comercial requerida (estado)", 14); r += 1
    header(ws, r, ["N°", "Documento / definición", "Estado", "Fecha", "Fuente", "Efecto en el modelo"]); r += 1
    DRT["doc_first"] = r
    docs = [("Contrato CIE-PERG / matriz (modalidad: venta de producto o servicio de maquila)", "Modalidad G26, tarifas G27-G28, ingresos"),
            ("Programa de maquila aprobado (MIC) y régimen de transición Ley 7547/2025", "Tributo G42-G43, IVA E78-E79"),
            ("Propiedad de materias primas e insumos (quién compra e importa)", "G29, inventarios y costos propios"),
            ("Condiciones de entrega, flete de exportación, seguros, garantías y rechazos", "E89 gastos de venta; E43 rechazo"),
            ("Naturaleza del alquiler de la nave (tercero, relacionada o costo económico)", "G40 alquiler en caja"),
            ("Cartera de pedidos / contratos comerciales / evidencia de demanda", "Volumen E05-E24, precios E26-E29"),
            ("Contrato de préstamo (desembolsos, tasa, comisiones, garantías)", "L01-L06; T2 S_DEUDA1"),
            ("Acta de aportes de socios y política de dividendos", "Aportes, F. política de dividendos"),
            ("Tasación / valor de realización de activos", "G24 realización de activo fijo"),
            ("Estudio técnico de capacidad (equipos, turnos, horas, cuellos de botella)", "Capacidad, E25")]
    for i, (d, e) in enumerate(docs):
        setc(ws, r, 1, i + 1, fmt=NF_INT0)
        setc(ws, r, 2, d)
        setc(ws, r, 3, "Pendiente", font=F_IN, fill=FILL_PEND)
        setc(ws, r, 4, None, font=F_IN, fill=FILL_IN, fmt=NF_DATE)
        setc(ws, r, 5, None, font=F_IN, fill=FILL_IN)
        setc(ws, r, 6, e)
        r += 1
    DRT["doc_last"] = r - 1
    dv_list(ws, f"C{DRT['doc_first']}:C{DRT['doc_last']}", ["Pendiente", "Recibido", "Validado", "No aplica"])
    r += 1
    section(ws, r, "T6. Campos mínimos para considerar completo el forecast actualizado", 14); r += 1
    header(ws, r, ["N°", "Campo mínimo", "Completo (1/0)", "", "", "Criterio"]); r += 1
    DRT["min_first"] = r
    minimos = [
        ("Último mes cerrado cargado", f'=IF(ISNUMBER(r_Corte),1,0)', "T1"),
        ("Mes real de arranque o confirmación de atraso", f'=IF(ISNUMBER(r_Arranque),1,0)', "T1"),
        ("Fecha calendario del mes 0", '=IF(ISNUMBER(p_FechaInicio),1,0)', "Supuestos G04"),
        ("Saldos de caja, clientes, proveedores e inventarios al corte",
         '=IF(AND(ISNUMBER(r_S_CAJA),ISNUMBER(r_S_CXC),ISNUMBER(r_S_PROVMAT),ISNUMBER(r_S_PROVMEN),ISNUMBER(r_S_INVMAT)),1,0)', "T2"),
        ("Saldo de deuda al corte", '=IF(ISNUMBER(r_S_DEUDA1),1,0)', "T2"),
        ("CAPEX ejecutado/pagado/comprometido por ítem", f'=IF(COUNT(C31:E{DRT["capex_last"]})>0,1,0)', "T3"),
        ("Movimientos reales mensuales (ventas, costos, cobros y pagos)", '=IF(COUNTIFS(dr_Cod,"<>",dr_Mes,">=0")>0,1,0)', "T7"),
        ("Contrato y modalidad de maquila validados", f'=IF(OR(C{DRT["doc_first"]}="Validado",C{DRT["doc_first"]}="Recibido"),1,0)', "T5-1"),
        ("Tipo de cambio real/proyectado cargado en escenario Actualizado", "=IF(Supuestos!$E$%d<>Supuestos!$D$%d,1,0)" % (SUP["tc1"], SUP["tc1"]), "Supuestos E49"),
        ("Tasas vigentes (Ke, deuda) con fecha y fundamento", 0, "Supuestos G10-G17 (marcar 1 al validar)"),
    ]
    for i, (lab, f, crit) in enumerate(minimos):
        setc(ws, r, 1, i + 1, fmt=NF_INT0)
        setc(ws, r, 2, lab)
        setc(ws, r, 3, f, fmt=NF_INT0, font=F_IN if not isinstance(f, str) else F_BASE, fill=FILL_IN if not isinstance(f, str) else None)
        setc(ws, r, 6, crit)
        r += 1
    DRT["min_last"] = r - 1
    DRT["min_tot"] = r
    setc(ws, r, 2, "Campos mínimos completos", font=F_BOLD)
    setc(ws, r, 3, f"=SUM(C{DRT['min_first']}:C{DRT['min_last']})&\" de {len(minimos)}\"", font=F_BOLD)
    add_name("r_MinCompletos", "DR", f"C{r}")
    r += 1
    setc(ws, r, 2, "Cantidad completa (número)")
    setc(ws, r, 3, f"=SUM(C{DRT['min_first']}:C{DRT['min_last']})", fmt=NF_INT0)
    add_name("r_MinN", "DR", f"C{r}")
    r += 2
    section(ws, r, "T7. Movimientos mensuales reales (formato largo). Filas 1 a 600", 14); r += 1
    setc(ws, r, 1, "EJEMPLO — NO SE INCLUYE EN CÁLCULOS (fuera del rango de la tabla):", font=F_ALERT)
    r += 1
    header(ws, r, ["ID", "Mes modelo", "Fecha", "Código", "Producto / ítem", "Tipo de dato", "Moneda", "Importe origen", "Cantidad",
                   "TC Gs/USD", "Importe USD", "Fuente", "Estado", "Comentario"]); r += 1
    ex = ["EJ-0001", 3, None, "NOM_D", "", "Devengado = pagado", "PYG", 1250000000, None, 7300, None, "Planilla IPS mar-26", "Ejemplo", "Formato ilustrativo"]
    for c, v in enumerate(ex):
        setc(ws, r, c + 1, v, font=F_SUB, fmt=NF_GS if c == 7 else None)
    setc(ws, r, 11, f"=H{r}/J{r}", font=F_SUB, fmt=NF_USD)
    r += 2
    header(ws, r, ["ID", "Mes modelo", "Fecha", "Código", "Producto / ítem", "Tipo de dato", "Moneda", "Importe origen", "Cantidad",
                   "TC Gs/USD", "Importe USD", "Fuente", "Estado", "Comentario"]); r += 1
    DRT["mov_first"] = r
    for i in range(600):
        setc(ws, r, 1, None, font=F_IN, fill=FILL_IN)
        setc(ws, r, 2, None, font=F_IN, fill=FILL_IN, fmt=NF_INT0)
        setc(ws, r, 3, None, font=F_IN, fill=FILL_IN, fmt=NF_DATE)
        setc(ws, r, 4, None, font=F_IN, fill=FILL_IN)
        setc(ws, r, 5, None, font=F_IN, fill=FILL_IN)
        setc(ws, r, 6, None, font=F_IN, fill=FILL_IN)
        setc(ws, r, 7, None, font=F_IN, fill=FILL_IN)
        setc(ws, r, 8, None, font=F_IN, fill=FILL_IN, fmt=NF_USD2)
        setc(ws, r, 9, None, font=F_IN, fill=FILL_IN, fmt=NF_INT)
        setc(ws, r, 10, None, font=F_IN, fill=FILL_IN, fmt=NF_INT)
        setc(ws, r, 11, f'=IF(H{r}="","",IF(G{r}="PYG",IF(N(J{r})=0,"FALTA TC",H{r}/J{r}),H{r}))', fmt=NF_USD2)
        setc(ws, r, 12, None, font=F_IN, fill=FILL_IN)
        setc(ws, r, 13, None, font=F_IN, fill=FILL_IN)
        setc(ws, r, 14, None, font=F_IN, fill=FILL_IN)
        r += 1
    DRT["mov_last"] = r - 1
    mf, ml = DRT["mov_first"], DRT["mov_last"]
    add_range_name("dr_Mes", "DR", f"$B${mf}:$B${ml}")
    add_range_name("dr_Cod", "DR", f"$D${mf}:$D${ml}")
    add_range_name("dr_USD", "DR", f"$K${mf}:$K${ml}")
    add_range_name("dr_Cant", "DR", f"$I${mf}:$I${ml}")
    add_range_name("dr_Est", "DR", f"$M${mf}:$M${ml}")
    dv_range(ws, f"D{mf}:D{ml}", f"=$A${DRT['cat_first']}:$A${DRT['cat_last']}")
    dv_list(ws, f"G{mf}:G{ml}", ["USD", "PYG"])
    dv_list(ws, f"M{mf}:M{ml}", ["Borrador", "Revisado", "Aprobado", "Excluido"])
    dv_list(ws, f"F{mf}:F{ml}", ["Devengado", "Pagado/cobrado", "Devengado = pagado", "Cantidad", "Compromiso"])
    ws.freeze_panes = "C4"


CODES = [
    ("U_PROD_G", "Gabinetes buenos producidos", "unidades", "Cantidad", "Produccion_Ventas: producción buena"),
    ("U_PROD_C", "Cajas buenas producidas", "unidades", "Cantidad", "Produccion_Ventas: producción buena"),
    ("U_RCH_G", "Gabinetes rechazados", "unidades", "Cantidad", "Produccion_Ventas: rechazos"),
    ("U_RCH_C", "Cajas rechazadas", "unidades", "Cantidad", "Produccion_Ventas: rechazos"),
    ("U_VTA_G", "Gabinetes vendidos / facturados", "unidades", "Cantidad", "Produccion_Ventas: ventas"),
    ("U_VTA_C", "Cajas vendidas / facturadas", "unidades", "Cantidad", "Produccion_Ventas: ventas"),
    ("VTA_G", "Facturación neta gabinetes", "USD", "Devengado", "Ventas gabinetes"),
    ("VTA_C", "Facturación neta cajas", "USD", "Devengado", "Ventas cajas"),
    ("VTA_CHAT", "Venta de chatarra", "USD", "Devengado = cobrado", "Otros ingresos"),
    ("COBRO", "Cobranzas totales de clientes (incl. contado y anticipos)", "USD", "Pagado/cobrado", "Capital_Trabajo: cobranzas"),
    ("MP", "Materia prima consumida", "USD", "Devengado", "Costo de producción"),
    ("INS", "Insumos consumidos", "USD", "Devengado", "Costo de producción"),
    ("CON", "Consumibles", "USD", "Devengado", "Costo de producción"),
    ("FLE", "Flete de materia prima", "USD", "Devengado", "Costo de producción"),
    ("DES", "Despachos de importación", "USD", "Devengado", "Costo de producción"),
    ("PKG", "Packaging", "USD", "Devengado", "Costo de producción"),
    ("COMPRA_MAT", "Compras de materiales a costo de adquisición", "USD", "Devengado", "Inventario de materiales"),
    ("PAGO_MAT", "Pagos a proveedores de materiales (incl. anticipos)", "USD", "Pagado/cobrado", "Capital_Trabajo"),
    ("NOM_D", "Costo total personal directo (remuneración + cargas + alimentación)", "USD", "Devengado = pagado", "Costos_Personal"),
    ("NOM_I", "Costo total personal indirecto de fábrica", "USD", "Devengado = pagado", "Costos_Personal"),
    ("NOM_A", "Costo total personal administrativo", "USD", "Devengado = pagado", "Costos_Personal"),
    ("ENE", "Energía eléctrica planta", "USD", "Devengado", "Costos de producción"),
    ("MAN", "Mantenimiento planta", "USD", "Devengado", "Costos de producción"),
    ("IMPP", "Imprevistos / otros de producción", "USD", "Devengado", "Costos de producción"),
    ("SEGC", "Seguro de caución", "USD", "Devengado", "Gastos variables de venta"),
    ("TRIB", "Tributo único de maquila", "USD", "Devengado", "Gastos variables de venta"),
    ("GVTA", "Gastos de venta y logística de exportación", "USD", "Devengado", "Gastos variables de venta"),
    ("ALQ", "Alquiler de nave (incluye IVA)", "USD", "Devengado", "Administración"),
    ("EXP", "Expensas", "USD", "Devengado", "Administración"),
    ("SRV", "Agua, comunicaciones y electricidad adm.", "USD", "Devengado", "Administración"),
    ("MAA", "Mantenimiento administración", "USD", "Devengado", "Administración"),
    ("SGA", "Seguros sobre activos", "USD", "Devengado", "Administración"),
    ("GG", "Gastos generales", "USD", "Devengado", "Administración"),
    ("HON", "Honorarios y otros fijos", "USD", "Devengado", "Administración"),
    ("IMPA", "Imprevistos / otros de administración", "USD", "Devengado", "Administración"),
    ("PREOP", "Costos preoperativos, pruebas, capacitación, pérdidas de arranque y extraordinarios", "USD", "Devengado", "Resultados (línea separada)"),
    ("PAGO_MEN", "Pagos a proveedores menores y obligaciones", "USD", "Pagado/cobrado", "Capital_Trabajo"),
    ("IVA_RECUP", "IVA recuperado / certificados cobrados", "USD", "Pagado/cobrado", "Deuda_Tributos"),
    ("CAPEX_DEV", "CAPEX devengado del mes", "USD", "Devengado", "Inversion_Activos"),
    ("CAPEX_PAG", "CAPEX pagado del mes", "USD", "Pagado/cobrado", "Inversion_Activos"),
    ("DESEMB_P1", "Desembolso préstamo original", "USD", "Pagado/cobrado", "Deuda_Tributos"),
    ("INT_P1", "Intereses préstamo original (sin IVA)", "USD", "Pagado/cobrado", "Deuda_Tributos"),
    ("AMORT_P1", "Amortización préstamo original", "USD", "Pagado/cobrado", "Deuda_Tributos"),
    ("DESEMB_P2", "Desembolso tramo 2", "USD", "Pagado/cobrado", "Deuda_Tributos"),
    ("INT_P2", "Intereses tramo 2 (sin IVA)", "USD", "Pagado/cobrado", "Deuda_Tributos"),
    ("AMORT_P2", "Amortización tramo 2", "USD", "Pagado/cobrado", "Deuda_Tributos"),
    ("APORTE", "Aportes de socios en efectivo", "USD", "Pagado/cobrado", "Deuda_Tributos: aportes"),
    ("DIVIDENDO", "Dividendos pagados", "USD", "Pagado/cobrado", "Resultados_Caja: tesorería"),
]

build_cp_tables()
build_ia_table()
build_datos_reales()


# ===========================================================================
# 4) MOTOR MENSUAL (especificaciones por fila; fórmulas idénticas por mes)
# ===========================================================================
def S(t):
    return ("sec", t)


def HDR():
    return ("hdr",)


def RW(key, label, unit, fn, fmt=NF_USD, tot="sum", note=None, bold=False):
    return ("row", key, label, unit, fn, fmt, tot, note, bold)


def hv(k, m):
    return A(f"{CUR[0]}.h_{k}", m)


def cal(k, m):
    return A(f"CAL.{k}", m)


def RS(code, m, qty=False):
    return f'SUMIFS({"dr_Cant" if qty else "dr_USD"},dr_Mes,{m},dr_Cod,"{code}",dr_Est,"<>Excluido")'


def realor(code, m, fc, qty=False):
    return f"IF({hv('freal', m)}=1,{RS(code, m, qty)},{fc})"


def prv(key, m, d="0"):
    return A(key, m - 1) if m > 0 else d


def nxt(key, m):
    return A(key, m + 1) if m < N else "0"


def lagf(key, m, Li, Lf):
    return (f"(IF({m}-{Li}>=0,INDEX({RG(key)},1,{m}-{Li}+1),0)*(1-{Lf})"
            f"+IF({m}-{Li}-1>=0,INDEX({RG(key)},1,{m}-{Li}),0)*{Lf})")


def fwd(key, m, k):
    """Valor de la fila 'key' k meses adelante (k puede ser nombre)."""
    return f"IF({m}+{k}<={N},INDEX({RG(key)},1,{m}+{k}+1),0)"


def over(m, name, calc):
    """Saldo: en el mes de corte usa el saldo real si fue cargado."""
    return f"IF(AND({m}=p_MesCorte,ISNUMBER({name})),{name},{calc})"


def rsum(key, m, n=12):
    a = max(0, m - n + 1)
    r = ROWS[key]
    return f"SUM({mc(a)}{r}:{mc(m)}{r})"


# ------------------------------------------------------------------ Calendario
def cal_specs():
    return [
        S("Calendario del modelo (meses relativos; fechas solo si se cargó la fecha del mes 0)"),
        RW("m", "Mes del modelo (0 = inicio de implantación / desembolso)", "n°", lambda m: m, NF_INT0, "none"),
        RW("fecha", "Fecha de fin de mes", "fecha", lambda m: f'=IF(p_FechaInicio="","s/f",EOMONTH(p_FechaInicio,{m}))', NF_MES, "none"),
        RW("anio", "Año de proyecto (0 = implantación)", "n°", lambda m: f"=IF({m}<=p_ImplMeses,0,ROUNDUP(({m}-p_ImplMeses)/12,0))", NF_INT0, "none"),
        RW("ini", "Inicio de año de proyecto (1/0)", "1/0", lambda m: f"=IF(AND({m}>p_ImplMeses,MOD({m}-p_ImplMeses-1,12)=0),1,0)", NF_INT0, "none"),
        RW("cierre", "Cierre de año de proyecto (1/0)", "1/0", lambda m: f"=IF(AND({m}>p_ImplMeses,MOD({m}-p_ImplMeses,12)=0),1,0)", NF_INT0, "none"),
        RW("mes_op", "Mes operativo desde el arranque efectivo", "n°", lambda m: f"=IF(AND({m}>=p_MesArranque,{m}<=p_MesFin),{m}-p_MesArranque+1,0)", NF_INT0, "none"),
        RW("anio_op", "Año operativo (curva de volumen, capacidad y dotación)", "n°",
           lambda m: f"=IF({cal('mes_op', m)}>0,MIN(p_AniosOp,ROUNDUP({cal('mes_op', m)}/12,0)),0)", NF_INT0, "none"),
        RW("f_op", "Operación (1/0)", "1/0", lambda m: f"=IF({cal('mes_op', m)}>0,1,0)", NF_INT0, "sum"),
        RW("f_hor", "Dentro del horizonte evaluado (1/0)", "1/0", lambda m: f"=IF({m}<=p_MesFin,1,0)", NF_INT0, "sum"),
        RW("f_pre", "Pre-operación: implantación o atraso con costos de espera (1/0)", "1/0",
           lambda m: f"=IF(AND({m}>=1,{m}<p_MesArranque,{m}<=p_MesFin),1,0)", NF_INT0, "sum"),
        RW("f_fin", "Mes final del horizonte (1/0)", "1/0", lambda m: f"=IF({m}=p_MesFin,1,0)", NF_INT0, "sum"),
        RW("mes_plan", "Mes operativo del plan original (sin atraso)", "n°",
           lambda m: f"=IF(AND({m}>p_ImplMeses,{m}<=p_ImplMeses+p_MesesOp),{m}-p_ImplMeses,0)", NF_INT0, "none"),
        RW("anio_plan", "Año operativo del plan original", "n°",
           lambda m: f"=IF({cal('mes_plan', m)}>0,ROUNDUP({cal('mes_plan', m)}/12,0),0)", NF_INT0, "none"),
        RW("f_real", "Mes real cerrado (1/0) — congelado", "1/0", lambda m: f"=IF({m}<=p_MesCorte,1,0)", NF_INT0, "sum"),
        RW("anio_px", "Índice de año para precios, costos y TC (año de proyecto, mín. 1)", "n°", lambda m: f"=MAX(1,{cal('anio', m)})", NF_INT0, "none"),
        RW("tc", "Tipo de cambio del mes (escenario efectivo)", "Gs/USD", lambda m: f"=INDEX(e_TC,MIN(14,{cal('anio', m)}+1))", NF_INT, "none"),
        RW("df_p", "Factor de descuento del proyecto (tasa anual efectiva, capitalización mensual)", "factor",
           lambda m: f"=1/(1+p_TasaProy)^({m}/12)", NF_USD4, "none"),
        RW("df_e", "Factor de descuento del capital propio (Ke)", "factor", lambda m: f"=1/(1+p_Ke)^({m}/12)", NF_USD4, "none"),
    ]


def hdr_rows(S_):
    return [("h_m", "Mes del modelo", "m", NF_INT0), ("h_fecha", "Fecha", "fecha", NF_MES), ("h_anio", "Año de proyecto", "anio", NF_INT0),
            ("h_aop", "Año operativo", "anio_op", NF_INT0), ("h_fop", "Operación (1/0)", "f_op", NF_INT0),
            ("h_freal", "Real cerrado (1/0)", "f_real", NF_INT0)]


# ------------------------------------------------------------------ Produccion_Ventas
def pv_specs():
    sp = [HDR(), S("Curva de aprendizaje (común a ambos productos)"),
          RW("rampa", "Factor de rampa (% del ritmo del plan alcanzado)", "%",
             lambda m: f"=IF({hv('fop', m)}=0,0,IF(e_Rampa<=0,1,MIN(1,e_Eff0+(1-e_Eff0)*({cal('mes_op', m)}-1)/e_Rampa)))", NF_PCT, "none")]
    for p, P, nm in [("g", "G", "Gabinetes (153 kg)"), ("c", "C", "Cajas (13 kg)")]:
        cap, vta = BP["cap_" + p], BP["vta_" + p]
        vol = "e_VolG" if p == "g" else "e_VolC"
        K = lambda s, p=p: f"PV.{s}_{p}"
        sp += [
            S(f"{nm} — unidades"),
            RW(f"cap_{p}", "Capacidad instalada del plan (u/mes)", "u",
               lambda m, cap=cap: f"=IF({hv('fop', m)}=1,INDEX(Base_PDF!$F${cap}:$O${cap},1,{hv('aop', m)})/12*(1+e_CapAdic),0)", NF_INT),
            RW(f"plan_{p}", "Plan PDF de ventas según año operativo (u/mes)", "u",
               lambda m, vta=vta: f"=IF({hv('fop', m)}=1,INDEX(Base_PDF!$F${vta}:$O${vta},1,{hv('aop', m)})/12,0)", NF_INT),
            RW(f"fvol_{p}", "Factor de volumen del escenario", "%", lambda m, vol=vol: f"=IF({hv('fop', m)}=1,INDEX({vol},{hv('aop', m)}),0)", NF_PCT, "none"),
            RW(f"dem_{p}", "Demanda del escenario (u)", "u", lambda m, K=K: f"={A(K('plan'), m)}*{A(K('fvol'), m)}", NF_INT),
            RW(f"vbase_{p}", "Ventas posibles con rampa (u)", "u", lambda m, K=K: f"={A(K('dem'), m)}*{A('PV.rampa', m)}", NF_INT),
            RW(f"prampa_{p}", "Volumen no producido por rampa (u)", "u", lambda m, K=K: f"={A(K('dem'), m)}-{A(K('vbase'), m)}", NF_INT),
            RW(f"patraso_{p}", "Demanda del plan no atendida por atraso (u)", "u",
               lambda m, vta=vta, vol=vol: (f"=IF(AND({cal('mes_plan', m)}>0,{m}<p_MesArranque),INDEX(Base_PDF!$F${vta}:$O${vta},1,{cal('anio_plan', m)})/12"
                                            f"*INDEX({vol},{cal('anio_plan', m)}),0)"), NF_INT),
            RW(f"holg_{p}", "Holgura de capacidad para recuperar (u buenas)", "u",
               lambda m, K=K: f"=IF({hv('fop', m)}=1,MAX(0,{A(K('cap'), m)}*{A('PV.rampa', m)}*(1-e_Rech)-{A(K('vbase'), m)}),0)", NF_INT),
            RW(f"rec_{p}", "Volumen recuperado (desplazamiento)", "u",
               lambda m, K=K: "=0" if m == 0 else f"=MIN({prv(K('pend'), m)},{A(K('holg'), m)})", NF_INT),
            RW(f"pend_{p}", "Pendiente recuperable acumulado (u)", "u",
               lambda m, K=K: f"={prv(K('pend'), m)}+({A(K('prampa'), m)}+{A(K('patraso'), m)})*e_Recup-{A(K('rec'), m)}", NF_INT, "last"),
            RW(f"pdef_{p}", "Pérdida definitiva de ventas (u)", "u", lambda m, K=K: f"=({A(K('prampa'), m)}+{A(K('patraso'), m)})*(1-e_Recup)", NF_INT),
            RW(f"vtas_u_{p}", "UNIDADES VENDIDAS", "u",
               lambda m, K=K, P=P: "=" + realor(f"U_VTA_{P}", m, f"{A(K('vbase'), m)}+{A(K('rec'), m)}", qty=True), NF_INT, bold=True),
            RW(f"ptobj_{p}", "Stock objetivo de PT (u) = ventas del mes siguiente x días PT", "u",
               lambda m, K=K: f"=IF({hv('fop', m)}=1,{nxt(K('vtas_u'), m)}*p_PTf,0)", NF_INT, "none"),
            RW(f"prod_{p}", "Producción buena (u)", "u",
               lambda m, K=K, P=P: "=" + realor(f"U_PROD_{P}", m, f"MAX(0,{A(K('vtas_u'), m)}+{A(K('ptobj'), m)}-{prv(K('ptu'), m)})", qty=True), NF_INT, bold=True),
            RW(f"ptu_{p}", "Inventario de PT (u): inicial + producción - ventas", "u",
               lambda m, K=K: f"={prv(K('ptu'), m)}+{A(K('prod'), m)}-{A(K('vtas_u'), m)}", NF_INT, "last"),
            RW(f"ini_{p}", "Unidades iniciadas (incluye rechazo)", "u",
               lambda m, K=K, P=P: (f"=IF({hv('freal', m)}=1,{A(K('prod'), m)}+{RS('U_RCH_' + P, m, True)},"
                                    f"IF(e_Rech<1,{A(K('prod'), m)}/(1-e_Rech),0))"), NF_INT),
            RW(f"rch_{p}", "Unidades rechazadas", "u", lambda m, K=K: f"={A(K('ini'), m)}-{A(K('prod'), m)}", NF_INT),
            RW(f"exc_{p}", "ALERTA: exceso de unidades iniciadas sobre capacidad", "u",
               lambda m, K=K: f"=MAX(0,{A(K('ini'), m)}-{A(K('cap'), m)}*1.0001)", NF_INT),
            RW(f"precio_{p}", "Precio / tarifa neta vigente (USD/u)", "USD/u",
               lambda m, P=P: (f"=IF(p_Modalidad=2,N(p_Tarifa{P}),e_Precio{P}1)*(1+e_PrecioVar)*(1+e_PrecioGr)^({cal('anio_px', m)}-1)"),
               NF_USD2, "none"),
            RW(f"vtas_{p}", "VENTAS (USD)", "USD",
               lambda m, K=K, P=P: "=" + realor(f"VTA_{P}", m, f"{A(K('vtas_u'), m)}*{A(K('precio'), m)}"), NF_USD, bold=True),
        ]
    sp += [
        S("Totales"),
        RW("vtas_prod", "Ventas de productos (USD)", "USD", lambda m: f"={A('PV.vtas_g', m)}+{A('PV.vtas_c', m)}", NF_USD, bold=True),
        RW("chat", "Venta de chatarra (USD)", "USD", lambda m: "=" + realor("VTA_CHAT", m, f"({A('PV.rch_g', m)}+{A('PV.rch_c', m)})*e_Chat")),
        RW("vtas_tot", "VENTAS TOTALES (USD)", "USD", lambda m: f"={A('PV.vtas_prod', m)}+{A('PV.chat', m)}", NF_USD, bold=True),
        RW("kg_prod", "Kg procesados (producción buena)", "kg", lambda m: f"={A('PV.prod_g', m)}*p_KgG+{A('PV.prod_c', m)}*p_KgC", NF_INT),
        RW("exc_tot", "Exceso total sobre capacidad (u)", "u", lambda m: f"={A('PV.exc_g', m)}+{A('PV.exc_c', m)}", NF_INT),
        RW("exc_flag", "Mes con producción superior a la capacidad (1/0)", "1/0", lambda m: f"=IF({A('PV.exc_tot', m)}>0.5,1,0)", NF_INT0),
    ]
    return sp


# ------------------------------------------------------------------ Costos_Personal (motor)
def usd_cm(code, m):
    return f"IF({cm(code, 'E')}=\"PYG\",{cm(code, 'F')}/{A('CP.tc', m)},{cm(code, 'F')})"


def cp_specs():
    sp = [HDR(), S("Tipo de cambio e índices"),
          RW("tc", "Tipo de cambio del mes (Gs/USD)", "Gs/USD", lambda m: f"={cal('tc', m)}", NF_INT, "none"),
          RW("tcn", "TC aplicado a la nómina (C1: único; sin C1: 8.000 del PDF)", "Gs/USD", lambda m: f"=IF(p_C1=1,{A('CP.tc', m)},p_TC_NomPDF)", NF_INT, "none"),
          RW("isal", "Índice salarial (ajuste anual en Gs)", "índice", lambda m: f"=(1+e_AjSal)^({cal('anio_px', m)}-1)", NF_USD4, "none"),
          RW("ifij", "Índice de costos fijos", "índice", lambda m: f"=(1+e_InflFij)^({cal('anio_px', m)}-1)", NF_USD4, "none"),
          RW("afij", "Actividad de costos fijos (1 operación; % espera; 0 fuera)", "%", lambda m: f"=IF({hv('fop', m)}=1,1,IF({cal('f_pre', m)}=1,p_EspFijos,0))", NF_PCT, "none"),
          RW("aene", "Actividad de energía y mantenimiento de planta", "%", lambda m: f"=IF({hv('fop', m)}=1,1,IF({cal('f_pre', m)}=1,p_EspEne,0))", NF_PCT, "none")]
    groups = [("d", "Directo", "p_EspDir", CPT["dot_d_aj"], CPT["masa_d_aj"], "NOM_D", "Personal directo"),
              ("i", "Indirecto fábrica", "p_EspInd", CPT["dot_i"], CPT["masa_i"], "NOM_I", "Personal indirecto de fábrica"),
              ("a", "Administración", "p_EspInd", CPT["dot_a"], CPT["masa_a"], "NOM_A", "Personal administrativo")]
    for k, cls, esp, rdot, rmasa, code, nm in groups:
        rd = f"$G${rdot}:$P${rdot}"
        rmm = f"$G${rmasa}:$P${rmasa}"
        sp += [S(nm + " (clase efectiva según C6)"),
               RW(f"dot_{k}", "Dotación activa (personas)", "pers.",
                  lambda m, rd=rd, esp=esp: f"=IF({hv('fop', m)}=1,INDEX({rd},1,{hv('aop', m)}),IF({cal('f_pre', m)}=1,INDEX({rd},1,1)*{esp},0))", NF_USD2, "none"),
               RW(f"rem_{k}", "Remuneración base (USD)", "USD",
                  lambda m, rmm=rmm, esp=esp: (f"=IF({hv('fop', m)}=1,INDEX({rmm},1,{hv('aop', m)}),IF({cal('f_pre', m)}=1,INDEX({rmm},1,1)*{esp},0))"
                                               f"*{A('CP.isal', m)}/{A('CP.tcn', m)}")),
               RW(f"pre_{k}", "Prestaciones sociales (coeficiente del plan)", "USD", lambda m, k=k: f"={A(f'CP.rem_{k}', m)}*p_Prest"),
               RW(f"ali_{k}", "Alimentación", "USD", lambda m, k=k: f"={A(f'CP.dot_{k}', m)}*p_AlimGs/{A('CP.tc', m)}"),
               RW(f"tot_{k}", "COSTO TOTAL " + nm.upper(), "USD",
                  lambda m, k=k, code=code: "=" + realor(code, m, f"{A(f'CP.rem_{k}', m)}+{A(f'CP.pre_{k}', m)}+{A(f'CP.ali_{k}', m)}"), NF_USD, bold=True)]
    sp += [S("Resumen de personal"),
           RW("dot_tot", "Dotación total activa", "pers.", lambda m: f"={A('CP.dot_d', m)}+{A('CP.dot_i', m)}+{A('CP.dot_a', m)}", NF_USD2, "max"),
           RW("rem_tot", "Remuneración base total (base de anticipos)", "USD", lambda m: f"={A('CP.rem_d', m)}+{A('CP.rem_i', m)}+{A('CP.rem_a', m)}"),
           RW("per_tot", "Costo total de personal", "USD", lambda m: f"={A('CP.tot_d', m)}+{A('CP.tot_i', m)}+{A('CP.tot_a', m)}", NF_USD, bold=True)]
    sp += [S("Materiales: costo unitario por año de proyecto (USD/u)"),
           RW("u_mp_g", "MP por gabinete", "USD/u", lambda m: f"=e_MPg*(1+e_MPvar)*(1+e_MPgr)^({cal('anio_px', m)}-1)", NF_USD4, "none"),
           RW("u_mp_c", "MP por caja (C2: 1% acumulativo; sin C2: criterio PDF)", "USD/u",
              lambda m: f"=e_MPc*(1+e_MPvar)*IF(p_C2=1,(1+e_MPgr)^({cal('anio_px', m)}-1),IF({cal('anio_px', m)}=1,1,1+e_MPgr))", NF_USD4, "none"),
           RW("u_ins_g", "Insumos por gabinete", "USD/u", lambda m: f"=e_INSg*(1+e_INSvar)*(1+e_INSgr)^({cal('anio_px', m)}-1)", NF_USD4, "none"),
           RW("u_ins_c", "Insumos por caja", "USD/u", lambda m: f"=e_INSc*(1+e_INSvar)*(1+e_INSgr)^({cal('anio_px', m)}-1)", NF_USD4, "none"),
           RW("u_con_g", "Consumibles por gabinete", "USD/u", lambda m: f"=e_CONg*(1+e_INSvar)*(1+e_INSgr)^({cal('anio_px', m)}-1)", NF_USD4, "none"),
           RW("u_con_c", "Consumibles por caja", "USD/u", lambda m: f"=e_CONc*(1+e_INSvar)*(1+e_INSgr)^({cal('anio_px', m)}-1)", NF_USD4, "none"),
           RW("u_fle_g", "Flete MP por gabinete", "USD/u", lambda m: "=e_FLEg*(1+e_FLEvar)", NF_USD4, "none"),
           RW("u_fle_c", "Flete MP por caja", "USD/u", lambda m: "=e_FLEc*(1+e_FLEvar)", NF_USD4, "none"),
           RW("own", "Factor de materiales propios (modalidad 2 = materiales del cliente)", "%", lambda m: "=IF(p_Modalidad=2,1-p_PropCliente,1)", NF_PCT, "none"),
           S("Materiales consumidos (USD)"),
           RW("mp", "Materia prima", "USD",
              lambda m: "=" + realor("MP", m, f"({A('PV.ini_g', m)}*{A('CP.u_mp_g', m)}+{A('PV.ini_c', m)}*{A('CP.u_mp_c', m)})*{A('CP.own', m)}")),
           RW("ins", "Insumos", "USD",
              lambda m: "=" + realor("INS", m, f"({A('PV.ini_g', m)}*{A('CP.u_ins_g', m)}+{A('PV.ini_c', m)}*{A('CP.u_ins_c', m)})*{A('CP.own', m)}")),
           RW("con", "Consumibles", "USD",
              lambda m: "=" + realor("CON", m, f"{A('PV.ini_g', m)}*{A('CP.u_con_g', m)}+{A('PV.ini_c', m)}*{A('CP.u_con_c', m)}")),
           RW("fle", "Flete marítimo y terrestre de MP", "USD",
              lambda m: "=" + realor("FLE", m, f"({A('PV.ini_g', m)}*{A('CP.u_fle_g', m)}+{A('PV.ini_c', m)}*{A('CP.u_fle_c', m)})*{A('CP.own', m)}")),
           RW("des", "Despachos de importación (% de MP + insumos)", "USD", lambda m: "=" + realor("DES", m, f"p_Desp*({A('CP.mp', m)}+{A('CP.ins', m)})")),
           RW("mat_land", "Consumo valuado a costo de adquisición (C5 incluye flete y despacho)", "USD",
              lambda m: f"={A('CP.mp', m)}+{A('CP.ins', m)}+{A('CP.con', m)}+IF(p_C5=1,{A('CP.fle', m)}+{A('CP.des', m)},0)", NF_USD, bold=True),
           RW("std_g", "Costo estándar de materiales gabinetes (para asignación)", "USD",
              lambda m: f"={A('PV.ini_g', m)}*(({A('CP.u_mp_g', m)}+{A('CP.u_ins_g', m)}+{A('CP.u_fle_g', m)})*{A('CP.own', m)}+{A('CP.u_con_g', m)})"),
           RW("std_c", "Costo estándar de materiales cajas (para asignación)", "USD",
              lambda m: f"={A('PV.ini_c', m)}*(({A('CP.u_mp_c', m)}+{A('CP.u_ins_c', m)}+{A('CP.u_fle_c', m)})*{A('CP.own', m)}+{A('CP.u_con_c', m)})"),
           RW("umat_g", "Materiales por gabinete bueno (valuación de PT)", "USD/u",
              lambda m: (f"=IF(e_Rech<1,(({A('CP.u_mp_g', m)}+{A('CP.u_ins_g', m)})*{A('CP.own', m)}*(1+IF(p_C5=1,p_Desp,0))+{A('CP.u_con_g', m)}"
                         f"+IF(p_C5=1,{A('CP.u_fle_g', m)}*{A('CP.own', m)},0))/(1-e_Rech),0)"), NF_USD4, "none"),
           RW("umat_c", "Materiales por caja buena (valuación de PT)", "USD/u",
              lambda m: (f"=IF(e_Rech<1,(({A('CP.u_mp_c', m)}+{A('CP.u_ins_c', m)})*{A('CP.own', m)}*(1+IF(p_C5=1,p_Desp,0))+{A('CP.u_con_c', m)}"
                         f"+IF(p_C5=1,{A('CP.u_fle_c', m)}*{A('CP.own', m)},0))/(1-e_Rech),0)"), NF_USD4, "none"),
           S("Otros costos de producción (USD)"),
           RW("ene", "Energía eléctrica planta", "USD",
              lambda m: "=" + realor("ENE", m, (f"{usd_cm('ENE', m)}/12*{A('CP.ifij', m)}*IF({hv('fop', m)}=1,(1-e_EneVar)+e_EneVar*{A('PV.kg_prod', m)}/p_KgBaseMes,"
                                                 f"IF({cal('f_pre', m)}=1,p_EspEne,0))"))),
           RW("man", "Mantenimiento y reparaciones planta", "USD",
              lambda m: "=" + realor("MAN", m, f"{usd_cm('MAN', m)}/12*(1+p_ManInc*(MAX(1,{hv('aop', m)})-1))*{A('CP.ifij', m)}*{A('CP.aene', m)}")),
           RW("pkg", "Packaging (base identificada del PDF)", "USD",
              lambda m: "=" + realor("PKG", m, (f"IF({hv('fop', m)}=1,p_Pkg*({A('CP.mp', m)}+{A('CP.ins', m)}+{A('CP.con', m)}+{A('CP.fle', m)}"
                                                 f"+{A('CP.rem_d', m)}+{A('CP.pre_d', m)}+{A('CP.ali_d', m)}),0)"))),
           RW("impp", "Imprevistos de producción", "USD",
              lambda m: "=" + realor("IMPP", m, (f"p_ImpP*({A('CP.mp', m)}+{A('CP.ins', m)}+{A('CP.con', m)}+{A('CP.fle', m)}+{A('CP.des', m)}+{A('CP.pkg', m)}"
                                                  f"+{A('CP.tot_d', m)}+{A('CP.tot_i', m)}+{A('CP.ene', m)}+{A('CP.man', m)}+{A('CP.segc', m)})"))),
           S("Gastos variables de venta y tributos (USD)"),
           RW("segc", "Seguro de caución maquila", "USD", lambda m: "=" + realor("SEGC", m, f"p_SegCau*{A('PV.vtas_prod', m)}")),
           RW("gvta", "Gastos de venta y logística de exportación", "USD", lambda m: "=" + realor("GVTA", m, f"e_GV*{A('PV.vtas_prod', m)}")),
           RW("va_nac", "Valor agregado nacional (base alternativa del tributo)", "USD",
              lambda m: (f"={A('CP.tot_d', m)}+{A('CP.tot_i', m)}+{A('CP.tot_a', m)}"
                         + "".join(f"+{A('CP.' + c.lower(), m)}*{cm(c, 'H')}" for c in ["ENE", "MAN", "SRV", "MAA", "SGA", "GG", "HON"])
                         + "".join(f"+{A('CP.' + c.lower(), m)}*{pref('CP')}$C${CPT['lf_' + c]}" for c, _, _ in LOCAL_FLAGS))),
           RW("btrib", "Base del tributo de maquila", "USD",
              lambda m: f"=IF(p_BaseTrib=2,MAX({A('PV.vtas_prod', m)},{A('CP.va_nac', m)}),{A('PV.vtas_prod', m)})"),
           RW("trib", "Tributo único de maquila", "USD", lambda m: "=" + realor("TRIB", m, f"p_Trib*{A('CP.btrib', m)}")),
           S("Gastos de administración (USD)"),
           RW("alq", "Alquiler de nave (incluye IVA)", "USD",
              lambda m: "=" + realor("ALQ", m, f"p_M2*p_AlqM2*(1+p_IVAalq)*(1+p_AlqG)^({cal('anio_px', m)}-1)*{A('CP.afij', m)}")),
           RW("exp", "Expensas", "USD", lambda m: "=" + realor("EXP", m, f"p_M2*p_ExpM2*(1+p_AlqG)^({cal('anio_px', m)}-1)*{A('CP.afij', m)}")),
           RW("srv", "Agua, comunicaciones y electricidad", "USD", lambda m: "=" + realor("SRV", m, f"{usd_cm('SRV', m)}/12*{A('CP.ifij', m)}*{A('CP.afij', m)}")),
           RW("maa", "Mantenimiento y reparaciones (adm.)", "USD", lambda m: "=" + realor("MAA", m, f"{usd_cm('MAA', m)}/12*{A('CP.ifij', m)}*{A('CP.afij', m)}")),
           RW("sga", "Seguros sobre activos fijos", "USD", lambda m: "=" + realor("SGA", m, f"{usd_cm('SGA', m)}/12*{A('CP.ifij', m)}*{A('CP.afij', m)}")),
           RW("gg", "Gastos generales", "USD", lambda m: "=" + realor("GG", m, f"{usd_cm('GG', m)}/12*{A('CP.ifij', m)}*{A('CP.afij', m)}")),
           RW("hon", "Honorarios y otros fijos", "USD", lambda m: "=" + realor("HON", m, f"{usd_cm('HON', m)}/12*{A('CP.ifij', m)}*{A('CP.afij', m)}")),
           RW("impa", "Imprevistos de administración (base sin alquiler ni expensas)", "USD",
              lambda m: "=" + realor("IMPA", m, (f"p_ImpA*({A('CP.tot_a', m)}+{A('CP.srv', m)}+{A('CP.maa', m)}+{A('CP.sga', m)}+{A('CP.gg', m)}+{A('CP.hon', m)})"))),
           RW("extra", "Costos preoperativos / extraordinarios reales", "USD", lambda m: f"=IF({hv('freal', m)}=1,{RS('PREOP', m)},0)"),
           S("IVA contenido en costos y resúmenes (USD)"),
           RW("iva_cont", "IVA contenido en costos gravados identificados", "USD",
              lambda m: (f"={A('CP.alq', m)}*p_IVAalq/(1+p_IVAalq)"
                         + "".join(f"+{A('CP.' + c.lower(), m)}*{cm(c, 'G')}/(1+{cm(c, 'G')})" for c in ["ENE", "MAN", "SRV", "MAA", "SGA", "GG", "HON"]))),
           RW("c_prod", "COSTO DE PRODUCCIÓN INCURRIDO", "USD",
              lambda m: "=" + "+".join(A('CP.' + k, m) for k in ["mp", "ins", "con", "fle", "des", "pkg", "tot_d", "tot_i", "ene", "man", "impp"]), NF_USD, bold=True),
           RW("c_varv", "Gastos variables de venta y tributo", "USD", lambda m: f"={A('CP.segc', m)}+{A('CP.gvta', m)}+{A('CP.trib', m)}", NF_USD, bold=True),
           RW("c_adm", "GASTOS DE ADMINISTRACIÓN", "USD",
              lambda m: "=" + "+".join(A('CP.' + k, m) for k in ["tot_a", "alq", "exp", "srv", "maa", "sga", "gg", "hon", "impa"]), NF_USD, bold=True),
           RW("conv", "Pool de conversión fabril (MO directa, indirectos, energía, mant., packaging)", "USD",
              lambda m: f"={A('CP.tot_d', m)}+{A('CP.tot_i', m)}+{A('CP.ene', m)}+{A('CP.man', m)}+{A('CP.pkg', m)}"),
           RW("c_var", "Costos variables (equilibrio)", "USD",
              lambda m: "=" + "+".join(A('CP.' + k, m) for k in ["mp", "ins", "con", "fle", "des", "pkg", "impp", "segc", "trib", "gvta"])),
           RW("c_fij", "Costos fijos operativos (equilibrio)", "USD",
              lambda m: f"={A('CP.c_prod', m)}+{A('CP.c_varv', m)}+{A('CP.c_adm', m)}+{A('CP.extra', m)}-{A('CP.c_var', m)}"),
           RW("preop", "Memo: costos incurridos en meses de espera / implantación", "USD",
              lambda m: f"=IF({cal('f_pre', m)}=1,{A('CP.c_prod', m)}+{A('CP.c_varv', m)}+{A('CP.c_adm', m)}+{A('CP.extra', m)},0)"),
           RW("acc_men", "Devengado a pagar vía proveedores menores (incl. IVA)", "USD",
              lambda m: ("=" + "+".join(A('CP.' + k, m) for k in ["ene", "man", "pkg", "segc", "gvta", "impp", "trib", "exp", "srv", "maa", "sga", "gg", "hon", "impa", "extra"])
                         + f"+{A('CP.alq', m)}*IF(p_AlqCaja=3,0,1)+IF(p_C5=1,0,{A('CP.fle', m)}+{A('CP.des', m)})")),
           RW("alq_esp", "Alquiler como costo económico sin desembolso (aporte en especie)", "USD",
              lambda m: f"=IF(p_AlqCaja=3,{A('CP.alq', m)},0)"),
           ]
    return sp


# ------------------------------------------------------------------ Inversion_Activos (motor)
def ia_specs():
    f, l = IAT["first"], IAT["last"]
    T = lambda c: f"${c}${f}:${c}${l}"
    sp = [HDR(), S("CAPEX: pagos y devengamiento (USD)"),
          RW("pag_real", "CAPEX pagado real", "USD", lambda m: f"=IF({hv('freal', m)}=1,{RS('CAPEX_PAG', m)},0)"),
          RW("dev_real", "CAPEX devengado real", "USD", lambda m: f"=IF({hv('freal', m)}=1,{RS('CAPEX_DEV', m)},0)"),
          RW("pag_fc", "CAPEX pagos previstos (devengado no pagado + comprometido + por contratar)", "USD",
             lambda m: f"=IF({hv('freal', m)}=1,0,SUMIF({T('T')},{m},{T('X')}))"),
          RW("dev_fc", "CAPEX devengamiento previsto (comprometido + por contratar)", "USD",
             lambda m: f"=IF({hv('freal', m)}=1,0,SUMIF({T('T')},{m},{T('Y')}))"),
          RW("pag", "CAPEX PAGADO TOTAL", "USD", lambda m: f"={A('IA.pag_real', m)}+{A('IA.pag_fc', m)}", NF_USD, bold=True),
          RW("dev", "CAPEX DEVENGADO TOTAL (altas)", "USD", lambda m: f"={A('IA.dev_real', m)}+{A('IA.dev_fc', m)}", NF_USD, bold=True),
          RW("pag_mant", "CAPEX de mantenimiento / reposición (para CFADS)", "USD",
             lambda m: f'=IF({hv("freal", m)}=1,0,SUMIFS({T("X")},{T("T")},{m},{T("D")},"Mantenimiento/Reposición"))'),
          RW("cum_dev", "Activo fijo bruto acumulado (devengado)", "USD", lambda m: f"={prv('IA.cum_dev', m)}+{A('IA.dev', m)}", NF_USD, "last"),
          RW("cum_pag", "CAPEX pagado acumulado", "USD", lambda m: f"={prv('IA.cum_pag', m)}+{A('IA.pag', m)}", NF_USD, "last"),
          RW("cxp", "Cuentas por pagar de CAPEX", "USD", lambda m: f"={A('IA.cum_dev', m)}-{A('IA.cum_pag', m)}", NF_USD, "last"),
          S("Depreciación / amortización por ítem (desde la puesta en servicio; USD)")]
    for code in IAT["codes"]:
        rr = IAT[code]
        sp.append(RW(f"dep_{code}", f"Depreciación {code}", "USD",
                     lambda m, rr=rr: (f"=IF(AND({m}>=$U${rr},N($V${rr})>0,{m}<$U${rr}+N($V${rr})*12,{m}<=p_MesFin),"
                                       f"$O${rr}*(1-$W${rr})/(N($V${rr})*12),0)")))
    first_dep = f"IA.dep_{IAT['codes'][0]}"
    last_dep = f"IA.dep_{IAT['codes'][-1]}"
    sp += [S("Resumen de activo fijo"),
           RW("dep", "DEPRECIACIÓN Y AMORTIZACIÓN TOTAL", "USD",
              lambda m: f"=SUM({mc(m)}{ROWS[first_dep]}:{mc(m)}{ROWS[last_dep]})", NF_USD, bold=True),
           RW("dep_acum", "Depreciación acumulada", "USD", lambda m: f"={prv('IA.dep_acum', m)}+{A('IA.dep', m)}", NF_USD, "last"),
           RW("costo_serv", "Costo de activos en servicio", "USD", lambda m: f'=SUMIF({T("U")},"<="&{m},{T("O")})', NF_USD, "last"),
           RW("af_neto", "ACTIVO FIJO NETO (devengado - depreciación acumulada)", "USD", lambda m: f"={A('IA.cum_dev', m)}-{A('IA.dep_acum', m)}", NF_USD, "last", bold=True),
           RW("ctrl_dep", "Control: depreciación acumulada ≤ base depreciable en servicio (0 = OK)", "USD",
              lambda m: f"=MAX(0,{A('IA.dep_acum', m)}-{A('IA.costo_serv', m)})", NF_USD2, "max")]
    return sp


# ------------------------------------------------------------------ Capital_Trabajo
def ct_specs():
    fh = lambda m: cal("f_hor", m)
    return [
        HDR(), S("Clientes (USD)"),
        RW("vtas", "Ventas de productos", "USD", lambda m: f"={A('PV.vtas_prod', m)}"),
        RW("cred", "Ventas a crédito", "USD", lambda m: f"={A('CT.vtas', m)}*e_Cred"),
        RW("cont", "Ventas al contado", "USD", lambda m: f"={A('CT.vtas', m)}*e_Cont"),
        RW("antap", "Anticipos de clientes aplicados a ventas", "USD", lambda m: f"={A('CT.vtas', m)}*e_Ant"),
        RW("antrec", "Anticipos de clientes recibidos (ventas futuras)", "USD",
           lambda m: f"=IF({hv('freal', m)}=1,0,e_Ant*{fwd('CT.vtas', m, 'e_AntMeses')})*{fh(m)}"),
        RW("incob", "Incobrables (baja de cartera)", "USD", lambda m: f"={A('CT.cred', m)}*e_Incob"),
        RW("cobcred_fc", "Cobranza prevista de ventas a crédito (cohortes con plazo interpolado)", "USD",
           lambda m: f"=(1-e_Incob)*{lagf('CT.cred', m, 'p_LcobI', 'p_LcobF')}"),
        RW("cobaj", "Cobranza del ajuste de saldo real al corte", "USD",
           lambda m: f"=IF(AND(p_MesCorte>=0,{m}=p_MesCorte+MAX(1,ROUND(p_Lcob,0))),{prv('CT.ajcxc', m)},0)*{fh(m)}"),
        RW("cob", "COBRANZAS TOTALES", "USD",
           lambda m: "=" + realor("COBRO", m, f"({A('CT.cont', m)}+{A('CT.antrec', m)}+{A('CT.cobcred_fc', m)}+{A('CT.cobaj', m)})*{fh(m)}"), NF_USD, bold=True),
        RW("cobcred", "Cobranza aplicada a cuentas por cobrar", "USD",
           lambda m: f"={A('CT.cob', m)}-{A('CT.cont', m)}-{A('CT.antrec', m)}-{A('CT.cobaj', m)}"),
        RW("cxc_calc", "Cuentas por cobrar calculadas", "USD",
           lambda m: f"={prv('CT.cxc', m)}+{A('CT.cred', m)}-{A('CT.incob', m)}-{A('CT.cobcred', m)}-{A('CT.cobaj', m)}", NF_USD, "none"),
        RW("cxc", "CUENTAS POR COBRAR (saldo)", "USD", lambda m: "=" + over(m, "r_S_CXC", A('CT.cxc_calc', m)), NF_USD, "last", bold=True),
        RW("cxc_dif", "Ajuste por saldo real al corte", "USD", lambda m: f"={A('CT.cxc', m)}-{A('CT.cxc_calc', m)}"),
        RW("ajcxc", "Ajuste pendiente de cobro", "USD", lambda m: f"={prv('CT.ajcxc', m)}+{A('CT.cxc_dif', m)}-{A('CT.cobaj', m)}", NF_USD, "last"),
        RW("antcli", "Anticipos de clientes (pasivo)", "USD", lambda m: f"={prv('CT.antcli', m)}+{A('CT.antrec', m)}-{A('CT.antap', m)}", NF_USD, "last"),
        S("Inventario de materiales (USD, valuado según C5)"),
        RW("cons", "Consumo de materiales", "USD", lambda m: f"={A('CP.mat_land', m)}"),
        RW("stk_obj", "Stock objetivo = consumo de los próximos días de stock", "USD",
           lambda m: (f"=IF(p_Sk>0,SUM(INDEX({RG('CT.cons')},1,MIN({NC},{m}+2)):INDEX({RG('CT.cons')},1,MIN({NC},{m}+1+p_Sk))),0)"
                      f"+p_Sf*INDEX({RG('CT.cons')},1,MIN({NC},{m}+2+p_Sk))"), NF_USD, "none"),
        RW("compras", "Compras de materiales", "USD",
           lambda m: "=" + realor("COMPRA_MAT", m, f"MAX(0,{A('CT.cons', m)}+{A('CT.stk_obj', m)}-{prv('CT.stk', m)})"), NF_USD, bold=True),
        RW("stk_calc", "Inventario de materiales calculado", "USD", lambda m: f"={prv('CT.stk', m)}+{A('CT.compras', m)}-{A('CT.cons', m)}", NF_USD, "none"),
        RW("stk", "INVENTARIO DE MATERIALES (saldo)", "USD", lambda m: "=" + over(m, "r_S_INVMAT", A('CT.stk_calc', m)), NF_USD, "last", bold=True),
        RW("stk_dif", "Ajuste por saldo real al corte", "USD", lambda m: f"={A('CT.stk', m)}-{A('CT.stk_calc', m)}"),
        S("Proveedores de materiales (USD)"),
        RW("adv", "Anticipos pagados a proveedores (compras futuras)", "USD",
           lambda m: f"=IF({hv('freal', m)}=1,0,e_AntProv*{fwd('CT.compras', m, 'e_AntProvMeses')})*{fh(m)}"),
        RW("advap", "Anticipos aplicados a compras del mes", "USD", lambda m: f"=IF({hv('freal', m)}=1,0,e_AntProv*{A('CT.compras', m)})"),
        RW("antprov", "Anticipos a proveedores (activo)", "USD", lambda m: f"={prv('CT.antprov', m)}+{A('CT.adv', m)}-{A('CT.advap', m)}", NF_USD, "last"),
        RW("accmp", "Compras a crédito (netas de anticipos)", "USD", lambda m: f"={A('CT.compras', m)}-{A('CT.advap', m)}"),
        RW("pagmp_fc", "Pago previsto de compras (plazo interpolado)", "USD", lambda m: f"={lagf('CT.accmp', m, 'p_LmpI', 'p_LmpF')}"),
        RW("pagaj_mp", "Pago del ajuste de saldo real al corte", "USD",
           lambda m: f"=IF(AND(p_MesCorte>=0,{m}=p_MesCorte+MAX(1,ROUND(p_Lmp,0))),{prv('CT.ajmp', m)},0)*{fh(m)}"),
        RW("pag_mat", "PAGOS A PROVEEDORES DE MATERIALES", "USD",
           lambda m: "=" + realor("PAGO_MAT", m, f"({A('CT.adv', m)}+{A('CT.pagmp_fc', m)}+{A('CT.pagaj_mp', m)})*{fh(m)}"), NF_USD, bold=True),
        RW("prov_calc", "Proveedores de materiales calculados", "USD",
           lambda m: f"={prv('CT.prov', m)}+{A('CT.accmp', m)}-({A('CT.pag_mat', m)}-{A('CT.adv', m)})", NF_USD, "none"),
        RW("prov", "PROVEEDORES DE MATERIALES (saldo)", "USD", lambda m: "=" + over(m, "r_S_PROVMAT", A('CT.prov_calc', m)), NF_USD, "last", bold=True),
        RW("prov_dif", "Ajuste por saldo real al corte", "USD", lambda m: f"={A('CT.prov', m)}-{A('CT.prov_calc', m)}"),
        RW("ajmp", "Ajuste pendiente de pago", "USD", lambda m: f"={prv('CT.ajmp', m)}+{A('CT.prov_dif', m)}-{A('CT.pagaj_mp', m)}", NF_USD, "last"),
        S("Proveedores menores y obligaciones (USD)"),
        RW("accmen", "Devengado a pagar (costos no salariales ni de materiales, incl. IVA y tributo)", "USD", lambda m: f"={A('CP.acc_men', m)}"),
        RW("pagmen_fc", "Pago previsto (plazo interpolado)", "USD", lambda m: f"={lagf('CT.accmen', m, 'p_LmenI', 'p_LmenF')}"),
        RW("pagaj_men", "Pago del ajuste de saldo real al corte", "USD",
           lambda m: f"=IF(AND(p_MesCorte>=0,{m}=p_MesCorte+MAX(1,ROUND(p_Lmen,0))),{prv('CT.ajmen', m)},0)*{fh(m)}"),
        RW("pag_men", "PAGOS A PROVEEDORES MENORES", "USD",
           lambda m: "=" + realor("PAGO_MEN", m, f"({A('CT.pagmen_fc', m)}+{A('CT.pagaj_men', m)})*{fh(m)}"), NF_USD, bold=True),
        RW("pmen_calc", "Proveedores menores calculados", "USD", lambda m: f"={prv('CT.pmen', m)}+{A('CT.accmen', m)}-{A('CT.pag_men', m)}", NF_USD, "none"),
        RW("pmen", "PROVEEDORES MENORES (saldo)", "USD", lambda m: "=" + over(m, "r_S_PROVMEN", A('CT.pmen_calc', m)), NF_USD, "last", bold=True),
        RW("pmen_dif", "Ajuste por saldo real al corte", "USD", lambda m: f"={A('CT.pmen', m)}-{A('CT.pmen_calc', m)}"),
        RW("ajmen", "Ajuste pendiente de pago", "USD", lambda m: f"={prv('CT.ajmen', m)}+{A('CT.pmen_dif', m)}-{A('CT.pagaj_men', m)}", NF_USD, "last"),
        S("Producto terminado y anticipos al personal (USD)"),
        RW("pt_calc", "PT calculado: materiales por unidad + conversión asignada por driver", "USD",
           lambda m: (f"={A('PV.ptu_g', m)}*{A('CP.umat_g', m)}+{A('PV.ptu_c', m)}*{A('CP.umat_c', m)}"
                      f"+IF({A('PV.prod_g', m)}*p_WG+{A('PV.prod_c', m)}*p_WC>0,{A('CP.conv', m)}*({A('PV.ptu_g', m)}*p_WG+{A('PV.ptu_c', m)}*p_WC)"
                      f"/({A('PV.prod_g', m)}*p_WG+{A('PV.prod_c', m)}*p_WC),0)"), NF_USD, "none"),
        RW("pt", "INVENTARIO DE PRODUCTO TERMINADO (saldo)", "USD", lambda m: "=" + over(m, "r_S_INVPT", A('CT.pt_calc', m)), NF_USD, "last", bold=True),
        RW("pt_dif", "Ajuste por saldo real al corte (se reconoce en el costo de ventas vía ΔPT)", "USD", lambda m: f"={A('CT.pt', m)}-{A('CT.pt_calc', m)}"),
        RW("antpers", "Anticipos al personal (activo)", "USD",
           lambda m: f"=IF({fh(m)}=1,{A('CP.rem_tot', m)}*e_DiasAntPers/p_DiasMes,{prv('CT.antpers', m)})", NF_USD, "last"),
        S("Capital de trabajo operativo neto (USD; excluye caja mínima y deuda)"),
        RW("ct_act", "Activo corriente operativo", "USD",
           lambda m: f"={A('CT.cxc', m)}+{A('CT.stk', m)}+{A('CT.pt', m)}+{A('CT.antprov', m)}+{A('CT.antpers', m)}+{A('DT.iva', m)}", NF_USD, "last"),
        RW("ct_pas", "Pasivo corriente operativo", "USD", lambda m: f"={A('CT.prov', m)}+{A('CT.pmen', m)}+{A('CT.antcli', m)}", NF_USD, "last"),
        RW("ctn", "CAPITAL DE TRABAJO OPERATIVO NETO", "USD", lambda m: f"={A('CT.ct_act', m)}-{A('CT.ct_pas', m)}", NF_USD, "max", bold=True),
        RW("dctn", "Variación del capital de trabajo operativo", "USD", lambda m: f"={A('CT.ctn', m)}-{prv('CT.ctn', m)}"),
    ]


# ------------------------------------------------------------------ Deuda_Tributos
def loan_specs(n, monto, mes, tasa, gracia, cuotas, com, code_d, code_i, code_a, sname):
    L = f"l{n}"
    K = lambda s: f"DT.{L}_{s}"
    return [
        S(f"Préstamo {n} (USD; tasa mensual = nominal anual / 12)"),
        RW(f"{L}_k", "Mes relativo al desembolso", "n°", lambda m: f"={m}-{mes}", NF_INT0, "none"),
        RW(f"{L}_th", "Saldo contractual teórico", "USD",
           lambda m: f"=IF({A(K('k'), m)}<0,0,IF({A(K('k'), m)}=0,{monto},{prv(K('th'), m)}-{A(K('tham'), m)}))", NF_USD, "none"),
        RW(f"{L}_tham", "Amortización contractual teórica (francés)", "USD",
           lambda m: (f"=IF(AND({A(K('k'), m)}>{gracia},{A(K('k'), m)}<={gracia}+{cuotas},{cuotas}>0),"
                      f"PMT({tasa}/12,{cuotas},-{monto})-{prv(K('th'), m)}*{tasa}/12,0)"), NF_USD),
        RW(f"{L}_desemb", "Desembolso", "USD", lambda m: "=" + realor(code_d, m, f"IF({A(K('k'), m)}=0,{monto},0)")),
        RW(f"{L}_sini", "Saldo inicial", "USD", lambda m: f"={prv(K('sal'), m)}", NF_USD, "none"),
        RW(f"{L}_int", "Intereses (sin IVA)", "USD", lambda m: "=" + realor(code_i, m, f"{A(K('sini'), m)}*{tasa}/12*{cal('f_hor', m)}")),
        RW(f"{L}_am", "Amortización de capital", "USD",
           lambda m: "=" + realor(code_a, m, (f"IF({A(K('k'), m)}<={gracia},0,IF({A(K('k'), m)}>={gracia}+{cuotas},{A(K('sini'), m)},"
                                              f"MIN({A(K('sini'), m)},{A(K('tham'), m)})))*{cal('f_hor', m)}"))),
        RW(f"{L}_iva", "IVA sobre intereses (costo)", "USD", lambda m: f"={A(K('int'), m)}*p_IVAint"),
        RW(f"{L}_com", "Comisión inicial", "USD", lambda m: f"=IF({A(K('k'), m)}=0,{monto}*{com},0)"),
        RW(f"{L}_scalc", "Saldo final calculado", "USD", lambda m: f"={A(K('sini'), m)}+{A(K('desemb'), m)}-{A(K('am'), m)}", NF_USD, "none"),
        RW(f"{L}_sal", "SALDO FINAL", "USD", lambda m: "=" + over(m, sname, A(K('scalc'), m)), NF_USD, "last", bold=True),
        RW(f"{L}_dif", "Ajuste por saldo real al corte", "USD", lambda m: f"={A(K('sal'), m)}-{A(K('scalc'), m)}"),
        RW(f"{L}_neg", "Control: saldo negativo (0 = OK)", "USD", lambda m: f"=MIN(0,{A(K('sal'), m)})", NF_USD2, "min"),
    ]


def dt_specs():
    fh = lambda m: cal("f_hor", m)
    sp = [HDR(), S("IVA soportado recuperable (USD)"),
          RW("ivacont", "IVA contenido en costos gravados identificados", "USD", lambda m: f"={A('CP.iva_cont', m)}"),
          RW("ivarec", "IVA recuperable devengado (crédito fiscal)", "USD",
             lambda m: (f"=IF(p_Modalidad=2,MIN({A('DT.ivacont', m)}*e_IVArec,p_TopeIVAServ*{A('PV.vtas_prod', m)}),"
                        f"{A('DT.ivacont', m)}*e_IVArec)")),
          RW("ivacob_fc", "Recuperación prevista (con desfase)", "USD",
             lambda m: f"=IF({m}-e_IVAlag>=0,INDEX({RG('DT.ivarec')},1,{m}-e_IVAlag+1),0)"),
          RW("ivaaj", "Recuperación del ajuste de saldo real al corte", "USD",
             lambda m: f"=IF(AND(p_MesCorte>=0,{m}=p_MesCorte+MAX(1,e_IVAlag)),{prv('DT.ivaajp', m)},0)*{fh(m)}"),
          RW("ivacob", "IVA RECUPERADO (cobro)", "USD",
             lambda m: "=" + realor("IVA_RECUP", m, f"({A('DT.ivacob_fc', m)}+{A('DT.ivaaj', m)})*{fh(m)}"), NF_USD, bold=True),
          RW("iva_calc", "Crédito fiscal calculado", "USD", lambda m: f"={prv('DT.iva', m)}+{A('DT.ivarec', m)}-{A('DT.ivacob', m)}", NF_USD, "none"),
          RW("iva", "CRÉDITO FISCAL IVA (saldo; requiere financiamiento hasta recuperarse)", "USD",
             lambda m: "=" + over(m, "r_S_IVA", A('DT.iva_calc', m)), NF_USD, "max", bold=True),
          RW("iva_dif", "Ajuste por saldo real al corte", "USD", lambda m: f"={A('DT.iva', m)}-{A('DT.iva_calc', m)}"),
          RW("ivaajp", "Ajuste pendiente de recuperar", "USD", lambda m: f"={prv('DT.ivaajp', m)}+{A('DT.iva_dif', m)}-{A('DT.ivaaj', m)}", NF_USD, "last")]
    sp += loan_specs(1, "p_L1_Monto", "p_L1_Mes", "p_L1_Tasa", "p_L1_Gracia", "p_L1_Cuotas", "p_L1_Com", "DESEMB_P1", "INT_P1", "AMORT_P1", "r_S_DEUDA1")
    sp += loan_specs(2, "e_L2_Monto", "e_L2_Mes", "e_L2_Tasa", "e_L2_Gracia", "e_L2_Cuotas", "0", "DESEMB_P2", "INT_P2", "AMORT_P2", "r_S_DEUDA2")
    sp += [S("Aportes de socios en efectivo (USD)"),
           RW("ap_ini", "Aporte propio inicial (plan)", "USD", lambda m: f"=IF({m}=0,Base_PDF!$E${BP['aporte_ini']},0)"),
           RW("ap_pdf", "Aportes adicionales del plan PDF (inicio de años 1 y 2)", "USD",
              lambda m: (f"=e_AportesPDF*(IF({m}=p_ImplMeses+1,Base_PDF!$E${BP['aporte_a1']},0)"
                         f"+IF({m}=p_ImplMeses+13,Base_PDF!$E${BP['aporte_a2']},0))")),
           RW("ap_plan", "Aportes planificados adicionales (Supuestos F)", "USD", lambda m: f"=SUMIF(p_ApMes,{m},p_ApMonto)"),
           RW("aportes", "APORTES EN EFECTIVO", "USD",
              lambda m: "=" + realor("APORTE", m, f"{A('DT.ap_ini', m)}+{A('DT.ap_pdf', m)}+{A('DT.ap_plan', m)}"), NF_USD, bold=True)]
    return sp


# ------------------------------------------------------------------ Resultados_Caja
def rc_specs():
    fh = lambda m: cal("f_hor", m)
    a = lambda k, m: A("RC." + k, m)
    return [
        HDR(), S("Estado de resultados (USD)"),
        RW("v_g", "Ventas gabinetes", "USD", lambda m: f"={A('PV.vtas_g', m)}"),
        RW("v_c", "Ventas cajas", "USD", lambda m: f"={A('PV.vtas_c', m)}"),
        RW("v_ch", "Venta de chatarra", "USD", lambda m: f"={A('PV.chat', m)}"),
        RW("v", "VENTAS NETAS", "USD", lambda m: f"={a('v_g', m)}+{a('v_c', m)}+{a('v_ch', m)}", NF_USD, bold=True),
        RW("c_mat", "Materia prima, insumos y consumibles", "USD", lambda m: f"={A('CP.mp', m)}+{A('CP.ins', m)}+{A('CP.con', m)}"),
        RW("c_fle", "Flete de materia prima", "USD", lambda m: f"={A('CP.fle', m)}"),
        RW("c_des", "Despachos de importación", "USD", lambda m: f"={A('CP.des', m)}"),
        RW("c_pkg", "Packaging", "USD", lambda m: f"={A('CP.pkg', m)}"),
        RW("c_mod", "Personal directo (con cargas y alimentación)", "USD", lambda m: f"={A('CP.tot_d', m)}"),
        RW("c_moi", "Personal indirecto de fábrica", "USD", lambda m: f"={A('CP.tot_i', m)}"),
        RW("c_ene", "Energía de planta", "USD", lambda m: f"={A('CP.ene', m)}"),
        RW("c_man", "Mantenimiento de planta", "USD", lambda m: f"={A('CP.man', m)}"),
        RW("c_imp", "Imprevistos de producción", "USD", lambda m: f"={A('CP.impp', m)}"),
        RW("c_prod", "Costo de producción incurrido", "USD", lambda m: f"={A('CP.c_prod', m)}", NF_USD, bold=True),
        RW("d_pt", "(-) Aumento del inventario de producto terminado", "USD", lambda m: f"={A('CT.pt', m)}-{prv('CT.pt', m)}"),
        RW("cv", "COSTO DE VENTAS", "USD", lambda m: f"={a('c_prod', m)}-{a('d_pt', m)}", NF_USD, bold=True),
        RW("mb", "MARGEN BRUTO", "USD", lambda m: f"={a('v', m)}-{a('cv', m)}", NF_USD, bold=True),
        RW("g_trib", "Tributo único de maquila (una sola vez, dentro del EBITDA)", "USD", lambda m: f"={A('CP.trib', m)}"),
        RW("g_segc", "Seguro de caución", "USD", lambda m: f"={A('CP.segc', m)}"),
        RW("g_gvta", "Gastos de venta y logística de exportación", "USD", lambda m: f"={A('CP.gvta', m)}"),
        RW("g_incob", "Incobrables", "USD", lambda m: f"={A('CT.incob', m)}"),
        RW("g_var", "Gastos variables de venta y tributo", "USD",
           lambda m: f"={a('g_trib', m)}+{a('g_segc', m)}+{a('g_gvta', m)}+{a('g_incob', m)}", NF_USD, bold=True),
        RW("g_adm", "Gastos de administración (incl. alquiler)", "USD", lambda m: f"={A('CP.c_adm', m)}", NF_USD, bold=True),
        RW("g_alq", "   de los cuales: alquiler de nave (incl. IVA)", "USD", lambda m: f"={A('CP.alq', m)}"),
        RW("g_extra", "Costos preoperativos / extraordinarios reales", "USD", lambda m: f"={A('CP.extra', m)}"),
        RW("iva_rec", "(+) IVA soportado recuperable reclasificado a crédito fiscal", "USD", lambda m: f"={A('DT.ivarec', m)}"),
        RW("ebitda", "EBITDA", "USD", lambda m: f"={a('mb', m)}-{a('g_var', m)}-{a('g_adm', m)}-{a('g_extra', m)}+{a('iva_rec', m)}", NF_USD, bold=True),
        RW("dep", "Depreciación y amortización", "USD", lambda m: f"={A('IA.dep', m)}"),
        RW("ebit", "EBIT (resultado operativo)", "USD", lambda m: f"={a('ebitda', m)}-{a('dep', m)}", NF_USD, bold=True),
        RW("f_int", "Intereses de préstamos", "USD", lambda m: f"={A('DT.l1_int', m)}+{A('DT.l2_int', m)}"),
        RW("f_lin", "Intereses de la línea hipotética", "USD", lambda m: f"={a('lin_int', m)}"),
        RW("f_iva", "IVA sobre intereses (no recuperable)", "USD", lambda m: f"={A('DT.l1_iva', m)}+{A('DT.l2_iva', m)}+{a('lin_iva', m)}"),
        RW("f_com", "Comisiones", "USD", lambda m: f"={A('DT.l1_com', m)}+{A('DT.l2_com', m)}"),
        RW("gfin", "Gastos financieros", "USD", lambda m: f"={a('f_int', m)}+{a('f_lin', m)}+{a('f_iva', m)}+{a('f_com', m)}", NF_USD, bold=True),
        RW("ebt", "Resultado antes de IRE", "USD", lambda m: f"={a('ebit', m)}-{a('gfin', m)}"),
        RW("ire", "IRE (parámetro; 0% según el plan)", "USD", lambda m: f"=p_IRE*{a('ebt', m)}"),
        RW("ni", "RESULTADO NETO", "USD", lambda m: f"={a('ebt', m)}-{a('ire', m)}", NF_USD, bold=True),
        RW("m_preop", "Memo: costos de espera / implantación incluidos arriba", "USD", lambda m: f"={A('CP.preop', m)}"),
        S("Flujo de caja — método directo (USD)"),
        RW("cob", "Cobranzas de clientes", "USD", lambda m: f"={A('CT.cob', m)}"),
        RW("cob_ch", "Cobro de chatarra", "USD", lambda m: f"={A('PV.chat', m)}*{fh(m)}"),
        RW("ivacob", "Recuperación de IVA", "USD", lambda m: f"={A('DT.ivacob', m)}"),
        RW("pag_mat", "Pagos a proveedores de materiales", "USD", lambda m: f"={A('CT.pag_mat', m)}"),
        RW("pag_men", "Pagos a proveedores menores, servicios, alquiler y tributo", "USD", lambda m: f"={A('CT.pag_men', m)}"),
        RW("pag_nom", "Pagos de nómina (incl. variación de anticipos al personal)", "USD",
           lambda m: f"=({A('CP.per_tot', m)}+{A('CT.antpers', m)}-{prv('CT.antpers', m)})*{fh(m)}"),
        RW("ire_pag", "IRE pagado", "USD", lambda m: f"={a('ire', m)}*{fh(m)}"),
        RW("cfo", "FLUJO OPERATIVO (antes de financiamiento)", "USD",
           lambda m: (f"={a('cob', m)}+{a('cob_ch', m)}+{a('ivacob', m)}-{a('pag_mat', m)}-{a('pag_men', m)}-{a('pag_nom', m)}-{a('ire_pag', m)}"),
           NF_USD, bold=True),
        RW("capex", "CAPEX pagado", "USD", lambda m: f"={A('IA.pag', m)}"),
        RW("alq_esp", "Alquiler económico sin desembolso (solo valuación)", "USD", lambda m: f"={A('CP.alq_esp', m)}"),
        RW("fcff_op", "FCFF antes de valor terminal", "USD", lambda m: f"={a('cfo', m)}-{a('capex', m)}-{a('alq_esp', m)}", NF_USD, bold=True),
        RW("vt", "Valor terminal del proyecto (mes final)", "USD", lambda m: f"=IF({cal('f_fin', m)}=1,{a('tv', m)},0)"),
        RW("fcff", "FLUJO LIBRE DEL PROYECTO (FCFF)", "USD", lambda m: f"={a('fcff_op', m)}+{a('vt', m)}", NF_USD, bold=True),
        RW("ind_chk", "Control: método indirecto (EBITDA - ΔCT - IRE) menos directo (0 = OK)", "USD",
           lambda m: (f"=ROUND(({a('ebitda', m)}+{a('alq_esp', m)}-{a('ire', m)}-({A('CT.ctn', m)}-{prv('CT.ctn', m)})"
                      f"+({A('CT.cxc_dif', m)}+{A('CT.stk_dif', m)}+{A('DT.iva_dif', m)}-{A('CT.prov_dif', m)}-{A('CT.pmen_dif', m)}))*{fh(m)}"
                      f"-{a('cfo', m)},2)"), NF_USD2, "max"),
        S("Financiamiento y tesorería (USD)"),
        RW("desemb", "Desembolsos de préstamos contratados", "USD", lambda m: f"={A('DT.l1_desemb', m)}+{A('DT.l2_desemb', m)}"),
        RW("amort", "Amortización de préstamos", "USD", lambda m: f"={A('DT.l1_am', m)}+{A('DT.l2_am', m)}"),
        RW("int_pag", "Intereses, IVA y comisiones pagados", "USD",
           lambda m: f"={A('DT.l1_int', m)}+{A('DT.l2_int', m)}+{A('DT.l1_iva', m)}+{A('DT.l2_iva', m)}+{A('DT.l1_com', m)}+{A('DT.l2_com', m)}"),
        RW("aportes", "Aportes de socios en efectivo", "USD", lambda m: f"={A('DT.aportes', m)}"),
        RW("caja_ini", "Caja inicial", "USD", lambda m: f"={prv('RC.caja', m)}", NF_USD, "none"),
        RW("lin_ini", "Línea hipotética: saldo inicial", "USD", lambda m: f"={prv('RC.lin', m)}", NF_USD, "none"),
        RW("lin_int", "Línea hipotética: intereses", "USD", lambda m: f"={a('lin_ini', m)}*e_TasaLinea/12*{fh(m)}"),
        RW("lin_iva", "Línea hipotética: IVA sobre intereses", "USD", lambda m: f"={a('lin_int', m)}*p_IVAint"),
        RW("x1", "Caja antes de dividendos y línea", "USD",
           lambda m: (f"={a('caja_ini', m)}+{a('cfo', m)}-{a('capex', m)}+{a('desemb', m)}-{a('amort', m)}-{a('int_pag', m)}+{a('aportes', m)}"
                      f"-{a('lin_int', m)}-{a('lin_iva', m)}"), NF_USD, "min"),
        RW("cmin", "Caja mínima operativa", "USD", lambda m: f"=({a('v', m)}*p_CajaMinDias/p_DiasMes+p_CajaMinFija)*{fh(m)}", NF_USD, "max"),
        RW("lin_rep", "Línea hipotética: cancelación", "USD", lambda m: f"=IF(e_Linea=1,MIN({a('lin_ini', m)},MAX(0,{a('x1', m)}-{a('cmin', m)})),0)"),
        RW("x2", "Caja disponible para dividendos", "USD", lambda m: f"={a('x1', m)}-{a('lin_rep', m)}", NF_USD, "none"),
        RW("ytd", "Resultado neto acumulado del año de proyecto", "USD",
           lambda m: f"=IF(OR({m}=0,{cal('ini', m)}=1),{a('ni', m)},{prv('RC.ytd', m)}+{a('ni', m)})", NF_USD, "none"),
        RW("decl", "Dividendos declarados al cierre del año (política x resultado, con límite de resultados acumulados)", "USD",
           lambda m: (f"=IF({cal('cierre', m)}=1,MAX(0,MIN(INDEX(p_Payout,MIN({NYEARS},{cal('anio', m)}))*{a('ytd', m)},"
                      f"{prv('RC.re', m)}+{a('ni', m)})),0)")),
        RW("re", "Resultados acumulados no distribuidos", "USD", lambda m: f"={prv('RC.re', m)}+{a('ni', m)}-{a('decl', m)}", NF_USD, "last"),
        RW("tesp", "Meses desde la declaración", "n°",
           lambda m: f"=IF({a('decl', m)}>0,0,IF({prv('RC.esp', m)}>0,{prv('RC.tesp', m)}+1,0))", NF_INT0, "none"),
        RW("hab", "Dividendos habilitados para pago (después del desfase)", "USD",
           lambda m: f"=IF(AND({prv('RC.esp', m)}>0,{a('tesp', m)}>=MAX(1,p_DivLag)),{prv('RC.esp', m)},0)"),
        RW("esp", "Dividendos declarados en espera", "USD", lambda m: f"={prv('RC.esp', m)}+{a('decl', m)}-{a('hab', m)}", NF_USD, "last"),
        RW("divpag", "DIVIDENDOS PAGADOS (solo con caja sobre el mínimo)", "USD",
           lambda m: "=" + realor("DIVIDENDO", m, (f"IF(p_DivRestr=1,MIN({prv('RC.dpp', m)}+{a('hab', m)},MAX(0,{a('x2', m)}-{a('cmin', m)})),"
                                                    f"{prv('RC.dpp', m)}+{a('hab', m)})*{fh(m)}")), NF_USD, bold=True),
        RW("dpp", "Dividendos por pagar", "USD", lambda m: f"={prv('RC.dpp', m)}+{a('hab', m)}-{a('divpag', m)}", NF_USD, "last"),
        RW("x3", "Caja antes de la línea hipotética", "USD", lambda m: f"={a('x2', m)}-{a('divpag', m)}", NF_USD, "min"),
        RW("brecha", "NECESIDAD DE FINANCIAMIENTO del mes (caja mínima - caja disponible)", "USD",
           lambda m: f"=MAX(0,{a('cmin', m)}-{a('x3', m)})*{fh(m)}", NF_USD, "max", bold=True),
        RW("lin_uso", "Línea hipotética: uso (NO APROBADA)", "USD",
           lambda m: f"=IF(e_Linea=1,MIN(MAX(0,e_Cupo-({a('lin_ini', m)}-{a('lin_rep', m)})),{a('brecha', m)}),0)"),
        RW("lin", "Línea hipotética: saldo", "USD", lambda m: f"={a('lin_ini', m)}-{a('lin_rep', m)}+{a('lin_uso', m)}", NF_USD, "max"),
        RW("ap_auto", "Aportes hipotéticos de cobertura de brecha (NO APROBADOS)", "USD",
           lambda m: f"=IF(p_ApAuto=1,MAX(0,{a('cmin', m)}-{a('x3', m)}-{a('lin_uso', m)}),0)*{fh(m)}", NF_USD, "sum"),
        RW("caja_calc", "Caja final calculada", "USD", lambda m: f"={a('x3', m)}+{a('lin_uso', m)}+{a('ap_auto', m)}", NF_USD, "none"),
        RW("caja", "CAJA FINAL", "USD", lambda m: "=" + over(m, "r_S_CAJA", a('caja_calc', m)), NF_USD, "min", bold=True),
        RW("caja_dif", "Ajuste por saldo real de caja al corte", "USD", lambda m: f"={a('caja', m)}-{a('caja_calc', m)}"),
        RW("brecha_nc", "BRECHA NO CUBIERTA (caja final bajo el mínimo)", "USD", lambda m: f"=MAX(0,{a('cmin', m)}-{a('caja', m)})*{fh(m)}", NF_USD, "max", bold=True),
        RW("cum_pre", "Caja acumulada antes de financiamiento adicional (solo préstamo original y aporte inicial)", "USD",
           lambda m: (f"={prv('RC.cum_pre', m)}+({a('cfo', m)}-{a('capex', m)}+{A('DT.l1_desemb', m)}-{A('DT.l1_am', m)}-{A('DT.l1_int', m)}"
                      f"-{A('DT.l1_iva', m)}-{A('DT.l1_com', m)}+{A('DT.ap_ini', m)})*{fh(m)}"), NF_USD, "min"),
        RW("nec_pre", "Necesidad de fondos adicionales acumulada (incluye caja mínima)", "USD",
           lambda m: f"=MAX(0,{a('cmin', m)}-{a('cum_pre', m)})*{fh(m)}", NF_USD, "max", bold=True),
        RW("deuda", "Deuda financiera total", "USD", lambda m: f"={A('DT.l1_sal', m)}+{A('DT.l2_sal', m)}+{a('lin', m)}", NF_USD, "max"),
        S("Balance de control (USD) — debe cuadrar todos los meses"),
        RW("b_act", "Activo total", "USD",
           lambda m: (f"={a('caja', m)}+{A('CT.cxc', m)}+{A('CT.stk', m)}+{A('CT.pt', m)}+{A('CT.antprov', m)}+{A('CT.antpers', m)}"
                      f"+{A('DT.iva', m)}+{A('IA.af_neto', m)}"), NF_USD, "last"),
        RW("b_pas", "Pasivo total", "USD",
           lambda m: (f"={A('CT.prov', m)}+{A('CT.pmen', m)}+{A('CT.antcli', m)}+{A('IA.cxp', m)}+{A('DT.l1_sal', m)}+{A('DT.l2_sal', m)}"
                      f"+{a('lin', m)}+{a('esp', m)}+{a('dpp', m)}"), NF_USD, "last"),
        RW("b_cap", "Aportes acumulados (efectivo y especie)", "USD", lambda m: f"={prv('RC.b_cap', m)}+{a('aportes', m)}+{a('ap_auto', m)}+{a('alq_esp', m)}", NF_USD, "last"),
        RW("b_ni", "Resultados acumulados (antes de dividendos)", "USD", lambda m: f"={prv('RC.b_ni', m)}+{a('ni', m)}", NF_USD, "last"),
        RW("b_div", "Dividendos declarados acumulados", "USD", lambda m: f"={prv('RC.b_div', m)}+{a('decl', m)}", NF_USD, "last"),
        RW("b_aj", "Ajustes de conciliación con saldos reales (a patrimonio; el de PT va a costo de ventas)", "USD",
           lambda m: (f"={prv('RC.b_aj', m)}+{a('caja_dif', m)}+{A('CT.cxc_dif', m)}+{A('CT.stk_dif', m)}+{A('DT.iva_dif', m)}"
                      f"-{A('CT.prov_dif', m)}-{A('CT.pmen_dif', m)}-{A('DT.l1_dif', m)}-{A('DT.l2_dif', m)}"), NF_USD, "last"),
        RW("b_pat", "Patrimonio", "USD", lambda m: f"={a('b_cap', m)}+{a('b_ni', m)}-{a('b_div', m)}+{a('b_aj', m)}", NF_USD, "last"),
        RW("b_chk", "CONTROL: activo - pasivo - patrimonio (0 = OK)", "USD", lambda m: f"=ROUND({a('b_act', m)}-{a('b_pas', m)}-{a('b_pat', m)},2)", NF_USD2, "max", bold=True),
        RW("b_chk_min", "Control (mínimo)", "USD", lambda m: f"={a('b_chk', m)}", NF_USD2, "min"),
        S("Valuación (USD)"),
        RW("df_p", "Factor de descuento del proyecto", "factor", lambda m: f"={cal('df_p', m)}", NF_USD4, "none"),
        RW("df_e", "Factor de descuento Ke", "factor", lambda m: f"={cal('df_e', m)}", NF_USD4, "none"),
        RW("tv_liq", "Valor de liquidación de activos operativos netos (si fuera el mes final)", "USD",
           lambda m: (f"={A('CT.cxc', m)}*p_RealCxC+({A('CT.stk', m)}+{A('CT.pt', m)})*p_RealInv+{A('CT.antprov', m)}+{A('CT.antpers', m)}"
                      f"+{A('DT.iva', m)}-{A('CT.prov', m)}-{A('CT.pmen', m)}-{A('CT.antcli', m)}-{A('IA.cxp', m)}+{A('IA.af_neto', m)}*p_RealAF-p_CostoCierre"),
           NF_USD, "none"),
        RW("tv_gc", "Valor de negocio en marcha (perpetuidad del EBITDA de 12 meses - reposición - CT incremental)", "USD",
           lambda m: (f"=IF(p_TasaProy>p_g,({rsum('RC.ebitda', m)}-{rsum('RC.dep', m)}-p_g*{A('CT.ctn', m)})*(1+p_g)/(p_TasaProy-p_g),0)"),
           NF_USD, "none"),
        RW("tv", "Valor terminal según método elegido", "USD", lambda m: f"=IF(p_TipoVT=1,{a('tv_liq', m)},{a('tv_gc', m)})", NF_USD, "none"),
        RW("fcff_pv", "FCFF descontado", "USD", lambda m: f"={a('fcff', m)}*{a('df_p', m)}"),
        RW("cum_f", "FCFF acumulado", "USD", lambda m: f"={prv('RC.cum_f', m)}+{a('fcff', m)}", NF_USD, "last"),
        RW("cum_fd", "FCFF descontado acumulado", "USD", lambda m: f"={prv('RC.cum_fd', m)}+{a('fcff_pv', m)}", NF_USD, "last"),
        RW("eq_tv", "Valor final atribuible a socios (caja + activos netos - deuda)", "USD",
           lambda m: f"=IF({cal('f_fin', m)}=1,{a('caja', m)}+{a('tv', m)}-{a('deuda', m)},0)"),
        RW("eq", "FLUJO DE LOS SOCIOS (-aportes en efectivo, hipotéticos y en especie + dividendos pagados + valor final)", "USD",
           lambda m: f"=-{a('aportes', m)}-{a('ap_auto', m)}-{a('alq_esp', m)}+{a('divpag', m)}+{a('eq_tv', m)}", NF_USD, bold=True),
        RW("eq_pv", "Flujo de socios descontado a Ke", "USD", lambda m: f"={a('eq', m)}*{a('df_e', m)}"),
        RW("cum_e", "Flujo de socios acumulado", "USD", lambda m: f"={prv('RC.cum_e', m)}+{a('eq', m)}", NF_USD, "last"),
        RW("fcfe", "Memo: FCFE potencial (no se suma a dividendos)", "USD",
           lambda m: (f"={a('fcff_op', m)}+{a('alq_esp', m)}-{a('int_pag', m)}-{a('lin_int', m)}-{a('lin_iva', m)}+{a('desemb', m)}-{a('amort', m)}"
                      f"+{a('lin_uso', m)}-{a('lin_rep', m)}")),
        RW("inc", "Flujo incremental de continuar desde el mes de revisión (sin costos hundidos)", "USD",
           lambda m: f"=IF({m}>p_MesRev,{a('fcff', m)},0)-IF({m}=MAX(0,p_MesRev),N(p_ValorSalida),0)"),
        RW("sc_f", "Cambio de signo del FCFF acumulado (1/0)", "1/0",
           lambda m: "=0" if m == 0 else f"=IF(({prv('RC.cum_f', m)}<0)<>({a('cum_f', m)}<0),1,0)", NF_INT0),
        RW("sc_e", "Cambio de signo del flujo de socios acumulado (1/0)", "1/0",
           lambda m: "=0" if m == 0 else f"=IF(({prv('RC.cum_e', m)}<0)<>({a('cum_e', m)}<0),1,0)", NF_INT0),
        RW("pb_f", "Payback simple FCFF (mes con interpolación, último cruce)", "mes",
           lambda m: "=0" if m == 0 else f"=IF(AND({prv('RC.cum_f', m)}<0,{a('cum_f', m)}>=0),{m - 1}+(-{prv('RC.cum_f', m)})/{a('fcff', m)},0)", NF_USD2, "max"),
        RW("pb_fd", "Payback descontado FCFF (mes)", "mes",
           lambda m: "=0" if m == 0 else f"=IF(AND({prv('RC.cum_fd', m)}<0,{a('cum_fd', m)}>=0),{m - 1}+(-{prv('RC.cum_fd', m)})/{a('fcff_pv', m)},0)", NF_USD2, "max"),
        RW("pb_e", "Payback simple socios (mes)", "mes",
           lambda m: "=0" if m == 0 else f"=IF(AND({prv('RC.cum_e', m)}<0,{a('cum_e', m)}>=0),{m - 1}+(-{prv('RC.cum_e', m)})/{a('eq', m)},0)", NF_USD2, "max"),
    ]


# ------------------------------------------------------------------ escritura del motor
import re
LINK_RE = re.compile(r"^=('?[A-Za-z_]+'?!)\$?[A-Z]{1,3}\$?\d+$")

ENGINE = [
    ("CAL", cal_specs, 4, "Calendario — meses del modelo, años, banderas de operación/horizonte/reales, TC y factores de descuento"),
    ("PV", pv_specs, 4, "Produccion_Ventas — capacidad, demanda, rampa, recuperación, producción buena, inventario PT, precios y ventas (mensual)"),
    ("CP", cp_specs, None, None),
    ("IA", ia_specs, None, None),
    ("CT", ct_specs, 4, "Capital_Trabajo — clientes, inventarios, proveedores y anticipos (roll-forward mensual, USD)"),
    ("DT", dt_specs, 4, "Deuda_Tributos — IVA recuperable, préstamo original, tramo 2 y aportes (mensual, USD)"),
    ("RC", rc_specs, 4, "Resultados_Caja — estado de resultados, caja directa, tesorería, balance de control y flujos de valuación (mensual, USD)"),
]
SPECS = {}


def assign_rows():
    for S_, fn, start, _ in ENGINE:
        if S_ == "CP":
            start = CPT["engine_start"]
        if S_ == "IA":
            start = IAT["engine_start"]
        sp = fn()
        SPECS[S_] = (sp, start)
        r = start
        for it in sp:
            if it[0] == "sec":
                r += 1
            elif it[0] == "hdr":
                for k, *_ in hdr_rows(S_):
                    ROWS[f"{S_}.{k}"] = r
                    r += 1
            else:
                key = f"{S_}.{it[1]}"
                assert key not in ROWS, key
                ROWS[key] = r
                r += 1


def write_engine():
    for S_, fn, _, ttl in ENGINE:
        ws = WS[S_]
        CUR[0] = S_
        sp, start = SPECS[S_]
        if ttl:
            title(ws, ttl, "Motor único: fórmulas homogéneas por mes. Columna E = total del horizonte o saldo al cierre del horizonte. "
                           "No editar fórmulas; los supuestos se cambian en Supuestos / Datos_Reales.")
        ws.column_dimensions["A"].width = 12
        ws.column_dimensions["B"].width = 62
        ws.column_dimensions["C"].width = 8
        ws.column_dimensions["D"].width = 6
        ws.column_dimensions["E"].width = 15
        for m in range(NC):
            ws.column_dimensions[mc(m)].width = 11.5
        r = start
        first_data = None
        for it in sp:
            if it[0] == "sec":
                for c in range(1, 6):
                    ws.cell(row=r, column=c).fill = FILL_SEC
                setc(ws, r, 2, it[1], font=F_SEC, fill=FILL_SEC)
                r += 1
            elif it[0] == "hdr":
                for k, lab, ck, fmt in hdr_rows(S_):
                    setc(ws, r, 2, lab, font=F_BOLD)
                    setc(ws, r, 1, k)
                    for m in range(NC):
                        c = setc(ws, r, FC + m, f"={A('CAL.' + ck, m)}", fmt=fmt,
                                 font=F_HDR if k == "h_m" else F_LINK, fill=FILL_HDR if k == "h_m" else None)
                    if k == "h_m":
                        for c in range(1, 6):
                            ws.cell(row=r, column=c).fill = FILL_HDR
                        setc(ws, r, 2, "Mes del modelo", font=F_HDR, fill=FILL_HDR)
                        setc(ws, r, 5, "Total / cierre", font=F_HDR, fill=FILL_HDR)
                    r += 1
                first_data = r
            else:
                _, key, label, unit, fnc, fmt, tot, note, bold = it
                setc(ws, r, 1, key, font=F_SUB)
                setc(ws, r, 2, label, font=F_BOLD if bold else F_BASE, fill=FILL_TOT if bold else None)
                setc(ws, r, 3, unit)
                if note:
                    setc(ws, r, 4, note)
                rng = f"$F{r}:${MCN}{r}"
                totf = {"sum": f"=SUM({rng})", "last": f"=INDEX({rng},1,p_MesFin+1)", "max": f"=MAX({rng})", "min": f"=MIN({rng})"}.get(tot)
                if totf:
                    setc(ws, r, 5, totf, fmt=fmt, font=F_BOLD if bold else F_BASE, fill=FILL_TOT)
                for m in range(NC):
                    v = fnc(m)
                    font = F_BASE
                    if isinstance(v, str) and LINK_RE.match(v):
                        font = F_LINK
                    if bold:
                        font = Font(name=FONT, size=9, bold=True, color=font.color)
                    setc(ws, r, FC + m, v, fmt=fmt, font=font)
                r += 1
        if S_ in ("CP", "IA"):
            ws.freeze_panes = f"F{start}"
        elif S_ == "CAL":
            ws.freeze_panes = f"F{start + 2}"
            for c in range(1, FC + NC):
                ws.cell(row=ROWS['CAL.m'], column=c).fill = FILL_HDR
                ws.cell(row=ROWS['CAL.m'], column=c).font = F_HDR
        else:
            ws.freeze_panes = f"F{first_data}"


# ===========================================================================
# 5) RESULTADOS_ANUALES: agregación por año de proyecto + indicadores
# ===========================================================================
RA = {}
KPI = {}
RAY0 = 5   # columna E = año 0 ; F..R = años 1..13


def ycol(y):
    return CL(RAY0 + y)


def build_ra():
    ws = WS["RA"]
    CUR[0] = "RA"
    title(ws, "Resultados_Anuales — agregación por año de proyecto, comparación con el PDF e indicadores (USD)",
          "Flujos = suma de meses del año; saldos = cierre del año (no suma de saldos); ratios = cálculo sobre componentes. "
          "Año de proyecto 1 = primeros 12 meses después de la implantación prevista.")
    for k, v in {1: 10, 2: 60, 3: 11, 4: 16}.items():
        ws.column_dimensions[CL(k)].width = v
    for y in range(0, NYEARS + 1):
        ws.column_dimensions[ycol(y)].width = 13
    YR = RG("CAL.anio")

    # -------------------------------------------------- tabla anual (se escribe primero para conocer filas)
    r0 = 70
    r = r0
    section(ws, r - 2, "B. Resumen anual del escenario activo", 18)
    header(ws, r - 1, ["Cód.", "Concepto", "Unidad", "Total / cierre"] + [f"Año {y}" for y in range(0, NYEARS + 1)])

    def row(key, label, unit, kind, src=None, fmt=NF_USD, bold=False, note=None, pdf=False):
        nonlocal r
        assert key not in RA, key
        RA[key] = r
        setc(ws, r, 1, key, font=F_SUB)
        setc(ws, r, 2, label, font=F_BOLD if bold else F_BASE, fill=FILL_HIST if pdf else (FILL_TOT if bold else None))
        setc(ws, r, 3, unit)
        for y in range(0, NYEARS + 1):
            c = ycol(y)
            if kind == "sum":
                f = f"=SUMIFS({RG(src)},{YR},{y})"
            elif kind == "end":
                f = f"=INDEX({RG(src)},1,MIN({NC},p_ImplMeses+12*{y}+1))"
            elif kind == "max":
                f = f"=_xlfn.MAXIFS({RG(src)},{YR},{y})"
            elif kind == "pdf":
                f = f"=Base_PDF!{CL(BPY0 + y)}${BP[src]}" if 1 <= y <= 10 else (f"=Base_PDF!$E${BP[src]}" if (y == 0 and src in ("fl19_pub", "fcff_pdf", "eq_pdf")) else None)
            else:
                f = src(y, c)
            if f is not None:
                setc(ws, r, RAY0 + y, f, fmt=fmt, font=F_LINK if kind == "pdf" else (F_BOLD if bold else F_BASE))
        if kind == "sum" or (kind == "pdf" and fmt == NF_USD):
            setc(ws, r, 4, f"=SUM(E{r}:{ycol(NYEARS)}{r})", fmt=fmt, font=F_BOLD, fill=FILL_TOT)
        elif kind == "end":
            setc(ws, r, 4, f"=INDEX({RG(src)},1,p_MesFin+1)", fmt=fmt, font=F_BOLD, fill=FILL_TOT)
        elif kind == "max":
            setc(ws, r, 4, f"=MAX(E{r}:{ycol(NYEARS)}{r})", fmt=fmt, font=F_BOLD, fill=FILL_TOT)
        if note:
            setc(ws, r, RAY0 + NYEARS + 1, note, font=F_SUB)
        r += 1

    def sec(t):
        nonlocal r
        section(ws, r, t, 18)
        r += 1

    def rr(k, c):
        return f"{c}{RA[k]}"

    sec("Operación (unidades)")
    row("u_vg", "Gabinetes vendidos", "u", "sum", "PV.vtas_u_g", NF_INT)
    row("u_vg_pdf", "Gabinetes vendidos — PDF", "u", "pdf", "vta_g", NF_INT, pdf=True)
    row("u_vc", "Cajas vendidas", "u", "sum", "PV.vtas_u_c", NF_INT)
    row("u_vc_pdf", "Cajas vendidas — PDF", "u", "pdf", "vta_c", NF_INT, pdf=True)
    row("u_pg", "Gabinetes producidos (buenos)", "u", "sum", "PV.prod_g", NF_INT)
    row("u_pc", "Cajas producidas (buenas)", "u", "sum", "PV.prod_c", NF_INT)
    row("u_ig", "Gabinetes iniciados", "u", "sum", "PV.ini_g", NF_INT)
    row("u_ic", "Cajas iniciadas", "u", "sum", "PV.ini_c", NF_INT)
    row("cap_g", "Capacidad gabinetes", "u", "sum", "PV.cap_g", NF_INT)
    row("cap_c", "Capacidad cajas", "u", "sum", "PV.cap_c", NF_INT)
    row("ut_g", "Utilización gabinetes (iniciadas / capacidad)", "%", "f", lambda y, c: f'=IF({rr("cap_g", c)}>0,{rr("u_ig", c)}/{rr("cap_g", c)},"")', NF_PCT)
    row("ut_c", "Utilización cajas", "%", "f", lambda y, c: f'=IF({rr("cap_c", c)}>0,{rr("u_ic", c)}/{rr("cap_c", c)},"")', NF_PCT)
    row("pdef", "Ventas perdidas definitivamente (u, ambos productos)", "u", "f",
        lambda y, c: f"=SUMIFS({RG('PV.pdef_g')},{YR},{y})+SUMIFS({RG('PV.pdef_c')},{YR},{y})", NF_INT)
    row("exc_m", "Meses con producción sobre capacidad (alerta)", "meses", "sum", "PV.exc_flag", NF_INT0)
    sec("Estado de resultados (USD)")
    row("v", "Ventas netas", "USD", "sum", "RC.v", bold=True)
    row("v_g", "   Gabinetes", "USD", "sum", "RC.v_g")
    row("v_c", "   Cajas", "USD", "sum", "RC.v_c")
    row("cv", "Costo de ventas", "USD", "sum", "RC.cv")
    row("mb", "Margen bruto", "USD", "sum", "RC.mb", bold=True)
    row("mb_p", "Margen bruto / ventas", "%", "f", lambda y, c: f'=IF({rr("v", c)}<>0,{rr("mb", c)}/{rr("v", c)},"")', NF_PCT)
    row("g_var", "Gastos variables de venta y tributo de maquila", "USD", "sum", "RC.g_var")
    row("g_adm", "Gastos de administración", "USD", "sum", "RC.g_adm")
    row("g_ext", "Costos preoperativos / extraordinarios reales", "USD", "sum", "RC.g_extra")
    row("iva_r", "IVA recuperable reclasificado", "USD", "sum", "RC.iva_rec")
    row("ebitda", "EBITDA", "USD", "sum", "RC.ebitda", bold=True)
    row("ebitda_p", "Margen EBITDA", "%", "f", lambda y, c: f'=IF({rr("v", c)}<>0,{rr("ebitda", c)}/{rr("v", c)},"")', NF_PCT)
    row("dep", "Depreciación y amortización", "USD", "sum", "RC.dep")
    row("ebit", "EBIT", "USD", "sum", "RC.ebit")
    row("gfin", "Gastos financieros (intereses + IVA + comisiones)", "USD", "sum", "RC.gfin")
    row("ire", "IRE", "USD", "sum", "RC.ire")
    row("ni", "Resultado neto", "USD", "sum", "RC.ni", bold=True)
    row("ni_p", "Resultado neto / ventas", "%", "f", lambda y, c: f'=IF({rr("v", c)}<>0,{rr("ni", c)}/{rr("v", c)},"")', NF_PCT)
    row("preop", "Memo: costos de espera / implantación", "USD", "sum", "RC.m_preop")
    sec("Comparación con el PDF publicado (periodos equivalentes solo si implantación prevista = 0)")
    row("v_pdf", "Ventas — PDF", "USD", "pdf", "ing_pub", pdf=True)
    row("ebitda_pdf", "EBITDA — PDF", "USD", "pdf", "ebitda_pub", pdf=True)
    row("ni_pdf", "Utilidad neta — PDF", "USD", "pdf", "ut_pub", pdf=True)
    row("dv", "Diferencia ventas (activo - PDF)", "USD", "f", lambda y, c: f"={rr('v', c)}-{rr('v_pdf', c)}" if 1 <= y <= 10 else None)
    row("debitda", "Diferencia EBITDA (activo - PDF)", "USD", "f", lambda y, c: f"={rr('ebitda', c)}-{rr('ebitda_pdf', c)}" if 1 <= y <= 10 else None)
    row("dni", "Diferencia resultado neto (activo - PDF)", "USD", "f", lambda y, c: f"={rr('ni', c)}-{rr('ni_pdf', c)}" if 1 <= y <= 10 else None)
    for k in ("dv", "debitda", "dni"):
        ws.cell(row=RA[k], column=4).value = f"=SUM(F{RA[k]}:O{RA[k]})"
        ws.cell(row=RA[k], column=4).number_format = NF_USD
    sec("Detalle de costos (USD)")
    for k, lab, src in [("c_mat", "Materia prima, insumos y consumibles", "RC.c_mat"), ("c_fle", "Flete de MP", "RC.c_fle"),
                        ("c_des", "Despachos de importación", "RC.c_des"), ("c_pkg", "Packaging", "RC.c_pkg"),
                        ("c_mod", "Personal directo", "RC.c_mod"), ("c_moi", "Personal indirecto de fábrica", "RC.c_moi"),
                        ("c_ene", "Energía de planta", "RC.c_ene"), ("c_man", "Mantenimiento de planta", "RC.c_man"),
                        ("c_imp", "Imprevistos de producción", "RC.c_imp"), ("g_trib", "Tributo único de maquila", "RC.g_trib"),
                        ("g_segc", "Seguro de caución", "RC.g_segc"), ("g_gvta", "Gastos de venta y logística", "RC.g_gvta"),
                        ("g_inc", "Incobrables", "RC.g_incob"), ("g_alq", "Alquiler de nave (incl. IVA)", "RC.g_alq"),
                        ("per", "Costo total de personal", "CP.per_tot")]:
        row(k, lab, "USD", "sum", src)
    row("dot", "Dotación máxima del año", "pers.", "max", "CP.dot_tot", NF_USD2)
    sec("Flujo de caja y financiamiento (USD)")
    row("cfo", "Flujo operativo", "USD", "sum", "RC.cfo", bold=True)
    row("capex", "CAPEX pagado", "USD", "sum", "RC.capex")
    row("capex_m", "CAPEX de mantenimiento / reposición", "USD", "sum", "IA.pag_mant")
    row("fcff_op", "FCFF antes de valor terminal", "USD", "sum", "RC.fcff_op")
    row("vt", "Valor terminal", "USD", "sum", "RC.vt")
    row("fcff", "FCFF (proyecto)", "USD", "sum", "RC.fcff", bold=True)
    row("fl19_pdf", "Flujo cuadro 19 PDF (mezcla perspectivas — solo referencia)", "USD", "pdf", "fl19_pub", pdf=True)
    row("fcff_pdf", "FCFF reexpresado con datos PDF (Base_PDF X01)", "USD", "pdf", "fcff_pdf", pdf=True)
    row("desemb", "Desembolsos de préstamos", "USD", "sum", "RC.desemb")
    row("amort", "Amortizaciones", "USD", "sum", "RC.amort")
    row("int_pag", "Intereses, IVA y comisiones", "USD", "sum", "RC.int_pag")
    row("lin_int", "Intereses + IVA línea hipotética", "USD", "f",
        lambda y, c: f"=SUMIFS({RG('RC.lin_int')},{YR},{y})+SUMIFS({RG('RC.lin_iva')},{YR},{y})")
    row("aportes", "Aportes de socios en efectivo", "USD", "sum", "RC.aportes")
    row("ap_auto", "Aportes hipotéticos de cobertura de brecha (no aprobados)", "USD", "sum", "RC.ap_auto")
    row("ap_esp", "Aportes en especie (alquiler sin desembolso)", "USD", "sum", "RC.alq_esp")
    row("divpag", "Dividendos pagados", "USD", "sum", "RC.divpag")
    row("decl", "Dividendos declarados", "USD", "sum", "RC.decl")
    row("eq", "Flujo de los socios", "USD", "sum", "RC.eq", bold=True)
    row("eq_pdf", "Flujo de socios reexpresado con datos PDF (Base_PDF X04)", "USD", "pdf", "eq_pdf", pdf=True)
    row("caja", "Caja al cierre", "USD", "end", "RC.caja", bold=True)
    row("brecha", "Necesidad de financiamiento máxima del año (mensual)", "USD", "max", "RC.brecha")
    row("brecha_nc", "Brecha no cubierta máxima del año", "USD", "max", "RC.brecha_nc")
    row("nec_pre", "Necesidad acumulada antes de financiamiento adicional (máx.)", "USD", "max", "RC.nec_pre")
    row("deuda", "Deuda financiera al cierre", "USD", "end", "RC.deuda")
    sec("Capital de trabajo al cierre (USD)")
    for k, lab, src in [("cxc", "Cuentas por cobrar", "CT.cxc"), ("stk", "Inventario de materiales", "CT.stk"), ("pt", "Inventario de PT", "CT.pt"),
                        ("iva", "Crédito fiscal IVA", "DT.iva"), ("antp", "Anticipos al personal y proveedores", None),
                        ("prov", "Proveedores de materiales", "CT.prov"), ("pmen", "Proveedores menores", "CT.pmen"),
                        ("ctn", "Capital de trabajo operativo neto", "CT.ctn")]:
        if src:
            row(k, lab, "USD", "end", src, bold=(k == "ctn"))
        else:
            row(k, lab, "USD", "f", lambda y, c: (f"=INDEX({RG('CT.antpers')},1,MIN({NC},p_ImplMeses+12*{y}+1))"
                                                  f"+INDEX({RG('CT.antprov')},1,MIN({NC},p_ImplMeses+12*{y}+1))"))
    row("ctn_pdf", "CT neto — PDF (cuadro 16; incluye caja 3 días)", "USD", "pdf", "ctn_pub", pdf=True)
    row("dso", "Días de cuentas por cobrar al cierre (CxC / ventas x 365; aproximación)", "días", "f",
        lambda y, c: f'=IF(AND({y}>0,{rr("v", c)}>0),{rr("cxc", c)}/{rr("v", c)}*p_DiasAnio,"")', NF_INT)
    row("dio", "Días de inventario de materiales (stock / consumo x 365)", "días", "f",
        lambda y, c: f'=IF({rr("c_mat", c)}>0,{rr("stk", c)}/({rr("c_mat", c)}+IF(p_C5=1,{rr("c_fle", c)}+{rr("c_des", c)},0))*p_DiasAnio,"")', NF_INT)
    sec("Servicio de la deuda y cobertura")
    row("ds", "Servicio de deuda exigible (capital + intereses + IVA + comisiones + línea)", "USD", "f",
        lambda y, c: f"={rr('amort', c)}+{rr('int_pag', c)}+{rr('lin_int', c)}")
    row("cfads", "CFADS = flujo operativo - CAPEX de mantenimiento", "USD", "f", lambda y, c: f"={rr('cfo', c)}-{rr('capex_m', c)}")
    row("dscr", "DSCR = CFADS / servicio de deuda", "x", "f", lambda y, c: f'=IF({rr("ds", c)}>0,{rr("cfads", c)}/{rr("ds", c)},"")', NF_X)
    row("ebitda_ds", "EBITDA / servicio de deuda (aproximación, no DSCR completo)", "x", "f",
        lambda y, c: f'=IF({rr("ds", c)}>0,{rr("ebitda", c)}/{rr("ds", c)},"")', NF_X)
    sec("Punto de equilibrio con mix del año (ventas en USD; unidades por producto)")
    row("cvar", "Costos variables", "USD", "f", lambda y, c: f"=SUMIFS({RG('CP.c_var')},{YR},{y})+SUMIFS({RG('CT.incob')},{YR},{y})")
    row("cfij", "Costos fijos operativos (neto de IVA recuperable)", "USD", "f",
        lambda y, c: f"=SUMIFS({RG('CP.c_fij')},{YR},{y})-{rr('iva_r', c)}")
    row("mc_p", "Margen de contribución / ventas", "%", "f", lambda y, c: f'=IF({rr("v", c)}>0,1-{rr("cvar", c)}/{rr("v", c)},"")', NF_PCT)
    row("pe_ebitda", "Equilibrio EBITDA = 0 (ventas)", "USD", "f", lambda y, c: f'=IF(N({rr("mc_p", c)})>0,{rr("cfij", c)}/{rr("mc_p", c)},"")')
    row("pe_ebit", "Equilibrio contable EBIT = 0 (incluye D&A)", "USD", "f",
        lambda y, c: f'=IF(N({rr("mc_p", c)})>0,({rr("cfij", c)}+{rr("dep", c)})/{rr("mc_p", c)},"")')
    row("pe_caja", "Equilibrio de cobertura de deuda (fijos + servicio de deuda)", "USD", "f",
        lambda y, c: f'=IF(N({rr("mc_p", c)})>0,({rr("cfij", c)}+{rr("ds", c)})/{rr("mc_p", c)},"")')
    row("pe_ug", "Gabinetes en equilibrio EBITDA (mix del año)", "u", "f",
        lambda y, c: f'=IF(AND(N({rr("pe_ebitda", c)})>0,{rr("v", c)}>0),{rr("pe_ebitda", c)}/{rr("v", c)}*{rr("u_vg", c)},"")', NF_INT)
    row("pe_uc", "Cajas en equilibrio EBITDA (mix del año)", "u", "f",
        lambda y, c: f'=IF(AND(N({rr("pe_ebitda", c)})>0,{rr("v", c)}>0),{rr("pe_ebitda", c)}/{rr("v", c)}*{rr("u_vc", c)},"")', NF_INT)
    row("ms", "Margen de seguridad sobre equilibrio EBITDA", "%", "f",
        lambda y, c: f'=IF(AND(N({rr("pe_ebitda", c)})>0,{rr("v", c)}>0),1-{rr("pe_ebitda", c)}/{rr("v", c)},"")', NF_PCT)
    sec("Rentabilidad por producto (asignación provisional por driver; rótulos: margen = (venta - costo)/venta; markup = (venta - costo)/costo)")
    row("stdg", "Costo estándar de materiales gabinetes (clave de asignación)", "USD", "sum", "CP.std_g")
    row("stdc", "Costo estándar de materiales cajas (clave de asignación)", "USD", "sum", "CP.std_c")
    row("wg", "Participación de gabinetes en el driver de conversión", "%", "f",
        lambda y, c: f'=IF(({rr("u_pg", c)}*p_WG+{rr("u_pc", c)}*p_WC)>0,{rr("u_pg", c)}*p_WG/({rr("u_pg", c)}*p_WG+{rr("u_pc", c)}*p_WC),"")', NF_PCT)
    for p, P, nm in [("g", "G", "GABINETES"), ("c", "C", "CAJAS")]:
        std, ostd = ("stdg", "stdc")
        mine = "stdg" if p == "g" else "stdc"
        share_w = (lambda c: f"N({rr('wg', c)})") if p == "g" else (lambda c: f"(1-N({rr('wg', c)}))")
        vk = "v_g" if p == "g" else "v_c"
        uk = "u_vg" if p == "g" else "u_vc"
        row(f"p{p}_v", f"{nm}: ventas", "USD", "f", lambda y, c, vk=vk: f"={rr(vk, c)}", bold=True)
        row(f"p{p}_mat", f"{nm}: materiales, flete y despacho (asignado por costo estándar)", "USD", "f",
            lambda y, c, mine=mine: f'=IF(({rr("stdg", c)}+{rr("stdc", c)})>0,({rr("c_mat", c)}+{rr("c_fle", c)}+{rr("c_des", c)})*{rr(mine, c)}/({rr("stdg", c)}+{rr("stdc", c)}),0)')
        row(f"p{p}_pkg", f"{nm}: packaging", "USD", "f",
            lambda y, c, mine=mine: f'=IF(({rr("stdg", c)}+{rr("stdc", c)})>0,{rr("c_pkg", c)}*{rr(mine, c)}/({rr("stdg", c)}+{rr("stdc", c)}),0)')
        row(f"p{p}_vv", f"{nm}: tributo, seguro, logística e incobrables (por ventas)", "USD", "f",
            lambda y, c, vk=vk: f'=IF(({rr("v_g", c)}+{rr("v_c", c)})>0,({rr("g_trib", c)}+{rr("g_segc", c)}+{rr("g_gvta", c)}+{rr("g_inc", c)})*{rr(vk, c)}/({rr("v_g", c)}+{rr("v_c", c)}),0)')
        row(f"p{p}_cvar", f"{nm}: costo variable", "USD", "f", lambda y, c, p=p: f"={rr('p' + p + '_mat', c)}+{rr('p' + p + '_pkg', c)}+{rr('p' + p + '_vv', c)}")
        row(f"p{p}_contr", f"{nm}: contribución marginal", "USD", "f", lambda y, c, p=p: f"={rr('p' + p + '_v', c)}-{rr('p' + p + '_cvar', c)}", bold=True)
        row(f"p{p}_contr_p", f"{nm}: margen de contribución sobre ventas", "%", "f",
            lambda y, c, p=p: f'=IF({rr("p" + p + "_v", c)}>0,{rr("p" + p + "_contr", c)}/{rr("p" + p + "_v", c)},"")', NF_PCT)
        row(f"p{p}_conv", f"{nm}: conversión fabril asignada (MO, indirectos, energía, mant., imprevistos)", "USD", "f",
            lambda y, c, sw=share_w: f"=({rr('c_mod', c)}+{rr('c_moi', c)}+{rr('c_ene', c)}+{rr('c_man', c)}+{rr('c_imp', c)})*{sw(c)}")
        row(f"p{p}_mfab", f"{nm}: margen después del costo de fabricación absorbido", "USD", "f",
            lambda y, c, p=p: f"={rr('p' + p + '_contr', c)}-{rr('p' + p + '_conv', c)}")
        row(f"p{p}_adm", f"{nm}: gastos comunes de administración asignados", "USD", "f",
            lambda y, c, sw=share_w: f"=({rr('g_adm', c)}+{rr('g_ext', c)}-{rr('iva_r', c)})*{sw(c)}")
        row(f"p{p}_res", f"{nm}: resultado asignado (antes de D&A, financieros y variación de PT)", "USD", "f",
            lambda y, c, p=p: f"={rr('p' + p + '_mfab', c)}-{rr('p' + p + '_adm', c)}", bold=True)
        row(f"p{p}_pu", f"{nm}: precio medio", "USD/u", "f", lambda y, c, p=p, uk=uk: f'=IF({rr(uk, c)}>0,{rr("p" + p + "_v", c)}/{rr(uk, c)},"")', NF_USD2)
        row(f"p{p}_cvu", f"{nm}: costo variable unitario", "USD/u", "f", lambda y, c, p=p, uk=uk: f'=IF({rr(uk, c)}>0,{rr("p" + p + "_cvar", c)}/{rr(uk, c)},"")', NF_USD2)
        row(f"p{p}_ctu", f"{nm}: costo total asignado unitario", "USD/u", "f",
            lambda y, c, p=p, uk=uk: f'=IF({rr(uk, c)}>0,({rr("p" + p + "_cvar", c)}+{rr("p" + p + "_conv", c)}+{rr("p" + p + "_adm", c)})/{rr(uk, c)},"")', NF_USD2)
        row(f"p{p}_mk", f"{nm}: markup sobre costo total asignado", "%", "f",
            lambda y, c, p=p: f'=IF(N({rr("p" + p + "_ctu", c)})>0,({rr("p" + p + "_pu", c)}-{rr("p" + p + "_ctu", c)})/{rr("p" + p + "_ctu", c)},"")', NF_PCT)
    sec("Valuación con convención anual (fin de año) — comparación con el método del PDF")
    row("df_a", "Factor de descuento anual del proyecto", "factor", "f", lambda y, c: f"=1/(1+p_TasaProy)^{y}", NF_USD4)
    RA["_end"] = r

    # -------------------------------------------------- bloque de indicadores (arriba)
    r = 4
    section(ws, r, "A. Indicadores clave del escenario activo (se recalculan con el selector)", 18); r += 1
    header(ws, r, ["Cód.", "Indicador", "Unidad", "Valor", "Nota"]); r += 1
    fcff, eq, dfp, dfe = RG("RC.fcff"), RG("RC.eq"), RG("RC.df_p"), RG("RC.df_e")

    def kpi(key, label, unit, f, fmt=NF_USD, note=None, bold=False):
        nonlocal r
        KPI[key] = r
        setc(ws, r, 1, key, font=F_SUB)
        setc(ws, r, 2, label, font=F_BOLD if bold else F_BASE)
        setc(ws, r, 3, unit)
        setc(ws, r, 4, f, fmt=fmt, font=F_BOLD if bold else F_BASE, fill=FILL_TOT)
        if note:
            setc(ws, r, 5, note, font=F_SUB)
        add_name("k_" + key, "RA", f"D{r}")
        r += 1

    def irr_txt(rng):
        return (f'IF(COUNTIF({rng},"<0")=0,"No existe: sin flujos negativos",IF(COUNTIF({rng},">0")=0,"No existe: sin flujos positivos",'
                f'IFERROR((1+IRR({rng},0.01))^12-1,"No converge")))')

    kpi("van_p", "VAN del proyecto (FCFF, descuento mensual a la tasa del proyecto)", "USD", f"=SUMPRODUCT({fcff},{dfp})", bold=True,
        note="FCFF: excluye préstamos, intereses, aportes y dividendos.")
    kpi("tir_p", "TIR del proyecto (anual equivalente de la TIR mensual)", "%", "=" + irr_txt(fcff), NF_PCT2, bold=True)
    kpi("tir_p_u", "Unicidad de la TIR del proyecto", "texto",
        f'=IF(SUM({RG("RC.sc_f")})=1,"Única: el flujo acumulado cambia de signo una vez","Revisar: el acumulado cambia de signo "&SUM({RG("RC.sc_f")})&" veces; priorizar VAN")', None)
    kpi("van_a", "VAN del proyecto con convención anual (flujos anuales a fin de año)", "USD",
        f"={ycol(0)}{RA['fcff']}+NPV(p_TasaProy,{ycol(1)}{RA['fcff']}:{ycol(NYEARS)}{RA['fcff']})",
        note="Comparable con el método del PDF; difiere del mensual por la oportunidad de los flujos.")
    kpi("tir_a", "TIR del proyecto con convención anual", "%", f'=IFERROR(IRR({ycol(0)}{RA["fcff"]}:{ycol(NYEARS)}{RA["fcff"]},0.1),"No converge")', NF_PCT2)
    kpi("van_e", "VAN de los socios (aportes, dividendos pagados y valor final; descuento a Ke)", "USD", f"=SUMPRODUCT({eq},{dfe})", bold=True,
        note="No suma el FCFE no distribuido: queda dentro del valor final.")
    kpi("tir_e", "TIR de los socios (anual equivalente)", "%", "=" + irr_txt(eq), NF_PCT2, bold=True)
    kpi("tir_e_u", "Unicidad de la TIR de los socios", "texto",
        f'=IF(SUM({RG("RC.sc_e")})=1,"Única","Revisar: el acumulado cambia de signo "&SUM({RG("RC.sc_e")})&" veces; priorizar VAN")', None)
    kpi("pb_s", "Payback simple del proyecto (FCFF)", "años", f'=IF(MAX({RG("RC.pb_f")})=0,"No recupera en el horizonte",MAX({RG("RC.pb_f")})/12)', NF_USD2)
    kpi("pb_d", "Payback descontado del proyecto (misma serie FCFF)", "años", f'=IF(MAX({RG("RC.pb_fd")})=0,"No recupera en el horizonte",MAX({RG("RC.pb_fd")})/12)', NF_USD2)
    kpi("pb_e", "Payback simple de los socios", "años", f'=IF(MAX({RG("RC.pb_e")})=0,"No recupera en el horizonte",MAX({RG("RC.pb_e")})/12)', NF_USD2)
    kpi("tasa", "Tasa del proyecto aplicada / estado", "%", "=p_TasaProy", NF_PCT4, note="=p_TasaEstado")
    ws.cell(row=KPI["tasa"], column=5).value = "=p_TasaEstado"
    kpi("van_inc", "VAN incremental de continuar vs. salir, desde el mes de revisión", "USD",
        f'=SUMPRODUCT({RG("RC.inc")},{dfp})/INDEX({dfp},1,MAX(1,p_MesRev+1))',
        note="Excluye costos hundidos. Si no se cargó el valor de salida (G62) equivale al VAN de continuar sin alternativa.")
    ws.cell(row=KPI["van_inc"], column=5).value = ('=IF(p_ValorSalida="","PENDIENTE: falta valor de la alternativa de salida (Supuestos G62); '
                                                   'el valor mostrado es el VAN de continuar","Continuar - salir, al mes de revisión "&p_MesRev)')
    r += 1
    kpi("capex_o", "CAPEX presupuesto original (todos los ítems)", "USD", f"=Inversion_Activos!$H${IAT['tot']}")
    kpi("capex_e", "CAPEX costo final estimado", "USD", f"=Inversion_Activos!$O${IAT['tot']}", bold=True)
    kpi("capex_d", "Desviación de CAPEX", "USD", f"=Inversion_Activos!$P${IAT['tot']}")
    kpi("capex_pag", "CAPEX pagado a la fecha (real)", "USD", f"=Inversion_Activos!$J${IAT['tot']}")
    kpi("capex_pend", "CAPEX pendiente de pago", "USD", f"=Inversion_Activos!$X${IAT['tot']}")
    kpi("v1", "Ventas del año de proyecto 1", "USD", f"={ycol(1)}{RA['v']}")
    kpi("e1", "EBITDA del año de proyecto 1", "USD", f"={ycol(1)}{RA['ebitda']}")
    kpi("e1p", "Margen EBITDA del año 1", "%", f"={ycol(1)}{RA['ebitda_p']}", NF_PCT)
    kpi("vt", "Ventas totales del horizonte", "USD", f"=D{RA['v']}")
    kpi("et", "EBITDA total del horizonte", "USD", f"=D{RA['ebitda']}")
    kpi("nit", "Resultado neto total del horizonte", "USD", f"=D{RA['ni']}")
    kpi("ap", "Aportes de socios en efectivo (total)", "USD", f"=D{RA['aportes']}", bold=True)
    kpi("ap_auto", "Aportes hipotéticos de cobertura de brecha (NO APROBADOS; solo si G63 = 1)", "USD", f"=D{RA['ap_auto']}")
    kpi("ap_esp", "Aportes en especie (alquiler sin desembolso)", "USD", f"=D{RA['ap_esp']}")
    kpi("div", "Dividendos pagados (total)", "USD", f"=D{RA['divpag']}")
    kpi("nec", "Pico de necesidad de fondos adicionales (antes de aportes adicionales, tramo 2 y línea)", "USD",
        f"=MAX({RG('RC.nec_pre')})", bold=True, note="Incluye caja mínima. Financiado en el plan con aportes adicionales.")
    kpi("nec_m", "Mes del pico de necesidad", "mes", f"=MATCH(MAX({RG('RC.nec_pre')}),{RG('RC.nec_pre')},0)-1", NF_INT0)
    kpi("brecha", "Necesidad mensual máxima con el financiamiento planificado", "USD", f"=MAX({RG('RC.brecha')})")
    kpi("brecha_m", "Mes de la necesidad máxima", "mes", f"=MATCH(MAX({RG('RC.brecha')}),{RG('RC.brecha')},0)-1", NF_INT0)
    kpi("brecha_nc", "Brecha no cubierta máxima (después de la línea hipotética)", "USD", f"=MAX({RG('RC.brecha_nc')})", bold=True)
    kpi("brecha_nc_n", "Meses con brecha no cubierta", "meses", f'=COUNTIF({RG("RC.brecha_nc")},">0.5")', NF_INT0)
    kpi("caja_min", "Caja final mínima del horizonte", "USD", f"=MIN({RG('RC.caja')})")
    kpi("deuda_max", "Deuda financiera máxima", "USD", f"=MAX({RG('RC.deuda')})")
    kpi("dscr_min", "DSCR mínimo (años con servicio de deuda)", "x",
        f'=IFERROR(_xlfn.MINIFS({ycol(1)}{RA["dscr"]}:{ycol(NYEARS)}{RA["dscr"]},{ycol(1)}{RA["ds"]}:{ycol(NYEARS)}{RA["ds"]},">0"),"Sin servicio de deuda")', NF_X)
    kpi("pe1", "Equilibrio EBITDA del año 1 (ventas)", "USD", f"={ycol(1)}{RA['pe_ebitda']}")
    kpi("pe1g", "   Gabinetes en equilibrio año 1 (mix del año)", "u", f"={ycol(1)}{RA['pe_ug']}", NF_INT)
    kpi("pe1c", "   Cajas en equilibrio año 1 (mix del año)", "u", f"={ycol(1)}{RA['pe_uc']}", NF_INT)
    kpi("ctn_max", "Capital de trabajo operativo neto máximo", "USD", f"=MAX({RG('CT.ctn')})")
    r += 1
    kpi("chk_bal", "Control balance: máxima diferencia absoluta (0 = OK)", "USD",
        f"=MAX(ABS(MAX({RG('RC.b_chk')})),ABS(MIN({RG('RC.b_chk')})))", NF_USD2)
    kpi("chk_ind", "Control directo vs indirecto: máxima diferencia absoluta (0 = OK)", "USD",
        f"=MAX(ABS(MAX({RG('RC.ind_chk')})),ABS(MIN({RG('RC.ind_chk')})))", NF_USD2)
    kpi("chk_cap", "Meses con producción sobre capacidad", "meses", f"=SUM({RG('PV.exc_flag')})", NF_INT0)
    kpi("chk_real", "Meses reales cerrados sin registros cargados", "meses",
        f'=IF(p_MesCorte<0,0,SUMPRODUCT(({RG("CAL.m")}<=p_MesCorte)*(COUNTIF(dr_Mes,{RG("CAL.m")})=0)))', NF_INT0)
    kpi("chk_grid", "Horizonte dentro de la rejilla (1 = OK)", "1/0", "=IF(p_MesFinTeor<=%d,1,0)" % N, NF_INT0)
    KPI["_end"] = r
    assert r < r0 - 3, (r, r0)

    # fila de VAN anual: columnas de valor presente
    ws.freeze_panes = "E6"


# ===========================================================================
# 6) VALIDACIONES
# ===========================================================================
VAL = {}


def bp(key, col="E"):
    return f"Base_PDF!${col}${BP[key]}"


def bpmaxabs(key):
    rng = f"Base_PDF!$F${BP[key]}:$O${BP[key]}"
    return f"MAX(ABS(MAX({rng})),ABS(MIN({rng})))"


def build_val():
    ws = WS["VAL"]
    CUR[0] = "VAL"
    title(ws, "Validaciones — controles de reproducción e integridad, registro de inconsistencias y pruebas",
          "Esta hoja revisa; no alimenta los cálculos del negocio. Estado «OK» cuando |diferencia| ≤ tolerancia.")
    for k, v in {1: 7, 2: 62, 3: 16, 4: 16, 5: 14, 6: 11, 7: 10, 8: 80}.items():
        ws.column_dimensions[CL(k)].width = v
    r = 4
    section(ws, r, "A. Controles con diferencia calculada y tolerancia", 8); r += 1
    header(ws, r, ["ID", "Control", "Valor calculado", "Referencia", "Diferencia", "Tolerancia", "Estado", "Explicación de la tolerancia / resultado"]); r += 1
    VAL["first"] = r
    ctrls = [
        ("Inversión inicial = USD 655.044 + USD 2.336.543", f"={bp('inv_tot_calc')}", f"={bp('inv_total')}", 0.5, "Exacto."),
        ("Fuentes iniciales = USD 200.000 + USD 2.791.587", f"={bp('fuentes_calc')}", f"={bp('inv_total')}", 0.5, "Exacto."),
        ("Inversión fija = infraestructura + CIF + despacho + software + mobiliario", f"={bp('inv_fija_calc')}", f"={bp('inv_fija')}", 1,
         "CIF publicado 492.394 redondea 492.394,10."),
        ("Ventas por producto suman ventas totales (PDF, máx. dif. anual)", f"={bpmaxabs('ing_sum_dif')}", 0, 1.01, "Redondeo de la publicación (±1)."),
        ("Ventas recalculadas (precio sin redondear) vs publicadas (máx. dif.)", f"={bpmaxabs('ing_dif')}", 0, 1, "Precisión interna del precio."),
        ("Reproducción cuadro 10a con bases identificadas (máx. dif. de todas las líneas)", f"=MAX(Base_PDF!$F${BP['d_absmax']}:$O${BP['d_absmax']})", 0, 10,
         "Hasta USD 9,3/año: coeficiente de prestaciones publicado 43,09% vs implícito ≈43,087%."),
        ("EBITDA recalculado vs publicado (máx. dif.)", f"={bpmaxabs('ebitda_dif')}", 0, 1.01, "Redondeo."),
        ("Utilidad neta recalculada vs publicada (máx. dif.)", f"={bpmaxabs('ut_dif')}", 0, 1.01, "Redondeo."),
        ("Préstamo: suma de amortizaciones del cronograma = monto", f"=Base_PDF!$P${BP['amort_calc']}", f"={bp('prestamo')}", 0.01, "Cancela el préstamo."),
        ("Préstamo: cuota francesa posterior (sin IVA)", f"={bp('cuota_calc')}", 63060.2066, 0.001, "Control independiente del prompt."),
        ("Préstamo: interés mensual en gracia", f"={bp('int_gracia')}", 20936.9025, 0.001, "Control independiente."),
        ("Préstamo: (amortización + intereses) anuales vs cuadro 13 (máx. dif.)", f"={bpmaxabs('deuda_dif')}", 0, 1, "Redondeo de la publicación."),
        ("VAN reproducido con tasa exacta vs publicado", f"={bp('van_calc')}", f"={bp('van_pub')}", 1, "Diferencia ≈ USD 0,60 por redondeo de los flujos."),
        ("TIR reproducida vs publicada", f"={bp('tir_calc')}", f"={bp('tir_pub')}", 0.00005, "22,6055% vs 22,61%."),
        ("CT neto reconstruido (cuadro 16) vs publicado (máx. dif.)", f"={bpmaxabs('ctn_dif')}", 0, 1, "Con proveedores menores publicados (base no identificada)."),
        ("Flujo cuadro 19 reproducido vs publicado (máx. dif.)", f"=MAX({bpmaxabs('fl19_dif')},ABS({bp('fl19_dif')}))", 0, 1.01, "Redondeo."),
        ("Saldo cuadro 18 reproducido vs publicado (máx. dif.)", f"={bpmaxabs('saldo18_dif')}", 0, 1.01, "Redondeo."),
        ("Motor: ventas anuales = suma de meses (total horizonte)", f"=Resultados_Anuales!$D${RA['v']}", f"=Resultados_Caja!$E${ROWS['RC.v']}", 0.01, "Año = suma de flujos mensuales."),
        ("Motor: balance de control cuadra todos los meses", "=k_chk_bal", 0, 0.01, "Activo - pasivo - patrimonio."),
        ("Motor: método directo = método indirecto todos los meses", "=k_chk_ind", 0, 0.01, "EBITDA - ΔCT operativo - IRE vs cobros - pagos."),
        ("Motor: préstamo 1 amortizado = desembolsado (si el horizonte cubre el plazo)",
         f"=Deuda_Tributos!$E${ROWS['DT.l1_am']}", f"=Deuda_Tributos!$E${ROWS['DT.l1_desemb']}-Deuda_Tributos!$E${ROWS['DT.l1_sal']}", 0.01,
         "Capital pagado + saldo final = desembolsos."),
        ("Motor: saldo de préstamos nunca negativo", f"=MIN(Deuda_Tributos!$E${ROWS['DT.l1_neg']},Deuda_Tributos!$E${ROWS['DT.l2_neg']})", 0, 0, "0 = OK."),
        ("Motor: depreciación acumulada ≤ base depreciable en servicio", f"=Inversion_Activos!$E${ROWS['IA.ctrl_dep']}", 0, 0.01, "0 = OK."),
        ("Motor: activo fijo neto final = devengado - depreciación acumulada",
         f"=Inversion_Activos!$E${ROWS['IA.af_neto']}", f"=Inversion_Activos!$E${ROWS['IA.cum_dev']}-Inversion_Activos!$E${ROWS['IA.dep_acum']}", 0.01, "Roll-forward."),
        ("Motor: inventario de PT en unidades nunca negativo", f"=MIN(MIN({RG('PV.ptu_g')}),MIN({RG('PV.ptu_c')}))", 0, 0.001, "Inicial + producción - ventas = final."),
        ("Mix de cobro: crédito + contado + anticipo = 100%", "=p_MixCobro", 1, 0.0001, "Si difiere, las ventas no se cobran completas."),
        ("CAPEX: costo final = ejecutado + comprometido + por contratar", f"=Inversion_Activos!$O${IAT['tot']}",
         f"=Inversion_Activos!$I${IAT['tot']}+Inversion_Activos!$L${IAT['tot']}+Inversion_Activos!$N${IAT['tot']}", 0.01, "Estados excluyentes."),
        ("CAPEX real: pagos mensuales cargados (T7) = pagado por ítem (T3)", f'=SUMIFS(dr_USD,dr_Cod,"CAPEX_PAG",dr_Est,"<>Excluido")',
         f"=Inversion_Activos!$J${IAT['tot']}", 1, "Debe conciliar cuando se carguen reales."),
        ("Meses reales cerrados sin registros cargados", "=k_chk_real", 0, 0, "Un faltante no equivale a cero."),
        ("Horizonte dentro de la rejilla mensual", "=k_chk_grid", 1, 0, f"Rejilla de {N} meses."),
        ("Reales cargados con mes real de arranque informado (1 = OK)", '=IF(AND(p_MesCorte>=0,p_ArranqueReal=""),0,1)', 1, 0,
         "Con reales, cargar T1 «mes real de arranque»; si no, el atraso del escenario podría contradecir la operación real."),
        ("Errores en indicadores clave (conteo)", f"=SUMPRODUCT(--ISERROR(Resultados_Anuales!$D${KPI['van_p']}:$D${KPI['_end'] - 1}))", 0, 0,
         "Indicadores no disponibles se informan con texto, no con cero."),
        ("Base depurada: ventas año 1 del motor = PDF (válido con la base sin modificar, sin choque)",
         f'=IF(AND(p_EscIdx=1,p_ShockIdx=1),Resultados_Anuales!$F${RA["v"]},"n/a")', f'=IF(AND(p_EscIdx=1,p_ShockIdx=1),{bp("ing_pub", "F")},"n/a")', 0.5,
         "Solo aplica con el escenario base y sin prueba de sensibilidad."),
        ("Base depurada: CxC al cierre del año 1 = ventas x 120/365 (validación)",
         f'=IF(AND(p_EscIdx=1,p_ShockIdx=1),Resultados_Anuales!$F${RA["cxc"]},"n/a")', f'=IF(AND(p_EscIdx=1,p_ShockIdx=1),{bp("ct_cxc_vtas", "F")},"n/a")', 1,
         "Con ventas mensuales uniformes el calendario de cobro reproduce la aproximación anual."),
    ]
    for i, (lab, calc, ref, tol, expl) in enumerate(ctrls):
        setc(ws, r, 1, f"V{i + 1:02d}")
        setc(ws, r, 2, lab)
        setc(ws, r, 3, calc, fmt=NF_USD4 if tol < 0.01 and tol > 0 else NF_USD2)
        setc(ws, r, 4, ref, fmt=NF_USD4 if tol < 0.01 and tol > 0 else NF_USD2, font=F_IN if isinstance(ref, (int, float)) else F_BASE)
        setc(ws, r, 5, f'=IF(OR(ISTEXT(C{r}),ISTEXT(D{r})),"n/a",C{r}-D{r})', fmt=NF_USD4 if tol < 0.01 and tol > 0 else NF_USD2)
        setc(ws, r, 6, tol, font=F_IN, fmt=NF_USD4 if tol < 0.01 and tol > 0 else NF_USD2)
        setc(ws, r, 7, f'=IF(E{r}="n/a","n/a",IF(ABS(E{r})<=F{r},"OK","REVISAR"))', font=F_BOLD)
        setc(ws, r, 8, expl)
        r += 1
    VAL["last"] = r - 1
    ws.conditional_formatting.add(f"G{VAL['first']}:G{VAL['last']}", CellIsRule(operator="equal", formula=['"OK"'], fill=FILL_OK))
    ws.conditional_formatting.add(f"G{VAL['first']}:G{VAL['last']}", CellIsRule(operator="equal", formula=['"REVISAR"'], fill=FILL_ALERT))
    setc(ws, r, 2, "Controles en estado REVISAR", font=F_BOLD)
    setc(ws, r, 3, f'=COUNTIF(G{VAL["first"]}:G{VAL["last"]},"REVISAR")', fmt=NF_INT0, font=F_BOLD)
    add_name("v_Revisar", "VAL", f"C{r}")
    r += 1
    setc(ws, r, 2, "Sin circularidades ni errores de fórmula: verificado con recálculo completo en LibreOffice (ver sección C).", font=F_SUB)
    r += 2

    section(ws, r, "B. Registro de inconsistencias del PDF (dato original, comprobación, diferencia, criterio propuesto y estado)", 8); r += 1
    header(ws, r, ["ID", "Inconsistencia (página)", "Dato publicado", "Comprobación", "Diferencia", "", "Estado", "Criterio propuesto y tratamiento en el modelo"]); r += 1
    VAL["inc_first"] = r
    incs = [
        ("TC de la nómina 8.000 vs 7.300 declarado (pp. 4-5, 30-34). Remuneración + prestaciones año 1", f"=Base_PDF!$F${BP['mod_8000']}+Base_PDF!$F${BP['moa_8000']}",
         f"=Base_PDF!$F${BP['mod_7300']}+Base_PDF!$F${BP['moa_7300']}", f"=Base_PDF!$F${BP['nom_fx_ef']}", "Corregido (C1)",
         "TC único por periodo; nómina en PYG convertida al TC del escenario. Efecto ≈ USD 59.730/año con 43,09%."),
        ("Margen 10b rotulado como margen pero es costo/precio (p. 20). Gabinetes año 1", f"={bp('mbp_g_pub', 'F')}", f"=Base_PDF!$F${BP['mg_g']}", None,
         "Corregido (rótulo)", "Margen = (precio - costo)/precio ≈ 42,8% (gabinetes) y 35,7% (cajas) sobre costo parcial. Rentabilidad por producto rotulada en Resultados_Anuales."),
        ("Cuentas por cobrar calculadas sobre costos (p. 24). Año 1", f"={bp('ct_cxc_pub', 'F')}", f"={bp('ct_cxc_vtas', 'F')}", f"={bp('ct_cxc_dif', 'F')}",
         "Corregido (motor)", "Clientes = facturación a crédito - cobranzas con calendario de cohortes (DSO). No se usa ventas x DSO/365 salvo como control."),
        ("MP de cajas sin 1% acumulativo (p. 2 vs p. 17). Año 10", f"={bp('mp_c_pub', 'O')}", f"={bp('mp_c_calc_acum', 'O')}", f"={bp('mp_c_dif_acum', 'O')}",
         "Corregido (C2) - confirmar", "Aplicar el texto (1% acumulativo) salvo que exista decisión comercial documentada."),
        ("VAN/TIR del cuadro 19 mezclan flujo del proyecto y del accionista (p. 27)", f"={bp('van_pub')}", f"={bp('van_fcff_pdf')}", None,
         "Corregido (motor)", "FCFF descontado a la tasa del proyecto y flujo de socios a Ke, por separado. El VAN publicado se conserva solo como referencia."),
        ("Payback parte de -2.791.587 y el VAN de -2.991.587 (p. 28)", -2791587, -2991587, 200000, "Corregido (motor)",
         "Misma serie para VAN, TIR y payback. Payback simple ≈ 5,71 y descontado ≈ 6,89 años sobre la serie publicada."),
        ("Aporte propio no termina en USD 200.000 (p. 26). Aportes adicionales vs ΔCT años 2-3", f"={bp('aportes_adic')}", f"={bp('dct_a23')}", None,
         "Documentado", "6,7%/93,3% describe solo el momento inicial. 939.710 = aportes adicionales ≠ ΔCT años 2-3 (1.087.311)."),
        ("Depreciación desde año 2 (tabla) / año 3 (nota) con ventas desde año 1 (p. 29)", f"={bp('p_DEP', 'F')}", f"=Resultados_Anuales!$F${RA['dep']}", None,
         "Corregido (C3)", "Depreciación desde la puesta en servicio de cada ítem."),
        ("CAPEX USD 100.000 del año 2 sin alta en el cuadro 20 (pp. 26-29)", f"={bp('capex2')}", 0, None, "Corregido (C4) - vida provisional",
         "Se incorpora como ítem ADD con vida útil provisional de 10 años (no informada). Confirmar clase y vida útil."),
        ("Recuperación de CT año 10: 3.974.293 vs CT cuadro 16 3.980.818 (pp. 24, 27)", f"={bp('rec_ct')}", f"={bp('ctn_pub', 'O')}", f"={bp('ct10_vs_rec')}",
         "Sin explicación en el PDF", "Valor terminal explícito: liquidación realizable o negocio en marcha (sin sumar ambos)."),
        ("Valor residual 59.417 = valor neto contable, no precio de venta (p. 27, 29)", f"={bp('vr_af')}", f"={bp('vnc_10')}", None, "Documentado",
         "Realización del activo fijo como parámetro (G24); requiere tasación."),
        ("Base directa 10b usa flete a USD 0,145/kg y difiere del 10a (pp. 19-21). Año 1", f"={bp('cd_g_pub', 'F')}+{bp('cd_c_pub', 'F')}",
         f"={bp('p_MP', 'F')}+{bp('p_INS', 'F')}+{bp('p_CON', 'F')}+{bp('p_FLE', 'F')}", f"={bp('base10b_dif', 'F')}", "Documentado",
         "El 10a implica USD 23,20/gabinete y USD 2,08/caja (ajuste exacto 4 años). Se usa el 10a; confirmar con el origen."),
        ("Caja con pérdida asignada año 1 por prorrateo de costos comunes (p. 20)", f"={bp('ing_c_pub', 'F')}", f"={bp('ctot_c_pub', 'F')}", f"={bp('res_c_asig', 'F')}",
         "Documentado", "La contribución directa es positiva. No eliminar el producto solo por absorber costos comunes; ver contribución y driver."),
        ("Capacidad: texto +10% vs tabla +46,7% gabinetes / +34,4% cajas (pp. 3, 14)", 0.10, f"={bp('capinc_g', 'H')}", None, "Pendiente de respaldo técnico",
         "Vincular capacidad con equipos, turnos y cuellos de botella; E25 permite capacidad adicional explícita."),
        ("Costos constantes con producción creciente (energía, salarios, insumos) (p. 19)", f"={bp('p_ENE', 'F')}", f"={bp('p_ENE', 'O')}", None, "Supuesto histórico",
         "Drivers editables: energía variable (E90), ajuste salarial (E46), inflación de fijos (E47), productividad (E45)."),
        ("Equilibrio en unidades mezcladas (p. 25). Año 1", f"={bp('pe_pub', 'F')}", f"={bp('pe_u', 'F')}", None, "Corregido (presentación)",
         "Equilibrio en ventas con mix del año y unidades por producto (Resultados_Anuales)."),
        ("Tributo 1% s/ ventas e IVA como costo (pp. 2, 7, 22)", f"={bp('trib')}", "Ley 7547/2025", None, "Pendiente de verificación oficial",
         "Se conserva como histórico. Base alternativa: mayor entre factura y valor agregado nacional (G43) e IVA recuperable (E78) como escenarios."),
        ("Alquiler llamado «costo de oportunidad» (p. 6)", f"={bp('p_ALQ', 'F')}", "Naturaleza", None, "Pendiente de confirmar",
         "G40: pago a tercero, a relacionada o costo económico sin desembolso (se mantiene en resultados y FCFF)."),
        ("Gastos de ventas en cero (p. 19)", 0, "Quién asume", None, "Pendiente de confirmar",
         "E89 permite cargar flete de exportación, seguros, comisiones, garantías y rechazos."),
        ("Fila «% amortizaciones» del cuadro 13 inconsistente con los montos (p. 22). Año 1", f"={bp('pctam_pub', 'F')}", f"={bp('pctam_calc', 'F')}", None,
         "Corregido", "Se recalcula sobre el cronograma mensual reproducido."),
        ("Prestaciones: rótulo 43,09% vs ratio implícito ≈ 43,087% (p. 19)", f"={bp('prest')}", f"={bp('p_PRD', 'F')}/{bp('p_MOD', 'F')}", None, "Diferencia de redondeo",
         "Coeficiente del plan, no tasa legal. Desagregar componentes cuando se aporten."),
        ("Proveedores menores 30 días: base de cálculo no identificada (p. 24)", f"={bp('ct_pmen_pub', 'F')}", "No reproducible", None, "Documentado",
         "El motor paga a 30 días todos los costos no salariales ni de materiales."),
        ("Cuadro 9a rotula «MP e insumos» pero incluye solo materia prima (p. 17)", f"={bp('p_MP', 'F')}", f"={bp('mp_g_pub', 'F')}+{bp('mp_c_pub', 'F')}", None, "Documentado", "Sin efecto numérico."),
        ("Cuadro 17 clasifica energía y mantenimiento como variables (p. 25)", f"={bp('cv_pub', 'F')}", "Clasificación", None, "Corregido (clasificación)",
         "Energía fija con porción variable editable; mantenimiento crece con el tiempo, no con el volumen."),
        ("Tasa ponderada con pesos iniciales (6,7%/93,3%) para 10 años; IVA de intereses (p. 28)", f"={bp('w_pub')}", f"={bp('tasa_exacta')}", None, "Pendiente de actualizar",
         "Métodos 2 (WACC con estructura objetivo) o 3 (manual) en G12-G17; sin escudo de IRE."),
        ("Dividendos devengados vs pagados con un año de desfase (pp. 23, 26)", f"={bp('div_decl_pub', 'O')}", f"={bp('u_div', 'O')}", None, "Corregido (motor)",
         "Declaración al cierre y pago con desfase (G53) solo si hay caja sobre el mínimo (G54)."),
        ("Stock y PT: PT valuado con gastos administrativos; stock sin flete (p. 24)", f"={bp('ct_pt_pub', 'F')}", f"=Resultados_Anuales!$F${RA['pt']}", None, "Corregido (C5)",
         "PT a costo de producción (materiales + conversión asignada); materiales a costo de adquisición."),
    ]
    for i, (lab, pub, chk, dif, st, crit) in enumerate(incs):
        setc(ws, r, 1, f"I{i + 1:02d}")
        setc(ws, r, 2, lab, align=WRAP)
        for c, v in [(3, pub), (4, chk)]:
            fmt = NF_PCT2 if (isinstance(v, float) and v < 1) else NF_USD
            if isinstance(v, str) and ("mbp" in v or "mg_" in v or "pctam" in v or "capinc" in v or "prest" in v or "trib" in v or "w_pub" in v or "tasa_exacta" in v or "p_PRD" in v):
                fmt = NF_PCT2
            setc(ws, r, c, v, fmt=fmt, font=F_LINK if isinstance(v, str) and v.startswith("=") else F_IN)
        if dif is not None:
            setc(ws, r, 5, dif, fmt=NF_USD, font=F_LINK if isinstance(dif, str) else F_IN)
        else:
            setc(ws, r, 5, f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),D{r}-C{r},"")', fmt=NF_USD2)
        setc(ws, r, 7, st, font=F_BOLD)
        setc(ws, r, 8, crit, align=WRAP)
        ws.row_dimensions[r].height = 30
        r += 1
    VAL["inc_last"] = r - 1
    r += 1
    section(ws, r, "C. Pruebas del motor realizadas antes de la entrega (capturas; ver Escenarios_Sensib para verificación en vivo)", 8); r += 1
    header(ws, r, ["ID", "Prueba", "Resultado base", "Resultado con cambio", "Variación", "", "Estado", "Comportamiento esperado / observado"]); r += 1
    VAL["test_first"] = r
    for i, t in enumerate(CAPT.get("tests", [])):
        setc(ws, r, 1, f"P{i + 1:02d}")
        setc(ws, r, 2, t["name"], align=WRAP)
        setc(ws, r, 3, t.get("base"), fmt=t.get("fmt", NF_USD), font=F_IN)
        setc(ws, r, 4, t.get("new"), fmt=t.get("fmt", NF_USD), font=F_IN)
        if isinstance(t.get("base"), (int, float)) and isinstance(t.get("new"), (int, float)):
            setc(ws, r, 5, f"=D{r}-C{r}", fmt=t.get("fmt", NF_USD))
        setc(ws, r, 7, t.get("status", ""), font=F_BOLD)
        setc(ws, r, 8, t.get("note", ""), align=WRAP)
        ws.row_dimensions[r].height = 30
        r += 1
    if not CAPT.get("tests"):
        setc(ws, r, 2, "Pruebas pendientes de captura.", font=F_ALERT)
        r += 1
    setc(ws, r, 2, CAPT.get("engine_note", "Capturas generadas con el motor de cálculo del entorno."), font=F_SUB)
    ws.freeze_panes = "C4"


# ===========================================================================
# 7) ESCENARIOS_SENSIB: capturas verificables
# ===========================================================================
CAPT = {}
CAPT_KEYS = [("van_p", "VAN proyecto", NF_USD), ("tir_p", "TIR proyecto", NF_PCT2), ("van_e", "VAN socios", NF_USD), ("tir_e", "TIR socios", NF_PCT2),
             ("e1", "EBITDA año 1", NF_USD), ("et", "EBITDA total", NF_USD), ("nec", "Pico necesidad adicional", NF_USD),
             ("brecha_nc", "Brecha no cubierta máx.", NF_USD), ("dscr_min", "DSCR mínimo", NF_X), ("pb_s", "Payback simple (años)", NF_USD2),
             ("ctn_max", "CT operativo máx.", NF_USD)]


def build_es():
    ws = WS["ES"]
    CUR[0] = "ES"
    title(ws, "Escenarios_Sensib — comparativas CAPTURADAS del motor (no se recalculan solas) y puente de correcciones",
          "Las tablas son capturas estáticas identificadas. Para actualizarlas: seleccionar el caso en Supuestos, recalcular y copiar los "
          "indicadores de Resultados_Anuales. La columna «Verificación en vivo» compara la captura con el motor cuando la selección coincide.")
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 44
    for c in range(3, 3 + len(CAPT_KEYS) + 2):
        ws.column_dimensions[CL(c)].width = 15
    stamp = CAPT.get("stamp", "SIN CAPTURA")
    r = 4
    setc(ws, r, 2, f"Captura: {stamp}", font=F_ALERT)
    r += 2

    def table(title_txt, rows, kind):
        nonlocal r
        section(ws, r, title_txt, 3 + len(CAPT_KEYS) + 1); r += 1
        header(ws, r, ["N°", "Caso"] + [lab for _, lab, _ in CAPT_KEYS] + ["Verificación en vivo (VAN proyecto: motor - captura)"])
        ws.row_dimensions[r].height = 36
        r += 1
        for i, row in enumerate(rows):
            setc(ws, r, 1, i + 1, fmt=NF_INT0)
            setc(ws, r, 2, row["name"])
            for j, (k, lab, fmt) in enumerate(CAPT_KEYS):
                v = row["vals"].get(k)
                setc(ws, r, 3 + j, v, fmt=fmt if isinstance(v, (int, float)) else None, font=F_IN)
            cond = row["cond"]
            setc(ws, r, 3 + len(CAPT_KEYS), f'=IF({cond},k_van_p-C{r},"Seleccionar este caso para verificar")', fmt=NF_USD2)
            r += 1
        r += 1

    if CAPT.get("scen"):
        table("A. Escenarios (sin prueba de sensibilidad, 6 correcciones activas, sin reales)", CAPT["scen"], "scen")
    if CAPT.get("shock"):
        table("B. Pruebas de sensibilidad sobre la Base depurada (choques ilustrativos, no premisas de negocio)", CAPT["shock"], "shock")
    if CAPT.get("bridge"):
        section(ws, r, "C. Puente del VAN del proyecto: del motor con criterios del PDF (C1-C6 = 0) a la Base depurada", 14); r += 1
        setc(ws, r, 2, "Orden de aplicación acumulativo C1 → C6. El efecto de cada paso depende del orden (interacciones); la suma de pasos concilia por construcción. "
                       "La columna «Efecto aislado» activa solo esa corrección sobre el motor con criterios del PDF; la diferencia con la suma es la interacción.", font=F_SUB)
        r += 1
        header(ws, r, ["Paso", "Estado del motor", "VAN proyecto", "Efecto del paso", "Efecto aislado", "VAN socios", "EBITDA total", "Pico necesidad adicional",
                       "Verificación en vivo"])
        ws.row_dimensions[r].height = 30
        r += 1
        b0 = r
        for i, row in enumerate(CAPT["bridge"]):
            setc(ws, r, 1, i, fmt=NF_INT0)
            setc(ws, r, 2, row["name"])
            setc(ws, r, 3, row["van_p"], fmt=NF_USD, font=F_IN)
            if i > 0:
                setc(ws, r, 4, f"=C{r}-C{r - 1}", fmt=NF_USD)
                setc(ws, r, 5, row.get("iso"), fmt=NF_USD, font=F_IN)
            setc(ws, r, 6, row["van_e"], fmt=NF_USD, font=F_IN)
            setc(ws, r, 7, row["et"], fmt=NF_USD, font=F_IN)
            setc(ws, r, 8, row["nec"], fmt=NF_USD, font=F_IN)
            setc(ws, r, 9, f'=IF(AND({row["cond"]}),k_van_p-C{r},"Seleccionar esta combinación de C1-C6")', fmt=NF_USD2)
            r += 1
        setc(ws, r, 2, "Suma de efectos de los pasos (concilia con la diferencia total)", font=F_BOLD)
        setc(ws, r, 4, f"=SUM(D{b0 + 1}:D{r - 1})", fmt=NF_USD, font=F_BOLD)
        setc(ws, r, 5, f"=SUM(E{b0 + 1}:E{r - 1})", fmt=NF_USD, font=F_BOLD)
        setc(ws, r, 3, f"=C{r - 1}-C{b0}", fmt=NF_USD, font=F_BOLD)
        r += 1
        setc(ws, r, 2, "Interacción = suma de pasos - suma de efectos aislados", font=F_BOLD)
        setc(ws, r, 4, f"=D{r - 1}-E{r - 1}", fmt=NF_USD)
        r += 2
        section(ws, r, "D. Conciliación metodológica del VAN publicado (anual, datos del PDF)", 14); r += 1
        rows = [("VAN publicado — cuadro 19 (mezcla proyecto y accionista)", f"={bp('van_calc')}"),
                ("VAN FCFF con datos PDF (excluye préstamo y servicio de deuda), misma tasa", f"={bp('van_fcff_pdf')}"),
                ("VAN de los socios con datos PDF a Ke 16,32%", f"={bp('van_eq_pdf')}"),
                ("VAN del motor con criterios del PDF (C1-C6 = 0), mensual — captura", CAPT["bridge"][0]["van_p"]),
                ("VAN de la Base depurada (motor, todas las correcciones) — captura", CAPT["bridge"][-1]["van_p"])]
        for lab, v in rows:
            setc(ws, r, 2, lab)
            setc(ws, r, 3, v, fmt=NF_USD, font=F_LINK if isinstance(v, str) else F_IN)
            r += 1
        setc(ws, r, 2, "Del paso 1 al 2: cambio de perspectiva (FCFF). Del 2 al 4: calendario mensual del CT (clientes sobre ventas, stock anticipado, "
                       "PT a costo), valor terminal por liquidación y oportunidad de flujos. Del 4 al 5: correcciones C1-C6.", font=F_SUB)
        r += 1
    ws.freeze_panes = "C4"


# ===========================================================================
# 8) RESUMEN
# ===========================================================================
def build_resumen():
    ws = WS["RES"]
    CUR[0] = "RES"
    title(ws, "Resumen ejecutivo — Maquila CIE-PERG: procesamiento de chapas para gabinetes y cajas metálicas (USD)")
    setc(ws, 2, 1, '="Escenario: "&p_Escenario&"  |  Prueba: "&p_Shock&"  |  "&INDEX(Supuestos!$D$%d:$H$%d,1,p_EscIdx)' % (SUP["scen_status"], SUP["scen_status"]),
         font=F_ALERT)
    for k, v in {1: 4, 2: 64, 3: 20, 4: 20, 5: 16, 6: 70}.items():
        ws.column_dimensions[CL(k)].width = v
    r = 4
    section(ws, r, "A. Estado del modelo", 6); r += 1
    items = [("Escenario activo", "=p_Escenario"), ("Prueba de sensibilidad activa", "=p_Shock"),
             ("Correcciones metodológicas activas (de 6)", "=p_NCorr"),
             ("Último mes cerrado con datos reales", '=IF(p_MesCorte<0,"Ninguno: proyección con supuestos históricos",p_MesCorte)'),
             ("Campos mínimos para forecast actualizado", "=r_MinCompletos"),
             ("Calendario", '=IF(p_FechaInicio="","PENDIENTE: meses relativos (sin fecha del mes 0)","Mes 0 = "&MONTH(p_FechaInicio)&"/"&YEAR(p_FechaInicio))'),
             ("Tasa del proyecto aplicada", '=FIXED(p_TasaProy*100,2)&"% — "&p_TasaEstado'),
             ("Ke aplicado (histórico)", '=FIXED(p_Ke*100,2)&"%"'),
             ("Controles de Validaciones en estado REVISAR", "=v_Revisar")]
    for lab, f in items:
        setc(ws, r, 2, lab)
        setc(ws, r, 3, f, font=F_BOLD)
        r += 1
    r += 1
    section(ws, r, "B. PDF publicado (histórico) vs escenario activo recalculado", 6); r += 1
    header(ws, r, ["", "Indicador", "PDF publicado Oct-2025 (histórico)", "Escenario activo (motor)", "Diferencia", "Nota"]); r += 1
    RES_T = r
    comp = [
        ("Inversión fija inicial / CAPEX estimado total", f"={bp('inv_fija')}", "=k_capex_e", NF_USD, "El motor incluye el CAPEX adicional del año 2 (USD 100.000)."),
        ("Capital operativo inicial (PDF) / CT operativo neto máximo (motor)", f"={bp('ct_ini')}", "=k_ctn_max", NF_USD,
         "Conceptos distintos: el PDF suma CT de un año al inicio; el motor lo construye mes a mes."),
        ("Ventas año 1", f"={bp('ing_pub', 'F')}", "=k_v1", NF_USD, "Año de proyecto 1."),
        ("Ventas totales 10 años", f"=Base_PDF!$P${BP['ing_pub']}", "=k_vt", NF_USD, ""),
        ("EBITDA año 1", f"={bp('ebitda_pub', 'F')}", "=k_e1", NF_USD, "Tributo de maquila incluido una sola vez dentro del EBITDA."),
        ("EBITDA total 10 años", f"=Base_PDF!$P${BP['ebitda_pub']}", "=k_et", NF_USD, ""),
        ("Resultado neto total", f"=Base_PDF!$P${BP['ut_pub']}", "=k_nit", NF_USD, ""),
        ("VAN (PDF: cuadro 19 mezcla perspectivas / motor: FCFF del proyecto)", f"={bp('van_pub')}", "=k_van_p", NF_USD,
         "PUBLICADO, METODOLOGÍA POR DEPURAR vs VAN del proyecto. No son comparables en sentido estricto."),
        ("TIR (PDF «con financiamiento» / motor: TIR del proyecto)", f"={bp('tir_pub')}", "=k_tir_p", NF_PCT2, "Una TIR alta no implica decisión conveniente sin VAN, caja y riesgo."),
        ("VAN de los socios a Ke (PDF reexpresado / motor)", f"={bp('van_eq_pdf')}", "=k_van_e", NF_USD, "Aportes, dividendos pagados y valor final."),
        ("TIR de los socios (PDF reexpresado / motor)", f"={bp('tir_eq_pdf')}", "=k_tir_e", NF_PCT2, ""),
        ("Payback descontado (PDF: serie desde -2.791.587 / motor: FCFF)", f"={bp('pb_pub')}", "=k_pb_d", NF_USD2, "Años. PDF «durante el 7mo año»; sobre la serie del VAN ≈ 6,89. Motor simple en Resultados_Anuales."),
        ("Aportes de socios en efectivo (total)", f"={bp('aporte_ini')}+{bp('aportes_adic')}", "=k_ap", NF_USD, ""),
        ("Pico de necesidad de fondos adicionales (antes de aportes adicionales)", f"={bp('aportes_adic')}", "=k_nec", NF_USD,
         "PDF: aportes adicionales planificados. Motor: necesidad mensual acumulada incluida la caja mínima."),
        ("Brecha no cubierta máxima (con financiamiento planificado)", 0, "=k_brecha_nc", NF_USD, "El motor no inserta fondos ficticios: la brecha queda visible."),
        ("Deuda financiera máxima", f"={bp('prestamo')}", "=k_deuda_max", NF_USD, ""),
        ("DSCR mínimo (CFADS / servicio de deuda)", "n/d en el PDF", "=k_dscr_min", NF_X, "EBITDA/servicio año 1 PDF ≈ 1,08x (aproximación, no DSCR)."),
        ("Equilibrio año 1 (PDF: unidades mezcladas / motor: ventas USD)", f"={bp('pe_pub', 'F')}", "=k_pe1", NF_USD,
         "Motor: ver gabinetes y cajas en equilibrio con el mix del año."),
    ]
    for lab, pdf, act, fmt, note in comp:
        setc(ws, r, 2, lab)
        setc(ws, r, 3, pdf, fmt=fmt, font=F_LINK if isinstance(pdf, str) and pdf.startswith("=") else F_IN, fill=FILL_HIST)
        setc(ws, r, 4, act, fmt=fmt, font=F_BOLD)
        setc(ws, r, 5, f'=IF(AND(ISNUMBER(C{r}),ISNUMBER(D{r})),D{r}-C{r},"")', fmt=fmt)
        setc(ws, r, 6, note, font=F_SUB)
        r += 1
    RES_END = r - 1
    r += 1
    section(ws, r, "C. Alertas específicas", 6); r += 1
    alerts = [
        '=IF(p_EscIdx=1,"Base depurada: supuestos de Oct-2025 con correcciones metodológicas; NO es situación real de 2026.",'
        'IF(p_EscIdx=2,INDEX(Supuestos!$D$%d:$H$%d,1,2),"Escenario "&p_Escenario&": "&INDEX(Supuestos!$D$%d:$H$%d,1,p_EscIdx)))' % ((SUP["scen_status"],) * 4),
        '=IF(p_ShockIdx>1,"PRUEBA DE SENSIBILIDAD ACTIVA («"&p_Shock&"»): los resultados son un choque ilustrativo. Volver a «Ninguna».","")',
        '=IF(k_brecha_nc>0.5,"Brecha de caja no cubierta: máx. USD "&FIXED(k_brecha_nc,0)&" en "&k_brecha_nc_n&" meses. El financiamiento planificado no alcanza; definir aportes, deuda o mejoras de CT.","Sin brecha de caja no cubierta con el financiamiento planificado.")',
        '="Pico de necesidad antes de aportes adicionales: USD "&FIXED(k_nec,0)&" en el mes "&k_nec_m&" (PDF planificó USD 939.710 de aportes adicionales)."',
        '=IF(k_chk_cap>0,k_chk_cap&" meses con producción sobre la capacidad del plan: requiere horas extra, turno o inversión explícitos.","")',
        '=IF(ISNUMBER(k_dscr_min),IF(k_dscr_min<1,"DSCR mínimo "&FIXED(k_dscr_min,2)&"x: el flujo operativo (incluida la formación de CT del arranque) no cubre el servicio de deuda en algún año; se financia con aportes o caja.",""),"")',
        '=IF(ISNUMBER(k_tir_p),IF(LEFT(k_tir_p_u,5)<>"Única","TIR del proyecto con posibles valores múltiples: priorizar VAN.",""),"TIR del proyecto no disponible: "&k_tir_p)',
        '=IF(p_MesCorte<0,"Sin datos reales cargados: la situación desde Oct-2025 no está incorporada. Ver Datos_Reales.","")',
        '=IF(p_FechaInicio="","Calendarización pendiente: el año 1 no se asocia a 2025 ni a 2026 hasta cargar la fecha del mes 0.","")',
        '=IF(p_MetodoTasa=1,"Tasa del proyecto = ponderada inicial del PDF (93% deuda a 9%): hipótesis histórica, no WACC vigente.","")',
        '=IF(OR(p_BaseTrib=1,p_Modalidad=1),"Tributo de maquila, IVA y modalidad contractual: criterio histórico del PDF, pendiente de verificación con Ley 7547/2025 y el contrato.","")',
        '=IF(p_NCorr<6,"Correcciones metodológicas incompletas ("&p_NCorr&" de 6): el resultado no es la base depurada.","")',
        '=IF(k_chk_bal+k_chk_ind>0.01,"ERROR DE INTEGRIDAD: balance o método directo/indirecto no cuadran. Revisar Validaciones.","")',
        '=IF(v_Revisar>0,v_Revisar&" controles en estado REVISAR (hoja Validaciones).","")',
        '=IF(p_ApAuto=1,"Aportes hipotéticos de cobertura ACTIVOS (no aprobados): USD "&FIXED(k_ap_auto,0)&" incluidos en el flujo de socios.",IF(k_brecha_nc>0.5,"El VAN/TIR de socios supone que la brecha se cubre sin costo; activar Supuestos G63 para medir el aporte adicional requerido.",""))',
        '=IF(AND(p_MesCorte>=0,p_ArranqueReal=""),"Hay reales cargados sin «mes real de arranque» (Datos_Reales T1): el atraso del escenario puede contradecir la operación real.","")',
    ]
    for f in alerts:
        setc(ws, r, 2, f, font=F_ALERT)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        r += 1
    r += 1
    section(ws, r, "D. Lectura para la decisión", 6); r += 1
    txt = ["• Esta versión reconstruye el plan de octubre de 2025, corrige su metodología y permite simular. No acredita la viabilidad actual del negocio: "
           "no se cargó ejecución real, compromisos vigentes ni condiciones contractuales.",
           "• Rentabilidad, liquidez y creación de valor se leen por separado: VAN/TIR del proyecto (FCFF), VAN/TIR de socios (Ke), brecha de caja y DSCR.",
           "• La base depurada mantiene ventas y precios del PDF; las diferencias de EBITDA y caja provienen de correcciones (C1-C6) y del calendario mensual del capital de trabajo.",
           "• La recomendación actualizada debe emitirse cuando se incorporen reales, CAPEX comprometido, contrato de maquila y tasas vigentes."]
    for t in txt:
        setc(ws, r, 2, t, align=WRAP)
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=6)
        ws.row_dimensions[r].height = 28
        r += 1
    r += 1
    # datos para el gráfico de CAPEX
    CH = r
    section(ws, r, "E. Gráficos (años de proyecto 1 a 10)", 6); r += 1
    setc(ws, r, 2, "CAPEX presupuesto original")
    setc(ws, r, 3, "=k_capex_o", fmt=NF_USD)
    r += 1
    setc(ws, r, 2, "CAPEX costo final estimado")
    setc(ws, r, 3, "=k_capex_e", fmt=NF_USD)
    capex_rows = (r - 1, r)
    r += 1
    add_charts(ws, r + 2, capex_rows)
    ws.freeze_panes = "A4"


def add_charts(ws, anchor_row, capex_rows):
    from openpyxl.chart import BarChart, LineChart, Reference
    ra = WS["RA"]
    cats = Reference(ra, min_col=RAY0 + 1, max_col=RAY0 + 10, min_row=RA["_end"] + 2)
    # fila auxiliar de etiquetas de años en RA
    lr = RA["_end"] + 2
    for y in range(1, 11):
        ra.cell(row=lr, column=RAY0 + y, value=f"Año {y}")
    ra.cell(row=lr, column=2, value="Etiquetas de gráficos (no editar)").font = F_SUB

    def ref(key):
        return Reference(ra, min_col=RAY0 + 1, max_col=RAY0 + 10, min_row=RA[key])

    def mk_bar(title_, keys, names, anchor, line_keys=None, line_names=None):
        ch = BarChart()
        ch.title = title_
        ch.y_axis.title = "USD"
        ch.height, ch.width = 7.5, 16
        for k, nm in zip(keys, names):
            ch.add_data(ref(k), from_rows=True, titles_from_data=False)
            ch.series[-1].tx = None
        for s, nm in zip(ch.series, names):
            from openpyxl.chart.series import SeriesLabel
            s.tx = SeriesLabel(v=nm)
        ch.set_categories(cats)
        if line_keys:
            ln = LineChart()
            for k, nm in zip(line_keys, line_names):
                ln.add_data(ref(k), from_rows=True, titles_from_data=False)
            from openpyxl.chart.series import SeriesLabel
            for s, nm in zip(ln.series, line_names):
                s.tx = SeriesLabel(v=nm)
                s.smooth = False
            ch += ln
        ws.add_chart(ch, anchor)

    def mk_line(title_, keys, names, anchor):
        from openpyxl.chart.series import SeriesLabel
        ch = LineChart()
        ch.title = title_
        ch.y_axis.title = "USD"
        ch.height, ch.width = 7.5, 16
        for k in keys:
            ch.add_data(ref(k), from_rows=True, titles_from_data=False)
        for s, nm in zip(ch.series, names):
            s.tx = SeriesLabel(v=nm)
        ch.set_categories(cats)
        ch.x_axis.tickLblPos = "low"
        for s_ in ch.series:
            s_.smooth = False
        ws.add_chart(ch, anchor)

    mk_bar("Ventas y EBITDA — escenario activo vs PDF", ["v", "v_pdf"], ["Ventas (activo)", "Ventas (PDF)"], f"B{anchor_row}",
           ["ebitda", "ebitda_pdf"], ["EBITDA (activo)", "EBITDA (PDF)"])
    mk_line("Caja al cierre y necesidad máxima de financiamiento", ["caja", "brecha", "nec_pre"],
            ["Caja al cierre", "Necesidad mensual máx.", "Necesidad antes de aportes adicionales"], f"D{anchor_row}")
    mk_line("Deuda financiera al cierre", ["deuda"], ["Deuda"], f"B{anchor_row + 16}")
    mk_bar("Contribución marginal por producto", ["pg_contr", "pc_contr"], ["Gabinetes", "Cajas"], f"D{anchor_row + 16}")
    from openpyxl.chart.series import SeriesLabel
    ch = BarChart()
    ch.title = "Inversión: presupuesto original vs costo final estimado"
    ch.height, ch.width = 7.5, 16
    ch.add_data(Reference(ws, min_col=3, min_row=capex_rows[0], max_row=capex_rows[1]), titles_from_data=False)
    ch.set_categories(Reference(ws, min_col=2, min_row=capex_rows[0], max_row=capex_rows[1]))
    ch.series[0].tx = SeriesLabel(v="USD")
    ws.add_chart(ch, f"B{anchor_row + 32}")


# ===========================================================================
# 9) GUÍA
# ===========================================================================
def build_guia():
    ws = WS["GUIA"]
    CUR[0] = "GUIA"
    title(ws, "Modelo financiero de maquila CIE-PERG — guía de uso",
          "Base documental: plan «Procesamiento de chapas para gabinetes metálicos» (Mujica & Saldivar, Oct-2025) e informe de validación (29/09/2026).")
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 120
    r = 4
    blocks = [
        ("Propósito", [
            ("", "Reconstruir el plan publicado, documentar sus inconsistencias y convertirlo en un modelo de actualización continua: reales + forecast, "
                 "escenarios y evaluación de rentabilidad, liquidez, financiamiento y creación de valor por separado."),
            ("", "Tres capas: (1) PDF publicado Oct-2025 (Base_PDF, inalterado); (2) Base depurada (supuestos históricos + correcciones C1-C6 + motor mensual); "
                 "(3) Reales + forecast (Datos_Reales hasta el último mes cerrado; escenario activo para el futuro).")]),
        ("Hojas", [
            ("Resumen", "Tablero ejecutivo: PDF vs escenario activo, alertas y gráficos."),
            ("Supuestos", "Selector único de escenario, prueba de sensibilidad, parámetros generales, préstamo, correcciones metodológicas, drivers por escenario, dividendos y aportes."),
            ("Calendario", "Meses del modelo (0 = inicio de implantación), años de proyecto y operativos, banderas de operación/horizonte/reales, TC y factores de descuento."),
            ("Produccion_Ventas", "Capacidad, demanda, rampa, recuperación de volumen, producción buena, rechazos, inventario de PT en unidades, precios y ventas."),
            ("Costos_Personal", "Nómina por puesto (PYG), maestro de costos fijos con moneda de origen, materiales, costos variables, fijos y resúmenes."),
            ("Inversion_Activos", "Presupuesto original, ejecutado, comprometido, por contratar, costo final estimado, pagos, altas y depreciación por ítem."),
            ("Capital_Trabajo", "Clientes por cohortes (DSO interpolado), inventarios, proveedores, anticipos y CT operativo neto (sin caja mínima)."),
            ("Deuda_Tributos", "IVA recuperable, préstamo original (reproduce el PDF), tramo 2 opcional y aportes."),
            ("Resultados_Caja", "Estado de resultados, caja por método directo con control indirecto, tesorería (caja mínima, brecha, dividendos, línea hipotética), balance de control y flujos de valuación."),
            ("Resultados_Anuales", "Agregación anual, comparación con el PDF, indicadores (VAN/TIR proyecto y socios, payback, DSCR, equilibrio, rentabilidad por producto)."),
            ("Escenarios_Sensib", "Capturas identificadas de escenarios, pruebas de sensibilidad y puente de correcciones, con verificación en vivo."),
            ("Datos_Reales", "Tablas vacías para fechas, saldos al corte, CAPEX por ítem, movimientos mensuales, documentación y campos mínimos."),
            ("Base_PDF", "Datos publicados, reconstrucción de los cuadros 1-20, cronograma del préstamo, VAN/TIR/payback reproducidos y reexpresión de perspectivas."),
            ("Validaciones", "Controles con tolerancia, registro de inconsistencias del PDF y pruebas realizadas.")]),
        ("Convenciones", [
            ("Colores", "Azul sobre amarillo claro = entrada editable. Amarillo intenso = pendiente de cargar. Verde = vínculo a otra hoja. Negro = fórmula. "
                        "Gris itálica = valor heredado de la base (sobrescribir para cambiar). Salmón = dato histórico del PDF."),
            ("Moneda", "Presentación en USD. Nómina y varios costos fijos tienen origen en PYG y se convierten con el TC de cada año; compras en USD no cambian con el TC. "
                       "Nunca se suman PYG con USD."),
            ("Tiempo", "Mes 0 = inicio de implantación/desembolso. Sin fecha del mes 0 el modelo trabaja en meses relativos (no se asume año 1 = 2025 o 2026)."),
            ("Signos", "Costos y pagos positivos en sus filas; los flujos (FCFF, socios) llevan signo: negativo = salida.")]),
        ("Cómo usar", [
            ("Cambiar escenario", "Supuestos!D5. Los escenarios Conservador y Recuperación son ILUSTRATIVOS. Actualizado y Personalizado heredan la base hasta que se sobrescriban (columna E/H)."),
            ("Prueba de sensibilidad", "Supuestos!D8 aplica un choque sobre el escenario activo y recalcula todo (tributos, CT, financiamiento e intereses). Volver a «Ninguna»."),
            ("Correcciones", "Supuestos sección C: poner 0 reproduce el criterio del PDF de esa corrección (para el puente de conciliación)."),
            ("Cierre mensual", "1) Cargar movimientos del mes en Datos_Reales T7 (código, mes, moneda, importe, TC, fuente, estado). 2) Cargar saldos al corte en T2 y CAPEX por ítem en T3. "
                               "3) Actualizar T1 «Último mes cerrado». 4) Revisar Validaciones (meses reales sin registros, ajustes de conciliación a patrimonio). "
                               "5) Ajustar supuestos futuros en la columna Actualizado. 6) Revisar Resumen y alertas."),
            ("Congelamiento", "Los meses ≤ último mes cerrado usan solo reales; los escenarios afectan únicamente el futuro. Una diferencia entre saldos reales y calculados "
                              "se registra como «ajuste de conciliación» en patrimonio y debe investigarse.")]),
        ("Métricas", [
            ("VAN/TIR del proyecto", "FCFF: flujo operativo después de tributos - CAPEX - ΔCT operativo, sin préstamos, intereses, aportes ni dividendos. Tasa del proyecto (G17), mensual equivalente."),
            ("VAN/TIR de socios", "Aportes (negativos), dividendos efectivamente pagados y valor final atribuible (caja + activos netos - deuda), descontados a Ke. No suma FCFE no distribuido."),
            ("Payback", "Simple y descontado sobre la misma serie; «No recupera en el horizonte» si no cruza."),
            ("Necesidad de financiamiento", "Pico antes de aportes adicionales (caja acumulada con préstamo original y aporte inicial vs caja mínima) y brecha no cubierta después del financiamiento planificado."),
            ("CFADS y DSCR", "CFADS = flujo operativo - CAPEX de mantenimiento. DSCR = CFADS / (capital + intereses + IVA + comisiones). EBITDA/servicio es solo una aproximación."),
            ("Equilibrio", "En ventas con el mix del año y unidades por producto; no se mezclan gabinetes y cajas como unidades equivalentes."),
            ("Valor terminal", "Liquidación realizable (por defecto) o negocio en marcha; nunca ambos.")]),
        ("Registro de avance", [
            ("Versión", "v1.0 — 29/09/2026. Libro construido con generar_modelo.py (reproducible)."),
            ("Terminado", "Base_PDF reproducida (diferencias ≤ redondeo), motor mensual integrado con balance y método directo/indirecto cuadrados, 5 escenarios, 16 pruebas, "
                          "puente de correcciones, registro de inconsistencias y tablas de reales."),
            ("Último control", "Ver Validaciones: controles V01-V33 y pruebas P01 en adelante (recálculo completo sin errores ni circularidades)."),
            ("Pendiente", "Datos reales desde Oct-2025, fecha del mes 0, CAPEX ejecutado/comprometido, contrato y modalidad de maquila, verificación de Ley 7547/2025 y Decreto 5714/2026, "
                          "naturaleza del alquiler, gastos de venta, tasas vigentes, TC real/proyectado, vida útil del CAPEX adicional, respaldo técnico de capacidad.")]),
    ]
    for ttl, rows in blocks:
        section(ws, r, ttl, 3); r += 1
        for a, b in rows:
            setc(ws, r, 2, a, font=F_BOLD)
            setc(ws, r, 3, b, align=WRAP)
            ws.row_dimensions[r].height = max(15, 13 * (1 + len(b) // 130))
            r += 1
        r += 1


# ===========================================================================
# 10) EJECUCIÓN
# ===========================================================================
def main():
    import json, os
    cap_file = os.environ.get("CAPTURAS")
    if cap_file and os.path.exists(cap_file):
        CAPT.update(json.load(open(cap_file, encoding="utf-8")))
    assign_rows()
    write_engine()
    build_ra()
    build_val()
    build_es()
    build_resumen()
    build_guia()
    for code, _ in SHEET_ORDER:
        WS[code].sheet_properties.tabColor = {"GUIA": "7F7F7F", "RES": "1F3864", "SUP": "FFC000", "DR": "FFC000", "BP": "C65911",
                                               "VAL": "548235", "ES": "548235"}.get(code, "2F5597")
        WS[code].sheet_view.zoomScale = 90
    from openpyxl.worksheet.properties import PageSetupProperties
    for code in ("GUIA", "RES", "SUP", "VAL", "ES", "DR", "BP", "RA"):
        ws = WS[code]
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
        ws.print_options.gridLines = False
    WS["RES"].print_area = "A1:F120"
    wb.active = 1
    wb.calculation.fullCalcOnLoad = True
    wb.save(OUT)
    print("Guardado:", OUT)


if __name__ == "__main__":
    main()
