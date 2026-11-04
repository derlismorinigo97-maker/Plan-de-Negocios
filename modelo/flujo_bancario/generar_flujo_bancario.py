"""Genera Flujo_Caja_Bancario_ECOSTAR.xlsx: flujo de caja mensual de 24 meses
(nov-2026 a oct-2028) para la solicitud del primer crédito de Ecostar.

Todo el cálculo queda en fórmulas de Excel; los datos de entrada están en la hoja
«Supuestos» (celdas azules sobre fondo amarillo). Fuentes: plan de negocio
CIE–PERG oct-2025 (modelo v7) e «Informe de inversión de CIE S.A. en Ecostar RV2»
(cierre 31/08/2026, proyección a 31/10/2026).

Uso:  python3 generar_flujo_bancario.py [ruta_salida.xlsx]
"""
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName

OUT = sys.argv[1] if len(sys.argv) > 1 else "Flujo_Caja_Bancario_ECOSTAR.xlsx"

# ---------------------------------------------------------------- estilos
F_IN = Font(color="1F4E9E")
FILL_IN = PatternFill("solid", fgColor="FFF7D6")
F_B = Font(bold=True)
F_T = Font(bold=True, size=14, color="1F3864")
F_H = Font(bold=True, color="FFFFFF")
FILL_H = PatternFill("solid", fgColor="1F3864")
FILL_S = PatternFill("solid", fgColor="D9E2F3")
FILL_K = PatternFill("solid", fgColor="E2EFDA")
THIN = Side(style="thin", color="BFBFBF")
B_TOP = Border(top=Side(style="thin", color="000000"))
NF_USD = '#,##0;[Red](#,##0);"-"'
NF_GS = '#,##0'
NF_PCT = '0.0%'
NF_X = '0.00"x"'

wb = Workbook()


def name(n, ref):
    wb.defined_names[n] = DefinedName(n, attr_text=ref)


def hdr(ws, r, vals, c0=1):
    for i, v in enumerate(vals):
        c = ws.cell(r, c0 + i, v)
        c.font, c.fill = F_H, FILL_H
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def sec(ws, r, text, ncol=8):
    for c in range(1, ncol + 1):
        ws.cell(r, c).fill = FILL_S
    ws.cell(r, 1, text).font = F_B


def inp(cell, v, nf=None):
    cell.value = v
    cell.font, cell.fill = F_IN, FILL_IN
    if nf:
        cell.number_format = nf


# ================================================================ SUPUESTOS
S = wb.active
S.title = "Supuestos"
S["A1"] = "ECOSTAR — Supuestos del flujo de caja bancario (24 meses)"
S["A1"].font = F_T
S["A2"] = ("Celdas azules sobre fondo amarillo = datos de entrada. Montos en USD salvo indicación. "
           "Fuentes: Plan CIE–PERG oct-2025 (modelo v7) e Informe de inversión CIE en Ecostar RV2.")
for col, w in zip("ABCDE", (6, 58, 16, 16, 70)):
    S.column_dimensions[col].width = w

r = 4
sec(S, r, "A. Parámetros generales", 5)
r += 1
hdr(S, r, ["", "Parámetro", "Valor", "Unidad", "Fuente / nota"])
GEN = [
    ("F_Ini", "Fecha del mes 1 del flujo", date(2026, 11, 1), "fecha", "Primer mes de operación autónoma de Ecostar.", "dd/mm/yyyy"),
    ("TC_A1", "Tipo de cambio — año 1 (nov-26 a oct-27)", 5890, "Gs/USD", "TC BCP estimado oct-26 (Informe RV2). El plan usaba 7.300.", NF_GS),
    ("TC_A2", "Tipo de cambio — año 2 (nov-27 a oct-28)", 5890, "Gs/USD", "Sin devaluación supuesta: criterio prudente para costos en Gs.", NF_GS),
    ("DiasCobro", "Días de cobro a clientes", 120, "días", "Plan de negocio. Se redondea a meses (120 = 4 meses).", NF_GS),
    ("DiasInvMP", "Inventario objetivo de materia prima", 120, "días de consumo", "Plan de negocio. Se compra solo lo que falta para el objetivo.", NF_GS),
    ("DiasPagoMP", "Días de pago a proveedores de materia prima", 0, "días", "Plan: pago contado/anticipado (importación).", NF_GS),
    ("CajaMin", "Caja mínima operativa", 50000, "USD", "Supuesto propio (≈ 2 semanas de egresos fijos). Ajustar a la política del banco.", NF_USD),
    ("Cargas", "Cargas sociales sobre salarios (aguinaldo, IPS, vacaciones)", 0.4309, "%", "Plan de negocio.", NF_PCT),
    ("Alim", "Alimentación del personal", 576000, "Gs/persona/mes", "Plan de negocio.", NF_GS),
    ("PisoDir", "Piso de dotación directa durante el ramp-up", 0.5, "% dotación", "La mano de obra directa no baja de este % aunque el volumen sea menor.", NF_PCT),
    ("Desp", "Despacho y gastos de importación (s/ materia prima + insumos comprados)", 0.03, "%", "Plan de negocio.", NF_PCT),
    ("Pkg", "Packaging (s/ materiales, otros variables y personal directo)", 0.01, "%", "Plan de negocio.", NF_PCT),
    ("ImpP", "Imprevistos de producción (s/ costos de producción)", 0.05, "%", "Plan de negocio.", NF_PCT),
    ("ImpA", "Imprevistos de administración (s/ personal adm. y fijos marcados)", 0.05, "%", "Plan de negocio.", NF_PCT),
    ("SegCau", "Seguro de caución y otros cargos (s/ ventas)", 0.001, "%", "Plan de negocio.", NF_PCT),
    ("IVAInt", "IVA sobre intereses", 0.10, "%", "Se trata como costo financiero (criterio del plan).", NF_PCT),
]
for n, lab, v, u, note, nf in GEN:
    r += 1
    S.cell(r, 2, lab)
    inp(S.cell(r, 3), v, nf)
    S.cell(r, 4, u)
    S.cell(r, 5, note)
    name(n, f"Supuestos!$C${r}")

r += 2
sec(S, r, "B. Régimen tributario (exportación convencional hasta aprobar la maquila)", 5)
r += 1
hdr(S, r, ["", "Parámetro", "Valor", "Unidad", "Fuente / nota"])
TRIB = [
    ("MesMaq", "Mes del flujo desde el que rige el régimen de maquila", 3, "n° de mes",
     "Resolución bicameral al 98% (cierre est. 30/11/2026). 3 = enero 2027. Antes: régimen general.", NF_GS),
    ("TribMaq", "Tributo único de maquila (s/ ventas) — desde MesMaq", 0.01, "%", "Ley 7547/2025: confirmar base (valor agregado vs. ventas) con Impuestos.", NF_PCT),
    ("IVAImp", "IVA de importación s/ materia prima + insumos — antes de MesMaq", 0.10, "%",
     "Régimen general: se paga en el despacho y se recupera por devolución al exportador.", NF_PCT),
    ("Arancel", "Arancel y tributos aduaneros no recuperables s/ MP — antes de MesMaq", 0.0, "%",
     "0% supuesto (origen Mercosur). CONFIRMAR con el despachante por NCM y origen.", NF_PCT),
    ("LagIVA", "Plazo de devolución del IVA al exportador", 6, "meses", "Plan de negocio: 6 meses. Depende de la DNIT.", NF_GS),
    ("IRE", "IRE (régimen general) — informativo", 0.10, "%",
     "Ver control en Flujo_24m: si el resultado de los meses en régimen general es positivo, agregar el IRE.", NF_PCT),
]
for n, lab, v, u, note, nf in TRIB:
    r += 1
    S.cell(r, 2, lab)
    inp(S.cell(r, 3), v, nf)
    S.cell(r, 4, u)
    S.cell(r, 5, note)
    name(n, f"Supuestos!$C${r}")

# ---- productos
r += 2
sec(S, r, "C. Productos (plan de negocio, año 1 y año 2 de operación)", 11)
r += 1
hdr(S, r, ["#", "Producto", "Precio año 1 (USD/u)", "Var. precio año 2", "Materia prima (USD/u)",
           "Var. MP año 2", "Insumos (USD/u)", "Otros variables (USD/u)", "Volumen año 1 (u)",
           "Volumen año 2 (u)", "Nota"])
PROD = [
    ("Gabinetes metálicos 153 kg", 270, 0.01, 82.65, 0.01, 45, 27.909333, 12000, 15000),
    ("Cajas metálicas 13 kg", 23, 0.01, 7.41, 0.01, 4.851, 2.73, 70000, 90000),
]
P0 = r + 1
for i, p in enumerate(PROD):
    r += 1
    S.cell(r, 1, i + 1)
    inp(S.cell(r, 2), p[0])
    for j, v in enumerate(p[1:]):
        nf = NF_PCT if j in (1, 3) else ('#,##0.00' if j in (0, 2, 4, 5) else NF_GS)
        inp(S.cell(r, 3 + j), v, nf)
    S.cell(r, 11, "Plan de negocio (Productos, modelo v7).")
P1 = r
for c in "FGHIJK":
    S.column_dimensions[c].width = 14
S.column_dimensions["K"].width = 40

# ---- ramp-up
r += 2
sec(S, r, "D. Ramp-up: % del volumen mensual del plan que se vende cada mes", 26)
r += 1
S.cell(r, 2, "Mes del flujo").font = F_B
RAMP_H = r
r += 1
S.cell(r, 2, "% del volumen del plan").font = F_B
RAMP_R = r
ramp = [0.25, 0.35, 0.45, 0.55, 0.65, 0.75, 0.85, 0.95, 1, 1, 1, 1] + [1] * 12
for t in range(24):
    S.cell(RAMP_H, 3 + t, t + 1).font = F_B
    inp(S.cell(RAMP_R, 3 + t), ramp[t], '0%')
r += 1
S.cell(r, 2, ("Supuesto propio a validar con pedidos firmes de CIE do Brasil. Real a sep-26: 715 cajas exportadas "
              "en 5 meses de producción, muy por debajo del plan. Es el supuesto más sensible para el banco."))
S.cell(r, 2).font = Font(italic=True, color="C00000")
name("Ramp", f"Supuestos!$C${RAMP_R}:$Z${RAMP_R}")

# ---- personal
r += 2
sec(S, r, "E. Personal (plan de negocio)", 7)
r += 1
hdr(S, r, ["#", "Puesto", "Clase", "Salario mensual (Gs)", "Dotación año 1", "Dotación año 2", "Nota"])
PERS = [
    ("Operador láser", "Directo", 4668900, 4, 4), ("Doblador", "Directo", 3650655, 8, 8),
    ("Soldador", "Directo", 3968900, 5, 5), ("Baño", "Directo", 3650655, 2, 2),
    ("Pintor", "Directo", 3650655, 4, 4), ("Montador", "Directo", 3150655, 10, 10),
    ("Ayudante", "Directo", 2798309, 10, 10), ("Gerente General", "Administración", 18000000, 1, 1),
    ("Gerente Industrial", "Indirecto", 9800000, 1, 1), ("Coord. Prod. Corte y Deformación", "Indirecto", 6550000, 1, 1),
    ("Coord. Prod. Soldadura", "Indirecto", 6550000, 1, 1), ("Coord. Prod. Trat. Sup. y Pintura", "Indirecto", 6550000, 1, 1),
    ("Coord. Prod. Montaje", "Indirecto", 6550000, 1, 1), ("Ingeniería", "Indirecto", 6550000, 3, 3),
    ("Mantenimiento", "Indirecto", 3650655, 2, 2), ("PCP", "Indirecto", 9800000, 1, 1),
    ("Analista PCP", "Indirecto", 2798309, 1, 1), ("Depósito", "Indirecto", 3150655, 1, 1),
    ("Compras", "Administración", 6550000, 1, 1), ("Logística/Comex", "Administración", 6550000, 1, 1),
    ("Coord. Control Calidad", "Indirecto", 6550000, 1, 1), ("Asistente Control Calidad", "Indirecto", 2798309, 1, 1),
    ("Coord. Procesos", "Indirecto", 6550000, 1, 1), ("Asistente Procesos", "Indirecto", 2798309, 1, 1),
    ("Administrador", "Administración", 9800000, 1, 1), ("Asistente Administrativo", "Administración", 2798309, 1, 1),
]
E0 = r + 1
for i, p in enumerate(PERS):
    r += 1
    S.cell(r, 1, i + 1)
    inp(S.cell(r, 2), p[0])
    inp(S.cell(r, 3), p[1])
    inp(S.cell(r, 4), p[2], NF_GS)
    inp(S.cell(r, 5), p[3], NF_GS)
    inp(S.cell(r, 6), p[4], NF_GS)
for _ in range(4):  # filas libres
    r += 1
    S.cell(r, 1, r - E0 + 1)
    for c in range(2, 7):
        inp(S.cell(r, c), None)
E1 = r
r += 1
S.cell(r, 2, "Total dotación").font = F_B
S.cell(r, 5, f"=SUM(E{E0}:E{E1})").font = F_B
S.cell(r, 6, f"=SUM(F{E0}:F{E1})").font = F_B
r += 1
S.cell(r, 2, "Ajuste salarial año 2")
inp(S.cell(r, 3), 0.0, NF_PCT)
S.cell(r, 5, "Plan: 0%. Salarios sin cargas sociales; clase Directo / Indirecto / Administración.")
name("AjSal", f"Supuestos!$C${r}")
name("Pe_Clase", f"Supuestos!$C${E0}:$C${E1}")
name("Pe_Sal", f"Supuestos!$D${E0}:$D${E1}")
name("Pe_D1", f"Supuestos!$E${E0}:$E${E1}")
name("Pe_D2", f"Supuestos!$F${E0}:$F${E1}")

# ---- costos fijos
r += 2
sec(S, r, "F. Costos fijos mensuales (plan de negocio)", 9)
r += 1
hdr(S, r, ["#", "Concepto", "Clase", "Moneda", "Monto mensual", "Crecimiento anual",
           "IVA incluido en el monto (informativo)", "Base imprevistos adm. (1/0)", "Nota"])
FIJ = [
    ("Energía eléctrica planta", "Producción", "PYG", 14000000, 0, 0, 0, "Plan. Real jun-oct: ≈ Gs 7,5 M/mes."),
    ("Mantenimiento y reparaciones planta", "Producción", "USD", 2475.67, 0.02, 0, 0, "Plan."),
    ("Alquiler de nave (Zona Franca Global)", "Administración", "USD", 17006, 0.03, 0.10, 0, "3.865 m² × USD 4 + IVA 10% (no recuperable). Real ≈ USD 16.000/mes."),
    ("Expensas", "Administración", "USD", 773, 0.03, 0, 0, "Plan."),
    ("Agua, comunicaciones y electricidad adm.", "Administración", "PYG", 750000, 0, 0, 1, "Plan."),
    ("Mantenimiento y reparaciones adm.", "Administración", "PYG", 3000000, 0, 0, 1, "Plan."),
    ("Seguros sobre activos fijos", "Administración", "PYG", 5000000, 0, 0, 1, "Plan."),
    ("Gastos generales", "Administración", "USD", 1838.33, 0, 0, 1, "Plan."),
    ("Honorarios y otros fijos (incl. outsourcing contable)", "Administración", "PYG", 15166667, 0, 0, 1, "Plan."),
]
F0 = r + 1
for i, f in enumerate(FIJ):
    r += 1
    S.cell(r, 1, i + 1)
    inp(S.cell(r, 2), f[0])
    inp(S.cell(r, 3), f[1])
    inp(S.cell(r, 4), f[2])
    inp(S.cell(r, 5), f[3], '#,##0.00')
    inp(S.cell(r, 6), f[4], NF_PCT)
    inp(S.cell(r, 7), f[5], NF_PCT)
    inp(S.cell(r, 8), f[6], NF_GS)
    S.cell(r, 9, f[7])
for _ in range(3):
    r += 1
    S.cell(r, 1, r - F0 + 1)
    for c in range(2, 9):
        inp(S.cell(r, c), None)
F1 = r
for n, c in (("Fi_Clase", "C"), ("Fi_Mon", "D"), ("Fi_Monto", "E"), ("Fi_Crec", "F"), ("Fi_IVA", "G"), ("Fi_Base", "H")):
    name(n, f"Supuestos!${c}${F0}:${c}${F1}")
S.column_dimensions["I"].width = 40

# ---- apertura y financiamiento
r += 2
sec(S, r, "G. Posición de apertura al 31/10/2026 y recupero a CIE", 5)
r += 1
hdr(S, r, ["", "Concepto", "Valor", "Unidad", "Fuente / nota"])
AP = [
    ("InvMP_Gs", "Inventario de materia prima recibido de CIE (stock + en tránsito, sin IVA)", 6690089008, "Gs",
     "Informe RV2: stock Gs 2.071 M + en tránsito Gs 4.619 M. Se convierte con TC año 1.", NF_GS),
    ("ActFijo_Gs", "Activo fijo de Ecostar (equipos, adecuaciones, mobiliario, incl. aporte en bienes)", 4500366770, "Gs",
     "Informe RV2 (implantación, partidas capitalizables). Solo para la depreciación.", NF_GS),
    ("Vida", "Vida útil promedio del activo fijo", 10, "años", "Plan de negocio.", NF_GS),
    ("CIE_Gs", "Deuda con CIE — base sin IVA (a facturar)", 15656350711, "Gs",
     "Informe RV2: Gs 17.074 M informados menos IVA incluido Gs 1.417 M. PENDIENTE de aprobación de Energy Capital.", NF_GS),
    ("IVA_CIE", "IVA de la factura de CIE a Ecostar", 0.10, "%", "Ecostar lo paga y lo recupera como exportador (plazo LagIVA).", NF_PCT),
    ("Aporte_Gs", "Aporte de capital pendiente — Energy Capital (51%)", 668000000, "Gs", "Composición del capital social (Informe RV3, lámina 9).", NF_GS),
    ("MesAporte", "Mes de integración del aporte", 1, "n° de mes", "Confirmar con el socio.", NF_GS),
    ("Capex2", "CAPEX adicional año 2", 100000, "USD", "Plan: inversión adicional año 2 (no desagregada).", NF_USD),
    ("MesCapex2", "Mes de pago del CAPEX adicional", 13, "n° de mes", "Plan.", NF_GS),
]
for n, lab, v, u, note, nf in AP:
    r += 1
    S.cell(r, 2, lab)
    inp(S.cell(r, 3), v, nf)
    S.cell(r, 4, u)
    S.cell(r, 5, note)
    name(n, f"Supuestos!$C${r}")
r += 1
S.cell(r, 2, "Factura total de CIE (con IVA), en USD").font = F_B
S.cell(r, 3, "=CIE_Gs*(1+IVA_CIE)/TC_A1").number_format = NF_USD
S.cell(r, 4, "USD")
S.cell(r, 5, "Facturada en Gs y convertida con el TC año 1 (ver riesgo cambiario en Notas).")
name("CIE_Tot", f"Supuestos!$C${r}")

r += 2
sec(S, r, "H. Financiamiento solicitado y saldo con CIE", 5)
r += 1
hdr(S, r, ["", "Concepto", "Valor", "Unidad", "Fuente / nota"])
FIN = [
    ("Pr_Monto", "Préstamo bancario — monto", 2900000, "USD", "A definir con el banco. Destino: cancelar la deuda con CIE (que cancela su préstamo de nov-26).", NF_USD),
    ("Pr_Mes", "Préstamo bancario — mes de desembolso", 1, "n° de mes", "Noviembre 2026.", NF_GS),
    ("Pr_Tasa", "Préstamo bancario — tasa nominal anual", 0.09, "%", "Plan de negocio. Reemplazar por la oferta del banco.", NF_PCT),
    ("Pr_Gracia", "Préstamo bancario — meses de gracia (solo intereses)", 6, "meses", "Plan de negocio.", NF_GS),
    ("Pr_Cuotas", "Préstamo bancario — cuotas de capital e interés (sistema francés)", 54, "cuotas", "Plan de negocio (5 años en total).", NF_GS),
    ("CIE_Tasa", "Saldo con CIE no cubierto por el préstamo — tasa anual", 0.09, "%", "Saldo = factura total − préstamo.", NF_PCT),
    ("CIE_Gracia", "Saldo con CIE — meses de gracia", 6, "meses", "", NF_GS),
    ("CIE_Cuotas", "Saldo con CIE — cuotas", 12, "cuotas", "", NF_GS),
]
for n, lab, v, u, note, nf in FIN:
    r += 1
    S.cell(r, 2, lab)
    inp(S.cell(r, 3), v, nf)
    S.cell(r, 4, u)
    S.cell(r, 5, note)
    name(n, f"Supuestos!$C${r}")
r += 1
S.cell(r, 2, "Pago a CIE con el préstamo (mes de desembolso)").font = F_B
S.cell(r, 3, "=MIN(Pr_Monto,CIE_Tot)").number_format = NF_USD
name("CIE_Pago1", f"Supuestos!$C${r}")
r += 1
S.cell(r, 2, "Saldo con CIE que queda financiado por CIE").font = F_B
S.cell(r, 3, "=MAX(0,CIE_Tot-CIE_Pago1)").number_format = NF_USD
name("CIE_Saldo", f"Supuestos!$C${r}")
S.freeze_panes = "C4"

# ================================================================ PRÉSTAMOS
PR = wb.create_sheet("Prestamos")
PR["A1"] = "Cronogramas de deuda (sistema francés, cuota mensual)"
PR["A1"].font = F_T
PR["A2"] = "Préstamo bancario y saldo financiado por CIE. Intereses sin IVA (el IVA se suma en el flujo)."
hdr(PR, 4, ["Mes", "Fecha", "Bco: saldo inicial", "Bco: desembolso", "Bco: interés", "Bco: amortización",
            "Bco: saldo final", "", "CIE: saldo inicial", "CIE: alta", "CIE: interés", "CIE: amortización", "CIE: saldo final"])
PR["O4"], PR["P4"] = "Cuota bco.", "=IFERROR(PMT(Pr_Tasa/12,Pr_Cuotas,-Pr_Monto),0)"
PR["O5"], PR["P5"] = "Cuota CIE", "=IFERROR(PMT(CIE_Tasa/12,CIE_Cuotas,-CIE_Saldo),IF(CIE_Cuotas>0,CIE_Saldo/CIE_Cuotas,0))"
PR["P4"].number_format = PR["P5"].number_format = NF_USD
name("CuotaBco", "Prestamos!$P$4")
name("CuotaCIE", "Prestamos!$P$5")
NM = 72
for m in range(1, NM + 1):
    rr = 4 + m
    PR.cell(rr, 1, m)
    PR.cell(rr, 2, f"=EDATE(F_Ini,A{rr}-1)").number_format = "mmm-yy"
    PR.cell(rr, 3, 0 if m == 1 else f"=G{rr-1}")
    PR.cell(rr, 4, f"=IF(A{rr}=Pr_Mes,Pr_Monto,0)")
    PR.cell(rr, 5, f"=C{rr}*Pr_Tasa/12")
    PR.cell(rr, 6, f"=IF(AND(A{rr}>Pr_Mes+Pr_Gracia,A{rr}<=Pr_Mes+Pr_Gracia+Pr_Cuotas),MIN(C{rr},CuotaBco-E{rr}),0)")
    PR.cell(rr, 7, f"=C{rr}+D{rr}-F{rr}")
    PR.cell(rr, 9, 0 if m == 1 else f"=M{rr-1}")
    PR.cell(rr, 10, f"=IF(A{rr}=Pr_Mes,CIE_Saldo,0)")
    PR.cell(rr, 11, f"=I{rr}*CIE_Tasa/12")
    PR.cell(rr, 12, f"=IF(AND(A{rr}>Pr_Mes+CIE_Gracia,A{rr}<=Pr_Mes+CIE_Gracia+CIE_Cuotas),MIN(I{rr},CuotaCIE-K{rr}),0)")
    PR.cell(rr, 13, f"=I{rr}+J{rr}-L{rr}")
    for c in (3, 4, 5, 6, 7, 9, 10, 11, 12, 13):
        PR.cell(rr, c).number_format = NF_USD
for c in range(1, 17):
    PR.column_dimensions[L(c)].width = 14
PR.column_dimensions["H"].width = 3
PR.freeze_panes = "C5"

# ================================================================ FLUJO 24m
FL = wb.create_sheet("Flujo_24m")
FL["A1"] = "ECOSTAR — Flujo de caja proyectado 24 meses (USD)"
FL["A1"].font = F_T
FL["A2"] = ("Método directo. Mes 1 = noviembre 2026. Año 1 = nov-26 a oct-27; año 2 = nov-27 a oct-28. "
            "Todo es fórmula: modificar solo la hoja Supuestos.")
C0, NT = 6, 24                       # columna F = mes 1
CA1, CA2, CT = C0 + NT, C0 + NT + 1, C0 + NT + 2   # Año 1, Año 2, Total
FL.column_dimensions["A"].width = 8
FL.column_dimensions["B"].width = 52
FL.column_dimensions["C"].width = 8
FL.column_dimensions["D"].width = 2
FL.column_dimensions["E"].width = 2
for c in range(C0, CT + 1):
    FL.column_dimensions[L(c)].width = 12.5
FL.freeze_panes = FL.cell(5, C0)

rows = {}
cur = [4]


def mcol(t):
    return L(C0 + t - 1)


def rng(key):
    return f"${L(C0)}${rows[key]}:${L(C0+NT-1)}${rows[key]}"


def line(key, label, fn, nf=NF_USD, tot="sum", bold=False, code=""):
    """fn(t, c) -> fórmula del mes t (c = letra de columna)."""
    rr = cur[0]
    rows[key] = rr
    FL.cell(rr, 1, code)
    FL.cell(rr, 2, label)
    for t in range(1, NT + 1):
        c = mcol(t)
        cell = FL.cell(rr, C0 + t - 1, fn(t, c))
        cell.number_format = nf
        if bold:
            cell.font = F_B
    a, b = L(C0), L(C0 + 11)
    a2, b2 = L(C0 + 12), L(C0 + 23)
    if tot == "sum":
        f1, f2, ft = f"=SUM({a}{rr}:{b}{rr})", f"=SUM({a2}{rr}:{b2}{rr})", f"=SUM({a}{rr}:{b2}{rr})"
    elif tot == "end":
        f1, f2, ft = f"={b}{rr}", f"={b2}{rr}", f"={b2}{rr}"
    elif tot == "start":
        f1, f2, ft = f"={a}{rr}", f"={a2}{rr}", f"={a}{rr}"
    elif tot == "min":
        f1, f2, ft = f"=MIN({a}{rr}:{b}{rr})", f"=MIN({a2}{rr}:{b2}{rr})", f"=MIN({a}{rr}:{b2}{rr})"
    elif tot == "max":
        f1, f2, ft = f"=MAX({a}{rr}:{b}{rr})", f"=MAX({a2}{rr}:{b2}{rr})", f"=MAX({a}{rr}:{b2}{rr})"
    else:
        f1 = f2 = ft = None
    for c, f in ((CA1, f1), (CA2, f2), (CT, ft)):
        if f:
            cell = FL.cell(rr, c, f)
            cell.number_format = nf
            cell.font = F_B
            cell.fill = FILL_K
    if bold:
        FL.cell(rr, 2).font = F_B
        for c in range(C0, CT + 1):
            FL.cell(rr, c).border = B_TOP
    cur[0] += 1
    return rr


def title(text):
    cur[0] += 1
    rr = cur[0]
    for c in range(1, CT + 1):
        FL.cell(rr, c).fill = FILL_S
    FL.cell(rr, 2, text).font = F_B
    cur[0] += 1


def R(key, c):
    return f"{c}{rows[key]}"


def lagged(key, lag_expr, t):
    return f"IF({t}>{lag_expr},INDEX({rng(key)},1,{t}-{lag_expr}),0)"


# encabezado de meses
hdr(FL, 4, ["", "Concepto", "", "", ""] + [None] * NT + ["Año 1", "Año 2", "Total 24 m"])
for t in range(1, NT + 1):
    FL.cell(4, C0 + t - 1, f"=EDATE(F_Ini,{t}-1)").number_format = "mmm-yy"
cur[0] = 5
line("t", "Mes n°", lambda t, c: t, NF_GS, None)
line("y", "Año de operación del flujo", lambda t, c: f"=IF({R('t', c)}<=12,1,2)", NF_GS, None)
line("tc", "Tipo de cambio (Gs/USD)", lambda t, c: f"=IF({R('y', c)}=1,TC_A1,TC_A2)", NF_GS, None)
line("ramp", "Ramp-up (% del volumen del plan)", lambda t, c: f"=INDEX(Ramp,1,{R('t', c)})", '0%', None)
line("gen", "Régimen general (1) / maquila (0)", lambda t, c: f"=IF({R('t', c)}<MesMaq,1,0)", NF_GS, None)

title("1. Volumen y precios")
for i in range(len(PROD)):
    pr = P0 + i
    line(f"u{i}", f"=\"Unidades — \"&Supuestos!$B${pr}",
         lambda t, c, pr=pr: f"=IF({R('y', c)}=1,Supuestos!$I${pr},Supuestos!$J${pr})/12*{R('ramp', c)}", NF_GS)
for i in range(len(PROD)):
    pr = P0 + i
    line(f"p{i}", f"=\"Precio USD/u — \"&Supuestos!$B${pr}",
         lambda t, c, pr=pr: f"=Supuestos!$C${pr}*IF({R('y', c)}=2,1+Supuestos!$D${pr},1)", '#,##0.00', None)

title("2. Estado de resultados devengado (USD)")
line("vta", "Ventas netas (exportación)",
     lambda t, c: "=" + "+".join(f"{R(f'u{i}', c)}*{R(f'p{i}', c)}" for i in range(len(PROD))), bold=True, code="VTA")
line("mp", "Materia prima consumida",
     lambda t, c: "=" + "+".join(f"{R(f'u{i}', c)}*Supuestos!$E${P0+i}*IF({R('y', c)}=2,1+Supuestos!$F${P0+i},1)"
                                 for i in range(len(PROD))))
line("ins", "Insumos",
     lambda t, c: "=" + "+".join(f"{R(f'u{i}', c)}*Supuestos!$G${P0+i}" for i in range(len(PROD))))
line("ov", "Otros costos variables",
     lambda t, c: "=" + "+".join(f"{R(f'u{i}', c)}*Supuestos!$H${P0+i}" for i in range(len(PROD))))
line("desp", "Despacho y gastos de importación", lambda t, c: f"=Desp*({R('mp', c)}+{R('ins', c)})")
line("aran", "Aranceles no recuperables (régimen general)", lambda t, c: f"=Arancel*{R('gen', c)}*{R('mp', c)}")


def pers(clase, c):
    d = (f"IF({R('y', c)}=1,SUMPRODUCT((Pe_Clase=\"{clase}\")*Pe_Sal*Pe_D1),"
         f"SUMPRODUCT((Pe_Clase=\"{clase}\")*Pe_Sal*Pe_D2)*(1+AjSal))")
    hc = (f"IF({R('y', c)}=1,SUMPRODUCT((Pe_Clase=\"{clase}\")*Pe_D1),SUMPRODUCT((Pe_Clase=\"{clase}\")*Pe_D2))")
    f = "MAX(" + R('ramp', c) + ",PisoDir)" if clase == "Directo" else "1"
    return f"({d}*(1+Cargas)+{hc}*Alim)*{f}/{R('tc', c)}"


line("pdir", "Personal directo (con cargas y alimentación)", lambda t, c: "=" + pers("Directo", c))
line("pind", "Personal indirecto de producción", lambda t, c: "=" + pers("Indirecto", c))


def fij(clase, c, base=False):
    b = "*Fi_Base" if base else ""
    return (f"SUMPRODUCT((Fi_Clase=\"{clase}\")*Fi_Monto*(1+Fi_Crec)^({R('y', c)}-1){b}*"
            f"((Fi_Mon=\"PYG\")/{R('tc', c)}+(Fi_Mon=\"USD\")))")


line("fp", "Costos fijos de producción", lambda t, c: "=" + fij("Producción", c))
line("pkg", "Packaging", lambda t, c: f"=Pkg*({R('mp', c)}+{R('ins', c)}+{R('ov', c)}+{R('pdir', c)})")
line("impp", "Imprevistos de producción",
     lambda t, c: f"=ImpP*({R('mp', c)}+{R('ins', c)}+{R('ov', c)}+{R('desp', c)}+{R('aran', c)}+{R('pdir', c)}+{R('pind', c)}+{R('fp', c)}+{R('pkg', c)})")
line("cprod", "Costo de producción y ventas",
     lambda t, c: f"=SUM({c}{rows['mp']}:{c}{rows['impp']})", bold=True)
line("mb", "Margen bruto", lambda t, c: f"={R('vta', c)}-{R('cprod', c)}", bold=True)
line("padm", "Personal de administración", lambda t, c: "=" + pers("Administración", c))
line("fa", "Costos fijos de administración", lambda t, c: "=" + fij("Administración", c))
line("impa", "Imprevistos de administración",
     lambda t, c: f"=ImpA*({R('padm', c)}+" + fij("Administración", c, True) + ")")
line("trib", "Tributo de maquila (desde la aprobación)", lambda t, c: f"=TribMaq*(1-{R('gen', c)})*{R('vta', c)}")
line("seg", "Seguro de caución y otros cargos", lambda t, c: f"=SegCau*{R('vta', c)}")
line("ebitda", "EBITDA", lambda t, c: f"={R('mb', c)}-SUM({c}{rows['padm']}:{c}{rows['seg']})", bold=True, code="EBITDA")
line("mgn", "Margen EBITDA", lambda t, c: f"=IF({R('vta', c)}=0,0,{R('ebitda', c)}/{R('vta', c)})", NF_PCT, None)
line("dep", "Depreciación",
     lambda t, c: f"=(ActFijo_Gs/TC_A1+IF({R('t', c)}>=MesCapex2,Capex2,0))/(Vida*12)")
line("intb", "Intereses préstamo bancario (con IVA)",
     lambda t, c: f"=INDEX(Prestamos!$E$5:$E${4+NM},{R('t', c)})*(1+IVAInt)")
line("intc", "Intereses saldo con CIE (con IVA)",
     lambda t, c: f"=INDEX(Prestamos!$K$5:$K${4+NM},{R('t', c)})*(1+IVAInt)")
line("res", "Resultado antes de impuestos",
     lambda t, c: f"={R('ebitda', c)}-{R('dep', c)}-{R('intb', c)}-{R('intc', c)}", bold=True)
line("resgen", "Control IRE: resultado de meses en régimen general",
     lambda t, c: f"={R('res', c)}*{R('gen', c)}")

title("3. Capital de trabajo: inventario de materia prima (USD)")
line("inv0", "Inventario inicial",
     lambda t, c: "=InvMP_Gs/TC_A1" if t == 1 else f"={mcol(t-1)}{rows['inv0']+4}", tot="start")
line("cons", "Consumo de materia prima", lambda t, c: f"={R('mp', c)}")
line("invobj", "Inventario objetivo", lambda t, c: f"={R('cons', c)}*DiasInvMP/30", tot="end")
line("comp", "Compras de materia prima",
     lambda t, c: f"=MAX(0,{R('cons', c)}+{R('invobj', c)}-{R('inv0', c)})")
line("inv1", "Inventario final", lambda t, c: f"={R('inv0', c)}+{R('comp', c)}-{R('cons', c)}", tot="end", bold=True)
line("mesesinv", "Meses de consumo en inventario",
     lambda t, c: f"=IF({R('cons', c)}=0,0,{R('inv1', c)}/{R('cons', c)})", '0.0', None)

title("4. Flujo de caja (método directo, USD)")
line("cob", "Cobranzas de exportación",
     lambda t, c: "=" + lagged("vta", "ROUND(DiasCobro/30,0)", R('t', c)), code="COB")
line("pmp", "(−) Pago de materia prima",
     lambda t, c: "=-" + lagged("comp", "ROUND(DiasPagoMP/30,0)", R('t', c)))
line("pvar", "(−) Insumos, otros variables, despacho, aranceles, packaging e imprevistos",
     lambda t, c: f"=-({R('ins', c)}+{R('ov', c)}+{R('desp', c)}+{R('aran', c)}+{R('pkg', c)}+{R('impp', c)})")
line("ppers", "(−) Personal (directo, indirecto y administración)",
     lambda t, c: f"=-({R('pdir', c)}+{R('pind', c)}+{R('padm', c)})")
line("pfij", "(−) Costos fijos e imprevistos de administración",
     lambda t, c: f"=-({R('fp', c)}+{R('fa', c)}+{R('impa', c)})")
line("ptrib", "(−) Tributo de maquila y seguro de caución", lambda t, c: f"=-({R('trib', c)}+{R('seg', c)})")
line("ivap", "(−) IVA de importación pagado (régimen general)",
     lambda t, c: f"=-IVAImp*{R('gen', c)}*({R('comp', c)}+{R('ins', c)})")
line("ivar", "(+) Devolución de IVA al exportador",
     lambda t, c: "=" + lagged("ivap", "LagIVA", R('t', c)).replace("INDEX(", "-INDEX(", 1)
     + f"+IF({R('t', c)}=Pr_Mes+LagIVA,CIE_Gs*IVA_CIE/TC_A1,0)")
line("fco", "FLUJO DE CAJA OPERATIVO", lambda t, c: f"=SUM({c}{rows['cob']}:{c}{rows['ivar']})", bold=True, code="FCO")
line("capex", "(−) CAPEX", lambda t, c: f"=-IF({R('t', c)}=MesCapex2,Capex2,0)")
line("fcl", "FLUJO DESPUÉS DE INVERSIONES (antes de financiamiento)",
     lambda t, c: f"={R('fco', c)}+{R('capex', c)}", bold=True)
line("apo", "(+) Aporte de capital", lambda t, c: f"=IF({R('t', c)}=MesAporte,Aporte_Gs/TC_A1,0)")
line("des", "(+) Desembolso del préstamo bancario",
     lambda t, c: f"=INDEX(Prestamos!$D$5:$D${4+NM},{R('t', c)})")
line("pcie", "(−) Pago a CIE con el préstamo (factura con IVA)",
     lambda t, c: f"=-IF({R('t', c)}=Pr_Mes,CIE_Pago1,0)")
line("intbp", "(−) Intereses préstamo bancario (con IVA)", lambda t, c: f"=-{R('intb', c)}")
line("amb", "(−) Amortización préstamo bancario",
     lambda t, c: f"=-INDEX(Prestamos!$F$5:$F${4+NM},{R('t', c)})")
line("intcp", "(−) Intereses saldo con CIE (con IVA)", lambda t, c: f"=-{R('intc', c)}")
line("amc", "(−) Amortización saldo con CIE",
     lambda t, c: f"=-INDEX(Prestamos!$L$5:$L${4+NM},{R('t', c)})")
line("ffin", "FLUJO DE FINANCIAMIENTO", lambda t, c: f"=SUM({c}{rows['apo']}:{c}{rows['amc']})", bold=True)
line("fn", "FLUJO NETO DEL MES", lambda t, c: f"={R('fcl', c)}+{R('ffin', c)}", bold=True)
line("c0", "Caja inicial", lambda t, c: 0 if t == 1 else f"={mcol(t-1)}{rows['fn']+2}", tot="start")
line("c1", "CAJA FINAL", lambda t, c: f"={R('c0', c)}+{R('fn', c)}", tot="end", bold=True, code="CAJA")
line("cmin", "Caja mínima operativa", lambda t, c: "=CajaMin", tot="end")
line("gap", "Excedente / (déficit) frente a la caja mínima",
     lambda t, c: f"={R('c1', c)}-{R('cmin', c)}", tot="min", bold=True)
line("need", "Financiamiento adicional requerido (acumulado)",
     lambda t, c: f"=MAX(0,-{R('gap', c)})", tot="max")

title("5. Deuda y cobertura")
line("debt", "Deuda financiera al cierre (banco + CIE)",
     lambda t, c: f"=INDEX(Prestamos!$G$5:$G${4+NM},{R('t', c)})+INDEX(Prestamos!$M$5:$M${4+NM},{R('t', c)})", tot="end")
line("sd", "Servicio de deuda (intereses + amortización)",
     lambda t, c: f"=-({R('intbp', c)}+{R('amb', c)}+{R('intcp', c)}+{R('amc', c)})")
line("fcopre", "Flujo operativo disponible para el servicio de deuda (FCO + CAPEX)",
     lambda t, c: f"={R('fcl', c)}")
rr = line("dscr", "DSCR (flujo disponible / servicio de deuda)",
          lambda t, c: f"=IF({R('sd', c)}=0,\"-\",{R('fcopre', c)}/{R('sd', c)})", NF_X, None)
for col in (CA1, CA2, CT):
    cl = L(col)
    FL.cell(rr, col, f"=IF({cl}{rows['sd']}=0,\"-\",{cl}{rows['fcopre']}/{cl}{rows['sd']})").number_format = NF_X
    FL.cell(rr, col).font = F_B
    FL.cell(rr, col).fill = FILL_K
line("deb", "Deuda / EBITDA anual",
     lambda t, c: "", NF_X, None)
for col in (CA1, CA2):
    cl = L(col)
    FL.cell(rows["deb"], col, f"=IF({cl}{rows['ebitda']}<=0,\"n/a\",{cl}{rows['debt']}/{cl}{rows['ebitda']})").number_format = NF_X
    FL.cell(rows["deb"], col).font = F_B
    FL.cell(rows["deb"], col).fill = FILL_K

title("6. Controles")
line("chk_inv", "Inventario negativo (debe ser 0)", lambda t, c: f"=IF({R('inv1', c)}<-0.5,1,0)", NF_GS)
line("chk_caja", "Caja final = caja inicial + flujo neto (diferencia)",
     lambda t, c: f"=ROUND({R('c1', c)}-{R('c0', c)}-{R('fn', c)},2)", NF_USD)
FL.cell(cur[0] + 1, 2, "CxC al cierre del mes 24 (ventas no cobradas)").font = F_B
FL.cell(cur[0] + 1, CT, f"={L(CT)}{rows['vta']}-{L(CT)}{rows['cob']}").number_format = NF_USD
rows["cxc"] = cur[0] + 1
FL.cell(cur[0] + 2, 2, "Saldo bancario al final del plazo (debe ser 0)").font = F_B
FL.cell(cur[0] + 2, CT, f"=ROUND(INDEX(Prestamos!$G$5:$G${4+NM},MIN({NM},Pr_Mes+Pr_Gracia+Pr_Cuotas)),0)").number_format = NF_USD

# ================================================================ RESUMEN
RS = wb.create_sheet("Resumen_Banco", 0)
RS["A1"] = "ECOSTAR S.A. — Flujo de caja proyectado para solicitud de crédito"
RS["A1"].font = F_T
RS["A2"] = "=\"Período: \"&TEXT(F_Ini,\"mmm-yyyy\")&\" a \"&TEXT(EDATE(F_Ini,23),\"mmm-yyyy\")&\" · USD · TC año 1: \"&TEXT(TC_A1,\"#,##0\")&\" Gs/USD\""
for col, w in zip("ABCDE", (52, 16, 16, 16, 60)):
    RS.column_dimensions[col].width = w
hdr(RS, 4, ["Indicador", "Año 1 (12 m)", "Año 2", "Total 24 m", "Lectura"])


def fl(key, col):
    return f"=Flujo_24m!{L(col)}{rows[key]}"


KP = [
    ("Ventas netas", "vta", NF_USD, "Exportación; supone el ramp-up de Supuestos D."),
    ("EBITDA", "ebitda", NF_USD, ""),
    ("Margen EBITDA", None, NF_PCT, ""),
    ("Flujo de caja operativo", "fco", NF_USD, "Incluye la inversión en capital de trabajo (cobro a 120 días)."),
    ("CAPEX", "capex", NF_USD, ""),
    ("Servicio de deuda (banco + CIE)", "sd", NF_USD, ""),
    ("DSCR (flujo disponible / servicio de deuda)", "dscr", NF_X, "Bancos suelen pedir ≥ 1,2x a partir del primer año completo."),
    ("Caja final", "c1", NF_USD, ""),
    ("Caja mínima del período (peor mes)", None, NF_USD, ""),
    ("Financiamiento adicional requerido (máximo)", "need", NF_USD, "Línea de capital de trabajo necesaria para no bajar de la caja mínima."),
    ("Deuda financiera al cierre", "debt", NF_USD, ""),
    ("Deuda / EBITDA", "deb", NF_X, ""),
]
rr = 4
for lab, key, nf, note in KP:
    rr += 1
    if lab.startswith("Caja mínima del período"):
        rmin = rr
    RS.cell(rr, 1, lab)
    for j, col in enumerate((CA1, CA2, CT)):
        if lab == "Margen EBITDA":
            f = f"=IF(Flujo_24m!{L(col)}{rows['vta']}=0,0,Flujo_24m!{L(col)}{rows['ebitda']}/Flujo_24m!{L(col)}{rows['vta']})"
        elif lab.startswith("Caja mínima del período"):
            a = (L(C0), L(C0 + 11)) if j == 0 else ((L(C0 + 12), L(C0 + 23)) if j == 1 else (L(C0), L(C0 + 23)))
            f = f"=MIN(Flujo_24m!{a[0]}{rows['c1']}:{a[1]}{rows['c1']})"
        elif key == "deb" and col == CT:
            continue
        else:
            f = fl(key, col)
        RS.cell(rr, 2 + j, f).number_format = nf
    RS.cell(rr, 5, note)
rr += 1
RS.cell(rr, 1, "Mes de menor caja")
RS.cell(rr, 4, f"=INDEX(Flujo_24m!{L(C0)}4:{L(C0+23)}4,MATCH(D{rmin},Flujo_24m!{L(C0)}{rows['c1']}:{L(C0+23)}{rows['c1']},0))").number_format = "mmm-yy"

rr += 2
sec(RS, rr, "Estructura del financiamiento solicitado", 5)
for lab, f, nf in (
    ("Préstamo bancario solicitado", "=Pr_Monto", NF_USD),
    ("Tasa nominal anual", "=Pr_Tasa", NF_PCT),
    ("Gracia / cuotas", "=Pr_Gracia&\" meses / \"&Pr_Cuotas&\" cuotas\"", None),
    ("Cuota mensual (capital + interés, sin IVA)", "=CuotaBco", NF_USD),
    ("Destino: cancelación de deuda con CIE (factura con IVA)", "=CIE_Pago1", NF_USD),
    ("Saldo con CIE financiado por CIE", "=CIE_Saldo", NF_USD),
    ("Aporte de capital de Energy Capital", "=Aporte_Gs/TC_A1", NF_USD),
):
    rr += 1
    RS.cell(rr, 1, lab)
    c = RS.cell(rr, 2, f)
    if nf:
        c.number_format = nf

rr += 2
sec(RS, rr, "Controles", 5)
for lab, f in (
    ("Inventario negativo en algún mes (0 = OK)", f"=SUM(Flujo_24m!{L(C0)}{rows['chk_inv']}:{L(C0+23)}{rows['chk_inv']})"),
    ("Diferencias de arrastre de caja (0 = OK)", f"=SUMPRODUCT(ABS(Flujo_24m!{L(C0)}{rows['chk_caja']}:{L(C0+23)}{rows['chk_caja']}))"),
    ("Resultado acumulado en meses de régimen general (si > 0, agregar IRE)", f"=Flujo_24m!{L(CT)}{rows['resgen']}"),
):
    rr += 1
    RS.cell(rr, 1, lab)
    RS.cell(rr, 2, f).number_format = NF_USD

# ================================================================ NOTAS
NO = wb.create_sheet("Notas")
NO.column_dimensions["A"].width = 4
NO.column_dimensions["B"].width = 130
NO["B1"] = "Notas, supuestos clave y pendientes"
NO["B1"].font = F_T
notes = [
    ("Supuestos clave", None),
    ("1", "Ecostar opera de forma autónoma desde nov-2026; los costos hasta oct-2026 los financió CIE y se trasladan como deuda (factura de CIE)."),
    ("2", "Volúmenes del plan de negocio (año 1: 12.000 gabinetes y 70.000 cajas) con un ramp-up de 25% a 100% en 9 meses. Es el supuesto más sensible: respaldarlo con pedidos u órdenes de compra del cliente."),
    ("3", "Exportación bajo régimen general hasta la aprobación de la maquila (mes indicado en Supuestos B): IVA de importación pagado y devuelto con desfase; sin tributo de maquila. Desde la aprobación: tributo del 1% y admisión temporaria sin IVA."),
    ("4", "El inventario de materia prima recibido de CIE (≈ USD 1,1 M, stock + tránsito) se consume antes de volver a comprar; reduce las compras de los primeros meses."),
    ("5", "La deuda con CIE se factura en guaraníes sin el IVA que CIE ya incluyó (base Gs 15.656 M) más IVA del 10%. Ecostar recupera ese IVA como exportador."),
    ("6", "Préstamo bancario con las condiciones del plan (9%, 6 meses de gracia, 54 cuotas). Reemplazar por la oferta del banco."),
    ("Pendientes que pueden cambiar el resultado", None),
    ("a", "Aprobación de Energy Capital del monto a recuperar (incluye 7% de gastos de funcionamiento e intereses al 9%). Si baja, baja la deuda y el préstamo."),
    ("b", "Fecha efectiva de aprobación del programa de maquila y tratamiento de la mercadería en tránsito importada a nombre de CIE."),
    ("c", "Confirmar con Impuestos: IVA de la factura de CIE, plazo real de devolución de IVA al exportador, IRE en el período de régimen general."),
    ("d", "Plazo de cobro real a CIE do Brasil (el plan supone 120 días): si se acorta a 60 días, la necesidad de capital de trabajo baja de forma significativa."),
    ("e", "Dotación y salarios reales vigentes (el flujo usa la dotación del plan)."),
    ("f", "Diferencia de cambio: la deuda se fija en Gs; si el guaraní se deprecia, baja en USD; si se aprecia, sube."),
]
rr = 2
for k, v in notes:
    rr += 1
    if v is None:
        NO.cell(rr, 2, k).font = F_B
    else:
        NO.cell(rr, 1, k)
        NO.cell(rr, 2, v).alignment = Alignment(wrap_text=True)

wb.calculation.fullCalcOnLoad = True
wb.save(OUT)
print("OK", OUT, "filas flujo:", cur[0])
