# -*- coding: utf-8 -*-
"""
Pruebas de la proyección rodante de 24 meses (hoja Detalle_24m).
Uso: python3 probar_detalle.py libro.xlsx recalc.py carpeta
Corre 11 casos en LibreOffice y verifica 25 resultados esperados (volumen, precio, materiales, personal,
costos fijos, TC, ventana móvil con cierre, escenario, ventana de 12 meses, guía y Proyeccion_24m).
"""
import sys, os, json, subprocess, pickle
from concurrent.futures import ThreadPoolExecutor
from openpyxl import load_workbook
BASE, RECALC, OUT = sys.argv[1:4]
def rows_by_code(ws):
    return {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value}
def det(wb, code, meses, val):
    ws = wb["Detalle_24m"]; r = rows_by_code(ws)[code]
    for m in meses: ws.cell(r, 5 + m).value = val
REALES = {"TC": [7400]*6, "U1": [600]*6, "U2": [4000]*6, "S1": [162000]*6, "S2": [92000]*6, "MAT": [120000]*6, "OVAR": [30000]*6,
          "PERP": [300000000]*6, "PERA": [60000000]*6, "FIJP": [15000]*6, "ADM": [25000]*6, "TRIB": [2540]*6,
          "CAPEX": [600000, 0, 0, 30000, 0, 0], "DESEMB": [2791587, 0, 0, 0, 0, 0], "INT": [23026]*6, "AMORT": [0]*6, "APORTE": [767782, 0, 0, 0, 0, 0], "DIV": [0]*6}
SALDOS = {"S_CAJA": 1500000, "S_CXC": 900000, "S_INV": 500000, "S_IVA": 0, "S_PROV": 100000, "S_DEUDA": 2791587}
def reales(wb):
    ws = wb["Reales"]; rr = rows_by_code(ws)
    for k, v in REALES.items():
        for m, x in enumerate(v, 1): ws.cell(rr[k], 5 + m).value = x
    for k, v in SALDOS.items(): ws.cell(rr[k], 11).value = v
    inv = wb["Inversion"]
    for r in range(6, 17): inv.cell(r, 7).value = inv.cell(r, 4).value; inv.cell(r, 8).value = inv.cell(r, 4).value
    inv.cell(17, 7).value = 35833.91; inv.cell(17, 8).value = 35833.91
    wb["Inicio"]["C11"] = 6
M24 = range(1, 25)
CASES = {
 "base": [],
 "vol": [lambda w: det(w, "Q1", M24, 1100), lambda w: det(w, "Q2", M24, 6500)],
 "vol23": [lambda w: det(w, "Q1", range(1, 24), 1100)],
 "precio": [lambda w: det(w, "P1", range(1, 13), 280)],
 "mp": [lambda w: det(w, "MP1", M24, 90)],
 "pers": [lambda w: det(w, "DO1", M24, 6), lambda w: det(w, "SA1", M24, 5000000)],
 "fijo": [lambda w: det(w, "FC1", M24, 20000000)],
 "tc": [lambda w: det(w, "TC", M24, 8000)],
 "rolling": [reales, lambda w: det(w, "Q1", range(1, 37), 1100)],
 "cons": [lambda w: det(w, "Q1", M24, 1100), lambda w: w["Inicio"].__setitem__("C13", "Conservador")],
 "vent12": [lambda w: det(w, "Q1", M24, 1100), lambda w: w["Inicio"].__setitem__("C14", 12)],
}
def run(item):
    name, edits = item
    wb = load_workbook(BASE)
    for f in edits: f(wb)
    p = os.path.join(OUT, f"t_{name}.xlsx"); wb.save(p)
    out = subprocess.run(["python3", RECALC, p, "900"], capture_output=True, text=True).stdout
    j = json.loads(out[out.index("{"):])
    w = load_workbook(p, data_only=True); ws = w["Calculo"]
    rows = {ws.cell(r, 1).value: [ws.cell(r, c).value or 0 for c in range(6, 150)] for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value}
    g = w["Guia_Carga"]; guia = {g.cell(r, 1).value: g.cell(r, 6).value for r in range(9, g.max_row + 1) if isinstance(g.cell(r, 1).value, int)}
    py = w["Proyeccion_24m"]; proy = {py.cell(r, 1).value: [py.cell(r, c).value for c in range(3, 30)] for r in range(1, py.max_row + 1) if py.cell(r, 1).value}
    hdr = [[py.cell(r, c).value for c in range(6, 30)] for r in range(4, 9)]
    return name, dict(err=j.get("total_errors"), rows=rows, guia=guia, proy=proy, hdr=hdr)
with ThreadPoolExecutor(6) as ex:
    R = dict(ex.map(run, CASES.items()))

B = R["base"]
r = lambda n, k: R[n]["rows"][k]
RES = []
def chk(lab, c):
    RES.append((lab, bool(c)))
    print(f"  [{'OK' if c else 'FALLA'}] {lab}")
for n in R:
    chk(f"{n}: sin errores de fórmula", R[n]["err"] == 0)
    chk(f"{n}: balance = 0", max(abs(x) for x in r(n, "V.CHK") + r(n, "P.CHK")) < 0.01)
    if n != "rolling":
        chk(f"{n}: presupuesto (bloque P) sin cambio", all(r(n, k) == B["rows"][k] for k in B["rows"] if k.startswith("P.")))
chk("volumen: gabinetes meses 1-24 = 1.100", all(abs(x - 1100) < 1e-9 for x in r("vol", "V.u1")[:24]))
chk("volumen: mes 25 vuelve al anual (1.500)", abs(r("vol", "V.u1")[24] - 1500) < 1e-9)
chk("volumen: cajas meses 1-24 = 6.500", all(abs(x - 6500) < 1e-9 for x in r("vol", "V.u2")[:24]))
chk("volumen: ventas gabinetes mes 13 = 1.100 × 272,70", abs(r("vol", "V.s1")[12] - 1100 * 270 * 1.01) < 1e-6)
chk("precio: ventas gabinetes meses 1-12 = 1.000 × 280", all(abs(x - 280000) < 1e-6 for x in r("precio", "V.s1")[:12]))
chk("precio: mes 13 vuelve al anual (1.250 × 272,70)", abs(r("precio", "V.s1")[12] - 1250 * 270 * 1.01) < 1e-6)
chk("materia prima: Δ mes 1 = 1.000 × (90 − 82,65)", abs((r("mp", "V.mp_f")[0] - B["rows"]["V.mp_f"][0]) - 1000 * (90 - 82.65)) < 1e-6)
chk("materia prima: Δ mes 13 = 1.250 × (90 − 82,65 × 1,01)", abs((r("mp", "V.mp_f")[12] - B["rows"]["V.mp_f"][12]) - 1250 * (90 - 82.65 * 1.01)) < 1e-6)
chk("materia prima: mes 25 sin cambio", abs(r("mp", "V.mp_f")[24] - B["rows"]["V.mp_f"][24]) < 1e-6)
exp = (6 * 5000000 - 4 * 4668900) / 7300 * 1.4309 + 2 * 576000 / 7300
chk("personal: Δ directo mes 1 = (dotación 6 × Gs 5.000.000 − 4 × 4.668.900) / TC × 1,4309 + alimentación", abs((r("pers", "V.dir_f")[0] - B["rows"]["V.dir_f"][0]) - exp) < 1e-6)
chk("personal: personas directas mes 1 = 45", abs(r("pers", "V.dir_n")[0] - 45) < 1e-9)
chk("costo fijo: Δ fijos de producción mes 1 = Gs 6.000.000 / 7.300", abs((r("fijo", "V.fp_f")[0] - B["rows"]["V.fp_f"][0]) - 6000000 / 7300) < 1e-6)
chk("TC: meses 1-24 = 8.000 y mes 25 = 7.300", all(x == 8000 for x in r("tc", "V.tc")[:24]) and r("tc", "V.tc")[24] == 7300)
chk("TC: personal administrativo mes 1 × 7.300 / 8.000", abs(r("tc", "V.PERA")[0] / B["rows"]["V.PERA"][0] - 7300 / 8000) < 1e-9)
u = r("rolling", "V.u1")
chk("ventana móvil (corte 6): meses 1-6 = reales (600)", all(abs(x - 600) < 1e-9 for x in u[:6]))
chk("ventana móvil: meses 7-30 = detalle (1.100)", all(abs(x - 1100) < 1e-9 for x in u[6:30]))
chk("ventana móvil: meses 31-36 = anual (detalle fuera de ventana no se usa)", all(abs(x - 1500) < 1e-9 for x in u[30:36]))
u = r("cons", "V.u1")
chk("Conservador: meses 1-3 sin ventas (atraso de 3 meses)", all(x == 0 for x in u[:3]))
chk("Conservador: meses 4-24 = 1.100 × 0,90", all(abs(x - 990) < 1e-9 for x in u[3:24]))
u = r("vent12", "V.u1")
chk("ventana de 12 meses: 1-12 detalle; 13-24 anual (1.250)", all(abs(x - 1100) < 1e-9 for x in u[:12]) and all(abs(x - 1250) < 1e-9 for x in u[12:24]))
g = lambda n: R[n]["guia"].get(29)
chk("guía paso 29: base «○ Usa supuestos anuales»", str(g("base")).startswith("○"))
chk("guía paso 29: 24 meses cargados «✔»", str(g("vol")).startswith("✔"))
chk("guía paso 29: falta un mes «⚠ ... completar hasta el mes 24»", str(g("vol23")).startswith("⚠") and "24" in str(g("vol23")))
h, p = R["rolling"]["hdr"], R["rolling"]["proy"]
chk("Proyeccion_24m: arranca en el mes 7 (corte 6) y termina en el 30", h[1][0] == 7 and h[1][23] == 30)
chk("Proyeccion_24m: origen de volúmenes = Detalle en los 24 meses", all(x == "Detalle" for x in h[3]))
chk("Proyeccion_24m: ventas del primer mes = Calculo (mes 7)", abs(p["V.vtas"][3] - r("rolling", "V.vtas")[6]) < 1e-6)
chk("Proyeccion_24m: total año 1 = suma de los meses 7-18", abs(p["V.vtas"][0] - sum(r("rolling", "V.vtas")[6:18])) < 1e-6)
chk("Proyeccion_24m: caja al cierre del año 2 = caja del mes 30", abs(p["V.CAJA"][1] - r("rolling", "V.CAJA")[29]) < 1e-6)
ok = sum(1 for _, c in RES if c)
print(f"Resultado: {ok} de {len(RES)} comprobaciones OK")
pickle.dump(dict(R=R, RES=RES), open(os.path.join(OUT, "detalle.pkl"), "wb"))
