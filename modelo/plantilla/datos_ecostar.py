# -*- coding: utf-8 -*-
"""
Carga inicial de ECOSTAR (maquila CIE–PERG: gabinetes y cajas metálicas).

Presupuesto original = plan «Procesamiento de chapas para gabinetes metálicos»
(Mujica & Saldivar, Oct-2025), con la metodología del modelo común:
- un solo tipo de cambio (Gs 7.300) también para la nómina;
- 1% acumulativo de materia prima para ambos productos (texto del plan);
- montos fijos con moneda de origen inferida en la versión detallada (v1).
Vigente = hereda el presupuesto (no se cargaron datos reales posteriores al plan).
"""

ANIOS = 10

DATOS = {
    "negocio": "ECOSTAR — Maquila CIE–PERG",
    "socios": "CIE S.A. / PERG",
    "etapa": "Operación",
    "producto_desc": "Procesamiento de chapas: gabinetes metálicos (153 kg) y cajas metálicas (13 kg)",
    "fecha_mes1": None,          # pendiente: el plan no fija fechas calendario
    "corte": None,               # sin datos reales cargados
    "inicio_real": None,
    "escenario": "Vigente",
    "ventana": 24,               # meses de proyección detallada (ventana móvil)
    "version_plan": "Plan CIE–PERG Oct-2025",
    "fecha_revision": None,      # pendiente: fecha de aprobación del plan vigente
    "frecuencia": "Anual",
    # ------------------------------------------------------------------ supuestos generales
    # clave: (valor presupuesto original, nota / fuente)
    "params": {
        "Inicio": (1, "PDF: inversión en el año 0 y operación desde el año 1 (mes 1 del modelo)."),
        "Anios": (10, "PDF: 10 años de operación."),
        "Merma": (0, "PDF: producción = ventas."),
        "Cargas": (0.4309, "PDF p. 19: coeficiente del plan (no validado como tasa legal)."),
        "Alim": (576000, "Inferido del PDF: USD 946,85/persona/año a Gs 7.300."),
        "PreDir": (0.5, "Supuesto provisional para meses de implantación (no aplica al plan original)."),
        "PreInd": (1, "Supuesto provisional para meses de implantación."),
        "Desp": (0.03, "PDF p. 19: 3% sobre materia prima + insumos."),
        "Pkg": (0.01, "PDF p. 19: 1% sobre materiales, otros variables y personal directo."),
        "ImpP": (0.05, "PDF p. 19: 5% sobre costos de producción (excluye tributo)."),
        "ImpA": (0.05, "PDF p. 19: 5% sobre administración (excluye alquiler y expensas)."),
        "Trib": (0.01, "PDF p. 7: 1% sobre el valor exportado (verificar Ley 7547/2025)."),
        "SegV": (0.001, "Seguro de caución: 0,1% de ventas (inferido del PDF)."),
        "ComV": (0, "PDF: sin gastos de venta (pendiente confirmar)."),
        "IRE": (0, "PDF: IRE no aplica bajo maquila (verificar)."),
        "DSO": (120, "PDF p. 8."),
        "DInv": (120, "PDF p. 8: materia prima e insumos."),
        "DPT": (10, "PDF p. 8."),
        "DPOmp": (0, "PDF p. 24."),
        "DPOot": (30, "PDF p. 8."),
        "CajaMin": (3, "PDF cuadro 16: 3 días de ventas."),
        "IVArec": (0, "PDF: IVA considerado costo (criterio prudencial)."),
        "IVAmeses": (6, "Supuesto provisional."),
        "IVAint": (0.10, "PDF p. 4."),
        "WACC": (0.0948937, "Tasa ponderada inicial del PDF (histórica: 93% deuda a 9%). Actualizar."),
        "Ke": (0.1632, "PDF p. 9 (histórica)."),
        "RealCT": (1, "Supuesto provisional: recuperación total del capital de trabajo al final."),
        "RealAF": (1, "Supuesto provisional: valor contable = valor de realización (requiere tasación)."),
        "MesDiv": (4, "Supuesto provisional: dividendos pagados en abril del año siguiente."),
    },
    # ------------------------------------------------------------------ supuestos por año (1..10)
    "anual": {
        "TC": [7300] * 10,
        "dPrecio": [0] + [0.01] * 9,
        "dMP": [0] + [0.01] * 9,
        "dOV": [0] * 10,
        "dSal": [0] * 10,
        "Payout": [0, 0, 0.30, 0.50] + [1.0] * 6,
    },
    # ------------------------------------------------------------------ productos
    # nombre, unidad, cap inicial, cap ampliada, año ampliación, precio, MP/u, insumos/u, otros variables/u, nota
    "productos": [
        ("Gabinetes metálicos 153 kg", "u", 15000, 22000, 3, 270, 82.65, 45, 27.909333,
         "Otros variables = consumibles 4,709 (inferido) + flete 23,20 (inferido del cuadro 10a)."),
        ("Cajas metálicas 13 kg", "u", 90000, 121000, 3, 23, 7.41, 4.851, 2.73,
         "Insumos = 339.570/70.000. Otros variables = consumibles 0,65 + flete 2,08 (inferido)."),
    ],
    "volumen": [
        [12000, 15000, 18000] + [21000] * 7,
        [70000, 90000, 110000] + [120000] * 7,
    ],
    # ------------------------------------------------------------------ personal: puesto, clase, moneda, salario mensual, dotación años 1-5 (6-10 = año 5)
    "personal": [
        ("Operador láser", "Directo", "PYG", 4668900, [4, 4, 5, 5, 6]),
        ("Doblador", "Directo", "PYG", 3650655, [8, 8, 9, 9, 10]),
        ("Soldador", "Directo", "PYG", 3968900, [5, 5, 6, 6, 6]),
        ("Baño", "Directo", "PYG", 3650655, [2, 2, 2, 2, 3]),
        ("Pintor", "Directo", "PYG", 3650655, [4, 4, 4, 5, 5]),
        ("Montador", "Directo", "PYG", 3150655, [10, 10, 10, 11, 11]),
        ("Ayudante", "Directo", "PYG", 2798309, [10, 10, 10, 11, 11]),
        ("Gerente General", "Administración", "PYG", 18000000, [1] * 5),
        ("Gerente Industrial", "Indirecto", "PYG", 9800000, [1] * 5),
        ("Coord. Prod. Corte y Deformación", "Indirecto", "PYG", 6550000, [1] * 5),
        ("Coord. Prod. Soldadura", "Indirecto", "PYG", 6550000, [1] * 5),
        ("Coord. Prod. Trat. Sup. y Pintura", "Indirecto", "PYG", 6550000, [1] * 5),
        ("Coord. Prod. Montaje", "Indirecto", "PYG", 6550000, [1] * 5),
        ("Ingeniería", "Indirecto", "PYG", 6550000, [3] * 5),
        ("Mantenimiento", "Indirecto", "PYG", 3650655, [2] * 5),
        ("PCP", "Indirecto", "PYG", 9800000, [1] * 5),
        ("Analista PCP", "Indirecto", "PYG", 2798309, [1] * 5),
        ("Depósito", "Indirecto", "PYG", 3150655, [1] * 5),
        ("Compras", "Administración", "PYG", 6550000, [1] * 5),
        ("Logística/Comex", "Administración", "PYG", 6550000, [1] * 5),
        ("Coord. Control Calidad", "Indirecto", "PYG", 6550000, [1] * 5),
        ("Asistente Control Calidad", "Indirecto", "PYG", 2798309, [1] * 5),
        ("Coord. Procesos", "Indirecto", "PYG", 6550000, [1] * 5),
        ("Asistente Procesos", "Indirecto", "PYG", 2798309, [1] * 5),
        ("Administrador", "Administración", "PYG", 9800000, [1] * 5),
        ("Asistente Administrativo", "Administración", "PYG", 2798309, [1] * 5),
    ],
    # ------------------------------------------------------------------ costos fijos: concepto, clase, moneda, monto mensual, crecimiento anual, desde año op (0 = implantación), IVA incluido, base imprevistos, nota
    "fijos": [
        ("Energía eléctrica planta", "Producción", "PYG", 14000000, 0, 1, 0, 1, "Gs 168 M/año (inferido del PDF)."),
        ("Mantenimiento y reparaciones planta", "Producción", "USD", 2475.67, 0.02, 1, 0, 1, "USD 29.708/año; PDF crece 2% lineal (aquí 2% compuesto)."),
        ("Alquiler de nave (3.865 m2 x USD 4 + IVA)", "Administración", "USD", 17006, 0.03, 0, 0.10, 0, "PDF: «costo de oportunidad». Confirmar si es pago real."),
        ("Expensas (USD 0,20/m2)", "Administración", "USD", 773, 0.03, 0, 0, 0, "PDF p. 19."),
        ("Agua, comunicaciones y electricidad", "Administración", "PYG", 750000, 0, 0, 0, 1, "Gs 9 M/año (inferido)."),
        ("Mantenimiento y reparaciones adm.", "Administración", "PYG", 3000000, 0, 0, 0, 1, "Gs 36 M/año (inferido)."),
        ("Seguros sobre activos fijos", "Administración", "PYG", 5000000, 0, 0, 0, 1, "Gs 60 M/año (inferido)."),
        ("Gastos generales", "Administración", "USD", 1838.33, 0, 0, 0, 1, "USD 22.060/año."),
        ("Honorarios y otros fijos", "Administración", "PYG", 15166667, 0, 0, 0, 1, "Gs 182 M/año (inferido)."),
    ],
    # ------------------------------------------------------------------ inversión: ítem, clase, presupuesto USD, mes de pago, vida útil
    "capex": [
        ("Adecuación de infraestructura", "Infraestructura", 87000, 1, 10),
        ("Conjunto línea de pintura", "Maquinaria", 183073.60, 1, 10),
        ("Cortadoras láser (2)", "Maquinaria", 84295.23, 1, 10),
        ("Manipulador", "Maquinaria", 10197.00, 1, 10),
        ("Dobladoras (4)", "Maquinaria", 151967.08, 1, 10),
        ("Soldadoras (6)", "Maquinaria", 3399.00, 1, 10),
        ("Lijadoras (4)", "Maquinaria", 589.16, 1, 10),
        ("Conjunto baño", "Maquinaria", 28325.01, 1, 10),
        ("Línea de montaje", "Maquinaria", 9857.10, 1, 10),
        ("Montacargas", "Maquinaria", 29458.01, 1, 10),
        ("Soldadoras capacitivas (2)", "Maquinaria", 6004.90, 1, 10),
        ("Software de gestión", "Intangible", 43478, 1, 5),
        ("Mobiliario", "Mobiliario", 17400, 1, 5),
        ("Inversión adicional año 2 (no desagregada)", "Pendiente", 100000, 13, 10),
    ],
    "capex_nota": "Maquinaria = CIF + despacho asignado (3% del CIF). Vida útil del ítem del año 2: provisional.",
    # ------------------------------------------------------------------ préstamos: nombre, monto, mes desembolso, tasa nominal anual, meses de gracia, cuotas
    "prestamos": [
        ("Préstamo bancario inicial", 2791587, 1, 0.09, 6, 54),
    ],
    # ------------------------------------------------------------------ aportes: concepto, mes, monto
    "aportes": [
        ("Aporte propio inicial", 1, 200000),
        ("Aporte adicional año 1 (plan)", 1, 567782),
        ("Aporte adicional año 2 (plan)", 13, 371928),
    ],
    # ------------------------------------------------------------------ referencia del plan publicado
    "publicado": {
        "Ventas": [4850000, 6181200, 7538539, 8685437, 8772292, 8860015, 8948615, 9038101, 9128482, 9219767],
        "EBITDA": [573024, 1029011, 1479759, 1886306, 1914796, 1973999, 2033657, 2093768, 2154335, 2215355],
        "Utilidad": [301922, 731144, 1238751, 1707489, 1804006, 1902407, 1974240, 2034352, 2094918, 2155939],
        "VAN": 4588736, "TIR": 0.2261,
        "nota": "Plan Oct-2025, cuadros 15 y 19. El VAN/TIR publicados mezclan flujo del proyecto y del accionista.",
    },
    "pendientes": [
        "Fecha calendario del mes 1 y mes real de inicio de operación.",
        "Ventas, costos, CAPEX, deuda y saldos reales desde octubre de 2025 (hoja Reales).",
        "CAPEX ejecutado, pagado, comprometido y por contratar por ítem (hoja Inversion).",
        "Contrato CIE–PERG: venta de producto o servicio de maquila; propiedad de materiales; logística de exportación.",
        "Régimen de la Ley 7547/2025 y Decreto 5714/2026: base del tributo y recuperación de IVA.",
        "Naturaleza del alquiler (pago real, relacionada o costo de oportunidad).",
        "Tasas vigentes (WACC, Ke) y tipo de cambio proyectado.",
        "Nómina actual y desglose de cargas sociales.",
    ],
}


# ---------------------------------------------------------------------------------------------------------------------
# Cifras publicadas del plan (Oct-2025), para el bloque B de las hojas de cuadros. Listas de 11 valores: año 0 y años 1-10.
# Fuente: cuadros financieros del plan (páginas impresas). Las líneas combinadas suman líneas publicadas (se indica cuáles).
_A = lambda *v: [None] + list(v)            # años 1-10, sin año 0
_K = lambda v: [None] + [v] * 10            # mismo valor en los años 1-10
_S = lambda *ls: [None if all(x is None for x in t) else sum(x or 0 for x in t) for t in zip(*ls)]

_ventas = _A(4850000, 6181200, 7538539, 8685437, 8772292, 8860015, 8948615, 9038101, 9128482, 9219767)
_u1 = _A(12000, 15000, 18000, 21000, 21000, 21000, 21000, 21000, 21000, 21000)
_u2 = _A(70000, 90000, 110000, 120000, 120000, 120000, 120000, 120000, 120000, 120000)
_mp = _A(1510500, 1925717, 2340854, 2686334, 2704216, 2722278, 2740519, 2758944, 2777552, 2796347)
_ins = _A(879570, 1111590, 1343610, 1527120, 1527120, 1527120, 1527120, 1527120, 1527120, 1527120)
_cons = _A(102012, 129140, 156267, 176895, 176895, 176895, 176895, 176895, 176895, 176895)
_flete = _A(424000, 535200, 646400, 736800, 736800, 736800, 736800, 736800, 736800, 736800)
_mod = _A(223678, 223678, 242111, 256510, 274466, 274466, 274466, 274466, 274466, 274466)
_prestp = _A(96376, 96376, 104318, 110522, 118258, 118258, 118258, 118258, 118258, 118258)
_alimp = _A(40715, 40715, 43555, 46396, 49236, 49236, 49236, 49236, 49236, 49236)
_energia = _K(23014)
_mantp = _A(29708, 30302, 30897, 31491, 32085, 32679, 33273, 33867, 34462, 35056)
_desp = _A(71702, 91119, 110534, 126404, 126940, 127482, 128029, 128582, 129140, 129704)
_seg = _A(4850, 6181, 7539, 8685, 8772, 8860, 8949, 9038, 9128, 9220)
_pkg = _A(32769, 40624, 48771, 55406, 55870, 56051, 56233, 56417, 56603, 56791)
_impp = _A(171945, 212683, 254893, 289279, 291684, 292657, 293640, 294632, 295634, 296645)
_trib = _A(48500, 61812, 75385, 86854, 87723, 88600, 89486, 90381, 91285, 92198)
_prod = _A(3659338, 4528150, 5428148, 6161710, 6213080, 6234396, 6255919, 6277651, 6299594, 6321750)
_adm = _A(617639, 624039, 630632, 637422, 644416, 651620, 659039, 666682, 674554, 682662)
_moa, _presta, _alima = _K(211643), _K(91190), _K(20831)
_agua, _manta, _seguros, _gg, _hon, _impa = _K(1233), _K(4932), _K(8219), _K(22060), _K(24932), _K(19252)
_alq = _A(204072, 210194, 216500, 222995, 229685, 236575, 243673, 250983, 258512, 266268)
_exp = _A(9276, 9554, 9841, 10136, 10440, 10753, 11076, 11408, 11751, 12103)
_dep = _A(0, 71592, 71592, 71592, 71592, 71592, 59417, 59417, 59417, 59417)
_fin = _A(271102, 226274, 169416, 107224, 39198, 0, 0, 0, 0, 0)
_tot = _A(4548078, 5450056, 6299788, 6977948, 6968286, 6957607, 6974375, 7003749, 7033564, 7063828)
_ebitda = _A(573024, 1029011, 1479759, 1886306, 1914796, 1973999, 2033657, 2093768, 2154335, 2215355)
_util = _A(301922, 731144, 1238751, 1707489, 1804006, 1902407, 1974240, 2034352, 2094918, 2155939)
_amort = _A(257526, 551019, 602708, 659246, 721088, 0, 0, 0, 0, 0)
_ctvar = [2336543, 538119, 549192, 448306, 26363, 16055, 16254, 16456, 16661, 16870, 0]
_z = _K(0)

PLAN_CUADROS = {
    # cuadro 1 (p. 11), cuadros 2, 5 y 6 (p. 12-13) y cuadro 18 (p. 26): montos totales
    "c01.inf": 87000, "c01.maq": 507166, "c01.otr": 60878 + 100000, "c01.fija": 655044 + 100000, "c01.ini": 655044, "c01.post": 100000,
    "c01.cop": 2336543, "c01.tini": 2991587, "c01.ap": 200000 + 567782 + 371928, "c01.pr": 2791587,
    "c01.items": [87000, 177741, 81840, 9900, 147541, 3300, 572, 27500, 9570, 28600, 5830, 43478, 17400, 100000],
    # cuadros 7 y 7a (p. 14-15)
    "c07.cap.1": _A(15000, 15000, *[22000] * 8), "c07.p.1": _u1, "c07.u.1": _u1,
    "c07.cap.2": _A(90000, 90000, *[121000] * 8), "c07.p.2": _u2, "c07.u.2": _u2,
    # cuadro 8 (p. 16)
    "c08.vtas": _ventas, "c08.iva": _z,
    "c08.s.1": _A(3240000, 4090500, 4957686, 5841807, 5900225, 5959227, 6018819, 6079007, 6139798, 6201195), "c08.q.1": _u1,
    "c08.s.2": _A(1610000, 2090700, 2580853, 2843631, 2872067, 2900788, 2929796, 2959094, 2988685, 3018571), "c08.q.2": _u2,
    # cuadros 9a y 9b (p. 17-18); «otros variables» = consumibles (9b) + flete de materia prima (10a)
    "c09.mp": _mp, "c09.u.1": _u1, "c09.u.2": _u2,
    "c09.mp.1": _A(991800, 1252148, 1517603, 1788242, 1806124, 1824186, 1842427, 1860852, 1879460, 1898255),
    "c09.mp.2": _A(518700, 673569, 823251, *[898092] * 7),
    "c09.ins": _ins, "c09.ins.1": _A(540000, 675000, 810000, *[945000] * 7), "c09.ins.2": _A(339570, 436590, 533610, *[582120] * 7),
    "c09.ov": _S(_cons, _flete),
    # cuadro 10a (p. 19)
    "c10.prod": _prod, "c10.mp": _mp, "c10.ins": _ins, "c10.ov": _S(_cons, _flete), "c10.mod": _mod, "c10.carp": _prestp, "c10.alip": _alimp,
    "c10.fp": _S(_energia, _mantp), "c10.desp": _desp, "c10.vs": _seg, "c10.pkg": _pkg, "c10.impp": _impp, "c10.trib": _trib,
    "c10.adm": _adm, "c10.moa": _moa, "c10.cara": _presta, "c10.alia": _alima,
    "c10.fa": _S(_agua, _alq, _manta, _seguros, _gg, _hon, _exp), "c10.impa": _impa, "c10.ivar": _z, "c10.dep": _dep, "c10.fin": _fin,
    "c10.uni": _S(_u1, _u2),
    "c10.fc.1": _energia, "c10.fc.2": _mantp, "c10.fc.3": _alq, "c10.fc.4": _exp, "c10.fc.5": _agua, "c10.fc.6": _manta,
    "c10.fc.7": _seguros, "c10.fc.8": _gg, "c10.fc.9": _hon,
    # cuadro 10b (p. 20)
    "c10b.tot": _tot,
    "c10b.cd.1": _A(1854532, 2330562, 2811700, 3298022, 3315905, 3333966, 3352208, 3370632, 3389241, 3408035), "c10b.u.1": _u1,
    "c10b.s.1": _A(3240000, 4090500, 4957686, 5841807, 5900225, 5959227, 6018819, 6079007, 6139798, 6201195),
    "c10b.cd.2": _A(1035720, 1338309, 1635711, *[1784412] * 7), "c10b.u.2": _u2,
    "c10b.s.2": _A(1610000, 2090700, 2580853, 2843631, 2872067, 2900788, 2929796, 2959094, 2988685, 3018571),
    # cuadro 11 (p. 21); «mantenimiento, servicios, alquiler y otros» = reparación y mantenimiento + electricidad y servicios + otros costos
    "c11.tot": _tot, "c11.mo": _A(684432, 684432, 713647, 737092, 765624, 765624, 765624, 765624, 765624, 765624),
    "c11.mp": _A(2916082, 3701646, 4487131, 5127149, 5145032, 5163093, 5181335, 5199759, 5218368, 5237162),
    "c11.ot": _S(_A(34640, 35234, 35828, 36422, 37016, 37611, 38205, 38799, 39393, 39987), _K(24247),
                 _A(497374, 553699, 612007, 660964, 670914, 679359, 688033, 696941, 706091, 715490)),
    "c11.imp": _A(120202, 152931, 185919, 213258, 214663, 216082, 217515, 218963, 220425, 221902), "c11.fin": _fin, "c11.dep": _dep,
    # cuadros 13 y 14 (p. 22)
    "c13.des": [2791587] + [0] * 10, "c13.am": _amort,
    "c13.sal": _A(2534061, 1983042, 1380334, 721088, 0, 0, 0, 0, 0, 0),
    "c13.int": _A(246456, 205704, 154015, 97476, 35635, 0, 0, 0, 0, 0), "c13.iva": _A(24646, 20570, 15401, 9748, 3563, 0, 0, 0, 0, 0),
    "c13.fin": _fin, "c14.deb": _z, "c14.cf": _z, "c14.rec": _z, "c14.sal": _z,
    # cuadro 15 (p. 23)
    "c15.vtas": _ventas, "c15.prod": _prod, "c15.adm": _adm, "c15.ebitda": _ebitda, "c15.dep": _dep, "c15.fin": _fin, "c15.ire": _z,
    "c15.ni": _util, "c15.dec": _A(0, 0, 371625, 853745, 1804006, 1902407, 1974240, 2034352, 2094918, 2155939),
    "c15.pol": _A(0, 0, 0.30, 0.50, 1, 1, 1, 1, 1, 1),
    # cuadro 16 (p. 24)
    "c16.bruto": _A(2400374, 2944744, 3501236, 3955637, 3983755, 4000572, 4017606, 4034861, 4052341, 4070049),
    "c16.caja": _A(39863, 50804, 61961, 71387, 72101, 72822, 73550, 74286, 75029, 75779),
    "c16.sue": _A(17890, 17890, 18647, 19239, 19977, 19977, 19977, 19977, 19977, 19977),
    "c16.stk": _A(819315, 1041023, 1262706, 1443403, 1449282, 1455220, 1461217, 1467274, 1473392, 1479571),
    "c16.pt": _A(117177, 141156, 165994, 186278, 187877, 188658, 189451, 190256, 191073, 191902),
    "c16.cxc": _A(1406129, 1693870, 1991928, 2235331, 2254519, 2263895, 2273411, 2283068, 2292870, 2302820),
    "c16.iva": _z, "c16.pas": _A(63831, 70082, 77382, 83478, 85233, 85995, 86775, 87574, 88393, 89232), "c16.pmp": _z,
    "c16.pot": _A(63831, 70082, 77382, 83478, 85233, 85995, 86775, 87574, 88393, 89232),
    "c16.ctn": _A(2336543, 2874662, 3423854, 3872159, 3898522, 3914577, 3930831, 3947287, 3963948, 3980818),
    "c16.var": _ctvar,
    # cuadro 17 (p. 25)
    "c17.vtas": _ventas, "c17.tot": _tot,
    "c17.var": _A(3298569, 4167382, 5038164, 5748282, 5771119, 5792435, 5813958, 5835690, 5857633, 5879790),
    "c17.uni": _S(_u1, _u2),
    # cuadro 18 (p. 26)
    "c18.ebitda": [0] + _ebitda[1:], "c18.sant": _A(0, 74058, 48512, 307842, 1029689, 1314400, 1468140, 1582933, 1685801, 1788914),
    "c18.prest": [2791587] + [0] * 10, "c18.lcd": [0] * 11, "c18.ap": [200000, 567782, 371928] + [0] * 8,
    "c18.inv": [655044, 0, 100000] + [0] * 8, "c18.ct": _ctvar, "c18.am": [0] + _amort[1:], "c18.lcr": [0] * 11,
    "c18.fin": [0] + _fin[1:], "c18.ire": [0] * 11,
    "c18.div": [0, 0, 0, 0, 371625, 853745, 1804006, 1902407, 1974240, 2034352, 2094918],
    # cuadro 19 (p. 27-28): valor residual = equipos y construcciones 59.417 + recuperación de capital operativo 3.974.293
    "c19.p_ing": [0] + _ebitda[1:10] + [2215355 + 59417 + 3974293], "c19.p_inv": [655044, 0, 100000] + [0] * 8, "c19.p_ct": _ctvar,
    "c19.p_am": [0] + _amort[1:], "c19.p_fin": [0] + _fin[1:], "c19.p_ire": [0] * 11,
    "c19.van": 4588736, "c19.tir": 0.2261, "c19.pb": 7,
    # cuadro 20 (p. 29)
    "c20.maq": _K(50717), "c20.inf": _K(8700), "c20.otr": _A(12176, 12176, 12176, 12176, 12176, 0, 0, 0, 0, 0),
    "c20.tot": _A(71592, 71592, 71592, 71592, 71592, 59417, 59417, 59417, 59417, 59417),
    # recursos humanos (p. 30-34; años 6-10 = año 5, como el cuadro 10a)
    "rh.ndir": _A(43, 43, 46, 49, 52, 52, 52, 52, 52, 52), "rh.nia": _K(22), "rh.ntot": _A(65, 65, 68, 71, 74, 74, 74, 74, 74, 74),
    "rh.cdir": _mod, "rh.cia": _moa, "rh.ctot": _A(435321, 435321, 453754, 468153, 486109, 486109, 486109, 486109, 486109, 486109),
}
DATOS["plan_cuadros"] = PLAN_CUADROS

# ---------------------------------------------------------------------------------------------------------------------
# Explicación de las diferencias modelo − plan (bloque C de cada cuadro). Causas numeradas en PLAN_NOTAS.
_C1 = "C1 nómina: salarios en Gs convertidos al TC 7.300 (el plan usa 8.000 implícito)."
_C2 = "C2 materia prima: 1% acumulativo también en las cajas (el plan no lo aplica desde el año 3)."
_CL = "Clasificación del personal: el plan pone a los 22 indirectos en administración; el modelo separa 17 de producción y 5 de administración."
_CT = "Capital de trabajo: el plan valúa clientes al costo, incluye caja y sueldos y lo forma al inicio del año; el modelo, a precio de venta y mes a mes."
_CAL = "Calendario: el plan lo registra en el año 0; en el modelo la operación empieza en el mes 1 (año 1)."
_DEP = "Depreciación: el modelo deprecia desde el año 1 y también la inversión adicional del año 2 (10.000 por año); el cuadro 10a del plan empieza en el año 2."
_FLETE = ("El cuadro 10b del plan asigna un flete total de 398.170 (gabinetes 22,19/u; cajas 1,89/u), menor que los 424.000 del cuadro 10a; "
          "el modelo usa el total del 10a (23,20 y 2,08 por unidad).")
PLAN_NOTAS_LINEAS = {
    "c09.mp": _C2, "c09.mp.2": _C2, "c09.mpu.2": _C2,
    "c10.prod": "Clasificación del personal (indirectos en producción), C1 nómina y C2 materia prima.", "c10.mp": _C2, "c10.mod": _C1,
    "c10.carp": "Cargas de los indirectos de producción (clasificación del personal) y C1 nómina.",
    "c10.alip": "Alimentación de los indirectos de producción (clasificación del personal).",
    "c10.fp": "Mantenimiento de planta: 2% compuesto en el modelo, 2% lineal en el plan.", "c10.fc.2": "2% compuesto en el modelo; 2% lineal en el plan.",
    "c10.desp": "Derivada de C2 (3% sobre materia prima e insumos).", "c10.pkg": "Derivada de C1 y C2 (1% sobre materiales y personal directo).",
    "c10.impp": "5% sobre costos de producción: incluye a los indirectos de producción (clasificación) y los efectos de C1 y C2.",
    "c10.adm": _CL, "c10.moa": _CL, "c10.cara": _CL, "c10.alia": _CL,
    "c10.impa": "5% sobre administración: la base del modelo no incluye a los indirectos de producción (clasificación del personal).",
    "c10.dep": _DEP, "c10.tot": "C1 nómina, C2 materia prima, mantenimiento y depreciación.", "c10.cpz": "Sigue a los costos totales.",
    "c10b.tot": "Ver cuadro 10a.", "c10b.cdt": _FLETE + " En cajas, además C2.", "c10b.cd.1": _FLETE, "c10b.cdu.1": _FLETE,
    "c10b.cd.2": _FLETE + " Además C2 desde el año 3.", "c10b.cdu.2": _FLETE + " Además C2 desde el año 3.",
    "c10b.mbu.1": "Sigue al costo directo unitario.", "c10b.mbu.2": "Sigue al costo directo unitario.",
    "c10b.mbp.1": "Margen sobre el precio. El plan publica como «margen bruto %» el cociente costo ÷ precio (57,2% en el año 1).",
    "c10b.mbp.2": "Margen sobre el precio. El plan publica como «margen bruto %» el cociente costo ÷ precio (64,3% en el año 1).",
    "c10b.ot": "C1, mantenimiento y depreciación, menos el flete que el cuadro 10b del plan no asigna.",
    "c10b.oa.1": "Sigue a los otros costos y al costo directo.", "c10b.oa.2": "Sigue a los otros costos y al costo directo.",
    "c10b.ct.1": "Sigue a los costos totales.", "c10b.ct.2": "Sigue a los costos totales.",
    "c10b.ctu.1": "Sigue a los costos totales.", "c10b.ctu.2": "Sigue a los costos totales.",
    "c11.tot": "Ver cuadro 10a.", "c11.mo": _C1 + " El total de personal no cambia por la clasificación.", "c11.mp": _C2,
    "c11.ot": "Mantenimiento compuesto, y packaging e imprevistos derivados de C1 y C2.", "c11.imp": "Despachos derivados de C2.", "c11.dep": _DEP,
    "c12.mo": "Sigue a los rubros del cuadro 11.", "c12.mp": "Sigue a los rubros del cuadro 11.", "c12.ot": "Sigue a los rubros del cuadro 11.",
    "c12.imp": "Sigue a los rubros del cuadro 11.", "c12.fin": "Sigue a los rubros del cuadro 11.", "c12.dep": "Sigue a los rubros del cuadro 11.",
    "c13.des": _CAL,
    "c14.cf": "El modelo informa el IVA incluido en el alquiler (10%), que igual se trata como costo; el plan no lo muestra.",
    "c15.prod": "Clasificación del personal (indirectos en producción), C1 nómina y C2 materia prima.", "c15.adm": _CL,
    "c15.ebitda": "C1 nómina, C2 materia prima de cajas y mantenimiento (ver Conciliacion_Plan_vs_Modelo_ECOSTAR.md).",
    "c15.dep": _DEP, "c15.uai": "EBITDA y depreciación.", "c15.ni": "EBITDA y depreciación.",
    "c15.dec": "Sigue a la utilidad del año.", "c15.ret": "Sigue a la utilidad del año.", "c15.mn": "Sigue a la utilidad del año.",
    "c16.bruto": "El plan incluye caja (3 días) y sueldos (15 días) y valúa clientes al costo; el modelo, a precio de venta.",
    "c16.sue": "El modelo no incluye sueldos y anticipos (punto abierto para decidir).",
    "c16.stk": "Base distinta: el modelo calcula 120 días sobre materia prima e insumos del último mes; el plan, sobre el costo anual de materiales con consumibles.",
    "c16.pt": "El modelo valúa el producto terminado al costo de producción del último mes.",
    "c16.cxc": "El plan valúa clientes al costo (120 días de costos); el modelo, al precio de venta (120 días de ventas).",
    "c16.pas": "Base distinta de proveedores menores (30 días): el modelo incluye otros variables, fijos, administración y tributo.",
    "c16.pot": "Base distinta de proveedores menores (30 días): el modelo incluye otros variables, fijos, administración y tributo.",
    "c16.ctn": "Suma de las diferencias de clientes, stock y proveedores.",
    "c16.var": "Calendario del capital de trabajo: el plan financia en el año 0 el capital del año 1; el modelo lo forma mes a mes.",
    "c17.tot": "Ver cuadro 10a.", "c17.var": "Clasificación: el plan trata energía y mantenimiento como variables; el modelo, como fijos. Además C2.",
    "c17.fij": "Clasificación de costos fijos y variables, C1 nómina y depreciación.",
    "c17.mc": "Sigue a los costos variables.", "c17.mcp": "Sigue a los costos variables.",
    "c17.pev": "Sigue a la clasificación de costos fijos y variables y a C1.", "c17.peu": "Sigue a la clasificación de costos fijos y variables y a C1.",
    "c17.ms": "Sigue al punto de equilibrio.",
    "c18.fue": "Calendario (préstamo y aporte inicial en el año 0 del plan), EBITDA y saldos de caja.", "c18.ebitda": "Ver cuadro 15.",
    "c18.sant": "Arrastre de los saldos de caja de cada año.", "c18.prest": _CAL, "c18.ap": _CAL, "c18.inv": _CAL,
    "c18.uso": "Calendario, capital de trabajo y dividendos.", "c18.ct": "Calendario del capital de trabajo: el plan financia en el año 0 el capital del año 1; el modelo lo forma mes a mes.",
    "c18.div": "Dividendos: el modelo paga solo con caja sobre el mínimo; el plan paga sin esa restricción.",
    "c18.sal": "Suma de las diferencias anteriores.",
    "c19.p_ing": "EBITDA (C1, C2) y valor residual: el modelo recupera el capital de trabajo y el activo fijo neto a valor contable.",
    "c19.p_inv": _CAL, "c19.p_ct": "Calendario y base del capital de trabajo (ver cuadro 16).",
    "c19.p_egr": "Suma de las diferencias anteriores.", "c19.p_flu": "Suma de las diferencias anteriores.",
    "c19.p_act": "Mismo método de descuento que el plan: la diferencia es de cifras.", "c19.p_acu": "Mismo método de descuento que el plan: la diferencia es de cifras.",
    "i_van": "Mismo método que el plan con las cifras del modelo: la diferencia es de cifras (EBITDA, capital de trabajo y calendario), no de método.",
    "i_vanb": "Con los flujos publicados, el método reproduce el VAN del plan: la diferencia es solo de cifras.",
    "i_tir": "Mismo método que el plan con las cifras del modelo.", "i_tirb": "Mismo método que el plan con las cifras del modelo.",
    "i_pb": "Mismo método que el plan con las cifras del modelo.",
    "i_wacc": ("El plan pondera solo el aporte inicial (200.000) y el préstamo; con todos los aportes previstos (1.139.710) la tasa ponderada sube. "
               "Ke y la tasa del préstamo son de 2025: actualizar."),
    "c20.otr": "La inversión adicional del año 2 (100.000) se deprecia en el modelo (10.000 por año); el plan no la deprecia.",
    "c20.tot": "La inversión adicional del año 2 (100.000) se deprecia en el modelo (10.000 por año); el plan no la deprecia.",
    "rh.cdir": _C1, "rh.cia": _C1, "rh.ctot": _C1,
    "c01.cop": _CT, "c01.tini": _CT,
    "c01.item": "Gastos de despacho (3% del CIF) incluidos en cada máquina; el plan los presenta en una línea aparte (14.772 en total).",
    "s_cop": _CT, "s_cost": "C1 nómina, C2 materia prima, mantenimiento y depreciación.",
    "s_ebitda": "C1 nómina, C2 materia prima de cajas y mantenimiento (ver Conciliacion_Plan_vs_Modelo_ECOSTAR.md).",
    "s_ni": "EBITDA y depreciación.", "s_ctn": _CT, "s_caja": "Menor EBITDA, mayor capital de trabajo y calendario de dividendos.",
    "s_pe": "Clasificación de costos fijos y variables y C1 nómina.",
    "s_van": "Mismo método que el plan con las cifras del modelo: la diferencia es de cifras, no de método.",
    "s_tir": "Mismo método que el plan con las cifras del modelo.", "s_pb": "Sigue al VAN con el método del plan.",
    "s_vanp": "Criterio del modelo (FCFF al WACC): otro flujo, no se compara con el VAN publicado.",
    "s_vane": "Criterio del modelo: flujo de los socios (aportes, dividendos y valor final) al Ke.",
}
PLAN_NOTAS = [
    "C1 nómina: el plan convierte los salarios en Gs con un TC implícito de 8.000 (p. 30-34: Gs 18.000.000 = USD 2.250), aunque declara 7.300 (p. 4). El modelo usa 7.300.",
    "C2 materia prima: el texto del plan (p. 2) dice 1% acumulativo desde el año 2; el cuadro 9a no lo aplica a las cajas desde el año 3 (7,48 fijo). El modelo lo aplica a ambos productos.",
    "Clasificación del personal: el plan pone a los 22 indirectos en administración; el modelo separa 17 indirectos de producción y 5 de administración. Cambia la apertura, no el total.",
    "Calendario: el plan registra la inversión, el préstamo, el aporte inicial y el capital operativo inicial en el año 0; en el modelo la operación empieza en el mes 1 (año 1).",
    "Capital de trabajo: el plan valúa clientes al costo, incluye caja (3 días) y sueldos (15 días) y forma el capital al inicio del año; el modelo valúa a precio de venta, trata la caja mínima aparte y no incluye sueldos.",
    "Depreciación: el cuadro 10a del plan empieza en el año 2 y no deprecia la inversión adicional del año 2; el cuadro 20 empieza en el año 1. El modelo deprecia desde la puesta en servicio.",
    "Mantenimiento de planta: 2% compuesto en el modelo; 2% lineal en el plan.",
    "Dividendos: el plan paga sin mirar la caja; el modelo paga solo con caja sobre el mínimo, por eso parte de los dividendos de los primeros años no se paga.",
    "VAN y TIR: el plan descuenta el «flujo de beneficios netos» (resta amortización e intereses, no suma el préstamo). El modelo informa el FCFF y el flujo de los socios, y recalcula el método del plan (cuadro 19).",
    "Inconsistencias internas del plan: flete del 10b (398.170) frente al 10a (424.000); «margen bruto %» del 10b = costo ÷ precio; depreciación del año 1 (10a: 0; cuadro 20: 71.592); pág. 28 con año 0 de −2.791.587 y saldo 4.788.736 frente a −2.991.587 y VAN 4.588.736 del cuadro 19.",
    "Despacho de maquinaria: el modelo lo incluye en cada máquina (3% del CIF); el plan lo presenta en una línea aparte (14.772).",
]
DATOS["plan_notas_lineas"] = PLAN_NOTAS_LINEAS
DATOS["plan_notas"] = PLAN_NOTAS
