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
