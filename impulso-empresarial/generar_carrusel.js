// Carrusel de LinkedIn «Impulso Empresarial» — 8 páginas verticales 8 x 10 in (4:5).
// Uso: NODE_PATH=<node_modules con pptxgenjs, react-icons, react, react-dom, sharp> \
//      PPTX_SKILL_DIR=<carpeta de la skill pptx> node generar_carrusel.js <salida.pptx>
// Todo el contenido es editable: textos, formas y el gráfico de la página 5 son nativos de PowerPoint.

const path = require("path");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const lu = require("react-icons/lu");

const OUT = path.resolve(process.argv[2] || "Impulso_Empresarial_LinkedIn.pptx");
const SKILL = process.env.PPTX_SKILL_DIR;

const THEME = {
  name: "Impulso Empresarial",
  headFontFace: "Arial",
  bodyFontFace: "Arial",
  colors: {
    dk1: "10243D", // azul oscuro: fondos oscuros y texto principal
    lt1: "FFFFFF",
    dk2: "4A5B73", // gris azulado: texto secundario sobre fondo claro
    lt2: "F7F8FC", // fondo claro
    accent1: "276EF1", // azul
    accent2: "82DEFF", // celeste
    accent3: "F4B860", // ámbar: solo para destacar la Línea B del ejemplo
    accent4: "1A3658", // tarjetas sobre fondo oscuro
    accent5: "E8F0FE", // tinte azul claro
    accent6: "B8C7DB", // texto secundario sobre fondo oscuro
    hlink: "276EF1",
    folHlink: "1A55C0",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.defineLayout({ name: "CARRUSEL_4x5", width: 8, height: 10 });
pres.layout = "CARRUSEL_4x5";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.author = "Derlis Javier Morinigo Ovelar";
pres.title = "Impulso Empresarial | De los datos a las decisiones";
pres.subject = "Idea de negocio desarrollada en una tesis de Ingeniería Comercial";
const C = pres.SchemeColor;

const MX = 0.6; // margen lateral
const CW = 6.8; // ancho útil
const TOTAL = 8;
const folio = (n) => `${String(n).padStart(2, "0")}/${String(TOTAL).padStart(2, "0")}`;

// Sombra suave; pptxgenjs modifica el objeto, así que se crea uno nuevo en cada uso.
const sombra = () => ({ type: "outer", color: HEX.dk1, opacity: 0.1, blur: 10, offset: 3, angle: 90 });

async function icono(nombre, hex, px = 384) {
  const svg = ReactDOMServer.renderToStaticMarkup(
    React.createElement(lu[nombre], { size: px, strokeWidth: 2 })
  ).replace(/currentColor/g, "#" + hex);
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------- Diseños (layouts) ----------
function pie(oscuro) {
  const color = oscuro ? C.accent6 : C.text2;
  return [
    {
      text: {
        text: "Impulso Empresarial",
        options: { x: MX, y: 9.3, w: 4, h: 0.32, fontSize: 12, bold: true, color, margin: 0, valign: "middle" },
      },
    },
    {
      placeholder: {
        options: {
          name: "folio", type: "body", x: 5.4, y: 9.3, w: 2.0, h: 0.32,
          fontSize: 12, bold: true, color, align: "right", valign: "middle", margin: 0,
        },
        text: "",
      },
    },
  ];
}

function antetitulo(oscuro) {
  return {
    placeholder: {
      options: {
        name: "eyebrow", type: "body", x: MX, y: 0.6, w: CW, h: 0.35,
        fontSize: 14, bold: true, charSpacing: 2, color: oscuro ? C.accent2 : C.accent1,
        margin: 0, valign: "top",
      },
      text: "",
    },
  };
}

pres.defineSlideMaster({
  title: "IE Portada",
  background: { color: C.text1 },
  objects: [
    antetitulo(true),
    {
      placeholder: {
        options: {
          name: "title", type: "title", x: MX, y: 1.15, w: CW, h: 3.2,
          fontSize: 48, bold: true, color: C.background1, align: "left", margin: 0, valign: "top", lineSpacingMultiple: 0.95,
        },
        text: "",
      },
    },
    ...pie(true),
  ],
});

for (const [titulo, oscuro] of [["IE Contenido claro", false], ["IE Contenido oscuro", true]]) {
  pres.defineSlideMaster({
    title: titulo,
    background: { color: oscuro ? C.text1 : C.background2 },
    objects: [
      antetitulo(oscuro),
      {
        placeholder: {
          options: {
            name: "title", type: "title", x: MX, y: 1.0, w: CW, h: 1.4,
            fontSize: 38, bold: true, color: oscuro ? C.background1 : C.text1,
            align: "left", margin: 0, valign: "top", lineSpacingMultiple: 0.95,
          },
          text: "",
        },
      },
      ...pie(oscuro),
    ],
  });
}

function nuevaPagina(n, masterName, seccion, eyebrow, titulo, notas) {
  const s = pres.addSlide({ masterName, sectionTitle: seccion });
  s.addText(eyebrow, { placeholder: "eyebrow" });
  s.addText(titulo, { placeholder: "title" });
  s.addText(folio(n), { placeholder: "folio" });
  s.addNotes(notas);
  return s;
}

const lineas = (...partes) =>
  partes.map((p, i) => {
    const [text, options] = Array.isArray(p) ? p : [p, {}];
    return { text, options: { ...options, breakLine: i < partes.length - 1 } };
  });

// Tres chips Ventas / Costos / Caja sobre fondo oscuro (portada y cierre).
async function chipsConceptos(s, y) {
  const conceptos = [
    ["Ventas", "LuShoppingCart"],
    ["Costos", "LuReceipt"],
    ["Caja", "LuWallet"],
  ];
  const w = 1.9, gap = (CW - 3 * w) / 2;
  const anchoTexto = { Ventas: 0.92, Costos: 0.95, Caja: 0.6 };
  const centros = [];
  for (const [i, [txt, ic]] of conceptos.entries()) {
    const x = MX + i * (w + gap);
    centros.push(x + w / 2);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w, h: 0.8, rectRadius: 0.12, fill: { color: C.accent4 }, line: { color: C.accent4 },
      objectName: `Chip ${txt}`,
    });
    const x0 = x + (w - (0.4 + 0.14 + anchoTexto[txt])) / 2;
    s.addImage({ data: await icono(ic, HEX.accent2), x: x0, y: y + 0.2, w: 0.4, h: 0.4, altText: txt });
    s.addText(txt, {
      x: x0 + 0.54, y, w: anchoTexto[txt] + 0.2, h: 0.8, fontSize: 20, bold: true, color: C.background1,
      margin: 0, valign: "middle", isTextBox: true, objectName: `Texto ${txt}`,
    });
  }
  return centros;
}

async function main() {
  // ============ 1. PORTADA ============
  pres.addSection({ title: "Apertura" });
  let s = nuevaPagina(
    1, "IE Portada", "Apertura", "PROYECTO DE TESIS · PARAGUAY",
    lineas("Tu negocio", "genera datos.", ["¿Los estás usando", { color: C.accent2 }], ["para decidir?", { color: C.accent2 }]),
    "Portada. Impulso Empresarial se presenta como una idea de negocio desarrollada en la tesis de Ingeniería Comercial de Derlis Javier Morinigo Ovelar; no como una empresa en operación. " +
      "El diagrama resume la propuesta: datos de ventas, costos y caja integrados en información para tomar decisiones. " +
      "Referencia: Trabajo Profesional de Grado (2.ª corrección, 13-09): título de la obra y resumen ejecutivo. No se usa el título académico completo de la tesis."
  );
  s.addText(
    [
      { text: "Impulso Empresarial:", options: { bold: true, color: C.background1, breakLine: true } },
      { text: "inteligencia de negocios para MIPYMES.", options: { color: C.accent6 } },
    ],
    { x: MX, y: 4.5, w: CW, h: 0.85, fontSize: 22, margin: 0, valign: "top", isTextBox: true, objectName: "Subtítulo" }
  );
  const yChip = 5.7;
  const centros = await chipsConceptos(s, yChip);
  const yUnion = 6.85, yCaja = 7.15;
  for (const cx of centros) {
    s.addShape(pres.shapes.LINE, { x: cx, y: yChip + 0.8, w: 0, h: yUnion - yChip - 0.8, line: { color: C.accent2, width: 1.75 }, objectName: "Conector" });
  }
  s.addShape(pres.shapes.LINE, { x: centros[0], y: yUnion, w: centros[2] - centros[0], h: 0, line: { color: C.accent2, width: 1.75 }, objectName: "Conector" });
  s.addShape(pres.shapes.LINE, { x: centros[1], y: yUnion, w: 0, h: yCaja - yUnion, line: { color: C.accent2, width: 1.75 }, objectName: "Conector" });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: MX, y: yCaja, w: CW, h: 0.85, rectRadius: 0.12, fill: { color: C.accent1 }, line: { color: C.accent1 },
    objectName: "Caja información",
  });
  s.addText("Información para tomar decisiones", {
    x: MX, y: yCaja, w: CW, h: 0.85, fontSize: 22, bold: true, color: C.background1, align: "center", valign: "middle",
    margin: 0, isTextBox: true, objectName: "Texto información",
  });
  s.addText("Derlis Javier Morinigo Ovelar", {
    x: MX, y: 8.4, w: CW, h: 0.45, fontSize: 20, bold: true, color: C.background1, margin: 0, valign: "middle",
    isTextBox: true, objectName: "Firma",
  });

  // ============ 2. EL PROBLEMA ============
  s = nuevaPagina(
    2, "IE Contenido claro", "Apertura", "EL PROBLEMA",
    "Tres preguntas para entender tu negocio.",
    "Las tres preguntas son situaciones habituales de gestión, no resultados de una encuesta ni estadísticas verificadas. " +
      "Separan tres conceptos distintos: vender (ventas), ganar (rentabilidad: lo que queda después de costos y gastos) y cobrar (caja y liquidez). " +
      "Un negocio puede vender más y aun así ganar menos, o ganar y no tener efectivo disponible. " +
      "Referencia: Trabajo Profesional de Grado (2.ª corrección, 13-09), 2.2.1: aplicaciones al seguimiento de ventas, costos y gastos, a la rentabilidad por producto o cliente y al control de inventarios y cobranzas. " +
      "La pregunta de caja (efectivo de las próximas semanas) va un paso más allá del alcance escrito en la tesis, que menciona cobranzas pero no proyección de caja."
  );
  const preguntas = [
    ["VENTAS", ["¿Qué productos o clientes", "aportan más?"], "LuShoppingCart"],
    ["RENTABILIDAD", ["¿Qué queda después de", "los costos y gastos?"], "LuPercent"],
    ["CAJA", ["¿Cuánto efectivo necesito", "en las próximas semanas?"], "LuWallet"],
  ];
  for (const [i, [lab, q, ic]] of preguntas.entries()) {
    const y = 2.65 + i * 1.67, h = 1.5;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: MX, y, w: CW, h, rectRadius: 0.14, fill: { color: C.background1 }, line: { color: "E3E8F0", width: 0.75 },
      shadow: sombra(), objectName: `Tarjeta ${lab}`,
    });
    s.addShape(pres.shapes.OVAL, { x: 0.88, y: y + 0.33, w: 0.84, h: 0.84, fill: { color: C.accent5 }, line: { color: C.accent5 }, objectName: `Círculo ${lab}` });
    s.addImage({ data: await icono(ic, HEX.accent1), x: 1.08, y: y + 0.53, w: 0.44, h: 0.44, altText: lab });
    s.addText(lab, {
      x: 2.0, y: y + 0.24, w: 5.1, h: 0.3, fontSize: 15, bold: true, charSpacing: 1.5, color: C.accent1,
      margin: 0, valign: "top", isTextBox: true,
    });
    s.addText(lineas(...q), {
      x: 2.0, y: y + 0.56, w: 5.15, h: 0.8, fontSize: 22, color: C.text1, margin: 0, valign: "top", isTextBox: true,
    });
  }
  s.addText(
    lineas(["Vender, ganar y cobrar", { color: C.accent1 }], ["son cosas diferentes.", { color: C.text1 }]),
    { x: MX, y: 7.85, w: CW, h: 1.0, fontSize: 28, bold: true, margin: 0, valign: "top", isTextBox: true, objectName: "Frase final" }
  );

  // ============ 3. LA PROPUESTA ============
  pres.addSection({ title: "Propuesta" });
  s = nuevaPagina(
    3, "IE Contenido oscuro", "Propuesta", "LA PROPUESTA",
    lineas("Impulso", ["Empresarial", { color: C.accent2 }]),
    "Propuesta de consultoría unipersonal en Inteligencia de Negocios (Business Intelligence, BI) para micro, pequeñas y medianas empresas. " +
      "«Ordenar, visualizar, decidir» es una síntesis comercial para este carrusel. Foco inicial: San Lorenzo y el área metropolitana de Paraguay. " +
      "Es una idea de negocio de la tesis: no se afirma que esté operando ni que tenga clientes. " +
      "Referencia: Trabajo Profesional de Grado (2.ª corrección, 13-09): resumen ejecutivo; 2.1 y Tabla 1 (consultora unipersonal, MIPYMES, San Lorenzo y área metropolitana, atención presencial y virtual); 2.2.2 (mercado de referencia)."
  );
  s.addText("Una consultoría de Inteligencia de Negocios pensada para las MIPYMES.", {
    x: MX, y: 2.6, w: CW, h: 0.85, fontSize: 22, color: C.background1, margin: 0, valign: "top", isTextBox: true,
    objectName: "Descripción",
  });
  const acciones = [
    ["Ordenar", "Integrar y validar los datos.", "LuLayers"],
    ["Visualizar", "Ver indicadores en tableros claros.", "LuChartColumn"],
    ["Decidir", "Elegir el siguiente paso con datos.", "LuSignpost"],
  ];
  const colW = 2.2, colGap = (CW - 3 * colW) / 2, yCirc = 4.0, dCirc = 1.0;
  const cxs = acciones.map((_, i) => MX + i * (colW + colGap) + colW / 2);
  s.addShape(pres.shapes.LINE, {
    x: cxs[0], y: yCirc + dCirc / 2, w: cxs[2] - cxs[0], h: 0, line: { color: C.accent2, width: 2 }, objectName: "Línea de flujo",
  });
  for (const [i, [verbo, desc, ic]] of acciones.entries()) {
    const cx = cxs[i], x = cx - colW / 2;
    s.addShape(pres.shapes.OVAL, {
      x: cx - dCirc / 2, y: yCirc, w: dCirc, h: dCirc, fill: { color: C.accent2 }, line: { color: C.text1, width: 4 },
      objectName: `Círculo ${verbo}`,
    });
    s.addImage({ data: await icono(ic, HEX.dk1), x: cx - 0.24, y: yCirc + 0.26, w: 0.48, h: 0.48, altText: verbo });
    s.addText(verbo, {
      x, y: 5.2, w: colW, h: 0.45, fontSize: 24, bold: true, color: C.background1, align: "center", margin: 0,
      valign: "top", isTextBox: true,
    });
    s.addText(desc, {
      x: x + 0.05, y: 5.7, w: colW - 0.1, h: 1.1, fontSize: 20, color: C.accent6, align: "center", margin: 0,
      valign: "top", isTextBox: true,
    });
  }
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: MX, y: 7.45, w: CW, h: 1.05, rectRadius: 0.14, fill: { color: C.accent4 }, line: { color: C.accent4 },
    objectName: "Caja alcance",
  });
  s.addImage({ data: await icono("LuMapPin", HEX.accent2), x: 0.9, y: 7.73, w: 0.48, h: 0.48, altText: "Ubicación" });
  s.addText(
    [
      { text: "Foco inicial: ", options: { bold: true, color: C.accent2 } },
      { text: "San Lorenzo", options: { color: C.background1, breakLine: true } },
      { text: "y el área metropolitana.", options: { color: C.background1 } },
    ],
    { x: 1.6, y: 7.45, w: 5.55, h: 1.05, fontSize: 20, margin: 0, valign: "middle", isTextBox: true, objectName: "Texto alcance" }
  );

  // ============ 4. LOS SERVICIOS PROPUESTOS ============
  s = nuevaPagina(
    4, "IE Contenido claro", "Propuesta", "LOS SERVICIOS PROPUESTOS",
    "Qué incluiría el servicio.",
    "Alcance propuesto del servicio: diagnóstico de información y necesidades, tableros de control, automatización de reportes, " +
      "definición y análisis de indicadores, capacitación y acompañamiento posterior. Herramientas: Excel y Power BI. " +
      "Lenguaje de propuesta: no se publican precios, paquetes comerciales ni compromisos contractuales. " +
      "Referencia: Trabajo Profesional de Grado (2.ª corrección, 13-09): resumen ejecutivo (diagnóstico, tableros, automatización de reportes, análisis de indicadores, mantenimiento mensual); " +
      "2.2.1 (capacitación inicial; mantenimiento y soporte mensual); Tabla 2 (integración de planillas Excel); 2.8.1 (Power BI). " +
      "Los niveles básico, estándar y avanzado y los precios de la Tabla 10 se omiten a propósito."
  );
  const servicios = [
    ["Diagnóstico", "Datos disponibles y necesidades de gestión.", "LuSearch"],
    ["Tableros de control", "Ventas, costos y resultados en un solo lugar.", "LuLayoutDashboard"],
    ["Reportes automatizados", "Menos tareas repetitivas y mayor consistencia.", "LuRefreshCw"],
    ["Acompañamiento", "Indicadores, capacitación y mejora continua.", "LuGraduationCap"],
  ];
  const cardW = (CW - 0.2) / 2, cardH = 3.1;
  for (const [i, [tit, desc, ic]] of servicios.entries()) {
    const x = MX + (i % 2) * (cardW + 0.2), y = 2.05 + Math.floor(i / 2) * (cardH + 0.2);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: cardW, h: cardH, rectRadius: 0.14, fill: { color: C.background1 }, line: { color: "E3E8F0", width: 0.75 },
      shadow: sombra(), objectName: `Tarjeta ${tit}`,
    });
    s.addShape(pres.shapes.OVAL, { x: x + 0.27, y: y + 0.27, w: 0.74, h: 0.74, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: `Círculo ${tit}` });
    s.addImage({ data: await icono(ic, HEX.lt1), x: x + 0.45, y: y + 0.45, w: 0.38, h: 0.38, altText: tit });
    s.addText(tit, {
      x: x + 0.27, y: y + 1.12, w: cardW - 0.45, h: 0.74, fontSize: 21, bold: true, color: C.text1, margin: 0,
      valign: "bottom", isTextBox: true, lineSpacingMultiple: 0.95,
    });
    s.addText(desc, {
      x: x + 0.27, y: y + 1.94, w: cardW - 0.42, h: 1.02, fontSize: 20, color: C.text2, margin: 0, valign: "top", isTextBox: true,
    });
  }
  s.addImage({ data: await icono("LuWrench", HEX.dk2), x: MX, y: 8.72, w: 0.32, h: 0.32, altText: "Herramientas" });
  s.addText("Herramientas como Excel y Power BI.", {
    x: MX + 0.45, y: 8.67, w: CW - 0.45, h: 0.42, fontSize: 18, color: C.text2, margin: 0, valign: "middle", isTextBox: true,
    objectName: "Herramientas",
  });

  // ============ 5. DEL DATO A LA DECISIÓN ============
  s = nuevaPagina(
    5, "IE Contenido claro", "Propuesta", "DEL DATO A LA DECISIÓN",
    lineas("Un tablero puede mostrar", "dónde revisar el margen."),
    [
      "EJEMPLO ILUSTRATIVO · DATOS SIMULADOS. No pertenecen a clientes, a CIE S.A. ni a resultados de la consultora.",
      "",
      "Datos de un mes, en millones de guaraníes (M = millones de guaraníes):",
      "Línea A: ventas 60; costo de ventas 42; utilidad bruta 18; margen bruto 30 %.",
      "Línea B: ventas 40; costo de ventas 34; utilidad bruta 6; margen bruto 15 %.",
      "Línea C: ventas 20; costo de ventas 14; utilidad bruta 6; margen bruto 30 %.",
      "Total: ventas 120; costo de ventas 90; utilidad bruta 30; margen bruto 25 %.",
      "",
      "Fórmula: margen bruto = (ventas − costo de ventas) / ventas.",
      "El 25 % total surge de 30 / 120, es decir, está ponderado por ventas. En este ejemplo el promedio simple de 30 %, 15 % y 30 % también da 25 %, pero solo por coincidencia: si cambia la mezcla de ventas, el promedio simple deja de coincidir y el cálculo correcto sigue siendo utilidad bruta total / ventas totales.",
      "Margen bruto: antes de gastos operativos, intereses e impuestos. No es utilidad neta.",
      "Cuentas por cobrar al cierre: Gs. 45 millones. Es un saldo adicional simulado; no son cobros del mes ni flujo de caja, ni una parte adicional de las ventas que deba sumarse a los Gs. 120 millones.",
      "Lectura: la Línea B tiene la mitad del margen de A y C. El tablero no da la respuesta; indica dónde revisar (precios, costos o descuentos) antes de decidir.",
      "El gráfico es nativo de PowerPoint: los datos se editan con clic derecho > Editar datos.",
      "Referencia: Trabajo Profesional de Grado (2.ª corrección, 13-09), 2.2.1: «un tablero comercial podrá mostrar la evolución de las ventas y los márgenes por producto». Los importes del ejemplo no provienen de la tesis.",
    ].join("\n")
  );
  // Panel del tablero
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: MX, y: 2.5, w: CW, h: 5.1, rectRadius: 0.14, fill: { color: C.background1 }, line: { color: "E3E8F0", width: 0.75 },
    shadow: sombra(), objectName: "Panel tablero",
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.85, y: 2.72, w: 5.0, h: 0.42, rectRadius: 0.21, fill: { color: C.text1 }, line: { color: C.text1 },
    objectName: "Etiqueta datos simulados",
  });
  s.addText("EJEMPLO ILUSTRATIVO · DATOS SIMULADOS", {
    x: 0.85, y: 2.72, w: 5.0, h: 0.42, fontSize: 13, bold: true, charSpacing: 1, color: C.background1, align: "center",
    valign: "middle", margin: 0, isTextBox: true, objectName: "Texto datos simulados",
  });
  const kpis = [
    [[{ text: "Gs. ", options: { fontSize: 18 } }, { text: "120 M", options: { fontSize: 28 } }], "Ventas del mes"],
    [[{ text: "25 %", options: { fontSize: 28 } }], "Margen bruto"],
    [[{ text: "Gs. ", options: { fontSize: 18 } }, { text: "45 M", options: { fontSize: 28 } }], "Cuentas por cobrar al cierre"],
  ];
  const kW = (6.3 - 0.3) / 3;
  for (const [i, [valor, etiqueta]] of kpis.entries()) {
    const x = 0.85 + i * (kW + 0.15), y = 3.32;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y, w: kW, h: 1.28, rectRadius: 0.1, fill: { color: C.accent5 }, line: { color: C.accent5 }, objectName: `KPI ${etiqueta}`,
    });
    s.addText(valor, {
      x: x + 0.15, y: y + 0.14, w: kW - 0.22, h: 0.5, bold: true, color: C.text1, margin: 0, valign: "bottom", isTextBox: true,
    });
    s.addText(etiqueta, {
      x: x + 0.15, y: y + 0.7, w: kW - 0.22, h: 0.52, fontSize: 15, color: C.text2, margin: 0, valign: "top", isTextBox: true,
    });
  }
  s.addText("Margen bruto por línea", {
    x: 0.85, y: 4.8, w: 5, h: 0.35, fontSize: 16, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true,
    objectName: "Título gráfico",
  });
  s.addChart(
    pres.charts.BAR,
    [{ name: "Margen bruto", labels: ["Línea A", "Línea B", "Línea C"], values: [0.3, 0.15, 0.3] }],
    {
      x: 0.75, y: 5.15, w: 6.5, h: 2.32, barDir: "col", barGapWidthPct: 70,
      chartColors: [HEX.accent1, HEX.accent3, HEX.accent1],
      showLegend: false, showTitle: false,
      showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0 %",
      dataLabelFontSize: 16, dataLabelFontBold: true, dataLabelColor: HEX.dk1, dataLabelFontFace: "+mn-lt",
      catAxisLabelFontSize: 16, catAxisLabelColor: HEX.dk1, catAxisLabelFontFace: "+mn-lt",
      catAxisLineShow: true, catAxisLineColor: "C9D3E1",
      valAxisMinVal: 0, valAxisMaxVal: 0.4, valAxisMajorUnit: 0.1, valAxisLabelFormatCode: "0 %",
      valAxisLabelFontSize: 12, valAxisLabelColor: HEX.dk2, valAxisLabelFontFace: "+mn-lt", valAxisLineShow: false,
      valGridLine: { color: "E3E8F0", size: 0.75 }, catGridLine: { style: "none" },
      objectName: "Gráfico margen bruto por línea",
      altText: "Gráfico de columnas con datos simulados: margen bruto Línea A 30 %, Línea B 15 %, Línea C 30 %.",
    }
  );
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: MX, y: 7.8, w: CW, h: 0.78, rectRadius: 0.12, fill: { color: C.text1 }, line: { color: C.text1 }, objectName: "Caja conclusión",
  });
  s.addShape(pres.shapes.OVAL, { x: 0.88, y: 8.08, w: 0.22, h: 0.22, fill: { color: C.accent3 }, line: { color: C.accent3 }, objectName: "Marcador Línea B" });
  s.addText(
    [
      { text: "Línea B: ", options: { color: C.accent3 } },
      { text: "revisar precios, costos o descuentos.", options: { color: C.background1 } },
    ],
    { x: 1.25, y: 7.8, w: 6.0, h: 0.78, fontSize: 19, bold: true, margin: 0, valign: "middle", isTextBox: true, objectName: "Conclusión" }
  );
  s.addText(lineas("M = millones de guaraníes.", "Margen bruto: antes de gastos operativos, intereses e impuestos."), {
    x: MX, y: 8.7, w: CW, h: 0.48, fontSize: 13, color: C.text2, margin: 0, valign: "top", isTextBox: true, objectName: "Aclaración",
  });

  // ============ 6. LA FORMA DE TRABAJAR ============
  s = nuevaPagina(
    6, "IE Contenido oscuro", "Propuesta", "LA FORMA DE TRABAJAR",
    lineas("Así se plantea", "el acompañamiento."),
    "Proceso propuesto en cuatro pasos: entender las decisiones, ordenar y validar los datos, implementar tablero y reportes según el alcance acordado, " +
      "y acompañar el uso de los indicadores. Modalidad: implementación por proyecto más seguimiento mensual. " +
      "No se fijan plazos, cantidad de reuniones ni compromisos contractuales. " +
      "Referencia: Trabajo Profesional de Grado (2.ª corrección, 13-09): 7.4, Tablas 24, 25 y 27 (diagnóstico y propuesta, ejecución del servicio, seguimiento posventa) y 2.2.1 (implementación por proyecto; mantenimiento como prestación recurrente). " +
      "En la tesis la prestación mensual se denomina «mantenimiento y soporte»; «seguimiento mensual» es la forma simplificada usada en el carrusel."
  );
  const pasos = [
    ["Entender", "Qué decisiones necesita tomar el negocio."],
    ["Ordenar", "Qué datos hay y qué falta validar."],
    ["Implementar", "Tablero y reportes según el alcance acordado."],
    ["Acompañar", ["Cómo usar los indicadores", "y mantenerlos útiles."]],
  ];
  const yPaso = (i) => 2.68 + i * 1.28, dNum = 0.7;
  s.addShape(pres.shapes.LINE, {
    x: MX + dNum / 2, y: yPaso(0) + dNum / 2, w: 0, h: yPaso(3) - yPaso(0), line: { color: C.accent1, width: 2 },
    objectName: "Línea de pasos",
  });
  for (const [i, [nombre, desc]] of pasos.entries()) {
    const y = yPaso(i);
    s.addShape(pres.shapes.OVAL, {
      x: MX, y, w: dNum, h: dNum, fill: { color: C.accent1 }, line: { color: C.text1, width: 3 }, objectName: `Paso ${i + 1}`,
    });
    s.addText(String(i + 1), {
      x: MX, y, w: dNum, h: dNum, fontSize: 22, bold: true, color: C.background1, align: "center", valign: "middle",
      margin: 0, isTextBox: true,
    });
    s.addText(nombre, {
      x: 1.6, y: y - 0.02, w: 5.8, h: 0.42, fontSize: 22, bold: true, color: C.background1, margin: 0, valign: "top", isTextBox: true,
    });
    s.addText(Array.isArray(desc) ? lineas(...desc) : desc, {
      x: 1.6, y: y + 0.42, w: 5.8, h: 0.72, fontSize: 20, color: C.accent6, margin: 0, valign: "top", isTextBox: true,
    });
  }
  const modalidad = [
    ["Implementación por proyecto", "LuFolderKanban"],
    ["Seguimiento mensual", "LuCalendarSync"],
  ];
  const chipW = 3.0, yMod = 8.0;
  for (const [i, [txt, ic]] of modalidad.entries()) {
    const x = i === 0 ? MX : MX + CW - chipW;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x, y: yMod, w: chipW, h: 0.92, rectRadius: 0.12, fill: { color: C.accent4 }, line: { color: C.accent4 }, objectName: `Chip ${txt}`,
    });
    s.addImage({ data: await icono(ic, HEX.accent2), x: x + 0.24, y: yMod + 0.26, w: 0.4, h: 0.4, altText: txt });
    s.addText(txt, {
      x: x + 0.8, y: yMod, w: chipW - 0.95, h: 0.92, fontSize: 18, bold: true, color: C.background1, margin: 0,
      valign: "middle", isTextBox: true,
    });
  }
  s.addText("+", {
    x: MX + chipW, y: yMod, w: CW - 2 * chipW, h: 0.92, fontSize: 30, bold: true, color: C.accent2, align: "center",
    valign: "middle", margin: 0, isTextBox: true, objectName: "Signo más",
  });

  // ============ 7. EL ORIGEN DE LA IDEA ============
  pres.addSection({ title: "Origen y cierre" });
  s = nuevaPagina(
    7, "IE Contenido claro", "Origen y cierre", "EL ORIGEN DE LA IDEA",
    lineas("Una idea que une", "finanzas y datos."),
    "Contexto profesional del autor, sin convertir la página en un currículum. Impulso Empresarial se desarrolló como proyecto de tesis de Ingeniería Comercial; " +
      "no se usa el título de Ingeniero Comercial ni se afirma que la tesis ya fue defendida. No se menciona al empleador actual ni logros comerciales de Impulso Empresarial. " +
      "Referencia: Trabajo Profesional de Grado (2.ª corrección, 13-09): portada (autor y carrera) y 4.8 (la experiencia profesional del consultor como parte de la ventaja competitiva)."
  );
  const dV = 2.0, yV = 2.6, xV1 = 2.25, xV2 = 3.75;
  s.addShape(pres.shapes.OVAL, { x: xV1, y: yV, w: dV, h: dV, fill: { color: C.accent1, transparency: 5 }, line: { color: C.accent1, transparency: 100 }, objectName: "Círculo finanzas" });
  s.addShape(pres.shapes.OVAL, { x: xV2, y: yV, w: dV, h: dV, fill: { color: C.accent2, transparency: 30 }, line: { color: C.accent2, transparency: 100 }, objectName: "Círculo datos" });
  s.addText("Finanzas", {
    x: xV1 + 0.05, y: yV + dV / 2 - 0.25, w: xV2 - xV1 - 0.1, h: 0.5, fontSize: 20, bold: true, color: C.background1,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("Datos", {
    x: xV1 + dV + 0.05, y: yV + dV / 2 - 0.25, w: xV2 - xV1 - 0.1, h: 0.5, fontSize: 20, bold: true, color: C.text1,
    align: "center", valign: "middle", margin: 0, isTextBox: true,
  });
  s.addText("Desarrollé Impulso Empresarial en mi tesis de Ingeniería Comercial.", {
    x: MX, y: 4.9, w: CW, h: 0.8, fontSize: 22, bold: true, color: C.text1, margin: 0, valign: "top", isTextBox: true,
    objectName: "Origen",
  });
  s.addText("Mi experiencia en costos, presupuestos y planificación financiera orienta esta propuesta.", {
    x: MX, y: 5.78, w: CW, h: 0.75, fontSize: 20, color: C.text2, margin: 0, valign: "top", isTextBox: true,
    objectName: "Experiencia",
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: MX, y: 6.72, w: CW, h: 1.05, rectRadius: 0.14, fill: { color: C.text1 }, line: { color: C.text1 }, objectName: "Caja frase",
  });
  s.addImage({ data: await icono("LuCompass", HEX.accent2), x: 0.92, y: 7.0, w: 0.5, h: 0.5, altText: "Brújula" });
  s.addText(lineas("Entender los números para", "elegir el siguiente paso."), {
    x: 1.65, y: 6.72, w: 5.5, h: 1.05, fontSize: 22, bold: true, color: C.background1, margin: 0, valign: "middle",
    isTextBox: true, objectName: "Frase destacada",
  });
  s.addShape(pres.shapes.OVAL, { x: MX, y: 8.08, w: 0.8, h: 0.8, fill: { color: C.accent5 }, line: { color: C.accent5 }, objectName: "Círculo autor" });
  s.addImage({ data: await icono("LuUserRound", HEX.accent1), x: MX + 0.19, y: 8.27, w: 0.42, h: 0.42, altText: "Autor" });
  s.addText("Derlis Javier Morinigo Ovelar", {
    x: 1.62, y: 8.08, w: 5.78, h: 0.42, fontSize: 20, bold: true, color: C.text1, margin: 0, valign: "middle",
    isTextBox: true, objectName: "Firma",
  });
  s.addText("Especialista en Planificación Financiera", {
    x: 1.62, y: 8.5, w: 5.78, h: 0.38, fontSize: 16, color: C.text2, margin: 0, valign: "middle",
    isTextBox: true, objectName: "Cargo",
  });

  // ============ 8. CIERRE ============
  s = nuevaPagina(
    8, "IE Portada", "Origen y cierre", "CONVERSEMOS",
    lineas("¿Qué dato te", "ayudaría a", ["decidir mejor?", { color: C.accent2 }]),
    "Invitación a conversar para seguir afinando la propuesta y a visitar el perfil de LinkedIn. " +
      "No se agregan teléfono, correo, QR, enlaces ni ofertas (por ejemplo, diagnósticos gratuitos). " +
      "Referencia: Trabajo Profesional de Grado (2.ª corrección, 13-09): 2.8.1 prevé LinkedIn como canal de contacto; el resumen ejecutivo indica que la captación y la disposición de pago deben validarse antes de implementar. " +
      "Las conversaciones que genere el carrusel son señales cualitativas, no una validación de mercado."
  );
  await chipsConceptos(s, 4.0);
  s.addText(lineas("Me interesa conversar con emprendedores", "y profesionales para seguir afinando", "esta propuesta."), {
    x: MX, y: 5.2, w: CW, h: 1.2, fontSize: 22, color: C.background1, margin: 0, valign: "top", isTextBox: true,
    objectName: "Invitación",
  });
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: MX, y: 6.85, w: CW, h: 0.95, rectRadius: 0.475, fill: { color: C.background1 }, line: { color: C.background1 },
    objectName: "Botón llamada a la acción",
  });
  s.addText("Visita mi perfil y conversemos.", {
    x: 1.0, y: 6.85, w: 5.3, h: 0.95, fontSize: 22, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true,
    objectName: "Llamada a la acción",
  });
  s.addShape(pres.shapes.OVAL, { x: 6.6, y: 6.99, w: 0.67, h: 0.67, fill: { color: C.accent1 }, line: { color: C.accent1 }, objectName: "Círculo flecha" });
  s.addImage({ data: await icono("LuArrowRight", HEX.lt1), x: 6.75, y: 7.14, w: 0.37, h: 0.37, altText: "Flecha" });
  s.addText("Derlis Javier Morinigo Ovelar", {
    x: MX, y: 8.3, w: CW, h: 0.45, fontSize: 20, bold: true, color: C.background1, margin: 0, valign: "middle",
    isTextBox: true, objectName: "Firma",
  });

  await pres.writeFile({ fileName: OUT });
  if (SKILL) {
    const { applyTheme } = require(path.join(SKILL, "scripts/apply_theme.js"));
    await applyTheme(OUT, THEME);
  }
  console.log("OK", OUT);
}

main().catch((e) => {
  console.error(e);
  process.exit(1);
});
