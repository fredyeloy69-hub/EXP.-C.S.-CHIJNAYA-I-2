import ExcelJS from "exceljs";
import { LOGO_PUNO_BASE64 } from "./logoPuno";

const ESTADO_LABEL = {
  completa: "COMPLETA",
  incompleta: "INCOMPLETA",
  vacia: "VACÍA",
};

const ESTADO_COLOR = {
  completa: "FF2ECC71",
  incompleta: "FFF39C12",
  vacia: "FFE74C3C",
};

const AZUL_OSCURO = "FF141E3C";
const NARANJA = "FFE67E22";

function comparaNatural(a, b) {
  return (a || "").localeCompare(b || "", undefined, { numeric: true, sensitivity: "base" });
}

function mezclarConBlanco(hexRgb, factor) {
  const r = parseInt(hexRgb.slice(2, 4), 16);
  const g = parseInt(hexRgb.slice(4, 6), 16);
  const b = parseInt(hexRgb.slice(6, 8), 16);
  const mezcla = (c) => Math.round(c + (255 - c) * factor).toString(16).padStart(2, "0").toUpperCase();
  return `FF${mezcla(r)}${mezcla(g)}${mezcla(b)}`;
}

const FACTORES_HOJA = [0.72, 0.78, 0.83, 0.87, 0.90, 0.93, 0.96];
function colorHoja(nivelVisual) {
  const factor = FACTORES_HOJA[Math.min((nivelVisual || 1) - 1, FACTORES_HOJA.length - 1)];
  return mezclarConBlanco(NARANJA, factor);
}

function insertarLogoYEncabezados(sheet, workbook, logoBase64, proyectoNombre, subtituloTexto, usuarioFirma) {
  // Inserción del Logo
  if (logoBase64) {
    try {
      const base64Data = logoBase64.includes('base64,') ? logoBase64.split('base64,')[1] : logoBase64;
      const imageId = workbook.addImage({
        base64: base64Data,
        extension: 'png',
      });
      sheet.addImage(imageId, {
        tl: { col: 0.2, row: 0.2 },
        ext: { width: 70, height: 70 }
      });
    } catch (e) {
      console.warn("No se pudo cargar el logo en el Excel:", e);
    }
  }

  sheet.mergeCells("A1:D1");
  sheet.getCell("A1").value = "GOBIERNO REGIONAL DE PUNO — GERENCIA REGIONAL DE INFRAESTRUCTURA";
  sheet.getCell("A1").font = { bold: true, size: 11 };
  sheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(1).height = 18;

  sheet.mergeCells("A2:D2");
  sheet.getCell("A2").value = "SUB GERENCIA DE ESTUDIOS DEFINITIVOS";
  sheet.getCell("A2").font = { bold: true, size: 9.5 };
  sheet.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(2).height = 15;

  sheet.mergeCells("A3:D3");
  const celdaProyecto = sheet.getCell("A3");
  celdaProyecto.value = `PROYECTO: ${proyectoNombre.toUpperCase()}`;
  celdaProyecto.font = { bold: true, size: 9.5, color: { argb: "FF4A2D0F" } };
  celdaProyecto.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
  sheet.getRow(3).height = 28; 

  sheet.mergeCells("A4:D4");
  const celdaTitulo = sheet.getCell("A4");
  celdaTitulo.value = subtituloTexto;
  celdaTitulo.font = { bold: true, size: 11, color: { argb: "FFFFFFFF" } };
  celdaTitulo.alignment = { horizontal: "center", vertical: "middle" };
  celdaTitulo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_OSCURO } };
  sheet.getRow(4).height = 22;

  sheet.mergeCells("A5:B5");
  sheet.getCell("A5").value = `Generado por: ${usuarioFirma} — ${new Date().toLocaleString("es-PE")}`;
  sheet.getCell("A5").font = { italic: true, size: 8.5, color: { argb: "FF666666" } };
}

/**
 * Genera y descarga un reporte en Excel (.xlsx) para un área específica.
 */
export async function generarReporteExcelPorArea(
  areaNombre,
  carpetasDelArea,
  opts = {}
) {
  const {
    logoBase64 = LOGO_PUNO_BASE64,
    proyectoNombre = '"MEJORAMIENTO DEL SERVICIO DE ATENCION DE SALUD BASICOS EN CHIJNAYA DISTRITO DE PUCARA DE LA PROVINCIA DE LAMPA DEPARTAMENTO DE PUNO"',
    usuarioFirma = "Sistema Chijnaya",
  } = opts;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Visor Chijnaya";
  workbook.created = new Date();

  const nombreHoja = areaNombre.replace(/[\\/*?:[\]]/g, "").slice(0, 31) || "Reporte";
  const sheet = workbook.addWorksheet(nombreHoja, {
    pageSetup: { 
      orientation: "portrait", 
      fitToPage: true, 
      fitToWidth: 1,
      margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }
    },
  });

  sheet.columns = [
    { width: 6 },
    { width: 55 },
    { width: 16 },
    { width: 55 },
  ];

  insertarLogoYEncabezados(sheet, workbook, logoBase64, proyectoNombre, `REPORTE DE AVANCE — ${areaNombre.toUpperCase()}`, usuarioFirma);

  const total = carpetasDelArea.length;
  const completas = carpetasDelArea.filter((c) => c.estado === "completa").length;
  const incompletas = carpetasDelArea.filter((c) => c.estado === "incompleta").length;
  const vacias = carpetasDelArea.filter((c) => c.estado === "vacia").length;

  sheet.mergeCells("C5:D5");
  sheet.getCell("C5").value = `Total: ${total}  ·  Completas: ${completas}  ·  Incompletas: ${incompletas}  ·  Vacías: ${vacias}`;
  sheet.getCell("C5").font = { italic: true, size: 8.5, color: { argb: "FF666666" } };
  sheet.getCell("C5").alignment = { horizontal: "right" };

  sheet.addRow(); // fila 6 en blanco

  const FILA_ENCABEZADO = 7;
  const filaEncabezado = sheet.getRow(FILA_ENCABEZADO);
  filaEncabezado.values = ["N°", "DESCRIPCIÓN DE CARPETA", "ESTADO", "DETALLE / OBSERVACIÓN"];
  filaEncabezado.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_OSCURO } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = { bottom: { style: "medium", color: { argb: "FF000000" } } };
  });
  filaEncabezado.height = 18;

  const grupos = {};
  const ordenGrupos = [];
  for (const c of carpetasDelArea) {
    const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
    const especialidad = partes.length > 1 ? partes[1] : "(raíz)";
    if (!grupos[especialidad]) {
      grupos[especialidad] = [];
      ordenGrupos.push(especialidad);
    }
    grupos[especialidad].push(c);
  }
  ordenGrupos.sort(comparaNatural);

  let contadorFila = 1;

  function agregarSubEncabezado(nombre, nivel) {
    const fila = sheet.addRow([`>  ${nombre}`]);
    sheet.mergeCells(`A${fila.number}:D${fila.number}`);
    const celda = fila.getCell(1);
    const esNivel1 = nivel === 1;
    const factoresPorNivel = [0, 0.40, 0.55, 0.68, 0.78, 0.86, 0.92]; 
    const factor = factoresPorNivel[Math.min(nivel - 1, factoresPorNivel.length - 1)];
    celda.font = { bold: true, size: esNivel1 ? 11 : 9.5, color: { argb: esNivel1 ? "FFFFFFFF" : "FF4A2D0F" } };
    celda.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: mezclarConBlanco(NARANJA, factor) },
    };
    celda.alignment = { horizontal: "left", vertical: "middle", indent: Math.max(0, nivel - 1) * 2 };
    fila.height = esNivel1 ? 19 : 16;
  }

  function agregarFila(c, nivelVisual) {
    const estado = c.estado || "incompleta";
    const nombreMostrado = c.nombre || (c.ruta || "").split(" / ").pop() || "-";
    const fila = sheet.addRow([contadorFila++, nombreMostrado, ESTADO_LABEL[estado] || estado.toUpperCase(), c.detalle || "-"]);
    const fondo = colorHoja(nivelVisual);

    fila.getCell(1).alignment = { horizontal: "center", vertical: "top" };
    fila.getCell(2).alignment = { wrapText: true, vertical: "top", indent: Math.max(0, (nivelVisual || 1) - 1) * 2 };
    fila.getCell(3).font = { bold: true, color: { argb: ESTADO_COLOR[estado] || "FF666666" } };
    fila.getCell(3).alignment = { horizontal: "center", vertical: "top" };
    fila.getCell(4).alignment = { wrapText: true, vertical: "top" };

    fila.eachCell((cell) => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fondo } };
      cell.border = {
        top: { style: "hair", color: { argb: "FFDDDDDD" } },
        bottom: { style: "hair", color: { argb: "FFDDDDDD" } },
        left: { style: "hair", color: { argb: "FFDDDDDD" } },
        right: { style: "hair", color: { argb: "FFDDDDDD" } },
      };
    });
  }

  const NIVEL_MAX_INTERMEDIO = 6;
  function agruparRecursivo(items, nivelIdxRuta, nivelVisual) {
    if (nivelIdxRuta > NIVEL_MAX_INTERMEDIO) {
      const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
      ordenado.forEach((c) => agregarFila(c, nivelVisual));
      return;
    }

    const subgrupos = {};
    let hayNivelMasProfundo = false;
    for (const c of items) {
      const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
      const clave = partes.length >= nivelIdxRuta + 1 ? partes[nivelIdxRuta] : null;
      if (clave) hayNivelMasProfundo = true;
      const key = clave || `__directo__${c.nombre || c.id}`;
      if (!subgrupos[key]) subgrupos[key] = [];
      subgrupos[key].push(c);
    }

    if (!hayNivelMasProfundo) {
      const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
      ordenado.forEach((c) => agregarFila(c, nivelVisual));
      return;
    }

    const entradas = Object.keys(subgrupos).map((key) => ({
      key,
      nombreOrden: key.startsWith("__directo__") ? subgrupos[key][0].nombre || "" : key,
      esGrupo: !key.startsWith("__directo__"),
    }));
    entradas.sort((a, b) => comparaNatural(a.nombreOrden, b.nombreOrden));

    for (const entrada of entradas) {
      if (entrada.esGrupo) {
        agregarSubEncabezado(entrada.key, nivelVisual);
        agruparRecursivo(subgrupos[entrada.key], nivelIdxRuta + 1, nivelVisual + 1);
      } else {
        const ordenado = [...subgrupos[entrada.key]].sort((a, b) => comparaNatural(a.nombre, b.nombre));
        ordenado.forEach((c) => agregarFila(c, nivelVisual));
      }
    }
  }

  for (const especialidad of ordenGrupos) {
    agregarSubEncabezado(especialidad.toUpperCase(), 1);
    agruparRecursivo(grupos[especialidad], 2, 2);
  }

  sheet.views = [{ state: "frozen", ySplit: FILA_ENCABEZADO }];

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `Reporte_${areaNombre.replace(/[^a-zA-Z0-9]+/g, "_")}_${new Date()
    .toISOString()
    .slice(0, 10)}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

/**
 * Genera y descarga un reporte Excel Consolidado Global con pestañas por área.
 */
export async function generarReporteExcelConsolidado(
  todasLasCarpetas,
  opts = {}
) {
  const {
    logoBase64 = LOGO_PUNO_BASE64,
    proyectoNombre = '"MEJORAMIENTO DEL SERVICIO DE ATENCION DE SALUD BASICOS EN CHIJNAYA DISTRITO DE PUCARA DE LA PROVINCIA DE LAMPA DEPARTAMENTO DE PUNO"',
    usuarioFirma = "Sistema Chijnaya",
  } = opts;

  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Visor Chijnaya";
  workbook.created = new Date();

  // Agrupar por área
  const areasMap = {};
  for (const c of todasLasCarpetas) {
    const area = c.area || "Sin área";
    if (!areasMap[area]) areasMap[area] = [];
    areasMap[area].push(c);
  }
  const ordenAreas = Object.keys(areasMap).sort((a, b) => a.localeCompare(b));

  for (const areaNombre of ordenAreas) {
    const carpetasDelArea = areasMap[areaNombre];
    const nombreHoja = areaNombre.replace(/[\\/*?:[\]]/g, "").slice(0, 31) || "Área";
    const sheet = workbook.addWorksheet(nombreHoja, {
      pageSetup: { 
        orientation: "portrait", 
        fitToPage: true, 
        fitToWidth: 1,
        margins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }
      },
    });

    sheet.columns = [
      { width: 6 },
      { width: 55 },
      { width: 16 },
      { width: 55 },
    ];

    insertarLogoYEncabezados(sheet, workbook, logoBase64, proyectoNombre, `CONSOLIDADO — ÁREA: ${areaNombre.toUpperCase()}`, usuarioFirma);

    const total = carpetasDelArea.length;
    const completas = carpetasDelArea.filter((c) => c.estado === "completa").length;
    const incompletas = carpetasDelArea.filter((c) => c.estado === "incompleta").length;
    const vacias = carpetasDelArea.filter((c) => c.estado === "vacia").length;

    sheet.mergeCells("C5:D5");
    sheet.getCell("C5").value = `Total: ${total}  ·  Completas: ${completas}  ·  Incompletas: ${incompletas}  ·  Vacías: ${vacias}`;
    sheet.getCell("C5").font = { italic: true, size: 8.5, color: { argb: "FF666666" } };
    sheet.getCell("C5").alignment = { horizontal: "right" };

    sheet.addRow(); // fila 6 en blanco

    const FILA_ENCABEZADO = 7;
    const filaEncabezado = sheet.getRow(FILA_ENCABEZADO);
    filaEncabezado.values = ["N°", "DESCRIPCIÓN DE CARPETA", "ESTADO", "DETALLE / OBSERVACIÓN"];
    filaEncabezado.eachCell((cell) => {
      cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_OSCURO } };
      cell.alignment = { vertical: "middle", horizontal: "center" };
      cell.border = { bottom: { style: "medium", color: { argb: "FF000000" } } };
    });
    filaEncabezado.height = 18;

    const grupos = {};
    const ordenGrupos = [];
    for (const c of carpetasDelArea) {
      const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
      const especialidad = partes.length > 1 ? partes[1] : "(raíz)";
      if (!grupos[especialidad]) {
        grupos[especialidad] = [];
        ordenGrupos.push(especialidad);
      }
      grupos[especialidad].push(c);
    }
    ordenGrupos.sort(comparaNatural);

    let contadorFila = 1;

    function agregarSubEncabezado(nombre, nivel) {
      const fila = sheet.addRow([`>  ${nombre}`]);
      sheet.mergeCells(`A${fila.number}:D${fila.number}`);
      const celda = fila.getCell(1);
      const esNivel1 = nivel === 1;
      const factoresPorNivel = [0, 0.40, 0.55, 0.68, 0.78, 0.86, 0.92]; 
      const factor = factoresPorNivel[Math.min(nivel - 1, factoresPorNivel.length - 1)];
      celda.font = { bold: true, size: esNivel1 ? 11 : 9.5, color: { argb: esNivel1 ? "FFFFFFFF" : "FF4A2D0F" } };
      celda.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: mezclarConBlanco(NARANJA, factor) },
      };
      celda.alignment = { horizontal: "left", vertical: "middle", indent: Math.max(0, nivel - 1) * 2 };
      fila.height = esNivel1 ? 19 : 16;
    }

    function agregarFila(c, nivelVisual) {
      const estado = c.estado || "incompleta";
      const nombreMostrado = c.nombre || (c.ruta || "").split(" / ").pop() || "-";
      const fila = sheet.addRow([contadorFila++, nombreMostrado, ESTADO_LABEL[estado] || estado.toUpperCase(), c.detalle || "-"]);
      const fondo = colorHoja(nivelVisual);

      fila.getCell(1).alignment = { horizontal: "center", vertical: "top" };
      fila.getCell(2).alignment = { wrapText: true, vertical: "top", indent: Math.max(0, (nivelVisual || 1) - 1) * 2 };
      fila.getCell(3).font = { bold: true, color: { argb: ESTADO_COLOR[estado] || "FF666666" } };
      fila.getCell(3).alignment = { horizontal: "center", vertical: "top" };
      fila.getCell(4).alignment = { wrapText: true, vertical: "top" };

      fila.eachCell((cell) => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fondo } };
        cell.border = {
          top: { style: "hair", color: { argb: "FFDDDDDD" } },
          bottom: { style: "hair", color: { argb: "FFDDDDDD" } },
          left: { style: "hair", color: { argb: "FFDDDDDD" } },
          right: { style: "hair", color: { argb: "FFDDDDDD" } },
        };
      });
    }

    const NIVEL_MAX_INTERMEDIO = 6;
    function agruparRecursivo(items, nivelIdxRuta, nivelVisual) {
      if (nivelIdxRuta > NIVEL_MAX_INTERMEDIO) {
        const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
        ordenado.forEach((c) => agregarFila(c, nivelVisual));
        return;
      }

      const subgrupos = {};
      let hayNivelMasProfundo = false;
      for (const c of items) {
        const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
        const clave = partes.length >= nivelIdxRuta + 1 ? partes[nivelIdxRuta] : null;
        if (clave) hayNivelMasProfundo = true;
        const key = clave || `__directo__${c.nombre || c.id}`;
        if (!subgrupos[key]) subgrupos[key] = [];
        subgrupos[key].push(c);
      }

      if (!hayNivelMasProfundo) {
        const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
        ordenado.forEach((c) => agregarFila(c, nivelVisual));
        return;
      }

      const entradas = Object.keys(subgrupos).map((key) => ({
        key,
        nombreOrden: key.startsWith("__directo__") ? subgrupos[key][0].nombre || "" : key,
        esGrupo: !key.startsWith("__directo__"),
      }));
      entradas.sort((a, b) => comparaNatural(a.nombreOrden, b.nombreOrden));

      for (const entrada of entradas) {
        if (entrada.esGrupo) {
          agregarSubEncabezado(entrada.key, nivelVisual);
          agruparRecursivo(subgrupos[entrada.key], nivelIdxRuta + 1, nivelVisual + 1);
        } else {
          const ordenado = [...subgrupos[entrada.key]].sort((a, b) => comparaNatural(a.nombre, b.nombre));
          ordenado.forEach((c) => agregarFila(c, nivelVisual));
        }
      }
    }

    for (const especialidad of ordenGrupos) {
      agregarSubEncabezado(especialidad.toUpperCase(), 1);
      agruparRecursivo(grupos[especialidad], 2, 2);
    }

    sheet.views = [{ state: "frozen", ySplit: FILA_ENCABEZADO }];
  }

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `Reporte_Consolidado_Global_${new Date().toISOString().slice(0, 10)}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

const CODIGO_PROYECTO = "CHIJNAYA";


// --- Palabras que indican una carpeta "contenedor de tipo de archivo" (PDF,
// Editable, CAD, Excel, etc.) — cuando el último nivel de una carpeta es uno
// de estos, la lista de separadores usa el penúltimo nivel en su lugar,
// porque el nombre del contenedor no aporta información real a un separador.
const PALABRAS_CONTENEDOR = [
  "PDF", "EDITABLE", "EDITABLES", "EXCEL", "PLANILLA", "PLANILLAS",
  "CAD", "DWG", "WORD", "DOC", "DOCX", "RVT", "REVIT",
];

function esCarpetaContenedor(nombre) {
  const upper = (nombre || "").toUpperCase();
  return PALABRAS_CONTENEDOR.some((palabra) => upper.includes(palabra));
}

function comparaNaturalSep(a, b) {
  return (a || "").localeCompare(b || "", undefined, { numeric: true, sensitivity: "base" });
}

// Para cada carpeta, calcula los segmentos de ruta "efectivos" para el árbol
// de separadores: se descarta el nombre del área (primer segmento) y, si el
// último nivel es un contenedor de tipo de archivo (PDF/EDITABLE/CAD/etc.),
// se usa el penúltimo nivel en su lugar.
function rutaEfectivaSeparador(carpeta) {
  const partes = (carpeta.ruta || carpeta.nombre || "").split(" / ").filter(Boolean);
  let segmentos = partes.slice(1); // quita el nombre del área
  if (segmentos.length === 0) return [];
  const ultimo = segmentos[segmentos.length - 1];
  if (segmentos.length > 1 && esCarpetaContenedor(ultimo)) {
    segmentos = segmentos.slice(0, -1);
  }
  return segmentos;
}

// Construye un árbol anidado a partir de las rutas efectivas de todas las
// carpetas de un área. Carpetas distintas que colapsan a la misma ruta
// efectiva (ej. "PDF" y "EDITABLE" bajo la misma carpeta madre) se
// deduplican automáticamente al insertarse en el mismo nodo del árbol.
function construirArbolSeparadores(carpetasDelArea) {
  const raiz = {};
  for (const c of carpetasDelArea) {
    const segmentos = rutaEfectivaSeparador(c);
    let nodo = raiz;
    for (const seg of segmentos) {
      if (!nodo[seg]) nodo[seg] = {};
      nodo = nodo[seg];
    }
  }
  return raiz;
}

// Extrae el número que YA viene en el nombre real de la carpeta en Drive
// (ej. "1.20_FICHA TÉCNICA" → item "1.20", descripción "FICHA TÉCNICA").
// Si el nombre no empieza con un número reconocible, el ítem queda vacío
// y la descripción es el nombre completo tal cual.
function extraerNumeroYDescripcion(nombre) {
  const texto = (nombre || "").trim();
  const match = texto.match(/^(\d+(?:\.\d+)*)[\.\_\-\s]+(.+)$/);
  if (match && match[2].trim()) {
    return { numero: match[1], descripcion: match[2].trim() };
  }
  return { numero: "", descripcion: texto };
}

// Recorre el árbol igual que generarFilasSeparadores, pero para la "Lista
// General": en vez de CONFIAR en el número que ya trae el nombre de cada
// carpeta (y saltarse las que no lo tienen), genera SU PROPIA numeración
// según la posición real de cada carpeta dentro del árbol — así nunca omite
// ninguna, sin importar cómo esté nombrada, y la numeración queda
// garantizada completa y correlativa (1, 1.1, 1.1.1, 1.1.2...).
function generarFilasGeneral(nodo, nivel, codigoPadre, resultado) {
  const claves = Object.keys(nodo).sort(comparaNaturalSep);
  claves.forEach((clave, i) => {
    const { descripcion } = extraerNumeroYDescripcion(clave);
    const codigo = codigoPadre ? `${codigoPadre}.${i + 1}` : String(i + 1);
    const tieneHijos = Object.keys(nodo[clave]).length > 0;
    resultado.push({ numero: codigo, nombre: descripcion || clave, nivel, esUltimoNivel: !tieneHijos });
    generarFilasGeneral(nodo[clave], nivel + 1, codigo, resultado);
  });
}

// Recorre el árbol en el mismo orden en que aparece en Drive (orden natural,
// respetando números). Por cada nodo extrae el número y la descripción
// reales desde su propio nombre.
//
// Reglas:
// - Si una carpeta NO tiene número reconocible en su nombre, se OMITE de la
//   lista (no genera fila) — pero sus hijos SÍ se recorren igual, respetando
//   la jerarquía (el nivel no avanza para ellos, ya que ocupan visualmente
//   el lugar de su padre omitido).
// - Cada fila que sí se agrega guarda su `nivel` (para el degradado de color)
//   y si es `esUltimoNivel` (no tiene hijos) — el último nivel no se sombrea.
function generarFilasSeparadores(nodo, nivel, resultado) {
  const claves = Object.keys(nodo).sort(comparaNaturalSep);
  for (const clave of claves) {
    const { numero, descripcion } = extraerNumeroYDescripcion(clave);
    const tieneHijos = Object.keys(nodo[clave]).length > 0;
    if (numero !== "") {
      resultado.push({ numero, nombre: descripcion, nivel, esUltimoNivel: !tieneHijos });
      generarFilasSeparadores(nodo[clave], nivel + 1, resultado);
    } else {
      // Carpeta sin número: se omite la fila, pero sus hijos se mantienen
      // en el mismo nivel visual (no se "hunde" un nivel por un padre que no existe en la lista).
      generarFilasSeparadores(nodo[clave], nivel, resultado);
    }
  }
}

function mezclarColorConBlanco(hexRgb, factor) {
  const r = parseInt(hexRgb.slice(2, 4), 16);
  const g = parseInt(hexRgb.slice(4, 6), 16);
  const b = parseInt(hexRgb.slice(6, 8), 16);
  const mezcla = (c) => Math.round(c + (255 - c) * factor).toString(16).padStart(2, "0").toUpperCase();
  return `FF${mezcla(r)}${mezcla(g)}${mezcla(b)}`;
}

const COLOR_ENCABEZADO_SEP = NARANJA; // acento del encabezado de la lista de separadores (naranja de Chijnaya)

// Un color base distinto por nivel, y una opacidad que se aclara mientras
// más profundo el nivel: nivel 1 el más oscuro/saturado, nivel 2 más claro
// — para que la jerarquía se note tanto en el color como en el tono.
const COLORES_NIVEL_SEPARADOR = [
  { base: "FF2E8B84", factor: 0.35 }, // nivel 1: azul verdoso, oscuro
  { base: "FFD2691E", factor: 0.50 }, // nivel 2: naranja, intermedio
  { base: "FFB33A3A", factor: 0.68 }, // nivel 3: rojo, más claro (reservado, no se usa por ahora)
];
function colorNivelSeparador(nivel) {
  const idx = Math.min(Math.max(nivel - 1, 0), COLORES_NIVEL_SEPARADOR.length - 1);
  const { base, factor } = COLORES_NIVEL_SEPARADOR[idx];
  return mezclarColorConBlanco(base, factor);
}

/**
 * Genera y descarga un Excel simple de dos columnas (N° / Descripción),
 * pensado para imprimir los separadores físicos de un expediente técnico —
 * NO es un reporte de avance, es un índice numerado plano de las secciones
 * reales del área, en el mismo orden en que están en Drive.
 *
 * @param {string} areaNombre - nombre del área (carpeta madre) a exportar
 * @param {Array} carpetasDelArea - carpetas (ya filtradas a esa área)
 */
export async function generarListaSeparadoresExcel(areaNombre, carpetasDelArea) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Visor Chijnaya";
  workbook.created = new Date();

  const nombreHoja = areaNombre.replace(/[\\/*?:[\]]/g, "").slice(0, 31) || "Separadores";
  const sheet = workbook.addWorksheet(nombreHoja, {
    pageSetup: { orientation: "portrait", fitToPage: true, fitToWidth: 1 },
  });

  sheet.columns = [
    { width: 14 },
    { width: 75 },
  ];

  if (LOGO_PUNO_BASE64) {
    try {
      const base64Data = LOGO_PUNO_BASE64.includes("base64,")
        ? LOGO_PUNO_BASE64.split("base64,")[1]
        : LOGO_PUNO_BASE64;
      const imageId = workbook.addImage({ base64: base64Data, extension: "png" });
      sheet.addImage(imageId, { tl: { col: 0.2, row: 0.2 }, ext: { width: 45, height: 50 } });
    } catch (e) {
      console.error("No se pudo cargar el logo en el Excel:", e);
    }
  }

  sheet.mergeCells("A1:B1");
  sheet.getCell("A1").value = "GOBIERNO REGIONAL DE PUNO — GERENCIA REGIONAL DE INFRAESTRUCTURA";
  sheet.getCell("A1").font = { bold: true, size: 11 };
  sheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(1).height = 18;

  sheet.mergeCells("A2:B2");
  sheet.getCell("A2").value = "SUB GERENCIA DE ESTUDIOS DEFINITIVOS";
  sheet.getCell("A2").font = { bold: true, size: 10 };
  sheet.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(2).height = 16;

  sheet.mergeCells("A3:B3");
  const celdaTitulo = sheet.getCell("A3");
  celdaTitulo.value = `ÍNDICE DE SEPARADORES — ${areaNombre.toUpperCase()}`;
  celdaTitulo.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
  celdaTitulo.alignment = { horizontal: "center", vertical: "middle" };
  celdaTitulo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_ENCABEZADO_SEP } };
  sheet.getRow(3).height = 22;

  sheet.mergeCells("A4:B4");
  sheet.getCell("A4").value = `Generado: ${new Date().toLocaleString("es-PE")}  ·  Lista plana para impresión de separadores físicos — no refleja estado de avance`;
  sheet.getCell("A4").font = { italic: true, size: 8.5, color: { argb: "FF666666" } };
  sheet.getCell("A4").alignment = { horizontal: "center" };
  sheet.getRow(4).height = 14;

  sheet.addRow([]); // fila 5 en blanco

  const FILA_ENCABEZADO = 6;
  const filaEncabezado = sheet.getRow(FILA_ENCABEZADO);
  filaEncabezado.values = ["N°", "DESCRIPCIÓN"];
  filaEncabezado.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_ENCABEZADO_SEP } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = { bottom: { style: "medium", color: { argb: "FF000000" } } };
  });
  filaEncabezado.height = 18;

  // --- Construir y recorrer el árbol jerárquico ---
  const arbol = construirArbolSeparadores(carpetasDelArea);
  const items = [];
  generarFilasSeparadores(arbol, 1, items);

  for (const item of items) {
    const fila = sheet.addRow([item.numero, item.nombre]);
    fila.getCell(1).alignment = { horizontal: "left", vertical: "middle" };
    fila.getCell(1).numFmt = "@"; // forzar texto, para que "1.10" no se lea como número
    fila.getCell(2).alignment = { horizontal: "left", vertical: "middle", wrapText: true };

    // Sombreado por nivel — mismo nivel, mismo color, con opacidad
    // creciente (nivel 1 el más oscuro, nivel 2 más claro). Solo se
    // sombrean los primeros 2 niveles; del nivel 3 en adelante, sin sombra.
    if (item.nivel <= 2) {
      const fondo = colorNivelSeparador(item.nivel);
      fila.eachCell((cell) => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: fondo } };
      });
    }

    // El nivel 1 además va en negrita, para reforzar la jerarquía visual.
    if (item.nivel === 1) {
      fila.getCell(1).font = { bold: true };
      fila.getCell(2).font = { bold: true };
    }

    fila.eachCell((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "FFAAAAAA" } },
        bottom: { style: "thin", color: { argb: "FFAAAAAA" } },
        left: { style: "thin", color: { argb: "FFAAAAAA" } },
        right: { style: "thin", color: { argb: "FFAAAAAA" } },
      };
    });
  }

  sheet.views = [{ state: "frozen", ySplit: FILA_ENCABEZADO }];
  sheet.headerFooter.oddFooter = "&L&I[Sistema Chijnaya] Índice de separadores&R&IPágina &P de &N";

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `Separadores_${CODIGO_PROYECTO}_${areaNombre.replace(/[^a-zA-Z0-9]+/g, "_")}_${new Date()
    .toISOString()
    .slice(0, 10)}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}


/**
 * Genera y descarga un Excel de 4 columnas (N° Orden / N° Separador /
 * Descripción / N° de Archivadores) — el "Índice General de la
 * Documentación" completo, para archivar físicamente el expediente.
 * A diferencia de la lista de separadores, esta SÍ lleva la columna de
 * archivadores (en blanco, para completar a mano) y un correlativo de
 * orden que solo numera las carpetas de nivel 1 (las demás filas quedan
 * en blanco en esa columna).
 *
 * @param {string} areaNombre - nombre del área (carpeta madre) a exportar
 * @param {Array} carpetasDelArea - carpetas (ya filtradas a esa área)
 */
export async function generarListaGeneralExcel(areaNombre, carpetasDelArea) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Visor Chijnaya";
  workbook.created = new Date();

  const nombreHoja = areaNombre.replace(/[\\/*?:[\]]/g, "").slice(0, 31) || "Lista General";
  const sheet = workbook.addWorksheet(nombreHoja, {
    pageSetup: { orientation: "portrait", fitToPage: true, fitToWidth: 1 },
  });

  sheet.columns = [
    { width: 12 },  // N° de Orden
    { width: 15 },  // N° de Separador
    { width: 58 },  // Descripción
    { width: 20 },  // N° de Archivadores (en blanco, para llenar a mano)
  ];

  if (LOGO_PUNO_BASE64) {
    try {
      const base64Data = LOGO_PUNO_BASE64.includes("base64,")
        ? LOGO_PUNO_BASE64.split("base64,")[1]
        : LOGO_PUNO_BASE64;
      const imageId = workbook.addImage({ base64: base64Data, extension: "png" });
      sheet.addImage(imageId, { tl: { col: 0.2, row: 0.2 }, ext: { width: 45, height: 50 } });
    } catch (e) {
      console.error("No se pudo cargar el logo en el Excel:", e);
    }
  }

  sheet.mergeCells("A1:D1");
  sheet.getCell("A1").value = "GOBIERNO REGIONAL DE PUNO — GERENCIA REGIONAL DE INFRAESTRUCTURA";
  sheet.getCell("A1").font = { bold: true, size: 11 };
  sheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(1).height = 18;

  sheet.mergeCells("A2:D2");
  sheet.getCell("A2").value = "SUB GERENCIA DE ESTUDIOS DEFINITIVOS";
  sheet.getCell("A2").font = { bold: true, size: 10 };
  sheet.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(2).height = 16;

  sheet.mergeCells("A3:D3");
  const celdaTitulo = sheet.getCell("A3");
  celdaTitulo.value = `ÍNDICE GENERAL DE LA DOCUMENTACIÓN — ${areaNombre.toUpperCase()}`;
  celdaTitulo.font = { bold: true, size: 12, color: { argb: "FFFFFFFF" } };
  celdaTitulo.alignment = { horizontal: "center", vertical: "middle" };
  celdaTitulo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_ENCABEZADO_SEP } };
  sheet.getRow(3).height = 22;

  sheet.mergeCells("A4:D4");
  sheet.getCell("A4").value = `Generado: ${new Date().toLocaleString("es-PE")}  ·  La columna "N° DE ARCHIVADORES" queda en blanco para completar a mano`;
  sheet.getCell("A4").font = { italic: true, size: 8.5, color: { argb: "FF666666" } };
  sheet.getCell("A4").alignment = { horizontal: "center" };
  sheet.getRow(4).height = 14;

  sheet.addRow([]); // fila 5 en blanco

  const FILA_ENCABEZADO = 6;
  const filaEncabezado = sheet.getRow(FILA_ENCABEZADO);
  filaEncabezado.values = ["N° ORDEN", "N° SEPARADOR", "DESCRIPCIÓN", "N° DE ARCHIVADORES"];
  filaEncabezado.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: COLOR_ENCABEZADO_SEP } };
    cell.alignment = { vertical: "middle", horizontal: "center", wrapText: true };
    cell.border = { bottom: { style: "medium", color: { argb: "FF000000" } } };
  });
  filaEncabezado.height = 28;

  // --- Construir y recorrer el árbol jerárquico ---
  const arbol = construirArbolSeparadores(carpetasDelArea);
  const items = [];
  generarFilasGeneral(arbol, 1, "", items);

  let contadorOrden = 0;
  for (const item of items) {
    // N° ORDEN solo numera las carpetas de nivel 1 (las principales); el
    // resto de las filas quedan en blanco en esa columna.
    const numeroOrden = item.nivel === 1 ? ++contadorOrden : "";
    const fila = sheet.addRow([numeroOrden, item.numero, item.nombre, ""]);
    fila.getCell(1).alignment = { horizontal: "center", vertical: "middle" };
    fila.getCell(2).alignment = { horizontal: "left", vertical: "middle" };
    fila.getCell(2).numFmt = "@"; // forzar texto, para que "1.10" no se lea como número
    fila.getCell(3).alignment = { horizontal: "left", vertical: "middle", wrapText: true };
    fila.getCell(4).alignment = { horizontal: "center", vertical: "middle" };

    // Sombreado por nivel — mismo nivel, mismo color, con opacidad
    // creciente (nivel 1 el más oscuro, nivel 2 más claro). Solo se
    // sombrean los primeros 2 niveles; del nivel 3 en adelante, sin sombra.
    // La columna "N° DE ARCHIVADORES" (4) siempre queda en amarillo claro,
    // para que se note que se llena a mano, sin importar el nivel.
    if (item.nivel <= 2) {
      const fondo = colorNivelSeparador(item.nivel);
      [1, 2, 3].forEach((col) => {
        fila.getCell(col).fill = { type: "pattern", pattern: "solid", fgColor: { argb: fondo } };
      });
    }
    fila.getCell(4).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFFFF2CC" } };

    // El nivel 1 además va en negrita, para reforzar la jerarquía visual.
    if (item.nivel === 1) {
      fila.getCell(1).font = { bold: true };
      fila.getCell(2).font = { bold: true };
      fila.getCell(3).font = { bold: true };
    }

    // Bordes aplicados explícitamente a las 4 columnas (no con eachCell),
    // porque la celda de "N° de Archivadores" tiene valor vacío y a veces
    // ExcelJS no la toma en cuenta al recorrer con eachCell.
    [1, 2, 3, 4].forEach((col) => {
      fila.getCell(col).border = {
        top: { style: "thin", color: { argb: "FFAAAAAA" } },
        bottom: { style: "thin", color: { argb: "FFAAAAAA" } },
        left: { style: "thin", color: { argb: "FFAAAAAA" } },
        right: { style: "thin", color: { argb: "FFAAAAAA" } },
      };
    });
  }

  sheet.views = [{ state: "frozen", ySplit: FILA_ENCABEZADO }];
  sheet.headerFooter.oddFooter = "&L&I[Sistema Chijnaya] Índice general de la documentación&R&IPágina &P de &N";

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `ListaGeneral_${CODIGO_PROYECTO}_${areaNombre.replace(/[^a-zA-Z0-9]+/g, "_")}_${new Date()
    .toISOString()
    .slice(0, 10)}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
