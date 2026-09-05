import ExcelJS from "exceljs";

const AZUL_OSCURO = "FF141E3C";
const NARANJA = "FFE67E22";
const AMARILLO_EDITABLE = "FFFFF2CC";
const BORDE_GRIS = { style: "thin", color: { argb: "FFBBBBBB" } };

// Los nombres de carpeta en Drive ya vienen con su propia numeración manual
// pegada al nombre (ej. "1.1_ÍNDICE GENERAL...", "2.6. RELACION DE INSUMOS") —
// como ya generamos el N° SEPARADOR en su propia columna, quitamos ese prefijo
// del nombre para no repetir el número dos veces.
function limpiarNombre(nombre) {
  return (nombre || "")
    .replace(/^\s*\d+(?:\.\d+)*[\s_.\-:]+\s*/, "")
    .trim();
}

/**
 * Genera y descarga un Excel (.xlsx) con el ÍNDICE GENERAL DE LA DOCUMENTACIÓN,
 * en el mismo formato de 4 columnas que usa la entidad para armar el expediente
 * físico (N° ORDEN | N° SEPARADOR | DESCRIPCIÓN | N° DE ARCHIVADORES), pero
 * generado en vivo a partir de la estructura real de carpetas del Drive —
 * en vez de armarlo a mano.
 *
 * Numeración:
 *  - N° ORDEN: solo en las especialidades de primer nivel (1, 2, 3...).
 *  - N° SEPARADOR: numeración jerárquica tipo "1.1", "1.1.1", "1.1.1.1",
 *    calculada según la profundidad real de cada carpeta dentro de su ruta.
 *  - DESCRIPCIÓN: nombre de la carpeta.
 *  - N° DE ARCHIVADORES: columna en blanco (amarilla) para que el usuario la
 *    complete a mano, ya que el número de archivador físico no existe en Drive.
 *
 * @param {string} areaNombre - nombre del área a exportar
 * @param {Array} carpetasDelArea - carpetas (ya filtradas a esa área)
 */
export async function generarIndiceGeneralExcel(areaNombre, carpetasDelArea) {
  const workbook = new ExcelJS.Workbook();
  workbook.creator = "Visor Chijnaya";
  workbook.created = new Date();

  const nombreHoja = areaNombre.replace(/[\\/*?:[\]]/g, "").slice(0, 31) || "Indice";
  const sheet = workbook.addWorksheet(nombreHoja, {
    pageSetup: { orientation: "portrait", fitToPage: true, fitToWidth: 1 },
  });

  sheet.columns = [
    { width: 10 }, // N° ORDEN
    { width: 12 }, // N° SEPARADOR
    { width: 70 }, // DESCRIPCION
    { width: 18 }, // N° DE ARCHIVADORES
  ];

  // --- Encabezado institucional ---
  sheet.mergeCells("A1:D1");
  sheet.getCell("A1").value = "GOBIERNO REGIONAL DE PUNO — GERENCIA REGIONAL DE INFRAESTRUCTURA";
  sheet.getCell("A1").font = { bold: true, size: 12 };
  sheet.getCell("A1").alignment = { horizontal: "center", vertical: "middle" };
  sheet.getRow(1).height = 20;

  sheet.mergeCells("A2:D2");
  sheet.getCell("A2").value = "SUB GERENCIA DE ESTUDIOS DEFINITIVOS";
  sheet.getCell("A2").font = { bold: true, size: 10 };
  sheet.getCell("A2").alignment = { horizontal: "center", vertical: "middle" };

  sheet.mergeCells("A3:D3");
  const celdaTitulo = sheet.getCell("A3");
  celdaTitulo.value = `ÍNDICE GENERAL DE LA DOCUMENTACIÓN — ${areaNombre.toUpperCase()}`;
  celdaTitulo.font = { bold: true, size: 13, color: { argb: "FFFFFFFF" } };
  celdaTitulo.alignment = { horizontal: "center", vertical: "middle" };
  celdaTitulo.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_OSCURO } };
  sheet.getRow(3).height = 24;

  sheet.mergeCells("A4:D4");
  sheet.getCell("A4").value = `Generado: ${new Date().toLocaleString("es-PE")} — la columna "N° DE ARCHIVADORES" queda en blanco para completar a mano`;
  sheet.getCell("A4").font = { italic: true, size: 9, color: { argb: "FF666666" } };

  sheet.addRow([]); // fila 5 en blanco, de separación

  // --- Encabezado de columnas (fila 6) ---
  const FILA_ENCABEZADO = 6;
  const filaEncabezado = sheet.getRow(FILA_ENCABEZADO);
  filaEncabezado.values = ["N° ORDEN", "N° SEPARADOR", "DESCRIPCIÓN", "N° DE ARCHIVADORES"];
  filaEncabezado.eachCell((cell) => {
    cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
    cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: AZUL_OSCURO } };
    cell.alignment = { vertical: "middle", horizontal: "center" };
    cell.border = { bottom: { style: "medium", color: { argb: "FF000000" } } };
  });
  filaEncabezado.height = 18;

  function comparaNatural(a, b) {
    return (a || "").localeCompare(b || "", undefined, { numeric: true, sensitivity: "base" });
  }

  // --- Agrupar por especialidad (partes[1] de la ruta), igual criterio que los otros exports ---
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

  function agregarFilaSeparador(nombre, codigo) {
    const nombreLimpio = limpiarNombre(nombre);
    const fila = sheet.addRow([null, codigo, nombreLimpio, null]);
    fila.getCell(2).alignment = { horizontal: "right", vertical: "top" };
    fila.getCell(3).alignment = { wrapText: true, vertical: "top", indent: Math.max(0, codigo.split(".").length - 1) };
    fila.getCell(4).fill = { type: "pattern", pattern: "solid", fgColor: { argb: AMARILLO_EDITABLE } };
    fila.getCell(4).border = { top: BORDE_GRIS, bottom: BORDE_GRIS, left: BORDE_GRIS, right: BORDE_GRIS };
    return fila;
  }

  function agregarFilaOrden(nombre, numero, archivadores) {
    const nombreLimpio = limpiarNombre(nombre);
    const fila = sheet.addRow([numero, null, nombreLimpio, archivadores || null]);
    fila.getCell(1).alignment = { horizontal: "center", vertical: "top" };
    fila.getCell(1).font = { bold: true };
    fila.getCell(3).font = { bold: true };
    fila.getCell(3).alignment = { wrapText: true, vertical: "top" };
    fila.eachCell({ includeEmpty: true }, (cell) => {
      cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: NARANJA } };
      cell.font = { ...(cell.font || {}), bold: true, color: { argb: "FFFFFFFF" } };
    });
    fila.getCell(4).border = { top: BORDE_GRIS, bottom: BORDE_GRIS, left: BORDE_GRIS, right: BORDE_GRIS };
    if (!archivadores) fila.getCell(4).alignment = { horizontal: "center" };
    return fila;
  }

  // Agrupación recursiva: arma el código jerárquico "1.1", "1.1.1", "1.1.1.1"...
  // según la profundidad real de la ruta, en el mismo orden natural que ya usan
  // exportarExcel.js / exportarReporte.js para especialidad/subcarpetas.
  function agruparRecursivo(items, nivelIdxRuta, codigoPadre) {
    const subgrupos = {};
    let hayNivelMasProfundo = false;
    for (const c of items) {
      const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
      const clave = partes.length > nivelIdxRuta + 1 ? partes[nivelIdxRuta] : null;
      if (clave) hayNivelMasProfundo = true;
      const key = clave || `__directo__${c.nombre || c.id}`;
      if (!subgrupos[key]) subgrupos[key] = [];
      subgrupos[key].push(c);
    }

    if (!hayNivelMasProfundo) {
      const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
      ordenado.forEach((c, i) => {
        agregarFilaSeparador(c.nombre || (c.ruta || "").split(" / ").pop() || "-", `${codigoPadre}.${i + 1}`);
      });
      return;
    }

    const entradas = Object.keys(subgrupos).map((key) => ({
      key,
      nombreOrden: key.startsWith("__directo__") ? subgrupos[key][0].nombre || "" : key,
      esGrupo: !key.startsWith("__directo__"),
    }));
    entradas.sort((a, b) => comparaNatural(a.nombreOrden, b.nombreOrden));

    entradas.forEach((entrada, i) => {
      const codigo = `${codigoPadre}.${i + 1}`;
      if (entrada.esGrupo) {
        agregarFilaSeparador(entrada.key, codigo);
        agruparRecursivo(subgrupos[entrada.key], nivelIdxRuta + 1, codigo);
      } else {
        const ordenado = [...subgrupos[entrada.key]].sort((a, b) => comparaNatural(a.nombre, b.nombre));
        ordenado.forEach((c) => {
          agregarFilaSeparador(c.nombre, codigo);
        });
      }
    });
  }

  ordenGrupos.forEach((especialidad, i) => {
    const numero = i + 1;
    agregarFilaOrden(especialidad.toUpperCase(), numero, null);
    agruparRecursivo(grupos[especialidad], 2, String(numero));
  });

  sheet.views = [{ state: "frozen", ySplit: FILA_ENCABEZADO }];

  const buffer = await workbook.xlsx.writeBuffer();
  const blob = new Blob([buffer], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `Indice_General_${areaNombre.replace(/[^a-zA-Z0-9]+/g, "_")}_${new Date()
    .toISOString()
    .slice(0, 10)}.xlsx`;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}
