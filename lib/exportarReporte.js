import jsPDF from "jspdf";
import autoTable from "jspdf-autotable";
import { LOGO_PUNO_BASE64 } from "./logoPuno";

const ESTADO_LABEL = {
  completa: "COMPLETA",
  incompleta: "INCOMPLETA",
  vacia: "VACÍA",
};

const ESTADO_RGB = {
  completa: [46, 204, 113],
  incompleta: [243, 156, 18],
  vacia: [231, 76, 60],
};

const NARANJA_HEADER = [230, 126, 34];
const TEAL = [23, 163, 152];
const TEAL_CLARO = [45, 212, 191];
const VERDE = [46, 204, 113];
const NARANJA = [243, 156, 18];
const ROJO = [231, 76, 60];
const AZUL_OSCURO = [20, 30, 60];

export function generarReportePorArea(areaNombre, carpetasDelArea, opts = {}) {
  const {
    logoBase64 = LOGO_PUNO_BASE64,
    proyectoNombre =
      '"MEJORAMIENTO DEL SERVICIO DE ATENCION DE SALUD BASICOS EN CHIJNAYA DISTRITO DE PUCARA DE LA PROVINCIA DE LAMPA DEPARTAMENTO DE PUNO"',
    usuarioFirma = "Sistema Chijnaya",
  } = opts;

  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 14;
  const anchoUtilTabla = pageWidth - marginX * 2; // 182 mm

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
  ordenGrupos.sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));

  const total = carpetasDelArea.length;
  const completas = carpetasDelArea.filter((c) => c.estado === "completa").length;
  const incompletas = carpetasDelArea.filter((c) => c.estado === "incompleta").length;
  const vacias = carpetasDelArea.filter((c) => c.estado === "vacia").length;
  let necesarios = 0;
  let completados = 0;
  for (const c of carpetasDelArea) {
    necesarios += c.archivosNecesarios || 0;
    completados += c.archivosCompletados || 0;
  }
  const pctArchivos = necesarios > 0 ? Math.round((completados / necesarios) * 100) : 0;

  const body = [];
  let contadorFila = 1;

  function comparaNatural(a, b) {
    return (a || "").localeCompare(b || "", undefined, { numeric: true, sensitivity: "base" });
  }

  function mezclarConBlanco(rgb, factor) {
    return rgb.map((c) => Math.round(c + (255 - c) * factor));
  }

  const FACTORES_HOJA = [0.72, 0.78, 0.83, 0.87, 0.90, 0.93, 0.96];
  function colorHoja(nivelVisual) {
    const factor = FACTORES_HOJA[Math.min(nivelVisual - 1, FACTORES_HOJA.length - 1)];
    return mezclarConBlanco(NARANJA_HEADER, factor);
  }

  function empujarFila(c, nivelVisual) {
    const estado = c.estado || "incompleta";
    const nombreMostrado = c.nombre || (c.ruta || "").split(" / ").pop() || "-";
    const sangriaIzq = 2 + Math.max(0, nivelVisual - 1) * 4.5;
    const fondo = colorHoja(nivelVisual);
    body.push([
      { content: String(contadorFila++), styles: { fillColor: fondo, halign: "center" } },
      { content: nombreMostrado, styles: { fillColor: fondo, cellPadding: { top: 2, right: 2, bottom: 2, left: sangriaIzq } } },
      {
        content: ESTADO_LABEL[estado] || estado.toUpperCase(),
        styles: { fillColor: fondo, textColor: ESTADO_RGB[estado] || [100, 100, 100], fontStyle: "bold", halign: "center" },
      },
      { content: c.detalle || "-", styles: { fillColor: fondo } },
    ]);
  }

  function empujarSubEncabezado(nombre, nivel) {
    const factoresPorNivel = [0.40, 0.55, 0.68, 0.78, 0.86, 0.92];
    const factor = factoresPorNivel[Math.min(nivel - 2, factoresPorNivel.length - 1)];
    const color = mezclarConBlanco(NARANJA_HEADER, factor);
    const sangriaIzq = 2 + (nivel - 1) * 4.5;
    
    const textoSub = `>  ${nombre}`;
    const lineasSub = doc.splitTextToSize(textoSub, anchoUtilTabla - sangriaIzq - 4);

    body.push([
      {
        content: lineasSub.join("\n"),
        colSpan: 4,
        styles: {
          fillColor: color,
          textColor: [70, 45, 15],
          fontStyle: "bold",
          fontSize: 9.0,
          halign: "left",
          overflow: "linebreak",
          cellPadding: { top: 2.5, right: 2, bottom: 2.5, left: sangriaIzq },
        },
      },
    ]);
  }

  const NIVEL_MAX_INTERMEDIO = 6;
  function agruparRecursivo(items, nivelIdx, nivelVisual) {
    if (nivelIdx > NIVEL_MAX_INTERMEDIO) {
      const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
      ordenado.forEach((c) => empujarFila(c, nivelVisual));
      return;
    }

    const subgrupos = {};
    let hayNivelMasProfundo = false;
    for (const c of items) {
      const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
      const clave = partes.length >= nivelIdx + 1 ? partes[nivelIdx] : null;
      if (clave) hayNivelMasProfundo = true;
      const key = clave || `__directo__${c.nombre || c.id}`;
      if (!subgrupos[key]) subgrupos[key] = [];
      subgrupos[key].push(c);
    }

    if (!hayNivelMasProfundo) {
      const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
      ordenado.forEach((c) => empujarFila(c, nivelVisual));
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
        empujarSubEncabezado(entrada.key, nivelIdx);
        agruparRecursivo(subgrupos[entrada.key], nivelIdx + 1, nivelVisual + 1);
      } else {
        const ordenado = [...subgrupos[entrada.key]].sort((a, b) => comparaNatural(a.nombre, b.nombre));
        ordenado.forEach((c) => empujarFila(c, nivelVisual));
      }
    }
  }

  for (const especialidad of ordenGrupos) {
    const textoEsp = especialidad.toUpperCase();
    const lineasEsp = doc.splitTextToSize(textoEsp, anchoUtilTabla - 6);

    body.push([
      { 
        content: lineasEsp.join("\n"), 
        colSpan: 4, 
        styles: { 
          fillColor: NARANJA_HEADER, 
          textColor: [255, 255, 255], 
          fontStyle: "bold", 
          fontSize: 10.5, 
          halign: "left", 
          overflow: "linebreak",
          cellPadding: 3 
        } 
      },
    ]);
    agruparRecursivo(grupos[especialidad], 2, 1);
  }

  function dibujarBloqueInstitucional() {
    let y = 12;
    if (logoBase64) {
      const anchoLogo = 13;
      const altoLogo = 14.6;
      doc.addImage(logoBase64, "PNG", marginX, y - 3, anchoLogo, altoLogo);
    }
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("GOBIERNO REGIONAL DE PUNO — GERENCIA REGIONAL DE INFRAESTRUCTURA", pageWidth / 2, y, { align: "center" });
    y += 5;
    doc.text("SUB GERENCIA DE ESTUDIOS DEFINITIVOS", pageWidth / 2, y, { align: "center" });
    y += 6;
    doc.setFontSize(8);
    
    const lineasProyecto = doc.splitTextToSize(proyectoNombre, pageWidth - marginX * 2 - 20);
    doc.text(lineasProyecto, pageWidth / 2, y, { align: "center" });
    
    return y + (lineasProyecto.length * 3.8) + 4;
  }

  function dibujarEncabezado() {
    let y = dibujarBloqueInstitucional();

    doc.setFillColor(...AZUL_OSCURO);
    doc.rect(marginX, y, pageWidth - marginX * 2, 9, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(10.5);
    doc.text(`REPORTE DE AVANCE — ${areaNombre.toUpperCase()}`, pageWidth / 2, y + 6.0, { align: "center" });
    doc.setTextColor(0, 0, 0);
    y += 15;

    const fecha = new Date().toLocaleString("es-PE");
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.text(`Generado: ${fecha}`, marginX, y);
    doc.text(`Total: ${total}  ·  Completas: ${completas}  ·  Pendientes: ${total - completas}`, pageWidth - marginX, y, { align: "right" });
    return y + 8;
  }

  function dibujarBarra(x, y, ancho, alto, pct, colorRgb, colorFondoRgb) {
    const fondo = colorFondoRgb || [20, 37, 41];
    doc.setFillColor(...fondo);
    doc.roundedRect(x, y, ancho, alto, alto / 2, alto / 2, "F");
    const anchoLleno = Math.max(alto, (ancho * Math.min(100, Math.max(0, pct))) / 100);
    doc.setFillColor(...colorRgb);
    doc.roundedRect(x, y, anchoLleno, alto, alto / 2, alto / 2, "F");
  }

  function dibujarTarjeta(x, y, ancho, alto, valor, etiqueta, colorRgb) {
    doc.setDrawColor(...colorRgb);
    doc.setLineWidth(0.6);
    doc.roundedRect(x, y, ancho, alto, 2, 2, "S");
    doc.setFont("helvetica", "bold");
    doc.setFontSize(20);
    doc.setTextColor(...colorRgb);
    doc.text(String(valor), x + ancho / 2, y + alto / 2 - 1, { align: "center" });
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.setTextColor(90, 90, 90);
    doc.text(etiqueta, x + ancho / 2, y + alto - 5, { align: "center" });
    doc.setTextColor(0, 0, 0);
  }

  // PORTADA — resumen gráfico
  let y = dibujarBloqueInstitucional();

  doc.setFillColor(...AZUL_OSCURO);
  doc.rect(marginX, y, pageWidth - marginX * 2, 11, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.text(`RESUMEN DE AVANCE — ${areaNombre.toUpperCase()}`, pageWidth / 2, y + 7.5, { align: "center" });
  doc.setTextColor(0, 0, 0);
  y += 18;

  const fechaGenerado = new Date().toLocaleString("es-PE");
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.5);
  doc.setTextColor(100, 100, 100);
  doc.text(`Generado: ${fechaGenerado}`, pageWidth / 2, y, { align: "center" });
  doc.setTextColor(0, 0, 0);
  y += 10;

  const pctCarpetas = total > 0 ? Math.round((completas / total) * 100) : 0;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("AVANCE POR CARPETAS", marginX, y);
  doc.setFontSize(16);
  doc.setTextColor(...TEAL);
  doc.text(`${pctCarpetas}%`, pageWidth - marginX, y, { align: "right" });
  doc.setTextColor(0, 0, 0);
  y += 4;
  dibujarBarra(marginX, y, pageWidth - marginX * 2, 6, pctCarpetas, TEAL);
  y += 12;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("AVANCE POR ARCHIVOS (más preciso)", marginX, y);
  doc.setFontSize(16);
  doc.setTextColor(...TEAL_CLARO);
  doc.text(`${pctArchivos}%`, pageWidth - marginX, y, { align: "right" });
  doc.setTextColor(0, 0, 0);
  y += 4;
  dibujarBarra(marginX, y, pageWidth - marginX * 2, 6, pctArchivos, TEAL_CLARO);
  y += 16;

  const anchoTarjeta = (pageWidth - marginX * 2 - 3 * 4) / 4;
  const altoTarjeta = 22;
  dibujarTarjeta(marginX, y, anchoTarjeta, altoTarjeta, total, "CARPETAS", TEAL);
  dibujarTarjeta(marginX + (anchoTarjeta + 4) * 1, y, anchoTarjeta, altoTarjeta, completas, "COMPLETAS", VERDE);
  dibujarTarjeta(marginX + (anchoTarjeta + 4) * 2, y, anchoTarjeta, altoTarjeta, incompletas, "INCOMPLETAS", NARANJA);
  dibujarTarjeta(marginX + (anchoTarjeta + 4) * 3, y, anchoTarjeta, altoTarjeta, vacias, "VACÍAS", ROJO);
  y += altoTarjeta + 12;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.text(`AVANCE POR ESPECIALIDAD (${ordenGrupos.length})`, marginX, y);
  y += 7;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  for (const especialidad of ordenGrupos) {
    if (y > pageHeight - 20) {
      doc.addPage();
      y = 18;
    }
    const items = grupos[especialidad];
    let espNecesarios = 0;
    let espCompletados = 0;
    for (const c of items) {
      espNecesarios += c.archivosNecesarios || 0;
      espCompletados += c.archivosCompletados || 0;
    }
    const pctEsp = espNecesarios > 0 ? Math.round((espCompletados / espNecesarios) * 100) : 0;
    const colorEsp = pctEsp >= 100 ? VERDE : pctEsp >= 50 ? TEAL_CLARO : NARANJA;

    const etiqueta = doc.splitTextToSize(especialidad, 70)[0];
    doc.setTextColor(60, 60, 60);
    doc.text(etiqueta, marginX, y + 3);
    doc.setTextColor(0, 0, 0);

    const barraX = marginX + 74;
    const barraAncho = pageWidth - marginX - 22 - barraX;
    dibujarBarra(barraX, y, barraAncho, 4.5, pctEsp, colorEsp);

    doc.setFont("helvetica", "bold");
    doc.setTextColor(...colorEsp);
    doc.text(`${pctEsp}%`, pageWidth - marginX, y + 3.3, { align: "right" });
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "normal");

    y += 8;
  }

  // DETALLE — tabla carpeta por carpeta
  doc.addPage();
  const primeraPaginaTabla = doc.internal.getCurrentPageInfo().pageNumber;
  const startY = dibujarEncabezado();

  autoTable(doc, {
    startY,
    margin: { left: marginX, right: marginX, top: startY },
    head: [["N°", "DESCRIPCIÓN", "ESTADO", "DETALLE"]],
    body,
    styles: { fontSize: 8.5, cellPadding: 2, valign: "middle", overflow: "linebreak" },
    headStyles: { fillColor: AZUL_OSCURO, textColor: 255, fontStyle: "bold" },
    columnStyles: {
      0: { cellWidth: 12, halign: "center" },
      1: { cellWidth: 89 },
      2: { cellWidth: 26, halign: "center" },
      3: { cellWidth: 55 },
    },
    didDrawPage: () => {
      if (doc.internal.getCurrentPageInfo().pageNumber > primeraPaginaTabla) {
        dibujarEncabezado();
      }
    },
  });

  const totalPaginas = doc.internal.getNumberOfPages();
  for (let p = 1; p <= totalPaginas; p++) {
    doc.setPage(p);
    doc.setFontSize(7.5);
    doc.setTextColor(120, 120, 120);
    doc.text(`Generado por: ${usuarioFirma}`, marginX, pageHeight - 8);
    doc.text(`Página ${p} de ${totalPaginas}`, pageWidth - marginX, pageHeight - 8, { align: "right" });
    doc.setTextColor(0, 0, 0);
  }

  const nombreArchivo = `Reporte_${areaNombre.replace(/[^a-zA-Z0-9]+/g, "_")}_${new Date()
    .toISOString()
    .slice(0, 10)}.pdf`;
  doc.save(nombreArchivo);
}

// ============================================================
// REPORTE CONSOLIDADO GLOBAL DE TODO EL PROYECTO
// ============================================================
export function generarReporteConsolidadoGlobal(todasLasCarpetas, opts = {}) {
  const {
    logoBase64 = LOGO_PUNO_BASE64,
    proyectoNombre =
      '"MEJORAMIENTO DEL SERVICIO DE ATENCION DE SALUD BASICOS IN CHIJNAYA DISTRITO DE PUCARA DE LA PROVINCIA DE LAMPA DEPARTAMENTO DE PUNO"',
    usuarioFirma = "Sistema Chijnaya",
  } = opts;

  const doc = new jsPDF({ orientation: "portrait", unit: "mm", format: "a4" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 14;
  const anchoUtilTabla = pageWidth - marginX * 2;

  // Agrupar por área
  const areasMap = {};
  for (const c of todasLasCarpetas) {
    const area = c.area || "Sin área";
    if (!areasMap[area]) areasMap[area] = [];
    areasMap[area].push(c);
  }
  const ordenAreas = Object.keys(areasMap).sort((a, b) => a.localeCompare(b));

  const totalCarpetas = todasLasCarpetas.length;
  const totalCompletas = todasLasCarpetas.filter((c) => c.estado === "completa").length;
  let totalNec = 0;
  let totalComp = 0;
  for (const c of todasLasCarpetas) {
    totalNec += c.archivosNecesarios || 0;
    totalComp += c.archivosCompletados || 0;
  }
  const pctGlobalArchivos = totalNec > 0 ? Math.round((totalComp / totalNec) * 100) : 0;

  const body = [];
  let contadorFila = 1;

  function comparaNatural(a, b) {
    return (a || "").localeCompare(b || "", undefined, { numeric: true, sensitivity: "base" });
  }

  function mezclarConBlanco(rgb, factor) {
    return rgb.map((c) => Math.round(c + (255 - c) * factor));
  }

  const FACTORES_HOJA = [0.72, 0.78, 0.83, 0.87, 0.90, 0.93, 0.96];
  function colorHoja(nivelVisual) {
    const factor = FACTORES_HOJA[Math.min(nivelVisual - 1, FACTORES_HOJA.length - 1)];
    return mezclarConBlanco(NARANJA_HEADER, factor);
  }

  function empujarFila(c, nivelVisual) {
    const estado = c.estado || "incompleta";
    const nombreMostrado = c.nombre || (c.ruta || "").split(" / ").pop() || "-";
    const sangriaIzq = 2 + Math.max(0, nivelVisual - 1) * 4.5;
    const fondo = colorHoja(nivelVisual);
    body.push([
      { content: String(contadorFila++), styles: { fillColor: fondo, halign: "center" } },
      { content: nombreMostrado, styles: { fillColor: fondo, cellPadding: { top: 2, right: 2, bottom: 2, left: sangriaIzq } } },
      {
        content: ESTADO_LABEL[estado] || estado.toUpperCase(),
        styles: { fillColor: fondo, textColor: ESTADO_RGB[estado] || [100, 100, 100], fontStyle: "bold", halign: "center" },
      },
      { content: c.detalle || "-", styles: { fillColor: fondo } },
    ]);
  }

  function empujarSubEncabezado(nombre, nivel) {
    const factoresPorNivel = [0.40, 0.55, 0.68, 0.78, 0.86, 0.92];
    const factor = factoresPorNivel[Math.min(nivel - 2, factoresPorNivel.length - 1)];
    const color = mezclarConBlanco(NARANJA_HEADER, factor);
    const sangriaIzq = 2 + (nivel - 1) * 4.5;
    
    const textoSub = `>  ${nombre}`;
    const lineasSub = doc.splitTextToSize(textoSub, anchoUtilTabla - sangriaIzq - 4);

    body.push([
      {
        content: lineasSub.join("\n"),
        colSpan: 4,
        styles: {
          fillColor: color,
          textColor: [70, 45, 15],
          fontStyle: "bold",
          fontSize: 9.0,
          halign: "left",
          overflow: "linebreak",
          cellPadding: { top: 2.5, right: 2, bottom: 2.5, left: sangriaIzq },
        },
      },
    ]);
  }

  const NIVEL_MAX_INTERMEDIO = 6;
  function agruparRecursivo(items, nivelIdx, nivelVisual) {
    if (nivelIdx > NIVEL_MAX_INTERMEDIO) {
      const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
      ordenado.forEach((c) => empujarFila(c, nivelVisual));
      return;
    }

    const subgrupos = {};
    let hayNivelMasProfundo = false;
    for (const c of items) {
      const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
      const clave = partes.length >= nivelIdx + 1 ? partes[nivelIdx] : null;
      if (clave) hayNivelMasProfundo = true;
      const key = clave || `__directo__${c.nombre || c.id}`;
      if (!subgrupos[key]) subgrupos[key] = [];
      subgrupos[key].push(c);
    }

    if (!hayNivelMasProfundo) {
      const ordenado = [...items].sort((a, b) => comparaNatural(a.nombre, b.nombre));
      ordenado.forEach((c) => empujarFila(c, nivelVisual));
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
        empujarSubEncabezado(entrada.key, nivelIdx);
        agruparRecursivo(subgrupos[entrada.key], nivelIdx + 1, nivelVisual + 1);
      } else {
        const ordenado = [...subgrupos[entrada.key]].sort((a, b) => comparaNatural(a.nombre, b.nombre));
        ordenado.forEach((c) => empujarFila(c, nivelVisual));
      }
    }
  }

  for (const areaNombre of ordenAreas) {
    const itemsArea = areasMap[areaNombre];
    const gruposEsp = {};
    const ordenEsp = [];
    for (const c of itemsArea) {
      const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
      const especialidad = partes.length > 1 ? partes[1] : "(raíz)";
      if (!gruposEsp[especialidad]) {
        gruposEsp[especialidad] = [];
        ordenEsp.push(especialidad);
      }
      gruposEsp[especialidad].push(c);
    }
    ordenEsp.sort((a, b) => comparaNatural(a, b));

    // Cabecera de Área
    const textoArea = `ÁREA: ${areaNombre.toUpperCase()}`;
    const lineasArea = doc.splitTextToSize(textoArea, anchoUtilTabla - 6);
    body.push([
      {
        content: lineasArea.join("\n"),
        colSpan: 4,
        styles: {
          fillColor: AZUL_OSCURO,
          textColor: [255, 255, 255],
          fontStyle: "bold",
          fontSize: 11,
          halign: "left",
          overflow: "linebreak",
          cellPadding: 3.5,
        },
      },
    ]);

    for (const esp of ordenEsp) {
      body.push([
        {
          content: `»  ${esp.toUpperCase()}`,
          colSpan: 4,
          styles: {
            fillColor: NARANJA_HEADER,
            textColor: [255, 255, 255],
            fontStyle: "bold",
            fontSize: 9.5,
            halign: "left",
            overflow: "linebreak",
            cellPadding: 2.5,
          },
        },
      ]);
      agruparRecursivo(gruposEsp[esp], 2, 1);
    }
  }

  function dibujarBloqueInstitucional() {
    let y = 12;
    if (logoBase64) {
      const anchoLogo = 13;
      const altoLogo = 14.6;
      doc.addImage(logoBase64, "PNG", marginX, y - 3, anchoLogo, altoLogo);
    }
    doc.setFont("helvetica", "bold");
    doc.setFontSize(10);
    doc.text("GOBIERNO REGIONAL DE PUNO — GERENCIA REGIONAL DE INFRAESTRUCTURA", pageWidth / 2, y, { align: "center" });
    y += 5;
    doc.text("SUB GERENCIA DE ESTUDIOS DEFINITIVOS", pageWidth / 2, y, { align: "center" });
    y += 6;
    doc.setFontSize(8);
    
    const lineasProyecto = doc.splitTextToSize(proyectoNombre, pageWidth - marginX * 2 - 20);
    doc.text(lineasProyecto, pageWidth / 2, y, { align: "center" });
    
    return y + (lineasProyecto.length * 3.8) + 4;
  }

  function dibujarEncabezadoGlobal() {
    let y = dibujarBloqueInstitucional();

    doc.setFillColor(...AZUL_OSCURO);
    doc.rect(marginX, y, pageWidth - marginX * 2, 9, "F");
    doc.setTextColor(255, 255, 255);
    doc.setFontSize(10.5);
    doc.text("REPORTE CONSOLIDADO GLOBAL — TODO EL PROYECTO", pageWidth / 2, y + 6.0, { align: "center" });
    doc.setTextColor(0, 0, 0);
    y += 15;

    const fecha = new Date().toLocaleString("es-PE");
    doc.setFont("helvetica", "normal");
    doc.setFontSize(8);
    doc.text(`Generado: ${fecha}`, marginX, y);
    doc.text(`Total Carpetas: ${totalCarpetas}  ·  Completas: ${totalCompletas}`, pageWidth - marginX, y, { align: "right" });
    return y + 8;
  }

  // PORTADA GLOBAL
  let y = dibujarBloqueInstitucional();

  doc.setFillColor(...AZUL_OSCURO);
  doc.rect(marginX, y, pageWidth - marginX * 2, 11, "F");
  doc.setTextColor(255, 255, 255);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(12);
  doc.text("CONSOLIDADO GENERAL DE AVANCE DEL PROYECTO", pageWidth / 2, y + 7.5, { align: "center" });
  doc.setTextColor(0, 0, 0);
  y += 18;

  const fechaGenerado = new Date().toLocaleString("es-PE");
  doc.setFont("helvetica", "normal");
  doc.setFontSize(8.5);
  doc.setTextColor(100, 100, 100);
  doc.text(`Generado: ${fechaGenerado}`, pageWidth / 2, y, { align: "center" });
  doc.setTextColor(0, 0, 0);
  y += 12;

  const pctCarpetasGlobal = totalCarpetas > 0 ? Math.round((totalCompletas / totalCarpetas) * 100) : 0;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("AVANCE GLOBAL POR CARPETAS", marginX, y);
  doc.setFontSize(16);
  doc.setTextColor(...TEAL);
  doc.text(`${pctCarpetasGlobal}%`, pageWidth - marginX, y, { align: "right" });
  doc.setTextColor(0, 0, 0);
  y += 6;

  // Barra global
  doc.setFillColor(20, 37, 41);
  doc.roundedRect(marginX, y, pageWidth - marginX * 2, 6, 3, 3, "F");
  const anchoLlenoGlobal = Math.max(6, ((pageWidth - marginX * 2) * Math.min(100, Math.max(0, pctCarpetasGlobal))) / 100);
  doc.setFillColor(...TEAL);
  doc.roundedRect(marginX, y, anchoLlenoGlobal, 6, 3, 3, "F");
  y += 14;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("AVANCE GLOBAL POR ARCHIVOS", marginX, y);
  doc.setFontSize(16);
  doc.setTextColor(...TEAL_CLARO);
  doc.text(`${pctGlobalArchivos}%`, pageWidth - marginX, y, { align: "right" });
  doc.setTextColor(0, 0, 0);
  y += 6;

  doc.setFillColor(20, 37, 41);
  doc.roundedRect(marginX, y, pageWidth - marginX * 2, 6, 3, 3, "F");
  const anchoLlenoArchivos = Math.max(6, ((pageWidth - marginX * 2) * Math.min(100, Math.max(0, pctGlobalArchivos))) / 100);
  doc.setFillColor(...TEAL_CLARO);
  doc.roundedRect(marginX, y, anchoLlenoArchivos, 6, 3, 3, "F");
  y += 18;

  // Resumen por Áreas en portada
  doc.setFont("helvetica", "bold");
  doc.setFontSize(11);
  doc.text(`DESGLOSE POR ÁREAS (${ordenAreas.length})`, marginX, y);
  y += 8;

  doc.setFont("helvetica", "normal");
  doc.setFontSize(8);
  for (const areaNombre of ordenAreas) {
    if (y > pageHeight - 25) {
      doc.addPage();
      y = 18;
    }
    const itemsArea = areasMap[areaNombre];
    let areaNec = 0;
    let areaComp = 0;
    for (const c of itemsArea) {
      areaNec += c.archivosNecesarios || 0;
      areaComp += c.archivosCompletados || 0;
    }
    const pctArea = areaNec > 0 ? Math.round((areaComp / areaNec) * 100) : 0;
    const colorArea = pctArea >= 100 ? VERDE : pctArea >= 50 ? TEAL_CLARO : NARANJA;

    const etiqueta = doc.splitTextToSize(areaNombre, 70)[0];
    doc.setTextColor(60, 60, 60);
    doc.text(etiqueta, marginX, y + 3);
    doc.setTextColor(0, 0, 0);

    const barraX = marginX + 74;
    const barraAncho = pageWidth - marginX - 22 - barraX;
    
    doc.setFillColor(20, 37, 41);
    doc.roundedRect(barraX, y, barraAncho, 4.5, 2.25, 2.25, "F");
    const anchoBarraArea = Math.max(4.5, (barraAncho * Math.min(100, Math.max(0, pctArea))) / 100);
    doc.setFillColor(...colorArea);
    doc.roundedRect(barraX, y, anchoBarraArea, 4.5, 2.25, 2.25, "F");

    doc.setFont("helvetica", "bold");
    doc.setTextColor(...colorArea);
    doc.text(`${pctArea}%`, pageWidth - marginX, y + 3.3, { align: "right" });
    doc.setTextColor(0, 0, 0);
    doc.setFont("helvetica", "normal");

    y += 8;
  }

  // TABLA GLOBAL DE DETALLE
  doc.addPage();
  const primeraPaginaTablaGlobal = doc.internal.getCurrentPageInfo().pageNumber;
  const startYGlobal = dibujarEncabezadoGlobal();

  autoTable(doc, {
    startY: startYGlobal,
    margin: { left: marginX, right: marginX, top: startYGlobal },
    head: [["N°", "DESCRIPCIÓN", "ESTADO", "DETALLE"]],
    body,
    styles: { fontSize: 8.5, cellPadding: 2, valign: "middle", overflow: "linebreak" },
    headStyles: { fillColor: AZUL_OSCURO, textColor: 255, fontStyle: "bold" },
    columnStyles: {
      0: { cellWidth: 12, halign: "center" },
      1: { cellWidth: 89 },
      2: { cellWidth: 26, halign: "center" },
      3: { cellWidth: 55 },
    },
    didDrawPage: () => {
      if (doc.internal.getCurrentPageInfo().pageNumber > primeraPaginaTablaGlobal) {
        dibujarEncabezadoGlobal();
      }
    },
  });

  const totalPaginas = doc.internal.getNumberOfPages();
  for (let p = 1; p <= totalPaginas; p++) {
    doc.setPage(p);
    doc.setFontSize(7.5);
    doc.setTextColor(120, 120, 120);
    doc.text(`Generado por: ${usuarioFirma}`, marginX, pageHeight - 8);
    doc.text(`Página ${p} de ${totalPaginas}`, pageWidth - marginX, pageHeight - 8, { align: "right" });
    doc.setTextColor(0, 0, 0);
  }

  const nombreArchivo = `Reporte_Consolidado_Global_${new Date().toISOString().slice(0, 10)}.pdf`;
  doc.save(nombreArchivo);
}

// ============================================================================
// REGISTRO DE ACTIVIDAD DE UN DÍA — PDF para saber quién entró al Drive,
// a qué hora, en qué carpeta y con qué tipo de archivo.
// ============================================================================
const ACCION_TEXTO_PDF = {
  archivo_subido: "Subió",
  archivo_reemplazado: "Reemplazó",
  archivo_borrado: "Borró",
  archivo_movido: "Movió",
  archivo_copiado: "Copió",
  carpeta_creada: "Creó carpeta",
  carpeta_borrada: "Borró carpeta",
  carpeta_movida: "Movió carpeta",
  carpeta_marcada_completa: "Marcó completa",
  carpeta_marcada_incompleta: "Marcó incompleta",
  carpeta_desmarcada: "Revirtió marca",
};

const ACCION_RGB_PDF = {
  archivo_subido: [30, 140, 80],
  archivo_reemplazado: [200, 110, 10],
  archivo_borrado: [192, 57, 43],
  archivo_movido: [46, 134, 171],
  archivo_copiado: [30, 150, 130],
  carpeta_creada: [23, 163, 152],
  carpeta_borrada: [192, 57, 43],
  carpeta_movida: [46, 134, 171],
};

function tipoDeArchivoPDF(ev) {
  if (!ev.tipo || ev.tipo.startsWith("carpeta_")) return "Carpeta";
  const m = String(ev.item || "").match(/\.([A-Za-z0-9]{1,6})$/);
  return m ? m[1].toUpperCase() : "Archivo";
}

/**
 * @param {string} fechaTexto - ej. "Lunes, 26 de octubre de 2026"
 * @param {string} fechaISO - ej. "2026-10-26" (para el nombre del archivo)
 * @param {Array} eventos - eventos del día ({ tipo, item, ruta, usuario, timestamp })
 */
export function generarReporteActividadDiaPDF(fechaTexto, fechaISO, eventos, opts = {}) {
  const {
    logoBase64 = LOGO_PUNO_BASE64,
    proyectoNombre =
      '"MEJORAMIENTO DEL SERVICIO DE ATENCION DE SALUD BASICOS EN CHIJNAYA DISTRITO DE PUCARA DE LA PROVINCIA DE LAMPA DEPARTAMENTO DE PUNO"',
    usuarioFirma = "Sistema Chijnaya",
  } = opts;

  const doc = new jsPDF({ orientation: "landscape", unit: "mm", format: "a4" });
  const pageWidth = doc.internal.pageSize.getWidth();
  const pageHeight = doc.internal.pageSize.getHeight();
  const marginX = 10;

  const fmtFecha = (d) =>
    d ? d.toLocaleDateString("es-PE", { timeZone: "America/Lima", day: "2-digit", month: "2-digit", year: "numeric" }) : "—";
  const fmtHora = (d) =>
    d ? d.toLocaleTimeString("es-PE", { timeZone: "America/Lima", hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false }) : "—";

  // Normalizar y ordenar de la primera a la última actividad
  const lista = (eventos || [])
    .map((e) => ({
      ...e,
      _fecha: e.timestamp?.toDate ? e.timestamp.toDate() : e.timestamp instanceof Date ? e.timestamp : null,
    }))
    .sort((a, b) => (a._fecha?.getTime() || 0) - (b._fecha?.getTime() || 0));

  // Encabezado institucional
  let y = 12;
  if (logoBase64) doc.addImage(logoBase64, "PNG", marginX, y - 3, 13, 14.6);
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.setTextColor(0, 0, 0);
  doc.text("GOBIERNO REGIONAL DE PUNO — GERENCIA REGIONAL DE INFRAESTRUCTURA", pageWidth / 2, y, { align: "center" });
  y += 5;
  doc.text("SUB GERENCIA DE ESTUDIOS DEFINITIVOS", pageWidth / 2, y, { align: "center" });
  y += 6;
  doc.setFontSize(8);
  const lineasProyecto = doc.splitTextToSize(proyectoNombre, pageWidth - marginX * 2 - 40);
  doc.text(lineasProyecto, pageWidth / 2, y, { align: "center" });
  y += lineasProyecto.length * 3.8 + 4;

  doc.setFont("helvetica", "bold");
  doc.setFontSize(13);
  doc.setTextColor(...TEAL);
  doc.text("REGISTRO DE ACTIVIDAD DEL DRIVE", pageWidth / 2, y + 3, { align: "center" });
  doc.setFontSize(10);
  doc.setTextColor(60, 60, 60);
  doc.text(`${fechaTexto} — ${lista.length} evento${lista.length !== 1 ? "s" : ""}`, pageWidth / 2, y + 9, { align: "center" });
  doc.setTextColor(0, 0, 0);
  y += 14;

  // --- Resumen por usuario ---
  const porUsuario = new Map();
  for (const e of lista) {
    const u = e.usuario || "Desconocido";
    if (!porUsuario.has(u)) {
      porUsuario.set(u, { subio: 0, reemplazo: 0, borro: 0, carpCreadas: 0, carpBorradas: 0, otros: 0, total: 0, primera: null, ultima: null });
    }
    const r = porUsuario.get(u);
    r.total++;
    if (e.tipo === "archivo_subido") r.subio++;
    else if (e.tipo === "archivo_reemplazado") r.reemplazo++;
    else if (e.tipo === "archivo_borrado") r.borro++;
    else if (e.tipo === "carpeta_creada") r.carpCreadas++;
    else if (e.tipo === "carpeta_borrada") r.carpBorradas++;
    else r.otros++;
    if (e._fecha) {
      if (!r.primera || e._fecha < r.primera) r.primera = e._fecha;
      if (!r.ultima || e._fecha > r.ultima) r.ultima = e._fecha;
    }
  }
  const filasResumen = [...porUsuario.entries()]
    .sort((a, b) => b[1].total - a[1].total)
    .map(([u, r]) => [u, fmtHora(r.primera), fmtHora(r.ultima), r.subio, r.reemplazo, r.borro, r.carpCreadas, r.carpBorradas, r.otros, r.total]);

  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.text("Resumen por usuario", marginX, y);
  autoTable(doc, {
    startY: y + 2,
    head: [["Usuario", "Primera hora", "Última hora", "Subió", "Reemplazó", "Borró arch.", "Carp. creadas", "Carp. borradas", "Otros", "Total"]],
    body: filasResumen,
    theme: "grid",
    margin: { left: marginX, right: marginX, bottom: 16 },
    styles: { fontSize: 8, cellPadding: 1.6, lineColor: [190, 190, 190], lineWidth: 0.1 },
    headStyles: { fillColor: AZUL_OSCURO, textColor: 255, halign: "center" },
    columnStyles: {
      0: { cellWidth: 80 },
      1: { halign: "center" }, 2: { halign: "center" },
      3: { halign: "center" }, 4: { halign: "center" }, 5: { halign: "center" },
      6: { halign: "center" }, 7: { halign: "center" }, 8: { halign: "center" },
      9: { halign: "center", fontStyle: "bold" },
    },
  });

  // --- Detalle cronológico ---
  let yDet = doc.lastAutoTable.finalY + 8;
  if (yDet > pageHeight - 40) {
    doc.addPage();
    yDet = 16;
  }
  doc.setFont("helvetica", "bold");
  doc.setFontSize(10);
  doc.setTextColor(0, 0, 0);
  doc.text("Detalle de actividad (de la primera a la última)", marginX, yDet);

  autoTable(doc, {
    startY: yDet + 2,
    head: [["N°", "Fecha", "Hora", "Usuario", "Acción", "Tipo", "Elemento", "Carpeta (ruta)"]],
    body: lista.map((e, i) => [
      i + 1,
      fmtFecha(e._fecha),
      fmtHora(e._fecha),
      e.usuario || "Desconocido",
      ACCION_TEXTO_PDF[e.tipo] || e.tipo || "—",
      tipoDeArchivoPDF(e),
      e.item || "",
      e.desde ? `${e.ruta || ""}  (desde: ${e.desde})` : e.ruta || "",
    ]),
    theme: "grid",
    margin: { left: marginX, right: marginX, bottom: 16 },
    styles: { fontSize: 7, cellPadding: 1.3, lineColor: [200, 200, 200], lineWidth: 0.1, overflow: "linebreak", valign: "middle" },
    headStyles: { fillColor: AZUL_OSCURO, textColor: 255, halign: "center" },
    alternateRowStyles: { fillColor: [244, 248, 248] },
    columnStyles: {
      0: { cellWidth: 11, halign: "center" },
      1: { cellWidth: 18, halign: "center" },
      2: { cellWidth: 15, halign: "center" },
      3: { cellWidth: 62 },
      4: { cellWidth: 24, fontStyle: "bold" },
      5: { cellWidth: 13, halign: "center" },
      6: { cellWidth: 50 },
      7: { cellWidth: "auto" },
    },
    didParseCell: (data) => {
      if (data.section === "body" && data.column.index === 4) {
        const tipo = lista[data.row.index]?.tipo;
        const rgb = ACCION_RGB_PDF[tipo];
        if (rgb) data.cell.styles.textColor = rgb;
      }
    },
  });

  const totalPaginas = doc.internal.getNumberOfPages();
  for (let p = 1; p <= totalPaginas; p++) {
    doc.setPage(p);
    doc.setFont("helvetica", "normal");
    doc.setFontSize(7.5);
    doc.setTextColor(120, 120, 120);
    doc.text(`Generado por: ${usuarioFirma}`, marginX, pageHeight - 8);
    doc.text(`Página ${p} de ${totalPaginas}`, pageWidth - marginX, pageHeight - 8, { align: "right" });
    doc.setTextColor(0, 0, 0);
  }
  doc.save(`Actividad_CHIJNAYA_${fechaISO}.pdf`);
}
