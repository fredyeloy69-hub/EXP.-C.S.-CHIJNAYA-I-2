import { adminDb } from "./firebaseAdmin";
import { scanDriveTree, getUltimoActorDeItem } from "./googleDrive";
import { FieldValue } from "firebase-admin/firestore";

const FOLDER_MIME = "application/vnd.google-apps.folder";
const WORD_EXT = [".docx", ".doc"];
const EXCEL_EXT = [".xlsx", ".xls"];
// DWG = plano editable de AutoCAD (el ORIGINAL). El .bak es solo su respaldo
// automático, NO es un archivo original — no cuenta para pedir su propio PDF.
const DWG_EXT = [".dwg"];
// RVT = modelo 3D de Revit (BIM). Igual que DWG, es un "editable" que necesita su PDF para imprimir.
const RVT_EXT = [".rvt", ".rte", ".rfa"];
// Otros formatos de modelo/cálculo estructural que también son "editables"
// originales — igual regla que DWG/RVT: necesitan al menos 1 PDF propio.
const MODELO_EXT = [".s2k", ".mcdx", ".msh", ".xsdm", ".fdb"];
// Archivos "basura" — se generan solos (logs, temporales, cachés de software
// tipo SAP2000/ETABS/AutoCAD) y NUNCA necesitan su propio PDF. Se reconocen
// como archivos (no quedan sueltos como "otro" sin explicación) pero no
// cuentan para nada en la regla de completitud.
// Incluye .ini/.tmp/.db y Thumbs.db/desktop.ini (config de Windows/carpeta):
// si son lo ÚNICO que hay en una carpeta, esta se trata como VACÍA.
const BASURA_EXT = [".out", ".tlog", ".ico", ".ini", ".tmp", ".temp", ".lock", ".ds_store", ".db"];
const BASURA_PREFIJO = /^k_/i; // ej. "K_NombreDelArchivo..." (temporales de SAP2000/ETABS)
const BASURA_NOMBRE_EXACTO = new Set(["thumbs.db", "desktop.ini", ".ds_store"]);

// Muestra "Nombre (correo@ejemplo.com)" si hay ambos datos disponibles;
// si solo hay uno de los dos, muestra ese; si no hay ninguno, "Desconocido".
function formatearUsuario(nombre, correo) {
  if (nombre && correo) return `${nombre} (${correo})`;
  return nombre || correo || "Desconocido";
}

// Calcula "YYYY-MM-DD" en HORA DE LIMA, sin importar en qué huso horario corra
// el código (Vercel corre en UTC, el navegador del usuario en hora de Lima).
// Antes se usaba new Date().toISOString().slice(0,10), que da la fecha en UTC —
// si algo pasaba después de las 7pm hora Perú, quedaba registrado como "el día
// siguiente" y no coincidía con lo que la persona veía en su calendario.
export function fechaLimaISO(fecha) {
  const partes = new Intl.DateTimeFormat("en-CA", {
    timeZone: "America/Lima",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
  }).formatToParts(fecha);
  const obj = {};
  for (const p of partes) obj[p.type] = p.value;
  return `${obj.year}-${obj.month}-${obj.day}`;
}

function getExt(name) {
  const lower = name.toLowerCase();
  const match = lower.match(/\.[a-z0-9]+$/);
  return match ? match[0] : "";
}

function classifyFile(name) {
  const ext = getExt(name);
  const nombreSinExt = name.slice(0, name.length - ext.length);
  if (BASURA_PREFIJO.test(nombreSinExt) || BASURA_EXT.includes(ext) || BASURA_NOMBRE_EXACTO.has(name.toLowerCase())) return "basura";
  if (ext === ".pdf") return "pdf";
  // .bak es un RESPALDO (de cualquier archivo: .dwg.bak, .evo.bak, etc.), nunca
  // el original — se reconoce como archivo aparte pero jamás pide su propio PDF.
  if (ext === ".bak") return "respaldo";
  if (WORD_EXT.includes(ext)) return "word";
  if (EXCEL_EXT.includes(ext)) return "excel";
  if (DWG_EXT.includes(ext)) return "dwg";
  if (RVT_EXT.includes(ext)) return "rvt";
  if (MODELO_EXT.includes(ext)) return "modelo";
  return "otro";
}

function buildPath(id, nodes, rootId) {
  const parts = [];
  let current = nodes[id];
  while (current && current.id !== rootId) {
    parts.unshift(current.name);
    current = nodes[current.parentId];
  }
  return parts.join(" / ");
}

// Primer segmento de la ruta = área / especialidad (ej. "PROYECTO CONTINGENCIA", "OTROS ARCHIVOS")
function extractArea(ruta) {
  if (!ruta) return "Sin área";
  const primerSegmento = ruta.split(" / ")[0];
  return primerSegmento || "Sin área";
}

// Convierte la lista plana de Drive en un mapa de nodos + arma la jerarquia
function buildTree(items, rootId) {
  const nodes = {};
  for (const item of items) {
    nodes[item.id] = {
      id: item.id,
      name: item.name,
      mimeType: item.mimeType,
      parentId: item.parents ? item.parents[0] : rootId,
      modifiedTime: item.modifiedTime,
      md5Checksum: item.md5Checksum || null,
      webViewLink: item.webViewLink || null,
      thumbnailLink: item.thumbnailLink || null,
      lastModifyingUser: formatearUsuario(
        item.lastModifyingUser?.displayName,
        item.lastModifyingUser?.emailAddress
      ),
      isFolder: item.mimeType === FOLDER_MIME,
      childFolderIds: [],
      files: [],
    };
  }

  for (const id in nodes) {
    const node = nodes[id];
    const parent = nodes[node.parentId];
    if (parent && node.isFolder) parent.childFolderIds.push(id);
  }

  for (const id in nodes) {
    const node = nodes[id];
    if (!node.isFolder) {
      const parent = nodes[node.parentId];
      if (parent) parent.files.push(node);
    }
  }

  return nodes;
}

// Devuelve { estado, detalle, tienePdf, tieneEditable, pdfCount, extensionesEditables }
// estado: completa | incompleta | vacia
//
// REGLA DE COMPLETITUD (cada editable necesita su propio PDF):
// - Cada archivo Word cuenta como 1 "editable" que necesita 1 PDF.
// - Cada archivo Excel cuenta como 1 "editable" que necesita 1 PDF.
// - DWG (+ su .bak) cuenta como UN SOLO editable que necesita al menos 1 PDF
//   (no importa cuántos .dwg/.bak haya, es el mismo plano).
// - RVT cuenta igual que DWG: un solo editable, necesita al menos 1 PDF.
// - PDFs requeridos = (# Word) + (# Excel) + (1 si hay DWG) + (1 si hay RVT).
// - Completa solo si pdfCount >= pdfsRequeridos Y pdfsRequeridos > 0.
function computeEstado(folderNode, override) {
  if (override?.forzada) {
    return {
      estado: "completa",
      detalle: `Marcada manualmente como completa${override.motivo ? ` — ${override.motivo}` : ""}`,
      tienePdf: true,
      tieneEditable: true,
      pdfCount: folderNode.files.filter((f) => classifyFile(f.name) === "pdf").length,
      extensionesEditables: [],
      forzada: true,
      marcadoPor: override.marcadoPor || null,
      motivo: override.motivo || null,
      marcadoEn: override.marcadoEn || null,
      archivosNecesarios: 2,
      archivosCompletados: 2,
    };
  }

  if (folderNode.files.length === 0) {
    return {
      estado: "vacia",
      detalle: "Carpeta vacía — sin archivos",
      tienePdf: false,
      tieneEditable: false,
      pdfCount: 0,
      extensionesEditables: [],
      archivosNecesarios: 2,
      archivosCompletados: 0,
    };
  }

  const pdfFiles = folderNode.files.filter((f) => classifyFile(f.name) === "pdf");
  const wordFiles = folderNode.files.filter((f) => classifyFile(f.name) === "word");
  const excelFiles = folderNode.files.filter((f) => classifyFile(f.name) === "excel");
  const dwgFiles = folderNode.files.filter((f) => classifyFile(f.name) === "dwg");
  const rvtFiles = folderNode.files.filter((f) => classifyFile(f.name) === "rvt");
  const modeloFiles = folderNode.files.filter((f) => classifyFile(f.name) === "modelo");
  const respaldoFiles = folderNode.files.filter((f) => classifyFile(f.name) === "respaldo");
  const basuraFiles = folderNode.files.filter((f) => classifyFile(f.name) === "basura");
  const otrosFiles = folderNode.files.filter((f) => classifyFile(f.name) === "otro");

  const otrosExts = Array.from(
    new Set(otrosFiles.map((f) => getExt(f.name).replace(".", "").toUpperCase()))
  ).filter(Boolean);
  const otrosLabel = otrosExts.length > 0 ? `archivos ${otrosExts.join(", ")}` : "otros archivos";

  const tienePdf = pdfFiles.length > 0;
  const tieneWord = wordFiles.length > 0;
  const tieneExcel = excelFiles.length > 0;
  const tieneDwg = dwgFiles.length > 0;
  const tieneRvt = rvtFiles.length > 0;
  const tieneModelo = modeloFiles.length > 0;
  const tieneEditable = tieneWord || tieneExcel || tieneDwg || tieneRvt || tieneModelo;

  // Cuántos PDFs hacen falta en total: 1 por cada Word, 1 por cada Excel,
  // 1 si hay DWG original (el .bak NO cuenta, es solo su respaldo), 1 si hay RVT,
  // 1 si hay algún otro modelo/cálculo editable (S2K, MCDX, MSH, XSDM, FDB).
  const pdfsRequeridos =
    wordFiles.length + excelFiles.length + (tieneDwg ? 1 : 0) + (tieneRvt ? 1 : 0) + (tieneModelo ? 1 : 0);

  // Cuántas "unidades editables" distintas hay (word cuenta cada archivo, dwg/rvt/modelo cuentan como 1 aunque tengan varios archivos)
  const unidadesEditables =
    wordFiles.length + excelFiles.length + (tieneDwg ? 1 : 0) + (tieneRvt ? 1 : 0) + (tieneModelo ? 1 : 0);

  // --- Medida GRANULAR a nivel de archivo (no de carpeta completa) ---
  // Para que el % de avance se mueva gradualmente en vez de saltar de golpe
  // cuando falta o sobra UN solo archivo. Los editables siempre cuentan como
  // "cumplidos" (si se contaron es porque existen); los PDFs cuentan hasta el
  // máximo que hace falta (no se premia de más por tener PDFs de sobra).
  // Si no hay ningún editable reconocido, se asume el mínimo esperado (1
  // editable + 1 PDF) como vara de medir, para no dejar la carpeta fuera del cálculo.
  const archivosNecesarios = tieneEditable ? unidadesEditables * 2 : 2;
  const archivosCompletados = tieneEditable
    ? unidadesEditables + Math.min(pdfFiles.length, pdfsRequeridos)
    : tienePdf
    ? 1
    : 0;

  // Arma el texto "Word", "Word y Excel", "Word, Excel y DWG", etc. según lo que haya
  const nombresList = [];
  if (tieneWord) nombresList.push(wordFiles.length > 1 ? `${wordFiles.length} Word` : "Word");
  if (tieneExcel) nombresList.push(excelFiles.length > 1 ? `${excelFiles.length} Excel` : "Excel");
  if (tieneDwg) nombresList.push("DWG");
  if (tieneRvt) nombresList.push("RVT");
  if (tieneModelo) nombresList.push(getExt(modeloFiles[0].name).replace(".", "").toUpperCase());
  let nombresEditables = "";
  if (nombresList.length === 1) nombresEditables = nombresList[0];
  else if (nombresList.length > 1) {
    nombresEditables = nombresList.slice(0, -1).join(", ") + " y " + nombresList[nombresList.length - 1];
  }

  const extensionesEditables = Array.from(
    new Set([
      ...(tieneWord ? [getExt(wordFiles[0].name)] : []),
      ...(tieneExcel ? [getExt(excelFiles[0].name)] : []),
      ...(tieneDwg ? [getExt(dwgFiles[0].name)] : []),
      ...(tieneRvt ? [".rvt"] : []),
      ...(tieneModelo ? [getExt(modeloFiles[0].name)] : []),
    ])
  );

  if (tieneEditable && pdfsRequeridos > 0 && pdfFiles.length >= pdfsRequeridos) {
    return {
      estado: "completa",
      detalle: `Completa — ${pdfFiles.length} PDF y ${nombresEditables}`,
      tienePdf,
      tieneEditable,
      pdfCount: pdfFiles.length,
      extensionesEditables,
      archivosNecesarios,
      archivosCompletados,
    };
  }

  if (tienePdf && !tieneEditable) {
    return {
      estado: "incompleta",
      detalle: `Falta Word, Excel, DWG, RVT u otro modelo — solo tiene ${pdfFiles.length} PDF`,
      tienePdf,
      tieneEditable,
      pdfCount: pdfFiles.length,
      extensionesEditables,
      archivosNecesarios,
      archivosCompletados,
    };
  }

  if (tieneEditable) {
    // Tiene editable(s) pero le faltan PDFs (0, o menos de los que hacen falta)
    const faltan = pdfsRequeridos - pdfFiles.length;
    const detalleFaltan =
      pdfFiles.length === 0
        ? `Falta${pdfsRequeridos > 1 ? "n" : ""} ${pdfsRequeridos} PDF${pdfsRequeridos > 1 ? "s" : ""} — solo tiene ${nombresEditables}`
        : `Faltan ${faltan} PDF${faltan > 1 ? "s" : ""} — tiene ${pdfFiles.length} de ${pdfsRequeridos} necesarios (${nombresEditables})`;
    return {
      estado: "incompleta",
      detalle: detalleFaltan,
      tienePdf,
      tieneEditable,
      pdfCount: pdfFiles.length,
      extensionesEditables,
      archivosNecesarios,
      archivosCompletados,
    };
  }

  // hay archivos, pero ninguno es PDF ni editable reconocido (ej. imagenes, ini, archivos basura, etc.)

  // Si TODO lo que hay es basura/respaldo (.ini, .bak, .tmp, thumbs.db, etc.)
  // y ningún archivo real, la carpeta se trata como VACÍA.
  if (otrosFiles.length === 0 && (respaldoFiles.length > 0 || basuraFiles.length > 0)) {
    const partesIgnoradas = [];
    if (basuraFiles.length > 0) partesIgnoradas.push(`${basuraFiles.length} archivo${basuraFiles.length > 1 ? "s" : ""} temporal${basuraFiles.length > 1 ? "es" : ""}/config (.ini, .tmp, etc.)`);
    if (respaldoFiles.length > 0) partesIgnoradas.push(`${respaldoFiles.length} respaldo${respaldoFiles.length > 1 ? "s" : ""} (.bak)`);
    return {
      estado: "vacia",
      detalle: `Carpeta vacía — solo tiene ${partesIgnoradas.join(" y ")}, que no cuenta como contenido real`,
      tienePdf: false,
      tieneEditable: false,
      pdfCount: 0,
      extensionesEditables: [],
      archivosNecesarios: 2,
      archivosCompletados: 0,
    };
  }

  const mensajeOtro = `Falta PDF y Word/Excel/DWG/RVT — solo tiene ${otrosLabel}${
    respaldoFiles.length > 0 ? " + respaldo(s) .bak" : ""
  }${basuraFiles.length > 0 ? " + archivo(s) basura/temporal" : ""}`;
  return {
    estado: "incompleta",
    detalle: mensajeOtro,
    tienePdf,
    tieneEditable,
    pdfCount: 0,
    extensionesEditables,
    archivosNecesarios,
    archivosCompletados,
  };
}

// Algunos usuarios, en vez de mezclar el editable y su PDF dentro de la MISMA
// carpeta final, crean una carpeta padre con DOS carpetas hermanas adentro:
// una para el editable (nombrada algo como "EDITABLE", "DIGITAL") y otra para
// el PDF (nombrada algo como "PDF", "IMPRIMIR", "DOCUMENTOS PDF"). En ese caso,
// ninguna de las dos por separado cumple la regla normal (una tiene editable
// sin PDF, la otra PDF sin editable) — hay que evaluarlas juntas, como si
// fueran una sola carpeta con todos esos archivos adentro.
const PATRON_CARPETA_EDITABLE = /(editable|digital)/i;
const PATRON_CARPETA_PDF = /(pdf|imprimir|documento)/i;

function detectarCombosEditablePdf(nodes) {
  const resultadoPorId = {}; // id de carpeta final -> resultado combinado de computeEstado
  for (const id in nodes) {
    const node = nodes[id];
    if (!node.isFolder) continue;

    const hijosFinales = node.childFolderIds
      .map((cid) => nodes[cid])
      .filter((h) => h && h.isFolder && h.childFolderIds.length === 0);
    if (hijosFinales.length < 2) continue;

    const hijoEditable = hijosFinales.find(
      (h) => PATRON_CARPETA_EDITABLE.test(h.name) && !PATRON_CARPETA_PDF.test(h.name)
    );
    const hijoPdf = hijosFinales.find(
      (h) => PATRON_CARPETA_PDF.test(h.name) && !PATRON_CARPETA_EDITABLE.test(h.name)
    );
    if (!hijoEditable || !hijoPdf || hijoEditable.id === hijoPdf.id) continue;

    const archivosCombinados = [...hijoEditable.files, ...hijoPdf.files];
    const resultado = computeEstado({ files: archivosCombinados }, null);
    resultadoPorId[hijoEditable.id] = { ...resultado, detalle: `${resultado.detalle} (junto con "${hijoPdf.name}")` };
    resultadoPorId[hijoPdf.id] = { ...resultado, detalle: `${resultado.detalle} (junto con "${hijoEditable.name}")` };
  }
  return resultadoPorId;
}

export async function runSync() {
  const rootId = process.env.DRIVE_ROOT_FOLDER_ID;
  const items = await scanDriveTree(rootId);
  const nodes = buildTree(items, rootId);

  // --- Detectar pares de carpetas "editable" + "pdf" separadas ---
  const combosEditablePdf = detectarCombosEditablePdf(nodes);

  // --- Cargar snapshot anterior ---
  const snapRef = adminDb.collection("_meta").doc("snapshot");
  const snapDoc = await snapRef.get();
  const prev = snapDoc.exists ? snapDoc.data().nodes || {} : {};

  // --- Cargar carpetas marcadas manualmente como completas (excepciones) ---
  // Esto vive en una colección aparte para no perderse cada vez que se
  // recalculan las carpetas en cada sync.
  const overridesSnap = await adminDb.collection("carpetasForzadas").get();
  const overrides = {};
  for (const d of overridesSnap.docs) {
    overrides[d.id] = d.data();
  }

  const eventos = [];
  const now = FieldValue.serverTimestamp();

  // --- Detectar carpetas nuevas / movidas ---
  for (const id in nodes) {
    const node = nodes[id];
    if (!node.isFolder) continue;
    const before = prev[id];
    if (!before) {
      eventos.push({
        tipo: "carpeta_creada",
        item: node.name,
        ruta: buildPath(id, nodes, rootId),
        usuario: node.lastModifyingUser,
        timestamp: now,
      });
    } else if (before.parentId !== node.parentId) {
      eventos.push({
        tipo: "carpeta_movida",
        item: node.name,
        ruta: buildPath(id, nodes, rootId),
        usuario: node.lastModifyingUser,
        timestamp: now,
      });
    }
  }

  // --- Detectar carpetas borradas ---
  const idsCarpetasBorradas = [];
  for (const id in prev) {
    if (prev[id].isFolder && !nodes[id]) {
      idsCarpetasBorradas.push(id);
      const actor = await getUltimoActorDeItem(id);
      let usuario = "Desconocido (ver Drive Activity)";
      if (actor) {
        usuario = actor;
      } else if (prev[id].lastModifyingUser && prev[id].lastModifyingUser !== "Desconocido") {
        usuario = `${prev[id].lastModifyingUser} (última persona que la modificó)`;
      }
      eventos.push({
        tipo: "carpeta_borrada",
        item: prev[id].name,
        ruta: prev[id].path || prev[id].name,
        usuario,
        timestamp: now,
      });
    }
  }

  // --- Detectar archivos nuevos / reemplazados ---
  for (const id in nodes) {
    const node = nodes[id];
    if (node.isFolder) continue;
    const before = prev[id];
    if (!before) {
      eventos.push({
        tipo: "archivo_subido",
        item: node.name,
        ruta: buildPath(node.parentId, nodes, rootId),
        usuario: node.lastModifyingUser,
        thumbnailLink: node.thumbnailLink || null,
        timestamp: now,
      });
    } else if (before.md5Checksum && before.md5Checksum !== node.md5Checksum) {
      eventos.push({
        tipo: "archivo_reemplazado",
        item: node.name,
        ruta: buildPath(node.parentId, nodes, rootId),
        usuario: node.lastModifyingUser,
        thumbnailLink: node.thumbnailLink || null,
        timestamp: now,
      });
    }
  }

  // --- Detectar archivos borrados ---
  for (const id in prev) {
    if (!prev[id].isFolder && !nodes[id]) {
      const actor = await getUltimoActorDeItem(id);
      let usuario = "Desconocido (ver Drive Activity)";
      if (actor) {
        usuario = actor;
      } else if (prev[id].lastModifyingUser && prev[id].lastModifyingUser !== "Desconocido") {
        usuario = `${prev[id].lastModifyingUser} (última persona que lo subió)`;
      }
      eventos.push({
        tipo: "archivo_borrado",
        item: prev[id].name,
        ruta: prev[id].path || prev[id].name,
        usuario,
        timestamp: now,
      });
    }
  }

  // --- Escribir eventos en el log de auditoria ---
  const batch = adminDb.batch();
  for (const evento of eventos) {
    const ref = adminDb.collection("eventos").doc();
    batch.set(ref, evento);
  }

  // --- Eliminar del visualizador cualquier carpeta que ya no exista en Drive ---
  // IMPORTANTE: esto NO se basa solo en comparar contra el snapshot anterior
  // (eso falla si la carpeta ya se había borrado en un sync viejo, antes de
  // que existiera esta limpieza — quedaba huérfana para siempre). En su lugar,
  // se compara la colección "carpetas" completa contra las carpetas finales
  // que existen AHORA MISMO en Drive, y se borra cualquier doc que sobre.
  const idsFinalesActuales = new Set();
  for (const id in nodes) {
    const node = nodes[id];
    if (node.isFolder && node.childFolderIds.length === 0) {
      idsFinalesActuales.add(id);
    }
  }
  const carpetasExistentesSnap = await adminDb.collection("carpetas").select().get();
  let huerfanasBorradas = 0;
  for (const doc of carpetasExistentesSnap.docs) {
    if (!idsFinalesActuales.has(doc.id)) {
      batch.delete(doc.ref);
      huerfanasBorradas++;
    }
  }

  // --- Escribir estado de cada carpeta final (hoja) ---
  let completas = 0,
    incompletas = 0,
    vacias = 0,
    totalFinales = 0,
    totalArchivosNecesarios = 0,
    totalArchivosCompletados = 0;

  for (const id in nodes) {
    const node = nodes[id];
    if (!node.isFolder) continue;
    const esFinal = node.childFolderIds.length === 0;
    if (!esFinal) continue;

    totalFinales++;
    const ruta = buildPath(id, nodes, rootId);
    // Prioridad: marca manual (override) > combo de carpetas editable+pdf hermanas > cálculo individual normal
    const resultadoEstado = overrides[id]?.forzada
      ? computeEstado(node, overrides[id])
      : combosEditablePdf[id] || computeEstado(node, null);
    const {
      estado,
      detalle,
      tienePdf,
      tieneEditable,
      pdfCount,
      extensionesEditables,
      forzada,
      marcadoPor,
      motivo,
      marcadoEn,
      archivosNecesarios,
      archivosCompletados,
    } = resultadoEstado;
    const area = extractArea(ruta);

    if (estado === "completa") completas++;
    if (estado === "incompleta") incompletas++;
    if (estado === "vacia") vacias++;
    totalArchivosNecesarios += archivosNecesarios || 0;
    totalArchivosCompletados += archivosCompletados || 0;

    const ref = adminDb.collection("carpetas").doc(id);
    batch.set(ref, {
      nombre: node.name,
      ruta,
      area,
      estado,
      detalle,
      tienePdf,
      tieneEditable,
      pdfCount,
      extensionesEditables,
      forzada: !!forzada,
      marcadoPor: marcadoPor || null,
      motivo: motivo || null,
      marcadoEn: marcadoEn || null,
      archivosNecesarios: archivosNecesarios || 0,
      archivosCompletados: archivosCompletados || 0,
      archivos: node.files.map((f) => ({
        nombre: f.name,
        tipo: classifyFile(f.name),
        modificadoPor: f.lastModifyingUser,
        modificadoEn: f.modifiedTime,
        link: f.webViewLink,
      })),
      actualizadoEn: now,
    });
  }

  const pctArchivos =
    totalArchivosNecesarios > 0 ? Math.round((totalArchivosCompletados / totalArchivosNecesarios) * 100) : 0;

  batch.set(adminDb.collection("_meta").doc("resumen"), {
    totalFinales,
    completas,
    incompletas,
    vacias,
    faltantes: incompletas + vacias,
    totalArchivosNecesarios,
    totalArchivosCompletados,
    pctArchivos,
    ultimaSync: now,
  });

  // --- Guardar punto de historial para el grafico de tendencia ---
  // Un documento por dia (id = fecha YYYY-MM-DD), se sobreescribe si hay
  // varios syncs el mismo dia, para no acumular puntos de mas.
  const fechaHoy = fechaLimaISO(new Date());
  const pctHoy = totalFinales > 0 ? Math.round((completas / totalFinales) * 100) : 0;
  batch.set(adminDb.collection("historial").doc(fechaHoy), {
    fecha: fechaHoy,
    totalFinales,
    completas,
    incompletas,
    vacias,
    pct: pctHoy,
    timestamp: now,
  });

  // --- Guardar contador AGREGADO de actividad por dia (para el mapa de calor) ---
  // IMPORTANTE: esto existe para evitar que el dashboard tenga que leer cientos
  // de documentos individuales de "eventos" cada vez que alguien abre la pagina
  // (eso fue lo que agoto la cuota gratuita de lecturas de Firestore). En vez de
  // eso, se guarda UN SOLO documento con un contador por dia y por tipo de evento,
  // y el dashboard lee ese unico documento sin importar cuantos eventos haya.
  const conteoPorTipoHoy = {};
  for (const evento of eventos) {
    conteoPorTipoHoy[evento.tipo] = (conteoPorTipoHoy[evento.tipo] || 0) + 1;
  }
  if (Object.keys(conteoPorTipoHoy).length > 0) {
    // OJO: tiene que ser un objeto REALMENTE anidado ({ "2026-08-21": { subido: N } }),
    // no una clave con punto en el nombre ({ "2026-08-21.subido": N }) — esto último
    // creaba un campo suelto llamado literalmente "2026-08-21.subido" en vez de
    // guardar "subido" DENTRO de "2026-08-21", y el dashboard nunca lo encontraba.
    const actividadUpdate = { [fechaHoy]: {} };
    for (const [tipo, cantidad] of Object.entries(conteoPorTipoHoy)) {
      actividadUpdate[fechaHoy][tipo] = FieldValue.increment(cantidad);
    }
    batch.set(adminDb.collection("_meta").doc("actividadPorDia"), actividadUpdate, { merge: true });
  }

  // --- Guardar snapshot para la proxima comparacion ---
  const snapshotToSave = {};
  for (const id in nodes) {
    snapshotToSave[id] = {
      name: nodes[id].name,
      isFolder: nodes[id].isFolder,
      parentId: nodes[id].parentId,
      md5Checksum: nodes[id].md5Checksum,
      lastModifyingUser: nodes[id].lastModifyingUser,
      path: buildPath(nodes[id].isFolder ? id : nodes[id].parentId, nodes, rootId),
    };
  }
  batch.set(snapRef, { nodes: snapshotToSave, updatedAt: now });

  await batch.commit();

  return { eventos: eventos.length, totalFinales, completas, incompletas, vacias };
}
