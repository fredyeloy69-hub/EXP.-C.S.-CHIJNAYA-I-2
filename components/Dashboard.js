"use client";

import { useEffect, useState, useRef } from "react";
import { db } from "@/lib/firebaseClient";
import {
  collection,
  doc,
  onSnapshot,
  orderBy,
  query,
  limit,
  where,
  getDocs,
} from "firebase/firestore";
import { generarReportePorArea, generarReporteConsolidadoGlobal } from "@/lib/exportarReporte";
import { generarReporteExcelPorArea, generarListaSeparadoresExcel, generarListaGeneralExcel } from "@/lib/exportarExcel";
import { LOGO_PUNO_BASE64 } from "@/lib/logoPuno";
import {
  getAuth,
  GoogleAuthProvider,
  signInWithPopup,
  onAuthStateChanged,
} from "firebase/auth";

const auth = getAuth(db.app);

const COLLAPSE_STORAGE_KEY = "chijnaya_grupos_colapsados";

const ESTADO_COLOR = {
  completa: "#2ecc71",
  incompleta: "#f39c12",
  vacia: "#e74c3c",
};

const AREA_COLORS = [
  "#17a398",
  "#2e86ab",
  "#2dd4bf",
  "#17a398",
  "#2ecc71",
  "#155e75",
  "#0d9488",
  "#9db3b0",
];

function colorForArea(area) {
  let hash = 0;
  for (let i = 0; i < area.length; i++) hash = area.charCodeAt(i) + ((hash << 5) - hash);
  return AREA_COLORS[Math.abs(hash) % AREA_COLORS.length];
}

const ESTADO_OPTIONS = [
  { value: "pendientes", label: "Pendientes", color: "#17a398" },
  { value: "incompleta", label: "Incompletas", color: "#f39c12" },
  { value: "vacia", label: "Vacías", color: "#e74c3c" },
  { value: "completa", label: "Completas", color: "#2ecc71" },
  { value: "todas", label: "Todas", color: "#b7c9c6" },
];

const EVENTO_LABEL = {
  archivo_subido: "subió",
  archivo_reemplazado: "reemplazó",
  archivo_borrado: "borró",
  carpeta_creada: "creó la carpeta",
  carpeta_borrada: "borró la carpeta",
  carpeta_movida: "movió la carpeta",
  carpeta_marcada_completa: "marcó como completa",
  carpeta_desmarcada: "desmarcó",
};

const EVENTO_COLOR = {
  archivo_subido: "#2ecc71",
  archivo_reemplazado: "#f39c12",
  archivo_borrado: "#e74c3c",
  carpeta_creada: "#17a398",
  carpeta_borrada: "#e74c3c",
  carpeta_movida: "#2e86ab",
  carpeta_marcada_completa: "#2dd4bf",
  carpeta_desmarcada: "#f39c12",
};

const EVENTO_ICONO = {
  archivo_subido: "↑",
  archivo_reemplazado: "⟲",
  archivo_borrado: "✕",
  carpeta_creada: "+",
  carpeta_borrada: "✕",
  carpeta_movida: "⇄",
  carpeta_marcada_completa: "✓",
  carpeta_desmarcada: "↺",
};

function fechaLimaISO(fecha) {
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

function formatearFechaLarga(fechaEntrada) {
  const fecha = typeof fechaEntrada === "string" ? new Date(fechaEntrada + "T12:00:00") : fechaEntrada;
  if (!fecha || isNaN(fecha.getTime())) return String(fechaEntrada);
  const texto = fecha.toLocaleDateString("es-PE", { weekday: "long", day: "numeric", month: "long", year: "numeric" });
  return texto.charAt(0).toUpperCase() + texto.slice(1);
}

function tiempoRelativo(date) {
  if (!date) return "...";
  const diffMs = Date.now() - date.getTime();
  const diffMin = Math.floor(diffMs / 60000);
  if (diffMin < 1) return "justo ahora";
  if (diffMin < 60) return `hace ${diffMin} min`;
  const diffH = Math.floor(diffMin / 60);
  if (diffH < 24) return `hace ${diffH} h`;
  const diffD = Math.floor(diffH / 24);
  if (diffD < 7) return `hace ${diffD} d`;
  return date.toLocaleDateString("es-PE");
}

export default function Dashboard() {
  const [resumen, setResumen] = useState(null);
  const [carpetas, setCarpetas] = useState([]);
  const [eventos, setEventos] = useState([]);
  const [filtroArea, setFiltroArea] = useState("Todas");
  const [filtroEstado, setFiltroEstado] = useState("pendientes");
  const [sincronizando, setSincronizando] = useState(false);
  const [mensajeSync, setMensajeSync] = useState(null);
  const [busqueda, setBusqueda] = useState("");
  const [colapsados, setColapsados] = useState({});
  const [colapsoListo, setColapsoListo] = useState(false);
  const [exportandoArea, setExportandoArea] = useState(null);
  const [exportandoExcelArea, setExportandoExcelArea] = useState(null);
  const [exportandoSeparadores, setExportandoSeparadores] = useState(null);
  const [exportandoListaGeneral, setExportandoListaGeneral] = useState(null);
  const [exportandoGlobal, setExportandoGlobal] = useState(false);
  const [modoPresentacion, setModoPresentacion] = useState(false);
  const [historial, setHistorial] = useState([]);
  const [actividadPorDia, setActividadPorDia] = useState({});
  const [marcandoId, setMarcandoId] = useState(null);
  const [mostrarMarcadas, setMostrarMarcadas] = useState(false);
  const [usuarioGoogle, setUsuarioGoogle] = useState(null);
  const [rangoDiasHeatmap, setRangoDiasHeatmap] = useState(84);

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (user) => {
      setUsuarioGoogle(user ? { email: user.email, displayName: user.displayName } : null);
    });
    return () => unsub();
  }, []);

  useEffect(() => {
    try {
      const guardado = localStorage.getItem(COLLAPSE_STORAGE_KEY);
      if (guardado) setColapsados(JSON.parse(guardado));
    } catch {}
    setColapsoListo(true);
  }, []);

  useEffect(() => {
    if (!colapsoListo) return;
    try {
      localStorage.setItem(COLLAPSE_STORAGE_KEY, JSON.stringify(colapsados));
    } catch {}
  }, [colapsados, colapsoListo]);

  function toggleGrupo(key) {
    setColapsados((prev) => ({ ...prev, [key]: !prev[key] }));
  }

  function handleExportarArea(areaNombre, carpetasDelArea) {
    setExportandoArea(areaNombre);
    try {
      const usuarioFirma = usuarioGoogle?.email || usuarioGoogle?.displayName || "Sistema Chijnaya";
      generarReportePorArea(areaNombre, carpetasDelArea, { usuarioFirma });
    } finally {
      setExportandoArea(null);
    }
  }

  async function handleExportarGlobal() {
    setExportandoGlobal(true);
    try {
      const usuarioFirma = usuarioGoogle?.email || usuarioGoogle?.displayName || "Sistema Chijnaya";
      generarReporteConsolidadoGlobal(carpetas, { usuarioFirma });
    } catch (err) {
      alert(`No se pudo generar el reporte consolidado: ${err.message}`);
    } finally {
      setExportandoGlobal(false);
    }
  }

  async function handleExportarExcelArea(areaNombre, carpetasDelArea) {
    setExportandoExcelArea(areaNombre);
    try {
      const timeoutPromise = new Promise((_, reject) =>
        setTimeout(() => reject(new Error("Tiempo de espera agotado al generar el Excel")), 10000)
      );
      await Promise.race([
        generarReporteExcelPorArea(areaNombre, carpetasDelArea),
        timeoutPromise,
      ]);
    } catch (err) {
      alert(`No se pudo generar el Excel: ${err.message}`);
    } finally {
      setExportandoExcelArea(null);
    }
  }

  // Lista de separadores: siempre usa TODAS las carpetas del área, sin
  // filtro por estado — es un índice de secciones, no un reporte de avance.
  async function handleExportarSeparadores(areaNombre, carpetasDelArea) {
    setExportandoSeparadores(areaNombre);
    try {
      const timeoutPromise = new Promise((_, reject) =>
        setTimeout(() => reject(new Error("Tiempo de espera agotado al generar el Excel")), 10000)
      );
      await Promise.race([
        generarListaSeparadoresExcel(areaNombre, carpetasDelArea),
        timeoutPromise,
      ]);
    } catch (err) {
      alert(`No se pudo generar la lista de separadores: ${err.message}`);
    } finally {
      setExportandoSeparadores(null);
    }
  }

  // Lista General (índice completo con columna N° DE ARCHIVADORES): también
  // usa TODAS las carpetas del área, sin filtro por estado.
  async function handleExportarListaGeneral(areaNombre, carpetasDelArea) {
    setExportandoListaGeneral(areaNombre);
    try {
      const timeoutPromise = new Promise((_, reject) =>
        setTimeout(() => reject(new Error("Tiempo de espera agotado al generar el Excel")), 10000)
      );
      await Promise.race([
        generarListaGeneralExcel(areaNombre, carpetasDelArea),
        timeoutPromise,
      ]);
    } catch (err) {
      alert(`No se pudo generar la lista general: ${err.message}`);
    } finally {
      setExportandoListaGeneral(null);
    }
  }

  async function handleMarcarCompleta(folderId, forzada, folderName, folderRuta) {
    let user = auth.currentUser;
    if (!user) {
      try {
        const cred = await signInWithPopup(auth, new GoogleAuthProvider());
        user = cred.user;
      } catch (err) {
        alert(`Necesitas iniciar sesión con Google para marcar/desmarcar carpetas. ${err.message || ""}`);
        return;
      }
    }

    let motivo = "";
    if (forzada) {
      motivo = window.prompt("¿Por qué se marca como completa?", "");
      if (motivo === null) return;
    } else {
      motivo = window.prompt("¿Por qué se desmarca?", "");
      if (motivo === null) return;
    }

    setMarcandoId(folderId);
    try {
      const idToken = await user.getIdToken();
      const res = await fetch("/api/marcar-completa", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ folderId, forzada, motivo, idToken, folderName, folderRuta }),
      });
      if (!res.ok) {
        const data = await res.json().catch(() => ({}));
        alert(`No se pudo actualizar: ${data.error || res.statusText}`);
      }
    } catch (err) {
      alert(`Error de conexión: ${err.message}`);
    } finally {
      setMarcandoId(null);
    }
  }

  async function handleSync() {
    setSincronizando(true);
    setMensajeSync(null);
    try {
      const res = await fetch("/api/manual-sync", { method: "POST" });
      const data = await res.json();
      if (data.ok) {
        setMensajeSync({ tipo: "ok", texto: `Listo — ${data.eventos} eventos nuevos detectados` });
      } else {
        setMensajeSync({ tipo: "error", texto: `Error: ${data.error || "desconocido"}` });
      }
    } catch (err) {
      setMensajeSync({ tipo: "error", texto: "Error de conexión al sincronizar" });
    } finally {
      setSincronizando(false);
    }
  }

  useEffect(() => {
    const unsubResumen = onSnapshot(doc(db, "_meta", "resumen"), (snap) => {
      if (snap.exists()) setResumen(snap.data());
    });

    const unsubCarpetas = onSnapshot(collection(db, "carpetas"), (snap) => {
      setCarpetas(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    });

    const eventosQuery = query(collection(db, "eventos"), orderBy("timestamp", "desc"), limit(50));
    const unsubEventos = onSnapshot(eventosQuery, (snap) => {
      setEventos(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    });

    const unsubActividadPorDia = onSnapshot(doc(db, "_meta", "actividadPorDia"), (snap) => {
      setActividadPorDia(snap.exists() ? snap.data() : {});
    });

    const historialQuery = query(collection(db, "historial"), orderBy("fecha", "asc"), limit(90));
    const unsubHistorial = onSnapshot(historialQuery, (snap) => {
      setHistorial(snap.docs.map((d) => ({ id: d.id, ...d.data() })));
    });

    return () => {
      unsubResumen();
      unsubCarpetas();
      unsubEventos();
      unsubActividadPorDia();
      unsubHistorial();
    };
  }, []);

  const pct = resumen && resumen.totalFinales
    ? Math.round((resumen.completas / resumen.totalFinales) * 100)
    : 0;

  const areas = Array.from(new Set(carpetas.map((c) => c.area || "Sin área"))).sort();

  const carpetasPorArea = {};
  for (const c of carpetas) {
    const a = c.area || "Sin área";
    if (!carpetasPorArea[a]) carpetasPorArea[a] = [];
    carpetasPorArea[a].push(c);
  }

  const carpetasForzadas = carpetas
    .filter((c) => c.forzada)
    .sort((a, b) => new Date(b.marcadoEn || 0) - new Date(a.marcadoEn || 0));

  const areaStats = {};
  const especialidadPorArea = {};
  for (const c of carpetas) {
    const a = c.area || "Sin área";
    if (!areaStats[a])
      areaStats[a] = { total: 0, completas: 0, incompletas: 0, vacias: 0, archivosNecesarios: 0, archivosCompletados: 0 };
    areaStats[a].total++;
    if (c.estado === "completa") areaStats[a].completas++;
    if (c.estado === "incompleta") areaStats[a].incompletas++;
    if (c.estado === "vacia") areaStats[a].vacias++;
    areaStats[a].archivosNecesarios += c.archivosNecesarios || 0;
    areaStats[a].archivosCompletados += c.archivosCompletados || 0;

    const partesRuta = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
    const especialidad = partesRuta.length > 1 ? partesRuta[1] : "(raíz)";
    if (!especialidadPorArea[a]) especialidadPorArea[a] = {};
    if (!especialidadPorArea[a][especialidad])
      especialidadPorArea[a][especialidad] = { total: 0, completas: 0, incompletas: 0, vacias: 0, archivosNecesarios: 0, archivosCompletados: 0 };
    
    especialidadPorArea[a][especialidad].total++;
    if (c.estado === "completa") especialidadPorArea[a][especialidad].completas++;
    if (c.estado === "incompleta") especialidadPorArea[a][especialidad].incompletas++;
    if (c.estado === "vacia") especialidadPorArea[a][especialidad].vacias++;
    especialidadPorArea[a][especialidad].archivosNecesarios += c.archivosNecesarios || 0;
    especialidadPorArea[a][especialidad].archivosCompletados += c.archivosCompletados || 0;
  }

  let listaBase = carpetas;
  if (filtroEstado === "pendientes") {
    listaBase = carpetas.filter((c) => c.estado !== "completa");
  } else if (filtroEstado !== "todas") {
    listaBase = carpetas.filter((c) => c.estado === filtroEstado);
  }

  let visibles = listaBase.sort((a, b) =>
    (a.ruta || "").localeCompare(b.ruta || "", undefined, { numeric: true, sensitivity: "base" })
  );

  if (filtroArea !== "Todas") {
    visibles = visibles.filter((c) => (c.area || "Sin área") === filtroArea);
  }

  if (busqueda.trim()) {
    const q = busqueda.trim().toLowerCase();
    visibles = visibles.filter((c) =>
      (c.nombre || "").toLowerCase().includes(q) || (c.ruta || "").toLowerCase().includes(q)
    );
  }

  const ESTADO_FILTRO_LABEL = {
    pendientes: "Pendientes (incompletas + vacías)",
    incompleta: "Solo incompletas",
    vacia: "Solo vacías",
    completa: "Solo completas",
    todas: "Todas las carpetas",
  };
  const areaLabel = filtroArea !== "Todas" ? ` · ${filtroArea}` : "";

  return (
    <div className="chijnaya-fondo-animado" style={{ minHeight: "100vh", width: "100%" }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

        .chijnaya-fondo-animado, .chijnaya-fondo-animado * {
          font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        }
        .chijnaya-fondo-animado {
          background: linear-gradient(
            -45deg,
            #010708,
            #0a2c33,
            #145e6a,
            #1c7d87,
            #145e6a,
            #0a2c33,
            #010708
          );
          background-size: 500% 500%;
          animation: chijnayaGradiente 8s ease infinite;
        }
        @keyframes chijnayaGradiente {
          0%   { background-position: 0% 50%; }
          50%  { background-position: 100% 50%; }
          100% { background-position: 0% 50%; }
        }
        @media (prefers-reduced-motion: reduce) {
          .chijnaya-fondo-animado { animation: none; }
        }
        .chijnaya-header-sticky {
          position: sticky;
          top: 0;
          z-index: 40;
          backdrop-filter: blur(10px);
          background: rgba(3,16,18,.88);
          border-bottom: 1px solid #1f4a4a;
        }
        .chijnaya-fade-in {
          animation: chijnayaFadeIn .28s ease both;
        }
        @keyframes chijnayaFadeIn {
          from { opacity: 0; transform: translateY(6px); }
          to   { opacity: 1; transform: translateY(0); }
        }
        @media (prefers-reduced-motion: reduce) {
          .chijnaya-fade-in { animation: none; }
        }
        .chijnaya-barra-avance {
          animation: chijnayaRayas 0.8s linear infinite;
        }
        @keyframes chijnayaRayas {
          from { background-position: 0 0, 0 0; }
          to   { background-position: 36px 0, 0 0; }
        }
        @media (prefers-reduced-motion: reduce) {
          .chijnaya-barra-avance { animation: none; }
        }
        .chijnaya-fondo-animado button:not(:disabled) {
          transition: transform .15s ease, filter .15s ease, box-shadow .15s ease;
        }
        .chijnaya-fondo-animado button:not(:disabled):hover {
          transform: translateY(-1.5px) scale(1.015);
          filter: brightness(1.12);
        }
        .chijnaya-fondo-animado button:not(:disabled):active {
          transform: translateY(0) scale(0.98);
          filter: brightness(0.96);
        }
        .chijnaya-tarjeta-viva {
          animation: chijnayaTarjetaEntrada .5s cubic-bezier(.25,.9,.35,1.25) both;
        }
        @keyframes chijnayaTarjetaEntrada {
          from { opacity: 0; transform: translateY(10px) scale(.96); }
          to   { opacity: 1; transform: translateY(0) scale(1); }
        }
        .chijnaya-celda-heatmap {
          animation: chijnayaCeldaEntrada .4s ease both;
        }
        @keyframes chijnayaCeldaEntrada {
          from { opacity: 0; transform: scale(.4); }
          to   { opacity: 1; transform: scale(1); }
        }
        .chijnaya-celda-heatmap:hover {
          transform: scale(1.35);
          transition: transform .12s ease;
          box-shadow: 0 0 8px rgba(45,212,191,.6);
          z-index: 70;
        }
        .chijnaya-celda-hoy {
          position: relative;
        }
        .chijnaya-celda-hoy::after {
          content: "";
          position: absolute;
          inset: 0;
          border-radius: inherit;
          border: 2px solid #eef7f5;
          animation: chijnayaHoyPulso 1.8s ease-out infinite;
          pointer-events: none;
        }
        @keyframes chijnayaHoyPulso {
          0%   { transform: scale(1); opacity: .9; }
          100% { transform: scale(1.9); opacity: 0; }
        }
        .chijnaya-modo-transicion {
          animation: chijnayaModoEntrada .35s cubic-bezier(.2,.85,.35,1.15) both;
        }
        @keyframes chijnayaModoEntrada {
          from { opacity: 0; transform: scale(.985); }
          to   { opacity: 1; transform: scale(1); }
        }
        @media (prefers-reduced-motion: reduce) {
          .chijnaya-tarjeta-viva, .chijnaya-celda-heatmap, .chijnaya-modo-transicion {
            animation: none !important;
          }
        }
      `}</style>

      <div className="chijnaya-header-sticky">
        <div
          style={{
            maxWidth: modoPresentacion ? "100%" : 1500,
            margin: "0 auto",
            padding: modoPresentacion ? "18px 48px" : "16px 28px",
            display: "flex",
            justifyContent: "space-between",
            alignItems: "flex-start",
            flexWrap: "wrap",
            gap: 12,
            color: "#eef7f5",
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 14 }}>
            <img
              src={LOGO_PUNO_BASE64}
              alt="Escudo Gobierno Regional de Puno"
              style={{ width: modoPresentacion ? 52 : 40, height: modoPresentacion ? 58 : 45, flexShrink: 0 }}
            />
            <div>
              <h1 style={{ fontSize: modoPresentacion ? 36 : 24, marginBottom: 4, fontWeight: 800, letterSpacing: -0.3 }}>Expediente Técnico — C.S. Chijnaya</h1>
              <p style={{ color: "#b7c9c6", marginTop: 0, marginBottom: 4, fontSize: modoPresentacion ? 16 : 14 }}>
                Estado en tiempo real de la carga de documentación
              </p>
              {resumen?.ultimaSync?.toDate && (
                <p style={{ color: "#8fa8a8", fontSize: 11, marginTop: 0 }}>
                  Última sincronización: {tiempoRelativo(resumen.ultimaSync.toDate())}
                  {usuarioGoogle && <span style={{ marginLeft: 8, color: "#2dd4bf" }}>· {usuarioGoogle.email}</span>}
                </p>
              )}
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
            {carpetasForzadas.length > 0 && (
              <button
                onClick={() => setMostrarMarcadas(true)}
                style={{
                  fontSize: 13,
                  fontWeight: 700,
                  padding: "14px 18px",
                  borderRadius: 14,
                  border: "1.5px solid #2dd4bf66",
                  background: "#0e2529",
                  color: "#2dd4bf",
                  cursor: "pointer",
                  display: "flex",
                  alignItems: "center",
                  gap: 8,
                  whiteSpace: "nowrap",
                }}
              >
                <span style={{ fontSize: 16 }}>✓</span>
                Marcadas manualmente ({carpetasForzadas.length})
              </button>
            )}

            <button
              onClick={handleExportarGlobal}
              disabled={exportandoGlobal || carpetas.length === 0}
              style={{
                fontSize: 13,
                fontWeight: 700,
                padding: "14px 18px",
                borderRadius: 14,
                border: "1.5px solid #17a39866",
                background: exportandoGlobal ? "#122e2e" : "#0e2529",
                color: exportandoGlobal ? "#8fa8a8" : "#7fe0d4",
                cursor: exportandoGlobal || carpetas.length === 0 ? "not-allowed" : "pointer",
                display: "flex",
                alignItems: "center",
                gap: 8,
                whiteSpace: "nowrap",
              }}
              title="Generar PDF consolidado de todo el proyecto"
            >
              <span style={{ fontSize: 16 }}>📑</span>
              {exportandoGlobal ? "Generando Global..." : "Reporte Consolidado PDF"}
            </button>

            <button
              onClick={() => setModoPresentacion((v) => !v)}
              style={{
                fontSize: 14,
                fontWeight: 700,
                padding: "14px 20px",
                borderRadius: 14,
                border: modoPresentacion ? "2px solid #17a398" : "1.5px solid #2b5c5c",
                background: modoPresentacion ? "#17a39822" : "#0e2529",
                color: modoPresentacion ? "#17a398" : "#dceeec",
                cursor: "pointer",
                display: "flex",
                alignItems: "center",
                gap: 8,
                whiteSpace: "nowrap",
              }}
            >
              <span style={{ fontSize: 18 }}>🖥</span>
              {modoPresentacion ? "Salir de presentación" : "Modo presentación"}
            </button>
            <div style={{ textAlign: "right" }}>
              <button
                onClick={handleSync}
                disabled={sincronizando}
                style={{
                  fontSize: 20,
                  fontWeight: 800,
                  padding: "22px 42px",
                  borderRadius: 16,
                  border: "2px solid #17a39888",
                  background: sincronizando ? "#122e2e" : "linear-gradient(90deg,#17a39833,#0e7c7233)",
                  color: sincronizando ? "#9db3b0" : "#7fe0d4",
                  cursor: sincronizando ? "not-allowed" : "pointer",
                  boxShadow: sincronizando ? "none" : "0 0 28px rgba(23,163,152,.4)",
                  letterSpacing: 0.3,
                }}
              >
                {sincronizando ? "⟳ Sincronizando..." : "⟳ Sincronizar ahora"}
              </button>
              {mensajeSync && (
                <p
                  style={{
                    fontSize: 11,
                    marginTop: 6,
                    color: mensajeSync.tipo === "ok" ? "#2ecc71" : "#e74c3c",
                  }}
                >
                  {mensajeSync.texto}
                </p>
              )}
            </div>
          </div>
        </div>
      </div>

      <div style={{ maxWidth: modoPresentacion ? "100%" : 1500, margin: "0 auto", padding: modoPresentacion ? "24px 48px 36px" : "24px 28px 32px", color: "#eef7f5" }}>

        {/* Barra de progreso */}
        <div
          style={{
            marginBottom: 20,
            background: "rgba(8,28,31,.75)",
            backdropFilter: "blur(6px)",
            borderRadius: 12,
            padding: "16px 18px",
            border: "1px solid #1f4a4a",
          }}
        >
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 4 }}>
            <span style={{ fontSize: 16, fontWeight: 700, color: "#dceeec", letterSpacing: 0.5 }}>
              <span style={{ color: "#17a398" }}>»» </span>AVANCE POR CARPETAS
            </span>
            <strong style={{ fontSize: 30, textShadow: "0 0 18px rgba(23,163,152,.6)" }}>{pct}%</strong>
          </div>
          <div style={{ fontSize: 11, color: "#8fa8a8", marginBottom: 8 }}>
            {resumen?.completas ?? "–"} de {resumen?.totalFinales ?? "–"} carpetas marcadas como completas
          </div>
          <div
            style={{
              height: 34,
              background: "#0a1e20",
              borderRadius: 17,
              overflow: "hidden",
              boxShadow: "inset 0 2px 6px rgba(0,0,0,.5), 0 0 0 1px #1f4a4a",
            }}
          >
            <div
              className="chijnaya-barra-avance"
              style={{
                width: `${pct}%`,
                height: "100%",
                backgroundImage:
                  "repeating-linear-gradient(45deg, rgba(255,255,255,.18) 0px, rgba(255,255,255,.18) 9px, transparent 9px, transparent 18px), linear-gradient(90deg,#17a398,#0e7c72)",
                backgroundSize: "36px 36px, 100% 100%",
                transition: "width .4s ease",
                boxShadow: "0 0 22px rgba(14,124,114,.65)",
                borderRadius: 17,
              }}
            />
          </div>
        </div>

        {/* Barra de archivos */}
        <div
          style={{
            marginBottom: 32,
            background: "rgba(8,28,31,.75)",
            backdropFilter: "blur(6px)",
            borderRadius: 12,
            padding: "16px 18px",
            border: "1px solid #1f4a4a",
          }}
        >
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "baseline", marginBottom: 4 }}>
            <span style={{ fontSize: 16, fontWeight: 700, color: "#dceeec", letterSpacing: 0.5 }}>
              <span style={{ color: "#2dd4bf" }}>»» </span>AVANCE POR ARCHIVOS <span style={{ fontSize: 11, color: "#8fa8a8", fontWeight: 400 }}>(más preciso)</span>
            </span>
            <strong style={{ fontSize: 30, color: "#2dd4bf", textShadow: "0 0 18px rgba(45,212,191,.6)" }}>
              {resumen?.pctArchivos ?? "–"}%
            </strong>
          </div>
          <div style={{ fontSize: 11, color: "#8fa8a8", marginBottom: 8 }}>
            {resumen?.totalArchivosCompletados ?? "–"} de {resumen?.totalArchivosNecesarios ?? "–"} archivos que hacen falta, ya están subidos
          </div>
          <div
            style={{
              height: 22,
              background: "#0a1e20",
              borderRadius: 11,
              overflow: "hidden",
              boxShadow: "inset 0 2px 6px rgba(0,0,0,.5), 0 0 0 1px #1f4a4a",
            }}
          >
            <div
              style={{
                width: `${resumen?.pctArchivos ?? 0}%`,
                height: "100%",
                background: "linear-gradient(90deg,#2dd4bf,#0d9488)",
                transition: "width .4s ease",
                boxShadow: "0 0 16px rgba(45,212,191,.5)",
                borderRadius: 11,
              }}
            />
          </div>
        </div>

        {/* Contadores con desglose */}
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12, marginBottom: 32 }}>
          <Card label="Carpetas finales" value={resumen?.totalFinales ?? "–"} color="#17a398" grande={modoPresentacion} />
          <Card label="Completas" value={resumen?.completas ?? "–"} color="#2ecc71" grande={modoPresentacion} />
          <Card label="Incompletas" value={resumen?.incompletas ?? "–"} color="#f39c12" grande={modoPresentacion} />
          <Card label="Vacías" value={resumen?.vacias ?? "–"} color="#e74c3c" grande={modoPresentacion} />
        </div>

        {/* Selector de Rango de Fechas para Actividad/Heatmap */}
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
          <div style={{ fontSize: 15, fontWeight: 700, color: "#dceeec" }}>
            Visualización de Actividad e Historial
          </div>
          <div style={{ display: "flex", gap: 6 }}>
            {[
              { label: "30 días", val: 30 },
              { label: "84 días", val: 84 },
              { label: "119 días", val: 119 },
            ].map((btn) => (
              <button
                key={btn.val}
                onClick={() => setRangoDiasHeatmap(btn.val)}
                style={{
                  fontSize: 11,
                  padding: "5px 12px",
                  borderRadius: 16,
                  border: `1px solid ${rangoDiasHeatmap === btn.val ? "#2dd4bf" : "#2b5c5c"}`,
                  background: rangoDiasHeatmap === btn.val ? "#2dd4bf22" : "#0e2529",
                  color: rangoDiasHeatmap === btn.val ? "#2dd4bf" : "#b7c9c6",
                  fontWeight: 600,
                  cursor: "pointer",
                }}
              >
                {btn.label}
              </button>
            ))}
          </div>
        </div>

        {/* Sección de Tendencia de avance y Actividad */}
        {historial.length >= 2 ? (
          <div style={{ display: "grid", gridTemplateColumns: "minmax(0, 1.4fr) minmax(0, 1fr)", gap: 16, marginBottom: 32 }}>
            <TendenciaChart historial={historial} grande={modoPresentacion} actividadPorDia={actividadPorDia} />
            <ActividadHeatmap actividadPorDia={actividadPorDia} diasCustom={rangoDiasHeatmap} grande={modoPresentacion} onMarcarCompleta={handleMarcarCompleta} marcandoId={marcandoId} />
          </div>
        ) : (
          <div style={{ marginBottom: 32 }}>
            <ActividadHeatmap actividadPorDia={actividadPorDia} diasCustom={rangoDiasHeatmap} grande={modoPresentacion} onMarcarCompleta={handleMarcarCompleta} marcandoId={marcandoId} />
          </div>
        )}

        {modoPresentacion && (
          <div
            className="chijnaya-fade-in chijnaya-modo-transicion"
            style={{
              display: "grid",
              gridTemplateColumns: "repeat(auto-fit, minmax(280px, 1fr))",
              gap: 28,
              justifyItems: "center",
              marginTop: 8,
            }}
          >
            {areas.map((a) => {
              const stats = areaStats[a] || { total: 0, completas: 0, incompletas: 0, vacias: 0, archivosNecesarios: 0, archivosCompletados: 0 };
              const pctArea =
                stats.archivosNecesarios > 0
                  ? Math.round((stats.archivosCompletados / stats.archivosNecesarios) * 100)
                  : 0;
              return (
                <AreaMiniCard
                  key={a}
                  area={a}
                  pct={pctArea}
                  total={stats.total}
                  incompletas={stats.incompletas}
                  vacias={stats.vacias}
                  color={colorForArea(a)}
                  active={false}
                  onClick={() => {}}
                  tamano={280}
                />
              );
            })}
          </div>
        )}

        {modoPresentacion && areas.length > 0 && (
          <div className="chijnaya-fade-in chijnaya-modo-transicion" style={{ marginTop: 36 }}>
            <div style={{ fontSize: 16, fontWeight: 700, color: "#dceeec", marginBottom: 20, display: "flex", alignItems: "center", gap: 8 }}>
              <span style={{ color: "#2dd4bf" }}>»» </span>AVANCE POR ESPECIALIDAD, POR ÁREA
            </div>
            {areas.map((a) => {
              const especialidadesDelArea = especialidadPorArea[a] || {};
              const nombresOrdenados = Object.keys(especialidadesDelArea).sort((x, y) =>
                x.localeCompare(y, undefined, { numeric: true, sensitivity: "base" })
              );
              if (nombresOrdenados.length === 0) return null;
              return (
                <div key={a} style={{ marginBottom: 28 }}>
                  <div
                    style={{
                      fontSize: 13,
                      fontWeight: 700,
                      color: "#eef7f5",
                      textTransform: "uppercase",
                      letterSpacing: 0.5,
                      marginBottom: 12,
                      padding: "8px 12px",
                      display: "flex",
                      alignItems: "center",
                      gap: 8,
                      background: "rgba(8,28,31,.65)",
                      borderRadius: 8,
                      borderBottom: `2px solid ${colorForArea(a)}`,
                    }}
                  >
                    <span
                      style={{
                        width: 10,
                        height: 10,
                        borderRadius: "50%",
                        background: colorForArea(a),
                        flexShrink: 0,
                        boxShadow: `0 0 6px ${colorForArea(a)}`,
                      }}
                    />
                    {a} <span style={{ color: "#8fa8a8", fontWeight: 400, textTransform: "none" }}>({nombresOrdenados.length} especialidades)</span>
                  </div>
                  <div
                    style={{
                      display: "grid",
                      gridTemplateColumns: "repeat(auto-fit, minmax(200px, 1fr))",
                      gap: 18,
                    }}
                  >
                    {nombresOrdenados.map((esp, i) => {
                      const s = especialidadesDelArea[esp];
                      const pctEsp = s.archivosNecesarios > 0 ? Math.round((s.archivosCompletados / s.archivosNecesarios) * 100) : 0;
                      return (
                        <EspecialidadMiniCard
                          key={esp}
                          nombre={esp}
                          pct={pctEsp}
                          total={s.total}
                          incompletas={s.incompletas}
                          vacias={s.vacias}
                          delay={i * 30}
                        />
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {!modoPresentacion && (
        <div className="chijnaya-modo-transicion" style={{ display: "grid", gridTemplateColumns: "minmax(0, 1.3fr) minmax(0, 1fr)", gap: 24 }}>
          <div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8, flexWrap: "wrap", gap: 8 }}>
              <h2 style={{ fontSize: 16, color: "#dceeec", margin: 0 }}>
                Carpetas — {ESTADO_FILTRO_LABEL[filtroEstado]}{areaLabel}
              </h2>
            </div>

            <input
              type="text"
              value={busqueda}
              onChange={(e) => setBusqueda(e.target.value)}
              placeholder="🔍 Buscar carpeta por nombre..."
              style={{
                width: "100%",
                boxSizing: "border-box",
                background: "#0e2529",
                color: "#eef7f5",
                border: "1px solid #2b5c5c",
                borderRadius: 8,
                padding: "9px 12px",
                fontSize: 13,
                marginBottom: 12,
                outline: "none",
              }}
            />

            <div style={{ display: "flex", gap: 10, marginBottom: 14, flexWrap: "wrap" }}>
              {areas.map((a) => (
                <div key={a} style={{ display: "flex", gap: 4 }}>
                  <button
                    onClick={() => handleExportarArea(a, carpetasPorArea[a] || [])}
                    disabled={exportandoArea === a}
                    style={{
                      fontSize: 11,
                      padding: "6px 12px",
                      borderRadius: "20px 0 0 20px",
                      border: "1px solid #2b5c5c",
                      background: "#0e2529",
                      color: exportandoArea === a ? "#8fa8a8" : "#dceeec",
                      fontWeight: 600,
                      cursor: exportandoArea === a ? "not-allowed" : "pointer",
                    }}
                    title={`Exportar reporte PDF de ${a}`}
                  >
                    📄 {exportandoArea === a ? "Generando..." : `PDF ${a}`}
                  </button>
                  <button
                    onClick={() => handleExportarExcelArea(a, carpetasPorArea[a] || [])}
                    disabled={exportandoExcelArea === a}
                    style={{
                      fontSize: 11,
                      padding: "6px 12px",
                      borderRadius: "0",
                      border: "1px solid #2b5c5c",
                      borderLeft: "none",
                      background: "#0e2529",
                      color: exportandoExcelArea === a ? "#8fa8a8" : "#2dd4bf",
                      fontWeight: 600,
                      cursor: exportandoExcelArea === a ? "not-allowed" : "pointer",
                    }}
                    title={`Exportar reporte Excel de ${a}`}
                  >
                    📊 {exportandoExcelArea === a ? "Generando..." : "Excel"}
                  </button>
                  <button
                    onClick={() => handleExportarSeparadores(a, carpetasPorArea[a] || [])}
                    disabled={exportandoSeparadores === a}
                    style={{
                      fontSize: 11,
                      padding: "6px 12px",
                      borderRadius: "0",
                      border: "1px solid #2b5c5c",
                      borderLeft: "none",
                      background: "#0e2529",
                      color: exportandoSeparadores === a ? "#8fa8a8" : "#f0b27a",
                      fontWeight: 600,
                      cursor: exportandoSeparadores === a ? "not-allowed" : "pointer",
                    }}
                    title={`Exportar lista para separadores de ${a}`}
                  >
                    📑 {exportandoSeparadores === a ? "Generando..." : "Lista para separadores"}
                  </button>
                  <button
                    onClick={() => handleExportarListaGeneral(a, carpetasPorArea[a] || [])}
                    disabled={exportandoListaGeneral === a}
                    style={{
                      fontSize: 11,
                      padding: "6px 12px",
                      borderRadius: "0 20px 20px 0",
                      border: "1px solid #2b5c5c",
                      borderLeft: "none",
                      background: "#0e2529",
                      color: exportandoListaGeneral === a ? "#8fa8a8" : "#f0d264",
                      fontWeight: 600,
                      cursor: exportandoListaGeneral === a ? "not-allowed" : "pointer",
                    }}
                    title={`Exportar lista general (con N° de archivador) de ${a}`}
                  >
                    📋 {exportandoListaGeneral === a ? "Generando..." : "Lista general"}
                  </button>
                </div>
              ))}
            </div>

            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))",
                gap: 8,
                marginBottom: 14,
              }}
            >
              <AreaMiniCard
                area="Todas"
                pct={resumen?.pctArchivos ?? 0}
                total={resumen?.totalFinales ?? 0}
                color="#17a398"
                active={filtroArea === "Todas"}
                onClick={() => setFiltroArea("Todas")}
              />
              {areas.map((a) => {
                const s = areaStats[a];
                if (!s) return null;
                const areaPct =
                  s.archivosNecesarios > 0 ? Math.round((s.archivosCompletados / s.archivosNecesarios) * 100) : 0;
                return (
                  <AreaMiniCard
                    key={a}
                    area={a}
                    pct={areaPct}
                    total={s.total}
                    color={colorForArea(a)}
                    active={filtroArea === a}
                    onClick={() => setFiltroArea(filtroArea === a ? "Todas" : a)}
                  />
                );
              })}
            </div>

            {filtroArea !== "Todas" && areaStats[filtroArea] && (
              <AreaProgressPanel area={filtroArea} stats={areaStats[filtroArea]} color={colorForArea(filtroArea)} />
            )}

            <div style={{ display: "flex", gap: 6, marginBottom: 12, flexWrap: "wrap", alignItems: "center" }}>
              {ESTADO_OPTIONS.map((opt) => (
                <button
                  key={opt.value}
                  onClick={() => setFiltroEstado(opt.value)}
                  style={chipStyle(filtroEstado === opt.value, opt.color)}
                >
                  {opt.label}
                </button>
              ))}
              <button
                onClick={() => setColapsados((prev) => ({ ...prev, __all: !prev.__all }))}
                style={{ ...chipStyle(false, "#b7c9c6"), fontWeight: 700 }}
              >
                {colapsados.__all ? "▸ Expandir todo" : "▾ Colapsar todo"}
              </button>
            </div>

            <div
              key={`${filtroEstado}-${filtroArea}-${busqueda}`}
              className="chijnaya-fade-in"
              style={{
                background: "rgba(21,27,43,.5)",
                backdropFilter: "blur(6px)",
                borderRadius: 12,
                overflow: "hidden",
                border: "1px solid #1f4a4a",
              }}
            >
              {visibles.length === 0 && (
                <div
                  className="chijnaya-fade-in"
                  style={{
                    display: "flex",
                    flexDirection: "column",
                    alignItems: "center",
                    justifyContent: "center",
                    textAlign: "center",
                    padding: "48px 24px",
                    gap: 8,
                  }}
                >
                  <div style={{ fontSize: 34, opacity: 0.7 }}>
                    {carpetas.length === 0 ? "⏳" : "🔍"}
                  </div>
                  <div style={{ fontSize: 15, fontWeight: 700, color: "#dceeec" }}>
                    {carpetas.length === 0 ? "Sin datos todavía" : "No hay carpetas que coincidan"}
                  </div>
                </div>
              )}
              {(() => {
                const grupos = {};
                const ordenGrupos = [];
                for (const c of visibles) {
                  const partes = (c.ruta || c.nombre || "").split(" / ").filter(Boolean);
                  const especialidad = partes.length > 1 ? partes[1] : "(raíz)";
                  const key = `${c.area || "Sin área"} / ${especialidad}`;
                  if (!grupos[key]) {
                    grupos[key] = { area: c.area || "Sin área", especialidad, items: [] };
                    ordenGrupos.push(key);
                  }
                  grupos[key].items.push(c);
                }
                ordenGrupos.sort((a, b) => a.localeCompare(b, undefined, { numeric: true, sensitivity: "base" }));

                return ordenGrupos.map((key) => {
                  const g = grupos[key];
                  const color = "#0a6058";
                  const pendientesGrupo = g.items.filter((c) => c.estado !== "completa").length;
                  const vaciasGrupo = g.items.filter((c) => c.estado === "vacia").length;
                  const tienePendientes = pendientesGrupo > 0;
                  const grupoColapsado = colapsados.__all ? !colapsados[key] : !!colapsados[key];
                  return (
                    <div key={key}>
                      <div
                        onClick={() => toggleGrupo(key)}
                        style={{
                          padding: "10px 16px 10px 14px",
                          background: "#e3f2f0",
                          borderLeft: `4px solid ${vaciasGrupo > 0 ? "#e74c3c" : tienePendientes ? "#f39c12" : "#2ecc71"}`,
                          borderTop: "1px solid #1f4a4a",
                          borderBottom: "1px solid #bcdcd8",
                          display: "flex",
                          alignItems: "baseline",
                          gap: 8,
                          cursor: "pointer",
                          userSelect: "none",
                        }}
                      >
                        <span style={{ fontSize: 12, color: "#2f625e", transform: grupoColapsado ? "rotate(-90deg)" : "none", display: "inline-block", transition: "transform .15s ease" }}>
                          ▾
                        </span>
                        <span style={{ fontSize: 11, fontWeight: 800, color, textTransform: "uppercase", letterSpacing: 0.4 }}>
                          {g.area}
                        </span>
                        <span style={{ color: "#5ba39d", fontSize: 12 }}>›</span>
                        <span style={{ fontSize: 14, fontWeight: 800, color: "#0d3b3b" }}>{g.especialidad}</span>
                        <span style={{ display: "flex", alignItems: "center", gap: 8, marginLeft: "auto" }}>
                          <MiniDona completas={g.items.length - pendientesGrupo} total={g.items.length} />
                          {tienePendientes && (
                            <span
                              style={{
                                fontSize: 10,
                                padding: "2px 8px",
                                borderRadius: 20,
                                background: (vaciasGrupo > 0 ? "#e74c3c" : "#f39c12") + "22",
                                color: vaciasGrupo > 0 ? "#c0392b" : "#0d6b62",
                                fontWeight: 700,
                              }}
                            >
                              {pendientesGrupo} pendiente{pendientesGrupo !== 1 ? "s" : ""}
                            </span>
                          )}
                          <span style={{ fontSize: 10, color: "#2f625e", fontWeight: 600 }}>
                            {g.items.length} carpeta{g.items.length !== 1 ? "s" : ""}
                          </span>
                        </span>
                      </div>
                      {!grupoColapsado && g.items.map((c) => {
                        const detalle = c.detalle || c.estado;
                        const driveUrl = `https://drive.google.com/drive/folders/${c.id}`;
                        return (
                          <div
                            key={c.id}
                            onClick={() => window.open(driveUrl, "_blank", "noopener,noreferrer")}
                            style={{
                              padding: "10px 16px 10px 24px",
                              borderBottom: "1px solid #1f4a4a",
                              cursor: "pointer",
                            }}
                            onMouseEnter={(e) => (e.currentTarget.style.background = "#173838")}
                            onMouseLeave={(e) => (e.currentTarget.style.background = "transparent")}
                          >
                            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 8 }}>
                              <RutaJerarquica ruta={c.ruta} nombre={c.nombre} skipLevels={2} />
                              <span
                                style={{
                                  fontSize: 10,
                                  padding: "2px 8px",
                                  borderRadius: 20,
                                  background: ESTADO_COLOR[c.estado] + "22",
                                  color: ESTADO_COLOR[c.estado],
                                  textTransform: "uppercase",
                                  fontWeight: 600,
                                  whiteSpace: "nowrap",
                                  flexShrink: 0,
                                }}
                              >
                                {c.estado}{c.forzada ? " · manual" : ""}
                              </span>
                            </div>
                            {c.forzada && (
                              <div
                                style={{
                                  marginTop: 6,
                                  padding: "6px 10px",
                                  background: "#2dd4bf14",
                                  border: "1px solid #2dd4bf33",
                                  borderRadius: 8,
                                  fontSize: 11,
                                  color: "#9db3b0",
                                }}
                              >
                                ✓ Marcada por <strong style={{ color: "#2dd4bf" }}>{c.marcadoPor || "alguien"}</strong>
                              </div>
                            )}
                            <div style={{ display: "flex", justifyContent: "flex-end", alignItems: "center", gap: 10, marginTop: 6 }}>
                              <span style={{ fontSize: 11, color: "#b7c9c6" }}>{detalle}</span>
                              <button
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleMarcarCompleta(c.id, !c.forzada, c.nombre, c.ruta);
                                }}
                                disabled={marcandoId === c.id}
                                style={{
                                  fontSize: 10,
                                  padding: "3px 9px",
                                  borderRadius: 20,
                                  border: c.forzada ? "1px solid #e74c3c66" : "1px solid #2ecc7166",
                                  background: "transparent",
                                  color: marcandoId === c.id ? "#5c7a78" : c.forzada ? "#e88f86" : "#7fe0a3",
                                  cursor: marcandoId === c.id ? "not-allowed" : "pointer",
                                  whiteSpace: "nowrap",
                                  flexShrink: 0,
                                }}
                              >
                                {marcandoId === c.id ? "..." : c.forzada ? "✕ Desmarcar" : "✓ Marcar completa"}
                              </button>
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  );
                });
              })()}
            </div>
          </div>

          <div>
            <h2 style={{ fontSize: 16, color: "#dceeec", marginBottom: 8 }}>Actividad reciente</h2>
            <div
              style={{
                background: "rgba(21,27,43,.5)",
                backdropFilter: "blur(6px)",
                borderRadius: 12,
                maxHeight: 480,
                overflowY: "auto",
                border: "1px solid #1f4a4a",
              }}
            >
              {eventos.length === 0 && (
                <p style={{ padding: 16, color: "#b7c9c6" }}>Sin eventos todavía.</p>
              )}
              {eventos.map((e) => {
                const color = EVENTO_COLOR[e.tipo] || "#9db3b0";
                const icono = EVENTO_ICONO[e.tipo] || "•";
                const fecha = e.timestamp?.toDate ? e.timestamp.toDate() : null;
                return (
                  <div
                    key={e.id}
                    style={{
                      display: "flex",
                      gap: 10,
                      padding: "10px 14px",
                      borderBottom: "1px solid #1f4a4a",
                    }}
                  >
                    <div
                      style={{
                        flexShrink: 0,
                        width: 26,
                        height: 26,
                        borderRadius: "50%",
                        background: color + "22",
                        border: `1.5px solid ${color}`,
                        color: color,
                        display: "flex",
                        alignItems: "center",
                        justifyContent: "center",
                        fontSize: 13,
                        fontWeight: 700,
                        marginTop: 1,
                      }}
                    >
                      {icono}
                    </div>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ fontSize: 13 }}>
                        <strong>{e.usuario}</strong> <span style={{ color }}>{EVENTO_LABEL[e.tipo] || e.tipo}</span> <strong>{e.item}</strong>
                      </div>
                      <div style={{ fontSize: 10, color: "#8fa8a8", marginTop: 2 }}>
                        {tiempoRelativo(fecha)}
                      </div>
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </div>
        )}

      </div>

      {mostrarMarcadas && (
        <div
          onClick={() => setMostrarMarcadas(false)}
          style={{
            position: "fixed",
            inset: 0,
            background: "rgba(2,10,11,.75)",
            backdropFilter: "blur(3px)",
            zIndex: 100,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            padding: 20,
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            className="chijnaya-fade-in"
            style={{
              background: "#0e2529",
              border: "1px solid #2b5c5c",
              borderRadius: 16,
              width: "min(1100px, 100%)",
              maxHeight: "85vh",
              display: "flex",
              flexDirection: "column",
              boxShadow: "0 20px 60px rgba(0,0,0,.5)",
            }}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                padding: "20px 26px",
                borderBottom: "1px solid #1f4a4a",
              }}
            >
              <div style={{ fontSize: 19, fontWeight: 700, color: "#dceeec" }}>
                ✓ Carpetas marcadas manualmente ({carpetasForzadas.length})
              </div>
              <button
                onClick={() => setMostrarMarcadas(false)}
                style={{
                  fontSize: 14,
                  padding: "7px 14px",
                  borderRadius: 8,
                  border: "1px solid #2b5c5c",
                  background: "transparent",
                  color: "#9db3b0",
                  cursor: "pointer",
                }}
              >
                ✕ Cerrar
              </button>
            </div>
            <div style={{ overflowY: "auto", padding: "16px 26px 26px" }}>
              {carpetasForzadas.length === 0 ? (
                <div style={{ color: "#8fa8a8", fontSize: 15, padding: "24px 0" }}>
                  No hay ninguna carpeta marcada manualmente todavía.
                </div>
              ) : (
                carpetasForzadas.map((c) => (
                  <div
                    key={c.id}
                    style={{
                      padding: "16px 18px",
                      marginBottom: 12,
                      background: "#0a1e2066",
                      border: "1px solid #1f4a4a",
                      borderRadius: 10,
                    }}
                  >
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12 }}>
                      <div style={{ minWidth: 0 }}>
                        <div style={{ fontWeight: 700, color: "#eef7f5", fontSize: 16 }}>{c.nombre}</div>
                        <div style={{ fontSize: 13.5, color: "#9db3b0", marginTop: 3 }}>{c.ruta}</div>
                      </div>
                      <button
                        onClick={() => handleMarcarCompleta(c.id, false, c.nombre, c.ruta)}
                        disabled={marcandoId === c.id}
                        style={{
                          flexShrink: 0,
                          fontSize: 12,
                          padding: "5px 12px",
                          borderRadius: 20,
                          border: "1px solid #e74c3c66",
                          background: "transparent",
                          color: marcandoId === c.id ? "#5c7a78" : "#e88f86",
                          cursor: marcandoId === c.id ? "not-allowed" : "pointer",
                        }}
                      >
                        {marcandoId === c.id ? "..." : "✕ Desmarcar"}
                      </button>
                    </div>
                  </div>
                ))
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

function RutaJerarquica({ ruta, nombre, skipLevels = 0 }) {
  let partes = (ruta || nombre || "").split(" / ").filter(Boolean);
  if (skipLevels > 0 && partes.length > skipLevels) {
    partes = partes.slice(skipLevels);
  }
  let mostrar = partes;

  return (
    <div style={{ display: "flex", flexWrap: "wrap", alignItems: "baseline", gap: 4, flex: 1, minWidth: 0 }}>
      {mostrar.map((p, i) => {
        const esUltimo = i === mostrar.length - 1;
        return (
          <span key={i} style={{ display: "inline-flex", alignItems: "baseline", gap: 4 }}>
            {i > 0 && <span style={{ color: "#2b5c5c", fontSize: 11 }}>›</span>}
            <span
              style={{
                fontSize: esUltimo ? 14 : 11,
                fontWeight: esUltimo ? 700 : 500,
                color: esUltimo ? "#eef7f5" : "#8fa8a8",
              }}
            >
              {p}
              {esUltimo && <span style={{ color: "#8fa8a8", marginLeft: 4 }}>↗</span>}
            </span>
          </span>
        );
      })}
    </div>
  );
}

function EspecialidadMiniCard({ nombre, pct, total, incompletas = 0, vacias = 0, delay }) {
  const size = 110;
  const stroke = 8;
  const radius = (size - stroke) / 2;
  const circumference = 2 * Math.PI * radius;
  const offset = circumference - (pct / 100) * circumference;
  const color = pct >= 100 ? "#2ecc71" : pct >= 50 ? "#2dd4bf" : "#f39c12";

  return (
    <div
      className="chijnaya-tarjeta-viva"
      style={{
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        gap: 8,
        padding: "18px 12px",
        borderRadius: 12,
        background: "#0e252966",
        border: "1px solid #1f4a4a",
        animationDelay: `${delay}ms`,
      }}
    >
      <svg width={size} height={size}>
        <circle cx={size / 2} cy={size / 2} r={radius} fill="none" stroke="#1f4a4a" strokeWidth={stroke} />
        <circle
          cx={size / 2}
          cy={size / 2}
          r={radius}
          fill="none"
          stroke={color}
          strokeWidth={stroke}
          strokeLinecap="round"
          strokeDasharray={circumference}
          strokeDashoffset={offset}
          transform={`rotate(-90 ${size / 2} ${size / 2})`}
        />
        <text x="50%" y="50%" textAnchor="middle" dominantBaseline="central" fontSize="20" fontWeight="800" fill="#eef7f5">
          {pct}%
        </text>
      </svg>
      <div style={{ fontSize: 13, fontWeight: 700, color: "#eef7f5", textAlign: "center", lineHeight: 1.3, maxWidth: 150 }}>
        {nombre}
      </div>
      <div style={{ fontSize: 12, color: "#9db3b0", fontWeight: 700 }}>{total} carpetas</div>
      <div style={{ fontSize: 13, textAlign: "center", display: "flex", gap: 10, marginTop: 4 }}>
        <span style={{ color: "#f39c12", fontWeight: 700 }}>{incompletas} inc.</span>
        <span style={{ color: "#e74c3c", fontWeight: 700 }}>{vacias} vacías</span>
      </div>
    </div>
  );
}

function AreaMiniCard({ area, pct, total, incompletas = 0, vacias = 0, color, active, onClick, tamano }) {
  const size = tamano || 128;
  const stroke = Math.max(9, Math.round(size * 0.06));
  const radius = (size - stroke) / 2;
  const circumference = 2 * Math.PI * radius;
  const offset = circumference - (pct / 100) * circumference;
  const fontPct = Math.round(size * 0.22);
  const fontLabel = Math.max(14, Math.round(size * 0.09));
  const fontCount = Math.max(12, Math.round(size * 0.07));

  return (
    <button
      onClick={onClick}
      style={{
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        gap: Math.round(size * 0.05),
        padding: "18px 14px",
        borderRadius: 14,
        border: `2px solid ${active ? color : "#1f4a4a"}`,
        background: active ? color + "18" : "#0e2529",
        cursor: "pointer",
        transition: "all .15s ease",
        width: tamano ? "100%" : "auto",
      }}
    >
      <svg width={size} height={size}>
        <circle cx={size / 2} cy={size / 2} r={radius} fill="none" stroke="#1f4a4a" strokeWidth={stroke} />
        <circle
          cx={size / 2}
          cy={size / 2}
          r={radius}
          fill="none"
          stroke={color}
          strokeWidth={stroke}
          strokeLinecap="round"
          strokeDasharray={circumference}
          strokeDashoffset={offset}
          transform={`rotate(-90 ${size / 2} ${size / 2})`}
        />
        <text x="50%" y="50%" textAnchor="middle" dominantBaseline="central" fontSize={fontPct} fontWeight="700" fill="#eef7f5">
          {pct}%
        </text>
      </svg>
      <div style={{ fontSize: fontLabel, fontWeight: 700, color: active ? color : "#dceeec", textAlign: "center", marginTop: 4 }}>
        {area}
      </div>
      <div style={{ fontSize: fontCount, color: "#9db3b0", fontWeight: 700 }}>{total} carpetas</div>
      {tamano && (
        <div style={{ fontSize: 14, textAlign: "center", display: "flex", gap: 12, marginTop: 6 }}>
          <span style={{ color: "#f39c12", fontWeight: 700 }}>{incompletas} inc.</span>
          <span style={{ color: "#e74c3c", fontWeight: 700 }}>{vacias} vacías</span>
        </div>
      )}
    </button>
  );
}

function AreaProgressPanel({ area, stats, color }) {
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 20, background: "#0e2529", border: `1px solid ${color}44`, borderRadius: 10, padding: "16px 20px", marginBottom: 12 }}>
      <div>
        <div style={{ fontSize: 14, fontWeight: 700, color, marginBottom: 6 }}>{area}</div>
        <div style={{ fontSize: 12, color: "#eef7f5", lineHeight: 1.7 }}>
          <div><span style={{ color: "#2ecc71", fontWeight: 700 }}>{stats.completas}</span> completas de <strong>{stats.total}</strong> carpetas</div>
          <div><span style={{ color: "#f39c12", fontWeight: 700 }}>{stats.incompletas}</span> incompletas · <span style={{ color: "#e74c3c", fontWeight: 700 }}>{stats.vacias}</span> vacías</div>
        </div>
      </div>
    </div>
  );
}

function MiniDona({ completas, total }) {
  const size = 22;
  const stroke = 4;
  const radius = (size - stroke) / 2;
  const circumference = 2 * Math.PI * radius;
  const pct = total > 0 ? completas / total : 0;
  const offset = circumference - pct * circumference;
  const color = pct >= 1 ? "#2ecc71" : pct > 0 ? "#f39c12" : "#e74c3c";

  return (
    <svg width={size} height={size}>
      <circle cx={size / 2} cy={size / 2} r={radius} fill="none" stroke="#bcdcd8" strokeWidth={stroke} />
      <circle
        cx={size / 2}
        cy={size / 2}
        r={radius}
        fill="none"
        stroke={color}
        strokeWidth={stroke}
        strokeLinecap="round"
        strokeDasharray={circumference}
        strokeDashoffset={offset}
        transform={`rotate(-90 ${size / 2} ${size / 2})`}
      />
    </svg>
  );
}

function chipStyle(active, color) {
  return {
    fontSize: 11,
    padding: "5px 10px",
    borderRadius: 20,
    border: `1px solid ${active ? color : "#2b5c5c"}`,
    background: active ? color + "33" : "#0e2529",
    color: active ? color : "#b7c9c6",
    fontWeight: 600,
    cursor: "pointer",
  };
}

function useCountUp(target) {
  const [display, setDisplay] = useState(target);
  const prevRef = useRef(target);

  useEffect(() => {
    if (typeof target !== "number") {
      setDisplay(target);
      prevRef.current = target;
      return;
    }
    const from = typeof prevRef.current === "number" ? prevRef.current : target;
    const to = target;
    if (from === to) {
      setDisplay(to);
      return;
    }
    const duracion = 650;
    const inicio = performance.now();
    let raf;
    function tick(ahora) {
      const t = Math.min(1, (ahora - inicio) / duracion);
      const suavizado = 1 - Math.pow(1 - t, 3);
      setDisplay(Math.round(from + (to - from) * suavizado));
      if (t < 1) {
        raf = requestAnimationFrame(tick);
      } else {
        prevRef.current = to;
      }
    }
    raf = requestAnimationFrame(tick);
    return () => cancelAnimationFrame(raf);
  }, [target]);

  return display;
}

function Card({ label, value, color, grande }) {
  const valorAnimado = useCountUp(value);
  return (
    <div
      className="chijnaya-tarjeta-viva"
      style={{
        background: "rgba(14,37,41,.65)",
        backdropFilter: "blur(6px)",
        borderRadius: 12,
        padding: grande ? "26px" : "18px",
        border: `1px solid ${color}33`,
        borderTop: `3px solid ${color}`,
        boxShadow: `0 0 20px ${color}22`,
      }}
    >
      <div style={{ fontSize: grande ? 44 : 28, fontWeight: 700, textShadow: `0 0 14px ${color}55` }}>{valorAnimado}</div>
      <div style={{ fontSize: grande ? 15 : 12, color: "#b7c9c6", letterSpacing: 0.3 }}>{label}</div>
    </div>
  );
}

function TendenciaChart({ historial, grande, actividadPorDia }) {
  const altoLinea = grande ? 420 : 280;
  const altoBarras = grande ? 100 : 70;
  const alto = altoLinea + altoBarras;
  
  const anchoPunto = grande ? 65 : 55; // Separación amplia para evitar amontonamientos
  const paddingIzq = 60;
  const paddingDer = 40;
  const anchoMinimo = grande ? 960 : 720;
  const ancho = Math.max(anchoMinimo, paddingIzq + paddingDer + historial.length * anchoPunto);
  const paddingArriba = 24;

  return (
    <div
      style={{
        background: "#0e2529",
        border: "1px solid #2b5c5c",
        borderRadius: 12,
        padding: grande ? "28px 32px" : "16px 18px",
      }}
    >
      <div style={{ fontSize: grande ? 20 : 15, fontWeight: 700, color: "#dceeec", marginBottom: 14, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <span>📈 Tendencia de avance {historial.length > 0 ? `(${historial.length} días)` : ""}</span>
        <span style={{ fontSize: 11, color: "#8fa8a8", fontWeight: 400 }}>Desliza horizontalmente ↔</span>
      </div>

      {historial.length < 2 ? (
        <div style={{ fontSize: 12, color: "#8fa8a8", padding: "20px 0" }}>
          Todavía no hay suficiente historial.
        </div>
      ) : (
        (() => {
          const puntos = historial.map((h, i) => {
            const x = paddingIzq + i * anchoPunto;
            const y = paddingArriba + altoLinea - paddingArriba - (h.pct / 100) * (altoLinea - paddingArriba * 2);
            const tiposDia = actividadPorDia?.[h.fecha] || {};
            const incidencias = Object.values(tiposDia).reduce((s, n) => s + n, 0);
            return { x, y, pct: h.pct, fecha: h.fecha, incidencias };
          });
          const pathLinea = puntos.map((p, i) => `${i === 0 ? "M" : "L"} ${p.x} ${p.y}`).join(" ");
          const pathArea =
            `M ${puntos[0].x} ${altoLinea - paddingArriba} ` +
            puntos.map((p) => `L ${p.x} ${p.y}`).join(" ") +
            ` L ${puntos[puntos.length - 1].x} ${altoLinea - paddingArriba} Z`;

          const maxIncidencias = Math.max(1, ...puntos.map((p) => p.incidencias));
          const yBaseBarras = altoLinea + altoBarras - 16;
          const pasoEtiqueta = 1;

          return (
            <div ref={(el) => { if (el && !el.dataset.scrolled) { el.scrollLeft = el.scrollWidth; el.dataset.scrolled = "1"; } }} style={{ overflowX: "auto", overflowY: "hidden", width: "100%", paddingBottom: 8, scrollbarWidth: "thin", scrollbarColor: "#2b5c5c #0e2529" }}>
              <svg viewBox={`0 0 ${ancho} ${alto}`} style={{ width: `${ancho}px`, height: `${alto}px`, display: "block" }}>
                {[0, 25, 50, 75, 100].map((v) => {
                  const y = paddingArriba + altoLinea - paddingArriba - (v / 100) * (altoLinea - paddingArriba * 2);
                  return (
                    <g key={v}>
                      <line x1={paddingIzq - 10} y1={y} x2={ancho - paddingDer} y2={y} stroke="#2b5c5c" strokeWidth="1" strokeDasharray="3,4" />
                      <text x={paddingIzq - 16} y={y + 4} textAnchor="end" fontSize={grande ? 14 : 12} fill="#8fa8a8" fontWeight="600">
                        {v}%
                      </text>
                    </g>
                  );
                })}

                <path d={pathArea} fill="url(#tendenciaGradient)" opacity="0.4" />
                <path
                  d={pathLinea}
                  fill="none"
                  stroke="#17a398"
                  strokeWidth={grande ? "4.5" : "3.5"}
                  strokeLinecap="round"
                  strokeLinejoin="round"
                />
                {puntos.map((p, i) => (
                  <circle key={i} cx={p.x} cy={p.y} r={i === puntos.length - 1 ? (grande ? 7 : 5.5) : (grande ? 5 : 3.5)} fill="#17a398" />
                ))}

                <defs>
                  <linearGradient id="tendenciaGradient" x1="0" y1="0" x2="0" y2="1">
                    <stop offset="0%" stopColor="#17a398" />
                    <stop offset="100%" stopColor="#17a398" stopOpacity="0" />
                  </linearGradient>
                </defs>

                <text x={puntos[puntos.length - 1].x} y={puntos[puntos.length - 1].y - 14} textAnchor="end" fontSize={grande ? 18 : 15} fontWeight="700" fill="#7fe0d4">
                  {puntos[puntos.length - 1].pct}%
                </text>

                <line x1={paddingIzq - 10} y1={altoLinea + 8} x2={ancho - paddingDer} y2={altoLinea + 8} stroke="#1f4a4a" strokeWidth="1" />
                <text x={paddingIzq} y={altoLinea + 18} fontSize={grande ? 12 : 11} fill="#8fa8a8" fontWeight="700">
                  INCIDENCIAS DEL DRIVE POR DÍA
                </text>
                {puntos.map((p, i) => {
                  const alturaBarrita = Math.max(3, (p.incidencias / maxIncidencias) * (altoBarras - 28));
                  return (
                    <rect
                      key={i}
                      x={p.x - 4}
                      y={yBaseBarras - alturaBarrita}
                      width="8"
                      height={alturaBarrita}
                      rx="2"
                      fill={p.incidencias > 0 ? "#2dd4bf" : "#1f4a4a"}
                    />
                  );
                })}
                {puntos.map((p, i) => {
                  if (i % pasoEtiqueta !== 0 && i !== puntos.length - 1) return null;
                  const fechaObj = new Date(p.fecha + "T12:00:00");
                  const etiqueta = isNaN(fechaObj.getTime())
                    ? p.fecha
                    : fechaObj.toLocaleDateString("es-PE", { day: "numeric", month: "short" });
                  return (
                    <text key={i} x={p.x} y={alto - 2} textAnchor="middle" fontSize={grande ? 13 : 11} fill="#c8e8e5" fontWeight="600">
                      {etiqueta}
                    </text>
                  );
                })}
              </svg>
            </div>
          );
        })()
      )}
    </div>
  );
}

function ActividadHeatmap({ actividadPorDia, diasCustom = 84, grande }) {
  const DIAS = diasCustom;
  const [tooltip, setTooltip] = useState(null);
  const [tooltipPos, setTooltipPos] = useState(null);
  const tooltipRef = useRef(null);
  const [diaSeleccionado, setDiaSeleccionado] = useState(null);
  const [eventosDelDia, setEventosDelDia] = useState(null);

  useEffect(() => {
    if (!tooltip) {
      setTooltipPos(null);
      return;
    }
    const el = tooltipRef.current;
    if (!el) return;
    const rect = el.getBoundingClientRect();
    let left = tooltip.anclaX - rect.width / 2;
    left = Math.max(8, Math.min(window.innerWidth - rect.width - 8, left));
    let top = tooltip.anclaY - rect.height - 10;
    if (top < 8) top = tooltip.anclaY + 18;
    setTooltipPos({ left, top });
  }, [tooltip]);

  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);

  const conteoPorDia = {};
  const conteoPorTipoTotal = {};
  for (const [fechaKey, tipos] of Object.entries(actividadPorDia || {})) {
    let totalDia = 0;
    for (const [tipo, cantidad] of Object.entries(tipos || {})) {
      totalDia += cantidad;
      conteoPorTipoTotal[tipo] = (conteoPorTipoTotal[tipo] || 0) + cantidad;
    }
    conteoPorDia[fechaKey] = totalDia;
  }

  const dias = [];
  for (let i = DIAS - 1; i >= 0; i--) {
    const d = new Date(hoy);
    d.setDate(d.getDate() - i);
    const key = fechaLimaISO(d);
    dias.push({ key, count: conteoPorDia[key] || 0, fecha: d });
  }

  function intensidad(count) {
    if (count === 0) return "#1f4a4a";
    if (count >= 11) return "#2dd4bf";
    if (count >= 4) return "#17a398";
    return "#0e7c72";
  }

  const semanas = [];
  for (let i = 0; i < dias.length; i += 7) {
    semanas.push(dias.slice(i, i + 7));
  }

  const celda = grande ? 30 : 17;
  const gap = grande ? 7 : 4;
  const tiposOrdenados = Object.keys(conteoPorTipoTotal).sort((a, b) => conteoPorTipoTotal[b] - conteoPorTipoTotal[a]);

  async function abrirDetalleDia(d) {
    setDiaSeleccionado(d);
    setEventosDelDia(null);
    try {
      const inicioUTC = new Date(`${d.key}T05:00:00.000Z`);
      const finUTC = new Date(inicioUTC.getTime() + 24 * 60 * 60 * 1000);
      const q = query(
        collection(db, "eventos"),
        where("timestamp", ">=", inicioUTC),
        where("timestamp", "<", finUTC),
        orderBy("timestamp", "desc")
      );
      const snap = await getDocs(q);
      setEventosDelDia(snap.docs.map((doc) => ({ id: doc.id, ...doc.data() })));
    } catch {
      setEventosDelDia([]);
    }
  }

  function mostrarTooltip(e, texto) {
    const rect = e.currentTarget.getBoundingClientRect();
    setTooltip({ anclaX: rect.left + rect.width / 2, anclaY: rect.top, texto });
  }

  return (
    <div
      style={{
        background: "#0e2529",
        border: "1px solid #2b5c5c",
        borderRadius: 12,
        padding: grande ? "22px 26px" : "16px 18px",
        overflowX: "auto",
      }}
    >
      <div style={{ fontSize: grande ? 18 : 14, fontWeight: 700, color: "#dceeec", marginBottom: grande ? 18 : 10 }}>
        🔥 Actividad ({DIAS} días)
      </div>
      <div style={{ display: "flex", gap: grande ? 40 : 20, alignItems: "flex-start", flexWrap: "wrap" }}>
        <div style={{ display: "flex", gap: gap }}>
          {semanas.map((semana, si) => (
            <div key={si} style={{ display: "flex", flexDirection: "column", gap: gap }}>
              {semana.map((d, di) => {
                const esHoy = d.key === fechaLimaISO(new Date());
                const textoTooltip = `${formatearFechaLarga(d.fecha)}${esHoy ? " (hoy)" : ""} — ${d.count} evento${d.count !== 1 ? "s" : ""}`;
                return (
                  <div
                    key={d.key}
                    className={`chijnaya-celda-heatmap${esHoy ? " chijnaya-celda-hoy" : ""}`}
                    onMouseEnter={(e) => mostrarTooltip(e, textoTooltip)}
                    onMouseLeave={() => setTooltip(null)}
                    onClick={() => abrirDetalleDia(d)}
                    style={{
                      width: celda,
                      height: celda,
                      borderRadius: grande ? 5 : 3,
                      background: intensidad(d.count),
                      animationDelay: `${(si * 7 + di) * 4}ms`,
                      cursor: "pointer",
                    }}
                  />
                );
              })}
            </div>
          ))}
        </div>

        <div style={{ display: "flex", flexDirection: "column", gap: 16, minWidth: 240 }}>
          {tiposOrdenados.length > 0 && (
            <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
              <div style={{ fontSize: 13, fontWeight: 700, color: "#9db3b0", textTransform: "uppercase", letterSpacing: 0.4 }}>
                Resumen del período
              </div>
              {tiposOrdenados.map((tipo) => (
                <div key={tipo} style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 15 }}>
                  <span
                    style={{
                      width: 28,
                      height: 28,
                      borderRadius: "50%",
                      background: (EVENTO_COLOR[tipo] || "#9db3b0") + "22",
                      border: `1.5px solid ${EVENTO_COLOR[tipo] || "#9db3b0"}`,
                      color: EVENTO_COLOR[tipo] || "#9db3b0",
                      display: "flex",
                      alignItems: "center",
                      justifyContent: "center",
                      fontSize: 13,
                      flexShrink: 0,
                    }}
                  >
                    {EVENTO_ICONO[tipo] || "•"}
                  </span>
                  <strong>{conteoPorTipoTotal[tipo]}</strong>
                  <span style={{ color: "#b7c9c6" }}>{EVENTO_LABEL[tipo] || tipo}</span>
                </div>
              ))}
            </div>
          )}
        </div>
      </div>

      {tooltip && (
        <div
          ref={tooltipRef}
          style={{
            position: "fixed",
            left: tooltipPos ? tooltipPos.left : tooltip.anclaX,
            top: tooltipPos ? tooltipPos.top : tooltip.anclaY,
            visibility: tooltipPos ? "visible" : "hidden",
            background: "#0e2529",
            border: "1px solid #2b5c5c",
            color: "#eef7f5",
            padding: "8px 12px",
            borderRadius: 7,
            fontSize: 11.5,
            fontWeight: 600,
            width: 150,
            zIndex: 200,
            pointerEvents: "none",
          }}
        >
          {tooltip.texto}
        </div>
      )}

      {diaSeleccionado && (
        <div
          onClick={() => setDiaSeleccionado(null)}
          style={{
            position: "fixed",
            inset: 0,
            background: "rgba(2,10,11,.8)",
            backdropFilter: "blur(4px)",
            zIndex: 150,
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            padding: 20,
          }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            className="chijnaya-fade-in"
            style={{
              background: "#0e2529",
              border: "1px solid #2b5c5c",
              borderRadius: 16,
              width: "min(750px, 100%)",
              maxHeight: "85vh",
              display: "flex",
              flexDirection: "column",
              boxShadow: "0 20px 60px rgba(0,0,0,.6)",
            }}
          >
            <div
              style={{
                display: "flex",
                justifyContent: "space-between",
                alignItems: "center",
                padding: "20px 26px",
                borderBottom: "1px solid #1f4a4a",
              }}
            >
              <div>
                <div style={{ fontSize: 18, fontWeight: 700, color: "#dceeec" }}>
                  Actividad del {formatearFechaLarga(diaSeleccionado.fecha)}
                </div>
                <div style={{ fontSize: 12, color: "#8fa8a8", marginTop: 2 }}>
                  {diaSeleccionado.count} evento{diaSeleccionado.count !== 1 ? "s" : ""} registrado{diaSeleccionado.count !== 1 ? "s" : ""}
                </div>
              </div>
              <button
                onClick={() => setDiaSeleccionado(null)}
                style={{
                  fontSize: 14,
                  padding: "7px 14px",
                  borderRadius: 8,
                  border: "1px solid #2b5c5c",
                  background: "transparent",
                  color: "#9db3b0",
                  cursor: "pointer",
                }}
              >
                ✕ Cerrar
              </button>
            </div>
            <div style={{ overflowY: "auto", padding: "16px 26px 26px", flex: 1 }}>
              {eventosDelDia === null ? (
                <div style={{ textAlign: "center", padding: "40px 0", color: "#8fa8a8" }}>
                  Cargando eventos del día...
                </div>
              ) : eventosDelDia.length === 0 ? (
                <div style={{ textAlign: "center", padding: "40px 0", color: "#8fa8a8" }}>
                  No se encontraron eventos detallados para esta fecha.
                </div>
              ) : (
                <div style={{ display: "flex", flexDirection: "column", gap: 10 }}>
                  {eventosDelDia.map((ev) => {
                    const color = EVENTO_COLOR[ev.tipo] || "#9db3b0";
                    const icono = EVENTO_ICONO[ev.tipo] || "•";
                    const fechaEv = ev.timestamp?.toDate ? ev.timestamp.toDate() : null;
                    const horaStr = fechaEv ? fechaEv.toLocaleTimeString("es-PE", { hour: "2-digit", minute: "2-digit", second: "2-digit" }) : "";
                    return (
                      <div
                        key={ev.id}
                        style={{
                          background: "#0a1e20",
                          border: "1px solid #1f4a4a",
                          borderRadius: 10,
                          padding: "12px 16px",
                          display: "flex",
                          alignItems: "flex-start",
                          gap: 12,
                        }}
                      >
                        <div
                          style={{
                            flexShrink: 0,
                            width: 30,
                            height: 30,
                            borderRadius: "50%",
                            background: color + "22",
                            border: `1.5px solid ${color}`,
                            color: color,
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            fontSize: 14,
                            fontWeight: 700,
                          }}
                        >
                          {icono}
                        </div>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontSize: 14, color: "#eef7f5" }}>
                            <strong>{ev.usuario || "Usuario"}</strong> <span style={{ color }}>{EVENTO_LABEL[ev.tipo] || ev.tipo}</span> <strong style={{ color: "#7fe0d4" }}>{ev.item}</strong>
                          </div>
                          {ev.ruta && (
                            <div style={{ fontSize: 12, color: "#8fa8a8", marginTop: 3, wordBreak: "break-all" }}>
                              📁 {ev.ruta}
                            </div>
                          )}
                          <div style={{ fontSize: 11, color: "#5b827f", marginTop: 6, display: "flex", justifyContent: "space-between" }}>
                            <span>{ev.tipo}</span>
                            <span>{horaStr}</span>
                          </div>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
