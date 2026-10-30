import { google } from "googleapis";

function getAuth() {
  const key = JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_KEY);
  return new google.auth.JWT({
    email: key.client_email,
    key: key.private_key,
    scopes: [
      "https://www.googleapis.com/auth/drive.readonly",
      "https://www.googleapis.com/auth/drive.activity.readonly",
    ],
  });
}

export function getDriveClient() {
  return google.drive({ version: "v3", auth: getAuth() });
}

export function getActivityClient() {
  return google.driveactivity({ version: "v2", auth: getAuth() });
}

const FIELDS =
  "files(id,name,mimeType,parents,trashed,createdTime,modifiedTime,md5Checksum,webViewLink,thumbnailLink,lastModifyingUser(displayName,emailAddress,permissionId),sharingUser(displayName,emailAddress,permissionId),owners(displayName,emailAddress,permissionId))";

// Trae todos los items (archivos y subcarpetas) dentro de UNA carpeta.
// Maneja paginacion si hay mas de 1000 items en esa carpeta.
async function listFolderContents(drive, folderId) {
  const items = [];
  let pageToken = undefined;
  do {
    const res = await drive.files.list({
      q: `'${folderId}' in parents and trashed = false`,
      fields: `nextPageToken, ${FIELDS}`,
      pageSize: 1000,
      pageToken,
    });
    items.push(...(res.data.files || []));
    pageToken = res.data.nextPageToken;
  } while (pageToken);
  return items;
}

// Trae TODOS los items (carpetas y archivos) debajo de rootId, recursivamente.
// Version paralelizada: procesa por niveles (BFS) en vez de uno por uno (DFS secuencial).
// En cada nivel, pide TODAS las carpetas de ese nivel al mismo tiempo con Promise.all,
// en vez de esperar cada respuesta antes de pedir la siguiente carpeta.
export async function scanDriveTree(rootId) {
  const drive = getDriveClient();
  const allItems = [];
  const visited = new Set([rootId]);

  let currentLevel = [rootId];

  while (currentLevel.length > 0) {
    // Pide todas las carpetas del nivel actual EN PARALELO
    const results = await Promise.all(
      currentLevel.map((folderId) => listFolderContents(drive, folderId))
    );

    const nextLevel = [];

    for (const items of results) {
      for (const file of items) {
        allItems.push(file);
        if (
          file.mimeType === "application/vnd.google-apps.folder" &&
          !visited.has(file.id)
        ) {
          visited.add(file.id);
          nextLevel.push(file.id);
        }
      }
    }

    currentLevel = nextLevel;
  }

  return allItems;
}

// Quién borró/movió un elemento (Drive Activity) y lectura del historial real de Drive.
export async function getUltimoActorDeItem(itemId) {
  // 1) Drive guarda quién mandó a la papelera cada archivo/carpeta (trashingUser).
  try {
    const drive = getDriveClient();
    const r = await drive.files.get({
      fileId: itemId,
      fields: "trashed,trashedTime,trashingUser(displayName,emailAddress)",
      supportsAllDrives: true,
    });
    const u = r.data && r.data.trashingUser;
    if (u && (u.displayName || u.emailAddress)) {
      return u.displayName && u.emailAddress ? `${u.displayName} (${u.emailAddress})` : u.emailAddress || u.displayName;
    }
  } catch (err) {
    console.log(`[actor] papelera ${itemId}: ${String((err && err.message) || err).slice(0, 160)}`);
  }

  // 2) Drive Activity (si la cuenta de servicio tiene acceso al historial)
  try {
    const activity = getActivityClient();

    const res = await activity.activity.query({
      requestBody: {
        itemName: `items/${itemId}`,
        pageSize: 10,
      },
    });

    const activities = res.data.activities || [];

    // Se prefiere la actividad de BORRAR o MOVER (esa es la que dice quién fue);
    // si no hay, la más reciente que tenga una persona identificada.
    const esAccionDeBorrarMover = (act) => {
      const k = Object.keys(act.primaryActionDetail || {})[0];
      return k === "delete" || k === "move";
    };
    const ordenadas = [...activities.filter(esAccionDeBorrarMover), ...activities.filter((a) => !esAccionDeBorrarMover(a))];
    for (const act of ordenadas) {
      const actor = (act.actors || []).find((a) => a?.user?.knownUser);

      if (actor) {
        const ku = actor.user.knownUser;
        return ku.personName || null;
      }
    }

    return null;
  } catch (err) {
    console.log(`[actor] activity ${itemId}: ${String((err && err.message) || err).slice(0, 160)}`);
    return null;
  }
}

// Personas con acceso a las carpetas (permisos): sirve para traducir el código
// "people/123456" que entrega Drive Activity a "Nombre (correo)". El id de
// permiso de Drive es el mismo número que el de la persona.
export async function getPersonasConAcceso(carpetaIds) {
  const drive = getDriveClient();
  const mapa = new Map();
  await Promise.all(
    carpetaIds.map(async (id) => {
      try {
        let pageToken;
        do {
          const r = await drive.permissions.list({
            fileId: id,
            fields: "nextPageToken,permissions(id,type,displayName,emailAddress)",
            pageSize: 100,
            pageToken,
            supportsAllDrives: true,
          });
          for (const p of r.data.permissions || []) {
            if (p.type === "user" && p.id && (p.displayName || p.emailAddress)) {
              mapa.set(
                String(p.id),
                p.displayName && p.emailAddress ? `${p.displayName} (${p.emailAddress})` : p.emailAddress || p.displayName
              );
            }
          }
          pageToken = r.data.nextPageToken;
        } while (pageToken);
      } catch (err) {
        console.log(`[personas] ${id}: ${String((err && err.message) || err).slice(0, 120)}`);
      }
    })
  );
  return mapa;
}

// Historial REAL de actividad de Drive (Drive Activity API): cada acción con su
// hora exacta y quién la hizo. Se consulta una vez por carpeta principal, desde
// la fecha indicada. Devuelve [{ id, tipo, time, personName }] donde tipo es
// create | edit | move | delete | rename | restore, y personName es el código
// "people/123…" de la persona (no trae el correo).
export async function getActividadDrive(carpetaIds, desde, maxMs = 15000) {
  const t0 = Date.now();
  const resultados = [];
  const vistos = new Set();
  const activity = getActivityClient();

  async function consultar(folderId) {
    let pageToken;
    let paginas = 0;
    try {
      do {
        if (Date.now() - t0 > maxMs) return;
        const res = await activity.activity.query({
          requestBody: {
            ancestorName: `items/${folderId}`,
            filter: `time >= "${desde.toISOString()}"`,
            pageSize: 100,
            pageToken,
          },
        });
        for (const act of res.data.activities || []) {
          const tipo = Object.keys(act.primaryActionDetail || {})[0] || null;
          const t = act.timestamp || (act.timeRange && act.timeRange.endTime) || null;
          const actor = (act.actors || []).find((x) => x?.user?.knownUser);
          const personName = actor ? actor.user.knownUser.personName || null : null;
          const a0 = (act.actors || [])[0] || {};
          const actorTipo = actor ? "usuario" : a0.administrator ? "administrador" : a0.system ? "sistema" : a0.impersonation ? "suplantacion" : a0.user && a0.user.unknownUser ? "desconocido" : "otro";
          for (const tg of act.targets || []) {
            const nombre = (tg.driveItem && tg.driveItem.name) || "";
            const id = nombre.replace(/^items\//, "");
            if (!id || !t || !tipo) continue;
            const clave = `${id}|${tipo}|${t}`;
            if (vistos.has(clave)) continue;
            vistos.add(clave);
            resultados.push({ id, tipo, time: new Date(t), personName, actorTipo });
          }
        }
        pageToken = res.data.nextPageToken;
        paginas++;
      } while (pageToken && paginas < 10);
    } catch (err) {
      console.log(`[actividad] ${folderId}: ${String((err && err.message) || err).slice(0, 160)}`);
    }
  }

  await Promise.all(carpetaIds.map((id) => consultar(id)));
  return resultados;
}
