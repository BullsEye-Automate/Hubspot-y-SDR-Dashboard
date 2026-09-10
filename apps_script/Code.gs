/**
 * BullsEye — Formulario de Reuniones
 * Google Apps Script Backend
 *
 * Funciones:
 *   doGet()  → retorna los desplegables desde "Maestra IA" + reuniones pendientes
 *   doPost() → guarda/actualiza reuniones en "API Reuniones - IA"
 *
 * Deployment: Publicar como aplicación web
 *   - Ejecutar como: Yo (tu cuenta Google)
 *   - Acceso: Cualquier persona (incluso anónima)
 */

const SPREADSHEET_ID = "11vYlkzNlRwmEpGbeDNWpceDlhh36-B9gIgXU_emuRRk";
const MAESTRA_TAB    = "Maestra IA";
const OUTPUT_TAB     = "API Reuniones - IA";
const CACHE_KEY_META     = "doget_meta";     // clientes, sdrs, propuestas, etc. (pequeño)
const CACHE_KEY_MEETINGS = "doget_meetings"; // reuniones pendientes (grande, separado)
const CACHE_TTL          = 300; // 5 minutos

// Columnas del sheet de salida
const OUTPUT_HEADERS = [
  "Cliente",
  "Origen",
  "Responsable",
  "Empresa",
  "Contactos/Correo",
  "Fecha de agendamiento",
  "Fecha de la reunión",
  "Hora",
  "País",
  "Realizado",
  "Sales Manager",
  "Cargo",
  "Correo",
  "Teléfono",
  "Industria",
  "Comentario de la reunión",
  "Propuesta/Oportunidad",
  "Hora envío formulario"
];

// ── GET: retorna desplegables + reuniones pendientes (con caché) ───────────
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || "";

    // Metas persistentes
    if (action === "getGoals") {
      return getGoals();
    }

    // Estas acciones no usan caché (son llamadas puntuales legacy)
    if (action === "getMeetings") {
      const sdr     = (e.parameter && e.parameter.sdr)     || "";
      const cliente = (e.parameter && e.parameter.cliente) || "";
      const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
      return getMeetingsForSDR(ss, sdr, cliente);
    }
    if (action === "getClientsForSDR") {
      const sdr = (e.parameter && e.parameter.sdr) || "";
      const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
      return getClientsForSDR(ss, sdr);
    }

    // Respuesta principal: intentar desde caché primero
    const cache    = CacheService.getScriptCache();
    const metaJson = cacheGet(cache, CACHE_KEY_META);
    const mtgsJson = cacheGet(cache, CACHE_KEY_MEETINGS);
    if (metaJson && mtgsJson) {
      const data = JSON.parse(metaJson);
      data.meetings = JSON.parse(mtgsJson);
      return ContentService
        .createTextOutput(JSON.stringify({ status: "ok", data: data }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Abrir hoja una sola vez
    const ss       = SpreadsheetApp.openById(SPREADSHEET_ID);
    const data     = getMaestraData(ss);
    const meetings = getPendingMeetings(ss);
    data.meetings  = meetings;

    // Cachear meta y meetings con chunking para evitar límite 100KB de CacheService
    const metaOnly = { clientes: data.clientes, sdrs: data.sdrs, origenes: data.origenes,
                       paises: data.paises, industrias: data.industrias, propuestas: data.propuestas };
    cachePut(cache, CACHE_KEY_META,     JSON.stringify(metaOnly),  CACHE_TTL);
    cachePut(cache, CACHE_KEY_MEETINGS, JSON.stringify(meetings),  CACHE_TTL);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", data: data }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

// ── POST: guarda nueva reunión o actualiza una existente ──────────────────
function doPost(e) {
  try {
    const params = e.parameter;
    if (params.action === "saveGoals") {
      saveGoals(params.goals_json || "{}");
    } else if (params.action === "update") {
      updateReunion(params);
    } else {
      saveReunion(params);
    }
    // Invalidar caché para que el próximo doGet lea datos frescos
    const c = CacheService.getScriptCache();
    cacheRemove(c, CACHE_KEY_META);
    cacheRemove(c, CACHE_KEY_MEETINGS);
    return jsonResponse({ status: "ok", message: "OK" });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

// ── Leer opciones de la Maestra IA ────────────────────────────────────────
function getMaestraData(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(MAESTRA_TAB);
  if (!sheet) throw new Error("No se encontró la pestaña 'Maestra IA'");

  const values  = sheet.getDataRange().getValues();
  const headers = values[0].map(h => h.toString().trim());

  const idx = (keyword, exclude) => headers.findIndex(h => {
    const hl = h.toLowerCase();
    return hl.includes(keyword) && (!exclude || !hl.includes(exclude));
  });

  const colCliente       = idx("cliente",    "status");
  const colStCliente     = idx("status",     "responsable");
  const colResponsable   = idx("responsable","status");
  const colStResponsable = headers.findIndex((h, i) =>
    h.toLowerCase().includes("status") && i > colResponsable
  );
  const colOrigen    = idx("origen");
  const colPais      = headers.findIndex(h => /pa[íi]s/i.test(h));
  const colIndustria = idx("industria");

  const result = { clientes:[], sdrs:[], origenes:[], paises:[], industrias:[], propuestas:[] };
  const seen   = {
    clientes:   new Set(),
    sdrs:       new Set(),
    origenes:   new Set(),
    paises:     new Set(),
    industrias: new Set()
  };

  for (let i = 1; i < values.length; i++) {
    const r   = values[i];
    const val = (col) => col >= 0 ? r[col]?.toString().trim() : "";

    const cliente   = val(colCliente);
    const stCliente = val(colStCliente).toLowerCase();
    if (cliente && stCliente === "activo" && !seen.clientes.has(cliente)) {
      result.clientes.push(cliente);
      seen.clientes.add(cliente);
    }

    const sdr   = val(colResponsable);
    const stSdr = val(colStResponsable).toLowerCase();
    if (sdr && stSdr === "activo" && !seen.sdrs.has(sdr)) {
      result.sdrs.push(sdr);
      seen.sdrs.add(sdr);
    }

    const origen = val(colOrigen);
    if (origen && !seen.origenes.has(origen)) {
      result.origenes.push(origen);
      seen.origenes.add(origen);
    }

    const pais = val(colPais);
    if (pais && !seen.paises.has(pais)) {
      result.paises.push(pais);
      seen.paises.add(pais);
    }

    const industria = val(colIndustria);
    if (industria && !seen.industrias.has(industria)) {
      result.industrias.push(industria);
      seen.industrias.add(industria);
    }
  }

  result.clientes.sort();
  result.sdrs.sort();
  result.origenes.sort();
  result.paises.sort();
  result.industrias.sort();

  // Leer opciones de validación de datos de la columna "Propuesta/Oportunidad"
  // Una sola llamada batch (getDataValidations) en vez de hasta 50 llamadas individuales
  try {
    const outSheet = ss.getSheetByName(OUTPUT_TAB);
    if (outSheet && outSheet.getLastRow() > 1) {
      const outHeadersL = outSheet.getRange(1, 1, 1, outSheet.getLastColumn()).getValues()[0]
        .map(h => h.toString().trim().toLowerCase());
      const colP = outHeadersL.findIndex(h => h.includes("propuesta") || h.includes("oportunidad"));
      if (colP >= 0) {
        const numRows    = Math.min(outSheet.getLastRow() - 1, 50);
        const validations = outSheet.getRange(2, colP + 1, numRows, 1).getDataValidations();
        for (let r = 0; r < validations.length; r++) {
          const dv = validations[r][0];
          if (dv && dv.getCriteriaType() === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
            result.propuestas = dv.getCriteriaValues()[0] || [];
            break;
          }
        }
      }
    }
  } catch(e) {}

  // Fallback: lista fija si no se encontró validación de datos en el sheet
  if (!result.propuestas.length) {
    result.propuestas = [
      "Si","No","Pendiente de envio","Licitación","Exploratoria",
      "Demo","Cierre perdido","Cliente","Feedback pendiente",
      "buscando Segunda reunión","No era el contacto - Escalar"
    ];
  }

  return result;
}

// ── Obtener TODAS las reuniones pendientes (para filtrado client-side) ────
function getPendingMeetings(ss) {
  if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(OUTPUT_TAB);
  if (!sheet) return [];

  const all = sheet.getDataRange().getValues();
  if (all.length < 2) return [];

  const headers = all[0].map(h => h.toString().trim().toLowerCase());

  const col = (kw1, kw2, exclude) => headers.findIndex(h =>
    h.includes(kw1) &&
    (!kw2    || h.includes(kw2)) &&
    (!exclude || !h.includes(exclude))
  );

  const colResponsable = col("responsable");
  const colRealizado   = col("realizado");
  const colCliente     = col("cliente");
  const colEmpresa     = col("empresa");
  const colContacto    = col("contacto");
  const colCargo       = col("cargo");
  const colPais        = headers.findIndex(h => /pa[íi]s/.test(h));
  const colFecha       = col("fecha", "reuni");
  const colHora        = col("hora");
  const colPropuesta   = headers.findIndex(h => h.includes("propuesta") || h.includes("oportunidad"));

  const pendingValues = ["pendiente", "no", "reagendar", ""];
  const meetings = [];

  for (let i = 1; i < all.length; i++) {
    const r             = all[i];
    const realizadoRaw  = colRealizado >= 0 ? r[colRealizado]?.toString().trim() : "";
    const realizado     = realizadoRaw.toLowerCase();

    if (!pendingValues.includes(realizado)) continue;

    meetings.push({
      row:       i + 1,
      sdr:       colResponsable >= 0 ? r[colResponsable]?.toString().trim() : "",
      cliente:   colCliente     >= 0 ? r[colCliente]?.toString().trim()     : "",
      empresa:   colEmpresa     >= 0 ? r[colEmpresa]?.toString().trim()     : "",
      contacto:  colContacto    >= 0 ? r[colContacto]?.toString().trim()    : "",
      cargo:     colCargo       >= 0 ? r[colCargo]?.toString().trim()       : "",
      pais:      colPais        >= 0 ? r[colPais]?.toString().trim()        : "",
      fecha:     colFecha       >= 0 ? _cellDate(r[colFecha])              : "",
      hora:      colHora        >= 0 ? r[colHora]?.toString().trim()        : "",
      // Se preserva el valor original (no en minúsculas) porque el formulario
      // compara "estado" contra los literales exactos de sus <option value>.
      estado:    realizadoRaw,
      propuesta: colPropuesta   >= 0 ? r[colPropuesta]?.toString().trim()   : "",
    });
  }

  return meetings;
}

// ── Guardar fila en "API Reuniones - IA" ──────────────────────────────────
function saveReunion(p) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(OUTPUT_TAB);

  // Crear pestaña con encabezados si no existe
  if (!sheet) {
    sheet = ss.insertSheet(OUTPUT_TAB);
    const headerRange = sheet.getRange(1, 1, 1, OUTPUT_HEADERS.length);
    headerRange.setValues([OUTPUT_HEADERS]);
    headerRange.setBackground("#1a1b4d");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    sheet.setFrozenRows(1);
  }

  const now = Utilities.formatDate(
    new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss"
  );

  // Matching flexible por header (tolera variaciones de nombre)
  function valueForHeader(h) {
    const hl = h.toString().toLowerCase().trim();
    if (hl === "cliente")                                        return p.cliente            || "";
    if (hl === "origen")                                         return p.origen             || "";
    if (hl.includes("responsable"))                              return p.responsable        || "";
    // empresa debe ir ANTES de contacto: "Empresa del contacto" contiene "contacto"
    // pero debe mapearse a empresa, no a contactos.
    if (hl.includes("empresa"))                                  return p.empresa            || "";
    if (hl.includes("contacto"))                                 return p.contactos          || "";
    if (hl.includes("fecha") && hl.includes("agend"))            return p.fecha_agendamiento || "";
    if (hl.includes("fecha") && hl.includes("reuni"))            return p.fecha_reunion      || "";
    // "Hora envío formulario" va ANTES del check simple de "hora"
    if (hl.includes("hora")  && hl.includes("env"))              return now;
    if (hl.includes("fecha") && hl.includes("env"))              return now;
    if (hl.includes("fecha") && hl.includes("registro"))         return now;
    if (hl === "hora" || hl.includes("hora"))                    return p.hora               || "";
    if (/pa[íi]s/.test(hl))                                     return p.pais               || "";
    if (hl === "realizado")                                      return p.realizado          || "";
    if (hl.includes("sales") || hl.includes("manager"))          return p.sales_manager      || "";
    if (hl === "cargo")                                          return p.cargo              || "";
    if (hl.includes("correo") || hl === "email")                 return p.correo             || "";
    if (hl.includes("tel"))                                      return p.telefono           || "";
    if (hl.includes("industria"))                                return p.industria          || "";
    if (hl.includes("comentario"))                               return p.comentario         || "";
    if (hl.includes("propuesta") || hl.includes("oportunidad"))  return p.propuesta          || "";
    return "";
  }

  // Leer headers actuales del sheet y construir la fila en el orden correcto
  const sheetHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const row = sheetHeaders.map(h => valueForHeader(h));

  sheet.appendRow(row);
}

// ── Obtener clientes únicos con reuniones pendientes para un SDR ──────────
function getClientsForSDR(ss, sdrName) {
  if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(OUTPUT_TAB);
  if (!sheet) return jsonResponse({ status: "ok", clientes: [] });

  const all = sheet.getDataRange().getValues();
  if (all.length < 2) return jsonResponse({ status: "ok", clientes: [] });

  const headers = all[0].map(h => h.toString().trim().toLowerCase());
  const col     = (kw1, kw2, exclude) => headers.findIndex(h =>
    h.includes(kw1) && (!kw2 || h.includes(kw2)) && (!exclude || !h.includes(exclude))
  );

  const colResponsable = col("responsable");
  const colRealizado   = col("realizado");
  const colCliente     = col("cliente");

  const pendingValues = ["pendiente", "no", "reagendar", ""];
  const seen     = new Set();
  const clientes = [];

  for (let i = 1; i < all.length; i++) {
    const r           = all[i];
    const responsable = colResponsable >= 0 ? r[colResponsable]?.toString().trim()               : "";
    const realizado   = colRealizado   >= 0 ? r[colRealizado]?.toString().trim().toLowerCase()   : "";
    const cliente     = colCliente     >= 0 ? r[colCliente]?.toString().trim()                   : "";

    if (responsable !== sdrName) continue;
    if (!pendingValues.includes(realizado)) continue;
    if (cliente && !seen.has(cliente)) {
      clientes.push(cliente);
      seen.add(cliente);
    }
  }

  clientes.sort();
  return jsonResponse({ status: "ok", clientes: clientes });
}

// ── Obtener reuniones pendientes para un SDR (legacy / puntual) ───────────
function getMeetingsForSDR(ss, sdrName, clienteFilter) {
  if (!ss) ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(OUTPUT_TAB);
  if (!sheet) return jsonResponse({ status: "ok", meetings: [] });

  const all = sheet.getDataRange().getValues();
  if (all.length < 2) return jsonResponse({ status: "ok", meetings: [] });

  const headers = all[0].map(h => h.toString().trim().toLowerCase());

  const col = (kw1, kw2, exclude) => headers.findIndex(h =>
    h.includes(kw1) &&
    (!kw2    || h.includes(kw2)) &&
    (!exclude || !h.includes(exclude))
  );

  const colResponsable = col("responsable");
  const colRealizado   = col("realizado");
  const colEmpresa     = col("empresa");
  const colContacto    = col("contacto");
  const colCargo       = col("cargo");
  const colPais        = headers.findIndex(h => /pa[íi]s/.test(h));
  const colFecha       = col("fecha", "reuni");
  const colHora        = col("hora");
  const colCliente     = col("cliente");
  const colPropuesta   = headers.findIndex(h => h.includes("propuesta") || h.includes("oportunidad"));

  const pendingValues = ["pendiente", "no", "reagendar", ""];
  const meetings      = [];

  for (let i = 1; i < all.length; i++) {
    const r           = all[i];
    const responsable = colResponsable >= 0 ? r[colResponsable]?.toString().trim() : "";
    const realizado   = colRealizado   >= 0 ? r[colRealizado]?.toString().trim()   : "";

    if (responsable !== sdrName) continue;
    if (!pendingValues.includes(realizado.toLowerCase())) continue;
    if (clienteFilter) {
      const c = colCliente >= 0 ? r[colCliente]?.toString().trim() : "";
      if (c !== clienteFilter) continue;
    }

    meetings.push({
      row:       i + 1,
      sdr:       responsable,
      empresa:   colEmpresa   >= 0 ? r[colEmpresa]?.toString().trim()   : "",
      contacto:  colContacto  >= 0 ? r[colContacto]?.toString().trim()  : "",
      cargo:     colCargo     >= 0 ? r[colCargo]?.toString().trim()     : "",
      pais:      colPais      >= 0 ? r[colPais]?.toString().trim()      : "",
      fecha:     colFecha     >= 0 ? _cellDate(r[colFecha])              : "",
      hora:      colHora      >= 0 ? r[colHora]?.toString().trim()      : "",
      estado:    realizado,
      cliente:   colCliente   >= 0 ? r[colCliente]?.toString().trim()   : "",
      propuesta: colPropuesta >= 0 ? r[colPropuesta]?.toString().trim() : "",
    });
  }

  return jsonResponse({ status: "ok", meetings: meetings });
}

// ── Actualizar fecha y estado de una reunión existente ────────────────────
function updateReunion(p) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(OUTPUT_TAB);
  if (!sheet) throw new Error("No se encontró la pestaña '" + OUTPUT_TAB + "'");

  const rowNum = parseInt(p.row, 10);
  if (isNaN(rowNum) || rowNum < 2) throw new Error("Número de fila inválido: " + p.row);

  const headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const headersL = headers.map(h => h.toString().trim().toLowerCase());
  const colIndex = (name) => headers.findIndex(h => h.toString().trim() === name);

  const colFecha      = colIndex("Fecha de la reunión");
  const colHora       = colIndex("Hora");
  const colRealizado  = colIndex("Realizado");
  const colComentario = colIndex("Comentario de la reunión");
  const colPropuesta  = headersL.findIndex(h => h.includes("propuesta") || h.includes("oportunidad"));

  if (colFecha >= 0 && p.fecha_reunion !== undefined) {
    sheet.getRange(rowNum, colFecha + 1).setValue(p.fecha_reunion);
  }
  if (colHora >= 0 && p.hora !== undefined) {
    sheet.getRange(rowNum, colHora + 1).setValue(p.hora);
  }
  if (colRealizado >= 0 && p.realizado !== undefined) {
    sheet.getRange(rowNum, colRealizado + 1).setValue(p.realizado);
  }
  if (colComentario >= 0 && p.comentario !== undefined && p.comentario !== "") {
    sheet.getRange(rowNum, colComentario + 1).setValue(p.comentario);
  }
  if (colPropuesta >= 0 && p.propuesta !== undefined) {
    sheet.getRange(rowNum, colPropuesta + 1).setValue(p.propuesta);
  }
}

// ── Metas persistentes ────────────────────────────────────────────────────────
// Se guardan como JSON en la celda A1 de la pestaña "Metas".
// Así sobreviven reinicios de Streamlit Cloud.
const GOALS_TAB = "Metas";

function getGoals() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(GOALS_TAB);
    if (!sheet) {
      return jsonResponse({ status: "ok", goals: { sdr: {}, cliente: {} } });
    }
    const raw = sheet.getRange(1, 1).getValue();
    const goals = raw ? JSON.parse(raw) : { sdr: {}, cliente: {} };
    return jsonResponse({ status: "ok", goals: goals });
  } catch (err) {
    return jsonResponse({ status: "ok", goals: { sdr: {}, cliente: {} } });
  }
}

function saveGoals(goalsJson) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(GOALS_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(GOALS_TAB);
    // Ocultar la pestaña para que no moleste al equipo
    sheet.hideSheet();
  }
  sheet.getRange(1, 1).setValue(goalsJson);
}

// ── Helper: respuesta JSON ─────────────────────────────────────────────────
function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Helpers de caché con chunking (supera el límite de 100KB por entrada) ──
// Divide el valor en trozos de 90KB y los guarda en claves numeradas.
// Así el tamaño efectivo es ilimitado (hasta el cupo total del script cache).
const CHUNK_SIZE = 90000; // 90KB por chunk (límite es 100KB)

function cachePut(cache, key, value, ttl) {
  const chunks = [];
  for (let i = 0; i < value.length; i += CHUNK_SIZE) {
    chunks.push(value.substring(i, i + CHUNK_SIZE));
  }
  const entries = {};
  entries[key + "__n"] = chunks.length.toString();
  chunks.forEach((chunk, idx) => { entries[key + "__" + idx] = chunk; });
  cache.putAll(entries, ttl);
}

function cacheGet(cache, key) {
  const countStr = cache.get(key + "__n");
  if (!countStr) return null;
  const count = parseInt(countStr, 10);
  const keys  = [];
  for (let i = 0; i < count; i++) keys.push(key + "__" + i);
  const parts = cache.getAll(keys);
  let result  = "";
  for (let i = 0; i < count; i++) {
    const part = parts[key + "__" + i];
    if (part === null || part === undefined) return null; // algún chunk expiró
    result += part;
  }
  return result;
}

function cacheRemove(cache, key) {
  const countStr = cache.get(key + "__n");
  if (countStr) {
    const count = parseInt(countStr, 10);
    for (let i = 0; i < count; i++) cache.remove(key + "__" + i);
    cache.remove(key + "__n");
  }
}

// ── Helper: formatear celda de fecha → "DD/MM/YYYY" ───────────────────────
// getValues() devuelve objetos Date para celdas con formato fecha.
// .toString() sobre ese Date genera un string largo que flatpickr parsea mal.
// Usamos Utilities.formatDate para garantizar el formato correcto.
function _cellDate(val) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  return val.toString().trim();
}
