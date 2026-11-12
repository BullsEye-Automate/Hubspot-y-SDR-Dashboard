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
  "ID Cliente",
  "Origen",
  "Responsable",
  "Empresa",
  "Contactos/Correo",
  "URL Linkedin",
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
                       paises: data.paises, industrias: data.industrias, propuestas: data.propuestas,
                       clientesMap: data.clientesMap };
    // No cachear si clientes o sdrs vinieron vacíos: evita que un hiccup puntual
    // (ej. una lectura que coincide con una edición en curso de "Maestra IA")
    // quede sirviéndose como respuesta "pegada" durante los 5 minutos de TTL.
    if (metaOnly.clientes.length > 0 && metaOnly.sdrs.length > 0) {
      cachePut(cache, CACHE_KEY_META,     JSON.stringify(metaOnly),  CACHE_TTL);
      cachePut(cache, CACHE_KEY_MEETINGS, JSON.stringify(meetings),  CACHE_TTL);
    }

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
  const colIdCliente = headers.findIndex(h =>
    h.toLowerCase().includes("id") && h.toLowerCase().includes("cliente")
  );

  const result = { clientes:[], sdrs:[], origenes:[], paises:[], industrias:[], propuestas:[], clientesMap:{} };
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
      const idCliente = colIdCliente >= 0 ? val(colIdCliente) : "";
      if (idCliente) result.clientesMap[cliente] = idCliente;
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
  const colReconf      = headers.indexOf(COL_RECONF.toLowerCase());
  const colReconfFecha = headers.indexOf(COL_RECONF_FECHA.toLowerCase());
  const colCorreo      = headers.indexOf("correo"); // exacto: "Contactos/Correo" es el nombre
  const colTelefono    = headers.findIndex(h => h.includes("tel"));

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
      reconfirmada:      colReconf      >= 0 ? _esSi(r[colReconf])            : false,
      fechaReconfirmada: colReconfFecha >= 0 ? r[colReconfFecha]?.toString().trim() : "",
      correo:    colCorreo      >= 0 ? r[colCorreo]?.toString().trim()      : "",
      telefono:  colTelefono    >= 0 ? r[colTelefono]?.toString().trim()    : "",
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

  const clientesMap = getMaestraData(ss).clientesMap; // Cliente -> ID Cliente (misma fuente que tu BUSCARV)

  // Matching flexible por header (tolera variaciones de nombre)
  function valueForHeader(h) {
    const hl = h.toString().toLowerCase().trim();
    if (hl === "cliente")                                        return p.cliente            || "";
    if (hl.includes("id") && hl.includes("cliente"))            return p.id_cliente || clientesMap[p.cliente] || "";
    if (hl.includes("id") && hl.includes("reuni"))              return Utilities.getUuid();
    if (hl === "origen")                                         return p.origen             || "";
    if (hl.includes("responsable"))                              return p.responsable        || "";
    // empresa debe ir ANTES de contacto: "Empresa del contacto" contiene "contacto"
    // pero debe mapearse a empresa, no a contactos.
    if (hl.includes("empresa"))                                  return p.empresa            || "";
    // linkedin debe ir ANTES de contacto: "URL Linkedin del contacto" contiene
    // "contacto" pero debe mapearse al link, no al nombre.
    if (hl.includes("linkedin"))                                 return p.linkedin_contacto  || "";
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

// Ejecutar UNA SOLA VEZ manualmente (seleccionarla en el desplegable de
// funciones, arriba al lado de "Depurar", y apretar "Ejecutar") para
// rellenar con un ID aleatorio las reuniones que ya existían antes de
// agregar la columna "ID Reunión".
function backfillMeetingIds() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(OUTPUT_TAB);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idCol = headers.findIndex(h => h.toString().trim() === "ID Reunión") + 1;
  if (idCol === 0) throw new Error('No encontré la columna "ID Reunión" — agregala primero en la hoja.');

  const lastRow = sheet.getLastRow();
  for (let r = 2; r <= lastRow; r++) {
    const cell = sheet.getRange(r, idCol);
    if (!cell.getValue()) cell.setValue(Utilities.getUuid());
  }
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

  // Fecha/hora antes del cambio: si la reunión se mueve, la reconfirmación
  // anterior deja de valer y el recordatorio debe volver a enviarse.
  const prevRow = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  p._fechaAnterior = colFecha >= 0 ? _cellDate(prevRow[colFecha]) : "";
  p._horaAnterior  = colHora  >= 0 ? _cellHora(prevRow[colHora])  : "";

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

  // ── Reconfirmación ──────────────────────────────────────────────────────
  if (p.reconfirmada !== undefined) {
    actualizarReconfirmacion(sheet, rowNum, p);
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


// ══════════════════════════════════════════════════════════════════════════
//  RECONFIRMACIÓN DE REUNIONES
// ══════════════════════════════════════════════════════════════════════════
//
//  Qué hace:
//   1. El SDR marca "Reconfirmada" en el formulario (modo Actualizar).
//      → Se guarda en el Sheet (columnas al final, se crean solas).
//      → Se actualiza el contacto en HubSpot (portal BullsEye):
//        "Reunión reconfirmada" = Sí y "Fecha reconfirmación reunión".
//      → (Opcional) Se envía un correo de aviso al cliente.
//   2. Un activador cada 15 min (revisarReconfirmaciones) avisa al SDR por
//      correo cuando le toca reconfirmar, y escala al supervisor si a 1 hora
//      de la reunión sigue sin reconfirmar.
//
//  Regla de cuándo reconfirmar:
//   - Reunión AM (antes de las 12:00, o sin hora): el día HÁBIL anterior,
//     aviso a las 16:00 (fines de semana y feriados de Chile se saltan).
//   - Reunión PM: aviso 3 horas antes, el mismo día.
//
//  Configuración (Apps Script → Configuración del proyecto → Propiedades
//  del script). Todas son opcionales salvo HUBSPOT_TOKEN para la sync:
//   HUBSPOT_TOKEN         Token de app privada del portal HubSpot BullsEye
//                         (scopes: crm.objects.contacts.read,
//                          crm.objects.contacts.write, crm.schemas.contacts.write)
//   SDR_EMAILS            JSON {"Nombre SDR tal como sale en el form":"correo"}
//   SUPERVISOR_EMAIL      Correo(s) para escalamiento, separados por coma
//   CLIENTES_RECONFIRMACION  Clientes a los que aplica, separados por coma
//                         (vacío = todos). Ej: "Simplex"
//   CLIENTE_AVISO_EMAILS  JSON {"Simplex":"correo1,correo2"} para avisar
//                         automáticamente al cliente cuando se reconfirma
//   FERIADOS_EXTRA        Fechas extra no hábiles, "2027-01-01,2027-04-02"
//
//  Instalación del activador: ejecutar UNA VEZ instalarActivadorReconfirmacion()
// ══════════════════════════════════════════════════════════════════════════

const COL_RECONF        = "Reconfirmada";
const COL_RECONF_FECHA  = "Fecha reconfirmación";
const COL_RECONF_ALERTA = "Alerta reconfirmación";
const COL_RECONF_SYNC   = "Sync HubSpot reconfirmación";
const RECONF_COLS       = [COL_RECONF, COL_RECONF_FECHA, COL_RECONF_ALERTA, COL_RECONF_SYNC];

const RECONF_CORTE_AM_PM   = 12;   // reuniones antes de esta hora = AM
const RECONF_HORA_AVISO_AM = 16;   // aviso el día hábil anterior a las 16:00
const RECONF_HORAS_PM      = 3;    // reuniones PM: aviso 3 h antes
const RECONF_HORAS_ESCALA  = 1;    // escalar al supervisor 1 h antes si sigue pendiente
const CAL_FERIADOS_CL      = "es.cl#holiday@group.v.calendar.google.com";

// Feriados de Chile ya conocidos (respaldo si el calendario de Google no responde)
const FERIADOS_CL_BASE = [
  "2026-09-18","2026-09-19","2026-10-12","2026-10-31","2026-11-01",
  "2026-12-08","2026-12-25"
];

const HS_PROP_RECONF       = "reunion_reconfirmada";
const HS_PROP_RECONF_FECHA = "fecha_reconfirmacion_reunion";

// ── Columnas de reconfirmación (se crean al final de la hoja si faltan) ────
function ensureReconfColumns(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => h.toString().trim());
  let next = lastCol + 1;
  RECONF_COLS.forEach(name => {
    if (headers.indexOf(name) < 0) {
      sheet.getRange(1, next).setValue(name)
        .setBackground("#1a1b4d").setFontColor("#ffffff").setFontWeight("bold");
      headers.push(name);
      next++;
    }
  });
  const map = {};
  headers.forEach((h, i) => { map[h] = i + 1; }); // 1-based
  return { headers: headers, col: map };
}

// ── Guardar reconfirmación desde el formulario ────────────────────────────
function actualizarReconfirmacion(sheet, rowNum, p) {
  const info   = ensureReconfColumns(sheet);
  const c      = info.col;
  const row    = sheet.getRange(rowNum, 1, 1, info.headers.length).getValues()[0];
  const get    = (name) => (c[name] ? row[c[name] - 1] : "");
  const yaEra  = _esSi(get(COL_RECONF));
  const marcar = _esSi(p.reconfirmada);

  // ¿Se movió la reunión? Entonces la reconfirmación y la alerta anteriores no valen.
  const seMovio =
    (p.fecha_reunion !== undefined && p.fecha_reunion !== "" && p.fecha_reunion !== p._fechaAnterior) ||
    (p.hora !== undefined && p.hora !== "" && p.hora !== p._horaAnterior);
  if (seMovio) {
    sheet.getRange(rowNum, c[COL_RECONF_ALERTA]).setValue("");
  }

  sheet.getRange(rowNum, c[COL_RECONF]).setValue(marcar ? "Sí" : "No");

  if (!marcar) {
    if (yaEra || seMovio) sheet.getRange(rowNum, c[COL_RECONF_FECHA]).setValue("");
    return;
  }
  if (yaEra && !seMovio) return; // ya estaba reconfirmada: no repetir avisos

  const ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  sheet.getRange(rowNum, c[COL_RECONF_FECHA]).setValue(ahora);

  // Datos de la reunión para HubSpot / aviso al cliente
  const headersL = info.headers.map(h => h.toLowerCase());
  const val = (pred) => { const i = headersL.findIndex(pred); return i >= 0 ? row[i] : ""; };
  const reunion = {
    cliente:  String(val(h => h === "cliente") || ""),
    empresa:  String(val(h => h.includes("empresa")) || ""),
    contacto: String(val(h => h.includes("contacto") && !h.includes("linkedin") && !h.includes("empresa")) || ""),
    cargo:    String(val(h => h === "cargo") || ""),
    correo:   String(val(h => h === "correo") || "").trim(),
    sdr:      String(val(h => h.includes("responsable")) || ""),
    fecha:    p.fecha_reunion || _cellDate(val(h => h.includes("fecha") && h.includes("reuni"))),
    hora:     p.hora || _cellHora(val(h => h === "hora")),
  };

  let sync;
  try {
    sync = syncReconfirmacionHubSpot(reunion);
  } catch (err) {
    sync = "Error: " + err.message;
  }
  sheet.getRange(rowNum, c[COL_RECONF_SYNC]).setValue(sync + " (" + ahora + ")");

  try { avisarClienteReconfirmacion(reunion); } catch (err) { Logger.log("Aviso cliente: " + err); }
}

// ── HubSpot (portal BullsEye) ─────────────────────────────────────────────
function syncReconfirmacionHubSpot(r) {
  const token = PropertiesService.getScriptProperties().getProperty("HUBSPOT_TOKEN");
  if (!token) return "Sin HUBSPOT_TOKEN";
  ensureHubSpotReconfProps(token);

  const id = buscarContactoHubSpot(token, r);
  if (!id) return "Contacto no encontrado en HubSpot";

  const res = _hs(token, "patch", "/crm/v3/objects/contacts/" + id, {
    properties: {
      [HS_PROP_RECONF]:       "Sí",
      [HS_PROP_RECONF_FECHA]: new Date().toISOString(),
    }
  });
  return res.code < 300 ? "OK contacto " + id : "Error HubSpot " + res.code + ": " + res.body.slice(0, 150);
}

function buscarContactoHubSpot(token, r) {
  // 1) Por correo (lo más confiable)
  if (r.correo && r.correo.indexOf("@") > 0) {
    const res = _hs(token, "post", "/crm/v3/objects/contacts/search", {
      filterGroups: [{ filters: [{ propertyName: "email", operator: "EQ", value: r.correo.toLowerCase() }] }],
      limit: 2
    });
    const json = _json(res);
    if (json && json.results && json.results.length) return json.results[0].id;
  }
  // 2) Por nombre del contacto dentro del cliente BullsEye (solo si hay UN resultado)
  if (r.contacto) {
    const body = { query: r.contacto, limit: 5, properties: ["company"] };
    if (r.cliente) {
      body.filterGroups = [{ filters: [{ propertyName: "cliente_bullseye", operator: "EQ", value: r.cliente }] }];
    }
    const json = _json(_hs(token, "post", "/crm/v3/objects/contacts/search", body));
    if (json && json.results) {
      let res = json.results;
      if (res.length > 1 && r.empresa) {
        const emp = r.empresa.toLowerCase();
        res = res.filter(x => (x.properties.company || "").toLowerCase().indexOf(emp) >= 0);
      }
      if (res.length === 1) return res[0].id;
    }
  }
  return null;
}

// Crea las propiedades de contacto si no existen (una sola vez)
function ensureHubSpotReconfProps(token) {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("HS_RECONF_PROPS_OK") === "1") return;
  const defs = [
    { name: HS_PROP_RECONF, label: "Reunión reconfirmada", type: "enumeration", fieldType: "select",
      groupName: "contactinformation",
      description: "El SDR reconfirmó la reunión con el prospecto (formulario de reuniones BullsEye).",
      options: [{ label: "Sí", value: "Sí", displayOrder: 0 }, { label: "No", value: "No", displayOrder: 1 }] },
    { name: HS_PROP_RECONF_FECHA, label: "Fecha reconfirmación reunión", type: "datetime", fieldType: "date",
      groupName: "contactinformation",
      description: "Cuándo el SDR marcó la reunión como reconfirmada." },
  ];
  defs.forEach(d => {
    const res = _hs(token, "post", "/crm/v3/properties/contacts", d);
    if (res.code >= 300 && res.code !== 409) {
      throw new Error("No se pudo crear la propiedad " + d.name + " (" + res.code + "): " + res.body.slice(0, 150));
    }
  });
  props.setProperty("HS_RECONF_PROPS_OK", "1");
}

function _hs(token, method, path, payload) {
  const resp = UrlFetchApp.fetch("https://api.hubapi.com" + path, {
    method: method,
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    payload: payload ? JSON.stringify(payload) : undefined,
    muteHttpExceptions: true,
  });
  return { code: resp.getResponseCode(), body: resp.getContentText() || "" };
}

function _json(res) {
  try { return res.code < 300 ? JSON.parse(res.body) : null; } catch (e) { return null; }
}

// ── Aviso automático al cliente (correo, opcional) ─────────────────────────
function avisarClienteReconfirmacion(r) {
  const destinos = _jsonProp("CLIENTE_AVISO_EMAILS")[r.cliente];
  if (!destinos) return;
  const asunto = "✅ Reunión reconfirmada: " + r.empresa + " — " + r.fecha + (r.hora ? " " + r.hora : "");
  const cuerpo = [
    "Hola,",
    "",
    "Te confirmamos que la siguiente reunión fue reconfirmada con el prospecto:",
    "",
    "🏢 Empresa: " + (r.empresa || "-"),
    "👤 Contacto: " + (r.contacto || "-") + (r.cargo ? " (" + r.cargo + ")" : ""),
    "📅 Fecha: " + (r.fecha || "-") + (r.hora ? " — " + r.hora + " hrs" : ""),
    "🧑‍💼 SDR: " + (r.sdr || "-"),
    "",
    "Equipo BullsEye",
  ].join("\n");
  MailApp.sendEmail(destinos, asunto, cuerpo, { name: "BullsEye — Reuniones" });
}

// ── Recordatorios y escalamiento (activador cada 15 min) ────────────────────
function revisarReconfirmaciones() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(OUTPUT_TAB);
    if (!sheet || sheet.getLastRow() < 2) return;

    const info     = ensureReconfColumns(sheet);
    const c        = info.col;
    const headersL = info.headers.map(h => h.toLowerCase());
    const idx      = (pred) => headersL.findIndex(pred);
    const iCliente  = idx(h => h === "cliente");
    const iSdr      = idx(h => h.includes("responsable"));
    const iEmpresa  = idx(h => h.includes("empresa"));
    const iContacto = idx(h => h.includes("contacto") && !h.includes("linkedin") && !h.includes("empresa"));
    const iFecha    = idx(h => h.includes("fecha") && h.includes("reuni"));
    const iHora     = idx(h => h === "hora");
    const iEstado   = idx(h => h === "realizado");

    const props       = PropertiesService.getScriptProperties();
    const sdrEmails   = _jsonProp("SDR_EMAILS");
    const supervisor  = props.getProperty("SUPERVISOR_EMAIL") || "";
    const soloClientes = (props.getProperty("CLIENTES_RECONFIRMACION") || "")
      .split(",").map(x => x.trim().toLowerCase()).filter(Boolean);

    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, info.headers.length).getValues();
    const now    = new Date();
    const tz     = Session.getScriptTimeZone();
    const avisosSdr = {};   // sdr -> [lineas]
    const escalados = [];   // lineas

    values.forEach((r, k) => {
      const rowNum = k + 2;
      const estado = String(r[iEstado] || "").trim().toLowerCase();
      if (estado !== "" && estado !== "pendiente") return;           // ya realizada / caída
      if (_esSi(r[c[COL_RECONF] - 1])) return;                         // ya reconfirmada
      const cliente = String(r[iCliente] || "").trim();
      if (soloClientes.length && soloClientes.indexOf(cliente.toLowerCase()) < 0) return;

      const inicio = _fechaHoraReunion(r[iFecha], r[iHora]);
      if (!inicio || inicio <= now) return;

      const plazo   = _momentoAvisoReconf(inicio, r[iHora]);
      const escala  = new Date(inicio.getTime() - RECONF_HORAS_ESCALA * 3600 * 1000);
      const alerta  = String(r[c[COL_RECONF_ALERTA] - 1] || "");
      const sdr     = String(r[iSdr] || "").trim();
      const linea   = "• " + cliente + " | " + (r[iEmpresa] || "-") + " — " + (r[iContacto] || "-") +
                      " | " + Utilities.formatDate(inicio, tz, "EEE dd/MM HH:mm") + " hrs";
      const stamp   = Utilities.formatDate(now, tz, "dd/MM HH:mm");

      if (now >= escala && alerta.indexOf("Escalada") !== 0) {
        escalados.push(linea + " | SDR: " + sdr);
        (avisosSdr[sdr] = avisosSdr[sdr] || []).push(linea + "  ⚠️ URGENTE (falta 1 hora)");
        sheet.getRange(rowNum, c[COL_RECONF_ALERTA]).setValue("Escalada " + stamp);
      } else if (now >= plazo && alerta === "") {
        (avisosSdr[sdr] = avisosSdr[sdr] || []).push(linea);
        sheet.getRange(rowNum, c[COL_RECONF_ALERTA]).setValue("Recordatorio " + stamp);
      }
    });

    Object.keys(avisosSdr).forEach(sdr => {
      const email = sdrEmails[sdr];
      if (!email) { Logger.log("Sin correo para SDR: " + sdr); return; }
      MailApp.sendEmail(email,
        "⏰ Reconfirma tus reuniones (" + avisosSdr[sdr].length + ")",
        "Hola " + sdr + ",\n\nEstas reuniones necesitan reconfirmación con el prospecto:\n\n" +
        avisosSdr[sdr].join("\n") +
        "\n\nCuando reconfirmes, márcala como ✅ Reconfirmada en el formulario (Actualizar reunión):\n" +
        "https://bullseye-automate.github.io/Hubspot-y-SDR-Dashboard/\n\nAsí el cliente queda notificado.",
        { name: "BullsEye — Reuniones" });
    });

    if (escalados.length && supervisor) {
      MailApp.sendEmail(supervisor,
        "🚨 Reuniones sin reconfirmar a 1 hora (" + escalados.length + ")",
        "Estas reuniones siguen sin reconfirmar y empiezan en menos de 1 hora:\n\n" + escalados.join("\n"),
        { name: "BullsEye — Reuniones" });
    }
  } finally {
    lock.releaseLock();
  }
}

// Ejecutar UNA VEZ a mano para instalar el activador cada 15 minutos
function instalarActivadorReconfirmacion() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "revisarReconfirmaciones")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("revisarReconfirmaciones").timeBased().everyMinutes(15).create();
  ensureReconfColumns(SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(OUTPUT_TAB));
}

// Prueba manual: muestra en el registro qué avisaría ahora, sin enviar nada
function probarReglasReconfirmacion() {
  const casos = [["13/11/2026", "09:30"], ["16/11/2026", "10:00"], ["13/11/2026", "15:00"], ["08/12/2026", "11:00"], ["09/12/2026", "10:00"]];
  const tz = Session.getScriptTimeZone();
  casos.forEach(([f, h]) => {
    const ini = _fechaHoraReunion(f, h);
    Logger.log(f + " " + h + " → avisar " + Utilities.formatDate(_momentoAvisoReconf(ini, h), tz, "EEE dd/MM HH:mm"));
  });
}

// ── Helpers de fecha ───────────────────────────────────────────────────────
function _momentoAvisoReconf(inicio, horaCell) {
  const hh = _horaNum(horaCell);
  if (hh === null || hh < RECONF_CORTE_AM_PM) {
    const d = _diaHabilAnterior(inicio);
    d.setHours(RECONF_HORA_AVISO_AM, 0, 0, 0);
    return d;
  }
  return new Date(inicio.getTime() - RECONF_HORAS_PM * 3600 * 1000);
}

function _diaHabilAnterior(fecha) {
  const d = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
  do { d.setDate(d.getDate() - 1); } while (!_esHabil(d));
  return d;
}

function _esHabil(d) {
  const dow = d.getDay();
  if (dow === 0 || dow === 6) return false;
  const key = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  return _feriados().indexOf(key) < 0;
}

let _feriadosCache = null;
function _feriados() {
  if (_feriadosCache) return _feriadosCache;
  const list = FERIADOS_CL_BASE.slice();
  (PropertiesService.getScriptProperties().getProperty("FERIADOS_EXTRA") || "")
    .split(",").map(x => x.trim()).filter(Boolean).forEach(x => list.push(x));
  try {
    let cal = CalendarApp.getCalendarById(CAL_FERIADOS_CL);
    if (!cal) cal = CalendarApp.subscribeToCalendar(CAL_FERIADOS_CL, { hidden: true });
    const desde = new Date(); desde.setDate(desde.getDate() - 7);
    const hasta = new Date(); hasta.setDate(hasta.getDate() + 60);
    cal.getEvents(desde, hasta).forEach(ev => {
      list.push(Utilities.formatDate(ev.getStartTime(), Session.getScriptTimeZone(), "yyyy-MM-dd"));
    });
  } catch (e) {
    Logger.log("Calendario de feriados no disponible, uso lista base: " + e);
  }
  _feriadosCache = list;
  return list;
}

function _fechaHoraReunion(fechaCell, horaCell) {
  let d;
  if (fechaCell instanceof Date) {
    d = new Date(fechaCell.getFullYear(), fechaCell.getMonth(), fechaCell.getDate());
  } else {
    const m = String(fechaCell || "").trim().match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
    if (!m) return null;
    d = new Date(+m[3], +m[2] - 1, +m[1]);
  }
  const hh = _horaNum(horaCell);
  const mm = _minNum(horaCell);
  d.setHours(hh === null ? 9 : hh, mm, 0, 0); // sin hora: se asume 09:00 (AM)
  return d;
}

function _horaNum(h) {
  if (h instanceof Date) return h.getHours();
  const m = String(h || "").trim().match(/^(\d{1,2}):(\d{2})/);
  return m ? +m[1] : null;
}
function _minNum(h) {
  if (h instanceof Date) return h.getMinutes();
  const m = String(h || "").trim().match(/^(\d{1,2}):(\d{2})/);
  return m ? +m[2] : 0;
}
function _cellHora(val) {
  if (!val) return "";
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
  return val.toString().trim();
}

function _esSi(v) {
  const s = String(v === true ? "si" : (v || "")).trim().toLowerCase();
  return s === "si" || s === "sí" || s === "true" || s === "1" || s === "yes";
}

function _jsonProp(name) {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty(name) || "{}"); }
  catch (e) { Logger.log("Propiedad " + name + " no es JSON válido"); return {}; }
}
