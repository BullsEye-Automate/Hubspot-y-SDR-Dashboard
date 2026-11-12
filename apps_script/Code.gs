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
//  1. El SDR marca "Reunión reconfirmada" en el formulario (modo Actualizar).
//     → Se guarda en el Sheet (4 columnas al final, se crean solas).
//     → Se actualiza el contacto en HubSpot BullsEye (si hay HUBSPOT_TOKEN).
//  2. Un activador cada 15 min (revisarReconfirmaciones) hace todo lo demás,
//     solo en días hábiles de Chile y solo para clientes activos:
//     · 08:30  Resumen diario a cada SDR y a cada líder de célula: reuniones
//              que hay que reconfirmar hoy y antes de qué hora (+ atrasadas).
//     · Recordatorio puntual al SDR cuando le toca reconfirmar:
//         reunión AM (antes de 13:00 o sin hora) → día hábil anterior 16:00
//         reunión PM (13:00 en adelante)          → mismo día 10:00
//     · Si se cumple el plazo límite sin reconfirmar, alerta al SDR y a su líder:
//         reunión AM → día hábil anterior 18:00 · reunión PM → mismo día 12:00
//       (así el correo al cliente de las 18:15 cubre las AM de mañana y el
//        de las 12:30 cubre las PM de hoy)
//     · 12:30  Correo al cliente: reuniones de esta tarde ya reconfirmadas.
//     · 18:15  Correo al cliente: reuniones del próximo día hábil ya reconfirmadas.
//
//  CONFIGURACIÓN — en el mismo Google Sheet (pestañas que se crean solas):
//   "Reconf - Equipo"   SDR | Correo SDR | Líder de célula | Correo líder
//   "Reconf - Clientes" Cliente | Activo (Sí/No) | Correos aviso diario | Saludo
//  Y en Propiedades del script (solo lo secreto):
//   HUBSPOT_TOKEN  token de app privada del portal HubSpot BullsEye
//                  (scopes: crm.objects.contacts.read/write, crm.schemas.contacts.write)
//   FERIADOS_EXTRA fechas extra no hábiles, "2027-01-01,2027-04-02" (opcional)
//
//  Instalación: ejecutar UNA VEZ instalarActivadorReconfirmacion()
// ══════════════════════════════════════════════════════════════════════════

const COL_RECONF        = "Reconfirmada";
const COL_RECONF_FECHA  = "Fecha reconfirmación";
const COL_RECONF_ALERTA = "Alerta reconfirmación";
const COL_RECONF_SYNC   = "Sync HubSpot reconfirmación";
const RECONF_COLS       = [COL_RECONF, COL_RECONF_FECHA, COL_RECONF_ALERTA, COL_RECONF_SYNC];

const TAB_EQUIPO   = "Reconf - Equipo";
const TAB_CLIENTES = "Reconf - Clientes";
const CLIENTES_PILOTO = ["Simplex"]; // se precargan como activos al crear la pestaña

const RECONF_CORTE_AM_PM    = 13;  // reuniones antes de esta hora = AM
const RECONF_AVISO_AM_HORA  = 16;  // AM: recordatorio el día hábil anterior 16:00
const RECONF_LIMITE_AM_HORA = 18;  // AM: plazo límite el día hábil anterior 18:00
const RECONF_AVISO_PM_HORA  = 10;  // PM: recordatorio el mismo día 10:00
const RECONF_LIMITE_PM_HORA = 12;  // PM: plazo límite el mismo día 12:00

// Envíos diarios [id, hora, minuto]
const RESUMENES = [
  ["equipo",          8, 30],
  ["cliente_tarde",  12, 30],   // reuniones de hoy en la tarde
  ["cliente_manana", 18, 15],   // reuniones del próximo día hábil
];

const CAL_FERIADOS_CL = "es.cl#holiday@group.v.calendar.google.com";
const FERIADOS_CL_BASE = [
  "2026-09-18","2026-09-19","2026-10-12","2026-10-31","2026-11-01",
  "2026-12-08","2026-12-25"
];

const HS_PROP_RECONF       = "reunion_reconfirmada";
const HS_PROP_RECONF_FECHA = "fecha_reconfirmacion_reunion";
const FORM_URL = "https://bullseye-automate.github.io/Hubspot-y-SDR-Dashboard/";

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

// ── Pestañas de configuración ───────────────────────────────────────────────
function ensureConfigTabs(ss) {
  let eq = ss.getSheetByName(TAB_EQUIPO);
  if (!eq) {
    eq = ss.insertSheet(TAB_EQUIPO);
    const hdr = ["SDR", "Correo SDR", "Líder de célula", "Correo líder"];
    eq.getRange(1, 1, 1, hdr.length).setValues([hdr])
      .setBackground("#1a1b4d").setFontColor("#ffffff").setFontWeight("bold");
    eq.setFrozenRows(1);
    try {
      const sdrs = getMaestraData(ss).sdrs || [];
      if (sdrs.length) eq.getRange(2, 1, sdrs.length, 1).setValues(sdrs.map(s => [s]));
    } catch (e) { Logger.log("No se pudo precargar SDRs: " + e); }
    eq.setColumnWidths(1, 4, 220);
  }
  let cl = ss.getSheetByName(TAB_CLIENTES);
  if (!cl) {
    cl = ss.insertSheet(TAB_CLIENTES);
    const hdr = ["Cliente", "Activo (Sí/No)", "Correos aviso diario (separados por coma)", "Saludo (opcional)"];
    cl.getRange(1, 1, 1, hdr.length).setValues([hdr])
      .setBackground("#1a1b4d").setFontColor("#ffffff").setFontWeight("bold");
    cl.setFrozenRows(1);
    cl.getRange(2, 1, CLIENTES_PILOTO.length, 2).setValues(CLIENTES_PILOTO.map(c => [c, "Sí"]));
    cl.getRange(2, 2, 50, 1).setDataValidation(
      SpreadsheetApp.newDataValidation().requireValueInList(["Sí", "No"], true).build());
    cl.setColumnWidths(1, 4, 240);
  }
}

function leerConfigReconf(ss) {
  ensureConfigTabs(ss);
  const equipo = {};   // sdr -> {correo, lider, correoLider}
  ss.getSheetByName(TAB_EQUIPO).getDataRange().getValues().slice(1).forEach(r => {
    const sdr = String(r[0] || "").trim();
    if (!sdr) return;
    equipo[sdr] = {
      correo:      String(r[1] || "").trim(),
      lider:       String(r[2] || "").trim(),
      correoLider: String(r[3] || "").trim(),
    };
  });
  const clientes = {}; // cliente(lower) -> {nombre, correos, saludo}
  ss.getSheetByName(TAB_CLIENTES).getDataRange().getValues().slice(1).forEach(r => {
    const nombre = String(r[0] || "").trim();
    if (!nombre || !_esSi(r[1])) return;
    clientes[nombre.toLowerCase()] = {
      nombre:  nombre,
      correos: String(r[2] || "").trim(),
      saludo:  String(r[3] || "").trim(),
    };
  });
  return { equipo: equipo, clientes: clientes };
}

// ── Guardar reconfirmación desde el formulario ────────────────────────────
function actualizarReconfirmacion(sheet, rowNum, p) {
  const info   = ensureReconfColumns(sheet);
  const c      = info.col;
  const row    = sheet.getRange(rowNum, 1, 1, info.headers.length).getValues()[0];
  const get    = (name) => (c[name] ? row[c[name] - 1] : "");
  const yaEra  = _esSi(get(COL_RECONF));
  const marcar = _esSi(p.reconfirmada);

  // ¿Se movió la reunión? La reconfirmación y las alertas anteriores dejan de valer.
  const seMovio =
    (p.fecha_reunion !== undefined && p.fecha_reunion !== "" && p.fecha_reunion !== p._fechaAnterior) ||
    (p.hora !== undefined && p.hora !== "" && p.hora !== p._horaAnterior);
  if (seMovio) sheet.getRange(rowNum, c[COL_RECONF_ALERTA]).setValue("");

  sheet.getRange(rowNum, c[COL_RECONF]).setValue(marcar ? "Sí" : "No");

  if (!marcar) {
    if (yaEra || seMovio) sheet.getRange(rowNum, c[COL_RECONF_FECHA]).setValue("");
    return;
  }
  if (yaEra && !seMovio) return;

  const ahora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  sheet.getRange(rowNum, c[COL_RECONF_FECHA]).setValue(ahora);

  const r = _reunionDeFila(info.headers, row);
  if (p.fecha_reunion) r.fecha = p.fecha_reunion;
  if (p.hora) r.hora = p.hora;

  let sync;
  try {
    sync = syncReconfirmacionHubSpot(r);
  } catch (err) {
    sync = "Error: " + err.message;
  }
  sheet.getRange(rowNum, c[COL_RECONF_SYNC]).setValue(sync + " (" + ahora + ")");
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
  if (r.correo && r.correo.indexOf("@") > 0) {
    const json = _json(_hs(token, "post", "/crm/v3/objects/contacts/search", {
      filterGroups: [{ filters: [{ propertyName: "email", operator: "EQ", value: r.correo.toLowerCase() }] }],
      limit: 2
    }));
    if (json && json.results && json.results.length) return json.results[0].id;
  }
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

// ══════════════════════════════════════════════════════════════════════════
//  ACTIVADOR CADA 15 MIN
// ══════════════════════════════════════════════════════════════════════════
function revisarReconfirmaciones() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return;
  try {
    const now = new Date();
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ctx = _contextoReconf(ss, now);
    if (!ctx) return;
    _alertasPuntuales(ctx);
    if (_esHabil(now)) _resumenesDiarios(ctx);
    _avisosClienteRezagados(ctx);
  } finally {
    lock.releaseLock();
  }
}

// Lee la hoja y arma la lista de reuniones pendientes de clientes activos
function _contextoReconf(ss, now) {
  const sheet = ss.getSheetByName(OUTPUT_TAB);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const info = ensureReconfColumns(sheet);
  const cfg  = leerConfigReconf(ss);
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, info.headers.length).getValues();

  const reuniones = [];
  values.forEach((row, k) => {
    const r = _reunionDeFila(info.headers, row);
    const estado = r.estado.toLowerCase();
    if (estado !== "" && estado !== "pendiente") return;           // realizada / caída / reagendar
    const cli = cfg.clientes[r.cliente.toLowerCase()];
    if (!cli) return;                                                // cliente no activo
    r.inicio = _fechaHoraReunion(r.fechaCell, r.horaCell);
    if (!r.inicio) return;
    r.rowNum = k + 2;
    r.plazos = _plazosReconf(r.inicio, r.horaCell);
    reuniones.push(r);
  });
  return { ss: ss, sheet: sheet, info: info, cfg: cfg, now: now, reuniones: reuniones,
           tz: Session.getScriptTimeZone() };
}

// ── Recordatorio puntual y alerta de plazo vencido ─────────────────────────
function _alertasPuntuales(ctx) {
  const { sheet, info, cfg, now, tz } = ctx;
  const colAlerta = info.col[COL_RECONF_ALERTA];
  const porSdr = {}, porLider = {};
  const stamp = Utilities.formatDate(now, tz, "dd/MM HH:mm");

  ctx.reuniones.forEach(r => {
    if (r.reconfirmada || r.inicio <= now) return;
    const eq = cfg.equipo[r.sdr] || {};
    if (now >= r.plazos.limite && r.alerta.indexOf("Vencida") !== 0) {
      _push(porSdr, r.sdr, r, "vencida");
      if (eq.correoLider) _push(porLider, eq.correoLider, r, "vencida");
      sheet.getRange(r.rowNum, colAlerta).setValue("Vencida " + stamp);
    } else if (now >= r.plazos.aviso && now < r.plazos.limite && r.alerta === "") {
      _push(porSdr, r.sdr, r, "aviso");
      sheet.getRange(r.rowNum, colAlerta).setValue("Recordatorio " + stamp);
    }
  });

  Object.keys(porSdr).forEach(sdr => {
    const correo = (cfg.equipo[sdr] || {}).correo;
    if (!correo) { Logger.log("Sin correo para SDR: " + sdr); return; }
    const items  = porSdr[sdr];
    const vencen = items.filter(x => x.tipo === "vencida").length;
    _enviar(correo,
      vencen ? "🚨 Reuniones sin reconfirmar — plazo vencido (" + vencen + ")"
             : "⏰ Te toca reconfirmar (" + items.length + ")",
      "Hola " + _primerNombre(sdr) + ",",
      vencen ? "Estas reuniones pasaron su plazo de reconfirmación. Reconfírmalas ahora:"
             : "Es momento de reconfirmar estas reuniones con el prospecto:",
      items.map(x => x.r), ctx, { sdr: false },
      "Cuando reconfirmes, márcala en el formulario (Actualizar reunión): " + FORM_URL);
  });

  Object.keys(porLider).forEach(correo => {
    _enviar(correo,
      "🚨 Reuniones de tu equipo sin reconfirmar (" + porLider[correo].length + ")",
      "Hola,",
      "Estas reuniones pasaron su plazo de reconfirmación y siguen pendientes:",
      porLider[correo].map(x => x.r), ctx, { sdr: true }, "");
  });
}

// ── Resúmenes diarios (una vez por día y por envío) ─────────────────────────
function _resumenesDiarios(ctx) {
  const props = PropertiesService.getScriptProperties();
  const hoy   = Utilities.formatDate(ctx.now, ctx.tz, "yyyy-MM-dd");
  RESUMENES.forEach(([id, hh, mm]) => {
    const desde = new Date(ctx.now); desde.setHours(hh, mm, 0, 0);
    const hasta = new Date(desde.getTime() + 3 * 3600 * 1000);
    const flag  = "RESUMEN_" + id + "_" + hoy;
    if (ctx.now < desde || ctx.now >= hasta || props.getProperty(flag)) return;
    props.setProperty(flag, "1");
    if (id === "equipo")         { _resumenEquipo(ctx); _limpiarMarcasAntiguas(ctx.now); }
    if (id === "cliente_tarde")  _resumenCliente(ctx, "tarde");
    if (id === "cliente_manana") _resumenCliente(ctx, "manana");
  });
}

// 08:30 — a cada SDR y a cada líder: qué reconfirmar hoy y antes de qué hora
function _resumenEquipo(ctx) {
  const { cfg, now } = ctx;
  const finHoy = new Date(now); finHoy.setHours(23, 59, 59, 999);
  const lista = ctx.reuniones
    .filter(r => !r.reconfirmada && r.inicio > now && r.plazos.limite <= finHoy)
    .sort((a, b) => a.plazos.limite - b.plazos.limite);

  const porSdr = {}, porLider = {};
  lista.forEach(r => {
    (porSdr[r.sdr] = porSdr[r.sdr] || []).push(r);
    const eq = cfg.equipo[r.sdr] || {};
    if (eq.correoLider) {
      const k = eq.correoLider;
      (porLider[k] = porLider[k] || { nombre: eq.lider, items: [] }).items.push(r);
    }
  });

  Object.keys(porSdr).forEach(sdr => {
    const correo = (cfg.equipo[sdr] || {}).correo;
    if (!correo) { Logger.log("Sin correo para SDR: " + sdr); return; }
    _enviar(correo, "📋 Reuniones para reconfirmar hoy (" + porSdr[sdr].length + ")",
      "Hola " + _primerNombre(sdr) + ",",
      "Estas son las reuniones que debes reconfirmar hoy, ordenadas por plazo:",
      porSdr[sdr], ctx, { sdr: false },
      "Al reconfirmar, márcala en el formulario (Actualizar reunión): " + FORM_URL);
  });

  Object.keys(porLider).forEach(correo => {
    const g = porLider[correo];
    _enviar(correo, "📋 Reconfirmaciones de tu célula hoy (" + g.items.length + ")",
      "Hola " + _primerNombre(g.nombre) + ",",
      "Estas son las reuniones que tu equipo debe reconfirmar hoy, ordenadas por plazo:",
      g.items, ctx, { sdr: true }, "");
  });
}

// 12:30 / 18:15 — al cliente: reuniones ya reconfirmadas
function _resumenCliente(ctx, cual) {
  const { cfg, now, tz } = ctx;
  let dia, titulo, intro;
  if (cual === "tarde") {
    dia    = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    titulo = "Reuniones reconfirmadas para esta tarde";
    intro  = "Te confirmamos que estas reuniones de hoy en la tarde ya fueron reconfirmadas con el prospecto:";
  } else {
    dia    = _diaHabilSiguiente(now);
    titulo = "Reuniones reconfirmadas para el " + Utilities.formatDate(dia, tz, "dd/MM");
    intro  = "Te confirmamos que estas reuniones del " + _nombreDia(dia) + " ya fueron reconfirmadas con el prospecto:";
  }
  const mismoDia = (d) => d.getFullYear() === dia.getFullYear() && d.getMonth() === dia.getMonth() && d.getDate() === dia.getDate();

  Object.keys(cfg.clientes).forEach(k => {
    const cli = cfg.clientes[k];
    if (!cli.correos) return;
    let delDia = ctx.reuniones.filter(r => r.cliente.toLowerCase() === k && mismoDia(r.inicio) && r.inicio > now);
    if (cual === "tarde") delDia = delDia.filter(r => r.inicio.getHours() >= RECONF_CORTE_AM_PM);
    const ok = delDia.filter(r => r.reconfirmada).sort((a, b) => a.inicio - b.inicio);
    if (!ok.length) return;
    const pendientes = delDia.length - ok.length;
    ok.forEach(r => _marcarAvisoCliente(r));
    const nota = cual === "manana"
      ? "Las reuniones de la tarde se reconfirman el mismo día; te avisaremos a las 12:30."
      : "";
    _enviar(cli.correos, "✅ " + titulo + " — " + cli.nombre,
      cli.saludo || "Hola,",
      intro, ok, ctx, { sdr: true, cliente: true },
      (pendientes ? pendientes + (pendientes === 1 ? " reunión más está" : " reuniones más están") +
        " en proceso de reconfirmación. " : "") + nota);
  });
}

// Reconfirmaciones que llegan DESPUÉS del correo diario (12:30 / 18:15):
// se avisan al cliente en la siguiente pasada (≤15 min), una por una.
function _avisosClienteRezagados(ctx) {
  const { cfg, now, tz } = ctx;
  ctx.reuniones.forEach(r => {
    if (!r.reconfirmada || r.inicio <= now || _yaAvisadoCliente(r)) return;
    const cli = cfg.clientes[r.cliente.toLowerCase()];
    if (!cli || !cli.correos) return;
    // ¿Ya pasó el correo diario que debía incluirla?
    let resumen;
    if (r.inicio.getHours() >= RECONF_CORTE_AM_PM) {
      resumen = new Date(r.inicio); resumen.setHours(12, 30, 0, 0);
    } else {
      resumen = _diaHabilAnterior(r.inicio); resumen.setHours(18, 15, 0, 0);
    }
    if (now < resumen) return;
    _enviar(cli.correos, "✅ Reunión reconfirmada: " + r.empresa + " — " +
        Utilities.formatDate(r.inicio, tz, "dd/MM HH:mm") + " hrs",
      cli.saludo || "Hola,",
      "Te confirmamos que esta reunión ya fue reconfirmada con el prospecto:",
      [r], ctx, { sdr: true, cliente: true }, "");
    _marcarAvisoCliente(r);
  });
}

function _claveAvisoCliente(r) {
  return "AVCLI_" + r.rowNum + "_" + Utilities.formatDate(r.inicio, Session.getScriptTimeZone(), "yyyyMMddHHmm");
}
function _yaAvisadoCliente(r) {
  return !!PropertiesService.getScriptProperties().getProperty(_claveAvisoCliente(r));
}
function _marcarAvisoCliente(r) {
  PropertiesService.getScriptProperties().setProperty(_claveAvisoCliente(r), "1");
}

// Borra marcas de avisos/resúmenes de más de 14 días (se llama en el resumen de las 08:30)
function _limpiarMarcasAntiguas(now) {
  const props = PropertiesService.getScriptProperties();
  const corte = new Date(now); corte.setDate(corte.getDate() - 14);
  const lim   = Utilities.formatDate(corte, Session.getScriptTimeZone(), "yyyyMMdd");
  Object.keys(props.getProperties()).forEach(k => {
    let m = k.match(/^AVCLI_\d+_(\d{8})/) || k.match(/^RESUMEN_.+_(\d{4})-(\d{2})-(\d{2})$/);
    if (!m) return;
    const f = m.length === 2 ? m[1] : m[1] + m[2] + m[3];
    if (f < lim) props.deleteProperty(k);
  });
}

// ── Correo HTML con tabla de reuniones ─────────────────────────────────────
function _enviar(destinos, asunto, saludo, intro, reuniones, ctx, opt, pie) {
  const tz  = ctx.tz;
  const th  = 'style="text-align:left;padding:8px 10px;background:#1a1b4d;color:#fff;font-size:13px"';
  const td  = 'style="padding:8px 10px;border-bottom:1px solid #e3e5ef;font-size:13px;vertical-align:top"';
  const filas = reuniones.map(r => {
    const reunion = Utilities.formatDate(r.inicio, tz, "EEE dd/MM HH:mm") + " hrs";
    const plazo   = opt.cliente ? "" :
      "<td " + td + "><b>" + Utilities.formatDate(r.plazos.limite, tz, "EEE dd/MM HH:mm") + "</b>" +
      (ctx.now >= r.plazos.limite ? ' <span style="color:#d32f2f">(vencido)</span>' : "") + "</td>";
    return "<tr>" +
      "<td " + td + "><b>" + _esc(r.empresa) + "</b>" + (opt.cliente ? "" : "<br><span style='color:#667'>" + _esc(r.cliente) + "</span>") + "</td>" +
      "<td " + td + ">" + _esc(r.contacto) + (r.cargo ? "<br><span style='color:#667'>" + _esc(r.cargo) + "</span>" : "") + "</td>" +
      "<td " + td + ">" + reunion + "</td>" +
      plazo +
      (opt.sdr ? "<td " + td + ">" + _esc(r.sdr) + "</td>" : "") +
      "</tr>";
  }).join("");
  const head = "<tr><th " + th + ">Empresa</th><th " + th + ">Contacto</th><th " + th + ">Reunión</th>" +
    (opt.cliente ? "" : "<th " + th + ">Reconfirmar antes de</th>") +
    (opt.sdr ? "<th " + th + ">SDR</th>" : "") + "</tr>";
  const html =
    '<div style="font-family:Arial,Helvetica,sans-serif;color:#1d1f33;max-width:720px">' +
    "<p>" + _esc(saludo) + "</p><p>" + _esc(intro) + "</p>" +
    '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%">' + head + filas + "</table>" +
    (pie ? '<p style="color:#555;font-size:13px">' + _esc(pie) + "</p>" : "") +
    '<p style="color:#888;font-size:12px">Equipo BullsEye</p></div>';
  const texto = saludo + "\n\n" + intro + "\n\n" + reuniones.map(r =>
    "• " + r.empresa + " — " + r.contacto + " | " + Utilities.formatDate(r.inicio, tz, "EEE dd/MM HH:mm") +
    (opt.cliente ? "" : " | reconfirmar antes de " + Utilities.formatDate(r.plazos.limite, tz, "dd/MM HH:mm")) +
    (opt.sdr ? " | SDR: " + r.sdr : "")).join("\n") + (pie ? "\n\n" + pie : "");
  MailApp.sendEmail(destinos, asunto, texto, { htmlBody: html, name: "BullsEye — Reuniones" });
}

// ══════════════════════════════════════════════════════════════════════════
//  INSTALACIÓN Y PRUEBAS
// ══════════════════════════════════════════════════════════════════════════

// Ejecutar UNA VEZ: crea pestañas de configuración, columnas y el activador
function instalarActivadorReconfirmacion() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === "revisarReconfirmaciones")
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("revisarReconfirmaciones").timeBased().everyMinutes(15).create();
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  ensureReconfColumns(ss.getSheetByName(OUTPUT_TAB));
  ensureConfigTabs(ss);
}

// Muestra en el registro qué se enviaría ahora, SIN enviar ni marcar nada
function simularReconfirmaciones() {
  const ctx = _contextoReconf(SpreadsheetApp.openById(SPREADSHEET_ID), new Date());
  if (!ctx) { Logger.log("Sin reuniones"); return; }
  ctx.reuniones.forEach(r => Logger.log([
    r.cliente, r.empresa, r.sdr,
    "reunión " + Utilities.formatDate(r.inicio, ctx.tz, "dd/MM HH:mm"),
    "aviso " + Utilities.formatDate(r.plazos.aviso, ctx.tz, "dd/MM HH:mm"),
    "límite " + Utilities.formatDate(r.plazos.limite, ctx.tz, "dd/MM HH:mm"),
    r.reconfirmada ? "RECONFIRMADA" : "pendiente",
  ].join(" | ")));
}

// ══════════════════════════════════════════════════════════════════════════
//  HELPERS
// ══════════════════════════════════════════════════════════════════════════
function _reunionDeFila(headers, row) {
  const hl  = headers.map(h => String(h).trim().toLowerCase());
  const at  = (pred) => { const i = hl.findIndex(pred); return i >= 0 ? row[i] : ""; };
  const str = (v) => String(v === null || v === undefined ? "" : v).trim();
  const iR  = hl.indexOf(COL_RECONF.toLowerCase());
  const iA  = hl.indexOf(COL_RECONF_ALERTA.toLowerCase());
  return {
    cliente:   str(at(h => h === "cliente")),
    sdr:       str(at(h => h.includes("responsable"))),
    empresa:   str(at(h => h.includes("empresa"))),
    contacto:  str(at(h => h.includes("contacto") && !h.includes("linkedin") && !h.includes("empresa"))),
    cargo:     str(at(h => h === "cargo")),
    correo:    str(at(h => h === "correo")),
    estado:    str(at(h => h === "realizado")),
    fechaCell: at(h => h.includes("fecha") && h.includes("reuni")),
    horaCell:  at(h => h === "hora"),
    fecha:     _cellDate(at(h => h.includes("fecha") && h.includes("reuni"))),
    hora:      _cellHora(at(h => h === "hora")),
    reconfirmada: iR >= 0 ? _esSi(row[iR]) : false,
    alerta:    iA >= 0 ? str(row[iA]) : "",
  };
}

function _push(obj, key, r, tipo) { (obj[key] = obj[key] || []).push({ r: r, tipo: tipo }); }

// aviso = cuándo recordarle al SDR; limite = plazo para reconfirmar
function _plazosReconf(inicio, horaCell) {
  const hh = _horaNum(horaCell);
  if (hh === null || hh < RECONF_CORTE_AM_PM) {
    const d = _diaHabilAnterior(inicio);
    const aviso  = new Date(d); aviso.setHours(RECONF_AVISO_AM_HORA, 0, 0, 0);
    const limite = new Date(d); limite.setHours(RECONF_LIMITE_AM_HORA, 0, 0, 0);
    return { aviso: aviso, limite: limite };
  }
  const aviso  = new Date(inicio); aviso.setHours(RECONF_AVISO_PM_HORA, 0, 0, 0);
  const limite = new Date(inicio); limite.setHours(RECONF_LIMITE_PM_HORA, 0, 0, 0);
  return { aviso: aviso, limite: limite };
}

function _diaHabilAnterior(fecha) {
  const d = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
  do { d.setDate(d.getDate() - 1); } while (!_esHabil(d));
  return d;
}

function _diaHabilSiguiente(fecha) {
  const d = new Date(fecha.getFullYear(), fecha.getMonth(), fecha.getDate());
  do { d.setDate(d.getDate() + 1); } while (!_esHabil(d));
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
  d.setHours(hh === null ? 9 : hh, _minNum(horaCell), 0, 0); // sin hora: se asume 09:00 (AM)
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

function _nombreDia(d) {
  return ["domingo","lunes","martes","miércoles","jueves","viernes","sábado"][d.getDay()] +
    " " + String(d.getDate()).padStart(2, "0") + "/" + String(d.getMonth() + 1).padStart(2, "0");
}
function _primerNombre(n) { return String(n || "").trim().split(/\s+/)[0] || ""; }
function _esc(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
function _esSi(v) {
  const s = String(v === true ? "si" : (v || "")).trim().toLowerCase();
  return s === "si" || s === "sí" || s === "true" || s === "1" || s === "yes";
}
