const SHEET_NAME = "storecheckdb";
const CATALOGO_SHEET_NAME = "catalogo";

// Cambia esto por una clave tuya.
const TOKEN = "10152021";
const AUDIT_START_WEEK = "2026-W20";

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action).trim() : "";

    if (action === "cadenas") {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(CATALOGO_SHEET_NAME);
      if (!sh) {
        return json_({ ok: false, error: `No existe pestaña ${CATALOGO_SHEET_NAME}` });
      }

      return json_({ ok: true, cadenas: obtenerCadenasCatalogo_(sh) });
    }

    if (action === "catalogo" || action === "catalogoCompleto") {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(CATALOGO_SHEET_NAME);
      if (!sh) {
        return json_({ ok: false, error: `No existe pestaña ${CATALOGO_SHEET_NAME}` });
      }

      const cadena = action === "catalogo" ? texto_(e.parameter.cadena) : "";
      if (action === "catalogo" && !cadena) {
        return json_({ ok: false, error: "Selecciona una cadena" });
      }

      return json_({ ok: true, cadena, rows: obtenerFilasCatalogo_(sh, cadena) });
    }

    if (action === "dashboardSemanal") {
      return obtenerDashboardSemanal_(e.parameter || {});
    }

    if (action === "ultimosPrecios") {
      const cadena = texto_(e.parameter.cadena);
      const estado = texto_(e.parameter.estado);
      if (!cadena || !estado) {
        return json_({ ok: false, error: "Selecciona cadena y estado" });
      }
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(SHEET_NAME);
      if (!sh) return json_({ ok: false, error: `No existe pestaña ${SHEET_NAME}` });
      return json_({
        ok: true,
        cadena,
        estado,
        precios: obtenerUltimosPreciosRegulares_(sh, cadena, estado)
      });
    }

    return ContentService
      .createTextOutput("OK")
      .setMimeType(ContentService.MimeType.TEXT);

  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json_({ ok: false, error: "Sin payload" });
    }

    const raw = e.postData.contents;
    const data = JSON.parse(raw);

    if (!data.token || data.token !== TOKEN) {
      return json_({ ok: false, error: "Token inválido" });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) return json_({ ok: false, error: `No existe pestaña ${SHEET_NAME}` });

    const ts = new Date();
    const cadena = data.cadena || "";
    const estado = data.estado || "";
    const familia = data.familia || "";

    const items = Array.isArray(data.items) ? data.items : [];
    if (items.length === 0) {
      return json_({ ok: false, error: "items vacío" });
    }

    const rows = items.map(it => ([
      ts,
      cadena,
      estado,
      familia,
      it.codigo || "",
      it.prod || "",
      it.marca || "",
      it.precioRegular || "",
      it.pr || it.precioPromocion || "",
      it.promo || "",
    ]));

    sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    CacheService.getScriptCache().remove(`dashboard-v3-${semanaIso_(ts)}`);

    return json_({ ok: true, inserted: rows.length });

  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function obtenerEncabezadosCatalogo_(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (!lastColumn || !sheet.getLastRow()) return { headers: [], cadenaIndex: -1 };
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0].map(texto_);
  const cadenaIndex = headers.findIndex(h => normalizarEncabezado_(h) === "cadena");
  return { headers, cadenaIndex };
}

function obtenerCadenasCatalogo_(sheet) {
  const lastRow = sheet.getLastRow();
  const meta = obtenerEncabezadosCatalogo_(sheet);
  if (lastRow < 2 || meta.cadenaIndex < 0) return [];

  const values = sheet.getRange(2, meta.cadenaIndex + 1, lastRow - 1, 1).getDisplayValues();
  return [...new Set(values.map(r => texto_(r[0])).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b));
}

function obtenerFilasCatalogo_(sheet, cadena) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  const meta = obtenerEncabezadosCatalogo_(sheet);
  if (lastRow < 2 || !lastColumn || meta.cadenaIndex < 0) return [];

  let rows = [];
  if (!cadena) {
    rows = sheet.getRange(2, 1, lastRow - 1, lastColumn).getDisplayValues();
  } else {
    // Primero se lee solamente la columna Cadena. Después se consultan las filas
    // coincidentes en bloques contiguos para no transportar el catálogo completo.
    const cadenas = sheet.getRange(2, meta.cadenaIndex + 1, lastRow - 1, 1).getDisplayValues();
    const indices = [];
    cadenas.forEach((r, i) => {
      if (texto_(r[0]) === cadena) indices.push(i);
    });

    for (let i = 0; i < indices.length;) {
      const inicio = indices[i];
      let fin = inicio;
      while (i + 1 < indices.length && indices[i + 1] === fin + 1) {
        i++;
        fin = indices[i];
      }
      rows = rows.concat(sheet.getRange(2 + inicio, 1, fin - inicio + 1, lastColumn).getDisplayValues());
      i++;
    }
  }

  return rows
    .filter(r => r.some(c => texto_(c)))
    .map(r => {
      const obj = {};
      meta.headers.forEach((h, i) => obj[h] = texto_(r[i]));
      return obj;
    });
}

function obtenerDashboardSemanal_(params) {
  const semana = String(params.week || semanaIsoActual_());
  if (!/^\d{4}-W\d{2}$/.test(semana) || semana < AUDIT_START_WEEK) {
    return json_({ ok: false, error: `Selecciona una semana desde ${AUDIT_START_WEEK}` });
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = `dashboard-v3-${semana}`;
  const cached = cache.get(cacheKey);
  if (cached) return json_(JSON.parse(cached));

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const lectura = leerCapturasSemana_(ss.getSheetByName(SHEET_NAME), semana);
  const resultado = {
    ok: true,
    version: "dashboard-v3",
    semana,
    inicio: AUDIT_START_WEEK,
    diagnostico: { filasCaptura: lectura.total, filasEnSemana: lectura.rows.length },
    capturas: lectura.rows
      .sort((a, b) => b.fecha - a.fecha)
      .map(r => ({
        cadena: r.cadena, estado: r.estado, codigo: r.codigo,
        fecha: Utilities.formatDate(r.fecha, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")
      }))
  };

  const serialized = JSON.stringify(resultado);
  if (serialized.length < 90000) cache.put(cacheKey, serialized, 300);
  return json_(resultado);
}

function leerFilas_(sheet) {
  if (!sheet) return { headers: [], rows: [] };
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return { headers: [], rows: [] };
  const headers = values[0].map(texto_);
  return {
    headers,
    rows: values.slice(1).filter(r => r.some(c => texto_(c))).map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = r[i]);
      return obj;
    })
  };
}

function obtenerUltimosPreciosRegulares_(sheet, cadena, estado) {
  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(10, sheet.getLastColumn());
  if (lastRow < 2) return [];

  const primeraFila = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(texto_);
  const tieneEncabezados = primeraFila.some(v => /^(fecha|timestamp)$/i.test(v)) &&
    primeraFila.some(v => /^cadena$/i.test(v));
  const columnas = obtenerColumnasCaptura_(primeraFila, tieneEncabezados);
  const dataStartRow = tieneEncabezados ? 2 : 1;
  const total = lastRow - dataStartRow + 1;
  if (total <= 0) return [];

  const values = sheet.getRange(dataStartRow, 1, total, lastColumn).getValues();
  const encontrados = new Set();
  const precios = [];

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    if (texto_(row[columnas.cadena]) !== cadena || texto_(row[columnas.estado]) !== estado) continue;
    const codigo = texto_(row[columnas.codigo]);
    const precioRegular = texto_(row[columnas.precioRegular]);
    if (!codigo || !precioRegular || encontrados.has(codigo)) continue;
    encontrados.add(codigo);
    const fecha = normalizarFecha_(row[columnas.fecha]);
    precios.push({
      codigo,
      precioRegular,
      fecha: fecha ? Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm") : ""
    });
  }

  return precios;
}

function obtenerColumnasCaptura_(primeraFila, tieneEncabezados) {
  const indice = nombre => {
    if (!tieneEncabezados) return nombre.fallback;
    const aliases = nombre.aliases.map(normalizarEncabezado_);
    const encontrado = primeraFila.findIndex(h => aliases.includes(normalizarEncabezado_(h)));
    return encontrado >= 0 ? encontrado : nombre.fallback;
  };
  return {
    fecha: indice({ aliases: ["Timestamp", "Fecha"], fallback: 0 }),
    cadena: indice({ aliases: ["Cadena", "Tienda"], fallback: 1 }),
    estado: indice({ aliases: ["Estado"], fallback: 2 }),
    familia: indice({ aliases: ["Familia"], fallback: 3 }),
    codigo: indice({ aliases: ["Codigo Producto", "Código Producto", "Codigo"], fallback: 4 }),
    producto: indice({ aliases: ["Producto"], fallback: 5 }),
    marca: indice({ aliases: ["Marca"], fallback: 6 }),
    precioRegular: indice({ aliases: ["Precio Regular"], fallback: 7 }),
    precio: indice({ aliases: ["Precio Promoción/Paquete", "Precio Promocion/Paquete", "Precio"], fallback: 8 }),
    promocion: indice({ aliases: ["Promoción", "Promocion"], fallback: 9 })
  };
}

function leerCapturasSemana_(sheet, semana) {
  if (!sheet) return { total: 0, rows: [] };
  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(10, sheet.getLastColumn());
  if (!lastRow) return { total: 0, rows: [] };

  const primeraFila = sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map(texto_);
  const tieneEncabezados = primeraFila.some(v => /^(fecha|timestamp)$/i.test(v)) &&
    primeraFila.some(v => /^cadena$/i.test(v));
  const columnas = obtenerColumnasCaptura_(primeraFila, tieneEncabezados);
  const dataStartRow = tieneEncabezados ? 2 : 1;
  const total = Math.max(0, lastRow - dataStartRow + 1);
  if (!total) return { total: 0, rows: [] };

  // Primera lectura: solamente Timestamp para localizar el bloque semanal.
  const fechas = sheet.getRange(dataStartRow, columnas.fecha + 1, total, 1).getValues();
  const indicesSemana = [];
  fechas.forEach((r, i) => {
    const fecha = normalizarFecha_(r[0]);
    if (fecha && semanaIso_(fecha) === semana) indicesSemana.push(i);
  });
  if (!indicesSemana.length) return { total, rows: [] };

  // Segunda lectura: únicamente las filas comprendidas entre la primera y última coincidencia.
  const firstIndex = indicesSemana[0];
  const lastIndex = indicesSemana[indicesSemana.length - 1];
  const blockStartRow = dataStartRow + firstIndex;
  const blockLength = lastIndex - firstIndex + 1;
  const values = sheet.getRange(blockStartRow, 1, blockLength, lastColumn).getValues();

  const rows = values
    .map(r => ({
      fecha: normalizarFecha_(r[columnas.fecha]),
      cadena: texto_(r[columnas.cadena]),
      estado: texto_(r[columnas.estado]),
      familia: texto_(r[columnas.familia]),
      codigo: texto_(r[columnas.codigo]),
      producto: texto_(r[columnas.producto]),
      marca: texto_(r[columnas.marca]),
      precioRegular: texto_(r[columnas.precioRegular]),
      precio: texto_(r[columnas.precio]),
      promocion: texto_(r[columnas.promocion])
    }))
    .filter(r => r.fecha && semanaIso_(r.fecha) === semana && r.cadena && r.codigo);

  return { total, rows };
}

function normalizarEncabezado_(value) {
  return texto_(value).toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function normalizarFecha_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const text = texto_(value);
  if (!text) return null;
  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) return parsed;
  const match = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (!match) return null;
  return new Date(Number(match[3]), Number(match[2]) - 1, Number(match[1]), Number(match[4] || 0), Number(match[5] || 0), Number(match[6] || 0));
}

function texto_(value) { return String(value == null ? "" : value).trim(); }
function claveCombinacion_(cadena, codigo) { return `${cadena}::${codigo}`; }

function semanaIso_(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const day = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - day);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  const week = Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
  return `${d.getUTCFullYear()}-W${String(week).padStart(2, "0")}`;
}

function semanaIsoActual_() { return semanaIso_(new Date()); }
