// ╔══════════════════════════════════════════════════════════╗
// ║   GOOGLE APPS SCRIPT — Alexander D&C v3.2                 ║
// ║   Cambios respecto a v3.1:                                 ║
// ║   3) CLAVE DE ACCESO: doGet y doPost ahora exigen un       ║
// ║      "token" que debe coincidir con ACCESS_TOKEN guardado  ║
// ║      en las Propiedades del Script. Sin la clave correcta, ║
// ║      el backend no entrega ni modifica ningún dato — no es ║
// ║      solo una pantalla de contraseña del lado del cliente. ║
// ║      CONFIGURA TU CLAVE: edita configurarClaveAcceso() más ║
// ║      abajo con tu clave, y ejecútala UNA vez desde el      ║
// ║      editor de Apps Script (▶ Ejecutar).                   ║
// ╚══════════════════════════════════════════════════════════╝

const HOJAS = {
  catalogo:     "Catalogo",
  clientes:     "Clientes",
  cuentas:      "Cuentas",
  cotizaciones: "Cotizaciones",
  reportes:     "Reportes",   // NUEVO
  config:       "Config"
};

const FOLDER_ID = "1ezV1YBMop_MS1sYyE-ERUcGtqiu33ztn";

// ── CLAVE DE ACCESO ───────────────────────────────────────────
// Ejecuta esta función UNA sola vez desde el editor de Apps Script (selecciona
// "configurarClaveAcceso" en el desplegable de funciones y presiona ▶ Ejecutar)
// para guardar tu clave. Después de ejecutarla puedes borrar la clave de acá si
// quieres, ya queda guardada de forma segura en las Propiedades del Script.
function configurarClaveAcceso() {
  PropertiesService.getScriptProperties().setProperty('ACCESS_TOKEN', 'CAMBIA_ESTA_CLAVE');
  Logger.log('Clave de acceso configurada.');
}

// Compara el token recibido contra el guardado. Si nunca configuraste una clave
// (ACCESS_TOKEN vacío), el sistema queda abierto como antes — no bloquea por error
// a nadie que todavía no haya corrido configurarClaveAcceso().
function verificarToken(tokenRecibido) {
  const esperado = PropertiesService.getScriptProperties().getProperty('ACCESS_TOKEN');
  if (!esperado) return true;
  return String(tokenRecibido || '') === String(esperado);
}

function respuestaNoAutorizado(cb) {
  const errData = { ok: false, error: "Clave de acceso incorrecta." };
  if (cb) {
    return ContentService
      .createTextOutput(cb + "(" + JSON.stringify(errData) + ")")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(errData))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET ────────────────────────────────────────────────────────
function doGet(e) {
  try {
    const cb = e && e.parameter && e.parameter.callback;
    const token = e && e.parameter && e.parameter.token;
    if (!verificarToken(token)) return respuestaNoAutorizado(cb);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const datos = {
      ok:           true,
      catalogo:     leerHoja(ss, HOJAS.catalogo),
      clientes:     leerHoja(ss, HOJAS.clientes),
      cuentas:      leerHoja(ss, HOJAS.cuentas),
      cotizaciones: leerHoja(ss, HOJAS.cotizaciones),
      reportes:     leerHoja(ss, HOJAS.reportes),   // NUEVO
      correlativo:  leerConfig(ss, "correlativo", 1)
    };
    if (cb) {
      return ContentService
        .createTextOutput(cb + "(" + JSON.stringify(datos) + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService
      .createTextOutput(JSON.stringify(datos))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── POST ───────────────────────────────────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (!verificarToken(payload.token)) return respuestaNoAutorizado(null);

    const ss      = SpreadsheetApp.getActiveSpreadsheet();

    if (payload.tipo === "guardarCotizacionConPdf") {
      return manejarGuardarPdfCotizacion(payload);
    }

    if (payload.tipo === "guardarFotoGaleria") {
      return manejarGuardarFotoGaleria(payload);
    }

    // NUEVO: respaldo en PDF de los Reportes de Obra (para "guardar mis reportes en mi Drive")
    if (payload.tipo === "guardarReportePdf") {
      return manejarGuardarPdfReporte(payload);
    }

    if (payload.tipo === "eliminarArchivoDrive") {
      return manejarEliminarArchivoDrive(payload);
    }

    // ── GUARDAR HOJAS NORMALES (incluye la nueva tabla "Reportes") ──
    var tabla = payload.tabla;
    var datos = payload.datos;

    if (!tabla) {
      return ContentService
        .createTextOutput(JSON.stringify({ error: "Petición sin 'tabla' ni 'tipo' reconocido: " + JSON.stringify(payload.tipo || null) }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (tabla === "Config") {
      if (Array.isArray(datos)) {
        datos.forEach(function(item) { guardarConfig(ss, item.key, item.value); });
      }
    } else {
      guardarHoja(ss, tabla, datos);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Cotizaciones: generar el PDF EXACTO que ve el usuario en pantalla ──────
function manejarGuardarPdfCotizacion(payload) {
  const carpeta = DriveApp.getFolderById(FOLDER_ID);

  let htmlCuerpo = payload.htmlFull;
  if (!htmlCuerpo) {
    htmlCuerpo = construirHtmlBasicoCotizacion(payload);
  }

  const archivoHtml  = Utilities.newBlob(htmlCuerpo, "text/html", "temp.html");
  const pdfBlob      = archivoHtml.getAs("application/pdf");
  const nombreLimpio = limpiarNombreArchivo(payload.cliente || "Sin_Nombre");
  pdfBlob.setName("Cotizacion_" + (payload.numero || "") + "_" + nombreLimpio + ".pdf");

  if (payload.pdfFileIdAnterior) {
    try { DriveApp.getFileById(payload.pdfFileIdAnterior).setTrashed(true); }
    catch (errBorrar) { /* ya no existe o no se encuentra: seguimos igual */ }
  }

  const nuevoPdf = carpeta.createFile(pdfBlob);
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, url: nuevoPdf.getUrl(), fileId: nuevoPdf.getId() }))
    .setMimeType(ContentService.MimeType.JSON);
}

function construirHtmlBasicoCotizacion(payload) {
  let filasHtml = "";
  if (payload.secciones) {
    const secciones = typeof payload.secciones === "string" ? JSON.parse(payload.secciones) : payload.secciones;
    secciones.forEach(function(sec) {
      if (sec.nombre) {
        filasHtml += "<tr><td colspan='4' style='padding:8px; font-weight:bold; font-size:12px; color:#0284c7; background:#f8fafc;'>" + sec.nombre + "</td></tr>";
      }
      (sec.items || []).forEach(function(item) {
        filasHtml += "<tr>" +
          "<td style='padding:8px; border-bottom:1px solid #eee; font-size:13px;'>" + (item.desc || item.descripcion || "") + "</td>" +
          "<td style='padding:8px; border-bottom:1px solid #eee; text-align:center; font-size:13px;'>" + (item.unid || item.unidad || "") + "</td>" +
          "<td style='padding:8px; border-bottom:1px solid #eee; text-align:right; font-size:13px;'>S/ " + parseFloat(item.precio || 0).toFixed(2) + "</td>" +
          "<td style='padding:8px; border-bottom:1px solid #eee; text-align:right; font-size:13px;'>S/ " + parseFloat(item.subtotal || 0).toFixed(2) + "</td>" +
          "</tr>";
      });
    });
  }
  return "<div style='font-family:Arial,sans-serif; padding:20px; color:#333; max-width:800px; margin:0 auto;'>" +
    "<table style='width:100%; margin-bottom:20px;'><tr>" +
      "<td><h2 style='margin:0; color:#0f172a;'>Alexander D&amp;C</h2>" +
          "<span style='color:#0284c7; font-size:12px; font-weight:bold;'>COTIZACIÓN</span></td>" +
      "<td style='text-align:right;'>" +
          "<h3 style='margin:0; color:#0f172a;'>" + (payload.numero || "") + "</h3>" +
          "<span style='font-size:12px; color:#64748b;'>Fecha: " + (payload.fecha || "") + "</span></td>" +
    "</tr></table>" +
    "<hr style='border:0; border-top:1px solid #e2e8f0; margin-bottom:15px;'>" +
    "<table style='width:100%; margin-bottom:20px; background-color:#f8fafc; padding:15px; border-radius:6px;'><tr>" +
      "<td><strong>Cliente:</strong> " + (payload.cliente || "") +
          "<br><strong>Documento:</strong> " + (payload.ruc || "") + "</td>" +
    "</tr></table>" +
    "<table style='width:100%; border-collapse:collapse; margin-bottom:25px;'>" +
      "<thead><tr style='background-color:#f1f5f9;'>" +
        "<th style='padding:10px; text-align:left;   font-size:12px; color:#475569;'>Descripción</th>" +
        "<th style='padding:10px; text-align:center; font-size:12px; color:#475569;'>Unid</th>" +
        "<th style='padding:10px; text-align:right;  font-size:12px; color:#475569;'>P. Unit</th>" +
        "<th style='padding:10px; text-align:right;  font-size:12px; color:#475569;'>Subtotal</th>" +
      "</tr></thead>" +
      "<tbody>" + filasHtml + "</tbody>" +
    "</table>" +
    "<table style='width:100%; margin-top:20px;'><tr>" +
      "<td style='width:60%;'></td>" +
      "<td style='width:40%;'>" +
        "<table style='width:100%; text-align:right; font-size:13px;'>" +
          "<tr><td style='padding:4px;'>Subtotal:</td><td style='padding:4px;'>S/ " + (payload.subtotal || 0) + "</td></tr>" +
          "<tr><td style='padding:4px;'>IGV (18%):</td><td style='padding:4px;'>S/ " + (payload.igv || 0) + "</td></tr>" +
          "<tr style='font-size:16px; color:#0f172a;'>" +
            "<td style='padding-top:10px; border-top:1px solid #cbd5e1;'><strong>Total:</strong></td>" +
            "<td style='padding-top:10px; border-top:1px solid #cbd5e1;'><strong>S/ " + (payload.total || 0) + "</strong></td>" +
          "</tr>" +
        "</table>" +
      "</td>" +
    "</tr></table>" +
    (payload.observaciones
      ? "<div style='font-size:11px; color:#64748b; margin-top:30px; padding:10px; border-left:3px solid #0284c7; background-color:#f8fafc;'>" +
          "<strong>Observaciones:</strong><br>" + String(payload.observaciones).replace(/\n/g, "<br>") + "</div>"
      : "") +
  "</div>";
}

// ── Reportes de Obra: guardar el PDF EXACTO que ve el usuario en pantalla ──
// Igual que con las cotizaciones, el frontend manda el HTML real del documento
// (payload.htmlFull) y aquí solo lo convertimos a PDF y lo guardamos en su propia
// subcarpeta de Drive ("Reportes de Obra"), para no mezclarlo con las fotos ni
// con los PDF de cotizaciones.
function manejarGuardarPdfReporte(payload) {
  const carpetaRaiz     = DriveApp.getFolderById(FOLDER_ID);
  const carpetaReportes = obtenerOCrearSubcarpeta(carpetaRaiz, "Reportes de Obra");

  const htmlCuerpo = payload.htmlFull || "<p>Reporte sin contenido</p>";
  const archivoHtml = Utilities.newBlob(htmlCuerpo, "text/html", "temp.html");
  const pdfBlob      = archivoHtml.getAs("application/pdf");
  const nombreLimpio = limpiarNombreArchivo(payload.titulo || payload.nombre || "Reporte_de_Obra");
  pdfBlob.setName(nombreLimpio + ".pdf");

  if (payload.pdfFileIdAnterior) {
    try { DriveApp.getFileById(payload.pdfFileIdAnterior).setTrashed(true); }
    catch (errBorrar) { /* ya no existe: seguimos igual */ }
  }

  const nuevoPdf = carpetaReportes.createFile(pdfBlob);
  try {
    nuevoPdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (errShare) { /* si falla el permiso, igual devolvemos el enlace */ }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, url: nuevoPdf.getUrl(), fileId: nuevoPdf.getId() }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Galería: subir/reemplazar una foto en Drive ────────────────────────────
// CAMBIO v3.1: ahora la foto se comparte como "Cualquiera con el enlace puede
// ver" y se devuelve, además del enlace de Drive (para abrir/descargar), una
// URL DIRECTA que sí funciona como <img src> desde cualquier dispositivo.
function manejarGuardarFotoGaleria(payload) {
  const lock = LockService.getScriptLock();
  let carpetaObra;
  try {
    lock.waitLock(20000);
    const carpetaRaiz    = DriveApp.getFolderById(FOLDER_ID);
    const carpetaGaleria = obtenerOCrearSubcarpeta(carpetaRaiz, "Galeria Fotografica");
    const nombreObraLimpio = limpiarNombreArchivo(payload.obraNombre || "Obra sin nombre");
    carpetaObra = obtenerOCrearSubcarpeta(carpetaGaleria, nombreObraLimpio);
  } finally {
    lock.releaseLock();
  }

  const dataUrl    = payload.imagenBase64 || "";
  const mimeMatch  = /^data:(image\/[a-zA-Z0-9.+-]+);base64,/.exec(dataUrl);
  const mimeType   = mimeMatch ? mimeMatch[1] : "image/jpeg";
  const base64Data = dataUrl.indexOf(",") >= 0 ? dataUrl.split(",")[1] : dataUrl;
  const bytes      = Utilities.base64Decode(base64Data);
  const blob       = Utilities.newBlob(bytes, mimeType, "foto.jpg");

  if (payload.fileIdAnterior) {
    try { DriveApp.getFileById(payload.fileIdAnterior).setTrashed(true); }
    catch (errBorrar) { /* ya no existe: seguimos igual */ }
  }

  const nombreObraLimpio = limpiarNombreArchivo(payload.obraNombre || "Obra sin nombre");
  const etiquetaTag = payload.etiqueta ? "_" + payload.etiqueta : "";
  const marcaTiempo = new Date().getTime();
  const nombreEtapa = limpiarNombreArchivo(payload.etapa || "foto");
  blob.setName(nombreObraLimpio + "_" + nombreEtapa + etiquetaTag + "_" + marcaTiempo + ".jpg");

  const archivo = carpetaObra.createFile(blob);
  archivo.setDescription(
    "Obra: "        + (payload.obraNombre || "")  + "\n" +
    "Cliente: "     + (payload.cliente    || "")  + "\n" +
    "Etapa: "       + (payload.etapa      || "")  + "\n" +
    "Etiqueta: "    + (payload.etiqueta   || "")  + "\n" +
    "Ubicación: "   + (payload.ubicacion  || "")  + "\n" +
    "Descripción: " + (payload.caption    || "")  + "\n" +
    "Fecha: "       + (payload.fecha      || "")
  );

  // NUEVO: compartir públicamente por enlace, para que la foto se pueda ver
  // como <img> desde cualquier dispositivo (no solo desde la cuenta dueña del Drive).
  try {
    archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (errShare) { /* si falla el permiso, igual devolvemos los enlaces */ }

  const urlDirecta = "https://drive.google.com/uc?export=view&id=" + archivo.getId();

  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      url: archivo.getUrl(),      // enlace para abrir/descargar en Drive
      urlDirecta: urlDirecta,     // NUEVO: enlace directo para usar como <img src>
      fileId: archivo.getId()
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function manejarEliminarArchivoDrive(payload) {
  const fileId = payload.fileId;
  if (!fileId) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: "Falta 'fileId' para eliminar." }))
      .setMimeType(ContentService.MimeType.JSON);
  }
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (err) {
    // Ya no existe: no pasa nada.
  }
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Utilidades Drive ───────────────────────────────────────────
function test()              { DriveApp.getRootFolder(); }
function autorizar()         { DriveApp.getRootFolder(); }
function verificarPermisos() {
  var carpeta = DriveApp.getRootFolder();
  Logger.log("Drive OK: " + carpeta.getName());
}
function testDrive() {
  var carpeta = DriveApp.getFolderById(FOLDER_ID);
  Logger.log(carpeta.getName());
}

function obtenerOCrearSubcarpeta(padre, nombre) {
  const iter = padre.getFoldersByName(nombre);
  return iter.hasNext() ? iter.next() : padre.createFolder(nombre);
}

function limpiarNombreArchivo(texto) {
  const limpio = String(texto).replace(/[^a-zA-Z0-9 _-]/g, "").trim().replace(/\s+/g, "_");
  return limpio || "Sin_Nombre";
}

// ── LEER hoja como array de objetos ───────────────────────────
function leerHoja(ss, nombre) {
  var sheet = ss.getSheetByName(nombre);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var rows    = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  return rows.map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) {
      if (h) obj[h] = row[i] === "" ? null : row[i];
    });
    return obj;
  });
}

// ── GUARDAR hoja desde array de objetos ───────────────────────
function guardarHoja(ss, nombre, datos) {
  var sheet = ss.getSheetByName(nombre);
  if (!sheet) sheet = ss.insertSheet(nombre);
  sheet.clearContents();
  if (!datos || datos.length === 0) return;
  var headers = Object.keys(datos[0]);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  var rows = datos.map(function(obj) {
    return headers.map(function(h) {
      var val = obj[h];
      if (val === null || val === undefined) return "";
      if (typeof val === "object") return JSON.stringify(val);
      return val;
    });
  });
  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
}

// ── LEER Config ───────────────────────────────────────────────
function leerConfig(ss, key, valorDefault) {
  var sheet = ss.getSheetByName(HOJAS.config);
  if (!sheet) return valorDefault;
  var datos = sheet.getDataRange().getValues();
  for (var i = 0; i < datos.length; i++) {
    if (String(datos[i][0]) === String(key)) return datos[i][1];
  }
  return valorDefault;
}

// ── GUARDAR Config ────────────────────────────────────────────
function guardarConfig(ss, key, value) {
  var sheet = ss.getSheetByName(HOJAS.config);
  if (!sheet) {
    sheet = ss.insertSheet(HOJAS.config);
    sheet.getRange(1, 1, 1, 2).setValues([["Clave", "Valor"]]);
  }
  var datos = sheet.getDataRange().getValues();
  for (var i = 0; i < datos.length; i++) {
    if (String(datos[i][0]) === String(key)) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}
