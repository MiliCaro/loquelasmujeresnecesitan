/**
 * Backend del formulario de organizaciones aliadas.
 * Recibe cada registro, guarda el logo en una carpeta de Drive
 * y agrega una fila a esta misma planilla.
 *
 * Pegar este código en: la planilla → Extensiones → Apps Script.
 * Luego: Implementar → Nueva implementación → Aplicación web
 *   - Ejecutar como: Yo (tu cuenta)
 *   - Quién tiene acceso: Cualquier usuario
 * Copiar la URL que termina en /exec y pasársela a Claude.
 */

var FOLDER_NAME = 'Logos aliadas - Lo que las Mujeres Necesitan';
var HEADERS = ['Fecha', 'Organización', 'Tipo', 'Provincia', 'Sitio web',
  'Contacto', 'Cargo', 'Email', 'WhatsApp', 'Difundir consulta',
  'Autoriza logo', 'Punto presencial', 'Mensaje', 'Logo (link)'];

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheets()[0];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    var logoUrl = '';
    if (data.logo_b64) {
      var folder = getFolder_();
      var bytes = Utilities.base64Decode(data.logo_b64);
      var safeOrg = String(data.organizacion || 'logo').replace(/[^\w\- ]+/g, '').substring(0, 60);
      var blob = Utilities.newBlob(bytes, data.logo_mime || 'application/octet-stream',
        safeOrg + ' - ' + (data.logo_name || 'logo'));
      var file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      logoUrl = file.getUrl();
    }

    sheet.appendRow([
      new Date(),
      safe_(data.organizacion),
      safe_(data.tipo),
      safe_(data.provincia),
      safe_(data.web),
      safe_(data.nombre),
      safe_(data.cargo),
      safe_(data.email),
      safe_(data.whatsapp),
      safe_(data.comp_difundir),
      safe_(data.comp_logo),
      safe_(data.comp_presencial),
      safe_(data.mensaje),
      logoUrl
    ]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};

  // Contador de respuestas de SurveyMonkey (para el sitio)
  if (params.action === 'count') {
    var count = getStoredCount_();
    var out = JSON.stringify({ ok: true, count: count });
    if (params.callback) {
      // JSONP: permite que el navegador lea la respuesta desde otro dominio
      return ContentService
        .createTextOutput(params.callback + '(' + out + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
  }

  return json_({ ok: true, msg: 'Endpoint activo' });
}

// ===== Contador SurveyMonkey =====
// Configurar en Configuración del proyecto → Propiedades del script:
//   SM_TOKEN      = token de acceso de SurveyMonkey
//   SM_SURVEY_ID  = id numérico de la encuesta (usar listSurveys() para verlo)
// Y crear un activador (trigger) horario que ejecute refreshResponseCount.

function refreshResponseCount() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('SM_TOKEN');
  var surveyId = props.getProperty('SM_SURVEY_ID');
  if (!token || !surveyId) return;

  var res = UrlFetchApp.fetch('https://api.surveymonkey.com/v3/surveys/' + surveyId, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) return;

  var data = JSON.parse(res.getContentText());
  if (data && typeof data.response_count === 'number') {
    props.setProperty('SM_COUNT', String(data.response_count));
  }
}

function getStoredCount_() {
  var v = PropertiesService.getScriptProperties().getProperty('SM_COUNT');
  return v ? parseInt(v, 10) : null;
}

// Ejecutar una vez desde el editor para ver el id de cada encuesta en los registros.
function listSurveys() {
  var token = PropertiesService.getScriptProperties().getProperty('SM_TOKEN');
  var res = UrlFetchApp.fetch('https://api.surveymonkey.com/v3/surveys?per_page=50', {
    method: 'get',
    headers: { Authorization: 'Bearer ' + token, Accept: 'application/json' },
    muteHttpExceptions: true
  });
  var data = JSON.parse(res.getContentText());
  (data.data || []).forEach(function (s) {
    Logger.log(s.id + '  →  ' + s.title);
  });
}

// Evita que Sheets interprete valores como fórmulas (ej: WhatsApp que empieza con "+")
function safe_(v) {
  var s = (v == null) ? '' : String(v);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function getFolder_() {
  var it = DriveApp.getFoldersByName(FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
