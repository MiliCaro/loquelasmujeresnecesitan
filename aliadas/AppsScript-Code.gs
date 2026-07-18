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

function doGet() {
  return json_({ ok: true, msg: 'Endpoint activo' });
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
