/**
 * Unicert Global - Receptor de formularios (Google Apps Script)
 * Sirve a la LANDING (certification.unicertglobal.org) y al SITIO (www.unicertglobal.org).
 *
 * Qué hace:
 *  - Guarda cada solicitud en dos planillas de Drive (carpeta "dir-cert"): una para la landing y otra para el sitio.
 *  - Envía un aviso a info@unicertglobal.org (con enlaces a los adjuntos).
 *  - Envía confirmación al cliente con su número de referencia (formulario del sitio).
 *  - Guarda los adjuntos en la carpeta de Drive "Unicert - Adjuntos".
 *
 * Pasos: script.google.com > Nuevo proyecto > pegar este código > Guardar
 *  > Ejecutar una vez "autorizar" (acepta permisos de Hojas, Drive y Gmail)
 *  > Implementar > Nueva implementación > Aplicación web
 *    Ejecutar como: Yo | Quién tiene acceso: Cualquier persona
 *  > Copiar la URL (termina en /exec) en:
 *      - landing: GOOGLE_SHEET_URL en index.html
 *      - sitio:   UG_CFG.GOOGLE_SCRIPT_URL en index.html
 *  Si cambia el código, hay que crear una "Nueva versión" de la implementación.
 */
var AVISO_A = 'info@unicertglobal.org';
var NOMBRE_HOJA = 'Unicert - Leads';
var CARPETA = 'Unicert - Adjuntos';
var COLS_LANDING = ['timestamp','nombreCompleto','empresa','pais','whatsapp','email','normaISO','numEmpleados','motivo','urgencia','comentario'];
var COLS_PROP = ['referencia','fecha','nombre','empresa','pais','ciudad','email','whatsapp','normas','actividad','trabajadores','mensaje','adjuntos','origen'];

// Planillas ya creadas en Drive (carpeta "dir-cert")
var ID_LANDING = '1dfhnZenKAES1YKR3ZnMQ_OZkttK-S4g1MuRrcAjoncE'; // Unicert - Leads Landing
var ID_SITIO   = '1tz78wvL3dyt9JL6EiUyYGWe9CYgXg7Mki8aBb8ix9Oc'; // Unicert - Propuestas Website

function tab_(name, cols) {
  var ss = SpreadsheetApp.openById(name === 'Landing' ? ID_LANDING : ID_SITIO);
  var sh = ss.getSheets()[0];
  if (sh.getLastRow() === 0) { sh.appendRow(cols); sh.setFrozenRows(1); }
  var w = cols.indexOf('whatsapp') + 1;
  if (w > 0) sh.getRange(1, w, sh.getMaxRows(), 1).setNumberFormat('@');
  return sh;
}
function fila_(sh, vals) {
  var r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, vals.length).setNumberFormat('@').setValues([vals.map(String)]);
}
function carpeta_() {
  var it = DriveApp.getFoldersByName(CARPETA);
  return it.hasNext() ? it.next() : DriveApp.createFolder(CARPETA);
}
function esc_(s) { return String(s || '').replace(/[&<>]/g, function (c) { return {'&':'&amp;','<':'&lt;','>':'&gt;'}[c]; }); }

function doPost(e) {
  try {
    var d = JSON.parse(e.postData.contents);
    if (d.referencia) return propuesta_(d);
    var sh = tab_('Landing', COLS_LANDING);
    fila_(sh, COLS_LANDING.map(function (c) { return d[c] || ''; }));
    MailApp.sendEmail({ to: AVISO_A, subject: 'Nuevo lead (landing) – ' + (d.empresa || ''),
      htmlBody: '<p><b>' + esc_(d.nombreCompleto) + '</b> – ' + esc_(d.empresa) + ' (' + esc_(d.pais) + ')</p><p>WhatsApp: ' + esc_(d.whatsapp) + '<br>Correo: ' + esc_(d.email) + '<br>Norma: ' + esc_(d.normaISO) + '<br>Empleados: ' + esc_(d.numEmpleados) + '<br>Motivo: ' + esc_(d.motivo) + '<br>Comentario: ' + esc_(d.comentario) + '</p>' });
    return ContentService.createTextOutput('ok');
  } catch (err) { return ContentService.createTextOutput('error: ' + err); }
}

function propuesta_(d) {
  var links = [];
  (d.adjuntos || []).forEach(function (a) {
    var blob = Utilities.newBlob(Utilities.base64Decode(a.datos), a.tipo, d.referencia + ' - ' + a.nombre);
    var f = carpeta_().createFile(blob);
    links.push(f.getUrl());
  });
  d.adjuntos = links.join('\n');
  fila_(tab_('Propuestas', COLS_PROP), COLS_PROP.map(function (c) { return d[c] || ''; }));
  var cuerpo = '<p>Referencia: <b>' + esc_(d.referencia) + '</b></p><p><b>' + esc_(d.nombre) + '</b> – ' + esc_(d.empresa) + '<br>' + esc_(d.pais) + ' ' + esc_(d.ciudad) + '<br>Correo: ' + esc_(d.email) + '<br>WhatsApp: ' + esc_(d.whatsapp) + '<br>Normas: ' + esc_(d.normas) + '<br>Trabajadores: ' + esc_(d.trabajadores) + '<br>Actividad: ' + esc_(d.actividad) + '<br>Mensaje: ' + esc_(d.mensaje) + '</p>' + (links.length ? '<p>Adjuntos:<br>' + links.join('<br>') + '</p>' : '');
  MailApp.sendEmail({ to: AVISO_A, replyTo: d.email, subject: 'Solicitud de propuesta ' + d.referencia + ' – ' + d.empresa, htmlBody: cuerpo });
  if (d.email) {
    MailApp.sendEmail({ to: d.email, replyTo: AVISO_A, name: 'Unicert Global',
      subject: 'Recibimos su solicitud ' + d.referencia,
      htmlBody: '<p>Hola ' + esc_(d.nombre) + ',</p><p>Recibimos su solicitud de propuesta para <b>' + esc_(d.empresa) + '</b>. Su número de referencia es <b>' + esc_(d.referencia) + '</b>.</p><p>Le responderemos con una propuesta a la brevedad. Si necesita algo más, escríbanos a ' + AVISO_A + ' o por WhatsApp al +595 987 560731 indicando la referencia.</p><p>Unicert Global</p>' });
  }
  return ContentService.createTextOutput('ok');
}

function doGet() { return ContentService.createTextOutput('Unicert formularios OK'); }
/** Ejecutar una vez a mano para autorizar permisos y crear la hoja. */
function autorizar() { tab_('Propuestas', COLS_PROP); tab_('Landing', COLS_LANDING); carpeta_(); Logger.log('OK'); }
