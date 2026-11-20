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

var ID_CERTS = '1LloFdJY--VTTjthoa-qIa9rsCifB_mSNqehhI40NGrw'; // Unicert Global - Base de datos de certificados
var HOJAS_CERT = {iso: 'Certificados ISO', curso: 'Cursos'};
var CAMPOS_ISO = {num:0, empresa:1, pais:2, norma:3, alcance:4, iaf:5, emision:6, vence:7, seguimiento:8, estado:9, emisor:10};
var CAMPOS_CURSO = {num:0, participante:1, curso:2, norma:3, horas:4, modalidad:5, inicio:6, fin:7, emision:8, resultado:9, estado:10, instructor:11, organizacion:12, pais:13};

function val_(v) { return v instanceof Date ? Utilities.formatDate(v, 'America/Asuncion', 'dd/MM/yyyy') : String(v == null ? '' : v).trim(); }
function norm_(s) { return String(s || '').replace(/\s+/g, '').toUpperCase(); }

/** Lee la hoja de certificados (cache 60 s). Solo devuelve campos publicos. */
function cert_(tipo) {
  var cache = CacheService.getScriptCache(), key = 'cert_' + tipo, hit = cache.get(key);
  if (hit) return JSON.parse(hit);
  var campos = tipo === 'iso' ? CAMPOS_ISO : CAMPOS_CURSO;
  var v = SpreadsheetApp.openById(ID_CERTS).getSheetByName(HOJAS_CERT[tipo]).getDataRange().getValues(), out = [];
  for (var i = 1; i < v.length; i++) {
    if (!val_(v[i][0])) continue;
    var o = {};
    Object.keys(campos).forEach(function (k) { o[k] = val_(v[i][campos[k]]); });
    out.push(o);
  }
  try { cache.put(key, JSON.stringify(out), 60); } catch (x) {}
  return out;
}

/** Verificador publico: ?action=verify&type=iso|curso&q=NUMERO  |  ?action=list&type=iso */
function doGet(e) {
  var p = (e && e.parameter) || {}, res;
  try {
    var tipo = p.type === 'curso' ? 'curso' : 'iso';
    if (p.action === 'list') {
      res = {ok: true, type: tipo, rows: tipo === 'iso' ? cert_('iso') : []};
    } else if (p.action === 'verify') {
      var q = norm_(p.q), rows = [];
      if (q.length >= (tipo === 'curso' ? 6 : 3)) {
        var all = cert_(tipo);
        rows = all.filter(function (r) { return norm_(r.num) === q; });
        if (!rows.length && tipo === 'iso') {
          var t = String(p.q || '').trim().toUpperCase();
          rows = all.filter(function (r) { return norm_(r.num).indexOf(q) > -1 || r.empresa.toUpperCase().indexOf(t) > -1; });
        }
      }
      res = {ok: true, type: tipo, rows: rows.slice(0, 20)};
    } else {
      return ContentService.createTextOutput('Unicert formularios OK');
    }
  } catch (err) { res = {ok: false, error: 'error'}; }
  return ContentService.createTextOutput(JSON.stringify(res)).setMimeType(ContentService.MimeType.JSON);
}
/** Ejecutar una vez a mano para autorizar permisos y crear la hoja. */
function autorizar() { tab_('Propuestas', COLS_PROP); tab_('Landing', COLS_LANDING); carpeta_(); cert_('iso'); cert_('curso'); Logger.log('OK'); }
