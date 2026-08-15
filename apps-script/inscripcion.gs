// Notificación interna al staff por cada preinscripción/inscripción: desactivada por ahora
// (lista vacía) — sendNotificationEmail_ se salta el envío cuando no hay destinatarios.
var NOTIFY_EMAILS_ = [];
var SITE_URL_ = 'https://acabrera93.github.io/Euromodelo/';
// Acceso simple al endpoint de simulación (solo lectura, no escribe en la Sheet): clave
// compartida + lista de correos autorizados. No es autenticación real, es un filtro básico.
var STAFF_KEY_ = 'euromodelo2026';
var STAFF_AUTHORIZED_EMAILS_ = [
  'juan.ladino@fundacionrevel.net',
  'juan.ovalle@fundacionrevel.net',
  'alejandro.cabrera@fundacionrevel.net',
  'andres.dewasseige@fundacionrevel.net'
];

var FIELD_LABELS_ = {
  enviado: 'Fecha de envío',
  ref_code: 'Código de referencia',
  email: 'Correo electrónico',
  password: 'Contraseña',
  nombre_completo: 'Nombre completo',
  tipo_documento: 'Tipo de documento',
  numero_documento: 'Número de documento',
  telefono: 'Teléfono',
  ciudad: 'Ciudad',
  institucion_educativa: 'Institución educativa',
  rol_opcion1: 'Rol — Opción 1',
  rol_opcion2: 'Rol — Opción 2',
  rol_opcion3: 'Rol — Opción 3',
  comision_opcion1: 'Comisión — Opción 1',
  comision_opcion2: 'Comisión — Opción 2',
  comision_opcion3: 'Comisión — Opción 3',
  partido: 'Partido político',
  resultado_brujula: 'Resultado Brújula Legislativa',
  resultado_brujula_comision: 'Resultado Brújula de Comisión',
  resultado_brujula_partido: 'Resultado Match Europeo',
  autoriza_datos: 'Autoriza tratamiento de datos',
  es_menor: 'Es menor de edad',
  autoriza_imagen: 'Autoriza derechos de imagen',
  rol_asignado: 'Rol asignado',
  comision_asignada: 'Comisión asignada',
  partido_asignado: 'Partido asignado',
  tipo: 'Tipo de cuenta',
  propuesta_url: 'Propuesta legislativa (PDF)',
  propuesta_estado: 'Estado de la propuesta',
  propuesta_comentario: 'Comentario del staff sobre la propuesta',
  asignacion_origen: 'Origen de la asignación',
  tipo_euromodelo: 'Regional o Nacional'
};

function normalizeName_(s) {
  return s.toString().toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
}

function findOrCreateSheet_(ss, candidateNames, fallbackName) {
  var normalizedCandidates = candidateNames.map(normalizeName_);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (normalizedCandidates.indexOf(normalizeName_(sheets[i].getName())) !== -1) {
      return sheets[i];
    }
  }
  return ss.insertSheet(fallbackName);
}

// Lee la fila 1 como cabeceras reales; si la hoja está vacía, la inicializa con canonicalHeaders.
// Así, si alguien reordena las columnas manualmente en la Sheet, los datos se siguen ubicando
// según el nombre de cada columna y no según una posición fija.
// Si la hoja ya existe pero le faltan columnas de canonicalHeaders (p.ej. porque venía de una
// versión anterior del script), las agrega al final en vez de ignorarlas.
function getHeaders_(sheet, canonicalHeaders) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(canonicalHeaders);
    return canonicalHeaders;
  }
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var missing = canonicalHeaders.filter(function(h) { return headers.indexOf(h) === -1; });
  if (missing.length > 0) {
    sheet.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    headers = headers.concat(missing);
  }
  return headers;
}

function escapeHtml_(s) {
  return s.toString()
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function randomPassword_() {
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var out = '';
  for (var i = 0; i < 6; i++) out += chars.charAt(Math.floor(Math.random() * chars.length));
  return out;
}

// ---------- Plantilla base de correo (cabecera con logo centrado) ----------
function emailShell_(titleHtml, bodyHtml) {
  var logoUrl = SITE_URL_ + 'assets/logo.png';
  return (
    '<div style="font-family:Arial,Helvetica,sans-serif; max-width:600px; margin:0 auto; background:#EDF1F6; padding:24px 0;">' +
      '<div style="max-width:600px; margin:0 auto; background:#ffffff; border-radius:14px; overflow:hidden; border:1px solid #D7DEE8;">' +
        '<div style="background:#0B2545; padding:28px 28px 22px; text-align:center;">' +
          '<img src="' + logoUrl + '" alt="Euromodelo Joven - Fundación Revel" style="height:44px; display:block; margin:0 auto 14px;">' +
          '<div style="color:#ffffff; font-size:20px; font-weight:700;">' + titleHtml + '</div>' +
        '</div>' +
        '<div style="padding:26px 28px;">' + bodyHtml + '</div>' +
        '<div style="background:#EDF1F6; padding:16px 28px; font-size:11.5px; color:#8695AC; text-align:center;">Fundación Revel — XX Euromodelo Joven 2026</div>' +
      '</div>' +
    '</div>'
  );
}

function credentialsCardHtml_(username, password) {
  return (
    '<div style="background:#0B2545; border-radius:12px; padding:22px; text-align:center; margin-bottom:22px;">' +
      '<p style="margin:0 0 14px; font-size:12px; color:rgba(255,255,255,.7); text-transform:uppercase; letter-spacing:.08em;">Tus credenciales de acceso</p>' +
      '<table style="width:100%; border-collapse:collapse;"><tr>' +
        '<td style="padding:8px; text-align:center;">' +
          '<div style="font-size:11px; color:#C9A227; text-transform:uppercase; letter-spacing:.08em; margin-bottom:4px;">Usuario</div>' +
          '<div style="font-size:15px; color:#ffffff; font-weight:700; word-break:break-all;">' + escapeHtml_(username) + '</div>' +
        '</td>' +
        '<td style="padding:8px; text-align:center;">' +
          '<div style="font-size:11px; color:#C9A227; text-transform:uppercase; letter-spacing:.08em; margin-bottom:4px;">Contraseña</div>' +
          '<div style="font-size:15px; color:#ffffff; font-weight:700;">' + escapeHtml_(password) + '</div>' +
        '</td>' +
      '</tr></table>' +
    '</div>'
  );
}

function loginButtonHtml_() {
  var loginUrl = SITE_URL_ + 'perfil.html';
  return (
    '<div style="text-align:center;">' +
      '<a href="' + loginUrl + '" style="display:inline-block; background:#C9A227; color:#0B2545; text-decoration:none; font-weight:700; font-size:14px; padding:13px 30px; border-radius:8px;">Iniciar sesión →</a>' +
    '</div>'
  );
}

// ---------- Correo al participante: confirmación de preinscripción ----------
function sendParticipantWelcomeEmail_(data) {
  var isProfesor = (data.tipo || '').toString() === 'Profesor';
  var introText = isProfesor
    ? 'Hola <b>' + escapeHtml_(data.nombre_completo || '') + '</b>, gracias por registrarte como profesor acompañante en el <b>XX Euromodelo Joven 2026</b>. Guarda tus credenciales: con ellas podrás iniciar sesión y ver el listado de estudiantes inscritos de tu colegio.'
    : 'Hola <b>' + escapeHtml_(data.nombre_completo || '') + '</b>, gracias por preinscribirte al <b>XX Euromodelo Joven 2026</b>. Guarda tus credenciales: las necesitarás para iniciar sesión y completar tu inscripción.';
  var body =
    '<p style="color:#3A4A63; font-size:14.5px; line-height:1.6; margin:0 0 18px;">' + introText + '</p>' +
    credentialsCardHtml_(data.email || '', data.password || '') +
    loginButtonHtml_() +
    '<p style="color:#8695AC; font-size:12px; line-height:1.6; margin:22px 0 0; text-align:center;">Si no reconoces esta preinscripción, puedes ignorar este correo.</p>';
  MailApp.sendEmail({
    to: data.email,
    subject: 'Tu preinscripción al XX Euromodelo Joven 2026',
    htmlBody: emailShell_('¡Preinscripción registrada!', body),
  });
}

// ---------- Correo al participante: nueva contraseña ----------
function sendPasswordResetEmail_(email, nombre, newPassword) {
  var body =
    '<p style="color:#3A4A63; font-size:14.5px; line-height:1.6; margin:0 0 18px;">Hola' + (nombre ? ' <b>' + escapeHtml_(nombre) + '</b>' : '') + ', generamos una nueva contraseña para tu cuenta del <b>XX Euromodelo Joven 2026</b> a pedido tuyo.</p>' +
    credentialsCardHtml_(email || '', newPassword || '') +
    loginButtonHtml_() +
    '<p style="color:#8695AC; font-size:12px; line-height:1.6; margin:22px 0 0; text-align:center;">Si no solicitaste este cambio, escríbenos de inmediato.</p>';
  MailApp.sendEmail({
    to: email,
    subject: 'Nueva contraseña — XX Euromodelo Joven 2026',
    htmlBody: emailShell_('Nueva contraseña generada', body),
  });
}

// ---------- Correo interno al staff (resumen de cada envío) ----------
function sendNotificationEmail_(isPre, data, canonicalHeaders, sheetUrl) {
  if (!NOTIFY_EMAILS_.length) return; // notificación desactivada: sin destinatarios, no se envía
  var tipoLabel = isPre ? 'Preinscripción' : 'Inscripción';
  var subject = 'Nueva ' + tipoLabel.toLowerCase() + ' — ' + (data.nombre_completo || 'Euromodelo Joven 2026');

  var rowsHtml = canonicalHeaders.map(function(key) {
    var value = data[key];
    if (!value) return '';
    var label = FIELD_LABELS_[key] || key;
    return '<tr>' +
      '<td style="padding:9px 12px; border-bottom:1px solid #EDF1F6; color:#0B2545; font-weight:600; font-size:13px; width:42%; vertical-align:top;">' + escapeHtml_(label) + '</td>' +
      '<td style="padding:9px 12px; border-bottom:1px solid #EDF1F6; color:#3A4A63; font-size:13px; vertical-align:top;">' + escapeHtml_(value) + '</td>' +
      '</tr>';
  }).join('');

  var body =
    '<p style="color:#4A5A73; font-size:14px; line-height:1.6; margin:0 0 18px;">Se registró una nueva <b>' + tipoLabel.toLowerCase() + '</b> en el sitio web del Euromodelo Joven 2026. Estos son los datos enviados:</p>' +
    '<table style="width:100%; border-collapse:collapse;">' + rowsHtml + '</table>' +
    '<div style="text-align:center; margin-top:26px;">' +
      '<a href="' + sheetUrl + '" style="display:inline-block; background:#1E3D73; color:#ffffff; text-decoration:none; font-weight:700; font-size:13.5px; padding:13px 28px; border-radius:8px;">Ver en la Sheet →</a>' +
    '</div>';

  MailApp.sendEmail({
    to: NOTIFY_EMAILS_.join(','),
    subject: subject,
    htmlBody: emailShell_('Nueva ' + tipoLabel.toLowerCase(), body),
  });
}

// Busca, en una columna dada, la fila que coincide con el valor. Devuelve el número
// de fila en la Sheet (1-indexado) o -1 si no la encuentra.
function findRowByColumn_(sheet, headers, columnName, value, caseInsensitive) {
  if (!value) return -1;
  var colIdx = headers.indexOf(columnName);
  if (colIdx === -1) return -1;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var values = sheet.getRange(2, colIdx + 1, lastRow - 1, 1).getValues();
  var target = caseInsensitive ? value.toString().toLowerCase() : value;
  for (var i = 0; i < values.length; i++) {
    var cellVal = caseInsensitive ? values[i][0].toString().toLowerCase() : values[i][0];
    if (cellVal === target) return i + 2;
  }
  return -1;
}

function mapRecordToUser_(record) {
  var user = {
    email: record.email || '',
    nombre: record.nombre_completo || '',
    tipoDocumento: record.tipo_documento || '',
    numDocumento: record.numero_documento || '',
    telefono: (record.telefono || '').toString().replace(/^\+57\s*/, ''),
    ciudad: record.ciudad || '',
    institucion: record.institucion_educativa || '',
    autorizaDatos: record.autoriza_datos || '',
    esMenor: record.es_menor || '',
    autorizaImagen: record.autoriza_imagen || '',
    refCode: record.ref_code || '',
    inscripcion: null,
    // Rol, comisión y partido finales: los completa el staff manualmente en la Sheet
    // una vez define la asignación. Si están presentes, reemplazan las preferencias
    // en el área personal del participante.
    rolAsignado: record.rol_asignado || '',
    comisionAsignada: record.comision_asignada || '',
    partidoAsignado: record.partido_asignado || '',
    tipo: record.tipo || 'Estudiante',
    propuestaUrl: record.propuesta_url || '',
    propuestaEstado: record.propuesta_estado || '',
    propuestaComentario: record.propuesta_comentario || '',
    tipoEuromodelo: record.tipo_euromodelo || 'Nacional',
  };
  if (record.rol_opcion1) {
    user.inscripcion = {
      rol1: record.rol_opcion1, rol2: record.rol_opcion2, rol3: record.rol_opcion3,
      comision1: record.comision_opcion1, comision2: record.comision_opcion2, comision3: record.comision_opcion3,
      partido: record.partido,
      resultadoBrujula: record.resultado_brujula || '',
      resultadoBrujulaComision: record.resultado_brujula_comision || '',
      resultadoBrujulaPartido: record.resultado_brujula_partido || '',
      enviado: record.enviado || '',
    };
  }
  return user;
}

var CANONICAL_HEADERS_ = [
  'enviado', 'ref_code', 'email', 'password', 'nombre_completo', 'tipo_documento', 'numero_documento',
  'telefono', 'ciudad', 'institucion_educativa', 'autoriza_datos', 'es_menor', 'autoriza_imagen',
  'rol_opcion1', 'rol_opcion2', 'rol_opcion3',
  'comision_opcion1', 'comision_opcion2', 'comision_opcion3', 'partido',
  // Resultados de las 3 brújulas, obligatorias al inscribirse; se pueden repetir desde el
  // área personal, y cada repetición reemplaza el valor anterior.
  'resultado_brujula', 'resultado_brujula_comision', 'resultado_brujula_partido',
  // Estas tres columnas quedan en blanco al enviar el formulario: el staff las completa
  // manualmente en la Sheet una vez define la asignación final de cada participante.
  'rol_asignado', 'comision_asignada', 'partido_asignado',
  // 'Estudiante' | 'Profesor'. En blanco en filas anteriores a este cambio: mapRecordToUser_
  // las trata como 'Estudiante' por defecto, sin necesidad de migrar la Sheet.
  'tipo',
  // Propuesta legislativa del comisario: URL del PDF en Drive, estado ('' | 'Pendiente' |
  // 'Aprobada' | 'Plenaria') y comentario del staff. Una nueva subida reemplaza url/estado/
  // comentario anteriores (mismo criterio que repetir una brújula).
  'propuesta_url', 'propuesta_estado', 'propuesta_comentario',
  // '' | 'Manual' | 'Automática'. Se recalcula cada vez que rol_asignado/comision_asignada/
  // partido_asignado cambian, desde handleUpdateAssignment_ (Manual) o handleApplyAssignment_
  // (Automática); se limpia junto con esos tres campos al borrar o deshacer.
  'asignacion_origen',
  // 'Regional' | 'Nacional'. En blanco se trata como 'Nacional' (mapRecordToUser_ y
  // filterByScope_). Si es 'Regional', la sede es la propia columna `ciudad` — no hace falta
  // una columna aparte para eso.
  'tipo_euromodelo'
];

var INSCRIPCION_FIELDS_ = [
  'rol_opcion1', 'rol_opcion2', 'rol_opcion3',
  'comision_opcion1', 'comision_opcion2', 'comision_opcion3',
  'partido', 'resultado_brujula', 'resultado_brujula_comision', 'resultado_brujula_partido'
];

// ---------- Autenticación: login / recuperar contraseña / cambiar contraseña ----------
function handleLogin_(sheet, headers, data) {
  var email = (data.email || '').toString().trim().toLowerCase();
  var password = (data.password || '').toString();
  var rowNum = findRowByColumn_(sheet, headers, 'email', email, true);
  if (rowNum === -1) return jsonOut_({ ok: false });

  var rowValues = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  var record = {};
  headers.forEach(function(h, i) { record[h] = rowValues[i]; });

  if ((record.password || '').toString() !== password) {
    return jsonOut_({ ok: false });
  }
  return jsonOut_({ ok: true, user: mapRecordToUser_(record) });
}

function handleForgotPassword_(sheet, headers, data) {
  var email = (data.email || '').toString().trim().toLowerCase();
  var rowNum = findRowByColumn_(sheet, headers, 'email', email, true);
  if (rowNum !== -1) {
    var newPassword = randomPassword_();
    var colIdx = headers.indexOf('password');
    if (colIdx !== -1) {
      var cell = sheet.getRange(rowNum, colIdx + 1);
      cell.setNumberFormat('@');
      cell.setValue(newPassword);
    }
    var nombreIdx = headers.indexOf('nombre_completo');
    var nombre = nombreIdx !== -1 ? sheet.getRange(rowNum, nombreIdx + 1).getValue() : '';
    try {
      sendPasswordResetEmail_(email, nombre, newPassword);
    } catch (err) {
      console.error('No se pudo enviar el correo de recuperación: ' + err);
    }
  }
  // Respuesta genérica siempre, para no revelar si el correo está registrado.
  return jsonOut_({ ok: true });
}

function handleUpdatePassword_(sheet, headers, data) {
  var email = (data.email || '').toString().trim().toLowerCase();
  var newPassword = (data.newPassword || '').toString();
  var rowNum = findRowByColumn_(sheet, headers, 'email', email, true);
  if (rowNum !== -1 && newPassword) {
    var colIdx = headers.indexOf('password');
    if (colIdx !== -1) {
      var cell = sheet.getRange(rowNum, colIdx + 1);
      cell.setNumberFormat('@');
      cell.setValue(newPassword);
    }
  }
  return jsonOut_({ ok: true });
}

// Repetir la Brújula Legislativa desde el área personal reemplaza el resultado anterior.
// Repetir cualquiera de las 3 brújulas desde el área personal reemplaza el resultado anterior.
var BRUJULA_UPDATE_FIELDS_ = ['resultado_brujula', 'resultado_brujula_comision', 'resultado_brujula_partido'];
function handleUpdateBrujula_(sheet, headers, data) {
  var refCode = (data.ref_code || '').toString();
  var field = (data.field || 'resultado_brujula').toString();
  var resultado = (data.resultado || '').toString();
  if (BRUJULA_UPDATE_FIELDS_.indexOf(field) === -1) return jsonOut_({ ok: false });
  var rowNum = findRowByColumn_(sheet, headers, 'ref_code', refCode);
  if (rowNum !== -1 && resultado) {
    var colIdx = headers.indexOf(field);
    if (colIdx !== -1) {
      var cell = sheet.getRange(rowNum, colIdx + 1);
      cell.setNumberFormat('@');
      cell.setValue(resultado);
    }
  }
  return jsonOut_({ ok: true });
}

// ---------- Cuentas de profesor: listar estudiantes de su colegio ----------
// Los profesores viven en la misma pestaña "preinscripciones" que los estudiantes (columna
// `tipo`), así que reusan tal cual handleLogin_/handleForgotPassword_/handleUpdatePassword_.
// Esta es la única función propia que necesitan: sin sesión (como el resto del proyecto),
// reenvían su email/password en cada llamada y se validan aquí mismo.
function handleListStudents_(sheet, headers, data) {
  var email = (data.email || '').toString().trim().toLowerCase();
  var password = (data.password || '').toString();
  var rowNum = findRowByColumn_(sheet, headers, 'email', email, true);
  if (rowNum === -1) return jsonOut_({ ok: false });

  var rowValues = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  var record = {};
  headers.forEach(function(h, i) { record[h] = rowValues[i]; });

  if ((record.password || '').toString() !== password) return jsonOut_({ ok: false });
  if ((record.tipo || '').toString() !== 'Profesor') return jsonOut_({ ok: false });

  var institucion = (record.institucion_educativa || '').toString().trim().toLowerCase();
  var lastRow = sheet.getLastRow();
  var students = [];
  if (lastRow >= 2) {
    var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    values.forEach(function(row) {
      var r = {};
      headers.forEach(function(h, i) { r[h] = row[i]; });
      if ((r.tipo || '').toString() === 'Profesor') return;
      if ((r.institucion_educativa || '').toString().trim().toLowerCase() !== institucion) return;
      students.push({
        nombre: r.nombre_completo || '',
        rol: r.rol_asignado || r.rol_opcion1 || '',
        comision: r.comision_asignada || r.comision_opcion1 || '',
        partido: r.partido_asignado || r.partido || '',
      });
    });
  }
  return jsonOut_({ ok: true, institucion: record.institucion_educativa || '', students: students });
}

// ---------- Propuestas legislativas de los comisarios ----------
var PROPUESTAS_FOLDER_NAME_ = 'Propuestas Euromodelo Joven 2026';
var PROPUESTA_MAX_BYTES_ = 8 * 1024 * 1024; // ~8MB, de sobra para un PDF de propuesta

function findOrCreateFolder_(name) {
  var folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder(name);
}

function findOrCreateSubfolder_(parentFolder, name) {
  var folders = parentFolder.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parentFolder.createFolder(name);
}

// Nombre de carpeta legible para una comisión: reusa los mismos code/title de SIM_COMISIONES_
// (definidos más abajo) para que "CLJ" y el nombre completo queden alineados con el resto del
// proyecto. Si la comisión aún no está asignada, cae en una carpeta de "Sin comisión" en vez de
// romper la subida.
function comisionFolderName_(comisionTitle) {
  var match = SIM_COMISIONES_.filter(function(c) { return c.title === comisionTitle; })[0];
  if (match) return match.code + ' — ' + match.title;
  return comisionTitle ? comisionTitle : 'Sin comisión';
}

// Árbol de carpetas: <raíz>/Nacional/<comisión>  ó  <raíz>/Regionales/<ciudad>/<comisión>.
// Cada nivel se crea automáticamente la primera vez que hace falta (mismo patrón find-or-create
// que ya usa findOrCreateFolder_), así que no hay que crear nada a mano en Drive.
function resolvePropuestaFolder_(record) {
  var root = findOrCreateFolder_(PROPUESTAS_FOLDER_NAME_);
  var tipoEuromodelo = (record.tipo_euromodelo || 'Nacional').toString();
  var bloqueFolder;
  if (tipoEuromodelo === 'Regional') {
    var regionalesFolder = findOrCreateSubfolder_(root, 'Regionales');
    var ciudad = (record.ciudad || 'Sin ciudad').toString();
    bloqueFolder = findOrCreateSubfolder_(regionalesFolder, ciudad);
  } else {
    bloqueFolder = findOrCreateSubfolder_(root, 'Nacional');
  }
  var comisionName = comisionFolderName_((record.comision_asignada || '').toString());
  return findOrCreateSubfolder_(bloqueFolder, comisionName);
}

// El comisario reenvía sus propias credenciales (mismo criterio sin sesión de todo el proyecto).
// Solo puede subir propuesta si su rol_asignado oficial (no la preferencia) es 'Comisario'.
// Cada subida reemplaza la url/estado/comentario anteriores, igual que repetir una brújula.
function handleUploadPropuesta_(sheet, headers, data) {
  var email = (data.email || '').toString().trim().toLowerCase();
  var password = (data.password || '').toString();
  var rowNum = findRowByColumn_(sheet, headers, 'email', email, true);
  if (rowNum === -1) return jsonOut_({ ok: false });

  var rowValues = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  var record = {};
  headers.forEach(function(h, i) { record[h] = rowValues[i]; });

  if ((record.password || '').toString() !== password) return jsonOut_({ ok: false });
  if ((record.rol_asignado || '').toString() !== 'Comisario') return jsonOut_({ ok: false, error: 'not_comisario' });

  var fileBase64 = (data.fileBase64 || '').toString();
  if (!fileBase64) return jsonOut_({ ok: false, error: 'no_file' });
  var bytes;
  try {
    bytes = Utilities.base64Decode(fileBase64);
  } catch (err) {
    return jsonOut_({ ok: false, error: 'invalid_file' });
  }
  if (bytes.length > PROPUESTA_MAX_BYTES_) return jsonOut_({ ok: false, error: 'file_too_large' });

  var fileName = 'Propuesta - ' + (record.nombre_completo || email) + ' (' + (record.ref_code || '') + ').pdf';
  var blob = Utilities.newBlob(bytes, 'application/pdf', fileName);
  var folder = resolvePropuestaFolder_(record);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  var url = file.getUrl();

  ['propuesta_url', 'propuesta_estado', 'propuesta_comentario'].forEach(function(field) {
    var colIdx = headers.indexOf(field);
    if (colIdx === -1) return;
    var value = field === 'propuesta_url' ? url : (field === 'propuesta_estado' ? 'Pendiente' : '');
    var cell = sheet.getRange(rowNum, colIdx + 1);
    cell.setNumberFormat('@');
    cell.setValue(value);
  });

  return jsonOut_({ ok: true, url: url, estado: 'Pendiente' });
}

// propuesta_estado sigue esta progresión: '' -> 'Pendiente' -> 'Aprobada' -> 'Plenaria'. Una
// propuesta en 'Plenaria' sigue contando como aprobada dentro de su propia comisión (por eso
// PROPUESTA_ESTADOS_VISIBLES_COMISION_ incluye ambas), y además se vuelve visible para todos los
// participantes del bloque en handleListPlenariaPropuestas_, sin importar su comisión.
var PROPUESTA_ESTADOS_VISIBLES_COMISION_ = ['Aprobada', 'Plenaria'];

// El participante reenvía sus propias credenciales; solo ve las propuestas ya aprobadas de su
// propia comisión asignada (si todavía no tiene comisión asignada, la lista viene vacía).
function handleListComisionPropuestas_(sheet, headers, data) {
  var email = (data.email || '').toString().trim().toLowerCase();
  var password = (data.password || '').toString();
  var rowNum = findRowByColumn_(sheet, headers, 'email', email, true);
  if (rowNum === -1) return jsonOut_({ ok: false });

  var rowValues = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  var record = {};
  headers.forEach(function(h, i) { record[h] = rowValues[i]; });
  if ((record.password || '').toString() !== password) return jsonOut_({ ok: false });

  var comision = (record.comision_asignada || '').toString();
  if (!comision) return jsonOut_({ ok: true, comision: '', proposals: [] });
  // Misma sede que quien pregunta: el mismo nombre de comisión existe de forma independiente
  // en cada bloque (Nacional, Regional-Ciudad), así que no basta con comparar el nombre solo.
  var tipoEuromodelo = (record.tipo_euromodelo || 'Nacional').toString();
  var ciudad = (record.ciudad || '').toString();

  var lastRow = sheet.getLastRow();
  var proposals = [];
  if (lastRow >= 2) {
    var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    values.forEach(function(row) {
      var r = {};
      headers.forEach(function(h, i) { r[h] = row[i]; });
      if ((r.rol_asignado || '').toString() !== 'Comisario') return;
      if ((r.comision_asignada || '').toString() !== comision) return;
      if ((r.tipo_euromodelo || 'Nacional').toString() !== tipoEuromodelo) return;
      if (tipoEuromodelo === 'Regional' && (r.ciudad || '').toString() !== ciudad) return;
      if (PROPUESTA_ESTADOS_VISIBLES_COMISION_.indexOf((r.propuesta_estado || '').toString()) === -1) return;
      if (!r.propuesta_url) return;
      proposals.push({ nombre: r.nombre_completo || '', url: r.propuesta_url });
    });
  }
  return jsonOut_({ ok: true, comision: comision, proposals: proposals });
}

// Propuestas que pasaron a la sesión Plenaria: visibles para TODOS los participantes del mismo
// bloque (Nacional, o Regional de su misma ciudad), sin importar su propia comisión — a
// diferencia de handleListComisionPropuestas_, que solo muestra las de la comisión propia.
function handleListPlenariaPropuestas_(sheet, headers, data) {
  var email = (data.email || '').toString().trim().toLowerCase();
  var password = (data.password || '').toString();
  var rowNum = findRowByColumn_(sheet, headers, 'email', email, true);
  if (rowNum === -1) return jsonOut_({ ok: false });

  var rowValues = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  var record = {};
  headers.forEach(function(h, i) { record[h] = rowValues[i]; });
  if ((record.password || '').toString() !== password) return jsonOut_({ ok: false });

  var tipoEuromodelo = (record.tipo_euromodelo || 'Nacional').toString();
  var ciudad = (record.ciudad || '').toString();

  var lastRow = sheet.getLastRow();
  var proposals = [];
  if (lastRow >= 2) {
    var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    values.forEach(function(row) {
      var r = {};
      headers.forEach(function(h, i) { r[h] = row[i]; });
      if ((r.rol_asignado || '').toString() !== 'Comisario') return;
      if ((r.tipo_euromodelo || 'Nacional').toString() !== tipoEuromodelo) return;
      if (tipoEuromodelo === 'Regional' && (r.ciudad || '').toString() !== ciudad) return;
      if ((r.propuesta_estado || '').toString() !== 'Plenaria') return;
      if (!r.propuesta_url) return;
      proposals.push({ nombre: r.nombre_completo || '', comision: r.comision_asignada || '', url: r.propuesta_url });
    });
  }
  return jsonOut_({ ok: true, proposals: proposals });
}

// ---------- Panel de administración: login individual de staff ----------
// Pestaña "admins" aparte (email/password/mustChangePassword), auto-sembrada la primera vez que
// cada correo de STAFF_AUTHORIZED_EMAILS_ inicia sesión, con STAFF_KEY_ como contraseña inicial.
// No hay sesiones ni tokens (mismo criterio que el resto del proyecto): cada acción de solo-admin
// reenvía adminEmail/adminPassword y se valida contra esta pestaña en cada request.
var ADMIN_HEADERS_ = ['email', 'password', 'mustChangePassword'];

function isAuthorizedStaffEmail_(email) {
  var target = (email || '').toString().trim().toLowerCase();
  return STAFF_AUTHORIZED_EMAILS_.some(function(e) { return e.toLowerCase() === target; });
}

function ensureAdminSheet_(ss) {
  var sheet = findOrCreateSheet_(ss, ['admins', 'administradores'], 'admins');
  var headers = getHeaders_(sheet, ADMIN_HEADERS_);
  return { sheet: sheet, headers: headers };
}

// Crea la fila de un correo autorizado la primera vez que se necesita (login o recuperación).
// Devuelve el número de fila (1-indexado).
function ensureAdminRow_(sheet, headers, email) {
  var rowNum = findRowByColumn_(sheet, headers, 'email', email, true);
  if (rowNum !== -1) return rowNum;
  var row = headers.map(function(h) {
    if (h === 'email') return email;
    if (h === 'password') return STAFF_KEY_;
    if (h === 'mustChangePassword') return 'true';
    return '';
  });
  var targetRow = sheet.getLastRow() + 1;
  var range = sheet.getRange(targetRow, 1, 1, row.length);
  range.setNumberFormat('@');
  range.setValues([row]);
  return targetRow;
}

function readAdminRecord_(sheet, headers, rowNum) {
  var rowValues = sheet.getRange(rowNum, 1, 1, headers.length).getValues()[0];
  var record = {};
  headers.forEach(function(h, i) { record[h] = rowValues[i]; });
  return record;
}

function handleAdminLogin_(ss, data) {
  var email = (data.email || '').toString().trim().toLowerCase();
  if (!isAuthorizedStaffEmail_(email)) return jsonOut_({ ok: false });

  var admin = ensureAdminSheet_(ss);
  var rowNum = ensureAdminRow_(admin.sheet, admin.headers, email);
  var record = readAdminRecord_(admin.sheet, admin.headers, rowNum);
  var password = (data.password || '').toString();
  if ((record.password || '').toString() !== password) return jsonOut_({ ok: false });

  return jsonOut_({ ok: true, mustChangePassword: (record.mustChangePassword || '').toString() === 'true' });
}

function handleAdminForgotPassword_(ss, data) {
  var email = (data.email || '').toString().trim().toLowerCase();
  if (isAuthorizedStaffEmail_(email)) {
    var admin = ensureAdminSheet_(ss);
    var rowNum = ensureAdminRow_(admin.sheet, admin.headers, email);
    var newPassword = randomPassword_();
    var pwIdx = admin.headers.indexOf('password');
    var mustChangeIdx = admin.headers.indexOf('mustChangePassword');
    var pwCell = admin.sheet.getRange(rowNum, pwIdx + 1);
    pwCell.setNumberFormat('@');
    pwCell.setValue(newPassword);
    admin.sheet.getRange(rowNum, mustChangeIdx + 1).setValue('true');
    try {
      sendPasswordResetEmail_(email, '', newPassword);
    } catch (err) {
      console.error('No se pudo enviar el correo de recuperación de admin: ' + err);
    }
  }
  // Respuesta genérica siempre, para no revelar si el correo está autorizado.
  return jsonOut_({ ok: true });
}

function handleAdminUpdatePassword_(ss, data) {
  var email = (data.email || '').toString().trim().toLowerCase();
  var newPassword = (data.newPassword || '').toString();
  if (!isAuthorizedStaffEmail_(email) || !newPassword) return jsonOut_({ ok: false });

  var admin = ensureAdminSheet_(ss);
  var rowNum = findRowByColumn_(admin.sheet, admin.headers, 'email', email, true);
  if (rowNum === -1) return jsonOut_({ ok: false });
  var pwIdx = admin.headers.indexOf('password');
  var mustChangeIdx = admin.headers.indexOf('mustChangePassword');
  var pwCell = admin.sheet.getRange(rowNum, pwIdx + 1);
  pwCell.setNumberFormat('@');
  pwCell.setValue(newPassword);
  admin.sheet.getRange(rowNum, mustChangeIdx + 1).setValue('false');
  return jsonOut_({ ok: true });
}

function verifyAdminCredentials_(ss, email, password) {
  email = (email || '').toString().trim().toLowerCase();
  if (!isAuthorizedStaffEmail_(email)) return false;
  var admin = ensureAdminSheet_(ss);
  var rowNum = findRowByColumn_(admin.sheet, admin.headers, 'email', email, true);
  if (rowNum === -1) return false;
  var record = readAdminRecord_(admin.sheet, admin.headers, rowNum);
  return (record.password || '').toString() === (password || '').toString();
}

// ---------- Panel de administración: listar y editar participantes ----------
function handleListParticipants_(sheet, headers, ss, data) {
  if (!verifyAdminCredentials_(ss, data.adminEmail, data.adminPassword)) return jsonOut_({ ok: false });

  var lastRow = sheet.getLastRow();
  var records = [];
  if (lastRow >= 2) {
    var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    records = values.map(function(row) {
      var record = {};
      headers.forEach(function(h, i) { record[h] = row[i]; });
      delete record.password; // nunca se envía al cliente
      return record;
    });
  }
  return jsonOut_({ ok: true, records: records });
}

var ADMIN_ASSIGNABLE_FIELDS_ = ['rol_asignado', 'comision_asignada', 'partido_asignado', 'propuesta_estado', 'propuesta_comentario'];
function handleUpdateAssignment_(sheet, headers, ss, data) {
  if (!verifyAdminCredentials_(ss, data.adminEmail, data.adminPassword)) return jsonOut_({ ok: false });

  var rowNum = findRowByColumn_(sheet, headers, 'ref_code', data.ref_code);
  if (rowNum === -1) return jsonOut_({ ok: false, error: 'not_found' });

  ADMIN_ASSIGNABLE_FIELDS_.forEach(function(field) {
    if (data[field] === undefined) return;
    var colIdx = headers.indexOf(field);
    if (colIdx === -1) return;
    var cell = sheet.getRange(rowNum, colIdx + 1);
    cell.setNumberFormat('@');
    cell.setValue(data[field] || '');
  });

  // Si el payload trae rol/comisión/partido, es una edición de fila de participante (Guardar o
  // Borrar), no una revisión de propuesta: se recalcula el origen de la asignación.
  if (data.rol_asignado !== undefined || data.comision_asignada !== undefined || data.partido_asignado !== undefined) {
    var origenIdx = headers.indexOf('asignacion_origen');
    if (origenIdx !== -1) {
      var tieneAsignacion = data.rol_asignado || data.comision_asignada || data.partido_asignado;
      var origenCell = sheet.getRange(rowNum, origenIdx + 1);
      origenCell.setNumberFormat('@');
      origenCell.setValue(tieneAsignacion ? 'Manual' : '');
    }
  }

  return jsonOut_({ ok: true });
}

// ---------- Sorteo automático de rol/comisión/partido (aplica el resultado a la Sheet) ----------
// Reusa runAssignmentSimulation_ (el mismo algoritmo que la vista previa "Simulación y sorteo"
// del panel de admin), pero esta vez SÍ escribe en la Sheet. Respeta cualquier asignación
// manual existente: solo toca a quienes tienen
// rol_asignado/comision_asignada/partido_asignado en blanco los tres. Guarda qué filas tocó (por
// email) en PropertiesService para poder deshacer esta corrida con handleUndoAssignment_.
var LAST_ASSIGNMENT_BACKUP_KEY_ = 'LAST_ASSIGNMENT_BACKUP_';

// Un participante 'Regional' pertenece a la sede de su propia columna `ciudad` — no hace falta
// una columna de sede aparte. 'Nacional' es un solo bloque, sin importar la ciudad.
function filterByScope_(records, tipoEuromodelo, ciudad) {
  return records.filter(function(r) {
    var tipo = (r.tipo_euromodelo || 'Nacional').toString();
    if (tipo !== tipoEuromodelo) return false;
    if (tipoEuromodelo === 'Regional' && (r.ciudad || '').toString() !== ciudad) return false;
    return true;
  });
}

// Clave de PropertiesService para el backup de "deshacer", una por bloque — así sortear en una
// sede no pisa el histórico de deshacer de otra.
function scopeBackupKey_(tipoEuromodelo, ciudad) {
  var key = LAST_ASSIGNMENT_BACKUP_KEY_ + (tipoEuromodelo || 'Nacional');
  if (tipoEuromodelo === 'Regional') key += '_' + (ciudad || '');
  return key;
}

function handleApplyAssignment_(sheet, headers, ss, data) {
  if (!verifyAdminCredentials_(ss, data.adminEmail, data.adminPassword)) return jsonOut_({ ok: false });

  var tipoEuromodelo = (data.tipoEuromodelo || 'Nacional').toString();
  var ciudad = (data.ciudad || '').toString();
  var mesaCount = Number(data.mesaCount);
  if (!isFinite(mesaCount) || mesaCount < 0) mesaCount = 2;

  var lastRow = sheet.getLastRow();
  var records = [];
  if (lastRow >= 2) {
    var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    records = values.map(function(row) {
      var record = {};
      headers.forEach(function(h, i) { record[h] = row[i]; });
      return record;
    });
  }
  records = filterByScope_(records, tipoEuromodelo, ciudad);

  var summary = runAssignmentSimulation_(records, mesaCount);
  var rolIdx = headers.indexOf('rol_asignado');
  var comIdx = headers.indexOf('comision_asignada');
  var parIdx = headers.indexOf('partido_asignado');
  var origenIdx = headers.indexOf('asignacion_origen');

  var touchedEmails = [];
  var applied = 0, skipped = 0;

  summary.participantes.forEach(function(p) {
    var rowNum = findRowByColumn_(sheet, headers, 'email', p.email, true);
    if (rowNum === -1) return;
    var currentRol = rolIdx !== -1 ? sheet.getRange(rowNum, rolIdx + 1).getValue() : '';
    var currentCom = comIdx !== -1 ? sheet.getRange(rowNum, comIdx + 1).getValue() : '';
    var currentPar = parIdx !== -1 ? sheet.getRange(rowNum, parIdx + 1).getValue() : '';
    if (currentRol || currentCom || currentPar) { skipped++; return; }

    if (rolIdx !== -1) { var c1 = sheet.getRange(rowNum, rolIdx + 1); c1.setNumberFormat('@'); c1.setValue(p.rol); }
    if (comIdx !== -1) { var c2 = sheet.getRange(rowNum, comIdx + 1); c2.setNumberFormat('@'); c2.setValue(p.comision); }
    if (parIdx !== -1) { var c3 = sheet.getRange(rowNum, parIdx + 1); c3.setNumberFormat('@'); c3.setValue(p.partido); }
    if (origenIdx !== -1) { var c4 = sheet.getRange(rowNum, origenIdx + 1); c4.setNumberFormat('@'); c4.setValue('Automática'); }
    touchedEmails.push(p.email);
    applied++;
  });

  PropertiesService.getScriptProperties().setProperty(scopeBackupKey_(tipoEuromodelo, ciudad), JSON.stringify(touchedEmails));
  return jsonOut_({ ok: true, applied: applied, skipped: skipped });
}

function handleUndoAssignment_(sheet, headers, ss, data) {
  if (!verifyAdminCredentials_(ss, data.adminEmail, data.adminPassword)) return jsonOut_({ ok: false });

  var tipoEuromodelo = (data.tipoEuromodelo || 'Nacional').toString();
  var ciudad = (data.ciudad || '').toString();
  var props = PropertiesService.getScriptProperties();
  var backupKey = scopeBackupKey_(tipoEuromodelo, ciudad);
  var raw = props.getProperty(backupKey);
  if (!raw) return jsonOut_({ ok: false, error: 'nothing_to_undo' });

  var emails = JSON.parse(raw);
  var rolIdx = headers.indexOf('rol_asignado');
  var comIdx = headers.indexOf('comision_asignada');
  var parIdx = headers.indexOf('partido_asignado');
  var origenIdx = headers.indexOf('asignacion_origen');

  var restored = 0;
  emails.forEach(function(email) {
    var rowNum = findRowByColumn_(sheet, headers, 'email', email, true);
    if (rowNum === -1) return;
    if (rolIdx !== -1) { var c1 = sheet.getRange(rowNum, rolIdx + 1); c1.setNumberFormat('@'); c1.setValue(''); }
    if (comIdx !== -1) { var c2 = sheet.getRange(rowNum, comIdx + 1); c2.setNumberFormat('@'); c2.setValue(''); }
    if (parIdx !== -1) { var c3 = sheet.getRange(rowNum, parIdx + 1); c3.setNumberFormat('@'); c3.setValue(''); }
    if (origenIdx !== -1) { var c4 = sheet.getRange(rowNum, origenIdx + 1); c4.setNumberFormat('@'); c4.setValue(''); }
    restored++;
  });

  props.deleteProperty(backupKey);
  return jsonOut_({ ok: true, restored: restored });
}

// ---------- Simulación de asignación de rol / comisión / partido ----------
// Corre bajo demanda, de solo lectura: nunca escribe en la Sheet. Pensada para reusarse más
// adelante como la asignación FINAL real (agregando un modo que sí persista los resultados),
// una vez cierren las inscripciones y se sepa el número definitivo de participantes.
var SIM_MIN_PER_COMISION_ = 10;

// Cupo de Primer Ministro: proporcional a los inscritos (1 por cada 10), redondeado al entero
// más cercano y ajustado hacia arriba si cae en un número par, porque el número final de
// Primeros Ministros siempre debe ser impar. Comisario y Europarlamentario no tienen tope.
var SIM_PM_RATIO_ = 1 / 10;

function computePrimerMinistroCap_(totalN) {
  if (totalN <= 0) return 0;
  var cap = Math.round(totalN * SIM_PM_RATIO_);
  if (cap < 1) cap = 1;
  if (cap % 2 === 0) cap += 1;
  return cap;
}

var SIM_COMISIONES_ = [
  { code: 'CLJ', title: 'Asuntos Constitucionales, Libertades Civiles, Justicia e Igualdad de Género', color: '#6B2D8C' },
  { code: 'EXT', title: 'Asuntos Exteriores, Derechos Humanos, Seguridad y Defensa', color: '#1E3D73' },
  { code: 'ECO', title: 'Asuntos Presupuestarios, Económicos y Monetarios', color: '#C9A227' },
  { code: 'CED', title: 'Cultura, Educación y Desarrollo Regional', color: '#2E7D6B' },
  { code: 'EAS', title: 'Empleo, Asuntos Sociales y Comercio Internacional', color: '#E07B39' },
  { code: 'IIE', title: 'Industria, Investigación y Energía', color: '#4E8CC7' },
  { code: 'MAS', title: 'Medio Ambiente, Agricultura, Desarrollo Rural y Salud Pública', color: '#3E9142' }
];

// Escaños reales del Parlamento Europeo, usados para escalar el cupo de partido al número de
// inscritos actual. Códigos/títulos alineados con los que ya usa perfil.html (no con
// partidos.html, que todavía usa códigos viejos para las mismas 7 bancadas).
var SIM_PARTIDOS_ = [
  { code: 'PPE', title: 'PPE — Partido Popular Europeo', color: '#2D5FA8', seats: 188 },
  { code: 'S&D', title: 'S&D — Alianza Progresista Socialista Demócrata', color: '#E33241', seats: 136 },
  { code: 'Patriots', title: 'Patriots — Patriotas por Europa', color: '#142850', seats: 84 },
  { code: 'ECR', title: 'ECR Group — Grupo de los Conservadores y Reformistas Europeos', color: '#4E8CC7', seats: 78 },
  { code: 'Renew', title: 'Renew Europe — Renovar Europa', color: '#F2B705', seats: 77 },
  { code: 'Greens/EFA', title: 'The Greens/EFA — Los Verdes/Alianza Libre Europea', color: '#3E9142', seats: 53 },
  { code: 'The Left', title: 'The Left — Izquierda Unitaria Europea/Verde Nórdica', color: '#8C1D2B', seats: 46 }
];

// Asigna rol respetando el cupo de Primer Ministro: prueba las 3 opciones del participante en
// orden y solo se detiene en 'Primer Ministro' si todavía queda cupo; si no, sigue a la
// siguiente opción (que nunca vuelve a ser 'Primer Ministro', porque cada participante lo tiene
// en una sola de sus 3 opciones). Comisario y Europarlamentario siempre tienen cupo libre.
function assignRol_(record, pmCap, pmCount) {
  var opciones = [record.rol_opcion1, record.rol_opcion2, record.rol_opcion3];
  for (var i = 0; i < opciones.length; i++) {
    var opcion = opciones[i];
    if (opcion !== 'Primer Ministro') return opcion;
    if (pmCount.total < pmCap) {
      pmCount.total++;
      return opcion;
    }
    // Cupo de Primer Ministro lleno: se prueba con la siguiente opción del participante.
  }
  return opciones[0] || '';
}

function assignPartido_(record, cap, count) {
  var preferido = record.partido;
  var brujula = record.resultado_brujula_partido;

  if (preferido && count[preferido] !== undefined && count[preferido] < cap[preferido]) {
    count[preferido]++;
    return { partido: preferido, origen: 'preferencia' };
  }
  if (brujula && brujula !== preferido && count[brujula] !== undefined && count[brujula] < cap[brujula]) {
    count[brujula]++;
    return { partido: brujula, origen: 'brujula' };
  }
  // Ninguna de las dos tiene cupo: se asigna al partido con más cupo libre restante.
  var best = null, bestFree = -1;
  SIM_PARTIDOS_.forEach(function(p) {
    var free = cap[p.title] - count[p.title];
    if (free > bestFree) { bestFree = free; best = p.title; }
  });
  if (best) count[best]++;
  return { partido: best || preferido || '', origen: 'cupo_libre' };
}

// Elige, entre las comisiones con superávit sobre el mínimo, al mejor candidato para mover hacia
// targetTitle: primero por afinidad (2ª/3ª opción o resultado de la Brújula de Comisión), luego
// por orden de llegada (se mueve a quien llegó más tarde, protegiendo a los primeros), y por
// último con un desempate suave por colegio (se prioriza no amontonar el mismo colegio).
function pickCandidateForRepair_(assignments, totals, targetTitle) {
  var pool = assignments.filter(function(a) {
    var surplus = (totals[a.comision] || 0) - SIM_MIN_PER_COMISION_;
    return a.comision !== targetTitle && surplus > 0;
  });
  if (pool.length === 0) return null;

  function affinityScore(a) {
    if (a.comisionPrefs[1] === targetTitle) return 3; // 2ª opción
    if (a.comisionPrefs[2] === targetTitle) return 2; // 3ª opción
    if (a.comisionBrujula === targetTitle) return 1; // Brújula de Comisión
    return 0;
  }

  var targetSchoolCount = {};
  assignments.forEach(function(a) {
    if (a.comision === targetTitle) {
      targetSchoolCount[a.institucion] = (targetSchoolCount[a.institucion] || 0) + 1;
    }
  });

  pool.sort(function(a, b) {
    var diff = affinityScore(b) - affinityScore(a);
    if (diff !== 0) return diff;
    var da = new Date(a.enviado).getTime() || 0;
    var db = new Date(b.enviado).getTime() || 0;
    if (db !== da) return db - da; // llegó más tarde primero (protege a los que llegaron antes)
    var ca = targetSchoolCount[a.institucion] || 0;
    var cb = targetSchoolCount[b.institucion] || 0;
    return ca - cb; // colegio menos representado en destino, primero
  });

  return pool[0];
}

function repairComisionMinimums_(assignments, mesaCount) {
  function computeTotals() {
    var totals = {};
    SIM_COMISIONES_.forEach(function(c) { totals[c.title] = mesaCount; });
    assignments.forEach(function(a) { totals[a.comision] = (totals[a.comision] || mesaCount) + 1; });
    return totals;
  }

  var totals = computeTotals();
  var under = SIM_COMISIONES_.filter(function(c) { return (totals[c.title] || 0) < SIM_MIN_PER_COMISION_; });
  under.sort(function(a, b) { return (totals[a.title] || 0) - (totals[b.title] || 0); }); // más urgente primero

  under.forEach(function(target) {
    var needed = SIM_MIN_PER_COMISION_ - (totals[target.title] || 0);
    for (var i = 0; i < needed; i++) {
      var candidate = pickCandidateForRepair_(assignments, totals, target.title);
      if (!candidate) break; // no hay más candidatos disponibles: queda por debajo del mínimo
      var origenAnterior = candidate.comision;
      totals[origenAnterior] = (totals[origenAnterior] || 0) - 1;
      candidate.comision = target.title;
      candidate.comisionOrigen = origenAnterior === candidate.comisionPrefs[1] ? 'reparacion_desde_preferencia_2'
        : origenAnterior === candidate.comisionPrefs[2] ? 'reparacion_desde_preferencia_3'
        : 'reparacion';
      totals[target.title] = (totals[target.title] || 0) + 1;
    }
  });
}

function buildSimulationSummary_(assignments, partidoCap, partidoCount, totalN, mesaCount) {
  var comisionSummary = SIM_COMISIONES_.map(function(c) {
    var members = assignments.filter(function(a) { return a.comision === c.title; });
    var comisarios = members.filter(function(a) { return a.rol === 'Comisario'; }).length;
    var diputados = members.length - comisarios;
    var totalConMesa = members.length + mesaCount;
    return {
      code: c.code, title: c.title, color: c.color,
      debate: members.length, mesa: mesaCount, total: totalConMesa,
      comisarios: comisarios, diputados: diputados,
      comisariosPct: members.length > 0 ? Math.round(comisarios / members.length * 100) : 0,
      cumpleMinimo: totalConMesa >= SIM_MIN_PER_COMISION_,
      faltan: Math.max(0, SIM_MIN_PER_COMISION_ - totalConMesa),
    };
  });

  var partidoSummary = SIM_PARTIDOS_.map(function(p) {
    var cupo = partidoCap[p.title] || 0;
    var asignados = partidoCount[p.title] || 0;
    return {
      code: p.code, title: p.title, color: p.color,
      cupo: cupo, asignados: asignados,
      pctLleno: cupo > 0 ? Math.round(asignados / cupo * 100) : 0,
    };
  });

  var participantes = assignments.map(function(a) {
    return {
      email: a.email, nombre: a.nombre, institucion: a.institucion,
      rol: a.rol, comision: a.comision, comisionOrigen: a.comisionOrigen,
      partido: a.partido, partidoOrigen: a.partidoOrigen,
    };
  });

  return { totalInscritos: totalN, comisiones: comisionSummary, partidos: partidoSummary, participantes: participantes };
}

function runAssignmentSimulation_(records, mesaCount) {
  // Solo entran quienes ya completaron la inscripción (tienen rol_opcion1).
  var participants = records.filter(function(r) { return r.rol_opcion1; });

  // Orden de llegada = fecha de envío de la inscripción.
  participants.sort(function(a, b) {
    var da = new Date(a.enviado).getTime() || 0;
    var db = new Date(b.enviado).getTime() || 0;
    return da - db;
  });

  var totalN = participants.length;

  // Cupo proporcional a escaños reales, por método de "mayor resto" (Hamilton): así la suma de
  // los cupos siempre da exactamente totalN, sin que el redondeo simple deje a alguien sin
  // partido posible o haga que uno termine por encima del 100% solo por el redondeo.
  var totalSeats = SIM_PARTIDOS_.reduce(function(s, p) { return s + p.seats; }, 0);
  var partidoCap = {}, partidoCount = {};
  if (totalN > 0) {
    var exactShares = SIM_PARTIDOS_.map(function(p) {
      var exact = p.seats / totalSeats * totalN;
      return { title: p.title, floor: Math.floor(exact), remainder: exact - Math.floor(exact) };
    });
    var assignedSoFar = exactShares.reduce(function(s, x) { return s + x.floor; }, 0);
    var remaining = totalN - assignedSoFar;
    exactShares.sort(function(a, b) { return b.remainder - a.remainder; });
    exactShares.forEach(function(x, i) {
      partidoCap[x.title] = x.floor + (i < remaining ? 1 : 0);
    });
  } else {
    SIM_PARTIDOS_.forEach(function(p) { partidoCap[p.title] = 0; });
  }
  SIM_PARTIDOS_.forEach(function(p) { partidoCount[p.title] = 0; });

  var pmCap = computePrimerMinistroCap_(totalN);
  var pmCount = { total: 0 };

  // Pasada 1: asignación directa. Sin tope de comisión por ahora ("no hay máximo por el
  // momento"), así que comisión es siempre la 1ª preferencia; partido y Primer Ministro sí
  // tienen cupo real.
  var assignments = participants.map(function(p) {
    var partidoResult = assignPartido_(p, partidoCap, partidoCount);
    return {
      email: p.email, nombre: p.nombre_completo, institucion: p.institucion_educativa, enviado: p.enviado,
      rol: assignRol_(p, pmCap, pmCount),
      comision: p.comision_opcion1,
      comisionOrigen: 'preferencia_1',
      comisionPrefs: [p.comision_opcion1, p.comision_opcion2, p.comision_opcion3],
      comisionBrujula: p.resultado_brujula_comision,
      partido: partidoResult.partido,
      partidoOrigen: partidoResult.origen,
    };
  });

  // Pasada 2: repara el mínimo de 10 por comisión (incluyendo mesa directiva) moviendo gente de
  // comisiones con superávit hacia las que están por debajo.
  repairComisionMinimums_(assignments, mesaCount);

  return buildSimulationSummary_(assignments, partidoCap, partidoCount, totalN, mesaCount);
}

function handleSimulate_(sheet, headers, ss, data) {
  if (!verifyAdminCredentials_(ss, data.adminEmail, data.adminPassword)) {
    return jsonOut_({ ok: false, error: 'unauthorized' });
  }
  var tipoEuromodelo = (data.tipoEuromodelo || 'Nacional').toString();
  var ciudad = (data.ciudad || '').toString();
  var mesaCount = Number(data.mesaCount);
  if (!isFinite(mesaCount) || mesaCount < 0) mesaCount = 2;

  var lastRow = sheet.getLastRow();
  var records = [];
  if (lastRow >= 2) {
    var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    records = values.map(function(row) {
      var record = {};
      headers.forEach(function(h, i) { record[h] = row[i]; });
      return record;
    });
  }
  records = filterByScope_(records, tipoEuromodelo, ciudad);

  var summary = runAssignmentSimulation_(records, mesaCount);
  return jsonOut_({ ok: true, summary: summary });
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Una sola pestaña para todo: la preinscripción crea la fila con los campos
  // de inscripción en blanco; la inscripción completa esa misma fila (por ref_code).
  // También sirve como tabla de usuarios para login / recuperación de contraseña.
  var sheet = findOrCreateSheet_(ss, ['preinscripciones', 'preinscripcion', 'preinscripción'], 'preinscripciones');
  var headers = getHeaders_(sheet, CANONICAL_HEADERS_);

  if (data.form === 'login') return handleLogin_(sheet, headers, data);
  if (data.form === 'forgot_password') return handleForgotPassword_(sheet, headers, data);
  if (data.form === 'update_password') return handleUpdatePassword_(sheet, headers, data);
  if (data.form === 'update_brujula') return handleUpdateBrujula_(sheet, headers, data);
  if (data.form === 'list_students') return handleListStudents_(sheet, headers, data);
  if (data.form === 'upload_propuesta') return handleUploadPropuesta_(sheet, headers, data);
  if (data.form === 'list_comision_propuestas') return handleListComisionPropuestas_(sheet, headers, data);
  if (data.form === 'list_plenaria_propuestas') return handleListPlenariaPropuestas_(sheet, headers, data);
  if (data.form === 'admin_simulate') return handleSimulate_(sheet, headers, ss, data);
  if (data.form === 'admin_login') return handleAdminLogin_(ss, data);
  if (data.form === 'admin_forgot_password') return handleAdminForgotPassword_(ss, data);
  if (data.form === 'admin_update_password') return handleAdminUpdatePassword_(ss, data);
  if (data.form === 'admin_list_participants') return handleListParticipants_(sheet, headers, ss, data);
  if (data.form === 'admin_update_assignment') return handleUpdateAssignment_(sheet, headers, ss, data);
  if (data.form === 'admin_apply_assignment') return handleApplyAssignment_(sheet, headers, ss, data);
  if (data.form === 'admin_undo_assignment') return handleUndoAssignment_(sheet, headers, ss, data);

  var isPre = data.form === 'preinscripcion';

  if (isPre) {
    var row = headers.map(function(columnName) { return data[columnName] || ''; });
    var targetRow = sheet.getLastRow() + 1;
    var range = sheet.getRange(targetRow, 1, 1, row.length);
    range.setNumberFormat('@'); // fuerza texto plano: evita que Sheets convierta documentos/teléfonos en número o fórmula
    range.setValues([row]);
  } else {
    var existingRow = findRowByColumn_(sheet, headers, 'ref_code', data.ref_code);
    if (existingRow !== -1) {
      INSCRIPCION_FIELDS_.forEach(function(field) {
        var colIdx = headers.indexOf(field);
        if (colIdx === -1) return;
        var cell = sheet.getRange(existingRow, colIdx + 1);
        cell.setNumberFormat('@');
        cell.setValue(data[field] || '');
      });
    } else {
      // No se encontró la fila de preinscripción original (caso raro): se agrega como fila nueva
      // para no perder los datos de la inscripción.
      var row2 = headers.map(function(columnName) { return data[columnName] || ''; });
      var targetRow2 = sheet.getLastRow() + 1;
      var range2 = sheet.getRange(targetRow2, 1, 1, row2.length);
      range2.setNumberFormat('@');
      range2.setValues([row2]);
    }
  }

  try {
    var sheetUrl = ss.getUrl() + '#gid=' + sheet.getSheetId();
    sendNotificationEmail_(isPre, data, CANONICAL_HEADERS_, sheetUrl);
    if (isPre && data.email) {
      sendParticipantWelcomeEmail_(data);
    }
  } catch (err) {
    // No queremos que un fallo de correo tumbe el guardado en la Sheet.
    console.error('No se pudo enviar el correo de notificación: ' + err);
  }

  return jsonOut_({ ok: true });
}
