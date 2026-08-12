var NOTIFY_EMAILS_ = ['juan.ladino@fundacionrevel.net', 'juan.ovalle@fundacionrevel.net'];
var SITE_URL_ = 'https://acabrera93.github.io/Euromodelo/';

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
  autoriza_datos: 'Autoriza tratamiento de datos',
  es_menor: 'Es menor de edad',
  autoriza_imagen: 'Autoriza derechos de imagen',
  rol_asignado: 'Rol asignado',
  comision_asignada: 'Comisión asignada',
  partido_asignado: 'Partido asignado'
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
  var body =
    '<p style="color:#3A4A63; font-size:14.5px; line-height:1.6; margin:0 0 18px;">Hola <b>' + escapeHtml_(data.nombre_completo || '') + '</b>, gracias por preinscribirte al <b>XX Euromodelo Joven 2026</b>. Guarda tus credenciales: las necesitarás para iniciar sesión y completar tu inscripción.</p>' +
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
  };
  if (record.rol_opcion1) {
    user.inscripcion = {
      rol1: record.rol_opcion1, rol2: record.rol_opcion2, rol3: record.rol_opcion3,
      comision1: record.comision_opcion1, comision2: record.comision_opcion2, comision3: record.comision_opcion3,
      partido: record.partido,
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
  // Estas tres columnas quedan en blanco al enviar el formulario: el staff las completa
  // manualmente en la Sheet una vez define la asignación final de cada participante.
  'rol_asignado', 'comision_asignada', 'partido_asignado'
];

var INSCRIPCION_FIELDS_ = [
  'rol_opcion1', 'rol_opcion2', 'rol_opcion3',
  'comision_opcion1', 'comision_opcion2', 'comision_opcion3',
  'partido'
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
