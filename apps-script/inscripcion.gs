var NOTIFY_EMAILS_ = ['juan.ladino@fundacionrevel.net', 'juan.ovalle@fundacionrevel.net'];

var FIELD_LABELS_ = {
  enviado: 'Fecha de envío',
  ref_code: 'Código de referencia',
  email: 'Correo electrónico',
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
  autoriza_imagen: 'Autoriza derechos de imagen'
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
function getHeaders_(sheet, canonicalHeaders) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(canonicalHeaders);
    return canonicalHeaders;
  }
  var lastCol = sheet.getLastColumn();
  return sheet.getRange(1, 1, 1, lastCol).getValues()[0];
}

function escapeHtml_(s) {
  return s.toString()
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

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

  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif; max-width:600px; margin:0 auto; background:#EDF1F6; padding:24px 0;">' +
      '<div style="max-width:600px; margin:0 auto; background:#ffffff; border-radius:14px; overflow:hidden; border:1px solid #D7DEE8;">' +
        '<div style="background:#0B2545; padding:24px 28px;">' +
          '<div style="color:#C9A227; font-family:\'Courier New\',monospace; font-size:11px; letter-spacing:2px; text-transform:uppercase;">XX Euromodelo Joven 2026</div>' +
          '<div style="color:#ffffff; font-size:21px; font-weight:700; margin-top:6px;">Nueva ' + tipoLabel.toLowerCase() + '</div>' +
        '</div>' +
        '<div style="padding:24px 28px;">' +
          '<p style="color:#4A5A73; font-size:14px; line-height:1.6; margin:0 0 18px;">Se registró una nueva <b>' + tipoLabel.toLowerCase() + '</b> en el sitio web del Euromodelo Joven 2026. Estos son los datos enviados:</p>' +
          '<table style="width:100%; border-collapse:collapse;">' + rowsHtml + '</table>' +
          '<div style="text-align:center; margin-top:26px;">' +
            '<a href="' + sheetUrl + '" style="display:inline-block; background:#1E3D73; color:#ffffff; text-decoration:none; font-weight:700; font-size:13.5px; padding:13px 28px; border-radius:8px;">Ver en la Sheet →</a>' +
          '</div>' +
        '</div>' +
        '<div style="background:#EDF1F6; padding:16px 28px; font-size:11.5px; color:#8695AC;">Fundación Revel — Notificación automática, no responder a este correo.</div>' +
      '</div>' +
    '</div>';

  MailApp.sendEmail({
    to: NOTIFY_EMAILS_.join(','),
    subject: subject,
    htmlBody: html,
  });
}

// Busca, en la columna ref_code, la fila que coincide con el código dado.
// Devuelve el número de fila en la Sheet (1-indexado) o -1 si no la encuentra.
function findRowByRefCode_(sheet, headers, refCode) {
  if (!refCode) return -1;
  var refColIdx = headers.indexOf('ref_code');
  if (refColIdx === -1) return -1;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return -1;
  var values = sheet.getRange(2, refColIdx + 1, lastRow - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === refCode) return i + 2;
  }
  return -1;
}

var INSCRIPCION_FIELDS_ = [
  'rol_opcion1', 'rol_opcion2', 'rol_opcion3',
  'comision_opcion1', 'comision_opcion2', 'comision_opcion3',
  'partido'
];

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var isPre = data.form === 'preinscripcion';

  // Una sola pestaña para todo: la preinscripción crea la fila con los campos
  // de inscripción en blanco; la inscripción completa esa misma fila (por ref_code)
  // en lugar de crear una fila nueva.
  var canonicalHeaders = [
    'enviado', 'ref_code', 'email', 'nombre_completo', 'tipo_documento', 'numero_documento',
    'telefono', 'ciudad', 'institucion_educativa', 'autoriza_datos', 'es_menor', 'autoriza_imagen',
    'rol_opcion1', 'rol_opcion2', 'rol_opcion3',
    'comision_opcion1', 'comision_opcion2', 'comision_opcion3', 'partido'
  ];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findOrCreateSheet_(ss, ['preinscripciones', 'preinscripcion', 'preinscripción'], 'preinscripciones');
  var headers = getHeaders_(sheet, canonicalHeaders);

  if (isPre) {
    var row = headers.map(function(columnName) { return data[columnName] || ''; });
    var targetRow = sheet.getLastRow() + 1;
    var range = sheet.getRange(targetRow, 1, 1, row.length);
    range.setNumberFormat('@'); // fuerza texto plano: evita que Sheets convierta documentos/teléfonos en número o fórmula
    range.setValues([row]);
  } else {
    var existingRow = findRowByRefCode_(sheet, headers, data.ref_code);
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
    sendNotificationEmail_(isPre, data, canonicalHeaders, sheetUrl);
  } catch (err) {
    // No queremos que un fallo de correo tumbe el guardado en la Sheet.
    console.error('No se pudo enviar el correo de notificación: ' + err);
  }

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
