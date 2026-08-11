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

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var isPre = data.form === 'preinscripcion';

  var canonicalHeaders = isPre
    ? ['enviado', 'ref_code', 'email', 'nombre_completo', 'tipo_documento', 'numero_documento',
       'telefono', 'ciudad', 'institucion_educativa', 'autoriza_datos', 'es_menor', 'autoriza_imagen']
    : ['enviado', 'ref_code', 'email', 'nombre_completo', 'tipo_documento', 'numero_documento',
       'telefono', 'ciudad', 'institucion_educativa',
       'rol_opcion1', 'rol_opcion2', 'rol_opcion3',
       'comision_opcion1', 'comision_opcion2', 'comision_opcion3',
       'partido', 'autoriza_datos', 'es_menor', 'autoriza_imagen'];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = isPre
    ? findOrCreateSheet_(ss, ['preinscripciones', 'preinscripcion', 'preinscripción'], 'preinscripciones')
    : findOrCreateSheet_(ss, ['inscripciones', 'inscripcion', 'inscripción'], 'inscripciones');

  var headers = getHeaders_(sheet, canonicalHeaders);

  var row = headers.map(function(columnName) { return data[columnName] || ''; });
  var targetRow = sheet.getLastRow() + 1;
  var range = sheet.getRange(targetRow, 1, 1, row.length);
  range.setNumberFormat('@'); // fuerza texto plano: evita que Sheets convierta documentos/teléfonos en número o fórmula
  range.setValues([row]);

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
