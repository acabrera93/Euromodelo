function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  var data = JSON.parse(e.postData.contents);

  var headers = [
    'enviado', 'ref_code', 'email', 'nombre_completo', 'tipo_documento', 'numero_documento',
    'telefono', 'ciudad', 'institucion_educativa',
    'rol_opcion1', 'rol_opcion2', 'rol_opcion3',
    'comision_opcion1', 'comision_opcion2', 'comision_opcion3',
    'partido', 'autoriza_datos', 'es_menor', 'autoriza_imagen'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  }

  var row = headers.map(function(key) { return data[key] || ''; });
  var targetRow = sheet.getLastRow() + 1;
  var range = sheet.getRange(targetRow, 1, 1, row.length);
  range.setNumberFormat('@'); // fuerza texto plano: evita que Sheets convierta documentos/teléfonos en número o fórmula
  range.setValues([row]);

  return ContentService.createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}
