function doGet(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Crear cabeceras si la hoja está vacía
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Fecha y Hora', 'Nombre', 'Adultos', 'Niños']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.getRange(1, 1, 1, 4).setBackground('#4a5828');
    sheet.getRange(1, 1, 1, 4).setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  var nombre  = e.parameter.nombre  || '';
  var adultos = e.parameter.adultos || '1';
  var ninos   = e.parameter.ninos   || '0';

  sheet.appendRow([new Date(), nombre, adultos, ninos]);

  return ContentService.createTextOutput('OK');
}
