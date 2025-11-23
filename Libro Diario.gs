function setupLibroDiario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Libro Diario');
  if (!sheet) {
    sheet = ss.insertSheet('Libro Diario');
  }

  var headers = ["ID Asiento", "Fecha", "Descripción", "Código de Cuenta", "Nombre de Cuenta", "Debe", "Haber"];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  sheet.setColumnWidths(1, 1, 100); // ID Asiento
  sheet.setColumnWidths(2, 1, 120); // Fecha
  sheet.setColumnWidths(3, 1, 300); // Descripción
  sheet.setColumnWidths(4, 1, 120); // Código de Cuenta
  sheet.setColumnWidths(5, 1, 200); // Nombre de Cuenta
  sheet.setColumnWidths(6, 2, 120); // Debe y Haber

  sheet.getRange("F:G").setNumberFormat("#,##0.00");

  sheet.setFrozenRows(1);
}