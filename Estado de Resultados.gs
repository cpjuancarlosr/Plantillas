function setupEstadoDeResultados() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Get chart of accounts data
  var coaSheet = ss.getSheetByName('Plan de Cuentas');
  if (!coaSheet) {
    ui.alert('No se encuentra la hoja "Plan de Cuentas". Por favor, ejecute la configuración del sistema primero.');
    return;
  }
  var accounts = coaSheet.getRange(2, 1, coaSheet.getLastRow() - 1, 4).getValues();

  // Filter for income and expense accounts
  var incomeAccounts = accounts.filter(function(acc) { return acc[2] === 'Ingresos' && acc[1] !== 'Ingresos'; });
  var expenseAccounts = accounts.filter(function(acc) { return acc[2] === 'Egresos' && acc[1] !== 'Egresos'; });

  // Prepare the sheet
  var sheet = ss.getSheetByName('Estado de Resultados');
  if (!sheet) {
    sheet = ss.insertSheet('Estado de Resultados');
  }
  sheet.clear();
  sheet.setColumnWidths(1, 1, 300); // Account Name
  sheet.setColumnWidths(2, 1, 150); // Value

  // --- Build the Report ---
  var currentRow = 1;

  // Title
  sheet.getRange(currentRow, 1, 1, 2).merge().setValue('Estado de Resultados')
       .setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  currentRow += 2;

  // Income Section
  sheet.getRange(currentRow, 1).setValue('Ingresos Operacionales').setFontWeight('bold');
  var incomeStartRow = currentRow + 1;
  incomeAccounts.forEach(function(acc) {
    currentRow++;
    sheet.getRange(currentRow, 1).setValue(acc[1]); // Account Name
    sheet.getRange(currentRow, 2).setFormula("=SUMIF('Libro Diario'!D:D, " + acc[0] + ", 'Libro Diario'!G:G) - SUMIF('Libro Diario'!D:D, " + acc[0] + ", 'Libro Diario'!F:F)");
  });
  var incomeEndRow = currentRow;
  currentRow++;
  sheet.getRange(currentRow, 1).setValue('Total Ingresos').setFontWeight('bold');
  sheet.getRange(currentRow, 2).setFormula('=SUM(B' + incomeStartRow + ':B' + incomeEndRow + ')').setFontWeight('bold');
  var totalIncomeCell = 'B' + currentRow;
  currentRow += 2;

  // Expense Section
  sheet.getRange(currentRow, 1).setValue('Costos y Gastos Operacionales').setFontWeight('bold');
  var expenseStartRow = currentRow + 1;
  expenseAccounts.forEach(function(acc) {
    currentRow++;
    sheet.getRange(currentRow, 1).setValue(acc[1]); // Account Name
    sheet.getRange(currentRow, 2).setFormula("=SUMIF('Libro Diario'!D:D, " + acc[0] + ", 'Libro Diario'!F:F) - SUMIF('Libro Diario'!D:D, " + acc[0] + ", 'Libro Diario'!G:G)");
  });
  var expenseEndRow = currentRow;
  currentRow++;
  sheet.getRange(currentRow, 1).setValue('Total Costos y Gastos').setFontWeight('bold');
  sheet.getRange(currentRow, 2).setFormula('=SUM(B' + expenseStartRow + ':B' + expenseEndRow + ')').setFontWeight('bold');
  var totalExpenseCell = 'B' + currentRow;
  currentRow += 2;

  // Net Income
  sheet.getRange(currentRow, 1).setValue('Utilidad Neta').setFontWeight('bold');
  sheet.getRange(currentRow, 2).setFormula('=' + totalIncomeCell + '-' + totalExpenseCell).setFontWeight('bold');

  // Formatting
  sheet.getRange(1, 2, currentRow, 1).setNumberFormat('$#,##0.00');
  sheet.getRange(totalIncomeCell).setBorder(true, null, null, null, null, null, null, SpreadsheetApp.BorderStyle.TOP_DOUBLE);
  sheet.getRange(totalExpenseCell).setBorder(true, null, null, null, null, null, null, SpreadsheetApp.BorderStyle.TOP_DOUBLE);
  sheet.getRange('B' + currentRow).setBorder(true, null, null, null, null, null, null, SpreadsheetApp.BorderStyle.TOP_DOUBLE);
}
