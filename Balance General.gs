function setupBalanceGeneral() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  // Get chart of accounts data
  var coaSheet = ss.getSheetByName('Plan de Cuentas');
  if (!coaSheet) {
    ui.alert('No se encuentra la hoja "Plan de Cuentas". Por favor, ejecute la configuración del sistema primero.');
    return;
  }
  var accounts = coaSheet.getRange(2, 1, coaSheet.getLastRow() - 1, 4).getValues();

  // Prepare the sheet
  var sheet = ss.getSheetByName('Balance General');
  if (!sheet) {
    sheet = ss.insertSheet('Balance General');
  }
  sheet.clear();
  sheet.setColumnWidths(1, 1, 250); // Asset Account Name
  sheet.setColumnWidths(2, 1, 120); // Asset Value
  sheet.setColumnWidths(3, 1, 50);  // Spacer
  sheet.setColumnWidths(4, 1, 250); // Liability/Equity Account Name
  sheet.setColumnWidths(5, 1, 120); // Liability/Equity Value

  var currentRow = 1;

  // Title
  sheet.getRange(currentRow, 1, 1, 5).merge().setValue('Balance General')
       .setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');
  currentRow += 2;

  // Headers
  sheet.getRange(currentRow, 1).setValue('Activos').setFontWeight('bold');
  sheet.getRange(currentRow, 4).setValue('Pasivos y Patrimonio').setFontWeight('bold');
  currentRow++;

  // --- Build the Report ---
  var assetRow = currentRow;
  var liabEqRow = currentRow;

  // Group accounts by type and subtype
  var assets = accounts.filter(function(acc) { return acc[2] === 'Activo' && acc[3]; });
  var liabilities = accounts.filter(function(acc) { return acc[2] === 'Pasivo' && acc[3]; });
  var equity = accounts.filter(function(acc) { return acc[2] === 'Patrimonio' && acc[3]; });

  // Function to process a section of accounts
  function processSection(startRow, accountList, isAsset) {
    var col = isAsset ? 1 : 4;
    var valCol = isAsset ? 2 : 5;
    var formulaSign = isAsset ? "F-G" : "G-F";
    var runningTotalFormula = "";

    var subtypes = [...new Set(accountList.map(item => item[3]))]; // Get unique subtypes

    subtypes.forEach(function(subtype) {
      sheet.getRange(startRow, col).setValue(subtype).setFontWeight('bold');
      startRow++;
      var subtotalFormula = "";

      accountList.filter(acc => acc[3] === subtype).forEach(function(acc) {
        sheet.getRange(startRow, col).setValue("  " + acc[1]); // Indent account name
        sheet.getRange(startRow, valCol).setFormula("=MAX(0, SUMIF('Libro Diario'!D:D, " + acc[0] + ", 'Libro Diario'!" + formulaSign.split('-')[0] + ":" + formulaSign.split('-')[0] + ") - SUMIF('Libro Diario'!D:D, " + acc[0] + ", 'Libro Diario'!" + formulaSign.split('-')[1] + ":" + formulaSign.split('-')[1] + "))");
        subtotalFormula += (subtotalFormula ? "+" : "") + sheet.getRange(startRow, valCol).getA1Notation();
        startRow++;
      });

      sheet.getRange(startRow, col).setValue("Total " + subtype).setFontWeight('bold');
      sheet.getRange(startRow, valCol).setFormula("=" + subtotalFormula).setFontWeight('bold');
      runningTotalFormula += (runningTotalFormula ? "+" : "") + sheet.getRange(startRow, valCol).getA1Notation();
      startRow++;
      startRow++; // Spacer row
    });

    return { nextRow: startRow, totalFormula: runningTotalFormula };
  }

  // Process Assets
  var assetResult = processSection(assetRow, assets, true);
  assetRow = assetResult.nextRow;
  sheet.getRange(assetRow, 1).setValue("Total Activos").setFontWeight('bold');
  var totalAssetsCell = sheet.getRange(assetRow, 2);
  totalAssetsCell.setFormula("=" + assetResult.totalFormula).setFontWeight('bold');

  // Process Liabilities
  var liabResult = processSection(liabEqRow, liabilities, false);
  liabEqRow = liabResult.nextRow;

  // Process Equity
  var equityResult = processSection(liabEqRow, equity, false);
  liabEqRow = equityResult.nextRow;

  // Add Net Income to Equity
  sheet.getRange(liabEqRow, 4).setValue("Utilidad del Ejercicio");
  var netIncomeFormula = "='Estado de Resultados'!B" + ss.getSheetByName('Estado de Resultados').getLastRow();
  sheet.getRange(liabEqRow, 5).setFormula(netIncomeFormula);
  var netIncomeCell = sheet.getRange(liabEqRow, 5).getA1Notation();
  liabEqRow++;

  sheet.getRange(liabEqRow, 4).setValue("Total Patrimonio").setFontWeight('bold');
  var totalEquityCell = sheet.getRange(liabEqRow, 5);
  totalEquityCell.setFormula("=" + equityResult.totalFormula + "+" + netIncomeCell).setFontWeight('bold');
  liabEqRow += 2;

  // Total Liabilities and Equity
  sheet.getRange(liabEqRow, 4).setValue("Total Pasivos y Patrimonio").setFontWeight('bold');
  var totalLiabEqCell = sheet.getRange(liabEqRow, 5);
  totalLiabEqCell.setFormula("=" + liabResult.totalFormula + "+" + totalEquityCell.getA1Notation()).setFontWeight('bold');

  // Final Formatting
  sheet.getRange("B:B, E:E").setNumberFormat('$#,##0.00');

  // Accounting Equation Check
  var checkRow = Math.max(assetRow, liabEqRow) + 2;
  sheet.getRange(checkRow, 1).setValue("Verificación (Activo = Pasivo + Patrimonio)").setFontWeight('bold');
  sheet.getRange(checkRow, 2).setFormula("=" + totalAssetsCell.getA1Notation() + "-" + totalLiabEqCell.getA1Notation()).setNumberFormat('$#,##0.00;(#,##0.00);"-"');

  // Conditional formatting for the check cell
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberEqualTo(0)
    .setBackground("#d9ead3") // green
    .setRanges([sheet.getRange(checkRow, 2)])
    .build();
  var rules = sheet.getConditionalFormatRules();
  rules.push(rule);
  sheet.setConditionalFormatRules(rules);
}