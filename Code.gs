function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Contabilidad')
    .addItem('Registrar Asiento Contable', 'showTransactionForm')
    .addSeparator()
    .addItem('Configurar Sistema', 'setupSystem')
    .addToUi();
}

function setupSystem() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Confirmación de Configuración',
    'Este proceso configurará las hojas de cálculo necesarias para el sistema de contabilidad. Cualquier dato existente en las hojas "Plan de Cuentas", "Libro Diario", "Estado de Resultados" y "Balance General" será eliminado. ¿Desea continuar?',
    ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    try {
      setupPlanDeCuentas();
      setupLibroDiario();
      setupEstadoDeResultados();
      setupBalanceGeneral();
      ui.alert('El sistema de contabilidad ha sido configurado exitosamente.');
    } catch (e) {
      ui.alert('Error durante la configuración: ' + e.message);
    }
  }
}

function showTransactionForm() {
  var html = HtmlService.createHtmlOutputFromFile('TransactionForm')
    .setWidth(400)
    .setHeight(500);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getChartOfAccounts() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Plan de Cuentas');
  return sheet.getDataRange().getValues();
}

function recordTransaction(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Libro Diario');
    var lastRow = sheet.getLastRow();
    var nextId = 1;
    if (lastRow > 1) {
      var range = sheet.getRange(2, 1, lastRow - 1, 1);
      var values = range.getValues();
      var maxId = values.reduce(function(max, row) {
        return Math.max(max, row[0]);
      }, 0);
      nextId = maxId + 1;
    }

    var date = new Date(data.date);
    var description = data.description;

    for (var i = 0; i < data.entries.length; i++) {
      var entry = data.entries[i];
      var accountInfo = entry.account.split(' - ');
      var accountCode = accountInfo[0];
      var accountName = accountInfo[1];
      var debit = parseFloat(entry.debit) || 0;
      var credit = parseFloat(entry.credit) || 0;

      sheet.appendRow([nextId, date, description, accountCode, accountName, debit, credit]);
    }

    return "Transacción registrada con éxito.";
  } catch (e) {
    return "Error al registrar la transacción: " + e.message;
  }
}