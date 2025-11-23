function setupPlanDeCuentas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Plan de Cuentas');
  if (!sheet) {
    sheet = ss.insertSheet('Plan de Cuentas');
  }

  var headers = ["Código", "Nombre de la Cuenta", "Tipo", "Subtipo"];
  var data = [
    ["1000", "Activo", "Activo", ""],
    ["1100", "Efectivo y Equivalentes", "Activo", "Activo Corriente"],
    ["1101", "Caja", "Activo", "Activo Corriente"],
    ["1102", "Bancos", "Activo", "Activo Corriente"],
    ["1200", "Cuentas por Cobrar", "Activo", "Activo Corriente"],
    ["1300", "Inventario", "Activo", "Activo Corriente"],
    ["1400", "Propiedad, Planta y Equipo", "Activo", "Activo no Corriente"],
    ["2000", "Pasivo", "Pasivo", ""],
    ["2100", "Cuentas por Pagar", "Pasivo", "Pasivo Corriente"],
    ["2200", "Préstamos Bancarios", "Pasivo", "Pasivo no Corriente"],
    ["3000", "Patrimonio", "Patrimonio", ""],
    ["3100", "Capital Social", "Patrimonio", ""],
    ["3200", "Resultados Acumulados", "Patrimonio", ""],
    ["4000", "Ingresos", "Ingresos", ""],
    ["4100", "Ventas", "Ingresos", "Ingresos Operacionales"],
    ["5000", "Egresos", "Egresos", ""],
    ["5100", "Costo de Ventas", "Egresos", "Egresos Operacionales"],
    ["5200", "Gastos Administrativos", "Egresos", "Egresos Operacionales"],
    ["5300", "Gastos de Ventas", "Egresos", "Egresos Operacionales"]
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(2, 1, data.length, data[0].length).setValues(data);

  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
}