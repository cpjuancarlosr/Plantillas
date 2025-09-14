/**
 * //== File: templates_oficina.gs
 * Generadores para las plantillas del módulo de Productividad de Oficina.
 */

/**
 * Función principal para crear todas las plantillas de Oficina.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - La hoja de cálculo activa.
 */
function createAllOfficeTemplates_(ss) {
  createAgendaTemplate_(ss);
  createMinutasTemplate_(ss);
  createDocsControlTemplate_(ss);
  createCuentasPagarTemplate_(ss);
  createCuentasCobrarTemplate_(ss);
  createBitacoraTemplate_(ss);
}

// --------------------------------------------------------------------------------
// 25. Agenda diaria/semana
// --------------------------------------------------------------------------------
function createAgendaTemplate_(ss) {
  const sn = 'C25_Agenda_Semanal';
  const s = getOrCreateSheet_(ss, sn, 15, 50);
  applyBaseSheetStyles_(s);
  s.getRange('A1').setValue('Agenda Semanal').setFontSize(14).setFontWeight('bold');
  const headers = ['Hora', 'Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes'];
  s.getRange('A3:F3').setValues([headers]);
  styleHeaderRow_(s, 'A3:F3');
  // Llenar horas
  const times = [];
  for (let i = 8; i < 19; i++) {
    times.push([`${i}:00`]);
  }
  s.getRange('A4:A14').setValues(times);
  logAction_(`Plantilla creada: ${sn}`);
}

// --------------------------------------------------------------------------------
// 26. Minutas y Acuerdos (RACI)
// --------------------------------------------------------------------------------
function createMinutasTemplate_(ss) {
  const sn = 'C26_Minutas_Acuerdos';
  const s = getOrCreateSheet_(ss, sn, 10, 100);
  applyBaseSheetStyles_(s);
  s.getRange('A1').setValue('Minuta de Reunión').setFontSize(14).setFontWeight('bold');
  s.getRange('A3').setValue('Tema:');
  s.getRange('A4').setValue('Fecha:');
  s.getRange('A5').setValue('Asistentes:');
  s.getRange('A7:F7').setValues([['#', 'Acuerdo / Tarea', 'Responsable (R)', 'Consultado (C)', 'Informado (I)', 'Fecha Compromiso']]);
  styleHeaderRow_(s, 'A7:F7');
  s.getRange('F8:F').setNumberFormat('dd/mm/yyyy');
  logAction_(`Plantilla creada: ${sn}`);
}

// --------------------------------------------------------------------------------
// 27. Control de Documentos
// --------------------------------------------------------------------------------
function createDocsControlTemplate_(ss) {
  const sn = 'C27_Control_Documentos';
  const s = getOrCreateSheet_(ss, sn, 10, 200);
  applyBaseSheetStyles_(s);
  s.getRange('A1:F1').setValues([['ID Doc', 'Nombre Documento', 'Categoría', 'Versión', 'Fecha Últ. Mod.', 'Enlace a Drive']]);
  styleHeaderRow_(s, 'A1:F1');
  s.getRange('E2:E').setNumberFormat('dd/mm/yyyy');
  // Se espera que el usuario pegue texto con hipervínculos en la columna F.
  logAction_(`Plantilla creada: ${sn}`);
}

// --------------------------------------------------------------------------------
// 28. Cuentas por Pagar
// --------------------------------------------------------------------------------
function createCuentasPagarTemplate_(ss) {
  const sn = 'C28_Cuentas_Por_Pagar';
  const s = getOrCreateSheet_(ss, sn, 15, 100);
  applyBaseSheetStyles_(s);
  const headers = ['ID Factura', 'Proveedor', 'Concepto', 'Fecha Factura', 'Fecha Vencimiento', 'Monto', 'Estado', 'Días Vencido'];
  s.getRange('A1:H1').setValues([headers]);
  styleHeaderRow_(s, 'A1:H1');
  s.getRange('H2:H').setFormulaR1C1('=IF(AND(RC[-1]<>"Pagado", RC[-3]<TODAY()), TODAY()-RC[-3], 0)');
  s.getRange('D:E').setNumberFormat('dd/mm/yyyy');
  s.getRange('F:F').setNumberFormat('$#,##0.00');

  // Tablero de antigüedad
  s.getRange('J1').setValue('Antigüedad de Saldos').setFontWeight('bold');
  s.getRange('J2:K2').setValues([['Rango', 'Monto']]);
  const rangos = [['Corriente'], ['1-30 días'], ['31-60 días'], ['60+ días']];
  s.getRange('J3:J6').setValues(rangos);
  s.getRange('K3').setFormula('=SUMIFS(F:F, H:H, "=0", G:G, "<>Pagado")');
  s.getRange('K4').setFormula('=SUMIFS(F:F, H:H, ">0", H:H, "<=30")');
  s.getRange('K5').setFormula('=SUMIFS(F:F, H:H, ">30", H:H, "<=60")');
  s.getRange('K6').setFormula('=SUMIFS(F:F, H:H, ">60")');
  s.getRange('K3:K6').setNumberFormat('$#,##0.00');

  logAction_(`Plantilla creada: ${sn}`);
}

// --------------------------------------------------------------------------------
// 29. Cuentas por Cobrar
// --------------------------------------------------------------------------------
function createCuentasCobrarTemplate_(ss) {
  const sn = 'C29_Cuentas_Por_Cobrar';
  const s = getOrCreateSheet_(ss, sn, 15, 100);
  applyBaseSheetStyles_(s);
  const headers = ['ID Factura', 'Cliente', 'Concepto', 'Fecha Factura', 'Fecha Vencimiento', 'Monto', 'Estado', 'Días Vencido'];
  s.getRange('A1:H1').setValues([headers]);
  styleHeaderRow_(s, 'A1:H1');
  s.getRange('H2:H').setFormulaR1C1('=IF(AND(RC[-1]<>"Cobrado", RC[-3]<TODAY()), TODAY()-RC[-3], 0)');
  s.getRange('D:E').setNumberFormat('dd/mm/yyyy');
  s.getRange('F:F').setNumberFormat('$#,##0.00');

  // Tablero de antigüedad (similar a Cuentas por Pagar)
  s.getRange('J1').setValue('Antigüedad de Saldos').setFontWeight('bold');
  s.getRange('J2:K2').setValues([['Rango', 'Monto']]);
  const rangos = [['Corriente'], ['1-30 días'], ['31-60 días'], ['60+ días']];
  s.getRange('J3:J6').setValues(rangos);
  s.getRange('K3').setFormula('=SUMIFS(F:F, H:H, "=0", G:G, "<>Cobrado")');
  s.getRange('K4').setFormula('=SUMIFS(F:F, H:H, ">0", H:H, "<=30")');
  s.getRange('K5').setFormula('=SUMIFS(F:F, H:H, ">30", H:H, "<=60")');
  s.getRange('K6').setFormula('=SUMIFS(F:F, H:H, ">60")');
  s.getRange('K3:K6').setNumberFormat('$#,##0.00');

  logAction_(`Plantilla creada: ${sn}`);
}

// --------------------------------------------------------------------------------
// 30. Bitácora de Actividades
// --------------------------------------------------------------------------------
function createBitacoraTemplate_(ss) {
  const sn = 'C30_Bitacora_Actividades';
  const s = getOrCreateSheet_(ss, sn, 10, 200);
  applyBaseSheetStyles_(s);
  const headers = ['Fecha', 'Hora Inicio', 'Hora Fin', 'Duración (min)', 'Actividad', 'Categoría', 'Nivel de Enfoque (1-5)'];
  s.getRange('A1:G1').setValues([headers]);
  styleHeaderRow_(s, 'A1:G1');
  s.getRange('D2:D').setFormulaR1C1('=IF(RC[-1]="","", (RC[-1]-RC[-2])*1440)');
  s.getRange('A:A').setNumberFormat('dd/mm/yyyy');
  s.getRange('B:C').setNumberFormat('hh:mm');

  // KPIs
  styleKpiCard_(s, 'I2:J3', 'Horas Registradas (Hoy)');
  s.getRange('I2').setFormula('=SUMIF(A:A, TODAY(), D:D)/60').setNumberFormat('0.0 "hrs"');

  logAction_(`Plantilla creada: ${sn}`);
}
