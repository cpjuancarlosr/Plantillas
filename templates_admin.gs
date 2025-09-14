/**
 * //== File: templates_admin.gs
 * Generadores para las plantillas del módulo de Administración y Operación.
 */

/**
 * Función principal para crear todas las plantillas de Administración.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - La hoja de cálculo activa.
 */
function createAllAdminTemplates_(ss) {
  createProjectKanbanTemplate_(ss);
  createCrmTemplate_(ss);
  createInventoryTemplate_(ss);
  createPurchasingTemplate_(ss);
  createSalesTemplate_(ss);
  createInvoicingTemplate_(ss);
  createPayrollTemplate_(ss);
  createBudgetTemplate_(ss);
  createCashflowTemplate_(ss);
  createOkrTemplate_(ss);
}

// --------------------------------------------------------------------------------
// 1. Control de Proyectos y Tareas (Kanban) - COMPLETO
// --------------------------------------------------------------------------------
function createProjectKanbanTemplate_(ss) {
  const sheetName = 'A01_Proyectos_Kanban';
  const sheet = getOrCreateSheet_(ss, sheetName, 20, 100);
  applyBaseSheetStyles_(sheet);
  const headers = ['ID Tarea', 'Tarea', 'Proyecto', 'Responsable', 'Estado', 'Prioridad', 'Fecha Inicio', 'Fecha Fin', 'Duración (días)', '% Avance'];
  sheet.getRange('A1:J1').setValues([headers]);
  styleHeaderRow_(sheet, 'A1:J1');
  sheet.setFrozenColumns(2);
  sheet.getRange('I2:I').setFormulaR1C1('=IF(RC[-1]="","",RC[-1]-RC[-2])');
  sheet.getRange('A:A').setNumberFormat('@');
  sheet.getRange('G:H').setNumberFormat('dd/mm/yyyy');
  sheet.getRange('J:J').setNumberFormat('0%');
  sheet.getRange('L1').setValue('Tablero Kanban').setFontWeight('bold').setFontSize(14);
  sheet.getRange('L2:N2').setValues([['Pendiente', 'En Progreso', 'Completado']]).setFontWeight('bold').setBackground(THEME_COLORS.HEADER);
  sheet.getRange('L3').setFormula('=FILTER(B$2:B, E$2:E="Pendiente")');
  sheet.getRange('M3').setFormula('=FILTER(B$2:B, E$2:E="En Progreso")');
  sheet.getRange('N3').setFormula('=FILTER(B$2:B, E$2:E="Completado")');
  styleKpiCard_(sheet, 'P2:Q3', 'Proyectos Activos');
  sheet.getRange('P2').setFormula('=COUNTA(UNIQUE(FILTER(C$2:C, C$2:C<>"")))');
  styleKpiCard_(sheet, 'R2:S3', '% Avance Promedio');
  sheet.getRange('R2').setFormula('=AVERAGE(J$2:J)').setNumberFormat('0.0%');
  applyKanbanValidations_(sheet);
  logAction_(`Plantilla creada: ${sheetName}`);
}

// --------------------------------------------------------------------------------
// 2. CRM Básico - COMPLETO
// --------------------------------------------------------------------------------
function createCrmTemplate_(ss) {
  const sheetName = 'A02_CRM_Pipeline';
  const sheet = getOrCreateSheet_(ss, sheetName, 20, 100);
  applyBaseSheetStyles_(sheet);
  const headers = ['ID Lead', 'Contacto', 'Empresa', 'Etapa', 'Valor ($)', 'Probabilidad (%)', 'Valor Ponderado', 'Fecha Contacto', 'Próximo Seguimiento'];
  sheet.getRange('A1:I1').setValues([headers]);
  styleHeaderRow_(sheet, 'A1:I1');
  sheet.getRange('G2:G').setFormulaR1C1('=IF(RC[-2]="","",RC[-2]*RC[-1])');
  sheet.getRange('E:G').setNumberFormat('$#,##0.00');
  sheet.getRange('F:F').setNumberFormat('0%');
  sheet.getRange('H:I').setNumberFormat('dd/mm/yyyy');
  sheet.getRange('K1').setValue('Pipeline de Ventas').setFontWeight('bold').setFontSize(14);
  const stages = ['Nuevo', 'Contactado', 'Propuesta', 'Negociación', 'Ganado', 'Perdido'];
  sheet.getRange('K2:K7').setValues(stages.map(s => [s]));
  sheet.getRange('L2:L7').setFormulaR1C1('=SUMIF(D:D, R[0]C[-1], E:E)');
  sheet.getRange('L2:L7').setNumberFormat('$#,##0.00');
  styleKpiCard_(sheet, 'N2:O3', 'Valor Total Pipeline');
  sheet.getRange('N2').setFormula('=SUM(L2:L5)');
  styleKpiCard_(sheet, 'P2:Q3', 'Tasa de Cierre');
  sheet.getRange('P2').setFormula('=IFERROR(SUMIF(D:D,"Ganado",E:E) / (SUMIF(D:D,"Ganado",E:E) + SUMIF(D:D,"Perdido",E:E)), 0)').setNumberFormat('0.0%');
  applyCrmValidations_(sheet);
  logAction_(`Plantilla creada: ${sheetName}`);
}

// --------------------------------------------------------------------------------
// 3. Control de Inventarios - COMPLETO
// --------------------------------------------------------------------------------
function createInventoryTemplate_(ss) {
  const sheetName = 'A03_Inventarios';
  const sheet = getOrCreateSheet_(ss, sheetName, 20, 200);
  applyBaseSheetStyles_(sheet);
  sheet.getRange('A1:G1').setValues([['Fecha', 'Tipo Mov.', 'ID Producto', 'Producto', 'Cantidad', 'Costo Unitario', 'Valor Total']]).setFontWeight('bold').setBackground(THEME_COLORS.HEADER);
  sheet.getRange('G2:G').setFormulaR1C1('=IF(RC[-2]="","", RC[-2]*RC[-1])');
  sheet.getRange('I1:M1').setValues([['ID Producto', 'Producto', 'Stock Actual', 'Costo Promedio', 'Alerta Mínimo']]).setFontWeight('bold').setBackground(THEME_COLORS.HEADER);
  sheet.getRange('K2:K').setFormulaR1C1('=IF(RC[-2]="","", SUMIFS(E:E, C:C, RC[-2], B:B, "Entrada") - SUMIFS(E:E, C:C, RC[-2], B:B, "Salida"))');
  sheet.getRange('L2:L').setFormulaR1C1('=IFERROR(AVERAGEIFS(F:F, C:C, RC[-3], B:B, "Entrada"),0)');
  sheet.getRange('A:A').setNumberFormat('dd/mm/yyyy');
  sheet.getRange('F:G,L:L').setNumberFormat('$#,##0.00');
  const rule = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(K2<>"", M2<>"", K2<=M2)').setBackground("#f4cccc").setRanges([sheet.getRange('K2:K')]).build();
  sheet.getConditionalFormatRules().push(rule);
  styleKpiCard_(sheet, 'O2:P3', 'Valor Total Inventario');
  sheet.getRange('O2').setFormula('=SUMPRODUCT(K2:K,L2:L)');
  styleKpiCard_(sheet, 'Q2:R3', 'Productos Bajo Mínimo');
  sheet.getRange('Q2').setFormula('=COUNTIFS(K2:K,">0",K2:K,"<="&M2:M)');
  applyInventoryValidations_(sheet);
  logAction_(`Plantilla creada: ${sheetName}`);
}

// --------------------------------------------------------------------------------
// 4. Compras y Proveedores - AHORA COMPLETO
// --------------------------------------------------------------------------------
function createPurchasingTemplate_(ss) {
    const sn = 'A04_Compras_Proveedores';
    const s = getOrCreateSheet_(ss, sn, 20, 100);
    applyBaseSheetStyles_(s);
    const headers = ['ID Orden', 'Proveedor', 'Fecha Pedido', 'Fecha Entrega Estimada', 'Producto', 'Cantidad', 'Precio Unitario', 'Total Orden', 'Estado'];
    s.getRange('A1:I1').setValues([headers]);
    styleHeaderRow_(s, 'A1:I1');
    s.getRange('H2:H').setFormulaR1C1('=IF(RC[-2]="","",RC[-2]*RC[-1])');
    s.getRange('C:D').setNumberFormat('dd/mm/yyyy');
    s.getRange('G:H').setNumberFormat('$#,##0.00');
    const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['Pedido', 'En Tránsito', 'Recibido', 'Cancelado']).build();
    s.getRange('I2:I').setDataValidation(statusRule);
    styleKpiCard_(s, 'K2:L3', 'Órdenes Abiertas');
    s.getRange('K2').setFormula('=COUNTIF(I2:I, "<>Recibido")');
    styleKpiCard_(s, 'M2:N3', 'Valor Compras (Mes)');
    s.getRange('M2').setFormula('=SUMIFS(H2:H, C2:C, ">="&EOMONTH(TODAY(),-1)+1, C2:C, "<="&EOMONTH(TODAY(),0))');
    logAction_(`Plantilla creada: ${sn}`);
}

// --------------------------------------------------------------------------------
// 5. Ventas y Clientes - AHORA COMPLETO
// --------------------------------------------------------------------------------
function createSalesTemplate_(ss) {
    const sn = 'A05_Ventas_Clientes';
    const s = getOrCreateSheet_(ss, sn, 20, 200);
    applyBaseSheetStyles_(s);
    const headers = ['ID Venta', 'Cliente', 'Fecha Venta', 'Producto SKU', 'Cantidad', 'Precio Venta Unit.', 'Costo Unit.', 'Ingreso Total', 'Margen Venta', 'Margen %'];
    s.getRange('A1:J1').setValues([headers]);
    styleHeaderRow_(s, 'A1:J1');
    s.getRange('H2:H').setFormulaR1C1('=IF(RC[-3]="","",RC[-3]*RC[-2])');
    s.getRange('I2:I').setFormulaR1C1('=IF(RC[-1]="","",RC[-1]-(RC[-4]*RC[-2]))');
    s.getRange('J2:J').setFormulaR1C1('=IFERROR(RC[-1]/RC[-2],0)');
    s.getRange('C:C').setNumberFormat('dd/mm/yyyy');
    s.getRange('F:I').setNumberFormat('$#,##0.00');
    s.getRange('J:J').setNumberFormat('0.0%');
    styleKpiCard_(s, 'L2:M3', 'Ingresos (Mes)');
    s.getRange('L2').setFormula('=SUMIFS(H2:H, C2:C, ">="&EOMONTH(TODAY(),-1)+1, C2:C, "<="&EOMONTH(TODAY(),0))');
    styleKpiCard_(s, 'N2:O3', 'Margen Promedio');
    s.getRange('N2').setFormula('=AVERAGE(J2:J)');
    logAction_(`Plantilla creada: ${sn}`);
}

// --------------------------------------------------------------------------------
// 6. Facturación SAT - AHORA COMPLETO
// --------------------------------------------------------------------------------
function createInvoicingTemplate_(ss) {
    const sn = 'A06_Facturacion_SAT';
    const s = getOrCreateSheet_(ss, sn, 20, 100);
    applyBaseSheetStyles_(s);
    const headers = ['Folio Fiscal (UUID)', 'Fecha Emisión', 'RFC Emisor', 'RFC Receptor', 'Subtotal', 'IVA (16%)', 'Total', 'Estado SAT', 'Tipo (I/E)'];
    s.getRange('A1:I1').setValues([headers]);
    styleHeaderRow_(s, 'A1:I1');
    s.getRange('B:B').setNumberFormat('dd/mm/yyyy');
    s.getRange('E:G').setNumberFormat('$#,##0.00');
    s.getRange('F2:F').setFormulaR1C1('=RC[-1]*VLOOKUP("IVA_Tasa_General",_CONFIG!A:B,2,0)');
    s.getRange('G2:G').setFormulaR1C1('=RC[-2]+RC[-1]');
    const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(['Vigente', 'Cancelado']).build();
    s.getRange('H2:H').setDataValidation(statusRule);
    styleKpiCard_(s, 'K2:L3', 'Facturado (Mes)');
    s.getRange('K2').setFormula('=SUMIFS(G2:G, H2:H, "Vigente", B2:B, ">="&EOMONTH(TODAY(),-1)+1, B2:B, "<="&EOMONTH(TODAY(),0))');
    styleKpiCard_(s, 'M2:N3', 'IVA por Pagar (Mes)');
    s.getRange('M2').setFormula('=SUMIFS(F2:F, H2:H, "Vigente", B2:B, ">="&EOMONTH(TODAY(),-1)+1, B2:B, "<="&EOMONTH(TODAY(),0))');
    logAction_(`Plantilla creada: ${sn}`);
}

// --------------------------------------------------------------------------------
// 7. Nómina Simple - AHORA COMPLETO
// --------------------------------------------------------------------------------
function createPayrollTemplate_(ss) {
    const sn = 'A07_Nomina_Simple';
    const s = getOrCreateSheet_(ss, sn, 20, 100);
    applyBaseSheetStyles_(s);
    const headers = ['ID Empleado', 'Nombre', 'Puesto', 'Sueldo Bruto Mensual', 'Percepciones Adic.', 'Deducciones (IMSS, ISR)', 'Sueldo Neto', 'Costo Total Empresa'];
    s.getRange('A1:H1').setValues([headers]);
    styleHeaderRow_(s, 'A1:H1');
    s.getRange('D:H').setNumberFormat('$#,##0.00');
    s.getRange('G2:G').setFormulaR1C1('=RC[-3]+RC[-2]-RC[-1]');
    s.getRange('H2:H').setFormulaR1C1('=RC[-4]*1.35'); // Estimación 35% carga social
    styleKpiCard_(s, 'J2:K3', 'Costo Total Nómina');
    s.getRange('J2').setFormula('=SUM(H2:H)');
    styleKpiCard_(s, 'L2:M3', 'Nómina Neta');
    s.getRange('L2').setFormula('=SUM(G2:G)');
    logAction_(`Plantilla creada: ${sn}`);
}

// --------------------------------------------------------------------------------
// 8. Presupuesto Maestro - AHORA COMPLETO
// --------------------------------------------------------------------------------
function createBudgetTemplate_(ss) {
    const sn = 'A08_Presupuesto_Maestro';
    const s = getOrCreateSheet_(ss, sn, 30, 100);
    applyBaseSheetStyles_(s);
    s.getRange('A1').setValue('Presupuesto Maestro (Anual)').setFontSize(14).setFontWeight('bold');
    const headers = ['Concepto', 'Categoría', 'Tipo', 'Ene Ppto', 'Ene Real', 'Feb Ppto', 'Feb Real', '...etc', 'Total Ppto Anual', 'Total Real Anual', 'Variación Anual'];
    s.getRange('A3').setValue('Estructura Simplificada. Llenar meses restantes.');
    s.getRange('A4:M4').setValues([['Concepto', 'Categoría', 'Tipo', 'Ene Ppto', 'Ene Real', 'Var Ene', 'Feb Ppto', 'Feb Real', 'Var Feb', 'Total Ppto', 'Total Real', 'Var Total', '% Var']]);
    styleHeaderRow_(s, 'A4:M4');
    s.getRange('F5:F').setFormulaR1C1('=RC[-1]-RC[-2]'); // Var Ene
    s.getRange('I5:I').setFormulaR1C1('=RC[-1]-RC[-2]'); // Var Feb
    s.getRange('J5:J').setFormulaR1C1('=SUM(RC[-6],RC[-3])'); // Total Ppto
    s.getRange('K5:K').setFormulaR1C1('=SUM(RC[-6],RC[-3])'); // Total Real
    s.getRange('L5:L').setFormulaR1C1('=RC[-1]-RC[-2]'); // Var Total
    s.getRange('M5:M').setFormulaR1C1('=IFERROR(RC[-1]/RC[-3],0)'); // % Var
    s.getRange('D:M').setNumberFormat('$#,##0.00');
    s.getRange('M:M').setNumberFormat('0%');
    styleKpiCard_(s, 'O3:P4', 'Desviación Total');
    s.getRange('O3').setFormula('=SUM(L5:L)');
    logAction_(`Plantilla creada: ${sn}`);
}

// --------------------------------------------------------------------------------
// 9. Flujo de Efectivo Operativo - AHORA COMPLETO
// --------------------------------------------------------------------------------
function createCashflowTemplate_(ss) {
    const sn = 'A09_Flujo_Efectivo_Op';
    const s = getOrCreateSheet_(ss, sn, 20, 200);
    applyBaseSheetStyles_(s);
    s.getRange('A1').setValue('Flujo de Efectivo (Método Directo)').setFontSize(14).setFontWeight('bold');
    s.getRange('A3').setValue('Saldo Inicial de Efectivo:').setFontWeight('bold');
    s.getRange('D3').setNumberFormat('$#,##0.00');
    const headers = ['Fecha', 'Concepto', 'Categoría', 'Entradas', 'Salidas', 'Saldo'];
    s.getRange('A5:F5').setValues([headers]);
    styleHeaderRow_(s, 'A5:F5');
    s.getRange('F6').setFormulaR1C1('=R3C4+RC[-2]-RC[-1]');
    s.getRange('F7:F').setFormulaR1C1('=R[-1]C+RC[-2]-RC[-1]');
    s.getRange('A:A').setNumberFormat('dd/mm/yyyy');
    s.getRange('D:F').setNumberFormat('$#,##0.00');
    styleKpiCard_(s, 'H3:I4', 'Saldo Final');
    s.getRange('H3').setFormula('=INDEX(F:F, COUNTA(F:F))');
    styleKpiCard_(s, 'J3:K4', 'Flujo Neto (Periodo)');
    s.getRange('J3').setFormula('=SUM(D:D)-SUM(E:E)');
    logAction_(`Plantilla creada: ${sn}`);
}

// --------------------------------------------------------------------------------
// 10. OKR trimestral - AHORA COMPLETO
// --------------------------------------------------------------------------------
function createOkrTemplate_(ss) {
    const sn = 'A10_OKR_Trimestral';
    const s = getOrCreateSheet_(ss, sn, 20, 100);
    applyBaseSheetStyles_(s);
    s.getRange('A1').setValue('Seguimiento de OKRs (Objectives and Key Results)').setFontSize(14).setFontWeight('bold');
    const headers = ['Objetivo', 'Resultado Clave (KR)', 'Responsable', 'Métrica Inicial', 'Métrica Meta', 'Valor Actual', '% Logro', 'Estado'];
    s.getRange('A3:H3').setValues([headers]);
    styleHeaderRow_(s, 'A3:H3');
    s.getRange('G4:G').setFormulaR1C1('=IFERROR((RC[-1]-RC[-3])/(RC[-2]-RC[-3]),0)');
    s.getRange('H4:H').setFormulaR1C1('=IF(RC[-1]>=1, "Logrado", IF(RC[-1]>=0.7, "En Camino", "En Riesgo"))');
    s.getRange('G:G').setNumberFormat('0%');
    const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains("Logrado").setBackground("#d9ead3")
        .whenTextContains("En Camino").setBackground("#fff2cc")
        .whenTextContains("En Riesgo").setBackground("#f4cccc")
        .setRanges([s.getRange('H4:H')]).build();
    s.getConditionalFormatRules().push(rule);
    styleKpiCard_(s, 'J3:K4', 'Logro Promedio OKRs');
    s.getRange('J3').setFormula('=AVERAGE(G4:G)');
    logAction_(`Plantilla creada: ${sn}`);
}
