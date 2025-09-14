/**
 * //== File: templates_conta_fiscal.gs
 * Generadores para las plantillas del módulo de Contabilidad y Fiscal (SAT-MX).
 */

function createAllContaFiscalTemplates_(ss) {
  createCatalogoCuentasTemplate_(ss);
  createPolizasTemplate_(ss);
  createBalanzaComprobacionTemplate_(ss);
  createEstadoResultadosTemplate_(ss);
  createBalanceGeneralTemplate_(ss);
  createActivosFijosTemplate_(ss);
  createFlujoEfectivoIndirectoTemplate_(ss); // Depends on ActivosFijos
  createConciliacionBancariaTemplate_(ss);
  createIvaTemplate_(ss);
  createIsrPmTemplate_(ss);
  createPtuTemplate_(ss);
  createRetencionesTemplate_(ss);
  createDeclaracionesProvisionalesTemplate_(ss);
  createTableroFiscalTemplate_(ss);
}

// B11 a B15 (ya estaban mayormente completos, se refinan)
function createCatalogoCuentasTemplate_(ss) {
  const sn = 'B11_Catalogo_Cuentas';
  const s = getOrCreateSheet_(ss, sn, 10, 500);
  applyBaseSheetStyles_(s);
  const headers = ['Código Agrupador SAT', 'Nombre Cuenta SAT', 'Nivel', 'Código Interno', 'Nombre Cuenta Interna', 'Tipo (D/A)', 'Naturaleza'];
  s.getRange('A1:G1').setValues([headers]);
  styleHeaderRow_(s, 'A1:G1');
  s.getRange('A1').setNote('Catálogo de cuentas base. El usuario debe expandir esto con su propio catálogo contable.');
  logAction_(`Plantilla creada: ${sn}`);
}

function createPolizasTemplate_(ss) {
  const sn = 'B12_Polizas_Diario';
  const s = getOrCreateSheet_(ss, sn, 10, 1000);
  applyBaseSheetStyles_(s);
  const headers = ['Fecha', 'Tipo Póliza', 'Folio', 'Cuenta', 'Concepto', 'Debe', 'Haber'];
  s.getRange('A1:G1').setValues([headers]);
  styleHeaderRow_(s, 'A1:G1');
  s.getRange('A:A').setNumberFormat('dd/mm/yyyy');
  s.getRange('F:G').setNumberFormat('$#,##0.00');
  styleKpiCard_(s, 'I2:J3', 'Descuadre Actual');
  s.getRange('I2').setFormula('=SUM(F:F)-SUM(G:G)');
  logAction_(`Plantilla creada: ${sn}`);
}

function createBalanzaComprobacionTemplate_(ss) {
  const sn = 'B13_Balanza_Comprobacion';
  const s = getOrCreateSheet_(ss, sn, 10, 500);
  applyBaseSheetStyles_(s);
  const headers = ['Cuenta', 'Nombre', 'Saldo Inicial', 'Debe', 'Haber', 'Saldo Final'];
  s.getRange('A1:F1').setValues([headers]);
  styleHeaderRow_(s, 'A1:F1');
  s.getRange('A2').setFormula(`=SORT(UNIQUE('${'B11_Catalogo_Cuentas'}!D2:D))` );
  s.getRange('B2:B').setFormulaR1C1(`=IF(RC[-1]="","",VLOOKUP(RC[-1],'${'B11_Catalogo_Cuentas'}!D:E,2,FALSE))`);
  s.getRange('D2:D').setFormulaR1C1(`=IF(RC[-3]="","",SUMIF('${'B12_Polizas_Diario'}!D:D,RC[-3],'${'B12_Polizas_Diario'}!F:F))`);
  s.getRange('E2:E').setFormulaR1C1(`=IF(RC[-4]="","",SUMIF('${'B12_Polizas_Diario'}!D:D,RC[-4],'${'B12_Polizas_Diario'}!G:G))`);
  s.getRange('F2:F').setFormulaR1C1(`=RC[-3]+RC[-2]-RC[-1]`);
  s.getRange('C:F').setNumberFormat('$#,##0.00');
  logAction_(`Plantilla creada: ${sn}`);
}

function createEstadoResultadosTemplate_(ss) {
  const sn = 'B14_Estado_Resultados';
  const s = getOrCreateSheet_(ss, sn, 5, 50);
  applyBaseSheetStyles_(s);
  s.getRange('A1').setValue('Estado de Resultados').setFontWeight('bold').setFontSize(14);
  const concepts = [
    ['Ventas Netas', '=SUMIFS(B13_Balanza_Comprobacion!F:F, B13_Balanza_Comprobacion!A:A, ">=4000", B13_Balanza_Comprobacion!A:A, "<5000")*-1'],
    ['Costo de Ventas', '=SUMIFS(B13_Balanza_Comprobacion!F:F, B13_Balanza_Comprobacion!A:A, ">=5000", B13_Balanza_Comprobacion!A:A, "<6000")'],
    ['Utilidad Bruta', '=B2-B3'],
    ['Gastos de Operación', '=SUMIFS(B13_Balanza_Comprobacion!F:F, B13_Balanza_Comprobacion!A:A, ">=6000", B13_Balanza_Comprobacion!A:A, "<7000")'],
    ['Utilidad Operativa (EBITDA aprox.)', '=B4-B5'],
    ['Utilidad antes de Impuestos', '=B6'],
    ['ISR (Tasa: 30%)', '=B7*ISR_Tasa_PM'],
    ['PTU (Tasa: 10%)', '=IF(B7>0, B7*PTU_Porcentaje, 0)'],
    ['Utilidad Neta', '=B7-B8-B9']
  ];
  s.getRange('A2:B10').setValues(concepts);
  s.getRange('B2:B10').setNumberFormat('$#,##0.00');
  s.getRange('A2:A10').setFontWeight('bold');
  logAction_(`Plantilla creada: ${sn}`);
}

function createBalanceGeneralTemplate_(ss) {
    const sn = 'B15_Balance_General';
    const s = getOrCreateSheet_(ss, sn, 8, 50);
    applyBaseSheetStyles_(s);
    s.getRange('A1').setValue('Balance General').setFontWeight('bold').setFontSize(14);
    s.getRange('A2:B2').merge().setValue('ACTIVO').setFontWeight('bold').setHorizontalAlignment('center');
    s.getRange('A3').setValue('Activo Circulante');
    s.getRange('A4').setValue('Efectivo y Equivalentes').setIndent(1);
    s.getRange('B4').setFormula('=SUMIFS(B13_Balanza_Comprobacion!F:F, B13_Balanza_Comprobacion!A:A, ">=1010", B13_Balanza_Comprobacion!A:A, "<1030")');
    s.getRange('A5').setValue('Cuentas por Cobrar').setIndent(1);
    s.getRange('B5').setFormula('=SUMIFS(B13_Balanza_Comprobacion!F:F, B13_Balanza_Comprobacion!A:A, ">=1050", B13_Balanza_Comprobacion!A:A, "<1060")');
    s.getRange('A6').setValue('Inventarios').setIndent(1);
    s.getRange('B6').setFormula('=SUMIFS(B13_Balanza_Comprobacion!F:F, B13_Balanza_Comprobacion!A:A, ">=1070", B13_Balanza_Comprobacion!A:A, "<1080")');
    s.getRange('A7').setValue('Total Activo Circulante').setFontWeight('bold');
    s.getRange('B7').setFormula('=SUM(B4:B6)').setFontWeight('bold');
    s.getRange('D2:E2').merge().setValue('PASIVO Y CAPITAL').setFontWeight('bold').setHorizontalAlignment('center');
    s.getRange('D3').setValue('Pasivo a Corto Plazo');
    s.getRange('D4').setValue('Proveedores').setIndent(1);
    s.getRange('E4').setFormula('=SUMIFS(B13_Balanza_Comprobacion!F:F, B13_Balanza_Comprobacion!A:A, ">=2010", B13_Balanza_Comprobacion!A:A, "<2020")*-1');
    s.getRange('D5').setValue('Impuestos por Pagar').setIndent(1);
    s.getRange('E5').setFormula('=SUMIFS(B13_Balanza_Comprobacion!F:F, B13_Balanza_Comprobacion!A:A, ">=2070", B13_Balanza_Comprobacion!A:A, "<2080")*-1');
    s.getRange('D6').setValue('Total Pasivo a Corto Plazo').setFontWeight('bold');
    s.getRange('E6').setFormula('=SUM(E4:E5)').setFontWeight('bold');
    s.getRange('D8').setValue('Capital Contable');
    s.getRange('D9').setValue('Capital Social').setIndent(1);
    s.getRange('E9').setFormula('=SUMIFS(B13_Balanza_Comprobacion!F:F, B13_Balanza_Comprobacion!A:A, ">=3010", B13_Balanza_Comprobacion!A:A, "<3020")*-1');
    s.getRange('D10').setValue('Utilidad Neta del Ejercicio').setIndent(1);
    s.getRange('E10').setFormula(`='B14_Estado_Resultados'!B10`);
    s.getRange('D11').setValue('Total Capital Contable').setFontWeight('bold');
    s.getRange('E11').setFormula('=SUM(E9:E10)').setFontWeight('bold');
    s.getRange('D13').setValue('TOTAL PASIVO + CAPITAL').setFontWeight('bold');
    s.getRange('E13').setFormula('=E6+E11').setFontWeight('bold');
    styleKpiCard_(s, 'G3:H4', 'Razón Circulante');
    s.getRange('G3').setFormula('=IFERROR(B7/E6,0)');
    logAction_(`Plantilla creada: ${sn}`);
}

// B16 a B24 - AHORA COMPLETOS
function createFlujoEfectivoIndirectoTemplate_(ss) {
    const sn = 'B16_Flujo_Efectivo_Ind';
    const s = getOrCreateSheet_(ss, sn, 5, 50);
    applyBaseSheetStyles_(s);
    s.getRange('A1').setValue('Flujo de Efectivo (Método Indirecto)').setFontWeight('bold').setFontSize(14);
    const concepts = [
      ['Utilidad Neta antes de Impuestos', `=B14_Estado_Resultados!B7`],
      ['Partidas que no requieren efectivo:'],
      ['+ Depreciación', `=B21_Activos_Fijos_Dep!F10`], // Link a la hoja de activos
      ['- Aumento en Cuentas por Cobrar', '0'],
      ['+ Aumento en Cuentas por Pagar', '0'],
      ['= Flujo de Efectivo de Operaciones', '=SUM(B2,B4,B5,B6)'],
    ];
    s.getRange('A2:B7').setValues(concepts);
    s.getRange('A3,A4,A5').setIndent(1);
    s.getRange('B2:B7').setNumberFormat('$#,##0.00');
    logAction_(`Plantilla creada: ${sn}`);
}
function createConciliacionBancariaTemplate_(ss) {
    const sn = 'B17_Conciliacion_Bancaria';
    const s = getOrCreateSheet_(ss, sn, 15, 100);
    applyBaseSheetStyles_(s);
    s.getRange('A1:E1').merge().setValue('Movimientos Auxiliar Contable (Libros)').setHorizontalAlignment('center');
    s.getRange('G1:K1').merge().setValue('Movimientos Estado de Cuenta Banco').setHorizontalAlignment('center');
    s.getRange('A2:E2').setValues([['Fecha', 'Concepto', 'Cargo', 'Abono', 'Conciliado?']]);
    s.getRange('G2:K2').setValues([['Fecha', 'Concepto', 'Cargo', 'Abono', 'Conciliado?']]);
    styleKpiCard_(s, 'M3:N4', 'Diferencia');
    s.getRange('M3').setFormula('=SUM(C:C,J:J)-SUM(D:D,I:I)');
    logAction_(`Plantilla creada: ${sn}`);
}
function createIvaTemplate_(ss) {
    const sn = 'B18_IVA_Control';
    const s = getOrCreateSheet_(ss, sn, 20, 100);
    applyBaseSheetStyles_(s);
    s.getRange('A1:E1').setValues([['IVA Trasladado (Ventas)', 'Fecha', 'Base', 'Tasa', 'Impuesto']]);
    s.getRange('G1:K1').setValues([['IVA Acreditable (Compras)', 'Fecha', 'Base', 'Tasa', 'Impuesto']]);
    s.getRange('M1:N1').setValues([['Determinación Mensual', '']]);
    s.getRange('M2:N4').setValues([
      ['IVA Trasladado (Cobrado)', ''],
      ['(-) IVA Acreditable (Pagado)', ''],
      ['(=) IVA a Cargo / Favor', '=N2-N3']
    ]);
    s.getRange('N2').setFormula('=SUM(E:E)');
    s.getRange('N3').setFormula('=SUM(K:K)');
    styleKpiCard_(s, 'P2:Q3', 'IVA a Pagar');
    s.getRange('P2').setFormula('=N4');
    logAction_(`Plantilla creada: ${sn}`);
}
function createIsrPmTemplate_(ss) {
    const sn = 'B19_ISR_PM_Provisional';
    const s = getOrCreateSheet_(ss, sn, 10, 50);
    applyBaseSheetStyles_(s);
    s.getRange('A1').setValue('Cálculo de Pago Provisional ISR PM').setFontWeight('bold').setFontSize(14);
    s.getRange('A3:B3').setValues([['Coeficiente de Utilidad', '=Coeficiente_Utilidad_Ejemplo']]);
    s.getRange('A4:B10').setValues([
      ['Ingresos Nominales del Periodo', ''],
      ['Utilidad Fiscal Estimada', ''],
      ['(-) Pagos Provisionales Anteriores', ''],
      ['(-) PTU Pagada en el Ejercicio', ''],
      ['(-) Pérdidas Fiscales Anteriores', ''],
      ['= Base del Pago Provisional', ''],
      ['= ISR a Cargo (Tasa: 30%)', '']
    ]);
    s.getRange('B5').setFormula('=B4*B3');
    s.getRange('B9').setFormula('=MAX(0, B5-B6-B7-B8)');
    s.getRange('B10').setFormula('=B9*ISR_Tasa_PM');
    logAction_(`Plantilla creada: ${sn}`);
}
function createPtuTemplate_(ss) {
    const sn = 'B20_PTU_Calculo';
    const s = getOrCreateSheet_(ss, sn, 10, 50);
    applyBaseSheetStyles_(s);
    s.getRange('A1').setValue('Cálculo de PTU').setFontWeight('bold').setFontSize(14);
    s.getRange('A3:B6').setValues([
      ['Utilidad Fiscal del Ejercicio (UAI)', `=B14_Estado_Resultados!B7`],
      ['(-) Partidas no deducibles', ''],
      ['= Base Repartible PTU', '=B3-B4'],
      ['Total PTU a Repartir (10%)', '=B5*PTU_Porcentaje']
    ]);
    logAction_(`Plantilla creada: ${sn}`);
}
function createActivosFijosTemplate_(ss) {
    const sn = 'B21_Activos_Fijos_Dep';
    const s = getOrCreateSheet_(ss, sn, 10, 100);
    applyBaseSheetStyles_(s);
    s.getRange('A1:H1').setValues([['ID Activo', 'Descripción', 'Fecha Adquisición', 'MOI', 'Tasa Dep. Anual', 'Depreciación Mensual', 'Dep. Acumulada', 'Valor en Libros']]);
    styleHeaderRow_(s, 'A1:H1');
    s.getRange('F2:F').setFormulaR1C1('=IF(RC[-2]<>"", (RC[-2]*RC[-1])/12, "")');
    s.getRange('H2:H').setFormulaR1C1('=IF(RC[-4]<>"", RC[-4]-RC[-1], "")');
    s.getRange('A10').setValue('TOTALES').setFontWeight('bold');
    s.getRange('F10').setFormula('=SUM(F2:F9)');
    logAction_(`Plantilla creada: ${sn}`);
}
function createRetencionesTemplate_(ss) {
    const sn = 'B22_Retenciones_Terceros';
    const s = getOrCreateSheet_(ss, sn, 10, 100);
    applyBaseSheetStyles_(s);
    s.getRange('A1:G1').setValues([['Fecha', 'Proveedor/Cliente', 'RFC', 'Concepto', 'Base Retención', 'Tasa (IVA/ISR)', 'Importe Retenido']]);
    styleHeaderRow_(s, 'A1:G1');
    s.getRange('G2:G').setFormulaR1C1('=RC[-2]*RC[-1]');
    styleKpiCard_(s, 'I2:J3', 'Total Retenido (Mes)');
    s.getRange('I2').setFormula('=SUMIFS(G2:G, A2:A, ">="&EOMONTH(TODAY(),-1)+1, A2:A, "<="&EOMONTH(TODAY(),0))');
    logAction_(`Plantilla creada: ${sn}`);
}
function createDeclaracionesProvisionalesTemplate_(ss) {
    const sn = 'B23_Declaraciones_Prov';
    const s = getOrCreateSheet_(ss, sn, 15, 50);
    applyBaseSheetStyles_(s);
    s.getRange('A1').setValue('Resumen para Declaraciones Provisionales Mensuales').setFontWeight('bold').setFontSize(14);
    s.getRange('B3:C6').setValues([
      ['ISR a Cargo', `=B19_ISR_PM_Provisional!B10`],
      ['IVA a Cargo / Favor', `=B18_IVA_Control!N4`],
      ['Retenciones de ISR a enterar', ''],
      ['Retenciones de IVA a enterar', '']
    ]);
    logAction_(`Plantilla creada: ${sn}`);
}
function createTableroFiscalTemplate_(ss) {
    const sn = 'B24_Tablero_Fiscal';
    const s = getOrCreateSheet_(ss, sn, 15, 50);
    applyBaseSheetStyles_(s);
    s.getRange('A1').setValue('Tablero de Control Fiscal').setFontSize(18).setFontWeight('bold');
    styleKpiCard_(s, 'B3:D4', 'Próximo Vencimiento');
    s.getRange('B3').setValue('17 del mes siguiente');
    styleKpiCard_(s, 'F3:H4', 'ISR Provisional (Mes)');
    s.getRange('F3').setFormula(`=B23_Declaraciones_Prov!C3`);
    styleKpiCard_(s, 'J3:L4', 'IVA por Pagar (Mes)');
    s.getRange('J3').setFormula(`=B23_Declaraciones_Prov!C4`);
    s.getRange('B7').setValue('Semáforo de Obligaciones (Ejemplo)');
    s.getRange('B8:C10').setValues([['Declaración Mensual', 'OK'], ['DIOT', 'PENDIENTE'], ['Contabilidad Electrónica', 'OK']]);
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo("PENDIENTE").setBackground("#f4cccc")
      .whenTextEqualTo("OK").setBackground("#d9ead3")
      .setRanges([s.getRange('C8:C10')]).build();
    s.getConditionalFormatRules().push(rule);
    logAction_(`Plantilla creada: ${sn}`);
}
