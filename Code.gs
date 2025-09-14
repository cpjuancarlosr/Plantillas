/**
 * @OnlyCurrentDoc
 * Sistema Contable Automatizado con XML (v1.0)
 * Autor: Jules, Arquitecto de Google Apps Script
 *
 * Este script principal maneja la configuración inicial del sistema y la creación del menú.
 */

// --- CONSTANTES GLOBALES ---
const TIMEZONE = 'America/Merida';
const THEME_COLORS = {
  BACKGROUND: '#ffffff',
  TEXT: '#111111',
  ACCENT: '#00A878', // Verde financiero
  GRID: '#e6e6e6',
  HEADER: '#f5f5f5'
};

/**
 * Se ejecuta cuando se abre la hoja de cálculo. Crea el menú personalizado.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Contabilidad XML')
    .addItem('1. Configuración Inicial del Sistema', 'initialSetup')
    .addSeparator()
    .addItem('2. Cargar Archivos XML', 'showSidebar') // Se implementará en el Paso 2 del plan
    .addToUi();
}

/**
 * Realiza la configuración inicial, creando todas las hojas necesarias.
 */
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setSpreadsheetTimeZone(TIMEZONE);

  // Crear cada hoja estructural
  createCatalogoCuentasSheet_(ss);
  createReglasSheet_(ss);
  createPolizasSheet_(ss);
  createCfdiLogSheet_(ss);
  createLogSheet_(ss);

  // Crear hojas de reportes financieros
  createBalanzaSheet_(ss);
  createEstadoResultadosSheet_(ss);
  createBalanceGeneralSheet_(ss);

  // Crear hojas de análisis fiscal y de negocio
  createIvaSheet_(ss);
  createIsrSheet_(ss);
  createDashboardSheet_(ss);

  // Poblar con datos iniciales para que el sistema sea usable desde el inicio
  seedInitialData_(ss);

  SpreadsheetApp.getUi().alert('¡Configuración completada!', 'Se han creado todas las hojas y se ha cargado un catálogo de cuentas y reglas de ejemplo. El sistema está listo para usarse.', SpreadsheetApp.getUi().ButtonSet.OK);
}


// --- FUNCIONES DE CREACIÓN DE HOJAS DE ANÁLISIS Y REPORTES ---

function createIvaSheet_(ss) {
  const sheetName = 'Cálculo de IVA';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) { sheet = ss.insertSheet(sheetName); } else { sheet.clear(); }

  sheet.getRange('A1').setValue('Cálculo de IVA Mensual').setFontWeight('bold').setFontSize(14);
  const concepts = [
    ['IVA Acreditable (de Gastos)', "=SUMIF('Balanza de Comprobación'!B:B, \"*IVA*\", 'Balanza de Comprobación'!E:E)"],
    ['IVA Trasladado (de Ingresos)', 0],
    ['= IVA a Cargo / (Favor)', '=B3-B2']
  ];
  sheet.getRange('A2:B4').setValues(concepts);
  sheet.getRange('B2:B4').setNumberFormat('$#,##0.00');
  sheet.getRange('A2:A4').setFontWeight('bold');
  sheet.getRange('B3').setNote('El IVA trasladado de ingresos no se calcula automáticamente en esta versión.');
}

function createIsrSheet_(ss) {
  const sheetName = 'Cálculo de ISR Provisional';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) { sheet = ss.insertSheet(sheetName); } else { sheet.clear(); }

  sheet.getRange('A1').setValue('Cálculo de Pago Provisional de ISR').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A3').setValue('Coeficiente de Utilidad').setFontWeight('bold');
  sheet.getRange('B3').setNumberFormat('0.00%').setNote('Ingresa aquí el coeficiente de utilidad de tu empresa.');

  const concepts = [
    ['Ingresos Nominales (del mes)', "='Estado de Resultados'!B2"],
    ['Utilidad Fiscal Estimada', '=B3*B4'],
    ['Tasa de ISR', '30.00%'],
    ['= ISR Causado', '=B5*B6'],
    ['(-) Pagos Provisionales Anteriores', ''],
    ['= ISR a Pagar', '=B7-B8']
  ];
  sheet.getRange('A4:B9').setValues(concepts);
  sheet.getRange('B4:B9').setNumberFormat('$#,##0.00');
  sheet.getRange('B6').setNumberFormat('0.00%');
}

function createDashboardSheet_(ss) {
  const sheetName = 'Tablero Financiero';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) { sheet = ss.insertSheet(sheetName); } else { sheet.clear(); }

  sheet.getRange('A1').setValue('Tablero Financiero').setFontWeight('bold').setFontSize(18);

  sheet.getRange('B3').setValue('Liquidez').setFontWeight('bold');
  sheet.getRange('B4').setValue('Razón Circulante');
  sheet.getRange('C4').setFormula("=IFERROR('Balance General'!B3 / 'Balance General'!E3, 0)").setNumberFormat('0.00');
  sheet.getRange('B4:C4').setNote('Activo Circulante / Pasivo Circulante. Idealmente > 1.5');

  sheet.getRange('E3').setValue('Rentabilidad').setFontWeight('bold');
  sheet.getRange('E4').setValue('Margen de Utilidad Neta');
  sheet.getRange('F4').setFormula("=IFERROR('Estado de Resultados'!B6 / 'Estado de Resultados'!B2, 0)").setNumberFormat('0.00%');
  sheet.getRange('E4:F4').setNote('Utilidad Neta / Ingresos Totales.');
}

function createBalanzaSheet_(ss) {
  const sheetName = 'Balanza de Comprobación';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) { sheet = ss.insertSheet(sheetName); } else { sheet.clear(); }

  const headers = ['Cuenta', 'Nombre', 'Debe', 'Haber', 'Saldo Final', 'Tipo'];
  sheet.getRange('A1:F1').setValues([headers]).setFontWeight('bold').setBackground(THEME_COLORS.HEADER);
  sheet.setColumnWidth(6, 150);

  // Fórmulas de ARRAYFORMULA para robustez
  sheet.getRange('A2').setFormula("=SORT(UNIQUE('Catálogo de Cuentas'!A2:A))");
  sheet.getRange('B2').setFormula("=ARRAYFORMULA(IF(A2:A=\"\",,IFERROR(VLOOKUP(A2:A, 'Catálogo de Cuentas'!A:B, 2, FALSE))))");
  sheet.getRange('C2').setFormula("=ARRAYFORMULA(IF(A2:A=\"\",,SUMIF('Pólizas (Diario General)'!C:C, A2:A, 'Pólizas (Diario General)'!E:E)))");
  sheet.getRange('D2').setFormula("=ARRAYFORMULA(IF(A2:A=\"\",,SUMIF('Pólizas (Diario General)'!C:C, A2:A, 'Pólizas (Diario General)'!F:F)))");
  sheet.getRange('E2').setFormula("=ARRAYFORMULA(IF(A2:A=\"\",,IF(VLOOKUP(A2:A, 'Catálogo de Cuentas'!A:D, 4, FALSE)=\"Deudora\", C2:C-D2:D, D2:D-C2:C)))");
  sheet.getRange('F2').setFormula("=ARRAYFORMULA(IF(A2:A=\"\",,IFERROR(VLOOKUP(A2:A, 'Catálogo de Cuentas'!A:C, 3, FALSE))))");

  sheet.getRange('C:E').setNumberFormat('$#,##0.00');
}

// --- DATOS INICIALES ---
function seedInitialData_(ss) {
  const catalogoSheet = ss.getSheetByName('Catálogo de Cuentas');
  const reglasSheet = ss.getSheetByName('Reglas de Categorización');

  // Verificar si ya hay datos para no duplicar
  if (catalogoSheet.getRange('A2').getValue() !== "") return;

  const catalogoData = [
    ['1101', 'Caja', 'Activo', 'Deudora'],
    ['1102', 'Bancos', 'Activo', 'Deudora'],
    ['1105', 'Clientes', 'Activo', 'Deudora'],
    ['1120', 'IVA Acreditable', 'Activo', 'Deudora'],
    ['2101', 'Proveedores', 'Pasivo', 'Acreedora'],
    ['2105', 'Acreedores Diversos', 'Pasivo', 'Acreedora'],
    ['4101', 'Ventas', 'Ingreso', 'Acreedora'],
    ['6101', 'Gastos de Oficina', 'Gasto', 'Deudora'],
    ['6102', 'Servicios Públicos (Luz, Agua)', 'Gasto', 'Deudora'],
    ['6103', 'Renta de Oficina', 'Gasto', 'Deudora']
  ];
  catalogoSheet.getRange(2, 1, catalogoData.length, 4).setValues(catalogoData);

  const reglasData = [
    ['RFC', 'CFE123456ABC', '6102', '1120', '2101'],
    ['RFC', 'TELMEX123ABC', '6102', '1120', '2101'],
    ['PalabraClave', 'Office Depot', '6101', '1120', '2105']
  ];
  reglasSheet.getRange(2, 1, reglasData.length, 5).setValues(reglasData);
}

function createEstadoResultadosSheet_(ss) {
  const sheetName = 'Estado de Resultados';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) { sheet = ss.insertSheet(sheetName); } else { sheet.clear(); }

  sheet.getRange('A1').setValue('Estado de Resultados').setFontWeight('bold').setFontSize(14);
  const concepts = [
    // Fórmulas corregidas para usar la columna de ayuda en la Balanza
    ['Ingresos', "=SUMIF('Balanza de Comprobación'!F:F, \"Ingreso\", 'Balanza de Comprobación'!E:E)"],
    ['(-) Costos', "=SUMIF('Balanza de Comprobación'!F:F, \"Costo\", 'Balanza de Comprobación'!E:E)"],
    ['= Utilidad Bruta', '=B2-B3'],
    ['(-) Gastos', "=SUMIF('Balanza de Comprobación'!F:F, \"Gasto\", 'Balanza de Comprobación'!E:E)"],
    ['= Utilidad Neta', '=B4-B5']
  ];
  sheet.getRange('A2:B6').setValues(concepts);
  sheet.getRange('B2:B6').setNumberFormat('$#,##0.00');
  sheet.getRange('A2:A6').setFontWeight('bold');
}

function createBalanceGeneralSheet_(ss) {
  const sheetName = 'Balance General';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) { sheet = ss.insertSheet(sheetName); } else { sheet.clear(); }

  sheet.getRange('A1').setValue('Balance General').setFontWeight('bold').setFontSize(14);

  sheet.getRange('A2').setValue('ACTIVOS').setFontWeight('bold');
  sheet.getRange('A3').setValue('Total Activos');
  // Fórmula corregida
  sheet.getRange('B3').setFormula("=SUMIF('Balanza de Comprobación'!F:F, \"Activo\", 'Balanza de Comprobación'!E:E)");

  sheet.getRange('D2').setValue('PASIVOS Y CAPITAL').setFontWeight('bold');
  sheet.getRange('D3').setValue('Total Pasivos');
  // Fórmula corregida
  sheet.getRange('E3').setFormula("=SUMIF('Balanza de Comprobación'!F:F, \"Pasivo\", 'Balanza de Comprobación'!E:E)");
  sheet.getRange('D4').setValue('Total Capital');
  // Fórmula corregida
  sheet.getRange('E4').setFormula("=SUMIF('Balanza de Comprobación'!F:F, \"Capital\", 'Balanza de Comprobación'!E:E)");
  sheet.getRange('D5').setValue('Utilidad del Ejercicio');
  sheet.getRange('E5').setFormula("='Estado de Resultados'!B6");

  sheet.getRange('A7').setValue('Total Activo').setFontWeight('bold');
  sheet.getRange('B7').setFormula('=B3');
  sheet.getRange('D7').setValue('Total Pasivo + Capital').setFontWeight('bold');
  sheet.getRange('E7').setFormula('=SUM(E3:E5)');
  sheet.getRange('G7').setValue('Verificación (debe ser 0)');
  sheet.getRange('H7').setFormula('=B7-E7');

  sheet.getRange('B:B,E:E,H:H').setNumberFormat('$#,##0.00');
}


// --- FUNCIONES DE CREACIÓN DE HOJAS ESTRUCTURALES ---

/**
 * Crea la hoja para el Catálogo de Cuentas.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function createCatalogoCuentasSheet_(ss) {
  const sheetName = 'Catálogo de Cuentas';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  const headers = ['Código Cuenta', 'Nombre de la Cuenta', 'Tipo (Activo, Pasivo, Capital, Ingreso, Costo, Gasto)', 'Naturaleza (Deudora/Acreedora)'];
  sheet.getRange('A1:D1').setValues([headers]).setFontWeight('bold').setBackground(THEME_COLORS.HEADER);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 200);
  sheet.getRange('A1:D1').setNote('Define aquí todas las cuentas contables que utilizarás.');
}

/**
 * Crea la hoja para las Reglas de Categorización.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function createReglasSheet_(ss) {
  const sheetName = 'Reglas de Categorización';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  const headers = ['Tipo de Regla (RFC/PalabraClave)', 'Valor (El RFC o la palabra a buscar)', 'Cuenta Contable (Cargo)', 'Cuenta de IVA (Cargo)', 'Cuenta por Pagar (Abono)'];
  sheet.getRange('A1:E1').setValues([headers]).setFontWeight('bold').setBackground(THEME_COLORS.HEADER);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 200);
  sheet.getRange('A1').setNote('Enseña al sistema cómo categorizar las facturas. El script buscará coincidencias en esta hoja para crear las pólizas.');
}

/**
 * Crea la hoja para las Pólizas (Diario General).
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function createPolizasSheet_(ss) {
  const sheetName = 'Pólizas (Diario General)';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  const headers = ['ID Póliza', 'Fecha', 'Cuenta Contable', 'Concepto', 'Debe', 'Haber', 'Origen (UUID del CFDI)'];
  sheet.getRange('A1:G1').setValues([headers]).setFontWeight('bold').setBackground(THEME_COLORS.HEADER);
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 350);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 300);
  sheet.getRange('E:F').setNumberFormat('$#,##0.00');
}

/**
 * Crea la hoja para el Log de CFDI.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function createCfdiLogSheet_(ss) {
  const sheetName = 'Log de CFDI';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clear();
  }

  const headers = ['Timestamp', 'Nombre Archivo', 'UUID', 'RFC Emisor', 'RFC Receptor', 'Total Factura', 'Estado', 'Detalle del Error'];
  sheet.getRange('A1:H1').setValues([headers]).setFontWeight('bold').setBackground(THEME_COLORS.HEADER);
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 150);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 120);
  sheet.setColumnWidth(7, 100);
  sheet.setColumnWidth(8, 400);
}

/**
 * Crea la hoja de LOG oculta para el script.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function createLogSheet_(ss) {
  const sheetName = '_LOG';
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName).hideSheet();
  } else {
    sheet.clear();
  }

  const headers = ['Timestamp', 'Función', 'Mensaje'];
  sheet.getRange('A1:C1').setValues([headers]).setFontWeight('bold');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 500);
}

/**
 * Muestra la barra lateral para la carga de archivos XML.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar.html')
      .setTitle('Cargar Facturas CFDI')
      .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Procesa el contenido de los archivos XML subidos desde la sidebar.
 * @param {Array<Object>} fileObjects Array de objetos con {fileName, content}.
 * @returns {string} Un mensaje de estado en HTML para la sidebar.
 */
function processXmlFiles(fileObjects) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cfdiLogSheet = ss.getSheetByName('Log de CFDI');
  const reglasSheet = ss.getSheetByName('Reglas de Categorización');
  const polizasSheet = ss.getSheetByName('Pólizas (Diario General)');
  const reglasData = reglasSheet.getRange(2, 1, reglasSheet.getLastRow() - 1, 5).getValues();

  let successCount = 0;
  let errorCount = 0;
  let needsRuleCount = 0;

  fileObjects.forEach(fileObject => {
    const timestamp = new Date();
    let logRow = [timestamp, fileObject.fileName, '', '', '', '', 'Iniciando', ''];
    const logRange = cfdiLogSheet.getRange(cfdiLogSheet.getLastRow() + 1, 1, 1, 8);
    logRange.setValues([logRow]);

    try {
      const doc = XmlService.parse(fileObject.content);
      const root = doc.getRootElement();
      const cfdi = XmlService.getNamespace('http://www.sat.gob.mx/cfd/4');
      const tfd = XmlService.getNamespace('http://www.sat.gob.mx/TimbreFiscalDigital');

      const emisor = root.getChild('Emisor', cfdi);
      const receptor = root.getChild('Receptor', cfdi);
      const timbre = root.getChild('Complemento', cfdi).getChild('TimbreFiscalDigital', tfd);

      const rfcEmisor = emisor.getAttribute('Rfc').getValue();
      const rfcReceptor = receptor.getAttribute('Rfc').getValue();
      const total = root.getAttribute('Total').getValue();
      const uuid = timbre.getAttribute('UUID').getValue();
      const fecha = new Date(root.getAttribute('Fecha').getValue());

      logRow[2] = uuid;
      logRow[3] = rfcEmisor;
      logRow[4] = rfcReceptor;
      logRow[5] = parseFloat(total);

      const rule = findCategorizationRule_(rfcEmisor, reglasData);

      if (rule) {
        createJournalEntries_(polizasSheet, rule, total, uuid, fecha);
        logRow[6] = 'Procesado';
        successCount++;
      } else {
        logRow[6] = 'Requiere Regla';
        logRow[7] = `No se encontró una regla para el RFC ${rfcEmisor}.`;
        needsRuleCount++;
      }
    } catch (e) {
      logRow[6] = 'Error';
      logRow[7] = e.message.slice(0, 500);
      errorCount++;
    }
    logRange.setValues([logRow]); // Update log with final status
  });

  return `Proceso finalizado: <br>
          - ${successCount} procesados con éxito.<br>
          - ${needsRuleCount} requieren regla.<br>
          - ${errorCount} con error.`;
}

/**
 * Busca una regla de categorización para un RFC emisor.
 * @param {string} rfc - El RFC a buscar.
 * @param {Array<Array<string>>} rulesData - Los datos de la hoja de reglas.
 * @returns {Object|null} Un objeto con la regla o null si no se encuentra.
 */
function findCategorizationRule_(rfc, rulesData) {
  for (let i = 0; i < rulesData.length; i++) {
    if (rulesData[i][0] === 'RFC' && rulesData[i][1] === rfc) {
      return {
        cargoAccount: rulesData[i][2],
        ivaAccount: rulesData[i][3],
        abonoAccount: rulesData[i][4]
      };
    }
  }
  return null;
}

/**
 * Crea los asientos de diario para una factura procesada.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - La hoja de Pólizas.
 * @param {Object} rule - La regla de categorización a aplicar.
 * @param {string} totalStr - El total de la factura como string.
 * @param {string} uuid - El UUID del CFDI.
 * @param {Date} fecha - La fecha del CFDI.
 */
function createJournalEntries_(sheet, rule, totalStr, uuid, fecha) {
  const total = parseFloat(totalStr);
  const iva = total / 1.16 * 0.16;
  const subtotal = total - iva;
  const polizaId = `P-${uuid.substring(0, 8)}`;

  const entries = [
    // Cargo a Gasto/Activo
    [polizaId, fecha, rule.cargoAccount, `Factura ${uuid}`, subtotal, 0, uuid],
    // Cargo a IVA Acreditable
    [polizaId, fecha, rule.ivaAccount, `IVA de Factura ${uuid}`, iva, 0, uuid],
    // Abono a Proveedor/Banco
    [polizaId, fecha, rule.abonoAccount, `Provisión Factura ${uuid}`, 0, total, uuid]
  ];

  sheet.getRange(sheet.getLastRow() + 1, 1, 3, 7).setValues(entries);
}
