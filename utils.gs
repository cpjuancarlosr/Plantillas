/**
 * //== File: utils.gs
 * Funciones de utilidad y helpers para operaciones comunes y optimización.
 */

const TIMEZONE = 'America/Merida';
const THEME_COLORS = {
  BACKGROUND: '#ffffff', // Blanco
  TEXT: '#111111',       // Negro
  ACCENT: '#00a878',     // Verde Financiero
  GRID: '#e6e6e6',       // Gris Suave
  HEADER: '#f5f5f5'      // Gris muy claro para encabezados
};

/**
 * Obtiene la hoja de cálculo activa de forma segura.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} La hoja de cálculo activa.
 */
function getActiveSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('No se pudo obtener la hoja de cálculo activa.');
  }
  ss.setSpreadsheetTimeZone(TIMEZONE);
  return ss;
}

/**
 * Registra una acción en la hoja de LOG.
 * @param {string} message - El mensaje a registrar.
 */
function logAction_(message) {
  const ss = getActiveSpreadsheet_();
  let logSheet = ss.getSheetByName('_LOG');
  if (!logSheet) {
    logSheet = ss.insertSheet('_LOG').hideSheet();
    logSheet.getRange('A1:C1').setValues([['Timestamp', 'Acción', 'Usuario']]).setFontWeight('bold');
    logSheet.setColumnWidth(1, 150);
    logSheet.setColumnWidth(2, 500);
  }
  const timestamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  const user = Session.getActiveUser().getEmail() || 'N/A';
  logSheet.appendRow([timestamp, message, user]);
}

/**
 * Crea o encuentra una hoja, y la formatea con un tamaño estándar.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - La hoja de cálculo.
 * @param {string} sheetName - El nombre de la hoja a crear/obtener.
 * @param {number} cols - Número de columnas.
 * @param {number} rows - Número de filas.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} La hoja creada o encontrada.
 */
function getOrCreateSheet_(ss, sheetName, cols = 26, rows = 100) {
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    sheet.clear(); // Limpia la hoja si ya existe
    sheet.clearFormats();
  } else {
    sheet = ss.insertSheet(sheetName);
  }
  // Asegura que hay suficientes filas/columnas
  if (sheet.getMaxRows() < rows) sheet.insertRowsAfter(sheet.getMaxRows(), rows - sheet.getMaxRows());
  if (sheet.getMaxColumns() < cols) sheet.insertColumnsAfter(sheet.getMaxColumns(), cols - sheet.getMaxColumns());
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).clearContent();
  return sheet;
}


/**
 * Aplica formato de lote usando la API avanzada de Sheets.
 * @param {string} spreadsheetId - El ID de la hoja de cálculo.
 * @param {Array<Object>} requests - Un array de objetos de solicitud de formato.
 */
function batchUpdate_(spreadsheetId, requests) {
  try {
    Sheets.Spreadsheets.batchUpdate({ requests: requests }, spreadsheetId);
    logAction_(`BatchUpdate ejecutado con ${requests.length} solicitudes.`);
  } catch (e) {
    console.error('Fallo en batchUpdate: ' + e.message);
    logAction_(`ERROR en batchUpdate: ${e.message}`);
  }
}

/**
 * Protege rangos específicos en una hoja.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - La hoja donde aplicar la protección.
 * @param {Array<string>} a1Notations - Array de rangos en notación A1 para proteger.
 * @param {string} description - Descripción de la protección.
 */
function protectRanges_(sheet, a1Notations, description = 'Celdas con fórmulas y encabezados (no editar)') {
  const protection = sheet.protect().setDescription(description);
  protection.setUnprotectedRanges(sheet.getRangeList(a1Notations).getRanges().map(r => r.getA1Notation()));

  // Restringir la edición a solo el propietario del documento.
  const me = Session.getEffectiveUser();
  protection.addEditor(me);
  protection.removeEditors(protection.getEditors());
  if (protection.canDomainEdit()) {
    protection.setDomainEdit(false);
  }
}

/**
 * Muestra un mensaje temporal en la barra lateral.
 * @param {string} title - El título del mensaje.
 * @param {string} message - El contenido del mensaje HTML.
 */
function showSidebarMessage_(title, message) {
  const html = `<div style="font-family: sans-serif; padding: 10px;"><h2>${title}</h2><p>${message}</p><p><i>Este mensaje se cerrará automáticamente.</i></p></div>`;
  const ui = HtmlService.createHtmlOutput(html).setWidth(300).setHeight(150);
  SpreadsheetApp.getUi().showSidebar(ui);
}
