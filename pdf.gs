/**
 * //== File: pdf.gs
 * Funciones para la exportación de hojas y tableros a formato PDF.
 */

const PDF_EXPORT_FOLDER = 'Export_PDFs';

/**
 * Función para el menú: Exporta la hoja actualmente activa a PDF.
 */
function exportActiveSheetToPdf() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName().startsWith('_')) {
    SpreadsheetApp.getUi().alert('No se pueden exportar hojas de sistema (las que empiezan con "_").');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  ui.showSidebar(HtmlService.createHtmlOutput('<p>Generando PDF...</p>').setWidth(300));

  try {
    const folder = getOrCreatePdfFolder_(PDF_EXPORT_FOLDER);
    const pdfOptions = {
      size: 'A4',
      portrait: 'false', // Horizontal
      margins: { top: 0.5, bottom: 0.5, left: 0.5, right: 0.5 }
    };

    const pdfFile = exportSheetToPdfHelper_(sheet, pdfOptions, folder);
    const alertMessage = `PDF creado con éxito. <a href="${pdfFile.getUrl()}" target="_blank">Abrir archivo</a>. Se guardó en la carpeta "${PDF_EXPORT_FOLDER}".`;
    const htmlOutput = HtmlService.createHtmlOutput(alertMessage).setWidth(400).setHeight(100);
    ui.showModalDialog(htmlOutput, 'Exportación Exitosa');

  } catch (e) {
    ui.alert('Error al exportar a PDF', e.message, ui.ButtonSet.OK);
    logAction_(`ERROR PDF: ${e.message}`);
  }
}

/**
 * Función para el menú: Exporta todos los tableros y plantillas clave a una carpeta con fecha.
 */
function exportAllDashboardsToPdf() {
  const ss = getActiveSpreadsheet_();
  const allSheets = ss.getSheets();
  // Filtrar para obtener solo las plantillas principales (excluir sistema y hojas secundarias)
  const sheetsToExport = allSheets.filter(s => /^[A-C]\d{2}_/.test(s.getName()));

  if (sheetsToExport.length === 0) {
    SpreadsheetApp.getUi().alert('No se encontraron plantillas para exportar. Por favor, créalas primero.');
    return;
  }

  const ui = SpreadsheetApp.getUi();
  ui.showSidebar(HtmlService.createHtmlOutput('<p>Iniciando exportación en lote... Esto puede tardar.</p>').setWidth(300));

  try {
    const timestamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd_HHmmss');
    const parentFolder = getOrCreatePdfFolder_(PDF_EXPORT_FOLDER);
    const batchFolder = parentFolder.createFolder(`Lote_${timestamp}`);

    sheetsToExport.forEach((sheet, index) => {
      SpreadsheetApp.getActiveSpreadsheet().toast(`Exportando ${index + 1}/${sheetsToExport.length}: ${sheet.getName()}`, 'Progreso', 10);
      const pdfOptions = {
        size: 'A4',
        portrait: sheet.getMaxColumns() > 15 ? 'false' : 'true', // Horizontal para hojas anchas
        margins: { top: 0.5, bottom: 0.5, left: 0.5, right: 0.5 }
      };
      exportSheetToPdfHelper_(sheet, pdfOptions, batchFolder);
    });

    const alertMessage = `Exportación en lote finalizada. Se crearon ${sheetsToExport.length} PDFs. <a href="${batchFolder.getUrl()}" target="_blank">Abrir carpeta de resultados</a>.`;
    const htmlOutput = HtmlService.createHtmlOutput(alertMessage).setWidth(400).setHeight(100);
    ui.showModalDialog(htmlOutput, 'Exportación Exitosa');
    logAction_(`Exportación en lote completada. ${sheetsToExport.length} archivos creados.`);

  } catch (e) {
    ui.alert('Error durante la exportación en lote', e.message, ui.ButtonSet.OK);
    logAction_(`ERROR LOTE PDF: ${e.message}`);
  }
}


/**
 * Obtiene o crea la carpeta de destino para los PDFs en Google Drive.
 * @param {string} folderName - El nombre de la carpeta a buscar o crear.
 * @returns {GoogleAppsScript.Drive.Folder} El objeto de la carpeta.
 */
function getOrCreatePdfFolder_(folderName) {
  const driveRoot = DriveApp.getRootFolder();
  const folders = driveRoot.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return driveRoot.createFolder(folderName);
}


/**
 * Helper genérico para exportar una hoja específica a PDF.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - La hoja a exportar.
 * @param {Object} pdfOptions - Opciones de configuración del PDF.
 * @param {GoogleAppsScript.Drive.Folder} folder - La carpeta de Drive donde guardar el PDF.
 * @returns {GoogleAppsScript.Drive.File} El archivo PDF creado.
 */
function exportSheetToPdfHelper_(sheet, pdfOptions, folder) {
  const ss = sheet.getParent();
  const ssId = ss.getId();
  const sheetId = sheet.getSheetId();

  const url = `https://docs.google.com/spreadsheets/d/${ssId}/export?` +
    `format=pdf&` +
    `gid=${sheetId}&` +
    `size=${pdfOptions.size || 'A4'}&` +
    `portrait=${pdfOptions.portrait || 'false'}&` +
    `fitw=true&` +
    `sheetnames=false&` +
    `printtitle=false&` +
    `gridlines=false&` +
    `top_margin=${pdfOptions.margins.top || 0.5}&` +
    `bottom_margin=${pdfOptions.margins.bottom || 0.5}&` +
    `left_margin=${pdfOptions.margins.left || 0.5}&` +
    `right_margin=${pdfOptions.margins.right || 0.5}`;

  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
      throw new Error(`Error al contactar los servidores de Google para generar el PDF. Código: ${response.getResponseCode()}. Mensaje: ${response.getContentText()}`);
  }

  const blob = response.getBlob().setName(`${sheet.getName()}.pdf`);
  return folder.createFile(blob);
}
