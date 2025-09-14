/**
 * @OnlyCurrentDoc
 * Este script gestiona la creación y administración de un conjunto de plantillas empresariales.
 * Autor: Jules, Arquitecto de Google Apps Script
 * Versión: 1.0
 * Fecha: 2023-10-27
 */

/**
 * Se ejecuta cuando se abre la hoja de cálculo. Crea el menú principal.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏢 Plantillas JC')
    .addItem('🚀 CREAR TODAS LAS PLANTILLAS', 'createAllTemplates')
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📂 Crear por Módulo')
      .addItem('Administración / Operación', 'createAdminModule')
      .addItem('Contabilidad / Fiscal (SAT)', 'createContaFiscalModule')
      .addItem('Productividad de Oficina', 'createOfficeModule'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('📄 Exportar a PDF')
      .addItem('Exportar Hoja Actual', 'exportActiveSheetToPdf')
      .addItem('Exportar Libro Completo (Tableros)', 'exportAllDashboardsToPdf'))
    .addSeparator()
    .addSubMenu(SpreadsheetApp.getUi().createMenu('🛠️ Herramientas')
      .addItem('Cargar Datos de Ejemplo', 'loadAllSeedData')
      .addItem('Limpiar Todos los Datos', 'clearAllData')
      .addItem('Regenerar Formatos y Tableros', 'regenerateAllStyling')
      .addItem('Resetear Validaciones de Datos', 'resetAllValidations'))
    .addSeparator()
    .addItem('🔥 RESETEAR TODO (Borrar Hojas)', 'deleteAllSheetsAndReset')
    .addToUi();
}

/**
 * Wrapper para crear todas las plantillas.
 */
function createAllTemplates() {
  const ss = getActiveSpreadsheet_();
  showSidebarMessage_('Iniciando Creación Total', 'Se generarán las 30 plantillas. Esto puede tardar varios minutos. Por favor, no cierres esta ventana.');

  createReadmeSheet_(ss);
  logAction_('INICIO: Creación de todas las plantillas.');

  // Crear rangos con nombre para variables globales (ej. tasas de impuestos)
  createNamedRanges_(ss);

  // Crear módulos en secuencia
  createAdminModule(false);
  createContaFiscalModule(false);
  createOfficeModule(false);

  logAction_('FIN: Todas las plantillas creadas exitosamente.');
  SpreadsheetApp.getUi().alert('¡Éxito!', 'Se han creado y configurado las 30 plantillas. ¡Bienvenido a tu nuevo centro de control!', SpreadsheetApp.getUi().ButtonSet.OK);
  ss.getSheetByName('README').activate();
}

/**
 * Crea el módulo de Administración.
 * @param {boolean} showAlert - Si se debe mostrar una alerta al finalizar.
 */
function createAdminModule(showAlert = true) {
  const ss = getActiveSpreadsheet_();
  if (showAlert) showSidebarMessage_('Módulo de Administración', 'Creando plantillas de operación...');
  logAction_('INICIO MÓDULO: Administración');
  createAllAdminTemplates_(ss);
  logAction_('FIN MÓDULO: Administración');
  if (showAlert) SpreadsheetApp.getUi().alert('Módulo de Administración creado.');
}

/**
 * Crea el módulo de Contabilidad y Fiscal.
 * @param {boolean} showAlert - Si se debe mostrar una alerta al finalizar.
 */
function createContaFiscalModule(showAlert = true) {
  const ss = getActiveSpreadsheet_();
  if (showAlert) showSidebarMessage_('Módulo Contable/Fiscal', 'Creando plantillas de contabilidad...');
  logAction_('INICIO MÓDULO: Contabilidad/Fiscal');
  createAllContaFiscalTemplates_(ss);
  logAction_('FIN MÓDULO: Contabilidad/Fiscal');
  if (showAlert) SpreadsheetApp.getUi().alert('Módulo Contable/Fiscal creado.');
}

/**
 * Crea el módulo de Productividad de Oficina.
 * @param {boolean} showAlert - Si se debe mostrar una alerta al finalizar.
 */
function createOfficeModule(showAlert = true) {
  const ss = getActiveSpreadsheet_();
  if (showAlert) showSidebarMessage_('Módulo de Oficina', 'Creando plantillas de productividad...');
  logAction_('INICIO MÓDULO: Oficina');
  createAllOfficeTemplates_(ss);
  logAction_('FIN MÓDULO: Oficina');
  if (showAlert) SpreadsheetApp.getUi().alert('Módulo de Oficina creado.');
}

/**
 * Wrapper para cargar todos los datos de ejemplo.
 */
function loadAllSeedData() {
  logAction_('ACCIÓN: Cargar datos de ejemplo.');
  seedAllTemplates_();
  SpreadsheetApp.getUi().alert('Datos de ejemplo cargados.');
}

/**
 * Wrapper para limpiar datos de captura en todas las plantillas.
 */
function clearAllData() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirmación', '¿Estás seguro de que deseas borrar TODOS los datos de captura de las plantillas? Las fórmulas se conservarán.', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    logAction_('ACCIÓN: Limpiar todos los datos.');
    clearAllTemplateData_();
    ui.alert('Proceso completado. Se han limpiado los datos.');
  }
}

/**
 * Wrapper para reaplicar todos los estilos.
 */
function regenerateAllStyling() {
  logAction_('ACCIÓN: Regenerar formatos.');
  applyStylingToAllSheets_();
  SpreadsheetApp.getUi().alert('Formatos y estilos regenerados en todas las plantillas.');
}

/**
 * Wrapper para reaplicar todas las validaciones.
 */
function resetAllValidations() {
  logAction_('ACCIÓN: Resetear validaciones.');
  applyValidationsToAllSheets_();
  SpreadsheetApp.getUi().alert('Validaciones de datos reseteadas en todas las plantillas.');
}

/**
 * Borra todas las hojas excepto la primera para empezar de cero.
 */
function deleteAllSheetsAndReset() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('¡ADVERTENCIA MÁXIMA!', 'Esta acción borrará PERMANENTEMENTE todas las hojas de este libro. ¿Estás absolutamente seguro?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    const ss = getActiveSpreadsheet_();
    const allSheets = ss.getSheets();
    // Crea una hoja temporal para no quedarse sin hojas
    const tempSheet = ss.insertSheet('Inicio');
    allSheets.forEach(sheet => {
      if (sheet.getName() !== 'Inicio') {
        ss.deleteSheet(sheet);
      }
    });
    ss.getSheetByName('Inicio').getRange('A1').setValue('Libro reseteado. Usa el menú "Plantillas JC" para empezar.');
    ss.setActiveSheet(tempSheet);
    logAction_('ACCIÓN: RESETEO TOTAL DEL LIBRO.');
    ui.alert('El libro ha sido reseteado.');
  }
}
