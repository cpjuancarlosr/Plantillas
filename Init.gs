/**
 * @fileoverview Inicializador del Sistema para el Business OS.
 *
 * Descripción:
 * Este archivo contiene la función principal de inicialización `ECD_OS_INIT`.
 * Esta función es responsable de configurar el entorno de la hoja de cálculo
 * por primera vez, asegurando que todas las hojas necesarias existan,
 * aplicando formatos básicos y estableciendo protecciones iniciales.
 * Es el punto de partida para configurar un nuevo cliente.
 *
 * @author ECD OS
 * @version 1.1
 */

/**
 * Función de Inicialización Principal del Sistema ECD OS.
 *
 * Esta función debe ser ejecutada una vez para configurar la hoja de cálculo.
 * Realiza las siguientes acciones:
 * 1. Verifica la existencia de todas las hojas obligatorias definidas en `CONFIG.SHEET_NAMES`.
 * 2. Crea cualquier hoja que falte.
 * 3. Aplica un formato visual base (fondo blanco, fuentes oscuras, sin gridlines).
 * 4. Congela los encabezados en las hojas de datos.
 * 5. Protege las celdas no editables (ej. celdas con fórmulas o KPIs).
 * 6. Crea el menú personalizado "ECD OS" en la interfaz de usuario.
 */
function ECD_OS_INIT() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheetNames = Object.values(CONFIG.SHEET_NAMES);
  const existingSheets = ss.getSheets().map(sheet => sheet.getName());

  // 1. Verificar y crear hojas faltantes
  allSheetNames.forEach(sheetName => {
    if (existingSheets.indexOf(sheetName) === -1) {
      ss.insertSheet(sheetName);
    }
    const sheet = ss.getSheetByName(sheetName);
    // 2. Aplicar formato base y ocultar gridlines
    if (sheet) {
      sheet.setGridLinesVisible(false);
      sheet.getRange('A1:Z1000').setBackground('#FFFFFF').setFontColor('#000000');
      // 3. Congelar encabezados (ej. la primera fila)
      sheet.setFrozenRows(1);
    }
  });

  // 4. Proteger rangos definidos en la configuración
  const protectedRanges = CONFIG.PROTECTED_RANGES;
  for (const key in protectedRanges) {
    const rangeString = protectedRanges[key];
    const range = ss.getRange(rangeString);
    if (range) {
      const protection = range.protect();
      // CORRECCIÓN: Asegurar que el dueño de la hoja pueda editar los rangos protegidos.
      // Esto es crucial para que el script pueda escribir en ellos.
      const me = Session.getEffectiveUser();
      protection.addEditor(me);
      // Opcional: remover otros editores si es necesario.
      // protection.removeEditors(protection.getEditors().filter(editor => editor.getEmail() !== me.getEmail()));
      protection.setDescription('Rango protegido por el sistema ECD OS.');
    }
  }

  // 5. Configurar validación de datos
  setupDataValidation();

  // 6. Aplicar seguridad basada en roles
  applyRoleBasedProtections();

  // 7. Crear menú personalizado (llamando a la función del archivo Menu.gs)
  createCustomMenu();

  // 8. Configurar los triggers automáticos
  setupTriggers();

  // 9. Mensaje de finalización
  SpreadsheetApp.getUi().alert('¡Sistema inicializado correctamente! La automatización ha sido configurada.');
}

/**
 * Crea o actualiza los triggers necesarios para la automatización del sistema.
 * Borra los triggers antiguos para evitar duplicados.
 */
function setupTriggers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Borrar todos los triggers existentes para este proyecto para evitar duplicados
  const existingTriggers = ScriptApp.getUserTriggers(ss);
  existingTriggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });

  // Crear trigger onEdit
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  // Crear trigger diario para recalcular el sistema
  ScriptApp.newTrigger('recalculateSystem')
    .forSpreadsheet(ss)
    .timeBased()
    .everyDays(1)
    .atHour(1) // Ejecutar a la 1 AM
    .create();

  // Crear trigger semanal para alertas fiscales
  ScriptApp.newTrigger('detectTaxDeadlines')
    .forSpreadsheet(ss)
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(8) // Ejecutar los lunes a las 8 AM
    .create();

  Logger.log('Triggers automáticos configurados correctamente.');
}
