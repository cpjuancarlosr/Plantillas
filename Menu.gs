/**
 * @fileoverview Gestión del Menú Personalizado para el Business OS.
 *
 * Descripción:
 * Este archivo centraliza la creación y gestión del menú "ECD OS" en la
 * interfaz de usuario de Google Sheets. Se activa automáticamente al abrir
 * el documento a través de la función `onOpen`.
 *
 * @author Tu Nombre/Empresa
 * @version 1.0
 */

/**
 * Trigger que se ejecuta automáticamente al abrir la hoja de cálculo.
 * Su única responsabilidad es llamar a la función que crea el menú.
 * Esto asegura que el menú esté siempre disponible para el usuario.
 */
function onOpen() {
  createCustomMenu();
}

/**
 * Crea el menú personalizado "ECD OS" en la interfaz de usuario.
 * Agrupa todas las acciones principales del sistema en un lugar accesible.
 */
function createCustomMenu() {
  SpreadsheetApp.getUi()
      .createMenu('ECD OS')
      .addItem('Inicializar Sistema', 'ECD_OS_INIT') // Llama a la función en Init.gs
      .addSeparator()
      .addItem('Recalcular Sistema', 'recalculateSystem') // Llama a la función en Calc.gs
      .addItem('Ejecutar Diagnóstico', 'runSystemCheck') // Llama a la función en Diagnostics.gs
      .addSeparator()
      .addSubMenu(SpreadsheetApp.getUi().createMenu('Gestión de Períodos')
          .addItem('Crear Nuevo Mes', 'createNewMonth') // Llama a la función en Periods.gs
          .addItem('Limpiar Inputs del Mes', 'clearCurrentMonthInputs')) // Llama a la función en Periods.gs
      .addSeparator()
      .addItem('Configuración', 'openConfiguration')
      .addToUi();
}

/**
 * Navega a la hoja de configuración.
 */
function openConfiguration() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAMES.CONFIGURACION);
  if (sheet) {
    SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  } else {
    SpreadsheetApp.getUi().alert('La hoja de configuración no ha sido encontrada.');
  }
}
