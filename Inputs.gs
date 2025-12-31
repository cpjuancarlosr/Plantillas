/**
 * @fileoverview Motor de Inputs para el Business OS.
 *
 * Descripción:
 * Este archivo contendrá todas las funciones relacionadas con la validación,
 * normalización y gestión de los datos que el usuario introduce en el sistema.
 * El objetivo es asegurar la calidad y consistencia de los datos antes de
 * que sean utilizados en los cálculos.
 *
 * @author ECD OS
 * @version 1.1
 */

/**
 * Configura las reglas de validación de datos para las celdas de categoría.
 * Crea menús desplegables basados en una lista central en la hoja de Configuración.
 */
function setupDataValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Obtener el rango que contiene la lista de categorías
  const categoriasSourceRange = ss.getRange(CONFIG.VALIDATION_RANGES.CATEGORIAS_SOURCE);
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(categoriasSourceRange).setAllowInvalid(false).build();

  // 2. Aplicar la regla a la hoja de Ingresos
  const ingresosSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.INGRESOS);
  if (ingresosSheet) {
    const targetRangeIngresos = ingresosSheet.getRange(CONFIG.VALIDATION_RANGES.INGRESOS_CATEGORIA_TARGET);
    targetRangeIngresos.setDataValidation(rule);
  } else {
    Logger.log(`No se encontró la hoja de Ingresos para aplicar la validación de datos.`);
  }

  // 3. Aplicar la regla a la hoja de Egresos
  const egresosSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.EGRESOS);
  if (egresosSheet) {
    const targetRangeEgresos = egresosSheet.getRange(CONFIG.VALIDATION_RANGES.EGRESOS_CATEGORIA_TARGET);
    targetRangeEgresos.setDataValidation(rule);
  } else {
    Logger.log(`No se encontró la hoja de Egresos para aplicar la validación de datos.`);
  }

  Logger.log('Reglas de validación de datos configuradas correctamente.');
}

/**
 * Función principal que se activa con el trigger onEdit.
 * Orquesta la validación y normalización de datos en tiempo real.
 * @param {Object} e El objeto de evento de edición.
 */
function onEdit(e) {
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  // Obtener nombres de las hojas activas para saber si la edición ocurrió en una de ellas
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CONFIGURACION);
  if (!configSheet) return;

  const activeIncomeSheet = configSheet.getRange(CONFIG.ACTIVE_PERIOD_CONFIG.ACTIVE_INCOME_SHEET_CELL).getValue();
  const activeExpenseSheet = configSheet.getRange(CONFIG.ACTIVE_PERIOD_CONFIG.ACTIVE_EXPENSE_SHEET_CELL).getValue();

  // Salir si la edición no es en una hoja de input activa y por debajo de la fila de inicio
  if ((sheetName !== activeIncomeSheet && sheetName !== activeExpenseSheet) || range.getRow() < CONFIG.INPUT_STRUCTURE.START_ROW) {
    return;
  }

  normalizeData(range);
  checkRowCompleteness(sheet, range.getRow());
}

/**
 * Normaliza el formato de los datos en las columnas de Fecha y Monto.
 * @param {Range} range El rango que fue editado.
 */
function normalizeData(range) {
  const column = range.getColumn();
  const value = range.getValue();

  // Normalizar Fecha
  if (column === CONFIG.INPUT_STRUCTURE.DATE_COLUMN && value) {
    if (Object.prototype.toString.call(value) === "[object Date]") {
      range.setNumberFormat('yyyy-mm-dd');
    }
  }

  // Normalizar Monto
  if (column === CONFIG.INPUT_STRUCTURE.AMOUNT_COLUMN && value) {
    if (typeof value !== 'number') {
      range.setValue(parseFloat(value) || 0);
    }
    range.setNumberFormat('$#,##0.00');
  }
}

/**
 * Verifica si una fila tiene todas las celdas obligatorias llenas.
 * Si no, colorea la fila para indicar que está incompleta.
 * @param {Sheet} sheet La hoja donde ocurrió la edición.
 * @param {number} rowNum El número de la fila a verificar.
 */
function checkRowCompleteness(sheet, rowNum) {
  const requiredColumns = CONFIG.INPUT_STRUCTURE.REQUIRED_COLUMNS;
  const startCol = 1;
  const numCols = sheet.getLastColumn();
  const rowRange = sheet.getRange(rowNum, startCol, 1, numCols);
  const rowValues = rowRange.getValues()[0];

  let isComplete = true;
  for (const colIndex of requiredColumns) {
    if (!rowValues[colIndex - 1]) { // -1 porque los arrays son base 0
      isComplete = false;
      break;
    }
  }

  // Aplicar o quitar el color de fondo
  if (isComplete) {
    rowRange.setBackground(null); // Quitar color de fondo
  } else {
    // Solo colorear si la fila no está completamente vacía
    const isRowEmpty = rowValues.every(cell => cell === "");
    if (!isRowEmpty) {
      rowRange.setBackground(CONFIG.INPUT_STRUCTURE.INCOMPLETE_ROW_COLOR);
    } else {
      rowRange.setBackground(null);
    }
  }
}
