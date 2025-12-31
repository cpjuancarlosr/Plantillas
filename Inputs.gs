/**
 * @fileoverview Motor de Inputs para el Business OS.
 *
 * Descripción:
 * Este archivo contendrá todas las funciones relacionadas con la validación,
 * normalización y gestión de los datos que el usuario introduce en el sistema.
 * El objetivo es asegurar la calidad y consistencia de los datos antes de
 * que sean utilizados en los cálculos.
 *
 * @author Tu Nombre/Empresa
 * @version 1.0
 */

/**
 * Configura las reglas de validación de datos para las celdas de categoría.
 * Crea menús desplegables basados en una lista central en la hoja de Configuración.
 */
function setupDataValidation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Obtener el rango que contiene la lista de categorías
  const categoriasSourceRange = ss.getRange(CONFIG.VALIDATION_RANGES.CATEGORIAS_SOURCE);
  const rule = SpreadsheetApp.newDataValidation().requireValueInRange(categoriasSourceRange).build();

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
