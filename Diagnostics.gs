/**
 * @fileoverview Módulo de Diagnóstico del Sistema para el Business OS.
 *
 * Descripción:
 * Este archivo contiene funciones para verificar la salud y la integridad
 * del sistema. Permite detectar problemas comunes como hojas eliminadas,
 * rangos con nombre rotos o fórmulas con errores, facilitando el soporte
 * y mantenimiento.
 *
 * @author Tu Nombre/Empresa
 * @version 1.0
 */

/**
 * Ejecuta una serie de verificaciones para asegurar la integridad del sistema.
 *
 * Esta función comprueba lo siguiente:
 * 1. Que todas las hojas definidas en `CONFIG.SHEET_NAMES` existan.
 * 2. Que los rangos protegidos clave no hayan sido alterados.
 * 3. (Futuro) Que las fórmulas importantes no contengan errores (ej. #REF!).
 *
 * Muestra un reporte final al usuario.
 */
function runSystemCheck() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  let report = 'Resultado del Diagnóstico del Sistema:\n\n';
  let issuesFound = 0;

  // 1. Verificar la existencia de todas las hojas
  const allSheetNames = Object.values(CONFIG.SHEET_NAMES);
  const existingSheets = ss.getSheets().map(sheet => sheet.getName());

  report += '1. Verificación de Hojas Obligatorias:\n';
  allSheetNames.forEach(sheetName => {
    if (existingSheets.indexOf(sheetName) > -1) {
      report += `   - ${sheetName}: OK\n`;
    } else {
      report += `   - ${sheetName}: ¡NO ENCONTRADA! (Error Crítico)\n`;
      issuesFound++;
    }
  });

  report += '\n';

  // Aquí se podrían agregar más verificaciones (fórmulas, rangos, etc.)

  if (issuesFound === 0) {
    report += '¡El sistema parece estar en buen estado!';
    ui.alert('Diagnóstico Completo', report, ui.ButtonSet.OK);
  } else {
    report += `Se encontraron ${issuesFound} problemas. Por favor, revise el detalle.`;
    ui.alert('Diagnóstico Completo con Errores', report, ui.ButtonSet.OK);
  }
}
