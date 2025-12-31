/**
 * @fileoverview Gestión de Períodos para el Business OS.
 *
 * Descripción:
 * Este archivo maneja la lógica de creación, archivo y bloqueo de períodos
 * contables (generalmente mensuales). Asegura que los datos históricos
 * se mantengan íntegros y prepara el sistema para un nuevo ciclo de inputs.
 *
 * @author Tu Nombre/Empresa
 * @version 1.0
 */

/**
 * Crea un nuevo período mensual.
 * Duplica las hojas de input (Ingresos, Egresos) y las renombra con el nuevo mes.
 * Ejemplo: "03. Ingresos" -> "03. Ingresos (Ene 2024)".
 */
function createNewMonth() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Crear Nuevo Mes', 'Introduce el nombre para el nuevo mes (ej. "Feb 2024"):', ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() == ui.Button.OK) {
    const newMonthName = response.getResponseText();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetsToDuplicate = CONFIG.PERIOD_MANAGEMENT.SHEETS_TO_DUPLICATE;

    ss.toast('Iniciando creación del nuevo mes...', 'ECD OS');

    sheetsToDuplicate.forEach(sheetName => {
      const originalSheet = ss.getSheetByName(sheetName);
      if (originalSheet) {
        const newSheet = originalSheet.copyTo(ss);
        newSheet.setName(`${sheetName} (${newMonthName})`);

        // REFACTORIZACIÓN: Usar el nuevo formato de rango más robusto.
        const editableRangeA1 = CONFIG.EDITABLE_RANGES[sheetName];
        if (editableRangeA1) {
          const dataRange = newSheet.getRange(editableRangeA1);
          dataRange.clearContent();
        }
      }
    });

    ui.alert(`Mes "${newMonthName}" creado. Las hojas de input han sido duplicadas.`);
  }
}

/**
 * Limpia los rangos de input del mes actual.
 * Borra el contenido de las celdas definidas en `CONFIG.EDITABLE_RANGES`.
 */
function clearCurrentMonthInputs() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirmar Limpieza', '¿Estás seguro de que quieres borrar todos los datos de input del mes actual? Esta acción no se puede deshacer.', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.toast('Limpiando inputs...', 'ECD OS');

    for (const sheetName in CONFIG.EDITABLE_RANGES) {
      const rangeA1 = CONFIG.EDITABLE_RANGES[sheetName];
      const sheet = ss.getSheetByName(sheetName);
      if (sheet) {
        const range = sheet.getRange(rangeA1);
        range.clearContent();
      }
    }

    ss.toast('Inputs limpiados.', 'ECD OS', 5);
  }
}

/**
 * Archiva y bloquea un mes cerrado.
 * Renombra la hoja y la protege contra futuras ediciones.
 */
function archiveCurrentMonth() {
  // Esta es una función más compleja que podría involucrar:
  // 1. Pedir al usuario el mes a cerrar.
  // 2. Renombrar las hojas de ese mes (ej. a "ZZ. Ingresos (Ene 2024)").
  // 3. Aplicar protección a toda la hoja, permitiendo solo a los 'owners'.
  SpreadsheetApp.getUi().alert('La función de archivar el mes aún no está completamente implementada.');
}
