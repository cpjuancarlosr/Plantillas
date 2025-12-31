/**
 * @fileoverview Gestión de Períodos para el Business OS.
 *
 * Descripción:
 * Este archivo maneja la lógica de creación, archivo y bloqueo de períodos
 * contables (generalmente mensuales). Asegura que los datos históricos
 * se mantengan íntegros y prepara el sistema para un nuevo ciclo de inputs.
 *
 * @author ECD OS
 * @version 1.1
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
    let newSheetNames = {};

    ss.toast('Iniciando creación del nuevo mes...', 'ECD OS');

    sheetsToDuplicate.forEach(sheetName => {
      const originalSheet = ss.getSheetByName(sheetName);
      if (originalSheet) {
        const newSheet = originalSheet.copyTo(ss);
        const newName = `${sheetName} (${newMonthName})`;
        newSheet.setName(newName);

        if (sheetName.includes('Ingresos')) {
          newSheetNames.income = newName;
        } else if (sheetName.includes('Egresos')) {
          newSheetNames.expense = newName;
        }

        const editableRangeA1 = CONFIG.EDITABLE_RANGES[sheetName];
        if (editableRangeA1) {
          const dataRange = newSheet.getRange(editableRangeA1);
          dataRange.clearContent();
        }
      }
    });

    // Actualizar el período activo para que los cálculos usen las nuevas hojas
    if (newSheetNames.income && newSheetNames.expense) {
      updateActivePeriod(newSheetNames.income, newSheetNames.expense);
    }

    ui.alert(`Mes "${newMonthName}" creado y activado. Las hojas de input han sido duplicadas.`);
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
 * Archiva y bloquea las hojas del período activo.
 * Renombra las hojas y las protege contra futuras ediciones.
 */
function archiveCurrentMonth() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Confirmar Archivo', '¿Estás seguro de que quieres archivar y bloquear el período activo? Esta acción no se puede deshacer.', ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CONFIGURACION);

    if (!configSheet) {
      ui.alert('Error: No se encontró la hoja de configuración.');
      return;
    }

    const activeIncomeSheetName = configSheet.getRange(CONFIG.ACTIVE_PERIOD_CONFIG.ACTIVE_INCOME_SHEET_CELL).getValue();
    const activeExpenseSheetName = configSheet.getRange(CONFIG.ACTIVE_PERIOD_CONFIG.ACTIVE_EXPENSE_SHEET_CELL).getValue();

    const sheetsToArchive = [ss.getSheetByName(activeIncomeSheetName), ss.getSheetByName(activeExpenseSheetName)];

    ss.toast('Archivando período...', 'ECD OS');

    sheetsToArchive.forEach(sheet => {
      if (sheet) {
        // 1. Renombrar la hoja
        sheet.setName(`Archivado - ${sheet.getName()}`);

        // 2. Proteger la hoja completamente
        const protection = sheet.protect();
        protection.setDescription('Período archivado. Solo lectura.');

        // 3. Asegurar que solo los 'owners' puedan (potencialmente) editar
        const owners = CONFIG.USER_ROLES.OWNERS;
        protection.addEditors(owners);

        // Opcional: Remover a todos los demás editores
        const editors = protection.getEditors();
        editors.forEach(editor => {
          if (owners.indexOf(editor.getEmail()) === -1) {
            protection.removeEditor(editor);
          }
        });
      }
    });

    // Limpiar las celdas del período activo
    configSheet.getRange(CONFIG.ACTIVE_PERIOD_CONFIG.ACTIVE_INCOME_SHEET_CELL).clearContent();
    configSheet.getRange(CONFIG.ACTIVE_PERIOD_CONFIG.ACTIVE_EXPENSE_SHEET_CELL).clearContent();

    ui.alert('El período activo ha sido archivado y protegido.');
  }
}

/**
 * Actualiza las celdas de configuración con los nombres de las hojas del nuevo período activo.
 * @param {string} newIncomeSheetName El nombre de la nueva hoja de ingresos.
 * @param {string} newExpenseSheetName El nombre de la nueva hoja de egresos.
 */
function updateActivePeriod(newIncomeSheetName, newExpenseSheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CONFIGURACION);

  if (!configSheet) {
    SpreadsheetApp.getUi().alert('Error Crítico: No se encontró la hoja "30. Configuración" para actualizar el período activo.');
    return;
  }

  configSheet.getRange(CONFIG.ACTIVE_PERIOD_CONFIG.ACTIVE_INCOME_SHEET_CELL).setValue(newIncomeSheetName);
  configSheet.getRange(CONFIG.ACTIVE_PERIOD_CONFIG.ACTIVE_EXPENSE_SHEET_CELL).setValue(newExpenseSheetName);

  Logger.log(`Período activo actualizado a: ${newIncomeSheetName}, ${newExpenseSheetName}`);
}
