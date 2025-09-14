/**
 * //== File: validations.gs
 * Lógica para crear y aplicar reglas de validación de datos.
 */

// Listas de Opciones para Validaciones
const KANBAN_STATUSES = ['Pendiente', 'En Progreso', 'Completado', 'En Espera'];
const TASK_PRIORITIES = ['Alta', 'Media', 'Baja'];
const CRM_STAGES = ['Nuevo', 'Contactado', 'Calificado', 'Propuesta', 'Negociación', 'Ganado', 'Perdido'];
const INVENTORY_MOVEMENTS = ['Entrada', 'Salida', 'Ajuste'];
const INVOICE_STATUS_SAT = ['Vigente', 'Cancelado'];
const PAYMENT_STATUS = ['Pagado', 'Pendiente', 'Vencido'];
const YES_NO = ['Sí', 'No'];

/**
 * Función principal para (re)aplicar validaciones a todas las hojas relevantes.
 * Se llama desde el menú "Resetear validaciones".
 */
function applyValidationsToAllSheets_() {
    const ss = getActiveSpreadsheet_();
    const allSheets = ss.getSheets();

    allSheets.forEach(sheet => {
        const sheetName = sheet.getName();
        // Usamos un switch para llamar a la función de validación correcta para cada hoja.
        switch (sheetName) {
            case 'A01_Proyectos_Kanban':
                applyKanbanValidations_(sheet);
                break;
            case 'A02_CRM_Pipeline':
                applyCrmValidations_(sheet);
                break;
            case 'A03_Inventarios':
                applyInventoryValidations_(sheet);
                break;
            // Añadir casos para otras hojas que requieran validación
            // ...
        }
    });
    logAction_('Todas las validaciones han sido reseteadas/aplicadas.');
}


/**
 * Aplica las validaciones de datos a la plantilla de Proyectos (Kanban).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja 'A01_Proyectos_Kanban'.
 */
function applyKanbanValidations_(sheet) {
    // Validación para la columna de Estado (E)
    const statusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(KANBAN_STATUSES)
        .setAllowInvalid(false)
        .setHelpText('Selecciona un estado de la lista.')
        .build();
    sheet.getRange('E2:E').setDataValidation(statusRule);

    // Validación para la columna de Prioridad (F)
    const priorityRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(TASK_PRIORITIES)
        .setAllowInvalid(false)
        .setHelpText('Selecciona una prioridad.')
        .build();
    sheet.getRange('F2:F').setDataValidation(priorityRule);

    logAction_(`Validaciones aplicadas a ${sheet.getName()}`);
}


/**
 * Aplica las validaciones de datos a la plantilla de CRM.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja 'A02_CRM_Pipeline'.
 */
function applyCrmValidations_(sheet) {
    const stageRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(CRM_STAGES)
        .setAllowInvalid(false)
        .setHelpText('Selecciona la etapa actual del lead.')
        .build();
    sheet.getRange('D2:D').setDataValidation(stageRule);

    logAction_(`Validaciones aplicadas a ${sheet.getName()}`);
}

/**
 * Aplica las validaciones de datos a la plantilla de Inventarios.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet La hoja 'A03_Inventarios'.
 */
function applyInventoryValidations_(sheet) {
    const movementRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(INVENTORY_MOVEMENTS)
        .setAllowInvalid(false)
        .setHelpText('Selecciona el tipo de movimiento.')
        .build();
    sheet.getRange('B2:B').setDataValidation(movementRule);

    logAction_(`Validaciones aplicadas a ${sheet.getName()}`);
}

// Se pueden añadir más funciones específicas para otras plantillas aquí.
// Ejemplo: applyCuentasPagarValidations_(sheet), etc.
// ...
