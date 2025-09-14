/**
 * //== File: formulas.gs
 * Generadores de fórmulas complejas y gestor de rangos con nombre.
 */

/**
 * Crea los rangos con nombre esenciales para los cálculos fiscales y financieros.
 * Estos rangos actúan como variables globales dentro de la hoja de cálculo.
 * Se almacenarán en una hoja de configuración oculta.
 */
function createNamedRanges_(ss) {
  const configSheetName = '_CONFIG';
  let configSheet = ss.getSheetByName(configSheetName);
  if (!configSheet) {
    configSheet = ss.insertSheet(configSheetName).hideSheet();
    configSheet.getRange('A1:B1').setValues([['Clave', 'Valor']]).setFontWeight('bold');
  }

  const namedRanges = {
    'IVA_Tasa_General': 0.16,
    'ISR_Tasa_PM': 0.30,
    'PTU_Porcentaje': 0.10,
    'Coeficiente_Utilidad_Ejemplo': 0.185 // Ejemplo para cálculo de ISR provisional
  };

  const rangeData = Object.entries(namedRanges);
  configSheet.getRange(2, 1, rangeData.length, 2).setValues(rangeData);

  rangeData.forEach((item, index) => {
    const rangeName = item[0];
    const cell = configSheet.getRange(index + 2, 2);
    try {
      ss.setNamedRange(rangeName, cell);
    } catch (e) {
      // El rango con nombre ya podría existir, lo cual está bien.
      console.warn(`No se pudo crear el rango con nombre ${rangeName}: ${e.message}`);
    }
  });

  configSheet.getRange('B2:B' + (rangeData.length + 1)).setNumberFormat('0.00%');

  logAction_('Rangos con nombre creados/actualizados: ' + Object.keys(namedRanges).join(', '));
}

/**
 * NOTA SOBRE LA GENERACIÓN DE FÓRMULAS:
 * Para este proyecto, se ha tomado la decisión de diseño de integrar las fórmulas
 * directamente en las funciones de creación de plantillas (en los archivos templates_*.gs)
 * usando `setFormula` o `setFormulaR1C1`.
 *
 * Este enfoque mantiene toda la lógica de una plantilla (estructura, estilo y fórmulas)
 * encapsulada en una única función, mejorando la legibilidad y el mantenimiento
 * a nivel de plantilla individual.
 *
 * Este archivo, `formulas.gs`, se utiliza principalmente para gestionar lógicas
 * que son verdaderamente globales, como la creación de Rangos con Nombre.
 * Si en el futuro se necesitaran fórmulas muy complejas y reutilizables en varias
 * plantillas, este sería el lugar ideal para crear funciones generadoras para ellas.
 *
 * Ejemplo de una función generadora de fórmula:
 *
 * function getVlookupFormula_(sheetName, searchKeyRange, dataRange, columnIndex) {
 *   return `=ARRAYFORMULA(IFERROR(VLOOKUP(${searchKeyRange}, '${sheetName}'!${dataRange}, ${columnIndex}, FALSE)))`;
 * }
 */
