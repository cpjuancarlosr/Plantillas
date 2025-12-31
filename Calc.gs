/**
 * @fileoverview Motor de Cálculos para el Business OS.
 *
 * Descripción:
 * Este archivo es el corazón del sistema. Contiene todas las funciones
 * que realizan los cálculos financieros clave, como proyecciones de caja,
 * cálculo de impuestos, KPIs, márgenes de contribución y simulaciones.
 *
 * @author ECD OS
 * @version 1.1
 */

/**
 * Función principal que orquesta todos los cálculos del sistema.
 * Se puede llamar manualmente desde el menú o automáticamente con un trigger.
 */
function recalculateSystem() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Iniciando recálculo del sistema...', 'ECD OS');

  // Ejecutar cálculos en un orden lógico
  const { totalIngresos } = recalculateKPIs(); // Es importante que los KPIs se calculen primero y devuelvan los ingresos brutos
  recalculateTaxes(totalIngresos); // Pasar los ingresos brutos al cálculo de impuestos
  recalculateCashflowProjections();
  recalculateMargins(totalIngresos);

  // Actualizar elementos dinámicos del Dashboard
  updateRiskSemaphore();
  generateDashboardRecommendations();

  // Al final, ejecutar las alertas para reflejar los nuevos datos
  checkAllAlerts();

  SpreadsheetApp.getActiveSpreadsheet().toast('¡Sistema recalculado con éxito!', 'ECD OS', 5);
}

/**
 * Recalcula las proyecciones de flujo de caja a 30 días.
 * Suma los gastos fijos y los resta de la caja actual.
 */
function recalculateCashflowProjections() {
  Logger.log('Recalculando proyección de caja...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Obtener las hojas y rangos necesarios
  const configSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CONFIGURACION);
  const dashboardSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);

  if (!configSheet || !dashboardSheet) {
    SpreadsheetApp.getUi().alert('Error Crítico: No se encontró la hoja de "Configuración" o "Dashboard". El cálculo no puede continuar.');
    Logger.log('Error: No se encontraron la hoja de Configuración o Dashboard.');
    return;
  }

  // 2. Sumar los gastos fijos desde la hoja de configuración
  const gastosFijosValues = configSheet.getRange(CONFIG.DATA_RANGES.GASTOS_FIJOS_SOURCE_RANGE).getValues();
  const totalGastosFijos = gastosFijosValues.reduce((sum, row) => sum + (parseFloat(row[0]) || 0), 0);

  // 3. Obtener la caja actual del dashboard
  const cajaActual = dashboardSheet.getRange(CONFIG.DATA_RANGES.DASHBOARD_CAJA_HOY_CELL).getValue();

  // 4. Calcular y escribir la caja proyectada
  const cajaProyectada = cajaActual - totalGastosFijos;
  dashboardSheet.getRange(CONFIG.DATA_RANGES.DASHBOARD_CAJA_PROYECTADA_CELL).setValue(cajaProyectada);

  Logger.log(`Proyección de caja actualizada: ${cajaProyectada}`);
}

/**
 * Recalcula los impuestos a provisionar.
 * Utiliza los ingresos brutos para aplicar la tasa de impuesto general.
 * @param {number} totalIngresos El total de ingresos brutos del período.
 */
function recalculateTaxes(totalIngresos) {
  Logger.log('Recalculando impuestos...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Obtener la hoja de impuestos
  const impuestosSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.IMPUESTOS);

  if (!impuestosSheet) {
    SpreadsheetApp.getUi().alert('Error Crítico: No se encontró la hoja de "Impuestos". El cálculo no puede continuar.');
    Logger.log('Error: No se encontró la hoja de Impuestos.');
    return;
  }

  // 2. Calcular el impuesto basado en los ingresos brutos pasados como argumento
  const taxRate = CONFIG.TAX_SETTINGS.GENERAL_TAX_RATE;
  const impuestoCalculado = totalIngresos * taxRate;

  // 4. Escribir el resultado en la hoja de Impuestos
  impuestosSheet.getRange(CONFIG.DATA_RANGES.IMPUESTOS_TOTAL_CELL).setValue(impuestoCalculado);

  Logger.log(`Impuestos calculados y actualizados: ${impuestoCalculado}`);
}

/**
 * Recalcula los Indicadores Clave de Desempeño (KPIs).
 * Lee los totales de Ingresos y Egresos, los escribe en el Dashboard y devuelve los ingresos brutos.
 * @return {{totalIngresos: number}} Un objeto con el total de ingresos brutos.
 */
function recalculateKPIs() {
  Logger.log('Recalculando KPIs...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Obtener los nombres de las hojas activas desde la configuración
  const configSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CONFIGURACION);
  if (!configSheet) {
    SpreadsheetApp.getUi().alert('Error Crítico: No se encontró la hoja de "Configuración".');
    return { totalIngresos: 0 };
  }
  const activeIncomeSheetName = configSheet.getRange(CONFIG.ACTIVE_PERIOD_CONFIG.ACTIVE_INCOME_SHEET_CELL).getValue();
  const activeExpenseSheetName = configSheet.getRange(CONFIG.ACTIVE_PERIOD_CONFIG.ACTIVE_EXPENSE_SHEET_CELL).getValue();

  // 2. Obtener las hojas necesarias usando los nombres del período activo
  const ingresosSheet = ss.getSheetByName(activeIncomeSheetName);
  const egresosSheet = ss.getSheetByName(activeExpenseSheetName);
  const dashboardSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);

  if (!ingresosSheet || !egresosSheet || !dashboardSheet) {
    SpreadsheetApp.getUi().alert(`Error Crítico: No se encontraron una o más hojas activas ("${activeIncomeSheetName}", "${activeExpenseSheetName}") o el "Dashboard". El cálculo no puede continuar.`);
    Logger.log('Error: No se encontraron una o más hojas requeridas para el cálculo de KPIs.');
    return { totalIngresos: 0 };
  }

  // 2. Definir rangos y leer datos desde CONFIG
  const ingresosRange = ingresosSheet.getRange(CONFIG.DATA_RANGES.INGRESOS_MONTO_COL);
  const egresosRange = egresosSheet.getRange(CONFIG.DATA_RANGES.EGRESOS_MONTO_COL);

  const ingresosValues = ingresosRange.getValues();
  const egresosValues = egresosRange.getValues();

  // 3. Calcular totales y saldo de caja
  const totalIngresos = ingresosValues.reduce((sum, row) => sum + (parseFloat(row[0]) || 0), 0);
  const totalEgresos = egresosValues.reduce((sum, row) => sum + (parseFloat(row[0]) || 0), 0);
  const saldoCajaActual = totalIngresos - totalEgresos;

  // 4. Escribir los resultados en el Dashboard usando rangos de CONFIG
  dashboardSheet.getRange(CONFIG.DATA_RANGES.DASHBOARD_CAJA_HOY_CELL).setValue(saldoCajaActual);
  dashboardSheet.getRange(CONFIG.DATA_RANGES.DASHBOARD_TOTAL_EGRESOS_CELL).setValue(totalEgresos);
  dashboardSheet.getRange(CONFIG.DATA_RANGES.DASHBOARD_INGRESOS_BRUTOS_CELL).setValue(totalIngresos);

  Logger.log(`KPIs actualizados: Saldo de Caja Actual = ${saldoCajaActual}, Egresos Totales = ${totalEgresos}, Ingresos Brutos = ${totalIngresos}`);

  return { totalIngresos };
}

/**
 * Recalcula el margen de contribución.
 * Resta los costos directos de los ingresos brutos.
 */
function recalculateMargins(totalIngresos) {
  Logger.log('Recalculando márgenes...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Obtener la hoja de Costos y Márgenes
  const marginsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.COSTOS_MARGENES);
  if (!marginsSheet) {
    SpreadsheetApp.getUi().alert('Error Crítico: No se encontró la hoja "12. Costos y Márgenes".');
    return;
  }

  // 2. Sumar los costos directos
  const costsValues = marginsSheet.getRange(CONFIG.DATA_RANGES.COSTS_SOURCE_RANGE).getValues();
  const totalCosts = costsValues.reduce((sum, row) => sum + (parseFloat(row[0]) || 0), 0);

  // 3. Calcular el margen de contribución y el porcentaje
  const contributionMargin = totalIngresos - totalCosts;
  const marginPercentage = (totalIngresos > 0) ? (contributionMargin / totalIngresos) : 0;

  // 4. Escribir los resultados en la hoja
  marginsSheet.getRange(CONFIG.DATA_RANGES.MARGIN_RESULT_CELL).setValue(contributionMargin);
  marginsSheet.getRange(CONFIG.DATA_RANGES.MARGIN_PERCENT_CELL).setValue(marginPercentage);

  Logger.log(`Márgenes calculados: Margen de Contribución = ${contributionMargin}, Porcentaje = ${marginPercentage}`);
}

/**
 * Ejecuta una simulación basada en parámetros.
 * Placeholder para la lógica detallada.
 */
function runSimulation(parameters) {
  // Lógica para el Simulador de Decisiones (Módulo 18).
  Logger.log('Ejecutando simulación con parámetros:', parameters);
}

/**
 * Genera y muestra recomendaciones automáticas en el Dashboard.
 * Basado en los KPIs y umbrales definidos en la configuración.
 */
function generateDashboardRecommendations() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);

  if (!dashboardSheet) {
    SpreadsheetApp.getUi().alert('Error Crítico: No se encontró la hoja de "Dashboard". Las recomendaciones no pueden ser generadas.');
    Logger.log('Error: No se encontró el Dashboard para generar recomendaciones.');
    return;
  }

  // CORRECCIÓN: Usar el saldo de caja para la recomendación de cobertura, no los ingresos brutos.
  const saldoCaja = dashboardSheet.getRange(CONFIG.DATA_RANGES.DASHBOARD_CAJA_HOY_CELL).getValue();
  const totalEgresos = dashboardSheet.getRange(CONFIG.DATA_RANGES.DASHBOARD_TOTAL_EGRESOS_CELL).getValue();

  let recommendations = [];

  // Lógica de Recomendación 1: Cobertura de Caja
  if (totalEgresos > 0) {
    const cashCoverage = saldoCaja / totalEgresos;
    const threshold = CONFIG.DASHBOARD_SETTINGS.RECOMMENDATION_THRESHOLDS.LOW_CASH_COVERAGE;
    if (cashCoverage < threshold) {
      recommendations.push(`- Tu caja actual cubre solo ${cashCoverage.toFixed(1)} meses de gastos. Considera reducir gastos variables.`);
    }
  }

  // (Aquí se podrían añadir más lógicas de recomendación en el futuro)
  // Ej. "Margen por debajo del 15%. Revisa costos en Hoja 12."
  // Ej. "Cobranza pendiente supera el 20% de la facturación. Acciona Hoja 05."

  const recommendationText = recommendations.length > 0 ? recommendations.join('\n') : 'Todo parece estar en orden. ¡Buen trabajo!';

  dashboardSheet.getRange(CONFIG.DASHBOARD_SETTINGS.RECOMMENDATIONS_CELL).setValue(recommendationText);

  Logger.log('Recomendaciones del dashboard actualizadas.');
}
