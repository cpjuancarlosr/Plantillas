/**
 * @fileoverview Motor de Alertas para el Business OS.
 *
 * Descripción:
 * Este archivo contiene toda la lógica para monitorear el estado del negocio
 * y generar alertas proactivas. Las funciones aquí implementadas se conectan
 * con los parámetros definidos en `Config.gs` para detectar riesgos.
 *
 * @author Tu Nombre/Empresa
 * @version 1.0
 */

/**
 * Función principal que orquesta la verificación de todas las alertas.
 * Puede ser llamada después de un recálculo o mediante un trigger.
 */
function checkAllAlerts() {
  Logger.log('Iniciando verificación de alertas...');

  detectCashflowRisk();
  detectTaxDeadlines();
  detectBudgetDeviations();
  detectClientDependency();

  Logger.log('Verificación de alertas completada.');
}

/**
 * Detecta si existe un riesgo de liquidez a corto plazo.
 * Lee los ingresos y egresos del dashboard y los compara.
 */
function detectCashflowRisk() {
  Logger.log('Verificando riesgo de flujo de caja...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);

  if (!dashboardSheet) {
    Logger.log('Error: No se encontró el Dashboard para la verificación de alertas.');
    return;
  }

  // Leer los valores que `recalculateKPIs` acaba de escribir, usando rangos de CONFIG
  const totalIngresos = dashboardSheet.getRange(CONFIG.DATA_RANGES.DASHBOARD_TOTAL_INGRESOS_CELL).getValue();
  const totalEgresos = dashboardSheet.getRange(CONFIG.DATA_RANGES.DASHBOARD_TOTAL_EGRESOS_CELL).getValue();

  // Asegurarse de que los egresos no sean cero para evitar división por cero
  if (totalEgresos <= 0) {
    Logger.log('No hay egresos registrados, no se puede calcular el riesgo de caja.');
    return;
  }

  // Calcular la cobertura de caja (cuántos "meses" de gastos cubren los ingresos)
  const cashCoverage = totalIngresos / totalEgresos;
  const threshold = CONFIG.RISK_THRESHOLDS.CASH_COVERAGE_MONTHS;

  if (cashCoverage < threshold) {
    const subject = 'Acción Requerida: Riesgo de Liquidez Detectado';
    const body = `Hola,\n\nSe ha detectado un posible riesgo de liquidez en el sistema.\n\n` +
                 `Ingresos del período: ${totalIngresos.toFixed(2)}\n` +
                 `Egresos del período: ${totalEgresos.toFixed(2)}\n` +
                 `Cobertura actual: ${cashCoverage.toFixed(2)} meses\n` +
                 `Umbral de riesgo: ${threshold} meses\n\n` +
                 `Recomendación: Revisa tus gastos variables y prioriza la cobranza pendiente.`;

    sendAlert(subject, body);
  } else {
    Logger.log(`Salud de caja estable. Cobertura: ${cashCoverage.toFixed(2)}, Umbral: ${threshold}`);
  }
}

/**
 * Detecta la proximidad de vencimientos fiscales importantes.
 * Revisa las fechas en `CONFIG.FISCAL_DATES`.
 */
function detectTaxDeadlines() {
  // Lógica para comprobar la fecha actual contra las fechas fiscales
  // Si una fecha está cerca (ej. `CONFIG.RISK_THRESHOLDS.TAX_REMINDER_DAYS`),
  // llamar a `sendAlert()`.
  Logger.log('Verificando vencimientos fiscales...');
}

/**
 * Detecta si los gastos reales se están desviando significativamente del presupuesto.
 * (Esta función requiere una hoja de presupuestos no definida en el plan original).
 */
function detectBudgetDeviations() {
  // Lógica para comparar gastos reales con presupuestados.
  Logger.log('Verificando desviaciones de presupuesto...');
}

/**
 * Detecta si existe una dependencia excesiva de un solo cliente.
 */
function detectClientDependency() {
  // Lógica para analizar la hoja de Ranking de Clientes
  // Si un cliente supera `CONFIG.RISK_THRESHOLDS.CLIENT_DEPENDENCY_PERCENTAGE`,
  // llamar a `sendAlert()`.
  Logger.log('Verificando dependencia de clientes...');
}


/**
 * Envía una alerta por correo electrónico.
 * @param {string} subject El asunto del correo.
 * @param {string} body El cuerpo del correo.
 */
function sendAlert(subject, body) {
  const recipient = CONFIG.ALERT_EMAILS.OWNER;
  if (recipient) {
    MailApp.sendEmail(recipient, `[ECD OS] ${subject}`, body);
    Logger.log(`Alerta enviada a ${recipient}: ${subject}`);
  }
}
