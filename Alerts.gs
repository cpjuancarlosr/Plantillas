/**
 * @fileoverview Motor de Alertas para el Business OS.
 *
 * Descripción:
 * Este archivo contiene toda la lógica para monitorear el estado del negocio
 * y generar alertas proactivas. Las funciones aquí implementadas se conectan
 * con los parámetros definidos en `Config.gs` para detectar riesgos.
 *
 * @author ECD OS
 * @version 1.1
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
 * Compara la caja proyectada con los gastos fijos del mes.
 */
function detectCashflowRisk() {
  Logger.log('Verificando riesgo de flujo de caja...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);
  const configSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.CONFIGURACION);

  if (!dashboardSheet || !configSheet) {
    SpreadsheetApp.getUi().alert('Error Crítico: No se encontró la hoja de "Dashboard" o "Configuración". Las alertas no pueden ser evaluadas.');
    Logger.log('Error: No se encontró el Dashboard o la hoja de Configuración.');
    return;
  }

  const cajaProyectada = dashboardSheet.getRange(CONFIG.DATA_RANGES.DASHBOARD_CAJA_PROYECTADA_CELL).getValue();
  const gastosFijosValues = configSheet.getRange(CONFIG.DATA_RANGES.GASTOS_FIJOS_SOURCE_RANGE).getValues();
  const totalGastosFijos = gastosFijosValues.reduce((sum, row) => sum + (parseFloat(row[0]) || 0), 0);

  // Si no hay gastos fijos, no se puede evaluar este riesgo.
  if (totalGastosFijos <= 0) {
    Logger.log('No hay gastos fijos definidos para calcular el riesgo de caja.');
    return;
  }

  const threshold = CONFIG.RISK_THRESHOLDS.CASH_COVERAGE_MONTHS;

  if (cajaProyectada < (totalGastosFijos * threshold)) {
    const subject = 'Acción Requerida: Riesgo de Liquidez Detectado';
    const body = `Hola,\n\nSe ha detectado un posible riesgo de liquidez en el sistema.\n\n` +
                 `Caja proyectada a 30 días: ${cajaProyectada.toFixed(2)}\n` +
                 `Gastos Fijos del Mes: ${totalGastosFijos.toFixed(2)}\n` +
                 `La caja proyectada cubre menos de ${threshold} meses de gastos fijos.\n\n` +
                 `Recomendación: Revisa tus gastos variables y prioriza la cobranza pendiente.`;

    sendAlert(subject, body);
  } else {
    Logger.log(`Salud de caja estable. Caja proyectada (${cajaProyectada.toFixed(2)}) cubre los gastos fijos (${totalGastosFijos.toFixed(2)}).`);
  }
}

/**
 * Detecta la proximidad de vencimientos fiscales importantes.
 * Revisa las fechas en `CONFIG.FISCAL_DATES`.
 */
function detectTaxDeadlines() {
  Logger.log('Verificando vencimientos fiscales...');
  const today = new Date();
  const reminderDays = CONFIG.RISK_THRESHOLDS.TAX_REMINDER_DAYS;

  for (const key in CONFIG.FISCAL_DATES) {
    const dateStr = CONFIG.FISCAL_DATES[key]; // Formato 'MM-DD'
    const [month, day] = dateStr.split('-').map(Number);

    // Crear una fecha para el vencimiento de este año
    const deadline = new Date(today.getFullYear(), month - 1, day);

    // Calcular la diferencia en días
    const timeDiff = deadline.getTime() - today.getTime();
    const daysUntilDeadline = Math.ceil(timeDiff / (1000 * 3600 * 24));

    // Enviar alerta si la fecha está dentro del umbral de recordatorio
    if (daysUntilDeadline > 0 && daysUntilDeadline <= reminderDays) {
      const subject = `Recordatorio: Vencimiento Fiscal Próximo`;
      const body = `Hola,\n\nEste es un recordatorio de que la fecha para "${key}" se acerca.\n\n` +
                   `Fecha de Vencimiento: ${day}/${month}/${today.getFullYear()}\n` +
                   `Días restantes: ${daysUntilDeadline}\n\n` +
                   `Por favor, asegúrate de tener todo en orden.`;

      sendAlert(subject, body);
    }
  }
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
  Logger.log('Verificando dependencia de clientes...');
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Obtener la hoja de Ranking de Clientes y los ingresos brutos del Dashboard
  const clientsSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.RANKING_CLIENTES);
  const dashboardSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);

  if (!clientsSheet || !dashboardSheet) {
    Logger.log('No se encontraron las hojas necesarias para el análisis de dependencia de clientes.');
    return;
  }

  const totalIngresos = dashboardSheet.getRange(CONFIG.DATA_RANGES.DASHBOARD_INGRESOS_BRUTOS_CELL).getValue();

  if (totalIngresos <= 0) {
    Logger.log('No hay ingresos registrados para analizar la dependencia de clientes.');
    return;
  }

  // 2. Leer los datos de ingresos por cliente
  const clientData = clientsSheet.getRange(CONFIG.DATA_RANGES.CLIENT_REVENUE_SOURCE_RANGE).getValues();
  const dependencyThreshold = CONFIG.RISK_THRESHOLDS.CLIENT_DEPENDENCY_PERCENTAGE;

  // 3. Iterar sobre cada cliente y verificar su concentración de ingresos
  clientData.forEach(row => {
    const clientName = row[0];
    const clientRevenue = parseFloat(row[1]);

    if (clientName && clientRevenue > 0) {
      const percentage = clientRevenue / totalIngresos;
      if (percentage > dependencyThreshold) {
        const subject = 'Alerta: Alta Dependencia de Cliente Detectada';
        const body = `Hola,\n\nSe ha detectado una alta dependencia en el cliente "${clientName}".\n\n` +
                     `Ingresos del cliente: ${clientRevenue.toFixed(2)}\n` +
                     `Ingresos totales: ${totalIngresos.toFixed(2)}\n` +
                     `Este cliente representa un ${ (percentage * 100).toFixed(1) }% de la facturación total, superando el umbral de riesgo del ${(dependencyThreshold * 100)}%.\n\n` +
                     `Recomendación: Considera estrategias para diversificar tu cartera de clientes.`;

        sendAlert(subject, body);
      }
    }
  });
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
