/**
 * @fileoverview Módulo de Experiencia de Usuario (UX) para el Business OS.
 *
 * Descripción:
 * Este archivo contiene funciones destinadas a mejorar la interacción del
 * usuario con la hoja de cálculo. Esto incluye mostrar mensajes claros,
 * toasts, pop-ups, indicadores visuales (como cambiar el color de una celda
 * basado en su valor) y tooltips. El objetivo es que el sistema se sienta
 * más como una aplicación y menos como una simple hoja de cálculo.
 *
 * @author ECD OS
 * @version 1.1
 */

/**
 * Actualiza el "Semáforo de Riesgo" en el Dashboard Ejecutivo.
 * Cambia el color de fondo de la celda designada basándose en la relación
 * entre ingresos y egresos.
 */
function updateRiskSemaphore() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboardSheet = ss.getSheetByName(CONFIG.SHEET_NAMES.DASHBOARD);

  if (!dashboardSheet) {
    SpreadsheetApp.getUi().alert('Error Crítico: No se encontró la hoja de "Dashboard". El semáforo de riesgo no puede ser actualizado.');
    Logger.log('Error: No se encontró el Dashboard para actualizar el semáforo de riesgo.');
    return;
  }

  // CORRECCIÓN: Usar los ingresos brutos para el ratio, no el saldo de caja.
  const ingresosBrutos = dashboardSheet.getRange(CONFIG.DATA_RANGES.DASHBOARD_INGRESOS_BRUTOS_CELL).getValue();
  const totalEgresos = dashboardSheet.getRange(CONFIG.DATA_RANGES.DASHBOARD_TOTAL_EGRESOS_CELL).getValue();

  const semaphoreCell = dashboardSheet.getRange(CONFIG.DASHBOARD_SETTINGS.RISK_SEMAPHORE_CELL);

  if (totalEgresos <= 0) {
    semaphoreCell.setBackground('#d3d3d3').setValue('N/A'); // Gris si no hay egresos
    return;
  }

  const ratio = ingresosBrutos / totalEgresos;
  const thresholds = CONFIG.DASHBOARD_SETTINGS.SEMAPHORE_THRESHOLDS;

  if (ratio > thresholds.GREEN) {
    semaphoreCell.setBackground('#90ee90').setValue('Saludable'); // Verde
  } else if (ratio > thresholds.YELLOW) {
    semaphoreCell.setBackground('#fffacd').setValue('Precaución'); // Amarillo
  } else {
    semaphoreCell.setBackground('#f08080').setValue('Riesgo'); // Rojo
  }

  Logger.log(`Semáforo de riesgo actualizado. Ratio: ${ratio}`);
}
