/**
 * @fileoverview Módulo de Seguridad y Control para el Business OS.
 *
 * Descripción:
 * Este archivo se encarga de la protección de la hoja de cálculo.
 * Incluye funciones para proteger rangos y hojas enteras, gestionar
 * roles de usuario (dueño vs. operador) y realizar auditorías básicas
 * de cambios para mantener la integridad de los datos.
 *
 * @author ECD OS
 * @version 1.1
 */

/**
 * Aplica protecciones a los rangos basadas en los roles definidos en Config.gs.
 * Esta función asegura que solo los 'OWNERS' puedan editar los rangos protegidos,
 * mientras que los 'OPERATORS' son eliminados de esos permisos.
 */
function applyRoleBasedProtections() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const protections = ss.getProtections(SpreadsheetApp.ProtectionType.RANGE);

  const owners = CONFIG.USER_ROLES.OWNERS;
  const operators = CONFIG.USER_ROLES.OPERATORS;

  protections.forEach(protection => {
    // Asegurar que solo los 'owners' tengan permiso de edición
    protection.getEditors().forEach(editor => {
      if (owners.indexOf(editor.getEmail()) === -1) {
        protection.removeEditor(editor);
      }
    });

    // Añadir todos los 'owners' a la protección
    protection.addEditors(owners);

    // Remover explícitamente a los 'operators' de los rangos protegidos
    operators.forEach(operatorEmail => {
      protection.removeEditor(operatorEmail);
    });
  });

  Logger.log('Protecciones basadas en roles aplicadas correctamente.');
}
