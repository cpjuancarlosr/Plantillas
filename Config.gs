/**
 * @fileoverview Archivo de Configuración Central para el Business OS.
 *
 * Descripción:
 * Este archivo contiene todas las variables, constantes y configuraciones
 * que controlan el comportamiento del sistema. El objetivo es centralizar
 * la gestión para que cualquier ajuste se realice en un único lugar,
 * facilitando el mantenimiento y la duplicación para nuevos clientes.
 *
 * @author Tu Nombre/Empresa
 * @version 1.0
 */

// Objeto principal de configuración para encapsular todas las variables.
const CONFIG = {

    // 1. Nombres de las Hojas de Cálculo (Sheet Names)
    // Se utilizan para referenciar hojas sin 'hardcodear' los nombres en el código.
    SHEET_NAMES: {
        DASHBOARD: '00. Dashboard Ejecutivo',
        FLUJO_CAJA: '01. Flujo de Caja',
        PROYECCION_CAJA: '02. Proyección de Caja',
        INGRESOS: '03. Ingresos',
        EGRESOS: '04. Egresos',
        COBRANZA: '05. Cuentas por Cobrar',
        IMPUESTOS: '08. Impuestos',
        PUNTO_EQUILIBRIO: '11. Punto de Equilibrio',
        COSTOS_MARGENES: '12. Costos y Márgenes',
        RANKING_CLIENTES: '14. Ranking de Clientes',
        SIMULADOR: '18. Simulador de Decisiones',
        RIESGOS: '21. Matriz de Riesgos',
        PROVEEDORES: '24. Control de Proveedores',
        CHECKLIST_CIERRE: '27. Checklist de Cierre',
        HISTORIAL: '28. Historial de Decisiones',
        CONFIGURACION: '30. Configuración'
    },

    // 2. Rangos de Celdas Editables (Input Ranges)
    // Define las áreas donde el usuario puede introducir datos.
    // La clave DEBE coincidir con el valor en SHEET_NAMES.
    EDITABLE_RANGES: {
        '03. Ingresos': 'B5:F100',
        '04. Egresos': 'B5:F100'
        // ... agregar otros rangos clave
    },

    // 3. Rangos Protegidos (Protected Ranges)
    // Celdas o rangos que deben ser bloqueados para evitar ediciones accidentales.
    PROTECTED_RANGES: {
        DASHBOARD_KPIS: '00. Dashboard Ejecutivo!B2:D5',
        FORMULAS_IMPUESTOS: '08. Impuestos!C:E',
        // ... agregar otros rangos con fórmulas
    },

    // 4. Rangos de Datos Específicos (Data Ranges)
    // Define dónde se encuentran los datos clave para los cálculos.
    DATA_RANGES: {
        INGRESOS_MONTO_COL: 'F5:F100',
        EGRESOS_MONTO_COL: 'F5:F100',
        DASHBOARD_TOTAL_INGRESOS_CELL: 'C2',
        DASHBOARD_TOTAL_EGRESOS_CELL: 'C3'
    },

    // 5. Parámetros de Alertas y Riesgos (Risk & Alert Parameters)
    // Umbrales para disparar notificaciones.
    RISK_THRESHOLDS: {
        // Si la caja proyectada a 30 días es menor que 1.5 veces los gastos fijos mensuales, se considera riesgo.
        CASH_COVERAGE_MONTHS: 1.5,
        // Si un solo cliente representa más del 25% de los ingresos totales, generar una alerta.
        CLIENT_DEPENDENCY_PERCENTAGE: 0.25,
        // Días antes del vencimiento de un impuesto para enviar un recordatorio.
        TAX_REMINDER_DAYS: 7
    },

    // 6. Configuración de Correo Electrónico (Email Settings)
    // Direcciones para enviar alertas automáticas.
    ALERT_EMAILS: {
        // Correo del dueño o tomador de decisiones principal.
        OWNER: 'correo@dueno.com',
        // Correo del contador o responsable financiero.
        ACCOUNTANT: 'correo@contador.com'
    },

    // 7. Fechas Fiscales Clave (Key Fiscal Dates)
    // Usado para recordatorios y cálculos. El formato es 'MM-DD'.
    FISCAL_DATES: {
        DECLARACION_MENSUAL: '03-20', // Día 20 de cada mes
        DECLARACION_ANUAL: '04-30'
    },

    // 8. Roles de Usuario (User Roles)
    // Define permisos básicos.
    USER_ROLES: {
        OWNER: 'owner',       // Acceso total, puede cambiar configuración.
        OPERATOR: 'operator'  // Acceso a inputs, pero no a configuración ni reportes sensibles.
    },

    // 9. Gestión de Períodos (Period Management)
    // Define qué hojas se deben duplicar al crear un nuevo mes.
    PERIOD_MANAGEMENT: {
        SHEETS_TO_DUPLICATE: [
            '03. Ingresos',
            '04. Egresos'
        ]
    }
};

// --- FIN DEL ARCHIVO DE CONFIGURACIÓN ---
