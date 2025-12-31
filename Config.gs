/**
 * @fileoverview Archivo de Configuración Central para el Business OS.
 *
 * Descripción:
 * Este archivo contiene todas las variables, constantes y configuraciones
 * que controlan el comportamiento del sistema. El objetivo es centralizar
 * la gestión para que cualquier ajuste se realice en un único lugar,
 * facilitando el mantenimiento y la duplicación para nuevos clientes.
 *
 * @author ECD OS
 * @version 1.1
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
        GASTOS_FIJOS_SOURCE_RANGE: '30. Configuración!D5:D50', // Rango para la lista de gastos fijos
        DASHBOARD_CAJA_HOY_CELL: 'C2', // Celda para el saldo de caja actual (Ingresos - Egresos)
        DASHBOARD_TOTAL_EGRESOS_CELL: 'C3',
        DASHBOARD_CAJA_PROYECTADA_CELL: 'C4',
        DASHBOARD_INGRESOS_BRUTOS_CELL: 'D2', // Celda para los ingresos brutos del período
        IMPUESTOS_TOTAL_CELL: 'C5', // Celda para el total de impuestos en la hoja de Impuestos
        COSTS_SOURCE_RANGE: '12. Costos y Márgenes!B5:B50', // Rango para la lista de costos directos
        MARGIN_RESULT_CELL: '12. Costos y Márgenes!D5',     // Celda para escribir el margen de contribución
        MARGIN_PERCENT_CELL: '12. Costos y Márgenes!E5',    // Celda para escribir el % de margen
        CLIENT_REVENUE_SOURCE_RANGE: '14. Ranking de Clientes!B5:C50' // Rango para leer Nombres de Cliente y sus Ingresos
    },

    // 5. Configuración del Dashboard Dinámico
    DASHBOARD_SETTINGS: {
        RISK_SEMAPHORE_CELL: 'E8', // Celda para el "Semáforo de Riesgo"
        RECOMMENDATIONS_CELL: 'B10', // Celda para las "Recomendaciones Automáticas"
        SEMAPHORE_THRESHOLDS: {
            GREEN: 2.0,  // Ratio (Ingresos/Egresos) > 2.0 es Verde
            YELLOW: 1.0  // Ratio > 1.0 y <= 2.0 es Amarillo
        },
        RECOMMENDATION_THRESHOLDS: {
            LOW_CASH_COVERAGE: 1.5 // Meses de cobertura de caja
        }
    },

    // 6. Configuración de Impuestos
    TAX_SETTINGS: {
        GENERAL_TAX_RATE: 0.16 // Tasa de impuesto general (ej. 16% IVA)
    },

    // 6. Parámetros de Alertas y Riesgos (Risk & Alert Parameters)
    // Umbrales para disparar notificaciones.
    RISK_THRESHOLDS: {
        // Si la caja proyectada a 30 días es menor que 1.5 veces los gastos fijos mensuales, se considera riesgo.
        CASH_COVERAGE_MONTHS: 1.5,
        // Si un solo cliente representa más del 25% de los ingresos totales, generar una alerta.
        CLIENT_DEPENDENCY_PERCENTAGE: 0.25,
        // Días antes del vencimiento de un impuesto para enviar un recordatorio.
        TAX_REMINDER_DAYS: 7
    },

    // 7. Configuración de Correo Electrónico (Email Settings)
    // Direcciones para enviar alertas automáticas.
    ALERT_EMAILS: {
        // Correo del dueño o tomador de decisiones principal.
        OWNER: 'correo@dueno.com',
        // Correo del contador o responsable financiero.
        ACCOUNTANT: 'correo@contador.com'
    },

    // 8. Fechas Fiscales Clave (Key Fiscal Dates)
    // Usado para recordatorios y cálculos. El formato es 'MM-DD'.
    FISCAL_DATES: {
        DECLARACION_MENSUAL: '03-20', // Día 20 de cada mes
        DECLARACION_ANUAL: '04-30'
    },

    // 9. Roles de Usuario (User Roles)
    // Define permisos básicos y las listas de correos para cada rol.
    USER_ROLES: {
        OWNERS: ['correo@dueno.com'], // Lista de correos con acceso total
        OPERATORS: ['correo@operador.com'], // Lista de correos con acceso limitado a inputs
        // Se pueden añadir más roles si es necesario
    },

    // 10. Gestión de Períodos (Period Management)
    // Define qué hojas se deben duplicar al crear un nuevo mes.
    PERIOD_MANAGEMENT: {
        SHEETS_TO_DUPLICATE: [
            '03. Ingresos',
            '04. Egresos'
        ]
    },

    // 12. Validación de Datos (Data Validation)
    VALIDATION_RANGES: {
        CATEGORIAS_SOURCE: '30. Configuración!B5:B50', // Rango que contiene la lista de categorías
        INGRESOS_CATEGORIA_TARGET: 'C5:C100',         // Rango para aplicar el desplegable en Ingresos
        EGRESOS_CATEGORIA_TARGET: 'C5:C100'           // Rango para aplicar el desplegable en Egresos
    },

    // 13. Configuración del Período Activo
    ACTIVE_PERIOD_CONFIG: {
        ACTIVE_INCOME_SHEET_CELL: '30. Configuración!F5', // Celda que almacena el nombre de la hoja de ingresos activa
        ACTIVE_EXPENSE_SHEET_CELL: '30. Configuración!F6'  // Celda que almacena el nombre de la hoja de egresos activa
    },

    // 14. Estructura y Reglas de Inputs
    INPUT_STRUCTURE: {
        DATE_COLUMN: 2,         // Número de la columna de Fecha (B)
        AMOUNT_COLUMN: 6,       // Número de la columna de Monto (F)
        REQUIRED_COLUMNS: [2, 3, 6], // Columnas obligatorias (Fecha, Categoría, Monto)
        INCOMPLETE_ROW_COLOR: '#fff8e1', // Color para resaltar filas incompletas (amarillo claro)
        START_ROW: 5            // Fila donde comienzan los datos
    }
};

// --- FIN DEL ARCHIVO DE CONFIGURACIÓN ---
