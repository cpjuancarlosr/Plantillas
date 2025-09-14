/**
 * //== File: seed.gs
 * Funciones para poblar las plantillas con datos de ejemplo realistas y para limpiarlos.
 */

function seedAllTemplates_() {
    const ss = getActiveSpreadsheet_();
    SpreadsheetApp.getActiveSpreadsheet().toast('Poblando plantillas con datos de ejemplo...', 'Progreso', 15);

    // El orden es importante para las dependencias
    seedCatalogoCuentasData_(ss);
    seedPolizasData_(ss); // Esto alimenta la Balanza, ER y BG

    seedKanbanData_(ss);
    seedCrmData_(ss);
    seedInventoryData_(ss);
    seedSalesData_(ss); // Usa datos que se reflejan en Pólizas
    seedPurchasingData_(ss);
    seedPayrollData_(ss);
    seedActivosFijosData_(ss);
    seedOkrData_(ss);
    seedCashflowData_(ss);
    seedCuentasPagarCobrarData_(ss);

    logAction_('Datos de ejemplo cargados en la mayoría de las plantillas.');
    SpreadsheetApp.getUi().alert('Se han cargado los datos de ejemplo.');
}

function clearAllTemplateData_() {
    const ss = getActiveSpreadsheet_();
    // Lista expandida de rangos a limpiar
    const sheetsToClear = [
        { name: 'A01_Proyectos_Kanban', range: 'A2:J100' },
        { name: 'A02_CRM_Pipeline', range: 'A2:I100' },
        { name: 'A03_Inventarios', range: 'A2:G100,I2:M100' },
        { name: 'A04_Compras_Proveedores', range: 'A2:I100' },
        { name: 'A05_Ventas_Clientes', range: 'A2:J100' },
        { name: 'A06_Facturacion_SAT', range: 'A2:I100' },
        { name: 'A07_Nomina_Simple', range: 'A2:H100' },
        { name: 'A08_Presupuesto_Maestro', range: 'A5:L100' },
        { name: 'A09_Flujo_Efectivo_Op', range: 'D3,A6:E100' },
        { name: 'A10_OKR_Trimestral', range: 'A4:F100' },
        { name: 'B11_Catalogo_Cuentas', range: 'A2:G500' },
        { name: 'B12_Polizas_Diario', range: 'A2:G1000' },
        { name: 'B17_Conciliacion_Bancaria', range: 'A3:E100,G3:K100' },
        { name: 'B18_IVA_Control', range: 'B2:E100,H2:K100' },
        { name: 'B19_ISR_PM_Provisional', range: 'B4,B6:B8' },
        { name: 'B21_Activos_Fijos_Dep', range: 'A2:H100' },
        { name: 'B22_Retenciones_Terceros', range: 'A2:G100' },
        { name: 'C28_Cuentas_Por_Pagar', range: 'A2:G100' },
        { name: 'C29_Cuentas_Por_Cobrar', range: 'A2:G100' }
    ];

    sheetsToClear.forEach(s => {
        const sheet = ss.getSheetByName(s.name);
        if (sheet) {
            sheet.getRange(s.range).clearContent();
        }
    });
    logAction_('Datos de ejemplo eliminados de todas las plantillas.');
}

// --- FUNCIONES DE SEEDING DETALLADAS ---

function seedCatalogoCuentasData_(ss) {
    const s = ss.getSheetByName('B11_Catalogo_Cuentas');
    if (!s) return;
    const data = [
      ['1010', 'Caja y Bancos', 1, '1010.01', 'Caja General', 'D', 'Activo'],
      ['1020', 'Bancos', 1, '1020.01', 'Banco Nacional MXN', 'D', 'Activo'],
      ['1050', 'Clientes', 1, '1050.01', 'Clientes Nacionales', 'D', 'Activo'],
      ['1070', 'Inventarios', 1, '1070.01', 'Almacén de Mercancías', 'D', 'Activo'],
      ['1500', 'Activo Fijo', 1, '1500.01', 'Equipo de Cómputo', 'D', 'Activo'],
      ['2010', 'Proveedores', 1, '2010.01', 'Proveedores Nacionales', 'A', 'Pasivo'],
      ['2070', 'Impuestos por Pagar', 1, '2070.01', 'IVA por Pagar', 'A', 'Pasivo'],
      ['3010', 'Capital Social', 1, '3010.01', 'Capital Social Fijo', 'A', 'Capital'],
      ['4010', 'Ventas', 1, '4010.01', 'Ventas de Productos', 'A', 'Ingresos'],
      ['5010', 'Costo de Ventas', 1, '5010.01', 'Costo de la Mercancía Vendida', 'D', 'Costos'],
      ['6010', 'Gastos de Venta', 1, '6010.01', 'Sueldos y Salarios Ventas', 'D', 'Gastos'],
      ['6020', 'Gastos de Administración', 1, '6020.01', 'Renta de Oficina', 'D', 'Gastos']
    ];
    s.getRange('A2:G13').setValues(data);
}

function seedPolizasData_(ss) {
    const s = ss.getSheetByName('B12_Polizas_Diario');
    if (!s) return;
    const data = [
        // Venta de 2 laptops
        [new Date(), 'Ingreso', 'I-1', '1050.01', 'Venta a Cliente X (2 Laptops)', 46400, 0],
        [new Date(), 'Ingreso', 'I-1', '4010.01', 'Venta a Cliente X (2 Laptops)', 0, 40000],
        [new Date(), 'Ingreso', 'I-1', '2070.01', 'IVA Trasladado Venta', 0, 6400],
        // Costo de esa venta
        [new Date(), 'Diario', 'D-1', '5010.01', 'Costo Venta (2 Laptops)', 37000, 0],
        [new Date(), 'Diario', 'D-1', '1070.01', 'Salida de Almacén', 0, 37000],
        // Pago de nómina
        [new Date(), 'Egreso', 'E-1', '6010.01', 'Pago de Nómina Ventas', 15000, 0],
        [new Date(), 'Egreso', 'E-1', '1020.01', 'Salida de Banco por Nómina', 0, 15000],
        // Pago de Renta
        [new Date(), 'Egreso', 'E-2', '6020.01', 'Pago de Renta Oficina', 12000, 0],
        [new Date(), 'Egreso', 'E-2', '1020.01', 'Salida de Banco por Renta', 0, 12000],
        // Cobro a cliente
        [new Date(), 'Ingreso', 'I-2', '1020.01', 'Cobro a Cliente X', 46400, 0],
        [new Date(), 'Ingreso', 'I-2', '1050.01', 'Aplicación Cobro', 0, 46400]
    ];
    s.getRange('A2:G12').setValues(data);
}

function seedKanbanData_(ss) {
    const sheet = ss.getSheetByName('A01_Proyectos_Kanban');
    if (!sheet) return;
    const data = [
        ['T-01', 'Desarrollar Landing Page', 'Proyecto Web', 'Ana', 'En Progreso', 'Alta', new Date(2023, 9, 1), new Date(2023, 9, 15), '', 0.75],
        ['T-02', 'Diseño de Mockups', 'Proyecto Web', 'Luis', 'Completado', 'Alta', new Date(2023, 8, 25), new Date(2023, 8, 30), '', 1],
        ['T-03', 'Campaña Marketing Q4', 'Marketing', 'Sara', 'Pendiente', 'Media', new Date(2023, 9, 10), new Date(2023, 11, 20), '', 0.1]
    ];
    sheet.getRange('A2:J4').setValues(data);
}

function seedCrmData_(ss) {
    const sheet = ss.getSheetByName('A02_CRM_Pipeline');
    if (!sheet) return;
    const data = [
        ['L-001', 'Juan Pérez', 'Empresa ABC', 'Propuesta', 50000, 0.6, '', new Date(2023, 9, 2), new Date(2023, 9, 10)],
        ['L-002', 'María García', 'Constructora XYZ', 'Ganado', 120000, 1, '', new Date(2023, 8, 15), ''],
        ['L-003', 'Carlos López', 'Startup Tech', 'Negociación', 75000, 0.8, '', new Date(2023, 9, 1), new Date(2023, 9, 8)]
    ];
    sheet.getRange('A2:I4').setValues(data);
}

function seedInventoryData_(ss) {
    const sheet = ss.getSheetByName('A03_Inventarios');
    if (!sheet) return;
    const products = [['P-101', 'Laptop Gamer X', '', '', 10], ['P-205', 'Monitor Curvo 27"', '', '', 15]];
    sheet.getRange('I2:M3').setValues(products);
    const moves = [
        [new Date(2023, 9, 1), 'Entrada', 'P-101', 'Laptop Gamer X', 20, 18500, ''],
        [new Date(2023, 9, 2), 'Entrada', 'P-205', 'Monitor Curvo 27"', 30, 4200, ''],
        [new Date(2023, 9, 5), 'Salida', 'P-101', 'Laptop Gamer X', 2, 0, ''] // Corresponde a la venta en pólizas
    ];
    sheet.getRange('A2:G4').setValues(moves);
}

function seedSalesData_(ss) {
    const s = ss.getSheetByName('A05_Ventas_Clientes');
    if (!s) return;
    const data = [
        ['V-001', 'Cliente X', new Date(), 'P-101', 2, 23200, 18500, '', '', ''],
        ['V-002', 'Cliente Z', new Date(), 'P-205', 5, 5500, 4200, '', '', '']
    ];
    s.getRange('A2:J3').setValues(data);
}

function seedPayrollData_(ss) {
    const s = ss.getSheetByName('A07_Nomina_Simple');
    if (!s) return;
    const data = [
        ['E-01', 'Ana López', 'Gerente de Ventas', 35000, 5000, 8500, '', ''],
        ['E-02', 'Luis Torres', 'Vendedor', 15000, 2000, 3500, '', '']
    ];
    s.getRange('A2:H3').setValues(data);
}

function seedActivosFijosData_(ss) {
    const s = ss.getSheetByName('B21_Activos_Fijos_Dep');
    if (!s) return;
    const data = [
        ['AF-01', 'Laptop Gerencia', new Date(2023, 1, 1), 45000, 0.30, '', '', ''],
        ['AF-02', 'Servidor Dell', new Date(2023, 1, 1), 80000, 0.25, '', '', '']
    ];
    s.getRange('A2:H3').setValues(data);
}

function seedOkrData_(ss) {
    const s = ss.getSheetByName('A10_OKR_Trimestral');
    if (!s) return;
    const data = [
      ['Aumentar Rentabilidad', 'Incrementar margen bruto de 25% a 30%', 'Finanzas', 0.25, 0.30, 0.28, '', ''],
      ['Mejorar Satisfacción Cliente', 'Aumentar NPS de 40 a 55', 'Ventas', 40, 55, 51, '', '']
    ];
    s.getRange('A4:H5').setValues(data);
}

function seedCashflowData_(ss) {
    const s = ss.getSheetByName('A09_Flujo_Efectivo_Op');
    if (!s) return;
    s.getRange('D3').setValue(50000); // Saldo inicial
    const data = [
        [new Date(), 'Cobro a Cliente X', 'Ventas', 46400, 0, ''],
        [new Date(), 'Pago de Nómina', 'Operación', 0, 15000, ''],
        [new Date(), 'Pago de Renta', 'Operación', 0, 12000, '']
    ];
    s.getRange('A6:F8').setValues(data);
}

function seedCuentasPagarCobrarData_(ss) {
    const cxp = ss.getSheetByName('C28_Cuentas_Por_Pagar');
    if (cxp) cxp.getRange('A2:G3').setValues([['FP-01', 'Proveedor de Laptops', 'Compra de inventario', new Date(), new Date(new Date().getTime() + 30 * 24 * 60 * 60 * 1000), 370000, 'Pendiente']]);
    const cxc = ss.getSheetByName('C29_Cuentas_Por_Cobrar');
    if (cxc) cxc.getRange('A2:G3').setValues([['FC-01', 'Cliente X', 'Venta de 2 Laptops', new Date(), new Date(new Date().getTime() + 30 * 24 * 60 * 60 * 1000), 46400, 'Pendiente']]);
}

// Stubs para otras funciones de seed, pueden ser implementadas de forma similar
function seedPurchasingData_(ss){}
