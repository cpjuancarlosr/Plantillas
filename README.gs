/**
 * //== File: README.gs
 * Generador para la hoja de bienvenida e instrucciones (README).
 */

/**
 * Crea y formatea la hoja README con instrucciones, mapa y glosario.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - La hoja de cálculo activa.
 */
function createReadmeSheet_(ss) {
  const sheetName = 'README';
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    ss.deleteSheet(sheet);
  }

  sheet = ss.insertSheet(sheetName, 0); // Insertar al principio
  ss.setActiveSheet(sheet);

  // Limpieza y formato básico
  sheet.getRange('A1:E50').clear();
  sheet.setColumnWidth(1, 400);
  sheet.setColumnWidth(2, 500);

  // --- Contenido ---

  // Título
  sheet.getRange('A1').setValue('🚀 Centro de Control Empresarial v1.0 (Plantillas JC)')
    .setFontSize(18).setFontWeight('bold').setFontFamily('Arial');
  sheet.getRange('A2').setValue('Bienvenido a tu suite de plantillas inteligentes. Todo está conectado para darte una visión 360° de tu negocio.')
    .setFontStyle('italic').setWrap(true);

  // Cómo Empezar
  sheet.getRange('A4').setValue('🎯 CÓMO EMPEZAR')
    .setFontSize(14).setFontWeight('bold');
  const steps = [
    ['1. Explora las Plantillas', 'Navega por las pestañas generadas. Cada una está diseñada para una función específica.'],
    ['2. Carga tus Propios Datos', 'Usa "Herramientas > Limpiar Todos los Datos" para borrar los ejemplos. Luego, empieza a capturar tu información real.'],
    ['3. Usa el Menú "Plantillas JC"', 'Todas las acciones principales (crear, exportar, limpiar) están en el menú superior.'],
    ['4. Exporta a PDF', 'Cualquier tablero o reporte puede ser exportado a PDF desde el menú "Exportar a PDF". Úsalo para tus juntas o archivos.'],
    ['5. Regenera si es necesario', 'Si algo se desajusta, las opciones "Regenerar Formatos" y "Resetear Validaciones" en el menú "Herramientas" pueden restaurar el orden.']
  ];
  sheet.getRange('A5:B9').setValues(steps).setWrap(true).setVerticalAlignment('top');
  sheet.getRange('A5:A9').setFontWeight('bold');

  // Mapa de Pestañas
  sheet.getRange('A11').setValue('🗺️ MAPA DE PESTAÑAS (ÍNDICE)')
    .setFontSize(14).setFontWeight('bold');
  const map = [
    ['Módulo Administración (A01-A10)', 'Control de operaciones, ventas, compras, proyectos y finanzas del día a día.'],
    ['Módulo Contable/Fiscal (B11-B24)', 'Corazón contable del sistema. Desde pólizas hasta estados financieros y cálculo de impuestos (SAT-MX).'],
    ['Módulo Oficina (C25-C30)', 'Herramientas para la productividad personal y de equipo. Agendas, minutas, control de documentos.']
  ];
  sheet.getRange('A12:B14').setValues(map).setWrap(true).setVerticalAlignment('top');
  sheet.getRange('A12:A14').setFontWeight('bold');

  // Glosario Fiscal
  sheet.getRange('A16').setValue('💡 GLOSARIO FISCAL BÁSICO (MÉXICO)')
      .setFontSize(14).setFontWeight('bold');
  const glossary = [
    ['CFDI', 'Comprobante Fiscal Digital por Internet. El estándar de factura electrónica en México.'],
    ['IVA Acreditable', 'El IVA que pagas en tus compras y gastos. Generalmente, puedes restarlo del IVA que cobras.'],
    ['IVA Trasladado', 'El IVA que cobras a tus clientes en tus ventas. Es el impuesto que "trasladas" al cliente.'],
    ['ISR PM', 'Impuesto Sobre la Renta para Personas Morales (empresas). Se calcula sobre las utilidades.'],
    ['Pago Provisional', 'Pagos mensuales a cuenta del impuesto anual (ISR). Se calculan con un coeficiente de utilidad.'],
    ['PTU', 'Participación de los Trabajadores en las Utilidades. Un derecho constitucional de los trabajadores a recibir un porcentaje de las ganancias de la empresa.']
  ];
  sheet.getRange('A17:B22').setValues(glossary).setWrap(true).setVerticalAlignment('top');
  sheet.getRange('A17:A22').setFontWeight('bold');

  logAction_('Hoja README creada/actualizada.');
}
