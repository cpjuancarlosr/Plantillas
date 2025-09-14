/**
 * //== File: styles.gs
 * Funciones para aplicar estilos, formatos y temas a las hojas de cálculo.
 */

// Definiciones de estilo reutilizables para batchUpdate
const FONT_FAMILY = 'Arial';

const BASE_TEXT_FORMAT = {
  foregroundColor: { red: 0.067, green: 0.067, blue: 0.067 }, // #111111
  fontFamily: FONT_FAMILY,
  fontSize: 10
};

const BOLD_TEXT_FORMAT = { ...BASE_TEXT_FORMAT, bold: true };

const HEADER_FORMAT = {
  userEnteredFormat: {
    backgroundColor: { red: 0.96, green: 0.96, blue: 0.96 }, // #f5f5f5
    textFormat: { ...BOLD_TEXT_FORMAT, fontSize: 11 },
    horizontalAlignment: 'CENTER',
    verticalAlignment: 'MIDDLE'
  }
};

const KPI_CARD_FORMAT = {
  userEnteredFormat: {
    backgroundColor: { red: 1, green: 1, blue: 1 },
    textFormat: { ...BOLD_TEXT_FORMAT, fontSize: 24, foregroundColor: { red: 0, green: 0.658, blue: 0.47 } }, // #00a878
    horizontalAlignment: 'CENTER',
    verticalAlignment: 'MIDDLE'
  }
};

const KPI_LABEL_FORMAT = {
  userEnteredFormat: {
    textFormat: { ...BASE_TEXT_FORMAT, fontSize: 9, italic: true },
    horizontalAlignment: 'CENTER',
    verticalAlignment: 'TOP'
  }
};

const SOFT_BORDER = {
  style: 'SOLID',
  width: 1,
  color: { red: 0.9, green: 0.9, blue: 0.9 } // #e6e6e6
};


/**
 * Aplica el tema de color personalizado a la hoja de cálculo.
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss - La hoja de cálculo.
 */
function applyCustomTheme_(ss) {
  const theme = ss.getPredefinedSpreadsheetThemes().find(t => t.getName() === 'Plantillas JC Theme');
  if (!theme) {
    const newTheme = SpreadsheetApp.newSpreadsheetTheme()
      .setPrimaryColor(THEME_COLORS.ACCENT)
      .setConcreteColor(SpreadsheetApp.ThemeColorType.BACKGROUND, SpreadsheetApp.newColor().setRgbColor(THEME_COLORS.BACKGROUND).build())
      .setConcreteColor(SpreadsheetApp.ThemeColorType.TEXT, SpreadsheetApp.newColor().setRgbColor(THEME_COLORS.TEXT).build())
      .setConcreteColor(SpreadsheetApp.ThemeColorType.ACCENT1, SpreadsheetApp.newColor().setRgbColor(THEME_COLORS.ACCENT).build())
      .setFontFamily(FONT_FAMILY)
      .build();
    ss.setSpreadsheetTheme(newTheme);
  } else {
    ss.setSpreadsheetTheme(theme);
  }
}

/**
 * Función central para aplicar todos los estilos a una hoja específica.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - La hoja a formatear.
 */
function applyBaseSheetStyles_(sheet) {
  const ssId = sheet.getParent().getId();
  const sheetId = sheet.getSheetId();

  const requests = [{
    updateSheetProperties: {
      properties: {
        sheetId: sheetId,
        gridProperties: {
          hideGridlines: true
        }
      },
      fields: 'gridProperties.hideGridlines'
    }
  }];

  // Formato general: texto y números
  const allCellsRange = {
    sheetId: sheetId,
    startRowIndex: 0,
    endRowIndex: sheet.getMaxRows(),
    startColumnIndex: 0,
    endColumnIndex: sheet.getMaxColumns()
  };

  requests.push({
    repeatCell: {
      range: allCellsRange,
      cell: {
        userEnteredFormat: {
          backgroundColor: { red: 1, green: 1, blue: 1 },
          textFormat: BASE_TEXT_FORMAT,
          numberFormat: { type: 'NUMBER', pattern: '#,##0.00' },
          verticalAlignment: 'MIDDLE'
        }
      },
      fields: 'userEnteredFormat(backgroundColor,textFormat,numberFormat,verticalAlignment)'
    }
  });

  batchUpdate_(ssId, requests);

  // Congelar la primera fila (encabezados)
  sheet.setFrozenRows(1);
}

/**
 * Crea una tarjeta KPI con borde.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - La hoja.
 * @param {string} rangeA1 - El rango para la tarjeta (ej. 'B2:D3').
 * @param {string} label - El texto para la etiqueta debajo de la tarjeta.
 */
function styleKpiCard_(sheet, rangeA1, label) {
    const range = sheet.getRange(rangeA1);
    const ssId = sheet.getParent().getId();
    const sheetId = sheet.getSheetId();

    const gridRange = {
        sheetId: sheetId,
        startRowIndex: range.getRow() - 1,
        endRowIndex: range.getLastRow(),
        startColumnIndex: range.getColumn() - 1,
        endColumnIndex: range.getLastColumn()
    };

    const requests = [
        // Fusionar celdas para el valor del KPI
        { mergeCells: { range: gridRange, mergeType: 'MERGE_ALL' } },
        // Aplicar formato de valor KPI
        { repeatCell: { range: gridRange, cell: KPI_CARD_FORMAT, fields: 'userEnteredFormat' } },
        // Aplicar bordes
        {
            updateBorders: {
                range: gridRange,
                top: SOFT_BORDER,
                bottom: SOFT_BORDER,
                left: SOFT_BORDER,
                right: SOFT_BORDER
            }
        }
    ];

    // Etiqueta debajo de la tarjeta
    const labelRow = range.getLastRow() + 1;
    const labelRange = sheet.getRange(labelRow, range.getColumn(), 1, range.getNumColumns());
    const labelGridRange = {
        sheetId: sheetId,
        startRowIndex: labelRow - 1,
        endRowIndex: labelRow,
        startColumnIndex: range.getColumn() - 1,
        endColumnIndex: range.getLastColumn()
    };

    requests.push(
        { mergeCells: { range: labelGridRange, mergeType: 'MERGE_ALL' } },
        { repeatCell: { range: labelGridRange, cell: KPI_LABEL_FORMAT, fields: 'userEnteredFormat' } }
    );

    labelRange.setValue(label);

    batchUpdate_(ssId, requests);
}

/**
 * Aplica el formato de encabezado a un rango.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - La hoja.
 * @param {string} rangeA1 - El rango de los encabezados.
 */
function styleHeaderRow_(sheet, rangeA1) {
    const range = sheet.getRange(rangeA1);
    const ssId = sheet.getParent().getId();
    const sheetId = sheet.getSheetId();

    const gridRange = {
        sheetId: sheetId,
        startRowIndex: range.getRow() - 1,
        endRowIndex: range.getLastRow(),
        startColumnIndex: range.getColumn() - 1,
        endColumnIndex: range.getLastColumn()
    };

    const requests = [{
        repeatCell: {
            range: gridRange,
            cell: HEADER_FORMAT,
            fields: 'userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)'
        }
    }];

    batchUpdate_(ssId, requests);
}

/**
 * Recorre todas las hojas y reaplica los estilos base.
 * Usado para la función de 'Regenerar Formatos'.
 */
function applyStylingToAllSheets_() {
    const allSheets = getActiveSpreadsheet_().getSheets();
    // Excluir hojas de sistema
    const sheetsToStyle = allSheets.filter(s => !s.getName().startsWith('_'));

    sheetsToStyle.forEach((sheet, index) => {
        SpreadsheetApp.getActiveSpreadsheet().toast(`Aplicando estilo a: ${sheet.getName()} (${index + 1}/${sheetsToStyle.length})`, 'Progreso', 5);
        applyBaseSheetStyles_(sheet);
        // Aquí se podrían llamar a las funciones específicas de cada plantilla para regenerar KPIs, etc.
        // Por simplicidad, esta versión solo aplica el formato base.
    });
}
