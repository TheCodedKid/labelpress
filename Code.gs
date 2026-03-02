/**
 * Avery Label Generator - Google Sheets Add-on
 * Generates printable PDF label sheets from spreadsheet data.
 *
 * SETUP: In your Google Sheet, go to Extensions > Apps Script,
 * then copy these files into the script editor.
 */

// ============================================================
// AVERY LABEL TEMPLATES (dimensions in inches)
// ============================================================
const AVERY_TEMPLATES = {
  "5160": {
    name: "5160 - Address Labels (30/sheet)",
    columns: 3,
    rows: 10,
    labelWidth: 2.625,
    labelHeight: 1.0,
    topMargin: 0.5,
    sideMargin: 0.19,
    horizontalGap: 0.125,
    verticalGap: 0.0,
    pageWidth: 8.5,
    pageHeight: 11.0,
    fontSize: 9,
    description: "Standard address/mailing labels"
  },
  "5161": {
    name: "5161 - Address Labels (20/sheet)",
    columns: 2,
    rows: 10,
    labelWidth: 4.0,
    labelHeight: 1.0,
    topMargin: 0.5,
    sideMargin: 0.15625,
    horizontalGap: 0.1875,
    verticalGap: 0.0,
    pageWidth: 8.5,
    pageHeight: 11.0,
    fontSize: 10,
    description: "Address labels, wider format"
  },
  "5163": {
    name: "5163 - Shipping Labels (10/sheet)",
    columns: 2,
    rows: 5,
    labelWidth: 4.0,
    labelHeight: 2.0,
    topMargin: 0.5,
    sideMargin: 0.15625,
    horizontalGap: 0.1875,
    verticalGap: 0.0,
    pageWidth: 8.5,
    pageHeight: 11.0,
    fontSize: 11,
    description: "Shipping/package labels"
  },
  "5164": {
    name: "5164 - Shipping Labels (6/sheet)",
    columns: 2,
    rows: 3,
    labelWidth: 4.0,
    labelHeight: 3.333,
    topMargin: 0.5,
    sideMargin: 0.15625,
    horizontalGap: 0.1875,
    verticalGap: 0.0,
    pageWidth: 8.5,
    pageHeight: 11.0,
    fontSize: 12,
    description: "Large shipping labels"
  },
  "5167": {
    name: "5167 - Return Address (80/sheet)",
    columns: 4,
    rows: 20,
    labelWidth: 1.75,
    labelHeight: 0.5,
    topMargin: 0.5,
    sideMargin: 0.3125,
    horizontalGap: 0.3125,
    verticalGap: 0.0,
    pageWidth: 8.5,
    pageHeight: 11.0,
    fontSize: 7,
    description: "Small return address labels"
  },
  "5162": {
    name: "5162 - Address Labels (14/sheet)",
    columns: 2,
    rows: 7,
    labelWidth: 4.0,
    labelHeight: 1.333,
    topMargin: 0.833,
    sideMargin: 0.15625,
    horizontalGap: 0.1875,
    verticalGap: 0.0,
    pageWidth: 8.5,
    pageHeight: 11.0,
    fontSize: 10,
    description: "Address labels, taller format"
  },
  "8160": {
    name: "8160 - Address Labels (30/sheet, Inkjet)",
    columns: 3,
    rows: 10,
    labelWidth: 2.625,
    labelHeight: 1.0,
    topMargin: 0.5,
    sideMargin: 0.19,
    horizontalGap: 0.125,
    verticalGap: 0.0,
    pageWidth: 8.5,
    pageHeight: 11.0,
    fontSize: 9,
    description: "Inkjet address labels (same layout as 5160)"
  }
};

// ============================================================
// MENU & SIDEBAR
// ============================================================

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 LabelPress')
    .addItem('Generate Labels...', 'showSidebar')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Avery Label Generator')
    .setWidth(350);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ============================================================
// DATA HELPERS (called from sidebar)
// ============================================================

function getTemplates() {
  return AVERY_TEMPLATES;
}

function getSheetHeaders() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  return headers.filter(h => h !== '');
}

function getSheetData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  return data.map(row => {
    const obj = {};
    headers.forEach((h, i) => {
      if (h !== '') obj[h] = row[i];
    });
    return obj;
  });
}

// ============================================================
// LABEL GENERATION
// ============================================================

/**
 * Render a template string like "{{Name}}\n{{City}}, {{State}} {{Zip}}"
 * against a data row. Lines that are empty after substitution are removed.
 */
function renderTemplate(templateStr, row) {
  const lines = templateStr.split('\n');
  const rendered = lines.map(line => {
    return line.replace(/\{\{(.+?)\}\}/g, (match, field) => {
      const val = row[field.trim()];
      return (val !== undefined && val !== null && val !== '') ? String(val) : '';
    });
  });
  // Remove lines that are only whitespace/punctuation after substitution
  return rendered.filter(line => line.replace(/[\s,;:\-\/]/g, '') !== '').join('\n');
}

/**
 * Helper: set a paragraph to near-invisible (1pt, zero spacing)
 */
function shrinkParagraph(para) {
  para.setFontSize(1);
  para.setSpacingBefore(0);
  para.setSpacingAfter(0);
  para.setLineSpacing(0);
}

/**
 * Main entry point called from the sidebar.
 */
function generateLabels(config) {
  let template;
  if (config.templateId === 'custom' && config.customTemplate) {
    template = config.customTemplate;
  } else {
    template = AVERY_TEMPLATES[config.templateId];
  }
  if (!template) throw new Error('Unknown template: ' + config.templateId);

  const data = getSheetData();
  if (data.length === 0) throw new Error('No data found (need headers in row 1, data in row 2+).');

  const labelTemplateStr = config.labelTemplate;
  if (!labelTemplateStr || !labelTemplateStr.trim()) {
    throw new Error('Label template is empty. Add fields like {{Name}}.');
  }

  const startPosition = config.startPosition || 0;
  const labelsPerPage = template.columns * template.rows;

  // Build label texts
  const labelTexts = [];
  for (let i = 0; i < startPosition; i++) labelTexts.push('');
  data.forEach(row => labelTexts.push(renderTemplate(labelTemplateStr, row)));

  const totalLabels = labelTexts.length;
  const totalPages = Math.ceil(totalLabels / labelsPerPage);

  // Create doc
  const doc = DocumentApp.create('Labels - ' + template.name + ' - ' + new Date().toLocaleDateString());
  const body = doc.getBody();

  body.setPageWidth(template.pageWidth * 72);
  body.setPageHeight(template.pageHeight * 72);
  body.setMarginTop(template.topMargin * 72);
  body.setMarginBottom(0);
  body.setMarginLeft(template.sideMargin * 72);
  body.setMarginRight(template.sideMargin * 72);

  const labelWidthPt = template.labelWidth * 72;
  const labelHeightPt = template.labelHeight * 72;

  for (let page = 0; page < totalPages; page++) {

    // For pages after the first, insert a minimal page break paragraph
    if (page > 0) {
      const pb = body.appendParagraph('');
      shrinkParagraph(pb);
      pb.appendPageBreak();
    }

    const table = body.appendTable();
    table.setBorderWidth(0);

    // On first iteration, remove the default empty paragraph that the doc starts with
    if (page === 0) {
      body.getChild(0).removeFromParent();
    }

    for (let row = 0; row < template.rows; row++) {
      const tableRow = table.appendTableRow();
      tableRow.setMinimumHeight(labelHeightPt);

      for (let col = 0; col < template.columns; col++) {
        const labelIndex = page * labelsPerPage + row * template.columns + col;
        const cell = tableRow.appendTableCell();

        cell.setWidth(labelWidthPt);
        cell.setPaddingTop(4);
        cell.setPaddingBottom(2);
        cell.setPaddingLeft(6);
        cell.setPaddingRight(4);

        const text = labelIndex < totalLabels ? labelTexts[labelIndex] : '';
        const lines = text !== '' ? text.split('\n') : [' '];

        lines.forEach((line, lineIdx) => {
          const safeText = line !== '' ? line : ' ';
          let para;
          if (lineIdx === 0) {
            para = cell.getChild(0).asParagraph();
            para.setText(safeText);
          } else {
            para = cell.appendParagraph(safeText);
          }
          para.setFontSize(template.fontSize);
          para.setSpacingBefore(0);
          para.setSpacingAfter(0);
          para.setLineSpacing(1.15);
          para.setFontFamily('Arial');
        });

        cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);

        // Add a spacer column for the gap between labels (not after last column)
        if (col < template.columns - 1 && template.horizontalGap > 0) {
          const spacer = tableRow.appendTableCell();
          spacer.setWidth(template.horizontalGap * 72);
          spacer.setPaddingTop(0);
          spacer.setPaddingBottom(0);
          spacer.setPaddingLeft(0);
          spacer.setPaddingRight(0);
          const spacerPara = spacer.getChild(0).asParagraph();
          spacerPara.setText(' ');
          spacerPara.setFontSize(1);
          spacerPara.setSpacingBefore(0);
          spacerPara.setSpacingAfter(0);
        }
      }
    }

    // Remove the empty first row appendTable() creates
    if (table.getNumRows() > template.rows) {
      table.removeRow(0);
    }
  }

  // Shrink any trailing paragraph so it doesn't cause a blank page
  const last = body.getChild(body.getNumChildren() - 1);
  if (last.getType() === DocumentApp.ElementType.PARAGRAPH) {
    shrinkParagraph(last.asParagraph());
  }

  doc.saveAndClose();

  // Convert to PDF
  const docFile = DriveApp.getFileById(doc.getId());
  const pdfBlob = docFile.getAs('application/pdf');
  pdfBlob.setName('Labels_' + config.templateId + '_' + new Date().toISOString().slice(0, 10) + '.pdf');
  const pdfFile = DriveApp.createFile(pdfBlob);

  return {
    pdfUrl: pdfFile.getUrl(),
    docUrl: 'https://docs.google.com/document/d/' + doc.getId() + '/edit',
    labelCount: totalLabels - startPosition,
    pageCount: totalPages
  };
}
