
/**
 * Create a sheet with the specified name in the workbook and return it.
 * @param {String} name Optional name of the Workbook, or leave blank to title it as the current date. 
 * @returns {SpreadsheetApp.Sheet} The created sheet. 
 */
function createSheet(name = null) {
  if (name === null) {
    name = (new Date()).toLocaleDateString();
  }
  let sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
  return sheet;
}

/**
 * Trim all rows after the last one in use.
 * @param {SpreadsheetApp.Sheet} sheet 
 */
function trimSheetRows(sheet) {
  let maxRows = sheet.getMaxRows(); 
  let lastRow = sheet.getLastRow();
  if (maxRows - lastRow != 0) {
    sheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }
}

/**
 * Set the header row (1st row) of the sheet to the specified Object[1][...]
 * @param {SpreadsheetApp.Sheet} sheet 
 * @param {String[][]} headerValues 
 */
function setHeaderRow(sheet, headerValues) {
  let headerRow = sheet.getRange(1, 1, 1, headerValues[0].length);
  headerRow.setFontWeight('bold');
  headerRow.setHorizontalAlignment('center');
  headerRow.setVerticalAlignment('middle');
  return headerRow = headerRow.setValues(headerValues);
}

/**
 * Hard read a single cell from the document.
 * @param {SpreadsheetApp.Sheet} sheet 
 * @param {number} row The row to retrieve the value from (as a 1-based number).
 * @param {number} column The column to retrieve the value from (as a 1-based number).
 */
function getCell(sheet, row, column) {
  sheet.getRange(row, column, 1, 1).getValue().toString();
}
