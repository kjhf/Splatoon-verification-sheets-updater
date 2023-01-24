
/**
 * Get all the values in the spreadsheet.
 * @param {SpreadsheetApp.Sheet} sheet 
 * @returns {Object[]} The values.
 */
function cacheAllValues(sheet) {
  CACHE = sheet.getSheetValues(1, 1, sheet.getMaxRows(), COLUMNS_TOTAL);
  debug(`sheet cached with length: ${CACHE.length} x ${COLUMNS_TOTAL}`);
  return CACHE;
}

/**
 * Get the specified row values from the cache.
 * @param {number|SpreadsheetApp.Range} row The row to retrieve the value from (as a range or 1-based number).
 * @returns The value[] from the specified row from the cache.
 */
 function getValuesFromCacheRow(row) {
  if (typeof row.getRow === "function") {
    row = row.getRow();
    if (row === null || row === undefined) {
      throw new TypeError(`Row has getRow function, but it returned null or undefined.`);
    }
  }
  if (row < 1 || row > CACHE.length) {
    return '';
  }
  return CACHE[row - 1];
}

/**
 * Get the specified column values from the cache.
 * @param {SpreadsheetApp.Sheet} sheet The sheet (to retrieve the rows)
 * @param {number|SpreadsheetApp.Range} column The column to retrieve the value from (as a range or 1-based number).
 * @returns The value[] from the specified column from the cache.
 */
 function getValuesFromCacheColumn(sheet, column) {
  if (typeof column.getColumn === "function") {
    column = column.getColumn();
    if (column === null || column === undefined) {
      throw new TypeError(`Column has getColumn function, but it returned null or undefined.`);
    }
  }
  return Array.from({length: sheet.getMaxRows()}, (_, i) => getValueFromCache(i+1, column));
}

/**
 * Get the specified cell value from the cache.
 * @param {number|SpreadsheetApp.Range} row The row to retrieve the value from (as a range or 1-based number).
 * @param {number} column The column to retrieve the value from as a Google Apps Scripts column index (1-based).
 * @returns The value from the specified row and column from the cache.
 */
function getValueFromCache(row, column) {
  if (row === null || row === undefined) {
    throw new TypeError(`Row cannot be null or undefined.`);
  }
  if (typeof row.getRow === "function") {
    row = row.getRow();
    if (row === null || row === undefined) {
      throw new TypeError(`Row has getRow function, but it returned null or undefined.`);
    }
  }
  if (row < 1 || row > CACHE.length) {
    return '';
  }
  if (column < 1 || column > COLUMNS_TOTAL) {
    return '';
  }
  return CACHE[row - 1][column - 1];
}

/**
 * Get the specified cell value from the cache.
 * @param {number|SpreadsheetApp.Range} row The row to retrieve the value from (as a range or 1-based number).
 * @param {number} column The column to retrieve the value from as a Google Apps Scripts column index (1-based).
 * @param {Object} value The value from the specified row and column from the cache.
 */
function setValueInCache(row, column, value) {
  if (row === null || row === undefined) {
    throw new TypeError(`Row cannot be null or undefined.`);
  }
  if (column === null || column === undefined) {
    throw new TypeError(`Column cannot be null or undefined.`);
  }
  if (value === null || value === undefined) {
    throw new TypeError(`Value cannot be null or undefined.`);
  }

  if (typeof row.getRow === "function") {
    row = row.getRow();
    if (row === null || row === undefined) {
      throw new TypeError(`Row has getRow function, but it returned null or undefined.`);
    }
  } else if (typeof row != "number") {
    throw new TypeError(`Row must be a number, actually ${typeof row} (${row}).`);
  }

  if (typeof column != "number") {
    throw new TypeError(`Column must be a number, actually ${typeof column} (${column}).`);
  }

  if (row < 1) {
    throw new Error(`Row ${row} is out of bounds (1-based!)`);
  }
  if (row > CACHE.length) {
    throw new Error(`Row ${row} is out of bounds, call createGetLastRow to stay in sync`);
  }
  if (column < 1) {
    throw new Error(`Column ${column} is out of bounds (1-based!)`);
  }
  if (column > COLUMNS_TOTAL) {
    throw new Error(`Column ${column} is out of bounds (>${COLUMNS_TOTAL})`);
  }

  try {
    CACHE[row - 1][column - 1] = value;
  }
  catch (e) {
    throw new Error(`Error setting value: ${e} for CACHE[${row - 1}][${column - 1}] where CACHE is ${CACHE.length} rows and ${COLUMNS_TOTAL} columns`);
  }
}

/**
 * Move the cells in both the sheet and cache (to preserve formatting).
 * @param {SpreadsheetApp.Range} source The range to move
 * @param {SpreadsheetApp.Range} destination The destination range
 */
 function moveRange(source, destination) {
  source.moveTo(destination);
  CACHE[destination.getRow() - 1][destination.getColumn() - 1] = CACHE[source.getRow() - 1][source.getColumn() - 1];
  CACHE[source.getRow() - 1][source.getColumn() - 1] = "";
}

/**
 * Add one row to the cache.
 */
function appendCacheRow() {
  CACHE.push(Array(COLUMNS_TOTAL).fill(""));
}

/**
 * Insert a column to the cache.
 */
function insertColumn(columnIndex, defaultValue = "") {
  for (let i = 0; i < CACHE.length; i++) {
    CACHE[i].splice(columnIndex, 0, defaultValue);
  }
}

/**
 * Commit the cache to the sheet.
 * @param {SpreadsheetApp.Sheet} sheet The sheet to commit to.
 */
function commitCache(sheet) {
  debug(`committing the cache...`);
  sheet.getRange(1, 1, CACHE.length, COLUMNS_TOTAL).setValues(CACHE);
}
