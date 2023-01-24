const VERSION = 2;

/**
 * Get the current version of the spreadsheet
 * @param {SpreadsheetApp.Sheet} sheet 
 */
function getVersion(sheet) {
  if (sheet.rows < 2 || sheet.columns < COLUMNS_TOTAL - 1) {
    debug(`Returning version 1 from rows/columns boundaries for sheet ${sheet.getName()}`);
    return 1;
  }
  else {
    try {
      let val = getCell(sheet, SHEET_INFO_COLUMN_VERSION_ROW, SHEET_INFO_COLUMN);
      if (val) {
        val = parseIntOrThrow(val);
        debug(`Returning version ${val}`);
      }
      else {
        debug(`Returning version 1 from empty sheet info for sheet ${sheet.getName()}`);
        return 1;
      }
    }
    catch (e) {
      debug(`Returning version 1 from catch for sheet ${sheet.getName()}`);
      return 1;
    }
  }
}

/**
 * Upgrade the spreadsheet to the latest version
 * @param {SpreadsheetApp.Sheet} sheet 
 */
function upgradeBattlefySheet(sheet) {
  currentVersion = getVersion(sheet);
  if (currentVersion == VERSION) {
    debug(`Sheet already latest version; skipping.`);
    return;
  }

  if (currentVersion == 1) {
    debug(`Beginning upgrade of ${sheet.getName()}.`);
    let battlefyId = getTournamentIdFromSheet(sheet);
    insertColumn(TEAM_LOGO_URL_COLUMN);
    initialiseBattlefySheet(sheet, battlefyId);
    debug(`Upgrade done for ${sheet.getName()}.`);
  } else {
    throw new Error("Unknown version: " + currentVersion);
  }
}
