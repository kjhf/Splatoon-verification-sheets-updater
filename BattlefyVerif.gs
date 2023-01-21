'use strict';
// Above declaration makes Javascript less jank
// See https://www.w3schools.com/js/js_strict.asp

const BATTLEFY_ID_REGEX = new RegExp("[A-Fa-f0-9]{20,30}","g");
const NUMBER_OF_PLAYERS = 8

const SEEDING_COLUMN = 1
const TEAM_NAME_COLUMN = 2
const PLAYER_NAMES_INDEX = 3
const PLAYER_NAMES_COLUMNS = Array.from({length: NUMBER_OF_PLAYERS}, (_, i) => i + PLAYER_NAMES_INDEX)
const DROPPED_PLAYER_NAMES_INDEX = (PLAYER_NAMES_INDEX + NUMBER_OF_PLAYERS)
const DROPPED_PLAYER_NAMES_COLUMNS = Array.from({length: NUMBER_OF_PLAYERS}, (_, i) => i + DROPPED_PLAYER_NAMES_INDEX)
const NOTES_COLUMN = DROPPED_PLAYER_NAMES_INDEX + NUMBER_OF_PLAYERS
const GAP_COLUMN = (NOTES_COLUMN + 1)
const TEAM_ID_COLUMN = (GAP_COLUMN + 1)
const PLAYER_IDS_INDEX = (TEAM_ID_COLUMN + 1)
const PLAYER_IDS_COLUMNS = Array.from({length: NUMBER_OF_PLAYERS}, (_, i) => i + PLAYER_IDS_INDEX)
const PLAYER_SLUGS_INDEX = (PLAYER_IDS_INDEX + NUMBER_OF_PLAYERS)
const PLAYER_SLUGS_COLUMNS = Array.from({length: NUMBER_OF_PLAYERS}, (_, i) => i + PLAYER_SLUGS_INDEX)
const DROPPED_PLAYER_IDS_INDEX = (PLAYER_SLUGS_INDEX + NUMBER_OF_PLAYERS)
const DROPPED_PLAYER_IDS_COLUMNS = Array.from({length: NUMBER_OF_PLAYERS}, (_, i) => i + DROPPED_PLAYER_IDS_INDEX)
const DROPPED_PLAYER_SLUGS_INDEX = (DROPPED_PLAYER_IDS_INDEX + NUMBER_OF_PLAYERS)
const DROPPED_PLAYER_SLUGS_COLUMNS = Array.from({length: NUMBER_OF_PLAYERS}, (_, i) => i + DROPPED_PLAYER_SLUGS_INDEX)
const UPDATE_TIME_COLUMN = DROPPED_PLAYER_SLUGS_INDEX + NUMBER_OF_PLAYERS
const GUTTER_TOURNAMENT_ID_COLUMN = UPDATE_TIME_COLUMN + 1
const COLUMNS_TOTAL = GUTTER_TOURNAMENT_ID_COLUMN + 1
const HEADER_VALUES = initHeaderValues();
var CACHE = [];

/**
 * Get the header values as a ranged string array with one value (one row) containing the headers' values (COLUMNS_TOTAL columns).
 * @returns {String[][]} The header values.
 */
function initHeaderValues() {
  let result /* {String[][]} */ = Array(1).fill(Array(COLUMNS_TOTAL).fill(""));

  // Note that the order of the columns is set by the constants at the top of the file and not the order here (it's an insertion into the predefined index)
  result[0][UPDATE_TIME_COLUMN - 1] = "Updated At";
  result[0][SEEDING_COLUMN - 1] = "Seed";
  result[0][TEAM_NAME_COLUMN - 1] = "Team Name";
  result[0][TEAM_ID_COLUMN - 1] = "Team Id";
  result[0][NOTES_COLUMN - 1] = "Notes";
  
  for (let i = 0; i < NUMBER_OF_PLAYERS; i++) {
    result[0][PLAYER_NAMES_COLUMNS[i] - 1] = `P${(i + 1)}`;
    result[0][PLAYER_IDS_COLUMNS[i] - 1] = `P${(i + 1)} Id`;
    result[0][PLAYER_SLUGS_COLUMNS[i] - 1] = `P${(i + 1)} Slug`;
    result[0][DROPPED_PLAYER_NAMES_COLUMNS[i] - 1] = `Dropped P${(i + 1)}`;
    result[0][DROPPED_PLAYER_IDS_COLUMNS[i] - 1] = `Dropped P${(i + 1)} Id`;
    result[0][DROPPED_PLAYER_SLUGS_COLUMNS[i] - 1] = `Dropped P${(i + 1)} Slug`;
  }
  return result;
}

function createBattlefySheet() {
  let ui = SpreadsheetApp.getUi();
  let response = ui.prompt('Please enter the Battlefy Tournament Id or its tournament link!', 
    'It will look something like \"622bbfd72282c9043b22e5f9\" or \"https://battlefy.com/mulloway-institute-of-turfing/minnow-cup-12-splat-zones-edition/622bbfd72282c9043b22e5f9/\"', ui.ButtonSet.OK);

  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.OK) {
    let input = response.getResponseText();
    if (input.startsWith('https://') && input.includes("battlefy.com")) {  // May or may not include www.
      let idIndex = input.search(BATTLEFY_ID_REGEX);
      if (idIndex == -1) {
        showSidebarFeedback(`The url specified doesn't contain a tournament id...`);
        return;
      }
      let endIndex = input.indexOf('/', idIndex);
      input = input.substring(idIndex, endIndex);
      debug(`Using id parsed from url: ${input}`);
    } 
    
    if (!isBattlefyId(input)) {
      showSidebarFeedback(`That's not a Battlefy id... ${input}`);
      return;
    }
    
    let battlefyUrl = getBattlefyUrl(input);
    if (urlPageExists(battlefyUrl)) {
      let sheet = createSheet(input);
      
      // Add the columns
      var headerRow = sheet.getRange(1, 1, 1, COLUMNS_TOTAL);
      headerRow.setFontWeight('bold');
      headerRow.setHorizontalAlignment('center');
      headerRow.setVerticalAlignment('middle');
      //headerRow.setFontSize(12);
      //sheet.setRowHeight(1, 34);
      //sheet.setColumnWidths(1, 6, 208);
      HEADER_VALUES[0][GUTTER_TOURNAMENT_ID_COLUMN - 1] = input;  // -1 because [] indexes
      headerRow = headerRow.setValues(HEADER_VALUES);
      
      // Delete rows after the last row
      var maxRows = sheet.getMaxRows(); 
      var lastRow = sheet.getLastRow();
      if (maxRows - lastRow != 0) {
        sheet.deleteRows(lastRow + 1, maxRows - lastRow);
      }

      debug(`Sheet created! ${input}`);
      doUpdateSheet(sheet);
      return sheet;
    } else {
      showSidebarFeedback(`Didn't get a good response from Battlefy! Not creating the sheet.`);
    }
  } else {
    debug(`User cancelled.`);
  }
}

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
 * Get the Battlefy tournament id from the sheet's saved id or (as a fall-back) its name.
 * @param {SpreadsheetApp.Sheet} sheet 
 * @returns {String?} The id, or null if not found.
 */
function getTournamentIdFromSheet(sheet) {
  try {
    return sheet.getRange(1, GUTTER_TOURNAMENT_ID_COLUMN, 1, 1).getValue().toString();
  }
  catch (e) {
    // Attempt fallback to sheetname.
    let sheetName = sheet.getName();
    if (isBattlefyId(sheetName)) {
      return sheetName;
    }
    // Nope.
    return null;
  }
}

/**
 * Update a Battlefy sheet determined from the UI.
 */
function beginUpdateBattlefySheet() {
  let battlefySheets = getBattlefySheets();
  switch (battlefySheets.length)
  {
    case 0:
    {
      showSidebarFeedback(`You have no Battlefy sheets in this workbook. Please add one first!`);
      return;
    }
    case 1:
    {
      // We can shortcut if there's only one candidate
      doUpdateSheet(battlefySheets[0]);
      return;
    }
    default:
    {
      let sheet = SpreadsheetApp.getActiveSheet();
      tournament_id = getTournamentIdFromSheet(sheet);
      
      if (isBattlefyId(tournament_id))
      {
        let ui = SpreadsheetApp.getUi();
        let response = ui.prompt(sheet.getName(), `Update the current page? (Tournament: ${tournament_id})`, ui.ButtonSet.YES_NO);

        // Process the user's response.
        if (response.getSelectedButton() == ui.Button.YES) {
          doUpdateSheet(sheet);
        }
      } else {
        showSidebarFeedback(`The active sheet is not a Battlefy page. Please first select the page you want to update.`);
      }
    }
  }
}

/**
 * Update the specified Battlefy sheet
 * @param {SpreadsheetApp.Sheet} sheet 
 */
function doUpdateSheet(sheet) {
  // Get all the teams from the Battlefy tournament -- 
  // Update the update time on a change.
  //
  // Treat a team as its persistent id - that way, if a team drops and signs up again its registration is kept.
  //
  // Foreach new team:
  // Add a new row
  // Add the team's name and persistent id
  // Foreach player on the team, add a cell for the player's username, id, slug 
  //
  // Foreach team not added, because it already exists:
  // Get the appropriate row for the team
  // Update the update time
  // If the row has been struck-out before, remove the strike-out (only if it's the whole row, not a single player!)
  // Add players that are new
  // For players no longer on the Bfy roster, move THE CELL to the next free in dropped (this keeps notes and formatting that the TOs have made)
  // If the seeding column has DROPPED in it, but the team has not dropped from Bfy, then warn in the feedback.
  //
  // Foreach team that has dropped:
  // Strikethrough the whole row, add DROPPED to the seeding column

  // First check the tournament URL and its data to make sure it's valid
  let tournament_id = getTournamentIdFromSheet(sheet);
  if (!tournament_id) {
    showSidebarFeedback(`The sheet is not a valid Battlefy sheet: ${sheet.getName()}`);
    return;
  }
  
  let battlefyUrl = getBattlefyUrl(tournament_id);
  let response = UrlFetchApp.fetch(battlefyUrl);
  let responseText = response.getContentText();
  
  debug(`responseText length: ${responseText.length}`);
  if (responseText.length == 0) {
    showSidebarFeedback("No incoming text data received from the Bfy server.");
    return;
  }

  let jsonData = JSON.parse(responseText);
  let incomingTeamsJSON = jsonData;
  debug(`incomingTeamsJSON length: ${incomingTeamsJSON.length}`);
  if (incomingTeamsJSON.length == 0) {
    showSidebarFeedback("No incoming JSON data was parsed from the Bfy server.");
    return;
  }
  
  // Cache the sheet (sets the CACHE var).
  // This is necessary because the Google API get and set requests are rate limited and so operations are really slow if you don't cache the sheet.
  cacheAllValues(sheet);
  debug(`sheet cached with length: ${CACHE.length} x ${COLUMNS_TOTAL}`);
  let knownTeamIds = getValuesFromCacheColumn(sheet, TEAM_ID_COLUMN);
  
  let teamIds = new Map();  // Keyed by team id, value is the row number
  for (let rowIndex = 0; rowIndex < knownTeamIds.length; rowIndex++) {
    let teamIdValue = knownTeamIds[rowIndex];
    if (teamIdValue) {
      teamIds.set(teamIdValue, rowIndex + 1); // One-based
    }
  }
  debug(`teamIds size: ${teamIds.size}, processing ...`);

  // Process the incoming teams to add or edit
  let incomingTeamIds = [];
  for (let key in incomingTeamsJSON) {
    let team = incomingTeamsJSON[key];
    let persistentTeamId = team.persistentTeamID;
    incomingTeamIds.push(persistentTeamId);

    let isKnownTeam = teamIds.has(persistentTeamId);
    if (isKnownTeam) {
      doUpdateSheetForKnownTeam(sheet, teamIds.get(persistentTeamId), team);
    } else {
      let row = doUpdateSheetForNewTeam(sheet, team);
      teamIds[persistentTeamId] = row;
    }
  }

  debug(`processing drops...`);
  // Check for dropped teams.
  for (let teamId in teamIds.keys()) {
    if (!incomingTeamIds.includes(teamId)) {
      // Dropped, do stuff.
      debug(`Handling dropped team id ${teamId}, row: ${teamIds[teamId]}`);
      doUpdateSheetForDroppedTeam(sheet, teamIds[teamId]);
    }
  }
  
  debug(`committing the cache...`);
  commitCache(sheet);
  debug(`Done!`);
}

/**
 * Update the specified Battlefy sheet for a known team
 * @param {SpreadsheetApp.Sheet} sheet The workbook sheet
 * @param {number|SpreadsheetApp.Range} row The row that the team is registered for
 * @param {String} teamJson The team JSON
 */
function doUpdateSheetForKnownTeam(sheet, row, teamJson) {
  if (!teamJson) {
    showSidebarFeedback(`Error: doUpdateSheetForKnownTeam: teamJson is null!`);
    return;
  }

  if (typeof row === "SpreadsheetApp.Range") {
    row = row.getRowIndex();  // row is now a number
  }

  //dumpObject(teamJson);
  debug(`doUpdateSheetForKnownTeam called for team ${teamJson.persistentTeamID}`);

  let teamName = teamJson.name;
  if (teamName == "") {
    showSidebarFeedback(`Warning: team has no name. Ignoring the entry. Has the JSON changed format? (Ask Slate/a dev about this one!).`);
    return;
  }

  let hasChanges = false;
  if (getValueFromCache(row, SEEDING_COLUMN) == "DROPPED") {
    showSidebarFeedback(`Warning: team ${teamName} has DROPPED in seeding but has incoming Bfy data.`);
  }

  if (getValueFromCache(row, TEAM_NAME_COLUMN) != teamName) {
    hasChanges = true;
    setValueInCache(row, TEAM_NAME_COLUMN, teamName);
    debug(`Team name has changed: Now ${teamName}`);
  }

  let players = teamJson.players;
  if (!players) {
    players = [];
    showSidebarFeedback(`Team doesn't have players in the JSON! Team: ${teamName}`);
  }

  debug(`Team has ${players.length} players`);

  // First let's take care of drops
  let incomingPlayerIds = [];
  let playersFix = [];
  for (let playerKey in players) {
    let player = null;
    if (players.hasOwnProperty(playerKey)) {
      player = players[playerKey];
    } else {
      player = playerKey;
    }
    let incomingPlayerId = player.persistentPlayerID ?? `SUB-ID-${player._id}`;
    incomingPlayerIds.push(incomingPlayerId);
    playersFix.push(player);
  }
  players = playersFix;
  
  for (let j = 0; j < PLAYER_IDS_COLUMNS.length; j++) {
    let thisPlayerId = getValueFromCache(row, PLAYER_IDS_COLUMNS[j]);
    if (!thisPlayerId) continue;

    if (!incomingPlayerIds.includes(thisPlayerId)) {
      debug(`Handling dropped player ${thisPlayerId} from row ${row} cell index ${j}`)
      // The id in the spreadsheet is not found in the JSON
      dropPlayerFromTeam(sheet, row, j);
      hasChanges = true;
    }
  } 

  // Now edits to existing members, and then additions
  for (let i = 0; i < players.length && i < NUMBER_OF_PLAYERS; i++) {
    let playerFound = false;
    let incomingPlayerName = players[i].inGameName;
    let incomingPlayerId = players[i].persistentPlayerID ?? `SUB-ID-${players[i]._id}`;
    let incomingPlayerSlug = players[i].userSlug ?? `SUB-USER-${players[i]._id}`;

    // Check if the incoming player already exists
    for (let j = 0; j < PLAYER_IDS_COLUMNS.length; j++) {
      let thisPlayerId = getValueFromCache(row, PLAYER_IDS_COLUMNS[j]);
      if (!thisPlayerId) continue;

      if (thisPlayerId == incomingPlayerId) {
        let thisPlayerName = getValueFromCache(row, PLAYER_NAMES_COLUMNS[j]);
        let thisPlayerSlug = getValueFromCache(row, PLAYER_SLUGS_COLUMNS[j]);
        if (thisPlayerName != incomingPlayerName) {
          debug(`Player name for [${j}] has changed: ${thisPlayerName} -> ${incomingPlayerName}.`);
          setValueInCache(row, PLAYER_NAMES_COLUMNS[j], incomingPlayerName);
          hasChanges = true;
        }
        if (thisPlayerSlug != incomingPlayerSlug) {
          debug(`Player slug for [${j}] has changed: ${thisPlayerSlug} -> ${incomingPlayerSlug}.`);
          setValueInCache(row, PLAYER_SLUGS_COLUMNS[j], incomingPlayerSlug);
          hasChanges = true;
        }
        playerFound = true;
        break;
      }
    }
    
    // If not, add it
    if (!playerFound) {
      addNewPlayerToTeam(row, players[i]);
      hasChanges = true;
    }
  }

  if (hasChanges) {
    setValueInCache(row, UPDATE_TIME_COLUMN, getDateNow());
  }
}

/**
 * Update the specified Battlefy sheet for a new team
 * @param {SpreadsheetApp.Sheet} sheet
 * @param {String} teamJson The team JSON
 * @returns {SpreadsheetApp.Range} The row that the team has been added to
 */
function doUpdateSheetForNewTeam(sheet, teamJson) {
  if (!teamJson) {
    showSidebarFeedback(`Error: doUpdateSheetForNewTeam: teamJson is null!`);
    return;
  }
  //dumpObject(teamJson);
  debug(`doUpdateSheetForNewTeam called for team ${teamJson.persistentTeamID}`);

  let teamName = teamJson.name;
  if (teamName == "") {
    showSidebarFeedback(`Warning: team has no name. Ignoring the entry. Has the JSON changed format? (Ask Slate/a dev about this one!).`);
    return;
  }

  let row = createGetLastRow(sheet);
  setValueInCache(row, UPDATE_TIME_COLUMN, getDateNow());
  setValueInCache(row, TEAM_NAME_COLUMN, teamName);
  setValueInCache(row, TEAM_ID_COLUMN, teamJson.persistentTeamID);
  let players = teamJson.players;
  if (players) {
    for (let i = 0; i < players.length && i < NUMBER_OF_PLAYERS; i++) {
      addNewPlayerToTeam(row, players[i], i);
    }
  }
  return row;
}


/**
 * Update the specified Battlefy sheet for a dropped team
 * @param {SpreadsheetApp.Sheet} sheet 
 * @param {SpreadsheetApp.Range} row
 */
function doUpdateSheetForDroppedTeam(sheet, row) {
  debug(`doUpdateSheetForDroppedTeam called for row ${row.getRow()}`);

  setValueInCache(row, UPDATE_TIME_COLUMN, getDateNow());
  setValueInCache(row, SEEDING_COLUMN, "DROPPED");
  setStrikeThrough(row, true);
}

/**
 * Update the specified team row to add a new player
 * @param {number|SpreadsheetApp.Range} row The team's row
 * @param {String} playerJson The JSON containing the player
 * @param {number} index Optional player index for the column to add the player to, otherwise find the next available (null).
 */
function addNewPlayerToTeam(row, playerJson, index = null) {
  if (index === null) {
    index = getFreePlayerSlot(row);
    if (index == -1) {
      showSidebarFeedback(`The spreadsheet can't handle this many players in the team (player: ${playerJson.name}, team: ${getValueFromCache(row, TEAM_NAME_COLUMN)})`);
      return;
    }
  }

  setValueInCache(row, PLAYER_NAMES_COLUMNS[index], playerJson.inGameName);
  setValueInCache(row, PLAYER_SLUGS_COLUMNS[index], playerJson.userSlug ?? `SUB-USER-${playerJson._id}`);
  setValueInCache(row, PLAYER_IDS_COLUMNS[index], playerJson.persistentPlayerID ?? `SUB-ID-${playerJson._id}`);
}

/**
 * Update the specified team row to drop a player
 * @param {SpreadsheetApp.Sheet} sheet 
 * @param {number} row The team's row
 * @param {number} playerIndex The index of the player in the row that is being dropped
 * @param {number} droppedIndex Optional player index for the column to move the drop the player to, otherwise find the next available (null).
 */
 function dropPlayerFromTeam(sheet, row, playerIndex, droppedIndex = null) {
  if (typeof playerIndex !== "number") {
    throw new TypeError(`playerIndex must be a number, actually ${typeof playerIndex} (${playerIndex}).`);
  }

  let idToMove = getValueFromCache(row, PLAYER_IDS_COLUMNS[playerIndex]);

  if (!idToMove) {
    throw new TypeError(`dropPlayerFromTeam expected to find an id to drop, but it returned empty/null.`); 
  }

  // Check if this player has already been dropped
  for (let i = 0; i < DROPPED_PLAYER_IDS_COLUMNS.length; i++) {
    let thisPlayerId = getValueFromCache(row, DROPPED_PLAYER_IDS_COLUMNS[i]);
    if (!thisPlayerId) continue;

    if (thisPlayerId == idToMove) {
      showSidebarFeedback(`The spreadsheet already has this player dropped. Overwriting values. (player: ${idToMove}, team: ${getValueFromCache(row, TEAM_NAME_COLUMN)})`);
      droppedIndex = i;
      break;
    }
  }

  if (droppedIndex === null) {
    droppedIndex = getFreeDroppedSlot(row);
    if (droppedIndex == -1) {
      showSidebarFeedback(`The spreadsheet can't handle this many dropped players in the team (player: ${idToMove}, team: ${getValueFromCache(row, TEAM_NAME_COLUMN)})`);
      return;
    }
  }

  // Copy over
  moveRange(sheet.getRange(row, PLAYER_NAMES_COLUMNS[playerIndex]), sheet.getRange(row, DROPPED_PLAYER_NAMES_COLUMNS[droppedIndex]));
  moveRange(sheet.getRange(row, PLAYER_IDS_COLUMNS[playerIndex]), sheet.getRange(row, DROPPED_PLAYER_IDS_COLUMNS[droppedIndex]));
  moveRange(sheet.getRange(row, PLAYER_SLUGS_COLUMNS[playerIndex]), sheet.getRange(row, DROPPED_PLAYER_SLUGS_COLUMNS[droppedIndex]));
}

/**
 * Get the next free Player index in the row (0-based)
 * @param {number|SpreadsheetApp.Range} row The team's row
 * @returns {number} The free index, from 0 to NUMBER_OF_PLAYERS, or -1 if all slots are full.
 */
function getFreePlayerSlot(row) {
  for (let i = 0; i < NUMBER_OF_PLAYERS; i++) {
    let value = getValueFromCache(row, PLAYER_NAMES_COLUMNS[i]);
    if (!value) {
      return i;
    }
  }
  return -1;
}

/**
 * Get the next free Dropped Player index in the row.
 * @param {number|SpreadsheetApp.Range} row The team's row
 * @returns {number} The free index, from 0 to NUMBER_OF_PLAYERS, or -1 if all slots are full.
 */
 function getFreeDroppedSlot(row) {
  for (let i = 0; i < NUMBER_OF_PLAYERS; i++) {
    let value = getValueFromCache(row, DROPPED_PLAYER_NAMES_COLUMNS[i]);
    if (!value) {
      return i;
    }
  }
  return -1;
}

/**
 * Get if the specified URL exists
 * @param {String} url 
 * @returns {Boolean}
 */
function urlPageExists(url) {
  const params = {"method" : "GET", "headers": {"X-HTTP-Method-Override": "HEAD"}};
  let response = UrlFetchApp.fetch(url, params);
  let code = response.getResponseCode();
  debug(`Response code: ${code} from ${url}`);
  return code >= 200 && code <= 399;
}

/**
 * Get if the string is a battlefy id by its length and format
 * @param {String} input 
 * @returns {Boolean}
 */
function isBattlefyId(input) {
  return input && 20 <= input.length && input.length < 30 && BATTLEFY_ID_REGEX.test(input)
}

/**
 * Get the backend cloud link for the Battlefy id
 * @param {String} id
 * @returns {String}
 */
function getBattlefyUrl(id) {
  return `https://dtmwra1jsgyb0.cloudfront.net/tournaments/${id}/teams`
}

/**
 * Get the Battlefy sheets from this Workbook
 * @returns {SpreadsheetApp.Sheet[]}
 */
function getBattlefySheets() {
  let sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  let battlefySheets = [];
  for (let i = 0; i < sheets.length; i++) {
    let candidate = sheets[i];
    let candidateId = getTournamentIdFromSheet(candidate);
    if (isBattlefyId(candidateId))
    {
      battlefySheets.push(candidate);
    }
  }
  return battlefySheets;
}

/**
 * Create a last row and return its Range.
 * @param {SpreadsheetApp.Sheet} sheet 
 * @returns {SpreadsheetApp.Range}
 */
function createGetLastRow(sheet) {
  let rowIndex = CACHE.length;
  sheet.insertRowAfter(rowIndex);
  CACHE.push(Array(COLUMNS_TOTAL).fill(""));
  let range = sheet.getRange(rowIndex + 1, 1, 1, COLUMNS_TOTAL);
  range.setFontWeight('normal');
  range.setHorizontalAlignment('left');
  return range;
}
