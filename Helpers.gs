'use strict'

/* 
 * Helper functions to support the main Code.
 */

function pause(title = 'Please confirm', message = 'Continue?') {
  let ui = SpreadsheetApp.getUi()
  return ui.alert(
    title,
    message,
    ui.ButtonSet.YES_NO) 
    == ui.Button.YES;
}

function debug(logMessage) {
  Logger.log(`[DEBUG] ${logMessage}`);
}

function getSheetById(id) {
  return SpreadsheetApp.getActive().getSheets().filter(
    function(s) {return s.getSheetId() === id;}
  )[0];
}

function showSidebarFeedback(message) {
   var html = HtmlService.createHtmlOutput()
      .setTitle('Slate says...')
      .setFaviconUrl('https://media.discordapp.net/attachments/471361750986522647/758104388824072253/icon.png')
      .append(message)
      ;
  SpreadsheetApp.getUi().showSidebar(html);
}

function showAlert(message) {
  let ui = SpreadsheetApp.getUi()
  ui.alert(
     'Slapp',
     message,
      ui.ButtonSet.OK);
}

function parseIntOrThrow(string, radix = null) {
  let num = parseInt(string, radix);
  if (isNaN(num)) {
    throw new TypeError("Not a number.");
  }
  return num;
}

/* Build a URL query options string like URLSearchParams (unavailable in GScript). Does not have a leading ?. */
function urlQueryBuilder(obj) {
  return Object.keys(obj).reduce(function(p, e, i) {
    return p + (i == 0 ? "" : "&") +
      (Array.isArray(obj[e]) ? obj[e].reduce(function(str, f, j) {
        return str + e + "=" + encodeURIComponent(f) + (j != obj[e].length - 1 ? "&" : "")
      },"") : e + "=" + encodeURIComponent(obj[e]));
  },"");
}

/**
 * Dumps a JS object to the debug log.
 * @param {Object} obj 
 * @returns {String} of the dumped object
 */
function dumpObject(obj) {
  let str = JSON.stringify(obj, null, 2);
  debug(str);
}


/**
 * Set the strikethrough property on the range. 
 * @param {SpreadsheetApp.Range} range The cells to apply the action to.
 * @param {Boolean} enabled true/false to set/clear
 */
function setStrikeThrough(range, enabled) {
  if (enabled) {
    range.setFontLine("line-through");
  } else {
    range.setFontLine("none");
  }
}

/**
 * Get the current date as an ISO string.
 * @returns {String} The date as an ISO string.
 */
function getDateNow() {
  return (new Date()).toISOString();
}
