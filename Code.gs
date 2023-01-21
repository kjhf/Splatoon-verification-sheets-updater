'use strict';
// Above declaration makes Javascript less jank
// See https://www.w3schools.com/js/js_strict.asp

/**
 * @OnlyCurrentDoc
 */
// Above declaration reduces the permissions that these Google functions need so it's not looking at all your Google accounts and sheets >.<
// See https://developers.google.com/apps-script/guides/services/authorization#manual_authorization_scopes_for_sheets_docs_slides_and_forms

// Note that `MITGoogleScripts.` prefix on the functions is an access into the library.
// For code that is already using these functions without the library as described in the readme, do not prefix the `MITGoogleScripts.`
function onOpen() {
  let battlefyMenu = SpreadsheetApp.getUi().createMenu('Battlefy');
  battlefyMenu.addItem('Create new sheet', 'MITGoogleScripts.createBattlefySheet');
  battlefyMenu.addItem('Update a current sheet', 'MITGoogleScripts.beginUpdateBattlefySheet');
  battlefyMenu.addToUi();
}
