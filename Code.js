/**
 * Runs when the add-on is installed.
 *
 * @param {object} e The event parameter for a simple onInstall trigger. To
 *     determine which authorization mode (ScriptApp.AuthMode) the trigger is
 *     running in, inspect e.authMode. (In practice, onInstall triggers always
 *     run in AuthMode.FULL, but onOpen triggers may be AuthMode.LIMITED or
 *     AuthMode.NONE).
 */
function onInstall(e) {
  onOpen();
}

function onOpen() {
  
}

function buildHomePage(){

  const movements = StockApp.getMovementSheetKeys(SpreadsheetApp.getActiveSpreadsheet())
  const card = StockApp.buildCard(movements);
  return card;

  //return [card];
}