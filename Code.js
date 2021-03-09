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

  //loadTranslations(); // TODO use module pattern instead
  
  // create menu
  const ui = SpreadsheetApp.getUi(); // Or DocumentApp.

  const addOnMenu = ui.createAddonMenu();
    
  //createStockMenu(ui,addOnMenu); // TODO use module pattern for stock and invoice
  createInvoiceMenu(ui,addOnMenu);

  addOnMenu
    .addSeparator()
    .addItem('Aide', 'not implemented yet')
    .addToUi();  
}