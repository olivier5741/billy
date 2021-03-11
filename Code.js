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
  onOpen()
}

function onOpen(e) {
}

function buildCommonHomePage(){
  return BillyApp.buildInstancesView()
}

function buildSheetsHomePage(){
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const parentFolder = DriveApp.getFileById(spreadsheet.getId()).getParents().next()

  Logger.log(spreadsheet.getName())
  
  if(BillyApp.isFolderAnInstance(parentFolder.getName()) == false)
    return BillyApp.buildInstancesView();
  
  const moduleKey = BillyApp.getModuleKeyFromSpreadsheetName(spreadsheet.getName())
  
  if(moduleKey == false || Object.keys(BillyApp.modules).includes(moduleKey) == false)
    return BillyApp.buildInstanceView({id: parentFolder.getId(), name: BillyApp.getInstanceNameFromFolder(parentFolder.getName())})    

  return BillyApp.modules[moduleKey].buildCard();
}