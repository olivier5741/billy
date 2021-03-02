const stockSheetName = "Stock"

function createStockMenu(ui,addOnMenu){
  addOnMenu
    .addSubMenu(ui.createMenu("Stock")
      .addItem("Créer une entrée", "createInSheet")
      .addItem("Créer une sortie", "createOutSheet")
      .addItem("Créer un inventaire", "createInventorySheet")
      .addSeparator()
      .addItem("Ajouter l'entrée au stock", "addInsToStock")
      .addItem("Soustraire la sortie au stock", "substractOutsToStock")
      .addItem("Appliquer l'inventaire au stock", "setInventoryToStock")
    );
}


/*

function applyMovementToStock(sheet = SpreadsheetApp.getActiveSheet()){
  const lastWord = sheet.getName().split(" ").pop();

  switch()
  createMovementSheet("entrée");
}
*/

function createInSheet(){
  createMovementSheet("entrée");
}

function createOutSheet(){
  createMovementSheet("sortie");
}

function createInventorySheet(){
  createMovementSheet("inventaire");
}

function createMovementSheet(movementName){
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = spreadSheet.getSheetByName(stockSheetName);

  const insSheet = stockSheet.copyTo(spreadSheet);
  insSheet.setName(Utilities.formatDate(new Date(), "GMT+1", "yyyy-MM-dd") + " " + movementName);
  insSheet.getRange("C2:C").clear();
  insSheet.activate();
}

function addInsToStock() {
  performMovementOnStock((sv,v) => sv + v);
}

function substractOutsToStock() {
  performMovementOnStock((sv,v) => sv - v);
}

function setInventoryToStock() {
  performMovementOnStock((sv,v) => v == "" ? sv : v);
}

function performMovementOnStock(
  action,
  sheet = SpreadsheetApp.getActiveSheet(),
  stockSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(stockSheetName)) {

  const values = sheet.getRange("A2:C").getValues();
  const stockRange = stockSheet.getRange("A2:C");
  const stock = stockRange.getValues();

  for(const i of values){
    const r = stock.find(s => s[0] == i[0])
    r[2] = action(r[2],i[2]);
  }

  stockRange.setValues(stock);

  stockSheet.activate();
  sheet.getParent().deleteSheet(sheet);
}
