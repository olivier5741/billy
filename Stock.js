const stockSheetName = "Stock"

function createStockMenu(ui,addOnMenu){

  loadTranslations();

  addOnMenu
    .addSubMenu(ui.createMenu(t("stock.menu.module.name"))
      .addItem(t("stock.menu.create.entry"), "createInSheet")
      .addItem(t("stock.menu.create.output"), "createOutSheet")
      .addItem(t("stock.menu.create.inventory"), "createInventorySheet")
      .addSeparator()
      .addItem(t("stock.menu.apply.movement"), "applyMovementToStock")
    );
}



function applyMovementToStock(sheet = SpreadsheetApp.getActiveSheet()){
  const lastWord = sheet.getName().split(" ").pop();


  createMovementSheet("entrée");
}


function createInSheet(){
  createMovementSheet(t("app.stock.sheet.name.entry"));
}

function createOutSheet(){
  createMovementSheet(t("app.stock.sheet.name.output"));
}

function createInventorySheet(){
  createMovementSheet(t("app.stock.sheet.name.inventory"));
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
