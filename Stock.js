const StockApp = (function() { 

var self ={}

self.key = "stock"
self.movementKeys = ["entry","output","inventory"]

self.movementFunctions = {
  entry: (sv,v) => sv + v,
  output: (sv,v) => sv - v,
  inventory: (sv,v) => v == "" ? sv : v
}

self.getMovementSheetKeys = function(ss){
  
  const sheetNames = getAllSheetNames(ss);

  return sheetNames.filter(s => s.startsWith("2")); // TODO hack
}

self.createMovementSheet = function(dto){
  const movementName = t("app.stock.sheet.name." + dto.key);
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const movements = self.getMovementSheetKeys(spreadSheet)
  const stockSheet = spreadSheet.getSheetByName(t("app.stock.sheet.name.stock"));

  const insSheet = stockSheet.copyTo(spreadSheet);
  const sheetName = Utilities.formatDate(new Date(), "GMT+1", "yyyy-MM-dd") + " " + movementName
  insSheet.setName(sheetName);
  insSheet.getRange("C2:C").clear();
  insSheet.activate();
  
  movements.push(sheetName)

  return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().updateCard(self.buildCard(movements)))
        .build();
}

self.applyMovementSheet = function (dto){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const movements = self.getMovementSheetKeys(ss)

  const stockSheet = ss.getSheetByName(t("app.stock.sheet.name.stock"));
  const sheet = ss.getSheetByName(dto.key);
  const sheetName = sheet.getName();
  const sheetNameLastWord = sheetName.split(" ").pop();
  const key = rt("app.stock.sheet.name.",sheetNameLastWord,APP_LANGUAGE)

  self.performMovement(self.movementFunctions[key.split(".").pop()],sheet,stockSheet)
  
  return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation().updateCard(self.buildCard(movements.filter(m => m != sheetName))))
        .build();
}

self.performMovement = function(action,sheet,stockSheet) {

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

self.buildCard = function(movements){
  const actionSection = CardService.newCardSection()

  for(const k of self.movementKeys){
    actionSection.addWidget(
        CardService.newTextButton()
          .setText(t("stock.menu.create." + k))
          .setOnClickAction(CardService.newAction()
            .setFunctionName("pipeline")
            .setParameters(
              {
                moduleKey: self.key, 
                moduleFunction: "createMovementSheet", 
                dtoKey: k
              })))
  } 

  const movementSection = CardService.newCardSection();

  for(const m of movements){
    movementSection.addWidget(
      CardService.newKeyValue()
        .setContent(m)
        .setOnClickAction(CardService.newAction().setFunctionName("test"))
        .setButton(CardService.newTextButton()
          .setText("Apply") // "stock.menu.apply.movement"
          .setOnClickAction(
            CardService.newAction()
            .setFunctionName("pipeline")
            .setParameters(
              {
                moduleKey: self.key, 
                moduleFunction: "applyMovementSheet", 
                dtoKey: m
              })))
          )
  }

  const builder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
             .setTitle(t("stock.menu.module.name")))
    .addSection(actionSection)
  
  if(movements.length > 0)
    builder.addSection(movementSection)

  return builder.build();
}

return self

})()
