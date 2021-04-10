const StockApp = (function() { 

var self ={}

self.key = "stock"
self.movementKeys = ["entry","output","inventory"]

self.movementFunctions = {
  entry: (sv,v) => sv + v,
  output: (sv,v) => sv - v,
  inventory: (sv,v) => v == "" ? sv : v
}

self.createMovementSheet = function(dto){
  const movementName = t("app.stock.sheet.name." + dto.key);
  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = spreadSheet.getSheetByName(t("app.stock.sheet.name.stock"));

  const insSheet = stockSheet.copyTo(spreadSheet);
  const sheetName = Utilities.formatDate(new Date(), "GMT+1", "yyyy-MM-dd") + " " + movementName
  insSheet.setName(sheetName);
  insSheet.getRange("C2:C").clear();
  insSheet.activate();
}

self.applyMovementSheet = function (){
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const stockSheet = ss.getSheetByName(t("app.stock.sheet.name.stock"));
  const sheet = ss.getActiveSheet()
  const sheetName = sheet.getName();
  const sheetNameLastWord = sheetName.split(" ").pop();
  const key = rt("app.stock.sheet.name.",sheetNameLastWord,APP_LANGUAGE)

  self.performMovement(self.movementFunctions[key.split(".").pop()],sheet,stockSheet)
}

self.performMovement = function(action,sheet = SpreadsheetApp.getActiveSheet(),stockSheet = SpreadsheetApp.getActiveSheet()) {  
  
  const values = sheet.getRange("A2:C").getValues();
  const stockRange = stockSheet.getRange("A2:C");
  const stock = stockRange.getValues();

  let nbNewRows = 0;

  for(const i of values){
    let r = stock.find(s => s[0] == i[0])
   
    if(r == undefined){
      r = i.slice()
      stock.push(r)  
      r[2] = 0
      nbNewRows++
    }

    r[2] = action(r[2],i[2]);
  }

  if(nbNewRows > 0)
    stockSheet.insertRows(stockRange.getLastRow(),nbNewRows);
  stockSheet.getRange("A2:C").setValues(stock);

  stockSheet.activate();
  sheet.getParent().deleteSheet(sheet);
}

self.buildCard = function(){
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

  actionSection.addWidget(  //CardService.newImageButton()..setAltText(t("stock.menu.apply.movement")).setIconUrl(buildIconUrl("diff-added"))
  
  CardService.newTextButton()
          .setText(t("stock.menu.apply.movement"))
          .setOnClickAction(CardService.newAction()
            .setFunctionName("pipeline")
            .setParameters(
              {
                moduleKey: self.key, 
                moduleFunction: "applyMovementSheet"
              })))

  const builder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
            // .setImageUrl("https://www.gstatic.com/images/icons/material/system/1x/pets_black_48dp.png")
             .setTitle(t("stock.menu.module.name")))
             
    .addSection(actionSection)

  return builder.build();
}

return self

})()
