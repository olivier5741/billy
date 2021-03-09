const MeApp = (function() { 

var self ={}

self.key = "me"

self.initializeSpreadsheet = function(spreadsheet = SpreadsheetApp.getActiveSpreadsheet()) {

  const properties = Object.getOwnPropertyNames(new Party)
  const translations = properties.map(p => t(`app.object.party.${p}.field.name`))

  const sheet = spreadsheet.getActiveSheet()

  sheet.setName(t("app.me.data.sheet.name"))
  
  sheet.getRange(1,1,translations.length,1).setValues(transpose([translations]))
  
  trimAroundRange(sheet.getRange(1,1,translations.length,2))
  sheet.autoResizeColumn(1);
  sheet.setColumnWidths(2,1,400)

}; 

return self

})()