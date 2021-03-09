function columnToLetter(column)
{
  let temp, letter = '';
  while (column > 0)
  {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function protectFormulaRangeWithWarning(range){
  const p = range.protect();
  p.setDescription("Formules. Merci de ne pas modifier.");
  p.setWarningOnly(true);
}

function sortSheetsByName() {
  const aSheets = new Array();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (let s in ss.getSheets())
  {
    aSheets.push(ss.getSheets()[s].getName());
  }
  if(aSheets.length)
  {
    const collator = new Intl.Collator(undefined, {numeric: true, sensitivity: 'base'});
    aSheets.sort(collator.compare);
    for (let i = 0; i < aSheets.length; i++)
    {
      const theSheet = ss.getSheetByName(aSheets[i]);
      if(theSheet.getIndex() != i + 1)
      {
        ss.setActiveSheet(theSheet);
        ss.moveActiveSheet(i + 1);
      }
    }
  }
}

function copySheetRangeProtectionWarnings(sheet, sheet2) { 
  const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (let i = 0; i < protections.length; i++) {
    const p = protections[i];
    const rangeNotation = p.getRange().getA1Notation();
    const p2 = sheet2.getRange(rangeNotation).protect();
    p2.setDescription(p.getDescription());
    p2.setWarningOnly(p.isWarningOnly());
  }  
}

function convertSheetToPdf(spreadsheet, sheet, pdfName, lastColumn = columnToLetter(spreadsheet.getLastColumn()),portrait = false) {
  
  const spreadsheetId = spreadsheet.getId()  
  const sheetId = sheet != null ? sheet.getSheetId() : undefined;  
  pdfName = pdfName ? pdfName : spreadsheet.getName();
  const url_base = "https://docs.google.com/spreadsheets/d/" + spreadsheet.getId() + "/"

  const url_ext = 'export?exportFormat=pdf&format=pdf'   //export as pdf

      // Print either the entire Spreadsheet or the specified sheet if optSheetId is provided
      + (sheetId ? ('&gid=' + sheetId) : ('&id=' + spreadsheetId)) 
      // following parameters are optional...
      + '&size=A4'      // paper size
      + '&portrait=' + portrait    // orientation, false for landscape
      + '&fitw=true'        // fit to width, false for actual size
      + '&top_margin=0.5'              
      + '&bottom_margin=0.5'          
      + '&left_margin=0.5'             
      + '&right_margin=0.5'          
      + '&sheetnames=false&printtitle=false&pagenumbers=false'  //hide optional headers and footers
      + '&gridlines=false'  // hide gridlines
      + '&fzr=false'       // do not repeat row headers (frozen rows) on each page
      + '&printnotes=false'
      + (sheetId ? '&range=a1:'+ lastColumn + spreadsheet.getLastRow() : '');      // range to export TODO more than n if more than 14 days

  const options = {
    headers: {
      'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken(),
    }
  }
  
  const response = UrlFetchApp.fetch(url_base + url_ext, options);
  return response.getBlob().setName(pdfName + '.pdf');
}

function getFoldersByNameOrCreate(mainFolder, name){
  var path = mainFolder.getFoldersByName(name);
  return path.hasNext() ? 
    path.next() : mainFolder.createFolder(name);
}