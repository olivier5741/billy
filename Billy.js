//const MODULE_KEYS = ["me","members","stock","planning","customers","quoting","invoicing"]

const MODULES = [MeApp];

class Party {
  constructor() {
    this.name = "";
    this.nameTitle = "";
    this.phoneNumber = "";
    this.email = "";
    this.addressLine1 = "";
    this.addressLine2 = "";
    this.addressLine3 = "";
    this.bankAccountNumber = "";
    this.vatNumber = "";
  }
}

function initialiseEnvironment(){

  const rootFolder = getRootFolder()

  for(m of MODULES){

    const fileName = t(`app.${m.key}.spreadsheet.name`)
    const files = rootFolder.getFilesByName(fileName)

    if(files.hasNext() == false){

      const ss = SpreadsheetApp.create(fileName);
      m.initializeSpreadsheet(ss);
      DriveApp.getFileById(ss.getId()).moveTo(rootFolder);
    }

  }
}

function getRootFolder(){
  const rootFolder = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()).getParents().next()
  return rootFolder
}