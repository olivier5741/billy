//const MODULE_KEYS = ["me","members","stock","planning","customers","quoting","invoicing"]

class BillyInstance {
  constructor(folderId, name){
    this.id = folderId
    this.name = name
  }
}

class BillyInstanceModule {
  constructor(fileId, name){
    this.id = fileId
    this.name = name
  }
}

const BillyApp = (function() { 

var self ={}

self.key = "billy"
self.moduleKeys = ["me","members","stock","planning","customers","quoting","invoicing"]
self.moduleKeysTranslated = [];

self.modules = {
  stock: StockApp
}

self.isFolderAnInstance = function(name){
  const endPattern = "- billy"
  return name.endsWith(endPattern)
}

self.getModuleKeysTranslated = function(){
  
  if(self.moduleKeysTranslated.length == 0){
    self.moduleKeysTranslated = self.moduleKeys.map(k => t(`app.${k}.spreadsheet.name`))
  }

  return self.moduleKeysTranslated
}

self.isSpreadsheetAModule = function(name){
  return self.getModuleKeysTranslated().includes(name)
}

self.getModuleKeyFromSpreadsheetName = function(name){
  const i = self.getModuleKeysTranslated().findIndex(t => t == name)

  return i ? self.moduleKeys[i] : undefined
}

self.getInstanceNameFromFolder = function(folderName){
  const endPattern = "- billy";
  return folderName.substring(0, folderName.length - endPattern.length );
}

self.findInstances = function(){
  
  const search = DriveApp.searchFolders("title contains 'billy'")

  const folders = []

  while (search.hasNext()) {
    const folder = search.next()
    const name = folder.getName()
    if(self.isFolderAnInstance(name))
      folders.push(new BillyInstance(folder.getId(),getInstanceNameFromFolder(name)));
  }  

  return folders;
}

self.findInstanceModules = function(instanceId){

  const files = DriveApp.getFolderById(instanceId).getFiles() // TODO what if many files

  const instanceModules = []

  while (files.hasNext()) {
    const file = files.next()
    const name = file.getName()
   // if(name.endsWith(endPattern))
    if(self.isSpreadsheetAModule(name))
      instanceModules.push(new BillyInstanceModule(file.getId(),name))
  }  
  Logger.log(instanceModules)
  return instanceModules
}

self.buildInstancesView = function(){
  const instances = self.findInstances();
  const actionSection = CardService.newCardSection()

  for(const i of instances){
    actionSection.addWidget(
        CardService.newTextButton()
          .setText(i.name)
          .setOnClickAction(CardService.newAction()
            .setFunctionName("pipeline")
            .setParameters(
              {
                moduleKey: self.key, 
                moduleFunction: "buildInstanceView", 
                dtoId: i.id,
                dtoName: i.name
              })))
  } 

  const builder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle("Instances"))
    .addSection(actionSection)
  return builder.build()
}

self.buildInstanceView = function(dto){

  Logger.log(dto)

  const modules = self.findInstanceModules(dto.id);
  const actionSection = CardService.newCardSection()

  for(const m of modules){    

    actionSection.addWidget(CardService.newTextButton()
          .setText(m.name)
          //.setOnClickAction(action)
          .setOpenLink(CardService.newOpenLink()        
            .setUrl(`https://docs.google.com/spreadsheets/d/${m.id}`)
            .setOpenAs(CardService.OpenAs.FULL_SIZE)
            .setOnClose(CardService.OnClose.RELOAD_ADD_ON)))
  }

  const builder = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader().setTitle(dto.name))

  if(modules.length != 0)
    builder.addSection(actionSection)

  return builder.build()
}

return self

})()

function test(){

}

const BILLY = {
   billy: BillyApp,
   stock: StockApp
 }

function pipeline(e){
  const params = e.parameters;  
  return BILLY[params.moduleKey][params.moduleFunction]({
    key: params.dtoKey,
    id: params.dtoId,
    name: params.dtoName
  })
}

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