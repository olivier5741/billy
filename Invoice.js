// TODO rename bill to invoice
const billTemplateSheetName = "zz_template_facture"
const billItemsTemplateSheetName = "zz_template_facture_pièces"
const customersSheetName = "zz_clients";

function createInvoiceMenu(ui,addOnMenu){
  addOnMenu
    .addSubMenu(ui.createMenu("Facturation")
      .addSubMenu(ui.createMenu("Créer une facture")
          .addItem("simple", "createBillBasic")
          .addItem("avec pièces", "createBillWithItems"))
        .addItem("Imprimer la facture via PDF","createMultipagePDF")
        .addItem("Archiver la facture","archiveBill")
        .addSeparator()
        .addItem("Trier les onglets alph.","sortSheetsByName")
    )
}

class Customer {
  constructor(name,title="",addressLine1="",addressLine2="",tvaNumber="") {
    this.name = name,
    this.title = title,
    this.addressLine1 = addressLine1,
    this.addressLine2 = addressLine2,
    this.tvaNumber = tvaNumber
  }

  getNameWithTitle(){
    if(this.title != "")
      return this.title + " " + this.name;
    return this.name;
  }
}

function createBillWithItems(){
  createBill("WithItems");
}

function createBillBasic(){
  createBill("Basic");
}

/**
 * Creates a bill sheet.
 *  
 * @param {number} billType The type of the bill (Basic or WithItems)
 * @return {void}
 */
function createBill(billType){

  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const currentRange = spreadSheet.getActiveRange();

  const customerSelected = currentRange.getSheet().getName() == customersSheetName 
    && currentRange.getNumRows() == 1 
    && currentRange.getNumColumns() >= 5;

  let customer;

  if(customerSelected){
    const row = currentRange.getValues()[0];
    customer = new Customer(row[0],row[1],row[2],row[3],row[4]);
  }

  if(customerSelected == false){
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt("Créer une facture", "Entrer le nom du client (ex: Durant Marc).\nTu peux également lancer cette action en sélectionnant le ligne du client dans la fiche _clients (cancel l'action en cours).", ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() == ui.Button.CANCEL){
      return;
    }
    
    customer = new Customer(response.getResponseText());
  }

  const template = spreadSheet.getSheetByName(billTemplateSheetName);
  const sheet = template.copyTo(spreadSheet);
  const sheetName = Utilities.formatDate(new Date(), "GMT+1", "yyyy-MM-dd") + " facture - " + customer.name;
  sheet.setName(sheetName);

  sheet.getRange("B9").setValue(Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy"));
  sheet.getRange("G1").setValue("???/" + (new Date()).getFullYear());

  if(customer.tvaNumber != "")
    sheet.getRange("B11").setValue(customer.tvaNumber); 

  const customerPostInfo = transpose(
    [[customer.getNameWithTitle(),customer.addressLine1,customer.addressLine2]]);
  sheet.getRange("F8:F10").setValues(customerPostInfo);

  if(billType == "Basic"){
    sheet.deleteRows(59,sheet.getMaxRows()-59+1);
    sheet.deleteRow(35);
  }

  if(billType == "WithItems"){
    
  }

  sheet.activate();
  spreadSheet.moveActiveSheet(1)
}

function archiveBill(){

  const spreadSheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadSheet.getActiveSheet();
  const sheetName = sheet.getName();
  
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
     "Archiver la facture",
     "Voulez-vous archiver la facture : " + sheetName,
      ui.ButtonSet.OK_CANCEL);

  if (result == ui.Button.CANCEL) {
    return;
  }
  
  const folder = DriveApp.getFileById(spreadSheet.getId()).getParents().next();
  const archiveFolder = getFoldersByNameOrCreate(folder,"Archives");

  const file = createMultipagePDF();
  file.moveTo(archiveFolder);  
  
  spreadSheet.deleteSheet(sheet);
}

function insertSheetWithName(spreadSheet,name){
    const sheet = spreadSheet.insertSheet();
    sheet.setName(name);
    return sheet;
}

function copyRangeValuesFormatAndColumnWidth(source,destination){
  destination.setValues(source.getValues());
  source.copyTo(destination,SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  source.copyTo(destination,SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
}

function findPageBreaks(sheet = SpreadsheetApp.getActiveSheet()){
  const pattern = "/saut_de_page";

  const arr = transpose(sheet.getRange("A:A").getValues())[0];
  const keys = [];
  arr.filter((e, i) => {
    if (e == pattern) {
      keys.push(i);
    }
  });

  return keys;
}

function createMultipagePDF(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getActiveSheet();
  const printFolderName = "À imprimer";
  const lastColumnLetter = "H";

  const rootFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  const printFolder = getFoldersByNameOrCreate(rootFolder,printFolderName);

  const printSpreadsheet = SpreadsheetApp.create(s.getName());
  const dummyPrintSheet = printSpreadsheet.getActiveSheet();
  const printFile = DriveApp.getFileById(printSpreadsheet.getId());
  printFile.moveTo(printFolder);

  const pageBreaks = findPageBreaks(s);
  const maxRow = s.getMaxRows();

  for(let i = 0; i < pageBreaks.length + 1; i++){
    const billSheet = insertSheetWithName(ss,"print_temporary_facture_" + i);

    const firstRow = i == 0 ? 1 : pageBreaks[i-1] + 2;
    const lastRow = pageBreaks.length > i ? pageBreaks[i] : maxRow;

    const sourceRange = s.getRange("A"+ firstRow +":" + lastColumnLetter + lastRow);

    copyRangeValuesFormatAndColumnWidth(sourceRange,billSheet.getRange("A1:" + lastColumnLetter + sourceRange.getNumRows()));

    billSheet.copyTo(printSpreadsheet);

     ss.deleteSheet(billSheet);
  }
  
  printSpreadsheet.deleteSheet(dummyPrintSheet); // TODO wrong name in french

  const pdf = convertSheetToPdf(printSpreadsheet,undefined,undefined,undefined,true);  
  const pdfFile = printFolder.createFile(pdf);
  
  printFile.setTrashed(true);

  return pdfFile;
}
