function test1(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const template = ss.getSheetByName("template invoice draft")
  TemplateApp.createSheetFromTemplate(ss,template,{})
}

function blockValuesInSheet(){
  const sheet = SpreadsheetApp.getActiveSheet();
  TemplateApp.block(sheet)
}

function unflatten(data) {
    "use strict";
    if (Object(data) !== data || Array.isArray(data))
        return data;
    var result = {}, cur, prop, parts, idx;
    for(var p in data) {
        cur = result, prop = "";
        parts = p.split(".");
        for(var i=0; i<parts.length; i++) {
            idx = !isNaN(parseInt(parts[i]));
            cur = cur[prop] || (cur[prop] = (idx ? [] : {}));
            prop = parts[i];
        }
        cur[prop] = data[p];
    }
    return result[""];
}

function flatten(data) {
    var result = {};
    function recurse (cur, prop) {
        if (Object(cur) !== cur) {
            result[prop] = cur;
        } else if (Array.isArray(cur)) {
             for(var i=0, l=cur.length; i<l; i++)
                 recurse(cur[i], prop ? prop+"."+i : ""+i);
            if (l == 0)
                result[prop] = [];
        } else {
            var isEmpty = true;
            for (var p in cur) {
                isEmpty = false;
                recurse(cur[p], prop ? prop+"."+p : p);
            }
            if (isEmpty)
                result[prop] = {};
        }
    }
    recurse(data, "");
    return result;
}

function test2(){
  const o = {
    sheet: {
      name: "The sheet name"
    },
    buyer: {
      name: "Olivier Wouters"
    }
  }

  const a =  unflatten(flatten(o));

  console.log(a)
}



const TemplateApp = (function(){
  const self = {}

  /**
   * Get the template sheet metadata
   * 
   * @param {SpreadsheetApp.Sheet} sheet the sheet
   * @return {object}
   */
  self.getData = function(sheet){
    // C2 to C5

    return new SheetData();
  }

  self.setFromData = function(sheet = SpreadsheetApp.getActiveSheet(),data){
    const sheetMetadata = sheet.getRange("A2:B5").getValues();
    const sheetMetadataTransposed = transpose(sheetMetadata)

    const keys = sheetMetadataTransposed[0];
    const froms = sheetMetadataTransposed[1];
  
    const flatData = flatten(data);

    for(let [i, k] of keys.entries()){
      if(k in flatData){
        froms[i] = flatData[k]
      }
    }

    Logger.log(froms)
    sheet.getRange("B2:B5").setValues(transpose([froms]))
    //const properties = 

  }

  self.block = function(sheet = SpreadsheetApp.getActiveSheet()){
    const toBlockValues = sheet.getRange("A7:C8").getValues()
    const toBlock = [].concat(...toBlockValues).filter(v => v)

    console.log(toBlock)

    const rangeListToBlock = sheet.getRangeList(toBlock)
    rangeListToBlock.getRanges()
      .forEach(r => r.setValues(r.getValues()))

  }

  /**
   * Creates a sheet by applying data to template sheet
   *  
   * @param {SpreadsheetApp.Spreadsheet} spreadsheet the spreadsheet where to create the sheet
   * @param {SpreadsheetApp.Sheet} templateSheet the template sheet
   * @param {object} data the data used by the template
   * @return {SpreadsheetApp.Sheet}
   */
  self.createSheetFromTemplate = function(spreadsheet,templateSheet,data){

    const sheet = templateSheet.copyTo(spreadsheet);    
    self.setFromData(sheet,{buyer:{name:"Sébastien"}})

    const sheetData = self.getData();

    sheet.setName(sheetData.sheet.name);


    return sheet;
  }

  class SheetData {
    constructor(){
      this.sheet = {
          name: "Olivier Wouters"        
      }
    }
  }

  return self
})()

const BillyCustomerModule = (function(){
  const self = {}

  

  return self
})()

// seller
// buyer
// products
// services
class Invoice {
  constructor(){
    this.name = "test"
    this.time = new Date()
    this.seller = new Party()
    this.buyer = new Party()
  }
}

class InvoiceModule{
  constructor(spreadsheet){
    this.key = "invoice"
    this.spreadsheet = spreadsheet
  }

  /**
   * Creates an invoice sheet.
   *  
   * @param {Invoice} the invoice entity
   * @param {Party} buyer the invoice buyer
   * @param {string} templateId the id of the template sheet
   * @return {void}
   */
  createInvoice(invoice,templateId){

    if(customerSelected == false){
      const ui = SpreadsheetApp.getUi();
      const response = ui.prompt("Créer une facture", "Entrer le nom du client (ex: Durant Marc).\nTu peux également lancer cette action en sélectionnant le ligne du client dans la fiche _clients (cancel l'action en cours).", ui.ButtonSet.OK_CANCEL);
      
      if (response.getSelectedButton() == ui.Button.CANCEL){
        return;
      }
      
      customer = new Customer(response.getResponseText());
    }

    const template = this.spreadsheet.getSheetByName(billTemplateSheetName);
    const sheet = template.copyTo(this.spreadsheet);
    const sheetName = Utilities.formatDate(new Date(), "GMT+1", "yyyy-MM-dd") + " facture - " + customer.name;
    sheet.setName(sheetName);

    sheet.getRange("B9").setValue(Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yyyy"));
    sheet.getRange("G1").setValue("???/" + (new Date()).getFullYear());

    if(customer.tvaNumber != "")
      sheet.getRange("B11").setValue(customer.tvaNumber); 

    const customerPostInfo = transpose(
      [[customer.getNameWithTitle(),customer.addressLine1,customer.addressLine2]]);
    sheet.getRange("F8:F10").setValues(customerPostInfo);

    if(invoiceType == "Basic"){
      sheet.deleteRows(59,sheet.getMaxRows()-59+1);
      sheet.deleteRow(35);
    }

    if(invoiceType == "WithItems"){
      
    }

    sheet.activate();
    this.spreadsheet.moveActiveSheet(1)
  }

  /**
   * Creates an invoice.
   *  
   * @param {id} the id of the invoice
   * @return {void}
   */
  archiveInvoice(id){

  }
}

const InvoiceModuleApp = (function() { 

var self ={}

/**
 * Get invoice module
 *  
 * @param {id} the id of spreadsheet
 * @return {InvoiceModule}
 */
self.getModule = function(id){
  return new InvoiceModule()
}

/**
 * Creates a bill sheet.
 *  
 * @param {number} billType The type of the bill (Basic or WithItems)
 * @return {Test}
 */
self.createInvoice = function(customer, invoiceType, spreadSheetId){

  

}

self.newTest = () => new Test()

return self

})()


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
