const FILE = SpreadsheetApp.getActiveSpreadsheet();
const WORKSHEET = FILE.getSheetByName("FORM INPUT"); //Form Sheet
const REGISTERS = FILE.getSheetByName("REGISTERS");
const RECEIPT = FILE.getSheetByName("RECEIPT");
const LAST_REGISTERS_ROW = REGISTERS.getLastRow();
const LAST_REGISTERS_COL = REGISTERS.getLastColumn();

const REG_COL = 2; //second column reserved for registers number
const REG_NUMBERING = 1; // first column is reserved for numbering increment
const RECEIPT_REGISTER_RANGE = 'A6:E6';

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('Riksa III')
      .addItem('Overstay receipt', 'osReceipt')
      .addSeparator()
      .addSubMenu(ui.createMenu('Sub-menu')
        .addItem('Second item', 'menuItem2'))
      .addToUi();
}

function receive(formData) {
  assignRegisterNumber(formData);
  // return value will be passed to client side onSuccess function
  // return value can be anything, since it will be ignored
  return true;
}

function osReceipt() {
  // To include other file in html files, use evaluate() method first.
  // This evaluates the template and returns an HtmlOutput object.
  // Any properties set on this HtmlTemplate object will be in scope when evaluating.
  // To debug errors in a template, examine the code using the getCode() method.
  var output = HtmlService.createTemplateFromFile("receiptForm")
    .evaluate()
    .setWidth(500)
    .setHeight(650);
  
  // Opens a modal dialog box in the user's editor with custom client-side content.
  // This method does not suspend the server-side script while the dialog is open.
  // To communicate with the server-side script, the client-side component must make asynchronous callbacks using the google.script API
  // for HtmlService.

  // To close the dialog programmatically, call google.script.host.close() on the client side of an HtmlService web app.

  SpreadsheetApp.getUi()
    .showModalDialog(output, "OVERSTAY RECEIPT");
}

function include(filename) {
  // Creates a new HtmlOutput object from a file in the code editor.
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function assignRegisterNumber( formData ) {
  
  const lastRegister = getLastRegister();
  
  const currentRegister = {};
  // if last register row is 1, set initial numbering to 1. Last row = 1, because the sheet has header
  if (LAST_REGISTERS_ROW === 1) {
    currentRegister.numbering = 1;
    currentRegister.register = currentRegister.numbering.toString().padStart(2, 0);
  } else {
    currentRegister.numbering = ++lastRegister.numbering;
    currentRegister.register = currentRegister.numbering.toString().padStart(2, 0);
  }

  let currentYear = Utilities.formatDate(new Date(), "GMT+08", "yyyy");

  // Formats date according to specification described in Java SE SimpleDateFormat class.
  // Please visit the specification at http://docs.oracle.com/javase/7/docs/api/java/text/SimpleDateFormat.html
  let currentDate = Utilities.formatDate(new Date(), "GMT+08", "dd-MM-yyyy");
  
  currentRegister.register = `NR/${currentRegister.register}/III/${currentYear}/TPI.III/`;
  setRegister(currentRegister, formData, currentDate);
}

function setRegister({ register, numbering}, { masaOverstay, nomorPesawat, officers, kodeNegara, nomorPaspor, nama }, date) {
  RECEIPT.getRange(RECEIPT_REGISTER_RANGE).setValue(register);
  REGISTERS.appendRow([numbering, register, kodeNegara.toUpperCase(), nomorPaspor.toUpperCase(), nomorPesawat.toUpperCase(), nama.toUpperCase(), masaOverstay, date, officers.toUpperCase()]);
  
}

const getLastRegister = () => {
  return {
    row: LAST_REGISTERS_ROW,
    numbering: REGISTERS.getRange(LAST_REGISTERS_ROW, REG_NUMBERING)
      .getValue(),
    register: REGISTERS.getRange(LAST_REGISTERS_ROW, REG_COL)
      .getValue(),
  };
}

function getOfficers() {
  let data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('officers in charge').getDataRange().getValues();
  return data;
}

function getCountries() {
  let data = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('country iso code list').getDataRange().getValues();
  return data;
}
