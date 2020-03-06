/*
 * CONFIGURAZIONE
 */
var debug = false;
var sheetID = "1O-5Ren-h9xJy04rYxQLGSf7dMmGYhgYWsWhk459rSVs";
var templateIdHW = "1Q6P7OovyFC3hPMVY_UT5o3-vCTvnIkhwOArSAL0XQ-0";
var templateIdSW = "1OL23gmlnvr4ZSEgZwo6eNfrtWB-4Y-1UXY5bs41rchw";
var templateIdOrder = "1q_Y1vMNIYHBrne-6irggMVtZqlJpBq8lqGf8K-HEvuU";
var modalHeight = 600;
/** Tipologia di esecuzione - "web" o "form" */
var launchType = "form";
// SEGNAPOSTO
var NR_ORDINE = "<NR_ORDINE>";
var NOME_RIFERIMENTO = "<NOME_RIFERIMENTO>";
var EMAIL_RIFERIMENTO = "<EMAIL_RIFERIMENTO>";
var RAGIONE_SOCIALE = "<RAGIONE_SOCIALE>";
var COMUNE = "<COMUNE>";
var PROVINCIA = "<PROVINCIA>";
var INDIRIZZO = "<INDIRIZZO>";
var CAP = "<CAP>";
var PIVA = "<PIVA>";
var DATA = "<DATA>";
// RIFERIMENTI CELLE SHEET ORDINE
var CELL_ADDRESS = "D7";
var CELL_DATA_1 = "G14";
var CELL_DATA_2 = "E38";
var CELL_ORDER_NR_1 = "A14";
var CELL_ORDER_NR_2 = "C14";
/*
 * FINE CONFIGURAZIONE
 */
/** Inizializzazione Frontend */
function onOpen(e) {
    DocumentApp.getUi()
        .createMenu("-Offerte OSD-")
        .addItem("Nuova Offerta", "showModalOfferte")
        .addItem("Nuovo Ordine", "showModalOrdini")
        .addToUi();
}
/** Inizializzazione */
function onInstall(e) {
    onOpen(e);
}
/** Per la distribuzione come webapp */
function doGet(e) {
    launchType = "web";
    return showHTMLOfferte(e);
}
/** Per la distribuzione come webapp */
function doPost(e) {
    launchType = "web";
    return showHTMLOfferte(e);
}
/** Mostra interfaccia HTML */
function showModalOfferte() {
    // APRE DIALOG HTML
    var html = HtmlService.createTemplateFromFile('Frontend');
    var template = html.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setHeight(modalHeight);
    DocumentApp.getUi().showModalDialog(template, 'Crea nuova Offerta');
}
/** Mostra interfaccia HTML */
function showModalOrdini() {
    // APRE DIALOG HTML
    var html = HtmlService.createTemplateFromFile('FrontendOrdine');
    var template = html.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setHeight(modalHeight);
    DocumentApp.getUi().showModalDialog(template, 'Crea nuovo Ordine');
}
/** Mostra interfaccia HTML */
function showHTMLOfferte(e) {
    // APRE DIALOG HTML
    var html = HtmlService.createTemplateFromFile('Frontend');
    var template = html.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setHeight(modalHeight)
        .setTitle('Crea nuova Offerta Olivetti Scuola Digitale');
    return HtmlService.createHtmlOutput(template);
}
/** Crea nuovo ordine */
function CreaOrdine(datiInput) {
    Logger.log("CreaOrdine: " + JSON.stringify(datiInput));
    var templateId = templateIdOrder;
    var templateDoc = DriveApp.getFileById(templateId);
    var date = new Date();
    var month = date.getMonth() + 1;
    var day = date.getDate();
    var year = date.getFullYear();
    var dateString = "" + day + "/" + month + "/" + year;
    Logger.log("Data corrente: " + dateString);
    var orderFullName = datiInput.orderName;
    var orderAddress = datiInput.orderAddress;
    // Crea un nuovo documento dal template e sostituisce i dati
    var newDoc = templateDoc.makeCopy(orderFullName);
    var newDocId = newDoc.getId();
    var file = SpreadsheetApp.openById(newDocId);
    var sheet = file.getSheets()[0];
    Logger.log("Inizio sostituzione datiInput: " + JSON.stringify(datiInput));
    cellReplaceText(sheet, CELL_ADDRESS, INDIRIZZO, orderAddress);
    cellReplaceText(sheet, CELL_ORDER_NR_1, NR_ORDINE, orderFullName);
    cellReplaceText(sheet, CELL_ORDER_NR_2, NR_ORDINE, orderFullName);
    cellReplaceText(sheet, CELL_DATA_1, DATA, dateString);
    cellReplaceText(sheet, CELL_DATA_1, DATA, dateString);
    Logger.log("Segnaposto sostituiti");
    // Mostro nuovamente l'interfaccia
    doGet(null);
}
function cellReplaceText(sheet, cell, template, replacement) {
    Logger.log("Sostituzione nella cella " + cell + " di " + template + " con " + replacement);
    var dateCell = sheet.getRange(cell);
    var dateCellValue = dateCell.getDisplayValues();
    var dateCellContent = dateCell.getDisplayValue().replace(template, replacement);
    sheet.getRange(cell).setValue(dateCellContent);
}
/** Crea nuova offerta */
function CreaOfferta(datiInput) {
    Logger.log("CreaOfferta: " + datiInput);
    // Seleziono il template in base alla tipologia di ordine
    var templateId;
    if (datiInput.orderType.toString() == "hw") {
        templateId = templateIdHW;
    }
    else {
        templateId = templateIdSW;
    }
    Logger.log("Template selezionato: " + templateId + "[" + JSON.stringify(datiInput.orderType) + "]");
    // TEMPLATE
    var templateDoc = DriveApp.getFileById(templateId);
    var lastOrderNr = getLastOrderNumber();
    var orderNumber = Number(lastOrderNr) + 1;
    var dateString = GetDateString();
    var orderFullName = orderNumber + "_" + dateString + "_" + datiInput.orderName;
    // Crea un nuovo documento dal template e sostituisce i dati
    var newDoc = templateDoc.makeCopy(orderFullName);
    var newDocId = newDoc.getId();
    var doc = DocumentApp.openById(newDocId);
    var body = doc.getBody();
    Logger.log("datiInput: " + JSON.stringify(datiInput));
    body.replaceText(NR_ORDINE, orderFullName);
    body.replaceText(NOME_RIFERIMENTO, datiInput.nomeRiferimento);
    body.replaceText(EMAIL_RIFERIMENTO, datiInput.emailRiferimento);
    body.replaceText(RAGIONE_SOCIALE, datiInput.ragioneSociale);
    body.replaceText(COMUNE, datiInput.comune);
    body.replaceText(CAP, datiInput.cap);
    body.replaceText(PROVINCIA, datiInput.prov);
    body.replaceText(INDIRIZZO, datiInput.indirizzo);
    body.replaceText(PIVA, datiInput.pIva);
    body.replaceText(DATA, new Date().toLocaleDateString("it"));
    Logger.log("Segnaposto sostituiti");
    InsertOrder(orderNumber, orderFullName, datiInput.orderType.toUpperCase(), datiInput.valore, datiInput.ragioneSociale);
    Logger.log("Ordine inserito");
    // Mostro nuovamente l'interfaccia
    doGet(null);
}
/** Recupera dallo Spreadsheet l'ultimo ordine creato */
function getLastOrderNumber() {
    var sheet = SpreadsheetApp.openById(sheetID);
    var data = sheet.getDataRange().getValues();
    if (debug)
        Logger.log("L'ultimo ordine è: " + data[data.length - 1]);
    return data[data.length - 1][0];
}
/** Inserisce l'ordine nello Spreadsheet degli ordini */
function InsertOrder(OrderNr, OrderName, OrderType, OrderValue, ClientName) {
    var ss = SpreadsheetApp.openById(sheetID);
    var sheet = ss.getSheets()[0];
    sheet.appendRow([OrderNr, OrderName, OrderType, OrderValue, ClientName]);
}
/** Chiede all'utente di inserire una stringa */
function promptForString(question, hint, canBeEmpty) {
    if (canBeEmpty === void 0) { canBeEmpty = true; }
    var ui = DocumentApp.getUi();
    var response = ui.prompt(question, hint, ui.ButtonSet.OK);
    while (response.getResponseText() == "" && !canBeEmpty) {
        ui.alert("Questo campo non può essere vuoto!", ui.ButtonSet.OK);
        response = ui.prompt(question, hint, ui.ButtonSet.OK);
    }
    return response.getResponseText();
}
/** Restituisce la data sotto forma di timestamp */
function GetDateString() {
    Date.prototype.yyyymmdd = function () {
        var mm = this.getMonth() + 1; // getMonth() is zero-based
        var dd = this.getDate();
        return [this.getFullYear(),
            (mm > 9 ? '' : '0') + mm,
            (dd > 9 ? '' : '0') + dd
        ].join('');
    };
    var date = new Date();
    return date.yyyymmdd();
}
/** Include un file nel template */
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
}
