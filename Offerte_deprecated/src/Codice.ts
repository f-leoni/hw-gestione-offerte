/*
 * CONFIGURAZIONE
 */

const debug = false;
const sheetID = "1O-5Ren-h9xJy04rYxQLGSf7dMmGYhgYWsWhk459rSVs";
const templateIdHW = "1Q6P7OovyFC3hPMVY_UT5o3-vCTvnIkhwOArSAL0XQ-0";
const templateIdSW = "1OL23gmlnvr4ZSEgZwo6eNfrtWB-4Y-1UXY5bs41rchw";
const templateIdOrder = "1q_Y1vMNIYHBrne-6irggMVtZqlJpBq8lqGf8K-HEvuU";
const modalHeight = 600;
/** Tipologia di esecuzione - "web" o "form" */
let launchType = "form";
// SEGNAPOSTO
const NR_ORDINE = "<NR_ORDINE>";
const NOME_RIFERIMENTO = "<NOME_RIFERIMENTO>";
const EMAIL_RIFERIMENTO = "<EMAIL_RIFERIMENTO>";
const RAGIONE_SOCIALE = "<RAGIONE_SOCIALE>";
const COMUNE = "<COMUNE>";
const PROVINCIA = "<PROVINCIA>";
const INDIRIZZO = "<INDIRIZZO>";
const CAP = "<CAP>";
const PIVA = "<PIVA>";
const DATA = "<DATA>";
// RIFERIMENTI CELLE SHEET ORDINE
const CELL_ADDRESS = "D7";
const CELL_DATA_1 = "G14";
const CELL_DATA_2 = "E38";
const CELL_ORDER_NR_1 = "A14";
const CELL_ORDER_NR_2 = "C14";

/*
 * FINE CONFIGURAZIONE
 */

/** Inizializzazione Frontend */
function onOpen(e: any) {
    DocumentApp.getUi()
        .createMenu("-Offerte OSD-")
        .addItem("Nuova Offerta", "showModalOfferte")
        .addItem("Nuovo Ordine", "showModalOrdini")
        .addToUi();
}

/** Inizializzazione */
function onInstall(e: any) {
    onOpen(e);
}

/** Per la distribuzione come webapp */
function doGet(e: any) {
    launchType = "web";
    return showHTMLOfferte(e);
}

/** Per la distribuzione come webapp */
function doPost(e: any) {
    launchType = "web";
    return showHTMLOfferte(e);
}

/** Mostra interfaccia HTML */
function showModalOfferte() {
    // APRE DIALOG HTML
    const html = HtmlService.createTemplateFromFile('Frontend');
    const template = html.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setHeight(modalHeight);
    DocumentApp.getUi().showModalDialog(template, 'Crea nuova Offerta');
}

/** Mostra interfaccia HTML */
function showModalOrdini() {
    // APRE DIALOG HTML
    const html = HtmlService.createTemplateFromFile('FrontendOrdine');
    const template = html.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setHeight(modalHeight);
    DocumentApp.getUi().showModalDialog(template, 'Crea nuovo Ordine');
}

/** Mostra interfaccia HTML */
function showHTMLOfferte(e: any) {
    // APRE DIALOG HTML
    const html = HtmlService.createTemplateFromFile('Frontend');
    const template = html.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setHeight(modalHeight)
        .setTitle('Crea nuova Offerta Olivetti Scuola Digitale');
    return HtmlService.createHtmlOutput(template);
}

/** Crea nuovo ordine */
function CreaOrdine(datiInput: DatoOrdine) {
    Logger.log("CreaOrdine: " + JSON.stringify(datiInput));
    const templateId = templateIdOrder;
    const templateDoc = DriveApp.getFileById(templateId);
    const date = new Date();
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear();
    const dateString: string = "" + day + "/" + month + "/" + year;
    Logger.log("Data corrente: " + dateString);
    const orderFullName = datiInput.orderName;
    const orderAddress = datiInput.orderAddress;
    // Crea un nuovo documento dal template e sostituisce i dati
    const newDoc = templateDoc.makeCopy(orderFullName);
    const newDocId = newDoc.getId();
    const file = SpreadsheetApp.openById(newDocId);
    const sheet = file.getSheets()[0];

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

function cellReplaceText(sheet: GoogleAppsScript.Spreadsheet.Sheet, cell: string, template: string, replacement: string) {
    Logger.log("Sostituzione nella cella " + cell + " di " + template + " con " + replacement);
    const dateCell = sheet.getRange(cell);
    const dateCellValue = dateCell.getDisplayValues()
    const dateCellContent = dateCell.getDisplayValue().replace(template, replacement);
    sheet.getRange(cell).setValue(dateCellContent);
}

/** Crea nuova offerta */
function CreaOfferta(datiInput: DatoOfferta) {
    Logger.log("CreaOfferta: " + datiInput);
    // Seleziono il template in base alla tipologia di ordine
    let templateId;
    if (datiInput.orderType.toString() == "hw") {
        templateId = templateIdHW;
    } else {
        templateId = templateIdSW;
    }
    Logger.log("Template selezionato: " + templateId + "[" + JSON.stringify(datiInput.orderType) + "]");

    // TEMPLATE
    const templateDoc = DriveApp.getFileById(templateId);
    const lastOrderNr = getLastOrderNumber();
    const orderNumber = Number(lastOrderNr) + 1;
    const dateString = GetDateString();
    const orderFullName = orderNumber + "_" + dateString + "_" + datiInput.orderName;
    // Crea un nuovo documento dal template e sostituisce i dati
    const newDoc = templateDoc.makeCopy(orderFullName);
    const newDocId = newDoc.getId();
    const doc = DocumentApp.openById(newDocId);
    const body = doc.getBody();
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
    const sheet = SpreadsheetApp.openById(sheetID);
    const data = sheet.getDataRange().getValues();
    if (debug)
        Logger.log("L'ultimo ordine è: " + data[data.length - 1]);
    return data[data.length - 1][0];
}

/** Inserisce l'ordine nello Spreadsheet degli ordini */
function InsertOrder(
    OrderNr: Number,
    OrderName: string,
    OrderType: string,
    OrderValue: string,
    ClientName: string) {
    const ss = SpreadsheetApp.openById(sheetID);
    const sheet = ss.getSheets()[0];
    sheet.appendRow([OrderNr, OrderName, OrderType, OrderValue, ClientName]);
}

/* TOOLS */

/** Legge un valore dalla configurazione  */
function ReadConfigValue(paramName: string) {
    var value = "";
    const A = 0;
    const B = 1;
    const sheet = SpreadsheetApp.getActive().getSheetByName('Config');
    var data = sheet.getDataRange().getValues();

    for (var i = 0; i < data.length; i++) {
        if (data[i][A] == paramName) { //[1] because column B
            Logger.log((i + 1))
            value = data[i][B].toString();
        }
    }
    return value;
}

/** Chiede all'utente di inserire una stringa */
function promptForString(question: string, hint: string, canBeEmpty: boolean = true) {
    const ui = DocumentApp.getUi();
    let response = ui.prompt(question, hint, ui.ButtonSet.OK);

    while (response.getResponseText() == "" && !canBeEmpty) {
        ui.alert("Questo campo non può essere vuoto!", ui.ButtonSet.OK)
        response = ui.prompt(question, hint, ui.ButtonSet.OK);
    }
    return response.getResponseText();
}

/** Restituisce la data sotto forma di timestamp */
function GetDateString() {
    Date.prototype.yyyymmdd = function () {
        const mm = this.getMonth() + 1; // getMonth() is zero-based
        const dd = this.getDate();
        return [this.getFullYear(),
        (mm > 9 ? '' : '0') + mm,
        (dd > 9 ? '' : '0') + dd
        ].join('');
    };
    const date = new Date();
    return date.yyyymmdd();
}

/** Nuovo metodo per oggeto Date */
declare interface Date {
    yyyymmdd(): string;
}

interface DatoOrdine {
    orderName: string;
    orderAddress: string;
}

interface DatoOfferta {
    orderName: string;
    nomeRiferimento: string;
    orderFullName: string;
    emailRiferimento: string;
    ragioneSociale: string;
    comune: string;
    cap: string;
    prov: string;
    indirizzo: string;
    pIva: string;
    valore: string;
    orderType: string;
}

/** Include un file nel template */
function include(filename: string) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
}

