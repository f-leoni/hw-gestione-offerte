/*
 * CONFIGURAZIONE
 */
const debug = false;
/** ID documenti  */
const invoicesFileID = "1eqLHaMm9DFpofqGtETu8Mu7sUCak7plWHyRV37lbV0A";
const invoiceTemplateID ="1i8VS32ksIrcUvJJOrn983xDgrE5qy9sbo4KfdPRODXs";
/** altezza in px della finestra modale */
const modalHeight = 600;
/** Righe file fatture */
const firstRow = 2;
const lastRow = 70;
const iRow = 22;
// SEGNAPOSTO
const RAGIONE_SOCIALE = "<RAGIONE_SOCIALE>";
const INDIRIZZO = "<INDIRIZZO>";
const CODICE_SAP = "<CODICE_SAP>";
const CODICE_CIG = "<CODICE_CIG>";
const NR_ORDINE = "<NR_ORDINE>";
const UFF_VENDITE ="<UFF_VENDITE>"; //OLSE
const DATA = "<DATA>";
const EWBS = "<EWBS>"; //A200V01920C030120000_001 
// RIFERIMENTI CELLE SHEET RICHIESTA FATTURA
const CELL_ADDRESS = "E9";
const CELL_DATA_1 = "G14";
const CELL_DATA_2_COL = "E";
const CELL_DATA_2_ROW = 24;
const CELL_ORDER_NR_1 = "B14";
const CELL_ORDER_NR_2 = "D14";
const TOTAL_START_COL = 11;
const QTY_START_COL = 9;
const EXAMPLE_ROW = 21;
const INVOICEROW_START_COL = 2;
const INVOICEROW_COLS_NR = 13;
// STILI TABELLA
const headerStyle: any = {};
headerStyle[DocumentApp.Attribute.BOLD] = true;
const cellStyle: any = {};
cellStyle[DocumentApp.Attribute.BOLD] = false;
cellStyle[DocumentApp.Attribute.UNDERLINE] = false;
const amountCellStyle: any = {};
amountCellStyle[DocumentApp.Attribute.BOLD] = false;
amountCellStyle[DocumentApp.Attribute.ITALIC] = false;
const slantedStyle: any = {};
slantedStyle[DocumentApp.Attribute.STRIKETHROUGH] = true;
slantedStyle[DocumentApp.Attribute.BOLD] = false;
slantedStyle[DocumentApp.Attribute.ITALIC] = true;
const paraStyle: any = {};
paraStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
paraStyle[DocumentApp.Attribute.LINE_SPACING] = 1;
paraStyle[DocumentApp.Attribute.BOLD] = false;
const pFooterStyle: any = {};
pFooterStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
pFooterStyle[DocumentApp.Attribute.LINE_SPACING] = 1;
pFooterStyle[DocumentApp.Attribute.BOLD] = true;
const footerStyle: any = {};
footerStyle[DocumentApp.Attribute.STRIKETHROUGH] = false;
footerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = "#f3f3f3";

/** Inizializzazione e installazione */
function onInstall(e: any) {
    onOpen(e);
}

/** Inizializzazione Frontend */
function onOpen(e: any) {
    SpreadsheetApp.getUi()
        .createMenu("-Fatture OSD-")
        .addItem("Nuova Fattura da selezione", "createInvoiceFromSelection")
        .addItem("Pulisci Foglio", "clearSheet")
        .addToUi();
}

/** Mostra interfaccia HTML */
function showModalFatture() {
    // APRE DIALOG HTML
    const html = HtmlService.createTemplateFromFile('FrontendFattura');
    const template = html.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setHeight(modalHeight);
    //SpreadsheetApp.getUi().showModalDialog(template, 'Crea nuova Fattura');
}

/** Crea fattura dalle righe selezionate  */
function createInvoiceFromSelection() {
    var activeSheet = SpreadsheetApp.getActiveSheet();
    var sheetName = activeSheet.getName();
    var selection = activeSheet.getSelection();
    Logger.log('Current Cell: ' + selection.getCurrentCell().getA1Notation());
    Logger.log('Active Range: ' + selection.getActiveRange().getA1Notation());
    var ranges =  selection.getActiveRangeList().getRanges();
    Logger.log('Ranges: ' + ranges.length);
    for (var i = 0; i < ranges.length; i++) {
        Logger.log('Active Ranges: ' + ranges[i].getA1Notation());
      }      
    Logger.log('Active Sheet: ' + selection.getActiveSheet().getName());
}

/** Crea nuova fattura */
function CreaFattura(datiInput: InvoiceItem) {
    var activeSheet = SpreadsheetApp.getActiveSheet();
    var sheetName = activeSheet.getName();
    Logger.log("Leggo i dati");
    const currentInvoice = LeggiDati(firstRow, lastRow);
    Logger.log("Esco");
    // return; 
    //Logger.log("CreaFAttura: " + JSON.stringify(datiInput));
    const templateId = invoiceTemplateID;
    const templateDoc = DriveApp.getFileById(templateId);
    const date = new Date();
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear();
    const dateString = "" + day + "/" + month + "/" + year;
    Logger.log("Data corrente: " + dateString);
    const orderFullName = sheetName + GetDateString();
    const orderAddress = currentInvoice[0].address;
    // Crea un nuovo documento dal template e sostituisce i dati
    const newDoc = templateDoc.makeCopy(orderFullName);
    const newDocId = newDoc.getId();
    const file = SpreadsheetApp.openById(newDocId);
    const sheet = file.getSheets()[0];

    for (let index = 0; index < currentInvoice.length; index++) {
        const currItem = currentInvoice[index];
        Logger.log("Inserisco riga " + index + ": " + JSON.stringify(currItem));
        sheet.insertRowAfter(21);
        sheet.getRange(iRow, 11, 1, 1)//(start row, start column, number of rows, number of columns
            .setFormulaR1C1("=R[0]C[-2]*R[0]C[-1]");
        sheet.getRange(iRow, 5, 1, 4).merge();
        sheet.getRange(iRow, 2, 1, 2).merge();
        sheet.getRange(iRow, 1, 1, 10)//(start row, start column, number of rows, number of columns
            .setValues([[
                currentInvoice.length - index,
                currItem.itemCode,
                "",
                "",
                currItem.itemDesc,
                "",
                "",
                "",
                currItem.nrItems,
                currItem.itemPrice
            ]]);
    }
    Logger.log("Inizio sostituzione datiInput: " + JSON.stringify(datiInput));
    cellReplaceText(sheet, CELL_ADDRESS, INDIRIZZO, orderAddress);
    cellReplaceText(sheet, CELL_ORDER_NR_1, NR_ORDINE, orderFullName);
    cellReplaceText(sheet, CELL_ORDER_NR_2, NR_ORDINE, orderFullName);
    cellReplaceText(sheet, CELL_DATA_1, DATA, dateString);
    const currData2Row = CELL_DATA_2_ROW + currentInvoice.length + 1;
    cellReplaceText(sheet, CELL_DATA_2_COL + currData2Row, DATA, dateString);

    /** Numero di righe da sommare per totale (numero item + 1) */
    const sumRows = currentInvoice.length + 1;
    /** Formula per totale */
    sheet.getRange(iRow + currentInvoice.length + 1, TOTAL_START_COL, 1, 1)//(start row, start column, number of rows, number of columns
        .setFormulaR1C1("=sum(R[-" + sumRows + "]C[0]:R[-1]C[0])");
    /** Formula per Q.tà */
    sheet.getRange(iRow + currentInvoice.length + 1, QTY_START_COL, 1, 1)//(start row, start column, number of rows, number of columns
        .setFormulaR1C1("=sum(R[-" + sumRows + "]C[0]:R[-1]C[0])");
    /** Elimino riga di esempio */
    sheet.deleteRow(EXAMPLE_ROW);

    Logger.log("Segnaposto sostituiti");
}

/** legge dal foglio i dati della fattura */
function LeggiDati(rigaIniziale: number, rigaFinale: number, updateStock: boolean = true) {
    /** Dati ordine corrente */
    let currentInvoice: InvoiceItem[] = new Array();
    Logger.log("Inizio LeggiDati");
    const file = SpreadsheetApp.openById(invoicesFileID);
    const ss = file.getSheets()[0];
    for (let riga: number = rigaFinale; riga >= rigaIniziale; riga--) {
        // 13 columns starting with column 2, so B-N range 
        const rangeToCheck = ss.getRange(riga, INVOICEROW_START_COL, 1, INVOICEROW_COLS_NR); 
        const readValues = rangeToCheck.getValues();
        const isChecked = readValues[0][0];
        if (isChecked) {
            Logger.log("Riga" + riga);
            const currValue = readValues[0][1].toString();
            let nrItems: number = parseInt(currValue);
            Logger.log("Numero oggetti: " + readValues);
            if (nrItems == 0 || isNaN(nrItems)) {
                nrItems = 1;
            }
            const itemCode: string = readValues[0][3].toString();
            const itemDescription: string = readValues[0][4].toString();
            const itemPrice: number = parseFloat(readValues[0][5].toString());
            const itemCessione: number = parseInt(readValues[0][7].toString());
            const itemOffer: number = parseInt(readValues[0][8].toString());
            const itemStock: number = parseInt(readValues[0][9].toString());
            Logger.log("Aggiungo item " + itemCode);
            AggiungiItem(currentInvoice, nrItems, itemCode, itemDescription, itemPrice, itemCessione, itemOffer, itemStock);
            if (updateStock) {
                AggiornaStock(ss, riga, nrItems);
            }
        }
    }
    Logger.log("***");
    Logger.log(JSON.stringify(currentInvoice));
    Logger.log("***");
    Logger.log("Fine LeggiDati");

    return currentInvoice;
}

/* TOOLS */
/** Aggiunge un item all'ordine */
function AggiungiItem(currentOrder: InvoiceItem[], nrItems: number, itemCode: string,
    itemDesc: string, itemPrice: number, itemCessione: number, itemOffer: number, itemStock: number) {
    const item: InvoiceItem = {
        nrItems: nrItems,
        itemCode: itemCode,
        itemDesc: itemDesc,
        itemPrice: itemPrice,
        itemCessione: itemCessione,
        itemOffer: itemOffer,
        itemStock: itemStock
    };
    currentOrder.push(item);
    Logger.log("Aggiunto item " + JSON.stringify(item));
}

/** Sastituisce nel contenuto di una cella un template con un valore  */
function cellReplaceText(sheet: GoogleAppsScript.Spreadsheet.Sheet, cell: string, template: string, replacement: string) {
    Logger.log("Sostituzione nella cella " + cell + " di " + template + " con " + replacement);
    const dateCell = sheet.getRange(cell);
    const dateCellContent = dateCell.getDisplayValue().replace(template, replacement);
    sheet.getRange(cell).setValue(dateCellContent);
}

/** Include un file nel template */
function include(filename: string) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
}

/** Aggiorna lo stock sottraendo la quantità venduta */
function AggiornaStock(ss: GoogleAppsScript.Spreadsheet.Sheet, row: number, nrItems: number) {
    Logger.log("Aggiorno lo stock");
    /*const currStock = parseInt(ss.getRange(row, stockCol).getValue().toString());
    const newStock = currStock - nrItems;
    Logger.log("Alla riga " + row + " da " + currStock + " a " + newStock);
    ss.getRange(row, stockCol).setValue(newStock);//*/
}

/** Clear script */
function clearSheet() {
}


/* 
 * TIPI E INTERFACCE 
 */

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

/** Nuovo metodo per oggetto Date */
declare interface Date {
    yyyymmdd(): string;
}

/** Interfacci descrizione item ordine */
interface InvoiceItem {
    address: string;
    nrItems: number;
    itemCode: string;
    itemDesc: string;
    itemPrice: number;
    itemCessione: number;
    itemOffer: number;
    itemStock: number;
}