/*
 * CONFIGURAZIONE
 */
const debug = false;
/** ID documenti  */
const OffersCatalogID = "1O-5Ren-h9xJy04rYxQLGSf7dMmGYhgYWsWhk459rSVs";
//const OffersCatalogID = "1hbd9DhHzkCzdp4d3DKTNb1dQZXGb_aLdH1ItkVpkP0I"; // COPIA PER TEST
const templateIdOrder = "1q_Y1vMNIYHBrne-6irggMVtZqlJpBq8lqGf8K-HEvuU";
const ListinoId = "1JYWsDb3bx1eHS3D8JmwHkptUbkJ9R9Lof_RQMegMx3k";
const templateIdHW = "1Q6P7OovyFC3hPMVY_UT5o3-vCTvnIkhwOArSAL0XQ-0";
const templateIdSW = "1OL23gmlnvr4ZSEgZwo6eNfrtWB-4Y-1UXY5bs41rchw";
//const ListinoId = "1jseXrwiSeRG7DPnvx3vGjDdxVagDRZCCKgpRL8Wr8Fg"; // COPIA PER TEST
/** altezza in px della finestra modale */
const modalHeight = 600;
/** Colonna in cui è memorizzato lo stock */
const stockCol = 11;
/** Righe file listino */
const firstRow = 2;
const lastRow = 70;
const iRow = 22;
/** Percentuale spese di spedizione (1,25%)  */
const transportPercent = 0.0125;
// SEGNAPOSTO
const INDIRIZZO = "<INDIRIZZO>";
const NR_ORDINE = "<NR_ORDINE>";
const DATA = "<DATA>";
const NOME_RIFERIMENTO = "<NOME_RIFERIMENTO>";
const EMAIL_RIFERIMENTO = "<EMAIL_RIFERIMENTO>";
const RAGIONE_SOCIALE = "<RAGIONE_SOCIALE>";
const COMUNE = "<COMUNE>";
const PROVINCIA = "<PROVINCIA>";
const CAP = "<CAP>";
const PIVA = "<PIVA>";
const DESCRIZIONE = "<DESCRIZIONE>";
const TABELLA = "<TABELLA>";
// RIFERIMENTI CELLE SHEET ORDINE
const CELL_ADDRESS = "D7";
const CELL_DATA_1 = "G14";
const CELL_DATA_2_COL = "E";
const CELL_DATA_2_ROW = 24;
const CELL_ORDER_NR_1 = "A14";
const CELL_ORDER_NR_2 = "C14";
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
        .createMenu("-Ordini OSD-")
        .addItem("Nuova Offerta", "showModalOfferte")
        .addItem("Nuovo Ordine", "showModalOrdini")
        .addItem("Pulisci Foglio", "clearSheet")
        .addToUi();
}

/** pulisce i dati del foglio */

function clearSheet() {
    const file = SpreadsheetApp.openById(ListinoId);
    const ss = file.getSheets()[0];

    ss.getRange(2, 2, lastRow).setValue(false);
    ss.getRange(2, 3, lastRow).setValue("");

}

/** Mostra interfaccia HTML */
function showModalOfferte() {
    // APRE DIALOG HTML
    const html = HtmlService.createTemplateFromFile('FrontendOfferta');
    const template = html.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setHeight(modalHeight);
    SpreadsheetApp.getUi().showModalDialog(template, 'Crea nuova Offerta');
}

/** Mostra interfaccia HTML */
function showModalOrdini() {
    // APRE DIALOG HTML
    const html = HtmlService.createTemplateFromFile('FrontendOrdine');
    const template = html.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setHeight(modalHeight);
    SpreadsheetApp.getUi().showModalDialog(template, 'Crea nuovo Ordine');
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
    // Aggiorna la variabile currentOrder
    const currentOrder = LeggiDati(firstRow, lastRow, false);
    Logger.log("currentOrder: " + JSON.stringify(currentOrder));
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
    // CREAZIONE TABELLA
    let totalCessione = 0;
    let totalOffer = 0;
    let totalItems = 0;
    const range = body.findText(TABELLA);
    const table = body.findElement(DocumentApp.ElementType.TABLE, range).getElement().asTable();
    // DATI ITEMS
    Logger.log("Creazione tabella items");
    //for (let index = 0; index < currentOrder.length; index++) {
    for (let index = currentOrder.length - 1; index >= 0; index--) {
        const currItem = currentOrder[index];
        totalCessione += currItem.itemCessione * currItem.nrItems;
        totalOffer += currItem.itemOffer * currItem.nrItems;
        totalItems += currItem.nrItems;
        Logger.log("Inserisco riga: " + index + "\n" + JSON.stringify(currItem));
        const currRow = table.appendTableRow();
        // Descrizione
        addCell(currRow, currItem.itemDesc + " (" + currItem.itemCode + ")", cellStyle, paraStyle);
        // Costo listino
        addCell(currRow, ToC(currItem.itemCessione), cellStyle, paraStyle, DocumentApp.HorizontalAlignment.RIGHT);
        // Offerta
        addCell(currRow, ToC(currItem.itemOffer), cellStyle, paraStyle, DocumentApp.HorizontalAlignment.RIGHT);
        // Quantità
        addCell(currRow, currItem.nrItems.toLocaleString(), cellStyle, paraStyle, DocumentApp.HorizontalAlignment.RIGHT);
    }
    // SPESE DI SPEDIZIONE
    const spedizioneCessione = Math.round(totalCessione * transportPercent);
    const spedizioneOfferta = Math.round(totalOffer * transportPercent);
    totalCessione += spedizioneCessione;
    totalOffer += spedizioneOfferta;
    const currRow = table.appendTableRow();
    // Descrizione
    addCell(currRow, "Spese di trasporto", cellStyle, paraStyle);
    // Costo listino
    addCell(currRow, ToC(spedizioneCessione), cellStyle, paraStyle, DocumentApp.HorizontalAlignment.RIGHT);
    // Offerta
    addCell(currRow, ToC(spedizioneOfferta), cellStyle, paraStyle, DocumentApp.HorizontalAlignment.RIGHT);
    // Quantità
    addCell(currRow, "1", cellStyle, paraStyle, DocumentApp.HorizontalAlignment.RIGHT);
    // FOOTER
    const footerRow = table.appendTableRow();
    addCell(footerRow, "TOTALE", slantedStyle, paraStyle);
    addCell(footerRow, " ", slantedStyle, paraStyle);
    addCell(footerRow, " ", slantedStyle, paraStyle);
    addCell(footerRow, ToC(totalCessione), slantedStyle, paraStyle, DocumentApp.HorizontalAlignment.RIGHT);
    const footerRow2 = table.appendTableRow();
    addCell(footerRow2, "PREZZO A VOI DEDICATO", footerStyle, pFooterStyle);
    addCell(footerRow2, " ", footerStyle, pFooterStyle);
    addCell(footerRow2, " ", footerStyle, pFooterStyle);
    addCell(footerRow2, ToC(totalOffer), footerStyle, pFooterStyle, DocumentApp.HorizontalAlignment.RIGHT);
    Logger.log("Creazione tabella items - FINE");
    Logger.log("datiInput: " + JSON.stringify(datiInput));
    // SOSTITUZIONE SEGNAPOSTO
    body.replaceText(TABELLA, "");
    body.replaceText(DATA, new Date().toLocaleDateString("it"));
    body.replaceText(DESCRIZIONE, datiInput.descrizione);
    body.replaceText(NR_ORDINE, orderFullName);
    body.replaceText(NOME_RIFERIMENTO, datiInput.nomeRiferimento);
    body.replaceText(EMAIL_RIFERIMENTO, datiInput.emailRiferimento);
    body.replaceText(RAGIONE_SOCIALE, datiInput.ragioneSociale);
    body.replaceText(COMUNE, datiInput.comune);
    body.replaceText(CAP, datiInput.cap);
    body.replaceText(PROVINCIA, datiInput.prov);
    body.replaceText(INDIRIZZO, datiInput.indirizzo);
    body.replaceText(PIVA, datiInput.pIva);
    Logger.log("Segnaposto sostituiti");
    InsertOrder(orderNumber, orderFullName, datiInput.orderType.toUpperCase(), totalOffer.toLocaleString(), datiInput.ragioneSociale);
    Logger.log("Ordine inserito");
}

function addCell(row: GoogleAppsScript.Document.TableRow, text: string, style: any, paragraphStyle: any, 
    hAlignment: GoogleAppsScript.Document.HorizontalAlignment = DocumentApp.HorizontalAlignment.LEFT) {
    const td1 = row.appendTableCell(text).setAttributes(style);
    const paragraph = td1.getChild(0).asParagraph();
    paragraph.setAttributes(paragraphStyle);
    paragraph.setAlignment(hAlignment);
}

/** Recupera dallo Spreadsheet l'ultimo ordine creato */
function getLastOrderNumber() {
    const sheet = SpreadsheetApp.openById(OffersCatalogID);
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
    const ss = SpreadsheetApp.openById(OffersCatalogID);
    const sheet = ss.getSheets()[0];
    sheet.appendRow([OrderNr, OrderName, OrderType, OrderValue, ClientName]);
}

/** Crea nuovo ordine */
function CreaOrdine(datiInput: DatoOrdine) {
    Logger.log("Leggo i dati");
    const currentOrder = LeggiDati(firstRow, lastRow);
    Logger.log("Esco");
    // return; 
    Logger.log("CreaOrdine: " + JSON.stringify(datiInput));
    const templateId = templateIdOrder;
    const templateDoc = DriveApp.getFileById(templateId);
    const date = new Date();
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear();
    const dateString = "" + day + "/" + month + "/" + year;
    Logger.log("Data corrente: " + dateString);
    const orderFullName = datiInput.orderName;
    const orderAddress = datiInput.orderAddress;
    // Crea un nuovo documento dal template e sostituisce i dati
    const newDoc = templateDoc.makeCopy(orderFullName);
    const newDocId = newDoc.getId();
    const file = SpreadsheetApp.openById(newDocId);
    const sheet = file.getSheets()[0];

    for (let index = 0; index < currentOrder.length; index++) {
        const currItem = currentOrder[index];
        Logger.log("Inserisco riga " + index + ": " + JSON.stringify(currItem));
        sheet.insertRowAfter(21);
        sheet.getRange(iRow, 11, 1, 1)//(start row, start column, number of rows, number of columns
            .setFormulaR1C1("=R[0]C[-2]*R[0]C[-1]");
        sheet.getRange(iRow, 5, 1, 4).merge();
        sheet.getRange(iRow, 2, 1, 2).merge();
        sheet.getRange(iRow, 1, 1, 10)//(start row, start column, number of rows, number of columns
            .setValues([[
                currentOrder.length - index,
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
    const currData2Row = CELL_DATA_2_ROW + currentOrder.length + 1;
    cellReplaceText(sheet, CELL_DATA_2_COL + currData2Row, DATA, dateString);

    /** Numero di righe da sommare per totale (numero item + 1) */
    const sumRows = currentOrder.length + 1;
    /** Formula per totale */
    sheet.getRange(iRow + currentOrder.length + 1, 11, 1, 1)//(start row, start column, number of rows, number of columns
        .setFormulaR1C1("=sum(R[-" + sumRows + "]C[0]:R[-1]C[0])");
    /** Formula per Q.tà */
    sheet.getRange(iRow + currentOrder.length + 1, 9, 1, 1)//(start row, start column, number of rows, number of columns
        .setFormulaR1C1("=sum(R[-" + sumRows + "]C[0]:R[-1]C[0])");
    /** Elimino riga di esempio */
    sheet.deleteRow(21);

    Logger.log("Segnaposto sostituiti");
}

/** legge dal foglio i dati dell'ordine (e aggiorna lo stock se updateStock è true) */
function LeggiDati(rigaIniziale: number, rigaFinale: number, updateStock: boolean = true) {
    /** Dati ordine corrente */
    let currentOrder: ItemOrdine[] = new Array();
    Logger.log("Inizio LeggiDati");
    const file = SpreadsheetApp.openById(ListinoId);
    const ss = file.getSheets()[0];
    for (let riga: number = rigaFinale; riga >= rigaIniziale; riga--) {
        const rangeToCheck = ss.getRange(riga, 2, 1, 10); // 10 columns starting with column 2, so B-J range 
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
            AggiungiItem(currentOrder, nrItems, itemCode, itemDescription, itemPrice, itemCessione, itemOffer, itemStock);
            if (updateStock) {
                AggiornaStock(ss, riga, nrItems);
            }
        }
    }
    Logger.log("***");
    Logger.log(JSON.stringify(currentOrder));
    Logger.log("***");
    Logger.log("Fine LeggiDati");

    return currentOrder;
}

/** Aggiunge un item all'ordine */
function AggiungiItem(currentOrder: ItemOrdine[], nrItems: number, itemCode: string,
    itemDesc: string, itemPrice: number, itemCessione: number, itemOffer: number, itemStock: number) {
    const item: ItemOrdine = {
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

/** Aggiorna lo stock sottraendo la quantità venduta */
function AggiornaStock(ss: GoogleAppsScript.Spreadsheet.Sheet, row: number, nrItems: number) {
    Logger.log("Aggiorno lo stock");
    const currStock = parseInt(ss.getRange(row, stockCol).getValue().toString());
    const newStock = currStock - nrItems;
    Logger.log("Alla riga " + row + " da " + currStock + " a " + newStock);
    ss.getRange(row, stockCol).setValue(newStock);
}

/* 
 * UTILITY
 */
/** sostituisce nel contenuto di una cella un template con un valore  */
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

/** Scrive come valuta */
function ToC(amount: number) {
    return "€ " + amount.toFixed(2).replace(".", ",").replace(/\B(?=(\d{3})+(?!\d))/g, ".");
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

/** Interfaccia dati ordine */
interface DatoOrdine {
    orderName: string;
    orderAddress: string;
}

/** Interfacci descrizione item ordine */
interface ItemOrdine {
    nrItems: number;
    itemCode: string;
    itemDesc: string;
    itemPrice: number;
    itemCessione: number;
    itemOffer: number;
    itemStock: number;
}

/** Interfacci descrizione item offerta */
interface DatoOfferta {
    orderName: string;
    descrizione: string;
    nomeRiferimento: string;
    orderFullName: string;
    emailRiferimento: string;
    ragioneSociale: string;
    comune: string;
    cap: string;
    prov: string;
    indirizzo: string;
    pIva: string;
    orderType: string;
}