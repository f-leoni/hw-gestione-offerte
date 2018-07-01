/*
 * CONFIGURAZIONE
 */
var debug = false;
/** ID documenti  */
var OffersCatalogID = "1O-5Ren-h9xJy04rYxQLGSf7dMmGYhgYWsWhk459rSVs";
//const OffersCatalogID = "1hbd9DhHzkCzdp4d3DKTNb1dQZXGb_aLdH1ItkVpkP0I"; // COPIA PER TEST
var templateIdOrder = "1q_Y1vMNIYHBrne-6irggMVtZqlJpBq8lqGf8K-HEvuU";
var ListinoId = "1JYWsDb3bx1eHS3D8JmwHkptUbkJ9R9Lof_RQMegMx3k";
var templateIdHW = "1Q6P7OovyFC3hPMVY_UT5o3-vCTvnIkhwOArSAL0XQ-0";
var templateIdSW = "1OL23gmlnvr4ZSEgZwo6eNfrtWB-4Y-1UXY5bs41rchw";
//const ListinoId = "1jseXrwiSeRG7DPnvx3vGjDdxVagDRZCCKgpRL8Wr8Fg"; // COPIA PER TEST
/** altezza in px della finestra modale */
var modalHeight = 600;
/** Colonna in cui è memorizzato lo stock */
var stockCol = 11;
/** Righe file listino */
var firstRow = 2;
var lastRow = 50;
var iRow = 22;
/** Percentuale spese di spedizione (1,25%)  */
var transportPercent = 0.0125;
// SEGNAPOSTO
var INDIRIZZO = "<INDIRIZZO>";
var NR_ORDINE = "<NR_ORDINE>";
var DATA = "<DATA>";
var NOME_RIFERIMENTO = "<NOME_RIFERIMENTO>";
var EMAIL_RIFERIMENTO = "<EMAIL_RIFERIMENTO>";
var RAGIONE_SOCIALE = "<RAGIONE_SOCIALE>";
var COMUNE = "<COMUNE>";
var PROVINCIA = "<PROVINCIA>";
var CAP = "<CAP>";
var PIVA = "<PIVA>";
var DESCRIZIONE = "<DESCRIZIONE>";
var TABELLA = "<TABELLA>";
// RIFERIMENTI CELLE SHEET ORDINE
var CELL_ADDRESS = "D7";
var CELL_DATA_1 = "G14";
var CELL_DATA_2_COL = "E";
var CELL_DATA_2_ROW = 24;
var CELL_ORDER_NR_1 = "A14";
var CELL_ORDER_NR_2 = "C14";
// STILI TABELLA
var headerStyle = {};
headerStyle[DocumentApp.Attribute.BOLD] = true;
var cellStyle = {};
cellStyle[DocumentApp.Attribute.BOLD] = false;
cellStyle[DocumentApp.Attribute.UNDERLINE] = false;
var amountCellStyle = {};
amountCellStyle[DocumentApp.Attribute.BOLD] = false;
amountCellStyle[DocumentApp.Attribute.ITALIC] = false;
var slantedStyle = {};
slantedStyle[DocumentApp.Attribute.STRIKETHROUGH] = true;
slantedStyle[DocumentApp.Attribute.BOLD] = false;
slantedStyle[DocumentApp.Attribute.ITALIC] = true;
var paraStyle = {};
paraStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
paraStyle[DocumentApp.Attribute.LINE_SPACING] = 1;
paraStyle[DocumentApp.Attribute.BOLD] = false;
var pFooterStyle = {};
pFooterStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
pFooterStyle[DocumentApp.Attribute.LINE_SPACING] = 1;
pFooterStyle[DocumentApp.Attribute.BOLD] = true;
var footerStyle = {};
footerStyle[DocumentApp.Attribute.STRIKETHROUGH] = false;
footerStyle[DocumentApp.Attribute.BACKGROUND_COLOR] = "#f3f3f3";
/** Inizializzazione */
function onInstall(e) {
    onOpen(e);
}
/** Inizializzazione Frontend */
function onOpen(e) {
    SpreadsheetApp.getUi()
        .createMenu("-Ordini OSD-")
        .addItem("Nuova Offerta", "showModalOfferte")
        .addItem("Nuovo Ordine", "showModalOrdini")
        .addToUi();
}
/** Mostra interfaccia HTML */
function showModalOfferte() {
    // APRE DIALOG HTML
    var html = HtmlService.createTemplateFromFile('FrontendOfferta');
    var template = html.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setHeight(modalHeight);
    SpreadsheetApp.getUi().showModalDialog(template, 'Crea nuova Offerta');
}
/** Mostra interfaccia HTML */
function showModalOrdini() {
    // APRE DIALOG HTML
    var html = HtmlService.createTemplateFromFile('FrontendOrdine');
    var template = html.evaluate()
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setHeight(modalHeight);
    SpreadsheetApp.getUi().showModalDialog(template, 'Crea nuovo Ordine');
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
    // Aggiorna la variabile currentOrder
    var currentOrder = LeggiDati(firstRow, lastRow, false);
    Logger.log("currentOrder: " + JSON.stringify(currentOrder));
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
    // CREAZIONE TABELLA
    var totalCessione = 0;
    var totalOffer = 0;
    var totalItems = 0;
    var range = body.findText(TABELLA);
    var table = body.findElement(DocumentApp.ElementType.TABLE, range).getElement().asTable();
    // DATI ITEMS
    Logger.log("Creazione tabella items");
    //for (let index = 0; index < currentOrder.length; index++) {
    for (var index = currentOrder.length - 1; index >= 0; index--) {
        var currItem = currentOrder[index];
        totalCessione += currItem.itemCessione * currItem.nrItems;
        totalOffer += currItem.itemOffer * currItem.nrItems;
        totalItems += currItem.nrItems;
        Logger.log("Inserisco riga: " + index + "\n" + JSON.stringify(currItem));
        var currRow_1 = table.appendTableRow();
        // Descrizione
        addCell(currRow_1, currItem.itemDesc + " (" + currItem.itemCode + ")", cellStyle, paraStyle);
        // Costo listino
        addCell(currRow_1, ToC(currItem.itemCessione), cellStyle, paraStyle, DocumentApp.HorizontalAlignment.RIGHT);
        // Offerta
        addCell(currRow_1, ToC(currItem.itemOffer), cellStyle, paraStyle, DocumentApp.HorizontalAlignment.RIGHT);
        // Quantità
        addCell(currRow_1, currItem.nrItems.toLocaleString(), cellStyle, paraStyle, DocumentApp.HorizontalAlignment.RIGHT);
    }
    // SPESE DI SPEDIZIONE
    var spedizioneCessione = Math.round(totalCessione * transportPercent);
    var spedizioneOfferta = Math.round(totalOffer * transportPercent);
    totalCessione += spedizioneCessione;
    totalOffer += spedizioneOfferta;
    var currRow = table.appendTableRow();
    // Descrizione
    addCell(currRow, "Spese di trasporto", cellStyle, paraStyle);
    // Costo listino
    addCell(currRow, ToC(spedizioneCessione), cellStyle, paraStyle, DocumentApp.HorizontalAlignment.RIGHT);
    // Offerta
    addCell(currRow, ToC(spedizioneOfferta), cellStyle, paraStyle, DocumentApp.HorizontalAlignment.RIGHT);
    // Quantità
    addCell(currRow, "1", cellStyle, paraStyle, DocumentApp.HorizontalAlignment.RIGHT);
    // FOOTER
    var footerRow = table.appendTableRow();
    addCell(footerRow, "TOTALE", slantedStyle, paraStyle);
    addCell(footerRow, " ", slantedStyle, paraStyle);
    addCell(footerRow, " ", slantedStyle, paraStyle);
    addCell(footerRow, ToC(totalCessione), slantedStyle, paraStyle, DocumentApp.HorizontalAlignment.RIGHT);
    var footerRow2 = table.appendTableRow();
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
function addCell(row, text, style, paragraphStyle, hAlignment) {
    if (hAlignment === void 0) { hAlignment = DocumentApp.HorizontalAlignment.LEFT; }
    var td1 = row.appendTableCell(text).setAttributes(style);
    var paragraph = td1.getChild(0).asParagraph();
    paragraph.setAttributes(paragraphStyle);
    paragraph.setAlignment(hAlignment);
}
/** Recupera dallo Spreadsheet l'ultimo ordine creato */
function getLastOrderNumber() {
    var sheet = SpreadsheetApp.openById(OffersCatalogID);
    var data = sheet.getDataRange().getValues();
    if (debug)
        Logger.log("L'ultimo ordine è: " + data[data.length - 1]);
    return data[data.length - 1][0];
}
/** Inserisce l'ordine nello Spreadsheet degli ordini */
function InsertOrder(OrderNr, OrderName, OrderType, OrderValue, ClientName) {
    var ss = SpreadsheetApp.openById(OffersCatalogID);
    var sheet = ss.getSheets()[0];
    sheet.appendRow([OrderNr, OrderName, OrderType, OrderValue, ClientName]);
}
/** Crea nuovo ordine */
function CreaOrdine(datiInput) {
    Logger.log("Leggo i dati");
    var currentOrder = LeggiDati(firstRow, lastRow);
    Logger.log("Esco");
    // return; 
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
    for (var index = 0; index < currentOrder.length; index++) {
        var currItem = currentOrder[index];
        Logger.log("Inserisco riga " + index + ": " + JSON.stringify(currItem));
        sheet.insertRowAfter(21);
        sheet.getRange(iRow, 11, 1, 1) //(start row, start column, number of rows, number of columns
            .setFormulaR1C1("=R[0]C[-2]*R[0]C[-1]");
        sheet.getRange(iRow, 5, 1, 4).merge();
        sheet.getRange(iRow, 2, 1, 2).merge();
        sheet.getRange(iRow, 1, 1, 10) //(start row, start column, number of rows, number of columns
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
    var currData2Row = CELL_DATA_2_ROW + currentOrder.length + 1;
    cellReplaceText(sheet, CELL_DATA_2_COL + currData2Row, DATA, dateString);
    /** Numero di righe da sommare per totale (numero item + 1) */
    var sumRows = currentOrder.length + 1;
    /** Formula per totale */
    sheet.getRange(iRow + currentOrder.length + 1, 11, 1, 1) //(start row, start column, number of rows, number of columns
        .setFormulaR1C1("=sum(R[-" + sumRows + "]C[0]:R[-1]C[0])");
    /** Formula per Q.tà */
    sheet.getRange(iRow + currentOrder.length + 1, 9, 1, 1) //(start row, start column, number of rows, number of columns
        .setFormulaR1C1("=sum(R[-" + sumRows + "]C[0]:R[-1]C[0])");
    /** Elimino riga di esempio */
    sheet.deleteRow(21);
    Logger.log("Segnaposto sostituiti");
}
/** legge dal foglio i dati dell'ordine (e aggiorna lo stock se updateStock è true) */
function LeggiDati(rigaIniziale, rigaFinale, updateStock) {
    if (updateStock === void 0) { updateStock = true; }
    /** Dati ordine corrente */
    var currentOrder = new Array();
    Logger.log("Inizio LeggiDati");
    var file = SpreadsheetApp.openById(ListinoId);
    var ss = file.getSheets()[0];
    for (var riga = rigaFinale; riga >= rigaIniziale; riga--) {
        var rangeToCheck = ss.getRange(riga, 2, 1, 10); // 10 columns starting with column 2, so B-J range 
        var readValues = rangeToCheck.getValues();
        var isChecked = readValues[0][0];
        if (isChecked) {
            Logger.log("Riga" + riga);
            var currValue = readValues[0][1].toString();
            var nrItems = parseInt(currValue);
            Logger.log("Numero oggetti: " + readValues);
            if (nrItems == 0 || isNaN(nrItems)) {
                nrItems = 1;
            }
            var itemCode = readValues[0][3].toString();
            var itemDescription = readValues[0][4].toString();
            var itemPrice = parseFloat(readValues[0][5].toString());
            var itemCessione = parseInt(readValues[0][7].toString());
            var itemOffer = parseInt(readValues[0][8].toString());
            var itemStock = parseInt(readValues[0][9].toString());
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
function AggiungiItem(currentOrder, nrItems, itemCode, itemDesc, itemPrice, itemCessione, itemOffer, itemStock) {
    var item = {
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
function AggiornaStock(ss, row, nrItems) {
    Logger.log("Aggiorno lo stock");
    var currStock = parseInt(ss.getRange(row, stockCol).getValue().toString());
    var newStock = currStock - nrItems;
    Logger.log("Alla riga " + row + " da " + currStock + " a " + newStock);
    ss.getRange(row, stockCol).setValue(newStock);
}
/*
 * UTILITY
 */
/** sostituisce nel contenuto di una cella un template con un valore  */
function cellReplaceText(sheet, cell, template, replacement) {
    Logger.log("Sostituzione nella cella " + cell + " di " + template + " con " + replacement);
    var dateCell = sheet.getRange(cell);
    var dateCellContent = dateCell.getDisplayValue().replace(template, replacement);
    sheet.getRange(cell).setValue(dateCellContent);
}
/** Include un file nel template */
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
}
/** Scrive come valuta */
function ToC(amount) {
    //TODO Manca separatore delle migliaia
    return "€ " + amount.toFixed(2).replace(".", ",");
}
/*
 * TIPI E INTERFACCE
 */
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
