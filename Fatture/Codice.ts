/*
 * CONFIGURAZIONE
 */
const debug = false;
/** ID documenti  */
const invoicesFileID = ReadConfigValue("ID file fatture"); //ID del file contenete le fatture
const invoiceTemplateID = ReadConfigValue("ID template fattura"); //ID del template di richiesta fatturazione
const invoiceFolderID = ReadConfigValue("ID cartella fatture"); //ID del template di richiesta fatturazione
//Logger.log("Cartella Fatture ID["+invoiceFolderID+"]");

/** Righe file fatture */
const firstRow = parseInt(ReadConfigValue("Prima riga fatture"));
const lastRow = parseInt(ReadConfigValue("Ultima riga fatture"));
const statusCol = parseInt(ReadConfigValue("Colonna stato fattura"));
const iRow = parseInt(ReadConfigValue("Riga dati template"));
// CARTELLE
const fattureFolderID = ReadConfigValue("ID cartella Fatture");
// SEGNAPOSTO
const RAGIONE_SOCIALE = ReadConfigValue("Segnaposto Ragione Sociale");
const INDIRIZZO = ReadConfigValue("Segnaposto Indirizzo");
const SAPCODE = ReadConfigValue("Segnaposto Codice SAP");
const CIGCODE = ReadConfigValue("Segnaposto Codice CIG");
const NR_ORDINE = ReadConfigValue("Segnaposto Nr Ordine");
const SALES_CODE = ReadConfigValue("Segnaposto Uff. Vendite"); //OLSE - OLSD
const DATA = ReadConfigValue("Segnaposto Data");
const EWBS = ReadConfigValue("Segnaposto EWBS"); //A200V01920C030120000_001 
const CANALE = ReadConfigValue("Segnaposto Canale"); // Es: IT01/Z2
// RIFERIMENTI CELLE SHEET RICHIESTA FATTURA
const CELL_ADDRESS = ReadConfigValue("Template Cella indirizzo");
const CELL_DATA_1 = ReadConfigValue("Template Cella data 1");
const CELL_DATA_2_COL = ReadConfigValue("Template Col data 2"); "E";
const CELL_DATA_2_ROW = ReadConfigValue("Template Riga data 2");
const CELL_CIGCODE = ReadConfigValue("Template Cella CIG");
const CELL_ORDER_NR_1 = ReadConfigValue("Template Cella Ordine");
const CELL_NAME = ReadConfigValue("Template Cella Ragione Sociale");
const CELL_SAPCODE = ReadConfigValue("Template Cella Codice SAP");
const CELL_EWBS = ReadConfigValue("Template Cella EWBS");
const CELL_SALESCODE = ReadConfigValue("Template Cella Ufficio Vendite");
const CELL_CANALE = ReadConfigValue("Template Cella Canale");
const TOTAL_START_COL = parseInt(ReadConfigValue("Colonna Totale"));
const QTY_START_COL = parseInt(ReadConfigValue("Colonna Quantità"));
const EXAMPLE_ROW = parseInt(ReadConfigValue("Riga Esempio"));
const INVOICEROW_START_COL = parseInt(ReadConfigValue("Colonna iniziale fattura"));
const INVOICEROW_COLS_NR = parseInt(ReadConfigValue("Numero Colonne Fattura"));
var filename = "";

/** Inizializzazione e installazione */
function onInstall(e: any) {
    onOpen(e);
}

/** Inizializzazione Frontend */
function onOpen(e: any) {
    SpreadsheetApp.getUi()
        .createMenu("-Fatture OSD-")
        .addItem("Nuova Fattura da selezione", "CreateInvoiceFromSelection")
        .addItem("Pulisci Foglio", "clearSheet")
        //.addItem("DBG Leggi righe selezionate", "LeggiRigheSelezionate")
        .addToUi();
}

/** Crea fattura dalle righe selezionate  */
function CreateInvoiceFromSelection() {
    var activeSheet = SpreadsheetApp.getActiveSheet();
    var sheetName = activeSheet.getName();
    Logger.log("Leggo i dati");
    var currentInvoices = LeggiDati(firstRow, lastRow);
    if(currentInvoices == undefined){
        return;
    }
    Logger.log("Dati letti");
    Logger.log(JSON.stringify(currentInvoices));
    if (currentInvoices.length == 0) {
        showOkPrompt("Non hai selezionato nessuna riga!");
        return;
    }
    //Logger.log("CreaFattura: " + JSON.stringify(datiInput));
    const templateId = invoiceTemplateID;
    const folderId = invoiceFolderID;
    const templateDoc = DriveApp.getFileById(templateId);
    const date = new Date();
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear();
    const dateString = "" + day + "/" + month + "/" + year;
    Logger.log("Data corrente: ${dateString}");
    const orderAddress = currentInvoices[0].address;
    const ragioneSociale = currentInvoices[0].name;
    const sapCode = currentInvoices[0].sapCode;
    const cigCode = currentInvoices[0].cigCode;
    const ewbs = currentInvoices[0].ewbs;
    const salesCode = currentInvoices[0].salesCode;
    var orderNr = currentInvoices[0].orderNr;
    if (currentInvoices.length > 1 && currentInvoices[0].orderNr.toLowerCase() != "contratto scuolabook") {
        orderNr = currentInvoices[0].orderNr + " e altri";
    }
    const canaleString = currentInvoices[0].channel;
    // Crea un nuovo documento dal template e sostituisce i dati
    const newDoc = templateDoc.makeCopy(filename);
    const newDocId = newDoc.getId();
    const file = SpreadsheetApp.openById(newDocId);
    const sheet = file.getSheets()[0];

    for (let index = 0; index < currentInvoices.length; index++) {
        const currItem = currentInvoices[index];
        Logger.log("Inserisco riga " + index + ": " + JSON.stringify(currItem));
        sheet.insertRowAfter(21);
        //Calcolo Importo Colonna 11
        sheet.getRange(iRow, 11, 1, 1)//(start row, start column, number of rows, number of columns
            .setFormulaR1C1("=R[0]C[-2]*R[0]C[-1]");
        sheet.getRange(iRow, 12, 1, 1)//(start row, start column, number of rows, number of columns
            .setValue(currItem.ewbs)
        sheet.getRange(iRow, 5, 1, 3).merge();
        sheet.getRange(iRow, 3, 1, 2).merge();
        var description = currItem.description;
        if (currItem.orderNr.toLowerCase() != "contratto scuolabook") {
            description = description + " (" + currItem.orderNr + ")";
        }
        sheet.getRange(iRow, 2, 1, 9)//(start row, start column, number of rows, number of columns
            .setValues([[
                currentInvoices.length - index, //B
                currItem.productCode,           //C
                "",                             //D
                description,                    //E
                "",                             //F
                "",                             //G
                "",                             //H
                currItem.nrItems,               //I
                currItem.price,                 //J
            ]]);
    }
    //Logger.log("Inizio sostituzione datiInput: " + JSON.stringify(datiInput));
    cellReplaceText(sheet, CELL_SAPCODE, SAPCODE, sapCode);
    cellReplaceText(sheet, CELL_ORDER_NR_1, NR_ORDINE, orderNr);
    cellReplaceText(sheet, CELL_CIGCODE, CIGCODE, cigCode);
    cellReplaceText(sheet, CELL_EWBS, EWBS, ewbs);
    cellReplaceText(sheet, CELL_NAME, RAGIONE_SOCIALE, ragioneSociale);
    cellReplaceText(sheet, CELL_ADDRESS, INDIRIZZO, orderAddress);
    cellReplaceText(sheet, CELL_ORDER_NR_1, NR_ORDINE, filename);
    cellReplaceText(sheet, CELL_SALESCODE, SALES_CODE, salesCode);
    cellReplaceText(sheet, CELL_DATA_1, DATA, dateString);
    cellReplaceText(sheet, CELL_CANALE, CANALE, canaleString);
    const currData2Row = CELL_DATA_2_ROW + currentInvoices.length + 1;
    cellReplaceText(sheet, CELL_DATA_2_COL + currData2Row, DATA, dateString);
    /** Numero di righe da sommare per totale (numero item + 1) */
    const sumRows = currentInvoices.length + 1;
    /** Formula per totale */
    sheet.getRange(iRow + currentInvoices.length, TOTAL_START_COL, 1, 1)//(start row, start column, number of rows, number of columns
        .setFormulaR1C1("=sum(R[-" + sumRows + "]C[0]:R[-1]C[0])");
    /** Formula per Q.tà */
    sheet.getRange(iRow + currentInvoices.length, QTY_START_COL, 1, 1)//(start row, start column, number of rows, number of columns
        .setFormulaR1C1("=sum(R[-" + sumRows + "]C[0]:R[-1]C[0])");
    /** Elimino riga di esempio */
    sheet.deleteRow(EXAMPLE_ROW);
    Logger.log("Segnaposto sostituiti");
    // Sposto la fattura nella cartella "Fatture" su Drive 
    Logger.log("Sposto il file nella cartella 'Fatture'");
    moveFiles(newDocId, folderId);
    Logger.log("Pulizia foglio");
    if(showYesNoPrompt("Creato  documento " + filename + ". Pulire la selezione?", "Operazione terminata")){
        clearSheet();
    }
    Logger.log("FINE");
}

/** legge vettore righe attive  */
function LeggiRigheSelezionate(rigaIniziale: number = firstRow, rigaFinale: number = lastRow) {
    //Logger.log("controllo le righe dalla " + firstRow + " alla " + lastRow);
    let activeRows: number[] = [];
    const file = SpreadsheetApp.openById(invoicesFileID);
    const ss = file.getSheets()[0];
    /*Parametri: riga, colonna, nrighe, nrcolonne */
    const rangeToCheck = ss.getRange(rigaIniziale, INVOICEROW_START_COL, rigaFinale - rigaIniziale, 1).getValues();
    //Logger.log("  Array è " + JSON.stringify(rangeToCheck));
    for (let i: number = 0; i < rangeToCheck.length; i++) {
        Logger.log("  Dato [0][" + i + "] Valore " + rangeToCheck[i][0]);
        if (rangeToCheck[i][0] == true) {
            activeRows.push(i + rigaIniziale);
            //Logger.log("E' attiva la riga " + i);
        }
    }
    return activeRows;
}

/** legge dal foglio i dati della fattura */
function LeggiDati(rigaIniziale: number, rigaFinale: number, updateStatus: boolean = true) {
    /** Dati ordine corrente */
    Logger.log("Inizio LeggiDati");
    let currentInvoice: InvoiceItem[] = new Array();
    const file = SpreadsheetApp.openById(invoicesFileID);
    const ss = file.getSheets()[0];
    const activeRows: Array<number> = LeggiRigheSelezionate();
    for (let i: number = activeRows.length - 1; i >= 0; i--) {
        const riga = activeRows[i];
        // 17 columns starting with column 2, so B-R range 
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
            // Ragione Sociale
            const itemYear: number = readValues[0][1].toString();
            //Logger.log("  itemYear: " + itemYear);
            const itemMonth: number = readValues[0][2].toString();
            //Logger.log("  itemMonth: " + itemMonth);
            const itemShortname: string = readValues[0][3].toString();
            //Logger.log("  itemShortname: " + itemShortname);
            const itemProductType: string = readValues[0][4].toString();
            //Logger.log("  productType: " + itemProductType);
            const itemOrderNr: string = readValues[0][5].toString();
            //Logger.log("  itemOrderNr: " + itemOrderNr);
            const itemName: string = readValues[0][6].toString();
            //Logger.log("  itemName: " + itemName);
            const itemAddress: string = readValues[0][7].toString();
            //Logger.log("  itemAddress: " + itemAddress);
            const itemSapCode: string = readValues[0][9].toString();
            if(itemSapCode.trim() == "") {
                Logger.log("Codice SAP non definito. Chiedo all'utente");
                if(!showYesNoPrompt("Codice SAP non definito. Si vuole continuare?", "Attenzione!")) {
                    Logger.log("Uscita per 'Codice SAP non definito'");
                    return undefined;
                }
            }
            //Logger.log("  itemSapCode: " + itemSapCode);
            const itemCigCode: string = readValues[0][10].toString();
            //Logger.log("  itemCigCode: " + itemCigCode);
            const itemEwbsCode: string = readValues[0][11].toString();
            //Logger.log("  itemEwbsCode: " + itemEwbsCode);
            const itemSalesCode: string = readValues[0][12].toString();
            //Logger.log("  itemSalesCode: " + itemSalesCode);
            const itemDescription: string = readValues[0][13].toString();
            //Logger.log("  itemDescription: " + itemDescription);
            const itemProductCode: string = readValues[0][14].toString();
            //Logger.log("  itemProductCode: " + itemProductCode);
            const itemQty: number = readValues[0][15].toString();
            //Logger.log("  itemQty: " + itemQty);
            const itemPrice: number = parseFloat(readValues[0][17].toString());
            //Logger.log("  itemPrice: " + readValues[0][17].toString());
            const itemChannel: string = readValues[0][20].toString();
            //Logger.log("  itemChannel: " + itemChannel);
            Logger.log("Aggiungo item " + itemName);
            AddiItem(currentInvoice,
                itemYear,
                itemMonth,
                itemName,
                itemShortname,
                itemAddress,
                itemSapCode,
                itemCigCode,
                itemEwbsCode,
                itemSalesCode,
                itemDescription,
                itemProductCode,
                itemQty,
                itemPrice,
                itemProductType,
                itemOrderNr,
                itemChannel
            );
            // Define filename only once in each run
            if (filename == "") {
                filename = "Ft_" + itemProductType + "_" + itemShortname + "_" + itemYear + pad(itemMonth, 2);
            }
            if (updateStatus) {
                UpdateStatus(ss, riga, filename);
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
function AddiItem(currentInvoice: InvoiceItem[], itemYear: number, itemMonth: number, itemName: string, itemShortname: string, 
    itemAddress: string,itemSapCode: string, itemCigCode: string, itemEwbsCode: string, itemSalesCode: string, itemDescription: string,
    itemProductCode: string, itemQty: number, itemPrice: number, itemProductType: string, itemOrderNr: string, itemChannel: string
    ) {
    const item: InvoiceItem = {
        name: itemName,
        shortname: itemShortname,
        year: itemYear,
        month: itemMonth,
        address: itemAddress,
        sapCode: itemSapCode,
        cigCode: itemCigCode,
        ewbs: itemEwbsCode,
        salesCode: itemSalesCode,
        description: itemDescription,
        productCode: itemProductCode,
        nrItems: itemQty,
        price: itemPrice,
        productType: itemProductType,
        orderNr: itemOrderNr,
        channel: itemChannel,
    };
    currentInvoice.push(item);
    Logger.log("Aggiunto item " + JSON.stringify(item));
} //*/

/** Legge un valore dalla configurazione  */
function ReadConfigValue(paramName: string) {
    var value = "";
    const A = 0;
    const B = 1;
    const sheet = SpreadsheetApp.getActive().getSheetByName('Config');
    var data = sheet.getDataRange().getValues();
    for (var i = 0; i < data.length; i++) {
        if (data[i][A] == paramName) {      
            value = data[i][B];
            //Logger.log("[" + data[i][A] + "]");     
            break;
        }
    }
    //Logger.log(paramName + " = [" + value + "]");
    return value;
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

/** Aggiorna lo stato aggiungendo il nome del file creato */
function UpdateStatus(ss: GoogleAppsScript.Spreadsheet.Sheet, row: number, newStatus: string) {
    const currStatus = parseInt(ss.getRange(row, statusCol).getValue().toString());
    Logger.log("Aggiorno lo stato alla riga " + row + " da " + currStatus + " a " + newStatus);
    ss.getRange(row, statusCol).setValue(newStatus);
    ss.getRange(row, statusCol).setBackground("#FFFF00");
}

/** Clear script */
function clearSheet() {
    const file = SpreadsheetApp.openById(invoicesFileID);
    const ss = file.getSheets()[0];
    ss.getRange(2, 2, lastRow).setValue(false);
}

/** Show a info prompt */
function showOkPrompt(text: string, title: string = "Attenzione") {
    var ui = SpreadsheetApp.getUi();
    var result = ui.alert(
        title,
        text,
        ui.ButtonSet.OK);
    return;
}

/** Show a yes/no choice prompt */
function showYesNoPrompt(text: string, title: string = "Attenzione") {
    var ui = SpreadsheetApp.getUi(); 
    var result = ui.alert(
        title,
        text,
        ui.ButtonSet.YES_NO);
    if (result == ui.Button.YES) {
        return true;
    }
    return false;
}

/** Apply zero padding */
function pad(num, size) {
    var s = num + "";
    while (s.length < size) s = "0" + s;
    return s;
}

/** Move file to folder */
function moveFiles(sourceFileId, targetFolderId) {
    Logger.log("Sposto file ID[" + sourceFileId + "] nella cartella ID[" + targetFolderId + "]");
    var file = DriveApp.getFileById(sourceFileId);
    file.getParents().next().removeFile(file);
    DriveApp.getFolderById(targetFolderId).addFile(file);
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

/** Interfaccia descrizione item ordine */
interface InvoiceItem {
    year: number,
    month: number,
    sapCode: string;
    cigCode: string;
    ewbs: string;
    name: string;
    shortname: string;
    salesCode: string;
    address: string;
    nrItems: number;
    description: string;
    productCode: string;
    price: number;
    productType: string;
    orderNr: string;
    channel: string;
}