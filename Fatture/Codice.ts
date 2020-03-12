/*
 * Configuration
 */
const debug = false;
let config: ConfigData;
let filename = "";

/** Initialization and installation */
function onInstall(e: any) {
    onOpen(e);
}

/** Frontend Initialization */
function onOpen(e: any) {
    //config = readConfig();
    SpreadsheetApp.getUi()
        .createMenu("-Fatture OSD-")
        .addItem("Nuova Fattura da selezione", "createInvoiceFromSelection")
        .addItem("Pulisci Foglio", "clearSheet")
        .addToUi();
    }

/** Create invoice from selected rows  */
function createInvoiceFromSelection() {
    var activeSheet = SpreadsheetApp.getActiveSheet();
    var sheetName = activeSheet.getName();
    Logger.log("Leggo i dati");
    SpreadsheetApp.getActiveSpreadsheet().toast(new Date().toLocaleString() + '\nLettura dati');
    if(!config) { 
        config = readConfig();
    }
    var currentInvoices = readData(config.firstRow, config.lastRow);
    SpreadsheetApp.getActiveSpreadsheet().toast(new Date().toLocaleString() + '\nLettura dati completata');
    if (currentInvoices == undefined) {
        return;
    }
    Logger.log("Dati letti");
    Logger.log(JSON.stringify(currentInvoices));
    if (currentInvoices.length == 0) {
        showOkPrompt("Non hai selezionato nessuna riga!");
        return;
    }
    //Logger.log("CreaFattura: " + JSON.stringify(datiInput));
    const templateId = config.invoiceTemplateID;
    const folderId = config.invoiceFolderID;
    const templateDoc = DriveApp.getFileById(templateId);
    const date = new Date();
    const month = date.getMonth() + 1;
    const day = date.getDate();
    const year = date.getFullYear();
    const dateString = "" + day + "/" + month + "/" + year;
    Logger.log("Data corrente: ${dateString}");
    const orderAddress = currentInvoices[0].address;
    const fiscalName = currentInvoices[0].fiscalName;
    const sapCode = currentInvoices[0].sapCode;
    const cigCode = currentInvoices[0].cigCode;
    const ewbs = currentInvoices[0].ewbs;
    const salesCode = currentInvoices[0].salesCode;
    var orderNr = currentInvoices[0].orderNr;
    if (currentInvoices.length > 1 && currentInvoices[0].orderNr.toLowerCase() != "contratto scuolabook") {
        orderNr = currentInvoices[0].orderNr + " e altri";
    }
    const channel = currentInvoices[0].channel;
    SpreadsheetApp.getActiveSpreadsheet().toast(new Date().toLocaleString() + '\nCreazione modulo');
    // Creates a new Sheet from template and replace data
    const newDoc = templateDoc.makeCopy(filename);
    const newDocId = newDoc.getId();
    const file = SpreadsheetApp.openById(newDocId);
    const sheet = file.getSheets()[0];

    for (let index = 0; index < currentInvoices.length; index++) {
        const currItem = currentInvoices[index];
        Logger.log("Inserisco riga " + index + ": " + JSON.stringify(currItem));
        sheet.insertRowAfter(21);
        // Calc value Col 11
        sheet.getRange(config.iRow, 11, 1, 1)//(start row, start column, number of rows, number of columns
            .setFormulaR1C1("=R[0]C[-2]*R[0]C[-1]");
        sheet.getRange(config.iRow, 12, 1, 1)//(start row, start column, number of rows, number of columns
            .setValue(currItem.ewbs)
        sheet.getRange(config.iRow, 5, 1, 3).merge();
        sheet.getRange(config.iRow, 3, 1, 2).merge();
        var description = currItem.description;
        if (currItem.orderNr.toLowerCase() != "contratto scuolabook") {
            description = description + " (" + currItem.orderNr + ")";
        }
        sheet.getRange(config.iRow, 2, 1, 9)//(start row, start column, number of rows, number of columns
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
    cellReplaceText(sheet, config.cellSapCode, config.sapcode, sapCode);
    cellReplaceText(sheet, config.cellOrderNr1, config.orderNr, orderNr);
    cellReplaceText(sheet, config.cellCigCode, config.cigCode, cigCode);
    cellReplaceText(sheet, config.cellEwbs, config.ewbs, ewbs);
    cellReplaceText(sheet, config.cellName, config.fiscalName, fiscalName);
    cellReplaceText(sheet, config.cellAddress, config.address, orderAddress);
    cellReplaceText(sheet, config.cellOrderNr1, config.orderNr, filename);
    cellReplaceText(sheet, config.cellSalesCode, config.salesCode, salesCode);
    cellReplaceText(sheet, config.cellData1, config.date, dateString);
    cellReplaceText(sheet, config.cellChannel, config.channel, channel);
    const currData2Row = config.cellData2Row + currentInvoices.length + 1;
    cellReplaceText(sheet, config.cellData2Col + currData2Row, config.date, dateString);
    /** Number of rows to be summed to obtrain total (items numbr + 1) */
    const sumRows = currentInvoices.length + 1;
    /** Formula for total */
    sheet.getRange(config.iRow + currentInvoices.length, config.totalStartCol, 1, 1)//(start row, start column, number of rows, number of columns
        .setFormulaR1C1("=sum(R[-" + sumRows + "]C[0]:R[-1]C[0])");
    /** Formula for Q.ty */
    sheet.getRange(config.iRow + currentInvoices.length, config.qtyStartCol, 1, 1)//(start row, start column, number of rows, number of columns
        .setFormulaR1C1("=sum(R[-" + sumRows + "]C[0]:R[-1]C[0])");
    /** Remove example row */
    sheet.deleteRow(config.exampleRow);
    Logger.log("Segnaposto sostituiti");
    SpreadsheetApp.getActiveSpreadsheet().toast(new Date().toLocaleString() + '\nCreazione modulo completata');
    // Move invoice into folder "Fatture" on GDrive 
    Logger.log("Sposto il file nella cartella 'Fatture'");
    moveFiles(newDocId, folderId);
    SpreadsheetApp.getActiveSpreadsheet().toast(new Date().toLocaleString() + '\nSpostamento file');
    Logger.log("Pulizia foglio");
    if (showYesNoPrompt("Creato  documento " + filename + ". Pulire la selezione?", "Operazione terminata")) {
        clearSheet();
    }
    Logger.log("FINE");
    SpreadsheetApp.getActiveSpreadsheet().toast(new Date().toLocaleString() + '\nCreazione modulo completata - FINE');

}

/** legge vettore righe attive  */
function readSelectedRows(firstRow: number = config.firstRow, lastRow: number = config.lastRow) {
    //Logger.log("controllo le righe dalla " + firstRow + " alla " + lastRow);
    let activeRows: number[] = [];
    const file = SpreadsheetApp.openById(config.invoicesFileID);
    const ss = file.getSheets()[0];
    // Parameters: row, column, nr rows, nr cols 
    const rangeToCheck = ss.getRange(firstRow, config.invoiceRowStartCol, lastRow - firstRow, 1).getValues();
    // Logger.log("  Array è " + JSON.stringify(rangeToCheck));
    for (let i: number = 0; i < rangeToCheck.length; i++) {
        Logger.log("  Dato [0][" + i + "] Valore " + rangeToCheck[i][0]);
        if (rangeToCheck[i][0] == true) {
            activeRows.push(i + firstRow);
            //Logger.log("E' attiva la riga " + i);
        }
    }
    return activeRows;
}

/** read invoice data from sheet */
function readData(firstRow: number, lastRow: number, updateStatus: boolean = true) {
    /** Current invoice data */
    Logger.log("Inizio LeggiDati");
    let currentInvoice: InvoiceItem[] = new Array();
    const file = SpreadsheetApp.openById(config.invoicesFileID);
    const ss = file.getSheets()[0];
    const activeRows: Array<number> = readSelectedRows();
    for (let i: number = activeRows.length - 1; i >= 0; i--) {
        const currRow = activeRows[i];
        // 21 columns starting with column 2, so B-V range 
        const rangeToCheck = ss.getRange(currRow, config.invoiceRowStartCol, 1, config.invoiceRowColsNr);
        const readValues = rangeToCheck.getValues();
        const isChecked = readValues[0][0];
        if (isChecked) {
            Logger.log("Riga" + currRow);
            const currValue = readValues[0][1].toString();
            let nrItems: number = parseInt(currValue);
            Logger.log("Numero oggetti: " + readValues);
            if (nrItems == 0 || isNaN(nrItems)) {
                nrItems = 1;
            }
            // Fiscal Name
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
            if (itemSapCode.trim() == "") {
                Logger.log("Codice SAP non definito. Chiedo all'utente");
                if (!showYesNoPrompt("Codice SAP non definito. Si vuole continuare?", "Attenzione!")) {
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
            addiItem(currentInvoice,
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
                if(fileExists(filename)){
                    filename = filename + "_" + new Date().getTime();//getDateString();
                }
            }
            if (updateStatus) {
                updateState(ss, currRow, filename);
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
/** Add an  item to invoice */
function addiItem(currentInvoice: InvoiceItem[], itemYear: number, itemMonth: number, itemName: string, itemShortname: string,
    itemAddress: string, itemSapCode: string, itemCigCode: string, itemEwbsCode: string, itemSalesCode: string, itemDescription: string,
    itemProductCode: string, itemQty: number, itemPrice: number, itemProductType: string, itemOrderNr: string, itemChannel: string
) {
    const item: InvoiceItem = {
        fiscalName: itemName,
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

/** Read some value fronm vonfig sheet  */
function readConfigValue(paramName: string) {
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

/** Replace template with value in given cell */
function cellReplaceText(sheet: GoogleAppsScript.Spreadsheet.Sheet, cell: string, template: string, replacement: string) {
    Logger.log("Sostituzione nella cella " + cell + " di " + template + " con " + replacement);
    const dateCell = sheet.getRange(cell);
    const dateCellContent = dateCell.getDisplayValue().replace(template, replacement);
    sheet.getRange(cell).setValue(dateCellContent);
}

/** Include a file in template */
function include(filename: string) {
    return HtmlService.createHtmlOutputFromFile(filename)
        .getContent();
}

/** Update status column with filename */
function updateState(ss: GoogleAppsScript.Spreadsheet.Sheet, row: number, newStatus: string) {
    const currStatus = parseInt(ss.getRange(row, config.statusCol).getValue().toString());
    Logger.log("Aggiorno lo stato alla riga " + row + " da " + currStatus + " a " + newStatus);
    ss.getRange(row, config.statusCol).setValue(newStatus);
    ss.getRange(row, config.statusCol).setBackground("#FFFF00");
}

/** Clear script */
function clearSheet() {
    const file = SpreadsheetApp.openById(config.invoicesFileID);
    const ss = file.getSheets()[0];
    ss.getRange(2, 2, config.lastRow).setValue(false);
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

/** return date as timestamp */
function getDateString() {
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

/** Check if a given filename exists */
function fileExists(filename) {
    var haBDs = DriveApp.getFilesByName(filename);
    if (!haBDs.hasNext()) {
        return true;
    }
    return false;
}

/** Read config data from "Config" sheet */
function readConfig() {
    SpreadsheetApp.getActiveSpreadsheet().toast(new Date().toLocaleString() + '\nLettura configurazione');
    let configuration: ConfigData = {
        /** Docs IDs */
        invoicesFileID: readConfigValue("ID file fatture"), //ID del file contenete le fatture
        invoiceTemplateID: readConfigValue("ID template fattura"), //ID del template di richiesta fatturazione
        invoiceFolderID: readConfigValue("ID cartella fatture"), //ID del template di richiesta fatturazione
        //Logger.log("Cartella Fatture ID["+invoiceFolderID+"]");
        /** Invoice file rows */
        firstRow: parseInt(readConfigValue("Prima riga fatture")),
        lastRow: parseInt(readConfigValue("Ultima riga fatture")),
        statusCol: parseInt(readConfigValue("Colonna stato fattura")),
        iRow: parseInt(readConfigValue("Riga dati template")),
        // Folders
        invoicesFolderID: readConfigValue("ID cartella Fatture"),
        // Bookmarks
        fiscalName: readConfigValue("Segnaposto Ragione Sociale"),
        address: readConfigValue("Segnaposto Indirizzo"),
        sapcode: readConfigValue("Segnaposto Codice SAP"),
        cigCode: readConfigValue("Segnaposto Codice CIG"),
        orderNr: readConfigValue("Segnaposto Nr Ordine"),
        salesCode: readConfigValue("Segnaposto Uff. Vendite"), //OLSE - OLSD
        date: readConfigValue("Segnaposto Data"),
        ewbs: readConfigValue("Segnaposto EWBS"), //A200V01920C030120000_001 
        channel: readConfigValue("Segnaposto Canale"), // Es: IT01/Z2
        // Cell addresses in template sheet
        cellAddress: readConfigValue("Template Cella indirizzo"),
        cellData1: readConfigValue("Template Cella data 1"),
        cellData2Col: readConfigValue("Template Col data 2"),
        cellData2Row: readConfigValue("Template Riga data 2"),
        cellCigCode: readConfigValue("Template Cella CIG"),
        cellOrderNr1: readConfigValue("Template Cella Ordine"),
        cellName: readConfigValue("Template Cella Ragione Sociale"),
        cellSapCode: readConfigValue("Template Cella Codice SAP"),
        cellEwbs: readConfigValue("Template Cella EWBS"),
        cellSalesCode: readConfigValue("Template Cella Ufficio Vendite"),
        cellChannel: readConfigValue("Template Cella Canale"),
        totalStartCol: parseInt(readConfigValue("Colonna Totale")),
        qtyStartCol: parseInt(readConfigValue("Colonna Quantità")),
        exampleRow: parseInt(readConfigValue("Riga Esempio")),
        invoiceRowStartCol: parseInt(readConfigValue("Colonna iniziale fattura")),
        invoiceRowColsNr: parseInt(readConfigValue("Numero Colonne Fattura")),
    }
    SpreadsheetApp.getActiveSpreadsheet().toast(new Date().toLocaleString() + '\nLettura configurazione completata');
    return configuration;
}

/* 
 * TYPES AND INTERFACES 
 */

/** Add a method to Date object */
declare interface Date {
    yyyymmdd(): string;
}

/** Interface which defines an invoice item */
interface InvoiceItem {
    year: number,
    month: number,
    sapCode: string;
    cigCode: string;
    ewbs: string;
    fiscalName: string;
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

/** Interface for Configuration data */
interface ConfigData {
    /** Docs IDs */
    invoicesFileID: string;
    invoiceTemplateID: string;
    invoiceFolderID: string;
    /** Invoices sheet rows */
    firstRow: number;
    lastRow: number;
    statusCol: number;
    iRow: number;
    // Folders
    invoicesFolderID: string;
    // Bookmarks
    fiscalName: string;
    address: string;
    sapcode: string;
    cigCode: string;
    orderNr: string;
    salesCode: string;
    date: string;
    ewbs: string;
    channel: string;
    // Cell addresses in template sheet
    cellAddress: string;
    cellData1: string;
    cellData2Col: string;
    cellData2Row: string;
    cellCigCode: string;
    cellOrderNr1: string;
    cellName: string;
    cellSapCode: string;
    cellEwbs: string;
    cellSalesCode: string;
    cellChannel: string;
    totalStartCol: number;
    qtyStartCol: number;
    exampleRow: number;
    invoiceRowStartCol: number;
    invoiceRowColsNr: number;
}