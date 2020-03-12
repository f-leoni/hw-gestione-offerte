/*
 * CONFIGURAZIONE
 */
const debug = false;
let config: ConfigData;
let filename = "";

/** Inizializzazione e installazione */
function onInstall(e: any) {
    SpreadsheetApp.getActiveSpreadsheet().toast(new Date().toLocaleString() + '\nLettura configurazione');
    config = readConfig();
    SpreadsheetApp.getActiveSpreadsheet().toast(new Date().toLocaleString() + '\nLettura configurazione completata');
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
    SpreadsheetApp.getActiveSpreadsheet().toast(new Date().toLocaleString() + '\nLettura dati');
    var currentInvoices = LeggiDati(config.firstRow, config.lastRow);
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
    SpreadsheetApp.getActiveSpreadsheet().toast(new Date().toLocaleString() + '\nCreazione modulo');
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
    cellReplaceText(sheet, config.cellOrderNr1, config.nrOrdine, orderNr);
    cellReplaceText(sheet, config.cellCigCode, config.cigCode, cigCode);
    cellReplaceText(sheet, config.cellEwbs, config.ewbs, ewbs);
    cellReplaceText(sheet, config.cellName, config.ragioneSociale, ragioneSociale);
    cellReplaceText(sheet, config.cellAddress, config.indirizzo, orderAddress);
    cellReplaceText(sheet, config.cellOrderNr1, config.nrOrdine, filename);
    cellReplaceText(sheet, config.cellSalesCode, config.salesCode, salesCode);
    cellReplaceText(sheet, config.cellData1, config.data, dateString);
    cellReplaceText(sheet, config.cellCanale, config.canale, canaleString);
    const currData2Row = config.cellData2Row + currentInvoices.length + 1;
    cellReplaceText(sheet, config.cellData2Col + currData2Row, config.data, dateString);
    /** Numero di righe da sommare per totale (numero item + 1) */
    const sumRows = currentInvoices.length + 1;
    /** Formula per totale */
    sheet.getRange(config.iRow + currentInvoices.length, config.totalStartCol, 1, 1)//(start row, start column, number of rows, number of columns
        .setFormulaR1C1("=sum(R[-" + sumRows + "]C[0]:R[-1]C[0])");
    /** Formula per Q.tà */
    sheet.getRange(config.iRow + currentInvoices.length, config.qtyStartCol, 1, 1)//(start row, start column, number of rows, number of columns
        .setFormulaR1C1("=sum(R[-" + sumRows + "]C[0]:R[-1]C[0])");
    /** Elimino riga di esempio */
    sheet.deleteRow(config.exampleRow);
    Logger.log("Segnaposto sostituiti");
    SpreadsheetApp.getActiveSpreadsheet().toast(new Date().toLocaleString() + '\nCreazione modulo completata');
    // Sposto la fattura nella cartella "Fatture" su Drive 
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
function LeggiRigheSelezionate(rigaIniziale: number = config.firstRow, rigaFinale: number = config.lastRow) {
    //Logger.log("controllo le righe dalla " + firstRow + " alla " + lastRow);
    let activeRows: number[] = [];
    const file = SpreadsheetApp.openById(config.invoicesFileID);
    const ss = file.getSheets()[0];
    /*Parametri: riga, colonna, nrighe, nrcolonne */
    const rangeToCheck = ss.getRange(rigaIniziale, config.invoiceRowStartCol, rigaFinale - rigaIniziale, 1).getValues();
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
    const file = SpreadsheetApp.openById(config.invoicesFileID);
    const ss = file.getSheets()[0];
    const activeRows: Array<number> = LeggiRigheSelezionate();
    for (let i: number = activeRows.length - 1; i >= 0; i--) {
        const riga = activeRows[i];
        // 21 columns starting with column 2, so B-V range 
        const rangeToCheck = ss.getRange(riga, config.invoiceRowStartCol, 1, config.invoiceRowColsNr);
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
    itemAddress: string, itemSapCode: string, itemCigCode: string, itemEwbsCode: string, itemSalesCode: string, itemDescription: string,
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

/** Controlla se esiste già un file con il nome passato come parametro */
function ChechFileExists(filename) {
    var haBDs = DriveApp.getFilesByName(filename);
    if (!haBDs.hasNext()) {
        return true;
    }
    return false;
}

/** Legge dati di configuraz<ione dal foglio "Config" */
function readConfig() {
    const configuration: ConfigData = {
        /** ID documenti  */
        invoicesFileID: ReadConfigValue("ID file fatture"), //ID del file contenete le fatture
        invoiceTemplateID: ReadConfigValue("ID template fattura"), //ID del template di richiesta fatturazione
        invoiceFolderID: ReadConfigValue("ID cartella fatture"), //ID del template di richiesta fatturazione
        //Logger.log("Cartella Fatture ID["+invoiceFolderID+"]");
        /** Righe file fatture */
        firstRow: parseInt(ReadConfigValue("Prima riga fatture")),
        lastRow: parseInt(ReadConfigValue("Ultima riga fatture")),
        statusCol: parseInt(ReadConfigValue("Colonna stato fattura")),
        iRow: parseInt(ReadConfigValue("Riga dati template")),
        // CARTELLE
        fattureFolderID: ReadConfigValue("ID cartella Fatture"),
        // SEGNAPOSTO
        ragioneSociale: ReadConfigValue("Segnaposto Ragione Sociale"),
        indirizzo: ReadConfigValue("Segnaposto Indirizzo"),
        sapcode: ReadConfigValue("Segnaposto Codice SAP"),
        cigCode: ReadConfigValue("Segnaposto Codice CIG"),
        nrOrdine: ReadConfigValue("Segnaposto Nr Ordine"),
        salesCode: ReadConfigValue("Segnaposto Uff. Vendite"), //OLSE - OLSD
        data: ReadConfigValue("Segnaposto Data"),
        ewbs: ReadConfigValue("Segnaposto EWBS"), //A200V01920C030120000_001 
        canale: ReadConfigValue("Segnaposto Canale"), // Es: IT01/Z2
        // RIFERIMENTI CELLE SHEET RICHIESTA FATTURA
        cellAddress: ReadConfigValue("Template Cella indirizzo"),
        cellData1: ReadConfigValue("Template Cella data 1"),
        cellData2Col: ReadConfigValue("Template Col data 2"),
        cellData2Row: ReadConfigValue("Template Riga data 2"),
        cellCigCode: ReadConfigValue("Template Cella CIG"),
        cellOrderNr1: ReadConfigValue("Template Cella Ordine"),
        cellName: ReadConfigValue("Template Cella Ragione Sociale"),
        cellSapCode: ReadConfigValue("Template Cella Codice SAP"),
        cellEwbs: ReadConfigValue("Template Cella EWBS"),
        cellSalesCode: ReadConfigValue("Template Cella Ufficio Vendite"),
        cellCanale: ReadConfigValue("Template Cella Canale"),
        totalStartCol: parseInt(ReadConfigValue("Colonna Totale")),
        qtyStartCol: parseInt(ReadConfigValue("Colonna Quantità")),
        exampleRow: parseInt(ReadConfigValue("Riga Esempio")),
        invoiceRowStartCol: parseInt(ReadConfigValue("Colonna iniziale fattura")),
        invoiceRowColsNr: parseInt(ReadConfigValue("Numero Colonne Fattura")),
    }
    return configuration;
}

/* 
 * TIPI E INTERFACCE 
 */

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

/** Interfaccia configurazione */
interface ConfigData {
    /** ID documenti  */
    invoicesFileID: string;
    invoiceTemplateID: string;
    invoiceFolderID: string;
    /** Righe file fatture */
    firstRow: number;
    lastRow: number;
    statusCol: number;
    iRow: number;
    // Cartelle
    fattureFolderID: string;
    // Segnaposto
    ragioneSociale: string;
    indirizzo: string;
    sapcode: string;
    cigCode: string;
    nrOrdine: string;
    salesCode: string;
    data: string;
    ewbs: string;
    canale: string;
    // Riferimenti celle sheet richiesta fattura
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
    cellCanale: string;
    totalStartCol: number;
    qtyStartCol: number;
    exampleRow: number;
    invoiceRowStartCol: number;
    invoiceRowColsNr: number;
}