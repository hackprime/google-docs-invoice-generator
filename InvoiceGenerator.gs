/** SETTINGS */

var
    homeFolderId = '0B6FfG98UNCKMY3hpbUlhelpIN00',
    // Base number for invoices. Next number = count_of_invoices + 1
    invoiceNumberBase = 1,
    // ID of invoices folder (Google Drive folder)
    invoicesFolderId = '0B6FfG98UNCKMUDFVZFRmdUdyaVU',
    // ID of invoice template (Google Document)
    invoiceTemplateId = '1QVEk6QlyZ3d60iH1YWDQ1URYYZ4dgQUrCFLU06Z5OoU',
    // New invoice filename template.
    invoiceFilenameTemplate = 'Invoice No.{invoiceID} of {datetime}',
    // Invoice field to Sheet index mapping
    dataRowToSheetIdx = {
        'name': 3,
        'developer': 4,
        'number': 5,
        'price': 10,
        'quantity': 6,
        'amount': 11,
        'manager': 15
    },
    // Invoice items table generation rules
    invoiceTableFields = [
        {
            index: 0,
            generator: function (data) {
                return data.idx;
            },
        }, {
            index: 1,
            generator: function (data) {
                return [data.name, data.developer, data.number].join(' ');
            },
        }, {
            index: 2,
            generator: function (data) {
                return 'units';
            },
        }, {
            index: 3,
            generator: function (data) {
                return data.price || 0;
            },
        }, {
            index: 4,
            generator: function (data) {
                return data.quantity || 0;
            },
        }, {
            index: 5,
            generator: function (data) {
                return (data.price || 0) * (data.quantity || 0);
            }
        }
    ],
    /**
     * Function that converts total cost of invoice into words representation,
     * e.g. 223 -> 'two hundred twenty three'
     * NOTE: be careful, function numberInWords() converts number into Russian
     *       sentance. If you need English version - implement it by yourself :)
     */
     invoiceTotalCostInWordsConverter = numberInWords;


/** BEFORE START */

/**
 * String interpolation property
 * '{a} and {b}'.supplant({a: 'you', b: 'me'})  // 'you and me'
 */
String.prototype.supplant = function (o) {
    return this.replace(/\{([^{}]*)\}/g, function (a, b) {
        var r = o[b];
        if (typeof r === 'string' || typeof r === 'number') {
            return r;
        } else {
            return a;
        }
    });
};

/**
 * Triggers on parent sheet oppening
 * '{a} and {b}'.supplant({a: 'you', b: 'me'})  // 'you and me'
 */
function onOpen() {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var entries = [{
        name: "Generate Invoice",
        functionName: "generateInvoice"
    }];
    spreadsheet.addMenu("Actions", entries);
}


/** UTILS */

/**
 * Prepend additional leading zero to number if length of number equals 1
 * leadingZero('4')  // '04'
 * leadingZero(14)  // '14'
 *
 * @param {string|number} number or string with size of 1 or 2 characters
 * @return {string} modified string with size of 2 characters
 */
function leadingZero(number) {
    if (String(number).length === 1) {
        return '0' + number;
    }
    return number;
}

/**
 * Returns string representation of Date() in specified format.
 * formattedDateTime('{y}.{m}.{d} {h}:{i}')  // '2016.02.26 3:24'
 *
 * @param {string} supplant-like format
 * @return {string} formatted date
 */
function formattedDateTime(format) {
    var date = new Date();
    return format.supplant({
        y: date.getYear(),
        m: leadingZero(date.getMonth() + 1),
        d: leadingZero(date.getDate()),
        h: leadingZero(date.getHours()),
        i: leadingZero(date.getMinutes())
    });
}

/**
 * Generates new invoice ID using formula:
 *     (f + invoiceNumberBase)
 * f - number of invoice files in invoices folder
 *
 * @return {number} new invoice ID
 */
function getNewInvoiceID() {
    var invoiceFiles = DriveApp.getFolderById(invoicesFolderId).getFiles(),
        filesCount = 0;

    // TODO: getting growing complexity ~O(n), bad practice.
    //       Where to store a number of files?
    while (invoiceFiles.hasNext()) {
        filesCount += 1;
        invoiceFiles.next();
    }
    return filesCount + invoiceNumberBase;
}

/**
 * Generates new invoice file name using invoiceFilenameTemplate as template:
 *
 * @return {string} new invoice file name
 */
function getNewInvoiceFileName() {
    return invoiceFilenameTemplate.supplant({
        invoiceID: getNewInvoiceID(),
        datetime: formattedDateTime('{y}.{m}.{d} {h}:{i}')
    });
}

/**
 * Extracts data from selected sheet cells to array of objects,
 * defined in dataRowToSheetIdx.
 *
 * @param {array} selected sheet cells
 * @return {array} extracted invoice rows data
 */
function getInvoiceData(selectedCells) {
    var invoiceData = [],
        idx,
        key,
        item;

    for (idx = 0; idx < selectedCells.length; idx += 1) {
        item = {idx: idx + 1}
        for (key in dataRowToSheetIdx) {
            item[key] = selectedCells[idx][dataRowToSheetIdx[key]];
        }
        invoiceData.push(item);
    }
    return invoiceData;
}

/**
 * Creates and returns new invoice document at invoicesFolderId
 *
 * @param {string} name of new file
 * @return {object} new document object
 */
function createInvoiceDocument(filename) {
    var doc = DriveApp.getFileById(invoiceTemplateId).makeCopy(filename);
    DriveApp.getFolderById(invoicesFolderId).addFile(doc);
    DriveApp.getFolderById(homeFolderId).removeFile(doc);
    return DocumentApp.openById(doc.getId());
}


/** VIEWS AND TRIGGERS */

/**
 * Main invoice generation function.
 */
function generateInvoice() {
    var sheet = SpreadsheetApp.getActiveSheet(),
        rows = sheet.getActiveRange(),
        selectedCells = rows.getValues(),
        invoiceData = getInvoiceData(selectedCells),
        invoiceID = getNewInvoiceID();

    if (invoiceData !== null) {
        var html,
            doc = createInvoiceDocument(getNewInvoiceFileName());
        renderInvoice(doc, {
            invoiceID: invoiceID,
            invoiceData: invoiceData,
            date: formattedDateTime('{d}.{m}.{y}')
        });

        html = HtmlService.createHtmlOutput([
            '<a href="',
            doc.getUrl(),
            '" target="_blank">Link to Invoice</a>'
        ].join(''))
            .setWidth(300)
            .setHeight(100)
            .setTitle('Invoice ready!');

        SpreadsheetApp.getActive().show(html);
    }
}

/**
 * Renders invoice document.
 *
 * @param {object} document object created from template
 * @param {object} rendering context
 */
function renderInvoice(doc, context) {
    var table = doc.getTables()[0],
        managers = [],
        amount = 0,
        rowData,
        row,
        i;

    for (i = 0; i < context.invoiceData.length; i += 1) {
        rowData = context.invoiceData[i];
        row = table.getRow(1).copy();
        invoiceTableFields.forEach(function (genItem) {
            row.getCell(genItem.index).setText(genItem.generator(rowData));
            if (managers.indexOf(rowData.manager) === -1) {
                managers.push(rowData.manager);
            }
        });
        table.insertTableRow(rowData.idx + 1, row);
    }
    table.removeRow(1);
    doc.replaceText("{date}", context.date);
    doc.replaceText("{number}", context.invoiceID);
    doc.replaceText("{manager}", managers.join(', '));
    amount = context.invoiceData.map(function (el) {
        return el.price * el.quantity;
    }).reduce(function (prev, cur) {
        return prev + cur;
    });
    doc.replaceText("{amount}", amount);
    doc.replaceText("{amount_string}",
        invoiceTotalCostInWordsConverter(amount).toUpperCase());

    doc.saveAndClose();
}
