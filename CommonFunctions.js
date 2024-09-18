const server = '18.171.8.199';
const port = 3306;
const dbName = 'jtl_cavingcrew_com';
const username = 'gsheets';
const password = 'athohKeecieD8xees';
const url = `jdbc:mysql://${server}:${port}/${dbName}`;
const cc_location = "placeholder";
const apidomain = "cavingcrew.com";
const apiusername = "ck_91675da0323ed5e5f5704a79a59288950db68efc";
const apipassword = "cs_7be15b56e20ef6006720147f4ce44ff472039328";

const colors = {
    lightRed: "#ffe6e6",
    lightGreen: "#e6ffe6",
    lightYellow: "#ffd898",
    lightBlue: "#e0ffff",
    grey: "#a9a9a9",
    pink: "#ff75d8",
    yellow: "#fad02c",
    white: "#FFFFFF"
};

function setupSheet(name) {
    const spreadsheet = SpreadsheetApp.getActive();
    const sheet = spreadsheet.getSheetByName(name);
    sheet.clearContents();
    sheet.clearFormats();
    return sheet;
}

function setupCell(name, range) {
    const spreadsheet = SpreadsheetApp.getActive();
    const sheet = spreadsheet.getSheetByName(name);
    let cellValue = sheet.getRange(range).getValue();

    if (isNaN(cellValue) || cellValue === "") {
        // Rerun eventListing
        eventListing();

        // Try again
        cellValue = sheet.getRange(range).getValue();

        if (isNaN(cellValue) || cellValue === "") {
            throw new Error("Invalid event selected - please try again");
        }
    }

    return cellValue;
}

function appendToSheet(sheet, results) {
    const metaData = results.getMetaData();
    const numCols = metaData.getColumnCount();
    const rows = [];

    // First row with column labels
    const colLabels = [];
    for (let col = 0; col < numCols; col++) {
        colLabels.push(metaData.getColumnLabel(col + 1));
    }
    rows.push(colLabels);

    // Remaining rows with results
    while (results.next()) {
        const row = [];
        for (let col = 0; col < numCols; col++) {
            row.push(results.getString(col + 1));
        }
        rows.push(row);
    }

    sheet.getRange(1, 1, rows.length, numCols).setValues(rows);

    // Set the font size of the rows with column labels to 14
    sheet.getRange(1, 1, 1, numCols).setFontSize(14);
    sheet.autoResizeColumns(1, numCols);
}

function getColumnRange(sheet, columnHeader) {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnIndex = headers.indexOf(columnHeader) + 1;
    if (columnIndex === 0) {
        throw new Error(`Column '${columnHeader}' not found in the sheet.`);
    }
    return sheet.getRange(2, columnIndex, sheet.getLastRow() - 1, 1);
}

function setColoursFormat(sheet, columnHeader, search, colour) {
    const range = getColumnRange(sheet, columnHeader);
    const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains(search)
        .setBackground(colour)
        .setRanges([range])
        .build();
    const rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
}

function setTextFormat(sheet, columnHeader, search, colour) {
    const range = getColumnRange(sheet, columnHeader);
    const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenTextContains(search)
        .setFontColor(colour)
        .setRanges([range])
        .build();
    const rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
}

function setWrapped(sheet, columnHeader) {
    const range = getColumnRange(sheet, columnHeader);
    range.setWrap(true);
    sheet.setColumnWidth(range.getColumn(), 300); // Set column width to 300 pixels
}

function setColoursFormatLessThanOrEqualTo(sheet, columnHeader, search, colour) {
    search = Number(search);
    const range = getColumnRange(sheet, columnHeader);
    const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenNumberLessThanOrEqualTo(search)
        .setBackground(colour)
        .setRanges([range])
        .build();
    const rules = sheet.getConditionalFormatRules();
    rules.push(rule);
    sheet.setConditionalFormatRules(rules);
}

function setNumberFormat(sheet, columnHeader, format) {
    const range = getColumnRange(sheet, columnHeader);
    range.setNumberFormat(format);
}

function setColumnWidth(sheet, columnHeader, width) {
    const range = getColumnRange(sheet, columnHeader);
    sheet.setColumnWidth(range.getColumn(), width);
}

function getIP() {
    const url = "http://api.ipify.org";
    const response = UrlFetchApp.fetch(url);
    Logger.log(response.getContentText());
}

function makeReport(stmt, reportConfig) {
    const sheet = setupSheet(reportConfig.sheetName);

    const results = stmt.executeQuery(reportConfig.query);

    appendToSheet(sheet, results);

    if (reportConfig.formatting) {
        reportConfig.formatting.forEach(format => {
            if (format.type === 'color') {
                setColoursFormat(sheet, format.column, format.search, format.color);
            } else if (format.type === 'text') {
                setTextFormat(sheet, format.column, format.search, format.color);
            } else if (format.type === 'wrap') {
                setWrapped(sheet, format.column);
            } else if (format.type === 'numberFormat') {
                setNumberFormat(sheet, format.column, format.format);
            } else if (format.type === 'colorLessThanOrEqual') {
                setColoursFormatLessThanOrEqualTo(sheet, format.column, format.value, format.color);
            } else if (format.type === 'columnWidth') {
                setColumnWidth(sheet, format.column, format.width);
            }
        });
    }

    results.close();
}
