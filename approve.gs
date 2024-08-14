function doGet(e) {
    // Find the ID of your spread sheet in a URL like below
    // example: https://docs.google.com/spreadsheets/d/{ID of spread sheet}/
    let spreadSheetId = "{Input your spreadsheet ID}";
    let spreadsheet = SpreadsheetApp.openById(spreadSheetId);
    let activeSheet = spreadsheet.getActiveSheet();

    const startRow = 1;
    const startColumn = 1
    const lastRow = 1;
    let lastColumn = activeSheet.getLastColumn();
    // target row: from first row to first row (first row only as it's a header)
    // target colum: from first colum to last colum (all columns)
    let header = activeSheet.getRange(startRow, startColumn, lastRow,lastColumn).getDisplayValues();
    const colmunName = "承認結果";
    const colmunIndex = header[0].indexOf(colmunName) + 1;

    let html = "";
    if (e.parameters.status == "approve") {
        html = "<h1>承認しました。</h1>";
        activeSheet.getRange(e.parameters.row, colmunIndex).setValue('承認');
    } else {
        html = "<h1>否認しました。</h1>";
        activeSheet.getRange(e.parameters.row, colmunIndex).setValue('否認');
    }

    return HtmlService.createHtmlOutput(html);
}

