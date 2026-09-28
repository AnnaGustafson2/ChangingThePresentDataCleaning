function cleanSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  sheets.forEach(sheet => {
    convertFormulas(sheet);
    deleteEmptyRowsBelow(sheet);
    deleteEmptyRowsBelow(sheet);
  });
}

function convertFormulas(sheet) {
  const range = sheet.getDataRange();
  range.setValues(range.getValues());
}

function deleteEmptyRowsBelow(sheet) {
  var maxRows = sheet.getMaxRows();
  var lastRow = sheet.getLastRow();
      
  if (maxRows - lastRow > 0) {
    sheet.deleteRows(lastRow + 1, maxRows - lastRow);
  }

}

function deleteEmptyColumnsRight(sheet) {
  const maxColumns = sheet.getMaxColumns();
  const lastColumn = sheet.getLastColumn();
  
  if (maxColumns > lastColumn) {
    const columnsToDelete = maxColumns - lastColumn;
    sheet.deleteColumns(lastColumn + 1, columnsToDelete);
  }
}