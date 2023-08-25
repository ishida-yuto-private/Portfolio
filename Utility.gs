function getSheetByName(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(sheetName);
}

// 対象文字列が含まれるrow番号を取得する
function getRowNumber(sheet, targetCol, targetStr) {
  let rowNumber = 1;
  let lastRow = sheet.getLastRow();
  while (lastRow >= rowNumber) {
    const targetRange = sheet.getRange(rowNumber, targetCol);
    if (targetStr == targetRange.getValue()) {
      break;
    }
    rowNumber++;
  }
  return rowNumber;
}
