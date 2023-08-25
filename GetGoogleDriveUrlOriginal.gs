function getListInFolder() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var confSheet = ss.getSheetByName("config");
  var sheetName = ss.getSheetByName("getOriginalUrl");
  var lastRow = sheetName.getLastRow();
  var lastCol = sheetName.getLastColumn();
  var rangeList = sheetName.getRange(2, 1, lastRow, lastCol);

  var folder_id = confSheet.getRange("B1").getValue();

  var url = "https://drive.google.com/drive/folders/" + folder_id;
  var paths = url.split("/"); // Separate URL into an array of strings by separating the string into substrings
  var folderId = paths[paths.length - 1];
  var folder = DriveApp.getFolderById(folderId);
  var childFolders = folder.getFolders();
  var files = folder.getFiles();
  var list = [];
  var rowIndex = 2;
  var colIndex = 1;

  // 初期化処理
  rangeList.clearContent();

  while (files.hasNext()) {
    var buff = files.next();
    list.push([buff.getName(), buff.getUrl()]);
  }

  while (childFolders.hasNext()) {
    var buff = childFolders.next();
    list.push([buff.getName(), buff.getUrl()]);
  }

  range = sheetName.getRange(rowIndex, colIndex, list.length, list[0].length);

  // 対象の範囲にまとめて書き出します
  range.setValues(list);
}
