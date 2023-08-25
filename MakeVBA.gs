/**
 * VBA生成
 */
function makeWordVBA() {
  // Wordテンプレートファイルマスタから情報取得
  const wordInfoList = _getWordInfoList();
  // 描画
  makeVBA(wordInfoList);
}

function makeVBA(wordInfo) {
  if (!wordInfo) {
    return;
  }
  makeImageVBA(wordInfo);
  Browser.msgBox("VBAの生成が完了しました。");

  return;
}

function makeImageVBA(wordInfo) {
  for (let i = 0; i < wordInfo.length; i++) {
    let sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    let num = i + 1;
    let name = "画像作成VBA(" + num + ")";
    sheet.setName(name);
    // 物件拡張情報定義から情報取得
    let wordItem = _getWordItem(num);
    // 描画
    _makeWordImageVBA(wordItem, num);
  }
}

/**
 * 設計書画像用VBAの出力
 */
function _makeWordImageVBA(wordItem, fileId) {
  let wordSheet = getSheetByName("画像作成VBA(" + fileId + ")");
  let targetRange = wordSheet.getRange(1, 1);
  wordSheet.clear();
  if (wordItem.length == 0) {
    return;
  }
  let sqlValueStrs = [];

  sqlValueStrs.push(`  replaceWord "{{${item.tagName}}}", "(${item.no})"`);

  for (const item of wordItem) {
    sqlValueStrs.push(`  highlight "(${item.no})"`);
  }
  let sqlStr = VBA_SCRIPT_1;
  sqlStr += sqlValueStrs.join("\n");
  sqlStr += VBA_SCRIPT_2;
  targetRange.setValue(sqlStr);
}
