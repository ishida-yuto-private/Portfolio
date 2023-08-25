function makeSql(dicition, backgroundColor) {
  const wordInfo = makeSqlForWord(backgroundColor, dicition);
  if (!wordInfo) {
    return;
  }
  if (dicition == "追加") {
    makeReplaseVba(wordInfo);
  }
  makeSqlForWordTagInfo(backgroundColor, dicition);
  makeSqlForTagDefinition(backgroundColor, dicition);
  makeSqlForContractExtraInfo(backgroundColor, dicition);
  Browser.msgBox("SQLの生成が完了しました。");

  return;
}

function trashSheet() {
  var sheets = SpreadsheetApp.getActive().getSheets();
  for (var i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().indexOf("画像作成VBA(") == 0) {
      SpreadsheetApp.getActive().deleteSheet(sheets[i]);
    }
  }
}

/**
 * 背景色クリアのみの処理
 */
function clearBackgrandsColor(colorToClear) {
  let spreadsheets = [];

  for (let i = 0; i < SHEET_NAME_BACKGRAND_COLORS.length; i++) {
    spreadsheets.push(getSheetByName(SHEET_NAME_BACKGRAND_COLORS[i]));
    for (var sheet of spreadsheets) {
      if (sheet != null) {
        let range = sheet.getDataRange();
        let values = range.getValues();
        for (let row = 0; row < values.length; row++) {
          for (let col = 0; col < values[row].length; col++) {
            var cell = range.getCell(row + 1, col + 1);
            var bgColor = cell.getBackground();
            if (bgColor == colorToClear) {
              cell.setBackground(null);
            }
            if (bgColor == colorToClear && (i == 1 || i == 2) && col == 2) {
              cell.setBackground("#e6e6e6");
            }
          }
        }
      }
    }
  }
}

function clearBackgrandsTagColor(color) {
  // 「OK」と「キャンセル」
  var fileId = Browser.inputBox(
    "一番数値の小さいファイルIDを入力してください",
    Browser.Buttons.OK_CANCEL
  );
  var lastFileId = Browser.inputBox(
    "一番数値の大きいファイルIDを入力してください",
    Browser.Buttons.OK_CANCEL
  );

  for (var fileId; fileId <= lastFileId; fileId++) {
    // ワードの情報を記載したシートを取得
    const sheetName = _getWordItemSheetName(fileId);
    const wordSheet = getSheetByName(sheetName);
    if (wordSheet === null) {
      console.log(`${sheetName}が存在しません。`);
      continue;
    }
    const startRow = getRowNumber(wordSheet, 1, TITLE_INTO_TAG_MASTER) + 1;
    const lastRow = wordSheet.getLastRow();
    var backColor;
    if (wordSheet != null) {
      let range = wordSheet.getDataRange();
      let values = range.getValues();
      for (let row = startRow; row < values.length; row++) {
        for (let col = 0; col < values[row].length; col++) {
          backColor = range.getCell(row + 1, col + 1).getBackground();
          if (backColor == color) {
            range.getCell(row + 1, col + 1).setBackground(null);
          }
        }
      }
    }
  }
  return;
}

/**
 * Word用VBAの出力
 */
function _renderWordItem(wordItem, fileId) {
  let wordSheet = getSheetByName("画像作成VBA(" + fileId + ")");
  let targetRange = wordSheet.getRange(1, 1);
  wordSheet.clear();
  if (wordItem.length == 0) {
    return;
  }
  let sqlValueStrs = [];
  for (const item of wordItem) {
    sqlValueStrs.push(`  replaceWord "{{${item.no}}}", "{{${item.tagName}}}"`);
  }
  let sqlStr = VBA_SCRIPT_1;
  sqlStr += sqlValueStrs.join("\n");
  sqlStr += VBA_SCRIPT_2;
  targetRange.setValue(sqlStr);
}

/**
 * ワードファイルの流し込み定義（WordItem）を取得
 */
function _getWordItem(fileId) {
  let wordItemList = [];
  const startIndex = 1;
  const rangeValuesOfWordTempalteItem = _getRangeValuesOfWordItem(fileId);
  if (rangeValuesOfWordTempalteItem === null) {
    return null;
  }
  for (const wordItemRowValues of rangeValuesOfWordTempalteItem) {
    const no = wordItemRowValues[WORD_ITEM_TAG_COL_NUMBERS["no"] - startIndex];
    const displayName =
      wordItemRowValues[WORD_ITEM_TAG_COL_NUMBERS["displayName"] - startIndex];
    if (Number.isInteger(no) && displayName !== "") {
      wordItemList.push(new WordItem(fileId, wordItemRowValues));
    }
  }
  return wordItemList;
}

/**
 * hoge_wordのUpdate文生成
 */
function makeSqlForWord(backgroundColor, dicition) {
  // Wordファイルマスタから情報取得
  const wordInfoList = _getWordInfoListToCreateSql(backgroundColor);
  if (!wordInfoList) {
    Browser.msgBox("表示ファイル名に不正な文字が含まれています");
    return false;
  }
  // 描画
  if (wordInfoList.length != 0) {
    _renderWordToCreateSql(wordInfoList, dicition);
  }
  return wordInfoList;
}

/**
 * word_tag_infoのUpdate文を生成
 */
function makeSqlForWordTagInfo(backgroundColor, dicition) {
  // ワード情報を取得
  const wordInfoList = _getWordInfoListToCreateSql(backgroundColor);
  let wordTagInfoList = [];
  for (const wordInfo of wordInfoList) {
    // 各ワードの流し込み情報を取得
    wordTagInfoList.push(
      _getWordTagInfoandBackgroundsListFromFileId(
        wordInfo.fileId,
        backgroundColor,
        wordInfo.sqlkey,
        dicition
      )
    );
    console.log(wordInfo.fileId);
  }
  // 描画
  if (wordTagInfoList.length != 0) {
    _renderWordTagInfoToCreateSql(wordTagInfoList, dicition);
  }
  return wordTagInfoList;
}

/**
 * tag_definitionのUpdate文を生成
 */
function makeSqlForTagDefinition(backgroundColor, dicition) {
  // 契約情報定義から情報取得
  const tagMaster = _getTagMasterSQL(backgroundColor);
  // 描画
  if (tagMaster.length != 0) {
    _renderTagDefinitionToCreateSql(tagMaster, dicition);
  }
  return tagMaster;
}

/**
 * contract_extra_infoのUpdate文を生成
 */
function makeSqlForContractExtraInfo(backgroundColor, dicition) {
  // 契約拡張情報定義から情報取得
  const contractItemExtraInfo = _getContractSQLItemExtraInfo(backgroundColor);
  // 描画
  if (contractItemExtraInfo.length != 0) {
    _renderContractItemExtraInfoToCreateSql(contractItemExtraInfo, dicition);
  }
  return contractItemExtraInfo;
}
