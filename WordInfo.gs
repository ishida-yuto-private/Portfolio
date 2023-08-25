/**
 * Wordファイル定義シート内のテキストを2次元配列にて取得
 */
function _getRangeValuesOfWordMaster() {
  const masterSheet = getSheetByName(SHEET_NAME_WORD_MASTER);
  const lastRow = masterSheet.getLastRow();
  let startRow = 2;
  let startCol = 1;
  return masterSheet
    .getRange(startRow, startCol, lastRow, WORD_MASTER_COL_NUMBERS["sqlkey"])
    .getValues();
}

// Wordファイル定義の背景色を取得
function _getRangeBackgroundsOfWordMaster() {
  const masterSheet = getSheetByName(SHEET_NAME_WORD_MASTER);
  const lastRow = masterSheet.getLastRow();
  let startRow = 2;
  let startCol = 1;
  return masterSheet
    .getRange(
      startRow,
      startCol,
      lastRow,
      WORD_MASTER_COL_NUMBERS["displayFileName"]
    )
    .getBackgrounds();
}

function _getWordInfoList() {
  const rangeValues = _getRangeValuesOfWordMaster();
  const startIndex = 1;
  let wordInfoList = [];
  for (const rowValues of rangeValues) {
    const displayOrder =
      rowValues[WORD_MASTER_COL_NUMBERS["displayOrder"] - startIndex];
    const displayFileName =
      rowValues[WORD_MASTER_COL_NUMBERS["displayileName"] - startIndex];
    if (Number.isInteger(displayOrder) && displayFileName !== "") {
      wordInfoList.push(new WordInfo(rowValues));
    }
  }
  return wordInfoList;
}

function _getWordInfoListToCreateSql(backgroundColor) {
  const rangeValues = _getRangeValuesOfWordMaster();
  const rangeBackgrands = _getRangeBackgroundsOfWordMaster();
  const startIndex = 1;
  const filteredArray = [];

  for (let i = 0; i < rangeValues.length; i++) {
    if (rangeBackgrands[i][0] === backgroundColor) {
      filteredArray.push(rangeValues[i]);
    }
  }

  let wordInfoList = [];
  for (const rowValues of filteredArray) {
    const displayOrder =
      rowValues[WORD_MASTER_COL_NUMBERS["displayOrder"] - startIndex];
    const displayFileName =
      rowValues[WORD_MASTER_COL_NUMBERS["displayFileName"] - startIndex];

    for (let n of NG_LIST) {
      // 以下どれかに当てはまる場合falseにする
      if (
        displayFileName.includes(n) ||
        !Number.isInteger(displayOrder) ||
        displayFileName === ""
      ) {
        return false;
      }
    }
    wordInfoList.push(new WordInfo(rowValues));
  }
  return wordInfoList;
}

class WordInfo {
  constructor(rowValues) {
    // colnumberが１始まりのため、rowValuesから値を取得する場合はindexから1を減らす
    const startIndex = 1;
    this.displayOrder =
      rowValues[WORD_MASTER_COL_NUMBERS["displayOrder"] - startIndex];
    this.displayFileName =
      rowValues[WORD_MASTER_COL_NUMBERS["displayFileName"] - startIndex];
    this.phaseId = rowValues[WORD_MASTER_COL_NUMBERS["phaseId"] - startIndex];
    this.fileId = rowValues[WORD_MASTER_COL_NUMBERS["fileId"] - startIndex];
    this.wordFileName =
      rowValues[WORD_MASTER_COL_NUMBERS["wordFileName"] - startIndex];
    // this.isAllArticleFlg = 1;
    this.sqlkey = rowValues[WORD_ITEM_TAG_COL_NUMBERS["sqlkey"]];
  }
}
