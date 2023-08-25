/**
 * Wordタグ定義シート内の<流し込み定義>内のテキストを2次元配列にて取得
 */
function _getRangeValuesOfWordItem(fileId) {
  // ワードの情報を記載したシートを取得
  const sheetName = _getWordItemSheetName(fileId);
  const wordSheet = getSheetByName(sheetName);
  if (wordSheet === null) {
    console.log(`${sheetName}が存在しません。`);
    return null;
  }
  const startRow = getRowNumber(wordSheet, 1, TITLE_INTO_TAG_MASTER) + 1;
  const lastRow = wordSheet.getLastRow();
  return wordSheet
    .getRange(startRow, 1, lastRow, WORD_ITEM_TAG_COL_NUMBERS["printType"])
    .getValues();
}

/**
 * Wordタグ定義シート内の背景色をを2次元配列にて取得
 */
function _getRangeBackgroundsOfWordItem(fileId) {
  // ワードの情報を記載したシートを取得
  const sheetName = _getWordItemSheetName(fileId);
  const wordSheet = getSheetByName(sheetName);
  if (wordSheet === null) {
    console.log(`${sheetName}が存在しません。`);
    return null;
  }
  const startRow = getRowNumber(wordSheet, 1, TITLE_INTO_TAG_MASTER) + 1;
  const lastRow = wordSheet.getLastRow();

  return wordSheet
    .getRange(startRow, 1, lastRow, WORD_ITEM_TAG_COL_NUMBERS["printType"])
    .getBackgrounds();
}

/**
 * ワードファイルの流し込み定義（WordItem）を取得
 */
function _getWordTagInfoListFromFileId(fileId) {
  let wordItemList = [];
  const startIndex = 1;
  const rangeValuesOfTagMaster = _getRangeValuesOfTagMaster();
  const rangeValuesOfWordItem = _getRangeValuesOfWordItem(fileId);
  if (rangeValuesOfWordItem === null) {
    return null;
  }
  for (const wordItemRowValues of rangeValuesOfWordItem) {
    const no = wordItemRowValues[WORD_ITEM_TAG_COL_NUMBERS["no"] - startIndex];
    const displayName =
      wordItemRowValues[WORD_ITEM_TAG_COL_NUMBERS["displayName"] - startIndex];
    const tagDefinitionId =
      wordItemRowValues[
        WORD_ITEM_TAG_COL_NUMBERS["tagDefinitionId"] - startIndex
      ];
    const key =
      wordItemRowValues[WORD_ITEM_TAG_COL_NUMBERS["key"] - startIndex];
    if (Number.isInteger(no) && displayName !== "") {
      const tagMasterRowValues = _getTagMasterRowValuesByItemId(
        tagDefinitionId,
        key,
        rangeValuesOfTagMaster
      );
      if (tagMasterRowValues !== null) {
        wordItemList.push(new WordItem(fileId, wordItemRowValues));
      }
    }
  }
  return wordItemList;
}

/**
 * ワードファイルの流し込み定義（WordItem）を取得
 */
function _getWordTagInfoandBackgroundsListFromFileId(
  fileId,
  backgroundColor,
  sqlkey,
  dicition
) {
  let wordItemList = [];
  const startIndex = 1;
  const rangeValuesOfTagMaster = _getRangeValuesOfTagMaster();
  const rangeValuesOfWordItem = _getRangeValuesOfWordItem(fileId);
  const backgroundArray = _getRangeBackgroundsOfWordItem(fileId);
  if (rangeValuesOfWordItem === null) {
    return null;
  }
  for (const wordItemRowValues of rangeValuesOfWordItem) {
    const no = wordItemRowValues[WORD_ITEM_TAG_COL_NUMBERS["no"] - startIndex];
    const displayName =
      wordItemRowValues[WORD_ITEM_TAG_COL_NUMBERS["displayName"] - startIndex];
    const tagDefinitionId =
      wordItemRowValues[
        WORD_ITEM_TAG_COL_NUMBERS["tagDefinitionId"] - startIndex
      ];
    const key =
      wordItemRowValues[WORD_ITEM_TAG_COL_NUMBERS["key"] - startIndex];
    if (Number.isInteger(no) && displayName !== "") {
      const tagMasterRowValues = _getTagMasterRowValuesByItemId(
        tagDefinitionId,
        key,
        rangeValuesOfTagMaster
      );
      if (tagMasterRowValues !== null)
        if (dicition == "追加") {
          wordItemList.push(new WordItem(fileId, wordItemRowValues, sqlkey));
        } else if (
          dicition == "修正" &&
          backgroundArray[no][0] === backgroundColor
        ) {
          wordItemList.push(new WordItem(fileId, wordItemRowValues, sqlkey));
        }
    }
  }
  return wordItemList;
}

function _getWordItemSheetName(fileId) {
  return SHEET_NAME_WORD.replace("{fileId}", fileId);
}

class WordItem {
  constructor(fileId, wordItemRowValues, sqlkey = 0) {
    const startIndex = 1;
    this.fileId = fileId;
    this.sqlkey = sqlkey;
    this.no = wordItemRowValues[WORD_ITEM_TAG_COL_NUMBERS["no"] - startIndex];
    this.displayName =
      wordItemRowValues[WORD_ITEM_TAG_COL_NUMBERS["displayName"] - startIndex];
    this.isRequired =
      wordItemRowValues[
        WORD_ITEM_TAG_COL_NUMBERS["isRequired"] - startIndex
      ] === MARU_STR;
    this.tagDefinitionId =
      wordItemRowValues[
        WORD_ITEM_TAG_COL_NUMBERS["tagDefinitionId"] - startIndex
      ];
    this.tagName =
      wordItemRowValues[WORD_ITEM_TAG_COL_NUMBERS["tagName"] - startIndex];
    this.category =
      wordItemRowValues[WORD_ITEM_TAG_COL_NUMBERS["category"] - startIndex];
    this.key = wordItemRowValues[WORD_ITEM_TAG_COL_NUMBERS["key"] - startIndex];
    this.printType =
      wordItemRowValues[WORD_ITEM_TAG_COL_NUMBERS["printType"] - startIndex];
  }
}
