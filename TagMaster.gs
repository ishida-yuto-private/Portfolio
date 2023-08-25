/**
 * ＜契約情報定義＞よりも下に契約情報からタグ定義を自動生成する
 */
function makeTags() {
  const tagMasterSheet = getSheetByName(SHEET_NAME_TAG_MASTER);

  let tagDefinitionId = 1;
  let rowNumber = getRowNumber(
    tagMasterSheet,
    1,
    TITLE_CONTRACT_EXTRA_INFO_MASTER
  );
  const lastRow = tagMasterSheet.getLastRow();

  // クリア
  rowNumber++;
  tagMasterSheet
    .getRange(rowNumber, 1, lastRow, TAG_MASTER_COL_NUMBERS["printType"])
    .setValue("");

  var range = tagMasterSheet.getRange("B:B"); // B列を取得
  var values = range.getValues().flat(); // 列の値を取得し、1次元配列に変換
  var numericValues = values.filter(function (value) {
    return typeof value === "number"; // 数値のみを抽出
  });
  tagDefinitionId += Math.max(...numericValues); // 数値の最大値を取得

  function setBorder(tagMasterSheet, x, y) {
    tagMasterSheet
      .getRange(x, y)
      .setBorder(true, true, true, true, false, false);
  }

  // Wordファイルマスタから情報取得してタグを描画
  const contractItemExtraInfo = _getContractItemExtraInfo();
  for (const contractItem of contractItemExtraInfo) {
    if (contractItem.name == "") {
      continue;
    }
    tagMasterSheet
      .getRange(rowNumber, TAG_MASTER_COL_NUMBERS["displayName"])
      .setValue(contractItem.name);
    tagMasterSheet
      .getRange(rowNumber, TAG_MASTER_COL_NUMBERS["tagDefinitionId"])
      .setValue(tagDefinitionId++);
    tagMasterSheet
      .getRange(rowNumber, TAG_MASTER_COL_NUMBERS["tagName"])
      .setValue(CEX_ADD + "_" + contractItem.id);
    tagMasterSheet
      .getRange(rowNumber, TAG_MASTER_COL_NUMBERS["category"])
      .setValue(TABLE_NAME_CONTRACT_EXTRA_INFO);
    tagMasterSheet
      .getRange(rowNumber, TAG_MASTER_COL_NUMBERS["key"])
      .setValue(contractItem.id);
    tagMasterSheet
      .getRange(rowNumber, TAG_MASTER_COL_NUMBERS["printType"])
      .setValue("");

    for (let i = 1; i < 100; i++) {
      if (i == 7) {
        continue;
      }
      setBorder(tagMasterSheet, rowNumber, i);
    }

    rowNumber++;
  }
}

// H列以降にワードファイルの名前及びタグの使用状況記載する
function makeTagUseStatus() {
  const tagMasterSheet = getSheetByName(SHEET_NAME_TAG_MASTER);
  // Wordファイルマスタから情報取得
  const wordInfoList = _getWordInfoList();

  let properties = PropertiesService.getScriptProperties();
  var colNumberKey = "colNumber"; //何行目まで処理したかを保存するときに使用するkey

  //途中から実行した場合、ここに何行目まで実行したかが入る
  var colNumber = parseInt(properties.getProperty(colNumberKey));
  if (!colNumber) {
    //初めて実行する場合はこっち
    colNumber = TAG_MASTER_COL_NUMBERS["useStatus"];
    // 値をクリア
    const lastCol = tagMasterSheet.getLastColumn();
    const lastRow = tagMasterSheet.getLastRow();
    tagMasterSheet
      .getRange(1, TAG_MASTER_COL_NUMBERS["useStatus"], lastRow, lastCol)
      .setValue("");
  }

  let startTime = new Date();
  //スクリプトプロパティにトリガーIDを保存するときに使用するkey名
  let triggerKey = "trigger";

  while (
    colNumber <
    wordInfoList.length + TAG_MASTER_COL_NUMBERS["useStatus"]
  ) {
    properties.setProperty(colNumberKey, colNumber);
    //開始時刻（startTime）と現時点の処理時点の時間を比較する
    let diff = parseInt((new Date() - startTime) / (1000 * 60));
    if (diff >= 5) {
      //トリガー(1分後)を登録する
      setTrigger(triggerKey, "makeTagUseStatus");
      return;
    }
    tagMasterSheet
      .getRange(1, colNumber)
      .setValue(
        wordInfoList[colNumber - TAG_MASTER_COL_NUMBERS["useStatus"]].fileId
      );
    tagMasterSheet
      .getRange(1, colNumber)
      .setNote(
        wordInfoList[colNumber - TAG_MASTER_COL_NUMBERS["useStatus"]]
          .displayFileName
      );
    _renderTagUseFlag(
      wordInfoList[colNumber - TAG_MASTER_COL_NUMBERS["useStatus"]].fileId,
      colNumber
    );
    colNumber++;
  }

  deleteTrigger(triggerKey);
  properties.deleteProperty("colNumber");
}

//指定したkeyに保存されているトリガーIDを使って、トリガーを削除する
function deleteTrigger(triggerKey) {
  var triggerId =
    PropertiesService.getScriptProperties().getProperty(triggerKey);

  if (!triggerId) return;

  ScriptApp.getProjectTriggers()
    .filter(function (trigger) {
      return trigger.getUniqueId() == triggerId;
    })
    .forEach(function (trigger) {
      ScriptApp.deleteTrigger(trigger);
    });
  PropertiesService.getScriptProperties().deleteProperty(triggerKey);
}

//トリガーを発行
function setTrigger(triggerKey, funcName) {
  deleteTrigger(triggerKey); //保存しているトリガーがあったら削除
  var dt = new Date();
  dt.setMinutes(dt.getMinutes() + 1); //１分後に再実行
  var triggerId = ScriptApp.newTrigger(funcName)
    .timeBased()
    .at(dt)
    .create()
    .getUniqueId();
  //あとでトリガーを削除するためにトリガーIDを保存しておく
  PropertiesService.getScriptProperties().setProperty(triggerKey, triggerId);
}

// 使用状況（○、△）を描画
function _renderTagUseFlag(fileId, colNumber) {
  const tagMasterSheet = getSheetByName(SHEET_NAME_TAG_MASTER);
  // 値をクリア
  const lastRow = tagMasterSheet.getLastRow();
  let rowNumber = 2;
  tagMasterSheet
    .getRange(rowNumber, colNumber, lastRow, colNumber)
    .setValue("");
  const wordItemList = _getWordTagInfoListFromFileId(fileId);
  if (wordItemList === null) {
    return;
  }
  while (lastRow >= rowNumber) {
    const itemId = tagMasterSheet
      .getRange(rowNumber, TAG_MASTER_COL_NUMBERS["tagDefinitionId"])
      .getValue();
    if (itemId === "") {
      rowNumber++;
      continue;
    }
    const wordItem = wordItemList.find((wordItem) => {
      return wordItem.tagDefinitionId === itemId;
    });
    if (wordItem) {
      const range = tagMasterSheet.getRange(rowNumber, colNumber);
      const value = wordItem.isRequired ? SANKAKU_STR : MARU_STR;
      range.setValue(value);
    }
    rowNumber++;
  }
}

// 該当のシートを使用有無をチェックする
function _getStatusOfUse(tagId, rangeValues) {
  const startIndex = 1;
  for (const rowValues of rangeValues) {
    const _tagId =
      rowValues[WORD_ITEM_TAG_COL_NUMBERS["tagDefinitionId"] - startIndex];
    if (tagId === _tagId) {
      return rowValues[WORD_ITEM_TAG_COL_NUMBERS["isRequired"] - startIndex] ===
        MARU_STR
        ? IS_OPTION
        : IS_REQUREID;
    }
  }
  return NO_USE;
}

function _getTagMasterRowValuesByItemId(
  tagDefinitionId,
  key,
  rangeValuesOfTagMaster
) {
  if (rangeValuesOfTagMaster == null) {
    rangeValuesOfTagMaster = _getRangeValuesOfTagMaster();
  }
  const startIndex = 1;
  for (const rowValues of rangeValuesOfTagMaster) {
    const _tagDefinitionId =
      rowValues[TAG_MASTER_COL_NUMBERS["tagDefinitionId"] - startIndex];
    const _key = rowValues[TAG_MASTER_COL_NUMBERS["key"] - startIndex];
    if (tagDefinitionId === _tagDefinitionId && key === _key) {
      return rowValues;
    }
  }
  return null;
}

function _getRangeValuesOfTagMaster() {
  const masterSheet = getSheetByName(SHEET_NAME_TAG_MASTER);
  const lastRow = masterSheet.getLastRow();
  let startRow = 2;
  let startCol = 1;
  return masterSheet
    .getRange(startRow, startCol, lastRow, TAG_MASTER_COL_NUMBERS["printType"])
    .getValues();
}

function _getRangeBackgroundsOfTagMaster() {
  const masterSheet = getSheetByName(SHEET_NAME_TAG_MASTER);
  const lastRow = masterSheet.getLastRow();
  let startRow = 2;
  let startCol = 1;
  return masterSheet
    .getRange(startRow, startCol, lastRow, TAG_MASTER_COL_NUMBERS["printType"])
    .getBackgrounds();
}

function _getTagMaster() {
  const rangeValues = _getRangeValuesOfTagMaster();
  const startIndex = 1;
  let tagMaster = [];
  for (const rowValues of rangeValues) {
    const displayName =
      rowValues[TAG_MASTER_COL_NUMBERS["displayName"] - startIndex];
    const tagName = rowValues[TAG_MASTER_COL_NUMBERS["tagName"] - startIndex];
    if (displayName !== "" && tagName !== "") {
      tagMaster.push(new TagMaster(rowValues));
    }
  }
  return tagMaster;
}

function _getTagMasterSQL(backgroundColor) {
  const rangeValues = _getRangeValuesOfTagMaster();
  const rangeBackgrands = _getRangeBackgroundsOfTagMaster();
  const startIndex = 1;
  const filteredArray = [];
  for (let i = 0; i < rangeValues.length; i++) {
    if (rangeBackgrands[i][0] === backgroundColor) {
      filteredArray.push(rangeValues[i]);
    }
  }
  let tagMaster = [];
  for (const rowValues of filteredArray) {
    const displayName =
      rowValues[TAG_MASTER_COL_NUMBERS["displayName"] - startIndex];
    const tagName = rowValues[TAG_MASTER_COL_NUMBERS["tagName"] - startIndex];
    if (displayName !== "" && tagName !== "") {
      tagMaster.push(new TagMaster(rowValues));
    }
  }
  return tagMaster;
}

class TagMaster {
  constructor(rowValues) {
    // colnumberが１始まりのため、rowValuesから値を取得する場合はindexから1を減らす
    const startIndex = 1;
    this.displayName =
      rowValues[TAG_MASTER_COL_NUMBERS["displayName"] - startIndex];
    this.tagDefinitionId =
      rowValues[TAG_MASTER_COL_NUMBERS["tagDefinitionId"] - startIndex];
    this.tagName = rowValues[TAG_MASTER_COL_NUMBERS["tagName"] - startIndex];
    this.category = rowValues[TAG_MASTER_COL_NUMBERS["category"] - startIndex];
    this.key = rowValues[TAG_MASTER_COL_NUMBERS["key"] - startIndex];
    this.printTypeStr =
      rowValues[TAG_MASTER_COL_NUMBERS["printType"] - startIndex];
    this.printType = PRINT_TYPE_MAP[this.printTypeStr];
  }
}
