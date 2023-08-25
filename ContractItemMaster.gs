/**
 * 契約拡張情報項目定義シート内のテキストを2次元配列にて取得
 */
function _getRangeValuesOfContractItemExtraInfo() {
  const masterSheet = getSheetByName(SHEET_NAME_CONTRACT_EXTRA_INFO_MASTER);
  const lastRow = masterSheet.getLastRow();
  let startRow = 2;
  let startCol = 1;
  return masterSheet
    .getRange(
      startRow,
      startCol,
      lastRow,
      CONTRACT_EXTRA_INFO_MASTER_COL_NUMBERS["isDisabled"]
    )
    .getValues();
}

/**
 * 契約拡張情報項目定義シート内のバックエンドカラーを2次元配列にて取得
 */
function _getRangeBackgroundsOfContractItemExtraInfo() {
  const masterSheet = getSheetByName(SHEET_NAME_CONTRACT_EXTRA_INFO_MASTER);
  const lastRow = masterSheet.getLastRow();
  let startRow = 2;
  let startCol = 1;
  return masterSheet
    .getRange(
      startRow,
      startCol,
      lastRow,
      CONTRACT_EXTRA_INFO_MASTER_COL_NUMBERS["isDisabled"]
    )
    .getBackgrounds();
}

/**
 * 契約拡張情報項目定義シート内のA列に値があるもの(displayOrder)のみにする
 */
function _getContractItemExtraInfo() {
  const rangeValues = _getRangeValuesOfContractItemExtraInfo();
  const startIndex = 1;
  let contractItemExtraInfo = [];
  for (const rowValues of rangeValues) {
    const displayOrder =
      rowValues[
        CONTRACT_EXTRA_INFO_MASTER_COL_NUMBERS["displayOrder"] - startIndex
      ];
    if (displayOrder !== "") {
      contractItemExtraInfo.push(new ContractItemExtraInfo(rowValues));
    }
  }
  return _sortContractItem(contractItemExtraInfo, "displayOrder");
}

function _getContractSQLItemExtraInfo(backgroundColor) {
  const rangeValues = _getRangeValuesOfContractItemExtraInfo();
  const rangeBackgrounds = _getRangeBackgroundsOfContractItemExtraInfo();
  const startIndex = 1;
  const filteredArray = [];

  for (let i = 0; i < rangeValues.length; i++) {
    for (let j = 0; j < rangeBackgrounds[i].length; j++) {
      if (rangeBackgrounds[i][j] === backgroundColor) {
        filteredArray.push(rangeValues[i]);
        continue;
      }
    }
  }

  let contractItemExtraInfo = [];
  for (const rowValues of filteredArray) {
    const displayOrder =
      rowValues[
        CONTRACT_EXTRA_INFO_MASTER_COL_NUMBERS["displayOrder"] - startIndex
      ];
    if (displayOrder !== "") {
      contractItemExtraInfo.push(new ContractItemExtraInfo(rowValues));
    }
  }
  return _sortContractItem(contractItemExtraInfo, "displayOrder");
}

/**
 * アイテムをソートする
 */
function _sortContractItem(contractItemExtraInfo, orderColumn, isAsc = true) {
  contractItemExtraInfo.sort((a, b) => {
    const aValue = a[orderColumn];
    const bValue = b[orderColumn];
    if (aValue < bValue) {
      return -1 * (isAsc ? 1 : -1);
    }
    if (aValue > bValue) {
      return 1 * (isAsc ? 1 : -1);
    }
    return 0;
  });
  return contractItemExtraInfo;
}

class ContractItemExtraInfo {
  constructor(rowValues) {
    // colnumberが１始まりのため、rowValuesから値を取得する場合はindexから1を減らす
    const startIndex = 1;
    this.displayOrder =
      rowValues[
        CONTRACT_EXTRA_INFO_MASTER_COL_NUMBERS["displayOrder"] - startIndex
      ];
    this.name =
      rowValues[CONTRACT_EXTRA_INFO_MASTER_COL_NUMBERS["name"] - startIndex];
    this.tagName =
      rowValues[CONTRACT_EXTRA_INFO_MASTER_COL_NUMBERS["tagName"] - startIndex];
    this.id =
      rowValues[CONTRACT_EXTRA_INFO_MASTER_COL_NUMBERS["id"] - startIndex];
    this.inputTypeStr =
      rowValues[
        CONTRACT_EXTRA_INFO_MASTER_COL_NUMBERS["inputType"] - startIndex
      ];
    this.inputType = INPUT_TYPE_MAP[this.inputTypeStr];
    this.isDisabled =
      rowValues[
        CONTRACT_EXTRA_INFO_MASTER_COL_NUMBERS["isDisabled"] - startIndex
      ];
  }
}
