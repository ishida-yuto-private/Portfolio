/**
 * wordからhoge_wordのUpdate文生成して、描画する
 */
function _renderWordToCreateSql(wordInfoList, dicition) {
  const sqlSheet = getSheetByName(SHEET_NAME_WORD_SQL);
  const gettenantcode = getSheetByName(SHEET_NAME_CONST);
  let tenantcode = gettenantcode.getRange(18, 2).getValue();
  file_path = tenantcode + "_{{id}}.docx";
  const targetRange = sqlSheet.getRange(1, 1);
  sqlSheet.clear();
  if (wordInfoList.length == 0) {
    return;
  }
  sqlValueStrs = [];
  if (dicition == "修正") {
    for (const w of wordInfoList) {
      sqlValueStrs.push(
        "UPDATE table_" +
          `${tenantcode}` +
          ".hoge_word SET `display_order` = " +
          `${w.displayOrder}` +
          ", `display_name` = " +
          `'${w.displayFileName}'` +
          ", `phase_id` = " +
          `${w.phaseId}` +
          ", `file_path` = " +
          `'${w.wordFileName}'` +
          ", `is_all_article_flg` = " +
          `1` +
          " WHERE `id` =" +
          `${w.fileId}`
      );
    }
    let sqlStr = sqlValueStrs.join(";\n");
    sqlStr += ";";
    targetRange.setValue(sqlStr);
  } else {
    for (const w of wordInfoList) {
      sqlValueStrs.push(
        `    ({{display_order}}, '${w.displayFileName}', ${w.phaseId}, {{id}}, '${file_path}', 1, now(), ${CREATE_USER}, NULL, NULL)`
      );
    }
    let sqlStr =
      "INSERT INTO table_" +
      `${tenantcode}` +
      ".hoge_word (`display_order`, `display_name`, `phase_id`, `id`, `file_path`, `is_all_article_flg`, `create_date`, `create_user`, `update_date`, `update_user`) VALUES";
    sqlStr += "\n";
    sqlStr += sqlValueStrs.join(",\n");
    sqlStr += ";";
    targetRange.setValue(sqlStr);
  }
}

/**
 * wordItemListからword_tag_infoのUpdate文生成して、描画する
 */
function _renderWordTagInfoToCreateSql(wordTagInfoList, dicition) {
  const gettenantcode = getSheetByName(SHEET_NAME_CONST);
  let tenantcode = gettenantcode.getRange(18, 2).getValue();
  _clearSqlSheet(SHEET_NAME_word_tag_INFO_SQL);
  if (wordTagInfoList.length == 0) {
    return;
  }

  const wordTagInfoSQLSheet = getSheetByName(SHEET_NAME_word_tag_INFO_SQL);
  let rowNumber = 1;
  for (const wordItem of wordTagInfoList) {
    const targetRange = wordTagInfoSQLSheet.getRange(rowNumber, 1);
    targetRange.setValue("");
    if (wordItem === null || wordItem.length == 0) {
      continue;
    }

    sqlValueStrs = [];
    if (dicition == "修正") {
      for (const item of wordItem) {
        if (item.isAutoCreateTag) {
          continue;
        }
        const isRequired = item.isRequired ? 0 : 1;

        sqlValueStrs.push(
          " UPDATE table_" +
            `${tenantcode}` +
            ".word_tag_info SET `is_required` = " +
            `${isRequired}` +
            " WHERE `hoge_word_id` = " +
            `${item.fileId}` +
            " AND " +
            " `tag_definition_id` = " +
            `${item.tagDefinitionId}`
        );
      }
      let sqlStr = sqlValueStrs.join(";\n");
      sqlStr += ";";
      targetRange.setValue(sqlStr);
      rowNumber++;
    } else {
      for (const item of wordItem) {
        if (item.isAutoCreateTag) {
          continue;
        }
        const isRequired = item.isRequired ? 0 : 1;

        sqlValueStrs.push(
          `    ({{${item.sqlkey}}}, ${item.tagDefinitionId}, ${isRequired}, now(), ${CREATE_USER}, NULL, NULL)`
        );
      }
      let sqlStr =
        "INSERT INTO table_" +
        `${tenantcode}` +
        ".word_tag_info (`hoge_word_id`, `tag_definition_id`, `is_required`, `create_date`, `create_user`, `update_date`, `update_user`) VALUES";
      sqlStr += "\n";
      sqlStr += sqlValueStrs.join(",\n");
      sqlStr += ";";
      targetRange.setValue(sqlStr);
      rowNumber++;
    }
  }
}

/**
 * contractItemExtraInfoからtag_masterのUpdate文生成して、描画する
 */
function _renderTagDefinitionToCreateSql(tagMaster, dicition) {
  const gettenantcode = getSheetByName(SHEET_NAME_CONST);
  let tenantcode = gettenantcode.getRange(18, 2).getValue();
  const sqlSheet = getSheetByName(SHEET_NAME_TAG_DEFINITION_SQL);
  const targetRange = sqlSheet.getRange(1, 1);
  sqlSheet.clear();
  if (tagMaster.length == 0) {
    return;
  }
  let sqltagMasterStrs = [];
  if (dicition == "修正") {
    for (const tagMasterItem of tagMaster) {
      if (CATEGORY_TYPE_LIST[tagMasterItem.category] == 1) {
        tagMasterItem.arrayFlag = USE;
        tagMasterItem.categoryType = 1;
      } else {
        tagMasterItem.arrayFlag = NO_USE;
        tagMasterItem.categoryType = 0;
      }
      sqltagMasterStrs.push(
        "UPDATE table_" +
          `${tenantcode}` +
          ".tag_definition SET `display_name` = " +
          `'${tagMasterItem.displayName}'` +
          ", `tag_name` = " +
          `'${tagMasterItem.tagName}'` +
          " , `category_type` = " +
          `${tagMasterItem.categoryType}` +
          ", `category` = " +
          `'${tagMasterItem.category}'` +
          ", `key` = " +
          `'${tagMasterItem.key}'` +
          " , `print_type` = " +
          `${tagMasterItem.printType}` +
          ", `array_flag` = " +
          `${tagMasterItem.arrayFlag}` +
          " WHERE `id` = " +
          `${tagMasterItem.tagDefinitionId}`
      );
    }
    let sqlStr = sqltagMasterStrs.join(";\n");
    sqlStr += ";";
    targetRange.setValue(sqlStr);
  } else {
    for (const tagMasterItem of tagMaster) {
      if (CATEGORY_TYPE_LIST[tagMasterItem.category] == 1) {
        tagMasterItem.arrayFlag = USE;
        tagMasterItem.categoryType = 1;
      } else {
        tagMasterItem.arrayFlag = NO_USE;
        tagMasterItem.categoryType = 0;
      }
      sqltagMasterStrs.push(
        `    ('${tagMasterItem.displayName}', ${tagMasterItem.tagDefinitionId}, '${tagMasterItem.tagName}', ${tagMasterItem.categoryType}, '${tagMasterItem.category}', '${tagMasterItem.key}', ${tagMasterItem.printType}, ${tagMasterItem.arrayFlag}, now(), ${CREATE_USER}, NULL, NULL)`
      );
    }
    let sqlStr =
      "INSERT INTO table_" +
      `${tenantcode}` +
      ".tag_definition (`display_name`, `id`, `tag_name`, `category_type`, `category`, `key`, `print_type`, `array_flag`, `create_date`, `create_user`, `update_date`, `update_user`) VALUES";
    sqlStr += "\n";
    sqlStr += sqltagMasterStrs.join(",\n");
    sqlStr += ";";
    targetRange.setValue(sqlStr);
  }
}

/**
 * contractItemExtraInfoからcontract_definitionのUpdate文生成して、描画する
 */
function _renderContractItemExtraInfoToCreateSql(
  contractItemExtraInfo,
  dicition
) {
  const gettenantcode = getSheetByName(SHEET_NAME_CONST);
  let tenantcode = gettenantcode.getRange(18, 2).getValue();
  const sqlSheet = getSheetByName(SHEET_NAME_CONTRACT_EXTRA_INFO_SQL);
  const targetRange = sqlSheet.getRange(1, 1);
  sqlSheet.clear();
  if (contractItemExtraInfo.length == 0) {
    return;
  }
  let sqlValueStrs = [];
  if (dicition == "修正") {
    for (const contractItem of contractItemExtraInfo) {
      sqlValueStrs.push(
        "UPDATE table_" +
          tenantcode +
          ".contract_extra_info SET `display_order` =" +
          `${contractItem.displayOrder}` +
          ", `display_name` = " +
          `'${contractItem.name}'` +
          ", `disabled` = " +
          `${contractItem.isDisabled}` +
          ", `input_type` = " +
          `'${contractItem.inputType}'` +
          " WHERE `id` =" +
          `${contractItem.id}`
      );
    }
    let sqlStr = sqlValueStrs.join(";\n");
    sqlStr += ";";
    targetRange.setValue(sqlStr);
  } else {
    for (const contractItem of contractItemExtraInfo) {
      sqlValueStrs.push(
        `    (${contractItem.displayOrder}, '${contractItem.name}', '${contractItem.inputType}', ${contractItem.id}, ${contractItem.isDisabled}, now(), ${CREATE_USER}, NULL, NULL)`
      );
    }
    let sqlStr =
      "INSERT INTO table_" +
      `${tenantcode}` +
      ".contract_extra_info (`display_order`, `display_name`, `input_type`, `id`, `disabled`, `create_date`, `create_user`, `update_date`, `update_user`) VALUES";
    sqlStr += "\n";
    sqlStr += sqlValueStrs.join(",\n");
    sqlStr += ";";
    targetRange.setValue(sqlStr);
  }
}
