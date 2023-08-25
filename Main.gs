// メニューバーへの追加

function onOpen() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet();

  //メニュー配列
  const developerMenu = [
    { name: "タグ項目生成", functionName: "makeTags" },
    { name: "タグ使用状況生成", functionName: "makeTagUseStatus" },
    { name: "ファイル一覧取得", functionName: "getListInFolder" },
    { name: "設計書画像作成用VBA生成", functionName: "makeWordVBA" },
  ];

  sheet.addMenu("開発メニュー", developerMenu); //メニューを追加
  var ui = SpreadsheetApp.getUi();
  var menu = ui.createMenu("対応内容・色選択");
  menu.addSubMenu(
    // サブメニューをメニューに追加する
    ui
      .createMenu("追加") // Uiクラスからメニューを作成する
      .addItem("黄", "addSqlYellow")
      .addItem("緑", "addSqlGreen")
      .addItem("青", "addSqlBlue")
      .addItem("ピンク", "addSqlPink")
      .addItem("赤", "addSqlRed")
  );
  menu.addSubMenu(
    ui
      .createMenu("修正")
      .addItem("黄", "revistionSqlYellow")
      .addItem("緑", "revistionSqlGreen")
      .addItem("青", "revistionSqlBlue")
      .addItem("ピンク", "revistionSqlPink")
      .addItem("赤", "revistionSqlRed")
  );
  menu.addSubMenu(
    ui
      .createMenu("クリア（タグ定義以外）")
      .addItem("黄", "clearYellow")
      .addItem("緑", "clearGreen")
      .addItem("青", "clearBlue")
      .addItem("ピンク", "clearPink")
      .addItem("赤", "clearRed")
      .addItem("VBAシート", "trashSheet")
  );
  menu.addSubMenu(
    ui
      .createMenu("クリア（タグ定義）")
      .addItem("黄", "clearYellowTag")
      .addItem("緑", "clearGreenTag")
      .addItem("青", "clearBlueTag")
      .addItem("ピンク", "clearPinkTag")
      .addItem("赤", "clearRedTag")
  );
  menu.addToUi();
}
