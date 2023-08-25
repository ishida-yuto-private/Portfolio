// 定数の定義
// テーブル名
const TABLE_NAME_CONTRACT_EXTRA_INFO = "contract_extra_info";
const TABLE_NAME = "table_name";
const TABLE_NAME_TAG_DEFINITION = "tag_definition";
const TABLE_NAME_WORD_TAG_INFO = "word_tag_info";

// シート名
const SHEET_NAME_WORD_MASTER = "Wordファイル定義";
const SHEET_NAME_CONTRACT_EXTRA_INFO_MASTER = "契約拡張情報項目定義";
const SHEET_NAME_WORD = "Wordタグ定義({fileId})";
const SHEET_NAME_TAG_MASTER = "tag_master";
const SHEET_NAME_WORD_SQL = "Given_sql(hoge_word)";
const SHEET_NAME_CONTRACT_EXTRA_INFO_SQL = "Given_sql(contract_extra_info)";
const SHEET_NAME_WORD_TAG_INFO_SQL = "Given_sql(word_tag_info)";
const SHEET_NAME_TAG_DEFINITION_SQL = "Given_sql(tag_definition)";
const SHEET_NAME_CONST = "Const";

const SHEET_NAME_BACKGRAND_COLORS = new Array(
  SHEET_NAME_WORD,
  SHEET_NAME_WORD_MASTER,
  SHEET_NAME_CONTRACT_EXTRA_INFO_MASTER,
  SHEET_NAME_TAG_MASTER
);

const SQL_SHEET_NAME_LIST = new Array(
  SHEET_NAME_CONTRACT_EXTRA_INFO_SQL,
  SHEET_NAME_WORD_SQL,
  SHEET_NAME_TAG_DEFINITION_SQL,
  SHEET_NAME_WORD_TAG_INFO_SQL
);

// タイトル
const TITLE_CONTRACT_EXTRA_INFO_MASTER = "＜契約拡張情報項目定義＞";
const TITLE_INTO_TAG_MASTER = "<流し込み定義>";

const CREATE_USER = "1";

const CEX_ADD = "CEX";

const USE = 1; // 使用
const NO_USE = 0; //未使用
const IS_OPTION = 1; //オプション使用
const IS_REQUREID = 2; //必須
const MARU_STR = "◯";
const SANKAKU_STR = "△";
//入力タイプ S:文字、R:文字(改行あり)、N:数値、D:日付
//印字タイプ NULL:変換なし、D:日付(年月日)、N:数値、C:数値(カンマ区切り)
const INPUT_TYPE_MAP = {
  文字: "S",
  "文字(改行あり)": "R",
  数値: "N",
  日付: "D",
};
const PRINT_TYPE_MAP = {
  "": "NULL",
  "日付(年月日)": "'D'",
  数値: "'N'",
  "数値(カンマ区切り)": "'C'",
};

const NG_LIST = [`¥`, `/`, `:`, `*`, `?`, `"`, `<`, `>`, `|`, `\\`];

const COLOR_MAP = {
  Yellow: "#ffff00",
  Green: "#00ff00",
  Blue: "#00ffff",
  Pink: "#ff00ff",
  Red: "#ff0000",
};

// WORD_MASTER_COL_NUMBERS = "Wordファイル定義";
const WORD_MASTER_COL_NUMBERS = {
  displayOrder: 1,
  displayFileName: 2,
  phaseId: 3,
  fileId: 6,
  wordFileName: 5,
  isAllArticleFlg: 6,
  sqlkey: 10,
};

// SHEET_NAME_CONTRACT_EXTRA_INFO_MASTER = "契約拡張情報項目定義";
const CONTRACT_EXTRA_INFO_MASTER_COL_NUMBERS = {
  displayOrder: 1,
  name: 2,
  tagName: 3,
  id: 4,
  inputType: 5,
  isDisabled: 6,
};

// SHEET_NAME_TAG_MASTER = "tag_master";
const TAG_MASTER_COL_NUMBERS = {
  displayName: 1,
  tagDefinitionId: 2,
  tagName: 3,
  category: 4,
  key: 5,
  printType: 6,
  empty: 7,
  useStatus: 8,
};

// SHEET_NAME_WORD = "Wordタグ定義({fileId})";
const WORD_ITEM_TAG_COL_NUMBERS = {
  no: 1,
  displayName: 2,
  isRequired: 3,
  tagDefinitionId: 4,
  tagName: 5,
  category: 6,
  key: 7,
  printType: 8,
  sqlkey: 9,
};

const CATEGORY_TYPE_LIST = {
  table_name: 1,
  contract_extra_info: 2,
};

const VBA_SCRIPT_1 =
  "'引数の単語をハイライト	\n" +
  "Function highlight(ByVal word As String)	\n" +
  " Dim range As range	\n" +
  " Set range = ActiveDocument.range(0, 0)	\n" +
  "	\n" +
  " With range.Find	\n" +
  "   .Text = word	\n" +
  "   .Forward = True	\n" +
  "   .Format = False	\n" +
  "   .MatchWholeWord = True     '完全に一致する単語だけを検索する	\n" +
  "   .MatchByte = True          '半角と全角を区別する	\n" +
  "   .MatchCase = True          '大文字と小文字の区別する	\n" +
  "	\n" +
  "   Do While .Execute = True	\n" +
  "     range.HighlightColorIndex = wdYellow	\n" +
  "   Loop	\n" +
  " End With	\n" +
  "End Function	\n" +
  "'置換	\n" +
  "Function replaceWord(ByVal f As String, ByVal t As String)	\n" +
  " Selection.Move wdStory, -1	\n" +
  "	\n" +
  " With Selection.Find	\n" +
  "     .MatchWholeWord = True     '完全に一致する単語だけを検索する	\n" +
  "     .MatchByte = True          '半角と全角を区別する	\n" +
  "     .MatchCase = True          '大文字と小文字の区別する	\n" +
  "     .Text = f	\n" +
  "     .Execute Replace:=wdReplaceAll, replacewith:=t	\n" +
  " End With	\n" +
  "End Function	\n" +
  "	\n" +
  "Sub 置換とマーカー()	\n" +
  " '画面の更新オフ	\n" +
  " Application.ScreenUpdating = False	\n" +
  "	\n" +
  "	\n" +
  " Dim workView As View	\n" +
  " Set workView = Documents(1).ActiveWindow.View	\n" +
  " Dim isRevisionsDisp As Boolean	\n" +
  "	\n" +
  "	\n" +
  " isRevisionsDisp = workView.ShowRevisionsAndComments	\n" +
  " workView.ShowRevisionsAndComments = False	" +
  " \n";

const VBA_SCRIPT_2 =
  " 	\n" +
  "	\n" +
  " '全ての処理終了後、画面の更新オン	\n" +
  " Application.ScreenUpdating = True	\n" +
  "	\n" +
  " workView.ShowRevisionsAndComments = isRevisionsDisp	\n" +
  "	\n" +
  ' MsgBox "チェック完了しました。"	\n' +
  "End Sub	\n" +
  "	\n" +
  "Private Sub Document_New()	\n" +
  "	\n" +
  "End Sub	";
