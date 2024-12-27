function _alert(message) {
  SpreadsheetApp.getUi().alert(message);
}

function _getHeaderRange(index) {
  const from = "A";
  const to = "E";
  return `${from}${index}:${to}${index}`;
}

function _getProjects() {
  const listSheet =
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName("リスト");
  return listSheet
    .getRange("D1:D3")
    .getValues()
    .map((row) => row[0])
    .filter((p) => p);
}

/**
 * 稼働工数管理シートの生成
 */
function generateManagementSheet() {
  // 変数初期化
  const targetYear = 2025;
  const targetMonth = 1;
  const sheetName = `${targetYear}年${targetMonth}月工数管理`;
  const headerBg = "#efefef";
  const sapp = SpreadsheetApp;
  const ss = sapp.getActiveSpreadsheet();

  // シートの存在確認
  const oldSheet = ss.getSheetByName(sheetName);
  if (oldSheet) {
    _alert(
      `シート「${sheetName}」は既に存在します。\n` +
        "新規作成をキャンセルしました。"
    );
    return; // 処理を終了
  }

  // 新規シート作成
  const index = sapp.getActiveSpreadsheet().getNumSheets();
  const sheet = ss.insertSheet(sheetName, index);

  // シートの列幅設定
  const columnWidths = {
    1: 500, // A列
    2: 100, // B列
    3: 100, // C列
    4: 100, // D列
    5: 100, // E列
  };

  for (const [col, width] of Object.entries(columnWidths)) {
    sheet.setColumnWidth(Number(col), width);
  }

  // --- プロジェクト別集計 ---
  // 集計ヘッダー
  let headerRange = _getHeaderRange(1); // A1:E1
  sheet
    .getRange(headerRange)
    .setValues([["プロジェクト別集計", "", "", "", ""]])
    .merge()
    .setBackground(headerBg)
    .setFontWeight("bold");
  // 項目ヘッダー
  headerRange = _getHeaderRange(2); // A2:E2
  sheet
    .getRange(headerRange)
    .setValues([["プロジェクト", "保守工数", "稼働実績", "残工数", ""]])
    .setFontWeight("bold");

  // プロジェクトの取得
  const projects = _getProjects();
  let row = 3;
  projects.forEach((project) => {
    // A3以降にプロジェクト名を設定
    sheet.getRange(`A${row}`).setValue(project);
    row++;
  });
  row++; // 1行空ける

  // プロジェクト毎の表示
  for (const project of projects) {

  }
}
