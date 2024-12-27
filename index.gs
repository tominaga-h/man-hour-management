function _alert(message) {
  SpreadsheetApp.getUi().alert(message);
}

/**
 * 稼働工数管理シートの生成
 */
function generateManagementSheet() {
  // 変数初期化
  const targetYear = 2024;
  const targetMonth = 12;
  const sheetName = `${targetYear}年${targetMonth}月工数管理`;
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
}
