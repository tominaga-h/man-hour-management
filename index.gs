function _alert(message) {
  SpreadsheetApp.getUi().alert(message);
}

function _getHeaderRange(index) {
	const from = "A";
	const to = "E";
	return `${from}${index}:${to}${index}`;
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

	// --- プロジェクト別集計 ---
	// 集計ヘッダー
	let headerRange = _getHeaderRange(1); // A1:E1
	sheet.getRange(headerRange)
		.setValues([["プロジェクト別集計", "", "", "", ""]])
		.merge()
		.setBackground(headerBg)
		.setFontWeight("bold")
	;
	// 項目ヘッダー
	headerRange = _getHeaderRange(2); // A2:E2
	sheet.getRange(headerRange)
		.setValues([["プロジェクト", "保守工数", "稼働実績", "残工数", ""]])
		.setFontWeight("bold")
	;
}
