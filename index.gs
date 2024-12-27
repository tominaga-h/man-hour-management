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

function _getDaysInMonth(year, month) {
  return new Date(year, month, 0).getDate();
}

function _getDateRange(sheet, row, dateStartCol, daysInMonth) {
	return sheet.getRange(
		row, dateStartCol, 1, daysInMonth
	);
}

/**
 * 稼働工数管理シートの生成
 */
function generateManagementSheet() {
  // 変数初期化
  const targetYear = 2025;
  const targetMonth = 1;
  const dateStartCol = 6;
  const daysInMonth = _getDaysInMonth(targetYear, targetMonth);
  const sheetName = `${targetYear}年${targetMonth}月工数管理`;
  const weekdayLabels = ["日", "月", "火", "水", "木", "金", "土"];
  const dateHeaders = [];
  const dayHeaders = [];
  const dateBgs = [];
  const headerBg = "#efefef";
  const issueTotalBg = "#d3e2ed";
  const totalBg = "#dbe9d6";
  const weekendBg = "#df9b99";
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
    4: 110, // D列
    5: 110, // E列
  };

  for (const [col, width] of Object.entries(columnWidths)) {
    sheet.setColumnWidth(Number(col), width);
  }

  // 日付データ準備
  for (let d = 1; d <= daysInMonth; d++) {
    const current = new Date(targetYear, targetMonth - 1, d, 12, 0, 0);
    const dayOfWeek = current.getDay();
    const wLabel = weekdayLabels[dayOfWeek];
    dateHeaders.push(`${d}`);
    dayHeaders.push(wLabel);
    dateBgs.push(dayOfWeek === 0 || dayOfWeek === 6 ? weekendBg : null);
  }

  // --- プロジェクト別集計 ---
  // 集計ヘッダー
  let range = _getHeaderRange(1); // A1:E1
  sheet
    .getRange(range)
    .setValues([["プロジェクト別集計", "", "", "", ""]])
    .merge()
    .setBackground(headerBg)
    .setFontWeight("bold");
  // 項目ヘッダー
  range = _getHeaderRange(2); // A2:E2
  sheet
    .getRange(range)
    .setValues([["プロジェクト", "保守工数", "稼働実績", "残工数", ""]])
    .setHorizontalAlignment("center")
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
    // プロジェクトヘッダー
    range = _getHeaderRange(row);
    sheet
      .getRange(range)
      .setValues([[project, "", "", "", ""]])
      .merge()
      .setBackground(headerBg)
      .setFontWeight("bold");
    row++; // 1行追加

    // 課題ヘッダー
    range = _getHeaderRange(row);
    sheet
      .getRange(range)
      .setValues([
        ["課題", "担当者", "工程", "予定工数(時間)", "稼働実績(時間)"],
      ])
      .setFontWeight("bold")
      .setHorizontalAlignment("center");
    row++;

    // 日付の表示
    const dateHeaderRange = _getDateRange(sheet, row - 2, dateStartCol, daysInMonth);
    const dayHeaderRange = _getDateRange(sheet, row - 1, dateStartCol, daysInMonth);
    dateHeaderRange
			.setValues([dateHeaders])
			.setBackgrounds([dateBgs])
			.setHorizontalAlignment("center");
    dayHeaderRange
			.setValues([dayHeaders])
			.setBackgrounds([dateBgs])
			.setHorizontalAlignment("center");

    for (let d = 1; d <= daysInMonth; d++) {
      sheet.setColumnWidth(dateStartCol-1 + d, 45); // 日付の列幅設定
    }

    // 3回課題のテンプレートを作成
    for (const _ of [1, 2, 3]) {
			_getDateRange(sheet, row, dateStartCol, daysInMonth).setBackgrounds([dateBgs]);
      row++;
			_getDateRange(sheet, row, dateStartCol, daysInMonth).setBackgrounds([dateBgs]);
      row++;
			_getDateRange(sheet, row, dateStartCol, daysInMonth).setBackgrounds([dateBgs]);
      row++; // 3行空ける
			_getDateRange(sheet, row, dateStartCol, daysInMonth).setBackgrounds([dateBgs]);

      // 課題合計セル
      range = _getHeaderRange(row);
      sheet
        .getRange(range)
        .setValues([["課題合計", "", "", "", ""]])
        .setBackground(issueTotalBg);
      row++;
			_getDateRange(sheet, row, dateStartCol, daysInMonth).setBackgrounds([dateBgs]);
    }

    // 総合計セル
    range = _getHeaderRange(row);
    sheet
      .getRange(range)
      .setValues([["総合計", "", "", "", ""]])
      .setFontWeight("bold")
      .setBackground(totalBg);
		_getDateRange(sheet, row, dateStartCol, daysInMonth).setBackgrounds([dateBgs]);
    row++;
    row++; // 1行空ける
  }
}
