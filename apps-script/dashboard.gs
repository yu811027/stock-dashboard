function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = getOrCreateSheet_(ss, CONFIG.SHEETS.DATA);
  const analysisSheet = getOrCreateSheet_(ss, CONFIG.SHEETS.ANALYSIS);
  const dashboardSheet = getOrCreateSheet_(ss, CONFIG.SHEETS.DASHBOARD);
  const checkSheet = getOrCreateSheet_(ss, CONFIG.SHEETS.CHECK);

  setupDataSheet_(dataSheet);

  const months = getRecentMonths_(CONFIG.MONTH_COUNT);

  let taiexRows = [];
  let etfRows = [];

  months.forEach(m => {
    taiexRows = taiexRows.concat(fetchTaiexMonth_(m));
    etfRows = etfRows.concat(fetchStockDayMonth_(CONFIG.SYMBOLS.ETF_00631L, m));
    Utilities.sleep(800);
  });

  const etfMap = {};
  etfRows.forEach(r => {
    etfMap[r.date] = r;
  });

  const output = taiexRows.map(t => {
    const e = etfMap[t.date] || {};
    return [
      t.date,
      t.close,
      "",
      e.close || "",
      "",
      "",
      ""
    ];
  });

  output.sort((a, b) => new Date(a[0]) - new Date(b[0]));

  if (output.length > 0) {
    dataSheet.getRange(2, 1, output.length, output[0].length).setValues(output);

    for (let i = 2; i <= output.length + 1; i++) {
      dataSheet.getRange(i, 3).setFormula(`=IFERROR(B${i}/B${i-1}-1,"")`);
      dataSheet.getRange(i, 6).setFormula(`=IFERROR((D${i}-E${i})/E${i},"")`);
    }
  }

  formatDataSheet_(dataSheet);
  setupAnalysisSheet_(analysisSheet, output.length + 1);
  setupDashboardSheet_(dashboardSheet);
  setupCheckSheet_(checkSheet, taiexRows, etfRows, output);

  SpreadsheetApp.getUi().alert("更新完成。");
}

function setupDataSheet_(sheet) {
  sheet.clear();
  const headers = [
    "日期",
    "加權指數收盤",
    "加權指數日報酬",
    "00631L收盤價",
    "00631L淨值",
    "折溢價",
    "備註"
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
}

function formatDataSheet_(sheet) {
  sheet.getRange("A1:G1").setFontWeight("bold").setBackground("#D9EAF7");
  sheet.getRange("A:A").setNumberFormat("yyyy-mm-dd");
  sheet.getRange("B:B").setNumberFormat("#,##0.00");
  sheet.getRange("C:C").setNumberFormat("0.00%");
  sheet.getRange("D:E").setNumberFormat("#,##0.00");
  sheet.getRange("F:F").setNumberFormat("0.00%");
  sheet.autoResizeColumns(1, 7);
}

function setupAnalysisSheet_(sheet, lastRow) {
  sheet.clear();

  const rows = [
    ["指標", "公式/數值"],
    ["加權指數三個月報酬", `=IFERROR(('每日資料'!B${lastRow}/'每日資料'!B2)-1,"")`],
    ["00631L收盤價三個月報酬", `=IFERROR(('每日資料'!D${lastRow}/'每日資料'!D2)-1,"")`],
    ["00631L淨值三個月報酬", `=IFERROR(('每日資料'!E${lastRow}/'每日資料'!E2)-1,"")`],
    ["理論2倍報酬", `=IFERROR(B2*2,"")`],
    ["00631L淨值 vs 理論2倍差異", `=IFERROR(B4-B5,"")`],
    ["平均折溢價", `=IFERROR(AVERAGE('每日資料'!F2:F${lastRow}),"")`],
    ["最大折價", `=IFERROR(MIN('每日資料'!F2:F${lastRow}),"")`],
    ["最大溢價", `=IFERROR(MAX('每日資料'!F2:F${lastRow}),"")`]
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange("A1:B1").setFontWeight("bold").setBackground("#FCE4D6");
  sheet.getRange("B2:B9").setNumberFormat("0.00%");
  sheet.autoResizeColumns(1, 2);
}

function setupDashboardSheet_(sheet) {
  sheet.clear();

  const rows = [
    ["項目", "結果"],
    ["買點評分", "待建立"],
    ["目前是否折價", `=IFERROR(IF(INDEX('每日資料'!F:F,COUNTA('每日資料'!A:A))<0,"折價","溢價/平價"),"")`],
    ["最新加權指數", `=LOOKUP(9^9,'每日資料'!B:B)`],
    ["最新00631L收盤價", `=LOOKUP(9^9,'每日資料'!D:D)`],
    ["最新00631L淨值", `=LOOKUP(9^9,'每日資料'!E:E)`],
    ["操作建議", "淨值補齊後再評估"]
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange("A1:B1").setFontWeight("bold").setBackground("#E2F0D9");
  sheet.autoResizeColumns(1, 2);
}

function setupCheckSheet_(sheet, taiexRows, etfRows, output) {
  sheet.clear();

  const rows = [
    ["項目", "結果"],
    ["加權指數筆數", taiexRows.length],
    ["00631L收盤價筆數", etfRows.length],
    ["合併後筆數", output.length],
    ["最後更新時間", Utilities.formatDate(new Date(), CONFIG.TIMEZONE, "yyyy-MM-dd HH:mm:ss")]
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange("A1:B1").setFontWeight("bold").setBackground("#FFF2CC");
  sheet.autoResizeColumns(1, 2);
}