/**
 * Yu Cheng 0050正二 Google Sheets 自動整理工具
 * 功能：
 * 1. 建立三張表：每日資料、績效分析、操作儀表板
 * 2. 抓近三個月台股加權指數收盤
 * 3. 抓 00631L 元大台灣50正2 每日收盤價
 * 4. 預留 00631L 淨值欄位，可貼上元大/TWSE淨值資料後自動算折溢價
 *
 * 使用方式：
 * Google Sheets → 擴充功能 → Apps Script → 貼上本程式 → 儲存
 * 回到 Sheets 重新整理，就會看到「0050正二工具」選單
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("0050正二工具")
    .addItem("建立/更新近三個月資料", "updateDashboard")
    .addToUi();
}

function updateDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSheet = getOrCreateSheet_(ss, "每日資料");
  const analysisSheet = getOrCreateSheet_(ss, "績效分析");
  const dashboardSheet = getOrCreateSheet_(ss, "操作儀表板");

  setupDataSheet_(dataSheet);
  const months = getRecentMonths_(3);

  let taiexRows = [];
  let etfRows = [];

  months.forEach(m => {
    taiexRows = taiexRows.concat(fetchTaiexMonth_(m));
    etfRows = etfRows.concat(fetchStockDayMonth_("00631L", m));
    Utilities.sleep(800);
  });

  const etfMap = {};
  etfRows.forEach(r => {
    etfMap[r.date] = r;
  });

  const output = [];
  taiexRows.forEach(t => {
    const e = etfMap[t.date] || {};
    output.push([
      t.date,
      t.close,
      "",
      e.close || "",
      "",          // 00631L淨值：請從元大/TWSE淨值資料貼入
      "",          // 折溢價公式
      ""           // 備註
    ]);
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
  setupCheckSheet_(ss, taiexRows, etfRows, output);

  SpreadsheetApp.getUi().alert("已完成更新。提醒：00631L 淨值欄位需要從元大投信/TWSE淨值資料貼入，貼入後折溢價會自動計算。");
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

function fetchTaiexMonth_(yyyymmdd) {
  const url = `https://www.twse.com.tw/rwd/zh/TAIEX/MI_5MINS_HIST?date=${yyyymmdd}&response=json`;
  const json = fetchJson_(url);
  const rows = [];
  if (!json || !json.data) return rows;

  json.data.forEach(r => {
    // 常見欄位：日期、開盤指數、最高指數、最低指數、收盤指數
    rows.push({
      date: normalizeTwDate_(r[0]),
      close: toNumber_(r[4])
    });
  });
  return rows.filter(r => r.date && r.close);
}

function fetchStockDayMonth_(stockNo, yyyymmdd) {
  const url = `https://www.twse.com.tw/rwd/zh/afterTrading/STOCK_DAY?date=${yyyymmdd}&stockNo=${stockNo}&response=json`;
  const json = fetchJson_(url);
  const rows = [];
  if (!json || !json.data) return rows;

  json.data.forEach(r => {
    // 常見欄位：日期、成交股數、成交金額、開盤價、最高價、最低價、收盤價...
    rows.push({
      date: normalizeTwDate_(r[0]),
      close: toNumber_(r[6])
    });
  });
  return rows.filter(r => r.date && r.close);
}

function fetchJson_(url) {
  try {
    const res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        "User-Agent": "Mozilla/5.0"
      }
    });
    if (res.getResponseCode() !== 200) return null;
    return JSON.parse(res.getContentText());
  } catch (e) {
    return null;
  }
}

function getRecentMonths_(n) {
  const result = [];
  const now = new Date();
  for (let i = 0; i < n; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    result.push(Utilities.formatDate(d, "Asia/Taipei", "yyyyMMdd"));
  }
  return result;
}

function normalizeTwDate_(s) {
  if (!s) return "";
  const parts = String(s).split("/");
  if (parts.length !== 3) return s;
  const year = Number(parts[0]) + 1911;
  const month = parts[1].padStart(2, "0");
  const day = parts[2].padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function toNumber_(v) {
  if (v === null || v === undefined || v === "") return "";
  return Number(String(v).replace(/,/g, ""));
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

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}
function setupCheckSheet_(ss, taiexRows, etfRows, output) {
  const sheet = getOrCreateSheet_(ss, "資料檢查");
  sheet.clear();

  const rows = [
    ["項目", "結果"],
    ["加權指數筆數", taiexRows.length],
    ["00631L收盤價筆數", etfRows.length],
    ["合併後筆數", output.length],
    ["最後更新時間", Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss")]
  ];

  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange("A1:B1").setFontWeight("bold").setBackground("#FFF2CC");
  sheet.autoResizeColumns(1, 2);
}
  

