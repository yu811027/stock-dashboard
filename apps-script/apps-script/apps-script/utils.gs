function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
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

function toNumber_(v) {
  if (v === null || v === undefined || v === "") return "";
  return Number(String(v).replace(/,/g, ""));
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
