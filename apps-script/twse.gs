function fetchTaiexMonth_(yyyymmdd) {
  const url = `https://www.twse.com.tw/rwd/zh/TAIEX/MI_5MINS_HIST?date=${yyyymmdd}&response=json`;
  const json = fetchJson_(url);
  const rows = [];

  if (!json || !json.data) return rows;

  json.data.forEach(r => {
    rows.push({
      date: normalizeTwDate_(r[0]),
      close: toNumber_(r[4])
    });
  });

  return rows.filter(r => r.date && r.close);
}

function getRecentMonths_(n) {
  const result = [];
  const now = new Date();

  for (let i = 0; i < n; i++) {
    const d = new Date(now.getFullYear(), now.getMonth() - i, 1);
    result.push(Utilities.formatDate(d, CONFIG.TIMEZONE, "yyyyMMdd"));
  }

  return result;
}