import fetch from "node-fetch";
import fs from "fs";
import xlsx from "xlsx";
import path from "path";

// -----------------------------
// 1. Excel から銘柄コードを読み込む
// -----------------------------
const workbook = xlsx.readFile("data_j.xlsx");
const sheet = workbook.Sheets["Sheet1"];
const rows = xlsx.utils.sheet_to_json(sheet);

let symbols = rows
  .map(r => String(r["コード"]).trim())
  .filter(code => code && code !== "undefined")
  .map(code => `${code}.T`);

if (symbols.length === 0) {
  console.log("ERROR: Excel から銘柄コードが読み取れませんでした。");
  process.exit(1);
}

// ★ 動作確認用：1件だけに制限
symbols = symbols.slice(0, 1);

// -----------------------------
// 2. Yahoo Finance API（1銘柄ずつ取得）
// -----------------------------
async function fetchSymbol(symbol) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?interval=1d&range=5d`;

  try {
    const res = await fetch(url);
    const json = await res.json();

    if (!json.chart || !json.chart.result) {
      return {
        error: json.chart?.error?.description || "Unknown error from Yahoo Finance"
      };
    }

    const item = json.chart.result[0];
    const timestamps = item.timestamp;
    const q = item.indicators.quote[0];

    if (!timestamps || timestamps.length < 2) {
      return { error: "Not enough historical data" };
    }

    const last = timestamps.length - 1;
    const todayIndex = last;
    const prevIndex = last - 1;

    return {
      prev: {
        o: q.open[prevIndex],
        h: q.high[prevIndex],
        l: q.low[prevIndex],
        c: q.close[prevIndex],
        v: q.volume[prevIndex]
      },
      today: {
        o: q.open[todayIndex],
        h: q.high[todayIndex],
        l: q.low[todayIndex],
        c: q.close[todayIndex],
        v: q.volume[todayIndex]
      }
    };

  } catch (err) {
    return { error: "Network or fetch error" };
  }
}

// -----------------------------
// 3. 全銘柄を順次取得
// -----------------------------
async function main() {
  let finalData = {};

  for (const symbol of symbols) {
    const data = await fetchSymbol(symbol);
    finalData[symbol] = data;
    await new Promise(r => setTimeout(r, 500));
  }

  // -----------------------------
  // 4. data.json を洗い替え（相対パス）
  // -----------------------------
  fs.writeFileSync("data.json", JSON.stringify(finalData, null, 2));

  // -----------------------------
  // 5. バックアップ処理（相対パス & JST）
  // -----------------------------
  const backupDir = "backup";

  if (!fs.existsSync(backupDir)) {
    fs.mkdirSync(backupDir, { recursive: true });
  }

  const now = new Date(Date.now() + 9 * 60 * 60 * 1000);
  const pad = n => String(n).padStart(2, "0");

  const timestamp =
    now.getFullYear().toString() +
    pad(now.getMonth() + 1) +
    pad(now.getDate()) + "_" +
    pad(now.getHours()) +
    pad(now.getMinutes()) +
    pad(now.getSeconds());

  const backupFile = path.join(backupDir, `data.json.${timestamp}`);
  fs.copyFileSync("data.json", backupFile);

  // -----------------------------
  // 6. バックアップは 3 個だけ保持
  // -----------------------------
  const files = fs
    .readdirSync(backupDir)
    .filter(f => f.startsWith("data.json."))
    .sort();

  while (files.length > 3) {
    const oldFile = files.shift();
    const oldPath = path.join(backupDir, oldFile);
    fs.unlinkSync(oldPath);
  }
}

main();
