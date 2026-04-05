import fetch from "node-fetch";
import fs from "fs";
import xlsx from "xlsx";

// -----------------------------
// 1. Excel から銘柄コードを読み込む
// -----------------------------
const workbook = xlsx.readFile("data_j.xlsx");
const sheet = workbook.Sheets["Sheet1"];
const rows = xlsx.utils.sheet_to_json(sheet);

// デバッグ：Excel の読み込み確認
console.log("Excel rows count:", rows.length);
console.log("First 3 rows:", rows.slice(0, 3));

// コード列を抽出（数字・アルファベット混在OK）
let symbols = rows
  .map(r => String(r["コード"]).trim())
  .filter(code => code && code !== "undefined")
  .map(code => `${code}.T`);

console.log("Extracted symbols (first 20):", symbols.slice(0, 20));

if (symbols.length === 0) {
  console.log("ERROR: Excel から銘柄コードが読み取れませんでした。");
  process.exit(1);
}

// -----------------------------
// 2. Yahoo Finance API（1銘柄ずつ取得）
// -----------------------------
async function fetchSymbol(symbol) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?interval=1d&range=5d`;

  try {
    const res = await fetch(url);
    const json = await res.json();

    // Yahoo が null を返した場合
    if (!json.chart || !json.chart.result) {
      console.log(`No data for ${symbol}:`, json.chart?.error?.description);
      return {
        error: json.chart?.error?.description || "Unknown error from Yahoo Finance"
      };
    }

    const item = json.chart.result[0];
    const timestamps = item.timestamp;
    const q = item.indicators.quote[0];

    if (!timestamps || timestamps.length < 3) {
      console.log(`Not enough data for ${symbol}`);
      return {
        error: "Not enough historical data"
      };
    }

    const last = timestamps.length - 1;
    const now = new Date();
    const hour = now.getHours();

    let todayIndex, prevIndex;

    // -----------------------------
    // 9時判定ロジック
    // -----------------------------
    if (hour < 9) {
      todayIndex = last - 1;
      prevIndex = last - 2;
    } else {
      todayIndex = last;
      prevIndex = last - 1;
    }

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
    console.log(`Error fetching ${symbol}:`, err);
    return {
      error: "Network or fetch error"
    };
  }
}

// -----------------------------
// 3. 全銘柄を順次取得
// -----------------------------
async function main() {
  let finalData = {};

  for (const symbol of symbols) {
    console.log(`Fetching ${symbol}...`);
    const data = await fetchSymbol(symbol);

    // 成功・失敗に関わらず data.json に記録
    finalData[symbol] = data;

    // Yahoo API 負荷対策
    await new Promise(r => setTimeout(r, 500));
  }

  // -----------------------------
  // 4. data.json に保存
  // -----------------------------
  fs.writeFileSync("data.json", JSON.stringify(finalData, null, 2));
  console.log("data.json updated successfully");
}

main();
