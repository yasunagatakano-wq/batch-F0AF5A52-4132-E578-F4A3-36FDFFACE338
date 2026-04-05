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

// コード列を抽出し、.T を付ける
let symbols = rows
  .map(r => String(r["コード"]).trim())
  .filter(code => code && code !== "undefined")
  .map(code => `${code}.T`);

// デバッグ：抽出された銘柄コード
console.log("Extracted symbols (first 20):", symbols.slice(0, 20));

if (symbols.length === 0) {
  console.log("ERROR: Excel から銘柄コードが読み取れませんでした。");
  process.exit(1);
}

// -----------------------------
// 2. Yahoo Finance v8 API を叩く
// -----------------------------
const chunkSize = 100;

function chunk(array, size) {
  const result = [];
  for (let i = 0; i < array.length; i += size) {
    result.push(array.slice(i, i + size));
  }
  return result;
}

const symbolChunks = chunk(symbols, chunkSize);

// -----------------------------
// 3. v8 API からローソク足を取得
// -----------------------------
async function fetchChunk(chunk) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${chunk.join(",")}?interval=1d&range=5d`;

  try {
    const res = await fetch(url);
    const json = await res.json();

    // デバッグ：API の生レスポンス
    console.log("API response sample:", JSON.stringify(json.chart, null, 2).slice(0, 500));

    if (!json.chart || !json.chart.result) {
      console.log("Invalid response:", json);
      return {};
    }

    const results = json.chart.result;
    const output = {};

    results.forEach(item => {
      if (!item || !item.meta || !item.timestamp) return;

      const symbol = item.meta.symbol;
      const timestamps = item.timestamp;
      const o = item.indicators.quote[0].open;
      const h = item.indicators.quote[0].high;
      const l = item.indicators.quote[0].low;
      const c = item.indicators.quote[0].close;
      const v = item.indicators.quote[0].volume;

      if (!timestamps || timestamps.length < 3) {
        console.log(`Not enough data for ${symbol}`);
        return;
      }

      const last = timestamps.length - 1; // 実行日
      const now = new Date();
      const hour = now.getHours();

      let todayIndex, prevIndex;

      // -----------------------------
      // 9時判定ロジック
      // -----------------------------
      if (hour < 9) {
        todayIndex = last - 1; // 前日
        prevIndex = last - 2;  // 前々日
      } else {
        todayIndex = last;     // 実行日
        prevIndex = last - 1;  // 前日
      }

      output[symbol] = {
        prev: {
          o: o[prevIndex],
          h: h[prevIndex],
          l: l[prevIndex],
          c: c[prevIndex],
          v: v[prevIndex]
        },
        today: {
          o: o[todayIndex],
          h: h[todayIndex],
          l: l[todayIndex],
          c: c[todayIndex],
          v: v[todayIndex]
        }
      };
    });

    return output;

  } catch (err) {
    console.error("Fetch error:", err);
    return {};
  }
}

// -----------------------------
// 4. 全チャンクを順次取得
// -----------------------------
async function main() {
  let finalData = {};

  for (let i = 0; i < symbolChunks.length; i++) {
    console.log(`Fetching chunk ${i + 1}/${symbolChunks.length}...`);
    const chunkData = await fetchChunk(symbolChunks[i]);
    finalData = { ...finalData, ...chunkData };

    await new Promise(r => setTimeout(r, 1500));
  }

  // -----------------------------
  // 5. data.json に保存
  // -----------------------------
  fs.writeFileSync("data.json", JSON.stringify(finalData, null, 2));
  console.log("data.json updated successfully");
}

main();
