import fetch from "node-fetch";
import fs from "fs";
import xlsx from "xlsx";

// -----------------------------
// 1. Excel から銘柄コードを読み込む
// -----------------------------
const workbook = xlsx.readFile("data_j.xlsx");
const sheet = workbook.Sheets["Sheet1"];
const rows = xlsx.utils.sheet_to_json(sheet);

// コード列を抽出し、.T を付ける
let symbols = rows
  .map(r => String(r["コード"]).trim())
  .filter(code => code && code !== "undefined")
  .map(code => `${code}.T`);

console.log(`Loaded ${symbols.length} symbols from Excel.`);

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

    if (!json.chart || !json.chart.result) {
      console.log("Invalid response:", json);
      return {};
    }

    const results = json.chart.result;
    const output = {};

    results.forEach(item => {
      const symbol = item.meta.symbol;
      const timestamps = item.timestamp;
      const o = item.indicators.quote[0].open;
      const h = item.indicators.quote[0].high;
      const l = item.indicators.quote[0].low;
      const c = item.indicators.quote[0].close;
      const v = item.indicators.quote[0].volume;

      if (!timestamps || timestamps.length < 3) return;

      const last = timestamps.length - 1; // 実行日
      const now = new Date();
      const hour = now.getHours();

      let todayIndex, prevIndex;

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

  fs.writeFileSync("data.json", JSON.stringify(finalData, null, 2));
  console.log("data.json updated successfully");
}

main();