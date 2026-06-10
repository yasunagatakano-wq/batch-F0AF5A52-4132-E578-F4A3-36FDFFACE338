import fetch from "node-fetch";
import fs from "fs";
import xlsx from "xlsx";
import path from "path";

// ---------------------------------------------------------
// 1. Excel から銘柄コードを読み込む（fetch.js と同じ）
// ---------------------------------------------------------
const workbook = xlsx.readFile("data/data_j.xlsx");
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

// ---------------------------------------------------------
// 2. Yahoo Finance API（足種別に取得）
// ---------------------------------------------------------
async function fetchCandles(symbol, interval, range) {
  const url = `https://query1.finance.yahoo.com/v8/finance/chart/${symbol}?interval=${interval}&range=${range}`;

  try {
    const res = await fetch(url);
    const json = await res.json();

    if (!json.chart || !json.chart.result) {
      return { error: json.chart?.error?.description || "Unknown error from Yahoo Finance" };
    }

    const item = json.chart.result[0];
    const timestamps = item.timestamp;
    const q = item.indicators.quote[0];

    if (!timestamps || timestamps.length === 0) {
      return { error: "Not enough historical data" };
    }

    // fetch.js と同じ形式で整形
    let result = {};

    for (let i = 0; i < timestamps.length; i++) {
      const ts = timestamps[i];
      const date = new Date(ts * 1000);

      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, "0");
      const d = String(date.getDate()).padStart(2, "0");
      const key = `${y}${m}${d}`;

      result[key] = {
        o: q.open[i],
        h: q.high[i],
        l: q.low[i],
        c: q.close[i],
        v: q.volume[i]
      };
    }

    return result;

  } catch (err) {
    return { error: "Network or fetch error" };
  }
}

// ---------------------------------------------------------
// 3. 全テクニカル条件の実行（isXxx 系）
// ---------------------------------------------------------
// ※ 実際の関数は別ファイルで定義する想定
//    ここでは import だけ記述（あなたの実装に合わせて調整可能）

import {
  isMaSlopeUpDaily,
  isMaSlopeDownDaily,
  isMaSlopeUpWeekly,
  isMaSlopeDownWeekly,
  isMaSlopeUpMonthly,
  isMaSlopeDownMonthly,
  isPerfectOrderDaily,
  isReversePerfectOrderDaily,
  isPerfectOrderWeekly,
  isReversePerfectOrderWeekly,
  isPerfectOrderMonthly,
  isReversePerfectOrderMonthly,
  isPrePerfectOrder,
  isPreReversePerfectOrder,
  isMaCongestionUp,
  isMaCongestionDown,
  isMaSpreadUp,
  isMaSpreadDown,
  isMa100TrendUp,
  isMa100TrendDown,
  isShimohanshin,
  isGyakushimohanshin,
  is5MaHighUpdate,
  is5MaLowUpdate,
  isSakataTripleTop,
  isSakataTripleBottom,
  isSakataSankuUp,
  isSakataSankuDown,
  isSakataSanpeiUp,
  isSakataSanpeiDown,
  isSakataSanpoUp,
  isSakataSanpoDown,
  isHeadAndShoulders,
  isDoubleBottom,
  isNichiDai,
  isGyakuNichiDai,
  isMonowakareUp,
  isMonowakareDown,
  isMonowakareCrossUp,
  isMonowakareCrossDown,
  isRule9Up,
  isRule9Down,
  isBbZoneBreakDown,
  isCycleUp,
  isCycleDown,
  isSentimentOverheat,
  isStayBox
} from "./heuristics_conditions/index.js";

// ---------------------------------------------------------
// 4. 条件実行まとめ
// ---------------------------------------------------------
function runAllConditions(daily, weekly, monthly) {
  return {
    // 1-1 移動平均線の傾き
    TECH_MA_SLOPE_UP_DAILY: isMaSlopeUpDaily(daily),
    TECH_MA_SLOPE_DOWN_DAILY: isMaSlopeDownDaily(daily),
    TECH_MA_SLOPE_UP_WEEKLY: isMaSlopeUpWeekly(weekly),
    TECH_MA_SLOPE_DOWN_WEEKLY: isMaSlopeDownWeekly(weekly),
    TECH_MA_SLOPE_UP_MONTHLY: isMaSlopeUpMonthly(monthly),
    TECH_MA_SLOPE_DOWN_MONTHLY: isMaSlopeDownMonthly(monthly),

    // 1-2 パーフェクトオーダー
    TECH_MA_PO_DAILY: isPerfectOrderDaily(daily),
    TECH_MA_RPO_DAILY: isReversePerfectOrderDaily(daily),
    TECH_MA_PO_WEEKLY: isPerfectOrderWeekly(weekly),
    TECH_MA_RPO_WEEKLY: isReversePerfectOrderWeekly(weekly),
    TECH_MA_PO_MONTHLY: isPerfectOrderMonthly(monthly),
    TECH_MA_RPO_MONTHLY: isReversePerfectOrderMonthly(monthly),

    // 1-3 前夜
    TECH_MA_PRE_PO: isPrePerfectOrder(daily),
    TECH_MA_PRE_RPO: isPreReversePerfectOrder(daily),

    // 1-4 密集
    TECH_MA_CONGESTION_UP: isMaCongestionUp(daily),
    TECH_MA_CONGESTION_DOWN: isMaCongestionDown(daily),

    // 1-5 間隔
    TECH_MA_SPREAD_UP: isMaSpreadUp(daily),
    TECH_MA_SPREAD_DOWN: isMaSpreadDown(daily),

    // 1-6 100日線
    TECH_MA100_TREND_UP: isMa100TrendUp(daily),
    TECH_MA100_TREND_DOWN: isMa100TrendDown(daily),

    // 2-1 下半身
    TECH_SHIMOHANSHIN: isShimohanshin(daily),
    TECH_GYAKU_SHIMOHANSHIN: isGyakushimohanshin(daily),

    // 2-2 5日線更新
    TECH_5MA_HIGH_UPDATE: is5MaHighUpdate(daily),
    TECH_5MA_LOW_UPDATE: is5MaLowUpdate(daily),

    // 3-1 酒田五法
    TECH_SAKATA_TRIPLE_TOP: isSakataTripleTop(daily),
    TECH_SAKATA_TRIPLE_BOTTOM: isSakataTripleBottom(daily),
    TECH_SAKATA_SANKU_UP: isSakataSankuUp(daily),
    TECH_SAKATA_SANKU_DOWN: isSakataSankuDown(daily),
    TECH_SAKATA_SANPEI_UP: isSakataSanpeiUp(daily),
    TECH_SAKATA_SANPEI_DOWN: isSakataSanpeiDown(daily),
    TECH_SAKATA_SANPO_UP: isSakataSanpoUp(daily),
    TECH_SAKATA_SANPO_DOWN: isSakataSanpoDown(daily),

    // 3-2 三尊
    TECH_HEAD_AND_SHOULDERS: isHeadAndShoulders(daily),

    // 3-3 W底
    TECH_DOUBLE_BOTTOM: isDoubleBottom(daily),

    // 3-4 N大
    TECH_NICHI_DAI: isNichiDai(daily),
    TECH_GYAKU_NICHI_DAI: isGyakuNichiDai(daily),

    // 4-1 ものわかれ
    TECH_MONOWAKARE_UP: isMonowakareUp(daily),
    TECH_MONOWAKARE_DOWN: isMonowakareDown(daily),

    // 4-2 ものわかれ（赤青交差）
    TECH_MONOWAKARE_CROSS_UP: isMonowakareCrossUp(daily, weekly, monthly),
    TECH_MONOWAKARE_CROSS_DOWN: isMonowakareCrossDown(daily, weekly, monthly),

    // 5. 9の法則
    TECH_RULE9_UP: isRule9Up(daily),
    TECH_RULE9_DOWN: isRule9Down(daily),

    // 6. ボリンジャー
    TECH_BB_ZONE_BREAK_DOWN: isBbZoneBreakDown(daily),

    // 7. サイクル
    TECH_CYCLE_UP: isCycleUp(daily),
    TECH_CYCLE_DOWN: isCycleDown(daily),

    // 7-2 話題性
    TECH_SENTIMENT_OVERHEAT: isSentimentOverheat(daily),

    // 8. ステイ
    TECH_STAY_BOX: isStayBox(daily)
  };
}

// ---------------------------------------------------------
// 5. メイン処理
// ---------------------------------------------------------
async function main() {
  let finalData = {};

  for (const symbol of symbols) {
    console.log(`Processing ${symbol} ...`);

    const daily = await fetchCandles(symbol, "1d", "1y");
    const weekly = await fetchCandles(symbol, "1wk", "5y");
    const monthly = await fetchCandles(symbol, "1mo", "10y");

    finalData[symbol] = runAllConditions(daily, weekly, monthly);

    await new Promise(r => setTimeout(r, 500)); // fetch.js と同じ
  }

  // heuristics.json を保存
  fs.writeFileSync("data/heuristics.json", JSON.stringify(finalData, null, 2));

  // ---------------------------------------------------------
  // バックアップ処理（fetch.js と完全同じ方式）
  // ---------------------------------------------------------
  const backupDir = "data/backup";

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

  const backupFile = path.join(backupDir, `heuristics.json.${timestamp}`);
  fs.copyFileSync("data/heuristics.json", backupFile);

  // ---------------------------------------------------------
  // バックアップは 8 個だけ保持
  // ---------------------------------------------------------
  const files = fs
    .readdirSync(backupDir)
    .filter(f => f.startsWith("heuristics.json."))
    .sort();

  while (files.length > 8) {
    const oldFile = files.shift();
    const oldPath = path.join(backupDir, oldFile);
    fs.unlinkSync(oldPath);
  }

  console.log("heuristics.json generation completed.");
}

main();
