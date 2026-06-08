import fetch from "node-fetch";
import fs from "fs";
import path from "path";
import xlsx from "xlsx";
import { JSDOM } from "jsdom";
import iconv from "iconv-lite";
import { getDocument } from "pdfjs-dist/legacy/build/pdf.mjs";

// ======================================================================
// Utility
// ======================================================================
function ensureDir(dir) {
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
}

function jstNow() {
  return new Date(Date.now() + 9 * 60 * 60 * 1000);
}

function timestamp() {
  const now = jstNow();
  const pad = n => String(n).padStart(2, "0");
  return (
    now.getFullYear().toString() +
    pad(now.getMonth() + 1) +
    pad(now.getDate()) + "_" +
    pad(now.getHours()) +
    pad(now.getMinutes()) +
    pad(now.getSeconds())
  );
}

function normalizeNFKC(s) {
  return s.normalize("NFKC");
}

// ======================================================================
// 1. JSF meigara.csv（Shift_JIS）
// ======================================================================
async function fetchKubunMap() {
  console.log("\n==============================");
  console.log("STEP 1: Fetch JSF meigara.csv");
  console.log("==============================");

  const url = "https://www.taisyaku.jp/data/meigara.csv";
  const res = await fetch(url);
  const buf = Buffer.from(await res.arrayBuffer());
  const csv = iconv.decode(buf, "shift_jis");

  const lines = csv.split(/\r?\n/);
  console.log("meigara.csv 行数 =", lines.length);
  console.log("1行目 =", lines[0]);
  console.log("2行目 =", lines[1]);

  const headerLine = lines[1];
  const headers = headerLine.split(",");
  console.log("ヘッダ =", headers);

  const codeIndex = headers.indexOf("コード");
  const kubunIndex = headers.indexOf("貸借銘柄区分（東証）");

  console.log("コード index =", codeIndex, "貸借区分 index =", kubunIndex);

  const kubunMap = {};

  for (let i = 2; i < lines.length; i++) {
    const line = lines[i];
    if (!line.trim()) continue;

    const cols = line.split(",");
    const rawCode = cols[codeIndex];
    const kubun = cols[kubunIndex];

    if (!rawCode) continue;

    const code = normalizeNFKC(String(rawCode)).padStart(4, "0");
    kubunMap[code] = String(kubun);
  }

  console.log("kubunMap 件数 =", Object.keys(kubunMap).length);
  console.log("kubunMap サンプル =", Object.entries(kubunMap).slice(0, 10));

  return kubunMap;
}

// ======================================================================
// 2. 楽天規制（rowspan 対応）
// ======================================================================
async function fetchRakutenRegulation() {
  console.log("\n==============================");
  console.log("STEP 2: Fetch Rakuten Regulation");
  console.log("==============================");

  const url = "https://www.rakuten-sec.co.jp/ITS/Companyfile/margin_restriction.html";
  const res = await fetch(url);
  const html = await res.text();

  const dom = new JSDOM(html);
  const document = dom.window.document;

  const BUY_BAN_KEYWORDS = ["新規買停止", "全取引停止"];
  const SELL_BAN_KEYWORDS = ["新規売停止", "全取引停止"];
  const TOKYO_KEYWORDS = ["東京"];

  const rows = [...document.querySelectorAll("table tr")];
  console.log("楽天 table 行数 =", rows.length);

  const regulationMap = {};
  const current = Array(6).fill(null);
  const rowspanLeft = Array(6).fill(0);

  for (const tr of rows) {
    const tds = [...tr.querySelectorAll("td")];
    if (tds.length === 0) continue;

    const logical = Array(6).fill(null);
    let td_i = 0;

    for (let col = 0; col < 6; col++) {
      if (rowspanLeft[col] > 0) {
        logical[col] = current[col];
        rowspanLeft[col]--;
      }
    }

    for (let col = 0; col < 6; col++) {
      if (logical[col] !== null) continue;
      if (td_i >= tds.length) break;

      const td = tds[td_i];
      const text = td.textContent.trim();
      logical[col] = text;

      const rs = td.getAttribute("rowspan");
      if (rs) {
        rowspanLeft[col] = parseInt(rs) - 1;
        current[col] = text;
      }

      td_i++;
    }

    const rawCode = logical[0];
    const market = logical[2];
    const text = logical[3];

    if (!rawCode) continue;

    const mCode = String(rawCode).match(/(\d{4})/);
    if (!mCode) continue;
    const code4 = mCode[1];

    let marketFlag = false;
    if (market) {
      const m = normalizeNFKC(String(market)).replace(/\s+/g, "");
      marketFlag = TOKYO_KEYWORDS.some(k => m.includes(k));
    }
    if (!marketFlag) continue;

    regulationMap[code4] = regulationMap[code4] || [];
    regulationMap[code4].push(text);
  }

  console.log("regulationMap 件数 =", Object.keys(regulationMap).length);
  console.log("regulationMap サンプル =", Object.entries(regulationMap).slice(0, 10));

  return { regulationMap, BUY_BAN_KEYWORDS, SELL_BAN_KEYWORDS };
}

// ======================================================================
// 3. JPX 週次 PDF
// ======================================================================
async function fetchJpxWeekly() {
  console.log("\n==============================");
  console.log("STEP 3: Fetch JPX Weekly PDF");
  console.log("==============================");

  const page = "https://www.jpx.co.jp/markets/statistics-equities/margin/05.html";
  const res = await fetch(page);
  const html = await res.text();

  const dom = new JSDOM(html);
  const document = dom.window.document;

  const pdfLinks = [...document.querySelectorAll("a")]
    .map(a => a.href)
    .filter(h => h.endsWith(".pdf") && h.includes("syumatsu"))
    .map(h => "https://www.jpx.co.jp" + h);

  console.log("PDFリンク数 =", pdfLinks.length);

  if (pdfLinks.length === 0) return {};

  const latest = pdfLinks.sort().slice(-1)[0];
  console.log("最新PDF =", latest);

  const pdfRes = await fetch(latest);
  const pdfBuf = await pdfRes.arrayBuffer();

  const pdfDoc = await getDocument({ data: pdfBuf }).promise;

  let fullText = "";
  for (let i = 1; i <= pdfDoc.numPages; i++) {
    const page = await pdfDoc.getPage(i);
    const content = await page.getTextContent();
    const strings = content.items.map(it => it.str).join(" ");
    fullText += "\n" + strings;
  }

  const blocks = fullText.split(/(?=[0-9A-Z]{4}0\s+JP\d{10})/);
  console.log("PDF ブロック数 =", blocks.length);

  const jpxMap = {};

  function parseNum(s) {
    s = s.replace(/,/g, "").replace(/\s+/g, "");
    if (s === "" || s === "▲") return 0;
    if (s.startsWith("▲")) return -parseInt(s.slice(1));
    return parseInt(s);
  }

  for (const block of blocks) {
    const m = block.match(/([0-9A-Z]{4}0)\s+JP\d{10}/);
    if (!m) continue;

    const rawCode5 = m[1];
    const code4 = normalizeNFKC(rawCode5.slice(0, 4));

    const afterIsin = block.split(/JP\d{10}/)[1] || "";
    const nums = (afterIsin.match(/[▲\-]?\s*[\d,]+/g) || []).map(parseNum);

    if (nums.length < 4) continue;

    const [sell, sellDiff, buy, buyDiff] = nums;
    const ratio = sell !== 0 ? Math.round((buy / sell) * 100) / 100 : null;

    jpxMap[code4] = { buy, buy_diff: buyDiff, sell, sell_diff: sellDiff, ratio };
  }

  console.log("jpxMap 件数 =", Object.keys(jpxMap).length);
  console.log("jpxMap サンプル =", Object.entries(jpxMap).slice(0, 10));

  return jpxMap;
}

// ======================================================================
// 4. JPX 日々公表 XLS
// ======================================================================
async function fetchJpxDaily() {
  console.log("\n==============================");
  console.log("STEP 4: Fetch JPX Daily XLS");
  console.log("==============================");

  const index = "https://www.jpx.co.jp/markets/statistics-equities/margin/index.html";
  const res = await fetch(index);
  const html = await res.text();

  const dom = new JSDOM(html);
  const document = dom.window.document;

  const links = [...document.querySelectorAll("a")]
    .map(a => a.href)
    .filter(h => /mtdailyk.*\.xls$/.test(h));

  console.log("日々公表 XLS リンク数 =", links.length);

  if (links.length === 0) return {};

  const latest = links.sort().slice(-1)[0];
  console.log("最新日々公表 XLS =", latest);

  const url = "https://www.jpx.co.jp" + latest;
  const buf = await (await fetch(url)).arrayBuffer();
  const wb = xlsx.read(buf, { type: "buffer" });
  const sheet = wb.Sheets[wb.SheetNames[0]];

  const rows = xlsx.utils.sheet_to_json(sheet, { header: 0, range: 5 });
  console.log("日々公表 行数 =", rows.length);

  const codeCol = "コード";
  const sellCol = "売残高 Outstanding Sales";
  const buyCol = "買残高 Outstanding Purchases";

  const dailyMap = {};

  for (const row of rows) {
    const raw = String(row[codeCol] || "").trim();
    if (!/^[0-9A-Z]{4}0$/.test(raw)) continue;

    const code4 = normalizeNFKC(raw.slice(0, 4));

    const sell = parseInt(String(row[sellCol] || "0").replace(/,/g, ""));
    const buy = parseInt(String(row[buyCol] || "0").replace(/,/g, ""));

    if (Number.isNaN(sell) || Number.isNaN(buy)) continue;

    dailyMap[code4] = { sell, buy };
  }

  console.log("dailyMap 件数 =", Object.keys(dailyMap).length);
  console.log("dailyMap サンプル =", Object.entries(dailyMap).slice(0, 10));

  return dailyMap;
}

// ======================================================================
// 5. 日々公表 → 週次 上書き
// ======================================================================
function applyDailyToWeekly(jpxMap, dailyMap) {
  console.log("\n==============================");
  console.log("STEP 5: Apply Daily → Weekly");
  console.log("==============================");

  for (const [code4, d] of Object.entries(dailyMap)) {
    const newBuy = d.buy;
    const newSell = d.sell;

    if (jpxMap[code4]) {
      const prevBuy = jpxMap[code4].buy ?? 0;
      const prevSell = jpxMap[code4].sell ?? 0;

      jpxMap[code4].buy = newBuy;
      jpxMap[code4].sell = newSell;
      jpxMap[code4].buy_diff = newBuy - prevBuy;
      jpxMap[code4].sell_diff = newSell - prevSell;
      jpxMap[code4].ratio = newSell !== 0 ? Math.round((newBuy / newSell) * 100) / 100 : null;
    } else {
      jpxMap[code4] = {
        buy: newBuy,
        buy_diff: null,
        sell: newSell,
        sell_diff: null,
        ratio: newSell !== 0 ? Math.round((newBuy / newSell) * 100) / 100 : null
      };
    }
  }

  console.log("上書き後 jpxMap 件数 =", Object.keys(jpxMap).length);
  console.log("上書き後 jpxMap サンプル =", Object.entries(jpxMap).slice(0, 10));
}

// ======================================================================
// 6. margin.json 統合
// ======================================================================
function buildMarginJson(kubunMap, regulationMap, BUY_BAN, SELL_BAN, jpxMap) {
  console.log("\n==============================");
  console.log("STEP 6: Build margin.json");
  console.log("==============================");

  const allCodes = [
    ...Object.keys(regulationMap),
    ...Object.keys(kubunMap),
    ...Object.keys(jpxMap)
  ];
  const uniqCodes = [...new Set(allCodes)];

  console.log("統合対象コード数 =", uniqCodes.length);

  const margin = {};

  for (const code of uniqCodes) {
    const kubun = kubunMap[code];

    if (kubun === undefined || kubun === null || kubun === "0") continue;

    const regs = regulationMap[code] || [];
    const jpx = jpxMap[code] || {};

    const hasSellBan = regs.some(r => SELL_BAN.some(k => r.includes(k)));
    const hasBuyBan = regs.some(r => BUY_BAN.some(k => r.includes(k)));

    const seiBuy = !hasBuyBan;
    const seiSell = kubun === "1" && !hasSellBan;

    margin[code] = {
      "貸借区分": kubun,
      "制度信用": {
        "買い建て": seiBuy,
        "売り建て": seiSell
      },
      "JPX信用買残": jpx.buy,
      "JPX信用買残前週比": jpx.buy_diff,
      "JPX信用売残": jpx.sell,
      "JPX信用売残前週比": jpx.sell_diff,
      "JPX信用倍率": jpx.ratio,
      "規制": regs
    };
  }

  console.log("margin.json 件数 =", Object.keys(margin).length);
  console.log("margin.json サンプル =", Object.entries(margin).slice(0, 10));

  return margin;
}

// ======================================================================
// 7. Main
// ======================================================================
async function main() {
  ensureDir("data");

  const kubunMap = await fetchKubunMap();
  const { regulationMap, BUY_BAN_KEYWORDS, SELL_BAN_KEYWORDS } =
    await fetchRakutenRegulation();
  const jpxMap = await fetchJpxWeekly();
  const dailyMap = await fetchJpxDaily();

  applyDailyToWeekly(jpxMap, dailyMap);

  const margin = buildMarginJson(
    kubunMap,
    regulationMap,
    BUY_BAN_KEYWORDS,
    SELL_BAN_KEYWORDS,
    jpxMap
  );

  fs.writeFileSync("data/margin.json", JSON.stringify(margin, null, 2), "utf-8");

  console.log("\n✔ margin.json 更新完了");
}

main();
