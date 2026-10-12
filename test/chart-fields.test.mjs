/**
 * 工具真的收得下它自己宣稱支援的圖表。
 *
 * ── 這條測試是為了什麼 ──
 * `word_insert_chart` 的列舉裡有 `timeline`、`map`、`scatter`、`box`、`plot`，
 * 它自己的說明裡還寫著「資料走 `curves`」並附了範例。但那份 JSON Schema 既沒有
 * `curves` 也沒有 `events`，而且 `additionalProperties: false`——所以模型照著說明
 * 送出的參數，在到達增益集之前就被擋掉了；更早一步，`required: ['categories']`
 * 讓這幾種圖連送出的機會都沒有，因為它們本來就不用 categories。
 *
 * 兩個 host 的轉送也各漏過一次：`timeline` 曾經整個不能用，因為 host 組物件時
 * 一個欄位一個欄位手寫，漏了 `events`。
 *
 * 這種壞法特別難追：模型是對的，渲染器回報「某個必要欄位沒給」，而真正漏掉的
 * 地方（schema 或 host 的那行 spread）在訊息裡完全沒有出現。所以改成一條鏈上
 * 三個環節都用同一份清單，並在這裡逐環比對：
 *
 *   渲染器讀的欄位 → CHART_DATA_KEYS → host 轉送 → 工具 schema 宣告
 *
 * 前兩環在 svg-chart.test.mjs，後兩環在這裡。
 */
import fs from 'node:fs';
import { CHART_DATA_KEYS } from '../src/taskpane/hosts/svgchart.js';

let pass = 0;
let fail = 0;
function ok(name, cond, detail) {
  if (cond) { pass++; console.log(`PASS  ${name}`); } else {
    fail++; console.log(`FAIL  ${name}${detail === undefined ? '' : `  → ${detail}`}`);
  }
}

const hostsDir = new URL('../src/taskpane/hosts/', import.meta.url);

/**
 * host 要用整批轉送，不要再手寫欄位。
 *
 * 檢查的是 `...chartData(args)` 有沒有出現在呼叫 `chartSvg` 的地方——這是唯一
 * 一種不會隨著新圖表型別而腐壞的寫法。
 */
for (const file of ['word.js', 'powerpoint.js']) {
  const src = fs.readFileSync(new URL(file, hostsDir), 'utf8');
  const call = src.slice(src.indexOf('chartSvg({'), src.indexOf('chartSvg({') + 1400);
  ok(`${file} 整批轉送資料欄位`, /\.\.\.chartData\(args\)/.test(call),
    call.split('\n').slice(0, 3).join(' '));
}

const serverDir = new URL('../../ai-agents-office/server/src/', import.meta.url);
if (!fs.existsSync(serverDir)) {
  console.log('      （找不到 ai-agents-office，跳過工具 schema 的跨 repo 檢查）');
  console.log(`\n${pass} passed, ${fail} failed`);
  process.exit(fail ? 1 : 0);
}

/**
 * 工具 schema 要宣告得出每一種圖的資料欄位。
 *
 * 只有兩個欄位例外，而且是因為那支工具真的不畫那種圖：甘特圖在 Word 是獨立的
 * `word_insert_gantt`，所以 `word_insert_chart` 不需要 `tasks`／`today`。
 */
const CASES = [
  { file: 'wordToolSpec.ts', tool: 'word_insert_chart', skip: ['tasks', 'today'] },
  { file: 'pptToolSpec.ts', tool: 'ppt_add_chart', skip: [] },
];

for (const { file, tool, skip } of CASES) {
  const src = fs.readFileSync(new URL(`services/${file}`, serverDir), 'utf8');
  const start = src.indexOf(`name: '${tool}'`);
  ok(`${file} 找得到 ${tool}`, start > 0, start);
  if (start < 0) continue;
  // 下一支工具開始之前都算這一支的範圍。
  const next = src.indexOf("    name: '", start + 20);
  const body = src.slice(start, next > 0 ? next : src.length);

  const want = CHART_DATA_KEYS.filter((k) => !skip.includes(k));
  const missing = want.filter((k) => !new RegExp(`^\\s{8}${k}:`, 'm').test(body));
  ok(`${tool} 宣告了所有資料欄位`, !missing.length, missing.join(', '));

  // additionalProperties: false 是對的——它擋掉打錯的參數名。但正因為它在，
  // 上面那條漏一個欄位就等於那種圖不能用。
  ok(`${tool} 仍然擋未知參數`, /additionalProperties: false/.test(body));

  /**
   * categories 不能是硬性要求。
   *
   * timeline 走 events、plot 走 curves、map 的分類是地區名——把 categories 列進
   * required 等於宣告「這幾種圖一律不准送」。缺什麼由增益集擋，訊息也比 schema
   * 的驗證錯誤清楚得多。
   */
  ok(`${tool} 沒有硬性要求 categories`, !/required: \['categories'\]/.test(body));
}

/**
 * mermaid 不該拿去畫時程。
 *
 * 甘特圖那條已經有了；時間軸是後來補的，而且原因不一樣——mermaid 的 `timeline`
 * 會把事件平均排開，而一條時間軸真正要講的就是「哪裡密、哪裡疏」。
 */
for (const [file, tool] of [['wordToolSpec.ts', 'word_insert_diagram'], ['pptToolSpec.ts', 'ppt_add_diagram']]) {
  const src = fs.readFileSync(new URL(`services/${file}`, serverDir), 'utf8');
  const start = src.indexOf(`name: '${tool}'`);
  const next = src.indexOf("    name: '", start + 20);
  const body = src.slice(start, next > 0 ? next : src.length);
  ok(`${tool} 把時程導去原生甘特圖`, /gantt/.test(body) && /不要用這支|不要用 mermaid/.test(body));
  ok(`${tool} 把沿革導去原生時間軸`, /timeline/.test(body) && /events/.test(body));
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
