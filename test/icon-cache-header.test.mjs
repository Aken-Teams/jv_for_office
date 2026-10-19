/**
 * 圖片不能送 `Cache-Control: no-store`。
 *
 * ── 這條規則是哪一天留下來的 ──
 * 三個頻道的增益集圖示全部顯示成一顆通用六邊形。Office 抓得到檔案——開發伺服器的
 * 請求紀錄裡是乾淨的 200，位元組數一個不差——然後畫預設圖。沒有錯誤訊息、沒有重試、
 * 什麼線索都沒有。
 *
 * 原因寫在微軟的疑難排解文件裡：
 *
 *   "Office Add-ins require the ability to cache image resources, and if the
 *    server hosting an image resource adds CACHE-CONTROL directives to the
 *    response header, Office automatically substitutes a generic or default image."
 *   https://learn.microsoft.com/office/dev/add-ins/testing/testing-and-troubleshooting
 *
 * 而開發伺服器對**所有**回應都送 no-store，理由是窗格在開發時一直在改、WebView
 * 快取得很兇。那個理由對 HTML/JS/CSS 完全成立，只是沒有人想到它會順手把圖示也關掉。
 *
 * ── 為什麼值得一個測試 ──
 * 找到這個原因花了一整個下午和五個錯誤的推論（sideload 機制、PNG 的 pHYs 區塊、
 * 網址快取鍵、manifest 版本號、圖示尺寸）。而標頭在第一次比對時就印出來過——
 * 三台主機都一樣，看起來正是「正常」的樣子。
 *
 * 一個沒有症狀、沒有訊息、而且看起來像正常設定的失敗，不能靠人記得。
 */
import fs from 'fs';

const src = fs.readFileSync(new URL('../tools/dev-server.mjs', import.meta.url), 'utf-8');

let failed = 0;
function check(name, ok, detail) {
  if (ok) { console.log(`PASS  ${name}`); return; }
  failed += 1;
  console.error(`FAIL  ${name}${detail ? `\n      ${detail}` : ''}`);
}

// 註解裡就引用了 'no-store' 和那段文件，拿原始字串比對會被註解滿足。
const code = src
  .replace(/\/\*[\s\S]*?\*\//g, (m) => ' '.repeat(m.length))
  .replace(/(^|[^:])\/\/[^\n]*/g, (m, p1) => p1 + ' '.repeat(m.length - p1.length));

check('Cache-Control 不是寫死的 no-store',
  !/'Cache-Control'\s*:\s*'no-store'/.test(code),
  '圖片一旦送 no-store，Office 會改畫通用圖示——而且不會有任何錯誤訊息。');

check('有分辨圖片副檔名',
  /\.(?:png|jpe?g|gif|svg|ico)\$?/i.test(code) && /png/i.test(code),
  '要能區分圖片和 HTML/JS/CSS：後者仍然需要 no-store，否則開發時會對著舊程式碼除錯。');

check('圖片走可快取的那一條',
  /max-age=\d+/.test(code),
  'Office 必須被允許快取圖片資源，否則它會替換成預設圖。');

check('非圖片仍然 no-store',
  /no-store/.test(code),
  'HTML/JS/CSS 的 no-store 是刻意的——窗格在開發時一直在改，WebView 快取得很兇。');

// 三元運算要在同一個 Cache-Control 上，而不是兩個分開的 writeHead——
// 分成兩條路的話，之後改其中一條很容易忘了另一條。
const ccLine = code.match(/'Cache-Control'[^\n]*/)?.[0] || '';
check('兩種情況共用同一個標頭設定',
  ccLine.includes('?') && ccLine.split('?')[1].includes(':'),
  ccLine || '找不到 Cache-Control');

console.log(failed ? `\n${failed} failed` : '\n全部通過');
process.exit(failed ? 1 : 0);
