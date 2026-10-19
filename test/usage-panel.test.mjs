/**
 * 個人區：把 /api/usage 畫成面板。
 *
 * ── 為什麼這個檔案是「真的執行」而不是靜態檢查 ──
 * 窗格的其他測試大多是讀原始碼，因為那些程式碼碰得到 Office.js 或 DOM。但 renderUsage
 * 這一段是純字串運算：給它一包 JSON，回一段 HTML。能跑就該跑——這個 session 已經有兩次
 * 「讀程式碼看不出來、跑一次立刻現形」的經驗（powerpoint.js 的 mixHex 未定義、Office 的
 * 圖示快取），而兩次都是靠實際執行才找到的。
 *
 * ── 這裡真正在防的是什麼 ──
 * 伺服器對管理員回的上限是 Infinity，而 JSON.stringify(Infinity) 是 **null**。
 * 所以面板收到的「無上限」長得跟「上限為 0」很像，而 used/0 是 Infinity、
 * Math.round(Infinity) 是 Infinity——一不小心就會畫出 "Infinity%" 或 "NaN%"，
 * 而且只會發生在管理員帳號上，也就是最不會有人回報的那種人。
 *
 * 其餘幾題都是同一類問題：某個數字缺席時，畫面要說得出話，不能印出 NaN。
 */
import fs from 'fs';

const src = fs.readFileSync(new URL('../src/taskpane/app.js', import.meta.url), 'utf-8');

let failed = 0;
function check(name, ok, detail) {
  if (ok) { console.log(`PASS  ${name}`); return; }
  failed += 1;
  console.error(`FAIL  ${name}${detail ? `\n      ${detail}` : ''}`);
}

// ── 把那一段抽出來單獨執行 ────────────────────────────────────────────────────
//
// 用標記字串切，而不是用行號：行號會在任何人上面加一行註解時失效，而那種失敗
// 看起來像是功能壞了。
const START = 'const USAGE_CATS = [';
const END = '/** Dates in the pane are always about';
const from = src.indexOf(START);
const to = src.indexOf(END);
check('找得到個人區那一段', from > 0 && to > from,
  `START=${from} END=${to}——標記字串被改掉的話要同步更新這裡`);
if (from < 0 || to < from) process.exit(1);

const block = src.slice(from, to);

// 這一段只用到這三個外部函式，其餘（fetch、$、state）只在 openUsage/loadUsage 裡，
// 那兩個不在測試範圍內——它們做的事是取資料，不是決定畫面長什麼樣。
const stubs = `
  const escapeHtml = (v) => String(v).replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  // 真正的 t() 會查翻譯表；這裡只要它把 {0} 代進去，因為測試斷言的是數字不是語言。
  const t = (s, ...a) => String(s).replace(/\\{(\\d+)\\}/g, (_, i) => a[i]);
  const $ = () => ({ hidden: true, setAttribute() {}, innerHTML: '' });
  const state = { apiBase: '', token: '' };
  const httpTrouble = (n) => 'HTTP ' + n;
`;

const mod = await import(
  'data:text/javascript;charset=utf-8,'
  + encodeURIComponent(`${stubs}\n${block}\nexport { renderUsage, renderDays, renderCats, money };`)
);
const { renderUsage, money } = mod;

// ── 測資 ─────────────────────────────────────────────────────────────────────

const today = new Date();
const ym = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}`;
const day = (n) => `${ym}-${String(n).padStart(2, '0')}`;

const normal = {
  isBeta: false,
  limit: 50,
  total: { totalInput: 120000, totalOutput: 34000, totalInvocations: 42, cost: 12.5 },
  summary: [
    { date: day(1), total_input: 1000, total_output: 400, invocation_count: 3, cost: 2.5 },
    { date: day(2), total_input: 3000, total_output: 900, invocation_count: 9, cost: 10.0 },
  ],
  byCategory: {
    addin: [{ date: day(1), cost: 8.0 }, { date: day(2), cost: 2.0 }],
    document: [{ date: day(2), cost: 2.5 }],
    team: [], email: [], km: [],
  },
};

const clone = (o) => JSON.parse(JSON.stringify(o));
const noNaN = (h) => !/NaN|Infinity|undefined|\[object/.test(h);

// 1. 一般情況
{
  const h = renderUsage(normal);
  check('一般：百分比算得對', h.includes('>25%<'), h.match(/usage-num">[^<]*/)?.[0]);
  check('一般：金額同時顯示已用與上限', h.includes('$12.50') && h.includes('$50.00'));
  check('一般：未達七成不上警示色', !/usage-fill (warn|over)/.test(h));
  check('一般：增益集那一列被標成 mine', /usage-row mine/.test(h));
  check('一般：增益集的金額與佔比都在', h.includes('$10.00') && h.includes('80%'));
  check('一般：沒有用過的來源不列出來', !h.includes('KM 助手') && !h.includes('AI 團隊'));
  check('一般：提問次數與 token 都有', h.includes('42') && h.includes('154,000'));
  check('一般：沒有 NaN / undefined', noNaN(h), h.slice(0, 400));
}

// 2. 管理員：伺服器回 Infinity，JSON 變成 null
{
  const d = clone(normal); d.limit = null;
  const h = renderUsage(d);
  check('無上限：說「無使用上限」而不是 0%', h.includes('無使用上限'), h.slice(0, 300));
  check('無上限：標題改印金額', h.includes('$12.50'));
  check('無上限：不畫進度條', !h.includes('usage-bar'));
  check('無上限：沒有 Infinity 或 NaN 漏出來', noNaN(h), h.slice(0, 400));
}

// 3. 上限是 0（設定失誤）——不能變成除以零
{
  const d = clone(normal); d.limit = 0;
  const h = renderUsage(d);
  check('上限 0：走無上限那條，不印 Infinity%', noNaN(h) && h.includes('無使用上限'));
}

// 4. 超量與接近超量
{
  const warn = clone(normal); warn.total.cost = 40;      // 80%
  const over = clone(normal); over.total.cost = 73.2;    // 146%
  check('接近上限：七成以上轉警示色', /usage-fill warn/.test(renderUsage(warn)));
  const h = renderUsage(over);
  check('超量：轉危險色', /usage-fill over/.test(h));
  check('超量：進度條夾在 100%', h.includes('width:100%') && !/width:1[1-9]\d/.test(h));
  check('超量：數字照實說超了多少，不夾成 100', h.includes('>146%<'),
    h.match(/usage-num">[^<]*/)?.[0]);
}

// 5. 完全沒有用量
{
  const h = renderUsage({ isBeta: false, limit: 50, total: {}, summary: [], byCategory: {} });
  check('零用量：講得出話', h.includes('還沒有用量'), h.slice(0, 300));
  check('零用量：不畫空圖表', !h.includes('usage-days'));
  check('零用量：沒有 NaN', noNaN(h), h.slice(0, 300));
}

// 6. 伺服器回了半包資料
{
  const h = renderUsage({});
  check('空回應：不炸、也不印 NaN', noNaN(h), h.slice(0, 300));
}

// 7. beta 模式：累計而非當月
{
  const d = clone(normal); d.isBeta = true;
  const h = renderUsage(d);
  check('beta：標題說「累計」', h.includes('累計已使用'));
  check('beta：不談每月歸零', !h.includes('歸零：') && h.includes('不歸零'));
}

// 8. summary 跨月——正式模式只能算當月，否則長條加起來對不上標題的數字
{
  const d = clone(normal);
  d.summary.push({ date: '2020-01-05', total_input: 1, total_output: 1, invocation_count: 1, cost: 999 });
  d.byCategory.addin.push({ date: '2020-01-05', cost: 999 });
  const h = renderUsage(d);
  check('跨月：舊資料不計入每日最高', !h.includes('$999.00'), h.match(/單日最高[^<]*/)?.[0]);
  check('跨月：舊資料不計入來源明細', h.includes('$10.00'));
}

// 9. money() 本身
{
  check('money：四捨五入到分', money(1.005) === '$1.00' || money(1.005) === '$1.01');
  check('money：給 null 回 $0.00', money(null) === '$0.00');
  check('money：給字串也算得出來', money('3.5') === '$3.50');
}

console.log(failed ? `\n${failed} failed` : '\n全部通過');
process.exit(failed ? 1 : 0);
