/**
 * 看得到既有的表格，而且填得進去。
 *
 * ── 這支測試的由來 ──
 * 「我要求就是 AI 要把原本檔案全部看一次，上面有甚麼都要看」。在這之前
 * `word_get_overview` 回報的是標題、段落數、大綱、圖片——**完全沒提表格**。
 * 整個 word.js 裡所有 `getCell` 都在「我自己剛建的那張表」裡面，所以一份公司的
 * 報價單範本送進來，模型看到的是一串段落文字：它讀得到「客戶名稱」四個字，
 * 卻不知道那四個字在第 2 列第 1 欄、右邊那格是空的。
 *
 * 所以這裡測兩件事：
 *   · 概覽講不講得出表格的尺寸和欄位名稱
 *   · 填得進去，而且**不會蓋掉已經有內容的格子**
 *
 * 第二條是重點。填表單跟覆寫是兩件事，分不清楚的工具會把人家填好的資料洗掉。
 */
import { executeTool } from '../src/taskpane/hosts/word.js';

let pass = 0;
let fail = 0;
function ok(name, cond, detail) {
  if (cond) { pass++; console.log(`PASS  ${name}`); } else {
    fail++; console.log(`FAIL  ${name}${detail === undefined ? '' : `  → ${detail}`}`);
  }
}

/** 一份公司的報價單範本：左邊欄位名稱、右邊留白。 */
const QUOTE = [
  ['客戶名稱', '', '報價日期', ''],
  ['聯絡人', '', '聯絡電話', ''],
  ['專案名稱', '', '幣別', 'TWD'],
];

function fakeWord(grids = [QUOTE]) {
  const log = { writes: [], syncs: 0 };
  const tables = grids.map((g) => ({
    rowCount: g.length,
    values: g.map((r) => r.slice()),
    getCell: (r, c) => ({
      set value(v) { log.writes.push({ r: r + 1, c: c + 1, v }); },
      body: { font: {}, insertParagraph: () => ({ font: {} }) },
      setCellPadding() {},
      set verticalAlignment(v) {},
      set shadingColor(v) {},
      set columnWidth(v) {},
    }),
    getBorder: () => ({}),
    load() {},
  }));
  const ctx = {
    document: {
      changeTrackingMode: 'Off',
      properties: { title: '報價單範本', lastAuthor: '', load() {} },
      getStyles: () => ({ getByNameOrNullObject: () => ({ isNullObject: true, font: {}, load() {} }) }),
      load() {},
      body: {
        text: '報價單',
        tables: { items: tables, load() {} },
        inlinePictures: { items: [], load() {} },
        paragraphs: { items: [], load() {} },
        getOoxml: () => ({ value: '<xml/>' }),
        insertTable: () => tables[0],
        insertOoxml: () => {},
        insertParagraph: () => ({ insertOoxml: () => {}, styleBuiltIn: '', insertBreak: () => {} }),
        load() {},
      },
    },
    sync: async () => { log.syncs += 1; },
    load() {},
  };
  global.Word = { run: async (fn) => fn(ctx), InsertLocation: {} };
  global.Office = {
    context: {
      requirements: { isSetSupported: () => true },
      document: { url: `q-${Math.random()}.docx`, settings: { get: () => null, set() {}, saveAsync() {} } },
    },
  };
  return log;
}

// ── 概覽看得到表格 ──────────────────────────────────────────────────────────
{
  fakeWord();
  const res = await executeTool('word_get_overview', {});
  const txt = String(res && res.content);
  ok('概覽講得出有幾張表', /表格：1 張/.test(txt), txt.slice(0, 200));
  ok('講得出尺寸', /3 列 × 4 欄/.test(txt), txt.match(/表格 1[^\n]*/));
  ok('講得出空了幾格', /5 格是空的/.test(txt), txt.match(/表格 1[^\n]*/));
  // 第一欄就是欄位清單——有這個，模型才認得出這是一張表單。
  ok('列出第一欄的欄位名稱',
    /客戶名稱/.test(txt) && /聯絡人/.test(txt) && /專案名稱/.test(txt),
    txt.match(/表格 1[^\n]*/));
  ok('告訴模型下一步用哪支工具',
    /word_read_table/.test(txt) && /word_fill_table/.test(txt));
}

// ── 讀一張表 ────────────────────────────────────────────────────────────────
{
  fakeWord();
  const res = await executeTool('word_read_table', { table: 1 });
  const txt = String(res && res.content);
  ok('讀得出整張表', /客戶名稱/.test(txt) && /TWD/.test(txt), txt.slice(0, 160));
  ok('有列號和欄號', /列1/.test(txt) && /欄1/.test(txt), txt.slice(0, 160));
  // 空格要標出來，不然模型分不清「空的」和「讀不到」。
  ok('空格標成（空）', /（空）/.test(txt));
}

// ── 用欄位名稱填 ────────────────────────────────────────────────────────────
{
  const log = fakeWord();
  const res = await executeTool('word_fill_table', {
    table: 1,
    fields: { 客戶名稱: '強茂股份有限公司', 報價日期: '2026-10-12', 聯絡電話: '07-123-4567' },
  });
  ok('填得進去', res && res.ok, JSON.stringify(res).slice(0, 140));
  const at = (r, c) => log.writes.find((w) => w.r === r && w.c === c);
  ok('客戶名稱寫到它右邊那一格', at(1, 2) && at(1, 2).v === '強茂股份有限公司',
    JSON.stringify(log.writes));
  ok('報價日期也是寫右邊', at(1, 4) && at(1, 4).v === '2026-10-12', JSON.stringify(log.writes));
  ok('第二列的聯絡電話也對', at(2, 4) && at(2, 4).v === '07-123-4567', JSON.stringify(log.writes));
  ok('只寫了三格，沒有多寫', log.writes.length === 3, JSON.stringify(log.writes));
}

// ── 不會蓋掉已經有內容的格子 ────────────────────────────────────────────────
{
  /**
   * 「幣別」右邊已經寫著 TWD。
   *
   * 填表單跟覆寫是兩件事。模型以為那格是空的而送了值過來，正確的反應是**不要寫**
   * 並且講出來——把人家填好的資料無聲洗掉，是這種工具最糟的失敗方式。
   */
  const log = fakeWord();
  const res = await executeTool('word_fill_table', {
    table: 1, fields: { 幣別: 'USD', 客戶名稱: '甲公司' },
  });
  ok('有內容的格子不會被蓋掉', !log.writes.some((w) => w.v === 'USD'),
    JSON.stringify(log.writes));
  ok('而且會講出來為什麼', /幣別/.test(String(res.content)) && /已經有內容/.test(String(res.content)),
    String(res.content));
  ok('其他欄位照樣填', log.writes.some((w) => w.v === '甲公司'));
}

// ── 用列欄填 ────────────────────────────────────────────────────────────────
{
  const log = fakeWord();
  await executeTool('word_fill_table', {
    table: 1, cells: [{ row: 3, col: 2, text: '封裝良率改善案' }],
  });
  ok('列欄直接指定也行', log.writes.length === 1
    && log.writes[0].r === 3 && log.writes[0].c === 2, JSON.stringify(log.writes));

  // 超出範圍要擋，而且要說清楚表格多大。
  const bad = await executeTool('word_fill_table', { table: 1, cells: [{ row: 9, col: 9, text: 'x' }] });
  ok('超出範圍會擋', bad && bad.ok === false && /3×4/.test(String(bad.error)),
    String(bad && bad.error));
}

// ── 找不到欄位的時候 ────────────────────────────────────────────────────────
{
  fakeWord();
  const bad = await executeTool('word_fill_table', { table: 1, fields: { 統一編號: '12345678' } });
  ok('找不到的欄位會明講', bad && bad.ok === false && /找不到「統一編號」/.test(String(bad.error)),
    String(bad && bad.error));
  ok('並且指路去 word_read_table', /word_read_table/.test(String(bad.error)));
}

// ── 欄位名稱的小差異要容得下 ────────────────────────────────────────────────
{
  /**
   * 範本上寫的是「客戶名稱：」或「客戶 名稱」，模型送的是「客戶名稱」。
   *
   * 全形空白、冒號、星號這些是版面，不是欄位名稱的一部分。差一個冒號就填不進去，
   * 會讓這支工具在真的範本上幾乎不能用——真的範本上到處都是冒號。
   */
  const log = fakeWord([[['客戶名稱：', '', '　報價日期 ', ''], ['＊聯絡人', '', '', '']]]);
  await executeTool('word_fill_table', {
    table: 1, fields: { 客戶名稱: 'A', 報價日期: 'B', 聯絡人: 'C' },
  });
  ok('冒號、空白、星號不影響比對', log.writes.length === 3,
    JSON.stringify(log.writes));
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
