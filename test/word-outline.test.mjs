/**
 * Pin the paragraph outline against the document that made it necessary.
 *
 * ── What is actually being tested ──
 * The navigator's only job is to make 「段落 19-78」 mean something to a person,
 * and it does that by collapsing two runs — tables and blank spacers — into one
 * row each while keeping every number honest. Those two collapses are ordinary
 * index arithmetic, and ordinary index arithmetic is where this breaks: a table
 * that ends one paragraph early sends every jump after it to the wrong place,
 * silently. Nothing throws. The row just scrolls somewhere else.
 *
 * So the fixture is the real quotation from the screenshot — letterhead, title,
 * a two-panel layout table, a heading, a 13×6 detail table, terms, a signature
 * block — and the assertions check the invariant that matters most:
 *
 *   every paragraph in the document is covered by exactly one row, once.
 *
 * If that holds, no jump can land on the wrong paragraph. If it does not, the
 * panel is lying about where things are, which is worse than not having it.
 */

// ── A document shaped like the one people actually build ─────────────────────

const P = (text, style = 'Normal', extra = {}) => ({
  text, style, styleBuiltIn: style, isListItem: false, tableNestingLevel: 0, ...extra,
});
const CELL = (text) => P(text, 'Normal', { tableNestingLevel: 1 });

const DOC = [
  P(''),                                            // 1
  P('日新數位整合股份有限公司'),                     // 2
  P('SUNRISE DIGITAL INTEGRATION CO., LTD.'),        // 3
  P('台北市信義區松高路 68 號 12 樓'),               // 4
  P('TEL 02-2758-1200　FAX 02-2758-1201'),           // 5
  P('sales@sunrise-digital.com.tw'),                 // 6
  P('報　價　單　QUOTATION', 'Title'),               // 7
  // The 客戶資料 | 報價資訊 pair — one layout table, ten paragraphs.
  ...['客戶資料', '報價資訊', '公司名稱', '宏遠實業股份有限公司', '報價單號',
    'Q2026-0821-014', '聯絡人', '吳孟樺', '報價日期', '2026-08-21'].map(CELL),  // 8–17
  P(''),                                            // 18
  P('報價明細', 'Heading1'),                         // 19
  // A 4×3 detail table — 12 paragraphs.
  ...['項次', '品名', '金額', '1', '網站規劃', '120,000', '2', '前端開發',
    '180,000', '', '總計', '300,000'].map(CELL),                                 // 20–31
  P(''),                                            // 32
  P(''),                                            // 33
  P('報價條件與說明', 'Heading1'),                    // 34
  P('本報價有效期限為 30 天。', 'ListParagraph', { isListItem: true }),  // 35
  P('付款方式為簽約後支付 30%。', 'ListParagraph', { isListItem: true }), // 36
  P('交付時程', 'Heading2'),                          // 37
  P('自簽約日起 90 個日曆天內完成。'),                // 38
];

// ── Boot the host with just enough Word around it ────────────────────────────

let selectionFrom = 0;

globalThis.Word = {
  run: async (fn) => fn({
    document: {
      body: {
        paragraphs: { items: DOC, load() {} },
        getOoxml: () => ({ value: '<w:body/>' }),
      },
      load: () => {},
    },
    sync: async () => {},
    load: () => {},
  }),
};
globalThis.Office = {
  context: { requirements: { isSetSupported: () => true }, document: { addHandlerAsync: () => {} } },
  EventType: { DocumentSelectionChanged: 'sel' },
};
globalThis.window = globalThis;
globalThis.document = {
  createElement: () => ({ getContext: () => null, style: {} }),
  getElementById: () => null,
  head: { appendChild: () => {} },
};

const host = await import('../src/taskpane/hosts/word.js');

let pass = 0, fail = 0;
const ok = (name, cond, detail) => {
  cond ? pass++ : fail++;
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}${cond ? '' : '  → ' + (detail || '')}`);
};

const { total, rows, error } = await host.getOutline();
ok('讀得到大綱', !error, error);
ok('段落總數正確', total === DOC.length, `${total} vs ${DOC.length}`);

// ── The invariant: every paragraph covered exactly once ──────────────────────

const covered = new Array(DOC.length + 1).fill(0);
for (const r of rows) for (let n = r.n; n <= r.to; n++) covered[n]++;
const uncovered = covered.slice(1).map((c, i) => (c === 0 ? i + 1 : 0)).filter(Boolean);
const doubled = covered.slice(1).map((c, i) => (c > 1 ? i + 1 : 0)).filter(Boolean);
ok('每個段落都被某一列涵蓋', !uncovered.length, `漏掉 ${uncovered.join(',')}`);
ok('沒有段落被兩列重複涵蓋', !doubled.length, `重複 ${doubled.join(',')}`);
ok('列的順序是遞增的', rows.every((r, i) => i === 0 || r.n === rows[i - 1].to + 1),
  rows.map((r) => `${r.n}-${r.to}`).join(' '));
ok('沒有反向區間', rows.every((r) => r.to >= r.n));

// ── The two collapses ────────────────────────────────────────────────────────

const tables = rows.filter((r) => r.kind === 'table');
ok('兩張表格各收成一列', tables.length === 2, `${tables.length} 列`);
ok('雙欄方框是段落 8-17', tables[0] && tables[0].n === 8 && tables[0].to === 17,
  tables[0] && `${tables[0].n}-${tables[0].to}`);
ok('明細表是段落 20-31', tables[1] && tables[1].n === 20 && tables[1].to === 31,
  tables[1] && `${tables[1].n}-${tables[1].to}`);
ok('表格列用開頭幾格當辨識文字', tables[0].text.includes('客戶資料'), tables[0].text);
ok('表格預覽不會把整張表倒出來', tables[1].text.split('　').length <= 4, tables[1].text);

const blanks = rows.filter((r) => r.kind === 'blank');
ok('連續空白段落收成一列', blanks.some((b) => b.n === 32 && b.to === 33),
  blanks.map((b) => `${b.n}-${b.to}`).join(' '));
ok('單一空白段落也照樣列出來', blanks.some((b) => b.n === 1 && b.to === 1),
  '空白段落會吃掉編號，藏起來反而讓編號看起來莫名其妙');
ok('空白列帶著它吃掉的段落數', blanks.find((b) => b.n === 32).span === 2);

// ── Structure ────────────────────────────────────────────────────────────────

const headings = rows.filter((r) => r.kind === 'heading');
ok('三個標題都認出來', headings.length === 3, headings.map((h) => h.text).join(' / '));
ok('標題層級對', headings.map((h) => h.level).join(',') === '1,1,2',
  headings.map((h) => h.level).join(','));
ok('Title 不算 Heading', rows.find((r) => r.n === 7).kind === 'body',
  '封面主標題不該撐開大綱縮排');
ok('清單項目標成 list', rows.filter((r) => r.kind === 'list').length === 2);
ok('內文就是 body', rows.find((r) => r.n === 38).kind === 'body');
ok('文字有截斷', rows.every((r) => !r.text || r.text.length <= 47),
  rows.map((r) => (r.text || '').length).join(','));

// 「只看標題」 keeps tables, because a table is structure and it is the row people
// most want to jump to.
const structural = rows.filter((r) => r.kind === 'heading' || r.kind === 'table');
ok('只看標題時剩下標題和表格', structural.length === 5, `${structural.length} 列`);

// ── Cursor ───────────────────────────────────────────────────────────────────

ok('meta 宣告 Word 有大綱', !!host.meta.outline);
ok('游標段落讀得到型別正確', typeof (await host.getCursorParagraph()) === 'number');

// ── Excel must not grow this ─────────────────────────────────────────────────

const excel = await import('../src/taskpane/hosts/excel.js').catch(() => null);
if (excel) {
  ok('Excel 沒有大綱（A1 本來就看得懂）', !excel.meta.outline);
} else {
  ok('Excel 主機載得起來', false, '載入失敗');
}

// ── Degenerate documents ─────────────────────────────────────────────────────

DOC.length = 0;
const empty = await host.getOutline();
ok('空文件不會爆', !empty.error && empty.rows.length === 0 && empty.total === 0,
  JSON.stringify(empty));

DOC.push(CELL('只有一張表'), CELL('沒有別的'));
const allTable = await host.getOutline();
ok('整份都是表格也只收成一列', allTable.rows.length === 1 && allTable.rows[0].to === 2,
  JSON.stringify(allTable.rows));

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
