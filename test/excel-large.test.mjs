/**
 * 大檔案：看不看得完，以及看不完的時候會怎樣。
 *
 * ── 這個檔案在守什麼 ──
 * 使用者丟進來的活頁簿通常不小——BOM、出貨明細、料號主檔動輒上萬列。以前有兩個
 * 上限擋在那裡，而且擋錯了地方：
 *
 *   1. `excel_read_range` 超過兩萬格就截斷，然後模型拿著六分之一的資料回答
 *      「最大的三筆是…」。每個數字都是真的讀到的，所以它完全不違反「絕對不要編
 *      造儲存格的值」那條規則，但結論是錯的，而且沒有人看得出來。
 *   2. 比對兩側各兩萬格，十欄就是兩千列。可是比對的 diff 是在窗格的 JavaScript
 *      裡算的、根本不會進模型——那個上限保護的是別人的資源。
 *
 * 所以修法不一樣：第一個是**把會產生殘缺答案的那條路拿掉**（超過就擋，指去
 * script），第二個是**分批讀，把天花板拉高**。
 *
 * 下面每一條都是在釘「安靜地答錯」這件事不能再發生。
 */
let pass = 0, fail = 0;
const ok = (name, cond, detail) => {
  cond ? pass++ : fail++;
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}${cond ? '' : '  → ' + (detail || '')}`);
};

const COL = (i) => {
  let n = i + 1, out = '';
  while (n > 0) { const r = (n - 1) % 26; out = String.fromCharCode(65 + r) + out; n = (n - r - 1) / 26; }
  return out;
};

/** How many slices Excel was asked for — the whole point of the chunked read. */
let LOADS = [];

/**
 * A sheet that can be any size without allocating it.
 *
 * Rows are generated on demand, so a 「50,000 列」 test costs nothing until
 * something actually reads them — and what it reads is checked against what it
 * asked for, which is how a chunked read that quietly drops a slice gets caught.
 */
function fakeSheet(name, rows, cols, cell = (r, c) => `r${r}c${c}`) {
  const range = (top, left, nRows, nCols) => {
    const r = {
      isNullObject: false,
      rowIndex: top, columnIndex: left, rowCount: nRows, columnCount: nCols,
      address: `${name}!${COL(left)}${top + 1}:${COL(left + nCols - 1)}${top + nRows}`,
      load(what) {
        if (!String(what).includes('values')) return;
        LOADS.push({ top, left, rows: nRows, cols: nCols });
        r.values = Array.from({ length: nRows }, (_, i) =>
          Array.from({ length: nCols }, (_, j) => cell(top + i, left + j)));
      },
    };
    return r;
  };
  return {
    name,
    isNullObject: false,
    load() {},
    getUsedRangeOrNullObject: () => range(0, 0, rows, cols),
    getRange: (a1) => {
      const m = /([A-Z]+)(\d+):([A-Z]+)(\d+)/.exec(String(a1).toUpperCase());
      if (!m) return range(0, 0, rows, cols);
      const num = (s) => [...s].reduce((a, ch) => a * 26 + (ch.charCodeAt(0) - 64), 0) - 1;
      return range(+m[2] - 1, num(m[1]), +m[4] - +m[2] + 1, num(m[3]) - num(m[1]) + 1);
    },
    getRangeByIndexes: (t, l, nr, nc) => {
      const r = range(t, l, nr, nc);
      // The report is written back through this, so the fake has to accept the
      // writes as well as answer the reads.
      r.format = { font: { bold: undefined }, fill: { color: undefined }, autofitColumns() {} };
      r.clear = () => {};
      return r;
    },
    getUsedRange: () => ({
      clear() {},
      format: { autofitColumns() {}, autofitRows() {} },
    }),
    activate() {},
  };
}

let SHEETS = {};
const ctx = {
  workbook: {
    worksheets: {
      getItemOrNullObject: (n) => SHEETS[n] || { isNullObject: true, load() {} },
      getItem: (n) => SHEETS[n],
      add: (n) => { SHEETS[n] = fakeSheet(n, 1, 1); return SHEETS[n]; },
      load() {},
      get items() { return Object.values(SHEETS); },
    },
  },
  sync: async () => {},
};

globalThis.Excel = { run: async (fn) => fn(ctx) };
globalThis.Office = {
  context: {
    requirements: { isSetSupported: () => true },
    document: { url: 'C:\\b.xlsx', addHandlerAsync: () => {} },
  },
  EventType: { DocumentSelectionChanged: 'sel' },
};
globalThis.window = globalThis;
globalThis.document = {
  createElement: () => ({ getContext: () => null, style: {} }),
  getElementById: () => null, head: { appendChild: () => {} },
};

const host = await import('../src/taskpane/hosts/excel.js');
const read = (args) => host.executeTool('excel_read_range', args);

// ── 讀取：超過上限要擋，不要截斷 ──────────────────────────────────────────────
{
  SHEETS = { 出貨明細: fakeSheet('出貨明細', 12431, 14) };
  LOADS = [];
  const r = await read({ sheet: '出貨明細' });

  ok('大範圍被擋下來，不是回一半', r.ok === false, JSON.stringify(r).slice(0, 160));
  // 最重要的一條：擋下來的時候一格都不該讀。讀了再丟掉只是浪費，讀了再回傳才是
  // 那個 bug。
  ok('而且完全沒有讀任何值', LOADS.length === 0, JSON.stringify(LOADS));

  const msg = String(r.content || r.error || '');
  ok('說出實際多大', /12431/.test(msg) && /14/.test(msg), msg.slice(0, 120));
  ok('指去 excel_run_script', /excel_run_script/.test(msg), msg);
  ok('說明 script 沒有格數限制', /沒有格數限制/.test(msg), msg);
  // 「不行」會被重試；「改用這個」不會。而且要擋住「那我讀少一點再回答」。
  ok('明講不要讀一部分再回答', /不要.*讀一部分|不要.*一部分/.test(msg), msg);
  ok('給了一個可以直接抄的範圍', /range: "[A-Z]+\d+:[A-Z]+\d+"/.test(msg), msg);
}

{
  // 刻意取樣是可以的——講出來就沒問題，而且會出現在對話紀錄裡。
  SHEETS = { 出貨明細: fakeSheet('出貨明細', 12431, 14) };
  LOADS = [];
  const r = await read({ sheet: '出貨明細', range: 'A1:N100' });
  ok('明確指定的小範圍照讀', r.ok === true, JSON.stringify(r).slice(0, 160));
  ok('讀到的就是指定的那塊', LOADS.length === 1 && LOADS[0].rows === 100, JSON.stringify(LOADS));
  ok('回傳裡沒有截斷警告', !/截斷/.test(String(r.content || '')), String(r.content || '').slice(0, 120));
}

{
  // 剛好在上限內的整張表要能讀完，不能被新的擋法誤傷。
  SHEETS = { 小表: fakeSheet('小表', 1000, 8) };
  LOADS = [];
  const r = await read({ sheet: '小表' });
  ok('上限內的表照讀', r.ok === true);
  ok('而且是整張 1000 列', LOADS.length === 1 && LOADS[0].rows === 1000, JSON.stringify(LOADS));
}

// ── 比對：分批讀，天花板拉高 ──────────────────────────────────────────────────
/**
 * 比對的上限跟讀取的上限保護的是不同東西。
 *
 * 讀取進模型的上下文，所以要小。比對的兩份資料在窗格的 JavaScript 裡 diff 完才
 * 產生報表，模型從頭到尾沒看過那些列——所以沿用讀取的上限，是拿別人的限制來卡
 * 自己。十欄兩千列就滿，而 BOM 隨便就超過。
 */
{
  SHEETS = {
    今年: fakeSheet('今年', 9000, 6, (r, c) => (c === 0 ? `P${r}` : `${r}-${c}`)),
    去年: fakeSheet('去年', 9000, 6, (r, c) => (c === 0 ? `P${r}` : `${r}-${c}`)),
    比對結果: fakeSheet('比對結果', 1, 1),
  };
  LOADS = [];
  const r = await host.executeTool('excel_compare_source',
    { sheet: '今年', compare_with_sheet: '去年', headers: false, output_sheet: '比對結果' });
  ok('54,000 格的比對跑得完', r.ok === true, JSON.stringify(r).slice(0, 200));

  const readRows = LOADS.filter((l) => l.cols === 6).reduce((a, l) => a + l.rows, 0);
  // 兩側各 9,000 列，全部都要被讀到——一列都不能少，少掉的會被報成「只在另一邊」。
  ok('兩側 9,000 列全部讀完', readRows === 18000, `${readRows}`);
  ok('而且是分批讀的，不是一次撈', LOADS.filter((l) => l.cols === 6).length > 2,
    `${LOADS.filter((l) => l.cols === 6).length} 批`);
  ok('沒有宣稱截斷', !/只比對了前/.test(String(r.content || '')), String(r.content || '').slice(0, 200));
}

{
  // 超過新天花板還是會截，但要說得出從第幾列開始沒比到——「有差異」和「沒看到」
  // 長得一模一樣，不講就會被當成發現。
  SHEETS = {
    巨表: fakeSheet('巨表', 200000, 6),
    對照: fakeSheet('對照', 200000, 6),
    結果: fakeSheet('結果', 1, 1),
  };
  LOADS = [];
  const r = await host.executeTool('excel_compare_source',
    { sheet: '巨表', compare_with_sheet: '對照', headers: false, output_sheet: '結果' });
  const text = String(r.content || r.error || '');
  ok('超過天花板時有講', /只比對了前/.test(text), text.slice(0, 220));
  ok('而且說得出從第幾列開始沒比到', /第 \d+ 列之後沒有比到/.test(text), text.slice(0, 220));
  ok('天花板比舊的兩千列高很多',
    (text.match(/只比對了前 (\d+) 列/) || [])[1] > 20000,
    (text.match(/只比對了前 (\d+) 列/) || [])[1]);
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
