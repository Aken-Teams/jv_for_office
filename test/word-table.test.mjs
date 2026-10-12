/**
 * Drive the two layout tools against a recording fake of Word.
 *
 * ── Why these two, and why this way ──
 * Both exist because of a side-by-side comparison with a real quotation
 * template, where the wording was already right and the page was still visibly
 * not it. Two differences did most of that damage:
 *
 *   · amounts left-aligned, and a 總計 row styled exactly like the rows above it
 *   · 客戶資料 and 報價資訊 stacked down the page instead of sitting side by side
 *
 * Neither is catchable by reading the code. Both are index arithmetic — which
 * column an alignment lands on, which rows count as totals, which cell each
 * panel is written into — and index arithmetic is exactly the kind of thing that
 * looks right and is off by one. An off-by-one here is not a crash; it is a
 * document where the wrong column is right-aligned, which nobody notices until a
 * customer does.
 *
 * So the fake records every property assignment per cell, and the assertions
 * check the grid that came out, not that the calls were made.
 */
import fs from 'fs';

// ── A Word that remembers what was done to it ────────────────────────────────

function fakeCell(r, c, log) {
  const props = {};
  const cell = {
    body: {
      font: {},
      insertOoxml: (xml, loc) => log.ooxml.push({ r, c, xml, loc }),
      insertParagraph: (text) => ({ text }),
    },
    setCellPadding: (loc, pt) => log.padding.push({ r, c, loc, pt }),
    _props: props,
  };
  // A property bag that records assignment rather than storing silently, so a
  // value written to the wrong cell is visible in the transcript.
  for (const name of ['horizontalAlignment', 'verticalAlignment', 'columnWidth', 'shadingColor', 'value']) {
    Object.defineProperty(cell, name, {
      set: (v) => { props[name] = v; },
      get: () => props[name],
    });
  }
  return cell;
}

/** 一張 A4、一吋邊界的文件，文字欄大約就是這麼寬。 */
const TABLE_WIDTH_PT = 451;

function fakeTable(rows, cols, log) {
  const cells = new Map();
  return {
    rowCount: rows,
    styleBuiltIn: '', headerRowCount: -1, styleBandedRows: null, styleFirstColumn: null,
    /**
     * 寬度要讀得回來。
     *
     * 真正的 Office 會在插入時把表格撐到文字欄寬，而自動欄寬是**先量再分配**——
     * 量不到就什麼都不做。假物件少一個 load()，整段 withTracking 會丟 TypeError，
     * 然後對齊、合計列、欄寬全部靜靜地沒被套用。
     */
    width: undefined,
    load(what) {
      if (String(what).includes('width')) this.width = TABLE_WIDTH_PT;
    },
    getCell(r, c) {
      const key = `${r},${c}`;
      if (!cells.has(key)) cells.set(key, fakeCell(r, c, log));
      return cells.get(key);
    },
    getBorder(location) {
      const b = { location };
      log.borders.push(b);
      return b;
    },
    _cells: cells,
  };
}

function fakeWord() {
  const log = { tables: [], ooxml: [], padding: [], borders: [], syncs: 0 };
  const makeTable = (r, c, loc, values) => {
    const t = fakeTable(r, c, log);
    log.tables.push({ rows: r, cols: c, loc, values, table: t });
    return t;
  };
  /**
   * 樣式集合。
   *
   * ensureBaseStyles 是**改樣式定義**，不是在段落上塗字型——這是「一開始就對」
   * 跟「最後再修一次」的差別。假物件答不出樣式，那段就被 catch 吃掉，而測試會以
   * 為它有跑。
   */
  log.styles = {};
  const fakeStyle = (name, exists) => {
    const st = {
      isNullObject: !exists,
      font: {},
      load() {},
    };
    for (const k of ['name', 'size', 'color', 'bold']) {
      Object.defineProperty(st.font, k, {
        set: (v) => {
          log.styles[name] = log.styles[name] || {};
          log.styles[name][k] = v;
        },
        get: () => (log.styles[name] || {})[k],
      });
    }
    return st;
  };
  // 一台繁體中文的 Word：內建樣式用的是中文名字，英文名查不到。
  const REAL_STYLES = new Set(['內文', '標題 1', '標題 2', '標題 3', '標題 4',
    '標題', '清單段落', '圖表標號', '引文']);

  const ctx = {
    document: {
      changeTrackingMode: 'Off',
      getStyles: () => ({
        getByNameOrNullObject: (n) => fakeStyle(n, REAL_STYLES.has(n)),
      }),
      body: {
        insertTable: makeTable,
        insertOoxml: (xml) => log.ooxml.push({ xml, loc: 'End' }),
        /**
         * 新段落也要能收 OOXML。
         *
         * 因為 `body.insertOoxml(..., 'End')` 會把片段的第一段黏進文件最後一段，
         * 所以寫入改成「先開一個空段落，再 Replace」。假的 body 如果回一個什麼
         * 都不會的空物件，那條路就會丟 TypeError——而更糟的情況是它悄悄吞掉，
         * 測試照樣綠燈，功能其實壞的。
         */
        insertParagraph: () => ({
          insertOoxml: (xml) => log.ooxml.push({ xml, loc: 'End' }),
          styleBuiltIn: '',
          insertBreak: () => {},
        }),
        /**
         * 不能是空的。
         *
         * 封面是先寫幾段 OOXML、再抓最後一段當錨點把資訊表接在後面。items 是空
         * 陣列的話錨點是 undefined，表格整段靜靜地沒建出來——而用 `if (table)`
         * 包起來的斷言就會「通過」，因為它根本沒跑。
         */
        paragraphs: {
          items: Array.from({ length: 8 }, () => ({
            insertTable: makeTable,
            insertOoxml: (xml, loc) => log.ooxml.push({ xml, loc }),
            insertParagraph: () => ({ styleBuiltIn: '' }),
            insertBreak: () => {},
            getRange: () => ({ insertField: () => ({}) }),
          })),
          load() {},
        },
        // executeTool photographs the document before every write so the change
        // is undoable; the fake answers with an empty body rather than nothing,
        // because a snapshot failure would be swallowed into ok:false and the
        // real assertions below would never run.
        getOoxml: () => ({ value: '<w:body/>' }),
      },
    },
    sync: async () => { log.syncs++; },
    load: () => {},
  };
  return {
    log,
    Word: {
      run: async (fn) => fn(ctx),
      // The host feature-gates on these; the fake claims the 1.3 surface the
      // tools actually use and nothing beyond it.
      RequirementSet: {},
    },
  };
}

// ── Boot the host module with just enough browser and Office around it ────────

const { Word, log } = fakeWord();
globalThis.Word = Word;
globalThis.Office = {
  context: {
    // Every gate answers yes for 1.x so the tools under test are reachable;
    // the point here is the arithmetic, not the gating.
    requirements: { isSetSupported: () => true },
    // 可寫：ensureBaseStyles 是一份文件只做一次，換文件才會再做一次。
    document: { url: 'C:/first.docx', addHandlerAsync: () => {} },
  },
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
const reset = () => { log.tables.length = 0; log.ooxml.length = 0; log.padding.length = 0; log.borders.length = 0; };

// ── word_insert_table ────────────────────────────────────────────────────────
// The shape of a real quotation: six columns, three of them money, three total
// rows at the bottom.

reset();
const QUOTE = [
  ['項次', '品名規格', '數量', '單位', '單價', '金額'],
  ['1', 'Wafer MRP 供需查詢系統', '1', '式', '480,000', '480,000'],
  ['2', '教育訓練', '2', '天', '15,000', '30,000'],
  ['', '', '', '', '小計', '510,000'],
  ['', '', '', '', '營業稅 5%', '25,500'],
  ['', '', '', '', '總計', '535,500'],
];
await host.executeTool('word_insert_table', {
  values: QUOTE,
  align: ['center', 'left', 'right', 'center', 'right', 'right'],
  widths: [36, 180, 40, 36, 78, 80],
  total_rows: 3,
});

const t = log.tables[0].table;
const at = (r, c, prop) => t.getCell(r, c)._props[prop];

ok('表格尺寸', log.tables[0].rows === 6 && log.tables[0].cols === 6,
  `${log.tables[0].rows}×${log.tables[0].cols}`);

// The alignment must follow the COLUMN, header included — a right-aligned money
// column under a left-aligned title is the exact thing this is fixing.
ok('金額欄整欄靠右（含表頭）',
  [0, 1, 2, 3, 4, 5].every((r) => at(r, 5, 'horizontalAlignment') === 'Right'),
  [0, 1, 2, 3, 4, 5].map((r) => at(r, 5, 'horizontalAlignment')).join('|'));
ok('品名欄整欄靠左',
  [0, 1, 2].every((r) => at(r, 1, 'horizontalAlignment') === 'Left'));
ok('項次欄置中',
  at(0, 0, 'horizontalAlignment') === 'Centered');
ok('對齊沒有串欄',
  at(0, 1, 'horizontalAlignment') === 'Left' && at(0, 2, 'horizontalAlignment') === 'Right',
  'left/right 相鄰時最容易錯位');

ok('欄寬逐欄對上', [36, 180, 40, 36, 78, 80].every((w, c) => at(0, c, 'columnWidth') === w),
  [0, 1, 2, 3, 4, 5].map((c) => at(0, c, 'columnWidth')).join('|'));

// total_rows counts from the BOTTOM. Row 2 is the last line item and must stay
// plain; rows 3–5 are 小計/稅/總計.
ok('最後三列是合計列', [3, 4, 5].every((r) => at(r, 0, 'shadingColor') === '#1F3864'),
  [0, 1, 2, 3, 4, 5].map((r) => at(r, 0, 'shadingColor') || '-').join('|'));
// 明細列現在也會上色（白／淡灰相間），但絕不能沾到合計列的深底。
ok('合計列的深底沒有外溢到明細列',
  [1, 2].every((r) => at(r, 0, 'shadingColor') !== '#1F3864'),
  [0, 1, 2].map((r) => at(r, 0, 'shadingColor') || '-').join('|'));
ok('明細列是白／淡灰相間',
  at(1, 0, 'shadingColor') === '#FFFFFF' && at(2, 0, 'shadingColor') === '#F2F2F2',
  `${at(1, 0, 'shadingColor')}|${at(2, 0, 'shadingColor')}`);
ok('表頭黑底白字',
  at(0, 0, 'shadingColor') === '#000000' && t.getCell(0, 0).body.font.color === '#FFFFFF',
  `${at(0, 0, 'shadingColor')}|${t.getCell(0, 0).body.font.color}`);
ok('明細列的字是黑的', t.getCell(1, 0).body.font.color === '#000000',
  String(t.getCell(1, 0).body.font.color));
ok('合計列白色粗體',
  [3, 4, 5].every((r) => t.getCell(r, 5).body.font.color === '#FFFFFF' && t.getCell(r, 5).body.font.bold === true));
ok('合計列仍然靠右', at(5, 5, 'horizontalAlignment') === 'Right',
  '深底白字不該覆蓋掉對齊');

// total_rows must never eat the header, however large the number.
reset();
await host.executeTool('word_insert_table', { values: [['A', 'B'], ['1', '2']], total_rows: 99 });
const t2 = log.tables[0].table;
ok('total_rows 過大也不會吃掉表頭',
  t2.getCell(0, 0)._props.shadingColor === '#000000',
  String(t2.getCell(0, 0)._props.shadingColor));

// Ragged input is normal from a model; short rows get padded rather than
// producing a table with holes in it.
reset();
await host.executeTool('word_insert_table', { values: [['A', 'B', 'C'], ['1']] });
ok('短列補齊', log.tables[0].values[1].length === 3, JSON.stringify(log.tables[0].values[1]));

// executeTool turns a throw into { ok: false, content } so the model gets to
// read the reason and retry, rather than the run dying. So refusals are read off
// the result, not caught.
const refuse = async (tool, args) => {
  const r = await host.executeTool(tool, args);
  return r.ok ? '' : String(r.content || r.error || '');
};

reset();
let refused = await refuse('word_insert_table', { values: [] });
ok('空表格會被擋掉', refused.includes('values'), refused);

// ── word_insert_layout ───────────────────────────────────────────────────────
// The 客戶資料 | 報價資訊 pair. One row, one cell per panel.

reset();
await host.executeTool('word_insert_layout', {
  columns: [
    { title: '客戶資料', fields: { 客戶名稱: '強合科技', 聯絡人: '陳詠樺' } },
    { title: '報價資訊', fields: { 報價單號: 'ZH2026072401', 報價日期: '2026-07-24' } },
  ],
  widths: [225, 225],
  panel: true,
});

ok('並排是一列兩欄', log.tables[0].rows === 1 && log.tables[0].cols === 2,
  `${log.tables[0].rows}×${log.tables[0].cols}`);
ok('兩欄各收到自己的內容', log.ooxml.length === 2, `${log.ooxml.length} 段`);
ok('左欄是客戶資料', log.ooxml[0].c === 0 && log.ooxml[0].xml.includes('客戶名稱'));
ok('右欄是報價資訊', log.ooxml[1].c === 1 && log.ooxml[1].xml.includes('報價單號'));
ok('內容沒有跑錯欄', !log.ooxml[0].xml.includes('報價單號') && !log.ooxml[1].xml.includes('客戶名稱'));
ok('用 Replace 蓋掉儲存格原本的空段落', log.ooxml.every((o) => o.loc === 'Replace'),
  '否則每個區塊上面都會多一行空白');
ok('欄寬對上', log.tables[0].table.getCell(0, 0)._props.columnWidth === 225);
ok('panel=true 有底色', log.tables[0].table.getCell(0, 1)._props.shadingColor === '#F2F6FA');
ok('panel=true 有框線', log.borders.some((b) => b.type === 'Single'), JSON.stringify(log.borders));
ok('內距四邊都設', log.padding.filter((p) => p.c === 0).length === 4);

// A letterhead is the same primitive with the borders off and the right-hand
// column right-aligned.
reset();
await host.executeTool('word_insert_layout', {
  columns: [
    { text: '強合科技股份有限公司' },
    { text: '台北市內湖區\n02-1234-5678', align: 'right' },
  ],
});
ok('信頭不畫框線', log.borders.every((b) => b.type === 'None'), JSON.stringify(log.borders.map((b) => b.type)));
ok('信頭沒有底色', log.tables[0].table.getCell(0, 0)._props.shadingColor === undefined);
ok('地址靠右', log.ooxml[1].xml.includes('w:val="right"'), log.ooxml[1].xml.slice(0, 200));
ok('地址的兩行各自成段', (log.ooxml[1].xml.match(/<w:p[ >]/g) || []).length === 2,
  '\\n 應該要分段，不是變成一行');

reset();
refused = await refuse('word_insert_layout', { columns: [{ text: '只有一個' }] });
ok('一欄會被擋掉並說明該用哪個工具', refused.includes('word_insert_block'), refused);

refused = await refuse('word_insert_layout', { columns: Array(5).fill({ text: 'x' }) });
ok('五欄會被擋掉', refused.includes('四欄'), refused);

// ── the rule ─────────────────────────────────────────────────────────────────

reset();
await host.executeTool('word_insert_block', { kind: 'rule', color: '#C00000' });
const rule = log.ooxml[0].xml;
ok('分隔線是段落下框線', rule.includes('<w:pBdr>') && rule.includes('w:bottom'));
ok('分隔線用指定顏色', rule.includes('w:color="C00000"'), rule.slice(0, 300));
ok('預設 1.5 點 = w:sz 12', rule.includes('w:sz="12"'), rule.slice(0, 300));

// thickness is separate from size on purpose: every other kind treats `size` as
// a font size, and a rule drawn at 10pt is a black bar.
reset();
await host.executeTool('word_insert_block', { kind: 'rule', size: 10 });
ok('size 不會被誤當成線粗', log.ooxml[0].xml.includes('w:sz="12"'),
  log.ooxml[0].xml.slice(0, 300));
reset();
await host.executeTool('word_insert_block', { kind: 'rule', thickness: 3 });
ok('thickness 才控制線粗', log.ooxml[0].xml.includes('w:sz="24"'), log.ooxml[0].xml.slice(0, 300));

// ── both tools must be declared, not just implemented ────────────────────────

const src = fs.readFileSync(new URL('../src/taskpane/hosts/word.js', import.meta.url), 'utf8');
for (const name of ['word_insert_layout']) {
  ok(`${name} 有列進 supportedTools`, new RegExp(`'${name}',`).test(src));
  ok(`${name} 有 i18n 進度字串`, new RegExp(`${name}: t\\(`).test(src));
}
const spec = fs.readFileSync('d:/github/ai-agents-office/server/src/services/wordToolSpec.ts', 'utf8');
ok('後端規格宣告了 word_insert_layout', spec.includes("name: 'word_insert_layout'"));
ok('後端規格宣告了表格對齊', spec.includes('align:') && spec.includes('total_rows:'));
ok('後端規格宣告了 rule', spec.includes("'rule'") && spec.includes('thickness:'));

/**
 * 沒給欄寬的時候，欄寬要自己算。
 *
 * ── 回報的畫面 ──
 * 「表格真的很醜」。Word 預設每一欄一樣寬，所以裝著「F-01」的編號欄跟裝著一整句
 * 說明的欄位一樣寬——結果就是編號折成兩行、表頭「編號」斷成 編／號、而說明欄被擠
 * 在四分之一的版面裡，旁邊三分之二是空的。
 *
 * 這支工具一直都收 `widths` 參數，看起來像答案但不是：模型只有想到的時候才會給，
 * 而它很少想到。**好的預設值勝過可用的選項**——跟版型自己帶版面、而不是要模型算
 * 座標是同一個道理。
 */
{
  const grid = [
    ['編號', '功能', '說明', '優先級'],
    ['F-01', '題型支援', '單選、複選、下拉、簡答、詳答、線性量表（1–10）、日期、時間', 'P0'],
    ['F-02', '題目編輯', '新增、複製、刪除、拖曳排序；可設為必填', 'P0'],
    ['F-03', '問卷範本', '提供活動報名、時間調查、滿意度、讀者調查等 8 種範本', 'P0'],
  ];
  reset();
  const r = await host.executeTool('word_insert_table', { values: grid });
  const t = log.tables[0].table;
  const w = [0, 1, 2, 3].map((c) => t.getCell(0, c).columnWidth);

  ok('四欄都拿到寬度', w.every((x) => typeof x === 'number' && x > 0), w.join('|'));
  // 說明欄裝的是一整句話，編號欄裝的是四個字——不該一樣寬。
  ok('說明欄最寬', w[2] === Math.max(...w), w.join('|'));
  ok('編號欄最窄', w[0] === Math.min(...w), w.join('|'));
  ok('說明欄至少是編號欄的兩倍', w[2] > w[0] * 2, w.join('|'));
  // 這是畫面上最刺眼的那個：表頭折行看起來就像壞掉。
  ok('編號欄夠寬，不會把「編號」折成兩行', w[0] >= 34, `${w[0]}`);
  ok('優先級欄夠寬', w[3] >= 34, `${w[3]}`);
  // 分配完要剛好是表格的寬度，不能多也不能少。
  const sum = w.reduce((a, b) => a + b, 0);
  ok('加起來等於表格寬度', Math.abs(sum - TABLE_WIDTH_PT) < 1, `${sum} vs ${TABLE_WIDTH_PT}`);
  ok('回報有講欄寬是自動分配的',
    /自動分配/.test(String(r.content || '')), String(r.content || ''));
}

{
  // 中日韓字元佔兩格。用 .length 算的話「優先級」會跟「abc」一樣寬，欄位就窄三成。
  reset();
  await host.executeTool('word_insert_table', {
    values: [['abc', '中文字'], ['x', 'y']],
  });
  const t = log.tables[0].table;
  ok('中文表頭比同字數的英文表頭寬',
    t.getCell(0, 1).columnWidth > t.getCell(0, 0).columnWidth,
    `${t.getCell(0, 0).columnWidth} vs ${t.getCell(0, 1).columnWidth}`);
}

{
  // 模型明確指定就照做——自動分配是預設值，不是政策。
  reset();
  await host.executeTool('word_insert_table', {
    values: [['a', 'b'], ['1', '2']], widths: [100, 351],
  });
  const t = log.tables[0].table;
  ok('有給 widths 就以 widths 為準',
    t.getCell(0, 0).columnWidth === 100 && t.getCell(0, 1).columnWidth === 351);
}

{
  reset();
  await host.executeTool('word_insert_table', {
    values: [['a', 'b'], ['1', '2']], auto_width: false,
  });
  const t = log.tables[0].table;
  ok('auto_width: false 就完全不碰欄寬',
    t.getCell(0, 0).columnWidth === undefined, `${t.getCell(0, 0).columnWidth}`);
}

/**
 * 一開始就對，而不是最後再修一次。
 *
 * ── 回報的問題 ──
 * 「他其實一開始都會做得非常醜…文字是細明體，標題、段落、條列的設計不一，然後她
 * 會在最後有一個專門看排版的流程把這些都修好，但其實我很希望他當下就修好。」
 *
 * 舊做法做不到：`word_apply_theme` 是走過每一段去塗——它只修得到**已經寫好**的
 * 東西，而之後寫進去的又會變回 Word 的 zh-TW 預設（細明體）。最後要一道排版流程
 * 收尾，本身就是「格式長在段落上而不是長在樣式上」的症狀。
 *
 * 改樣式定義就從源頭解決：「內文」從此是微軟正黑體，之後每一段、每個項目、每個
 * 圖說、每個表格儲存格都第一次就對，而且會一直對。
 */
{
  // 換一份沒碰過的文件，才看得到「第一次寫入」的行為——上面那些表格測試已經把
  // 這份文件的樣式設過了，而那個鎖本身就是要測的東西之一。
  Office.context.document.url = 'C:/fresh.docx';
  reset();
  log.styles = {};
  await host.executeTool('word_insert_text', { text: '第一句話' });

  ok('寫第一個字之前就先改樣式', Object.keys(log.styles).length > 0,
    JSON.stringify(log.styles));
  // 這是整件事的重點：細明體是 Word 的 zh-TW 預設，不改樣式就永遠是它。
  ok('內文字型是微軟正黑體', log.styles['內文'] && log.styles['內文'].name === '微軟正黑體',
    JSON.stringify(log.styles['內文']));
  ok('標題 1 也改了', !!log.styles['標題 1'], JSON.stringify(log.styles['標題 1']));
  ok('標題階層是遞減的',
    log.styles['標題 1'].size > log.styles['標題 2'].size
    && log.styles['標題 2'].size > log.styles['標題 3'].size,
    [1, 2, 3].map((n) => log.styles[`標題 ${n}`].size).join(' > '));
  // 第四層如果比內文還小，讀起來就不像標題了。
  ok('標題 4 不會比內文小', log.styles['標題 4'].size > log.styles['內文'].size,
    `${log.styles['標題 4'].size} vs ${log.styles['內文'].size}`);
  // 條列用的是「清單段落」樣式——不一起改，條列就會跟內文長得不一樣。
  ok('清單段落一起改了', log.styles['清單段落'].name === '微軟正黑體');
  // 表格儲存格裡的文字是「內文」樣式的段落，所以改內文就一起涵蓋到了。
  ok('圖說也一起改了', log.styles['圖表標號'].name === '微軟正黑體');
  ok('圖說比內文小一點', log.styles['圖表標號'].size < log.styles['內文'].size);

  // 查不到的名字要安靜跳過，不能讓整段寫入失敗。
  ok('英文名查不到也不會壞', !log.styles.Normal && !log.styles['Heading 1'],
    JSON.stringify(Object.keys(log.styles)));
}

{
  // 一份文件只做一次——每次寫入都重跑是白花一趟來回。
  Office.context.document.url = 'C:/once.docx';
  reset();
  log.styles = {};
  await host.executeTool('word_insert_text', { text: 'a' });
  const first = JSON.stringify(log.styles);
  ok('新文件的第一次寫入有改樣式', Object.keys(log.styles).length > 0, first.slice(0, 80));
  log.styles = {};
  await host.executeTool('word_insert_text', { text: 'b' });
  ok('第二次寫入不會再改一次樣式', Object.keys(log.styles).length === 0,
    JSON.stringify(log.styles).slice(0, 80));

  // 但換了一份文件就要重做——SharedRuntime 會讓窗格活過換檔案。
  Office.context.document.url = 'C:/another.docx';
  log.styles = {};
  await host.executeTool('word_insert_text', { text: 'c' });
  ok('換一份文件會重新設定樣式', Object.keys(log.styles).length > 0,
    JSON.stringify(log.styles).slice(0, 80));
}

{
  // 表格預設顏色：藍色的底配黑字，表格最該做好的「看得清楚」反而最差。
  reset();
  await host.executeTool('word_insert_table', { values: [['a', 'b'], ['1', '2']] });
  const style = log.tables[0].table.styleBuiltIn;
  /**
   * 內建樣式是一場賭博：它實際畫出什麼取決於文件的佈景主題，而那是建檔的人的。
   * 表格唯一的工作就是看得清楚，這件事不值得賭——所以 TableGrid 只負責框線，
   * 顏色自己上。
   */
  ok('預設樣式只提供框線，不帶顏色', style === 'TableGrid', style);
  ok('沒有 Accent 色號', !/Accent/.test(style), style);
  // 自己畫條紋的時候，樣式自己的條紋要關掉——兩組錯開半列的條紋比哪一組都糟。
  ok('關掉內建的條紋', log.tables[0].table.styleBandedRows === false,
    String(log.tables[0].table.styleBandedRows));
}

{
  // 指定了樣式就是選了一個外觀，不該被我們蓋掉。
  reset();
  await host.executeTool('word_insert_table',
    { values: [['a'], ['1']], style: 'GridTable4_Accent1' });
  const t3 = log.tables[0].table;
  ok('指定樣式時不自己上色', t3.getCell(0, 0)._props.shadingColor === undefined,
    String(t3.getCell(0, 0)._props.shadingColor));
}

{
  reset();
  await host.executeTool('word_insert_table', { values: [['a'], ['1']], paint: false });
  const t4 = log.tables[0].table;
  ok('paint: false 也不上色', t4.getCell(0, 0)._props.shadingColor === undefined);
}

{
  reset();
  await host.executeTool('word_insert_table',
    { values: [['a']], style: 'GridTable4_Accent1' });
  ok('使用者指定顏色還是照做',
    log.tables[0].table.styleBuiltIn === 'GridTable4_Accent1');
}

{
  /**
   * 封面的文件資訊表跟內文表格要長得一樣。
   *
   * 一份文件裡的表格對「表格長什麼樣」各說各話，正是這整件事要擋掉的。而這張沒有
   * 表頭列——它是標籤／值的配對——所以條紋要從第一列開始數，也不該有黑底。
   */
  reset();
  await host.executeTool('word_insert_cover', {
    title: '個人問卷平台 產品需求文件',
    subtitle: 'Personal Survey Platform — PRD',
    meta: { 文件版本: 'v1.0', 建立日期: '2026-10-11', 文件狀態: '草稿' },
  });
  const info = log.tables[0] && log.tables[0].table;
  if (info) {
    ok('封面資訊表第一列是白的，不是黑底',
      info.getCell(0, 0)._props.shadingColor === '#FFFFFF',
      String(info.getCell(0, 0)._props.shadingColor));
    ok('封面資訊表第二列是淡灰',
      info.getCell(1, 0)._props.shadingColor === '#F2F2F2',
      String(info.getCell(1, 0)._props.shadingColor));
    ok('封面資訊表的字是黑的',
      info.getCell(0, 1).body.font.color === '#000000',
      String(info.getCell(0, 1).body.font.color));
  }
  // 「首頁文字我希望盡量都是黑色」——標題和副標都不該是藍的。
  const xml = log.ooxml.map((o) => o.xml).join('');
  ok('封面沒有藍色文字', !/w:color w:val="(1F3864|2E5C9A|4F81BD|16314F)"/.test(xml),
    (xml.match(/w:color w:val="[0-9A-Fa-f]{6}"/g) || []).slice(0, 6).join(' '));
  ok('封面有黑色文字', /w:color w:val="000000"/.test(xml));
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
