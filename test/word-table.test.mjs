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

function fakeTable(rows, cols, log) {
  const cells = new Map();
  return {
    rowCount: rows,
    styleBuiltIn: '', headerRowCount: -1, styleBandedRows: null, styleFirstColumn: null,
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
  const ctx = {
    document: {
      changeTrackingMode: 'Off',
      body: {
        insertTable: makeTable,
        insertOoxml: (xml) => log.ooxml.push({ xml, loc: 'End' }),
        insertParagraph: () => ({}),
        paragraphs: { items: [], load() {} },
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
    document: { addHandlerAsync: () => {} },
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
ok('明細列沒有被染到', [0, 1, 2].every((r) => at(r, 0, 'shadingColor') === undefined));
ok('合計列白色粗體',
  [3, 4, 5].every((r) => t.getCell(r, 5).body.font.color === '#FFFFFF' && t.getCell(r, 5).body.font.bold === true));
ok('合計列仍然靠右', at(5, 5, 'horizontalAlignment') === 'Right',
  '深底白字不該覆蓋掉對齊');

// total_rows must never eat the header, however large the number.
reset();
await host.executeTool('word_insert_table', { values: [['A', 'B'], ['1', '2']], total_rows: 99 });
const t2 = log.tables[0].table;
ok('total_rows 過大也不會吃掉表頭', t2.getCell(0, 0)._props.shadingColor === undefined,
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

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
