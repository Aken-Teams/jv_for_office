/**
 * Drive the script-undo snapshot against a stand-in for Office.js.
 *
 * excel_write_range knows the range it is about to overwrite, so its snapshot is
 * easy to believe. A script's is not: the before-image is of the whole workbook,
 * and putting it back has to deal with content written PAST the old used range,
 * sheets the script created, and sheets it deleted — none of which a plain
 * "write the old values back" would handle. That is the part worth pinning, and
 * it cannot be checked by reading the code.
 *
 * The scenario below is one script doing all four things at once, plus a sheet
 * the PERSON adds afterwards, which undo must not touch.
 */
import fs from 'fs';

const src = fs.readFileSync(new URL('../src/taskpane/excelTools.js', import.meta.url), 'utf8');
const from = src.indexOf('/* ── snapshot:start');
const to = src.indexOf('/* ── snapshot:end');
const bareAt = src.indexOf('function bareAddress(');
if (from < 0 || to < 0 || bareAt < 0) {
  console.error('找不到 snapshot:start / snapshot:end 標記，excelTools.js 的結構變了。');
  process.exit(1);
}
const mod = src.slice(bareAt, src.indexOf('\n}\n', bareAt) + 3)
  + src.slice(from, to)
  // restoreSnapshot carries its own `export` in the source; only the two
  // module-private helpers need one adding.
  + '\nexport { captureWorkbook, finishWorkbookSnapshot };\n';
const tmp = new URL('./.snapshot-under-test.mjs', import.meta.url);
fs.writeFileSync(tmp, mod);

// ── A very small Excel ────────────────────────────────────────────────────────
// Cells hold a formula string and a number format. Enough for the snapshot,
// which only ever reads and writes those two.

const colName = (i) => {
  let s = '', n = i;
  for (;;) { s = String.fromCharCode(65 + (n % 26)) + s; n = Math.floor(n / 26) - 1; if (n < 0) break; }
  return s;
};
const colIndex = (s) => [...s].reduce((n, c) => n * 26 + (c.charCodeAt(0) - 64), 0) - 1;
const parseA1 = (a) => {
  const [tl, br = tl] = a.split(':');
  const m = (r) => { const x = /^([A-Z]+)(\d+)$/.exec(r); return [Number(x[2]) - 1, colIndex(x[1])]; };
  const [r1, c1] = m(tl), [r2, c2] = m(br);
  return { r1, c1, r2, c2 };
};

class FakeRange {
  constructor(sheet, box, isNull = false) {
    this.sheet = sheet;
    this.box = box;
    this.isNullObject = isNull;
    this.address = isNull ? '' : `${sheet.name}!${colName(box.c1)}${box.r1 + 1}:${colName(box.c2)}${box.r2 + 1}`;
    this.rowCount = isNull ? 0 : box.r2 - box.r1 + 1;
    this.columnCount = isNull ? 0 : box.c2 - box.c1 + 1;
  }
  /** The fake is always loaded; load() only has to not blow up. */
  load() { return this; }
  grid(pick) {
    const out = [];
    for (let r = this.box.r1; r <= this.box.r2; r++) {
      const row = [];
      for (let c = this.box.c1; c <= this.box.c2; c++) row.push(pick(this.sheet.get(r, c)));
      out.push(row);
    }
    return out;
  }
  get formulas() { return this.grid((cell) => cell.f); }
  set formulas(v) { this.write(v, 'f'); }
  get numberFormat() { return this.grid((cell) => cell.nf); }
  set numberFormat(v) { this.write(v, 'nf'); }
  write(v, key) {
    for (let r = 0; r < v.length; r++) {
      for (let c = 0; c < v[r].length; c++) this.sheet.set(this.box.r1 + r, this.box.c1 + c, key, v[r][c]);
    }
  }
  clearContents() {
    for (let r = this.box.r1; r <= this.box.r2; r++) {
      for (let c = this.box.c1; c <= this.box.c2; c++) this.sheet.set(r, c, 'f', '');
    }
  }
}

class FakeSheet {
  constructor(name) { this.name = name; this.cells = new Map(); }
  get(r, c) { return this.cells.get(`${r},${c}`) || { f: '', nf: 'General' }; }
  set(r, c, key, value) {
    const cur = { ...this.get(r, c) };
    cur[key] = value;
    this.cells.set(`${r},${c}`, cur);
  }
  getRange(a1) { return new FakeRange(this, parseA1(a1)); }
  /** valuesOnly — the bounding box of cells that actually hold something. */
  getUsedRangeOrNullObject() {
    let r1 = Infinity, c1 = Infinity, r2 = -1, c2 = -1;
    for (const [k, v] of this.cells) {
      if (v.f === '' || v.f === null) continue;
      const [r, c] = k.split(',').map(Number);
      r1 = Math.min(r1, r); c1 = Math.min(c1, c); r2 = Math.max(r2, r); c2 = Math.max(c2, c);
    }
    return r2 < 0 ? new FakeRange(this, null, true) : new FakeRange(this, { r1, c1, r2, c2 });
  }
}

class FakeSheets {
  constructor(book) { this.book = book; }
  get items() { return this.book.sheets; }
  load() { return this; }
  getItem(name) {
    const s = this.book.sheets.find((x) => x.name === name);
    if (!s) throw new Error(`找不到工作表「${name}」`);
    return s;
  }
  add(name) { const s = new FakeSheet(name); this.book.sheets.push(s); return s; }
  delete(name) { this.book.sheets = this.book.sheets.filter((x) => x.name !== name); }
}

class FakeBook {
  constructor() { this.sheets = []; this.worksheets = new FakeSheets(this); }
}

const book = new FakeBook();
// delete() is a method on the sheet in Office.js; route it back to the collection.
FakeSheet.prototype.delete = function () { book.worksheets.delete(this.name); };

globalThis.Excel = {
  run: async (cb) => cb({ workbook: book, sync: async () => {} }),
};

const { captureWorkbook, finishWorkbookSnapshot, restoreSnapshot } = await import(tmp);

// ── The scenario ──────────────────────────────────────────────────────────────

const fill = (sheet, a1, value) => { const { r1, c1 } = parseA1(a1); sheet.set(r1, c1, 'f', value); };
const read = (sheet, a1) => { const { r1, c1 } = parseA1(a1); return sheet.get(r1, c1).f; };
const names = () => book.sheets.map((s) => s.name);

const data = book.worksheets.add('損益表');
fill(data, 'A1', '科目'); fill(data, 'B1', '金額');
fill(data, 'A2', '營收'); fill(data, 'B2', '1000');
const scratch = book.worksheets.add('原始資料');
fill(scratch, 'A1', '不要弄丟我');

const before = { b2: read(data, 'B2'), a1: read(scratch, 'A1') };

const snap = await captureWorkbook();

// A script that overwrites, writes past the old used range, adds a sheet and
// deletes one — all in a single approved call.
fill(data, 'B2', '=SUM(9999)');
fill(data, 'D9', '寫到舊使用範圍以外');
const report = book.worksheets.add('報表');
fill(report, 'A1', 'AI 產生的');
book.worksheets.delete('原始資料');

await finishWorkbookSnapshot(snap);

// …and only afterwards, the person adds a sheet of their own. Undo must leave
// this alone: it is absent from `before` AND from `after`.
const mine = book.worksheets.add('我的筆記');
fill(mine, 'A1', '手寫的');

const result = await restoreSnapshot(snap);

// ── What must be true ─────────────────────────────────────────────────────────

const checks = [
  ['覆寫的儲存格回到原值', read(book.worksheets.getItem('損益表'), 'B2') === before.b2],
  ['寫到舊範圍以外的內容被清掉', read(book.worksheets.getItem('損益表'), 'D9') === ''],
  ['AI 新增的工作表被移除', !names().includes('報表')],
  ['AI 刪掉的工作表被重建', names().includes('原始資料')],
  ['重建的工作表帶回內容', read(book.worksheets.getItem('原始資料'), 'A1') === before.a1],
  ['使用者自己加的工作表沒有被動到', names().includes('我的筆記')],
  ['使用者自己加的內容還在', read(book.worksheets.getItem('我的筆記'), 'A1') === '手寫的'],
  ['回報了移除的工作表', result.deleted.join() === '報表'],
  ['回報了重建的工作表', result.recreated.join() === '原始資料'],
];

// A workbook past the ceiling must report "no snapshot" rather than a partial
// one — the pane keys the 「拿不回來」 warning off exactly this null.
const big = new FakeSheet('大表');
for (let r = 0; r < 500; r++) for (let c = 0; c < 500; c++) big.set(r, c, 'f', 'x');
book.sheets.push(big);
checks.push(['超過上限時回傳 null，不是半份快照', (await captureWorkbook()) === null]);

let pass = 0, fail = 0;
for (const [name, ok] of checks) {
  ok ? pass++ : fail++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}`);
}
console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
