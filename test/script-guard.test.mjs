/**
 * Drive makeScriptGuard against a stand-in for Office.js.
 *
 * The cases that matter are the ones the source-text patterns cannot see:
 * computed property names, aliased objects, a mutator reached through a
 * callback. If the Proxy only caught what the regex already caught, it would
 * not be worth the layer.
 */
import fs from 'fs';

// Sliced out of the real source between the guard:start / guard:end markers, so
// this runs the shipping code rather than a copy of it. Move the markers and
// this test tells you immediately.
const src = fs.readFileSync(new URL('../src/taskpane/excelTools.js', import.meta.url), 'utf8');
const from = src.indexOf('/* ── guard:start');
const to = src.indexOf('/* ── guard:end');
if (from < 0 || to < 0) {
  console.error('找不到 guard:start / guard:end 標記，excelTools.js 的結構變了。');
  process.exit(1);
}
const mod = src.slice(from, to) + '\nexport { makeScriptGuard };\n';
const tmp = new URL('./.guard-under-test.mjs', import.meta.url);
fs.writeFileSync(tmp, mod);
const { makeScriptGuard } = await import(tmp);

class Range {
  constructor() { this.values = null; this.formulas = null; this.format = new Fmt(); }
  load() { return this; }
  getUsedRange() { return new Range(); }
  clear() { this.cleared = true; }
  delete() { this.deleted = true; }
  // Writes that carry no assignment and no obviously destructive name — the
  // shape the old denylist could not see.
  autoFill() { this.filled = true; }
  setCellProperties() { this.propsSet = true; }
  find() { return new Range(); }
}
class Fmt { constructor() { this.fill = new Fill(); } autofitColumns() {} }
class Fill { constructor() { this.color = '#fff'; } }
class Sheets {
  constructor() { this.items = [new Sheet()]; }
  getItem() { return this.items[0]; }
  add() { const s = new Sheet(); this.items.push(s); return s; }
  load() { return this; }
}
class Sheet {
  constructor() { this.name = '工作表1'; }
  getRange() { return new Range(); }
  delete() {}
  copy() { return new Sheet(); }
  findAll() { return new Range(); }
}
/** workbook.functions.sum(…) — computes and returns, changes nothing. */
class Functions { sum() { return new FunctionResult(); } vlookup() { return new FunctionResult(); } }
class FunctionResult { constructor() { this.value = 3; } load() { return this; } }
/** Office.js proxy lifetime bookkeeping, not workbook content. */
class Tracked { add() {} remove() {} }
class Workbook {
  constructor() { this.worksheets = new Sheets(); this.functions = new Functions(); }
  close() {}
  save() {}
}
class Ctx {
  constructor() { this.workbook = new Workbook(); this.trackedObjects = new Tracked(); }
  sync() { return Promise.resolve(); }
}
const FakeExcel = { run: (cb) => cb(new Ctx()) };

const cases = [
  ['plain assignment', (c) => { c.workbook.worksheets.getItem().getRange().values = [[1]]; }, true],
  ['computed property name', (c) => { const r = c.workbook.worksheets.getItem().getRange(); r['val' + 'ues'] = [[1]]; }, true],
  ['nested format assignment', (c) => { c.workbook.worksheets.getItem().getRange().format.fill.color = 'red'; }, true],
  ['computed method name', (c) => { const r = c.workbook.worksheets.getItem().getRange(); r['cle' + 'ar'](); }, true],
  ['aliased mutator', (c) => { const r = c.workbook.worksheets.getItem().getRange(); const f = r.delete.bind(r); f(); }, true],
  ['worksheet add', (c) => { c.workbook.worksheets.add(); }, true],
  ['autofit via format', (c) => { c.workbook.worksheets.getItem().getRange().format.autofitColumns(); }, true],
  ['delete operator', (c) => { const r = c.workbook.worksheets.getItem().getRange(); delete r.values; }, true],
  ['pure read', (c) => { const r = c.workbook.worksheets.getItem().getRange(); r.load('values'); return r.values; }, false],
  ['read via getUsedRange', (c) => c.workbook.worksheets.getItem().getRange().getUsedRange().load(), false],
  ['mutate a values array copy', (c) => { const v = [[1, 2]]; v[0][0] = 9; return v; }, false],

  // ── Writes that carried no assignment and were absent from the old denylist ──
  // Each of these reached the workbook with no confirmation card and no
  // snapshot, because the server's source patterns missed them too.
  ['worksheet.copy()', (c) => { c.workbook.worksheets.getItem().copy(); }, true],
  ['range.autoFill()', (c) => { c.workbook.worksheets.getItem().getRange().autoFill('B2:B100'); }, true],
  ['range.setCellProperties()', (c) => { c.workbook.worksheets.getItem().getRange().setCellProperties([[{}]]); }, true],
  ['workbook.close()', (c) => { c.workbook.close(); }, true],
  ['workbook.save()', (c) => { c.workbook.save(); }, true],

  // The door the assignment rule could not see: `.items` is a plain Array, so it
  // was handed back unwrapped along with every Office object inside it.
  ['write through collection .items', (c) => {
    c.workbook.worksheets.load('items/name');
    c.workbook.worksheets.items[0].getRange('A1').values = [[1]];
  }, true],

  // ── Reads that are not named get*, and must not start asking for permission ──
  ['read through collection .items', (c) => {
    c.workbook.worksheets.load('items/name');
    return c.workbook.worksheets.items[0].getRange('A1').load('values');
  }, false],
  ['range.find()', (c) => c.workbook.worksheets.getItem().getRange().find('營收'), false],
  ['worksheet.findAll()', (c) => c.workbook.worksheets.getItem().findAll('營收'), false],
  ['workbook.functions.sum()', (c) => c.workbook.functions.sum([[1, 2]]).value, false],
  ['trackedObjects.add()', (c) => { c.trackedObjects.add(c.workbook.worksheets.getItem().getRange()); }, false],
];

let pass = 0, fail = 0;
for (const [name, fn, shouldBlock] of cases) {
  const guard = makeScriptGuard(false);   // writes NOT approved
  let blocked = false, err = '';
  try { fn(guard.wrap(new Ctx())); } catch (e) { blocked = true; err = e.message.slice(0, 40); }
  const ok = blocked === shouldBlock;
  ok ? pass++ : fail++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${shouldBlock ? 'block' : 'allow'}  ${name}${blocked ? '  → ' + err : ''}`);
}

// The same scripts must run untouched once a person has approved the write.
let allowed = 0;
for (const [name, fn, shouldBlock] of cases) {
  if (!shouldBlock) continue;
  const guard = makeScriptGuard(true);
  try { fn(guard.wrap(new Ctx())); allowed++; } catch (e) { fail++; console.log('FAIL  approved-run threw: ' + name + ' → ' + e.message); }
}
console.log(`\napproved runs completed: ${allowed}/${cases.filter(c => c[2]).length}`);

// A callback handed to Office must come back wrapped, or it is an unguarded door.
const guard = makeScriptGuard(false);
let viaCallback = false;
try { guard.wrapNamespace(FakeExcel).run((c) => { c.workbook.worksheets.getItem().getRange().values = [[1]]; }); }
catch { viaCallback = true; }
console.log(`${viaCallback ? 'PASS' : 'FAIL'}  block  write through Excel.run callback`);
viaCallback ? pass++ : fail++;

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
