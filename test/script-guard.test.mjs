/**
 * Drive makeScriptGuard against a stand-in for Office.js.
 *
 * The cases that matter are the ones the source-text patterns cannot see:
 * computed property names, aliased objects, a mutator reached through a
 * callback. If the Proxy only caught what the regex already caught, it would
 * not be worth the layer.
 */
import fs from 'fs';

const src = fs.readFileSync(new URL('../src/taskpane/excelTools.js', import.meta.url), 'utf8');
const from = src.indexOf('const MUTATING_METHODS');
const to = src.indexOf('async function runScript(');
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
}
class Fmt { constructor() { this.fill = new Fill(); } autofitColumns() {} }
class Fill { constructor() { this.color = '#fff'; } }
class Sheets {
  constructor() { this.items = [new Sheet()]; }
  getItem() { return this.items[0]; }
  add() { const s = new Sheet(); this.items.push(s); return s; }
}
class Sheet { getRange() { return new Range(); } delete() {} }
class Workbook { constructor() { this.worksheets = new Sheets(); } }
class Ctx {
  constructor() { this.workbook = new Workbook(); }
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
