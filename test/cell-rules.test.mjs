/**
 * 下拉選單、定義名稱、儲存格註解——做表單的三件事。
 *
 * ── 為什麼值得是三支工具而不是叫 AI 寫 script ──
 * 這三件事本來用 `excel_run_script` 就做得到，所以看起來像便利功能。不是。
 * 「把這欄做成下拉選單」是試算表最常被要求的事情之一——報價單、申請表之所以是
 * 「表單」而不是「一堆格子」，大半就靠它——而走 script 代表每次都要現寫程式、
 * 每次都跳一張確認卡片。
 *
 * 更要緊的是**模型找不到它不知道存在的能力**：沒有具名工具，它根本不會把
 * 「不要讓人亂打」辨識成資料驗證這件事。
 *
 * ── 這裡釘的三件事 ──
 *   1. 選項有兩種給法，而且差很多。寫死的有 Excel 自己的 255 字上限，會變動的
 *      要指向範圍——給錯會在使用者加第十個選項的時候才爆。
 *   2. 擋下來要說人話。Excel 預設的「這個值不符合儲存格定義的資料驗證限制」沒有
 *      告訴任何人到底能填什麼。
 *   3. 拿掉規則、刪名稱、刪註解要算破壞性——那是在移除別人刻意放的東西。
 */
import fs from 'fs';

let pass = 0, fail = 0;
const ok = (name, cond, detail) => {
  cond ? pass++ : fail++;
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}${cond ? '' : '  → ' + (detail || '')}`);
};

/** 每一次對 Office 的寫入，依序記下來。 */
let WROTE = [];

function fakeRange(address) {
  const r = {
    address,
    dataValidation: {
      set rule(v) { WROTE.push({ what: 'rule', address, value: v }); r._rule = v; },
      get rule() { return r._rule; },
      set errorAlert(v) { WROTE.push({ what: 'errorAlert', address, value: v }); },
      set prompt(v) { WROTE.push({ what: 'prompt', address, value: v }); },
      clear() { WROTE.push({ what: 'clearValidation', address }); },
    },
    load() {},
  };
  return r;
}

let SHEETS = {};
let NAMES = {};
let COMMENTS = {};

const fakeSheet = (name) => ({
  name, isNullObject: false, load() {},
  getRange: (a1) => fakeRange(a1),
});

const ctx = {
  workbook: {
    worksheets: {
      getItemOrNullObject: (n) => SHEETS[n] || { isNullObject: true, load() {} },
      load() {},
    },
    names: {
      get items() {
        return Object.entries(NAMES).map(([name, v]) => ({ name, ...v }));
      },
      load() {},
      getItemOrNullObject: (n) => (NAMES[n]
        ? { isNullObject: false, load() {}, delete() { delete NAMES[n]; } }
        : { isNullObject: true, load() {} }),
      add(name, range, comment) {
        NAMES[name] = { formula: `=${range.address}`, comment: comment || '' };
        WROTE.push({ what: 'name', name, formula: NAMES[name].formula });
      },
    },
    comments: {
      getItemByCellOrNullObject: (at) => (COMMENTS[at]
        ? {
          isNullObject: false, load() {},
          delete() { delete COMMENTS[at]; WROTE.push({ what: 'deleteComment', at }); },
          replies: { add(t) { COMMENTS[at].replies.push(t); WROTE.push({ what: 'reply', at, text: t }); } },
        }
        : { isNullObject: true, load() {} }),
      add(at, text) {
        COMMENTS[at] = { text, replies: [] };
        WROTE.push({ what: 'comment', at, text });
      },
    },
  },
  sync: async () => {},
};

globalThis.Excel = { run: async (fn) => fn(ctx) };
let API_LEVEL = 99;
globalThis.Office = {
  context: {
    requirements: { isSetSupported: (_set, v) => parseFloat(v) <= API_LEVEL },
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
const reset = () => { WROTE = []; SHEETS = { 報價單: fakeSheet('報價單'), 設定: fakeSheet('設定') }; NAMES = {}; COMMENTS = {}; };
const call = (tool, args) => host.executeTool(tool, args);
const said = (r) => String(r.content || r.error || '');

// ── 下拉選單 ──────────────────────────────────────────────────────────────────
{
  reset();
  const r = await call('excel_data_validation', {
    sheet: '報價單', range: 'D2:D200',
    options: ['現金', '月結30天', '月結60天'],
    error_message: '請從下拉選單選付款條件。',
  });
  ok('下拉選單做得出來', r.ok, said(r));
  const rule = WROTE.find((w) => w.what === 'rule');
  ok('寫的是 list 規則', !!rule && !!rule.value.list, JSON.stringify(rule));
  ok('選項用逗號串起來', rule.value.list.source === '現金,月結30天,月結60天', rule.value.list.source);
  ok('預設會顯示下拉箭頭', rule.value.list.inCellDropDown === true);
  // Excel 預設的錯誤訊息沒告訴任何人能填什麼，所以這條要能寫進去。
  ok('錯誤訊息有寫進去', WROTE.some((w) => w.what === 'errorAlert' && /下拉選單/.test(w.value.message)));
  ok('錯誤樣式預設是 Stop', WROTE.find((w) => w.what === 'errorAlert').value.style === 'Stop');
}

{
  // 會變動的選項要指向範圍：改那一欄，所有下拉選單一起變。
  reset();
  const r = await call('excel_data_validation', {
    sheet: '報價單', range: 'B2:B200', source: '設定!$A$1:$A$9', options: ['會被忽略'],
  });
  ok('有 source 就以 source 為準',
    WROTE.find((w) => w.what === 'rule').value.list.source === '設定!$A$1:$A$9', said(r));
}

{
  // 255 是 Excel 自己的上限，不是我們定的——要在寫下去之前就講，不然會寫一半失敗。
  reset();
  const r = await call('excel_data_validation', {
    sheet: '報價單', range: 'D2', options: Array.from({ length: 40 }, (_, i) => `選項${i}字字字字`),
  });
  ok('超過 255 字會先擋下來', /255/.test(said(r)), said(r));
  ok('而且一個字都還沒寫進去', WROTE.length === 0, JSON.stringify(WROTE));
  ok('有講替代做法是改用 source', /source/.test(said(r)), said(r));
}

{
  reset();
  const r = await call('excel_data_validation', {
    sheet: '報價單', range: 'E2:E200', rule: 'wholeNumber', operator: 'Between', min: 1, max: 9999,
  });
  const rule = WROTE.find((w) => w.what === 'rule');
  ok('數值範圍規則', !!rule && rule.value.wholeNumber, said(r));
  ok('上下界都帶進去', rule.value.wholeNumber.formula1 === 1 && rule.value.wholeNumber.formula2 === 9999);
}

{
  reset();
  const r = await call('excel_data_validation',
    { sheet: '報價單', range: 'E2', rule: 'wholeNumber', operator: 'Between', min: 1 });
  ok('Between 少了 max 會說清楚', !r.ok || /max/.test(said(r)), said(r));
  ok('沒有寫進去半套規則', WROTE.length === 0);
}

{
  reset();
  const r = await call('excel_data_validation', { sheet: '報價單', range: 'D2:D200', rule: 'clear' });
  ok('清除驗證', WROTE.some((w) => w.what === 'clearValidation'), said(r));
}

{
  // 舊版 Office 沒有這個 API——要說得出替代做法，不要只丟一句不支援。
  reset();
  API_LEVEL = 1.7;
  const r = await call('excel_data_validation', { sheet: '報價單', range: 'D2', options: ['a'] });
  ok('舊版 Office 講得出版本需求', /1\.8/.test(said(r)), said(r));
  ok('而且告訴使用者可以自己設', /資料驗證/.test(said(r)), said(r));
  API_LEVEL = 99;
}

// ── 定義名稱 ──────────────────────────────────────────────────────────────────
{
  reset();
  const r = await call('excel_define_name', { sheet: '報價單', name: '單價表', range: 'C2:C40' });
  ok('定義名稱做得出來', r.ok && !!NAMES['單價表'], said(r));
  // 這才是它存在的理由：插入欄列之後公式還是對的。
  ok('回報有講為什麼值得用', /插入欄列|不會跑掉/.test(said(r)), said(r));
}

{
  // 「把單價表改成指到 C2:C60」是最常見的第二次呼叫，不該失敗。
  reset();
  await call('excel_define_name', { sheet: '報價單', name: '單價表', range: 'C2:C40' });
  const r = await call('excel_define_name', { sheet: '報價單', name: '單價表', range: 'C2:C60' });
  ok('同名會直接更新，不是報錯', r.ok && /更新/.test(said(r)), said(r));
  ok('指到新的範圍', NAMES['單價表'].formula === '=C2:C60', JSON.stringify(NAMES));
}

{
  reset();
  await call('excel_define_name', { sheet: '報價單', name: '單價表', range: 'C2:C40' });
  const r = await call('excel_define_name', { op: 'list' });
  ok('列得出來', /單價表/.test(said(r)), said(r));
  const gone = await call('excel_define_name', { op: 'delete', name: '不存在的' });
  ok('刪不存在的會說清楚', /沒有叫/.test(said(gone)), said(gone));
}

// ── 註解 ──────────────────────────────────────────────────────────────────────
{
  reset();
  const r = await call('excel_cell_comment', { sheet: '報價單', cell: 'F12', text: '這格是手動調的' });
  ok('加得上註解', r.ok && !!COMMENTS['報價單!F12'], said(r));

  const again = await call('excel_cell_comment', { sheet: '報價單', cell: 'F12', text: '再補一句' });
  // Excel 自己的模型就是這樣：同一格的第二則是回覆，不是另一個註解。
  ok('同一格再加會變成回覆', /回覆/.test(said(again)), said(again));
  ok('原本那則還在', COMMENTS['報價單!F12'].text === '這格是手動調的');

  const del = await call('excel_cell_comment', { op: 'delete', sheet: '報價單', cell: 'F12' });
  ok('刪得掉', del.ok && !COMMENTS['報價單!F12'], said(del));
}

{
  reset();
  const r = await call('excel_cell_comment', { sheet: '不存在', cell: 'A1', text: 'x' });
  ok('工作表不存在會說清楚', /找不到工作表/.test(said(r)), said(r));
}

// ── 哪些要先問過使用者 ────────────────────────────────────────────────────────
/**
 * 設規則不會拿走任何東西，清掉規則會——那是在移除別人刻意放的防呆。刪定義名稱更
 * 嚴重：每一條用到它的公式會同時壞掉，而且壞得很安靜。
 */
const serverDir = new URL('../../ai-agents-office/server/src/', import.meta.url);
if (fs.existsSync(serverDir)) {
  const spec = await import(new URL('services/excelToolSpec.ts', serverDir).href).catch(() => null);
  if (spec) {
    const d = (n, a) => spec.isDestructiveTool(n, a);
    ok('設下拉選單不用先問', d('excel_data_validation', { rule: 'list' }) === false);
    ok('清掉驗證要先問', d('excel_data_validation', { rule: 'clear' }) === true);
    ok('加定義名稱不用先問', d('excel_define_name', { op: 'add' }) === false);
    ok('刪定義名稱要先問', d('excel_define_name', { op: 'delete' }) === true);
    ok('加註解不用先問', d('excel_cell_comment', { op: 'add' }) === false);
    ok('刪註解要先問', d('excel_cell_comment', { op: 'delete' }) === true);

    const names = spec.EXCEL_TOOLS.map((t) => t.name);
    for (const t of ['excel_data_validation', 'excel_define_name', 'excel_cell_comment']) {
      ok(`${t} 有在伺服器的工具清單裡`, names.includes(t));
    }
    // 兩種給法的差別要寫在工具說明裡，不然模型只會用寫死的那種。
    const dv = spec.EXCEL_TOOLS.find((t) => t.name === 'excel_data_validation');
    ok('說明有講 255 字上限', /255/.test(dv.description));
    ok('說明有講會變動的要用 source', /source/.test(dv.description));
    ok('說明有叫它寫錯誤訊息', /error_message/.test(dv.description));
  }
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
