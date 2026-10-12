/**
 * 最後那一關：整份版面的機械性檢查。
 *
 * ── 這支測試在守什麼 ──
 * 「有些時候會變成他表格跨頁、多一整個空白頁」。這些毛病的共同點是：寫的時候每一步
 * 都對，是**累積**出來的，所以只有在最後退一步看整份才看得到。
 *
 * 而它最該守住的其實是**不要做太多**：使用者明講了「新章節新一頁並沒有真的很需要，
 * 有些適合有些不適合」。一支會自作主張把章節切開的工具，造成的工比它省的多。
 */
import { executeTool } from '../src/taskpane/hosts/word.js';

let pass = 0;
let fail = 0;
function ok(name, cond, detail) {
  if (cond) { pass++; console.log(`PASS  ${name}`); } else {
    fail++; console.log(`FAIL  ${name}${detail === undefined ? '' : `  → ${detail}`}`);
  }
}

function fakeWord(texts, tableSpecs = []) {
  const log = { deleted: [], headerSet: [], breaks: [] };
  const paragraphs = texts.map((t, i) => ({
    text: t,
    style: '', styleBuiltIn: '',
    delete: () => log.deleted.push(i),
    insertBreak: (kind, where) => log.breaks.push({ i, kind, where }),
    load() {},
  }));
  const tables = tableSpecs.map((spec, i) => {
    const t = {
      rowCount: spec.rows,
      values: spec.values,
      _header: spec.headerRowCount ?? 0,
      get headerRowCount() { return this._header; },
      set headerRowCount(v) { this._header = v; log.headerSet.push({ i, v }); },
      load() {},
    };
    return t;
  });
  const ctx = {
    document: {
      changeTrackingMode: 'Off',
      getStyles: () => ({ getByNameOrNullObject: () => ({ isNullObject: true, font: {}, load() {} }) }),
      body: {
        paragraphs: { items: paragraphs, load() {} },
        tables: { items: tables, load() {} },
        getOoxml: () => ({ value: '<xml/>' }),
        insertParagraph: () => ({ insertOoxml: () => {}, styleBuiltIn: '', insertBreak: () => {} }),
        insertOoxml: () => {},
        load() {},
      },
      load() {},
    },
    sync: async () => {},
    load() {},
  };
  global.Word = { run: async (fn) => fn(ctx), InsertLocation: {} };
  global.Office = {
    context: {
      requirements: { isSetSupported: () => true },
      document: { url: `r-${Math.random()}.docx`, settings: { get: () => null, set() {}, saveAsync() {} } },
    },
  };
  return log;
}

// ── 連續空段落：一整頁空白的真正來源 ──────────────────────────────────────
{
  const log = fakeWord(['標題', '內文一', '', '', '', '', '', '內文二']);
  const res = await executeTool('word_review_layout', {});
  ok('查得出連續空段落', /連續空段落/.test(String(res.content)), String(res.content).slice(0, 120));
  // 五個空的留一個、刪四個——段落之間要留呼吸，但不是留五段。
  ok('留一段、其餘刪掉', log.deleted.length === 4, `刪了 ${log.deleted.length} 段`);
  /**
   * 由後往前刪。
   *
   * 由前往後刪的話，刪掉第 2 段之後原本的第 3 段就變成第 2 段，後面每一個索引都
   * 錯位——而錯位刪掉的是**有內容的段落**，使用者不會立刻發現。
   */
  const order = log.deleted.slice();
  ok('由後往前刪，索引才不會位移',
    order.every((v, k) => k === 0 || v < order[k - 1]), order.join(' '));
}

// ── 結尾的空白全部清掉 ────────────────────────────────────────────────────
{
  const log = fakeWord(['內文', '', '']);
  await executeTool('word_review_layout', {});
  ok('結尾的空白不留', log.deleted.length === 2, `刪了 ${log.deleted.length} 段`);
}

// ── 乾淨的文件不要亂動 ────────────────────────────────────────────────────
{
  const log = fakeWord(['標題', '', '內文一', '', '內文二']);
  const res = await executeTool('word_review_layout', {});
  ok('段落之間的單一空行不算問題', log.deleted.length === 0, `刪了 ${log.deleted.length} 段`);
  ok('沒問題就說沒問題', /沒有查到/.test(String(res.content)), String(res.content).slice(0, 80));
}

// ── 長表格的表頭 ──────────────────────────────────────────────────────────
{
  const header = ['編號', '模組', '功能說明', '優先級'];
  const body = Array.from({ length: 14 }, (_, i) => [`F-${i}`, '結果', '即時統計圖表', 'P0']);
  const log = fakeWord(['內文'], [{ rows: 15, values: [header, ...body] }]);
  const res = await executeTool('word_review_layout', {});
  ok('長表格的表頭改成跨頁重複',
    log.headerSet.length === 1 && log.headerSet[0].v === 1, JSON.stringify(log.headerSet));
  ok('有講出來', /表頭/.test(String(res.content)));

  /**
   * 做不到的要照實說。
   *
   * 「整列不要被切成兩頁」是 OOXML 的 `w:cantSplit`，Word JS API 沒有對應屬性。
   * 只設表頭重複卻回報成「表格跨頁處理好了」，會讓使用者以為不用再檢查。
   */
  ok('把做不到的事講清楚',
    /沒有對應的設定|做不到/.test(String(res.content)), String(res.content).slice(-200));
}

// ── 第一列不像表頭就不要動 ────────────────────────────────────────────────
{
  // 第一列是資料，而且有一格是空的——把它當表頭重複出來比不重複更難看。
  const rows = Array.from({ length: 12 }, () => ['2026-01-01', '', '這是一段很長的說明文字，不會是表頭']);
  const log = fakeWord(['內文'], [{ rows: 12, values: rows }]);
  const res = await executeTool('word_review_layout', {});
  ok('第一列不像表頭就不設', log.headerSet.length === 0, JSON.stringify(log.headerSet));
  ok('但會提醒它會跨頁', /看起來不是表頭/.test(String(res.content)), String(res.content).slice(0, 160));
}

// ── 短表格不要管 ──────────────────────────────────────────────────────────
{
  const log = fakeWord(['內文'], [{ rows: 4, values: [['項目', '內容'], ['a', 'b'], ['c', 'd'], ['e', 'f']] }]);
  await executeTool('word_review_layout', {});
  ok('短表格不會被動到', log.headerSet.length === 0);
}

// ── 刻意不做的事 ──────────────────────────────────────────────────────────
{
  /**
   * 「新章節新一頁並沒有真的很需要，因為有時候真的要看，有些適合有些不適合」。
   *
   * 所以這支一個分頁符號都不可以插。這條斷言比任何一條「有沒有修好」更重要——
   * 自作主張切開一份文件，使用者要花更多力氣改回來。
   */
  const log = fakeWord(['第一章', '內文', '第二章', '內文', '第三章']);
  const res = await executeTool('word_review_layout', {});
  ok('不會自己插分頁', log.breaks.length === 0, JSON.stringify(log.breaks));
  ok('而且會說明為什麼沒做', /編輯判斷/.test(String(res.content)), String(res.content).slice(-120));
}

// ── 只看不改 ──────────────────────────────────────────────────────────────
{
  const log = fakeWord(['內文', '', '', '', '', '內文二']);
  const res = await executeTool('word_review_layout', { fix: false });
  ok('fix: false 不動文件', log.deleted.length === 0, `刪了 ${log.deleted.length} 段`);
  ok('但還是會報告查到什麼', /連續空段落/.test(String(res.content)));
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
