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
  const log = { deleted: [], headerSet: [], breaks: [], rewrites: [] };
  /**
   * 字串就是普通段落；要測標題就給 {text, style}。
   *
   * 舊的假物件把 styleBuiltIn 寫死成空字串，於是 headingLevel 一律回 0——「不會自己
   * 插分頁」那條斷言之所以永遠綠的，不是因為程式不插，是因為**這份假文件裡根本沒有
   * 標題**。斷言測不到它宣稱在測的東西，比沒有斷言更糟。
   */
  const paragraphs = texts.map((spec, i) => {
    const t = typeof spec === 'string' ? spec : spec.text;
    const node = {
      text: t,
      style: '', styleBuiltIn: typeof spec === 'string' ? '' : (spec.style || ''),
      delete: () => log.deleted.push(i),
      insertBreak: (kind, where) => log.breaks.push({ i, kind, where }),
      insertText: (txt) => { node.text = txt; log.rewrites.push({ i, txt }); return node; },
      load() {},
    };
    return node;
  });
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

// ── 章節編號與大章節換頁 ──────────────────────────────────────────────────
{
  /**
   * 政策換了，而且是使用者自己換的。
   *
   * 原本這裡釘的是「一個分頁符號都不可以插」，理由是「有些適合有些不適合」。
   * 後來看到真實產出才發現問題不在「要不要換頁」，在**沒有東西分得出大章節和
   * 小節**——兩級標題都只有一條線，份量太接近。
   *
   * 所以門檻不是「所有標題」，是「這是不是一個大章節」：編得到號的才換頁。
   * 原本那條顧慮（三行的小節獨佔一頁很空）由這個門檻擋掉，不是靠完全不做。
   */
  const H1 = (text) => ({ text, style: 'Heading 1' });
  const H2 = (text) => ({ text, style: 'Heading 2' });
  const log = fakeWord([H1('產品概述'), '內文', H2('背景'), '內文', H1('目標使用者'), '內文', H1('功能需求')]);
  await executeTool('word_review_layout', {});
  // 三個大章節，第一個不換頁——它緊接在目錄後面，目錄自己已經把頁收掉了。
  ok('只有第二個之後的大章節換頁', log.breaks.length === 2, JSON.stringify(log.breaks));
  ok('換的是分頁', log.breaks.every((b) => b.kind === 'Page' && b.where === 'Before'),
    JSON.stringify(log.breaks));
  const got = log.rewrites.map((r) => r.txt);
  ok('大章節編成一、二、三', got.includes('一、產品概述') && got.includes('二、目標使用者')
    && got.includes('三、功能需求'), JSON.stringify(got));
  ok('小節編成 1.1（跟著所屬的大章節走）', got.includes('1.1 背景'), JSON.stringify(got));

  // 再跑一次不該變成「一、一、產品概述」——這個工具本來就會被跑很多次。
  const again = fakeWord([H1('一、產品概述'), '內文', H2('1.1 背景'), H1('二、目標使用者')]);
  await executeTool('word_review_layout', {});
  ok('重複跑不會疊加編號', again.rewrites.length === 0, JSON.stringify(again.rewrites));
}

// ── 只看不改 ──────────────────────────────────────────────────────────────
{
  const log = fakeWord(['內文', '', '', '', '', '內文二']);
  const res = await executeTool('word_review_layout', { fix: false });
  ok('fix: false 不動文件', log.deleted.length === 0, `刪了 ${log.deleted.length} 段`);
  ok('但還是會報告查到什麼', /連續空段落/.test(String(res.content)));
}

console.log(`\n${pass} passed, ${fail} failed`);
// ── 收尾機制：不靠使用者下對 prompt ─────────────────────────────────────────
{
  /**
   * 「我擔心都看使用的 prompt，如果他沒下，產品品質就不穩」。
   *
   * 章節編號和大章節換頁都住在 word_review_layout 裡，而它要靠模型自己想到要呼叫。
   * 這幾條釘的是門檻：動過結構才補、模型跑過就不補、文件太小也不補。寧可少做，
   * 也不要在使用者只改了一句話的時候重寫他的標題。
   */
  const { finishRun } = await import('../src/taskpane/hosts/word.js');
  const H1 = (text) => ({ text, style: 'Heading 1' });

  // 沒動過結構 → 不補
  fakeWord([H1('產品概述'), '內文', H1('目標使用者')]);
  ok('沒動過結構就不補跑', !(await finishRun()));

  // 動過結構、模型沒跑 → 補
  const a = fakeWord([H1('產品概述'), '內文', H1('目標使用者')]);
  await executeTool('word_insert_cover', { title: '測試' }).catch(() => {});
  fakeWord([H1('產品概述'), '內文', H1('目標使用者')], [], a);
  const note = await finishRun();
  ok('動過結構就會補跑並說出來', typeof note === 'string');

  // 模型自己跑過 → 不重複
  fakeWord([H1('產品概述'), '內文', H1('目標使用者')]);
  await executeTool('word_insert_table', { rows: [['a', 'b']] }).catch(() => {});
  await executeTool('word_review_layout', {});
  ok('模型自己跑過就不再補一次', !(await finishRun()));

  // 只有一個大章節 → 不值得編號
  fakeWord([H1('備忘'), '內文']);
  await executeTool('word_insert_table', { rows: [['a', 'b']] }).catch(() => {});
  ok('只有一個大章節就不補', !(await finishRun()));
}

process.exit(fail ? 1 : 0);
