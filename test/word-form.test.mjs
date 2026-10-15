/**
 * 表單，不是報表。
 *
 * ── 這支測試的由來 ──
 * 「企業內部那種履歷文件他做得不夠好……表格還超大，都沒有特別排版設計過，因為
 * 傳統企業履歷都是用比較特殊和完整的設計」。截圖裡的員工履歷表是這樣壞的：
 *
 *   · 每一節一張黑色表頭的表，五節就五張表，整份文件攤成五頁
 *   · 列高是填寫空間的兩倍，因為沒有人規定過列要多高
 *   · 「電子郵件」那一列只湊得出一組欄位，旁邊就留了兩個沒有標題的空格子
 *
 * 最後一條是最糟的：空格子在表單上讀起來像是有人忘了寫欄位名稱。
 *
 * 所以這裡測的是「印出來是不是一張表單」，而不是函式有沒有被呼叫到：
 * 合併有沒有發生、標籤欄有多寬、列高有沒有講、區段帶是不是整列。
 */
import { executeTool } from '../src/taskpane/hosts/word.js';

let pass = 0;
let fail = 0;
function ok(name, cond, detail) {
  if (cond) { pass++; console.log(`PASS  ${name}`); } else {
    fail++; console.log(`FAIL  ${name}${detail === undefined ? '' : `  → ${detail}`}`);
  }
}

const TABLE_WIDTH_PT = 451;

function fakeCell(r, c, log) {
  const props = {};
  const cell = {
    body: {
      font: {},
      paragraphs: { getFirst: () => ({ set alignment(v) { props.align = v; } }) },
      insertParagraph: (text, loc) => { log.extraParas.push({ r, c, text, loc }); return {}; },
      insertOoxml: () => {},
    },
    setCellPadding: (loc, pt) => log.padding.push({ r, c, loc, pt }),
    _props: props,
  };
  for (const name of ['verticalAlignment', 'columnWidth', 'shadingColor', 'value']) {
    Object.defineProperty(cell, name, { set: (v) => { props[name] = v; }, get: () => props[name] });
  }
  Object.defineProperty(cell.body, 'font', {
    value: new Proxy({}, { set: (t, k, v) => { props[`font.${String(k)}`] = v; return true; } }),
  });
  return cell;
}

/**
 * 會合併的假表格。
 *
 * 真的 Word 合併之後，那一列的格子會變少，後面的欄位索引整個往前移。假物件要是
 * 不管合併、每個索引都照樣給一個格子，那「合併之後還去塗原本的格子」這種錯就測
 * 不出來——而那正是真的 Word 會丟例外的地方。
 */
function fakeTable(rowCount, cols, log, opts = {}) {
  const cells = new Map();
  const merges = [];
  return {
    rowCount,
    styleBuiltIn: '', headerRowCount: -1, styleBandedRows: null, styleFirstColumn: null,
    width: undefined,
    load(what) { if (String(what).includes('width')) this.width = TABLE_WIDTH_PT; },
    rows: {
      items: Array.from({ length: rowCount }, (_, i) => {
        const row = {};
        Object.defineProperty(row, 'preferredHeight', {
          set: (v) => { log.heights.push({ row: i, pt: v }); },
          get: () => undefined,
        });
        return row;
      }),
      load() {},
    },
    mergeCells(r1, c1, r2, c2) {
      if (opts.noMerge) throw new Error('mergeCells is not available');
      merges.push({ row: r1, from: c1, to: c2 });
      log.merges.push({ row: r1, from: c1, to: c2 });
    },
    getCell(r, c) {
      // 合併過的列，剩下的格子數會變少。
      const used = merges.filter((m) => m.row === r)
        .reduce((a, m) => a + (m.to - m.from), 0);
      if (c >= cols - used) throw new Error(`格子 ${r},${c} 已經被合併掉了`);
      const key = `${r},${c}`;
      if (!cells.has(key)) cells.set(key, fakeCell(r, c, log));
      return cells.get(key);
    },
    getBorder(location) { const b = { location }; log.borders.push(b); return b; },
    insertParagraph: (text, loc) => {
      log.paras.push({ text, where: `table:${loc}` });
      return { styleBuiltIn: '', alignment: '', font: {} };
    },
    _cells: cells,
    _merges: merges,
  };
}

function fakeWord(opts = {}) {
  const log = {
    tables: [], padding: [], borders: [], merges: [], heights: [], extraParas: [],
    paras: [], deleted: [], syncs: 0,
  };
  const makeTable = (r, c, loc, values) => {
    const t = fakeTable(r, c, log, opts);
    log.tables.push({ rows: r, cols: c, loc, values, table: t });
    return t;
  };
  const style = () => ({ isNullObject: true, font: {}, load() {} });
  const ctx = {
    document: {
      changeTrackingMode: 'Off',
      getStyles: () => ({ getByNameOrNullObject: style }),
      body: {
        insertTable: makeTable,
        insertOoxml: () => {},
        insertParagraph: (text) => {
          log.paras.push({ text, where: 'body' });
          return {
            insertOoxml: () => {},
            insertTable: makeTable,
            styleBuiltIn: '',
            alignment: '',
            font: {},
            insertBreak: () => {},
          };
        },
        paragraphs: {
          get items() {
            const made = log.paras.map((q) => ({
              text: String(q.text ?? ''),
              insertTable: makeTable,
              insertOoxml: () => {},
              insertParagraph: () => ({ styleBuiltIn: '' }),
              delete: () => log.deleted.push(`para:${q.text}`),
            }));
            return [...made, ...Array.from({ length: 4 }, () => ({
              text: '',
              insertTable: makeTable,
              insertOoxml: () => {},
              insertParagraph: () => ({ styleBuiltIn: '' }),
              delete: () => log.deleted.push('para:'),
            }))];
          },
          load() {},
        },
        getOoxml: () => ({ value: '<xml/>' }),
        /**
         * 續接時會去找「文件裡最後一張表」把它拆掉。假物件要記得自己建過哪些表，
         * 不然那段會丟例外、被 catch 吃掉，然後測試看到的是兩張表而不是一張。
         */
        tables: {
          get items() { return log.tables.map(() => ({ delete: () => log.deleted.push('table') })); },
          load() {},
        },
        load() {},
      },
    },
    sync: async () => { log.syncs += 1; },
    load() {},
  };
  global.Word = {
    run: async (fn) => fn(ctx),
    InsertLocation: { start: 'Start', end: 'End', before: 'Before', after: 'After', replace: 'Replace' },
  };
  global.Office = {
    context: {
      requirements: { isSetSupported: () => true },
      document: { url: `form-${Math.random()}.docx`, settings: { get: () => null, set() {}, saveAsync() {} } },
    },
  };
  return log;
}

// ── 一張真的履歷表 ──────────────────────────────────────────────────────────
{
  const log = fakeWord();
  const res = await executeTool('word_insert_form', {
    title: '員工履歷表',
    fields: [
      { section: '基本資料' },
      { label: '中文姓名' },
      { label: '英文姓名' },
      { label: '性別' },
      { label: '出生日期', value: '　　年　　月　　日' },
      // 奇數個——最後這一欄要吃掉剩下的寬度，不可以留空格子。
      { label: '電子郵件' },
      { label: '戶籍地址', span: true },
      { section: '自傳' },
      { label: '自傳內容', lines: 8 },
    ],
  });
  ok('有回報', res && res.ok && String(res.content).includes('表單'),
    JSON.stringify(res).slice(0, 120));

  const t = log.tables[0];
  ok('只插了一張表', log.tables.length === 1, `${log.tables.length} 張`);
  ok('四欄（兩組標籤＋填寫）', t && t.cols === 4, t && t.cols);

  /**
   * 列數要剛好。
   *
   * 基本資料帶、姓名列、性別/生日列、電子郵件列、戶籍地址列、自傳帶、自傳列
   * ——七列。多一列代表有一列是空的，少一列代表有欄位被吃掉了。
   */
  ok('七列', t && t.rows === 7, t && t.rows);

  // 區段帶整列合併。
  const sect = log.merges.filter((m) => m.from === 0 && m.to === 3);
  ok('兩個區段帶都整列合併', sect.length === 2, JSON.stringify(log.merges));

  // 奇數欄位：電子郵件的填寫格要延伸到最後一欄，不是留兩個空格子。
  const stretched = log.merges.filter((m) => m.from === 1 && m.to === 3);
  ok('湊不滿的那一列，填寫格吃掉剩下的寬度', stretched.length >= 1,
    JSON.stringify(log.merges));

  // 戶籍地址 span 也一樣。
  ok('span 的欄位佔滿整列', stretched.length === 3, `${stretched.length} 個`);

  // 標籤欄灰、填寫欄白。
  const cellAt = (r, c) => t.table._cells.get(`${r},${c}`);
  /**
   * 表單配色改成跟文件主題同一家了。
   *
   * 原本是寫死的灰——那在整份文件都黑白的時候沒問題，但表格早就跟著主題走，於是
   * 一份履歷裡前兩張表是灰白的、後兩張是綠底的，而且剛好在同一頁上對照得一清二楚。
   *
   * 這兩條斷言的是**深淺關係**而不是色碼：表單的字一樣大，階層全靠深淺——區塊標題
   * 要比欄位名深。色碼以後還會調，這個關係不該變。
   */
  ok('標籤欄是最淡的那一階', cellAt(1, 0) && cellAt(1, 0).shadingColor === '#F5F8F6',
    cellAt(1, 0) && cellAt(1, 0).shadingColor);
  ok('填寫欄是白的', cellAt(1, 1) && cellAt(1, 1).shadingColor === '#FFFFFF',
    cellAt(1, 1) && cellAt(1, 1).shadingColor);
  ok('區段帶比標籤深（表單的階層靠深淺，不靠字級）',
    cellAt(0, 0) && cellAt(0, 0).shadingColor === '#EAF1ED'
    && cellAt(0, 0).shadingColor !== cellAt(1, 0).shadingColor,
    cellAt(0, 0) && cellAt(0, 0).shadingColor);
  ok('標籤是粗體', cellAt(1, 0) && cellAt(1, 0)._props['font.bold'] === true);
  ok('填寫欄不是粗體', cellAt(1, 1) && cellAt(1, 1)._props['font.bold'] === false);

  /**
   * 標籤欄要照標籤算寬度，而且不能吃掉填寫的空間。
   *
   * 最長的標籤是「出生日期」四個中文字。欄寬要放得下它，又不能超過一組欄位可用
   * 寬度的四成——表單的重點是給人寫字的地方。
   */
  const lw = cellAt(1, 0).columnWidth;
  ok('標籤欄放得下最長的標籤', lw >= 4 * 11, `${lw}`);
  ok('標籤欄沒有吃掉填寫空間', lw <= (TABLE_WIDTH_PT / 2) * 0.42 + 0.5, `${lw}`);
  ok('填寫欄比標籤欄寬', cellAt(1, 1).columnWidth > lw,
    `${cellAt(1, 1).columnWidth} vs ${lw}`);

  // 列高要講出來，不然 Word 會用空段落的高度決定。
  ok('每一列都設了列高', log.heights.length === 7, `${log.heights.length} 列`);
  ok('列高是 24 點', log.heights.every((h) => h.pt === 24), JSON.stringify(log.heights[0]));

  // 上下內距要小，不然列印出來一頁放不了幾列。
  const tops = log.padding.filter((q) => q.loc === 'Top');
  ok('上下內距壓到 2 點', tops.length > 0 && tops.every((q) => q.pt <= 2),
    JSON.stringify(tops[0]));

  // lines: 8 要真的給八行。
  ok('自傳給了八行的空間', log.extraParas.length === 7, `${log.extraParas.length} 行`);

  // 框線是細的，不是報表那種粗黑。
  ok('設了框線', log.borders.length >= 1);

  /**
   * title 和 note 不可以只寫在 schema 裡。
   *
   * 這一整輪最常見的壞法就是「工具說明寫了這個參數，程式卻沒讀」——模型照著送，
   * 然後什麼都沒發生，而且沒有任何訊息指向漏掉的那一行。
   */
  ok('標題印在表單上方', log.paras.some((q) => q.text === '員工履歷表' && q.where === 'body'),
    JSON.stringify(log.paras));
}

// ── 標題與下方小字 ──────────────────────────────────────────────────────────
{
  const log = fakeWord();
  await executeTool('word_insert_form', {
    title: '出差申請單',
    note: '本表所蒐集的個人資料僅供人事管理使用。',
    fields: [{ label: '申請人' }, { label: '部門' }],
  });
  ok('標題有畫', log.paras.some((q) => q.text === '出差申請單'), JSON.stringify(log.paras));
  ok('說明在表單下面',
    log.paras.some((q) => q.where === 'table:After' && /個人資料/.test(String(q.text))),
    JSON.stringify(log.paras));
}

// ── 合併不支援的 Word ──────────────────────────────────────────────────────
{
  /**
   * 沒有 mergeCells 也要畫得出來。
   *
   * `mergeCells` 是 WordApi 1.4。舊版 Word 上合併會丟例外——那時候寧可畫出一張
   * 格子多一點的表單，也不要整個失敗。
   */
  const log = fakeWord({ noMerge: true });
  const res = await executeTool('word_insert_form', {
    fields: [{ section: '基本資料' }, { label: '姓名' }, { label: '電話' },
      { label: '地址', span: true }],
  });
  ok('舊版 Word 還是畫得出表單', res && res.ok && String(res.content).includes('表單'),
    JSON.stringify(res).slice(0, 120));
  ok('沒有合併就不硬合併', log.merges.length === 0, `${log.merges.length} 次`);
  const t = log.tables[0];
  ok('格子照樣塗好', t && t.table._cells.size >= 4, t && t.table._cells.size);
}

// ── 擋下不會用的情況 ────────────────────────────────────────────────────────
{
  fakeWord();
  // executeTool 不丟例外，它回 { ok: false, error }——那是送回模型的那段字。
  const bad = await executeTool('word_insert_form', { fields: [] });
  const msg = String(bad && bad.error);
  ok('沒有欄位會說清楚怎麼給', /section/.test(msg) && /lines/.test(msg), msg);
}

// ── 分成好幾次送，還是一張表 ────────────────────────────────────────────────
{
  /**
   * 「還不錯，只是他會重複好幾個」。
   *
   * 模型會把一張表單拆成好幾次呼叫送過來——一次基本資料、一次任職資料、一次緊急
   * 聯絡人——而每一次都開一張新表，於是文件裡出現一張完整的表單、一條孤伶伶的
   * 「基本資料」灰帶、一張只有兩個欄位的殘表，然後又一張完整的。
   *
   * 工具說明裡早就寫了「一個陣列給完」，它照樣拆。規則擋不住的事情要用機制擋：
   * 上一筆寫入也是表單的話，這一次就把它拆掉重建，合成同一張。
   */
  const log = fakeWord();
  await executeTool('word_insert_form', {
    title: '員工履歷表',
    fields: [{ section: '基本資料' }, { label: '中文姓名' }, { label: '英文姓名' }],
  });
  const first = log.tables.length;
  const res = await executeTool('word_insert_form', {
    fields: [{ section: '任職資料' }, { label: '到職日期' }, { label: '所屬部門' }],
  });
  ok('第二次是接上去，不是另開一張',
    res && res.ok && /同一張表單/.test(String(res.content)), JSON.stringify(res).slice(0, 140));
  ok('把前一張拆掉了', log.deleted.includes('table'), JSON.stringify(log.deleted));

  // 重建出來的那張要同時有兩段的欄位。
  const built = log.tables[log.tables.length - 1];
  const flat = built.values.flat().join('|');
  ok('合併後兩段都在', /基本資料/.test(flat) && /任職資料/.test(flat) && /所屬部門/.test(flat), flat);
  // 兩條區段帶 + 兩列各兩組欄位 = 四列。多一列代表有一列是空的。
  ok('列數是兩段加起來', built.rows === 4, `${built.rows} 列`);
  ok('第一次只有一張表', first === 1, `${first}`);

  // 標題不可以再印一次——截圖裡出現過「員工履歷表員工履歷表」。
  ok('舊標題被收走了', log.deleted.includes('para:員工履歷表'),
    JSON.stringify(log.deleted));
}

// ── 中間插了別的東西就不再續接 ──────────────────────────────────────────────
{
  const log = fakeWord();
  await executeTool('word_insert_form', { fields: [{ label: '姓名' }] });
  await executeTool('word_insert_text', { text: '以下為附註' });
  const res = await executeTool('word_insert_form', { fields: [{ label: '電話' }] });
  ok('中間有別的寫入就重新開一張',
    res && res.ok && !/同一張表單/.test(String(res.content)),
    String(res && res.content).slice(0, 60));
  ok('沒有去刪別人的東西', !log.deleted.includes('table'), JSON.stringify(log.deleted));
}

console.log(`\n${pass} passed, ${fail} failed`);
/**
 * 標籤要撐成同一個寬度——中文表單的分散對齊。
 *
 * 「我真的不知道為甚麼兩欄設計的最後一列就是沒有 tab」。放大看才發現：不是最後一列
 * 跑掉，是前面幾列被撐開了而它沒有——**最後一列才是正常的**。那表示分散對齊是模型
 * 一列一列設的，設到最後漏了。這正是「一段一段決定」會出錯的典型：每列分開看都對，
 * 合起來就是不齊。
 *
 * Word JS 的 Alignment 沒有 Distributed，所以真的對齊屬性設不了。中文表單傳統上本來
 * 就是手動補全形空白做出來的——這裡做同一件事，只是由程式統一做，不會漏掉任何一列。
 */
{
  const log = fakeWord();
  await executeTool('word_insert_form', {
    fields: [
      { label: '客戶名稱', value: '強合科技' },
      { label: '聯絡人', value: '王小明' },
      { label: '電話', value: '02-8765-4321' },
      { label: '統一編號', value: '87654321' },
    ],
    columns: 1,
  });
  const cells = (log.tables[0] && log.tables[0].values) || [];
  const labels = cells.map((r) => String((r || [])[0] || ''));
  ok('每個標籤都被撐到同一個寬度',
    new Set(labels.filter(Boolean).map((t) => [...t].length)).size === 1,
    JSON.stringify(labels));
  ok('最長的那個不會被動到', labels.includes('客戶名稱') || labels.includes('統一編號'),
    JSON.stringify(labels));
}

process.exit(fail ? 1 : 0);
