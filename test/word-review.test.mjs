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
  const log = { deleted: [], headerSet: [], breaks: [], rewrites: [], fontSets: [], deletedTables: [], styleSets: [] };
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
      style: '',
      // 內容住在儲存格裡還是正文裡——版面檢查靠這個分辨表單和報告。
      tableNestingLevel: typeof spec === 'string' ? 0 : (spec.inCell ? 1 : 0),
      // 只裝一張圖的段落 text 是空的——清空段落的邏輯必須看得見這個。
      inlinePictures: { items: (typeof spec !== 'string' && spec.pic) ? [{}] : [], load() {} },
      delete: () => log.deleted.push(i),
      insertBreak: (kind, where) => log.breaks.push({ i, kind, where }),
      insertText: (txt) => { node.text = txt; log.rewrites.push({ i, txt }); return node; },
      // 封面檢查要問字級。沒有 font 的假物件測不到那段，而真實的 Word 一定給得出來。
      font: typeof spec === 'string' ? undefined : (spec.size === undefined ? undefined : {
        size: spec.size,
        load() {},
        set bold(v) { log.fontSets.push({ i, bold: v }); },
        set color(v) { log.fontSets.push({ i, color: v }); },
      }),
      load() {},
    };
    Object.defineProperty(node, 'styleBuiltIn', {
      get() { return node._style ?? (typeof spec === 'string' ? '' : (spec.style || '')); },
      set(v) { node._style = v; log.styleSets.push({ i, v }); },
    });
    return node;
  });
  const tables = tableSpecs.map((spec, i) => {
    const t = {
      delete: () => log.deletedTables.push(i),
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
  /**
   * 換頁的門檻是「底下有小節」，不是「是個大章節」。
   *
   * 合約是反例：每一條都是大章節，但常常只有兩三行。全部換頁就變成一頁一條，整份
   * 文件空得發慌——「新章節新一頁」要解決的是讀者找不到章節邊界，一頁三行不會讓
   * 邊界更清楚，只是把文件拉長。
   *
   * 這個樣本三個大章節都有編號，但只有兩個底下有小節；再扣掉第一個（它緊接在目錄
   * 後面，目錄已經把頁收掉了），所以只該有一個換頁。
   */
  const log = fakeWord([
    H1('產品概述'), '內文', H2('背景'), '內文',
    H1('目標使用者'), '內文', H2('使用者輪廓'), '內文',
    H1('附則'), '內文',
  ]);
  await executeTool('word_review_layout', {});
  ok('只有「底下有小節」且不是第一個的大章節才換頁',
    log.breaks.length === 1, JSON.stringify(log.breaks));
  ok('沒有小節的章節不換頁（合約那種一條三行的不要獨佔一頁）',
    !log.breaks.some((b) => b.i === 8), JSON.stringify(log.breaks));
  ok('換的是分頁', log.breaks.every((b) => b.kind === 'Page' && b.where === 'Before'),
    JSON.stringify(log.breaks));
  const got = log.rewrites.map((r) => r.txt);
  ok('大章節編成一、二、三', got.includes('一、產品概述') && got.includes('二、目標使用者')
    && got.includes('三、附則'), JSON.stringify(got));
  // 編號跟換頁是兩回事：沒有小節的章節照樣有編號，只是不換頁。
  ok('小節編成 1.1 / 2.1（跟著所屬的大章節走）',
    got.includes('1.1 背景') && got.includes('2.1 使用者輪廓'), JSON.stringify(got));

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

// ── 封面檢查 ──────────────────────────────────────────────────────────────
{
  /**
   * 這幾條是一個 bug 留下來的。
   *
   * 封面檢查原本用猜的：「文件前十段裡最大的那一段就是封面標題」。那在有封面的文件裡
   * 剛好成立，在沒有封面的文件裡什麼都不成立——合約沒有封面，於是它抓到本文段落，
   * 放大成 30pt 粗體置中。而收尾檢查每一輪都跑、文件是一段一段長出來的，所以每一輪
   * 那個窗口裡都有不同段落，整份合約到處是巨大的句子。
   *
   * 修法是改成要有憑據：只修 word_insert_cover 自己做過、而且文字還對得上的那一段。
   * 下面第一條是這個 bug 的回歸測試，也是這組裡最重要的一條——「不該動的時候不要動」
   * 比「該修的時候有修」更要緊，因為前者壞掉的時候使用者是整份文件被弄花。
   */
  const T = (text, size) => ({ text, size });

  const contract = fakeWord([
    T('軟體委託開發合約書', 12),
    T('甲方委託乙方開發軟體系統，雙方同意訂立本合約，條款如下：', 12),
    T('一、乙方受甲方委託，開發「＿＿＿系統」。', 12),
  ]);
  await executeTool('word_review_layout', {});
  ok('沒有封面的文件一個字都不會被放大', contract.fontSets.length === 0,
    JSON.stringify(contract.fontSets));

  const noCover = fakeWord([T('備忘', 10), T('內容', 10)]);
  await executeTool('word_review_layout', {});
  ok('看起來像標題也不算——沒做過封面就是沒有封面', noCover.fontSets.length === 0,
    JSON.stringify(noCover.fontSets));
}

// ── 檢查要因文件類型而異 ──────────────────────────────────────────────────
{
  /**
   * 「最後檢查有點狹窄，好像什麼類似的資料都這樣檢查，但應該因應不同文件有不同檢查」。
   *
   * 章節編號和大章節換頁是報告類文件的規矩。履歷表沒有章節——把「基本資料」編成
   * 「一、基本資料」並讓它獨佔一頁是錯的，那四個字下面只有一張表。一視同仁的檢查
   * 在報告上是幫忙，在表單上是破壞，而且破壞得很徹底：編號寫進標題文字裡、換頁把
   * 一頁的表單拆成四頁。
   *
   * 認法是看內容住在哪裡——表單的字幾乎全在儲存格裡。
   */
  const H1 = (text) => ({ text, style: 'Heading 1' });
  const cell = (text) => ({ text, inCell: true });

  // 履歷：兩個區塊標題，其餘全在儲存格裡
  const form = fakeWord([
    H1('基本資料'), cell('姓名'), cell(''), cell('員工編號'), cell(''),
    H1('聯絡資訊'), cell('行動電話'), cell(''), cell('電子郵件'), cell(''),
  ]);
  await executeTool('word_review_layout', {});
  ok('表單不會被編號', form.rewrites.length === 0, JSON.stringify(form.rewrites));
  ok('表單不會被換頁', form.breaks.length === 0, JSON.stringify(form.breaks));

  // 報告：同樣兩個 H1，但內容在正文段落裡
  const report = fakeWord([
    H1('產品概述'), '內文', { text: '背景', style: 'Heading 2' }, '內文',
    H1('目標使用者'), '內文', { text: '輪廓', style: 'Heading 2' }, '內文',
    H1('功能需求'), '內文',
  ]);
  await executeTool('word_review_layout', {});
  ok('報告照樣編號', report.rewrites.length > 0, JSON.stringify(report.rewrites));
}

// ── 圖片不是空白 ──────────────────────────────────────────────────────────
{
  /**
   * 「月曆也有 BUG，好像被最後檢查刪除了」。
   *
   * 年曆、圖表、流程圖都是以行內圖片的形式掛在一個段落上，而那個段落的 text 是空字串。
   * 空段落清理看的就是 text，於是它把圖整個刪掉——使用者看到「AI 說做好了，但文件是
   * 空的」，訊息裡還明寫著「多出來的空白段落也清掉了」。
   *
   * 前提錯了：段落空不空該看它有沒有內容，不是有沒有文字。
   */
  const pic = { text: '', pic: true };
  const log = fakeWord(['前言', '', '', pic, '', '', '', '結語']);
  await executeTool('word_review_layout', {});
  const picIndex = 3;
  ok('裝著圖的段落不會被當成空白刪掉', !log.deleted.includes(picIndex),
    JSON.stringify(log.deleted));
  ok('旁邊真的空白的還是會清掉', log.deleted.length > 0, JSON.stringify(log.deleted));
}

// ── 表格內部的格式一致性 ──────────────────────────────────────────────────
{
  /**
   * 「出席人員的請假那列就是跟上面格式呈現不一致，但 AI 沒有檢查和修正」。
   *
   * 原本的檢查全部是「整份文件」層級的——封面標題、章節編號、章節換頁、空段落、
   * 長表格表頭。沒有一條在看一張表的內部，所以這種毛病從頭到尾沒有人看得見。
   *
   * 偵測條件要可以量，不能是「看起來怪怪的」：這張表多數列的第一格都有短字（標籤配值
   * 的形狀），而某幾列的第一格是空的、內容卻從第二欄開始——那就是對不齊。
   *
   * 只報告不自動修。自動修要把文字從一格搬到另一格，而「該搬到哪一格」得看那張表在講
   * 什麼；同一個函式已經因為「自動修正的範圍比判斷的範圍大」誤判過三次。偵測不需要
   * 語意，修正需要——所以修正留給看得到語意的那一方。
   */
  const tbl = (values) => ({ rows: values.length, values, headerRowCount: 0 });

  const bad = fakeWord(['表格前'], [tbl([
    ['出席', '林志豪、陳怡君'],
    ['列席', '周俊廷'],
    ['', '吳宗翰（財務部）'],
  ])]);
  const res = await executeTool('word_review_layout', {});
  ok('抓得到對不齊的那一列', /欄位對不齊/.test(String(res.content)), String(res.content).slice(0, 160));
  ok('不會自己動手搬格子', bad.rewrites.length === 0, JSON.stringify(bad.rewrites));

  const good = fakeWord(['表格前'], [tbl([
    ['出席', '林志豪、陳怡君'],
    ['列席', '周俊廷'],
    ['請假', '吳宗翰（財務部）'],
  ])]);
  const res2 = await executeTool('word_review_layout', {});
  ok('整齊的表格不會被誤報', !/欄位對不齊/.test(String(res2.content)),
    String(res2.content).slice(0, 160));
  void good;

  // 空白表單不是對不齊——標籤有、值沒填，那是等著人填的。
  fakeWord(['表格前'], [tbl([['姓名', ''], ['電話', ''], ['地址', '']])]);
  const res3 = await executeTool('word_review_layout', {});
  ok('留白等人填的表單不會被誤報', !/欄位對不齊/.test(String(res3.content)),
    String(res3.content).slice(0, 160));
}

// ── 整張空白的表格 ────────────────────────────────────────────────────────
{
  /**
   * 「檔案多一個空白的沒完整刪除，但她還是留著」。
   *
   * 空段落清理看的是段落，而表格不是段落。一張 20 列 4 欄的空表格在 items 裡是 80 個
   * 段落，但它們全部在表格裡，所以連空段落那條路都走不到——整張留在文件裡佔掉一頁。
   *
   * 分辨「空表格」和「等著填的表單」靠的是：表單一定有欄位名（履歷的「姓名」「電話」
   * 在那裡，只有值是空的），整張連一個字都沒有的才是殘留物。這個條件嚴格到幾乎不可能
   * 誤判，所以它可以自動修——跟那些「偵測容易、修正需要語意」的不一樣。
   */
  const tbl = (values) => ({ rows: values.length, values, headerRowCount: 0 });

  const junk = fakeWord(['前言'], [tbl([['', '', ''], ['', '', ''], ['', '', '']])]);
  await executeTool('word_review_layout', {});
  ok('整張空白的表格會被刪掉', junk.deletedTables.length === 1,
    JSON.stringify(junk.deletedTables));

  const form = fakeWord(['前言'], [tbl([['姓名', ''], ['電話', ''], ['地址', '']])]);
  await executeTool('word_review_layout', {});
  ok('等著填的表單不會被刪', form.deletedTables.length === 0,
    JSON.stringify(form.deletedTables));
}

// ── 同一組編號條款的樣式一致 ──────────────────────────────────────────────
{
  /**
   * 「五、客戶確認本報價後⋯」那裡沒有 | 的標示。
   *
   * 備註一到四左邊都有區塊線，第五條沒有——模型一段一段套樣式，套到最後一段漏了。
   * 跟標籤分散對齊漏掉最後一列、跟字型字級忽大忽小，是同一個病的第三次出現。
   *
   * 修在檢查裡而不是某個工具裡：這些條款可能是 word_insert_block 做的、也可能是
   * word_insert_text 做的，修在其中一個工具裡，模型走另一條路就又漏掉。檢查看的是
   * 結果，不管它怎麼來的。
   *
   * 多數決在這裡是對的：漏掉的永遠是少數，而「大家都長這樣」本來就是要的答案。
   */
  const P = (text, style) => ({ text, style });
  const log = fakeWord([
    '備註',
    P('一、本報價有效期限為 30 天。', 'Quote'),
    P('二、付款方式：簽約後支付 40%。', 'Quote'),
    P('三、交付時程：45 個工作天。', 'Quote'),
    P('四、不含文案撰寫。', 'Quote'),
    P('五、客戶確認後請簽章回傳。', ''),
  ]);
  await executeTool('word_review_layout', {});
  ok('漏掉樣式的那一段被拉回來', log.styleSets.some((x) => x.i === 5 && x.v === 'Quote'),
    JSON.stringify(log.styleSets));
  ok('本來就一致的不會被動到', log.styleSets.filter((x) => x.i !== 5).length === 0,
    JSON.stringify(log.styleSets));
}

process.exit(fail ? 1 : 0);
