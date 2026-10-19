/**
 * 產生使用者操作手冊（.docx）。
 *
 * ── 為什麼是腳本而不是一份手工文件 ──
 * 手冊裡的內容會跟著側邊欄一起變——按鈕增減、起手式換句、額度門檻調整。手工維護的
 * 文件在第二次改版之後就會開始跟產品不一致，而「說明文件寫的跟實際不一樣」比沒有
 * 說明文件更糟：使用者照著做、做不出來，然後不再相信整份文件。
 *
 * 寫成腳本的話，改一行再跑一次就好，而且 diff 看得出來改了什麼。
 *
 * ── 排版的依據 ──
 * 對照的是使用者給的那份參考文件（久億藥局產品規劃書）的水準：有封面、章節標題有
 * 色塊和級別、表格有標頭底色、提醒框跟內文分得開。不是把文字倒進 Word 就算數。
 *
 * ── 相依 ──
 * 用 ai-agents-office 那邊的 docx 套件，跟 make-icons.mjs 借 sharp 是同一個做法：
 * 這支腳本偶爾才跑一次，兩個 repo 本來就並排簽出。路徑若變了就在這裡安裝 docx。
 *
 * 跑：npm run manual
 */
import fs from 'fs';
import path from 'path';
import { createRequire } from 'module';
import { fileURLToPath } from 'url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(ROOT, 'docs', 'JVision增益集-操作手冊.docx');

const DOCX_PATHS = [
  path.resolve(ROOT, '..', 'ai-agents-office', 'server', 'node_modules', 'docx'),
  'docx',
];

function loadDocx() {
  const require = createRequire(import.meta.url);
  for (const p of DOCX_PATHS) {
    try { return require(p); } catch { /* 試下一個 */ }
  }
  throw new Error(
    '找不到 docx 套件。請把 ai-agents-office 簽出在這個 repo 旁邊，或執行：npm i -D docx',
  );
}

const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, HeadingLevel,
  Header, Footer, PageNumber, PageBreak, TableOfContents, HeightRule,
  convertInchesToTwip,
} = loadDocx();

// ── 色票 ─────────────────────────────────────────────────────────────────────
//
// 跟網頁版手冊同一組。三個應用程式用的是 Microsoft 自己的品牌色——使用者每天在
// 工作列上看到，所以拿它當標記時，顏色本身就帶著資訊。
const C = {
  ink: '16201D',
  ink2: '4A5B55',
  ink3: '788A84',
  line: 'DFE5E2',
  line2: 'B9C5C0',
  accent: '0F6B3F',
  accentSoft: 'E8F2EC',
  bg2: 'F6F8F7',
  warnBg: 'FDF6E8',
  warnLine: 'C9922E',
  danger: 'A4262C',
  word: '2B579A',
  excel: '1D7044',
  ppt: 'C43E1C',
};

const FONT = '微軟正黑體';
const MONO = 'Consolas';

// ── 小工具 ───────────────────────────────────────────────────────────────────

/**
 * 內文一行。
 *
 * text 可以是字串，也可以是片段陣列——片段用 {t, b, mono, color} 描述，
 * 這樣「這一句裡有兩個粗體字」不必拆成三個段落。
 */
function p(text, opts = {}) {
  const runs = (Array.isArray(text) ? text : [{ t: text }]).map(seg =>
    new TextRun({
      text: seg.t,
      bold: !!seg.b,
      font: seg.mono ? MONO : FONT,
      size: (opts.size || 21),
      color: seg.color || opts.color || C.ink,
      shading: seg.mono ? { type: ShadingType.CLEAR, fill: 'EEF2F0' } : undefined,
    }));
  return new Paragraph({
    children: runs,
    spacing: { after: opts.after ?? 140, line: 300 },
    indent: opts.indent,
    alignment: opts.align,
  });
}

/** 章節標題。底下那條線是分隔，不是裝飾——它讓目視掃描時章節邊界很清楚。 */
function h2(text, n) {
  return new Paragraph({
    children: [
      new TextRun({ text: `${String(n).padStart(2, '0')}　`, font: MONO, size: 22, color: C.accent, bold: true }),
      new TextRun({ text, font: FONT, size: 30, bold: true, color: C.ink }),
    ],
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 420, after: 180 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: C.ink, space: 6 } },
  });
}

function h3(text, color) {
  return new Paragraph({
    children: [new TextRun({ text, font: FONT, size: 23, bold: true, color: color || C.ink })],
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 260, after: 100 },
  });
}

/** 項目符號。用 Word 內建的 bullet 清單，這樣使用者自己要加一行也接得上。 */
function li(text, level = 0) {
  const runs = (Array.isArray(text) ? text : [{ t: text }]).map(seg =>
    new TextRun({
      text: seg.t, bold: !!seg.b, font: seg.mono ? MONO : FONT, size: 21,
      color: seg.color || C.ink,
      shading: seg.mono ? { type: ShadingType.CLEAR, fill: 'EEF2F0' } : undefined,
    }));
  return new Paragraph({ children: runs, bullet: { level }, spacing: { after: 70, line: 290 } });
}

/** 編號步驟。 */
function step(n, title, body) {
  return new Paragraph({
    children: [
      new TextRun({ text: `${n}.　`, font: MONO, size: 21, bold: true, color: C.accent }),
      new TextRun({ text: title, font: FONT, size: 21, bold: true, color: C.ink }),
      ...(body ? [new TextRun({ text: `　${body}`, font: FONT, size: 21, color: C.ink2 })] : []),
    ],
    spacing: { after: 110, line: 300 },
    indent: { left: 340, hanging: 340 },
  });
}

/**
 * 提醒框。
 *
 * 做成單格表格而不是縮排段落：Word 的縮排在不同版面設定下會跑掉，而一個有底色和
 * 左邊色條的方塊在任何縮放比例下都還是一個方塊。
 */
function note(title, lines, tone = 'info') {
  const fill = tone === 'warn' ? C.warnBg : C.accentSoft;
  const bar = tone === 'warn' ? C.warnLine : C.accent;
  const body = [
    new Paragraph({
      children: [new TextRun({ text: title, font: FONT, size: 21, bold: true, color: C.ink })],
      spacing: { after: 70 },
    }),
    ...lines.map(l => p(l, { after: 0, size: 20, color: C.ink2 })),
  ];
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
      left: { style: BorderStyle.SINGLE, size: 18, color: bar },
    },
    rows: [new TableRow({
      children: [new TableCell({
        children: body,
        shading: { type: ShadingType.CLEAR, fill },
        margins: { top: 160, bottom: 160, left: 200, right: 200 },
      })],
    })],
  });
}

/**
 * 資料表。
 *
 * 標頭有底色、外框只留上下兩條——欄線太多會讓表格看起來像報稅單。
 */
function table(head, rows, widths) {
  const cell = (text, { bold = false, fill, color } = {}) => new TableCell({
    children: [new Paragraph({
      children: (Array.isArray(text) ? text : [{ t: text }]).map(seg => new TextRun({
        text: seg.t, bold: bold || !!seg.b, font: seg.mono ? MONO : FONT, size: 20,
        color: seg.color || color || C.ink,
      })),
      spacing: { after: 0, line: 280 },
    })],
    shading: fill ? { type: ShadingType.CLEAR, fill } : undefined,
    margins: { top: 100, bottom: 100, left: 140, right: 140 },
  });

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: widths,
    borders: {
      top: { style: BorderStyle.SINGLE, size: 8, color: C.ink },
      bottom: { style: BorderStyle.SINGLE, size: 8, color: C.ink },
      left: { style: BorderStyle.NONE }, right: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: C.line },
    },
    rows: [
      new TableRow({
        tableHeader: true,
        children: head.map(t => cell(t, { bold: true, fill: C.bg2, color: C.ink3 })),
      }),
      ...rows.map(r => new TableRow({ children: r.map(t => cell(t)) })),
    ],
  });
}

/** 提問對照：左邊不好、右邊好。並排是重點——差別要用看的，不是用讀的。 */
function askPair(bad, badWhy, good, goodWhy) {
  const box = (label, labelColor, quote, why) => new TableCell({
    children: [
      new Paragraph({
        children: [new TextRun({ text: label, font: FONT, size: 17, bold: true, color: labelColor })],
        spacing: { after: 60 },
      }),
      new Paragraph({
        children: [new TextRun({ text: `「${quote}」`, font: FONT, size: 20, color: C.ink })],
        spacing: { after: 70, line: 280 },
      }),
      new Paragraph({
        children: [new TextRun({ text: why, font: FONT, size: 18, color: C.ink3 })],
        spacing: { after: 0, line: 270 },
      }),
    ],
    shading: { type: ShadingType.CLEAR, fill: C.bg2 },
    margins: { top: 140, bottom: 140, left: 160, right: 160 },
  });
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    columnWidths: [4680, 4680],
    borders: {
      top: { style: BorderStyle.SINGLE, size: 4, color: C.line },
      bottom: { style: BorderStyle.SINGLE, size: 4, color: C.line },
      left: { style: BorderStyle.SINGLE, size: 4, color: C.line },
      right: { style: BorderStyle.SINGLE, size: 4, color: C.line },
      insideVertical: { style: BorderStyle.SINGLE, size: 4, color: C.line },
      insideHorizontal: { style: BorderStyle.NONE },
    },
    rows: [new TableRow({
      children: [
        box('不夠具體', C.danger, bad, badWhy),
        box('好問法', C.accent, good, goodWhy),
      ],
    })],
  });
}

/** 起手式：一個應用程式一塊，左邊色條標明是誰的。 */
function hostBlock(name, color, prompts) {
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.NONE }, bottom: { style: BorderStyle.NONE },
      right: { style: BorderStyle.NONE }, insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
      left: { style: BorderStyle.SINGLE, size: 18, color },
    },
    rows: [new TableRow({
      children: [new TableCell({
        children: [
          new Paragraph({
            children: [new TextRun({ text: name, font: FONT, size: 21, bold: true, color })],
            spacing: { after: 80 },
          }),
          ...prompts.map(t => new Paragraph({
            children: [new TextRun({ text: `「${t}」`, font: FONT, size: 20, color: C.ink })],
            spacing: { after: 50, line: 280 },
          })),
        ],
        shading: { type: ShadingType.CLEAR, fill: C.bg2 },
        margins: { top: 160, bottom: 160, left: 200, right: 200 },
      })],
    })],
  });
}

/**
 * 截圖要放的位置。
 *
 * -- 為什麼先留空格而不是等有圖再排版 --
 * 文字和圖是兩個人在不同時間做的事。先把格子挖好，版面就定了——之後把圖貼進來，
 * 前後文不會被推走、頁碼不會整份重跑。反過來做的話，每補一張圖就要重看一次全文。
 *
 * -- 為什麼格子裡要寫說明 --
 * 「這裡放圖」沒有用；三個月後沒有人記得是放哪一張。寫清楚要截什麼畫面、
 * 什麼狀態，截圖的人照著做就好。
 */
function shot(caption, rows = 7) {
  const empty = Array.from({ length: rows }, () => new Paragraph({
    children: [], spacing: { after: 0, line: 300 },
  }));
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.DASHED, size: 6, color: C.line2 },
      bottom: { style: BorderStyle.DASHED, size: 6, color: C.line2 },
      left: { style: BorderStyle.DASHED, size: 6, color: C.line2 },
      right: { style: BorderStyle.DASHED, size: 6, color: C.line2 },
      insideHorizontal: { style: BorderStyle.NONE },
      insideVertical: { style: BorderStyle.NONE },
    },
    rows: [new TableRow({
      children: [new TableCell({
        children: [
          new Paragraph({
            children: [new TextRun({
              text: `〔截圖〕${caption}`,
              font: FONT, size: 18, color: C.ink3, italics: true,
            })],
            alignment: AlignmentType.CENTER,
            spacing: { after: 60 },
          }),
          ...empty,
        ],
        margins: { top: 160, bottom: 160, left: 160, right: 160 },
      })],
    })],
  });
}

const gap = (h = 160) => new Paragraph({ children: [], spacing: { after: h } });

// ── 封面 ─────────────────────────────────────────────────────────────────────

const cover = [
  gap(1400),
  new Paragraph({
    children: [new TextRun({ text: '操作手冊', font: FONT, size: 19, bold: true, color: C.accent, characterSpacing: 60 })],
    spacing: { after: 120 },
  }),
  new Paragraph({
    children: [new TextRun({ text: 'JVision for Office', font: FONT, size: 60, bold: true, color: C.ink })],
    spacing: { after: 60 },
  }),
  new Paragraph({
    children: [],
    border: { bottom: { style: BorderStyle.SINGLE, size: 18, color: C.accent, space: 2 } },
    spacing: { after: 260 },
  }),
  p([
    { t: '在 Word、Excel、PowerPoint 的側邊欄，直接對' },
    { t: '你正開著的那份檔案', b: true },
    { t: '提問。它讀得到內容、格式和你選取的範圍，也能在你確認後改回去——不是產生一個新檔案給你下載。' },
  ], { size: 23, color: C.ink2, after: 400 }),
  new Paragraph({
    children: [
      new TextRun({ text: '■ ', font: FONT, size: 21, color: C.word }),
      new TextRun({ text: 'Word　　', font: FONT, size: 21, bold: true, color: C.word }),
      new TextRun({ text: '■ ', font: FONT, size: 21, color: C.excel }),
      new TextRun({ text: 'Excel　　', font: FONT, size: 21, bold: true, color: C.excel }),
      new TextRun({ text: '■ ', font: FONT, size: 21, color: C.ppt }),
      new TextRun({ text: 'PowerPoint', font: FONT, size: 21, bold: true, color: C.ppt }),
    ],
    spacing: { after: 0 },
  }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── 目錄 ─────────────────────────────────────────────────────────────────────
//
// 用 Word 真正的目錄欄位，不是手寫清單——手寫的沒有頁碼，而使用者拿到的是一份
// 要列印、要翻頁的文件。
//
// 代價是欄位剛產生時是空的，要 Word 去算一次。所以 Document 開了 updateFields，
// Word 開檔時會自己更新（某些版本會先問一句「要更新功能變數嗎」，按「是」）。
// 真的沒更新的話：全選後按 F9。
const SECTIONS = [
  '安裝', '登入', '介面導覽', '怎麼問', '確認、復原、停止',
  '資料來源與附件', '跨應用交接', '額度', '疑難排解', '常見問題',
];

const toc = [
  h3('目錄'),
  new TableOfContents('目錄', {
    hyperlink: true,
    // 1-2 代表收 Heading1 和 Heading2：章節和小節都進目錄，再深就太碎了。
    headingStyleRange: '1-2',
  }),
  new Paragraph({ children: [new PageBreak()] }),
];

// ── 內文 ─────────────────────────────────────────────────────────────────────

const body = [
  // 01 安裝
  h2('安裝', 1),

  h3('先確認你要裝哪一個'),
  p([
    { t: '有兩個版本，' }, { t: '安裝檔長得不一樣、連的伺服器也不一樣', b: true },
    { t: '。裝錯的話資料會進到另一套系統，所以拿到檔案時先看檔名。' },
  ]),
  table(
    ['版本', '安裝檔', '這是什麼'],
    [
      [[{ t: '正式區', b: true }], [{ t: 'JVisionSetup-jv.exe', mono: true }],
        '日常使用的版本。資料都是真的，做出來的文件可以直接用。'],
      [[{ t: '測試區', b: true }], [{ t: 'JVisionSetup-test.exe', mono: true }],
        '試用新功能的版本。資料獨立，不會影響正式區。'],
    ],
    [1500, 3100, 4760],
  ),
  gap(),
  note('測試區的側邊欄最上面有一條橘色橫幅', [
    [{ t: '上面寫著「測試　·　這不是正式版，測試完請記得移除」。' },
     { t: '看到那條橫幅就代表你現在在測試區', b: true },
     { t: '——重要的文件不要在這裡做。正式區沒有這條橫幅。' }],
  ], 'warn'),
  gap(),
  shot('測試區側邊欄最上方的橘色橫幅（正式區沒有這一條）', 5),
  gap(),

  h3('安裝步驟'),
  step(1, '關閉所有 Office 程式。', 'Word、Excel、PowerPoint 都要關，包含最小化的視窗。安裝程式會自己檢查，沒關的話它會停下來告訴你。'),
  step(2, '執行安裝檔。', '不需要系統管理員權限，所有設定只寫進你自己的帳號。'),
  step(3, '選要安裝哪幾個應用程式。', '預設三個都裝。之後想改，重跑一次安裝檔就行。'),
  step(4, '打開 Word（或 Excel、PowerPoint）。', '「常用」索引標籤的最右邊會多一顆 JVision 按鈕，點它就開側邊欄。'),
  gap(),
  shot('安裝程式主畫面——勾選要安裝的應用程式那一步', 7),
  gap(),
  shot('Word 功能區「常用」最右邊的 JVision 按鈕（整列功能區，把按鈕圈出來）', 5),
  gap(),
  note('Windows 可能會攔下載', [
    [{ t: '從瀏覽器下載時，Chrome 或 Windows SmartScreen 可能顯示警告。那是因為安裝檔還沒有購買程式碼簽章憑證，' },
     { t: '不是檔案有問題', b: true },
     { t: '。選「保留」或「仍要執行」即可，或請 IT 用內部派送管道發給你。' }],
  ], 'warn'),
  gap(),
  p([
    { t: '兩個版本可以同時裝，在增益集清單裡是兩個不同的項目，' },
    { t: '名稱上看得出來哪個是測試區', b: true }, { t: '。' },
  ]),
  gap(60),
  h3('移除'),
  p([
    { t: 'Windows 設定 →「應用程式」→ 找到 ' }, { t: 'JVision for Office', b: true },
    { t: ' → 解除安裝。或執行安裝目錄下的 ' }, { t: '解除安裝.cmd', mono: true },
    { t: '。移除只會拿掉側邊欄，' }, { t: '你的檔案一個都不會動', b: true }, { t: '。' },
  ]),

  // 02 登入
  h2('登入', 2),
  p([
    { t: '側邊欄第一次打開會要求登入。帳號和 ' }, { t: 'AI Agents Office 網站是同一組', b: true },
    { t: '，不用另外申請。登入狀態會保留，之後開其他檔案不必再登一次。' },
  ]),
  table(
    ['方式', '適用', '說明'],
    [
      [[{ t: 'AD 帳號', b: true }], '公司內部', '用你每天開機的網域帳號密碼。這條路同時會接通「郵件」資料來源。'],
      [[{ t: '電子信箱', b: true }], '全部', '在 AI Agents Office 網站註冊的信箱與密碼。'],
      [[{ t: 'Google', b: true }], '全部', '在「電子信箱」分頁下方。公司內部版本僅限名單內的帳號。'],
    ],
    [1700, 1400, 6260],
  ),
  shot('登入畫面——顯示「AD 帳號」和「電子信箱」兩個分頁，以及下方的 Google 登入按鈕', 7),
  gap(),
  gap(),
  note('看到哪幾種，由伺服器決定', [
    '公司內部版本以 AD 為主、電子信箱為輔；對外版本沒有 AD 分頁。你看到的畫面就是你這套部署該有的樣子。',
  ]),
  gap(),
  h3('登出'),
  p([{ t: '點右上角的' }, { t: '人像', b: true }, { t: '圖示，面板最下面有「登出」。' }]),

  // 03 介面導覽
  h2('介面導覽', 3),
  h3('上方工具列'),
  table(
    ['按鈕', '作用'],
    [
      [[{ t: '盾牌', b: true }], [{ t: '切換' }, { t: '每次確認', b: true }, { t: '和' }, { t: '自動允許', b: true }, { t: '。顏色會跟著變，因為它決定 AI 能不能直接改你的檔案。' }]],
      [[{ t: '時鐘', b: true }], '這份檔案過去的對話紀錄。唯讀，看完會把你送回目前這段對話。'],
      [[{ t: '新對話', b: true }], '重新開始。舊紀錄留在時鐘裡，不會消失。'],
      [[{ t: '復原', b: true }], '退回 AI 上一步的修改。旁邊的數字是還能退幾步。'],
      [[{ t: '地球', b: true }], 'AI 回覆要用哪種語言。只影響之後的回覆，不會翻譯檔案裡既有的內容。'],
      [[{ t: '人像', b: true }], '額度、用量來源、申請額度、登出。'],
    ],
    [1600, 7760],
  ),
  gap(),
  shot('側邊欄上方工具列整列——六顆按鈕由左至右（盾牌、時鐘、新對話、復原、地球、人像）', 4),
  gap(),
  h3('中間'),
  li([{ t: '檔名列', b: true }, { t: '——目前對著哪份檔案。Excel 會再顯示現在是哪張工作表。' }]),
  li([{ t: '導覽列', b: true }, { t: '——Word 點「段落」、Excel 點工作表名稱，會列出結構，點一行就跳到那個位置。' }]),
  li([{ t: '對話區', b: true }, { t: '——AI 的回覆與它做過的每一個動作。' }]),
  gap(60),
  h3('下方輸入區'),
  li([{ t: '＋', b: true }, { t: '——選資料來源、上傳附件、插入常用指令。' }]),
  li([{ t: '輸入框', b: true }, { t: '——' }, { t: 'Enter', mono: true }, { t: ' 送出，' }, { t: 'Shift+Enter', mono: true }, { t: ' 換行。' }]),
  li([{ t: '選取範圍標籤', b: true }, { t: '——你在檔案裡選了東西，這裡會顯示；不想帶就按 ×。' }]),
  li([{ t: '右下角', b: true }, { t: '——模型選擇。部分部署不開放自選，就不會顯示。' }]),
  li([{ t: '送出後', b: true }, { t: '——箭頭變成停止鍵。' }]),

  // 04 怎麼問
  h2('怎麼問', 4),
  p([
    { t: '它看得到整份檔案，所以' }, { t: '不用把內容貼給它', b: true },
    { t: '。直接說你要什麼，愈具體愈好。' },
  ]),
  askPair(
    '幫我改一下這份文件', '改什麼？改成什麼樣？它只能猜，然後多半猜不中。',
    '把第三節改寫成正式公文語氣，長度不要超過原本的一半', '說了範圍、方向、限制——三樣都具體。',
  ),
  gap(),
  askPair(
    '做一份簡報', '它會先問你主題，等於多花一輪。',
    '做一份「智慧工廠數位轉型」提案簡報給製造業客戶，8 頁：封面、問題、解法、三階段導入、預期效益、案例、報價、下一步',
    '主題、對象、頁數、每頁內容都給了，一次到位。',
  ),
  gap(),
  h3('先選再問'),
  p([
    { t: '在檔案裡選一段文字、一塊儲存格或一個物件，再提問——輸入框上方會出現標籤，表示這次提問會帶上它。' },
    { t: '「這段」「這張表」「這個圖」這類說法，要先選才有意義。', b: true },
  ]),
  shot('輸入框上方出現選取範圍標籤的樣子（先在文件裡選一段文字再截）', 4),
  gap(),
  h3('各應用程式的起手式'),
  hostBlock('Word', C.word, [
    '幫我摘要這份文件的重點', '把選取的段落改寫成正式公文',
    '檢查全文的錯字和用詞', '整理標題階層並產生目錄',
  ]),
  gap(100),
  hostBlock('Excel', C.excel, [
    '這張表在算什麼？', '幫我找出所有錯誤的公式',
    '選取範圍幫我做一張摘要', '這些數字有什麼趨勢？',
  ]),
  gap(100),
  hostBlock('PowerPoint', C.ppt, [
    '幫我把這份簡報做成一致的版面', '這張投影片的物件幫我對齊',
    '依這些重點做三張投影片', '幫我做一張架構流程圖',
  ]),
  gap(),
  note('中途改主意沒關係', [
    '做到一半想換方向，直接在輸入框講就好，不必等它做完。它會在下一步把你的新要求併進去。',
  ]),

  // 05 確認、復原、停止
  h2('確認、復原、停止', 5),
  p([
    { t: '讀取永遠不會問你，寫入才會。', b: true },
    { t: '這條線是固定的——它看檔案不需要你同意，動檔案需要。' },
  ]),
  h3('每次確認 vs 自動允許'),
  table(
    ['模式', '行為', '什麼時候用'],
    [
      [[{ t: '每次確認', b: true }], '每一個修改動作都先問過你，你看得到它要改什麼。', '改重要的檔案、或你還不熟悉它的時候。'],
      [[{ t: '自動允許', b: true }], '直接修改，不再逐項詢問。', '連續做很多步的工作，例如整份簡報的排版。'],
    ],
    [1700, 4100, 3560],
  ),
  gap(),
  note('自動允許也有例外', [
    [{ t: '刪除整段文字、刪除投影片、刪除工作表，一律還是會問。', b: true },
     { t: '那些動作沒辦法完整復原，所以不管開哪個模式都要你點頭。' }],
  ], 'warn'),
  gap(),
  shot('確認框「AI 想要修改這份文件」——看得到它要改什麼，以及「允許／不要」兩顆按鈕', 6),
  gap(),
  h3('復原'),
  p([
    { t: '工具列的' }, { t: '復原', b: true }, { t: '鍵退回 AI 的上一步，旁邊的數字是還剩幾步可退。它和 Word/Excel 自己的 ' },
    { t: 'Ctrl+Z', mono: true }, { t: ' 是兩套——這顆按鈕退的是「AI 做的那一步」，可能包含好幾個動作。' },
  ]),
  h3('停止'),
  p([
    { t: '送出後箭頭會變成停止鍵。按下去會停在目前這一步，' },
    { t: '已經做完的部分留在檔案裡', b: true }, { t: '——需要的話用復原鍵退回。' },
  ]),

  // 06 資料來源
  h2('資料來源與附件', 6),
  p([{ t: '按輸入框左邊的 ' }, { t: '＋', b: true }, { t: '。' }]),
  table(
    ['項目', '作用'],
    [
      [[{ t: '搜尋網路', b: true }], '需要外部最新資訊時開啟，例如查市場數據或法規。'],
      [[{ t: '郵件', b: true }], [{ t: '讓它讀得到你的信件。' }, { t: '需要用 AD 帳號登入', b: true }, { t: '，信箱登入不會接通。' }]],
      [[{ t: 'KM', b: true }], '公司知識庫。問內部流程、規範、產品資料時開。'],
      [[{ t: '上傳附件', b: true }], '把另一份檔案交給它——比對兩份報表、依既有報告做簡報。'],
      [[{ t: '常用指令', b: true }], '管理者預先準備好的提示詞範本，選了會填進輸入框，送出前可以改。'],
    ],
    [1900, 7460],
  ),
  shot('按下輸入框左邊的「＋」展開的選單——網路、郵件、KM、附件、常用指令', 6),
  gap(),
  gap(),
  note('沒必要就不要全開', [
    '每個來源都會讓它多花時間查。問的是眼前這份檔案的事，一個都不用開。',
  ]),

  // 07 交接
  h2('跨應用交接', 7),
  p('在 Excel 整理完資料，想拿到 Word 寫成報告——直接說：'),
  p([{ t: '「把這張銷售明細整理成一份給主管看的分析報告，我等一下在 Word 繼續」' }],
    { indent: { left: 340 }, color: C.ink2 }),
  shot('Word 側邊欄上方的交接橫幅', 4),
  gap(),
  p([
    { t: '然後打開 Word 的側邊欄，上方會出現一條' }, { t: '交接橫幅', b: true },
    { t: '，點它就接著做。三個應用程式之間都可以這樣傳。' },
  ]),

  // 08 額度
  h2('額度', 8),
  p([{ t: '點右上角' }, { t: '人像', b: true }, { t: '圖示。面板上方是你的帳號，下面是這個月用掉多少。' }]),
  li([{ t: '百分比', b: true }, { t: '——目前用掉的比例。超過七成會轉成警示色。' }]),
  li([{ t: '每日用量', b: true }, { t: '——這個月每一天的長條，今天那根是深色。' }]),
  li([{ t: '用量來源', b: true }, { t: '——拆開給你看是誰用掉的。額度是整個帳號共用的，網站上產文件、AI 團隊、信件助手都扣同一池，不是只有側邊欄。' }]),
  gap(60),
  shot('個人區面板——帳號名稱、本月百分比、每日用量長條、用量來源明細', 8),
  gap(),
  h3('快用完的時候'),
  p([
    { t: '到九成時，對話上方會出現一條提醒，可以直接申請調整上限。點「申請額度」會打開個人區，' },
    { t: '寫一句用途', b: true }, { t: '再送出——理由是必填的，管理者需要它來判斷。' },
  ]),
  shot('額度快用完時對話上方的提醒橫幅，以及個人區下方的申請表單', 7),
  gap(),
  p('送出後面板會顯示「等待管理者審核」，結果以電子郵件通知。同一時間只能有一筆申請在排隊。'),

  // 09 疑難排解
  h2('疑難排解', 9),
  table(
    ['狀況', '怎麼處理'],
    [
      ['側邊欄顯示「連不上伺服器」', '先確認你連得上公司網路或 VPN。確認後仍然不行，把畫面上那串伺服器位址一起回報給 IT。'],
      ['側邊欄一片空白', '關掉側邊欄再開一次。還是空白就重開 Office。'],
      ['功能區沒有 JVision 按鈕', '確認安裝時有勾選這個應用程式，然後完全關閉 Office 再開——包含最小化的視窗。'],
      ['它說看不到我選的東西', '重新選一次，確認輸入框上方出現了選取範圍標籤。'],
      ['改壞了', '按工具列的復原鍵。要退更多步就連按，數字會告訴你還剩幾步。'],
      ['一直說額度不足', '開人像面板看目前用量，從那裡申請調整上限。'],
      ['回覆語言不對', '點地球圖示改。這只影響之後的回覆，檔案裡既有的文字不會被翻譯。'],
    ],
    [3100, 6260],
  ),

  // 10 常見問題
  h2('常見問題', 10),
  h3('我的檔案會被上傳到哪裡嗎？'),
  p([
    { t: '不會。側邊欄是請你電腦上的 Office 幫它讀某個範圍、寫某幾格——' },
    { t: '檔案本身不會離開你的電腦', b: true },
    { t: '。對話內容會送到公司自己的伺服器處理。' },
  ]),
  h3('它會不會把我的公式弄壞？'),
  p('它是透過 Office 本身的介面操作，所以公式關聯、樞紐、格式都還在。這也是為什麼它引用的每個數字都能點回原始位置。'),
  h3('可以同時開好幾個檔案嗎？'),
  p('可以。每份檔案有自己的對話紀錄，側邊欄會跟著你目前作用中的那一份切換。'),
  h3('對話紀錄留多久？'),
  p('留在 AI Agents Office 裡，從時鐘圖示隨時找得到。「開新對話」只是重新開始，不會刪掉舊的。'),
  h3('為什麼我沒有模型選擇？'),
  p('部分部署由公司統一指定模型，不開放自選，所以右下角不會出現。這是設定，不是故障。'),
  h3('可以在 Mac 或網頁版用嗎？'),
  p('目前以 Windows 桌面版 Office 為主。其他平台請詢問 IT。'),

  gap(300),
  new Paragraph({
    children: [new TextRun({
      text: '遇到手冊沒寫到的狀況，或發現它做錯了，請回報給 IT——附上你當時問了什麼，會快很多。',
      font: FONT, size: 19, color: C.ink3,
    })],
    border: { top: { style: BorderStyle.SINGLE, size: 4, color: C.line, space: 8 } },
    spacing: { before: 200 },
  }),
];

// ── 組起來 ───────────────────────────────────────────────────────────────────

const doc = new Document({
  // 目錄是 Word 的功能變數，產生時是空的——這個旗標讓 Word 開檔時自己算一次頁碼。
  // 沒有它的話使用者打開會看到一片空白的目錄，而那是最糟的第一印象。
  features: { updateFields: true },
  creator: 'JVision for Office',
  title: 'JVision for Office 操作手冊',
  description: '增益集使用者操作手冊',
  styles: {
    default: {
      document: { run: { font: FONT, size: 21, color: C.ink } },
    },
  },
  sections: [{
    properties: {
      page: {
        margin: {
          top: convertInchesToTwip(1), bottom: convertInchesToTwip(0.9),
          left: convertInchesToTwip(1), right: convertInchesToTwip(1),
        },
      },
    },
    // 頁首從第二頁起才有意義，但 Word 的首頁不同設定會連頁碼一起拿掉，
    // 所以頁首留白、只在頁尾放頁碼——封面多一個頁碼不礙事，少了頁碼才礙事。
    footers: {
      default: new Footer({
        children: [new Paragraph({
          children: [
            new TextRun({ text: 'JVision for Office · 操作手冊　　', font: FONT, size: 16, color: C.ink3 }),
            new TextRun({ children: [PageNumber.CURRENT], font: MONO, size: 16, color: C.ink3 }),
          ],
          alignment: AlignmentType.RIGHT,
        })],
      }),
    },
    children: [...cover, ...toc, ...body],
  }],
});

fs.mkdirSync(path.dirname(OUT), { recursive: true });
const buf = await Packer.toBuffer(doc);
fs.writeFileSync(OUT, buf);
console.log(`已產生 ${path.relative(ROOT, OUT)}  ${(buf.length / 1024).toFixed(1)} KB`);
console.log(`  ${SECTIONS.length} 個章節：${SECTIONS.join('、')}`);
