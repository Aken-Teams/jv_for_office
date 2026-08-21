/**
 * hosts/excel — the Office.js half of the bridge, for Excel.
 *
 * The server sends `{tool, args}`; this module runs it against the workbook the
 * user actually has open and returns text the model can read. It contains NO AI
 * logic and no network code on purpose: it is the one part that can only be
 * tested inside a real Excel, so everything else was kept out of it.
 *
 * One of three interchangeable host modules — see hosts/index.js for the
 * contract every one of them fills, and docs/multi-host.md for why they live in
 * the same project as one shared shell rather than in three copies of it.
 *
 * Two rules shape all of it:
 *
 *  1. Reads return CSV, not JSON. A 50×8 block is ~2.5k characters as CSV and
 *     ~9k as JSON — and the model reads CSV perfectly well.
 *
 *  2. Everything is capped. An unbounded read of a 200k-row sheet would blow
 *     the context window, the 1MB request-body limit on the API, and the user's
 *     token budget, in that order.
 */
import { t } from '../i18n.js';
import { makeScriptGuard } from './guard.js';

/**
 * The tools THIS build of the pane can actually execute.
 *
 * Sent to the server on every chat so it can hide from the model anything this
 * client doesn't implement. Without it, a pane running older code answers
 * "未知的工具：excel_manage_sheet" and the agent just retries in a loop — which is
 * exactly what happened, and it is not a rare edge case: the manifest declares
 * SharedRuntime with lifetime="long", so closing and reopening the task pane does
 * NOT reload this file. Only a full Excel restart does. Any user who updates the
 * add-in without quitting Excel lands here.
 */
const BASE_TOOLS = [
  'excel_get_overview',
  'excel_get_selection',
  'excel_read_range',
  'excel_search',
  'excel_write_range',
  'excel_manage_sheet',
  'excel_format_range',
  'excel_create_chart',
  'excel_structure_op',
  'excel_table_ops',
  // 資料清理：去重、統一寫法、補空白、拆併欄位。PRD 3.2 列了這四項，
  // 在這支工具存在之前它們全都只能靠 run_script 現寫程式。
  'excel_clean_data',
  // Native precedent tracing needs ExcelApi 1.12, but there is a formula-parsing
  // fallback that works on any build, so this one is always offered.
  'excel_trace_precedents',
  'excel_sheet_style',
  // Compute inside Excel instead of hauling the data into the model's context.
  // See runScript below for why this one matters more than all the others.
  'excel_run_script',
  // Answered by the pane's own UI (see askChoice in app.js), not by Office.js —
  // listed here so the server offers it to the model.
  'excel_ask_user',
];

/** Cheap and synchronous; Office.js answers from a local table. */
export function hasApi(version) {
  try {
    return Office.context.requirements.isSetSupported('ExcelApi', version);
  } catch {
    return false;
  }
}

/**
 * Tools this build can execute ON THIS HOST.
 *
 * Pivot tables need ExcelApi 1.12 to place fields into rows/columns/values —
 * `pivotTables.add()` alone is 1.8, but a pivot with no fields in it is useless.
 * Rather than let the agent call it and fail on older builds, the tool is simply
 * not offered: the server already filters what the model sees by this list, so an
 * unsupported host degrades to "fewer abilities" with no confusing retries.
 */
export function supportedTools() {
  return hasApi('1.12') ? [...BASE_TOOLS, 'excel_create_pivot'] : [...BASE_TOOLS];
}

/**
 * Everything the shell used to assume.
 *
 * app.js knew fourteen tool names, an API path, an MCP prefix, a citation
 * syntax and the word 「活頁簿」. None of that was ever shell business — it was
 * Excel business that happened to be written in the shell because there was
 * only one host. Moved here, adding a host is adding a file.
 *
 * The strings are Traditional Chinese SOURCE text, not translations: they are
 * the keys i18n.js looks up, so 活頁簿 and 文件 are separate dictionary entries
 * and each host's wording gets translated on its own terms.
 */
export const meta = {
  id: 'excel',
  /** Which Office.HostType this module is for. */
  officeHost: 'Excel',
  /** URL segment — /api/excel/chat, /api/excel/tool-result. */
  api: 'excel',
  /** MCP tools reach the pane prefixed: mcp__excel__excel_run_script. */
  mcpPrefix: 'mcp__excel__',
  /**
   * Handed over with the first message so the agent does not have to spend a
   * round trip asking for what it always asks for first.
   */
  overview: { tool: 'excel_get_overview', args: { preview_rows: 5 } },
  /** Answered by the pane's own UI (askChoice), never by Office.js. */
  askUserTool: 'excel_ask_user',
  /** The one write whose cost can be counted BEFORE it happens. */
  inspectTools: ['excel_write_range', 'excel_clean_data'],
  /** Model-written code: needs a snapshot, and its own progress caption. */
  scriptTool: 'excel_run_script',

  /**
   * [[損益表!C42]] → a link that selects that cell. Two capture groups; the
   * shell does not care what they mean, only that goTo() understands them.
   */
  citation: {
    pattern: /\[\[([^\]!]+)!(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?)\]\]/g,
    label: (a, b) => `${a}!${b}`,
  },

  /** What a call is doing while it runs. */
  toolLabels: () => ({
    excel_get_overview: t('正在讀取活頁簿結構…'),
    excel_get_selection: t('正在讀取選取範圍…'),
    excel_read_range: t('正在讀取資料…'),
    excel_search: t('正在搜尋…'),
    excel_write_range: t('等待你確認寫入…'),
    excel_manage_sheet: t('正在調整工作表…'),
    excel_format_range: t('正在套用格式…'),
    excel_create_chart: t('正在建立圖表…'),
    excel_structure_op: t('等待你確認結構變更…'),
    excel_table_ops: t('正在整理表格…'),
    excel_clean_data: t('等待你確認資料清理…'),
    excel_create_pivot: t('等待你確認建立樞紐…'),
    excel_trace_precedents: t('正在追蹤公式來源…'),
    excel_sheet_style: t('正在調整版面…'),
    excel_run_script: t('正在執行程式…'),
  }),

  /**
   * What the model is doing BEFORE the call is dispatched. Composing a long
   * call is generation time — for excel_run_script that is an entire script,
   * usually the longest single stretch of a turn.
   */
  preparingLabels: () => ({
    excel_run_script: t('正在寫程式…'),
    excel_write_range: t('正在整理要寫入的資料…'),
    excel_create_chart: t('正在設定圖表…'),
    excel_create_pivot: t('正在設定樞紐分析表…'),
    excel_format_range: t('正在決定格式…'),
    excel_clean_data: t('正在檢查資料…'),
    excel_manage_sheet: t('正在準備工作表變更…'),
  }),

  /** Sentences that name what this host edits. See the note above. */
  words: {
    loadingTools: '載入 Excel 工具…',
    // ── The pane's own furniture, reached from the markup by data-h ──────────
    // Here rather than written into taskpane.html because all three hosts load
    // the SAME page, and a Word user reading 「選一段儲存格」 is not a
    // translation problem — it is a wrong-product problem.
    title: 'JVision for Excel',
    brandSub: 'for Excel',
    loginLead: '登入你的 AI Agents Office 帳號，開始對這份活頁簿提問。',
    histTip: '這份活頁簿的對話紀錄',
    newChatTip: '開新對話（這份活頁簿的紀錄會留著）',
    docNameTitle: '目前開啟的活頁簿',
    emptyTitle: '問這份活頁簿任何事',
    emptySub: '我讀得到儲存格、公式和你選取的範圍，也能在你確認後寫回去。',
    placeholder: '對這份活頁簿提問…',
    /** The four openers on the empty state. Questions this host can answer. */
    chips: [
      '這張表在算什麼？',
      '幫我找出所有錯誤的公式',
      '選取範圍幫我做一張摘要',
      '這些數字有什麼趨勢？',
    ],
    /**
     * First-run carousel, slides 1-3. The fourth is about the web app and is
     * the same wherever you are, so it stays in the markup.
     */
    intro: [
      {
        title: '在 Excel 裡直接問',
        body: '不用切視窗，也不用把資料複製到別的地方。選一段儲存格，用講的就好。',
        feats: ['選了就懂', '用問的', '就地作答'],
      },
      {
        title: '它看得懂這張表',
        body: '讀得到儲存格、公式和你選取的範圍。可以解釋這張表在算什麼、幫你找出壞掉的公式。',
        feats: ['追公式來源', '揪出錯誤', '解釋算法'],
      },
      {
        title: '動手前都會問過你',
        body: '寫公式、排版、做圖表和樞紐分析表都做得到。每一次要動你的活頁簿之前都會先問，做完也還能復原。',
        feats: ['先問過你', '隨時復原', '隨時喊停'],
      },
    ],
    untitled: '（未命名活頁簿）',
    confirmTitle: 'AI 想要修改這份活頁簿',
    noHistory: '這份活頁簿還沒有對話紀錄。',
    stopped: '已停止。已經做完的部分留在活頁簿裡，需要的話可以用上方的復原鍵退回。',
    localeChanged: '之後的回覆會用{0}。活頁簿裡既有的資料不會被翻譯。',
    autoApproveOn: '已改成自動允許：AI 修改活頁簿時不會再詢問。刪除工作表或整列整欄仍會問，因為那沒辦法復原。',
    autoApproveOff: '已改回每次確認：AI 要修改活頁簿前都會先問過你。',
    noSnapshot: '這份活頁簿太大，程式執行前沒辦法先備份，所以這次的修改復原鍵拿不回來。',
    undoScope: '提醒：AI 做的修改不會進到 Excel 的復原紀錄，Ctrl+Z 退不回來。要退回請用這裡的復原鍵。',
    locationLabel: '目前的工作表：{0}',
    locationMissing: '找不到工作表「{0}」',
  },
};

/** Hard ceiling per read. ~20k cells of typical data lands near 100-200KB of CSV. */
const MAX_READ_CELLS = 20000;
/** Second ceiling, on characters — one column of long text can be small in cells but huge in bytes. */
const MAX_READ_CHARS = 200000;
/** Rows previewed per sheet in the overview. Enough to see headers and a sample. */
const DEFAULT_PREVIEW_ROWS = 5;
const MAX_PREVIEW_ROWS = 20;
const MAX_SEARCH_RESULTS = 50;

// ── CSV ───────────────────────────────────────────────────────────────────────

function csvCell(v) {
  if (v === null || v === undefined) return '';
  const s = String(v);
  return /[",\n\r]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s;
}

/** Grid → CSV, stopping at MAX_READ_CHARS so one monster column can't run away. */
function toCsv(grid) {
  const lines = [];
  let chars = 0;
  for (let r = 0; r < grid.length; r++) {
    const line = grid[r].map(csvCell).join(',');
    chars += line.length + 1;
    if (chars > MAX_READ_CHARS) {
      lines.push(`… 已達字元上限，只顯示前 ${r} 列（共 ${grid.length} 列）`);
      break;
    }
    lines.push(line);
  }
  return lines.join('\n');
}

/** 0-based column index → 'A', 'Z', 'AA' … */
function colName(index) {
  let n = index + 1, s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function a1(rowIndex, colIndex) {
  return `${colName(colIndex)}${rowIndex + 1}`;
}

/** Strip the "Sheet1!" prefix Office.js puts on addresses — the sheet is reported separately. */
function bareAddress(address) {
  const i = address.lastIndexOf('!');
  return i >= 0 ? address.slice(i + 1) : address;
}

// ── Tools ─────────────────────────────────────────────────────────────────────

/**
 * Structure only: sheet names, sizes, and a few preview rows each.
 *
 * This is deliberately the ONLY thing the agent gets for free at the start of a
 * conversation. Sending whole sheets up front is what makes naive versions of
 * this product unusable on real workbooks — Claude for Excel sends metadata here
 * too, and fetches ranges on demand.
 */
async function getOverview(args) {
  const previewRows = Math.min(Math.max(1, args.preview_rows || DEFAULT_PREVIEW_ROWS), MAX_PREVIEW_ROWS);

  return Excel.run(async (ctx) => {
    const wb = ctx.workbook;
    wb.load('name');
    const sheets = wb.worksheets;
    sheets.load('items/name,items/position,items/visibility');
    await ctx.sync();

    const used = sheets.items.map((s) => {
      const u = s.getUsedRangeOrNullObject(true);
      u.load('isNullObject,rowIndex,columnIndex,rowCount,columnCount,address');
      return u;
    });
    await ctx.sync();

    // Second round trip: now that we know where each sheet's data starts, grab
    // just the top few rows of each. One sync for all sheets, not one per sheet.
    const previews = sheets.items.map((s, i) => {
      const u = used[i];
      if (u.isNullObject) return null;
      const rows = Math.min(previewRows, u.rowCount);
      const cols = Math.min(u.columnCount, 40); // 40 columns is plenty to recognise a table
      const r = s.getRangeByIndexes(u.rowIndex, u.columnIndex, rows, cols);
      r.load('values');
      return r;
    });
    await ctx.sync();

    const parts = [`活頁簿：${wb.name}`, `工作表數：${sheets.items.length}`, ''];
    sheets.items.forEach((s, i) => {
      const u = used[i];
      const hidden = s.visibility !== 'Visible' ? '（隱藏）' : '';
      if (u.isNullObject) {
        parts.push(`## ${s.name}${hidden}\n（空白工作表）\n`);
        return;
      }
      const truncCols = u.columnCount > 40 ? `，預覽只顯示前 40 欄` : '';
      parts.push(
        `## ${s.name}${hidden}`,
        `已使用範圍：${bareAddress(u.address)}（${u.rowCount} 列 × ${u.columnCount} 欄）`,
        `前 ${Math.min(previewRows, u.rowCount)} 列預覽${truncCols}：`,
        toCsv(previews[i].values),
        '',
      );
    });
    return parts.join('\n');
  });
}

async function getSelection() {
  return Excel.run(async (ctx) => {
    const r = ctx.workbook.getSelectedRange();
    r.load('address,rowCount,columnCount,values,formulas');
    await ctx.sync();
    const cells = r.rowCount * r.columnCount;
    if (cells > MAX_READ_CELLS) {
      return `使用者選取了 ${r.address}（${r.rowCount} 列 × ${r.columnCount} 欄），太大無法整段回傳。`
        + `請用 excel_read_range 讀其中一部分。`;
    }
    return `使用者目前選取：${r.address}（${r.rowCount} 列 × ${r.columnCount} 欄）\n\n值：\n${toCsv(r.values)}`;
  });
}

async function readRange(args) {
  const sheetName = String(args.sheet || '');
  const mode = args.mode || 'values';

  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
    sheet.load('isNullObject,name');
    await ctx.sync();
    if (sheet.isNullObject) {
      const all = ctx.workbook.worksheets;
      all.load('items/name');
      await ctx.sync();
      return `錯誤：找不到工作表「${sheetName}」。現有的工作表：${all.items.map(s => s.name).join('、')}`;
    }

    // Size first, contents second. Loading a huge range's values is the slow and
    // expensive part — we want to know we're about to do it before we do.
    let range;
    if (args.range) {
      range = sheet.getRange(String(args.range));
    } else {
      range = sheet.getUsedRangeOrNullObject(true);
    }
    range.load('isNullObject,address,rowIndex,columnIndex,rowCount,columnCount');
    await ctx.sync();

    if (range.isNullObject) return `工作表「${sheetName}」是空的。`;

    let rowCount = range.rowCount;
    let truncated = false;
    const maxRows = Math.max(1, Math.floor(MAX_READ_CELLS / Math.max(1, range.columnCount)));
    if (rowCount > maxRows) { rowCount = maxRows; truncated = true; }

    const target = sheet.getRangeByIndexes(range.rowIndex, range.columnIndex, rowCount, range.columnCount);
    if (mode === 'values' || mode === 'both') target.load('values');
    if (mode === 'formulas' || mode === 'both') target.load('formulas');
    target.load('address');
    await ctx.sync();

    const head = [`工作表 ${sheetName}，範圍 ${bareAddress(target.address)}`
      + `（${rowCount} 列 × ${range.columnCount} 欄）`];
    if (truncated) {
      head.push(`⚠ 原始範圍有 ${range.rowCount} 列，已截斷為前 ${rowCount} 列。`
        + `需要後面的資料請指定明確的 range 分批讀取。`);
    }
    const body = [];
    if (mode === 'values' || mode === 'both') body.push('### 值\n' + toCsv(target.values));
    if (mode === 'formulas' || mode === 'both') body.push('### 公式\n' + toCsv(target.formulas));
    return head.join('\n') + '\n\n' + body.join('\n\n');
  });
}

/**
 * Locate cells by content.
 *
 * Uses Excel's own findAll where available (fast, native) and falls back to a
 * capped JS scan when the host build doesn't support it — the fallback matters
 * because findAll needs a newer requirement set than the SharedRuntime floor we
 * ship against, so some supported desktops won't have it.
 */
async function search(args) {
  const query = String(args.query || '');
  if (!query) return '錯誤：query 為必填。';
  const limit = Math.min(Math.max(1, args.max_results || MAX_SEARCH_RESULTS), MAX_SEARCH_RESULTS);
  const onlySheet = args.sheet ? String(args.sheet) : null;

  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load('items/name');
    await ctx.sync();

    const targets = onlySheet
      ? sheets.items.filter(s => s.name === onlySheet)
      : sheets.items;
    if (!targets.length) return `錯誤：找不到工作表「${onlySheet}」。`;

    const hits = [];
    for (const sheet of targets) {
      if (hits.length >= limit) break;
      try {
        const found = sheet.getUsedRangeOrNullObject(true)
          .findAllOrNullObject(query, { completeMatch: false, matchCase: false });
        found.load('isNullObject,areas/items/address,areas/items/values');
        await ctx.sync();
        if (found.isNullObject) continue;
        for (const area of found.areas.items) {
          for (let r = 0; r < area.values.length; r++) {
            for (let c = 0; c < area.values[r].length; c++) {
              if (hits.length >= limit) break;
              hits.push({ sheet: sheet.name, address: bareAddress(area.address), value: area.values[r][c] });
            }
          }
        }
      } catch {
        // findAll unavailable on this build — scan the used range ourselves.
        const u = sheet.getUsedRangeOrNullObject(true);
        u.load('isNullObject,rowIndex,columnIndex,rowCount,columnCount');
        await ctx.sync();
        if (u.isNullObject) continue;
        const rows = Math.min(u.rowCount, Math.floor(MAX_READ_CELLS / Math.max(1, u.columnCount)));
        const block = sheet.getRangeByIndexes(u.rowIndex, u.columnIndex, rows, u.columnCount);
        block.load('values');
        await ctx.sync();
        const needle = query.toLowerCase();
        for (let r = 0; r < block.values.length && hits.length < limit; r++) {
          for (let c = 0; c < block.values[r].length && hits.length < limit; c++) {
            const v = block.values[r][c];
            if (v !== null && v !== '' && String(v).toLowerCase().includes(needle)) {
              hits.push({ sheet: sheet.name, address: a1(u.rowIndex + r, u.columnIndex + c), value: v });
            }
          }
        }
      }
    }

    if (!hits.length) return `在${onlySheet ? `工作表「${onlySheet}」` : '整份活頁簿'}裡找不到含有「${query}」的儲存格。`;
    const lines = hits.map(h => `${h.sheet}!${h.address}\t${csvCell(h.value)}`);
    const capped = hits.length >= limit ? `\n（已達 ${limit} 筆上限，可能還有更多）` : '';
    return `找到 ${hits.length} 筆：\n工作表!位址\t內容\n${lines.join('\n')}${capped}`;
  });
}

/**
 * Look at a write target WITHOUT writing — powers the confirmation dialog.
 *
 * Counting the non-empty cells about to be overwritten is the whole point: it is
 * what turns "允許這次寫入？" into "這會蓋掉 14 格既有資料", which is the
 * difference between a dialog people read and one they click through.
 */
/**
 * The confirmation card's look-before-you-ask step. One entry point taking the
 * tool name, because the shell only knows names — WHICH calls can be costed, and
 * how, is the host's business and nobody else's.
 */
export async function inspectWrite(tool, args) {
  if (tool === 'excel_clean_data') return inspectClean(args);
  return inspectWriteRange(args);
}

/** What excel_write_range is about to overwrite. */
async function inspectWriteRange(args) {
  const grid = args.formulas || args.values || [];
  const rows = grid.length;
  const cols = rows ? grid[0].length : 0;

  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItemOrNullObject(String(args.sheet || ''));
    sheet.load('isNullObject');
    await ctx.sync();
    if (sheet.isNullObject) return { ok: false, error: `找不到工作表「${args.sheet}」` };

    const target = sheet.getRange(String(args.start_cell)).getResizedRange(rows - 1, cols - 1);
    target.load('address,values,formulas');
    await ctx.sync();

    let occupied = 0, hadFormula = 0;
    for (const row of target.values) for (const v of row) if (v !== '' && v !== null) occupied++;
    for (const row of target.formulas) for (const f of row) if (typeof f === 'string' && f.startsWith('=')) hadFormula++;

    return {
      ok: true,
      address: target.address,
      rows, cols,
      cells: rows * cols,
      occupied,
      hadFormula,
      kind: args.formulas ? 'formulas' : 'values',
    };
  });
}

/**
 * Write, after snapshotting what was there.
 *
 * The snapshot is not optional bookkeeping — Office.js writes cannot be relied on
 * to land in Excel's own undo stack, so without this the user has no way back
 * from a bad write. Restoring `formulas` recovers constants and formulas alike.
 */
async function writeRange(args) {
  const grid = args.formulas || args.values;
  const asFormulas = !!args.formulas;
  const rows = grid.length;
  const cols = grid[0].length;
  const sheetName = String(args.sheet || '');

  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
    sheet.load('isNullObject');
    await ctx.sync();
    if (sheet.isNullObject) throw new Error(`找不到工作表「${sheetName}」`);

    const target = sheet.getRange(String(args.start_cell)).getResizedRange(rows - 1, cols - 1);
    target.load('address,formulas,numberFormat,rowIndex,columnIndex');
    await ctx.sync();

    const snapshot = {
      sheet: sheetName,
      address: bareAddress(target.address),
      formulas: target.formulas.map(r => r.slice()),
      numberFormat: target.numberFormat.map(r => r.slice()),
    };

    if (asFormulas) target.formulas = grid;
    else target.values = grid;
    await ctx.sync();

    // Grow whatever the new content no longer fits in. Widen-only, which is what
    // makes it safe on every single write: it can remove ### and clipped text, and
    // it cannot undo a layout somebody chose. Purely cosmetic, so a failure here
    // must not fail a write that already landed.
    let fitted = null;
    try {
      fitted = await fitLayout(ctx, [{
        sheet,
        rowIndex: target.rowIndex,
        columnIndex: target.columnIndex,
        rowCount: rows,
        columnCount: cols,
      }]);
    } catch { /* 版面沒調到不影響已寫入的資料 */ }

    return {
      text: `已寫入 ${sheetName}!${bareAddress(target.address)}（${rows} 列 × ${cols} 欄，${asFormulas ? '公式' : '值'}）。`
        + describeFit(fitted),
      snapshot,
    };
  });
}

// ── Structure, formatting, charts ─────────────────────────────────────────────

/** Resolve "工作表!A1:C9" or a bare "A1:C9" against a default sheet. */
function resolveRange(ctx, defaultSheet, ref) {
  const i = String(ref).lastIndexOf('!');
  if (i < 0) return ctx.workbook.worksheets.getItem(defaultSheet).getRange(String(ref));
  const name = String(ref).slice(0, i).replace(/^'|'$/g, '');
  return ctx.workbook.worksheets.getItem(name).getRange(String(ref).slice(i + 1));
}

async function manageSheet(args) {
  const op = String(args.op || '');
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;

    if (op === 'add') {
      const name = String(args.name || '').trim();
      if (name) {
        const dup = sheets.getItemOrNullObject(name);
        dup.load('isNullObject');
        await ctx.sync();
        if (!dup.isNullObject) return `錯誤：已經有一張叫「${name}」的工作表了。改個名字，或直接寫進現有那張。`;
      }
      const s = name ? sheets.add(name) : sheets.add();
      if (typeof args.position === 'number') s.position = args.position;
      s.activate();
      s.load('name,position');
      await ctx.sync();
      return `已新增工作表「${s.name}」並切換過去（位置 ${s.position}）。`;
    }

    const target = sheets.getItemOrNullObject(String(args.name || ''));
    target.load('isNullObject,name');
    await ctx.sync();
    if (target.isNullObject) {
      sheets.load('items/name');
      await ctx.sync();
      return `錯誤：找不到工作表「${args.name}」。現有的：${sheets.items.map(s => s.name).join('、')}`;
    }

    if (op === 'rename') {
      const from = target.name;
      target.name = String(args.new_name || '');
      await ctx.sync();
      return `已把「${from}」改名為「${args.new_name}」。`;
    }
    if (op === 'delete') {
      const n = target.name;
      target.delete();
      await ctx.sync();
      return `已刪除工作表「${n}」。`;
    }
    if (op === 'activate') {
      target.activate();
      await ctx.sync();
      return `已切換到「${target.name}」。`;
    }
    if (op === 'move') {
      target.position = Number(args.position) || 0;
      await ctx.sync();
      return `已把「${target.name}」移到位置 ${args.position}。`;
    }
    return `錯誤：未知的 op「${op}」。`;
  });
}

/**
 * Excel's UI shows column width in CHARACTERS (default 8.43). Office.js
 * `format.columnWidth` is in POINTS (default ~64). Everyone — including the
 * model — thinks in the UI's units, so `column_width` is accepted as characters
 * and converted here.
 *
 * Getting this wrong is not subtle: "set the width to 12" meant 12 points, i.e.
 * about 1.5 characters, which turns a table into one letter per line with very
 * tall rows. The tool description used to say 字元數 while the code passed the
 * number straight through as points.
 *
 * The formula is Excel's own: width = chars × (width of '0') + padding, which for
 * the default Calibri 11 is ≈ 7px per character plus 5px of cell padding.
 */
const charsToPoints = (chars) => Math.round(Number(chars) * 7 + 5);

/* ── fit:start ─ test/fit-layout.test.mjs slices the source from here ─ */

/** Floor and ceiling for anything that sets a column width, autofit included. */
const MIN_COL_POINTS = 30;    // ≈ 3.5 characters — narrower than this is unreadable
const MAX_COL_POINTS = 400;   // ≈ 56 characters — one long cell shouldn't span the screen

/** Same idea for heights. Row height is in points in Excel's own UI too, so no conversion. */
const MIN_ROW_POINTS = 15;    // Excel's default at Calibri 11
const MAX_ROW_POINTS = 180;   // ≈ 9 wrapped lines

/**
 * How much of one write gets measured. Autofit is a round trip to Excel per
 * column, so an unbounded sweep over a 50 000-row sheet is seconds of freeze
 * spent on columns nobody is looking at. Generous for the report-shaped output
 * this add-in produces, cheap for everything else.
 */
const MAX_FIT_COLUMNS = 60;
const MAX_FIT_ROWS = 200;
/** Sheets swept after a script. A workbook with more than this is not a report. */
const MAX_FIT_SHEETS = 8;

/**
 * Size columns and rows to the content that is in them.
 *
 * This exists because of what a generated workbook actually looked like. The
 * model builds a report with excel_run_script, reaches for
 * `format.columnWidth = 12` meaning 12 CHARACTERS, Office.js reads it as 12
 * POINTS — about a character and a half — and the finished report is a wall of
 * `####` in columns sixteen pixels wide. charsToPoints above fixes that for
 * excel_format_range, but a raw script bypasses that path entirely, and no
 * wording in a tool description prevents it reliably. So the layout is repaired
 * afterwards, from the content, where the right answer is knowable instead of
 * guessed.
 *
 * `shrink: false` — the default — only ever GROWS a column or row, and that is
 * the whole reason this is safe to run automatically on every write: a wide
 * column, a tall banner row, any layout a person tuned by hand cannot be taken
 * away by a write that happened to touch one cell. Only the unreadable direction
 * gets corrected. Hidden rows and columns are skipped for the same reason — a
 * width of zero is a decision, not a mistake.
 *
 * @param targets  [{ sheet, rowIndex, columnIndex, rowCount, columnCount }]
 * @returns {{columns: number, rows: number}} how many actually changed size
 */
async function fitLayout(ctx, targets, opts = {}) {
  const shrink = !!opts.shrink;
  const cols = [];
  const rows = [];

  for (const t of targets || []) {
    if (!t || !t.sheet) continue;
    const nc = Math.min(Number(t.columnCount) || 0, MAX_FIT_COLUMNS);
    for (let i = 0; i < nc; i++) {
      // The ENTIRE column, not the block that was written: autofitColumns()
      // measures only the cells inside the range it is called on, so autofitting
      // a header row alone sizes every column to its 2-3 character title and the
      // numbers underneath collapse into ###. Whole column is what Excel's own
      // 「自動調整欄寬」 means, so that is what this does.
      cols.push(t.sheet.getRangeByIndexes(0, t.columnIndex + i, 1, 1).getEntireColumn());
    }
    // All-or-nothing on rows, rather than the first 200 of a long block: a big
    // data dump has uniform automatic heights and nothing to fix, while the
    // hand-built report where a height WAS set by hand is always small.
    const nr = Number(t.rowCount) || 0;
    if (nr > 0 && nr <= MAX_FIT_ROWS) {
      for (let i = 0; i < nr; i++) {
        rows.push(t.sheet.getRangeByIndexes(t.rowIndex + i, 0, 1, 1).getEntireRow());
      }
    }
  }
  if (!cols.length && !rows.length) return { columns: 0, rows: 0 };

  for (const c of cols) c.load('format/columnWidth,columnHidden');
  for (const r of rows) r.load('format/rowHeight,rowHidden');
  await ctx.sync();
  const wasWidth = cols.map(c => c.format.columnWidth);
  const wasHeight = rows.map(r => r.format.rowHeight);
  const colHidden = cols.map(c => !!c.columnHidden);
  const rowHidden = rows.map(r => !!r.rowHidden);

  // Columns first, always. How tall a wrapped row needs to be depends on how wide
  // its columns ended up, so measuring rows first measures a layout that is about
  // to change — which is the "one row is half the screen" look.
  const columns = await fitAxis(ctx, cols, wasWidth, colHidden, {
    autofit: (x) => x.format.autofitColumns(),
    read: (x) => x.format.columnWidth,
    write: (x, v) => { x.format.columnWidth = v; },
    prop: 'format/columnWidth',
    min: MIN_COL_POINTS,
    max: MAX_COL_POINTS,
    shrink,
  });
  const fittedRows = await fitAxis(ctx, rows, wasHeight, rowHidden, {
    autofit: (x) => x.format.autofitRows(),
    read: (x) => x.format.rowHeight,
    write: (x, v) => { x.format.rowHeight = v; },
    prop: 'format/rowHeight',
    min: MIN_ROW_POINTS,
    max: MAX_ROW_POINTS,
    shrink,
  });
  return { columns, rows: fittedRows };
}

/**
 * One axis of fitLayout: autofit, measure what that gave, then correct whatever
 * came back out of bounds — including putting a size back when autofit shrank
 * something it was not allowed to shrink.
 */
async function fitAxis(ctx, items, was, hidden, spec) {
  if (!items.length) return 0;

  for (let i = 0; i < items.length; i++) if (!hidden[i]) spec.autofit(items[i]);
  await ctx.sync();

  for (const it of items) it.load(spec.prop);
  await ctx.sync();

  let changed = 0;
  const fixes = [];
  for (let i = 0; i < items.length; i++) {
    if (hidden[i]) continue;
    const before = was[i];
    const fitted = spec.read(items[i]);
    if (typeof before !== 'number' || typeof fitted !== 'number' || before <= 0) continue;
    // Autofit alone is not enough at either end. Excel refuses to widen a column
    // whose cells wrap — it grows the row instead — so a wrapped column comes back
    // at whatever narrow width it already had; and a column holding one
    // 500-character note autofits to something absurd. Hence the clamp.
    let want = Math.min(spec.max, Math.max(spec.min, fitted));
    if (!spec.shrink && before > want) want = before;
    if (Math.abs(want - fitted) > 0.5) fixes.push([items[i], want]);
    if (Math.abs(want - before) > 0.5) changed++;
  }
  for (const [it, v] of fixes) spec.write(it, v);
  if (fixes.length) await ctx.sync();
  return changed;
}

/** The half-sentence a tool result adds when fitLayout actually moved something. */
function describeFit(fitted) {
  if (!fitted) return '';
  const parts = [];
  if (fitted.columns) parts.push(`${fitted.columns} 欄的欄寬`);
  if (fitted.rows) parts.push(`${fitted.rows} 列的列高`);
  return parts.length ? `已依內容自動調整${parts.join('、')}。` : '';
}

/* ── fit:end ──────────────────────────────────────── */

const BORDER_EDGES = {
  all: ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight', 'InsideHorizontal', 'InsideVertical'],
  outline: ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'],
  bottom: ['EdgeBottom'],
  top: ['EdgeTop'],
};

/**
 * Apply only the properties that were supplied — everything omitted keeps its
 * current value. Deliberately NOT snapshotted for undo: formatting can't lose
 * data, and a half-correct format restore is worse than none.
 */
async function formatRange(args) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItemOrNullObject(String(args.sheet || ''));
    sheet.load('isNullObject');
    await ctx.sync();
    if (sheet.isNullObject) return `錯誤：找不到工作表「${args.sheet}」。`;

    const r = sheet.getRange(String(args.range));
    // Needed by the autofit clamp below, which walks the range column by column.
    const range = r;
    range.load('rowIndex,columnIndex,columnCount');
    await ctx.sync();

    const applied = [];

    if (args.number_format !== undefined) { r.numberFormat = [[args.number_format]]; applied.push('數字格式'); }
    if (args.bold !== undefined) { r.format.font.bold = !!args.bold; applied.push(args.bold ? '粗體' : '取消粗體'); }
    if (args.italic !== undefined) { r.format.font.italic = !!args.italic; applied.push('斜體'); }
    if (args.font_name !== undefined) { r.format.font.name = String(args.font_name); applied.push('字體'); }
    if (args.indent !== undefined) { r.format.indentLevel = Math.min(15, Math.max(0, Number(args.indent))); applied.push('縮排'); }
    if (args.font_size !== undefined) { r.format.font.size = Number(args.font_size); applied.push('字級'); }
    if (args.font_color !== undefined) { r.format.font.color = String(args.font_color); applied.push('字色'); }
    if (args.fill_color !== undefined) {
      if (String(args.fill_color).toLowerCase() === 'none') r.format.fill.clear();
      else r.format.fill.color = String(args.fill_color);
      applied.push('底色');
    }
    if (args.h_align !== undefined) { r.format.horizontalAlignment = String(args.h_align); applied.push('水平對齊'); }
    if (args.v_align !== undefined) { r.format.verticalAlignment = String(args.v_align); applied.push('垂直對齊'); }
    if (args.wrap_text !== undefined) { r.format.wrapText = !!args.wrap_text; applied.push('自動換行'); }
    if (args.column_width !== undefined) {
      const pts = Math.min(MAX_COL_POINTS, Math.max(MIN_COL_POINTS, charsToPoints(args.column_width)));
      r.format.columnWidth = pts;
      applied.push('欄寬');
    }
    if (args.row_height !== undefined) { r.format.rowHeight = Number(args.row_height); applied.push('列高'); }

    if (args.borders !== undefined) {
      const edges = BORDER_EDGES[String(args.borders)] || [];
      if (String(args.borders) === 'none') {
        for (const e of BORDER_EDGES.all) r.format.borders.getItem(e).style = 'None';
      } else {
        for (const e of edges) {
          const b = r.format.borders.getItem(e);
          // Style must be set before colour/weight — on a border that is currently
          // 'None' the other properties have nothing to apply to and are dropped.
          b.style = String(args.border_style || 'Continuous');
          b.weight = String(args.border_weight || 'Thin');
          if (args.border_color) b.color = String(args.border_color);
        }
      }
      applied.push('框線');
    }

    if (args.merge === true) { r.merge(false); applied.push('合併'); }
    else if (args.merge === false) { r.unmerge(); applied.push('取消合併'); }

    await ctx.sync();

    // Autofit runs after the content-affecting changes land, or it measures the
    // old text. `shrink: true` because this one was asked for explicitly — unlike
    // the automatic pass after a write, "自動調整欄寬" here means the real thing,
    // narrowing included.
    if (args.autofit_columns) {
      const fitted = await fitLayout(ctx, [{
        sheet,
        rowIndex: range.rowIndex,
        columnIndex: range.columnIndex,
        rowCount: 0,   // rows are handled by autofit_rows below, after the widths settle
        columnCount: range.columnCount,
      }], { shrink: true });
      applied.push(fitted.columns ? `自動欄寬（${fitted.columns} 欄）` : '自動欄寬');
    }

    // Rows last, and after the columns are final: row height depends on how wide
    // the columns ended up, so autofitting rows first measures the wrong layout.
    // Without this a wrapped column leaves rows stuck at the height they needed
    // when the column was narrow — the "one row is half the screen" look.
    if (args.autofit_rows) {
      r.getEntireRow().format.autofitRows();
      await ctx.sync();
      applied.push('自動列高');
    }

    return applied.length
      ? `已設定 ${args.sheet}!${args.range} 的${applied.join('、')}。`
      : '沒有指定任何要改的格式屬性。';
  });
}

async function createChart(args) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItemOrNullObject(String(args.sheet || ''));
    sheet.load('isNullObject');
    await ctx.sync();
    if (sheet.isNullObject) return `錯誤：找不到工作表「${args.sheet}」。`;

    const data = resolveRange(ctx, String(args.sheet), args.data_range);
    data.load('rowIndex,columnIndex,rowCount,columnCount,address');
    await ctx.sync();

    const chart = sheet.charts.add(String(args.chart_type), data, String(args.series_by || 'Auto'));
    if (args.title) { chart.title.text = String(args.title); chart.title.visible = true; }

    // Default placement: just to the right of the source data, so a chart never
    // lands on top of the numbers it came from.
    const anchor = args.anchor_cell
      ? sheet.getRange(String(args.anchor_cell))
      : sheet.getRangeByIndexes(data.rowIndex, data.columnIndex + data.columnCount + 1, 1, 1);
    chart.setPosition(anchor);
    chart.width = Number(args.width) || 360;
    chart.height = Number(args.height) || 220;

    // A default Excel chart carries a legend even for a single series, where it
    // just eats width and says nothing. Let the caller turn it off, label the
    // points directly, and pick a built-in style.
    if (args.legend_position !== undefined) {
      if (String(args.legend_position) === 'None') chart.legend.visible = false;
      else { chart.legend.visible = true; chart.legend.position = String(args.legend_position); }
    }
    if (args.show_data_labels) {
      chart.dataLabels.showValue = true;
    }
    if (args.chart_style !== undefined) {
      chart.style = Number(args.chart_style);
    }

    chart.load('name');
    await ctx.sync();
    return `已在「${args.sheet}」建立${args.chart_type}圖表「${chart.name}」，資料來源 ${bareAddress(data.address)}。`;
  });
}

const CLEAR_APPLY_TO = { contents: 'Contents', formats: 'Formats', all: 'All' };

async function structureOp(args) {
  const op = String(args.op || '');
  const count = Math.max(1, Number(args.count) || 1);
  const index = Math.max(1, Number(args.index) || 1);

  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItemOrNullObject(String(args.sheet || ''));
    sheet.load('isNullObject');
    await ctx.sync();
    if (sheet.isNullObject) return `錯誤：找不到工作表「${args.sheet}」。`;

    if (op === 'clear') {
      if (!args.range) return '錯誤：clear 需要 range。';
      sheet.getRange(String(args.range)).clear(CLEAR_APPLY_TO[String(args.clear_what || 'contents')] || 'Contents');
      await ctx.sync();
      return `已清除 ${args.sheet}!${args.range}。`;
    }

    // index is 1-based for humans; getRangeByIndexes is 0-based.
    const isRow = op.endsWith('_rows');
    const band = isRow
      ? sheet.getRangeByIndexes(index - 1, 0, count, 1).getEntireRow()
      : sheet.getRangeByIndexes(0, index - 1, 1, count).getEntireColumn();

    if (op === 'insert_rows') band.insert('Down');
    else if (op === 'delete_rows') band.delete('Up');
    else if (op === 'insert_columns') band.insert('Right');
    else if (op === 'delete_columns') band.delete('Left');
    else return `錯誤：未知的 op「${op}」。`;

    await ctx.sync();
    const unit = isRow ? '列' : '欄';
    const verb = op.startsWith('insert') ? '插入' : '刪除';
    return `已在「${args.sheet}」第 ${index} ${unit}${verb} ${count} ${unit}。`;
  });
}

async function tableOps(args) {
  const op = String(args.op || '');
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItemOrNullObject(String(args.sheet || ''));
    sheet.load('isNullObject');
    await ctx.sync();
    if (sheet.isNullObject) return `錯誤：找不到工作表「${args.sheet}」。`;

    const hasHeaders = args.has_headers !== false;

    if (op === 'sort') {
      if (!args.range) return '錯誤：sort 需要 range。';
      sheet.getRange(String(args.range)).sort.apply(
        [{ key: Number(args.sort_column) || 0, ascending: args.ascending !== false }],
        false, hasHeaders,
      );
      await ctx.sync();
      return `已依第 ${Number(args.sort_column) || 0} 欄${args.ascending === false ? '降冪' : '升冪'}排序 ${args.range}。`;
    }

    if (op === 'autofilter') {
      if (!args.range) return '錯誤：autofilter 需要 range。';
      sheet.autoFilter.apply(sheet.getRange(String(args.range)), Number(args.filter_column) || 0, {
        filterOn: 'Values',
        values: (args.filter_values || []).map(String),
      });
      await ctx.sync();
      return `已在 ${args.range} 第 ${Number(args.filter_column) || 0} 欄套用篩選。`;
    }

    if (op === 'create_table') {
      if (!args.range) return '錯誤：create_table 需要 range。';
      const t = sheet.tables.add(String(args.range), hasHeaders);
      if (args.table_name) t.name = String(args.table_name);
      // A built-in table style gives banded rows, a styled header and hover
      // highlighting in one property — far more consistent than painting cells.
      if (args.table_style) t.style = String(args.table_style);
      t.load('name,style');
      await ctx.sync();
      return `已把 ${args.range} 轉成表格「${t.name}」${args.table_style ? `，套用樣式 ${t.style}` : ''}。`;
    }

    if (op === 'freeze_panes') {
      const rows = Number(args.freeze_rows) || 0;
      const cols = Number(args.freeze_columns) || 0;
      if (rows && cols) sheet.freezePanes.freezeAt(sheet.getRangeByIndexes(0, 0, rows, cols));
      else if (rows) sheet.freezePanes.freezeRows(rows);
      else if (cols) sheet.freezePanes.freezeColumns(cols);
      else sheet.freezePanes.removeFrozenPanes();
      await ctx.sync();
      return `已凍結前 ${rows} 列、前 ${cols} 欄。`;
    }

    if (op === 'conditional_format') {
      if (!args.range) return '錯誤：conditional_format 需要 range。';
      const r = sheet.getRange(String(args.range));
      const type = String(args.format_type || 'ColorScale');

      if (type === 'ColorScale') {
        const cf = r.conditionalFormats.add('ColorScale');
        cf.colorScale.criteria = {
          minimum: { formula: null, type: 'LowestValue', color: '#F8696B' },
          midpoint: { formula: '50', type: 'Percentile', color: '#FFEB84' },
          maximum: { formula: null, type: 'HighestValue', color: '#63BE7B' },
        };
      } else if (type === 'DataBar') {
        r.conditionalFormats.add('DataBar');
      } else if (type === 'GreaterThan' || type === 'LessThan') {
        const cf = r.conditionalFormats.add('CellValue');
        cf.cellValue.rule = { formula1: String(args.threshold ?? 0), operator: type };
        cf.cellValue.format.font.color = type === 'GreaterThan' ? '#006100' : '#9C0006';
        cf.cellValue.format.fill.color = type === 'GreaterThan' ? '#C6EFCE' : '#FFC7CE';
      } else if (type === 'ContainsText') {
        const cf = r.conditionalFormats.add('ContainsText');
        cf.textComparison.rule = { operator: 'Contains', text: String(args.text || '') };
        cf.textComparison.format.fill.color = '#FFEB9C';
      } else if (type === 'TopBottom') {
        const cf = r.conditionalFormats.add('TopBottom');
        cf.topBottom.rule = { rank: Number(args.threshold) || 10, type: 'TopItems' };
        cf.topBottom.format.fill.color = '#C6EFCE';
      } else {
        return `錯誤：不支援的 format_type「${type}」。`;
      }
      await ctx.sync();
      return `已在 ${args.range} 套用${type}條件式格式。`;
    }

    return `錯誤：未知的 op「${op}」。`;
  });
}

/* ── clean:start ─ test/clean-data.test.mjs slices the source from here ─ */

/**
 * Reading a whole block into JS is cheap; the cap is about not hanging Excel on
 * an accidental A:XFD. Far above the 20 000-cell READ limit on purpose — those
 * cells go into the model's context, these never leave the pane.
 */
const MAX_CLEAN_CELLS = 200000;

/** Full-width ！ through ～ sit exactly 0xFEE0 above their ASCII twins. */
const toHalfWidth = (s) => s
  .replace(/[！-～]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
  .replace(/　/g, ' ');

/**
 * Text that means a number → a number. Handles what exported reports actually
 * contain: thousands separators, a currency mark, a trailing %, and accounting's
 * (1,234) for negatives.
 *
 * Returns null when the text is not a number — and null is also the answer for
 * '0012', because an ID with leading zeros is text and "cleaning" it to 12
 * destroys it. Same for anything past 15 digits, which is where IEEE doubles
 * start silently rounding account numbers.
 */
function textToNumber(raw) {
  const s = toHalfWidth(String(raw)).trim();
  if (!s) return null;
  const neg = /^\((.*)\)$/.exec(s);
  const body = neg ? neg[1] : s;
  const pct = /%$/.test(body);
  const bare = body.replace(/%$/, '').replace(/[,\s]/g, '').replace(/^(NT|[$¥€£])+/i, '');
  if (!/^-?\d+(\.\d+)?$/.test(bare)) return null;
  if (/^-?0\d/.test(bare)) return null;              // 0012 is an ID, not 12
  if (bare.replace(/[-.]/g, '').length > 15) return null;
  const n = Number(bare);
  if (!Number.isFinite(n)) return null;
  return (neg ? -1 : 1) * (pct ? n / 100 : n);
}

/** Excel's day 0 is 1899-12-30, which is how its 1900-leap-year bug shakes out. */
const EXCEL_EPOCH_UTC = Date.UTC(1899, 11, 30);

/**
 * Text that means a date → an Excel serial number.
 *
 * 民國年 is in here because it is the reason this exists: a government or ERP
 * export writes 113/1/5, Excel stores it as text, every date formula downstream
 * returns #VALUE!, and no amount of 儲存格格式 fixes it because the cell was never
 * a date. A year of three digits or fewer is read as 民國 and 1911 added.
 *
 * Deliberately narrow: y/m/d and y-m-d only. Nothing here guesses whether 3/4/24
 * is March or April — that is a question about the person's locale, not about
 * the data, and getting it wrong silently is worse than not converting at all.
 */
function textToDateSerial(raw) {
  const s = toHalfWidth(String(raw)).trim().replace(/[年月]/g, '/').replace(/日$/, '');
  const m = /^(\d{1,4})[/-](\d{1,2})[/-](\d{1,2})$/.exec(s);
  if (!m) return null;
  let y = Number(m[1]);
  const mo = Number(m[2]);
  const d = Number(m[3]);
  if (m[1].length <= 3) y += 1911;                   // 民國
  if (y < 1900 || y > 9999 || mo < 1 || mo > 12 || d < 1 || d > 31) return null;
  const t = Date.UTC(y, mo - 1, d);
  const back = new Date(t);
  // Rejects 2024/2/31, which Date.UTC would have rolled forward into March.
  if (back.getUTCMonth() !== mo - 1 || back.getUTCDate() !== d) return null;
  return Math.round((t - EXCEL_EPOCH_UTC) / 86400000);
}

const isBlank = (v) => v === '' || v === null || v === undefined;
const isFormula = (f) => typeof f === 'string' && f.startsWith('=');

/**
 * Work out what a clean-up would change — from the values alone, with no Office
 * object anywhere in it.
 *
 * Pure on purpose, and that is the whole design of this tool. The confirmation
 * card has to say 「會刪掉 37 列重複資料」 BEFORE anything happens, and PRD 3.1
 * requires that number to be honest. Computing it twice — once to show, once to
 * do — is exactly how the two drift apart. So the card and the edit call this
 * same function: the card renders the counts and throws the grid away, the edit
 * writes the grid. It also means every rule in here is testable without an Excel.
 *
 * @param values    what the block currently shows
 * @param formulas  the same block's formulas, so cells that compute are left alone
 * @returns {{grid, numberFormat, cells, rows, blocked, note}}
 *          grid          the block as it should end up (null = nothing to do)
 *          numberFormat  same shape, sparse; only where the op has an opinion
 *          cells / rows  how many cells changed / how many rows went away
 *          blocked       a sentence for the model when the op must not run at all
 */
function planClean(op, args, values, formulas) {
  const rows = values.length;
  const cols = rows ? values[0].length : 0;
  const out = values.map(r => r.slice());
  const nf = values.map(r => r.map(() => null));
  const hasHeaders = args.has_headers !== false;
  const first = hasHeaders ? 1 : 0;
  const nothing = { grid: null, numberFormat: null, cells: 0, rows: 0, blocked: null, note: '' };

  /** Which columns of the block this op may touch. Omitted = all of them. */
  const inScope = (c) => !Array.isArray(args.columns) || args.columns.length === 0
    || args.columns.map(Number).includes(c);

  if (op === 'dedupe') {
    // Moving a row up rewrites nothing inside its formulas, so every relative
    // reference in it would quietly start pointing at a different row. There is
    // no safe way to dedupe across that, and doing it anyway produces a sheet
    // that looks right and computes wrong — the worst outcome a cleaning tool
    // has available. Refuse, and say what to do instead.
    for (let r = first; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        if (isFormula(formulas[r][c])) {
          return {
            ...nothing,
            blocked: '這個範圍裡有公式。去重會把資料列往上搬，公式的相對參照就會指到別的列，'
              + '結果是看起來對但算錯。請先把公式貼成值，或改成只框選純資料的欄位再重試。',
          };
        }
      }
    }
    const keys = Array.isArray(args.key_columns) && args.key_columns.length
      ? args.key_columns.map(Number).filter(c => c >= 0 && c < cols)
      : Array.from({ length: cols }, (_, i) => i);
    const seen = new Set();
    const kept = values.slice(0, first);
    let removed = 0;
    for (let r = first; r < rows; r++) {
      // Compared as trimmed text: an export where one row says 'A001 ' and the
      // next says 'A001' has two rows that any person would call the same.
      // Joined on NUL because it is the one character a cell cannot hold — any
      // printable separator lets ['a b','c'] and ['a','b c'] collide into the
      // same key and silently drop a row that was never a duplicate.
      const k = keys.map(c => String(values[r][c] ?? '').trim()).join('\u0000');
      if (seen.has(k)) { removed++; continue; }
      seen.add(k);
      kept.push(values[r]);
    }
    if (!removed) return { ...nothing, note: '沒有重複的資料列。' };
    const grid = kept.map(r => r.slice());
    while (grid.length < rows) grid.push(new Array(cols).fill(''));
    return { grid, numberFormat: null, cells: 0, rows: removed, blocked: null, note: `剩下 ${kept.length - first} 列` };
  }

  if (op === 'normalize') {
    const trim = args.trim !== false;
    const collapse = args.collapse_spaces !== false;
    const half = !!args.half_width;
    const toNum = !!args.to_number;
    const toDate = !!args.to_date;
    const letterCase = String(args.letter_case || 'none');
    let changed = 0;
    for (let r = first; r < rows; r++) {
      for (let c = 0; c < cols; c++) {
        if (!inScope(c) || isFormula(formulas[r][c])) continue;
        const v = values[r][c];
        if (typeof v !== 'string' || v === '') continue;
        let s = v;
        if (half) s = toHalfWidth(s);
        if (trim) s = s.trim();
        if (collapse) s = s.replace(/\s+/g, ' ');
        if (letterCase === 'upper') s = s.toUpperCase();
        else if (letterCase === 'lower') s = s.toLowerCase();
        let value = s;
        if (toDate) {
          const serial = textToDateSerial(s);
          if (serial !== null) { value = serial; nf[r][c] = 'yyyy/m/d'; }
        }
        if (typeof value === 'string' && toNum) {
          const n = textToNumber(s);
          if (n !== null) value = n;
        }
        if (value !== v) { out[r][c] = value; changed++; }
      }
    }
    return changed
      ? { grid: out, numberFormat: nf, cells: changed, rows: 0, blocked: null, note: '' }
      : { ...nothing, note: '這個範圍裡沒有需要統一的寫法。' };
  }

  if (op === 'fill') {
    const method = String(args.method || 'down');
    let changed = 0;
    if (method === 'down') {
      // The merged-cell export: a category written once with the rows under it
      // left blank. Reads fine to a person, groups and pivots terribly.
      for (let c = 0; c < cols; c++) {
        if (!inScope(c)) continue;
        let carry = null;
        for (let r = first; r < rows; r++) {
          if (isFormula(formulas[r][c])) { carry = null; continue; }
          if (!isBlank(values[r][c])) { carry = values[r][c]; continue; }
          if (carry === null) continue;
          out[r][c] = carry;
          changed++;
        }
      }
    } else {
      const filler = args.fill_value === undefined ? 0 : args.fill_value;
      for (let r = first; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
          if (!inScope(c) || isFormula(formulas[r][c]) || !isBlank(values[r][c])) continue;
          out[r][c] = filler;
          changed++;
        }
      }
    }
    return changed
      ? { grid: out, numberFormat: null, cells: changed, rows: 0, blocked: null, note: '' }
      : { ...nothing, note: '這個範圍裡沒有空白格要補。' };
  }

  if (op === 'split') {
    if (cols < 2) {
      return {
        ...nothing,
        blocked: '拆欄的 range 要同時含來源欄和放結果的欄，例如 B2:D50 是把 B 拆進 B、C、D。'
          + '先確認右邊那幾欄是空的，或接受它們會被蓋掉。',
      };
    }
    const sep = String(args.delimiter ?? ',');
    const keep = !!args.keep_source;
    // keep_source: the source column stays as it was and the parts start one to
    // its right. Otherwise the source is replaced by its own first part, which is
    // what Excel's 資料剖析 does and what people expect.
    const firstOut = keep ? 1 : 0;
    const slots = cols - firstOut;
    let changed = 0;
    let overflow = 0;
    for (let r = first; r < rows; r++) {
      const v = values[r][0];
      if (isFormula(formulas[r][0]) || isBlank(v)) continue;
      const parts = String(v).split(sep).map(p => p.trim());
      if (parts.length > slots) overflow++;
      for (let i = 0; i < slots; i++) {
        // The tail is cleared as well as filled, or a shorter row leaves the
        // previous row's third part sitting next to it.
        const want = i < parts.length ? parts[i] : '';
        if (out[r][firstOut + i] !== want) { out[r][firstOut + i] = want; changed++; }
      }
    }
    if (!changed) return { ...nothing, note: `沒有任何一格含有分隔符號「${sep}」。` };
    return {
      grid: out,
      numberFormat: null,
      cells: changed,
      rows: 0,
      blocked: null,
      note: overflow ? `有 ${overflow} 列拆出來超過 ${slots} 欄，多的被丟掉了` : '',
    };
  }

  if (op === 'merge_columns') {
    if (cols < 2) return { ...nothing, blocked: '併欄的 range 至少要兩欄。' };
    const sep = String(args.separator ?? '');
    const keep = !!args.keep_source;
    let changed = 0;
    for (let r = first; r < rows; r++) {
      const parts = [];
      let skip = false;
      for (let c = 0; c < cols; c++) {
        if (isFormula(formulas[r][c])) { skip = true; break; }
        const v = values[r][c];
        // Blank cells are skipped rather than joined, or a missing middle column
        // turns every row into 「台北市－－信義路」.
        if (!isBlank(v)) parts.push(String(v).trim());
      }
      if (skip || !parts.length) continue;
      const joined = parts.join(sep);
      if (out[r][0] !== joined) { out[r][0] = joined; changed++; }
      if (!keep) {
        for (let c = 1; c < cols; c++) {
          if (isBlank(out[r][c])) continue;
          out[r][c] = '';
          changed++;
        }
      }
    }
    return changed
      ? { grid: out, numberFormat: null, cells: changed, rows: 0, blocked: null, note: '' }
      : { ...nothing, note: '沒有需要合併的內容。' };
  }

  return { ...nothing, blocked: `未知的 op「${op}」。可用的是 dedupe、normalize、fill、split、merge_columns。` };
}

/* ── clean:end ──────────────────────────────────────── */

/** What the confirmation card says a clean-up is about to cost. */
const CLEAN_RISK = {
  dedupe: (p) => t('會刪掉 {0} 列重複資料（{1}）。', p.rows, p.note),
  normalize: (p) => t('會統一 {0} 格的寫法。', p.cells),
  fill: (p) => t('會補上 {0} 格空白。', p.cells),
  split: (p) => t('會改寫 {0} 格，原本在那裡的內容會被蓋掉。', p.cells),
  merge_columns: (p) => t('會改寫 {0} 格，原本在那裡的內容會被蓋掉。', p.cells),
};

/**
 * Load the block a clean op works on, and plan the change without touching it.
 * Shared by inspectClean (show it) and cleanData (do it).
 */
async function readCleanBlock(ctx, args) {
  const sheet = ctx.workbook.worksheets.getItemOrNullObject(String(args.sheet || ''));
  sheet.load('isNullObject');
  await ctx.sync();
  if (sheet.isNullObject) return { error: `找不到工作表「${args.sheet}」` };

  // OrNullObject, not getUsedRange: on an empty sheet the plain one throws, and
  // 「這張表是空的」 is an answer the model can act on, not an error to report.
  const target = args.range ? sheet.getRange(String(args.range)) : sheet.getUsedRangeOrNullObject(true);
  target.load('isNullObject,address,rowCount,columnCount,values,formulas,numberFormat');
  await ctx.sync();
  if (target.isNullObject) return { error: `工作表「${args.sheet}」是空的，沒有東西可以清理。` };

  const cells = target.rowCount * target.columnCount;
  if (cells > MAX_CLEAN_CELLS) {
    return { error: `範圍太大（${cells} 格，上限 ${MAX_CLEAN_CELLS} 格）。請分批處理。` };
  }
  return { sheet, target, plan: planClean(String(args.op || ''), args, target.values, target.formulas) };
}

/** Look at a clean-up WITHOUT doing it — powers the confirmation card. */
async function inspectClean(args) {
  return Excel.run(async (ctx) => {
    const b = await readCleanBlock(ctx, args);
    if (b.error) return { ok: false, error: b.error };
    if (b.plan.blocked) return { ok: false, error: b.plan.blocked };

    const risk = CLEAN_RISK[String(args.op || '')];
    return {
      ok: true,
      address: bareAddress(b.target.address),
      // No grid means the op found nothing to change. Saying that on the card is
      // better than an empty confirmation the person has to puzzle over.
      risk: b.plan.grid && risk ? risk(b.plan) : (b.plan.note || t('這個範圍沒有需要清理的地方。')),
    };
  });
}

/**
 * 資料清理 —— 去重、統一格式、補齊缺漏、拆欄、併欄。
 *
 * Every op rewrites ONE rectangle and never inserts or deletes anything, which
 * is what makes the existing range snapshot enough to undo all five: put the old
 * formulas and number formats back over the same address and the sheet is exactly
 * as it was. Dedupe compacts its rows upward inside the block and blanks the tail
 * rather than deleting rows, for that same reason — a real row delete drags
 * everything below the block up with it, and no range snapshot can put that back.
 */
async function cleanData(args) {
  return Excel.run(async (ctx) => {
    const b = await readCleanBlock(ctx, args);
    if (b.error) return { text: `錯誤：${b.error}` };
    if (b.plan.blocked) return { text: `錯誤：${b.plan.blocked}` };
    if (!b.plan.grid) return { text: b.plan.note || '這個範圍沒有需要清理的地方，什麼都沒改。' };

    const address = bareAddress(b.target.address);
    const snapshot = {
      sheet: String(args.sheet || ''),
      address,
      formulas: b.target.formulas.map(r => r.slice()),
      numberFormat: b.target.numberFormat.map(r => r.slice()),
    };

    // numberFormat merged rather than replaced: writing the plan's grid over the
    // whole block would reset every currency and percentage in it to General.
    if (b.plan.numberFormat) {
      b.target.numberFormat = b.target.numberFormat
        .map((row, r) => row.map((cur, c) => b.plan.numberFormat[r][c] || cur));
    }
    b.target.values = b.plan.grid;
    await ctx.sync();

    const done = b.plan.rows ? `已刪掉 ${b.plan.rows} 列重複資料` : `已清理 ${b.plan.cells} 格`;
    return {
      text: `${done}${b.plan.note ? `（${b.plan.note}）` : ''}，範圍 ${args.sheet}!${address}。`,
      snapshot,
    };
  });
}

/**
 * Sheet-level appearance.
 *
 * `showGridlines = false` is the one that matters. Those grey lines are the
 * single strongest signal that something is "a spreadsheet"; with them off, the
 * fills and borders the agent draws stop reading as cell decoration and start
 * reading as layout. Every Excel dashboard and calendar that looks designed
 * turns them off first.
 *
 * All of this is ExcelApi 1.7 — well below the SharedRuntime floor the manifest
 * already requires — but each property is set independently so an unexpectedly
 * old host degrades property by property instead of failing the whole call.
 */
async function sheetStyle(args) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItemOrNullObject(String(args.sheet || ''));
    sheet.load('isNullObject');
    await ctx.sync();
    if (sheet.isNullObject) return `錯誤：找不到工作表「${args.sheet}」。`;

    const applied = [];
    if (args.show_gridlines !== undefined) {
      sheet.showGridlines = !!args.show_gridlines;
      applied.push(args.show_gridlines ? '顯示格線' : '關閉格線');
    }
    if (args.show_headings !== undefined) {
      sheet.showHeadings = !!args.show_headings;
      applied.push(args.show_headings ? '顯示欄列標題' : '隱藏欄列標題');
    }
    if (args.tab_color !== undefined) {
      sheet.tabColor = String(args.tab_color);
      applied.push('分頁顏色');
    }
    if (args.standard_width !== undefined) {
      sheet.standardWidth = charsToPoints(args.standard_width);
      applied.push('預設欄寬');
    }
    if (!applied.length) return '沒有指定任何要改的項目。';

    await ctx.sync();
    return `已設定工作表「${args.sheet}」：${applied.join('、')}。`;
  });
}

// ── Pivot tables ──────────────────────────────────────────────────────────────

const PIVOT_AREAS = [
  ['rows', 'rowHierarchies'],
  ['columns', 'columnHierarchies'],
  ['filters', 'filterHierarchies'],
];

async function createPivot(args) {
  if (!hasApi('1.12')) {
    return '錯誤：這個 Excel 版本不支援用程式建立樞紐分析表（需要 ExcelApi 1.12）。'
      + '改用 excel_write_range 直接寫一份算好的彙總表給使用者。';
  }

  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItemOrNullObject(String(args.sheet || ''));
    sheet.load('isNullObject');
    await ctx.sync();
    if (sheet.isNullObject) return `錯誤：找不到工作表「${args.sheet}」。`;

    const existing = ctx.workbook.pivotTables.getItemOrNullObject(String(args.name));
    existing.load('isNullObject');
    await ctx.sync();
    if (!existing.isNullObject) return `錯誤：已經有一個叫「${args.name}」的樞紐分析表，換個名字。`;

    const source = resolveRange(ctx, String(args.sheet), args.source_range);
    const dest = sheet.getRange(String(args.destination_cell));
    const pivot = sheet.pivotTables.add(String(args.name), source, dest);

    // Field names come from the source header row. Load them first so a typo gets
    // a list of what IS available instead of a bare ItemNotFound.
    pivot.hierarchies.load('items/name');
    await ctx.sync();
    const fields = pivot.hierarchies.items.map(h => h.name);

    const wanted = [
      ...PIVOT_AREAS.flatMap(([key]) => (args[key] || [])),
      ...((args.values || []).map(v => v.field)),
    ].map(String);
    const unknown = wanted.filter(f => !fields.includes(f));
    if (unknown.length) {
      pivot.delete();
      await ctx.sync();
      return `錯誤：來源資料裡沒有「${unknown.join('、')}」這些欄位。可用的欄位：${fields.join('、')}`;
    }

    for (const [key, collection] of PIVOT_AREAS) {
      for (const f of args[key] || []) pivot[collection].add(pivot.hierarchies.getItem(String(f)));
    }
    for (const v of args.values || []) {
      const dh = pivot.dataHierarchies.add(pivot.hierarchies.getItem(String(v.field)));
      dh.summarizeBy = v.summarize_by || 'Sum';
    }
    await ctx.sync();

    const parts = PIVOT_AREAS
      .map(([key]) => (args[key] || []).length ? `${key}：${args[key].join('、')}` : null)
      .filter(Boolean);
    if ((args.values || []).length) {
      parts.push(`值：${args.values.map(v => `${v.field}(${v.summarize_by || 'Sum'})`).join('、')}`);
    }
    return `已在 ${args.sheet}!${args.destination_cell} 建立樞紐分析表「${args.name}」。${parts.join('；')}`;
  });
}

// ── Formula precedents ────────────────────────────────────────────────────────

/**
 * Pull cell references out of a formula string.
 *
 * The fallback for hosts without ExcelApi 1.12. Only finds DIRECT references and
 * can't follow named ranges, so it says so in the output rather than quietly
 * pretending to be the real thing. The lookbehind stops it matching the "G10"
 * inside a function name like LOG10.
 */
const FORMULA_REF = /(?<![A-Za-z0-9_$])(?:('[^']+'|[A-Za-z_一-鿿][\w一-鿿.]*)!)?(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?)(?![\w(])/g;

function parseFormulaRefs(formula, defaultSheet) {
  const out = [];
  for (const m of String(formula).matchAll(FORMULA_REF)) {
    const sheet = (m[1] || defaultSheet).replace(/^'|'$/g, '');
    const addr = m[2];
    const key = `${sheet}!${addr}`;
    if (!out.some(r => r.key === key)) out.push({ key, sheet, addr });
  }
  return out;
}

const MAX_PRECEDENTS = 30;

async function tracePrecedents(args) {
  const sheetName = String(args.sheet || '');
  const cell = String(args.cell || '');
  const all = args.mode === 'all';

  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
    sheet.load('isNullObject');
    await ctx.sync();
    if (sheet.isNullObject) return `錯誤：找不到工作表「${sheetName}」。`;

    const target = sheet.getRange(cell);
    target.load('address,values,formulas,numberFormat');
    await ctx.sync();

    const formula = target.formulas[0][0];
    const value = target.values[0][0];
    const head = `${sheetName}!${bareAddress(target.address)} 目前的值是 ${JSON.stringify(value)}`;

    if (typeof formula !== 'string' || !formula.startsWith('=')) {
      return `${head}，它是直接輸入的常數，沒有公式，也就沒有前導參照。`;
    }

    // getPrecedents (all levels) needs 1.14; getDirectPrecedents needs 1.12.
    // Both throw ItemNotFound when a formula references nothing (e.g. =TODAY()).
    const canAll = hasApi('1.14');
    const canDirect = hasApi('1.12');
    let addresses = null;
    let native = false;

    if ((all && canAll) || canDirect) {
      try {
        const areas = all && canAll ? target.getPrecedents() : target.getDirectPrecedents();
        areas.load('addresses');
        await ctx.sync();
        addresses = areas.addresses;
        native = true;
      } catch {
        addresses = [];
        native = true;
      }
    }

    let refs;
    let note = '';
    if (native) {
      refs = (addresses || []).slice(0, MAX_PRECEDENTS).map(a => {
        const i = a.lastIndexOf('!');
        return { sheet: i < 0 ? sheetName : a.slice(0, i).replace(/^'|'$/g, ''), addr: i < 0 ? a : a.slice(i + 1) };
      });
      if (all && !canAll) note = '（這個 Excel 版本只支援直接參照，沒有一路追到底）';
    } else {
      refs = parseFormulaRefs(formula, sheetName).slice(0, MAX_PRECEDENTS);
      note = '（這個 Excel 版本不支援原生追蹤，以下是直接從公式字串解析出來的，不含具名範圍與間接參照）';
    }

    if (!refs.length) return `${head}\n公式：${formula}\n這個公式沒有參照其他儲存格。`;

    // One batch load for every precedent, not one round trip each.
    const loaded = refs.map(r => {
      const s = ctx.workbook.worksheets.getItemOrNullObject(r.sheet);
      const rr = s.getRange(r.addr);
      rr.load('address,values,formulas');
      return { ref: r, sheet: s, range: rr };
    });
    loaded.forEach(l => l.sheet.load('isNullObject'));
    await ctx.sync();

    const lines = loaded.map(({ ref, sheet: s, range }) => {
      if (s.isNullObject) return `${ref.sheet}!${ref.addr}\t(找不到工作表)`;
      const v = range.values[0]?.[0];
      const f = range.formulas[0]?.[0];
      const isFormula = typeof f === 'string' && f.startsWith('=');
      return `${ref.sheet}!${ref.addr}\t${JSON.stringify(v)}\t${isFormula ? f : ''}`;
    });

    return `${head}\n公式：${formula}\n\n前導參照（${all ? '全部層級' : '直接'}）${note}：\n`
      + `位址\t值\t它自己的公式\n${lines.join('\n')}`
      + (addresses && addresses.length > MAX_PRECEDENTS
        ? `\n（共 ${addresses.length} 個，只列出前 ${MAX_PRECEDENTS} 個）` : '');
  });
}

/* ── snapshot:start ─ test/undo-snapshot.test.mjs slices the source from here ─ */

/**
 * Ceiling on a whole-workbook snapshot, in cells.
 *
 * Unlike MAX_READ_CELLS this is not about context — nothing here reaches the
 * model. It is about the seconds the person spends watching 執行中… after they
 * clicked 允許, and about holding two string arrays of that size in the pane.
 * Past this the snapshot is skipped and the pane says so, because a snapshot
 * that takes ten seconds to build would get blamed on the script.
 */
const MAX_SNAPSHOT_CELLS = 200000;

/**
 * Before-image of the whole workbook, for undoing a script.
 *
 * excel_write_range knows its target range and can snapshot exactly it. A script
 * does not: it can touch any sheet, and it decides which as it runs — by which
 * time the before-state is gone. The options are to have the model declare its
 * targets (it would be wrong sometimes, and wrong here means data with no way
 * back) or to photograph everything first. This does the second, bounded.
 *
 * What comes back is content, not appearance: formulas and number formats of
 * each sheet's used range. Fills, fonts and borders the script applies are NOT
 * undone — the same limit excel_write_range's snapshot has always had.
 *
 * Returns null when the workbook is too big to photograph. The caller must treat
 * that as "this change cannot be undone" and say so, not fail.
 */
async function captureWorkbook() {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load('items/name');
    await ctx.sync();

    const items = sheets.items;
    // valuesOnly: a sheet whose formatting reaches row 10000 but whose content
    // stops at row 40 should cost 40 rows, not 10000.
    const used = items.map((s) => {
      const r = s.getUsedRangeOrNullObject(true);
      r.load('address,rowCount,columnCount,isNullObject');
      return r;
    });
    await ctx.sync();

    let cells = 0;
    for (const r of used) if (!r.isNullObject) cells += r.rowCount * r.columnCount;
    if (cells > MAX_SNAPSHOT_CELLS) return null;

    for (const r of used) if (!r.isNullObject) r.load('formulas,numberFormat');
    await ctx.sync();

    return {
      kind: 'workbook',
      /** Sheet names before the script ran. */
      before: items.map((s) => s.name),
      /** Filled in by finishWorkbookSnapshot once the script is done. */
      after: null,
      sheets: items.map((s, i) => (used[i].isNullObject
        ? { name: s.name, address: null, formulas: null, numberFormat: null }
        : {
          name: s.name,
          address: bareAddress(used[i].address),
          formulas: used[i].formulas.map((r) => r.slice()),
          numberFormat: used[i].numberFormat.map((r) => r.slice()),
        })),
    };
  });
}

/**
 * Record the sheet list AFTER the script, so undo can delete what the script
 * added without touching what the person added afterwards.
 *
 * Comparing only against `before` cannot tell those apart: a sheet someone
 * created themselves five minutes later is equally absent from it, and undo
 * would delete their work. A sheet is the script's only if it is missing from
 * `before` AND present in `after`.
 *
 * Failing to read the list leaves `after` null, which makes undo delete nothing.
 * An incomplete undo is recoverable; deleting the wrong sheet is not.
 */
async function finishWorkbookSnapshot(snapshot) {
  if (!snapshot || snapshot.kind !== 'workbook') return snapshot;
  try {
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load('items/name');
      await ctx.sync();
      snapshot.after = sheets.items.map((s) => s.name);
    });
  } catch {
    snapshot.after = null;
  }
  return snapshot;
}

/**
 * Put a snapshot back. Restoring formulas covers both constants and formulas.
 *
 * Returns `{ label, notes }` already worded and translated, because the two
 * snapshot shapes undo very different amounts of ground and only this module
 * knows which — 「{0} 張工作表」 is not a sentence the shell should be able to
 * say, now that the shell also runs in Word.
 */
export async function restoreSnapshot(snapshot) {
  if (snapshot && snapshot.kind === 'workbook') return restoreWorkbook(snapshot);
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(snapshot.sheet);
    const target = sheet.getRange(snapshot.address);
    target.numberFormat = snapshot.numberFormat;
    target.formulas = snapshot.formulas;
    await ctx.sync();
    return {
      kind: 'range',
      sheet: snapshot.sheet,
      address: snapshot.address,
      label: `${snapshot.sheet}!${snapshot.address}`,
      notes: [],
    };
  });
}

/**
 * Undo a script: put every sheet's contents back, remove what it added, rebuild
 * what it deleted.
 *
 * Clearing the current used range before writing the before-image back is what
 * makes this an undo rather than a partial one — a script that wrote past the
 * old used range (A1:F10 before, A1:F400 after) would otherwise leave rows
 * 11-400 sitting there. `clear('Contents')` and not a bare `clear()`, which is
 * clear-everything: formatting that was already on those cells is not this
 * snapshot's to throw away.
 *
 * Note there is no `Range.clearContents()` in Office.js however much the name
 * suggests itself — that spelling is Excel VBA's. It was what this used to call,
 * and every undo of a script that had written anything died on
 * 「current[i].clearContents is not a function」 while an undo of a script that
 * had only ADDED sheets went through, because that path never reached the call.
 */
async function restoreWorkbook(snapshot) {
  return Excel.run(async (ctx) => {
    const coll = ctx.workbook.worksheets;
    coll.load('items/name');
    await ctx.sync();
    const present = coll.items.map((s) => s.name);

    const deleted = (snapshot.after || [])
      .filter((n) => !snapshot.before.includes(n) && present.includes(n));
    for (const name of deleted) coll.getItem(name).delete();

    // One pass to find out which used ranges exist, so the clears below do not
    // need a sync each.
    const current = snapshot.sheets.map((s) => {
      if (!present.includes(s.name)) return null;
      const r = coll.getItem(s.name).getUsedRangeOrNullObject(true);
      r.load('isNullObject');
      return r;
    });
    await ctx.sync();

    const recreated = [];
    snapshot.sheets.forEach((s, i) => {
      const exists = present.includes(s.name);
      // A sheet the script deleted comes back with its contents but at the end
      // of the tab strip and without its formatting. Reported, not hidden.
      const sheet = exists ? coll.getItem(s.name) : coll.add(s.name);
      if (!exists) recreated.push(s.name);
      if (exists && current[i] && !current[i].isNullObject) current[i].clear('Contents');
      if (s.address) {
        const target = sheet.getRange(s.address);
        target.numberFormat = s.numberFormat;
        target.formulas = s.formulas;
      }
    });
    await ctx.sync();

    const notes = [];
    if (deleted.length) notes.push(t('已移除 AI 新增的工作表：{0}', deleted.join('、')));
    if (recreated.length) {
      notes.push(t('AI 刪掉的工作表已重建：{0}。內容回來了，格式沒有，位置在最後面。', recreated.join('、')));
    }
    return {
      kind: 'workbook',
      sheets: snapshot.sheets.length,
      deleted,
      recreated,
      label: t('{0} 張工作表的內容', snapshot.sheets.length),
      notes,
    };
  });
}

/* ── snapshot:end ──────────────────────────────────────── */

/** Scroll to and select a cell — powers the [[Sheet!A1]] citation links. */
export async function goTo(sheetName, address) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
    sheet.load('isNullObject');
    await ctx.sync();
    if (sheet.isNullObject) return false;
    sheet.activate();
    const r = sheet.getRange(address);
    r.select();
    await ctx.sync();
    return true;
  });
}

// ── The user's current selection ──────────────────────────────────────────────

/**
 * Cells previewed with a selection. Enough for a row, a header, a handful of
 * cells — the shapes people actually point at — without putting a whole screen
 * of data into every message.
 */
const MAX_SELECTION_PREVIEW_CELLS = 60;

/**
 * What the user has selected right now, in two forms.
 *
 * `label` goes on the chip above the composer; `context` is what rides along with
 * the message. This exists because "which column is wrong?" is a question people
 * often cannot answer — not out of carelessness, but because the thing that looks
 * wrong to them has no name. They can always point at it, though.
 *
 * Returns null when nothing useful is selected.
 */
export async function getSelectionRef() {
  return Excel.run(async (ctx) => {
    const r = ctx.workbook.getSelectedRange();
    const sh = ctx.workbook.worksheets.getActiveWorksheet();
    r.load('address,rowCount,columnCount,values');
    sh.load('name');
    await ctx.sync();

    // `address` already carries the sheet ("缺料警示!A4:L4"); strip it for the chip
    // so a long sheet name doesn't push the useful half out of a 350px bar.
    const full = String(r.address || '');
    const local = full.includes('!') ? full.slice(full.indexOf('!') + 1) : full;
    if (!local) return null;

    const cells = r.rowCount * r.columnCount;
    const size = `${r.rowCount} 列 × ${r.columnCount} 欄`;
    const head = `使用者在側邊欄提問時選取的範圍：${sh.name}!${local}（${size}）`;

    if (cells > MAX_SELECTION_PREVIEW_CELLS) {
      return { label: `${sh.name}!${local}`, context: `${head}\n（範圍較大，內容請自己用 excel_read_range 讀）` };
    }
    return { label: `${sh.name}!${local}`, context: `${head}\n內容：\n${toCsv(r.values)}` };
  });
}

/**
 * Call `onChange` when the selection moves.
 *
 * Debounced: dragging across a range fires this per cell, and the chip only needs
 * to settle. Failure is non-fatal — the chip then only refreshes when a message is
 * sent, which is still useful.
 */
export async function watchSelection(onChange) {
  let timer = null;
  try {
    await Excel.run(async (ctx) => {
      ctx.workbook.onSelectionChanged.add(async () => {
        if (timer) clearTimeout(timer);
        timer = setTimeout(async () => {
          try { onChange(await getSelectionRef()); } catch { /* ignore */ }
        }, 150);
      });
      await ctx.sync();
    });
  } catch { /* older host: no live tracking */ }
}

export async function getDocumentName() {
  return Excel.run(async (ctx) => {
    ctx.workbook.load('name');
    await ctx.sync();
    return ctx.workbook.name;
  });
}
/**
 * The worksheet the user is currently on.
 *
 * Shown in the pane so it is never a mystery which sheet the agent is talking
 * about — "這一段" and excel_get_selection both resolve against this one, and a
 * workbook with 工作表1 / 工作表5 / 缺料警示 gives no other clue which is in play.
 */
/** The custom document property that carries our id. */
const WORKBOOK_ID_PROP = 'JVisionWorkbookId';

/**
 * A stable identifier for this workbook, for keying conversation history.
 *
 * Office.js has no built-in workbook id, and everything it does expose changes
 * when the file is renamed or moved — which is the one thing this must survive.
 * People rename files constantly; losing the history when they do reads as a bug,
 * not as a filing problem. So we mint a GUID into a custom document property the
 * first time and read it back after.
 *
 * The cost: writing that property marks the workbook as modified. Callers must
 * therefore only ask for the id at a point where the person is already doing
 * something to the file — sending a message, not opening the pane.
 *
 * Returns null when the id cannot be stored (a read-only or protected file, or a
 * host below ExcelApi 1.7). The caller falls back to the file name, which is
 * worse but not broken.
 */
export async function getDocumentKey() {
  if (!hasApi('1.7')) return null;
  try {
    return await Excel.run(async (ctx) => {
      const custom = ctx.workbook.properties.custom;
      const existing = custom.getItemOrNullObject(WORKBOOK_ID_PROP);
      existing.load('isNullObject,value');
      await ctx.sync();
      if (!existing.isNullObject && existing.value) return String(existing.value);

      // crypto.randomUUID needs a secure context; the pane is always https.
      const id = (self.crypto && self.crypto.randomUUID)
        ? self.crypto.randomUUID()
        : String(Date.now()) + '-' + Math.random().toString(16).slice(2);
      custom.add(WORKBOOK_ID_PROP, id);
      await ctx.sync();
      return id;
    });
  } catch {
    // Read-only, protected, or a host that will not take the property.
    return null;
  }
}

export async function getLocation() {
  return Excel.run(async (ctx) => {
    const sh = ctx.workbook.worksheets.getActiveWorksheet();
    sh.load('name');
    await ctx.sync();
    return sh.name;
  });
}

/**
 * Call `onChange` whenever the user switches sheets.
 *
 * Registered once; Office.js keeps the handler alive for the pane's lifetime.
 * Failure is non-fatal — the name simply stops tracking, which is better than a
 * broken pane on a host that lacks the event (it needs ExcelApi 1.7).
 */
export async function watchLocation(onChange) {
  try {
    await Excel.run(async (ctx) => {
      ctx.workbook.worksheets.onActivated.add(async () => {
        try { onChange(await getLocation()); } catch { /* ignore */ }
      });
      await ctx.sync();
    });
  } catch { /* older host: no live tracking */ }
}


/**
 * Scratch space shared across script calls, for the pane's lifetime.
 *
 * This is what makes "scan in one script, write in the next" possible without the
 * data passing through the model at all. Before it existed the model had exactly
 * one channel between two scripts — its own return value — so a task like
 * "find the 99 shortage rows, then write them into a sheet" forced all 99 rows
 * out through the context window and back in again. That measured 19.6 KB and
 * 17 seconds of dead time per hop.
 *
 * Not persisted and not sent anywhere: it lives in this module, in the pane.
 */
const scriptStore = Object.create(null);

/** One short line per key, so the model remembers what it parked and under what name. */
function describeStore() {
  const keys = Object.keys(scriptStore);
  if (!keys.length) return '';
  return `store 目前有：${keys.map((k) => {
    const v = scriptStore[k];
    if (Array.isArray(v)) return `${k}（${v.length} 筆陣列）`;
    if (v && typeof v === 'object') return `${k}（物件，${Object.keys(v).length} 個鍵）`;
    return `${k}（${typeof v}）`;
  }).join('、')}`;
}

// ── run_script ────────────────────────────────────────────────────────────────
/**
 * Result cap.
 *
 * Measured: with a 20 KB cap the model filled it — 19.6 KB of scan output, then
 * 17 seconds of silence while it read the whole thing back in. That is the exact
 * "10,000 cells into the context window" problem this tool was added to avoid,
 * arriving through a different door.
 *
 * The cap is not the fix on its own, though; the model returned everything
 * because it needed those rows later and had nowhere else to put them. `store`
 * below is that somewhere. The cap just makes the wrong path fail loudly, and
 * the truncation message names the right one.
 */
const MAX_SCRIPT_RESULT_CHARS = 4000;
const MAX_SCRIPT_LOG_LINES = 100;
/**
 * Client-side limits, both kept UNDER the server bridge's own timeouts
 * (45s for reads, 120s for confirmed writes) so a slow script reports itself
 * with an actionable message instead of the bridge timing out anonymously.
 */
const SCRIPT_TIMEOUT_READ_MS = 40000;
const SCRIPT_TIMEOUT_WRITE_MS = 100000;

function stringifyResult(v, cap, tailNote) {
  let s;
  if (typeof v === 'string') s = v;
  else {
    try { s = JSON.stringify(v, null, 2); } catch { s = String(v); }
    if (s === undefined) s = String(v);   // JSON.stringify(undefined) is undefined
  }
  return s.length > cap ? s.slice(0, cap) + tailNote : s;
}

/**
 * Run model-written JavaScript against the live workbook.
 *
 * This is the most important tool in the set, and the reason is context, not
 * convenience. Reading a 125-item × 26-period sheet through excel_read_range puts
 * ~10,000 numbers into the model's context and makes the MODEL do the arithmetic,
 * token by token — measured at 342s and still going on a real sheet. The same
 * question answered by a script is one round trip and returns three lines.
 *
 * `new Function` rather than eval: the body sees exactly the four bindings passed
 * to it and not this module's scope. That is scope hygiene, NOT a security
 * sandbox — the code shares the pane's realm and could reach `window` if it tried.
 * The real gate is upstream: the server marks anything that looks like a write as
 * needing confirmation (see scriptMutatesWorkbook), and the pane will not run it
 * until a human clicks.
 */
async function runScript(args, allowWrites) {
  const code = String(args.code || '').trim();
  if (!code) throw new Error('code 是必要的');

  const logs = [];
  const log = (...parts) => {
    if (logs.length >= MAX_SCRIPT_LOG_LINES) return;
    logs.push(parts.map(p => (typeof p === 'string' ? p : stringifyResult(p, 500, '…'))).join(' '));
  };

  let body;
  try {
    body = new Function(
      // `context` as well as `ctx`: every Office.js doc, sample and Script Lab
      // snippet writes Excel.run(async (context) => …), so that is what the model
      // reaches for. Measured — three consecutive calendar attempts died on
      // 「context is not defined」, and the agent then spent the rest of the run
      // deliberating over how to avoid a fourth. Fighting the convention was the
      // mistake; providing both names costs nothing.
      'ctx', 'context', 'workbook', 'Excel', 'log', 'store',
      '"use strict"; return (async () => {\n' + code + '\n})();',
    );
  } catch (e) {
    throw new Error(`程式碼語法錯誤：${(e && e.message) || String(e)}`);
  }

  const limit = String(args.mode || '') === 'write' ? SCRIPT_TIMEOUT_WRITE_MS : SCRIPT_TIMEOUT_READ_MS;

  const guard = makeScriptGuard(allowWrites);
  const run = Excel.run(async (ctx) => {
    const g = guard.wrap(ctx);
    const value = await body(g, g, g.workbook, guard.wrapNamespace(Excel), log, scriptStore);
    // A script that queued mutations but never synced would silently do nothing;
    // syncing again here is free when the script already did it.
    await ctx.sync();
    return value;
  });

  // Note: this races the timeout, it does not cancel the script — Excel.run has no
  // abort. The point is to hand the model a message it can act on rather than let
  // the bridge time out with no explanation.
  const value = await Promise.race([
    run,
    new Promise((_, reject) => setTimeout(
      () => reject(new Error(
        `程式執行超過 ${Math.round(limit / 1000)} 秒。`
        + '常見原因是在迴圈裡呼叫 ctx.sync()——改成一次 load 整塊、sync 一次，剩下用純 JS 算。',
      )),
      limit,
    )),
  ]);

  // A script builds a whole report in one go, and that is exactly where the
  // layout goes wrong: nothing between the model and Office.js converts
  // `columnWidth = 12` from the characters it meant into the points Office reads,
  // so a report arrives as columns of ###. Repair it from the content the same
  // widen-only way a write does — see fitLayout. Only worth doing for a script
  // that actually changed something, and never worth failing a good script over.
  let fitted = null;
  if (allowWrites && guard.didMutate()) {
    try {
      fitted = await Excel.run(async (ctx) => {
        const sheets = ctx.workbook.worksheets;
        sheets.load('items/name');
        await ctx.sync();

        // Which sheets the script touched is not knowable — the guard records that
        // something was written, not where — so every sheet's used range is
        // measured. valuesOnly, or a sheet that merely has formatting way out at
        // BZ400 drags the sweep across four hundred empty columns.
        const used = sheets.items.slice(0, MAX_FIT_SHEETS)
          .map(s => ({ sheet: s, range: s.getUsedRangeOrNullObject(true) }));
        for (const u of used) u.range.load('isNullObject,rowIndex,columnIndex,rowCount,columnCount');
        await ctx.sync();

        return fitLayout(ctx, used.filter(u => !u.range.isNullObject).map(u => ({
          sheet: u.sheet,
          rowIndex: u.range.rowIndex,
          columnIndex: u.range.columnIndex,
          rowCount: u.range.rowCount,
          columnCount: u.range.columnCount,
        })));
      });
    } catch { /* 版面沒調到不影響 script 已經寫好的內容 */ }
  }

  const parts = [];
  if (logs.length) parts.push('log:\n' + logs.join('\n'));
  const fitNote = describeFit(fitted);
  if (fitNote) parts.push(fitNote);
  // Remind the model what it parked and under which name — otherwise the next
  // script guesses the key and gets undefined.
  const stored = describeStore();
  if (stored) parts.push(stored);
  parts.push(
    value === undefined
      ? '（程式沒有 return 任何東西。要拿到結果請用 return。）'
      : stringifyResult(
        value,
        MAX_SCRIPT_RESULT_CHARS,
        '\n…（結果超過 4 KB 已截斷。要留給下一支 script 用的資料請放進 store，'
        + '例如 store.rows = rows —— 它跨呼叫保留，而且不會經過你的上下文。'
        + '若要寫進工作表，就在這一支 script 裡直接寫完，不要先回傳。）',
      ),
  );
  return parts.join('\n\n');
}

/**
 * Dispatch one tool call.
 *
 * Always resolves — never throws. A failure has to reach the model as a readable
 * tool result ("找不到工作表「損益表」") so it can correct itself; an exception
 * here would just strand the agent waiting for a result that never comes.
 */
/**
 * @param opts.writeApproved  whether a person agreed to this call changing the
 *   workbook. Only excel_run_script reads it — every other tool is either a read
 *   or was flagged destructive by name, and the confirmation already happened.
 */
export async function executeTool(tool, args, opts = {}) {
  /**
   * Held out here so the catch below can hand it back too. A script that wrote
   * three sheets and then threw has changed the workbook exactly as much as one
   * that succeeded — dropping the snapshot on the error path would leave the
   * user with the mess and no way back, which is the one case they need it most.
   */
  let scriptSnapshot;
  try {
    switch (tool) {
      case 'excel_get_overview':  return { ok: true, content: await getOverview(args || {}) };
      case 'excel_get_selection': return { ok: true, content: await getSelection() };
      case 'excel_read_range':    return { ok: true, content: await readRange(args || {}) };
      case 'excel_search':        return { ok: true, content: await search(args || {}) };
      case 'excel_write_range': {
        const r = await writeRange(args || {});
        return { ok: true, content: r.text, snapshot: r.snapshot };
      }
      case 'excel_manage_sheet':   return { ok: true, content: await manageSheet(args || {}) };
      case 'excel_format_range':   return { ok: true, content: await formatRange(args || {}) };
      case 'excel_create_chart':   return { ok: true, content: await createChart(args || {}) };
      case 'excel_structure_op':   return { ok: true, content: await structureOp(args || {}) };
      case 'excel_table_ops':      return { ok: true, content: await tableOps(args || {}) };
      case 'excel_clean_data': {
        const r = await cleanData(args || {});
        return { ok: true, content: r.text, snapshot: r.snapshot };
      }
      case 'excel_create_pivot':   return { ok: true, content: await createPivot(args || {}) };
      case 'excel_trace_precedents': return { ok: true, content: await tracePrecedents(args || {}) };
      case 'excel_sheet_style':    return { ok: true, content: await sheetStyle(args || {}) };
      case 'excel_run_script': {
        // Only for approved writes: a read-mode script cannot change anything
        // (the guard sees to that), so photographing the workbook first would be
        // seconds of delay bought for nothing.
        if (opts.writeApproved) scriptSnapshot = await captureWorkbook();
        const text = await runScript(args || {}, !!opts.writeApproved);
        await finishWorkbookSnapshot(scriptSnapshot);
        return { ok: true, content: text, snapshot: scriptSnapshot || undefined };
      }
      default:
        return { ok: false, error: `未知的工具：${tool}` };
    }
  } catch (e) {
    // Office.js errors carry a `code` that is far more actionable than the message
    // (e.g. InvalidArgument on a malformed A1 address).
    const code = e && e.code ? `[${e.code}] ` : '';
    let msg = `${code}${(e && e.message) || String(e)}`;
    // A ReferenceError means the script used a name this sandbox does not define.
    // The bare message says what broke but not what to use instead, and an agent
    // that cannot tell the difference just rewrites the same script again.
    if (tool === 'excel_run_script' && e instanceof ReferenceError) {
      msg += '。這段程式可以用的名稱只有：context（= ctx）、workbook、Excel、log、store。'
        + '沒有 Office Scripts 的 ExcelScript，也沒有 require / import / fetch / document。';
    }
    // A failed script may still have written something before it threw.
    await finishWorkbookSnapshot(scriptSnapshot);
    return { ok: false, error: msg, snapshot: scriptSnapshot || undefined };
  }
}
