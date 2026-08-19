/**
 * excelTools — the Office.js half of the bridge.
 *
 * The server sends `{tool, args}`; this module runs it against the workbook the
 * user actually has open and returns text the model can read. It contains NO AI
 * logic and no network code on purpose: it is the one part that can only be
 * tested inside a real Excel, so everything else was kept out of it.
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
  // Native precedent tracing needs ExcelApi 1.12, but there is a formula-parsing
  // fallback that works on any build, so this one is always offered.
  'excel_trace_precedents',
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
export async function inspectWrite(args) {
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
    target.load('address,formulas,numberFormat');
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

    return {
      text: `已寫入 ${sheetName}!${bareAddress(target.address)}（${rows} 列 × ${cols} 欄，${asFormulas ? '公式' : '值'}）。`,
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

const BORDER_EDGES = {
  all: ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight', 'InsideHorizontal', 'InsideVertical'],
  outline: ['EdgeTop', 'EdgeBottom', 'EdgeLeft', 'EdgeRight'],
  bottom: ['EdgeBottom'],
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
    const applied = [];

    if (args.number_format !== undefined) { r.numberFormat = [[args.number_format]]; applied.push('數字格式'); }
    if (args.bold !== undefined) { r.format.font.bold = !!args.bold; applied.push(args.bold ? '粗體' : '取消粗體'); }
    if (args.italic !== undefined) { r.format.font.italic = !!args.italic; applied.push('斜體'); }
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
    if (args.column_width !== undefined) { r.format.columnWidth = Number(args.column_width); applied.push('欄寬'); }
    if (args.row_height !== undefined) { r.format.rowHeight = Number(args.row_height); applied.push('列高'); }

    if (args.borders !== undefined) {
      const edges = BORDER_EDGES[String(args.borders)] || [];
      if (String(args.borders) === 'none') {
        for (const e of BORDER_EDGES.all) r.format.borders.getItem(e).style = 'None';
      } else {
        for (const e of edges) {
          const b = r.format.borders.getItem(e);
          b.style = 'Continuous';
          b.weight = 'Thin';
        }
      }
      applied.push('框線');
    }

    if (args.merge === true) { r.merge(false); applied.push('合併'); }
    else if (args.merge === false) { r.unmerge(); applied.push('取消合併'); }

    await ctx.sync();

    // Autofit runs after the content-affecting changes land, or it measures the
    // old text.
    //
    // getEntireColumn() matters: Range.format.autofitColumns() measures ONLY the
    // cells inside the given range. Autofitting a header row alone therefore sizes
    // every column to its 2-3 character title, and the numbers underneath collapse
    // into ###. Measuring the whole column is what "自動調整欄寬" means in Excel's
    // own UI, so that is what this does.
    if (args.autofit_columns) {
      r.getEntireColumn().format.autofitColumns();
      await ctx.sync();
      applied.push('自動欄寬');
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
      t.load('name');
      await ctx.sync();
      return `已把 ${args.range} 轉成表格「${t.name}」。`;
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

/** Put a snapshot back. Restoring formulas covers both constants and formulas. */
export async function restoreSnapshot(snapshot) {
  return Excel.run(async (ctx) => {
    const sheet = ctx.workbook.worksheets.getItem(snapshot.sheet);
    const target = sheet.getRange(snapshot.address);
    target.numberFormat = snapshot.numberFormat;
    target.formulas = snapshot.formulas;
    await ctx.sync();
  });
}

/** Scroll to and select a cell — powers the [[Sheet!A1]] citation links. */
export async function selectCell(sheetName, address) {
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

export async function getWorkbookName() {
  return Excel.run(async (ctx) => {
    ctx.workbook.load('name');
    await ctx.sync();
    return ctx.workbook.name;
  });
}

/**
 * Dispatch one tool call.
 *
 * Always resolves — never throws. A failure has to reach the model as a readable
 * tool result ("找不到工作表「損益表」") so it can correct itself; an exception
 * here would just strand the agent waiting for a result that never comes.
 */
export async function executeTool(tool, args) {
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
      case 'excel_create_pivot':   return { ok: true, content: await createPivot(args || {}) };
      case 'excel_trace_precedents': return { ok: true, content: await tracePrecedents(args || {}) };
      default:
        return { ok: false, error: `未知的工具：${tool}` };
    }
  } catch (e) {
    // Office.js errors carry a `code` that is far more actionable than the message
    // (e.g. InvalidArgument on a malformed A1 address).
    const code = e && e.code ? `[${e.code}] ` : '';
    return { ok: false, error: `${code}${(e && e.message) || String(e)}` };
  }
}
