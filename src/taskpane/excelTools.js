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
