/**
 * compare — diff the open workbook against a file the user attached.
 *
 * ── Why this is a module and not a prompt ──
 * PRD 3.2 asks for cross-workbook comparison, and AC-40 grades it on whether the
 * difference values are CORRECT. The obvious implementation — let the model read
 * the uploaded file with excel_read_file, read the sheet with excel_read_range,
 * and write out what it believes the differences are — puts a language model in
 * the middle of a numeric comparison. That is the one job it should never hold.
 * Five hundred rows retyped through a tool call is five hundred chances to drop
 * a digit, and the failure is invisible: the comparison sheet looks exactly as
 * confident when it is wrong.
 *
 * So the model chooses WHAT to compare and this module does the comparing. Both
 * sides arrive as real data — the sheet through Office.js, the source through the
 * pane's own fetch — and what comes out is arithmetic.
 *
 * ── No Office.js in here ──
 * Arrays in, arrays out, which is what makes the judgement calls below testable:
 * what counts as the same value, how rows line up when one side has extra ones,
 * what to do when there is no key. See test/compare.test.mjs. The Office half —
 * reading the range, writing the sheet — stays in excel.js.
 */

/** Report at most this many differing cells. Past it the sheet is not a report, it is a second copy. */
export const MAX_DIFF_ROWS = 5000;

/**
 * Source formats that carry a grid.
 *
 * A PDF or a .docx does not. Text can be extracted from either, but the rows and
 * columns are gone by then — so rather than guess at a table and diff against the
 * guess, those are refused here and the model is pointed at excel_read_file,
 * which is the honest tool for prose.
 */
export const GRID_EXTS = new Set(['xlsx', 'xls', 'csv', 'tsv']);

/**
 * One line of a CSV, honouring quotes.
 *
 * Written out rather than split(',') because a comma inside "台北市, 信義區" is
 * not a column boundary, and getting that wrong shifts every column after it —
 * which then reports as a difference in every single row. A comparison tool that
 * fails this way is worse than none: it produces a large, specific, wrong answer.
 */
function parseDelimited(line, delimiter) {
  const out = [];
  let cell = '';
  let quoted = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (quoted) {
      if (ch === '"') {
        // "" inside a quoted cell is one literal quote, not the end of it.
        if (line[i + 1] === '"') { cell += '"'; i++; } else quoted = false;
      } else cell += ch;
      continue;
    }
    if (ch === '"' && cell === '') { quoted = true; continue; }
    if (ch === delimiter) { out.push(cell); cell = ''; continue; }
    cell += ch;
  }
  out.push(cell);
  return out;
}

const cellText = (v) => (v === null || v === undefined ? '' : String(v));

/**
 * The uploaded file's text → the sheets inside it.
 *
 * The shapes handled here are the ones the server's extractor actually produces
 * (see services/fileTextExtract.ts): a spreadsheet comes back as `# SheetName`
 * followed by tab-separated rows, and a csv/tsv comes back as itself. Anything
 * else returns no sheets, and the caller says so rather than diffing prose.
 */
export function parseSourceGrid(text, ext) {
  const e = String(ext || '').toLowerCase();
  if (!GRID_EXTS.has(e)) {
    return { sheets: [], reason: `「${e || '這種格式'}」沒有列與欄的結構，沒辦法逐格比對。` };
  }

  const raw = String(text == null ? '' : text).replace(/\r\n?/g, '\n');
  if (!raw.trim()) return { sheets: [], reason: '這個檔案是空的。' };

  if (e === 'csv' || e === 'tsv') {
    const delimiter = e === 'tsv' ? '\t' : ',';
    // A trailing newline is how text files end, not an empty final row.
    const lines = raw.split('\n');
    if (lines.length && lines[lines.length - 1] === '') lines.pop();
    return { sheets: [{ name: '', rows: lines.map((l) => parseDelimited(l, delimiter)) }], reason: '' };
  }

  // Spreadsheet: `# SheetName` opens each sheet, tabs separate the cells.
  const sheets = [];
  let current = null;
  for (const line of raw.split('\n')) {
    if (line.startsWith('# ')) {
      current = { name: line.slice(2).trim(), rows: [] };
      sheets.push(current);
      continue;
    }
    if (line === '') continue;
    // A row before any header means the extractor changed shape under us. Take
    // it rather than drop it — losing rows silently is the worst failure here.
    if (!current) { current = { name: '', rows: [] }; sheets.push(current); }
    current.rows.push(line.split('\t'));
  }
  const kept = sheets.filter((s) => s.rows.length);
  return { sheets: kept, reason: kept.length ? '' : '這個檔案裡沒有任何資料列。' };
}

/** Pick the sheet the caller named, or the biggest one when they named none. */
export function pickSheet(sheets, wanted) {
  if (!sheets || !sheets.length) return null;
  const want = String(wanted || '').trim();
  if (want) {
    return sheets.find((s) => s.name === want)
      || sheets.find((s) => s.name.toLowerCase() === want.toLowerCase())
      || null;
  }
  // Tallest wins, because the sheet with the data is very rarely the cover sheet
  // that happens to come first.
  return [...sheets].sort((a, b) => b.rows.length - a.rows.length)[0];
}

/**
 * A cell as a number, or null when it is not one.
 *
 * Thousands separators, a trailing %, and accounting parentheses are stripped
 * because all three are how a number LOOKS, not what it is. An empty cell is
 * deliberately not zero: a blank where a figure should be is a difference worth
 * reporting, and folding it to 0 would hide exactly the rows somebody is
 * comparing to find.
 */
export function toNumber(v) {
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (v === null || v === undefined) return null;
  const s = String(v).trim();
  if (!s) return null;
  const cleaned = s.replace(/,/g, '').replace(/^\((.*)\)$/, '-$1');
  const pct = cleaned.endsWith('%');
  const body = pct ? cleaned.slice(0, -1) : cleaned;
  if (!/^[+-]?(\d+\.?\d*|\.\d+)$/.test(body)) return null;
  const n = Number(body);
  if (!Number.isFinite(n)) return null;
  return pct ? n / 100 : n;
}

/**
 * Are these two cells the same value?
 *
 * Numbers are compared AS numbers, and that is the important half. The workbook
 * hands back 1234.5 and the source file hands back the string "1,234.50" — the
 * same number, written by two different formatters. Comparing those as text
 * reports every formatted cell in the file as a difference, which buries the
 * handful of real ones. `tolerance` covers the other direction: a value that
 * survived a round trip through a text format and came back 0.30000000000000004.
 */
export function sameValue(a, b, opts = {}) {
  const { trim = true, ignoreCase = false, tolerance = 0 } = opts;
  const na = toNumber(a);
  const nb = toNumber(b);
  if (na !== null && nb !== null) {
    return tolerance > 0 ? Math.abs(na - nb) <= tolerance : na === nb;
  }
  let sa = cellText(a);
  let sb = cellText(b);
  if (trim) { sa = sa.trim(); sb = sb.trim(); }
  if (ignoreCase) { sa = sa.toLowerCase(); sb = sb.toLowerCase(); }
  return sa === sb;
}

/**
 * Line the two sides' columns up.
 *
 * By header name when both sides have one, because a column inserted in the
 * middle of the source file is the most common way two versions of the same
 * report differ — and matching by position after that reports every column to
 * its right as changed. By position when there are no headers, which is the best
 * available answer rather than a good one.
 *
 * `fallbackWidth` is not optional in practice, it just looks it. With
 * `headers: false` there ARE no header rows to measure, so without a width
 * taken from the data this returned zero columns and the diff compared nothing
 * at all — reporting two completely different grids as identical, which is the
 * one answer this module must never give.
 */
export function alignColumns(curHeader, srcHeader, fallbackWidth = 0) {
  if (curHeader && srcHeader) {
    const srcIndex = new Map();
    srcHeader.forEach((h, i) => {
      const k = cellText(h).trim().toLowerCase();
      if (k && !srcIndex.has(k)) srcIndex.set(k, i);
    });
    const pairs = [];
    const usedSrc = new Set();
    curHeader.forEach((h, i) => {
      const k = cellText(h).trim().toLowerCase();
      const j = k ? srcIndex.get(k) : undefined;
      if (j !== undefined) usedSrc.add(j);
      pairs.push({ label: cellText(h) || `第 ${i + 1} 欄`, cur: i, src: j === undefined ? null : j });
    });
    // Columns the source has and the workbook does not are still news — they are
    // the 差異欄 half of what the PRD asks for.
    srcHeader.forEach((h, j) => {
      if (usedSrc.has(j)) return;
      pairs.push({ label: cellText(h) || `來源第 ${j + 1} 欄`, cur: null, src: j });
    });
    return pairs;
  }
  const width = Math.max(
    curHeader ? curHeader.length : 0,
    srcHeader ? srcHeader.length : 0,
    Number(fallbackWidth) || 0,
  );
  return Array.from({ length: width }, (_, i) => ({ label: `第 ${i + 1} 欄`, cur: i, src: i }));
}

/** Match rows by the value in the key column; unmatched rows on either side stand alone. */
function pairByKey(curBody, srcBody, keyCur, keySrc, opts) {
  const norm = (v) => {
    let s = cellText(v);
    if (opts.trim) s = s.trim();
    if (opts.ignoreCase) s = s.toLowerCase();
    // '001' and 1 are the same customer id on two sides of an export.
    const n = toNumber(s);
    return n === null ? s : String(n);
  };

  const srcBuckets = new Map();
  srcBody.forEach((row, i) => {
    const k = norm(row[keySrc]);
    if (!srcBuckets.has(k)) srcBuckets.set(k, []);
    srcBuckets.get(k).push(i);
  });

  const pairs = [];
  const takenSrc = new Set();
  curBody.forEach((row, i) => {
    const k = norm(row[keyCur]);
    const bucket = srcBuckets.get(k);
    // Duplicate keys pair in occurrence order — the second 「A001」 on one side
    // goes with the second on the other, which is the least surprising rule
    // available once the key has stopped being unique.
    const j = bucket && bucket.length ? bucket.shift() : null;
    if (j !== null) takenSrc.add(j);
    pairs.push({ c: i, s: j, key: cellText(row[keyCur]) });
  });
  srcBody.forEach((row, j) => {
    if (takenSrc.has(j)) return;
    pairs.push({ c: null, s: j, key: cellText(row[keySrc]) });
  });
  return pairs;
}

/** No key: row 1 against row 1. Honest, and said as such in the report header. */
function pairByPosition(curBody, srcBody) {
  const n = Math.max(curBody.length, srcBody.length);
  const pairs = [];
  for (let i = 0; i < n; i++) {
    pairs.push({
      c: i < curBody.length ? i : null,
      s: i < srcBody.length ? i : null,
      key: `第 ${i + 1} 列`,
    });
  }
  return pairs;
}

/** A whole row flattened, for the "this row exists on one side only" report line. */
function rowText(row, columns, side) {
  if (!row) return '';
  const parts = [];
  for (const col of columns) {
    const i = side === 'cur' ? col.cur : col.src;
    if (i === null) continue;
    const v = cellText(row[i]);
    if (v !== '') parts.push(`${col.label}=${v}`);
    if (parts.length >= 8) { parts.push('…'); break; }
  }
  return parts.join('；');
}

/**
 * The comparison.
 *
 * `key_column` is what makes this useful rather than merely correct. Without it,
 * a single row inserted at the top of one side shifts everything below it and
 * every row reports as changed — technically true, and useless. With a key, rows
 * are matched by identity, so an inserted row reports as one addition and
 * nothing else moves.
 */
export function diffGrids(current, source, opts = {}) {
  const {
    headers = true,
    keyColumn = null,
    trim = true,
    ignoreCase = false,
    tolerance = 0,
    limit = MAX_DIFF_ROWS,
  } = opts;

  const cur = (current || []).map((r) => (Array.isArray(r) ? r : []));
  const src = (source || []).map((r) => (Array.isArray(r) ? r : []));

  const curHeader = headers && cur.length ? cur[0] : null;
  const srcHeader = headers && src.length ? src[0] : null;
  const curBody = headers ? cur.slice(1) : cur;
  const srcBody = headers ? src.slice(1) : src;
  /** Row 1 of the range is the header when there is one, so data starts at 2. */
  const rowOffset = headers ? 2 : 1;

  // Measured with a loop rather than Math.max(...rows.map(…)): a spread of
  // 100 000 arguments is a stack overflow, and a workbook that size is exactly
  // the one somebody reaches for this tool with.
  let bodyWidth = 0;
  for (const r of curBody) if (r.length > bodyWidth) bodyWidth = r.length;
  for (const r of srcBody) if (r.length > bodyWidth) bodyWidth = r.length;

  const columns = alignColumns(curHeader, srcHeader, bodyWidth);
  const notes = [];

  const wantKey = keyColumn !== null && keyColumn !== undefined && Number.isFinite(Number(keyColumn));
  const keyCur = wantKey ? Number(keyColumn) : null;
  // The key is named by its position in the WORKBOOK; find where that column
  // landed on the source side rather than assuming it is the same index.
  const keyPair = wantKey ? columns.find((c) => c.cur === keyCur) : null;
  const keySrc = keyPair ? keyPair.src : keyCur;
  const keyed = wantKey && keySrc !== null && keySrc !== undefined;
  if (wantKey && !keyed) {
    notes.push(`來源檔裡找不到對應的鍵欄位「${keyPair ? keyPair.label : keyCur}」，已改用列的先後順序比對。`);
  }

  const pairs = keyed
    ? pairByKey(curBody, srcBody, keyCur, keySrc, { trim, ignoreCase })
    : pairByPosition(curBody, srcBody);

  const diffs = [];
  let changedRows = 0;
  let onlyCurrent = 0;
  let onlySource = 0;
  let sameRows = 0;
  let truncated = false;

  for (const pair of pairs) {
    if (diffs.length >= limit) { truncated = true; break; }
    const { c, s, key } = pair;

    if (c !== null && s === null) {
      onlyCurrent++;
      diffs.push({
        status: 'onlyCurrent',
        key,
        column: '（整列）',
        cur: rowText(curBody[c], columns, 'cur'),
        src: '',
        row: c + rowOffset,
        col: null,
      });
      continue;
    }
    if (c === null && s !== null) {
      onlySource++;
      diffs.push({
        status: 'onlySource',
        key,
        column: '（整列）',
        cur: '',
        src: rowText(srcBody[s], columns, 'src'),
        row: null,
        col: null,
      });
      continue;
    }

    let rowChanged = false;
    for (const col of columns) {
      if (diffs.length >= limit) { truncated = true; break; }
      const a = col.cur === null ? '' : curBody[c][col.cur];
      const b = col.src === null ? '' : srcBody[s][col.src];
      if (sameValue(a, b, { trim, ignoreCase, tolerance })) continue;
      rowChanged = true;
      diffs.push({
        status: 'changed',
        key,
        column: col.label,
        cur: cellText(a),
        src: cellText(b),
        row: c + rowOffset,
        col: col.cur,
      });
    }
    if (rowChanged) changedRows++;
    else sameRows++;
  }

  if (truncated) notes.push(`差異太多，只列出前 ${limit} 筆。建議縮小比對範圍再跑一次。`);

  return {
    keyed,
    columns,
    diffs,
    notes,
    counts: {
      changedRows,
      onlyCurrent,
      onlySource,
      sameRows,
      cells: diffs.filter((d) => d.status === 'changed').length,
      curRows: curBody.length,
      srcRows: srcBody.length,
    },
  };
}

const STATUS_LABEL = {
  changed: '值不同',
  onlyCurrent: '只在目前活頁簿',
  onlySource: '只在來源檔',
};

/**
 * The diff table's header row.
 *
 * Exported because excel.js has to find this row again to bold it, and it cannot
 * count its way there: the summary block above grows by one row for every note
 * the comparison produced. Matching on the text means the two files cannot drift
 * apart the way an index would.
 */
export const REPORT_HEAD = ['狀態', '鍵值', '欄位', '目前活頁簿', '來源檔', '位置'];

/** 0-based column index → 'A', 'Z', 'AA'. Same rule as colName in excel.js. */
function colLetters(index) {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

/** Where in the workbook a difference sits. Column-precise when the column is known. */
function addressOf(d, firstRow, firstCol) {
  if (d.row === null) return '';
  const row = d.row + firstRow - 1;
  if (d.col === null || d.col === undefined) return `第 ${row} 列`;
  return `${colLetters(d.col + firstCol)}${row}`;
}

/**
 * The comparison sheet, as a grid ready to write.
 *
 * The summary block goes ON the sheet rather than only into the reply, because
 * the sheet outlives the conversation. Somebody opening this file next week has
 * to be able to tell what was compared against what, and on which key — a bare
 * list of differences with no provenance is a list nobody can act on.
 *
 * `位置` holds the workbook address rather than a second copy of the value, so
 * every reported difference can be checked against the cell it came from.
 * `firstRow` / `firstCol` shift the diff's own coordinates back onto the real
 * sheet: the diff counts from the top-left of the compared RANGE, and a range
 * starting at C5 would otherwise report every address five rows and two columns
 * out — a report that points at the wrong cells is worse than one with none.
 */
export function buildReport(diff, info) {
  const { sourceName, sheetName, keyLabel, rangeText, firstRow = 1, firstCol = 0 } = info;
  const c = diff.counts;
  const grid = [
    ['比對報告'],
    ['來源檔', sourceName],
    ['比對範圍', `${sheetName}!${rangeText}`],
    ['比對方式', diff.keyed ? `以「${keyLabel}」對應列` : '依列的先後順序對應（沒有指定鍵欄位）'],
    ['目前活頁簿列數', c.curRows],
    ['來源檔列數', c.srcRows],
    ['值不同的儲存格', c.cells],
    ['有差異的列', c.changedRows],
    ['只在目前活頁簿的列', c.onlyCurrent],
    ['只在來源檔的列', c.onlySource],
    ['完全相同的列', c.sameRows],
  ];
  for (const note of diff.notes) grid.push(['注意', note]);
  grid.push(['']);
  grid.push([...REPORT_HEAD]);

  if (!diff.diffs.length) {
    grid.push(['沒有差異', '', '', '', '', '']);
    return grid;
  }
  for (const d of diff.diffs) {
    grid.push([
      STATUS_LABEL[d.status] || d.status,
      d.key,
      d.column,
      d.cur,
      d.src,
      d.row === null ? '' : `${sheetName}!${addressOf(d, firstRow, firstCol)}`,
    ]);
  }
  return grid;
}
