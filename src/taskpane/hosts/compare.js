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

/* ── text→value ── moved here from excel.js; see the note at its old site ──── */

/** Full-width ！ through ～ sit exactly 0xFEE0 above their ASCII twins. */
export const toHalfWidth = (s) => s
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
export function textToNumber(raw) {
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
export function textToDateSerial(raw) {
  if (raw instanceof Date) {
    return Number.isFinite(raw.getTime())
      ? Math.round((Date.UTC(raw.getUTCFullYear(), raw.getUTCMonth(), raw.getUTCDate())
        - EXCEL_EPOCH_UTC) / 86400000)
      : null;
  }
  const s = toHalfWidth(String(raw)).trim()
    .replace(/[年月]/g, '/').replace(/日$/, '')
    // An ISO timestamp is a date with a time bolted on; the day is the part two
    // spreadsheets can be expected to agree about.
    .replace(/[T ]\d{1,2}:\d{2}(:\d{2})?(\.\d+)?(Z|[+-]\d{2}:?\d{2})?$/, '');
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

/* ── text→value:end ───────────────────────────────────────────────────────── */

/**
 * A cell as a number, or null when it is not one.
 *
 * Looser than textToNumber on purpose, and the difference is deliberate: this
 * one is for COMPARING, that one is for REWRITING a cell. '0012' must stay text
 * when clean_data is about to overwrite the cell, because turning an id into 12
 * destroys it — but when the question is only 「are these two cells the same
 * value」, '0012' and 12 plainly are, and reporting them as a difference would
 * fill the report with noise on every id column that survived a CSV round trip.
 */
export function toNumber(v) {
  if (typeof v === 'number') return Number.isFinite(v) ? v : null;
  if (v === null || v === undefined) return null;
  const s = toHalfWidth(String(v)).trim();
  if (!s) return null;

  // Delegated rather than reimplemented. The first version of this had its own
  // regex and quietly disagreed with textToNumber about NT$2000 — it stripped
  // commas and parentheses but not the currency mark, so a workbook holding 2000
  // and an export holding NT$2000 came back as a difference. One parser for
  // "what does this text mean as a number"; this function only widens it.
  const direct = textToNumber(s);
  if (direct !== null) return direct;

  // Leading zeros are the one place the two questions genuinely differ.
  // textToNumber refuses '0012' because clean_data is about to REWRITE the cell
  // and turning an id into 12 destroys it. Here the question is only whether two
  // cells hold the same value, and '0012' and 12 plainly do — reporting them as
  // a difference fills the report with noise on every id column that has been
  // through a CSV.
  const unpadded = s.replace(/^(-?)0+(?=\d)/, '$1');
  return unpadded === s ? null : textToNumber(unpadded);
}

/**
 * The window in which a bare number might really be a date.
 *
 * Excel hands an add-in a date as a serial — 45292, not 「2024/1/1」 — while a
 * file that has been through a text extractor hands over the words. Without this
 * they never match, and EVERY row of any report with a date column comes back as
 * a difference. That is not a rare edge: it is most of the files anybody would
 * reach for this tool with.
 *
 * The trade-off, stated plainly: a genuine quantity that happens to land between
 * these bounds will be read as a date when the other side holds a date string,
 * and a real difference is then hidden. The window is 1980-01-01 to 2079-12-31,
 * and the coercion only fires when the OTHER side parses as a date — so it takes
 * a column that is a number on one side and a date on the other, which is
 * already broken data. Narrower than a bug in every dated report.
 */
const SERIAL_MIN = 29221;
const SERIAL_MAX = 65744;
const looksLikeSerial = (n) => Number.isInteger(n) && n >= SERIAL_MIN && n <= SERIAL_MAX;

/**
 * This cell as a date serial, or null.
 *
 * `asNumber` is passed in rather than recomputed so the caller's single parse is
 * reused — and so this can tell 「a number that might be a serial」 from 「text
 * that spells a date」 without parsing twice.
 */
function asDateSerial(v, asNumber) {
  if (asNumber !== null) return looksLikeSerial(asNumber) ? asNumber : null;
  return textToDateSerial(v);
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

  // Dates, second. One side arrives from Excel as a serial and the other from a
  // text extractor as words, so this is the only place they can be made to meet.
  const da = asDateSerial(a, na);
  const db = asDateSerial(b, nb);
  if (da !== null && db !== null) return da === db;

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

/** Joins the parts of a composite key.   cannot occur in a cell. */
const KEY_SEP = ' ';

/** What the report shows for a composite key: 「PN-1001 / W02 / L2402」. */
const keyText = (row, cols) => cols.map((i) => cellText(row[i])).join(' / ');

/**
 * Match rows by their key; unmatched rows on either side stand alone.
 *
 * ── Why the key is a LIST ──
 * One column is not enough for the data people actually compare. An inventory
 * snapshot is keyed on 料號＋倉別＋批號 and the same 料號 appears in every
 * warehouse; a rolling forecast is keyed on 料號＋廠區＋客戶. Keyed on 料號
 * alone, every one of those rows is a duplicate, pairing falls back to
 * occurrence order, and two exports that sorted differently — which is most
 * pairs of exports — get matched row-for-row at random.
 *
 * Measured on a six-row inventory fixture that had exactly ONE real change:
 * keyed on 料號 alone it produced fifteen differences, all but one spurious, and
 * mis-paired the real one so its numbers were wrong too. That is the failure
 * this signature exists to prevent, and it is worse than no tool at all.
 */
function pairByKey(curBody, srcBody, keyCur, keySrc, opts) {
  const norm = (row, cols) => cols.map((i) => {
    let s = cellText(row[i]);
    if (opts.trim) s = s.trim();
    if (opts.ignoreCase) s = s.toLowerCase();
    // '001' and 1 are the same customer id on two sides of an export.
    const n = toNumber(s);
    return n === null ? s : String(n);
  }).join(KEY_SEP);

  const srcBuckets = new Map();
  srcBody.forEach((row, i) => {
    const k = norm(row, keySrc);
    if (!srcBuckets.has(k)) srcBuckets.set(k, []);
    srcBuckets.get(k).push(i);
  });

  const pairs = [];
  const takenSrc = new Set();
  curBody.forEach((row, i) => {
    const bucket = srcBuckets.get(norm(row, keyCur));
    // Still duplicates after the composite key? Then pair in occurrence order —
    // the second 「A001」 on one side goes with the second on the other, which is
    // the least surprising rule available once the key has stopped identifying.
    const j = bucket && bucket.length ? bucket.shift() : null;
    if (j !== null) takenSrc.add(j);
    pairs.push({ c: i, s: j, key: keyText(row, keyCur) });
  });
  srcBody.forEach((row, j) => {
    if (takenSrc.has(j)) return;
    pairs.push({ c: null, s: j, key: keyText(row, keySrc) });
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

  // A single number stays valid — most comparisons really do have one key — but
  // the list is the shape the hard cases need. See pairByKey.
  const wanted = (Array.isArray(keyColumn) ? keyColumn : [keyColumn])
    .filter((k) => k !== null && k !== undefined && Number.isFinite(Number(k)))
    .map(Number);

  // Each key column is named by its position in the WORKBOOK; find where each
  // landed on the source side rather than assuming the indexes line up.
  const keyCur = [];
  const keySrc = [];
  const missing = [];
  for (const k of wanted) {
    const pair = columns.find((c) => c.cur === k);
    if (pair && pair.src !== null && pair.src !== undefined) {
      keyCur.push(k);
      keySrc.push(pair.src);
    } else {
      missing.push(pair ? pair.label : `第 ${k + 1} 欄`);
    }
  }
  const keyed = keyCur.length > 0;
  if (missing.length) {
    notes.push(keyed
      // Partial is worth saying out loud: the comparison still ran, but on a
      // weaker key than was asked for, and that changes how much to trust it.
      ? `來源檔裡找不到鍵欄位「${missing.join('、')}」，這次只用 ${keyCur.length} 個欄位對應。`
      : `來源檔裡找不到鍵欄位「${missing.join('、')}」，已改用列的先後順序比對。`);
  }

  const pairs = keyed
    ? pairByKey(curBody, srcBody, keyCur, keySrc, { trim, ignoreCase })
    : pairByPosition(curBody, srcBody);

  /**
   * Columns that exist on one side only, reported ONCE instead of per row.
   *
   * PRD 3.2 asks for 差異列, 差異欄 and 差異值 as three separate things, and this
   * is the middle one. Comparing a one-sided column cell by cell means every row
   * carries a difference against an empty string — on a rolling forecast, where
   * the oldest month drops off and a new one appears at the end, that is two
   * spurious differences per row on top of the real ones. Said once at the top,
   * it is the single fact it actually is: this column is new, that one is gone.
   */
  const common = columns.filter((c) => c.cur !== null && c.src !== null);
  const columnsOnlyCurrent = columns.filter((c) => c.src === null).map((c) => c.label);
  const columnsOnlySource = columns.filter((c) => c.cur === null).map((c) => c.label);

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
    for (const col of common) {
      if (diffs.length >= limit) { truncated = true; break; }
      const a = curBody[c][col.cur];
      const b = srcBody[s][col.src];
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
    columnsOnlyCurrent,
    columnsOnlySource,
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

/**
 * The two sides, named by what they actually are.
 *
 * Fixed labels do not survive the second caller. 「目前活頁簿」 vs 「來源檔」 reads
 * correctly when the comparison is against an uploaded file and becomes nonsense
 * the moment it is one sheet against another in the SAME workbook — both sides
 * are then the current workbook, and a reader has no way to tell which column is
 * which. Naming the sides after the sheet and the file makes the report legible
 * without knowing which kind of comparison produced it.
 */
const statusLabels = (left, right) => ({
  changed: '值不同',
  onlyCurrent: `只在${left}`,
  onlySource: `只在${right}`,
});

/**
 * First cell of the diff table's header row.
 *
 * Exported because excel.js has to find that row again to bold it, and it cannot
 * count its way there: the summary block above grows by one row for every note
 * the comparison produced. The other five headers vary with the comparison, but
 * this one never does — so it is the anchor.
 */
export const REPORT_HEAD_CELL = '狀態';

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
  const {
    leftLabel, rightLabel, sheetName, keyLabel, rangeText,
    rightRangeText = '', firstRow = 1, firstCol = 0,
  } = info;
  const status = statusLabels(leftLabel, rightLabel);
  const c = diff.counts;
  const grid = [
    ['比對報告'],
    ['比對對象', `${leftLabel}　←→　${rightLabel}`],
    ['本表範圍', `${sheetName}!${rangeText}`],
  ];
  // Only meaningful when the other side is a range too — an uploaded file has no
  // address to state, and an empty row would just look like something failed.
  if (rightRangeText) grid.push(['對照範圍', rightRangeText]);
  grid.push(
    ['比對方式', diff.keyed ? `以「${keyLabel}」對應列` : '依列的先後順序對應（沒有指定鍵欄位）'],
  );
  // 差異欄, stated once. A new or dropped column is one fact about the two
  // files, not a difference in every row of them.
  if (diff.columnsOnlyCurrent && diff.columnsOnlyCurrent.length) {
    grid.push([`只在 ${leftLabel} 的欄位`, diff.columnsOnlyCurrent.join('、')]);
  }
  if (diff.columnsOnlySource && diff.columnsOnlySource.length) {
    grid.push([`只在 ${rightLabel} 的欄位`, diff.columnsOnlySource.join('、')]);
  }
  grid.push(
    [`${leftLabel} 列數`, c.curRows],
    [`${rightLabel} 列數`, c.srcRows],
    ['值不同的儲存格', c.cells],
    ['有差異的列', c.changedRows],
    [`只在 ${leftLabel} 的列`, c.onlyCurrent],
    [`只在 ${rightLabel} 的列`, c.onlySource],
    ['完全相同的列', c.sameRows],
  );
  for (const note of diff.notes) grid.push(['注意', note]);
  grid.push(['']);
  grid.push([REPORT_HEAD_CELL, '鍵值', '欄位', leftLabel, rightLabel, '位置']);

  if (!diff.diffs.length) {
    grid.push(['沒有差異', '', '', '', '', '']);
    return grid;
  }
  for (const d of diff.diffs) {
    grid.push([
      status[d.status] || d.status,
      d.key,
      d.column,
      d.cur,
      d.src,
      d.row === null ? '' : `${sheetName}!${addressOf(d, firstRow, firstCol)}`,
    ]);
  }
  return grid;
}
