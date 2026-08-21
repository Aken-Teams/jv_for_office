/**
 * Drive fitLayout against a stand-in for Office.js.
 *
 * The rule worth pinning is the one that makes this safe to run after EVERY
 * write: widen-only. Autofit on its own is destructive — it happily narrows a
 * column somebody widened on purpose and flattens a banner row — so fitLayout
 * measures, then puts back anything autofit shrank. That is invisible in the
 * code path (it looks like an ordinary clamp) and impossible to check by hand in
 * Excel without a workbook laid out exactly right, which is what this is for.
 *
 * The other half is the bug that started it: a script writing
 * `columnWidth = 12` meaning characters, Office.js reading points, and the whole
 * report coming out as ###. Column A below is that column.
 */
import fs from 'fs';

const src = fs.readFileSync(new URL('../src/taskpane/hosts/excel.js', import.meta.url), 'utf8');
const from = src.indexOf('/* ── fit:start');
const to = src.indexOf('/* ── fit:end');
if (from < 0 || to < 0) {
  console.error('找不到 fit:start / fit:end 標記，hosts/excel.js 的結構變了。');
  process.exit(1);
}
const tmp = new URL('./.fit-under-test.mjs', import.meta.url);
fs.writeFileSync(tmp, src.slice(from, to)
  + '\nexport { fitLayout, describeFit, MIN_COL_POINTS, MAX_COL_POINTS, MAX_FIT_ROWS };\n');
const { fitLayout, describeFit, MIN_COL_POINTS, MAX_COL_POINTS, MAX_FIT_ROWS } = await import(tmp);

// ── A very small Excel ────────────────────────────────────────────────────────
// Columns and rows have a CURRENT size and a NATURAL size — what autofit would
// give them for the content they hold. Everything under test is a decision about
// those two numbers.

function makeSheet(spec) {
  const ops = [];
  const width = [...(spec.width || [])];
  const height = [...(spec.height || [])];
  const natWidth = spec.natWidth || [];
  const natHeight = spec.natHeight || [];
  const hiddenCols = spec.hiddenCols || [];
  const hiddenRows = spec.hiddenRows || [];

  const column = (i) => ({
    load() { return this; },
    columnHidden: hiddenCols.includes(i),
    format: {
      get columnWidth() { return width[i]; },
      set columnWidth(v) { ops.push(`width ${i} := ${v}`); width[i] = v; },
      autofitColumns() { ops.push(`autofit col ${i}`); width[i] = natWidth[i]; },
    },
  });
  const row = (i) => ({
    load() { return this; },
    rowHidden: hiddenRows.includes(i),
    format: {
      get rowHeight() { return height[i]; },
      set rowHeight(v) { ops.push(`height ${i} := ${v}`); height[i] = v; },
      autofitRows() { ops.push(`autofit row ${i}`); height[i] = natHeight[i]; },
    },
  });

  return {
    ops,
    width,
    height,
    sheet: {
      getRangeByIndexes(r, c) {
        return { getEntireColumn: () => column(c), getEntireRow: () => row(r) };
      },
    },
  };
}

const ctx = { sync: async () => {} };

// ── Checks ────────────────────────────────────────────────────────────────────

let pass = 0, fail = 0;
const check = (name, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}`);
  if (!ok) { console.log(`      得到 ${JSON.stringify(got)}，預期 ${JSON.stringify(want)}`); fail++; }
  else pass++;
};

// Five columns, one of each situation the sweep runs into.
//   0  the points-vs-characters bug: 12pt wide, content needs 88
//   1  a column somebody deliberately widened; autofit would take it to 40
//   2  one very long note: autofit says 900
//   3  a short label column already wider than its content
//   4  hidden
const COLS = {
  width: [16, 300, 64, 64, 0],
  natWidth: [88, 40, 900, 10, 88],
  hiddenCols: [4],
};

{
  const s = makeSheet(COLS);
  const got = await fitLayout(ctx, [{ sheet: s.sheet, rowIndex: 0, columnIndex: 0, rowCount: 0, columnCount: 5 }]);
  check('太窄的欄被加寬到內容需要的寬度', s.width[0], 88);
  check('刻意加寬的欄不會被縮回去', s.width[1], 300);
  check('超長內容夾在上限', s.width[2], MAX_COL_POINTS);
  check('已經夠寬的欄維持原狀', s.width[3], 64);
  check('隱藏的欄完全沒被碰', s.width[4], 0);
  check('沒對隱藏欄呼叫 autofit', s.ops.includes('autofit col 4'), false);
  check('只回報真的變動的欄數', got, { columns: 2, rows: 0 });
}

{
  // Explicitly requested autofit (excel_format_range) is the real thing: the same
  // columns, but narrowing is now allowed and the floor still applies.
  const s = makeSheet(COLS);
  await fitLayout(ctx, [{ sheet: s.sheet, rowIndex: 0, columnIndex: 0, rowCount: 0, columnCount: 5 }], { shrink: true });
  check('shrink 時刻意加寬的欄會縮成內容寬度', s.width[1], 40);
  check('shrink 時仍不會窄到看不懂', s.width[3], MIN_COL_POINTS);
}

{
  // Rows: squashed ones grow, a banner row keeps its height.
  const s = makeSheet({
    width: [64], natWidth: [64],
    height: [12, 60, 0], natHeight: [30, 15, 30],
    hiddenRows: [2],
  });
  const got = await fitLayout(ctx, [{ sheet: s.sheet, rowIndex: 0, columnIndex: 0, rowCount: 3, columnCount: 1 }]);
  check('被壓扁的列長回內容需要的高度', s.height[0], 30);
  check('刻意加高的標題列沒被壓扁', s.height[1], 60);
  check('隱藏的列完全沒被碰', s.height[2], 0);
  check('欄寬先於列高處理', s.ops.indexOf('autofit col 0') < s.ops.indexOf('autofit row 0'), true);
  check('欄與列分開計數', got, { columns: 0, rows: 1 });
}

{
  // A long block is data, not a hand-built layout: measuring 200-odd row heights
  // to change nothing is the one case where this pass is pure cost.
  const s = makeSheet({ width: [16], natWidth: [88], height: [], natHeight: [] });
  const got = await fitLayout(ctx, [{
    sheet: s.sheet, rowIndex: 0, columnIndex: 0, rowCount: MAX_FIT_ROWS + 1, columnCount: 1,
  }]);
  check('超過上限的列數就整批跳過列高', s.ops.some(o => o.startsWith('autofit row')), false);
  check('欄寬照樣處理', got, { columns: 1, rows: 0 });
}

check('沒有變動就不多嘴', describeFit({ columns: 0, rows: 0 }), '');
check('有變動就講清楚改了什麼', describeFit({ columns: 3, rows: 1 }), '已依內容自動調整3 欄的欄寬、1 列的列高。');

fs.rmSync(tmp);
console.log(`\n${pass} passed, ${fail} failed`);
if (fail) process.exit(1);
