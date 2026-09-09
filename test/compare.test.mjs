/**
 * The comparison engine, which AC-40 grades on being RIGHT.
 *
 * ── Why this file is long ──
 * Everything here is a case where the obvious implementation gives a confident
 * wrong answer rather than an error. A comma inside a quoted CSV cell shifts
 * every column after it; comparing 「1,234.50」 to 1234.5 as text reports every
 * formatted cell in the file as a difference; matching rows by position after
 * somebody inserted one at the top reports the entire sheet as changed. None of
 * those throw. All of them produce a large, specific, plausible comparison sheet
 * that is wrong, which is the worst outcome this feature has — worse than
 * refusing to run, because somebody will act on it.
 *
 * So the cases below are the ones where being wrong is silent.
 */
import {
  parseSourceGrid, pickSheet, sameValue, toNumber, alignColumns,
  diffGrids, buildReport, MAX_DIFF_ROWS, REPORT_HEAD_CELL,
} from '../src/taskpane/hosts/compare.js';

let pass = 0, fail = 0;
const ok = (name, got, want = true) => {
  const good = JSON.stringify(got) === JSON.stringify(want);
  console.log(`${good ? 'PASS' : 'FAIL'}  ${name}`);
  if (good) pass++;
  else { console.log(`      得到 ${JSON.stringify(got)}，預期 ${JSON.stringify(want)}`); fail++; }
};

// ── parseSourceGrid ──────────────────────────────────────────────────────────

ok('csv 切成列與欄',
  parseSourceGrid('a,b\n1,2\n', 'csv').sheets[0].rows, [['a', 'b'], ['1', '2']]);

// 這是整個模組最容易錯、錯了又最難發現的一條。
ok('引號裡的逗號不是欄位分隔',
  parseSourceGrid('客戶,地址\nA,"台北市, 信義區"\n', 'csv').sheets[0].rows[1],
  ['A', '台北市, 信義區']);

ok('引號裡的兩個雙引號是一個雙引號',
  parseSourceGrid('a\n"他說""好"""\n', 'csv').sheets[0].rows[1], ['他說"好"']);

ok('結尾換行不會變成一列空的',
  parseSourceGrid('a,b\n1,2\n', 'csv').sheets[0].rows.length, 2);

ok('tsv 用 tab 切', parseSourceGrid('a\tb\n1\t2', 'tsv').sheets[0].rows[1], ['1', '2']);

{
  const r = parseSourceGrid('# 一月\nA\t1\nB\t2\n# 二月\nC\t3', 'xlsx');
  ok('xlsx 依 # 分成多張工作表', r.sheets.map(s => s.name), ['一月', '二月']);
  ok('每張工作表的列都在', r.sheets[0].rows, [['A', '1'], ['B', '2']]);
}

ok('沒有列與欄的格式直接說做不到', parseSourceGrid('一些內文', 'pdf').sheets, []);
ok('而且說得出原因', parseSourceGrid('一些內文', 'pdf').reason.includes('沒辦法逐格比對'));
ok('空檔案也說得出原因', parseSourceGrid('   ', 'csv').reason.includes('空的'));

// ── pickSheet ────────────────────────────────────────────────────────────────

{
  const sheets = [{ name: '封面', rows: [[1]] }, { name: '明細', rows: [[1], [2], [3]] }];
  ok('沒指定就挑資料最多的，不是第一張', pickSheet(sheets).name, '明細');
  ok('指定了就用指定的', pickSheet(sheets, '封面').name, '封面');
  ok('大小寫不同也找得到', pickSheet([{ name: 'Data', rows: [[1]] }], 'data').name, 'Data');
  ok('指定了不存在的就回 null，不要默默挑一張', pickSheet(sheets, '不存在'), null);
}

// ── toNumber / sameValue ─────────────────────────────────────────────────────

ok('千分位是外觀不是數值', toNumber('1,234.50'), 1234.5);
ok('百分比換算成小數', toNumber('12.5%'), 0.125);
ok('會計負號括號', toNumber('(1,200)'), -1200);
ok('空白不是 0——把空白當成 0 會蓋掉真正要找的那幾列', toNumber(''), null);
ok('文字不是數字', toNumber('N/A'), null);

ok('數字跟它的格式化寫法算相同', sameValue(1234.5, '1,234.50'));
ok('前後空白不算差異', sameValue(' 台北 ', '台北'));
ok('大小寫預設算差異', sameValue('ABC', 'abc'), false);
ok('要求忽略大小寫時就不算', sameValue('ABC', 'abc', { ignoreCase: true }));
ok('浮點誤差用容差吃掉', sameValue(0.3, 0.30000000000000004, { tolerance: 1e-9 }));
ok('沒給容差時就是嚴格比', sameValue(0.3, 0.30000000000000004), false);
ok('空白對空白算相同', sameValue('', null));
ok('空白對 0 算差異', sameValue('', 0), false);

// ── alignColumns ─────────────────────────────────────────────────────────────

{
  // 來源檔在中間插了一欄——這是兩個版本最常見的差異形態。
  const cols = alignColumns(['代號', '名稱', '金額'], ['代號', '部門', '名稱', '金額']);
  const byLabel = Object.fromEntries(cols.map(c => [c.label, c]));
  ok('同名欄位對得起來，不是按位置',
    [byLabel['名稱'].cur, byLabel['名稱'].src], [1, 2]);
  ok('金額沒有被右移的欄位帶歪',
    [byLabel['金額'].cur, byLabel['金額'].src], [2, 3]);
  ok('只在來源有的欄位也列出來', byLabel['部門'].cur, null);
}

ok('目前活頁簿有、來源沒有的欄位也留著',
  alignColumns(['a', 'b', 'c'], ['a', 'b']).map(c => [c.cur, c.src]),
  [[0, 0], [1, 1], [2, null]]);

ok('沒有標題列時按位置對，寬度由資料決定',
  alignColumns(null, null, 3).map(c => [c.cur, c.src]), [[0, 0], [1, 1], [2, 2]]);

// headers:false 時沒有標題列可以量寬度。少了資料寬度這條，欄位清單會是空的，
// 兩份完全不同的資料會被回報成「沒有差異」——這個模組最不能給的答案。
{
  const d = diffGrids([['A', '1']], [['A', '2']], { headers: false });
  ok('headers:false 時仍然比得出差異', d.counts.cells, 1);
  ok('headers:false 時列號從第 1 列算', d.diffs[0].row, 1);
}

// ── diffGrids：值的差異 ──────────────────────────────────────────────────────

{
  const cur = [['代號', '金額'], ['A', '100'], ['B', '200']];
  const src = [['代號', '金額'], ['A', '100'], ['B', '250']];
  const d = diffGrids(cur, src, { keyColumn: 0 });
  ok('只回報真正不同的那一格', d.counts.cells, 1);
  ok('相同的列被算成相同', d.counts.sameRows, 1);
  ok('差異記下兩邊的值', [d.diffs[0].cur, d.diffs[0].src], ['200', '250']);
  ok('差異記下欄位名', d.diffs[0].column, '金額');
  ok('差異記得住鍵值', d.diffs[0].key, 'B');
  // 標題列佔掉第 1 列，所以第二筆資料在第 3 列。
  ok('列號算進標題列', d.diffs[0].row, 3);
}

// ── diffGrids：這條是 key 存在的理由 ────────────────────────────────────────

{
  const cur = [['代號', '金額'], ['A', '1'], ['B', '2'], ['C', '3']];
  // 來源檔在最前面多了一列。
  const src = [['代號', '金額'], ['X', '9'], ['A', '1'], ['B', '2'], ['C', '3']];

  const keyed = diffGrids(cur, src, { keyColumn: 0 });
  ok('有鍵：多出來的那列就是一筆新增，其他都沒動', keyed.counts.onlySource, 1);
  ok('有鍵：沒有任何值被誤報成不同', keyed.counts.cells, 0);

  const positional = diffGrids(cur, src, { keyColumn: null });
  ok('沒有鍵：整份都被判成不同（所以預設要用鍵）', positional.counts.cells > 0);
}

// ── diffGrids：整列只在一邊 ─────────────────────────────────────────────────

{
  const cur = [['代號'], ['A'], ['B']];
  const src = [['代號'], ['A'], ['C']];
  const d = diffGrids(cur, src, { keyColumn: 0 });
  ok('只在目前活頁簿的列有算到', d.counts.onlyCurrent, 1);
  ok('只在來源檔的列有算到', d.counts.onlySource, 1);
  const only = d.diffs.find(x => x.status === 'onlySource');
  ok('只在來源檔的列沒有活頁簿位置可指', only.row, null);
}

// ── diffGrids：鍵欄位在來源檔不存在 ─────────────────────────────────────────

{
  const d = diffGrids([['代號', '值'], ['A', '1']], [['名稱', '值'], ['A', '1']], { keyColumn: 0 });
  ok('找不到鍵欄位時退回順序比對，而不是整個失敗', d.keyed, false);
  ok('而且把退回這件事講出來', d.notes.some(n => n.includes('找不到對應的鍵欄位')));
}

// ── diffGrids：重複鍵 ───────────────────────────────────────────────────────

{
  const cur = [['代號', '值'], ['A', '1'], ['A', '2']];
  const src = [['代號', '值'], ['A', '1'], ['A', '9']];
  const d = diffGrids(cur, src, { keyColumn: 0 });
  ok('重複鍵按出現順序配對，第二筆對第二筆', d.counts.cells, 1);
  ok('沒有把重複鍵誤判成新增或刪除',
    [d.counts.onlyCurrent, d.counts.onlySource], [0, 0]);
}

// ── diffGrids：上限 ─────────────────────────────────────────────────────────

{
  const cur = [['k', 'v'], ...Array.from({ length: 20 }, (_, i) => [`k${i}`, 'a'])];
  const src = [['k', 'v'], ...Array.from({ length: 20 }, (_, i) => [`k${i}`, 'b'])];
  const d = diffGrids(cur, src, { keyColumn: 0, limit: 5 });
  ok('超過上限就停', d.diffs.length, 5);
  ok('而且明講被截斷了', d.notes.some(n => n.includes('只列出前 5 筆')));
}

ok('預設上限是模組宣告的那個', MAX_DIFF_ROWS, 5000);

// ── buildReport ─────────────────────────────────────────────────────────────

const oneDiff = () => diffGrids(
  [['代號', '金額'], ['A', '100']],
  [['代號', '金額'], ['A', '120']],
  { keyColumn: 0 },
);

{
  const grid = buildReport(oneDiff(), {
    leftLabel: '損益表', rightLabel: '去年.xlsx', sheetName: '損益表',
    keyLabel: '代號', rangeText: 'A1:B2', firstRow: 1, firstCol: 0,
  });
  const flat = grid.map(r => r.join('|')).join('\n');
  ok('報表寫出比對的兩邊', flat.includes('損益表') && flat.includes('去年.xlsx'));
  ok('報表寫出比對範圍', flat.includes('損益表!A1:B2'));
  ok('報表寫出用哪個鍵對應', flat.includes('以「代號」對應列'));
  ok('表頭用兩邊真正的名字當欄名', flat.includes('狀態|鍵值|欄位|損益表|去年.xlsx|位置'));
  ok('差異那列指得回儲存格', flat.includes('損益表!B2'));
  ok('沒有對照範圍時不印出空的那一列', flat.includes('對照範圍'), false);
}

// 同一份活頁簿裡兩張工作表對比時，「目前活頁簿 vs 來源檔」這種欄名沒有意義
// ——兩邊都是目前活頁簿，看報表的人分不出哪欄是哪張表。
{
  const grid = buildReport(oneDiff(), {
    leftLabel: '一月', rightLabel: '二月', sheetName: '一月',
    keyLabel: '代號', rangeText: 'A1:B2', rightRangeText: '二月!A1:B2',
  });
  const flat = grid.map(r => r.join('|')).join('\n');
  ok('跨工作表時表頭是兩張表的名字', flat.includes('狀態|鍵值|欄位|一月|二月|位置'));
  ok('跨工作表時有印出對照範圍', flat.includes('對照範圍|二月!A1:B2'));
  ok('計數列也用表名，不是「目前活頁簿」', flat.includes('一月 列數') && flat.includes('二月 列數'));
}

{
  const d = diffGrids([['代號'], ['A'], ['B']], [['代號'], ['A'], ['C']], { keyColumn: 0 });
  const grid = buildReport(d, {
    leftLabel: '一月', rightLabel: '二月', sheetName: '一月',
    keyLabel: '代號', rangeText: 'A1:A3',
  });
  const flat = grid.map(r => r.join('|')).join('\n');
  ok('「只在某一邊」的狀態也用真正的表名', flat.includes('只在一月') && flat.includes('只在二月'));
}

{
  // 比對範圍不是從 A1 開始時，位置必須跟著移，否則報表指向錯的儲存格。
  const grid = buildReport(oneDiff(), {
    leftLabel: '表', rightLabel: 's.xlsx', sheetName: '表',
    keyLabel: '代號', rangeText: 'C5:D6', firstRow: 5, firstCol: 2,
  });
  ok('範圍不從 A1 開始時位置有跟著位移',
    grid.some(r => r[5] === '表!D6'));
}

{
  const same = diffGrids([['a'], ['1']], [['a'], ['1']], { keyColumn: 0 });
  const grid = buildReport(same, {
    leftLabel: '表', rightLabel: 's.csv', sheetName: '表', keyLabel: 'a', rangeText: 'A1:A2',
  });
  ok('完全相同時明確寫「沒有差異」，不是留一張空表',
    grid.some(r => r[0] === '沒有差異'));
}

// excel.js 靠找這個字串定位表頭列再套粗體。摘要區每多一條 note 就多一列，
// 所以它不能用數的——這條測試釘住那個錨點還在。
{
  const d = oneDiff();
  d.notes.push('隨便一條注意事項');
  const grid = buildReport(d, {
    leftLabel: '表', rightLabel: 's.csv', sheetName: '表', keyLabel: '代號', rangeText: 'A1:B2',
  });
  const at = grid.findIndex(r => r[0] === REPORT_HEAD_CELL);
  ok('多了一條 note 之後仍然找得到表頭列', at > 0);
  ok('而且找到的確實是表頭列，不是摘要區的某一列', grid[at][5], '位置');
}

console.log(`\n${pass} passed, ${fail} failed`);
if (fail) process.exit(1);
