/**
 * Drive planClean — the rules behind excel_clean_data — with no Excel involved.
 *
 * This is the payoff for making the planner pure. Every one of these cases is a
 * decision about someone's real data (is '0012' a number or a part number? is
 * 113/1/5 a date?), and being able to state them as a table means the answers are
 * pinned rather than re-derived by whoever next reads the code.
 *
 * The confirmation card counts come from this same function, so a wrong number
 * here is a lie on the card — which is the one thing PRD 3.1 says it must not be.
 */
import fs from 'fs';

const src = fs.readFileSync(new URL('../src/taskpane/hosts/excel.js', import.meta.url), 'utf8');
const from = src.indexOf('/* ── clean:start');
const to = src.indexOf('/* ── clean:end');
if (from < 0 || to < 0) {
  console.error('找不到 clean:start / clean:end 標記，hosts/excel.js 的結構變了。');
  process.exit(1);
}
const tmp = new URL('./.clean-under-test.mjs', import.meta.url);
// The text→value helpers moved to compare.js — the comparison engine asks the
// same three questions ("is this a number", "is this a date", "are these the
// same"), and a parser whose entire value is its edge cases must exist once.
// They are imported into the slice and re-exported, so these tests still cover
// them wherever they live.
fs.writeFileSync(tmp,
  "import { toHalfWidth, textToNumber, textToDateSerial } from '../src/taskpane/hosts/compare.js';\n"
  + src.slice(from, to)
  + '\nexport { planClean, textToNumber, textToDateSerial, toHalfWidth };\n');
const { planClean, textToNumber, textToDateSerial } = await import(tmp);

let pass = 0, fail = 0;
const check = (name, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}`);
  if (!ok) { console.log(`      得到 ${JSON.stringify(got)}，預期 ${JSON.stringify(want)}`); fail++; }
  else pass++;
};

/** planClean wants a formulas grid alongside the values; most cases have none. */
const noFormulas = (v) => v.map(r => r.map(() => ''));
const plan = (op, args, values, formulas) => planClean(op, args, values, formulas || noFormulas(values));

// ── 去重 ──────────────────────────────────────────────────────────────────────

{
  const values = [
    ['編號', '姓名', '部門'],
    ['A001', '陳小明', '業務'],
    ['A002', '林小華', '財務'],
    ['A001', '陳小明', '業務'],
    ['A001 ', '陳小明', '業務'],   // 尾巴多一個空白，人看起來是同一筆
  ];
  const p = plan('dedupe', {}, values);
  check('去重刪掉 2 列', p.rows, 2);
  check('去重不動標題列', p.grid[0], ['編號', '姓名', '部門']);
  check('保留第一次出現的那筆', p.grid[1], ['A001', '陳小明', '業務']);
  check('後面的資料往上補', p.grid[2], ['A002', '林小華', '財務']);
  check('尾巴補空白列，塊的大小不變', p.grid.length, values.length);
  check('尾巴是空的', p.grid[3], ['', '', '']);
}

{
  const values = [
    ['編號', '姓名', '金額'],
    ['A001', '陳小明', 100],
    ['A001', '陳小明', 250],       // 只有金額不同
  ];
  check('預設比對整列，金額不同就不算重複', plan('dedupe', {}, values).rows, 0);
  check('指定 key_columns 後只看那幾欄', plan('dedupe', { key_columns: [0] }, values).rows, 1);
}

{
  const values = [['編號', '小計'], ['A001', 100], ['A002', 200]];
  const formulas = [['編號', '小計'], ['A001', '=B2*1.05'], ['A002', 200]];
  const p = plan('dedupe', {}, values, formulas);
  check('範圍裡有公式時拒絕去重', !!p.blocked, true);
  check('拒絕時什麼都不改', p.grid, null);
}

// ── 統一格式 ─────────────────────────────────────────────────────────────────

{
  const values = [
    ['品名', '金額', '日期'],
    ['  蘋果   汁 ', '1,234', '113/1/5'],
    ['ＡＢＣ－１', '(500)', '2024-02-29'],
    ['橘子', '12%', '2024/2/31'],       // 2/31 不存在，不該被硬轉
  ];
  const p = plan('normalize', { half_width: true, to_number: true, to_date: true }, values);
  check('去頭尾空白並壓縮中間的連續空白', p.grid[1][0], '蘋果 汁');
  check('全形轉半形', p.grid[2][0], 'ABC-1');
  check('千分位文字轉成數字', p.grid[1][1], 1234);
  check('會計括號負數', p.grid[2][1], -500);
  check('百分比轉小數', p.grid[3][1], 0.12);
  check('民國年轉成 Excel 日期序號', p.grid[1][2], textToDateSerial('2024/1/5'));
  check('轉出來的日期帶上日期格式', p.numberFormat[1][2], 'yyyy/m/d');
  check('不存在的日期原封不動', p.grid[3][2], '2024/2/31');
}

{
  // ' 88 ' is here so the block has something to change and a grid comes back —
  // the three cells being asserted are precisely the ones that must NOT.
  const values = [['代號'], ['0012'], ['1234567890123456'], ['A1'], [' 88 ']];
  const p = plan('normalize', { to_number: true }, values);
  check('前導零的代號不會被轉成數字', p.grid[1][0], '0012');
  check('超過 15 位數不轉（會被浮點數吃掉精度）', p.grid[2][0], '1234567890123456');
  check('本來就不是數字的放著', p.grid[3][0], 'A1');
  check('該轉的還是有轉', p.grid[4][0], 88);
  check('只算真的動到的格數', p.cells, 1);
}

check('整批沒得改時 grid 是 null',
  plan('normalize', { to_number: true }, [['代號'], ['0012'], ['A1']]).grid, null);

{
  const values = [['A', 'B'], [' x ', ' y ']];
  const p = plan('normalize', { columns: [1] }, values);
  check('指定 columns 時只動那一欄', p.grid[1], [' x ', 'y']);
}

// ── 補齊缺漏 ─────────────────────────────────────────────────────────────────

{
  const values = [
    ['部門', '姓名', '獎金'],
    ['業務', '陳小明', 100],
    ['', '林小華', ''],
    ['', '王小美', 300],
    ['財務', '李小強', ''],
  ];
  const p = plan('fill', { method: 'down', columns: [0] }, values);
  check('往下補承接上面的值', [p.grid[2][0], p.grid[3][0]], ['業務', '業務']);
  check('遇到新的值就換成新的', p.grid[4][0], '財務');
  check('沒指定的欄不動', p.grid[2][2], '');
  check('只算真的補了幾格', p.cells, 2);

  const z = plan('fill', { method: 'value', fill_value: 0, columns: [2] }, values);
  check('填固定值只填空白格', [z.grid[2][2], z.grid[4][2], z.grid[1][2]], [0, 0, 100]);
}

// ── 拆欄 ─────────────────────────────────────────────────────────────────────

{
  const values = [
    ['地址', '', ''],
    ['台北市,信義區,信義路五段', '', ''],
    ['新北市,板橋區', '舊資料', '舊資料'],
    ['桃園市,中壢區,中大路,300號', '', ''],   // 比欄位多
  ];
  const p = plan('split', { delimiter: ',' }, values);
  check('來源欄被第一段取代', p.grid[1][0], '台北市');
  check('其餘各段往右放', [p.grid[1][1], p.grid[1][2]], ['信義區', '信義路五段']);
  check('段數不足時把右邊清乾淨，不留上一列的殘影', [p.grid[2][1], p.grid[2][2]], ['板橋區', '']);
  check('放不下的段數有回報', p.note.includes('1 列'), true);

  const k = plan('split', { delimiter: ',', keep_source: true }, values);
  check('keep_source 時來源欄原封不動', k.grid[1][0], '台北市,信義區,信義路五段');
  check('keep_source 時從右邊一欄開始放', [k.grid[1][1], k.grid[1][2]], ['台北市', '信義區']);
}

check('拆欄範圍只有一欄時拒絕', !!plan('split', {}, [['地址'], ['台北市,信義區']]).blocked, true);

// ── 併欄 ─────────────────────────────────────────────────────────────────────

{
  const values = [
    ['縣市', '區', '路名'],
    ['台北市', '信義區', '信義路五段'],
    ['新北市', '', '中山路'],          // 中間缺一段
  ];
  const p = plan('merge_columns', { separator: '' }, values);
  check('併成一欄放在最左邊', p.grid[1][0], '台北市信義區信義路五段');
  check('中間空白直接跳過，不會併出連續分隔符號', p.grid[2][0], '新北市中山路');
  check('預設清掉被併走的欄', [p.grid[1][1], p.grid[1][2]], ['', '']);

  const k = plan('merge_columns', { separator: '-', keep_source: true }, values);
  check('keep_source 保留原本各欄', k.grid[1][2], '信義路五段');
  check('分隔符號有用上', k.grid[1][0], '台北市-信義區-信義路五段');
}

// ── 邊界 ─────────────────────────────────────────────────────────────────────

check('未知的 op 被擋下來', !!plan('shuffle', {}, [['a'], ['b']]).blocked, true);
check('沒東西可改時 grid 是 null', plan('fill', { method: 'down' }, [['A'], ['x']]).grid, null);
check('textToNumber 不吃空字串', textToNumber(''), null);
check('textToNumber 吃 NT$ 前綴', textToNumber('NT$1,200'), 1200);
check('textToDateSerial 認得 1900-01-01 的序號', textToDateSerial('1900-01-01'), 2);

fs.rmSync(tmp);
console.log(`\n${pass} passed, ${fail} failed`);
if (fail) process.exit(1);
