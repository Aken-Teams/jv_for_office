/**
 * Pin the one line of excel_get_overview that tells the agent a picture exists.
 *
 * Small, but it is the whole of the "stop being blind without knowing it" fix:
 * before this line, a monthly report that was half a pasted screenshot got a
 * confident answer about the other half and no hint anything was missing. The
 * two things worth pinning are that it stays silent when there is nothing to
 * say — a line on every sheet of every workbook is noise nobody reads — and
 * that a sheet full of decorations does not push the actual data preview out of
 * the model's context.
 */
import fs from 'fs';

const src = fs.readFileSync(new URL('../src/taskpane/hosts/excel.js', import.meta.url), 'utf8');
const from = src.indexOf('/* ── shapes:start');
const to = src.indexOf('/* ── shapes:end');
if (from < 0 || to < 0) {
  console.error('找不到 shapes:start / shapes:end 標記，hosts/excel.js 的結構變了。');
  process.exit(1);
}
const tmp = new URL('./.shapes-under-test.mjs', import.meta.url);
fs.writeFileSync(tmp, src.slice(from, to) + '\nexport { describeShapes, MAX_SHAPES_LISTED };\n');
const { describeShapes, MAX_SHAPES_LISTED } = await import(tmp);

let pass = 0, fail = 0;
const check = (name, got, want) => {
  const ok = JSON.stringify(got) === JSON.stringify(want);
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}`);
  if (!ok) { console.log(`      得到 ${JSON.stringify(got)}，預期 ${JSON.stringify(want)}`); fail++; }
  else pass++;
};

const shape = (name, type = 'Image', width = 320, height = 240) => ({ name, type, width, height });
const collection = (items) => ({ items });

check('沒有圖片就完全不出聲', describeShapes(collection([])), '');
check('連載入都沒發生（舊版 Excel）也不出聲', describeShapes(null), '');

{
  const line = describeShapes(collection([shape('圖片 1'), shape('文字方塊 2', 'GeometricShape', 180.4, 60.6)]));
  check('講出總數', line.includes('2 個圖片／圖案'), true);
  check('名稱是拿來給 excel_read_image 用的，所以要原樣列出', line.includes('「圖片 1」'), true);
  check('類型翻成人看得懂的字', line.includes('圖案／文字方塊'), true);
  check('尺寸四捨五入，不要小數點', line.includes('180×61'), true);
  check('明說看不到內容並指路', line.includes('excel_read_image'), true);
}

{
  // 一張表放 20 個裝飾圖案是有的事，全部列出來會把真正的資料預覽擠出去。
  const many = Array.from({ length: MAX_SHAPES_LISTED + 5 }, (_, i) => shape(`裝飾 ${i + 1}`));
  const line = describeShapes(collection(many));
  check('總數仍然誠實', line.includes(`${MAX_SHAPES_LISTED + 5} 個`), true);
  check('只點名前幾個', line.includes(`「裝飾 ${MAX_SHAPES_LISTED}」`), true);
  check('超過的不點名', line.includes(`「裝飾 ${MAX_SHAPES_LISTED + 1}」`), false);
  check('但有講還有幾個沒列', line.includes('另有 5 個未列出'), true);
}

fs.rmSync(tmp);
console.log(`\n${pass} passed, ${fail} failed`);
if (fail) process.exit(1);
