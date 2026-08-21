/**
 * Drive the SVG chart renderer against the shapes of data that break charts.
 *
 * ── What this is guarding ──
 * PowerPoint has no chart API, so these are drawn by hand — which means the
 * arithmetic is ours, and the classic way for that arithmetic to fail is a NaN
 * coordinate. An SVG renderer treats `x="NaN"` the way canvas treats it: the
 * element is silently dropped. The chart comes out missing a bar, or empty, and
 * from inside PowerPoint it looks exactly like "the add-in did nothing".
 *
 * No amount of reading the code catches an axis that collapses on one particular
 * shape of data, so every awkward shape gets driven through here instead:
 *
 *   all values identical · all zeros · one category · one series · many series ·
 *   negatives · huge and tiny in the same chart · a pie with a single slice
 *
 * The assertion is blunt on purpose. It does not check that the chart is pretty;
 * it checks that the markup is well-formed and every coordinate is a real
 * number, which is the difference between a chart and a blank rectangle.
 */
import { chartSvg, svgProblem, CHART_TYPES } from '../src/taskpane/hosts/svgchart.js';

let pass = 0, fail = 0;
const ok = (name, cond, detail) => {
  cond ? pass++ : fail++;
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}${cond ? '' : '  → ' + (detail || '')}`);
};

// ── Every type × every awkward shape ─────────────────────────────────────────

const SHAPES = [
  ['一般資料', { categories: ['一月', '二月', '三月'], values: [10, 25, 18] }],
  ['全部一樣', { categories: ['A', 'B', 'C'], values: [7, 7, 7] }],
  ['全部是零', { categories: ['A', 'B', 'C'], values: [0, 0, 0] }],
  ['只有一個分類', { categories: ['唯一'], values: [42] }],
  ['很多分類', { categories: Array.from({ length: 14 }, (_, i) => `第${i + 1}季`), values: Array.from({ length: 14 }, (_, i) => i * 3) }],
  ['多個系列', { categories: ['A', 'B'], values: [[1, 2], [3, 4], [5, 6]], series: ['甲', '乙', '丙'] }],
  ['負數', { categories: ['A', 'B'], values: [-5, 12] }],
  ['大小懸殊', { categories: ['小', '大'], values: [0.4, 9800000] }],
  ['缺值的短列', { categories: ['A', 'B', 'C'], values: [[1], [2, 3, 4]] }],
  ['非數字混進來', { categories: ['A', 'B'], values: ['abc', 5] }],
  ['沒有分類名稱', { values: [3, 6, 9] }],
];

for (const type of CHART_TYPES) {
  let bad = 0;
  const notes = [];
  for (const [label, spec] of SHAPES) {
    let svg;
    try {
      svg = chartSvg({ ...spec, type, title: `${type} · ${label}` });
    } catch (e) {
      bad++; notes.push(`${label}：丟出 ${e.message}`); continue;
    }
    const problem = svgProblem(svg);
    if (problem) { bad++; notes.push(`${label}：${problem}`); continue; }
    if (!svg.startsWith('<svg') || !svg.endsWith('</svg>')) { bad++; notes.push(`${label}：不是完整的 SVG`); }
  }
  ok(`${type}：11 種資料形狀都畫得出乾淨的 SVG`, bad === 0, notes.join(' / '));
}

// ── The things a chart on a slide has to get right ───────────────────────────

const bar = chartSvg({ type: 'column', categories: ['一月', '二月'], values: [10, 20], title: '營收' });
ok('標題會畫出來', bar.includes('營收'));
ok('分類名稱會畫出來', bar.includes('一月') && bar.includes('二月'));
ok('尺寸寫進 svg 屬性', /width="760"/.test(bar) && /height="380"/.test(bar),
  'PowerPoint 用這個決定插入後的大小，沒有的話會落在預設尺寸要手動調');
ok('有 viewBox', /viewBox="0 0 760 380"/.test(bar));
// Chinese must survive as literal characters. Numeric entities (&#x4E00;)
// would still render, but PowerPoint's 轉換為圖形 turns the SVG into real text
// boxes — and entities there come out as the escape sequence, not the word.
ok('中文以原字元寫入，不是數值實體', /<text[^>]*>一月<\/text>/.test(bar) && !/&#x?\d/i.test(bar),
  bar.slice(bar.indexOf('一月') - 60, bar.indexOf('一月') + 12));

// XML injection through a label would produce markup, not text.
const nasty = chartSvg({ type: 'column', categories: ['<script>x</script>', 'B&B'], values: [1, 2], title: '"引號" & <標籤>' });
ok('標籤裡的 XML 會被逸出', !nasty.includes('<script>') && nasty.includes('&lt;script&gt;'),
  '圖表資料可能來自使用者上傳的檔案，不能當標記處理');
ok('& 和引號都逸出了', nasty.includes('&amp;') && nasty.includes('&quot;'));
ok('逸出後仍然是合法的 SVG', svgProblem(nasty) === '', svgProblem(nasty));

// A "nice" axis is the whole reason a chart is legible from the back of a room.
const axis = chartSvg({ type: 'column', categories: ['A'], values: [37] });
ok('座標軸取整齊的刻度', /">40</.test(axis) || /">50</.test(axis),
  '37 的軸頂應該是 40 或 50，不是 37——不然刻度會變成 0/9.25/18.5');

const bigNums = chartSvg({ type: 'column', categories: ['A'], values: [12345678] });
ok('大數字會縮寫', /萬|億/.test(bigNums), '12345678 不該原樣印在軸上');

// ── Pie: the degenerate cases that produce an invisible chart ────────────────

const onePie = chartSvg({ type: 'pie', categories: ['全部'], values: [100] });
ok('單一切片畫成整個圓而不是空的弧',
  onePie.includes('<circle') && svgProblem(onePie) === '',
  '起點和終點重合時弧線路徑會塌掉，必須改畫圓');
ok('圓餅圖有百分比', /100%/.test(onePie));

const zeroPie = chartSvg({ type: 'pie', categories: ['A', 'B'], values: [0, 0] });
ok('全零的圓餅圖說明沒有資料而不是畫出空白', zeroPie.includes('沒有資料'), svgProblem(zeroPie));

const donut = chartSvg({ type: 'donut', categories: ['是', '否'], values: [7, 3], center: '70%', center_label: '採用率' });
ok('甜甜圈有中心文字', donut.includes('70%') && donut.includes('採用率'));
ok('甜甜圈有挖空的內圓', (donut.match(/<circle/g) || []).length >= 1);

// ── Refusals ─────────────────────────────────────────────────────────────────

let msg = '';
try { chartSvg({ type: 'radar', categories: ['A'], values: [1] }); } catch (e) { msg = e.message; }
ok('不支援的類型會被擋掉並列出可用的', msg.includes('column') && msg.includes('pie'), msg);

msg = '';
try { chartSvg({ type: 'column', categories: ['A'] }); } catch (e) { msg = e.message; }
ok('沒有 values 會被擋掉', msg.includes('values'), msg);

// ── The theme reaches the chart ──────────────────────────────────────────────

const themed = chartSvg({ type: 'column', categories: ['A'], values: [1], title: 'T' },
  { strong: '#AA0011', ink: '#112233', muted: '#445566', line: '#778899', bg: '#FEFEFE' });
ok('配色會套進圖表', themed.includes('#AA0011') && themed.includes('#445566'),
  '圖表跟投影片同一套色才不會看起來是貼上去的');
ok('背景用配色的底色', themed.includes('#FEFEFE'));

const clear = chartSvg({ type: 'column', categories: ['A'], values: [1], transparent: true });
ok('可以要求透明背景', /fill="none"/.test(clear.slice(0, 300)));

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
