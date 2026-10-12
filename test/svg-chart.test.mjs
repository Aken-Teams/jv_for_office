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
import fs from 'node:fs';
import { chartSvg, svgProblem, CHART_TYPES, CHART_DATA_KEYS, chartData }
  from '../src/taskpane/hosts/svgchart.js';

/** 一套完整的配色，新圖表都吃得到——缺欄位的 fallback 由上面的 loop 驗。 */
const TH = { ink: '#3D3D3D', strong: '#1A3A5C', second: '#2E6DA4', accent: '#C0392B',
  muted: '#8A93A0', line: '#E4E8EC', bg: '#FFFFFF' };

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

/**
 * Two types do not take `categories` + `values` at all.
 *
 * `timeline` takes `events`, and `radar` refuses fewer than three axes on
 * purpose — a two-axis radar is a line, and drawing one would be worse than
 * saying no. Running them through this loop would be testing that they honour a
 * contract they deliberately do not have. They get their own fixtures further
 * down, exercising the shapes that ARE awkward for them.
 */
/**
 * Five types are excluded, all for one reason: they do not take
 * `categories` + `values`.
 *
 *   timeline → `events`
 *   scatter  → `points`
 *   box      → arrays of raw observations per group, not one number each
 *   map      → region NAMES
 *   radar    → refuses fewer than three axes, on purpose
 *
 * Running them through this loop would be testing that they honour a contract
 * they deliberately do not have. Each gets its own fixtures further down,
 * covering the shapes that ARE awkward for it.
 */
const SHAPE_FUZZ = CHART_TYPES.filter(
  (t) => !['timeline', 'radar', 'map', 'scatter', 'box', 'gantt', 'plot'].includes(t));

for (const type of SHAPE_FUZZ) {
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
// 曾經用 'gantt' 當這裡的代表，後來甘特圖真的做出來了——一個「不支援的類型」的
// 樣本在它被支援的那天就不再是樣本，所以換成一個不會哪天長出來的名字。
try { chartSvg({ type: 'sankey', categories: ['A'], values: [1] }); } catch (e) { msg = e.message; }
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

// ── 瀑布、漏斗、雷達、時間軸 ─────────────────────────────────────────────────
//
// 這四種都是「長條圖畫不出來的問題」。驗的是幾何不變量，不是外觀——沒有瀏覽器
// 可以在這裡渲染，而且會悄悄壞掉的本來就是座標：SVG 裡的 NaN 不會報錯，那個
// 元素直接消失，看起來就像增益集什麼都沒做。

const nums = (svg, attr) => [...svg.matchAll(new RegExp(`${attr}="(-?[\\\\d.]+)"`, 'g'))]
  .map((m) => Number(m[1]));
const sane = (svg) => !/NaN|Infinity|undefined/.test(svg)
  && (svg.match(/</g) || []).length === (svg.match(/>/g) || []).length;

{
  // 走勢是 100 →(+32) →(+18) →(−22) → 128。瀑布圖的內容就是這條路徑。
  const wf = chartSvg({
    type: 'waterfall', title: '營收橋接',
    categories: ['=期初', '新客', '續約', '流失', '=期末'],
    values: [100, 32, 18, -22, 128], total: [0, 4],
  }, TH);
  ok('瀑布圖畫得出來', sane(wf), svgProblem(wf));
  const bars = [...wf.matchAll(/<rect x="([\d.]+)" y="([\d.]+)" width="([\d.]+)" height="([\d.]+)"[^>]*fill="(#[0-9A-Fa-f]{6})"/g)]
    .map((m) => ({ x: +m[1], y: +m[2], w: +m[3], h: +m[4], fill: m[5] }))
    .filter((b) => b.w < 200);           // 排除背景
  ok('五個步驟五根柱子', bars.length === 5, `${bars.length} 根`);
  // 每一根要從上一根結束的地方開始——這是瀑布圖跟長條圖唯一的差別。
  const ends = bars.map((b) => b.y + b.h);
  ok('第二根從第一根的頂端接上去', Math.abs(bars[1].y + bars[1].h - bars[0].y) < 1.5,
    `第一根頂 ${bars[0].y}，第二根底 ${ends[1]}`);
  // 下降要跟上升不同色，否則讀者得自己看數字才知道方向。
  ok('下降跟上升不同色', bars[3].fill !== bars[1].fill, `${bars[3].fill} vs ${bars[1].fill}`);
  ok('合計又是第三種顏色', bars[0].fill !== bars[1].fill && bars[0].fill !== bars[3].fill);
  ok('等號只是標記，不會印出來', !wf.includes('=期初') && wf.includes('期初'));
  ok('增減有正負號', wf.includes('+32') && wf.includes('-22'));
  // 起點 100、終點 128，中間跑到 150——軸要蓋住整條路徑，不是只蓋住數值。
  const ys = bars.flatMap((b) => [b.y, b.y + b.h]);
  ok('所有柱子都在畫布裡', Math.min(...ys) >= 0 && Math.max(...ys) <= 380,
    `${Math.min(...ys)}..${Math.max(...ys)}`);
}

{
  const fn = chartSvg({
    type: 'funnel', title: '導入漏斗',
    categories: ['曝光', '點擊', '試用', '付費'], values: [10000, 3200, 860, 210],
  }, TH);
  ok('漏斗圖畫得出來', sane(fn), svgProblem(fn));
  // 沒有轉換率的漏斗就是一張裝飾用的長條圖：形狀告訴你它變窄了，那是讀者本來
  // 就知道的事；它沒告訴你在哪一段掉最兇。
  ok('每一段之間標出轉換率', (fn.match(/轉換 /g) || []).length === 3, fn.match(/轉換 \d+%/g));
  ok('也標出流失率', fn.includes('流失'));
  ok('轉換率算對', fn.includes('轉換 32%'), fn.match(/轉換 \d+%/g));
  const widths = [...fn.matchAll(/<rect x="[\d.]+" y="[\d.]+" width="([\d.]+)"/g)]
    .map((m) => +m[1]).filter((w) => w < 700);
  ok('寬度逐段變窄', widths.every((w, i) => i === 0 || w <= widths[i - 1] + 0.5), widths.join(' '));
}

{
  const rd = chartSvg({
    type: 'radar', title: '六面向',
    categories: ['資料', '流程', '治理', '人才', '工具', '成效'],
    values: [[4, 3, 2, 3, 5, 2], [5, 5, 4, 4, 5, 4]], series: ['現況', '目標'],
  }, TH);
  ok('雷達圖畫得出來', sane(rd), svgProblem(rd));
  ok('有四層網格 + 兩個資料多邊形',
    (rd.match(/<polygon/g) || []).length === 6, (rd.match(/<polygon/g) || []).length);
  ok('六個面向的標籤都在', ['資料', '流程', '治理', '人才', '工具', '成效']
    .every((k) => rd.includes(k)));
  ok('兩個系列都有圖例', rd.includes('現況') && rd.includes('目標'));
  ok('資料區域半透明，兩層疊起來還看得到下面',
    rd.includes('fill-opacity="0.18"'));

  // 三個以上就分不出來了，第四層會長得像資料但作用是裝飾。
  const many = chartSvg({
    type: 'radar', categories: ['A', 'B', 'C'],
    values: [[1, 2, 3], [2, 3, 1], [3, 1, 2], [1, 1, 1]],
  }, TH);
  ok('最多畫三個系列', (many.match(/fill-opacity/g) || []).length === 3);

  let bad = '';
  try { chartSvg({ type: 'radar', categories: ['A', 'B'], values: [[1, 2]] }, TH); }
  catch (e) { bad = e.message; }
  ok('兩個面向的雷達圖會被擋掉', bad.includes('三個'), bad);
}

{
  // 時間軸的內容就是「間隔不平均」：三月兩件、十一月一件才是重點，平均排開會
  // 把那件事抹掉。
  const tl = chartSvg({
    type: 'timeline', title: '導入時程',
    events: [
      { at: 2024, label: '試點', note: '兩條產線' },
      { at: 2024.75, label: '擴大' },
      { at: 2026, label: '全面上線', highlight: true },
    ],
  }, TH);
  ok('時間軸畫得出來', sane(tl), svgProblem(tl));
  const dots = [...tl.matchAll(/<circle cx="([\d.]+)"/g)].map((m) => +m[1]);
  ok('三個事件三個點', dots.length === 3, `${dots.length} 個`);
  const g1 = dots[1] - dots[0];
  const g2 = dots[2] - dots[1];
  ok('間隔照實際數值，不是平均排開', Math.abs(g1 - g2) > 20, `${g1} vs ${g2}`);
  ok('標註上下交錯，不會擠在同一邊',
    new Set([...tl.matchAll(/<text x="[\d.]+" y="([\d.]+)"/g)].map((m) => +m[1] > 170)).size === 2);
  ok('重點事件會被標出來', tl.includes('r="8"'));

  // 給不了刻度的時候平均排開是誠實的——我們沒有的尺度不能自己發明。
  const words = chartSvg({
    type: 'timeline',
    events: [{ at: '第一階段', label: 'A' }, { at: '第二階段', label: 'B' },
      { at: '第三階段', label: 'C' }],
  }, TH);
  const even = [...words.matchAll(/<circle cx="([\d.]+)"/g)].map((m) => +m[1]);
  ok('非數字的刻度就平均排開',
    Math.abs((even[1] - even[0]) - (even[2] - even[1])) < 1.5, even.join(' '));

  let none = '';
  try { chartSvg({ type: 'timeline', events: [] }, TH); } catch (e) { none = e.message; }
  ok('沒有事件會被擋掉並說要什麼', none.includes('events'), none);
}

// ── 地圖 ────────────────────────────────────────────────────────────────────
//
// 地圖需要真的地理資料——沒有任何算術能生出一條海岸線。所以那份資料是下載一次、
// 簡化過、放進 repo 的。這裡驗的是「它真的在裡面，而且畫得出來」。
{
  const { MAPS } = await import('../src/taskpane/hosts/mapdata.js');
  ok('三份地圖都在', ['world', 'china', 'taiwan'].every((k) => MAPS[k]));
  ok('世界有主要國家', ['China', 'Taiwan', 'Japan', 'United States of America']
    .every((k) => MAPS.world[k]), Object.keys(MAPS.world).length + ' 國');
  // 這兩份是這個產品實際要賣的市場，所以它們必須在。
  ok('中國有省級', !!(MAPS.china['广东'] && MAPS.china['江苏']),
    Object.keys(MAPS.china).length + ' 省');
  ok('台灣有縣市級', !!(MAPS.taiwan['台北'] && MAPS.taiwan['高雄']),
    Object.keys(MAPS.taiwan).length + ' 縣市');
  // 簡化過頭會變成多邊形，簡化不夠檔案會太大。
  const pts = (m) => Object.values(MAPS[m]).flat().reduce((a, r) => a + r.length, 0);
  ok('點數壓在可用範圍內',
    pts('world') < 6000 && pts('china') < 4000 && pts('taiwan') < 2000,
    `world ${pts('world')} china ${pts('china')} taiwan ${pts('taiwan')}`);
  ok('每個環都還是封閉的多邊形',
    Object.values(MAPS.taiwan).flat().every((r) => r.length > 3));
}

for (const which of ['world', 'china', 'taiwan']) {
  const cats = which === 'world' ? ['China', 'Taiwan', 'Japan']
    : which === 'china' ? ['广东', '江苏', '北京'] : ['台北', '高雄', '台中'];
  const svg = chartSvg({ type: 'map', map: which, title: which,
    categories: cats, values: [30, 20, 10] }, TH);
  ok(`${which}：畫得出乾淨的 SVG`, svg.startsWith('<svg') && svg.endsWith('</svg>')
    && !/NaN|undefined|Infinity/.test(svg), svgProblem(svg));
  // 每個地區都要落在畫布裡，不然就是投影或縮放算錯了。
  const pts = [...svg.matchAll(/[ML](-?[\d.]+),(-?[\d.]+)/g)].map((m) => [+m[1], +m[2]]);
  ok(`${which}：全部畫在畫布內`,
    pts.every(([x, y]) => x >= -1 && x <= 761 && y >= -1 && y <= 381),
    `x ${Math.min(...pts.map((q) => q[0])).toFixed(0)}..${Math.max(...pts.map((q) => q[0])).toFixed(0)}`);
  // 給了值的地區要有不同深淺，不然這張圖沒有在表達任何東西。
  const fills = new Set([...svg.matchAll(/<path d="[^"]*" fill="(#[0-9A-Fa-f]{6})"/g)]
    .map((m) => m[1].toUpperCase()));
  ok(`${which}：深淺有分`, fills.size >= 3, [...fills].join(' '));
  ok(`${which}：圖例有地區名和數字`, svg.includes(cats[0]) && svg.includes('30'));
}
{
  // 沒給值的地區畫成最淺色而不是消失——一個洞讀起來是壞掉，淺色讀起來是「沒資料」。
  const svg = chartSvg({ type: 'map', map: 'taiwan', categories: ['台北'], values: [5] }, TH);
  const paths = (svg.match(/<path /g) || []).length;
  ok('沒給值的地區照樣畫出來', paths >= 18, `${paths} 個區`);
}
{
  const svg = chartSvg({ type: 'map', map: 'world', categories: ['Narnia', 'Japan'], values: [1, 2] }, TH);
  ok('名字對不上的地區會講出來', svg.includes('找不到') && svg.includes('Narnia'));
}
{
  let msg = '';
  try { chartSvg({ type: 'map', map: 'mars', categories: ['a'], values: [1] }, TH); } catch (e) { msg = e.message; }
  ok('沒有的地圖會被擋下並列出可用的', msg.includes('world') && msg.includes('taiwan'), msg);
}
{
  // 一個值都沒有也不該爆——底圖本身就是有意義的輸出。
  const svg = chartSvg({ type: 'map', map: 'china' }, TH);
  ok('沒有資料也畫得出底圖', svg.includes('<path') && svgProblem(svg) === '');
}

// ── 配色要真的傳到 renderer ─────────────────────────────────────────────────
//
// chartSvg 原本只把 ink/strong/muted/line/bg 五個角色交給 renderer。新的圖表
// 需要更多：瀑布用 accent 畫下降、時間軸用 accent 標重點、地圖用 tint 畫沒資料
// 的地區。那些角色不在物件裡，於是各自退回這個模組自己的預設——在藍綠色的簡報上
// 畫出一張深藍配大紅的圖，而且沒有任何地方會回報。
{
  const P = { ink: '#111111', strong: '#0F6E72', second: '#4AA7A5', accent: '#E0743A',
    muted: '#7E9694', line: '#CFE3E1', tint: '#E8F4F3', bg: '#FFFFFF' };
  const wf = chartSvg({ type: 'waterfall', categories: ['=初', '增', '減', '=末'],
    values: [100, 30, -20, 110], total: [0, 3] }, P);
  ok('瀑布圖的下降色來自簡報的 accent', wf.includes('#E0743A'));
  ok('瀑布圖的合計色來自簡報的 second', wf.includes('#4AA7A5'));

  const tl = chartSvg({ type: 'timeline',
    events: [{ at: 2024, label: 'A' }, { at: 2026, label: 'B', highlight: true }] }, P);
  ok('時間軸的重點站用簡報的 accent', tl.includes('#E0743A'));

  const mp = chartSvg({ type: 'map', map: 'world', categories: ['Japan'], values: [5] }, P);
  ok('地圖沒資料的地區用簡報的 tint', mp.includes('#E8F4F3'));

  const fn = chartSvg({ type: 'funnel', categories: ['A', 'B'], values: [10, 5] }, P);
  // mixHex 輸出是小寫，比對要不分大小寫。
  ok('漏斗圖的漸層從簡報的 strong 出發', /#0f6e72/i.test(fn), fn.slice(0, 200));
}

// ── 地圖的第二種用法：標示，不是數值 ───────────────────────────────────────
//
// 歷史報告沒有「每個地區一個數字」，所以它永遠不會畫地圖——從外面看就像這個功能
// 不存在。但「這幾個地方」才是投影片最常要地圖做的事。
{
  const P = { strong: '#0F6E72', accent: '#E0743A', tint: '#E8F4F3', ink: '#111', bg: '#FFF',
    muted: '#7E9694', line: '#CFE3E1' };
  const svg = chartSvg({ type: 'map', map: 'world', title: '四大古文明',
    categories: ['Egypt', 'Iraq', 'India', 'China'] }, P);
  ok('沒有數值也畫得出地圖', svg.startsWith('<svg') && svgProblem(svg) === '');
  const fills = new Set([...svg.matchAll(/<path d="[^"]*" fill="(#[0-9A-Fa-f]{6})"/gi)]
    .map((m) => m[1].toUpperCase()));
  // 沒有刻度可以坐落，所以標示就是滿色——再淡一點只會讀成「有點被選到」。
  ok('標示模式只有兩種顏色', fills.size === 2, [...fills].join(' '));
  ok('標示用強調色', fills.has('#E0743A'));
  ok('圖例列出名字', svg.includes('Egypt') && svg.includes('China'));
  // 給的順序通常是時序或敘事順序，重排會把唯一的排序丟掉。
  ok('圖例照給的順序，不重排',
    svg.indexOf('Egypt') < svg.indexOf('China'));
  const withValues = chartSvg({ type: 'map', map: 'world',
    categories: ['Japan', 'China'], values: [5, 50] }, P);
  ok('有數值就回到深淺模式', withValues.includes('50'));
}

// ── 地圖：名字說了算 ───────────────────────────────────────────────────────
//
// 一頁講中國省份的投影片畫出了整張世界地圖、全部空白，底下一行「找不到：广东、
// 山东、河南」。工具照著被告知的做了——map 預設 world，沒有人說別的——結果是
// 用一張全球地圖去說明一個國家的事。
{
  const P = { ink: '#111', strong: '#0F6E72', accent: '#E0743A', tint: '#E8F4F3',
    bg: '#FFF', muted: '#7E9694', line: '#CFE3E1' };
  const cn = chartSvg({ type: 'map', map: 'world',
    categories: ['广东', '山东', '河南', '江苏'], values: [13, 10, 9.9, 8.5] }, P);
  ok('中國省名會自動換成中國地圖', cn.includes('已改用china地圖'));
  ok('而且不再回報找不到', !cn.includes('找不到'));

  const tw = chartSvg({ type: 'map', categories: ['台北', '高雄', '台中'], values: [1, 2, 3] }, P);
  ok('台灣縣市不給 map 也會挑對', tw.includes('已改用taiwan地圖'));

  // 一兩個打錯字不該把一張刻意的世界地圖換掉。
  const w = chartSvg({ type: 'map', map: 'world',
    categories: ['United States of America', 'China', 'Narnia'], values: [1, 2, 3] }, P);
  ok('世界地圖有少數名字對不上時不換', !w.includes('已改用'));
  ok('但仍然把對不上的講出來', w.includes('Narnia'));

  // 「United States of America」原本會直接壓到右邊的數字上。
  const names = [...w.matchAll(/<text x="([\d.]+)"[^>]*>([^<]*)<\/text>/g)]
    .filter((m) => /United|China/.test(m[2]));
  ok('過長的名字會截斷', names.some((m) => m[2].endsWith('…')),
    names.map((m) => m[2]).join(' | '));
  ok('短名字不會被動到', names.some((m) => m[2] === 'China'));
}

// ── 散佈圖、堆疊、箱型 ──────────────────────────────────────────────────────
{
  const P = { ink: '#111', strong: '#0F6E72', second: '#4AA7A5', accent: '#E0743A',
    muted: '#7E9694', line: '#CFE3E1', tint: '#E8F4F3', bg: '#FFFFFF' };

  // 散佈圖是唯一能講「相關性」的——其他每一種都在比大小。
  const sc = chartSvg({ type: 'scatter', title: '規模 vs 成熟度',
    x_axis: '員工數 →', y_axis: '↑ 成熟度',
    points: [{ x: 120, y: 2.1, label: 'A' }, { x: 800, y: 3.4, label: 'B' },
      { x: 2400, y: 4.2, label: 'C' }] }, P);
  ok('散佈圖畫得出來', sc.startsWith('<svg') && svgProblem(sc) === '');
  ok('三個點', (sc.match(/<circle/g) || []).length === 3);
  // 沒有軸名的散佈圖就是兩個沒有名字的數字。
  ok('兩個軸名都畫出來', sc.includes('員工數 →') && sc.includes('↑ 成熟度'));
  ok('點少的時候會標名字', sc.includes('>A<'));
  // 80 到 90 之間的資料從零畫起就是角落一團。
  const tight = chartSvg({ type: 'scatter', x_axis: 'x', y_axis: 'y',
    points: Array.from({ length: 6 }, (_, i) => ({ x: 80 + i, y: 80 + i })) }, P);
  const cys = [...tight.matchAll(/<circle cx="([\d.]+)" cy="([\d.]+)"/g)].map((m) => +m[2]);
  ok('資料離零很遠時軸不從零開始',
    Math.max(...cys) - Math.min(...cys) > 100,
    `垂直展開 ${(Math.max(...cys) - Math.min(...cys)).toFixed(0)}`);
  // 點太多就不標名字，不然是一團字不是一張圖。
  const many = chartSvg({ type: 'scatter', x_axis: 'x', y_axis: 'y',
    points: Array.from({ length: 20 }, (_, i) => ({ x: i, y: i % 7, label: `P${i}` })) }, P);
  ok('點很多就不標名字', !many.includes('>P0<'));
  let bad = '';
  try { chartSvg({ type: 'scatter', values: [1, 2] }, P); } catch (e) { bad = e.message; }
  ok('沒有 points 會被擋下並說要什麼', bad.includes('points'), bad);

  // 堆疊：總和是柱高、組成是分段。
  const st = chartSvg({ type: 'stacked', categories: ['2023', '2024'],
    values: [[10, 20], [5, 10]], series: ['硬體', '服務'] }, P);
  ok('堆疊圖畫得出來', svgProblem(st) === '');
  const segs = [...st.matchAll(/<rect x="([\d.]+)" y="([\d.]+)" width="([\d.]+)" height="([\d.]+)"/g)]
    .map((m) => ({ x: +m[1], y: +m[2], w: +m[3], h: +m[4] })).filter((r) => r.w < 200 && r.h > 5);
  const col1 = segs.filter((r) => Math.abs(r.x - segs[0].x) < 1).sort((a, b) => a.y - b.y);
  ok('同一根的分段互相接續',
    col1.length >= 2 && Math.abs(col1[0].y + col1[0].h - col1[1].y) < 1.5,
    col1.map((r) => `${r.y.toFixed(0)}+${r.h.toFixed(0)}`).join(' '));
  ok('兩個系列都有圖例', st.includes('硬體') && st.includes('服務'));

  const pct = chartSvg({ type: 'stacked100', categories: ['A', 'B'],
    values: [[10, 40], [90, 60]], series: ['x', 'y'] }, P);
  const cols = {};
  for (const m of pct.matchAll(/<rect x="([\d.]+)" y="[\d.]+" width="([\d.]+)" height="([\d.]+)"/g)) {
    if (+m[2] > 200) continue;
    cols[(+m[1]).toFixed(0)] = (cols[(+m[1]).toFixed(0)] || 0) + +m[3];
  }
  const hs = Object.values(cols);
  ok('百分比堆疊每根等高', hs.length >= 2 && Math.abs(hs[0] - hs[1]) < 2, hs.join(' '));
  ok('百分比堆疊標的是百分比', /\d+%/.test(pct));

  // 箱型圖：分布，不只是平均。
  const bx = chartSvg({ type: 'box', title: '交期', categories: ['A 廠', 'B 廠'],
    values: [[5, 6, 6, 7, 7, 8, 9], [2, 3, 5, 6, 9, 12, 30]] }, P);
  ok('箱型圖畫得出來', svgProblem(bx) === '');
  ok('兩個箱子', (bx.match(/fill-opacity="0.22"/g) || []).length === 2);
  // 30 距離四分位距超過 1.5 倍，是離群值。
  ok('離群值畫成空心圓', bx.includes('fill="none"'));
  ok('兩組名稱都在', bx.includes('A 廠') && bx.includes('B 廠'));
  // 中位數是唯一會被讀出來的數字，所以它是箱子上最重的一筆。
  ok('中位數線最粗', bx.includes('stroke-width="2.6"'));
  const explicit = chartSvg({ type: 'box', categories: ['X'],
    boxes: [{ min: 1, q1: 3, median: 5, q3: 7, max: 9, outliers: [15] }] }, P);
  ok('也吃現成的五數摘要', svgProblem(explicit) === '' && explicit.includes('<rect'));
  bad = '';
  try { chartSvg({ type: 'box', categories: ['A'], values: [1, 2, 3] }, P); } catch (e) { bad = e.message; }
  ok('一維陣列會被擋下並說清楚格式', bad.includes('[[...]'), bad);
}

/**
 * 圖例不能把地圖擠掉。
 *
 * 「地圖有點小，反而圖例太大」——固定 150pt 在半頁寬的圖上就是三分之一，而那塊
 * 地理是用地圖的唯一理由。改成跟著畫布寬度走，下限留在放得下一個截短的地名。
 */
{
  const mapOf = (w, h) => {
    const s = chartSvg({ type: 'map', map: 'taiwan', title: '縣市', width: w, height: h,
      categories: ['新北', '高雄', '台中', '台北'], values: [397, 278, 275, 246] }, TH);
    const pts = [...s.matchAll(/[ML](-?[\d.]+),(-?[\d.]+)/g)].map((m) => [+m[1], +m[2]]);
    return {
      svg: s,
      w: Math.max(...pts.map((p) => p[0])) - Math.min(...pts.map((p) => p[0])),
      lowest: Math.max(...[...s.matchAll(/<text x="[\d.]+" y="([\d.]+)"/g)].map((m) => +m[1])),
    };
  };
  const small = mapOf(430, 300);
  const big = mapOf(900, 460);
  ok('窄畫布的圖例會縮', /font-size="10.5"/.test(small.svg));
  ok('寬畫布的圖例不縮', /font-size="12"/.test(big.svg));
  ok('圖例沒有掉出畫布', small.lowest <= 300, `${small.lowest}`);
  ok('地圖拿到的寬度比圖例多', small.w > 430 * 0.3, `${small.w.toFixed(0)}`);
}

/**
 * 甘特圖——為什麼不用 mermaid 畫。
 *
 * 「我不建議用 mermaid 來呈現」，而那張截圖把理由講完了：mermaid 的 gantt 把**任務
 * 名寫在長條上**，所以只要排程有兩件事時間重疊，字就疊在一起；而且它有幾天印幾天，
 * 最後整排日期擠成一團。
 *
 * 兩個失敗是同一個錯誤——把資料裡有的東西全部畫出來，然後祈禱它放得下。所以這裡釘
 * 的就是那兩個相反的決定：**名字在左欄不在長條上**、**日期標籤的數量由寬度決定**。
 */
{
  const TH = { ink: '#3D3D3D', strong: '#1A3A5C', accent: '#C0392B', muted: '#8A93A0',
    line: '#D5DBE0', tint: '#EEF3F8', bg: '#FFFFFF' };
  const PLAN = [
    { name: '需求確認與設計', group: 'MVP (P0)', start: '2026-10-13', end: '2026-10-24' },
    { name: '編輯器與題型', group: 'MVP (P0)', start: '2026-10-27', end: '2026-11-14' },
    { name: '發佈與填答頁', group: 'MVP (P0)', start: '2026-11-03', end: '2026-11-21' },
    { name: '外觀主題與 AI 出題', group: 'V1.1 (P1)', start: '2026-12-15', end: '2027-01-09' },
    { name: '測驗模式與 PDF 報告', group: 'V1.2 (P2)', start: '2027-02-02', end: '2027-03-06' },
  ];
  const svg = chartSvg({ type: 'gantt', title: '版本規劃與時程', width: 700,
    today: '2026-12-20', tasks: PLAN }, TH);
  ok('甘特圖畫得出乾淨的 SVG', !svgProblem(svg), svgProblem(svg));

  // 任務名完整出現在左欄——這是 mermaid 那張圖最明顯的毛病。
  const labels = [...svg.matchAll(/<text[^>]*>([^<]+)</g)].map((m) => m[1]);
  for (const t of PLAN) {
    ok(`「${t.name}」完整畫出來，沒有被截斷`, labels.includes(t.name),
      labels.filter((l) => l.startsWith(t.name.slice(0, 3))).join('|'));
  }
  ok('階段名稱變成色帶標籤', ['MVP (P0)', 'V1.1 (P1)', 'V1.2 (P2)'].every((g) => labels.includes(g)));

  // 日期標籤數量是算出來的，不是有幾天印幾天。
  const dates = labels.filter((l) => /^\d{4}\/\d+\/\d+$|^\d+\/\d+$/.test(l));
  ok('日期標籤沒有多到擠成一團', dates.length <= Math.floor(700 / 42), `${dates.length} 個`);
  ok('日期標籤也不會少到看不出時間軸', dates.length >= 3, `${dates.length} 個`);
  // 跨年的時程要帶年份，不然 12/22 和 1/5 看起來像差兩週。
  ok('跨年的時程會帶年份', dates.every((d) => /^\d{4}\//.test(d)), dates.join(' '));

  ok('今天在範圍內就畫線', labels.includes('今天'));
  const away = chartSvg({ type: 'gantt', width: 700, today: '2030-01-01', tasks: PLAN }, TH);
  // 一條釘在左邊緣的紅線，對一份講明年的圖表什麼也沒說，只像畫壞了。
  ok('今天不在範圍內就不畫', !/今天/.test(away));
}

{
  const TH = { ink: '#3D3D3D', strong: '#1A3A5C', bg: '#FFFFFF' };
  // 高度跟著列數走——20 列塞進 380 點是一列六點，那是一團糊。
  const five = chartSvg({ type: 'gantt', width: 600, tasks: Array.from({ length: 5 },
    (_, i) => ({ name: `T${i}`, start: '2026-01-01', days: 5 })) }, TH);
  const twenty = chartSvg({ type: 'gantt', width: 600, tasks: Array.from({ length: 20 },
    (_, i) => ({ name: `T${i}`, start: '2026-01-01', days: 5 })) }, TH);
  const h = (x) => Number(/height="(\d+)"/.exec(x)[1]);
  ok('高度跟著工作數量長', h(twenty) > h(five), `${h(five)} → ${h(twenty)}`);

  // days 是「做幾天」的簡寫，跟 end 擇一。
  const byDays = chartSvg({ type: 'gantt', width: 400,
    tasks: [{ name: 'a', start: '2026-01-01', days: 10 }] }, TH);
  ok('days 可以代替 end', !svgProblem(byDays), svgProblem(byDays));

  // 一天的工作要看得見，不能是一條髮絲。
  const oneDay = chartSvg({ type: 'gantt', width: 800,
    tasks: [{ name: '甲', start: '2026-01-01', end: '2026-01-01' },
      { name: '乙', start: '2026-01-02', end: '2026-06-30' }] }, TH);
  const widths = [...oneDay.matchAll(/<rect[^>]*width="([\d.]+)"[^>]*rx="3"/g)].map((m) => Number(m[1]));
  ok('一天的工作仍然看得見', widths.length > 0 && Math.min(...widths) >= 4,
    widths.join('|'));

  let threw = '';
  try { chartSvg({ type: 'gantt', tasks: [] }, TH); } catch (e) { threw = e.message; }
  ok('沒有 tasks 會說清楚格式', /YYYY-MM-DD/.test(threw), threw);
  try { chartSvg({ type: 'gantt', tasks: [{ name: 'x', start: '明天' }] }, TH); } catch (e) { threw = e.message; }
  ok('日期格式不對也會說清楚', /YYYY-MM-DD/.test(threw), threw);
}

/**
 * 數學圖：y 對 x，兩條都是真的數值軸。
 *
 * ── 為什麼其他圖做不到 ──
 * 這個模組在此之前每一種圖的 X 軸都是**分類軸**——位置是名字平均排開的，對季度和
 * 產品是對的，對數學是錯的。N² 和 N log N 畫在平均刻度上，那張圖唯一要講的「差距
 * 會越拉越開」正好被平均刻度藏起來。
 */
{
  const TH = { ink: '#000', strong: '#1F3864', second: '#595959', accent: '#C0392B',
    muted: '#595959', line: '#D9D9D9', tint: '#F2F2F2', bg: '#FFF' };
  const N = [256, 1024, 4096, 16384, 65536];
  const svg = chartSvg({ type: 'plot', title: 'FFT 與直接計算的運算量', width: 640,
    log_x: true, log_y: true, x_axis: 'N（點數）', y_axis: '複數乘法次數',
    curves: [
      { name: '直接計算 N²', points: N.map((n) => [n, n * n]), dots: true },
      { name: 'FFT N log N', points: N.map((n) => [n, n * Math.log2(n)]), dashed: true },
    ] }, TH);
  ok('數學圖畫得出乾淨的 SVG', !svgProblem(svg), svgProblem(svg));
  ok('兩條曲線都畫了', (svg.match(/<path /g) || []).length === 2);
  ok('兩條線就要有圖例', /直接計算/.test(svg) && /FFT/.test(svg));
  ok('虛線畫成虛線', /stroke-dasharray/.test(svg));
  // 數值軸不會自己說明自己，軸名在這裡比在分類圖上更要緊。
  ok('軸名有標出來', /N（點數）/.test(svg) && /複數乘法次數/.test(svg));
  ok('對數刻度要講明', (svg.match(/對數刻度/g) || []).length === 2);

  // 刻度標的是真實數值，不是它的對數——寫 4.2 沒有人看得懂那是 16384。
  const ticks = [...svg.matchAll(/font-size="11"[^>]*>([^<]+)</g)].map((m) => m[1]);
  ok('刻度標的是真實數值不是對數', ticks.some((t) => /[k萬億]|\d{3}/.test(t)), ticks.join(' '));

  // 對數軸碰到 0 或負數就不能畫，要先講而不是畫出 -Infinity。
  let threw = '';
  try {
    chartSvg({ type: 'plot', log_y: true,
      curves: [{ name: 'x', points: [[1, 0], [2, 5]] }] }, TH);
  } catch (e) { threw = e.message; }
  ok('log_y 遇到 0 會說清楚', /大於 0/.test(threw), threw);

  try { chartSvg({ type: 'plot', curves: [] }, TH); } catch (e) { threw = e.message; }
  ok('沒有 curves 會說清楚格式', /points/.test(threw), threw);
  try {
    chartSvg({ type: 'plot', curves: [{ name: 'a', points: [[1, 1]] }] }, TH);
  } catch (e) { threw = e.message; }
  ok('一個點不成一條線', /至少要兩個點/.test(threw), threw);

  // 線性軸：資料靠近零就從零起，離零很遠就不要硬拉——跟散佈圖同一條規則。
  const near = chartSvg({ type: 'plot', width: 400,
    curves: [{ name: 'a', points: [[0, 1], [10, 9]] }] }, TH);
  ok('線性軸資料靠近零時從零起', !svgProblem(near), svgProblem(near));
}

/**
 * 一個系列不該有圖例，而且不該蓋在 X 軸標籤上。
 *
 * ── 回報的畫面 ──
 * 一張長條圖的底部，「加速倍數」連同它的色塊壓在「N=1,024」和「N=4,096」上面。
 *
 * ── 真正的原因 ──
 * `frame()` 在 `series.length > 1` 時才留圖例的空間，而 `legend()` 只要系列**有
 * 名字**就畫——所以單一具名系列會拿到一個沒有地方放的圖例。那不是版面擠不下，是
 * 兩個函式對「到底有沒有圖例」各說各話。
 */
{
  const TH = { ink: '#000', strong: '#1F3864', muted: '#595959', line: '#D9D9D9',
    tint: '#F2F2F2', bg: '#FFF' };
  const one = chartSvg({ type: 'column', title: '加速倍數', width: 640,
    categories: ['N=256', 'N=1,024', 'N=4,096'], values: [32, 102, 341],
    series: ['ZZONLYZZ'] }, TH);
  // 標題已經講完這張圖在講什麼了，再放一個色塊配同一句話只是噪音。
  ok('單一系列不畫圖例', !/ZZONLYZZ/.test(one));

  const two = chartSvg({ type: 'column', width: 640, categories: ['A', 'B'],
    values: [[1, 2], [3, 4]], series: ['ZZAZZ', 'ZZBZZ'] }, TH);
  ok('兩個系列照樣有圖例', /ZZAZZ/.test(two) && /ZZBZZ/.test(two));
}

/**
 * 時間軸的標籤要留在畫布內，而且同一側不能疊。
 *
 * ── 回報的畫面 ──
 * 一張傅立葉沿革的時間軸，第一個事件的字跑到畫面左邊外面，最後一個跑到右邊外面，
 * 中間兩個互相壓在一起。
 *
 * ── 原因有兩層 ──
 * 標籤一律置中在刻度上，而頭尾的刻度離邊只有 24 點——一句 12 個中文字的標籤是
 * 150 點寬，有一半本來就在畫面外。第二層更難看出來：第一次修的時候用「平均間距
 * × 1.7」當可用寬度，但時間軸的間距本來就不平均（1750、1805、1822 擠在 215 年
 * 的前三分之一），平均值會承諾每個標籤兩倍於它實際擁有的空間。
 *
 * 所以這裡直接量：把每個 <text> 的左右邊界算出來，比對畫布和同側的鄰居。這種事
 * 看程式碼看不出來，只有量過才知道。
 */
{
  const TH = { ink: '#000', strong: '#1F3864', accent: '#C00', muted: '#595959',
    line: '#D9D9D9', tint: '#F2F2F2', bg: '#FFF' };
  // 使用者截圖裡的那份資料，連字數都一樣。
  const events = [
    { label: '白努利、尤拉研究三角級數', at: '1750' },
    { label: '高斯使用快速演算法（未發表）', at: '1805' },
    { label: '傅立葉《熱的解析理論》', at: '1822' },
    { label: 'Danielson 與 Lanczos 提出 N log N 方法', at: '1942' },
    { label: 'Cooley 與 Tukey 發表 FFT', at: '1965' },
  ];
  const measure = (txt, size) => [...txt]
    .reduce((a, c) => a + (/[\u2E80-\uFF60]/.test(c) ? size : size * 0.55), 0);

  // 三種寬度：Word 內文欄寬、投影片常用寬、以及窄到必須斷行的情況。
  for (const W of [760, 640, 460]) {
    const H = Math.round(W * 0.47);
    const svg = chartSvg({ type: 'timeline', title: '傅立葉與 FFT 的發展',
      events, width: W, height: H }, TH);
    const boxes = [];
    for (const m of svg.matchAll(
      /<text x="([-\d.]+)" y="([-\d.]+)"[^>]*font-size="([\d.]+)"[^>]*text-anchor="(\w+)"[^>]*>([^<]*)</g)) {
      const [, xs, ys, sz, anchor, txt] = m;
      const x = Number(xs); const size = Number(sz);
      const w = measure(txt, size);
      const l = anchor === 'start' ? x : (anchor === 'end' ? x - w : x - w / 2);
      boxes.push({ txt, size, l, r: l + w, y: Number(ys) });
    }
    ok(`${W} 點寬：有量到標籤`, boxes.length >= events.length * 2, boxes.length);

    const out = boxes.filter((b) => b.l < -1 || b.r > W + 1 || b.y < 0 || b.y > H);
    ok(`${W} 點寬：沒有字跑到畫布外`, !out.length,
      out.map((b) => `${b.txt} ${b.l.toFixed(0)}~${b.r.toFixed(0)}`).join(' / '));

    // 同一行（y 差 14 點以內）而且左右有交集，就是疊在一起。
    const words = boxes.filter((b) => b.size > 12);
    const hit = [];
    for (let i = 0; i < words.length; i++) {
      for (let j = i + 1; j < words.length; j++) {
        const a = words[i]; const b = words[j];
        if (Math.abs(a.y - b.y) < 14 && a.l < b.r && b.l < a.r) hit.push(`${a.txt}／${b.txt}`);
      }
    }
    ok(`${W} 點寬：同側標籤沒有互疊`, !hit.length, hit.join(' / '));
  }

  // 放得下就不要斷行——第一版修完以後它連 760 點寬都還在斷，因為中點切法把一半
  // 的間距讓給了根本不需要的鄰居。
  const wide = chartSvg({ type: 'timeline', events, width: 760, height: 357 }, TH);
  ok('空間夠就整句一行', /白努利、尤拉研究三角級數</.test(wide));

  // 窄到真的放不下，就斷行而不是畫成一團。
  const narrow = chartSvg({ type: 'timeline', events, width: 420, height: 220 }, TH);
  ok('窄的時候改斷行', !/白努利、尤拉研究三角級數</.test(narrow)
    && /白努利/.test(narrow) && /級數/.test(narrow));
}

/**
 * 每一個渲染器會讀的欄位，CHART_DATA_KEYS 都要有。
 *
 * ── 這條測試擋的是什麼 ──
 * `timeline` 曾經整個不能用：工具的列舉裡有它、說明裡有 `events` 的範例，但 host
 * 組給 `chartSvg` 的物件是一個欄位一個欄位手寫的，漏了 `events`。接著 `plot`、
 * `box`、`map`、`scatter` 在兩個 host 上又各漏一次。
 *
 * 症狀最惡劣的地方在於：模型送的參數是對的，渲染器卻回報「curves 是必要的」，而
 * 兩邊都沒有指向那個漏抄的地方。所以改成一份清單，並且在這裡用原始碼反查——新增
 * 圖表型別時漏掉欄位，這條會先紅。
 */
{
  const src = fs.readFileSync(new URL('../src/taskpane/hosts/svgchart.js', import.meta.url), 'utf8');
  const read = [...new Set([...src.matchAll(/spec\.([a-z_]+)/g)].map((m) => m[1]))];
  // host 自己決定的那幾個不經過資料清單。
  const own = new Set(['type', 'title', 'width', 'height', 'transparent']);
  const missing = read.filter((k) => !own.has(k) && !CHART_DATA_KEYS.includes(k));
  ok('渲染器讀的欄位都在 CHART_DATA_KEYS 裡', !missing.length, missing.join(', '));

  // 反過來也要擋：清單裡列了沒人讀的欄位，host 會白抄，說明也會騙人。
  const unread = CHART_DATA_KEYS.filter((k) => !read.includes(k));
  ok('CHART_DATA_KEYS 沒有多餘的欄位', !unread.length, unread.join(', '));

  const picked = chartData({ type: 'timeline', events: [1], curves: [2], width: 9, junk: 1 });
  ok('chartData 只挑資料欄位',
    picked.events && picked.curves && picked.width === undefined && picked.junk === undefined,
    JSON.stringify(picked));
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
