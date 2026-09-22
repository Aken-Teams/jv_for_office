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
// `map` joins them: it takes region NAMES, so feeding it 「一月／二月」 would be
// testing that it honours a contract it does not have.
const SHAPE_FUZZ = CHART_TYPES.filter((t) => !['timeline', 'radar', 'map'].includes(t));

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
try { chartSvg({ type: 'gantt', categories: ['A'], values: [1] }); } catch (e) { msg = e.message; }
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

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
