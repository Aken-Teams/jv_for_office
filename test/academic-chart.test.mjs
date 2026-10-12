/**
 * 學術文件的那組圖。
 *
 * ── 為什麼要單獨一支 ──
 * 「WORD 就是比較偏向幾種應用，一種是學術文件，一種是企業內部的文件……學術文件
 * 的圖偏少」。上面那十七種回答的是商業問題：誰賣得多、錢花到哪、專案何時結束。
 * 論文問的是別的——這東西怎麼分布、這兩件事會不會一起動、差異有沒有大過雜訊——
 * 而那些問題用商業圖表湊不出來。
 *
 * 這支測的是「有沒有照統計的規矩畫」，不是「畫得好不好看」：
 *
 *   · 直方圖吃的是原始數值，不是數好的次數
 *   · 森林圖的無效果線、比值型要用對數軸
 *   · 存活曲線是階梯，不是直線
 *   · 誤差棒不可以被軸切掉
 *   · 氣泡大小要按面積，不是按半徑
 *
 * 最後一條是這類圖最常見的謊：半徑加倍，墨水變四倍，讀者就把兩倍看成四倍。
 */
import fs from 'node:fs';
import { chartSvg, svgProblem, CHART_TYPES } from '../src/taskpane/hosts/svgchart.js';

let pass = 0;
let fail = 0;
function ok(name, cond, detail) {
  if (cond) { pass++; console.log(`PASS  ${name}`); } else {
    fail++; console.log(`FAIL  ${name}${detail === undefined ? '' : `  → ${detail}`}`);
  }
}

const TH = {
  ink: '#000000', strong: '#1F3864', second: '#595959', accent: '#C0392B',
  muted: '#595959', line: '#D9D9D9', tint: '#F2F2F2', ground: '#F7F7F7', bg: '#FFFFFF',
};

/** 固定種子的常態亂數——測試不可以每次跑出不一樣的資料。 */
function sample(n, mu, sd, seed = 7) {
  let st = seed;
  const rnd = () => { st = (st * 1103515245 + 12345) & 0x7fffffff; return st / 0x7fffffff; };
  return Array.from({ length: n }, () => mu
    + sd * Math.sqrt(-2 * Math.log(rnd() || 1e-9)) * Math.cos(2 * Math.PI * rnd()));
}

const measure = (t, size) => [...t]
  .reduce((a, c) => a + (/[⺀-｠]/.test(c) ? size : size * 0.55), 0);

/** 每一段文字的框，旋轉的另外算。 */
function boxesOf(svg) {
  return [...svg.matchAll(
    /<text x="([-\d.]+)" y="([-\d.]+)"[^>]*font-size="([\d.]+)"[^>]*text-anchor="(\w+)"([^>]*)>([^<]*)</g)]
    .filter((m) => m[6].trim())
    .map((m) => {
      const x = Number(m[1]); const y = Number(m[2]); const size = Number(m[3]);
      const w = measure(m[6], size);
      if (/rotate\(-?9[05]/.test(m[5])) {
        return { txt: m[6], size, l: x - size, r: x + size, y: y - w / 2, tall: w };
      }
      const l = m[4] === 'start' ? x : (m[4] === 'end' ? x - w : x - w / 2);
      return { txt: m[6], size, l, r: l + w, y, tall: 0 };
    });
}

function clean(name, svg, W, H) {
  ok(`${name}：SVG 乾淨`, !svgProblem(svg), svgProblem(svg));
  const out = boxesOf(svg).filter((b) => b.l < -1 || b.r > W + 1
    || b.y < -1 || b.y + b.tall > H + 1);
  ok(`${name}：沒有字跑出畫布`, !out.length,
    out.map((b) => `「${b.txt}」${b.l.toFixed(0)}~${b.r.toFixed(0)}`).join('、'));
}

// ── 這六種都真的註冊了 ──────────────────────────────────────────────────────
{
  const want = ['histogram', 'heatmap', 'violin', 'forest', 'survival', 'venn'];
  ok('六種學術圖都在支援清單裡',
    want.every((t) => CHART_TYPES.includes(t)),
    want.filter((t) => !CHART_TYPES.includes(t)).join(', '));
}

// ── 直方圖 ──────────────────────────────────────────────────────────────────
{
  const v = sample(200, 1.08, 0.035);
  const svg = chartSvg({ type: 'histogram', title: '厚度分布', values: v,
    normal: true, x_axis: '厚度（mm）', width: 450 }, TH);
  const H = Number(/height="(\d+)"/.exec(svg)[1]);
  clean('histogram', svg, 450, H);

  // 長條要互相貼著——中間有縫的是長條圖，不是直方圖。
  const bars = [...svg.matchAll(/<rect x="([\d.]+)" y="[\d.]+" width="([\d.]+)"/g)]
    .map((m) => ({ x: Number(m[1]), w: Number(m[2]) }))
    .filter((b) => b.w > 2 && b.w < 80)
    .sort((a, b) => a.x - b.x);
  const gaps = bars.slice(1).map((b, i) => b.x - (bars[i].x + bars[i].w));
  ok('直方圖的長條是連著的', gaps.every((g) => g < 2), `最大縫隙 ${Math.max(...gaps).toFixed(1)}`);

  // n、平均、標準差要在圖上——沒有這三個數字的直方圖沒辦法跟別張比。
  ok('標了 n、平均、標準差', /n = 200/.test(svg) && /平均/.test(svg) && /標準差/.test(svg));

  // 常態曲線要畫在「次數」的尺度上。密度曲線直接畫會變成貼著零的一條線。
  const curve = /<path d="(M[^"]+)" fill="none" stroke="#C0392B"/.exec(svg);
  ok('常態曲線有畫出來', !!curve);
  if (curve) {
    const ys = [...curve[1].matchAll(/[ML][\d.]+,([\d.]+)/g)].map((m) => Number(m[1]));
    ok('常態曲線在圖裡面，不是壓在軸上',
      Math.min(...ys) > 20 && Math.max(...ys) < H,
      `y 範圍 ${Math.min(...ys).toFixed(0)}~${Math.max(...ys).toFixed(0)}`);
  }

  let threw = '';
  try { chartSvg({ type: 'histogram', values: [3] }, TH); } catch (e) { threw = e.message; }
  ok('只有一個數值會說清楚要原始資料', /原始數值/.test(threw), threw);
}

// ── 熱力圖 ──────────────────────────────────────────────────────────────────
{
  const corr = [[1, 0.72, -0.31], [0.72, 1, -0.12], [-0.31, -0.12, 1]];
  const svg = chartSvg({ type: 'heatmap', title: '相關矩陣', matrix: corr,
    rows: ['爐溫', '壓力', '良率'], cols: ['爐溫', '壓力', '良率'], width: 450 }, TH);
  clean('heatmap', svg, 450, Number(/height="(\d+)"/.exec(svg)[1]));
  ok('列名沒有被截斷', !/…/.test(svg), (/>([^<]*…)</.exec(svg) || [])[1]);
  ok('格子裡有數字', /0\.72/.test(svg) && /-0\.31/.test(svg));

  // 跨越零就要用雙向配色：相關係數的正負是重點，同一個方向的深淺會把它抹掉。
  const fills = [...svg.matchAll(/<rect[^>]*fill="(#[0-9A-Fa-f]{6})"/g)].map((m) => m[1]);
  ok('正負用不同方向的顏色', new Set(fills).size >= 4, [...new Set(fills)].join(' '));

  let threw = '';
  try { chartSvg({ type: 'heatmap', matrix: [[1, 2], [3]] }, TH); } catch (e) { threw = e.message; }
  ok('每列長度不一樣會擋下來', /一樣/.test(threw), threw);
}

// ── 小提琴圖 ────────────────────────────────────────────────────────────────
{
  // 第二組故意做成雙峰：箱型圖看不出來，小提琴圖看得出來，這就是它存在的理由。
  const one = sample(80, 10, 1, 3);
  const two = [...sample(40, 8, 0.6, 11), ...sample(40, 12, 0.6, 29)];
  const svg = chartSvg({ type: 'violin', title: '兩組', width: 450,
    groups: [{ label: '單峰', values: one }, { label: '雙峰', values: two }] }, TH);
  clean('violin', svg, 450, Number(/height="(\d+)"/.exec(svg)[1]));
  ok('每組標了 n', (svg.match(/n = \d+/g) || []).length === 2);

  // 雙峰那一側的輪廓應該有兩個凸起：把路徑的寬度沿著 y 掃一遍，看有幾個局部極大。
  const paths = [...svg.matchAll(/<path d="M([^"]+)" fill="#[0-9A-Fa-f]{6}" stroke=/g)];
  ok('兩把小提琴都畫出來了', paths.length === 2, `${paths.length} 個`);
  if (paths.length === 2) {
    const pts = paths[1][1].split('L').map((q) => q.split(',').map(Number));
    const half = pts.slice(0, Math.floor(pts.length / 2));
    const xs = half.map((q) => q[0]);
    let peaks = 0;
    for (let i = 2; i < xs.length - 2; i++) {
      if (xs[i] > xs[i - 2] && xs[i] > xs[i + 2]) peaks += 1;
    }
    ok('雙峰的資料畫出兩個凸起', peaks >= 2, `${peaks} 個凸起`);
  }
}

// ── 森林圖 ──────────────────────────────────────────────────────────────────
{
  const svg = chartSvg({ type: 'forest', title: '各廠', log_x: true, x_axis: '勝算比',
    rows: [
      { label: '高雄一廠', value: 1.42, lo: 1.08, hi: 1.87, weight: 28 },
      { label: '無錫廠', value: 0.94, lo: 0.71, hi: 1.24, weight: 18 },
    ],
    summary: { label: '合計（隨機效果）', value: 1.31, lo: 1.09, hi: 1.58 },
    width: 450 }, TH);
  const H = Number(/height="(\d+)"/.exec(svg)[1]);
  clean('forest', svg, 450, H);
  ok('長標籤沒有被截斷', !/…/.test(svg), (/>([^<]*…)</.exec(svg) || [])[1]);
  ok('高度跟著列數走，不是固定 380', H < 300, `${H}`);

  // 無效果線：比值型的效果量要畫在 1，不是 0。
  const dashed = /<line x1="([\d.]+)"[^>]*stroke-dasharray/.exec(svg);
  ok('有畫無效果線', !!dashed);
  ok('合計畫成菱形', /<path d="M[\d.]+,[\d.]+L[\d.]+,[\d.]+L[\d.]+,[\d.]+L[\d.]+,[\d.]+Z"/.test(svg));
  ok('右欄有效果量和信賴區間', /1\.42 \(1\.08–1\.87\)/.test(svg));

  let threw = '';
  try {
    chartSvg({ type: 'forest', log_x: true, rows: [{ label: 'a', value: -1, lo: -2, hi: 0 }] }, TH);
  } catch (e) { threw = e.message; }
  ok('log_x 碰到非正數會說清楚', /大於 0/.test(threw), threw);
}

// ── 存活曲線 ────────────────────────────────────────────────────────────────
{
  const svg = chartSvg({ type: 'survival', title: '壽命', x_axis: '小時', width: 450,
    curves: [
      { name: '新製程', points: [[0, 1], [200, 0.98], [700, 0.88], [1400, 0.79]], censored: [520] },
      { name: '舊製程', points: [[0, 1], [350, 0.86], [900, 0.63], [1400, 0.52]], dashed: true },
    ] }, TH);
  clean('survival', svg, 450, Number(/height="(\d+)"/.exec(svg)[1]));

  /**
   * 一定要是階梯。
   *
   * 事件之間什麼都沒發生，用直線連起來等於宣稱有穩定的下降——在真實的研究裡，
   * 那是「一個月內掉了一半」和「一年內陸續發生」的差別。
   */
  const d = /<path d="(M[^"]+)" fill="none"/.exec(svg)[1];
  ok('畫成階梯（只有水平和垂直）', /^M[\d.]+,[\d.]+(?:H[\d.]+V[\d.]+)+$/.test(d), d.slice(0, 60));
  /**
   * 設限標記要真的多畫出東西來。
   *
   * 數總線條數會把格線也算進去，那條斷言就算把 censored 整個拿掉也會過。改成
   * 直接比對「有設限」和「沒設限」兩張圖差幾條線——差幾條就該等於幾個設限點。
   */
  const bare = chartSvg({ type: 'survival', width: 450,
    curves: [
      { name: '新製程', points: [[0, 1], [200, 0.98], [700, 0.88], [1400, 0.79]] },
      { name: '舊製程', points: [[0, 1], [350, 0.86], [900, 0.63], [1400, 0.52]], dashed: true },
    ] }, TH);
  const count = (x) => (x.match(/<line /g) || []).length;
  ok('設限有標記', count(svg) - count(bare) === 1,
    `有設限 ${count(svg)} 條、沒設限 ${count(bare)} 條`);
  ok('兩條曲線都在', /新製程/.test(svg) && /舊製程/.test(svg));
}

// ── 范氏圖 ──────────────────────────────────────────────────────────────────
{
  const svg = chartSvg({ type: 'venn', title: '三種檢測', width: 450,
    sets: ['AOI', 'X-Ray', '電測'],
    overlaps: { A: 142, B: 88, C: 61, AB: 37, AC: 24, BC: 19, ABC: 11 } }, TH);
  clean('venn', svg, 450, Number(/height="(\d+)"/.exec(svg)[1]));
  ok('三個集合的名字都在圖上',
    ['AOI', 'X-Ray', '電測'].every((n) => svg.includes(`>${n}<`)),
    ['AOI', 'X-Ray', '電測'].filter((n) => !svg.includes(`>${n}<`)).join(', '));
  ok('七個區的數字都在', [142, 88, 61, 37, 24, 19, 11].every((n) => svg.includes(`>${n}<`)));
  ok('三個圓', (svg.match(/<circle /g) || []).length === 3);

  let threw = '';
  try { chartSvg({ type: 'venn', sets: ['只有一個'] }, TH); } catch (e) { threw = e.message; }
  ok('少於兩個集合會說清楚', /兩到三個/.test(threw), threw);
}

// ── 誤差棒 ──────────────────────────────────────────────────────────────────
{
  const svg = chartSvg({ type: 'column', title: '各組平均', width: 450,
    categories: ['對照', 'A', 'B'], values: [12.4, 15.8, 14.1],
    error: [1.2, 1.9, 1.5], error_label: '± 1 SD' }, TH);
  clean('errorbar', svg, 450, 380);
  ok('誤差棒的說明有標出來', /± 1 SD/.test(svg));

  /**
   * 軸要先把誤差棒算進去。
   *
   * 最高的一根是 15.8 ± 1.9 = 17.7，所以軸不能停在 16——區間被畫到圖外，比
   * 不畫還糟，因為它看起來像是「上限就到這裡」。
   */
  const ys = [...svg.matchAll(/<line x1="[\d.]+" y1="([\d.]+)" x2="[\d.]+" y2="([\d.]+)"/g)]
    .flatMap((m) => [Number(m[1]), Number(m[2])]);
  ok('誤差棒沒有被軸切掉', Math.min(...ys) > 30, `最高點 y=${Math.min(...ys).toFixed(0)}`);

  // 形狀對不上就不要硬畫——畫在錯的長條上比不畫更糟。
  const bad = chartSvg({ type: 'column', width: 450, categories: ['a', 'b'],
    values: [1, 2], error: [1, 2, 3] }, TH);
  ok('長度對不上就忽略，不亂畫', !svgProblem(bad));

  const line2 = chartSvg({ type: 'line', width: 450, categories: ['1月', '2月'],
    values: [[97.2, 96.8]], series: ['A'], error: [[0.4, 0.6]], error_label: '± 1 SE' }, TH);
  ok('折線圖也畫得出誤差棒', /± 1 SE/.test(line2) && !svgProblem(line2));
}

// ── 迴歸線與氣泡 ────────────────────────────────────────────────────────────
{
  // y = 0.5x + 2 完全線性 → R² 應該是 1。
  const exact = Array.from({ length: 8 }, (_, i) => ({ x: i + 1, y: 0.5 * (i + 1) + 2 }));
  const svg = chartSvg({ type: 'scatter', width: 450, points: exact, fit: 'linear' }, TH);
  const r2 = /R² = ([\d.]+)/.exec(svg);
  ok('完全線性的資料 R² = 1.00', r2 && r2[1] === '1.00', r2 && r2[1]);
  const eq = /y = ([\d.]+)x \+ ([\d.]+)/.exec(svg);
  ok('斜率和截距算對了', eq && eq[1] === '0.5' && eq[2] === '2', eq && eq.slice(1).join(' / '));

  const band = chartSvg({ type: 'scatter', width: 450, fit: 'linear', fit_band: true,
    points: [{ x: 1, y: 2 }, { x: 2, y: 2.4 }, { x: 3, y: 3.5 }, { x: 4, y: 3.9 },
      { x: 5, y: 5.2 }] }, TH);
  ok('信賴帶畫得出來', /<polygon points="[^"]+" fill="#1F3864" opacity="0.12"/.test(band));

  /**
   * 氣泡按面積，不是按半徑。
   *
   * 四倍的值要畫成兩倍的半徑。直接把值當半徑的話墨水會變十六倍，而那正是氣泡圖
   * 最常見的那種謊。
   */
  const bub = chartSvg({ type: 'scatter', width: 450,
    points: [{ x: 1, y: 1, size: 25 }, { x: 2, y: 2, size: 100 }] }, TH);
  const rs = [...bub.matchAll(/<circle[^>]*r="([\d.]+)"/g)].map((m) => Number(m[1]));
  ok('有兩顆大小不同的氣泡', rs.length === 2 && rs[0] !== rs[1], rs.join(' '));
  if (rs.length === 2) {
    // r = 4 + 14·√(v/max)：25 → 4+7 = 11，100 → 4+14 = 18。
    ok('大小按面積換算', Math.abs(rs[0] - 11) < 0.5 && Math.abs(rs[1] - 18) < 0.5, rs.join(' '));
  }
}

// ── Word 真的開得了這些圖 ───────────────────────────────────────────────────
{
  const word = fs.readFileSync(new URL('../src/taskpane/hosts/word.js', import.meta.url), 'utf8');
  const types = ['histogram', 'heatmap', 'violin', 'forest', 'survival', 'venn'];
  // 這幾種不吃 categories/values，要在 SELF_SHAPED 裡，不然會被那條檢查擋掉。
  const self = /const SELF_SHAPED = \[([\s\S]*?)\];/.exec(word);
  ok('word.js 知道這些圖自己帶資料',
    self && types.every((t) => self[1].includes(`'${t}'`)),
    types.filter((t) => !self || !self[1].includes(`'${t}'`)).join(', '));
  // 插完的回報訊息要講得出中文名字，不然會說「已插入 histogram 圖」。
  const labels = /const CHART_LABEL = \{([\s\S]*?)\};/.exec(word);
  ok('每一種都有中文名字',
    labels && types.every((t) => new RegExp(`${t}:`).test(labels[1])),
    types.filter((t) => !labels || !new RegExp(`${t}:`).test(labels[1])).join(', '));
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
