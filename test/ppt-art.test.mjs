/**
 * The illustration library — ppt_find_art, and the SVG half of ppt_add_photo.
 *
 * ── Why illustrations got built at all ──
 * A stock photograph has to be ABOUT something. On a slide about warehouse
 * throughput a warehouse photo works; on a slide about 「流程優化」 the same photo
 * is decoration pretending to be an argument, and everybody can tell. That is
 * what makes a deck look like a template. A flat illustration makes no such
 * claim — it fills the space, carries a mood, and nobody mistakes it for
 * evidence — so it is the right answer for exactly the slides where a photo
 * would be filler.
 *
 * ── The two things that break silently ──
 *   1. An SVG sent through Office.CoercionType.Image produces a broken-picture
 *      placeholder. Same call, same arguments, nothing thrown. So the coercion
 *      type has to be pinned, not assumed.
 *   2. A scrim over a transparent illustration is a grey rectangle with a
 *      drawing faintly inside it. White type over a pale illustration is
 *      invisible. Both look like a rendering bug and neither raises anything.
 *
 * The catalogue itself lives in the server repo, so the search half skips when
 * that is not checked out beside this one.
 */
import fs from 'fs';

let pass = 0, fail = 0;
const ok = (name, cond, detail) => {
  cond ? pass++ : fail++;
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}${cond ? '' : '  → ' + (detail || '')}`);
};

// ── The pane: an SVG must go in as a vector ──────────────────────────────────

let LOG = [];
let slide;
function fakeShape(kind, opts = {}) {
  const s = {
    id: `a${LOG.length}`, kind, type: 'GeometricShape', name: '', deleted: false,
    left: opts.left, top: opts.top, width: opts.width, height: opts.height,
    _text: opts.text || '', _font: {}, _fill: null, _alpha: undefined,
    fill: {
      setSolidColor(c) { s._fill = c; },
      clear() { s._fill = null; },
      set transparency(v) { s._alpha = v; },
      get transparency() { return s._alpha; },
    },
    lineFormat: { visible: undefined, color: undefined, weight: undefined },
    load() {}, delete() { s.deleted = true; },
  };
  Object.defineProperty(s, 'textFrame', {
    get: () => ({
      get hasText() { return !!s._text; },
      textRange: {
        get text() { return s._text; }, set text(v) { s._text = v; },
        font: s._font, paragraphFormat: { horizontalAlignment: undefined },
      },
      get verticalAlignment() { return s._valign; },
      set verticalAlignment(v) { s._valign = v; },
      wordWrap: undefined, autoSizeSetting: undefined,
      leftMargin: 0, rightMargin: 0, topMargin: 0, bottomMargin: 0,
      load() {},
    }),
  });
  return s;
}
function fakeSlide() {
  const items = [];
  return {
    id: 'slide-1', _all: items,
    shapes: {
      get items() { return items.filter((x) => !x.deleted); },
      load() {},
      addGeometricShape: (g, o) => { const x = fakeShape('geo', o); x.geo = g; items.push(x); return x; },
      addTextBox: (t, o) => { const x = fakeShape('text', { ...o, text: t }); items.push(x); return x; },
      addLine: (c, o) => { const x = fakeShape('line', o); items.push(x); return x; },
      getItemAt: (i) => items.filter((x) => !x.deleted)[i],
    },
  };
}
slide = fakeSlide();

const ctx = {
  presentation: {
    slides: { get items() { return [slide]; }, load() {}, add() {}, getItemAt: () => slide },
    setSelectedSlides() {},
  },
  sync: async () => {}, load: () => {},
};

let RESPONSE = null;
let FETCHED = [];
globalThis.fetch = async (url) => { FETCHED.push(String(url)); return RESPONSE; };
globalThis.btoa = (s) => Buffer.from(s, 'binary').toString('base64');
globalThis.PowerPoint = { run: async (fn) => fn(ctx) };

let COERCION = '1.2';
globalThis.Office = {
  context: {
    requirements: {
      isSetSupported: (name, v) => (name === 'ImageCoercion'
        ? Number(v) <= Number(COERCION) : true),
    },
    document: {
      url: 'C:\\d.pptx',
      addHandlerAsync: () => {},
      setSelectedDataAsync: (data, opts, cb) => {
        LOG.push({ what: 'image', coercion: opts.coercionType, data, box: {
          left: opts.imageLeft, top: opts.imageTop,
          width: opts.imageWidth, height: opts.imageHeight,
        } });
        cb({ status: 'ok' });
      },
    },
  },
  EventType: { DocumentSelectionChanged: 'sel' },
  CoercionType: { Image: 'image', XmlSvg: 'xmlsvg' },
  AsyncResultStatus: { Succeeded: 'ok' },
};
globalThis.window = globalThis;
globalThis.document = {
  createElement: () => ({ getContext: () => null, style: {} }),
  getElementById: () => null, head: { appendChild: () => {} },
};

const host = await import('../src/taskpane/hosts/powerpoint.js');

const SVG_URL = 'https://stories.freepiklabs.com/storage/12245/Forest-01.svg';
const SVG_BODY = '<svg xmlns="http://www.w3.org/2000/svg"><circle r="4"/></svg>';
const svgResponse = () => ({
  ok: true, status: 200,
  headers: { get: (k) => (k.toLowerCase() === 'content-type' ? 'image/svg+xml' : null) },
  text: async () => SVG_BODY,
  arrayBuffer: async () => new ArrayBuffer(8),
});
async function place(args, response = svgResponse()) {
  slide = fakeSlide(); LOG = []; FETCHED = []; RESPONSE = response;
  const r = await host.executeTool('ppt_add_photo', args);
  return { r, log: LOG, url: FETCHED[0], shapes: slide.shapes.items };
}

ok('ppt_find_art 有被宣告出去', host.supportedTools().includes('ppt_find_art'),
  '伺服器自己回答這個工具，但工具清單是從增益集報上去的——沒宣告就等於不存在');
ok('而且有進度字串', !!host.meta.toolLabels().ppt_find_art);

{
  const { r, log } = await place({ url: SVG_URL, mode: 'block' });
  ok('插畫放得進去', r.ok, r.content);
  // base64 走 Image 通道會變成一個破圖佔位符，而且什麼都不會丟出來。
  ok('SVG 走向量通道，不是圖片通道', log[0].coercion === 'xmlsvg', String(log[0].coercion));
  ok('送進去的是 SVG 原文，不是 base64', log[0].data === SVG_BODY);
}
{
  // Freepik 的網址不是我們的，而且向量圖也沒有裁切這回事。
  const { url } = await place({ url: SVG_URL, mode: 'side' });
  ok('不會去改插畫網址的查詢字串', url === SVG_URL, url);
}
{
  const { shapes } = await place({ url: SVG_URL, mode: 'hero', title: '標題' });
  // 透明底的向量圖蓋上黑色半透明，結果是一塊灰色方塊裡面隱約有張圖。
  ok('插畫不會被蓋上遮罩', !shapes.some((s) => s._fill === '#000000'),
    shapes.map((s) => s._fill).join(' '));
  const t = shapes.find((s) => s._text === '標題');
  ok('而且標題不是白字', t._font.color !== '#FFFFFF', String(t._font.color));
  const lum = (c) => { const x = String(c).replace('#', '');
    return (parseInt(x.slice(0, 2), 16) + parseInt(x.slice(2, 4), 16) + parseInt(x.slice(4, 6), 16)) / 765; };
  ok('是深色的，在淺色插畫上讀得到', lum(t._font.color) < 0.5, String(t._font.color));
}
{
  // 照片還是要有遮罩跟白字——這條是確認上面那個分支沒有把照片也一起改掉。
  const { shapes } = await place({ url: 'https://images.unsplash.com/photo-1', mode: 'hero', title: '標題' }, {
    ok: true, status: 200,
    headers: { get: () => 'image/jpeg' },
    arrayBuffer: async () => new ArrayBuffer(64),
    text: async () => '',
  });
  ok('照片照樣有遮罩', shapes.some((s) => s._fill === '#000000'));
  ok('照片照樣是白字', shapes.find((s) => s._text === '標題')._font.color === '#FFFFFF');
}
{
  COERCION = '1.1';
  const { r } = await place({ url: SVG_URL, mode: 'block' });
  COERCION = '1.2';
  ok('舊版 Office 放不了向量圖時會講清楚原因',
    !r.ok && /1\.2/.test(r.content), r.content);
}

// ── The catalogue ────────────────────────────────────────────────────────────

const serverDir = new URL('../../ai-agents-office/server/src/', import.meta.url);
if (!fs.existsSync(serverDir)) {
  console.log('      （找不到 ai-agents-office，跳過插畫庫的跨 repo 檢查）');
} else {
  const src = fs.readFileSync(new URL('services/artLibrary.ts', serverDir), 'utf8');
  ok('插畫庫是讀檔案，不是把 417 個網址塞進提示詞',
    src.includes('storyset-library.json') && !src.includes('photo-1'));
  ok('找不到就回報分類清單，而不是硬給一個結果',
    src.includes('插畫庫裡沒有') && src.includes('一張不相干的插畫比留白更糟'));
  ok('讀不到檔案時會說清楚，並給替代做法', src.includes('找不到 storyset-library.json'));
  // 模糊比對什麼都找得到，等於模型分不出「命中」跟「聳肩」，然後就把聳肩放上去了。
  ok('沒有命中就回空的，不做模糊比對', src.includes('score > 0'));
  ok('分類命中的權重高於標籤', /cat === term\) score \+= 10/.test(src));

  // Node 22 會直接剝掉型別執行 .ts；能載入就真的跑一次搜尋。
  const lib = await import(new URL('services/artLibrary.ts', serverDir).href)
    .catch(() => null);
  if (lib) {
    const hits = lib.searchArt('teamwork', 5);
    ok('搜尋真的找得到東西', hits.length > 0 && hits.every((h) => /^https:/.test(h.url)),
      JSON.stringify(hits.slice(0, 1)));
    ok('回的是相關分類，不是隨便前幾筆',
      hits.every((h) => /team/.test(h.category) || /team/i.test(h.label)),
      hits.map((h) => h.category).join(' '));
    ok('沒有命中就回空陣列', lib.searchArt('zzzzz-not-a-topic', 5).length === 0);
    const miss = lib.describeArtSearch('zzzzz-not-a-topic', 5);
    ok('沒命中的時候把分類清單交出去',
      miss.content.includes('teamwork') && miss.content.includes('沒有'), miss.content.slice(0, 80));
    ok('而且叫它不要硬塞', /比留白更糟/.test(miss.content));
    ok('limit 有效', lib.searchArt('a', 3).length <= 3);
    ok('分類數量符合預期', lib.artCategories().length >= 40,
      String(lib.artCategories().length));
  }
  const cat = new URL('../../ai-agents-office/client/public/storyset-library.json', import.meta.url);
  if (fs.existsSync(cat)) {
    const json = JSON.parse(fs.readFileSync(cat, 'utf8'));
    const cats = Object.keys(json.categories || {});
    const n = cats.reduce((a, k) => a + json.categories[k].length, 0);
    ok('圖庫確實有東西', cats.length >= 40 && n >= 400, `${cats.length} 類 / ${n} 張`);
    ok('每一筆都有網址', cats.every((k) => json.categories[k].every((x) => x && x.url)));
  }

  const bridge = fs.readFileSync(new URL('services/excelBridge.ts', serverDir), 'utf8');
  // 走到增益集再繞回來問我們自己，是多一條會壞掉的路。
  ok('ppt_find_art 由伺服器直接回答，不繞到增益集',
    /tool === 'ppt_find_art'/.test(bridge) && bridge.includes('describeArtSearch'));
  const spec = fs.readFileSync(new URL('services/pptToolSpec.ts', serverDir), 'utf8');
  ok('ppt_find_art 不是破壞性操作',
    !/name: 'ppt_find_art',[\s\S]{0,900}?destructive: true/.test(spec));
  ok('工具說明講清楚插畫跟照片的分工',
    /ppt_find_art[\s\S]{0,1200}?假裝成論點/.test(spec));
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
