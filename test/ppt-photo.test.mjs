/**
 * ppt_add_photo — the tool that finally puts a full-bleed image on a slide.
 *
 * ── Why this gets its own file and its own fakes ──
 * It is the only PowerPoint tool that touches the network, and the only one
 * whose output depends on two APIs cooperating in the right ORDER:
 * setSelectedDataAsync appends the picture, and everything drawn afterwards
 * lands on top of it. Get that backwards and the scrim ends up under the
 * photograph — which looks fine in code review and wrong on a projector.
 *
 * The three things pinned here are the three that are silently wrong when they
 * break:
 *
 *   1. ORDER — image first, scrim and type after.
 *   2. CROP — a 16:9 photo stretched into a 384×540 column distorts faces.
 *      setSelectedDataAsync cannot crop, so the crop has to be asked for in the
 *      URL, and only Unsplash is allowed to have its query string rewritten.
 *   3. THE SCRIM FALLBACK — fill.transparency is not in every build. Without the
 *      fallback, an old Office gets an opaque black rectangle covering the
 *      photo, which is worse than having no scrim at all.
 */

let SEQ = 0;
function fakeShape(kind, opts = {}) {
  const s = {
    id: `s${++SEQ}`, kind, type: 'GeometricShape', name: '', deleted: false,
    left: opts.left, top: opts.top, width: opts.width, height: opts.height,
    _text: opts.text || '', _font: {}, _fill: null, _alpha: undefined,
    fill: {
      setSolidColor(c) { s._fill = c; },
      clear() { s._fill = null; },
      set transparency(v) {
        if (NO_TRANSPARENCY) throw new Error('NotSupported');
        s._alpha = v;
      },
      get transparency() { return s._alpha; },
    },
    lineFormat: { visible: undefined, color: undefined, weight: undefined },
    load() {},
    delete() { s.deleted = true; },
  };
  Object.defineProperty(s, 'textFrame', {
    get: () => ({
      get hasText() { return !!s._text; },
      textRange: {
        get text() { return s._text; },
        set text(v) { s._text = v; },
        font: s._font,
        paragraphFormat: { horizontalAlignment: undefined },
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

let NO_TRANSPARENCY = false;
/** Everything that happened, in order — the ordering assertions live off this. */
let LOG = [];
let slide;

function fakeSlide() {
  const items = [];
  return {
    id: 'slide-1',
    _all: items,
    shapes: {
      get items() { return items.filter((s) => !s.deleted); },
      load() {},
      addGeometricShape: (g, o) => {
        const s = fakeShape('geo', o); s.geo = g; items.push(s);
        LOG.push({ what: 'shape', geo: g, box: o });
        return s;
      },
      addTextBox: (t, o) => {
        const s = fakeShape('text', { ...o, text: t }); s.type = 'TextBox'; items.push(s);
        LOG.push({ what: 'text', text: t, box: o });
        return s;
      },
      addLine: (c, o) => { const s = fakeShape('line', o); items.push(s); return s; },
      getItemAt: (i) => items.filter((s) => !s.deleted)[i],
    },
  };
}
slide = fakeSlide();

const ctx = {
  presentation: {
    slides: {
      get items() { return [slide]; },
      load() {},
      add() {},
      getItemAt: () => slide,
    },
    setSelectedSlides() {},
  },
  sync: async () => {},
  load: () => {},
};

/** What the last fetch was asked for — the crop assertions read this. */
let FETCHED = [];
let RESPONSE = null;
globalThis.fetch = async (url) => {
  FETCHED.push(String(url));
  if (RESPONSE instanceof Error) throw RESPONSE;
  return RESPONSE;
};
const imageResponse = (bytes = 64) => ({
  ok: true,
  status: 200,
  headers: { get: (k) => (k.toLowerCase() === 'content-type' ? 'image/jpeg' : null) },
  arrayBuffer: async () => new ArrayBuffer(bytes),
});

globalThis.btoa = (s) => Buffer.from(s, 'binary').toString('base64');
globalThis.PowerPoint = { run: async (fn) => fn(ctx) };
globalThis.Office = {
  context: {
    requirements: { isSetSupported: () => true },
    document: {
      url: 'C:\\d.pptx',
      addHandlerAsync: () => {},
      setSelectedDataAsync: (data, opts, cb) => {
        LOG.push({ what: 'image', bytes: data.length, box: {
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
  getElementById: () => null,
  head: { appendChild: () => {} },
};

const host = await import('../src/taskpane/hosts/powerpoint.js');

let pass = 0, fail = 0;
const ok = (name, cond, detail) => {
  cond ? pass++ : fail++;
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}${cond ? '' : '  → ' + (detail || '')}`);
};
const UNSPLASH = 'https://images.unsplash.com/photo-1518770660439-4636190af475';
const SVG_ANY = 'https://stories.freepiklabs.com/storage/1/a.svg';
async function shoot(args, response = imageResponse()) {
  slide = fakeSlide(); LOG = []; FETCHED = []; RESPONSE = response;
  const r = await host.executeTool('ppt_add_photo', args);
  return { r, log: LOG, url: FETCHED[0], shapes: slide.shapes.items };
}

// ── The tool exists and is treated as a write ────────────────────────────────

ok('ppt_add_photo 在工具清單裡', host.supportedTools().includes('ppt_add_photo'));
ok('而且有進度字串', !!host.meta.toolLabels().ppt_add_photo);

// ── Order: the image goes down first, the scrim on top of it ─────────────────
//
// setSelectedDataAsync 是「附加到最上層」，所以遮罩必須晚於圖片。順序反了的話
// 程式碼看起來一模一樣，只有投影出來才會發現照片被一塊黑蓋住。
{
  const { log } = await shoot({ url: UNSPLASH, mode: 'hero', title: '2026 趨勢' });
  const iImage = log.findIndex((e) => e.what === 'image');
  const iScrim = log.findIndex((e) => e.what === 'shape');
  const iTitle = log.findIndex((e) => e.text === '2026 趨勢');
  ok('圖片先進去', iImage === 0, JSON.stringify(log.map((e) => e.what)));
  ok('遮罩壓在圖片上面', iScrim > iImage);
  ok('標題又壓在遮罩上面', iTitle > iScrim);
  ok('hero 是滿版 960×540',
    log[iImage].box.width === 960 && log[iImage].box.height === 540
      && log[iImage].box.left === 0 && log[iImage].box.top === 0,
    JSON.stringify(log[iImage].box));
  ok('白色大標', log[iTitle] && slide.shapes.items
    .find((s) => s._text === '2026 趨勢')._font.color === '#FFFFFF');
}

// ── The crop, asked for in the URL because the API cannot crop ───────────────

{
  const { url } = await shoot({ url: UNSPLASH, mode: 'side' });
  const u = new URL(url);
  ok('side 模式要直式裁切',
    u.searchParams.get('w') === '768' && u.searchParams.get('h') === '1080'
      && u.searchParams.get('fit') === 'crop',
    url);
}
{
  const { url } = await shoot({ url: UNSPLASH, mode: 'hero' });
  ok('hero 模式要 16:9 大圖', new URL(url).searchParams.get('w') === '1920', url);
}
{
  // 別人的網址我們不懂，動它的 query string 只會把圖弄壞。
  const other = 'https://cdn.example.com/a.jpg?token=abc';
  const { url } = await shoot({ url: other, mode: 'hero' });
  ok('非 Unsplash 的網址原封不動', url === other, url);
}
{
  // 「unsplash.com.evil.tw」不是 Unsplash。比對要認邊界，不能用 includes。
  const evil = 'https://images.unsplash.com.evil.tw/x.jpg';
  const { url } = await shoot({ url: evil, mode: 'hero' });
  ok('只認真正的 unsplash.com 網域', url === evil, url);
}

// ── side: the photo takes 40%, and says where the other 60% is ───────────────

{
  const { r, log } = await shoot({ url: UNSPLASH, mode: 'side', title: '標題' });
  const img = log.find((e) => e.what === 'image');
  ok('side 是滿高的 384 寬直欄',
    img.box.width === 384 && img.box.height === 540 && img.box.top === 0,
    JSON.stringify(img.box));
  ok('預設靠右', img.box.left === 960 - 384, String(img.box.left));
  const title = slide.shapes.items.find((s) => s._text === '標題');
  ok('文字放在照片的另一邊', title.left + 456 <= img.box.left,
    `文字 ${title.left}+456，照片從 ${img.box.left}`);
  ok('回覆告訴模型剩下的空間在哪', /60%/.test(r.content) && /456/.test(r.content),
    r.content);
}
{
  const { log } = await shoot({ url: UNSPLASH, mode: 'side', position: 'left', title: '標題' });
  const img = log.find((e) => e.what === 'image');
  ok('指定 left 就靠左', img.box.left === 0);
  ok('文字跟著換到右邊',
    slide.shapes.items.find((s) => s._text === '標題').left >= 384);
}

// ── The scrim, and what happens when the build cannot do transparency ────────

{
  const { shapes } = await shoot({ url: UNSPLASH, mode: 'hero', title: 'T', scrim: 0.6 });
  const veil = shapes.find((s) => s._fill === '#000000');
  ok('遮罩濃度可調', Math.abs(veil._alpha - 0.4) < 1e-9, String(veil._alpha));
}
{
  const { shapes } = await shoot({ url: UNSPLASH, mode: 'hero', title: 'T', scrim: 5 });
  // 夾到 0.9 之後 transparency 是 0.1——浮點數算出來是 0.0999…，所以留一點容差。
  ok('濃度會被夾在合理範圍內',
    shapes.find((s) => s._fill === '#000000')._alpha >= 0.09,
    String(shapes.find((s) => s._fill === '#000000')._alpha));
}
{
  NO_TRANSPARENCY = true;
  const { r, shapes } = await shoot({ url: UNSPLASH, mode: 'hero', title: 'T' });
  NO_TRANSPARENCY = false;
  // 沒有 transparency 就是一塊不透明的黑——那比沒有遮罩糟得多，寧可拿掉。
  ok('舊版 Office 做不出半透明時，遮罩會被移除而不是留一塊黑',
    !shapes.some((s) => s._fill === '#000000'),
    shapes.map((s) => s._fill).join(' '));
  ok('而且整個呼叫還是成功的', r.ok, r.content);
}

// ── Failures say which half broke ────────────────────────────────────────────

{
  const { r } = await shoot({ url: UNSPLASH }, new Error('Failed to fetch'));
  ok('抓不到圖的時候講清楚是 CORS 或外網',
    !r.ok && /CORS|外網/.test(r.content), r.content);
  ok('而且告訴使用者改用拖曳', /拖進對話框/.test(r.content), r.content);
}
{
  const { r } = await shoot({ url: UNSPLASH }, {
    ok: true, status: 200,
    headers: { get: () => 'text/html' },
    arrayBuffer: async () => new ArrayBuffer(8),
  });
  ok('回來的不是圖片會被擋下來', !r.ok && /不是圖片/.test(r.content), r.content);
}
{
  const { r } = await shoot({ url: UNSPLASH }, {
    ok: false, status: 404, headers: { get: () => null }, arrayBuffer: async () => new ArrayBuffer(0),
  });
  ok('404 會說網址失效', !r.ok && /404/.test(r.content), r.content);
}
{
  const { r } = await shoot({ url: UNSPLASH }, imageResponse(7 * 1024 * 1024));
  ok('太大的圖會被擋下來並教怎麼縮小',
    !r.ok && /太大/.test(r.content) && /w=1600/.test(r.content), r.content);
}
{
  const { r } = await shoot({ url: 'file:///C:/secret.png' });
  ok('只接受 http(s)', !r.ok && /http/.test(r.content), r.content);
  ok('而且根本沒發出請求', FETCHED.length === 0);
}
{
  const { r } = await shoot({ url: UNSPLASH, mode: 'collage' });
  ok('未知的 mode 會被擋掉並列出可用的值',
    !r.ok && /hero/.test(r.content) && /side/.test(r.content), r.content);
}

// ── 圖片不能被拉變形 ────────────────────────────────────────────────────────
//
// setSelectedDataAsync 只會拉伸，沒有「放得下就好」這種模式。照片被擠 10% 還是
// 照片；一張畫在 1:1 畫布上的扁平插畫塞進 384×540 的直欄，就明顯是一張被拉過的
// 圖——「這樣會很怪」是正確的反應。
const svgRatio = (w, h) => ({
  ok: true, status: 200,
  headers: { get: () => 'image/svg+xml' },
  text: async () => `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ${w} ${h}"><circle r="4"/></svg>`,
  arrayBuffer: async () => new ArrayBuffer(8),
});
{
  const { log } = await shoot({ url: SVG_ANY, mode: 'side' }, svgRatio(500, 500));
  const b = log.find((e) => e.what === 'image').box;
  ok('1:1 的插畫放進直欄仍然是 1:1',
    Math.abs(b.width / b.height - 1) < 0.01, JSON.stringify(b));
  ok('而且置中在那個欄位裡', Math.abs((b.left + b.width / 2) - (960 - 192)) < 2,
    JSON.stringify(b));
  ok('沒有超出欄位', b.width <= 384 && b.height <= 540, JSON.stringify(b));
}
{
  const { log } = await shoot({ url: SVG_ANY, mode: 'hero' }, svgRatio(400, 300));
  const b = log.find((e) => e.what === 'image').box;
  // 向量圖在任何模式都不拉伸：被拉過的插畫一眼就看得出來，照片不會。
  ok('插畫當滿版背景時也不拉伸',
    Math.abs(b.width / b.height - 4 / 3) < 0.01, JSON.stringify(b));
}
{
  const { log } = await shoot({ url: 'https://images.unsplash.com/photo-1', mode: 'hero' });
  const b = log.find((e) => e.what === 'image').box;
  ok('照片當滿版背景時照樣填滿', b.width === 960 && b.height === 540, JSON.stringify(b));
}
{
  // 讀不出尺寸就退回填滿——也就是原本的行為，對背景而言是對的。
  const { r } = await shoot({ url: SVG_ANY, mode: 'block' }, {
    ok: true, status: 200,
    headers: { get: () => 'image/svg+xml' },
    text: async () => '<svg xmlns="http://www.w3.org/2000/svg"><circle r="4"/></svg>',
    arrayBuffer: async () => new ArrayBuffer(8),
  });
  ok('讀不出比例也不會失敗', r.ok, r.content);
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
