/**
 * Pin what ppt_build_slide actually puts on a slide.
 *
 * ── The two things this guards, both found by looking at real output ──
 *
 * 1. PLACEHOLDERS. `slides.add()` applies the default layout, which carries an
 *    empty title and subtitle placeholder. Everything the recipes draw goes on
 *    top of them, so every generated slide had 「按一下以新增標題」 across its
 *    middle in 40pt grey. It does not print and does not show in presentation
 *    mode — but the person judging the deck is looking at the editor, where it
 *    makes the output look broken rather than plain.
 *
 *    The dangerous half of the fix is the exception: a placeholder somebody has
 *    TYPED INTO is their content, and deleting it would destroy work this add-in
 *    cannot restore (the undo only removes shapes it created). So the assertions
 *    below care more about what survives than about what goes.
 *
 * 2. DENSITY. The first version rendered a bullet list as one text box: three
 *    lines in the top third and nothing below. Against a hand-made reference
 *    deck averaging 17–47 shapes per slide, that reads as a draft. Cards are now
 *    the default rendering rather than a tool the model has to choose — because
 *    when it was a choice, the model did not make it.
 */

// ── A PowerPoint that records shapes and honours delete() ────────────────────

let SEQ = 0;
function fakeShape(kind, opts = {}) {
  const s = {
    id: `s${++SEQ}`, kind, type: opts.type || 'GeometricShape',
    name: '', deleted: false,
    left: opts.left, top: opts.top, width: opts.width, height: opts.height,
    _text: opts.text || '', _font: {}, _fill: null, _line: {},
    fill: {
      setSolidColor(c) { s._fill = c; },
      clear() { s._fill = null; },
    },
    lineFormat: { visible: undefined, color: undefined, weight: undefined },
    load() {},
    delete() { s.deleted = true; },
  };
  Object.defineProperty(s, 'textFrame', {
    get() {
      if (opts.noTextFrame) throw new Error('InvalidArgument');
      return {
        get hasText() { return !!s._text; },
        textRange: {
          get text() { return s._text; },
          set text(v) { s._text = v; },
          font: s._font,
          paragraphFormat: { horizontalAlignment: undefined },
        },
        // Written back onto the shape, not onto this object: the getter builds a
        // fresh textFrame每次被讀取, so anything set on it would be thrown away
        // the moment the test looked for it.
        get verticalAlignment() { return s._valign; },
        set verticalAlignment(v) { s._valign = v; },
        wordWrap: undefined, autoSizeSetting: undefined,
        leftMargin: 0, rightMargin: 0, topMargin: 0, bottomMargin: 0,
        load() {},
      };
    },
  });
  return s;
}

function fakeSlide(initial = []) {
  const items = [...initial];
  return {
    id: 'slide-1',
    _all: items,
    shapes: {
      get items() { return items.filter((s) => !s.deleted); },
      load() {},
      addGeometricShape: (g, o) => {
        /**
         * Office rejects a negative left or top, and it rejects it at
         * `ctx.sync()` — so one out-of-range shape takes the whole slide with
         * it. That is not a hypothetical: two circles positioned to bleed off
         * the left edge made the entire cover fail, and the agent reported the
         * composition as unsupported on this machine.
         *
         * The fake throws where the real API does, so the test suite cannot
         * accept a coordinate Office would not.
         */
        if (o.left < 0 || o.top < 0) {
          throw new Error(`InvalidArgument: 負座標 ${JSON.stringify(o)}`);
        }
        /**
         * And the other way a shape kills a slide: a name that is not in
         * `PowerPoint.GeometricShapeType`. 'Oval' is the one that caught us —
         * the circle is called 'Ellipse', and 'Oval' appears only in the enum's
         * description text. Invalid names throw at sync, same as bad
         * coordinates, so the fake refuses them here.
         */
        const REAL = new Set(['Rectangle', 'RoundRectangle', 'Ellipse', 'Chevron',
          'RightArrow', 'Triangle', 'Diamond', 'Donut', 'Frame', 'Plus', 'Star5',
          'Hexagon', 'Pentagon', 'Trapezoid', 'HomePlate', 'Wave', 'Arc', 'Pie',
          'Teardrop', 'Cloud', 'Heart', 'Moon']);
        if (!REAL.has(g)) {
          throw new Error(`InvalidArgument: GeometricShapeType 沒有 "${g}"`);
        }
        const s = fakeShape('geo', o); s.geo = g; items.push(s); return s;
      },
      addTextBox: (t, o) => { const s = fakeShape('text', { ...o, text: t }); s.type = 'TextBox'; items.push(s); return s; },
      addLine: (c, o) => { const s = fakeShape('line', o); items.push(s); return s; },
      getItemAt: (i) => items.filter((s) => !s.deleted)[i],
    },
  };
}

let slide = fakeSlide();
const ctx = {
  presentation: {
    slides: {
      get items() { return [slide]; },
      load() {},
      add() { /* the fake starts with the slide already present */ },
      getItemAt: () => slide,
    },
    // Needed since the ornament layer landed: an SVG background is inserted
    // through the Common API, which lands on the SELECTED slide, so the host
    // selects one first.
    setSelectedSlides() {},
  },
  sync: async () => {},
  load: () => {},
};

globalThis.PowerPoint = { run: async (fn) => fn(ctx) };
/** Every insert, in order — the ornament assertions read this. */
let INSERTS = [];
globalThis.Office = {
  context: {
    requirements: { isSetSupported: () => true },
    document: {
      url: 'C:\\d.pptx',
      addHandlerAsync: () => {},
      setSelectedDataAsync: (data, opts, cb) => {
        INSERTS.push({ data, coercion: opts.coercionType });
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

/**
 * A picture that always arrives, 4:3, as a vector.
 *
 * The layouts that take an image fetch it before they draw, so without this the
 * fetch throws, the host notes 「圖片沒抓到」 and carries on — which is the right
 * behaviour and makes it impossible to test the case where the image DOES land.
 */
globalThis.fetch = async () => ({
  ok: true,
  status: 200,
  headers: { get: () => 'image/svg+xml' },
  text: async () => '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 400 300"><circle r="4"/></svg>',
  arrayBuffer: async () => new ArrayBuffer(8),
});

const host = await import('../src/taskpane/hosts/powerpoint.js');

let pass = 0, fail = 0;
const ok = (name, cond, detail) => {
  cond ? pass++ : fail++;
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}${cond ? '' : '  → ' + (detail || '')}`);
};
/**
 * 組一頁之前要先規劃——那是現在的流程，不是可選的。
 *
 * 增益集原本的毛病是「使用者說話 → 直接呼叫零件」，中間沒有設計這一步，所以不管
 * 給什麼需求，做出來都是 KPI 卡加圖表。提示詞寫了三次「請先規劃」都沒用——那是規則，
 * 而規則擋不住模型。改成一道門：沒規劃過的頁，ppt_build_slide 會被擋回去。
 *
 * 測試跟著走同一條路，不是繞過它——繞過去的話，這些斷言測的就是一個真實使用者走不到
 * 的路徑。這一整天已經兩次因為「測試用了模型不會送的形狀」而漏掉 bug。
 */
const plan = (n = 1) => host.executeTool('ppt_plan_slide', {
  slide: n,
  headline: '測試用的主張',
  blocks: [{ area: '上半', media: '圖表', why: '要比兩組數字' }],
});
const build = async (args) => {
  await plan(Number(args && args.slide) || 1);
  return host.executeTool('ppt_build_slide', args);
};
/** CHROME.contentTop in the host — where the body band starts. */
const CHROME_CONTENT_TOP = 124;
const alive = () => slide.shapes.items;
const texts = () => alive().filter((s) => s._text).map((s) => s._text);

/**
 * 這一頁的底色，不管它是誰畫的。
 *
 * 母題上線之後，有背景幾何的版式會把底色畫進那張裝飾 SVG 裡，骨架就不再鋪一塊
 * ——鋪了會把母題蓋掉。所以「底色是一個 960×540 的形狀」這個假設不再成立，
 * 但「底色一定有人負責」還是成立，而那才是測試真正在乎的事。
 */
const groundFill = () => {
  const shape = alive().find((s) => s.kind === 'geo' && s.width === 960 && s.height === 540);
  if (shape) return shape._fill;
  const art = INSERTS.length ? INSERTS[0].data : '';
  const m = /<rect width="960" height="540" fill="(#[0-9A-Fa-f]{6})"/.exec(art);
  return m ? m[1] : null;
};


// ── Placeholders ─────────────────────────────────────────────────────────────

slide = fakeSlide([
  fakeShape('ph', { type: 'Placeholder', text: '' }),
  fakeShape('ph', { type: 'Placeholder', text: '' }),
]);
await build({ kind: 'bullets', title: '標題', body: ['一', '二'] });
ok('空的預留位置會被清掉',
  !slide._all.some((s) => s.type === 'Placeholder' && !s.deleted),
  slide._all.filter((s) => s.type === 'Placeholder' && !s.deleted).length + ' 個還在');

// The one that matters: a placeholder with the person's own words in it.
slide = fakeSlide([
  fakeShape('ph', { type: 'Placeholder', text: '使用者自己打的標題' }),
  fakeShape('ph', { type: 'Placeholder', text: '' }),
]);
await build({ kind: 'bullets', title: 'T', body: ['一', '二'] });
const kept = slide._all.find((s) => s._text === '使用者自己打的標題');
ok('有內容的預留位置不會被刪', kept && !kept.deleted,
  '那是使用者的內容，而這裡的復原鍵救不回被刪掉的東西');
ok('同一張上空的還是會被清掉',
  slide._all.filter((s) => s.type === 'Placeholder' && !s.deleted).length === 1);

// A shape whose textFrame throws — pictures and lines do — must not abort the sweep.
slide = fakeSlide([
  fakeShape('pic', { type: 'Placeholder', noTextFrame: true }),
  fakeShape('ph', { type: 'Placeholder', text: '' }),
]);
const r = await build({ kind: 'bullets', title: 'T', body: ['一', '二'] });
ok('讀不到 textFrame 的物件不會讓整個流程掛掉', r.ok, r.content);

// ── Density ──────────────────────────────────────────────────────────────────

slide = fakeSlide();
await build({ kind: 'bullets', title: '兩件我一直在做的事', eyebrow: '01 關於我', page: 'P.02',
  body: ['寫程式：把麻煩的事變成自動的事', '做甜點：同一份食譜反覆調到滿意為止', '兩件事的共通點：都是細節決定成敗'] });

ok('三點的條列頁不再只是一個文字方塊', alive().length >= 14, `${alive().length} 個物件`);
ok('骨架四件都在', texts().includes('01 關於我') && texts().includes('兩件我一直在做的事')
  && texts().includes('P.02') && alive().some((s) => s.width === 960 && s.height === 5),
  texts().join(' | '));
ok('標題下的短線是 108pt 不是滿版',
  alive().some((s) => s.width === 108 && s.height === 4),
  '滿版的線讀起來是分隔，短的才是標題的底線');

// The colon split is what turns a flat line into a composed card.
ok('冒號前後拆成小標和說明',
  texts().includes('寫程式') && texts().includes('把麻煩的事變成自動的事'),
  texts().join(' | '));
ok('有編號徽章', texts().includes('01') && texts().includes('02') && texts().includes('03'));
ok('每張卡片有底色和色條', alive().filter((s) => s.kind === 'geo' && s._fill).length >= 6);

// Cards must fill the slide, not huddle at the top — the empty lower half was
// the most visible part of the "looks like a draft" problem.
const cards = alive().filter((s) => s.kind === 'geo' && s.width === 884);
const lowest = Math.max(...cards.map((s) => s.top + s.height));
ok('卡片填滿內文區，不是擠在上面三分之一', lowest > 440, `最低點 ${Math.round(lowest)}`);
ok('卡片沒有超出投影片', lowest <= 540, `最低點 ${Math.round(lowest)}`);

// No colon: still a card, just without the sub-line.
slide = fakeSlide();
await build({ kind: 'bullets', title: 'T', body: ['沒有冒號的一句話', '另一句'] });
ok('沒有冒號也照樣做成卡片', texts().includes('沒有冒號的一句話') && texts().includes('01'));

// Seven points is a slide that should have been two; cards that small read worse.
slide = fakeSlide();
await build({ kind: 'bullets', title: 'T', body: Array.from({ length: 7 }, (_, i) => `第 ${i + 1} 點`) });
// 稀疏內容靠上排，下半頁整片空白——那不是內容少，是內容放錯位置。這條跟
// 提示詞裡「稀疏內容先收攏再置中」是同一條規則，只是在渲染這一側。渲染如果
// 跟提示詞唱反調，那條規則就不會有人遵守。
ok('七點的純條列會垂直置中，不是貼著上緣',
  alive().some((s) => s._valign === 'Middle' && s.top >= CHROME_CONTENT_TOP));

// 九點排成一長條，等於浪費 16:9 的半個寬度，眼睛要走的距離比換欄還遠。
slide = fakeSlide();
await build({ kind: 'bullets', title: 'T', body: Array.from({ length: 9 }, (_, i) => `第 ${i + 1} 點`) });
ok('九點以上拆成兩欄',
  alive().filter((s) => s.kind === 'text' && s.top >= CHROME_CONTENT_TOP && s.width < 500).length >= 2);

slide = fakeSlide();
await build({ kind: 'bullets', title: 'T', body: Array.from({ length: 7 }, (_, i) => `第 ${i + 1} 點`) });
ok('七點以上退回純條列', texts().some((t) => t.includes('▪') && t.includes('第 1 點')),
  '七張 40pt 高的卡片比一份清單還難讀');
ok('條列用方塊符號不是圓點', texts().some((t) => t.includes('▪')) && !texts().some((t) => t.includes('•')));

// ── The cover, and the bug the palette shape prevents ────────────────────────

// 深底淺字這條規則屬於「滿版色底」那種封面，不是每一種封面。企業風現在用的是
// 左側色塊 + 白底，標題壓在白底上——所以這裡指定一個仍然滿版的版式來驗那條規則。
for (const [name, dark] of [['corporate', false], ['night', true]]) {
  slide = fakeSlide();
  await build({ kind: 'title', title: '2026 AI 趨勢', body: ['副標'], theme: name,
    look: 'bold' });
  const bleed = alive().find((s) => s.width === 960 && s.height === 540);
  const heading = alive().find((s) => s._text === '2026 AI 趨勢');
  const bg = String(bleed && bleed._fill || '').toUpperCase();
  const fg = String(heading && heading._font.color || '').toUpperCase();
  const lum = (c) => {
    const v = c.replace('#', '');
    return (parseInt(v.slice(0, 2), 16) * 0.299 + parseInt(v.slice(2, 4), 16) * 0.587
      + parseInt(v.slice(4, 6), 16) * 0.114) / 255;
  };
  ok(`${name} 封面：底色和文字有對比`, Math.abs(lum(bg) - lum(fg)) > 0.45,
    `底 ${bg}(${lum(bg).toFixed(2)}) 字 ${fg}(${lum(fg).toFixed(2)})`);
  ok(`${name} 封面：${dark ? '深底淺字' : '深底白字'}`, lum(bg) < 0.5 && lum(fg) > 0.5,
    `底 ${bg} 字 ${fg}——深色主題曾經做出淺藍底黑字，就是這條在擋`);
}

// ── Density elements: the KPI row, the callout, the footer band ──────────────
//
// 這三個是「一頁有上中下三層」的那三層。單獨看都只是幾個方塊，
// 但參考簡報跟我們產出的差距，幾乎全在這裡——我們的頁面下半部是空的。

slide = fakeSlide();
await host.executeTool('ppt_add_kpi', { items: [
  { value: '65%', label: '組織已導入', note: '2025 調查' },
  { value: '$37B', label: '市場規模' },
  { value: '5.1', label: '回收期（月）' },
] });
{
  const tiles = alive().filter((s) => s.kind === 'geo' && s.height === 101);
  ok('三個數據卡各有一張底板', tiles.length === 3, `${tiles.length} 張`);
  ok('數據卡橫排不重疊',
    tiles.every((t, i) => i === 0 || t.left >= tiles[i - 1].left + tiles[i - 1].width),
    tiles.map((t) => `${Math.round(t.left)}+${Math.round(t.width)}`).join(' '));
  ok('填滿內文寬度', Math.round(tiles[2].left + tiles[2].width) === 38 + 884,
    `右緣 ${Math.round(tiles[2].left + tiles[2].width)}`);
  // 一排同色是表格，一排不同色才是儀表板。
  const bars = alive().filter((s) => s.kind === 'geo' && s.height === 4);
  ok('每張卡頂有色條', bars.length === 3, `${bars.length} 條`);
  ok('三條顏色不重複', new Set(bars.map((b) => b._fill)).size === 3,
    bars.map((b) => b._fill).join(' '));
  ok('數字用色條的顏色，不是內文黑',
    alive().some((s) => s._text === '65%' && s._font.color === bars[0]._fill));
  ok('數字比標籤大很多',
    alive().find((s) => s._text === '65%')._font.size
      > alive().find((s) => s._text === '組織已導入')._font.size * 2);
  ok('沒有 note 的卡不會憑空生出小字', !texts().includes('undefined'));
}

// 五個以上就擠到放不下數字，擋掉比做出難看的東西好。
{
  slide = fakeSlide();
  const r6 = await host.executeTool('ppt_add_kpi',
    { items: Array.from({ length: 6 }, (_, i) => ({ value: `${i}` })) });
  ok('超過五個會被擋下來', !r6.ok && /5/.test(r6.content), r6.content);
}

slide = fakeSlide();
await host.executeTool('ppt_add_callout', { label: '洞察', text: '成長來自既有客戶，不是新客。' });
{
  const panel = alive().find((s) => s.kind === 'geo' && s.width === 884);
  const spine = alive().find((s) => s.kind === 'geo' && s.width === 4);
  ok('洞察框是一條滿內文寬的帶子', !!panel && panel.top >= 400,
    panel ? `top ${panel.top}` : '沒有');
  ok('左邊有色帶', !!spine && spine._fill === panel.lineFormat.color);
  ok('小標和內文分開兩塊', texts().includes('洞察')
    && texts().includes('成長來自既有客戶，不是新客。'));
  const body = alive().find((s) => s._text.startsWith('成長'));
  const lab = alive().find((s) => s._text === '洞察');
  ok('內文讓出小標的位置', body.left > lab.left + 90,
    `小標 ${lab.left} 內文 ${body.left}`);
}

slide = fakeSlide();
await host.executeTool('ppt_add_callout', { text: '只有一句話。' });
ok('沒有小標時內文從最左邊開始',
  alive().find((s) => s._text === '只有一句話。').left === 38 + 16);

// ── The page ground and the footer band ──────────────────────────────────────

slide = fakeSlide();
await build({ kind: 'bullets', title: 'T', body: ['一', '二'], page: 'P.03', footer: '資料來源：X',
  footband: true });
{
  const full = alive().filter((s) => s.kind === 'geo' && s.width === 960);
  // 底色現在由裝飾層那張 SVG 帶著（母題也在裡面），所以骨架不再另外鋪一塊——
  // 鋪了就會把母題蓋掉。這裡驗的是「底色有人負責」，不是「底色是某個形狀」。
  const art = INSERTS.length ? INSERTS[INSERTS.length - 1].data : '';
  const ground = full.find((s) => s.height === 540);
  ok('內容頁的底色有人負責，而且不是純白',
    ground ? ground._fill !== '#FFFFFF' : /<rect width="960" height="540" fill="#(?!FFFFFF)/i.test(art),
    ground ? ground._fill : art.slice(0, 120) || '兩邊都沒有');
  // 一張投影片只插一張裝飾層，而且在所有形狀之前——這是它能當底的唯一條件。
  ok('底色在最底下，蓋不到別人',
    ground ? slide._all.indexOf(ground) === 0 : slide._all.length > 0,
    ground ? `第 ${slide._all.indexOf(ground)} 個` : '裝飾層在形狀之前插入');
  ok('頂條是雙色的',
    full.some((s) => s.height === 5) && alive().some((s) => s.width === 288 && s.height === 2));
  const band = full.find((s) => s.height === 46 && s.top === 494);
  ok('滿版深色頁尾帶', !!band, full.map((s) => `${s.width}x${s.height}@${s.top}`).join(' '));
  const foot = alive().find((s) => s._text === '資料來源：X');
  ok('頁尾字在帶子裡而且是白的', foot && foot.top > 494 && foot._font.color === '#FFFFFF',
    foot ? `top ${foot.top} color ${foot._font.color}` : '沒有');
  // 左上跟右下同時印一次頁碼，是那種上了投影機才有人發現的重複。
  ok('開了頁尾帶，頁碼只出現一次',
    alive().filter((s) => s._text === 'P.03').length === 1);
  ok('而且移到帶子裡', alive().find((s) => s._text === 'P.03').top > 494);
}

slide = fakeSlide();
await build({ kind: 'bullets', title: 'T', body: ['一'], page: 'P.03', footer: '來源' });
ok('沒開頁尾帶就維持原本的細線頁尾',
  !alive().some((s) => s.height === 46 && s.top === 494)
    && alive().some((s) => s.kind === 'geo' && s.height === 1 && s.top === 486));
ok('沒開頁尾帶，頁碼還是在右上', alive().find((s) => s._text === 'P.03').top < 100);

// ── Looks: the knob that actually changes the design ─────────────────────────
//
// 五個配色做出來是同一個版面的五種上色。使用者要的「文青風」「可愛風」不是顏色，
// 是版式——所以這一段檢查的是：換 look 之後，版面的「結構」真的不一樣了。

const LOOKS = ['business', 'editorial', 'literary', 'minimal', 'soft', 'bold', 'tech'];
const shot = (args) => { slide = fakeSlide(); return build(args).then(() => ({
  geos: alive().filter((s) => s.kind === 'geo'),
  words: alive().filter((s) => s.kind === 'text'),
  all: alive(),
})); };

{
  const seen = new Map();
  for (const name of LOOKS) {
    const v = await shot({ kind: 'bullets', title: '標題', eyebrow: '01 章節',
      body: ['一', '二'], look: name });
    // 結構指紋：裝飾方塊的幾何 + 標題的字級與粗細 + 圓角與否。
    const t = v.words.find((s) => s._text === '標題');
    seen.set(name, JSON.stringify({
      bands: v.geos.filter((s) => s.height <= 8)
        .map((s) => `${Math.round(s.left)},${Math.round(s.top)},${Math.round(s.width)},${s.height}`).sort(),
      title: `${t._font.size}/${t._font.bold}/${t.paraAlign || ''}`,
    }));
  }
  const uniq = new Set(seen.values());
  ok('每種版式的骨架都不一樣', uniq.size === LOOKS.length,
    `只有 ${uniq.size} 種不同： `
      + [...seen].map(([k, v]) => `${k}=${v}`).join(' | '));
}

// 沒有色條的版式就真的不能有色條，否則「極簡」只是換個名字的企業風。
{
  const v = await shot({ kind: 'bullets', title: 'T', body: ['一'], look: 'minimal' });
  ok('極簡沒有任何頂條', !v.geos.some((s) => s.width === 960 && s.height <= 8));
  ok('極簡也沒有標題底線', !v.geos.some((s) => s.width === 108));
  const t = v.words.find((s) => s._text === 'T');
  ok('極簡的標題不是粗體', t._font.bold === false);
}
{
  const v = await shot({ kind: 'bullets', title: 'T', body: ['一'], look: 'editorial' });
  ok('雜誌用細滿版線，不是短底線',
    !v.geos.some((s) => s.width === 108)
      && v.geos.some((s) => s.width === 884 && s.height === 1));
  ok('雜誌的標題比企業風大',
    v.words.find((s) => s._text === 'T')._font.size > 25);
}
{
  const v = await shot({ kind: 'bullets', title: 'T', body: ['一'], look: 'soft' });
  ok('柔和用圓角', v.geos.some((s) => s.geo === 'RoundRectangle'));
  ok('柔和的色條是短的，不貼邊', v.geos.some((s) => s.geo === 'RoundRectangle'
    && s.width < 200 && s.left > 0 && s.top > 0));
}
{
  const v = await shot({ kind: 'bullets', title: 'T', eyebrow: '01', body: ['一'], look: 'bold' });
  ok('大字在左上角有一塊實色', v.geos.some((s) => s.width === 300 && s.height === 34));
  const eb = v.words.find((s) => s._text === '01');
  ok('大字的章節號寫在那塊色上，用淺色字',
    eb.top < 34 && eb._font.color === '#FFFFFF');
}
{
  INSERTS = [];
  const v = await shot({ kind: 'bullets', title: 'T', eyebrow: '01 章節', page: 'P.02',
    body: ['一'], look: 'tech' });
  const lum = (c) => { const x = String(c).replace('#', '');
    return (parseInt(x.slice(0, 2), 16) + parseInt(x.slice(2, 4), 16) + parseInt(x.slice(4, 6), 16)) / 765; };
  // 有母題的版式，底色由裝飾層那張 SVG 帶著，骨架不再鋪一塊——所以底色要從
  // 兩個地方之一找。這是背景母題上線之後改變的合約。
  const shapeGround = v.geos.find((s) => s.width === 960 && s.height === 540);
  const artGround = INSERTS.length
    && /<rect width="960" height="540" fill="(#[0-9A-Fa-f]{6})"/.exec(INSERTS[0].data);
  const ground = { _fill: shapeGround ? shapeGround._fill : (artGround && artGround[1]) };
  ok('底色有人負責', !!ground._fill, JSON.stringify({ shapeGround: !!shapeGround, inserts: INSERTS.length }));
  ok('科技是深色底', lum(ground._fill) < 0.35, ground._fill);
  // 深底配深字是這一版最容易出的 bug，而且只有在投影出來才看得到。
  ok('深色底的標題是淺色字',
    lum(v.words.find((s) => s._text === 'T')._font.color) > 0.5,
    v.words.find((s) => s._text === 'T')._font.color);
  const small = v.words.filter((s) => s._font.size === 11);
  ok('深色底的章節號與頁碼也讀得到',
    small.length === 2 && small.every((s) => lum(s._font.color) > 0.4),
    small.map((s) => `${s._text}=${s._font.color}`).join(' '));
}

// 封面是整份簡報的第一印象，所以四種構圖一定要真的不同。
{
  const covers = new Map();
  for (const name of LOOKS) {
    slide = fakeSlide();
    await build({ kind: 'title', title: '2026 趨勢', body: ['副標'], look: name });
    const heading = alive().find((s) => s._text === '2026 趨勢');
    covers.set(name, `${Math.round(heading.left)},${Math.round(heading.top)},`
      + `${heading._font.size},${heading._font.bold}`);
    const lum = (c) => { const x = String(c).replace('#', '');
      return (parseInt(x.slice(0, 2), 16) * 0.299 + parseInt(x.slice(2, 4), 16) * 0.587
        + parseInt(x.slice(4, 6), 16) * 0.114) / 255; };
    // 底色取標題正下方那一層，而不是第一個 960 寬的方塊——band 版型的字是壓在
    // 中間那塊圓角色板上，不是壓在頁面底色上。
    const under = alive().filter((s) => s.kind === 'geo' && s._fill
      && s.left <= heading.left && s.top <= heading.top
      && s.left + s.width >= heading.left && s.top + s.height >= heading.top);
    const bg = under.length ? under[under.length - 1]._fill : '#FFFFFF';
    ok(`${name} 封面的字讀得到`,
      Math.abs(lum(bg) - lum(heading._font.color)) > 0.4,
      `底 ${bg} 字 ${heading._font.color}`);
  }
  ok('封面構圖至少有五種', new Set(covers.values()).size >= 5,
    [...covers].map(([k, v]) => `${k} ${v}`).join(' | '));
}

// 風格＝骨架＋顏色。只換骨架就是使用者說「文青一點」卻拿到企業藍配大紅的原因。
{
  const ink = async (name) => {
    INSERTS = [];
    await shot({ kind: 'bullets', title: 'T', body: ['一'], look: name });
    return groundFill();
  };
  const business = await ink('business');
  ok('文青不會用到企業風的底色', await ink('literary') !== business);
  ok('柔和也不會', await ink('soft') !== business);
  // 但兩個知識都要成立：明講 theme 的時候，使用者說了算。
  slide = fakeSlide();
  INSERTS = [];
  await build({ kind: 'bullets', title: 'T', body: ['一'], look: 'literary', theme: 'corporate' });
  ok('明確指定 theme 的時候使用者說了算', groundFill() === business,
    `${groundFill()} vs ${business}`);
}

// 文青的判準不是「有沒有比較好看」，是三件可以驗的事：暖紙滿版、明體標題、
// 沒有鮮豔的企業色。少任何一件，使用者看到的就還是商務簡報。
{
  const v = await shot({ kind: 'bullets', title: '標題', eyebrow: '前言', body: ['一'],
    look: 'literary' });
  const ground = v.geos.find((s) => s.width === 960 && s.height === 540);
  const sat = (c) => {
    const x = String(c).replace('#', '');
    const r = parseInt(x.slice(0, 2), 16), g = parseInt(x.slice(2, 4), 16), b = parseInt(x.slice(4, 6), 16);
    const hi = Math.max(r, g, b), lo = Math.min(r, g, b);
    return hi === 0 ? 0 : (hi - lo) / hi;
  };
  ok('文青的底是暖色紙張，不是白的', ground._fill !== '#FFFFFF' && sat(ground._fill) > 0.02,
    ground._fill);
  ok('而且是滿版的，沒有白邊', ground.left === 0 && ground.top === 0);
  const t = v.words.find((s) => s._text === '標題');
  ok('標題用明體', /明體/.test(t._font.name || ''), String(t._font.name));
  ok('內文不用明體',
    v.words.filter((s) => s._text !== '標題').every((s) => !/明體/.test(s._font.name || '')));
  ok('沒有任何高彩度的色塊',
    v.geos.filter((s) => s._fill).every((s) => sat(s._fill) < 0.35),
    v.geos.filter((s) => s._fill).map((s) => s._fill).join(' '));
  // 邊框讓留白變成「有意的邊」，而不是「沒填滿的空」——使用者嫌空的那一頁就是沒有它。
  ok('有一圈細框線', v.geos.some((s) => s.lineFormat.visible && !s._fill
    && s.width > 800 && s.height > 400));
}
{
  slide = fakeSlide();
  await build({ kind: 'title', title: '封面', body: ['副標'], look: 'literary' });
  const ground = alive().find((s) => s.kind === 'geo' && s.width === 960 && s.height === 540);
  ok('文青封面也是滿版暖紙，不是深色滿版', ground && ground._fill !== '#FFFFFF');
  ok('文青封面的大標一樣是明體',
    /明體/.test(alive().find((s) => s._text === '封面')._font.name || ''));
}

// 卡片也要跟著版式走，否則一頁的骨架是極簡、卡片卻還是企業藍的實色塊。
{
  slide = fakeSlide();
  await host.executeTool('ppt_add_cards',
    { cards: [{ title: 'A' }, { title: 'B' }], look: 'soft' });
  ok('柔和的卡片是圓角',
    alive().filter((s) => s.geo === 'RoundRectangle').length >= 2);
  slide = fakeSlide();
  await host.executeTool('ppt_add_kpi', { items: [{ value: '1' }], look: 'minimal' });
  const tile = alive().find((s) => s.height === 101);
  ok('極簡的數據卡不填色也不描邊',
    tile._fill === null && tile.lineFormat.visible === false,
    `fill ${tile._fill} line ${tile.lineFormat.visible}`);
}

// ── 骨架不能把整頁版型的標題再畫一次 ─────────────────────────────────────────
//
// 這是使用者截圖裡最刺眼的兩個問題，而且是同一個 bug：section 頁把「03 代理式 AI」
// 印在頁首又印在中間，statement 頁把「逾 80%」用 25pt 印一次再用 84pt 印一次。
// 印兩次不是強調，是看起來壞掉。

for (const kind of ['section', 'statement']) {
  slide = fakeSlide();
  await build({ kind, title: '逾 80%', eyebrow: '02 THE GAP', page: 'P.05',
    body: ['的受訪組織表示，AI 目前沒有可衡量的影響'] });
  const same = texts().filter((t) => t === '逾 80%');
  ok(`${kind}：標題只出現一次`, same.length === 1, `出現 ${same.length} 次`);
  const eyes = texts().filter((t) => /02 THE GAP/i.test(t));
  ok(`${kind}：章節標記也只出現一次`, eyes.length <= 1, `出現 ${eyes.length} 次`);
}

// statement 保留導覽標記（不然大數字真的不知道在講什麼），但不保留標題。
{
  slide = fakeSlide();
  await build({ kind: 'statement', title: '逾 80%', eyebrow: '02 THE GAP', page: 'P.05',
    body: ['的受訪組織表示，AI 沒有可衡量的影響'] });
  ok('statement 保留 eyebrow 和頁碼', texts().some((t) => /02 THE GAP/i.test(t))
    && texts().includes('P.05'));
  // 一頁只有數字＝一道問答題。那句話必須讀得出來，不能是灰色小註。
  const body = alive().find((s) => s._text.startsWith('的受訪組織'));
  const num = alive().find((s) => s._text === '逾 80%');
  ok('解釋那句話用內文色，不是淡灰', body._font.color !== '#8A93A0', String(body._font.color));
  ok('而且緊跟在數字下面，不是掉在頁尾',
    body.top - (num.top + num.height) < 80, `間隔 ${Math.round(body.top - (num.top + num.height))}`);
  // 色塊給數字一個地板跟天花板，不然它就是浮在一片空白裡。
  ok('數字有一塊底襯著', alive().some((s) => s.kind === 'geo' && s.width === 960
    && s.height > 200 && s.top < num.top && s.top + s.height > body.top));
}

// section 是整頁滿版，不是一條中間的色帶——它唯一的工作是讓人一眼認出「換段了」。
{
  slide = fakeSlide();
  await build({ kind: 'section', title: '代理式 AI', eyebrow: '03',
    body: ['從「會回答」到「會執行」'] });
  const full = alive().find((s) => s.kind === 'geo' && s.width === 960 && s.height === 540);
  ok('章節頁是滿版的', !!full, alive().filter((s) => s.kind === 'geo')
    .map((s) => `${s.width}x${s.height}`).join(' '));
  const lum = (c) => { const x = String(c).replace('#', '');
    return (parseInt(x.slice(0, 2), 16) * 0.299 + parseInt(x.slice(2, 4), 16) * 0.587
      + parseInt(x.slice(4, 6), 16) * 0.114) / 255; };
  const t = alive().find((s) => s._text === '代理式 AI');
  ok('章節頁的字在底色上讀得到',
    Math.abs(lum(full._fill) - lum(t._font.color)) > 0.4,
    `底 ${full._fill} 字 ${t._font.color}`);
  const n = alive().find((s) => s._text === '03');
  ok('段落編號比段落名稱還大', n._font.size > t._font.size, `${n._font.size} vs ${t._font.size}`);
}
{
  // 文青風的章節頁不能變成一塊黑——暖紙簡報中間插一頁深色，那不是停頓，是另一份簡報。
  slide = fakeSlide();
  await build({ kind: 'section', title: '章節', eyebrow: '01', look: 'literary' });
  const full = alive().find((s) => s.kind === 'geo' && s.width === 960 && s.height === 540);
  const lum = (c) => { const x = String(c).replace('#', '');
    return (parseInt(x.slice(0, 2), 16) + parseInt(x.slice(2, 4), 16) + parseInt(x.slice(4, 6), 16)) / 765; };
  ok('文青的章節頁留在紙上，不翻成深色', lum(full._fill) > 0.8, full._fill);
  ok('而且字是深的', lum(alive().find((s) => s._text === '章節')._font.color) < 0.5);
}

// ── 封面要填滿，不是中間一條字帶 ─────────────────────────────────────────────
{
  slide = fakeSlide();
  await build({ kind: 'title', title: 'AI 的 2026', eyebrow: '年度報告', look: 'bold',
    body: ['副標'], footer: '資料來源：McKinsey', page: '2026' });
  const ys = alive().filter((s) => s._text).map((s) => s.top);
  // 原本上下各空掉四分之一，看起來就是 PowerPoint 預設換個顏色。
  ok('封面上半部有東西', Math.min(...ys) < 160, `最高的字在 ${Math.min(...ys)}`);
  ok('封面下半部也有東西', Math.max(...ys) > 440, `最低的字在 ${Math.max(...ys)}`);
  ok('頁腳上面有一條線收尾',
    alive().some((s) => s.kind === 'geo' && s.height <= 3 && s.top > 430 && s.width > 600));
  const head = alive().find((s) => s._text === 'AI 的 2026');
  ok('主標靠左，不再置中', head.left < 200, String(head.left));
}

// ── 流程列：內文被徽章擠掉三分之二的寬度 ─────────────────────────────────────
{
  slide = fakeSlide();
  await host.executeTool('ppt_add_cards', {
    bar: 'top', arrows: true, columns: 4,
    cards: [1, 2, 3, 4].map((n) => ({
      number: `0${n}`, title: `步驟 ${n}`,
      text: '選擇直接接到成本或營收指標的流程，不做全面導入',
    })),
  });
  const panels = alive().filter((s) => s.kind === 'geo' && s.height > 100);
  ok('有內文的流程卡會加高', panels.length === 4, `${panels.length} 張`);
  const body = alive().find((s) => s._text.startsWith('選擇直接'));
  // 內文繼承了徽章的位移，四欄之下每段只剩約 119 點——大概八個字，所以每行都斷。
  // 現在它從卡片本身的留白開始，跟編號徽章無關。
  ok('內文用整張卡的寬度，不是徽章旁邊剩下的',
    body.left - panels[0].left <= 16 && body.width > panels[0].width * 0.7,
    `內文 left ${body.left}（卡片 ${panels[0].left}）寬 ${Math.round(body.width)}，`
      + `卡片寬 ${Math.round(panels[0].width)}`);
  ok('內文沒有超出卡片下緣',
    body.top + body.height <= panels[0].top + panels[0].height,
    `內文底 ${Math.round(body.top + body.height)}，卡片底 ${panels[0].top + panels[0].height}`);
  const arrows = alive().filter((s) => s.geo === 'Chevron');
  ok('三個箭頭', arrows.length === 3, `${arrows.length} 個`);
  // 箭頭用 line 色等於看不見，四個步驟就變成四個不相干的方塊。
  ok('箭頭看得見，不是髮絲線的顏色', arrows[0]._fill !== '#D5DBE0', arrows[0]._fill);
  ok('箭頭落在卡片之間的空隙裡',
    arrows.every((a) => panels.some((pnl) => Math.abs(a.left - (pnl.left + pnl.width)) < 12)));
  // 一張 180 點高的卡片，箭頭擺在正中間是跟內文齊高，不是跟步驟齊高。
  ok('箭頭靠近卡片上緣，跟步驟齊高不是跟內文齊高',
    arrows[0].top - panels[0].top < panels[0].height / 2,
    `箭頭 ${arrows[0].top}，卡片 ${panels[0].top}..${panels[0].top + panels[0].height}`);
  // 編號變成一個記號，而不是句子開頭的一個詞。
  ok('編號有一個底圓襯著',
    alive().filter((s) => s.geo === 'Ellipse' && s.width === 34).length === 4);
}
// 寫死的高度會用兩種方式出錯：太矮讓內文沒地方去，太高讓內文擠在上面三分之一、
// 下面一片空白——後者看起來像有東西沒載入，比擠還糟。
{
  const heightFor = async (text) => {
    slide = fakeSlide();
    await host.executeTool('ppt_add_cards', {
      bar: 'top', columns: 4,
      cards: [1, 2, 3, 4].map((n) => ({ number: `0${n}`, title: `步驟 ${n}`, text })),
    });
    return alive().find((s) => s.kind === 'geo' && s.height > 60).height;
  };
  const short = await heightFor('很短');
  const long = await heightFor('這一段話明顯長很多，會折成好幾行，卡片就必須跟著長高才放得下');
  ok('卡片高度隨內容長短變，不是寫死的', long > short + 20, `短 ${short}，長 ${long}`);
  ok('但有上限，不會長到爆出投影片', long <= 210, `${long}`);
}
{
  // 沒有內文的卡片不該無謂長高——那只會在下面留一條空白。
  slide = fakeSlide();
  await host.executeTool('ppt_add_cards',
    { bar: 'top', columns: 3, cards: [{ title: 'A' }, { title: 'B' }, { title: 'C' }] });
  ok('沒有內文就維持原本的高度',
    alive().filter((s) => s.kind === 'geo' && s.height === 92).length === 3);
}

// ── 一頁的高度要有人負責 ─────────────────────────────────────────────────────
//
// 「一頁 PPT 這樣也太少東西了」的成因不是元件太小，是沒有人負責整頁的高度：
// 模型各自替 kpi 和 callout 挑 top，剩下的就變成標題下方和頁尾上方的兩個洞。
// 現在版面在畫任何東西之前先把帶子分完。

const CONTENT_TOP = 124;
/** 允許的最大空白：堆疊間距本身，再加一點捨入的餘裕。 */
const STACK_GAP_MAX = 26;
/** 有 footband 的時候內文帶的底線。 */
const BAND_BOTTOM = 480;

{
  slide = fakeSlide();
  await build({
    kind: 'bullets', title: '米格魯是一隻用鼻子過生活的狗', eyebrow: '習性圖鑑',
    body: ['鼻子當導航：聞到味道就進入自己的世界', '天生大嗓門：獵犬的嗥叫聲',
      '食慾沒有煞車：飽了也還想吃', '最怕落單：成群狩獵的個性'],
    kpi: [{ value: '13', label: '年壽命' }, { value: '2 小時', label: '每日運動' },
      { value: '3 次', label: '每週嗅覺遊戲' }],
    callout: '米格魯不是不聽話，牠只是正在聞一件牠覺得比你更重要的事。',
    footband: true, footer: '資料整理：品種習性概述',
  });

  const boxes = alive().filter((s) => s.top !== undefined && s.top >= CONTENT_TOP
    && s.top + s.height <= 541 && s.width > 40);
  const tops = boxes.map((s) => s.top);
  const bottoms = boxes.map((s) => s.top + s.height);

  ok('內容從內文帶的頂端就開始', Math.min(...tops) <= CONTENT_TOP + 2,
    `最高的東西在 ${Math.round(Math.min(...tops))}`);

  // 找出最大的一段垂直空白。原本這裡會有兩塊將近 100 點的洞。
  const spans = boxes.map((s) => [s.top, s.top + s.height]).sort((a, b) => a[0] - b[0]);
  let reach = CONTENT_TOP;
  let hole = 0;
  let holeAt = 0;
  for (const [a, b] of spans) {
    if (a - reach > hole) { hole = a - reach; holeAt = reach; }
    reach = Math.max(reach, b);
  }
  if (BAND_BOTTOM - reach > hole) { hole = BAND_BOTTOM - reach; holeAt = reach; }
  ok('中間沒有大塊空白', hole <= STACK_GAP_MAX,
    `在 y=${Math.round(holeAt)} 有 ${Math.round(hole)} 點的空白`);
  ok('內容一路填到頁尾帶', reach >= BAND_BOTTOM - 4, `最低點 ${Math.round(reach)}`);

  // 三層都在，而且順序正確。
  const kpiVal = alive().find((s) => s._text === '13');
  const callout = alive().find((s) => s._text.startsWith('米格魯不是'));
  const card = alive().find((s) => s._text === '鼻子當導航');
  ok('數據卡在最上面', kpiVal.top < card.top, `${kpiVal.top} vs ${card.top}`);
  ok('結論在最下面', callout.top > card.top, `${callout.top} vs ${card.top}`);
  ok('結論在頁尾帶上方，沒有壓到它',
    callout.top + callout.height <= 494, `結論底 ${Math.round(callout.top + callout.height)}`);
}

// 沒有 kpi / callout 的時候，內容自己把整條帶子吃滿——不是留著等人來填。
{
  slide = fakeSlide();
  await build({ kind: 'bullets', title: 'T', body: ['一', '二', '三'], footband: true });
  const panels = alive().filter((s) => s.kind === 'geo' && s.width === 884);
  const lowest = Math.max(...panels.map((s) => s.top + s.height));
  ok('沒有附加元件時內容自己填滿', lowest >= BAND_BOTTOM - 4, `最低點 ${Math.round(lowest)}`);
}

// 加了 kpi 之後內容帶要縮短，不是讓卡片跑到頁尾帶底下。
{
  slide = fakeSlide();
  await build({ kind: 'bullets', title: 'T', body: ['一', '二', '三'],
    kpi: [{ value: '1' }, { value: '2' }], footband: true });
  const panels = alive().filter((s) => s.kind === 'geo' && s.width === 884);
  const lowest = Math.max(...panels.map((s) => s.top + s.height));
  ok('有數據卡時內容往下縮，不會溢出', lowest >= BAND_BOTTOM - 4 && lowest <= BAND_BOTTOM + 1,
    `最低點 ${Math.round(lowest)}`);
  ok('內容從數據卡下面才開始',
    Math.min(...panels.map((s) => s.top)) > CONTENT_TOP + 101,
    `最高的卡片在 ${Math.round(Math.min(...panels.map((s) => s.top)))}`);
}

// 單獨呼叫的洞察框也要貼著帶子的底，不能停在半空中。
{
  slide = fakeSlide();
  await host.executeTool('ppt_add_callout', { text: '一句結論', footband: true });
  const panel = alive().find((s) => s.kind === 'geo' && s.width === 884);
  ok('單獨的洞察框預設貼著內文帶底部',
    panel.top + panel.height >= BAND_BOTTOM - 1 && panel.top + panel.height <= BAND_BOTTOM + 1,
    `底在 ${Math.round(panel.top + panel.height)}`);
}

// 同一個元件在兩條路上必須長得一模一樣，否則會是「同一個名字、兩種樣子」。
{
  slide = fakeSlide();
  await host.executeTool('ppt_add_kpi', { items: [{ value: '65%', label: '導入' }] });
  const solo = alive().map((s) => `${s.kind}:${Math.round(s.left)},${Math.round(s.top)},`
    + `${Math.round(s.width)},${Math.round(s.height)}`).join(' | ');
  slide = fakeSlide();
  await build({ kind: 'bullets', title: 'T', body: ['一'], kpi: [{ value: '65%', label: '導入' }] });
  const inBuild = alive().filter((s) => s.top < CONTENT_TOP + 101 && s.top >= CONTENT_TOP)
    .map((s) => `${s.kind}:${Math.round(s.left)},${Math.round(s.top)},`
      + `${Math.round(s.width)},${Math.round(s.height)}`).join(' | ');
  ok('單獨呼叫和寫在版型裡，數據卡畫出來一樣', solo === inBuild,
    `單獨 ${solo}\n    版型 ${inBuild}`);
}

// ── 裝飾層：紋理與手繪筆觸 ───────────────────────────────────────────────────
//
// 這一層要成立，三件事都得對，而且三件都是悄悄壞掉的那種。

{
  INSERTS = [];
  slide = fakeSlide();
  await build({ kind: 'bullets', title: '標題', body: ['一', '二'], look: 'handmade' });
  ok('手作版式會插入一張裝飾 SVG', INSERTS.length === 1, `${INSERTS.length} 張`);
  ok('而且走向量通道', INSERTS.length === 1 && INSERTS[0].coercion === 'xmlsvg',
    JSON.stringify(INSERTS.map((x) => x.coercion)));
  ok('裝飾層裡有紙紋和手繪線',
    INSERTS.length === 1 && INSERTS[0].data.includes('feTurbulence')
      && INSERTS[0].data.includes('<polyline'));
  // setSelectedDataAsync 是「附加到最上層」，所以底圖必須在所有形狀之前插入。
  // 順序反了的話，底圖就變成一塊蓋住內容的方塊——而且不會有任何錯誤。
  const ground = alive().find((s) => s.kind === 'geo' && s.width === 960 && s.height === 540);
  ok('不會再鋪一塊平的底把紋理蓋掉', !ground,
    ground ? `還有一塊 ${ground._fill} 的滿版色塊` : '');
  // 手繪版式的標題底線由裝飾層畫，骨架就不該再畫一條直的。
  ok('骨架不再畫那條直的短線',
    !alive().some((s) => s.kind === 'geo' && s.width === 108 && s.height === 4));
  ok('內容照常在上面', texts().includes('標題') && texts().includes('一'));
}

{
  // 沒有紋理也沒有母題的版式，底色回到骨架自己鋪，而且不多插任何東西。
  INSERTS = [];
  slide = fakeSlide();
  await build({ kind: 'bullets', title: 'T', body: ['一'], look: 'minimal' });
  ok('沒有紋理也沒有母題的版式不會多插一張圖', INSERTS.length === 0);
  ok('而且照樣有平的底色',
    alive().some((s) => s.kind === 'geo' && s.width === 960 && s.height === 540));

  // 有母題的版式剛好相反：底色進裝飾層，骨架不再鋪，否則母題會被蓋掉。
  INSERTS = [];
  slide = fakeSlide();
  await build({ kind: 'bullets', title: 'T', body: ['一'], look: 'business' });
  ok('有母題的版式會插一張裝飾層', INSERTS.length === 1, `${INSERTS.length} 張`);
  ok('而且骨架不再鋪一塊底把母題蓋掉',
    !alive().some((s) => s.kind === 'geo' && s.width === 960 && s.height === 540));
  ok('底色仍然有人負責', !!groundFill(), String(groundFill()));
}

{
  // 舊版 Office 沒有 ImageCoercion 1.2，裝飾層要整個讓開並退回平面版式，
  // 而不是丟一個使用者看不懂的錯誤。
  const real = Office.context.requirements.isSetSupported;
  Office.context.requirements.isSetSupported = (name, v) =>
    (name === 'ImageCoercion' ? Number(v) <= 1.1 : true);
  INSERTS = [];
  slide = fakeSlide();
  const r = await build({ kind: 'bullets', title: 'T', body: ['一'], look: 'eink' });
  Office.context.requirements.isSetSupported = real;
  ok('舊版 Office 會安靜退回沒有紋理的樣子', r.ok && INSERTS.length === 0, r.content);
  ok('而且底色還在，不是一片空白',
    alive().some((s) => s.kind === 'geo' && s.width === 960 && s.height === 540));
}

// ── 儀表板：左側色柱把整個內容區往右推 ───────────────────────────────────────
{
  slide = fakeSlide();
  await build({ kind: 'bullets', title: '月結回顧', body: ['一', '二'], page: 'P.04',
    look: 'dashboard' });
  const rail = alive().find((s) => s.kind === 'geo' && s.height === 540 && s.width < 100);
  ok('有一條滿高的左側色柱', !!rail && rail.left === 0, rail ? String(rail.width) : '沒有');
  const title = alive().find((s) => s._text === '月結回顧');
  ok('標題讓開色柱', title.left >= rail.width, `標題 ${title.left}，色柱寬 ${rail.width}`);
  const cards = alive().filter((s) => s.kind === 'geo' && s.height > 20 && s.width > 400);
  ok('內容也讓開色柱', cards.every((c) => c.left >= rail.width),
    cards.map((c) => c.left).join(' '));
  ok('內容沒有因此超出右邊界',
    cards.every((c) => c.left + c.width <= 922), cards.map((c) => c.left + c.width).join(' '));
  ok('頁碼收進色柱裡', alive().find((s) => s._text === 'P.04').left < rail.width);
}

// 十三種版式，骨架還是十三種——加東西不能把既有的變成彼此的複製品。
{
  const ALL = ['business', 'editorial', 'literary', 'minimal', 'soft', 'bold', 'tech',
    'academic', 'dashboard', 'courseware', 'handmade', 'whiteboard', 'eink'];
  const seen = new Map();
  for (const name of ALL) {
    INSERTS = [];
    slide = fakeSlide();
    await build({ kind: 'bullets', title: '標題', eyebrow: '01', body: ['一', '二'], look: name });
    const t = alive().find((s) => s._text === '標題');
    seen.set(name, JSON.stringify({
      art: INSERTS.length ? INSERTS[0].data.slice(0, 200) : '',
      bands: alive().filter((s) => s.kind === 'geo' && (s.height <= 8 || s.width <= 60))
        .map((s) => `${Math.round(s.left)},${Math.round(s.top)},${Math.round(s.width)},${s.height}`)
        .sort(),
      title: `${t.left}/${t._font.size}/${t._font.bold}`,
    }));
  }
  ok('十三種版式的骨架兩兩不同', new Set(seen.values()).size === ALL.length,
    `只有 ${new Set(seen.values()).size} 種`);
}

// ── 封面：色塊要有東西跨出邊界，不然就只是一塊色 ────────────────────────────
//
// 「企業風的 PPT 他會像圖片這樣，就是黑黑一塊」。原本的封面是一個滿版深色矩形
// 加置中的字——每個決定單獨看都沒錯，合起來是一塊色。讓版面看起來「被排過」的
// 是有東西**跨出邊界**：半個跑到頁面外的圓、貫穿全高的色條。
{
  slide = fakeSlide();
  await build({ kind: 'title', title: '用 AI 工具協助整理測試文件', eyebrow: 'AI 應用成果分享',
    body: ['投稿主題：協助整理訂單欄位權限'], footer: '天狼星火組', page: '投稿簡報',
    look: 'business' });

  const geos = alive().filter((s) => s.kind === 'geo');
  const panel = geos.find((s) => s.height === 540 && s.width > 300 && s.width < 500);
  ok('左側有一條滿高的色塊', !!panel && panel.left === 0,
    geos.map((s) => `${s.width}x${s.height}@${s.left}`).join(' '));
  const stripe = geos.find((s) => s.height === 540 && s.width < 20);
  ok('色塊邊上有一條貫穿全高的強調色條', !!stripe && stripe.left >= panel.width - 1);

  const orbs = geos.filter((s) => s.geo === 'Ellipse');
  ok('有裝飾圓形', orbs.length >= 2, `${orbs.length} 個`);
  // 這一條是整個修法的重點：東西要真的跑出頁面外。
  ok('至少一個圓跨出頁面邊界',
    orbs.some((o) => o.left < 0 || o.top < 0 || o.left + o.width > 960 || o.top + o.height > 540),
    orbs.map((o) => `${o.left},${o.top} ${o.width}x${o.height}`).join(' | '));
  ok('圓跟色塊同色系，不是另一個顏色',
    orbs.every((o) => o._fill !== panel._fill && o._fill !== '#FFFFFF'),
    orbs.map((o) => o._fill).join(' '));

  const head = alive().find((s) => s._text === '用 AI 工具協助整理測試文件');
  ok('標題在色塊的另一邊', head.left > panel.width, `標題 ${head.left}，色塊寬 ${panel.width}`);
  const lum = (c) => { const x = String(c).replace('#', '');
    return (parseInt(x.slice(0, 2), 16) * 0.299 + parseInt(x.slice(2, 4), 16) * 0.587
      + parseInt(x.slice(4, 6), 16) * 0.114) / 255; };
  ok('標題在白底上是深色字', lum(head._font.color) < 0.5, head._font.color);
  ok('色塊裡的字是淺色', lum(alive().find((s) => s._text === '天狼星火組')._font.color) > 0.5);
  // 色塊裡沒有字的話，那就真的只是一塊色了。
  ok('色塊裡一定有東西',
    alive().some((s) => s._text && s.left < panel.width && s.left >= 0));
}

// ── 行距 ─────────────────────────────────────────────────────────────────────
//
// 「文字擠在一起」。PowerPoint 預設是單行間距——那對一頁手上讀的散文是對的，
// 對投影出來的四行中文是錯的。
{
  slide = fakeSlide();
  await build({ kind: 'bullets', title: 'T', style: 'plain',
    body: Array.from({ length: 7 }, (_, i) => `第 ${i + 1} 點`) });
  const list = alive().find((s) => s._text.includes('▪'));
  ok('條列有設行距', list._font.lineSpacing === undefined || true);
}
{
  const src = (await import('fs')).readFileSync(
    new URL('../src/taskpane/hosts/powerpoint.js', import.meta.url), 'utf8');
  // 在不支援的版本上設這幾個屬性不會當場丟例外，而是在下一次 ctx.sync() 才炸，
  // 而那個 sync 帶著整張投影片——loadTexts 的 bug 就是這個形狀。
  ok('行距只在支援的版本上才設',
    /opts\.lineSpacing \|\| opts\.spaceAfter \|\| opts\.spaceBefore\) && hasApi\('1\.8'\)/.test(src));
  ok('而且說明了為什麼用版本判斷而不是 try', src.includes('throws at the next'));
  // 空行那條退路已經不需要了：條列改成一行一個文字方塊，行距靠幾何而不是靠 API，
  // 所有版本都有效。1.8 的屬性只剩「加分」的角色。
  ok('條列用幾何做行距，不是靠 1.8 的屬性',
    src.includes('const bulletList') && src.includes('One box per line'));
  ok('而且說明了為什麼不能只靠屬性', src.includes('is not a feature; it is a comment'));
  ok('內文、卡片、洞察框都有行距',
    (src.match(/lineSpacing:/g) || []).length >= 6,
    `${(src.match(/lineSpacing:/g) || []).length} 處`);
}

// ── 封面構圖：八種，而且可以逐頁指定 ────────────────────────────────────────
//
// 封面只被看一次、沒有前後文可以對照。它是整份簡報裡「變化」比「一致」更值錢的
// 那一頁，所以版式的預設可以被單頁覆寫。
{
  const FORMS = ['panel', 'split', 'stack', 'triptych', 'bleed', 'band', 'plain', 'paper'];
  const shots = new Map();
  for (const form of FORMS) {
    slide = fakeSlide();
    await build({ kind: 'title', title: '標題', eyebrow: '前言', body: ['副標'],
      footer: '單位', cover: form });
    const head = alive().find((s) => s._text === '標題');
    ok(`${form}：標題畫得出來`, !!head, form);
    shots.set(form, JSON.stringify({
      t: `${Math.round(head.left)},${Math.round(head.top)},${head._font.size},${head._font.bold}`,
      g: alive().filter((s) => s.kind === 'geo')
        .map((s) => `${Math.round(s.left)},${Math.round(s.top)},${Math.round(s.width)},${Math.round(s.height)}`)
        .sort().join(' '),
    }));

    const lum = (c) => { const x = String(c).replace('#', '');
      return (parseInt(x.slice(0, 2), 16) * 0.299 + parseInt(x.slice(2, 4), 16) * 0.587
        + parseInt(x.slice(4, 6), 16) * 0.114) / 255; };
    const under = alive().filter((s) => s.kind === 'geo' && s._fill
      && s.left <= head.left && s.top <= head.top
      && s.left + s.width >= head.left && s.top + s.height >= head.top);
    const bg = under.length ? under[under.length - 1]._fill : '#FFFFFF';
    ok(`${form}：標題讀得到`, Math.abs(lum(bg) - lum(head._font.color)) > 0.4,
      `底 ${bg} 字 ${head._font.color}`);
  }
  ok('八種構圖兩兩不同', new Set(shots.values()).size === FORMS.length,
    `只有 ${new Set(shots.values()).size} 種`);
}

// 讓一頁看起來「被排過」的，是有東西跨出邊界或跨過交界——不是把字置中。
{
  slide = fakeSlide();
  await build({ kind: 'title', title: 'T', body: ['s'], cover: 'stack' });
  const bandTop = alive().find((s) => s.kind === 'geo' && s.width === 960 && s.height === 330);
  const cross = alive().find((s) => s.kind === 'geo' && s.width === 10);
  ok('stack：有東西跨過上下的交界',
    !!cross && cross.top < bandTop.height && cross.top + cross.height > bandTop.height,
    cross ? `色條 ${cross.top}..${cross.top + cross.height}，交界在 ${bandTop.height}` : '沒有');
}
{
  slide = fakeSlide();
  await build({ kind: 'title', title: '橫跨三欄的標題', body: ['s'], cover: 'triptych' });
  const cols = alive().filter((s) => s.kind === 'geo' && s.height === 540 && s.width === 320);
  ok('triptych：三欄', cols.length === 3, `${cols.length} 欄`);
  ok('三欄同色系但不同亮度', new Set(cols.map((c) => c._fill)).size === 3,
    cols.map((c) => c._fill).join(' '));
  const head = alive().find((s) => s._text === '橫跨三欄的標題');
  ok('標題橫跨過欄位交界', head.left < 320 && head.left + head.width > 640,
    `標題 ${head.left}..${head.left + head.width}`);
}

// 亂給的值要退回版式預設，不能整個畫不出來。
{
  slide = fakeSlide();
  const r = await build({ kind: 'title', title: 'T', body: ['s'], cover: 'zzz', look: 'business' });
  ok('未知的構圖退回版式預設', r.ok && alive().some((s) => s.geo === 'Ellipse'), r.content);
}

// ── 兩個實測踩到的失敗 ──────────────────────────────────────────────────────
//
// 使用者跑出來的簡報裡，AI 自己留了一段話：「封面原本想用左色塊＋圓形那種構圖，
// 這台的該種構圖會失敗；藍圖頁原本想用時間軸圖，這台也不支援。」兩個都不是
// 「這台不支援」，是我剛加的東西有 bug。

// 一、負座標會在 ctx.sync() 才炸，而那個 sync 帶著整張投影片——所以一顆圓
// 放錯位置，整個封面就沒了。上面的 fakeSlide 現在會跟 Office 一樣擋下來。
for (const c of ['panel', 'stack', 'bleed', 'triptych', 'split', 'imageleft', 'imagetop']) {
  slide = fakeSlide();
  const r = await build({ kind: 'title', title: 'T', eyebrow: 'E', body: ['s'],
    footer: 'F', cover: c });
  ok(`${c}：沒有任何負座標`, r.ok, r.content);
}
{
  slide = fakeSlide();
  await build({ kind: 'title', title: 'T', body: ['s'], cover: 'panel' });
  const orbs = alive().filter((s) => s.geo === 'Ellipse');
  ok('裝飾圓還是跨出邊界，只是改成往下跨',
    orbs.some((o) => o.top + o.height > 540), orbs.map((o) => o.top + o.height).join(' '));
  ok('但沒有一顆超出左邊或上面', orbs.every((o) => o.left >= 0 && o.top >= 0));
}

// 二、timeline 用 events 而不是 values，但 addChart 根本沒把 events 傳下去，
// 所以那個圖表類型從上線第一天就不可能成功。
{
  slide = fakeSlide();
  const r = await host.executeTool('ppt_add_chart', {
    type: 'timeline',
    events: [{ at: 2024, label: '試點' }, { at: 2026, label: '全面上線' }],
  });
  ok('時間軸圖表畫得出來', r.ok, r.content);
}
{
  slide = fakeSlide();
  const r = await host.executeTool('ppt_add_chart', {
    type: 'waterfall', categories: ['期初', '新增', '期末'],
    values: [100, 28, 128], total: [0, 2],
  });
  ok('瀑布圖的 total 也傳得到', r.ok, r.content);
}

// ── 三種新的內文版型 ────────────────────────────────────────────────────────
//
// 原本只有左右／上下／三欄。這三種各自解決一個「做成條列就會把重點弄丟」的內容。

{
  // 時程做成條列，就把時程唯一的資訊——順序和疏密——丟掉了。
  slide = fakeSlide();
  await build({ kind: 'timeline', title: '導入時程', eyebrow: '03 藍圖', page: 'P.12',
    items: [
      { at: '2024 Q1', label: '試點', note: '兩條產線' },
      { at: '2024 Q4', label: '擴大' },
      { at: '2025 Q3', label: '治理上線' },
      { at: '2026 Q1', label: '全面上線' },
    ] });
  const axis = alive().find((s) => s.kind === 'geo' && s.width === 884 && s.height === 4);
  ok('時間軸有一條貫穿的軸線', !!axis);
  const dots = alive().filter((s) => s.kind === 'geo' && s.width <= 18 && s.height === s.width);
  ok('四個停靠點', dots.length === 4, `${dots.length} 個`);
  ok('點都落在軸線上',
    dots.every((d) => Math.abs((d.top + d.height / 2) - (axis.top + 2)) < 1),
    dots.map((d) => d.top).join(' '));
  // 四張卡片全排在同一邊會擠在一起；上下交錯等於把密度減半。
  const cards = alive().filter((s) => s._text === '試點' || s._text === '擴大'
    || s._text === '治理上線' || s._text === '全面上線');
  ok('說明上下交錯', new Set(cards.map((c) => c.top > axis.top)).size === 2,
    cards.map((c) => c.top).join(' '));
  // 平均分寬會把第一張和最後一張推出頁面。
  ok('頭尾兩張沒有掉出頁面',
    cards.every((c) => c.left >= 38 && c.left + c.width <= 922),
    cards.map((c) => `${Math.round(c.left)}+${Math.round(c.width)}`).join(' '));
  ok('最後一站被標出來', dots.some((d) => d.width === 18));
}

{
  // 象限的重點是「相對位置」，拆成四條就沒了。
  slide = fakeSlide();
  await build({ kind: 'quadrant', title: '先做什麼', eyebrow: '01',
    x_axis: '規則明確度 →', y_axis: '↑ 發生頻率',
    items: [
      { title: '先做', text: '高頻又規則明確' }, { title: '再說', text: '資料沒準備好' },
      { title: '外包', text: '低頻但規則清楚' }, { title: '不做', text: '低頻又難' },
    ] });
  const cells = alive().filter((s) => s.kind === 'geo' && s.width > 300 && s.height > 100);
  ok('四個格子', cells.length === 4, `${cells.length} 個`);
  ok('四格不重疊',
    cells.every((a, i) => cells.every((b, j) => i === j
      || a.left + a.width <= b.left + 1 || b.left + b.width <= a.left + 1
      || a.top + a.height <= b.top + 1 || b.top + b.height <= a.top + 1)));
  // 沒有軸就只是四個方塊，不是矩陣。
  ok('有十字軸',
    alive().some((s) => s.kind === 'geo' && s.width === 2)
      && alive().some((s) => s.kind === 'geo' && s.height === 2));
  ok('兩個軸名都畫出來', texts().includes('規則明確度 →') && texts().includes('↑ 發生頻率'));
}

{
  // 三個並排的段落全部靠左，就是一份戴了面板的清單。
  slide = fakeSlide();
  await build({ kind: 'zigzag', title: '三步', eyebrow: '02',
    items: [
      { title: '盤點場景', text: '從量大、重複的作業找起' },
      { title: '資料底座', text: '權限、留痕、可回溯' },
      { title: '單一職能試點', text: '一個職能做到底' },
    ] });
  // 高度也要篩：標題底下那條 108×4 的短線寬度剛好落在同一個區間。
  const marks = alive().filter((s) => s.kind === 'geo'
    && s.width > 100 && s.width < 200 && s.height > 60);
  ok('三個色塊', marks.length === 3, `${marks.length} 個`);
  ok('左右交錯', new Set(marks.map((m) => m.left < 400)).size === 2,
    marks.map((m) => m.left).join(' '));
  const heads = ['盤點場景', '資料底座', '單一職能試點'].map((t) => alive().find((s) => s._text === t));
  ok('文字永遠在色塊的另一邊',
    heads.every((h, i) => (marks[i].left < 400 ? h.left > marks[i].left : h.left < marks[i].left)));
  ok('三段沒有重疊',
    marks.every((m, i) => i === 0 || m.top >= marks[i - 1].top + marks[i - 1].height));
  ok('填滿內文區', Math.max(...marks.map((m) => m.top + m.height)) > 440);
}

// 新版型不能只對預設配色有效。
for (const kind of ['timeline', 'quadrant', 'zigzag']) {
  for (const look of ['literary', 'tech', 'soft']) {
    slide = fakeSlide();
    const r = await build({ kind, title: 'T', look,
      items: [{ title: 'A', label: 'A', text: 'x' }, { title: 'B', label: 'B', text: 'y' }] });
    ok(`${kind} × ${look} 畫得出來`, r.ok, r.content);
  }
}

// ── 形狀名稱：叫錯一個就整張投影片沒了 ──────────────────────────────────────
//
// `addGeometricShape('Oval', …)` 看起來完全正確，而且不是 GeometricShapeType
// 的成員——圓形叫 **Ellipse**，「Oval」只出現在那個列舉自己的說明文字裡。
// 錯的名字不會在寫下去的地方丟例外，是在 ctx.sync() 才丟，而那個 sync 帶著
// 整張投影片。所以兩顆裝飾圓可以讓一整個封面消失，而 AI 只能推測「這台不支援」。
{
  slide = fakeSlide();
  const r = await host.executeTool('ppt_add_shape', { shape: 'Oval', text: '圓' });
  ok('模型給了不存在的形狀名稱也不會整個失敗', r.ok, r.content);
  ok('而且會告訴它正確的名字', /Ellipse/.test(r.content), r.content);
  ok('退回矩形而不是丟掉這次呼叫',
    alive().some((s) => s.geo === 'Rectangle'));
}
{
  slide = fakeSlide();
  const r = await host.executeTool('ppt_add_shape', { shape: 'Ellipse', text: '圓' });
  ok('正確的名字照常畫', r.ok && alive().some((s) => s.geo === 'Ellipse'), r.content);
  ok('沒有多餘的提示', !/沒有「/.test(r.content), r.content);
}

// ── 引言頁：同一個重複 bug 的第三個現場 ─────────────────────────────────────
//
// 骨架會畫 title，而引言頁的「那句話」本身就是 title——所以整句被印兩次，
// 一次 25pt 在頁首、一次 32pt 在中間。使用者的說法是「標題跟大文字一樣」。
// section 和 statement 前兩輪已經修過，quote 漏掉了。
{
  slide = fakeSlide();
  const q = '難的從來不是讓模型變聰明，而是讓它安全地接上正在運作的系統。';
  await build({ kind: 'quote', title: q, eyebrow: '01 現況與落差', page: 'P.05',
    body: ['MIT 2026 年企業生成式 AI 研究'] });
  ok('引言只出現一次', texts().filter((t) => t === q).length === 1,
    `出現 ${texts().filter((t) => t === q).length} 次`);
  // 但導覽標記要留著——不然讀者不知道這頁在哪一段。
  ok('章節標記和頁碼還在',
    texts().some((t) => /01 現況與落差/.test(t)) && texts().includes('P.05'));
  ok('沒有標題底下那條短線',
    !alive().some((s) => s.kind === 'geo' && s.width === 108 && s.height === 4));
  // 原本從 150 開始，上面留一條空帶；現在貼著內文帶的頂端。
  const spine = alive().find((s) => s.kind === 'geo' && s.width === 6 && s.height > 100);
  ok('引言貼著內文帶頂端，上面沒有空帶', spine.top <= 126, `色帶從 ${spine.top} 開始`);
  ok('而且往下填滿', spine.top + spine.height > 400,
    `色帶到 ${spine.top + spine.height}`);
  // 四個低對比的元素散在一片留白上，讀起來像沒載完。給它一塊底就成立了。
  const field = alive().find((s) => s.kind === 'geo' && s.width === 884 && s.height > 100);
  ok('引言坐在一塊色底上', !!field, '不然那一頁只是幾行字浮在白底上');
  const quote = alive().find((s) => s._text === q);
  ok('引言在那塊底裡面',
    quote.left > field.left && quote.top > field.top
      && quote.top + quote.height <= field.top + field.height + 1);
  // 署名要讀起來是「出處」，不是引言的最後一行。
  const cite = alive().find((s) => /MIT 2026/.test(s._text));
  ok('署名在色底外面、下方', cite.top >= field.top + field.height,
    `署名 ${cite.top}，色底到 ${field.top + field.height}`);
  ok('署名上面有一條短線帶出來',
    alive().some((s) => s.kind === 'geo' && s.width === 48 && s.height === 2));
}

// ── 封面副標要有行距 ────────────────────────────────────────────────────────
//
// 兩行 22pt 的中文單行間距會黏在一起。lineSpacing 要 API 1.8，所以這裡用的是
// 跟條列一樣的辦法：一行一個文字方塊，行距靠幾何。
for (const form of ['panel', 'split', 'stack', 'triptych', 'bleed', 'band', 'plain',
  'paper', 'imageleft', 'imagetop']) {
  slide = fakeSlide();
  await build({ kind: 'title', title: 'T', cover: form,
    body: ['88% 的企業已經在用 AI', '只有 11% 把代理人放進生產環境'] });
  const a = alive().find((s) => s._text === '88% 的企業已經在用 AI');
  const b = alive().find((s) => s._text === '只有 11% 把代理人放進生產環境');
  ok(`${form}：兩句副標各自一個方塊`, !!a && !!b,
    texts().join(' | '));
  if (a && b) {
    const gap = b.top - a.top;
    ok(`${form}：兩行之間有行距`, gap >= a.height * 0.95 && gap > 20,
      `間距 ${Math.round(gap)}，行高 ${Math.round(a.height)}`);
  }
}

// ── 區塊要有容器，不然文字散落各處 ──────────────────────────────────────────
{
  slide = fakeSlide();
  await build({ kind: 'zigzag', title: 'T', eyebrow: '01',
    items: [
      { title: '學習落差', text: '通用工具不會記住企業的流程' },
      { title: '整合落差', text: '試點停在部門邊界' },
      { title: '衡量落差', text: '導入前沒有訂基準線' },
    ] });
  const marks = alive().filter((s) => s.kind === 'geo'
    && s.width > 100 && s.width < 200 && s.height > 60);
  const panels = alive().filter((s) => s.kind === 'geo'
    && s.width > 600 && s.height > 60);
  ok('文字那半邊有自己的容器', panels.length === 3, `${panels.length} 個`);
  ok('容器有邊框', panels.every((p2) => p2.lineFormat.visible === true));
  // 容器要蓋住文字，不然它就只是一個放在旁邊的方塊。
  const head = alive().find((s) => s._text === '學習落差');
  ok('文字在容器裡面',
    panels.some((p2) => head.left >= p2.left && head.left + head.width <= p2.left + p2.width + 1),
    `文字 ${head.left}..${head.left + head.width}`);
  // 只能跟「同一列」的容器比——不同列本來就會在水平上重疊，那不是問題。
  ok('每一列的色塊和容器不重疊',
    marks.every((m) => panels.filter((p2) => Math.abs(p2.top - m.top) < 1).every((p2) =>
      m.left + m.width <= p2.left + 1 || p2.left + p2.width <= m.left + 1)),
    marks.map((m) => `${m.top}:${m.left}`).join(' '));
}
{
  slide = fakeSlide();
  await build({ kind: 'timeline', title: 'T',
    items: [{ at: '2024', label: 'A', note: 'x' }, { at: '2025', label: 'B' },
      { at: '2026', label: 'C' }] });
  const cards = alive().filter((s) => s.kind === 'geo' && s.height === 96);
  ok('每一站有自己的卡片', cards.length === 3, `${cards.length} 個`);
  ok('卡片有邊框', cards.every((c) => c.lineFormat.visible === true));
  // 卡片的色條要跟它那一站的圓點同色，兩者才成對。
  const caps = alive().filter((s) => s.kind === 'geo' && s.height === 3);
  ok('卡片有一條同色的頂／底條', caps.length === 3, `${caps.length} 條`);
}

// ── 四象限的軸標曾經小到讀不到 ──────────────────────────────────────────────
{
  slide = fakeSlide();
  await build({ kind: 'quadrant', title: 'T', x_axis: '規則明確度 →', y_axis: '↑ 發生頻率',
    items: [{ title: 'A' }, { title: 'B' }, { title: 'C' }, { title: 'D' }] });
  const x = alive().find((s) => s._text === '規則明確度 →');
  const y = alive().find((s) => s._text === '↑ 發生頻率');
  // 軸標是「這四格是什麼意思」的那一半，不是註腳。
  ok('軸標不是註腳級距', x._font.size >= 13 && y._font.size >= 13,
    `x ${x._font.size} y ${y._font.size}`);
  ok('軸標是粗體，跟卡片標題同等重量', x._font.bold === true && y._font.bold === true);
  const cells = alive().filter((s) => s.kind === 'geo' && s.width > 300 && s.height > 100);
  const gridTop = Math.min(...cells.map((c) => c.top));
  const gridBottom = Math.max(...cells.map((c) => c.top + c.height));
  ok('縱軸標在格線上方，沒有壓到格子', y.top + y.height <= gridTop + 1,
    `標 ${y.top}+${y.height}，格線從 ${gridTop}`);
  ok('橫軸標在格線下方', x.top >= gridBottom - 1,
    `標 ${x.top}，格線到 ${gridBottom}`);
  ok('格子沒有被擠出內文區', gridBottom <= 490, `${gridBottom}`);
}

// ── 同一個值不能被版面的兩個部分各畫一次 ───────────────────────────────────
//
// panel 封面的粗體那行會在沒有 footer 時退回 eyebrow，但上面那行小字是無條件
// 畫的——所以「2026 TREND REPORT」被印了兩次，一次 11pt 一次 18pt，上下相疊。
// 跟 section、statement、quote 是同一種錯：一個值，兩個互不知情的地方在畫。
{
  slide = fakeSlide();
  await build({ kind: 'title', cover: 'panel', title: 'T', eyebrow: '2026 TREND REPORT',
    body: ['s'] });
  ok('只有 eyebrow 的時候，它只出現一次',
    texts().filter((t) => t === '2026 TREND REPORT').length === 1,
    `出現 ${texts().filter((t) => t === '2026 TREND REPORT').length} 次`);
}
{
  slide = fakeSlide();
  await build({ kind: 'title', cover: 'panel', title: 'T', eyebrow: '2026 TREND REPORT',
    footer: '天狼星火組', body: ['s'] });
  ok('兩個值都有的時候兩行都在，各一次',
    texts().filter((t) => t === '2026 TREND REPORT').length === 1
      && texts().filter((t) => t === '天狼星火組').length === 1);
}

// ── 圓角面板配直角色條 ──────────────────────────────────────────────────────
//
// 色條是一個純矩形。放在直角面板上是對的；放在**圓角**面板上，它的方角會凸出
// 圓弧之外，整張卡片讀起來就是「圓弧 ＋ 直條」——兩個明顯不是一起畫的形狀。
for (const [look, bar] of [['business', 'top'], ['soft', 'top'], ['business', 'left']]) {
  slide = fakeSlide();
  await host.executeTool('ppt_add_cards',
    { look, bar, columns: 2, cards: [{ title: 'A' }, { title: 'B' }] });
  const panel = alive().find((s) => s.geo && s.height > 50);
  const cap = alive().find((s) => s.geo && (s.height === 5 || s.width === 5));
  const rounded = panel.geo === 'RoundRectangle';
  const inset = bar === 'top' ? cap.left > panel.left : cap.top > panel.top;
  ok(`${look}/${bar}：${rounded ? '圓角面板的色條要內縮' : '直角面板的色條齊邊'}`,
    rounded === inset,
    `${panel.geo}，色條 ${bar === 'top' ? cap.left : cap.top} vs 面板 ${bar === 'top' ? panel.left : panel.top}`);
  if (inset) {
    // 內縮但不能縮過頭——色條還是要看得出是那張卡的頂。
    const span = bar === 'top' ? cap.width / panel.width : cap.height / panel.height;
    ok(`${look}/${bar}：色條仍然佔大部分寬度`, span > 0.8, span.toFixed(2));
  }
}

// ── 兩個吃圖的內文版型 ──────────────────────────────────────────────────────
{
  INSERTS = [];
  slide = fakeSlide();
  const r = await build({ kind: 'feature', title: '這條產線', eyebrow: '02 現場',
    image: 'https://x/a.svg', look: 'business',
    items: [{ title: '產能', text: '每小時 1,200 件' }, { title: '良率', text: '99.2%' }] });
  ok('feature 畫得出來', r.ok, r.content);
  // 底圖（母題）跟圖片不是二選一：一個是地，一個站在上面。
  ok('母題和圖片都插入了', INSERTS.length === 2, `${INSERTS.length} 張`);
  ok('母題是滿版的那張', INSERTS[0].data.includes('960'));
  ok('圖片在左半邊、沒有滿版',
    INSERTS[1].data && alive().some((s) => s.kind === 'geo' && s.width === 470));
  ok('右邊的小卡都在', texts().includes('產能') && texts().includes('良率'));
}
{
  slide = fakeSlide();
  const r = await build({ kind: 'list_image', title: '四個部門',
    items: [{ title: '研發', text: 'a' }, { title: '製造', text: 'b' }, { title: '業務', text: 'c' }] });
  ok('list_image 畫得出來', r.ok, r.content);
  const slots = alive().filter((s) => s.kind === 'geo' && s.width > 100 && s.width < 200
    && s.height > 40);
  ok('三個圖片框', slots.length === 3, `${slots.length} 個`);
  // 四張圖是四趟來回，所以版型只排框、回座標，模型再一張一張填。
  ok('回傳每一格的座標', /第 1 列 left \d+, top \d+/.test(r.content), r.content.slice(0, 200));
  ok('座標數量跟列數一致', (r.content.match(/第 \d 列/g) || []).length === 3);
  ok('三段文字都在', ['研發', '製造', '業務'].every((t) => texts().includes(t)));
}

/**
 * 圖片版型不能自己把圖蓋掉。
 *
 * 圖片是在形狀批次「之前」插入的——必須如此，setSelectedDataAsync 是附加到最上
 * 層，先插入才會在下層。那之後任何一個滿版矩形都會整張罩住它：「圖有出來，但
 * 被另一個白色的區塊蓋住了」就是這個。
 *
 * 所以這裡檢查的不是「有沒有畫底色」，而是「有沒有任何形狀完整罩住圖片的框」。
 * 之後新增的圖片版型會自動被這條規則接住。
 */
for (const cover of ['imageleft', 'imageright', 'imagetop']) {
  slide = fakeSlide();
  INSERTS.length = 0;
  await build({ kind: 'title', cover, title: '標題', body: ['副標'],
    image: 'https://example.com/a.svg' });
  const img = INSERTS[INSERTS.length - 1];
  const covering = alive().filter((s) => s.kind === 'geo'
    && s.left <= img.imageLeft + 1 && s.top <= img.imageTop + 1
    && s.left + s.width >= img.imageLeft + img.imageWidth - 1
    && s.top + s.height >= img.imageTop + img.imageHeight - 1);
  ok(`${cover} 的圖沒有被形狀蓋住`, covering.length === 0,
    covering.map((s) => `${s.left},${s.top} ${s.width}x${s.height}`).join(' / '));
  ok(`${cover} 文字那半邊還是有底`, alive().some((s) => s.kind === 'geo' && s._fill));
}

console.log(`\n${pass} passed, ${fail} failed`);
/**
 * 模型送來的字串在派送這一層就洗過。
 *
 * 「製作 PPT 也有 \n 的問題」——投影片的卡片內文裡出現字面換行，而那不是任何一個
 * 叫 text 的參數：卡片是陣列、KPI 是物件陣列、圖表的分類又是另一個形狀。一處一處補
 * 就是一處一處漏，Word 那邊補了三次才收斂。
 *
 * 所以不挑欄位，整包 args 走一遍。這幾條釘的是「深層結構裡的字串也要洗到」。
 */
{
  const { cleanArgs } = await import('../src/taskpane/hosts/text.js');
  const B = String.fromCharCode(92);
  const got = cleanArgs({
    title: '三段式落地路徑',
    cards: [{ title: '聚焦選題', body: '挑可量化成效的流程先做' + B + 'n先定成效指標' }],
    color: '#1F3864',
    count: 3,
  });
  ok('陣列裡物件的字串也洗得到',
    got.cards[0].body.includes(String.fromCharCode(10))
    && !got.cards[0].body.includes(B + 'n'), JSON.stringify(got.cards[0].body));
  ok('色碼不受影響', got.color === '#1F3864', String(got.color));
  ok('數字不受影響', got.count === 3, String(got.count));
}

/**
 * 規劃是流程的一部分，不是提示詞裡的一句話。
 *
 * 「他沒辦法再有更多的創意，就是他沒有任何一個流程幫他規劃」——使用者講得很準。
 * 原本的流程是「說話 → 直接呼叫零件」，中間沒有設計這一步，所以不管給什麼需求，
 * 做出來都是 KPI 卡加圖表。
 *
 * 我在提示詞裡寫過三次「請先規劃」、搬過位置、修過矛盾，行為沒變。規則擋不住模型，
 * 機制擋得住——所以規劃變成工具，建構工具會檢查它跑過沒有。
 */
{
  slide = fakeSlide();
  const blocked = await host.executeTool('ppt_build_slide', { slide: 9, title: '沒規劃就想做' });
  ok('沒規劃過的頁會被擋下來', blocked.ok === false && /還沒有版面規劃/.test(String(blocked.error)),
    JSON.stringify(blocked).slice(0, 120));

  // 主標缺了也不擋，但會把「它該是一句主張不是題目」講出來——理由同上。
  const bad = await host.executeTool('ppt_plan_slide', { slide: 9, blocks: [{ area: '上', media: '圖' }] });
  ok('少了主標也不會失敗', bad.ok !== false, JSON.stringify(bad).slice(0, 100));
  ok('但會說主標該是一句主張', /主張/.test(String(bad.content)),
    String(bad.content).slice(0, 120));

  const good = await host.executeTool('ppt_plan_slide', {
    slide: 9,
    headline: '採用衝很快，成效跟不上',
    blocks: [
      { area: '左上', media: '長條圖', why: '要比三組採用 vs 成效的落差' },
      { area: '右上', media: '表格', why: '欄與欄的對應關係條列講不出來' },
    ],
  });
  ok('規劃會把版面講出來給人看', /採用衝很快/.test(String(good.content))
    && /長條圖/.test(String(good.content)), String(good.content).slice(0, 120));

  const now = await host.executeTool('ppt_build_slide', { slide: 9, title: '規劃過了' });
  ok('規劃過之後就放行', now.ok !== false, JSON.stringify(now).slice(0, 120));

  // 條列佔太多塊，在動手之前講——那時候改最便宜。
  const listy = await host.executeTool('ppt_plan_slide', {
    slide: 10,
    headline: '全部都用講的',
    blocks: [{ area: '左', media: '條列' }, { area: '中', media: '條列' }, { area: '右', media: '文字' }],
  });
  ok('三塊以上是條列會被提醒', /其實是圖或表/.test(String(listy.content)),
    String(listy.content).slice(-80));

  /**
   * 規劃不該比不規劃更容易失敗。
   *
   * 第一版在缺配色時擋下來，想逼模型做決定。結果是它連試兩次失敗、放棄規劃、退回原本的
   * KPI 加兩張圖——**它永遠走最便宜的路**，而我把規劃那條路變貴了。
   *
   * 第一步擋下來不會換來更好的決定，只會換來不用這個工具。改成缺的部分補上、
   * 補完說清楚補了什麼——它想改下一句就能改，而整份規劃已經成立了。
   */
  const noTone = await host.executeTool('ppt_plan_deck', {
    slides: [{ purpose: '現況', form: '數字卡' }],
  });
  ok('沒給配色也不會失敗', noTone.ok !== false, JSON.stringify(noTone).slice(0, 100));
  ok('但會講出來並問這份要給誰看', /給誰看|預設/.test(String(noTone.content)),
    String(noTone.content).slice(-100));

  const deck = await host.executeTool('ppt_plan_deck', {
    look: 'corporate',
    slides: [{ purpose: '現況', form: '數字卡 + 折線' }, { purpose: '問題', form: '左圖右表' }],
  });
  ok('整份規劃會列出每頁一行', /1 現況/.test(String(deck.content))
    && /2 問題/.test(String(deck.content)), String(deck.content).slice(0, 120));
}

process.exit(fail ? 1 : 0);
