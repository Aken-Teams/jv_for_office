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
      addGeometricShape: (g, o) => { const s = fakeShape('geo', o); s.geo = g; items.push(s); return s; },
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
  },
  sync: async () => {},
  load: () => {},
};

globalThis.PowerPoint = { run: async (fn) => fn(ctx) };
globalThis.Office = {
  context: {
    requirements: { isSetSupported: () => true },
    document: { url: 'C:\\d.pptx', addHandlerAsync: () => {} },
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
const build = (args) => host.executeTool('ppt_build_slide', args);
/** CHROME.contentTop in the host — where the body band starts. */
const CHROME_CONTENT_TOP = 124;
const alive = () => slide.shapes.items;
const texts = () => alive().filter((s) => s._text).map((s) => s._text);

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

for (const [name, dark] of [['corporate', false], ['night', true]]) {
  slide = fakeSlide();
  await build({ kind: 'title', title: '2026 AI 趨勢', body: ['副標'], theme: name });
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
  const ground = full.find((s) => s.height === 540);
  // 純白的投影片沒有「頁面」，所以白色卡片也就沒有邊——整頁看起來像直接打字上去的。
  ok('內容頁有淡色底，不是純白', !!ground && ground._fill !== '#FFFFFF',
    ground ? ground._fill : '沒有底');
  ok('底色是第一個畫的，蓋不到別人',
    slide._all.indexOf(ground) === 0, `第 ${slide._all.indexOf(ground)} 個`);
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
  const v = await shot({ kind: 'bullets', title: 'T', eyebrow: '01 章節', page: 'P.02',
    body: ['一'], look: 'tech' });
  const ground = v.geos.find((s) => s.width === 960 && s.height === 540);
  const lum = (c) => { const x = String(c).replace('#', '');
    return (parseInt(x.slice(0, 2), 16) + parseInt(x.slice(2, 4), 16) + parseInt(x.slice(4, 6), 16)) / 765; };
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
    const v = await shot({ kind: 'bullets', title: 'T', body: ['一'], look: name });
    return (v.geos.find((s) => s.width === 960 && s.height === 540) || {})._fill;
  };
  const business = await ink('business');
  ok('文青不會用到企業風的底色', await ink('literary') !== business);
  ok('柔和也不會', await ink('soft') !== business);
  // 但兩個知識都要成立：明講 theme 的時候，使用者說了算。
  slide = fakeSlide();
  await build({ kind: 'bullets', title: 'T', body: ['一'], look: 'literary', theme: 'corporate' });
  ok('明確指定 theme 的時候使用者說了算',
    alive().find((s) => s.kind === 'geo' && s.width === 960 && s.height === 540)._fill
      === business);
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

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
