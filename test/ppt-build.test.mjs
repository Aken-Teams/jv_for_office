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

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
