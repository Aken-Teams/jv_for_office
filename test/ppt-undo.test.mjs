/**
 * PPT 的復原——這個增益集敢不敢碰別人既有簡報的那條底線。
 *
 * ── 為什麼這件事值得一個獨立的測試檔 ──
 * 增益集的定位是「在已經存在的檔案上工作」。那復原就不是加分項，是前提：一個會
 * 改你的簡報又退不回去的工具，使用者會學會不要讓它靠近既有檔案，然後我們做的所
 * 有版面設計就只能用在新檔案上。
 *
 * 原本的做法只記「這次新增了哪些物件」，所以改寫第 5 頁的文字、刪掉一張投影片都
 * 退不回來。現在改成逆操作清單。
 *
 * ── 這裡要釘住的三件事 ──
 *   1. 改之前要先讀。沒讀就沒得還原，而且錯過的時機補不回來。
 *   2. 「讀不到」和「本來就沒有」是兩回事。透明的形狀被還原成白色，是比沒還原
 *      更糟的結果——它看起來成功了。
 *   3. 還原不了的要講出來。圖片讀不回來、重建的投影片會接在最後面、重建的形狀
 *      一律是矩形——這些都必須出現在回報裡，不能默默還原一半。
 */
let pass = 0, fail = 0;
const ok = (name, cond, detail) => {
  cond ? pass++ : fail++;
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}${cond ? '' : '  → ' + (detail || '')}`);
};

let SEQ = 0;

/**
 * 一個會像 Office 那樣回答問題的假形狀。
 *
 * 關鍵是 fill 和 font 要「讀得回來」——真正的 Office 會在 load 之後讓你讀
 * `fill.type` 和 `fill.foregroundColor`，而復原的整個前提就是這些讀得到。舊的假
 * 物件只模擬了「設定」沒有模擬「讀取」，所以就算捕捉完全壞掉，測試也會通過。
 */
function fakeShape(opts = {}) {
  const s = {
    id: opts.id || `s${++SEQ}`,
    name: opts.name || '',
    type: opts.type || 'GeometricShape',
    deleted: false,
    left: opts.left, top: opts.top, width: opts.width, height: opts.height,
    _text: opts.text === undefined ? '' : opts.text,
    // `load` 不是裝飾。真正的 Office 每一層都是要先 load 才讀得到的 proxy，
    // 而捕捉那段是包在 try/catch 裡的——假物件少一個 load，整段捕捉就被靜靜吃
    // 掉，然後測試會全部通過而復原其實壞的。
    _font: {
      name: opts.font, size: opts.size, color: opts.color, bold: opts.bold,
      italic: opts.italic, load() {},
    },
    fill: {
      type: opts.fill === undefined ? 'NoFill' : 'Solid',
      foregroundColor: opts.fill,
      setSolidColor(c) { s.fill.type = 'Solid'; s.fill.foregroundColor = c; },
      clear() { s.fill.type = 'NoFill'; s.fill.foregroundColor = undefined; },
      load() {},
    },
    lineFormat: {
      visible: opts.lineVisible, color: opts.lineColor, weight: opts.lineWeight,
      load() {},
    },
    load() {},
    delete() { s.deleted = true; },
  };
  Object.defineProperty(s, 'textFrame', {
    get() {
      // 圖片和線條沒有文字框，而且問了會丟例外——這正是會把整批 sync 毒死的那
      // 種錯，捕捉的程式必須自己擋住。
      if (opts.noTextFrame) throw new Error('InvalidArgument: no text frame');
      return {
        hasText: !!s._text,
        textRange: {
          get text() { return s._text; },
          set text(v) { s._text = v; },
          font: s._font,
          paragraphFormat: {},
          load() {},
        },
        load() {},
        verticalAlignment: undefined, wordWrap: undefined, autoSizeSetting: undefined,
        leftMargin: 0, rightMargin: 0, topMargin: 0, bottomMargin: 0,
      };
    },
  });
  return s;
}

const REAL_GEO = new Set(['Rectangle', 'RoundRectangle', 'Ellipse', 'Chevron',
  'RightArrow', 'Triangle', 'Diamond', 'Donut', 'Frame', 'Plus', 'Star5',
  'Hexagon', 'Pentagon', 'Trapezoid', 'HomePlate', 'Wave', 'Arc', 'Pie',
  'Teardrop', 'Cloud', 'Heart', 'Moon']);

function fakeSlide(id, initial = []) {
  const items = [...initial];
  const self = {
    id,
    _all: items,
    // 投影片自己會從 deck 裡消失，跟 Office 一樣——所以刪除後面的編號會位移，
    // 而那正是刪除必須由後往前做的原因。
    delete() {
      const at = deck.indexOf(self);
      if (at >= 0) deck.splice(at, 1);
    },
    shapes: {
      get items() { return items.filter((s) => !s.deleted); },
      load() {},
      addGeometricShape(g, o) {
        if (!REAL_GEO.has(g)) throw new Error(`InvalidArgument: 沒有 "${g}"`);
        if (o.left < 0 || o.top < 0) throw new Error('InvalidArgument: 負座標');
        const s = fakeShape({ ...o, type: 'GeometricShape' });
        items.push(s);
        return s;
      },
      addTextBox(t, o) {
        const s = fakeShape({ ...o, text: t, type: 'TextBox' });
        items.push(s);
        return s;
      },
      addLine(c, o) { const s = fakeShape({ ...o, type: 'Line', noTextFrame: true }); items.push(s); return s; },
      getItemAt(i) { return items.filter((s) => !s.deleted)[i]; },
      getItem(id2) {
        const hit = items.find((s) => s.id === id2 && !s.deleted);
        if (!hit) throw new Error(`ItemNotFound: ${id2}`);
        return hit;
      },
    },
  };
  return self;
}

let deck = [];
let added = 0;
const ctx = {
  presentation: {
    slides: {
      get items() { return deck; },
      load() {},
      add() { deck.push(fakeSlide(`added-${++added}`)); },
      getItemAt(i) { return deck[i]; },
    },
    setSelectedSlides() {},
  },
  sync: async () => {},
  load: () => {},
};

globalThis.PowerPoint = { run: async (fn) => fn(ctx) };
globalThis.Office = {
  context: {
    requirements: { isSetSupported: () => true },
    document: {
      url: 'C:\\d.pptx',
      addHandlerAsync: () => {},
      setSelectedDataAsync: (d, o, cb) => cb({ status: 'ok' }),
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
const run = (tool, args) => host.executeTool(tool, args);
const undo = (snap) => host.restoreSnapshot(snap);
const shapesOf = (i) => deck[i].shapes.items;

// ── 改寫文字 ──────────────────────────────────────────────────────────────────
{
  const original = fakeShape({ id: 'a1', text: '原本的標題', left: 10, top: 20, width: 300, height: 40, fill: '#112233' });
  deck = [fakeSlide('s1', [original])];
  const r = await run('ppt_set_text', { slide: 1, shape: 1, text: '被 AI 改掉的' });
  ok('改寫成功', r.ok, r.content);
  ok('文字真的被改了', original._text === '被 AI 改掉的');
  ok('有留下快照', !!r.snapshot, JSON.stringify(r.snapshot));

  const back = await undo(r.snapshot);
  ok('文字退回原本的內容', original._text === '原本的標題', original._text);
  ok('回報說還原了物件', /還原 1 個/.test(back.note), back.note);
}

// ── 調整物件：位置與填色 ──────────────────────────────────────────────────────
{
  const box = fakeShape({ id: 'b1', text: 'x', left: 50, top: 60, width: 200, height: 100, fill: '#AABBCC' });
  deck = [fakeSlide('s1', [box])];
  const r = await run('ppt_format_shape', { slide: 1, shape: 1, left: 400, top: 300, fill: '#FF0000' });
  ok('位置和填色都被改了', box.left === 400 && box.fill.foregroundColor === '#FF0000');

  await undo(r.snapshot);
  ok('位置退回去', box.left === 50 && box.top === 60, `${box.left},${box.top}`);
  ok('填色退回去', box.fill.foregroundColor === '#AABBCC', box.fill.foregroundColor);
  ok('寬高沒有被動到', box.width === 200 && box.height === 100);
}

/**
 * 透明必須還原成透明。
 *
 * 「讀不到填色」和「本來就沒有填色」在程式裡都很容易變成同一個 undefined，然後
 * 復原就把一個透明的形狀塗成白色。那比沒有復原更糟，因為它看起來成功了。
 */
{
  const clear = fakeShape({ id: 'c1', text: 'x', left: 0, top: 0, width: 100, height: 50 });
  deck = [fakeSlide('s1', [clear])];
  ok('起始狀態是沒有填色', clear.fill.type === 'NoFill');
  const r = await run('ppt_format_shape', { slide: 1, shape: 1, fill: '#123456' });
  ok('被填上顏色', clear.fill.foregroundColor === '#123456');
  await undo(r.snapshot);
  ok('退回成透明，不是白色', clear.fill.type === 'NoFill' && !clear.fill.foregroundColor,
    `${clear.fill.type} ${clear.fill.foregroundColor}`);
}

// ── 排列 ──────────────────────────────────────────────────────────────────────
{
  const a = fakeShape({ id: 'p1', text: 'a', left: 10, top: 10, width: 100, height: 50 });
  const b = fakeShape({ id: 'p2', text: 'b', left: 300, top: 80, width: 100, height: 50 });
  const c = fakeShape({ id: 'p3', text: 'c', left: 600, top: 200, width: 100, height: 50 });
  deck = [fakeSlide('s1', [a, b, c])];
  const r = await run('ppt_arrange', { slide: 1, how: 'left' });
  ok('三個都靠左了', a.left === 10 && b.left === 10 && c.left === 10);

  await undo(r.snapshot);
  ok('三個位置全部退回', a.left === 10 && b.left === 300 && c.left === 600,
    `${a.left} ${b.left} ${c.left}`);
  ok('上下位置沒被誤動', b.top === 80 && c.top === 200);
}

// ── 刪除物件 ──────────────────────────────────────────────────────────────────
{
  const keep = fakeShape({ id: 'k1', text: '留著' });
  const doomed = fakeShape({
    id: 'd1', name: 'jv:card:x', text: '被刪掉的卡片',
    left: 40, top: 120, width: 260, height: 90, fill: '#0F6E72',
  });
  deck = [fakeSlide('s1', [keep, doomed])];
  const r = await run('ppt_delete_shape', { slide: 1, shape: 2 });
  ok('真的刪掉了', doomed.deleted === true);
  ok('另一個沒被動到', keep.deleted === false);

  const back = await undo(r.snapshot);
  const rebuilt = shapesOf(0).find((s) => s._text === '被刪掉的卡片');
  ok('被刪掉的物件重建回來了', !!rebuilt, shapesOf(0).map((s) => s._text).join('/'));
  ok('位置一起回來', rebuilt && rebuilt.left === 40 && rebuilt.top === 120);
  ok('填色一起回來', rebuilt && rebuilt.fill.foregroundColor === '#0F6E72');
  ok('名字一起回來', rebuilt && rebuilt.name === 'jv:card:x');
  // 誠實：重建不是原件。PowerPoint 讀不回原本是圓形還是箭頭。
  ok('有講重建的形狀一律是矩形', /矩形/.test(back.note), back.note);
}

/**
 * 圖片刪掉就是拿不回來，而且必須說。
 *
 * 讀不回圖片的位元組，所以重建不了。沉默地還原五個、漏掉第六個，比一開始就說
 * 「有一個拿不回來」糟——前者讓人以為復原成功了。
 */
{
  const pic = fakeShape({ id: 'pic1', type: 'Image', noTextFrame: true, left: 0, top: 0, width: 400, height: 300 });
  const txt = fakeShape({ id: 't1', text: '說明文字', left: 10, top: 400, width: 200, height: 40 });
  deck = [fakeSlide('s1', [pic, txt])];
  const r = await run('ppt_delete_shape', { slide: 1, all: true });
  ok('刪除當下就說了圖片拿不回來', /拿不回來/.test(r.content), r.content);

  const back = await undo(r.snapshot);
  ok('文字那個重建回來了', shapesOf(0).some((s) => s._text === '說明文字'));
  ok('圖片沒有被假裝還原', !shapesOf(0).some((s) => s.width === 400 && s.height === 300));
  ok('回報有講幾個還不了', /沒辦法還原/.test(back.note), back.note);
}

// ── 刪除投影片 ────────────────────────────────────────────────────────────────
{
  deck = [
    fakeSlide('s1', [fakeShape({ id: 'x1', text: '第一張' })]),
    fakeSlide('s2', [
      fakeShape({ id: 'x2', text: '要被刪的標題', left: 60, top: 40, width: 500, height: 60, fill: '#333333' }),
      fakeShape({ id: 'x3', text: '要被刪的內文', left: 60, top: 120, width: 500, height: 200 }),
    ]),
    fakeSlide('s3', [fakeShape({ id: 'x4', text: '第三張' })]),
  ];
  const r = await run('ppt_delete_slide', { slide: 2 });
  ok('投影片少了一張', deck.length === 2, `${deck.length}`);
  // 舊版這裡寫的是「這個動作復原鍵拿不回來」。現在拿得回來，但要說清楚拿回來的
  // 是什麼。
  ok('不再宣稱拿不回來', !/復原鍵拿不回來/.test(r.content), r.content);
  ok('有說會接在最後面', /最後面/.test(r.content), r.content);

  const back = await undo(r.snapshot);
  ok('投影片重建了', deck.length === 3, `${deck.length}`);
  const last = deck[deck.length - 1].shapes.items;
  ok('兩個物件都回來了', last.length === 2, `${last.length}`);
  ok('文字回來了', last.some((s) => s._text === '要被刪的標題'));
  ok('位置回來了', last.some((s) => s.left === 60 && s.top === 40));
  ok('回報有講它接在最後面', /最後面/.test(back.note), back.note);
}

// ── 新增：原本就會的事不能壞掉 ────────────────────────────────────────────────
{
  deck = [fakeSlide('s1', [fakeShape({ id: 'orig', text: '使用者自己打的' })])];
  const r = await run('ppt_add_text', { slide: 1, text: 'AI 加的', left: 10, top: 10, width: 200, height: 40 });
  ok('加上去了', shapesOf(0).some((s) => s._text === 'AI 加的'));
  const back = await undo(r.snapshot);
  ok('新增的被移除', !shapesOf(0).some((s) => s._text === 'AI 加的'));
  ok('使用者原本的還在', shapesOf(0).some((s) => s._text === '使用者自己打的'));
  ok('回報說移除了新增的', /移除新增的/.test(back.note), back.note);
}

// ── 復原本身不能壞掉 ──────────────────────────────────────────────────────────
{
  const s = fakeShape({ id: 'z1', text: '原文', left: 0, top: 0, width: 100, height: 50 });
  deck = [fakeSlide('s1', [s])];
  const r = await run('ppt_set_text', { slide: 1, shape: 1, text: '改過' });
  // 使用者自己把它刪了，然後才按復原。快照描述的東西已經不在了。
  s.deleted = true;
  let threw = null;
  const back = await undo(r.snapshot).catch((e) => { threw = e; return null; });
  ok('物件已經不在也不會丟例外', !threw, threw && threw.message);
  ok('而且說得出沒有東西需要退回', back && /沒有東西需要退回|還原 0|已/.test(back.note), back && back.note);
}

{
  const back = await undo(null);
  ok('沒有快照時回一句話，不是例外', back.kind === 'none' && !!back.note);
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
