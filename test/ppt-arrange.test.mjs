/**
 * Pin the PowerPoint host's geometry.
 *
 * ── Why this file exists and what it guards ──
 * `ppt_arrange` is the tool that justifies the add-in: a web app can generate a
 * deck, but it cannot open the file someone already has, read where the six
 * boxes on slide 4 actually sit, and line them up. That only works if the
 * arithmetic is right — and it is arithmetic that fails silently. A distribute
 * that is off by one gap does not throw; it produces a slide that still looks
 * almost aligned, which is worse than one that obviously is not.
 *
 * PowerPoint also gives no safety net to catch it later. There are no tracked
 * changes, and the pane's undo only deletes shapes it created — it cannot put a
 * moved box back. So a bad move is permanent from the add-in's side, and these
 * assertions are the only thing standing in front of that.
 *
 * The fixture is a deliberately awkward row: different widths, out of order, and
 * not evenly spaced. Equal-sized boxes already in order would pass almost any
 * implementation, including several wrong ones.
 */

// ── A PowerPoint that records what was set ───────────────────────────────────

function shape(id, left, top, width, height) {
  const s = { id, name: `shape-${id}`, type: 'GeometricShape' };
  for (const [k, v] of [['left', left], ['top', top], ['width', width], ['height', height]]) {
    let cur = v;
    Object.defineProperty(s, k, { get: () => cur, set: (n) => { cur = n; }, enumerable: true });
  }
  return s;
}

let SHAPES = [];
const ctx = {
  presentation: {
    slides: {
      items: [],
      load() {},
      getItemAt: (i) => ctx.presentation.slides.items[i],
    },
  },
  sync: async () => {},
  load: () => {},
};
const slide = { id: 's1', shapes: { items: [], load() {} } };
ctx.presentation.slides.items = [slide];

globalThis.PowerPoint = { run: async (fn) => fn(ctx) };
globalThis.Office = {
  context: {
    requirements: { isSetSupported: () => true },
    document: { url: 'C:\\decks\\test.pptx', addHandlerAsync: () => {} },
  },
  EventType: { DocumentSelectionChanged: 'sel' },
  CoercionType: { Image: 'image', XmlSvg: 'xmlsvg' },
  AsyncResultStatus: { Succeeded: 'succeeded' },
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

/** Uneven on purpose: different widths, out of order, unequal gaps. */
const fixture = () => {
  SHAPES = [
    shape('a', 100, 200, 120, 60),
    shape('b', 400, 150, 60, 90),
    shape('c', 250, 300, 200, 40),
  ];
  slide.shapes.items = SHAPES;
  return SHAPES;
};
const at = (i) => {
  const s = SHAPES[i];
  return [Math.round(s.left), Math.round(s.top), Math.round(s.width), Math.round(s.height)];
};
const run = (args) => host.executeTool('ppt_arrange', args);

// ── Alignment ────────────────────────────────────────────────────────────────

fixture();
await run({ how: 'left' });
ok('靠左對齊：都貼到最小的 left', SHAPES.every((s) => s.left === 100),
  SHAPES.map((s) => s.left).join(','));
ok('靠左對齊不會動到 top', at(1)[1] === 150);
ok('靠左對齊不會改變寬度', at(2)[2] === 200);

fixture();
await run({ how: 'right' });
// The right edge of the fixture is c: 250 + 200 = 450. b is 400 + 60 = 460 — so
// 460 is the edge, and every shape's LEFT has to be edge minus its own width.
ok('靠右對齊：右緣切齊，各自扣掉自己的寬度',
  SHAPES.map((s) => s.left + s.width).every((r) => r === 460),
  SHAPES.map((s) => `${s.left}+${s.width}`).join(' '));

fixture();
await run({ how: 'top' });
ok('靠上對齊', SHAPES.every((s) => s.top === 150), SHAPES.map((s) => s.top).join(','));

fixture();
await run({ how: 'bottom' });
// Bottoms: 260, 240, 340 → 340.
ok('靠下對齊', SHAPES.every((s) => s.top + s.height === 340),
  SHAPES.map((s) => s.top + s.height).join(','));

fixture();
await run({ how: 'center_h' });
// Span 100..460, mid = 280. Each centred on 280.
ok('水平置中對齊：每個都以中線為準',
  SHAPES.every((s) => Math.round(s.left + s.width / 2) === 280),
  SHAPES.map((s) => Math.round(s.left + s.width / 2)).join(','));

fixture();
await run({ how: 'center_v' });
// Span 150..340, mid = 245.
ok('垂直置中對齊',
  SHAPES.every((s) => Math.round(s.top + s.height / 2) === 245),
  SHAPES.map((s) => Math.round(s.top + s.height / 2)).join(','));

// ── Distribute — the one that is easy to get subtly wrong ────────────────────

fixture();
await run({ how: 'distribute_h' });
// Sorted by current left: a(100,w120), c(250,w200), b(400,w60).
// Span 100..460 = 360. Widths total 380 — WIDER than the span, so the gap is
// negative and they overlap. That is correct behaviour, not a bug: PowerPoint
// does the same. What must hold is that the ends stay put and the order is kept.
const lefts = SHAPES.map((s) => s.left);
ok('水平分散：最左邊的不動', Math.round(SHAPES[0].left) === 100, String(SHAPES[0].left));
ok('水平分散：最右邊的仍然切齊右緣',
  Math.round(SHAPES[1].left + SHAPES[1].width) === 460,
  `${SHAPES[1].left}+${SHAPES[1].width}`);
ok('水平分散：不會把順序打亂',
  lefts[0] < lefts[2] && lefts[2] < lefts[1],
  `a=${lefts[0]} c=${lefts[2]} b=${lefts[1]}　分散不該重新排序`);

// A case with room, where the gaps must actually come out equal.
SHAPES = [shape('a', 0, 0, 50, 10), shape('b', 90, 0, 50, 10), shape('c', 300, 0, 50, 10)];
slide.shapes.items = SHAPES;
await run({ how: 'distribute_h' });
const gaps = [SHAPES[1].left - (SHAPES[0].left + 50), SHAPES[2].left - (SHAPES[1].left + 50)];
ok('水平分散：間距真的相等', Math.abs(gaps[0] - gaps[1]) < 0.001, gaps.join(' vs '));
ok('水平分散：兩端不動', SHAPES[0].left === 0 && SHAPES[2].left === 300,
  `${SHAPES[0].left} / ${SHAPES[2].left}`);

fixture();
await run({ how: 'distribute_v' });
ok('垂直分散：最上面的不動', Math.round(SHAPES[1].top) === 150, String(SHAPES[1].top));
ok('垂直分散：最下面的仍然切齊下緣',
  Math.round(SHAPES[2].top + SHAPES[2].height) === 340,
  `${SHAPES[2].top}+${SHAPES[2].height}`);

// ── Sizing ───────────────────────────────────────────────────────────────────

fixture();
await run({ how: 'same_width' });
ok('統一寬度：用現有最大的', SHAPES.every((s) => s.width === 200),
  SHAPES.map((s) => s.width).join(','));
ok('統一寬度不會動到高度', at(0)[3] === 60);

fixture();
await run({ how: 'same_width', value: 111 });
ok('統一寬度：可以指定數值', SHAPES.every((s) => s.width === 111));

fixture();
await run({ how: 'same_size' });
ok('統一大小：寬和高一起', SHAPES.every((s) => s.width === 200 && s.height === 90),
  SHAPES.map((s) => `${s.width}x${s.height}`).join(' '));

// ── Selecting which shapes ───────────────────────────────────────────────────

fixture();
await run({ how: 'left', shapes: [1, 3] });
ok('只動指定的物件', SHAPES[0].left === 100 && SHAPES[2].left === 100 && SHAPES[1].left === 400,
  SHAPES.map((s) => s.left).join(','));

// ── Refusals ─────────────────────────────────────────────────────────────────

const refuse = async (args) => {
  const r = await host.executeTool('ppt_arrange', args);
  return r.ok ? '' : String(r.content || '');
};
fixture();
ok('未知的 how 會被擋掉並列出可用的值',
  (await refuse({ how: 'diagonal' })).includes('distribute_h'));
slide.shapes.items = [SHAPES[0]];
ok('只有一個物件會被擋掉', (await refuse({ how: 'left' })).includes('兩個'));

// ── The contract the shell depends on ────────────────────────────────────────

slide.shapes.items = fixture();
ok('meta 用投影片當座標系統', host.meta.citation.pattern.test('[[投影片 3]]'));
ok('meta 宣告了畫布尺寸', host.meta.canvas.width === 960 && host.meta.canvas.height === 540);
for (const fn of ['executeTool', 'supportedTools', 'restoreSnapshot', 'goTo',
  'getDocumentName', 'getDocumentKey', 'getLocation', 'watchLocation',
  'getSelectionRef', 'watchSelection', 'inspectWrite', 'insertImage']) {
  ok(`主機契約：${fn}`, typeof host[fn] === 'function');
}

// ── 刪除要一次刪完 ───────────────────────────────────────────────────────────
//
// 每刪一個物件就是一張確認卡。清空一張投影片重做＝八張卡、八次「允許」——
// 按到第三次就沒人在看內容了，那個確認門等於白做。一次決定，一張卡。
{
  const fresh = (n) => {
    slide.shapes.items = Array.from({ length: n }, (_, i) => ({
      id: `d${i + 1}`, tag: `#${i + 1}`, gone: false, delete() { this.gone = true; },
    }));
  };
  const left = () => slide.shapes.items.filter((x) => !x.gone).map((x) => x.tag).join(' ');

  fresh(8);
  let r = await host.executeTool('ppt_delete_shape', { shapes: [3, 5, 8] });
  ok('一次刪多個', r.ok && left() === '#1 #2 #4 #6 #7', left());
  // 這是最容易寫錯的地方：由小到大刪的話，刪掉 3 之後原本的 5 已經變成 4。
  ok('編號是讀取當下的編號，不會因為前面刪掉而位移', left() === '#1 #2 #4 #6 #7', left());
  ok('回覆講清楚刪了幾個', /3 個物件/.test(r.content), r.content);

  fresh(4);
  r = await host.executeTool('ppt_delete_shape', { all: true });
  ok('all: true 會清空整張', r.ok && left() === '', left());

  fresh(3);
  r = await host.executeTool('ppt_delete_shape', { shape: 2 });
  ok('單一編號還是能用', r.ok && left() === '#1 #3', left());

  fresh(3);
  r = await host.executeTool('ppt_delete_shape', { shapes: [2, 2, 2] });
  ok('重複的編號只算一次', r.ok && left() === '#1 #3', left());

  // 刪除救不回來，所以「刪了兩個才發現第三個編號是錯的」是不能接受的結果。
  fresh(3);
  r = await host.executeTool('ppt_delete_shape', { shapes: [1, 2, 9] });
  ok('有一個編號錯就整批不動', !r.ok && left() === '#1 #2 #3', left());
  ok('而且說清楚是哪個錯、該怎麼辦',
    /9/.test(r.content) && /ppt_read_slide/.test(r.content), r.content);

  fresh(3);
  r = await host.executeTool('ppt_delete_shape', {});
  ok('什麼都沒給會被擋下來並列出可用的參數',
    !r.ok && /shapes/.test(r.content) && /all/.test(r.content), r.content);
  slide.shapes.items = [];
}

const tools = host.supportedTools();
ok('工具清單不是空的', tools.length > 10, String(tools.length));
// ask_user is excluded on purpose, and the Word host omits it for the same
// reason: the backend's describeToolCall returns the QUESTION, and app.js prefers
// that summary over any label. 「這份簡報是給客戶還是主管看？」 is a better status
// line than 「等你回答…」, so a generic caption here would be a downgrade.
const needsLabel = tools.filter((n) => n !== 'ppt_run_script' && n !== host.meta.askUserTool);
ok('每個工具都有 i18n 進度字串',
  needsLabel.every((n) => host.meta.toolLabels()[n]),
  needsLabel.filter((n) => !host.meta.toolLabels()[n]).join(', '));

// PowerPoint has no reorder API, so there must be no tool implying otherwise.
ok('沒有假裝可以搬移投影片',
  !tools.some((n) => /move|reorder|sort/i.test(n)),
  'API 沒有這個能力，提供做不到的工具比不提供更糟');

// ── Undo is by inverse, and must say what it cannot do ───────────────────────

const empty = await host.restoreSnapshot(null);
ok('沒有快照時的復原是安全的', empty.kind === 'none');
const note = (await host.restoreSnapshot({ kind: 'created', shapes: [] })).note || '';
ok('復原說明不會過度承諾', !/回到原狀|完整還原/.test(note), note);

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
