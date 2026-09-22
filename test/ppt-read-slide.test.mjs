/**
 * Reading a slide — specifically, reading it when there is a picture on it.
 *
 * ── The bug ──
 * `loadTexts` queued a textFrame load for every shape and ran one `ctx.sync()`.
 * The per-shape `try/catch` around `s.textFrame` looked like it handled the
 * shapes that cannot hold text, and did not: reading the property only QUEUES
 * work. The throw happens later, inside the sync, and one rejected sync rejects
 * the whole batch. The catch then returned '' for every shape on the slide.
 *
 * What the person saw was 「那頁的文案我讀不到 —— 工具只看得到物件的位置、看不到
 * 裡面的字」, followed by the agent placing a picture by guesswork and landing it
 * on top of the words. Nothing errored. The slide simply read as empty.
 *
 * ── Why it is worth a file of its own ──
 * It needed a picture on the slide to fire, and three things we added recently
 * put one there: the texture layer, ppt_add_photo, and generated backgrounds.
 * The feature work made the blindness the normal case.
 */

let SYNCS = 0;
/** Shapes whose queued load will blow up the batch, the way a picture's does. */
function fakeShape(opts) {
  const s = {
    id: opts.id, name: opts.name || '', type: opts.type || 'GeometricShape',
    left: opts.left ?? 0, top: opts.top ?? 0, width: opts.width ?? 100, height: opts.height ?? 20,
    _text: opts.text || '', _poison: !!opts.poison, _loaded: false,
  };
  Object.defineProperty(s, 'textFrame', {
    get() {
      // A picture does not throw HERE. This is the whole point of the bug.
      return {
        load() { s._loaded = true; },
        get hasText() { return !!s._text; },
        textRange: { get text() { return s._text; } },
      };
    },
  });
  return s;
}

let SHAPES = [];
const ctx = {
  presentation: {
    slides: {
      get items() { return [{ id: 's1', shapes: { get items() { return SHAPES; }, load() {} } }]; },
      load() {},
    },
  },
  load: () => {},
  sync: async () => {
    SYNCS++;
    // The real failure: the sync rejects if ANY shape in this batch had a load
    // queued against a text frame it does not have.
    const armed = SHAPES.filter((s) => s._poison && s._loaded);
    SHAPES.forEach((s) => { s._loaded = false; });
    if (armed.length) throw new Error('InvalidArgument');
  },
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
  getElementById: () => null, head: { appendChild: () => {} },
};

const host = await import('../src/taskpane/hosts/powerpoint.js');

let pass = 0, fail = 0;
const ok = (name, cond, detail) => {
  cond ? pass++ : fail++;
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}${cond ? '' : '  → ' + (detail || '')}`);
};
const read = () => host.executeTool('ppt_read_slide', { slide: 1 });

// ── A picture must not blind the reader to the rest of the slide ─────────────

SHAPES = [
  fakeShape({ id: 'a', name: '標題', text: '米格魯是一隻跟著鼻子走的家庭犬' }),
  // The full-bleed background that was on the real slide.
  fakeShape({ id: 'b', name: '背景', type: 'Image', poison: true, width: 960, height: 540 }),
  fakeShape({ id: 'c', name: '重點', text: '重點不是訓練牠別聞' }),
];
{
  SYNCS = 0;
  const r = await read();
  ok('整頁還是讀得到', r.ok, r.content);
  ok('圖片前面的文字沒有消失', r.content.includes('米格魯是一隻跟著鼻子走的家庭犬'),
    r.content);
  // 原本的 bug 在這裡：一張圖就讓後面所有文字一起變空。
  ok('圖片後面的文字也沒有消失', r.content.includes('重點不是訓練牠別聞'), r.content);
  ok('圖片本身照樣列出來', r.content.includes('背景') && r.content.includes('960×540'));
}

// 不能只靠「壞掉就一個一個重讀」——那條路一張投影片要幾十次來回。
// 正常情況必須是一次批次搞定。
SHAPES = [
  fakeShape({ id: 'a', name: 'T', text: '一' }),
  fakeShape({ id: 'b', name: 'B', text: '二' }),
  fakeShape({ id: 'c', name: 'C', text: '三' }),
];
{
  SYNCS = 0;
  const r = await read();
  ok('沒有圖片的時候只掃一批', SYNCS <= 2, `${SYNCS} 次 sync`);
  ok('三段文字都在', ['一', '二', '三'].every((t) => r.content.includes(`「${t}」`)), r.content);
}

// 已知不能裝文字的型別直接跳過，連排都不排——這是快的那條路。
SHAPES = [
  fakeShape({ id: 'a', name: 'T', text: '文字' }),
  fakeShape({ id: 'b', name: 'P', type: 'Image', poison: true }),
  fakeShape({ id: 'c', name: 'L', type: 'Line', poison: true }),
  fakeShape({ id: 'd', name: 'G', type: 'Group', poison: true }),
];
{
  SYNCS = 0;
  const r = await read();
  ok('圖片、線條、群組都被跳過，批次沒有被毒到', SYNCS <= 2, `${SYNCS} 次 sync`);
  ok('文字還在', r.content.includes('「文字」'), r.content);
}

// 沒見過的型別（不同 Office 版本會冒出新值）不能讓整頁變空白。
SHAPES = [
  fakeShape({ id: 'a', name: 'T', text: '讀得到' }),
  fakeShape({ id: 'b', name: 'X', type: 'SomethingNewIn2028', poison: true }),
];
{
  const r = await read();
  ok('沒見過的型別會退回逐一讀取，不是整頁放棄',
    r.content.includes('「讀得到」'), r.content);
}

// ── 回答問題的時候可以附圖 ───────────────────────────────────────────────────
//
// AI 問「或直接截一張現在頁面的圖給我」，而那張卡只收文字——於是人只能用文字
// 描述一張截圖，正好是這個問題想避開的事。
{
  const fsm = await import('fs');
  const app = fsm.readFileSync(new URL('../src/taskpane/app.js', import.meta.url), 'utf8');
  const card = app.slice(app.indexOf('function askChoice'), app.indexOf('function isIrreversible'));

  ok('問題卡可以貼圖', card.includes("addEventListener('paste'"));
  ok('也可以拖進來', card.includes("addEventListener('drop'"));
  ok('還有一顆按鈕可以選檔案', card.includes('choice-pic-btn'));
  // 附上的圖要加入「這一輪」的清單，不是下一則訊息的佇列——不然模型這一輪看不到。
  ok('圖加入這一輪的附件清單',
    card.includes('state.sentAttachments = [...state.sentAttachments'));
  // 沒講出來的話，模型根本不知道有圖。
  ok('答案會告訴模型圖在第幾張',
    card.includes('我附了') && card.includes('added.join'));
  ok('只貼圖不打字也能送出',
    card.includes('say.value.trim() || added.length'),
    '「就是這個，你看圖」是人真的會給的答案');
  ok('縮圖會顯示出來', card.includes('choice-pic'),
    '送出去就看不到了，人要能確認自己貼對');
  ok('非圖片檔會被擋掉', card.includes('image/') && card.includes('file.type'));
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
