/**
 * ppt_run_script 真的跑得起來嗎？
 *
 * ── 為什麼需要這個檔案 ──
 * 這個工具在增益集裡實作完整，但**伺服器的 pptToolSpec.ts 從來沒有宣告它**，
 * 所以模型不知道它存在、一次也沒呼叫過。一段從來沒被執行過的程式碼，和一段壞掉的
 * 程式碼，在外部觀察上是同一件事。
 *
 * 而這正是成本問題的關鍵：Word 和 Excel 的 run_script 讓模型能用「寫一支迴圈」
 * 取代「逐個列舉六十個矩形的座標」，兩者的輸出 token 差一個量級。PPT 少了這條路，
 * 於是只能一個一個畫——工具呼叫紀錄裡 ppt_add_shape 84 次、ppt_delete_shape 77 次，
 * 每一次都是一輪完整的模型往返。
 *
 * 開啟它之前要先確定它是好的。這個 session 已經有兩次「讀程式碼看不出來、跑一次
 * 立刻現形」的經驗（powerpoint.js 的 mixHex 未定義、Office 的圖示快取），所以這裡
 * 不讀程式碼，直接執行。
 *
 * ── 這個檔案不測什麼 ──
 * 沙箱的安全邊界由 script-guard.test.mjs 負責（那支測的是 makeScriptGuard 本身，
 * 三個宿主共用）。這裡只回答一件事：PowerPoint 這條路接得起來嗎。
 */
import assert from 'assert';

let failed = 0;
function check(name, fn) {
  try {
    const r = fn();
    if (r instanceof Promise) return r.then(
      () => console.log(`PASS  ${name}`),
      (e) => { failed++; console.error(`FAIL  ${name}\n      ${e.message}`); });
    console.log(`PASS  ${name}`);
  } catch (e) {
    failed++;
    console.error(`FAIL  ${name}\n      ${e.message}`);
  }
  return undefined;
}

// ── 最小的 PowerPoint 替身 ───────────────────────────────────────────────────
//
// 只做到 runScript 會碰到的部分：ctx.sync、presentation.slides。
// 替身要像真的輸入，不是像好寫的輸入——所以 slides.items 是陣列、每個 shape 有
// 真的 left/top/width/height，腳本寫出來的迴圈才碰得到東西。
let synced = 0;
const added = [];

function fakeShape(name, box = {}) {
  return {
    name,
    left: box.left ?? 0, top: box.top ?? 0,
    width: box.width ?? 100, height: box.height ?? 50,
    fill: { setSolidColor(c) { this._fill = c; } },
    textFrame: { textRange: { text: '', font: {} } },
    delete() { this._deleted = true; },
    load: () => {},
  };
}

const slide = {
  shapes: {
    items: [fakeShape('a'), fakeShape('b')],
    load: () => {},
    addGeometricShape(_kind, opts) {
      const s = fakeShape(`new${added.length}`, opts);
      added.push(s);
      slide.shapes.items.push(s);
      return s;
    },
  },
};

const ctx = {
  presentation: { slides: { items: [slide], load: () => {}, getItemAt: () => slide } },
  sync: async () => { synced++; },
  load: () => {},
  trackedObjects: { add: () => {}, remove: () => {} },
};

globalThis.PowerPoint = {
  run: async (fn) => fn(ctx),
  GeometricShapeType: { rectangle: 'rectangle' },
};
globalThis.Office = {
  context: {
    requirements: { isSetSupported: () => true },
    document: { url: 'C:\\d.pptx', addHandlerAsync: () => {} },
  },
  CoercionType: { Text: 'text' },
  onReady: () => {},
};

const host = await import('../src/taskpane/hosts/powerpoint.js');

/**
 * executeTool 的契約是 `{ ok, content }` / `{ ok: false, error }`——**不拋例外**。
 *
 * 這一點我第一版測試寫錯了：假設它會拋，結果十題全紅，紅的原因全是「沒有拋」。
 * 那種失敗訊息會讓人以為產品壞了，實際上壞的是測試的假設。
 *
 * ok() 取成功時的文字，fail() 取失敗時的訊息，兩個分開才不會把「成功但內容不對」
 * 和「根本沒執行」混為一談。
 */
const call = (args, opts = { writeApproved: true }) =>
  host.executeTool('ppt_run_script', args, opts);

async function ok(args, opts) {
  const r = await call(args, opts);
  assert.ok(r.ok, `預期成功，但回了：${r.error || r.content}`);
  return String(r.content ?? '');
}

async function fail(args, opts) {
  const r = await call(args, opts);
  assert.ok(!r.ok, `預期失敗，但成功了：${r.content}`);
  return String(r.error ?? r.content ?? '');
}

// ── 1. 最基本的：它跑得起來嗎 ────────────────────────────────────────────────
//
// 這一題如果紅了，代表這個工具從來就不能用——而沒有人會知道，因為模型不曾呼叫它。
await check('唯讀腳本跑得起來，而且拿得到回傳值', async () => {
  const out = await ok({ code: 'return 1 + 1;', mode: 'read' }, {});
  assert.ok(out.includes('2'), `回傳的是：${out}`);
});

await check('log() 的輸出會被收集起來', async () => {
  const out = await ok({ code: 'log("第一行"); log("第二行"); return "ok";', mode: 'read' }, {});
  assert.ok(out.includes('第一行') && out.includes('第二行'), out);
});

await check('沒有回傳值也要講得出話，不是空白', async () => {
  const out = await ok({ code: 'const x = 1;', mode: 'read' }, {});
  assert.ok(out.trim().length > 0, '回了空字串——模型會以為工具壞了');
});

// ── 2. 真正的用途：迴圈 ──────────────────────────────────────────────────────
//
// 這是整件事的重點。模型用「寫一支迴圈」取代「列舉六十個矩形」，輸出 token 才會
// 從三千降到三十。迴圈碰不到簡報的話，這條路就沒有意義。
await check('腳本摸得到投影片與形狀', async () => {
  const out = await ok({
    code: 'const s = presentation.slides.items[0]; s.shapes.load("items"); '
      + 'await context.sync(); return s.shapes.items.length;',
    mode: 'read',
  }, {});
  assert.ok(out.includes('2'), `應該看到兩個既有形狀，實際：${out}`);
});

await check('迴圈真的能產生多個物件', async () => {
  added.length = 0;
  const out = await ok({
    mode: 'write',
    code: 'const s = presentation.slides.items[0];'
      + 'for (let i = 0; i < 12; i++) {'
      + '  s.shapes.addGeometricShape(PowerPoint.GeometricShapeType.rectangle,'
      + '    { left: 40 + i * 70, top: 200, width: 60, height: 40 });'
      + '}'
      + 'return 12;',
  });
  assert.strictEqual(added.length, 12, `只產生了 ${added.length} 個`);
  // 座標要真的是迴圈算出來的，不是全部疊在一起
  assert.strictEqual(added[0].left, 40);
  assert.strictEqual(added[11].left, 40 + 11 * 70);
  assert.ok(out.includes('12'), out);
});

// ── 3. 寫入要先經過同意 ─────────────────────────────────────────────────────
await check('沒有核准就不准寫', async () => {
  const msg = await fail({ mode: 'write', code: 'presentation.slides.items[0].shapes.addGeometricShape("rectangle", {});' }, {});
  assert.match(msg, /同意|approve/i, msg);
});

await check('唯讀腳本不需要核准', async () => {
  const out = await ok({ code: 'return "read";', mode: 'read' }, {});
  assert.ok(out.includes('read'), out);
});

// ── 4. 壞掉的腳本要說得出原因 ───────────────────────────────────────────────
//
// 模型靠錯誤訊息修正自己。「執行失敗」這種訊息會讓它重試一模一樣的程式碼。
await check('語法錯誤要回報得出來', async () => {
  const msg = await fail({ code: 'const = ;', mode: 'read' }, {});
  assert.ok(msg.trim().length > 0, '失敗了卻說不出原因——模型只會原樣重試一次');
});

await check('執行期錯誤要帶出原本的訊息', async () => {
  const msg = await fail({ code: 'throw new Error("我自己丟的");', mode: 'read' }, {});
  assert.match(msg, /我自己丟的/, msg);
});

await check('空的 code 要擋下來', async () => {
  const msg = await fail({ code: '   ', mode: 'read' }, {});
  assert.match(msg, /code/, msg);
});

// ── 5. 收尾 ─────────────────────────────────────────────────────────────────
await check('執行完會 sync，變更才會真的寫進簡報', async () => {
  const before = synced;
  await ok({ code: 'return 1;', mode: 'read' }, {});
  assert.ok(synced > before, 'sync 沒有被呼叫——畫出來的東西不會出現在簡報裡');
});

console.log(failed ? `\n${failed} failed` : '\n全部通過');
process.exit(failed ? 1 : 0);
