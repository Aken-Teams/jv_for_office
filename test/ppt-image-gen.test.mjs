/**
 * 地端生圖 — the model gateway, and the prompt that is most of the product.
 *
 * ── Why the prompt gets tested like code ──
 * 「我自己下 prompt 都產生醜醜的東西，AI 下 prompt 產生的圖都蠻美的」. The difference
 * is four instructions nobody thinks to write: a composition with room for the
 * title, the deck's own palette, a style register, and NO LETTERS. Drop any one
 * and the output is the same model producing a worse picture — silently, because
 * an ugly background does not throw.
 *
 * ── The three things that must never regress ──
 *   1. NO TEXT. This model renders Chinese unusually well, which makes it more
 *      tempting to let it, not less — and a title baked into a raster cannot be
 *      corrected without regenerating the page.
 *   2. DETERMINISM. The seed comes from the request, because the endpoint is
 *      byte-deterministic (verified). Rebuilding slide 4 must give back the same
 *      background; artwork that shifts reads as an unreliable tool, not variety.
 *   3. ONE AT A TIME. Single GPU. Two concurrent calls do not halve the wait,
 *      they double it for both — and a retry after a timeout makes it worse.
 */
import fs from 'fs';

let pass = 0, fail = 0;
const ok = (name, cond, detail) => {
  cond ? pass++ : fail++;
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}${cond ? '' : '  → ' + (detail || '')}`);
};

const serverDir = new URL('../../ai-agents-office/server/src/', import.meta.url);
if (!fs.existsSync(serverDir)) {
  console.log('      （找不到 ai-agents-office，跳過地端生圖的跨 repo 檢查）');
  process.exit(0);
}

// ── The composed prompt ──────────────────────────────────────────────────────

const ip = await import(new URL('services/imagePrompt.ts', serverDir).href).catch(() => null);
if (ip) {
  const hero = ip.composeImagePrompt({
    subject: 'a modern automated warehouse interior',
    mode: 'hero', look: 'business', theme: 'corporate',
  });

  ok('主體有進去', hero.prompt.includes('automated warehouse'));
  // 圖裡的字改不了。這是唯一沒有事後補救辦法的失敗。
  ok('正面提示詞裡明講不要有字', /no text, no letters/.test(hero.prompt));
  ok('反面提示詞也擋一次', /\btext\b/.test(hero.negative) && /letters/.test(hero.negative));
  ok('反面提示詞連圖表也擋掉', /chart|graph|infographic/.test(hero.negative),
    '有數字的東西要走 ppt_add_chart，不能讓它用畫的');

  // 沒有這一條，滿版背景就是一張漂亮但沒辦法放標題的圖。
  ok('hero 會要求留一塊空白給標題',
    /empty area|for text to sit over/.test(hero.prompt), hero.prompt);
  ok('hero 是 16:9', hero.size === '1024x576', hero.size);
  ok('side 是直式', ip.composeImagePrompt({ subject: 'x', mode: 'side' }).size === '768x1024');

  // 配色講給模型聽要用字，不是 hex——而這是「這張圖屬於這份簡報」跟
  // 「一張剛好放在旁邊的漂亮圖片」的差別。
  ok('配色用文字描述，不是色碼',
    /blue|grey|gray/.test(hero.prompt) && !/#[0-9A-Fa-f]{6}/.test(hero.prompt), hero.prompt);
  const linen = ip.composeImagePrompt({ subject: 'x', theme: 'linen', look: 'literary' });
  ok('換配色就換顏色描述', /terracotta|sage/.test(linen.prompt), linen.prompt);
  ok('換版式就換調性', linen.prompt !== hero.prompt.replace('a modern automated warehouse interior', 'x'));

  // ── 插畫跟背景要的是相反的東西 ─────────────────────────────────────────
  //
  // 所有規則原本都是為**背景**寫的：空氣感、低對比、柔和均勻的光——因為 hero
  // 上面還要壓深色遮罩和標題。同一套指令拿去畫一張扁平插畫，就產出一張淡到
  // 縮圖看不見的圖。兩次抱怨（「有點可怕」和「這 2 張圖不行」）是同一個根因：
  // 一套打光規則套在兩種不同的工作上。
  {
    const flat = ip.composeImagePrompt({ subject: 'an assembly line', style: 'flat',
      theme: 'teal', mode: 'block' });
    ok('插畫要求實心飽和', /solid saturated fills/.test(flat.prompt));
    ok('插畫要求清楚的輪廓', /strong silhouette/.test(flat.prompt));
    ok('插畫要求白底', /white background/.test(flat.prompt));
    // 這幾句是背景用的，套在扁平向量上就是把它洗白。
    ok('插畫不帶背景用的打光詞', !/soft even lighting|calm and welcoming/.test(flat.prompt));
    ok('插畫的配色不帶亮度推力',
      !/airy|plenty of light|gentle daylight|bright and even/.test(flat.prompt), flat.prompt);
    // 主體不需要為標題留洞，那會把圖推到自己畫框的一角。
    ok('插畫不留標題的空位', !/empty area/.test(flat.prompt));
    ok('但仍然是這份簡報的色系', /teal|aqua/.test(flat.prompt));

    const bg = ip.composeImagePrompt({ subject: 'an assembly line', style: 'abstract',
      theme: 'teal', mode: 'hero' });
    ok('背景維持原本的空氣感規則', /calm and welcoming/.test(bg.prompt));
    ok('背景仍然留空位給標題', /empty area/.test(bg.prompt));

    // 分流看的是「風格」和「位置」兩件事。一張照片放進 feature 版型是要被看的，
    // 不是襯在標題底下——第一版只看風格，照片就跟著吃到背景那一套，一樣會偏淡。
    const photoPanel = ip.composeImagePrompt({ subject: 'a product', style: 'photo',
      theme: 'teal', mode: 'block' });
    ok('面板裡的照片不吃亮度推力',
      !/airy|plenty of light|gentle daylight|bright and even/.test(photoPanel.prompt),
      photoPanel.prompt);
    ok('面板裡的照片不留標題空位', !/empty area/.test(photoPanel.prompt));
    ok('面板裡的照片要求看得清楚', /clearly readable/.test(photoPanel.prompt));
    // 但它仍然是照片，所以打光語言要留著——「實心飽和」對照片沒有意義。
    ok('面板裡的照片仍然有打光語言', /soft even lighting/.test(photoPanel.prompt));
    ok('而且不會被當成插畫', !/solid saturated/.test(photoPanel.prompt));

    const photoHero = ip.composeImagePrompt({ subject: 'a product', style: 'photo',
      theme: 'teal', mode: 'hero' });
    ok('滿版的照片還是背景', /empty area/.test(photoHero.prompt));

    // 沒有人把扁平插畫擺在標題底下當底，所以它在哪個位置都是主體。
    const drawHero = ip.composeImagePrompt({ subject: 'a factory', style: 'flat', mode: 'hero' });
    ok('插畫在 hero 也還是主體', /solid saturated/.test(drawHero.prompt));
  }

  // 步數也跟著位置走：襯在遮罩底下的十步夠，面板裡的不夠。
  ok('面板裡的照片步數比背景多',
    ip.stepsFor('photo', false, 'block') > ip.stepsFor('photo', false, 'hero'),
    `block ${ip.stepsFor('photo', false, 'block')} vs hero ${ip.stepsFor('photo', false, 'hero')}`);
  ok('插畫仍然最高', ip.stepsFor('flat', false, 'hero') > ip.stepsFor('photo', false, 'block'));

  // 步數：抽象背景十步就成形，扁平插畫十步還是一團糊。
  ok('插畫用比較多步數',
    ip.stepsFor('flat', false, 'hero') > ip.stepsFor('abstract', false, 'hero'),
    `flat ${ip.stepsFor('flat', false, 'hero')} vs abstract ${ip.stepsFor('abstract', false, 'hero')}`);
  // 「背景」現在要連位置一起看——abstract 放在面板裡也不是背景。
  ok('滿版背景維持低步數（快）', ip.stepsFor('abstract', false, 'hero') <= 12);
  ok('精細模式最高', ip.stepsFor('flat', true) > ip.stepsFor('flat'));

  // ── 「有點可怕，很像命案現場」 ───────────────────────────────────────────
  //
  // 第一版把配色裡的強調色也交給影像模型：「deep navy 加一點暖紅」。在向量
  // 圖裡那是一條 4pt 的線；在擴散模型裡那是一個**光源**——深色場景配紅色光源
  // 就是事故現場。這是分類錯誤，不是措辭問題。
  ok('配色描述裡沒有強調色',
    !/accent/i.test(hero.prompt), hero.prompt);
  ok('尤其沒有紅色',
    !/red/i.test(hero.prompt), '強調色屬於畫在圖上面的形狀，不屬於圖本身');

  // 給了工業主題卻沒說氣氛，模型會往「有張力」去——而工廠的張力跟工安事故長一樣。
  ok('每張圖都墊一層氣氛底線',
    /calm and welcoming/.test(hero.prompt) && /nothing ominous/.test(hero.prompt));
  ok('氣氛底線跟主題無關，一律都加',
    /nothing ominous/.test(ip.composeImagePrompt({ subject: 'x', theme: 'night' }).prompt));

  ok('反面提示詞擋掉會嚇人的打光',
    ['low key', 'ominous', 'horror', 'red glow', 'neon rim light', 'emergency lighting']
      .every((k) => hero.negative.includes(k)),
    hero.negative);
  // 但不能連「暗」本身都擋掉——第一次修就是這樣，night 和 tech 變成灰白的，
  // 那是另一種錯。會嚇人的是打光方式，不是亮度。
  ok('沒有擋掉「暗」本身',
    !/(^|[, ])dark([, ]|$)/.test(hero.negative),
    'night／tech 本來就該是深色的');
  ok('深色配色還是深色',
    /deep charcoal|midnight/.test(ip.composeImagePrompt({ subject: 'x', theme: 'night' }).prompt));
  ok('但要求它平靜',
    /calm and quiet/.test(ip.composeImagePrompt({ subject: 'x', theme: 'night' }).prompt));

  // 端點本身是位元組等價的確定性，所以種子固定就等於整份簡報可重現。
  const a = ip.composeImagePrompt({ subject: 'same', mode: 'hero', theme: 'corporate' });
  const b = ip.composeImagePrompt({ subject: 'same', mode: 'hero', theme: 'corporate' });
  ok('同樣的請求給同樣的種子', a.seed === b.seed, `${a.seed} vs ${b.seed}`);
  ok('不同的請求給不同的種子',
    a.seed !== ip.composeImagePrompt({ subject: 'other', mode: 'hero' }).seed);
  ok('種子在 32 位元範圍內', a.seed > 0 && a.seed < 2147483647, String(a.seed));

  // 使用者寫的東西長度沒有上限，提示詞有。
  ok('主體會被截斷', ip.composeImagePrompt({ subject: 'x'.repeat(900) }).prompt.length < 900);
  ok('沒給 mode 就當 hero', ip.composeImagePrompt({ subject: 'x' }).size === '1024x576');
  ok('亂給 mode 也退回 hero', ip.composeImagePrompt({ subject: 'x', mode: 'zzz' }).size === '1024x576');
}

// ── The gateway ──────────────────────────────────────────────────────────────

const gw = fs.readFileSync(new URL('services/modelGateway.ts', serverDir), 'utf8');

// guidance > 1 讓每一步跑兩次 transformer，時間正好加倍，而這個模型不需要它。
ok('guidance 寫死 1', /guidance: 1,/.test(gw));
ok('而且解釋了為什麼', gw.includes('twice per step'));
// 單機 GPU。兩個同時送不會變快，只會一起變慢。
ok('請求排隊，一次一張', /function serialise/.test(gw) && /let queue/.test(gw));
ok('一個失敗不會卡住後面', gw.includes('must survive a failed job'));
// 15～56 秒的工作用預設 30 秒的 timeout 等於永遠失敗。
ok('timeout 給到 300 秒', /TIMEOUT_MS = 300000/.test(gw));
ok('逾時會叫使用者不要重送', gw.includes('不要重送'));

// AC-19：不得靜默降級。
ok('降級政策預設是拒絕而不是偷偷改用雲端',
  /onUnavailable: process\.env\.IMAGE_FALLBACK_POLICY === 'cloud' \? 'cloud' : 'refuse'/.test(gw));
ok('拒絕的時候會講出來', gw.includes('不會改用雲端模型'));
ok('降級成功的時候也會標記', /fellBack/.test(gw));
// AC-18：要能證明「全程沒有對外傳輸」，就得有地方留紀錄。
ok('每次呼叫都留稽核紀錄', (gw.match(/logAiCall\(/g) || []).length >= 3);
ok('紀錄標為地端', /authMode: 'local'/.test(gw));

// 端點與金鑰不能外洩到增益集，那是一台跑在使用者電腦上的 WebView。
ok('端點來自設定不是寫死', /config\.localLlmBaseUrl/.test(gw));
ok('沒設定會說清楚要管理員去設', gw.includes('尚未設定'));
// 生成的圖是暫存的裝飾品，不是誰的文件——但也不能無限累積。
ok('圖片有 TTL', /TTL_MS = 30 \* 60 \* 1000/.test(gw));
ok('也有數量上限', /MAX_HELD/.test(gw));
ok('id 是 128 位元亂數', /randomBytes\(16\)/.test(gw));
ok('而且說明了為什麼這條路不掛 auth', gw.includes('That is the trade'));

// ── Wiring ───────────────────────────────────────────────────────────────────

const bridge = fs.readFileSync(new URL('services/excelBridge.ts', serverDir), 'utf8');
ok('生圖由伺服器直接回答', /tool === 'ppt_generate_image'/.test(bridge));
// 步數改成看「在畫什麼」而不是一個常數——背景十步就成形，插畫十步還是一團糊。
ok('步數交給 stepsFor 決定', /stepsFor\(String\(args\.style/.test(bridge));
// 用設定裡的位址組網址，部署方式一換就默默失效。
ok('回傳網址用連線當下的來源組出來', /run\.origin/.test(bridge));

const route = fs.readFileSync(new URL('routes/ppt.ts', serverDir), 'utf8');
ok('有一條取圖的路由', /router\.get\('\/image\/:id\.png'/.test(route));
ok('而且刻意排在 authMiddleware 前面',
  route.indexOf("router.get('/image/:id.png'") < route.indexOf('router.use(authMiddleware)'));
ok('過期的圖會回 404，不是空白', route.includes('這張圖已經過期或不存在'));

const spec = fs.readFileSync(new URL('services/pptToolSpec.ts', serverDir), 'utf8');
const gen = spec.slice(spec.indexOf("name: 'ppt_generate_image'"), spec.indexOf("name: 'ppt_find_art'"));
ok('工具說明點名什麼時候不要用它',
  gen.includes('ppt_add_chart') && gen.includes('ppt_add_diagram') && gen.includes('ppt_find_art'));
ok('而且說出理由是中文字畫不好又改不了',
  gen.includes('中文字畫不好') && gen.includes('改不了'));
ok('提醒一次只能算一張', gen.includes('不要連續呼叫'));
ok('subject 要求只寫畫什麼', gen.includes('只寫「畫什麼」'));

// ── The rules the model reads ────────────────────────────────────────────────

const ctx = fs.readFileSync(new URL('services/pptContext.ts', serverDir), 'utf8');
ok('有一張「這張圖要做什麼」的決策表', ctx.includes('先決定「這張圖要做什麼」'));
ok('五種需求各自指名工具',
  ['ppt_add_chart', 'ppt_add_diagram', 'ppt_find_art', 'ppt_add_photo', 'ppt_generate_image']
    .every((k) => ctx.includes(k)));
ok('生圖只用在沒有特定意義的位置', ctx.includes('只是要好看'));
ok('並且要求多用扁平插畫', ctx.includes('要多用'));
ok('一份簡報最多兩到三張生成圖', ctx.includes('最多兩到三張'));

// ── The add-in side ──────────────────────────────────────────────────────────

globalThis.window = globalThis;
globalThis.document = {
  createElement: () => ({ getContext: () => null, style: {} }),
  getElementById: () => null, head: { appendChild: () => {} },
};
globalThis.Office = {
  context: {
    requirements: { isSetSupported: () => true },
    document: { url: 'C:\\d.pptx', addHandlerAsync: () => {} },
  },
  EventType: { DocumentSelectionChanged: 'sel' },
  CoercionType: { Image: 'image', XmlSvg: 'xmlsvg' },
  AsyncResultStatus: { Succeeded: 'ok' },
};
globalThis.PowerPoint = { run: async () => {} };
const host = await import('../src/taskpane/hosts/powerpoint.js');
// 伺服器自己回答這個工具，但工具清單是從增益集報上去的——沒宣告就等於不存在。
ok('ppt_generate_image 有被宣告出去', host.supportedTools().includes('ppt_generate_image'));
ok('而且有進度字串', !!host.meta.toolLabels().ppt_generate_image);
ok('進度字串講出大概要等多久',
  /秒/.test(host.meta.toolLabels().ppt_generate_image),
  '等十幾秒沒有任何訊息，使用者會以為當掉了');

// ── 寫實照片開關 ──────────────────────────────────────────────────────────────
/**
 * 這個開關必須同時切掉兩件事：模型「問得出口」和伺服器「答得出來」。
 *
 * 只切一邊都不夠。只改 schema，一個舊版窗格或設定中途改變的 run 就繞過去了；
 * 只擋伺服器，模型會一直重試一個它以為存在的選項，燒掉回合數。所以兩邊都測。
 *
 * 背景和插畫不在開關後面，而且必須確認沒被波及——會被誤認成「記錄」的只有照片，
 * 把整個生圖能力一起關掉是另一回事，不是使用者按那顆按鈕的意思。
 */
const toolSpec = await import(new URL('services/pptToolSpec.ts', serverDir).href).catch(() => null);
if (toolSpec) {
  const styleOf = (tools) => tools.find((x) => x.name === 'ppt_generate_image')
    .inputSchema.properties.style;
  const onTools = toolSpec.pptToolsFor({ allowPhoto: true });
  const offTools = toolSpec.pptToolsFor({ allowPhoto: false });

  ok('預設（沒給參數）不動原本的規格', toolSpec.pptToolsFor() === toolSpec.PPT_TOOLS);
  ok('開著的時候 photo 還在', styleOf(onTools).enum.includes('photo'));
  ok('關掉以後 photo 不在 enum 裡', !styleOf(offTools).enum.includes('photo'));
  // 這才是重點：擋的是「說明」還是「schema」。說明是建議，enum 是牆。
  ok('關掉以後說明裡也不再介紹 photo',
    !/- photo /.test(styleOf(offTools).description), styleOf(offTools).description);
  ok('關掉以後有講替代做法',
    /abstract|背景|插畫/.test(offTools.find((x) => x.name === 'ppt_generate_image').description));

  for (const style of ['abstract', 'flat', 'handdrawn', 'isometric', 'gradient', 'texture']) {
    ok(`關掉照片不影響 ${style}`, styleOf(offTools).enum.includes(style));
  }
  ok('其他工具一個都沒少', offTools.length === toolSpec.PPT_TOOLS.length);
  ok('沒有動到原本的常數', styleOf(toolSpec.PPT_TOOLS).enum.includes('photo'));

  // 伺服器端的那道牆。excelBridge 匯入不進來（Node 的型別剝離不會把 './x.js'
  // 轉回 .ts），但真正在判斷的是這個純函式——bridge 只是呼叫它，所以測它就是
  // 測那道牆。
  const off = toolSpec.imageStyleRefusal('photo', false);
  ok('關掉時伺服器擋下寫實照片', typeof off === 'string' && off.length > 0);
  // 「不行」會被重試，「改用這個」不會。
  ok('而且講了要改用什麼', /abstract|flat|插畫/.test(off || ''), off);
  ok('也講了在哪裡開', /選單|開啟/.test(off || ''), off);
  ok('說的替代畫法在 enum 裡真的存在',
    ['abstract', 'gradient', 'flat', 'handdrawn']
      .every((v) => styleOf(offTools).enum.includes(v)));

  ok('關掉照片不影響背景', toolSpec.imageStyleRefusal('abstract', false) === null);
  ok('關掉照片不影響插畫', toolSpec.imageStyleRefusal('flat', false) === null);
  ok('沒寫 style 不會被擋', toolSpec.imageStyleRefusal(undefined, false) === null);
  ok('大小寫不能繞過', toolSpec.imageStyleRefusal('Photo', false) !== null);
  ok('開著的時候放行', toolSpec.imageStyleRefusal('photo', true) === null);
}

// ── 窗格那顆按鈕 ──────────────────────────────────────────────────────────────
const app = fs.readFileSync(new URL('../src/taskpane/app.js', import.meta.url), 'utf8');
const html = fs.readFileSync(new URL('../src/taskpane/taskpane.html', import.meta.url), 'utf8');
ok('選單裡有這顆開關', /data-cap="photo"/.test(html));
ok('預設是關的', /localStorage.getItem\(LS_PHOTO_GEN\) === '1'/.test(app));
ok('會記住使用者的選擇', /localStorage.setItem\(LS_PHOTO_GEN/.test(app));
ok('會送給伺服器', /allowPhoto: state.photoGen/.test(app));
// 只在 PowerPoint 出現——另外兩個 host 沒有圖片工具可管。
ok('只在 PowerPoint 顯示', /src-cap'\).hidden = !\(host && host.meta.id === 'powerpoint'\)/.test(app));

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
