/**
 * 個人常用指令 — the rules that are security rules, not features. PRD 3.5.
 *
 * ── Why this test reads the server's source ──
 * The interesting guarantees here are absences: no sharing column, no query that
 * omits the owner filter, no share/export button. None of those can be asserted
 * by calling the API, because a passing call proves only that the path somebody
 * thought of is closed. What AC-43 actually asks is that there is no path at
 * all — and that is a property of the code, so the code is what gets read.
 *
 * It is a cross-repo check and skips when the server is not checked out beside
 * this one, exactly like the tool-name check in no-duplicates.test.mjs.
 */
import fs from 'fs';

let pass = 0, fail = 0;
const ok = (name, cond) => {
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}`);
  if (cond) pass++; else fail++;
};

const serverDir = new URL('../../ai-agents-office/server/src/', import.meta.url);
if (!fs.existsSync(serverDir)) {
  console.log('      （找不到 ai-agents-office，跳過常用指令的跨 repo 檢查）');
  process.exit(0);
}

const routeSrc = fs.readFileSync(new URL('routes/promptTemplates.ts', serverDir), 'utf8');
const db = fs.readFileSync(new URL('db.ts', serverDir), 'utf8');

/**
 * The file with its comments removed.
 *
 * Needed because the comments in that file EXPLAIN the rules — 「carries no
 * tools, no connectors」 — so a check that scans the raw text finds the words it
 * is looking for in the prose that says they are absent, and fails a file that
 * is correct. Comments are documentation; the code is the claim.
 */
const route = routeSrc
  .replace(/\/\*[\s\S]*?\*\//g, '')
  .replace(/^\s*\/\/.*$/gm, '');

// ── 僅本人可見（AC-43）─────────────────────────────────────────────────────
//
// Every read of a row has to carry the owner. A SELECT that forgets it is not a
// bug that shows up in testing — it works perfectly until the day two people
// use the product.
{
  // Whole statements, so the SELECT list is visible and the ownership probe can
  // be told apart from a query that simply forgot the filter.
  const selects = [...route.matchAll(/SELECT[\s\S]*?FROM prompt_templates[^`']*/g)].map(m => m[0]);
  ok('至少有查詢可以檢查', selects.length >= 3);
  ok('每一個查詢不是帶 user_id，就是那唯一一個只取擁有者的探測',
    selects.every(q => /user_id = \?/.test(q)
      || /^SELECT user_id FROM prompt_templates WHERE id = \?$/.test(q.trim())));
  // 那個探測只能取到 user_id 一個欄位——它的用途是「這個 id 是別人的嗎」，
  // 取到內容就等於把別人的常用指令讀出來了。
  ok('擁有者探測只取 user_id，不會連內容一起撈出來',
    (selects.filter(q => !/user_id = \?/.test(q)) || [])
      .every(q => /^SELECT user_id FROM/.test(q.trim())));
  ok('寫入與刪除也都帶 user_id',
    [...route.matchAll(/(UPDATE|DELETE FROM) prompt_templates[^`']*/g)]
      .every(m => /user_id = \?/.test(m[0])));
}

// 這一條存在的理由：只有一個地方能用 id 把整列取出來，越權才有唯一的攔截點。
ok('用 id 取整列只有 mine() 一個入口',
  (route.match(/SELECT \* FROM prompt_templates/g) || []).length === 1);

ok('查不到時回 404 而不是 403——403 等於確認那個 id 存在',
  route.includes('404') && !/status\(403\)/.test(route));

ok('越權存取記進稽核（AC-43）',
  route.includes('logSecurityEvent') && route.includes('嘗試存取他人的常用指令'));

// ── 不提權（AC-43b）────────────────────────────────────────────────────────
//
// A template is text. If it ever gains a tools/connector/skip-confirmation
// field it stops being one, and the "anyone may create these" rule stops being
// safe — so the absence is worth pinning rather than trusting to review.
ok('常用指令不綁工具、連接器或免確認旗標',
  !/\b(tools|mcp|connector|dataSources|skipConfirm|auto_approve)\b/i.test(route));

// ── 不能分享或匯出 ─────────────────────────────────────────────────────────
{
  const table = (db.match(/CREATE TABLE IF NOT EXISTS prompt_templates[\s\S]*?\)\s*ENGINE/) || [''])[0];
  ok('資料表存在', table.length > 0);
  ok('資料表沒有任何可以擴散出去的欄位',
    table.length > 0 && !/\b(shared|visibility|team_id|org_id|public|is_public)\b/i.test(table));
  ok('綁定使用者，且使用者刪除時一起刪',
    /user_id\s+VARCHAR\(36\)\s+NOT NULL/.test(table) && /ON DELETE CASCADE/.test(table));
}

// ── 稽核：記錄該次對話套用了哪一則 ─────────────────────────────────────────
ok('三個增益集的 /chat 都把套用的常用指令寫進 messages',
  ['excel', 'word', 'ppt'].every((host) => {
    const src = fs.readFileSync(new URL(`routes/${host}.ts`, serverDir), 'utf8');
    return src.includes('templateForAudit') && src.includes('prompt_template_id');
  }));

ok('稽核用的 id 由伺服器驗證擁有者，不是照單全收',
  route.includes('export async function templateForAudit') && /templateForAudit[\s\S]*?mine\(/.test(route));

// ── 窗格：介面上沒有分享／匯出／上架按鈕 ───────────────────────────────────
{
  const html = fs.readFileSync(new URL('../src/taskpane/taskpane.html', import.meta.url), 'utf8');
  const menu = (html.match(/<div id="cmd-menu"[\s\S]*?<div id="cmd-note"/) || [''])[0];
  ok('指令與技能選單在窗格裡', menu.length > 0);
  ok('選單裡沒有分享、匯出或上架',
    menu.length > 0 && !/(分享|匯出|上架|share|export|publish)/i.test(menu));
  // PRD 3.5 要求技能與常用指令在介面上分開命名：技能綁工具與權限，常用指令
  // 只是文字。合併成一個入口沒問題，混成一區就有問題。
  const app2 = fs.readFileSync(new URL('../src/taskpane/app.js', import.meta.url), 'utf8');
  ok('兩者在同一個選單裡但分區標示',
    app2.includes('公司技能') && app2.includes('我的常用指令'));
}

// ── 存檔前的注入掃描 ───────────────────────────────────────────────────────
//
// 這道防線擋的是「常駐化」。工作簿裡的注入本來是一次性的：影響一輪、確認閘門
// 擋下破壞、就沒了。但存成常用指令會打破這件事——產生那段文字的 agent 一直在
// 讀那份工作簿，而使用者會因為「是我自己存的」而信任它，然後每個月重播一次。
{
  ok('建立時會掃', /router\.post[\s\S]*?screen\(userId/.test(route));
  // 只在建立時掃等於留了一條有紀錄的繞道：建一則乾淨的，再改成髒的。
  ok('編輯時也會掃', /router\.patch[\s\S]*?screen\(userId/.test(route));
  ok('掃到就不存，而且回得出原因',
    /if \(refused\)[\s\S]*?status\(400\)/.test(route));
  ok('擋下來要記稽核', /screen[\s\S]*?logSecurityEvent\([^)]*'prompt_injection', 'high'/.test(route));
  ok('用散文的掃描器，不是檔案那條分塊路徑',
    route.includes('analyzeInput(') && !route.includes('analyzeFileContentChunked'));
}

// 掃描器抓的是注入手法，不是破壞性動詞——這條確認我們沒有在誤擋正常用法。
{
  const guard = fs.readFileSync(new URL('services/inputGuard.ts', serverDir), 'utf8');
  const formula = (guard.match(/function detectFormulaInjection[\s\S]*?\n}/) || [''])[0];
  ok('公式注入只抓 CMD/EXEC 那一類，不會誤判 =SUM()',
    formula.includes('CMD') && formula.includes('EXEC') && !/\bSUM\b/.test(formula));
}

// ── 不可以用瀏覽器的對話框 ─────────────────────────────────────────────────
//
// Office 增益集的 webview **不支援** prompt / confirm / alert，而且是無聲的失敗：
// prompt 回 null，程式讀成「使用者按了取消」，於是點「存成常用指令」什麼事都
// 不會發生，也不會有任何錯誤訊息。這是 bug 最糟的形狀——沒有東西可以搜尋。
{
  const app = fs.readFileSync(new URL('../src/taskpane/app.js', import.meta.url), 'utf8');
  const code = app.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
  ok('沒有 window.prompt', !/window\.prompt\s*\(/.test(code));
  ok('沒有 window.confirm', !/window\.confirm\s*\(/.test(code));
  ok('沒有 window.alert', !/window\.alert\s*\(/.test(code));
  ok('改用窗格自己的命名卡片', code.includes('function askName('));
  ok('刪除確認也改用窗格自己的卡片', code.includes('function askYesNo('));
}

// ── 預覽與建議名稱要去掉 markdown ──────────────────────────────────────────
//
// 存下來的一律是原始 markdown（模型讀那個最準，套用時也要原樣回來），
// 但「名稱」和清單裡那一行摘要是給人看的。agent 寫出來的流程開頭常是
// 「## 流程」或「**先讀三張規則表**」，拿那個當名字等於給人看標點符號。
{
  const app = fs.readFileSync(new URL('../src/taskpane/app.js', import.meta.url), 'utf8');
  ok('有把 markdown 剝掉的函式', app.includes('function plainText('));
  const fl = (app.match(/const firstLine = \(s\) => \{[\s\S]*?\n};/) || [''])[0];
  ok('摘要那一行用剝過的文字', fl.includes('plainText(s)'));
  // 「## 流程」剝完是「流程」——那是標題在做它的事，不是摘要在做它的事。
  ok('會跳過太短的段落標題，挑真的有話的那一行', /length >= 4/.test(fl));
  ok('建議名稱也用剝過的文字', /function suggestName\([\s\S]*?firstLine\(content\)/.test(app));
  // 存進去的必須還是原文，不是剝過的。
  const save = (app.match(/async function saveTemplateText\(content\)[\s\S]*?\n}/) || [''])[0];
  ok('存的是原文，不是剝過的預覽',
    save.includes('content: body') && !/content: plainText/.test(save));
}

// ── 整理出來的流程要主動送到眼前 ───────────────────────────────────────────
//
// 這條是整個「錄製」功能成不成立的關鍵。讓使用者自己去對話裡找那則回覆，
// 他會存看起來最重要的那一則——而看起來最重要的是 agent 的「回填完成，已驗過」
// 報告，那是**結果**。存下結果，下個月的執行學不到任何東西。
{
  const app = fs.readFileSync(new URL('../src/taskpane/app.js', import.meta.url), 'utf8');
  ok('整理完會主動問要不要存', app.includes('function offerProcedure('));
  ok('那一輪跑完就問，不用自己找 bubble',
    /state\.procedureRun[\s\S]{0,300}offerProcedure\(bubble\.dataset\.raw\)/.test(app));
  // 送出去的是一大串工具名稱，直接當成使用者訊息顯示只會把回覆淹掉。
  ok('送出的提示詞跟顯示的文字可以分開',
    /addBubble\('user', opts\.display \|\| message/.test(app));

  // 套用常用指令送回去的常常是 agent 寫的流程（含表格）。使用者訊息若只做
  // escape，那段會變成一整片原始的 | 符號——使用者自己存的文字看起來壞掉。
  ok('使用者訊息也走 markdown（renderMarkdown 內部會先 escape）',
    /el\.innerHTML = renderMarkdown\(text \|\| ''\)/.test(app));
  ok('不再只對使用者訊息做 escape', !/role === 'user' \? escapeHtml\(text\)/.test(app));

  // ── 只能有一個輸入源 ──────────────────────────────────────────────────
  //
  // 原本選單裡另有一個搜尋框，而選單是用輸入框打「/」叫出來的——游標在輸入框，
  // 篩選卻在另一個欄位，而且那個欄位清單短的時候還不出現。結果就是按了「/」
  // 打字，字進了輸入框、選單原封不動，看起來像篩選壞掉。
  // 現在：選單開著 ⟺ 輸入框以「/」開頭，「/」後面就是查詢字串。
  ok('輸入框以 / 開頭就是選單開著', /e\.target\.value\.startsWith\('\/'\)/.test(app));
  ok('查詢字串直接讀輸入框', /function slashQuery\(\)[\s\S]*?v\.slice\(1\)/.test(app));
  ok('沒有第二個搜尋框了', !app.includes('cmd-find'));
  ok('離開選單會把 /query 一起收掉', app.includes('function closeSlash('));
  // 「/月結」是查詢不是問句——送出鍵亮著等於邀請使用者送出一段斜線加半個名字。
  ok('slash 模式下送不出去', /const has = !inSlashMode\(\)/.test(app));
  // 找指令和存指令是相反的差事，搜尋中把存檔列整個收起來而不是留一列灰的。
  ok('slash 模式下存檔那一列整個收起來', /save\.hidden = inSlashMode\(\)/.test(app));

  // ── 「整理成流程」問在存檔的當下，不是藏在選單裡 ─────────────────────
  //
  // 它本來是指令選單裡的一顆按鈕，離任何東西都有兩層遠，旁邊還是一串常用指令。
  // 沒被告知的人永遠找不到它，找到的人也猜不出它讀的是工具紀錄而不是對話。
  ok('存檔時才問要不要整理', app.includes('async function saveAsTemplate('));
  ok('bubble 的存檔鈕走這條', /saveAsTemplate\(bubble\.dataset\.raw/.test(app));
  ok('沒有紀錄就不問——對什麼都沒跑過的對話問「要整理嗎」等於提議做白工',
    /if \(!kept\.length\) \{ await saveTemplateText\(text\); return; \}/.test(app));
  ok('選單裡不再有那顆按鈕', !app.includes('tpl-record'));
  // 已經在存檔那一步說要存了，整理完再問一次等於要人確認自己剛做的決定。
  ok('從存檔流程來的整理，寫完直接命名存檔',
    /const asked = state\.procedureRun === 'save'/.test(app));
  // 整段貼進 350px 的輸入框，使用者連自己打的字都看不到。
  ok('套用是掛一張引用晶片，不是整段貼進輸入框',
    app.includes('function attachTemplate(') && /state\.attachedTemplate = tpl/.test(app));
  // AC-42 要求可再修改——點晶片就展開成全文。
  ok('晶片點得開，展開成可編輯的全文', app.includes('function expandTemplate('));
  ok('送出時才把引用展開成文字',
    /const msg = tpl \? \(typed \? `\$\{tpl\.content\}/.test(app));
}

// ── 選單裡的動作列不可以被當成資料來源停用 ─────────────────────────────────
//
// 「常用指令」第一版整列是灰的，點不動。它跟「上傳檔案」一樣掛在 ＋ 選單裡，
// 但 renderSourceMenu 把每一個 .src-item 都當成資料來源，查 state.available
// 查不到就 disabled = true，title 還會變成「undefined 尚未連結」。
// 現在改成看 ARIA role——會切換來源的是 menuitemcheckbox，會做事的是 menuitem。
{
  const html = fs.readFileSync(new URL('../src/taskpane/taskpane.html', import.meta.url), 'utf8');
  const app = fs.readFileSync(new URL('../src/taskpane/app.js', import.meta.url), 'utf8');

  const rows = [...html.matchAll(/<button[^>]*class="src-item"[^>]*>/g)].map(m => m[0]);
  const actions = rows.filter(r => /role="menuitem"/.test(r));
  ok('＋ 選單裡有動作列（上傳檔案、指令與技能）', actions.length >= 2);
  ok('指令與技能是動作列，不是資料來源',
    actions.some(r => /data-src="cmd"/.test(r)));

  const fn = (app.match(/function renderSourceMenu\(\)[\s\S]*?\n}/) || [''])[0];
  ok('render 會先依 role 放行動作列', fn.includes("!== 'menuitemcheckbox'"));
  // 動作列一旦回到這個清單裡，就等於又把兩個問題混在一起了。
  ok('動作列不是靠寫進 ALWAYS_AVAILABLE 才活著',
    !/ALWAYS_AVAILABLE = new Set\(\[[^\]]*'(file|cmd)'/.test(app));
}

// ── 套用後要能再修改（AC-42）──────────────────────────────────────────────
//
// Filling the composer, not sending it. A template that sent itself would have
// to be perfectly general to be useful, and the specific ones are the useful ones.
{
  const app = fs.readFileSync(new URL('../src/taskpane/app.js', import.meta.url), 'utf8');
  // 套用改成掛一張引用晶片（整段貼進 350px 的輸入框會把使用者自己要打的那句
  // 埋掉），但 AC-42 的「可再修改」靠點開晶片展開成全文來滿足。
  const att = (app.match(/function attachTemplate\(id\)[\s\S]*?\n}/) || [''])[0];
  ok('套用不會直接送出', att.length > 0 && !/\bsend\(/.test(att));
  ok('套用會記下是哪一則，供稽核', att.includes('state.appliedTemplateId = tpl.id'));

  const exp = (app.match(/function expandTemplate\(\)[\s\S]*?\n}/) || [''])[0];
  ok('展開會把全文放進輸入框', exp.includes("$('input')") && exp.includes('tpl.content'));
  ok('展開後游標在最後，方便接著補這次的條件', exp.includes('setSelectionRange'));
  ok('展開會保留使用者已經打的字，不是蓋掉', exp.includes('typed'));
}

console.log(`\n${pass} passed, ${fail} failed`);
if (fail) process.exit(1);
