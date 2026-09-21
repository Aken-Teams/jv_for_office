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
  const menu = (html.match(/<div id="tpl-menu"[\s\S]*?<\/div>\s*<!--/) || [''])[0];
  ok('常用指令選單在窗格裡', menu.length > 0);
  ok('選單裡沒有分享、匯出或上架',
    menu.length > 0 && !/(分享|匯出|上架|share|export|publish)/i.test(menu));
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
  ok('＋ 選單裡有動作列（上傳檔案、常用指令）', actions.length >= 2);
  ok('常用指令是動作列，不是資料來源',
    actions.some(r => /data-src="tpl"/.test(r)));

  const fn = (app.match(/function renderSourceMenu\(\)[\s\S]*?\n}/) || [''])[0];
  ok('render 會先依 role 放行動作列', fn.includes("!== 'menuitemcheckbox'"));
  // 動作列一旦回到這個清單裡，就等於又把兩個問題混在一起了。
  ok('動作列不是靠寫進 ALWAYS_AVAILABLE 才活著',
    !/ALWAYS_AVAILABLE = new Set\(\[[^\]]*'(file|tpl)'/.test(app));
}

// ── 套用後要能再修改（AC-42）──────────────────────────────────────────────
//
// Filling the composer, not sending it. A template that sent itself would have
// to be perfectly general to be useful, and the specific ones are the useful ones.
{
  const app = fs.readFileSync(new URL('../src/taskpane/app.js', import.meta.url), 'utf8');
  const fn = (app.match(/function applyTemplate\(id\)[\s\S]*?\n}/) || [''])[0];
  ok('套用是填進輸入框', fn.includes("$('input')") && fn.includes('input.value = tpl.content'));
  ok('套用不會直接送出', fn.length > 0 && !/\bsend\(/.test(fn));
  ok('游標移到最後，方便接著補這次的條件', fn.includes('setSelectionRange'));
}

console.log(`\n${pass} passed, ${fail} failed`);
if (fail) process.exit(1);
