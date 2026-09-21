/**
 * 跨增益集交接 (PRD 3.5, AC-44/45/46).
 *
 * ── The constraint the whole design answers to ──
 * The Office add-in platform does not let one host drive another. Excel's pane
 * cannot open PowerPoint; PowerPoint's pane cannot reach back into a workbook.
 * So 「turn this analysis into a deck」 is necessarily two actions with a parcel
 * in between, and every check here is about that parcel: who may read it, how
 * long it lives, and — the one that decides whether the feature is any good —
 * that it carries CONTENT rather than a finished file.
 *
 * That last one is not a detail. AC-44 requires the generated slides to inherit
 * the deck's own layouts and theme colours, and AC-45 requires Word's changes to
 * arrive as tracked changes. A server has no open document and can do neither.
 * Only the receiving add-in can, with its native tools, in the file that is
 * actually open — so the parcel must never become a download.
 */
import fs from 'fs';

let pass = 0, fail = 0;
const ok = (name, cond) => {
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}`);
  if (cond) pass++; else fail++;
};

const app = fs.readFileSync(new URL('../src/taskpane/app.js', import.meta.url), 'utf8');
const html = fs.readFileSync(new URL('../src/taskpane/taskpane.html', import.meta.url), 'utf8');

const serverDir = new URL('../../ai-agents-office/server/src/', import.meta.url);
if (!fs.existsSync(serverDir)) {
  console.log('      （找不到 ai-agents-office，跳過交接包的跨 repo 檢查）');
  process.exit(0);
}
const routeSrc = fs.readFileSync(new URL('routes/handoffs.ts', serverDir), 'utf8');
const route = routeSrc.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
const db = fs.readFileSync(new URL('db.ts', serverDir), 'utf8');
const spec = fs.readFileSync(new URL('services/excelToolSpec.ts', serverDir), 'utf8');
const bridge = fs.readFileSync(new URL('services/excelBridge.ts', serverDir), 'utf8');

// ── 僅本人可見（AC-46）────────────────────────────────────────────────────
{
  const selects = [...route.matchAll(/SELECT[\s\S]*?FROM handoffs[^`']*/g)].map(m => m[0]);
  ok('每一個查詢不是帶 user_id，就是那唯一一個擁有者探測',
    selects.length >= 2 && selects.every(q => /user_id = \?/.test(q)
      || /^SELECT user_id FROM handoffs WHERE id = \?$/.test(q.trim())));
  ok('寫入與刪除都帶 user_id',
    [...route.matchAll(/(UPDATE|DELETE FROM) handoffs[^`']*/g)]
      .filter(m => !/expires_at < NOW\(\)/.test(m[0]))
      .every(m => /user_id = \?/.test(m[0])));
  ok('用 id 取整列只有 mine() 一個入口',
    (route.match(/SELECT \* FROM handoffs/g) || []).length === 1);
  ok('越權存取記進稽核', route.includes('嘗試存取他人的交接包'));
}

// ── 24 小時，而且是真的刪掉（AC-46）──────────────────────────────────────
{
  ok('有效期 24 小時', /TTL_HOURS = 24/.test(route));
  // 逾期只是標記的話，那些數字還留在資料表裡。它們是別人活頁簿裡的真實金額。
  ok('逾期是刪除，不是標記', /DELETE FROM handoffs WHERE expires_at < NOW\(\)/.test(route));
  ok('每次讀取都會清掉逾期的', /async function sweep\(\)[\s\S]*?DELETE FROM handoffs/.test(route));

  const table = (db.match(/CREATE TABLE IF NOT EXISTS handoffs[\s\S]*?\)\s*ENGINE/) || [''])[0];
  ok('資料表存在', table.length > 0);
  ok('綁使用者，使用者刪除時一起刪',
    /user_id\s+VARCHAR\(36\)\s+NOT NULL/.test(table) && /ON DELETE CASCADE/.test(table));
  ok('有 expires_at 欄位', /expires_at\s+DATETIME NOT NULL/.test(table));
  ok('沒有任何可以擴散出去的欄位',
    table.length > 0 && !/\b(shared|visibility|team_id|public)\b/i.test(table));
}

// ── 清單只給中繼資料，內容要等接受才交出去 ────────────────────────────────
//
// 開個側邊欄就把別人活頁簿裡的數字拉進一份還沒決定要用它的文件，是不必要的。
{
  const list = (route.match(/router\.get\('\/',[\s\S]*?\n\}\);/) || [''])[0];
  ok('清單不回傳 payload', list.length > 0 && !/payload/.test(list));
  const accept = (route.match(/router\.post\('\/:id\/accept',[\s\S]*?\n\}\);/) || [''])[0];
  ok('接受時才交出 payload', /payload: row\.payload/.test(accept));
  ok('接受後標記已收，避免重複出現', /consumed_at = NOW\(\)/.test(accept));
  ok('已經被接收過的會擋下來', /已經被接收過/.test(accept));
}

// ── Excel 端：打包工具 ────────────────────────────────────────────────────
{
  ok('excel_create_handoff 在工具規格裡', spec.includes("name: 'excel_create_handoff'"));
  ok('列為伺服器端工具，不會送到窗格',
    /EXCEL_SERVER_TOOLS = \[[^\]]*'excel_create_handoff'/.test(spec));
  // 那邊的 agent 看不到這份活頁簿，工具說明必須講清楚 content 要自成一份材料。
  ok('工具說明要求內容自成一份材料', spec.includes('它只看得到你寫進去的東西'));
  ok('工具說明要求附上儲存格引用', spec.includes('[[工作表!C42]]'));
  ok('橋接層自己回答，不經過窗格', /tool\.endsWith\('_create_handoff'\)/.test(bridge));
  ok('來源檔名跟著 run 走', /documentName: string;/.test(bridge));
}

// ── 接收端 ────────────────────────────────────────────────────────────────
{
  ok('窗格頂端有接收橫幅', html.includes('id="handoff-bar"'));
  ok('橫幅預設隱藏', /<div id="handoff-bar"[^>]*hidden/.test(html));
  ok('依 host 取自己的交接包',
    /api\/handoffs\?target=\$\{encodeURIComponent\(host\.meta\.api\)\}/.test(app));
  ok('接受與忽略都有', app.includes('function acceptHandoff(') && app.includes('function ignoreHandoff('));

  // 這一條是整個功能好不好用的分界：接受之後由「這一邊」的 agent 用原生工具
  // 在當下開著的檔案裡生成，而不是伺服器產一個檔給人下載。
  const accept = (app.match(/async function acceptHandoff\(id\)[\s\S]*?\n}/) || [''])[0];
  ok('接受會把內容交給這一邊的 agent 生成', /void send\(/.test(accept));
  ok('明確要求沿用既有版面與樣式', accept.includes('沿用這份檔案既有的版面與樣式'));
  ok('明確要求不要另外建立新檔', accept.includes('不要另外建立新檔'));
  ok('明確要求數字以交接內容為準', accept.includes('不要自行補充或推估'));
  // AC-46：逾期要講出來，不是安靜地失敗。
  ok('逾期會講出來', /addNote\(body\.error \|\| t\('這個交接包已經失效了。'\)/.test(accept));
  // 對話裡顯示的是一句話，完整內容收在雪佛龍底下——跟常用指令同一個處理。
  ok('對話裡不貼整包內容', /display: t\('接受從 \{0\} 帶過來的分析/.test(accept));
  ok('完整內容仍然查得到', /folded: body\.payload/.test(accept));
}

console.log(`\n${pass} passed, ${fail} failed`);
if (fail) process.exit(1);
