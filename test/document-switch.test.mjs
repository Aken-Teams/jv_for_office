/**
 * What has to be forgotten when a different document arrives.
 *
 * ── The bug ──
 * SharedRuntime keeps the pane alive across workbooks — that is the whole point
 * of it — but `documentKey` was only minted when it was EMPTY, so it kept the
 * first workbook's id for as long as Excel stayed open. Open A, ask something,
 * open B: the header said B while the key still said A. B's messages were filed
 * under A's history, B's history came up blank, and A's thread ended up split.
 * The report was 「同一個檔案，對話紀錄沒有存起來」.
 *
 * The undo stack is the one that could have destroyed data. Its snapshots hold
 * cell values from the old workbook, and restoreSnapshot writes into whichever
 * workbook is ACTIVE — so pressing 復原 after switching would have written A's
 * old values into B.
 *
 * ── Why this is a source test ──
 * Reproducing it needs two workbooks in one Excel process. What can be checked
 * without Excel is that the reset exists, that it covers everything scoped to a
 * document, and that both readers of the key go through the switch check rather
 * than trusting a stale value.
 */
import fs from 'fs';

const app = fs.readFileSync(new URL('../src/taskpane/app.js', import.meta.url), 'utf8');

let pass = 0, fail = 0;
const ok = (name, cond) => {
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}`);
  if (cond) pass++; else fail++;
};

const fn = (app.match(/function documentSwitched\(\)[\s\S]*?\n}/) || [''])[0];
ok('有「換了文件」這條路', fn.length > 0);

// Each of these is scoped to one document. A stale one is not untidy, it is wrong.
ok('換文件會丟掉 documentKey', /state\.documentKey = ''/.test(fn));
ok('換文件會丟掉 sessionId——那是另一份文件那場對話的 CLI session',
  /state\.sessionId = null/.test(fn));
ok('換文件會清空復原堆疊——快照是舊活頁簿的值，還原會寫進新的那一份',
  /state\.undoStack = \[\]/.test(fn));
ok('換文件會清空錄製紀錄', /state\.recorded = \[\]/.test(fn));
ok('清空復原堆疊之後有更新按鈕狀態', /updateUndoButton\(\)/.test(fn));

const refresh = (app.match(/async function refreshDocumentName\(\)[\s\S]*?\n}/) || [''])[0];
ok('改名不算換文件的第一次比較（沒有前一個名字就不觸發）',
  /if \(was && state\.documentName && state\.documentName !== was\)/.test(refresh));

// Renaming also lands in documentSwitched, and that is safe: clearing the key
// only forces a re-read, and a renamed file still carries the same custom
// property — so it comes back with the same id and keeps its history.
ok('重新取 key 之前會先檢查有沒有換文件',
  /async function ensureDocumentKey\(\)[\s\S]*?await refreshDocumentName\(\)/.test(app));

// 兩個讀 key 的地方都必須走同一條路，否則其中一個又會信任過期的值。
ok('歷史紀錄用 ensureDocumentKey', /const key = \(await ensureDocumentKey\(\)\)/.test(app));
ok('送出訊息也用 ensureDocumentKey', /await ensureDocumentKey\(\);/.test(app));
ok('沒有人再直接取 documentKey 了',
  (app.match(/host\.getDocumentKey\(\)/g) || []).length === 1);

// ── 窗格送的參數名，伺服器要收得到 ─────────────────────────────────────────
//
// excel.ts 的歷史紀錄路由讀的是 `workbookKey`，而窗格早就改送 `documentKey` 了
// ——word.ts 和 ppt.ts 當初都跟著改，只有 excel.ts 漏掉。於是 key 永遠是空的，
// 路由的空值防呆直接回空陣列，每一份活頁簿都顯示「還沒有對話紀錄」。
// 資料一直好端端在資料表裡，只是沒有任何路徑碰得到它。
//
// 這種錯不會噴任何東西：兩邊各自都合法，只是講不同的話。
{
  const serverDir = new URL('../../ai-agents-office/server/src/', import.meta.url);
  if (!fs.existsSync(serverDir)) {
    console.log('      （找不到 ai-agents-office，跳過參數名的跨 repo 檢查）');
  } else {
    // 窗格用什麼名字送。
    const sent = [...app.matchAll(/conversations\?(\w+)=/g)].map(m => m[1]);
    ok('窗格送的是 documentKey', sent.length > 0 && sent.every(n => n === 'documentKey'));

    for (const host of ['excel', 'word', 'ppt']) {
      const src = fs.readFileSync(new URL(`routes/${host}.ts`, serverDir), 'utf8');
      const route = (src.match(/router\.get\('\/conversations',[\s\S]*?\n}\);/) || [''])[0];
      ok(`${host}: 歷史紀錄路由收得到 documentKey`,
        /req\.query\.documentKey/.test(route));
      // 舊拼法要留著：manifest 會把窗格的 JS 釘住，沒重開 Excel 的人現在還在送舊的。
      ok(`${host}: 舊的 workbookKey 也還收`, /req\.query\.workbookKey/.test(route));
    }
  }
}

console.log(`\n${pass} passed, ${fail} failed`);
if (fail) process.exit(1);
