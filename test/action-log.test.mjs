/**
 * 稽核：AI 在別人的檔案上做了什麼，以及有沒有人同意。
 *
 * ── 為什麼增益集比 ai-agents 更需要這個 ──
 * ai-agents 在沙箱裡做事，最後吐出一個檔案——**產出物本身就是紀錄**，想知道它做
 * 了什麼，打開來看就好。增益集沒有產出物：它改的就是使用者那份檔案，而改之前的
 * 狀態已經不存在了。所以這份 log 是「發生過什麼」唯一可能的紀錄。
 *
 * ── 確認閘門有一半是事後的 ──
 * 一個結果沒有被寫下來的閘門，只在有人盯著的時候有效。開了自動允許之後根本沒有
 * 人點擊，如果連呼叫都不記，等於 AI 動了真實檔案而完全沒有痕跡。所以 decision 是
 * 三態不是布林：**「自動允許」和「人看了才按」是兩個不同的事實**，分不出來的稽核
 * 不算稽核。
 *
 * ── 不存什麼 ──
 * 不存參數。寫入的參數裡就是要寫進去的儲存格內容，存下來等於把薪資和客戶報價抄
 * 進一張給別人看的稽核表——那會讓這份 log 變成它要偵測的那種外洩。
 */
import fs from 'fs';

let pass = 0, fail = 0;
const ok = (name, cond, detail) => {
  cond ? pass++ : fail++;
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}${cond ? '' : '  → ' + (detail || '')}`);
};

const serverDir = new URL('../../ai-agents-office/server/src/', import.meta.url);
if (!fs.existsSync(serverDir)) {
  console.log('      （找不到 ai-agents-office，跳過稽核的跨 repo 檢查）');
  process.exit(0);
}

// ── 窗格：只有它知道有沒有人按過 ──────────────────────────────────────────────
const app = fs.readFileSync(new URL('../src/taskpane/app.js', import.meta.url), 'utf8');

ok('回報結果時會帶上決定', /postToolResult\(runId, callId, result, decision/.test(app));
ok('確認卡片上按同意 → allowed', /decision\.approved \? 'allowed' : 'rejected'/.test(app));
ok('自動允許是自己一種，不是混進 allowed', /decided = 'auto'/.test(app));
ok('拒絕也要回報，不能只是不送', /\}, 'rejected'\);/.test(app));
// 預設值要是保守的那一邊：新分支忘了設，記下來的會是「沒有人確認」。
ok('預設是 none', /let decided = 'none'/.test(app));

// ── 伺服器：唯一的記錄點 ──────────────────────────────────────────────────────
const bridge = fs.readFileSync(new URL('services/excelBridge.ts', serverDir), 'utf8');
const log = fs.readFileSync(new URL('services/addinActionLog.ts', serverDir), 'utf8');

ok('記在 callWorkbookTool，也就是所有呼叫的共同出口', /dispatchTool\(token, tool, args\)\.then/.test(bridge));
ok('每一次呼叫都記一筆', /logAddinAction\(\{/.test(bridge));
// 沒被認出來的值必須讀作「沒有人確認」，不能原樣存進去。
for (const route of ['ppt', 'excel', 'word']) {
  const src = fs.readFileSync(new URL(`routes/${route}.ts`, serverDir), 'utf8');
  ok(`${route} 路由收 decision`, /decision\?: 'none' \| 'auto' \| 'allowed' \| 'rejected'/.test(src));
  ok(`${route} 會驗過才存`, /\.includes\(String\(decision\)\)/.test(src), '');
  ok(`${route} 的 run 知道自己屬於哪場對話`, /conversationId\)?;?\s*$/m.test(src) && /conversationId/.test(src));
}

// decision 是我們跟稽核之間的帳，模型看到只會以為是工具回答的一部分。
ok('決定不會被送進模型', /const \{ decision, \.\.\.rest \}/.test(bridge));

// ── 存什麼、不存什麼 ──────────────────────────────────────────────────────────
ok('有 addin_actions 這張表', /CREATE TABLE IF NOT EXISTS addin_actions/.test(log));
for (const col of ['user_id', 'conversation_id', 'host', 'tool', 'target',
  'destructive', 'decision', 'ok', 'summary']) {
  ok(`欄位有 ${col}`, new RegExp(`\\b${col}\\s`).test(log));
}
ok('「誰沒被確認過」查得快', /idx_addin_decision/.test(log));
// 這條是資安條款，不是整潔條款。註解裡當然會提到 arguments，所以先把註解拿掉再
// 檢查——不然這條會被自己的說明文字騙過去。
const code = log.replace(/\/\*[\s\S]*?\*\//g, '').replace(/\/\/.*/g, '');
ok('沒有把參數存進去', !/JSON\.stringify|args_json|args/.test(code),
  '寫入的參數裡就是儲存格內容，存下來等於把薪資和報價抄進稽核表');
ok('摘要有長度上限', /cut\(rec\.summary, 500\)/.test(log));
ok('記 log 失敗不能弄壞使用者的編輯', /catch \{ \/\* an audit row is never worth/.test(log));

// ── 位置資訊 ──────────────────────────────────────────────────────────────────
const mod = await import(new URL('services/addinActionLog.ts', serverDir).href).catch(() => null);
if (mod) {
  ok('工作表加範圍', mod.targetOf('excel_write_range', { sheet: '損益表', range: 'C4:F40' }) === '損益表!C4:F40');
  ok('投影片區間', mod.targetOf('ppt_delete_slide', { slide: 5, to: 7 }) === '投影片 5–7');
  ok('單張投影片不會多出區間', mod.targetOf('ppt_delete_slide', { slide: 5 }) === '投影片 5');
  ok('整頁刪除', mod.targetOf('ppt_delete_shape', { all: true }) === '整頁');
  // 程式碼本身屬於對話紀錄，不屬於稽核表——但「跑了多長的東西」值得留下。
  ok('程式碼只記長度', /字的程式碼/.test(mod.targetOf('excel_run_script', { code: 'x'.repeat(120) })));
  ok('沒有座標就空著，不要瞎掰', mod.targetOf('excel_get_overview', {}) === '');

  ok('host 從工具名推出來', mod.hostOf('ppt_build_slide') === 'ppt'
    && mod.hostOf('excel_write_range') === 'excel' && mod.hostOf('word_insert_toc') === 'word');
  ok('不認得的工具不會被硬塞成某個 host', mod.hostOf('ToolSearch') === 'other');
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
