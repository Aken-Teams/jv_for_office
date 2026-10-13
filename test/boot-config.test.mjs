/**
 * 部署設定必須在開機時讀，不能只在登入畫面讀。
 *
 * ── 這個檔案是哪一天留下來的 ──
 * loadConfig()（讀 config.json，決定要連哪一台伺服器、掛不掛測試區橫幅）當初只寫在
 * showLogin() 裡。而開機的最後一行是「有 token 就直接進聊天，否則才進登入畫面」。
 *
 * 於是**已經登入的人**走的是另一條路，config.json 從頭到尾沒被讀過，state.apiBase
 * 停在檔案頂端那個寫死的 DEFAULT_API_BASE —— http://localhost:12054，開發用的。
 *
 * 部署到別台機器之後的症狀是：側邊欄去打一個那台沒有的位址，錯誤訊息長得跟網路故障
 * 一模一樣；登出再登入就好了，因為登入畫面才會讀設定。這個「要重新登入才能用」的線索
 * 花了兩輪才指回真正的原因，因為它看起來像是 localStorage 留了髒資料。
 *
 * ── 為什麼用靜態檢查 ──
 * 窗格沒有建置步驟也沒有 DOM 測試環境，Office.onReady 跑不起來。但這個 bug 的形狀是
 * 「呼叫點在錯的地方」，那是讀原始碼就看得出來的事：只要 loadConfig() 的呼叫又只剩
 * 登入畫面那一處，它就回來了。
 */
import fs from 'fs';

const src = fs.readFileSync(new URL('../src/taskpane/app.js', import.meta.url), 'utf-8');
let failed = 0;
function check(name, ok, detail) {
  if (ok) { console.log(`PASS  ${name}`); return; }
  failed += 1;
  console.error(`FAIL  ${name}${detail ? `\n      ${detail}` : ''}`);
}

/**
 * 取**最後一次**出現，不是第一次。
 *
 * 上面那段說明 bug 的註解裡就引用了被測的那行程式碼；用 indexOf 會匹配到註解裡的
 * 那一份，然後檢查永遠是紅的——而原因跟被測的東西毫無關係。這個測試自己踩過一次。
 */
function lastMatch(text, re) {
  let at = -1;
  for (const m of text.matchAll(new RegExp(re.source, 'g'))) at = m.index;
  return at;
}

const readyAt = src.indexOf('Office.onReady(');
check('app.js 裡有 Office.onReady', readyAt !== -1);
const readyBody = readyAt === -1 ? '' : src.slice(readyAt);

check(
  '開機流程會讀 config.json（await loadConfig()）',
  /await\s+loadConfig\s*\(\s*\)/.test(readyBody),
  'loadConfig() 不在 Office.onReady 裡面。已經登入的人就不會讀到部署設定，'
  + 'apiBase 會停在 DEFAULT_API_BASE（http://localhost:12054）。',
);

const loginAt = src.indexOf('function showLogin(');
check('app.js 裡有 showLogin()', loginAt !== -1);
const loginBody = loginAt === -1 ? '' : src.slice(loginAt, src.indexOf('\n}', loginAt));

check(
  '登入畫面不是唯一讀設定的地方',
  !/\bloadConfig\s*\(\s*\)/.test(loginBody),
  '只有登入畫面會讀設定——那正是當初的 bug。',
);

// 讀設定要排在決定「進聊天還是進登入」之前，不然已經登入的人還是會先用到舊位址。
const branchAt = lastMatch(readyBody, /if\s*\(\s*state\.token\s*\)\s*showChat\s*\(/);
const configAt = lastMatch(readyBody, /await\s+loadConfig\s*\(\s*\)/);
check(
  '讀設定排在 showChat()/showLogin() 的分支之前',
  configAt !== -1 && branchAt !== -1 && configAt < branchAt,
  `loadConfig 在第 ${configAt} 字元，分支在第 ${branchAt} 字元。`,
);

// 橫幅跟著設定走：markChannel 是 loadConfig 的一部分，所以上面那條通過的話，
// 測試區的橘色橫幅在「已經登入」的路徑上也會出現——Excel 那張截圖就是缺了它。
check(
  '頻道橫幅跟著設定一起掛上（markChannel 在 loadConfig 裡）',
  /async function loadConfig[\s\S]*?markChannel\s*\(\s*\)[\s\S]*?\n}/.test(src),
  'markChannel() 不在 loadConfig() 裡，測試區橫幅會跟部署設定走散。',
);

if (failed) { console.error(`\n${failed} 項沒過。`); process.exit(1); }
console.log('\n開機就會讀部署設定。');
