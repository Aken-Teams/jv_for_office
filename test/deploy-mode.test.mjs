/**
 * 強茂版（pro-panjit）不顯示模型選單——而且要在「已經登入的人」身上也生效。
 *
 * ── 為什麼需要這個檔案 ──
 * 這個功能有一個特別容易壞的地方：deployMode 來自 /api/auth/mode，而那個端點原本
 * **只在登入畫面被呼叫**（loadAuthMode 寫在 showLogin 裡）。於是「已經登入的人」
 * 這條路徑從頭到尾不知道自己連的是哪一套部署。
 *
 * 症狀會非常刁鑽：登出再登入就正常，平常開啟側邊欄就是看得到選單——而強茂的使用者
 * 裝好之後就一直是登入狀態，也就是說，這個功能會**剛好在唯一需要它的那群人身上失效**。
 *
 * 這跟當初 loadConfig() 的 bug 是同一個形狀（見 boot-config.test.mjs），那次花了
 * 四輪才找到原因。同樣的錯誤不該再花第二次。
 *
 * ── 為什麼用靜態檢查 ──
 * 窗格沒有建置步驟也沒有 DOM 測試環境，Office.onReady 跑不起來。但這個 bug 的形狀是
 * 「呼叫點在錯的地方」，那是讀原始碼就看得出來的事。
 */
import fs from 'fs';

const raw = fs.readFileSync(new URL('../src/taskpane/app.js', import.meta.url), 'utf-8');

let failed = 0;
function check(name, ok, detail) {
  if (ok) { console.log(`PASS  ${name}`); return; }
  failed += 1;
  console.error(`FAIL  ${name}${detail ? `\n      ${detail}` : ''}`);
}

/**
 * 先把註解拿掉再檢查。
 *
 * 上面那幾段說明裡就寫著 'pro-panjit' 和 loadDeployMode()，用原始字串去比對的話，
 * 檢查會被註解滿足——然後在程式碼真的被刪掉的那天依然是綠的。boot-config.test.mjs
 * 踩過這個坑，它的解法是取最後一次出現；這裡直接把註解消掉，更不容易出錯。
 *
 * 換成等長空白而不是刪除，位置比對（誰在誰前面）才不會偏掉。
 */
const src = raw
  .replace(/\/\*[\s\S]*?\*\//g, (m) => ' '.repeat(m.length))
  .replace(/(^|[^:])\/\/[^\n]*/g, (m, p1) => p1 + ' '.repeat(m.length - p1.length));

// ── 1. 這個函式存在，而且真的去打那個端點 ────────────────────────────────────
check('有 loadDeployMode()', /async function loadDeployMode\s*\(/.test(src));
const body = src.slice(src.indexOf('async function loadDeployMode'),
  src.indexOf('async function loadAuthMode'));
check('loadDeployMode 去讀 /api/auth/mode', body.includes('/api/auth/mode'), body.trim().slice(0, 200));
check('loadDeployMode 寫進 state.deployMode', /state\.deployMode\s*=/.test(body));

// ── 2. 開機就呼叫，而且在分支之前 ────────────────────────────────────────────
//
// 這是整個檔案的重點。下面三個位置的先後順序，就是這個 bug 會不會回來的全部。
const branch = src.search(/if\s*\(\s*state\.token\s*\)\s*showChat\(\)/);
check('找得到開機的分支', branch > 0);

const calls = [...src.matchAll(/\bloadDeployMode\s*\(\s*\)/g)].map(m => m.index);
// 定義本身不算呼叫
const invocations = calls.filter(i => !src.slice(Math.max(0, i - 30), i).includes('async function'));
check('loadDeployMode 真的有被呼叫', invocations.length > 0,
  '只定義不呼叫的話，deployMode 永遠是空字串，選單永遠顯示。');
check('在 if (state.token) showChat() 之前就呼叫',
  invocations.some(i => i < branch),
  `呼叫位置=${invocations.join(',')} 分支位置=${branch}——`
  + '在分支之後才讀的話，已登入的人永遠讀不到，而那正是強茂使用者的常態。');

// showLogin 裡面呼叫也可以，但**不能只有那裡**——那就是原本的 bug。
const showLogin = src.indexOf('function showLogin');
const showLoginEnd = src.indexOf('\n}', showLogin);
check('不是只寫在 showLogin() 裡',
  invocations.some(i => i < showLogin || i > showLoginEnd),
  '只在登入畫面讀設定，等於只有剛登入的人拿得到——這就是 loadConfig() 當初的 bug。');

// await，不是 void。沒等的話 renderModelMenu() 會先畫出選單再收掉，閃一下。
const at = invocations.find(i => i < branch);
check('是 await 不是 void（否則選單會先閃一下）',
  /await\s+$/.test(src.slice(Math.max(0, at - 12), at)),
  `呼叫前面是：${JSON.stringify(src.slice(Math.max(0, at - 12), at))}`);

// ── 3. renderModelMenu 擋得住 ────────────────────────────────────────────────
const rm = src.slice(src.indexOf('function renderModelMenu'));
const rmBody = rm.slice(0, rm.indexOf('\n}'));
check('renderModelMenu 檢查 pro-panjit', rmBody.includes("'pro-panjit'"), rmBody.slice(0, 300));
check('renderModelMenu 在算清單之前就擋掉',
  rmBody.indexOf("'pro-panjit'") < rmBody.indexOf('list.length < 2'),
  '要在「只有一個模型」那條判斷之前——否則兩個模型的時候會漏掉。');
check('renderModelMenu 擋掉時把 wrap 藏起來',
  /pro-panjit'\s*\)\s*\{[^}]*wrap\.hidden\s*=\s*true/.test(rmBody), rmBody.slice(0, 400));

// ── 4. 看門狗不能對正常狀態喊狼來了 ──────────────────────────────────────────
const wd = src.slice(src.indexOf('function modelMenuWatchdog'));
const wdBody = wd.slice(0, wd.indexOf('\n}\n'));
check('看門狗認得 pro-panjit 是正常的', wdBody.includes("'pro-panjit'"),
  '少了這一行，每台強茂的機器每次開側邊欄都會在六秒後掛出一條診斷橫幅，'
  + '而那是設計好的行為不是故障。');
check('看門狗在 bootTrouble 之前就 return',
  wdBody.indexOf("'pro-panjit'") < wdBody.indexOf('bootTrouble'), wdBody.slice(0, 400));

// ── 5. 預設值的方向要是安全的那一邊 ──────────────────────────────────────────
check('state.deployMode 預設是空字串（＝顯示選單）',
  /deployMode:\s*''/.test(src),
  '猜錯的方向要選代價小的那邊：多一個選單比少一個功能好解釋。');

console.log(failed ? `\n${failed} failed` : '\n全部通過');
process.exit(failed ? 1 : 0);
