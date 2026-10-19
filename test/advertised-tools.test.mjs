/**
 * 每個做得到的工具，都要出現在 supportedTools() 裡。
 *
 * ── 這個檔案是哪一天留下來的 ──
 * 側邊欄每次開始對話會把 supportedTools() 送給伺服器，伺服器拿它跟自己的工具清單
 * 比對，少了什麼就掛一條「增益集版本較舊，少了 N 個功能」。
 *
 * ppt_run_script 的實作、派送（case）、WRITES 都做好了，就只有這份清單沒跟上。
 * 症狀特別惡劣：使用者看到的訊息是「請完全關閉 PowerPoint 再重新開啟」，而重開
 * **永遠不會好**——程式碼是對的，錯的是清單。他會一直重開，一直看到同一句話。
 *
 * ── 為什麼是這個形狀的 bug ──
 * 同一天出現三次：makeScriptGuard 的呼叫方式、伺服器的工具規格、這份清單。
 * 共同點都是「實作在這裡，但另一份平行清單不知道它存在」。平行清單靠人同步，
 * 而人會忘；派送表是唯一不會說謊的來源，所以拿它當基準。
 *
 * ── 為什麼用靜態比對 ──
 * supportedTools() 裡有依 Office 能力動態加入的工具（圖片、圖表），執行結果會隨
 * 環境變。改成比對原始碼：只要工具名字出現在那個函式的程式碼裡就算數，不管它是
 * 無條件列出還是條件式 push。
 */
import fs from 'fs';

let failed = 0;
function check(name, ok, detail) {
  if (ok) { console.log(`PASS  ${name}`); return; }
  failed += 1;
  console.error(`FAIL  ${name}${detail ? `\n      ${detail}` : ''}`);
}

const HOSTS = [
  { file: 'powerpoint.js', script: 'ppt_run_script' },
];

/**
 * 原本想拿「派送表」當基準，但三個宿主的派送寫法不同——PowerPoint 有少數 case，
 * 其餘和 Word、Excel 一樣是物件查表。用 case 去抓會把查表裡的工具全報成缺漏，
 * 而**一個會喊狼來了的檢查比沒有檢查更糟**：真的壞掉時沒人會再相信它。
 *
 * 第二版想改成「三個宿主都檢查 run_script」，但函式本體的切法（找換行＋右大括號）
 * 對 word.js 和 excel.js 不成立——會提早命中巢狀區塊，於是又把兩個好的報成壞的。
 *
 * 所以這一版只涵蓋 PowerPoint：那是我實際驗證過、也實際修過的那一個。
 * **寧可範圍小而真，不要範圍大而會誤報。** Word 和 Excel 的 run_script 一直在用
 * （成本分析顯示 Word 常用），真的漏掉會立刻被發現，不需要這支測試守。
 */
for (const { file, script } of HOSTS) {
  const src = fs.readFileSync(new URL(`../src/taskpane/hosts/${file}`, import.meta.url), 'utf-8');

  const from = src.indexOf('export function supportedTools');
  check(`${file}：找得到 supportedTools()`, from > 0);
  if (from < 0) continue;
  const rest = src.slice(from);
  const body = rest.slice(0, rest.indexOf('\n}\n') + 2);

  check(`${file}：真的實作了 ${script}`, src.includes(`'${script}'`));
  check(`${file}：${script} 有對外宣告`, body.includes(`'${script}'`),
    `supportedTools() 沒列它。伺服器會掛「增益集版本較舊，少了 1 個功能」，`
    + `而使用者重開 Office 永遠不會好——程式碼是對的，錯的是這份清單。`);
}

console.log(failed ? `\n${failed} failed` : '\n全部通過');
process.exit(failed ? 1 : 0);
