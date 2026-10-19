/**
 * tools/ 底下的 .ps1 必須存成「UTF-8 帶 BOM」。
 *
 * ── 這個檔案是哪一天留下來的 ──
 * install-channels.ps1 寫好、語法沒問題、在 Git Bash 裡看起來完全正常，然後
 * `npm run install-channels` 吐出十個解析錯誤：
 *
 *     運算式或陳述式中有未預期的 'test' = 'JVision' 語彙基元。
 *     '[' 後面遺漏型別名稱。
 *     字串遺漏結尾字元: "。
 *
 * 沒有一句話提到編碼。真正的原因是 **Windows PowerShell 5.1 在沒有 BOM 的時候，
 * 會用系統 ANSI 字碼頁去讀 .ps1**（這台是 cp950）。於是每個中文字都被解成兩個
 * 亂碼位元組，其中偶爾會湊出一個 `'` 或 `"`，把字串的結尾吃掉——錯誤訊息因此指向
 * 引號和括號，離真正的原因有十萬八千里。
 *
 * PowerShell 7 沒有這個問題（它預設 UTF-8），所以這個 bug 只在使用者那台出現，
 * 而那正是最糟的那種：寫的人看不到，跑的人看不懂。
 *
 * ── 為什麼用測試擋，而不是寫進文件 ──
 * 「存檔記得加 BOM」是一條沒有人會記得的規則，而且任何編輯器、任何腳本、任何
 * 一次 git 操作都可能把它拿掉。規則靠人遵守，機制不用。
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const TOOLS = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'tools');
const BOM = Buffer.from([0xef, 0xbb, 0xbf]);

let failed = 0;
function check(name, ok, detail) {
  if (ok) { console.log(`PASS  ${name}`); return; }
  failed += 1;
  console.error(`FAIL  ${name}${detail ? `\n      ${detail}` : ''}`);
}

const files = fs.readdirSync(TOOLS).filter(f => f.toLowerCase().endsWith('.ps1'));
check('tools/ 底下找得到 .ps1', files.length > 0, `在 ${TOOLS} 一個都沒有`);

for (const f of files) {
  const buf = fs.readFileSync(path.join(TOOLS, f));
  const hasBom = buf.subarray(0, 3).equals(BOM);

  // 純 ASCII 的腳本沒有 BOM 也不會壞——ANSI 和 UTF-8 對 ASCII 的解讀相同。
  // 所以只要求「有非 ASCII 字元的檔案」要帶 BOM，不然這個測試會變成一條
  // 跟壞掉無關的潔癖規則，而那種測試遲早會被人加例外繞過去。
  const nonAscii = buf.subarray(hasBom ? 3 : 0).some(b => b > 0x7f);

  if (!nonAscii) {
    check(`${f}：純 ASCII，不需要 BOM`, true);
    continue;
  }
  check(`${f}：有中文，必須帶 UTF-8 BOM`, hasBom,
    'Windows PowerShell 5.1 會改用系統 ANSI 字碼頁讀它，中文變亂碼、引號被吃掉，'
    + '然後報一堆跟編碼無關的語法錯誤。');
}

console.log(failed ? `\n${failed} failed` : '\n全部通過');
process.exit(failed ? 1 : 0);
