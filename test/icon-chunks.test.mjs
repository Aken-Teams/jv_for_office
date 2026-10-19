/**
 * 增益集圖示的 PNG 裡不能有 pHYs 區塊。
 *
 * ── 這是什麼問題 ──
 * Office 會把檔案抓下來（dev server 的請求紀錄裡是乾乾淨淨的 200、位元組數也對），
 * 然後**拒絕畫出來**，在清單和功能區都顯示一顆通用的六邊形。沒有錯誤訊息，沒有
 * 重試，而且在下次改 manifest 版本號之前不會再去抓一次。
 *
 * 跟一個已知會顯示的 Office 增益集（Claude for Excel，取自 pivot.claude.ai）逐區塊
 * 比對之後，唯一的結構差異是 sharp 會寫出一個 pHYs 區塊（實體像素密度，96 DPI）：
 *
 *     會顯示：  IHDR IDAT IEND
 *     不顯示：  IHDR pHYs IDAT IEND
 *
 * 兩邊都是 32x32、8-bit RGBA、帶透明度，檔案大小差不到 30 位元組。所以透明度從來
 * 不是原因；Office 顯然會讀 pHYs，然後做了某件以「什麼都不畫」收場的事。
 *
 * ── 為什麼需要這個測試 ──
 * make-icons.mjs 產 assets/icon-*.png 的時候會把這些區塊拿掉，但**頻道專用的那幾組**
 * （assets/test/、assets/localtest/、assets/jv/）不是它產的——是人工做好再放進來的。
 * 於是那條規則只保護了一半的檔案，而壞掉的那一半花了整整一個下午、四個錯誤的推論
 * 才找到原因（先後怪過 sideload 機制、網址快取鍵、manifest 版本號）。
 *
 * 真正的教訓是：這個 bug 完全沒有症狀可以讀，只能靠比對。所以它必須被檢查，不能
 * 靠人記得。任何人把一張新圖丟進 assets/ 底下，這個測試就會替他檢查。
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const ASSETS = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', 'assets');

let failed = 0;
function check(name, ok, detail) {
  if (ok) { console.log(`PASS  ${name}`); return; }
  failed += 1;
  console.error(`FAIL  ${name}${detail ? `\n      ${detail}` : ''}`);
}

/** 只保留關鍵區塊是 make-icons.mjs 的規則；這裡列出實際有哪些，好讓失敗訊息說得出話。 */
function chunks(file) {
  const d = fs.readFileSync(file);
  if (d.length < 8 || d[0] !== 0x89 || d[1] !== 0x50) return null;   // 不是 PNG
  const out = [];
  let p = 8;
  while (p + 8 <= d.length) {
    const len = d.readUInt32BE(p);
    const type = d.toString('ascii', p + 4, p + 8);
    out.push(type);
    if (type === 'IEND') break;
    p += 12 + len;
  }
  return out;
}

/** assets/ 底下所有 icon-*.png，含頻道子資料夾。 */
function iconFiles(dir) {
  const found = [];
  for (const e of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, e.name);
    if (e.isDirectory()) found.push(...iconFiles(full));
    else if (/^icon-\d+\.png$/.test(e.name)) found.push(full);
  }
  return found;
}

const files = iconFiles(ASSETS);
check('assets/ 底下找得到圖示', files.length > 0, `在 ${ASSETS} 一個都沒有`);

// manifest 要的四個尺寸。少一個的話 Office 會退回通用圖示，而那又是一個沒有訊息的失敗。
// 64 特別重要：HighResolutionIconUrl 規定必須是 64x64，指到 80 的檔是文件上寫明的違規。
const WANT = [16, 32, 64, 80];

const dirs = [...new Set(files.map(f => path.dirname(f)))].sort();
for (const d of dirs) {
  const rel = path.relative(ASSETS, d) || '（根目錄）';
  const have = files.filter(f => path.dirname(f) === d)
    .map(f => Number(f.match(/icon-(\d+)\.png$/)[1])).sort((a, b) => a - b);
  check(`${rel}：四個尺寸都在`, WANT.every(w => have.includes(w)),
    `有 ${have.join('/')}，少了 ${WANT.filter(w => !have.includes(w)).join('/')}`);
}

for (const f of files) {
  const rel = path.relative(ASSETS, f).replace(/\\/g, '/');
  const list = chunks(f);
  if (!list) { check(`${rel}：是 PNG`, false, '檔頭不是 PNG'); continue; }

  const extra = list.filter(c => !['IHDR', 'PLTE', 'IDAT', 'IEND'].includes(c));
  check(`${rel}：只有關鍵區塊`, extra.length === 0,
    `多了 ${extra.join(', ')} —— Office 會抓到檔案卻畫成通用六邊形。`
    + '跑 npm run icons 重產，或照 make-icons.mjs 的 stripAncillaryChunks 處理過再放進來。');
}

console.log(failed ? `\n${failed} failed` : '\n全部通過');
process.exit(failed ? 1 : 0);
