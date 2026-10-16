/**
 * 產生非正式環境的 manifest。
 *
 * ── 為什麼不能直接改正式那三個檔 ──
 * Office 是用 manifest 裡的 `<Id>` 這個 GUID 認增益集的。同一個 GUID 的兩份
 * manifest，Office 會當成**同一個**增益集——裝了測試版就是覆蓋掉正式版，不是並存。
 * 所以測試區要有自己的 GUID 和自己的名字，使用者的清單裡才會看到兩個分得開的項目。
 *
 * ── GUID 要穩定，不能每次重產都換 ──
 * 隨機產的話，每重跑一次腳本就變成一個「新的增益集」，使用者的清單裡會越積越多，
 * 而且舊的那些還得一個一個手動移除。這裡是從原本的 GUID 加上頻道名稱雜湊出來的：
 * 同一個頻道永遠得到同一個 GUID，不同頻道一定不同。
 *
 * ── 正式區不產 ──
 * `DEPLOY_CHANNEL=prod`（或沒設）的時候直接結束。正式區用的就是 manifests/ 底下
 * 原本那三個檔，這支腳本碰都不碰它們。
 *
 * 用法：npm run manifests
 */
import fs from 'fs';
import path from 'path';
import crypto from 'crypto';
import { fileURLToPath } from 'url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SRC = path.join(ROOT, 'manifests');

function readEnv(file) {
  if (!fs.existsSync(file)) return {};
  const out = {};
  for (const line of fs.readFileSync(file, 'utf-8').split('\n')) {
    const m = line.match(/^\s*([A-Z0-9_]+)\s*=\s*(.*)$/i);
    if (!m) continue;
    out[m[1]] = m[2].replace(/\s+#.*$/, '').trim().replace(/^["']|["']$/g, '');
  }
  return out;
}

const env = { ...readEnv(path.join(ROOT, '.env')), ...process.env };
const channel = String(env.DEPLOY_CHANNEL || 'prod').trim().toLowerCase();

if (channel === 'prod' || channel === 'production' || !channel) {
  console.log('DEPLOY_CHANNEL 是 prod（或沒設）——正式區直接用 manifests/ 底下原本的檔案，沒有產生任何東西。');
  process.exit(0);
}

const KNOWN = { test: '測試', staging: '預備', dev: '開發', uat: '驗收' };
const suffix = String(env.DEPLOY_LABEL || '').trim() || KNOWN[channel] || channel;

/**
 * 同一個 GUID + 同一個頻道 = 同一個新 GUID。
 *
 * 取雜湊的前 16 個位元組排成 UUID 的格式，並把版本位元設成 4——Office 只要求是個
 * 合法的 GUID，但照規矩寫可以避免某些驗證工具抱怨。
 */
function derivedGuid(original, ch) {
  const h = crypto.createHash('sha256').update(`${original}|${ch}`).digest();
  const b = Buffer.from(h.subarray(0, 16));
  b[6] = (b[6] & 0x0f) | 0x40;
  b[8] = (b[8] & 0x3f) | 0x80;
  const x = b.toString('hex');
  return `${x.slice(0, 8)}-${x.slice(8, 12)}-${x.slice(12, 16)}-${x.slice(16, 20)}-${x.slice(20)}`;
}

const OUT = path.join(SRC, channel);
fs.mkdirSync(OUT, { recursive: true });

const files = fs.readdirSync(SRC).filter((f) => f.endsWith('.xml'));
if (!files.length) {
  console.error(`${SRC} 底下沒有 .xml，沒有東西可以產。`);
  process.exit(1);
}

const urlBase = String(env.DEPLOY_URL_BASE || '').trim().replace(/\/+$/, '');
let missingIcons = false;

for (const file of files) {
  let xml = fs.readFileSync(path.join(SRC, file), 'utf-8');

  // 1. 換 GUID——沒有這一步，裝測試版等於覆蓋正式版。
  const idMatch = /<Id>([^<]+)<\/Id>/.exec(xml);
  if (!idMatch) { console.error(`${file} 裡找不到 <Id>，跳過。`); continue; }
  const newId = derivedGuid(idMatch[1].trim(), channel);
  xml = xml.replace(/<Id>[^<]+<\/Id>/, `<Id>${newId}</Id>`);

  /**
   * 2. 改顯示名稱——標記放**前面**。
   *
   * 原本是加在後面：「JVision for Word（測試）」。但 Office 的增益集挑選畫面是
   * 磁磚格狀排列，名字會被截斷成「JVision for...」——兩個磁磚長得一模一樣，標記
   * 在被切掉的那一半裡，等於沒做。
   *
   * 放前面的話第一眼就是「【測試】JVi...」，截不截斷都看得出來。
   */
  xml = xml.replace(
    /<DisplayName DefaultValue="([^"]+)"\s*\/>/,
    (_, name) => `<DisplayName DefaultValue="【${suffix}】${name}" />`,
  );

  // 3. 說明也標一下，免得從清單的描述看不出來。
  xml = xml.replace(
    /<Description DefaultValue="([^"]+)"\s*\/>/,
    (_, d) => `<Description DefaultValue="【${suffix}版本，請勿用於正式文件】${d}" />`,
  );

  /**
   * 4. 圖示換成有標記的那一組。
   *
   * 增益集的挑選畫面是磁磚，名字會被截斷——光靠名字分辨不夠可靠。assets/<頻道>/
   * 底下那幾個圖示在下緣壓了一條橘色，縮到 16 像素或灰階列印都還看得出來。
   */
  /**
   * 只有在那組圖示真的存在時才改寫路徑。
   *
   * ── 不檢查的代價 ──
   * 這一行原本無條件改寫，於是 `localtest` 這種沒有人做過圖示的頻道，manifest 會指向
   * assets/localtest/icon-32.png——一個不存在的檔案。Office 每次載入增益集都去要那四個
   * 檔、每次都 404，清單裡顯示的是空白圖示。
   *
   * 功能不受影響，所以它可以安靜地壞很久——而那正是最該被擋下來的那種問題。
   *
   * 沒有就維持原本的路徑：圖示跟正式區一樣（看不出是測試版），但至少看得見。
   * 要有標記的話在 assets/<頻道>/ 放四個檔就會自動接上。
   */
  const iconDir = path.join(ROOT, 'assets', channel);
  if (fs.existsSync(iconDir)) {
    xml = xml.replace(/assets\/(icon-\d+\.png)/g, `assets/${channel}/$1`);
  } else {
    missingIcons = true;
  }

  /**
   * 5. 功能區按鈕上的字。
   *
   * 「上面這裡的 icon 好像很難改」——圖示其實跟著上面那條規則換掉了（功能區用的也是
   * assets/icon-*.png）。沒換到的是**按鈕的文字**：那是 `<bt:String>` 的資源，跟
   * DisplayName 是兩回事，所以兩顆按鈕都還叫「JVision」。
   *
   * 功能區的空間很窄，塞不下「【測試】」，所以用一個字：「JVision·測」。兩顆按鈕
   * 並排的時候，差一個字就分得出來。
   */
  const short = suffix.slice(0, 1);
  xml = xml.replace(
    /<bt:String id="(JVisionGroup\.Label|TaskpaneButton\.Label)" DefaultValue="([^"]+)"\s*\/>/g,
    (_, id, v) => `<bt:String id="${id}" DefaultValue="${v}·${short}" />`,
  );
  // 提示文字和「已安裝」那句也標一下——滑鼠停留和安裝完的提示都是會被讀到的地方。
  xml = xml.replace(
    /<bt:String id="GetStarted\.Title" DefaultValue="([^"]+)"\s*\/>/,
    (_, v) => `<bt:String id="GetStarted.Title" DefaultValue="【${suffix}】${v}" />`,
  );
  xml = xml.replace(
    /<bt:String id="TaskpaneButton\.Tooltip" DefaultValue="([^"]+)"\s*\/>/,
    (_, v) => `<bt:String id="TaskpaneButton.Tooltip" DefaultValue="【${suffix}版本】${v}" />`,
  );

  // 6. 主機位址。只有在指定了 DEPLOY_URL_BASE 的時候才換——沒指定就維持原樣，只有在指定了 DEPLOY_URL_BASE 的時候才換——沒指定就維持原樣，
  //    因為在本機測試的時候那個位址本來就是對的。
  if (urlBase) {
    xml = xml.replace(/https:\/\/localhost:3443/g, urlBase);
  }

  fs.writeFileSync(path.join(OUT, file), xml, 'utf-8');
  console.log(`  ${file}  → Id ${newId}`);
}

console.log(`\n已產生 ${files.length} 份 manifest 到 manifests/${channel}/`);
console.log(`  名稱後面會加上「（${suffix}）」，說明開頭加上「【${suffix}版本，請勿用於正式文件】」`);
if (missingIcons) {
  console.log(`  assets/${channel}/ 不存在，所以圖示路徑沒有改寫——`
    + '清單裡的圖示會跟正式區一樣。');
  console.log(`  要有「${suffix}」標記的話，在 assets/${channel}/ 放 icon-16/32/64/80.png。`);
}
if (!urlBase) {
  console.log('  位址維持 https://localhost:3443——要部署到別台機器請設 DEPLOY_URL_BASE。');
}
console.log('\n給使用者的是 manifests/' + channel + '/ 底下這幾份；'
  + '正式區的檔案一個字都沒動。');
