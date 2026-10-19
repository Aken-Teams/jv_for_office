/**
 * 把 video/scene.html 逐幀截圖，用 ffmpeg 合成 MP4。
 *
 * ── 為什麼是逐幀而不是錄螢幕 ──
 * 錄螢幕錄到的是「當下這台機器跑得多順」：掉幀、字體還沒載完、動畫被背景程式卡一下，
 * 每次錄出來都不一樣，而且重錄一次就要重來一遍。
 *
 * 逐幀的做法是跟頁面要「第 t 秒長什麼樣」，拍下來，再要下一幀。頁面沒有自己的時鐘，
 * 所以慢一點快一點都不影響結果——**同樣的輸入永遠得到同樣的影片**。要把某一段放慢，
 * 改一個數字重跑就好，不必重演一次。
 *
 * ── 為什麼用系統的 Chrome ──
 * puppeteer 會自己下載一份 Chromium（約 150MB）。這台已經有 Chrome 155，用
 * puppeteer-core 接現成的就好；少一份要同步更新的二進位檔。
 *
 * 跑：
 *   npm run video              完整輸出
 *   npm run video -- --to 6    只算前 6 秒（調分鏡時用）
 *   npm run video -- --fps 15  降幀率快速預覽
 */
import fs from 'fs';
import path from 'path';
import { spawn } from 'child_process';
import { fileURLToPath, pathToFileURL } from 'url';
import { createRequire } from 'module';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const SCENE = path.join(ROOT, 'video', 'scene.html');
const FRAMES = path.join(ROOT, 'video', '.frames');
const OUT = path.join(ROOT, 'video', 'JVision-demo.mp4');

const CHROME = [
  'C:/Program Files/Google/Chrome/Application/chrome.exe',
  'C:/Program Files (x86)/Google/Chrome/Application/chrome.exe',
  `${process.env.LOCALAPPDATA || ''}/Google/Chrome/Application/chrome.exe`,
  'C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',
].find(p => p && fs.existsSync(p));

// ── 參數 ─────────────────────────────────────────────────────────────────────
const argv = process.argv.slice(2);
const arg = (name, dflt) => {
  const i = argv.indexOf(`--${name}`);
  return i >= 0 && argv[i + 1] ? Number(argv[i + 1]) : dflt;
};
const FPS = arg('fps', 30);
const FROM = arg('from', 0);
const TO = arg('to', 0);              // 0 = 問頁面自己有多長
const W = 1920;
const H = 1080;

function run(cmd, args, opts = {}) {
  return new Promise((resolve, reject) => {
    const ps = spawn(cmd, args, { stdio: ['ignore', 'pipe', 'pipe'], ...opts });
    let err = '';
    ps.stderr.on('data', d => { err += d; });
    ps.on('error', reject);
    ps.on('close', code => code === 0 ? resolve() : reject(new Error(`${cmd} 離開代碼 ${code}\n${err.slice(-1500)}`)));
  });
}

if (!CHROME) {
  console.error('找不到 Chrome 或 Edge。請安裝其中一個，或在 CHROME 陣列裡補上路徑。');
  process.exit(1);
}
if (!fs.existsSync(SCENE)) {
  console.error(`找不到分鏡檔：${path.relative(ROOT, SCENE)}`);
  process.exit(1);
}

const require = createRequire(import.meta.url);
let puppeteer;
try { puppeteer = require('puppeteer-core'); }
catch { console.error('缺少 puppeteer-core。執行：npm i -D puppeteer-core'); process.exit(1); }

console.log(`Chrome：${CHROME}`);
fs.rmSync(FRAMES, { recursive: true, force: true });
fs.mkdirSync(FRAMES, { recursive: true });

const browser = await puppeteer.launch({
  executablePath: CHROME,
  headless: 'shell',
  args: [
    `--window-size=${W},${H}`,
    '--hide-scrollbars',
    '--force-device-scale-factor=1',
    // 關掉字型平滑的機器差異，讓不同機器跑出來的幀一致。
    '--font-render-hinting=none',
    '--disable-lcd-text',
    '--allow-file-access-from-files',
  ],
});

try {
  const page = await browser.newPage();
  await page.setViewport({ width: W, height: H, deviceScaleFactor: 1 });

  // 頁面裡的 console.error 要浮上來——不然分鏡寫錯只會得到一片空白的影片。
  page.on('pageerror', e => console.error('  [頁面錯誤]', e.message));
  page.on('console', m => { if (m.type() === 'error') console.error('  [console]', m.text()); });

  await page.goto(pathToFileURL(SCENE).href, { waitUntil: 'networkidle0' });

  // 分鏡自己說它多長。寫死在這裡的話，改分鏡就要改兩個地方，而其中一個一定會忘。
  const duration = await page.evaluate(() => window.SCENE?.duration ?? 0);
  if (!duration) {
    throw new Error('分鏡沒有宣告 window.SCENE.duration —— 不知道要算到第幾秒。');
  }
  const end = TO || duration;
  const total = Math.round((end - FROM) * FPS);
  console.log(`分鏡長度 ${duration}s，輸出 ${FROM}s–${end}s @ ${FPS}fps（${total} 幀）`);

  const t0 = Date.now();
  for (let i = 0; i < total; i++) {
    const t = FROM + i / FPS;
    // 頁面沒有自己的時鐘：我們給它時間，它畫出那一刻。
    await page.evaluate((sec) => window.SCENE.seek(sec), t);
    await page.screenshot({
      path: path.join(FRAMES, `f${String(i).padStart(5, '0')}.png`),
      optimizeForSpeed: true,
    });
    if (i % 30 === 0 || i === total - 1) {
      const pct = Math.round(((i + 1) / total) * 100);
      const rate = (i + 1) / ((Date.now() - t0) / 1000);
      process.stdout.write(`\r  ${String(pct).padStart(3)}%  ${i + 1}/${total} 幀  ${rate.toFixed(1)} 幀/秒   `);
    }
  }
  process.stdout.write('\n');
} finally {
  await browser.close();
}

console.log('合成 MP4…');
await run('ffmpeg', [
  '-y', '-framerate', String(FPS),
  '-i', path.join(FRAMES, 'f%05d.png'),
  // yuv420p + even dimensions：少了這個，很多播放器（含 PowerPoint 內嵌）會放不出來。
  '-c:v', 'libx264', '-pix_fmt', 'yuv420p',
  '-crf', '18', '-preset', 'slow',
  '-movflags', '+faststart',
  OUT,
]);

fs.rmSync(FRAMES, { recursive: true, force: true });
const kb = fs.statSync(OUT).size / 1024;
console.log(`完成：${path.relative(ROOT, OUT)}  ${(kb / 1024).toFixed(1)} MB`);
