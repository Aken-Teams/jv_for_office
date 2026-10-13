/**
 * .env → config.json
 *
 * The task pane is static files with no build step, so there is no point at which
 * an environment variable could be baked into the JavaScript. Instead the pane
 * fetches config.json at startup.
 *
 * Why not just let the user type the server address (which is what the login
 * screen used to do): nobody outside this repo knows or should have to know where
 * the API lives, and a typo produces a login failure that looks like a wrong
 * password. It is deployment configuration, not a user preference.
 *
 * config.json is a plain file on the same host as the pane, so a deployment can
 * also just edit it in place — no regeneration, no redeploy of anything else.
 *
 * Run: npm run config   (npm run dev does it automatically)
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const ENV_FILE = path.join(ROOT, '.env');
const OUT = path.join(ROOT, 'config.json');

const DEFAULTS = { API_BASE: 'http://localhost:12054', DEPLOY_CHANNEL: 'prod' };

/** Minimal .env reader — KEY=value, # comments, no interpolation. */
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

// Real environment variables win, so CI or a deploy script can override without
// writing a file.
const env = { ...DEFAULTS, ...readEnv(ENV_FILE), ...process.env };

/**
 * 哪一套環境。
 *
 * ── 為什麼需要 ──
 * 「我擔心有些人增益集一打開就有測試區、正式區版本，他會搞混」。兩套同時裝著的
 * 時候，側邊欄長得一模一樣——同一個標題、同一個圖示、同一個配色。使用者分不出
 * 自己正在改的那份文件是跑在哪一套上。
 *
 * 所以頻道寫進 config.json，窗格開起來就把它掛在最上面。走 .env 是因為這是部署
 * 設定不是程式差異：同一份程式碼部署兩次，差別只在一個環境變數。
 *
 * `prod` 以外的值都會顯示橫幅——預設是 prod，所以忘記設定的那一邊是**正式區沒有
 * 橫幅**，而不是測試區沒有警告。寧可正式區多一條橫幅被罵，也不要測試區安靜。
 */
const channel = String(env.DEPLOY_CHANNEL || DEFAULTS.DEPLOY_CHANNEL).trim().toLowerCase();
const config = {
  apiBase: String(env.API_BASE || DEFAULTS.API_BASE).replace(/\/+$/, ''),
  channel,
  /** 橫幅上的字。沒設就用頻道名稱，讓部署的人可以寫「測試區（10/15 前請卸載）」。 */
  channelLabel: String(env.DEPLOY_LABEL || '').trim(),
};

fs.writeFileSync(OUT, JSON.stringify(config, null, 2) + '\n', 'utf-8');

const source = process.env.API_BASE ? '環境變數'
  : fs.existsSync(ENV_FILE) ? '.env'
  : '預設值（沒有 .env，可從 .env.example 複製一份）';
console.log(`config.json 已產生 — apiBase = ${config.apiBase}（來源：${source}）`);
console.log(`  頻道 = ${config.channel}`
  + (config.channel === 'prod' ? '（正式區，不會顯示橫幅）' : '（窗格會顯示橫幅）'));
