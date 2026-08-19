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

const DEFAULTS = { API_BASE: 'http://localhost:12054' };

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

const config = {
  apiBase: String(env.API_BASE || DEFAULTS.API_BASE).replace(/\/+$/, ''),
};

fs.writeFileSync(OUT, JSON.stringify(config, null, 2) + '\n', 'utf-8');

const source = process.env.API_BASE ? '環境變數'
  : fs.existsSync(ENV_FILE) ? '.env'
  : '預設值（沒有 .env，可從 .env.example 複製一份）';
console.log(`config.json 已產生 — apiBase = ${config.apiBase}（來源：${source}）`);
