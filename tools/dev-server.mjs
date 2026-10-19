/**
 * Dev server for the task pane.
 *
 * Office add-ins must be served over HTTPS — including on localhost, including
 * during development. There is no http:// escape hatch: Excel simply refuses to
 * load the pane. `office-addin-dev-certs` mints and installs a locally-trusted
 * certificate the first time it runs (Windows will show a one-time prompt).
 *
 * Static only, no bundler. The pane is plain ES modules, which WebView2 (the
 * engine behind the desktop task pane) handles natively — so there is no build
 * step between editing a file and reloading the pane.
 */
import https from 'https';
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import devCerts from 'office-addin-dev-certs';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const PORT = parseInt(process.env.PORT || '3443', 10);

// Installing the dev certificate adds a root CA to the machine's trust store —
// a system-level change, and one that pops a Windows prompt. So this server
// checks and instructs rather than doing it behind your back. Run the install
// yourself, once, and every later start is silent.
if (!devCerts.isCaCertificateInstalled() || !devCerts.verifyCertificates()) {
  console.error('尚未安裝本機開發憑證。Office 增益集一定要走 HTTPS，先執行一次：\n');
  console.error('  npx office-addin-dev-certs install\n');
  console.error('（會跳出 Windows 憑證安裝確認，按同意。之後就不會再問。）');
  process.exit(1);
}

/**
 * Request log — to console AND to a file.
 *
 * When Excel draws a generic ribbon icon there is no way to tell "Office never
 * asked for the image" from "Office asked and rejected it" except by watching
 * the wire. The file copy exists because the console output lives in whichever
 * terminal happens to be running the server, which is not always the one where
 * the question gets asked. The User-Agent separates the task pane's WebView from
 * Office's own fetcher (manifest, ribbon icons).
 */
const LOG_FILE = path.join(ROOT, 'devserver.log');
const NL = '\n';
try {
  fs.writeFileSync(LOG_FILE, `--- 啟動 ${new Date().toISOString()} ---${NL}`);
} catch { /* non-fatal */ }

function log(line) {
  console.log(line);
  try { fs.appendFileSync(LOG_FILE, line + NL); } catch { /* non-fatal */ }
}

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.xml': 'text/xml; charset=utf-8',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.ico': 'image/x-icon',
};

/**
 * The origin the pane is allowed to talk to, from config.json.
 *
 * Read per request rather than once at boot: config.json is edited while the
 * dev server is running, and a CSP that still names yesterday's port fails in a
 * way that looks like the server is down.
 */
function apiOrigin() {
  try {
    const cfg = JSON.parse(fs.readFileSync(path.join(ROOT, 'config.json'), 'utf8'));
    const u = new URL(String(cfg.apiBase || ''));
    return u.origin;
  } catch {
    return 'http://localhost:12054';
  }
}

const server = https.createServer(await devCerts.getHttpsServerOptions(), (req, res) => {
  const urlPath = decodeURIComponent(new URL(req.url, `https://localhost:${PORT}`).pathname);
  const rel = urlPath === '/' ? '/src/taskpane/taskpane.html' : urlPath;
  const file = path.join(ROOT, rel);

  /**
   * Path guard.
   *
   * ── Two holes this closes ──
   * `file.startsWith(ROOT)` passes for a SIBLING directory whose name starts
   * with the same letters: with ROOT at `…/jv_for_office`, a request for
   * `../jv_for_office_private/keys.txt` resolves to a path that starts with
   * ROOT and is served. Comparing against `ROOT + sep` is the whole fix.
   *
   * And nothing stopped `/.env` — which in this repo is where the API
   * credentials live. A dev server on localhost is still a server, and the pane
   * runs inside Office's WebView where any page can reach it.
   */
  const rootWithSep = ROOT.endsWith(path.sep) ? ROOT : ROOT + path.sep;
  if (file !== ROOT && !file.startsWith(rootWithSep)) {
    res.writeHead(403).end('Forbidden');
    return;
  }
  // Dotfiles and the git directory are never web assets.
  if (rel.split(/[\\/]/).some((seg) => seg.startsWith('.') && seg !== '.' && seg !== '..')) {
    res.writeHead(403).end('Forbidden');
    return;
  }

  const ua = (req.headers['user-agent'] || '').slice(0, 70);
  const stamp = new Date().toISOString().slice(11, 23);

  fs.readFile(file, (err, data) => {
    if (err) {
      log(`${stamp}  404  ${rel}  [${ua}]`);
      res.writeHead(404).end('Not found');
      return;
    }
    /**
     * Fill in the CSP's API origin on the way out.
     *
     * The policy in taskpane.html carries `%API_ORIGIN%` because a meta CSP is
     * static while `apiBase` is read from config.json at runtime. Serving the
     * page is the moment both are known, so this is where they meet. A deploy
     * that serves these files some other way has to do the same substitution —
     * see the comment on the meta tag.
     */
    let body = data;
    if (path.extname(file).toLowerCase() === '.html') {
      body = Buffer.from(String(data).replaceAll('%API_ORIGIN%', apiOrigin()), 'utf8');
    }

    log(`${stamp}  200  ${rel}  ${body.length}B  [${ua}]`);
    /**
     * 增益集的圖示**不能**送 no-store。
     *
     * ── 這一行害過什麼 ──
     * 微軟的疑難排解文件寫得很明確：「Office Add-ins require the ability to cache
     * image resources, and if the server hosting an image resource adds
     * CACHE-CONTROL directives to the response header, Office automatically
     * substitutes a generic or default image.」
     *
     * 也就是說，Office 會把圖抓下來（伺服器紀錄裡是乾淨的 200、位元組數也對），
     * 發現自己不准快取，然後畫一顆通用的六邊形——沒有錯誤訊息，什麼都沒有。
     *
     * 這個症狀花了一整個下午、五個錯誤的推論才找到原因（先後怪過 sideload 機制、
     * PNG 的 pHYs 區塊、網址快取鍵、manifest 版本號、圖示尺寸）。而標頭在第一次
     * 比對的時候就印出來過，只是當時三個主機都一樣，看起來像是「正常」。
     *
     * ── 所以只有圖示例外 ──
     * HTML/JS/CSS 還是要 no-store：窗格在開發時一直在改，WebView 快取得很兇，
     * 沒有這個就會對著舊程式碼除錯。這正是當初加它的理由，那個理由依然成立。
     *
     * 圖示不一樣——它們幾乎不變，而且換了之後本來就要升 manifest 版本號才會生效。
     *
     * ⚠ 部署的靜態站台要做同一件事。線上那兩台現在也送 no-store，所以 jv 和測試區
     *   的圖示壞在同一個原因上。
     */
    const isIcon = /\.(png|jpe?g|gif|svg|ico)$/i.test(file);
    res.writeHead(200, {
      'Content-Type': MIME[path.extname(file).toLowerCase()] || 'application/octet-stream',
      'Cache-Control': isIcon ? 'public, max-age=3600' : 'no-store',
    });
    res.end(body);
  });
});

server.listen(PORT, () => {
  console.log(`JVision for Excel — 任務窗格開發伺服器`);
  console.log(`  https://localhost:${PORT}/src/taskpane/taskpane.html`);
  console.log(`  manifest: ${path.join(ROOT, 'manifest.xml')}`);
});
