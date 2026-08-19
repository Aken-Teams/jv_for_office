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

const server = https.createServer(await devCerts.getHttpsServerOptions(), (req, res) => {
  const urlPath = decodeURIComponent(new URL(req.url, `https://localhost:${PORT}`).pathname);
  const rel = urlPath === '/' ? '/src/taskpane/taskpane.html' : urlPath;
  const file = path.join(ROOT, rel);

  // Path traversal guard: a dev server still shouldn't hand out the whole disk.
  if (!file.startsWith(ROOT)) { res.writeHead(403).end('Forbidden'); return; }

  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404).end('Not found'); return; }
    res.writeHead(200, {
      'Content-Type': MIME[path.extname(file).toLowerCase()] || 'application/octet-stream',
      // The pane is edited constantly during development and Excel's WebView
      // caches aggressively — without this you debug against stale code.
      'Cache-Control': 'no-store',
    });
    res.end(data);
  });
});

server.listen(PORT, () => {
  console.log(`JVision for Excel — 任務窗格開發伺服器`);
  console.log(`  https://localhost:${PORT}/src/taskpane/taskpane.html`);
  console.log(`  manifest: ${path.join(ROOT, 'manifest.xml')}`);
});
