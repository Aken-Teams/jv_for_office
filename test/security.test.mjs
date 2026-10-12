/**
 * 資安加固，照暫緩清單一條一條來。
 *
 * ── 這幾項為什麼是這幾項 ──
 * `excel_run_script` / `word_run_script` 是用 `new Function` 建的，函式本體跑在
 * 全域範圍，拿得到 `window`、`fetch`、`localStorage`——而工作階段的 token 就存在
 * localStorage 裡。Proxy guard 包的是 Office 物件，包不住整個 realm。
 *
 * 而且 JS 這一層補不起來：`[].constructor.constructor('return fetch')()` 一行就
 * 繞過任何 `with` 做的假作用域。唯一腳本吵不贏的邊界是瀏覽器在網路層執行的那條。
 *
 * 所以這支測的是**那條邊界在不在**，不是「有沒有藏好全域變數」——後者本來就擋不住。
 */
import fs from 'node:fs';

let pass = 0;
let fail = 0;
function ok(name, cond, detail) {
  if (cond) { pass++; console.log(`PASS  ${name}`); } else {
    fail++; console.log(`FAIL  ${name}${detail === undefined ? '' : `  → ${detail}`}`);
  }
}

const read = (f) => fs.readFileSync(new URL(f, import.meta.url), 'utf8');

// ── CSP ─────────────────────────────────────────────────────────────────────
{
  const html = read('../src/taskpane/taskpane.html');
  const m = /<meta http-equiv="Content-Security-Policy" content="([\s\S]*?)"\s*\/>/.exec(html);
  ok('有 CSP', !!m);
  const policy = m ? m[1].replace(/\s+/g, ' ').trim() : '';
  const directive = (name) => (new RegExp(`${name} ([^;]*)`).exec(policy) || [])[1] || '';

  ok('default-src 是 none', directive('default-src').trim() === "'none'", directive('default-src'));
  ok('object-src 是 none', directive('object-src').trim() === "'none'", directive('object-src'));
  ok('base-uri 是 none', directive('base-uri').trim() === "'none'", directive('base-uri'));

  /**
   * connect-src 是這份政策的全部意義。
   *
   * 腳本讀得到 token 這件事擋不住；能擋的是它把 token 送出去。所以這一行不可以有
   * 萬用主機——`*` 或 `https:` 都等於沒設。
   */
  const connect = directive('connect-src');
  ok('connect-src 有列出來', !!connect.trim(), connect);
  ok('connect-src 沒有開放任意主機',
    !/(^|\s)\*(\s|$)/.test(connect) && !/(^|\s)https:(\s|$)/.test(connect), connect);
  ok('connect-src 帶著 API 的佔位符', /%API_ORIGIN%/.test(connect), connect);

  // Office.js 從微軟的 CDN 載入，擋掉的話整個增益集起不來。
  ok('script-src 放行 Office.js 的 CDN',
    /appsforoffice\.microsoft\.com/.test(directive('script-src')), directive('script-src'));

  /**
   * 佔位符一定要有人替換。
   *
   * meta CSP 是靜態的、`apiBase` 是執行時從 config.json 讀的，兩者不可能都成立。
   * 開發時由 dev-server 送出頁面時替換；部署時負責送這個檔案的那一端要做同一件事。
   * 沒人替換的話，窗格連不到自己的伺服器——而且會在 console 吵，那是對的失敗方式。
   */
  const dev = read('../tools/dev-server.mjs');
  ok('dev-server 會替換佔位符', /%API_ORIGIN%/.test(dev) && /apiOrigin\(\)/.test(dev));
  ok('替換的來源是 config.json', /config\.json/.test(dev));

  // 真的替換得出一個合法的 origin。
  const cfg = JSON.parse(read('../config.json'));
  const origin = new URL(cfg.apiBase).origin;
  const filled = connect.replaceAll('%API_ORIGIN%', origin);
  ok('替換後是一個具體的來源', filled.includes(origin) && !filled.includes('%'), filled);
}

// ── dev-server 的路徑檢查 ───────────────────────────────────────────────────
{
  const dev = read('../tools/dev-server.mjs');

  /**
   * 同名前綴的兄弟目錄。
   *
   * ROOT 是 `…/jv_for_office` 的話，`…/jv_for_office_private/keys.txt` 也通過
   * `startsWith(ROOT)`——那是一個完全不同的目錄。差一個分隔符號而已。
   */
  ok('路徑檢查帶了分隔符號', /rootWithSep|ROOT \+ path\.sep|ROOT \+ sep/.test(dev),
    (/startsWith\([^)]*\)/.exec(dev) || [])[0]);
  // 註解裡會提到舊的寫法（那段註解就是在解釋這個 bug），所以先把註解剝掉再驗。
  const code = dev.replace(/\/\*[\s\S]*?\*\//g, '').replace(/^\s*\/\/.*$/gm, '');
  ok('程式碼裡不再只用 startsWith(ROOT)', !/startsWith\(ROOT\)/.test(code),
    (/.*startsWith\(ROOT\).*/.exec(code) || [])[0]);

  // 點檔：這個 repo 的 .env 就放著 API 憑證。
  ok('擋掉點檔', /startsWith\('\.'\)/.test(dev), '沒看到點檔規則');

  // 真的跑一次判斷邏輯，不是只檢查字串有沒有出現。
  const ROOT = 'd:\\github\\jv_for_office';
  const sep = '\\';
  const rootWithSep = ROOT + sep;
  const sibling = 'd:\\github\\jv_for_office_private\\keys.txt';
  ok('兄弟目錄會被擋', !(sibling === ROOT || sibling.startsWith(rootWithSep)), sibling);
  ok('自己底下的檔案照樣過',
    `${ROOT}${sep}src${sep}taskpane${sep}app.js`.startsWith(rootWithSep));
  ok('/.env 會被點檔規則擋',
    '/.env'.split(/[\\/]/).some((s) => s.startsWith('.') && s !== '.' && s !== '..'));
}

// ── 給模型的錯誤訊息要講真話 ────────────────────────────────────────────────
{
  /**
   * 「沒有 require / import / fetch / document」。
   *
   * 前兩個是真的，後兩個不是——腳本跑在全域範圍，fetch 和 document 都在。對模型講
   * 不準確的話比不講更糟：它會據此推論這個環境是封閉的，然後寫出依賴那個前提的
   * 程式，而那個前提剛好就是這整份清單要處理的問題。
   */
  for (const f of ['../src/taskpane/hosts/excel.js', '../src/taskpane/hosts/word.js']) {
    const src = read(f);
    const line = /也沒有 [^']*/.exec(src);
    ok(`${f.split('/').pop()} 不再宣稱沒有 fetch`,
      !/也沒有 require \/ import \/ fetch/.test(src), line && line[0]);
    ok(`${f.split('/').pop()} 有說明 fetch 其實存在`,
      /fetch 和 document 這個環境是有的|fetch.*是有的/.test(src));
  }
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
