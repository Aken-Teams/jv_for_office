/**
 * 錯誤訊息：使用者看得懂，而且知道下一步。
 *
 * ── 這裡守的是什麼 ──
 * 主要失敗路徑上原本寫的是「確認伺服器有在跑，而且 CORS_ORIGINS 允許 …」——那是
 * 寫給工程師的指示，顯示給正在打報價單的人看。另一條是「伺服器錯誤 (500)」，它
 * 告訴對方「有東西壞了」，而那件事他早就發現了，然後沒有給他任何能做的事。
 *
 * ── 但細節不能丟掉 ──
 * 真的壞掉的時候 IT 需要狀態碼和位址，而唯一能把它們轉達出去的人，就是坐在螢幕
 * 前面的那個。所以是「人話在前、可轉達的字串在後」，不是把技術細節刪掉——刪掉
 * 會讓介面友善而讓客服電話無解。
 */
import fs from 'fs';
const app = fs.readFileSync('src/taskpane/app.js', 'utf8');
const i18n = fs.readFileSync('src/taskpane/i18n.js', 'utf8');
let pass = 0, fail = 0;
const ok = (n, c, d) => { c ? pass++ : fail++; console.log(`${c ? 'PASS' : 'FAIL'}  ${n}${c ? '' : '  → ' + (d || '')}`); };
ok('沒有裸的狀態碼當訊息', !/伺服器錯誤 \(\{0\}\)|操作失敗（\{0\}）/.test(app));
ok('沒有把 CORS_ORIGINS 講給使用者聽', !/CORS_ORIGINS 允許 \{1\}/.test(app));
ok('連不上伺服器不再只是丟例外訊息', !/連不上伺服器（\{0\}）/.test(app));
ok('技術細節有保留下來給 IT', /note-detail/.test(app) && /troubleDetail/.test(app));
ok('離線和連不到分開講', /navigator\.onLine/.test(app));
for (const code of ['408', '429', '413', '404']) {
  ok(`${code} 有自己的說法`, new RegExp(`n === ${code}|n >= 500`).test(app));
}
ok('五百系列講「不是你操作的問題」', /不是你操作的問題/.test(app));
ok('伺服器自己的訊息優先', /err\.error \|\| httpTrouble/.test(app));
ok('登入失敗也不丟狀態碼', !/登入失敗 \(\$\{res\.status\}\)/.test(app));
const keys = [...app.matchAll(/t\('([^']{12,})'\)/g)].map(m => m[1]);
const missing = keys.filter(k => /伺服器|網路|要求|內容太大/.test(k) && !i18n.includes(`'${k}'`));
ok('新訊息都有翻譯', missing.length === 0, missing.join(' | '));
/**
 * 伺服器的 CLI 比使用者挑的模型還舊。
 *
 * 原文是「Claude Code 2.1.231 does not support this model; version 2.1.280 or
 * newer is required. Run 'claude update'…」——正確，而且是寫給管理那台伺服器的人
 * 看的。但看到它的人是從選單挑了一個模型，他手上沒有那台機器的 shell，而且那句
 * 建議對他是**錯的**：在自己筆電上跑 claude update 修不了任何東西。
 *
 * 所以要講真正的原因、給他現在就能做的事（換模型）、再講 IT 要做什麼。
 */
{
  const cli = fs.readFileSync(
    new URL('../../ai-agents-office/server/src/services/claudeCli.ts', import.meta.url), 'utf8');
  const m = cli.match(/const skew = t\.match\((\/.*?\/i)\)/);
  ok('有認出版本落差這種錯誤', !!m);
  if (m) {
    const real = "API Error: 400 Claude Code 2.1.231 does not support this model; "
      + "version 2.1.280 or newer is required. Run 'claude update', or update the "
      + "Claude desktop app, then try again.";
    const hit = real.toLowerCase().match(eval(m[1]));
    ok('抓得到真實的那句錯誤', !!hit, real.slice(0, 60));
    ok('抓得出兩個版本號', hit && hit[1] === '2.1.231' && hit[2] === '2.1.280',
      hit && hit.slice(1).join(' / '));
  }
  const msg = (cli.match(/伺服器上的 AI 程式版本太舊[\s\S]{0,400}?;/) || [''])[0];
  ok('講了是伺服器端的問題，不是使用者的', /伺服器上/.test(msg));
  // 看到這句的人沒有那台機器的 shell——要給他現在就能做的事。
  ok('給了使用者現在就能做的事：換模型', /換一個模型/.test(msg), msg.slice(0, 120));
  ok('也講了 IT 要做什麼', /通知 IT/.test(msg), msg.slice(0, 120));
  ok('版本號會被引回去，方便轉達', /skew\[1\]/.test(msg) && /skew\[2\]/.test(msg));
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
