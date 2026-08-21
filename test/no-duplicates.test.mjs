/**
 * Catch the duplicates that nothing else does.
 *
 * ── Why this test exists ──
 * Two people editing the same files produced, in one afternoon: two identical
 * 36-line blocks in taskpane.css, two `.attach-file` rules with different
 * bodies, and — on the server — two `router.post('/attachments')` registrations
 * where Express silently uses the first and drops the second.
 *
 * Not one of those is a syntax error, a type error, or a failing test. The
 * compiler is happy with a duplicated object key, the browser is happy with a
 * duplicated selector, and Express is happy with a duplicated route. They are
 * found by reading, which means they are found late or not at all.
 *
 * So: the file is parsed for the shapes where "twice" is always wrong.
 *
 * ── What is deliberately NOT flagged ──
 * Repetition that is meant. Three `$('log').addEventListener('click')` handlers
 * matching `.chip`, `.cite` and `a.link` are three concerns, not one mistake,
 * and a check that cannot tell those apart would be turned off within a week.
 */
import fs from 'fs';

const read = (p) => fs.readFileSync(new URL(p, import.meta.url), 'utf8');
const problems = [];
const check = (name, fn) => {
  const before = problems.length;
  fn();
  console.log(`${problems.length === before ? 'PASS' : 'FAIL'}  ${name}`);
};

// ── Element ids ───────────────────────────────────────────────────────────────
// A repeated id makes document.getElementById return whichever came first, so
// half the wiring silently attaches to the wrong element.
check('taskpane.html 沒有重複的 id', () => {
  const html = read('../src/taskpane/taskpane.html');
  const seen = new Map();
  for (const m of html.matchAll(/\sid="([^"]+)"/g)) seen.set(m[1], (seen.get(m[1]) || 0) + 1);
  for (const [id, n] of seen) if (n > 1) problems.push(`  id="${id}" 出現 ${n} 次`);
});

// ── CSS selectors ─────────────────────────────────────────────────────────────
// The later rule wins for the properties it sets and the earlier one still
// applies for the rest, so a duplicate is not "dead code" — it is two half-live
// definitions of the same thing.
check('taskpane.css 沒有重複的頂層選擇器', () => {
  const css = read('../src/taskpane/taskpane.css').replace(/\/\*[\s\S]*?\*\//g, '');
  const seen = new Map();
  let depth = 0;
  let buf = '';
  for (const ch of css) {
    if (ch === '{') {
      if (depth === 0) {
        const sel = buf.trim().replace(/\s+/g, ' ');
        if (sel && !sel.startsWith('@')) seen.set(sel, (seen.get(sel) || 0) + 1);
      }
      depth++;
      buf = '';
    } else if (ch === '}') {
      depth = Math.max(0, depth - 1);
      buf = '';
    } else if (depth === 0) {
      buf += ch;
    }
  }
  for (const [sel, n] of seen) if (n > 1) problems.push(`  選擇器 "${sel}" 出現 ${n} 次`);
});

// ── i18n keys ─────────────────────────────────────────────────────────────────
// A duplicated key in an object literal is not an error in JavaScript — the last
// one wins, and the translation someone carefully wrote is simply gone.
check('i18n 字典沒有重複的 key', () => {
  const src = read('../src/taskpane/i18n.js');
  for (const dict of ['zhCN', 'en']) {
    const start = src.indexOf(`const ${dict} = {`);
    if (start < 0) { problems.push(`  找不到 ${dict} 字典`); continue; }
    const body = src.slice(start, src.indexOf('\n};', start));
    const seen = new Map();
    for (const m of body.matchAll(/^\s+'((?:[^'\\]|\\.)*)':/gm)) seen.set(m[1], (seen.get(m[1]) || 0) + 1);
    for (const [k, n] of seen) if (n > 1) problems.push(`  ${dict} 裡 '${k.slice(0, 40)}' 有 ${n} 個`);
  }
});

// ── Both dictionaries carry the same keys ─────────────────────────────────────
// A key in zh-CN but not in en is a sentence that falls back to Chinese for an
// English user — which reads as a bug in the product, not as a missing string.
check('簡中與英文字典的 key 一致', () => {
  const src = read('../src/taskpane/i18n.js');
  const keysOf = (dict) => {
    const start = src.indexOf(`const ${dict} = {`);
    const body = src.slice(start, src.indexOf('\n};', start));
    return new Set([...body.matchAll(/^\s+'((?:[^'\\]|\\.)*)':/gm)].map(m => m[1]));
  };
  const zh = keysOf('zhCN');
  const en = keysOf('en');
  for (const k of zh) if (!en.has(k)) problems.push(`  英文字典少了 '${k.slice(0, 40)}'`);
  for (const k of en) if (!zh.has(k)) problems.push(`  簡中字典少了 '${k.slice(0, 40)}'`);
});

// ── Top-level declarations ────────────────────────────────────────────────────
// Two `function foo` in one module means the second quietly replaces the first,
// which is exactly what a careless paste produces.
check('模組沒有重複的頂層宣告', () => {
  for (const rel of [
    '../src/taskpane/app.js', '../src/taskpane/i18n.js',
    '../src/taskpane/hosts/word.js', '../src/taskpane/hosts/excel.js',
    '../src/taskpane/hosts/omml.js', '../src/taskpane/hosts/guard.js',
    '../src/taskpane/hosts/index.js',
  ]) {
    const src = read(rel);
    const seen = new Map();
    for (const m of src.matchAll(/^(?:export\s+)?(?:async\s+)?(?:function|const|let|class)\s+([A-Za-z_$][\w$]*)/gm)) {
      seen.set(m[1], (seen.get(m[1]) || 0) + 1);
    }
    const file = rel.split('/').pop();
    for (const [name, n] of seen) if (n > 1) problems.push(`  ${file}: ${name} 宣告了 ${n} 次`);
  }
});

// ── The tools the pane offers must exist ──────────────────────────────────────
// The reverse direction is NOT checked: the server legitimately carries tools
// the pane never lists — ones gated on a requirement set, and ones the server
// answers itself (see WORD_SERVER_TOOLS). A pane offering a tool the server has
// never heard of, though, is an agent that will be told 未知的工具 and retry.
check('窗格宣告的工具在伺服器都存在', () => {
  const server = new URL('../../ai-agents-office/server/src/services/', import.meta.url);
  if (!fs.existsSync(server)) {
    console.log('      （找不到 ai-agents-office，跳過跨 repo 檢查）');
    return;
  }
  for (const [host, spec] of [['excel', 'excelToolSpec.ts'], ['word', 'wordToolSpec.ts']]) {
    const specPath = new URL(spec, server);
    if (!fs.existsSync(specPath)) continue;
    const declared = new Set(
      [...fs.readFileSync(specPath, 'utf8').matchAll(/^\s+name: '((?:excel|word)_[a-z_]+)',/gm)].map(m => m[1]),
    );
    const paneSrc = read(`../src/taskpane/hosts/${host}.js`);
    // Both the base list and everything pushed behind a requirement-set gate.
    const offered = new Set([
      ...[...paneSrc.matchAll(/^\s+'((?:excel|word)_[a-z_]+)',/gm)].map(m => m[1]),
      ...[...paneSrc.matchAll(/tools\.push\(([^)]*)\)/g)]
        .flatMap(m => [...m[1].matchAll(/'((?:excel|word)_[a-z_]+)'/g)].map(x => x[1])),
    ]);
    for (const t of offered) if (!declared.has(t)) problems.push(`  ${host}: 窗格宣告 ${t}，伺服器沒有`);
  }
});

if (problems.length) {
  console.log('\n' + problems.join('\n'));
  console.log(`\n${problems.length} 項重複或不一致`);
  process.exit(1);
}
console.log('\n沒有重複。');
