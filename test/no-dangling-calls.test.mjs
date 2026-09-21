/**
 * Every function the pane calls has to exist.
 *
 * ── Why this is not what a linter already does for us ──
 * There is no build step. The pane is raw ES modules that WebView2 loads
 * directly, so nothing between the editor and Excel ever resolves a name.
 * `node --check` parses; it does not link. A call to a function that was deleted
 * is perfectly valid JavaScript right up until the line runs — and the line that
 * runs is usually the one somebody is demonstrating.
 *
 * ── The afternoon that produced this file ──
 * Merging two menus into one meant deleting three functions and rewriting the
 * callers. The scripts that did it worked by finding a function and walking back
 * to its doc comment, which is fine until a neighbour has already been removed
 * and the walk lands on the wrong comment. It took `plainText`, `loadSkills` and
 * `templateRequest` with it — silently, three separate times. Two were caught by
 * unrelated tests that happened to name them; the third was found by hand.
 *
 * So the check is crude on purpose: collect what the file defines, collect what
 * it calls, and complain about the difference. It cannot see scope and does not
 * try — see IGNORED for how the noise is handled.
 */
import fs from 'fs';

const FILES = [
  '../src/taskpane/app.js',
  '../src/taskpane/recorder.js',
  '../src/taskpane/hosts/compare.js',
];

/**
 * Names this check is not smart enough to resolve, and does not need to be.
 *
 * Language keywords that are followed by a paren, browser and Office globals,
 * and — the only interesting category — identifiers that are only ever local:
 * parameters, closures and destructured bindings. Teaching it scope would make
 * it a compiler; listing the handful it trips on keeps it a test.
 */
const IGNORED = new Set([
  // keywords and operators that look like calls
  'if', 'for', 'while', 'switch', 'catch', 'return', 'typeof', 'function',
  'async', 'await', 'void', 'of', 'in', 'var', 'new', 'do', 'else', 'yield',
  // platform
  'fetch', 'setTimeout', 'clearTimeout', 'setInterval', 'clearInterval',
  'requestAnimationFrame', 'Promise', 'Object', 'Array', 'JSON', 'String',
  'Number', 'Boolean', 'Math', 'Date', 'Set', 'Map', 'WeakMap', 'WeakSet',
  'Error', 'RegExp', 'Symbol', 'Proxy', 'Reflect', 'BigInt',
  'console', 'document', 'window', 'localStorage', 'sessionStorage', 'crypto',
  'self', 'navigator', 'location', 'history', 'FormData', 'FileReader', 'Image',
  'Blob', 'File', 'AbortController', 'URL', 'URLSearchParams', 'TextDecoder',
  'TextEncoder', 'Intl', 'isNaN', 'isFinite', 'parseInt', 'parseFloat',
  'encodeURIComponent', 'decodeURIComponent', 'structuredClone',
  'queueMicrotask', 'atob', 'btoa', 'alert', 'confirm', 'prompt',
  'Office', 'Excel', 'Word', 'PowerPoint', 'OfficeExtension',
  // CSS and escape sequences inside template strings
  'calc', 'translateX', 'translateY', 'rgba', 'rgb', 'var', 'u0000',
]);

let pass = 0, fail = 0;
const ok = (name, cond, detail) => {
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}`);
  if (cond) pass++;
  else { if (detail) console.log(detail); fail++; }
};

for (const rel of FILES) {
  const url = new URL(rel, import.meta.url);
  if (!fs.existsSync(url)) { ok(`${rel} 存在`, false); continue; }

  // Comments go first: prose mentions function names constantly, and this file
  // learned that lesson from a sibling test that failed on its own doc block.
  const src = fs.readFileSync(url, 'utf8')
    .replace(/\/\*[\s\S]*?\*\//g, '')
    .replace(/^\s*\/\/.*$/gm, '');

  const defined = new Set();
  for (const m of src.matchAll(/(?:async\s+)?function\s+([A-Za-z_$][\w$]*)/g)) defined.add(m[1]);
  for (const m of src.matchAll(/^\s*(?:export\s+)?(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*=/gm)) {
    defined.add(m[1]);
  }
  for (const m of src.matchAll(/^import\s*\{([^}]*)\}/gm)) {
    for (const n of m[1].split(',')) defined.add(n.trim().split(/\s+as\s+/).pop());
  }
  for (const m of src.matchAll(/^import\s+([A-Za-z_$][\w$]*)/gm)) defined.add(m[1]);
  // Locals: parameters, destructured bindings, and inner consts. Collected so a
  // closure called by name does not read as a missing top-level function.
  for (const m of src.matchAll(/(?:const|let)\s*\{([^}]*)\}\s*=/g)) {
    for (const n of m[1].split(',')) defined.add(n.trim().split(/[:=]/)[0].trim());
  }
  for (const m of src.matchAll(/\(([^)]*)\)\s*=>/g)) {
    for (const n of m[1].split(',')) {
      // Leading punctuation happens because the param scan starts at the wrong
      // paren in `new Promise((resolve) => …)`; trimming it is cheaper than
      // teaching the regex to balance brackets.
      const bare = n.trim().replace(/^[^A-Za-z_$]+/, '').split(/[:=]/)[0].trim();
      if (/^[A-Za-z_$][\w$]*$/.test(bare)) defined.add(bare);
    }
  }
  for (const m of src.matchAll(/function\s*[A-Za-z_$\w]*\s*\(([^)]*)\)/g)) {
    for (const n of m[1].split(',')) {
      // Leading punctuation happens because the param scan starts at the wrong
      // paren in `new Promise((resolve) => …)`; trimming it is cheaper than
      // teaching the regex to balance brackets.
      const bare = n.trim().replace(/^[^A-Za-z_$]+/, '').split(/[:=]/)[0].trim();
      if (/^[A-Za-z_$][\w$]*$/.test(bare)) defined.add(bare);
    }
  }

  const missing = new Set();
  // `(?<![.\w$])` so method calls are skipped: `a.foo()` is a property, and
  // whether it exists is a question about `a`, not about this file.
  for (const m of src.matchAll(/(?<![.\w$])([A-Za-z_$][\w$]*)\s*\(/g)) {
    if (!defined.has(m[1]) && !IGNORED.has(m[1])) missing.add(m[1]);
  }

  const list = [...missing].sort();
  ok(`${rel.replace('../src/taskpane/', '')} 沒有呼叫不存在的函式`,
    list.length === 0,
    list.length ? `      找不到定義：${list.join('、')}` : '');

  /**
   * And the other direction: declared, never reached.
   *
   * The dangling-call check above is blind to this. `expandTemplate` survived a
   * menu merge perfectly intact while the click handler that called it was
   * deleted — so the chip could be neither opened nor removed, and nothing in
   * the file was broken enough to notice. A function nobody calls is either dead
   * or unwired, and unwired is a bug wearing dead code's clothes.
   *
   * Only top-level `function` declarations, and only in files with no exports to
   * speak of: an exported helper is called from somewhere this check cannot see.
   */
  const declared = [...src.matchAll(/^(?:async\s+)?function\s+([A-Za-z_$][\w$]*)/gm)].map(m => m[1]);
  const exported = new Set(
    [...src.matchAll(/^export\s+(?:async\s+)?function\s+([A-Za-z_$][\w$]*)/gm)].map(m => m[1]));
  const orphans = declared.filter((name) => {
    if (exported.has(name)) return false;
    // Its own declaration is one mention; anything else means somebody uses it.
    return (src.match(new RegExp(`\\b${name}\\b`, 'g')) || []).length <= 1;
  });
  ok(`${rel.replace('../src/taskpane/', '')} 沒有宣告了卻沒人呼叫的函式`,
    orphans.length === 0,
    orphans.length ? `      沒人呼叫：${orphans.join('、')}` : '');
}

console.log(`\n${pass} passed, ${fail} failed`);
if (fail) process.exit(1);
