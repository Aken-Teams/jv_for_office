/**
 * Pin the part of ppt_get_overview that AC-11 rests on.
 *
 * ── The bug this covers ──
 * `ppt_add_slide` has always accepted layout_id + master_id, and nothing
 * anywhere told the model where to get them. So every generated slide landed on
 * the default layout, and slides added to a deck built on the company master did
 * not match the ones already in it. The ids are only reachable from the overview,
 * so the overview is where they have to be said — and the two things worth
 * pinning are that they ARE said, and that the block stays silent when there is
 * nothing to say. A line printed on every deck is noise nobody reads, and the
 * model pays for it in context on every single conversation.
 *
 * The colour half is a heuristic standing in for an API that does not exist:
 * PowerPoint exposes no theme to read, so the palette is recovered by counting
 * the colours already on the slides. What is tested here is the judgement in
 * that count — same colour written two ways, colours that carry no signal, and
 * the cap — not the Office plumbing that feeds it.
 */
import fs from 'fs';

const src = fs.readFileSync(new URL('../src/taskpane/hosts/powerpoint.js', import.meta.url), 'utf8');
const from = src.indexOf('/* ── palette:start');
const to = src.indexOf('/* ── palette:end');
if (from < 0 || to < 0) {
  console.error('找不到 palette:start / palette:end 標記，hosts/powerpoint.js 的結構變了。');
  process.exit(1);
}
const tmp = new URL('./.palette-under-test.mjs', import.meta.url);
fs.writeFileSync(tmp, src.slice(from, to)
  + '\nexport { normaliseHex, rankColours, describeLayouts, describeColours, THEME_REPORT };\n');
const { normaliseHex, rankColours, describeLayouts, describeColours, THEME_REPORT } = await import(tmp);

let pass = 0, fail = 0;
const ok = (name, got, want = true) => {
  const good = JSON.stringify(got) === JSON.stringify(want);
  console.log(`${good ? 'PASS' : 'FAIL'}  ${name}`);
  if (good) pass++;
  else { console.log(`      得到 ${JSON.stringify(got)}，預期 ${JSON.stringify(want)}`); fail++; }
};

// ── normaliseHex ─────────────────────────────────────────────────────────────
ok('補上 # 並轉大寫', normaliseHex('1f3864'), '#1F3864');
ok('已經有 # 的也轉大寫', normaliseHex('#1f3864'), '#1F3864');
ok('前後空白不算數', normaliseHex('  #1F3864 '), '#1F3864');
ok('不是六位十六進位就當作沒有', normaliseHex('red'), '');
ok('null 不會炸', normaliseHex(null), '');
ok('三位縮寫不接受——展開會猜錯', normaliseHex('#fff'), '');

// ── rankColours ──────────────────────────────────────────────────────────────
ok('同一個顏色的兩種寫法算同一個',
  rankColours(['1F3864', '#1f3864', '#C0392B']), ['#1F3864', '#C0392B']);
ok('白與黑不計', rankColours(['#FFFFFF', '#000000', '#C0392B']), ['#C0392B']);
ok('近白不會被誤殺——#FEFEFE 是有人選的顏色',
  rankColours(['#FEFEFE']), ['#FEFEFE']);
ok('用得多的排前面',
  rankColours(['#AAAAAA', '#BBBBBB', '#BBBBBB']), ['#BBBBBB', '#AAAAAA']);
ok('平手時維持出現順序，結果才是穩定的',
  rankColours(['#AAAAAA', '#BBBBBB']), ['#AAAAAA', '#BBBBBB']);
{
  // 一份簡報用到 30 種顏色是常有的事，全部列出來就不是「主題色」了。
  const many = Array.from({ length: THEME_REPORT + 4 }, (_, i) => `#${String(i + 11).repeat(3)}`);
  ok('超過上限就截斷', rankColours(many).length, THEME_REPORT);
}
ok('全部都沒訊號就回空陣列', rankColours(['#FFFFFF', 'red', '']), []);

// ── describeLayouts ──────────────────────────────────────────────────────────
ok('沒有母片就完全不出聲', describeLayouts([]), '');
ok('讀不到（舊版 PowerPoint）也不出聲', describeLayouts(null), '');
ok('有母片但沒有版面配置，一樣不出聲',
  describeLayouts([{ id: 'm1', name: '母片', layouts: [] }]), '');

{
  const text = describeLayouts([
    { id: 'm1', name: 'PANJIT 母片', layouts: [{ id: 'l1', name: '標題投影片' }, { id: 'l2', name: '標題及內容' }] },
  ]);
  ok('母片 id 有列出來', text.includes('master_id=m1'));
  ok('每個版面配置的 id 都有', text.includes('layout_id=l1') && text.includes('layout_id=l2'));
  ok('名稱原樣列出，模型才挑得到', text.includes('標題及內容'));
  ok('明說要成對給', text.includes('成對'));
  // 只給 id 而不說什麼時候該用，等於推薦一條更差的路：模型會拿原生版面配置
  // 配預留位置，繞過所有版型，做出 PowerPoint 的預設長相。
  ok('說清楚什麼時候才用這些 id', text.includes('原本就有設計好的版面'));
  ok('其他情況要用 ppt_build_slide', text.includes('其他情況一律用 ppt_build_slide'));
  ok('說明原生版面配置給的是預設長相不是設計', text.includes('那不是設計，只是沒有設計'));
}

ok('沒有名字的版面配置不會變成空白列',
  describeLayouts([{ id: 'm1', name: '', layouts: [{ id: 'l1', name: '' }] }]).includes('(未命名)'));

// ── describeColours ──────────────────────────────────────────────────────────
ok('兩邊都空就不出聲', describeColours({ fills: [], texts: [] }), '');
{
  const text = describeColours({ fills: ['#1F3864'], texts: ['#3D3D3D'] });
  ok('填色和文字色分開講', text.includes('填色：') && text.includes('文字色：'));
  ok('顏色列得出來', text.includes('#1F3864') && text.includes('#3D3D3D'));
  // 這句才是 AC-11 真正要的：讀到既有配色之後，不要再自作主張套一組。
  ok('明說沿用而不是另外套一組', text.includes('沿用'));
  ok('明說不要呼叫 ppt_apply_theme', text.includes('ppt_apply_theme'));
}
ok('只有填色時不會印出空的文字色欄位',
  describeColours({ fills: ['#1F3864'], texts: [] }).includes('文字色'), false);

fs.rmSync(tmp);
console.log(`\n${pass} passed, ${fail} failed`);
if (fail) process.exit(1);
