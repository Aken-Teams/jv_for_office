/**
 * Drive the two pure functions that decide how a paragraph goes into Word.
 *
 * Both exist because of failures seen in a real document, and both are pure
 * string work — so unlike everything else in the Word host, they can be pinned
 * without Word being involved.
 *
 *  · bodyParagraphOoxml — text written after a display equation was landing
 *    INSIDE the equation: Cambria Math, centred, and not wrapping, so it ran off
 *    the right edge of the page. Two rounds of fixes tried to correct the
 *    formatting afterwards; this one states the whole paragraph so there is
 *    nothing to inherit. If the XML it emits is malformed, Word rejects the
 *    insert wholesale and the fix silently stops working.
 *
 *  · headingRefused — 「注意事項：終端算力有限，模型需要壓縮或蒸餾……」 was being
 *    styled as a heading. A heading is a name, not a sentence.
 */
import fs from 'fs';

// Sliced rather than imported: these are internal to the host module, and
// exporting them purely for a test would widen its surface for no reason.
const src = fs.readFileSync(new URL('../src/taskpane/hosts/word.js', import.meta.url), 'utf8');
const cut = (from, to) => src.slice(src.indexOf(from), src.indexOf(to));
const mod = cut('const xmlEscape', 'function tocOoxml')
  + '\nexport { bodyParagraphOoxml, headingRefused };\n';
const tmp = new URL('./.paragraph-under-test.mjs', import.meta.url);
fs.writeFileSync(tmp, mod);
const { bodyParagraphOoxml, headingRefused } = await import(tmp);

function xmlProblem(xml) {
  const stack = [];
  let m;
  const re = /<(\/?)([A-Za-z:][\w:.-]*)([^>]*?)(\/?)>/g;
  while ((m = re.exec(xml))) {
    const [, closing, name, , selfClose] = m;
    if (selfClose === '/') continue;
    if (closing === '/') { if (stack.pop() !== name) return `</${name}> 對不上`; }
    else stack.push(name);
  }
  return stack.length ? `沒有關閉：${stack.join(', ')}` : '';
}
const visible = (xml) => (xml.match(/<w:t[^>]*>[\s\S]*?<\/w:t>/g) || [])
  .map((t) => t.replace(/<[^>]+>/g, ''))
  .join('')
  .replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"')
  .replace(/&apos;/g, "'").replace(/&amp;/g, '&');

let pass = 0, fail = 0;
const ok = (name, cond, detail) => {
  cond ? pass++ : fail++;
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}${cond ? '' : '  → ' + (detail || '')}`);
};

// ── bodyParagraphOoxml ────────────────────────────────────────────────────────
const CASES = [
  ['一般內文', '這是一段內文。', { style: 'Normal', align: 'Left' }, ['<w:pStyle w:val="Normal"/>', 'w:val="left"']],
  ['置中', '標題', { style: 'Normal', align: 'Centered' }, ['w:val="center"']],
  ['兩端對齊', '本文', { style: 'Normal', align: 'Justified' }, ['w:val="both"']],
  ['標題樣式', '六種方式對照', { style: 'Heading1', align: 'Left' }, ['<w:pStyle w:val="Heading1"/>']],
  ['首行縮排兩個字', '公文內容', { style: 'Normal', align: 'Justified', firstLineIndent: 24 }, ['w:firstLine="480"']],
  ['左縮排', '引文', { style: 'Normal', align: 'Left', leftIndent: 12 }, ['w:left="240"']],
  ['空白段落', '', { style: 'Normal', align: 'Left' }, []],
];
for (const [name, text, opts, needs] of CASES) {
  const xml = bodyParagraphOoxml(text, opts);
  const bad = xmlProblem(xml);
  const missing = needs.filter((n) => !xml.includes(n));
  const textOk = !text || visible(xml) === text;
  ok(`OOXML：${name}`, !bad && !missing.length && textOk,
    bad || (missing.length ? `少了 ${missing.join(', ')}` : `文字變成「${visible(xml)}」`));
}

// The whole point: no run properties, no maths, nothing carried in from a
// neighbour. A paragraph that mentions Cambria Math or oMath has inherited.
{
  const xml = bodyParagraphOoxml('答案：x = 3 時有最小值 2。', { style: 'Normal', align: 'Left' });
  ok('OOXML 不帶任何數學或字型繼承', !/oMath|Cambria|w:rFonts/.test(xml), xml.slice(0, 120));
}
{
  const xml = bodyParagraphOoxml('a < b & c > "d"', { style: 'Normal', align: 'Left' });
  ok('OOXML 跳脫危險字元', !xmlProblem(xml) && visible(xml) === 'a < b & c > "d"', visible(xml));
}

// ── headingRefused ────────────────────────────────────────────────────────────
const HEADINGS = [
  // [名稱, style, 文字, 應該被擋下嗎]
  ['短名稱可以當標題', 'Heading1', '六種方式對照', false],
  ['英數短名稱可以', 'Heading2', '邊緣 AI', false],
  ['句號結尾要擋', 'Heading3', '注意事項：定位是輔助而非取代。', true],
  ['問號結尾要擋', 'Heading2', '這樣做對嗎？', true],
  ['整段敘述要擋', 'Heading3',
    '注意事項：終端算力有限，模型需要壓縮或蒸餾，且裝置分散使版本更新的管理成本較高', true],
  ['內文樣式不受限制', 'Normal', '注意事項：定位是輔助而非取代，要保留人為確認的環節。', false],
  ['沒給樣式不受限制', '', '隨便什麼句子。', false],
  ['剛好在上限內', 'Heading1', '一'.repeat(30), false],
  ['超過上限一個字', 'Heading1', '一'.repeat(31), true],
  // 頓號在中文標題裡是正常的（「導論、背景與範圍」），不能一併擋掉。
  ['頓號不算敘述', 'Heading1', '導論、背景與範圍', false],
  ['冒號接短名稱還是標題', 'Heading2', '附錄：常用公式', false],
];
for (const [name, style, text, shouldBlock] of HEADINGS) {
  const why = headingRefused(style, text);
  ok(`標題守則：${name}`, Boolean(why) === shouldBlock, why || '沒有擋下來');
}

console.log(`\n${pass} passed, ${fail} failed`);
/**
 * 文字入口的三道正規化：字面換行、分開的底線、tab。
 *
 * 三者同一類——模型不是不知道怎麼排，是它在組字串的時候用了「在純文字裡有效、
 * 在 Word 裡沒效」的手法。tab 尤其隱蔽：Word 的 tab 是跳到下一個定位停駐點，
 * 而停駐點由段落樣式決定，模型看不到那些點在哪裡——所以同一段文字在不同樣式下
 * 落點不一樣，使用者看到的就是「有時候會因為 tab 而不小心把樣式跑版」。
 *
 * 這種事規則擋不住，入口擋得住。
 */
{
  /**
   * 三道正規化搬到 hosts/text.js 了——三個宿主共用，不再是 Word 專屬。
   * 所以這裡直接 import，不用再從 word.js 切片。
   */
  const { cleanText } = await import('../src/taskpane/hosts/text.js');

  ok('開頭的 tab 被拿掉（縮排該交給 left_indent）',
    cleanText('\t匯款帳戶') === '匯款帳戶');
  ok('中間的 tab 換成看得見的間隔，不是會跑的定位點',
    cleanText('戶名\t恆遠資訊') === '戶名\u3000恆遠資訊');
  ok('沒有 tab 的文字原樣不動', cleanText('一、付款方式') === '一、付款方式');

  ok('字面的反斜線 n 變成真的換行',
    cleanText('A。\\n2. B。') === 'A。\n2. B。');
  ok('真的換行不受影響', cleanText('A\nB') === 'A\nB');

  ok('分開的底線接成一條', /^簽章：_{6,}$/.test(cleanText('簽章：_ _ _ _ _ _')));
  ok('單獨的底線不動', cleanText('A_ B_ 正常') === 'A_ B_ 正常');
}

process.exit(fail ? 1 : 0);
