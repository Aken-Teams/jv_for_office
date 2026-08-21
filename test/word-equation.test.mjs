/**
 * Drive the LaTeX → OMML converter.
 *
 * Word has no equation API, so this conversion is entirely ours — and it is the
 * one piece of the Word add-in that can be checked completely without Word,
 * because it is a pure function from a string to a string. That makes the test
 * worth more than usual: the alternative to catching a malformed tag here is a
 * user watching Word reject the whole insertion with no explanation.
 *
 * Three things are checked, in increasing strictness:
 *   1. the output is well-formed XML with balanced tags
 *   2. the structural element the notation calls for is actually present
 *      (a fraction produces <m:f>, a matrix produces <m:m>, and so on)
 *   3. the visible characters survive — a formula that renders as an empty box
 *      is worse than one that renders slightly wrong
 */
import { latexToOmml, equationOoxml } from '../src/taskpane/hosts/omml.js';

/**
 * The characters a reader would actually see.
 *
 * Asserted instead of raw substrings because multi-letter identifiers split
 * across runs legitimately — `PV` is two italic variables in every maths
 * typesetting system, LaTeX and Word's own editor included. Checking
 * `xml.includes('PV')` would pin an implementation detail; checking the visible
 * text pins the thing that matters.
 */
function visibleText(xml) {
  return (xml.match(/<m:t[^>]*>[\s\S]*?<\/m:t>/g) || [])
    .map((t) => t.replace(/<[^>]+>/g, ''))
    .join('')
    .replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"').replace(/&apos;/g, "'")
    .replace(/&amp;/g, '&');
}

/** Tag balance, without pulling in an XML parser. */
function xmlProblem(xml) {
  const stack = [];
  const re = /<(\/?)([A-Za-z:][\w:.-]*)([^>]*?)(\/?)>/g;
  let m;
  while ((m = re.exec(xml))) {
    const [, closing, name, attrs, selfClose] = m;
    if (selfClose === '/') continue;
    if (closing === '/') {
      if (stack.pop() !== name) return `</${name}> 對不上`;
    } else {
      stack.push(name);
    }
    if (attrs.includes('<')) return `${name} 的屬性裡有沒跳脫的 <`;
  }
  return stack.length ? `沒有關閉：${stack.join(', ')}` : '';
}

const CASES = [
  // [名稱, LaTeX, 必須出現的元素, 必須看得到的字元]
  ['分數', String.raw`\frac{a}{b}`, ['<m:f>', '<m:num>', '<m:den>'], ['a', 'b']],
  ['上標', String.raw`x^{2}`, ['<m:sSup>', '<m:sup>'], ['x', '2']],
  ['下標', String.raw`a_{ij}`, ['<m:sSub>', '<m:sub>'], ['a', 'ij']],
  ['上下標合併成一個', String.raw`x_i^2`, ['<m:sSubSup>'], ['x', 'i', '2']],
  ['上下標反序也合併', String.raw`x^2_i`, ['<m:sSubSup>'], ['x', 'i', '2']],
  ['平方根', String.raw`\sqrt{x+1}`, ['<m:rad>', 'degHide m:val="1"'], ['x', '+', '1']],
  ['n 次方根', String.raw`\sqrt[3]{8}`, ['<m:rad>', 'degHide m:val="0"', '<m:deg>'], ['3', '8']],
  ['求和含上下限', String.raw`\sum_{i=1}^{n} i^2`, ['<m:nary>', 'undOvr', 'm:chr m:val="∑"'], ['i', '1', 'n']],
  ['積分的限在旁邊', String.raw`\int_0^1 f(x) dx`, ['<m:nary>', 'subSup', 'm:chr m:val="∫"'], ['0', '1', 'f']],
  ['連乘', String.raw`\prod_{k=1}^{m} a_k`, ['<m:nary>', 'm:chr m:val="∏"'], ['k', 'm']],
  ['希臘字母', String.raw`\alpha + \beta = \Gamma`, [], ['α', 'β', 'Γ']],
  ['運算子', String.raw`a \times b \leq c \neq d \pm e`, [], ['×', '≤', '≠', '±']],
  ['無窮與偏微分', String.raw`\frac{\partial f}{\partial x} \to \infty`, ['<m:f>'], ['∂', '∞', '→']],
  ['括號自動放大', String.raw`\left( \frac{a}{b} \right)`, ['<m:d>', 'begChr m:val="("', 'endChr m:val=")"'], ['a', 'b']],
  ['方括號', String.raw`\left[ x \right]`, ['begChr m:val="["', 'endChr m:val="]"'], ['x']],
  ['函數', String.raw`\sin(x) + \log(y)`, ['<m:func>', '<m:fName>'], ['sin', 'log', 'x', 'y']],
  ['極限', String.raw`\lim_{x \to 0} \frac{\sin x}{x}`, ['<m:func>', '<m:sSub>'], ['lim', '0', '→']],
  ['矩陣', String.raw`\begin{pmatrix} a & b \\ c & d \end{pmatrix}`, ['<m:m>', '<m:mr>', 'begChr m:val="("'], ['a', 'b', 'c', 'd']],
  ['行列式', String.raw`\begin{vmatrix} 1 & 2 \\ 3 & 4 \end{vmatrix}`, ['<m:m>', 'begChr m:val="|"'], ['1', '2', '3', '4']],
  ['重音符號', String.raw`\bar{x} + \hat{y} + \vec{v}`, ['<m:acc>'], ['x', 'y', 'v']],
  ['上劃線', String.raw`\overline{AB}`, ['<m:bar>', 'm:pos m:val="top"'], ['AB']],
  ['文字', String.raw`x = \text{速度} \times t`, ['<m:nor/>'], ['速度', '×']],
  ['二項式', String.raw`\binom{n}{k}`, ['<m:f>', 'noBar'], ['n', 'k']],
  ['巢狀分數', String.raw`\frac{\frac{a}{b}}{\frac{c}{d}}`, ['<m:f>'], ['a', 'b', 'c', 'd']],
  ['深層巢狀', String.raw`\sqrt{\frac{\sum_{i=1}^{n}(x_i-\bar{x})^2}{n-1}}`, ['<m:rad>', '<m:f>', '<m:nary>', '<m:acc>'], ['n', 'x']],
  ['真實案例：標準差', String.raw`\sigma = \sqrt{\frac{1}{N}\sum_{i=1}^{N}(x_i-\mu)^2}`, ['<m:rad>', '<m:nary>'], ['σ', 'μ', 'N']],
  ['真實案例：現值', String.raw`PV = \sum_{t=1}^{n} \frac{CF_t}{(1+r)^t}`, ['<m:nary>', '<m:f>'], ['PV', 'CF', 'r']],

  // ── 不該讓它壞掉的輸入 ──
  ['空字串', '', [], []],
  ['只有純文字', 'E = mc^2', ['<m:sSup>'], ['E', 'm', 'c', '2']],
  ['不認得的指令不會炸', String.raw`\foobar{x} + 1`, [], ['x', '1']],
  ['沒關的大括號', String.raw`\frac{a}{b`, ['<m:f>'], ['a', 'b']],
  ['沒有 \\right 的 \\left', String.raw`\left( x`, ['<m:d>'], ['x']],
  ['XML 危險字元有跳脫', String.raw`a < b & c > d`, ['&lt;', '&amp;', '&gt;'], ['<', '&', '>']],
  ['連續上標', String.raw`x^{y^{z}}`, ['<m:sSup>'], ['x', 'y', 'z']],
];

let pass = 0, fail = 0;
for (const [name, latex, needs, chars] of CASES) {
  const problems = [];
  let xml = '';
  try {
    xml = latexToOmml(latex);
  } catch (e) {
    problems.push('丟出例外：' + e.message);
  }
  if (!problems.length) {
    const bad = xmlProblem(xml);
    if (bad) problems.push('XML 不合法：' + bad);
    for (const n of needs) if (!xml.includes(n)) problems.push(`少了 ${n}`);
    const seen = visibleText(xml);
    for (const c of chars) if (!seen.includes(c)) problems.push(`看不到「${c}」`);
  }
  problems.length ? fail++ : pass++;
  console.log(`${problems.length ? 'FAIL' : 'PASS'}  ${name}${problems.length ? '  → ' + problems.join('；') : ''}`);
}

// The package wrapper is what Word actually receives; a malformed one is
// rejected wholesale with no clue as to which part was wrong.
const pkg = equationOoxml(String.raw`\frac{a}{b}`, true);
const pkgChecks = [
  ['整包是合法 XML', !xmlProblem(pkg)],
  ['宣告了 math 命名空間', pkg.includes('officeDocument/2006/math')],
  ['display 模式包了 oMathPara', pkg.includes('<m:oMathPara>')],
  ['行內模式不包 oMathPara', !equationOoxml('x', false).includes('<m:oMathPara>')],
  ['有 document.xml 這個 part', pkg.includes('/word/document.xml')],
];
for (const [name, ok] of pkgChecks) {
  ok ? pass++ : fail++;
  console.log(`${ok ? 'PASS' : 'FAIL'}  ${name}`);
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
