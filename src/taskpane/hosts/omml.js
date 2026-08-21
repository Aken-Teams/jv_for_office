/**
 * omml — LaTeX in, Office Math Markup out.
 *
 * ── Why this exists ──
 * The Word JavaScript API has no equation object. Not a limited one — none. The
 * only way an add-in can put mathematics into a document is to hand Word OOXML
 * containing OMML, which is what this file produces.
 *
 * That is worth the effort because the alternatives are all bad: a picture of an
 * equation cannot be edited, cannot be searched, and goes blurry when the
 * document is scaled; typed-out Unicode ("x^2 + sqrt(y)") is not mathematics, it
 * is a description of mathematics. What comes out of here is a real Word
 * equation object — the reader can click into it and change a symbol.
 *
 * And LaTeX is the right input, because it is the notation the model already
 * thinks in. Asking it to emit OMML directly would mean asking it to write a
 * verbose XML dialect from memory, which it would do less reliably than the
 * notation it has seen a million times.
 *
 * ── Scope ──
 * The subset that appears in engineering and financial documents: fractions,
 * scripts, roots, n-ary operators with limits, delimiters, matrices, accents,
 * named functions, Greek, and the usual operators. Not: custom macros, \\newcommand,
 * arrays with column specs, commutative diagrams, chemistry.
 *
 * Anything unrecognised falls through as literal text rather than throwing. A
 * formula that renders with one odd symbol in it is repairable by the person
 * reading it; a tool call that failed is not.
 *
 * ── This file touches no Office API ──
 * Which is the point: it is pure string→string, so test/word-equation.test.mjs
 * can drive every branch in Node without Word being involved.
 */

const GREEK = {
  alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', epsilon: 'ϵ', varepsilon: 'ε',
  zeta: 'ζ', eta: 'η', theta: 'θ', vartheta: 'ϑ', iota: 'ι', kappa: 'κ',
  lambda: 'λ', mu: 'μ', nu: 'ν', xi: 'ξ', pi: 'π', varpi: 'ϖ', rho: 'ρ',
  varrho: 'ϱ', sigma: 'σ', varsigma: 'ς', tau: 'τ', upsilon: 'υ', phi: 'ϕ',
  varphi: 'φ', chi: 'χ', psi: 'ψ', omega: 'ω',
  Gamma: 'Γ', Delta: 'Δ', Theta: 'Θ', Lambda: 'Λ', Xi: 'Ξ', Pi: 'Π',
  Sigma: 'Σ', Upsilon: 'Υ', Phi: 'Φ', Psi: 'Ψ', Omega: 'Ω',
};

const SYMBOLS = {
  times: '×', div: '÷', cdot: '⋅', pm: '±', mp: '∓', ast: '∗', star: '⋆',
  leq: '≤', le: '≤', geq: '≥', ge: '≥', neq: '≠', ne: '≠', ll: '≪', gg: '≫',
  approx: '≈', equiv: '≡', cong: '≅', sim: '∼', simeq: '≃', propto: '∝',
  infty: '∞', partial: '∂', nabla: '∇', forall: '∀', exists: '∃', neg: '¬',
  in: '∈', notin: '∉', ni: '∋', subset: '⊂', subseteq: '⊆', supset: '⊃',
  supseteq: '⊇', cup: '∪', cap: '∩', emptyset: '∅', varnothing: '∅',
  rightarrow: '→', to: '→', leftarrow: '←', gets: '←', leftrightarrow: '↔',
  Rightarrow: '⇒', Leftarrow: '⇐', Leftrightarrow: '⇔', mapsto: '↦',
  ldots: '…', dots: '…', cdots: '⋯', vdots: '⋮', ddots: '⋱',
  angle: '∠', perp: '⊥', parallel: '∥', therefore: '∴', because: '∵',
  degree: '°', circ: '∘', prime: '′', pm2: '±', oplus: '⊕', otimes: '⊗',
  langle: '⟨', rangle: '⟩', hbar: 'ℏ', ell: 'ℓ', Re: 'ℜ', Im: 'ℑ',
  aleph: 'ℵ', wp: '℘', surd: '√', top: '⊤', bot: '⊥', vee: '∨', wedge: '∧',
  sum: '∑', prod: '∏', int: '∫',   // only used when they appear without limits
};

/** Operators that take limits above and below (or after, when inline). */
const NARY = {
  sum: '∑', prod: '∏', coprod: '∐', bigcup: '⋃', bigcap: '⋂',
  int: '∫', iint: '∬', iiint: '∭', oint: '∮', bigoplus: '⨁', bigotimes: '⨂',
};
/** Integrals put their limits beside the sign; sums put them above and below. */
const SUBSCRIPT_STYLE = { int: true, iint: true, iiint: true, oint: true };

const FUNCS = new Set([
  'sin', 'cos', 'tan', 'cot', 'sec', 'csc', 'arcsin', 'arccos', 'arctan',
  'sinh', 'cosh', 'tanh', 'coth', 'log', 'ln', 'lg', 'exp', 'lim', 'limsup',
  'liminf', 'max', 'min', 'sup', 'inf', 'det', 'dim', 'ker', 'deg', 'gcd', 'arg', 'mod',
]);

/** Combining marks — they sit ON the character, which is what m:acc wants. */
const ACCENTS = {
  hat: '̂', widehat: '̂', bar: '̄', vec: '⃗',
  dot: '̇', ddot: '̈', tilde: '̃', widetilde: '̃',
  acute: '́', grave: '̀', check: '̌', breve: '̆',
};

const DELIMS = {
  '(': '(', ')': ')', '[': '[', ']': ']', '\\{': '{', '\\}': '}',
  '|': '|', '\\|': '‖', '\\langle': '⟨', '\\rangle': '⟩',
  '\\lceil': '⌈', '\\rceil': '⌉', '\\lfloor': '⌊', '\\rfloor': '⌋',
  '.': '',
};

const MATRIX_FENCES = {
  matrix: ['', ''], pmatrix: ['(', ')'], bmatrix: ['[', ']'],
  Bmatrix: ['{', '}'], vmatrix: ['|', '|'], Vmatrix: ['‖', '‖'],
  cases: ['{', ''],
};

const esc = (s) => String(s).replace(/[&<>"']/g, (c) => (
  { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]
));

// ── Tokenizer ─────────────────────────────────────────────────────────────────

/**
 * LaTeX into a flat token list.
 *
 * `\command` is one token including its backslash, so the parser can look a
 * command up without re-scanning. Whitespace between tokens is dropped — in
 * mathematics it carries no meaning, and keeping it would put stray gaps into
 * every rendered formula.
 */
function tokenize(src) {
  const out = [];
  let i = 0;
  const s = String(src);
  while (i < s.length) {
    const c = s[i];
    if (/\s/.test(c)) { i++; continue; }
    if (c === '\\') {
      const m = /^\\([A-Za-z]+|.)/.exec(s.slice(i));
      if (!m) { i++; continue; }
      out.push({ t: 'cmd', v: m[1] });
      i += m[0].length;
      continue;
    }
    if ('{}^_&'.includes(c)) { out.push({ t: c }); i++; continue; }
    // Digits clump: 3.14 should be one run, not four.
    const num = /^\d+(\.\d+)?/.exec(s.slice(i));
    if (num) { out.push({ t: 'txt', v: num[0] }); i += num[0].length; continue; }
    out.push({ t: 'txt', v: c });
    i++;
  }
  return out;
}

// ── Parser ────────────────────────────────────────────────────────────────────

/**
 * Tokens into a node tree.
 *
 * `stop` lets the same routine parse a whole formula, the inside of a `{}`, and
 * one cell of a matrix — the caller says which token ends the run.
 */
function parseNodes(toks, state, stop) {
  const nodes = [];
  while (state.i < toks.length) {
    const tk = toks[state.i];
    if (stop && stop(tk)) break;

    if (tk.t === '}') break;
    if (tk.t === '^' || tk.t === '_') {
      state.i++;
      const arg = parseGroup(toks, state);
      const base = nodes.pop() || { k: 'run', text: '' };
      // x_i^2 and x^2_i mean the same thing, so a second script folds into the
      // first rather than nesting — nesting renders as a script of a script.
      if (base.k === 'script') {
        nodes.push({
          k: 'script',
          base: base.base,
          sub: tk.t === '_' ? arg : base.sub,
          sup: tk.t === '^' ? arg : base.sup,
        });
      } else {
        nodes.push({ k: 'script', base, sub: tk.t === '_' ? arg : null, sup: tk.t === '^' ? arg : null });
      }
      continue;
    }
    if (tk.t === '{') {
      state.i++;
      const inner = parseNodes(toks, state, null);
      if (toks[state.i] && toks[state.i].t === '}') state.i++;
      nodes.push({ k: 'group', body: inner });
      continue;
    }
    if (tk.t === 'txt') { state.i++; nodes.push({ k: 'run', text: tk.v }); continue; }
    // A bare & outside a matrix is not a separator, it is the character itself.
    // parseMatrix stops on & before this line is ever reached, so the first
    // version's bare `continue` only ever discarded real content.
    if (tk.t === '&') { state.i++; nodes.push({ k: 'run', text: '&' }); continue; }

    // ── commands ──
    const name = tk.v;
    state.i++;

    if (name === 'frac' || name === 'dfrac' || name === 'tfrac' || name === 'cfrac') {
      nodes.push({ k: 'frac', num: parseGroup(toks, state), den: parseGroup(toks, state) });
      continue;
    }
    if (name === 'binom') {
      nodes.push({ k: 'binom', num: parseGroup(toks, state), den: parseGroup(toks, state) });
      continue;
    }
    if (name === 'sqrt') {
      let deg = null;
      if (toks[state.i] && toks[state.i].t === 'txt' && toks[state.i].v === '[') {
        state.i++;
        deg = parseNodes(toks, state, (t) => t.t === 'txt' && t.v === ']');
        if (toks[state.i]) state.i++;   // consume ]
      }
      nodes.push({ k: 'rad', deg, body: parseGroup(toks, state) });
      continue;
    }
    if (ACCENTS[name]) {
      nodes.push({ k: 'acc', chr: ACCENTS[name], body: parseGroup(toks, state) });
      continue;
    }
    if (name === 'overline') { nodes.push({ k: 'bar', pos: 'top', body: parseGroup(toks, state) }); continue; }
    if (name === 'underline') { nodes.push({ k: 'bar', pos: 'bot', body: parseGroup(toks, state) }); continue; }
    if (name === 'text' || name === 'mathrm' || name === 'operatorname' || name === 'mbox') {
      nodes.push({ k: 'text', body: parseGroup(toks, state) });
      continue;
    }
    if (name === 'left') {
      const open = readDelim(toks, state);
      const body = parseNodes(toks, state, (t) => t.t === 'cmd' && t.v === 'right');
      let close = '';
      if (toks[state.i] && toks[state.i].t === 'cmd' && toks[state.i].v === 'right') {
        state.i++;
        close = readDelim(toks, state);
      }
      nodes.push({ k: 'delim', open, close, body });
      continue;
    }
    if (name === 'begin') {
      const env = readWord(toks, state);
      if (MATRIX_FENCES[env]) { nodes.push(parseMatrix(toks, state, env)); continue; }
      // An environment we do not model: drop the wrapper, keep the contents.
      continue;
    }
    if (name === 'end') { readWord(toks, state); continue; }
    if (NARY[name]) {
      let sub = null;
      let sup = null;
      for (let guard = 0; guard < 2; guard++) {
        const nx = toks[state.i];
        if (!nx || (nx.t !== '_' && nx.t !== '^')) break;
        state.i++;
        if (nx.t === '_') sub = parseGroup(toks, state); else sup = parseGroup(toks, state);
      }
      nodes.push({
        k: 'nary', chr: NARY[name], sub, sup,
        under: !SUBSCRIPT_STYLE[name],
        body: parseNodes(toks, state, stop),
      });
      continue;
    }
    if (FUNCS.has(name)) {
      let sub = null;
      if (toks[state.i] && toks[state.i].t === '_') { state.i++; sub = parseGroup(toks, state); }
      nodes.push({ k: 'func', name, sub, body: parseNodes(toks, state, stop) });
      continue;
    }
    if (GREEK[name]) { nodes.push({ k: 'run', text: GREEK[name], italic: true }); continue; }
    if (SYMBOLS[name]) { nodes.push({ k: 'run', text: SYMBOLS[name] }); continue; }
    if (name === 'quad') { nodes.push({ k: 'run', text: ' ' }); continue; }
    if (name === 'qquad') { nodes.push({ k: 'run', text: '  ' }); continue; }
    if (name === ',' || name === ';' || name === ':' || name === '!' || name === ' ') {
      nodes.push({ k: 'run', text: name === '!' ? '' : ' ' });
      continue;
    }
    if (name === '\\') { nodes.push({ k: 'br' }); continue; }
    if (name === '{' || name === '}' || name === '%' || name === '$' || name === '#' || name === '_' || name === '&') {
      nodes.push({ k: 'run', text: name });
      continue;
    }
    // Unknown command: keep the word so the reader can see what was meant.
    nodes.push({ k: 'run', text: name });
  }
  return nodes;
}

/** The next single item: a braced group, or one token. */
function parseGroup(toks, state) {
  const tk = toks[state.i];
  if (!tk) return [];
  if (tk.t === '{') {
    state.i++;
    const inner = parseNodes(toks, state, null);
    if (toks[state.i] && toks[state.i].t === '}') state.i++;
    return inner;
  }
  const one = parseNodes(toks, state, () => true);
  if (one.length) return one;
  // parseNodes with an always-true stop consumed nothing; take the token itself.
  state.i++;
  if (tk.t === 'txt') return [{ k: 'run', text: tk.v }];
  if (tk.t === 'cmd') {
    if (GREEK[tk.v]) return [{ k: 'run', text: GREEK[tk.v], italic: true }];
    if (SYMBOLS[tk.v]) return [{ k: 'run', text: SYMBOLS[tk.v] }];
    return [{ k: 'run', text: tk.v }];
  }
  return [];
}

function readDelim(toks, state) {
  const tk = toks[state.i];
  if (!tk) return '';
  state.i++;
  if (tk.t === 'txt') return DELIMS[tk.v] !== undefined ? DELIMS[tk.v] : tk.v;
  if (tk.t === 'cmd') {
    const key = '\\' + tk.v;
    if (DELIMS[key] !== undefined) return DELIMS[key];
    return SYMBOLS[tk.v] || '';
  }
  return '';
}

/** `\begin{pmatrix}` — the environment name arrives as ordinary tokens. */
function readWord(toks, state) {
  if (!toks[state.i] || toks[state.i].t !== '{') return '';
  state.i++;
  let name = '';
  while (toks[state.i] && toks[state.i].t !== '}') {
    name += toks[state.i].v || '';
    state.i++;
  }
  if (toks[state.i]) state.i++;
  return name;
}

function parseMatrix(toks, state, env) {
  const rows = [];
  let row = [];
  for (;;) {
    const cell = parseNodes(toks, state, (t) =>
      t.t === '&' || (t.t === 'cmd' && (t.v === '\\' || t.v === 'end')));
    row.push(cell);
    const tk = toks[state.i];
    if (!tk) break;
    if (tk.t === '&') { state.i++; continue; }
    if (tk.t === 'cmd' && tk.v === '\\') { state.i++; rows.push(row); row = []; continue; }
    if (tk.t === 'cmd' && tk.v === 'end') { state.i++; readWord(toks, state); break; }
    break;
  }
  if (row.length && (row.length > 1 || row[0].length)) rows.push(row);
  const [open, close] = MATRIX_FENCES[env] || ['', ''];
  return { k: 'matrix', rows, open, close };
}

// ── Emitter ───────────────────────────────────────────────────────────────────

const run = (text, italic) =>
  '<m:r>' + (italic === false ? '<m:rPr><m:nor/></m:rPr>' : '')
  + `<m:t xml:space="preserve">${esc(text)}</m:t></m:r>`;

function emit(nodes) {
  return (nodes || []).map(emitNode).join('');
}

function emitNode(n) {
  switch (n.k) {
    case 'run': return run(n.text);
    case 'group': return emit(n.body);
    case 'text': return run(plain(n.body), false);
    case 'br': return '';   // line breaks inside one equation are not modelled
    case 'frac':
      return `<m:f><m:fPr><m:type m:val="bar"/></m:fPr><m:num>${emit(n.num)}</m:num><m:den>${emit(n.den)}</m:den></m:f>`;
    case 'binom':
      return '<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>'
        + `<m:f><m:fPr><m:type m:val="noBar"/></m:fPr><m:num>${emit(n.num)}</m:num><m:den>${emit(n.den)}</m:den></m:f>`
        + '</m:e></m:d>';
    case 'rad':
      return n.deg && n.deg.length
        ? `<m:rad><m:radPr><m:degHide m:val="0"/></m:radPr><m:deg>${emit(n.deg)}</m:deg><m:e>${emit(n.body)}</m:e></m:rad>`
        : `<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e>${emit(n.body)}</m:e></m:rad>`;
    case 'script': {
      const base = `<m:e>${emit([n.base])}</m:e>`;
      if (n.sub && n.sup) {
        return `<m:sSubSup>${base}<m:sub>${emit(n.sub)}</m:sub><m:sup>${emit(n.sup)}</m:sup></m:sSubSup>`;
      }
      if (n.sub) return `<m:sSub>${base}<m:sub>${emit(n.sub)}</m:sub></m:sSub>`;
      return `<m:sSup>${base}<m:sup>${emit(n.sup)}</m:sup></m:sSup>`;
    }
    case 'acc':
      return `<m:acc><m:accPr><m:chr m:val="${esc(n.chr)}"/></m:accPr><m:e>${emit(n.body)}</m:e></m:acc>`;
    case 'bar':
      return `<m:bar><m:barPr><m:pos m:val="${n.pos}"/></m:barPr><m:e>${emit(n.body)}</m:e></m:bar>`;
    case 'delim':
      return '<m:d><m:dPr>'
        + `<m:begChr m:val="${esc(n.open)}"/><m:endChr m:val="${esc(n.close)}"/>`
        + `</m:dPr><m:e>${emit(n.body)}</m:e></m:d>`;
    case 'nary':
      return '<m:nary><m:naryPr>'
        + `<m:chr m:val="${esc(n.chr)}"/>`
        + `<m:limLoc m:val="${n.under ? 'undOvr' : 'subSup'}"/>`
        + `<m:subHide m:val="${n.sub ? '0' : '1'}"/><m:supHide m:val="${n.sup ? '0' : '1'}"/>`
        + '</m:naryPr>'
        + `<m:sub>${emit(n.sub)}</m:sub><m:sup>${emit(n.sup)}</m:sup><m:e>${emit(n.body)}</m:e></m:nary>`;
    case 'func': {
      const nm = n.sub
        ? `<m:sSub><m:e>${run(n.name, false)}</m:e><m:sub>${emit(n.sub)}</m:sub></m:sSub>`
        : run(n.name, false);
      return `<m:func><m:funcPr/><m:fName>${nm}</m:fName><m:e>${emit(n.body)}</m:e></m:func>`;
    }
    case 'matrix': {
      const cols = Math.max(...n.rows.map((r) => r.length), 1);
      const mcs = '<m:mcs><m:mc><m:mcPr>'
        + `<m:count m:val="${cols}"/><m:mcJc m:val="center"/>`
        + '</m:mcPr></m:mc></m:mcs>';
      const body = '<m:m><m:mPr>' + mcs + '</m:mPr>'
        + n.rows.map((r) => '<m:mr>' + r.map((cell) => `<m:e>${emit(cell)}</m:e>`).join('') + '</m:mr>').join('')
        + '</m:m>';
      if (!n.open && !n.close) return body;
      return '<m:d><m:dPr>'
        + `<m:begChr m:val="${esc(n.open)}"/><m:endChr m:val="${esc(n.close)}"/>`
        + `</m:dPr><m:e>${body}</m:e></m:d>`;
    }
    default: return '';
  }
}

/** \text{...} is prose, so it is flattened back to characters. */
function plain(nodes) {
  return (nodes || []).map((n) => {
    if (n.k === 'run') return n.text;
    if (n.k === 'group' || n.k === 'text') return plain(n.body);
    return '';
  }).join('');
}

/** The maths itself, without any Word paragraph around it. */
export function latexToOmml(latex) {
  const toks = tokenize(latex);
  const nodes = parseNodes(toks, { i: 0 }, null);
  return `<m:oMath>${emit(nodes)}</m:oMath>`;
}

const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
  + 'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"';

/**
 * A complete OOXML package holding one equation.
 *
 * `display` wraps it in m:oMathPara, which is what makes Word treat it as a
 * block equation on its own centred line rather than as maths sitting inside a
 * sentence.
 */
export function equationOoxml(latex, display) {
  const math = latexToOmml(latex);
  const para = display
    ? `<w:p><m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr>${math}</m:oMathPara></w:p>`
    : `<w:p>${math}</w:p>`;
  return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
    + '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">'
    + '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document.xml"/>'
    + '</Relationships></pkg:xmlData></pkg:part>'
    + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">'
    + `<pkg:xmlData><w:document ${NS}><w:body>${para}</w:body></w:document></pkg:xmlData>`
    + '</pkg:part></pkg:package>';
}
