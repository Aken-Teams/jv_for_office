/**
 * hosts/word — the Office.js half of the bridge, for Word.
 *
 * Same contract as hosts/excel (see hosts/index.js), and deliberately the same
 * shape, but Word is not a spreadsheet with prose in it and the differences that
 * matter are these three:
 *
 *  1. **The address is a paragraph number.** Word has no A1. Everything the
 *     model reads comes back numbered, and everything it writes is aimed at
 *     those numbers — which is also what makes [[段落 42]] a citation the reader
 *     can click.
 *
 *  2. **Writes go in as tracked changes.** This is the whole reason the Word
 *     add-in is safer than the Excel one: Word has a native review mechanism
 *     and an add-in can turn it on, so a rewrite arrives as a revision the
 *     person accepts or rejects line by line rather than as text that has
 *     already replaced their own. PRD 3.3 makes this the default and so does
 *     this module — see withTracking().
 *
 *  3. **Undo is exact.** Excel's snapshot restores formulas and number formats
 *     and admits it cannot bring fills and borders back. Word's before-image is
 *     the body's OOXML, which is everything — text, styles, tables, formatting.
 *     Restoring it is one call.
 *
 * Reads are capped for the same reason as in Excel: an unbounded read of a
 * 400-page contract would blow the context window, then the request body limit,
 * then the token budget, in that order.
 */
import { t } from '../i18n.js';
import { makeScriptGuard } from './guard.js';

/** Cheap and synchronous; Office.js answers from a local table. */
export function hasApi(version) {
  try {
    return Office.context.requirements.isSetSupported('WordApi', version);
  } catch {
    return false;
  }
}

const BASE_TOOLS = [
  'word_get_overview',
  'word_get_selection',
  'word_read_range',
  'word_search',
  'word_write_range',
  'word_insert_text',
  'word_delete_range',
  'word_format_range',
  'word_insert_table',
  'word_insert_toc',
  'word_manage_list',
  'word_header_footer',
  'word_insert_break',
  'word_insert_chart',
  'word_replace_all',
  // Compute and edit inside Word instead of hauling the document through the
  // model's context. Same rationale as excel_run_script.
  'word_run_script',
  // Answered by the pane's own UI, not by Office.js.
  'word_ask_user',
];

/**
 * Tools this build can execute ON THIS HOST.
 *
 * Tracked changes need WordApi 1.4 and comments need 1.4 as well. Rather than
 * let the agent call them and fail on an older build, they are simply not
 * offered — the server filters what the model sees by this list, so an old Word
 * degrades to "fewer abilities" instead of a loop of failed calls.
 *
 * Note what this means for a build below 1.4: writes still work, but they are
 * DIRECT, and the pane says so. That is the honest degradation — refusing to
 * write at all on an old Word would be worse, and pretending it was tracked
 * would be a lie about the one guarantee this host is built on.
 */
export function supportedTools() {
  const tools = [...BASE_TOOLS];
  if (hasApi('1.4')) tools.push('word_tracked_changes', 'word_comment');
  return tools;
}

/** Ceilings per read. Prose is denser than cells, so the character cap binds first. */
const MAX_READ_CHARS = 120000;
const MAX_READ_PARAGRAPHS = 600;
const MAX_SEARCH_RESULTS = 50;
/** Characters of each paragraph shown in the overview outline. */
const OUTLINE_SNIPPET = 60;
/** How much of the opening the overview ships with the first message. */
const OVERVIEW_PREVIEW_CHARS = 1500;

// ── Paragraph plumbing ────────────────────────────────────────────────────────

/**
 * Every paragraph, numbered from 1, with the style that says what it is.
 *
 * One load and one sync for the whole document. Word documents are hundreds of
 * paragraphs, not hundreds of thousands of cells, so reading the lot once and
 * working in JavaScript is both simpler and faster than a query per paragraph —
 * the same lesson excel_run_script exists to teach.
 */
async function loadParagraphs(ctx) {
  const paras = ctx.document.body.paragraphs;
  paras.load('items/text,items/style,items/styleBuiltIn,items/isListItem');
  await ctx.sync();
  return paras.items;
}

/** Heading levels come back as 'Heading 1' / styleBuiltIn 'Heading1'. */
function headingLevel(p) {
  const s = String(p.styleBuiltIn || p.style || '');
  const m = /^Heading\s*([1-9])$/i.exec(s.replace(/\s+/g, ' ').trim());
  return m ? Number(m[1]) : 0;
}

const snip = (text, n) => {
  const s = String(text || '').replace(/\s+/g, ' ').trim();
  return s.length > n ? s.slice(0, n) + '…' : s;
};

/**
 * Resolve `from`/`to` against the document, 1-based and inclusive.
 *
 * Out-of-range numbers are clamped rather than rejected: the model is working
 * from a count it read a few calls ago, and failing the whole call because the
 * document got two paragraphs shorter helps nobody.
 */
function resolveRange(items, args) {
  const total = items.length;
  let from = Math.max(1, Math.floor(Number(args.from ?? args.paragraph ?? 1)));
  let to = args.to === undefined && args.paragraph === undefined && args.from === undefined
    ? total
    : Math.floor(Number(args.to ?? args.from ?? args.paragraph));
  if (!Number.isFinite(to)) to = from;
  from = Math.min(from, total);
  to = Math.min(Math.max(to, from), total);
  return { from, to };
}

/** Numbered lines, which is what makes [[段落 42]] mean anything. */
function renderParagraphs(items, from, to) {
  const out = [];
  let chars = 0;
  let stoppedAt = 0;
  for (let i = from; i <= to; i++) {
    const p = items[i - 1];
    if (!p) break;
    const level = headingLevel(p);
    const mark = level ? `[H${level}] ` : (p.isListItem ? '• ' : '');
    const line = `${i}\t${mark}${String(p.text || '').trim()}`;
    if (chars + line.length > MAX_READ_CHARS || (i - from) >= MAX_READ_PARAGRAPHS) { stoppedAt = i; break; }
    chars += line.length;
    out.push(line);
  }
  let text = out.join('\n');
  if (stoppedAt) {
    text += `\n…（讀到第 ${stoppedAt - 1} 段就達到上限。要繼續請用 from=${stoppedAt} 再讀一次，`
      + '或改用 word_run_script 在 Word 裡算完只回結論。）';
  }
  return text;
}

// ── Tracked changes ───────────────────────────────────────────────────────────

const TRACK_MODES = { off: 'Off', mine: 'TrackMineOnly', all: 'TrackAll' };

/**
 * Run a write with revision marking on, then put the setting back.
 *
 * The point of restoring it: change tracking is the PERSON's setting, shown in
 * their ribbon, and an add-in that silently leaves it on has changed how their
 * next hour of typing behaves. Turning it on for the duration of our own write
 * gets the revision marks — which is what we actually want — without taking
 * over the document's configuration.
 *
 * `track: false` in the args is the escape hatch for "just replace it", and the
 * result text always says which of the two happened, because "did that go in as
 * a revision?" is the question the whole design rests on.
 */
async function withTracking(ctx, args, fn) {
  const wanted = args.track === false ? null : TRACK_MODES.all;
  if (!wanted || !hasApi('1.4')) {
    const value = await fn();
    return { value, tracked: false, reason: hasApi('1.4') ? 'track=false' : 'unsupported' };
  }

  const doc = ctx.document;
  doc.load('changeTrackingMode');
  await ctx.sync();
  const previous = String(doc.changeTrackingMode || TRACK_MODES.off);

  if (previous !== TRACK_MODES.all) {
    doc.changeTrackingMode = TRACK_MODES.all;
    await ctx.sync();
  }
  try {
    const value = await fn();
    await ctx.sync();
    return { value, tracked: true };
  } finally {
    if (previous !== TRACK_MODES.all) {
      doc.changeTrackingMode = previous;
      await ctx.sync();
    }
  }
}

/** One sentence appended to every write, so the review state is never a guess. */
function trackingNote(r) {
  if (r.tracked) return '\n（以追蹤修訂寫入，可在 Word 的「校閱」裡逐條接受或拒絕。）';
  if (r.reason === 'unsupported') {
    return '\n（這個 Word 版本不支援用增益集開啟追蹤修訂，所以是直接寫入。要退回請用側邊欄的復原鍵。）';
  }
  return '\n（依要求直接寫入，沒有標成修訂。要退回請用側邊欄的復原鍵。）';
}

// ── Tools: reading ────────────────────────────────────────────────────────────

/**
 * Structure first, content second — the same bargain as excel_get_overview.
 *
 * The outline is the useful half: a model that knows the document runs
 * 「1 摘要 / 2 範圍 / 14 付款條件」 can ask for paragraph 14 instead of reading
 * forty pages to find it.
 */
async function getOverview(args) {
  return Word.run(async (ctx) => {
    const doc = ctx.document;
    const props = doc.properties;
    props.load('title,lastAuthor');
    const body = doc.body;
    body.load('text');
    if (hasApi('1.4')) doc.load('changeTrackingMode');
    await ctx.sync();

    const items = await loadParagraphs(ctx);
    const text = String(body.text || '');
    const words = text.trim() ? text.trim().split(/\s+|(?<=[一-鿿])/).filter(Boolean).length : 0;

    const outline = items
      .map((p, i) => ({ p, i: i + 1, level: headingLevel(p) }))
      .filter(x => x.level > 0)
      .map(x => `${'  '.repeat(x.level - 1)}${x.i}\t${'#'.repeat(x.level)} ${snip(x.p.text, OUTLINE_SNIPPET)}`);

    const parts = [
      `文件：${props.title || '（未命名）'}`,
      `段落數：${items.length}　字數：約 ${words}`,
    ];
    if (hasApi('1.4')) {
      const mode = String(doc.changeTrackingMode || TRACK_MODES.off);
      parts.push(`追蹤修訂：${mode === TRACK_MODES.off ? '目前關閉（本增益集寫入時會自行開啟）' : mode}`);
    }
    parts.push('');
    parts.push(outline.length
      ? `章節大綱（段落編號＋標題）：\n${outline.join('\n')}`
      : '（這份文件沒有使用標題樣式，所以沒有大綱。需要定位時請用 word_search。）');

    const previewChars = Math.max(200, Math.min(Number(args.preview_chars) || OVERVIEW_PREVIEW_CHARS, 4000));
    parts.push('');
    parts.push(`開頭預覽：\n${renderParagraphs(items, 1, Math.min(items.length, 12)).slice(0, previewChars)}`);
    parts.push('');
    parts.push('段落編號是這份文件的位址：讀取、改寫、格式都用它，引用時寫成 [[段落 42]] 使用者點了會跳過去。');
    return parts.join('\n');
  });
}

/**
 * What the person had selected when they asked.
 *
 * This is the whole answer to 「這段幫我改一下」 — they cannot always name the
 * paragraph, but the selection says which one they mean. The paragraph numbers
 * come from matching against the body list, so the model gets an address it can
 * write to rather than a loose quotation.
 */
async function getSelectionInfo(ctx) {
  const sel = ctx.document.getSelection();
  sel.load('text');
  const selParas = sel.paragraphs;
  selParas.load('items/text');
  await ctx.sync();

  const items = await loadParagraphs(ctx);
  const wanted = selParas.items.map(p => String(p.text || ''));
  let from = 0;
  if (wanted.length) {
    // Match the run of selected paragraph texts against the document list. Word
    // gives no stable paragraph id below WordApi 1.6, and an exact run of texts
    // is unambiguous in every real document — a false match needs the same
    // several paragraphs repeated verbatim, and the cost of one would be a
    // heading shown in a status bar, not a misplaced write.
    for (let i = 0; i + wanted.length <= items.length; i++) {
      let hit = true;
      for (let j = 0; j < wanted.length; j++) {
        if (String(items[i + j].text || '') !== wanted[j]) { hit = false; break; }
      }
      if (hit) { from = i + 1; break; }
    }
  }
  const to = from ? from + wanted.length - 1 : 0;

  // The heading the selection sits under, for the bar at the top of the pane.
  let heading = '';
  if (from) {
    for (let i = from - 1; i >= 1; i--) {
      if (headingLevel(items[i - 1])) { heading = snip(items[i - 1].text, 30); break; }
    }
  }
  return { sel, items, from, to, heading, text: String(sel.text || '') };
}

async function getSelection() {
  return Word.run(async (ctx) => {
    const info = await getSelectionInfo(ctx);
    if (!info.text.trim()) return '使用者目前沒有選取任何文字（游標可能只是停在某處）。';
    const where = info.from ? `段落 ${info.from}${info.to > info.from ? `-${info.to}` : ''}` : '（無法對到段落編號）';
    return `使用者選取的範圍：${where}${info.heading ? `，位於「${info.heading}」底下` : ''}\n`
      + `字數：約 ${info.text.length}\n內容：\n${info.text}`;
  });
}

async function readRange(args) {
  return Word.run(async (ctx) => {
    const items = await loadParagraphs(ctx);
    if (!items.length) return '這份文件是空的。';

    // Addressing by heading, because that is how a person says it: 「把第三章
    // 讀一下」 rather than 「讀第 47 到 61 段」.
    if (args.heading) {
      const needle = String(args.heading).trim();
      const at = items.findIndex(p => headingLevel(p) && String(p.text || '').trim().includes(needle));
      if (at < 0) return `找不到標題「${needle}」。用 word_get_overview 可以看到完整大綱。`;
      const level = headingLevel(items[at]);
      let end = items.length;
      for (let i = at + 1; i < items.length; i++) {
        const l = headingLevel(items[i]);
        if (l && l <= level) { end = i; break; }
      }
      return `「${snip(items[at].text, 40)}」（段落 ${at + 1}-${end}）：\n${renderParagraphs(items, at + 1, end)}`;
    }

    const { from, to } = resolveRange(items, args);
    return `段落 ${from}-${to}（全文共 ${items.length} 段）：\n${renderParagraphs(items, from, to)}`;
  });
}

async function search(args) {
  const query = String(args.query || '').trim();
  if (!query) throw new Error('query 是必要的');
  const limit = Math.min(Math.max(1, Number(args.max_results) || MAX_SEARCH_RESULTS), MAX_SEARCH_RESULTS);

  return Word.run(async (ctx) => {
    const items = await loadParagraphs(ctx);
    const flags = args.match_case ? '' : 'i';
    let re;
    try {
      re = new RegExp(args.regex ? query : query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags);
    } catch (e) {
      throw new Error(`query 不是合法的規則運算式：${e.message}`);
    }

    const hits = [];
    for (let i = 0; i < items.length && hits.length < limit; i++) {
      const text = String(items[i].text || '');
      if (!re.test(text)) continue;
      hits.push(`${i + 1}\t${snip(text, 120)}`);
    }
    if (!hits.length) return `找不到「${query}」。`;
    return `找到 ${hits.length} 段含有「${query}」${hits.length >= limit ? `（只列出前 ${limit} 筆）` : ''}：\n`
      + `段落\t內容\n${hits.join('\n')}`;
  });
}

// ── Tools: writing ────────────────────────────────────────────────────────────

/**
 * What a write is about to cost, for the confirmation card.
 *
 * The shell shows this before anyone clicks: how many paragraphs get replaced
 * and how much text disappears. 「改寫第 12-14 段」 tells you nothing about
 * whether you are losing a sentence or three pages.
 */
export async function inspectWrite(_tool, args) {
  return Word.run(async (ctx) => {
    const items = await loadParagraphs(ctx);
    if (!items.length) return { ok: false, error: '這份文件是空的，沒有可以取代的段落。' };
    const { from, to } = resolveRange(items, args);
    let occupied = 0;
    let chars = 0;
    for (let i = from; i <= to; i++) {
      const text = String(items[i - 1].text || '').trim();
      if (text) { occupied++; chars += text.length; }
    }
    return {
      ok: true,
      address: to > from ? `段落 ${from}-${to}` : `段落 ${from}`,
      occupied,
      // The shell's risk line counts 「格」; for prose the honest unit is
      // characters, and it is reported through the same field so the card needs
      // no host-specific branch.
      hadFormula: 0,
      chars,
    };
  });
}

/** Split incoming text into paragraphs the way the model most likely meant it. */
function asParagraphs(args) {
  if (Array.isArray(args.paragraphs)) return args.paragraphs.map(String);
  const text = args.text === undefined ? '' : String(args.text);
  return text.split(/\r?\n/);
}

async function writeRange(args) {
  const lines = asParagraphs(args);
  return Word.run(async (ctx) => {
    const items = await loadParagraphs(ctx);
    if (!items.length) throw new Error('這份文件是空的，沒有可以取代的段落。');
    const { from, to } = resolveRange(items, args);

    const result = await withTracking(ctx, args, async () => {
      const target = items[from - 1];
      // Replace the first paragraph's text, then add the rest after it, then
      // drop what is left of the old range. Doing it in this order keeps the
      // first paragraph's own style — which is what makes a rewritten heading
      // still look like a heading.
      target.insertText(lines[0] ?? '', 'Replace');
      let anchor = target;
      for (let i = 1; i < lines.length; i++) anchor = anchor.insertParagraph(lines[i], 'After');
      for (let i = from + 1; i <= to; i++) items[i - 1].delete();
      await ctx.sync();
      return null;
    });

    return `已改寫段落 ${from}${to > from ? `-${to}` : ''}，換成 ${lines.length} 段。` + trackingNote(result);
  });
}

async function insertText(args) {
  const lines = asParagraphs(args);
  const at = String(args.at || 'end');
  return Word.run(async (ctx) => {
    const items = await loadParagraphs(ctx);
    const result = await withTracking(ctx, args, async () => {
      let anchor;
      if (at === 'start' || (at === 'before' && !args.paragraph)) {
        anchor = ctx.document.body.insertParagraph(lines[0] ?? '', 'Start');
      } else if (at === 'end' || !items.length) {
        anchor = ctx.document.body.insertParagraph(lines[0] ?? '', 'End');
      } else {
        const { from } = resolveRange(items, args);
        anchor = items[from - 1].insertParagraph(lines[0] ?? '', at === 'before' ? 'Before' : 'After');
      }
      for (let i = 1; i < lines.length; i++) anchor = anchor.insertParagraph(lines[i], 'After');
      if (args.style) {
        // Applied to the run we just inserted only — never to what was there.
        anchor.styleBuiltIn = String(args.style).replace(/\s+/g, '');
      }
      await ctx.sync();
      return null;
    });
    const where = at === 'end' ? '文件最後' : at === 'start' ? '文件開頭' : `段落 ${args.paragraph || args.from} ${at === 'before' ? '之前' : '之後'}`;
    return `已在${where}插入 ${lines.length} 段。` + trackingNote(result);
  });
}

async function deleteRange(args) {
  return Word.run(async (ctx) => {
    const items = await loadParagraphs(ctx);
    if (!items.length) throw new Error('這份文件是空的。');
    const { from, to } = resolveRange(items, args);
    const result = await withTracking(ctx, args, async () => {
      for (let i = from; i <= to; i++) items[i - 1].delete();
      await ctx.sync();
      return null;
    });
    return `已刪除段落 ${from}${to > from ? `-${to}` : ''}（共 ${to - from + 1} 段）。` + trackingNote(result);
  });
}

const ALIGNMENTS = { left: 'Left', center: 'Centered', centered: 'Centered', right: 'Right', justify: 'Justified', justified: 'Justified' };

/**
 * Formatting.
 *
 * Not wrapped in tracked changes on purpose: Word records formatting revisions
 * too, and a document where every style tweak is a revision to accept is a
 * document nobody reviews properly. The line this add-in draws is the same one
 * Excel draws — ask about what destroys, let appearance through — and formatting
 * is fully covered by the pane's own undo.
 */
async function formatRange(args) {
  return Word.run(async (ctx) => {
    const items = await loadParagraphs(ctx);
    if (!items.length) throw new Error('這份文件是空的。');
    const { from, to } = resolveRange(items, args);
    const applied = [];

    for (let i = from; i <= to; i++) {
      const p = items[i - 1];
      if (args.style) p.styleBuiltIn = String(args.style).replace(/\s+/g, '');
      if (args.alignment) p.alignment = ALIGNMENTS[String(args.alignment).toLowerCase()] || 'Left';
      if (args.left_indent !== undefined) p.leftIndent = Number(args.left_indent);
      if (args.first_line_indent !== undefined) p.firstLineIndent = Number(args.first_line_indent);
      if (args.line_spacing !== undefined) p.lineSpacing = Number(args.line_spacing);
      if (args.space_before !== undefined) p.spaceBefore = Number(args.space_before);
      if (args.space_after !== undefined) p.spaceAfter = Number(args.space_after);
      const f = p.font;
      if (args.bold !== undefined) f.bold = !!args.bold;
      if (args.italic !== undefined) f.italic = !!args.italic;
      if (args.underline !== undefined) f.underline = args.underline ? 'Single' : 'None';
      if (args.font) f.name = String(args.font);
      if (args.size !== undefined) f.size = Number(args.size);
      if (args.color) f.color = String(args.color);
      if (args.highlight) f.highlightColor = String(args.highlight);
    }
    for (const k of ['style', 'alignment', 'bold', 'italic', 'underline', 'font', 'size', 'color', 'highlight',
      'left_indent', 'first_line_indent', 'line_spacing', 'space_before', 'space_after']) {
      if (args[k] !== undefined) applied.push(k);
    }
    await ctx.sync();
    return `已對段落 ${from}${to > from ? `-${to}` : ''} 套用：${applied.join('、') || '（沒有指定任何格式）'}。`;
  });
}

async function manageList(args) {
  const kind = String(args.type || 'bullet').toLowerCase();
  return Word.run(async (ctx) => {
    const items = await loadParagraphs(ctx);
    const { from, to } = resolveRange(items, args);
    if (kind === 'none') {
      for (let i = from; i <= to; i++) items[i - 1].detachFromList();
      await ctx.sync();
      return `已取消段落 ${from}-${to} 的清單格式。`;
    }
    const first = items[from - 1];
    const list = first.startNewList();
    list.load('id');
    await ctx.sync();
    if (kind === 'number' || kind === 'ordered') list.setLevelNumbering(0, 'Arabic', ['.']);
    else list.setLevelBullet(0, 'Solid');
    for (let i = from + 1; i <= to; i++) items[i - 1].attachToList(list.id, 0);
    await ctx.sync();
    return `已把段落 ${from}-${to} 設成${kind === 'number' || kind === 'ordered' ? '編號' : '項目符號'}清單。`;
  });
}

async function insertTable(args) {
  const values = Array.isArray(args.values) ? args.values.map(r => (Array.isArray(r) ? r.map(String) : [String(r)])) : [];
  if (!values.length) throw new Error('values 是必要的，格式是二維陣列（第一列通常是表頭）。');
  const cols = Math.max(...values.map(r => r.length));
  const rows = values.map(r => [...r, ...Array(cols - r.length).fill('')]);

  return Word.run(async (ctx) => {
    const items = await loadParagraphs(ctx);
    const result = await withTracking(ctx, args, async () => {
      let table;
      if (args.paragraph && items.length) {
        const { from } = resolveRange(items, args);
        table = items[from - 1].insertTable(rows.length, cols, String(args.at) === 'before' ? 'Before' : 'After', rows);
      } else {
        table = ctx.document.body.insertTable(rows.length, cols, 'End', rows);
      }
      table.styleBuiltIn = String(args.style || 'GridTable4_Accent1').replace(/\s+/g, '');
      table.headerRowCount = args.header === false ? 0 : 1;
      await ctx.sync();
      return null;
    });
    return `已插入 ${rows.length} 列 × ${cols} 欄的表格。` + trackingNote(result);
  });
}

/**
 * A real table of contents — which in Word means a field, not text.
 *
 * ── Why the first version came out empty ──
 * It inserted the TOC field and nothing else. That is technically a correct
 * table of contents, and it displays as a blank space with the word 目錄 above
 * it, because a field shows nothing until Word evaluates it — and Word does not
 * evaluate a freshly inserted one. Telling the user to press 「更新目錄」 was a
 * true instruction attached to something that looked broken.
 *
 * So the field goes in WITH a cached result: the headings we already read,
 * styled TOC1/TOC2/TOC3, sitting between `separate` and `end` exactly where
 * Word puts its own. The contents page is readable immediately, and the moment
 * anyone updates fields it is replaced by the real thing with page numbers.
 *
 * Page numbers are the one part we cannot supply — Office.js does not expose
 * pagination — so they are absent rather than wrong.
 */
const xmlEscape = (v) => String(v).replace(/[&<>"']/g, (c) => (
  { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' }[c]
));

const OOXML_HEAD = '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">'
  + '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">'
  + '<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
  + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/word/document.xml"/>'
  + '</Relationships></pkg:xmlData></pkg:part>'
  + '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">'
  + '<pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>';
const OOXML_TAIL = '</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>';

/** One paragraph, optionally styled. */
const para = (text, style) =>
  '<w:p>' + (style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : '')
  + (text ? `<w:r><w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r>` : '')
  + '</w:p>';

function tocOoxml(entries, heading) {
  const body = [];
  if (heading) body.push(para(heading, 'TOCHeading'));

  // The field opens inside the FIRST entry's paragraph and closes in a paragraph
  // of its own — the shape Word itself writes, and the one it re-reads cleanly.
  const first = entries[0];
  body.push(
    `<w:p><w:pPr><w:pStyle w:val="TOC${first.level}"/></w:pPr>`
    + '<w:r><w:fldChar w:fldCharType="begin" w:dirty="1"/></w:r>'
    + '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
    + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    + `<w:r><w:t xml:space="preserve">${xmlEscape(first.text)}</w:t></w:r></w:p>`,
  );
  for (const e of entries.slice(1)) body.push(para(e.text, `TOC${e.level}`));
  body.push('<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>');

  return OOXML_HEAD + body.join('') + OOXML_TAIL;
}

async function insertToc(args) {
  return Word.run(async (ctx) => {
    const items = await loadParagraphs(ctx);
    const entries = items
      .map((p) => ({ level: Math.min(3, headingLevel(p)), text: String(p.text || '').trim() }))
      .filter((e) => e.level > 0 && e.text);

    // An empty contents page is worse than no contents page: it looks like the
    // tool failed. Say what is actually missing instead.
    if (!entries.length) {
      throw new Error(
        '這份文件沒有任何段落套用「標題」樣式，所以目錄會是空的。'
        + '請先用 word_format_range 把各章節標題設成 Heading1／Heading2（style 參數），再插入目錄。',
      );
    }

    const heading = args.heading === false ? '' : String(args.heading || '目錄');
    const ooxml = tocOoxml(entries, heading);

    // Deliberately NOT tracked, even though every other write is.
    //
    // A TOC is a field: what a reviewer would see in the revisions pane is an
    // insertion of machinery, not of words, and there is nothing meaningful to
    // accept or reject. It is also the one write that is purely derived — it
    // says nothing the headings did not already say. The snapshot still covers
    // it, so 復原上一步 takes it straight back out.
    const result = await withTracking(ctx, { track: false }, async () => {
      if (args.paragraph && items.length) {
        const { from } = resolveRange(items, args);
        items[from - 1].insertOoxml(ooxml, String(args.at) === 'before' ? 'Before' : 'After');
      } else {
        ctx.document.body.insertOoxml(ooxml, 'Start');
      }
      await ctx.sync();
      return null;
    });

    return `已插入目錄，收錄 ${entries.length} 個標題（${entries.filter(e => e.level === 1).length} 個第一層）。`
      + '項目現在就看得到了；頁碼要等使用者按一次「參考資料 → 更新目錄」才會出現，'
      + '因為 Office.js 讀不到分頁資訊。之後標題有增減，同樣按更新就會同步。'
      + trackingNote(result);
  });
}

async function replaceAll(args) {
  const find = String(args.find ?? '');
  const replace = String(args.replace ?? '');
  if (!find) throw new Error('find 是必要的');
  return Word.run(async (ctx) => {
    const results = ctx.document.body.search(find, {
      matchCase: !!args.match_case,
      matchWholeWord: !!args.whole_word,
    });
    results.load('items/text');
    await ctx.sync();
    const n = results.items.length;
    if (!n) return `找不到「${find}」，沒有改動任何地方。`;

    const result = await withTracking(ctx, args, async () => {
      for (const r of results.items) r.insertText(replace, 'Replace');
      await ctx.sync();
      return null;
    });
    return `已把 ${n} 處「${find}」換成「${replace}」。` + trackingNote(result);
  });
}

async function trackedChanges(args) {
  if (!hasApi('1.4')) throw new Error('這個 Word 版本不支援從增益集操作追蹤修訂（需要 WordApi 1.4）。');
  const op = String(args.op || 'status');
  return Word.run(async (ctx) => {
    const doc = ctx.document;

    if (op === 'set_mode') {
      const mode = TRACK_MODES[String(args.mode || 'all').toLowerCase()] || TRACK_MODES.all;
      doc.changeTrackingMode = mode;
      await ctx.sync();
      return `已把追蹤修訂設成 ${mode}。`;
    }

    if (op === 'status') {
      doc.load('changeTrackingMode');
      await ctx.sync();
      return `追蹤修訂目前是：${doc.changeTrackingMode}`;
    }

    if (!hasApi('1.6')) throw new Error('列出或接受個別修訂需要 WordApi 1.6，這個版本只能改模式。');
    const changes = doc.body.getTrackedChanges();
    changes.load('items/text,items/type,items/author');
    await ctx.sync();

    if (op === 'list') {
      if (!changes.items.length) return '這份文件目前沒有未處理的修訂。';
      const lines = changes.items.slice(0, MAX_SEARCH_RESULTS)
        .map((c, i) => `${i + 1}\t${c.type}\t${c.author || ''}\t${snip(c.text, 80)}`);
      return `未處理的修訂共 ${changes.items.length} 筆：\n編號\t類型\t作者\t內容\n${lines.join('\n')}`;
    }
    if (op === 'accept_all') { changes.acceptAll(); await ctx.sync(); return '已接受全部修訂。'; }
    if (op === 'reject_all') { changes.rejectAll(); await ctx.sync(); return '已拒絕全部修訂，文件回到修訂前的樣子。'; }
    throw new Error(`未知的 op：${op}`);
  });
}

async function comment(args) {
  if (!hasApi('1.4')) throw new Error('這個 Word 版本不支援註解 API（需要 WordApi 1.4）。');
  const op = String(args.op || 'add');
  return Word.run(async (ctx) => {
    if (op === 'list') {
      const comments = ctx.document.body.getComments();
      comments.load('items/content,items/authorName,items/resolved');
      await ctx.sync();
      if (!comments.items.length) return '這份文件目前沒有註解。';
      return `註解共 ${comments.items.length} 則：\n`
        + comments.items.map((c, i) => `${i + 1}\t${c.authorName || ''}\t${c.resolved ? '已解決' : '未解決'}\t${snip(c.content, 100)}`).join('\n');
    }

    const text = String(args.text || '').trim();
    if (!text) throw new Error('text 是必要的');
    const items = await loadParagraphs(ctx);
    if (!items.length) throw new Error('這份文件是空的。');
    const { from } = resolveRange(items, args);
    // A comment is the right shape for 「這一條有風險」: it says something about
    // the text without changing it, which is exactly what a review note is.
    items[from - 1].getRange().insertComment(text);
    await ctx.sync();
    return `已在段落 ${from} 加上註解。`;
  });
}


// ── Layout ────────────────────────────────────────────────────────────────────

/**
 * Header and footer, with a page number that is visible straight away.
 *
 * Same trap as the table of contents: a page number is a FIELD, and a field
 * Word has not evaluated shows nothing. So it goes in with a cached result — a
 * literal 1 sitting where the field's answer belongs — and every page after the
 * first is correct as soon as Word repaginates, which it does on its own.
 */
function pageNumberOoxml(style) {
  const field = (instr, cached) =>
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
    + `<w:r><w:instrText xml:space="preserve"> ${instr} </w:instrText></w:r>`
    + '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
    + `<w:r><w:t>${cached}</w:t></w:r>`
    + '<w:r><w:fldChar w:fldCharType="end"/></w:r>';

  const runs = style === 'x_of_y'
    ? '<w:r><w:t xml:space="preserve">第 </w:t></w:r>' + field('PAGE', '1')
      + '<w:r><w:t xml:space="preserve"> 頁，共 </w:t></w:r>' + field('NUMPAGES', '1')
      + '<w:r><w:t xml:space="preserve"> 頁</w:t></w:r>'
    : field('PAGE', '1');

  return OOXML_HEAD
    + `<w:p><w:pPr><w:jc w:val="center"/></w:pPr>${runs}</w:p>`
    + OOXML_TAIL;
}

const HF_ALIGN = { left: 'Left', center: 'Centered', centered: 'Centered', right: 'Right' };

async function headerFooter(args) {
  const where = String(args.where || 'footer').toLowerCase();
  const isHeader = where === 'header';
  return Word.run(async (ctx) => {
    const sections = ctx.document.sections;
    sections.load('items');
    await ctx.sync();
    if (!sections.items.length) throw new Error('讀不到文件的節（section），無法設定頁首頁尾。');

    let touched = 0;
    for (const section of sections.items) {
      const body = isHeader ? section.getHeader('Primary') : section.getFooter('Primary');
      if (args.clear) { body.clear(); touched++; continue; }
      body.clear();
      if (args.text) {
        const p = body.insertParagraph(String(args.text), 'Start');
        p.alignment = HF_ALIGN[String(args.alignment || 'center').toLowerCase()] || 'Centered';
      }
      if (args.page_number) {
        body.insertOoxml(pageNumberOoxml(String(args.page_number)), 'End');
      }
      touched++;
    }
    await ctx.sync();

    const what = args.clear ? '已清空' : '已設定';
    return `${what}${isHeader ? '頁首' : '頁尾'}（${touched} 個節）。`
      + (args.page_number ? '頁碼是欄位，第一頁現在就看得到，其餘頁 Word 重新分頁後會自己算。' : '');
  });
}

const BREAKS = { page: 'Page', line: 'Line', section: 'SectionNext', section_continuous: 'SectionContinuous' };

async function insertBreak(args) {
  const kind = BREAKS[String(args.type || 'page').toLowerCase()] || 'Page';
  return Word.run(async (ctx) => {
    const items = await loadParagraphs(ctx);
    const result = await withTracking(ctx, args, async () => {
      if (args.paragraph && items.length) {
        const { from } = resolveRange(items, args);
        items[from - 1].insertBreak(kind, String(args.at) === 'before' ? 'Before' : 'After');
      } else {
        ctx.document.body.insertBreak(kind, 'End');
      }
      await ctx.sync();
      return null;
    });
    const label = kind === 'Page' ? '分頁符號' : kind === 'Line' ? '換行符號' : '分節符號';
    return `已插入${label}。` + trackingNote(result);
  });
}

// ── Charts ────────────────────────────────────────────────────────────────────

/**
 * Charts, drawn in the pane and inserted as a picture.
 *
 * Word's JavaScript API has no chart object — Excel's `charts.add()` has no
 * counterpart here, and the only way to put a NATIVE chart into a document is
 * to build an OOXML part containing an embedded workbook, which is a
 * disproportionate amount of machinery to maintain for something that would
 * still not be editable from this add-in.
 *
 * So the pane draws it. It is a web page with a canvas; a bar chart is fifty
 * lines of arithmetic. What goes into the document is a PNG.
 *
 * The honest cost: the reader cannot click the chart and change the numbers,
 * and it does not re-render if the underlying figures change. Both are stated
 * in the tool description so the model does not promise otherwise. What it buys
 * is that 「幫我把這段數字畫成圖」 works at all, and looks like something you
 * would put in front of a customer.
 */
const CHART_COLORS = ['#2563eb', '#16a34a', '#ea580c', '#9333ea', '#0891b2', '#ca8a04', '#dc2626', '#4b5563'];
/** Rendered at 2x so it stays crisp when Word scales it to the text column. */
const CHART_SCALE = 2;

export function drawChart(spec) {
  const W = 900;
  const H = spec.type === 'pie' ? 460 : 500;
  const canvas = document.createElement('canvas');
  canvas.width = W * CHART_SCALE;
  canvas.height = H * CHART_SCALE;
  const c = canvas.getContext('2d');
  c.scale(CHART_SCALE, CHART_SCALE);
  c.fillStyle = '#ffffff';
  c.fillRect(0, 0, W, H);
  c.textBaseline = 'middle';
  const font = (size, weight) => `${weight || ''} ${size}px "Microsoft JhengHei UI", "Microsoft JhengHei", sans-serif`;

  const cats = spec.categories;
  const series = spec.series;
  let top = 24;

  if (spec.title) {
    c.fillStyle = '#111827';
    c.font = font(22, '600');
    c.textAlign = 'center';
    c.fillText(spec.title, W / 2, top + 12);
    top += 44;
  }

  // Legend, only when there is more than one thing to tell apart.
  if (series.length > 1) {
    c.font = font(14);
    c.textAlign = 'left';
    const widths = series.map(s => c.measureText(s.name).width + 30);
    let x = (W - widths.reduce((a, b) => a + b, 0)) / 2;
    series.forEach((sr, i) => {
      c.fillStyle = CHART_COLORS[i % CHART_COLORS.length];
      c.fillRect(x, top + 4, 12, 12);
      c.fillStyle = '#374151';
      c.fillText(sr.name, x + 18, top + 10);
      x += widths[i];
    });
    top += 30;
  }

  if (spec.type === 'pie') {
    const values = series[0].values.map(v => Math.max(0, Number(v) || 0));
    const total = values.reduce((a, b) => a + b, 0) || 1;
    const cx = W / 2 - 90;
    const cy = top + (H - top) / 2 - 10;
    const r = Math.min(H - top - 60, 300) / 2;
    let angle = -Math.PI / 2;
    values.forEach((v, i) => {
      const slice = (v / total) * Math.PI * 2;
      c.beginPath();
      c.moveTo(cx, cy);
      c.arc(cx, cy, r, angle, angle + slice);
      c.closePath();
      c.fillStyle = CHART_COLORS[i % CHART_COLORS.length];
      c.fill();
      c.strokeStyle = '#ffffff';
      c.lineWidth = 2;
      c.stroke();
      angle += slice;
    });
    // Labels beside the pie rather than on it: Chinese category names do not fit
    // inside a slice, and a leader-line salad is worse than a list.
    c.textAlign = 'left';
    c.font = font(14);
    const lx = cx + r + 40;
    let ly = cy - (values.length - 1) * 13;
    values.forEach((v, i) => {
      c.fillStyle = CHART_COLORS[i % CHART_COLORS.length];
      c.fillRect(lx, ly - 6, 12, 12);
      c.fillStyle = '#374151';
      c.fillText(`${cats[i] || ''}　${((v / total) * 100).toFixed(1)}%`, lx + 18, ly);
      ly += 26;
    });
    return canvas;
  }

  // ── Column / bar / line share one plot area ──
  const horizontal = spec.type === 'bar';
  const all = series.flatMap(sr => sr.values.map(v => Number(v) || 0));
  const rawMax = Math.max(...all, 0);
  const rawMin = Math.min(...all, 0);
  // Round the axis out to something a person would have chosen.
  const span = rawMax - rawMin || 1;
  const step = Math.pow(10, Math.floor(Math.log10(span / 4))) * [1, 2, 5, 10].find(m => span / 4 <= Math.pow(10, Math.floor(Math.log10(span / 4))) * m);
  const max = Math.ceil(rawMax / step) * step;
  const min = Math.min(0, Math.floor(rawMin / step) * step);

  const fmt = (v) => (Math.abs(v) >= 10000 ? (v / 10000).toFixed(1) + '萬' : String(Math.round(v * 100) / 100));
  c.font = font(13);
  const labelW = horizontal
    ? Math.min(160, Math.max(...cats.map(t => c.measureText(String(t)).width)) + 16)
    : Math.max(...[min, max].map(v => c.measureText(fmt(v)).width)) + 16;

  const left = 20 + labelW;
  const right = W - 28;
  const bottom = H - 46;
  const plotH = bottom - top;
  const plotW = right - left;

  // Gridlines, and the numbers that make them mean something.
  const ticks = 4;
  c.strokeStyle = '#e5e7eb';
  c.lineWidth = 1;
  c.fillStyle = '#6b7280';
  for (let i = 0; i <= ticks; i++) {
    const v = min + ((max - min) * i) / ticks;
    if (horizontal) {
      const x = left + (plotW * i) / ticks;
      c.beginPath(); c.moveTo(x, top); c.lineTo(x, bottom); c.stroke();
      c.textAlign = 'center';
      c.fillText(fmt(v), x, bottom + 16);
    } else {
      const y = bottom - (plotH * i) / ticks;
      c.beginPath(); c.moveTo(left, y); c.lineTo(right, y); c.stroke();
      c.textAlign = 'right';
      c.fillText(fmt(v), left - 10, y);
    }
  }

  const at = (v) => (v - min) / (max - min || 1);
  const slot = (horizontal ? plotH : plotW) / cats.length;

  if (spec.type === 'line') {
    series.forEach((sr, si) => {
      c.strokeStyle = CHART_COLORS[si % CHART_COLORS.length];
      c.lineWidth = 2.5;
      c.beginPath();
      sr.values.forEach((raw, i) => {
        const x = left + slot * (i + 0.5);
        const y = bottom - plotH * at(Number(raw) || 0);
        i === 0 ? c.moveTo(x, y) : c.lineTo(x, y);
      });
      c.stroke();
      c.fillStyle = CHART_COLORS[si % CHART_COLORS.length];
      sr.values.forEach((raw, i) => {
        const x = left + slot * (i + 0.5);
        const y = bottom - plotH * at(Number(raw) || 0);
        c.beginPath(); c.arc(x, y, 4, 0, Math.PI * 2); c.fill();
      });
    });
  } else {
    const groupPad = slot * 0.22;
    const barSize = (slot - groupPad * 2) / series.length;
    series.forEach((sr, si) => {
      c.fillStyle = CHART_COLORS[si % CHART_COLORS.length];
      sr.values.forEach((raw, i) => {
        const v = Number(raw) || 0;
        if (horizontal) {
          const y = top + slot * i + groupPad + barSize * si;
          const x0 = left + plotW * at(Math.min(0, v));
          const x1 = left + plotW * at(v);
          c.fillRect(x0, y, Math.max(1, x1 - x0), barSize * 0.86);
        } else {
          const x = left + slot * i + groupPad + barSize * si;
          const y0 = bottom - plotH * at(Math.min(0, v));
          const y1 = bottom - plotH * at(v);
          c.fillRect(x, Math.min(y0, y1), barSize * 0.86, Math.max(1, Math.abs(y1 - y0)));
        }
      });
    });
  }

  // Category labels.
  c.fillStyle = '#374151';
  c.font = font(13);
  cats.forEach((label, i) => {
    if (horizontal) {
      c.textAlign = 'right';
      c.fillText(String(label), left - 10, top + slot * (i + 0.5));
    } else {
      c.textAlign = 'center';
      c.fillText(String(label), left + slot * (i + 0.5), bottom + 18);
    }
  });

  c.strokeStyle = '#9ca3af';
  c.beginPath();
  c.moveTo(left, top); c.lineTo(left, bottom); c.lineTo(right, bottom);
  c.stroke();
  return canvas;
}

async function insertChart(args) {
  const type = String(args.type || 'column').toLowerCase();
  const cats = Array.isArray(args.categories) ? args.categories.map(String) : [];
  let series = Array.isArray(args.series) ? args.series : [];
  // A single unnamed row of numbers is the common case; accept it without
  // making the model wrap it in the full shape.
  if (!series.length && Array.isArray(args.values)) series = [{ name: args.title || '數列 1', values: args.values }];
  series = series
    .filter(sr => sr && Array.isArray(sr.values))
    .map((sr, i) => ({ name: String(sr.name || `數列 ${i + 1}`), values: sr.values }));

  if (!cats.length || !series.length) {
    throw new Error('需要 categories（分類名稱）和 series（數列）。例如 '
      + 'categories=["Q1","Q2"], series=[{name:"營收", values:[120,140]}]。');
  }
  if (series.some(sr => sr.values.length !== cats.length)) {
    throw new Error(`每個數列的數字個數必須跟 categories 一樣多（目前 categories 有 ${cats.length} 個）。`);
  }
  if (type === 'pie' && series.length > 1) {
    throw new Error('圓餅圖只能有一個數列。要比較多個數列請用 column 或 bar。');
  }

  const canvas = drawChart({ type, title: String(args.title || ''), categories: cats, series });
  const base64 = canvas.toDataURL('image/png').split(',')[1];

  return Word.run(async (ctx) => {
    const items = await loadParagraphs(ctx);
    const result = await withTracking(ctx, args, async () => {
      let anchor;
      if (args.paragraph && items.length) {
        const { from } = resolveRange(items, args);
        anchor = items[from - 1].insertParagraph('', String(args.at) === 'before' ? 'Before' : 'After');
      } else {
        anchor = ctx.document.body.insertParagraph('', 'End');
      }
      anchor.alignment = 'Centered';
      anchor.insertInlinePictureFromBase64(base64, 'Replace');
      if (args.caption) {
        const cap = anchor.insertParagraph(String(args.caption), 'After');
        cap.alignment = 'Centered';
        // The built-in Caption style is what a real document uses, and it is what
        // 「插入圖表標號」 later picks up.
        cap.styleBuiltIn = 'Caption';
      }
      await ctx.sync();
      return null;
    });
    return `已插入${type === 'pie' ? '圓餅' : type === 'line' ? '折線' : type === 'bar' ? '橫條' : '直條'}圖`
      + `（${cats.length} 個分類、${series.length} 個數列）。`
      + '圖是以圖片插入的——Word 的增益集 API 沒有原生圖表物件，所以它不能在 Word 裡點開改數字；'
      + '數字有變要重新產生一張。' + trackingNote(result);
  });
}

// ── Undo ──────────────────────────────────────────────────────────────────────

/**
 * Cap on the before-image, in characters of OOXML.
 *
 * OOXML is verbose — a page of prose is tens of KB — so this is generous by
 * design. Past it the snapshot is skipped and the pane says the change cannot be
 * taken back, which is better than spending ten seconds building one.
 */
const MAX_SNAPSHOT_CHARS = 8000000;

/**
 * Photograph the document before a write.
 *
 * `body.getOoxml()` is the whole body — text, styles, tables, numbering — so
 * restoring it is exact in a way Excel's cell snapshot never manages. What it
 * does NOT cover is headers, footers and comments, which live outside the body;
 * a tool that only touches those is also a tool that cannot lose a paragraph, so
 * the trade is worth naming and keeping.
 */
async function captureDocument() {
  return Word.run(async (ctx) => {
    const ooxml = ctx.document.body.getOoxml();
    await ctx.sync();
    const xml = String(ooxml.value || '');
    if (!xml || xml.length > MAX_SNAPSHOT_CHARS) return null;
    return { kind: 'document', ooxml: xml };
  });
}

export async function restoreSnapshot(snapshot) {
  return Word.run(async (ctx) => {
    // Restoring must not itself become a revision: undo is the person taking
    // their document back, not another edit for them to review. So tracking goes
    // off for the duration and then back to whatever it was.
    const doc = ctx.document;
    let previous = null;
    if (hasApi('1.4')) {
      doc.load('changeTrackingMode');
      await ctx.sync();
      previous = String(doc.changeTrackingMode || TRACK_MODES.off);
      if (previous !== TRACK_MODES.off) { doc.changeTrackingMode = TRACK_MODES.off; await ctx.sync(); }
    }
    try {
      doc.body.clear();
      doc.body.insertOoxml(snapshot.ooxml, 'Replace');
      await ctx.sync();
    } finally {
      if (previous && previous !== TRACK_MODES.off) { doc.changeTrackingMode = previous; await ctx.sync(); }
    }
    return { kind: 'document', label: t('這份文件的內容'), notes: [] };
  });
}

// ── Citations, naming, location ───────────────────────────────────────────────

/** Follow [[段落 42]] — select it so the person lands on the sentence cited. */
export async function goTo(from, to) {
  const start = Number(from);
  if (!Number.isFinite(start)) return false;
  const end = Number(to) || start;
  try {
    return await Word.run(async (ctx) => {
      const paras = ctx.document.body.paragraphs;
      paras.load('items');
      await ctx.sync();
      const items = paras.items;
      if (!items.length || start > items.length) return false;
      const first = items[start - 1];
      const last = items[Math.min(end, items.length) - 1] || first;
      // Selecting the span rather than the first paragraph: a citation that
      // covers three paragraphs should highlight three.
      first.getRange('Start').expandTo(last.getRange('End')).select();
      await ctx.sync();
      return true;
    });
  } catch {
    return false;
  }
}

export async function getDocumentName() {
  // Read the URL directly, not inside a Word.run: it is a synchronous property
  // of the common Office context, and making the file name depend on a batch
  // that can fail is how the bar ended up saying 「（未命名文件）」 for a
  // document that has a perfectly good name.
  const url = (Office.context.document && Office.context.document.url) || '';
  if (url) {
    // SharePoint and OneDrive hand back a web URL with a query string on it.
    const clean = url.split(/[?#]/)[0];
    const file = decodeURIComponent(clean.split(/[\\/]/).pop() || '');
    if (file) return file;
  }
  // A document that has never been saved genuinely has no file name — Word calls
  // it 文件1 and that string is not available to an add-in. The title property
  // is the last thing worth trying before giving up.
  try {
    return await Word.run(async (ctx) => {
      const props = ctx.document.properties;
      props.load('title');
      await ctx.sync();
      return String(props.title || '');
    });
  } catch {
    return '';
  }
}

const DOCUMENT_ID_PROP = 'JVisionDocumentId';

/**
 * A stable identifier for this document, for keying conversation history.
 *
 * Same reasoning as Excel's: everything Office exposes changes when the file is
 * renamed or moved, and that is the one thing this has to survive. Minting the
 * id writes a custom property, which marks the file as modified — so callers ask
 * only when the person is already doing something to it, never on pane open.
 */
export async function getDocumentKey() {
  if (!hasApi('1.3')) return null;
  try {
    return await Word.run(async (ctx) => {
      const custom = ctx.document.properties.customProperties;
      const existing = custom.getItemOrNullObject(DOCUMENT_ID_PROP);
      existing.load('isNullObject,value');
      await ctx.sync();
      if (!existing.isNullObject && existing.value) return String(existing.value);

      const id = (self.crypto && self.crypto.randomUUID)
        ? self.crypto.randomUUID()
        : String(Date.now()) + '-' + Math.random().toString(16).slice(2);
      custom.add(DOCUMENT_ID_PROP, id);
      await ctx.sync();
      return id;
    });
  } catch {
    return null;
  }
}

/**
 * Where the cursor is, and what is selected — one read, many listeners.
 *
 * ── The bug this replaces ──
 * The first version called `ctx.document.onSelectionChanged.add(...)`, which is
 * an EXCEL api. Word.Document has no such event, so the call threw, the throw
 * landed in a catch that treated it as "an older build has no event", and the
 * selection chip silently never appeared. Word raises the COMMON Office.js
 * event instead — the one that works in every host.
 *
 * ── Why one handler and not two ──
 * Both the chip and the location bar need the same answer, and computing it
 * loads every paragraph in the document. Registering the two watchers
 * separately meant doing that twice per click. Now the handler runs once and
 * fans out, debounced, because dragging a selection fires this continuously.
 */
const selectionListeners = [];
const locationListeners = [];
let handlerBound = false;
let debounce = null;

/** Long enough to swallow a drag, short enough that a click feels immediate. */
const SELECTION_DEBOUNCE_MS = 200;

function bindSelectionHandler() {
  if (handlerBound) return;
  handlerBound = true;
  try {
    Office.context.document.addHandlerAsync(
      Office.EventType.DocumentSelectionChanged,
      () => {
        clearTimeout(debounce);
        debounce = setTimeout(() => { void fanOutSelection(); }, SELECTION_DEBOUNCE_MS);
      },
    );
  } catch {
    // No event on this build: the bar and the chip stay on whatever the opening
    // read found. Degraded, not broken.
    handlerBound = false;
  }
}

async function fanOutSelection() {
  let ref = null;
  let where = '';
  try {
    await Word.run(async (ctx) => {
      const info = await getSelectionInfo(ctx);
      ref = selectionRefFrom(info);
      where = locationFrom(info);
    });
  } catch { /* a failed read should not kill the handler for the next one */ }
  for (const fn of selectionListeners) { try { fn(ref); } catch { /* ignore */ } }
  for (const fn of locationListeners) { try { fn(where); } catch { /* ignore */ } }
}

/**
 * The bar under the file name.
 *
 * Shows the paragraph NUMBER rather than the heading, because that number is
 * the address: it is what word_read_range takes, and what [[段落 42]] means. Seeing
 * it move as the cursor moves is what teaches someone that the two are the same
 * thing. The heading rides along when there is room for it.
 */
function locationFrom(info) {
  if (!info.from) return '';
  const at = info.to > info.from ? `段落 ${info.from}-${info.to}` : `段落 ${info.from}`;
  return info.heading ? `${at} · ${info.heading}` : at;
}

export async function getLocation() {
  try {
    return await Word.run(async (ctx) => locationFrom(await getSelectionInfo(ctx)));
  } catch {
    return '';
  }
}

export async function watchLocation(onChange) {
  locationListeners.push(onChange);
  bindSelectionHandler();
}

/** Cap on what rides along with the message. A whole chapter is not a chip. */
const MAX_SELECTION_PREVIEW_CHARS = 4000;

function selectionRefFrom(info) {
  const text = info.text.trim();
  if (!text) return null;

  const where = info.from
    ? `段落 ${info.from}${info.to > info.from ? `-${info.to}` : ''}`
    : `${text.length} 字`;
  const head = `使用者在側邊欄提問時選取的範圍：${where}`
    + `${info.heading ? `（位於「${info.heading}」底下）` : ''}`;

  return {
    // Already the finished chip text — the shell must not have to know that
    // Excel joins with 「!」 and Word does not.
    label: info.heading ? `${where} · ${info.heading}` : where,
    context: text.length > MAX_SELECTION_PREVIEW_CHARS
      ? `${head}
（選取範圍較長，只附上開頭；完整內容請用 word_read_range 讀）
${text.slice(0, MAX_SELECTION_PREVIEW_CHARS)}…`
      : `${head}
內容：
${text}`,
  };
}

export async function getSelectionRef() {
  return Word.run(async (ctx) => selectionRefFrom(await getSelectionInfo(ctx)));
}

export async function watchSelection(onChange) {
  selectionListeners.push(onChange);
  bindSelectionHandler();
}

// ── run_script ────────────────────────────────────────────────────────────────

const MAX_SCRIPT_RESULT_CHARS = 4000;
const MAX_SCRIPT_LOG_LINES = 100;
const SCRIPT_TIMEOUT_READ_MS = 40000;
const SCRIPT_TIMEOUT_WRITE_MS = 100000;

/** Scratch space shared across script calls, for the pane's lifetime. */
const scriptStore = Object.create(null);

function describeStore() {
  const keys = Object.keys(scriptStore);
  if (!keys.length) return '';
  return `store 目前有：${keys.map((k) => {
    const v = scriptStore[k];
    if (Array.isArray(v)) return `${k}（${v.length} 筆陣列）`;
    if (v && typeof v === 'object') return `${k}（物件，${Object.keys(v).length} 個鍵）`;
    return `${k}（${typeof v}）`;
  }).join('、')}`;
}

function stringifyResult(v, cap, tailNote) {
  let s;
  if (typeof v === 'string') s = v;
  else {
    try { s = JSON.stringify(v, null, 2); } catch { s = String(v); }
    if (s === undefined) s = String(v);
  }
  return s.length > cap ? s.slice(0, cap) + tailNote : s;
}

/**
 * Run model-written JavaScript against the live document.
 *
 * Worth as much here as in Excel, for a different reason. There it was
 * arithmetic; here it is the sweep — 「把所有全形括號換成半形」,
 * 「找出每一個沒有加單位的數字」 — where the alternative is pulling the whole
 * document through the model's context and pushing it back paragraph by
 * paragraph.
 *
 * `new Function` rather than eval: the body sees the bindings passed to it and
 * not this module's scope. Scope hygiene, NOT a sandbox — the guard is what
 * stands between an unapproved call and the document, and the person clicking
 * 允許 is what stands behind that.
 */
async function runScript(args, allowWrites) {
  const code = String(args.code || '').trim();
  if (!code) throw new Error('code 是必要的');

  const logs = [];
  const log = (...parts) => {
    if (logs.length >= MAX_SCRIPT_LOG_LINES) return;
    logs.push(parts.map(p => (typeof p === 'string' ? p : stringifyResult(p, 500, '…'))).join(' '));
  };

  let body;
  try {
    body = new Function(
      'ctx', 'context', 'document', 'body', 'Word', 'log', 'store',
      '"use strict"; return (async () => {\n' + code + '\n})();',
    );
  } catch (e) {
    throw new Error(`程式碼語法錯誤：${(e && e.message) || String(e)}`);
  }

  const limit = String(args.mode || '') === 'write' ? SCRIPT_TIMEOUT_WRITE_MS : SCRIPT_TIMEOUT_READ_MS;
  const guard = makeScriptGuard(allowWrites, 'word_run_script');

  const run = Word.run(async (ctx) => {
    const g = guard.wrap(ctx);
    const value = await body(g, g, g.document, g.document.body, guard.wrapNamespace(Word), log, scriptStore);
    await ctx.sync();
    return value;
  });

  const value = await Promise.race([
    run,
    new Promise((_, reject) => setTimeout(
      () => reject(new Error(
        `程式執行超過 ${Math.round(limit / 1000)} 秒。`
        + '常見原因是在迴圈裡呼叫 ctx.sync()——改成一次 load 整份 paragraphs、sync 一次，剩下用純 JS 算。',
      )),
      limit,
    )),
  ]);

  const parts = [];
  if (logs.length) parts.push('log:\n' + logs.join('\n'));
  const stored = describeStore();
  if (stored) parts.push(stored);
  parts.push(
    value === undefined
      ? '（程式沒有 return 任何東西。要拿到結果請用 return。）'
      : stringifyResult(value, MAX_SCRIPT_RESULT_CHARS,
        '\n…（結果超過 4 KB 已截斷。要留給下一支 script 用的資料請放進 store，'
        + '例如 store.rows = rows —— 它跨呼叫保留，而且不會經過你的上下文。'
        + '若要寫進文件，就在這一支 script 裡直接寫完，不要先回傳。）'),
  );
  return parts.join('\n\n');
}

// ── Dispatch ──────────────────────────────────────────────────────────────────

/**
 * Run one tool call. Always resolves — never throws.
 *
 * A failure has to reach the model as a readable tool result so it can correct
 * itself; an exception here would strand the agent waiting for a result that
 * never comes.
 */
export async function executeTool(tool, args, opts = {}) {
  /** Set before anything that writes, so a half-finished change is undoable. */
  let snapshot;
  const a = args || {};
  try {
    switch (tool) {
      case 'word_get_overview':  return { ok: true, content: await getOverview(a) };
      case 'word_get_selection': return { ok: true, content: await getSelection() };
      case 'word_read_range':    return { ok: true, content: await readRange(a) };
      case 'word_search':        return { ok: true, content: await search(a) };

      case 'word_write_range':
      case 'word_insert_text':
      case 'word_delete_range':
      case 'word_format_range':
      case 'word_insert_table':
      case 'word_insert_toc':
      case 'word_manage_list':
      case 'word_header_footer':
      case 'word_insert_break':
      case 'word_insert_chart':
      case 'word_replace_all': {
        // Every one of these changes the document, so every one is photographed
        // first — including the ones that did not need a click. Tracked changes
        // cover a rewrite, but they do not cover a style sweep, and 「復原上一步」
        // has to mean the same thing whichever tool ran.
        snapshot = await captureDocument();
        const run = {
          word_write_range: writeRange,
          word_insert_text: insertText,
          word_delete_range: deleteRange,
          word_format_range: formatRange,
          word_insert_table: insertTable,
          word_insert_toc: insertToc,
          word_manage_list: manageList,
          word_header_footer: headerFooter,
          word_insert_break: insertBreak,
          word_insert_chart: insertChart,
          word_replace_all: replaceAll,
        }[tool];
        return { ok: true, content: await run(a), snapshot };
      }

      case 'word_tracked_changes': {
        // accept/reject rewrite the document as surely as a write does.
        if (String(a.op || '').startsWith('accept') || String(a.op || '').startsWith('reject')) {
          snapshot = await captureDocument();
        }
        return { ok: true, content: await trackedChanges(a), snapshot };
      }
      case 'word_comment':       return { ok: true, content: await comment(a) };
      case 'word_run_script': {
        if (opts.writeApproved) snapshot = await captureDocument();
        const text = await runScript(a, !!opts.writeApproved);
        return { ok: true, content: text, snapshot };
      }
      default:
        return { ok: false, error: `未知的工具：${tool}` };
    }
  } catch (e) {
    const code = e && e.code ? `[${e.code}] ` : '';
    let msg = `${code}${(e && e.message) || String(e)}`;
    if (tool === 'word_run_script' && e instanceof ReferenceError) {
      msg += '。這段程式可以用的名稱只有：context（= ctx）、document、body、Word、log、store。'
        + '沒有 Office Scripts 的 ExcelScript，也沒有 require / import / fetch。';
    }
    return { ok: false, error: msg, snapshot };
  }
}

/**
 * Everything the shell used to assume — see hosts/excel.js for the long version
 * of why this lives in the host module and not in app.js.
 */
export const meta = {
  id: 'word',
  officeHost: 'Word',
  api: 'word',
  mcpPrefix: 'mcp__word__',
  overview: { tool: 'word_get_overview', args: { preview_chars: OVERVIEW_PREVIEW_CHARS } },
  askUserTool: 'word_ask_user',
  inspectTools: ['word_write_range'],
  scriptTool: 'word_run_script',

  /**
   * [[段落 42]] or [[段落 12-15]] → a link that selects it.
   *
   * Word has no A1, so the citation is the paragraph number the model was given
   * when it read the text. The second group is optional; the shell passes an
   * empty string through when a citation names one paragraph.
   */
  citation: {
    pattern: /\[\[段落\s*(\d{1,5})(?:\s*[-–~]\s*(\d{1,5}))?\]\]/g,
    label: (a, b) => (b ? `段落 ${a}-${b}` : `段落 ${a}`),
  },

  toolLabels: () => ({
    word_get_overview: t('正在讀取文件結構…'),
    word_get_selection: t('正在讀取選取範圍…'),
    word_read_range: t('正在閱讀內文…'),
    word_search: t('正在搜尋…'),
    word_write_range: t('等待你確認改寫…'),
    word_insert_text: t('正在插入內容…'),
    word_delete_range: t('等待你確認刪除…'),
    word_format_range: t('正在套用格式…'),
    word_insert_table: t('正在插入表格…'),
    word_insert_toc: t('正在建立目錄…'),
    word_manage_list: t('正在整理清單…'),
    word_header_footer: t('正在設定頁首頁尾…'),
    word_insert_break: t('正在調整分頁…'),
    word_insert_chart: t('正在繪製圖表…'),
    word_replace_all: t('等待你確認取代…'),
    word_tracked_changes: t('正在處理追蹤修訂…'),
    word_comment: t('正在加註解…'),
    word_run_script: t('正在執行程式…'),
  }),

  preparingLabels: () => ({
    word_run_script: t('正在寫程式…'),
    word_write_range: t('正在草擬改寫內容…'),
    word_insert_text: t('正在草擬要插入的內容…'),
    word_insert_table: t('正在整理表格…'),
    word_insert_chart: t('正在設計圖表…'),
    word_format_range: t('正在決定格式…'),
  }),

  words: {
    loadingTools: '載入 Word 工具…',
    title: 'JVision for Word',
    brandSub: 'for Word',
    loginLead: '登入你的 AI Agents Office 帳號，開始對這份文件提問。',
    histTip: '這份文件的對話紀錄',
    newChatTip: '開新對話（這份文件的紀錄會留著）',
    docNameTitle: '目前開啟的文件',
    untitled: '（未命名文件）',
    confirmTitle: 'AI 想要修改這份文件',
    noHistory: '這份文件還沒有對話紀錄。',
    stopped: '已停止。已經做完的部分留在文件裡，需要的話可以用上方的復原鍵退回。',
    localeChanged: '之後的回覆會用{0}。文件裡既有的內容不會被翻譯。',
    autoApproveOn: '已改成自動允許：AI 修改這份文件時不會再詢問。刪除整段仍會問，因為那沒辦法復原。',
    autoApproveOff: '已改回每次確認：AI 要修改文件前都會先問過你。',
    noSnapshot: '這份文件太大，改動前沒辦法先備份，所以這次的修改復原鍵拿不回來。',
    // Word's version of this note says something different from Excel's, and
    // better: the changes ARE in a review mechanism the person already knows.
    undoScope: '提醒：AI 的改動是以追蹤修訂寫入的，可以在 Word 的「校閱」裡逐條接受或拒絕；要一次全退回請用這裡的復原鍵。',
    locationLabel: '目前的位置：{0}',
    locationMissing: '找不到段落 {0}',
    emptyTitle: '問這份文件任何事',
    emptySub: '我讀得到章節、樣式和你選取的段落。改動會以追蹤修訂寫入，你可以逐條決定。',
    placeholder: '對這份文件提問…',
    chips: [
      '幫我摘要這份文件的重點',
      '把選取的段落改寫成正式公文',
      '檢查全文的錯字和用詞',
      '整理標題階層並產生目錄',
    ],
    intro: [
      {
        title: '在 Word 裡直接問',
        body: '不用切視窗，也不用把草稿貼到別的地方。選一段文字，用講的就好。',
        feats: ['選了就懂', '用問的', '就地改寫'],
      },
      {
        title: '它讀得懂這份文件',
        body: '讀得到章節階層、樣式和你選取的段落。可以摘要重點、抓出錯字、比對條款差異。',
        feats: ['抓重點', '揪錯字', '比對差異'],
      },
      {
        title: '改動都是追蹤修訂',
        body: '改寫、潤飾、套樣式都做得到。動筆之前會先問過你，寫進去的是修訂，可以逐條接受或拒絕。',
        feats: ['先問過你', '逐條接受', '隨時復原'],
      },
    ],
  },
};
