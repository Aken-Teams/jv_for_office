/**
 * PowerPoint — the host module for presentations.
 *
 * ── What makes this one different ──
 * Excel and Word both have a rich API and a natural coordinate system: A1, and
 * the paragraph ordinal. PowerPoint has neither. It has a *canvas* — every shape
 * carries left/top/width/height in points, and nothing lays anything out for
 * you. That inverts the job. In Word the model writes prose and the styles
 * handle the page; here the model has to decide where things go, and a
 * presentation where the boxes do not line up looks wrong from the back of the
 * room even when every word is right.
 *
 * So the centre of gravity of this module is geometry, not text. See GRID below.
 *
 * ── What the API actually supports, checked rather than assumed ──
 * The PowerPoint JS API is much narrower than Word's and the gaps are not where
 * you would guess. Verified against the requirement-set docs, 2026-08:
 *
 *   1.2  slides: count, getItemAt, delete, insertSlidesFromBase64
 *   1.3  slides.add, tags, shapes collection accessors
 *   1.4  THE workhorse: addGeometricShape / addTextBox / addLine, and
 *        Shape.left/top/width/height/name/type/fill/lineFormat/textFrame,
 *        TextFrame margins + verticalAlignment + wordWrap + autoSizeSetting,
 *        TextRange.font (name/size/bold/italic/color) and
 *        paragraphFormat.horizontalAlignment
 *   1.5  hyperlinks
 *   1.6  getSelectedSlides / getSelectedShapes / setSelectedSlides
 *   1.8  addTable, addGroup          ← Office 2504+, NOT on volume-licensed
 *   1.9  table formatting            ← same
 *   1.10 rotation, slide background, exportAsBase64Presentation
 *
 * Three of those gaps shape everything below, and each one is a decision the
 * model has to be told about rather than left to discover:
 *
 *   · `shapes.addPicture` exists but is PREVIEW ONLY — Microsoft's own docs say
 *     not to ship it. Pictures therefore go in through the Common API
 *     (`setSelectedDataAsync` + ImageCoercion), which is supported back to
 *     Office 2021. That route also accepts SVG, which is a better deal than
 *     Word got: a diagram inserted here is VECTOR, stays sharp on a projector,
 *     and the person can right-click → convert to shapes and edit it.
 *
 *   · `slides.add()` returns void and always APPENDS. There is no "insert after
 *     slide 3" and no way to get the new slide back directly — you add, sync,
 *     then read the last one. Both facts are stated in the tool descriptions,
 *     because a model that assumes otherwise produces a deck in the wrong order
 *     and reports success.
 *
 *   · There is NO reorder API at all. Not hidden, not gated — absent. So there
 *     is no ppt_move_slide, and the tool descriptions say to build in order
 *     rather than fixing it afterwards. Offering a tool that cannot work is
 *     worse than not offering one.
 *
 * ── Undo ──
 * Word snapshots the whole body as OOXML before every write, which is cheap and
 * total. PowerPoint has no equivalent below 1.10, and requiring 1.10 would mean
 * requiring Office 2601. So undo here is by INVERSE OPERATION: everything this
 * module creates is named and recorded, and undo deletes what it made. That
 * covers adding, which is nearly all of what an agent does to a deck. It does
 * not cover edits to shapes the person made themselves, and restoreSnapshot says
 * so rather than quietly restoring less than it claims.
 */

import { t } from '../i18n.js';
import { makeScriptGuard } from './guard.js';
// Lazily loaded inside: the 2.9 MB bundle it pulls in is only fetched the first
// time somebody asks for a diagram.
import { renderMermaid } from './mermaid.js';
// Pure string work — no canvas, no DOM. See svgchart.js for why vector.
import { chartSvg, CHART_TYPES } from './svgchart.js';

// ── Capability gates ─────────────────────────────────────────────────────────

export function hasApi(version) {
  try {
    return !!Office.context.requirements.isSetSupported('PowerPointApi', version);
  } catch {
    return false;
  }
}

/**
 * Can this build put a picture on a slide, and in what format?
 *
 * Two different requirement sets, and the difference is worth keeping: 1.1 is
 * raster, 1.2 adds SVG. A deck with vector diagrams is a materially better
 * artefact, so the SVG path is preferred and the PNG path is the fallback rather
 * than the default.
 */
function hasImageCoercion(version) {
  try {
    return !!Office.context.requirements.isSetSupported('ImageCoercion', version);
  } catch {
    return false;
  }
}

// ── The slide as a canvas ────────────────────────────────────────────────────

/**
 * Where things go on a 16:9 slide, in points.
 *
 * ── Why this exists at all ──
 * Nothing in PowerPoint's API lays anything out. `addTextBox` takes left, top,
 * width and height and does exactly what it is told, so if the model invents
 * coordinates per call, every slide in the deck has its title in a slightly
 * different place. Nobody can name the defect — the deck just looks amateur.
 *
 * One grid, used by every tool here, is what makes ten slides look like one
 * deck. The numbers are the standard 13.333in × 7.5in slide at 96 dpi, with the
 * margins PowerPoint's own built-in layouts use, so content lines up with any
 * native placeholder the person adds later.
 */
const GRID = {
  W: 960,
  H: 540,
  margin: 60,
  /** Title band: full width, top of the slide. */
  title: { left: 60, top: 48, width: 840, height: 72 },
  /** Body band: everything under the title. */
  body: { left: 60, top: 140, width: 840, height: 340 },
  /** Two columns with a gutter that is wide enough to read as a gap. */
  gutter: 32,
  /** Footer strip — page numbers, source notes. */
  footer: { left: 60, top: 494, width: 840, height: 28 },
};

/** Half-width column geometry, derived so the two always agree. */
const COL_W = (GRID.body.width - GRID.gutter) / 2;

/**
 * Type scale.
 *
 * Presentation type is much larger than document type and the jumps have to be
 * bigger: at 3 metres, 18pt and 20pt are the same size. These are the sizes the
 * themes below vary around.
 */
const TYPE = { display: 40, title: 32, heading: 22, body: 18, small: 14, caption: 11 };

/**
 * Themes.
 *
 * Deliberately the same idea as the Word host's: a small named set, so 「幫我把
 * 這份簡報做得專業一點」 has an answer that is one call rather than thirty font
 * assignments. Each is a full palette — an accent alone is not a theme, and a
 * deck where only the titles changed colour looks half-finished.
 */
/**
 * Themes.
 *
 * ── Rebuilt after measuring a deck that actually looks good ──
 * The first version gave each theme an accent, a tint and a body colour, which
 * sounds sufficient and is not. Pulling apart a reference deck built by hand
 * showed why: across ten slides it used TEN colours total, and they were not
 * three roles with variations — they were six distinct jobs, each used the same
 * way every time. Repetition is what makes a deck look designed; a palette that
 * only names "the accent" leaves the other five jobs to be improvised per slide.
 *
 * So each theme names every job:
 *
 *   strong   headings, number badges, the primary card bar        #1A3A5C
 *   accent   the top strip and the short rule under the title     #C0392B
 *   second   the alternate card bar, so a row of five is not flat #2E6DA4
 *   tint     card fill                                            #EEF3F8
 *   muted    eyebrows, page references, axis labels               #8A93A0
 *   ink      body text — NOT black. #3D3D3D reads softer on a projector
 *   line     hairlines and gridlines                              #D5DBE0
 *
 * ── cover / onCover exist because of a bug this shape prevents ──
 * The old title recipe filled the whole slide with `accent` and wrote in
 * `onAccent`. That works for four of the five palettes and inverts on the dark
 * one, where the accent is a LIGHT blue meant to sit on a dark ground — so the
 * dark theme's cover came out pale blue with near-black text. Naming the cover
 * fill and its text colour directly makes every palette state its own answer
 * rather than having one inferred from a rule that does not generalise.
 */
const THEMES = {
  corporate: {
    label: '企業藍', font: '微軟正黑體',
    ink: '#3D3D3D', strong: '#1A3A5C', accent: '#C0392B', second: '#2E6DA4',
    tint: '#EEF3F8', muted: '#8A93A0', line: '#D5DBE0',
    bg: '#FFFFFF', cover: '#1A3A5C', onCover: '#FFFFFF',
  },
  slate: {
    label: '沉穩灰', font: '微軟正黑體',
    ink: '#33383D', strong: '#2F4858', accent: '#B5651D', second: '#5E7C8B',
    tint: '#EFF2F4', muted: '#87909A', line: '#D8DDE1',
    bg: '#FFFFFF', cover: '#2F4858', onCover: '#FFFFFF',
  },
  fresh: {
    label: '清新綠', font: '微軟正黑體',
    ink: '#33403A', strong: '#0E6B3A', accent: '#C0392B', second: '#4C9B72',
    tint: '#EDF6F1', muted: '#849289', line: '#D6E2DB',
    bg: '#FFFFFF', cover: '#0E6B3A', onCover: '#FFFFFF',
  },
  warm: {
    label: '暖橘', font: '微軟正黑體',
    ink: '#40352D', strong: '#9C3D12', accent: '#1A3A5C', second: '#D08A50',
    tint: '#FBF2EA', muted: '#9C8B7C', line: '#E6D8CB',
    bg: '#FFFFFF', cover: '#9C3D12', onCover: '#FFFFFF',
  },
  night: {
    label: '深色', font: '微軟正黑體',
    // Every role inverts together. `cover` is the dark ground and `onCover` the
    // light type — the opposite of the light themes, stated rather than derived.
    ink: '#E8ECF1', strong: '#7FB2FF', accent: '#FF6B5B', second: '#4C9AFF',
    tint: '#1E2530', muted: '#98A3B3', line: '#2C3542',
    bg: '#141A22', cover: '#0D1219', onCover: '#F5F7FA',
  },
};

const DEFAULT_THEME = 'corporate';

const theme = (name) => THEMES[String(name || '').toLowerCase()] || THEMES[DEFAULT_THEME];

// ── Small helpers ────────────────────────────────────────────────────────────

const num = (v, fallback) => (Number.isFinite(Number(v)) ? Number(v) : fallback);
const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
const snip = (text, n) => {
  const s = String(text || '').replace(/\s+/g, ' ').trim();
  return s.length > n ? s.slice(0, n) + '…' : s;
};

/** PowerPoint wants #RRGGBB; models write 'red', '1F3864', '#1f3864'. */
function colour(value, fallback) {
  const s = String(value || '').trim();
  if (!s) return fallback;
  if (/^#[0-9a-f]{6}$/i.test(s)) return s.toUpperCase();
  if (/^[0-9a-f]{6}$/i.test(s)) return '#' + s.toUpperCase();
  return s;
}

const H_ALIGN = { left: 'Left', center: 'Center', centre: 'Center', centered: 'Center', right: 'Right', justify: 'Justify', justified: 'Justify' };
const V_ALIGN = { top: 'Top', middle: 'Middle', center: 'Middle', bottom: 'Bottom' };

/**
 * Everything this module creates gets a name carrying this prefix.
 *
 * Two jobs. Undo needs to know what it made, and 「把剛剛那張圖移到右邊」 needs a
 * way to address a shape that nobody gave a name to. PowerPoint auto-names
 * shapes 「TextBox 4」, which is unstable across edits; a name we assign is not.
 */
const TAG = 'jv';
let seq = 0;
const stamp = (kind) => `${TAG}:${kind}:${Date.now().toString(36)}${(seq++).toString(36)}`;

/** Shapes created during the current tool call, for undo-by-inverse. */
let made = [];
const remember = (slideId, shapeId) => made.push({ slideId, shapeId });

// ── Reading the deck ─────────────────────────────────────────────────────────

/**
 * Load slides with the properties every read path needs.
 *
 * `shapes/items/...` in one load rather than a call per slide: the bridge round
 * trip dominates everything else here, and a 30-slide deck asked shape by shape
 * takes long enough that the pane looks hung.
 */
async function loadSlides(ctx, opts = {}) {
  const slides = ctx.presentation.slides;
  const want = opts.shapes === false
    ? 'items/id'
    : 'items/id,items/shapes/items/id,items/shapes/items/name,items/shapes/items/type,'
      + 'items/shapes/items/left,items/shapes/items/top,items/shapes/items/width,items/shapes/items/height';
  slides.load(want);
  await ctx.sync();
  return slides.items;
}

/**
 * Text out of a shape, or '' when it has none.
 *
 * `shape.textFrame` throws InvalidArgument for shapes that cannot hold text —
 * pictures, lines — so this must be queued per shape and tolerated per shape,
 * not loaded in bulk with everything else.
 */
async function loadTexts(ctx, shapes) {
  const frames = shapes.map((s) => {
    try {
      const f = s.textFrame;
      f.load('hasText,textRange/text');
      return f;
    } catch {
      return null;
    }
  });
  try {
    await ctx.sync();
  } catch {
    return shapes.map(() => '');
  }
  return frames.map((f) => {
    if (!f) return '';
    try {
      return f.hasText ? String(f.textRange.text || '') : '';
    } catch {
      return '';
    }
  });
}

const OVERVIEW_PREVIEW_CHARS = 60;

/**
 * The deck, as a map the model can plan against.
 *
 * Per slide: its number, how many shapes it holds, and the text of the largest
 * one — which is very nearly always the title, and is a far more useful label
 * than the slide ID.
 */
async function getOverview(args = {}) {
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    if (!slides.length) {
      return '這份簡報目前沒有投影片。用 ppt_add_slide 新增一張再開始。';
    }

    const limit = clamp(num(args.slides, 40), 1, 200);
    const lines = [];
    for (let i = 0; i < Math.min(slides.length, limit); i++) {
      const shapes = slides[i].shapes.items || [];
      const texts = await loadTexts(ctx, shapes);
      // Largest text-bearing shape ≈ the title. Cheap, and right far more often
      // than trusting a placeholder type that half these decks do not use.
      let best = '';
      let bestArea = -1;
      shapes.forEach((s, j) => {
        if (!texts[j]) return;
        const area = (Number(s.width) || 0) * (Number(s.height) || 0);
        if (area > bestArea) { bestArea = area; best = texts[j]; }
      });
      const withText = texts.filter(Boolean).length;
      lines.push(`投影片 ${i + 1}：${best ? snip(best, OVERVIEW_PREVIEW_CHARS) : '（沒有文字）'}`
        + `　[${shapes.length} 個物件${withText ? `、${withText} 個有文字` : ''}]`);
    }
    if (slides.length > limit) lines.push(`…另外還有 ${slides.length - limit} 張投影片。`);

    return `這份簡報有 ${slides.length} 張投影片。\n\n${lines.join('\n')}\n\n`
      + `版面尺寸以 ${GRID.W}×${GRID.H} 點（16:9）計算，內文範圍是左 ${GRID.margin} 到右 ${GRID.W - GRID.margin}。\n`
      + '要看某一張的完整內容和每個物件的座標，用 ppt_read_slide。';
  });
}

/** One slide in full: every shape, with the geometry needed to rearrange it. */
async function readSlide(args) {
  const n = num(args.slide, 0);
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    if (!slides.length) return '這份簡報沒有投影片。';
    const idx = n ? clamp(n, 1, slides.length) - 1 : 0;
    const shapes = slides[idx].shapes.items || [];
    if (!shapes.length) return `投影片 ${idx + 1} 是空的。`;

    const texts = await loadTexts(ctx, shapes);
    const rows = shapes.map((s, i) => {
      const geo = `(${Math.round(s.left)},${Math.round(s.top)}) ${Math.round(s.width)}×${Math.round(s.height)}`;
      return `${i + 1}. ${s.name || '(未命名)'}　${s.type}　${geo}`
        + (texts[i] ? `\n    「${snip(texts[i], 200)}」` : '');
    });
    return `投影片 ${idx + 1}，共 ${shapes.length} 個物件：\n\n${rows.join('\n')}\n\n`
      + '座標是點，原點在投影片左上角。要移動或對齊用 ppt_arrange，改文字用 ppt_set_text。';
  });
}

/** What the person has selected right now. Needs 1.6. */
async function getSelection() {
  if (!hasApi('1.6')) return '這個 Office 版本讀不到選取範圍（需要 PowerPointApi 1.6）。';
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.getSelectedSlides();
    slides.load('items/id');
    const all = ctx.presentation.slides;
    all.load('items/id');
    await ctx.sync();

    const order = new Map((all.items || []).map((s, i) => [s.id, i + 1]));
    const picked = (slides.items || []).map((s) => order.get(s.id)).filter(Boolean);
    if (!picked.length) return '目前沒有選取任何投影片。';
    return picked.length === 1
      ? `目前選取的是投影片 ${picked[0]}。`
      : `目前選取的是投影片 ${picked.join('、')}。`;
  });
}

// ── Slides ───────────────────────────────────────────────────────────────────

async function addSlide(args) {
  if (!hasApi('1.3')) throw new Error('新增投影片需要 PowerPointApi 1.3，這個 Office 版本太舊。');
  const count = clamp(num(args.count, 1), 1, 30);
  return PowerPoint.run(async (ctx) => {
    const before = ctx.presentation.slides;
    before.load('items/id');
    await ctx.sync();
    const start = before.items.length;

    for (let i = 0; i < count; i++) {
      // Both IDs or neither: passing only a layoutId throws, because a layout is
      // only meaningful relative to its master.
      const opts = args.layout_id && args.master_id
        ? { slideMasterId: String(args.master_id), layoutId: String(args.layout_id) }
        : undefined;
      ctx.presentation.slides.add(opts);
    }
    await ctx.sync();

    return `已新增 ${count} 張投影片，現在是第 ${start + 1}`
      + (count > 1 ? `–${start + count}` : '') + ' 張。\n'
      + '注意：PowerPoint 的 API 只能把新投影片加在**最後面**，沒有「插在第 N 張後面」，'
      + '也沒有搬移投影片的方法。所以請照最終順序由前往後建立。';
  });
}

async function deleteSlide(args) {
  const from = num(args.slide, 0);
  if (!from) throw new Error('slide 是必要的（投影片編號，從 1 開始）。');
  const to = num(args.to, from);
  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx, { shapes: false });
    const a = clamp(from, 1, slides.length);
    const b = clamp(Math.max(to, a), a, slides.length);
    // Backwards: deleting slide 3 renumbers everything after it, so a forward
    // loop deletes the wrong slides from the second iteration onward.
    for (let i = b - 1; i >= a - 1; i--) slides[i].delete();
    await ctx.sync();
    return `已刪除投影片 ${a}${b > a ? `–${b}` : ''}（共 ${b - a + 1} 張）。這個動作復原鍵拿不回來。`;
  });
}

// ── Putting things on a slide ────────────────────────────────────────────────

/** Resolve the slide a tool should act on: explicit number, else the last one. */
async function targetSlide(ctx, args) {
  const slides = ctx.presentation.slides;
  slides.load('items/id');
  await ctx.sync();
  if (!slides.items.length) throw new Error('這份簡報沒有投影片，先用 ppt_add_slide 新增一張。');
  const n = num(args.slide, 0);
  const idx = n ? clamp(n, 1, slides.items.length) - 1 : slides.items.length - 1;
  return { slide: slides.items[idx], index: idx + 1, total: slides.items.length };
}

/**
 * Apply the typography and box settings every text-bearing shape shares.
 *
 * One function rather than repeated blocks: a text box and a shape with a label
 * in it should obey the same rules, and when they drift the deck stops looking
 * like one deck.
 */
function styleText(shape, opts) {
  const frame = shape.textFrame;
  const range = frame.textRange;
  const font = range.font;

  if (opts.font) font.name = String(opts.font);
  if (opts.size) font.size = Number(opts.size);
  if (opts.bold !== undefined) font.bold = !!opts.bold;
  if (opts.italic !== undefined) font.italic = !!opts.italic;
  if (opts.color) font.color = colour(opts.color, undefined);

  const h = H_ALIGN[String(opts.align || '').toLowerCase()];
  if (h) range.paragraphFormat.horizontalAlignment = h;

  const v = V_ALIGN[String(opts.valign || '').toLowerCase()];
  if (v) frame.verticalAlignment = v;

  if (opts.wrap !== undefined) frame.wordWrap = !!opts.wrap;
  // Text that overflows its box is the single most common way an auto-built
  // slide looks broken, so shrink-to-fit is the default rather than an option.
  frame.autoSizeSetting = opts.autofit === false ? 'AutoSizeNone' : 'AutoSizeTextToFitShape';
  if (opts.padding !== undefined) {
    const p = Number(opts.padding);
    frame.leftMargin = p; frame.rightMargin = p;
    frame.topMargin = p; frame.bottomMargin = p;
  }
}

/** Fill and outline, shared by shapes and boxes. */
function styleBox(shape, opts) {
  if (opts.fill) {
    shape.fill.setSolidColor(colour(opts.fill, '#FFFFFF'));
  } else if (opts.fill === null || opts.fill === '') {
    shape.fill.clear();
  }
  if (opts.line_color || opts.line_width !== undefined) {
    const lf = shape.lineFormat;
    lf.visible = true;
    if (opts.line_color) lf.color = colour(opts.line_color, '#000000');
    if (opts.line_width !== undefined) lf.weight = Number(opts.line_width);
  } else if (opts.line === false) {
    shape.lineFormat.visible = false;
  }
}

const box = (args, fallback) => ({
  left: num(args.left, fallback.left),
  top: num(args.top, fallback.top),
  width: num(args.width, fallback.width),
  height: num(args.height, fallback.height),
});

async function addText(args) {
  if (!hasApi('1.4')) throw new Error('新增文字方塊需要 PowerPointApi 1.4，這個 Office 版本太舊。');
  const text = String(args.text ?? '');
  if (!text) throw new Error('text 是必要的。');
  const th = theme(args.theme);

  return PowerPoint.run(async (ctx) => {
    const { slide, index } = await targetSlide(ctx, args);
    const geo = box(args, GRID.body);
    const shape = slide.shapes.addTextBox(text, geo);
    shape.name = stamp('text');
    styleText(shape, {
      font: args.font || th.font,
      size: num(args.size, TYPE.body),
      bold: args.bold,
      italic: args.italic,
      color: args.color || th.ink,
      align: args.align,
      valign: args.valign,
      wrap: args.wrap,
      autofit: args.autofit,
      padding: args.padding,
    });
    styleBox(shape, args);
    shape.load('id');
    await ctx.sync();
    remember(slide.id, shape.id);
    return `已在投影片 ${index} 加上文字方塊（${Math.round(geo.left)},${Math.round(geo.top)}）`
      + ` ${Math.round(geo.width)}×${Math.round(geo.height)}。`;
  });
}

async function addShape(args) {
  if (!hasApi('1.4')) throw new Error('新增圖形需要 PowerPointApi 1.4，這個 Office 版本太舊。');
  const kind = String(args.shape || 'Rectangle');
  const th = theme(args.theme);

  return PowerPoint.run(async (ctx) => {
    const { slide, index } = await targetSlide(ctx, args);
    const geo = box(args, { left: GRID.margin, top: GRID.body.top, width: 240, height: 120 });
    const shape = slide.shapes.addGeometricShape(kind, geo);
    shape.name = stamp('shape');
    styleBox(shape, { fill: args.fill || th.strong, ...args });
    if (args.text) {
      shape.textFrame.textRange.text = String(args.text);
      styleText(shape, {
        font: args.font || th.font,
        size: num(args.size, TYPE.body),
        bold: args.bold !== undefined ? args.bold : true,
        color: args.color || th.onCover,
        align: args.align || 'center',
        valign: args.valign || 'middle',
      });
    }
    shape.load('id');
    await ctx.sync();
    remember(slide.id, shape.id);
    return `已在投影片 ${index} 加上 ${kind}（${Math.round(geo.left)},${Math.round(geo.top)}）。`;
  });
}

async function addLine(args) {
  if (!hasApi('1.4')) throw new Error('新增線條需要 PowerPointApi 1.4。');
  const th = theme(args.theme);
  return PowerPoint.run(async (ctx) => {
    const { slide, index } = await targetSlide(ctx, args);
    // For a line, PowerPoint reads left/top as the START point and width/height
    // as the offset to the END point — not as a bounding box. Getting this wrong
    // draws a line somewhere plausible but incorrect, so it is spelled out here
    // and in the tool description.
    const geo = box(args, { left: GRID.margin, top: GRID.body.top, width: GRID.body.width, height: 0 });
    const line = slide.shapes.addLine(String(args.connector || 'Straight'), geo);
    line.name = stamp('line');
    const lf = line.lineFormat;
    lf.visible = true;
    lf.color = colour(args.line_color || args.color, th.strong);
    lf.weight = num(args.line_width, 1.5);
    line.load('id');
    await ctx.sync();
    remember(slide.id, line.id);
    return `已在投影片 ${index} 畫一條線，從 (${Math.round(geo.left)},${Math.round(geo.top)})`
      + ` 到 (${Math.round(geo.left + geo.width)},${Math.round(geo.top + geo.height)})。`;
  });
}

/** Rewrite the text of a shape that is already there. */
async function setText(args) {
  if (!hasApi('1.4')) throw new Error('改寫文字需要 PowerPointApi 1.4。');
  const which = num(args.shape, 0);
  if (!which) throw new Error('shape 是必要的（ppt_read_slide 列出的物件編號）。');
  const text = String(args.text ?? '');

  return PowerPoint.run(async (ctx) => {
    const { slide, index } = await targetSlide(ctx, args);
    const shapes = slide.shapes;
    shapes.load('items/id,items/name');
    await ctx.sync();
    const items = shapes.items || [];
    if (which > items.length) throw new Error(`投影片 ${index} 只有 ${items.length} 個物件，沒有第 ${which} 個。`);
    const shape = items[which - 1];

    shape.textFrame.textRange.text = text;
    styleText(shape, args);
    await ctx.sync();
    return `已改寫投影片 ${index} 第 ${which} 個物件的文字。`;
  });
}

/** Move, resize, recolour a shape that is already there. */
async function formatShape(args) {
  if (!hasApi('1.4')) throw new Error('調整物件需要 PowerPointApi 1.4。');
  const which = num(args.shape, 0);
  if (!which) throw new Error('shape 是必要的（ppt_read_slide 列出的物件編號）。');

  return PowerPoint.run(async (ctx) => {
    const { slide, index } = await targetSlide(ctx, args);
    const shapes = slide.shapes;
    shapes.load('items/id');
    await ctx.sync();
    const items = shapes.items || [];
    if (which > items.length) throw new Error(`投影片 ${index} 只有 ${items.length} 個物件。`);
    const shape = items[which - 1];

    const moved = [];
    for (const k of ['left', 'top', 'width', 'height']) {
      if (args[k] !== undefined) { shape[k] = Number(args[k]); moved.push(k); }
    }
    styleBox(shape, args);
    if (args.font || args.size || args.color || args.bold !== undefined || args.align) styleText(shape, args);
    await ctx.sync();
    return `已調整投影片 ${index} 第 ${which} 個物件${moved.length ? `（${moved.join('、')}）` : ''}。`;
  });
}

async function deleteShape(args) {
  const which = num(args.shape, 0);
  if (!which) throw new Error('shape 是必要的。');
  return PowerPoint.run(async (ctx) => {
    const { slide, index } = await targetSlide(ctx, args);
    const shapes = slide.shapes;
    shapes.load('items/id');
    await ctx.sync();
    const items = shapes.items || [];
    if (which > items.length) throw new Error(`投影片 ${index} 只有 ${items.length} 個物件。`);
    items[which - 1].delete();
    await ctx.sync();
    return `已刪除投影片 ${index} 的第 ${which} 個物件。`;
  });
}

// ── Layout: the part a web app cannot do ─────────────────────────────────────

/**
 * Align, distribute and size a set of shapes together.
 *
 * ── Why this is the tool that justifies the add-in ──
 * Everything else here has an equivalent somewhere: a web app can generate a
 * deck, and generating is easy. What it cannot do is open the file the person
 * already has, look at where the six boxes on slide 4 actually sit, and line
 * them up. That requires reading live geometry out of the open document, which
 * is precisely what an add-in is for and a generator is not.
 *
 * The operations are the ones PowerPoint's own ribbon has, because those are the
 * ones people already know they want — and because a person asking 「幫我把這幾
 * 個框對齊」 means the same thing PowerPoint means.
 */
const ARRANGE = new Set(['left', 'right', 'top', 'bottom', 'center_h', 'center_v',
  'distribute_h', 'distribute_v', 'same_width', 'same_height', 'same_size']);

async function arrange(args) {
  if (!hasApi('1.4')) throw new Error('排列物件需要 PowerPointApi 1.4。');
  const how = String(args.how || '').toLowerCase();
  if (!ARRANGE.has(how)) {
    throw new Error(`how 必須是這些其中之一：${[...ARRANGE].join('、')}。`);
  }

  return PowerPoint.run(async (ctx) => {
    const { slide, index } = await targetSlide(ctx, args);
    const shapes = slide.shapes;
    shapes.load('items/id,items/left,items/top,items/width,items/height');
    await ctx.sync();

    const all = shapes.items || [];
    // An explicit list, or everything on the slide. "Everything" is the common
    // case for 「這張投影片幫我對齊一下」 and saves a read round trip.
    const picked = Array.isArray(args.shapes) && args.shapes.length
      ? args.shapes.map((n) => all[num(n, 0) - 1]).filter(Boolean)
      : all;
    if (picked.length < 2) throw new Error('至少要有兩個物件才能對齊或分散。');

    const L = picked.map((s) => Number(s.left) || 0);
    const T = picked.map((s) => Number(s.top) || 0);
    const W = picked.map((s) => Number(s.width) || 0);
    const H = picked.map((s) => Number(s.height) || 0);
    const minL = Math.min(...L);
    const maxR = Math.max(...L.map((l, i) => l + W[i]));
    const minT = Math.min(...T);
    const maxB = Math.max(...T.map((tt, i) => tt + H[i]));

    if (how === 'left') picked.forEach((s) => { s.left = minL; });
    else if (how === 'right') picked.forEach((s, i) => { s.left = maxR - W[i]; });
    else if (how === 'top') picked.forEach((s) => { s.top = minT; });
    else if (how === 'bottom') picked.forEach((s, i) => { s.top = maxB - H[i]; });
    else if (how === 'center_h') {
      const mid = (minL + maxR) / 2;
      picked.forEach((s, i) => { s.left = mid - W[i] / 2; });
    } else if (how === 'center_v') {
      const mid = (minT + maxB) / 2;
      picked.forEach((s, i) => { s.top = mid - H[i] / 2; });
    } else if (how === 'distribute_h' || how === 'distribute_v') {
      const horiz = how === 'distribute_h';
      // Sorted by current position, so distributing does not also reorder them —
      // a distribute that shuffles the boxes is not what anybody meant.
      const order = picked.map((s, i) => i).sort((a, b) => (horiz ? L[a] - L[b] : T[a] - T[b]));
      const span = horiz ? maxR - minL : maxB - minT;
      const used = order.reduce((sum, i) => sum + (horiz ? W[i] : H[i]), 0);
      const gap = (span - used) / (order.length - 1);
      let at = horiz ? minL : minT;
      for (const i of order) {
        if (horiz) { picked[i].left = at; at += W[i] + gap; }
        else { picked[i].top = at; at += H[i] + gap; }
      }
    } else if (how === 'same_width' || how === 'same_size') {
      const w = num(args.value, Math.max(...W));
      picked.forEach((s) => { s.width = w; });
    }
    if (how === 'same_height' || how === 'same_size') {
      const h = num(args.value, Math.max(...H));
      picked.forEach((s) => { s.height = h; });
    }

    await ctx.sync();
    return `已對投影片 ${index} 的 ${picked.length} 個物件套用「${how}」。`;
  });
}

// ── The chrome every slide wears ─────────────────────────────────────────────

/**
 * The frame that makes ten slides look like one deck.
 *
 * ── What this is, and why it is not the model's decision ──
 * Measuring the reference deck, the striking thing was not any individual slide.
 * It was that the same four elements sat at exactly the same coordinates on
 * every one of them:
 *
 *   a full-bleed 5pt accent strip along the very top
 *   an eyebrow at (38,22) — section number and name, 11pt, grey
 *   the title at (38,42), 25pt, in the strong colour
 *   a SHORT accent rule at (38,89), 108pt wide — not full width
 *
 * That last detail is the one worth noticing: the rule stops at 108 points. A
 * full-width rule under a title reads as a divider between two things; a short
 * one reads as an underline belonging to the title. It is a small decision, made
 * once, repeated ten times — and repetition is where the impression of design
 * actually comes from.
 *
 * None of it is exposed as a parameter, because a model that chooses the strip's
 * height per slide will eventually choose differently, and then the deck has a
 * defect nobody can name. The model supplies words; the geometry is fixed.
 */
const CHROME = {
  strip: { left: 0, top: 0, width: GRID.W, height: 5 },
  eyebrow: { left: 38, top: 22, width: 828, height: 19 },
  title: { left: 38, top: 42, width: 878, height: 42 },
  rule: { left: 38, top: 89, width: 108, height: 4 },
  /** Content starts below the rule, with room to breathe. */
  contentTop: 124,
  contentLeft: 38,
  contentWidth: 884,
};

// ── One call, one finished slide ─────────────────────────────────────────────

/**
 * Build a whole slide from a named recipe.
 *
 * ── Why a recipe tool exists next to the primitives ──
 * A title, a rule, four bullets and a footer is six calls through the primitives
 * above, each a bridge round trip, each carrying coordinates the model made up.
 * That is slow AND it is where consistency goes to die: two slides built that
 * way have their titles 4pt apart and nobody can say why.
 *
 * This is the same trade the Word host makes with word_insert_block. The
 * primitives stay, because a recipe can never cover everything; but the ordinary
 * case — which is most cases — becomes one call against the shared GRID, and
 * every slide built this way lines up with every other one.
 */
const RECIPES = new Set(['title', 'section', 'bullets', 'two_column', 'compare', 'statement', 'quote']);

async function buildSlide(args) {
  if (!hasApi('1.4')) throw new Error('組版需要 PowerPointApi 1.4，這個 Office 版本太舊。');
  const kind = String(args.kind || 'bullets').toLowerCase();
  if (!RECIPES.has(kind)) throw new Error(`kind 必須是：${[...RECIPES].join('、')}。`);
  const th = theme(args.theme);

  return PowerPoint.run(async (ctx) => {
    // A recipe fills a slide, so it defaults to making a new one rather than
    // piling a second layout on top of whatever is already there.
    if (args.slide === undefined && args.append !== false && hasApi('1.3')) {
      ctx.presentation.slides.add();
      await ctx.sync();
    }
    const { slide, index } = await targetSlide(ctx, args);
    const put = { text: [], shape: [] };

    const text = (content, geo, style) => {
      const s = slide.shapes.addTextBox(String(content), geo);
      s.name = stamp('build');
      styleText(s, { font: th.font, wrap: true, ...style });
      put.text.push(s);
      return s;
    };
    const band = (geo, fill) => {
      const s = slide.shapes.addGeometricShape('Rectangle', geo);
      s.name = stamp('band');
      s.fill.setSolidColor(fill);
      s.lineFormat.visible = false;
      put.shape.push(s);
      return s;
    };

    const title = String(args.title || '');
    const body = Array.isArray(args.body) ? args.body.map(String)
      : (args.body ? String(args.body).split(/\r?\n/) : []);
    // A small square rather than 「•」: the reference deck uses one, and a round
    // dot is the mark every word processor puts there by default.
    const bullet = (lines) => lines.filter((l) => l.trim()).map((l) => `▪　${l.trim()}`).join('\n');

    // ── The chrome, on every slide that is not a full-bleed cover ──
    // Fixed geometry, applied the same way every time. See CHROME for why none
    // of it is a parameter.
    const dressed = kind !== 'title' && args.chrome !== false;
    if (dressed) {
      band(CHROME.strip, th.accent);
      if (args.eyebrow) {
        text(String(args.eyebrow).toUpperCase(), CHROME.eyebrow,
          { size: 11, color: th.muted, valign: 'middle' });
      }
      if (title) {
        text(title, CHROME.title, { size: 25, bold: true, color: th.strong, valign: 'middle' });
        band(CHROME.rule, th.accent);
      }
      if (args.page) {
        text(String(args.page), { left: 800, top: 22, width: 116, height: 19 },
          { size: 11, color: th.muted, align: 'right', valign: 'middle' });
      }
    }
    const CT = CHROME.contentTop;
    const CL = CHROME.contentLeft;
    const CW = CHROME.contentWidth;

    if (kind === 'title') {
      // The cover: the one slide where full-bleed and centred type are right.
      // `cover`/`onCover` rather than accent/onAccent — see THEMES for the bug
      // that distinction exists to prevent.
      band({ left: 0, top: 0, width: GRID.W, height: GRID.H }, th.cover);
      band({ left: 0, top: 0, width: GRID.W, height: 6 }, th.accent);
      text(title, { left: 80, top: 190, width: GRID.W - 160, height: 96 },
        { size: TYPE.display, bold: true, color: th.onCover, align: 'center', valign: 'middle' });
      // A short rule between title and subtitle, centred. The same idea as the
      // one under a slide title: it belongs to the words above it.
      band({ left: (GRID.W - 120) / 2, top: 300, width: 120, height: 3 }, th.accent);
      if (body.length) {
        text(body.join('　·　'), { left: 80, top: 322, width: GRID.W - 160, height: 40 },
          { size: TYPE.heading, color: th.onCover, align: 'center' });
      }
      if (args.footer) {
        text(String(args.footer), { left: 80, top: 452, width: GRID.W - 160, height: 28 },
          { size: TYPE.small, color: th.onCover, align: 'center' });
      }
    } else if (kind === 'section') {
      // A divider. The big number is what tells the audience where they are, and
      // it is the element people actually navigate by.
      band({ left: 0, top: 186, width: GRID.W, height: 168 }, th.tint);
      band({ left: 0, top: 186, width: 8, height: 168 }, th.accent);
      if (args.eyebrow) {
        text(String(args.eyebrow), { left: 76, top: 212, width: 200, height: 56 },
          { size: 54, bold: true, color: th.accent, valign: 'middle' });
      }
      text(title, { left: args.eyebrow ? 180 : 76, top: 220, width: 700, height: 48 },
        { size: TYPE.title, bold: true, color: th.strong, valign: 'middle' });
      if (body.length) {
        text(body[0], { left: args.eyebrow ? 180 : 76, top: 274, width: 700, height: 30 },
          { size: TYPE.body, color: th.muted });
      }
    } else if (kind === 'statement') {
      // One number, big. The slide that has exactly one job.
      text(title, { left: CL, top: 176, width: CW, height: 120 },
        { size: 84, bold: true, color: th.strong, align: 'center', valign: 'bottom' });
      band({ left: (GRID.W - 120) / 2, top: 310, width: 120, height: 4 }, th.accent);
      if (body.length) {
        text(body.join('\n'), { left: 140, top: 330, width: GRID.W - 280, height: 80 },
          { size: TYPE.heading, color: th.muted, align: 'center' });
      }
    } else if (kind === 'quote') {
      band({ left: CL, top: 150, width: 4, height: 200 }, th.accent);
      text('“', { left: CL + 20, top: 132, width: 90, height: 90 },
        { size: 96, bold: true, color: th.tint });
      text(title, { left: CL + 34, top: 168, width: CW - 60, height: 164 },
        { size: TYPE.title, color: th.ink, valign: 'middle' });
      if (body.length) {
        text(`— ${body.join('，')}`, { left: CL + 34, top: 344, width: CW - 60, height: 32 },
          { size: TYPE.body, color: th.muted });
      }
    } else if (kind === 'two_column' || kind === 'compare') {
      const left = Array.isArray(args.left) ? args.left.map(String) : String(args.left || '').split(/\r?\n/);
      const right = Array.isArray(args.right) ? args.right.map(String) : String(args.right || '').split(/\r?\n/);
      const colW = (CW - 28) / 2;
      const cols = [
        { head: args.left_title, lines: left, x: CL, tone: th.strong },
        { head: args.right_title, lines: right, x: CL + colW + 28, tone: th.second },
      ];
      const colH = 340;
      for (const c of cols) {
        let y = CT;
        if (kind === 'compare') {
          // Two panels with their own bars, so the sides read as two things.
          // Without it the gutter is the only cue, and at projector distance a
          // 28pt gap is not a cue.
          band({ left: c.x, top: y, width: colW, height: colH }, th.tint);
          band({ left: c.x, top: y, width: colW, height: 5 }, c.tone);
        }
        if (c.head) {
          text(String(c.head), { left: c.x + 18, top: y + 16, width: colW - 36, height: 32 },
            { size: TYPE.heading, bold: true, color: c.tone });
          y += 52;
        }
        text(bullet(c.lines), { left: c.x + 18, top: y + 10, width: colW - 36, height: CT + colH - y - 22 },
          { size: TYPE.body, color: th.ink });
      }
    } else {
      // bullets — the ordinary content slide, and the one that has to be right.
      text(bullet(body), { left: CL, top: CT, width: CW, height: 330 },
        { size: TYPE.body, color: th.ink });
    }

    if (args.footer && kind !== 'title') {
      band({ left: CL, top: 486, width: CW, height: 1 }, th.line);
      text(String(args.footer), { left: CL, top: 494, width: CW, height: 24 },
        { size: TYPE.caption, color: th.muted });
    }

    const all = [...put.shape, ...put.text];
    all.forEach((s) => s.load('id'));
    await ctx.sync();
    all.forEach((s) => remember(slide.id, s.id));

    return `已用「${kind}」版型做好投影片 ${index}，${all.length} 個物件，套用「${th.label}」配色。\n`
      + '每個物件都可以再用 ppt_format_shape 微調，或用 ppt_arrange 對齊。';
  });
}

/**
 * A row of cards.
 *
 * ── Why this is the single highest-value addition ──
 * Counting shapes on the reference deck: 17 to 47 per slide, against the four to
 * six the recipes were producing. That gap IS the difference in how the two look,
 * and almost all of it is cards — a tinted panel with a coloured bar, repeated
 * across the slide. Slides 2 and 6 of the reference are both nothing but rows of
 * these with different parameters.
 *
 * And the density is free. Forty shapes and five shapes are queued into the SAME
 * PowerPoint.run batch and cost the SAME single sync — one round trip either way.
 * The add-in was not slow because it was doing less; it was doing less for no
 * reason.
 *
 * ── The two variants, both taken from the reference ──
 *   bar: 'left'  a wide card with a 5pt spine down its left edge. For lists —
 *                agenda items, findings. Reads as a row you scan downward.
 *   bar: 'top'   a narrower card with a 5pt cap. For a row of peers — the five
 *                stages of a process, four pillars. Reads as a row you scan across.
 *
 * Colours cycle strong → second → accent so a row of five is not a flat wall of
 * one hue, which is the thing that makes auto-generated slides look generated.
 */
async function addCards(args) {
  if (!hasApi('1.4')) throw new Error('卡片需要 PowerPointApi 1.4，這個 Office 版本太舊。');
  const cards = Array.isArray(args.cards) ? args.cards.filter(Boolean) : [];
  if (!cards.length) throw new Error('cards 是必要的，至少要一張。');
  if (cards.length > 8) throw new Error('一排最多 8 張卡片——再多在 A4 寬的投影片上就窄到讀不了了。');

  const th = theme(args.theme);
  const bar = String(args.bar || (cards.length > 3 ? 'top' : 'left')).toLowerCase();
  const cols = clamp(num(args.columns, bar === 'top' ? cards.length : Math.min(2, cards.length)), 1, 8);
  const rows = Math.ceil(cards.length / cols);

  const gutter = num(args.gutter, 24);
  const left = num(args.left, CHROME.contentLeft);
  const top = num(args.top, CHROME.contentTop);
  const totalW = num(args.width, CHROME.contentWidth);
  const cw = (totalW - gutter * (cols - 1)) / cols;
  const ch = num(args.height, bar === 'top' ? 92 : 68);
  const pitchY = ch + num(args.row_gap, 18);

  const cycle = [th.strong, th.second, th.accent];

  return PowerPoint.run(async (ctx) => {
    const { slide, index } = await targetSlide(ctx, args);
    const made_ = [];
    const add = (shape) => { made_.push(shape); return shape; };

    cards.forEach((card, i) => {
      const c = i % cols;
      const r = Math.floor(i / cols);
      const x = left + (cw + gutter) * c;
      const y = top + pitchY * r;
      const accent = colour(card.color, cycle[i % cycle.length]);

      // The panel. roundRect for a top-bar card and plain rect for a left-bar
      // one, matching the reference — a spine down the side of a rounded box
      // does not sit flush against the corner radius.
      const panel = add(slide.shapes.addGeometricShape(
        bar === 'top' ? 'RoundRectangle' : 'Rectangle',
        { left: x, top: y, width: cw, height: ch }));
      panel.name = stamp('card');
      panel.fill.setSolidColor(colour(card.fill, bar === 'top' ? th.bg : th.tint));
      if (bar === 'top') {
        panel.lineFormat.visible = true;
        panel.lineFormat.color = accent;
        panel.lineFormat.weight = 1;
      } else {
        panel.lineFormat.visible = false;
      }

      // The bar. A separate shape rather than a thick border, because a border
      // is drawn on all four sides and this must be on exactly one.
      const spine = add(slide.shapes.addGeometricShape('Rectangle', bar === 'top'
        ? { left: x, top: y, width: cw, height: 5 }
        : { left: x, top: y, width: 5, height: ch }));
      spine.name = stamp('bar');
      spine.fill.setSolidColor(accent);
      spine.lineFormat.visible = false;

      const padX = bar === 'top' ? 14 : 16;
      let tx = x + padX + (bar === 'left' ? 5 : 0);
      let tw = cw - padX * 2 - (bar === 'left' ? 5 : 0);
      let ty = y + (bar === 'top' ? 16 : 12);

      // The number badge. Big, in the accent, set to the LEFT of the text rather
      // than above it — it is an index, not a heading.
      if (card.number) {
        const badge = add(slide.shapes.addTextBox(String(card.number),
          { left: tx, top: ty, width: 52, height: 40 }));
        badge.name = stamp('num');
        styleText(badge, { font: th.font, size: 24, bold: true, color: accent, valign: 'top' });
        badge.fill.clear();
        badge.lineFormat.visible = false;
        tx += 56;
        tw -= 56;
      }

      if (card.title) {
        const title = add(slide.shapes.addTextBox(String(card.title),
          { left: tx, top: ty, width: tw, height: card.text ? 24 : ch - 24 }));
        title.name = stamp('ct');
        styleText(title, {
          font: th.font, size: num(args.size, 14), bold: true,
          color: card.title_color ? colour(card.title_color) : th.ink, valign: 'top',
        });
        title.fill.clear();
        title.lineFormat.visible = false;
        ty += 26;
      }
      if (card.text) {
        const body = add(slide.shapes.addTextBox(String(card.text),
          { left: tx, top: ty, width: tw, height: y + ch - ty - 10 }));
        body.name = stamp('cb');
        styleText(body, { font: th.font, size: num(args.small, 11.5), color: th.muted, valign: 'top' });
        body.fill.clear();
        body.lineFormat.visible = false;
      }
      // The right-hand reference (「P.03」 in the reference deck), for agendas.
      if (card.note && bar === 'left') {
        const note = add(slide.shapes.addTextBox(String(card.note),
          { left: x + cw - 74, top: y + ch / 2 - 14, width: 58, height: 28 }));
        note.name = stamp('cn');
        styleText(note, { font: th.font, size: 11, color: th.muted, align: 'right', valign: 'middle' });
        note.fill.clear();
        note.lineFormat.visible = false;
      }
    });

    // Arrows between cards, for a process row. Drawn after the cards so they sit
    // in the gutters rather than under a panel.
    if (args.arrows && bar === 'top' && cols > 1) {
      for (let c = 0; c < Math.min(cols, cards.length) - 1; c++) {
        const ax = left + (cw + gutter) * c + cw + 4;
        const arrow = add(slide.shapes.addGeometricShape('RightArrow',
          { left: ax, top: top + ch / 2 - 7, width: gutter - 8, height: 14 }));
        arrow.name = stamp('arrow');
        arrow.fill.setSolidColor(th.line);
        arrow.lineFormat.visible = false;
      }
    }

    made_.forEach((sh) => sh.load('id'));
    await ctx.sync();
    made_.forEach((sh) => remember(slide.id, sh.id));

    return `已在投影片 ${index} 排了 ${cards.length} 張卡片`
      + `（${cols} 欄${rows > 1 ? ` × ${rows} 列` : ''}、${bar === 'top' ? '頂部' : '左側'}色條）`
      + `，共 ${made_.length} 個物件。`;
  });
}

// ── Pictures, diagrams, charts ───────────────────────────────────────────────

/**
 * Put an image on the current slide.
 *
 * ── Why this does not use the PowerPoint API ──
 * `slide.shapes.addPicture` exists and would be the obvious call, but Microsoft
 * marks it PREVIEW ONLY with an explicit "do not use in production". So pictures
 * go through the Common API instead: `setSelectedDataAsync` with ImageCoercion,
 * supported back to Office 2021 on Windows, Mac, iPad and the web.
 *
 * The trade is that it is callback-based, it lands on whatever slide is
 * SELECTED rather than one we name, and it cannot be batched with the shape work
 * above. All three are worth it to not ship a preview API.
 */
function insertImageData(base64, coercion, geo) {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(base64, {
      coercionType: coercion,
      imageLeft: geo.left,
      imageTop: geo.top,
      imageWidth: geo.width,
      imageHeight: geo.height,
    }, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve(true);
      else reject(new Error((res.error && res.error.message) || '插入圖片失敗。'));
    });
  });
}

/** Fit a natural size into the body area, keeping the aspect ratio. */
function fitToBody(w, h, args) {
  const maxW = num(args.width, GRID.body.width);
  const maxH = num(args.height, GRID.body.height);
  const scale = Math.min(maxW / w, maxH / h, 1);
  const width = Math.round(w * scale);
  const height = Math.round(h * scale);
  return {
    width,
    height,
    left: num(args.left, Math.round((GRID.W - width) / 2)),
    top: num(args.top, Math.round(GRID.body.top + (GRID.body.height - height) / 2)),
  };
}

async function addImage(args) {
  if (!hasImageCoercion('1.1')) {
    throw new Error('這個 Office 版本不支援插入圖片（需要 ImageCoercion 1.1，Office 2021 以上）。');
  }
  const data = String(args.data || '').replace(/^data:[^,]+,/, '');
  if (!data) throw new Error('data 是必要的（base64 圖片資料）。');
  const geo = fitToBody(num(args.natural_width, 800), num(args.natural_height, 450), args);
  await insertImageData(data, Office.CoercionType.Image, geo);
  return `已插入圖片，${geo.width}×${geo.height} 點，位置 (${geo.left},${geo.top})。\n`
    + '它加在**目前選取的投影片**上——Common API 沒辦法指定投影片。';
}

/** The Word host's insertImage contract: a pasted picture from the composer. */
export async function insertImage(pic, args = {}) {
  const data = String((pic && pic.data) || '').replace(/^data:[^,]+,/, '');
  if (!data) throw new Error('沒有圖片資料。');
  if (!hasImageCoercion('1.1')) {
    throw new Error('這個 Office 版本不支援插入圖片（需要 ImageCoercion 1.1）。');
  }
  const geo = fitToBody(num(pic.width, 800), num(pic.height, 450), args);
  await insertImageData(data, Office.CoercionType.Image, geo);
  return `已把圖片放到目前的投影片，${geo.width}×${geo.height} 點。`;
}

/**
 * A mermaid diagram, as vector where possible.
 *
 * ── The one place PowerPoint beats Word ──
 * The Word host renders mermaid to a PNG because that is all Word takes. Here,
 * ImageCoercion 1.2 accepts SVG — so the diagram goes in as VECTOR: sharp at any
 * projector resolution, and the person can right-click → 轉換為圖形 and edit the
 * boxes. That is a genuinely better artefact, so SVG is the default and the PNG
 * path is the fallback for builds without 1.2, not the other way round.
 */
async function addDiagram(args) {
  const code = String(args.code || '').trim();
  if (!code) throw new Error('code 是必要的（mermaid 語法）。');
  if (!hasImageCoercion('1.1')) {
    throw new Error('這個 Office 版本不支援插入圖片（需要 ImageCoercion 1.1，Office 2021 以上）。');
  }

  const png = await renderMermaid(code, { theme: args.theme, font: args.font });
  const geo = fitToBody(png.width, png.height, args);

  const vector = hasImageCoercion('1.2') && args.raster !== true && png.svg;
  if (vector) {
    await insertImageData(png.svg, Office.CoercionType.XmlSvg, geo);
    return `已插入向量圖表，${geo.width}×${geo.height} 點。\n`
      + '它是 SVG，投影時不會糊掉；使用者在 PowerPoint 裡按右鍵「轉換為圖形」之後，'
      + '還可以直接拖動裡面的方塊。';
  }
  await insertImageData(png.data, Office.CoercionType.Image, geo);
  return `已插入圖表，${geo.width}×${geo.height} 點（點陣圖）。\n`
    + '這個 Office 版本不支援 SVG（要 ImageCoercion 1.2），所以放的是圖片，不能在 PowerPoint 裡編輯內容。';
}

/**
 * A chart, drawn here and inserted as vector.
 *
 * ── The trade, stated plainly ──
 * PowerPoint's JS API cannot create a chart. There is no gate to lift and no
 * newer requirement set that adds one — the capability is absent. So this draws
 * the chart itself and inserts it as SVG.
 *
 * What that buys: it is vector, so it stays sharp on a projector, and the person
 * can right-click → 轉換為圖形 and edit every bar as a real object.
 *
 * What it costs, and the tool description says so rather than letting the model
 * imply otherwise: it is NOT a PowerPoint chart object. There is no linked
 * datasheet. Changing a number means asking for the chart again.
 */
async function addChart(args) {
  if (!hasImageCoercion('1.2')) {
    throw new Error(
      '這個 Office 版本不支援插入向量圖（需要 ImageCoercion 1.2，Office 2021 以上）。'
      + '圖表沒辦法用點陣圖替代——投影出來會糊，而且不能編輯。'
      + '請改用 ppt_add_cards 把數字做成數據卡。');
  }
  const th = theme(args.theme);
  const svg = chartSvg({
    type: args.type,
    title: args.title,
    categories: args.categories,
    values: args.values,
    series: args.series,
    colors: args.colors || [th.strong, th.second, th.accent],
    data_labels: args.data_labels,
    area: args.area,
    center: args.center,
    center_label: args.center_label,
    width: num(args.width, 760),
    height: num(args.height, 380),
  }, th);

  const geo = fitToBody(num(args.width, 760), num(args.height, 380), args);
  await insertImageData(svg, Office.CoercionType.XmlSvg, geo);
  return `已插入${args.title ? `「${args.title}」` : ''}圖表，${geo.width}×${geo.height} 點。\n`
    + '它是 SVG 向量圖：投影不會糊，使用者按右鍵「轉換為圖形」之後每一根長條都可以個別編輯。\n'
    + '但它**不是 PowerPoint 的圖表物件**，沒有連結的資料表——要改數字就重新產生一張。';
}

// ── Theme sweep ──────────────────────────────────────────────────────────────

/**
 * Push one palette through every slide.
 *
 * Only touches shapes this module made — see TAG. Repainting a person's own
 * boxes because they happened to be on the same slide is not a theme, it is
 * damage, and it is not undoable here.
 */
async function applyTheme(args) {
  if (!hasApi('1.4')) throw new Error('套用配色需要 PowerPointApi 1.4。');
  const th = theme(args.theme);
  const onlyOurs = args.all !== true;

  return PowerPoint.run(async (ctx) => {
    const slides = await loadSlides(ctx);
    let touched = 0;
    for (const slide of slides) {
      for (const shape of slide.shapes.items || []) {
        if (onlyOurs && !String(shape.name || '').startsWith(`${TAG}:`)) continue;
        try {
          shape.textFrame.textRange.font.name = th.font;
          touched++;
        } catch {
          // Lines and pictures have no text frame. Not an error.
        }
      }
    }
    await ctx.sync();
    return `已把「${th.label}」的字型套到 ${touched} 個物件`
      + (onlyOurs ? '（只動這個增益集建立的物件；使用者自己放的內容沒有被改）' : '（包含使用者自己放的物件）')
      + `。\n配色：主色 ${th.strong}、強調 ${th.accent}、內文 ${th.ink}、輔助 ${th.muted}。`;
  });
}

// ── Script ───────────────────────────────────────────────────────────────────

const scriptGuard = makeScriptGuard;

async function runScript(args, opts = {}) {
  if (!opts.writeApproved && !args.read_only) {
    throw new Error('這段程式會修改簡報，需要先取得使用者同意。');
  }
  const code = String(args.code || '').trim();
  if (!code) throw new Error('code 是必要的。');

  const logs = [];
  const store = {};
  return PowerPoint.run(async (ctx) => {
    const guard = scriptGuard(!args.read_only, 'ppt_run_script');
    const guarded = guard(ctx);
    // eslint-disable-next-line no-new-func
    const fn = new Function('context', 'presentation', 'PowerPoint', 'log', 'store',
      `"use strict"; return (async () => { ${code} })();`);
    const out = await fn(guarded, guarded.presentation, PowerPoint,
      (...a) => logs.push(a.map(String).join(' ')), store);
    await ctx.sync();
    return [
      out === undefined ? '' : `回傳值：${JSON.stringify(out)}`,
      logs.length ? `log：\n${logs.join('\n')}` : '',
    ].filter(Boolean).join('\n') || '執行完成，沒有輸出。';
  });
}

// ── Undo ─────────────────────────────────────────────────────────────────────

/**
 * What this tool call created, so it can be taken back.
 *
 * Not a snapshot. PowerPoint has no cheap whole-file snapshot below API 1.10
 * (Office 2601, released weeks ago), and gating undo on that would mean almost
 * nobody has it. Recording what was made and deleting it covers additions, which
 * is nearly everything an agent does to a deck.
 *
 * What it does NOT cover is stated plainly in restoreSnapshot rather than
 * silently doing less than the button promises.
 */
async function captureCreated() {
  const snapshot = made.length ? { kind: 'created', shapes: made.slice() } : null;
  made = [];
  return snapshot;
}

export async function restoreSnapshot(snapshot) {
  if (!snapshot || snapshot.kind !== 'created' || !snapshot.shapes.length) {
    return { kind: 'none', note: '這次的動作沒有可以退回的新增物件。' };
  }
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load('items/id');
    await ctx.sync();
    const bySlide = new Map((slides.items || []).map((s) => [s.id, s]));

    let removed = 0;
    for (const { slideId, shapeId } of snapshot.shapes) {
      const slide = bySlide.get(slideId);
      if (!slide) continue;
      try {
        slide.shapes.getItem(shapeId).delete();
        removed++;
      } catch {
        // Already gone — the person deleted it themselves. Fine.
      }
    }
    await ctx.sync();
    return {
      kind: 'created',
      removed,
      note: `已移除這次新增的 ${removed} 個物件。`
        + '注意：PowerPoint 沒有整份快照的 API，所以復原只能拿掉新增的東西，'
        + '改過的文字和搬過的位置退不回來——那些請用 PowerPoint 自己的 Ctrl+Z。',
    };
  });
}

// ── Shell contract ───────────────────────────────────────────────────────────

/** Jump to a slide. The citation system's other half. */
export async function goTo(from) {
  const n = Number(from);
  if (!Number.isFinite(n) || !hasApi('1.6')) return false;
  try {
    return await PowerPoint.run(async (ctx) => {
      const slides = ctx.presentation.slides;
      slides.load('items/id');
      await ctx.sync();
      const items = slides.items || [];
      if (n < 1 || n > items.length) return false;
      ctx.presentation.setSelectedSlides([items[n - 1].id]);
      await ctx.sync();
      return true;
    });
  } catch {
    return false;
  }
}

export async function getDocumentName() {
  try {
    return String(Office.context.document.url || '').split(/[\\/]/).pop() || '';
  } catch {
    return '';
  }
}

export async function getDocumentKey() {
  try {
    const url = String(Office.context.document.url || '');
    if (url) return `ppt:${url}`;
  } catch { /* fall through */ }
  return '';
}

/** Which slide the person is on, for the bar at the top of the pane. */
async function currentSlide() {
  if (!hasApi('1.6')) return 0;
  try {
    return await PowerPoint.run(async (ctx) => {
      const sel = ctx.presentation.getSelectedSlides();
      sel.load('items/id');
      const all = ctx.presentation.slides;
      all.load('items/id');
      await ctx.sync();
      if (!sel.items.length) return 0;
      const order = new Map((all.items || []).map((s, i) => [s.id, i + 1]));
      return order.get(sel.items[0].id) || 0;
    });
  } catch {
    return 0;
  }
}

export async function getLocation() {
  const n = await currentSlide();
  return n ? `投影片 ${n}` : '';
}

const locationListeners = [];
const selectionListeners = [];
let bound = false;

/**
 * One handler for both listeners.
 *
 * PowerPoint has no onSelectionChanged on the presentation object — the same
 * trap the Word host hit. The Common API event is what exists, so it is what is
 * used, and it is bound once and fanned out rather than registered twice.
 */
function bindSelectionHandler() {
  if (bound) return;
  bound = true;
  try {
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
      void getLocation().then((loc) => locationListeners.forEach((fn) => fn(loc)));
      void getSelectionRef().then((ref) => selectionListeners.forEach((fn) => fn(ref)));
    });
  } catch {
    bound = false;
  }
}

export async function watchLocation(onChange) {
  locationListeners.push(onChange);
  bindSelectionHandler();
}

export async function getSelectionRef() {
  const n = await currentSlide();
  if (!n) return null;
  return {
    label: `投影片 ${n}`,
    context: `使用者在側邊欄提問時，正在看投影片 ${n}。`
      + '需要知道上面有什麼的話，用 ppt_read_slide 讀它。',
  };
}

export async function watchSelection(onChange) {
  selectionListeners.push(onChange);
  bindSelectionHandler();
}

/** What a write would cost, for the confirmation card. */
export async function inspectWrite(tool, args) {
  const a = args || {};
  if (tool === 'ppt_delete_slide') {
    const from = num(a.slide, 0);
    const to = num(a.to, from);
    return `要刪除投影片 ${from}${to > from ? `–${to}` : ''}，共 ${Math.max(1, to - from + 1)} 張。`
      + '刪除的投影片這裡的復原鍵拿不回來。';
  }
  if (tool === 'ppt_build_slide') {
    return `要新增一張「${a.kind || 'bullets'}」版型的投影片：${snip(a.title || '（沒有標題）', 40)}`;
  }
  return '';
}

export function supportedTools() {
  const tools = [
    'ppt_get_overview',
    'ppt_read_slide',
    'ppt_get_selection',
    'ppt_add_slide',
    'ppt_delete_slide',
    'ppt_build_slide',
    'ppt_add_text',
    'ppt_add_shape',
    'ppt_add_line',
    'ppt_add_cards',
    'ppt_set_text',
    'ppt_format_shape',
    'ppt_delete_shape',
    'ppt_arrange',
    'ppt_apply_theme',
    'ppt_ask_user',
    'ppt_read_attachment',
    'ppt_insert_attachment',
    'ppt_read_file',
  ];
  // Pictures are a Common API capability, not a PowerPointApi one, so they are
  // gated on their own requirement set rather than on the host version.
  if (hasImageCoercion('1.1')) tools.push('ppt_add_image', 'ppt_add_diagram');
  // Charts are vector-only on purpose: a rasterised chart on a projector is
  // worse than no chart, so this needs 1.2 rather than falling back to 1.1.
  if (hasImageCoercion('1.2')) tools.push('ppt_add_chart');
  return tools;
}

const WRITES = new Set([
  'ppt_add_slide', 'ppt_delete_slide', 'ppt_build_slide', 'ppt_add_text',
  'ppt_add_shape', 'ppt_add_line', 'ppt_add_cards', 'ppt_add_chart',
  'ppt_set_text', 'ppt_format_shape',
  'ppt_delete_shape', 'ppt_arrange', 'ppt_apply_theme', 'ppt_add_image',
  'ppt_add_diagram', 'ppt_run_script',
]);

export async function executeTool(tool, args, opts = {}) {
  const a = args || {};
  made = [];
  try {
    switch (tool) {
      case 'ppt_get_overview':  return { ok: true, content: await getOverview(a) };
      case 'ppt_read_slide':    return { ok: true, content: await readSlide(a) };
      case 'ppt_get_selection': return { ok: true, content: await getSelection() };

      case 'ppt_run_script':
        return { ok: true, content: await runScript(a, opts) };

      default: {
        const run = {
          ppt_add_slide: addSlide,
          ppt_delete_slide: deleteSlide,
          ppt_build_slide: buildSlide,
          ppt_add_text: addText,
          ppt_add_shape: addShape,
          ppt_add_line: addLine,
          ppt_add_cards: addCards,
          ppt_add_chart: addChart,
          ppt_set_text: setText,
          ppt_format_shape: formatShape,
          ppt_delete_shape: deleteShape,
          ppt_arrange: arrange,
          ppt_apply_theme: applyTheme,
          ppt_add_image: addImage,
          ppt_add_diagram: addDiagram,
        }[tool];
        if (!run) return { ok: false, content: `這個增益集不認得工具 ${tool}。` };
        const content = await run(a);
        return { ok: true, content, snapshot: WRITES.has(tool) ? await captureCreated() : undefined };
      }
    }
  } catch (e) {
    const code = e && e.code ? `[${e.code}] ` : '';
    return { ok: false, content: `${code}${(e && e.message) || String(e)}` };
  }
}

export const meta = {
  id: 'powerpoint',
  officeHost: 'PowerPoint',
  api: 'ppt',
  mcpPrefix: 'mcp__ppt__',
  overview: { tool: 'ppt_get_overview', args: {} },
  askUserTool: 'ppt_ask_user',
  attachmentTool: 'ppt_read_attachment',
  insertAttachmentTool: 'ppt_insert_attachment',
  inspectTools: ['ppt_delete_slide', 'ppt_build_slide'],
  scriptTool: 'ppt_run_script',

  /** Slides are the coordinate system, the way A1 is Excel's. */
  citation: {
    pattern: /\[\[投影片\s*(\d{1,4})\]\]/g,
    label: (a) => `投影片 ${a}`,
  },

  /** The slide canvas, so the pane can show it and the prompt can quote it. */
  canvas: { width: GRID.W, height: GRID.H, grid: GRID, chrome: CHROME, type: TYPE, themes: THEMES },

  toolLabels: () => ({
    ppt_get_overview: t('正在讀取簡報結構…'),
    ppt_read_slide: t('正在讀這張投影片…'),
    ppt_get_selection: t('正在讀取選取範圍…'),
    ppt_add_slide: t('正在新增投影片…'),
    ppt_delete_slide: t('等待你確認刪除…'),
    ppt_build_slide: t('正在組一張投影片…'),
    ppt_add_text: t('正在加上文字…'),
    ppt_add_shape: t('正在加上圖形…'),
    ppt_add_line: t('正在畫線…'),
    ppt_add_cards: t('正在排卡片…'),
    ppt_add_chart: t('正在畫圖表…'),
    ppt_set_text: t('正在改寫文字…'),
    ppt_format_shape: t('正在調整物件…'),
    ppt_delete_shape: t('正在移除物件…'),
    ppt_arrange: t('正在對齊排列…'),
    ppt_apply_theme: t('正在套用配色…'),
    ppt_add_image: t('正在插入圖片…'),
    ppt_add_diagram: t('正在畫流程圖…'),
    ppt_read_attachment: t('正在看你貼的圖…'),
    ppt_insert_attachment: t('正在把圖片放進簡報…'),
    ppt_read_file: t('正在讀你上傳的檔案…'),
    ppt_run_script: t('正在執行程式…'),
  }),

  preparingLabels: () => ({
    ppt_build_slide: t('正在設計投影片…'),
    ppt_arrange: t('正在盤算版面…'),
    ppt_add_diagram: t('正在設計流程圖…'),
    ppt_add_cards: t('正在設計版面…'),
  }),

  words: {
    loadingTools: '載入 PowerPoint 工具…',
    title: 'JVision for PowerPoint',
    brandSub: 'for PowerPoint',
    loginLead: '登入你的 AI Agents Office 帳號，開始對這份簡報提問。',
    histTip: '這份簡報的對話紀錄',
    newChatTip: '開新對話（這份簡報的紀錄會留著）',
    docNameTitle: '目前開啟的簡報',
    untitled: '（未命名簡報）',
    confirmTitle: 'AI 想要修改這份簡報',
    noHistory: '這份簡報還沒有對話紀錄。',
    stopped: '已停止。已經做好的投影片留在檔案裡，需要的話可以用上方的復原鍵退回。',
    localeChanged: '之後的回覆會用{0}。簡報裡既有的內容不會被翻譯。',
    autoApproveOn: '已改成自動允許：AI 修改這份簡報時不會再詢問。刪除投影片仍會問，因為那沒辦法復原。',
    autoApproveOff: '已改回每次確認：AI 要修改簡報前都會先問過你。',
    noSnapshot: 'PowerPoint 沒有整份備份的 API，所以復原鍵只能拿掉新增的物件。',
    // Honest about a real limitation rather than implying a fuller undo.
    undoScope: '提醒：PowerPoint 沒有整份快照的 API，所以這裡的復原鍵只能移除 AI 新增的物件；'
      + '改過的文字和搬過的位置請用 PowerPoint 自己的 Ctrl+Z。',
    locationLabel: '目前的位置：{0}',
    locationMissing: '找不到投影片 {0}',
    emptyTitle: '問這份簡報任何事',
    emptySub: '我讀得到每張投影片上物件的位置和文字，可以幫你排版、對齊，或直接做出新的投影片。',
    placeholder: '對這份簡報提問…',
    chips: [
      '幫我把這份簡報做成一致的版面',
      '這張投影片的物件幫我對齊',
      '依這些重點做三張投影片',
      '幫我做一張架構流程圖',
    ],
    intro: [
      {
        title: '在 PowerPoint 裡直接做',
        body: '不用切視窗，也不用把大綱貼到別的地方。看著投影片，用講的就好。',
        feats: ['看著改', '用問的', '就地排版'],
      },
      {
        title: '它看得到版面',
        body: '讀得到每個物件的座標和大小，所以可以真的幫你對齊、分散、統一尺寸——這是網頁版做不到的。',
        feats: ['讀得到座標', '對齊分散', '統一版型'],
      },
      {
        title: '一句話一張投影片',
        body: '封面、章節頁、條列、雙欄對照、大數字，都有現成版型，位置和字級是同一套格線。',
        feats: ['現成版型', '一致格線', '向量圖表'],
      },
    ],
  },
};
