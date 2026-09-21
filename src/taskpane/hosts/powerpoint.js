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
 *   ground   the PAGE, not the card — off-white, not white           #F7F9FB
 *   warn     a fourth tone, so a row of four KPIs is not three+one   #C9A227
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
    bg: '#FFFFFF', ground: '#F7F9FB', warn: '#C9A227', cover: '#1A3A5C', onCover: '#FFFFFF',
  },
  slate: {
    label: '沉穩灰', font: '微軟正黑體',
    ink: '#33383D', strong: '#2F4858', accent: '#B5651D', second: '#5E7C8B',
    tint: '#EFF2F4', muted: '#87909A', line: '#D8DDE1',
    bg: '#FFFFFF', ground: '#F5F7F8', warn: '#A88B4A', cover: '#2F4858', onCover: '#FFFFFF',
  },
  fresh: {
    label: '清新綠', font: '微軟正黑體',
    ink: '#33403A', strong: '#0E6B3A', accent: '#C0392B', second: '#4C9B72',
    tint: '#EDF6F1', muted: '#849289', line: '#D6E2DB',
    bg: '#FFFFFF', ground: '#F5FAF7', warn: '#C99A2E', cover: '#0E6B3A', onCover: '#FFFFFF',
  },
  warm: {
    label: '暖橘', font: '微軟正黑體',
    ink: '#40352D', strong: '#9C3D12', accent: '#1A3A5C', second: '#D08A50',
    tint: '#FBF2EA', muted: '#9C8B7C', line: '#E6D8CB',
    bg: '#FFFFFF', ground: '#FDF8F3', warn: '#2E7D6B', cover: '#9C3D12', onCover: '#FFFFFF',
  },
  night: {
    label: '深色', font: '微軟正黑體',
    // Every role inverts together. `cover` is the dark ground and `onCover` the
    // light type — the opposite of the light themes, stated rather than derived.
    ink: '#E8ECF1', strong: '#7FB2FF', accent: '#FF6B5B', second: '#4C9AFF',
    tint: '#1E2530', muted: '#98A3B3', line: '#2C3542',
    bg: '#141A22', ground: '#141A22', warn: '#F2C14E', cover: '#0D1219', onCover: '#F5F7FA',
  },
  linen: {
    // 文青。Warm and desaturated the whole way through: paper rather than white,
    // ink rather than black, a terracotta with the shout taken out of it.
    // Corporate red on a warm ground is what made the first attempt look like a
    // company report wearing a scarf.
    label: '亞麻', font: '微軟正黑體',
    ink: '#4A443C', strong: '#3E4A45', accent: '#A8705A', second: '#7D8B80',
    tint: '#F1EDE5', muted: '#9A9187', line: '#E2DAD0',
    bg: '#FFFDF9', ground: '#F7F2E9', warn: '#B8925A',
    cover: '#3E4A45', onCover: '#F7F3EC',
  },
  muted: {
    // 莫蘭迪。Every colour sits at roughly the same low saturation, which is what
    // makes the family read as deliberate rather than as washed out.
    label: '莫蘭迪', font: '微軟正黑體',
    ink: '#4C4F52', strong: '#5C6B73', accent: '#A98467', second: '#8FA3A8',
    tint: '#EDEFEE', muted: '#97A0A3', line: '#DDE2E1',
    bg: '#FCFCFB', ground: '#F4F5F4', warn: '#B0A08A',
    cover: '#4A5860', onCover: '#F4F6F5',
  },
  mono: {
    // 墨黑。No hue anywhere except the one accent. The most forgiving palette in
    // the set: nothing in it can clash with anything.
    label: '墨黑', font: '微軟正黑體',
    ink: '#2B2B2B', strong: '#1C1C1C', accent: '#8C6A4A', second: '#6B6B6B',
    tint: '#F0F0F0', muted: '#8E8E8E', line: '#DCDCDC',
    bg: '#FFFFFF', ground: '#F6F6F5', warn: '#A8A8A8',
    cover: '#1C1C1C', onCover: '#FAFAFA',
  },
};

/**
 * ── The second knob: how a slide is BUILT, not what colour it is ──
 *
 * THEMES above changes the paint. It does not change the look, and that turned
 * out to be the whole complaint: five palettes produced five recolourings of one
 * layout. Swap corporate for fresh and you get the same bar, the same 25pt bold
 * left-aligned title, the same short rule, the same tinted cards — in green.
 * Nobody calls that 「文青風」 or 「可愛風」; they call it the blue one and the
 * green one.
 *
 * A LOOK changes the things that actually carry style:
 *
 *   masthead   whether there is a bar at the top at all, and what kind
 *   rule       the mark under the title — short, full-width hairline, or none
 *   title      size, weight and alignment of the biggest word on the page
 *   ground     white page, tinted page, or a coloured one
 *   shape      square corners or round — this alone separates 企業 from 可愛
 *   card       filled, outlined, or nothing but type
 *   cover      how slide 1 is composed
 *
 * ── Why looks and themes are separate ──
 * Because the pairings are not fixed. 極簡 in warm orange and 極簡 in night are
 * both coherent; fusing look into palette would need thirty entries to say what
 * six and five say here. The cost is that the model has two knobs to keep steady
 * instead of one, which the prompt handles by saying: pick both once, then stop
 * touching them.
 *
 * ── About the fonts ──
 * Only families Office actually ships on a Chinese Windows are named. A missing
 * family does not error — PowerPoint silently substitutes — so an exotic choice
 * here would look right on my machine and wrong on the customer's, which is the
 * worst kind of styling bug because nothing reports it.
 */
const LOOKS = {
  business: {
    label: '企業', desc: '雙色頂條、短底線、方角卡片。給客戶、主管、對外提案。',
    theme: 'corporate',
    masthead: 'double', rule: 'short', ground: 'tint',
    titleSize: 25, titleBold: true, titleAlign: 'left', eyebrowCase: 'upper',
    shape: 'Rectangle', card: 'tint', cover: 'bleed', topPad: 0,
  },
  editorial: {
    label: '雜誌', desc: '沒有色條，標題大而不粗，下面一條細細的滿版線。像一篇報導的開頭。',
    font: '微軟正黑體 Light',
    theme: 'mono',
    masthead: 'none', rule: 'full', ground: 'white',
    titleSize: 30, titleBold: false, titleAlign: 'left', eyebrowCase: 'keep',
    shape: 'Rectangle', card: 'outline', cover: 'split', topPad: 10,
  },
  minimal: {
    label: '極簡', desc: '什麼裝飾都沒有，只有字和留白。內容本身夠強的時候最好看。',
    font: '微軟正黑體 Light',
    theme: 'mono',
    masthead: 'none', rule: 'none', ground: 'white',
    titleSize: 22, titleBold: false, titleAlign: 'left', eyebrowCase: 'keep',
    shape: 'Rectangle', card: 'plain', cover: 'plain', topPad: 18,
  },
  soft: {
    label: '柔和', desc: '圓角、粗一點的色條、標題置中。給內部分享、教學、比較輕鬆的場合。',
    theme: 'muted',
    masthead: 'round', rule: 'none', ground: 'tint',
    titleSize: 26, titleBold: true, titleAlign: 'center', eyebrowCase: 'keep',
    shape: 'RoundRectangle', card: 'tint', cover: 'band', topPad: 4,
  },
  bold: {
    label: '大字', desc: '標題佔掉上面一大塊，色塊當底。給只有一個重點、要講得很用力的場合。',
    theme: 'corporate',
    masthead: 'block', rule: 'none', ground: 'white',
    titleSize: 34, titleBold: true, titleAlign: 'left', eyebrowCase: 'upper',
    shape: 'Rectangle', card: 'tint', cover: 'bleed', topPad: 6,
  },
  tech: {
    label: '科技', desc: '細線、單色、深色底走向。給技術簡報、資料儀表板。',
    theme: 'night',
    masthead: 'hairline', rule: 'none', ground: 'strong',
    titleSize: 24, titleBold: true, titleAlign: 'left', eyebrowCase: 'upper',
    shape: 'Rectangle', card: 'outline', cover: 'bleed', topPad: 0,
  },
  literary: {
    label: '文青',
    desc: '暖色紙張滿版、明體標題、細框線。沒有企業色，沒有色塊。',
    // The serif is the single strongest signal and also the riskiest, so this is
    // a family Windows actually ships rather than a fashionable one: 新細明體 is
    // on every machine in Taiwan. Open the deck somewhere it is not and
    // PowerPoint substitutes silently — the slide still reads, it just loses the
    // mood. That is the failure mode we can live with.
    font: '微軟正黑體', titleFont: '新細明體',
    theme: 'linen',
    masthead: 'frame', rule: 'none', ground: 'tint',
    titleSize: 30, titleBold: false, titleAlign: 'left', eyebrowCase: 'keep',
    shape: 'Rectangle', card: 'plain', cover: 'paper', topPad: 14,
  },
};

const DEFAULT_LOOK = 'business';

const look = (name) => LOOKS[String(name || '').toLowerCase()] || LOOKS[DEFAULT_LOOK];

/**
 * The palette for a call, with the look's own choice as the default.
 *
 * ── The bug this fixes ──
 * Asking for 文青 used to change the skeleton and nothing else, so the slide came
 * back in corporate navy with a bright red spine down the side. A style is a
 * skeleton AND a palette; delivering half of one is not delivering a lighter
 * version of it, it is delivering something else.
 *
 * An explicit `theme` still wins, because 「文青風但用我們的企業色」 is a real
 * request and the two knobs are still independent — the look only supplies the
 * default nobody stated.
 */
const paletteFor = (args, lk) => theme(args.theme || lk.theme);

/**
 * The page ground a look asks for, resolved against the palette.
 *
 * `strong` is the one worth explaining: it paints the page in the palette's dark
 * colour, which on a light theme gives a dark slide with light type. That is the
 * 科技 look, and it means `tech` + `corporate` is a navy deck without needing a
 * sixth palette to say so.
 */
function groundOf(th, lk) {
  if (lk.ground === 'white') return th.bg;
  if (lk.ground === 'strong') return th.cover;
  return th.ground;
}

/**
 * How a panel is drawn, given the look — the one decision cards, KPI tiles and
 * callouts all have to make the same way.
 *
 *   tint     filled, faint border. The default, and the densest-looking.
 *   outline  no fill, a 1pt line in the accent. Lighter; lets the ground show.
 *   plain    neither. Only the coloured bar survives, which on 極簡 is the point.
 *
 * Spelled out here rather than in each renderer because three copies of this
 * drifting apart is exactly how a deck ends up with tinted cards on one slide
 * and outlined ones on the next.
 */
function skin(th, lk, accent, fallbackFill) {
  if (lk.card === 'plain') return { fill: null, line: false, color: accent };
  if (lk.card === 'outline') return { fill: th.bg, line: true, color: accent };
  return { fill: fallbackFill, line: true, color: th.line };
}

/** RoundRectangle if the look asks for it, otherwise whatever the caller wanted. */
const boxOf = (lk, base) => (lk.shape === 'RoundRectangle' ? 'RoundRectangle' : base);

/** Text colour that survives whatever groundOf just chose. */
const onGround = (th, lk, role) =>
  (lk.ground === 'strong' ? th.onCover : th[role]);

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
 * Slide masters and the layouts under them.
 *
 * ── Why the overview has to carry this ──
 * `ppt_add_slide` has taken `layout_id` + `master_id` since it was written, and
 * nothing anywhere told the model where those ids come from. So every new slide
 * was born on the default layout, and a deck built on the company master got
 * pages that did not match the ones already in it — the exact failure AC-11
 * names. The ids are only reachable here, so this is where they have to be said.
 *
 * Returns [] rather than throwing on an older build: the overview is the first
 * call of every conversation and it must not be the thing that fails.
 */
async function loadLayouts(ctx) {
  if (!hasApi('1.3')) return [];
  const masters = ctx.presentation.slideMasters;
  masters.load('items/id,items/name,items/layouts/items/id,items/layouts/items/name');
  try {
    await ctx.sync();
  } catch {
    return [];
  }
  return (masters.items || []).map((m) => ({
    id: String(m.id || ''),
    name: String(m.name || ''),
    layouts: ((m.layouts && m.layouts.items) || []).map((l) => ({
      id: String(l.id || ''),
      name: String(l.name || ''),
    })),
  }));
}

/* ── palette:start ── pure, and covered by test/ppt-overview.test.mjs ── */

/** How far the colour sample reaches. A deck states its palette in its first few pages. */
const THEME_SAMPLE_SLIDES = 12;
const THEME_SAMPLE_SHAPES = 240;
const THEME_REPORT = 5;

/** Pure white and pure black are in every deck ever made, so they say nothing. */
const NO_SIGNAL = new Set(['#FFFFFF', '#000000']);

/** '1F3864' and '#1f3864' are the same colour; the tally has to agree they are. */
function normaliseHex(raw) {
  const v = String(raw == null ? '' : raw).trim();
  if (/^#[0-9a-f]{6}$/i.test(v)) return v.toUpperCase();
  if (/^[0-9a-f]{6}$/i.test(v)) return ('#' + v).toUpperCase();
  return '';
}

/**
 * Most-used colours first, capped at THEME_REPORT.
 *
 * Split out of the sampling so the part with the judgement in it — what counts
 * as the same colour, what carries no signal, how many are worth reporting —
 * can be tested without an Office context to stand it up in.
 *
 * Ties keep first-seen order: Map iterates by insertion and sort is stable, so
 * two colours used once each come out in slide order rather than at random.
 */
function rankColours(values) {
  const tally = new Map();
  for (const raw of values) {
    const hex = normaliseHex(raw);
    if (!hex || NO_SIGNAL.has(hex)) continue;
    tally.set(hex, (tally.get(hex) || 0) + 1);
  }
  return [...tally.entries()]
    .sort((a, b) => b[1] - a[1])
    .slice(0, THEME_REPORT)
    .map(([hex]) => hex);
}

/** The masters/layouts block of the overview, or '' when there is nothing to say. */
function describeLayouts(masters) {
  if (!masters || !masters.length) return '';
  const lines = [];
  for (const m of masters) {
    if (!m.layouts || !m.layouts.length) continue;
    lines.push(`母片「${m.name || '(未命名)'}」 master_id=${m.id}`);
    for (const l of m.layouts) {
      lines.push(`  ・${l.name || '(未命名)'}　layout_id=${l.id}`);
    }
  }
  if (!lines.length) return '';
  // ── The boundary, and why it is stated here ──
  // The first version of this block said 「沿用簡報既有的版面就從這裡挑」 and
  // nothing else. The model read that as a general preference, started every
  // deck with ppt_add_slide + a native layout, and typed into the placeholders
  // — which bypasses every recipe in this module. The covers came out as plain
  // PowerPoint defaults: white ground, default font, no rule, orphaned words.
  // Nothing failed, and the output looked like the model had chosen it.
  //
  // Native layouts are for INHERITING a deck somebody else designed. The recipes
  // are for MAKING slides look designed. Offering the ids without saying which
  // situation they are for was offering a worse route and recommending it.
  return '\n可用的版面配置（layout_id 與 master_id 要成對給，只給一個會失敗）：\n' + lines.join('\n')
    + '\n【什麼時候用這些 id】只有在**這份簡報原本就有設計好的版面**、'
    + '而使用者要你「照現有格式繼續加」的時候。那種情況用 ppt_add_slide 帶上'
    + ' master_id + layout_id，新頁面才會跟原有的頁面一致。\n'
    + '【其他情況一律用 ppt_build_slide】空白簡報、從頭做一份、或使用者沒有'
    + '特別要求沿用格式——都用版型。原生版面配置給的是 PowerPoint 的預設長相'
    + '（白底、預設字型、空的預留位置），那不是設計，只是沒有設計。\n';
}

/** The palette block of the overview, or '' when the deck gave nothing away. */
function describeColours({ fills, texts }) {
  if (!fills.length && !texts.length) return '';
  const parts = [];
  if (fills.length) parts.push(`填色：${fills.join('、')}`);
  if (texts.length) parts.push(`文字色：${texts.join('、')}`);
  return '\n這份簡報已經在用的顏色（依出現次數排序，白與黑不計）：\n'
    + parts.join('\n')
    + '\n新增的內容請沿用這些顏色，不要另外套一組配色；'
    + '除非使用者明講要改，否則不要呼叫 ppt_apply_theme。\n';
}

/* ── palette:end ── */

/**
 * The colours this deck already uses, most-used first.
 *
 * ── There is no theme to read ──
 * PowerPoint's JS API exposes no `presentation.theme`: the six accent colours a
 * designer set on the master are not reachable from an add-in at all. What IS
 * reachable is every shape's own resolved fill and font colour — and a shape
 * filled from the theme palette reports the theme's hex. So the palette is
 * recovered by counting what is on the slides rather than by asking for it.
 *
 * Fills and text are counted separately because they answer different
 * questions. The model needs an accent to fill a bar with and an ink to set body
 * text in, and the most common fill is very rarely the right ink.
 *
 * Best-effort throughout. `shape.fill` throws for shapes that cannot carry one
 * and `foregroundColor` throws again when the fill is a gradient or a picture,
 * so every read is guarded on its own — one exotic shape must not cost the whole
 * sample, and a failed sample must not cost the overview.
 */
async function sampleColours(ctx, slides) {
  if (!hasApi('1.4')) return { fills: [], texts: [] };

  const fills = [];
  const fonts = [];
  let budget = THEME_SAMPLE_SHAPES;
  for (const slide of slides.slice(0, THEME_SAMPLE_SLIDES)) {
    for (const shape of (slide.shapes.items || [])) {
      if (budget <= 0) break;
      budget--;
      try {
        const f = shape.fill;
        f.load('type,foregroundColor');
        fills.push(f);
      } catch { /* this shape cannot carry a fill */ }
      try {
        const font = shape.textFrame.textRange.font;
        font.load('color');
        fonts.push(font);
      } catch { /* this shape holds no text */ }
    }
  }

  try {
    await ctx.sync();
  } catch {
    return { fills: [], texts: [] };
  }

  const read = (items, get) => items.map((item) => {
    try {
      return get(item);
    } catch {
      return '';
    }
  });

  return {
    // A non-solid fill has no single colour to report, so it is skipped rather
    // than guessed at from wherever the gradient happens to start.
    fills: rankColours(read(fills, (f) => (f.type === 'Solid' ? f.foregroundColor : ''))),
    texts: rankColours(read(fonts, (f) => f.color)),
  };
}

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
      return '這份簡報目前沒有投影片。用 ppt_add_slide 新增一張再開始。\n'
        + describeLayouts(await loadLayouts(ctx));
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

    // After the text pass, not before: both of these sync, and a deck with an
    // exotic shape in it can cost one of them. Ordering them last means the
    // slide list — the part nothing else works without — is already in hand.
    const masters = await loadLayouts(ctx);
    const palette = await sampleColours(ctx, slides);

    return `這份簡報有 ${slides.length} 張投影片。\n\n${lines.join('\n')}\n`
      + describeLayouts(masters)
      + describeColours(palette)
      + `\n版面尺寸以 ${GRID.W}×${GRID.H} 點（16:9）計算，內文範圍是左 ${GRID.margin} 到右 ${GRID.W - GRID.margin}。\n`
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

    // A blank slide the person will fill in themselves keeps its placeholders —
    // that is what they are for. Only the ones this module then draws over get
    // cleared, which happens in ppt_build_slide.
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

/**
 * Delete the empty placeholders a new slide is born with.
 *
 * ── The bug this fixes, which looked like a design problem ──
 * `slides.add()` applies the default layout, and that layout carries a title and
 * a subtitle placeholder. Everything this module draws goes ON TOP of them — so
 * every generated slide had 「按一下以新增標題」 sitting across its middle in 40pt
 * grey, with a dashed box round it.
 *
 * They do not print and they do not show in presentation mode, which is why this
 * is easy to miss when reasoning about the code. But the person judging the
 * output is looking at the EDITOR, and in the editor the deck looks broken. No
 * amount of better layout compensates for a placeholder prompt over the top of it.
 *
 * Only EMPTY ones go. A placeholder somebody has typed into is their content,
 * and this add-in does not delete content it did not create.
 */
async function clearPlaceholders(ctx, slide) {
  const shapes = slide.shapes;
  shapes.load('items/id,items/type');
  await ctx.sync();

  const candidates = (shapes.items || []).filter((s) => String(s.type) === 'Placeholder');
  if (!candidates.length) return 0;

  // textFrame throws for shapes that cannot hold text, so each is queued and
  // tolerated individually rather than in one bulk load.
  const frames = candidates.map((s) => {
    try {
      const f = s.textFrame;
      f.load('hasText');
      return f;
    } catch {
      return null;
    }
  });
  try {
    await ctx.sync();
  } catch {
    return 0;
  }

  let gone = 0;
  candidates.forEach((s, i) => {
    let empty = true;
    try {
      empty = !frames[i] || !frames[i].hasText;
    } catch {
      empty = true;
    }
    if (empty) { s.delete(); gone++; }
  });
  if (gone) await ctx.sync();
  return gone;
}

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

/**
 * Delete shapes — several at once, on purpose.
 *
 * ── Why this takes a list and not a number ──
 * It used to take one index. Clearing a slide before rebuilding it therefore
 * meant eight calls, and because deleting is destructive, eight separate
 * confirmation cards: 「刪除第 8 個物件」「刪除第 7 個物件」… Nobody reads the
 * eighth one. They click 允許 eight times, which trains them to click 允許
 * without looking — and that is the exact habit the confirmation gate exists to
 * prevent. A gate that fires eight times for one decision is worse than no gate,
 * because it is the same protection with the attention spent.
 *
 * One call, one decision, one card: 「刪除投影片 5 的 8 個物件」.
 *
 * ── Deleting from the back ──
 * The indices the model quotes come from ppt_read_slide, and they are positions
 * in a list. Delete number 3 and everything after it shifts down one, so a
 * forward loop deletes 3, then what used to be 5, then what used to be 8.
 * Descending order is the only order where every index still means what it meant
 * when it was read.
 */
async function deleteShape(args) {
  return PowerPoint.run(async (ctx) => {
    const { slide, index } = await targetSlide(ctx, args);
    const shapes = slide.shapes;
    shapes.load('items/id');
    await ctx.sync();
    const items = shapes.items || [];

    const asked = args.all === true
      ? items.map((_, i) => i + 1)
      : [...(Array.isArray(args.shapes) ? args.shapes : []),
        ...(args.shape === undefined ? [] : [args.shape])]
        .map((n) => num(n, 0)).filter(Boolean);

    if (!asked.length) throw new Error('要刪哪些？給 shape（單一編號）、shapes（編號陣列）或 all: true。');

    // Everything is checked before anything is deleted. A half-applied delete is
    // the worst outcome here: the undo cannot bring shapes back, so a call that
    // removes four things and then fails on the fifth has destroyed four things
    // for nothing.
    const uniq = [...new Set(asked)].sort((a, b) => a - b);
    const bad = uniq.filter((n) => n < 1 || n > items.length);
    if (bad.length) {
      throw new Error(`投影片 ${index} 只有 ${items.length} 個物件，`
        + `但你要刪第 ${bad.join('、')} 個。先用 ppt_read_slide 看一次目前的編號。`);
    }

    uniq.slice().reverse().forEach((n) => items[n - 1].delete());
    await ctx.sync();
    return uniq.length === 1
      ? `已刪除投影片 ${index} 的第 ${uniq[0]} 個物件。`
      : `已刪除投影片 ${index} 的 ${uniq.length} 個物件（第 ${uniq.join('、')} 個）。`;
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
  /**
   * A second, shorter bar in the accent colour, tucked under the first.
   *
   * It costs one shape and it is the difference between a line and a masthead.
   * 288 is 30% of the width — short enough to read as a deliberate marker rather
   * than a second rule that failed to reach the edge.
   */
  strip2: { left: 0, top: 5, width: 288, height: 2 },
  /**
   * The full-bleed dark band along the bottom.
   *
   * ── Why a band and not a line of grey text ──
   * A slide with white above and white below has the content floating in it. The
   * reference deck closes every content page with a dark strip, and the effect is
   * structural rather than decorative: it gives the page a floor, so the body has
   * something to sit ON. It also absorbs whatever whitespace the content did not
   * use, which is the actual cure for 「看起來很空」.
   */
  footband: { left: 0, top: 494, width: GRID.W, height: 46 },
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
  const lk = look(args.look);
  const th = paletteFor(args, lk);
  const face = lk.font || th.font;
  // 文青 sets its titles in a Ming face and its body in a sans. That contrast is
  // most of the mood, and it costs one extra field.
  const display = lk.titleFont || face;

  return PowerPoint.run(async (ctx) => {
    // A recipe fills a slide, so it defaults to making a new one rather than
    // piling a second layout on top of whatever is already there.
    if (args.slide === undefined && args.append !== false && hasApi('1.3')) {
      ctx.presentation.slides.add();
      await ctx.sync();
    }
    const { slide, index } = await targetSlide(ctx, args);
    // Before anything is drawn: the layout's empty prompts would otherwise sit
    // on top of everything below.
    await clearPlaceholders(ctx, slide);
    const put = { text: [], shape: [] };

    const text = (content, geo, style) => {
      const s = slide.shapes.addTextBox(String(content), geo);
      s.name = stamp('build');
      styleText(s, { font: face, wrap: true, ...style });
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
    /**
     * How much of the frame this recipe wants.
     *
     * ── The bug this replaces ──
     * The chrome used to go on every slide that was not the cover. But `section`
     * and `statement` are WHOLE-PAGE compositions that draw their own title —
     * so a section slide printed 「03 代理式 AI」 in the header and then again in
     * the middle of the page, and a statement slide printed 「逾 80%」 as a 25pt
     * heading above the same words at 84pt. Twice is not emphasis. It reads as a
     * rendering fault, and it is the reason those slides looked unfinished while
     * the ordinary ones looked fine.
     *
     * 'marks' is the middle case and exists for `statement` specifically: the
     * eyebrow and the page number are NAVIGATION — they answer 「我們講到哪了」 —
     * and a slide carrying one huge number needs that more than any other, since
     * the number itself says nothing about where it sits in the argument. What it
     * does not need is a title, because the number is the title.
     */
    const frame = args.chrome === false ? 'none'
      : (kind === 'title' || kind === 'section') ? 'none'
        : kind === 'statement' ? 'marks'
          : 'full';
    const dressed = frame !== 'none';
    if (dressed) {
      // The page ground goes down FIRST so everything else lands on top of it.
      //
      // Off-white rather than white, and this is not a small difference: a pure
      // white slide has no page, so a white card on it has no edge and the
      // content reads as typed straight onto the projector. Two percent of grey
      // gives the cards something to be lighter than, and that is what makes a
      // layout look built rather than filled in.
      band({ left: 0, top: 0, width: GRID.W, height: GRID.H }, groundOf(th, lk));

      // The masthead. Each variant is a different answer to the same question —
      // how does the audience know a new slide started — and swapping between
      // them changes the deck's character more than any colour does.
      if (lk.masthead === 'double') {
        band(CHROME.strip, th.accent);
        band(CHROME.strip2, th.strong);
      } else if (lk.masthead === 'round') {
        // Inset and rounded: a bar that does not touch the edges reads friendly
        // rather than official. That is the whole of 柔和.
        const r = slide.shapes.addGeometricShape('RoundRectangle',
          { left: 38, top: 20, width: 96, height: 8 });
        r.name = stamp('band');
        r.fill.setSolidColor(th.accent);
        r.lineFormat.visible = false;
        put.shape.push(r);
      } else if (lk.masthead === 'block') {
        // A solid slab behind the eyebrow — the title then sits on white below
        // it, and the contrast is what makes 大字 loud without a bigger font.
        band({ left: 0, top: 0, width: 300, height: 34 }, th.strong);
      } else if (lk.masthead === 'hairline') {
        band({ left: 0, top: 0, width: GRID.W, height: 1 }, th.line);
      } else if (lk.masthead === 'frame') {
        // A hairline box inset from all four edges. It does two things at once:
        // it is the only ornament on the page, and it turns the margin into a
        // deliberate border instead of leftover space — which is the difference
        // between 「留白」 and 「很空」.
        const f = slide.shapes.addGeometricShape('Rectangle',
          { left: 24, top: 24, width: GRID.W - 48, height: GRID.H - 48 });
        f.name = stamp('frame');
        f.fill.clear();
        f.lineFormat.visible = true;
        f.lineFormat.color = th.line;
        f.lineFormat.weight = 1;
        put.shape.push(f);
      }

      const eyeY = lk.masthead === 'block' ? 8 : CHROME.eyebrow.top + lk.topPad;
      if (args.eyebrow) {
        const label = lk.eyebrowCase === 'upper'
          ? String(args.eyebrow).toUpperCase() : String(args.eyebrow);
        text(label, { ...CHROME.eyebrow, left: lk.masthead === 'block' ? 16 : CHROME.eyebrow.left, top: eyeY },
          {
            size: 11,
            color: lk.masthead === 'block' ? th.onCover : onGround(th, lk, 'muted'),
            valign: 'middle',
            align: lk.titleAlign === 'center' ? 'center' : 'left',
          });
      }
      if (title && frame === 'full') {
        text(title, { ...CHROME.title, top: CHROME.title.top + lk.topPad },
          {
            font: display,
            size: lk.titleSize, bold: lk.titleBold,
            color: onGround(th, lk, 'strong'),
            align: lk.titleAlign, valign: 'middle',
          });
        if (lk.rule === 'short') {
          band(lk.titleAlign === 'center'
            ? { ...CHROME.rule, left: (GRID.W - CHROME.rule.width) / 2, top: CHROME.rule.top + lk.topPad }
            : { ...CHROME.rule, top: CHROME.rule.top + lk.topPad }, th.accent);
        } else if (lk.rule === 'full') {
          // Full width and one point high. A short rule underlines the title; a
          // hairline across the page separates the header from the body, which
          // is the move a magazine makes and a slide deck usually does not.
          band({ left: CHROME.contentLeft, top: CHROME.rule.top + lk.topPad + 4,
            width: CHROME.contentWidth, height: 1 }, th.line);
        }
      }
      // With a footer band the page number lives down there instead — printing it
      // in both corners is the kind of duplicate nobody notices until it is on a
      // projector.
      if (args.page && !args.footband) {
        text(String(args.page), { left: 800, top: 22 + lk.topPad, width: 116, height: 19 },
          { size: 11, color: onGround(th, lk, 'muted'), align: 'right', valign: 'middle' });
      }
    }
    const CL = CHROME.contentLeft;
    const CW = CHROME.contentWidth;

    /**
     * The composer: one call, one whole page, no holes.
     *
     * ── What this replaces ──
     * Building a dense slide used to be three calls — build_slide, then
     * add_kpi, then add_callout — and each one chose its own `top`. No single
     * call could see the others, so the leftovers landed wherever the model's
     * arithmetic happened to leave them: a void under the title and another
     * above the footer. 「一頁 PPT 這樣也太少東西了」 was that, exactly.
     *
     * Here the band is divided before anything is drawn. The fixed-height
     * blocks take their share off the top and the bottom, and the recipe's own
     * content gets ALL of what is left — so the page is full by construction
     * rather than by the model guessing well.
     */
    const kit = { th, lk, face };
    const pen = penFor(slide, (sh) => put.shape.push(sh), (sh) => put.text.push(sh));
    let bandTop = CHROME.contentTop + lk.topPad;
    let bandBottom = bandBottomFor(args);

    const kpis = (Array.isArray(args.kpi) ? args.kpi : []).filter(Boolean).slice(0, 5);
    if (kpis.length && dressed) {
      drawKpiRow(pen, kit, kpis,
        { left: CL, top: bandTop, width: CW, height: KPI_H });
      bandTop += KPI_H + STACK_GAP;
    }

    const callout = typeof args.callout === 'string'
      ? { text: args.callout }
      : (args.callout && args.callout.text ? args.callout : null);
    if (callout && dressed) {
      drawCallout(pen, kit, callout,
        { left: CL, top: bandBottom - CALLOUT_H, width: CW, height: CALLOUT_H });
      bandBottom -= CALLOUT_H + STACK_GAP;
    }

    // What the recipes lay out against. CT stays the name every branch below
    // already uses; it is simply no longer a constant.
    const CT = bandTop;
    const BODY_H = Math.max(120, bandBottom - bandTop);

    if (kind === 'title') {
      // ── The cover, composed four different ways ──
      //
      // Slide 1 sets the expectation for everything after it, so this is where a
      // look has to be most visible. All four keep the same promise — the title
      // is the largest thing on the page and it has enough contrast to read from
      // the back — and differ in everything else.
      const sub = body.join('　·　');

      if (lk.cover === 'split') {
        // A magazine cover: white page, one narrow colour edge, type set left and
        // ragged right. Nothing is centred, which is exactly why it reads as
        // designed rather than as a template.
        band({ left: 0, top: 0, width: GRID.W, height: GRID.H }, th.bg);
        band({ left: 0, top: 0, width: 14, height: GRID.H }, th.accent);
        if (args.eyebrow) {
          text(String(args.eyebrow), { left: 96, top: 150, width: 700, height: 24 },
            { size: TYPE.small, color: th.accent });
        }
        text(title, { left: 96, top: 182, width: 740, height: 110 },
          { size: 42, bold: false, color: th.strong, valign: 'middle' });
        band({ left: 96, top: 306, width: 740, height: 1 }, th.line);
        if (sub) {
          text(sub, { left: 96, top: 320, width: 740, height: 40 },
            { size: TYPE.heading, color: th.muted });
        }
        if (args.footer) {
          text(String(args.footer), { left: 96, top: 456, width: 740, height: 24 },
            { size: TYPE.small, color: th.muted });
        }
      } else if (lk.cover === 'band') {
        // A rounded slab floating on a tinted page. Inset on all four sides, so
        // the page frames it — the friendliest of the four and the only one that
        // stays centred.
        band({ left: 0, top: 0, width: GRID.W, height: GRID.H }, th.ground);
        const slab = slide.shapes.addGeometricShape('RoundRectangle',
          { left: 64, top: 132, width: GRID.W - 128, height: 276 });
        slab.name = stamp('band');
        slab.fill.setSolidColor(th.cover);
        slab.lineFormat.visible = false;
        put.shape.push(slab);
        text(title, { left: 104, top: 188, width: GRID.W - 208, height: 92 },
          { size: 38, bold: true, color: th.onCover, align: 'center', valign: 'middle' });
        band({ left: (GRID.W - 96) / 2, top: 292, width: 96, height: 3 }, th.accent);
        if (sub) {
          text(sub, { left: 104, top: 310, width: GRID.W - 208, height: 40 },
            { size: TYPE.heading, color: th.onCover, align: 'center' });
        }
        if (args.footer) {
          text(String(args.footer), { left: 64, top: 432, width: GRID.W - 128, height: 24 },
            { size: TYPE.small, color: th.muted, align: 'center' });
        }
      } else if (lk.cover === 'paper') {
        // Full-bleed warm paper with a hairline frame. Nothing is white, which
        // is the point: 「一頁滿版」 does not have to mean a photograph, it can
        // just mean the page has a colour and the colour reaches the edges.
        band({ left: 0, top: 0, width: GRID.W, height: GRID.H }, th.ground);
        const f = slide.shapes.addGeometricShape('Rectangle',
          { left: 24, top: 24, width: GRID.W - 48, height: GRID.H - 48 });
        f.name = stamp('frame');
        f.fill.clear();
        f.lineFormat.visible = true;
        f.lineFormat.color = th.line;
        f.lineFormat.weight = 1;
        put.shape.push(f);
        if (args.eyebrow) {
          text(String(args.eyebrow), { left: 88, top: 176, width: 700, height: 22 },
            { size: TYPE.small, color: th.accent });
        }
        text(title, { left: 88, top: 208, width: 784, height: 104 },
          { font: display, size: 40, bold: false, color: th.strong, valign: 'middle' });
        band({ left: 88, top: 322, width: 64, height: 2 }, th.accent);
        if (sub) {
          text(sub, { left: 88, top: 338, width: 784, height: 40 },
            { font: display, size: TYPE.heading, color: th.ink });
        }
        if (args.footer) {
          text(String(args.footer), { left: 88, top: 436, width: 784, height: 24 },
            { size: TYPE.caption, color: th.muted });
        }
      } else if (lk.cover === 'plain') {
        // Type and nothing else. The hardest one to get wrong and the easiest to
        // dismiss — which is the point: when the content is strong, decoration
        // is what makes it look weaker.
        band({ left: 0, top: 0, width: GRID.W, height: GRID.H }, th.bg);
        if (args.eyebrow) {
          text(String(args.eyebrow), { left: 96, top: 196, width: 700, height: 22 },
            { size: TYPE.small, color: th.muted });
        }
        text(title, { left: 96, top: 224, width: 768, height: 96 },
          { size: 40, bold: false, color: th.strong, valign: 'middle' });
        if (sub) {
          text(sub, { left: 96, top: 326, width: 768, height: 36 },
            { size: TYPE.heading, color: th.muted });
        }
        if (args.footer) {
          text(String(args.footer), { left: 96, top: 456, width: 768, height: 24 },
            { size: TYPE.small, color: th.muted });
        }
      } else {
        // 'bleed' — the full-colour cover. `cover`/`onCover` rather than
        // accent/onAccent: see THEMES for the bug that distinction prevents.
        //
        // ── Why this stopped being centred ──
        // Everything used to sit in a band across the middle, which left the top
        // quarter and the bottom quarter empty and made the slide read as a
        // PowerPoint default with the colours changed. Ranged left with a mark at
        // the top and a rule at the foot, the same words occupy the whole page —
        // the type has not grown, it has been given somewhere to sit.
        band({ left: 0, top: 0, width: GRID.W, height: GRID.H }, th.cover);
        band({ left: 0, top: 0, width: GRID.W, height: 6 }, th.accent);
        band({ left: 0, top: 6, width: 288, height: 2 }, th.onCover);
        if (args.eyebrow) {
          text(String(args.eyebrow).toUpperCase(), { left: 80, top: 116, width: 700, height: 24 },
            { size: TYPE.small, color: th.accent });
        }
        text(title, { left: 80, top: 168, width: 800, height: 128 },
          { font: display, size: TYPE.display, bold: true, color: th.onCover, valign: 'middle' });
        // A short rule between title and subtitle. The same idea as the one under
        // a slide title: it belongs to the words above it.
        band({ left: 80, top: 306, width: 120, height: 3 }, th.accent);
        if (sub) {
          text(sub, { left: 80, top: 326, width: 780, height: 60 },
            { size: TYPE.heading, color: th.onCover });
        }
        // The foot. A hairline the width of the type block, then the metadata —
        // this is what stops the bottom third being a hole.
        band({ left: 80, top: 452, width: 800, height: 1 }, th.accent);
        if (args.footer) {
          text(String(args.footer), { left: 80, top: 464, width: 640, height: 24 },
            { size: TYPE.small, color: th.onCover });
        }
        if (args.page) {
          text(String(args.page), { left: 740, top: 464, width: 140, height: 24 },
            { size: TYPE.small, color: th.onCover, align: 'right' });
        }
      }
    } else if (kind === 'section') {
      /**
       * A divider, and the whole page is the divider.
       *
       * ── Why this is now full-bleed ──
       * It used to be a tinted band across the middle of an otherwise ordinary
       * slide, which made it look like a content page whose content had not
       * arrived — and because the frame was drawn on top of it as well, the
       * section number appeared twice. A deck's section slides have one job: to
       * be unmistakably NOT a content slide, so the audience registers a break
       * without reading anything. Flooding the page does that in a way a band
       * never can.
       *
       * The 文青 look keeps its paper instead of going dark: a black slab in the
       * middle of a warm-paper deck is not a pause, it is a different deck.
       */
      const paper = lk.cover === 'paper';
      const ground = paper ? th.ground : th.cover;
      const ink = paper ? th.strong : th.onCover;
      band({ left: 0, top: 0, width: GRID.W, height: GRID.H }, ground);
      band({ left: 0, top: 0, width: GRID.W, height: 6 }, th.accent);
      if (args.eyebrow) {
        // The number is what people actually navigate by, so it is the largest
        // thing on the page — larger than the section's own name.
        text(String(args.eyebrow), { left: 80, top: 186, width: 220, height: 96 },
          { font: display, size: 76, bold: true, color: th.accent, valign: 'middle' });
      }
      const sx = args.eyebrow ? 220 : 80;
      text(title, { left: sx, top: 196, width: 880 - sx, height: 60 },
        { font: display, size: 36, bold: !paper, color: ink, valign: 'middle' });
      band({ left: sx, top: 264, width: 96, height: 3 }, th.accent);
      if (body.length) {
        text(body[0], { left: sx, top: 282, width: 880 - sx, height: 48 },
          { size: TYPE.heading, color: paper ? th.muted : th.onCover });
      }
      if (args.footer) {
        text(String(args.footer), { left: 80, top: 466, width: 640, height: 24 },
          { size: TYPE.caption, color: paper ? th.muted : th.onCover });
      }
      if (args.page) {
        text(String(args.page), { left: 740, top: 466, width: 140, height: 24 },
          { size: TYPE.caption, color: paper ? th.muted : th.onCover, align: 'right' });
      }
    } else if (kind === 'statement') {
      /**
       * One number, big — and the one thing it must never be is 「不明所以」.
       *
       * A slide that says 逾 80% and nothing else is a quiz. The reading order
       * has to be: what this is about (the eyebrow, drawn by the frame above) →
       * the figure → the sentence that makes it mean something → where it came
       * from. All four, or the number is decoration.
       *
       * The tinted panel is doing structural work rather than decorative: it
       * gives the figure a floor and a ceiling, so the slide is a composition
       * with something in the middle instead of a number floating in a void.
       */
      band({ left: 0, top: 150, width: GRID.W, height: 252 }, th.tint);
      band({ left: 0, top: 150, width: GRID.W, height: 3 }, th.accent);
      text(title, { left: CL, top: 176, width: CW, height: 116 },
        { font: display, size: 84, bold: true, color: th.strong, align: 'center', valign: 'middle' });
      band({ left: (GRID.W - 120) / 2, top: 300, width: 120, height: 4 }, th.accent);
      if (body.length) {
        // The sentence is the half that stops it being a quiz, so it is set at
        // reading size in the ink — not as a grey caption under a big number.
        text(body.join('\n'), { left: 120, top: 318, width: GRID.W - 240, height: 70 },
          { size: TYPE.heading, color: th.ink, align: 'center' });
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
      const colH = BODY_H - 8;
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
      // ── bullets ──
      // The ordinary content slide, and the one that decides whether the deck
      // looks made or generated.
      //
      // The first version put the whole list in one text box. Three lines of
      // text in the top third of the slide and nothing below — technically the
      // content, visually a draft. Counting the reference deck: 17 to 47 shapes
      // per slide against the six that produced.
      //
      // The fix is not to give the model a richer tool and hope it reaches for
      // one. It did not: told to use ppt_build_slide, it used ppt_build_slide,
      // and got back a text box. Density has to be what happens by DEFAULT.
      //
      // So a short list renders as cards — a tinted panel and a coloured spine
      // per point, sized to fill the content area rather than to fit the words.
      // It costs nothing: the extra shapes go into the same batch and the same
      // single sync.
      const items = body.filter((l) => String(l).trim());
      const asCards = args.style !== 'plain' && items.length >= 2 && items.length <= 6;

      if (asCards) {
        const AREA_TOP = CT;
        const AREA_H = BODY_H;
        const gap = items.length > 4 ? 12 : 16;
        const ch = (AREA_H - gap * (items.length - 1)) / items.length;
        const cycle = [th.strong, th.second, th.accent];

        items.forEach((raw, i) => {
          const y = AREA_TOP + (ch + gap) * i;
          const tone = cycle[i % cycle.length];
          band({ left: CL, top: y, width: CW, height: ch }, th.tint);
          band({ left: CL, top: y, width: 5, height: ch }, tone);

          // 「寫程式：把麻煩的事變成自動的事」 is how people and models both write
          // a point — a label, then the substance. Splitting on the colon turns
          // that into a heading and a body rather than one flat line, which is
          // most of what makes the card look composed instead of padded.
          const s = String(raw).trim();
          const cut = /^([^：:]{1,14})[：:]\s*(.+)$/.exec(s);
          const head = cut ? cut[1] : s;
          const rest = cut ? cut[2] : '';

          // The index, in the spine's colour. Small — it orders the points, it
          // is not the point.
          text(String(i + 1).padStart(2, '0'),
            { left: CL + 22, top: y + (rest ? 14 : ch / 2 - 16), width: 46, height: 32 },
            { size: 20, bold: true, color: tone, valign: 'middle' });

          text(head,
            { left: CL + 74, top: y + (rest ? 12 : ch / 2 - 17), width: CW - 100, height: rest ? 30 : 34 },
            { size: rest ? 17 : 18, bold: true, color: th.ink, valign: 'middle' });

          if (rest) {
            text(rest, { left: CL + 74, top: y + 44, width: CW - 100, height: ch - 54 },
              { size: 13.5, color: th.muted, valign: 'top' });
          }
        });
      } else {
        // Seven or more points is a slide that should have been two, but the
        // model has already decided. Cards at that count are 40pt tall and read
        // worse than a plain list, so the plain list is the honest fallback.
        //
        // ── Centred, not top-aligned ──
        // This path was the one the card layout above was built to escape, and
        // it kept the original defect: eight short lines pinned to the top of a
        // 348pt band leave the bottom HALF of the slide empty, which reads as an
        // unfinished draft. Nothing is missing — the content simply sits in the
        // wrong place. The prompt tells the model 「稀疏內容先收攏再置中」 and
        // this is the same rule applied where the geometry is actually decided,
        // because a rule the renderer contradicts is a rule nobody follows.
        //
        // Two columns once the list is long: eight lines down one side of a
        // 16:9 slide wastes half its width, and the eye has further to travel
        // than a column break would cost.
        const wide = items.length >= 8;
        if (wide) {
          const half = Math.ceil(items.length / 2);
          const colW = (CW - GRID.gutter) / 2;
          text(bullet(items.slice(0, half)),
            { left: CL, top: CT, width: colW, height: BODY_H },
            { size: TYPE.body, color: th.ink, valign: 'middle' });
          text(bullet(items.slice(half)),
            { left: CL + colW + GRID.gutter, top: CT, width: colW, height: BODY_H },
            { size: TYPE.body, color: th.ink, valign: 'middle' });
        } else {
          text(bullet(items), { left: CL, top: CT, width: CW, height: BODY_H },
            { size: TYPE.body, color: th.ink, valign: 'middle' });
        }
      }
    }

    if (args.footband && kind !== 'title') {
      // The heavy close. Text is optional — the band does its structural job
      // whether or not anything is written in it.
      band(CHROME.footband, th.strong);
      band({ left: 0, top: CHROME.footband.top, width: 288, height: 3 }, th.accent);
      if (args.footer) {
        text(String(args.footer), { left: CL, top: CHROME.footband.top + 12, width: CW - 120, height: 22 },
          { size: TYPE.small, color: '#FFFFFF', valign: 'middle' });
      }
      if (args.page) {
        text(String(args.page), { left: GRID.W - 140, top: CHROME.footband.top + 12, width: 100, height: 22 },
          { size: TYPE.small, color: '#FFFFFF', align: 'right', valign: 'middle' });
      }
    } else if (args.footer && kind !== 'title') {
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
/**
 * Where the body of a slide lives, top to bottom.
 *
 * ── Why this exists as a number and not as a habit ──
 * The complaint was 「一頁 PPT 這樣也太少東西了」, and the cause was not that the
 * elements were too small. It was that nothing owned the page's HEIGHT. The
 * model placed a card row at its own guess of `top`, then a callout at another
 * guess, and the leftovers became two holes — one under the title and one above
 * the footer — that no single tool could see, because each one only knew about
 * itself.
 *
 * So the band is stated once and the composer below divides it. 472 rather than
 * 486 when there is no footer band: a hairline footer needs the room, and a
 * block that ends exactly on the rule reads as an overflow.
 */
const bandBottomFor = (args) => (args.footband ? 480 : args.footer ? 472 : 490);

/** Fixed heights, so the composer can divide what is left. */
const KPI_H = 101;
const CALLOUT_H = 52;
/** Between stacked blocks. Smaller than the page margin on purpose — these are
 *  parts of one composition, not separate things that happen to share a slide. */
const STACK_GAP = 18;

/**
 * A drawing surface both the standalone tools and ppt_build_slide can write to.
 *
 * The two used to draw the same element in two places: addKpi built tiles with
 * `slide.shapes.add*` and an array, while buildSlide had its own `text()`/`band()`
 * helpers. Duplicating the geometry would mean a KPI row inside a built slide
 * slowly drifting away from a KPI row added to one — same name, two looks, and
 * no test could catch it because both would be self-consistent.
 */
function penFor(slide, onShape, onText) {
  return {
    rect(geo, opts = {}) {
      const sh = slide.shapes.addGeometricShape(opts.shape || 'Rectangle', geo);
      sh.name = stamp(opts.tag || 'box');
      if (opts.fill) sh.fill.setSolidColor(opts.fill); else sh.fill.clear();
      sh.lineFormat.visible = !!opts.line;
      if (opts.line) {
        sh.lineFormat.color = opts.line;
        sh.lineFormat.weight = opts.weight || 1;
      }
      onShape(sh);
      return sh;
    },
    label(content, geo, style = {}) {
      const sh = slide.shapes.addTextBox(String(content), geo);
      sh.name = stamp(style.tag || 'txt');
      styleText(sh, { wrap: true, ...style });
      sh.fill.clear();
      sh.lineFormat.visible = false;
      onText(sh);
      return sh;
    },
  };
}

/**
 * The KPI row, drawn into a box.
 *
 * `box.height` is honoured rather than fixed, so the same code makes the 101pt
 * row the standalone tool asks for and whatever height the slide composer has
 * left over. The type positions are proportional to it for the same reason.
 */
function drawKpiRow(pen, kit, items, box, opts = {}) {
  const { th, lk, face } = kit;
  const gap = num(opts.gap, 15);
  const w = (box.width - gap * (items.length - 1)) / items.length;
  const cycle = [th.strong, th.second, th.accent, th.warn || th.second, th.muted];

  items.forEach((item, i) => {
    const x = box.left + (w + gap) * i;
    const tone = colour(item.color, cycle[i % cycle.length]);
    const sk = skin(th, lk, tone, colour(opts.fill, th.bg));

    pen.rect({ left: x, top: box.top, width: w, height: box.height },
      { shape: boxOf(lk, 'Rectangle'), tag: 'kpi', fill: sk.fill, line: sk.line && sk.color });
    // On the panel's top edge, not inside it: a border would draw all four sides.
    pen.rect({ left: x, top: box.top, width: w, height: 4 },
      { shape: boxOf(lk, 'Rectangle'), tag: 'kpibar', fill: tone });

    const pad = 14;
    const tx = x + pad;
    const tw = w - pad * 2;
    // The figure takes the room the label and note do not, so a tall row grows
    // the number rather than leaving a gap under it.
    const labelH = item.label ? 21 : 0;
    const noteH = item.note ? 16 : 0;
    const valueH = Math.max(36, box.height - 22 - labelH - noteH);
    pen.label(item.value ?? '', { left: tx, top: box.top + 11, width: tw, height: valueH },
      { font: face, size: 30, bold: true, color: tone, valign: 'middle', tag: 'kpival' });
    let y = box.top + 11 + valueH;
    if (item.label) {
      pen.label(item.label, { left: tx, top: y, width: tw, height: labelH },
        { font: face, size: 12, bold: true, color: th.strong, valign: 'middle', tag: 'kpilab' });
      y += labelH;
    }
    if (item.note) {
      pen.label(item.note, { left: tx, top: y, width: tw, height: noteH },
        { font: face, size: 9, color: th.muted, valign: 'middle', tag: 'kpinote' });
    }
  });
}

/** The insight strip, drawn into a box. One paragraph — see ppt_add_callout. */
function drawCallout(pen, kit, spec, box) {
  const { th, lk, face } = kit;
  const tone = colour(spec.color, th.strong);
  const sk = skin(th, lk, tone, colour(spec.fill, th.tint));

  pen.rect(box, {
    shape: boxOf(lk, 'Rectangle'), tag: 'callout',
    // The conclusion keeps its accent border on every skin — a hairline would
    // make it read as one more card.
    fill: sk.fill, line: tone,
  });
  pen.rect({ left: box.left, top: box.top, width: 4, height: box.height },
    { tag: 'calloutbar', fill: tone });

  let tx = box.left + 16;
  let tw = box.width - 32;
  if (spec.label) {
    pen.label(spec.label, { left: tx, top: box.top + 4, width: 96, height: box.height - 8 },
      { font: face, size: 11, bold: true, color: tone, valign: 'middle', tag: 'calloutlab' });
    tx += 100;
    tw -= 100;
  }
  pen.label(spec.text, { left: tx, top: box.top + 4, width: tw, height: box.height - 8 },
    { font: face, size: 12, color: th.ink, valign: 'middle', tag: 'calloutbody' });
}

/**
 * A row of metric tiles — the element that makes a slide look finished.
 *
 * ── Why this is its own tool and not a shape of ppt_add_cards ──
 * A card carries a sentence; a KPI tile carries a NUMBER, and the two want
 * opposite proportions. A card is wide and short so the text has a line to run
 * along; a tile is tall enough to stack a large figure over a label over a
 * footnote, and narrow enough that four sit side by side. Squeezing one into the
 * other's geometry is how you get four cards with tiny numbers in them.
 *
 * ── The numbers come from a deck that works ──
 * Measured off the reference: tiles 212×101 at y=133, 15pt apart across the
 * 884pt content width; a 4pt bar across the TOP of each in its own colour; the
 * figure at 30pt in that same colour; the label at 12 in the ink; the footnote
 * at 9 in the muted grey. Those proportions are why the row reads as one band
 * rather than as four boxes, so they are defaults rather than suggestions.
 *
 * Each tile takes its own colour by default. A row of four in one colour is a
 * table; a row of four in four colours is a dashboard, and the difference costs
 * nothing.
 */
async function addKpi(args) {
  if (!hasApi('1.4')) throw new Error('數據卡需要 PowerPointApi 1.4，這個 Office 版本太舊。');
  const items = (Array.isArray(args.items) ? args.items : []).filter(Boolean);
  if (!items.length) throw new Error('items 是必要的，至少要一個。');
  if (items.length > 5) {
    throw new Error('一排最多 5 個數據卡——再多每個就窄到放不下數字了。挑最重要的幾個。');
  }

  const lk = look(args.look);
  const th = paletteFor(args, lk);
  const kit = { th, lk, face: lk.font || th.font };
  const gap = num(args.gap, 15);
  const box = {
    left: num(args.left, CHROME.contentLeft),
    top: num(args.top, CHROME.contentTop),
    width: num(args.width, CHROME.contentWidth),
    height: num(args.height, KPI_H),
  };

  return PowerPoint.run(async (ctx) => {
    const { slide, index } = await targetSlide(ctx, args);
    const made_ = [];
    const keep = (sh) => made_.push(sh);
    // Drawn by the same routine ppt_build_slide uses. Two copies of this
    // geometry would drift — same tool name, two looks, and no test could catch
    // it because each copy would be self-consistent.
    drawKpiRow(penFor(slide, keep, keep), kit, items, box, { gap, fill: args.fill });

    made_.forEach((s) => s.load('id'));
    await ctx.sync();
    made_.forEach((s) => remember(slide.id, s.id));
    const w = (box.width - gap * (items.length - 1)) / items.length;
    return `已在投影片 ${index} 放上 ${items.length} 個數據卡，`
      + `${Math.round(w)}×${box.height} 點，從 (${box.left},${box.top}) 開始。`;
  });
}

/**
 * The insight strip — one sentence, tinted, usually at the foot of the slide.
 *
 * ── What it is for ──
 * A slide full of figures says what happened. This says what it MEANS, and the
 * reference deck puts one on nearly every content page for that reason. It also
 * does something structural: a tinted band across the bottom gives the page a
 * third layer, so the content stops floating in the middle of the slide.
 *
 * Deliberately one paragraph. A callout that grows into a list is the slide's
 * content wearing a different colour, and then nothing on the page is the
 * conclusion.
 */
async function addCallout(args) {
  if (!hasApi('1.4')) throw new Error('洞察框需要 PowerPointApi 1.4，這個 Office 版本太舊。');
  const body = String(args.text || '').trim();
  if (!body) throw new Error('text 是必要的（要講的那句結論）。');

  const lk = look(args.look);
  const th = paletteFor(args, lk);
  const kit = { th, lk, face: lk.font || th.font };
  const h = num(args.height, CALLOUT_H);
  const box = {
    left: num(args.left, CHROME.contentLeft),
    width: num(args.width, CHROME.contentWidth),
    height: h,
    // Sits on the floor of the content band by default, which is where a
    // conclusion goes — and, just as importantly, is what stops a strip of
    // nothing opening up between the last block and the footer.
    top: num(args.top, bandBottomFor(args) - h),
  };

  return PowerPoint.run(async (ctx) => {
    const { slide, index } = await targetSlide(ctx, args);
    const made_ = [];
    const keep = (sh) => made_.push(sh);
    // Same routine ppt_build_slide uses — see addKpi for why that matters.
    drawCallout(penFor(slide, keep, keep), kit,
      { text: body, label: args.label, color: args.color, fill: args.fill }, box);

    made_.forEach((s) => s.load('id'));
    await ctx.sync();
    made_.forEach((s) => remember(slide.id, s.id));
    return `已在投影片 ${index} 放上洞察框，寬 ${box.width} 點，位置 (${box.left},${box.top})。`;
  });
}

async function addCards(args) {
  if (!hasApi('1.4')) throw new Error('卡片需要 PowerPointApi 1.4，這個 Office 版本太舊。');
  const cards = Array.isArray(args.cards) ? args.cards.filter(Boolean) : [];
  if (!cards.length) throw new Error('cards 是必要的，至少要一張。');
  if (cards.length > 8) throw new Error('一排最多 8 張卡片——再多在 A4 寬的投影片上就窄到讀不了了。');

  const lk = look(args.look);
  const th = paletteFor(args, lk);
  const face = lk.font || th.font;
  const bar = String(args.bar || (cards.length > 3 ? 'top' : 'left')).toLowerCase();
  const cols = clamp(num(args.columns, bar === 'top' ? cards.length : Math.min(2, cards.length)), 1, 8);
  const rows = Math.ceil(cards.length / cols);

  const gutter = num(args.gutter, 24);
  const left = num(args.left, CHROME.contentLeft);
  const top = num(args.top, CHROME.contentTop);
  const totalW = num(args.width, CHROME.contentWidth);
  const cw = (totalW - gutter * (cols - 1)) / cols;
  // 92 was the height of a card carrying a heading. A process step carries a
  // heading AND two or three lines under it, and at 92 those lines had nowhere
  // to go — which is what made the four-step row look cramped rather than
  // designed. Measured against the content rather than guessed once.
  const hasBody = cards.some((c) => c && c.text);
  const ch = num(args.height, bar === 'top' ? (hasBody ? 118 : 92) : 68);
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
      const sk = skin(th, lk, accent,
        colour(card.fill, bar === 'top' ? th.bg : th.tint));
      const panel = add(slide.shapes.addGeometricShape(
        boxOf(lk, bar === 'top' ? 'RoundRectangle' : 'Rectangle'),
        { left: x, top: y, width: cw, height: ch }));
      panel.name = stamp('card');
      if (sk.fill) panel.fill.setSolidColor(card.fill ? colour(card.fill) : sk.fill);
      else panel.fill.clear();
      // A left-bar card is normally borderless — the tint alone separates it.
      // An outlined look has no tint to rely on, so it gets the line instead.
      if (bar === 'top' || lk.card === 'outline') {
        panel.lineFormat.visible = sk.line;
        panel.lineFormat.color = sk.line ? accent : sk.color;
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
      // Where the body starts and how wide it runs, decided BEFORE the badge
      // shifts tx. See the block below the badge for why that matters.
      const bodyX = tx;
      const bodyW = tw;

      // The number badge. Big, in the accent, set to the LEFT of the text rather
      // than above it — it is an index, not a heading.
      if (card.number) {
        const badge = add(slide.shapes.addTextBox(String(card.number),
          { left: tx, top: ty, width: 52, height: 40 }));
        badge.name = stamp('num');
        styleText(badge, { font: face, size: 24, bold: true, color: accent, valign: 'top' });
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
          font: face, size: num(args.size, 14), bold: true,
          color: card.title_color ? colour(card.title_color) : th.ink, valign: 'top',
        });
        title.fill.clear();
        title.lineFormat.visible = false;
        ty += 26;
      }
      if (card.text) {
        /**
         * The body runs the FULL width of the card, not the width left over
         * beside the number.
         *
         * ── The bug, and why it only showed up on process rows ──
         * `tx`/`tw` are shifted right by 56 points when there is a number badge,
         * because the title has to sit beside it. The body was inheriting that
         * shift, so on a four-across row each paragraph was rendering into about
         * 119 points — barely eight characters — and every line broke. It looked
         * like the text was too long. It was not; it was being given a third of
         * the card to live in.
         *
         * The badge is an index at the top-left corner. Nothing below the title
         * needs to avoid it.
         */
        const body = add(slide.shapes.addTextBox(String(card.text),
          { left: bodyX, top: ty, width: bodyW, height: y + ch - ty - 10 }));
        body.name = stamp('cb');
        styleText(body, { font: face, size: num(args.small, 11.5), color: th.muted, valign: 'top' });
        body.fill.clear();
        body.lineFormat.visible = false;
      }
      // The right-hand reference (「P.03」 in the reference deck), for agendas.
      if (card.note && bar === 'left') {
        const note = add(slide.shapes.addTextBox(String(card.note),
          { left: x + cw - 74, top: y + ch / 2 - 14, width: 58, height: 28 }));
        note.name = stamp('cn');
        styleText(note, { font: face, size: 11, color: th.muted, align: 'right', valign: 'middle' });
        note.fill.clear();
        note.lineFormat.visible = false;
      }
    });

    // Arrows between cards, for a process row. Drawn after the cards so they sit
    // in the gutters rather than under a panel.
    if (args.arrows && bar === 'top' && cols > 1) {
      for (let c = 0; c < Math.min(cols, cards.length) - 1; c++) {
        const ax = left + (cw + gutter) * c + cw + 5;
        // In `line` these were all but invisible — line is the hairline colour,
        // meant for gridlines under other things, and an arrow nobody can see
        // turns a process into four unrelated boxes. `muted` is the quietest
        // colour that still reads as a mark on the page.
        const arrow = add(slide.shapes.addGeometricShape('RightArrow',
          { left: ax, top: top + ch / 2 - 6, width: gutter - 10, height: 12 }));
        arrow.name = stamp('arrow');
        arrow.fill.setSolidColor(th.muted);
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
/**
 * Put the cursor on the slide an image is about to land on.
 *
 * ── The bug this fixes ──
 * Every picture in this module goes in through setSelectedDataAsync, which is a
 * Common API: it inserts into the CURRENT SELECTION and has no idea which slide
 * the caller meant. Building a deck never changes the selection, so a run that
 * created eight slides and drew a chart on each of them put all eight charts on
 * whichever slide the person happened to be looking at — slide 1, in practice.
 * Nothing failed; the deck was simply wrong, and wrong in a way that looks like
 * the model chose it.
 *
 * Every other tool here resolves its target through targetSlide, which defaults
 * to the LAST slide — the one just built. This makes the image path agree with
 * the rest, by moving the selection first.
 *
 * Needs PowerPointApi 1.6 for setSelectedSlides. Below that the old behaviour is
 * all there is, so the caller is told rather than silently misplacing the image.
 */
async function focusSlideFor(args) {
  if (!hasApi('1.6')) {
    return num(args.slide, 0)
      ? '（這個 Office 版本沒辦法指定投影片，圖片會插在目前選取的那一張。）'
      : '';
  }
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load('items/id');
    await ctx.sync();
    const items = slides.items || [];
    if (!items.length) throw new Error('這份簡報沒有投影片，先用 ppt_add_slide 新增一張。');
    const n = num(args.slide, 0);
    const idx = n ? clamp(n, 1, items.length) - 1 : items.length - 1;
    ctx.presentation.setSelectedSlides([items[idx].id]);
    await ctx.sync();
    return '';
  });
}

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
  const note = await focusSlideFor(args);
  await insertImageData(data, Office.CoercionType.Image, geo);
  return `已插入圖片，${geo.width}×${geo.height} 點，位置 (${geo.left},${geo.top})。${note}`;
}

/**
 * ── Photographs, and why the add-in fetches them itself ──
 *
 * The web platform puts a photo on a slide by writing an <img src> into HTML and
 * letting the browser do the rest. PowerPoint has no equivalent: Office.js takes
 * image BYTES, base64, through setSelectedDataAsync. There is no URL form.
 *
 * So somebody has to do the download, and the choice is the pane or the server.
 * This does it in the pane, for one reason worth stating plainly: a server route
 * that fetches any URL a language model asks for is an SSRF proxy sitting inside
 * the corporate network, reachable by anything that can reach the bridge. The
 * pane runs in the user's own WebView with the user's own network position — it
 * can already reach exactly what the user can reach, and nothing more. That is a
 * much smaller thing to get wrong.
 *
 * The cost is CORS: the remote host must allow a cross-origin read.
 * images.unsplash.com does. Most corporate file shares do not, which is why the
 * error below says which half failed instead of just 「插入圖片失敗」.
 */

/**
 * Is this an illustration rather than a photograph?
 *
 * It matters because the two go in through different doors. A JPEG becomes
 * base64 and goes through Office.CoercionType.Image. An SVG has to go through
 * XmlSvg as TEXT — base64 through the Image door produces a broken picture
 * placeholder, silently. Same call, same arguments, nothing thrown.
 *
 * Going in as a vector is also the better outcome on its own: the illustration
 * stays sharp when somebody drags it bigger, and it survives being printed.
 */
const isSvg = (url, type) => /^image\/svg/i.test(type || '') || /\.svg(\?|$)/i.test(url);

/** The crop each mode actually needs. 16:9 unless the slot says otherwise. */
const PHOTO_MODES = {
  hero: { w: 1920, h: 1080 },
  side: { w: 768, h: 1080 },
  block: { w: 1280, h: 720 },
};

/**
 * Ask the CDN for the crop we are going to draw.
 *
 * A 16:9 photo stretched into a 384×540 side panel is visibly wrong — faces go
 * narrow, horizons bend — and setSelectedDataAsync stretches, it does not crop.
 * Unsplash will crop server-side if asked, so ask: the model supplies a photo,
 * this decides its shape. Every other host is left alone, because there the
 * query string means something we do not know.
 */
function cropUrl(raw, mode) {
  const want = PHOTO_MODES[mode] || PHOTO_MODES.block;
  try {
    const u = new URL(raw);
    if (isSvg(raw, '')) return raw;
    if (!/(^|\.)unsplash\.com$/i.test(u.hostname)) return raw;
    u.searchParams.set('w', String(want.w));
    u.searchParams.set('h', String(want.h));
    u.searchParams.set('fit', 'crop');
    u.searchParams.set('q', '80');
    return u.toString();
  } catch {
    return raw;
  }
}

/** ArrayBuffer → base64, in chunks: apply() on a 2MB array overflows the stack. */
function toBase64(buf) {
  const bytes = new Uint8Array(buf);
  let out = '';
  const CHUNK = 0x8000;
  for (let i = 0; i < bytes.length; i += CHUNK) {
    out += String.fromCharCode.apply(null, bytes.subarray(i, i + CHUNK));
  }
  return btoa(out);
}

/** 6MB. Past that PowerPoint gets slow and the deck gets unmailable. */
const PHOTO_MAX = 6 * 1024 * 1024;

async function fetchPhoto(url, mode) {
  let res;
  try {
    res = await fetch(cropUrl(url, mode));
  } catch (e) {
    throw new Error(`抓不到那張圖：${e.message}。`
      + '通常是圖片主機不允許跨來源讀取（CORS），或這台電腦連不到外網。'
      + '換一個公開圖床的網址，或請使用者把圖片直接拖進對話框。');
  }
  if (!res.ok) throw new Error(`抓不到那張圖（HTTP ${res.status}）。網址可能已經失效。`);
  const type = String(res.headers.get('content-type') || '');
  if (type && !/^image\//i.test(type)) {
    throw new Error(`那個網址回來的不是圖片（${type}）。請給圖片檔本身的網址。`);
  }
  if (isSvg(url, type)) {
    const svg = await res.text();
    if (svg.length > PHOTO_MAX) throw new Error('這張向量圖太大了，換一張。');
    return { svg };
  }
  const buf = await res.arrayBuffer();
  if (buf.byteLength > PHOTO_MAX) {
    throw new Error(`圖片太大（${Math.round(buf.byteLength / 1024 / 1024)}MB）。`
      + '在網址加上 ?w=1600 之類的參數把它縮小。');
  }
  return { data: toBase64(buf) };
}

/**
 * Put a photograph on a slide — composed, not merely placed.
 *
 * ── Three modes, three different jobs ──
 *   hero   full-bleed behind a darkened scrim, title on top. Opens a deck or a
 *          section, and it is the thing the reference decks do that we had no
 *          answer for at all.
 *   side   a full-height 40% panel with the content beside it. The magazine
 *          split: it makes an ordinary content slide look composed without the
 *          photo needing to carry any argument.
 *   block  contained in the content area. For a photo that IS the evidence.
 *
 * ── The scrim is not decoration ──
 * White type over an unknown photograph is unreadable about half the time, and
 * we never get to look at the photograph first. So the darkening is applied
 * always and its strength is a parameter — a judgement made in the open rather
 * than blind.
 */
async function addPhoto(args) {
  if (!hasImageCoercion('1.1')) {
    throw new Error('這個 Office 版本不支援插入圖片（需要 ImageCoercion 1.1，Office 2021 以上）。');
  }
  const url = String(args.url || '').trim();
  if (!/^https?:\/\//i.test(url)) throw new Error('url 是必要的，而且要是 http(s) 開頭的圖片網址。');
  const mode = String(args.mode || 'block').toLowerCase();
  if (!PHOTO_MODES[mode]) throw new Error(`mode 必須是：${Object.keys(PHOTO_MODES).join('、')}。`);

  const lk = look(args.look);
  const th = paletteFor(args, lk);
  const face = lk.font || th.font;
  const right = String(args.position || 'right').toLowerCase() !== 'left';

  const geo = mode === 'hero'
    ? { left: 0, top: 0, width: GRID.W, height: GRID.H }
    : mode === 'side'
      ? { left: right ? GRID.W - 384 : 0, top: 0, width: 384, height: GRID.H }
      : {
        left: num(args.left, CHROME.contentLeft),
        top: num(args.top, CHROME.contentTop),
        width: num(args.width, CHROME.contentWidth),
        height: num(args.height, 300),
      };

  const got = await fetchPhoto(url, mode);
  const note = await focusSlideFor(args);
  if (got.svg) {
    // Vectors need ImageCoercion 1.2. Saying so beats letting the call fail with
    // whatever Office decides to call it.
    if (!hasImageCoercion('1.2')) {
      throw new Error('這個 Office 版本放不了向量插畫（需要 ImageCoercion 1.2）。'
        + '改用 ppt_find_art 以外的做法：放照片，或用 ppt_add_diagram 畫圖解。');
    }
    await insertImageData(got.svg, Office.CoercionType.XmlSvg, geo);
  } else {
    await insertImageData(got.data, Office.CoercionType.Image, geo);
  }

  // Everything below is drawn AFTER the image, and that ordering is the whole
  // trick: setSelectedDataAsync appends, so a scrim created first would end up
  // underneath the photograph it is supposed to darken.
  const title = String(args.title || '');
  const sub = String(args.subtitle || '');
  const caption = String(args.caption || '');
  if (!title && !sub && !caption && mode !== 'hero') {
    return `已插入圖片，${geo.width}×${geo.height} 點，位置 (${geo.left},${geo.top})。${note}`;
  }

  return PowerPoint.run(async (ctx) => {
    const { slide, index } = await targetSlide(ctx, args);
    const made_ = [];
    const add = (shape) => { made_.push(shape); return shape; };
    const veil = (box, opacity) => {
      const s = add(slide.shapes.addGeometricShape('Rectangle', box));
      s.name = stamp('scrim');
      s.fill.setSolidColor('#000000');
      // transparency is not in every build's requirement set. Without it this is
      // an opaque black rectangle covering the photograph — so a failure here
      // removes the scrim rather than leaving a black slide.
      try {
        s.fill.transparency = opacity;
      } catch {
        s.delete();
        made_.pop();
        return null;
      }
      s.lineFormat.visible = false;
      return s;
    };
    const say = (content, box, style) => {
      const s = add(slide.shapes.addTextBox(String(content), box));
      s.name = stamp('photo');
      styleText(s, { font: face, wrap: true, ...style });
      s.fill.clear();
      s.lineFormat.visible = false;
      return s;
    };

    if (mode === 'hero') {
      // No scrim over a vector illustration: those have transparent or white
      // grounds, so darkening them produces a grey rectangle with a drawing
      // faintly visible inside it rather than a readable background.
      if (!got.svg) {
        const strength = clamp(num(args.scrim, 0.45), 0, 0.9);
        veil({ left: 0, top: 0, width: GRID.W, height: GRID.H }, 1 - strength);
      }
      // White reads over a darkened photo. Over a pale illustration it vanishes,
      // so the type falls back to the palette's own ink.
      const onArt = got.svg ? th.strong : '#FFFFFF';
      const subArt = got.svg ? th.ink : '#FFFFFF';
      if (args.eyebrow) {
        say(String(args.eyebrow), { left: 80, top: 214, width: GRID.W - 160, height: 24 },
          { size: TYPE.small, color: onArt });
      }
      if (title) {
        say(title, { left: 80, top: 244, width: GRID.W - 160, height: 96 },
          { size: TYPE.display, bold: true, color: onArt, valign: 'middle' });
      }
      if (sub) {
        say(sub, { left: 80, top: 348, width: GRID.W - 160, height: 40 },
          { size: TYPE.heading, color: subArt });
      }
    } else if (mode === 'side') {
      // Words go on the other 60%. The return value repeats those coordinates so
      // the model can keep using the ordinary tools for the rest of the slide
      // instead of guessing where the photo stopped.
      const tx = right ? 60 : 444;
      if (title) {
        say(title, { left: tx, top: 150, width: 456, height: 72 },
          { size: lk.titleSize, bold: lk.titleBold, color: th.strong, valign: 'middle' });
      }
      if (sub) {
        say(sub, { left: tx, top: 230, width: 456, height: 120 },
          { size: TYPE.body, color: th.ink });
      }
    }
    if (caption) {
      say(caption, mode === 'hero'
        ? { left: 80, top: 470, width: GRID.W - 160, height: 20 }
        : { left: geo.left, top: geo.top + geo.height + 6, width: geo.width, height: 20 },
        { size: TYPE.caption, color: mode === 'hero' && !got.svg ? '#FFFFFF' : th.muted });
    }

    made_.forEach((s) => s.load('id'));
    await ctx.sync();
    made_.forEach((s) => remember(slide.id, s.id));
    const room = mode === 'side'
      ? `\n文字請放在另外 60%：left ${right ? 60 : 444}、寬 456。`
      : '';
    return `已在投影片 ${index} 用「${mode}」方式放上照片。${note}${room}`;
  });
}

/** The Word host's insertImage contract: a pasted picture from the composer. */
export async function insertImage(pic, args = {}) {
  const data = String((pic && pic.data) || '').replace(/^data:[^,]+,/, '');
  if (!data) throw new Error('沒有圖片資料。');
  if (!hasImageCoercion('1.1')) {
    throw new Error('這個 Office 版本不支援插入圖片（需要 ImageCoercion 1.1）。');
  }
  const geo = fitToBody(num(pic.width, 800), num(pic.height, 450), args);
  await focusSlideFor(args);
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

  const note = await focusSlideFor(args);
  const vector = hasImageCoercion('1.2') && args.raster !== true && png.svg;
  if (vector) {
    await insertImageData(png.svg, Office.CoercionType.XmlSvg, geo);
    return `已插入向量圖表，${geo.width}×${geo.height} 點。${note}\n`
      + '它是 SVG，投影時不會糊掉；使用者在 PowerPoint 裡按右鍵「轉換為圖形」之後，'
      + '還可以直接拖動裡面的方塊。';
  }
  await insertImageData(png.data, Office.CoercionType.Image, geo);
  return `已插入圖表，${geo.width}×${geo.height} 點（點陣圖）。${note}\n`
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
  const note = await focusSlideFor(args);
  await insertImageData(svg, Office.CoercionType.XmlSvg, geo);
  return `已插入${args.title ? `「${args.title}」` : ''}圖表，${geo.width}×${geo.height} 點。${note}\n`
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

/**
 * ── 為什麼這裡沒有備忘稿（演講稿）工具 ──
 *
 * PRD 3.4 的能力表列了「演講稿生成：為每頁生成備忘稿」。做不到，而且不是版本問題
 * ——PowerPoint 的 JavaScript API **從來沒有**暴露過備忘稿：`Slide` 上沒有 notes、
 * 沒有 notesPage，也沒有任何 requirement set 加過它。這是 office-js 上的公開需求
 * （OfficeDev/office-js#3269），至今未實作。查過一次就夠了，不要再查。
 *
 * 三條繞道，以及為什麼都沒走：
 *
 *   1. `presentation.insertSlidesFromBase64()`（PowerPointApi 1.2）確實可以把一份
 *      在伺服器上做好、備忘稿都寫好的 pptx 插進來。但那等於整份重做——它插的是新
 *      投影片，貼不回使用者已經有的那幾張，而「在當下這份檔案原生生成」正是這個
 *      產品存在的理由。
 *   2. 把講稿放進投影片外面的文字方塊：東西是在檔案裡了，但它不是備忘稿，簡報者
 *      模式看不到，列印和某些檢視反而看得到。比沒有更糟。
 *   3. 假裝有這個工具、實際上什麼都沒寫：不考慮。
 *
 * 所以講稿目前只出現在對話裡，由使用者自己貼進備忘稿欄。要改成產品功能的話，正確
 * 的做法是把這個限制寫回 PRD 3.4（那節本來就有一段技術風險），而不是在這裡硬做。
 */
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
    'ppt_add_kpi',
    'ppt_add_callout',
    'ppt_set_text',
    'ppt_format_shape',
    'ppt_delete_shape',
    'ppt_arrange',
    'ppt_apply_theme',
    'ppt_ask_user',
    'ppt_read_attachment',
    'ppt_insert_attachment',
    'ppt_read_file',
    // Answered by the server, not here — but it still has to be advertised.
    'ppt_find_art',
  ];
  // Pictures are a Common API capability, not a PowerPointApi one, so they are
  // gated on their own requirement set rather than on the host version.
  if (hasImageCoercion('1.1')) tools.push('ppt_add_image', 'ppt_add_photo', 'ppt_add_diagram');
  // Charts are vector-only on purpose: a rasterised chart on a projector is
  // worse than no chart, so this needs 1.2 rather than falling back to 1.1.
  if (hasImageCoercion('1.2')) tools.push('ppt_add_chart');
  return tools;
}

const WRITES = new Set([
  'ppt_add_slide', 'ppt_delete_slide', 'ppt_build_slide', 'ppt_add_text',
  'ppt_add_shape', 'ppt_add_line', 'ppt_add_cards', 'ppt_add_kpi', 'ppt_add_callout', 'ppt_add_chart',
  'ppt_set_text', 'ppt_format_shape',
  'ppt_delete_shape', 'ppt_arrange', 'ppt_apply_theme', 'ppt_add_image', 'ppt_add_photo',
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
          ppt_add_kpi: addKpi,
          ppt_add_callout: addCallout,
          ppt_add_chart: addChart,
          ppt_set_text: setText,
          ppt_format_shape: formatShape,
          ppt_delete_shape: deleteShape,
          ppt_arrange: arrange,
          ppt_apply_theme: applyTheme,
          ppt_add_image: addImage,
          ppt_add_photo: addPhoto,
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
  canvas: { width: GRID.W, height: GRID.H, grid: GRID, chrome: CHROME, type: TYPE, themes: THEMES, looks: LOOKS },

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
    ppt_add_kpi: t('正在排數據卡…'),
    ppt_add_callout: t('正在寫結論框…'),
    ppt_add_chart: t('正在畫圖表…'),
    ppt_set_text: t('正在改寫文字…'),
    ppt_format_shape: t('正在調整物件…'),
    ppt_delete_shape: t('正在移除物件…'),
    ppt_arrange: t('正在對齊排列…'),
    ppt_apply_theme: t('正在套用配色…'),
    ppt_add_image: t('正在插入圖片…'),
    ppt_add_photo: t('正在抓圖片並排版…'),
    ppt_add_diagram: t('正在畫流程圖…'),
    ppt_read_attachment: t('正在看你貼的圖…'),
    ppt_insert_attachment: t('正在把圖片放進簡報…'),
    ppt_find_art: t('正在找插畫…'),
    ppt_read_file: t('正在讀你上傳的檔案…'),
    ppt_run_script: t('正在執行程式…'),
  }),

  preparingLabels: () => ({
    ppt_build_slide: t('正在設計投影片…'),
    ppt_arrange: t('正在盤算版面…'),
    ppt_add_diagram: t('正在設計流程圖…'),
    ppt_add_cards: t('正在設計版面…'),
    ppt_add_kpi: t('正在挑要放哪幾個數字…'),
    ppt_add_callout: t('正在想這頁的結論…'),
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
