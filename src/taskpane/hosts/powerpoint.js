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
import { cleanArgs } from './text.js';
import { makeScriptGuard } from './guard.js';
// Lazily loaded inside: the 2.9 MB bundle it pulls in is only fetched the first
// time somebody asks for a diagram.
import { renderMermaid } from './mermaid.js';
// Pure string work — no canvas, no DOM. See svgchart.js for why vector.
import { chartSvg, CHART_TYPES, chartData } from './svgchart.js';
import { ornamentSvg } from './svgart.js';

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
  teal: {
    // 藍綠。Reads as technical without reading as cold — the family most decks
    // reach for when 企業藍 has been used on everything already.
    label: '藍綠', font: '微軟正黑體',
    ink: '#2F3E3F', strong: '#0F6E72', accent: '#E0743A', second: '#4AA7A5',
    tint: '#E8F4F3', muted: '#7E9694', line: '#CFE3E1',
    bg: '#FFFFFF', ground: '#F2F9F8', warn: '#D8A13C',
    cover: '#0F5D61', onCover: '#F1FAF9',
  },
  sky: {
    // 淡藍。Deliberately the lightest palette in the set: the cover is a mid
    // blue rather than a dark one, so a whole deck in it never goes heavy.
    label: '淡藍', font: '微軟正黑體',
    ink: '#33414F', strong: '#2A6FB0', accent: '#F2994A', second: '#79B4E3',
    tint: '#EAF3FB', muted: '#8399AD', line: '#D3E4F3',
    bg: '#FFFFFF', ground: '#F4F9FD', warn: '#E4B34A',
    cover: '#3A83C4', onCover: '#FFFFFF',
  },
  amber: {
    // 黃橘。Warm and high-energy without going red — the palette for a launch,
    // a workshop, anything that should feel like a beginning.
    label: '黃橘', font: '微軟正黑體',
    ink: '#40382C', strong: '#B5730F', accent: '#2C6E7F', second: '#E2A93B',
    tint: '#FDF4E3', muted: '#9C8C74', line: '#EDDDC2',
    bg: '#FFFFFF', ground: '#FFFBF3', warn: '#C0563A',
    cover: '#A8660C', onCover: '#FFF8EC',
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
    /**
     * 預設風格不帶裝飾性的曲線。
     *
     * ── 為什麼把 waves 拿掉 ──
     * 波浪是**裝飾**——它不承擔任何資訊。而一頁上其他的東西都在工作：色條標示層級、
     * 底板分隔區塊、箭頭表示順序。波浪跟它們不是同一種語言，擺在一起會打架。
     *
     * 直線和方塊的好處是跟內容同一套語彙：一條線可以是分隔、底線、色條；一個方塊
     * 可以是底板、卡片、表格的一格。**而波浪只能是波浪。**
     *
     * 換成 grid（方格點陣）而不是拿掉：母題有沒有會改變骨架的繪製路徑（有母題時
     * 底色進裝飾層、沒有就自己鋪一塊），直接設 none 會連帶影響其他版型的格子數。
     * grid 保留同一條路徑，但畫的是直線和方塊。
     *
     * 這是預設風格，所以它影響的是每一份沒特別指定的簡報——拿掉的收益最大。
     * 想要曲線的場合仍然有（soft、handmade），那些要明確指定才會出現。
     */
    label: '企業', desc: '雙色頂條、短底線、方角卡片。給客戶、主管、對外提案。',
    theme: 'corporate', motif: 'grid',
    masthead: 'double', rule: 'short', ground: 'tint',
    titleSize: 25, titleBold: true, titleAlign: 'left', eyebrowCase: 'upper',
    shape: 'Rectangle', card: 'tint', cover: 'panel', topPad: 0,
  },
  editorial: {
    label: '雜誌', desc: '沒有色條，標題大而不粗，下面一條細細的滿版線。像一篇報導的開頭。',
    font: '微軟正黑體 Light',
    theme: 'mono', motif: 'corner',
    masthead: 'none', rule: 'full', ground: 'white',
    titleSize: 30, titleBold: false, titleAlign: 'left', eyebrowCase: 'keep',
    shape: 'Rectangle', card: 'outline', cover: 'split', topPad: 10,
  },
  minimal: {
    label: '極簡', desc: '什麼裝飾都沒有，只有字和留白。內容本身夠強的時候最好看。',
    font: '微軟正黑體 Light',
    theme: 'mono', motif: 'none',
    masthead: 'none', rule: 'none', ground: 'white',
    titleSize: 22, titleBold: false, titleAlign: 'left', eyebrowCase: 'keep',
    shape: 'Rectangle', card: 'plain', cover: 'plain', topPad: 18,
  },
  soft: {
    label: '柔和', desc: '圓角、粗一點的色條、標題置中。給內部分享、教學、比較輕鬆的場合。',
    theme: 'muted', motif: 'arcs',
    masthead: 'round', rule: 'none', ground: 'tint',
    titleSize: 26, titleBold: true, titleAlign: 'center', eyebrowCase: 'keep',
    shape: 'RoundRectangle', card: 'tint', cover: 'stack', topPad: 4,
  },
  bold: {
    label: '大字', desc: '標題佔掉上面一大塊，色塊當底。給只有一個重點、要講得很用力的場合。',
    theme: 'corporate', motif: 'ribbon',
    masthead: 'block', rule: 'none', ground: 'white',
    titleSize: 34, titleBold: true, titleAlign: 'left', eyebrowCase: 'upper',
    shape: 'Rectangle', card: 'tint', cover: 'bleed', topPad: 6,
  },
  tech: {
    label: '科技', desc: '細線、單色、深色底走向。給技術簡報、資料儀表板。',
    theme: 'night', motif: 'grid',
    masthead: 'hairline', rule: 'none', ground: 'strong',
    titleSize: 24, titleBold: true, titleAlign: 'left', eyebrowCase: 'upper',
    shape: 'Rectangle', card: 'outline', cover: 'triptych', topPad: 0,
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
    theme: 'linen', motif: 'none',
    masthead: 'frame', rule: 'none', ground: 'tint',
    titleSize: 30, titleBold: false, titleAlign: 'left', eyebrowCase: 'keep',
    shape: 'Rectangle', card: 'plain', cover: 'paper', topPad: 14,
  },
  academic: {
    label: '學術', desc: '整頁細框、密集分欄、編號分明。論文口試、技術審查、內部評審。',
    theme: 'corporate', motif: 'grid',
    masthead: 'frame', rule: 'short', ground: 'white',
    titleSize: 22, titleBold: true, titleAlign: 'left', eyebrowCase: 'keep',
    shape: 'Rectangle', card: 'outline', cover: 'band', topPad: 0,
  },
  dashboard: {
    label: '儀表板', desc: '左側色柱、點陣底、方格面板。給數據回顧、月結報告。',
    theme: 'slate', motif: 'topo',
    // The rail is the whole skeleton: a full-height column down the left edge
    // with everything indented past it. No other look moves the content area,
    // which is exactly why this one does not read as a recolour of any of them.
    masthead: 'rail', rail: 52, texture: 'dots',
    rule: 'none', ground: 'tint',
    titleSize: 24, titleBold: true, titleAlign: 'left', eyebrowCase: 'upper',
    shape: 'Rectangle', card: 'tint', cover: 'bleed', topPad: 0,
  },
  courseware: {
    label: '教材', desc: '大字、圓角、步驟分明、留白大方。教學、內訓、說明會。',
    theme: 'fresh', motif: 'waves',
    masthead: 'round', rule: 'short', ground: 'tint',
    titleSize: 28, titleBold: true, titleAlign: 'left', eyebrowCase: 'keep',
    shape: 'RoundRectangle', card: 'tint', cover: 'band', topPad: 6,
  },
  handmade: {
    label: '手作', desc: '紙紋底、手繪框線。暖色、鬆一點的場合——提案、分享、對外簡介。',
    theme: 'linen', motif: 'none',
    masthead: 'none', rule: 'none', ground: 'tint',
    texture: 'paper', strokes: 'hand',
    titleSize: 28, titleBold: true, titleAlign: 'left', eyebrowCase: 'keep',
    shape: 'RoundRectangle', card: 'plain', cover: 'paper', topPad: 10,
  },
  whiteboard: {
    label: '白板', desc: '接近全白、手繪線條、沒有色塊。講解、拆解、邊講邊畫的內容。',
    theme: 'mono', motif: 'none',
    masthead: 'none', rule: 'none', ground: 'white',
    strokes: 'hand',
    titleSize: 26, titleBold: false, titleAlign: 'left', eyebrowCase: 'keep',
    shape: 'Rectangle', card: 'plain', cover: 'plain', topPad: 10,
  },
  eink: {
    label: '電子紙', desc: '紙紋、明體、純灰階、細線。長文、報告、需要被讀而不是被看的內容。',
    font: '微軟正黑體', titleFont: '新細明體',
    theme: 'mono', motif: 'none',
    masthead: 'none', rule: 'full', ground: 'tint',
    texture: 'paper',
    titleSize: 30, titleBold: false, titleAlign: 'left', eyebrowCase: 'keep',
    shape: 'Rectangle', card: 'outline', cover: 'paper', topPad: 12,
  },
};

const DEFAULT_LOOK = 'business';

const look = (name) => LOOKS[String(name || deckLook || '').toLowerCase()] || LOOKS[DEFAULT_LOOK];

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
const paletteFor = (args, lk) => theme(args.theme || deckPalette || lk.theme);

/**
 * 整份簡報決定的配色，由 ppt_plan_deck 寫進來。
 *
 * ── 為什麼要讓它「黏住」 ──
 * 配色參數（theme / look）一直都有，但它是**每次呼叫各自指定**的。模型不會每次都帶，
 * 所以永遠落回同一個預設——使用者的感受是「不管下什麼需求，PPT 都是淡藍底加 KPI 卡」。
 *
 * 規劃階段決定一次，之後每一個工具自動跟著走。這跟 Word 那邊 accent 的做法一樣：
 * **讓一個決定貫穿整份，而不是讓每一次呼叫各自決定。** 明確指定 theme 的那一次仍然
 * 優先——「這一頁我要深色」是合理的要求，只是它不該是預設。
 */
let deckPalette = null;

/** 同理：整份決定的版型風格。 */
let deckLook = null;

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

/**
 * Blend two hex colours.
 *
 * Wanted for decoration that is 「the same colour, slightly apart」 — a circle on
 * a panel that has to read as part of it rather than as a second element. Doing
 * this by hand would mean a second hex per theme for every such shape, which is
 * eight more values to keep in step and eight more chances for one to be wrong.
 */
function mix(a, b, t) {
  const hex = (c) => {
    let v = String(c || '').replace('#', '');
    // 「#FFF」 is legal CSS and parsed to garbage by a naive slice: 'FF', 'F', ''
    // gives NaN on the last two channels, and the blend comes out a muddy
    // brown that looks like a deliberate colour choice.
    if (v.length === 3) v = v.split('').map((ch) => ch + ch).join('');
    return [0, 2, 4].map((i) => parseInt(v.slice(i, i + 2), 16) || 0);
  };
  const [r1, g1, b1] = hex(a);
  const [r2, g2, b2] = hex(b);
  const k = Math.max(0, Math.min(1, Number(t) || 0));
  const to = (x, y) => Math.round(x + (y - x) * k).toString(16).padStart(2, '0');
  return `#${to(r1, r2)}${to(g1, g2)}${to(b1, b2)}`.toUpperCase();
}

/**
 * Shape names Office actually accepts.
 *
 * ── Why this list exists ──
 * `addGeometricShape('Oval', …)` looks right, reads right, and is not a member
 * of `PowerPoint.GeometricShapeType`. The circle is called **Ellipse**; "Oval"
 * appears only in the enum's own description text. An invalid name does not
 * throw where it is written — it throws at `ctx.sync()`, which carries the whole
 * slide, so two decorative circles took an entire cover down and the agent
 * reported the composition as unsupported on that machine.
 *
 * The enum has about 180 members and we use six. Listing the six and falling
 * back to a rectangle turns a guess that destroys a slide into a shape that is
 * merely the wrong shape.
 */
const SHAPES_OK = new Set([
  'Rectangle', 'RoundRectangle', 'Ellipse', 'Chevron', 'RightArrow', 'Triangle',
  'Diamond', 'Donut', 'Frame', 'Plus', 'Star5', 'Hexagon', 'Pentagon', 'Trapezoid',
  'HomePlate', 'Wave', 'Arc', 'Pie', 'Teardrop', 'Cloud', 'Heart', 'Moon',
]);
const geoType = (name) => (SHAPES_OK.has(String(name)) ? String(name) : 'Rectangle');

/**
 * A coloured bar laid on a panel, without the corners fighting.
 *
 * ── The problem, which only shows up on the rounded looks ──
 * The cap across the top of a card is a plain rectangle. On a square panel that
 * is right. On a ROUNDED one the bar's square corners stick out past the
 * panel's curve, so the card reads as 「圓弧 ＋ 直條」 — two shapes that were
 * clearly not drawn together.
 *
 * Inset rather than rounded: a rounded bar 4 points tall renders as a lozenge,
 * which is worse. Pulling it in by the corner radius means it starts where the
 * panel's straight edge starts, and the two read as one object.
 */
function capBox(shape, box, vertical) {
  // Keyed off the PANEL's resolved shape, not the look's preference: a top-bar
  // card is a RoundRectangle in every look, so asking the look would have left
  // 企業風 with the same mismatched corners it started with.
  if (shape !== 'RoundRectangle') return box;
  const inset = 9;
  return vertical
    ? { ...box, top: box.top + inset, height: Math.max(2, box.height - inset * 2) }
    : { ...box, left: box.left + inset, width: Math.max(2, box.width - inset * 2) };
}

/** RoundRectangle if the look asks for it, otherwise whatever the caller wanted. */
const boxOf = (lk, base) => geoType(lk.shape === 'RoundRectangle' ? 'RoundRectangle' : base);

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
/**
 * 把一個顏色往某個方向推 t（0~1）。t=0 是原色，t=1 是目標色。
 *
 * ── 這個函式為什麼在這裡出現第二份 ──
 * 原本只寫在 word.js 裡。我在 ppt_add_infographic 用了它，卻沒有在這支檔案裡定義——
 * 結果是五種資訊圖**全部**在推導顏色那一步炸掉，而 Office 的錯誤訊息是空的，
 * 所以模型只看到「未知的失敗」。
 *
 * 有趣的是它不是每次都炸：`tone()` 在 items 自己帶了 color 的時候會提早回傳，
 * 根本不會走到這裡。所以「指定顏色就成功、不指定就失敗」——那正是模型回報的現象，
 * 而我花了三輪在猜座標和 schema。
 *
 * 不從 word.js 匯入是因為那會把整個 Word 宿主拉進 PowerPoint 的載入路徑。
 * 十行的顏色工具，兩邊各放一份比綁在一起好。
 */
function mixHex(hex, toward, t) {
  const h = String(hex).replace('#', '');
  const g = String(toward).replace('#', '');
  const ch = (i) => {
    const a = parseInt(h.slice(i * 2, i * 2 + 2), 16);
    const b = parseInt(g.slice(i * 2, i * 2 + 2), 16);
    return Math.round(a + (b - a) * t).toString(16).padStart(2, '0');
  };
  return `#${ch(0)}${ch(1)}${ch(2)}`.toUpperCase();
}

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

/**
 * What this tool call did, expressed as things that can be taken back.
 *
 * ── Why a list of operations and not a snapshot ──
 * PowerPoint has no cheap whole-file snapshot below API 1.10 (Office 2601), so
 * the Word approach — grab the whole body as OOXML, put it back — is not
 * available to almost anybody. What IS available is reading a shape's state
 * before touching it, which is enough to invert every edit this module makes.
 *
 * ── Why 「只記新增」 was not enough ──
 * It covered additions only, so rewriting the text on slide 5 or deleting a
 * shape could not be taken back, and the note said so. But this add-in's whole
 * position is that it works on files that already exist — and a tool that edits
 * your existing deck without being able to put it back teaches people not to let
 * it near one. Undo is not a feature here, it is the condition for the rest.
 *
 * Each entry is an inverse, applied in reverse order:
 *   created  — it was not there before → delete it
 *   restore  — it was there and we changed it → put the old values back
 *   recreate — it was there and we deleted it → build it again
 *   slide    — a whole slide went → rebuild it from what it held
 */
let undoOps = [];
const remember = (slideId, shapeId) => undoOps.push({ op: 'created', slideId, shapeId });

/**
 * Shapes we can rebuild from what we can read.
 *
 * A picture's bytes and a group's structure are not readable through the API, so
 * a deleted one is gone for good. Saying which is the point: 「已還原 5 個物件，
 * 1 張圖片沒辦法還原」 is a true sentence somebody can act on, and silently
 * restoring five of six is not.
 */
const REBUILDABLE = new Set(['geometricshape', 'textbox']);

/** Types with no text frame — asking for one throws and poisons the batch. */
const NO_TEXT = new Set(['image', 'picture', 'line', 'group', 'unsupported']);

/**
 * Queue every property the undo needs, tolerating the ones a shape cannot have.
 *
 * `ctx.sync()` rejects the WHOLE batch if any single load is invalid, so a
 * picture in the middle of a slide would otherwise take the whole capture down
 * with it — the same failure that made ppt_read_slide return empty text for
 * every shape on a slide that happened to contain one photo.
 */
function queueState(shape) {
  const st = { shape, fill: null, line: null, range: null, font: null };
  shape.load('id,name,type,left,top,width,height');
  try { shape.fill.load('type,foregroundColor'); st.fill = shape.fill; } catch { /* no fill */ }
  try { shape.lineFormat.load('visible,color,weight'); st.line = shape.lineFormat; } catch { /* none */ }
  if (!NO_TEXT.has(String(shape.type || '').toLowerCase())) {
    try {
      const f = shape.textFrame;
      f.load('hasText');
      f.textRange.load('text');
      f.textRange.font.load('name,size,color,bold,italic');
      st.range = f.textRange;
      st.font = f.textRange.font;
    } catch { /* pictures and lines have none */ }
  }
  return st;
}

/** Read one back after the sync. Every property is optional on purpose. */
function readState(st) {
  const s = st.shape;
  const grab = (fn, fallback) => { try { const v = fn(); return v === undefined ? fallback : v; } catch { return fallback; } };
  const out = {
    shapeId: grab(() => String(s.id || ''), ''),
    name: grab(() => String(s.name || ''), ''),
    type: grab(() => String(s.type || ''), ''),
    box: {
      left: grab(() => Number(s.left), undefined),
      top: grab(() => Number(s.top), undefined),
      width: grab(() => Number(s.width), undefined),
      height: grab(() => Number(s.height), undefined),
    },
  };
  if (st.fill) {
    // A gradient or a picture fill reads as something other than Solid. We can
    // put a solid colour back but not a gradient, so anything else is recorded
    // as 「unknown」 and left alone rather than flattened to a colour it never had.
    const kind = grab(() => String(st.fill.type || ''), '');
    if (/solid/i.test(kind)) out.fill = grab(() => String(st.fill.foregroundColor || ''), '');
    else if (/nofill|none/i.test(kind)) out.fill = null;
  }
  if (st.line) {
    out.line = {
      visible: grab(() => st.line.visible, undefined),
      color: grab(() => st.line.color, undefined),
      weight: grab(() => st.line.weight, undefined),
    };
  }
  if (st.range) out.text = grab(() => String(st.range.text ?? ''), undefined);
  if (st.font) {
    out.font = {
      name: grab(() => st.font.name, undefined),
      size: grab(() => st.font.size, undefined),
      color: grab(() => st.font.color, undefined),
      bold: grab(() => st.font.bold, undefined),
      italic: grab(() => st.font.italic, undefined),
    };
  }
  return out;
}

/**
 * Capture a set of shapes, and never let the capture break the edit.
 *
 * If reading the old state fails, the edit still happens — it just cannot be
 * taken back, and restoreSnapshot says how many. The alternative is refusing to
 * edit because we could not prepare an undo, which trades a recoverable問題 for
 * an unusable tool.
 */
async function captureShapes(ctx, shapes) {
  const queued = shapes.map((s) => { try { return queueState(s); } catch { return null; } });
  try {
    await ctx.sync();
  } catch {
    return [];
  }
  return queued.map((q) => { try { return q ? readState(q) : null; } catch { return null; } })
    .filter(Boolean);
}

/** This call changed these shapes; here is what they looked like first. */
const rememberBefore = (slideId, states) => {
  for (const before of states) {
    if (before && before.shapeId) undoOps.push({ op: 'restore', slideId, before });
  }
};

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

/** Shape types that can never hold text. Queuing a load against one poisons the batch. */
const TEXTLESS = new Set(['image', 'line', 'group', 'unsupported', 'picture']);

/**
 * Text out of every shape on a slide, or '' for the ones that have none.
 *
 * ── The bug this is written against, which was worse than it looked ──
 * The first version queued a textFrame load for every shape and ran ONE sync.
 * The per-shape try/catch around `s.textFrame` looks like it covers the failure
 * but does not: reading the property queues work, it does not execute it. The
 * throw happens later, inside `ctx.sync()`, and a sync that rejects rejects for
 * the WHOLE batch. So the catch returned '' for every shape on the slide.
 *
 * The visible symptom was 「那頁的文案我讀不到」: one picture anywhere on a slide
 * and the model went blind to all of its text — then edited it by guessing, and
 * landed things on top of words it could not see.
 *
 * It also got worse exactly as the product got better. A slide had to contain a
 * picture to trigger it, and the texture layer, the photo modes and the
 * generated backgrounds all put one there.
 *
 * ── Two defences, because one is not enough ──
 * Skipping the types that cannot hold text avoids the throw in the ordinary
 * case. But `type` is a string from a growing enum and a build we have not seen
 * may report something new, so a batch that still fails falls back to reading
 * one shape at a time. That costs a round trip per shape and only happens when
 * the cheap path has already failed — which is the right way round, because the
 * alternative is silently returning nothing.
 */
async function loadTexts(ctx, shapes) {
  const queue = (list) => list.map((s) => {
    if (TEXTLESS.has(String(s && s.type || '').toLowerCase())) return null;
    try {
      const f = s.textFrame;
      f.load('hasText,textRange/text');
      return f;
    } catch {
      return null;
    }
  });
  const read = (f) => {
    if (!f) return '';
    try {
      return f.hasText ? String(f.textRange.text || '') : '';
    } catch {
      return '';
    }
  };

  const frames = queue(shapes);
  try {
    await ctx.sync();
    return frames.map(read);
  } catch {
    // The batch was poisoned by at least one shape. Read them individually so
    // the rest of the slide still comes back.
    const out = [];
    for (const s of shapes) {
      const [f] = queue([s]);
      if (!f) { out.push(''); continue; }
      try {
        await ctx.sync();
        out.push(read(f));
      } catch {
        out.push('');
      }
    }
    return out;
  }
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
    const slides = await loadSlides(ctx);
    const a = clamp(from, 1, slides.length);
    const b = clamp(Math.max(to, a), a, slides.length);

    /**
     * Write down what each slide held before it goes.
     *
     * ── What comes back, and what does not ──
     * A rebuilt slide is a rebuild. The boxes, their text, colours and positions
     * come back; a photograph does not, and neither does the original position
     * in the deck — PowerPoint's API appends, so a restored slide 3 lands at the
     * end. That is worse than a real undo and much better than the old answer,
     * which was 「這個動作復原鍵拿不回來」 and meant it.
     */
    const taken = [];
    for (let i = a - 1; i <= b - 1; i++) {
      taken.push({
        index: i + 1,
        shapes: await captureShapes(ctx, slides[i].shapes.items || []),
      });
    }
    // Reversed so the ops list, which is replayed backwards, rebuilds them in
    // their original order.
    for (const t of taken.slice().reverse()) undoOps.push({ op: 'slide', ...t });

    // Backwards: deleting slide 3 renumbers everything after it, so a forward
    // loop deletes the wrong slides from the second iteration onward.
    for (let i = b - 1; i >= a - 1; i--) slides[i].delete();
    await ctx.sync();
    const lost = taken.flatMap((t) => t.shapes)
      .filter((x) => !REBUILDABLE.has(String(x.type || '').toLowerCase())).length;
    return `已刪除投影片 ${a}${b > a ? `–${b}` : ''}（共 ${b - a + 1} 張）。`
      + '\n復原鍵可以把內容重建回來，但會接在簡報最後面，'
      + '而且是重建不是原件'
      + (lost ? `，其中 ${lost} 個圖片或群組拿不回來` : '')
      + '。順序要對請用 PowerPoint 自己的 Ctrl+Z。';
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
/**
 * At or below this size, a box is carrying body copy rather than display type,
 * and is not allowed to shrink. 16 is the smallest the design rules call
 * readable from the back of a room; anything set at or under it has no room to
 * give away.
 */
const AUTOFIT_FLOOR = 16;

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

  /**
   * Line spacing and the gap between paragraphs.
   *
   * ── Why this is not cosmetic ──
   * 「文字擠在一起」 was the complaint, and it is the difference between a slide
   * that looks typeset and one that looks pasted. PowerPoint's own default is
   * single spacing, which is right for a page of prose read at arm's length and
   * wrong for four lines of Chinese projected across a room.
   *
   * ── Why it is gated rather than tried ──
   * ParagraphFormat only gained lineSpacing / spaceBefore / spaceAfter in
   * PowerPointApi 1.8. Setting an unknown property on an Office.js proxy does
   * not throw where it is written — it throws at the next `ctx.sync()`, and that
   * sync carries the whole slide. The textFrame bug in loadTexts was exactly
   * this shape, so this asks the requirement set instead of hoping.
   */
  if ((opts.lineSpacing || opts.spaceAfter || opts.spaceBefore) && hasApi('1.8')) {
    const pf = range.paragraphFormat;
    try {
      if (opts.lineSpacing) pf.lineSpacing = Number(opts.lineSpacing);
      if (opts.spaceAfter) pf.spaceAfter = Number(opts.spaceAfter);
      if (opts.spaceBefore) pf.spaceBefore = Number(opts.spaceBefore);
    } catch { /* an older build than it claimed; the text is still readable */ }
  }

  const v = V_ALIGN[String(opts.valign || '').toLowerCase()];
  if (v) frame.verticalAlignment = v;

  if (opts.wrap !== undefined) frame.wordWrap = !!opts.wrap;
  /**
   * Shrink to fit — but not for body copy.
   *
   * ── Why the default had to be split ──
   * Shrink-to-fit exists because text hanging out of its box is the most
   * obvious way an auto-built slide looks broken. It has a failure of its own
   * though, and it is worse: there is no floor. A caption that does not fit
   * comes back at six points, which nobody can read and which does not look
   * like an error — it looks like the add-in chose that size.
   * 「openai 底下小字只有 6px，人眼幾乎看不到」 was exactly that.
   *
   * Office.js has no minimum-size setting, so the only lever is which boxes are
   * allowed to shrink at all. Display type can: a 40pt heading dropping to 32 is
   * still a heading. Body copy cannot — it either fits or it overflows, and an
   * overflow is a bug somebody can SEE and report.
   */
  const bodySize = Number(opts.size) && Number(opts.size) <= AUTOFIT_FLOOR;
  frame.autoSizeSetting = (opts.autofit === false || (bodySize && opts.autofit === undefined))
    ? 'AutoSizeNone'
    : 'AutoSizeTextToFitShape';
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
  /**
   * The model picks this name, so it is the likeliest place to be handed one
   * Office does not have. A bad name throws at sync and takes everything queued
   * with it, so an unknown name becomes a rectangle plus a sentence rather than
   * a failed call.
   */
  const asked = String(args.shape || 'Rectangle');
  const kind = geoType(asked);
  const swapped = kind !== asked;
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
    return `已在投影片 ${index} 加上 ${kind}（${Math.round(geo.left)},${Math.round(geo.top)}）。`
      + (swapped
        ? `
（PowerPoint 沒有「${asked}」這個形狀，改用了矩形。`
          + '圓形要用 Ellipse，箭頭是 RightArrow、Chevron。）'
        : '');
  });
}

/**
 * 一次畫一整批物件。
 *
 * ── 為什麼非有不可 ──
 * 現有的工具全部是**模組**：ppt_add_kpi 給我資料我排成數字卡、ppt_add_table 排成表、
 * ppt_add_cards 排成卡片。每一個的格式都是固定的，所以模型能做的事只有「挑模組、填資料」。
 * 使用者的形容很準：「我這頁要有哪些模組，有表格，就做表格，沒了」。
 *
 * 真正好的那些看板不是這樣做的。它們是**一塊一塊畫出來的**：表格是一格一格的矩形，
 * 卡片是底板加色條加三層字，區塊之間還有箭頭和小結論。所以它們能做到四個區塊再各自
 * 拆細——因為沒有任何模組在限制粒度。
 *
 * 這件事理論上用 ppt_add_shape 就做得到，但一次呼叫畫一個物件：一份看板六十幾個物件
 * 就是六十次往返。模型退回模組不是不會畫，**是太貴**。
 *
 * 所以這個工具存在的意義只有一個：**讓模組變成捷徑，而不是天花板**。合適就用
 * ppt_add_table，不合適就自己畫。
 *
 * ── 為什麼一次 sync ──
 * 六十個物件分六十次往返要好幾分鐘，而且中間任何一次失敗就留下半張投影片。排進同一個
 * 批次，一次 sync 送出去——這也是 ppt_add_cards 那些模組本來就在做的事，只是它們的
 * 形狀是寫死的。
 */
/**
 * 最後看一次整份，只修「壞掉」的，不碰「設計」。
 *
 * ── 這一條界線是整個工具的重點 ──
 * 版面檢查最容易犯的錯是**管太寬**：它看得到座標，就忍不住去對齊、去等距、去統一大小。
 * 但那些是設計決定——模型剛剛為這一頁想過，檢查沒有資格覆蓋它。Word 那邊已經為此
 * 誤判過三次（封面字級被放大、合約一頁一條、年曆被刪），每一次都是**自動修正的範圍
 * 超過了它能判斷的範圍**。
 *
 * 所以這裡只抓「不管什麼設計都不該發生」的三件事：
 *   · 物件跑到投影片外面——那是絕對的，960×540 之外沒有東西看得到
 *   · 兩個有字的物件疊在一起——字疊字一定讀不了
 *   · 字小到看不清楚
 *
 * 前兩件裡只有第一件自動修（夾回邊界內，位置是算得出來的）；疊在一起**只報告**，
 * 因為該讓誰讓開是設計問題，而我們只看得到座標。
 */
async function reviewLayout(args) {
  if (!hasApi('1.4')) throw new Error('版面檢查需要 PowerPointApi 1.4。');
  const fix = args.fix !== false;
  const W = GRID.W;
  const H = GRID.H;

  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load('items/id');
    await ctx.sync();
    const all = slides.items || [];
    const only = Number(args.slide) || 0;
    const targets = only ? all.slice(only - 1, only) : all;
    if (!targets.length) throw new Error('找不到要檢查的投影片。');

    const loaded = targets.map((sl) => {
      const shapes = sl.shapes;
      /**
       * 文字也要載入。
       *
       * ── 沒載入的代價 ──
       * 第一版只載座標，於是「這一塊是不是空的」只能看「裡面有沒有別的物件」——
       * 而一個自己有字的方塊（標題、六角形裡的 S、說明文字）裡面當然沒有別的物件，
       * 就全部被報成空的。使用者一次收到 13 個誤判，而其中一個都不是真的。
       *
       * **一個會喊狼來了的檢查，比沒有檢查更糟**——它會讓人學會忽略所有警告，
       * 包括真的那一個。
       */
      shapes.load('items/id,items/name,items/left,items/top,items/width,items/height,'
        + 'items/textFrame/textRange/text');
      return shapes;
    });
    await ctx.sync();

    const found = [];
    const fixed = [];

    loaded.forEach((shapes, k) => {
      const n = only || k + 1;
      const items = (shapes.items || []).map((sh, i) => ({
        sh,
        i: i + 1,
        l: Number(sh.left) || 0,
        t: Number(sh.top) || 0,
        w: Number(sh.width) || 0,
        h: Number(sh.height) || 0,
        name: String(sh.name || ''),
        text: (() => {
          try { return String(sh.textFrame.textRange.text || '').trim(); } catch { return ''; }
        })(),
      }));

      // 1. 跑到版面外。這個是絕對的，夾回來就對了。
      for (const o of items) {
        const over = o.l < -1 || o.t < -1 || o.l + o.w > W + 1 || o.t + o.h > H + 1;
        if (!over) continue;
        found.push(`第 ${n} 頁第 ${o.i} 個物件跑到版面外`
          + `（${Math.round(o.l)},${Math.round(o.t)} ${Math.round(o.w)}×${Math.round(o.h)}）`);
        if (!fix) continue;
        try {
          // 先夾寬高（比版面還大的東西夾位置沒有用），再夾位置。
          const w = Math.min(o.w, W);
          const h = Math.min(o.h, H);
          if (w !== o.w) o.sh.width = w;
          if (h !== o.h) o.sh.height = h;
          o.sh.left = Math.min(Math.max(0, o.l), W - w);
          o.sh.top = Math.min(Math.max(0, o.t), H - h);
          fixed.push(`把第 ${n} 頁第 ${o.i} 個物件移回版面內`);
        } catch { /* 鎖住的物件就只報告 */ }
      }

      /**
       * 2. 有底色、夠大、裡面一個字都沒有的區塊。
       *
       * ── 為什麼這也算壞掉 ──
       * 使用者看到的是「右邊那塊只有標題，底下是空的」——那通常是某一步失敗了，
       * 或者畫了底板卻忘了填內容。而縮圖上看不出來，是打開檔案才發現。
       *
       * 只報告不自動修：裡面該放什麼不是檢查能決定的。而且條件刻意訂得嚴（面積要
       * 夠大、範圍內真的一個東西都沒有），不然每一塊底板都會被誤報。
       */
      for (const o of items) {
        if (o.w * o.h < 12000) continue;
        // 自己有字就不是空的——這是第一版漏掉的那一條，而它造成了全部的誤判。
        if (o.text) continue;
        const inside = items.some((q) => q !== o
          && q.l >= o.l - 2 && q.t >= o.t - 2
          && q.l + q.w <= o.l + o.w + 2 && q.t + q.h <= o.t + o.h + 2);
        if (inside) continue;
        found.push(`第 ${n} 頁第 ${o.i} 個區塊是空的`
          + `（${Math.round(o.w)}×${Math.round(o.h)}，裡面沒有任何東西）`
          + '——把內容補上，或是把這塊刪掉。不要留著一個空框交出去。');
      }

      /**
       * 2. 兩個物件疊在一起。
       *
       * 只報告不自動修：該讓誰讓開是設計問題。而且底板和它上面的字本來就重疊——
       * 那是正常的堆疊，不是跑版。所以門檻訂在「兩個都有相當面積、而且重疊超過
       * 四成」，底板配小字那種組合不會被誤判。
       */
      for (let a = 0; a < items.length; a += 1) {
        for (let b = a + 1; b < items.length; b += 1) {
          const x = items[a];
          const y = items[b];
          const ox = Math.min(x.l + x.w, y.l + y.w) - Math.max(x.l, y.l);
          const oy = Math.min(x.t + x.h, y.t + y.h) - Math.max(x.t, y.t);
          if (ox <= 2 || oy <= 2) continue;
          const area = ox * oy;
          const small = Math.min(x.w * x.h, y.w * y.h);
          if (!small || area / small < 0.4) continue;
          /**
           * 字壓在底板上是正常的堆疊，不是跑版。
           *
           * 面積比只擋得掉「大底板配小元件」，擋不掉「說明文字框跟它的底板差不多大」
           * ——而那正是最常見的組合，所以第一版把每一塊說明都報成重疊 100%。
           *
           * 真正的判準是：**有字的和沒字的疊在一起本來就該疊**。兩個都有字才是問題，
           * 因為字疊字一定讀不了。
           */
          if (!x.text || !y.text) continue;
          const big = Math.max(x.w * x.h, y.w * y.h);
          if (big / small > 3) continue;
          found.push(`第 ${n} 頁第 ${x.i} 和第 ${y.i} 個物件疊在一起`
            + `（重疊 ${Math.round((area / small) * 100)}%）——這兩塊要讓開一個`);
        }
      }
    });

    await ctx.sync();
    if (!found.length) return '看過了，沒有物件跑到版面外、也沒有疊在一起的。版面是好的。';
    return `檢查完 ${targets.length} 頁，查到 ${found.length} 件：\n`
      + found.map((x) => `· ${x}`).join(`\n`)
      + (fixed.length ? `\n\n修好的：\n` + fixed.map((x) => `· ${x}`).join(`\n`) : '')
      + `\n\n疊在一起和空的區塊沒有自動修——那是設計決定，我只看得到座標。**你自己判斷該怎麼改，然後動手**：讓開一個、補上內容、或是刪掉。改完再跑一次這個工具確認。`;
  });
}

/**
 * 常見的資訊圖，座標算好的。
 *
 * ── 為什麼這個該是模組，而版面不該 ──
 * 「不要用模組、自己畫」那條針對的是**版面**——一頁該分成幾塊、各佔哪裡，那是每一頁
 * 都該重新想的事。但 SWOT、金字塔、環形流程是**圖表類型**，跟長條圖同一類：它們的
 * 結構本來就該一致，而且座標算一次就不用每次重算。
 *
 * 實測證明模型畫得出來（給了座標之後），但每次都要重算一遍三角函數和梯形寬度，
 * 算錯一個數字就會有縫或重疊。把算好的放進來，模型的心力就能花在**這一頁其餘的
 * 部分**——而那才是它該想的。
 *
 * 一個工具五種圖，不是五個工具：它們共用同一套配色推導和同一套回傳格式，拆開只會
 * 讓工具清單變長、而模型更難挑。
 */
async function addInfographic(args) {
  if (!hasApi('1.4')) throw new Error('資訊圖需要 PowerPointApi 1.4。');
  const kind = String(args.kind || '').toLowerCase();
  const KINDS = ['swot', 'pyramid', 'cycle', 'steps', 'venn'];
  if (!KINDS.includes(kind)) throw new Error(`kind 必須是：${KINDS.join('、')}。`);
  const items = Array.isArray(args.items) ? args.items.filter(Boolean) : [];
  if (!items.length) throw new Error('items 是必要的，每個至少要有 label。');

  const lk = look(args.look);
  const th = paletteFor(args, lk);
  const face = lk.font || th.font;
  const ink = args.text_color || th.ink;
  const onBlock = args.block_text || th.onCover;
  // 每個區塊的顏色：給了就用，沒給就從主色推一組同色相的深淺階。
  const tone = (i, n) => {
    const given = items[i] && items[i].color;
    if (given) return colour(given);
    const base = colour(args.accent, th.strong);
    return mixHex(base, '#FFFFFF', n > 1 ? (i / (n - 1)) * 0.55 : 0);
  };

  const L = num(args.left, 40);
  const T = num(args.top, 120);
  /**
   * 預設寬度分三種，因為這三種圖的旁邊通常放不一樣的東西。
   *
   * ── cycle 要再往左一點 ──
   * 環形圖右邊幾乎一定有一份編號清單（01~05 配說明），而圓形的視覺重心比它的外框
   * 更往外——一個 110 寬的圓繞在半徑 150 上，實際佔的視覺寬度是 520，比方形的同尺寸
   * 元件更顯得「滿」。照 560 置中的話，整簇會往右頂到清單。
   *
   * 所以 cycle 收窄到 500：圓簇的中心落在 290 附近，右邊留得出清單的位置，
   * 而左邊也不會空一大塊。
   */
  /**
   * 預設寬度是**這張圖自己**的寬度，不是整頁的寬度。
   *
   * ── 金字塔原本為什麼畫不出來 ──
   * 預設給 880（幾乎整頁），於是梯形撐滿版面，而右邊的說明欄從 936 開始、畫到 1236——
   * 整排掉在 960 的版面外。算一次就看得出來，但我當初是照「內容區寬度」設的，
   * 沒想到這張圖的說明在**圖的右邊**，不在圖的下面。
   *
   * 會把說明放在右邊的（pyramid）要留出那塊空間；自成一塊的（steps）才用整頁寬。
   */
  const W = num(args.width, kind === 'cycle' ? 500
    : kind === 'swot' ? 560
      : kind === 'pyramid' ? 400
        : kind === 'venn' ? 420 : 880);

  const plan = [];
  const push = (o) => plan.push(o);

  if (kind === 'swot') {
    // 四個六角形聚成一簇，說明在兩側。實測過的座標。
    /**
     * 四個六角形之間要留縫。
     *
     * ── 原本為什麼會互相蓋住 ──
     * 我把它們排在剛好相鄰的位置（水平差一個寬、垂直差 120 而高是 130）。六角形的
     * 外框是矩形，但它的**斜邊在矩形內縮**——兩個外框剛好相鄰的時候，看起來是分開的；
     * 可是垂直方向我少算了 10，於是下面那排直接壓在上面那排的斜邊上。
     *
     * 而且六角形相鄰不像方塊：方塊貼在一起是一道線，六角形貼在一起是一片交錯的斜角。
     * 留縫才看得出來那是四個獨立的區塊。
     */
    const hw = 150;
    const hh = 130;
    const gapX = 14;
    const gapY = 14;
    const cx = L + W / 2;
    const x0 = cx - hw - gapX / 2;
    const x1 = cx + gapX / 2;
    const y0 = T + 50;
    const y1 = y0 + hh + gapY;
    const spots = [[x0, y0], [x1, y0], [x0, y1], [x1, y1]];
    items.slice(0, 4).forEach((it, i) => {
      push({ type: 'rect', shape: 'Hexagon', left: spots[i][0], top: spots[i][1],
        width: hw, height: hh, fill: tone(i, 4),
        text: String(it.key || it.label || '').slice(0, 2), color: onBlock, size: 40, bold: true });
    });
  } else if (kind === 'pyramid') {
    /**
     * 梯形寬度遞增，置中疊起來。
     *
     * 層高固定、寬度等差：這樣每一層的斜邊角度一致，疊起來才像一個完整的三角形。
     * 寬度用比例算而不是寫死，整張圖才縮放得了。
     */
    const n = Math.min(items.length, 6);
    const hFull = num(args.height, 290);
    const rowH = hFull / n;
    const halfW = W * 0.42;
    const cx = L + W / 2;
    /**
     * 第一項在**最底層**，也就是最寬的那一層。
     *
     * ── 為什麼會畫反 ──
     * 工具說明寫「由下而上給，最底層放第一個」，但程式碼把 items[0] 放在最上面——
     * 文件和實作講的是相反的事。模型照文件送「基礎營運 40%」當第一項，結果 40% 跑到
     * 最窄的頂端，它得自己清掉十二個物件重畫一次。
     *
     * 選擇讓程式碼配合文件，不是反過來：講一件事的時候，「基礎」本來就先出口——
     * 先講地基再講屋頂是自然的順序，而金字塔的地基在下面。
     */
    items.slice(0, n).forEach((it, i) => {
      const row = n - 1 - i;          // 第一項畫在最後一列（最下面）
      const w = halfW * (0.3 + (0.7 * (n - i)) / n) * 2;
      const y = T + rowH * row;
      push({ type: 'rect', shape: 'Trapezoid', left: cx - w / 2, top: y,
        width: w, height: rowH - 4, fill: tone(i, n),
        text: String(it.value || ''), color: onBlock, size: 20, bold: true });
      // 引線接到右邊的說明，顏色跟著那一層走。
      push({ type: 'rect', left: cx + w / 2 + 8, top: y + rowH / 2 - 1,
        width: Math.max(12, L + W - (cx + w / 2) - 20), height: 1, fill: tone(i, n) });
      // 說明欄加寬到 420：原本 300 的話右邊會空掉一大塊，而那塊沒有別的東西要放。
      push({ type: 'text', left: L + W + 16, top: y + 2,
        width: 420, height: rowH - 8, text: String(it.label || ''),
        color: ink, size: 14, bold: true, align: 'left' });
    });
  } else if (kind === 'cycle') {
    const n = Math.min(Math.max(items.length, 3), 7);
    const d = num(args.node, 110);
    const cx = L + W / 2;
    // 往上 20：圓簇的視覺重心比外框更低（下面兩個節點佔的空間大），照外框置中會偏下。
    const cy = T + 170;
    const r = num(args.radius, 150);
    /**
     * 先畫軌道，再畫節點——順序就是疊的順序。
     *
     * 沒有那條環的話，五個圓只是五個圓，看不出它們是一個循環。用一個比節點圈大一點的
     * 空心圓當軌道，節點蓋在上面，缺口自然就被節點補起來。
     */
    push({ type: 'rect', shape: 'Ellipse', left: cx - r, top: cy - r,
      width: r * 2, height: r * 2, fill: null,
      line_color: colour(args.track_color, th.line || th.muted), line_width: 1.5 });
    items.slice(0, n).forEach((it, i) => {
      const a = (-Math.PI / 2) + (i * 2 * Math.PI) / n;
      // 編號寫進標籤裡，而不是另外畫一個小圓——那會變成第二層要對齊的東西。
      const seq = String(i + 1).padStart(2, '0');
      const label = it.no === false ? String(it.label || '') : `${seq}
${String(it.label || '')}`;
      push({ type: 'rect', shape: 'Ellipse',
        left: cx + r * Math.cos(a) - d / 2, top: cy + r * Math.sin(a) - d / 2,
        width: d, height: d, fill: tone(i, n),
        text: label, color: onBlock, size: 13, bold: true });
    });
    if (args.hub) {
      /**
       * 軸心的字要看得見。
       *
       * 底是白的，所以字色必須是深的——用 ink 而不是 onBlock（那是給色塊上的字用的，
       * 通常是白色）。實測畫出來是一個空白圓，就是因為白底配到了白字。
       *
       * `||` 在這裡會吃掉空字串，所以用明確的判斷：使用者把字色設成 '' 的時候，
       * 應該退回 ink，而不是變成 undefined 讓 styleText 去猜。
       */
      const hubText = args.hub_text || ink;
      push({ type: 'rect', shape: 'Ellipse', left: cx - 60, top: cy - 60,
        width: 120, height: 120, fill: args.hub_fill || '#FFFFFF',
        line_color: colour(args.accent, th.strong), line_width: 1.5,
        text: String(args.hub), color: hubText, size: 14, bold: true, valign: 'middle' });
    }
  } else if (kind === 'steps') {
    const n = Math.min(items.length, 6);
    const w = (W + 30 * (n - 1)) / n;
    items.slice(0, n).forEach((it, i) => {
      push({ type: 'rect', shape: 'Chevron', left: L + (w - 30) * i, top: T,
        width: w, height: num(args.height, 80), fill: tone(i, n),
        text: String(it.label || ''), color: onBlock, size: 14, bold: true });
    });
  } else {
    /**
     * 文氏圖：三個圓，交集另外補。
     *
     * fill 設不了半透明，所以交集不會自己變色——要手算混色另外疊一塊上去。
     * 不補的話三個圓就只是互相遮住，看不出交集在哪裡。
     */
    const d = num(args.node, 220);
    const cx = L + W / 2;
    // 往下 18：上面兩個圓的標籤在圓內上緣，貼著標題區會顯得擠。
    const vt = T + 18;
    const pos = [[cx - d * 0.38, vt], [cx - d * 0.62 + d, vt], [cx - d / 2, vt + d * 0.55]];
    const cols = [tone(0, 3), tone(1, 3), tone(2, 3)];
    /**
     * 三個圓直接用半透明，交集自己就會疊出來。
     *
     * 0.72 是試出來的平衡：再實一點交集看不出來，再淡一點圓本身的顏色就站不住。
     * 文字用深色而不是白色——半透明的底撐不起白字。
     */
    pos.forEach((pt, i) => {
      push({ type: 'rect', shape: 'Ellipse', left: pt[0], top: pt[1], width: d, height: d,
        fill: cols[i], opacity: num(args.opacity, 0.72),
        text: String((items[i] && items[i].label) || ''),
        color: args.block_text || ink, size: 14, bold: true,
        valign: i === 2 ? 'bottom' : 'top' });
    });
    /**
     * 中心只放一塊，而且只在有字的時候。
     *
     * ── 第一版多畫了什麼 ──
     * 我原本在工具說明裡叫模型「交集要另外補混色塊」，結果它在每兩個圓的交會處都補了
     * 一塊——三個暗色小圓浮在圖上，看起來像沒對齊的瑕疵，而不是交集。
     *
     * 兩兩交集本來就不需要標示：三個圓重疊的形狀**本身就說明了交集在哪裡**。真正要
     * 指出來的只有正中央那一塊，因為那是整張圖的結論。沒有結論文字的話連它都不要畫。
     */
    if (args.center) {
      const mid = mixHex(mixHex(cols[0], cols[1], 0.5), cols[2], 0.34);
      push({ type: 'rect', shape: 'Ellipse', left: cx - 42, top: vt + d * 0.42,
        width: 84, height: 84, fill: mid,
        text: String(args.center), color: onBlock, size: 13, bold: true });
    }
  }

  return drawShapes({ ...args, shapes: plan, theme: args.theme, font: face });
}

async function drawShapes(args) {
  if (!hasApi('1.4')) throw new Error('批次繪製需要 PowerPointApi 1.4，這個 Office 版本太舊。');
  const items = Array.isArray(args.shapes) ? args.shapes.filter(Boolean) : [];
  if (!items.length) {
    throw new Error('shapes 是必要的：一個陣列，每個元素是一個物件。'
      + '例如 [{ type:"rect", left:30, top:104, width:490, height:220, fill:"#233559" },'
      + ' { type:"text", left:30, top:12, text:"標題", size:26, bold:true, color:"#FFFFFF" }]。');
  }
  if (items.length > 120) {
    throw new Error(`一次最多 120 個物件，這次有 ${items.length} 個。`
      + '再多的話一張投影片會慢到難以編輯——拆成兩次，或看看有沒有哪一塊其實該用'
      + ' ppt_add_table、ppt_add_cards 這些現成的。');
  }

  const th = theme(args.theme);
  const face = args.font || th.font;

  return PowerPoint.run(async (ctx) => {
    const { slide, index } = await targetSlide(ctx, args);
    let made = 0;
    const unknown = new Set();
    const failed = [];

    /**
     * 分批 sync，而且一個壞掉不拖垮整批。
     *
     * ── 為什麼 ──
     * Office.js 把排進去的操作一次送出，所以**其中任何一個參數錯了，整批都回不來**——
     * 使用者看到的是「畫版面沒成功」，六十個物件一個都沒上去，而真正出問題的可能只有
     * 一個顏色字串。
     *
     * 一張看板的價值在於它整張做完。一個格子畫不出來是瑕疵，整張沒畫出來是失敗，
     * 這兩件事差很遠。所以每 30 個送一次，而且單一物件的錯誤只記下來、不往外丟。
     */
    const CHUNK = 30;
    let pending = 0;
    const flush = async () => {
      if (!pending) return;
      try { await ctx.sync(); } catch (e) {
        failed.push(`第 ${made - pending + 1}~${made} 個：${(e && e.message) || e}`);
      }
      pending = 0;
    };

    for (const it of items) {
      try {
      const type = String(it.type || 'rect').toLowerCase();
      const geo = {
        left: num(it.left, 0),
        top: num(it.top, 0),
        width: num(it.width, type === 'line' ? 120 : 160),
        height: num(it.height, type === 'line' ? 0 : 40),
      };

      if (type === 'line') {
        const ln = slide.shapes.addLine('StraightConnector1', geo);
        ln.name = stamp('draw');
        try {
          ln.lineFormat.color = colour(it.color || it.fill, th.strong);
          ln.lineFormat.weight = num(it.weight, 2);
        } catch { /* 舊版給不了就用預設 */ }
        made += 1;
        pending += 1;
        if (pending >= CHUNK) await flush();
        continue;
      }

      /**
       * text 也是一個矩形，只是沒有填色。
       *
       * 用 addTextBox 的話它的寬高行為跟矩形不一樣（會自動縮），同一份版面裡兩種
       * 盒子模型混用，對齊就會差幾點——而看板最吃的就是對齊。
       */
      const asked = type === 'text' ? 'Rectangle' : (it.shape || type);
      const kind = geoType(asked === 'rect' ? 'Rectangle' : asked);
      if (it.shape && kind !== it.shape) unknown.add(String(it.shape));
      const sh = slide.shapes.addGeometricShape(kind, geo);
      sh.name = stamp('draw');

      if (type === 'text' && it.fill === undefined) sh.fill.clear();
      else if (it.fill === null || it.fill === '') sh.fill.clear();
      else sh.fill.setSolidColor(colour(it.fill, th.tint));
      /**
       * 半透明。
       *
       * ── 我當初說做不到，那是沒查就下的結論 ──
       * ShapeFill 有 transparency 這個屬性（0 是不透明、1 是全透明）。我沒試就在文氏圖
       * 那邊叫模型「交集另外補混色塊」，結果它在每兩個圓的交會處補了一塊暗色小圓，
       * 看起來像瑕疵——**為了繞過一個不存在的限制，產生了一個真的問題**。
       *
       * 這裡用 opacity（0~1，直覺是「不透明度」）而不是 transparency，因為設計的人
       * 想的是「這塊要多實」，不是「要多空」。沒給就是完全不透明，行為不變。
       *
       * 包 try：舊版 Office 給不了就維持實心——那時候圖還是畫得出來，只是沒有疊色。
       */
      if (it.opacity !== undefined) {
        const o = Math.max(0, Math.min(1, Number(it.opacity)));
        try { sh.fill.transparency = 1 - o; } catch { /* 舊版沒有，維持實心 */ }
      }
      try {
        if (it.line_color) {
          sh.lineFormat.color = colour(it.line_color);
          sh.lineFormat.weight = num(it.line_width, 1);
        } else {
          sh.lineFormat.visible = false;
        }
      } catch { /* 舊版沒有框線控制 */ }

      if (it.text !== undefined && it.text !== null && String(it.text) !== '') {
        sh.textFrame.textRange.text = String(it.text);
        styleText(sh, {
          font: it.font || face,
          size: num(it.size, 11),
          bold: it.bold === true,
          color: it.color || th.ink,
          align: it.align || (type === 'text' ? 'left' : 'center'),
          valign: it.valign || 'middle',
        });
        try {
          // 盒子就是盒子——不要因為字多就自己長大，那會毀掉整份版面的對齊。
          sh.textFrame.wordWrap = it.wrap !== false;
          sh.textFrame.autoSizeSetting = 'AutoSizeNone';
        } catch { /* 舊版沒有這些設定 */ }
      }
      made += 1;
      pending += 1;
      if (pending >= CHUNK) await flush();
      } catch (e) {
        // 這一個畫不出來就跳過，其餘照畫——整張做完比完美重要。
        failed.push(`${String(it.text || it.type || '物件').slice(0, 12)}：${(e && e.message) || e}`);
      }
    }

    await flush();
    return `已在投影片 ${index} 畫上 ${made - failed.length} 個物件`
      + (failed.length ? `，${failed.length} 個沒畫成` : '')
      + '。'
      + (failed.length
        ? `\n沒成功的：${failed.slice(0, 5).join('；')}`
          + (failed.length > 5 ? `（另外還有 ${failed.length - 5} 個）` : '')
          + `\n這幾個通常是顏色字串或座標的格式問題。**把它們重畫一次**——修正參數之後再呼叫 ppt_draw，只畫沒成功的那幾個就好。`
        : '')
      + (unknown.size
        ? `\n（PowerPoint 沒有「${[...unknown].join('、')}」這些形狀，改用了矩形。`
          + '圓形是 Ellipse，箭頭是 RightArrow、Chevron。）'
        : '')
      + '\n每一個都是獨立的物件，可以在 PowerPoint 裡直接點選拖動。'
      + '要整塊重做的話跟我說，我重畫比一個一個搬快。';
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

    // Before, not after. This is the whole reason 「AI 把我第 5 頁的字改掉了」
    // used to be unrecoverable: nobody wrote down what it said.
    rememberBefore(slide.id, await captureShapes(ctx, [shape]));

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

    rememberBefore(slide.id, await captureShapes(ctx, [shape]));

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

    // Read them before they are gone. What comes back is a description, not the
    // shape — good enough to rebuild a box with its text, colour and position,
    // and not good enough for a photograph. restoreSnapshot says which is which
    // rather than quietly restoring some of them.
    const gone = await captureShapes(ctx, uniq.map((n) => items[n - 1]));
    for (const before of gone) {
      if (before) undoOps.push({ op: 'recreate', slideId: slide.id, before });
    }

    uniq.slice().reverse().forEach((n) => items[n - 1].delete());
    await ctx.sync();
    const lost = gone.filter((b) => !REBUILDABLE.has(String(b.type || '').toLowerCase())).length;
    const caveat = lost ? `\n其中 ${lost} 個（圖片或群組）復原鍵拿不回來。` : '';
    return (uniq.length === 1
      ? `已刪除投影片 ${index} 的第 ${uniq[0]} 個物件。`
      : `已刪除投影片 ${index} 的 ${uniq.length} 個物件（第 ${uniq.join('、')} 個）。`) + caveat;
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

    // Every shape it is about to move, not just the ones that end up moving:
    // which ones move depends on the operation, and working that out twice is
    // how the two copies drift apart.
    rememberBefore(slide.id, await captureShapes(ctx, picked));

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
/**
 * The layouts a deck is built from.
 *
 * ── Why `section` and `statement` are no longer offered ──
 * They stayed in the tool's enum for weeks while the prompt told the model to
 * use them sparingly, and the model kept reaching for them anyway. That is not
 * the model being disobedient; it is what a prompt rule IS — advisory. When a
 * thing has to stop happening, the reliable lever is the schema, not the
 * paragraph: they are gone from the enum in pptToolSpec, so the model cannot
 * pick them.
 *
 * The code stays because the capability is not wrong, only the default. Putting
 * either back is one word in that enum.
 *
 * What replaces them, and honestly does the job better:
 *   · chapters → the `eyebrow` on every content slide. 「01 資本與成本」 sits on
 *     each page of the section rather than on one page before it, so somebody
 *     who joins halfway still knows where they are.
 *   · a big number → `kpi`, next to the content it came from, where it has a
 *     context instead of a whole slide to itself.
 */
/**
 * 封面版型留著——但它不是「一種」封面，是八種構圖。
 *
 * ── 我差點把它拿掉，然後發現不該 ──
 * 「每一份封面都長一樣」是真的，所以我一度把這個版型移掉，想逼它用 ppt_draw 自己畫。
 * 測試立刻指出代價：`cover` 參數底下有 panel / stack / bleed / triptych / split /
 * imageleft / imagetop 八種構圖，而它們全都掛在這個版型上。拿掉 title 等於一次砍掉
 * 八種封面，換來「每次重畫」——那是用更少的選擇去換更多的變化，不划算。
 *
 * 所以問題不是「沒有選擇」，是**模型從來不換**：八種構圖它每次都挑同一種。
 * 那要靠提示詞和規劃階段解決，不是靠刪掉能力。
 */
const RECIPES = new Set(['title', 'section', 'bullets', 'two_column', 'compare',
  'statement', 'quote', 'timeline', 'quadrant', 'zigzag', 'feature', 'list_image']);

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

  /**
   * The decorative layer, and why it cannot be drawn with the rest.
   *
   * Paper grain and a hand-drawn frame are SVG, and SVG reaches a slide through
   * `setSelectedDataAsync` — a Common API that APPENDS to the current selection.
   * Appending puts it on top, and a background on top of the content is not a
   * background. So it goes in first, in its own step, before the shape batch
   * that then lands above it.
   *
   * One document for the whole layer rather than one per ornament: each insert
   * is its own round trip, and a slide wanting paper plus a frame plus a rule
   * would otherwise pay for three.
   */
  let imageNote = '';
  let slotNote = '';
  /**
   * Which motif runs under this slide.
   *
   * A deck's background is a decision made once, so the look supplies it and the
   * model may override it per DECK — never per slide. The same marks in the same
   * places on every page is the entire source of the 「模板感」; varying it turns a
   * template back into a collage.
   */
  const motif = args.motif === undefined ? (lk.motif || 'none') : String(args.motif);
  const wantsArt = lk.texture || lk.strokes === 'hand' || (motif && motif !== 'none');
  const ornament = wantsArt && hasImageCoercion('1.2')
    ? ornamentSvg({
      texture: lk.texture,
      strokes: lk.strokes,
      motif,
      tone: th.strong,
      accent: th.accent,
      theme: th,
      ground: groundOf(th, lk),
      ink: th.strong,
      title: kind !== 'title' && kind !== 'section' && args.chrome !== false && !!args.title,
      frame: lk.masthead === 'frame' || lk.strokes === 'hand',
      footband: !!args.footband,
      topPad: lk.topPad,
    })
    : null;

  /**
   * A picture that is part of the composition rather than the whole of it.
   *
   * ── The complaint this answers ──
   * 「首頁都是產圖的版型…我寧願是多版型」, and behind it a sharper point: a
   * full-bleed photograph with a title on top is the LEAST striking thing a
   * cover can be unless the photograph is extraordinary. A picture holding a
   * quarter of the page, with type and colour doing the rest, is a composition;
   * a picture holding all of it is a background.
   *
   * So the image goes in as a band or a column and the recipe lays out around
   * it. It shares the ornament's path for the same reason — setSelectedDataAsync
   * appends, so anything that must sit UNDER the shapes has to go first.
   */
  const IMAGE_BOX = {
    imageleft: { left: 0, top: 0, width: 384, height: GRID.H },
    imageright: { left: GRID.W - 384, top: 0, width: 384, height: GRID.H },
    imagetop: { left: 0, top: 0, width: GRID.W, height: 208 },
  };
  /**
   * Where a content slide's picture goes.
   *
   * Unlike the cover forms these sit INSIDE the content band, below the title —
   * the page still has its heading, its eyebrow and its footer, and the picture
   * is one element on it rather than the ground it stands on.
   */
  const BODY_IMAGE = {
    feature: { left: CHROME.contentLeft, top: 132, width: 470, height: 306 },
    list_image: null,
  };
  const wantsImage = String(args.image || '').trim()
    && (IMAGE_BOX[String(args.cover || '').toLowerCase()]
      || BODY_IMAGE[String(args.kind || '').toLowerCase()]);

  /**
   * Everything that has to go UNDER the shapes, in the order it stacks.
   *
   * ── The bug this replaces ──
   * The ornament and the picture were two independent blocks, the second
   * guarded by `!ornament`. So on any look with a background motif — which is
   * most of them — a slide that asked for a picture silently got none: no
   * error, no note, just a cover with an empty column where the photograph was
   * supposed to be. The two are not alternatives; one is the ground and the
   * other sits on it.
   */
  if (ornament || wantsImage) {
    await PowerPoint.run(async (ctx) => {
      if (args.slide === undefined && args.append !== false && hasApi('1.3')) {
        ctx.presentation.slides.add();
        await ctx.sync();
      }
    });
    await focusSlideFor(args.slide === undefined ? {} : args);

    if (ornament) {
      await insertImageData(ornament, Office.CoercionType.XmlSvg,
        { left: 0, top: 0, width: GRID.W, height: GRID.H });
    }
    if (wantsImage) {
      try {
        const got = await fetchPhoto(String(args.image),
          wantsImage.width > wantsImage.height ? 'block' : 'side');
        await insertImageData(got.svg || got.data,
          got.svg ? Office.CoercionType.XmlSvg : Office.CoercionType.Image,
          // Illustrations and photographs alike keep their own proportions
          // here: this slot is a frame in a layout, not a background to fill.
          fitInside(wantsImage, naturalSize(got)));
      } catch (e) {
        // A slide that arrives without its picture is a slide. One that refuses
        // to be built because a CDN was slow is a broken tool.
        imageNote = `
（圖片沒抓到：${e.message} 版面照做，圖請稍後再補。）`;
      }
    }
  }

  return PowerPoint.run(async (ctx) => {
    // A recipe fills a slide, so it defaults to making a new one rather than
    // piling a second layout on top of whatever is already there.
    if (!ornament && !wantsImage && args.slide === undefined && args.append !== false && hasApi('1.3')) {
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

    /**
     * A bullet list as one text box PER LINE.
     *
     * ── Why not one box and a line-spacing property ──
     * That was the first attempt and on this machine it did nothing, because
     * `ParagraphFormat.lineSpacing` needs PowerPointApi 1.8 and most installs
     * are on 1.4–1.6. A feature that silently does nothing on the version people
     * actually run is not a feature; it is a comment.
     *
     * Geometry works on every build. One box per line, positioned on a pitch we
     * choose, gives exactly the air the property would have — and it gives more
     * besides: each line can be vertically centred in its own slot, so a point
     * that wraps to two lines still sits on the same rhythm as the ones that do
     * not. A single box cannot do that at any API version.
     *
     * The list is centred in its band rather than pinned to the top, because
     * three points in a 350pt band pinned to the top is the 「下半頁是空的」
     * complaint in its original form.
     */
    const bulletList = (lines, box, style = {}) => {
      const items = lines.filter((l) => String(l).trim());
      if (!items.length) return;
      const size = style.size || TYPE.body;
      // 2.1× the type size: enough that the lines read as separate points, not
      // so much that a four-point list turns into a page of gaps.
      const pitch = Math.min(style.pitch || size * 2.1, box.height / items.length);
      const used = pitch * items.length;
      const top = style.top === true
        ? box.top
        : box.top + Math.max(0, (box.height - used) / 2);
      items.forEach((line, i) => {
        text(`▪　${String(line).trim()}`,
          { left: box.left, top: top + pitch * i, width: box.width, height: pitch },
          { size, color: th.ink, valign: 'middle', ...style, pitch: undefined, top: undefined });
      });
    };

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
        // `quote` belongs in this group for the same reason `statement` does:
        // the quotation IS the title, so drawing both prints the sentence twice
        // — once at 25pt in the header and once at 32pt in the middle. 「標題跟
        // 大文字一樣」 was that, and it is the third place this exact mistake has
        // turned up. Any recipe whose main element is `title` goes here.
        : (kind === 'statement' || kind === 'quote') ? 'marks'
          : 'full';
    const dressed = frame !== 'none';
    // Everything below lays out against this, so indenting here is all the
    // rail look needs — no recipe has to know it exists.
    const RAIL = lk.masthead === 'rail' ? (lk.rail || 52) : 0;
    if (dressed) {
      // The page ground goes down FIRST so everything else lands on top of it.
      //
      // Off-white rather than white, and this is not a small difference: a pure
      // white slide has no page, so a white card on it has no edge and the
      // content reads as typed straight onto the projector. Two percent of grey
      // gives the cards something to be lighter than, and that is what makes a
      // layout look built rather than filled in.
      // Skipped when the ornament is already down there carrying the ground,
      // the texture and (for a hand look) the frame and the rule. Painting a
      // flat rectangle over it would hide the entire decorative layer.
      if (!ornament) band({ left: 0, top: 0, width: GRID.W, height: GRID.H }, groundOf(th, lk));

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
      } else if (lk.masthead === 'rail') {
        // A full-height column down the left edge. The content is indented past
        // it below — this is the only look that moves the content area, and
        // that is precisely why it does not read as a recolour of another one.
        band({ left: 0, top: 0, width: lk.rail || 52, height: GRID.H }, th.strong);
        band({ left: 0, top: 0, width: lk.rail || 52, height: 5 }, th.accent);
        if (args.page) {
          text(String(args.page), { left: 0, top: GRID.H - 44, width: lk.rail || 52, height: 24 },
            { size: 10, color: th.onCover, align: 'center', valign: 'middle' });
        }
      } else if (lk.masthead === 'frame' && !ornament) {
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
        text(label, {
          ...CHROME.eyebrow,
          left: lk.masthead === 'block' ? 16 : CHROME.eyebrow.left + RAIL,
          width: CHROME.eyebrow.width - RAIL,
          top: eyeY,
        },
          {
            size: 11,
            color: lk.masthead === 'block' ? th.onCover : onGround(th, lk, 'muted'),
            valign: 'middle',
            align: lk.titleAlign === 'center' ? 'center' : 'left',
          });
      }
      if (title && frame === 'full') {
        text(title, {
          ...CHROME.title,
          left: CHROME.title.left + RAIL,
          width: CHROME.title.width - RAIL,
          top: CHROME.title.top + lk.topPad,
        },
          {
            font: display,
            size: lk.titleSize, bold: lk.titleBold,
            color: onGround(th, lk, 'strong'),
            align: lk.titleAlign, valign: 'middle',
          });
        if (ornament && lk.strokes === 'hand') {
          // Already drawn, with a wobble. See ornamentSvg.
        } else if (lk.rule === 'short') {
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
    const CL = CHROME.contentLeft + RAIL;
    const CW = CHROME.contentWidth - RAIL;

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

      /**
       * The cover's subtitle, set with actual leading.
       *
       * ── Why it is not just a text box ──
       * A cover subtitle is one or two lines of Chinese at 22pt, and at single
       * spacing two lines of that touch. `ParagraphFormat.lineSpacing` would fix
       * it and needs PowerPointApi 1.8, which most installs do not have — so the
       * same trick the bullet lists use applies here: one box per line, on a
       * pitch we choose.
       *
       * When the model supplied several `body` entries they are already separate
       * lines, which is the good case. When it supplied one long sentence there
       * is nothing to split on, so it stays a single box — a wrapped line still
       * being tight is better than breaking a sentence at an arbitrary point.
       */
      const subLines = (lines, box, style = {}) => {
        const items = (lines || []).map((l) => String(l).trim()).filter(Boolean);
        if (!items.length) return;
        const size = style.size || TYPE.heading;
        const pitch = Math.min(style.pitch || size * 1.9, box.height / items.length);
        items.forEach((line, i) => {
          text(line, { left: box.left, top: box.top + pitch * i, width: box.width, height: pitch },
            { size, valign: 'middle', ...style, pitch: undefined });
        });
      };

      /**
       * Which composition, and why the look does not get the last word.
       *
       * A look names the cover it was designed around, and that is the right
       * default. But a deck's first page is the one slide where variety is worth
       * more than consistency — it is seen once, against nothing, and it sets
       * the expectation for everything after it. So the model may name a form
       * outright, and the prompt encourages it to: 「上下、左右、滿版、三欄」 are
       * different answers to the same page, not different levels of quality.
       */
      /**
       * Circles on a flat colour block, cut off by its edges.
       *
       * ── Why every plain panel gets them ──
       * A rectangle of one colour is not a design decision; it is the absence of
       * one, and it is what made the first covers read as 「黑黑一塊」. Two circles
       * in the same hue a step lighter cost two shapes and change the block from
       * a fill into a surface — and because they are cut off by the panel's own
       * edges, they imply the block continues past what you can see.
       *
       * Same hue rather than a second colour, deliberately: a contrasting circle
       * is an element the reader has to account for. This one is texture.
       */
      const orbs = (box, base) => {
        const halo = mix(base, th.onCover, 0.12);
        const draw = (geo) => {
          /**
           * Never a negative coordinate.
           *
           * ── What this cost before the clamp ──
           * The circles were positioned at `box.left - r * 0.34` to bleed off the
           * left edge, which is negative whenever the block starts at 0. Office
           * rejects that, and it rejects it at `ctx.sync()` — so the failure took
           * the ENTIRE cover with it, not just the circle. The agent's own note
           * to the user said it: 「封面原本想用左色塊＋圓形那種構圖，這台的該種構圖
           * 會失敗」. Nothing was wrong with the composition; two numbers were out
           * of range.
           *
           * The bleed still happens — it just happens off the BOTTOM, where the
           * coordinate stays positive and the overflow is in the width and
           * height instead.
           */
          const safe = {
            left: Math.max(0, geo.left), top: Math.max(0, geo.top),
            width: Math.max(1, geo.width), height: Math.max(1, geo.height),
          };
          const sh = slide.shapes.addGeometricShape(geoType('Ellipse'), safe);
          sh.name = stamp('orb');
          sh.fill.setSolidColor(halo);
          sh.lineFormat.visible = false;
          put.shape.push(sh);
        };
        const r = Math.min(box.width, box.height) * 0.82;
        // The large one sits in the block's bottom-left and runs off the bottom
        // of the page — the boundary is still crossed, with every coordinate
        // positive. The small one stays inside the block's right edge, because
        // on a 400pt panel anything past it would land on the white where the
        // title goes.
        draw({ left: box.left, top: box.top + box.height - r * 0.45, width: r, height: r });
        const r2 = r * 0.34;
        draw({ left: box.left + box.width - r2 * 1.25, top: box.top + box.height * 0.16,
          width: r2, height: r2 });
      };

      const FORMS = new Set(['bleed', 'split', 'band', 'plain', 'paper', 'panel',
        'stack', 'triptych', 'imageleft', 'imageright', 'imagetop']);
      const form = FORMS.has(String(args.cover || '').toLowerCase())
        ? String(args.cover).toLowerCase()
        : lk.cover;

      if (form === 'split') {
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
        subLines(body, { left: 96, top: 320, width: 740, height: 56 },
          { size: TYPE.heading, color: th.muted });
        if (args.footer) {
          text(String(args.footer), { left: 96, top: 456, width: 740, height: 24 },
            { size: TYPE.small, color: th.muted });
        }
      } else if (form === 'band') {
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
        subLines(body, { left: 104, top: 310, width: GRID.W - 208, height: 56 },
          { size: TYPE.heading, color: th.onCover, align: 'center' });
        if (args.footer) {
          text(String(args.footer), { left: 64, top: 432, width: GRID.W - 128, height: 24 },
            { size: TYPE.small, color: th.muted, align: 'center' });
        }
      } else if (form === 'panel') {
        /**
         * A colour block down one side, with shapes that run off the edge.
         *
         * ── Why the old cover had to go ──
         * 「企業風的 PPT 他會像圖片這樣，就是黑黑一塊」. It was: a flat rectangle of
         * navy with type in the middle of it. Nothing was wrong with any single
         * decision and the result was a coloured slab.
         *
         * What makes a real cover look made is that something CROSSES A BOUNDARY
         * — a circle half off the page, a rule that starts outside the margin.
         * That reads as a composition somebody positioned; type centred inside a
         * rectangle reads as type that had nowhere else to go. So the circles
         * below are deliberately placed to bleed off the left and bottom edges,
         * and the accent stripe runs the full height rather than stopping
         * politely at the margins.
         */
        band({ left: 0, top: 0, width: GRID.W, height: GRID.H }, th.bg);
        band({ left: 0, top: 0, width: 400, height: GRID.H }, th.cover);

        orbs({ left: 0, top: 0, width: 400, height: GRID.H }, th.cover);

        band({ left: 400, top: 0, width: 9, height: GRID.H }, th.accent);

        /**
         * Two lines in the panel, or one — never the same line twice.
         *
         * The bold line falls back to the eyebrow so the panel is never an empty
         * block of colour, but the small line above it was drawn regardless. So
         * a cover with an eyebrow and no footer printed 「2026 TREND REPORT」 at
         * 11pt and again at 18pt, one under the other. Same shape of mistake as
         * section, statement and quote: one value rendered by two parts of the
         * layout that did not know about each other.
         *
         * 18pt rather than 22: 「McKinsey / Gartner / MIT / PwC 公開研究」 wraps
         * inside 300 points and at 22 the two lines touch.
         */
        const org = String(args.footer || args.eyebrow || '');
        const kicker = args.eyebrow && String(args.eyebrow) !== org ? String(args.eyebrow) : '';
        if (kicker) {
          text(kicker, { left: 62, top: 96, width: 300, height: 22 },
            { size: TYPE.small, color: mix(th.cover, th.onCover, 0.72) });
        }
        text(org, { left: 62, top: kicker ? 120 : 104, width: 300, height: 56 },
          { font: display, size: 18, bold: true, color: th.onCover, valign: 'middle' });
        band({ left: 62, top: 178, width: 96, height: 4 }, th.accent);

        text(title, { left: 470, top: 128, width: 430, height: 132 },
          { font: display, size: 34, bold: true, color: th.strong, valign: 'middle' });
        band({ left: 470, top: 274, width: 140, height: 5 }, th.accent);
        subLines(body, { left: 470, top: 296, width: 430, height: 84 },
          { size: TYPE.body, color: th.muted });
        if (args.page) {
          text(String(args.page), { left: 470, top: 456, width: 430, height: 22 },
            { size: TYPE.caption, color: th.muted });
        }
      } else if (form === 'imageleft' || form === 'imageright') {
        /**
         * 左右 — a picture down one side, everything else on the other.
         *
         * The picture is 40% of the page, so it has to be doing a job rather
         * than filling space, and the type sits on a plain ground where it is
         * unambiguously readable. This is the composition to reach for when the
         * subject HAS a picture; the fully abstract ones below are for when it
         * does not.
         */
        const onRight = form === 'imageright';
        /**
         * Only the text half gets painted.
         *
         * ── The bug ──
         * This used to lay a full-bleed `bg` rectangle down first, the way the
         * other covers do. But the picture is inserted BEFORE the shape batch —
         * it has to be, so it sits underneath — and a full-bleed rectangle drawn
         * afterwards covers it completely. 「圖有出來，但被另一個白色的區塊蓋住了」
         * was exactly that: the image was there, under a sheet of white.
         *
         * So the ground stops at the picture's edge. The tint is doing real work
         * too — it makes the half without the picture a surface rather than a
         * blank — but it must not cross the seam.
         */
        band({ left: onRight ? 0 : 392, top: 0, width: GRID.W - 392, height: GRID.H }, th.ground);
        // A thin bar of colour against the picture's edge, so the two halves are
        // joined rather than merely adjacent.
        band({ left: onRight ? GRID.W - 392 : 384, top: 0, width: 8, height: GRID.H }, th.accent);
        const tx = onRight ? 72 : 456;
        const tw = 432;
        if (args.eyebrow) {
          text(String(args.eyebrow).toUpperCase(), { left: tx, top: 128, width: tw, height: 24 },
            { size: TYPE.small, color: th.accent });
        }
        text(title, { left: tx, top: 162, width: tw, height: 132 },
          { font: display, size: 34, bold: true, color: th.strong, valign: 'middle' });
        band({ left: tx, top: 306, width: 120, height: 5 }, th.accent);
        subLines(body, { left: tx, top: 328, width: tw, height: 84 },
          { size: TYPE.body, color: th.muted });
        if (args.footer) {
          text(String(args.footer), { left: tx, top: 452, width: tw, height: 24 },
            { size: TYPE.caption, color: th.muted });
        }
      } else if (form === 'imagetop') {
        /**
         * 上下 — a strip of picture across the top, the page below it.
         *
         * A band rather than a background: 208 points of 540, so it reads as a
         * window onto the subject and the rest of the slide is still a designed
         * surface. The colour bar under it is what stops the join looking like a
         * crop.
         */
        // Below the picture only — see imageLeft for what a full-bleed ground
        // does to an image that was inserted before the shapes.
        band({ left: 0, top: 216, width: GRID.W, height: GRID.H - 216 }, th.bg);
        band({ left: 0, top: 208, width: GRID.W, height: 8 }, th.accent);
        if (args.eyebrow) {
          text(String(args.eyebrow).toUpperCase(), { left: 80, top: 252, width: 800, height: 24 },
            { size: TYPE.small, color: th.accent });
        }
        text(title, { left: 80, top: 286, width: 800, height: 104 },
          { font: display, size: 38, bold: true, color: th.strong, valign: 'middle' });
        band({ left: 80, top: 398, width: 130, height: 5 }, th.accent);
        subLines(body, { left: 80, top: 418, width: 800, height: 56 },
          { size: TYPE.body, color: th.muted });
        if (args.footer) {
          text(String(args.footer), { left: 80, top: 492, width: 800, height: 22 },
            { size: TYPE.caption, color: th.muted });
        }
      } else if (form === 'stack') {
        /**
         * 上下 — a colour band across the top two thirds, type below it.
         *
         * The one element that makes it a composition rather than two stacked
         * rectangles is the rule that CROSSES the boundary: it starts inside the
         * colour and ends on the white. Without it the eye reads a coloured strip
         * with a caption; with it, one page.
         */
        band({ left: 0, top: 0, width: GRID.W, height: GRID.H }, th.bg);
        band({ left: 0, top: 0, width: GRID.W, height: 330 }, th.cover);
        orbs({ left: 0, top: 0, width: GRID.W, height: 330 }, th.cover);
        if (args.eyebrow) {
          text(String(args.eyebrow).toUpperCase(), { left: 80, top: 74, width: 700, height: 24 },
            { size: TYPE.small, color: mix(th.cover, th.onCover, 0.7) });
        }
        text(title, { left: 80, top: 116, width: 800, height: 150 },
          { font: display, size: 40, bold: true, color: th.onCover, valign: 'middle' });
        // Crosses from the colour onto the page.
        band({ left: 80, top: 302, width: 10, height: 56 }, th.accent);
        subLines(body, { left: 110, top: 372, width: 760, height: 64 },
          { size: TYPE.heading, color: th.ink });
        if (args.footer) {
          text(String(args.footer), { left: 110, top: 452, width: 760, height: 24 },
            { size: TYPE.small, color: th.muted });
        }
      } else if (form === 'triptych') {
        /**
         * 三欄 — three full-height columns of one hue, the title laid across all
         * of them.
         *
         * Same principle again, turned ninety degrees: the type crosses two
         * boundaries, so the columns read as one surface rather than as three
         * panels that happen to be adjacent. Cheap — five shapes — and it looks
         * nothing like any other cover in the set.
         */
        const w = GRID.W / 3;
        band({ left: 0, top: 0, width: w, height: GRID.H }, th.cover);
        band({ left: w, top: 0, width: w, height: GRID.H }, mix(th.cover, th.onCover, 0.1));
        band({ left: w * 2, top: 0, width: w, height: GRID.H }, mix(th.cover, th.onCover, 0.2));
        band({ left: 0, top: 0, width: GRID.W, height: 8 }, th.accent);
        if (args.eyebrow) {
          text(String(args.eyebrow).toUpperCase(), { left: 80, top: 150, width: 800, height: 24 },
            { size: TYPE.small, color: mix(th.cover, th.onCover, 0.75) });
        }
        text(title, { left: 80, top: 186, width: 800, height: 130 },
          { font: display, size: 40, bold: true, color: th.onCover, valign: 'middle' });
        band({ left: 80, top: 330, width: 130, height: 5 }, th.accent);
        subLines(body, { left: 80, top: 352, width: 800, height: 60 },
          { size: TYPE.heading, color: mix(th.cover, th.onCover, 0.82) });
        if (args.footer) {
          text(String(args.footer), { left: 80, top: 456, width: 800, height: 24 },
            { size: TYPE.small, color: mix(th.cover, th.onCover, 0.65) });
        }
      } else if (form === 'paper') {
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
        subLines(body, { left: 88, top: 338, width: 784, height: 56 },
          { font: display, size: TYPE.heading, color: th.ink });
        if (args.footer) {
          text(String(args.footer), { left: 88, top: 436, width: 784, height: 24 },
            { size: TYPE.caption, color: th.muted });
        }
      } else if (form === 'plain') {
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
        subLines(body, { left: 96, top: 326, width: 768, height: 56 },
          { size: TYPE.heading, color: th.muted });
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
        orbs({ left: 0, top: 0, width: GRID.W, height: GRID.H }, th.cover);
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
        subLines(body, { left: 80, top: 326, width: 780, height: 64 },
          { size: TYPE.heading, color: th.onCover });
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
      // Pulled up to sit directly under the chrome. It used to start at 150,
      // which left a band of nothing between the eyebrow and the panel — and a
      // page whose top third is empty reads as unfinished however big the number
      // in the middle of it is.
      band({ left: 0, top: 120, width: GRID.W, height: 292 }, th.tint);
      band({ left: 0, top: 120, width: GRID.W, height: 3 }, th.accent);
      text(title, { left: CL, top: 148, width: CW, height: 126 },
        { font: display, size: 84, bold: true, color: th.strong, align: 'center', valign: 'middle' });
      band({ left: (GRID.W - 120) / 2, top: 288, width: 120, height: 4 }, th.accent);
      if (body.length) {
        // The sentence is the half that stops it being a quiz, so it is set at
        // reading size in the ink — not as a grey caption under a big number.
        //
        // Narrower than the content band on purpose: a line of Chinese running
        // the full 884 points is too long a measure to scan, and the version
        // that wrapped to three ragged centred lines was the one people called
        // 「擠在一起」.
        text(body.join('\n'), { left: 170, top: 304, width: GRID.W - 340, height: 96 },
          { size: TYPE.body, color: th.ink, align: 'center', lineSpacing: 1.45 });
      }
    } else if (kind === 'timeline') {
      /**
       * A timeline as the page's主視覺, not as a chart dropped onto it.
       *
       * ── Why this is a layout and not just ppt_add_chart ──
       * The chart version draws a timeline into a box. This IS the page: the
       * axis runs the full width, the stops alternate above and below, and each
       * one gets room for a real sentence rather than a label. A roadmap slide
       * is the commonest thing a deck needs and the commonest thing to render as
       * a bulleted list, which throws away the only information a roadmap has —
       * that these things are in an order and some of them are further apart.
       */
      const stops = (Array.isArray(args.items) ? args.items : body.map((t) => ({ label: t })))
        .filter(Boolean).slice(0, 6);
      const midY = CT + BODY_H / 2;
      band({ left: CL, top: midY - 2, width: CW, height: 4 }, th.line);
      const step = stops.length > 1 ? CW / (stops.length - 1) : 0;
      stops.forEach((raw, i) => {
        const item = typeof raw === 'string' ? { label: raw } : (raw || {});
        const cx = stops.length > 1 ? CL + step * i : CL + CW / 2;
        const up = i % 2 === 0;
        const tone = i === stops.length - 1 ? th.accent : th.strong;
        // The dot sits ON the axis, so it is positioned by its centre.
        const dot = item.highlight === true || i === stops.length - 1 ? 18 : 13;
        band({ left: cx - dot / 2, top: midY - dot / 2, width: dot, height: dot }, tone);
        const boxW = Math.min(step * 0.92 || 200, 210);
        const boxH = 96;
        const by = up ? midY - 34 - boxH : midY + 34;
        // Clamped so the first and last cards stay on the page — a timeline's
        // ends are exactly where an evenly divided width puts them off it.
        const bx = clamp(cx - boxW / 2, CL, CL + CW - boxW);
        // Same reasoning as zigzag: without a container each stop is three
        // loose lines of text and the reader has to group them by proximity.
        const card = slide.shapes.addGeometricShape(boxOf(lk, 'Rectangle'),
          { left: bx, top: by, width: boxW, height: boxH });
        card.name = stamp('tlcard');
        card.fill.setSolidColor(th.bg);
        card.lineFormat.visible = true;
        card.lineFormat.color = th.line;
        card.lineFormat.weight = 1;
        put.shape.push(card);
        // A cap in the stop's own colour, so the card and its dot are the pair.
        band(capBox(boxOf(lk, 'Rectangle'), { left: bx, top: up ? by + boxH - 3 : by, width: boxW, height: 3 }, false), tone);
        if (item.at) {
          text(String(item.at), { left: bx, top: up ? by + boxH - 26 : by, width: boxW, height: 24 },
            { size: TYPE.small, bold: true, color: tone, align: 'center' });
        }
        text(String(item.label || ''),
          { left: bx, top: up ? by + boxH - 60 : by + 26, width: boxW, height: 34 },
          { size: 16, bold: true, color: th.ink, align: 'center', valign: 'middle' });
        if (item.note) {
          text(String(item.note),
            { left: bx, top: up ? by : by + 60, width: boxW, height: 36 },
            { size: 11.5, color: th.muted, align: 'center', lineSpacing: 1.25 });
        }
      });
    } else if (kind === 'quadrant') {
      /**
       * Two axes, four boxes.
       *
       * The layout for 「這兩個維度交叉之後有四種情況」 — priority matrices, risk
       * grids, build-vs-buy. Rendered as four bullets it loses the thing that
       * makes it a matrix: WHERE each item sits relative to the other three.
       *
       * The axis labels are the half people forget. A quadrant chart with four
       * captioned boxes and no axes is just four boxes.
       */
      const q = (Array.isArray(args.items) ? args.items : []).slice(0, 4);
      // The grid gives up a strip at the top and the bottom so the axis labels
      // have somewhere to stand that is not on top of a cell.
      const qTop = CT + 8;
      const qH = BODY_H - 30;
      const gx = CL + CW / 2;
      const gy = qTop + qH / 2;
      const cellW = CW / 2 - 14;
      const cellH = qH / 2 - 14;
      const tones = [th.strong, th.second, th.accent, th.muted];
      [[CL, qTop], [gx + 14, qTop], [CL, gy + 14], [gx + 14, gy + 14]].forEach((pos, i) => {
        const item = q[i] || {};
        const tone = colour(item.color, tones[i]);
        band({ left: pos[0], top: pos[1], width: cellW, height: cellH }, th.tint);
        band(capBox(boxOf(lk, 'Rectangle'), { left: pos[0], top: pos[1], width: cellW, height: 4 }, false), tone);
        text(String(item.title || ''),
          { left: pos[0] + 18, top: pos[1] + 18, width: cellW - 36, height: 28 },
          { size: TYPE.heading, bold: true, color: tone });
        if (item.text) {
          text(String(item.text),
            { left: pos[0] + 18, top: pos[1] + 52, width: cellW - 36, height: cellH - 70 },
            { size: 13.5, color: th.ink, valign: 'top', lineSpacing: 1.35 });
        }
      });
      // The axes, drawn over the cells so the cross reads as one mark.
      band({ left: gx - 1, top: qTop, width: 2, height: qH }, th.line);
      band({ left: CL, top: gy - 1, width: CW, height: 2 }, th.line);
      /**
       * The axis labels, and why they were unreadable.
       *
       * They were set at caption size in a 20pt strip squeezed against the edge
       * of the content band — 「太小了，被擠在一邊看不清楚」. An axis label is not
       * a footnote: it is the half of a matrix that says what the four boxes
       * MEAN, so it is set at the same weight as a card heading.
       *
       * Office.js cannot rotate a text box, so the vertical axis cannot run up
       * the side. It gets its own line above the grid instead, with an arrow to
       * carry the direction the rotation would have shown.
       */
      if (args.y_axis) {
        text(String(args.y_axis), { left: CL, top: qTop - 26, width: CW / 2, height: 22 },
          { size: 13, bold: true, color: th.strong, valign: 'middle' });
      }
      if (args.x_axis) {
        // Below the grid and ranged right: an x-axis reads left to right, so its
        // label belongs at the end of the run, not centred under it.
        text(String(args.x_axis),
          { left: CL + CW / 2, top: qTop + qH + 4, width: CW / 2, height: 22 },
          { size: 13, bold: true, color: th.strong, align: 'right', valign: 'middle' });
      }
    } else if (kind === 'zigzag') {
      /**
       * 圖文交錯 — alternating blocks, each with its side flipped.
       *
       * Three points down a page all aligned left is a list wearing panels. The
       * same three with the marker alternating left and right is a rhythm, and
       * the eye tracks it without being told to. Costs nothing: the geometry is
       * the same, mirrored on odd rows.
       */
      const rows = (Array.isArray(args.items) ? args.items : body.map((t) => ({ title: t })))
        .filter(Boolean).slice(0, 4);
      const gap = 14;
      const rowH = (BODY_H - gap * (rows.length - 1)) / Math.max(1, rows.length);
      const markW = Math.min(180, CW * 0.26);
      rows.forEach((raw, i) => {
        const item = typeof raw === 'string' ? { title: raw } : (raw || {});
        const y = CT + (rowH + gap) * i;
        const right = i % 2 === 1;
        const tone = colour(item.color, [th.strong, th.second, th.accent, th.warn || th.second][i % 4]);
        const markX = right ? CL + CW - markW : CL;
        // Inside the panel now, so the padding is the panel's rather than a gap.
        const textX = (right ? CL : CL + markW) + 22;
        const textW = CW - markW - 44;

        /**
         * A panel behind the words, not only behind the number.
         *
         * Without it the text floats on the page and the eye has to work out
         * which paragraph belongs to which marker — 「文字散落各處」. The panel
         * costs one shape per row and makes the pairing unambiguous.
         *
         * Tinted rather than outlined, with the marker's colour as a hairline:
         * four outlined boxes in a column read as a form to fill in.
         */
        const panel = slide.shapes.addGeometricShape(boxOf(lk, 'Rectangle'),
          // Starts where the marker ends, not where the marker starts — on a
          // left-marker row those are the same point and the panel would sit
          // on top of the number.
          { left: right ? CL : CL + markW, top: y, width: CW - markW, height: rowH });
        panel.name = stamp('zzpanel');
        panel.fill.setSolidColor(th.tint);
        panel.lineFormat.visible = true;
        panel.lineFormat.color = th.line;
        panel.lineFormat.weight = 1;
        put.shape.push(panel);

        band({ left: markX, top: y, width: markW, height: rowH }, tone);
        text(String(item.number || String(i + 1).padStart(2, '0')),
          { left: markX, top: y + rowH / 2 - 26, width: markW, height: 52 },
          { font: display, size: 38, bold: true, color: th.onCover, align: 'center', valign: 'middle' });

        text(String(item.title || ''),
          { left: textX, top: y + (item.text ? 10 : rowH / 2 - 18), width: textW, height: 32 },
          { size: TYPE.heading, bold: true, color: th.ink,
            align: right ? 'right' : 'left', valign: 'middle' });
        if (item.text) {
          text(String(item.text),
            { left: textX, top: y + 44, width: textW, height: rowH - 52 },
            { size: 14, color: th.muted, lineSpacing: 1.35,
              align: right ? 'right' : 'left', valign: 'top' });
        }
      });
    } else if (kind === 'feature') {
      /**
       * 大圖配小卡 — one picture holding the left, a short stack of cards on the
       * right.
       *
       * ── What this is for ──
       * The page where one thing is the subject and everything else is what you
       * need to know about it: a product, a site, a case. A bulleted list next
       * to a caption cannot do it, because the picture has to be big enough to
       * be looked AT rather than referred to.
       *
       * The picture is inserted before this batch (see wantsImage) and keeps its
       * own proportions, so the panel below is a frame for whatever shape turns
       * up — not a box the picture was stretched into.
       */
      const rows = (Array.isArray(args.items) ? args.items : body.map((t) => ({ title: t })))
        .filter(Boolean).slice(0, 4);
      const imgW = 470;
      const gap = 26;
      const rx = CL + imgW + gap;
      const rw = CW - imgW - gap;

      // A hairline frame around the picture's slot. With an image it reads as a
      // mount; without one it is an honest empty frame rather than a hole.
      const slot = slide.shapes.addGeometricShape(boxOf(lk, 'Rectangle'),
        { left: CL, top: CT + 8, width: imgW, height: Math.max(120, BODY_H - 16) });
      slot.name = stamp('featslot');
      slot.fill.clear();
      slot.lineFormat.visible = true;
      slot.lineFormat.color = th.line;
      slot.lineFormat.weight = 1;
      put.shape.push(slot);

      const rowH = rows.length
        ? (BODY_H - 8 - 12 * (rows.length - 1)) / rows.length
        : BODY_H;
      rows.forEach((raw, i) => {
        const item = typeof raw === 'string' ? { title: raw } : (raw || {});
        const y = CT + 8 + (rowH + 12) * i;
        const tone = colour(item.color,
          [th.strong, th.second, th.accent, th.warn || th.second][i % 4]);
        band({ left: rx, top: y, width: rw, height: rowH }, th.tint);
        band(capBox(boxOf(lk, 'Rectangle'),
          { left: rx, top: y, width: 5, height: rowH }, true), tone);
        text(String(item.title || ''),
          { left: rx + 20, top: y + (item.text ? 12 : rowH / 2 - 16), width: rw - 36, height: 30 },
          { size: TYPE.heading, bold: true, color: th.ink, valign: 'middle' });
        if (item.text) {
          text(String(item.text),
            { left: rx + 20, top: y + 46, width: rw - 36, height: rowH - 56 },
            { size: 13.5, color: th.muted, valign: 'top', lineSpacing: 1.3 });
        }
      });
    } else if (kind === 'list_image') {
      /**
       * 圖文列表 — a row per item, each with its own small picture.
       *
       * ── Why the pictures are drawn as slots rather than fetched ──
       * Four rows means four images, and four `setSelectedDataAsync` calls are
       * four round trips plus four downloads before the slide exists. So this
       * lays out the frames and returns their coordinates; the model fills them
       * with `ppt_add_photo` afterwards, one at a time, and the page is readable
       * the whole way through rather than blank until the last one lands.
       */
      const rows = (Array.isArray(args.items) ? args.items : body.map((t) => ({ title: t })))
        .filter(Boolean).slice(0, 4);
      const gap = 12;
      const rowH = rows.length
        ? (BODY_H - gap * (rows.length - 1)) / rows.length
        : BODY_H;
      const thumbW = Math.min(150, rowH * 1.5);
      rows.forEach((raw, i) => {
        const item = typeof raw === 'string' ? { title: raw } : (raw || {});
        const y = CT + (rowH + gap) * i;
        const tone = colour(item.color,
          [th.strong, th.second, th.accent, th.warn || th.second][i % 4]);
        band({ left: CL, top: y, width: CW, height: rowH }, th.tint);
        band(capBox(boxOf(lk, 'Rectangle'),
          { left: CL, top: y, width: 5, height: rowH }, true), tone);

        const slot = slide.shapes.addGeometricShape(boxOf(lk, 'Rectangle'),
          { left: CL + 18, top: y + 10, width: thumbW, height: rowH - 20 });
        slot.name = stamp('slot');
        slot.fill.setSolidColor(th.bg);
        slot.lineFormat.visible = true;
        slot.lineFormat.color = th.line;
        slot.lineFormat.weight = 1;
        put.shape.push(slot);

        const tx = CL + 18 + thumbW + 20;
        const tw = CW - (thumbW + 56);
        text(String(item.title || ''),
          { left: tx, top: y + (item.text ? 12 : rowH / 2 - 16), width: tw, height: 30 },
          { size: TYPE.heading, bold: true, color: th.ink, valign: 'middle' });
        if (item.text) {
          text(String(item.text),
            { left: tx, top: y + 44, width: tw, height: rowH - 54 },
            { size: 13.5, color: th.muted, valign: 'top', lineSpacing: 1.3 });
        }
      });
      // The coordinates go back in the tool result so the next call can fill them.
      slotNote = `\n圖片位置（用 ppt_add_photo，mode: "block"）：`
        + rows.map((_, i) => {
          const y = Math.round(CT + (rowH + gap) * i) + 10;
          return `第 ${i + 1} 列 left ${CL + 18}, top ${y}, `
            + `width ${Math.round(thumbW)}, height ${Math.round(rowH - 20)}`;
        }).join('；');
    } else if (kind === 'quote') {
      // Pulled up and stretched now that the header no longer carries a title:
      // the quotation starts where the content band starts, so the page has no
      // dead strip above it.
      /**
       * A quotation, given a page.
       *
       * ── What was thin about the first version ──
       * A hairline down the left, the sentence, a dash and a name. Correct and
       * unmemorable — on a page with this much air, four elements at low
       * contrast read as a slide that did not finish loading.
       *
       * Three changes, each cheap: a tinted field so the quotation sits ON
       * something; the opening mark large and overlapping the field's corner,
       * which is what makes it decoration rather than punctuation; and the
       * attribution on its own rule below, so it reads as a credit instead of a
       * trailing line of the quote.
       */
      const qH = Math.max(170, BODY_H - 78);
      band({ left: CL, top: CT, width: CW, height: qH }, th.tint);
      band({ left: CL, top: CT, width: 6, height: qH }, th.accent);
      text('“', { left: CL + 26, top: CT - 26, width: 120, height: 120 },
        { font: display, size: 120, bold: true, color: mix(th.tint, th.strong, 0.22) });
      text(title, { left: CL + 84, top: CT + 30, width: CW - 140, height: qH - 60 },
        { font: display, size: TYPE.title, color: th.ink, valign: 'middle' });
      if (body.length) {
        band({ left: CL + 84, top: CT + qH + 22, width: 48, height: 2 }, th.accent);
        text(body.join('，'),
          { left: CL + 84, top: CT + qH + 32, width: CW - 140, height: 30 },
          { size: TYPE.small, color: th.muted });
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
        // Top-aligned rather than centred: two columns whose lists are each
        // centred in their own panel end up starting at different heights when
        // one side has four points and the other has three, and the eye reads
        // that as a mistake.
        bulletList(c.lines,
          { left: c.x + 18, top: y + 10, width: colW - 36, height: CT + colH - y - 22 },
          { align: 'left', top: true });
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
              { size: 13.5, color: th.muted, valign: 'top', lineSpacing: 1.3 });
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
          bulletList(items.slice(0, half), { left: CL, top: CT, width: colW, height: BODY_H });
          bulletList(items.slice(half),
            { left: CL + colW + GRID.gutter, top: CT, width: colW, height: BODY_H });
        } else {
          bulletList(items, { left: CL, top: CT, width: CW, height: BODY_H });
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

    return `已用「${kind}」版型做好投影片 ${index}，${all.length} 個物件，套用「${th.label}」配色。${imageNote}${slotNote}\n`
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
      const sh = slide.shapes.addGeometricShape(geoType(opts.shape || 'Rectangle'), geo);
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
    pen.rect(capBox(boxOf(lk, 'Rectangle'), { left: x, top: box.top, width: w, height: 4 }, false),
      { tag: 'kpibar', fill: tone });

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
  pen.rect(capBox(boxOf(lk, 'Rectangle'), { left: box.left, top: box.top, width: 4, height: box.height }, true),
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
    { font: face, size: 12, color: th.ink, valign: 'middle', lineSpacing: 1.3, tag: 'calloutbody' });
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

/**
 * How wide one card will be, needed BEFORE the height can be estimated.
 *
 * The real width is computed further down from the same inputs; this repeats
 * that arithmetic rather than reordering the function, because the alternative
 * is a `let` that is assigned in two places and read in between.
 */
function cwEstimate(args, cards) {
  const bar = String(args.bar || (cards.length > 3 ? 'top' : 'left')).toLowerCase();
  const cols = clamp(num(args.columns, bar === 'top' ? cards.length : Math.min(2, cards.length)), 1, 8);
  const gutter = num(args.gutter, 24);
  const totalW = num(args.width, CHROME.contentWidth);
  return (totalW - gutter * (cols - 1)) / cols;
}

/**
 * 一張表格，用方塊拼出來的。
 *
 * ── 為什麼這個工具非有不可 ──
 * 「他 PPT 100% 都長這樣，上面 KPI 下放 2 張圖，沒了」——查下去才發現原因不是它沒創意，
 * 是**清單裡根本沒有表格**。使用者拿來參考的兩張好投影片都是表格撐起來的（落差分析表、
 * 檢驗站對照表、時程表），而增益集做不出來，只好退回它有的東西：卡片和圖表。
 *
 * 這是能力缺口，不是提示詞能解決的。再怎麼叫它「多樣化設計」，它手上沒有那個零件。
 *
 * ── 為什麼用方塊拼而不用 PowerPoint 的表格物件 ──
 * PowerPointApi 沒有建立表格的方法（只讀得到既有的）。方塊反而更可控：每一格的底色、
 * 字色、對齊都能單獨設，做得出「落差為負數標紅字」「最後一列是平均、底色再深一階」
 * 這種真正讓表格好讀的細節——那正是參考圖裡最講究的地方。
 *
 * 代價是使用者不能用 Word 那種表格工具改它，要直接點進某一格改字。回報訊息會講。
 */
async function addTable(args) {
  if (!hasApi('1.4')) throw new Error('表格需要 PowerPointApi 1.4，這個 Office 版本太舊。');
  const rows = Array.isArray(args.rows) ? args.rows.filter(Array.isArray) : [];
  if (!rows.length) throw new Error('rows 是必要的，二維陣列。第一列通常是欄名。');
  const cols = Math.max(...rows.map((r) => r.length));
  if (cols > 8) throw new Error('一張表最多 8 欄——再多在投影片上每欄都窄到讀不了。');
  if (rows.length > 12) throw new Error('一張表最多 12 列。超過的話那是文件不是投影片，'
    + '挑重點或拆成兩頁。');

  const lk = look(args.look);
  const th = paletteFor(args, lk);
  const face = lk.font || th.font;

  const left = num(args.left, CHROME.contentLeft);
  const top = num(args.top, CHROME.contentTop);
  const width = num(args.width, CHROME.contentWidth);
  const rowH = clamp(num(args.row_height, 26), 16, 60);
  // 列與列之間留一點縫：格線用「看得見的底色邊界」做，比畫線乾淨，而且縫寬可控。
  const gap = clamp(num(args.gap, 1), 0, 4);

  /**
   * 欄寬：第一欄是標籤，通常比較長；其餘平分。
   *
   * 給了 widths 就照給的（比例或點數都接受），因為有些表的第三欄才是重點。
   */
  const given = Array.isArray(args.widths) ? args.widths.map(Number).filter((n) => n > 0) : [];
  let widths;
  if (given.length === cols) {
    const sum = given.reduce((a, b) => a + b, 0);
    widths = given.map((w) => (w / sum) * width);
  } else {
    const firstW = Math.min(width * 0.34, Math.max(width / cols, 90));
    const rest = (width - firstW) / Math.max(1, cols - 1);
    widths = Array.from({ length: cols }, (_, c) => (c === 0 ? firstW : rest));
  }
  const xOf = (c) => left + widths.slice(0, c).reduce((a, b) => a + b, 0);

  const hasHeader = args.header !== false;
  // 最後一列是合計／平均的話再深一階——它是結論，不是又一筆資料。
  const totalRow = args.total_row === true ? rows.length - 1 : -1;
  /** 哪幾欄是數字：靠右對齊，而且用等寬數字才對得齊。 */
  const numeric = new Set((Array.isArray(args.numeric_columns) ? args.numeric_columns : [])
    .map((n) => Number(n) - 1).filter((n) => n >= 0));
  /** 負數或帶 ▲▼ 的值要標色——那是表格裡唯一該有顏色的東西。 */
  const warn = colour(args.warn_color, '#FF6B6B');

  return PowerPoint.run(async (ctx) => {
    const { slide, index } = await targetSlide(ctx, args);
    const made_ = [];
    const add = (shape) => { made_.push(shape); return shape; };

    rows.forEach((row, r) => {
      const isHead = hasHeader && r === 0;
      const isTotal = r === totalRow;
      const y = top + (rowH + gap) * r;
      for (let c = 0; c < cols; c += 1) {
        const cell = add(slide.shapes.addGeometricShape('Rectangle',
          { left: xOf(c), top: y, width: widths[c] - gap, height: rowH }));
        cell.name = stamp('tcell');
        const fill = isHead ? (args.header_fill || th.strong)
          : isTotal ? (args.total_fill || th.tint)
            // 斑馬紋：淡到幾乎看不見，它的工作是讓眼睛橫著走，不是分區。
            : (r % 2 === 0 ? th.bg : (args.band_fill || th.tint));
        cell.fill.setSolidColor(colour(fill));
        try { cell.lineFormat.visible = false; } catch { /* 舊版沒有就算了 */ }

        const text = String((row[c] ?? ''));
        cell.textFrame.textRange.text = text;
        const negative = /^[−-]|^\(.*\)$|▼/.test(text.trim());
        styleText(cell, {
          font: face,
          size: num(args.size, isHead ? 11.5 : 11),
          bold: isHead || isTotal || c === 0,
          color: isHead ? (args.header_text || th.bg)
            : negative ? warn
              : (args.text_color || th.ink),
          align: numeric.has(c) || (isHead && numeric.has(c)) ? 'right' : 'left',
        });
      }
    });

    await ctx.sync();
    const h = (rowH + gap) * rows.length;
    return `已在投影片 ${index} 放上 ${rows.length} 列 × ${cols} 欄的表格`
      + `（${Math.round(left)},${Math.round(top)}，寬 ${Math.round(width)}、高 ${Math.round(h)}）。
`
      + '表格是用方塊拼成的，不是 PowerPoint 的表格物件——要改某一格就直接點進那一格改字，'
      + '不能整列插入或刪除。要改結構的話跟我說，我重做一張。';
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
  /**
   * How tall a card is, worked out from what goes in it.
   *
   * ── The two failures a constant produces, and it produces both ──
   * Too short and the body has nowhere to go. Too tall and the text sits in the
   * top third of a box with an empty lower half — which is what the five-step
   * process row looked like, and it reads worse than crowding because it looks
   * like something failed to load.
   *
   * So the height is estimated from the longest body: characters that fit per
   * line at this width, rounded up to lines, plus the room the heading and the
   * padding actually take. It is an estimate and it does not need to be better
   * than one — the text is vertically centred inside whatever it produces, so
   * being a line out costs a few points of air rather than a clipped sentence.
   */
  const longest = Math.max(0, ...cards.map((c) => String((c && c.text) || '').length));
  const headLines = cards.some((c) => c && c.title) ? 1 : 0;
  const ch = num(args.height, (() => {
    if (bar !== 'top') return 68;
    if (!longest) return 92;
    const perLine = Math.max(6, Math.floor((cwEstimate(args, cards) - 28) / 12));
    const lines = Math.ceil(longest / perLine);
    return clamp(40 + headLines * 26 + lines * 19, 92, 210);
  })());
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
      /**
       * 圓角交給風格決定，不要因為色條在上面就自己圓起來。
       *
       * ── 症狀 ──
       * 一排卡片裡，編號那塊是方的、說明那塊是圓的——因為編號走的是另一條路徑。
       * 兩種圓角混在同一組元件上，看起來不是設計，是沒對齊。
       *
       * ── 原本為什麼這樣寫 ──
       * 「上色條配圓角」在只有卡片的那一頁是好看的。但卡片很少單獨出現——它旁邊通常
       * 有編號塊、底板、表格，而那些都是方的。一個元件自己決定圓角，就會跟鄰居打架。
       *
       * 圓角是整份的語彙，不是單一元件的裝飾。要圓角就選 soft 或 courseware，
       * 那時候整份都會圓；其餘情況一律方角。
       */
      const panel = add(slide.shapes.addGeometricShape(
        boxOf(lk, 'Rectangle'),
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
      const spine = add(slide.shapes.addGeometricShape('Rectangle',
        capBox(panel.geo || boxOf(lk, bar === 'top' ? 'RoundRectangle' : 'Rectangle'), bar === 'top'
          ? { left: x, top: y, width: cw, height: 5 }
          : { left: x, top: y, width: 5, height: ch }, bar !== 'top')));
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
        // A tinted disc behind the index. Two extra shapes for the row, and it
        // is the difference between a numbered list and a set of steps — the
        // number stops being a word at the start of the line and becomes a mark.
        if (bar === 'top') {
          const disc = add(slide.shapes.addGeometricShape(geoType('Ellipse'),
            { left: tx - 3, top: ty - 2, width: 34, height: 34 }));
          disc.name = stamp('numdot');
          disc.fill.setSolidColor(mix(accent, th.bg, 0.86));
          disc.lineFormat.visible = false;
        }
        const badge = add(slide.shapes.addTextBox(String(card.number),
          { left: tx - 3, top: ty - 2, width: 34, height: 34 }));
        badge.name = stamp('num');
        styleText(badge, {
          font: face, size: bar === 'top' ? 15 : 24, bold: true, color: accent,
          align: bar === 'top' ? 'center' : 'left', valign: 'middle',
        });
        badge.fill.clear();
        badge.lineFormat.visible = false;
        tx += bar === 'top' ? 42 : 56;
        tw -= bar === 'top' ? 42 : 56;
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
          { left: bodyX, top: ty, width: bodyW, height: y + ch - ty - 12 }));
        body.name = stamp('cb');
        // Centred in the space below the heading. Top-aligned was what produced
        // 「內容擠在上面、下面一片空白」 whenever the estimate ran generous.
        styleText(body, { font: face, size: num(args.small, 12.5), color: th.muted,
          valign: 'middle', lineSpacing: 1.3 });
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
        // A chevron rather than a fat arrow, and sitting a third of the way down
        // rather than halfway: at the vertical centre of a 180pt card the arrow
        // floats in the middle of nothing, level with body text instead of with
        // the step itself.
        const arrow = add(slide.shapes.addGeometricShape('Chevron',
          { left: ax, top: top + 24, width: gutter - 8, height: 18 }));
        arrow.name = stamp('arrow');
        arrow.fill.setSolidColor(mix(th.muted, th.bg, 0.35));
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

/**
 * A picture's own proportions, read out of the bytes.
 *
 * ── Why this matters more for illustrations than for photographs ──
 * `setSelectedDataAsync` stretches to whatever box it is given; it has no
 * fit-inside mode. On a photograph used as a full-bleed background that is
 * survivable — a landscape squeezed 10% is a landscape. On a flat illustration
 * it is not: a figure drawn on a 1:1 canvas stretched into a 384×540 column is
 * visibly a stretched drawing, and 「這樣會很怪」 is the correct reaction.
 *
 * So the natural size comes out of the file. SVG carries it in the viewBox or
 * the width/height attributes; PNG carries it in the IHDR chunk, which is always
 * the first chunk and always at the same offset. Returns null when neither is
 * readable, and the caller then falls back to filling the box — the old
 * behaviour, which is right for a background.
 */
function naturalSize(got) {
  if (got.svg) {
    const vb = /viewBox\s*=\s*"[\s\d.+-]*?([\d.]+)[\s,]+([\d.]+)\s*"/.exec(got.svg);
    if (vb) return { w: Number(vb[1]), h: Number(vb[2]) };
    const w = /width\s*=\s*"([\d.]+)/.exec(got.svg);
    const h = /height\s*=\s*"([\d.]+)/.exec(got.svg);
    if (w && h) return { w: Number(w[1]), h: Number(h[1]) };
    return null;
  }
  try {
    // PNG: 8-byte signature, then the IHDR length+type, then width and height as
    // big-endian 32-bit integers at byte 16 and 20.
    const head = atob(String(got.data).slice(0, 64));
    if (head.slice(1, 4) !== 'PNG') return null;
    const be = (at) => head.charCodeAt(at) * 16777216 + head.charCodeAt(at + 1) * 65536
      + head.charCodeAt(at + 2) * 256 + head.charCodeAt(at + 3);
    const w = be(16);
    const h = be(20);
    return w > 0 && h > 0 ? { w, h } : null;
  } catch {
    return null;
  }
}

/** Centre a picture's own shape inside the slot it was given. */
function fitInside(box, nat) {
  if (!nat || !nat.w || !nat.h) return box;
  const scale = Math.min(box.width / nat.w, box.height / nat.h);
  const width = Math.round(nat.w * scale);
  const height = Math.round(nat.h * scale);
  return {
    left: Math.round(box.left + (box.width - width) / 2),
    top: Math.round(box.top + (box.height - height) / 2),
    width,
    height,
  };
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
  /**
   * Fill the slot, or fit inside it?
   *
   * A `hero` is a background: filling is the point, and a few per cent of
   * stretch on a photograph is invisible. Anything else is a picture somebody
   * will look AT, so it keeps its own proportions and is centred in the slot.
   * Vector illustrations are never stretched, in any mode — a distorted drawing
   * is obvious in a way a distorted photograph is not.
   */
  const stretch = mode === 'hero' && !got.svg && args.fit !== 'contain';
  const placed = stretch ? geo : fitInside(geo, naturalSize(got));
  const note = await focusSlideFor(args);
  if (got.svg) {
    // Vectors need ImageCoercion 1.2. Saying so beats letting the call fail with
    // whatever Office decides to call it.
    if (!hasImageCoercion('1.2')) {
      throw new Error('這個 Office 版本放不了向量插畫（需要 ImageCoercion 1.2）。'
        + '改用 ppt_find_art 以外的做法：放照片，或用 ppt_add_diagram 畫圖解。');
    }
    await insertImageData(got.svg, Office.CoercionType.XmlSvg, placed);
  } else {
    await insertImageData(got.data, Office.CoercionType.Image, placed);
  }

  // Everything below is drawn AFTER the image, and that ordering is the whole
  // trick: setSelectedDataAsync appends, so a scrim created first would end up
  // underneath the photograph it is supposed to darken.
  const title = String(args.title || '');
  const sub = String(args.subtitle || '');
  const caption = String(args.caption || '');
  if (!title && !sub && !caption && mode !== 'hero') {
    return `已插入圖片，${placed.width}×${placed.height} 點，位置 (${placed.left},${placed.top})。`
      + `${placed.width !== geo.width || placed.height !== geo.height
        ? '（保持原比例置中，沒有拉伸。）' : ''}${note}`;
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

  // The deck's palette, not its NAME — see renderMermaid for what the old
  // call was actually asking mermaid for.
  const lk = look(args.look);
  const png = await renderMermaid(code,
    { palette: paletteFor(args, lk), font: args.font || (lk.font || paletteFor(args, lk).font) });
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
    /**
     * Every data field, copied wholesale.
     *
     * This used to name them one at a time, and the list was wrong twice:
     * `timeline` shipped unable to receive its own `events`, and `plot`, `box`,
     * `map` and `scatter` were all in the tool's enum while their data fields
     * were being dropped here. A chart type that cannot be given its data is
     * worse than one that is missing — the agent passes the right arguments,
     * the renderer reports them as absent, and nothing points at this function.
     */
    ...chartData(args),
    type: args.type,
    title: args.title,
    colors: args.colors || [th.strong, th.second, th.accent],
    width: num(args.width, 760),
    // A Gantt's height follows its row count, so it is left for the renderer to
    // decide unless somebody asked for one.
    height: args.height === undefined && String(args.type) === 'gantt'
      ? undefined : num(args.height, 380),
  }, th);

  const drawn = /height="(\d+)"/.exec(svg);
  const drawnH = drawn ? Number(drawn[1]) : num(args.height, 380);
  const geo = fitToBody(num(args.width, 760), drawnH, args);
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
    // Two passes. The first works out what will be touched and writes down how
    // it looks now; only then does anything change. Deciding and mutating in one
    // loop would mean capturing shapes we then skip, and skipping shapes we then
    // capture — and the undo would be a record of a different edit.
    const wanted = slides.map((slide) => ({
      slide,
      shapes: (slide.shapes.items || []).filter((shape) =>
        !onlyOurs || String(shape.name || '').startsWith(`${TAG}:`)),
    })).filter((g) => g.shapes.length);

    for (const g of wanted) {
      rememberBefore(g.slide.id, await captureShapes(ctx, g.shapes));
    }

    let touched = 0;
    for (const g of wanted) {
      for (const shape of g.shapes) {
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
 * Everything this tool call did, as a list of inverses.
 *
 * Collected per call and handed to the shell, which keeps the undo stack. The
 * list is emptied here so the next call starts clean — a tool that fails
 * half-way still returns what it managed to record, and restoring that is
 * correct: those shapes really were changed.
 */
async function captureCreated() {
  const snapshot = undoOps.length ? { kind: 'ops', ops: undoOps.slice() } : null;
  undoOps = [];
  return snapshot;
}

/** Put one shape's recorded values back. Anything unrecorded is left alone. */
function applyState(shape, before) {
  const b = before || {};
  const box = b.box || {};
  for (const k of ['left', 'top', 'width', 'height']) {
    if (Number.isFinite(box[k])) shape[k] = box[k];
  }
  // undefined means 「we could not read it」, null means 「there was no fill」.
  // Those are different instructions and collapsing them would paint a
  // transparent shape white every time an undo ran.
  if (b.fill !== undefined) {
    try {
      if (b.fill === null) shape.fill.clear();
      else if (b.fill) shape.fill.setSolidColor(b.fill);
    } catch { /* a shape that cannot be filled was not filled by us either */ }
  }
  if (b.line) {
    try {
      const lf = shape.lineFormat;
      if (b.line.visible !== undefined) lf.visible = b.line.visible;
      if (b.line.color !== undefined) lf.color = b.line.color;
      if (b.line.weight !== undefined) lf.weight = b.line.weight;
    } catch { /* ignore */ }
  }
  if (b.text !== undefined || b.font) {
    try {
      const range = shape.textFrame.textRange;
      if (b.text !== undefined) range.text = b.text;
      if (b.font) {
        for (const k of ['name', 'size', 'color', 'bold', 'italic']) {
          if (b.font[k] !== undefined) range.font[k] = b.font[k];
        }
      }
    } catch { /* no text frame */ }
  }
}

/** Build a shape back from its description. Returns false when we cannot. */
function rebuild(slide, before) {
  const b = before || {};
  if (!REBUILDABLE.has(String(b.type || '').toLowerCase())) return false;
  const box = b.box || {};
  const at = {
    left: Number.isFinite(box.left) ? box.left : 0,
    top: Number.isFinite(box.top) ? box.top : 0,
    width: Number.isFinite(box.width) ? box.width : 120,
    height: Number.isFinite(box.height) ? box.height : 40,
  };
  // Rectangle, not the original geometry: PowerPoint does not read back which
  // geometric shape a shape is. A rebuilt ellipse comes back square, which is
  // visible and fixable — unlike losing it, which is neither.
  const shape = /textbox/i.test(String(b.type || ''))
    ? slide.shapes.addTextBox(String(b.text || ''), at)
    : slide.shapes.addGeometricShape('Rectangle', at);
  if (b.name) shape.name = b.name;
  applyState(shape, /textbox/i.test(String(b.type || '')) ? { ...b, text: undefined } : b);
  return true;
}

/**
 * Undo.
 *
 * Applied in reverse: the last thing done is the first thing taken back, which
 * is the only order that survives a call that both changed a shape and then
 * deleted it.
 */
export async function restoreSnapshot(snapshot) {
  const ops = (snapshot && snapshot.kind === 'ops' && snapshot.ops) || [];
  if (!ops.length) return { kind: 'none', note: '這次的動作沒有可以退回的東西。' };

  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.slides;
    slides.load('items/id');
    await ctx.sync();
    const bySlide = new Map((slides.items || []).map((x) => [x.id, x]));

    const done = { removed: 0, restored: 0, rebuilt: 0, slides: 0, lost: 0 };
    for (const entry of ops.slice().reverse()) {
      const slide = entry.slideId ? bySlide.get(entry.slideId) : null;
      if (entry.op === 'created') {
        if (!slide) continue;
        try { slide.shapes.getItem(entry.shapeId).delete(); done.removed++; } catch {
          // Already gone — the person deleted it themselves. Fine.
        }
      } else if (entry.op === 'restore') {
        if (!slide) continue;
        try { applyState(slide.shapes.getItem(entry.before.shapeId), entry.before); done.restored++; } catch {
          // The shape it describes is no longer there. Nothing to put back.
        }
      } else if (entry.op === 'recreate') {
        if (!slide) continue;
        try { if (rebuild(slide, entry.before)) done.rebuilt++; else done.lost++; } catch { done.lost++; }
      } else if (entry.op === 'slide') {
        try {
          slides.add();
          await ctx.sync();
          slides.load('items/id');
          await ctx.sync();
          const fresh = slides.items[slides.items.length - 1];
          // A new slide is born with the default layout's empty title and
          // subtitle prompts. Leaving them on a rebuilt slide would put 「按一下
          // 以新增標題」 behind the content we are restoring — the same problem
          // buildSlide already solves, solved the same way.
          try { await clearPlaceholders(ctx, fresh); } catch { /* not fatal */ }
          for (const before of entry.shapes || []) {
            try { if (rebuild(fresh, before)) done.rebuilt++; else done.lost++; } catch { done.lost++; }
          }
          done.slides++;
        } catch {
          done.lost += (entry.shapes || []).length;
        }
      }
    }
    await ctx.sync();

    const said = [];
    if (done.removed) said.push(`移除新增的 ${done.removed} 個物件`);
    if (done.restored) said.push(`還原 ${done.restored} 個被改過的物件`);
    if (done.slides) said.push(`重建 ${done.slides} 張投影片`);
    if (done.rebuilt) said.push(`重建 ${done.rebuilt} 個被刪掉的物件`);

    const caveats = [];
    if (done.lost) {
      caveats.push(`${done.lost} 個物件沒辦法還原（圖片和群組讀不回來）`);
    }
    if (done.slides) {
      caveats.push('重建的投影片會接在最後面，不是原本的位置——'
        + 'PowerPoint 的 API 只能往後加。順序要對請用 PowerPoint 自己的 Ctrl+Z');
    }
    if (done.rebuilt) {
      caveats.push('重建的形狀一律是矩形，因為 PowerPoint 讀不回原本是圓形還是箭頭');
    }

    return {
      kind: 'ops',
      ...done,
      note: (said.length ? `已${said.join('、')}。` : '沒有東西需要退回。')
        + (caveats.length ? `\n要注意：${caveats.join('；')}。` : ''),
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

/** Where this deck's id lives. Same name Word uses, for one thing to look for. */
const DECK_ID_KEY = 'JVisionDocumentId';

/**
 * The add-in's own storage inside the file.
 *
 * ── Why this and not a custom document property ──
 * Excel and Word each mint a UUID into a custom document property, which travels
 * with the file and survives a rename. PowerPointApi has no equivalent, so this
 * used to fall back to the file's PATH — and a path is not an identity.
 *
 * `Office.context.document.settings` is the Common API's answer to exactly this:
 * a small store that belongs to this add-in, lives inside the document, and goes
 * wherever the document goes. It is not gated on a PowerPointApi requirement set
 * because it is not a PowerPointApi feature.
 */
function deckSettings() {
  try {
    const s = Office.context.document.settings;
    return (s && typeof s.get === 'function' && typeof s.set === 'function') ? s : null;
  } catch {
    return null;
  }
}

/** Read this deck's id, minting one the first time. '' when it cannot be stored. */
async function storedDeckId() {
  const store = deckSettings();
  if (!store) return '';
  try {
    const saved = store.get(DECK_ID_KEY);
    if (saved) return String(saved);
  } catch {
    return '';
  }
  const id = (self.crypto && self.crypto.randomUUID)
    ? self.crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(16).slice(2)}`;
  try {
    store.set(DECK_ID_KEY, id);
    // Only trust it once it is actually in the file. A id that lives in memory
    // and vanishes on reopen is worse than no id: it looks like a new deck every
    // session, which is a different wrong answer from the one being fixed here.
    const saved = await new Promise((resolve) => {
      store.saveAsync((res) => resolve(!!res && res.status === Office.AsyncResultStatus.Succeeded));
    });
    return saved ? id : '';
  } catch {
    return '';
  }
}

/**
 * Which presentation this is — the key the conversation hangs off.
 *
 * ── The bug this replaces ──
 * It used to return `ppt:${Office.context.document.url}`, i.e. the file's path.
 * A path answers 「哪個位置」, and the question here is 「哪一份檔案」. Those come
 * apart in both directions, and both were reported:
 *
 *   · rename a deck → new path → its history is stranded on the old name
 *   · make a NEW deck with the same name in the same folder → same path →
 *     it inherits the previous file's entire conversation
 *
 * The second one is the dangerous half. Somebody opens a fresh 報價簡報.pptx and
 * the assistant appears to remember a deck they have never seen, which is also
 * how one person's work becomes visible in another file's transcript.
 *
 * So the id goes INSIDE the file, the way Excel's and Word's already do. The
 * path stays only as a last resort for a build where settings cannot be saved —
 * unchanged from today's behaviour, so nothing that works now stops working.
 */
export async function getDocumentKey() {
  const id = await storedDeckId();
  if (id) return `ppt:${id}`;
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
/**
 * 這份簡報規劃過了嗎，每一頁又規劃了什麼。
 *
 * ── 為什麼規劃要變成一個工具 ──
 * 增益集原本的流程是「使用者說話 → 模型直接呼叫零件」，中間沒有任何一個階段叫設計。
 * 所以不管給什麼需求，它都從手上的零件盒裡挑——KPI 卡、圖表、條列——因為沒有任何
 * 一步**要求**它先決定這一頁長什麼樣。做出來的東西像「從固定套件裡挑幾個」，
 * 不像有人規劃過。
 *
 * 我在提示詞裡寫過三次「請先規劃」，搬過位置、修過矛盾，行為沒變。那是**規則**，
 * 而規則擋不住模型。這幾天反覆驗證的是：**機制擋得住**。
 *
 * 所以規劃變成一個工具，而且建構工具會檢查它跑過沒有——模型繞不過去，也不需要記得。
 *
 * ── 為什麼不是只記一個布林值 ──
 * 規劃的內容要留著：後面的檢查要拿它對照（說好要放一張表，最後沒放，查得出來），
 * 側邊欄也要顯示出來，讓使用者在它動手之前就看得到方向對不對。
 */
let deckPlan = null;
const slidePlans = new Map();

/** 需要先有版面規劃才能動手的工具——會真的在投影片上放東西的那些。 */
const NEEDS_PLAN = new Set([
  // 「組一頁」才需要規劃。改一行字、搬一個方塊是編輯，不是設計——
  // 連那個都擋的話，使用者說「把標題改成 X」還要先規劃一次版面，那是找麻煩。
  'ppt_build_slide',
]);

/**
 * 整份簡報的走向：每頁一行，講這一頁要證明什麼、用什麼形式。
 *
 * 一頁的簡報也要跑——「一頁」正是最需要規劃的情況，因為所有東西要擠進同一張紙，
 * 位置是先決的，不是做到哪算哪。
 */
async function planDeck(args) {
  const slides = Array.isArray(args.slides) ? args.slides : [];
  if (!slides.length) {
    throw new Error('slides 是必要的：每頁一個物件，至少要有 { purpose, form }。'
      + '例如 [{ purpose: "現況：市場規模與成長", form: "數字卡 + 折線圖" }]。');
  }
  /**
   * 配色是規劃階段的決定，不是每一頁各自挑。
   *
   * 不給的話就擋下來——這跟 headline 一樣，是「做之前非決定不可」的那一類。一份給客戶
   * 的提案、一份內部檢討、一份財報分析，該長的樣子完全不同；而預設值只有一種，
   * 所以不逼它決定，結果就永遠是同一種。
   */
  /**
   * 沒給配色不是錯誤——用預設，然後說出來。
   *
   * ── 第一版為什麼錯 ──
   * 我原本在這裡擋下來，想逼模型做決定。結果是它連試兩次失敗、放棄規劃、退回原本的
   * KPI 加兩張圖——**它永遠走最便宜的路**，而我把規劃那條路變貴了。
   *
   * 第一步擋下來不會換來更好的決定，只會換來不用這個工具。
   */
  /**
   * 沒指定配色的話，從主題推一個——而不是永遠落回同一個預設。
   *
   * ── 為什麼不能只有一個預設 ──
   * 「不指定就用預設」聽起來合理，但預設只有一個值，所以旅遊、財報、資安、產品發表
   * 全部長成同一套藍綠。使用者的感受是「不管我下什麼需求，PPT 都一樣」——而那不是
   * 模型沒創意，是這條路上只有一個出口。
   *
   * ── 為什麼用雜湊而不是隨機 ──
   * 隨機的話同一份簡報重跑一次就換色，使用者會以為是 bug，而且沒辦法重現。
   * 從主題文字推導：同一個主題永遠同一套（可重現），不同主題自然不同（有變化）。
   *
   * 模型自己指定的仍然優先——這只是「沒人決定時」的出口，而那個出口現在有九個方向。
   */
  const picked = !args.theme && !args.look;
  if (picked) {
    const seed = `${args.headline || ''}${args.audience || ''}`
      + slides.map((x) => String((x && x.purpose) || '')).join('');
    let hash = 0;
    for (const ch of seed) hash = (hash * 31 + ch.codePointAt(0)) >>> 0;
    const pool = Object.keys(THEMES);
    if (pool.length) deckPalette = pool[hash % pool.length];
  }
  if (args.theme) deckPalette = String(args.theme);
  else if (!picked) deckPalette = null;
  deckLook = args.look ? String(args.look) : null;
  deckPlan = {
    headline: String(args.headline || ''),
    audience: String(args.audience || ''),
    theme: deckPalette,
    look: deckLook,
    slides: slides.map((x, i) => ({
      n: i + 1,
      purpose: String((x && x.purpose) || ''),
      form: String((x && x.form) || ''),
    })),
  };
  slidePlans.clear();
  const lines = deckPlan.slides.map((x) => `${x.n} ${x.purpose}（${x.form}）`);
  const tone = [deckPalette && `配色 ${deckPalette}`, deckLook && `風格 ${deckLook}`]
    .filter(Boolean).join('、') || '配色沿用預設';
  return `整份規劃好了（${tone}），共 ${lines.length} 頁：\n${lines.join('\n')}\n\n`
    + '接下來每做一頁之前，先用 ppt_plan_slide 規劃那一頁的版面。'
    + (picked
      ? `\n\n沒指定配色，所以照主題挑了「${deckPalette}」。這份要給誰看？對外提案、內部檢討、財報分析，`
        + '該長的樣子完全不同——想換的話再跑一次這個工具並帶上 theme 或 look，'
        + '整份都會跟著改，不用一頁一頁調。'
      : '');
}

/**
 * 單頁的版面：分幾塊、各佔哪裡、每塊什麼媒材。
 *
 * 要求講出**理由**不是為了好看——說不出「為什麼這塊用圖不用條列」，通常代表那塊
 * 其實不需要存在，或者它該換一種形式。這一步把那個判斷逼到動手之前。
 */
async function planSlide(args) {
  const n = Math.max(1, Number(args.slide) || 1);
  const blocks = Array.isArray(args.blocks) ? args.blocks : [];
  if (!blocks.length) {
    throw new Error('blocks 是必要的：這一頁分成幾塊、各佔哪裡、每塊用什麼媒材。'
      + '例如 [{ area: "左上", media: "長條圖", why: "要比三組採用 vs 成效的落差" }]。');
  }
  /**
   * 主標缺了也不要擋——理由同 planDeck：規劃不該比不規劃更容易失敗。
   *
   * 留一個提示，然後讓這一頁的規劃成立。模型看得到，而且不會因為少一個欄位就整個
   * 退回去用原本那套。
   */
  const headline = String(args.headline || '').trim();
  const plan = {
    headline,
    blocks: blocks.map((b) => ({
      area: String((b && b.area) || ''),
      media: String((b && b.media) || ''),
      why: String((b && b.why) || ''),
      // 座標是選填的，但給了就等於這一頁有了設計稿——後面每個工具照著放，排不歪。
      left: Number(b && b.left) || null,
      top: Number(b && b.top) || null,
      width: Number(b && b.width) || null,
      height: Number(b && b.height) || null,
    })),
  };
  slidePlans.set(n, plan);
  const lines = plan.blocks.map((b) => {
    const at = (b.left !== null && b.top !== null)
      ? `（${b.left},${b.top}${b.width ? ` ${b.width}×${b.height || '?'}` : ''}）` : '';
    return `· ${b.area}${at}：${b.media}${b.why ? ` — ${b.why}` : ''}`;
  });
  // 條列佔太多塊，通常代表有東西該畫成圖——在動手之前講，改起來最便宜。
  const listy = plan.blocks.filter((b) => /條列|清單|文字|bullet/i.test(b.media)).length;
  const warn = listy > 2
    ? '\n\n提醒：這一頁有三塊以上是條列。回頭看看有沒有哪一塊其實是圖或表——通常有。'
    : '';
  const lead = headline
    ? `主標「${headline}」`
    : '主標還沒決定——那句話要寫成一句主張（「採用衝很快，成效跟不上」），'
      + '不是題目名稱（「AI 趨勢分析」）。它決定這一頁其餘所有東西怎麼排，先想好再做。';
  return `第 ${n} 頁的版面：\n${lead}\n${lines.join('\n')}${warn}`;
}

/**
 * 還沒規劃就想動手的時候，把它擋回去。
 *
 * 回傳的是一段錯誤訊息，不是拋例外——模型看得到、照著做就過得了，而且這比在提示詞裡
 * 寫一百次「記得先規劃」有效，因為它不是提醒，是一道門。
 */
function planGate(tool, args) {
  if (!NEEDS_PLAN.has(tool)) return '';
  const n = Math.max(1, Number(args && (args.slide ?? args.slide_index)) || 1);
  if (slidePlans.has(n)) return '';
  return `第 ${n} 頁還沒有版面規劃。先呼叫 ppt_plan_slide 決定這一頁：`
    + '要讓人記得哪一句話（headline，寫成主張不是題目）、分成幾塊、各佔哪裡、'
    + '每塊用什麼媒材以及為什麼。\n\n'
    + '這一步不能跳過：先想內容再找地方放，做出來的頁面是把區塊依序堆下去的結果——'
    + '每一塊都對，但整頁沒有人規劃過。'
    + (deckPlan ? '' : '\n\n整份還沒規劃的話，也先跑一次 ppt_plan_deck。');
}

export function supportedTools() {
  const tools = [
    'ppt_get_overview',
    'ppt_read_slide',
    'ppt_get_selection',
    'ppt_plan_deck',
    'ppt_plan_slide',
    'ppt_add_slide',
    'ppt_delete_slide',
    'ppt_build_slide',
    'ppt_add_text',
    'ppt_review_layout',
    'ppt_add_infographic',
    'ppt_draw',
    'ppt_add_shape',
    'ppt_add_line',
    'ppt_add_table',
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
    // Answered by the server, not here — but they still have to be advertised.
    'ppt_find_art',
    'ppt_generate_image',
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
  /**
   * 模型送來的字串先洗過——字面的換行、tab、被空格隔開的底線。
   *
   * 做在派送這一層而不是每個工具裡：文字不只從 args.text 進來，卡片是陣列、KPI 是
   * 物件陣列、表格是二維陣列。一處一處補就是一處一處漏——這個問題在 Word 補了三次
   * 才收斂，不要在這裡重走一遍。
   */
  args = cleanArgs(args);
  // 還沒規劃就想動手的話，擋回去。見 planGate。
  const gate = planGate(tool, args);
  if (gate) return { ok: false, error: gate };

  const a = args || {};
  // One call, one undo entry. Cleared at the start rather than the end so a
  // call that throws part-way leaves behind what it actually managed to do.
  undoOps = [];
  try {
    switch (tool) {
      case 'ppt_get_overview':  return { ok: true, content: await getOverview(a) };
      case 'ppt_read_slide':    return { ok: true, content: await readSlide(a) };
      case 'ppt_get_selection': return { ok: true, content: await getSelection() };

      case 'ppt_run_script':
        return { ok: true, content: await runScript(a, opts) };

      default: {
        const run = {
          ppt_plan_deck: planDeck,
          ppt_plan_slide: planSlide,
          ppt_add_slide: addSlide,
          ppt_delete_slide: deleteSlide,
          ppt_build_slide: buildSlide,
          ppt_add_text: addText,
          ppt_review_layout: reviewLayout,
          ppt_add_infographic: addInfographic,
          ppt_draw: drawShapes,
          ppt_add_shape: addShape,
          ppt_add_line: addLine,
          ppt_add_table: addTable,
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
    /**
     * Office 的錯誤要把 debugInfo 挖出來。
     *
     * ── 為什麼非做不可 ──
     * 「兩次失敗都只回報『未知的失敗』，沒有說哪裡出錯，所以我沒辦法確定原因」——
     * 這是模型自己的話，而它是對的。RichApi 的錯誤常常 `message` 是空的或極短，
     * 真正有用的（哪個屬性、哪個呼叫、Office 的說法）全在 `debugInfo` 裡。
     *
     * 不挖出來的話，模型只能猜：它改參數、改寫法、再試一次，失敗三輪之後放棄改用
     * 手畫——而使用者看到的是「這個功能好像壞了」，沒有人知道為什麼。
     *
     * 我自己也為此繞了三圈：猜過座標越界、猜過顏色、猜過 schema 擋掉，
     * 全都不是。**一個不說原因的錯誤，會讓所有人一起用猜的。**
     *
     * Word 的宿主早就這樣做了，PPT 這邊漏了。
     */
    const code = e && e.code ? `[${e.code}] ` : '';
    let msg = (e && e.message) || String(e);
    const info = e && e.debugInfo;
    if (info) {
      const where = info.errorLocation ? `在 ${info.errorLocation}` : '';
      const detail = info.message && info.message !== (e && e.message) ? info.message : '';
      const extra = [where, detail].filter(Boolean).join('：');
      if (extra) msg += `（${extra}）`;
    }
    if (!String(msg).trim()) msg = 'Office 沒有給錯誤訊息——通常是某個屬性的值超出範圍。';
    return { ok: false, content: `${code}${msg}` };
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
    ppt_plan_deck: t('正在規劃整份的走向…'),
    ppt_plan_slide: t('正在規劃這一頁的版面…'),
    ppt_build_slide: t('正在組一張投影片…'),
    ppt_add_text: t('正在加上文字…'),
    ppt_review_layout: t('正在檢查版面…'),
    ppt_add_infographic: t('正在畫資訊圖…'),
    ppt_draw: t('正在畫版面…'),
    ppt_add_shape: t('正在加上圖形…'),
    ppt_add_line: t('正在畫線…'),
    ppt_add_table: t('正在排表格…'),
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
    ppt_generate_image: t('正在生成圖片（十幾秒）…'),
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
