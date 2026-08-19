/**
 * Builds the ribbon icons (16/32/80 px) from the brand logo.
 *
 * The manifest needs real square PNGs at real URLs — Excel won't draw the ribbon
 * button without them. The source logo is 254×176, i.e. NOT square, so it is
 * fitted into a transparent square rather than stretched; squashing a logo to fit
 * a ribbon slot is the kind of thing people notice without being able to say why.
 *
 * Uses sharp from the ai-agents-office server package rather than adding a
 * dependency here: this script runs once when the artwork changes, and the two
 * repos are checked out side by side anyway. If that path ever moves, install
 * sharp locally and drop the SHARP_PATHS lookup.
 *
 * Run: npm run icons
 */
import fs from 'fs';
import path from 'path';
import { createRequire } from 'module';
import { fileURLToPath } from 'url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(ROOT, 'assets');
const SOURCE = path.join(ROOT, 'public', 'logo2.png');

const SHARP_PATHS = [
  path.resolve(ROOT, '..', 'ai-agents-office', 'server', 'node_modules', 'sharp'),
  'sharp',
];

function loadSharp() {
  const require = createRequire(import.meta.url);
  for (const p of SHARP_PATHS) {
    try { return require(p); } catch { /* try the next one */ }
  }
  throw new Error(
    'sharp not found. Either check out ai-agents-office beside this repo, '
    + 'or run: npm i -D sharp',
  );
}

/**
 * Strip every ancillary PNG chunk, leaving only the critical ones
 * (IHDR / PLTE / IDAT / IEND).
 *
 * This is not tidiness — it is the fix. Office fetched our icons perfectly well
 * (verified in the dev server's request log: three 200s, correct byte counts)
 * and then refused to draw them, falling back to the generic add-in cube. The
 * only structural difference from a known-good Office ribbon icon — Claude for
 * Excel's, fetched from pivot.claude.ai — was that sharp emits a `pHYs` chunk
 * (physical pixel density, 96 DPI) and theirs has none:
 *
 *     Claude:  IHDR IDAT IEND          → renders
 *     ours:    IHDR pHYs IDAT IEND     → generic cube
 *
 * Both are 32x32, 8-bit RGBA with transparency, within 30 bytes of each other.
 * So transparency was never the problem; Office evidently reads pHYs and does
 * something with it that ends in nothing being drawn.
 *
 * Written by hand rather than via a sharp option because sharp derives pHYs from
 * the source image's density and offers no reliable way to omit it.
 */
function stripAncillaryChunks(buf) {
  const KEEP = new Set(['IHDR', 'PLTE', 'IDAT', 'IEND']);
  const out = [buf.subarray(0, 8)];               // signature
  let p = 8;
  while (p < buf.length) {
    const len = buf.readUInt32BE(p);
    const type = buf.toString('ascii', p + 4, p + 8);
    const end = p + 12 + len;                     // len + type + data + crc
    if (KEEP.has(type)) out.push(buf.subarray(p, end));
    if (type === 'IEND') break;
    p = end;
  }
  return Buffer.concat(out);
}

/**
 * 16 / 32 / 80 are the ribbon sizes Office asks for (bt:Image in VersionOverrides).
 *
 * 64 is separate and NOT optional: the manifest reference states that for task
 * pane add-ins HighResolutionIconUrl "must be 64 x 64 pixels". We had it pointed
 * at the 80px file, which is a documented violation — and the dialog it feeds
 * (Insert > Add-ins) was showing the generic placeholder because of it.
 * https://learn.microsoft.com/javascript/api/manifest/highresolutioniconurl
 */
const SIZES = [16, 32, 64, 80];

/**
 * A logo drawn edge to edge looks cramped next to Excel's own ribbon icons, which
 * all carry a little breathing room. 6% inset at 32px is roughly 2 device pixels —
 * enough to sit right, small enough not to shrink the mark.
 */
const INSET = 0.06;

/**
 * At 16px the full logo is unreadable: the constellation and the ring collapse
 * into one smudge. So the smallest icon uses the ring alone, which is the half
 * that still reads at that size. Standard practice — favicons do the same thing.
 *
 * The crop is hand-measured, and deliberately so. Both automated approaches fail
 * on this artwork: colour can't separate the two marks (the constellation has
 * cyan nodes too), and neither can stroke density (its dots are as heavy as the
 * ring). They also overlap horizontally, so there is no clean vertical cut.
 * Re-measure this box if the logo is ever redrawn.
 */
const RING_CROP = { left: 126, top: 20, width: 122, height: 134 };
const SIMPLIFY_BELOW = 32;   // sizes strictly below this use the ring alone

const sharp = loadSharp();

if (!fs.existsSync(SOURCE)) {
  console.error(`找不到來源圖：${SOURCE}`);
  process.exit(1);
}

fs.mkdirSync(OUT, { recursive: true });

const meta = await sharp(SOURCE).metadata();
console.log(`來源：${path.relative(ROOT, SOURCE)}  ${meta.width}×${meta.height}`);

for (const size of SIZES) {
  const inner = Math.round(size * (1 - INSET * 2));
  const pad = Math.round((size - inner) / 2);
  const simplify = size < SIMPLIFY_BELOW;

  const base = simplify ? sharp(SOURCE).extract(RING_CROP) : sharp(SOURCE);
  const buf = await base
    .resize(inner, inner, {
      fit: 'contain',                                   // keep the aspect ratio
      background: { r: 0, g: 0, b: 0, alpha: 0 },        // transparent, not white:
    })                                                   // the ribbon has its own
    .extend({                                            // background in light AND
      top: pad, bottom: size - inner - pad,              // dark Office themes
      left: pad, right: size - inner - pad,
      background: { r: 0, g: 0, b: 0, alpha: 0 },
    })
    .png({ compressionLevel: 9 })
    .toBuffer();

  const clean = stripAncillaryChunks(buf);
  const file = path.join(OUT, `icon-${size}.png`);
  fs.writeFileSync(file, clean);
  console.log(`  icon-${size}.png  ${clean.length} bytes  ${simplify ? '（只取圓環）' : '（完整 logo）'}`);
}

// Full-resolution copy for the task pane header, where the logo can keep its
// natural proportions instead of being padded into a square.
const paneLogo = path.join(OUT, 'logo.png');
fs.copyFileSync(SOURCE, paneLogo);
console.log(`  logo.png  (原尺寸，供側邊欄使用)`);
