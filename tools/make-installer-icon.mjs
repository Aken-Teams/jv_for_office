/**
 * public/logo → installer/JVision.ico
 *
 * An .exe with no icon shows the generic "unknown application" glyph in Explorer,
 * which is exactly the look that makes someone hesitate before double-clicking —
 * the same objection that started this whole installer.
 *
 * .ico is a plain container format: a 6-byte header, one 16-byte directory entry
 * per size, then the image data. Windows Vista and later accept PNG payloads
 * directly, so there is no BMP/AND-mask encoding to do; sharp resizes, this file
 * packs. That means no ico-specific dependency for one icon.
 *
 * Run: npm run installer:icon   (build-installer.ps1 picks the file up if present)
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath, pathToFileURL } from 'url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

// sharp lives in the ai-agents-office server package — see make-icons.mjs for why
// this repo does not carry its own copy.
const SHARP_PATHS = [
  path.resolve(ROOT, '..', 'ai-agents-office', 'server', 'node_modules', 'sharp'),
  path.resolve(ROOT, '..', 'ai-agents-office', 'node_modules', 'sharp'),
  path.resolve(ROOT, 'node_modules', 'sharp'),
];

const SOURCE = path.join(ROOT, 'assets', 'logo.png');
const OUT = path.join(ROOT, 'installer', 'JVision.ico');

/**
 * Sizes Windows actually asks for: 16 in the title bar, 32 on the desktop, 48 in
 * Explorer's medium view, 256 for the extra-large view and the Store-style
 * property sheet. 64 and 128 fill the gaps so nothing is ever upscaled.
 */
const SIZES = [16, 32, 48, 64, 128, 256];

async function loadSharp() {
  for (const p of SHARP_PATHS) {
    if (fs.existsSync(p)) return (await import(pathToFileURL(path.join(p, 'lib', 'index.js')).href)).default;
  }
  throw new Error(
    'sharp 找不到。它是從 ai-agents-office 借用的，請先確認那個專案有安裝過依賴。',
  );
}

function pack(images) {
  const header = Buffer.alloc(6);
  header.writeUInt16LE(0, 0);            // reserved
  header.writeUInt16LE(1, 2);            // 1 = icon
  header.writeUInt16LE(images.length, 4);

  const dir = Buffer.alloc(16 * images.length);
  // Data starts after the header and the whole directory.
  let offset = 6 + 16 * images.length;

  images.forEach(({ size, data }, i) => {
    const e = 16 * i;
    // 256 is stored as 0 — the field is one byte, so 256 does not fit.
    dir.writeUInt8(size >= 256 ? 0 : size, e + 0);
    dir.writeUInt8(size >= 256 ? 0 : size, e + 1);
    dir.writeUInt8(0, e + 2);            // palette entries (0 = truecolour)
    dir.writeUInt8(0, e + 3);            // reserved
    dir.writeUInt16LE(1, e + 4);         // colour planes
    dir.writeUInt16LE(32, e + 6);        // bits per pixel
    dir.writeUInt32LE(data.length, e + 8);
    dir.writeUInt32LE(offset, e + 12);
    offset += data.length;
  });

  return Buffer.concat([header, dir, ...images.map(i => i.data)]);
}

const sharp = await loadSharp();
if (!fs.existsSync(SOURCE)) throw new Error(`找不到來源圖 ${SOURCE}`);

const images = [];
for (const size of SIZES) {
  const data = await sharp(SOURCE)
    .resize(size, size, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
    .png({ compressionLevel: 9 })
    .toBuffer();
  images.push({ size, data });
}

fs.writeFileSync(OUT, pack(images));
console.log(`JVision.ico 已產生（${SIZES.join('、')} px，${(fs.statSync(OUT).size / 1024).toFixed(1)} KB）`);
console.log('接著跑 npm run installer:build 讓 exe 帶上它。');
