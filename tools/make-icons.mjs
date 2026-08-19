/**
 * Generates the ribbon icons (16/32/80 px) as PNGs, with zero dependencies.
 *
 * The manifest requires real image files at real URLs — Excel will not render the
 * ribbon button without them — so we need *something* checked in before the
 * add-in can load at all. This writes a plain green tile so the pipeline works on
 * day one; replace assets/icon-*.png with the real artwork whenever design has it.
 *
 * Node's zlib does the compression, so this is a ~60-line PNG encoder rather than
 * an image library.
 */
import fs from 'fs';
import path from 'path';
import zlib from 'zlib';
import { fileURLToPath } from 'url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.join(ROOT, 'assets');

const CRC_TABLE = (() => {
  const t = new Int32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c;
  }
  return t;
})();

function crc32(buf) {
  let c = -1;
  for (let i = 0; i < buf.length; i++) c = CRC_TABLE[(c ^ buf[i]) & 0xff] ^ (c >>> 8);
  return (c ^ -1) >>> 0;
}

function chunk(type, data) {
  const len = Buffer.alloc(4);
  len.writeUInt32BE(data.length);
  const body = Buffer.concat([Buffer.from(type, 'ascii'), data]);
  const crc = Buffer.alloc(4);
  crc.writeUInt32BE(crc32(body));
  return Buffer.concat([len, body, crc]);
}

/** Excel green with a lighter inner square, so the shape reads at 16px. */
function pixel(x, y, size) {
  const pad = Math.max(1, Math.round(size * 0.16));
  const inner = x >= pad && y >= pad && x < size - pad && y < size - pad;
  const innerPad = Math.max(2, Math.round(size * 0.34));
  const core = x >= innerPad && y >= innerPad && x < size - innerPad && y < size - innerPad;
  if (core) return [255, 255, 255, 255];
  if (inner) return [33, 115, 70, 255];
  return [0, 0, 0, 0];
}

function makePng(size) {
  const raw = Buffer.alloc((size * 4 + 1) * size);
  let p = 0;
  for (let y = 0; y < size; y++) {
    raw[p++] = 0; // filter: none
    for (let x = 0; x < size; x++) {
      const [r, g, b, a] = pixel(x, y, size);
      raw[p++] = r; raw[p++] = g; raw[p++] = b; raw[p++] = a;
    }
  }
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(size, 0);
  ihdr.writeUInt32BE(size, 4);
  ihdr[8] = 8;   // bit depth
  ihdr[9] = 6;   // colour type: RGBA
  return Buffer.concat([
    Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
    chunk('IHDR', ihdr),
    chunk('IDAT', zlib.deflateSync(raw, { level: 9 })),
    chunk('IEND', Buffer.alloc(0)),
  ]);
}

fs.mkdirSync(OUT, { recursive: true });
for (const size of [16, 32, 80]) {
  const file = path.join(OUT, `icon-${size}.png`);
  fs.writeFileSync(file, makePng(size));
  console.log(`wrote ${file}`);
}
