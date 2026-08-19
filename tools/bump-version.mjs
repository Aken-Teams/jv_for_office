/**
 * Bumps the manifest <Version>.
 *
 * Office decides whether to re-read an add-in by (add-in Id, version). Same
 * version means "nothing changed" — so it keeps serving the ribbon from its own
 * cache and never re-downloads the icons, even after the icon cache is cleared
 * and even though the URLs return the new files perfectly well. That is exactly
 * the state that shows a generic cube on the ribbon.
 *
 * So: any change Office reads out of the manifest — icons included, because it
 * caches them per manifest version rather than per URL — needs a bump. Then
 * re-run installer\安裝.cmd so the copy under %LOCALAPPDATA% carries the new
 * version, and restart Excel.
 *
 * Run: npm run bump
 */
import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const MANIFEST = path.join(ROOT, 'manifest.xml');

const xml = fs.readFileSync(MANIFEST, 'utf-8');
const match = xml.match(/<Version>(\d+)\.(\d+)\.(\d+)\.(\d+)<\/Version>/);
if (!match) {
  console.error('manifest.xml 裡找不到四段式的 <Version>。');
  process.exit(1);
}

const [, major, minor, patch] = match.map(Number);
// Bump the third segment: the fourth is reserved for build numbers, and Office
// rejects a version below 1.0.
const next = `${major}.${minor}.${patch + 1}.0`;

fs.writeFileSync(MANIFEST, xml.replace(match[0], `<Version>${next}</Version>`), 'utf-8');

console.log(`manifest 版本：${match[1]}.${match[2]}.${match[3]}.${match[4]} → ${next}`);
console.log('');
console.log('接下來（順序不能顛倒）：');
console.log('  1. 完全關閉 Excel');
console.log('  2. installer\\安裝.cmd        ← 把新版本的 manifest 複製過去');
console.log('  3. npm run clear-cache        ← 清掉舊圖示');
console.log('  4. 確認 npm run dev 有在跑，再開 Excel');
