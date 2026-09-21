/**
 * Build the files for testing excel_compare_source by hand.
 *
 * ── Why fixtures and not just unit tests ──
 * test/compare.test.mjs covers the engine, but the engine is only half of it.
 * The other half is everything that happens on the way in: what Excel hands an
 * add-in for a date cell (a serial), what a text extractor makes of the same
 * cell (words), how a CSV writes a number somebody formatted. Those cannot be
 * asserted from Node — they need a real workbook, a real upload, and somebody
 * looking at the result.
 *
 * So these files are built to contain every trap at once, and 答案卷.md says
 * what the correct output is for each one. A fixture without an answer key is
 * not a test; it is just a file.
 *
 *   node tools/make-compare-fixtures.mjs [輸出資料夾]
 *
 * Default output is d:/tmp/jv-compare-fixtures. Nothing is written into the
 * repo: these are throwaway files for one afternoon of clicking.
 */
import fs from 'fs';
import path from 'path';
import { createRequire } from 'module';
import { fileURLToPath } from 'url';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT = path.resolve(process.argv[2] || 'd:/tmp/jv-compare-fixtures');

/** exceljs lives in the server package — same borrowing as make-icons.mjs. */
function loadExcelJS() {
  const require = createRequire(import.meta.url);
  for (const p of [
    path.resolve(ROOT, '..', 'ai-agents-office', 'server', 'node_modules', 'exceljs'),
    path.resolve(ROOT, '..', 'ai-agents-office', 'node_modules', 'exceljs'),
    'exceljs',
  ]) {
    try { return require(p); } catch { /* try the next one */ }
  }
  throw new Error('找不到 exceljs。請確認隔壁的 ai-agents-office 有裝過依賴。');
}

const ExcelJS = loadExcelJS();
const d = (iso) => new Date(`${iso}T00:00:00Z`);

// ── 本月：the sheet the person opens and compares FROM ────────────────────────
//
// Twelve rows is small enough to check by eye and long enough that a
// misalignment is obvious rather than arguable.
const HEAD = ['代號', '客戶', '日期', '數量', '金額', '備註'];
const THIS_MONTH = [
  ['A001', '台積電', d('2024-01-05'), 100, 1234.5, ''],
  // A comma inside a value. In the CSV twin this is the quoted-cell trap: parse
  // it wrong and every column to the right shifts, which reports as a difference
  // on every single row.
  ['A002', '鴻海, 深圳廠', d('2024-01-06'), 50, 2000, ''],
  ['A003', '聯發科', d('2024-01-07'), 0, 0, ''],
  ['A004', '廣達', d('2024-01-08'), 12, 300, '急件'],
  // Blank, not zero. A blank where a figure belongs IS a difference, and any
  // implementation that folds one into the other hides it.
  ['A005', '緯創', d('2024-01-09'), 7, null, ''],
  ['A006', '仁寶', d('2024-01-10'), 3, 99.99, ''],
  ['A007', '和碩', d('2024-01-11'), 8, 150, ''],
  ['A008', '英業達', d('2024-01-12'), 20, 400, ''],
  // Duplicate key, twice over: the pairing rule has to send the first to the
  // first and the second to the second, not report two additions and two removals.
  ['A009', '日月光', d('2024-01-13'), 5, 88, '第一筆'],
  ['A009', '日月光', d('2024-01-13'), 6, 90, '第二筆'],
  ['A010', '台達電', d('2024-01-14'), 9, 175, ''],
  ['A011', '光寶', d('2024-01-15'), 4, 60, ''],
];

// ── 上月：the second sheet, for same-workbook comparison ──────────────────────
//
// Column order differs from 本月 on purpose — 備註 has moved to the front. Align
// by header and this is invisible; align by position and every column reports as
// changed, which is the failure this trap exists to catch.
const LAST_HEAD = ['備註', '代號', '客戶', '日期', '數量', '金額'];
const LAST_MONTH = [
  // An inserted row at the TOP. Without a key everything below shifts by one and
  // the whole sheet reads as changed; with 代號 as the key it is one addition.
  ['', 'X999', '新客戶', d('2023-12-31'), 1, 10],
  ['', 'A001', '台積電', d('2024-01-05'), 100, 1234.5],
  ['', 'A002', '鴻海, 深圳廠', d('2024-01-06'), 50, 2000],
  ['', 'A003', '聯發科', d('2024-01-07'), 5, 0],          // 數量 0 → 5
  ['急件', 'A004', '廣達', d('2024-01-08'), 12, 300],
  ['', 'A005', '緯創', d('2024-01-09'), 7, 0],            // 金額 空白 → 0
  ['', 'A006', '仁寶', d('2024-01-10'), 3, 99.99],
  // A007 is missing entirely — 只在本月.
  ['', 'A008', '英業達', d('2024-01-12'), 20, 400],
  ['第一筆', 'A009', '日月光', d('2024-01-13'), 5, 88],
  ['第二筆', 'A009', '日月光', d('2024-01-13'), 6, 95],   // 第二筆的金額 90 → 95
  ['', 'A010', '台達電', d('2024-01-14'), 9, 175],
  ['', 'A011', '光寶', d('2024-01-15'), 4, 60],
  ['', 'A012', '華碩', d('2024-01-16'), 6, 120],          // 只在上月
];

async function writeMain(file) {
  const wb = new ExcelJS.Workbook();

  const a = wb.addWorksheet('本月');
  a.addRow(HEAD);
  THIS_MONTH.forEach((r) => a.addRow(r));
  a.getColumn(3).numFmt = 'yyyy/m/d';
  a.getColumn(5).numFmt = '#,##0.00';

  const b = wb.addWorksheet('上月');
  b.addRow(LAST_HEAD);
  LAST_MONTH.forEach((r) => b.addRow(r));
  b.getColumn(4).numFmt = 'yyyy/m/d';
  b.getColumn(6).numFmt = '#,##0.00';

  // A third sheet that is not part of any comparison, so picking the wrong one
  // by accident is visible rather than silent.
  const c = wb.addWorksheet('說明');
  c.addRow(['這張表不是拿來比對的。']);
  c.addRow(['「本月」對「上月」＝同活頁簿比對；「本月」對上傳的檔案＝跨檔案比對。']);

  await wb.xlsx.writeFile(file);
}

// ── 來源檔：the upload, for cross-file comparison against 本月 ─────────────────
//
// Same data as 上月, but every value that CAN be written a different way IS —
// this is the side that goes through the server's text extractor, so it is where
// type and formatting traps actually bite.
async function writeSource(file) {
  const wb = new ExcelJS.Workbook();
  const s = wb.addWorksheet('明細');
  s.addRow(HEAD);

  const rows = [
    ['A001', '台積電', d('2024-01-05'), 100, 1234.5, ''],
    ['A002', '鴻海, 深圳廠', d('2024-01-06'), 50, 2000, ''],
    ['A003', '聯發科', d('2024-01-07'), 5, 0, ''],
    ['A004', '廣達', d('2024-01-08'), 12, 300, '急件'],
    ['A005', '緯創', d('2024-01-09'), 7, 0, ''],
    ['A006', '仁寶', d('2024-01-10'), 3, 99.99, ''],
    ['A008', '英業達', d('2024-01-12'), 20, 400, ''],
    ['A009', '日月光', d('2024-01-13'), 5, 88, '第一筆'],
    ['A009', '日月光', d('2024-01-13'), 6, 95, '第二筆'],
    ['A010', '台達電', d('2024-01-14'), 9, 175, ''],
    ['A011', '光寶', d('2024-01-15'), 4, 60, ''],
    ['A012', '華碩', d('2024-01-16'), 6, 120, ''],
  ];
  rows.forEach((r) => s.addRow(r));
  s.getColumn(3).numFmt = 'yyyy/m/d';

  // A second sheet with more rows than the real one, so `source_sheet` has to be
  // honoured: pick "the biggest" and this decoy wins and the answer is wrong.
  const decoy = wb.addWorksheet('工作表2');
  decoy.addRow(['這張是誘餌', '列數比「明細」多']);
  for (let i = 0; i < 40; i++) decoy.addRow([`X${i}`, i]);

  await wb.xlsx.writeFile(file);
}

/**
 * The CSV twin: same rows, written the way an exported report writes them.
 *
 * Thousands separators, a currency mark, accounting parentheses, a 民國 date, a
 * percentage and an id with a leading zero. None of these are differences — they
 * are the same values wearing different clothes — and a comparison that reports
 * them fills the sheet with noise and buries the four real ones.
 */
function csvText() {
  const lines = [
    '代號,客戶,日期,數量,金額,備註',
    'A001,台積電,113/1/5,100,"1,234.50",',              // 民國年 + 千分位
    'A002,"鴻海, 深圳廠",2024/1/6,50.0,NT$2000,',       // 引號逗號 + 貨幣 + 50 vs 50.0
    'A003,聯發科,2024-01-07,5,0,',
    'A004,廣達,2024/1/8,12,300,急件',
    'A005,緯創,2024/1/9,7,0,',
    'A006,仁寶,2024/1/10,3,99.99,',
    'A008,英業達,2024/1/12,20,400,',
    'A009,日月光,2024/1/13,5,88,第一筆',
    'A009,日月光,2024/1/13,6,95,第二筆',
    'A010,台達電,2024/1/14,9,"(175)",',                // 會計負數：這是真差異
    'A011,光寶,2024/1/15,4,60,',
    'A012,華碩,2024/1/16,6,120,',
  ];
  return lines.join('\r\n') + '\r\n';                   // CRLF, like Excel exports
}

const ANSWER = `# 比對測試 — 答案卷

三個測試。每一個的正確結果都寫在下面，跑出來的東西跟這裡不一樣就是有 bug。

**共通**：一定要帶 \`key_column = 代號\`。不帶的話會退回「依列的先後順序對應」，
上月最前面那一列 X999 會讓後面每一列都位移——那不是 bug，是沒給鍵的正確行為，
報表開頭也會寫明。想看那個行為長怎樣，可以故意不給鍵跑一次。

---

## 測試一：同一份活頁簿，本月 vs 上月

開啟 \`比對測試-主檔.xlsx\`，在側邊欄說：

> 幫我比對「本月」和「上月」這兩張工作表，用代號當鍵

**應該要有 4 筆差異：**

| 狀態 | 鍵值 | 欄位 | 本月 | 上月 |
|---|---|---|---|---|
| 值不同 | A003 | 數量 | 0 | 5 |
| 值不同 | A005 | 金額 | (空白) | 0 |
| 值不同 | A009 | 金額 | 90 | 95 |
| 只在本月 | A007 | （整列） | 和碩… | |
| 只在上月 | A012 | （整列） | | 華碩… |
| 只在上月 | X999 | （整列） | | 新客戶… |

（值不同 3 格、只在本月 1 列、只在上月 2 列。）

**這個測試在驗什麼：**

1. **欄位順序不同不算差異。** 上月的「備註」在第一欄，本月在最後一欄。
   按標題名對齊就看不出來；按位置對齊的話每一欄都會被回報成不同。
2. **A005 的空白 vs 0 是真差異。** 把空白當成 0 的實作會漏掉這筆。
3. **重複鍵 A009 有兩筆**，要第一筆對第一筆、第二筆對第二筆。
   配對錯的話會變成 2 筆新增 + 2 筆刪除，而不是 1 格值不同。
4. **X999 插在最前面**，有鍵的話它只是一筆新增，不會讓其他列位移。

---

## 測試二：跨檔案，本月 vs 上傳的 xlsx

同一份主檔開著，把 \`比對測試-來源.xlsx\` 拖進對話框，然後說：

> 用剛上傳的檔案比對「本月」，來源用「明細」那張表，鍵是代號

**差異應該跟測試一完全一樣**（來源檔的「明細」內容等同上月，只是少了 X999）：

| 狀態 | 鍵值 | 欄位 |
|---|---|---|
| 值不同 | A003 | 數量 |
| 值不同 | A005 | 金額 |
| 值不同 | A009 | 金額 |
| 只在本月 | A007 | （整列） |
| 只在來源 | A012 | （整列） |

**這個測試在驗什麼：**

1. **日期欄不可以出現任何差異。** 這是最重要的一條。Excel 交給增益集的日期是
   序號（45296 之類），上傳檔經過文字抽取後是字串。沒處理好的話
   **每一列的日期都會被回報成不同**，12 筆假差異把 3 筆真差異埋掉。
   看到日期欄整片都是差異 → 有 bug。
2. **\`source_sheet\` 有被聽進去。** 來源檔裡有一張「工作表2」列數比「明細」多。
   沒指定的話工具會挑列數最多的那張，結果就會整個錯。
   請明確講「用明細那張」；也可以故意不講，確認它挑了誘餌之後報表開頭有寫是哪張。

---

## 測試三：CSV，型別與格式的陷阱

把 \`比對測試-來源.csv\` 拖進去，比對「本月」，鍵一樣是代號。

**只應該有 4 筆差異：**

| 狀態 | 鍵值 | 欄位 | 本月 | 來源 |
|---|---|---|---|---|
| 值不同 | A003 | 數量 | 0 | 5 |
| 值不同 | A005 | 金額 | (空白) | 0 |
| 值不同 | A009 | 金額 | 90 | 95 |
| 值不同 | A010 | 金額 | 175 | -175 |
| 只在本月 | A007 | （整列） | 和碩… | |
| 只在來源 | A012 | （整列） | | 華碩… |

**下面這些「看起來不一樣」的，全部都不可以被回報成差異：**

| 本月 | 來源 CSV | 為什麼相同 |
|---|---|---|
| 2024/1/5 | \`113/1/5\` | 民國年 |
| 1234.5 | \`"1,234.50"\` | 千分位是外觀 |
| 2000 | \`NT$2000\` | 貨幣符號是外觀 |
| 50 | \`50.0\` | 同一個數 |
| 鴻海, 深圳廠 | \`"鴻海, 深圳廠"\` | 引號裡的逗號不是欄位分隔 |

**而 A010 的 \`(175)\` 是真差異**：會計寫法的括號代表負數，本月是 175、來源是 -175。
這一條是用來確認「寬鬆」沒有寬鬆過頭——把括號當裝飾忽略掉的實作會漏掉這筆。

如果 A002 那一列出現一堆欄位都不同，就是 CSV 的引號逗號沒解析對，
整列從第 2 欄開始往右位移了。

---

## 測試四：不能比對的格式

把 \`不能比對.txt\` 拖進去要求比對。

**應該**：明確說 txt 沒有列與欄的結構、沒辦法逐格比對，並建議改用
\`excel_read_file\` 讀內容。**不應該**：硬做出一張報表，或是沉默地失敗。

---

## 另外可以順手看的

- **不指定 \`output_sheet\`** → 開一張新的「比對 本月」，不跳確認視窗。
  再跑一次應該變成「比對 本月 (2)」，不是報錯。
- **指定一張已經存在的工作表** → 應該跳確認視窗，說會覆寫。
- **按「復原上一步」** → 新開的比對工作表整張消失；覆寫既有表的話內容回到原狀。
- **報表「位置」欄** → 點得動，會跳到「本月」對應的那一格。
`;

// ════════════════════════════════════════════════════════════════════════════
// 進階：滾動預測與庫存快照
//
// The basic fixture above is a clean table with one key column. Neither of the
// two things people actually compare looks like that. A forecast is wide, its
// months roll every version, and a row is identified by 料號＋廠區＋客戶. An
// inventory snapshot repeats the same 料號 in every warehouse and is identified
// by 料號＋倉別＋批號. Keyed on one column, both produce a confidently wrong
// report — which is why they are here.
// ════════════════════════════════════════════════════════════════════════════

const FC_V2_HEAD = ['料號', '廠區', '客戶', '2024-02', '2024-03', '2024-04', '2024-05', '合計'];
const FC_V2 = [
  ['PN-1001', '台中', 'ACME', 1200, 1300, 1250, 1400, 5150],
  ['PN-1001', '蘇州', 'ACME', 800, 850, 800, 900, 3350],
  ['PN-1002', '台中', 'ACME', 500, 520, 540, 560, 2120],
  ['PN-1002', '台中', 'BETA', 300, 310, 320, 330, 1260],
  ['PN-1003', '蘇州', 'BETA', 0, 0, 100, 200, 300],
  ['PN-1004', '台中', 'GAMMA', 60, 60, 60, 60, 240],
];

// 上一版：月份整排往前一格，列序也不同——ERP 匯出很少保證排序穩定。
const FC_V1_HEAD = ['料號', '廠區', '客戶', '2024-01', '2024-02', '2024-03', '2024-04', '合計'];
const FC_V1 = [
  ['PN-1002', '台中', 'BETA', 290, 300, 310, 320, 1220],
  ['PN-1001', '蘇州', 'ACME', 780, 800, 850, 800, 3230],
  ['PN-1001', '台中', 'ACME', 1100, 1200, 1300, 1250, 4850],
  ['PN-1002', '台中', 'ACME', 480, 500, 600, 540, 2120],   // 2024-03 才是真差異
  ['PN-1003', '蘇州', 'BETA', 0, 0, 0, 100, 100],
  ['PN-1005', '台中', 'DELTA', 40, 40, 40, 40, 160],        // 只在上版
];

const INV_HEAD = ['料號', '倉別', '批號', '數量', '單位', '效期', '狀態'];
const INV_TODAY = [
  ['PN-1001', 'W01', 'L2401', 1200, 'PCS', d('2025-06-30'), '可用'],
  ['PN-1001', 'W02', 'L2401', 300, 'PCS', d('2025-06-30'), '可用'],
  ['PN-1001', 'W02', 'L2402', 450, 'PCS', d('2025-09-30'), '可用'],
  ['PN-1002', 'W01', 'L2403', 80, 'PCS', d('2025-03-31'), '待驗'],
  ['PN-1002', 'W03', 'L2403', 20, 'PCS', d('2025-03-31'), '可用'],
  ['PN-1003', 'W01', 'L2404', 0, 'PCS', d('2026-01-31'), '凍結'],
];

async function writeAdvanced(file) {
  const wb = new ExcelJS.Workbook();

  const a = wb.addWorksheet('預測-本版');
  a.addRow(FC_V2_HEAD);
  FC_V2.forEach((r) => a.addRow(r));

  const b = wb.addWorksheet('預測-上版');
  b.addRow(FC_V1_HEAD);
  FC_V1.forEach((r) => b.addRow(r));

  const c = wb.addWorksheet('庫存-今日');
  c.addRow(INV_HEAD);
  INV_TODAY.forEach((r) => c.addRow(r));
  c.getColumn(6).numFmt = 'yyyy/m/d';

  await wb.xlsx.writeFile(file);
}

/**
 * Yesterday's inventory, as an ERP writes it.
 *
 * Row order differs, quantities carry thousands separators and trailing spaces,
 * and 效期 is 民國. Exactly one value really changed.
 */
function inventoryCsv() {
  return [
    '料號,倉別,批號,數量,單位,效期,狀態',
    'PN-1002,W03,L2403,20,PCS,114/3/31,可用',
    'PN-1001,W02,L2402,450,PCS,114/9/30,可用',
    'PN-1001,W01,L2401,"1,200",PCS,114/6/30,可用',
    'PN-1002,W01,L2403,100,PCS,114/3/31,待驗',
    'PN-1001,W02,L2401,300 ,PCS,114/6/30,可用',
    'PN-1003,W01,L2404,0,PCS,115/1/31,凍結',
  ].join('\r\n') + '\r\n';
}

const ADVANCED = `# 進階測試 — 預測與庫存

\`比對測試-進階.xlsx\` 有三張表：預測-本版、預測-上版、庫存-今日。
搭配 \`庫存-昨日.csv\`。

**這兩個測試的重點只有一件事：鍵要給好幾個欄位。**
只給料號的話，同一個料號在每個倉別／每個廠區都有一筆，全部變成重複鍵，
配對退回「出現順序」——而兩邊列序刻意不同，結果就是隨機配對。

---

## 測試五：滾動預測，本版 vs 上版

> 比對「預測-本版」和「預測-上版」，鍵用 料號、廠區、客戶

**應該要有 5 筆值不同 + 2 筆整列：**

| 狀態 | 鍵值 | 欄位 | 本版 | 上版 |
|---|---|---|---|---|
| 值不同 | PN-1002 / 台中 / ACME | 2024-03 | 520 | 600 |
| 值不同 | PN-1001 / 台中 / ACME | 合計 | 5150 | 4850 |
| 值不同 | PN-1001 / 蘇州 / ACME | 合計 | 3350 | 3230 |
| 值不同 | PN-1002 / 台中 / BETA | 合計 | 1260 | 1220 |
| 值不同 | PN-1003 / 蘇州 / BETA | 合計 | 300 | 100 |
| 只在本版 | PN-1004 / 台中 / GAMMA | （整列） | | |
| 只在上版 | PN-1005 / 台中 / DELTA | （整列） | | |

**四筆「合計」差異是對的，不是 bug**：兩版涵蓋的月份不一樣（本版 2–5 月、
上版 1–4 月），合計本來就不該相等。真正的預測變動只有 PN-1002/台中/ACME 的 2024-03。

**摘要區應該要有這兩列：**

- 只在 預測-本版 的欄位：**2024-05**
- 只在 預測-上版 的欄位：**2024-01**

滾動掉的那個月和新長出來的那個月，是**整份檔案的一個事實，講一次就好**。
如果它變成「每一列的 2024-01 和 2024-05 都不同」（多 12 筆假差異），就是有 bug。

**只給料號當鍵會怎樣**：實測值不同會從 5 筆爆到 21 筆，而且廠區、客戶欄自己
會出現「台中 → 蘇州」這種差異——那是配對錯了的招牌症狀。可以故意跑一次看看。

---

## 測試六：庫存快照，今日 vs 昨日

開著 \`比對測試-進階.xlsx\`，把 \`庫存-昨日.csv\` 拖進去：

> 用剛上傳的檔案比對「庫存-今日」，鍵用 料號、倉別、批號

**正確答案只有一筆：**

| 狀態 | 鍵值 | 欄位 | 今日 | 昨日 |
|---|---|---|---|---|
| 值不同 | PN-1002 / W01 / L2403 | 數量 | 80 | 100 |

**其他全部都不可以出現：**

| 今日 | 昨日 CSV | 為什麼相同 |
|---|---|---|
| 1200 | \`"1,200"\` | 千分位 |
| 300 | \`300 \` | 尾隨空白（ERP 匯出常有） |
| 2025/6/30 | \`114/6/30\` | 民國年 |
| 2026/1/31 | \`115/1/31\` | 民國年跨年 |

而且六列的順序在兩邊完全不同，正確配對之後不應該有任何「只在某一邊」。

**只給料號當鍵會怎樣**：實測是 **15 筆差異，其中 14 筆是假的，
而唯一那筆真的還配錯對**（把 W01 的 80 拿去跟 W03 的 20 比）。
這就是為什麼工具說明裡要求寧可多給鍵也不要少給。
`;

fs.mkdirSync(OUT, { recursive: true });
await writeMain(path.join(OUT, '比對測試-主檔.xlsx'));
await writeAdvanced(path.join(OUT, '比對測試-進階.xlsx'));
fs.writeFileSync(path.join(OUT, '庫存-昨日.csv'), '﻿' + inventoryCsv(), 'utf8');
fs.writeFileSync(path.join(OUT, '答案卷-進階.md'), ADVANCED, 'utf8');
await writeSource(path.join(OUT, '比對測試-來源.xlsx'));
fs.writeFileSync(path.join(OUT, '比對測試-來源.csv'), '\ufeff' + csvText(), 'utf8');
fs.writeFileSync(
  path.join(OUT, '不能比對.txt'),
  '這是一個純文字檔，沒有列與欄。\n拿它去比對應該要被工具擋下來，並建議改用 excel_read_file。\n',
  'utf8',
);
fs.writeFileSync(path.join(OUT, '答案卷.md'), ANSWER, 'utf8');

console.log(`測試檔產生在 ${OUT}`);
for (const f of fs.readdirSync(OUT)) {
  console.log(`  ${f}  ${fs.statSync(path.join(OUT, f)).size} B`);
}
console.log('\n先看 答案卷.md，它寫了每個測試的正確結果。');
