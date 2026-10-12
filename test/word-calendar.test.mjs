/**
 * 月曆。
 *
 * ── 這支測試在守什麼 ──
 * 日期算術。一份月曆只要有一天排錯，整份就是廢的，而「排錯一天」看起來跟排對
 * 一模一樣——沒有人會逐格去對 2026 年 3 月 1 日是不是星期日。所以這裡把會出錯的
 * 地方全部釘住：每個月的第一天落在哪一欄、閏年的二月、跨週的最後一列、以及週起始
 * 改成星期一之後整排要跟著移。
 *
 * 另外一條是刻意的「不做」：**不內建任何國定假日**。各國不一樣、每年都在動，
 * 而一份把上班日標成放假的月曆比完全不標更糟。
 */
import {
  executeTool, calendarGridFor, yearCalendarSvgFor,
} from '../src/taskpane/hosts/word.js';

let pass = 0;
let fail = 0;
function ok(name, cond, detail) {
  if (cond) { pass++; console.log(`PASS  ${name}`); } else {
    fail++; console.log(`FAIL  ${name}${detail === undefined ? '' : `  → ${detail}`}`);
  }
}

// ── 日期算術 ────────────────────────────────────────────────────────────────
{
  // 2026-01-01 是星期四：第一週前面要空四格（日一二三）。
  const jan = calendarGridFor(2026, 1, 0);
  ok('一月第一天落在星期四',
    JSON.stringify(jan[0]) === JSON.stringify([null, null, null, null, 1, 2, 3]),
    JSON.stringify(jan[0]));
  ok('一月有 31 天', jan.flat().filter((d) => d !== null).length === 31);
  ok('最後一列補滿七格', jan[jan.length - 1].length === 7);

  // 週起始改星期一，整排要往左移一格。
  const janMon = calendarGridFor(2026, 1, 1);
  ok('改成週一起算，第一天就移一格',
    JSON.stringify(janMon[0]) === JSON.stringify([null, null, null, 1, 2, 3, 4]),
    JSON.stringify(janMon[0]));

  ok('閏年二月 29 天',
    calendarGridFor(2024, 2, 0).flat().filter((d) => d !== null).length === 29);
  ok('平年二月 28 天',
    calendarGridFor(2025, 2, 0).flat().filter((d) => d !== null).length === 28);
  // 2100 不是閏年——世紀年要能被 400 整除才是。
  ok('2100 年二月 28 天',
    calendarGridFor(2100, 2, 0).flat().filter((d) => d !== null).length === 28);

  // 每個月的天數都要對，一個一個驗。
  const want = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  const got = want.map((_, i) => calendarGridFor(2026, i + 1, 0).flat()
    .filter((d) => d !== null).length);
  ok('2026 十二個月的天數全對', JSON.stringify(got) === JSON.stringify(want), got.join(' '));

  // 每一列都是連號，中間不能跳掉或重複。
  const flat = calendarGridFor(2026, 5, 0).flat().filter((d) => d !== null);
  ok('日期是連號的', flat.every((d, i) => d === i + 1), flat.join(' '));
}

// ── 畫出來的年曆 ────────────────────────────────────────────────────────────
{
  const svg = yearCalendarSvgFor(2026, { holidays: ['2026-01-01', '2026-10-10'] });
  ok('沒有 NaN', !/NaN|undefined/.test(svg));
  ok('十二個月都在',
    Array.from({ length: 12 }, (_, i) => `${i + 1} 月`).every((m) => svg.includes(`>${m}<`)));

  const W = Number(/width="([\d.]+)"/.exec(svg)[1]);
  const H = Number(/height="([\d.]+)"/.exec(svg)[1]);
  const xs = [...svg.matchAll(/<text x="([-\d.]+)" y="([-\d.]+)"/g)]
    .map((m) => [Number(m[1]), Number(m[2])]);
  ok('每個字都在畫布裡',
    xs.every(([x, y]) => x >= -1 && x <= W + 1 && y >= -1 && y <= H + 1),
    `${W}x${H}`);

  /**
   * 星期日要是紅的，而且不能只有星期日的「日」字紅。
   *
   * 整個第一欄的日期都該是紅的——一份只有標題紅、日期黑的月曆，等於沒標。
   */
  const reds = (svg.match(/fill="#C00000"/g) || []).length;
  ok('紅色的日子夠多（週末 + 假日）', reds > 100, `${reds} 個`);

  // 沒給假日就一個都不標——不要自己編。
  const plain = yearCalendarSvgFor(2026, {});
  const plainReds = (plain.match(/fill="#C00000"/g) || []).length;
  ok('沒給假日時只有週末是紅的', plainReds > 50 && plainReds < reds + 1,
    `${plainReds} vs ${reds}`);

  // 單休的公司
  const sixDay = yearCalendarSvgFor(2026, { saturday: false });
  ok('關掉週六就少一半紅字',
    (sixDay.match(/fill="#C00000"/g) || []).length < plainReds * 0.75,
    `${(sixDay.match(/fill="#C00000"/g) || []).length} vs ${plainReds}`);
}

// ── 走完整支工具 ────────────────────────────────────────────────────────────
/**
 * 一個剛好夠 svgToPng 跑完的瀏覽器。
 *
 * 轉點陣圖要 `Image` 和 canvas，node 沒有。少了這段，年曆那條會以
 * 「Image is not defined」失敗——那不是程式的錯，是環境的，但也不能因此就不測：
 * 年曆走的是跟月曆完全不同的一條路（一張圖 vs 一張表），不測等於那半邊沒人看。
 *
 * 這裡只假裝到「圖片載得進來、畫布吐得出 base64」為止。畫出來長什麼樣是另一回事，
 * 那個在 yearCalendarSvgFor 那段用 SVG 直接量。
 */
function fakeBrowser() {
  global.Image = class {
    set src(v) { this._src = v; setTimeout(() => this.onload && this.onload(), 0); }
    get src() { return this._src; }
  };
  global.document = {
    createElement: () => ({
      width: 0,
      height: 0,
      getContext: () => ({ fillRect() {}, drawImage() {}, set fillStyle(v) {} }),
      toDataURL: () => 'data:image/png;base64,QQ==',
    }),
  };
}

function fakeWord() {
  fakeBrowser();
  const log = { tables: [], cells: [], pics: [], paras: [] };
  const makeCell = (r, c) => {
    const props = { r, c };
    log.cells.push(props);
    return {
      body: {
        font: {},
        insertParagraph: (text) => { log.paras.push({ r, c, text }); return { font: {} }; },
        insertOoxml: () => {},
      },
      setCellPadding: () => {},
      set verticalAlignment(v) { props.valign = v; },
      set shadingColor(v) { props.fill = v; },
      set columnWidth(v) { props.width = v; },
    };
  };
  const makeTable = (rows, cols, loc, values) => {
    const t = {
      styleBuiltIn: '', headerRowCount: -1, styleBandedRows: null, styleFirstColumn: null,
      width: 451,
      load() {},
      getCell: makeCell,
      getBorder: () => ({}),
      rows: { items: [], load() {} },
    };
    log.tables.push({ rows, cols, loc, values, table: t });
    return t;
  };
  const ctx = {
    document: {
      changeTrackingMode: 'Off',
      getStyles: () => ({ getByNameOrNullObject: () => ({ isNullObject: true, font: {}, load() {} }) }),
      body: {
        insertTable: makeTable,
        insertOoxml: () => {},
        insertParagraph: () => ({
          insertOoxml: () => {},
          styleBuiltIn: '',
          alignment: '',
          insertBreak: () => {},
          insertInlinePictureFromBase64: (data) => {
            const pic = { data, width: 0, height: 0 };
            log.pics.push(pic);
            return pic;
          },
        }),
        paragraphs: { items: [], load() {} },
        getOoxml: () => ({ value: '<xml/>' }),
        load() {},
      },
    },
    sync: async () => {},
    load() {},
  };
  global.Word = { run: async (fn) => fn(ctx), InsertLocation: {} };
  global.Office = {
    context: {
      requirements: { isSetSupported: () => true },
      document: { url: `cal-${Math.random()}.docx`, settings: { get: () => null, set() {}, saveAsync() {} } },
    },
  };
  return log;
}

{
  const log = fakeWord();
  const res = await executeTool('word_insert_calendar', {
    year: 2026, month: 3, lines: 3,
    holidays: [{ date: '2026-03-01', label: '補假' }],
  });
  ok('單月版做得出來', res && res.ok, JSON.stringify(res).slice(0, 120));
  const t = log.tables[0];
  ok('七欄', t && t.cols === 7, t && t.cols);
  // 2026 年 3 月：1 號是星期日，31 天 → 五列 + 表頭。
  ok('列數對', t && t.rows === 6, t && t.rows);
  ok('第一列是星期標題',
    t && JSON.stringify(t.values[0]) === JSON.stringify(['日', '一', '二', '三', '四', '五', '六']),
    t && JSON.stringify(t.values[0]));
  ok('三月一號排在星期日那欄', t && t.values[1][0] === '1', t && t.values[1][0]);
  ok('假日名稱寫進格子裡', log.paras.some((q) => q.text === '補假'),
    JSON.stringify(log.paras.slice(0, 3)));
  // 每格要留可以寫字的空行。
  const blanks = log.paras.filter((q) => q.text === '').length;
  ok('每格留了三行', blanks >= 31 * 3, `${blanks} 行`);
}

{
  const log = fakeWord();
  const res = await executeTool('word_insert_calendar', { year: 2026 });
  ok('年曆版做得出來', res && res.ok, JSON.stringify(res).slice(0, 120));
  ok('年曆是一張圖', log.pics.length === 1 && log.tables.length === 0,
    `${log.pics.length} 張圖、${log.tables.length} 張表`);
  ok('沒給假日時會講一聲', /國定假日每年都在動/.test(String(res.content)),
    String(res.content).slice(0, 120));
}

{
  fakeWord();
  const bad = await executeTool('word_insert_calendar', { year: 2026, month: 13 });
  ok('月份超出範圍會擋', bad && bad.ok === false && /1 到 12/.test(String(bad.error)),
    String(bad && bad.error));
  const bad2 = await executeTool('word_insert_calendar', { year: 1200 });
  ok('年份超出範圍會擋', bad2 && bad2.ok === false, String(bad2 && bad2.error));
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
