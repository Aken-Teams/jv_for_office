/**
 * 側欄版面。
 *
 * ── 這支測試的由來 ──
 * 「有辦法多一些範本或是風格如圖這樣嗎？……我認為整體 WORD 的設計可以，但是對於
 * 多元的版面是缺乏的」，附圖是一份設計感履歷：左邊一條深色欄放聯絡方式和技能條，
 * 右邊放經歷。在這之前所有東西都是整頁寬的，那個版面根本到不了。
 *
 * 這裡守的是三件會讓它變回「一張表格」的事：
 *
 *   · 框線要關掉——看得到格線就不是版面了
 *   · 色塊欄要真的上色，而且欄寬要照指定的走
 *   · 文字要是**真的 Word 文字**，不是一張圖
 *
 * 最後一條最重要。用圖片做會長得一模一樣，然後使用者一個字都改不了。
 */
import { executeTool, barsSvgFor } from '../src/taskpane/hosts/word.js';

let pass = 0;
let fail = 0;
function ok(name, cond, detail) {
  if (cond) { pass++; console.log(`PASS  ${name}`); } else {
    fail++; console.log(`FAIL  ${name}${detail === undefined ? '' : `  → ${detail}`}`);
  }
}

const TABLE_WIDTH_PT = 451;

function fakeWord() {
  const log = { tables: [], cells: [], ooxml: [], borders: [], pics: [], padding: [] };
  const makeCell = (r, c) => {
    const props = { r, c };
    log.cells.push(props);
    return {
      body: {
        font: {},
        insertOoxml: (xml, loc) => log.ooxml.push({ c, xml, loc }),
        insertParagraph: () => ({
          insertInlinePictureFromBase64: (data) => {
            const pic = { data, width: 0, height: 0 };
            log.pics.push(pic);
            return pic;
          },
        }),
      },
      setCellPadding: (loc, pt) => log.padding.push({ c, loc, pt }),
      set verticalAlignment(v) { props.valign = v; },
      set columnWidth(v) { props.width = v; },
      set shadingColor(v) { props.fill = v; },
    };
  };
  const makeTable = (rows, cols, loc) => {
    const t = {
      styleBuiltIn: '', headerRowCount: -1, styleBandedRows: null, styleFirstColumn: null,
      width: undefined,
      load(w) { if (String(w).includes('width')) this.width = TABLE_WIDTH_PT; },
      getCell: (r, c) => makeCell(r, c),
      getBorder: (where) => { const b = { where }; log.borders.push(b); return b; },
    };
    log.tables.push({ rows, cols, loc, table: t });
    return t;
  };
  const ctx = {
    document: {
      changeTrackingMode: 'Off',
      getStyles: () => ({ getByNameOrNullObject: () => ({ isNullObject: true, font: {}, load() {} }) }),
      body: {
        insertTable: makeTable,
        insertOoxml: () => {},
        insertParagraph: () => ({ insertOoxml: () => {}, styleBuiltIn: '', insertBreak: () => {} }),
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
      document: { url: `cv-${Math.random()}.docx`, settings: { get: () => null, set() {}, saveAsync() {} } },
    },
  };
  return log;
}

const CV = {
  sidebar: [
    { title: '王小明', subtitle: '資深製程工程師' },
    { heading: '聯絡方式' },
    { lines: ['0912-345-678', 'ming@example.com', '高雄市楠梓區'] },
    { heading: '技能' },
    { bars: [{ label: 'SPC', value: 0.9 }, { label: 'Python', value: 70 }] },
  ],
  main: [
    { heading: '工作經歷' },
    { entries: [{ title: '製程工程師', meta: '強茂　2019–2024', detail: '負責封裝良率改善。' }] },
    { heading: '學歷' },
    { entries: [{ title: '材料工程研究所', meta: '成功大學　2017–2019' }] },
  ],
};

// ── 版面本身 ────────────────────────────────────────────────────────────────
{
  const log = fakeWord();
  const res = await executeTool('word_insert_sidebar', CV);
  ok('有回報', res && res.ok, JSON.stringify(res).slice(0, 120));

  const t = log.tables[0];
  ok('一列兩欄', t && t.rows === 1 && t.cols === 2, t && `${t.rows}x${t.cols}`);

  // 框線關掉——看得到格線就變回一張表格了。
  ok('框線關掉', log.borders.length >= 1 && log.borders[0].where === 'All');

  const side = log.cells.find((q) => q.fill);
  ok('色塊欄有上色', side && side.fill === '#1F3864', side && side.fill);
  ok('色塊欄在左邊', side && side.c === 0, side && side.c);
  ok('色塊欄寬度照預設 150', side && side.width === 150, side && side.width);
  const body = log.cells.find((q) => !q.fill);
  ok('正文欄吃掉剩下的寬度', body && Math.abs(body.width - (TABLE_WIDTH_PT - 150)) < 1,
    body && body.width);

  /**
   * 文字要是真的文字。
   *
   * 兩欄各寫一段 OOXML 進去，裡面要看得到真的 <w:t>。整頁做成圖片會長得一樣，
   * 然後使用者一個字都改不了——那不是版面，是截圖。
   */
  ok('兩欄都寫了內容', log.ooxml.length === 2, `${log.ooxml.length} 次`);
  const asideXml = log.ooxml.find((q) => q.c === 0).xml;
  const mainXml = log.ooxml.find((q) => q.c === 1).xml;
  ok('姓名是真的文字', /<w:t[^>]*>王小明</.test(asideXml));
  ok('聯絡方式一行一條',
    ['0912-345-678', 'ming@example.com', '高雄市楠梓區'].every((x) => asideXml.includes(x)));
  ok('經歷的三段都在',
    /製程工程師/.test(mainXml) && /2019–2024/.test(mainXml) && /封裝良率改善/.test(mainXml));

  // 小標底下要有分隔線，那是這種版型的招牌。
  ok('小標有分隔線', /<w:pBdr><w:bottom/.test(asideXml));
  // 深色欄上的字要是白的，不然看不見。
  ok('深色欄的字是白的', /<w:color w:val="FFFFFF"\/>/.test(asideXml));
  ok('正文欄的字是黑的', /<w:color w:val="000000"\/>/.test(mainXml));

  /**
   * 技能條：畫不出來的時候要留下名字。
   *
   * 轉點陣圖要靠瀏覽器的 canvas，node 裡沒有，所以這裡走的是退路。退路本身才是
   * 要守的東西——畫不出圖就把技能名稱當文字印出來，而不是整段消失。使用者會看到
   * 「少了一塊」，卻不會知道少的是什麼。
   */
  ok('畫不出長條時，技能名稱不會消失',
    /SPC/.test(asideXml) && /Python/.test(asideXml), asideXml.slice(0, 200));

  // 長條本身在瀏覽器裡才畫得出來，這裡只驗 SVG 的算式沒有 NaN。
  const svg = barsSvgFor([{ label: 'SPC', value: 0.9 }, { label: 'Python', value: 0.7 }]);
  ok('長條的 SVG 沒有 NaN', !/NaN|undefined/.test(svg), svg.slice(0, 160));
  const widths = [...svg.matchAll(/<rect[^>]*width="([\d.]+)"/g)].map((m) => Number(m[1]));
  ok('九成的長條比七成的長', widths[1] > widths[3], widths.join(' '));
}

// ── 右側欄 ──────────────────────────────────────────────────────────────────
{
  const log = fakeWord();
  await executeTool('word_insert_sidebar', {
    side: 'right', width: 200, fill: '#2E6DA4', main: [{ heading: '正文' }],
    sidebar: [{ lines: ['一'] }],
  });
  const side = log.cells.find((q) => q.fill);
  ok('色塊欄換到右邊', side && side.c === 1, side && side.c);
  ok('寬度與顏色聽指定的', side && side.width === 200 && side.fill === '#2E6DA4',
    side && `${side.width} ${side.fill}`);
}

// ── 擋下不會用的情況 ────────────────────────────────────────────────────────
{
  fakeWord();
  const bad = await executeTool('word_insert_sidebar', {});
  const msg = String(bad && bad.error);
  ok('兩邊都空會說清楚怎麼給', /entries/.test(msg) && /bars/.test(msg), msg);
}

console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
