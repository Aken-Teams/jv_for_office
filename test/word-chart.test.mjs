/**
 * Drive the Word chart renderer against a recording canvas.
 *
 * Word has no chart API, so the pane draws the picture itself — which means the
 * arithmetic is ours, and the classic way for that arithmetic to fail is a NaN
 * coordinate: canvas silently ignores the call, the PNG comes out blank, and
 * from inside Word it looks exactly like "the add-in did nothing". No amount of
 * reading the code catches an axis that collapses on a particular shape of data,
 * so every awkward shape gets driven through here instead:
 *
 *   all values identical · all zeros · negatives · one category · one series ·
 *   several series · huge and tiny numbers in the same chart
 *
 * The assertion is deliberately blunt. It does not check that the chart is
 * pretty — it checks that every number handed to the canvas is a real number and
 * lands inside the bitmap, which is the difference between a chart and a blank.
 */

/** Records every call and every coordinate, so the test can inspect them. */
function fakeCanvas() {
  const calls = [];
  const numbers = [];
  const track = (name, args) => {
    calls.push({ name, args });
    for (const a of args) if (typeof a === 'number') numbers.push({ name, value: a });
  };
  const ctx = {
    canvas: null,
    fillStyle: '', strokeStyle: '', lineWidth: 1, font: '', textAlign: '', textBaseline: '',
    scale: (...a) => track('scale', a),
    fillRect: (...a) => track('fillRect', a),
    fillText: (t, ...a) => track('fillText', a),
    measureText: (t) => ({ width: String(t).length * 7 }),
    beginPath: () => track('beginPath', []),
    closePath: () => track('closePath', []),
    moveTo: (...a) => track('moveTo', a),
    lineTo: (...a) => track('lineTo', a),
    arc: (...a) => track('arc', a),
    fill: () => track('fill', []),
    stroke: () => track('stroke', []),
  };
  return {
    width: 0, height: 0,
    getContext: () => ctx,
    toDataURL: () => 'data:image/png;base64,AAAA',
    _calls: calls,
    _numbers: numbers,
  };
}

let made = [];
globalThis.document = {
  createElement(tag) {
    if (tag !== 'canvas') throw new Error('unexpected element: ' + tag);
    const c = fakeCanvas();
    made.push(c);
    return c;
  },
};

const { drawChart } = await import('../src/taskpane/hosts/word.js');

const CASES = [
  ['直條圖，一般資料', {
    type: 'column', title: '各季營收', categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series: [{ name: '營收', values: [120, 340, 280, 410] }],
  }],
  ['橫條圖，多數列', {
    type: 'bar', title: '', categories: ['台灣', '美國', '歐洲'],
    series: [{ name: '2025', values: [12, 8, 5] }, { name: '2026', values: [18, 11, 9] }],
  }],
  ['折線圖，三個數列', {
    type: 'line', title: '趨勢', categories: ['一月', '二月', '三月', '四月'],
    series: [
      { name: 'A', values: [1, 2, 3, 4] },
      { name: 'B', values: [4, 3, 2, 1] },
      { name: 'C', values: [2, 2, 2, 2] },
    ],
  }],
  ['圓餅圖', {
    type: 'pie', title: '佔比', categories: ['甲', '乙', '丙'],
    series: [{ name: '佔比', values: [50, 30, 20] }],
  }],
  // ── The shapes that break axis arithmetic ──
  ['所有值相同', {
    type: 'column', title: '', categories: ['a', 'b', 'c'],
    series: [{ name: 'x', values: [7, 7, 7] }],
  }],
  ['全部是零', {
    type: 'column', title: '', categories: ['a', 'b'],
    series: [{ name: 'x', values: [0, 0] }],
  }],
  ['含負數', {
    type: 'column', title: '盈虧', categories: ['a', 'b', 'c'],
    series: [{ name: 'x', values: [-50, 120, -10] }],
  }],
  ['只有一個分類', {
    type: 'column', title: '', categories: ['只有這個'],
    series: [{ name: 'x', values: [42] }],
  }],
  ['數量級差很多', {
    type: 'line', title: '', categories: ['a', 'b', 'c'],
    series: [{ name: 'x', values: [0.5, 1, 999999] }],
  }],
  ['圓餅圖全零', {
    type: 'pie', title: '', categories: ['甲', '乙'],
    series: [{ name: 'x', values: [0, 0] }],
  }],
];

let pass = 0, fail = 0;
for (const [name, spec] of CASES) {
  made = [];
  let bad = null;
  try {
    const canvas = drawChart(spec);
    if (!canvas || !canvas.width || !canvas.height) bad = '畫布沒有尺寸';
    const nums = canvas._numbers;
    const nan = nums.find(n => !Number.isFinite(n.value));
    if (!bad && nan) bad = `${nan.name}() 收到 ${nan.value}`;
    // Anything drawn a long way outside the bitmap is a geometry bug even when
    // it is a finite number — it just will not be in the picture.
    const off = nums.find(n => Math.abs(n.value) > 5000);
    if (!bad && off) bad = `${off.name}() 的座標跑到 ${Math.round(off.value)}，畫布只有 ${canvas.width / 2}`;
    if (!bad && canvas._calls.length < 5) bad = '幾乎沒畫東西';
  } catch (e) {
    bad = '丟出例外：' + e.message;
  }
  bad ? fail++ : pass++;
  console.log(`${bad ? 'FAIL' : 'PASS'}  ${name}${bad ? '  → ' + bad : ''}`);
}

// The three shapes the tool refuses outright are checked at the executeTool
// layer, not here; what this file pins is that anything that DOES reach the
// renderer comes out as a picture rather than as a blank rectangle.
console.log(`\n${pass} passed, ${fail} failed`);
process.exit(fail ? 1 : 0);
