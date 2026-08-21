/**
 * Charts as SVG, for slides.
 *
 * ── Why this exists at all ──
 * The PowerPoint JS API has no way to create a chart. Not gated behind a
 * requirement set — absent. The reference deck this was measured against has
 * four native bar charts in `ppt/charts/`, and nothing in Office.js can produce
 * one. So a presentation add-in either draws its own or tells people to insert
 * charts by hand, and the second is not a product.
 *
 * ── Why SVG rather than the canvas the Word host uses ──
 * Word can only take a raster picture, so its charts are PNG and that is the end
 * of it. PowerPoint accepts SVG (ImageCoercion 1.2), and the difference is not
 * cosmetic:
 *
 *   · it stays sharp at any projector resolution, and on a 4K screen a PNG
 *     scaled to half a slide looks soft in a way people notice
 *   · PowerPoint can convert it to real shapes — right-click → 轉換為圖形 — after
 *     which every bar and label is an editable object with the right colours
 *
 * The honest cost, which the tool description states rather than hides: it is
 * not a PowerPoint CHART object. There is no linked data, so changing a number
 * means regenerating, not typing in a datasheet.
 *
 * ── Why it is pure string work ──
 * No canvas, no DOM, no measurement APIs. Everything here is arithmetic and
 * template strings, which means the whole module can be driven by tests with
 * nothing pretending to be a browser — the failure this guards against is a NaN
 * coordinate, and a NaN in SVG is silently dropped by the renderer exactly the
 * way it is silently dropped by canvas.
 */

const xmlEscape = (s) => String(s == null ? '' : s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;').replace(/'/g, '&apos;');

/** Coordinates rounded to 2dp: SVG does not need more, and the file is smaller. */
const n2 = (v) => {
  const x = Number(v);
  return Number.isFinite(x) ? Math.round(x * 100) / 100 : 0;
};

/**
 * A number a person can read at three metres.
 *
 * 1234567 → 1.23M. Axis labels are the one place in a deck where an unrounded
 * number is always wrong: nobody reads 1234567 off a wall, and the digits past
 * the third are noise that makes the label wide enough to collide with its
 * neighbour.
 */
function short(v) {
  const x = Number(v) || 0;
  const a = Math.abs(x);
  if (a >= 1e8) return (x / 1e8).toFixed(a >= 1e9 ? 0 : 1) + '億';
  if (a >= 1e4) return (x / 1e4).toFixed(a >= 1e5 ? 0 : 1) + '萬';
  if (a >= 1000) return (x / 1000).toFixed(a >= 1e4 ? 0 : 1) + 'k';
  return String(Math.round(x * 100) / 100);
}

/**
 * A "nice" axis top — 1, 2, 2.5 or 5 times a power of ten.
 *
 * Without this the axis reads 0 / 3.7 / 7.4 / 11.1, which is technically correct
 * and unreadable. The whole point of a chart on a slide is that the shape is
 * grasped in a second, and an axis nobody can do arithmetic against defeats it.
 */
function niceMax(v) {
  const x = Number(v) || 0;
  if (x <= 0) return 1;
  const mag = Math.pow(10, Math.floor(Math.log10(x)));
  const norm = x / mag;
  const step = norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10;
  return step * mag;
}

const FONT = '"Microsoft JhengHei UI", "Microsoft JhengHei", "PingFang TC", sans-serif';

/**
 * Default series colours.
 *
 * Five, in the order the reference deck uses them, and deliberately not more: a
 * chart that needs a sixth colour is a chart with too many series to read from
 * the back of a room, and the honest fix is fewer series, not more hues.
 */
const SERIES = ['#1A3A5C', '#2E6DA4', '#C0392B', '#8FB4D6', '#8A93A0'];

function text(x, y, s, o = {}) {
  if (s === '' || s == null) return '';
  return `<text x="${n2(x)}" y="${n2(y)}" font-family='${o.font || FONT}'`
    + ` font-size="${o.size || 12}" fill="${o.color || '#3D3D3D'}"`
    + (o.bold ? ' font-weight="600"' : '')
    + ` text-anchor="${o.anchor || 'start'}"`
    + (o.baseline ? ` dominant-baseline="${o.baseline}"` : '')
    + `>${xmlEscape(s)}</text>`;
}
const rect = (x, y, w, h, fill, extra = '') =>
  `<rect x="${n2(x)}" y="${n2(y)}" width="${n2(Math.max(0, w))}" height="${n2(Math.max(0, h))}" fill="${fill}"${extra}/>`;
const line = (x1, y1, x2, y2, stroke, w = 1) =>
  `<line x1="${n2(x1)}" y1="${n2(y1)}" x2="${n2(x2)}" y2="${n2(y2)}" stroke="${stroke}" stroke-width="${w}"/>`;

/**
 * Normalise whatever the model sent into series of equal length.
 *
 * Models send `values: [1,2,3]` for one series and `[[1,2],[3,4]]` for several,
 * and ragged rows when a category is missing a number. All three arrive here and
 * exactly one shape leaves, because every renderer below would otherwise need
 * its own copy of the same defensive code.
 */
function normalise(spec) {
  const cats = (Array.isArray(spec.categories) ? spec.categories : []).map(String);
  let raw = Array.isArray(spec.values) ? spec.values : [];
  if (!raw.length) throw new Error('values 是必要的。');
  if (!Array.isArray(raw[0])) raw = [raw];

  const width = cats.length || Math.max(...raw.map((r) => r.length));
  const series = raw.map((row, i) => ({
    name: String((spec.series || [])[i] || (raw.length > 1 ? `系列 ${i + 1}` : '')),
    color: (spec.colors || [])[i] || SERIES[i % SERIES.length],
    // Padded, and non-numbers become 0 rather than NaN — a NaN coordinate makes
    // the bar silently disappear, which reads as "the add-in did nothing".
    data: Array.from({ length: width }, (_, j) => Number(row[j]) || 0),
  }));
  const labels = Array.from({ length: width }, (_, i) => cats[i] || String(i + 1));
  return { labels, series, width };
}

/** The plot box, once the axis labels and legend have taken their room. */
function frame(spec, opts) {
  const W = Number(spec.width) || 760;
  const H = Number(spec.height) || 380;
  const pad = {
    top: spec.title ? 44 : 16,
    right: 20,
    bottom: opts.legend ? 52 : 34,
    left: opts.axis ? 56 : 16,
  };
  return { W, H, pad, x: pad.left, y: pad.top, w: W - pad.left - pad.right, h: H - pad.top - pad.bottom };
}

function chrome(spec, f, theme) {
  let out = '';
  if (spec.title) {
    out += text(f.x, 26, spec.title, { size: 17, bold: true, color: theme.strong });
  }
  return out;
}

function legend(series, f, theme) {
  if (series.length < 2 && !series[0].name) return '';
  // Centred as a block: a legend that starts at the left edge of a wide plot
  // reads as debris rather than as part of the chart.
  const items = series.filter((s) => s.name);
  if (!items.length) return '';
  const w = items.map((s) => 14 + 6 + s.name.length * 12 + 18);
  let x = f.x + (f.w - w.reduce((a, b) => a + b, 0)) / 2;
  const y = f.H - 16;
  return items.map((s, i) => {
    const out = rect(x, y - 9, 12, 12, s.color, ' rx="2"')
      + text(x + 18, y, s.name, { size: 12.5, color: theme.muted, baseline: 'middle' });
    x += w[i];
    return out;
  }).join('');
}

// ── Renderers ────────────────────────────────────────────────────────────────

function barChart(spec, theme) {
  const { labels, series } = normalise(spec);
  const horizontal = spec.type === 'bar';
  const f = frame(spec, { axis: true, legend: series.length > 1 });
  const top = niceMax(Math.max(...series.flatMap((s) => s.data), 0));
  const TICKS = 4;
  let out = chrome(spec, f, theme);

  // Gridlines behind everything, in the tint colour — visible enough to read a
  // value against, faint enough not to compete with the bars.
  for (let i = 0; i <= TICKS; i++) {
    const v = (top / TICKS) * i;
    if (horizontal) {
      const x = f.x + (f.w / TICKS) * i;
      out += line(x, f.y, x, f.y + f.h, theme.line);
      out += text(x, f.y + f.h + 18, short(v), { size: 11, color: theme.muted, anchor: 'middle' });
    } else {
      const y = f.y + f.h - (f.h / TICKS) * i;
      out += line(f.x, y, f.x + f.w, y, theme.line);
      out += text(f.x - 8, y, short(v), { size: 11, color: theme.muted, anchor: 'end', baseline: 'middle' });
    }
  }

  const groups = labels.length;
  const band = (horizontal ? f.h : f.w) / groups;
  // A quarter of the band as padding, split between the two sides, then divided
  // among the series. Bars that touch read as one block.
  const inner = band * 0.72;
  const bw = inner / series.length;

  labels.forEach((label, g) => {
    const base = (horizontal ? f.y : f.x) + band * g + (band - inner) / 2;
    series.forEach((s, si) => {
      const ratio = top ? s.data[g] / top : 0;
      if (horizontal) {
        const y = base + bw * si;
        out += rect(f.x, y, f.w * ratio, bw * 0.86, s.color, ' rx="2"');
        if (spec.data_labels !== false) {
          out += text(f.x + f.w * ratio + 6, y + bw * 0.43, short(s.data[g]),
            { size: 11, color: theme.ink, baseline: 'middle' });
        }
      } else {
        const h = f.h * ratio;
        const x = base + bw * si;
        out += rect(x, f.y + f.h - h, bw * 0.86, h, s.color, ' rx="2"');
        if (spec.data_labels !== false) {
          out += text(x + bw * 0.43, f.y + f.h - h - 7, short(s.data[g]),
            { size: 11, color: theme.ink, anchor: 'middle' });
        }
      }
    });
    if (horizontal) {
      out += text(f.x - 8, base + inner / 2, label,
        { size: 12, color: theme.ink, anchor: 'end', baseline: 'middle' });
    } else {
      out += text(f.x + band * g + band / 2, f.y + f.h + 20, label,
        { size: 12, color: theme.ink, anchor: 'middle' });
    }
  });

  out += line(f.x, f.y + f.h, f.x + f.w, f.y + f.h, theme.muted, 1.2);
  out += legend(series, f, theme);
  return out;
}

function lineChart(spec, theme) {
  const { labels, series } = normalise(spec);
  const f = frame(spec, { axis: true, legend: series.length > 1 });
  const top = niceMax(Math.max(...series.flatMap((s) => s.data), 0));
  const TICKS = 4;
  let out = chrome(spec, f, theme);

  for (let i = 0; i <= TICKS; i++) {
    const y = f.y + f.h - (f.h / TICKS) * i;
    out += line(f.x, y, f.x + f.w, y, theme.line);
    out += text(f.x - 8, y, short((top / TICKS) * i),
      { size: 11, color: theme.muted, anchor: 'end', baseline: 'middle' });
  }

  // Points sit at band CENTRES, not at the edges: a line starting hard against
  // the axis has its first marker half-clipped.
  const step = labels.length > 1 ? f.w / (labels.length - 1) : 0;
  const px = (i) => (labels.length > 1 ? f.x + step * i : f.x + f.w / 2);
  const py = (v) => f.y + f.h - (top ? (v / top) * f.h : 0);

  series.forEach((s) => {
    const pts = s.data.map((v, i) => `${n2(px(i))},${n2(py(v))}`).join(' ');
    if (spec.area) {
      out += `<polygon points="${n2(px(0))},${n2(f.y + f.h)} ${pts} ${n2(px(s.data.length - 1))},${n2(f.y + f.h)}"`
        + ` fill="${s.color}" opacity="0.12"/>`;
    }
    out += `<polyline points="${pts}" fill="none" stroke="${s.color}" stroke-width="2.5"`
      + ' stroke-linejoin="round" stroke-linecap="round"/>';
    s.data.forEach((v, i) => {
      out += `<circle cx="${n2(px(i))}" cy="${n2(py(v))}" r="3.5" fill="#FFFFFF" stroke="${s.color}" stroke-width="2"/>`;
    });
  });

  labels.forEach((label, i) => {
    out += text(px(i), f.y + f.h + 20, label, { size: 12, color: theme.ink, anchor: 'middle' });
  });
  out += line(f.x, f.y + f.h, f.x + f.w, f.y + f.h, theme.muted, 1.2);
  out += legend(series, f, theme);
  return out;
}

function pieChart(spec, theme) {
  const { labels, series } = normalise(spec);
  const data = series[0].data;
  const total = data.reduce((a, b) => a + b, 0);
  const f = frame(spec, { axis: false, legend: false });
  const donut = spec.type === 'donut';
  let out = chrome(spec, f, theme);

  // The pie sits left of centre so the labels have somewhere to live; centring
  // it and putting labels around the edge is what produces the classic
  // unreadable pie with leader lines crossing each other.
  const r = Math.min(f.w * 0.42, f.h * 0.46);
  const cx = f.x + r + 10;
  const cy = f.y + f.h / 2;

  if (!total) {
    return out + text(cx, cy, '（沒有資料）', { size: 13, color: theme.muted, anchor: 'middle' });
  }

  let angle = -Math.PI / 2;   // start at twelve o'clock, the way people read a pie
  data.forEach((v, i) => {
    const sweep = (v / total) * Math.PI * 2;
    const colour = (spec.colors || [])[i] || SERIES[i % SERIES.length];
    const x1 = cx + r * Math.cos(angle);
    const y1 = cy + r * Math.sin(angle);
    const x2 = cx + r * Math.cos(angle + sweep);
    const y2 = cy + r * Math.sin(angle + sweep);
    // A single slice covering the whole circle cannot be drawn as an arc — the
    // start and end points coincide and the path collapses to nothing.
    if (data.filter((d) => d > 0).length === 1) {
      out += `<circle cx="${n2(cx)}" cy="${n2(cy)}" r="${n2(r)}" fill="${colour}"/>`;
    } else {
      out += `<path d="M ${n2(cx)} ${n2(cy)} L ${n2(x1)} ${n2(y1)}`
        + ` A ${n2(r)} ${n2(r)} 0 ${sweep > Math.PI ? 1 : 0} 1 ${n2(x2)} ${n2(y2)} Z"`
        + ` fill="${colour}"/>`;
    }
    angle += sweep;
  });

  if (donut) {
    out += `<circle cx="${n2(cx)}" cy="${n2(cy)}" r="${n2(r * 0.58)}" fill="${theme.bg || '#FFFFFF'}"/>`;
    if (spec.center) {
      out += text(cx, cy - 4, String(spec.center),
        { size: 26, bold: true, color: theme.strong, anchor: 'middle', baseline: 'middle' });
      if (spec.center_label) {
        out += text(cx, cy + 20, String(spec.center_label),
          { size: 12, color: theme.muted, anchor: 'middle', baseline: 'middle' });
      }
    }
  }

  // Legend as a right-hand list with the percentage — far easier to read than
  // labels scattered round the rim, and it never collides with itself.
  const lx = cx + r + 28;
  const lh = 26;
  let ly = cy - (labels.length * lh) / 2 + lh / 2;
  labels.forEach((label, i) => {
    const colour = (spec.colors || [])[i] || SERIES[i % SERIES.length];
    out += rect(lx, ly - 6, 12, 12, colour, ' rx="2"');
    out += text(lx + 20, ly, label, { size: 12.5, color: theme.ink, baseline: 'middle' });
    out += text(f.W - 16, ly, `${Math.round((data[i] / total) * 100)}%`,
      { size: 12.5, bold: true, color: theme.strong, anchor: 'end', baseline: 'middle' });
    ly += lh;
  });
  return out;
}

const RENDER = { column: barChart, bar: barChart, line: lineChart, area: lineChart, pie: pieChart, donut: pieChart };
export const CHART_TYPES = Object.keys(RENDER);

/**
 * Chart spec → a complete SVG document.
 *
 * Returns the markup as a string. The caller hands it to
 * `setSelectedDataAsync` with `Office.CoercionType.XmlSvg`, which is what puts
 * it on the slide as a vector object.
 */
export function chartSvg(spec = {}, theme = {}) {
  const type = String(spec.type || 'column').toLowerCase();
  const render = RENDER[type];
  if (!render) throw new Error(`不支援的圖表類型：${type}。可用的有 ${CHART_TYPES.join('、')}。`);

  const t = {
    ink: theme.ink || '#3D3D3D',
    strong: theme.strong || '#1A3A5C',
    muted: theme.muted || '#8A93A0',
    line: theme.line || '#D5DBE0',
    bg: theme.bg || '#FFFFFF',
  };
  const W = Number(spec.width) || 760;
  const H = Number(spec.height) || 380;
  const body = render({ ...spec, type, width: W, height: H }, t);

  // width/height in points AND a viewBox: PowerPoint sizes the inserted object
  // from the attributes, and without them it lands at some default and has to be
  // resized by hand.
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}"`
    + ` viewBox="0 0 ${W} ${H}">`
    + rect(0, 0, W, H, spec.transparent ? 'none' : t.bg)
    + body
    + '</svg>';
}

/** Unbalanced or NaN-bearing markup, for the tests. Empty string = fine. */
export function svgProblem(svg) {
  const nan = /(?:^|[^A-Za-z])NaN|Infinity|undefined|="null"/.exec(svg);
  if (nan) return `座標裡有 ${nan[0]}`;
  const stack = [];
  const re = /<(\/?)([A-Za-z][\w:-]*)([^>]*?)(\/?)>/g;
  let m;
  while ((m = re.exec(svg))) {
    const [, closing, name, , selfClose] = m;
    if (selfClose === '/') continue;
    if (closing === '/') { if (stack.pop() !== name) return `</${name}> 對不上`; }
    else stack.push(name);
  }
  return stack.length ? `沒有關閉：${stack.join(', ')}` : '';
}
