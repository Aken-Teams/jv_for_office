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

import { MAPS } from './mapdata.js';

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
  /**
   * Room for the axis labels — and a horizontal bar chart needs far more of it.
   *
   * ── The bug this fixes ──
   * 56 points is right for a column chart, where the left margin holds numbers
   * like 「75」. On a BAR chart the same margin holds the category names, and
   * 「顯著」回報」 arrived as 「著」回報」 — silently clipped, with the missing half
   * looking like a label somebody typed wrong.
   *
   * So the margin is measured against the longest label rather than fixed. A CJK
   * glyph is about as wide as the type is tall; latin runs narrower, so the
   * estimate splits the difference and caps at a third of the canvas — past that
   * the labels are the chart and the answer is shorter names.
   */
  const labelRoom = () => {
    const cats = (Array.isArray(spec.categories) ? spec.categories : []).map(String);
    if (!cats.length) return 56;
    const widest = Math.max(...cats.map((c) => {
      const cjk = (c.match(/[\u3000-\u9fff\uff00-\uffef]/g) || []).length;
      return cjk * 11.5 + (c.length - cjk) * 6.5;
    }));
    return Math.min(Math.round(widest) + 16, Math.round((Number(spec.width) || 760) / 3));
  };
  const pad = {
    top: spec.title ? 44 : 16,
    right: 20,
    bottom: opts.legend ? 52 : 34,
    left: opts.axis ? (opts.wideLabels ? Math.max(56, labelRoom()) : 56) : 16,
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
  const f = frame(spec, { axis: true, legend: series.length > 1, wideLabels: horizontal });
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

/**
 * A waterfall — how a number got from one value to another.
 *
 * ── Why a column chart cannot do this job ──
 * The question a waterfall answers is 「這中間發生了什麼」: revenue was 100, it is
 * now 118, and here are the four things that moved it. Plotted as columns those
 * four things are four unrelated bars and the reader has to do the running total
 * in their head. The floating bar IS the explanation — each one starts where the
 * last one finished, so the path from start to end is a shape rather than an
 * arithmetic exercise.
 *
 * Positive and negative get different colours because that is the one thing the
 * eye should get for free, and totals get a third because a total is not a
 * change — it is a resting point, and drawing it like a change is the classic
 * way to make a waterfall lie.
 *
 * Mark a bar as a total with `total: [0, 5]` (indices) or by giving its category
 * a leading `=`.
 */
function waterfallChart(spec, theme) {
  const { labels, series } = normalise(spec);
  const data = series[0].data;
  const f = frame(spec, { axis: true, legend: false });
  let out = chrome(spec, f, theme);

  const totals = new Set((Array.isArray(spec.total) ? spec.total : []).map(Number));
  labels.forEach((l, i) => { if (/^=/.test(l)) totals.add(i); });
  const clean = labels.map((l) => String(l).replace(/^=/, ''));

  // Walk the series once to find where each bar starts and ends, and how far the
  // running total ever travels — the axis has to cover the PATH, not the values.
  const steps = [];
  let run = 0;
  let lo = 0;
  let hi = 0;
  data.forEach((v, i) => {
    if (totals.has(i)) {
      steps.push({ from: 0, to: v, kind: 'total' });
      run = v;
    } else {
      steps.push({ from: run, to: run + v, kind: v < 0 ? 'down' : 'up' });
      run += v;
    }
    lo = Math.min(lo, steps[i].from, steps[i].to);
    hi = Math.max(hi, steps[i].from, steps[i].to);
  });

  const max = niceMax(hi);
  const min = lo < 0 ? -niceMax(-lo) : 0;
  const span = (max - min) || 1;
  const yOf = (v) => f.y + f.h - ((v - min) / span) * f.h;

  for (let g = 0; g <= 4; g++) {
    const y = f.y + (f.h / 4) * g;
    out += line(f.x, y, f.x + f.w, y, theme.line || '#E4E8EC', 1);
    out += text(f.x - 10, y, short(max - (span / 4) * g),
      { size: 11, color: theme.muted || '#8A93A0', anchor: 'end', baseline: 'middle' });
  }

  const slot = f.w / Math.max(1, data.length);
  const bw = Math.min(64, slot * 0.6);
  const up = (spec.colors || [])[0] || theme.strong || SERIES[0];
  const down = (spec.colors || [])[1] || theme.accent || SERIES[2];
  const tot = (spec.colors || [])[2] || theme.second || SERIES[1];

  steps.forEach((st, i) => {
    const cx = f.x + slot * i + slot / 2;
    const top = yOf(Math.max(st.from, st.to));
    const bot = yOf(Math.min(st.from, st.to));
    const fill = st.kind === 'total' ? tot : st.kind === 'down' ? down : up;
    // A step that changes nothing still needs to be visible, or the chart has a
    // gap where a category should be.
    out += rect(cx - bw / 2, top, bw, Math.max(2, bot - top), fill);
    // The connector: what makes the bars read as one path instead of a row.
    if (i < steps.length - 1 && !totals.has(i + 1)) {
      const y = yOf(st.to);
      out += line(cx + bw / 2, y, f.x + slot * (i + 1) + slot / 2 - bw / 2, y,
        theme.muted || '#8A93A0', 1);
    }
    const v = st.kind === 'total' ? st.to : st.to - st.from;
    out += text(cx, top - 8, (v > 0 && st.kind !== 'total' ? '+' : '') + short(v),
      { size: 11.5, bold: true, color: fill, anchor: 'middle' });
    out += text(cx, f.y + f.h + 18, clean[i],
      { size: 11.5, color: theme.ink || '#3D3D3D', anchor: 'middle' });
  });
  return out;
}

/**
 * A funnel — how many survive each stage.
 *
 * ── The one thing that makes it worth drawing ──
 * The conversion rate between consecutive stages, printed between them. A funnel
 * without those is a decorative bar chart: the shape tells you it narrows, which
 * the reader already assumed, and not where it narrows WORST. The drop is the
 * finding, so it gets a label and it gets one whether or not anybody asked.
 */
function funnelChart(spec, theme) {
  const { labels, series } = normalise(spec);
  const data = series[0].data;
  const f = frame(spec, { axis: false, legend: false });
  let out = chrome(spec, f, theme);

  const top = Math.max(...data) || 1;
  const rows = data.length;
  const gap = 12;
  const h = Math.max(18, (f.h - gap * (rows - 1)) / rows);
  // Room on the right for the stage name and its count, on the left for the drop.
  const labelW = 220;
  const plotW = f.w - labelW;

  data.forEach((v, i) => {
    const y = f.y + (h + gap) * i;
    const w = Math.max(8, (v / top) * plotW);
    const cx = f.x + plotW / 2;
    const tone = (spec.colors || [])[i]
      || mixHex(theme.strong || SERIES[0], theme.bg || '#FFFFFF', i / Math.max(1, rows));
    out += rect(cx - w / 2, y, w, h, tone, ' rx="3"');
    out += text(cx, y + h / 2, short(v),
      { size: 13, bold: true, color: '#FFFFFF', anchor: 'middle', baseline: 'middle' });
    out += text(f.x + plotW + 20, y + h / 2 - 7, labels[i],
      { size: 12.5, bold: true, color: theme.ink || '#3D3D3D', baseline: 'middle' });

    if (i > 0) {
      const prev = data[i - 1] || 0;
      const keep = prev ? (v / prev) * 100 : 0;
      // Printed in the gap, which is the only place it belongs: it describes the
      // transition, not either stage.
      out += text(f.x + plotW + 20, y + h / 2 + 10,
        `轉換 ${keep.toFixed(0)}%　流失 ${(100 - keep).toFixed(0)}%`,
        { size: 11, color: theme.muted || '#8A93A0', baseline: 'middle' });
    }
  });
  return out;
}

/** Blend two hex colours — the funnel fades its stages through one hue. */
function mixHex(a, b, t) {
  const hex = (c) => {
    const v = String(c || '').replace('#', '');
    return [0, 2, 4].map((i) => parseInt(v.slice(i, i + 2), 16) || 0);
  };
  const [r1, g1, b1] = hex(a);
  const [r2, g2, b2] = hex(b);
  const k = Math.max(0, Math.min(0.62, Number(t) || 0));
  const to = (x, y) => Math.round(x + (y - x) * k).toString(16).padStart(2, '0');
  return `#${to(r1, r2)}${to(g1, g2)}${to(b1, b2)}`;
}

/**
 * A radar — one thing profiled across several axes, or two things compared.
 *
 * ── Where this earns its place, and where it does not ──
 * It is the right chart for 「我們在這六個面向上各多強」 and the wrong chart for
 * almost everything else: areas on a radar are not comparable between axes, so
 * anybody reading magnitudes off it is being misled. Three series is already too
 * many to separate, so the spec caps it — a fourth ring would look like data and
 * function as decoration.
 */
function radarChart(spec, theme) {
  const { labels, series } = normalise(spec);
  const f = frame(spec, { axis: false, legend: series.length > 1 });
  let out = chrome(spec, f, theme);

  const n = labels.length;
  if (n < 3) throw new Error('雷達圖至少要三個面向，不然畫出來是一條線。');
  const cx = f.x + f.w / 2;
  const cy = f.y + f.h / 2;
  const r = Math.min(f.w, f.h) / 2 - 34;
  const max = niceMax(Math.max(...series.flatMap((s) => s.data)));
  // Start at twelve o'clock: a radar whose first axis is on the right reads as
  // rotated, and nobody can say why.
  const at = (i, frac) => {
    const a = (Math.PI * 2 * i) / n - Math.PI / 2;
    return [cx + Math.cos(a) * r * frac, cy + Math.sin(a) * r * frac];
  };

  for (let ring = 1; ring <= 4; ring++) {
    const pts = Array.from({ length: n }, (_, i) => at(i, ring / 4).map(n2).join(','));
    out += `<polygon points="${pts.join(' ')}" fill="none"`
      + ` stroke="${theme.line || '#E4E8EC'}" stroke-width="1"/>`;
  }
  for (let i = 0; i < n; i++) {
    const [x, y] = at(i, 1);
    out += line(cx, cy, x, y, theme.line || '#E4E8EC', 1);
    const [lx, ly] = at(i, 1.16);
    out += text(lx, ly, labels[i], {
      size: 11.5, color: theme.ink || '#3D3D3D', baseline: 'middle',
      anchor: Math.abs(lx - cx) < 4 ? 'middle' : lx > cx ? 'start' : 'end',
    });
  }

  series.slice(0, 3).forEach((sr) => {
    const pts = sr.data.map((v, i) => at(i, Math.max(0, v) / max).map(n2).join(','));
    out += `<polygon points="${pts.join(' ')}" fill="${sr.color}" fill-opacity="0.18"`
      + ` stroke="${sr.color}" stroke-width="2" stroke-linejoin="round"/>`;
    sr.data.forEach((v, i) => {
      const [x, y] = at(i, Math.max(0, v) / max);
      out += `<circle cx="${n2(x)}" cy="${n2(y)}" r="3" fill="${sr.color}"/>`;
    });
  });

  if (series.length > 1) {
    let lx = f.x;
    series.slice(0, 3).forEach((sr) => {
      out += rect(lx, f.H - f.pad.bottom + 16, 12, 12, sr.color, ' rx="2"');
      out += text(lx + 18, f.H - f.pad.bottom + 26, sr.name,
        { size: 12, color: theme.ink || '#3D3D3D', baseline: 'middle' });
      lx += 28 + String(sr.name).length * 12;
    });
  }
  return out;
}

/**
 * A timeline — events on one axis, in order.
 *
 * ── Why this is a chart and not a diagram ──
 * It could be drawn with shapes, and that is what a mermaid flowchart would do:
 * boxes joined by arrows, spaced evenly. But a timeline's whole content is that
 * the gaps are NOT even — two things in March and one in November is the point,
 * and even spacing erases it. Placing marks against a measured axis is chart
 * work, so it lives here.
 *
 * Labels alternate above and below the line. Eight events on one side collide;
 * alternating halves the density for free, and it is what every printed timeline
 * does for the same reason.
 */
function timelineChart(spec, theme) {
  const f = frame(spec, { axis: false, legend: false });
  let out = chrome(spec, f, theme);

  const items = (Array.isArray(spec.events) ? spec.events : []).filter(Boolean);
  if (!items.length) throw new Error('events 是必要的：[{ label, at, note }]。');

  const y = f.y + f.h / 2;
  const x0 = f.x + 24;
  const x1 = f.x + f.w - 24;
  // Numbers place marks proportionally; anything else falls back to even spacing,
  // which is honest — we cannot invent a scale we were not given.
  const nums = items.map((e) => Number(e.at));
  const scaled = nums.every((v) => Number.isFinite(v)) && new Set(nums).size > 1;
  const lo = scaled ? Math.min(...nums) : 0;
  const hi = scaled ? Math.max(...nums) : 1;
  const xOf = (i) => (scaled
    ? x0 + ((nums[i] - lo) / (hi - lo)) * (x1 - x0)
    : x0 + (items.length === 1 ? (x1 - x0) / 2 : ((x1 - x0) / (items.length - 1)) * i));

  out += line(x0 - 12, y, x1 + 12, y, theme.line || '#E4E8EC', 3);
  const done = (spec.colors || [])[0] || theme.strong || SERIES[0];
  const mark = (spec.colors || [])[1] || theme.accent || SERIES[2];

  items.forEach((e, i) => {
    const x = xOf(i);
    const hot = e.highlight === true;
    const tone = hot ? mark : done;
    const above = i % 2 === 0;
    const stem = 26;
    out += line(x, y, x, above ? y - stem : y + stem, tone, 1.5);
    out += `<circle cx="${n2(x)}" cy="${n2(y)}" r="${hot ? 8 : 6}" fill="${tone}"`
      + ` stroke="${theme.bg || '#FFFFFF'}" stroke-width="2.5"/>`;

    const ly = above ? y - stem - 8 : y + stem + 16;
    out += text(x, ly, String(e.at ?? ''),
      { size: 11, bold: true, color: tone, anchor: 'middle' });
    out += text(x, above ? ly - 16 : ly + 16, String(e.label ?? ''),
      { size: 12.5, bold: true, color: theme.ink || '#3D3D3D', anchor: 'middle' });
    if (e.note) {
      out += text(x, above ? ly - 32 : ly + 32, String(e.note),
        { size: 11, color: theme.muted || '#8A93A0', anchor: 'middle' });
    }
  });
  return out;
}

/**
 * A choropleth — regions shaded by a number.
 *
 * ── What a map is actually for on a slide, and what it is not ──
 * It answers 「哪裡多、哪裡少」 in one glance, and it answers nothing else well.
 * Reading a precise value off a shaded country is impossible, and comparing two
 * similar shades across a page is worse than comparing two bars. So the legend
 * carries the numbers and the map carries the pattern — and when the question is
 * really 「排名」, a bar chart is the honest answer and the tool description says
 * so.
 *
 * ── The projection ──
 * Equirectangular, scaled to fit the box with the aspect ratio preserved. Not
 * because it is a good projection — it stretches the poles badly — but because
 * it is the one whose distortion everybody has already seen a thousand times,
 * and a slide is not the place to explain why Greenland changed shape.
 *
 * ── Names ──
 * The world data is keyed by English name, the two Chinese maps by Chinese name.
 * A region nobody supplied a value for is drawn in the empty tone rather than
 * skipped: a hole where Vietnam should be reads as a rendering fault, while a
 * pale Vietnam reads as 「no data」, which is what it means.
 */
function mapChart(spec, theme) {
  const which = String(spec.map || 'world').toLowerCase();
  const geo = MAPS[which];
  if (!geo) {
    throw new Error(`沒有「${spec.map}」這份地圖。可用的有：${Object.keys(MAPS).join('、')}。`);
  }

  const values = new Map();
  const cats = Array.isArray(spec.categories) ? spec.categories.map(String) : [];
  const nums = Array.isArray(spec.values) ? spec.values : [];
  cats.forEach((c, i) => {
    const v = Number(nums[i]);
    if (Number.isFinite(v)) values.set(c, v);
  });
  // Named but unknown is worth saying out loud: a typo in a region name would
  // otherwise just be a country that never lights up.
  const unknown = cats.filter((c) => !geo[c]);

  const f = frame(spec, { axis: false, legend: false });
  let out = chrome(spec, f, theme);

  // The bounding box of everything being drawn, so a China map fills its box
  // instead of sitting in the middle of a world-sized one.
  let x0 = 180;
  let y0 = 90;
  let x1 = -180;
  let y1 = -90;
  for (const rings of Object.values(geo)) {
    for (const r of rings) {
      for (const [lon, lat] of r) {
        if (lon < x0) x0 = lon;
        if (lon > x1) x1 = lon;
        if (lat < y0) y0 = lat;
        if (lat > y1) y1 = lat;
      }
    }
  }
  const legendW = values.size ? 150 : 0;
  const plotW = f.w - legendW;
  const scale = Math.min(plotW / ((x1 - x0) || 1), f.h / ((y1 - y0) || 1));
  const offX = f.x + (plotW - (x1 - x0) * scale) / 2;
  const offY = f.y + (f.h - (y1 - y0) * scale) / 2;
  // Latitude increases northward and SVG's y increases downward.
  const px = (lon) => offX + (lon - x0) * scale;
  const py = (lat) => offY + (y1 - lat) * scale;

  const vals = [...values.values()];
  const lo = vals.length ? Math.min(...vals) : 0;
  const hi = vals.length ? Math.max(...vals) : 1;
  const strong = (spec.colors || [])[0] || theme.strong || SERIES[0];
  const empty = theme.tint || '#EEF3F8';
  const shade = (v) => {
    if (v === undefined) return empty;
    // A floor of 0.18 so the lowest region is still visibly part of the data
    // rather than indistinguishable from one with no value at all.
    const t = hi === lo ? 1 : 0.18 + 0.82 * ((v - lo) / (hi - lo));
    return mixHex(empty, strong, t);
  };

  for (const [name, rings] of Object.entries(geo)) {
    const fill = shade(values.get(name));
    const d = rings.map((r) => `M${r.map(([lon, lat]) => `${n2(px(lon))},${n2(py(lat))}`).join('L')}Z`)
      .join('');
    out += `<path d="${d}" fill="${fill}" stroke="${theme.bg || '#FFFFFF'}"`
      + ` stroke-width="0.6" stroke-linejoin="round"/>`;
  }

  /**
   * The legend is the numbers.
   *
   * Nobody reads a value off a shade, so the regions that were given one are
   * listed with it — largest first, because the question a map raises is always
   * 「誰最多」. Capped at eight: past that the list is the chart.
   */
  if (values.size) {
    const rows = [...values.entries()].sort((a, b) => b[1] - a[1]).slice(0, 8);
    const lx = f.x + plotW + 16;
    let ly = f.y + 8;
    for (const [name, v] of rows) {
      out += rect(lx, ly, 12, 12, shade(v), ' rx="2"');
      out += text(lx + 19, ly + 6, name,
        { size: 12, color: theme.ink || '#3D3D3D', baseline: 'middle' });
      out += text(f.W - 16, ly + 6, short(v),
        { size: 12, bold: true, color: theme.strong || SERIES[0], anchor: 'end', baseline: 'middle' });
      ly += 22;
    }
    if (values.size > rows.length) {
      out += text(lx, ly + 6, `（另有 ${values.size - rows.length} 個未列出）`,
        { size: 10.5, color: theme.muted || '#8A93A0', baseline: 'middle' });
      ly += 20;
    }
    if (unknown.length) {
      out += text(lx, ly + 6, `找不到：${unknown.slice(0, 3).join('、')}`,
        { size: 10.5, color: theme.muted || '#8A93A0', baseline: 'middle' });
    }
  }
  return out;
}

const RENDER = {
  column: barChart, bar: barChart, line: lineChart, area: lineChart,
  pie: pieChart, donut: pieChart,
  waterfall: waterfallChart, funnel: funnelChart, radar: radarChart,
  timeline: timelineChart, map: mapChart,
};
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
