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

/**
 * Is a legend worth the space it takes?
 *
 * ── The bug this is the fix for ──
 * `frame()` reserved legend room when `series.length > 1`, while `legend()`
 * drew one whenever a series had a NAME — so a single named series got a legend
 * with nowhere to put it, and it landed on top of the category labels. That is
 * the overlapping 「加速倍數」 sitting across 「N=1,024」 and 「N=4,096」 in the
 * report: not a layout that ran out of room, two functions disagreeing about
 * whether there was a legend at all.
 *
 * One predicate now, used by both, so they cannot drift again. And the answer
 * for one series is no: the chart's title already says what the bars are, and a
 * legend that names the only thing on the chart is a swatch beside a restatement.
 */
function wantsLegend(series) {
  return series.filter((s) => s && s.name).length >= 2;
}

function legend(series, f, theme) {
  if (!wantsLegend(series)) return '';
  // Centred as a block: a legend that starts at the left edge of a wide plot
  // reads as debris rather than as part of the chart.
  const items = series.filter((s) => s.name);
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
  const f = frame(spec, { axis: true, legend: wantsLegend(series), wideLabels: horizontal });
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
  const f = frame(spec, { axis: true, legend: wantsLegend(series) });
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
    let v = String(c || '').replace('#', '');
    // 「#FFF」 is legal CSS and parsed to garbage by a naive slice: 'FF', 'F', ''
    // gives NaN on the last two channels, and the blend comes out a muddy
    // brown that looks like a deliberate colour choice.
    if (v.length === 3) v = v.split('').map((ch) => ch + ch).join('');
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
  const f = frame(spec, { axis: false, legend: wantsLegend(series) });
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

  /**
   * How much horizontal room each label gets, and where it may sit.
   *
   * ── The two ways this overflowed ──
   * 「跑版這塊要修正」, with a timeline whose first and last labels ran off the
   * canvas and whose middle ones ran into each other. Both came from drawing
   * every label centred on its mark at whatever width it happened to be:
   *
   *   · an edge mark is 24 points from the edge, so half of a 140-point label
   *     is simply outside the picture;
   *   · a 20-character Chinese label is 250 points wide, and five marks across
   *     640 points are 128 apart.
   *
   * ── Why a per-label slot and not one shared width ──
   * The obvious fix is `room = average gap × 2`, since alternating sides means
   * a label's nearest same-side neighbour is two marks away. It is wrong for
   * exactly the reason this chart exists: the gaps are NOT even. 1750, 1805 and
   * 1822 put marks one and three only 72 years apart on a 215-year axis, so the
   * average gap promises each of them twice the room it actually has, and the
   * two labels above the line overlap while the average says they fit.
   *
   * So every label gets its own slot, bounded by the midpoint to its same-side
   * neighbours and by the canvas at the ends. The label wraps to its slot and is
   * then centred on its mark, nudged back inside if centring would cross a
   * boundary. Nothing can reach into another label's slot, whatever the dates.
   */
  const X = items.map((_, i) => xOf(i));
  const edgeL = f.x;
  const edgeR = f.x + f.w;

  /**
   * Try the labels at their natural width first.
   *
   * Cutting every label down to the midpoint between its neighbours is safe but
   * mean: it hands half the gap to a neighbour that may not want it, so a label
   * breaks into two lines with 50 points of white space beside it. `packLabels`
   * asks the real question instead — does this label need to reach further right
   * than its neighbour needs to reach left? — and only when the answer is yes
   * does anything get wrapped.
   */
  const natural = items.map((e) => Math.max(
    textW(String(e.label ?? ''), 12.5),
    e.note ? textW(String(e.note), 11) : 0,
  ));
  const layout = new Array(items.length);
  for (const parity of [0, 1]) {
    const side = items.map((_, i) => i).filter((i) => i % 2 === parity);
    if (!side.length) continue;
    const packed = packLabels(side.map((i) => X[i]), side.map((i) => natural[i]),
      edgeL, edgeR, 12);
    if (packed) {
      // Half a point of slack: the packer measured the whole string at once and
      // the wrapper adds up its pieces, so the same label can come out a
      // rounding error wider and break into two lines for no visible reason.
      side.forEach((i, k) => { layout[i] = { room: natural[i] + 0.5, cx: packed[k] }; });
      continue;
    }
    // Too much text for the axis: fall back to fixed slots bounded by the
    // midpoint to each same-side neighbour, and wrap into them. Nothing can
    // reach into another label's slot after this, whatever the dates are.
    for (const i of side) {
      const prev = i >= 2 ? X[i - 2] : null;
      const next = i + 2 < items.length ? X[i + 2] : null;
      const lo = prev === null ? edgeL : (X[i] + prev) / 2;
      const hi = next === null ? edgeR : (X[i] + next) / 2;
      layout[i] = { room: Math.max(64, Math.min(hi - lo, f.w * 0.45)), lo, hi };
    }
  }

  items.forEach((e, i) => {
    const x = xOf(i);
    const hot = e.highlight === true;
    const tone = hot ? mark : done;
    const above = i % 2 === 0;
    const stem = 26;
    out += line(x, y, x, above ? y - stem : y + stem, tone, 1.5);
    out += `<circle cx="${n2(x)}" cy="${n2(y)}" r="${hot ? 8 : 6}" fill="${tone}"`
      + ` stroke="${theme.bg || '#FFFFFF'}" stroke-width="2.5"/>`;

    const L = layout[i];
    /**
     * Where a block of text that wide ends up.
     *
     * A packed label already has its centre. A wrapped one is centred on its
     * mark and then nudged back inside its slot — the mark cannot move, since
     * its position IS the date, so the text slides along instead and stays
     * beside its own stem rather than drifting onto a neighbour's.
     */
    const place = (w) => {
      if (L.cx != null) return L.cx;
      const half = w / 2;
      const lo = Math.max(L.lo, edgeL);
      const hi = Math.max(L.hi, lo);
      if (w >= hi - lo) return lo + half;
      return Math.min(Math.max(x, lo + half), hi - half);
    };

    const ly = above ? y - stem - 8 : y + stem + 16;
    // The year stays over its own stem even when the label block has slid away:
    // it is 24 points wide, it always fits, and it is what ties the text back to
    // the mark.
    const yw = textW(String(e.at ?? ''), 11) / 2;
    out += text(Math.min(Math.max(x, edgeL + yw), edgeR - yw), ly, String(e.at ?? ''),
      { size: 11, bold: true, color: tone, anchor: 'middle' });

    // Labels run away from the axis: upward ones stack upward, downward ones
    // downward, so neither grows back across the line.
    const lines = wrapLabel(String(e.label ?? ''), L.room, 12.5);
    const lx = place(Math.max(0, ...lines.map((ln) => textW(ln, 12.5))));
    const step = 15;
    let cursor = above ? ly - 16 : ly + 16;
    for (const ln of (above ? lines.slice().reverse() : lines)) {
      out += text(lx, cursor, ln,
        { size: 12.5, bold: true, color: theme.ink || '#3D3D3D', anchor: 'middle' });
      cursor += above ? -step : step;
    }
    if (e.note) {
      const nl = wrapLabel(String(e.note), L.room, 11).slice(0, 2);
      const nx = place(Math.max(0, ...nl.map((ln) => textW(ln, 11))));
      for (const ln of nl) {
        out += text(nx, cursor, ln,
          { size: 11, color: theme.muted || '#8A93A0', anchor: 'middle' });
        cursor += above ? -13 : 13;
      }
    }
  });
  return out;
}

/**
 * Place labels at their preferred spots, pushed apart only as far as needed.
 *
 * Two sweeps, which is how every map label and axis tick is packed: left to
 * right, each box shoved right until it clears the one before it; then right to
 * left, each box shoved back left until it clears the canvas edge and the box
 * after it. Whatever survives both sweeps is a layout with no overlap and no
 * overflow. If something cannot survive them, the text genuinely does not fit at
 * this width and `null` says so, which is the caller's signal to wrap instead of
 * drawing the pile.
 */
function packLabels(spots, widths, lo, hi, gap) {
  const n = spots.length;
  const pos = spots.slice();
  for (let i = 0; i < n; i++) {
    const least = i === 0
      ? lo + widths[i] / 2
      : pos[i - 1] + widths[i - 1] / 2 + gap + widths[i] / 2;
    pos[i] = Math.max(pos[i], least);
  }
  for (let i = n - 1; i >= 0; i--) {
    const most = i === n - 1
      ? hi - widths[i] / 2
      : pos[i + 1] - widths[i + 1] / 2 - gap - widths[i] / 2;
    pos[i] = Math.min(pos[i], most);
  }
  for (let i = 0; i < n; i++) {
    if (pos[i] - widths[i] / 2 < lo - 0.5) return null;
    if (pos[i] + widths[i] / 2 > hi + 0.5) return null;
    if (i && pos[i] - widths[i] / 2 < pos[i - 1] + widths[i - 1] / 2 + gap - 0.5) return null;
  }
  return pos;
}

/**
 * How wide a string is, near enough to lay out against.
 *
 * A CJK glyph is square — as wide as the type is tall — while latin averages a
 * bit over half that. There is no font metric available in an SVG string, so the
 * same estimate has to be used by whatever measures and whatever places, or a
 * label fits by one function's arithmetic and overflows by the other's.
 */
function textW(str, size) {
  return [...String(str || '')].reduce((a, c) => a
    + (/[\u2E80-\uA4CF\uAC00-\uD7A3\uF900-\uFAFF\uFE10-\uFE6F\uFF00-\uFF60]/.test(c)
      ? size : size * 0.55), 0);
}

/**
 * Break a label into lines that fit the room it has.
 *
 * Breaks between Chinese characters anywhere and between Latin words only —
 * splitting 「Lanczos」 across two lines reads as a typo, while Chinese has no
 * word spaces and is broken at any character by every typesetter there is.
 *
 * Capped at three lines with an ellipsis: a label tall enough to reach the axis
 * is a label that should have been shorter, and silently drawing it anyway is
 * how a timeline turns into the pile of overlapping text this exists to stop.
 */
function wrapLabel(str, room, size) {
  const text_ = String(str || '').trim();
  if (!text_) return [];
  // Latin runs stay whole; every other character is its own breakable unit.
  const units = text_.match(/[A-Za-z0-9][A-Za-z0-9.'-]*|\s+|[\s\S]/g) || [];
  const lines = [];
  let line = '';
  let w = 0;
  for (const u of units) {
    const uw = textW(u, size);
    if (w + uw > room && line) {
      lines.push(line.trim());
      line = /^\s+$/.test(u) ? '' : u;
      w = /^\s+$/.test(u) ? 0 : uw;
    } else {
      line += u;
      w += uw;
    }
    if (lines.length === 3) break;
  }
  if (line.trim() && lines.length < 3) lines.push(line.trim());
  if (lines.length === 3 && w > room) lines[2] = `${lines[2].slice(0, -1)}…`;
  return lines;
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
  const asked = String(spec.map || 'world').toLowerCase();
  if (!MAPS[asked]) {
    throw new Error(`沒有「${spec.map}」這份地圖。可用的有：${Object.keys(MAPS).join('、')}。`);
  }

  /**
   * Which map the names are actually from.
   *
   * ── The failure this replaces ──
   * A slide about Chinese provinces came out as a world map with everything
   * blank and a footnote reading 「找不到：广东、山东、河南」. The tool had done
   * exactly what it was told — `map` defaulted to world and nobody said
   * otherwise — and the result was a picture of the whole planet illustrating a
   * page about one country.
   *
   * The names say which map they belong to, so they are asked. Only when the
   * requested map matches NOTHING and another matches most of them: a deliberate
   * world map with one or two unrecognised names stays a world map.
   */
  const hits = (m) => (Array.isArray(spec.categories) ? spec.categories : [])
    .filter((c) => MAPS[m][String(c)]).length;
  let which = asked;
  const total = (Array.isArray(spec.categories) ? spec.categories : []).length;
  if (total && hits(asked) === 0) {
    const better = Object.keys(MAPS)
      .map((m) => [m, hits(m)])
      .sort((a, b) => b[1] - a[1])[0];
    if (better && better[1] > total / 2) which = better[0];
  }
  const geo = MAPS[which];
  const swapped = which !== asked;

  const cats = Array.isArray(spec.categories) ? spec.categories.map(String) : [];
  const nums = Array.isArray(spec.values) ? spec.values : [];

  /**
   * Two ways to use a map, and the second one is why a history deck never got one.
   *
   * With numbers it is a choropleth: shade by value, legend carries the figures.
   * Without them it is a REFERENCE map — 「這幾個地方」 — and that is most of what
   * a slide actually wants a map for: where the trading ports were, which markets
   * we sell into, which provinces the pilot covers. The first version required
   * values, so a deck with regions but no statistics simply never drew one, and
   * from the outside that looked like the feature was missing.
   */
  const highlight = cats.length > 0 && !nums.some((v) => Number.isFinite(Number(v)));
  const values = new Map();
  cats.forEach((c, i) => {
    const v = Number(nums[i]);
    if (Number.isFinite(v)) values.set(c, v);
  });
  const marked = new Set(highlight ? cats : []);
  // Named but unknown is worth saying out loud: a typo in a region name would
  // otherwise just be a country that never lights up.
  // Deferred until the map is chosen — see below.

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
  /**
   * How much of the canvas the legend may take.
   *
   * ── Why a constant was wrong ──
   * 150 points is a quarter of a full-width chart and more than a third of a
   * half-width one — and a map on half a slide is already small. 「地圖有點小，
   * 反而圖例太大」 is what that ratio looks like: the list wins the page and the
   * geography, which is the only reason to use a map at all, loses it.
   *
   * Proportional, with a floor that still fits 「United Sta…  1.4k」 and a
   * ceiling so a very wide chart does not grow a very wide list.
   */
  const legendW = (values.size || marked.size)
    ? Math.max(104, Math.min(150, Math.round(f.w * 0.24)))
    : 0;
  // Type shrinks with the column, or a narrow legend clips every name.
  const legendSize = legendW < 128 ? 10.5 : 12;
  const rowH = legendW < 128 ? 19 : 22;
  // Room for the rows plus the two footnotes that may follow them.
  const maxRows = Math.max(3, Math.min(8, Math.floor((f.h - 44) / rowH)));
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
  const shade = (v, name) => {
    // A highlighted region gets the full accent: there is no scale to sit on,
    // so anything paler would read as 「a bit selected」, which is not a thing.
    if (marked.has(name)) return (spec.colors || [])[0] || theme.accent || SERIES[2];
    if (v === undefined) return empty;
    // A floor of 0.18 so the lowest region is still visibly part of the data
    // rather than indistinguishable from one with no value at all.
    const t = hi === lo ? 1 : 0.18 + 0.82 * ((v - lo) / (hi - lo));
    return mixHex(empty, strong, t);
  };

  for (const [name, rings] of Object.entries(geo)) {
    const fill = shade(values.get(name), name);
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
  if (values.size || marked.size) {
    const rows = marked.size
      // Highlight mode: the list is the names, in the order they were given —
      // that order is usually chronological or narrative, and re-sorting it
      // would throw away the only ordering the caller had.
      ? [...marked].slice(0, maxRows).map((n) => [n, null])
      : [...values.entries()].sort((a, b) => b[1] - a[1]).slice(0, maxRows);
    const lx = f.x + plotW + 16;
    let ly = f.y + 8;
    for (const [name, v] of rows) {
      out += rect(lx, ly + (rowH - 12) / 2 - 5, 11, 11, shade(v, name), ' rx="2"');
      /**
       * The name column has a fixed width and 「United States of America」 does
       * not fit in it — it ran straight under the number on its right. SVG has
       * no text-overflow, so the string is cut here.
       *
       * Measured rather than counted: a CJK glyph is about as wide as the type
       * is tall, latin is roughly half that, so 「广东」 and 「Bangladesh」 both
       * get as many characters as actually fit.
       */
      const room = (f.W - 16) - (lx + 17) - (v === null ? 4 : legendSize * 3.4);
      let shown = name;
      let used = 0;
      for (let i = 0; i < name.length; i++) {
        used += /[　-鿿＀-￯]/.test(name[i]) ? legendSize : legendSize * 0.55;
        if (used > room) { shown = `${name.slice(0, Math.max(1, i - 1))}…`; break; }
      }
      out += text(lx + 17, ly + 6, shown,
        { size: legendSize, color: theme.ink || '#3D3D3D', baseline: 'middle' });
      if (v !== null) {
        out += text(f.W - 14, ly + 6, short(v),
          { size: legendSize, bold: true, color: theme.strong || SERIES[0],
            anchor: 'end', baseline: 'middle' });
      }
      ly += rowH;
    }
    const total = marked.size || values.size;
    if (total > rows.length) {
      out += text(lx, ly + 6, `（另有 ${total - rows.length} 個未列出）`,
        { size: legendSize - 1.5, color: theme.muted || '#8A93A0', baseline: 'middle' });
      ly += rowH - 2;
    }
    if (swapped) {
      out += text(lx, ly + 6, `（已改用${which}地圖）`,
        { size: legendSize - 1.5, color: theme.muted || '#8A93A0', baseline: 'middle' });
      ly += rowH - 4;
    }
    if (unknown.length) {
      out += text(lx, ly + 6, `找不到：${unknown.slice(0, 3).join('、')}`,
        { size: legendSize - 1.5, color: theme.muted || '#8A93A0', baseline: 'middle' });
    }
  }
  return out;
}

/**
 * A scatter — do these two numbers move together?
 *
 * ── The argument nothing else here could make ──
 * Every other chart in this module compares magnitudes across categories. None
 * of them can say 「規模越大，成熟度越高」, which is a claim analysis decks make
 * constantly and currently have to make in prose. A scatter says it in one
 * glance, and says it honestly: if the cloud has no shape, there is no
 * relationship, and the chart admits that where a sentence would not.
 *
 * Points may be labelled. Labels are drawn only when there are few enough for
 * them not to collide — past a dozen the names are noise and the shape is the
 * point.
 */
function scatterChart(spec, theme) {
  const pts = (Array.isArray(spec.points) ? spec.points : []).filter(Boolean)
    .map((p) => ({
      x: Number(p.x), y: Number(p.y), label: String(p.label || ''),
      group: String(p.group || ''),
    }))
    .filter((p) => Number.isFinite(p.x) && Number.isFinite(p.y));
  if (!pts.length) throw new Error('points 是必要的：[{ x, y, label }]。');

  const f = frame(spec, { axis: true, legend: false });
  let out = chrome(spec, f, theme);

  const xs = pts.map((p) => p.x);
  const ys = pts.map((p) => p.y);
  // Axes start at zero unless the data sits far from it — a scatter of values
  // between 80 and 90 plotted from zero is a single blob in one corner.
  /**
   * The axis range for one dimension.
   *
   * Zero-based when the data is anywhere near zero, because a scatter that
   * floats above a cut axis exaggerates every difference. Hugged when it is
   * not: six values between 80 and 85 plotted from zero are a single dot.
   *
   * The TOP has to follow the same decision as the bottom. Rounding it up to a
   * nice number while the bottom hugs the data puts the gap straight back —
   * 79.4 to 100 for data that only spans 80 to 85.
   */
  const range = (vals) => {
    const lo = Math.min(...vals);
    const hi = Math.max(...vals);
    const spread = (hi - lo) || Math.abs(hi) || 1;
    if (lo >= 0 && lo < spread * 0.6) return [0, niceMax(hi)];
    return [lo - spread * 0.12, hi + spread * 0.12];
  };
  const [x0, x1] = range(xs);
  const [y0, y1] = range(ys);
  const px = (v) => f.x + ((v - x0) / ((x1 - x0) || 1)) * f.w;
  const py = (v) => f.y + f.h - ((v - y0) / ((y1 - y0) || 1)) * f.h;

  for (let i = 0; i <= 4; i++) {
    const gy = f.y + (f.h / 4) * i;
    out += line(f.x, gy, f.x + f.w, gy, theme.line);
    out += text(f.x - 8, gy, short(y1 - ((y1 - y0) / 4) * i),
      { size: 11, color: theme.muted, anchor: 'end', baseline: 'middle' });
    const gx = f.x + (f.w / 4) * i;
    out += text(gx, f.y + f.h + 18, short(x0 + ((x1 - x0) / 4) * i),
      { size: 11, color: theme.muted, anchor: 'middle' });
  }

  const groups = [...new Set(pts.map((p) => p.group))];
  const tone = (g) => (spec.colors || [])[Math.max(0, groups.indexOf(g))]
    || SERIES[Math.max(0, groups.indexOf(g)) % SERIES.length];
  const showLabels = spec.labels !== false && pts.length <= 12;

  for (const p of pts) {
    out += `<circle cx="${n2(px(p.x))}" cy="${n2(py(p.y))}" r="6"`
      + ` fill="${tone(p.group)}" fill-opacity="0.75"`
      + ` stroke="${theme.bg || '#FFFFFF'}" stroke-width="1.5"/>`;
    if (showLabels && p.label) {
      out += text(px(p.x), py(p.y) - 12, p.label,
        { size: 11, color: theme.ink, anchor: 'middle' });
    }
  }

  // The axis names matter more here than on any other chart: a scatter without
  // them is two unlabelled numbers and nobody can read the claim off it.
  if (spec.x_axis) {
    out += text(f.x + f.w, f.y + f.h + 32, String(spec.x_axis),
      { size: 11.5, bold: true, color: theme.strong, anchor: 'end' });
  }
  if (spec.y_axis) {
    out += text(f.x - 44, f.y - 6, String(spec.y_axis),
      { size: 11.5, bold: true, color: theme.strong });
  }
  if (groups.length > 1) {
    let lx = f.x;
    for (const g of groups) {
      out += rect(lx, f.H - 22, 11, 11, tone(g), ' rx="2"');
      out += text(lx + 17, f.H - 16, g || '（未分組）',
        { size: 11.5, color: theme.ink, baseline: 'middle' });
      lx += 30 + String(g).length * 12;
    }
  }
  return out;
}

/**
 * Stacked bars — the total changed, and so did what it is made of.
 *
 * ── Why grouped bars cannot answer this ──
 * Side-by-side bars show each series' own trend and hide the total; a reader
 * asked 「整體有沒有成長」 has to add three bars in their head for every
 * category. Stacking puts the total on the page as the height of the column and
 * the mix as the divisions inside it, which is exactly the question revenue
 * breakdowns and market-share slides are asking.
 *
 * `stacked100` normalises every column to full height, which drops the total
 * entirely and shows only the mix. That is the right chart when the totals are
 * not comparable — and the wrong one when they are, because it hides growth.
 */
function stackedChart(spec, theme) {
  const { labels, series } = normalise(spec);
  const pct = spec.type === 'stacked100';
  const f = frame(spec, { axis: true, legend: wantsLegend(series) });

  const totals = labels.map((_, i) => series.reduce((a, sr) => a + Math.max(0, sr.data[i]), 0));
  const top = pct ? 100 : niceMax(Math.max(...totals, 0));
  let out = chrome(spec, f, theme);

  for (let i = 0; i <= 4; i++) {
    const y = f.y + f.h - (f.h / 4) * i;
    out += line(f.x, y, f.x + f.w, y, theme.line);
    out += text(f.x - 8, y, pct ? `${i * 25}%` : short((top / 4) * i),
      { size: 11, color: theme.muted, anchor: 'end', baseline: 'middle' });
  }

  const band = f.w / labels.length;
  const bw = Math.min(72, band * 0.56);
  labels.forEach((label, i) => {
    const cx = f.x + band * i + band / 2;
    const total = totals[i] || 1;
    let y = f.y + f.h;
    series.forEach((sr) => {
      const v = Math.max(0, sr.data[i]);
      const h = (pct ? (v / total) * 100 : v) / (top || 1) * f.h;
      if (h <= 0) return;
      y -= h;
      out += rect(cx - bw / 2, y, bw, h, sr.color);
      // Only label a segment tall enough to hold the number. A 9pt slice with a
      // number jammed into it is less readable than one without.
      if (spec.data_labels !== false && h > 20) {
        out += text(cx, y + h / 2, pct ? `${Math.round((v / total) * 100)}%` : short(v),
          { size: 11, bold: true, color: '#FFFFFF', anchor: 'middle', baseline: 'middle' });
      }
    });
    out += text(cx, f.y + f.h + 18, label,
      { size: 11.5, color: theme.ink, anchor: 'middle' });
  });

  if (series.length > 1) {
    let lx = f.x;
    for (const sr of series) {
      out += rect(lx, f.H - 22, 11, 11, sr.color, ' rx="2"');
      out += text(lx + 17, f.H - 16, sr.name,
        { size: 11.5, color: theme.ink, baseline: 'middle' });
      lx += 30 + String(sr.name).length * 12;
    }
  }
  return out;
}

/**
 * A box plot — how spread out each group is, not just where its middle sits.
 *
 * ── The question only this answers ──
 * Two factories with the same average lead time can be completely different
 * businesses: one delivers in 5–7 days every time, the other averages the same
 * number by mixing 2-day rushes with 20-day disasters. A column chart of means
 * shows those as identical bars. The box is the variation, and variation is
 * usually the thing the meeting is actually about.
 *
 * ── Stated plainly, because it matters for a slide ──
 * Most audiences cannot read a box plot. It is the right chart when the room is
 * technical or the spread IS the finding; it is the wrong one when somebody
 * just wants to compare four numbers, and the tool description says so. Given a
 * choice between a chart that is correct and a chart that is understood, a
 * presentation wants the second.
 *
 * Takes raw observations per group and computes the summary, because raw
 * numbers are what people have. An explicit five-number summary is accepted too
 * for data that arrived pre-aggregated.
 */
function boxChart(spec, theme) {
  const quartile = (sorted, q) => {
    const pos = (sorted.length - 1) * q;
    const lo = Math.floor(pos);
    const hi = Math.ceil(pos);
    return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
  };

  /** One group's five-number summary, plus whatever fell outside the whiskers. */
  const summarise = (raw, label) => {
    if (raw && typeof raw === 'object' && !Array.isArray(raw) && raw.median !== undefined) {
      return {
        label: String(raw.label || label || ''),
        min: Number(raw.min), q1: Number(raw.q1), med: Number(raw.median),
        q3: Number(raw.q3), max: Number(raw.max),
        out: (Array.isArray(raw.outliers) ? raw.outliers : []).map(Number).filter(Number.isFinite),
      };
    }
    const nums = (Array.isArray(raw) ? raw : []).map(Number).filter(Number.isFinite)
      .sort((a, b) => a - b);
    if (!nums.length) return null;
    const q1 = quartile(nums, 0.25);
    const q3 = quartile(nums, 0.75);
    // Tukey's 1.5×IQR: the convention, so a reader who knows box plots reads
    // ours the way they read every other one.
    const fence = (q3 - q1) * 1.5;
    const inside = nums.filter((v) => v >= q1 - fence && v <= q3 + fence);
    return {
      label: String(label || ''),
      min: inside.length ? inside[0] : nums[0],
      q1,
      med: quartile(nums, 0.5),
      q3,
      max: inside.length ? inside[inside.length - 1] : nums[nums.length - 1],
      out: nums.filter((v) => v < q1 - fence || v > q3 + fence),
    };
  };

  const cats = (Array.isArray(spec.categories) ? spec.categories : []).map(String);
  const source = Array.isArray(spec.boxes) && spec.boxes.length ? spec.boxes
    : (Array.isArray(spec.values) && Array.isArray(spec.values[0]) ? spec.values : null);
  if (!source) {
    throw new Error('箱型圖要每一組的原始數值：values: [[...], [...]]，'
      + '或 boxes: [{ label, min, q1, median, q3, max }]。');
  }
  const boxes = source.map((raw, i) => summarise(raw, cats[i] || `第 ${i + 1} 組`)).filter(Boolean);
  if (!boxes.length) throw new Error('箱型圖沒有拿到任何有效數值。');

  const f = frame(spec, { axis: true, legend: false });
  let out = chrome(spec, f, theme);

  const all = boxes.flatMap((b) => [b.min, b.max, ...b.out]);
  const lo = Math.min(...all);
  const hi = Math.max(...all);
  // Padded rather than zero-based: a distribution between 4.8 and 5.2 plotted
  // from zero is four identical lines.
  const pad = (hi - lo) * 0.12 || 1;
  const y0 = lo - pad;
  const y1 = hi + pad;
  const py = (v) => f.y + f.h - ((v - y0) / ((y1 - y0) || 1)) * f.h;

  for (let i = 0; i <= 4; i++) {
    const gy = f.y + (f.h / 4) * i;
    out += line(f.x, gy, f.x + f.w, gy, theme.line);
    out += text(f.x - 8, gy, short(y1 - ((y1 - y0) / 4) * i),
      { size: 11, color: theme.muted, anchor: 'end', baseline: 'middle' });
  }

  const band = f.w / boxes.length;
  const bw = Math.min(56, band * 0.46);
  boxes.forEach((b, i) => {
    const cx = f.x + band * i + band / 2;
    const tone = (spec.colors || [])[i] || theme.strong || SERIES[0];

    // Whisker, then the box on top of it, so the line does not show through a
    // translucent fill.
    out += line(cx, py(b.max), cx, py(b.min), tone, 1.4);
    out += line(cx - bw / 4, py(b.max), cx + bw / 4, py(b.max), tone, 1.4);
    out += line(cx - bw / 4, py(b.min), cx + bw / 4, py(b.min), tone, 1.4);

    const top = py(b.q3);
    const h = Math.max(2, py(b.q1) - top);
    out += rect(cx - bw / 2, top, bw, h, theme.bg || '#FFFFFF');
    out += `<rect x="${n2(cx - bw / 2)}" y="${n2(top)}" width="${n2(bw)}" height="${n2(h)}"`
      + ` fill="${tone}" fill-opacity="0.22" stroke="${tone}" stroke-width="1.4"/>`;
    // The median is the one number people read off a box, so it is the heaviest
    // mark on it.
    out += line(cx - bw / 2, py(b.med), cx + bw / 2, py(b.med), tone, 2.6);

    for (const v of b.out) {
      out += `<circle cx="${n2(cx)}" cy="${n2(py(v))}" r="3" fill="none"`
        + ` stroke="${theme.accent || SERIES[2]}" stroke-width="1.4"/>`;
    }
    out += text(cx, f.y + f.h + 18, b.label,
      { size: 11.5, color: theme.ink, anchor: 'middle' });
    if (spec.data_labels !== false) {
      out += text(cx + bw / 2 + 6, py(b.med), short(b.med),
        { size: 10.5, bold: true, color: tone, baseline: 'middle' });
    }
  });
  return out;
}

/**
 * A Gantt chart — what happens when, drawn rather than described.
 *
 * ── Why this exists instead of mermaid's `gantt` ──
 * 「我不建議用 mermaid 來呈現」, with a screenshot that made the case: mermaid
 * writes each task's NAME ON ITS BAR, so on any real schedule the labels run
 * into the bars beside them and the chart becomes a pile of overlapping text.
 * Underneath it, every date on the axis was printed whether or not there was
 * room, which turned the bottom of the figure into a smear of digits.
 *
 * Both failures are the same mistake — drawing everything the data contains and
 * hoping it fits. So this one makes two decisions up front, and they are the
 * whole design:
 *
 *   1. **Names live in a left-hand gutter, never on the bars.** A column of
 *      left-aligned labels is readable at any row count; a label floating on a
 *      bar is readable only when the bar happens to be long enough.
 *   2. **The number of date ticks is derived from the width available**, not
 *      from the number of days. Fewer, legible dates beat all of them.
 *
 * Sections (「MVP (P0)」, 「V1.1 (P1)」) become tinted bands behind their rows,
 * which is what makes a phased plan readable as phases rather than as a list.
 */
const DAY_MS = 86400000;

/** `2026-10-11` → a day number. Returns NaN for anything else, on purpose. */
function dayOf(value) {
  const m = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(String(value || '').trim());
  if (!m) return NaN;
  return Date.UTC(Number(m[1]), Number(m[2]) - 1, Number(m[3])) / DAY_MS;
}

/** A day number back to `10/11`, or `2026/10/11` when the year matters. */
function dayLabel(day, withYear) {
  const d = new Date(day * DAY_MS);
  const md = `${d.getUTCMonth() + 1}/${d.getUTCDate()}`;
  return withYear ? `${d.getUTCFullYear()}/${md}` : md;
}

/**
 * How many date labels fit, and how far apart they should sit.
 *
 * Measured against the room on the axis rather than the length of the project:
 * the same six-month plan drawn 900 points wide and 400 points wide needs a
 * different number of dates, and printing the same number on both is how the
 * narrow one ends up illegible.
 */
function tickStep(days, plotWidth, withYear) {
  const labelPt = withYear ? 62 : 42;
  const room = Math.max(2, Math.floor(plotWidth / labelPt));
  // Steps a person reads as intervals: a week, a fortnight, a month, a quarter.
  for (const step of [1, 2, 3, 7, 14, 30, 60, 90, 180, 365]) {
    if (days / step <= room) return step;
  }
  return Math.ceil(days / room);
}

function ganttChart(spec, theme) {
  const raw = Array.isArray(spec.tasks) ? spec.tasks : [];
  const tasks = raw.map((t) => {
    const start = dayOf(t && t.start);
    // `end` wins; `days` is the convenience form for 「這件事做五天」.
    const len = Math.max(1, Number(t && t.days) || 0);
    const end = t && t.end ? dayOf(t.end) : start + len - 1;
    return {
      name: String((t && t.name) || ''),
      group: String((t && t.group) || ''),
      start,
      end,
      done: Math.max(0, Math.min(1, Number(t && t.progress) || 0)),
    };
  }).filter((t) => t.name && Number.isFinite(t.start) && Number.isFinite(t.end) && t.end >= t.start);

  if (!tasks.length) {
    throw new Error('tasks 是必要的：[{ name, start: "2026-10-11", end: "2026-10-20" }]。'
      + '日期要用 YYYY-MM-DD；也可以用 days 代替 end。');
  }

  const W = Number(spec.width) || 760;
  const H = Number(spec.height) || Math.max(180, 64 + tasks.length * 26 + 34);
  const first = Math.min(...tasks.map((t) => t.start));
  const last = Math.max(...tasks.map((t) => t.end));
  const span = Math.max(1, last - first + 1);
  /**
   * Show the year when it is ambiguous, not just when the plan is long.
   *
   * A five-month plan running Oct→Mar is under 300 days, so a span-only test
   * prints 「12/22」 and 「1/5」 side by side with nothing saying they are twelve
   * months apart in reading order but two weeks apart in fact.
   */
  const withYear = span > 300
    || new Date(first * DAY_MS).getUTCFullYear() !== new Date(last * DAY_MS).getUTCFullYear();

  /**
   * The gutter holds the section label AND the task name, side by side.
   *
   * Measuring it from the names alone was the first version's bug: the section
   * column was then drawn over the first 56 points of that space, and every
   * task name lost its last few characters to make room — 「需求確認與設計」 came
   * out as 「需求確…」, which reads as data someone truncated rather than as a
   * layout that ran out of room.
   */
  const grouped = tasks.some((t) => t.group);
  const groupW = grouped
    ? Math.min(96, Math.max(44, Math.max(...tasks.map((t) => labelWidth(t.group))) + 10))
    : 0;
  const nameW = Math.min(
    groupW + Math.max(70, Math.max(...tasks.map((t) => labelWidth(t.name))) + 14),
    Math.round(W * 0.46),
  );
  const top = spec.title ? 44 : 16;
  const axisH = 26;
  const x0 = 12 + nameW;
  const plotW = Math.max(60, W - x0 - 16);
  const rows = tasks.length;
  const rowH = Math.max(16, Math.min(30, (H - top - axisH - 8) / rows));
  const barH = Math.max(8, rowH - 9);
  const atDay = (d) => x0 + ((d - first) / span) * plotW;

  let out = '';
  if (spec.title) out += text(12, 26, spec.title, { size: 17, bold: true, color: theme.strong });

  /**
   * Section bands first, so everything else sits on top of them.
   *
   * Drawn as one band per CONTIGUOUS run rather than per distinct name: tasks
   * are listed in the order they happen, so a group that appears twice is two
   * phases with the same label and deserves two bands.
   */
  const bands = [];
  for (let i = 0; i < rows; i++) {
    const g = tasks[i].group;
    const prev = bands[bands.length - 1];
    if (prev && prev.group === g) prev.to = i;
    else bands.push({ group: g, from: i, to: i });
  }
  const tintFor = (i) => mixHex(theme.bg || '#FFFFFF',
    (spec.colors || [])[i % ((spec.colors || []).length || 1)]
      || SERIES[i % SERIES.length], 0.12);
  bands.forEach((b, i) => {
    if (!b.group) return;
    const y = top + b.from * rowH;
    out += rect(0, y, W, (b.to - b.from + 1) * rowH, tintFor(i));
    out += text(10, y + ((b.to - b.from + 1) * rowH) / 2, b.group,
      { size: 11.5, bold: true, color: theme.muted || '#8A93A0', baseline: 'middle' });
  });

  // Grid and dates. One pass, so a tick and its label can never disagree.
  const step = tickStep(span, plotW, withYear);
  const axisY = top + rows * rowH;
  for (let d = first; d <= last; d += step) {
    out += line(atDay(d), top, atDay(d), axisY, theme.line || '#D5DBE0');
    out += text(atDay(d), axisY + 16, dayLabel(d, withYear),
      { size: 10.5, color: theme.muted || '#8A93A0', anchor: 'middle' });
  }
  out += line(x0, axisY, x0 + plotW, axisY, theme.line || '#D5DBE0');

  tasks.forEach((t, i) => {
    const y = top + i * rowH;
    const tone = (spec.colors || [])[bands.findIndex((b) => i >= b.from && i <= b.to) % ((spec.colors || []).length || 1)]
      || SERIES[Math.max(0, bands.findIndex((b) => i >= b.from && i <= b.to)) % SERIES.length];
    // The name, in the gutter. Left-aligned and clipped by the gutter width
    // rather than by the bar, so a long name never lands on top of a neighbour.
    out += text(12 + groupW, y + rowH / 2, fit(t.name, nameW - groupW - 6),
      { size: 11.5, color: theme.ink || '#3D3D3D', baseline: 'middle' });

    const bx = atDay(t.start);
    // +1 because a task that starts and ends on the same day still lasts a day;
    // a floor so a one-day task is a mark rather than a hairline.
    const bw = Math.max(4, ((t.end - t.start + 1) / span) * plotW);
    out += rect(bx, y + (rowH - barH) / 2, bw, barH, tone, ' rx="3"');
    if (t.done > 0) {
      out += rect(bx, y + (rowH - barH) / 2, bw * t.done, barH,
        mixHex(tone, '#000000', 0.25), ' rx="3"');
    }
  });

  /**
   * Today, when today is inside the plan.
   *
   * Drawn last so it reads as an overlay, and skipped entirely when it falls
   * outside the range — a red line pinned to the left edge of a chart about next
   * year says nothing and looks like a rendering fault.
   */
  const now = dayOf(spec.today);
  if (Number.isFinite(now) && now >= first && now <= last) {
    out += line(atDay(now), top - 4, atDay(now), axisY, theme.accent || '#C0392B', 1.5);
    out += text(atDay(now), top - 8, '今天',
      { size: 10, color: theme.accent || '#C0392B', anchor: 'middle' });
  }

  return out;
}

/**
 * How tall this chart wants to be when nobody said.
 *
 * A Gantt's height is a function of its row count — the other charts all fit the
 * same box whatever they hold, a Gantt does not. Twenty tasks in a 380-point
 * frame is six points a row, which is a smear.
 */
ganttChart.preferredHeight = (spec) => {
  const n = (Array.isArray(spec.tasks) ? spec.tasks : []).length || 1;
  return Math.max(180, (spec.title ? 44 : 16) + n * 26 + 34);
};

/** Rough on-screen width of a label, CJK counted double. */
function labelWidth(s) {
  let w = 0;
  for (const ch of String(s || '')) {
    w += /[\u2E80-\uA4CF\uAC00-\uD7A3\uF900-\uFAFF\uFE10-\uFE6F\uFF00-\uFF60]/.test(ch) ? 11.5 : 6;
  }
  return w;
}

/** Trim a label to the room it has, with an ellipsis so the cut is visible. */
function fit(s, room) {
  const str = String(s || '');
  if (labelWidth(str) <= room) return str;
  let out = '';
  for (const ch of str) {
    if (labelWidth(out + ch) > room - 8) break;
    out += ch;
  }
  return out + '…';
}

/**
 * A function plot — y against x, on two numeric axes.
 *
 * ── Why the other charts could not do this ──
 * 「可以加一些數學的圖嗎？例如線性圖之類的」, asked while writing a report about
 * the Fourier transform. Every chart in this module until now has a CATEGORY
 * axis: the x positions are names laid out evenly, which is right for quarters
 * and products and wrong for mathematics. You cannot draw N² against N log N on
 * an evenly spaced axis — the whole point of that figure is that the gap grows,
 * and even spacing is exactly what hides it.
 *
 * So this one has a real x scale. Curves are given as points and drawn in the
 * order supplied, which also makes it the right chart for a measured series, a
 * fitted line, or anything where x is a quantity rather than a label.
 *
 * ── Log scales, because the interesting ones need them ──
 * Algorithmic complexity, frequency response, signal power: the figures a
 * technical document wants are usually several orders of magnitude tall, and on
 * a linear axis they are a flat line with one spike. `log_y` (and `log_x`) turn
 * those into the straight lines the reader is looking for.
 */
function plotChart(spec, theme) {
  const raw = Array.isArray(spec.curves) ? spec.curves : [];
  const curves = raw.map((c, i) => ({
    name: String((c && c.name) || ''),
    color: (spec.colors || [])[i] || SERIES[i % SERIES.length],
    dashed: !!(c && c.dashed),
    dots: (c && c.dots) === true,
    points: (Array.isArray(c && c.points) ? c.points : [])
      .map((pt) => (Array.isArray(pt) ? { x: Number(pt[0]), y: Number(pt[1]) } : { x: Number(pt && pt.x), y: Number(pt && pt.y) }))
      .filter((pt) => Number.isFinite(pt.x) && Number.isFinite(pt.y)),
  })).filter((c) => c.points.length >= 2);

  if (!curves.length) {
    throw new Error('curves 是必要的：[{ name: "N log N", points: [[256, 2048], [1024, 10240]] }]。'
      + '每條線至少要兩個點，points 可以是 [x, y] 或 { x, y }。');
  }

  const logX = spec.log_x === true;
  const logY = spec.log_y === true;
  const all = curves.flatMap((c) => c.points);
  // A log axis cannot show zero or negative values, so the data decides whether
  // the request is possible rather than the renderer drawing -Infinity.
  if (logX && all.some((p) => p.x <= 0)) throw new Error('log_x 需要所有 x 都大於 0。');
  if (logY && all.some((p) => p.y <= 0)) throw new Error('log_y 需要所有 y 都大於 0。');

  const tx = (v) => (logX ? Math.log10(v) : v);
  const ty = (v) => (logY ? Math.log10(v) : v);
  const xs = all.map((p) => tx(p.x));
  const ys = all.map((p) => ty(p.y));
  const pad = (lo, hi) => {
    if (hi === lo) return [lo - 1, hi + 1];
    const room = (hi - lo) * 0.06;
    return [lo - room, hi + room];
  };
  // A linear y axis starts at zero when the data is near it — a bar-like reading
  // of a curve that never approaches zero is misleading in the other direction,
  // so the same 「is zero in sight」 test the scatter uses applies here.
  const rawMinY = Math.min(...ys);
  const rawMaxY = Math.max(...ys);
  const [y0, y1] = logY
    ? pad(rawMinY, rawMaxY)
    : pad(rawMinY >= 0 && rawMinY < (rawMaxY - rawMinY) * 0.6 ? 0 : rawMinY, rawMaxY);
  const [x0, x1] = pad(Math.min(...xs), Math.max(...xs));

  const f = frame(spec, { axis: true, legend: wantsLegend(curves) });
  let out = chrome(spec, f, theme);
  const px = (v) => f.x + ((tx(v) - x0) / ((x1 - x0) || 1)) * f.w;
  const py = (v) => f.y + f.h - ((ty(v) - y0) / ((y1 - y0) || 1)) * f.h;

  /** A tick's label: the real value, not the logarithm of it. */
  const back = (t, isLog) => (isLog ? 10 ** t : t);
  for (let i = 0; i <= 4; i++) {
    const gy = f.y + (f.h / 4) * i;
    out += line(f.x, gy, f.x + f.w, gy, theme.line);
    out += text(f.x - 8, gy, short(back(y1 - ((y1 - y0) / 4) * i, logY)),
      { size: 11, color: theme.muted, anchor: 'end', baseline: 'middle' });
    const gx = f.x + (f.w / 4) * i;
    out += line(gx, f.y, gx, f.y + f.h, theme.line);
    out += text(gx, f.y + f.h + 18, short(back(x0 + ((x1 - x0) / 4) * i, logX)),
      { size: 11, color: theme.muted, anchor: 'middle' });
  }

  for (const c of curves) {
    const d = c.points
      .map((pt, i) => `${i ? 'L' : 'M'}${n2(px(pt.x))},${n2(py(pt.y))}`)
      .join('');
    out += `<path d="${d}" fill="none" stroke="${c.color}" stroke-width="2.2"`
      + ' stroke-linejoin="round" stroke-linecap="round"'
      + (c.dashed ? ' stroke-dasharray="6 4"' : '') + '/>';
    // Markers only when asked. On a smooth curve they are clutter; on a handful
    // of measured points they are the data and the line is the interpolation.
    if (c.dots) {
      for (const pt of c.points) {
        out += `<circle cx="${n2(px(pt.x))}" cy="${n2(py(pt.y))}" r="3.2" fill="${c.color}"`
          + ` stroke="${theme.bg || '#FFFFFF'}" stroke-width="1.2"/>`;
      }
    }
  }

  /**
   * Axis names matter more here than on a category chart.
   *
   * 「N」 and 「運算量」 are the difference between a figure and a squiggle — a
   * category axis labels itself, a numeric one does not.
   */
  if (spec.x_axis) {
    out += text(f.x + f.w, f.y + f.h + 34, String(spec.x_axis)
      + (logX ? '（對數刻度）' : ''),
    { size: 11.5, bold: true, color: theme.strong, anchor: 'end' });
  }
  if (spec.y_axis) {
    out += text(f.x - 44, f.y - 8, String(spec.y_axis) + (logY ? '（對數刻度）' : ''),
      { size: 11.5, bold: true, color: theme.strong });
  }
  out += legend(curves, f, theme);
  return out;
}

const RENDER = {
  column: barChart, bar: barChart, line: lineChart, area: lineChart,
  pie: pieChart, donut: pieChart,
  waterfall: waterfallChart, funnel: funnelChart, radar: radarChart,
  timeline: timelineChart, map: mapChart,
  scatter: scatterChart, stacked: stackedChart, stacked100: stackedChart,
  gantt: ganttChart, plot: plotChart,
  box: boxChart,
};
export const CHART_TYPES = Object.keys(RENDER);

/**
 * Chart spec → a complete SVG document.
 *
 * Returns the markup as a string. The caller hands it to
 * `setSelectedDataAsync` with `Office.CoercionType.XmlSvg`, which is what puts
 * it on the slide as a vector object.
 */
/**
 * Every field a renderer reads off the spec.
 *
 * ── Why this list has to exist ──
 * Each host built its own object to hand to `chartSvg`, naming the fields one by
 * one. Every new chart type therefore needed the same field added in three
 * places, and the place that got forgotten was always a host — so the tool
 * advertised `type: "plot"` in its enum, described `curves` in its own
 * documentation, and then dropped `curves` on the floor on the way here. The
 * renderer saw a spec with no data and said 「curves 是必要的」 about an argument
 * the model had in fact supplied. `timeline` shipped broken exactly this way
 * once already, and `plot`, `box`, `map` and `scatter` were broken this way
 * again on both hosts.
 *
 * A host now copies these keys wholesale. Adding a chart type means adding its
 * field here, in the module that reads it, and the hosts follow for free.
 *
 * `width`, `height`, `type` and `title` are deliberately absent: a host decides
 * those itself from the slide or the text column, rather than passing them on.
 */
export const CHART_DATA_KEYS = [
  'categories', 'values', 'series', 'colors', 'labels', 'data_labels',
  'events', 'total', 'points', 'boxes', 'map', 'area',
  'curves', 'x_axis', 'y_axis', 'log_x', 'log_y',
  'center', 'center_label', 'tasks', 'today',
];

/** Pick the data fields out of a tool's arguments, leaving its own options behind. */
export function chartData(args = {}) {
  const out = {};
  for (const k of CHART_DATA_KEYS) {
    if (args[k] !== undefined) out[k] = args[k];
  }
  return out;
}

export function chartSvg(spec = {}, theme = {}) {
  const type = String(spec.type || 'column').toLowerCase();
  const render = RENDER[type];
  if (!render) throw new Error(`不支援的圖表類型：${type}。可用的有 ${CHART_TYPES.join('、')}。`);

  /**
   * The palette the renderers actually get.
   *
   * ── Why this list grew ──
   * It used to carry five roles, which was everything the bar and line charts
   * needed. The newer ones need more: a waterfall colours its falls with the
   * accent, a timeline marks its last stop with it, a map shades unset regions
   * with the tint. Those roles were not in the object, so each of them silently
   * fell back to this module's own defaults — producing a chart in navy and red
   * on a teal deck, with nothing anywhere reporting a problem.
   *
   * Defaults are still filled in per role, so a caller that passes a partial
   * palette gets a coherent chart rather than an undefined colour.
   */
  const t = {
    ink: theme.ink || '#3D3D3D',
    strong: theme.strong || '#1A3A5C',
    second: theme.second || '#2E6DA4',
    accent: theme.accent || '#C0392B',
    muted: theme.muted || '#8A93A0',
    line: theme.line || '#D5DBE0',
    tint: theme.tint || '#EEF3F8',
    ground: theme.ground || '#F7F9FB',
    bg: theme.bg || '#FFFFFF',
  };
  const W = Number(spec.width) || 760;
  // A renderer may ask for a height when the caller did not give one — see
  // ganttChart.preferredHeight, where the row count decides it.
  const H = Number(spec.height)
    || (typeof render.preferredHeight === 'function' ? render.preferredHeight(spec) : 380);
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
