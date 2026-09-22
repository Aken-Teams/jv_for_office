/**
 * mermaid — turn diagram source into a picture the document can hold.
 *
 * ── Why the library is vendored ──
 * The pane had no runtime dependencies and no build step, and that was worth
 * something: edit a file, refresh, done. This breaks that, and it is worth being
 * explicit about why rather than letting it happen quietly.
 *
 * Word has no diagram API. Neither does the server — the platform's own mermaid
 * support renders in a BROWSER (generate-slides emits `<pre class="mermaid">`
 * for a page to run), so there is no image pipeline to reuse. That leaves three
 * options: render here, stand up a headless browser on the server, or draw a
 * handful of diagram types by hand on a canvas. The pane IS a browser, so
 * rendering here costs no round trip and no deployment; the other two cost a
 * server component or most of what mermaid is for.
 *
 * ── Why the UMD build and not the ESM one ──
 * `mermaid.esm.min.mjs` is 26 KB because it code-splits: it fetches a chunk per
 * diagram type at render time. Vendoring it would mean vendoring the chunk
 * directory and hoping the paths resolve. `mermaid.min.js` is one 2.9 MB file
 * with everything in it, which is larger on disk and simpler in every other way.
 *
 * ── Why it is loaded lazily ──
 * 2.9 MB parsed at startup would be paid by every person who opens the pane,
 * and almost nobody asks for a diagram. The script tag is injected the first
 * time one is requested and the module is reused afterwards.
 */

let loading = null;

/**
 * Get mermaid, loading it once.
 *
 * A UMD bundle assigns to a global rather than resolving a module, so this is a
 * script tag and a promise rather than an import.
 */
function loadMermaid() {
  if (window.mermaid) return Promise.resolve(window.mermaid);
  if (loading) return loading;

  loading = new Promise((resolve, reject) => {
    const el = document.createElement('script');
    // Relative to taskpane.html, and same-origin: no CDN, so this works on a
    // machine with no internet and needs nothing from the CSP.
    el.src = './vendor/mermaid.min.js';
    el.onload = () => (window.mermaid
      ? resolve(window.mermaid)
      : reject(new Error('mermaid 載入了但沒有註冊全域物件')));
    el.onerror = () => {
      loading = null;
      reject(new Error('載不到 mermaid.min.js。確認 src/taskpane/vendor/ 裡有這個檔案。'));
    };
    document.head.appendChild(el);
  });
  return loading;
}

/** Rendered at 2x so Word does not scale a blurry bitmap up to the text column. */
const SCALE = 2;
/** Beyond this the PNG is bigger than the request body allows, and unreadable anyway. */
const MAX_EDGE = 2400;

/**
 * Diagram source → PNG, as base64 without the data: prefix.
 *
 * ── The one setting that matters ──
 * `htmlLabels: false`. By default mermaid puts label text inside
 * `<foreignObject>`, which is HTML embedded in the SVG — and a canvas cannot
 * rasterise that: the picture comes out with every box empty. Turning it off
 * makes mermaid emit real SVG `<text>`, which rasterises. It is the difference
 * between a diagram and a set of blank rectangles.
 */
export async function renderMermaid(code, opts = {}) {
  const mermaid = await loadMermaid();
  const font = opts.font || '"Microsoft JhengHei UI", "Microsoft JhengHei", sans-serif';

  mermaid.initialize({
    startOnLoad: false,
    // Diagram source can arrive from a document the person did not write, so it
    // is treated the way every other outside input here is.
    securityLevel: 'strict',
    htmlLabels: false,
    flowchart: { htmlLabels: false, useMaxWidth: false },
    sequence: { useMaxWidth: false },
    gantt: { useMaxWidth: false },
    class: { htmlLabels: false, useMaxWidth: false },
    /**
     * Recolouring mermaid, and why the old call never did.
     *
     * `theme` here is MERMAID's theme name — one of default / base / dark /
     * forest / neutral. The caller was handing it the deck's palette name
     * (`corporate`, `linen`, …), which mermaid does not recognise, so every
     * diagram silently fell back to the default lavender boxes. A flowchart in
     * purple on a teal deck reads as clip-art someone pasted in.
     *
     * The supported way to recolour is `base` plus themeVariables, so that is
     * what a palette turns into. Without one, the old behaviour stands.
     */
    theme: opts.palette ? 'base' : (opts.theme || 'default'),
    themeVariables: opts.palette ? {
      background: opts.palette.bg || '#FFFFFF',
      primaryColor: opts.palette.tint || '#EEF3F8',
      primaryTextColor: opts.palette.ink || '#3D3D3D',
      primaryBorderColor: opts.palette.strong || '#1A3A5C',
      secondaryColor: opts.palette.bg || '#FFFFFF',
      tertiaryColor: opts.palette.ground || '#F7F9FB',
      // Edges and their labels: the arrows are the argument in a flowchart, so
      // they get the deck's own ink rather than mermaid's grey.
      lineColor: opts.palette.strong || '#1A3A5C',
      textColor: opts.palette.ink || '#3D3D3D',
      mainBkg: opts.palette.tint || '#EEF3F8',
      nodeBorder: opts.palette.strong || '#1A3A5C',
      clusterBkg: opts.palette.ground || '#F7F9FB',
      clusterBorder: opts.palette.line || '#D5DBE0',
      titleColor: opts.palette.strong || '#1A3A5C',
      edgeLabelBackground: opts.palette.bg || '#FFFFFF',
      fontFamily: font,
      fontSize: '15px',
    } : undefined,
    fontFamily: font,
  });

  // A unique id per render: mermaid keys internal state off it, and reusing one
  // makes the second diagram in a document come back as the first.
  const id = 'jv-mmd-' + Math.random().toString(36).slice(2);
  let svg;
  try {
    ({ svg } = await mermaid.render(id, String(code)));
  } catch (e) {
    // mermaid's parse errors name the line and the token, which is exactly what
    // the model needs to fix its own syntax. Passed through rather than wrapped.
    throw new Error(`圖表語法有誤：${(e && e.message) || String(e)}`);
  } finally {
    // mermaid leaves its scratch element behind on failure.
    const junk = document.getElementById('d' + id) || document.getElementById(id);
    if (junk && junk.parentNode) junk.parentNode.removeChild(junk);
  }

  // The SVG carries its own size, but not always as attributes — read the
  // viewBox when width/height are percentages or missing.
  const box = /viewBox="([\d.\-\s]+)"/.exec(svg);
  let w = 800;
  let h = 600;
  if (box) {
    const n = box[1].trim().split(/\s+/).map(Number);
    if (n.length === 4 && n[2] > 0 && n[3] > 0) { w = n[2]; h = n[3]; }
  }
  const cap = Math.min(1, MAX_EDGE / Math.max(w * SCALE, h * SCALE));
  const cw = Math.max(1, Math.round(w * SCALE * cap));
  const ch = Math.max(1, Math.round(h * SCALE * cap));

  // A data: URL rather than a blob: URL — a blob is another thing to revoke, and
  // an add-in that leaks object URLs leaks them for the life of the SharedRuntime.
  const url = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(
    svg.includes('xmlns=') ? svg : svg.replace('<svg', '<svg xmlns="http://www.w3.org/2000/svg"'),
  );

  const img = await new Promise((resolve, reject) => {
    const el = new Image();
    el.onload = () => resolve(el);
    el.onerror = () => reject(new Error('SVG 轉成圖片時失敗——通常是圖裡有這個環境沒有的字型。'));
    el.src = url;
  });

  const canvas = document.createElement('canvas');
  canvas.width = cw;
  canvas.height = ch;
  const c = canvas.getContext('2d');
  // White, not transparent: a transparent PNG on a white page looks fine until
  // someone prints it or pastes it onto a coloured slide.
  c.fillStyle = '#ffffff';
  c.fillRect(0, 0, cw, ch);
  c.drawImage(img, 0, 0, cw, ch);

  return {
    data: canvas.toDataURL('image/png').split(',')[1],
    // The SVG source is carried alongside the raster, not instead of it. Word
    // can only take the PNG; PowerPoint can take either, and with ImageCoercion
    // 1.2 it takes this one — vector, sharp on a projector, and convertible to
    // editable shapes. Rendering twice to serve both would be wasteful when the
    // SVG is already in hand.
    svg,
    width: w,
    height: h,
  };
}
