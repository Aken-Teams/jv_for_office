/**
 * recorder — turn what actually happened into something worth saving.
 *
 * ── The problem this solves, which VBA got wrong ──
 * Excel's macro recorder records KEYSTROKES, which is why recorded macros are
 * notorious: they are full of `Selection.Select`, of going somewhere and coming
 * back, of doing a thing wrong and then doing it right. The recorder cannot tell
 * which steps were mistakes, so it keeps all of them, and the macro reproduces
 * the fumbling along with the result.
 *
 * We can tell. The pane does not just see WHAT was called, it sees how each call
 * ENDED — whether the tool failed, whether the person declined it at the
 * confirmation card, whether they hit undo afterwards, whether a later call
 * wrote over the same place. Those four signals are facts we recorded, not
 * guesses, so the dead ends can be removed before anybody tries to describe the
 * procedure.
 *
 * That split matters. Asking a model 「which of these steps was a mistake」 gets
 * an answer that is sometimes right. Asking it 「write these six steps up as an
 * instruction」, having already removed the four that were undone, gets an answer
 * that is reliably right — because the hard part was done with arithmetic.
 *
 * ── No Office.js and no DOM in here ──
 * Arrays in, arrays out, so every rule below is testable without an Excel to
 * stand up. See test/recorder.test.mjs.
 */

/** Reads are cheap to repeat and carry no risk, so they never need pruning for safety. */
const READ_PREFIXES = ['get_', 'read_', 'search', 'trace_'];

/**
 * Is this call a read?
 *
 * Matched on the part after the host prefix — `excel_read_range` → `read_range`
 * — so one rule covers all three hosts rather than three lists that drift.
 */
export function isRead(tool) {
  const name = String(tool || '').replace(/^(excel|word|ppt)_/, '');
  return READ_PREFIXES.some((p) => name.startsWith(p));
}

/**
 * Arguments that say WHERE a call landed.
 *
 * Used to spot a later call that wrote over an earlier one. Deliberately a small
 * fixed list rather than the whole args object: two writes to the same range
 * with different VALUES are the same target, and that is exactly the pair where
 * only the last one belongs in the procedure.
 */
const TARGET_ARGS = ['sheet', 'range', 'output_sheet', 'slide', 'to', 'name', 'shape'];

/** A stable identity for "the place this call acted on". */
export function targetOf(entry) {
  const a = entry.args || {};
  const parts = TARGET_ARGS
    .filter((k) => a[k] !== undefined && a[k] !== null && a[k] !== '')
    .map((k) => `${k}=${String(a[k])}`);
  return `${entry.tool}|${parts.join('&')}`;
}

/**
 * Remove the dead ends.
 *
 * Four rules, in the order they matter:
 *
 *   1. FAILED — the tool returned an error. Whatever came next is the real step.
 *   2. REJECTED — the person said no at the confirmation card. It never happened.
 *   3. UNDONE — they hit 復原 afterwards. It happened and was taken back, which
 *      for the purpose of writing down a procedure is the same as not happening.
 *   4. SUPERSEDED — a later call wrote to the same target. Only the last stands.
 *
 * Reads are exempt from rule 4: reading the same range twice is not one read
 * overwriting another, and a procedure that says 「先讀標題列」 is better for
 * having that step in it. Exact repeats are still collapsed, because a read
 * issued twice with identical arguments says nothing the first one did not.
 */
export function prune(entries) {
  const live = (entries || []).filter((e) => e && e.tool && !e.failed && !e.rejected && !e.undone);

  // Walk backwards so "the last one wins" falls out of the order rather than
  // needing a second pass to decide which of the survivors is latest.
  const seenTarget = new Set();
  const seenExact = new Set();
  const kept = [];
  for (let i = live.length - 1; i >= 0; i--) {
    const e = live[i];
    const exact = `${e.tool}|${JSON.stringify(e.args || {})}`;
    if (seenExact.has(exact)) continue;
    seenExact.add(exact);

    if (!isRead(e.tool)) {
      const target = targetOf(e);
      if (seenTarget.has(target)) continue;
      seenTarget.add(target);
    }
    kept.push(e);
  }
  return kept.reverse();
}

/** What was thrown away, so the pane can say so instead of silently shrinking. */
export function pruneSummary(entries, kept) {
  const all = entries || [];
  return {
    total: all.length,
    kept: kept.length,
    failed: all.filter((e) => e.failed).length,
    rejected: all.filter((e) => e.rejected).length,
    undone: all.filter((e) => e.undone).length,
  };
}

/** Values long enough to be DATA rather than a setting are not worth carrying. */
const MAX_ARG_CHARS = 80;

/**
 * One argument, as the model should see it.
 *
 * Big payloads are replaced by their shape. A `values` array holding 400 rows of
 * somebody's ledger is not a parameter of the procedure — it is this month's
 * data — and putting it in a saved template would both bloat it and store real
 * figures on the server for replay next month.
 */
function briefValue(v) {
  if (Array.isArray(v)) {
    // A grid is data — this month's numbers, not part of the procedure.
    if (Array.isArray(v[0])) return `（${v.length} 列 × ${v[0].length} 欄的資料）`;
    // A short flat list is a SETTING, and usually the most important one in the
    // call: `key_column: ['料號','倉別','批號']` is the single hardest thing to
    // get right and the whole reason the procedure is worth saving. Summarising
    // it to 「3 筆資料」 threw away the answer and kept the packaging.
    const flat = v.map((x) => String(x)).join('、');
    if (v.length <= 8 && flat.length <= MAX_ARG_CHARS) return flat;
    return `（${v.length} 筆）`;
  }
  if (v && typeof v === 'object') return '（一組設定）';
  const s = String(v);
  return s.length > MAX_ARG_CHARS ? s.slice(0, MAX_ARG_CHARS) + '…' : s;
}

/**
 * The pruned log, written out for the model to turn into an instruction.
 *
 * Tool names are kept verbatim rather than translated: they are what the model
 * will call next time, and 「比對工具」 is not something it can invoke. The values
 * are abbreviated because the point of the saved procedure is the SHAPE of the
 * work — which sheets, which key columns, what order — not this month's numbers.
 */
export function describe(kept) {
  if (!kept.length) return '';
  return kept.map((e, i) => {
    const a = e.args || {};
    const args = Object.keys(a)
      .filter((k) => a[k] !== undefined && a[k] !== null && a[k] !== '')
      .map((k) => `${k}=${briefValue(a[k])}`)
      .join('、');
    return `${i + 1}. ${e.tool}${args ? `（${args}）` : ''}`;
  }).join('\n');
}

/**
 * The instruction that asks for the write-up.
 *
 * Written here rather than inline at the call site because it is the whole
 * feature: everything above only decides WHICH steps to hand over, and this
 * decides what gets made of them. The three rules it carries are the three ways
 * a generated template goes wrong — too specific to reuse, too vague to follow,
 * or quietly carrying this month's data into next month.
 */
export function buildPrompt(kept, summary) {
  const dropped = summary.total - summary.kept;
  return '請把我剛剛在這個對話裡實際做過的事，整理成一段**之後可以重複使用的指令**。\n\n'
    + '下面是實際執行過的步驟（已經自動剔除失敗、被我拒絕、以及被我復原的步驟'
    + (dropped > 0 ? `，共剔除 ${dropped} 步` : '')
    + '）：\n\n'
    + describe(kept)
    + '\n\n寫的時候請遵守：\n'
    + '1. **寫成我下次可以直接說出口的一段話**，不是步驟編號清單，也不要提到工具名稱。\n'
    + '2. **把這次的具體值換成佔位**——寫「用本月的檔案」而不是「用 2024-03-報表.xlsx」，'
    + '但**保留那些不會變的設定**，例如用哪幾個欄位當鍵、要不要忽略大小寫。\n'
    + '3. **只寫最短路徑**。上面的步驟已經是結果，不要把我摸索的過程寫進去。\n\n'
    + '只要輸出那段指令本身，前後不要加說明。';
}
