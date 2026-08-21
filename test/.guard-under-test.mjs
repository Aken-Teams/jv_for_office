/* ── guard:start ─ test/script-guard.test.mjs slices the source from here ──── */

/**
 * Methods that READ. Anything not in here counts as changing the workbook.
 *
 * Stated in this direction after the other one failed. The first version carried
 * a list of MUTATING methods, and a denylist is only worth what its completeness
 * is worth: `worksheet.copy()`, `range.autoFill()`, `range.setCellProperties()`,
 * `workbook.close()`, `table.clearFilters()` and the whole `applyValuesFilter`
 * family were all absent from it. Every one of those wrote to the workbook with
 * no confirmation card and no snapshot — and the server's own pattern matcher
 * missed them too, so nothing was watching. Office.js gains methods with every
 * release, so the list would have gone stale again on its own schedule.
 *
 * Reading has a shape that writing does not: `get…` for navigation and
 * retrieval, plus the handful of names below. So this rule can be stated once
 * and stay true, which is the same property the assignment rule above has, and
 * the reason the two now read alike — reading is a closed set, writing is
 * whatever is left over.
 *
 * The price is false positives: a read this set does not know about is treated
 * as a write and the model is told to re-request with mode="write", which puts
 * a confirmation card in front of something harmless. That is the exact habit
 * this add-in works to avoid, so a read found missing belongs in here — not
 * tolerated as an acceptable cost.
 */
const READ_METHODS = new Set([
  'load', 'sync', 'toJSON', 'run', 'track', 'untrack',
  // Search hands back ranges without touching them. Not named get*, but reads.
  'find', 'findOrNullObject', 'findAll', 'findAllOrNullObject',
]);

/**
 * Sub-objects whose whole method surface reads.
 *
 * `workbook.functions` is Excel's function library — `functions.sum(range)`,
 * `functions.vlookup(…)`. Every one of them computes and returns, none are named
 * get*, and enumerating them would be its own stale list. `trackedObjects.add()`
 * manages Office.js proxy lifetimes, not workbook content, and its name would
 * otherwise trip the rule on the way past.
 */
const READ_ONLY_SUBTREES = new Set(['functions', 'trackedObjects']);

const isReadMethod = (name) => name.startsWith('get') || READ_METHODS.has(name);

/** Reach the object underneath a wrapper, so Office.js never sees a Proxy. */
const RAW = Symbol('raw');

function makeScriptGuard(allowWrites) {
  const cache = new WeakMap();
  /** Separate, because the same object can be reached in both modes. */
  const readOnlyCache = new WeakMap();
  let touched = null;

  const deny = (what) => {
    touched = what;
    if (allowWrites) return;
    throw new Error(
      `這段程式想修改活頁簿（${what}），但這次呼叫沒有申請寫入權限。`
      + '請重新呼叫 excel_run_script 並帶上 mode="write"，使用者會看到確認視窗。'
      + '（若這其實只是讀取：唯讀的做法是 get… 開頭的方法、load、sync。）',
    );
  };

  // Plain data comes back from Office as ordinary arrays and objects — .values is
  // a real Array. Wrapping those would cost a proxy per cell and catch nothing,
  // because assigning into them changes a JS array, not the workbook.
  const isOfficeObject = (v) => {
    if (v === null) return false;
    const k = typeof v;
    if (k !== 'object' && k !== 'function') return false;
    if (Array.isArray(v) || v instanceof Date || v instanceof Promise) return false;
    const c = v.constructor;
    return !(c === Object || c === undefined);
  };

  const unwrap = (v) => (isOfficeObject(v) && v[RAW]) || v;

  // `force` exists for the Excel namespace itself. It is a plain object, so the
  // heuristic below rejects it — and leaving it unwrapped hands any script an
  // unguarded context through Excel.run(cb).
  //
  // `readOnly` marks everything reached through a READ_ONLY_SUBTREES branch, so
  // `functions.sum()` is not mistaken for a mutation on the way past.
  function wrap(target, force, readOnly) {
    // A collection's `.items` is a plain Array holding real Office objects, and
    // the isOfficeObject heuristic rejects arrays — right for `.values` and
    // `.formulas`, which hold primitives, but it handed `worksheets.items` back
    // UNWRAPPED and every object inside it was an open door:
    //
    //     sheets.load('items/name'); await ctx.sync();
    //     sheets.items[0].getRange('A1').values = [[1]];   // nothing watching
    //
    // which is not an exotic way to write it — it is what the Office.js samples
    // do. Arrays of Office objects get wrapped element by element; data arrays
    // are handed back exactly as they came.
    if (Array.isArray(target)) {
      return (target.length && isOfficeObject(target[0]))
        ? target.map(v => wrap(v, false, readOnly))
        : target;
    }
    if (!force && !isOfficeObject(target)) return target;
    if (target === null || (typeof target !== 'object' && typeof target !== 'function')) return target;
    const store = readOnly ? readOnlyCache : cache;
    const hit = store.get(target);
    if (hit) return hit;

    const proxy = new Proxy(target, {
      get(t, prop) {
        if (prop === RAW) return t;
        // Reflect against the target, not the receiver: Office objects keep state
        // in internal slots and reject a Proxy as `this`.
        const v = Reflect.get(t, prop, t);
        const name = String(prop);
        const childReadOnly = readOnly || READ_ONLY_SUBTREES.has(name);
        if (typeof v !== 'function') return wrap(v, false, childReadOnly);
        return function (...a) {
          if (!readOnly && !isReadMethod(name)) deny(name + '()');
          // A callback passed to Office (Excel.run, forEach) is handed raw
          // objects when Office calls it back — re-wrap them on the way in, or
          // the script gets an unguarded context through the front door.
          const args = a.map(x => (typeof x === 'function'
            ? (...back) => x(...back.map(b => wrap(b, false, readOnly)))
            : unwrap(x)));
          return wrap(Reflect.apply(v, t, args), false, readOnly);
        };
      },
      set(t, prop, value) {
        deny(String(prop) + ' =');
        return Reflect.set(t, prop, unwrap(value), t);
      },
      deleteProperty(t, prop) {
        deny('delete ' + String(prop));
        return Reflect.deleteProperty(t, prop);
      },
    });

    store.set(target, proxy);
    return proxy;
  }

  return {
    wrap: (v) => wrap(v, false, false),
    wrapNamespace: (v) => wrap(v, true, false),
    didMutate: () => touched,
  };
}


export { makeScriptGuard };
