const MUTATING_METHODS = new Set([
  'add', 'addFromBase64', 'delete', 'remove', 'insert', 'clear', 'clearContents', 'clearFormats',
  'merge', 'unmerge', 'copyFrom', 'moveTo', 'autofitColumns', 'autofitRows', 'convertToRange',
  'applyFilter', 'clearFilter', 'reapplyFilters', 'apply', 'freezeRows', 'freezeColumns',
  'freezePanes', 'setPosition', 'setDynamicArray', 'setData', 'replaceAll', 'group', 'ungroup',
  'refresh', 'refreshAll', 'activate', 'select', 'setName', 'setFormat', 'setValues',
  'addConditionalFormat', 'removeDuplicates', 'sort',
]);

/** Reach the object underneath a wrapper, so Office.js never sees a Proxy. */
const RAW = Symbol('raw');

function makeScriptGuard(allowWrites) {
  const cache = new WeakMap();
  let touched = null;

  const deny = (what) => {
    touched = what;
    if (allowWrites) return;
    throw new Error(
      `這段程式想修改活頁簿（${what}），但這次呼叫沒有申請寫入權限。`
      + '請重新呼叫 excel_run_script 並帶上 mode="write"，使用者會看到確認視窗。',
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
  function wrap(target, force) {
    if (!force && !isOfficeObject(target)) return target;
    if (target === null || (typeof target !== 'object' && typeof target !== 'function')) return target;
    const hit = cache.get(target);
    if (hit) return hit;

    const proxy = new Proxy(target, {
      get(t, prop) {
        if (prop === RAW) return t;
        // Reflect against the target, not the receiver: Office objects keep state
        // in internal slots and reject a Proxy as `this`.
        const v = Reflect.get(t, prop, t);
        if (typeof v !== 'function') return wrap(v);
        const name = String(prop);
        return function (...a) {
          if (MUTATING_METHODS.has(name)) deny(name + '()');
          // A callback passed to Office (Excel.run, forEach) is handed raw
          // objects when Office calls it back — re-wrap them on the way in, or
          // the script gets an unguarded context through the front door.
          const args = a.map(x => (typeof x === 'function'
            ? (...back) => x(...back.map(wrap))
            : unwrap(x)));
          return wrap(Reflect.apply(v, t, args));
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

    cache.set(target, proxy);
    return proxy;
  }

  return {
    wrap: (v) => wrap(v, false),
    wrapNamespace: (v) => wrap(v, true),
    didMutate: () => touched,
  };
}


export { makeScriptGuard };
