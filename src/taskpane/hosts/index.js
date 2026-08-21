/**
 * hosts — pick the module that knows how to drive whatever Office app this is.
 *
 * Every host module exports the same surface. The shell (app.js) holds a
 * reference to one of them and never asks which; that is what lets 2,000 lines
 * of chat UI, sign-in, streaming, confirmation and undo be written once instead
 * of three times. See docs/multi-host.md.
 *
 * ── Why the import is dynamic ──
 * `import()` and not a static import of all three, because the risk this
 * structure exists to control is Excel's bugs reaching Word. A static import
 * would evaluate every host module in every host — one bad module-level line in
 * excel.js and the Word pane fails to start, which is exactly the coupling three
 * separate projects were supposed to buy. Dynamic import means Word never loads
 * Excel's code at all. The guarantee comes from the structure, not from anyone
 * remembering to be careful.
 *
 * ── The contract ──
 * Each module exports:
 *
 *   meta              descriptor — see hosts/excel.js for what is in it
 *   supportedTools()  what THIS build can execute on THIS host
 *   executeTool(tool, args, opts)  → { ok, content?, error?, snapshot? }
 *   inspectWrite(tool, args)       → what a write would cost, for the dialog
 *   restoreSnapshot(snap)          → { kind, … } describing what came back
 *   goTo(a, b)                     → follow a citation; false if not found
 *   getDocumentName()              → title bar text
 *   getDocumentKey()               → stable id, so a renamed file keeps history
 *   getLocation()  / watchLocation(cb)   → where in the file the user is
 *   getSelectionRef() / watchSelection(cb) → what they have selected
 *
 * executeTool must ALWAYS resolve. A tool that throws strands the agent waiting
 * for a result that never comes.
 */

/**
 * Office.HostType → module path.
 *
 * Missing entries are the answer for the hosts we do not serve: Office loads
 * the pane wherever the manifest says it may, and "Outlook" landing here should
 * produce the polite refusal in loadHost(), not a 404 on a chunk.
 */
const MODULES = {
  Excel: () => import('./excel.js'),
  PowerPoint: () => import('./powerpoint.js'),
  Word: () => import('./word.js'),
};

/**
 * Resolve the host module for the app we are running inside.
 *
 * Returns null when this is not a host we implement — the caller shows the
 * "wrong app" message. Office.HostType values are strings ('Excel', 'Word'), so
 * comparing on `info.host` directly keeps this module free of any Office API
 * call at import time; it is loaded before Office.onReady in some paths.
 */
export async function loadHost(officeHost) {
  const load = MODULES[String(officeHost)];
  if (!load) return null;
  return load();
}

/** For the "this add-in only works in …" message, before a host is resolved. */
export const SUPPORTED_HOSTS = Object.keys(MODULES);
