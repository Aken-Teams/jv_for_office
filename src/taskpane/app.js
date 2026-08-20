/**
 * app — the task pane: chat UI, SSE client, confirmation gate, undo stack.
 *
 * The interesting part is `handleToolRequest`. Everything else is a chat window.
 *
 * ── Why the confirmation lives HERE ──
 * The server runs the Claude CLI with --permission-mode dontAsk: there is no human
 * in that loop, so the CLI's own permission prompts are useless to us. The only
 * place a person can actually see a proposed write and decide is this pane. So the
 * server marks destructive calls (`needsConfirm`) and this file is what stands
 * between an instruction hidden in a downloaded spreadsheet and someone's real
 * data. Do not "streamline" it away.
 */
import {
  executeTool, inspectWrite, restoreSnapshot, selectCell, getWorkbookName, supportedTools,
  getActiveSheetName, watchActiveSheet,
} from './excelTools.js';

/** Last-resort fallback if config.json is missing or unreadable. */
const DEFAULT_API_BASE = 'http://localhost:12054';
const LS_TOKEN = 'jv_excel_token';
/**
 * Only set when someone typed an address into the revealed 伺服器 field. Kept
 * separate from the configured value so a deployment that changes config.json
 * isn't silently overridden by whatever happened to be cached in this pane.
 */
const LS_API_OVERRIDE = 'jv_excel_api_override';
/**
 * Whether writes execute without a click. Persisted, unlike `sources`: this one
 * is a standing preference about how someone wants to work, and re-arming it
 * every time Excel restarts would just train them to switch it on without
 * reading. The cost of persisting it is that a forgotten setting is exactly how
 * an injected instruction lands — which is why the header button is coloured as
 * a warning whenever it is on, and why irreversible operations still ask.
 */
const LS_AUTO_APPROVE = 'jv_excel_auto_approve';

const state = {
  apiBase: localStorage.getItem(LS_API_OVERRIDE) || DEFAULT_API_BASE,
  token: localStorage.getItem(LS_TOKEN) || '',
  /** Claude CLI session — echoed back each turn so the agent keeps its context. */
  sessionId: null,
  runId: null,
  streaming: false,
  workbookName: '',
  /** The sheet the user is on — shown so it is clear where the agent is working. */
  sheetName: '',
  /** Snapshots of everything we've overwritten this session, newest last. */
  undoStack: [],
  /** Which cross-source MCPs the server says this user COULD mount. */
  available: { email: false, km: false },
  /** Why they aren't available, shown on the disabled toggle. */
  sourceHint: '正在確認可用的資料來源…',
  /** 'ad' | 'email' — which sign-in form is showing. */
  loginMode: 'email',
  /** Where the web app lives, for first-time AD setup. */
  webUrl: '',
  /**
   * Which ones are switched ON right now. Off by default and held only in
   * memory: the pane runs in a SharedRuntime, so this survives closing and
   * reopening the pane but resets when Excel is quit. Reaching someone's mailbox
   * should expire with the session, not persist quietly in localStorage.
   */
  sources: [],
  /** true = destructive tools run without a click. See LS_AUTO_APPROVE. */
  autoApprove: localStorage.getItem(LS_AUTO_APPROVE) === '1',
};

const $ = (id) => document.getElementById(id);

// ── Markdown (a deliberate subset) ────────────────────────────────────────────

function escapeHtml(s) {
  return s.replace(/[&<>"']/g, (c) => (
    { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]
  ));
}

/**
 * Citations are the whole trust mechanism: [[損益表!C42]] becomes a link that
 * selects that cell. If the model cites nothing, the user has no way to check it —
 * which is why the system prompt insists on them.
 */
function linkCitations(html) {
  return html.replace(
    /\[\[([^\]!]+)!(\$?[A-Za-z]{1,3}\$?\d{1,7}(?::\$?[A-Za-z]{1,3}\$?\d{1,7})?)\]\]/g,
    (_m, sheet, addr) =>
      `<a href="#" class="cite" data-sheet="${escapeHtml(sheet)}" data-addr="${escapeHtml(addr)}">${escapeHtml(sheet)}!${escapeHtml(addr)}</a>`,
  );
}

/**
 * Inline formatting.
 *
 * Code spans are pulled out into placeholders BEFORE the other rules run, then
 * put back at the end — otherwise `**not bold**` inside a code span would get
 * bolded, which is exactly backwards.
 */
function inlineMd(s) {
  const codes = [];
  // NUL as the delimiter, not spaces or brackets: it cannot occur in model
  // output, so the restore pass can never collide with a number in the prose.
  let out = escapeHtml(s).replace(/`([^`]+)`/g, (_m, c) => `\u0000${codes.push(c) - 1}\u0000`);

  // [text](url). Only http(s) and mailto survive as links — the text of a message
  // can originate in a spreadsheet cell, so `javascript:` must never become an
  // href here. Anything else degrades to plain text.
  out = out.replace(/\[([^\]\n]+)\]\(([^)\s]+)\)/g, (_m, text, url) =>
    /^(https?:|mailto:)/i.test(url) ? `<a class="link" href="${url}">${text}</a>` : text);

  out = out.replace(/\*\*([^*\n]+)\*\*/g, '<strong>$1</strong>');
  out = out.replace(/(^|[^*])\*([^*\n]+)\*/g, '$1<em>$2</em>');
  out = linkCitations(out);
  return out.replace(/\u0000(\d+)\u0000/g, (_m, i) => `<code>${codes[i]}</code>`);
}

const RE_FENCE = /^\s*```/;
const RE_HR = /^\s{0,3}(-{3,}|\*{3,}|_{3,})\s*$/;
const RE_HEAD = /^(#{1,6})\s+(.*)$/;
const RE_QUOTE = /^\s{0,3}>\s?(.*)$/;
const RE_ROW = /^\s*\|.*\|\s*$/;
const RE_SEP = /^\s*\|?[\s:|-]+\|?\s*$/;
const RE_ITEM = /^(\s*)([-*+]|\d+[.)])\s+(.*)$/;

/** Nested lists from indentation. Two levels is all a 350px column can show. */
function renderList(items) {
  let html = '';
  const stack = [];
  const close = () => { html += `</${stack.pop()}>`; };
  for (const it of items) {
    while (stack.length > it.level + 1) close();
    if (stack.length === it.level + 1 && stack[stack.length - 1] !== it.type) close();
    while (stack.length < it.level + 1) { html += `<${it.type}>`; stack.push(it.type); }
    html += `<li>${inlineMd(it.text)}</li>`;
  }
  while (stack.length) close();
  return html;
}

function renderTable(rows) {
  const cells = rows
    .filter(r => !RE_SEP.test(r))                    // drop the |---|---| separator row
    .map(r => r.replace(/^\s*\|/, '').replace(/\|\s*$/, '').split('|').map(c => c.trim()));
  if (!cells.length) return '';
  const head = cells.shift();
  return '<div class="tablewrap"><table><thead><tr>'
    + head.map(c => `<th>${inlineMd(c)}</th>`).join('')
    + '</tr></thead><tbody>'
    + cells.map(r => `<tr>${r.map(c => `<td>${inlineMd(c)}</td>`).join('')}</tr>`).join('')
    + '</tbody></table></div>';
}

/**
 * Block-level markdown → HTML.
 *
 * A deliberate subset, but a complete one for what the model actually emits:
 * headings, paragraphs, lists (nested), tables, fenced code, block quotes,
 * horizontal rules. Written as a line scanner rather than regex-over-the-whole-
 * string because this runs on every streamed chunk — dozens of times per answer.
 */
function renderMarkdown(src) {
  const lines = src.split('\n');
  const out = [];
  let i = 0;

  while (i < lines.length) {
    const line = lines[i];

    // Fenced code. An unterminated fence still renders — during streaming we are
    // usually looking at a half-arrived block, and showing it beats showing nothing.
    if (RE_FENCE.test(line)) {
      const buf = [];
      i++;
      while (i < lines.length && !RE_FENCE.test(lines[i])) buf.push(lines[i++]);
      i++;
      out.push(`<pre><code>${escapeHtml(buf.join('\n'))}</code></pre>`);
      continue;
    }

    if (!line.trim()) { i++; continue; }

    if (RE_HR.test(line)) { out.push('<hr>'); i++; continue; }

    const h = line.match(RE_HEAD);
    if (h) {
      // h1/h2 in a 350px pane would be comically large, so the whole scale is
      // shifted down two levels.
      const level = Math.min(6, h[1].length + 2);
      out.push(`<h${level}>${inlineMd(h[2])}</h${level}>`);
      i++;
      continue;
    }

    if (RE_QUOTE.test(line)) {
      const buf = [];
      while (i < lines.length && RE_QUOTE.test(lines[i])) buf.push(lines[i++].match(RE_QUOTE)[1]);
      out.push(`<blockquote>${inlineMd(buf.join(' '))}</blockquote>`);
      continue;
    }

    if (RE_ROW.test(line)) {
      const buf = [];
      while (i < lines.length && RE_ROW.test(lines[i])) buf.push(lines[i++]);
      out.push(renderTable(buf));
      continue;
    }

    if (RE_ITEM.test(line)) {
      const items = [];
      let m;
      while (i < lines.length && (m = lines[i].match(RE_ITEM))) {
        items.push({
          level: Math.min(1, Math.floor(m[1].replace(/\t/g, '  ').length / 2)),
          type: /\d/.test(m[2]) ? 'ol' : 'ul',
          text: m[3],
        });
        i++;
      }
      out.push(renderList(items));
      continue;
    }

    // Paragraph: consume until a blank line or the start of another block.
    // Single newlines become <br> rather than spaces — models use them as real
    // line breaks, and collapsing them loses the structure the model intended.
    const buf = [];
    while (i < lines.length) {
      const l = lines[i];
      if (!l.trim() || RE_FENCE.test(l) || RE_HR.test(l) || RE_HEAD.test(l)
          || RE_QUOTE.test(l) || RE_ROW.test(l) || RE_ITEM.test(l)) break;
      buf.push(l);
      i++;
    }
    out.push(`<p>${buf.map(inlineMd).join('<br>')}</p>`);
  }

  return out.join('');
}

// ── Chat rendering ────────────────────────────────────────────────────────────

function scrollDown() {
  const log = $('log');
  log.scrollTop = log.scrollHeight;
}

/** The empty state is a starting point, not a permanent header — it goes on first use. */
function clearEmptyState() {
  const e = $('empty');
  if (e) e.remove();
}

function addBubble(role, text) {
  clearEmptyState();
  const el = document.createElement('div');
  el.className = `bubble ${role}`;
  el.innerHTML = role === 'user' ? escapeHtml(text) : renderMarkdown(text || '');
  $('log').appendChild(el);
  scrollDown();
  return el;
}

/**
 * Progress panel.
 *
 * The wait here is long and lumpy — measured on this stack, 7-9 seconds pass
 * between "send" and the first tool call while the CLI process spawns and the
 * model loads the deferred MCP tools. A single unchanging 「思考中…」 for that
 * long reads as "hung", so this shows the actual steps as they happen, each with
 * how long it took.
 *
 * Every signal it needs was already arriving on the SSE stream and being
 * discarded: `spawn_started` fires when the CLI process really starts (i.e. the
 * concurrency queue let go), and `tool_activity` reports every tool the model
 * touches, including the ToolSearch call that dominates the opening gap.
 */
const progress = {
  el: null,
  steps: [],     // { label, t0, t1, detail }
  timer: null,
  /** Rolling tail of the model's thinking, used to caption the running step. */
  think: '',
};

const fmtElapsed = (ms) => (ms < 10000 ? `${(ms / 1000).toFixed(1)}s` : `${Math.round(ms / 1000)}s`);

function renderProgress() {
  if (!progress.el) return;
  const now = Date.now();
  progress.el.innerHTML = progress.steps.map((st) => {
    const done = st.t1 !== undefined;
    const ms = (done ? st.t1 : now) - st.t0;
    const mark = st.fail
      ? '<span class="mark">✕</span>'
      : done
        ? '<span class="mark">✓</span>'
        : '<span class="dots" aria-hidden="true"><i></i><i></i><i></i></span>';
    const hint = st.hint ? ` title="${escapeHtml(st.hint)}"` : '';
    return `<div class="step${done ? ' done' : ''}${st.fail ? ' fail' : ''}"${hint}>${mark}`
      + `<span class="step-label">${escapeHtml(st.label)}</span>`
      + `<span class="step-ms">${fmtElapsed(ms)}</span></div>`
      // Only on the step that is still running: a finished step's reasoning is
      // history, and keeping it would turn the panel into a wall of text.
      + (st.detail && !done ? `<div class="step-detail">${escapeHtml(st.detail)}</div>` : '');
  }).join('');
}

/**
 * Show what the model is actually reasoning about, under the running step.
 *
 * The server has been forwarding `thinking` deltas all along (claudeCli emits
 * them, the excel route passes them straight through) — the pane just dropped
 * them on the floor, which is why 30 seconds of the model composing its next
 * script looked identical to a hung spinner.
 *
 * Only the latest sentence is shown. The full stream is far too long for a 350px
 * pane, and "what is it doing right now" is answered by the most recent line, not
 * by all of them. Splitting on sentence enders means the text settles into a
 * readable phrase instead of scrolling character by character.
 */
function progressDetail(text) {
  const step = progress.steps[progress.steps.length - 1];
  if (!step || step.t1 !== undefined) return;
  step.detail = text;
  renderProgress();
}

/** Caption the running step with the model's latest line of reasoning. */
function progressThinking(chunk) {
  progress.think = (progress.think + String(chunk || '')).slice(-400);
  const parts = progress.think.split(/[\n。．.!?！？]+/).map(s => s.trim()).filter(Boolean);
  const last = parts[parts.length - 1] || '';
  progressDetail(last.length > 64 ? `${last.slice(0, 64)}…` : last);
}

/**
 * Advance to a new step. Repeating the current label is a no-op, so the
 * 「思考中」 that follows every tool result doesn't pile up duplicates.
 */
function progressStep(label) {
  const current = progress.steps[progress.steps.length - 1];
  if (current && current.t1 === undefined && current.label === label) return;

  clearEmptyState();
  if (!progress.el) {
    progress.el = document.createElement('div');
    progress.el.className = 'bubble ai thinking';
    progress.el.setAttribute('role', 'status');
    $('log').appendChild(progress.el);
    progress.timer = setInterval(renderProgress, 250);
  }
  if (current && current.t1 === undefined) current.t1 = Date.now();
  // A new step means a new train of thought; the old caption would be a lie.
  progress.think = '';
  progress.steps.push({ label, t0: Date.now() });
  renderProgress();
  scrollDown();
}

/**
 * A step that ended badly, closed immediately.
 *
 * `hint` carries the raw message as a tooltip rather than into the transcript:
 * a tool failure is addressed to the AGENT, which reads it and tries another way.
 * Printing 「pt.getRange is not a function」 in red made a task that then finished
 * successfully look like the add-in had crashed. The line stays — silently
 * swallowing a failed attempt would be worse — but it is grey, and the
 * JavaScript is available on hover for whoever wants it.
 */
function progressFailed(label, hint) {
  progressStep(label);
  const st = progress.steps[progress.steps.length - 1];
  if (st) { st.t1 = Date.now(); st.fail = true; st.hint = hint || ''; }
  renderProgress();
}

/** Tear the panel down and, if it did real work, leave one line behind. */
function progressDone(summarise) {
  if (progress.timer) { clearInterval(progress.timer); progress.timer = null; }
  const steps = progress.steps;
  if (progress.el) { progress.el.remove(); progress.el = null; }
  progress.steps = [];
  progress.think = '';

  if (summarise && steps.length > 1) {
    const total = (steps[steps.length - 1].t1 || Date.now()) - steps[0].t0;
    addNote(`${steps.length} 個步驟 · 共 ${fmtElapsed(total)}`, 'note');
  }
}

/** Kept for callers that just want "show this one thing" / "clear". */
function setStatus(text) {
  if (text) progressStep(text);
  else progressDone(false);
}

function addNote(text, kind = 'note') {
  clearEmptyState();
  const el = document.createElement('div');
  el.className = `note ${kind}`;
  if (kind === 'err') el.setAttribute('role', 'alert');
  el.textContent = text;
  $('log').appendChild(el);
  scrollDown();
  return el;
}

/** Office.js quotes sheet names containing spaces: 'SP 分析報告'!A1:F36 → SP 分析報告!A1:F36 */
function prettyAddress(address) {
  return String(address).replace(/'/g, '');
}

/**
 * Collapse a decided card into one quiet line.
 *
 * A card that has been answered must stop looking like a question. Leaving the
 * orange warning treatment on resolved cards made the transcript read as a wall
 * of unresolved alarms, and it also said the same thing three times over (the
 * summary, the address, and a separate success note underneath). Now the card
 * reports its own outcome and the note is gone.
 */
function resolveCard(card, kind, text, hint) {
  card.className = `confirm resolved ${kind}`;
  card.removeAttribute('role');
  // Only a genuine 'err' is announced to assistive tech. A 'retry' is the agent's
  // problem, not the reader's — see the mark table below.
  if (kind === 'err') card.setAttribute('role', 'alert');
  else card.removeAttribute('aria-live');
  if (hint) card.title = hint; else card.removeAttribute('title');
  const mark = { ok: '✓', denied: '×', err: '!', retry: '✕', running: '' }[kind] ?? '';
  card.innerHTML = mark
    ? `<span class="mark">${mark}</span>${escapeHtml(text)}`
    : `<span class="dots" aria-hidden="true"><i></i><i></i><i></i></span>${escapeHtml(text)}`;
  scrollDown();
}

/**
 * A question with clickable answers.
 *
 * The agent used to write 「1. 清除該列左側資料欄… 2. 先刪掉樞紐分析表… 你要哪個做法？」
 * as prose, leaving the person to type a reply describing an option that was
 * already right there. It is a multiple-choice question, so it is rendered as one.
 *
 * The free-text row is not a fallback nobody uses: when both offered options are
 * wrong, being forced to pick one is worse than not asking at all.
 *
 * Resolves to the chosen label, or whatever was typed.
 */
function askChoice({ question, options }) {
  return new Promise((resolve) => {
    clearEmptyState();
    const list = Array.isArray(options) ? options.filter(o => o && o.label) : [];
    const card = document.createElement('div');
    card.className = 'confirm choice';
    card.setAttribute('role', 'alertdialog');
    card.innerHTML = `
      <div class="confirm-head">
        <svg class="ic ic-sm" viewBox="0 0 20 20" aria-hidden="true"><use href="#i-ask"/></svg>
        <span class="confirm-title">AI 需要你決定</span>
      </div>
      <div class="confirm-body">${escapeHtml(String(question || '請你決定一件事'))}</div>
      <div class="choice-list">
        ${list.map((o, i) => `
          <button class="choice-opt" data-i="${i}">
            <span class="choice-label">${escapeHtml(o.label)}</span>
            ${o.detail ? `<span class="choice-detail">${escapeHtml(o.detail)}</span>` : ''}
          </button>`).join('')}
      </div>
      <input class="confirm-say" type="text" autocomplete="off"
             placeholder="都不是？直接告訴它你要怎麼做" aria-label="自己輸入做法" />`;
    $('log').appendChild(card);
    scrollDown();

    const say = card.querySelector('.confirm-say');
    const done = (answer) => {
      resolveCard(card, 'ok', `你選了：${answer}`);
      resolve(answer);
    };
    for (const btn of card.querySelectorAll('.choice-opt')) {
      btn.addEventListener('click', () => done(list[Number(btn.dataset.i)].label));
    }
    say.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && say.value.trim()) { e.preventDefault(); done(say.value.trim()); }
    });
  });
}

/**
 * The confirmation card. Resolves true/false on click and hands back the element
 * so the caller can report what actually happened in the same place the user
 * made the decision.
 */
/**
 * Does this call destroy data the undo button cannot bring back?
 *
 * The server decides, because the answer depends on the CONTENTS of the call —
 * a script containing sheet.getRange("371:373").delete() is exactly as final as
 * excel_structure_op delete_rows, and only the server sees the code. Used for
 * two things: the red card, and the one carve-out 自動允許 keeps asking about.
 *
 * If you want 自動允許 to mean literally everything, change this to `false`.
 */
function isIrreversible(payload) {
  return payload.risk === 'high';
}

/** A card in the log for something that ran without being asked about. */
function autoApprovedCard(summary) {
  clearEmptyState();
  const card = document.createElement('div');
  card.className = 'confirm';
  $('log').appendChild(card);
  // Still rendered, still in the transcript, still undoable. 自動允許 removes the
  // click, not the record — otherwise nobody could tell afterwards what the agent
  // had done on its own.
  resolveCard(card, 'running', `自動允許：${summary}`);
  scrollDown();
  return card;
}

/**
 * The confirmation card.
 *
 * Three ways out, not two. "拒絕" alone tells the agent nothing except to stop,
 * so it comes back and asks what to do instead — a whole extra round trip to
 * learn something the person already knew when they clicked. The text field
 * carries that intent with the answer: type "改用 IFS，不要巢狀 IF" and the
 * refusal arrives WITH the correction, so the next call is the corrected one.
 *
 * Resolves { approved, card, note }.
 */
function askConfirm(summary, detail, meta, highRisk) {
  return new Promise((resolve) => {
    clearEmptyState();
    const card = document.createElement('div');
    // Two levels, because a font change and a sheet deletion should not shout
    // with the same voice. Everything reaching this card mutates the workbook;
    // `danger` is reserved for the ones the undo button cannot take back.
    card.className = highRisk ? 'confirm danger' : 'confirm';
    card.setAttribute('role', 'alertdialog');
    const risk = detail && detail.occupied
      ? `<div class="risk">目標範圍已有 ${detail.occupied} 格有內容${detail.hadFormula ? `（其中 ${detail.hadFormula} 格是公式）` : ''}，會被覆蓋。</div>`
      : '';
    card.innerHTML = `
      <div class="confirm-head">
        <svg class="ic ic-sm" viewBox="0 0 20 20" aria-hidden="true"><use href="#i-warn"/></svg>
        <span class="confirm-title">${highRisk ? '這個動作沒辦法復原' : 'AI 想要修改這份活頁簿'}</span>
      </div>
      <div class="confirm-body">${escapeHtml(summary)}</div>
      ${meta ? `<div class="confirm-meta">${escapeHtml(meta)}</div>` : ''}
      ${detail && detail.address ? `<div class="confirm-addr">${escapeHtml(prettyAddress(detail.address))}</div>` : ''}
      ${risk}
      ${state.sources.length ? `<div class="confirm-src">此對話可讀取：${escapeHtml(sourceLabel(state.sources))}。寫入內容可能包含這些來源的資料。</div>` : ''}
      <input class="confirm-say" type="text" autocomplete="off"
             placeholder="想改成什麼？（選填）" aria-label="給 AI 的說明（選填）" />
      <div class="confirm-actions">
        <button class="btn-approve">允許</button>
        <button class="btn-deny">拒絕</button>
      </div>`;
    $('log').appendChild(card);
    scrollDown();

    const say = card.querySelector('.confirm-say');
    const deny = card.querySelector('.btn-deny');

    // The deny button relabels itself once there is something to say: with an
    // instruction attached it is no longer "no", it is "do it this way instead",
    // and the label should say which one the click will do.
    const syncLabel = () => { deny.textContent = say.value.trim() ? '改成這樣' : '拒絕'; };
    say.addEventListener('input', syncLabel);

    const done = (ok) => {
      const note = say.value.trim();
      // ✓ / × rather than emoji: they inherit colour and exist in every font in
      // the stack, so they render identically on Windows and in dark mode.
      if (ok) resolveCard(card, 'running', '執行中…');
      else resolveCard(card, 'denied', note ? `改成這樣：${note}` : `已拒絕：${summary}`);
      resolve({ approved: ok, card, note });
    };
    card.querySelector('.btn-approve').addEventListener('click', () => done(true));
    deny.addEventListener('click', () => done(false));
    // Enter sends the instruction rather than approving. Someone who has just
    // typed 「不要這樣做」 pressing Enter must not have it executed.
    say.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && say.value.trim()) { e.preventDefault(); done(false); }
    });
  });
}

// ── Bridge: run a tool the agent asked for ────────────────────────────────────

async function postToolResult(runId, callId, result) {
  try {
    await fetch(`${state.apiBase}/api/excel/tool-result`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${state.token}` },
      body: JSON.stringify({
        runId, callId,
        ok: result.ok !== false,
        content: result.content,
        error: result.error,
      }),
    });
  } catch (e) {
    // Nothing useful to do: the server-side call will hit its own timeout and the
    // agent will be told the workbook stopped responding.
    console.error('[jv] tool-result post failed', e);
  }
}

const TOOL_LABEL = {
  excel_get_overview: '正在讀取活頁簿結構…',
  excel_get_selection: '正在讀取選取範圍…',
  excel_read_range: '正在讀取資料…',
  excel_search: '正在搜尋…',
  excel_write_range: '等待你確認寫入…',
  excel_manage_sheet: '正在調整工作表…',
  excel_format_range: '正在套用格式…',
  excel_create_chart: '正在建立圖表…',
  excel_structure_op: '等待你確認結構變更…',
  excel_table_ops: '正在整理表格…',
  excel_create_pivot: '等待你確認建立樞紐…',
  excel_trace_precedents: '正在追蹤公式來源…',
  excel_sheet_style: '正在調整版面…',
  excel_run_script: '正在執行程式…',
};

/**
 * What the model is doing BEFORE the call is dispatched.
 *
 * Distinct from TOOL_LABEL, which describes the call executing. Composing a long
 * tool call is generation time — for excel_run_script that is the model writing
 * an entire script, and it is usually the longest single stretch of a turn.
 */
const PREPARING_LABEL = {
  excel_run_script: '正在寫程式…',
  excel_write_range: '正在整理要寫入的資料…',
  excel_create_chart: '正在設定圖表…',
  excel_create_pivot: '正在設定樞紐分析表…',
  excel_format_range: '正在決定格式…',
  excel_manage_sheet: '正在準備工作表變更…',
};

/** MCP tools arrive prefixed (`mcp__excel__excel_run_script`). */
function preparingLabel(name) {
  const short = String(name).replace(/^mcp__excel__/, '');
  return PREPARING_LABEL[short] || TOOL_LABEL[short] || `正在準備 ${short}…`;
}

const fmtSize = (n) => (n >= 1024 ? `${(n / 1024).toFixed(1)} KB` : `${n || 0} 字元`);

async function handleToolRequest(payload) {
  const { callId, tool, args, needsConfirm, summary, meta } = payload;
  // Prefer the server's specific description ("讀取 工作表1!A1:AZ375") over the
  // generic per-tool label — it is the same line the confirmation card shows, so
  // the trace and the dialog agree about what happened.
  setStatus(summary || TOOL_LABEL[tool] || `正在執行 ${tool}…`);

  // Answered by a person, not by Office.js — so it is handled here and never
  // reaches executeTool. The bridge gives this call a human-scale timeout.
  if (tool === 'excel_ask_user') {
    const answer = await askChoice(args || {});
    await postToolResult(state.runId, callId, { ok: true, content: answer });
    setStatus('思考中…');
    return;
  }

  let card = null;
  /** Anything the person typed on the confirmation card. */
  let userNote = '';

  if (needsConfirm) {
    // Look before asking: the dialog is only useful if it says what will be lost.
    // Only cell writes can be inspected this way — for a sheet delete or a sort
    // there is no "target range" to count, so those fall back to the summary line.
    let detail = null;
    if (tool === 'excel_write_range') {
      try { detail = await inspectWrite(args); } catch { /* fall back to the plain summary */ }
      if (detail && detail.ok === false) {
        await postToolResult(state.runId, callId, { ok: false, error: detail.error });
        setStatus('');
        return;
      }
    }
    // 自動允許 skips the click, but only after the same inspection: the snapshot
    // that inspectWrite triggers is what makes the write undoable, so it still
    // runs and a broken target still fails here rather than mid-write.
    if (state.autoApprove && !isIrreversible(payload)) {
      card = autoApprovedCard(summary);
    } else {
      const decision = await askConfirm(summary, detail, meta, isIrreversible(payload));
      card = decision.card;
      userNote = decision.note || '';
      if (!decision.approved) {
        // A refusal that carries the person's own instruction is worth far more
        // than a bare "no": the agent can act on it immediately instead of
        // spending a round trip asking what they wanted.
        await postToolResult(state.runId, callId, {
          ok: false,
          error: userNote
            ? `使用者不同意這樣做，他的說明是：「${userNote}」。請照他說的調整，不要重複同一個呼叫。`
            : '使用者拒絕了這次操作。請不要重試，改問他想怎麼做。',
        });
        setStatus('');
        return;
      }
    }
  }

  const result = await executeTool(tool, args);
  // Approving WITH a note means "do this, and here is something you should know".
  // Riding it back on the tool result gets it to the agent within the same turn,
  // instead of the person having to repeat themselves in the next message.
  if (userNote && result.ok !== false) {
    result.content = `${result.content || ''}

【使用者補充】${userNote}`;
  }
  if (result.snapshot) {
    state.undoStack.push(result.snapshot);
    updateUndoButton();
  }

  // The outcome lands back in the card the user just answered, instead of a
  // separate line below it — same information, one place, and the card stops
  // looking like an open question.
  if (card) {
    // 'retry', not 'err'. An approved action that Excel rejected is something the
    // AGENT then works around — the run in the screenshot went on to finish. Both
    // halves matter: the person consented to something that did not happen, so the
    // line stays; but 「[PivotTableRangeConflict] 操作失敗…」 in red reads as a
    // crash. Grey line, plain words, raw message on hover.
    if (result.ok === false) resolveCard(card, 'retry', '沒成功，AI 正在換方法', result.error || '執行失敗');
    else resolveCard(card, 'ok', result.content || summary);
  } else if (result.ok === false) {
    progressFailed('沒成功，AI 正在換方法', result.error || '執行失敗');
  }

  await postToolResult(state.runId, callId, result);
  setStatus('思考中…');
  // How much came back is the number that matters here. excel_run_script exists so
  // the model gets a CONCLUSION, not a data dump; a result measured in tens of KB
  // means a script is returning raw rows and has quietly reintroduced the very
  // problem the tool was added to solve — and it explains the silent gap that
  // follows, which is the model reading all of it back in.
  const size = String(result.content || result.error || '').length;
  if (size) progressDetail(`已讀入 ${fmtSize(size)} 的結果`);
}

function updateUndoButton() {
  const n = state.undoStack.length;
  $('undo').disabled = n === 0;
  const badge = $('undo-count');
  badge.textContent = n;
  badge.hidden = n === 0;
}

// ── The chat stream ───────────────────────────────────────────────────────────

async function send(message) {
  if (state.streaming) return;
  state.streaming = true;
  state.runId = crypto.randomUUID();
  $('send').disabled = true;

  addBubble('user', message);
  progressStep('連線到伺服器…');

  let bubble = null;
  let text = '';

  try {
    // Prefetch the workbook structure on the first turn and ship it WITH the
    // question.
    //
    // Measured on this stack: the bridge round trip is ~140ms, but the model
    // round trip that precedes it is 7-9 SECONDS (CLI spawn → ToolSearch to load
    // the deferred MCP tools → decide to call excel_get_overview). Since the
    // agent opens essentially every conversation by asking for exactly this, we
    // can hand it over for free and skip that entire leg. Simple questions can
    // then be answered with no tool call at all.
    //
    // First turn only: later turns resume the CLI session, which already has it.
    let workbookContext;
    if (!state.sessionId) {
      const r = await executeTool('excel_get_overview', { preview_rows: 5 });
      if (r.ok) workbookContext = r.content;
    }

    const res = await fetch(`${state.apiBase}/api/excel/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${state.token}` },
      body: JSON.stringify({
        message,
        runId: state.runId,
        sessionId: state.sessionId || undefined,
        workbookName: state.workbookName || undefined,
        workbookContext,
        // What this build can execute. The server hides everything else from the
        // model, so a stale pane degrades to "fewer abilities" instead of an
        // agent looping on 未知的工具.
        // Computed per host: pivot tables drop out on builds below ExcelApi 1.12.
        clientTools: supportedTools(),
        dataSources: state.sources,
      }),
    });

    if (res.status === 401) { logout('登入已過期，請重新登入。'); return; }
    if (!res.ok) {
      const err = await res.json().catch(() => ({}));
      addNote(err.error || `伺服器錯誤 (${res.status})`, 'err');
      return;
    }

    const reader = res.body.getReader();
    const decoder = new TextDecoder();
    let buf = '';

    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      buf += decoder.decode(value, { stream: true });

      let idx;
      while ((idx = buf.indexOf('\n\n')) >= 0) {
        const chunk = buf.slice(0, idx);
        buf = buf.slice(idx + 2);
        for (const line of chunk.split('\n')) {
          if (!line.startsWith('data:')) continue;   // ': keepalive' comments
          let ev;
          try { ev = JSON.parse(line.slice(5).trim()); } catch { continue; }

          if (ev.type === 'spawn_started') {
            // The concurrency/auth queue let go and the CLI process actually started.
            progressStep('AI 啟動中…');
          } else if (ev.type === 'tool_activity') {
            // Three different things arrive on this event, and dropping the last
            // two is what made a 42-second wait look like a frozen spinner: the
            // model spends that time WRITING the tool call (a whole script, in
            // the case of excel_run_script), and the server has been announcing
            // it from the first token.
            const d = ev.data || {};
            const t = d.tool;
            // A tool call starting means the remark before it is finished. Segment
            // here as well as on tool_request: ToolSearch sits between two
            // preambles and produces no tool_request, which is how
            // 「先載入工具並跑資料分析。先掃描工作表1…」 ended up as one sentence.
            if (d.status === 'running' && t) { bubble = null; text = ''; }
            if (t === 'ToolSearch') {
              // The CLI loading the deferred MCP tools — the single longest step
              // of the opening gap, so it gets named rather than hidden.
              progressStep('載入 Excel 工具…');
            } else if (d.status === 'writing') {
              progressDetail(`已寫出 ${fmtSize(d.chars)}`);
            } else if (d.status === 'running' && t) {
              progressStep(preparingLabel(t));
            }
          } else if (ev.type === 'text') {
            // Drop the panel BEFORE creating the bubble, so the bubble takes
            // its place at the end of the log rather than appearing above it.
            progressDone(true);
            text += ev.data;
            if (!bubble) bubble = addBubble('ai', '');
            bubble.innerHTML = renderMarkdown(text);
            scrollDown();
          } else if (ev.type === 'thinking') {
            // Caption the running step with what the model is reasoning about.
            progressThinking(ev.data);
          } else if (ev.type === 'tool_request') {
            // Close the current bubble. Whatever the model says AFTER this tool
            // comes back is a new remark, not a continuation of the last one —
            // appending it to the same bubble is what produced run-on text like
            // 「先讀一下資料全貌…先把完整資料讀進來…」, two separate sentences
            // glued together. A fresh bubble per turn also lets the tool step
            // land between them, which is what makes the trace readable.
            bubble = null;
            text = '';
            // Not awaited: the stream must keep draining while Office.js works and
            // while a confirmation dialog waits on a human.
            void handleToolRequest(ev.data);
          } else if (ev.type === 'session_id') {
            state.sessionId = ev.data;
          } else if (ev.type === 'info') {
            addNote(typeof ev.data === 'string' ? ev.data : JSON.stringify(ev.data), 'note');
          } else if (ev.type === 'error') {
            addNote(typeof ev.data === 'string' ? ev.data : JSON.stringify(ev.data), 'err');
          } else if (ev.type === 'done') {
            if (ev.data && ev.data.sessionId) state.sessionId = ev.data.sessionId;
            progressDone(false);
          }
        }
      }
    }
  } catch (e) {
    addNote(`連線失敗：${e.message}。確認伺服器有在跑，而且 CORS_ORIGINS 允許 ${location.origin}。`, 'err');
  } finally {
    state.streaming = false;
    state.runId = null;
    $('send').disabled = !$('input').value.trim();
    progressDone(false);
  }
}

// ── Auth ──────────────────────────────────────────────────────────────────────

async function login(email, password, apiBase) {
  const res = await fetch(`${apiBase}/api/auth/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ email, password }),
  });
  const body = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(body.error || `登入失敗 (${res.status})`);
  return body.token;
}

/**
 * AD sign-in — the primary path in pro-panjit.
 *
 * Worth knowing: this is also what acquires the user's Outlook token
 * (authenticateOutlook runs server-side on success), so signing in with AD is
 * what makes the 郵件 data source light up. Email/password sign-in never does.
 *
 * First-time AD users get a setup wizard instead of a token. That flow is
 * several screens of account claiming and is not something to reimplement in a
 * 350px pane — we send them to the web app once and they come back.
 */
async function loginAd(username, password, domain, apiBase) {
  const res = await fetch(`${apiBase}/api/auth/ad/login`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ username, password, domain: domain || undefined }),
  });
  const body = await res.json().catch(() => ({}));
  if (!res.ok) throw new Error(body.error || `登入失敗 (${res.status})`);
  if (body.firstLogin) {
    throw new Error(
      '這是你第一次用 AD 帳號登入，請先到 AI Agents Office 網頁版'
      + (state.webUrl ? `（${state.webUrl}）` : '')
      + '完成首次設定，再回來登入。',
    );
  }
  return body.token;
}

/** Draw the sign-in form the server actually accepts. */
function setLoginMode(mode) {
  state.loginMode = mode;
  $('fields-ad').hidden = mode !== 'ad';
  $('fields-email').hidden = mode !== 'email';
  for (const tab of document.querySelectorAll('#login-tabs .tab')) {
    tab.setAttribute('aria-selected', String(tab.dataset.mode === mode));
  }
}

/**
 * Read the deployed configuration.
 *
 * The API address is deployment configuration, not a user preference — nobody
 * outside this repo knows where the server lives, and a typo produces a login
 * failure that looks exactly like a wrong password. So it comes from
 * config.json (generated from .env) and the field stays hidden.
 *
 * An explicit override wins: it only gets set when someone typed an address
 * after a connection failure, which is precisely the case where config.json is
 * the thing that's wrong.
 */
async function loadConfig() {
  const override = localStorage.getItem(LS_API_OVERRIDE);
  if (override) { state.apiBase = override; return; }
  try {
    const res = await fetch('../../config.json', { cache: 'no-store' });
    if (res.ok) {
      const cfg = await res.json();
      if (cfg.apiBase) state.apiBase = String(cfg.apiBase).replace(/\/+$/, '');
    }
  } catch { /* fall back to the default */ }
}

/** Reveal the server field so a wrong address can be fixed without a redeploy. */
function revealServerField(msg) {
  $('api-base-field').hidden = false;
  $('api-base').value = state.apiBase;
  if (msg) $('login-error').textContent = msg;
}

/**
 * Ask which sign-in methods exist before drawing the form. In pro-panjit,
 * email/password is whitelist-only — showing it as the default hands a PANJIT
 * employee a form that can only ever answer 帳號或密碼錯誤.
 */
async function loadAuthMode() {
  let mode = { adLogin: false, emailLogin: true };
  try {
    const res = await fetch(`${state.apiBase}/api/auth/mode`);
    if (res.ok) mode = await res.json();
  } catch {
    // The configured server didn't answer. Show the field so it can be corrected
    // here rather than requiring someone to redeploy config.json.
    revealServerField(`連不上伺服器 ${state.apiBase}，請確認位址或聯繫管理者。`);
  }
  state.webUrl = mode.webUrl || '';
  $('login-tabs').hidden = !(mode.adLogin && mode.emailLogin);
  setLoginMode(mode.adLogin ? 'ad' : 'email');
}

function logout(reason) {
  state.token = '';
  state.sessionId = null;
  localStorage.removeItem(LS_TOKEN);
  showLogin(reason);
}

function showLogin(msg) {
  $('login').classList.add('show');
  $('chat').classList.remove('show');
  $('login-error').textContent = msg || '';
  $('api-base').value = state.apiBase;
  void (async () => { await loadConfig(); await loadAuthMode(); })();
}

function showChat() {
  $('login').classList.remove('show');
  $('chat').classList.add('show');
  updateUndoButton();
  updateApproveToggle();
  updateSourceToggle();
  void loadDataSources();
}

/**
 * AD 組織清單。
 *
 * 同一份清單也在網頁版登入頁 client/src/app/login/page.tsx (AD_DOMAINS)。
 * 兩邊要一起改 —— 值是送給 /api/auth/ad/login 的 domain，打錯會登入失敗。
 * 之後如果組織會常變動，比較好的做法是讓 /api/auth/mode 回傳這份清單，
 * 這樣改組織就不用重新部署增益集。
 */
const AD_DOMAINS = [
  { value: 'PANJIT', label: 'PANJIT（台灣）' },
  { value: 'PYNMAX', label: 'PYNMAX（璟茂）' },
  { value: 'WXPJ',   label: 'WXPJ（無錫強茂）' },
  { value: 'PJWS',   label: 'PJWS（強茂深圳）' },
  { value: 'GDPJ',   label: 'GDPJ（蘇州群鑫）' },
  { value: 'PJXZ',   label: 'PJXZ（強茂徐州）' },
  { value: 'PJSD',   label: 'PJSD（山東強茂）' },
];

/** Remembered per pane so a returning user doesn't re-pick their own company. */
const LS_DOMAIN = 'jv_excel_ad_domain';

function fillDomainOptions() {
  const sel = $('ad-domain');
  sel.innerHTML = AD_DOMAINS
    .map(d => `<option value="${d.value}">${escapeHtml(d.label)}</option>`)
    .join('');
  const saved = localStorage.getItem(LS_DOMAIN);
  sel.value = AD_DOMAINS.some(d => d.value === saved) ? saved : AD_DOMAINS[0].value;
}

const SOURCE_NAMES = { email: '郵件', km: 'KM' };
const sourceLabel = (keys) => keys.map(k => SOURCE_NAMES[k] || k).join('、');

/**
 * Reflect the toggle's state.
 *
 * Always visible, disabled when nothing is mountable — with the reason on the
 * control. The first version hid it instead, which just moved the confusion:
 * a control you cannot find is indistinguishable from a feature that does not
 * exist, and the honest answer ("KM 未設定 / 尚未連結 Outlook") is short enough
 * to simply say.
 */
/**
 * Show which worksheet the agent is looking at.
 *
 * Hidden rather than blank when unknown: an empty chevron next to the filename
 * reads as a rendering bug, and on a host too old for the activation event the
 * name would simply never arrive.
 */
function setSheetName(name) {
  state.sheetName = name || '';
  const el = $('wb-sheet');
  el.textContent = state.sheetName;
  el.title = state.sheetName ? `目前的工作表：${state.sheetName}` : '';
  el.hidden = !state.sheetName;
}
/**
 * The 每次確認 / 自動允許 button.
 *
 * Both the tooltip and the aria-label spell out the current state AND what a
 * click does, because the only other signal is colour — and a control that
 * silently disables the confirmation step should never be ambiguous.
 */
function updateApproveToggle() {
  const btn = $('approve-mode');
  if (!btn) return;
  const on = state.autoApprove;
  btn.setAttribute('aria-pressed', String(on));
  // Symmetric and short enough to stay on ONE line at the narrowest pane width
  // anyone is likely to drag to. With only two states, naming the current one and
  // saying it toggles is unambiguous; what each mode actually means is spelled out
  // in the transcript the moment it changes.
  btn.dataset.tip = on ? '自動允許，點一下切換' : '每次確認，點一下切換';
  btn.setAttribute('aria-label', on ? '目前是自動允許，點一下改回每次確認' : '目前是每次確認，點一下改成自動允許');
}

function updateSourceToggle() {
  const btn = $('sources');
  btn.hidden = false;
  const keys = Object.keys(SOURCE_NAMES).filter(k => state.available[k]);

  if (!keys.length) {
    btn.disabled = true;
    btn.setAttribute('aria-pressed', 'false');
    btn.title = state.sourceHint || '這個環境沒有可用的外部資料來源';
    $('sources-label').textContent = '郵件／KM 未連結';
    return;
  }

  btn.disabled = false;
  const on = state.sources.length > 0;
  btn.setAttribute('aria-pressed', String(on));
  btn.title = on
    ? `這個對話可以讀取${sourceLabel(state.sources)}（唯讀）。點一下關閉。`
    : `點一下讓這個對話也能讀取${sourceLabel(keys)}（唯讀）。`;
  $('sources-label').textContent = on ? `${sourceLabel(state.sources)} 已開啟` : sourceLabel(keys);
}

/** Ask the server which cross-source MCPs this user could actually mount. */
async function loadDataSources() {
  try {
    const res = await fetch(`${state.apiBase}/api/excel/ping`, {
      headers: { Authorization: `Bearer ${state.token}` },
    });
    if (!res.ok) {
      state.sourceHint = `無法向伺服器確認資料來源（HTTP ${res.status}）。`;
    } else {
      const body = await res.json();
      if (!body.dataSources) {
        // An older server build that predates this endpoint's response shape.
        state.sourceHint = '伺服器尚未套用更新，請重新啟動 server 後再試。';
      } else {
        state.available = body.dataSources;
        state.sourceHint = body.dataSources.hint || '';
      }
    }
  } catch (e) {
    state.sourceHint = `連不上伺服器（${e.message}）。`;
  }
  updateSourceToggle();
}

/** Grow the composer with the text, up to the CSS max-height, then let it scroll. */
function autoGrow(el) {
  el.style.height = 'auto';
  el.style.height = `${el.scrollHeight}px`;
}

// ── Boot ──────────────────────────────────────────────────────────────────────

Office.onReady(async (info) => {
  if (info.host !== Office.HostType.Excel) {
    document.body.innerHTML = '<p style="padding:16px">這個增益集只能在 Excel 中使用。</p>';
    return;
  }

  fillDomainOptions();

  try { state.workbookName = await getWorkbookName(); } catch { /* non-fatal */ }
  $('wb-name').textContent = state.workbookName || '（未命名活頁簿）';
  $('wb-name').title = state.workbookName || '';

  try { setSheetName(await getActiveSheetName()); } catch { /* non-fatal */ }
  // Live, not just at startup: someone clicking to another tab has moved what
  // 「這一段」 and excel_get_selection refer to, and the bar should say so.
  void watchActiveSheet(setSheetName);

  for (const tab of document.querySelectorAll('#login-tabs .tab')) {
    tab.addEventListener('click', () => {
      $('login-error').textContent = '';
      setLoginMode(tab.dataset.mode);
    });
  }

  $('login-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    $('login-error').textContent = '';
    const typed = $('api-base-field').hidden ? '' : $('api-base').value.trim().replace(/\/+$/, '');
    const apiBase = typed || state.apiBase || DEFAULT_API_BASE;
    try {
      const token = state.loginMode === 'ad'
        ? await loginAd($('ad-username').value.trim(), $('ad-password').value, $('ad-domain').value, apiBase)
        : await login($('email').value.trim(), $('password').value, apiBase);
      state.token = token; state.apiBase = apiBase;
      localStorage.setItem(LS_TOKEN, token);
      if (typed) localStorage.setItem(LS_API_OVERRIDE, typed);
      if (state.loginMode === 'ad') localStorage.setItem(LS_DOMAIN, $('ad-domain').value);
      showChat();
    } catch (err) {
      $('login-error').textContent = err.message;
    }
  });

  $('composer').addEventListener('submit', (e) => {
    e.preventDefault();
    const input = $('input');
    const msg = input.value.trim();
    if (!msg) return;
    input.value = '';
    input.style.height = 'auto';
    $('send').disabled = true;
    void send(msg);
  });

  $('input').addEventListener('input', (e) => {
    autoGrow(e.target);
    $('send').disabled = !e.target.value.trim() || state.streaming;
  });

  // Suggestion chips are real prompts: one click asks the question. Delegated so
  // the chips can be regenerated without rebinding.
  $('log').addEventListener('click', (e) => {
    const chip = e.target.closest('.chip');
    if (!chip || state.streaming) return;
    void send(chip.textContent.trim());
  });

  // Enter sends, Shift+Enter is a newline — the convention every chat UI uses.
  $('input').addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && !e.shiftKey) {
      e.preventDefault();
      $('composer').requestSubmit();
    }
  });

  $('undo').addEventListener('click', async () => {
    const snap = state.undoStack.pop();
    if (!snap) return;
    updateUndoButton();
    try {
      await restoreSnapshot(snap);
      addNote(`已復原 ${snap.sheet}!${snap.address}`, 'ok');
    } catch (e) {
      addNote(`復原失敗：${e.message}`, 'err');
    }
  });

  $('sources').addEventListener('click', () => {
    const keys = Object.keys(SOURCE_NAMES).filter(k => state.available[k]);
    // One switch for all available sources: a per-source menu in a 350px bar is
    // more UI than the decision deserves — the question is "can it leave this
    // workbook", not "which of the two".
    state.sources = state.sources.length ? [] : keys;
    updateSourceToggle();
    addNote(state.sources.length
      ? `已開啟${sourceLabel(state.sources)}存取，這個對話可以讀取它們（唯讀）。`
      : '已關閉外部資料來源，這個對話只看得到活頁簿。', 'note');
  });

  $('approve-mode').addEventListener('click', () => {
    state.autoApprove = !state.autoApprove;
    localStorage.setItem(LS_AUTO_APPROVE, state.autoApprove ? '1' : '0');
    updateApproveToggle();
    // Spelled out in the transcript, not just the tooltip. Someone turning this
    // on should read once what they are turning off; someone scrolling back
    // should be able to see when it changed.
    addNote(state.autoApprove
      ? '已改成自動允許：AI 修改活頁簿時不會再詢問。刪除工作表或整列整欄仍會問，因為那沒辦法復原。'
      : '已改回每次確認：AI 要修改活頁簿前都會先問過你。', 'note');
  });

  $('logout').addEventListener('click', () => logout(''));

  // Citation clicks are delegated: bubbles are re-rendered on every streamed
  // chunk, so per-link handlers would be rebound dozens of times per message.
  $('log').addEventListener('click', async (e) => {
    const a = e.target.closest('.cite');
    if (!a) return;
    e.preventDefault();
    const ok = await selectCell(a.dataset.sheet, a.dataset.addr);
    if (!ok) addNote(`找不到工作表「${a.dataset.sheet}」`, 'err');
  });

  // External links open in the real browser. Letting one navigate the pane would
  // replace the add-in with a web page and there is no back button in here.
  $('log').addEventListener('click', (e) => {
    const a = e.target.closest('a.link');
    if (!a) return;
    e.preventDefault();
    window.open(a.href, '_blank');
  });

  if (state.token) showChat(); else showLogin('');
});
