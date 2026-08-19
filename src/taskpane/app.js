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
  executeTool, inspectWrite, restoreSnapshot, selectCell, getWorkbookName, SUPPORTED_TOOLS,
} from './excelTools.js';

/** Matches PORT in ai-agents-office/.env. Editable on the login screen. */
const DEFAULT_API_BASE = 'http://localhost:12054';
const LS_TOKEN = 'jv_excel_token';
const LS_API = 'jv_excel_api';

const state = {
  apiBase: localStorage.getItem(LS_API) || DEFAULT_API_BASE,
  token: localStorage.getItem(LS_TOKEN) || '',
  /** Claude CLI session — echoed back each turn so the agent keeps its context. */
  sessionId: null,
  runId: null,
  streaming: false,
  workbookName: '',
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
 * The "working on it" indicator lives IN the transcript, at the end, where the
 * answer is about to appear — not pinned to the bottom of the pane. When it sat
 * down by the composer there was no way to tell which part of the log it belonged
 * to, or whether the text above it was finished.
 *
 * It is always the last element, so a tool call that happens mid-answer shows up
 * below the text so far, in the order things actually happened.
 */
let thinkingEl = null;

function setStatus(text) {
  if (thinkingEl) { thinkingEl.remove(); thinkingEl = null; }
  if (!text) return;
  clearEmptyState();
  thinkingEl = document.createElement('div');
  thinkingEl.className = 'bubble ai thinking';
  thinkingEl.setAttribute('role', 'status');
  thinkingEl.innerHTML =
    '<span class="dots" aria-hidden="true"><i></i><i></i><i></i></span>'
    + `<span>${escapeHtml(text)}</span>`;
  $('log').appendChild(thinkingEl);
  scrollDown();
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
function resolveCard(card, kind, text) {
  card.className = `confirm resolved ${kind}`;
  card.removeAttribute('role');
  if (kind === 'err') card.setAttribute('role', 'alert');
  const mark = { ok: '✓', denied: '×', err: '!', running: '' }[kind] ?? '';
  card.innerHTML = mark
    ? `<span class="mark">${mark}</span>${escapeHtml(text)}`
    : `<span class="dots" aria-hidden="true"><i></i><i></i><i></i></span>${escapeHtml(text)}`;
  scrollDown();
}

/**
 * The confirmation card. Resolves true/false on click and hands back the element
 * so the caller can report what actually happened in the same place the user
 * made the decision.
 */
function askConfirm(summary, detail) {
  return new Promise((resolve) => {
    clearEmptyState();
    const card = document.createElement('div');
    card.className = 'confirm';
    card.setAttribute('role', 'alertdialog');
    const risk = detail && detail.occupied
      ? `<div class="risk">目標範圍已有 ${detail.occupied} 格有內容${detail.hadFormula ? `（其中 ${detail.hadFormula} 格是公式）` : ''}，會被覆蓋。</div>`
      : '';
    card.innerHTML = `
      <div class="confirm-head">
        <svg class="ic ic-sm" viewBox="0 0 20 20" aria-hidden="true"><use href="#i-warn"/></svg>
        <span class="confirm-title">AI 想要修改這份活頁簿</span>
      </div>
      <div class="confirm-body">${escapeHtml(summary)}</div>
      ${detail && detail.address ? `<div class="confirm-addr">${escapeHtml(prettyAddress(detail.address))}</div>` : ''}
      ${risk}
      ${state.sources.length ? `<div class="confirm-src">此對話可讀取：${escapeHtml(sourceLabel(state.sources))}。寫入內容可能包含這些來源的資料。</div>` : ''}
      <div class="confirm-actions">
        <button class="btn-approve">允許</button>
        <button class="btn-deny">拒絕</button>
      </div>`;
    $('log').appendChild(card);
    scrollDown();

    const done = (ok) => {
      // ✓ / × rather than emoji: they inherit colour and exist in every font in
      // the stack, so they render identically on Windows and in dark mode.
      if (ok) resolveCard(card, 'running', '執行中…');
      else resolveCard(card, 'denied', `已拒絕：${summary}`);
      resolve({ approved: ok, card });
    };
    card.querySelector('.btn-approve').addEventListener('click', () => done(true));
    card.querySelector('.btn-deny').addEventListener('click', () => done(false));
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
};

async function handleToolRequest(payload) {
  const { callId, tool, args, needsConfirm, summary } = payload;
  setStatus(TOOL_LABEL[tool] || `正在執行 ${tool}…`);
  let card = null;

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
    const decision = await askConfirm(summary, detail);
    card = decision.card;
    if (!decision.approved) {
      await postToolResult(state.runId, callId, { ok: false, error: '使用者拒絕了這次操作。請不要重試，改問他想怎麼做。' });
      setStatus('');
      return;
    }
  }

  const result = await executeTool(tool, args);
  if (result.snapshot) {
    state.undoStack.push(result.snapshot);
    updateUndoButton();
  }

  // The outcome lands back in the card the user just answered, instead of a
  // separate line below it — same information, one place, and the card stops
  // looking like an open question.
  if (card) {
    if (result.ok === false) resolveCard(card, 'err', result.error || '執行失敗');
    else resolveCard(card, 'ok', result.content || summary);
  } else if (result.ok === false) {
    addNote(result.error || '執行失敗', 'err');
  }

  await postToolResult(state.runId, callId, result);
  setStatus('思考中…');
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
  setStatus('思考中…');

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
        clientTools: SUPPORTED_TOOLS,
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

          if (ev.type === 'text') {
            // Drop the indicator BEFORE creating the bubble, so the bubble takes
            // its place at the end of the log rather than appearing above it.
            setStatus('');
            text += ev.data;
            if (!bubble) bubble = addBubble('ai', '');
            bubble.innerHTML = renderMarkdown(text);
            scrollDown();
          } else if (ev.type === 'tool_request') {
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
            setStatus('');
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
    setStatus('');
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
 * Ask which sign-in methods exist before drawing the form. In pro-panjit,
 * email/password is whitelist-only — showing it as the default hands a PANJIT
 * employee a form that can only ever answer 帳號或密碼錯誤.
 */
async function loadAuthMode() {
  let mode = { adLogin: false, emailLogin: true };
  try {
    const res = await fetch(`${state.apiBase}/api/auth/mode`);
    if (res.ok) mode = await res.json();
  } catch { /* unreachable server: fall back to email + let them fix the URL */ }
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
  void loadAuthMode();
}

function showChat() {
  $('login').classList.remove('show');
  $('chat').classList.add('show');
  updateUndoButton();
  updateSourceToggle();
  void loadDataSources();
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

  try { state.workbookName = await getWorkbookName(); } catch { /* non-fatal */ }
  $('wb-name').textContent = state.workbookName || '（未命名活頁簿）';

  for (const tab of document.querySelectorAll('#login-tabs .tab')) {
    tab.addEventListener('click', () => {
      $('login-error').textContent = '';
      setLoginMode(tab.dataset.mode);
    });
  }

  $('login-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    $('login-error').textContent = '';
    const apiBase = $('api-base').value.trim().replace(/\/+$/, '') || DEFAULT_API_BASE;
    try {
      const token = state.loginMode === 'ad'
        ? await loginAd($('ad-username').value.trim(), $('ad-password').value, $('ad-domain').value.trim(), apiBase)
        : await login($('email').value.trim(), $('password').value, apiBase);
      state.token = token; state.apiBase = apiBase;
      localStorage.setItem(LS_TOKEN, token);
      localStorage.setItem(LS_API, apiBase);
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
