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
import { executeTool, inspectWrite, restoreSnapshot, selectCell, getWorkbookName } from './excelTools.js';

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

function inlineMd(s) {
  let out = escapeHtml(s);
  out = out.replace(/`([^`]+)`/g, '<code>$1</code>');
  out = out.replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>');
  out = out.replace(/(^|[^*])\*([^*]+)\*/g, '$1<em>$2</em>');
  return linkCitations(out);
}

function renderMarkdown(src) {
  const lines = src.split('\n');
  const html = [];
  let inCode = false, codeBuf = [];
  let listType = null, listBuf = [];
  let tableBuf = [];

  const flushList = () => {
    if (!listBuf.length) return;
    html.push(`<${listType}>${listBuf.map(li => `<li>${inlineMd(li)}</li>`).join('')}</${listType}>`);
    listBuf = []; listType = null;
  };
  const flushTable = () => {
    if (!tableBuf.length) return;
    const rows = tableBuf
      .filter(r => !/^\s*\|?[\s:|-]+\|?\s*$/.test(r))          // drop the |---|---| separator
      .map(r => r.replace(/^\s*\|/, '').replace(/\|\s*$/, '').split('|').map(c => c.trim()));
    if (rows.length) {
      const head = rows.shift();
      html.push(
        '<div class="tablewrap"><table><thead><tr>'
        + head.map(c => `<th>${inlineMd(c)}</th>`).join('')
        + '</tr></thead><tbody>'
        + rows.map(r => `<tr>${r.map(c => `<td>${inlineMd(c)}</td>`).join('')}</tr>`).join('')
        + '</tbody></table></div>',
      );
    }
    tableBuf = [];
  };

  for (const line of lines) {
    if (/^```/.test(line)) {
      if (inCode) { html.push(`<pre><code>${escapeHtml(codeBuf.join('\n'))}</code></pre>`); codeBuf = []; }
      inCode = !inCode;
      continue;
    }
    if (inCode) { codeBuf.push(line); continue; }

    if (/^\s*\|.*\|\s*$/.test(line)) { flushList(); tableBuf.push(line); continue; }
    flushTable();

    const h = line.match(/^(#{1,4})\s+(.*)$/);
    if (h) { flushList(); html.push(`<h${h[1].length + 2}>${inlineMd(h[2])}</h${h[1].length + 2}>`); continue; }

    const ul = line.match(/^\s*[-*]\s+(.*)$/);
    if (ul) { if (listType !== 'ul') { flushList(); listType = 'ul'; } listBuf.push(ul[1]); continue; }

    const ol = line.match(/^\s*\d+[.)]\s+(.*)$/);
    if (ol) { if (listType !== 'ol') { flushList(); listType = 'ol'; } listBuf.push(ol[1]); continue; }

    flushList();
    if (line.trim() === '') continue;
    html.push(`<p>${inlineMd(line)}</p>`);
  }
  flushList(); flushTable();
  if (inCode && codeBuf.length) html.push(`<pre><code>${escapeHtml(codeBuf.join('\n'))}</code></pre>`);
  return html.join('');
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

function setStatus(text) {
  const s = $('status');
  $('status-text').textContent = text || '';
  s.classList.toggle('on', !!text);
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

/**
 * The confirmation card. Resolves true/false on click, and — importantly — stays
 * in the transcript afterwards showing what was decided, so the user can scroll
 * back and see exactly which writes they approved.
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
      ${detail && detail.address ? `<div class="confirm-addr">${escapeHtml(detail.address)}</div>` : ''}
      ${risk}
      <div class="confirm-actions">
        <button class="btn-approve">允許</button>
        <button class="btn-deny">拒絕</button>
      </div>`;
    $('log').appendChild(card);
    scrollDown();

    const done = (ok) => {
      card.querySelector('.confirm-actions').innerHTML =
        `<span class="decided">${ok ? '✓ 已允許' : '✕ 已拒絕'}</span>`;
      resolve(ok);
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
};

async function handleToolRequest(payload) {
  const { callId, tool, args, needsConfirm, summary } = payload;
  setStatus(TOOL_LABEL[tool] || `正在執行 ${tool}…`);

  if (needsConfirm) {
    // Look before asking: the dialog is only useful if it says what will be lost.
    let detail = null;
    try { detail = await inspectWrite(args); } catch { /* fall back to the plain summary */ }
    if (detail && detail.ok === false) {
      await postToolResult(state.runId, callId, { ok: false, error: detail.error });
      setStatus('');
      return;
    }
    const approved = await askConfirm(summary, detail);
    if (!approved) {
      await postToolResult(state.runId, callId, { ok: false, error: '使用者拒絕了這次操作。請不要重試，改問他想怎麼做。' });
      setStatus('');
      return;
    }
  }

  const result = await executeTool(tool, args);
  if (result.snapshot) {
    state.undoStack.push(result.snapshot);
    updateUndoButton();
    addNote(`已寫入 ${result.snapshot.sheet}!${result.snapshot.address}`, 'ok');
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
    const res = await fetch(`${state.apiBase}/api/excel/chat`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', Authorization: `Bearer ${state.token}` },
      body: JSON.stringify({
        message,
        runId: state.runId,
        sessionId: state.sessionId || undefined,
        workbookName: state.workbookName || undefined,
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
}

function showChat() {
  $('login').classList.remove('show');
  $('chat').classList.add('show');
  updateUndoButton();
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

  $('login-form').addEventListener('submit', async (e) => {
    e.preventDefault();
    $('login-error').textContent = '';
    const apiBase = $('api-base').value.trim().replace(/\/+$/, '') || DEFAULT_API_BASE;
    try {
      const token = await login($('email').value.trim(), $('password').value, apiBase);
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

  if (state.token) showChat(); else showLogin('');
});
