/**
 * 使用技能 (PRD 3.1) — the wiring that decides what a task pane may offer.
 *
 * ── The mistake this file exists to prevent ──
 * Every skill written before the add-ins existed GENERATES A NEW FILE:
 * pptx-gen, docx-gen, xlsx-gen. The add-ins do the opposite — they edit the
 * document already open. Listing those in the pane would not be untidy, it
 * would be wrong: somebody picks 「PowerPoint Generator」 inside Excel and gets a
 * download instead of a changed workbook, and nothing about that reads as a bug
 * until they go looking for the file.
 *
 * So the rule is opt-IN: a skill reaches a task pane only by naming that host.
 * Every check below is a way that rule could quietly invert.
 */
import fs from 'fs';

let pass = 0, fail = 0;
const ok = (name, cond) => {
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}`);
  if (cond) pass++; else fail++;
};

const serverDir = new URL('../../ai-agents-office/server/src/', import.meta.url);
if (!fs.existsSync(serverDir)) {
  console.log('      （找不到 ai-agents-office，跳過技能的跨 repo 檢查）');
  process.exit(0);
}

const loader = fs.readFileSync(new URL('skills/loader.ts', serverDir), 'utf8');
const app = fs.readFileSync(new URL('../src/taskpane/app.js', import.meta.url), 'utf8');
const html = fs.readFileSync(new URL('../src/taskpane/taskpane.html', import.meta.url), 'utf8');

// ── 預設是 web-only，而且必須是「不寫就不給」────────────────────────────────
ok('沒宣告 hosts 的技能預設只給網頁平台', /\|\|\s*\['web'\]|:\s*\['web'\]/.test(loader));
ok('技能要主動宣告 host 才會出現在增益集', loader.includes('data.hosts'));
ok('router 不會出現在給人挑的清單裡', /skillsForHost[\s\S]*?role !== 'router'/.test(loader));

// ── 三個增益集都用 host 過濾，而不是直接照 id 取 ────────────────────────────
//
// getSkill(id) 會取到任何一個技能，包含另一個介面的。用 skillsForHost 再 find，
// 跨介面的 id 就只是找不到，而不是被照做。
for (const host of ['excel', 'word', 'ppt']) {
  const src = fs.readFileSync(new URL(`routes/${host}.ts`, serverDir), 'utf8');
  ok(`${host}: 用 skillsForHost 解析，不是 getSkill`,
    src.includes(`skillsForHost('${host}')`) && !/getSkill\(/.test(src));
  // 技能是加在主提示詞後面的。取代掉的話，agent 會有專業但沒有手——
  // 工具約定、確認規則、引用格式全在主提示詞裡。
  ok(`${host}: 技能是附加在主提示詞之後，不是取代它`,
    /const systemPrompt = \(mounted\.length[\s\S]{0,400}\+ \(skill \?/.test(src));
  ok(`${host}: 指定到不存在的技能會退回自動調度並講出來`,
    src.includes('找不到指定的技能'));
}

// ── 窗格 ────────────────────────────────────────────────────────────────────
{
  const row = (html.match(/<button[^>]*id="src-skill"[^>]*>/) || [''])[0];
  ok('技能那一列預設是隱藏的', row.includes('hidden'));
  ok('技能那一列是動作列，不是資料來源開關', /role="menuitem"/.test(row));

  const fn = (app.match(/function renderSkillMenu\(\)[\s\S]*?\n}/) || [''])[0];
  ok('沒有技能可選時整列收起來，而不是開一張空清單',
    /row\.hidden = !state\.skills\.length/.test(fn));
  ok('清單第一項是「自動調度」，而且是預設', fn.includes("id: ''") && fn.includes('自動調度'));

  ok('送出訊息時帶上 skillId', /skillId: state\.skill \|\| undefined/.test(app));
  ok('請求會帶 host，不是整包拿回來自己篩',
    /api\/skills\?host=\$\{encodeURIComponent\(host\.meta\.api\)\}/.test(app));
  // 換過 host 或 IT 下架之後，留著一個已經不存在的選擇會每輪都被伺服器拒絕。
  ok('選到的技能不在清單裡時會自動清掉',
    /state\.skill && !state\.skills\.some/.test(app));
}

// ── 至少要有一個增益集技能，否則這個功能等於沒上 ────────────────────────────
{
  const dirs = fs.readdirSync(new URL('skills/', serverDir), { withFileTypes: true })
    .filter(d => d.isDirectory());
  const addin = dirs.filter((d) => {
    const p = new URL(`skills/${d.name}/SKILL.md`, serverDir);
    if (!fs.existsSync(p)) return false;
    const head = fs.readFileSync(p, 'utf8').split('---')[1] || '';
    return /hosts:\s*\[[^\]]*(excel|word|ppt)/.test(head);
  });
  ok('至少有一個技能宣告了增益集 host', addin.length >= 1);

  // 對帳技能的價值全在這兩條規則上——少了它們就只是一段客套話。
  const rec = new URL('skills/excel-reconcile/SKILL.md', serverDir);
  if (fs.existsSync(rec)) {
    const body = fs.readFileSync(rec, 'utf8');
    ok('對帳技能明講不要自己比對', body.includes('絕對不要自己比對'));
    ok('對帳技能明講鍵通常不只一欄', body.includes('料號 ＋ 倉別 ＋ 批號'));
    ok('對帳技能要求不確定就問使用者', body.includes('excel_ask_user'));
  }
}

console.log(`\n${pass} passed, ${fail} failed`);
if (fail) process.exit(1);
