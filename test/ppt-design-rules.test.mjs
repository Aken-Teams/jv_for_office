/**
 * The PowerPoint add-in's design rules.
 *
 * ── Why a prompt gets a test ──
 * README says it plainly: what decides whether the output looks any good is the
 * 排版與美感 section of the system prompt, not the tools. That makes this file
 * the highest-leverage text in the PowerPoint half of the product — and the
 * easiest to weaken by accident, because prose has no compiler and a rule that
 * quietly disappears in an edit takes the quality with it and nothing fails.
 *
 * So the rules with teeth get pinned. Not the wording — the wording should be
 * free to improve — but the fact that each rule is still SAID.
 *
 * ── Scope ──
 * The add-in only. The web platform builds decks through skills/pptx-gen with a
 * different toolset and is deliberately untouched; the check at the end makes
 * sure this file is not quietly measuring that one instead.
 */
import fs from 'fs';

let pass = 0, fail = 0;
const ok = (name, cond) => {
  console.log(`${cond ? 'PASS' : 'FAIL'}  ${name}`);
  if (cond) pass++; else fail++;
};

const serverDir = new URL('../../ai-agents-office/server/src/', import.meta.url);
if (!fs.existsSync(serverDir)) {
  console.log('      （找不到 ai-agents-office，跳過 PPT 設計規則的跨 repo 檢查）');
  process.exit(0);
}
const ctx = fs.readFileSync(new URL('services/pptContext.ts', serverDir), 'utf8');

// ── 節奏：一張一張合格、整份很平，是最常見的失敗 ──────────────────────────
ok('有整份簡報的節奏這一節', ctx.includes('## 整份簡報的節奏'));
ok('定義了頁面角色', ['開場', '導覽', '證據', '關係', '收束'].every(r => ctx.includes(r)));
ok('定義了強度', ctx.includes('強度'));
// 這兩條是節奏規則裡唯一「可檢查」的部分——沒有它們，角色只是分類法。
ok('要求至少一個停頓與一個高峰',
  /強度 ≤2 的停頓[\s\S]{0,20}≥4 的高峰/.test(ctx));
ok('限制同一種角色不能連續太多頁', ctx.includes('連續超過三頁'));

// ── 禁止清單：命名比原則有效 ──────────────────────────────────────────────
ok('有不要做的事這一節', ctx.includes('## 不要做的事'));
// 這一條是針對我們自己的：ppt_add_cards 做出來很漂亮，所以會被用在不該用的地方。
ok('點名等寬圓角卡片牆', ctx.includes('等寬圓角卡片牆'));
ok('而且說清楚卡片該用在哪', ctx.includes('三到五個並列、地位相同的項目'));
ok('點名沒有功能的色塊', ctx.includes('沒有功能的色塊'));
ok('點名通用漸層', ctx.includes('通用漸層背景'));
// 內容少就放大字撐版面，是「看起來很空」最常見的錯誤解法。
ok('內容少要先收攏而不是放大',
  ctx.includes('不要把東西放大撐版面') && ctx.includes('資訊變多了才往外擴張'));

// ── 標題與數字 ────────────────────────────────────────────────────────────
ok('主標要寫主張不是標籤', ctx.includes('主標寫主張，不是標籤'));
ok('關鍵數字至少跟標題一樣醒目', ctx.includes('關鍵數字至少要跟它的標題一樣醒目'));
ok('斷行要斷在語意的地方', ctx.includes('斷行要斷在語意的地方'));

// ── 配色的職責 ────────────────────────────────────────────────────────────
ok('主色與強調色分工', ctx.includes('主色跟強調色各有各的工作'));
ok('重心用真實內容算，不用透明框假裝置中', ctx.includes('假裝置中'));

// ── 大綱階段就要標角色 ────────────────────────────────────────────────────
//
// 投影片不能重新排序。節奏如果等到生成之後才發現不對，就已經來不及了。
{
  const skill = new URL('skills/ppt-narrative/SKILL.md', serverDir);
  if (fs.existsSync(skill)) {
    const body = fs.readFileSync(skill, 'utf8');
    ok('敘事線技能要求大綱就標角色與強度', body.includes('[證據‧3]'));
    ok('而且說明為什麼不能等到生成後', body.includes('已經不能重新排序'));
  }
}

// ── 範圍：不可以順手改到網頁平台那條路 ────────────────────────────────────
{
  const users = [];
  for (const f of fs.readdirSync(new URL('routes/', serverDir))) {
    if (!f.endsWith('.ts')) continue;
    const src = fs.readFileSync(new URL(`routes/${f}`, serverDir), 'utf8');
    if (src.includes('PPT_ASSISTANT_SYSTEM_PROMPT')) users.push(f);
  }
  ok('這份提示詞只服務增益集那條路', users.length === 1 && users[0] === 'ppt.ts');
}

console.log(`\n${pass} passed, ${fail} failed`);
if (fail) process.exit(1);
