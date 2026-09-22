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

// ── 版式：使用者說「文青一點」的時候不能只換顏色 ──────────────────────────
//
// 「我全部都覺得醜」的原因是五個配色做出來是同一個版面的五種上色。規則要明說
// 風格的要求該用 look 回應，否則模型會拿換配色來交差——那是答非所問。
ok('有版式這一節', ctx.includes('### 版式'));
ok('每種版式都列出來',
  ['business', 'editorial', 'literary', 'minimal', 'soft', 'bold', 'tech',
    'academic', 'dashboard', 'courseware', 'handmade', 'whiteboard', 'eink']
    .every(k => ctx.includes(k)));
// 紋理那三個最容易被誤會成「要接生圖 API」。規則要把這件事講死，
// 否則模型會去問使用者要金鑰，或乾脆不敢用。
ok('說明紋理是本機用 SVG 算的，不是 AI 生圖', ctx.includes('不是 AI 生圖'));
ok('而且強調轉換為圖形後還能編輯', ctx.includes('轉換為圖形'));
ok('講出為什麼不走整頁生成', ctx.includes('改一個錯字就要重畫一整頁'));
// 使用者說「文青」的時候，模型要知道那是 literary，不是把 editorial 換個顏色。
ok('文青對應到 literary', /literary[\s\S]{0,80}文青/.test(ctx));
ok('風格自帶配色，不要另外指定 theme', ctx.includes('每個版式自帶配色'));
ok('而且解釋只換一半會發生什麼事', ctx.includes('只換一半就不是那個風格'));
ok('說清楚配色與版式是兩回事', ctx.includes('決定顏色') && ctx.includes('決定長相'));
ok('風格要求不能用換配色回應', ctx.includes('不要用換配色來回應風格的要求'));
ok('沒指定就用預設，不要自己挑特別的',
  ctx.includes('不要自己挑一個特別的'));
ok('版式整份只選一次', ctx.includes('一路用到底'));

// ── 視覺不是使用者要開口才有的東西 ──────────────────────────────────────────
//
// 使用者說的是「幫我做一份 AI 趨勢簡報」，然後看到十頁純文字，覺得這個工具很醜。
// 他不會回頭想「是我沒講」——他也不該需要講。
ok('規劃階段就要決定每頁的視覺', ctx.includes('規劃的時候就要決定每一頁的視覺'));
ok('連續兩頁純文字就要停下來想', ctx.includes('連續兩頁純文字'));
ok('一半以上的內容頁要有文字以外的東西', ctx.includes('一半以上的內容頁'));
ok('封面和章節頁一定要有視覺', ctx.includes('封面和章節頁一定要有視覺'));
// 直接假設沒有合適的素材，是簡報看起來很空的主因。
ok('不確定就先去找，不要先假設沒有', ctx.includes('不確定放什麼的時候，先去找'));

// ── 封面 ───────────────────────────────────────────────────────────────────
ok('有封面這一節', ctx.includes('### 封面值得多花一點力氣'));
ok('八種構圖都列出來',
  ['panel', 'split', 'stack', 'triptych', 'bleed', 'band', 'plain', 'paper']
    .every((k) => ctx.includes(k)));
ok('提醒不要每次都同一種', ctx.includes('不要每次都用同一種構圖'));
ok('適合的話封面就放滿版背景圖', ctx.includes('封面就放一張滿版背景圖'));
// 「適合」要定義，否則要嘛每頁都放圖、要嘛一張都不敢放。
ok('而且定義了什麼叫適合', ctx.includes('什麼叫「適合」'));
ok('不適合的題目有替代方案', ctx.includes('純幾何的構圖反而更俐落'));

// ── 三層結構：密度的具體做法 ──────────────────────────────────────────────
//
// 「下半頁是空的」是我們跟參考簡報最大的差距，而模型對空版面的預設解法是
// 把字放大。規則要說的不只是「不要放大」，還要說「那該放什麼」——
// 沒有替代做法的禁令不會被遵守。
ok('有上中下三層這一節', ctx.includes('## 讓一頁有「上中下三層」'));
ok('三層各指名了工具',
  ctx.includes('ppt_add_kpi') && ctx.includes('ppt_add_callout'));
// 三次呼叫就是三個各自猜出來的 top，落差就是使用者看到的空白。
ok('一頁只呼叫一次 ppt_build_slide', ctx.includes('一頁就呼叫一次'));
ok('而且說清楚為什麼拆開會出現空白', ctx.includes('沒有人負責整頁的高度'));
ok('單獨的那兩個工具留給事後補件', ctx.includes('事後補一個到既有頁面上'));
ok('說清楚數據卡跟卡片的差別',
  ctx.includes('卡片裝一句話，數據卡裝一個數字'));
ok('洞察框只寫一句', ctx.includes('只放一句'));
ok('內容少的頁面要開頁尾帶', ctx.includes('footband'));
ok('但不要求每頁都湊滿三層', ctx.includes('三層不必每頁都湊滿'));

// ── 照片：能力有了，節制才是規則 ────────────────────────────────────────────
//
// 「一頁滿版」是使用者要的，但一份簡報每頁都貼美圖就從簡報變成相簿。
// 這一節要同時說得出「怎麼做」跟「什麼時候別做」。
// 插畫比照片更常是對的答案，而模型的預設反應是去找照片——規則要把這件事顛倒過來。
ok('有插畫這一節', ctx.includes('### 插畫'));
ok('說清楚為什麼無關的照片會讓簡報像模板', ctx.includes('假裝成論點'));
ok('抽象題目用插畫、具體東西用照片',
  ctx.includes('抽象的題目 → 插畫') && ctx.includes('具體的東西 → 照片'));
ok('插畫不要當滿版底圖', ctx.includes('不要當 hero 滿版底圖'));
ok('畫風整份要一致', ctx.includes('整份簡報固定用同一種'));

ok('有照片這一節', ctx.includes('### 照片'));
ok('三種模式都說明用途',
  ctx.includes('hero') && ctx.includes('side') && ctx.includes('block'));
ok('一份簡報的照片有上限', ctx.includes('最多兩到三張'));
ok('無關的美圖不要放', ctx.includes('跟內容無關的美圖不要放'));
ok('hero 一定要有遮罩', ctx.includes('hero 一定要有遮罩'));
ok('連續用 side 要左右交替', ctx.includes('左右要交替'));
// 抓不到圖的時候不斷換網址重試，是最浪費使用者時間的失敗模式。
ok('抓不到圖不要一直重試', ctx.includes('失敗不要一直換網址重試'));

// ── 章節與大數字：一致性，以及「不明所以」 ────────────────────────────────
//
// 使用者實測後的兩個抱怨。大數字沒有那句解釋就是一道問答題；章節頁有些有有些沒有，
// 觀眾會以為自己漏看了一頁——不一致比沒有更糟。
ok('大數字不能單獨出現', ctx.includes('數字不能單獨出現'));
ok('而且說清楚三件事要到齊',
  ctx.includes('eyebrow') && ctx.includes('一句話說這個數字代表什麼'));
ok('章節頁要嘛全有要嘛全無', ctx.includes('要嘛每一段都有章節頁，要嘛一張都不要放'));
ok('而且解釋為什麼不一致更糟', ctx.includes('以為自己漏看了一頁'));
ok('章節編號要連續', ctx.includes('從 01 連續編下去'));

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
