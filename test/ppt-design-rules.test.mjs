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

// ── 圖表不是只有長條圖 ──────────────────────────────────────────────────────
ok('有圖表選型這一節', ctx.includes('### 圖表不是只有長條圖'));
ok('長條圖以外的圖表都列出來',
  ['waterfall', 'funnel', 'radar', 'timeline', 'map'].every((k) => ctx.includes(k)));
// 地圖最容易被拿來「讓簡報有國際感」，那正是不該放的理由。
ok('地圖只回答哪裡多哪裡少', ctx.includes('地圖只回答'));
// 歷史、據點、事件發生地都沒有「每個地區一個數字」——而那正是最常要地圖的場合。
// 只做深淺圖的話，那種題目永遠不會畫地圖，從外面看就像功能不存在。
ok('地圖有標示模式', ctx.includes('只給地區名、不給數字就是標示圖'));

// ── 這一輪新增的圖表與版型 ──────────────────────────────────────────────────
ok('散佈圖是唯一能講相關性的', ctx.includes('散佈圖是唯一能講'));
ok('堆疊圖說明並排長條看不出總和', ctx.includes('並排長條看不出總和'));
ok('百分比堆疊會藏起成長', ctx.includes('把成長藏起來'));
// 箱型圖正確但多數人看不懂——簡報要的是被看懂。
ok('箱型圖限技術場合', ctx.includes('箱型圖給技術場合'));
ok('而且說出為什麼要克制', ctx.includes('要的是被看懂，不是正確'));

ok('有圖文版型這一節', ctx.includes('### 圖文版型'));
ok('feature 與 list_image 都說明了',
  ctx.includes('feature') && ctx.includes('list_image'));
// 實測過：生出來的圖縮到 150×94 就看不出是什麼。
ok('小縮圖用內建插畫不要生圖', ctx.includes('小縮圖一律用'));
ok('而且是實測的結論', ctx.includes('實測過'));
ok('而且說明那才是最常需要地圖的場合', ctx.includes('最常需要地圖的場合'));

// 主標可以是主張，但不能是一句看不出在講什麼的漂亮話。
ok('主標要看得出在講什麼', ctx.includes('主張要看得出在講什麼'));
ok('給了可操作的檢查方式', ctx.includes('把主標單獨唸出來'));
ok('要比排名就用長條圖', ctx.includes('要比排名就用'));
ok('不要為了國際感放地圖', ctx.includes('只是想讓簡報有國際感就不要'));
ok('時程類內容要用時間軸', ctx.includes('時程、里程碑、導入階段'));
// 「現況 vs 目標」本來就是要並排看的，拆成兩頁反而讓人對照不起來。
ok('一頁可以放兩張圖表', ctx.includes('一頁可以放兩張圖表'));
// 「應該常常這樣做」而且給了座標——不給座標，模型就不會自己算。
ok('而且直接給出左右兩張的座標',
  ctx.includes('left: 38, width: 430') && ctx.includes('left: 492, width: 430'));
ok('點名雷達圖被用得太少', ctx.includes('雷達圖被用得太少'));
ok('插畫會自動保持原比例', ctx.includes('自動保持原比例置中'));
ok('不要每張都是直條圖', ctx.includes('不要每張圖表都是直條圖'));

// ── 配色不要永遠深藍配大紅 ──────────────────────────────────────────────────
ok('三個亮一點的配色都在', ['teal', 'sky', 'amber'].every((k) => ctx.includes(k)));
ok('明說企業簡報不等於深藍配大紅', ctx.includes('企業簡報不等於深藍配大紅'));
ok('提醒深色整份用下來會很重', ctx.includes('整份用下來會很重'));

// ── 章節頁與大數字是「打斷」，要克制 ────────────────────────────────────────
//
// 看報告的人是一路往下讀的。中間插一頁滿版深色或一個巨大的數字，如果那一頁
// 沒有強到值得被打斷，就只是突兀。
// 這兩個版型已經從工具的 enum 拿掉了，不是叫模型少用——提示詞規則是建議，
// 要讓一件事真的停止發生，可靠的槓桿是 schema 不是段落。
ok('明說這兩個版型已經拿掉', ctx.includes('這兩個版型已經拿掉了，不是叫你少用'));
ok('說明它們是打斷', ctx.includes('是「打斷」'));
ok('章節改成寫進 eyebrow', ctx.includes('章節寫進每一頁的'));
ok('大數字改成放進 kpi', ctx.includes('關鍵數字放進'));
ok('eyebrow 要「編號＋段落名」而且連續', ctx.includes('編號從 01 連續編下去'));
{
  const spec = fs.readFileSync(new URL('services/pptToolSpec.ts', serverDir), 'utf8');
  const kinds = /enum: \['title',[\s\S]*?\],/.exec(spec)[0];
  ok('工具的 enum 真的沒有 section', !kinds.includes("'section'"), kinds);
  ok('也沒有 statement', !kinds.includes("'statement'"), kinds);
  ok('但工具說明有解釋為什麼', spec.includes('這是刻意拿掉的'));
}

// ── 封面圖要跟標題有關 ──────────────────────────────────────────────────────
//
// 標題寫「AI 從工具變成同事」，背景是一堆抽象線條——使用者第一個反應是
// 「這張圖跟內容有什麼關係」，那比沒有圖更糟。
ok('封面圖要跟標題有關', ctx.includes('圖要跟標題有關'));
ok('而且說出無關的圖比沒有圖更糟', ctx.includes('那比沒有圖更糟'));
ok('想不出來就不要硬放', ctx.includes('想不出跟主題有關的畫面'));

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

// ── 一份報告要有多少東西 ────────────────────────────────────────────────────
//
// 「扣掉章節、目錄、首頁，其實內文只有 6 頁，然後每頁都是一張圖或一個結論，
// 感覺就不像一個 PPT 報告。」這是目前最大的問題，而且根因在前面：查得不夠深，
// 手上就真的只有那幾個數字。
ok('有資料量這一節', ctx.includes('## 一份報告要有多少東西'));
ok('明說查資料的深度決定簡報的厚度', ctx.includes('查資料的深度決定簡報的厚度'));
ok('要求一個主題查三到五個角度', ctx.includes('三到五個角度'));
// 那幾個角度彼此不相關，排隊查只會把等待疊起來——使用者看到的是一個
// 轉了兩分鐘的側邊欄。
ok('查詢要並行送出，不要一輪一輪等', ctx.includes('不要一輪一輪等'));
// 判斷句不算內容。「導入率很高」跟「88% 已導入、只有 29% 看到回報」是兩回事。
ok('每個小節都要有自己的數字', ctx.includes('都要有自己的數字或具體事實'));
ok('開工前要先盤點手上有幾個事實', ctx.includes('少於 15 個就再查'));
ok('一頁要有不只一個東西', ctx.includes('### 每一頁要有不只一個東西'));

// ── 不要每頁都長一樣 ────────────────────────────────────────────────────────
//
// 「上面一定是 KPI、下面一定是結論、中間換東西」連續五頁，讀的人第三頁就會
// 開始跳著讀，因為他已經知道每一頁哪裡有什麼。
ok('有版面輪替這一節', ctx.includes('### 不要每一頁都長一樣'));
ok('說清楚一致是骨架的工作不是內容的工作', ctx.includes('一致是骨架的工作'));
ok('列出可以輪著換的版面', ['two_column', 'zigzag', 'timeline', 'quadrant', 'quote']
  .every((k) => ctx.includes(k)));
ok('同一種不能連續兩次以上', ctx.includes('同一種不要連續出現兩次以上'));
// KPI 每頁都擺，那三個數字就不再是重點了。
ok('KPI 不是每頁都要有', ctx.includes('KPI 不是每頁都要有'));
ok('而且給了大概幾頁才配', ctx.includes('三到四頁'));
ok('三樣東西不等於固定那三格', ctx.includes('不夠的是內容，不是格子'));
ok('給出一頁的標準配置', ctx.includes('數字（發生了什麼）'));
ok('「報告」這個字對應到的頁數有寫出來', ctx.includes('12～16'));
ok('而且說明為什麼六頁不夠', ctx.includes('「報告」這個字是有份量的'));

// 章節頁與大數字從「克制」收緊到「預設零張」——使用者實測後仍然覺得突兀。

// ── 背景母題：模板感的來源 ──────────────────────────────────────────────────
//
// 「就很像 PPT 模板那樣每頁都有這個背景，而我們的圖表、文字放在這背景上。」
// 這比每頁生一張圖好在四件事：一致、免費、永遠不離題、而且可以編輯。
ok('有背景母題這一節', ctx.includes('### 背景：整份共用一層很淡的幾何'));
ok('七種母題都列出來',
  ['waves', 'diamonds', 'arcs', 'grid', 'topo', 'ribbon', 'corner'].every((k) => ctx.includes(k)));
ok('整份只挑一次', ctx.includes('整份只挑一次'));
ok('而且說出為什麼不能每頁換', ctx.includes('每頁換一個就不是模板，是拼貼'));
ok('強調它是形狀不是圖片', ctx.includes('形狀不是圖片'));
// 這一條是要把模型從「背景＝生圖」的直覺上拉回來。
ok('明說不要為了背景去生圖', ctx.includes('不要為了做背景去生圖'));
ok('也說清楚生圖該留給什麼場合', ctx.includes('不是拿來鋪底'));

// ── 圖片佔比 ────────────────────────────────────────────────────────────────
ok('三種圖片版型都列出來',
  ['imageLeft', 'imageRight', 'imageTop'].every((k) => ctx.includes(k)));
// 滿版一張圖再壓標題，除非圖極精彩，否則衝擊感反而更低。
ok('說明圖片佔一部分通常比滿版好', ctx.includes('圖片佔一部分，通常比滿版好看'));
ok('提醒不要每次封面都是滿版產圖', ctx.includes('不要每份簡報的封面都是滿版產圖'));
ok('生圖畫法預設不是照片', ctx.includes('預設不是照片'));
ok('照片只在有實體畫面時用', ctx.includes('只在題目真的有實體畫面的時候用'));
ok('同一份固定一種畫法', ctx.includes('同一份簡報固定一種畫法'));

// ── 封面上那三行字 ─────────────────────────────────────────────────────────
//
// 封面只有三個欄位，而它們最常被寫壞：主標寫成題目、副標寫成目錄。
ok('有封面文案這一節', ctx.includes('### 封面上那三行字怎麼寫'));
ok('主標要寫主張不是題目', ctx.includes('主標寫主張'));
ok('而且給了對照的例子', ctx.includes('是題目；') && ctx.includes('是主張'));
ok('副標不是目錄', ctx.includes('副標不是目錄'));
ok('副標要讓主標更有力', ctx.includes('讓主標更有力'));
ok('三行有字數上限', ctx.includes('不要超過四十個字'));
ok('說明封面為什麼特別要短', ctx.includes('唯一一頁沒有前後文'));

// ── 封面 ───────────────────────────────────────────────────────────────────
ok('有封面這一節', ctx.includes('### 封面值得多花一點力氣'));
ok('八種構圖都列出來',
  ['panel', 'split', 'stack', 'triptych', 'bleed', 'band', 'plain', 'paper']
    .every((k) => ctx.includes(k)));
ok('提醒不要每次都同一種', ctx.includes('不要每次都用同一種構圖'));
// 這條原本寫的是「適合的話就放滿版背景圖」，後來被推翻了：滿版一張圖再壓標題，
// 除非那張圖極為精彩，否則只是一個背景，衝擊感反而比構圖低。
ok('封面用圖的做法有講清楚', ctx.includes('cover') && ctx.includes('image'));
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
