/**
 * The pruning rules — the half of "recording" that VBA never got right.
 *
 * ── Why this matters more than the write-up ──
 * Excel's macro recorder keeps every keystroke, including the ones that were
 * mistakes, so a recorded macro reproduces the fumbling along with the result.
 * The fix is not a better writer; it is knowing which steps were dead ends
 * before anybody tries to describe them. We know, because the pane recorded how
 * each call ended.
 *
 * So these tests are about facts, not judgement: a call that failed, one the
 * person declined, one they undid, one a later call wrote over. Every case below
 * is a way a saved procedure could end up teaching 「do it wrong, then fix it」.
 */
import {
  isRead, targetOf, prune, pruneSummary, describe, buildPrompt,
} from '../src/taskpane/recorder.js';

let pass = 0, fail = 0;
const ok = (name, got, want = true) => {
  const good = JSON.stringify(got) === JSON.stringify(want);
  console.log(`${good ? 'PASS' : 'FAIL'}  ${name}`);
  if (good) pass++;
  else { console.log(`      得到 ${JSON.stringify(got)}，預期 ${JSON.stringify(want)}`); fail++; }
};
const tools = (list) => prune(list).map(e => e.tool);

// ── isRead ───────────────────────────────────────────────────────────────────
ok('三個 host 的讀取都認得', [
  isRead('excel_read_range'), isRead('word_get_overview'), isRead('ppt_get_selection'),
  isRead('excel_search'), isRead('excel_trace_precedents'),
], [true, true, true, true, true]);
ok('寫入不是讀取', [
  isRead('excel_write_range'), isRead('ppt_build_slide'), isRead('word_insert_table'),
], [false, false, false]);

// ── 四條剪枝規則 ─────────────────────────────────────────────────────────────

ok('失敗的步驟丟掉，留下後來成功那次',
  tools([
    { tool: 'excel_write_range', args: { sheet: 'A', range: 'A1' }, failed: true },
    { tool: 'excel_write_range', args: { sheet: 'A', range: 'B1' } },
  ]), ['excel_write_range']);

ok('使用者在確認視窗拒絕的，等於沒發生',
  tools([
    { tool: 'excel_structure_op', args: { sheet: 'A' }, rejected: true },
    { tool: 'excel_format_range', args: { sheet: 'A' } },
  ]), ['excel_format_range']);

ok('被復原的步驟丟掉',
  tools([
    { tool: 'excel_write_range', args: { sheet: 'A', range: 'A1' }, undone: true },
    { tool: 'excel_create_chart', args: { sheet: 'A' } },
  ]), ['excel_create_chart']);

ok('寫到同一個地方兩次，只留最後那次',
  prune([
    { tool: 'excel_write_range', args: { sheet: '報表', range: 'A1:B2', values: [[1]] } },
    { tool: 'excel_write_range', args: { sheet: '報表', range: 'A1:B2', values: [[2]] } },
  ]).map(e => e.args.values), [[[2]]]);

ok('寫到不同地方的不會互相吃掉',
  tools([
    { tool: 'excel_write_range', args: { sheet: '報表', range: 'A1' } },
    { tool: 'excel_write_range', args: { sheet: '報表', range: 'C1' } },
  ]), ['excel_write_range', 'excel_write_range']);

// 讀取重複不是「覆蓋」，而且「先讀標題列」本來就該留在流程裡。
ok('讀同一張表的不同範圍都留著',
  tools([
    { tool: 'excel_read_range', args: { sheet: 'A', range: 'A1:B2' } },
    { tool: 'excel_read_range', args: { sheet: 'A', range: 'D1:E9' } },
  ]), ['excel_read_range', 'excel_read_range']);

ok('一模一樣的重複呼叫還是收斂成一次',
  tools([
    { tool: 'excel_get_overview', args: {} },
    { tool: 'excel_get_overview', args: {} },
  ]), ['excel_get_overview']);

// ── 使用者真實會遇到的情況：摸索出正確的鍵欄位 ──────────────────────────────
//
// 先用料號比（結果一堆假差異）→ 復原 → 改用複合鍵再比一次。
// 存成常用指令時應該只留下對的那一次，不然下次會重演「先做錯再做對」。
{
  const session = [
    { tool: 'excel_get_overview', args: {} },
    { tool: 'excel_read_range', args: { sheet: '庫存', range: 'A1:G3' } },
    { tool: 'excel_compare_source', args: { sheet: '庫存', key_column: '料號' }, undone: true },
    { tool: 'excel_compare_source', args: { sheet: '庫存', key_column: ['料號', '倉別', '批號'] } },
  ];
  const kept = prune(session);
  ok('摸索的那一次被剔除，只留下對的做法', kept.length, 3);
  ok('留下來的是複合鍵那一次',
    kept[kept.length - 1].args.key_column, ['料號', '倉別', '批號']);

  const sum = pruneSummary(session, kept);
  ok('摘要說得出剔除了幾步、為什麼', [sum.total, sum.kept, sum.undone], [4, 3, 1]);
}

ok('全部都是死路時回空陣列',
  prune([{ tool: 'excel_write_range', args: {}, failed: true }]), []);
ok('沒有紀錄也不會炸', prune(null), []);
ok('壞掉的紀錄會被跳過', prune([null, { args: {} }, { tool: 'excel_search', args: {} }]).length, 1);

// ── targetOf ────────────────────────────────────────────────────────────────
ok('同一個目標的簽章一致，值不一樣也算同一個地方',
  targetOf({ tool: 'w', args: { sheet: 'A', range: 'A1', values: [[1]] } })
    === targetOf({ tool: 'w', args: { sheet: 'A', range: 'A1', values: [[9]] } }));
ok('不同工作表不是同一個目標',
  targetOf({ tool: 'w', args: { sheet: 'A' } }) !== targetOf({ tool: 'w', args: { sheet: 'B' } }));

// ── describe：交給模型的內容 ────────────────────────────────────────────────
{
  const text = describe(prune([
    { tool: 'excel_compare_source', args: { sheet: '庫存', key_column: ['料號', '倉別'] } },
  ]));
  ok('工具名原樣保留，模型下次才叫得動', text.includes('excel_compare_source'));
  // 鍵欄位是整個流程最難搞對、也最值得存的一件事。把它縮成「2 筆資料」
  // 等於把答案丟掉、只留下包裝。
  ok('複合鍵完整留著', text.includes('料號') && text.includes('倉別'));
}

// 這次的資料不是流程的一部分，而且把幾百列數字存進伺服器再每月重播是另一個問題。
{
  const big = describe(prune([
    { tool: 'excel_write_range', args: { sheet: 'A', values: Array.from({ length: 400 }, () => [1, 2, 3]) } },
  ]));
  ok('大筆資料只講形狀，不會把這個月的數字寫進指令裡',
    big.includes('400 列') && !big.includes('[1,2,3]'));
}

// ── buildPrompt ─────────────────────────────────────────────────────────────
{
  const kept = prune([{ tool: 'excel_compare_source', args: { sheet: 'A' } }]);
  const p = buildPrompt(kept, pruneSummary([...kept, { tool: 'x', args: {}, undone: true }], kept));
  ok('要求寫成可以直接說出口的一段話', p.includes('直接說出口'));
  ok('有講剔除了幾步', p.includes('共剔除 1 步'));

  // 實測發現的問題：第一版只叫模型看工具清單，還加了一句「不要把摸索過程
  // 寫進去」。模型照字面做——把對話裡的東西全丟了，寫出一份沒有工作表名稱、
  // 沒有 Key、沒有單位規則的清單，比直接存 AI 原本的回覆還糟。
  // 工具紀錄只有「做了什麼」，「為什麼」全在對話裡。
  ok('明講對話本身也是材料', p.includes('這整段對話'));
  ok('明講兩份材料都要用', p.includes('兩份都要用'));
  ok('點名工作表要寫進去', p.includes('工作表'));
  ok('點名 Key／對應關係要寫進去', p.includes('Key'));
  ok('點名單位與期別規則要寫進去', p.includes('單位換算') && p.includes('期別'));
  ok('點名驗證方式要寫進去', p.includes('確認結果是對的'));
  // 換佔位只該換「這一次的值」——規則和欄位名不是值。
  ok('規則與欄位名不可以被當成佔位換掉',
    p.includes('規則、欄位名稱、工作表名稱不是值'));
  // 「摸索過程」限縮成試錯，而不是泛指對話。
  ok('剔除的只有試錯，不是整段對話', p.includes('試錯') && !p.includes('最短路徑'));
}

console.log(`\n${pass} passed, ${fail} failed`);
if (fail) process.exit(1);
