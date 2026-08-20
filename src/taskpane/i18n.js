/**
 * Pane text in the language the user picked.
 *
 * The add-in writes as much into the transcript as the agent does — progress
 * captions, confirmation cards, notes about what just changed. Pinning only the
 * agent's replies produced the thing that started this: a Simplified Chinese
 * answer sitting under a Traditional Chinese sentence explaining that replies
 * are now in Simplified Chinese.
 *
 * Keys are the Traditional Chinese source strings themselves. That keeps the
 * real sentence visible at the call site, makes zh-TW free, and means an
 * untranslated string degrades to Chinese rather than to a key name like
 * `progress.thinking`. Placeholders are {0}, {1} — positional, because word
 * order moves between these languages and a sentence that hard-codes it is a
 * sentence that will read badly in at least one of them.
 *
 * SCOPE: everything after sign-in. The login and first-time setup screens stay
 * Traditional Chinese, and that is deliberate — they run before we know who the
 * person is, so there is no setting to follow yet. Getting those right means
 * reading the browser locale, which is a different decision.
 *
 * Three languages, because three is what the product has: PATCH /api/auth/profile
 * rejects anything outside [zh-TW,zh-CN,en]. Offering 日本語 here would be
 * offering a setting nothing else in the product can honour. A fourth language
 * starts on the account side, then gets a map below.
 */

const zhCN = {
  '正在確認可用的資料來源…': '正在确认可用的数据源…',
  'AI 啟動中…': 'AI 启动中…',
  '連線到伺服器…': '连接到服务器…',
  '思考中…': '思考中…',
  '正在想：{0}': '正在想：{0}',
  '{0} 個步驟 · 共 {1}': '{0} 个步骤 · 共 {1}',
  'AI 需要你決定': 'AI 需要你决定',
  '請你決定一件事': '请你决定一件事',
  '都不是？直接告訴它你要怎麼做': '都不是？直接告诉它你要怎么做',
  '自己輸入做法': '自己输入做法',
  '你選了：{0}': '你选了：{0}',
  '自動允許：{0}': '自动允许：{0}',
  '目標範圍已有 {0} 格有內容{1}，會被覆蓋。': '目标范围已有 {0} 格有内容{1}，会被覆盖。',
  '（其中 {0} 格是公式）': '（其中 {0} 格是公式）',
  '這個動作沒辦法復原': '这个操作没办法撤销',
  'AI 想要修改這份活頁簿': 'AI 想要修改这个工作簿',
  '此對話可讀取：{0}。寫入內容可能包含這些來源的資料。': '此对话可读取：{0}。写入内容可能包含这些来源的数据。',
  '想改成什麼？（選填）': '想改成什么？（选填）',
  '給 AI 的說明（選填）': '给 AI 的说明（选填）',
  '允許': '允许',
  '拒絕': '拒绝',
  '改成這樣': '改成这样',
  '執行中…': '执行中…',
  '完成': '完成',
  '詳細內容': '详细内容',
  '收合': '收起',
  '改成這樣：{0}': '改成这样：{0}',
  '已拒絕：{0}': '已拒绝：{0}',
  '正在讀取活頁簿結構…': '正在读取工作簿结构…',
  '正在讀取選取範圍…': '正在读取选定区域…',
  '正在讀取資料…': '正在读取数据…',
  '正在搜尋…': '正在搜索…',
  '等待你確認寫入…': '等待你确认写入…',
  '正在調整工作表…': '正在调整工作表…',
  '正在套用格式…': '正在应用格式…',
  '正在建立圖表…': '正在创建图表…',
  '等待你確認結構變更…': '等待你确认结构变更…',
  '正在整理表格…': '正在整理表格…',
  '等待你確認建立樞紐…': '等待你确认创建数据透视表…',
  '正在追蹤公式來源…': '正在追踪公式来源…',
  '正在調整版面…': '正在调整版面…',
  '正在執行程式…': '正在执行程序…',
  '正在寫程式…': '正在写程序…',
  '正在整理要寫入的資料…': '正在整理要写入的数据…',
  '正在設定圖表…': '正在设置图表…',
  '正在設定樞紐分析表…': '正在设置数据透视表…',
  '正在決定格式…': '正在决定格式…',
  '正在準備工作表變更…': '正在准备工作表变更…',
  '正在準備 {0}…': '正在准备 {0}…',
  '正在執行 {0}…': '正在执行 {0}…',
  '{0} 字元': '{0} 字符',
  '沒成功，AI 正在換方法': '没成功，AI 正在换方法',
  '執行失敗': '执行失败',
  '已讀入 {0} 的結果': '已读入 {0} 的结果',
  '停止（Esc）': '停止（Esc）',
  '停止 AI 工作': '停止 AI 工作',
  '送出（Enter）': '发送（Enter）',
  '送出': '发送',
  '已開新對話。之前的紀錄留在 AI Agents Office 裡，這裡從頭開始。': '已开新对话。之前的记录留在 AI Agents Office 里，这里从头开始。',
  '載入中…': '加载中…',
  '這份活頁簿還沒有對話紀錄。': '这个工作簿还没有对话记录。',
  '未命名': '未命名',
  '{0} 則提問': '{0} 条提问',
  '目前': '当前',
  '讀不到紀錄：{0}': '读不到记录：{0}',
  '今天 {0}': '今天 {0}',
  '昨天 {0}': '昨天 {0}',
  '{0} 天前': '{0} 天前',
  '過往紀錄': '过往记录',
  '（唯讀）': '（只读）',
  '回到目前對話': '回到当前对话',
  '讀不到這段紀錄：{0}': '读不到这段记录：{0}',
  '停止中…': '停止中…',
  '登入已過期，請重新登入。': '登录已过期，请重新登录。',
  '伺服器錯誤 ({0})': '服务器错误 ({0})',
  '載入 Excel 工具…': '加载 Excel 工具…',
  '已寫出 {0}': '已写出 {0}',
  '已停止。已經做完的部分留在活頁簿裡，需要的話可以用上方的復原鍵退回。': '已停止。已经做完的部分留在工作簿里，需要的话可以用上方的撤销键退回。',
  '連線失敗：{0}。確認伺服器有在跑，而且 CORS_ORIGINS 允許 {1}。': '连接失败：{0}。确认服务器有在跑，而且 CORS_ORIGINS 允许 {1}。',
  '搜尋網路': '搜索网络',
  '郵件': '邮件',
  'KM 文件': 'KM 文档',
  '這則訊息會一起送出你選取的 {0}，讓 AI 知道你指的是哪裡。點 × 可以不帶。': '这条消息会一起发送你选定的 {0}，让 AI 知道你指的是哪里。点 × 可以不带。',
  '目前的工作表：{0}': '当前的工作表：{0}',
  '自動允許，點一下切換': '自动允许，点一下切换',
  '每次確認，點一下切換': '每次确认，点一下切换',
  '目前是自動允許，點一下改回每次確認': '当前是自动允许，点一下改回每次确认',
  '目前是每次確認，點一下改成自動允許': '当前是每次确认，点一下改成自动允许',
  '{0} 尚未連結': '{0} 尚未连接',
  '這個對話可以讀取：{0}（唯讀）。': '这个对话可以读取：{0}（只读）。',
  '開啟後，這個對話可以讀取這些來源（唯讀）。': '开启后，这个对话可以读取这些来源（只读）。',
  '無法向伺服器確認資料來源（HTTP {0}）。': '无法向服务器确认数据源（HTTP {0}）。',
  '伺服器尚未套用更新，請重新啟動 server 後再試。': '服务器尚未应用更新，请重启 server 后再试。',
  '連不上伺服器（{0}）。': '连不上服务器（{0}）。',
  '（未命名活頁簿）': '（未命名工作簿）',
  '已復原 {0}': '已撤销 {0}',
  '復原失敗：{0}': '撤销失败：{0}',
  '之後的回覆會用{0}。活頁簿裡既有的資料不會被翻譯。': '之后的回复会用{0}。工作簿里已有的数据不会被翻译。',
  '之後的回覆會跟著你帳號的語言設定。': '之后的回复会跟着你账号的语言设置。',
  '已改用 {0}。{1}': '已改用 {0}。{1}',
  '已開啟{0}，這個對話可以讀取它（唯讀）。': '已开启{0}，这个对话可以读取它（只读）。',
  '已關閉{0}。': '已关闭{0}。',
  '已改成自動允許：AI 修改活頁簿時不會再詢問。刪除工作表或整列整欄仍會問，因為那沒辦法復原。': '已改成自动允许：AI 修改工作簿时不会再询问。删除工作表或整行整列仍会问，因为那没办法撤销。',
  '已改回每次確認：AI 要修改活頁簿前都會先問過你。': '已改回每次确认：AI 要修改工作簿前都会先问过你。',
  '找不到工作表「{0}」': '找不到工作表「{0}」',
  '跟隨帳號設定': '跟随账号设置',
  '用 AI Agents Office 上的語言': '用 AI Agents Office 上的语言',
  '這份活頁簿的對話紀錄': '这个工作簿的对话记录',
  '對話紀錄': '对话记录',
  '開新對話（這份活頁簿的紀錄會留著）': '开新对话（这个工作簿的记录会留着）',
  '開新對話': '开新对话',
  '復原上一步 AI 做的修改': '撤销上一步 AI 做的修改',
  '復原上一步': '撤销上一步',
  'AI 回覆語言': 'AI 回复语言',
  '登出': '退出登录',
  '目前開啟的活頁簿': '当前打开的工作簿',
  '問這份活頁簿任何事': '问这个工作簿任何事',
  '我讀得到儲存格、公式和你選取的範圍，也能在你確認後寫回去。': '我读得到单元格、公式和你选定的区域，也能在你确认后写回去。',
  '這張表在算什麼？': '这张表在算什么？',
  '幫我找出所有錯誤的公式': '帮我找出所有错误的公式',
  '選取範圍幫我做一張摘要': '选定区域帮我做一张摘要',
  '這些數字有什麼趨勢？': '这些数字有什么趋势？',
  '這則訊息不要帶選取範圍': '这条消息不要带选定区域',
  '外部資料來源': '外部数据源',
  '對這份活頁簿提問…': '对这个工作簿提问…',
  '訊息': '消息',
  'Enter 送出 · Shift+Enter 換行': 'Enter 发送 · Shift+Enter 换行',
  '這個增益集只能在 Excel 中使用。': '这个加载项只能在 Excel 中使用。',
};

const en = {
  '正在確認可用的資料來源…': 'Checking which sources are available…',
  'AI 啟動中…': 'Starting up…',
  '連線到伺服器…': 'Connecting to the server…',
  '思考中…': 'Thinking…',
  '正在想：{0}': 'Thinking: {0}',
  '{0} 個步驟 · 共 {1}': '{0} steps · {1} total',
  'AI 需要你決定': 'A decision for you',
  '請你決定一件事': 'Please choose.',
  '都不是？直接告訴它你要怎麼做': 'None of these? Say what you want instead',
  '自己輸入做法': 'Describe it yourself',
  '你選了：{0}': 'You chose: {0}',
  '自動允許：{0}': 'Auto-approved: {0}',
  '目標範圍已有 {0} 格有內容{1}，會被覆蓋。': '{0} cells in the target range already have content{1} and will be overwritten.',
  '（其中 {0} 格是公式）': ' ({0} of them formulas)',
  '這個動作沒辦法復原': 'This cannot be undone',
  'AI 想要修改這份活頁簿': 'The assistant wants to change this workbook',
  '此對話可讀取：{0}。寫入內容可能包含這些來源的資料。': 'This conversation can read: {0}. What gets written may include data from them.',
  '想改成什麼？（選填）': 'Want it done differently? (optional)',
  '給 AI 的說明（選填）': 'Instructions for the assistant (optional)',
  '允許': 'Allow',
  '拒絕': 'Decline',
  '改成這樣': 'Do it this way',
  '執行中…': 'Running…',
  '完成': 'Done',
  '詳細內容': 'Details',
  '收合': 'Hide',
  '改成這樣：{0}': 'Changed to: {0}',
  '已拒絕：{0}': 'Declined: {0}',
  '正在讀取活頁簿結構…': 'Reading the workbook structure…',
  '正在讀取選取範圍…': 'Reading your selection…',
  '正在讀取資料…': 'Reading data…',
  '正在搜尋…': 'Searching…',
  '等待你確認寫入…': 'Waiting for you to approve a write…',
  '正在調整工作表…': 'Adjusting the worksheet…',
  '正在套用格式…': 'Applying formatting…',
  '正在建立圖表…': 'Creating a chart…',
  '等待你確認結構變更…': 'Waiting for you to approve a structural change…',
  '正在整理表格…': 'Working on the table…',
  '等待你確認建立樞紐…': 'Waiting for you to approve a PivotTable…',
  '正在追蹤公式來源…': 'Tracing formula precedents…',
  '正在調整版面…': 'Adjusting the layout…',
  '正在執行程式…': 'Running…',
  '正在寫程式…': 'Working out the steps…',
  '正在整理要寫入的資料…': 'Preparing the data to write…',
  '正在設定圖表…': 'Setting up the chart…',
  '正在設定樞紐分析表…': 'Setting up the PivotTable…',
  '正在決定格式…': 'Deciding on formatting…',
  '正在準備工作表變更…': 'Preparing worksheet changes…',
  '正在準備 {0}…': 'Preparing {0}…',
  '正在執行 {0}…': 'Running {0}…',
  '{0} 字元': '{0} characters',
  '沒成功，AI 正在換方法': "That didn't work — trying another way",
  '執行失敗': 'Failed',
  '已讀入 {0} 的結果': 'Read {0} of results',
  '停止（Esc）': 'Stop (Esc)',
  '停止 AI 工作': 'Stop',
  '送出（Enter）': 'Send (Enter)',
  '送出': 'Send',
  '已開新對話。之前的紀錄留在 AI Agents Office 裡，這裡從頭開始。': 'New conversation. The earlier one is kept in AI Agents Office; this one starts fresh.',
  '載入中…': 'Loading…',
  '這份活頁簿還沒有對話紀錄。': 'No conversations for this workbook yet.',
  '未命名': 'Untitled',
  '{0} 則提問': '{0} questions',
  '目前': 'current',
  '讀不到紀錄：{0}': "Couldn't load history: {0}",
  '今天 {0}': 'Today {0}',
  '昨天 {0}': 'Yesterday {0}',
  '{0} 天前': '{0} days ago',
  '過往紀錄': 'Past conversation',
  '（唯讀）': '(read-only)',
  '回到目前對話': 'Back to the current conversation',
  '讀不到這段紀錄：{0}': "Couldn't load that conversation: {0}",
  '停止中…': 'Stopping…',
  '登入已過期，請重新登入。': 'Your session expired. Please sign in again.',
  '伺服器錯誤 ({0})': 'Server error ({0})',
  '載入 Excel 工具…': 'Loading Excel tools…',
  '已寫出 {0}': 'Wrote {0}',
  '已停止。已經做完的部分留在活頁簿裡，需要的話可以用上方的復原鍵退回。': 'Stopped. What was already done stays in the workbook — use undo above to roll it back.',
  '連線失敗：{0}。確認伺服器有在跑，而且 CORS_ORIGINS 允許 {1}。': 'Connection failed: {0}. Check the server is running and that CORS_ORIGINS allows {1}.',
  '搜尋網路': 'Web search',
  '郵件': 'Mail',
  'KM 文件': 'KM documents',
  '這則訊息會一起送出你選取的 {0}，讓 AI 知道你指的是哪裡。點 × 可以不帶。': 'This message will include your selection ({0}) so the assistant knows where you mean. Click × to leave it out.',
  '目前的工作表：{0}': 'Current sheet: {0}',
  '自動允許，點一下切換': 'Auto-approve — click to switch',
  '每次確認，點一下切換': 'Ask every time — click to switch',
  '目前是自動允許，點一下改回每次確認': 'Currently auto-approving; click to ask every time',
  '目前是每次確認，點一下改成自動允許': 'Currently asking every time; click to auto-approve',
  '{0} 尚未連結': '{0} is not connected',
  '這個對話可以讀取：{0}（唯讀）。': 'This conversation can read: {0} (read-only).',
  '開啟後，這個對話可以讀取這些來源（唯讀）。': 'Turn one on and this conversation can read it (read-only).',
  '無法向伺服器確認資料來源（HTTP {0}）。': "Couldn't check sources with the server (HTTP {0}).",
  '伺服器尚未套用更新，請重新啟動 server 後再試。': 'The server has not picked up the update yet — restart it and try again.',
  '連不上伺服器（{0}）。': "Can't reach the server ({0}).",
  '（未命名活頁簿）': '(unsaved workbook)',
  '已復原 {0}': 'Undid {0}',
  '復原失敗：{0}': 'Undo failed: {0}',
  '之後的回覆會用{0}。活頁簿裡既有的資料不會被翻譯。': 'Replies will be in {0} from now on. Data already in the workbook is left as it is.',
  '之後的回覆會跟著你帳號的語言設定。': 'Replies will follow the language set on your account.',
  '已改用 {0}。{1}': 'Now using {0}. {1}',
  '已開啟{0}，這個對話可以讀取它（唯讀）。': '{0} is on — this conversation can read it (read-only).',
  '已關閉{0}。': '{0} is off.',
  '已改成自動允許：AI 修改活頁簿時不會再詢問。刪除工作表或整列整欄仍會問，因為那沒辦法復原。': "Auto-approve is on: changes to the workbook won't ask first. Deleting sheets, rows or columns still asks, because that cannot be undone.",
  '已改回每次確認：AI 要修改活頁簿前都會先問過你。': 'Back to asking: you will be asked before anything in the workbook changes.',
  '找不到工作表「{0}」': 'No sheet named "{0}"',
  '跟隨帳號設定': 'Follow my account',
  '用 AI Agents Office 上的語言': 'Use the language set in AI Agents Office',
  '這份活頁簿的對話紀錄': 'Conversations for this workbook',
  '對話紀錄': 'History',
  '開新對話（這份活頁簿的紀錄會留著）': 'New conversation (this workbook keeps its history)',
  '開新對話': 'New conversation',
  '復原上一步 AI 做的修改': 'Undo the assistant’s last change',
  '復原上一步': 'Undo',
  'AI 回覆語言': 'Reply language',
  '登出': 'Sign out',
  '目前開啟的活頁簿': 'The workbook you have open',
  '問這份活頁簿任何事': 'Ask anything about this workbook',
  '我讀得到儲存格、公式和你選取的範圍，也能在你確認後寫回去。': 'I can read cells, formulas and whatever you select — and write back once you approve.',
  '這張表在算什麼？': 'What is this sheet calculating?',
  '幫我找出所有錯誤的公式': 'Find every broken formula',
  '選取範圍幫我做一張摘要': 'Summarise what I have selected',
  '這些數字有什麼趨勢？': 'What trends are in these numbers?',
  '這則訊息不要帶選取範圍': "Don't include my selection in this message",
  '外部資料來源': 'External sources',
  '對這份活頁簿提問…': 'Ask about this workbook…',
  '訊息': 'Message',
  'Enter 送出 · Shift+Enter 換行': 'Enter to send · Shift+Enter for a new line',
  '這個增益集只能在 Excel 中使用。': 'This add-in only works in Excel.',
};

const DICTS = { 'zh-CN': zhCN, en };

let dict = null;

/** '' (follow account) and zh-TW both mean: the source strings, untranslated. */
export function setLocale(id) {
  dict = DICTS[id] || null;
}

/**
 * Translate, then substitute {0}, {1}… positionally.
 *
 * Arguments are interpolated after lookup, never before — a sheet name that
 * happens to contain {0} is data, and data must not be able to reach into the
 * template. Missing entries fall through to the Traditional Chinese source.
 */
export function t(zh, ...args) {
  const s = (dict && dict[zh]) || zh;
  return args.length ? s.replace(/\{(\d+)\}/g, (m, i) => (args[i] === undefined ? m : String(args[i]))) : s;
}

/**
 * Re-render the markup's own text.
 *
 * data-t sets textContent; the -ph/-tip/-al/-title variants set placeholder,
 * data-tip, aria-label and title. Attributes rather than a lookup table in JS,
 * so a string added to the HTML is translated by the same edit that adds it —
 * a separate list is a list that goes stale.
 */
export function applyStaticText(root = document) {
  const map = [
    ['data-t', (el, v) => { el.textContent = v; }],
    ['data-t-ph', (el, v) => { el.placeholder = v; }],
    ['data-t-tip', (el, v) => el.setAttribute('data-tip', v)],
    ['data-t-al', (el, v) => el.setAttribute('aria-label', v)],
    ['data-t-title', (el, v) => { el.title = v; }],
  ];
  for (const [attr, set] of map) {
    for (const el of root.querySelectorAll(`[${attr}]`)) set(el, t(el.getAttribute(attr)));
  }
}
