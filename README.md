# JVision for Office

在 Excel 側邊欄直接對「當下開啟的那份活頁簿」提問 —— AI 讀得到儲存格、公式、選取範圍，
也能在你按下確認後寫回去。對標 Claude for Excel，但後端是我們自己的
[ai-agents-office](../ai-agents-office)，資料流全程在自己的機器上。

**目前已完成 Excel 與 Word。** 兩者共用同一份外殼（`src/taskpane/app.js`），
主應用程式專屬的那一層在 `src/taskpane/hosts/`，由 `Office.onReady` 動態載入——
Word 不會載入 Excel 那份。PowerPoint 還沒做；
為什麼這樣放、以及更版怎麼互不影響，見 [docs/multi-host.md](docs/multi-host.md)。
下面的說明以 Excel 為例。

## 這跟「產生一個 Excel 檔給你下載」差在哪

差在**活的**。模型手上沒有你的檔案，它是透過工具呼叫，請你電腦上的 Excel 幫它讀某個範圍、
寫某幾格。所以公式關聯不會壞、樞紐不會掉、格式還在，而且它引用的每個數字都能點回原始儲存格。

## 架構

```
Excel 桌面版（WebView2）
  任務窗格 = 一個 HTTPS 網頁
      │  ① POST /api/excel/chat  ─────────────> SSE（模型輸出 + 工具請求）
      │  ⑥ POST /api/excel/tool-result ───────> Office.js 執行結果
      ▼
ai-agents-office / Express
      │  excelBridge.ts   runId → SSE 連線 + 等待中的呼叫
      │  ② spawnClaude(..., mcpExcelRunToken)
      │  ⑤ /internal/excel/tool   ← 只收 127.0.0.1 + run token
      ▼
Claude CLI ──stdio──> excelMcp.ts（獨立 subprocess）
      ③ 模型呼叫 excel_read_range
      ④ POST 回 ⑤ ──────────────────────────────┘
```

為什麼要繞這麼一圈：Claude CLI 把每個 MCP server 都 spawn 成獨立 process，那個 process
摸不到 Express 記憶體裡的 SSE 連線，只能用 HTTP loopback 繞回來。這是既有架構決定的，
不是設計上的偏好。詳見 [excelBridge.ts](../ai-agents-office/server/src/services/excelBridge.ts) 的檔頭。

## 目前實作了什麼

| 工具 | 能做什麼 | 要確認 |
|---|---|---|
| `excel_get_overview` | 工作表清單、大小、前幾列預覽 | |
| `excel_get_selection` | 使用者當下選取的範圍 | |
| `excel_read_range` | 讀值／讀公式／兩者，回 CSV | |
| `excel_search` | 找含某段文字的儲存格 | |
| `excel_read_attachment` | 看使用者貼進對話框的圖片 | |
| `excel_read_image` | 看工作表上貼著的圖片內容（使用者問到才讀） | |
| `excel_write_range` | 寫入值或公式 | ✔ |
| `excel_clean_data` | 去重、統一寫法、補空白、拆併欄位 | ✔ |
| `excel_manage_sheet` | 新增／更名／刪除／切換／搬移工作表 | 刪除、更名 |
| `excel_format_range` | 數字格式、粗體、字色底色、對齊、框線、欄寬列高、合併 | |
| `excel_create_chart` | 直條／橫條／折線／圓餅／散佈／區域／雷達圖 | |
| `excel_structure_op` | 插入刪除整列整欄、清除範圍 | ✔ |
| `excel_table_ops` | 排序、篩選、轉表格、凍結窗格、條件式格式 | 排序 |
| `excel_create_pivot` | 建立真的樞紐分析表（可自己拖欄位的活物件） | ✔ |
| `excel_trace_precedents` | 追一格公式是從哪些儲存格算出來的 | |
| `excel_sheet_style` | 關格線、隱藏欄列標題、分頁顏色、預設欄寬 | |
| `excel_compare_source` | 逐格比對兩份資料——同活頁簿的兩張工作表，或跟使用者上傳的檔案——結果寫成一張比對報表 | 覆寫既有工作表時 |

儲存格引用 `[[工作表!C42]]` 會渲染成可點的連結，點了跳到那一格。
寫入前會存快照，「復原上一步」拿得回來。

圖片可以直接貼進對話框（或拖進去）——最常見的用法是把一張報表截圖丟進來，
請它打成 Excel 表格。圖片留在增益集裡，模型要看才會呼叫 `excel_read_attachment` 取走；
沒問到的圖不會離開這台電腦。工作表上「貼著」的圖片則只在使用者提到它時才讀，
因為圖片是最難用肉眼檢查的夾帶管道。

格式方面另外支援：字體、框線顏色／粗細／樣式、縮排、內建表格樣式、圖表圖例位置與資料標籤。

**還沒有的：** 巨集／VBA —— Office.js 本身就不支援，沒有繞道。

### 關於「美化」

工具只是旋鈕，真正決定成品好不好看的是 system prompt 裡的「排版與美感」那一節
（[excelContext.ts](../ai-agents-office/server/src/services/excelContext.ts)）。
那裡寫的是可執行的規則而不是形容詞：關格線、一個主色、字級層次、
用列高做留白、數字一定要設格式、收尾要 autofit + 凍結標題列，
以及儀表板／行事曆／報表三種成品的具體步驟。

改「它做出來的東西好不好看」要去改那一節，不是加更多屬性。

### 需要較新的 Excel 版本

| 功能 | 需要 | 版本不夠時 |
|---|---|---|
| 樞紐分析表 | ExcelApi **1.12** | 工具不會出現在模型眼前，改用算好的彙總表 |
| 追蹤公式（直接參照） | ExcelApi **1.12** | 退回解析公式字串，只找得到直接參照 |
| 追蹤公式（追到底、跨表） | ExcelApi **1.14** | 退回只追直接參照 |

版本偵測用 `Office.context.requirements.isSetSupported()`，結果會併進送給伺服器的
`clientTools`——所以舊版 Excel 上模型根本看不到樞紐工具，不會呼叫了才失敗。

### 確認閘門的界線

會**弄丟資料**的操作才跳確認：覆寫儲存格、插入刪除列欄、刪除或更名工作表、排序。
格式、圖表、新增工作表、篩選、凍結窗格直接執行。

這條線是刻意畫的——每個裝飾性動作都要按一次「允許」，人就會養成不看內容直接按的習慣，
那正是被注入的破壞性指令混過去的方式。

## 先決條件

- **Excel**：Microsoft 365 桌面版，組建 16.0.13127.20296 以上。
  Excel 2016／2019 買斷版**不支援**（manifest 需要 SharedRuntime），iPad／Android 也不支援。
- **Node.js 18+**
- **ai-agents-office 的 server 跑得起來**（MySQL、Claude CLI 已登入）

## 啟動（開發）

**① 安裝本機 HTTPS 憑證（只需一次）**

Office 增益集強制走 HTTPS，localhost 也不例外。

```bash
npx office-addin-dev-certs install
```

會跳一次 Windows 憑證安裝確認。這個步驟刻意不放進 `npm run dev` 自動做 ——
它會往系統信任區塞一張根憑證，該由你自己按下同意。

**② 起任務窗格的靜態伺服器**

```bash
npm install
npm run dev          # https://localhost:3443
```

沒有打包步驟。任務窗格是原生 ES modules，WebView2 直接吃，改完存檔重新整理窗格就生效。

唯一的例外是 `src/taskpane/vendor/mermaid.min.js`（2.9 MB，MIT，從
`ai-agents-office/client/node_modules/mermaid/dist/` 複製過來）。Word 沒有圖表 API，
伺服器也沒有 mermaid → 圖片的管線，而窗格本身就是瀏覽器，所以在這裡渲染
（mermaid → SVG → canvas → PNG）零伺服器來回。用 UMD 單檔而不是 ESM 版，
因為 ESM 那份只有 26 KB，它會再去載 chunks。**只在第一次要畫圖時才載入**，
不影響窗格開啟速度。要升級就重新複製一次。

**③ 起後端**

```bash
cd ../ai-agents-office
pnpm run dev:server
```

`.env` 的 `CORS_ORIGINS` 必須含 `https://localhost:3443`（已加好）。少了這行，
任務窗格的每個請求都會被瀏覽器擋掉，而且錯誤訊息不會告訴你是 CORS。

**④ 安裝到 Excel**

雙擊 [installer/安裝.cmd](installer/安裝.cmd)，然後**完全關閉 Excel 再重新開啟**。

常用分頁最右邊會出現「JVision 助手」。之後改網頁程式碼只要重新整理窗格；
只有改了 `manifests/` 裡的 manifest 才需要重跑一次安裝。

只想裝其中一個主應用程式的話：

```powershell
.\installer\install.ps1 -Hosts word
```

要移除就雙擊 [installer/解除安裝.cmd](installer/解除安裝.cmd)，或到 Windows 的
「應用程式和功能」找 JVision for Excel。

**⑤ 設定 API 位址**

```bash
cp .env.example .env      # 首次
# 編輯 .env 裡的 API_BASE，對應 ai-agents-office 的 PORT
```

`npm run dev` 啟動時會把 `.env` 轉成 `config.json`，任務窗格開場讀它。
**使用者看不到也不用輸入伺服器位址**——那是部署設定，不是使用者偏好，
打錯字造成的登入失敗看起來會跟密碼錯誤一模一樣。

連不上設定的伺服器時，登入畫面才會把「伺服器」欄位放出來讓人當場修正，
輸入的值存成明確覆寫（跟設定值分開，所以之後改 `config.json` 不會被舊快取蓋掉）。

正式部署時直接改主機上的 `config.json` 即可，不用重新產生或重新部署其他東西。

**⑥ 登入**

用 AD 工號或電子信箱登入（見下方「登入方式」），拿到的 JWT 存在窗格自己的 localStorage。

## 要怎麼發給使用者

三條路，選哪一條看你在哪個階段：

| 方式 | 使用者要做什麼 | 需要 IT | 適合 |
|---|---|---|---|
| **M365 系統管理中心集中部署** | **什麼都不用做**，按鈕自己出現 | 要，全域管理員 | 正式全公司推 |
| **installer/安裝.cmd** | 雙擊一次，重開 Excel | 不用 | 試行、少數人、IT 還沒點頭前 |
| 共用資料夾受信任目錄 | 手動設四層對話框 | 不用 | 只建議開發時用 |

### 集中部署（正式的答案）

Microsoft 365 系統管理中心 → 設定 → 整合式應用程式 → 增益集 → 上傳自訂應用程式 →
上傳 `manifests/excel.xml`、`manifests/word.xml` 和 `manifests/powerpoint.xml`（**三份分開上傳**，它們是三個獨立的增益集）
→ 指派給使用者或群組。使用者端**完全不用動作**，下次開 Excel
按鈕就在那裡，IT 也能隨時收回。這是 Claude for Excel 給企業客戶的方式，也應該是我們的終局。

前提是 manifest 裡的網址要是使用者連得到的內網主機，不能是 localhost。

### 安裝程式（試行階段用這個）

[installer/](installer/) 底下那兩個 `.cmd` 就是給一般使用者雙擊的。它做的事：

1. 把選定的每份 manifest 複製到 `%LOCALAPPDATA%\JVision for Office\`
   —— 放在下載資料夾的檔案遲早會被刪掉，之後 Excel 就會載入失敗
2. 在 `HKCU\SOFTWARE\Microsoft\Office\16.0\Wef\Developer` 寫一個字串值
   （名稱＝manifest 的 `<Id>`，資料＝manifest 路徑），Excel 啟動時就會讀到它
3. 註冊一筆解除安裝資訊，讓它出現在 Windows 的「應用程式和功能」裡

**全部在 HKCU 和 %LOCALAPPDATA%，不需要系統管理員權限**，也不影響同一台電腦的其他使用者。
第 2 步就是微軟自己的 `office-addin-dev-settings` 工具在做的事，只是我們把它包成雙擊即可。

打包給不同環境用 `-TaskpaneUrl`，安裝程式會在複製時就把網址換掉：

```powershell
.\install.ps1 -TaskpaneUrl https://jvision.panjit.com.tw
```

所以你只要維護一份 manifest，不必為每個環境各存一份。

### 單一執行檔（發給一般使用者用這個）

[installer/JVisionSetup.exe](installer/JVisionSetup.exe) 把三份 manifest **嵌進 exe 裡**，
所以發出去的就是一個檔案，沒有「記得連 manifests 資料夾一起包」這種事。
雙擊會開一個對話框，可以勾要裝哪幾個主應用程式、改伺服器網址。

IT 派送：

```powershell
JVisionSetup.exe /silent                          # 三個都裝
JVisionSetup.exe /silent /hosts:word              # 只裝 Word
JVisionSetup.exe /silent /hosts:excel,ppt         # 不裝 Word
JVisionSetup.exe /silent /url:https://jvision.panjit.com.tw
JVisionSetup.exe /silent /uninstall
```

離開代碼：0 成功、1 失敗、2 這台電腦沒有 Office 16.0、3 使用者取消。
「2」不是錯誤，是「這台機器本來就不是目標」，派送系統可以把它對到 not-applicable。

**改過 manifest 之後要重新建置**，否則 exe 裡還是舊的版本號和網址：

```bash
npm run installer:build          # 用 .NET Framework 內建的 csc.exe，不需要裝 SDK
```

建置時會檢查三份 manifest 的版本號一致——不一致的話使用者會看到一個版本、裝到另一個。

**還沒簽章。** 沒簽章的 exe 會被 SmartScreen 攜，企業環境的防毒也常常直接隔離；
要大量派送的話，`npm run installer:build -- -Sign -CertThumbprint <指紋>` 可以帶公司憑證簽。

不想用 exe 的話，把 `installer\` 和 `manifests\` 壓成 zip 也一樣可以，
使用者雙擊 `安裝.cmd`。

## 除錯

- **先在 Excel 網頁版開發**。同一份 manifest 可以側載到 Excel for the web，F12 就有完整
  DevTools，迭代速度差一個數量級。桌面版只用來做最後驗證。
- 桌面版要接 DevTools：`npx office-addin-debugging start manifests/excel.xml desktop`
  （Word 就換成 `manifests/word.xml`）。
- **MCP 那一段看不到 stderr**：Claude CLI 會吃掉 MCP 的 stderr。開發模式下
  `excelMcp` 會把每次工具呼叫寫進 `ai-agents-office/workspace/excel-mcp-debug.log`，
  agent 說「我沒有 Excel 工具」的時候第一個看這裡。
- 模型載不到工具通常是 `ToolSearch` 被關掉了 —— MCP 工具在 CLI 裡是延遲載入的，
  少了 ToolSearch 就永遠拿不到。

### 改了東西但 Excel 沒反應？先確認是哪一種快取

這個平台有三層各自獨立的快取，症狀很像但解法完全不同：

| 症狀 | 原因 | 解法 |
|---|---|---|
| 改了 `taskpane.*` 但窗格沒變 | WebView 的頁面快取 | 重新整理窗格 |
| 改了 `app.js` / `hosts/*.js` 但行為沒變 | **SharedRuntime 讓 JS 常駐**，關窗格不會重載 | **完全關閉 Excel 再開** |
| 換了圖示但功能區還是舊的／顯示通用方塊 | Office 依「增益集 ID + manifest 版本」判斷要不要重讀，版本沒動就不重抓 | `npm run bump` → 重跑安裝 → 清快取 → 開 Excel |

第三種最容易誤判成「圖檔壞了」。先用 `Invoke-WebRequest https://localhost:3443/assets/icon-32.png`
確認圖是正常服務出去的——是的話就純粹是快取問題。

**換圖示的完整步驟**（順序不能顛倒）：

```powershell
npm run icons          # 從 public/logo2.png 重新產生
npm run bump           # manifest 版本 +1 —— 沒有這步，後面全部白做
# 完全關閉 Excel
installer\安裝.cmd      # 把新版本的 manifest 複製到 %LOCALAPPDATA%
npm run clear-cache    # 清掉舊圖示（會順便檢查伺服器有沒有在跑）
# 確認 npm run dev 有在跑，再開 Excel
```

**`npm run dev` 必須在開 Excel 之前就跑著。** 清完快取後的第一次 Excel 啟動，
是 Office 唯一一次去抓功能區圖示的機會；抓不到就退回通用立方體，而且在下次
改版本號之前都不會再試。症狀跟「圖檔壞掉」一模一樣，但其實只是伺服器沒開。
（正式部署時圖示放在常駐主機上，不會有這個問題——這純粹是本機開發的坑。）

**版本號那步是關鍵。** Office 不是用 URL 或檔案內容判斷圖示有沒有變，是用
「增益集 ID + manifest 版本」。版本一樣它就完全不重新處理這個增益集——
即使你已經把圖示快取刪光、即使那些網址回應完全正常。症狀就是功能區顯示一個通用立方體。

### 已知問題：開發側載時功能區圖示顯示為通用立方體

用 `installer\安裝.cmd`（登錄檔側載）安裝時，功能區和「增益集」對話框會顯示 Office 的
通用立方體，而不是我們的 logo。**側邊欄裡的 logo 正常**，功能完全不受影響。

這件事查過一輪，以下都**不是**原因，不要再重複查：

| 假設 | 驗證方式 | 結果 |
|---|---|---|
| 圖檔格式不對 | 逐位元組比對 IHDR | 與 Claude for Excel 的圖示完全相同 |
| `pHYs` chunk | 移除後重測 | 無效 |
| 透明通道 | Claude 的也是 RGBA | 不是原因 |
| 抓不到圖 | dev server 請求記錄 | 三個尺寸都 HTTP 200，位元組數正確 |
| 快取沒清 | 清空整個 `Wef`（9322 檔） | 無效 |
| 版本沒 bump | 每次都確認已安裝副本同步 | 無效 |
| manifest 結構 | 與 Claude 的 manifest 逐元素比對 | 完全相同 |
| 重複註冊 | 檢查 `Wef\Developer` | 只有一筆，Id 與路徑都正確 |
| **圖檔本身** | **把 Claude 的圖檔位元組放到我們的網址下** | **仍是立方體 → 與圖檔無關** |

最後那一列是決定性的：用別人能正常顯示的圖檔，在我們的註冊方式下一樣顯示不出來。

剩下的唯一差異是**載入管道**——Claude 在「我的增益集」（正式部署），我們在「開發人員增益集」
（登錄檔側載）。所以推論是 Office 不對登錄檔側載的開發用增益集套用自訂功能區圖示。

**這對正式部署沒有影響**：M365 系統管理中心的集中部署跟 Claude 走同一條管道。
要在部署前確認的話，最快的驗證是把同一份 manifest 側載到 Excel 網頁版
（增益集 → 上傳我的增益集），那是完全不同的管道，不需要系統管理員權限。

## 安全模型

- **寫入一定要人按確認。** 後端跑 CLI 時用 `--permission-mode dontAsk`（伺服器上沒有人可以按
  Allow），所以確認閘門實作在任務窗格裡。確認視窗會明講「這會蓋掉 N 格既有資料，其中 M 格是公式」。
- **寫入前一定會快照。** Office.js 的寫入不能倚賴 Excel 原生的 Ctrl+Z，所以每次寫入前會把目標範圍的
  公式和數字格式存下來，「復原上一步」按鈕拿得回來。`excel_run_script` 事前不知道自己會碰哪裡，
  所以改成整份活頁簿：每張工作表使用範圍的公式與數字格式。復原時會把寫超出舊範圍的內容清掉、
  移除腳本新增的工作表、重建它刪掉的工作表（判斷依據是腳本執行前後兩份工作表清單的差集，
  所以事後使用者自己開的工作表不會被動到）。兩者都只還原**內容**，不還原填色框線字型；
  活頁簿超過 20 萬格就跳過快照，並在對話裡明講這次拿不回來。

- **腳本能做什麼由執行時決定，不是由它長什麼樣決定。** 伺服器用 regex 掃原始碼判斷腳本會不會寫入，
  那是啟發式；真正的控制在窗格這邊——每個 Office 物件都包了 Proxy，任何屬性 assignment、
  以及任何不在唯讀清單（`get…`、`load`、`sync` 等）裡的方法呼叫，在沒有取得寫入許可時就地擋下。
  規則寫成「唯讀的有哪些」而不是「會寫入的有哪些」，這樣 Office.js 加新 API 時它不會過期。
- **活頁簿內容視為不可信輸入。** 外來範本可以在儲存格裡藏針對 agent 的指令。system prompt 明確
  要求模型把那些當資料而不是命令，加上寫入閘門，構成兩道防線。
- **身分就是 token。** 每次 run 產生一次性的 bridge token，經 env 注入 MCP subprocess，
  對應到唯一一條 SSE 連線。模型無法指定「哪個使用者」「哪份活頁簿」—— 它只能碰到啟動它的那一個。
- **內部端點只收 loopback。** `/internal/excel/tool` 沒有 JWT，靠的是「必須來自 127.0.0.1」
  加上「必須帶有效 run token」兩個條件。

## 正式部署要改什麼

1. 任務窗格改由正式的 HTTPS 主機提供（憑證要是內部 CA 或公開 CA 簽的）
2. 網址**不要手改 manifest** —— 用 `install.ps1 -TaskpaneUrl https://你的主機`，
   或先產生一份改好網址的 manifest 再上傳到系統管理中心
3. 正式站和開發站要並存的話，正式站那份 manifest 換一個新的 `<Id>` GUID，
   否則兩份會互相覆蓋
4. 後端 `CORS_ORIGINS` 加上那個主機
5. 部署走 Microsoft 365 系統管理中心（見上面「要怎麼發給使用者」）。
   內部工具不需要上 AppSource，改版也不用審核。
6. 圖示已是正式 logo。改了 `public/logo2.png` 之後跑 `npm run icons` 重新產生即可
