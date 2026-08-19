# JVision for Excel

在 Excel 側邊欄直接對「當下開啟的那份活頁簿」提問 —— AI 讀得到儲存格、公式、選取範圍，
也能在你按下確認後寫回去。對標 Claude for Excel，但後端是我們自己的
[ai-agents-office](../ai-agents-office)，資料流全程在自己的機器上。

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
| `excel_write_range` | 寫入值或公式 | ✔ |
| `excel_manage_sheet` | 新增／更名／刪除／切換／搬移工作表 | 刪除、更名 |
| `excel_format_range` | 數字格式、粗體、字色底色、對齊、框線、欄寬列高、合併 | |
| `excel_create_chart` | 直條／橫條／折線／圓餅／散佈／區域／雷達圖 | |
| `excel_structure_op` | 插入刪除整列整欄、清除範圍 | ✔ |
| `excel_table_ops` | 排序、篩選、轉表格、凍結窗格、條件式格式 | 排序 |

儲存格引用 `[[工作表!C42]]` 會渲染成可點的連結，點了跳到那一格。
寫入前會存快照，「復原上一步」拿得回來。

**還沒有的：** 追蹤公式前導參照（`trace_precedents`）、樞紐分析表、巨集／VBA。
前兩個是之後要補的，巨集是 Office.js 本身就不支援。

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
只有改了 `manifest.xml` 才需要重跑一次安裝。

要移除就雙擊 [installer/解除安裝.cmd](installer/解除安裝.cmd)，或到 Windows 的
「應用程式和功能」找 JVision for Excel。

**⑤ 登入**

側邊欄會問伺服器位址（預設 `http://localhost:12054`，對應 ai-agents-office 的 `PORT`）
和你的帳號密碼，拿到的 JWT 存在窗格自己的 localStorage。

## 要怎麼發給使用者

三條路，選哪一條看你在哪個階段：

| 方式 | 使用者要做什麼 | 需要 IT | 適合 |
|---|---|---|---|
| **M365 系統管理中心集中部署** | **什麼都不用做**，按鈕自己出現 | 要，全域管理員 | 正式全公司推 |
| **installer/安裝.cmd** | 雙擊一次，重開 Excel | 不用 | 試行、少數人、IT 還沒點頭前 |
| 共用資料夾受信任目錄 | 手動設四層對話框 | 不用 | 只建議開發時用 |

### 集中部署（正式的答案）

Microsoft 365 系統管理中心 → 設定 → 整合式應用程式 → 增益集 → 上傳自訂應用程式 →
上傳 `manifest.xml` → 指派給使用者或群組。使用者端**完全不用動作**，下次開 Excel
按鈕就在那裡，IT 也能隨時收回。這是 Claude for Excel 給企業客戶的方式，也應該是我們的終局。

前提是 manifest 裡的網址要是使用者連得到的內網主機，不能是 localhost。

### 安裝程式（試行階段用這個）

[installer/](installer/) 底下那兩個 `.cmd` 就是給一般使用者雙擊的。它做的事：

1. 把 `manifest.xml` 複製到 `%LOCALAPPDATA%\JVision for Excel\`
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

發佈時把 `installer\` 和 `manifest.xml` 一起壓成 zip 給使用者就行。想更像「一個程式」的話，
可以用 Inno Setup 把這兩個腳本包成單一個 `.exe` 並簽章——那一層還沒做，需要的話再加。

## 除錯

- **先在 Excel 網頁版開發**。同一份 manifest 可以側載到 Excel for the web，F12 就有完整
  DevTools，迭代速度差一個數量級。桌面版只用來做最後驗證。
- 桌面版要接 DevTools：`npx office-addin-debugging start manifest.xml desktop`。
- **MCP 那一段看不到 stderr**：Claude CLI 會吃掉 MCP 的 stderr。開發模式下
  `excelMcp` 會把每次工具呼叫寫進 `ai-agents-office/workspace/excel-mcp-debug.log`，
  agent 說「我沒有 Excel 工具」的時候第一個看這裡。
- 模型載不到工具通常是 `ToolSearch` 被關掉了 —— MCP 工具在 CLI 裡是延遲載入的，
  少了 ToolSearch 就永遠拿不到。

## 安全模型

- **寫入一定要人按確認。** 後端跑 CLI 時用 `--permission-mode dontAsk`（伺服器上沒有人可以按
  Allow），所以確認閘門實作在任務窗格裡。確認視窗會明講「這會蓋掉 N 格既有資料，其中 M 格是公式」。
- **寫入前一定會快照。** Office.js 的寫入不能倚賴 Excel 原生的 Ctrl+Z，所以每次寫入前會把目標範圍的
  公式和數字格式存下來，「復原上一步」按鈕拿得回來。
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
6. `assets/icon-*.png` 換成真的圖（現在是 `npm run icons` 產的純色占位圖）
