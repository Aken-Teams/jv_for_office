# 一個專案，三個 Office 主應用程式

Excel、PowerPoint、Word 三個增益集要怎麼放，以及為什麼這樣放。
決定於 2026-08-21，在 Excel 版本 1.0.49 之後、PowerPoint 版動工之前。

---

## 決策摘要

| 項目 | 決定 | 一句話理由 |
|---|---|---|
| 專案 | **一個**（`jv_for_office`） | 外殼有 3,300 行跟主應用程式無關，複製三份一定會漂移 |
| 外殼程式碼 | **共用** | 修一次，三個主應用程式都好 |
| 主應用程式專屬程式碼 | **各自獨立檔案，動態載入** | Word 從頭到尾不會執行 Excel 的那份 |
| manifest | **三份**，各管一個主應用程式 | 可以只裝 Word、可以單獨移除 |
| 窗格網址 | **帶版號** | 各自決定何時升版、可以單獨退版 |
| 安裝程式 | **一顆**，`/hosts:` 決定裝哪些 | IT 派送一次，範圍自己挑 |

---

## 為什麼是一個專案

先量過再決定的，不是憑感覺。Excel 版做到 1.0.49 時的行數：

| 檔案 | 行數 | 換一個主應用程式要重寫嗎 |
|---|---|---|
| `app.js` | 2,014 | 幾乎不用 |
| `taskpane.css` | 1,018 | 不用 |
| `i18n.js` | 334 | 只有少數字彙（活頁簿／工作表／儲存格） |
| `excelTools.js` | 1,494 | **整份** |

關鍵的一個數字：**`app.js` 裡碰到 Excel API 的地方只有一處**，就是 `Office.onReady` 裡
的主應用程式判斷。登入、AD 首次設定、SSE 串流、進度面板、確認卡、復原、對話紀錄、
模型選單、資料來源、語系、停止鍵——這兩千行沒有一行知道自己在 Excel 裡。

接縫本來就存在，就是 `excelTools.js` 對外匯出的那十幾個函式。

分成三個專案的代價是這 3,366 行變成三份。它們會漂移，不是可能而是一定。這個專案
2026-08 那幾天修掉的每一件事——變數遮蔽讓整個窗格掛掉、腳本寫入的攔截層、Ctrl+Z、
帳號被刪除時顯示錯誤原因——每一件都會變成要做三次，而且第三次一定會漏。

### 什麼情況該拆成三個專案

- 三個增益集由不同的人維護
- 發版節奏必須完全分開，且無法用帶版號的網址解決
- 其中一個要獨立授權或獨立賣

以上都不成立時不要拆。**從合併拆開很容易，從三份合起來很難**——真要合的時候，三份
程式碼早就不一樣了。

---

## 目錄結構

```
jv_for_office/
  manifests/
    excel.xml          三份由同一個範本產生，只差主應用程式名稱和 Id
    powerpoint.xml
    word.xml
  src/taskpane/
    app.js             外殼。不知道自己在哪個主應用程式
    i18n.js
    taskpane.css
    taskpane.html
    hosts/
      index.js         依 Office.context.host 動態載入下面其中一個
      excel.js         原 excelTools.js
      powerpoint.js
      word.js
  installer/           一顆 exe，/hosts: 決定裝哪些
  test/
  docs/
```

伺服器端（`ai-agents-office`）對應：

```
routes/office.ts       原 excel.ts，帶 host 參數
services/
  toolSpec/excel.ts | powerpoint.ts | word.ts
  context/excel.ts  | powerpoint.ts | word.ts
  securityRules.ts   三個共用
```

資料庫 `conversations.workbook_key` 要改名 `document_key`——Word 沒有活頁簿。

---

## 三份 manifest，不是一份

一份 manifest 可以同時宣告三個主應用程式：

```xml
<Hosts>
  <Host Name="Workbook" />
  <Host Name="Document" />
  <Host Name="Presentation" />
</Hosts>
```

**但不要這樣做。** 那樣裝了就是三個一起來，只想用 Word 的人也會在 Excel 看到按鈕，
而且沒辦法單獨移除其中一個。

三份 manifest 各自一個 Id、一個註冊表項目，三份都指向同一個 `taskpane.html`，
`Office.onReady` 自己判斷在哪。於是：

```
JVisionSetup.exe /silent                    → 三個都裝
JVisionSetup.exe /silent /hosts:word        → 只有 Word
JVisionSetup.exe /silent /hosts:excel,ppt   → 不裝 Word
```

IT 可以先只推 Excel，PowerPoint 做好了再推，也可以單獨拿掉某一個。

版本號由建置腳本從同一個來源寫進三份，不會有人手動改到不一致。

---

## 帶版號的網址

這是「Excel 更版會不會影響 Word」的真正答案，而且**跟專案怎麼切完全無關**。

窗格是從伺服器載入的。所以就算切成三個專案、只要部署到同一台，一更新三個主應用程式
立刻都換。切專案解決不了這件事。

解法是讓每份 manifest 指向自己的版號路徑：

```
Excel  manifest → /v1.0.50/taskpane.html
Word   manifest → /v1.0.49/taskpane.html   ← 沒動
```

部署新版只是多一個資料夾，**沒有人會被移動**。改哪份 manifest 的網址，哪個主應用程式
才升。Excel 升上去發現有問題，把網址改回 `/v1.0.49/` 就退版了，Word 全程沒感覺。

> 2026-08 當時還沒有正式部署腳本（`tools/dev-server.mjs` 直接讀檔案樹），所以這件事
> 要在寫部署腳本的時候一起做。上線之後再改會麻煩很多。

---

## 風險與處理

| 風險 | 處理 |
|---|---|
| Excel 的 bug 跑到 Word | `hosts/*.js` 動態載入，Word 不會載入 Excel 那份。靠結構，不靠紀律 |
| **共用外殼的 bug 影響三個** | 這是真的，不粉飾。用測試攔截；換來的是修一次三個都好 |
| Excel 更版影響 Word | 帶版號的網址，各自決定何時升 |
| 想單獨移除某個主應用程式 | 三份 manifest |
| 發版節奏綁在一起 | 帶版號的網址已經解耦；真的不夠再拆專案 |

共用外殼那一條沒有辦法消除，只能降低。降低的方式是 `npm test` 要涵蓋外殼，不是只有
主應用程式專屬的部分。

---

## 動工順序

1. **先做主應用程式抽象化重構，只用 Excel 驗證。** 不要邊寫 PowerPoint 邊抽——那樣
   會同時在 debug 兩個主應用程式和一個還沒定型的抽象層。重構前後 Excel 的行為必須
   完全一樣。
2. `excelTools.js` → `hosts/excel.js`，加上 `hosts/index.js` 的動態載入。
3. 把 `app.js` 裡剩下的 Excel 假設抽出去：`getWorkbookKey`、`getActiveSheetName`、
   `getSelectionRef`、`inspectWrite`、上方的工作表列，以及 i18n 字典裡的
   「活頁簿」「工作表」「儲存格」。
4. manifest 拆成三份 + 建置腳本產生。
5. 安裝程式加上 `/hosts:`。
6. **PowerPoint 先於 Word**（理由見下）。

---

## 為什麼 PowerPoint 排在 Word 前面

跟架構無關，是定位問題。

`ai-agents-office` 網頁版本來就會產出文件。Word 版增益集和它會撞：「幫我改這段」在
網頁上貼一貼也做得到。簡報的版面操作則是網頁版最難取代的——投影片的位置、對齊、
母片、動畫，都必須在那份檔案裡動。價值差距最明顯的是 PowerPoint。

Word 該先想清楚它跟網頁版的分工再動手。
