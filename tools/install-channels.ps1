<#
  把 manifests\<頻道>\ 底下的檔案裝到 %LOCALAPPDATA%，然後清圖示快取。

  ── 為什麼需要這支腳本 ──
  步驟不多，但**順序錯了就沒有效果**，而且失敗的樣子跟「沒做」一模一樣：清單裡
  一顆通用六邊形，功能區一顆通用六邊形，沒有任何錯誤訊息。這個順序在開發過程中
  被弄錯兩次，兩次都花了好幾輪才發現錯在順序而不是錯在檔案。

  正確順序，每一步都有理由：

    1. 關閉 Office      ── Office 開著的時候快取檔案被鎖住，清掉的是清得動的那些，
                           剩下的留著，於是「清過了」和「清乾淨了」不是同一件事。
    2. 複製 manifest    ── 要先讓新的版本號就位。
    3. 清快取           ── 在複製之後。反過來做的話，清掉的是舊版本留下的有效快取，
                           然後沒有任何東西去補，Office 下次啟動又記一次失敗。
    4. 確認圖示抓得到   ── Office 第一次抓失敗就會記住，而且在下次改版本號之前
                           不會再試。所以這一步寧可先擋下來，也不要讓它去記一筆失敗。

  用法：
    powershell -ExecutionPolicy Bypass -File tools\install-channels.ps1
    powershell -ExecutionPolicy Bypass -File tools\install-channels.ps1 -Channels jv,test
#>
param(
  # 預設三個都裝。只裝其中幾個就傳 -Channels jv,test。
  [string[]]$Channels = @('jv', 'test', 'localtest')
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $PSScriptRoot

# 頻道 → %LOCALAPPDATA% 底下的資料夾名。安裝程式用的是中文後綴，這裡要對得上，
# 否則會裝到一個 Office 沒有登記的位置，而且一樣沒有錯誤訊息。
$folders = @{
  'jv'        = 'JVision for Office（JV）'
  'test'      = 'JVision for Office（測試）'
  'localtest' = 'JVision for Office（本地測試）'
}

Write-Host "=== 安裝 manifest 並清快取 ===" -ForegroundColor Cyan
Write-Host ""

# ── 1. Office 必須完全關閉 ──────────────────────────────────────────────────
$apps = Get-Process WINWORD, EXCEL, POWERPNT -ErrorAction SilentlyContinue
if ($apps) {
  Write-Host "Office 還開著，先關掉再跑一次：" -ForegroundColor Yellow
  $apps | Select-Object ProcessName, Id | Format-Table -AutoSize
  Write-Host "  （不自動關閉，因為可能有沒存檔的文件。）"
  exit 1
}

# ── 2. 複製 manifest ────────────────────────────────────────────────────────
$icons = @()
$done = 0
foreach ($c in $Channels) {
  $src = Join-Path $root "manifests\$c"
  if (-not (Test-Path $src)) {
    Write-Host "  $c：manifests\$c 不存在，跳過（先跑 npm run manifests）" -ForegroundColor Yellow
    continue
  }
  if (-not $folders.ContainsKey($c)) {
    Write-Host "  $c：不認得這個頻道，跳過" -ForegroundColor Yellow
    continue
  }
  $dst = Join-Path $env:LOCALAPPDATA $folders[$c]
  if (-not (Test-Path $dst)) {
    Write-Host "  $c：還沒安裝過（找不到 $dst），跳過" -ForegroundColor Yellow
    continue
  }

  Copy-Item (Join-Path $src '*.xml') $dst -Force
  $word = Join-Path $dst 'word.xml'
  $ver = (Select-String -Path $word -Pattern '<Version>([\d.]+)</Version>').Matches[0].Groups[1].Value
  $url = (Select-String -Path $word -Pattern 'IconUrl DefaultValue="([^"]+)"').Matches[0].Groups[1].Value
  Write-Host ("  {0,-10} v{1,-9} {2}" -f $c, $ver, $url) -ForegroundColor Green
  $icons += $url
  $done++
}

if ($done -eq 0) { Write-Host "`n沒有裝到任何東西。" -ForegroundColor Yellow; exit 1 }

# ── 3. 清快取（一定在複製之後）──────────────────────────────────────────────
Write-Host ""
$wef = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'
$n = 0
foreach ($sub in @('Resources', 'AppCommands', 'AggregatedCache')) {
  $dir = Join-Path $wef $sub
  if (Test-Path $dir) {
    $n += (Get-ChildItem $dir -Recurse -File -ErrorAction SilentlyContinue).Count
    Remove-Item (Join-Path $dir '*') -Recurse -Force -ErrorAction SilentlyContinue
  }
}
Write-Host "  已清掉 $n 個快取檔案"

# ── 4. 圖示真的抓得到才算完成 ───────────────────────────────────────────────
#
# 這一步不是禮貌性的檢查。Office 抓不到圖的話會把失敗記住，而且**在下次改版本號
# 之前不會再試** —— 也就是說，現在讓它去抓一次失敗的代價，是接下來整個版本都看不到
# 圖示。寧可在這裡紅字擋下來。
Write-Host ""
$bad = 0
foreach ($u in ($icons | Sort-Object -Unique)) {
  try {
    $r = Invoke-WebRequest $u -TimeoutSec 10 -UseBasicParsing
    if ($r.StatusCode -eq 200) { Write-Host "  OK   $u" -ForegroundColor Green }
    else { Write-Host "  $($r.StatusCode)  $u" -ForegroundColor Red; $bad++ }
  } catch {
    Write-Host "  抓不到  $u" -ForegroundColor Red
    $bad++
  }
}

Write-Host ""
if ($bad -gt 0) {
  Write-Host "有 $bad 個圖示抓不到——現在開 Office 的話它會記住失敗，" -ForegroundColor Red
  Write-Host "而且要等到下次 npm run bump 才會再試一次。" -ForegroundColor Red
  Write-Host "localhost 抓不到通常是 npm run dev 沒在跑。" -ForegroundColor Yellow
  exit 1
}

Write-Host "好了，可以開 Office 了。" -ForegroundColor Cyan
