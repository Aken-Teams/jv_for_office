<#
.SYNOPSIS
  把 manifest 的變更（圖示、標籤、網址）套用到 Excel —— 一個指令做完整套。

.DESCRIPTION
  這件事有四個步驟，順序錯一步就完全無效，而且失敗的樣子跟成功前一模一樣
  （功能區顯示一個通用立方體），根本看不出是哪一步漏了。所以把順序寫死在這裡：

    1. 確認任務窗格伺服器在跑
         清完快取後的第一次 Excel 啟動，是 Office 唯一一次去抓功能區圖示的機會。
         抓不到就記成失敗，在下次改版本號之前都不會再試。

    2. 確認 Excel 已完全關閉
         Office 在關閉時才把快取寫回磁碟，開著清等於沒清。

    3. 版本號 +1，並重新安裝
         Office 用「增益集 ID + manifest 版本」判斷要不要重讀。版本沒動它連看都
         不看，即使你已經把圖示快取刪光、即使那些網址回應完全正常。
         而且 Excel 讀的是 %LOCALAPPDATA% 底下那份副本，不是 repo 裡的檔案，
         所以一定要重跑安裝程式。

    4. 清掉圖示快取

  跑完之後開 Excel，就這樣。

.EXAMPLE
  npm run refresh
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path)

Write-Host ""
Write-Host "=== 套用增益集變更 ===" -ForegroundColor Cyan
Write-Host ""

# ── 1. 伺服器 ────────────────────────────────────────────────────────────────
try {
  $probe = Invoke-WebRequest -Uri 'https://localhost:3443/assets/icon-32.png' -UseBasicParsing -TimeoutSec 5
  if ($probe.StatusCode -ne 200) { throw "HTTP $($probe.StatusCode)" }
  Write-Host "  [1/4] 任務窗格伺服器：正常" -ForegroundColor Green
} catch {
  Write-Host "  [1/4] 任務窗格伺服器沒有回應。" -ForegroundColor Red
  Write-Host ""
  Write-Host "        請先在另一個終端機視窗執行 npm run dev，讓它一直開著，" -ForegroundColor Yellow
  Write-Host "        再回來跑一次這個指令。" -ForegroundColor Yellow
  Write-Host ""
  Write-Host "        （沒有它，Office 抓不到圖示，這一輪會白做。）"
  Write-Host ""
  exit 1
}

# ── 2. Excel ─────────────────────────────────────────────────────────────────
$excel = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($excel) {
  Write-Host "  [2/4] Excel 正在執行（pid $($excel.Id -join ', ')）。" -ForegroundColor Red
  Write-Host ""
  Write-Host "        請存檔並完全關閉 Excel，再跑一次這個指令。" -ForegroundColor Yellow
  Write-Host "        （Office 是在關閉時才把快取寫回磁碟的，開著清等於沒清。）"
  Write-Host ""
  exit 1
}
Write-Host "  [2/4] Excel 已關閉" -ForegroundColor Green

# ── 3. 版本 + 安裝 ───────────────────────────────────────────────────────────
Push-Location $root
try {
  # Read the version straight from the file, before and after. Parsing node's
  # stdout looked tidier but breaks here: node writes UTF-8 while PowerShell 5.1
  # reads the console in the ANSI codepage, so the Chinese in that line arrives
  # as mojibake, Select-String matches nothing, and .ToString() throws on null.
  $manifestPath = Join-Path $root 'manifest.xml'
  $rx = '<Version>([\d.]+)</Version>'
  $before = [regex]::Match((Get-Content $manifestPath -Raw), $rx).Groups[1].Value

  & node (Join-Path $root 'tools\bump-version.mjs') | Out-Null
  if ($LASTEXITCODE -ne 0) { throw 'bump-version 失敗' }

  $after = [regex]::Match((Get-Content $manifestPath -Raw), $rx).Groups[1].Value
  Write-Host "  [3/4] manifest 版本：$before → $after" -ForegroundColor Green

  $install = & powershell -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root 'installer\install.ps1') 2>&1
  if ($LASTEXITCODE -ne 0) { $install | Write-Host; throw '安裝失敗' }
  Write-Host "        已更新 %LOCALAPPDATA% 底下的 manifest 副本"
} finally {
  Pop-Location
}

# ── 4. 快取 ──────────────────────────────────────────────────────────────────
$wef = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'
$cleared = 0
foreach ($name in @('Resources', 'AppCommands')) {
  $dir = Join-Path $wef $name
  if (-not (Test-Path $dir)) { continue }
  $files = @(Get-ChildItem $dir -Recurse -File -ErrorAction SilentlyContinue)
  if (-not $files.Count) { continue }
  try {
    Remove-Item (Join-Path $dir '*') -Recurse -Force -ErrorAction Stop
    $cleared += $files.Count
  } catch { Write-Host "        $name 清除失敗：$($_.Exception.Message)" -ForegroundColor Yellow }
}
Write-Host "  [4/4] 已清除圖示快取（$cleared 個檔案）" -ForegroundColor Green

Write-Host ""
Write-Host "全部就緒 —— 現在開啟 Excel。" -ForegroundColor Green
Write-Host '（npm run dev 請保持在跑，Excel 啟動時才抓得到圖示。）'
Write-Host ""
