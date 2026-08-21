<#
.SYNOPSIS
  安裝 JVision for Office 增益集（單一使用者，不需要系統管理員權限）。

.DESCRIPTION
  Office 在 Windows 上有一個「每位使用者自行註冊」的機制：只要在
  HKCU\SOFTWARE\Microsoft\Office\16.0\Wef\Developer 底下放一個字串值
  （名稱 = manifest 的 <Id>，資料 = manifest 檔案的絕對路徑），Office 下次啟動就會
  載入那個增益集。微軟自己的 office-addin-dev-settings 工具做的就是這件事。

  所以「安裝」實際上只有三個動作：
    1. 把每個主應用程式的 manifest 複製到 %LOCALAPPDATA% 底下一個穩定的位置
       （放在使用者桌面或下載資料夾的檔案會被刪掉，之後 Office 就載入失敗）
    2. 每份 manifest 各寫一個登錄值
    3. 在「應用程式和功能」註冊一筆解除安裝資訊，讓它看起來像個正常的程式

  全部都在 HKCU 和 %LOCALAPPDATA%，所以不需要管理員權限，也不會影響其他使用者。

  ── 為什麼是「每個主應用程式一份 manifest」 ──
  一份 manifest 其實可以同時宣告 Workbook 和 Document，但那樣裝了就是兩個一起來，
  只想用 Word 的人也會在 Excel 看到按鈕，而且沒辦法單獨移除其中一個。三份各自一個
  Id、一個登錄值，指向同一個 taskpane.html，由 Office.onReady 判斷自己在哪。
  詳見 docs/multi-host.md。

.PARAMETER Hosts
  要安裝哪幾個主應用程式，逗號分隔。預設全部。

      .\install.ps1 -Hosts word
      .\install.ps1 -Hosts excel,word

.PARAMETER TaskpaneUrl
  任務窗格的網址。不給就沿用 manifest 裡原本的值（開發時是
  https://localhost:3443）。正式部署時傳內網網址，例如：

      .\install.ps1 -TaskpaneUrl https://jvision.panjit.com.tw

  這樣同一份安裝程式可以裝到不同環境，不必為每個環境維護一份 manifest。

.EXAMPLE
  .\install.ps1
  .\install.ps1 -Hosts word -TaskpaneUrl https://jvision.panjit.com.tw
#>
[CmdletBinding()]
param(
  [string]$TaskpaneUrl,
  [string]$Hosts = 'excel,word'
)

$ErrorActionPreference = 'Stop'

$AppName      = 'JVision for Office'
$AppShortName = 'JVisionForOffice'
$InstallDir   = Join-Path $env:LOCALAPPDATA $AppName
$WefKey       = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer'
$UninstallKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$AppShortName"

# 主應用程式一覽。加 PowerPoint 時在這裡多一列，其餘都不用動。
$KnownHosts = [ordered]@{
  excel = @{ File = 'excel.xml'; App = 'Excel'; Process = 'EXCEL' }
  word  = @{ File = 'word.xml';  App = 'Word';  Process = 'WINWORD' }
}

function Write-Step($msg) { Write-Host "  $msg" }
function Fail($msg) { Write-Host ""; Write-Host "安裝失敗：$msg" -ForegroundColor Red; Write-Host ""; exit 1 }

Write-Host ""
Write-Host "=== 安裝 $AppName ===" -ForegroundColor Cyan
Write-Host ""

# ── 0. 決定要裝哪幾個 ─────────────────────────────────────────────────────────
$selected = @()
foreach ($h in ($Hosts -split ',')) {
  $key = $h.Trim().ToLower()
  if (-not $key) { continue }
  if (-not $KnownHosts.Contains($key)) { Fail "不認得的主應用程式「$key」。可以用的是：$($KnownHosts.Keys -join '、')。" }
  if ($selected -notcontains $key) { $selected += $key }
}
if (-not $selected.Count) { Fail "-Hosts 至少要指定一個主應用程式。" }

# ── 1. 找到 manifest 資料夾 ───────────────────────────────────────────────────
# 安裝程式資料夾裡優先，其次是 repo 根目錄（開發時直接從 installer\ 執行的情況）。
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$manifestDir = Join-Path $here 'manifests'
if (-not (Test-Path $manifestDir)) { $manifestDir = Join-Path (Split-Path -Parent $here) 'manifests' }
if (-not (Test-Path $manifestDir)) { Fail "找不到 manifests 資料夾（找過 $here 和它的上層資料夾）。" }

# ── 2. 檢查 Office 版本 ───────────────────────────────────────────────────────
# 這個機制只存在於 Office 16.0（2016 之後的版本編號，含 Microsoft 365）。
# 找不到就代表這台機器不可能載入這個增益集，早點講比裝完才發現好。
if (-not (Test-Path 'HKCU:\SOFTWARE\Microsoft\Office\16.0')) {
  Fail "這台電腦沒有偵測到 Office 16.0（Microsoft 365 / Office 2016 以後）。$AppName 需要 Microsoft 365 版的 Office。"
}

New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
if (-not (Test-Path $WefKey)) { New-Item -Path $WefKey -Force | Out-Null }

# ── 3. 每個主應用程式：複製 manifest、寫登錄值 ────────────────────────────────
$version = $null
$provider = 'PANJIT'
$installed = @()

foreach ($key in $selected) {
  $spec = $KnownHosts[$key]
  $source = Join-Path $manifestDir $spec.File
  if (-not (Test-Path $source)) { Fail "找不到 $($spec.File)（在 $manifestDir）。" }

  try {
    [xml]$manifest = Get-Content -Path $source -Encoding UTF8 -Raw
  } catch {
    Fail "$($spec.File) 不是合法的 XML：$($_.Exception.Message)"
  }
  $addinId = $manifest.OfficeApp.Id
  if (-not $addinId) { Fail "$($spec.File) 裡沒有 <Id>。" }
  $version = $manifest.OfficeApp.Version
  if ($manifest.OfficeApp.ProviderName) { $provider = $manifest.OfficeApp.ProviderName }

  $target = Join-Path $InstallDir $spec.File
  if ($TaskpaneUrl) {
    # 把 manifest 裡所有舊網址換成指定的網址。開發用的 localhost 和任何先前設定的
    # 內網位址都會被換掉，所以重跑安裝程式換環境是安全的。
    $newBase = $TaskpaneUrl.TrimEnd('/')
    $raw = Get-Content -Path $source -Raw -Encoding UTF8
    $raw = [regex]::Replace($raw, 'https://localhost:3443', $newBase)
    Set-Content -Path $target -Value $raw -Encoding UTF8 -NoNewline
  } else {
    Copy-Item -Path $source -Destination $target -Force
  }

  # 官方工具用 manifest 路徑當值名稱是舊寫法，現在用的是 Id。兩種都清掉再寫，
  # 避免同一個增益集在清單裡出現兩次。
  Remove-ItemProperty -Path $WefKey -Name $target -ErrorAction SilentlyContinue
  New-ItemProperty -Path $WefKey -Name $addinId -Value $target -PropertyType String -Force | Out-Null

  Write-Step "$($spec.App)：已註冊（$addinId）"
  $installed += $spec
}

if ($TaskpaneUrl) { Write-Step "任務窗格網址：$($TaskpaneUrl.TrimEnd('/'))" }
Write-Step "版本：$version"
Write-Step "安裝位置：$InstallDir"

# 解除安裝程式必須跟著一起裝進去 —— 使用者刪掉下載資料夾之後，
# 「應用程式和功能」裡那顆解除安裝按鈕還是要能動。
Copy-Item -Path (Join-Path $here 'uninstall.ps1') -Destination (Join-Path $InstallDir 'uninstall.ps1') -Force
$iconSource = Join-Path (Split-Path -Parent $here) 'assets\icon-80.png'
if (Test-Path $iconSource) { Copy-Item $iconSource (Join-Path $InstallDir 'icon-80.png') -Force }

# ── 4. 讓它出現在「應用程式和功能」 ───────────────────────────────────────────
# 一筆，不是每個主應用程式一筆：使用者心裡那個「程式」是 JVision，不是三個東西。
$uninstallCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$InstallDir\uninstall.ps1`""
# 舊版是以「JVision for Excel」登記的。留著會在「應用程式和功能」裡看到兩筆，
# 而且舊的那筆指向一個已經不存在的解除安裝腳本。
$legacyKey = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\JVisionForExcel'
if (Test-Path $legacyKey) { Remove-Item -Path $legacyKey -Recurse -Force }
$legacyDir = Join-Path $env:LOCALAPPDATA 'JVision for Excel'
if (Test-Path $legacyDir) { Remove-Item -Path $legacyDir -Recurse -Force -ErrorAction SilentlyContinue }

New-Item -Path $UninstallKey -Force | Out-Null
$props = @{
  DisplayName     = "$AppName（$(($installed | ForEach-Object { $_.App }) -join '、')）"
  DisplayVersion  = $version
  Publisher       = $provider
  InstallLocation = $InstallDir
  UninstallString = $uninstallCmd
  NoModify        = 1
  NoRepair        = 1
}
foreach ($k in $props.Keys) {
  $type = 'String'
  if ($props[$k] -is [int]) { $type = 'DWord' }
  New-ItemProperty -Path $UninstallKey -Name $k -Value $props[$k] -PropertyType $type -Force | Out-Null
}
Write-Step "已登記到「應用程式和功能」"

# ── 5. 提醒重開 ───────────────────────────────────────────────────────────────
# Office 只在啟動時讀這個登錄機碼，所以開著的視窗不會出現按鈕。
# 這裡刻意不強制關掉 —— 使用者可能有未存檔的檔案。
Write-Host ""
Write-Host "安裝完成。" -ForegroundColor Green

$running = @($installed | Where-Object { Get-Process $_.Process -ErrorAction SilentlyContinue })
if ($running.Count) {
  Write-Host ""
  Write-Host "！$(($running | ForEach-Object { $_.App }) -join '、') 目前是開著的。請存檔後完全關閉再重新開啟，" -ForegroundColor Yellow
  Write-Host "  增益集按鈕才會出現。" -ForegroundColor Yellow
}
Write-Host ""
Write-Host "開啟 $(($installed | ForEach-Object { $_.App }) -join ' / ') 後，在「常用」分頁最右邊會看到「JVision」。"
Write-Host ""
