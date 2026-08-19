<#
.SYNOPSIS
  安裝 JVision for Excel 增益集（單一使用者，不需要系統管理員權限）。

.DESCRIPTION
  Office 在 Windows 上有一個「每位使用者自行註冊」的機制：只要在
  HKCU\SOFTWARE\Microsoft\Office\16.0\Wef\Developer 底下放一個字串值
  （名稱 = manifest 的 <Id>，資料 = manifest 檔案的絕對路徑），Excel 下次啟動就會
  載入那個增益集。微軟自己的 office-addin-dev-settings 工具做的就是這件事。

  所以「安裝」實際上只有三個動作：
    1. 把 manifest.xml 複製到 %LOCALAPPDATA% 底下一個穩定的位置
       （放在使用者桌面或下載資料夾的檔案會被刪掉，之後 Excel 就載入失敗）
    2. 寫那一個登錄值
    3. 在「應用程式和功能」註冊一筆解除安裝資訊，讓它看起來像個正常的程式

  全部都在 HKCU 和 %LOCALAPPDATA%，所以不需要管理員權限，也不會影響其他使用者。

.PARAMETER TaskpaneUrl
  任務窗格的網址。不給就沿用 manifest.xml 裡原本的值（開發時是
  https://localhost:3443）。正式部署時傳內網網址，例如：

      .\install.ps1 -TaskpaneUrl https://jvision.panjit.com.tw

  這樣同一份安裝程式可以裝到不同環境，不必為每個環境維護一份 manifest。

.EXAMPLE
  .\install.ps1
  .\install.ps1 -TaskpaneUrl https://jvision.panjit.com.tw
#>
[CmdletBinding()]
param(
  [string]$TaskpaneUrl
)

$ErrorActionPreference = 'Stop'

$AppName      = 'JVision for Excel'
$AppShortName = 'JVisionForExcel'
$InstallDir   = Join-Path $env:LOCALAPPDATA $AppName
$WefKey       = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer'
$UninstallKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$AppShortName"

function Write-Step($msg) { Write-Host "  $msg" }
function Fail($msg) { Write-Host ""; Write-Host "安裝失敗：$msg" -ForegroundColor Red; Write-Host ""; exit 1 }

Write-Host ""
Write-Host "=== 安裝 $AppName ===" -ForegroundColor Cyan
Write-Host ""

# ── 1. 找到並讀取 manifest ────────────────────────────────────────────────────
# 安裝程式資料夾裡優先，其次是 repo 根目錄（開發時直接從 installer\ 執行的情況）。
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceManifest = Join-Path $here 'manifest.xml'
if (-not (Test-Path $sourceManifest)) { $sourceManifest = Join-Path (Split-Path -Parent $here) 'manifest.xml' }
if (-not (Test-Path $sourceManifest)) { Fail "找不到 manifest.xml（找過 $here 和它的上層資料夾）。" }

try {
  [xml]$manifest = Get-Content -Path $sourceManifest -Encoding UTF8 -Raw
} catch {
  Fail "manifest.xml 不是合法的 XML：$($_.Exception.Message)"
}

$addinId = $manifest.OfficeApp.Id
$version = $manifest.OfficeApp.Version
if (-not $addinId) { Fail "manifest.xml 裡沒有 <Id>。" }
Write-Step "增益集 ID：$addinId"
Write-Step "版本：$version"

# ── 2. 檢查 Office 版本 ───────────────────────────────────────────────────────
# 這個機制只存在於 Office 16.0（2016 之後的版本編號，含 Microsoft 365）。
# 找不到就代表這台機器的 Excel 不可能載入這個增益集，早點講比裝完才發現好。
if (-not (Test-Path 'HKCU:\SOFTWARE\Microsoft\Office\16.0')) {
  Fail "這台電腦沒有偵測到 Office 16.0（Microsoft 365 / Office 2016 以後）。JVision for Excel 需要 Microsoft 365 版的 Excel。"
}

# ── 3. 複製到穩定位置 ─────────────────────────────────────────────────────────
New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
$targetManifest = Join-Path $InstallDir 'manifest.xml'

if ($TaskpaneUrl) {
  # 把 manifest 裡所有舊網址換成指定的網址。開發用的 localhost 和任何先前設定的
  # 內網位址都會被換掉，所以重跑安裝程式換環境是安全的。
  $newBase = $TaskpaneUrl.TrimEnd('/')
  $raw = Get-Content -Path $sourceManifest -Raw -Encoding UTF8
  $raw = [regex]::Replace($raw, 'https://localhost:3443', $newBase)
  Set-Content -Path $targetManifest -Value $raw -Encoding UTF8 -NoNewline
  Write-Step "任務窗格網址：$newBase"
} else {
  Copy-Item -Path $sourceManifest -Destination $targetManifest -Force
  $src = $manifest.OfficeApp.DefaultSettings.SourceLocation.DefaultValue
  Write-Step "任務窗格網址：$src"
}

# 解除安裝程式必須跟著一起裝進去 —— 使用者刪掉下載資料夾之後，
# 「應用程式和功能」裡那顆解除安裝按鈕還是要能動。
Copy-Item -Path (Join-Path $here 'uninstall.ps1') -Destination (Join-Path $InstallDir 'uninstall.ps1') -Force
$iconSource = Join-Path (Split-Path -Parent $here) 'assets\icon-80.png'
if (Test-Path $iconSource) { Copy-Item $iconSource (Join-Path $InstallDir 'icon-80.png') -Force }
Write-Step "安裝位置：$InstallDir"

# ── 4. 註冊給 Excel ───────────────────────────────────────────────────────────
if (-not (Test-Path $WefKey)) { New-Item -Path $WefKey -Force | Out-Null }
# 官方工具用 manifest 路徑當值名稱是舊寫法，現在用的是 Id。兩種都清掉再寫，
# 避免同一個增益集在清單裡出現兩次。
Remove-ItemProperty -Path $WefKey -Name $targetManifest -ErrorAction SilentlyContinue
New-ItemProperty -Path $WefKey -Name $addinId -Value $targetManifest -PropertyType String -Force | Out-Null
Write-Step "已註冊到 Excel"

# ── 5. 讓它出現在「應用程式和功能」 ───────────────────────────────────────────
$uninstallCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File `"$InstallDir\uninstall.ps1`""
New-Item -Path $UninstallKey -Force | Out-Null
$props = @{
  DisplayName     = $AppName
  DisplayVersion  = $version
  Publisher       = $manifest.OfficeApp.ProviderName
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

# ── 6. 提醒重開 Excel ─────────────────────────────────────────────────────────
# Excel 只在啟動時讀這個登錄機碼，所以開著的 Excel 不會出現按鈕。
# 這裡刻意不強制關掉 Excel —— 使用者可能有未存檔的活頁簿。
$excel = Get-Process EXCEL -ErrorAction SilentlyContinue
Write-Host ""
Write-Host "安裝完成。" -ForegroundColor Green
if ($excel) {
  Write-Host ""
  Write-Host "！Excel 目前是開著的。請存檔後完全關閉 Excel 再重新開啟，" -ForegroundColor Yellow
  Write-Host "  增益集按鈕才會出現。" -ForegroundColor Yellow
}
Write-Host ""
Write-Host "開啟 Excel 後，在「常用」分頁最右邊會看到「JVision」。"
Write-Host ""
