<#
.SYNOPSIS
  解除安裝 JVision for Excel。

.DESCRIPTION
  把 install.ps1 做的三件事反過來做：拿掉 Excel 的註冊、拿掉「應用程式和功能」的
  項目、刪掉安裝資料夾。同樣全部在 HKCU 和 %LOCALAPPDATA%，不需要管理員權限。

  這份腳本會被複製進安裝資料夾，所以「應用程式和功能」裡的解除安裝按鈕在使用者
  刪掉原始下載檔之後仍然有效。
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'

$AppName      = 'JVision for Office'
$AppShortName = 'JVisionForOffice'
$InstallDir   = Join-Path $env:LOCALAPPDATA $AppName
$WefKey       = 'HKCU:\SOFTWARE\Microsoft\Office\16.0\Wef\Developer'
$UninstallKey = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\$AppShortName"

Write-Host ""
Write-Host "=== 解除安裝 $AppName ===" -ForegroundColor Cyan
Write-Host ""

# ── 1. 取消 Office 的註冊 ──────────────────────────────────────────────────────
# 從已安裝的 manifest 讀回 Id。讀不到（檔案被手動刪了）就退而求其次，
# 掃整個 Developer 機碼，清掉任何指向我們安裝資料夾的值 —— 留下一個指向不存在
# 檔案的登錄值，會讓 Excel 每次啟動都嘗試載入一個載不到的增益集。
# 每個主應用程式一份 manifest，所以這裡收集的是一串 Id，不是一個。
$addinIds = @()
if (Test-Path $InstallDir) {
  foreach ($f in Get-ChildItem -Path $InstallDir -Filter '*.xml' -ErrorAction SilentlyContinue) {
    try { [xml]$m = Get-Content -Path $f.FullName -Encoding UTF8 -Raw; if ($m.OfficeApp.Id) { $addinIds += $m.OfficeApp.Id } } catch { }
  }
}

if (Test-Path $WefKey) {
  foreach ($id in $addinIds) {
    Remove-ItemProperty -Path $WefKey -Name $id -ErrorAction SilentlyContinue
  }
  $key = Get-Item -Path $WefKey
  foreach ($name in $key.GetValueNames()) {
    $data = $key.GetValue($name)
    if ($data -is [string] -and $data -like "$InstallDir*") {
      Remove-ItemProperty -Path $WefKey -Name $name -ErrorAction SilentlyContinue
    }
  }
  Write-Host "  已取消 Office 的註冊"
}

# ── 2. 從「應用程式和功能」移除 ───────────────────────────────────────────────
if (Test-Path $UninstallKey) {
  Remove-Item -Path $UninstallKey -Recurse -Force
  Write-Host "  已從「應用程式和功能」移除"
}

# ── 3. 刪掉安裝資料夾 ─────────────────────────────────────────────────────────
# 這份腳本自己就在那個資料夾裡。Windows 不會鎖住執行中的 .ps1（跟 .exe 不同），
# 所以可以直接刪，不需要那種「排程下次開機刪除」的把戲。
if (Test-Path $InstallDir) {
  try {
    Remove-Item -Path $InstallDir -Recurse -Force
    Write-Host "  已刪除安裝資料夾"
  } catch {
    Write-Host "  安裝資料夾刪除失敗（可能有檔案被鎖住）：$InstallDir" -ForegroundColor Yellow
  }
}

$excel = Get-Process EXCEL -ErrorAction SilentlyContinue
Write-Host ""
Write-Host "已解除安裝。" -ForegroundColor Green
if ($excel) {
  Write-Host ""
  Write-Host "！Excel 目前是開著的，重新啟動 Excel 後按鈕才會消失。" -ForegroundColor Yellow
}
Write-Host ""
