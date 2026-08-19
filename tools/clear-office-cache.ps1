<#
.SYNOPSIS
  清掉 Office 的增益集圖示快取，讓 Excel 重新抓 manifest 裡的圖示。

.DESCRIPTION
  換了 assets\icon-*.png 之後，Excel 還是會顯示舊圖示——甚至顯示一個通用的方塊。
  原因是 Office 會把功能區圖示存進：

      %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\Resources    ← 圖檔本身
      %LOCALAPPDATA%\Microsoft\Office\16.0\Wef\AppCommands  ← 功能區命令定義

  這些快取沒有過期機制，manifest 的網址沒變它就不會重抓。重開 Excel 也沒用，
  因為快取在磁碟上而不是在記憶體裡。

  這裡刪掉的都是純快取，Office 下次啟動會自己重建。不會動到你的增益集註冊
  （那在登錄檔），所以清完不用重裝。

.PARAMETER All
  連 Wef 底下其他快取一起清（含 manifest 快取）。改了 manifest 內容但 Excel
  行為沒跟著變時用這個。範圍較大，其他增益集也會重新抓一次資料。

.EXAMPLE
  .\clear-office-cache.ps1
  .\clear-office-cache.ps1 -All
#>
[CmdletBinding()]
param([switch]$All)

$ErrorActionPreference = 'Stop'
$wef = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\Wef'

Write-Host ""
Write-Host "=== 清除 Office 增益集圖示快取 ===" -ForegroundColor Cyan
Write-Host ""

# Office rewrites these folders as it shuts down, so clearing them under a running
# Excel achieves nothing. Refuse rather than pretend it worked.
$excel = Get-Process EXCEL -ErrorAction SilentlyContinue
if ($excel) {
  Write-Host "Excel 正在執行（pid $($excel.Id -join ', ')）。" -ForegroundColor Yellow
  Write-Host "請先存檔並完全關閉 Excel，再執行一次這個腳本。" -ForegroundColor Yellow
  Write-Host "（Excel 關閉時會把快取寫回去，開著清等於沒清。）"
  Write-Host ""
  exit 1
}

if (-not (Test-Path $wef)) {
  Write-Host "找不到快取目錄，不用清：$wef"
  exit 0
}

$targets = if ($All) {
  Get-ChildItem $wef -Directory -ErrorAction SilentlyContinue
} else {
  @('Resources', 'AppCommands') |
    ForEach-Object { Join-Path $wef $_ } |
    Where-Object { Test-Path $_ } |
    ForEach-Object { Get-Item $_ }
}

$total = 0
foreach ($dir in $targets) {
  $files = @(Get-ChildItem $dir.FullName -Recurse -File -ErrorAction SilentlyContinue)
  if (-not $files.Count) { continue }
  try {
    Remove-Item (Join-Path $dir.FullName '*') -Recurse -Force -ErrorAction Stop
    Write-Host ("  已清 {0,-16} {1} 個檔案" -f $dir.Name, $files.Count)
    $total += $files.Count
  } catch {
    Write-Host ("  {0} 清除失敗：{1}" -f $dir.Name, $_.Exception.Message) -ForegroundColor Yellow
  }
}

Write-Host ""
if ($total) {
  Write-Host "共清掉 $total 個快取檔案。" -ForegroundColor Green
} else {
  Write-Host "快取本來就是空的。"
}
# The cache is now empty, so the very next Excel start is the ONE chance Office
# takes to download the ribbon icons. If the dev server isn't up at that exact
# moment the fetch fails, Office falls back to a generic cube, and it will not
# retry until the manifest version changes again. Silent, and indistinguishable
# from a broken icon file — so check now, while it can still be fixed.
Write-Host "接著：" -ForegroundColor Cyan
$devUp = $false
try {
  $probe = Invoke-WebRequest -Uri 'https://localhost:3443/assets/icon-32.png' -UseBasicParsing -TimeoutSec 5
  $devUp = $probe.StatusCode -eq 200
} catch { $devUp = $false }

if ($devUp) {
  Write-Host "  ✓ 任務窗格伺服器有在跑，圖示抓得到" -ForegroundColor Green
  Write-Host "  → 直接開啟 Excel，新圖示會在啟動時下載"
} else {
  Write-Host ""
  Write-Host "  ！先不要開 Excel。" -ForegroundColor Yellow
  Write-Host "    https://localhost:3443 沒有回應，代表 npm run dev 沒在跑。" -ForegroundColor Yellow
  Write-Host "    現在開 Excel 的話 Office 抓不到圖示，會顯示一個通用立方體，" -ForegroundColor Yellow
  Write-Host "    而且在下次改版本號之前都不會再試一次。" -ForegroundColor Yellow
  Write-Host ""
  Write-Host "    正確順序：先在另一個視窗跑 npm run dev，再開 Excel。"
}
Write-Host ""
