#Requires -Version 5.1
<#
.SYNOPSIS
  把 installer\JVisionSetup.cs 編成 installer\JVisionSetup.exe。

.DESCRIPTION
  用 csc.exe —— .NET Framework 4.x 內建的 C# 編譯器，每一台 Windows 10/11 都有。
  刻意不用 dotnet SDK、MSBuild 或任何 NuGet 套件：這個安裝程式是要交給客戶 IT 的，
  他們拿到原始碼應該在自己的機器上直接重建得出來，不必先裝一套開發環境。

  目前的 manifest.xml 會被嵌進 exe 裡，所以發出去的就是單一檔案。也就是說
  **改過 manifest 之後要重跑這個腳本**，否則 exe 裡還是舊的版本號和網址。

.PARAMETER Sign
  建置後用憑證簽章。沒有簽章的 exe 會被 SmartScreen 攔（「Windows 已保護您的電腦」），
  企業環境的防毒也常常直接隔離——要大量派送，簽章不是加分項而是必要條件。

      .\tools\build-installer.ps1 -Sign -CertThumbprint A1B2C3...

.EXAMPLE
  .\tools\build-installer.ps1
#>
[CmdletBinding()]
param(
  [switch]$Sign,
  [string]$CertThumbprint,
  [string]$TimestampUrl = 'http://timestamp.digicert.com'
)

$ErrorActionPreference = 'Stop'

$root      = Split-Path -Parent $PSScriptRoot
$installer = Join-Path $root 'installer'
$source    = Join-Path $installer 'JVisionSetup.cs'
$outExe    = Join-Path $installer 'JVisionSetup.exe'
$logo      = Join-Path $root 'assets\logo.png'

# ── 哪一套環境 ───────────────────────────────────────────────────────────────
#
# 安裝程式是把 manifest **嵌進 exe** 的，而且只讀 manifests\ 最上層那三個——所以
# 「我開這個看不到測試區安裝」是對的：產在 manifests\test\ 底下的那三份，建置的
# 時候根本沒被看到。
#
# 這裡照 .env 的 DEPLOY_CHANNEL 決定要嵌哪一組，並且：
#   · 測試區的 exe 另外取名，不要蓋掉正式區那一支
#   · 安裝程式自己的視窗標題也要標出來，不然按下去之前看不出裝的是哪一套
$envFile = Join-Path $root '.env'
$channel = 'prod'
$label   = ''
if (Test-Path $envFile) {
  foreach ($line in Get-Content -Path $envFile -Encoding UTF8) {
    if ($line -match '^\s*DEPLOY_CHANNEL\s*=\s*(.+?)\s*$') { $channel = $Matches[1].Trim('"', "'").ToLower() }
    if ($line -match '^\s*DEPLOY_LABEL\s*=\s*(.+?)\s*$')   { $label   = $Matches[1].Trim('"', "'") }
  }
}
if ($env:DEPLOY_CHANNEL) { $channel = $env:DEPLOY_CHANNEL.ToLower() }
if ($env:DEPLOY_LABEL)   { $label   = $env:DEPLOY_LABEL }

$isProd = ($channel -eq 'prod' -or $channel -eq 'production' -or -not $channel)
if (-not $isProd) {
  $known = @{ test = '測試'; staging = '預備'; dev = '開發'; uat = '驗收' }
  if (-not $label) { $label = if ($known.ContainsKey($channel)) { $known[$channel] } else { $channel } }
  $manifestDir = Join-Path $root "manifests\$channel"
  $outExe      = Join-Path $installer "JVisionSetup-$channel.exe"
  Write-Host "頻道：$channel（$label）— 嵌入 manifests\$channel\，輸出 $(Split-Path -Leaf $outExe)" -ForegroundColor Yellow
  if (-not (Test-Path $manifestDir)) {
    throw "找不到 $manifestDir。先跑 npm run manifests 產生非正式環境的 manifest。"
  }
} else {
  $manifestDir = Join-Path $root 'manifests'
}

if (-not (Test-Path $source))   { throw "找不到 $source" }
if (-not (Test-Path $manifestDir)) { throw "找不到 $manifestDir" }
$manifests = @(Get-ChildItem -Path $manifestDir -Filter '*.xml' | Sort-Object Name)
if (-not $manifests.Count) { throw "$manifestDir 裡沒有 manifest。" }
if (-not (Test-Path $logo))     { throw "找不到 $logo" }

# csc.exe：64 位元優先，找不到就退回 32 位元（很舊的機器）。
$csc = @(
  "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe",
  "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $csc) {
  throw "找不到 csc.exe。這台電腦需要 .NET Framework 4.x（Windows 10/11 預設就有）。"
}

# 版本號從 manifest 來，這樣 exe 的檔案內容和它安裝的東西不會對不上。
# 三份 manifest 的版本號由 npm run bump 寫成一致，這裡取第一份就行——
# 但還是檢查一下，因為不一致的話使用者會看到一個版本、裝到另一個。
$versions = $manifests | ForEach-Object { ([xml](Get-Content -Path $_.FullName -Encoding UTF8 -Raw)).OfficeApp.Version }
$version = $versions | Select-Object -First 1
if (-not $version) { throw "$($manifests[0].Name) 裡沒有 <Version>。" }
if (($versions | Sort-Object -Unique).Count -gt 1) {
  throw "manifests 裡的版本號不一致（$($versions -join '、')）。先跑 npm run bump。"
}

# AssemblyInfo 一起編進去，exe 的「內容」頁籤才有正常的說明和版本，
# 而不是一片空白 —— 空白的屬性頁正是使用者不敢點的那種檔案。
$asmInfo = Join-Path $installer 'AssemblyInfo.cs'
@"
using System.Reflection;
[assembly: AssemblyTitle("JVision for Office 安裝程式")]
[assembly: AssemblyProduct("JVision for Office")]
[assembly: AssemblyCompany("PANJIT")]
[assembly: AssemblyDescription("安裝 JVision for Office 增益集（Excel / Word / PowerPoint）")]
[assembly: AssemblyVersion("$version")]
[assembly: AssemblyFileVersion("$version")]
[assembly: AssemblyCopyright("Copyright (c) PANJIT")]
// 著作權欄位在「內容 → 詳細資料」裡；資安人員檢查一個不明來源的 exe 時，
// 看的就是這一頁。簽章解決「發行者」那一行，這裡解決其餘所有欄位。
"@ | Set-Content -Path $asmInfo -Encoding UTF8

# 安裝程式自己的標題也要標。改的是一份暫存副本，原始碼不動——不然切回正式區的
# 時候會留下一個改壞的檔案。
if (-not $isProd) {
  $patched = Join-Path $env:TEMP "JVisionSetup-$channel.cs"
  # AppName 決定安裝資料夾（LocalAppData\<AppName>），ShortName 決定「應用程式和
  # 功能」那筆解除安裝紀錄的登錄機碼。**兩個都要換**——只換 AppName 的話資料夾分開了，
  # 但兩個版本仍然共用同一筆解除安裝紀錄，移除其中一個會把另一個的紀錄一起帶走。
  $tag = ($channel -replace '[^a-zA-Z0-9]', '')
  (Get-Content -Path $source -Encoding UTF8 -Raw) `
    -replace 'AppName\s*=\s*"JVision for Office"', "AppName = `"JVision for Office（$label）`"" `
    -replace 'ShortName\s*=\s*"JVisionForOffice"', "ShortName = `"JVisionForOffice_$tag`"" `
    | Set-Content -Path $patched -Encoding UTF8
  $source = $patched
}

$args = @(
  '/nologo'
  '/target:winexe'          # 沒有主控台視窗
  '/platform:anycpu'
  '/optimize+'
  "/out:$outExe"
  "/resource:$logo,logo.png"
  '/reference:System.dll'
  '/reference:System.Core.dll'
  '/reference:System.Drawing.dll'
  '/reference:System.Windows.Forms.dll'
  '/reference:System.Xml.dll'
  '/reference:System.Xml.Linq.dll'
)

# 每份 manifest 各嵌一個資源，名字就是檔名——JVisionSetup.cs 的 HostSpec.File
# 拿同一個字串去讀它，所以這裡多一份檔、那邊多一列，就多一個主應用程式。
foreach ($mf in $manifests) { $args += "/resource:$($mf.FullName),$($mf.Name)" }

# 圖示：exe 在檔案總管裡不能長得像個沒人認識的東西。
$ico = Join-Path $installer 'JVision.ico'
if (Test-Path $ico) { $args += "/win32icon:$ico" }
else { Write-Host "（沒有 installer\JVision.ico，exe 會用預設圖示。跑 npm run installer:icon 可以產生）" -ForegroundColor DarkGray }

$args += $source
$args += $asmInfo

Write-Host "編譯 JVisionSetup.exe（版本 $version）…"
& $csc @args
if ($LASTEXITCODE -ne 0) { throw "編譯失敗（csc 離開代碼 $LASTEXITCODE）。" }
Remove-Item $asmInfo -Force -ErrorAction SilentlyContinue

if ($Sign) {
  if (-not $CertThumbprint) { throw "-Sign 需要一起給 -CertThumbprint。" }
  $signtool = Get-ChildItem "${env:ProgramFiles(x86)}\Windows Kits\10\bin" -Recurse -Filter signtool.exe -ErrorAction SilentlyContinue |
    Where-Object { $_.FullName -match 'x64' } | Select-Object -First 1
  if (-not $signtool) { throw "找不到 signtool.exe（需要 Windows SDK）。" }
  & $signtool.FullName sign /sha1 $CertThumbprint /fd SHA256 /tr $TimestampUrl /td SHA256 $outExe
  if ($LASTEXITCODE -ne 0) { throw "簽章失敗。" }
  Write-Host "已簽章。" -ForegroundColor Green
}

$size = [math]::Round((Get-Item $outExe).Length / 1KB, 1)

# ── 交付物：檢查碼與 manifest 包 ────────────────────────────────────────────
#
# 這兩樣不是為了我們自己，是為了**對方的資安和 IT**：
#
#   · SHA-256 —— 客戶把 exe 從你的伺服器下載回去之後，要能驗證它沒被掉包。
#     沒有這個檔的話，對方的資安多半會直接退回來問，而那一來一往是幾天。
#
#   · manifest 包 —— 有些客戶不讓員工自己跑 exe，而是由 IT 從 Microsoft 365
#     系統管理中心統一部署。那條路只要三個 .xml，不要 exe。先備好，問起來就有。
#
# 兩個都放在 exe 旁邊，一起給出去。
$hash = (Get-FileHash -Algorithm SHA256 $outExe).Hash.ToLower()
$hashFile = "$outExe.sha256"
# 用 sha256sum 的格式（雜湊 + 兩個空格 + 檔名），對方拿 certutil 或 sha256sum 都驗得了。
"$hash  $(Split-Path -Leaf $outExe)" | Set-Content -Path $hashFile -Encoding ascii

$zipName = if ($isProd) { 'JVision-manifests.zip' } else { "JVision-manifests-$channel.zip" }
$zipPath = Join-Path $installer $zipName
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path (Join-Path $manifestDir '*.xml') -DestinationPath $zipPath

Write-Host ""
Write-Host "完成：$outExe（$size KB）" -ForegroundColor Green
Write-Host "  SHA-256：$hash"
Write-Host "  檢查碼檔：$hashFile"
Write-Host "  manifest 包：$zipPath（給走管理中心部署的客戶）"
if (-not $Sign) {
  Write-Host ""
  Write-Host "！這個 exe 沒有簽章。" -ForegroundColor Yellow
  Write-Host "  自己測試沒問題，但要發給全公司之前請用公司憑證簽過——" -ForegroundColor Yellow
  Write-Host "  沒簽章的 exe 會被 SmartScreen 擋，也常被防毒直接隔離。" -ForegroundColor Yellow
}
Write-Host ""
Write-Host "IT 派送用法："
Write-Host "  JVisionSetup.exe /silent /url:https://jvision.panjit.com.tw
  JVisionSetup.exe /silent /hosts:word            # 只裝 Word"
Write-Host "  JVisionSetup.exe /silent /uninstall"
Write-Host "  離開代碼：0 成功、1 失敗、2 沒有 Office 16.0、3 使用者取消"
Write-Host "  必須在使用者身分下執行（設定寫在 HKCU，不是 HKLM）。"
Write-Host ""
