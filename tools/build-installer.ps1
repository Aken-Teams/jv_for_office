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
$manifestDir = Join-Path $root 'manifests'
$outExe    = Join-Path $installer 'JVisionSetup.exe'
$logo      = Join-Path $root 'assets\logo.png'

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
[assembly: AssemblyDescription("安裝 JVision for Office 增益集（Excel / Word）")]
[assembly: AssemblyVersion("$version")]
[assembly: AssemblyFileVersion("$version")]
"@ | Set-Content -Path $asmInfo -Encoding UTF8

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
Write-Host ""
Write-Host "完成：$outExe（$size KB）" -ForegroundColor Green
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
