$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$softwareDir = Get-ChildItem -Path $root -Directory |
  Where-Object { Test-Path (Join-Path $_.FullName "NPOI.dll") } |
  Select-Object -First 1 -ExpandProperty FullName
if (-not $softwareDir) {
  throw "Could not find software directory containing NPOI.dll under $root"
}
$outDir = Join-Path $PSScriptRoot "bin"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$csc = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$out = Join-Path $outDir "QuotaLearningImporter.exe"
$pluginOut = Join-Path $outDir "RecoQuotaRecommend.dll"
$npoi = Join-Path $softwareDir "NPOI.dll"
$npoiOoxml = Join-Path $softwareDir "NPOI.OOXML.dll"
$npoiOpenXml4Net = Join-Path $softwareDir "NPOI.OpenXml4Net.dll"
$npoiOpenXmlFormats = Join-Path $softwareDir "NPOI.OpenXmlFormats.dll"
$source = Join-Path $PSScriptRoot "QuotaLearningImporter.cs"
$pluginSource = Join-Path $PSScriptRoot "QuotaRecommendPanel.cs"

& $csc /nologo /target:exe /out:$out `
  /reference:$npoi `
  /reference:$npoiOoxml `
  /reference:$npoiOpenXml4Net `
  /reference:$npoiOpenXmlFormats `
  $source

& $csc /nologo /target:library /out:$pluginOut `
  /reference:System.Windows.Forms.dll `
  /reference:System.Drawing.dll `
  /reference:System.Data.dll `
  $pluginSource

Copy-Item -LiteralPath $npoi -Destination $outDir -Force
Copy-Item -LiteralPath $npoiOoxml -Destination $outDir -Force
Copy-Item -LiteralPath $npoiOpenXml4Net -Destination $outDir -Force
Copy-Item -LiteralPath $npoiOpenXmlFormats -Destination $outDir -Force
Copy-Item -LiteralPath (Join-Path $softwareDir "ICSharpCode.SharpZipLib.dll") -Destination $outDir -Force
Copy-Item -LiteralPath $pluginOut -Destination $softwareDir -Force

$dataSource = Join-Path $root "RecoQuotaData"
if (Test-Path $dataSource) {
  Copy-Item -LiteralPath $dataSource -Destination $softwareDir -Recurse -Force
}

Write-Host "Built $out"
Write-Host "Built $pluginOut"
Write-Host "Deployed RecoQuotaRecommend.dll to $softwareDir"
