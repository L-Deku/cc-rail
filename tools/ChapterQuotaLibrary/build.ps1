$ErrorActionPreference = "Stop"

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$outDir = Join-Path $PSScriptRoot "bin"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

# Find a software directory containing NPOI.dll (same discovery as RecoQuotaRecommend/build.ps1)
$referenceDir = Get-ChildItem -LiteralPath $root -Directory -Recurse |
  Where-Object {
    (Test-Path -LiteralPath (Join-Path $_.FullName "NPOI.dll")) -and
    (
      (Test-Path -LiteralPath (Join-Path $_.FullName "RejjNet2020.exe")) -or
      (Test-Path -LiteralPath (Join-Path $_.FullName "ReJJGSNet2024.exe")) -or
      (Test-Path -LiteralPath (Join-Path $_.FullName "ReJJQDNet2024.exe"))
    )
  } |
  Sort-Object FullName |
  Select-Object -First 1 -ExpandProperty FullName

if (-not $referenceDir) {
  throw "Could not find a software directory containing NPOI.dll under $root"
}

$npoi = Join-Path $referenceDir "NPOI.dll"
$npoiOoxml = Join-Path $referenceDir "NPOI.OOXML.dll"
$npoiOpenXml4Net = Join-Path $referenceDir "NPOI.OpenXml4Net.dll"
$npoiOpenXmlFormats = Join-Path $referenceDir "NPOI.OpenXmlFormats.dll"
$sharpZipLib = Join-Path $referenceDir "ICSharpCode.SharpZipLib.dll"

$csc = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if (-not (Test-Path -LiteralPath $csc)) {
  $csc = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\csc.exe"
}
$out = Join-Path $outDir "ChapterQuotaLibrary.exe"

& $csc /nologo /target:exe /out:$out `
  /reference:System.Data.dll `
  /reference:System.Core.dll `
  /reference:$npoi `
  /reference:$npoiOoxml `
  /reference:$npoiOpenXml4Net `
  /reference:$npoiOpenXmlFormats `
  (Join-Path $PSScriptRoot "ChapterQuotaLibrary.cs")
if ($LASTEXITCODE -ne 0) {
  throw "Build ChapterQuotaLibrary failed with exit code $LASTEXITCODE"
}

Copy-Item -LiteralPath $npoi -Destination $outDir -Force
Copy-Item -LiteralPath $npoiOoxml -Destination $outDir -Force
Copy-Item -LiteralPath $npoiOpenXml4Net -Destination $outDir -Force
Copy-Item -LiteralPath $npoiOpenXmlFormats -Destination $outDir -Force
Copy-Item -LiteralPath $sharpZipLib -Destination $outDir -Force

Write-Host "Built $out"
