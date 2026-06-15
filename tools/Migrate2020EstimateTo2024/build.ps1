$ErrorActionPreference = "Stop"

$outDir = Join-Path $PSScriptRoot "bin"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$csc = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\csc.exe"
if (-not (Test-Path -LiteralPath $csc)) {
  throw "Could not find 32-bit csc.exe: $csc"
}

$out = Join-Path $outDir "Migrate2020EstimateTo2024.exe"
$source = Join-Path $PSScriptRoot "Migrate2020EstimateTo2024.cs"

& $csc /nologo /platform:x86 /target:exe /out:$out `
  /reference:System.Data.dll `
  /reference:System.IO.Compression.dll `
  /reference:System.IO.Compression.FileSystem.dll `
  /reference:System.Web.Extensions.dll `
  /reference:System.Xml.dll `
  /reference:System.Xml.Linq.dll `
  $source

if ($LASTEXITCODE -ne 0) {
  throw "Build failed with exit code $LASTEXITCODE"
}

Write-Host "Built $out"
