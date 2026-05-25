$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$softwareDir = Get-ChildItem -Path $root -Directory |
  Where-Object { Test-Path (Join-Path $_.FullName "RejjNet2020.exe") } |
  Select-Object -First 1 -ExpandProperty FullName
if (-not $softwareDir) {
  throw "Could not find software directory containing RejjNet2020.exe under $root"
}

$outDir = Join-Path $PSScriptRoot "bin"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$csc = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$out = Join-Path $outDir "RecoPluginLoader.dll"
$source = Join-Path $PSScriptRoot "AutoLoadDomainManager.cs"

& $csc /nologo /target:library /out:$out $source

Copy-Item -LiteralPath $out -Destination $softwareDir -Force

Write-Host "Built $out"
Write-Host "Deployed to $softwareDir"
