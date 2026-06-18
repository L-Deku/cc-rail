$ErrorActionPreference = "Stop"

$outDir = Join-Path $PSScriptRoot "bin"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$cecil = Join-Path $PSScriptRoot "packages\Mono.Cecil.0.11.5\lib\net40\Mono.Cecil.dll"
if (-not (Test-Path -LiteralPath $cecil)) {
  throw "Mono.Cecil.dll not found: $cecil"
}

$csc = Join-Path $env:WINDIR "Microsoft.NET\Framework\v4.0.30319\csc.exe"
$out = Join-Path $outDir "InspectRecoValidation.exe"

& $csc /nologo /target:exe /out:$out /reference:$cecil (Join-Path $PSScriptRoot "Program.cs")
if ($LASTEXITCODE -ne 0) {
  throw "Build failed with exit code $LASTEXITCODE"
}

Copy-Item -LiteralPath $cecil -Destination $outDir -Force
Write-Host "Built $out"
