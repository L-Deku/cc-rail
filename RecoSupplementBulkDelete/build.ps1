$ErrorActionPreference = "Stop"

# 独立编译并部署补充材料/补充设备批量删除插件 RecoSupplementBulkDelete.dll。
# 日常仍推荐运行 RecoQuotaRecommend\build.ps1，它会一并编译并部署本插件；本脚本只用于单独快速迭代。

$root = Split-Path -Parent $PSScriptRoot
$outDir = Join-Path $root "RecoQuotaRecommend\bin"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

$csc = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$out = Join-Path $outDir "RecoSupplementBulkDelete.dll"
$source = Join-Path $PSScriptRoot "SupplementBulkDeletePlugin.cs"

& $csc /nologo /target:library /out:$out `
  /reference:System.Windows.Forms.dll `
  /reference:System.Drawing.dll `
  /reference:System.Data.dll `
  $source
if ($LASTEXITCODE -ne 0) { throw "Build RecoSupplementBulkDelete failed with exit code $LASTEXITCODE" }
Write-Host "Built $out"

if ($args -contains '-BuildOnly') {
  exit 0
}

$softwareDir = Join-Path $root "2024铁路工程云计价系统网络版V1.0\铁路工程云计价系统网络版V1.0"
if (-not (Test-Path -LiteralPath (Join-Path $softwareDir "ReJJGSNet2024.exe"))) {
  throw "Formal software directory not found: $softwareDir"
}
Copy-Item -LiteralPath $out -Destination $softwareDir -Force
Write-Host "Deployed to $softwareDir"
