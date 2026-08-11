param(
  [string]$OutputDir,
  [string]$SeedDataDir,
  [switch]$Force
)

$ErrorActionPreference = "Stop"

function Copy-RequiredFile {
  param(
    [string]$Source,
    [string]$Destination
  )

  if (-not (Test-Path -LiteralPath $Source)) {
    throw "Missing required source file: $Source"
  }
  $parent = Split-Path -Parent $Destination
  if (-not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
  }
  Copy-Item -LiteralPath $Source -Destination $Destination -Force
}

function Write-Utf8BomFile {
  param(
    [string]$Path,
    [string[]]$Lines
  )

  [System.IO.File]::WriteAllLines($Path, $Lines, (New-Object System.Text.UTF8Encoding($true)))
}

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($OutputDir)) {
  $OutputDir = Join-Path $repoRoot "release\同事插件分层发布"
}
if ([string]::IsNullOrWhiteSpace($SeedDataDir)) {
  $SeedDataDir = Join-Path $repoRoot "铁路基本建设工程投资控制系统2020网络版V0503021201\RecoQuotaData"
}

$repoRoot = (Resolve-Path -LiteralPath $repoRoot).Path
$outputParent = Split-Path -Parent $OutputDir
if (-not (Test-Path -LiteralPath $outputParent)) {
  New-Item -ItemType Directory -Path $outputParent -Force | Out-Null
}
$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
if (-not $OutputDir.StartsWith($repoRoot + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
  throw "OutputDir must stay inside the repository: $OutputDir"
}
if (-not (Test-Path -LiteralPath $SeedDataDir)) {
  throw "Seed data directory does not exist: $SeedDataDir"
}

if (Test-Path -LiteralPath $OutputDir) {
  if (-not $Force) {
    throw "Output already exists. Re-run with -Force to rebuild: $OutputDir"
  }
  Remove-Item -LiteralPath $OutputDir -Recurse -Force
}

$commonDir = Join-Path $OutputDir "00-首次安装公共包"
$expandDir = Join-Path $OutputDir "01-综合扩展功能包"
$quotaDir = Join-Path $OutputDir "02-推荐定额功能包-首次安装"
$updateDir = Join-Path $OutputDir "90-后续更新文件"
foreach ($dir in @($commonDir, $expandDir, $quotaDir, $updateDir)) {
  New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

$binDir = Join-Path $repoRoot "RecoQuotaRecommend\bin"
Copy-RequiredFile -Source (Join-Path $binDir "RecoPluginLoader.dll") -Destination (Join-Path $commonDir "RecoPluginLoader.dll")
Copy-RequiredFile -Source (Join-Path $binDir "0Harmony.dll") -Destination (Join-Path $commonDir "0Harmony.dll")
Copy-RequiredFile -Source (Join-Path $PSScriptRoot "InstallColleaguePlugins.ps1") -Destination (Join-Path $commonDir "InstallPlugins.ps1")
Copy-RequiredFile -Source (Join-Path $PSScriptRoot "InstallColleaguePlugins.cmd") -Destination (Join-Path $commonDir "安装插件.cmd")

Copy-RequiredFile -Source (Join-Path $binDir "RecoExpandPanel.dll") -Destination (Join-Path $expandDir "RecoExpandPanel.dll")
$iconSource = Join-Path $repoRoot "tools\RecoExpandPanel\icons"
if (Test-Path -LiteralPath $iconSource) {
  Copy-Item -LiteralPath $iconSource -Destination (Join-Path $expandDir "RecoExpandPanelIcons") -Recurse -Force
}

Copy-RequiredFile -Source (Join-Path $binDir "RecoQuotaRecommend.dll") -Destination (Join-Path $quotaDir "RecoQuotaRecommend.dll")
$seedTarget = Join-Path $quotaDir "RecoQuotaData"
New-Item -ItemType Directory -Path $seedTarget -Force | Out-Null
foreach ($name in @(
  "quota-index.jsonl",
  "material-index.jsonl",
  "chapter-entries.jsonl",
  "chapter-quota-library.jsonl"
)) {
  Copy-RequiredFile -Source (Join-Path $SeedDataDir $name) -Destination (Join-Path $seedTarget $name)
}
foreach ($name in @("excel-link-units.txt", "fill-templates")) {
  $source = Join-Path $SeedDataDir $name
  if (Test-Path -LiteralPath $source) {
    Copy-Item -LiteralPath $source -Destination (Join-Path $seedTarget $name) -Recurse -Force
  }
}

$expandUpdate = Join-Path $updateDir "综合扩展更新"
$quotaUpdate = Join-Path $updateDir "推荐定额更新"
$commonUpdate = Join-Path $updateDir "公共组件更新"
foreach ($dir in @($expandUpdate, $quotaUpdate, $commonUpdate)) {
  New-Item -ItemType Directory -Path $dir -Force | Out-Null
}
Copy-RequiredFile -Source (Join-Path $binDir "RecoExpandPanel.dll") -Destination (Join-Path $expandUpdate "RecoExpandPanel.dll")
Copy-RequiredFile -Source (Join-Path $binDir "RecoQuotaRecommend.dll") -Destination (Join-Path $quotaUpdate "RecoQuotaRecommend.dll")
Copy-RequiredFile -Source (Join-Path $binDir "RecoPluginLoader.dll") -Destination (Join-Path $commonUpdate "RecoPluginLoader.dll")
Copy-RequiredFile -Source (Join-Path $binDir "0Harmony.dll") -Destination (Join-Path $commonUpdate "0Harmony.dll")

Write-Utf8BomFile -Path (Join-Path $OutputDir "使用说明.txt") -Lines @(
  "首次安装：",
  "1. 关闭 ReJJGSNet2024 和 RejjNet2020。",
  "2. 把 00 公共包及所需功能包内的文件合并复制到软件根目录。",
  "3. 双击安装插件.cmd。脚本只配置 ReJJGSNet2024.exe 和 RejjNet2020.exe。",
  "4. ReJJQDNet2024.exe 不会加载本插件。",
  "",
  "后续更新：",
  "只发送 90-后续更新文件 中对应功能的 DLL，让同事覆盖到软件根目录。",
  "不要在普通更新中发送 RecoQuotaData，以免覆盖同事自己的参考池和模板。",
  "覆盖 DLL 前必须关闭两个目标软件。"
)

$forbidden = Get-ChildItem -LiteralPath $OutputDir -Recurse -File | Where-Object {
  $_.Extension -in @(".cs", ".pdb", ".sln", ".csproj") -or
  $_.Name -eq "deepseek-settings.json" -or
  $_.Name -eq "agent-undo.jsonl" -or
  $_.Name -like "*.log" -or
  $_.Name -like "*.bak*"
}
if ($forbidden) {
  throw "Forbidden files found in release: $($forbidden.FullName -join ', ')"
}

$manifestPath = Join-Path $OutputDir "文件清单-SHA256.txt"
$manifestLines = New-Object System.Collections.Generic.List[string]
Get-ChildItem -LiteralPath $OutputDir -Recurse -File |
  Where-Object { $_.FullName -ne $manifestPath } |
  Sort-Object FullName |
  ForEach-Object {
    $relative = $_.FullName.Substring($OutputDir.Length).TrimStart('\')
    $hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash
    [void]$manifestLines.Add(($hash + "  " + $relative))
  }
Write-Utf8BomFile -Path $manifestPath -Lines $manifestLines.ToArray()

Write-Host "Release built: $OutputDir"
