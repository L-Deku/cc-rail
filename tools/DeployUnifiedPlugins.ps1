param(
  [switch]$Deploy,
  [switch]$SkipBuild,
  [switch]$AllowNonMain
)

$ErrorActionPreference = "Stop"

function New-UString {
  param([int[]]$Codes)
  $chars = foreach ($code in $Codes) { [char]$code }
  return -join $chars
}

function Write-Section {
  param([string]$Title)
  Write-Host ""
  Write-Host "== $Title =="
}

function Get-Hash {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) {
    return ""
  }
  return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

function Get-Hash16 {
  param([string]$Path)
  $hash = Get-Hash -Path $Path
  if ([string]::IsNullOrEmpty($hash)) {
    return "missing"
  }
  return $hash.Substring(0, 16)
}

function Read-DllText {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) {
    return ""
  }
  $bytes = [System.IO.File]::ReadAllBytes($Path)
  return [System.Text.Encoding]::Unicode.GetString($bytes) + "`n" + [System.Text.Encoding]::UTF8.GetString($bytes)
}

function Test-AllMarkers {
  param(
    [string]$Path,
    [string[]]$Markers
  )
  $text = Read-DllText -Path $Path
  foreach ($marker in $Markers) {
    if (-not $text.Contains($marker)) {
      return $false
    }
  }
  return $true
}

function Test-AnyMarker {
  param(
    [string]$Path,
    [string[]]$Markers
  )
  $text = Read-DllText -Path $Path
  foreach ($marker in $Markers) {
    if ($text.Contains($marker)) {
      return $true
    }
  }
  return $false
}

function Assert-SourceDlls {
  param([string]$BinDir)

  $required = @(
    "RecoQuotaRecommend.dll",
    "RecoExpandPanel.dll",
    "RecoPluginLoader.dll",
    "0Harmony.dll"
  )
  foreach ($name in $required) {
    $path = Join-Path $BinDir $name
    if (-not (Test-Path -LiteralPath $path)) {
      throw "Missing build output: $path"
    }
  }

  $expand = Join-Path $BinDir "RecoExpandPanel.dll"
  $quota = Join-Path $BinDir "RecoQuotaRecommend.dll"
  $loader = Join-Path $BinDir "RecoPluginLoader.dll"

  if (-not (Test-AllMarkers -Path $expand -Markers @("ExcelInstantQuantityInputRuntime", "AgentChat", "ExcelLinkPanel"))) {
    throw "RecoExpandPanel.dll is missing Excel instant input, agent chat, or Excel link markers."
  }
  if (-not (Test-AnyMarker -Path $expand -Markers @("TemplateFillPanel", "Template fill"))) {
    throw "RecoExpandPanel.dll is missing template-fill markers."
  }
  if (-not (Test-AllMarkers -Path $quota -Markers @("ReferenceQuotaPoolFeature", "QuotaInlineSearchFeature", "ApplySfDetails", "IsAllowedReferencePoolItem"))) {
    throw "RecoQuotaRecommend.dll is missing reference-pool/candidate/SF/original-quota markers."
  }
  if (-not (Test-AllMarkers -Path $loader -Markers @("AutoLoadDomainManager"))) {
    throw "RecoPluginLoader.dll is missing loader marker."
  }
}

function Test-SoftwareDir {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath (Join-Path $Path "NPOI.dll"))) {
    return $false
  }
  foreach ($exe in @("RejjNet2020.exe", "ReJJGSNet2024.exe", "ReJJQDNet2024.exe")) {
    if (Test-Path -LiteralPath (Join-Path $Path $exe)) {
      return $true
    }
  }
  return $false
}

function Find-SoftwareDirs {
  param(
    [string]$Root,
    [string]$Group
  )

  $result = New-Object System.Collections.Generic.List[object]
  if (-not (Test-Path -LiteralPath $Root)) {
    return $result
  }

  $candidates = New-Object System.Collections.Generic.List[object]
  [void]$candidates.Add((Get-Item -LiteralPath $Root))
  Get-ChildItem -LiteralPath $Root -Directory -Recurse -ErrorAction SilentlyContinue |
    ForEach-Object { [void]$candidates.Add($_) }

  $seen = @{}
  foreach ($candidate in $candidates) {
    if ($seen.ContainsKey($candidate.FullName)) {
      continue
    }
    $seen[$candidate.FullName] = $true
    if (Test-SoftwareDir -Path $candidate.FullName) {
      [void]$result.Add([pscustomobject]@{
        Group = $Group
        Path = $candidate.FullName
      })
    }
  }
  return $result
}

function Ensure-PluginConfig {
  param(
    [string]$ConfigPath,
    [string]$TemplatePath
  )

  if (-not (Test-Path -LiteralPath $ConfigPath)) {
    if (-not [string]::IsNullOrWhiteSpace($TemplatePath) -and (Test-Path -LiteralPath $TemplatePath)) {
      Copy-Item -LiteralPath $TemplatePath -Destination $ConfigPath -Force
    } else {
      Set-Content -LiteralPath $ConfigPath -Encoding UTF8 -Value @(
        '<?xml version="1.0" encoding="utf-8"?>',
        '<configuration>',
        '  <startup useLegacyV2RuntimeActivationPolicy="true">',
        '    <supportedRuntime version="v4.0" sku=".NETFramework,Version=v4.6" />',
        '  </startup>',
        '  <runtime />',
        '</configuration>'
      )
    }
  }

  $backupPath = $ConfigPath + ".pre-plugin-loader.bak"
  if (-not (Test-Path -LiteralPath $backupPath)) {
    Copy-Item -LiteralPath $ConfigPath -Destination $backupPath -Force
  }

  [xml]$xml = Get-Content -LiteralPath $ConfigPath -Raw
  if ($null -eq $xml.configuration) {
    throw "Invalid config file: $ConfigPath"
  }

  $runtime = $xml.configuration.runtime
  if ($null -eq $runtime) {
    $runtime = $xml.CreateElement("runtime")
    [void]$xml.configuration.AppendChild($runtime)
  }

  $managerAssembly = $runtime.appDomainManagerAssembly
  if ($null -eq $managerAssembly) {
    $managerAssembly = $xml.CreateElement("appDomainManagerAssembly")
    [void]$runtime.PrependChild($managerAssembly)
  }
  $managerAssembly.SetAttribute("value", "RecoPluginLoader, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null")

  $managerType = $runtime.appDomainManagerType
  if ($null -eq $managerType) {
    $managerType = $xml.CreateElement("appDomainManagerType")
    $insertAfter = $runtime.appDomainManagerAssembly
    if ($null -ne $insertAfter -and $null -ne $insertAfter.NextSibling) {
      [void]$runtime.InsertBefore($managerType, $insertAfter.NextSibling)
    } else {
      [void]$runtime.AppendChild($managerType)
    }
  }
  $managerType.SetAttribute("value", "RecoPluginLoader.AutoLoadDomainManager")

  $xml.Save($ConfigPath)
}

function Ensure-TargetConfigs {
  param([string]$SoftwareDir)

  $template = Join-Path $SoftwareDir "RecoNet2024.exe.config"
  if (-not (Test-Path -LiteralPath $template)) {
    $template = Join-Path $SoftwareDir "RejjNet2020.exe.config"
  }

  foreach ($exeName in @("RejjNet2020.exe", "ReJJGSNet2024.exe", "ReJJQDNet2024.exe")) {
    $exePath = Join-Path $SoftwareDir $exeName
    if (Test-Path -LiteralPath $exePath) {
      Ensure-PluginConfig -ConfigPath ($exePath + ".config") -TemplatePath $template
    }
  }
}

function Test-TargetConfig {
  param([string]$SoftwareDir)

  $configs = New-Object System.Collections.Generic.List[string]
  foreach ($exeName in @("RejjNet2020.exe", "ReJJGSNet2024.exe", "ReJJQDNet2024.exe")) {
    $exePath = Join-Path $SoftwareDir $exeName
    if (Test-Path -LiteralPath $exePath) {
      [void]$configs.Add($exePath + ".config")
    }
  }
  if ($configs.Count -eq 0) {
    return $false
  }
  foreach ($config in $configs) {
    if (-not (Test-Path -LiteralPath $config)) {
      return $false
    }
    $text = Get-Content -LiteralPath $config -Raw
    if (-not $text.Contains("RecoPluginLoader.AutoLoadDomainManager")) {
      return $false
    }
  }
  return $true
}

function Copy-PluginFile {
  param(
    [string]$Source,
    [string]$DestinationDir,
    [string]$Timestamp
  )

  $destination = Join-Path $DestinationDir (Split-Path -Leaf $Source)
  $sourceHash = Get-Hash -Path $Source
  $destHash = Get-Hash -Path $destination
  if ($sourceHash -eq $destHash) {
    return "same"
  }

  if (Test-Path -LiteralPath $destination) {
    $backup = $destination + ".bak-unified-" + $Timestamp
    Copy-Item -LiteralPath $destination -Destination $backup -Force
  }
  Copy-Item -LiteralPath $Source -Destination $destination -Force
  return "copied"
}

$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = Split-Path -Parent $scriptRoot
$sourceDir = Join-Path $repoRoot "RecoQuotaRecommend"
$buildScript = Join-Path $sourceDir "build.ps1"
$binDir = Join-Path $sourceDir "bin"
$aiRoot = Split-Path -Parent $repoRoot

$specialLineName = New-UString @(0x81EA, 0x52A8, 0x9884, 0x7B97, 0x4E13, 0x7528, 0x7EBF)
$xuName = (New-UString @(0x94C1, 0x8DEF, 0x5DE5, 0x7A0B, 0x4E91, 0x8BA1, 0x4EF7, 0x7CFB, 0x7EDF, 0x7F51, 0x7EDC, 0x7248)) + "V1.0-" + (New-UString @(0x5F90, 0x603B))

$managedRoots = @(
  [pscustomobject]@{ Group = "source-auto"; Path = $repoRoot },
  [pscustomobject]@{ Group = "special-line"; Path = (Join-Path $aiRoot $specialLineName) },
  [pscustomobject]@{ Group = "xu"; Path = (Join-Path $aiRoot $xuName) }
)

Write-Section "Source"
Write-Host "RepoRoot: $repoRoot"
Write-Host "BuildScript: $buildScript"
Write-Host "BinDir: $binDir"

if (-not (Test-Path -LiteralPath $buildScript)) {
  throw "Missing build script: $buildScript"
}

$branch = (& git -C $repoRoot rev-parse --abbrev-ref HEAD 2>$null)
if ($LASTEXITCODE -eq 0) {
  Write-Host "GitBranch: $branch"
  if ($branch -ne "main" -and -not $AllowNonMain) {
    throw "This unified deploy script must run from branch main. Use -AllowNonMain only for emergency testing."
  }
  $gitStatus = (& git -C $repoRoot status --short --branch 2>$null)
  foreach ($line in $gitStatus) {
    Write-Host "GitStatus: $line"
  }
} else {
  Write-Warning "Git branch check failed; continuing without branch validation."
}

if (-not $SkipBuild) {
  Write-Section "Build"
  & powershell.exe -ExecutionPolicy Bypass -File $buildScript
  if ($LASTEXITCODE -ne 0) {
    throw "Build failed with exit code $LASTEXITCODE"
  }
} else {
  Write-Section "Build"
  Write-Host "Skipped build because -SkipBuild was supplied."
}

Assert-SourceDlls -BinDir $binDir

$pluginFiles = @(
  "RecoQuotaRecommend.dll",
  "RecoExpandPanel.dll",
  "RecoPluginLoader.dll",
  "0Harmony.dll"
)

$sourceHashes = @{}
foreach ($file in $pluginFiles) {
  $sourcePath = Join-Path $binDir $file
  $sourceHashes[$file] = Get-Hash -Path $sourcePath
}

Write-Section "Targets"
$targets = New-Object System.Collections.Generic.List[object]
foreach ($root in $managedRoots) {
  $found = Find-SoftwareDirs -Root $root.Path -Group $root.Group
  foreach ($target in $found) {
    [void]$targets.Add($target)
  }
}

if ($targets.Count -eq 0) {
  throw "No supported software directories were found."
}

$targets | Sort-Object Group, Path | Format-Table Group, Path -AutoSize

if ($Deploy) {
  Write-Section "Deploy"
  $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
  foreach ($target in ($targets | Sort-Object Group, Path)) {
    Write-Host "Deploying to [$($target.Group)] $($target.Path)"
    foreach ($file in $pluginFiles) {
      $sourcePath = Join-Path $binDir $file
      $result = Copy-PluginFile -Source $sourcePath -DestinationDir $target.Path -Timestamp $timestamp
      Write-Host "  $file : $result"
    }
    Ensure-TargetConfigs -SoftwareDir $target.Path
  }
} else {
  Write-Section "Deploy"
  Write-Host "Dry run only. Add -Deploy to copy plugin DLLs to all targets."
}

Write-Section "Verification"
$rows = New-Object System.Collections.Generic.List[object]
foreach ($target in ($targets | Sort-Object Group, Path)) {
  $quotaPath = Join-Path $target.Path "RecoQuotaRecommend.dll"
  $expandPath = Join-Path $target.Path "RecoExpandPanel.dll"
  $loaderPath = Join-Path $target.Path "RecoPluginLoader.dll"
  foreach ($file in $pluginFiles) {
    $path = Join-Path $target.Path $file
    $hash = Get-Hash -Path $path
    [void]$rows.Add([pscustomobject]@{
      Group = $target.Group
      File = $file
      Hash16 = Get-Hash16 -Path $path
      SameAsSource = (-not [string]::IsNullOrEmpty($hash) -and $hash -eq $sourceHashes[$file])
      Path = $target.Path
    })
  }

  $expandOk = (Test-AllMarkers -Path $expandPath -Markers @("ExcelInstantQuantityInputRuntime", "AgentChat", "ExcelLinkPanel")) -and
    (Test-AnyMarker -Path $expandPath -Markers @("TemplateFillPanel", "Template fill"))
  $quotaOk = Test-AllMarkers -Path $quotaPath -Markers @("ReferenceQuotaPoolFeature", "QuotaInlineSearchFeature", "ApplySfDetails", "IsAllowedReferencePoolItem")
  $loaderOk = Test-AllMarkers -Path $loaderPath -Markers @("AutoLoadDomainManager")
  $configOk = Test-TargetConfig -SoftwareDir $target.Path
  Write-Host ("FeatureCheck [{0}] Expand={1} Quota={2} Loader={3} Config={4} Path={5}" -f $target.Group, $expandOk, $quotaOk, $loaderOk, $configOk, $target.Path)
}

$rows | Format-Table Group, File, Hash16, SameAsSource, Path -AutoSize

$missing = $rows | Where-Object { -not $_.SameAsSource }
if ($Deploy) {
  if ($missing.Count -gt 0) {
    throw "Deploy verification failed: some plugin files do not match source hashes."
  }
  Write-Host "Unified deployment completed successfully."
} else {
  if ($missing.Count -gt 0) {
    Write-Warning "Dry run found plugin files that do not match source. Run again with -Deploy after closing the target software."
  } else {
    Write-Host "All target plugin files already match source hashes."
  }
}
