param(
  [string]$SoftwareDir = $PSScriptRoot
)

$ErrorActionPreference = "Stop"

function New-MinimalConfig {
  param(
    [string]$Path,
    [string]$Framework
  )

  $lines = @(
    '<?xml version="1.0" encoding="utf-8"?>',
    '<configuration>',
    '  <startup useLegacyV2RuntimeActivationPolicy="true">',
    ('    <supportedRuntime version="v4.0" sku="{0}" />' -f $Framework),
    '  </startup>',
    '  <runtime />',
    '</configuration>'
  )
  [System.IO.File]::WriteAllLines($Path, $lines, (New-Object System.Text.UTF8Encoding($true)))
}

function Ensure-PluginConfig {
  param(
    [string]$ConfigPath,
    [string[]]$TemplatePaths,
    [string]$Framework
  )

  if (-not (Test-Path -LiteralPath $ConfigPath)) {
    $template = $TemplatePaths | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
    if ($template) {
      Copy-Item -LiteralPath $template -Destination $ConfigPath
    } else {
      New-MinimalConfig -Path $ConfigPath -Framework $Framework
    }
  }

  $xml = New-Object System.Xml.XmlDocument
  $xml.PreserveWhitespace = $true
  $xml.Load($ConfigPath)
  if ($null -eq $xml.configuration) {
    throw "Invalid config file: $ConfigPath"
  }

  $changed = $false
  $runtime = $xml.configuration.runtime
  if ($null -eq $runtime) {
    $runtime = $xml.CreateElement("runtime")
    [void]$xml.configuration.AppendChild($runtime)
    $changed = $true
  }

  $managerAssembly = $runtime.appDomainManagerAssembly
  if ($null -eq $managerAssembly) {
    $managerAssembly = $xml.CreateElement("appDomainManagerAssembly")
    [void]$runtime.PrependChild($managerAssembly)
    $changed = $true
  }
  $assemblyValue = "RecoPluginLoader, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null"
  if ($managerAssembly.GetAttribute("value") -ne $assemblyValue) {
    $managerAssembly.SetAttribute("value", $assemblyValue)
    $changed = $true
  }

  $managerType = $runtime.appDomainManagerType
  if ($null -eq $managerType) {
    $managerType = $xml.CreateElement("appDomainManagerType")
    [void]$runtime.AppendChild($managerType)
    $changed = $true
  }
  $typeValue = "RecoPluginLoader.AutoLoadDomainManager"
  if ($managerType.GetAttribute("value") -ne $typeValue) {
    $managerType.SetAttribute("value", $typeValue)
    $changed = $true
  }

  if (-not $changed) {
    return "unchanged"
  }

  $timestamp = Get-Date -Format "yyyyMMdd-HHmmss"
  $backupPath = $ConfigPath + ".pre-reco-plugin-" + $timestamp + ".bak"
  Copy-Item -LiteralPath $ConfigPath -Destination $backupPath
  $xml.Save($ConfigPath)
  return "updated; backup=" + (Split-Path -Leaf $backupPath)
}

if ([string]::IsNullOrWhiteSpace($SoftwareDir)) {
  throw "SoftwareDir is empty."
}
if (-not (Test-Path -LiteralPath $SoftwareDir)) {
  throw "Software directory does not exist: $SoftwareDir"
}
$SoftwareDir = (Resolve-Path -LiteralPath $SoftwareDir).Path

$required = @("RecoPluginLoader.dll", "0Harmony.dll")
foreach ($name in $required) {
  if (-not (Test-Path -LiteralPath (Join-Path $SoftwareDir $name))) {
    throw "Missing required file: $name"
  }
}
if (-not (Test-Path -LiteralPath (Join-Path $SoftwareDir "RecoExpandPanel.dll")) -and
    -not (Test-Path -LiteralPath (Join-Path $SoftwareDir "RecoQuotaRecommend.dll"))) {
  throw "No feature DLL found. Copy at least one feature package into the software directory."
}

$targets = @(
  [pscustomobject]@{
    Exe = "ReJJGSNet2024.exe"
    Process = "ReJJGSNet2024"
    Framework = ".NETFramework,Version=v4.6.2"
    Templates = @("RecoNet2024.exe.config")
  },
  [pscustomobject]@{
    Exe = "RejjNet2020.exe"
    Process = "RejjNet2020"
    Framework = ".NETFramework,Version=v4.6"
    Templates = @("RecoNet2020.exe.config", "RecoNet2020.vshost.exe.config")
  }
)

$present = @($targets | Where-Object { Test-Path -LiteralPath (Join-Path $SoftwareDir $_.Exe) })
if ($present.Count -eq 0) {
  throw "Neither ReJJGSNet2024.exe nor RejjNet2020.exe was found in: $SoftwareDir"
}

foreach ($target in $present) {
  if (Get-Process -Name $target.Process -ErrorAction SilentlyContinue) {
    throw "Close $($target.Exe) before installing the plugin."
  }
}

Write-Host "Software directory: $SoftwareDir"
foreach ($target in $present) {
  $configPath = Join-Path $SoftwareDir ($target.Exe + ".config")
  $templates = @($target.Templates | ForEach-Object { Join-Path $SoftwareDir $_ })
  $result = Ensure-PluginConfig -ConfigPath $configPath -TemplatePaths $templates -Framework $target.Framework
  Write-Host ("Configured {0}: {1}" -f $target.Exe, $result)
}

Write-Host "Skipped by design: ReJJQDNet2024.exe"
Write-Host "Plugin installation completed."
