$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$outDir = Join-Path $PSScriptRoot "bin"
New-Item -ItemType Directory -Path $outDir -Force | Out-Null

function Add-SoftwareTarget {
  param(
    [System.Collections.ArrayList]$Targets,
    [string]$Path
  )

  if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) {
    return
  }

  $resolved = (Resolve-Path -LiteralPath $Path).Path
  if (-not ($Targets -contains $resolved)) {
    [void]$Targets.Add($resolved)
  }
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
  if ($xml.configuration -eq $null) {
    throw "Invalid config file: $ConfigPath"
  }

  $runtime = $xml.configuration.runtime
  if ($runtime -eq $null) {
    $runtime = $xml.CreateElement("runtime")
    [void]$xml.configuration.AppendChild($runtime)
  }

  $managerAssembly = $runtime.appDomainManagerAssembly
  if ($managerAssembly -eq $null) {
    $managerAssembly = $xml.CreateElement("appDomainManagerAssembly")
    [void]$runtime.PrependChild($managerAssembly)
  }
  $managerAssembly.SetAttribute("value", "RecoPluginLoader, Version=0.0.0.0, Culture=neutral, PublicKeyToken=null")

  $managerType = $runtime.appDomainManagerType
  if ($managerType -eq $null) {
    $managerType = $xml.CreateElement("appDomainManagerType")
    $insertAfter = $runtime.appDomainManagerAssembly
    if ($insertAfter -ne $null -and $insertAfter.NextSibling -ne $null) {
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

  $exeNames = @("RejjNet2020.exe", "ReJJGSNet2024.exe", "ReJJQDNet2024.exe")
  foreach ($exeName in $exeNames) {
    $exePath = Join-Path $SoftwareDir $exeName
    if (Test-Path -LiteralPath $exePath) {
      Ensure-PluginConfig -ConfigPath ($exePath + ".config") -TemplatePath $template
    }
  }
}

function Assert-NativeSuccess {
  param([string]$Step)

  if ($LASTEXITCODE -ne 0) {
    throw "$Step failed with exit code $LASTEXITCODE"
  }
}

function Assert-DllContainsText {
  param(
    [string]$Path,
    [string]$FeatureName,
    [string[]]$Markers
  )

  if (-not (Test-Path -LiteralPath $Path)) {
    throw "Missing DLL for feature check: $Path"
  }

  $bytes = [System.IO.File]::ReadAllBytes($Path)
  $text = [System.Text.Encoding]::Unicode.GetString($bytes) + "`n" + [System.Text.Encoding]::UTF8.GetString($bytes)
  foreach ($marker in $Markers) {
    if ($text.Contains($marker)) {
      return
    }
  }

  throw "Build output $Path does not contain required feature '$FeatureName'. Check tools\RecoExpandPanel source files before deploying."
}

$targets = New-Object System.Collections.ArrayList
Get-ChildItem -LiteralPath $root -Directory -Recurse |
  Where-Object {
    (Test-Path -LiteralPath (Join-Path $_.FullName "NPOI.dll")) -and
    (
      (Test-Path -LiteralPath (Join-Path $_.FullName "RejjNet2020.exe")) -or
      (Test-Path -LiteralPath (Join-Path $_.FullName "ReJJGSNet2024.exe")) -or
      (Test-Path -LiteralPath (Join-Path $_.FullName "ReJJQDNet2024.exe"))
    )
  } |
  Sort-Object FullName |
  ForEach-Object { Add-SoftwareTarget -Targets $targets -Path $_.FullName }

if ($targets.Count -eq 0) {
  throw "Could not find any supported software directory under $root"
}

$referenceDir = $targets | Where-Object { Test-Path -LiteralPath (Join-Path $_ "NPOI.dll") } | Select-Object -First 1
if (-not $referenceDir) {
  throw "Could not find NPOI.dll in any target software directory"
}

$csc = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
$importerOut = Join-Path $outDir "QuotaLearningImporter.exe"
$quotaOut = Join-Path $outDir "RecoQuotaRecommend.dll"
$loaderOut = Join-Path $outDir "RecoPluginLoader.dll"
$expandOut = Join-Path $outDir "RecoExpandPanel.dll"

$npoi = Join-Path $referenceDir "NPOI.dll"
$npoiOoxml = Join-Path $referenceDir "NPOI.OOXML.dll"
$npoiOpenXml4Net = Join-Path $referenceDir "NPOI.OpenXml4Net.dll"
$npoiOpenXmlFormats = Join-Path $referenceDir "NPOI.OpenXmlFormats.dll"
$sharpZipLib = Join-Path $referenceDir "ICSharpCode.SharpZipLib.dll"
$harmony = Join-Path $PSScriptRoot "packages\Lib.Harmony.2.3.3\package\lib\net452\0Harmony.dll"
if (-not (Test-Path -LiteralPath $harmony)) {
  throw "Could not find Harmony runtime: $harmony"
}
$systemRuntime = Get-ChildItem -LiteralPath "$env:WINDIR\Microsoft.NET\assembly\GAC_MSIL\System.Runtime" -Recurse -Filter "System.Runtime.dll" |
  Select-Object -First 1 -ExpandProperty FullName
if (-not $systemRuntime) {
  throw "Could not find System.Runtime.dll"
}
$quotaSources = @(
  (Join-Path $PSScriptRoot "QuotaRecommendPanel.cs"),
  (Join-Path $PSScriptRoot "RecoQuotaInlineSearchFeature.cs"),
  (Join-Path $PSScriptRoot "RecoReferenceQuotaPoolFeature.cs")
)
$expandSources = Get-ChildItem -LiteralPath (Join-Path $root "tools\RecoExpandPanel") -Filter "*.cs" |
  Sort-Object Name |
  ForEach-Object { $_.FullName }

& $csc /nologo /target:exe /out:$importerOut `
  /reference:$npoi `
  /reference:$npoiOoxml `
  /reference:$npoiOpenXml4Net `
  /reference:$npoiOpenXmlFormats `
  (Join-Path $PSScriptRoot "QuotaLearningImporter.cs")
Assert-NativeSuccess "Build QuotaLearningImporter"

& $csc /nologo /target:library /out:$quotaOut `
  /reference:System.Windows.Forms.dll `
  /reference:System.Drawing.dll `
  /reference:System.Data.dll `
  /reference:System.Web.Extensions.dll `
  /reference:$systemRuntime `
  /reference:$harmony `
  $quotaSources
Assert-NativeSuccess "Build RecoQuotaRecommend"

& $csc /nologo /target:library /out:$loaderOut `
  (Join-Path $root "RecoPluginLoader\AutoLoadDomainManager.cs")
Assert-NativeSuccess "Build RecoPluginLoader"

& $csc /nologo /target:library /out:$expandOut `
  /reference:System.Windows.Forms.dll `
  /reference:System.Drawing.dll `
  /reference:System.Data.dll `
  /reference:System.Xml.dll `
  /reference:System.Xml.Linq.dll `
  /reference:System.Core.dll `
  /reference:System.Web.Extensions.dll `
  /reference:System.IO.Compression.dll `
  /reference:System.IO.Compression.FileSystem.dll `
  /reference:$npoi `
  /reference:$npoiOoxml `
  /reference:$npoiOpenXml4Net `
  /reference:$npoiOpenXmlFormats `
  /reference:$sharpZipLib `
  $expandSources
Assert-NativeSuccess "Build RecoExpandPanel"
Assert-DllContainsText -Path $expandOut -FeatureName "Excel quantity instant input" -Markers @("ExcelInstantQuantityInputRuntime")
Assert-DllContainsText -Path $expandOut -FeatureName "Template fill" -Markers @("TemplateFillPanel", "模板铺量")
Assert-DllContainsText -Path $expandOut -FeatureName "Agent chat" -Markers @("AgentChat")

Copy-Item -LiteralPath $npoi -Destination $outDir -Force
Copy-Item -LiteralPath $npoiOoxml -Destination $outDir -Force
Copy-Item -LiteralPath $npoiOpenXml4Net -Destination $outDir -Force
Copy-Item -LiteralPath $npoiOpenXmlFormats -Destination $outDir -Force
Copy-Item -LiteralPath $sharpZipLib -Destination $outDir -Force
Copy-Item -LiteralPath $harmony -Destination $outDir -Force

foreach ($softwareDir in $targets) {
  Copy-Item -LiteralPath $loaderOut -Destination $softwareDir -Force
  Copy-Item -LiteralPath $expandOut -Destination $softwareDir -Force
  Copy-Item -LiteralPath $quotaOut -Destination $softwareDir -Force
  Copy-Item -LiteralPath $harmony -Destination $softwareDir -Force

  $iconSource = Join-Path $root "tools\RecoExpandPanel\icons"
  if (Test-Path -LiteralPath $iconSource) {
    $iconTarget = Join-Path $softwareDir "RecoExpandPanelIcons"
    New-Item -ItemType Directory -Path $iconTarget -Force | Out-Null
    Copy-Item -Path (Join-Path $iconSource "*") -Destination $iconTarget -Force
  }

  $dataTarget = Join-Path $softwareDir "RecoQuotaData"
  New-Item -ItemType Directory -Path $dataTarget -Force | Out-Null
  if ($softwareDir -notmatch "2024") {
    $dataSource = Join-Path $root "RecoQuotaData"
    if (Test-Path -LiteralPath $dataSource) {
      Copy-Item -Path (Join-Path $dataSource "*") -Destination $dataTarget -Force
    }
  } else {
    $mappingPath = Join-Path $dataTarget "mapping-boxes.jsonl"
    if (-not (Test-Path -LiteralPath $mappingPath)) {
      New-Item -ItemType File -Path $mappingPath -Force | Out-Null
    }
    # 2024 targets still need the chapter-entry library files (kept per-method inside the files)
    foreach ($chapterFile in @("chapter-entries.jsonl", "chapter-quota-library.jsonl")) {
      $chapterSource = Join-Path (Join-Path $root "RecoQuotaData") $chapterFile
      if (Test-Path -LiteralPath $chapterSource) {
        Copy-Item -LiteralPath $chapterSource -Destination $dataTarget -Force
      }
    }
  }

  Ensure-TargetConfigs -SoftwareDir $softwareDir
  Write-Host "Deployed plugins to $softwareDir"
}

Write-Host "Built $importerOut"
Write-Host "Built $quotaOut"
Write-Host "Built $loaderOut"
Write-Host "Built $expandOut"
