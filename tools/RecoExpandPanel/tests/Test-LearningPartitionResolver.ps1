param(
  [string]$QuotaDll = '',
  [string]$ExpandDll = ''
)

$ErrorActionPreference = 'Stop'

function Assert-Equal {
  param($Actual, $Expected, [string]$Message)
  if ($Actual -ne $Expected) { throw $Message }
}

function Invoke-Resolver {
  param([Reflection.Assembly]$Assembly, [string]$ProcessName, [string]$ModuleFileName)
  $type = $Assembly.GetType('LearningPartitionIdentity', $true)
  $method = $type.GetMethod('ResolveFromProcessIdentity', [Reflection.BindingFlags]'Static,NonPublic')
  if ($method -eq $null) { throw 'ResolveFromProcessIdentity was not found.' }
  [object[]]$arguments = New-Object object[] 2
  $arguments[0] = $ProcessName
  $arguments[1] = $ModuleFileName
  return [string]$method.Invoke($null, $arguments)
}

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $testDir '..\..\..')).Path
$source = (Resolve-Path -LiteralPath (Join-Path $repoRoot 'RecoShared\LearningPartitionIdentity.cs')).Path
$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
$tempRoot = Join-Path $tempBase ('RecoPartitionResolver-' + [Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($tempRoot) | Out-Null

try {
  if ([string]::IsNullOrWhiteSpace($QuotaDll) -xor [string]::IsNullOrWhiteSpace($ExpandDll)) {
    throw 'QuotaDll and ExpandDll must be supplied together.'
  }
  if ([string]::IsNullOrWhiteSpace($QuotaDll)) {
    $quotaProbe = Join-Path $tempRoot 'RecoQuotaRecommend.IdentityProbe.dll'
    $expandProbe = Join-Path $tempRoot 'RecoExpandPanel.IdentityProbe.dll'
    foreach ($output in @($quotaProbe, $expandProbe)) {
      & $csc /nologo /target:library /out:$output $source
      if ($LASTEXITCODE -ne 0) { throw "Identity probe compilation failed: $LASTEXITCODE" }
    }
  } else {
    $quotaProbe = (Resolve-Path -LiteralPath $QuotaDll).Path
    $expandProbe = (Resolve-Path -LiteralPath $ExpandDll).Path
  }

  $assemblies = @(
    [Reflection.Assembly]::Load([IO.File]::ReadAllBytes($quotaProbe)),
    [Reflection.Assembly]::Load([IO.File]::ReadAllBytes($expandProbe))
  )
  $cases = @(
    @('RejjNet2020', '', '2020'),
    @('helper', 'C:\apps\ReJJGSNet2024.exe', '2024'),
    @('helper', 'C:\apps\ReJJQDNet2024.exe', '2024'),
    @('unknown-host', 'C:\apps\helper.exe', ''),
    @('unknown-host', 'C:\folder-2024\helper.exe', ''),
    @('RejjNet2020', 'C:\apps\ReJJGSNet2024.exe', '')
  )

  foreach ($case in $cases) {
    $left = Invoke-Resolver $assemblies[0] $case[0] $case[1]
    $right = Invoke-Resolver $assemblies[1] $case[0] $case[1]
    Assert-Equal $left $case[2] ('Unexpected partition for process probe: ' + $case[0])
    Assert-Equal $right $case[2] ('Second DLL probe returned a different partition: ' + $case[0])
    Assert-Equal $left $right ('Cross-DLL partition result drifted: ' + $case[0])
  }

  Write-Host 'PASS B16 resolver: process identity only, unknown/conflict rejected'
  Write-Host 'PASS cross-DLL resolver outputs are identical'
}
finally {
  $resolvedTempRoot = [IO.Path]::GetFullPath($tempRoot)
  if (-not $resolvedTempRoot.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Refusing to clean unexpected test directory.'
  }
  if ([IO.Directory]::Exists($resolvedTempRoot)) {
    [IO.Directory]::Delete($resolvedTempRoot, $true)
  }
}
