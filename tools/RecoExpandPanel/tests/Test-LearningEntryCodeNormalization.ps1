param(
  [string]$QuotaDll = '',
  [string]$ExpandDll = ''
)

$ErrorActionPreference = 'Stop'

function Assert-Equal {
  param($Actual, $Expected, [string]$Message)
  if ($Actual -ne $Expected) { throw $Message }
}

function Invoke-Normalizer {
  param([Reflection.Assembly]$Assembly, $Value)
  $type = $Assembly.GetType('LearningPartitionIdentity', $true)
  $method = $type.GetMethod('NormalizeLearningEntryCode', [Reflection.BindingFlags]'Static,NonPublic')
  if ($method -eq $null) { throw 'NormalizeLearningEntryCode was not found.' }
  [object[]]$arguments = New-Object object[] 1
  $arguments[0] = $Value
  return [string]$method.Invoke($null, $arguments)
}

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $testDir '..\..\..')).Path
$source = (Resolve-Path -LiteralPath (Join-Path $repoRoot 'RecoShared\LearningPartitionIdentity.cs')).Path
$csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
$tempRoot = Join-Path $tempBase ('RecoEntryCodeNormalization-' + [Guid]::NewGuid().ToString('N'))
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

  $cases = New-Object System.Collections.ArrayList
  [void]$cases.Add(@(' 0616-02-01 ', '0616-02-01'))
  [void]$cases.Add(@('0010-02', '0010-02'))
  [void]$cases.Add(@('12', '12'))
  foreach ($dash in @([char]0x2010, [char]0x2011, [char]0x2012, [char]0x2013, [char]0x2014, [char]0x2212, [char]0xFF0D)) {
    [void]$cases.Add(@(('0616' + $dash + '02-01'), '0616-02-01'))
  }
  foreach ($invalid in @('06 16-02', '0616--02', '0616-', '2', '3', '4', '5', '', $null)) {
    [void]$cases.Add(@($invalid, ''))
  }

  foreach ($case in $cases) {
    $left = Invoke-Normalizer $assemblies[0] $case[0]
    $right = Invoke-Normalizer $assemblies[1] $case[0]
    Assert-Equal $left $case[1] 'Unexpected normalized entry code.'
    Assert-Equal $right $case[1] 'Second DLL probe normalized the entry code differently.'
    Assert-Equal $left $right 'Cross-DLL entry-code normalization drifted.'
  }

  Write-Host 'PASS T25 pure normalization samples'
  Write-Host 'PASS cross-DLL entry-code outputs are identical'
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
