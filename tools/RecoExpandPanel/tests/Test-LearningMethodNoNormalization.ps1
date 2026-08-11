param(
  [Parameter(Mandatory)][string]$QuotaDll,
  [Parameter(Mandatory)][string]$ExpandDll
)

$ErrorActionPreference = 'Stop'

function Invoke-MethodNormalizer {
  param([Reflection.Assembly]$Assembly, [string]$Value)
  $type = $Assembly.GetType('LearningPartitionIdentity', $true)
  $method = $type.GetMethod('NormalizeLearningMethodNo', [Reflection.BindingFlags]'Static,NonPublic')
  if ($method -eq $null) { throw 'NormalizeLearningMethodNo was not found.' }
  [object[]]$arguments = New-Object object[] 1
  $arguments[0] = $Value
  return [string]$method.Invoke($null, $arguments)
}

$assemblies = @(
  [Reflection.Assembly]::Load([IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $QuotaDll).Path)),
  [Reflection.Assembly]::Load([IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $ExpandDll).Path))
)
$method101 = (-join @([char]0x31, [char]0x30, [char]0x31, [char]0x53F7, [char]0x6587, [char]0x4F30, [char]0x7B97))
$method30 = (-join @([char]0x33, [char]0x30, [char]0x53F7, [char]0x6587))
$method2024 = 'TB 10801' + [char]0x2014 + '2024'
$railwayLaw = (-join @([char]0x56FD, [char]0x94C1, [char]0x79D1, [char]0x6CD5))
$cases = @(
  @('101-estimate 2020', $method101),
  @('TB10801-2024', $method2024),
  @('2024', $method2024),
  @('2020', $method30),
  @($railwayLaw, $method30),
  @('unknown', '')
)

foreach ($case in $cases) {
  $left = Invoke-MethodNormalizer $assemblies[0] $case[0]
  $right = Invoke-MethodNormalizer $assemblies[1] $case[0]
  if ($left -ne $case[1] -or $right -ne $case[1] -or $left -ne $right) {
    throw 'Method-number normalization result mismatch.'
  }
}

Write-Host 'PASS B2 method-number normalization: 101 priority, 2024, 2020, unknown'
Write-Host 'PASS cross-DLL method-number outputs are identical'
