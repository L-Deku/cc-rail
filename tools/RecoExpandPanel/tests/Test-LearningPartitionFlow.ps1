param(
  [Parameter(Mandatory)][string]$ExpandDll,
  [switch]$Worker
)

$ErrorActionPreference = 'Stop'

function Assert-True {
  param([bool]$Condition, [string]$Message)
  if (-not $Condition) { throw $Message }
}

function Set-Field {
  param([Type]$Type, $Instance, [string]$Name, $Value)
  $field = $Type.GetField($Name, [Reflection.BindingFlags]'Instance,Public,NonPublic')
  if ($field -eq $null) { throw "Missing field: $Name" }
  $field.SetValue($Instance, $Value)
}

function Get-Field {
  param([Type]$Type, $Instance, [string]$Name)
  $field = $Type.GetField($Name, [Reflection.BindingFlags]'Instance,Public,NonPublic')
  if ($field -eq $null) { throw "Missing field: $Name" }
  return $field.GetValue($Instance)
}

if (-not $Worker) {
  $tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
  $parentTempRoot = Join-Path $tempBase ('RecoLearningPartitionFlow-' + [Guid]::NewGuid().ToString('N'))
  [IO.Directory]::CreateDirectory($parentTempRoot) | Out-Null
  $workerDll = Join-Path $parentTempRoot 'RecoExpandPanel.dll'
  [IO.File]::Copy((Resolve-Path -LiteralPath $ExpandDll).Path, $workerDll, $false)
  $workerExit = 1
  try {
    $scriptPath = (Resolve-Path -LiteralPath $MyInvocation.MyCommand.Path).Path
    & powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath -Worker -ExpandDll $workerDll
    $workerExit = $LASTEXITCODE
  }
  finally {
    $resolvedParentTempRoot = [IO.Path]::GetFullPath($parentTempRoot)
    if (-not $resolvedParentTempRoot.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase)) {
      throw 'Refusing to clean unexpected test directory.'
    }
    if ([IO.Directory]::Exists($resolvedParentTempRoot)) {
      [IO.Directory]::Delete($resolvedParentTempRoot, $true)
    }
  }
  if ($workerExit -ne 0) { throw "Learning partition worker failed: $workerExit" }
  return
}

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $testDir '..\..\..')).Path
$referenceDir = Get-ChildItem -LiteralPath $repoRoot -Directory -Recurse | Where-Object {
  (Test-Path -LiteralPath (Join-Path $_.FullName 'NPOI.dll')) -and (
    (Test-Path -LiteralPath (Join-Path $_.FullName 'RejjNet2020.exe')) -or
    (Test-Path -LiteralPath (Join-Path $_.FullName 'ReJJGSNet2024.exe')) -or
    (Test-Path -LiteralPath (Join-Path $_.FullName 'ReJJQDNet2024.exe')))
} | Select-Object -First 1
if ($referenceDir -eq $null) { throw 'No offline plugin reference directory was found.' }
foreach ($name in @('ICSharpCode.SharpZipLib.dll', 'NPOI.dll', 'NPOI.OOXML.dll', 'NPOI.OpenXml4Net.dll', 'NPOI.OpenXmlFormats.dll')) {
  $path = Join-Path $referenceDir.FullName $name
  if (Test-Path -LiteralPath $path -PathType Leaf) {
    [void][Reflection.Assembly]::LoadFrom($path)
  }
}

$tempRoot = Split-Path -Parent (Resolve-Path -LiteralPath $ExpandDll).Path

try {
  $assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $ExpandDll).Path)
  $formType = $assembly.GetType('RecoNet.FormPanel', $true)
  $groupType = $formType.GetNestedType('MappingFeedbackGroup', [Reflection.BindingFlags]'NonPublic')
  Assert-True ($groupType -ne $null) 'Learning group type was not found.'

  $listType = [Collections.Generic.List``1].MakeGenericType($groupType)
  $group = [Activator]::CreateInstance($groupType, $true)
  Set-Field $groupType $group 'QuantityName' 'partition-order-probe'
  Set-Field $groupType $group 'Method' '101-estimate 2020'
  Set-Field $groupType $group 'MethodNo' ''
  Set-Field $groupType $group 'SoftwarePartition' '2020'
  $groups = [Activator]::CreateInstance($listType)
  [void]$groups.Add($group)

  $prepare = $formType.GetMethod('PrepareLearningGroupsForSql', [Reflection.BindingFlags]'Static,NonPublic')
  [object[]]$prepareArgs = New-Object object[] 1
  $prepareArgs[0] = $groups.PSObject.BaseObject
  [void]$prepare.Invoke($null, $prepareArgs)
  $expected101 = (-join @([char]0x31, [char]0x30, [char]0x31, [char]0x53F7, [char]0x6587, [char]0x4F30, [char]0x7B97))
  Assert-True ((Get-Field $groupType $group 'MethodNo') -eq $expected101) 'MethodNo was calculated after Method lost its 101 identity.'
  Assert-True ((Get-Field $groupType $group 'Method') -eq '2020') 'Legacy Method was not normalized to 2020.'
  Assert-True ((Get-Field $groupType $group 'SoftwarePartition') -eq '2020') 'Existing software partition was overwritten.'

  $unknown = [Activator]::CreateInstance($groupType, $true)
  Set-Field $groupType $unknown 'QuantityName' 'unknown-process-probe'
  Set-Field $groupType $unknown 'Method' '101-estimate 2020'
  $unknownGroups = [Activator]::CreateInstance($listType)
  [void]$unknownGroups.Add($unknown)

  $openAttemptsField = $formType.GetField('learningDbConnectionOpenAttempts', [Reflection.BindingFlags]'Static,NonPublic')
  $beforeOpenAttempts = [long]$openAttemptsField.GetValue($null)
  $record = $formType.GetMethod('RecordMappingGroupsToLearningDb', [Reflection.BindingFlags]'Static,NonPublic')
  [object[]]$recordArgs = New-Object object[] 2
  $recordArgs[0] = $unknownGroups.PSObject.BaseObject
  $recordArgs[1] = 'offline-unknown-process-test'
  [void]$record.Invoke($null, $recordArgs)
  $afterOpenAttempts = [long]$openAttemptsField.GetValue($null)

  $dataDir = Join-Path $tempRoot 'RecoQuotaData'
  $mappingPath = Join-Path $dataDir 'mapping-boxes.jsonl'
  $outboxPath = Join-Path $dataDir 'learning-db-outbox.jsonl'
  $deadLetterPath = Join-Path $dataDir 'learning-db-outbox.dead-letter.jsonl'
  Assert-True (-not (Test-Path -LiteralPath $mappingPath)) 'Unknown process wrote local mapping-boxes data.'
  Assert-True (-not (Test-Path -LiteralPath $outboxPath)) 'Unknown process entered the ordinary retry outbox.'
  Assert-True (-not (Test-Path -LiteralPath $deadLetterPath)) 'Unknown process wrote a local dead-letter record.'
  Assert-True ($beforeOpenAttempts -eq $afterOpenAttempts) 'Unknown process attempted to open SQL.'

  Write-Host 'PASS B3 normalization order preserves 101-estimate MethodNo before Method=2020'
  Write-Host 'PASS SQL-only unknown process: no local mapping, outbox, dead-letter, or SQL attempt'
}
finally {
  # Worker process exits immediately after this block; the parent removes the now-unlocked directory.
}
