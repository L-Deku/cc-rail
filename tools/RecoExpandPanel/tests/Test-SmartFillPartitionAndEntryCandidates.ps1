param(
  [Parameter(Mandatory)][string]$ExpandDll
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

function Set-Field {
  param([Type]$Type, $Instance, [string]$Name, $Value)
  $field = $Type.GetField($Name, [Reflection.BindingFlags]'Public,NonPublic,Instance,Static')
  if ($field -eq $null) { throw "Missing field: $Name" }
  $field.SetValue($Instance, $Value)
}

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..')).Path
$source = [IO.File]::ReadAllText((Join-Path $repoRoot 'tools\RecoExpandPanel\SmartFillFeature.cs'), [Text.Encoding]::UTF8)
foreach ($forbidden in @("method=''", "method = ''", '名称学习命中，全库兜底', '名称兼容命中，全库兜底', '"全库兜底"')) {
  if ($source.Contains($forbidden)) { throw "SmartFill 仍含禁止的空办法/全库兜底：$forbidden" }
}
foreach ($required in @(
  'm.software_partition=@software_partition',
  'software_partition=@software_partition AND method_no=@method_no',
  'r.software_partition=@software_partition AND r.method_no=@method_no',
  'const int MaxEntryCombinations = 16')) {
  if (-not $source.Contains($required)) { throw "SmartFill 缺少分区或候选上限门禁：$required" }
}

$assembly = [Reflection.Assembly]::Load([IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $ExpandDll).Path))
$formType = $assembly.GetType('RecoNet.FormPanel', $true)
$nested = [Reflection.BindingFlags]'Public,NonPublic'
$all = [Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$snapshotType = $formType.GetNestedType('SmartLearningSnapshot', $nested)
$scopeType = $formType.GetNestedType('SmartLearningScope', $nested)
$entryType = $formType.GetNestedType('SmartMapEntry', $nested)
$targetType = $formType.GetNestedType('SmartBoxTarget', $nested)
$statType = $formType.GetNestedType('SmartEntryStat', $nested)
$scoreType = $formType.GetNestedType('SmartMapCandidateScore', $nested)
$resolve = $formType.GetMethod('ResolveSmartTargetEntryCombinations', $all)
$combinationKey = $formType.GetMethod('BuildSmartEntryCombinationKey', $all)
$canAuto = $formType.GetMethod('CanAutoSelectSmartMapEntry', $all)
if ($null -eq $resolve -or $null -eq $combinationKey -or $null -eq $canAuto) {
  throw '缺少条目组合或原子自动勾选测试入口。'
}

$snapshot = [Activator]::CreateInstance($snapshotType, $true).PSObject.BaseObject
Set-Field $snapshotType $snapshot 'MethodNo' '30号文'
$allScope = $scopeType.GetMethod('CreateAll', [Reflection.BindingFlags]'Public,NonPublic,Static').Invoke($null, $null)
Set-Field $snapshotType $snapshot 'SelectedScope' $allScope

$mapping = [Activator]::CreateInstance($entryType, $true).PSObject.BaseObject
Set-Field $entryType $mapping 'BoxId' 'box-combination-test'
$target = [Activator]::CreateInstance($targetType, $true).PSObject.BaseObject
Set-Field $targetType $target 'Kind' 'quota'
Set-Field $targetType $target 'Code' 'DY-1'
Set-Field $targetType $target 'Name' '候选测试定额'
Set-Field $targetType $target 'Unit' '个'
[void]$entryType.GetField('Targets', $all).GetValue($mapping).PSObject.BaseObject.Add($target)

$entryByQuota = $snapshotType.GetField('EntryByQuota', $all).GetValue($snapshot).PSObject.BaseObject
$statListType = $entryByQuota.GetType().GetGenericArguments()[1]
$stats = [Activator]::CreateInstance($statListType).PSObject.BaseObject
$projectEntriesType = $resolve.GetParameters()[1].ParameterType
$projectEntries = [Activator]::CreateInstance($projectEntriesType).PSObject.BaseObject
for ($i = 1; $i -le 20; $i++) {
  $prefix = if (($i % 2) -eq 0) { '03' } else { '04' }
  $entryCode = $prefix + '01-' + $i.ToString('00')
  $stat = [Activator]::CreateInstance($statType, $true).PSObject.BaseObject
  Set-Field $statType $stat 'EntryCode' $entryCode
  Set-Field $statType $stat 'EntryName' ('候选条目' + $i)
  Set-Field $statType $stat 'ProjectCount' (100 - $i)
  [void]$stats.Add($stat)
  $projectEntries[$entryCode] = [long]$i
}
$entryByQuota['DY-1'] = $stats

function Invoke-Combinations($Scope) {
  Set-Field $snapshotType $snapshot 'SelectedScope' $Scope
  [object[]]$arguments = New-Object object[] 6
  $arguments[0] = $snapshot
  $arguments[1] = $projectEntries
  $arguments[2] = $mapping
  $arguments[3] = '候选测试|'
  $arguments[4] = $null
  $arguments[5] = $false
  $result = $resolve.Invoke($null, $arguments)
  return [pscustomobject]@{ Combinations=$result; Truncated=[bool]$arguments[5] }
}

$allResult = Invoke-Combinations $allScope
if ($allResult.Combinations.Count -ne 16 -or -not $allResult.Truncated) {
  throw "条目组合硬上限错误：count=$($allResult.Combinations.Count), truncated=$($allResult.Truncated)"
}

$entryScope = [Activator]::CreateInstance($scopeType, $true).PSObject.BaseObject
Set-Field $scopeType $entryScope 'Kind' 'Entry'
Set-Field $scopeType $entryScope 'EntryCode' '03'
$scopedResult = Invoke-Combinations $entryScope
if ($scopedResult.Combinations.Count -ne 10 -or $scopedResult.Truncated) {
  throw '专业范围内候选数量错误。'
}
foreach ($combination in $scopedResult.Combinations) {
  if (-not ([string]$combination[0].EntryCode).StartsWith('03', [StringComparison]::Ordinal)) {
    throw '专业范围外条目进入了最终候选组合。'
  }
}

$scoreA = [Activator]::CreateInstance($scoreType, $true).PSObject.BaseObject
$scoreB = [Activator]::CreateInstance($scoreType, $true).PSObject.BaseObject
Set-Field $scoreType $scoreA 'TargetEntries' $allResult.Combinations[0]
Set-Field $scoreType $scoreB 'TargetEntries' $allResult.Combinations[1]
$keyA = [string]$combinationKey.Invoke($null, @($scoreA))
$keyB = [string]$combinationKey.Invoke($null, @($scoreB))
if ($keyA -eq $keyB -or $keyA.Length -ne 12 -or $keyB.Length -ne 12) {
  throw '候选 Key 没有包含目标条目组合身份。'
}

$scoreListType = [Collections.Generic.List``1].MakeGenericType($scoreType)
$scores = [Activator]::CreateInstance($scoreListType).PSObject.BaseObject
$blocked = [Activator]::CreateInstance($scoreType, $true).PSObject.BaseObject
Set-Field $scoreType $blocked 'CurrentTargetsValid' $true
Set-Field $scoreType $blocked 'HasEntry' $false
Set-Field $scoreType $blocked 'HasCurrentContext' $true
[void]$scores.Add($blocked)
[object[]]$autoArguments = New-Object object[] 1
$autoArguments[0] = $scores.PSObject.BaseObject
if ([bool]$canAuto.Invoke($null, $autoArguments)) { throw '组件任一成员缺条目时仍允许自动勾选。' }

Write-Host 'PASS A1/B14：无空办法回退，普通关系按软件分区，条目/公式按分区+办法号'
Write-Host 'PASS A2/A3：专业范围为最终条目硬限制，且不存在全库兜底'
Write-Host 'PASS A4/A5/T11/T12/T16：候选组合上限16、Key含组合身份、缺成员整组不自动勾选'
