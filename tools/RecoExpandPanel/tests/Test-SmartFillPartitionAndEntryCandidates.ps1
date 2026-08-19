param(
  [Parameter(Mandatory)][string]$ExpandDll
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

function Set-Field {
  param([Type]$Type, $Instance, [string]$Name, $Value)
  $field = $Type.GetField($Name, [Reflection.BindingFlags]'Public,NonPublic,Instance,Static')
  if ($null -eq $field) { throw "Missing field: $Name" }
  $field.SetValue($Instance, $Value)
}

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..')).Path
$source = [IO.File]::ReadAllText((Join-Path $repoRoot 'tools\RecoExpandPanel\SmartFillFeature.cs'), [Text.Encoding]::UTF8)
foreach ($forbidden in @(
  "method=''", "method = ''", '名称学习命中，全库兜底', '名称兼容命中，全库兜底', '"全库兜底"',
  'SmartEntryStat', 'SmartTargetEntryResolution', 'ResolveSmartTargetEntryCombinations',
  'BuildSmartEntryCombinationKey', 'EntryByQuota', 'EntryBySignature', 'LocalContextKeys',
  'MaxEntryCombinations')) {
  if ($source.Contains($forbidden)) { throw "SmartFill 仍含禁止的全库回退或条目推理：$forbidden" }
}
foreach ($required in @(
  'm.software_partition=@software_partition',
  'software_partition=@software_partition AND method_no=@method_no',
  'r.software_partition=@software_partition AND r.method_no=@method_no',
  'FilterSmartHitsByScope',
  'ScopeEntriesByBox')) {
  if (-not $source.Contains($required)) { throw "SmartFill 缺少分区或持久化专业范围门禁：$required" }
}

$assembly = [Reflection.Assembly]::Load([IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $ExpandDll).Path))
$formType = $assembly.GetType('RecoNet.FormPanel', $true)
$nested = [Reflection.BindingFlags]'Public,NonPublic'
$all = [Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$entryType = $formType.GetNestedType('SmartMapEntry', $nested)
$targetType = $formType.GetNestedType('SmartBoxTarget', $nested)
$scoreType = $formType.GetNestedType('SmartMapCandidateScore', $nested)
$canAuto = $formType.GetMethod('CanAutoSelectSmartMapEntry', $all)
$candidateLabel = $formType.GetMethod('BuildSmartCandidateLabel', $all)
if ($null -eq $entryType -or $null -eq $targetType -or $null -eq $scoreType -or
    $null -eq $canAuto -or $null -eq $candidateLabel) {
  throw '缺少新的组件排序或候选标签行为入口。'
}
foreach ($removedField in @('EntryCode','EntryName','TargetEntries','HasEntry','HasCurrentContext')) {
  if ($null -ne $scoreType.GetField($removedField, $all)) {
    throw "候选分数仍携带条目推理字段：$removedField"
  }
}

function New-Score([string]$BoxId, [int]$Weight, [bool]$CurrentMethod) {
  $entry = [Activator]::CreateInstance($entryType, $true).PSObject.BaseObject
  Set-Field $entryType $entry 'BoxId' $BoxId
  Set-Field $entryType $entry 'Weight' $Weight
  $target = [Activator]::CreateInstance($targetType, $true).PSObject.BaseObject
  Set-Field $targetType $target 'Kind' 'quota'
  Set-Field $targetType $target 'Code' 'DY-1'
  Set-Field $targetType $target 'Name' '候选测试定额'
  Set-Field $targetType $target 'Unit' '个'
  [void]$entryType.GetField('Targets', $all).GetValue($entry).PSObject.BaseObject.Add($target)
  $score = [Activator]::CreateInstance($scoreType, $true).PSObject.BaseObject
  Set-Field $scoreType $score 'Entry' $entry
  Set-Field $scoreType $score 'CurrentTargetsValid' $true
  Set-Field $scoreType $score 'HasCurrentMethodMapping' $CurrentMethod
  return $score
}

$scoreListType = [Collections.Generic.List``1].MakeGenericType($scoreType)
$scores = [Activator]::CreateInstance($scoreListType).PSObject.BaseObject
[void]$scores.Add((New-Score 'box-high' 50 $true))
[void]$scores.Add((New-Score 'box-low' 40 $true))
[object[]]$autoArguments = New-Object object[] 1
$autoArguments[0] = $scores.PSObject.BaseObject
if ([bool]$canAuto.Invoke($null, $autoArguments)) { throw '两个完整组件小权重差时不应静默选择。' }
Set-Field $entryType ($scoreType.GetField('Entry', $all).GetValue($scores[1])) 'Weight' 20
if (-not [bool]$canAuto.Invoke($null, $autoArguments)) { throw '当前办法完整组件权重差30时应允许自动选择。' }

$label = [string]$candidateLabel.Invoke($null, @($scores[0]))
if ($label -ne 'DY-1（候选测试定额 / 个）' -or $label -match '条目|权重|当前办法') {
  throw "候选标签仍泄漏条目推理或内部排序证据：$label"
}

Write-Host 'PASS A1/B14：无空办法回退，普通关系按软件分区，专业范围按分区+办法号'
Write-Host 'PASS A2/A3：条目不再作为推理输出，候选标签不含条目'
Write-Host 'PASS A4/A5：组件仍按完整身份和权重门槛整组自动选择'
