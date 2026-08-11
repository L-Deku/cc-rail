param(
  [Parameter(Mandatory)][string]$ExpandDll
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

function Set-Field {
  param([Type]$Type, $Instance, [string]$Name, $Value)
  $field = $Type.GetField($Name, [Reflection.BindingFlags]'Instance,Public,NonPublic')
  if ($field -eq $null) { throw "Missing field: $Name" }
  $field.SetValue($Instance, $Value)
}

function Get-Md5 {
  param([string]$Text)
  $md5 = [Security.Cryptography.MD5]::Create()
  try {
    return ([BitConverter]::ToString($md5.ComputeHash([Text.Encoding]::UTF8.GetBytes($Text)))).Replace('-', '').ToLowerInvariant()
  }
  finally { $md5.Dispose() }
}

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..')).Path
. (Join-Path $repoRoot 'tools\RecoLearning\Common.ps1')
$assembly = [Reflection.Assembly]::Load([IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $ExpandDll).Path))
$formType = $assembly.GetType('RecoNet.FormPanel', $true)
$nested = [Reflection.BindingFlags]'Public,NonPublic'
$all = [Reflection.BindingFlags]'Static,Instance,Public,NonPublic'
$groupType = $formType.GetNestedType('MappingFeedbackGroup', $nested)
$targetType = $formType.GetNestedType('MappingFeedbackTarget', $nested)
$operandType = $formType.GetNestedType('QuantityFormulaOperandInfo', $nested)
$group = [Activator]::CreateInstance($groupType, $true)
$target = [Activator]::CreateInstance($targetType, $true)
$operand = [Activator]::CreateInstance($operandType, $true)

Set-Field $groupType $group 'QuantityName' 'FormulaProbe'
Set-Field $groupType $group 'SoftwarePartition' '2020'
Set-Field $groupType $group 'MethodNo' '30号文'
Set-Field $groupType $group 'Method' '2020'
Set-Field $targetType $target 'Kind' 'quota'
Set-Field $targetType $target 'Code' 'DY-1'
Set-Field $targetType $target 'Unit' ' 10ｍ³ '
Set-Field $targetType $target 'FormulaTemplate' 'V0/100'
Set-Field $targetType $target 'EntryCode' ('0101' + [char]0x2014 + '01')
Set-Field $operandType $operand 'Signature' 'OPERAND|'
Set-Field $operandType $operand 'Name' 'operand'
Set-Field $operandType $operand 'Unit' '㎥'
[void]$groupType.GetField('FormulaOperands', $all).GetValue($group).PSObject.BaseObject.Add($operand)

$hashMethod = $formType.GetMethod('BuildLearningFormulaRuleHash', $all)
[object[]]$arguments = @($group.PSObject.BaseObject, $target.PSObject.BaseObject)
$actual = [string]$hashMethod.Invoke($null, $arguments)
$targetUnit = Get-NormalizedLearningFormulaUnit ' 10ｍ³ '
$operandUnit = Get-NormalizedLearningFormulaUnit '㎥'
$raw = 'FORMULAPROBE||quota:DY-1|' + $targetUnit + '|V0/100|2020|30号文|0101-01|OPERAND|@' + $operandUnit
$expected = Get-Md5 $raw
if ($actual -ne $expected) { throw "C#/PowerShell 公式哈希不一致：$actual / $expected" }

Set-Field $groupType $group 'MethodNo' '101号文估算'
$methodHash = [string]$hashMethod.Invoke($null, $arguments)
if ($methodHash -eq $actual) { throw '只改变 method_no 时公式哈希没有变化。' }
Set-Field $groupType $group 'MethodNo' '30号文'
Set-Field $groupType $group 'SoftwarePartition' '2024'
$partitionHash = [string]$hashMethod.Invoke($null, $arguments)
if ($partitionHash -eq $actual) { throw '只改变 software_partition 时公式哈希没有变化。' }

$rebuild = [IO.File]::ReadAllText((Join-Path $repoRoot 'tools\RecoLearning\Rebuild-Aggregates.ps1'), [Text.Encoding]::UTF8)
if ($rebuild -match "Get-ExtraText `$extra 'formula_rule_hash'" -or
    $rebuild -notmatch '\$ruleHash = Get-Md5Hex \$ruleRaw') {
  throw '全量重建仍可能信任旧 formula_rule_hash。'
}
$requiredFormulaGate = '$formulaPartition -eq '''' -or $formulaMethodNo -eq '''' -or $formulaEntryCode -eq '''''
if (-not $rebuild.Contains($requiredFormulaGate)) {
  throw '全量重建没有同时要求分区、办法号和条目编号。'
}

Write-Host 'PASS B6/T4/T18/T19：跨语言单位归一化一致；分区和办法号进入新哈希；旧哈希无条件失效'
Write-Host 'PASS B7/T14：全量公式要求精确分区、办法号和有效条目编号'
