$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# 审查 §2.3 / §2.4 反射回归：
#   1. IsLearningDbWritableGroup / IsLearningDbStructurallyValidGroup 对 可推荐 / 不可推荐 / 结构不合规 三类组的判定；
#   2. HasUnsupportedLearningDbMethod 只在整批证据全部不可推荐时才判永久失败；
#   3. BuildSmartRowQuantitySignature(TargetQtyRow) 与 BuildSmartQuantitySignature(name, unit) 同口径，且无尾缀名称恒等。
# 纯内存反射，不连接 SQL，不落盘。

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..'))
$dll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL)) { $env:RECO_EXPAND_DLL } else { Join-Path $repoRoot 'RecoQuotaRecommend\bin\RecoExpandPanel.dll' }
if (-not (Test-Path -LiteralPath $dll)) { throw "Missing DLL: $dll" }
$dllDir = Split-Path -Parent $dll
foreach ($dependency in @('NPOI.dll', 'NPOI.OpenXmlFormats.dll', 'NPOI.OpenXml4Net.dll', 'NPOI.OOXML.dll', 'ICSharpCode.SharpZipLib.dll')) {
    $dependencyPath = Join-Path $dllDir $dependency
    if (Test-Path -LiteralPath $dependencyPath) { [void][System.Reflection.Assembly]::LoadFrom($dependencyPath) }
}

$assembly = [System.Reflection.Assembly]::LoadFrom($dll)
$panelType = $assembly.GetType('RecoNet.FormPanel', $true)
$allFlags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$nestedFlags = [System.Reflection.BindingFlags]'Public,NonPublic'

function Require-Method([Type]$Owner, [string]$Name) {
    $method = $Owner.GetMethod($Name, $allFlags)
    if ($null -eq $method) { throw "Missing regression seam: $($Owner.FullName).$Name" }
    return $method
}
function Require-NestedType([Type]$Owner, [string]$Name) {
    $nested = $Owner.GetNestedType($Name, $nestedFlags)
    if ($null -eq $nested) { throw "Missing nested type: $($Owner.FullName)+$Name" }
    return $nested
}
function Unwrap([object]$Value) {
    if ($null -eq $Value) { return $null }
    return $Value.PSObject.BaseObject
}
function Set-Field([object]$Instance, [Type]$Owner, [string]$Name, [object]$Value) {
    $field = $Owner.GetField($Name, $allFlags)
    if ($null -eq $field) { throw "Missing field: $($Owner.FullName).$Name" }
    $field.SetValue((Unwrap $Instance), $Value)
}

$groupType = Require-NestedType $panelType 'MappingFeedbackGroup'
$targetType = Require-NestedType $panelType 'MappingFeedbackTarget'
$batchType = Require-NestedType $panelType 'LearningDbOutboxBatch'
$rowType = Require-NestedType $panelType 'TargetQtyRow'

$isWritable = Require-Method $panelType 'IsLearningDbWritableGroup'
$isStructural = Require-Method $panelType 'IsLearningDbStructurallyValidGroup'
$isRecommendable = Require-Method $panelType 'IsLearningFeedbackGroupRecommendable'
$hasUnsupported = Require-Method $panelType 'HasUnsupportedLearningDbMethod'
$createBatch = Require-Method $panelType 'CreateLearningDbOutboxBatch'
$rowSignature = Require-Method $panelType 'BuildSmartRowQuantitySignature'
$quantitySignature = Require-Method $panelType 'BuildSmartQuantitySignature'
$normalizeSignature = Require-Method $panelType 'NormalizeForSignature'

# 核对签名，避免旧 DLL 或参数变更导致误判。
$hasUnsupportedParams = $hasUnsupported.GetParameters()
if ($hasUnsupportedParams.Count -ne 2 -or $hasUnsupportedParams[0].ParameterType -ne $batchType -or -not $hasUnsupportedParams[1].IsOut) {
    throw "Unexpected HasUnsupportedLearningDbMethod signature: $($hasUnsupported)"
}
$listType = [System.Collections.Generic.List`1].MakeGenericType($groupType)
$createBatchParams = $createBatch.GetParameters()
if ($createBatchParams.Count -ne 2 -or $createBatchParams[0].ParameterType -ne [string] -or $createBatchParams[1].ParameterType -ne $listType) {
    throw "Unexpected CreateLearningDbOutboxBatch signature: $($createBatch)"
}
$rowSignatureParams = $rowSignature.GetParameters()
if ($rowSignatureParams.Count -ne 1 -or $rowSignatureParams[0].ParameterType -ne $rowType) {
    throw "Unexpected BuildSmartRowQuantitySignature signature: $($rowSignature)"
}
Write-Host 'PASS reflection seams: IsLearningDbWritableGroup / IsLearningDbStructurallyValidGroup / HasUnsupportedLearningDbMethod(LearningDbOutboxBatch, out string) / CreateLearningDbOutboxBatch(string, List<MappingFeedbackGroup>) / BuildSmartRowQuantitySignature(TargetQtyRow)'

$suffix = [Guid]::NewGuid().ToString('N')
$method = '2024'
$methodNo = 'TB 10801—2024'
$softwarePartition = '2024'

function New-Target([string]$Kind, [string]$Code, [string]$Name, [string]$Unit, [string]$EntryName = '') {
    $target = [Activator]::CreateInstance($targetType, $true).PSObject.BaseObject
    Set-Field $target $targetType 'Kind' $Kind
    Set-Field $target $targetType 'Code' $Code
    Set-Field $target $targetType 'Name' $Name
    Set-Field $target $targetType 'Unit' $Unit
    if (-not [String]::IsNullOrWhiteSpace($EntryName)) { Set-Field $target $targetType 'EntryName' $EntryName }
    Write-Output -NoEnumerate $target
}
function New-Group([string]$QuantityName, [string]$EntryName, [object[]]$Targets) {
    $group = [Activator]::CreateInstance($groupType, $true).PSObject.BaseObject
    Set-Field $group $groupType 'QuantityName' $QuantityName
    Set-Field $group $groupType 'QuantityUnit' 'm3'
    Set-Field $group $groupType 'EntryCode' '0309-01-03-01'
    Set-Field $group $groupType 'EntryName' $EntryName
    Set-Field $group $groupType 'Method' $method
    Set-Field $group $groupType 'MethodNo' $methodNo
    Set-Field $group $groupType 'SoftwarePartition' $softwarePartition
    Set-Field $group $groupType 'BoxId' ('box-filter-' + $suffix.Substring(0, 20))
    $list = $groupType.GetField('Targets', $allFlags).GetValue($group).PSObject.BaseObject
    foreach ($target in $Targets) { [void]$list.Add((Unwrap $target)) }
    Write-Output -NoEnumerate $group
}
function Invoke-Bool([System.Reflection.MethodInfo]$Method, [object]$Argument) {
    $args = [object[]]::new(1)
    $args[0] = Unwrap $Argument
    return [bool]$Method.Invoke($null, $args)
}

# A：结构合规且可推荐（普通定额目标，非上下文敏感码）。
$groupA = New-Group ('FILTER-A-' + $suffix) '桥涵工程' @((New-Target 'quota' ('TEST-A-' + $suffix.Substring(0, 16)) 'ordinary target' 'm3'))
# B：结构合规但不可推荐——SF 目标绑在非“设备购置费”条目（IsLearningGroupRecommendable: sf != equipmentEntry -> false）。
$groupB = New-Group ('FILTER-B-' + $suffix) '桥涵工程' @((New-Target 'quota' 'SF' 'sf target' '元'))
# C：结构不合规——没有任何目标。
$groupC = New-Group ('FILTER-C-' + $suffix) '桥涵工程' @()

# ---------- 断言 1：组级过滤 ----------
if (-not (Invoke-Bool $isStructural $groupA)) { throw 'A should be structurally valid' }
if (-not (Invoke-Bool $isRecommendable $groupA)) { throw 'A should be recommendable' }
if (-not (Invoke-Bool $isWritable $groupA)) { throw 'IsLearningDbWritableGroup(A) should be true' }

if (-not (Invoke-Bool $isStructural $groupB)) { throw 'IsLearningDbStructurallyValidGroup(B) should be true' }
if (Invoke-Bool $isRecommendable $groupB) { throw 'B (SF on non-equipment entry) should NOT be recommendable' }
if (Invoke-Bool $isWritable $groupB) { throw 'IsLearningDbWritableGroup(B) should be false' }

if (Invoke-Bool $isStructural $groupC) { throw 'C (no targets) should NOT be structurally valid' }
if (Invoke-Bool $isWritable $groupC) { throw 'IsLearningDbWritableGroup(C) should be false' }

# 额外：辅助码缺单位也应判不可推荐但结构合规（覆盖 §2.3 另一条不可推荐路径）。
$groupB2 = New-Group ('FILTER-B2-' + $suffix) '桥涵工程' @((New-Target 'quota' ('TEST-B2-' + $suffix.Substring(0, 15)) 'ordinary' 'm3'), (New-Target 'quota' 'SH' 'aux without unit' ''))
if (-not (Invoke-Bool $isStructural $groupB2)) { throw 'B2 (SH missing unit) should be structurally valid' }
if (Invoke-Bool $isWritable $groupB2) { throw 'IsLearningDbWritableGroup(B2, SH missing unit) should be false' }

if ($null -ne $isWritable.Invoke($null, @([object]$null)) -and [bool]$isWritable.Invoke($null, @([object]$null))) { throw 'IsLearningDbWritableGroup(null) should be false' }
Write-Host 'PASS IsLearningDbWritableGroup: A=true, B(SF non-equipment)=false, B2(SH no unit)=false, C(no targets)=false; IsLearningDbStructurallyValidGroup(B)=true, (C)=false'

# ---------- 断言 2：HasUnsupportedLearningDbMethod 仅整批全无效才永久失败 ----------
function New-Batch([object[]]$Groups) {
    $list = [Activator]::CreateInstance($listType)
    foreach ($group in $Groups) { [void]$list.Add((Unwrap $group)) }
    $args = [object[]]::new(2)
    $args[0] = 'test-group-filter'
    $args[1] = $list
    return $createBatch.Invoke($null, $args)
}
function Test-Unsupported([object]$Batch) {
    $args = [object[]]::new(2)
    $args[0] = Unwrap $Batch
    $args[1] = $null
    $result = [bool]$hasUnsupported.Invoke($null, $args)
    return @{ Unsupported = $result; Reason = [string]$args[1] }
}

$batchAB = New-Batch @($groupA, $groupB)
$batchPartition = [string]$batchType.GetField('SoftwarePartition', $allFlags).GetValue($batchAB)
if ($batchPartition -ne $softwarePartition) { throw "Batch partition should be '$softwarePartition', got '$batchPartition'" }
$groupsCount = $batchType.GetField('Groups', $allFlags).GetValue($batchAB).Count
if ($groupsCount -ne 2) { throw "Batch A+B should carry 2 groups, got $groupsCount" }

$partial = Test-Unsupported $batchAB
if ($partial.Unsupported) { throw "Partial-invalid batch (A+B) should NOT be permanent failure; reason='$($partial.Reason)'" }
if (-not [String]::IsNullOrEmpty($partial.Reason)) { throw "Partial-invalid batch should leave reason empty, got '$($partial.Reason)'" }

$onlyA = Test-Unsupported (New-Batch @($groupA))
if ($onlyA.Unsupported) { throw "All-valid batch (A) should NOT be permanent failure; reason='$($onlyA.Reason)'" }

$onlyB = Test-Unsupported (New-Batch @($groupB))
if (-not $onlyB.Unsupported) { throw 'All-invalid batch (only B) should be permanent failure' }
if ($onlyB.Reason -notmatch '^unsupported_learning_group_partition_0_mismatch_0_method_0_evidence_1$') {
    throw "All-invalid batch reason mismatch: '$($onlyB.Reason)'"
}

$bothInvalid = Test-Unsupported (New-Batch @($groupB, $groupB2))
if (-not $bothInvalid.Unsupported) { throw 'All-invalid batch (B+B2) should be permanent failure' }
if ($bothInvalid.Reason -notmatch '_evidence_2$') { throw "All-invalid batch (B+B2) reason mismatch: '$($bothInvalid.Reason)'" }
Write-Host "PASS HasUnsupportedLearningDbMethod: A+B=false (partial), A=false, B=true ($($onlyB.Reason)), B+B2=true ($($bothInvalid.Reason))"

# ---------- 断言 3：BuildSmartRowQuantitySignature 与 BuildSmartQuantitySignature 同口径 ----------
function New-Row([string]$RawName, [string]$Unit) {
    $row = [Activator]::CreateInstance($rowType, $true).PSObject.BaseObject
    Set-Field $row $rowType 'Row' 10
    Set-Field $row $rowType 'RawName' $RawName
    Set-Field $row $rowType 'Unit' $Unit
    Set-Field $row $rowType 'Chapter' ''
    Write-Output -NoEnumerate $row
}
function Get-RowSignature([object]$Row) {
    $args = [object[]]::new(1)
    $args[0] = Unwrap $Row
    return [string]$rowSignature.Invoke($null, $args)
}
function Get-QuantitySignature([string]$Name, [string]$Unit) {
    $args = [object[]]::new(2)
    $args[0] = $Name
    $args[1] = $Unit
    return [string]$quantitySignature.Invoke($null, $args)
}
function Get-Normalized([string]$Text) {
    $args = [object[]]::new(1)
    $args[0] = $Text
    return [string]$normalizeSignature.Invoke($null, $args)
}

$rowWithSuffix = Get-RowSignature (New-Row '混凝土 m3' 'm3')
$functionWithSuffix = Get-QuantitySignature '混凝土 m3' 'm3'
$rowPlain = Get-RowSignature (New-Row '混凝土' 'm3')
$expectedConcrete = (Get-Normalized '混凝土') + '|'
if ($rowWithSuffix -ne $functionWithSuffix) { throw "Row signature '$rowWithSuffix' != BuildSmartQuantitySignature '$functionWithSuffix'" }
if ($rowWithSuffix -ne $rowPlain) { throw "Row signature with unit suffix '$rowWithSuffix' != plain row signature '$rowPlain'" }
if ($rowWithSuffix -ne $expectedConcrete) { throw "Row signature '$rowWithSuffix' != NormalizeForSignature('混凝土')+'|' '$expectedConcrete'" }

$rowClamp = Get-RowSignature (New-Row '设备线夹' '')
$expectedClamp = (Get-Normalized '设备线夹') + '|'
if ($rowClamp -ne $expectedClamp) { throw "Identity signature mismatch: '$rowClamp' != '$expectedClamp' (unit suffix must not be stripped)" }
if ($rowClamp -ne (Get-QuantitySignature '设备线夹' '')) { throw 'Row signature for 设备线夹 should equal BuildSmartQuantitySignature(设备线夹, empty)' }

$rowNull = Get-RowSignature $null
if ($rowNull -ne '|') { throw "Null row signature should be '|', got '$rowNull'" }
Write-Host "PASS BuildSmartRowQuantitySignature: '混凝土 m3'/m3 == '混凝土'/m3 == '$rowWithSuffix'; '设备线夹'/'' == '$rowClamp' (identity); null == '|'"

Write-Host 'PASS Test-LearningDbGroupFilter'
