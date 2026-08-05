$ErrorActionPreference = 'Stop'

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..'))
$dll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL)) { $env:RECO_EXPAND_DLL } else { Join-Path $repoRoot 'RecoQuotaRecommend\bin\RecoExpandPanel.dll' }
if (-not (Test-Path -LiteralPath $dll)) { throw "Missing DLL: $dll" }
$dllDir = Split-Path -Parent $dll
foreach ($dependency in @('NPOI.dll', 'NPOI.OpenXmlFormats.dll', 'NPOI.OpenXml4Net.dll', 'NPOI.OOXML.dll', 'ICSharpCode.SharpZipLib.dll')) {
    $dependencyPath = Join-Path $dllDir $dependency
    if (Test-Path -LiteralPath $dependencyPath) { [void][System.Reflection.Assembly]::LoadFrom($dependencyPath) }
}

$panelType = [System.Reflection.Assembly]::LoadFrom($dll).GetType('RecoNet.FormPanel', $true)
$allFlags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$nestedFlags = [System.Reflection.BindingFlags]'Public,NonPublic'
$groupType = $panelType.GetNestedType('MappingFeedbackGroup', $nestedFlags)
$targetType = $panelType.GetNestedType('MappingFeedbackTarget', $nestedFlags)
$operandType = $panelType.GetNestedType('QuantityFormulaOperandInfo', $nestedFlags)
$group = [Activator]::CreateInstance($groupType, $true).PSObject.BaseObject
$target = [Activator]::CreateInstance($targetType, $true).PSObject.BaseObject
$operand = [Activator]::CreateInstance($operandType, $true).PSObject.BaseObject

$suffix = [Guid]::NewGuid().ToString('N')
$rawName = 'CODEXSQLROLLBACK' + $suffix
$signature = $rawName + '|'
$targetCode = 'TEST-' + $suffix.Substring(0, 20)
$entryCode = '12-01'
$entryName = 'rollback entry'
$method = '2024'
$boxId = 'box-test-' + $suffix.Substring(0, 24)
$engineeringType = $entryCode.Substring(0, 2)
$contextTargetCode = 'CTX-' + $suffix.Substring(0, 18)
$conflictTargetCode = 'CTX-CONFLICT-' + $suffix.Substring(0, 9)
$contextNameA = $rawName + '-A'
$contextNameB = $rawName + '-B'
$pureContextName = $rawName + '-PURE'
$conflictName = $rawName + '-CONFLICT'
$pureSfName = $rawName + '-PURE-SF'
$mixedEntryName = $rawName + '-MIXED-ENTRY'
$invalidSfName = $rawName + '-INVALID-SF'
$pureSfBoxId = 'box-sf-' + $suffix.Substring(0, 24)
$mixedEntryBoxId = 'box-mixed-' + $suffix.Substring(0, 21)
$invalidSfBoxId = 'box-invalid-' + $suffix.Substring(0, 19)

$groupType.GetField('QuantityName', $allFlags).SetValue($group, $rawName)
$groupType.GetField('QuantityUnit', $allFlags).SetValue($group, 'm2')
$groupType.GetField('EntryCode', $allFlags).SetValue($group, $entryCode)
$groupType.GetField('EntryName', $allFlags).SetValue($group, $entryName)
$groupType.GetField('Method', $allFlags).SetValue($group, $method)
$groupType.GetField('BoxId', $allFlags).SetValue($group, $boxId)
$targetType.GetField('Kind', $allFlags).SetValue($target, 'quota')
$targetType.GetField('Code', $allFlags).SetValue($target, $targetCode)
$targetType.GetField('Name', $allFlags).SetValue($target, 'rollback target')
$targetType.GetField('Unit', $allFlags).SetValue($target, 'm3')
$targetType.GetField('FormulaTemplate', $allFlags).SetValue($target, 'V0*0.2')
$targets = $groupType.GetField('Targets', $allFlags).GetValue($group).PSObject.BaseObject
[void]$targets.Add($target)
$operandType.GetField('Name', $allFlags).SetValue($operand, $rawName)
$operandType.GetField('Unit', $allFlags).SetValue($operand, 'm2')
$operandType.GetField('Signature', $allFlags).SetValue($operand, $signature)
$operands = $groupType.GetField('FormulaOperands', $allFlags).GetValue($group).PSObject.BaseObject
[void]$operands.Add($operand)

$getConnection = $panelType.GetMethod('GetLearningDbConnectionString', $allFlags)
$upsert = $panelType.GetMethod('UpsertBindingGroupAggregates', $allFlags)
function New-ContextAggregateGroup([string]$quantityName, [string]$ordinaryCode,
    [string]$auxiliaryName, [string]$auxiliaryUnit, [string]$secondAuxiliaryName = '', [string]$secondAuxiliaryUnit = '') {
    $testGroup = [Activator]::CreateInstance($groupType, $true).PSObject.BaseObject
    foreach ($pair in @{ QuantityName=$quantityName; QuantityUnit='100m3'; EntryCode='0309-01-03-01'; EntryName='桥涵工程'; Method='2024'; BoxId=('box-shared-' + $suffix.Substring(0, 20)) }.GetEnumerator()) {
        $groupType.GetField($pair.Key, $allFlags).SetValue($testGroup, $pair.Value)
    }
    $testTargets = $groupType.GetField('Targets', $allFlags).GetValue($testGroup).PSObject.BaseObject
    if (-not [String]::IsNullOrWhiteSpace($ordinaryCode)) {
        $ordinaryTarget = [Activator]::CreateInstance($targetType, $true).PSObject.BaseObject
        foreach ($definition in @(@('Kind','quota'),@('Code',$ordinaryCode),@('Name','ordinary target'),@('Unit','100m3'))) {
            $targetType.GetField($definition[0], $allFlags).SetValue($ordinaryTarget, $definition[1])
        }
        [void]$testTargets.Add($ordinaryTarget)
    }
    $auxiliaryDefinitions = New-Object System.Collections.ArrayList
    [void]$auxiliaryDefinitions.Add([object[]]@($auxiliaryName,$auxiliaryUnit))
    if (-not [String]::IsNullOrWhiteSpace($secondAuxiliaryName)) {
        [void]$auxiliaryDefinitions.Add([object[]]@($secondAuxiliaryName,$secondAuxiliaryUnit))
    }
    foreach ($auxiliary in $auxiliaryDefinitions) {
        $auxiliaryTarget = [Activator]::CreateInstance($targetType, $true).PSObject.BaseObject
        foreach ($definition in @(@('Kind','quota'),@('Code','SH'),@('Name',[string]$auxiliary[0]),@('Unit',[string]$auxiliary[1]))) {
            $targetType.GetField($definition[0], $allFlags).SetValue($auxiliaryTarget, $definition[1])
        }
        [void]$testTargets.Add($auxiliaryTarget)
    }
    Write-Output -NoEnumerate $testGroup
}
function New-TargetEntryAggregateGroup([string]$quantityName, [string]$testBoxId, $definitions) {
    $testGroup = [Activator]::CreateInstance($groupType, $true).PSObject.BaseObject
    foreach ($pair in @{ QuantityName=$quantityName; QuantityUnit='元'; Method='2024'; BoxId=$testBoxId; AcceptedCount=1 }.GetEnumerator()) {
        $groupType.GetField($pair.Key, $allFlags).SetValue($testGroup, $pair.Value)
    }
    $testTargets = $groupType.GetField('Targets', $allFlags).GetValue($testGroup).PSObject.BaseObject
    foreach ($definition in $definitions) {
        $testTarget = [Activator]::CreateInstance($targetType, $true).PSObject.BaseObject
        foreach ($field in @(
            @('Kind',$definition[0]), @('Code',$definition[1]), @('Name',$definition[2]),
            @('Unit',$definition[3]), @('EntryCode',$definition[4]), @('EntryName',$definition[5])
        )) {
            $targetType.GetField($field[0], $allFlags).SetValue($testTarget, $field[1])
        }
        [void]$testTargets.Add($testTarget)
    }
    Write-Output -NoEnumerate $testGroup
}
$connectionString = [string]$getConnection.Invoke($null, $null)
$connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
$connection.Open()
$methodColumn = $connection.CreateCommand()
$methodColumn.CommandText = "SELECT CASE WHEN COL_LENGTH('dbo.SignatureBoxMap','method') IS NULL THEN 0 ELSE 1 END"
if ([int]$methodColumn.ExecuteScalar() -eq 0) {
    $connection.Dispose()
    Write-Host 'Test-LearningDbIncrementalAggregate: SKIP (schema.sql SignatureBoxMap.method migration not applied)'
    exit 0
}
$formulaTableState = $connection.CreateCommand()
$formulaTableState.CommandText = "SELECT CASE WHEN OBJECT_ID('dbo.QuantityFormulaRule','U') IS NOT NULL AND OBJECT_ID('dbo.QuantityFormulaOperand','U') IS NOT NULL THEN 1 ELSE 0 END"
$formulaTablesExisted = [int]$formulaTableState.ExecuteScalar() -eq 1
$engineeringTemplateState = $connection.CreateCommand()
$engineeringTemplateState.CommandText = "SELECT CASE WHEN OBJECT_ID('dbo.EngineeringTemplate','U') IS NULL THEN 0 ELSE 1 END"
if ([int]$engineeringTemplateState.ExecuteScalar() -eq 0) {
    $connection.Dispose()
    throw 'Missing dbo.EngineeringTemplate; apply the existing learning schema before running the incremental aggregate test'
}
$transaction = $connection.BeginTransaction()
try {
    $ensureFormulaTables = $connection.CreateCommand()
    $ensureFormulaTables.Transaction = $transaction
    $ensureFormulaTables.CommandText = @'
IF OBJECT_ID('dbo.QuantityFormulaRule','U') IS NULL
CREATE TABLE dbo.QuantityFormulaRule (
  rule_hash CHAR(32) PRIMARY KEY, anchor_signature NVARCHAR(450) NOT NULL,
  target_kind NVARCHAR(20) NOT NULL, target_code NVARCHAR(100) NOT NULL,
  target_unit NVARCHAR(50) NOT NULL, formula_template NVARCHAR(2000) NOT NULL,
  method NVARCHAR(50) NOT NULL DEFAULT(''), entry_code NVARCHAR(100) NOT NULL DEFAULT(''),
  sample_count INT NOT NULL DEFAULT(0), first_seen DATETIME2(0) NULL, last_seen DATETIME2(0) NULL
);
IF OBJECT_ID('dbo.QuantityFormulaOperand','U') IS NULL
CREATE TABLE dbo.QuantityFormulaOperand (
  rule_hash CHAR(32) NOT NULL, operand_index INT NOT NULL,
  operand_signature NVARCHAR(450) NOT NULL, operand_name NVARCHAR(1000) NOT NULL DEFAULT(''),
  operand_unit NVARCHAR(50) NOT NULL DEFAULT(''),
  CONSTRAINT PK_QuantityFormulaOperand_Test PRIMARY KEY(rule_hash, operand_index)
);
'@
    [void]$ensureFormulaTables.ExecuteNonQuery()

    $invokeArgs = New-Object 'object[]' 3
    $invokeArgs[0] = $connection.PSObject.BaseObject
    $invokeArgs[1] = $transaction.PSObject.BaseObject
    $invokeArgs[2] = $group
    [void]$upsert.Invoke($null, $invokeArgs)
    [void]$upsert.Invoke($null, $invokeArgs)

    $command = $connection.CreateCommand()
    $command.Transaction = $transaction
    $command.CommandText = @'
SELECT
  (SELECT COUNT(*) FROM dbo.QuantityAlias WHERE raw_name=@name AND signature=@signature) +
  (SELECT COUNT(*) FROM dbo.SignatureBoxMap WHERE signature=@signature AND method=@method AND box_id=@box) +
  (SELECT COUNT(*) FROM dbo.QuotaBoxTarget WHERE box_id=@box AND target_code=@code) +
  (SELECT COUNT(*) FROM dbo.SignatureEntryMap WHERE signature=@signature AND target_code=@code AND method=@method AND entry_code=@entry) +
  (SELECT COUNT(*) FROM dbo.EngineeringTemplate WHERE method=@method AND engineering_type=@type AND entry_code=@entry AND box_id=@box) +
  (SELECT COUNT(*) FROM dbo.QuantityFormulaRule WHERE anchor_signature=@signature AND target_code=@code) +
  (SELECT COUNT(*) FROM dbo.QuantityFormulaOperand WHERE operand_signature=@signature)
'@
    [void]$command.Parameters.AddWithValue('@name', $rawName)
    [void]$command.Parameters.AddWithValue('@signature', $signature)
    [void]$command.Parameters.AddWithValue('@box', $boxId)
    [void]$command.Parameters.AddWithValue('@code', $targetCode)
    [void]$command.Parameters.AddWithValue('@method', $method)
    [void]$command.Parameters.AddWithValue('@entry', $entryCode)
    [void]$command.Parameters.AddWithValue('@type', $engineeringType)
    $insideCount = [int]$command.ExecuteScalar()
    if ($insideCount -ne 7) { throw "Expected seven aggregate rows inside transaction, got $insideCount" }

    $counts = $connection.CreateCommand()
    $counts.Transaction = $transaction
    $counts.CommandText = @'
SELECT
  (SELECT seen_count FROM dbo.QuantityAlias WHERE raw_name=@name AND signature=@signature),
  (SELECT accepted_count FROM dbo.SignatureBoxMap WHERE signature=@signature AND method=@method AND box_id=@box),
  (SELECT sample_count FROM dbo.SignatureEntryMap WHERE signature=@signature AND target_code=@code AND method=@method AND entry_code=@entry),
  (SELECT sample_count FROM dbo.EngineeringTemplate WHERE method=@method AND engineering_type=@type AND entry_code=@entry AND box_id=@box),
  (SELECT sample_count FROM dbo.QuantityFormulaRule WHERE anchor_signature=@signature AND target_code=@code)
'@
    [void]$counts.Parameters.AddWithValue('@name', $rawName)
    [void]$counts.Parameters.AddWithValue('@signature', $signature)
    [void]$counts.Parameters.AddWithValue('@box', $boxId)
    [void]$counts.Parameters.AddWithValue('@code', $targetCode)
    [void]$counts.Parameters.AddWithValue('@method', $method)
    [void]$counts.Parameters.AddWithValue('@entry', $entryCode)
    [void]$counts.Parameters.AddWithValue('@type', $engineeringType)
    $reader = $counts.ExecuteReader()
    try {
        if (-not $reader.Read() -or $reader.GetInt32(0) -ne 2 -or $reader.GetInt32(1) -ne 2 -or
            $reader.GetInt32(2) -ne 2 -or $reader.GetInt32(3) -ne 2 -or $reader.GetInt32(4) -ne 2) {
            throw 'Repeated aggregate upsert did not update existing rows exactly once'
        }
    }
    finally {
        $reader.Dispose()
    }

    $metadata = $connection.CreateCommand()
    $metadata.Transaction = $transaction
    $metadata.CommandText = 'SELECT target_name + ''|'' + target_unit FROM dbo.QuotaBoxTarget WHERE box_id=@box AND target_code=@code'
    [void]$metadata.Parameters.AddWithValue('@box', $boxId)
    [void]$metadata.Parameters.AddWithValue('@code', $targetCode)
    if ([string]$metadata.ExecuteScalar() -ne 'rollback target|m3') { throw 'QuotaBoxTarget metadata was not persisted' }

    $alias = $connection.CreateCommand()
    $alias.Transaction = $transaction
    $alias.CommandText = 'SELECT COUNT(*) FROM dbo.QuantityAlias WHERE raw_name=@name AND signature=@signature AND quantity_unit=''m2'''
    [void]$alias.Parameters.AddWithValue('@name', $rawName)
    [void]$alias.Parameters.AddWithValue('@signature', $signature)
    $aliasCount = [int]$alias.ExecuteScalar()
    if ($aliasCount -ne 1) {
        $aliasDebug = $connection.CreateCommand()
        $aliasDebug.Transaction = $transaction
        $aliasDebug.CommandText = 'SELECT raw_name + ''|'' + quantity_unit + ''|'' + signature + ''|'' + CAST(seen_count AS nvarchar(20)) FROM dbo.QuantityAlias WHERE signature=@signature'
        [void]$aliasDebug.Parameters.AddWithValue('@signature', $signature)
        $actualAlias = [string]$aliasDebug.ExecuteScalar()
        throw "name-level QuantityAlias row mismatch: count=$aliasCount actual=$actualAlias"
    }

    $entry = $connection.CreateCommand()
    $entry.Transaction = $transaction
    $entry.CommandText = 'SELECT entry_name FROM dbo.SignatureEntryMap WHERE signature=@signature AND target_code=@code AND method=@method AND entry_code=@entry'
    [void]$entry.Parameters.AddWithValue('@signature', $signature)
    [void]$entry.Parameters.AddWithValue('@code', $targetCode)
    [void]$entry.Parameters.AddWithValue('@method', $method)
    [void]$entry.Parameters.AddWithValue('@entry', $entryCode)
    if ([string]$entry.ExecuteScalar() -ne $entryName) { throw 'SignatureEntryMap did not persist real method/entry name' }

    $groupType.GetField('AcceptedCount', $allFlags).SetValue($group, 0)
    $groupType.GetField('CorrectedCount', $allFlags).SetValue($group, 0)
    $groupType.GetField('RejectedCount', $allFlags).SetValue($group, 1)
    [void]$upsert.Invoke($null, $invokeArgs)
    $delta = $connection.CreateCommand()
    $delta.Transaction = $transaction
    $delta.CommandText = @'
SELECT accepted_count,corrected_count,rejected_count,weight,
  (SELECT sample_count FROM dbo.SignatureEntryMap WHERE signature=@signature AND target_code=@code AND method=@method AND entry_code=@entry),
  (SELECT sample_count FROM dbo.EngineeringTemplate WHERE method=@method AND engineering_type=@type AND entry_code=@entry AND box_id=@box),
  (SELECT sample_count FROM dbo.QuantityFormulaRule WHERE anchor_signature=@signature AND target_code=@code)
FROM dbo.SignatureBoxMap WHERE signature=@signature AND method=@method AND box_id=@box
'@
    [void]$delta.Parameters.AddWithValue('@signature', $signature)
    [void]$delta.Parameters.AddWithValue('@code', $targetCode)
    [void]$delta.Parameters.AddWithValue('@method', $method)
    [void]$delta.Parameters.AddWithValue('@entry', $entryCode)
    [void]$delta.Parameters.AddWithValue('@type', $engineeringType)
    [void]$delta.Parameters.AddWithValue('@box', $boxId)
    $deltaReader = $delta.ExecuteReader()
    try {
        if (-not $deltaReader.Read() -or $deltaReader.GetInt32(0) -ne 2 -or $deltaReader.GetInt32(1) -ne 0 -or
            $deltaReader.GetInt32(2) -ne 1 -or $deltaReader.GetInt32(3) -ne 10 -or
            $deltaReader.GetInt32(4) -ne 2 -or $deltaReader.GetInt32(5) -ne 2 -or $deltaReader.GetInt32(6) -ne 2) {
            throw 'Rejection delta did not lower weight without strengthening entry/template/formula evidence'
        }
    }
    finally { $deltaReader.Dispose() }

    $groupType.GetField('CorrectedCount', $allFlags).SetValue($group, 1)
    $groupType.GetField('RejectedCount', $allFlags).SetValue($group, 0)
    [void]$upsert.Invoke($null, $invokeArgs)
    $correction = $connection.CreateCommand()
    $correction.Transaction = $transaction
    $correction.CommandText = @'
SELECT corrected_count,rejected_count,weight,
  (SELECT sample_count FROM dbo.EngineeringTemplate WHERE method=@method AND engineering_type=@type AND entry_code=@entry AND box_id=@box)
FROM dbo.SignatureBoxMap WHERE signature=@signature AND method=@method AND box_id=@box
'@
    [void]$correction.Parameters.AddWithValue('@signature', $signature)
    [void]$correction.Parameters.AddWithValue('@method', $method)
    [void]$correction.Parameters.AddWithValue('@type', $engineeringType)
    [void]$correction.Parameters.AddWithValue('@entry', $entryCode)
    [void]$correction.Parameters.AddWithValue('@box', $boxId)
    $correctionReader = $correction.ExecuteReader()
    try {
        if (-not $correctionReader.Read() -or $correctionReader.GetInt32(0) -ne 1 -or
            $correctionReader.GetInt32(1) -ne 1 -or $correctionReader.GetInt32(2) -ne 30 -or
            $correctionReader.GetInt32(3) -ne 3) {
            throw 'Correction delta did not raise the corrected relation weight and template evidence'
        }
    }
    finally { $correctionReader.Dispose() }

    $formula = $connection.CreateCommand()
    $formula.Transaction = $transaction
    $formula.CommandText = @'
SELECT r.formula_template + '|' + r.target_unit + '|' + o.operand_name + '|' + o.operand_unit
FROM dbo.QuantityFormulaRule r
JOIN dbo.QuantityFormulaOperand o ON o.rule_hash=r.rule_hash AND o.operand_index=0
WHERE r.anchor_signature=@signature AND r.target_code=@code
'@
    [void]$formula.Parameters.AddWithValue('@signature', $signature)
    [void]$formula.Parameters.AddWithValue('@code', $targetCode)
    if ([string]$formula.ExecuteScalar() -ne ('V0*0.2|m3|' + $rawName + '|m2')) { throw 'Formula rule/operand metadata was not persisted' }

    $contextGroupA = New-ContextAggregateGroup $contextNameA $contextTargetCode '消纳费' 'm3'
    $contextGroupB = New-ContextAggregateGroup $contextNameB $contextTargetCode 'PE' 'm'
    $pureContextGroup = New-ContextAggregateGroup $pureContextName '' 'FAS 联动接入' '项'
    foreach ($contextGroup in @($contextGroupA,$contextGroupB,$pureContextGroup)) {
        $contextArgs = New-Object 'object[]' 3
        $contextArgs[0] = $connection.PSObject.BaseObject
        $contextArgs[1] = $transaction.PSObject.BaseObject
        $contextArgs[2] = $contextGroup.PSObject.BaseObject
        [void]$upsert.Invoke($null, $contextArgs)
    }
    $contextQuery = $connection.CreateCommand()
    $contextQuery.Transaction = $transaction
    $contextQuery.CommandText = @'
SELECT sh.target_name + '|' + sh.target_unit + '|' + sh.box_id
FROM dbo.QuotaBoxTarget ordinary
JOIN dbo.QuotaBoxTarget sh ON sh.box_id=ordinary.box_id AND sh.target_code='SH'
WHERE ordinary.target_code=@code
ORDER BY sh.target_name,sh.target_unit
'@
    [void]$contextQuery.Parameters.AddWithValue('@code', $contextTargetCode)
    $contextRows = @()
    $contextReader = $contextQuery.ExecuteReader()
    try { while ($contextReader.Read()) { $contextRows += $contextReader.GetString(0) } } finally { $contextReader.Dispose() }
    if ($contextRows.Count -ne 2 -or $contextRows[0] -notmatch '^PE\|m\|' -or $contextRows[1] -notmatch '^消纳费\|m3\|' -or
        ($contextRows[0].Split('|')[-1] -eq $contextRows[1].Split('|')[-1])) {
        throw "Context-sensitive SH identities were not split into stable boxes: $($contextRows -join '; ')"
    }
    $pureContextQuery = $connection.CreateCommand()
    $pureContextQuery.Transaction = $transaction
    $pureContextQuery.CommandText = @'
SELECT COUNT(*)
FROM dbo.QuantityAlias a
JOIN dbo.SignatureBoxMap m ON m.signature=a.signature AND m.method=@method
JOIN dbo.QuotaBoxTarget t ON t.box_id=m.box_id
WHERE a.raw_name=@name AND t.target_code='SH' AND t.target_name=@target_name AND t.target_unit=@target_unit
'@
    [void]$pureContextQuery.Parameters.AddWithValue('@method', $method)
    [void]$pureContextQuery.Parameters.AddWithValue('@name', $pureContextName)
    [void]$pureContextQuery.Parameters.AddWithValue('@target_name', 'FAS 联动接入')
    [void]$pureContextQuery.Parameters.AddWithValue('@target_unit', '项')
    if ([int]$pureContextQuery.ExecuteScalar() -ne 1) {
        throw '名称和单位完整的纯 SH 组件未进入增量聚合'
    }

    $pureSfDefinitions = New-Object System.Collections.ArrayList
    [void]$pureSfDefinitions.Add([object[]]@('quota','SF','设备购置费','元','0802-01','设备购置费'))
    $pureSfGroup = New-TargetEntryAggregateGroup $pureSfName $pureSfBoxId $pureSfDefinitions
    $mixedDefinitions = New-Object System.Collections.ArrayList
    [void]$mixedDefinitions.Add([object[]]@('quota','EY-299','安装设备','台','0801-01','安装工程费'))
    [void]$mixedDefinitions.Add([object[]]@('quota','SF','设备购置费','元','0802-01','设备购置费'))
    [void]$mixedDefinitions.Add([object[]]@('quota','ZLF','装料费','m3','0801-01','安装工程费'))
    [void]$mixedDefinitions.Add([object[]]@('quota','SH','配合费','项','0803-01','其他工程费'))
    $mixedEntryGroup = New-TargetEntryAggregateGroup $mixedEntryName $mixedEntryBoxId $mixedDefinitions
    $invalidDefinitions = New-Object System.Collections.ArrayList
    [void]$invalidDefinitions.Add([object[]]@('quota','EY-299','安装设备','台','0802-01','设备购置费'))
    [void]$invalidDefinitions.Add([object[]]@('quota','SF','设备购置费','元','0801-01','安装工程费'))
    $invalidSfGroup = New-TargetEntryAggregateGroup $invalidSfName $invalidSfBoxId $invalidDefinitions
    foreach ($targetEntryGroup in @($pureSfGroup,$mixedEntryGroup,$invalidSfGroup)) {
        $targetEntryArgs = New-Object 'object[]' 3
        $targetEntryArgs[0] = $connection.PSObject.BaseObject
        $targetEntryArgs[1] = $transaction.PSObject.BaseObject
        $targetEntryArgs[2] = $targetEntryGroup.PSObject.BaseObject
        [void]$upsert.Invoke($null, $targetEntryArgs)
    }
    $targetEntryQuery = $connection.CreateCommand()
    $targetEntryQuery.Transaction = $transaction
    $targetEntryQuery.CommandText = @'
SELECT
  (SELECT COUNT(*) FROM dbo.EngineeringTemplate WHERE box_id=@pure_box AND entry_code='0802-01'),
  (SELECT COUNT(*) FROM dbo.SignatureEntryMap WHERE signature=@pure_sig AND target_code='SF' AND entry_code='0802-01'),
  (SELECT COUNT(*) FROM dbo.EngineeringTemplate WHERE box_id=@mixed_box),
  (SELECT COUNT(*) FROM dbo.QuantityAlias WHERE raw_name=@invalid_name)
'@
    [void]$targetEntryQuery.Parameters.AddWithValue('@pure_box', $pureSfBoxId)
    [void]$targetEntryQuery.Parameters.AddWithValue('@pure_sig', $pureSfName + '|')
    [void]$targetEntryQuery.Parameters.AddWithValue('@mixed_box', $mixedEntryBoxId)
    [void]$targetEntryQuery.Parameters.AddWithValue('@invalid_name', $invalidSfName)
    $targetEntryReader = $targetEntryQuery.ExecuteReader()
    try {
        if (-not $targetEntryReader.Read() -or $targetEntryReader.GetInt32(0) -ne 1 -or
            $targetEntryReader.GetInt32(1) -ne 1 -or $targetEntryReader.GetInt32(2) -ne 2 -or
            $targetEntryReader.GetInt32(3) -ne 0) {
            throw '纯 SF 未归入设备购置费范围、混合框未按普通/SF 两条目归集，或 SF 违规组进入了聚合'
        }
    }
    finally { $targetEntryReader.Dispose() }

    $conflictGroup = New-ContextAggregateGroup $conflictName $conflictTargetCode '消纳费' 'm3' 'PE' 'm'
    $conflictArgs = New-Object 'object[]' 3
    $conflictArgs[0] = $connection.PSObject.BaseObject
    $conflictArgs[1] = $transaction.PSObject.BaseObject
    $conflictArgs[2] = $conflictGroup.PSObject.BaseObject
    [void]$upsert.Invoke($null, $conflictArgs)
    $conflictQuery = $connection.CreateCommand()
    $conflictQuery.Transaction = $transaction
    $conflictQuery.CommandText = 'SELECT COUNT(*) FROM dbo.QuantityAlias WHERE raw_name=@name'
    [void]$conflictQuery.Parameters.AddWithValue('@name', $conflictName)
    if ([int]$conflictQuery.ExecuteScalar() -ne 0) { throw 'Same-event SH conflicts must remain only in BindingLog and not enter aggregates' }
}
finally {
    $transaction.Rollback()
    $transaction.Dispose()
}

$verify = $connection.CreateCommand()
$verify.CommandText = @'
SELECT
  (SELECT COUNT(*) FROM dbo.QuantityAlias WHERE raw_name=@name) +
  (SELECT COUNT(*) FROM dbo.SignatureBoxMap WHERE signature=@signature) +
  (SELECT COUNT(*) FROM dbo.QuotaBoxTarget WHERE target_code=@code) +
  (SELECT COUNT(*) FROM dbo.SignatureEntryMap WHERE signature=@signature AND target_code=@code) +
  (SELECT COUNT(*) FROM dbo.EngineeringTemplate WHERE entry_code=@entry AND box_id=@box)
'@
if ($formulaTablesExisted) {
    $verify.CommandText += ' + (SELECT COUNT(*) FROM dbo.QuantityFormulaRule WHERE anchor_signature=@signature AND target_code=@code) + (SELECT COUNT(*) FROM dbo.QuantityFormulaOperand WHERE operand_signature=@signature)'
}
[void]$verify.Parameters.AddWithValue('@name', $rawName)
[void]$verify.Parameters.AddWithValue('@signature', $signature)
[void]$verify.Parameters.AddWithValue('@code', $targetCode)
[void]$verify.Parameters.AddWithValue('@entry', $entryCode)
[void]$verify.Parameters.AddWithValue('@box', $boxId)
$afterCount = [int]$verify.ExecuteScalar()
$contextVerify = $connection.CreateCommand()
$contextVerify.CommandText = 'SELECT (SELECT COUNT(*) FROM dbo.QuotaBoxTarget WHERE target_code IN (@context,@conflict)) + (SELECT COUNT(*) FROM dbo.QuantityAlias WHERE raw_name IN (@nameA,@nameB,@pureName,@conflictName))'
[void]$contextVerify.Parameters.AddWithValue('@context', $contextTargetCode)
[void]$contextVerify.Parameters.AddWithValue('@conflict', $conflictTargetCode)
[void]$contextVerify.Parameters.AddWithValue('@nameA', $contextNameA)
[void]$contextVerify.Parameters.AddWithValue('@nameB', $contextNameB)
[void]$contextVerify.Parameters.AddWithValue('@pureName', $pureContextName)
[void]$contextVerify.Parameters.AddWithValue('@conflictName', $conflictName)
$contextAfterCount = [int]$contextVerify.ExecuteScalar()
$connection.Dispose()
if ($afterCount -ne 0 -or $contextAfterCount -ne 0) { throw "Rollback left $afterCount ordinary and $contextAfterCount context test rows" }

Write-Host 'Test-LearningDbIncrementalAggregate: PASS (ordinary idempotence, exact pure SH aggregation, SH identity split/conflict guard, no residue after rollback)'
