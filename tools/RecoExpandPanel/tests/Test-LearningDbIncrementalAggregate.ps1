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
$entryCode = 'ENTRY-' + $suffix.Substring(0, 20)
$entryName = 'rollback entry'
$method = '2024'
$boxId = 'box-test-' + $suffix.Substring(0, 24)

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
  (SELECT COUNT(*) FROM dbo.QuantityFormulaRule WHERE anchor_signature=@signature AND target_code=@code) +
  (SELECT COUNT(*) FROM dbo.QuantityFormulaOperand WHERE operand_signature=@signature)
'@
    [void]$command.Parameters.AddWithValue('@name', $rawName)
    [void]$command.Parameters.AddWithValue('@signature', $signature)
    [void]$command.Parameters.AddWithValue('@box', $boxId)
    [void]$command.Parameters.AddWithValue('@code', $targetCode)
    [void]$command.Parameters.AddWithValue('@method', $method)
    [void]$command.Parameters.AddWithValue('@entry', $entryCode)
    $insideCount = [int]$command.ExecuteScalar()
    if ($insideCount -ne 6) { throw "Expected six aggregate rows inside transaction, got $insideCount" }

    $counts = $connection.CreateCommand()
    $counts.Transaction = $transaction
    $counts.CommandText = @'
SELECT
  (SELECT seen_count FROM dbo.QuantityAlias WHERE raw_name=@name AND signature=@signature),
  (SELECT accepted_count FROM dbo.SignatureBoxMap WHERE signature=@signature AND method=@method AND box_id=@box),
  (SELECT sample_count FROM dbo.SignatureEntryMap WHERE signature=@signature AND target_code=@code AND method=@method AND entry_code=@entry),
  (SELECT sample_count FROM dbo.QuantityFormulaRule WHERE anchor_signature=@signature AND target_code=@code)
'@
    [void]$counts.Parameters.AddWithValue('@name', $rawName)
    [void]$counts.Parameters.AddWithValue('@signature', $signature)
    [void]$counts.Parameters.AddWithValue('@box', $boxId)
    [void]$counts.Parameters.AddWithValue('@code', $targetCode)
    [void]$counts.Parameters.AddWithValue('@method', $method)
    [void]$counts.Parameters.AddWithValue('@entry', $entryCode)
    $reader = $counts.ExecuteReader()
    try {
        if (-not $reader.Read() -or $reader.GetInt32(0) -ne 2 -or $reader.GetInt32(1) -ne 2 -or $reader.GetInt32(2) -ne 2 -or $reader.GetInt32(3) -ne 2) {
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
  (SELECT sample_count FROM dbo.QuantityFormulaRule WHERE anchor_signature=@signature AND target_code=@code)
FROM dbo.SignatureBoxMap WHERE signature=@signature AND method=@method AND box_id=@box
'@
    [void]$delta.Parameters.AddWithValue('@signature', $signature)
    [void]$delta.Parameters.AddWithValue('@code', $targetCode)
    [void]$delta.Parameters.AddWithValue('@method', $method)
    [void]$delta.Parameters.AddWithValue('@entry', $entryCode)
    [void]$delta.Parameters.AddWithValue('@box', $boxId)
    $deltaReader = $delta.ExecuteReader()
    try {
        if (-not $deltaReader.Read() -or $deltaReader.GetInt32(0) -ne 2 -or $deltaReader.GetInt32(1) -ne 0 -or
            $deltaReader.GetInt32(2) -ne 1 -or $deltaReader.GetInt32(3) -ne 10 -or
            $deltaReader.GetInt32(4) -ne 2 -or $deltaReader.GetInt32(5) -ne 2) {
            throw 'Rejection delta did not lower weight without strengthening entry/formula evidence'
        }
    }
    finally { $deltaReader.Dispose() }

    $groupType.GetField('CorrectedCount', $allFlags).SetValue($group, 1)
    $groupType.GetField('RejectedCount', $allFlags).SetValue($group, 0)
    [void]$upsert.Invoke($null, $invokeArgs)
    $correction = $connection.CreateCommand()
    $correction.Transaction = $transaction
    $correction.CommandText = 'SELECT corrected_count,rejected_count,weight FROM dbo.SignatureBoxMap WHERE signature=@signature AND method=@method AND box_id=@box'
    [void]$correction.Parameters.AddWithValue('@signature', $signature)
    [void]$correction.Parameters.AddWithValue('@method', $method)
    [void]$correction.Parameters.AddWithValue('@box', $boxId)
    $correctionReader = $correction.ExecuteReader()
    try {
        if (-not $correctionReader.Read() -or $correctionReader.GetInt32(0) -ne 1 -or
            $correctionReader.GetInt32(1) -ne 1 -or $correctionReader.GetInt32(2) -ne 30) {
            throw 'Correction delta did not raise the corrected relation weight'
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
  (SELECT COUNT(*) FROM dbo.SignatureEntryMap WHERE signature=@signature AND target_code=@code)
'@
if ($formulaTablesExisted) {
    $verify.CommandText += ' + (SELECT COUNT(*) FROM dbo.QuantityFormulaRule WHERE anchor_signature=@signature AND target_code=@code) + (SELECT COUNT(*) FROM dbo.QuantityFormulaOperand WHERE operand_signature=@signature)'
}
[void]$verify.Parameters.AddWithValue('@name', $rawName)
[void]$verify.Parameters.AddWithValue('@signature', $signature)
[void]$verify.Parameters.AddWithValue('@code', $targetCode)
$afterCount = [int]$verify.ExecuteScalar()
$connection.Dispose()
if ($afterCount -ne 0) { throw "Rollback left $afterCount test rows" }

Write-Host 'Test-LearningDbIncrementalAggregate: PASS (visible in transaction, no residue after rollback)'
