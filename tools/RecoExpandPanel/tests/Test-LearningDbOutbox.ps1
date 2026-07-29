$ErrorActionPreference = 'Stop'

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..'))
$sourceDll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL)) { $env:RECO_EXPAND_DLL } else { Join-Path $repoRoot 'RecoQuotaRecommend\bin\RecoExpandPanel.dll' }
if (-not (Test-Path -LiteralPath $sourceDll)) { throw "Missing DLL: $sourceDll" }

$testRoot = Join-Path $repoRoot ('RecoQuotaRecommend\obj\learningdb-outbox-' + [Guid]::NewGuid().ToString('N'))
[void][System.IO.Directory]::CreateDirectory($testRoot)
$dll = Join-Path $testRoot 'RecoExpandPanel.dll'
Copy-Item -LiteralPath $sourceDll -Destination $dll -Force
$sourceDir = Split-Path -Parent $sourceDll
foreach ($dependency in @('NPOI.dll', 'NPOI.OpenXmlFormats.dll', 'NPOI.OpenXml4Net.dll', 'NPOI.OOXML.dll', 'ICSharpCode.SharpZipLib.dll')) {
    $dependencyPath = Join-Path $sourceDir $dependency
    if (Test-Path -LiteralPath $dependencyPath) {
        Copy-Item -LiteralPath $dependencyPath -Destination $testRoot -Force
        [void][System.Reflection.Assembly]::LoadFrom((Join-Path $testRoot $dependency))
    }
}

$connectionString = $null
$rawName = $null
$signature = $null
$targetCode = $null
$boxId = $null
try {
    $panelType = [System.Reflection.Assembly]::LoadFrom($dll).GetType('RecoNet.FormPanel', $true)
    $allFlags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
    $nestedFlags = [System.Reflection.BindingFlags]'Public,NonPublic'
    $normalizeMethod = $panelType.GetMethod('NormalizeLearningDbMethod', $allFlags)
    $method30 = '30' + [char]0x53F7 + [char]0x6587
    $method101 = '101' + [char]0x53F7 + [char]0x6587 + [char]0x4F30 + [char]0x7B97
    foreach ($methodCase in @(
        @{ Input = $method30; Expected = '2020' },
        @{ Input = $method101; Expected = '2020' },
        @{ Input = '101-estimate'; Expected = '2020' },
        @{ Input = 'TB10801-2024'; Expected = '2024' }
    )) {
        $methodArgs = New-Object 'object[]' 1
        $methodArgs[0] = $methodCase.Input
        $actualMethod = [string]$normalizeMethod.Invoke($null, $methodArgs)
        if ($actualMethod -ne $methodCase.Expected) {
            throw "Method normalization mismatch: $($methodCase.Input) -> $actualMethod"
        }
    }
    $blankMethodArgs = New-Object 'object[]' 1
    $blankMethodArgs[0] = ''
    if ([string]$normalizeMethod.Invoke($null, $blankMethodArgs) -ne '') { throw 'Blank historical method must remain read-only compatibility data' }
    $groupType = $panelType.GetNestedType('MappingFeedbackGroup', $nestedFlags)
    $targetType = $panelType.GetNestedType('MappingFeedbackTarget', $nestedFlags)
    $operandType = $panelType.GetNestedType('QuantityFormulaOperandInfo', $nestedFlags)
    $group = [Activator]::CreateInstance($groupType, $true).PSObject.BaseObject
    $target = [Activator]::CreateInstance($targetType, $true).PSObject.BaseObject
    $operand = [Activator]::CreateInstance($operandType, $true).PSObject.BaseObject

    $suffix = [Guid]::NewGuid().ToString('N')
    $rawName = 'CODEXOUTBOX' + $suffix
    $signature = $rawName + '|'
    $targetCode = 'OUT-' + $suffix.Substring(0, 20)
    $boxId = 'box-out-' + $suffix.Substring(0, 24)
    $entryCode = 'ENTRY-' + $suffix.Substring(0, 20)

    $groupType.GetField('QuantityName', $allFlags).SetValue($group, $rawName)
    $groupType.GetField('QuantityUnit', $allFlags).SetValue($group, 'm2')
    $groupType.GetField('EntryCode', $allFlags).SetValue($group, $entryCode)
    $groupType.GetField('EntryName', $allFlags).SetValue($group, 'outbox entry')
    $groupType.GetField('Method', $allFlags).SetValue($group, '2024')
    $groupType.GetField('BoxId', $allFlags).SetValue($group, $boxId)
    $targetType.GetField('Kind', $allFlags).SetValue($target, 'quota')
    $targetType.GetField('Code', $allFlags).SetValue($target, $targetCode)
    $targetType.GetField('Name', $allFlags).SetValue($target, 'outbox target')
    $targetType.GetField('Unit', $allFlags).SetValue($target, 'm3')
    $targetType.GetField('FormulaTemplate', $allFlags).SetValue($target, 'V0*0.2')
    [void]$groupType.GetField('Targets', $allFlags).GetValue($group).PSObject.BaseObject.Add($target)
    $operandType.GetField('Name', $allFlags).SetValue($operand, $rawName)
    $operandType.GetField('Unit', $allFlags).SetValue($operand, 'm2')
    $operandType.GetField('Signature', $allFlags).SetValue($operand, $signature)
    [void]$groupType.GetField('FormulaOperands', $allFlags).GetValue($group).PSObject.BaseObject.Add($operand)

    $listType = [System.Collections.Generic.List``1].MakeGenericType($groupType)
    $groups = [Activator]::CreateInstance($listType).PSObject.BaseObject
    [void]$groups.Add($group)
    $invokeArgs = New-Object 'object[]' 2
    $invokeArgs[0] = 'name-match'
    $invokeArgs[1] = $groups

    $getConnection = $panelType.GetMethod('GetLearningDbConnectionString', $allFlags)
    $connectionString = [string]$getConnection.Invoke($null, $null)
    $schemaProbe = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $schemaProbe.Open()
    try {
        $methodColumn = $schemaProbe.CreateCommand()
        $methodColumn.CommandText = "SELECT CASE WHEN COL_LENGTH('dbo.SignatureBoxMap','method') IS NULL THEN 0 ELSE 1 END"
        if ([int]$methodColumn.ExecuteScalar() -eq 0) {
            $rawName = $null
            Write-Host 'Test-LearningDbOutbox: SKIP (schema.sql SignatureBoxMap.method migration not applied)'
            return
        }
    }
    finally { $schemaProbe.Dispose() }
    $connectionField = $panelType.GetField('learningDbConnectionString', $allFlags)
    $record = $panelType.GetMethod('RecordBindingEventsToLearningDb', $allFlags)
    $replay = $panelType.GetMethod('ReplayPendingLearningDbEvents', $allFlags)
    $pendingKeysMethod = $panelType.GetMethod('LoadPendingLearningMappingKeys', $allFlags)
    $consume = $panelType.GetMethod('ConsumeLearningDbDurableResult', $allFlags)

    $connectionField.SetValue($null, 'Server=127.0.0.1,1;Database=NoDb;User ID=x;Password=x;Connect Timeout=1')
    $groupType.GetField('Method', $allFlags).SetValue($group, '')
    if ([bool]$record.Invoke($null, $invokeArgs)) { throw 'Blank method binding was incorrectly marked durable' }
    $blankConsumeArgs = New-Object 'object[]' 1
    $blankConsumeArgs[0] = $groups
    if ([bool]$consume.Invoke($null, $blankConsumeArgs)) { throw 'Blank method binding must not confirm MappingFeedbackRecorded' }
    $outboxPath = Join-Path $testRoot 'RecoQuotaData\learning-db-outbox.jsonl'
    if (Test-Path -LiteralPath $outboxPath) { throw 'Blank method binding entered the active outbox' }

    $groupType.GetField('Method', $allFlags).SetValue($group, '2024')
    if (-not [bool]$record.Invoke($null, $invokeArgs)) { throw 'SQL failure was not durably queued to the local outbox' }
    $consumeArgs = New-Object 'object[]' 1
    $consumeArgs[0] = $groups
    if (-not [bool]$consume.Invoke($null, $consumeArgs)) { throw 'Durable outbox result did not allow MappingFeedbackRecorded confirmation' }
    if ([bool]$consume.Invoke($null, $consumeArgs)) { throw 'Durable result must be consumed only once' }

    if (-not (Test-Path -LiteralPath $outboxPath)) { throw 'Outbox file was not created next to the isolated plugin' }
    $outboxLines = [System.IO.File]::ReadAllLines($outboxPath, [System.Text.Encoding]::UTF8)
    if ($outboxLines.Count -ne 1) { throw "Expected one outbox batch, got $($outboxLines.Count)" }
    $savedLine = $outboxLines[0]
    if ($savedLine -match 'Password=|Server=|User ID=') { throw 'Outbox leaked a database connection string' }
    foreach ($marker in $rawName, $targetCode, 'V0*0.2', $signature, $boxId) {
        if ($savedLine -notlike ('*' + $marker + '*')) { throw "Outbox lost required replay metadata: $marker" }
    }
    $pendingKeys = $pendingKeysMethod.Invoke($null, $null).PSObject.BaseObject
    if (-not $pendingKeys.Contains($signature + "`n" + $boxId)) { throw 'Pending mapping key was not exposed to SmartFill' }

    $laterBatchId = [Guid]::NewGuid().ToString('N')
    $laterLine = [regex]::Replace($savedLine, '"batch_id":"[0-9a-fA-F]{32}"', '"batch_id":"' + $laterBatchId + '"', 1)
    [System.IO.File]::WriteAllLines($outboxPath, [string[]]@($savedLine, $laterLine), (New-Object System.Text.UTF8Encoding($false)))
    [void]$replay.Invoke($null, $null)
    $retainedLines = [System.IO.File]::ReadAllLines($outboxPath, [System.Text.Encoding]::UTF8)
    if ($retainedLines.Count -ne 2) { throw 'Transient SQL failure did not retain the failed head batch and untouched later batch' }
    $prematureDeadLetterPath = Join-Path $testRoot 'RecoQuotaData\learning-db-outbox.dead-letter.jsonl'
    if (Test-Path -LiteralPath $prematureDeadLetterPath) { throw 'Transient SQL failure was incorrectly moved to dead-letter storage' }

    $poisonBatchId = [Guid]::NewGuid().ToString('N')
    $poisonLine = [regex]::Replace($savedLine, '"batch_id":"[0-9a-fA-F]{32}"', '"batch_id":"' + $poisonBatchId + '"', 1)
    $poisonLine = $poisonLine.Replace('"g0_method":"2024"', '"g0_method":""')
    if ($poisonLine -eq $savedLine -or $poisonLine -notlike ('*' + $poisonBatchId + '*') -or
        $poisonLine -like '*"g0_method":"2024"*') { throw 'Failed to prepare a deterministic poison outbox batch' }
    [System.IO.File]::WriteAllLines($outboxPath, [string[]]@($poisonLine, $savedLine), (New-Object System.Text.UTF8Encoding($false)))

    $connectionField.SetValue($null, $connectionString)
    [void]$replay.Invoke($null, $null)
    if ((Get-Content -LiteralPath $outboxPath -ErrorAction SilentlyContinue | Measure-Object -Line).Lines -ne 0) { throw 'Committed outbox batch was not cleared' }
    $deadLetterPath = Join-Path $testRoot 'RecoQuotaData\learning-db-outbox.dead-letter.jsonl'
    if (-not (Test-Path -LiteralPath $deadLetterPath)) { throw 'Poison outbox batch was not moved to dead-letter storage' }
    $deadLetterLines = [System.IO.File]::ReadAllLines($deadLetterPath, [System.Text.Encoding]::UTF8)
    if ($deadLetterLines.Count -ne 1 -or $deadLetterLines[0] -notlike ('*' + $poisonBatchId + '*') -or
        $deadLetterLines[0] -notlike '*dead_letter_reason*') { throw 'Dead-letter record did not preserve the rejected poison batch and reason' }
    if ($deadLetterLines[0] -match 'Password=|Server=|User ID=') { throw 'Dead-letter storage leaked a database connection string' }

    $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $connection.Open()
    try {
        $counts = $connection.CreateCommand()
        $counts.CommandText = @'
SELECT
 (SELECT COUNT(*) FROM dbo.BindingLog WHERE quantity_name=@name AND target_code=@code),
 (SELECT seen_count FROM dbo.QuantityAlias WHERE raw_name=@name AND signature=@signature),
 (SELECT accepted_count FROM dbo.SignatureBoxMap WHERE signature=@signature AND method='2024' AND box_id=@box),
 (SELECT sample_count FROM dbo.SignatureEntryMap WHERE signature=@signature AND target_code=@code),
 (SELECT sample_count FROM dbo.QuantityFormulaRule WHERE anchor_signature=@signature AND target_code=@code)
'@
        [void]$counts.Parameters.AddWithValue('@name', $rawName)
        [void]$counts.Parameters.AddWithValue('@code', $targetCode)
        [void]$counts.Parameters.AddWithValue('@signature', $signature)
        [void]$counts.Parameters.AddWithValue('@box', $boxId)
        $reader = $counts.ExecuteReader()
        try {
            if (-not $reader.Read() -or $reader.GetInt32(0) -ne 1 -or $reader.GetInt32(1) -ne 1 -or
                $reader.GetInt32(2) -ne 1 -or $reader.GetInt32(3) -ne 1 -or $reader.GetInt32(4) -ne 1) {
                throw 'First outbox replay did not commit exactly one event and one aggregate increment'
            }
        }
        finally { $reader.Dispose() }
    }
    finally { $connection.Dispose() }

    [System.IO.File]::WriteAllText($outboxPath, $savedLine + [Environment]::NewLine, (New-Object System.Text.UTF8Encoding($false)))
    [void]$replay.Invoke($null, $null)
    $verify = New-Object System.Data.SqlClient.SqlConnection($connectionString)
    $verify.Open()
    try {
        $unchanged = $verify.CreateCommand()
        $unchanged.CommandText = @'
SELECT
 (SELECT COUNT(*) FROM dbo.BindingLog WHERE quantity_name=@name AND target_code=@code) +
 (SELECT seen_count FROM dbo.QuantityAlias WHERE raw_name=@name AND signature=@signature) +
 (SELECT accepted_count FROM dbo.SignatureBoxMap WHERE signature=@signature AND method='2024' AND box_id=@box) +
 (SELECT sample_count FROM dbo.SignatureEntryMap WHERE signature=@signature AND target_code=@code) +
 (SELECT sample_count FROM dbo.QuantityFormulaRule WHERE anchor_signature=@signature AND target_code=@code)
'@
        [void]$unchanged.Parameters.AddWithValue('@name', $rawName)
        [void]$unchanged.Parameters.AddWithValue('@code', $targetCode)
        [void]$unchanged.Parameters.AddWithValue('@signature', $signature)
        [void]$unchanged.Parameters.AddWithValue('@box', $boxId)
        if ([int]$unchanged.ExecuteScalar() -ne 5) { throw 'Duplicate outbox replay incremented learning counts again' }
    }
    finally { $verify.Dispose() }
}
finally {
    if (-not [String]::IsNullOrWhiteSpace($connectionString) -and -not [String]::IsNullOrWhiteSpace($rawName)) {
        $cleanup = New-Object System.Data.SqlClient.SqlConnection($connectionString)
        try {
            $cleanup.Open()
            $cmd = $cleanup.CreateCommand()
            $cmd.CommandText = @'
DELETE FROM dbo.QuantityFormulaOperand WHERE rule_hash IN (SELECT rule_hash FROM dbo.QuantityFormulaRule WHERE anchor_signature=@signature AND target_code=@code);
DELETE FROM dbo.QuantityFormulaRule WHERE anchor_signature=@signature AND target_code=@code;
DELETE FROM dbo.SignatureEntryMap WHERE signature=@signature AND target_code=@code;
DELETE FROM dbo.SignatureBoxMap WHERE signature=@signature AND method='2024' AND box_id=@box;
DELETE FROM dbo.QuotaBoxTarget WHERE box_id=@box AND target_code=@code;
DELETE FROM dbo.QuotaBox WHERE box_id=@box;
DELETE FROM dbo.QuantityAlias WHERE raw_name=@name AND signature=@signature;
DELETE FROM dbo.BindingLog WHERE quantity_name=@name AND target_code=@code;
'@
            [void]$cmd.Parameters.AddWithValue('@name', $rawName)
            [void]$cmd.Parameters.AddWithValue('@code', $targetCode)
            [void]$cmd.Parameters.AddWithValue('@signature', $signature)
            [void]$cmd.Parameters.AddWithValue('@box', $boxId)
            [void]$cmd.ExecuteNonQuery()
        }
        finally { $cleanup.Dispose() }
    }
    $outboxDataDir = Join-Path $testRoot 'RecoQuotaData'
    if (Test-Path -LiteralPath $outboxDataDir) {
        Remove-Item -LiteralPath $outboxDataDir -Recurse -Force
    }
}

Write-Host 'Test-LearningDbOutbox: PASS (failure queued, replay committed, duplicate replay idempotent, cleanup complete)'
