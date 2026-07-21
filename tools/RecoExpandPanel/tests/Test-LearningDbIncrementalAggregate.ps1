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
$group = [Activator]::CreateInstance($groupType, $true).PSObject.BaseObject
$target = [Activator]::CreateInstance($targetType, $true).PSObject.BaseObject

$suffix = [Guid]::NewGuid().ToString('N')
$rawName = 'CODEXSQLROLLBACK' + $suffix
$signature = $rawName + '|'
$targetCode = 'TEST-' + $suffix.Substring(0, 20)
$entryCode = 'ENTRY-' + $suffix.Substring(0, 20)
$entryName = 'rollback entry'
$method = '2024'
$boxId = 'box-test-' + $suffix.Substring(0, 24)

$groupType.GetField('QuantityName', $allFlags).SetValue($group, $rawName)
$groupType.GetField('QuantityUnit', $allFlags).SetValue($group, 't')
$groupType.GetField('EntryCode', $allFlags).SetValue($group, $entryCode)
$groupType.GetField('EntryName', $allFlags).SetValue($group, $entryName)
$groupType.GetField('Method', $allFlags).SetValue($group, $method)
$groupType.GetField('BoxId', $allFlags).SetValue($group, $boxId)
$targetType.GetField('Kind', $allFlags).SetValue($target, 'quota')
$targetType.GetField('Code', $allFlags).SetValue($target, $targetCode)
$targetType.GetField('Name', $allFlags).SetValue($target, 'rollback target')
$targetType.GetField('Unit', $allFlags).SetValue($target, 't')
$targets = $groupType.GetField('Targets', $allFlags).GetValue($group).PSObject.BaseObject
[void]$targets.Add($target)

$getConnection = $panelType.GetMethod('GetLearningDbConnectionString', $allFlags)
$upsert = $panelType.GetMethod('UpsertBindingGroupAggregates', $allFlags)
$connectionString = [string]$getConnection.Invoke($null, $null)
$connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
$connection.Open()
$transaction = $connection.BeginTransaction()
try {
    $invokeArgs = New-Object 'object[]' 3
    $invokeArgs[0] = $connection.PSObject.BaseObject
    $invokeArgs[1] = $transaction.PSObject.BaseObject
    $invokeArgs[2] = $group
    [void]$upsert.Invoke($null, $invokeArgs)

    $command = $connection.CreateCommand()
    $command.Transaction = $transaction
    $command.CommandText = @'
SELECT
  (SELECT COUNT(*) FROM dbo.QuantityAlias WHERE raw_name=@name AND signature=@signature) +
  (SELECT COUNT(*) FROM dbo.SignatureBoxMap WHERE signature=@signature AND box_id=@box) +
  (SELECT COUNT(*) FROM dbo.QuotaBoxTarget WHERE box_id=@box AND target_code=@code) +
  (SELECT COUNT(*) FROM dbo.SignatureEntryMap WHERE signature=@signature AND target_code=@code AND method=@method AND entry_code=@entry)
'@
    [void]$command.Parameters.AddWithValue('@name', $rawName)
    [void]$command.Parameters.AddWithValue('@signature', $signature)
    [void]$command.Parameters.AddWithValue('@box', $boxId)
    [void]$command.Parameters.AddWithValue('@code', $targetCode)
    [void]$command.Parameters.AddWithValue('@method', $method)
    [void]$command.Parameters.AddWithValue('@entry', $entryCode)
    $insideCount = [int]$command.ExecuteScalar()
    if ($insideCount -ne 4) { throw "Expected four aggregate rows inside transaction, got $insideCount" }

    $metadata = $connection.CreateCommand()
    $metadata.Transaction = $transaction
    $metadata.CommandText = 'SELECT target_name + ''|'' + target_unit FROM dbo.QuotaBoxTarget WHERE box_id=@box AND target_code=@code'
    [void]$metadata.Parameters.AddWithValue('@box', $boxId)
    [void]$metadata.Parameters.AddWithValue('@code', $targetCode)
    if ([string]$metadata.ExecuteScalar() -ne 'rollback target|t') { throw 'QuotaBoxTarget metadata was not persisted' }

    $alias = $connection.CreateCommand()
    $alias.Transaction = $transaction
    $alias.CommandText = 'SELECT COUNT(*) FROM dbo.QuantityAlias WHERE raw_name=@name AND signature=@signature AND quantity_unit=''t'''
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
[void]$verify.Parameters.AddWithValue('@name', $rawName)
[void]$verify.Parameters.AddWithValue('@signature', $signature)
[void]$verify.Parameters.AddWithValue('@code', $targetCode)
$afterCount = [int]$verify.ExecuteScalar()
$connection.Dispose()
if ($afterCount -ne 0) { throw "Rollback left $afterCount test rows" }

Write-Host 'Test-LearningDbIncrementalAggregate: PASS (visible in transaction, no residue after rollback)'
