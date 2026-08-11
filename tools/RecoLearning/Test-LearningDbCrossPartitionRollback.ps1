param(
  [string]$TargetDatabase = '',
  [switch]$ExecuteLive,
  [string]$EvidenceDirectory = ''
)

$ErrorActionPreference = 'Stop'

if (-not $ExecuteLive) {
  throw 'Live execution is disabled. Pass -ExecuteLive explicitly.'
}
if (-not [string]::Equals($TargetDatabase, 'RecoLearning', [System.StringComparison]::Ordinal)) {
  throw 'TargetDatabase must be exactly RecoLearning.'
}

$hostNames = @('RejjNet2020','ReJJGSNet2024','ReJJQDNet2024')
$runningHosts = @($hostNames | ForEach-Object { Get-Process -Name $_ -ErrorAction SilentlyContinue })
if ($runningHosts.Count -ne 0) {
  throw 'All supported host processes must be closed before the E3 live rollback test.'
}

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..'))
$commonScript = Join-Path $PSScriptRoot 'Common.ps1'
$dllPath = Join-Path $repoRoot 'RecoQuotaRecommend\bin\RecoExpandPanel.dll'
if (-not (Test-Path -LiteralPath $commonScript)) { throw 'Missing RecoLearning Common.ps1.' }
if (-not (Test-Path -LiteralPath $dllPath)) { throw 'Missing RecoExpandPanel.dll.' }
. $commonScript

if ([string]::IsNullOrWhiteSpace($EvidenceDirectory)) {
  $stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
  $EvidenceDirectory = Join-Path $repoRoot ('artifacts\e3-sqlonly\' + $stamp)
}
[void][System.IO.Directory]::CreateDirectory($EvidenceDirectory)
$evidencePath = Join-Path $EvidenceDirectory 'e3-cross-partition-rollback.json'

$runId = [Guid]::NewGuid().ToString('N')
$quantityName = 'CODEXE3' + $runId.ToUpperInvariant()
$runStarted = [DateTime]::UtcNow
$stage = 'initialize'
$transaction = $null
$connection = $null
$failure = $null
$rollbackState = 'not-started'
$assertions = New-Object System.Collections.Generic.List[object]
$residue = @{}
$baselineCounts = @{}
$afterCounts = @{}

function Add-Assertion([string]$Name, [long]$Actual, [long]$Expected) {
  $passed = $Actual -eq $Expected
  $script:assertions.Add([pscustomobject]@{
    name = $Name
    actual = $Actual
    expected = $Expected
    passed = $passed
  })
  if (-not $passed) { throw "Assertion failed: $Name (actual=$Actual expected=$Expected)" }
}

function Assert-ExactLearningDatabase([System.Data.SqlClient.SqlConnection]$Connection) {
  $command = $Connection.CreateCommand()
  try {
    $command.CommandText = 'SELECT DB_NAME()'
    $actual = [string]$command.ExecuteScalar()
    if (-not [string]::Equals($actual, 'RecoLearning', [System.StringComparison]::Ordinal)) {
      throw 'Connected database is not the exact RecoLearning database.'
    }
  }
  finally { $command.Dispose() }
}

function Get-TableCounts([System.Data.SqlClient.SqlConnection]$Connection) {
  $command = $Connection.CreateCommand()
  try {
    $command.CommandText = @'
SELECT 'BindingLog' AS table_name, COUNT_BIG(*) AS row_count FROM dbo.BindingLog
UNION ALL SELECT 'QuantityAlias', COUNT_BIG(*) FROM dbo.QuantityAlias
UNION ALL SELECT 'QuotaBox', COUNT_BIG(*) FROM dbo.QuotaBox
UNION ALL SELECT 'QuotaBoxTarget', COUNT_BIG(*) FROM dbo.QuotaBoxTarget
UNION ALL SELECT 'SignatureBoxMap', COUNT_BIG(*) FROM dbo.SignatureBoxMap
UNION ALL SELECT 'SignatureEntryMap', COUNT_BIG(*) FROM dbo.SignatureEntryMap
UNION ALL SELECT 'EngineeringTemplate', COUNT_BIG(*) FROM dbo.EngineeringTemplate
'@
    $reader = $command.ExecuteReader()
    try {
      $counts = @{}
      while ($reader.Read()) { $counts[$reader.GetString(0)] = $reader.GetInt64(1) }
      return $counts
    }
    finally { $reader.Dispose() }
  }
  finally { $command.Dispose() }
}

function Assert-CurrentPartitionSchema([System.Data.SqlClient.SqlConnection]$Connection) {
  $command = $Connection.CreateCommand()
  try {
    $command.CommandText = @'
SELECT CASE WHEN
  OBJECT_ID('dbo.BindingLog','U') IS NOT NULL AND
  OBJECT_ID('dbo.QuantityAlias','U') IS NOT NULL AND
  OBJECT_ID('dbo.QuotaBox','U') IS NOT NULL AND
  OBJECT_ID('dbo.QuotaBoxTarget','U') IS NOT NULL AND
  OBJECT_ID('dbo.SignatureBoxMap','U') IS NOT NULL AND
  OBJECT_ID('dbo.SignatureEntryMap','U') IS NOT NULL AND
  OBJECT_ID('dbo.EngineeringTemplate','U') IS NOT NULL AND
  COL_LENGTH('dbo.BindingLog','software_partition') IS NOT NULL AND
  COL_LENGTH('dbo.BindingLog','method_no') IS NOT NULL AND
  COL_LENGTH('dbo.SignatureBoxMap','software_partition') IS NOT NULL AND
  COL_LENGTH('dbo.SignatureEntryMap','software_partition') IS NOT NULL AND
  COL_LENGTH('dbo.SignatureEntryMap','method_no') IS NOT NULL AND
  COL_LENGTH('dbo.EngineeringTemplate','software_partition') IS NOT NULL AND
  COL_LENGTH('dbo.EngineeringTemplate','method_no') IS NOT NULL
THEN 1 ELSE 0 END
'@
    if ([int]$command.ExecuteScalar() -ne 1) { throw 'Required partition schema is not present.' }
  }
  finally { $command.Dispose() }

  $keyCommand = $Connection.CreateCommand()
  try {
    $keyCommand.CommandText = @'
SELECT OBJECT_NAME(i.object_id) AS table_name, c.name AS column_name, ic.key_ordinal
FROM sys.indexes i
JOIN sys.index_columns ic ON ic.object_id=i.object_id AND ic.index_id=i.index_id
JOIN sys.columns c ON c.object_id=ic.object_id AND c.column_id=ic.column_id
WHERE i.is_primary_key=1
  AND i.object_id IN (OBJECT_ID('dbo.SignatureBoxMap'),OBJECT_ID('dbo.SignatureEntryMap'),OBJECT_ID('dbo.EngineeringTemplate'))
ORDER BY table_name, ic.key_ordinal
'@
    $reader = $keyCommand.ExecuteReader()
    try {
      $keys = @{}
      while ($reader.Read()) {
        $table = $reader.GetString(0)
        if (-not $keys.ContainsKey($table)) { $keys[$table] = New-Object System.Collections.Generic.List[string] }
        $keys[$table].Add($reader.GetString(1))
      }
    }
    finally { $reader.Dispose() }
    $expected = @{
      SignatureBoxMap = 'software_partition,signature,box_id'
      SignatureEntryMap = 'software_partition,method_no,signature,target_code,entry_code'
      EngineeringTemplate = 'software_partition,method_no,engineering_type,entry_code,box_id'
    }
    foreach ($table in $expected.Keys) {
      $actual = if ($keys.ContainsKey($table)) { [string]::Join(',', $keys[$table].ToArray()) } else { '' }
      if (-not [string]::Equals($actual, $expected[$table], [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Partition primary key mismatch: $table"
      }
    }
  }
  finally { $keyCommand.Dispose() }
}

function Set-PrivateField($Type, $Instance, [string]$Name, $Value, $Flags) {
  $field = $Type.GetField($Name, $Flags)
  if ($null -eq $field) { throw "Missing reflected field: $Name" }
  $field.SetValue($Instance, $Value)
}

function New-ProbeGroup($GroupType, $TargetType, $Flags, [string]$Partition, [string]$MethodNo,
  [string]$EntryCode, [string]$BoxId, [string]$TargetCode) {
  $group = [Activator]::CreateInstance($GroupType, $true).PSObject.BaseObject
  Set-PrivateField $GroupType $group 'QuantityName' $script:quantityName $Flags
  Set-PrivateField $GroupType $group 'QuantityUnit' 'm' $Flags
  Set-PrivateField $GroupType $group 'Method' $Partition $Flags
  Set-PrivateField $GroupType $group 'MethodNo' $MethodNo $Flags
  Set-PrivateField $GroupType $group 'SoftwarePartition' $Partition $Flags
  Set-PrivateField $GroupType $group 'ProjectId' ('codex-e3-rollback-' + $script:runId) $Flags
  Set-PrivateField $GroupType $group 'EntryCode' $EntryCode $Flags
  Set-PrivateField $GroupType $group 'EntryName' ('E3 entry ' + $Partition) $Flags
  Set-PrivateField $GroupType $group 'BoxId' $BoxId $Flags
  Set-PrivateField $GroupType $group 'AcceptedCount' 1 $Flags
  Set-PrivateField $GroupType $group 'UserAction' 'accepted' $Flags

  $target = [Activator]::CreateInstance($TargetType, $true).PSObject.BaseObject
  Set-PrivateField $TargetType $target 'Kind' 'quota' $Flags
  Set-PrivateField $TargetType $target 'Code' $TargetCode $Flags
  Set-PrivateField $TargetType $target 'Name' ('E3 target ' + $Partition) $Flags
  Set-PrivateField $TargetType $target 'Unit' 'm' $Flags
  Set-PrivateField $TargetType $target 'EntryCode' $EntryCode $Flags
  Set-PrivateField $TargetType $target 'EntryName' ('E3 entry ' + $Partition) $Flags
  $targets = $GroupType.GetField('Targets', $Flags).GetValue($group).PSObject.BaseObject
  [void]$targets.Add($target)

  return [pscustomobject]@{
    group = $group
    partition = $Partition
    method_no = $MethodNo
    entry_code = $EntryCode
    box_id = $BoxId
    target_code = $TargetCode
  }
}

function Get-ScalarCount([System.Data.SqlClient.SqlConnection]$Connection,
  [System.Data.SqlClient.SqlTransaction]$Transaction, [string]$Sql, [hashtable]$Parameters) {
  $command = $Connection.CreateCommand()
  try {
    if ($null -ne $Transaction) { $command.Transaction = $Transaction }
    $command.CommandTimeout = 5
    $command.CommandText = $Sql
    foreach ($key in $Parameters.Keys) { [void]$command.Parameters.AddWithValue('@' + $key, $Parameters[$key]) }
    return [int]$command.ExecuteScalar()
  }
  finally { $command.Dispose() }
}

function Get-PartitionRelationCount([System.Data.SqlClient.SqlConnection]$Connection,
  [System.Data.SqlClient.SqlTransaction]$Transaction, [string]$Partition, [string]$Signature, [string]$TargetCode) {
  return Get-ScalarCount $Connection $Transaction @'
SELECT COUNT(*)
FROM dbo.SignatureBoxMap m
JOIN dbo.QuotaBox b ON b.box_id=m.box_id AND b.status='active'
JOIN dbo.QuotaBoxTarget t ON t.box_id=m.box_id
WHERE m.weight>0 AND m.software_partition=@software_partition
  AND m.signature=@signature AND t.target_kind='quota' AND t.target_code=@target_code
'@ @{ software_partition=$Partition; signature=$Signature; target_code=$TargetCode }
}

function Get-PartitionRelationTotal([System.Data.SqlClient.SqlConnection]$Connection,
  [System.Data.SqlClient.SqlTransaction]$Transaction, [string]$Partition, [string]$Signature) {
  return Get-ScalarCount $Connection $Transaction @'
SELECT COUNT(*)
FROM dbo.SignatureBoxMap m
JOIN dbo.QuotaBox b ON b.box_id=m.box_id AND b.status='active'
JOIN dbo.QuotaBoxTarget t ON t.box_id=m.box_id
WHERE m.weight>0 AND m.software_partition=@software_partition AND m.signature=@signature
'@ @{ software_partition=$Partition; signature=$Signature }
}

function Get-EntryCount([System.Data.SqlClient.SqlConnection]$Connection,
  [System.Data.SqlClient.SqlTransaction]$Transaction, $Probe, [string]$Partition, [string]$MethodNo) {
  return Get-ScalarCount $Connection $Transaction @'
SELECT COUNT(*) FROM dbo.SignatureEntryMap
WHERE software_partition=@software_partition AND method_no=@method_no
  AND signature=@signature AND target_code=@target_code AND entry_code=@entry_code
'@ @{ software_partition=$Partition; method_no=$MethodNo; signature=$script:signature; target_code=$Probe.target_code; entry_code=$Probe.entry_code }
}

function Get-EngineeringTemplateCount([System.Data.SqlClient.SqlConnection]$Connection,
  [System.Data.SqlClient.SqlTransaction]$Transaction, $Probe, [string]$Partition, [string]$MethodNo) {
  return Get-ScalarCount $Connection $Transaction @'
SELECT COUNT(*) FROM dbo.EngineeringTemplate
WHERE software_partition=@software_partition AND method_no=@method_no
  AND engineering_type=@engineering_type AND entry_code=@entry_code AND box_id=@box_id
'@ @{ software_partition=$Partition; method_no=$MethodNo; engineering_type=$Probe.entry_code.Substring(0,2); entry_code=$Probe.entry_code; box_id=$Probe.box_id }
}

function Add-ProbeAggregate($UpsertMethod, [System.Data.SqlClient.SqlConnection]$Connection,
  [System.Data.SqlClient.SqlTransaction]$Transaction, $Probe) {
  $arguments = New-Object 'object[]' 3
  $arguments[0] = $Connection.PSObject.BaseObject
  $arguments[1] = $Transaction.PSObject.BaseObject
  $arguments[2] = $Probe.group.PSObject.BaseObject
  [void]$UpsertMethod.Invoke($null, $arguments)
}

try {
  $stage = 'load-assembly'
  $dllDirectory = Split-Path -Parent $dllPath
  foreach ($dependency in @('NPOI.dll','NPOI.OpenXmlFormats.dll','NPOI.OpenXml4Net.dll','NPOI.OOXML.dll','ICSharpCode.SharpZipLib.dll')) {
    $dependencyPath = Join-Path $dllDirectory $dependency
    if (Test-Path -LiteralPath $dependencyPath) { [void][System.Reflection.Assembly]::LoadFrom($dependencyPath) }
  }
  $assembly = [System.Reflection.Assembly]::LoadFrom($dllPath)
  $panelType = $assembly.GetType('RecoNet.FormPanel', $true)
  $flags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
  $nestedFlags = [System.Reflection.BindingFlags]'Public,NonPublic'
  $groupType = $panelType.GetNestedType('MappingFeedbackGroup', $nestedFlags)
  $targetType = $panelType.GetNestedType('MappingFeedbackTarget', $nestedFlags)
  $upsertMethod = $panelType.GetMethod('UpsertBindingGroupAggregates', $flags)
  $normalizeMethod = $panelType.GetMethod('NormalizeForSignature', $flags)
  if ($null -eq $groupType -or $null -eq $targetType -or $null -eq $upsertMethod -or $null -eq $normalizeMethod) {
    throw 'Required production learning reflection surface is unavailable.'
  }
  $normalizeArguments = New-Object 'object[]' 1
  $normalizeArguments[0] = $quantityName
  $script:signature = ([string]$normalizeMethod.Invoke($null, $normalizeArguments)) + '|'

  $probe2024 = New-ProbeGroup $groupType $targetType $flags '2024' 'TB 10801—2024' '0301-01-01' ('e3-24-' + $runId) ('E3-24-' + $runId.Substring(0,20))
  $probe2020 = New-ProbeGroup $groupType $targetType $flags '2020' '30号文' '0702-02-06-02' ('e3-20-' + $runId) ('E3-20-' + $runId.Substring(0,20))

  $stage = 'open-learning-database'
  $connectionString = Get-RecoConnectionString -Database $TargetDatabase
  $connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)
  $connection.Open()
  Assert-ExactLearningDatabase $connection

  $stage = 'schema-gate'
  Assert-CurrentPartitionSchema $connection
  $baselineCounts = Get-TableCounts $connection

  $stage = 'begin-rollback-transaction'
  $transaction = $connection.BeginTransaction([System.Data.IsolationLevel]::Serializable)
  $rollbackState = 'pending'

  $stage = 'write-2024-probe'
  Add-ProbeAggregate $upsertMethod $connection $transaction $probe2024
  Add-Assertion '2024 relation visible in 2024 after first probe' (Get-PartitionRelationCount $connection $transaction '2024' $script:signature $probe2024.target_code) 1
  Add-Assertion '2024 relation invisible in 2020 after first probe' (Get-PartitionRelationTotal $connection $transaction '2020' $script:signature) 0

  $stage = 'write-2020-probe'
  Add-ProbeAggregate $upsertMethod $connection $transaction $probe2020

  Add-Assertion '2024 own relation visible' (Get-PartitionRelationCount $connection $transaction '2024' $script:signature $probe2024.target_code) 1
  Add-Assertion '2024 cannot see 2020 relation' (Get-PartitionRelationCount $connection $transaction '2024' $script:signature $probe2020.target_code) 0
  Add-Assertion '2024 relation total remains one' (Get-PartitionRelationTotal $connection $transaction '2024' $script:signature) 1
  Add-Assertion '2020 own relation visible' (Get-PartitionRelationCount $connection $transaction '2020' $script:signature $probe2020.target_code) 1
  Add-Assertion '2020 cannot see 2024 relation' (Get-PartitionRelationCount $connection $transaction '2020' $script:signature $probe2024.target_code) 0
  Add-Assertion '2020 relation total remains one' (Get-PartitionRelationTotal $connection $transaction '2020' $script:signature) 1

  Add-Assertion '2024 entry visible in 2024 method' (Get-EntryCount $connection $transaction $probe2024 '2024' $probe2024.method_no) 1
  Add-Assertion '2024 entry same-method cross-partition invisible in 2020' (Get-EntryCount $connection $transaction $probe2024 '2020' $probe2024.method_no) 0
  Add-Assertion '2024 entry wrong method invisible in 2024' (Get-EntryCount $connection $transaction $probe2024 '2024' $probe2020.method_no) 0
  Add-Assertion '2020 entry visible in 2020 method' (Get-EntryCount $connection $transaction $probe2020 '2020' $probe2020.method_no) 1
  Add-Assertion '2020 entry same-method cross-partition invisible in 2024' (Get-EntryCount $connection $transaction $probe2020 '2024' $probe2020.method_no) 0
  Add-Assertion '2020 entry wrong method invisible in 2020' (Get-EntryCount $connection $transaction $probe2020 '2020' $probe2024.method_no) 0

  Add-Assertion '2024 engineering scope visible in 2024 method' (Get-EngineeringTemplateCount $connection $transaction $probe2024 '2024' $probe2024.method_no) 1
  Add-Assertion '2024 engineering scope same-method cross-partition invisible in 2020' (Get-EngineeringTemplateCount $connection $transaction $probe2024 '2020' $probe2024.method_no) 0
  Add-Assertion '2024 engineering scope wrong method invisible in 2024' (Get-EngineeringTemplateCount $connection $transaction $probe2024 '2024' $probe2020.method_no) 0
  Add-Assertion '2020 engineering scope visible in 2020 method' (Get-EngineeringTemplateCount $connection $transaction $probe2020 '2020' $probe2020.method_no) 1
  Add-Assertion '2020 engineering scope same-method cross-partition invisible in 2024' (Get-EngineeringTemplateCount $connection $transaction $probe2020 '2024' $probe2020.method_no) 0
  Add-Assertion '2020 engineering scope wrong method invisible in 2020' (Get-EngineeringTemplateCount $connection $transaction $probe2020 '2020' $probe2024.method_no) 0

  $stage = 'transaction-assertions-complete'
}
catch {
  $failure = $_
}
finally {
  if ($null -ne $transaction) {
    try {
      $transaction.Rollback()
      $rollbackState = 'rolled-back'
    }
    catch {
      $rollbackState = 'rollback-failed:' + $_.Exception.GetType().FullName
      if ($null -eq $failure) { $failure = $_ }
    }
    finally { $transaction.Dispose() }
  }
  if ($null -ne $connection) { $connection.Dispose() }
}

try {
  $stage = 'fresh-connection-residue-gate'
  $verifyConnectionString = Get-RecoConnectionString -Database $TargetDatabase
  $verifyConnection = New-Object System.Data.SqlClient.SqlConnection($verifyConnectionString)
  try {
    $verifyConnection.Open()
    Assert-ExactLearningDatabase $verifyConnection
    $afterCounts = Get-TableCounts $verifyConnection
    $residueSql = @{
      BindingLog = 'SELECT COUNT(*) FROM dbo.BindingLog WHERE quantity_name=@quantity_name'
      QuantityAlias = 'SELECT COUNT(*) FROM dbo.QuantityAlias WHERE raw_name=@quantity_name OR signature=@signature'
      QuotaBox = 'SELECT COUNT(*) FROM dbo.QuotaBox b WHERE b.box_id IN (@box_2024,@box_2020) OR EXISTS(SELECT 1 FROM dbo.QuotaBoxTarget t WHERE t.box_id=b.box_id AND t.target_code IN (@target_2024,@target_2020))'
      QuotaBoxTarget = 'SELECT COUNT(*) FROM dbo.QuotaBoxTarget WHERE box_id IN (@box_2024,@box_2020) OR target_code IN (@target_2024,@target_2020)'
      SignatureBoxMap = 'SELECT COUNT(*) FROM dbo.SignatureBoxMap WHERE signature=@signature AND box_id IN (@box_2024,@box_2020)'
      SignatureEntryMap = 'SELECT COUNT(*) FROM dbo.SignatureEntryMap WHERE signature=@signature AND target_code IN (@target_2024,@target_2020)'
      EngineeringTemplate = 'SELECT COUNT(*) FROM dbo.EngineeringTemplate WHERE box_id IN (@box_2024,@box_2020)'
    }
    $residueParameters = @{
      quantity_name=$quantityName; signature=$script:signature
      box_2024=$probe2024.box_id; box_2020=$probe2020.box_id
      target_2024=$probe2024.target_code; target_2020=$probe2020.target_code
    }
    foreach ($table in $residueSql.Keys) {
      $residue[$table] = Get-ScalarCount $verifyConnection $null $residueSql[$table] $residueParameters
      if ($residue[$table] -ne 0 -and $null -eq $failure) {
        $failure = New-Object System.InvalidOperationException("Rollback residue detected in $table")
      }
    }
    foreach ($table in $baselineCounts.Keys) {
      if (-not $afterCounts.ContainsKey($table)) { throw "Missing post-rollback table count: $table" }
      Add-Assertion ('global table count unchanged:' + $table) ([long]$afterCounts[$table]) ([long]$baselineCounts[$table])
    }
  }
  finally { $verifyConnection.Dispose() }
}
catch {
  if ($null -eq $failure) { $failure = $_ }
}

$result = [ordered]@{
  test = 'E3 SQL-only cross-partition rollback'
  status = if ($null -eq $failure -and $rollbackState -eq 'rolled-back') { 'passed' } else { 'failed' }
  started_at_utc = $runStarted.ToString('o')
  completed_at_utc = [DateTime]::UtcNow.ToString('o')
  target_database = 'RecoLearning'
  run_id = $runId
  host_process_count = 0
  dll_sha256 = (Get-FileHash -LiteralPath $dllPath -Algorithm SHA256).Hash
  rollback_state = $rollbackState
  failed_stage = if ($null -eq $failure) { '' } else { $stage }
  failure_type = if ($null -eq $failure) { '' } else { $failure.Exception.GetType().FullName }
  assertions = $assertions.ToArray()
  residue_counts = $residue
  baseline_table_counts = $baselineCounts
  after_table_counts = $afterCounts
}
$json = $result | ConvertTo-Json -Depth 8
[System.IO.File]::WriteAllText($evidencePath, $json, (New-Object System.Text.UTF8Encoding($false)))

if ($result.status -ne 'passed') {
  throw "E3 rollback acceptance failed at $stage; evidence: $evidencePath"
}
Write-Host "Test-LearningDbCrossPartitionRollback: PASS; evidence=$evidencePath"
