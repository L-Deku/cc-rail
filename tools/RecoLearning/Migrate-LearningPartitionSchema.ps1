[CmdletBinding(DefaultParameterSetName = 'Prepare')]
param(
  [Parameter(Mandatory = $true, ParameterSetName = 'Prepare')][switch]$Prepare,
  [Parameter(Mandatory = $true, ParameterSetName = 'RecordBackup')][switch]$RecordBackup,
  [Parameter(Mandatory = $true, ParameterSetName = 'DeploymentPreflight')][switch]$RecordDeploymentPreflight,
  [Parameter(Mandatory = $true, ParameterSetName = 'Backfill')][switch]$Backfill,
  [Parameter(Mandatory = $true, ParameterSetName = 'Finalize')][switch]$Finalize,
  [Parameter(Mandatory = $true, ParameterSetName = 'Abort')][switch]$Abort,

  [Parameter(Mandatory = $true, ParameterSetName = 'Prepare')]
  [Parameter(Mandatory = $true, ParameterSetName = 'RecordBackup')]
  [Parameter(Mandatory = $true, ParameterSetName = 'DeploymentPreflight')]
  [Parameter(Mandatory = $true, ParameterSetName = 'Backfill')]
  [Parameter(Mandatory = $true, ParameterSetName = 'Finalize')]
  [Parameter(Mandatory = $true, ParameterSetName = 'Abort')]
  [ValidateNotNullOrEmpty()][string]$TargetDatabase,

  [Parameter(Mandatory = $true, ParameterSetName = 'RecordBackup')]
  [Parameter(Mandatory = $true, ParameterSetName = 'DeploymentPreflight')]
  [Parameter(Mandatory = $true, ParameterSetName = 'Backfill')]
  [Parameter(Mandatory = $true, ParameterSetName = 'Finalize')]
  [Parameter(Mandatory = $true, ParameterSetName = 'Abort')]
  [ValidatePattern('^[a-fA-F0-9]{32}$')][string]$RunId,

  [Parameter(Mandatory = $true, ParameterSetName = 'RecordBackup')]
  [ValidateNotNullOrEmpty()][string]$BackupPath,

  [Parameter(Mandatory = $true, ParameterSetName = 'DeploymentPreflight')]
  [ValidateNotNullOrEmpty()][string]$ArtifactManifest,
  [Parameter(Mandatory = $true, ParameterSetName = 'DeploymentPreflight')]
  [ValidateNotNullOrEmpty()][string]$DeploymentEvidence,

  [Parameter(Mandatory = $true, ParameterSetName = 'Backfill')]
  [ValidateNotNullOrEmpty()][string]$QuotaIndex2020,
  [Parameter(Mandatory = $true, ParameterSetName = 'Backfill')]
  [ValidateNotNullOrEmpty()][string]$QuotaIndex2024
)

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Common.ps1"

if (-not [string]::Equals($TargetDatabase, 'RecoLearning', [StringComparison]::Ordinal)) {
  throw 'This migration is permanently restricted to the exact database RecoLearning.'
}

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
$stateDirectory = Join-Path $PSScriptRoot 'migration-state'
$targetConnectionString = $null
$dbIdentitySha256 = $null

function New-TargetConnection {
  return New-Object Data.SqlClient.SqlConnection $targetConnectionString
}

function Assert-SoftwareStopped {
  $running = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -in @('RejjNet2020','ReJJGSNet2024','ReJJQDNet2024') })
  if ($running.Count -ne 0) {
    throw ('Migration requires all target software processes stopped: ' + (($running | ForEach-Object { $_.ProcessName + ':' + $_.Id }) -join ','))
  }
}

function Assert-OutboxEmpty {
  $pending = New-Object System.Collections.Generic.List[string]
  foreach ($file in @(Get-ChildItem -LiteralPath $repoRoot -Recurse -File -Filter 'learning-db-outbox.jsonl' -ErrorAction SilentlyContinue)) {
    $hasPending = $false
    foreach ($line in [IO.File]::ReadLines($file.FullName, [Text.Encoding]::UTF8)) {
      if (-not [string]::IsNullOrWhiteSpace($line)) { $hasPending = $true; break }
    }
    if ($hasPending) {
      [void]$pending.Add($file.FullName)
    }
  }
  if ($pending.Count -ne 0) { throw ('Learning outbox is not empty: ' + ($pending -join ';')) }
}

function Get-StatePath([string]$Id) {
  return Join-Path $stateDirectory ('partition-' + $Id.ToLowerInvariant() + '.json')
}

function Read-State([string]$Id) {
  $path = Get-StatePath $Id
  if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Migration state not found: $path" }
  return Get-Content -LiteralPath $path -Raw -Encoding UTF8 | ConvertFrom-Json
}

function Write-StateAtomic($State) {
  $path = Get-StatePath ([string]$State.run_id)
  $temp = $path + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
  $previous = $path + '.' + [Guid]::NewGuid().ToString('N') + '.previous'
  try {
    [IO.File]::WriteAllText($temp, ($State | ConvertTo-Json -Depth 12), (New-Object Text.UTF8Encoding($true)))
    if ([IO.File]::Exists($path)) { [IO.File]::Replace($temp, $path, $previous, $true) } else { [IO.File]::Move($temp, $path) }
  } finally {
    if ([IO.File]::Exists($temp)) { [IO.File]::Delete($temp) }
    if ([IO.File]::Exists($previous)) {
      try { [IO.File]::Delete($previous) }
      catch { Write-Warning 'Migration state was committed, but its temporary previous copy could not be removed.' }
    }
  }
  return $path
}

function Assert-State($State, [string]$Expected) {
  if (-not [string]::Equals([string]$State.target_database, $TargetDatabase, [StringComparison]::Ordinal) -or
      -not [string]::Equals([string]$State.db_identity_sha256, $dbIdentitySha256, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Migration state database identity mismatch.'
  }
  if (-not [string]::Equals([string]$State.state, $Expected, [StringComparison]::Ordinal)) {
    throw "Migration state must be '$Expected', actual '$($State.state)'."
  }
}

function Invoke-Scalar($Connection, $Transaction, [string]$Sql, [hashtable]$Parameters = @{}) {
  $command = $Connection.CreateCommand(); $command.Transaction = $Transaction; $command.CommandText = $Sql; $command.CommandTimeout = 600
  foreach ($key in $Parameters.Keys) { [void]$command.Parameters.AddWithValue('@' + $key, $Parameters[$key]) }
  try { return $command.ExecuteScalar() } finally { $command.Dispose() }
}

function Invoke-NonQuery($Connection, $Transaction, [string]$Sql, [hashtable]$Parameters = @{}) {
  $command = $Connection.CreateCommand(); $command.Transaction = $Transaction; $command.CommandText = $Sql; $command.CommandTimeout = 3600
  foreach ($key in $Parameters.Keys) { [void]$command.Parameters.AddWithValue('@' + $key, $Parameters[$key]) }
  try { return $command.ExecuteNonQuery() } finally { $command.Dispose() }
}

function Invoke-Table($Connection, $Transaction, [string]$Sql, [hashtable]$Parameters = @{}) {
  $command = $Connection.CreateCommand(); $command.Transaction = $Transaction; $command.CommandText = $Sql; $command.CommandTimeout = 600
  foreach ($key in $Parameters.Keys) { [void]$command.Parameters.AddWithValue('@' + $key, $Parameters[$key]) }
  $table = New-Object Data.DataTable
  $adapter = New-Object Data.SqlClient.SqlDataAdapter $command
  try { [void]$adapter.Fill($table); return ,$table } finally { $adapter.Dispose(); $command.Dispose() }
}

function Assert-TransactionTarget($Connection, $Transaction, [string]$ExpectedRunState = '') {
  $actual = [string](Invoke-Scalar $Connection $Transaction 'SELECT DB_NAME();')
  if (-not [string]::Equals($actual, $TargetDatabase, [StringComparison]::Ordinal)) { throw "Transaction target mismatch: $actual" }
  if ($ExpectedRunState.Length -gt 0) {
    $state = Read-State $RunId
    Assert-State $state $ExpectedRunState
  }
}

function Get-StructureFingerprint($Connection) {
  $sql = @'
SELECT CONVERT(NVARCHAR(128),s.name)+N'.'+CONVERT(NVARCHAR(128),o.name)+N'|'+CONVERT(NVARCHAR(20),c.column_id)+N'|'+
  CONVERT(NVARCHAR(128),c.name)+N'|'+CONVERT(NVARCHAR(128),t.name)+N'|'+CONVERT(NVARCHAR(20),c.max_length)+N'|'+
  CONVERT(NVARCHAR(5),c.is_nullable)+N'|'+ISNULL(CONVERT(NVARCHAR(128),i.name),N'')+N'|'+
  CONVERT(NVARCHAR(5),ISNULL(i.is_primary_key,0))+N'|'+CONVERT(NVARCHAR(20),ISNULL(ic.key_ordinal,0)) AS schema_line
FROM sys.objects o
JOIN sys.schemas s ON s.schema_id=o.schema_id
JOIN sys.columns c ON c.object_id=o.object_id
JOIN sys.types t ON t.user_type_id=c.user_type_id
LEFT JOIN sys.index_columns ic ON ic.object_id=o.object_id AND ic.column_id=c.column_id
LEFT JOIN sys.indexes i ON i.object_id=ic.object_id AND i.index_id=ic.index_id
WHERE o.type='U' AND s.name='dbo'
ORDER BY s.name,o.name,c.column_id,i.index_id,ic.key_ordinal;
'@
  $table = Invoke-Table $Connection $null $sql
  $text = [string]::Join("`n", @($table.Rows | ForEach-Object { [string]$_.schema_line }))
  $hash = [Security.Cryptography.SHA256]::Create()
  try { return ([BitConverter]::ToString($hash.ComputeHash([Text.Encoding]::UTF8.GetBytes($text))).Replace('-', '')) }
  finally { $hash.Dispose() }
}

function Get-RowCounts($Connection, $Transaction = $null) {
  $tables = @('BindingLog','QuantityAlias','QuotaBox','QuotaBoxTarget','SignatureBoxMap','QuantityFormulaRule','QuantityFormulaOperand','SignatureEntryMap','EngineeringTemplate','SheetTemplateRow')
  $result = [ordered]@{}
  foreach ($table in $tables) {
    $result[$table] = [long](Invoke-Scalar $Connection $Transaction ('SELECT COUNT_BIG(*) FROM dbo.' + $table + ';'))
  }
  return $result
}

function Assert-RowCountsEqual($Expected, $Actual) {
  foreach ($property in $Expected.PSObject.Properties) {
    if ([long]$property.Value -ne [long]$Actual[$property.Name]) { throw "Row count changed for $($property.Name)." }
  }
}

function Get-BindingLogInvariant($Connection, $Transaction) {
  return [string](Invoke-Scalar $Connection $Transaction @'
SELECT CONVERT(NVARCHAR(30),COUNT_BIG(*))+N'|'+ISNULL(MIN(event_hash),'')+N'|'+ISNULL(MAX(event_hash),'')+N'|'+
  CONVERT(NVARCHAR(30),ISNULL(CHECKSUM_AGG(BINARY_CHECKSUM(id,occurred_at,imported_at,source,method,project_id,entry_code,entry_name,
    quantity_name,quantity_unit,target_kind,target_code,target_name,target_unit,group_key,event_hash,
    CONVERT(NVARCHAR(4000),extra))),0))
FROM dbo.BindingLog;
'@)
}

function Read-ArtifactManifest([string]$ManifestPath) {
  $resolved = (Resolve-Path -LiteralPath $ManifestPath -ErrorAction Stop).Path
  $workspaceBoundary = $repoRoot.TrimEnd('\') + '\'
  if (-not $resolved.StartsWith($workspaceBoundary, [StringComparison]::OrdinalIgnoreCase)) { throw 'Artifact manifest must stay inside the workspace.' }
  $manifest = Get-Content -LiteralPath $resolved -Raw -Encoding UTF8 | ConvertFrom-Json
  $directory = Split-Path -Parent $resolved
  $actualFiles = @(Get-ChildItem -LiteralPath $directory -File -Force | Select-Object -ExpandProperty Name | Sort-Object)
  $allowed = @('RecoExpandPanel.dll','RecoQuotaRecommend.dll','artifact-manifest.json') | Sort-Object
  if ([string]::Join('|',$actualFiles) -ne [string]::Join('|',$allowed)) { throw 'Artifact directory whitelist mismatch.' }
  if (@($manifest.files).Count -ne 2) { throw 'Artifact manifest must bind exactly two DLLs.' }
  foreach ($file in @($manifest.files)) {
    if ([string]$file.name -notin @('RecoExpandPanel.dll','RecoQuotaRecommend.dll')) { throw 'Artifact manifest contains an unapproved file.' }
    $filePath = Join-Path $directory ([string]$file.name)
    if ((Get-FileHash -LiteralPath $filePath -Algorithm SHA256).Hash -ne [string]$file.sha256) { throw "Artifact hash mismatch: $($file.name)" }
  }
  return [pscustomobject]@{ Path=$resolved; Sha256=(Get-FileHash -LiteralPath $resolved -Algorithm SHA256).Hash; Manifest=$manifest }
}

function Get-ApprovedRuntimeDirectories {
  return @(
    [pscustomobject]@{ runtime_id='2020'; path=(Resolve-Path -LiteralPath (Join-Path $repoRoot '铁路基本建设工程投资控制系统2020网络版V0503021201') -ErrorAction Stop).Path },
    [pscustomobject]@{ runtime_id='2024'; path=(Resolve-Path -LiteralPath (Join-Path $repoRoot '2024铁路工程云计价系统网络版V1.0\铁路工程云计价系统网络版V1.0') -ErrorAction Stop).Path }
  )
}

function Read-DeploymentEvidence([string]$EvidencePath, $Artifact) {
  $resolved = (Resolve-Path -LiteralPath $EvidencePath -ErrorAction Stop).Path
  $workspaceBoundary = $repoRoot.TrimEnd('\') + '\'
  if (-not $resolved.StartsWith($workspaceBoundary, [StringComparison]::OrdinalIgnoreCase)) { throw 'Deployment evidence must stay inside the workspace.' }
  $evidence = Get-Content -LiteralPath $resolved -Raw -Encoding UTF8 | ConvertFrom-Json
  if ([string]$evidence.run_id -ne $RunId -or [string]$evidence.target_database -ne $TargetDatabase -or
      [string]$evidence.artifact_manifest_sha256 -ne $Artifact.Sha256 -or
      [string]$evidence.artifact_set_id -ne [string]$Artifact.Manifest.artifact_set_id) { throw 'Deployment evidence identity mismatch.' }

  $approved = @(Get-ApprovedRuntimeDirectories)
  if ($approved.Count -ne 2) { throw 'Approved runtime directory count must be exactly two.' }
  $approvedById = @{}
  foreach ($runtime in $approved) {
    if ($approvedById.ContainsKey([string]$runtime.runtime_id)) { throw 'Approved runtime id is duplicated.' }
    $approvedById[[string]$runtime.runtime_id] = $runtime
  }
  $runtimeRows = @($evidence.runtime_directories)
  if ($runtimeRows.Count -ne 2) { throw 'Deployment evidence must bind exactly two runtime directories.' }
  $seenRuntimePaths = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  foreach ($runtime in $runtimeRows) {
    $runtimeId = [string]$runtime.runtime_id
    if (-not $approvedById.ContainsKey($runtimeId)) { throw "Unapproved runtime id: $runtimeId" }
    $runtimePath = (Resolve-Path -LiteralPath ([string]$runtime.path) -ErrorAction Stop).Path
    if (-not [string]::Equals($runtimePath, [string]$approvedById[$runtimeId].path, [StringComparison]::OrdinalIgnoreCase)) { throw "Runtime directory mismatch: $runtimeId" }
    if (-not $seenRuntimePaths.Add($runtimePath)) { throw "Duplicate runtime directory: $runtimePath" }
    if (-not $runtime.sentinel_verified) { throw "Runtime sentinel was not verified: $runtimePath" }
  }

  $targets = @($evidence.targets)
  if ($targets.Count -ne 4) { throw 'Deployment evidence must bind exactly four DLL targets.' }
  $seenTargets = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  $seenMatrix = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  $evidenceBoundary = (Split-Path -Parent $resolved).TrimEnd('\') + '\'
  foreach ($target in $targets) {
    $runtimeId = [string]$target.runtime_id
    if (-not $approvedById.ContainsKey($runtimeId)) { throw "Unapproved deployment runtime: $runtimeId" }
    $targetPath = (Resolve-Path -LiteralPath ([string]$target.target_path) -ErrorAction Stop).Path
    $leaf = Split-Path -Leaf $targetPath
    if ($leaf -notin @('RecoExpandPanel.dll','RecoQuotaRecommend.dll')) { throw "Unapproved deployment target: $leaf" }
    if (-not [string]::Equals((Split-Path -Parent $targetPath), [string]$approvedById[$runtimeId].path, [StringComparison]::OrdinalIgnoreCase)) { throw "Deployment target is outside its approved runtime directory: $targetPath" }
    if (-not $seenTargets.Add($targetPath)) { throw "Duplicate deployment target path: $targetPath" }
    $matrixKey = $runtimeId + '|' + $leaf.ToLowerInvariant()
    if (-not $seenMatrix.Add($matrixKey)) { throw "Duplicate deployment matrix target: $matrixKey" }
    $artifactFile = @($Artifact.Manifest.files | Where-Object { [string]$_.name -eq $leaf })
    if ($artifactFile.Count -ne 1 -or [string]$target.new_sha256 -ne [string]$artifactFile[0].sha256) {
      throw "Deployment target hash is not bound to the artifact manifest: $leaf"
    }
    $oldHash = (Get-FileHash -LiteralPath $targetPath -Algorithm SHA256).Hash
    if ($oldHash -ne [string]$target.old_sha256) { throw "Deployment target changed after evidence capture: $targetPath" }
    if (-not $target.sentinel_verified) { throw "Target sentinel was not verified: $targetPath" }
    $rollback = (Resolve-Path -LiteralPath ([string]$target.rollback_path) -ErrorAction Stop).Path
    $rollbackHash = (Get-FileHash -LiteralPath $rollback -Algorithm SHA256).Hash
    if (-not $rollback.StartsWith($evidenceBoundary, [StringComparison]::OrdinalIgnoreCase) -or
        $rollbackHash -ne [string]$target.rollback_sha256 -or $rollbackHash -ne $oldHash) { throw 'Rollback DLL evidence mismatch.' }
  }
  foreach ($runtimeId in @('2020','2024')) {
    foreach ($leaf in @('RecoExpandPanel.dll','RecoQuotaRecommend.dll')) {
      if (-not $seenMatrix.Contains($runtimeId + '|' + $leaf)) { throw "Deployment matrix is incomplete: $runtimeId|$leaf" }
    }
  }

  $mappingRows = @($evidence.mapping_files)
  if ($mappingRows.Count -ne 2) { throw 'Deployment evidence must bind exactly two mapping files.' }
  $seenMapping = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  foreach ($mapping in $mappingRows) {
    $runtimeId = [string]$mapping.runtime_id
    if (-not $approvedById.ContainsKey($runtimeId)) { throw "Unapproved mapping runtime: $runtimeId" }
    if (-not $seenMapping.Add($runtimeId)) { throw "Duplicate mapping runtime: $runtimeId" }
    $expectedPath = [IO.Path]::GetFullPath((Join-Path ([string]$approvedById[$runtimeId].path) 'RecoQuotaData\mapping-boxes.jsonl'))
    $mappingPath = [IO.Path]::GetFullPath([string]$mapping.path)
    if (-not [string]::Equals($mappingPath, $expectedPath, [StringComparison]::OrdinalIgnoreCase)) { throw "Mapping path mismatch: $runtimeId" }
    if ($null -eq $mapping.exists -or $mapping.exists -isnot [bool]) { throw "Mapping existence flag is invalid: $runtimeId" }
    $actualExists = [IO.File]::Exists($expectedPath)
    if ([bool]$mapping.exists -ne $actualExists) { throw "Mapping existence changed after evidence capture: $runtimeId" }
    if ($actualExists) {
      $mappingHash = (Get-FileHash -LiteralPath $expectedPath -Algorithm SHA256).Hash
      if ($mappingHash -ne [string]$mapping.sha256 -or [long](Get-Item -LiteralPath $expectedPath).Length -ne [long]$mapping.bytes) { throw "Mapping evidence mismatch: $runtimeId" }
    } elseif (-not [string]::IsNullOrEmpty([string]$mapping.sha256) -or [long]$mapping.bytes -ne 0) {
      throw "Missing mapping evidence must have empty hash and zero bytes: $runtimeId"
    }
  }
  return [pscustomobject]@{ Path=$resolved; Sha256=(Get-FileHash -LiteralPath $resolved -Algorithm SHA256).Hash; Evidence=$evidence }
}

function Assert-BackupPath([string]$ServerPath, [string]$Id) {
  if ([string]::IsNullOrWhiteSpace($ServerPath) -or -not [IO.Path]::IsPathRooted($ServerPath) -or
      $ServerPath.IndexOf($Id, [StringComparison]::OrdinalIgnoreCase) -lt 0) {
    throw 'BackupPath must be an absolute SQL Server-side path containing the run_id.'
  }
}

function Get-BackupHeader($Connection, [string]$ServerPath) {
  $table = Invoke-Table $Connection $null 'RESTORE HEADERONLY FROM DISK=@backup_path;' @{ backup_path=$ServerPath }
  if ($table.Rows.Count -ne 1) { throw 'Backup HEADERONLY must return exactly one backup set.' }
  $row = $table.Rows[0]
  if ([string]$row.DatabaseName -ne $TargetDatabase) { throw 'Backup database name mismatch.' }
  return [ordered]@{
    database_name=[string]$row.DatabaseName; position=[int]$row.Position; backup_finish_date=([datetime]$row.BackupFinishDate).ToUniversalTime().ToString('o')
    backup_size=[long]$row.BackupSize; compressed_backup_size=[long]$row.CompressedBackupSize
    first_lsn=[string]$row.FirstLSN; last_lsn=[string]$row.LastLSN
  }
}

function Assert-BackupVerified($Connection, $State) {
  [void](Invoke-NonQuery $Connection $null 'RESTORE VERIFYONLY FROM DISK=@backup_path WITH CHECKSUM;' @{ backup_path=[string]$State.backup_path })
  $header = Get-BackupHeader $Connection ([string]$State.backup_path)
  foreach ($name in @('database_name','position','backup_finish_date','backup_size','compressed_backup_size','first_lsn','last_lsn')) {
    if ([string]$header[$name] -ne [string]$State.backup_header.$name) { throw "Backup header changed: $name" }
  }
}

function Get-BaseQuotaCode([string]$Code) {
  $value = ([string]$Code).Trim().ToUpperInvariant()
  $positions = @($value.IndexOf('*'), $value.IndexOf('/')) | Where-Object { $_ -ge 0 }
  if ($positions.Count -gt 0) { return $value.Substring(0, ($positions | Measure-Object -Minimum).Minimum) }
  return $value
}

function Read-QuotaCodes([string]$IndexPath) {
  $resolved = (Resolve-Path -LiteralPath $IndexPath -ErrorAction Stop).Path
  $codes = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  foreach ($line in [IO.File]::ReadLines($resolved, [Text.Encoding]::UTF8)) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $row = $line | ConvertFrom-Json
    $code = Get-BaseQuotaCode ([string]$row.quota_code)
    if ($code.Length -gt 0) { [void]$codes.Add($code) }
  }
  Write-Output -NoEnumerate $codes
}

function Add-ExclusiveCodesToTemp($Connection, $Transaction, [string]$Index2020, [string]$Index2024) {
  $codes2020 = Read-QuotaCodes $Index2020
  $codes2024 = Read-QuotaCodes $Index2024
  [void](Invoke-NonQuery $Connection $Transaction 'CREATE TABLE #ExclusiveQuotaCode(code NVARCHAR(100) NOT NULL, only_partition NVARCHAR(10) NOT NULL, PRIMARY KEY(code,only_partition));')
  $table = New-Object Data.DataTable
  [void]$table.Columns.Add('code',[string]); [void]$table.Columns.Add('only_partition',[string])
  foreach ($code in $codes2020) { if (-not $codes2024.Contains($code)) { [void]$table.Rows.Add($code,'2020') } }
  foreach ($code in $codes2024) { if (-not $codes2020.Contains($code)) { [void]$table.Rows.Add($code,'2024') } }
  $bulk = New-Object Data.SqlClient.SqlBulkCopy($Connection, [Data.SqlClient.SqlBulkCopyOptions]::Default, $Transaction)
  try {
    $bulk.DestinationTableName='#ExclusiveQuotaCode'; [void]$bulk.ColumnMappings.Add('code','code'); [void]$bulk.ColumnMappings.Add('only_partition','only_partition'); $bulk.WriteToServer($table)
  } finally { $bulk.Close() }
}

Assert-SoftwareStopped

[void][IO.Directory]::CreateDirectory($stateDirectory)
$targetConnectionString = Get-RecoConnectionString -Database $TargetDatabase
$connectionBuilder = New-Object Data.SqlClient.SqlConnectionStringBuilder $targetConnectionString
$canonicalDataSource = $connectionBuilder.DataSource.Trim().ToLowerInvariant()
$canonicalCatalog = $connectionBuilder.InitialCatalog.Trim()
$identityInput = $canonicalDataSource + '|' + $canonicalCatalog
$sha = [Security.Cryptography.SHA256]::Create()
try { $dbIdentitySha256 = ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($identityInput))).Replace('-', '')) }
finally { $sha.Dispose() }

if ($Prepare) {
  Assert-OutboxEmpty
  foreach ($file in @(Get-ChildItem -LiteralPath $stateDirectory -File -Filter 'partition-*.json' -ErrorAction SilentlyContinue)) {
    $old = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8 | ConvertFrom-Json
    if ([string]$old.db_identity_sha256 -eq $dbIdentitySha256 -and [string]$old.state -in @('prepared','backed_up','deployment_ready','backfilled')) {
      throw "An active migration run already exists: $($old.run_id)"
    }
  }
  $connection = New-TargetConnection
  try {
    $connection.Open()
    $transaction = $connection.BeginTransaction()
    try {
      Assert-TransactionTarget $connection $transaction
      $schema = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'schema.sql'), [Text.Encoding]::UTF8)
      [void](Invoke-NonQuery $connection $transaction $schema)
      $transaction.Commit()
    } catch { $transaction.Rollback(); throw }
    $run = [Guid]::NewGuid().ToString('N')
    $state = [ordered]@{
      run_id=$run; state='prepared'; target_database=$TargetDatabase; db_identity_sha256=$dbIdentitySha256
      prepared_at=[DateTime]::UtcNow.ToString('o'); structure_fingerprint=(Get-StructureFingerprint $connection)
      prepare_row_counts=(Get-RowCounts $connection)
      backup_path=''; backup_header=$null; backup_row_counts=$null; backed_up_at=''
      artifact_manifest_path=''; artifact_manifest_sha256=''; deployment_evidence_path=''; deployment_evidence_sha256=''; deployment_ready_at=''
      backfilled_at=''; isolation_audit_path=''; binding_log_count=0; partition_distribution=@()
      consumed_at=''; aborted_at=''; final_structure_fingerprint=''
    }
    $statePath = Write-StateAtomic $state
    [pscustomobject]@{ RunId=$run; State='prepared'; StatePath=$statePath }
  } finally { $connection.Dispose() }
  return
}

$state = Read-State $RunId

if ($RecordBackup) {
  Assert-State $state 'prepared'; Assert-OutboxEmpty; Assert-BackupPath $BackupPath $RunId
  $connection = New-TargetConnection
  try {
    $connection.Open()
    if ((Get-StructureFingerprint $connection) -ne [string]$state.structure_fingerprint) { throw 'Structure fingerprint changed after Prepare.' }
    $beforeCounts = Get-RowCounts $connection
    Assert-RowCountsEqual $state.prepare_row_counts $beforeCounts
    $exists = Invoke-Table $connection $null 'EXEC master.dbo.xp_fileexist @backup_path;' @{ backup_path=$BackupPath }
    if ($exists.Rows.Count -eq 0) { throw 'SQL Server-side backup path probe is unavailable.' }
    $backupExists = [int]$exists.Rows[0][0] -ne 0
    if ($backupExists) {
      [void](Invoke-NonQuery $connection $null 'RESTORE VERIFYONLY FROM DISK=@backup_path WITH CHECKSUM;' @{ backup_path=$BackupPath })
      $header = Get-BackupHeader $connection $BackupPath
      $preparedAt = [DateTime]::Parse([string]$state.prepared_at, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()
      $backupFinishedAt = [DateTime]::Parse([string]$header.backup_finish_date, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind).ToUniversalTime()
      if ($backupFinishedAt -lt $preparedAt) { throw 'Existing backup predates this migration run Prepare state.' }
    } else {
      [void](Invoke-NonQuery $connection $null 'BACKUP DATABASE [RecoLearning] TO DISK=@backup_path WITH COPY_ONLY,CHECKSUM;' @{ backup_path=$BackupPath })
      [void](Invoke-NonQuery $connection $null 'RESTORE VERIFYONLY FROM DISK=@backup_path WITH CHECKSUM;' @{ backup_path=$BackupPath })
      $header = Get-BackupHeader $connection $BackupPath
    }
    $afterCounts = Get-RowCounts $connection
    foreach ($name in $beforeCounts.Keys) { if ([long]$beforeCounts[$name] -ne [long]$afterCounts[$name]) { throw "Row count changed during backup: $name" } }
    $state.state='backed_up'; $state.backup_path=$BackupPath; $state.backup_header=$header; $state.backup_row_counts=$beforeCounts
    $state.backed_up_at=[DateTime]::UtcNow.ToString('o'); $state.structure_fingerprint=(Get-StructureFingerprint $connection)
    $statePath = Write-StateAtomic $state
    [pscustomobject]@{ RunId=$RunId; State='backed_up'; StatePath=$statePath; BackupPath=$BackupPath; ExistingBackupReused=$backupExists }
  } finally { $connection.Dispose() }
  return
}

if ($RecordDeploymentPreflight) {
  Assert-State $state 'backed_up'; Assert-OutboxEmpty
  $artifact = Read-ArtifactManifest $ArtifactManifest
  $evidence = Read-DeploymentEvidence $DeploymentEvidence $artifact
  $connection = New-TargetConnection
  try {
    $connection.Open(); Assert-BackupVerified $connection $state
    if ((Get-StructureFingerprint $connection) -ne [string]$state.structure_fingerprint) { throw 'Structure fingerprint changed after backup.' }
  } finally { $connection.Dispose() }
  $state.state='deployment_ready'; $state.artifact_manifest_path=$artifact.Path; $state.artifact_manifest_sha256=$artifact.Sha256
  $state.deployment_evidence_path=$evidence.Path; $state.deployment_evidence_sha256=$evidence.Sha256; $state.deployment_ready_at=[DateTime]::UtcNow.ToString('o')
  $statePath = Write-StateAtomic $state
  [pscustomobject]@{ RunId=$RunId; State='deployment_ready'; StatePath=$statePath }
  return
}

if ($Backfill) {
  Assert-State $state 'deployment_ready'; Assert-OutboxEmpty
  $artifact = Read-ArtifactManifest ([string]$state.artifact_manifest_path)
  if ($artifact.Sha256 -ne [string]$state.artifact_manifest_sha256) { throw 'Artifact manifest changed after preflight.' }
  $evidence = Read-DeploymentEvidence ([string]$state.deployment_evidence_path) $artifact
  if ($evidence.Sha256 -ne [string]$state.deployment_evidence_sha256) { throw 'Deployment evidence changed after preflight.' }
  $connection = New-TargetConnection
  $audit = $null
  try {
    $connection.Open(); Assert-BackupVerified $connection $state
    $transaction = $connection.BeginTransaction([Data.IsolationLevel]::Serializable)
    try {
      Assert-TransactionTarget $connection $transaction 'deployment_ready'
      $beforeCount = [long](Invoke-Scalar $connection $transaction 'SELECT COUNT_BIG(*) FROM dbo.BindingLog;')
      $beforeInvariant = Get-BindingLogInvariant $connection $transaction
      Add-ExclusiveCodesToTemp $connection $transaction $QuotaIndex2020 $QuotaIndex2024
      $backfillSql = @'
CREATE TABLE #Inference(id BIGINT NOT NULL PRIMARY KEY, method_partition NVARCHAR(10) NOT NULL, project_partition NVARCHAR(10) NOT NULL, final_partition NVARCHAR(10) NOT NULL, reason NVARCHAR(100) NOT NULL);
INSERT #Inference(id,method_partition,project_partition,final_partition,reason)
SELECT id,
  CASE WHEN REPLACE(LOWER(method),' ','') LIKE '%101%' THEN '2020'
       WHEN REPLACE(LOWER(method),' ','') LIKE '%2024%' THEN '2024'
       WHEN REPLACE(LOWER(method),' ','') LIKE '%2020%' OR REPLACE(LOWER(method),' ','') LIKE '%30%' THEN '2020' ELSE '' END,
  CASE WHEN source='import:excel-links' AND project_id LIKE '%RecoData2024%' THEN '2024'
       WHEN source='import:excel-links' AND project_id LIKE '%RecoData2020%' THEN '2020' ELSE '' END,
  '', ''
FROM dbo.BindingLog;
UPDATE #Inference SET reason='method_project_conflict'
WHERE method_partition<>'' AND project_partition<>'' AND method_partition<>project_partition;
UPDATE #Inference SET final_partition=CASE WHEN method_partition<>'' THEN method_partition ELSE project_partition END
WHERE reason='';
;WITH group_partition AS (
  SELECT b.group_key, MIN(i.final_partition) AS pmin, MAX(i.final_partition) AS pmax
  FROM dbo.BindingLog b JOIN #Inference i ON i.id=b.id WHERE i.final_partition<>'' GROUP BY b.group_key
)
UPDATE i SET final_partition=g.pmin
FROM #Inference i JOIN dbo.BindingLog b ON b.id=i.id JOIN group_partition g ON g.group_key=b.group_key
WHERE i.final_partition='' AND i.reason='' AND g.pmin=g.pmax;
;WITH conflicting_group AS (
  SELECT b.group_key FROM dbo.BindingLog b JOIN #Inference i ON i.id=b.id
  WHERE i.final_partition<>'' GROUP BY b.group_key HAVING MIN(i.final_partition)<>MAX(i.final_partition)
)
UPDATE i SET final_partition='',reason='group_partition_conflict'
FROM #Inference i JOIN dbo.BindingLog b ON b.id=i.id JOIN conflicting_group g ON g.group_key=b.group_key;
UPDATE i SET final_partition='',reason='exclusive_quota_conflict'
FROM #Inference i JOIN dbo.BindingLog b ON b.id=i.id JOIN #ExclusiveQuotaCode q
  ON q.code=UPPER(CASE WHEN CHARINDEX('*',b.target_code)>0 THEN LEFT(b.target_code,CHARINDEX('*',b.target_code)-1)
                       WHEN CHARINDEX('/',b.target_code)>0 THEN LEFT(b.target_code,CHARINDEX('/',b.target_code)-1) ELSE b.target_code END)
WHERE i.final_partition<>'' AND i.final_partition<>q.only_partition;
UPDATE #Inference SET reason='unresolved' WHERE final_partition='' AND reason='';
UPDATE b SET software_partition=i.final_partition,
  method_no=CASE WHEN i.final_partition='2024' THEN N'TB 10801'+NCHAR(8212)+N'2024' ELSE N'' END
FROM dbo.BindingLog b JOIN #Inference i ON i.id=b.id;
SELECT b.id,b.quantity_name,b.target_code,b.target_name,b.target_unit,b.source,b.project_id,b.occurred_at,b.group_key,i.reason
FROM dbo.BindingLog b JOIN #Inference i ON i.id=b.id WHERE i.final_partition='' ORDER BY b.id;
'@
      $audit = Invoke-Table $connection $transaction $backfillSql
      $afterCount = [long](Invoke-Scalar $connection $transaction 'SELECT COUNT_BIG(*) FROM dbo.BindingLog;')
      $afterInvariant = Get-BindingLogInvariant $connection $transaction
      if ($afterCount -ne $beforeCount -or $afterInvariant -ne $beforeInvariant) { throw 'Backfill changed BindingLog row count or pre-existing columns.' }
      $transaction.Commit()
    } catch { $transaction.Rollback(); throw }
    $auditPath = Join-Path $stateDirectory ('partition-' + $RunId + '-isolated.csv')
    $audit | Export-Csv -LiteralPath $auditPath -NoTypeInformation -Encoding UTF8
    $distribution = Invoke-Table $connection $null "SELECT software_partition,method_no,COUNT_BIG(*) AS row_count FROM dbo.BindingLog GROUP BY software_partition,method_no ORDER BY software_partition,method_no;"
    $state.state='backfilled'; $state.backfilled_at=[DateTime]::UtcNow.ToString('o'); $state.isolation_audit_path=$auditPath
    $state.binding_log_count=[long](Invoke-Scalar $connection $null 'SELECT COUNT_BIG(*) FROM dbo.BindingLog;'); $state.partition_distribution=@($distribution | Select-Object software_partition,method_no,row_count)
    $statePath = Write-StateAtomic $state
    [pscustomobject]@{ RunId=$RunId; State='backfilled'; StatePath=$statePath; IsolationAudit=$auditPath; IsolatedRows=$audit.Rows.Count }
  } finally { $connection.Dispose() }
  return
}

if ($Finalize) {
  Assert-State $state 'backfilled'; Assert-OutboxEmpty
  $artifact = Read-ArtifactManifest ([string]$state.artifact_manifest_path)
  if ($artifact.Sha256 -ne [string]$state.artifact_manifest_sha256) { throw 'Artifact manifest changed after preflight.' }
  $connection = New-TargetConnection
  try {
    $connection.Open(); Assert-BackupVerified $connection $state
    $existingRun = [string](Invoke-Scalar $connection $null "SELECT CONVERT(NVARCHAR(128),value) FROM sys.extended_properties WHERE class=0 AND name='RecoLearningPartitionRunId';")
    if ($existingRun.Length -gt 0) {
      if ($existingRun -ne $RunId) { throw 'Database was finalized by a different migration run.' }
      $state.state='consumed'; $state.consumed_at=[DateTime]::UtcNow.ToString('o'); $statePath=Write-StateAtomic $state
      [pscustomobject]@{ RunId=$RunId; State='consumed'; StatePath=$statePath; RecoveredLocalState=$true }; return
    }
    $currentCounts = Get-RowCounts $connection
    Assert-RowCountsEqual $state.backup_row_counts $currentCounts
    $transaction = $connection.BeginTransaction([Data.IsolationLevel]::Serializable)
    try {
      Assert-TransactionTarget $connection $transaction 'backfilled'
      if ([long](Invoke-Scalar $connection $transaction 'SELECT COUNT_BIG(*) FROM dbo.BindingLog;') -ne [long]$state.binding_log_count) { throw 'BindingLog row count changed before Finalize.' }
      $ddl2 = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'finalize-partition-schema.sql'), [Text.Encoding]::UTF8)
      [void](Invoke-NonQuery $connection $transaction $ddl2)
      [void](Invoke-NonQuery $connection $transaction @'
IF EXISTS(SELECT 1 FROM sys.extended_properties WHERE class=0 AND name='RecoLearningSchemaVersion')
  EXEC sys.sp_updateextendedproperty @name=N'RecoLearningSchemaVersion',@value=N'partition-v1';
ELSE EXEC sys.sp_addextendedproperty @name=N'RecoLearningSchemaVersion',@value=N'partition-v1';
IF EXISTS(SELECT 1 FROM sys.extended_properties WHERE class=0 AND name='RecoLearningPartitionRunId')
  EXEC sys.sp_updateextendedproperty @name=N'RecoLearningPartitionRunId',@value=@run_id;
ELSE EXEC sys.sp_addextendedproperty @name=N'RecoLearningPartitionRunId',@value=@run_id;
'@ @{ run_id=$RunId })
      $transaction.Commit()
    } catch { $transaction.Rollback(); throw }
    $state.state='consumed'; $state.consumed_at=[DateTime]::UtcNow.ToString('o'); $state.final_structure_fingerprint=(Get-StructureFingerprint $connection)
    $statePath=Write-StateAtomic $state
    [pscustomobject]@{ RunId=$RunId; State='consumed'; StatePath=$statePath; RecoveredLocalState=$false }
  } finally { $connection.Dispose() }
  return
}

if ($Abort) {
  if ([string]$state.state -notin @('prepared','backed_up','deployment_ready','backfilled')) { throw "Run cannot be aborted from state '$($state.state)'." }
  if ([string]$state.db_identity_sha256 -ne $dbIdentitySha256 -or [string]$state.target_database -ne $TargetDatabase) { throw 'Abort database identity mismatch.' }
  $connection = New-TargetConnection
  try { $connection.Open(); if ([string](Invoke-Scalar $connection $null 'SELECT DB_NAME();') -ne $TargetDatabase) { throw 'Abort connection target mismatch.' } }
  finally { $connection.Dispose() }
  $state.state='aborted'; $state.aborted_at=[DateTime]::UtcNow.ToString('o'); $statePath=Write-StateAtomic $state
  [pscustomobject]@{ RunId=$RunId; State='aborted'; StatePath=$statePath }
  return
}
