$ErrorActionPreference = 'Stop'

function Assert-True([bool]$Condition, [string]$Message) {
  if (-not $Condition) { throw $Message }
}

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $testDir '..\..\..')).Path
$migrationPath = Join-Path $repoRoot 'tools\RecoLearning\Migrate-LearningPartitionSchema.ps1'
$schemaPath = Join-Path $repoRoot 'tools\RecoLearning\schema.sql'
$finalizePath = Join-Path $repoRoot 'tools\RecoLearning\finalize-partition-schema.sql'
$migration = [IO.File]::ReadAllText($migrationPath)
$schema = [IO.File]::ReadAllText($schemaPath)
$finalize = [IO.File]::ReadAllText($finalizePath)

Assert-True (-not $schema.Contains('TRUNCATE TABLE')) 'DDL-1 schema contains TRUNCATE.'
Assert-True (-not $schema.Contains('DROP CONSTRAINT')) 'DDL-1 schema switches an existing primary key.'
foreach ($table in @('QuantityFormulaOperand','QuantityFormulaRule','SignatureEntryMap','EngineeringTemplate','SignatureBoxMap','QuotaBoxTarget','QuotaBox','QuantityAlias','SheetTemplateRow')) {
  Assert-True $finalize.Contains('TRUNCATE TABLE dbo.' + $table) ("DDL-2 missing TRUNCATE: " + $table)
}
Assert-True ($finalize.Contains('IF @@TRANCOUNT = 0') -and $finalize.Contains("RAISERROR('DDL-2 requires") -and $finalize.Contains('QUOTENAME(@constraint_name)')) 'DDL-2 transaction/dynamic constraint guard missing.'
Assert-True (-not $migration.Contains('CONCAT(') -and -not $finalize.Contains('THROW ')) 'Migration uses SQL syntax newer than the supported SQL Server boundary.'
Assert-True ($migration.Contains('[ValidateNotNullOrEmpty()][string]$TargetDatabase') -and -not $migration.Contains("[string]`$TargetDatabase =")) 'TargetDatabase is optional or defaulted.'
Assert-True ($migration.Contains('Get-RecoConnectionString -Database $TargetDatabase') -and -not $migration.Contains('Invoke-RecoNonQuery') -and -not $migration.Contains('Invoke-RecoQuery')) 'Migration has an unscoped database helper.'
$processGateAt = $migration.LastIndexOf('Assert-SoftwareStopped', [StringComparison]::Ordinal)
$stateCreateAt = $migration.LastIndexOf('[void][IO.Directory]::CreateDirectory($stateDirectory)', [StringComparison]::Ordinal)
$credentialReadAt = $migration.LastIndexOf('Get-RecoConnectionString -Database $TargetDatabase', [StringComparison]::Ordinal)
Assert-True ($processGateAt -ge 0 -and $stateCreateAt -gt $processGateAt -and $credentialReadAt -gt $processGateAt) 'Credential/state initialization occurs before the live process gate.'
Assert-True ($migration.Contains("if (`$targets.Count -ne 4)") -and $migration.Contains('Duplicate deployment target path:') -and
  $migration.Contains('Deployment matrix is incomplete:') -and $migration.Contains("if (`$mappingRows.Count -ne 2)")) 'Deployment evidence is not bound to the complete four-target/two-mapping matrix.'
Assert-True ($migration.Contains('[IO.File]::Replace($temp, $path, $previous, $true)') -and
  -not $migration.Contains('[IO.File]::Replace($temp, $path, $null, $true)')) 'Existing migration state still uses an invalid empty File.Replace backup path.'
Assert-True ($migration.Contains('Existing backup predates this migration run Prepare state.') -and
  $migration.Contains('ExistingBackupReused=$backupExists') -and
  $migration.Contains('Assert-RowCountsEqual $state.prepare_row_counts $beforeCounts')) 'RecordBackup cannot safely resume after a post-backup state-write failure.'
foreach ($mode in @('Prepare','RecordBackup','DeploymentPreflight','Backfill','Finalize','Abort')) {
  Assert-True $migration.Contains("ParameterSetName = '$mode'") ("Migration mode missing: " + $mode)
}

$atomicRoot = Join-Path $repoRoot ('obj\partition-state-atomic-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($atomicRoot)
try {
  $functionStart = $migration.IndexOf('function Write-StateAtomic', [StringComparison]::Ordinal)
  $functionEnd = $migration.IndexOf('function Assert-State', $functionStart, [StringComparison]::Ordinal)
  Assert-True ($functionStart -ge 0 -and $functionEnd -gt $functionStart) 'Could not extract Write-StateAtomic for the PowerShell 5 runtime test.'
  $stateDirectory = $atomicRoot
  function Get-StatePath([string]$Id) { return Join-Path $stateDirectory ('partition-' + $Id.ToLowerInvariant() + '.json') }
  Invoke-Expression $migration.Substring($functionStart, $functionEnd - $functionStart)
  $testRun = [Guid]::NewGuid().ToString('N')
  $testState = [pscustomobject]@{ run_id=$testRun; state='prepared' }
  $testPath = Write-StateAtomic $testState
  $testState.state = 'backed_up'
  [void](Write-StateAtomic $testState)
  $persisted = [IO.File]::ReadAllText($testPath, [Text.Encoding]::UTF8) | ConvertFrom-Json
  Assert-True ([string]$persisted.state -eq 'backed_up') 'PowerShell 5 could not atomically replace an existing migration state file.'
  Assert-True (@(Get-ChildItem -LiteralPath $atomicRoot -File | Where-Object { $_.Name -like '*.tmp' -or $_.Name -like '*.previous' }).Count -eq 0) 'Atomic migration state update left a temporary file.'
} finally {
  if ([IO.Directory]::Exists($atomicRoot)) { [IO.Directory]::Delete($atomicRoot, $true) }
}

$startInfo = New-Object Diagnostics.ProcessStartInfo
$startInfo.FileName = 'powershell.exe'
$startInfo.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $migrationPath + '" -Prepare -TargetDatabase RecoData2020'
$startInfo.UseShellExecute = $false; $startInfo.CreateNoWindow = $true
$startInfo.RedirectStandardOutput = $true; $startInfo.RedirectStandardError = $true
$process = New-Object Diagnostics.Process; $process.StartInfo = $startInfo
try {
  Assert-True $process.Start() 'Could not start invalid-target migration probe.'
  $stdout = $process.StandardOutput.ReadToEnd(); $stderr = $process.StandardError.ReadToEnd(); $process.WaitForExit()
  Assert-True ($process.ExitCode -ne 0) 'Migration accepted a business database target.'
  Assert-True (($stdout + $stderr).Contains('restricted to the exact database RecoLearning')) 'Invalid-target rejection was not explicit.'
} finally { $process.Dispose() }

Assert-True ([IO.File]::ReadAllText((Join-Path $repoRoot '.gitignore')).Contains('/tools/RecoLearning/migration-state/')) 'Migration state directory is not ignored.'
Write-Host 'PASS B1a/B1b/B15 database target, DDL split, state, and transaction boundaries (no SQL connection opened)'
