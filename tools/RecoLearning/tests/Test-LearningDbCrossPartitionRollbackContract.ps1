$ErrorActionPreference = 'Stop'

$scriptPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'Test-LearningDbCrossPartitionRollback.ps1'
if (-not (Test-Path -LiteralPath $scriptPath)) {
  throw "Missing E3 rollback acceptance script: $scriptPath"
}

$source = [System.IO.File]::ReadAllText($scriptPath, [System.Text.Encoding]::UTF8)

function Assert-Contains([string]$Text, [string]$Expected, [string]$Message) {
  if ($Text.IndexOf($Expected, [System.StringComparison]::Ordinal) -lt 0) { throw $Message }
}

function Assert-NotContains([string]$Text, [string]$Forbidden, [string]$Message) {
  if ($Text.IndexOf($Forbidden, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { throw $Message }
}

Assert-Contains $source '[string]$TargetDatabase' 'The live test must require an explicit target database.'
Assert-Contains $source '[switch]$ExecuteLive' 'The live test must require an explicit live-execution switch.'
Assert-Contains $source "'RecoLearning'" 'The live test must pin the exact learning database.'
Assert-Contains $source "'RejjNet2020','ReJJGSNet2024','ReJJQDNet2024'" 'The live test must gate all supported hosts.'
Assert-Contains $source "DB_NAME()" 'The live test must verify the connected database name.'
Assert-Contains $source 'IsolationLevel]::Serializable' 'The live test must use a serializable transaction.'
Assert-Contains $source "GetMethod('UpsertBindingGroupAggregates'" 'The live test must exercise the production aggregate writer.'
Assert-Contains $source 'm.software_partition=@software_partition' 'The relation assertion must use the SmartFill partition predicate.'
Assert-Contains $source 'software_partition=@software_partition AND method_no=@method_no' 'Entry assertions must use partition and method number.'
Assert-Contains $source 'same-method cross-partition' 'Entry and scope assertions must change only the partition before testing method isolation.'
Assert-Contains $source 'global table count unchanged:' 'The live test must compare every baseline table count after rollback.'
Assert-Contains $source '$transaction.Rollback()' 'The live test must unconditionally roll back its transaction.'
Assert-Contains $source 'New-Object System.Data.SqlClient.SqlConnection' 'The live test must verify residue through a fresh connection.'

foreach ($table in 'BindingLog','QuantityAlias','QuotaBox','QuotaBoxTarget','SignatureBoxMap','SignatureEntryMap','EngineeringTemplate') {
  Assert-Contains $source $table "The residue gate is missing table $table."
}

foreach ($forbidden in 'Commit(','CREATE TABLE','ALTER TABLE','DROP TABLE','TRUNCATE TABLE','DELETE FROM','mapping-boxes','learning.jsonl','outbox','dead-letter','pending overlay') {
  Assert-NotContains $source $forbidden "The E3 rollback script contains forbidden text: $forbidden"
}
Assert-NotContains $source 'INSERT INTO dbo.BindingLog' 'The rollback test must not consume the non-transactional BindingLog identity seed.'

Write-Host 'Test-LearningDbCrossPartitionRollbackContract: PASS'
