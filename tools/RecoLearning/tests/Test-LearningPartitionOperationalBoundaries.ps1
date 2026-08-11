$ErrorActionPreference = 'Stop'

function Assert-True([bool]$Condition, [string]$Message) {
  if (-not $Condition) { throw $Message }
}

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $testDir '..\..\..')).Path
$paths = @(
  (Join-Path $repoRoot 'tools\RecoLearning\Import-JsonlLibraries.ps1'),
  (Join-Path $repoRoot 'tools\RecoLearning\Repair-MappingBoxesIdentityConflicts.ps1'),
  (Join-Path $repoRoot 'tools\RecoLearning\Migrate-LearningPartitionSchema.ps1'),
  (Join-Path $repoRoot 'tools\RecoLearning\Invoke-PartitionFileCutover.ps1'),
  (Join-Path $repoRoot 'tools\RecoLearning\Export-PartitionMappingBoxes.ps1'),
  (Join-Path $repoRoot 'RecoQuotaRecommend\build.ps1'),
  (Join-Path $repoRoot 'tools\BuildColleaguePluginRelease.ps1')
)
foreach ($path in $paths) {
  $tokens = $null; $errors = $null
  [void][Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
  Assert-True ($errors.Count -eq 0) ("PowerShell 5 parse failure: " + $path + ' ' + (($errors | ForEach-Object { $_.Message }) -join '; '))
}

$import = [IO.File]::ReadAllText($paths[0], [Text.Encoding]::UTF8)
$rejectIndex = $import.IndexOf('if ($ImportBindingHistory)')
$hasPermanentReject = $import.Contains('throw "mapping-boxes/learning -> BindingLog') -and $import.Contains('software_partition/method_no')
$hasHistoryConnection = $import.Contains('$historyConnection')
$hasBindingLogTarget = $import.Contains("'dbo.BindingLog'")
Assert-True ($rejectIndex -ge 0 -and $hasPermanentReject -and -not $hasHistoryConnection -and -not $hasBindingLogTarget) ("Deprecated binding import was not permanently removed before SQL access: reject=$rejectIndex marker=$hasPermanentReject connection=$hasHistoryConnection target=$hasBindingLogTarget")

$migration = [IO.File]::ReadAllText($paths[2], [Text.Encoding]::UTF8)
foreach ($marker in @("restricted to the exact database RecoLearning",'Assert-TransactionTarget','BACKUP DATABASE [RecoLearning]','RESTORE VERIFYONLY','finalize-partition-schema.sql','RecoLearningPartitionRunId')) {
  Assert-True $migration.Contains($marker) ("Migration safety marker missing: " + $marker)
}
$cutover = [IO.File]::ReadAllText($paths[3], [Text.Encoding]::UTF8)
foreach ($marker in @('File cutover is permanently restricted','mapping-isolation.json','bin-rollback','Restore-TouchedFiles')) {
  Assert-True $cutover.Contains($marker) ("File cutover safety marker missing: " + $marker)
}
$export = [IO.File]::ReadAllText($paths[4], [Text.Encoding]::UTF8)
foreach ($marker in @('Mapping export is permanently restricted','material-index.jsonl','missing_or_ambiguous_metadata','mapping-export-report.json')) {
  Assert-True $export.Contains($marker) ("Mapping export safety marker missing: " + $marker)
}
$build = [IO.File]::ReadAllText($paths[5], [Text.Encoding]::UTF8)
Assert-True ($build.Contains('mapping-boxes.jsonl') -and $build.Contains('Where-Object { -not [string]::Equals')) 'Build data copy does not explicitly exclude mapping-boxes.'
$release = [IO.File]::ReadAllText($paths[6], [Text.Encoding]::UTF8)
Assert-True ($release.Contains('.mapping-boxes.empty') -and $release.Contains('Copy-RequiredFile -Source $seedSource')) 'Release seed is not forced through an empty mapping-boxes source.'
$repair = [IO.File]::ReadAllText($paths[1], [Text.Encoding]::UTF8)
foreach ($marker in @('Assert-SoftwareStopped','ExpectedSha256','pre-identity-repair-','mapping-boxes.lock','Post-repair analysis')) {
  Assert-True $repair.Contains($marker) ("Repair safety marker missing: " + $marker)
}
Write-Host 'PASS B11/B12/B15/B17 operational script boundaries and Windows PowerShell 5 parsing'
