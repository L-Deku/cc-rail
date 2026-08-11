$ErrorActionPreference = 'Stop'

function Assert-True([bool]$Condition, [string]$Message) {
  if (-not $Condition) { throw $Message }
}

function Write-Utf8Json([string]$Path, $Value) {
  [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 12), (New-Object Text.UTF8Encoding($true)))
}

function Assert-Rejected([scriptblock]$Action, [string]$Message) {
  $rejected = $false
  try { & $Action } catch { $rejected = $true }
  Assert-True $rejected $Message
}

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $testDir '..\..\..')).Path
$sourceGenerator = Join-Path $repoRoot 'tools\RecoLearning\New-PartitionDeploymentEvidence.ps1'
$sourceMigration = Join-Path $repoRoot 'tools\RecoLearning\Migrate-LearningPartitionSchema.ps1'
$fixtureRoot = Join-Path $repoRoot ('obj\partition-evidence-fixture-' + [Guid]::NewGuid().ToString('N'))
$fixtureBoundary = (Join-Path $repoRoot 'obj\partition-evidence-fixture-')

try {
  $toolDir = Join-Path $fixtureRoot 'tools\RecoLearning'
  $stateDir = Join-Path $toolDir 'migration-state'
  $artifactDir = Join-Path $fixtureRoot 'artifact'
  $runtime2020 = Join-Path $fixtureRoot '铁路基本建设工程投资控制系统2020网络版V0503021201'
  $runtime2024 = Join-Path $fixtureRoot '2024铁路工程云计价系统网络版V1.0\铁路工程云计价系统网络版V1.0'
  foreach ($directory in @($toolDir,$stateDir,$artifactDir,$runtime2020,$runtime2024)) { [void][IO.Directory]::CreateDirectory($directory) }
  [IO.File]::Copy($sourceGenerator, (Join-Path $toolDir 'New-PartitionDeploymentEvidence.ps1'), $false)

  $runId = [Guid]::NewGuid().ToString('N')
  $artifactFiles = New-Object System.Collections.Generic.List[object]
  foreach ($name in @('RecoExpandPanel.dll','RecoQuotaRecommend.dll')) {
    $path = Join-Path $artifactDir $name
    [IO.File]::WriteAllText($path, 'new-' + $name, [Text.Encoding]::UTF8)
    [void]$artifactFiles.Add([pscustomobject]@{ name=$name; path=$path; sha256=(Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash })
  }
  $manifestPath = Join-Path $artifactDir 'artifact-manifest.json'
  Write-Utf8Json $manifestPath ([ordered]@{ artifact_set_id='fixture-artifact'; files=$artifactFiles.ToArray() })
  foreach ($runtime in @([pscustomobject]@{id='2020';path=$runtime2020},[pscustomobject]@{id='2024';path=$runtime2024})) {
    foreach ($name in @('RecoExpandPanel.dll','RecoQuotaRecommend.dll')) {
      [IO.File]::WriteAllText((Join-Path $runtime.path $name), ('old-' + $runtime.id + '-' + $name), [Text.Encoding]::UTF8)
    }
    $dataDir = Join-Path $runtime.path 'RecoQuotaData'
    [void][IO.Directory]::CreateDirectory($dataDir)
    [IO.File]::WriteAllText((Join-Path $dataDir 'mapping-boxes.jsonl'), ('mapping-' + $runtime.id), [Text.Encoding]::UTF8)
  }
  $statePath = Join-Path $stateDir ('partition-' + $runId + '.json')
  Write-Utf8Json $statePath ([ordered]@{ run_id=$runId; target_database='RecoLearning'; state='backed_up' })

  $fixtureGenerator = Join-Path $toolDir 'New-PartitionDeploymentEvidence.ps1'
  $startInfo = New-Object Diagnostics.ProcessStartInfo
  $startInfo.FileName = 'powershell.exe'
  $startInfo.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $fixtureGenerator + '" -RunId ' + $runId + ' -TargetDatabase RecoLearning -ArtifactManifest "' + $manifestPath + '"'
  $startInfo.UseShellExecute = $false; $startInfo.CreateNoWindow = $true
  $startInfo.RedirectStandardOutput = $true; $startInfo.RedirectStandardError = $true
  $process = New-Object Diagnostics.Process; $process.StartInfo = $startInfo
  try {
    Assert-True $process.Start() 'Could not start the deployment evidence generator fixture.'
    $stdout = $process.StandardOutput.ReadToEnd(); $stderr = $process.StandardError.ReadToEnd(); $process.WaitForExit()
    Assert-True ($process.ExitCode -eq 0) ('Deployment evidence generator failed: ' + $stdout + $stderr)
  } finally { $process.Dispose() }

  $evidenceDirectory = Join-Path $stateDir ('partition-' + $runId + '-deployment')
  $evidencePath = Join-Path $evidenceDirectory 'deployment-evidence.json'
  $evidence = Get-Content -LiteralPath $evidencePath -Raw -Encoding UTF8 | ConvertFrom-Json
  Assert-True (@($evidence.runtime_directories).Count -eq 2) 'Runtime directory evidence count mismatch.'
  Assert-True (@($evidence.targets).Count -eq 4) 'DLL target evidence count mismatch.'
  Assert-True (@($evidence.mapping_files).Count -eq 2) 'Mapping evidence count mismatch.'
  foreach ($target in @($evidence.targets)) {
    Assert-True (Test-Path -LiteralPath ([string]$target.rollback_path) -PathType Leaf) 'Rollback DLL is missing.'
    Assert-True ((Get-FileHash -LiteralPath ([string]$target.rollback_path) -Algorithm SHA256).Hash -eq [string]$target.old_sha256) 'Rollback DLL hash mismatch.'
    Assert-True ((Get-FileHash -LiteralPath ([string]$target.target_path) -Algorithm SHA256).Hash -eq [string]$target.old_sha256) 'Generator changed a runtime DLL.'
  }
  foreach ($runtimePath in @($runtime2020,$runtime2024)) {
    Assert-True (@(Get-ChildItem -LiteralPath $runtimePath -File -Filter '*.sentinel').Count -eq 0) 'Generator left a runtime sentinel.'
  }

  $migration = [IO.File]::ReadAllText($sourceMigration, [Text.Encoding]::UTF8)
  $contractStart = $migration.IndexOf('function Read-ArtifactManifest', [StringComparison]::Ordinal)
  $contractEnd = $migration.IndexOf('function Assert-BackupPath', $contractStart, [StringComparison]::Ordinal)
  Assert-True ($contractStart -ge 0 -and $contractEnd -gt $contractStart) 'Could not extract the deployment evidence contract.'
  $repoRoot = $fixtureRoot
  $TargetDatabase = 'RecoLearning'
  $RunId = $runId
  Invoke-Expression $migration.Substring($contractStart, $contractEnd - $contractStart)
  $artifact = Read-ArtifactManifest $manifestPath
  [void](Read-DeploymentEvidence $evidencePath $artifact)

  function Write-TamperedEvidence([string]$Name, [scriptblock]$Mutate) {
    $copy = ([IO.File]::ReadAllText($evidencePath, [Text.Encoding]::UTF8) | ConvertFrom-Json)
    & $Mutate $copy
    $path = Join-Path $evidenceDirectory ($Name + '.json')
    Write-Utf8Json $path $copy
    return $path
  }

  $twoTargets = Write-TamperedEvidence 'two-targets' { param($x); $x.targets=@($x.targets | Select-Object -First 2) }
  Assert-Rejected { Read-DeploymentEvidence $twoTargets $artifact } 'Validator accepted two DLL targets.'
  $fiveTargets = Write-TamperedEvidence 'five-targets' { param($x); $x.targets=@($x.targets)+@($x.targets[0]) }
  Assert-Rejected { Read-DeploymentEvidence $fiveTargets $artifact } 'Validator accepted a fifth DLL target.'
  $duplicateTarget = Write-TamperedEvidence 'duplicate-target' { param($x); $x.targets[3].target_path=$x.targets[0].target_path }
  Assert-Rejected { Read-DeploymentEvidence $duplicateTarget $artifact } 'Validator accepted a duplicate target path.'
  $thirdDirectory = Join-Path $fixtureRoot 'third-runtime'
  [void][IO.Directory]::CreateDirectory($thirdDirectory)
  [IO.File]::Copy([string]$evidence.targets[3].target_path, (Join-Path $thirdDirectory ([string]$evidence.targets[3].dll_name)), $false)
  $thirdTarget = Write-TamperedEvidence 'third-runtime' { param($x); $x.targets[3].target_path=Join-Path $thirdDirectory ([string]$x.targets[3].dll_name) }
  Assert-Rejected { Read-DeploymentEvidence $thirdTarget $artifact } 'Validator accepted a third runtime directory.'
  $crossRollback = Write-TamperedEvidence 'cross-rollback' { param($x); $x.targets[0].rollback_path=$x.targets[1].rollback_path }
  Assert-Rejected { Read-DeploymentEvidence $crossRollback $artifact } 'Validator accepted a rollback DLL from another matrix target.'
  $oneMapping = Write-TamperedEvidence 'one-mapping' { param($x); $x.mapping_files=@($x.mapping_files | Select-Object -First 1) }
  Assert-Rejected { Read-DeploymentEvidence $oneMapping $artifact } 'Validator accepted one mapping file.'

  Write-Host 'PASS W3 four-target/two-mapping evidence generation, rollback copies, sentinel cleanup, and tamper rejection'
}
finally {
  if ([IO.Directory]::Exists($fixtureRoot)) {
    $resolvedFixture = [IO.Path]::GetFullPath($fixtureRoot)
    if (-not $resolvedFixture.StartsWith($fixtureBoundary, [StringComparison]::OrdinalIgnoreCase)) { throw 'Fixture cleanup escaped the expected workspace boundary.' }
    $fixtureItem = Get-Item -LiteralPath $resolvedFixture -Force
    if (($fixtureItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Fixture cleanup target is a reparse point.' }
    [IO.Directory]::Delete($resolvedFixture, $true)
  }
}
