$ErrorActionPreference = 'Stop'

function Assert-True([bool]$Condition, [string]$Message) {
  if (-not $Condition) { throw $Message }
}

function Write-Utf8Json([string]$Path, $Value) {
  [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 12), (New-Object Text.UTF8Encoding($true)))
}

function Invoke-Cutover([string]$Script, [string[]]$Arguments, [bool]$ExpectSuccess, [bool]$EnableFailureInjection) {
  $startInfo = New-Object Diagnostics.ProcessStartInfo
  $startInfo.FileName = 'powershell.exe'
  $startInfo.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $Script + '" ' + ($Arguments -join ' ')
  $startInfo.UseShellExecute = $false
  $startInfo.CreateNoWindow = $true
  $startInfo.RedirectStandardOutput = $true
  $startInfo.RedirectStandardError = $true
  if ($EnableFailureInjection) { $startInfo.EnvironmentVariables['RECO_PARTITION_CUTOVER_TEST'] = '1' }
  $process = New-Object Diagnostics.Process
  $process.StartInfo = $startInfo
  try {
    Assert-True $process.Start() 'Could not start cutover fixture process.'
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()
    if ($ExpectSuccess) { Assert-True ($process.ExitCode -eq 0) ('Cutover fixture failed: ' + $stdout + $stderr) }
    else { Assert-True ($process.ExitCode -ne 0) 'Injected cutover failure unexpectedly succeeded.' }
  }
  finally { $process.Dispose() }
}

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $testDir '..\..\..')).Path
$sourceScript = Join-Path $repoRoot 'tools\RecoLearning\Invoke-PartitionFileCutover.ps1'
$fixtureRoot = Join-Path $repoRoot ('obj\partition-cutover-fixture-' + [Guid]::NewGuid().ToString('N'))
$fixtureBoundary = Join-Path $repoRoot 'obj\partition-cutover-fixture-'

try {
  $toolDir = Join-Path $fixtureRoot 'tools\RecoLearning'
  $stateDir = Join-Path $toolDir 'migration-state'
  $artifactDir = Join-Path $fixtureRoot 'artifact'
  $binDir = Join-Path $fixtureRoot 'RecoQuotaRecommend\bin'
  $runtime2020 = Join-Path $fixtureRoot '铁路基本建设工程投资控制系统2020网络版V0503021201'
  $runtime2024 = Join-Path $fixtureRoot '2024铁路工程云计价系统网络版V1.0\铁路工程云计价系统网络版V1.0'
  foreach ($directory in @($toolDir,$stateDir,$artifactDir,$binDir,$runtime2020,$runtime2024)) { [void][IO.Directory]::CreateDirectory($directory) }
  $fixtureScript = Join-Path $toolDir 'Invoke-PartitionFileCutover.ps1'
  [IO.File]::Copy($sourceScript, $fixtureScript, $false)

  $runId = [Guid]::NewGuid().ToString('N')
  $artifactFiles = New-Object System.Collections.Generic.List[object]
  $oldBinHashes = @{}
  foreach ($name in @('RecoExpandPanel.dll','RecoQuotaRecommend.dll')) {
    $artifactPath = Join-Path $artifactDir $name
    [IO.File]::WriteAllText($artifactPath, 'new-' + $name, [Text.Encoding]::UTF8)
    [void]$artifactFiles.Add([pscustomobject]@{name=$name;path=$artifactPath;sha256=(Get-FileHash -LiteralPath $artifactPath -Algorithm SHA256).Hash})
    $binPath = Join-Path $binDir $name
    [IO.File]::WriteAllText($binPath, 'old-bin-' + $name, [Text.Encoding]::UTF8)
    $oldBinHashes[$name] = (Get-FileHash -LiteralPath $binPath -Algorithm SHA256).Hash
  }
  $manifestPath = Join-Path $artifactDir 'artifact-manifest.json'
  Write-Utf8Json $manifestPath ([ordered]@{artifact_set_id='fixture';files=$artifactFiles.ToArray()})

  $targetRows = New-Object System.Collections.Generic.List[object]
  $mappingRows = New-Object System.Collections.Generic.List[object]
  foreach ($runtime in @([pscustomobject]@{id='2020';path=$runtime2020},[pscustomobject]@{id='2024';path=$runtime2024})) {
    foreach ($file in $artifactFiles.ToArray()) {
      $targetPath = Join-Path $runtime.path ([string]$file.name)
      [IO.File]::WriteAllText($targetPath, ('old-' + $runtime.id + '-' + [string]$file.name), [Text.Encoding]::UTF8)
      $oldHash = (Get-FileHash -LiteralPath $targetPath -Algorithm SHA256).Hash
      $rollbackPath = Join-Path $stateDir ('rollback\' + $runtime.id + '\' + [string]$file.name)
      [void][IO.Directory]::CreateDirectory((Split-Path -Parent $rollbackPath))
      [IO.File]::Copy($targetPath,$rollbackPath,$false)
      [void]$targetRows.Add([pscustomobject]@{runtime_id=$runtime.id;target_path=$targetPath;dll_name=[string]$file.name;old_sha256=$oldHash;new_sha256=[string]$file.sha256;rollback_path=$rollbackPath;rollback_sha256=$oldHash})
    }
    $dataDir = Join-Path $runtime.path 'RecoQuotaData'
    [void][IO.Directory]::CreateDirectory($dataDir)
    $mappingPath = Join-Path $dataDir 'mapping-boxes.jsonl'
    [IO.File]::WriteAllLines($mappingPath,@('{"record_type":"mapping_box"}','{"record_type":"mapping_context"}','not-json'),(New-Object Text.UTF8Encoding($true)))
    [void]$mappingRows.Add([pscustomobject]@{runtime_id=$runtime.id;path=$mappingPath;exists=$true;sha256=(Get-FileHash -LiteralPath $mappingPath -Algorithm SHA256).Hash;bytes=(Get-Item -LiteralPath $mappingPath).Length})
  }

  $evidenceDir = Join-Path $stateDir ('partition-' + $runId + '-deployment')
  [void][IO.Directory]::CreateDirectory($evidenceDir)
  $evidencePath = Join-Path $evidenceDir 'deployment-evidence.json'
  Write-Utf8Json $evidencePath ([ordered]@{run_id=$runId;target_database='RecoLearning';artifact_manifest_path=$manifestPath;targets=$targetRows.ToArray();mapping_files=$mappingRows.ToArray()})
  $statePath = Join-Path $stateDir ('partition-' + $runId + '.json')
  Write-Utf8Json $statePath ([ordered]@{run_id=$runId;target_database='RecoLearning';state='consumed';deployment_evidence_path=$evidencePath;deployment_evidence_sha256=(Get-FileHash -LiteralPath $evidencePath -Algorithm SHA256).Hash;artifact_manifest_path=$manifestPath;artifact_manifest_sha256=(Get-FileHash -LiteralPath $manifestPath -Algorithm SHA256).Hash})

  Invoke-Cutover $fixtureScript @('-IsolateMappings','-RunId',$runId,'-TargetDatabase','RecoLearning') $true $false
  $isolationPath = Join-Path $evidenceDir 'mapping-isolation.json'
  $isolation = Get-Content -LiteralPath $isolationPath -Raw -Encoding UTF8 | ConvertFrom-Json
  Assert-True (@($isolation.files).Count -eq 2) 'Mapping isolation file count mismatch.'
  foreach ($file in @($isolation.files)) {
    Assert-True (-not (Test-Path -LiteralPath ([string]$file.source_path))) 'Mapping source survived isolation.'
    Assert-True ((Get-FileHash -LiteralPath ([string]$file.archive_path) -Algorithm SHA256).Hash -eq [string]$file.sha256) 'Mapping archive hash mismatch.'
    Assert-True ([int]$file.before_counts.total -eq 3) 'Mapping isolation line count mismatch.'
  }

  Invoke-Cutover $fixtureScript @('-DeployDlls','-RunId',$runId,'-TargetDatabase','RecoLearning','-InjectFailureAfterTarget','1') $false $true
  foreach ($target in $targetRows.ToArray()) { Assert-True ((Get-FileHash -LiteralPath ([string]$target.target_path) -Algorithm SHA256).Hash -eq [string]$target.old_sha256) 'Injected failure did not restore a runtime DLL.' }
  foreach ($name in $oldBinHashes.Keys) { Assert-True ((Get-FileHash -LiteralPath (Join-Path $binDir $name) -Algorithm SHA256).Hash -eq $oldBinHashes[$name]) 'Injected failure did not restore bin.' }
  Assert-True (-not (Test-Path -LiteralPath (Join-Path $evidenceDir 'bin-rollback'))) 'Failed cutover left a bin rollback directory.'

  Invoke-Cutover $fixtureScript @('-DeployDlls','-RunId',$runId,'-TargetDatabase','RecoLearning') $true $false
  foreach ($target in $targetRows.ToArray()) { Assert-True ((Get-FileHash -LiteralPath ([string]$target.target_path) -Algorithm SHA256).Hash -eq [string]$target.new_sha256) 'Successful cutover runtime hash mismatch.' }
  foreach ($file in $artifactFiles.ToArray()) { Assert-True ((Get-FileHash -LiteralPath (Join-Path $binDir ([string]$file.name)) -Algorithm SHA256).Hash -eq [string]$file.sha256) 'Successful cutover bin hash mismatch.' }
  Assert-True (Test-Path -LiteralPath (Join-Path $evidenceDir 'dll-deployment.json') -PathType Leaf) 'Successful cutover result is missing.'
  Write-Host 'PASS C3/D1 mapping isolation, bin promotion, four-target cutover, and full rollback'
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
