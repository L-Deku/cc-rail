[CmdletBinding(DefaultParameterSetName = 'Isolate')]
param(
  [Parameter(Mandatory = $true, ParameterSetName = 'Isolate')][switch]$IsolateMappings,
  [Parameter(Mandatory = $true, ParameterSetName = 'Deploy')][switch]$DeployDlls,
  [Parameter(Mandatory = $true)][ValidatePattern('^[a-fA-F0-9]{32}$')][string]$RunId,
  [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$TargetDatabase,
  [Parameter(ParameterSetName = 'Deploy')][ValidateRange(0,4)][int]$InjectFailureAfterTarget = 0
)

$ErrorActionPreference = 'Stop'
if (-not [string]::Equals($TargetDatabase, 'RecoLearning', [StringComparison]::Ordinal)) { throw 'File cutover is permanently restricted to the exact database RecoLearning.' }
if ($InjectFailureAfterTarget -gt 0 -and [Environment]::GetEnvironmentVariable('RECO_PARTITION_CUTOVER_TEST') -ne '1') { throw 'Failure injection is test-only.' }

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
$stateDirectory = Join-Path $PSScriptRoot 'migration-state'
$statePath = Join-Path $stateDirectory ('partition-' + $RunId.ToLowerInvariant() + '.json')

function Assert-SoftwareStopped {
  $running = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -in @('RejjNet2020','ReJJGSNet2024','ReJJQDNet2024') })
  if ($running.Count -ne 0) { throw ('File cutover requires all target software processes stopped: ' + (($running | ForEach-Object { $_.ProcessName + ':' + $_.Id }) -join ',')) }
}

function Write-JsonAtomic([string]$Path, $Value) {
  $temp = $Path + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
  $previous = $Path + '.' + [Guid]::NewGuid().ToString('N') + '.previous'
  try {
    [IO.File]::WriteAllText($temp, ($Value | ConvertTo-Json -Depth 12), (New-Object Text.UTF8Encoding($true)))
    if ([IO.File]::Exists($Path)) { [IO.File]::Replace($temp,$Path,$previous,$true) } else { [IO.File]::Move($temp,$Path) }
  } finally {
    if ([IO.File]::Exists($temp)) { [IO.File]::Delete($temp) }
    if ([IO.File]::Exists($previous)) { [IO.File]::Delete($previous) }
  }
}

function Get-TextSha256([string]$Value) {
  $sha = [Security.Cryptography.SHA256]::Create()
  try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Value)))).Replace('-','') }
  finally { $sha.Dispose() }
}

function Restore-TouchedFiles($Touched) {
  $errors = New-Object System.Collections.Generic.List[string]
  foreach ($target in @($Touched.ToArray() | Sort-Object order -Descending)) {
    $temp = [string]$target.target_path + '.' + [Guid]::NewGuid().ToString('N') + '.restore.tmp'
    $previous = [string]$target.target_path + '.' + [Guid]::NewGuid().ToString('N') + '.failed-new'
    try {
      [IO.File]::Copy([string]$target.rollback_path,$temp,$false)
      [IO.File]::Replace($temp,[string]$target.target_path,$previous,$true)
      if ((Get-FileHash -LiteralPath ([string]$target.target_path) -Algorithm SHA256).Hash -ne [string]$target.old_sha256) { throw 'Restored target hash mismatch.' }
    } catch { [void]$errors.Add(([string]$target.target_path + ': ' + $_.Exception.GetType().Name)) }
    finally {
      if ([IO.File]::Exists($temp)) { [IO.File]::Delete($temp) }
      if ([IO.File]::Exists($previous)) { [IO.File]::Delete($previous) }
    }
  }
  return $errors.ToArray()
}

Assert-SoftwareStopped
if (-not (Test-Path -LiteralPath $statePath -PathType Leaf)) { throw "Migration state not found: $statePath" }
$state = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([string]$state.run_id -ne $RunId -or [string]$state.target_database -ne $TargetDatabase -or [string]$state.state -ne 'consumed') { throw 'File cutover requires the matching run in consumed state.' }
$evidencePath = (Resolve-Path -LiteralPath ([string]$state.deployment_evidence_path) -ErrorAction Stop).Path
$artifactPath = (Resolve-Path -LiteralPath ([string]$state.artifact_manifest_path) -ErrorAction Stop).Path
if ((Get-FileHash -LiteralPath $evidencePath -Algorithm SHA256).Hash -ne [string]$state.deployment_evidence_sha256) { throw 'Deployment evidence changed after W3.' }
if ((Get-FileHash -LiteralPath $artifactPath -Algorithm SHA256).Hash -ne [string]$state.artifact_manifest_sha256) { throw 'Artifact manifest changed after W3.' }
$evidence = Get-Content -LiteralPath $evidencePath -Raw -Encoding UTF8 | ConvertFrom-Json
$artifact = Get-Content -LiteralPath $artifactPath -Raw -Encoding UTF8 | ConvertFrom-Json
if (@($evidence.targets).Count -ne 4 -or @($evidence.mapping_files).Count -ne 2) { throw 'Cutover evidence matrix is incomplete.' }
$runDirectory = Split-Path -Parent $evidencePath

if ($IsolateMappings) {
  $resultPath = Join-Path $runDirectory 'mapping-isolation.json'
  if (Test-Path -LiteralPath $resultPath) { throw "Mapping isolation result already exists: $resultPath" }
  $lineAudit = New-Object System.Collections.Generic.List[string]
  [void]$lineAudit.Add('runtime_id,line_no,record_type,line_sha256')
  $plans = New-Object System.Collections.Generic.List[object]
  foreach ($mapping in @($evidence.mapping_files | Sort-Object runtime_id)) {
    $source = [string]$mapping.path
    $archive = $source + '.pre-partition-' + $RunId.ToLowerInvariant() + '.bak'
    if ([bool]$mapping.exists) {
      if (-not (Test-Path -LiteralPath $source -PathType Leaf) -or (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne [string]$mapping.sha256) { throw "Mapping file changed after W3: $source" }
      if (Test-Path -LiteralPath $archive) { throw "Mapping archive already exists: $archive" }
      $counts = @{ mapping_box=0; mapping_context=0; other=0; total=0 }
      $lineNo = 0
      foreach ($line in [IO.File]::ReadLines($source,[Text.Encoding]::UTF8)) {
        $lineNo++; $counts.total++
        $recordType = 'other'
        if (-not [string]::IsNullOrWhiteSpace($line)) {
          try { $parsed=$line | ConvertFrom-Json; if([string]$parsed.record_type -in @('mapping_box','mapping_context')){$recordType=[string]$parsed.record_type} } catch {}
        }
        $counts[$recordType]++
        [void]$lineAudit.Add(([string]$mapping.runtime_id + ',' + $lineNo + ',' + $recordType + ',' + (Get-TextSha256 $line)))
      }
      [void]$plans.Add([pscustomobject]@{runtime_id=[string]$mapping.runtime_id;source_path=$source;archive_path=$archive;sha256=[string]$mapping.sha256;before_counts=$counts})
    } else { [void]$plans.Add([pscustomobject]@{runtime_id=[string]$mapping.runtime_id;source_path=$source;archive_path='';sha256='';before_counts=@{mapping_box=0;mapping_context=0;other=0;total=0}}) }
  }
  $moved = New-Object System.Collections.Generic.List[object]
  try {
    foreach ($plan in $plans.ToArray()) {
      if ([string]$plan.archive_path -eq '') { continue }
      [IO.File]::Move([string]$plan.source_path,[string]$plan.archive_path)
      [void]$moved.Add($plan)
      if ([IO.File]::Exists([string]$plan.source_path) -or (Get-FileHash -LiteralPath ([string]$plan.archive_path) -Algorithm SHA256).Hash -ne [string]$plan.sha256) { throw "Mapping isolation verification failed: $($plan.source_path)" }
    }
  } catch {
    $failure=$_
    foreach ($plan in @($moved.ToArray() | Sort-Object runtime_id -Descending)) { if([IO.File]::Exists([string]$plan.archive_path) -and -not [IO.File]::Exists([string]$plan.source_path)){[IO.File]::Move([string]$plan.archive_path,[string]$plan.source_path)} }
    throw $failure
  }
  $auditPath = Join-Path $runDirectory 'mapping-isolation-lines.csv'
  [IO.File]::WriteAllLines($auditPath,$lineAudit.ToArray(),(New-Object Text.UTF8Encoding($true)))
  $result=[ordered]@{run_id=$RunId.ToLowerInvariant();state='isolated';completed_at_utc=[DateTime]::UtcNow.ToString('o');files=$plans.ToArray();line_audit_path=$auditPath;line_audit_sha256=(Get-FileHash -LiteralPath $auditPath -Algorithm SHA256).Hash}
  Write-JsonAtomic $resultPath $result
  [pscustomobject]@{RunId=$RunId;State='isolated';ResultPath=$resultPath;FileCount=$plans.Count;ArchivedCount=$moved.Count}
  return
}

if ($DeployDlls) {
  $resultPath=Join-Path $runDirectory 'dll-deployment.json'
  if (Test-Path -LiteralPath $resultPath) { throw "DLL deployment result already exists: $resultPath" }
  $isolationPath = Join-Path $runDirectory 'mapping-isolation.json'
  if (-not (Test-Path -LiteralPath $isolationPath -PathType Leaf)) { throw 'D1 requires completed mapping isolation evidence.' }
  $isolation = Get-Content -LiteralPath $isolationPath -Raw -Encoding UTF8 | ConvertFrom-Json
  foreach($file in @($isolation.files)){if([string]$file.archive_path -ne '' -and ((Test-Path -LiteralPath ([string]$file.source_path)) -or (Get-FileHash -LiteralPath ([string]$file.archive_path) -Algorithm SHA256).Hash -ne [string]$file.sha256)){throw 'Mapping isolation evidence changed before D1.'}}
  $artifactDirectory=Split-Path -Parent $artifactPath
  $binDirectory=Join-Path $repoRoot 'RecoQuotaRecommend\bin'
  $binRollbackDirectory=Join-Path $runDirectory 'bin-rollback'
  if (Test-Path -LiteralPath $binRollbackDirectory) { throw "Bin rollback directory already exists: $binRollbackDirectory" }
  [void][IO.Directory]::CreateDirectory($binRollbackDirectory)
  $promoted=New-Object System.Collections.Generic.List[object]
  $touched=New-Object System.Collections.Generic.List[object]
  $order=0
  try {
    foreach($artifactFile in @($artifact.files | Sort-Object name)){
      $name=[string]$artifactFile.name
      if($name -notin @('RecoExpandPanel.dll','RecoQuotaRecommend.dll')){throw "Unexpected artifact DLL: $name"}
      $artifactSource=Join-Path $artifactDirectory $name
      $binPath=Join-Path $binDirectory $name
      if(-not (Test-Path -LiteralPath $binPath -PathType Leaf)){throw "Bin source target is missing: $binPath"}
      $binOldHash=(Get-FileHash -LiteralPath $binPath -Algorithm SHA256).Hash
      $binRollback=Join-Path $binRollbackDirectory $name
      [IO.File]::Copy($binPath,$binRollback,$false)
      if((Get-FileHash -LiteralPath $binRollback -Algorithm SHA256).Hash -ne $binOldHash){throw "Bin rollback copy hash mismatch: $binPath"}
      $temp=$binPath+'.'+[Guid]::NewGuid().ToString('N')+'.new.tmp'; $previous=$binPath+'.'+[Guid]::NewGuid().ToString('N')+'.previous'
      try{
        [IO.File]::Copy($artifactSource,$temp,$false)
        [IO.File]::Replace($temp,$binPath,$previous,$true)
        [void]$promoted.Add([pscustomobject]@{order=$promoted.Count+1;target_path=$binPath;rollback_path=$binRollback;old_sha256=$binOldHash})
        if((Get-FileHash -LiteralPath $binPath -Algorithm SHA256).Hash -ne [string]$artifactFile.sha256){throw "Promoted bin hash mismatch: $binPath"}
      } finally {if([IO.File]::Exists($temp)){[IO.File]::Delete($temp)};if([IO.File]::Exists($previous)){[IO.File]::Delete($previous)}}
    }
    foreach($target in @($evidence.targets | Sort-Object runtime_id,dll_name)){
      $order++; $targetPath=[string]$target.target_path; $source=Join-Path $binDirectory ([string]$target.dll_name)
      if((Get-FileHash -LiteralPath $targetPath -Algorithm SHA256).Hash -ne [string]$target.old_sha256){throw "Target changed before D1: $targetPath"}
      if((Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash -ne [string]$target.new_sha256){throw "Bin source hash mismatch before D1: $source"}
      $temp=$targetPath+'.'+[Guid]::NewGuid().ToString('N')+'.new.tmp'; $previous=$targetPath+'.'+[Guid]::NewGuid().ToString('N')+'.previous'
      try{[IO.File]::Copy($source,$temp,$false); [IO.File]::Replace($temp,$targetPath,$previous,$true); [void]$touched.Add([pscustomobject]@{order=$order;target_path=$targetPath;rollback_path=[string]$target.rollback_path;old_sha256=[string]$target.old_sha256}); if((Get-FileHash -LiteralPath $targetPath -Algorithm SHA256).Hash -ne [string]$target.new_sha256){throw "Deployed target hash mismatch: $targetPath"}} finally {if([IO.File]::Exists($temp)){[IO.File]::Delete($temp)};if([IO.File]::Exists($previous)){[IO.File]::Delete($previous)}}
      if($InjectFailureAfterTarget -eq $order){throw "Injected deployment failure after target $order"}
    }
  } catch {
    $failure=$_
    $rollbackErrors=New-Object System.Collections.Generic.List[string]
    foreach($message in @(Restore-TouchedFiles $touched)){[void]$rollbackErrors.Add($message)}
    foreach($message in @(Restore-TouchedFiles $promoted)){[void]$rollbackErrors.Add($message)}
    if($rollbackErrors.Count -ne 0){throw ('D1 failed and rollback was incomplete: '+($rollbackErrors -join ';')+'; original='+$failure.Exception.GetType().Name)}
    if([IO.Directory]::Exists($binRollbackDirectory)){[IO.Directory]::Delete($binRollbackDirectory,$true)}
    throw $failure
  }
  Write-JsonAtomic $resultPath ([ordered]@{run_id=$RunId.ToLowerInvariant();state='deployed';completed_at_utc=[DateTime]::UtcNow.ToString('o');bin_source=$binDirectory;bin_rollback_directory=$binRollbackDirectory;targets=@($evidence.targets)})
  [pscustomobject]@{RunId=$RunId;State='deployed';ResultPath=$resultPath;TargetCount=$touched.Count}
}
