[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)]
  [ValidatePattern('^[a-fA-F0-9]{32}$')][string]$RunId,

  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()][string]$TargetDatabase,

  [Parameter(Mandatory = $true)]
  [ValidateNotNullOrEmpty()][string]$ArtifactManifest
)

$ErrorActionPreference = 'Stop'

if (-not [string]::Equals($TargetDatabase, 'RecoLearning', [StringComparison]::Ordinal)) {
  throw 'Deployment evidence generation is permanently restricted to the exact database RecoLearning.'
}

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
$workspaceBoundary = $repoRoot.TrimEnd('\') + '\'
$stateDirectory = Join-Path $PSScriptRoot 'migration-state'

function Assert-SoftwareStopped {
  $running = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -in @('RejjNet2020','ReJJGSNet2024','ReJJQDNet2024') })
  if ($running.Count -ne 0) {
    throw ('Deployment evidence generation requires all target software processes stopped: ' + (($running | ForEach-Object { $_.ProcessName + ':' + $_.Id }) -join ','))
  }
}

function Assert-OutboxEmpty {
  $pending = New-Object System.Collections.Generic.List[string]
  foreach ($file in @(Get-ChildItem -LiteralPath $repoRoot -Recurse -File -Filter 'learning-db-outbox.jsonl' -ErrorAction SilentlyContinue)) {
    foreach ($line in [IO.File]::ReadLines($file.FullName, [Text.Encoding]::UTF8)) {
      if (-not [string]::IsNullOrWhiteSpace($line)) { [void]$pending.Add($file.FullName); break }
    }
  }
  if ($pending.Count -ne 0) { throw ('Learning outbox is not empty: ' + ($pending -join ';')) }
}

function Get-ApprovedRuntimeDirectories {
  return @(
    [pscustomobject]@{ runtime_id='2020'; path=(Join-Path $repoRoot '铁路基本建设工程投资控制系统2020网络版V0503021201') },
    [pscustomobject]@{ runtime_id='2024'; path=(Join-Path $repoRoot '2024铁路工程云计价系统网络版V1.0\铁路工程云计价系统网络版V1.0') }
  )
}

function Read-Artifact([string]$Path) {
  $resolved = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
  if (-not $resolved.StartsWith($workspaceBoundary, [StringComparison]::OrdinalIgnoreCase)) { throw 'Artifact manifest must stay inside the workspace.' }
  $directory = Split-Path -Parent $resolved
  $actualFiles = @(Get-ChildItem -LiteralPath $directory -File -Force | Select-Object -ExpandProperty Name | Sort-Object)
  $allowed = @('RecoExpandPanel.dll','RecoQuotaRecommend.dll','artifact-manifest.json') | Sort-Object
  if ([string]::Join('|',$actualFiles) -ne [string]::Join('|',$allowed)) { throw 'Artifact directory whitelist mismatch.' }
  $manifest = Get-Content -LiteralPath $resolved -Raw -Encoding UTF8 | ConvertFrom-Json
  if ([string]::IsNullOrWhiteSpace([string]$manifest.artifact_set_id) -or @($manifest.files).Count -ne 2) { throw 'Artifact manifest identity or DLL count is invalid.' }
  $seen = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
  foreach ($file in @($manifest.files)) {
    $name = [string]$file.name
    if ($name -notin @('RecoExpandPanel.dll','RecoQuotaRecommend.dll') -or -not $seen.Add($name)) { throw "Artifact manifest contains an invalid or duplicate DLL: $name" }
    $filePath = Join-Path $directory $name
    if ((Get-FileHash -LiteralPath $filePath -Algorithm SHA256).Hash -ne [string]$file.sha256) { throw "Artifact hash mismatch: $name" }
  }
  return [pscustomobject]@{ path=$resolved; sha256=(Get-FileHash -LiteralPath $resolved -Algorithm SHA256).Hash; directory=$directory; manifest=$manifest }
}

function Test-Sentinel([string]$Directory, [string]$RuntimeId) {
  $token = [Guid]::NewGuid().ToString('N')
  $sentinel = Join-Path $Directory ('.reco-partition-' + $RunId.ToLowerInvariant() + '-' + $token + '.sentinel')
  $content = 'run_id=' + $RunId.ToLowerInvariant() + "`nruntime_id=" + $RuntimeId + "`ntoken=" + $token
  try {
    if ([IO.File]::Exists($sentinel)) { throw "Sentinel path already exists: $sentinel" }
    [IO.File]::WriteAllText($sentinel, $content, (New-Object Text.UTF8Encoding($true)))
    $persisted = [IO.File]::ReadAllText($sentinel, [Text.Encoding]::UTF8)
    if (-not [string]::Equals($persisted, $content, [StringComparison]::Ordinal)) { throw "Sentinel content mismatch: $Directory" }
    return (Get-FileHash -LiteralPath $sentinel -Algorithm SHA256).Hash
  }
  finally {
    if ([IO.File]::Exists($sentinel)) { [IO.File]::Delete($sentinel) }
    if ([IO.File]::Exists($sentinel)) { throw "Sentinel could not be removed: $sentinel" }
  }
}

Assert-SoftwareStopped
Assert-OutboxEmpty

$statePath = Join-Path $stateDirectory ('partition-' + $RunId.ToLowerInvariant() + '.json')
if (-not (Test-Path -LiteralPath $statePath -PathType Leaf)) { throw "Migration state not found: $statePath" }
$state = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([string]$state.run_id -ne $RunId -or [string]$state.target_database -ne $TargetDatabase -or [string]$state.state -ne 'backed_up') {
  throw 'Deployment evidence requires the matching RecoLearning run in backed_up state.'
}

$artifact = Read-Artifact $ArtifactManifest
$runtimes = @(Get-ApprovedRuntimeDirectories)
if ($runtimes.Count -ne 2) { throw 'Approved runtime directory count must be exactly two.' }
$runtimePaths = New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
$targetSnapshots = New-Object System.Collections.Generic.List[object]
$mappingSnapshots = New-Object System.Collections.Generic.List[object]
foreach ($runtime in $runtimes) {
  $runtime.path = (Resolve-Path -LiteralPath ([string]$runtime.path) -ErrorAction Stop).Path
  $directoryItem = Get-Item -LiteralPath $runtime.path -Force
  if (($directoryItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Runtime directory cannot be a reparse point: $($runtime.path)" }
  if (-not $runtimePaths.Add([string]$runtime.path)) { throw "Duplicate approved runtime directory: $($runtime.path)" }
  foreach ($artifactFile in @($artifact.manifest.files | Sort-Object name)) {
    $targetPath = Join-Path $runtime.path ([string]$artifactFile.name)
    if (-not (Test-Path -LiteralPath $targetPath -PathType Leaf)) { throw "Existing deployment target is missing: $targetPath" }
    $targetItem = Get-Item -LiteralPath $targetPath -Force
    if (($targetItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Deployment target cannot be a reparse point: $targetPath" }
    [void]$targetSnapshots.Add([pscustomobject]@{
      runtime_id=[string]$runtime.runtime_id; target_path=$targetPath; dll_name=[string]$artifactFile.name
      old_sha256=(Get-FileHash -LiteralPath $targetPath -Algorithm SHA256).Hash; new_sha256=[string]$artifactFile.sha256
    })
  }
  $mappingPath = Join-Path $runtime.path 'RecoQuotaData\mapping-boxes.jsonl'
  $mappingExists = Test-Path -LiteralPath $mappingPath -PathType Leaf
  if ($mappingExists -and ((Get-Item -LiteralPath $mappingPath -Force).Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw "Mapping file cannot be a reparse point: $mappingPath" }
  [void]$mappingSnapshots.Add([pscustomobject]@{
    runtime_id=[string]$runtime.runtime_id; path=$mappingPath; exists=[bool]$mappingExists
    sha256=if($mappingExists){(Get-FileHash -LiteralPath $mappingPath -Algorithm SHA256).Hash}else{''}
    bytes=if($mappingExists){[long](Get-Item -LiteralPath $mappingPath).Length}else{[long]0}
  })
}
if ($targetSnapshots.Count -ne 4 -or $mappingSnapshots.Count -ne 2) { throw 'Deployment evidence preflight did not produce the complete four-target/two-mapping matrix.' }

$finalDirectory = Join-Path $stateDirectory ('partition-' + $RunId.ToLowerInvariant() + '-deployment')
if (Test-Path -LiteralPath $finalDirectory) { throw "Deployment evidence directory already exists: $finalDirectory" }
$stagingDirectory = $finalDirectory + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
[void][IO.Directory]::CreateDirectory($stagingDirectory)
try {
  $runtimeRows = New-Object System.Collections.Generic.List[object]
  $sentinelByRuntime = @{}
  foreach ($runtime in $runtimes) {
    $sentinelHash = Test-Sentinel ([string]$runtime.path) ([string]$runtime.runtime_id)
    $sentinelByRuntime[[string]$runtime.runtime_id] = $sentinelHash
    [void]$runtimeRows.Add([pscustomobject]@{ runtime_id=[string]$runtime.runtime_id; path=[string]$runtime.path; sentinel_verified=$true; sentinel_sha256=$sentinelHash })
  }

  $targetRows = New-Object System.Collections.Generic.List[object]
  foreach ($target in @($targetSnapshots | Sort-Object runtime_id,dll_name)) {
    $rollbackRelative = Join-Path (Join-Path 'rollback' ([string]$target.runtime_id)) ([string]$target.dll_name)
    $rollbackStaging = Join-Path $stagingDirectory $rollbackRelative
    [void][IO.Directory]::CreateDirectory((Split-Path -Parent $rollbackStaging))
    [IO.File]::Copy([string]$target.target_path, $rollbackStaging, $false)
    $rollbackHash = (Get-FileHash -LiteralPath $rollbackStaging -Algorithm SHA256).Hash
    if ($rollbackHash -ne [string]$target.old_sha256) { throw "Rollback copy hash mismatch: $($target.target_path)" }
    [void]$targetRows.Add([pscustomobject]@{
      runtime_id=[string]$target.runtime_id; target_path=[string]$target.target_path; dll_name=[string]$target.dll_name
      old_sha256=[string]$target.old_sha256; new_sha256=[string]$target.new_sha256
      rollback_path=(Join-Path $finalDirectory $rollbackRelative); rollback_sha256=$rollbackHash; sentinel_verified=$true
    })
  }

  $evidence = [ordered]@{
    run_id=$RunId.ToLowerInvariant(); target_database=$TargetDatabase; generated_at_utc=[DateTime]::UtcNow.ToString('o')
    artifact_set_id=[string]$artifact.manifest.artifact_set_id; artifact_manifest_path=[string]$artifact.path; artifact_manifest_sha256=[string]$artifact.sha256
    runtime_directories=$runtimeRows.ToArray(); targets=$targetRows.ToArray(); mapping_files=@($mappingSnapshots.ToArray() | Sort-Object runtime_id)
  }
  $evidenceStaging = Join-Path $stagingDirectory 'deployment-evidence.json'
  [IO.File]::WriteAllText($evidenceStaging, ($evidence | ConvertTo-Json -Depth 10), (New-Object Text.UTF8Encoding($true)))
  [IO.Directory]::Move($stagingDirectory, $finalDirectory)
  $stagingDirectory = $null
}
finally {
  if ($stagingDirectory -and [IO.Directory]::Exists($stagingDirectory)) { [IO.Directory]::Delete($stagingDirectory, $true) }
}

$evidencePath = Join-Path $finalDirectory 'deployment-evidence.json'
[pscustomobject]@{
  RunId=$RunId.ToLowerInvariant(); State='evidence_ready'; EvidencePath=$evidencePath
  EvidenceSHA256=(Get-FileHash -LiteralPath $evidencePath -Algorithm SHA256).Hash
  TargetCount=4; MappingFileCount=2; RuntimeDirectoryCount=2
}
