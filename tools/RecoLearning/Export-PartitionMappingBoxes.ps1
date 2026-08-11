[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][ValidatePattern('^[a-fA-F0-9]{32}$')][string]$RunId,
  [Parameter(Mandatory = $true)][ValidateNotNullOrEmpty()][string]$TargetDatabase,
  [string]$FixtureDataPath = ''
)

$ErrorActionPreference = 'Stop'
if (-not [string]::Equals($TargetDatabase, 'RecoLearning', [StringComparison]::Ordinal)) { throw 'Mapping export is permanently restricted to the exact database RecoLearning.' }
if ($FixtureDataPath -ne '' -and [Environment]::GetEnvironmentVariable('RECO_PARTITION_EXPORT_TEST') -ne '1') { throw 'Fixture data is test-only.' }

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path
$stateDirectory = Join-Path $PSScriptRoot 'migration-state'
$statePath = Join-Path $stateDirectory ('partition-' + $RunId.ToLowerInvariant() + '.json')
$auxiliaryCodes = @('SF','SH','SQ','ZLF','LF','YF','TLF','GF','JF','XGT1')

function Assert-SoftwareStopped {
  $running = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -in @('RejjNet2020','ReJJGSNet2024','ReJJQDNet2024') })
  if ($running.Count -ne 0) { throw ('Mapping export requires all target software processes stopped: ' + (($running | ForEach-Object { $_.ProcessName + ':' + $_.Id }) -join ',')) }
}

function Write-JsonAtomic([string]$Path, $Value) {
  $temp = $Path + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'
  $previous = $Path + '.' + [Guid]::NewGuid().ToString('N') + '.previous'
  try {
    [IO.File]::WriteAllText($temp, ($Value | ConvertTo-Json -Depth 12), (New-Object Text.UTF8Encoding($true)))
    if ([IO.File]::Exists($Path)) { [IO.File]::Replace($temp,$Path,$previous,$true) } else { [IO.File]::Move($temp,$Path) }
  }
  finally {
    if ([IO.File]::Exists($temp)) { [IO.File]::Delete($temp) }
    if ([IO.File]::Exists($previous)) { [IO.File]::Delete($previous) }
  }
}

function Get-BaseTargetCode([string]$Code) {
  $value = ([string]$Code).Trim().ToUpperInvariant()
  $multiplier = $value.IndexOfAny([char[]]@('*',[char]0x00D7))
  if ($multiplier -ge 0) { $value = $value.Substring(0,$multiplier) }
  $slash = $value.LastIndexOf('/')
  if ($slash -gt 0 -and $slash -lt $value.Length - 1 -and $value.Substring($slash + 1) -match '^\d+$') { $value = $value.Substring(0,$slash) }
  return $value.Replace('参','').Replace('换','').Replace('借','').Trim()
}

function Get-InferredTargetKind([string]$Code) {
  $baseCode = Get-BaseTargetCode $Code
  if ($baseCode -match '^\d+$') { return 'material' }
  if ($baseCode.IndexOf('-',[StringComparison]::Ordinal) -ge 0) { return 'quota' }
  return 'aux'
}

function Get-HexUtf8([string]$Value) {
  return ([BitConverter]::ToString([Text.Encoding]::UTF8.GetBytes([string]$Value))).Replace('-','')
}

function Get-ApprovedRuntimes {
  return @(
    [pscustomobject]@{runtime_id='2020';path=(Join-Path $repoRoot '铁路基本建设工程投资控制系统2020网络版V0503021201')},
    [pscustomobject]@{runtime_id='2024';path=(Join-Path $repoRoot '2024铁路工程云计价系统网络版V1.0\铁路工程云计价系统网络版V1.0')}
  )
}

function Get-RuntimePartitions([string]$Path) {
  $partitions = New-Object System.Collections.Generic.List[string]
  if (Test-Path -LiteralPath (Join-Path $Path 'RejjNet2020.exe') -PathType Leaf) { [void]$partitions.Add('2020') }
  if ((Test-Path -LiteralPath (Join-Path $Path 'ReJJGSNet2024.exe') -PathType Leaf) -or
      (Test-Path -LiteralPath (Join-Path $Path 'ReJJQDNet2024.exe') -PathType Leaf)) { [void]$partitions.Add('2024') }
  return $partitions.ToArray()
}

function Read-MetadataIndex([string]$Path, [string]$CodeProperty, [string]$NameProperty, [string]$UnitProperty) {
  if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { throw "Metadata index is missing: $Path" }
  $index = @{}
  foreach ($line in [IO.File]::ReadLines($Path,[Text.Encoding]::UTF8)) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $row = $line | ConvertFrom-Json
    $code = ([string]$row.$CodeProperty).Trim().ToUpperInvariant()
    if ($code -eq '') { continue }
    if (-not $index.ContainsKey($code)) { $index[$code] = @{} }
    $name = ([string]$row.$NameProperty).Trim()
    $unit = ([string]$row.$UnitProperty).Trim()
    $identity = $name + "`n" + $unit
    if (-not $index[$code].ContainsKey($identity)) { $index[$code][$identity] = [pscustomobject]@{name=$name;unit=$unit} }
  }
  return $index
}

function Resolve-IndexedMetadata($Index, [string]$Code) {
  $key = ([string]$Code).Trim().ToUpperInvariant()
  if (-not $Index.ContainsKey($key)) { return $null }
  $matches = @($Index[$key].Values)
  if ($matches.Count -ne 1 -or [string]::IsNullOrWhiteSpace([string]$matches[0].name) -or [string]::IsNullOrWhiteSpace([string]$matches[0].unit)) { return $null }
  return $matches[0]
}

function Get-ExportData {
  if ($FixtureDataPath -ne '') {
    $resolved = (Resolve-Path -LiteralPath $FixtureDataPath -ErrorAction Stop).Path
    return (Get-Content -LiteralPath $resolved -Raw -Encoding UTF8 | ConvertFrom-Json)
  }
  . (Join-Path $PSScriptRoot 'Common.ps1')
  $maps = Invoke-RecoQuery -Database $TargetDatabase -Sql @'
SELECT software_partition,signature,box_id,weight,accepted_count,corrected_count,rejected_count,
       CONVERT(NVARCHAR(33),last_used_at,126) AS last_used_at
FROM dbo.SignatureBoxMap
WHERE software_partition IN (N'2020',N'2024');
'@
  $aliases = Invoke-RecoQuery -Database $TargetDatabase -Sql 'SELECT signature,raw_name,quantity_unit FROM dbo.QuantityAlias;'
  $targets = Invoke-RecoQuery -Database $TargetDatabase -Sql 'SELECT box_id,target_kind,target_code,target_name,target_unit FROM dbo.QuotaBoxTarget;'
  return [pscustomobject]@{maps=@($maps.Rows);aliases=@($aliases.Rows);targets=@($targets.Rows)}
}

if ($FixtureDataPath -eq '') { Assert-SoftwareStopped }
if (-not (Test-Path -LiteralPath $statePath -PathType Leaf)) { throw "Migration state not found: $statePath" }
$state = Get-Content -LiteralPath $statePath -Raw -Encoding UTF8 | ConvertFrom-Json
if ([string]$state.run_id -ne $RunId -or [string]$state.target_database -ne $TargetDatabase -or [string]$state.state -ne 'consumed') { throw 'Mapping export requires the matching run in consumed state.' }
$evidencePath = (Resolve-Path -LiteralPath ([string]$state.deployment_evidence_path) -ErrorAction Stop).Path
if ((Get-FileHash -LiteralPath $evidencePath -Algorithm SHA256).Hash -ne [string]$state.deployment_evidence_sha256) { throw 'Deployment evidence changed before D4.' }
$evidence = Get-Content -LiteralPath $evidencePath -Raw -Encoding UTF8 | ConvertFrom-Json
$runDirectory = Split-Path -Parent $evidencePath
$isolationPath = Join-Path $runDirectory 'mapping-isolation.json'
$deploymentPath = Join-Path $runDirectory 'dll-deployment.json'
if (-not (Test-Path -LiteralPath $isolationPath -PathType Leaf) -or -not (Test-Path -LiteralPath $deploymentPath -PathType Leaf)) { throw 'D4 requires completed C3 isolation and D1 deployment evidence.' }
$isolation = Get-Content -LiteralPath $isolationPath -Raw -Encoding UTF8 | ConvertFrom-Json

$runtimes = New-Object System.Collections.Generic.List[object]
$runtimeById = @{}
foreach ($runtime in @(Get-ApprovedRuntimes)) {
  $runtime.path = (Resolve-Path -LiteralPath ([string]$runtime.path) -ErrorAction Stop).Path
  [string[]]$installedPartitions = @(Get-RuntimePartitions ([string]$runtime.path))
  if ($installedPartitions.Count -eq 0) { throw "Approved runtime contains no recognized host executable: $($runtime.path)" }
  $runtime | Add-Member -NotePropertyName partitions -NotePropertyValue $installedPartitions -Force
  $runtimeById[[string]$runtime.runtime_id] = $runtime
  [void]$runtimes.Add($runtime)
}
if ($runtimeById.Count -ne 2) { throw 'D4 runtime directory count must be exactly two.' }
foreach ($file in @($isolation.files)) {
  $runtimeId = [string]$file.runtime_id
  if (-not $runtimeById.ContainsKey($runtimeId)) { throw "Unexpected isolated runtime: $runtimeId" }
  $expectedOutput = Join-Path ([string]$runtimeById[$runtimeId].path) 'RecoQuotaData\mapping-boxes.jsonl'
  if (-not [string]::Equals([IO.Path]::GetFullPath([string]$file.source_path),[IO.Path]::GetFullPath($expectedOutput),[StringComparison]::OrdinalIgnoreCase)) { throw "Isolation output path mismatch: $runtimeId" }
  if (Test-Path -LiteralPath $expectedOutput) { throw "D4 output already exists: $expectedOutput" }
  if (([string]$file.archive_path -ne '') -and ((Get-FileHash -LiteralPath ([string]$file.archive_path) -Algorithm SHA256).Hash -ne [string]$file.sha256)) { throw "Mapping archive changed before D4: $runtimeId" }
  $runtimeById[$runtimeId] | Add-Member -NotePropertyName output_path -NotePropertyValue $expectedOutput -Force
}

$metadataByPartition = @{}
foreach ($partition in @('2020','2024')) {
  $dataDirectory = Join-Path ([string]$runtimeById[$partition].path) 'RecoQuotaData'
  $metadataByPartition[$partition] = [pscustomobject]@{
    quota=(Read-MetadataIndex (Join-Path $dataDirectory 'quota-index.jsonl') 'quota_code' 'quota_name' 'quota_unit')
    material=(Read-MetadataIndex (Join-Path $dataDirectory 'material-index.jsonl') 'material_code' 'material_name' 'material_unit')
  }
}

$data = Get-ExportData
$maps = @($data.maps | Where-Object { [string]$_.software_partition -in @('2020','2024') })
$aliasesBySignature = @{}
foreach ($alias in @($data.aliases)) {
  $signature = [string]$alias.signature
  if (-not $aliasesBySignature.ContainsKey($signature)) { $aliasesBySignature[$signature] = @{} }
  $identity = ([string]$alias.raw_name) + "`n" + ([string]$alias.quantity_unit)
  if (-not $aliasesBySignature[$signature].ContainsKey($identity)) { $aliasesBySignature[$signature][$identity] = $alias }
}
$targetsByBox = @{}
foreach ($target in @($data.targets)) {
  $boxId = [string]$target.box_id
  if (-not $targetsByBox.ContainsKey($boxId)) { $targetsByBox[$boxId] = @{} }
  $identity = ([string]$target.target_kind).ToLowerInvariant() + "`n" + ([string]$target.target_code).ToUpperInvariant()
  if (-not $targetsByBox[$boxId].ContainsKey($identity)) { $targetsByBox[$boxId][$identity] = $target }
}

$resolvedByPartitionBox = @{}
$isolatedByPartitionBox = @{}
foreach ($map in $maps) {
  $partition = [string]$map.software_partition
  $boxId = [string]$map.box_id
  $partitionBox = $partition + "`n" + $boxId
  if ($resolvedByPartitionBox.ContainsKey($partitionBox) -or $isolatedByPartitionBox.ContainsKey($partitionBox)) { continue }
  if (-not $targetsByBox.ContainsKey($boxId) -or $targetsByBox[$boxId].Count -eq 0) { $isolatedByPartitionBox[$partitionBox] = 'missing_targets'; continue }
  $resolvedTargets = New-Object System.Collections.Generic.List[object]
  $failureReason = ''
  foreach ($target in @($targetsByBox[$boxId].Values)) {
    $code = ([string]$target.target_code).Trim()
    $baseCode = Get-BaseTargetCode $code
    $kind = Get-InferredTargetKind $code
    $resolved = $null
    if ($kind -eq 'aux' -or $auxiliaryCodes -contains $baseCode) {
      if (-not [string]::IsNullOrWhiteSpace([string]$target.target_name) -and -not [string]::IsNullOrWhiteSpace([string]$target.target_unit)) { $resolved = [pscustomobject]@{name=([string]$target.target_name).Trim();unit=([string]$target.target_unit).Trim()} }
    }
    elseif ($kind -eq 'quota') { $resolved = Resolve-IndexedMetadata $metadataByPartition[$partition].quota $baseCode }
    elseif ($kind -eq 'material') { $resolved = Resolve-IndexedMetadata $metadataByPartition[$partition].material $baseCode }
    else { $failureReason = 'unsupported_kind:' + $kind }
    if ($null -eq $resolved) { if ($failureReason -eq '') { $failureReason = 'missing_or_ambiguous_metadata:' + $kind + ':' + $code }; break }
    [void]$resolvedTargets.Add([pscustomobject]@{kind=$kind;code=$code;name=[string]$resolved.name;unit=[string]$resolved.unit})
  }
  if ($failureReason -ne '') { $isolatedByPartitionBox[$partitionBox] = $failureReason }
  else { $resolvedByPartitionBox[$partitionBox] = $resolvedTargets.ToArray() }
}

$linesByPartition = @{'2020'=(New-Object System.Collections.Generic.List[string]);'2024'=(New-Object System.Collections.Generic.List[string])}
$reports = @{}
foreach ($partition in @('2020','2024')) {
  $partitionMaps = @($maps | Where-Object { [string]$_.software_partition -eq $partition })
  $relationCount = $partitionMaps.Count
  $targetExpansionCount = 0
  $aliasExpansionCount = 0
  $isolatedLineCount = 0
  $quotaTargetCount = 0
  $materialTargetCount = 0
  $auxTargetCount = 0
  foreach ($map in $partitionMaps) {
    $boxId = [string]$map.box_id
    $partitionBox = $partition + "`n" + $boxId
    [object[]]$aliases = @()
    if($aliasesBySignature.ContainsKey([string]$map.signature)){ $aliases = @($aliasesBySignature[[string]$map.signature].Values) }
    [object[]]$rawTargets = @()
    if($targetsByBox.ContainsKey($boxId)){ $rawTargets = @($targetsByBox[$boxId].Values) }
    $targetExpansionCount += $rawTargets.Count
    $aliasExpansionCount += $aliases.Count
    if ($aliases.Count -eq 0) { $isolatedByPartitionBox[$partitionBox] = 'missing_alias' }
    if ($isolatedByPartitionBox.ContainsKey($partitionBox)) { $isolatedLineCount += ($rawTargets.Count * $aliases.Count); continue }
    [object[]]$resolvedTargets = @($resolvedByPartitionBox[$partitionBox])
    foreach ($target in $resolvedTargets) {
      if ([string]$target.kind -eq 'quota' -and -not ($auxiliaryCodes -contains (Get-BaseTargetCode ([string]$target.code)))) { $quotaTargetCount++ }
      elseif ([string]$target.kind -eq 'material') { $materialTargetCount++ }
      else { $auxTargetCount++ }
      foreach ($alias in $aliases) {
        $row = [ordered]@{
          record_type='mapping_box'; software_partition=$partition; method_no=''; box_id=$boxId
          target_kind=[string]$target.kind; target_code=[string]$target.code; target_name=[string]$target.name; target_unit=[string]$target.unit
          quantity_name=[string]$alias.raw_name; quantity_unit=[string]$alias.quantity_unit
          weight=([string]$map.weight); accepted_count=([string]$map.accepted_count); corrected_count=([string]$map.corrected_count); rejected_count=([string]$map.rejected_count); last_used_at=([string]$map.last_used_at)
        }
        $json = $row | ConvertTo-Json -Compress
        $sortKey = (Get-HexUtf8 $boxId) + '|' + (Get-HexUtf8 ([string]$target.kind)) + '|' + (Get-HexUtf8 ([string]$target.code)) + '|' + (Get-HexUtf8 ([string]$alias.raw_name)) + '|' + (Get-HexUtf8 ([string]$alias.quantity_unit))
        [void]$linesByPartition[$partition].Add($sortKey + "`t" + $json)
      }
    }
  }
  $isolatedKeys = @($isolatedByPartitionBox.Keys | Where-Object { $_.StartsWith($partition + "`n",[StringComparison]::Ordinal) })
  $reports[$partition] = [ordered]@{
    relation_count=$relationCount; target_expansion_count=$targetExpansionCount; alias_expansion_count=$aliasExpansionCount
    final_line_count=$linesByPartition[$partition].Count; quota_target_count=$quotaTargetCount; material_target_count=$materialTargetCount; auxiliary_target_count=$auxTargetCount
    isolated_box_count=$isolatedKeys.Count; isolated_line_count=$isolatedLineCount
    isolated_boxes=@($isolatedKeys | Sort-Object | ForEach-Object {[ordered]@{box_id=$_.Substring($partition.Length+1);reason=[string]$isolatedByPartitionBox[$_]}})
  }
}

$stagingDirectory = Join-Path $runDirectory ('mapping-export.' + [Guid]::NewGuid().ToString('N') + '.tmp')
[void][IO.Directory]::CreateDirectory($stagingDirectory)
$createdOutputs = New-Object System.Collections.Generic.List[string]
$runtimeReports = [ordered]@{}
try {
  $staged = @{}
  foreach ($runtime in @($runtimes.ToArray() | Sort-Object runtime_id)) {
    $sortableList = New-Object System.Collections.Generic.List[string]
    foreach ($partition in @($runtime.partitions | Sort-Object)) {
      foreach ($line in $linesByPartition[$partition].ToArray()) { [void]$sortableList.Add($partition + '|' + $line) }
    }
    $sortable = $sortableList.ToArray()
    [Array]::Sort($sortable,[StringComparer]::Ordinal)
    $jsonLines = New-Object System.Collections.Generic.List[string]
    foreach ($sortableLine in $sortable) { [void]$jsonLines.Add($sortableLine.Substring($sortableLine.IndexOf("`t") + 1)) }
    $stagedPath = Join-Path $stagingDirectory ([string]$runtime.runtime_id + '-mapping-boxes.jsonl')
    $content = if($jsonLines.Count -eq 0){''}else{($jsonLines.ToArray() -join "`n") + "`n"}
    [IO.File]::WriteAllText($stagedPath,$content,(New-Object Text.UTF8Encoding($false)))
    $staged[[string]$runtime.runtime_id] = $stagedPath
  }
  foreach ($runtime in @($runtimes.ToArray() | Sort-Object runtime_id)) {
    $runtimeId = [string]$runtime.runtime_id
    $outputPath = [string]$runtime.output_path
    [IO.File]::Move([string]$staged[$runtimeId],$outputPath)
    [void]$createdOutputs.Add($outputPath)
    $runtimeReports[$runtimeId] = [ordered]@{
      included_partitions=@($runtime.partitions | Sort-Object)
      final_line_count=[IO.File]::ReadAllLines($outputPath,[Text.Encoding]::UTF8).Count
      output_path=$outputPath
      output_sha256=(Get-FileHash -LiteralPath $outputPath -Algorithm SHA256).Hash
    }
  }
}
catch {
  $failure=$_
  foreach ($path in $createdOutputs.ToArray()) { if ([IO.File]::Exists($path)) { [IO.File]::Delete($path) } }
  throw $failure
}
finally {
  if ([IO.Directory]::Exists($stagingDirectory)) { [IO.Directory]::Delete($stagingDirectory,$true) }
}

$reportPath = Join-Path $runDirectory 'mapping-export-report.json'
Write-JsonAtomic $reportPath ([ordered]@{run_id=$RunId.ToLowerInvariant();state='exported';completed_at_utc=[DateTime]::UtcNow.ToString('o');partitions=[ordered]@{'2020'=$reports['2020'];'2024'=$reports['2024']};runtime_files=$runtimeReports})
[pscustomobject]@{RunId=$RunId;State='exported';ReportPath=$reportPath;Rows2020=$reports['2020'].final_line_count;Rows2024=$reports['2024'].final_line_count;IsolatedBoxes2020=$reports['2020'].isolated_box_count;IsolatedBoxes2024=$reports['2024'].isolated_box_count}
