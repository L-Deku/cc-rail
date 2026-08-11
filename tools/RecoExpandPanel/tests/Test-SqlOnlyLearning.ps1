$ErrorActionPreference = 'Stop'

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceDir = Split-Path -Parent $testDir

function Get-Section([string]$text, [string]$startMarker, [string]$endMarker) {
    $start = $text.IndexOf($startMarker, [StringComparison]::Ordinal)
    if ($start -lt 0) { throw "Missing start marker: $startMarker" }
    $end = $text.IndexOf($endMarker, $start + $startMarker.Length, [StringComparison]::Ordinal)
    if ($end -lt 0) { throw "Missing end marker: $endMarker" }
    return $text.Substring($start, $end - $start)
}

$smart = [IO.File]::ReadAllText((Join-Path $sourceDir 'SmartFillFeature.cs'))
$snapshot = Get-Section $smart 'private static SmartLearningSnapshot LoadSmartLearningSnapshot' 'private static bool ShouldWarnSmartLibraryPartitionMissing'
if ($snapshot.Contains('LoadMappingBoxRows(')) { throw 'SmartFill still reads local mapping-boxes when SQL is unavailable.' }
if ($snapshot.Contains('MergePendingLocalMappingsIntoSmartSnapshot(')) { throw 'SmartFill still overlays pending local mappings onto the SQL snapshot.' }
if ($snapshot.Contains('fallback to jsonl') -or $snapshot.Contains('jsonl fallback')) { throw 'SmartFill still advertises a local jsonl fallback.' }

$scopes = Get-Section $smart 'private static List<SmartLearningScope> LoadSmartLearningScopes' 'private static bool IsSmartClassifiedEntryCode'
if ($scopes.Contains('AgentCreateWorkConnection(mainForm)')) { throw 'SmartFill scope loading still clones the host project connection.' }
if (-not $scopes.Contains('GetOpenProjectConnection(mainForm)')) { throw 'SmartFill scope loading does not reuse the open host project connection.' }

$preview = $smart.Substring($smart.IndexOf('private static List<FillPreviewItem> BuildPreview_SmartFill', [StringComparison]::Ordinal))
if ($preview.Contains('AgentCreateWorkConnection(mainForm)')) { throw 'SmartFill preview still clones the host project connection.' }
if (-not $preview.Contains('GetOpenProjectConnection(mainForm)')) { throw 'SmartFill preview does not reuse the open host project connection.' }
if (-not $preview.Contains('return new List<FillPreviewItem>();')) { throw 'SmartFill preview does not fail closed when project context is unavailable.' }

$excelLink = [IO.File]::ReadAllText((Join-Path $sourceDir 'ExcelLinkFeature.cs'))
$store = Get-Section $excelLink 'private static void RecordMappingGroupsToLearningDb' 'private static LocalMappingSaveResult SaveMappingGroupsToLocalFile'
if ($store.Contains('SaveMappingGroupsToLocalFile(') -or $store.Contains('mapping-boxes.jsonl')) { throw 'Binding feedback still writes the local mapping store.' }
if (-not $store.Contains('RecordBindingEventsToLearningDb(source, groups)')) { throw 'Binding feedback is not routed to RecoLearning SQL.' }

$learning = [IO.File]::ReadAllText((Join-Path $sourceDir 'LearningDbFeature.cs'))
$record = Get-Section $learning 'private static bool RecordBindingEventsToLearningDb' 'private static void WriteBindingEventsToLearningDb'
foreach ($forbidden in @('TryAppendLearningDbOutbox(', 'TryAppendLearningDbDeadLetter(', 'TryReplayPendingLearningDbEvents(')) {
    if ($record.Contains($forbidden)) { throw "SQL-only binding still persists a local queue: $forbidden" }
}

$nameMatch = [IO.File]::ReadAllText((Join-Path $sourceDir 'TemplateFillNameMatch.cs'))
$namePreview = Get-Section $nameMatch 'private static List<FillPreviewItem> BuildPreview_NameDriven' 'private sealed class ProjectQuota'
foreach ($forbidden in @('LoadMappingBoxRows(', 'BuildMappingBoxIndex(', 'LookupMappingBox(')) {
    if ($namePreview.Contains($forbidden)) { throw "Name-driven preview still pairs from local learning: $forbidden" }
}

$repoRoot = Split-Path -Parent (Split-Path -Parent $sourceDir)
$build = [IO.File]::ReadAllText((Join-Path $repoRoot 'RecoQuotaRecommend\build.ps1'))
if ($build.Contains('/target:exe /out:$importerOut')) { throw 'Build still emits the legacy local learning importer.' }
foreach ($requiredExclusion in @('mapping-boxes.jsonl', 'learning.jsonl', 'learning.csv', 'learning-summary.txt')) {
    if (-not $build.Contains($requiredExclusion)) { throw "Build does not exclude local learning artifact: $requiredExclusion" }
}

$release = [IO.File]::ReadAllText((Join-Path $repoRoot 'tools\BuildColleaguePluginRelease.ps1'))
foreach ($forbidden in @('mapping-boxes.jsonl', 'learning.jsonl', '.mapping-boxes.empty')) {
    if ($release.Contains($forbidden)) { throw "Release still packages local learning artifact: $forbidden" }
}

Write-Host 'PASS: binding writes and recommendation pairing are SQL-only; local learning is not active.' -ForegroundColor Green
