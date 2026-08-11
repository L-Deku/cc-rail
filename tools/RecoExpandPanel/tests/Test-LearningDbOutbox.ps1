$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
$learning = [IO.File]::ReadAllText((Join-Path $repoRoot 'tools\RecoExpandPanel\LearningDbFeature.cs'))
$smart = [IO.File]::ReadAllText((Join-Path $repoRoot 'tools\RecoExpandPanel\SmartFillFeature.cs'))

$start = $learning.IndexOf('private static bool RecordBindingEventsToLearningDb', [StringComparison]::Ordinal)
$end = $learning.IndexOf('private static void WriteBindingEventsToLearningDb', $start, [StringComparison]::Ordinal)
if ($start -lt 0 -or $end -le $start) { throw 'Cannot locate RecordBindingEventsToLearningDb.' }
$body = $learning.Substring($start, $end - $start)
foreach ($forbidden in @('TryAppendLearningDbOutbox(', 'TryAppendLearningDbDeadLetter(', 'TryReplayPendingLearningDbEvents(')) {
    if ($body.Contains($forbidden)) { throw "SQL-only binding still persists a local queue: $forbidden" }
}
if (-not $body.Contains('TryWriteLearningDbBatch(batch, out failureReason)')) { throw 'Binding no longer attempts the RecoLearning SQL transaction.' }
if (-not $body.Contains('return fullyWritten;')) { throw 'Binding does not return the SQL durability result.' }
if ($smart.Contains('LoadPendingLearningMappingKeys(') -or $smart.Contains('MergePendingLocalMappingsIntoSmartSnapshot(')) {
    throw 'Recommendation still consumes pending local learning.'
}

Write-Host 'PASS: local learning outbox/dead-letter is disabled on the active SQL-only path.' -ForegroundColor Green
