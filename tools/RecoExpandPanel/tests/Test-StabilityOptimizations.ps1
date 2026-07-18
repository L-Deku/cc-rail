$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..')).Path

function Read-Source([string]$relativePath) {
    $path = Join-Path $repoRoot $relativePath
    return [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
}

function Assert-Contains([string]$text, [string]$expected, [string]$message) {
    if (-not $text.Contains($expected)) {
        throw $message
    }
}

function Assert-NotContains([string]$text, [string]$unexpected, [string]$message) {
    if ($text.Contains($unexpected)) {
        throw $message
    }
}

$mappingStore = Read-Source 'RecoQuotaRecommend\MappingStore.cs'
$quotaPanel = Read-Source 'RecoQuotaRecommend\QuotaRecommendPanel.cs'
$searchIndex = Read-Source 'RecoQuotaRecommend\SearchIndexStore.cs'
$chapterStore = Read-Source 'RecoQuotaRecommend\ChapterLibraryStore.cs'
$referencePool = Read-Source 'RecoQuotaRecommend\RecoReferenceQuotaPoolFeature.cs'
$excelLink = Read-Source 'tools\RecoExpandPanel\ExcelLinkFeature.cs'
$formPanel = Read-Source 'tools\RecoExpandPanel\FormPanel.cs'
$agentExecutor = Read-Source 'tools\RecoExpandPanel\AgentExecutor.cs'
$chapterTool = Read-Source 'tools\ChapterQuotaLibrary\ChapterQuotaLibrary.cs'

Assert-Contains $mappingStore 'private static bool TryWithMappingBoxesLock(Action action, int timeoutMilliseconds)' 'MappingStore has no testable lock-timeout guard.'
Assert-Contains $mappingStore 'if (!TryWithMappingBoxesLock(delegate' 'MappingStore still writes after a mutex timeout.'
Assert-Contains $mappingStore 'File.Replace(temp, filePath, backup, true);' 'MappingStore does not atomically replace the mapping file with a backup.'
Assert-NotContains $mappingStore "File.Delete(path);`r`n            }`r`n            File.Move(temp, path);" 'MappingStore still deletes the live file before moving the replacement.'

Assert-Contains $excelLink 'private static bool TryWithMappingBoxesLock(Action action, int timeoutMilliseconds)' 'RecoExpandPanel mapping feedback has no lock-timeout guard.'
Assert-Contains $excelLink 'WriteAllLinesAtomic(path, rows.Select(ToFlatJson).ToArray(), Encoding.UTF8);' 'RecoExpandPanel mapping feedback is not written atomically.'
Assert-Contains $chapterTool 'throw new TimeoutException("Timed out waiting for mapping-boxes lock.");' 'ChapterQuotaLibrary still continues after a mapping lock timeout.'

Assert-Contains $quotaPanel 'private const long MaxLogBytes = 5L * 1024L * 1024L;' 'RecoQuotaRecommend log rotation threshold is missing.'
Assert-Contains $quotaPanel 'private const int LogBackupCount = 3;' 'RecoQuotaRecommend log backup count is missing.'
Assert-Contains $formPanel 'private const long MaxLogBytes = 5L * 1024L * 1024L;' 'RecoExpandPanel log rotation threshold is missing.'
Assert-Contains $formPanel 'private const int LogBackupCount = 3;' 'RecoExpandPanel log backup count is missing.'

Assert-Contains $searchIndex 'private static readonly object CacheLock = new object();' 'SearchIndexStore shared snapshot lock is missing.'
Assert-Contains $searchIndex 'File.ReadLines(quotaPath, Encoding.UTF8)' 'SearchIndexStore still buffers the entire quota index as string lines.'
Assert-Contains $searchIndex 'File.ReadLines(materialPath, Encoding.UTF8)' 'SearchIndexStore still buffers the entire material index as string lines.'
Assert-Contains $chapterStore 'private static readonly Dictionary<string, ChapterLibraryCacheEntry> StoreCache' 'ChapterLibraryStore shared snapshot cache is missing.'
Assert-Contains $chapterStore 'PoolKey(methodNo, entryCode)' 'Chapter pool method_no and entry_code isolation was removed.'
Assert-Contains $referencePool 'private static readonly object ReferenceDataCacheLock = new object();' 'Reference quota rich-data cache is missing.'

Assert-Contains $formPanel 'private const int InstalledPollIntervalMs = 15000;' 'Installed menu polling was not reduced to 15 seconds.'
Assert-Contains $formPanel 'installTimer.Interval = InstalledPollIntervalMs;' 'The installer timer never switches to the steady-state interval.'
Assert-Contains $formPanel 'mainForm.FormClosed += InstalledMainFormClosed;' 'The installer timer is not tied to the host form lifetime.'
Assert-Contains $agentExecutor 'private static readonly HashSet<string> AgentCloneConnectionFailures' 'Known clone-connection failures are not cached.'
Assert-Contains $agentExecutor 'AgentCloneConnectionFailures.Add(connectionKey);' 'A successful fallback does not cache the failed clone identity.'
$fallbackOpen = $agentExecutor.IndexOf('fallbackConn.Open();', [StringComparison]::Ordinal)
$cacheFailure = $agentExecutor.IndexOf('AgentCloneConnectionFailures.Add(connectionKey);', [StringComparison]::Ordinal)
if ($fallbackOpen -lt 0 -or $cacheFailure -lt $fallbackOpen) {
    throw 'Clone failure is cached before the fallback connection succeeds.'
}

Write-Host 'PASS: storage safety, shared snapshots, and steady-state retry guards are present.'
