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
$autoMatch = Read-Source 'tools\RecoExpandPanel\AutoMatchFeature.cs'
$formPanel = Read-Source 'tools\RecoExpandPanel\FormPanel.cs'
$agentExecutor = Read-Source 'tools\RecoExpandPanel\AgentExecutor.cs'
$chapterTool = Read-Source 'tools\ChapterQuotaLibrary\ChapterQuotaLibrary.cs'
$localMappingStore = Read-Source 'RecoShared\LocalMappingFileStore.cs'

Assert-Contains $mappingStore 'LocalMappingFileStore.Save(path, softwarePartition, sourceOperation, 5000' 'MappingStore does not use the shared mapping-file writer.'
Assert-Contains $excelLink 'LocalMappingFileStore.Save(path, partition, sourceOperation, 5000' 'RecoExpandPanel mapping feedback does not use the shared mapping-file writer.'
Assert-NotContains $mappingStore 'TryWithMappingBoxesLock' 'MappingStore reintroduced a private competing mapping-file lock path.'
Assert-NotContains $excelLink 'private static bool TryWithMappingBoxesLock' 'RecoExpandPanel reintroduced a private competing mapping-file lock path.'
Assert-Contains $localMappingStore 'internal const string MutexName = "RecoQuotaData.mapping-boxes.lock";' 'Shared mapping-file mutex name is missing.'
Assert-Contains $localMappingStore 'acquired = mutex.WaitOne(timeoutMilliseconds);' 'Shared mapping-file lock-timeout guard is missing.'
Assert-Contains $localMappingStore 'File.Replace(temp, path, null, true);' 'Shared mapping-file writer does not atomically replace the live file.'
Assert-NotContains $localMappingStore 'File.Delete(path)' 'Shared mapping-file writer deletes the live file before replacement.'
Assert-Contains $chapterTool 'throw new TimeoutException("TagMappingBoxes' 'ChapterQuotaLibrary still continues after a mapping lock timeout.'

Assert-Contains $excelLink 'private static WeakReference CachedSpreadsheetApplication' 'Excel/WPS application cache still keeps a strong global COM reference.'
Assert-NotContains $excelLink 'private static object CachedSpreadsheetApplication;' 'Strong Excel/WPS application cache was reintroduced.'
$collectStart = $excelLink.IndexOf('private static void CollectExcelChildWindows', [StringComparison]::Ordinal)
$collectEnd = $excelLink.IndexOf('private static string GetWindowClassName', $collectStart, [StringComparison]::Ordinal)
if ($collectStart -lt 0 -or $collectEnd -le $collectStart) {
    throw 'Could not locate the Excel child-window collector.'
}
$collectBody = $excelLink.Substring($collectStart, $collectEnd - $collectStart)
if (($collectBody.Split(@('CollectExcelChildWindows('), [StringSplitOptions]::None).Count - 1) -ne 1) {
    throw 'CollectExcelChildWindows still recursively re-enumerates descendants already returned by EnumChildWindows.'
}
Assert-Contains $autoMatch 'HashSet<string> scannedApplications = new HashSet<string>(StringComparer.Ordinal);' 'Open-workbook discovery does not deduplicate Excel/WPS application instances.'
Assert-Contains $autoMatch 'TryMarkSpreadsheetApplicationScanned(activeApplication, scannedApplications)' 'The active Excel/WPS application is not guarded against duplicate scans.'
Assert-Contains $excelLink 'String.IsNullOrWhiteSpace(link.QuantityName)' 'Opening the Excel link panel still refreshes names for every stored binding.'
Assert-Contains $excelLink 'private const int ExcelLinkPollIntervalMs = 5000;' 'Excel link background polling is still too aggressive.'
Assert-Contains $excelLink 'timer.Enabled = knownWriteTimes.Count > 0;' 'Excel link polling does not stop when the project has no active bindings.'
Assert-Contains $excelLink '.Distinct(StringComparer.OrdinalIgnoreCase)' 'Excel link polling still checks the same workbook path once per binding.'

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
Assert-Contains $agentExecutor 'private static SqlConnection GetOpenProjectConnection(Form mainForm)' 'Read-only UI paths cannot reuse the open host project connection.'
Assert-Contains $agentExecutor 'EnsureOpen(conn);' 'Closed host project connections are not safely reopened on the same connection.'
if ($agentExecutor -match 'AgentDbPassword|fallbackConn\.Open\(\)|User ID=|Password=') {
    throw 'An insecure project-connection fallback was reintroduced.'
}

Write-Host 'PASS: storage safety, shared snapshots, and steady-state retry guards are present.'
