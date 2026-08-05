$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if ([String]::IsNullOrWhiteSpace($env:RECO_QUOTA_DLL)) {
    throw 'Set RECO_QUOTA_DLL to a source-built RecoQuotaRecommend.dll.'
}

$work = Join-Path $env:TEMP ('reco-snapshot-test-' + [Guid]::NewGuid().ToString('N'))
$dataDir = Join-Path $work 'RecoQuotaData'
[System.IO.Directory]::CreateDirectory($dataDir) | Out-Null

function Write-Utf8Lines([string]$path, [string[]]$lines) {
    [System.IO.File]::WriteAllLines($path, $lines, [System.Text.Encoding]::UTF8)
}

function Invoke-ZeroArg($method) {
    return $method.Invoke($null, (New-Object object[] 0))
}

function Invoke-Scope($method, $store, [string]$methodNo) {
    $args = New-Object object[] 3
    $args[0] = $methodNo
    $args[1] = '0101'
    $args[2] = '测试条目'
    return $method.Invoke($store, $args)
}

try {
    $testDll = Join-Path $work 'RecoQuotaRecommend.dll'
    [System.IO.File]::Copy($env:RECO_QUOTA_DLL, $testDll, $true)
    Write-Utf8Lines (Join-Path $dataDir 'quota-index.jsonl') @(
        '{"quota_code":"QA-1","quota_name":"甲定额","quota_unit":"m","book_category":"预算定额","search_text":"甲定额"}',
        '{"quota_code":"QB-1","quota_name":"乙定额","quota_unit":"m","book_category":"预算定额","search_text":"乙定额"}',
        '{"quota_code":"QD-1","quota_name":"丁定额","quota_unit":"m","book_category":"预算定额","search_text":"丁定额"}'
    )
    Write-Utf8Lines (Join-Path $dataDir 'material-index.jsonl') @()
    Write-Utf8Lines (Join-Path $dataDir 'chapter-entries.jsonl') @(
        '{"method":"2020","method_no":"30号文","entry_code":"0101","entry_name":"测试条目2020","entry_type":"小计"}',
        '{"method":"2020","method_no":"101号文","entry_code":"0101","entry_name":"测试条目101","entry_type":"小计"}',
        '{"method":"2024","method_no":"TB 10801—2024","entry_code":"0101","entry_name":"测试条目2024","entry_type":"小计"}'
    )
    Write-Utf8Lines (Join-Path $dataDir 'chapter-quota-library.jsonl') @(
        '{"method":"2020","method_no":"30号文","entry_code":"0101","entry_name":"测试条目","target_kind":"quota","quota_code":"QA-1","quota_name":"甲定额","quota_unit":"m"}',
        '{"method":"2020","method_no":"101号文","entry_code":"0101","entry_name":"测试条目101","target_kind":"quota","quota_code":"QB-1","quota_name":"乙定额","quota_unit":"m"}',
        '{"method":"2024","method_no":"TB 10801—2024","entry_code":"0101","entry_name":"测试条目2024","target_kind":"quota","quota_code":"QD-1","quota_name":"丁定额","quota_unit":"m"}'
    )

    $assembly = [System.Reflection.Assembly]::LoadFrom($testDll)
    $flags = [System.Reflection.BindingFlags]'NonPublic,Static,Instance,Public'

    $searchType = $assembly.GetType('RecoQuotaRecommend.SearchIndexStore', $true)
    $loadSearch = $searchType.GetMethod('LoadOrBuild', $flags)
    $search1 = Invoke-ZeroArg $loadSearch
    $search2 = Invoke-ZeroArg $loadSearch
    if (-not [Object]::ReferenceEquals($search1, $search2)) {
        throw 'SearchIndexStore did not reuse an unchanged snapshot.'
    }
    [System.IO.File]::AppendAllText((Join-Path $dataDir 'quota-index.jsonl'), '{"quota_code":"QC-1","quota_name":"丙定额","quota_unit":"m","book_category":"预算定额","search_text":"丙定额"}' + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
    $search3 = Invoke-ZeroArg $loadSearch
    if ([Object]::ReferenceEquals($search1, $search3) -or $search3.QuotaCount -ne 4) {
        throw 'SearchIndexStore did not reload after the quota index changed.'
    }

    $chapterType = $assembly.GetType('RecoQuotaRecommend.ChapterLibraryStore', $true)
    $loadChapter = $chapterType.GetMethod('Load', $flags)
    $chapter1 = Invoke-ZeroArg $loadChapter
    $chapter2 = Invoke-ZeroArg $loadChapter
    if (-not [Object]::ReferenceEquals($chapter1, $chapter2)) {
        throw 'ChapterLibraryStore did not reuse an unchanged snapshot.'
    }
    $resolveScope = $chapterType.GetMethods($flags) | Where-Object { $_.Name -eq 'ResolveScope' -and $_.GetParameters().Count -eq 3 } | Select-Object -First 1
    $scope30 = Invoke-Scope $resolveScope $chapter1 '30号文'
    $scope101 = Invoke-Scope $resolveScope $chapter1 '101号文'
    $scope2024 = Invoke-Scope $resolveScope $chapter1 'TB 10801—2024'
    if (-not $scope30.PoolKeys.Contains('quota:QA-1') -or $scope30.PoolKeys.Contains('quota:QB-1')) {
        throw '30号文章节池 mixed records from another method number.'
    }
    if (-not $scope101.PoolKeys.Contains('quota:QB-1') -or $scope101.PoolKeys.Contains('quota:QA-1')) {
        throw '101号文章节池 mixed records from another method number.'
    }
    if (-not $scope2024.PoolKeys.Contains('quota:QD-1') -or $scope2024.PoolKeys.Contains('quota:QA-1')) {
        throw '2024章节池 mixed records from another method number.'
    }
    if ($scope30.EntryName -ne '测试条目2020' -or $scope2024.EntryName -ne '测试条目2024') {
        throw 'Chapter entry metadata is not isolated by method_no|entry_code.'
    }

    $resolveMethod = $chapterType.GetMethod('ResolveMethodKeyForHost', $flags)
    [object[]]$host2020Args = @('D:\AI文件\铁路工程云计价系统网络版V1.0', 'RejjNet2020.exe', $true, $true)
    [object[]]$host2024Args = @('D:\AI文件\铁路工程云计价系统网络版V1.0', 'ReJJGSNet2024.exe', $true, $true)
    if ($resolveMethod.Invoke($null, $host2020Args) -ne '2020' -or $resolveMethod.Invoke($null, $host2024Args) -ne '2024') {
        throw 'Shared-directory host detection did not prefer the current process.'
    }

    $referenceType = $assembly.GetType('RecoQuotaRecommend.ReferenceQuotaPoolFeature', $true)
    $loadQuotaIndex = $referenceType.GetMethod('LoadQuotaIndex', $flags)
    $quotaIndex1 = Invoke-ZeroArg $loadQuotaIndex
    $quotaIndex2 = Invoke-ZeroArg $loadQuotaIndex
    if (-not [Object]::ReferenceEquals($quotaIndex1, $quotaIndex2)) {
        throw 'Reference quota index did not reuse an unchanged snapshot.'
    }
    $loadPool = $referenceType.GetMethod('LoadPool', $flags)
    $poolArgs = New-Object object[] 2
    $poolArgs[0] = '2024'
    $poolArgs[1] = $quotaIndex1
    $richPool1 = $loadPool.Invoke($null, $poolArgs)
    $richPool2 = $loadPool.Invoke($null, $poolArgs)
    if ([Object]::ReferenceEquals($richPool1, $richPool2)) {
        throw 'Reference quota pool returned its mutable cached dictionary directly.'
    }
    if (-not $richPool1.ContainsKey('30号文|0101') -or -not $richPool1.ContainsKey('101号文|0101') -or -not $richPool1.ContainsKey('TB10801-2024|0101')) {
        throw ('Reference quota pool is not isolated by method_no|entry_code. Keys=' + [String]::Join(',', [string[]]@($richPool1.Keys)))
    }

    [System.IO.File]::AppendAllText((Join-Path $dataDir 'chapter-quota-library.jsonl'), '{"method":"2020","method_no":"30号文","entry_code":"0101","entry_name":"测试条目","target_kind":"quota","quota_code":"QC-1","quota_name":"丙定额","quota_unit":"m"}' + [Environment]::NewLine, [System.Text.Encoding]::UTF8)
    $chapter3 = Invoke-ZeroArg $loadChapter
    if (-not [Object]::ReferenceEquals($chapter1, $chapter3)) {
        throw 'ChapterLibraryStore replaced the shared object instead of refreshing it in place.'
    }
    $scope30After = Invoke-Scope $resolveScope $chapter3 '30号文'
    if (-not $scope30After.PoolKeys.Contains('quota:QC-1')) {
        throw 'ChapterLibraryStore reload did not include the appended quota.'
    }
    $richPool3 = $loadPool.Invoke($null, $poolArgs)
    if ($richPool3['30号文|0101'].Count -ne 2 -or $richPool3['101号文|0101'].Count -ne 1 -or $richPool3['TB10801-2024|0101'].Count -ne 1) {
        throw 'Reference quota pool invalidation or method isolation failed after append.'
    }
}
finally {
    if ([System.IO.Directory]::Exists($work)) {
        try {
            [System.IO.Directory]::Delete($work, $true)
        }
        catch [System.UnauthorizedAccessException] {
            # 当前 PowerShell 进程会锁住已 LoadFrom 的测试 DLL；进程退出后由调用方清理。
            Write-Verbose $_.Exception.Message
        }
    }
}

Write-Host 'PASS: shared snapshots, file invalidation, mutable pool copies, and method_no isolation.'
