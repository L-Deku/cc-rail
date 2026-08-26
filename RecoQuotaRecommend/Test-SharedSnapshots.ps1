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
        '{"method":"2020","method_no":"30号文","entry_code":"0101-98","entry_name":"测试条目2020","entry_type":"小计"}',
        '{"method":"2020","method_no":"30号文","entry_code":"0101-99","entry_name":"测试条目2020","entry_type":"小计"}',
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

    # 当前条目从父级池继承两条；物化后可批量增加，硬删除只改当前条目且不产生 deleted=1。
    $resolveEdit = $chapterType.GetMethods($flags) | Where-Object { $_.Name -eq 'ResolveScopeForUserEdit' -and $_.GetParameters().Count -eq 3 } | Select-Object -First 1
    $replacePool = $chapterType.GetMethod('ReplaceUserQuotaPool', $flags)
    $mutationType = $assembly.GetType('RecoQuotaRecommend.ReferencePoolMutationItem', $true)
    $mutationListType = [System.Collections.Generic.List``1].MakeGenericType($mutationType)

    function New-MutationItem([string]$code, [string]$name) {
        $item = [Activator]::CreateInstance($mutationType, $true)
        $item.Kind = 'quota'
        $item.Code = $code
        $item.Name = $name
        $item.Unit = 'm'
        return $item
    }

    function Invoke-ReplacePool([string]$entryCode, [object[]]$items) {
        $typedItems = [Activator]::CreateInstance($mutationListType)
        foreach ($item in $items) {
            $typedItems.Add($item.PSObject.BaseObject)
        }
        $replaceArgs = New-Object object[] 5
        $replaceArgs[0] = '30号文'
        $replaceArgs[1] = $entryCode
        $replaceArgs[2] = '测试条目2020'
        $replaceArgs[3] = $typedItems.PSObject.BaseObject
        $replaceArgs[4] = ''
        $ok = $replacePool.Invoke($chapter1, $replaceArgs)
        if (-not $ok) {
            throw ('ReplaceUserQuotaPool failed: ' + [string]$replaceArgs[4])
        }
    }

    $childBeforeArgs = New-Object object[] 3
    $childBeforeArgs[0] = '30号文'
    $childBeforeArgs[1] = '0101-99'
    $childBeforeArgs[2] = '测试条目2020'
    $childBefore = $resolveEdit.Invoke($chapter1, $childBeforeArgs)
    if ($childBefore.MatchedEntryCode -ne '0101' -or $childBefore.PoolKeys.Count -ne 2) {
        throw 'Child entry did not inherit the two-row source pool before editing.'
    }

    Invoke-ReplacePool '0101-99' @(
        (New-MutationItem 'QA-1' '甲定额'),
        (New-MutationItem 'QC-1' '丙定额'),
        (New-MutationItem 'QB-1' '乙定额')
    )
    $chapterAfterAdd = Invoke-ZeroArg $loadChapter
    $childAfterAdd = $resolveEdit.Invoke($chapterAfterAdd, $childBeforeArgs)
    if ($childAfterAdd.MatchedEntryCode -ne '0101-99' -or $childAfterAdd.PoolKeys.Count -ne 3 -or -not $childAfterAdd.IsExplicitPool) {
        throw 'Materialized child pool did not retain inherited rows plus the added row.'
    }

    Invoke-ReplacePool '0101-99' @(
        (New-MutationItem 'QA-1' '甲定额'),
        (New-MutationItem 'QC-1' '丙定额')
    )
    $libraryLines = [System.IO.File]::ReadAllLines((Join-Path $dataDir 'chapter-quota-library.jsonl'), [System.Text.Encoding]::UTF8)
    $childLines = @($libraryLines | Where-Object { $_ -like '*"method_no":"30号文"*' -and $_ -like '*"entry_code":"0101-99"*' })
    if ($childLines.Count -ne 3 -or ($childLines -join "`n") -match '"deleted":"1"' -or ($childLines -join "`n") -match '"quota_code":"QB-1"') {
        throw 'Hard delete left a tombstone or retained the removed quota row.'
    }
    $sourceAfterDelete = Invoke-Scope $resolveScope $chapterAfterAdd '30号文'
    if ($sourceAfterDelete.PoolKeys.Count -ne 2 -or -not $sourceAfterDelete.PoolKeys.Contains('quota:QA-1') -or -not $sourceAfterDelete.PoolKeys.Contains('quota:QC-1')) {
        throw 'Editing the child pool changed the source pool.'
    }
    $siblingArgs = New-Object object[] 3
    $siblingArgs[0] = '30号文'
    $siblingArgs[1] = '0101-98'
    $siblingArgs[2] = '测试条目2020'
    $siblingAfterDelete = $resolveEdit.Invoke($chapterAfterAdd, $siblingArgs)
    if ($siblingAfterDelete.MatchedEntryCode -ne '0101' -or $siblingAfterDelete.PoolKeys.Count -ne 2) {
        throw 'Editing the child pool changed a sibling entry that shares the source pool.'
    }

    Invoke-ReplacePool '0101-99' @()
    $chapterAfterEmpty = Invoke-ZeroArg $loadChapter
    $childAfterEmpty = $resolveEdit.Invoke($chapterAfterEmpty, $childBeforeArgs)
    if ($childAfterEmpty.MatchedEntryCode -ne '0101-99' -or -not $childAfterEmpty.IsExplicitPool -or -not $childAfterEmpty.Strict -or $childAfterEmpty.PoolKeys.Count -ne 0) {
        throw 'An explicitly empty current-entry pool fell back to the inherited source pool.'
    }
    $richPoolAfterEmpty = $loadPool.Invoke($null, $poolArgs)
    if (-not $richPoolAfterEmpty.ContainsKey('30号文|0101-99') -or $richPoolAfterEmpty['30号文|0101-99'].Count -ne 0) {
        throw 'Reference pool display cache did not preserve the explicit empty current-entry pool.'
    }
    $libraryLinesAfterEmpty = [System.IO.File]::ReadAllLines((Join-Path $dataDir 'chapter-quota-library.jsonl'), [System.Text.Encoding]::UTF8)
    $emptyChildLines = @($libraryLinesAfterEmpty | Where-Object { $_ -like '*"method_no":"30号文"*' -and $_ -like '*"entry_code":"0101-99"*' })
    if ($emptyChildLines.Count -ne 1 -or $emptyChildLines[0] -notlike '*"record_type":"entry_quota_pool"*' -or $emptyChildLines[0] -like '*"deleted":"1"*') {
        throw 'Delete-all did not leave exactly one explicit empty-pool marker.'
    }
    if (@(Get-ChildItem -LiteralPath $dataDir -File | Where-Object { $_.Name -like 'chapter-quota-library.jsonl.*.tmp' -or $_.Name -like 'chapter-quota-library.jsonl.*.old' }).Count -ne 0) {
        throw 'Atomic hard delete left temporary or old snapshot files behind.'
    }

    # 真实 WinForms 控件冒烟：安装后上下两个表都支持多选，释放后恢复宿主原设置。
    Add-Type -AssemblyName System.Windows.Forms
    $runtimeType = $referenceType.GetNestedType('Runtime', [System.Reflection.BindingFlags]'NonPublic')
    $runtimeCtor = $runtimeType.GetConstructors($flags) | Select-Object -First 1
    $testForm = New-Object System.Windows.Forms.Form
    $hostGrid = New-Object System.Windows.Forms.DataGridView
    $hostGrid.MultiSelect = $false
    [void]$hostGrid.Columns.Add('chapter_sequence', '条目序号')
    [void]$hostGrid.Columns.Add('quota_code', '定额编号')
    [void]$hostGrid.Rows.Add('100', 'QA-1')
    [void]$hostGrid.Rows.Add('100', 'QC-1')
    [void]$hostGrid.Rows.Add('101', 'QB-1')
    $hostGrid.CurrentCell = $hostGrid.Rows[0].Cells[1]
    $tabControl = New-Object System.Windows.Forms.TabControl
    $referencePage = New-Object System.Windows.Forms.TabPage
    $referencePage.Text = '参考定额'
    [void]$tabControl.TabPages.Add($referencePage)
    $testForm.Controls.Add($tabControl)
    $runtimeArgs = New-Object object[] 2
    $runtimeArgs[0] = $testForm.PSObject.BaseObject
    $runtimeArgs[1] = $hostGrid.PSObject.BaseObject
    $runtime = $runtimeCtor.Invoke($runtimeArgs)
    $installRuntime = $runtimeType.GetMethod('Install', $flags)
    if (-not $installRuntime.Invoke($runtime, (New-Object object[] 0))) {
        throw 'Reference quota runtime did not install in the WinForms multi-select smoke test.'
    }
    $refGridField = $runtimeType.GetField('refGrid', $flags)
    $installedRefGrid = $refGridField.GetValue($runtime)
    if (-not $hostGrid.MultiSelect -or -not $installedRefGrid.MultiSelect) {
        throw 'Add/delete grids did not enable multi-select.'
    }
    $doubleBufferedProperty = [System.Windows.Forms.DataGridView].GetProperty('DoubleBuffered', [System.Reflection.BindingFlags]'Instance,NonPublic')
    if (-not $doubleBufferedProperty.GetValue($installedRefGrid, $null)) {
        throw 'Reference quota grid did not enable double buffering.'
    }
    $timerField = $runtimeType.GetField('timer', $flags)
    $runtimeTimer = $timerField.GetValue($runtime)
    $runtimeTimer.Stop()
    $runtimeType.GetField('suppressRefresh', $flags).SetValue($runtime, $true)
    $runtimeType.GetMethod('ScheduleRefresh', $flags).Invoke($runtime, (New-Object object[] 0)) | Out-Null
    if ($runtimeTimer.Enabled) {
        throw 'Suppressed native apply still scheduled a re-entrant reference-pool refresh.'
    }
    $runtimeType.GetField('suppressRefresh', $flags).SetValue($runtime, $false)
    $runtimeType.GetField('lastGridChapterSequence', $flags).SetValue($runtime, '100')
    $runtimeTimer.Stop()
    $hostGrid.CurrentCell = $hostGrid.Rows[1].Cells[1]
    if ($runtimeTimer.Enabled) {
        throw 'Moving between quota rows in the same chapter scheduled a redundant refresh.'
    }
    $hostGrid.CurrentCell = $hostGrid.Rows[2].Cells[1]
    if (-not $runtimeTimer.Enabled) {
        throw 'Moving to a different chapter sequence did not schedule a reference-pool refresh.'
    }
    $runtimeTimer.Stop()
    $installedRefGrid.Rows.Add('QA-1') | Out-Null
    $installedRefGrid.Rows.Add('QC-1') | Out-Null
    $installedRefGrid.Rows.Add('QB-1') | Out-Null
    $installedRefGrid.ClearSelection()
    $installedRefGrid.Rows[0].Selected = $true
    $installedRefGrid.Rows[2].Selected = $true
    $selectedRowsMethod = $runtimeType.GetMethod('SelectedRowsOrCurrent', $flags)
    $selectedArgs = New-Object object[] 1
    $selectedArgs[0] = $installedRefGrid.PSObject.BaseObject
    [object[]]$selectedRows = @($selectedRowsMethod.Invoke($null, $selectedArgs))
    if ($selectedRows.Count -ne 2 -or $selectedRows[0].Index -ne 0 -or $selectedRows[1].Index -ne 2) {
        throw 'Ctrl/Shift multi-row selection was not collected deterministically.'
    }
    $disposeRuntime = $runtimeType.GetMethod('Dispose', $flags)
    [void]$disposeRuntime.Invoke($runtime, (New-Object object[] 0))
    if ($hostGrid.MultiSelect) {
        throw 'Runtime disposal did not restore the host grid multi-select setting.'
    }
    $testForm.Dispose()
    $hostGrid.Dispose()
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

Write-Host 'PASS: shared snapshots, method_no isolation, materialized child pools, multi-row snapshots, and hard delete.'
