$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..'))
$dll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL)) {
    $env:RECO_EXPAND_DLL
} else {
    Join-Path $repoRoot 'RecoQuotaRecommend\bin\RecoExpandPanel.dll'
}
if (-not (Test-Path -LiteralPath $dll)) { throw "Missing DLL: $dll" }

$dllDir = Split-Path -Parent $dll
foreach ($dependency in @('NPOI.dll', 'NPOI.OpenXmlFormats.dll', 'NPOI.OpenXml4Net.dll', 'NPOI.OOXML.dll', 'ICSharpCode.SharpZipLib.dll')) {
    $dependencyPath = Join-Path $dllDir $dependency
    if (Test-Path -LiteralPath $dependencyPath) { [void][System.Reflection.Assembly]::LoadFrom($dependencyPath) }
}
Add-Type -AssemblyName System.Windows.Forms

$assembly = [System.Reflection.Assembly]::LoadFrom($dll)
$type = $assembly.GetType('RecoNet.FormPanel', $true)
$panelType = $type.GetNestedType('TemplateFillPanel', [System.Reflection.BindingFlags]'Public,NonPublic')
$scopeType = $type.GetNestedType('SmartLearningScope', [System.Reflection.BindingFlags]'Public,NonPublic')
$flags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$constructor = $panelType.GetConstructor(
    [System.Reflection.BindingFlags]'Public,NonPublic,Instance',
    $null,
    [Type[]]@([System.Windows.Forms.Form], [bool]),
    $null)
if ($null -eq $constructor) { throw 'Missing smart-only TemplateFillPanel constructor' }

$owner = New-Object System.Windows.Forms.Form
$arguments = New-Object 'object[]' 2
$arguments[0] = $owner.PSObject.BaseObject
$arguments[1] = $true
$panel = $null
try {
    $panel = $constructor.Invoke($arguments)
    $button = $panelType.GetField('btnSmartLearningScope', $flags).GetValue($panel)
    $workbook = $panelType.GetField('cmbTargetWorkbook', $flags).GetValue($panel)
    $tree = $panelType.GetField('smartLearningScopeTree', $flags).GetValue($panel)
    if (-not $panel.Controls.Contains($button) -or $panel.Controls.Contains($workbook)) {
        throw '推荐定额窗口没有以推荐学习库替换目标Excel控件'
    }
    $learningLabel = @($panel.Controls | Where-Object { $_ -is [System.Windows.Forms.Label] -and $_.Text -eq '推荐学习库' })[0]
    if ($null -eq $learningLabel -or $learningLabel.Width -lt 70 -or $button.Width -gt 170) {
        throw '推荐学习库表头仍可能换行或选择框未缩窄'
    }
    if (-not $tree.ShowPlusMinus -or $tree.Nodes.Count -eq 0 -or $tree.Nodes[0].Nodes.Count -ne 0) {
        throw '推荐学习库没有把全部学习库与专业目录改为同级节点'
    }

    $format = $panelType.GetMethod('BuildSmartLearningScopeText', $flags)
    function Format-Scope([string]$Code, [string]$Name) {
        $scope = [Activator]::CreateInstance($scopeType, $true).PSObject.BaseObject
        $scopeType.GetField('Kind', $flags).SetValue($scope, 'Entry')
        $scopeType.GetField('EntryCode', $flags).SetValue($scope, $Code)
        $scopeType.GetField('DisplayName', $flags).SetValue($scope, $Name)
        return [string]$format.Invoke($null, @($scope))
    }
    if ((Format-Scope '03' '桥涵') -ne '三、桥涵') { throw '专业节点显示格式错误' }
    if ((Format-Scope '0305' '特大桥') -ne '05.特大桥') { throw '分部节点显示格式错误' }
    if ((Format-Scope '0309-01-03-01' '主体工程') -ne '0309-01-03-01 主体工程') {
        throw '完整条目叶子显示格式错误'
    }

    $snapshotType = $type.GetNestedType('SmartLearningSnapshot', [System.Reflection.BindingFlags]'Public,NonPublic')
    $entryType = $type.GetNestedType('SmartMapEntry', [System.Reflection.BindingFlags]'Public,NonPublic')
    $filter = $type.GetMethod('FilterSmartHitsByScope', $flags)
    if ($null -eq $snapshotType -or $null -eq $entryType -or $null -eq $filter) {
        throw '缺少推荐学习库范围过滤行为入口'
    }

    $snapshot = [Activator]::CreateInstance($snapshotType, $true).PSObject.BaseObject
    $snapshotType.GetField('Method', $flags).SetValue($snapshot, '2020')
    $scopeMap = $snapshotType.GetField('ScopeEntriesByBox', $flags).GetValue($snapshot)
    $persistedCodes = New-Object 'System.Collections.Generic.HashSet[string]'
    [void]$persistedCodes.Add('03')
    $scopeMap.Add('persisted-box', $persistedCodes)

    function New-SmartEntry([string]$BoxId, [string]$ContextCode) {
        $entry = [Activator]::CreateInstance($entryType, $true).PSObject.BaseObject
        $entryType.GetField('BoxId', $flags).SetValue($entry, $BoxId)
        $contexts = $entryType.GetField('LocalContextKeys', $flags).GetValue($entry)
        [void]$contexts.Add("2020`n$ContextCode")
        return $entry
    }
    $persistedEntry = New-SmartEntry 'persisted-box' '04'
    $localOnlyEntry = New-SmartEntry 'local-only-box' '05'
    $listType = [System.Collections.Generic.List``1].MakeGenericType($entryType)
    $hits = [Activator]::CreateInstance($listType).PSObject.BaseObject
    [void]$hits.Add($persistedEntry)
    [void]$hits.Add($localOnlyEntry)

    function New-Scope([string]$Kind, [string]$EntryCode) {
        $scope = [Activator]::CreateInstance($scopeType, $true).PSObject.BaseObject
        $scopeType.GetField('Kind', $flags).SetValue($scope, $Kind)
        $scopeType.GetField('EntryCode', $flags).SetValue($scope, $EntryCode)
        return $scope
    }
    function Invoke-ScopeFilter($Scope) {
        $invokeArgs = New-Object 'object[]' 3
        $invokeArgs[0] = $snapshot
        $invokeArgs[1] = $hits
        $invokeArgs[2] = $Scope
        return $filter.Invoke($null, $invokeArgs)
    }
    function Get-FilteredBoxIds($Result) {
        $ids = @()
        foreach ($entry in $Result) { $ids += [string]$entryType.GetField('BoxId', $flags).GetValue($entry) }
        return $ids
    }

    $entryScope = New-Scope 'Entry' '05'
    [string[]]$firstEntryIds = @(Get-FilteredBoxIds (Invoke-ScopeFilter $entryScope))
    [string[]]$secondEntryIds = @(Get-FilteredBoxIds (Invoke-ScopeFilter $entryScope))
    [string[]]$unclassifiedIds = @(Get-FilteredBoxIds (Invoke-ScopeFilter (New-Scope 'Unclassified' '')))
    if ($firstEntryIds.Count -ne 1 -or $firstEntryIds[0] -ne 'local-only-box' -or
        $secondEntryIds.Count -ne 1 -or $secondEntryIds[0] -ne 'local-only-box') {
        throw '本机上下文的范围过滤结果在连续调用后发生了变化'
    }
    if ($unclassifiedIds.Count -ne 1 -or $unclassifiedIds[0] -ne 'local-only-box') {
        throw '只有本机上下文、无 EngineeringTemplate 证据的组件未保持未归类'
    }
    if ($scopeMap.Count -ne 1 -or -not $scopeMap.ContainsKey('persisted-box') -or
        $persistedCodes.Count -ne 1 -or -not $persistedCodes.Contains('03') -or
        $persistedCodes.Contains('04') -or $scopeMap.ContainsKey('local-only-box')) {
        throw '范围过滤就地修改了 SmartLearningSnapshot'
    }
    Write-Host 'Test-SmartFillLearningScopeUi: PASS'
}
finally {
    if ($null -ne $panel) { $panel.Dispose() }
    $owner.Dispose()
}
