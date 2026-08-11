# 推荐定额(学习库智能铺量)标记测试:引擎、UI 挂接、菜单入口、老窗口删除。
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
$smartPath = Join-Path $repoRoot 'tools\RecoExpandPanel\SmartFillFeature.cs'
if (-not (Test-Path -LiteralPath $smartPath)) { throw '缺少 SmartFillFeature.cs' }
$smart = Get-Content -LiteralPath $smartPath -Raw -Encoding UTF8
$panel = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\TemplateFillPanel.cs') -Raw -Encoding UTF8
$excelLink = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\ExcelLinkFeature.cs') -Raw -Encoding UTF8
$learningDbSource = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\LearningDbFeature.cs') -Raw -Encoding UTF8
$rebuildAggregates = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoLearning\Rebuild-Aggregates.ps1') -Raw -Encoding UTF8
. (Join-Path $repoRoot 'tools\RecoLearning\Common.ps1')
$quotaPanel = Get-Content -LiteralPath (Join-Path $repoRoot 'RecoQuotaRecommend\QuotaRecommendPanel.cs') -Raw -Encoding UTF8
$oldDialog = Get-Content -LiteralPath (Join-Path $repoRoot 'RecoQuotaRecommend\RecommendDialog.cs') -Raw -Encoding UTF8
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

if ($smart -notmatch 'BuildPreview_SmartFill') { throw '缺少 BuildPreview_SmartFill' }
if ($smart -notmatch 'LoadSmartLearningSnapshot') { throw '缺少学习库快照加载 LoadSmartLearningSnapshot' }
if ($smart -match 'LoadMappingBoxRows\(' -or $smart -match '本地映射\(jsonl回退\)') { throw '推荐定额仍可能从本地学习配对' }
if ($smart -notmatch 'local learning is disabled') { throw 'SQL 失败没有明确关闭本地学习' }
if ($smart -notmatch 'IsLibraryQuota = true') { throw '缺少库内定额原生粘贴路径' }
if ($smart -notmatch 'IsNameDriven = true') { throw '推荐定额项必须 IsNameDriven=true,否则不回流学习库' }
if ($smart -notmatch 'TemplateName = "推荐定额"') { throw '预览项名称未改为推荐定额' }
if ($smart -notmatch 'EntryBySignatureQuota') { throw '缺少签名级条目证据' }
if ($smart -notmatch 'prefixVotes') { throw '缺少工程前缀投票' }
if ($smart -notmatch 'preferredPrefixes') { throw '缺少前缀过滤消歧' }
if ($smart -notmatch 'SmartLearningScope') { throw '缺少推荐学习库范围模型' }
if ($smart -notmatch 'LoadSmartLearningScopes') { throw '缺少推荐学习库目录加载' }
if ($smart -match '全库兜底') { throw '专业范围未命中后仍存在全库兜底' }
if ($smart -notmatch 'Status = "未匹配",\s*Selected = false') { throw '未匹配行仍可能默认勾选' }
if ($panel -notmatch '三·推荐定额') { throw 'cmbMode 缺少第三模式(推荐定额)' }
if ($panel -notmatch 'BuildPreview_SmartFill') { throw 'OnPreview 未挂接推荐定额分支' }
if ($panel -notmatch 'smartOnly') { throw '缺少 smartOnly 独立窗口模式' }
if ($panel -notmatch '推荐学习库' -or $panel -notmatch 'ToolStripDropDown' -or $panel -notmatch 'ShowPlusMinus = true') {
    throw '推荐定额窗口缺少折叠式推荐学习库目录'
}
if ($panel -notmatch 'TryResolveSmartActiveWorkbook' -or $panel -notmatch '当前活动工作簿尚未保存') {
    throw '推荐定额窗口没有改为读取当前活动且已保存的 Excel/WPS 工作簿'
}
if ($panel -notmatch 'CellPainting') { throw '缺少一量对多的工程量名合并绘制' }
if ($excelLink -notmatch '"推荐定额"') { throw '缺少推荐定额菜单入口' }
if ($excelLink -match '打开智能铺量面板') { throw '旧菜单名"打开智能铺量面板"未清除' }
if ($quotaPanel -match 'ShowRecommendDialog') { throw '老推荐定额窗口入口未删除' }
if ($oldDialog -match ': Form') { throw '老推荐定额窗口类未删除(仍继承 Form)' }
if ($learningDbSource -match 'Math\.Min\(100' -or $learningDbSource -match '>100 THEN 100' -or
    $excelLink -match 'Math\.Min\(100' -or $rebuildAggregates -match '\[Math\]::Min\(100') {
    throw '学习权重仍存在 100 上限'
}

$panelType = [System.Reflection.Assembly]::LoadFrom($dll).GetType('RecoNet.FormPanel', $true)
$flags = [System.Reflection.BindingFlags]'Public,NonPublic,Static'
$scoreMethod = $panelType.GetMethod('BuildSmartFuzzyScoresIfUnmatched', $flags)
if ($null -eq $scoreMethod) { throw '缺少可行为验证的模糊打分延后入口' }
$arguments = New-Object 'object[]' 3
$arguments[0] = $true
$arguments[1] = '精确签名命中'
$arguments[2] = $null
$scores = $scoreMethod.Invoke($null, $arguments)
if ($null -eq $scores -or $scores.Count -ne 0) { throw '精确签名命中后仍进入了模糊打分路径' }

$fuzzyPosition = $smart.IndexOf('BuildSmartFuzzyScoresIfUnmatched(matched', [StringComparison]::Ordinal)
if ($fuzzyPosition -lt 0) { throw '缺少范围内模糊打分入口' }

$allFlags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$candidateType = $panelType.GetNestedType('SmartMapCandidateScore', [System.Reflection.BindingFlags]'Public,NonPublic')
$entryType = $panelType.GetNestedType('SmartMapEntry', [System.Reflection.BindingFlags]'Public,NonPublic')
$targetType = $panelType.GetNestedType('SmartBoxTarget', [System.Reflection.BindingFlags]'Public,NonPublic')
$entryStatType = $panelType.GetNestedType('SmartEntryStat', [System.Reflection.BindingFlags]'Public,NonPublic')
$targetResolutionType = $panelType.GetNestedType('SmartTargetEntryResolution', [System.Reflection.BindingFlags]'Public,NonPublic')
$mappingFeedbackTargetType = $panelType.GetNestedType('MappingFeedbackTarget', [System.Reflection.BindingFlags]'Public,NonPublic')
$routeType = $panelType.GetNestedType('SmartMethodRoute', [System.Reflection.BindingFlags]'Public,NonPublic')
$snapshotType = $panelType.GetNestedType('SmartLearningSnapshot', [System.Reflection.BindingFlags]'Public,NonPublic')
$canAutoSelect = $panelType.GetMethod('CanAutoSelectSmartMapEntry', $allFlags)
$isClassifiedEntryCode = $panelType.GetMethod('IsSmartClassifiedEntryCode', $allFlags)
$resolveRoute = $panelType.GetMethod('ResolveSmartMethodRoute', $allFlags)
$resolveEntryName = $panelType.GetMethod('ResolveSmartEntryName', $allFlags)
$shouldWarnPartition = $panelType.GetMethod('ShouldWarnSmartLibraryPartitionMissing', $allFlags)
$orderCandidates = $panelType.GetMethod('OrderSmartMapCandidateScores', $allFlags)
$resolveTargetEntries = $panelType.GetMethod('ResolveSmartTargetEntries', $allFlags)
$isPrimaryTarget = $panelType.GetMethod('IsPrimaryLearningTarget', $allFlags)
$isEngineeringScopeTarget = $panelType.GetMethod('IsEngineeringScopeLearningTarget', $allFlags)
$isLearningGroupRecommendable = $panelType.GetMethod('IsLearningGroupRecommendable', $allFlags)
$isSmartTargetSetRecommendable = $panelType.GetMethod('IsSmartTargetSetRecommendable', $allFlags)
$isSingleQuotaTargetBox = $panelType.GetMethod('IsSingleQuotaTargetBox', $allFlags)
$hasCompatibleSpecifications = $panelType.GetMethod('HaveCompatibleSmartSpecificationNumbers', $allFlags)
if ($null -eq $candidateType -or $null -eq $entryType -or $null -eq $targetType -or $null -eq $routeType -or
    $null -eq $entryStatType -or $null -eq $targetResolutionType -or $null -eq $mappingFeedbackTargetType -or
    $null -eq $snapshotType -or $null -eq $canAutoSelect -or $null -eq $resolveRoute -or
    $null -eq $resolveEntryName -or $null -eq $shouldWarnPartition -or
    $null -eq $isClassifiedEntryCode -or
    $null -eq $orderCandidates -or $null -eq $resolveTargetEntries -or
    $null -eq $isPrimaryTarget -or $null -eq $isEngineeringScopeTarget -or $null -eq $isLearningGroupRecommendable -or
    $null -eq $isSmartTargetSetRecommendable -or $null -eq $isSingleQuotaTargetBox -or $null -eq $hasCompatibleSpecifications) {
    throw '缺少跨专业同名冲突判定入口'
}
if ($null -ne $candidateType.GetField('PendingLocal', $allFlags)) {
    throw 'PendingLocal 不应在 SmartMapCandidateScore 重复存储'
}
$routeCases = @(
    @('101号文估算', '2020', '101-estimate', '101号文估算'),
    @('101-estimate', '2020', '101-estimate', '101号文估算'),
    @('国铁科法〔2017〕30号文', '2020', '2020', '30号文'),
    @('2020', '2020', '2020', '30号文'),
    @('2024', '2024', '2024', 'TB 10801—2024'),
    @('TB10801-2024', '2024', '2024', 'TB 10801—2024')
)
foreach ($routeCase in $routeCases) {
    $routeArgs = New-Object 'object[]' 1
    $routeArgs[0] = $routeCase[0]
    $route = $resolveRoute.Invoke($null, $routeArgs)
    if ($routeType.GetField('RawMethod', $allFlags).GetValue($route) -ne $routeCase[0] -or
        $routeType.GetField('LearningMethod', $allFlags).GetValue($route) -ne $routeCase[1] -or
        $routeType.GetField('LibraryMethod', $allFlags).GetValue($route) -ne $routeCase[2] -or
        $routeType.GetField('MethodNo', $allFlags).GetValue($route) -ne $routeCase[3]) {
        throw "办法分区路由错误：$($routeCase[0])"
    }
}
if ([regex]::Matches($smart, 'FROM dbo\.ChapterEntry WHERE method=@library_method AND method_no=@method_no').Count -ne 2 -or
    $smart -notmatch 'q\.method=@library_method AND q\.method_no=@method_no' -or
    $smart -match "method\s+IN\s*\(\s*'2020'\s*,\s*'101-estimate'" -or
    -not $smart.Contains('WHERE m.weight > 0 AND m.software_partition=@software_partition') -or
    -not $smart.Contains('WHERE software_partition=@software_partition AND method_no=@method_no')) {
    throw '参考库、普通关系分区或条目办法号未精确路由'
}
$basePartitionPosition = $smart.IndexOf('SELECT COUNT(*) FROM dbo.EntryQuota', [StringComparison]::Ordinal)
$semiJoinPosition = $smart.IndexOf('SELECT quota_code, entry_code, entry_name, project_count FROM dbo.EntryQuota', [StringComparison]::Ordinal)
if ($basePartitionPosition -lt 0 -or $semiJoinPosition -lt 0 -or $basePartitionPosition -gt $semiJoinPosition) {
    throw '零命中告警没有在 EntryQuota 半连接前检查精确基础分区'
}
$warnArgs = New-Object 'object[]' 1
$warnArgs[0] = 1
if ([bool]$shouldWarnPartition.Invoke($null, $warnArgs)) { throw '基础分区存在但半连接可为 0 时被误告警' }
$warnArgs[0] = 0
if (-not [bool]$shouldWarnPartition.Invoke($null, $warnArgs)) { throw 'MethodNo 错误导致基础分区为 0 时未告警' }
$nameSnapshot = [Activator]::CreateInstance($snapshotType, $true).PSObject.BaseObject
$projectNames = $snapshotType.GetField('ProjectEntryNameByCode', $allFlags).GetValue($nameSnapshot)
$projectNames['0101'] = '当前项目条目名'
$nameArgs = New-Object 'object[]' 3
$nameArgs[0] = $nameSnapshot
$nameArgs[1] = '0101'
$nameArgs[2] = '30号文学习侧条目名'
if ($resolveEntryName.Invoke($null, $nameArgs) -ne '当前项目条目名') {
    throw '101 号文缺少 ChapterEntry 分区时未优先使用当前项目条目名'
}
[void]$projectNames.Remove('0101')
$learningNames = $snapshotType.GetField('LearningEntryNameByCode', $allFlags).GetValue($nameSnapshot)
$learningNames['0101'] = '精确分区 ChapterEntry 条目名'
if ($resolveEntryName.Invoke($null, $nameArgs) -ne '精确分区 ChapterEntry 条目名') {
    throw '当前项目无条目名时未优先使用精确分区 ChapterEntry'
}

function New-SmartTarget([string]$Code, [string]$Name, [string]$Unit, [string]$Kind = 'quota') {
    $target = [Activator]::CreateInstance($targetType, $true).PSObject.BaseObject
    $targetType.GetField('Kind', $allFlags).SetValue($target, $Kind)
    $targetType.GetField('Code', $allFlags).SetValue($target, $Code)
    $targetType.GetField('Name', $allFlags).SetValue($target, $Name)
    $targetType.GetField('Unit', $allFlags).SetValue($target, $Unit)
    return $target
}
function Add-SmartEntryStat($Dictionary, [string]$Key, [string]$EntryCode, [string]$EntryName, [bool]$CurrentMethod) {
    if (-not $Dictionary.ContainsKey($Key)) {
        $Dictionary.Add($Key, [Activator]::CreateInstance($smartEntryStatListType).PSObject.BaseObject)
    }
    $stat = [Activator]::CreateInstance($entryStatType, $true).PSObject.BaseObject
    $entryStatType.GetField('EntryCode', $allFlags).SetValue($stat, $EntryCode)
    $entryStatType.GetField('EntryName', $allFlags).SetValue($stat, $EntryName)
    $entryStatType.GetField('ProjectCount', $allFlags).SetValue($stat, 3)
    $entryStatType.GetField('CurrentMethodEvidence', $allFlags).SetValue($stat, $CurrentMethod)
    [void]$Dictionary[$Key].Add($stat)
}
function Invoke-ResolveTargetEntries($Snapshot, $ProjectEntries, $Entry, [string]$Signature) {
    $invokeArgs = New-Object 'object[]' 5
    $invokeArgs[0] = $Snapshot
    $invokeArgs[1] = $ProjectEntries
    $invokeArgs[2] = $Entry
    $invokeArgs[3] = $Signature
    $invokeArgs[4] = $null
    return $resolveTargetEntries.Invoke($null, $invokeArgs)
}

$targetSnapshot = [Activator]::CreateInstance($snapshotType, $true).PSObject.BaseObject
$snapshotType.GetField('Method', $allFlags).SetValue($targetSnapshot, '2024')
$targetProjectNames = $snapshotType.GetField('ProjectEntryNameByCode', $allFlags).GetValue($targetSnapshot)
$targetProjectNames['0801'] = '安装工程费'
$targetProjectNames['0802'] = '设备购置费'
$projectEntryType = $resolveTargetEntries.GetParameters()[1].ParameterType
$targetProjectEntries = [Activator]::CreateInstance($projectEntryType).PSObject.BaseObject
$targetProjectEntries['0801'] = [long]801
$targetProjectEntries['0802'] = [long]802
$statsByTarget = $snapshotType.GetField('EntryBySignatureQuota', $allFlags).GetValue($targetSnapshot)
$smartEntryStatListType = $statsByTarget.GetType().GetGenericArguments()[1]
$targetSignature = 'target-entry|'
$targetEntry = [Activator]::CreateInstance($entryType, $true).PSObject.BaseObject
$ordinaryTarget = New-SmartTarget 'EY-299' '安装定额' '台'
$sfTarget = New-SmartTarget 'SF' '设备购置费' '元'
[void]$entryType.GetField('Targets', $allFlags).GetValue($targetEntry).Add($ordinaryTarget)
[void]$entryType.GetField('Targets', $allFlags).GetValue($targetEntry).Add($sfTarget)
Add-SmartEntryStat $statsByTarget ($targetSignature + "`nEY-299") '0801' '历史安装条目' $true
Add-SmartEntryStat $statsByTarget ($targetSignature + "`nSF") '0802' '历史设备条目' $true
$resolvedTargets = @(Invoke-ResolveTargetEntries $targetSnapshot $targetProjectEntries $targetEntry $targetSignature)
$ordinaryResolved = @($resolvedTargets | Where-Object { $_.Target.Code -eq 'EY-299' })[0]
$sfResolved = @($resolvedTargets | Where-Object { $_.Target.Code -eq 'SF' })[0]
if ($ordinaryResolved.EntryCode -ne '0801' -or $ordinaryResolved.EntryName -ne '安装工程费' -or
    $sfResolved.EntryCode -ne '0802' -or $sfResolved.EntryName -ne '设备购置费' -or
    -not $ordinaryResolved.FromCurrentContext -or -not $sfResolved.FromCurrentContext) {
    throw '普通定额与 SF 未按目标分别解析到安装工程费/设备购置费条目'
}

$statsByTarget.Clear()
Add-SmartEntryStat $statsByTarget ($targetSignature + "`nEY-299") '0802' '设备购置费' $true
Add-SmartEntryStat $statsByTarget ($targetSignature + "`nSF") '0801' '安装工程费' $true
$blockedTargets = @(Invoke-ResolveTargetEntries $targetSnapshot $targetProjectEntries $targetEntry $targetSignature)
$blockedOrdinary = @($blockedTargets | Where-Object { $_.Target.Code -eq 'EY-299' })[0]
$blockedSf = @($blockedTargets | Where-Object { $_.Target.Code -eq 'SF' })[0]
if (-not [String]::IsNullOrWhiteSpace($blockedOrdinary.EntryCode) -or
    $blockedOrdinary.Issue -notlike '*设备购置费条目只能写入 SF*' -or
    -not [String]::IsNullOrWhiteSpace($blockedSf.EntryCode) -or
    $blockedSf.Issue -notlike '*SF 必须写入设备购置费条目*') {
    throw 'SF 双向条目约束未同时阻断普通定额落设备购置费及 SF 落普通条目'
}

$statsByTarget.Clear()
$followerEntry = [Activator]::CreateInstance($entryType, $true).PSObject.BaseObject
$followerOrdinary = New-SmartTarget 'EY-299' '安装定额' '台'
$zlfTarget = New-SmartTarget 'ZLF' '装料费' 'm3'
[void]$entryType.GetField('Targets', $allFlags).GetValue($followerEntry).Add($followerOrdinary)
[void]$entryType.GetField('Targets', $allFlags).GetValue($followerEntry).Add($zlfTarget)
Add-SmartEntryStat $statsByTarget ($targetSignature + "`nEY-299") '0801' '安装工程费' $true
$followerTargets = @(Invoke-ResolveTargetEntries $targetSnapshot $targetProjectEntries $followerEntry $targetSignature)
$resolvedFollower = @($followerTargets | Where-Object { $_.Target.Code -eq 'ZLF' })[0]
if ($resolvedFollower.EntryCode -ne '0801' -or $resolvedFollower.FromCurrentContext) {
    throw 'ZLF/LF 跟随普通定额时应继承条目，但不得伪装成目标级当前办法证据'
}

foreach ($case in @(
    @('quota','EY-299',$true,$true),
    @('quota','SF',$false,$true),
    @('quota','SH',$false,$false),
    @('quota','ZLF',$false,$false),
    @('quota','LF',$false,$false),
    @('material','1009001',$false,$false)
)) {
    $primaryArgs = New-Object 'object[]' 2; $primaryArgs[0] = $case[0]; $primaryArgs[1] = $case[1]
    if ([bool]$isPrimaryTarget.Invoke($null, $primaryArgs) -ne [bool]$case[2] -or
        [bool]$isEngineeringScopeTarget.Invoke($null, $primaryArgs) -ne [bool]$case[3]) {
        throw "普通主目标或工程范围归集分类错误：$($case[0])/$($case[1])"
    }
}

function New-FeedbackTarget([string]$Code, [string]$EntryName, [string]$Kind = 'quota') {
    $target = [Activator]::CreateInstance($mappingFeedbackTargetType, $true).PSObject.BaseObject
    $mappingFeedbackTargetType.GetField('Kind', $allFlags).SetValue($target, $Kind)
    $mappingFeedbackTargetType.GetField('Code', $allFlags).SetValue($target, $Code)
    $mappingFeedbackTargetType.GetField('Name', $allFlags).SetValue($target, $Code)
    $mappingFeedbackTargetType.GetField('Unit', $allFlags).SetValue($target, '元')
    $mappingFeedbackTargetType.GetField('EntryName', $allFlags).SetValue($target, $EntryName)
    return $target
}
$feedbackListType = [System.Collections.Generic.List``1].MakeGenericType($mappingFeedbackTargetType)
function Test-LearningTargets([object[]]$Targets) {
    $list = [Activator]::CreateInstance($feedbackListType).PSObject.BaseObject
    foreach ($target in $Targets) { [void]$list.Add($target) }
    $args = New-Object 'object[]' 2; $args[0] = $list; $args[1] = ''
    return [bool]$isLearningGroupRecommendable.Invoke($null, $args)
}
if (-not (Test-LearningTargets @((New-FeedbackTarget 'EY-299' '安装工程费'), (New-FeedbackTarget 'SF' '设备购置费'))) -or
    (Test-LearningTargets @((New-FeedbackTarget 'EY-299' '设备购置费'))) -or
    (Test-LearningTargets @((New-FeedbackTarget 'SF' '安装工程费'))) -or
    (Test-LearningTargets @((New-FeedbackTarget 'SF' '设备购置费' 'material'))) -or
    (Test-LearningTargets @((New-FeedbackTarget 'ZLF' '安装工程费'))) -or
    (Test-LearningTargets @((New-FeedbackTarget 'SH' '安装工程费'))) -or
    -not (Test-LearningTargets @((New-FeedbackTarget 'EY-299' '安装工程费'), (New-FeedbackTarget 'ZLF' '安装工程费'))) -or
    -not (Test-LearningTargets @((New-FeedbackTarget 'SF' '设备购置费')))) {
    throw '持久化入口没有对 SF 双向条目约束做防御性校验'
}

$smartTargetListType = [System.Collections.Generic.List``1].MakeGenericType($targetType)
function Test-SmartTargetSet([object[]]$Targets) {
    $list = [Activator]::CreateInstance($smartTargetListType).PSObject.BaseObject
    foreach ($target in $Targets) { [void]$list.Add($target) }
    $args = New-Object 'object[]' 1
    $args[0] = $list
    return [bool]$isSmartTargetSetRecommendable.Invoke($null, $args)
}
if ((Test-SmartTargetSet @((New-SmartTarget 'ZLF' '装料费' 'm3'))) -or
    (Test-SmartTargetSet @((New-SmartTarget 'SH' '设备费' '项'))) -or
    (Test-SmartTargetSet @((New-SmartTarget 'SF' '设备购置费' '元' 'material'))) -or
    -not (Test-SmartTargetSet @((New-SmartTarget 'EY-299' '安装定额' '台'), (New-SmartTarget 'ZLF' '装料费' 'm3'))) -or
    -not (Test-SmartTargetSet @((New-SmartTarget 'SF' '设备购置费' '元')))) {
    throw '历史纯辅助聚合框未在 SmartFill 读取端过滤，或误伤混合组件/纯 SF 设备费'
}
$singleAuxEntry = [Activator]::CreateInstance($entryType, $true).PSObject.BaseObject
[void]$entryType.GetField('Targets', $allFlags).GetValue($singleAuxEntry).Add((New-SmartTarget 'ZLF' '装料费' 'm3'))
if ([bool]$isSingleQuotaTargetBox.Invoke($null, @($singleAuxEntry))) {
    throw '单 ZLF 聚合框不得作为普通单定额框自动采纳'
}
$specArgs = [object[]]@('Φ100X10MMCPVC管', 'Φ150X10MMCPVC管')
if ([bool]$hasCompatibleSpecifications.Invoke($null, $specArgs)) { throw '不同公称直径因共同数字10被当作规格兼容' }
$specArgs = [object[]]@('Φ100X10MMCPVC管', 'Φ10X100MMCPVC管')
if ([bool]$hasCompatibleSpecifications.Invoke($null, $specArgs)) { throw '规格数字顺序不一致时不得模糊匹配' }
$specArgs = [object[]]@('Φ100X10MMCPVC管', 'Φ100X10MMCPVC管')
if (-not [bool]$hasCompatibleSpecifications.Invoke($null, $specArgs)) { throw '同规格异形符号归一后应保持兼容' }
if ($smart -notmatch 'targetEntries\.All\(item => item\.FromCurrentContext\)' -or
    $smart -notmatch 'targetEntries\.All\(item => item != null && !String\.IsNullOrWhiteSpace\(item\.EntryCode\)\)') {
    throw 'HasEntry/HasCurrentContext 未按组内全部目标判定'
}
$classifiedCases = @{
    '12-01' = $true
    'SF' = $false
    'XGT1' = $false
    '12A' = $false
}
foreach ($case in $classifiedCases.GetEnumerator()) {
    $caseArgs = New-Object 'object[]' 1
    $caseArgs[0] = [string]$case.Key
    if ([bool]$isClassifiedEntryCode.Invoke($null, $caseArgs) -ne [bool]$case.Value) {
        throw "C# 条目分类过滤错误：$($case.Key)"
    }
}
$rebuildFilter = [regex]::Match($rebuildAggregates, '(?ms)^function Test-ClassifiedEntryCode\s*\{.*?^\}')
if (-not $rebuildFilter.Success -or -not $learningDbSource.Contains('if (!IsSmartClassifiedEntryCode(entryCode)) continue;')) {
    throw ('增量写入与全量重算未同时接入条目分类过滤: rebuild=' + $rebuildFilter.Success +
        ', incremental=' + $learningDbSource.Contains('if (!IsSmartClassifiedEntryCode(entryCode)) continue;'))
}
. ([ScriptBlock]::Create($rebuildFilter.Value))
foreach ($case in $classifiedCases.GetEnumerator()) {
    if ([bool](Test-ClassifiedEntryCode ([string]$case.Key)) -ne [bool]$case.Value) {
        throw "重算脚本条目分类过滤错误：$($case.Key)"
    }
}
function New-SmartCandidate([int]$Weight) {
    $entry = [Activator]::CreateInstance($entryType, $true).PSObject.BaseObject
    $entryType.GetField('Weight', $allFlags).SetValue($entry, $Weight)
    $candidate = [Activator]::CreateInstance($candidateType, $true).PSObject.BaseObject
    $candidateType.GetField('Entry', $allFlags).SetValue($candidate, $entry)
    foreach ($fieldName in @('HasEntry', 'HasCurrentContext', 'HasCurrentMethodMapping', 'CurrentTargetsValid')) {
        $candidateType.GetField($fieldName, $allFlags).SetValue($candidate, $true)
    }
    return $candidate
}
function Set-SmartEvidence($Candidate, [int]$Accepted, [int]$Corrected, [int]$Rejected) {
    $entry = $candidateType.GetField('Entry', $allFlags).GetValue($Candidate)
    $entryType.GetField('AcceptedCount', $allFlags).SetValue($entry, $Accepted)
    $entryType.GetField('CorrectedCount', $allFlags).SetValue($entry, $Corrected)
    $entryType.GetField('RejectedCount', $allFlags).SetValue($entry, $Rejected)
}
function Set-SmartCandidateFlag($Candidate, [string]$FieldName, [bool]$Value) {
    $candidateType.GetField($FieldName, $allFlags).SetValue($Candidate, $Value)
}
function Add-SmartTarget($Candidate, [string]$Kind, [string]$Code) {
    $entry = $candidateType.GetField('Entry', $allFlags).GetValue($Candidate)
    $target = [Activator]::CreateInstance($targetType, $true).PSObject.BaseObject
    $targetType.GetField('Kind', $allFlags).SetValue($target, $Kind)
    $targetType.GetField('Code', $allFlags).SetValue($target, $Code)
    [void]$entryType.GetField('Targets', $allFlags).GetValue($entry).Add($target)
}
function Test-CanAutoSelect([object[]]$Items) {
    $list = [Activator]::CreateInstance($candidateListType).PSObject.BaseObject
    foreach ($item in $Items) { [void]$list.Add($item) }
    $args = New-Object 'object[]' 1
    $args[0] = $list
    return [bool]$canAutoSelect.Invoke($null, $args)
}
$candidateListType = [System.Collections.Generic.List``1].MakeGenericType($candidateType)
$candidates = [Activator]::CreateInstance($candidateListType).PSObject.BaseObject
[void]$candidates.Add((New-SmartCandidate 20))
[void]$candidates.Add((New-SmartCandidate 20))
$canAutoArgs = New-Object 'object[]' 1
$canAutoArgs[0] = $candidates
if ([bool]$canAutoSelect.Invoke($null, $canAutoArgs)) {
    throw '跨专业同名的同权重候选不应自动勾选'
}

$topOneAccepted = New-SmartCandidate 10
Set-SmartEvidence $topOneAccepted 1 0 0
$emptyMethodHigh = New-SmartCandidate 100
Set-SmartCandidateFlag $emptyMethodHigh 'HasCurrentMethodMapping' $false
if (Test-CanAutoSelect @($topOneAccepted, $emptyMethodHigh)) {
    throw '当前办法候选只有 1 次 accepted 时不应压过空办法高权重候选'
}
$topTwoAccepted = New-SmartCandidate 20
Set-SmartEvidence $topTwoAccepted 2 0 0
if (-not (Test-CanAutoSelect @($topTwoAccepted, $emptyMethodHigh))) {
    throw '当前办法候选累计 2 次 accepted 后未自动采纳'
}
$topCorrected = New-SmartCandidate 20
Set-SmartEvidence $topCorrected 0 1 0
if (-not (Test-CanAutoSelect @($topCorrected, $emptyMethodHigh))) {
    throw '当前办法候选 1 次 corrected 后未自动采纳'
}
$sameMethodA = New-SmartCandidate 20
$sameMethodB = New-SmartCandidate 20
Set-SmartEvidence $sameMethodA 2 0 0
Set-SmartEvidence $sameMethodB 2 0 0
if (Test-CanAutoSelect @($sameMethodA, $sameMethodB)) {
    throw '同权重且都有当前办法证据时不应自动采纳'
}
$rejectedTop = New-SmartCandidate 0
Set-SmartEvidence $rejectedTop 2 0 3
$strongerSecond = New-SmartCandidate 10
Set-SmartEvidence $strongerSecond 1 0 0
if (Test-CanAutoSelect @($rejectedTop, $strongerSecond)) {
    throw 'top 被 rejected 压到低于 second 后仍自动采纳'
}
$singleEmptyMethod = New-SmartCandidate 10
Set-SmartCandidateFlag $singleEmptyMethod 'HasCurrentMethodMapping' $false
Add-SmartTarget $singleEmptyMethod 'quota' 'Q-ONLY'
if (-not (Test-CanAutoSelect @($singleEmptyMethod))) {
    throw '单 quota 目标的唯一空办法候选未使用当前办法唯一条目证据自动采纳'
}
$multiEmptyMethod = New-SmartCandidate 20
Set-SmartCandidateFlag $multiEmptyMethod 'HasCurrentMethodMapping' $false
Add-SmartTarget $multiEmptyMethod 'quota' 'Q-1'
Add-SmartTarget $multiEmptyMethod 'quota' 'Q-2'
if (Test-CanAutoSelect @($multiEmptyMethod)) {
    throw '多目标空办法组件框不应自动采纳'
}

$singleTarget = New-SmartCandidate 20
Add-SmartTarget $singleTarget 'quota' 'DY-1250'
$completeComponent = New-SmartCandidate 20
Add-SmartTarget $completeComponent 'quota' 'DY-1250'
Add-SmartTarget $completeComponent 'quota' 'SF'
$componentRankInput = [Activator]::CreateInstance($candidateListType).PSObject.BaseObject
[void]$componentRankInput.Add($singleTarget)
[void]$componentRankInput.Add($completeComponent)
$componentRankArgs = New-Object 'object[]' 1
$componentRankArgs[0] = $componentRankInput
$componentRanked = $orderCandidates.Invoke($null, $componentRankArgs)
if (-not [Object]::ReferenceEquals($componentRanked[0], $completeComponent)) {
    throw '证据权重相同时，完整组件框应优先于其单条候选'
}
Write-Host 'Test-SmartFill: PASS'
