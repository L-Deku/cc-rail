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
if ($smart -notmatch 'AttachSmartCandidateOptions') { throw '高分组件自动选中后没有保留其他有效组件下拉。' }
if ($smart -notmatch 'LoadSmartLearningSnapshot') { throw '缺少学习库快照加载 LoadSmartLearningSnapshot' }
if ($smart -match 'LoadMappingBoxRows\(' -or $smart -match '本地映射\(jsonl回退\)') { throw '推荐定额仍可能从本地学习配对' }
if ($smart -notmatch 'local learning is disabled') { throw 'SQL 失败没有明确关闭本地学习' }
if ($smart -notmatch 'IsLibraryQuota = true') { throw '缺少库内目标候选标记' }
if ($smart -notmatch 'IsNameDriven = true') { throw '推荐定额项必须 IsNameDriven=true,否则不回流学习库' }
if ($smart -notmatch 'TemplateName = "推荐定额"') { throw '预览项名称未改为推荐定额' }
if ($smart -match 'EntryBySignatureQuota|prefixVotes|preferredPrefixes|ResolveSmartTargetEntries') { throw '已下线的条目推断仍残留' }
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
$mappingFeedbackTargetType = $panelType.GetNestedType('MappingFeedbackTarget', [System.Reflection.BindingFlags]'Public,NonPublic')
$routeType = $panelType.GetNestedType('SmartMethodRoute', [System.Reflection.BindingFlags]'Public,NonPublic')
$snapshotType = $panelType.GetNestedType('SmartLearningSnapshot', [System.Reflection.BindingFlags]'Public,NonPublic')
$previewItemType = $panelType.GetNestedType('FillPreviewItem', [System.Reflection.BindingFlags]'Public,NonPublic')
$nameCandidateType = $panelType.GetNestedType('NameQuotaCandidateGroup', [System.Reflection.BindingFlags]'Public,NonPublic')
$canAutoSelect = $panelType.GetMethod('CanAutoSelectSmartMapEntry', $allFlags)
$attachCandidateOptions = $panelType.GetMethod('AttachSmartCandidateOptions', $allFlags)
$isClassifiedEntryCode = $panelType.GetMethod('IsSmartClassifiedEntryCode', $allFlags)
$resolveRoute = $panelType.GetMethod('ResolveSmartMethodRoute', $allFlags)
$orderCandidates = $panelType.GetMethod('OrderSmartMapCandidateScores', $allFlags)
$isPrimaryTarget = $panelType.GetMethod('IsPrimaryLearningTarget', $allFlags)
$isEngineeringScopeTarget = $panelType.GetMethod('IsEngineeringScopeLearningTarget', $allFlags)
$isLearningGroupRecommendable = $panelType.GetMethod('IsLearningGroupRecommendable', $allFlags)
$isSmartTargetSetRecommendable = $panelType.GetMethod('IsSmartTargetSetRecommendable', $allFlags)
$isSingleQuotaTargetBox = $panelType.GetMethod('IsSingleQuotaTargetBox', $allFlags)
$hasCompatibleSpecifications = $panelType.GetMethod('HaveCompatibleSmartSpecificationNumbers', $allFlags)
if ($null -eq $candidateType -or $null -eq $entryType -or $null -eq $targetType -or $null -eq $routeType -or
    $null -eq $mappingFeedbackTargetType -or $null -eq $previewItemType -or $null -eq $nameCandidateType -or
    $null -eq $snapshotType -or $null -eq $canAutoSelect -or $null -eq $resolveRoute -or
    $null -eq $attachCandidateOptions -or
    $null -eq $isClassifiedEntryCode -or
    $null -eq $orderCandidates -or
    $null -eq $isPrimaryTarget -or $null -eq $isEngineeringScopeTarget -or $null -eq $isLearningGroupRecommendable -or
    $null -eq $isSmartTargetSetRecommendable -or $null -eq $isSingleQuotaTargetBox -or $null -eq $hasCompatibleSpecifications) {
    throw '缺少跨专业同名冲突判定入口'
}

$previewListType = [System.Collections.Generic.List``1].MakeGenericType($previewItemType)
$candidateListType = [System.Collections.Generic.List``1].MakeGenericType($nameCandidateType)
$activePreview = [Activator]::CreateInstance($previewListType).PSObject.BaseObject
$activeItem = [Activator]::CreateInstance($previewItemType, $true).PSObject.BaseObject
[void]$activePreview.Add($activeItem)
$candidateOptions = [Activator]::CreateInstance($candidateListType).PSObject.BaseObject
foreach ($definition in @(
    [pscustomobject]@{ Key='smart:three'; Label='DY-310 + DY-480 + 7015511*1.02' },
    [pscustomobject]@{ Key='smart:two'; Label='DY-480 + 7015511*1.02' }
)) {
    $candidate = [Activator]::CreateInstance($nameCandidateType, $true).PSObject.BaseObject
    $nameCandidateType.GetField('Key', $allFlags).SetValue($candidate, $definition.Key)
    $nameCandidateType.GetField('Label', $allFlags).SetValue($candidate, $definition.Label)
    [void]$candidateOptions.Add($candidate)
}
$attachArgs = New-Object 'object[]' 3
$attachArgs[0] = $activePreview
$attachArgs[1] = $candidateOptions
$attachArgs[2] = $false
$attachCandidateOptions.Invoke($null, $attachArgs)
if (-not [bool]$previewItemType.GetField('Selected', $allFlags).GetValue($activeItem) -or
    [bool]$previewItemType.GetField('NeedExactNameConfirmation', $allFlags).GetValue($activeItem) -or
    $previewItemType.GetField('SelectedNameQuotaCandidateKey', $allFlags).GetValue($activeItem) -ne 'smart:three' -or
    $previewItemType.GetField('NameQuotaCandidates', $allFlags).GetValue($activeItem).Count -ne 2) {
    throw '最高分组件默认勾选后未保留三条/两条完整组件下拉。'
}
Write-Host 'PASS 最高分组件默认勾选且保留其他有效完整组件下拉'
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
    $smart -match "method\s+IN\s*\(\s*'2020'\s*,\s*'101-estimate'" -or
    -not $smart.Contains('WHERE m.weight > 0 AND m.software_partition=@software_partition') -or
    -not $smart.Contains('WHERE software_partition=@software_partition AND method_no=@method_no')) {
    throw '参考库、普通关系分区或条目办法号未精确路由'
}

function New-SmartTarget([string]$Code, [string]$Name, [string]$Unit, [string]$Kind = 'quota') {
    $target = [Activator]::CreateInstance($targetType, $true).PSObject.BaseObject
    $targetType.GetField('Kind', $allFlags).SetValue($target, $Kind)
    $targetType.GetField('Code', $allFlags).SetValue($target, $Code)
    $targetType.GetField('Name', $allFlags).SetValue($target, $Name)
    $targetType.GetField('Unit', $allFlags).SetValue($target, $Unit)
    return $target
}
foreach ($case in @(
    @('quota','EY-299',$true,$true),
    @('quota','SF',$false,$true),
    @('quota','SH',$false,$true),
    @('quota','ZLF',$false,$true),
    @('quota','LF',$false,$true),
    @('material','1009001',$false,$true)
)) {
    $primaryArgs = New-Object 'object[]' 2; $primaryArgs[0] = $case[0]; $primaryArgs[1] = $case[1]
    if ([bool]$isPrimaryTarget.Invoke($null, $primaryArgs) -ne [bool]$case[2] -or
        [bool]$isEngineeringScopeTarget.Invoke($null, $primaryArgs) -ne [bool]$case[3]) {
        throw "普通主目标或工程范围归集分类错误：$($case[0])/$($case[1])"
    }
}
Write-Host 'PASS EngineeringTemplate 工程范围门禁接纳纯 SH/SF/ZLF/LF 与正式材料'

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
    -not (Test-LearningTargets @((New-FeedbackTarget 'ZLF' '安装工程费'))) -or
    -not (Test-LearningTargets @((New-FeedbackTarget 'SH' '安装工程费'))) -or
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
if (-not (Test-SmartTargetSet @((New-SmartTarget 'ZLF' '装料费' 'm3'))) -or
    -not (Test-SmartTargetSet @((New-SmartTarget 'SH' '设备费' '项'))) -or
    (Test-SmartTargetSet @((New-SmartTarget 'SF' '设备购置费' '元' 'material'))) -or
    -not (Test-SmartTargetSet @((New-SmartTarget 'EY-299' '安装定额' '台'), (New-SmartTarget 'ZLF' '装料费' 'm3'))) -or
    -not (Test-SmartTargetSet @((New-SmartTarget 'SF' '设备购置费' '元')))) {
    throw '完整身份的纯辅助组件未开放，或误伤混合组件/纯 SF 设备费安全约束'
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
    foreach ($fieldName in @('HasCurrentMethodMapping', 'CurrentTargetsValid')) {
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
