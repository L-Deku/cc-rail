# 推荐定额(学习库智能铺量)标记测试:引擎、UI 挂接、菜单入口、老窗口删除。
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
$smartPath = Join-Path $repoRoot 'tools\RecoExpandPanel\SmartFillFeature.cs'
if (-not (Test-Path -LiteralPath $smartPath)) { throw '缺少 SmartFillFeature.cs' }
$smart = Get-Content -LiteralPath $smartPath -Raw
$panel = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\TemplateFillPanel.cs') -Raw
$excelLink = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\ExcelLinkFeature.cs') -Raw
$learningDb = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\LearningDbFeature.cs') -Raw
$rebuildAggregates = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoLearning\Rebuild-Aggregates.ps1') -Raw
$quotaPanel = Get-Content -LiteralPath (Join-Path $repoRoot 'RecoQuotaRecommend\QuotaRecommendPanel.cs') -Raw
$oldDialog = Get-Content -LiteralPath (Join-Path $repoRoot 'RecoQuotaRecommend\RecommendDialog.cs') -Raw
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
if ($smart -notmatch 'BuildMappingBoxIndex|LoadMappingBoxRows') { throw '缺少 jsonl 回退' }
if ($smart -notmatch 'IsLibraryQuota = true') { throw '缺少库内定额原生粘贴路径' }
if ($smart -notmatch 'IsNameDriven = true') { throw '推荐定额项必须 IsNameDriven=true,否则不回流学习库' }
if ($smart -notmatch 'TemplateName = "推荐定额"') { throw '预览项名称未改为推荐定额' }
if ($smart -notmatch 'EntryBySignatureQuota') { throw '缺少签名级条目证据' }
if ($smart -notmatch 'prefixVotes') { throw '缺少工程前缀投票' }
if ($smart -notmatch 'preferredPrefixes') { throw '缺少前缀过滤消歧' }
if ($smart -notmatch 'SmartLearningScope') { throw '缺少推荐学习库范围模型' }
if ($smart -notmatch 'LoadSmartLearningScopes') { throw '缺少推荐学习库目录加载' }
if ($smart -notmatch '全库兜底') { throw '缺少专业学习库无命中后的全库兜底标记' }
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
if ($learningDb -match 'Math\.Min\(100' -or $learningDb -match '>100 THEN 100' -or
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

$globalExactPosition = $smart.IndexOf('"名称学习命中，全库兜底"', [StringComparison]::Ordinal)
$fuzzyPosition = $smart.IndexOf('BuildSmartFuzzyScoresIfUnmatched(matched', [StringComparison]::Ordinal)
if ($globalExactPosition -lt 0 -or $fuzzyPosition -lt 0 -or $globalExactPosition -gt $fuzzyPosition) {
    throw '全库精确命中未优先于模糊打分'
}

$allFlags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$candidateType = $panelType.GetNestedType('SmartMapCandidateScore', [System.Reflection.BindingFlags]'Public,NonPublic')
$entryType = $panelType.GetNestedType('SmartMapEntry', [System.Reflection.BindingFlags]'Public,NonPublic')
$targetType = $panelType.GetNestedType('SmartBoxTarget', [System.Reflection.BindingFlags]'Public,NonPublic')
$canAutoSelect = $panelType.GetMethod('CanAutoSelectSmartMapEntry', $allFlags)
$applyPending = $panelType.GetMethod('ApplyPendingLocalSmartMapEntry', $allFlags)
$orderCandidates = $panelType.GetMethod('OrderSmartMapCandidateScores', $allFlags)
if ($null -eq $candidateType -or $null -eq $entryType -or $null -eq $targetType -or $null -eq $canAutoSelect -or
    $null -eq $applyPending -or $null -eq $orderCandidates) {
    throw '缺少跨专业同名冲突判定入口'
}
if ($null -ne $candidateType.GetField('PendingLocal', $allFlags)) {
    throw 'PendingLocal 不应在 SmartMapCandidateScore 重复存储'
}
function New-SmartCandidate([int]$Weight, [bool]$PendingLocal = $false) {
    $entry = [Activator]::CreateInstance($entryType, $true).PSObject.BaseObject
    $entryType.GetField('Weight', $allFlags).SetValue($entry, $Weight)
    $entryType.GetField('PendingLocal', $allFlags).SetValue($entry, $PendingLocal)
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

$sqlEntry = [Activator]::CreateInstance($entryType, $true).PSObject.BaseObject
$entryType.GetField('BoxId', $allFlags).SetValue($sqlEntry, 'sql-box')
$entryType.GetField('Weight', $allFlags).SetValue($sqlEntry, 90)
$pendingArgs = New-Object 'object[]' 3
$pendingArgs[0] = $sqlEntry
$pendingArgs[1] = 'sql-box'
$pendingArgs[2] = 30
$mergedEntry = $applyPending.Invoke($null, $pendingArgs)
if ($entryType.GetField('Weight', $allFlags).GetValue($mergedEntry) -ne 90 -or
    -not [bool]$entryType.GetField('PendingLocal', $allFlags).GetValue($mergedEntry)) {
    throw '本机 pending(weight 30) 覆盖了 SQL weight 90，或未设置 PendingLocal'
}

$rankInput = [Activator]::CreateInstance($candidateListType).PSObject.BaseObject
$ordinary = New-SmartCandidate 90 $false
$pending = New-SmartCandidate 90 $true
[void]$rankInput.Add($ordinary)
[void]$rankInput.Add($pending)
$rankArgs = New-Object 'object[]' 1
$rankArgs[0] = $rankInput
$ranked = $orderCandidates.Invoke($null, $rankArgs)
if (-not [Object]::ReferenceEquals($ranked[0], $pending)) {
    throw '同权重候选中 PendingLocal 未优先排序'
}
Write-Host 'Test-SmartFill: PASS'
