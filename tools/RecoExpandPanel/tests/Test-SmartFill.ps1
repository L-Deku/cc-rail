# 推荐定额(学习库智能铺量)标记测试:引擎、UI 挂接、菜单入口、老窗口删除。
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
$smartPath = Join-Path $repoRoot 'tools\RecoExpandPanel\SmartFillFeature.cs'
if (-not (Test-Path -LiteralPath $smartPath)) { throw '缺少 SmartFillFeature.cs' }
$smart = Get-Content -LiteralPath $smartPath -Raw
$panel = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\TemplateFillPanel.cs') -Raw
$excelLink = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\ExcelLinkFeature.cs') -Raw
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
$canAutoSelect = $panelType.GetMethod('CanAutoSelectSmartMapEntry', $allFlags)
if ($null -eq $candidateType -or $null -eq $entryType -or $null -eq $canAutoSelect) {
    throw '缺少跨专业同名冲突判定入口'
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
$candidateListType = [System.Collections.Generic.List``1].MakeGenericType($candidateType)
$candidates = [Activator]::CreateInstance($candidateListType).PSObject.BaseObject
[void]$candidates.Add((New-SmartCandidate 20))
[void]$candidates.Add((New-SmartCandidate 20))
$canAutoArgs = New-Object 'object[]' 1
$canAutoArgs[0] = $candidates
if ([bool]$canAutoSelect.Invoke($null, $canAutoArgs)) {
    throw '跨专业同名的同权重候选不应自动勾选'
}
Write-Host 'Test-SmartFill: PASS'
