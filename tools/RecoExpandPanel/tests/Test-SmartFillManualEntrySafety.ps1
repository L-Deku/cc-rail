$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$sourceDir = Split-Path -Parent $PSScriptRoot
$repoRoot = Split-Path -Parent (Split-Path -Parent $sourceDir)
$dll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL)) {
    $env:RECO_EXPAND_DLL
} else {
    Join-Path $repoRoot 'RecoQuotaRecommend\bin\RecoExpandPanel.dll'
}
if (-not (Test-Path -LiteralPath $dll)) { throw "Missing DLL: $dll" }
$panel = [IO.File]::ReadAllText((Join-Path $sourceDir 'TemplateFillPanel.cs'))
$feature = [IO.File]::ReadAllText((Join-Path $sourceDir 'TemplateFillFeature.cs'))
$smart = [IO.File]::ReadAllText((Join-Path $sourceDir 'SmartFillFeature.cs'))
$nameMatch = [IO.File]::ReadAllText((Join-Path $sourceDir 'TemplateFillNameMatch.cs'))
$excelLink = [IO.File]::ReadAllText((Join-Path $sourceDir 'ExcelLinkFeature.cs'))

$requiredPanel = @(
    'SmartPreviewContext',
    'smartPreviewReady',
    'currentEntryWritable',
    'RefreshApplyEnabled',
    'InvalidateSmartPreview',
    'grid.ClearSelection()',
    'DataGridViewSelectionMode.FullRowSelect',
    'GetSelectedSmartTargetRows',
    'Smart fill apply rejected',
    'PromptSmartSfBindingChoice',
    'SmartSfBindingChoice.Append',
    'MergePreviewTargetGroup',
    'ValidateSmartSfEntryConstraint',
    'DescribeSmartSfEntryConflict',
    'replaceWarning',
    'appendWarning'
)
foreach ($marker in $requiredPanel) {
    if (-not $panel.Contains($marker)) {
        throw "Missing smart panel safety marker: $marker"
    }
}

$requiredFeature = @(
    'SfRedirect',
    'SfEntryBlocked',
    'SfEntryBlockReason',
    'EntrySource',
    'ApplyFillToSelectedEntry',
    'LoadSmartFillStructuralRow',
    'BuildSmartFillL2Row',
    'HasSmartFillConstructedIdentity',
    'IsSmartFillSourceIdentityMatch',
    'foreach (PreparedSmartFillItem plan in prepared)',
    '项目业务事务已整体回滚，未学习。失败原因：',
    'Smart fill apply begin',
    'Smart fill apply plan',
    'Smart fill apply ok',
    'Smart fill apply blocked',
    'Smart fill apply failed',
    'ProjectConnection = conn',
    'ProjectConnectionIdentity = GetProjectConnectionIdentity(conn)'
)
foreach ($marker in $requiredFeature) {
    if (-not $feature.Contains($marker)) {
        throw "Missing smart apply safety marker: $marker"
    }
}
foreach ($obsoleteNativeMarker in @(
    'NativeInsertState',
    'SmartNativeInsertRecord',
    'ExecuteSmartNativeInsertGroup',
    'TrySendSmartNativeKeyCommand',
    'TryCompensateSmartNativeRows',
    'SmartFillWriteLayer.L3'
)) {
    if ($feature.Contains($obsoleteNativeMarker)) {
        throw "推荐定额仍残留旧 L3 原生输入链：$obsoleteNativeMarker"
    }
}
foreach ($constructedMarker in @(
    'if (!HasSmartFillConstructedIdentity(item))',
    'if (!IsContextSensitiveLearningCode(item.QuotaCode)) item.LearnedUnitPrice = 0m;',
    'plan.SourceRow = BuildSmartFillL2Row(structuralRow, item);'
)) {
    if (-not $feature.Contains($constructedMarker)) {
        throw "正式定额没有统一进入结构模板构造：$constructedMarker"
    }
}
$applyStart = $feature.IndexOf('private static string ApplyFillToSelectedEntry', [StringComparison]::Ordinal)
$applyEnd = $feature.IndexOf('// 写入：把选中预览项对应的源定额行', $applyStart, [StringComparison]::Ordinal)
if ($applyStart -lt 0 -or $applyEnd -le $applyStart) { throw '缺少推荐定额当前条目写入事务入口' }
$applyBody = $feature.Substring($applyStart, $applyEnd - $applyStart)
if ([regex]::Matches($applyBody, 'BeginTransaction\(').Count -ne 1) {
    throw '推荐定额写入必须只有一个项目事务，不能再拆分原生输入和 SQL 写入'
}
foreach ($transactionMarker in @(
    'long markerId = InsertQuotaRowReturnId(conn, transaction, markerSource);',
    'long newId = InsertQuotaRowReturnId(conn, transaction, row);',
    'transaction.Rollback();',
    '项目业务事务已整体回滚，未学习。失败原因：'
)) {
    if (-not $applyBody.Contains($transactionMarker)) {
        throw "marker 与业务行没有受同一事务保护：$transactionMarker"
    }
}
$crossDbStart = $smart.IndexOf('private static Dictionary<string, object> LoadCrossDbQuotaRow', [StringComparison]::Ordinal)
$crossDbEnd = $smart.IndexOf('private static List<SmartMapCandidateScore> RankSmartMapEntries', $crossDbStart, [StringComparison]::Ordinal)
if ($crossDbStart -lt 0 -or $crossDbEnd -le $crossDbStart) {
    throw '缺少跨库完整源行加载入口'
}
$crossDbBody = $smart.Substring($crossDbStart, $crossDbEnd - $crossDbStart)
foreach ($marker in @(
    'candidates.OrderByDescending(value => value.BindingId)',
    'TryLoadSmartSourceRowFromConnection',
    'return values;'
)) {
    if (-not $crossDbBody.Contains($marker)) {
        throw "跨库源行缺少同编号异义拒绝或候选回退门禁：$marker"
    }
}
$sourceRowStart = $crossDbBody.IndexOf('private static Dictionary<string, object> TryLoadSmartSourceRowFromConnection', [StringComparison]::Ordinal)
if ($sourceRowStart -lt 0) { throw '缺少跨库/同库共用的完整源行身份核对入口' }
$sourceRowBody = $crossDbBody.Substring($sourceRowStart)
foreach ($marker in @('NormalizeForSignature(actualName)', 'NormalizeForSignature(actualUnit)',
    'if (!IsSmartFillSourceIdentityMatch(item, values)) return null;', 'return values;')) {
    if (-not $sourceRowBody.Contains($marker)) { throw "完整源行缺少身份核对：$marker" }
}
if ($sourceRowBody.IndexOf('if (!IsSmartFillSourceIdentityMatch(item, values)) return null;', [StringComparison]::Ordinal) -gt
    $sourceRowBody.IndexOf('return values;', [StringComparison]::Ordinal)) {
    throw '跨库源行在完整身份核对前已经返回，无法安全回退到旧候选'
}
foreach ($marker in @('approvedEntrySequence', '写入前项目或条目已变化',
    'currentSmartEntry.EntrySequence != approvedEntrySequence')) {
    if (-not $panel.Contains($marker)) { throw "推荐定额缺少临写入前项目/条目二次核对：$marker" }
}
$previewContextStart = $panel.IndexOf('private bool IsSmartPreviewContextCurrent', [StringComparison]::Ordinal)
$previewContextEnd = $panel.IndexOf('private bool StampSelectedSmartEntries', $previewContextStart, [StringComparison]::Ordinal)
if ($previewContextStart -lt 0 -or $previewContextEnd -le $previewContextStart) {
    throw '缺少推荐定额预览上下文校验入口'
}
$previewContextBody = $panel.Substring($previewContextStart, $previewContextEnd - $previewContextStart)
foreach ($unitGateMarker in @('Object.ReferenceEquals(context.ProjectConnection', 'ProjectConnectionIdentity',
    'context.CurrentUnitId != entry.UnitId', 'context.CurrentUnitCode', '项目或当前单元已切换，请重新预览')) {
    if (-not $previewContextBody.Contains($unitGateMarker)) {
        throw "同一预览不得跨单元写入的门禁被误删：$unitGateMarker"
    }
}
if (-not $feature.Contains('out bool succeeded') -or
    -not $feature.Contains('succeeded = true;') -or
    $panel.Contains('smartResult.StartsWith(') -or
    $panel.Contains('if (smartSucceeded) InvalidateSmartPreview();')) {
    throw '推荐写入成功后不得清空预览，用户还要继续选择其他组写入别的条目'
}
$smartApplyStart = $panel.IndexOf('private void OnApply()', [StringComparison]::Ordinal)
$smartApplyEnd = $panel.IndexOf('private void SetBusy(', $smartApplyStart, [StringComparison]::Ordinal)
if ($smartApplyStart -lt 0 -or $smartApplyEnd -le $smartApplyStart) { throw '缺少推荐定额界面写入入口' }
$smartApplyBody = $panel.Substring($smartApplyStart, $smartApplyEnd - $smartApplyStart)
foreach ($forbiddenUiMarker in @('确认把选中且勾选的', 'Hide();', 'mainForm.Activate();',
    'mainForm.BringToFront();', 'RestoreSmartPanelAfterHostWrite();')) {
    if ($smartApplyBody.Contains($forbiddenUiMarker)) {
        throw "推荐写入仍会弹前置确认或隐藏窗口：$forbiddenUiMarker"
    }
}
if (-not $smartApplyBody.Contains('MessageBox.Show(this, smartResult, "推荐定额")')) {
    throw '推荐写入完成后的结果确定窗口被误删'
}
foreach ($sfBindingMarker in @('PromptSmartSfBindingChoice', 'SmartSfBindingChoice.Replace',
    'SmartSfBindingChoice.Append', 'MergePreviewTargetGroup(oldGroup, replacements)',
    'replace.Text = "替换"', 'append.Text = "补充"', 'cancel.Text = "取消"',
    'replaceWarning', 'appendWarning', 'Smart fill sf replace blocked',
    'Smart fill sf append blocked',
    'replace.Enabled = String.IsNullOrWhiteSpace(replaceWarning)',
    'append.Enabled = String.IsNullOrWhiteSpace(appendWarning)',
    'ValidateSmartSfEntryConstraint(conn, currentSmartEntry, replacements',
    'ValidateSmartSfEntryConstraint(conn, currentSmartEntry, appendCandidates',
    'RefreshSmartSfEntryState();', 'target.LearningFeedbackAttempted = false')) {
    if (-not $panel.Contains($sfBindingMarker)) {
        throw "SF 右键绑定缺少替换/补充行为：$sfBindingMarker"
    }
}
if ($panel.Contains('Smart fill apply rejected: exception') -or
    -not $panel.Contains('Smart fill apply failed: 界面异常')) {
    throw '推荐定额界面异常未正确归类为 failed'
}
if (-not $applyBody.Contains('elapsedMs=') -or
    $applyBody.Contains('return blocked("transaction_failed"')) {
    throw '推荐定额 blocked 缺少耗时，或事务 failed 仍重复记 blocked'
}
if (-not $panel.Contains('smartOnly ? "推荐定额" : "模板铺量"')) {
    throw '推荐定额异常弹框标题仍被硬编码成模板铺量'
}
$currentEntryStart = $panel.IndexOf('private bool TryResolveCurrentSmartEntry', [StringComparison]::Ordinal)
$currentEntryEnd = $panel.IndexOf('private static TreeNode ResolveSmartHostTreeNode', $currentEntryStart, [StringComparison]::Ordinal)
if ($currentEntryStart -lt 0 -or $currentEntryEnd -le $currentEntryStart) {
    throw '缺少推荐定额当前条目解析入口'
}
$currentEntryBody = $panel.Substring($currentEntryStart, $currentEntryEnd - $currentEntryStart)
foreach ($marker in @(
    'ResolveChapterNo(mainForm, conn, node)',
    'from 章节表 where 条目编号=@code',
    'IsOptionalSmartTreeSequenceConsistent(seqText, sequence)'
)) {
    if (-not $currentEntryBody.Contains($marker)) {
        throw "当前条目没有按界面条目编号回查项目章节表：$marker"
    }
}
if ([regex]::Matches($currentEntryBody, 'from 章节表').Count -ne 1) {
    throw '当前条目解析应只执行一次章节表查询，不得增加逐行或重复往返'
}
if ($currentEntryBody.Contains('当前树节点缺少可核对的条目序号') -or
    $currentEntryBody.Contains('当前树节点缺少可核对的条目序号或编号') -or
    $feature.Contains('当前树节点缺少可核对的条目序号或编号')) {
    throw '真实宿主树节点未暴露 Tag 序号/编号时仍会被直接拒绝'
}

foreach ($marker in @('String.Equals(targetConn.Database, candidate.DatabaseName',
    'GetProjectConnectionIdentity(targetConn)', 'TryLoadSmartSourceRowFromConnection')) {
    if (-not $crossDbBody.Contains($marker)) {
        throw "当前项目历史右键绑定缺少同库完整源行回读：$marker"
    }
}

if (-not $nameMatch.Contains('row_number() over(partition by 定额编号, 工程或费用项目名称, 单位') -or
    -not $nameMatch.Contains('case when 单价 is not null and 单价<>0 then 0 else 1 end, 定额序号')) {
    throw '当前项目存在同身份 ZLF 时未优先选择非零单价完整行'
}
foreach ($marker in @('ResolveSmartPreviewUnitPrice',
    'LearnedUnitPrice = ResolveSmartPreviewUnitPrice(target, currentQuota)',
    'if (IsContextSensitiveLearningCode(target.Code) && item.LearnedUnitPrice == 0m)')) {
    if (-not $smart.Contains($marker)) { throw "ZLF 当前项目单价未在预览阶段提前生效：$marker" }
}
foreach ($marker in @('LoadQuotaUnitPriceForLearning',
    'select 单价 from 定额输入 where 定额序号=@id',
    'link.UnitPrice = FilterLearningTargetUnitPrice(quotaCode,',
    'IsContextSensitiveLearningCode(quotaCode)')) {
    if (-not $excelLink.Contains($marker)) {
        throw "右键绑定未按定额序号从当前项目行回读单价：$marker"
    }
}

foreach ($forbidden in @(
    'EntryByQuota',
    'EntryBySignatureQuota',
    'ResolveSmartTargetEntryCombinations',
    'prefixVotes',
    'preferredPrefixes'
)) {
    if ($smart.Contains($forbidden)) {
        throw "Obsolete entry inference remains: $forbidden"
    }
}

$dllDir = Split-Path -Parent $dll
foreach ($dependency in @('NPOI.dll', 'NPOI.OpenXmlFormats.dll', 'NPOI.OpenXml4Net.dll', 'NPOI.OOXML.dll', 'ICSharpCode.SharpZipLib.dll')) {
    $dependencyPath = Join-Path $dllDir $dependency
    if (Test-Path -LiteralPath $dependencyPath) { [void][Reflection.Assembly]::LoadFrom($dependencyPath) }
}
$formType = [Reflection.Assembly]::LoadFrom($dll).GetType('RecoNet.FormPanel', $true)
$nested = [Reflection.BindingFlags]'Public,NonPublic'
$flags = [Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$itemType = $formType.GetNestedType('FillPreviewItem', $nested)
$buildL2 = $formType.GetMethod('BuildSmartFillL2Row', $flags)
$hasConstructedIdentity = $formType.GetMethod('HasSmartFillConstructedIdentity', $flags)
$mergePreviewTargetGroup = $formType.GetMethod('MergePreviewTargetGroup', $flags)
$resolveSourceDatabase = $formType.GetMethod('ResolveSmartSourceDatabaseName', $flags)
$shouldLoadCurrentQuotaTarget = $formType.GetMethod('ShouldLoadCurrentSmartQuotaTarget', $flags)
$smartTargetType = $formType.GetNestedType('SmartBoxTarget', $nested)
$projectQuotaType = $formType.GetNestedType('ProjectQuota', $nested)
$resolveSmartPreviewUnitPrice = $formType.GetMethod('ResolveSmartPreviewUnitPrice', $flags)
$filterLearningTargetUnitPrice = $formType.GetMethod('FilterLearningTargetUnitPrice', $flags)
$resolveLearningTargetKind = $formType.GetMethod('ResolveLearningTargetKind', $flags)
$panelType = $formType.GetNestedType('TemplateFillPanel', $nested)
$resolveTreeNode = if ($null -eq $panelType) { $null } else { $panelType.GetMethod('ResolveSmartHostTreeNode', $flags) }
$isEditableGrid = if ($null -eq $panelType) { $null } else { $panelType.GetMethod('IsEditableAgentQuotaGrid', $flags) }
$isOptionalTreeSequenceConsistent = if ($null -eq $panelType) { $null } else { $panelType.GetMethod('IsOptionalSmartTreeSequenceConsistent', $flags) }
$describeSfConflict = if ($null -eq $panelType) { $null } else { $panelType.GetMethod('DescribeSmartSfEntryConflict', $flags) }
if ($null -eq $itemType -or $null -eq $buildL2 -or $null -eq $hasConstructedIdentity -or
    $null -eq $mergePreviewTargetGroup -or
    $null -eq $resolveSourceDatabase -or
    $null -eq $shouldLoadCurrentQuotaTarget -or
    $null -eq $smartTargetType -or $null -eq $projectQuotaType -or
    $null -eq $resolveSmartPreviewUnitPrice -or $null -eq $filterLearningTargetUnitPrice -or
    $null -eq $resolveLearningTargetKind -or
    $null -eq $resolveTreeNode -or $null -eq $isEditableGrid -or
    $null -eq $isOptionalTreeSequenceConsistent -or $null -eq $describeSfConflict) {
    throw '缺少推荐定额统一构造写入的可测试行为入口'
}

foreach ($auxiliaryCode in @('ZLF', 'SH', 'SF', 'LF', 'SQ', 'YF', 'TLF', 'GF', 'JF', 'XGT1')) {
    if ([decimal]$filterLearningTargetUnitPrice.Invoke($null, @($auxiliaryCode, [decimal]58.5)) -ne [decimal]58.5) {
        throw "辅助码 $auxiliaryCode 的软件行单价未进入学习链"
    }
}
foreach ($numberedTarget in @('PY-393', 'BC00-4', '1009001003', '1009001003*1.02')) {
    if ([decimal]$filterLearningTargetUnitPrice.Invoke($null, @($numberedTarget, [decimal]58.5)) -ne 0) {
        throw "正式定额或编号材料 $numberedTarget 被错误学习单价"
    }
}
Write-Host 'PASS 单价学习仅限辅助码，正式定额和编号材料统一归零'
if ($resolveLearningTargetKind.Invoke($null, @('', 'PY-393')) -ne 'quota' -or
    $resolveLearningTargetKind.Invoke($null, @('', '1009001003')) -ne 'material' -or
    $resolveLearningTargetKind.Invoke($null, @('', '1009001003*1.02')) -ne 'material') {
    throw '右键绑定完整关系未正确区分正式定额与编号材料'
}
Write-Host 'PASS 右键绑定完整关系保留目标类型'

$zlfTarget = [Activator]::CreateInstance($smartTargetType, $true).PSObject.BaseObject
$smartTargetType.GetField('Kind', $flags).SetValue($zlfTarget, 'quota')
$smartTargetType.GetField('Code', $flags).SetValue($zlfTarget, 'ZLF')
$smartTargetType.GetField('Name', $flags).SetValue($zlfTarget, 'Ф150×12mm CPVC管')
$smartTargetType.GetField('Unit', $flags).SetValue($zlfTarget, 'm')
$zlfArgs = New-Object 'object[]' 1
$zlfArgs[0] = $zlfTarget
if (-not [bool]$shouldLoadCurrentQuotaTarget.Invoke($null, $zlfArgs)) {
    throw '完整名称+单位的 ZLF 辅助码仍被排除在当前项目定额查找之外'
}
$currentZlf = [Activator]::CreateInstance($projectQuotaType, $true).PSObject.BaseObject
$projectQuotaType.GetField('Code', $flags).SetValue($currentZlf, 'ZLF')
$projectQuotaType.GetField('Name', $flags).SetValue($currentZlf, 'Ф150×12mm CPVC管')
$projectQuotaType.GetField('Unit', $flags).SetValue($currentZlf, 'm')
$projectQuotaType.GetField('UnitPrice', $flags).SetValue($currentZlf, [decimal]58.5)
$smartTargetType.GetField('UnitPrice', $flags).SetValue($zlfTarget, [decimal]0)
$previewPriceArgs = New-Object 'object[]' 2
$previewPriceArgs[0] = $zlfTarget
$previewPriceArgs[1] = $currentZlf
if ([decimal]$resolveSmartPreviewUnitPrice.Invoke($null, $previewPriceArgs) -ne [decimal]58.5) {
    throw 'ZLF 预览未优先显示当前项目完整身份行的非零单价'
}
$projectQuotaType.GetField('UnitPrice', $flags).SetValue($currentZlf, [decimal]0)
$smartTargetType.GetField('UnitPrice', $flags).SetValue($zlfTarget, [decimal]12.5)
if ([decimal]$resolveSmartPreviewUnitPrice.Invoke($null, $previewPriceArgs) -ne [decimal]12.5) {
    throw 'ZLF 当前项目行为 0 时未回退学习库非零单价'
}
$formalTarget = [Activator]::CreateInstance($smartTargetType, $true).PSObject.BaseObject
$smartTargetType.GetField('Kind', $flags).SetValue($formalTarget, 'quota')
$smartTargetType.GetField('Code', $flags).SetValue($formalTarget, 'PY-393')
$smartTargetType.GetField('UnitPrice', $flags).SetValue($formalTarget, [decimal]999)
$projectQuotaType.GetField('UnitPrice', $flags).SetValue($currentZlf, [decimal]888)
$formalPriceArgs = New-Object 'object[]' 2
$formalPriceArgs[0] = $formalTarget
$formalPriceArgs[1] = $currentZlf
if ([decimal]$resolveSmartPreviewUnitPrice.Invoke($null, $formalPriceArgs) -ne 0) {
    throw '正式定额错误读取了历史或当前项目的学习单价'
}
$smartTargetType.GetField('Unit', $flags).SetValue($zlfTarget, '')
if ([bool]$shouldLoadCurrentQuotaTarget.Invoke($null, $zlfArgs)) {
    throw '缺少单位的 ZLF 辅助码被误当成可精确匹配的当前项目定额'
}
Write-Host 'PASS ZLF 按完整名称+单位进入当前项目查找，预览优先非零单价，缺失身份时拒绝'

if ($resolveSourceDatabase.Invoke($null, @('old-server|OldDb', 'current-server|CurrentDb')) -ne 'CurrentDb' -or
    $resolveSourceDatabase.Invoke($null, @('old-server|OldDb', '')) -ne 'OldDb' -or
    $resolveSourceDatabase.Invoke($null, @('LegacyDb', '')) -ne 'LegacyDb') {
    throw 'BindingLog 项目身份没有被稳定解析为真实来源数据库名'
}
Write-Host 'PASS 历史/当前来源端点身份解析为真实数据库名'

$hostTree = New-Object System.Windows.Forms.TreeView
$selectedNode = New-Object System.Windows.Forms.TreeNode '界面实际选中条目'
$staleCurrNode = New-Object System.Windows.Forms.TreeNode '宿主旧 CurrNode'
[void]$hostTree.Nodes.Add($selectedNode)
$hostTree.SelectedNode = $selectedNode
$resolvedNode = $resolveTreeNode.Invoke($null, @($hostTree.PSObject.BaseObject, $staleCurrNode.PSObject.BaseObject))
if (-not [Object]::ReferenceEquals($resolvedNode, $selectedNode.PSObject.BaseObject)) {
    throw '推荐定额没有优先采用章节树当前真实选中节点'
}
$resolvedFallback = $resolveTreeNode.Invoke($null, @($null, $staleCurrNode.PSObject.BaseObject))
if (-not [Object]::ReferenceEquals($resolvedFallback, $staleCurrNode.PSObject.BaseObject)) {
    throw '章节树不可用时没有回退宿主 CurrNode'
}
$hostTree.SelectedNode = $null
$resolvedTransientFallback = $resolveTreeNode.Invoke($null, @($hostTree.PSObject.BaseObject, $staleCurrNode.PSObject.BaseObject))
if (-not [Object]::ReferenceEquals($resolvedTransientFallback, $staleCurrNode.PSObject.BaseObject)) {
    throw '章节树 SelectedNode 暂时为空时没有回退宿主 CurrNode'
}
$hostTree.Dispose()
Write-Host 'PASS 当前条目优先采用章节树 SelectedNode，空选中时安全回退 CurrNode'

$agentGrid = New-Object System.Windows.Forms.DataGridView
$agentGrid.AllowUserToAddRows = $false
$quotaColumn = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
$quotaColumn.Name = '定额编号'
$quotaColumn.HeaderText = '定额编号'
$quotaColumn.ReadOnly = $true
[void]$agentGrid.Columns.Add($quotaColumn)
[void]$agentGrid.Rows.Add()
$gridArgs = New-Object 'object[]' 1
$gridArgs[0] = $agentGrid.PSObject.BaseObject
if (-not [bool]$isEditableGrid.Invoke($null, $gridArgs)) {
    throw '宿主自管末尾空白行被误判为不可输入定额'
}
$agentGrid.Rows.Clear()
if ([bool]$isEditableGrid.Invoke($null, $gridArgs)) {
    throw '没有可定位行的定额表被误判为可输入'
}
$agentGrid.Dispose()
Write-Host 'PASS 宿主自管空白行与实际定额粘贴路径采用同一可写判定'

if (-not [bool]$isOptionalTreeSequenceConsistent.Invoke($null, @('', [long]8123)) -or
    -not [bool]$isOptionalTreeSequenceConsistent.Invoke($null, @($null, [long]8123)) -or
    -not [bool]$isOptionalTreeSequenceConsistent.Invoke($null, @('8123', [long]8123))) {
    throw 'Tag 未暴露条目序号或序号相同时，当前条目未被接受'
}
if ([bool]$isOptionalTreeSequenceConsistent.Invoke($null, @('8124', [long]8123)) -or
    [bool]$isOptionalTreeSequenceConsistent.Invoke($null, @('invalid', [long]8123))) {
    throw '树节点明确提供的序号与项目章节表不一致时没有拒绝'
}
Write-Host 'PASS 真实宿主 Tag 缺序号时允许按界面编号解析，明确序号仍作交叉校验'

function New-Item([string]$Code, [string]$Name, [string]$Unit, [string]$Quantity) {
    $item = [Activator]::CreateInstance($itemType).PSObject.BaseObject
    foreach ($pair in @{ TargetKind='quota'; QuotaCode=$Code; SourceName=$Name; Unit=$Unit; QuantityText=$Quantity }.GetEnumerator()) {
        $itemType.GetField($pair.Key, $flags).SetValue($item, $pair.Value)
    }
    return $item
}
function New-CompleteRow([string]$Code, [string]$Name, [string]$Unit, [string]$Quantity) {
    $row = [Collections.Generic.Dictionary[string,object]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($field in @(
        '定额编号','工程或费用项目名称','单位','总概算序号','条目序号','顺号',
        '工程数量输入','工程数量','单价','基价','工费','料费','机费','人工费','材料费','机械费',
        '设备费','主材费','价差','定额调整','单重','合重')) { $row[$field] = [decimal]0 }
    $row['定额编号'] = $Code
    $row['工程或费用项目名称'] = $Name
    $row['单位'] = $Unit
    $row['工程数量输入'] = $Quantity
    $row['工程数量'] = [decimal]::Parse($Quantity, [Globalization.CultureInfo]::InvariantCulture)
    $row['定额序号'] = [long]99
    $row['额外宿主字段'] = '保留'
    return $row
}
$aux = New-Item 'SH' '弃土消纳费' 'm3' '3'
$itemType.GetField('LearnedUnitPrice', $flags).SetValue($aux, [decimal]12.5)
$structural = New-CompleteRow 'LY-9' '结构模板行' '100m3' '1'
$l2 = $buildL2.Invoke($null, @($structural, $aux))
if ($l2['定额编号'] -ne 'SH' -or $l2['工程或费用项目名称'] -ne '弃土消纳费' -or
    $l2['单位'] -ne 'm3' -or $l2['工程数量输入'] -ne '3' -or [decimal]$l2['工程数量'] -ne 3 -or
    [decimal]$l2['单价'] -ne [decimal]12.5 -or
    [decimal]$l2['基价'] -ne 0 -or $l2.ContainsKey('定额序号') -or $l2['额外宿主字段'] -ne '保留') {
    throw 'L2 未在完整宿主结构副本上正确覆盖辅助码身份、数量和学习单价'
}
Write-Host 'PASS L2 保留宿主行结构并覆盖辅助码身份、数量和学习单价'

$formal = New-Item 'PY-415' '充填式注浆 Φ560×33.2mm' '10m' '252/10'
$itemType.GetField('LearnedUnitPrice', $flags).SetValue($formal, [decimal]0)
$identityArgs = New-Object 'object[]' 1
$identityArgs[0] = $formal
if (-not [bool]$hasConstructedIdentity.Invoke($null, $identityArgs)) {
    throw '有完整名称和单位的正式定额未获准进入结构模板构造'
}
$missingName = New-Item 'PY-415' '' '10m' '252/10'
$identityArgs[0] = $missingName
if ([bool]$hasConstructedIdentity.Invoke($null, $identityArgs)) {
    throw '缺少名称的正式定额未阻断整组构造写入'
}
$missingUnit = New-Item 'PY-415' '充填式注浆 Φ560×33.2mm' '' '252/10'
$identityArgs[0] = $missingUnit
if ([bool]$hasConstructedIdentity.Invoke($null, $identityArgs)) {
    throw '缺少单位的正式定额未阻断整组构造写入'
}
$formalRow = $buildL2.Invoke($null, @($structural, $formal))
if ($formalRow['定额编号'] -ne 'PY-415' -or
    $formalRow['工程或费用项目名称'] -ne '充填式注浆 Φ560×33.2mm' -or
    $formalRow['单位'] -ne '10m' -or $formalRow['工程数量输入'] -ne '252/10' -or
    [decimal]$formalRow['工程数量'] -ne [decimal]25.2 -or [decimal]$formalRow['单价'] -ne 0 -or
    [decimal]$formalRow['基价'] -ne 0 -or [decimal]$formalRow['合重'] -ne 0 -or
    $formalRow.ContainsKey('定额序号') -or $formalRow['额外宿主字段'] -ne '保留') {
    throw '无完整源行的正式定额未按编号、名称、单位、数量和零价构造完整业务行'
}
Write-Host 'PASS 无完整源行的正式定额统一构造完整业务行，缺名称或单位时阻断'

$ordinary = New-Item 'PY-415' '充填式注浆 Φ560×33.2mm' '10m' '252/10'
$itemType.GetField('ChosenItemNo', $flags).SetValue($ordinary, '0821-01-04-05-01')
$sf = New-Item 'SF' '设备购置费' '元' '1'
$itemType.GetField('ChosenItemNo', $flags).SetValue($sf, '0821-01-04-05-02')
$existingType = $mergePreviewTargetGroup.GetParameters()[0].ParameterType
$additionType = $mergePreviewTargetGroup.GetParameters()[1].ParameterType
$existing = [Activator]::CreateInstance($existingType).PSObject.BaseObject
$additions = [Activator]::CreateInstance($additionType).PSObject.BaseObject
[void]$existing.Add($ordinary)
[void]$additions.Add($sf)
[void]$additions.Add($sf)
$mergeArgs = New-Object 'object[]' 2
$mergeArgs[0] = $existing
$mergeArgs[1] = $additions
$merged = $mergePreviewTargetGroup.Invoke($null, $mergeArgs)
if ($merged.Count -ne 2 -or
    @($merged | Where-Object { $itemType.GetField('QuotaCode', $flags).GetValue($_) -eq 'PY-415' }).Count -ne 1 -or
    @($merged | Where-Object { $itemType.GetField('QuotaCode', $flags).GetValue($_) -eq 'SF' }).Count -ne 1) {
    throw 'SF 补充未保留普通定额、加入 SF 或去除同身份重复项'
}
Write-Host 'PASS SF 补充保留原组件并加入唯一设备费目标'

$equipmentConflict = $describeSfConflict.Invoke($null, @($true, $true, $true, $true))
$missingSiblingConflict = $describeSfConflict.Invoke($null, @($false, $true, $true, $false))
$resolvedSiblingConflict = $describeSfConflict.Invoke($null, @($false, $true, $true, $true))
$ordinaryOnlyConflict = $describeSfConflict.Invoke($null, @($false, $false, $true, $false))
if ($equipmentConflict -ne '设备购置费条目只接受 SF，所选组件整组未写入' -or
    $missingSiblingConflict -ne '未找到唯一同级设备购置费条目，所选组件整组未写入' -or
    $resolvedSiblingConflict -ne '' -or $ordinaryOnlyConflict -ne '') {
    throw 'SF 双向约束纯判定与既有阻断文案不一致'
}
Write-Host 'PASS SF 替换与补充候选可共用双向约束纯判定'

Write-Host 'PASS SmartFill constructed-row write safety contract'
