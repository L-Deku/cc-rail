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

$requiredPanel = @(
    'SmartPreviewContext',
    'smartPreviewReady',
    'currentEntryWritable',
    'RefreshApplyEnabled',
    'InvalidateSmartPreview',
    'grid.ClearSelection()',
    'DataGridViewSelectionMode.FullRowSelect',
    'GetSelectedSmartTargetRows'
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
    'NativeInsertState',
    'Submitted',
    'Confirming',
    'PartiallyConfirmed',
    'Indeterminate',
    'CompensationFailed',
    'DateTime.Now.AddSeconds(10)',
    'WaitAgentUiIdle(250)',
    'stableCount >= 3',
    'IsSmartFillSourceIdentityMatch',
    'ProjectConnection = conn',
    'ProjectConnectionIdentity = GetProjectConnectionIdentity(conn)'
)
foreach ($marker in $requiredFeature) {
    if (-not $feature.Contains($marker)) {
        throw "Missing smart apply safety marker: $marker"
    }
}
if ($feature.IndexOf('else if (IsContextSensitiveLearningCode(item.QuotaCode))', [StringComparison]::Ordinal) -lt 0 -or
    $feature.IndexOf('else if (IsContextSensitiveLearningCode(item.QuotaCode))', [StringComparison]::Ordinal) -gt
    $feature.IndexOf('plan.Layer = SmartFillWriteLayer.L3;', [StringComparison]::Ordinal)) {
    throw '辅助码没有在 L3 正式编号原生输入之前固定分流到 L2'
}
$crossDbStart = $smart.IndexOf('private static Dictionary<string, object> LoadCrossDbQuotaRow', [StringComparison]::Ordinal)
$crossDbEnd = $smart.IndexOf('private static List<SmartMapCandidateScore> RankSmartMapEntries', $crossDbStart, [StringComparison]::Ordinal)
if ($crossDbStart -lt 0 -or $crossDbEnd -le $crossDbStart) {
    throw '缺少跨库完整源行加载入口'
}
$crossDbBody = $smart.Substring($crossDbStart, $crossDbEnd - $crossDbStart)
foreach ($marker in @(
    'candidates.OrderByDescending(value => value.BindingId)',
    'NormalizeForSignature(actualName)',
    'NormalizeForSignature(actualUnit)',
    'IsSmartFillSourceIdentityMatch(item, values)) continue;',
    'return values;'
)) {
    if (-not $crossDbBody.Contains($marker)) {
        throw "跨库源行缺少同编号异义拒绝或候选回退门禁：$marker"
    }
}
if ($crossDbBody.IndexOf('IsSmartFillSourceIdentityMatch(item, values)) continue;', [StringComparison]::Ordinal) -gt
    $crossDbBody.IndexOf('return values;', [StringComparison]::Ordinal)) {
    throw '跨库源行在完整身份核对前已经返回，无法安全回退到旧候选'
}
foreach ($marker in @('approvedEntrySequence', '确认后项目或条目已变化',
    'currentSmartEntry.EntrySequence != approvedEntrySequence')) {
    if (-not $panel.Contains($marker)) { throw "用户确认后缺少临写入前项目/条目二次核对：$marker" }
}
if (-not $feature.Contains('out bool succeeded') -or
    -not $feature.Contains('succeeded = true;') -or
    $panel.Contains('smartResult.StartsWith(') -or
    -not $panel.Contains('if (smartSucceeded) InvalidateSmartPreview();')) {
    throw '推荐写入仍依赖成功提示字符串判断预览是否失效'
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
if ($currentEntryBody.Contains('当前树节点缺少可核对的条目序号')) {
    throw '树节点未暴露条目序号时仍会被直接拒绝'
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
$planType = $formType.GetNestedType('PreparedSmartFillItem', $nested)
$recordType = $formType.GetNestedType('SmartNativeInsertRecord', $nested)
$classify = $formType.GetMethod('ClassifySmartNativeRows', $flags)
$buildL2 = $formType.GetMethod('BuildSmartFillL2Row', $flags)
$panelType = $formType.GetNestedType('TemplateFillPanel', $nested)
$resolveTreeNode = if ($null -eq $panelType) { $null } else { $panelType.GetMethod('ResolveSmartHostTreeNode', $flags) }
$isEditableGrid = if ($null -eq $panelType) { $null } else { $panelType.GetMethod('IsEditableAgentQuotaGrid', $flags) }
$isOptionalTreeSequenceConsistent = if ($null -eq $panelType) { $null } else { $panelType.GetMethod('IsOptionalSmartTreeSequenceConsistent', $flags) }
if ($null -eq $itemType -or $null -eq $planType -or $null -eq $recordType -or
    $null -eq $classify -or $null -eq $buildL2 -or $null -eq $resolveTreeNode -or $null -eq $isEditableGrid -or
    $null -eq $isOptionalTreeSequenceConsistent) {
    throw '缺少 L2 构造或 L3 结构化确认的可测试行为入口'
}

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
    throw '条目编号唯一命中时，树节点缺少序号或序号相同未被接受'
}
if ([bool]$isOptionalTreeSequenceConsistent.Invoke($null, @('8124', [long]8123)) -or
    [bool]$isOptionalTreeSequenceConsistent.Invoke($null, @('invalid', [long]8123))) {
    throw '树节点序号与项目章节表不一致时没有拒绝'
}
Write-Host 'PASS 树节点序号为可选交叉校验，条目编号唯一命中即可解析'

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
function New-Record($Item, [int]$ExpectedCount) {
    $plan = [Activator]::CreateInstance($planType, $true).PSObject.BaseObject
    $planType.GetField('Item', $flags).SetValue($plan, $Item)
    $record = [Activator]::CreateInstance($recordType, $true).PSObject.BaseObject
    $recordType.GetField('ExpectedCount', $flags).SetValue($record, $ExpectedCount)
    [void]$recordType.GetField('Items', $flags).GetValue($record).Add($plan)
    return $record
}
function Invoke-Classify($Record, $Rows) {
    $args = New-Object 'object[]' 4
    $args[0] = $Record
    $args[1] = $Rows
    $args[2] = $null
    $args[3] = $null
    return [pscustomobject]@{ State=[string]$classify.Invoke($null, $args); Owned=$args[2]; Unowned=$args[3] }
}

$official = New-Item 'LY-1' '测试定额' 'm3' '2'
$record = New-Record $official 1
$rowDictionaryType = $classify.GetParameters()[1].ParameterType
$rows = [Activator]::CreateInstance($rowDictionaryType).PSObject.BaseObject
$rows.Add([long]101, (New-CompleteRow 'LY-1' '测试定额' 'm3' '2'))
$confirmed = Invoke-Classify $record $rows
if ($confirmed.State -ne 'Confirmed' -or $confirmed.Owned.Count -ne 1 -or $confirmed.Unowned.Count -ne 0) {
    throw 'L3 完整身份与数量相符的单行未进入 Confirmed'
}
$recordType.GetField('ExpectedCount', $flags).SetValue($record, 2)
$partial = Invoke-Classify $record $rows
if ($partial.State -ne 'PartiallyConfirmed' -or $partial.Owned.Count -ne 1) {
    throw 'L3 少于预期的新行未进入 PartiallyConfirmed'
}
$recordType.GetField('ExpectedCount', $flags).SetValue($record, 1)
$rows[101]['工程或费用项目名称'] = ''
$shell = Invoke-Classify $record $rows
if ($shell.State -ne 'Indeterminate' -or $shell.Owned.Count -ne 1) {
    throw '只有编号和数量的原生壳行被误认为完整写入'
}
$rows.Clear()
$rows.Add([long]102, (New-CompleteRow 'OTHER-1' '其他定额' 'm3' '2'))
$unowned = Invoke-Classify $record $rows
if ($unowned.State -ne 'Indeterminate' -or $unowned.Unowned.Count -ne 1 -or $unowned.Owned.Count -ne 0) {
    throw '无法归属的同期新行未与本批可补偿 ID 隔离'
}
$rows.Clear()
$failed = Invoke-Classify $record $rows
if ($failed.State -ne 'Failed') { throw '未检测到新行时未进入 Failed' }
Write-Host 'PASS L3 状态机按完整身份、数量和归属区分 Confirmed/Partial/Indeterminate/Failed'

$aux = New-Item 'SH' '弃土消纳费' 'm3' '3'
$itemType.GetField('LearnedUnitPrice', $flags).SetValue($aux, [decimal]12.5)
$structural = New-CompleteRow 'LY-9' '结构模板行' '100m3' '1'
$l2 = $buildL2.Invoke($null, @($structural, $aux))
if ($l2['定额编号'] -ne 'SH' -or $l2['工程或费用项目名称'] -ne '弃土消纳费' -or
    $l2['单位'] -ne 'm3' -or [decimal]$l2['单价'] -ne [decimal]12.5 -or
    [decimal]$l2['基价'] -ne 0 -or $l2.ContainsKey('定额序号') -or $l2['额外宿主字段'] -ne '保留') {
    throw 'L2 未在完整宿主结构副本上正确覆盖辅助码身份、数量和学习单价'
}
Write-Host 'PASS L2 保留宿主行结构并覆盖辅助码身份、数量和学习单价'

Write-Host 'PASS SmartFill manual-entry safety contract'
