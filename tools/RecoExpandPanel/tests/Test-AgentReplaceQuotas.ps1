$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$dll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL)) {
    $env:RECO_EXPAND_DLL
} else {
    Join-Path (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..')).Path 'RecoQuotaRecommend\bin\RecoExpandPanel.dll'
}

Add-Type -Path $dll
$formType = [RecoNet.FormPanel]
$flags = [System.Reflection.BindingFlags]'NonPublic,Static,Instance,Public'

$failures = New-Object System.Collections.Generic.List[string]

function Assert-True([bool]$condition, [string]$message) {
    if (-not $condition) {
        $script:failures.Add($message)
    }
}

# 1) AgentCommand 上的 replace_quotas 字段
$commandType = $formType.GetNestedType('AgentCommand', $flags)
Assert-True ($null -ne $commandType) 'AgentCommand nested type missing'
Assert-True ($null -ne $commandType.GetField('FromCodes', $flags)) 'AgentCommand.FromCodes missing'
Assert-True ($null -ne $commandType.GetField('ToQuotas', $flags)) 'AgentCommand.ToQuotas missing'

# 2) AgentInsertGroup 的定位插入字段
$groupType = $formType.GetNestedType('AgentInsertGroup', $flags)
Assert-True ($null -ne $groupType) 'AgentInsertGroup nested type missing'
foreach ($fieldName in @('UnitId', 'ItemSequence', 'AfterOrderNo')) {
    Assert-True ($null -ne $groupType.GetField($fieldName, $flags)) "AgentInsertGroup.$fieldName missing"
}

# 3) 新增/抽取出来的方法都在
foreach ($methodName in @(
        'BuildReplaceQuotasPlan',
        'ReorderAgentInsertedRows',
        'AgentQuotaCodeMatchWithSuffix',
        'LoadAgentUnitOptions',
        'LoadAgentItemOptions',
        'ResolveAgentScopeRows',
        'ResolveAgentNameRows',
        'AgentQuotaCodeKind',
        'ShowAgentPanelWindow')) {
    Assert-True ($null -ne $formType.GetMethod($methodName, $flags)) "method $methodName missing"
}

# 4) 点选面板与选择对话框类型存在
foreach ($typeName in @('AgentPanelWindow', 'AgentPickerDialog', 'AgentUnitOption', 'AgentItemOption', 'AgentTargetEntry')) {
    Assert-True ($null -ne $formType.GetNestedType($typeName, $flags)) "nested type $typeName missing"
}

# 4a) 目标清单支持编号与名称混用
$panelType = $formType.GetNestedType('AgentPanelWindow', $flags)
Assert-True ($null -ne $panelType.GetMethod('TakeTargetsFromHostGrid', $flags)) 'TakeTargetsFromHostGrid missing'
Assert-True ($null -ne $panelType.GetField('targetEntries', $flags)) 'targetEntries field missing'
foreach ($m in @('TargetCodes', 'TargetNames', 'ApplyTarget')) {
    Assert-True ($null -ne $panelType.GetMethod($m, $flags)) "panel method $m missing"
}

# 4b) 编号分类：材料编号 / 补充定额(手填单价) / 普通定额
$kindMethod = $formType.GetMethod('AgentQuotaCodeKind', $flags)
function Invoke-Kind([string]$code) {
    $invokeArgs = New-Object object[] 1
    $invokeArgs[0] = $code
    return [string]$kindMethod.Invoke($null, $invokeArgs)
}

Assert-True ((Invoke-Kind '1294861') -eq '材料') 'long pure-digit code should be classified as 材料'
Assert-True ((Invoke-Kind '0305') -eq '定额') 'short pure-digit code should not be classified as 材料'
foreach ($supplement in @('SF', 'SH', 'SQ', 'ZLF', 'LF', 'TLF')) {
    Assert-True ((Invoke-Kind $supplement) -eq '补充') "$supplement should be classified as 补充"
}
Assert-True ((Invoke-Kind 'SF*9') -eq '补充') 'supplement code with suffix should still be 补充'
Assert-True ((Invoke-Kind 'LY-21') -eq '定额') 'normal quota code should be classified as 定额'

# 5) 定额编号匹配规则：整号相同或整号+乘除后缀才算命中
$matchMethod = $formType.GetMethod('AgentQuotaCodeMatchWithSuffix', $flags)

function Invoke-Match([string]$code, [string[]]$fromCodes) {
    $list = New-Object 'System.Collections.Generic.List[string]'
    foreach ($item in $fromCodes) { $list.Add($item) }
    $invokeArgs = New-Object object[] 3
    $invokeArgs[0] = $code
    $invokeArgs[1] = $list.PSObject.BaseObject
    $invokeArgs[2] = $null
    $matched = [bool]$matchMethod.Invoke($null, $invokeArgs)
    return New-Object psobject -Property @{ Matched = $matched; Suffix = [string]$invokeArgs[2] }
}

$r = Invoke-Match 'LY-21' @('LY-21')
Assert-True ($r.Matched -and $r.Suffix -eq '') 'exact code should match with empty suffix'

$r = Invoke-Match 'LY-21*9' @('LY-21')
Assert-True ($r.Matched -and $r.Suffix -eq '*9') 'code with multiply suffix should match and report suffix'

$r = Invoke-Match 'LY-21/1.1' @('LY-21')
Assert-True ($r.Matched -and $r.Suffix -eq '/1.1') 'code with divide suffix should match and report suffix'

$r = Invoke-Match 'LY-210' @('LY-21')
Assert-True (-not $r.Matched) 'longer code without operator separator must not match'

$r = Invoke-Match 'LY-22' @('LY-21', 'LY-22')
Assert-True ($r.Matched -and $r.Suffix -eq '') 'multi from-code list should match the second entry'

$r = Invoke-Match 'ZLF-04' @('LY-21', 'LY-22')
Assert-True (-not $r.Matched) 'unrelated code must not match'

# 6) Describe() 能描述 1 对 1 / 多对 1 / 1 对多
$quotaInputType = $formType.GetNestedType('AgentQuotaInput', $flags)
$describeMethod = $commandType.GetMethod('Describe', $flags)

function New-QuotaInput([string]$code, [string]$quantity) {
    $quota = [Activator]::CreateInstance($quotaInputType, $true)
    $quotaInputType.GetField('Code', $flags).SetValue($quota, $code)
    $quotaInputType.GetField('Quantity', $flags).SetValue($quota, $quantity)
    return $quota.PSObject.BaseObject
}

function Describe-Replace([string[]]$fromCodes, $toQuotas) {
    $command = [Activator]::CreateInstance($commandType, $true)
    $commandType.GetField('Type', $flags).SetValue($command, 'replace_quotas')
    $fromList = $commandType.GetField('FromCodes', $flags).GetValue($command)
    foreach ($code in $fromCodes) { $fromList.Add($code) }
    $toList = $commandType.GetField('ToQuotas', $flags).GetValue($command)
    foreach ($quota in $toQuotas) { $toList.Add($quota.PSObject.BaseObject) }
    return [string]$describeMethod.Invoke($command.PSObject.BaseObject, @())
}

$text = Describe-Replace @('LY-21') @((New-QuotaInput 'QY-100' ''))
Assert-True ($text.Contains('LY-21') -and $text.Contains('QY-100')) "1-to-1 describe wrong: $text"

$text = Describe-Replace @('LY-21', 'LY-22') @((New-QuotaInput 'QY-100' '150'))
Assert-True ($text.Contains('LY-21') -and $text.Contains('LY-22') -and $text.Contains('QY-100(150)')) "many-to-1 describe wrong: $text"

$text = Describe-Replace @('LY-21') @((New-QuotaInput 'LY-22' '100'), (New-QuotaInput 'LY-23' '80'))
Assert-True ($text.Contains('LY-22(100)') -and $text.Contains('LY-23(80)')) "1-to-many describe wrong: $text"

# 7) AI 通道不得产出 replace_quotas（一对多数量不可控，只走点选界面）
$validate = $formType.GetMethod('ValidateAgentCommandShape', $flags)
if ($null -ne $validate) {
    $command = [Activator]::CreateInstance($commandType, $true)
    $commandType.GetField('Type', $flags).SetValue($command, 'replace_quotas')
    $invokeArgs = New-Object object[] 1
    $invokeArgs[0] = $command
    $result = $validate.Invoke($null, $invokeArgs)
    Assert-True ($null -ne $result) 'ValidateAgentCommandShape should reject replace_quotas from the AI channel'
}

# 8) 新增行改为事务内 SQL 插入（结构行克隆），不再走剪贴板粘贴
foreach ($methodName in @(
        'ExecuteAgentInsertGroupSql',
        'ResolveAgentInsertIdentity',
        'ResolveAgentInsertQuotaIdentities',
        'ResolveAgentInsertUnitIds',
        'ResolveAgentItemSequence',
        'LoadAgentStructuralRow',
        'AllocateAgentInsertOrders')) {
    Assert-True ($null -ne $formType.GetMethod($methodName, $flags)) "method $methodName missing"
}

$quotaInput = $formType.GetNestedType('AgentQuotaInput', $flags)
foreach ($f in @('Unit', 'SameCodeQuotaSequence')) {
    Assert-True ($null -ne $quotaInput.GetField($f, $flags)) "AgentQuotaInput.$f missing"
}
Assert-True ($null -ne $groupType.GetField('AnchorQuotaSequence', $flags)) 'AgentInsertGroup.AnchorQuotaSequence missing'

# 9) 源码级：智能指令的执行路径不再调用剪贴板粘贴，且跨单元硬拒已解除
$executorPath = Join-Path $PSScriptRoot '..\AgentExecutor.cs'
if (Test-Path -LiteralPath $executorPath) {
    $executor = [System.IO.File]::ReadAllText((Resolve-Path -LiteralPath $executorPath).Path, [System.Text.Encoding]::UTF8)
    $planStart = $executor.IndexOf('private static string ExecuteAgentPlan(')
    Assert-True ($planStart -ge 0) 'ExecuteAgentPlan not found in source'
    if ($planStart -ge 0) {
        $planBody = $executor.Substring($planStart, [Math]::Min(9000, $executor.Length - $planStart))
        Assert-True ($planBody.Contains('ExecuteAgentInsertGroupSql')) 'ExecuteAgentPlan should insert via SQL'
        Assert-True (-not ($planBody -match 'ExecuteAgentInsertGroup\(')) 'ExecuteAgentPlan must not use the clipboard paste path'
        Assert-True ($planBody.Contains('using (SqlTransaction transaction = conn.BeginTransaction())')) 'ExecuteAgentPlan should open one SQL transaction'
        Assert-True ($planBody.Contains('ExecuteAgentInsertGroupSql(conn, transaction')) 'insert groups must use the active transaction'
        Assert-True ($planBody.Contains('transaction.Commit()')) 'ExecuteAgentPlan should commit the transaction'
        Assert-True ($planBody.Contains('transaction.Rollback()')) 'ExecuteAgentPlan should roll back on failure'
        Assert-True ($planBody.Contains('undo.Rows.Clear()')) 'rolled-back groups must not keep undo rows'
        Assert-True ($planBody.Contains('if (affected != 1)')) 'changed or missing target rows must roll back the group'
        Assert-True ($planBody.Contains('Object.ReferenceEquals(conn, plan.ProjectConnection)')) 'project connection object gate must remain'
        Assert-True ($planBody.Contains('plan.ProjectConnectionIdentity')) 'project connection identity gate must remain'
    }

    $insertStart = $executor.IndexOf('private static int ExecuteAgentInsertGroupSql(')
    $insertEnd = $executor.IndexOf('private static void SetAgentRowValue(', $insertStart)
    Assert-True ($insertStart -ge 0 -and $insertEnd -gt $insertStart) 'ExecuteAgentInsertGroupSql source block not found'
    if ($insertStart -ge 0 -and $insertEnd -gt $insertStart) {
        $insertBody = $executor.Substring($insertStart, $insertEnd - $insertStart)
        Assert-True ($insertBody.Contains('LoadTemplateFullRow(conn, transaction')) 'same-code and anchor rows must be cloned inside the transaction'
        Assert-True ($insertBody.Contains('LoadAgentStructuralRow(conn, transaction')) 'fallback structural row must be read inside the transaction'
        # 不能用推荐定额那份 22 列清单卡：智能指令只写少数几列，费用列是"存在才清零"。
        # 列少的项目会被那道门禁整批拦下（现场实测：替换和新增全部失败）。
        Assert-True ($insertBody.Contains('FindMissingAgentInsertColumns')) 'insert gate must check only the columns the agent writes'
        Assert-True (-not $insertBody.Contains('HasSmartFillRequiredSourceColumns')) 'insert must not reuse the SmartFill column gate'
        Assert-True ($insertBody.Contains('SetAgentRowValue(row, "工程或费用项目名称", quota.Name)')) 'inserted name must be written'
        Assert-True ($insertBody.Contains('SetAgentRowValue(row, "单位", quota.Unit)')) 'inserted unit must be written'
        Assert-True ($insertBody.Contains('SetAgentRowValue(row, "工程数量输入", quantityInput)')) 'inserted quantity expression must be written'
        Assert-True ($insertBody.Contains('SetAgentRowValue(row, "工程数量", quantity)')) 'inserted calculated quantity must be written'
        Assert-True ($insertBody.Contains('InsertQuotaRowReturnId(conn, transaction, row)')) 'insert must use the active transaction'
    }

    $orderStart = $executor.IndexOf('private static int AllocateAgentInsertOrders(')
    $orderEnd = $executor.IndexOf('private static void BuildDeletePlan(', $orderStart)
    Assert-True ($orderStart -ge 0 -and $orderEnd -gt $orderStart) 'AllocateAgentInsertOrders source block not found'
    if ($orderStart -ge 0 -and $orderEnd -gt $orderStart) {
        $orderBody = $executor.Substring($orderStart, $orderEnd - $orderStart)
        Assert-True ($orderBody.Contains('AgentUndoRecord undo')) 'shifted order rows must receive the undo record'
        Assert-True ($orderBody.Contains('undo.Rows.Add(new AgentUndoRow')) 'shifted order rows must be undoable'
        Assert-True ($orderBody.Contains('ExecuteAgentFieldUpdate(conn, transaction')) 'order shifts must stay in the active transaction'
    }

    Assert-True ($executor.Contains('quota.Name = sourceQuota.Name;')) 'replace extra rows must retain the resolved name'
    Assert-True ($executor.Contains('quota.Unit = sourceQuota.Unit;')) 'replace extra rows must retain the resolved unit'
    Assert-True ($executor.Contains('quota.SameCodeQuotaSequence = sourceQuota.SameCodeQuotaSequence;')) 'replace extra rows must retain the same-code clone source'
    Assert-True ($executor.Contains('oldValues["单位"] = anchor.Unit;')) 'undo must restore the original unit'
    Assert-True (-not $executor.Contains('LoadSmartFillStructuralRow(conn, group.UnitId')) 'agent structural reads must not escape the active transaction'
    Assert-True (-not $executor.Contains('一条定额拆成多条时只能作用于一个单元')) 'cross-unit block should be removed'
    Assert-True (-not $executor.Contains('一条定额拆成多条时只能作用于当前正在显示的单元')) 'current-unit block should be removed'
}

if ($failures.Count -gt 0) {
    foreach ($failure in $failures) { Write-Host "FAIL: $failure" }
    throw "Test-AgentReplaceQuotas: $($failures.Count) assertion(s) failed."
}

Write-Host 'PASS: replace-quotas model, matching rule and panel wiring.'
