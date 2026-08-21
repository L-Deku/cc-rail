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
        'LoadAgentQuotaCandidates',
        'ResolveAgentScopeRows',
        'AgentQuotaCodeKind',
        'ShowAgentPanelWindow')) {
    Assert-True ($null -ne $formType.GetMethod($methodName, $flags)) "method $methodName missing"
}

# 4) 点选面板与选择对话框类型存在
foreach ($typeName in @('AgentPanelWindow', 'AgentPickerDialog', 'AgentUnitOption', 'AgentItemOption', 'AgentQuotaCandidate')) {
    Assert-True ($null -ne $formType.GetNestedType($typeName, $flags)) "nested type $typeName missing"
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

if ($failures.Count -gt 0) {
    foreach ($failure in $failures) { Write-Host "FAIL: $failure" }
    throw "Test-AgentReplaceQuotas: $($failures.Count) assertion(s) failed."
}

Write-Host 'PASS: replace-quotas model, matching rule and panel wiring.'
