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
$parse = $formType.GetMethod('TryParseAgentFallback', $flags)

function Parse-AgentCommand([string]$text) {
    $invokeArgs = New-Object object[] 2
    $invokeArgs[0] = $text
    if (-not [bool]$parse.Invoke($null, $invokeArgs)) {
        throw "Command was not handled: $text"
    }

    $result = $invokeArgs[1]
    $resultType = $result.GetType()
    $error = [string]$resultType.GetField('Error', $flags).GetValue($result)
    if (-not [String]::IsNullOrEmpty($error)) {
        throw "Command parse failed: $error"
    }

    $commands = $resultType.GetField('Commands', $flags).GetValue($result)
    if ($commands.Count -ne 1) {
        throw "Expected one command, got $($commands.Count): $text"
    }

    return $commands[0]
}

function Read-CommandField($command, [string]$name) {
    $value = $command.GetType().GetField($name, $flags).GetValue($command)
    Write-Output -NoEnumerate $value
}

function Assert-OldCommandRejected([string]$text, [string]$replacement) {
    $invokeArgs = New-Object object[] 2
    $invokeArgs[0] = $text
    if (-not [bool]$parse.Invoke($null, $invokeArgs)) {
        throw "Old command was not intercepted: $text"
    }

    $result = $invokeArgs[1]
    $error = [string]$result.GetType().GetField('Error', $flags).GetValue($result)
    if ([String]::IsNullOrEmpty($error) -or $error.IndexOf($replacement, [StringComparison]::Ordinal) -lt 0) {
        throw "Old command did not point to '$replacement': $text / $error"
    }
}

function Assert-Field($command, [string]$name, [string]$expected) {
    $actual = [string](Read-CommandField $command $name)
    if ($actual -ne $expected) {
        throw "Expected $name='$expected', got '$actual'."
    }
}

function Assert-List($command, [string]$name, [string[]]$expected) {
    $actual = Read-CommandField $command $name
    if ($actual.Count -ne $expected.Count) {
        throw "Expected $name count $($expected.Count), got $($actual.Count)."
    }
    for ($i = 0; $i -lt $expected.Count; $i++) {
        if ([string]$actual[$i] -ne $expected[$i]) {
            throw "Expected $name[$i]='$($expected[$i])', got '$($actual[$i])'."
        }
    }
}

$priceCodeSet = Parse-AgentCommand '单价 SH 3500'
Assert-Field $priceCodeSet 'Type' 'set_unit_price'
Assert-Field $priceCodeSet 'Target' 'unit_price'
Assert-Field $priceCodeSet 'Value' '3500'
Assert-List $priceCodeSet 'Items' @('@currentitem')
Assert-List $priceCodeSet 'QuotaFilter' @('SH')

$priceNameSet = Parse-AgentCommand '单价 ≤25t自卸汽车运土 增运1km 615.98 单元=04'
Assert-Field $priceNameSet 'Type' 'set_unit_price'
Assert-Field $priceNameSet 'QuotaName' '≤25t自卸汽车运土 增运1km'
Assert-Field $priceNameSet 'Value' '615.98'
Assert-List $priceNameSet 'Items' @()
Assert-List $priceNameSet 'Units' @('04')

$priceCodeMultiply = Parse-AgentCommand '单价 0101-01 SH *10'
Assert-Field $priceCodeMultiply 'Type' 'multiply_quantity'
Assert-Field $priceCodeMultiply 'Target' 'unit_price'
Assert-Field $priceCodeMultiply 'Operator' '*'
Assert-Field $priceCodeMultiply 'Factor' '10'
Assert-List $priceCodeMultiply 'Items' @('0101-01')
Assert-List $priceCodeMultiply 'QuotaFilter' @('SH')

$priceNameDivide = Parse-AgentCommand '单价 ≤25t自卸汽车运土 增运1km /2 单元=所有'
Assert-Field $priceNameDivide 'Type' 'multiply_quantity'
Assert-Field $priceNameDivide 'Target' 'unit_price'
Assert-Field $priceNameDivide 'QuotaName' '≤25t自卸汽车运土 增运1km'
Assert-Field $priceNameDivide 'Operator' '/'
Assert-Field $priceNameDivide 'Factor' '2'

$quantityCodeSet = Parse-AgentCommand '工程数量 LY-21 100'
Assert-Field $quantityCodeSet 'Type' 'set_quantity'
Assert-Field $quantityCodeSet 'Value' '100'
Assert-List $quantityCodeSet 'Items' @('@currentitem')
Assert-List $quantityCodeSet 'QuotaFilter' @('LY-21')

$quantityNameSet = Parse-AgentCommand '工程数量 0101-01 ≤25t自卸汽车运土 增运1km 100'
Assert-Field $quantityNameSet 'Type' 'set_quantity'
Assert-Field $quantityNameSet 'QuotaName' '≤25t自卸汽车运土 增运1km'
Assert-List $quantityNameSet 'Items' @('0101-01')

$quantityNameMultiply = Parse-AgentCommand '工程数量 ≤25t自卸汽车运土 增运1km *10'
Assert-Field $quantityNameMultiply 'Type' 'multiply_quantity'
Assert-Field $quantityNameMultiply 'Target' 'quantity'
Assert-Field $quantityNameMultiply 'QuotaName' '≤25t自卸汽车运土 增运1km'
Assert-Field $quantityNameMultiply 'Factor' '10'

$quantityNameRemove = Parse-AgentCommand '工程数量 0821-01-04-09-02 FAS 联动接入、门禁系统一键释放联动接入 删除*0 单元=所有'
Assert-Field $quantityNameRemove 'Type' 'remove_text'
Assert-Field $quantityNameRemove 'Target' 'quantity'
Assert-Field $quantityNameRemove 'QuotaName' 'FAS 联动接入、门禁系统一键释放联动接入'
Assert-Field $quantityNameRemove 'RemoveText' '*0'
Assert-List $quantityNameRemove 'Items' @('0821-01-04-09-02')
Assert-List $quantityNameRemove 'QuotaFilter' @()
Assert-List $quantityNameRemove 'Units' @('所有')

$quantityCodeListMultiply = Parse-AgentCommand '工程数量 0821-01-04-09-02 DY-873、DY-1169 *0 单元=所有'
Assert-Field $quantityCodeListMultiply 'Type' 'multiply_quantity'
Assert-Field $quantityCodeListMultiply 'Target' 'quantity'
Assert-Field $quantityCodeListMultiply 'Factor' '0'
Assert-List $quantityCodeListMultiply 'Items' @('0821-01-04-09-02')
Assert-List $quantityCodeListMultiply 'QuotaFilter' @('DY-873', 'DY-1169')
Assert-List $quantityCodeListMultiply 'Units' @('所有')

$quantityItemSet = Parse-AgentCommand '工程数量 0305 0'
Assert-Field $quantityItemSet 'Type' 'set_quantity'
Assert-Field $quantityItemSet 'Value' '0'
Assert-List $quantityItemSet 'Items' @('0305')
Assert-List $quantityItemSet 'QuotaFilter' @()

$quantityMaterialSet = Parse-AgentCommand '工程数量 1294861 5'
Assert-Field $quantityMaterialSet 'Type' 'set_quantity'
Assert-List $quantityMaterialSet 'Items' @('@currentitem')
Assert-List $quantityMaterialSet 'QuotaFilter' @('1294861')

Assert-OldCommandRejected '改单价 ≤25t自卸汽车运土 增运1km 615.98 单元=04' '单价'
Assert-OldCommandRejected '设单价 SH 3500' '单价'
Assert-OldCommandRejected '设数量 LY-21 100' '工程数量'

Write-Host 'PASS: unified price and quantity command parsing.'
