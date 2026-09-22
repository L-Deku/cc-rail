$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# Regression for remove_text (right-click "删系数" and Ctrl+Q "删除*x" / "删除数量" / "去掉编号" all go through it).
# Quantity: multiply always writes "(expr)*f", so removal may only undo such a layer -
#   the trailing fragment, or a fragment sitting between ")" and ")" (a layer later wrapped by another multiply).
#   Fragments the user typed inside the expression ("*6" in "(5*6*8*2.5*1/10)*8") must never be touched.
# Bugs covered: "(52.232+515.52+309.6*0.05/10)*0" minus "*0" became "309.6.05" (Replace-all),
#   "(3*12)*1" minus "*1" silently became "(32)", "(5*6*8*2.5*1/10)*8" minus "*6" removed the inner *6.
# Also: factor text normalization so "×6" / "*1.0" typed in Ctrl+Q match the "*6" / "*1" multiply wrote.

$dll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL)) {
    $env:RECO_EXPAND_DLL
} else {
    Join-Path (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..')).Path 'RecoQuotaRecommend\bin\RecoExpandPanel.dll'
}

Add-Type -Path $dll
$formType = [RecoNet.FormPanel]
$flags = [System.Reflection.BindingFlags]'NonPublic,Static'
$anyFlags = [System.Reflection.BindingFlags]'NonPublic,Static,Instance,Public'

function Get-Method([string]$name) {
    $method = $formType.GetMethod($name, $flags)
    if ($null -eq $method) {
        throw "Method not found: $name"
    }
    return $method
}

$removeQuantity = Get-Method 'TryRemoveAgentQuantityFragment'
$removeFragment = Get-Method 'TryRemoveAgentFragment'
$normalizeFragment = Get-Method 'NormalizeAgentOperatorFragment'
$parseFactor = Get-Method 'TryParseFactor'
$parseAgent = Get-Method 'TryParseAgentFallback'

function Invoke-Remove($method, [string]$text, [string]$fragment) {
    $invokeArgs = New-Object object[] 3
    $invokeArgs[0] = $text
    $invokeArgs[1] = $fragment
    $ok = [bool]$method.Invoke($null, $invokeArgs)
    return @{ Ok = $ok; Result = [string]$invokeArgs[2] }
}

$failures = New-Object System.Collections.Generic.List[string]

function Assert-Removed($method, [string]$label, [string]$text, [string]$fragment, [string]$expected) {
    $r = Invoke-Remove $method $text $fragment
    if (-not $r.Ok -or $r.Result -ne $expected) {
        $failures.Add("[$label] '$text' - '$fragment': expected ok='$expected', got ok=$($r.Ok) '$($r.Result)'")
    }
}

function Assert-Skipped($method, [string]$label, [string]$text, [string]$fragment) {
    $r = Invoke-Remove $method $text $fragment
    if ($r.Ok) {
        $failures.Add("[$label] '$text' - '$fragment': expected skip, got '$($r.Result)'")
    }
}

function Assert-Equal([string]$label, [string]$actual, [string]$expected) {
    if ($actual -ne $expected) {
        $failures.Add("[$label] expected '$expected', got '$actual'")
    }
}

# --- quantity expression: undo exactly one multiply layer ---
# the two rows from the field report: inner *0.05 / *0.16 must survive
Assert-Removed $removeQuantity 'qty' '(52.232+515.52+309.6*0.05/10)*0' '*0' '52.232+515.52+309.6*0.05/10'
Assert-Removed $removeQuantity 'qty' '(154.66+1082.59*0.16/10)*0'      '*0' '154.66+1082.59*0.16/10'
# silent corruption cases with the old Replace-all
Assert-Removed $removeQuantity 'qty' '(3*12)*1'   '*1' '3*12'
Assert-Removed $removeQuantity 'qty' '(2.5*2)*2'  '*2' '2.5*2'
Assert-Removed $removeQuantity 'qty' '(1.325)*0'  '*0' '1.325'
# user-typed operators inside the expression are not multiply layers -> skip
Assert-Skipped $removeQuantity 'qty' '(5*6*8*2.5*1/10)*8' '*6'
Assert-Skipped $removeQuantity 'qty' '(5*6*8*2.5*1/10)*8' '/10'
Assert-Removed $removeQuantity 'qty' '(5*6*8*2.5*1/10)*8' '*8' '5*6*8*2.5*1/10'
Assert-Skipped $removeQuantity 'qty' '(2+3)*6*7' '*6'
Assert-Removed $removeQuantity 'qty' '(2+3)*6*7' '*7' '(2+3)*6'
Assert-Skipped $removeQuantity 'qty' '(36.69/100)*10' '/100'
Assert-Removed $removeQuantity 'qty' '(36.69/100)*10' '*10' '36.69/100'
# nested multiply: removing an inner layer also drops the parentheses that layer added
Assert-Removed $removeQuantity 'qty' '((1.325)*0)*2'   '*0' '(1.325)*2'
Assert-Removed $removeQuantity 'qty' '((1.325)*0)*0'   '*0' '(1.325)*0'
Assert-Removed $removeQuantity 'qty' '((5*6*8)*6)*8'   '*6' '(5*6*8)*8'
Assert-Removed $removeQuantity 'qty' '(((2+3))*0)*2'   '*0' '((2+3))*2'
Assert-Removed $removeQuantity 'qty' '((2+3))*8'       '*8' '(2+3)'
Assert-Removed $removeQuantity 'qty' '(((1.325)*0)*2)*3' '*0' '((1.325)*2)*3'
# trailing hand-written conversion (Ctrl+Q "删除数量 /100") still works
Assert-Removed $removeQuantity 'qty' '36.69/100' '/100' '36.69'
Assert-Removed $removeQuantity 'qty' '5*6*8*2.5*1/10*6' '*6' '5*6*8*2.5*1/10'
# fragment only inside a number, or absent -> skip
Assert-Skipped $removeQuantity 'qty' '1.325*0.05' '*0'
Assert-Skipped $removeQuantity 'qty' '(1.325)*0.5' '*0'
Assert-Skipped $removeQuantity 'qty' '36.69/1000' '/100'
Assert-Skipped $removeQuantity 'qty' '36.69/2100' '100'
Assert-Skipped $removeQuantity 'qty' '(5+1100)*2' '100'
Assert-Skipped $removeQuantity 'qty' '1.325' '*0'
Assert-Skipped $removeQuantity 'qty' '' '*0'

# --- quota code (右键 定额编号/删系数, Ctrl+Q 去掉编号): appended without parentheses, each *f token is a coefficient ---
Assert-Removed $removeFragment 'code' 'SY-343*10'  '*10' 'SY-343'
Assert-Removed $removeFragment 'code' 'SY-343*0'   '*0'  'SY-343'
Assert-Removed $removeFragment 'code' 'SY-343*1*0' '*1'  'SY-343*0'
Assert-Removed $removeFragment 'code' 'SY-343/2*2' '/2'  'SY-343*2'
Assert-Skipped $removeFragment 'code' 'SY-343*10'  '*1'
Assert-Skipped $removeFragment 'code' 'SY-343*0.5' '*0'
Assert-Skipped $removeFragment 'code' 'SY-343'     '*0'

# --- factor text normalization: multiply suffix and removal fragment must agree ---
function Get-FactorSuffix([string]$text) {
    $invokeArgs = New-Object object[] 3
    $invokeArgs[0] = $text
    if (-not [bool]$parseFactor.Invoke($null, $invokeArgs)) { throw "TryParseFactor failed: $text / $($invokeArgs[2])" }
    $factor = $invokeArgs[1]
    return [string]$factor.GetType().GetProperty('Suffix', $anyFlags).GetValue($factor, $null)
}

Assert-Equal 'menu factor 1.0'   (Get-FactorSuffix '1.0')   '*1'
Assert-Equal 'menu factor *0.50' (Get-FactorSuffix '*0.50') '*0.5'
Assert-Equal 'menu factor ×6'    (Get-FactorSuffix '×6')    '*6'
Assert-Equal 'menu factor /2.000' (Get-FactorSuffix '/2.000') '/2'
Assert-Equal 'menu factor 100'   (Get-FactorSuffix '100')   '*100'
Assert-Equal 'menu factor 0'     (Get-FactorSuffix '0')     '*0'

function Normalize-Fragment([string]$text) { return [string]$normalizeFragment.Invoke($null, @($text)) }
Assert-Equal 'norm ×6'    (Normalize-Fragment '×6')    '*6'
Assert-Equal 'norm ÷100'  (Normalize-Fragment '÷100')  '/100'
Assert-Equal 'norm *1.0'  (Normalize-Fragment '*1.0')  '*1'
Assert-Equal 'norm /XG1'  (Normalize-Fragment '/XG1')  '/XG1'
Assert-Equal 'norm plain' (Normalize-Fragment 'abc')   'abc'

function Parse-AgentCommand([string]$text) {
    $invokeArgs = New-Object object[] 2
    $invokeArgs[0] = $text
    if (-not [bool]$parseAgent.Invoke($null, $invokeArgs)) { throw "Command was not handled: $text" }
    $result = $invokeArgs[1]
    $err = [string]$result.GetType().GetField('Error', $anyFlags).GetValue($result)
    if (-not [String]::IsNullOrEmpty($err)) { throw "Command parse failed: $err" }
    $commands = $result.GetType().GetField('Commands', $anyFlags).GetValue($result)
    if ($commands.Count -ne 1) { throw "Expected one command, got $($commands.Count): $text" }
    return $commands[0]
}

function Read-CommandField($command, [string]$name) {
    return [string]$command.GetType().GetField($name, $anyFlags).GetValue($command)
}

$c = Parse-AgentCommand '工程数量 删除×6'
Assert-Equal 'agent 工程数量 删除×6 type' (Read-CommandField $c 'Type') 'remove_text'
Assert-Equal 'agent 工程数量 删除×6 text' (Read-CommandField $c 'RemoveText') '*6'
$c = Parse-AgentCommand '工程数量 *1.0'
Assert-Equal 'agent 工程数量 *1.0 factor' (Read-CommandField $c 'Factor') '1'
$c = Parse-AgentCommand '定额编号 删除*1.0'
Assert-Equal 'agent 定额编号 删除*1.0' (Read-CommandField $c 'RemoveText') '*1'
$c = Parse-AgentCommand '删除数量 ÷100'
Assert-Equal 'agent 删除数量 ÷100 target' (Read-CommandField $c 'Target') 'quantity'
Assert-Equal 'agent 删除数量 ÷100 text' (Read-CommandField $c 'RemoveText') '/100'
$c = Parse-AgentCommand '去掉调整 /XG1'
Assert-Equal 'agent 去掉调整 /XG1 target' (Read-CommandField $c 'Target') 'adjustment'
Assert-Equal 'agent 去掉调整 /XG1 text' (Read-CommandField $c 'RemoveText') '/XG1'

if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Host "FAIL: $_" }
    throw "$($failures.Count) remove-fragment case(s) failed."
}

Write-Host 'PASS: remove_text only undoes multiply layers, never cuts numbers or user-typed operators; factor text normalized.'
