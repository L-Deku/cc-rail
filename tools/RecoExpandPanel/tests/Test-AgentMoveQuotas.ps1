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

function Parse-MoveCommand([string]$text) {
    $invokeArgs = New-Object object[] 2
    $invokeArgs[0] = $text
    $handled = [bool]$parse.Invoke($null, $invokeArgs)
    if (-not $handled) {
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

function Read-MoveParseError([string]$text) {
    $invokeArgs = New-Object object[] 2
    $invokeArgs[0] = $text
    if (-not [bool]$parse.Invoke($null, $invokeArgs)) {
        throw "Command was not handled: $text"
    }

    $result = $invokeArgs[1]
    return [string]$result.GetType().GetField('Error', $flags).GetValue($result)
}

$all = Parse-MoveCommand '移动定额 0305 到 0306'
if ((Read-CommandField $all 'Type') -ne 'move_quotas') { throw 'Move command type was not parsed.' }
if ((Read-CommandField $all 'SourceItem') -ne '0305') { throw 'Move source item was not parsed.' }
if ((Read-CommandField $all 'TargetItems').Count -ne 1 -or (Read-CommandField $all 'TargetItems')[0] -ne '0306') {
    throw 'Move target item was not parsed.'
}
if ((Read-CommandField $all 'QuotaFilter').Count -ne 0) { throw 'Unfiltered move unexpectedly has a quota filter.' }

$filtered = Parse-MoveCommand '移动定额 0305 LY-21、LY-22 到0306 单元=04'
$filters = Read-CommandField $filtered 'QuotaFilter'
if ($filters.Count -ne 2 -or $filters[0] -ne 'LY-21' -or $filters[1] -ne 'LY-22') {
    throw 'Move quota filter was not parsed.'
}
$units = Read-CommandField $filtered 'Units'
if ($units.Count -ne 1 -or $units[0] -ne '04') { throw 'Move unit filter was not parsed.' }

if ([String]::IsNullOrEmpty((Read-MoveParseError '移动定额 0305 到 0306、0307'))) {
    throw 'Move command unexpectedly accepted multiple targets.'
}
if ([String]::IsNullOrEmpty((Read-MoveParseError '移动定额 0305 到 0305'))) {
    throw 'Move command unexpectedly accepted the same source and target.'
}

Write-Host 'PASS: move-quotas deterministic command parsing.'
