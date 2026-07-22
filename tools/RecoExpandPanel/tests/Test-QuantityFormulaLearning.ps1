$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..'))
$dll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL)) { $env:RECO_EXPAND_DLL } else { Join-Path $repoRoot 'RecoQuotaRecommend\bin\RecoExpandPanel.dll' }
if (-not (Test-Path -LiteralPath $dll)) { throw "Missing DLL: $dll" }
$dllDir = Split-Path -Parent $dll
foreach ($dependency in @('NPOI.dll', 'NPOI.OpenXmlFormats.dll', 'NPOI.OpenXml4Net.dll', 'NPOI.OOXML.dll', 'ICSharpCode.SharpZipLib.dll')) {
    $dependencyPath = Join-Path $dllDir $dependency
    if (Test-Path -LiteralPath $dependencyPath) { [void][System.Reflection.Assembly]::LoadFrom($dependencyPath) }
}

$type = [System.Reflection.Assembly]::LoadFrom($dll).GetType('RecoNet.FormPanel', $true)
$flags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$nestedFlags = [System.Reflection.BindingFlags]'Public,NonPublic'

$unitScale = $type.GetMethod('TryBuildExcelLinkUnitScaleSuffix', $flags)
function Assert-Unit([string]$From, [string]$To, [bool]$ExpectedOk, [string]$ExpectedSuffix) {
    $args = [object[]]::new(3); $args[0] = $From; $args[1] = $To; $args[2] = $null
    $ok = [bool]$unitScale.Invoke($null, $args)
    if ($ok -ne $ExpectedOk -or ($ok -and [string]$args[2] -ne $ExpectedSuffix)) {
        throw "Unit conversion mismatch: $From -> $To, ok=$ok suffix='$($args[2])'"
    }
}
Assert-Unit 'm3' '10m3' $true '/10'
Assert-Unit 'kg' 't' $true '/1000'
Assert-Unit 'm2' 'm3' $false ''
Assert-Unit '天然密实方' '压实方' $false ''
Assert-Unit '天然密实方' '天然密实方' $true ''
Write-Host 'PASS 标准同量纲实时换算，跨量纲和不同方态不静默换算'

$fixture = Join-Path ([IO.Path]::GetTempPath()) ('reco-formula-learning-' + [Guid]::NewGuid().ToString('N') + '.xlsx')
$book = $null
try {
    $book = New-Object NPOI.XSSF.UserModel.XSSFWorkbook
    $sheet = $book.CreateSheet('Sheet1')
    foreach ($definition in @(
        [pscustomobject]@{ Row = 0; Name = '桩根数'; Unit = '根'; Value = 10 },
        [pscustomobject]@{ Row = 1; Name = '每根长度'; Unit = 'm'; Value = 5 },
        [pscustomobject]@{ Row = 2; Name = '桩半径'; Unit = 'm'; Value = 0.4 },
        [pscustomobject]@{ Row = 3; Name = '桩截面积'; Unit = 'm2'; Value = 100 }
    )) {
        $row = $sheet.CreateRow($definition.Row)
        $row.CreateCell(0).SetCellValue($definition.Name)
        $row.CreateCell(1).SetCellValue($definition.Unit)
        $row.CreateCell(5).SetCellValue([double]$definition.Value)
    }
    $stream = [IO.File]::Create($fixture)
    try { $book.Write($stream) } finally { $stream.Dispose() }

    $linkType = $type.GetNestedType('ExcelQuotaLink', $nestedFlags)
    $dictType = [Collections.Generic.Dictionary``2].MakeGenericType($linkType, [string])
    $links = [Activator]::CreateInstance($dictType)
    foreach ($definition in @(
        [pscustomobject]@{ Cell = 'F1'; Expression = 'F1*F2*F3*F3*3.14'; Code = 'TEST-PILE'; Name = '桩身混凝土'; Unit = 'm3' },
        [pscustomobject]@{ Cell = 'F4'; Expression = 'F4*0.2'; Code = 'TEST-SLAB'; Name = '板混凝土'; Unit = 'm3' }
    )) {
        $link = [Activator]::CreateInstance($linkType)
        foreach ($pair in @{
            ExcelPath=$fixture; WorksheetName='Sheet1'; CellAddress=$definition.Cell; Expression=$definition.Expression;
            QuotaCode=$definition.Code; QuotaName=$definition.Name; QuotaUnit=$definition.Unit;
            EntryCode='0101-01'; EntryName='测试条目'; Method='2024'
        }.GetEnumerator()) { $linkType.GetProperty($pair.Key, $flags).SetValue($link, $pair.Value, $null) }
        $links.Add($link, '')
    }

    $buildGroups = $type.GetMethod('BuildBindingFeedbackGroups', $flags)
    $groups = $buildGroups.Invoke($null, @($links.PSObject.BaseObject))
    $groupType = $type.GetNestedType('MappingFeedbackGroup', $nestedFlags)
    $targetType = $type.GetNestedType('MappingFeedbackTarget', $nestedFlags)
    if ($groups.Count -ne 2) {
        $learnedNames = @($groups | ForEach-Object { [string]$groupType.GetField('QuantityName', $flags).GetValue($_) }) -join ', '
        throw "Expected 2 learned groups, got $($groups.Count): $learnedNames"
    }
    $composite = @($groups | Where-Object { $groupType.GetField('QuantityName', $flags).GetValue($_) -eq '桩根数' })[0]
    $scalar = @($groups | Where-Object { $groupType.GetField('QuantityName', $flags).GetValue($_) -eq '桩截面积' })[0]
    $compositeOperands = $groupType.GetField('FormulaOperands', $flags).GetValue($composite)
    $compositeTargets = $groupType.GetField('Targets', $flags).GetValue($composite)
    $scalarOperands = $groupType.GetField('FormulaOperands', $flags).GetValue($scalar)
    $scalarTargets = $groupType.GetField('Targets', $flags).GetValue($scalar)
    if ($compositeOperands.Count -ne 3 -or [string]$targetType.GetField('FormulaTemplate', $flags).GetValue($compositeTargets[0]) -ne 'V0*V1*V2*V2*3.14') {
        throw 'Composite formula template/operands were not learned correctly'
    }
    if ($scalarOperands.Count -ne 1 -or [string]$targetType.GetField('FormulaTemplate', $flags).GetValue($scalarTargets[0]) -ne 'V0*0.2') {
        throw 'Single-operand cross-unit factor was not stored as a formula'
    }
    Write-Host 'PASS 单系数和根数×长度×半径²×3.14统一保存为变量公式'

    $readRows = $type.GetMethod('ReadTargetQtyRowsWithChapters', $flags)
    $readArgs = [object[]]::new(4); $readArgs[0] = $fixture.PSObject.BaseObject; $readArgs[1] = ([string]'Sheet1').PSObject.BaseObject; $readArgs[2] = 6; $readArgs[3] = $null
    $targetRows = $readRows.Invoke($null, $readArgs.PSObject.BaseObject)
    $formulaRuleType = $type.GetNestedType('SmartFormulaRule', $nestedFlags)
    $formulaOperandType = $type.GetNestedType('SmartFormulaOperand', $nestedFlags)
    $rule = [Activator]::CreateInstance($formulaRuleType, $true).PSObject.BaseObject
    $formulaRuleType.GetField('TargetUnit', $flags).SetValue($rule, 'm3')
    $formulaRuleType.GetField('Template', $flags).SetValue($rule, 'V0*V1*V2*V2*3.14')
    $ruleOperands = $formulaRuleType.GetField('Operands', $flags).GetValue($rule)
    for ($i = 0; $i -lt $compositeOperands.Count; $i++) {
        $sourceOperand = $compositeOperands[$i]
        $operand = [Activator]::CreateInstance($formulaOperandType, $true).PSObject.BaseObject
        $formulaOperandType.GetField('Index', $flags).SetValue($operand, $i)
        foreach ($field in @('Signature','Name','Unit')) {
            $formulaOperandType.GetField($field, $flags).SetValue($operand, $sourceOperand.GetType().GetField($field, $flags).GetValue($sourceOperand))
        }
        [void]$ruleOperands.Add($operand)
    }
    $evaluateFormula = $type.GetMethod('TryEvaluateSmartFormula', $flags)
    $evalArgs = [object[]]::new(5); $evalArgs[0] = $rule; $evalArgs[1] = $targetRows; $evalArgs[2] = $targetRows[0]; $evalArgs[3] = 'm3'; $evalArgs[4] = $null
    if (-not [bool]$evaluateFormula.Invoke($null, $evalArgs)) { throw 'Composite formula could not be evaluated from current sheet operands' }
    $tryDecimal = $type.GetMethod('TryEvaluateDecimal', $flags, $null, [Type[]]@([string], [decimal].MakeByRefType(), [string].MakeByRefType()), $null)
    $decimalArgs = [object[]]::new(3); $decimalArgs[0] = [string]$evalArgs[4]; $decimalArgs[1] = [decimal]0; $decimalArgs[2] = $null
    if (-not [bool]$tryDecimal.Invoke($null, $decimalArgs) -or [Math]::Abs([decimal]$decimalArgs[1] - [decimal]25.12) -gt [decimal]0.000001) {
        throw "Formula result mismatch: '$($evalArgs[4])' => $($decimalArgs[1])"
    }
    Write-Host "PASS 当前表参数公式计算正确：$($evalArgs[4]) = $($decimalArgs[1])"

    # 跨单位公式只能命中当前办法+当前条目，或显式空条目的通用规则；不得回退到其他条目。
    $snapshotType = $type.GetNestedType('SmartLearningSnapshot', $nestedFlags)
    $smartTargetType = $type.GetNestedType('SmartBoxTarget', $nestedFlags)
    $snapshot = [Activator]::CreateInstance($snapshotType, $true).PSObject.BaseObject
    $snapshotType.GetField('Method', $flags).SetValue($snapshot, '2024')
    $target = [Activator]::CreateInstance($smartTargetType, $true).PSObject.BaseObject
    $smartTargetType.GetField('Kind', $flags).SetValue($target, 'quota')
    $smartTargetType.GetField('Code', $flags).SetValue($target, 'TEST-PILE')
    $signatureMethod = $type.GetMethod('BuildSmartQuantitySignature', $flags)
    $anchorSignature = [string]$signatureMethod.Invoke($null, [object[]]@('桩根数', '根'))
    $formulaKeyMethod = $type.GetMethod('BuildSmartFormulaKey', $flags)
    $formulaKey = [string]$formulaKeyMethod.Invoke($null, [object[]]@($anchorSignature, 'quota', 'TEST-PILE'))
    $formulaByKey = $snapshotType.GetField('FormulaByKey', $flags).GetValue($snapshot)
    $ruleListType = [Collections.Generic.List``1].MakeGenericType($formulaRuleType)
    $ruleList = [Activator]::CreateInstance($ruleListType)
    $formulaRuleType.GetField('RuleHash', $flags).SetValue($rule, 'rule-context-test')
    $formulaRuleType.GetField('Method', $flags).SetValue($rule, '2024')
    $formulaRuleType.GetField('EntryCode', $flags).SetValue($rule, 'OTHER-ENTRY')
    [void]$ruleList.Add($rule)
    $formulaByKey.Add($formulaKey, $ruleList)
    $resolveFormula = $type.GetMethod('TryResolveSmartFormula', $flags)
    function Invoke-ResolveFormula([string]$EntryCode) {
        $resolveArgs = [object[]]::new(10)
        $resolveArgs[0] = $snapshot
        $resolveArgs[1] = $targetRows
        $resolveArgs[2] = $targetRows[0]
        $resolveArgs[3] = $target
        $resolveArgs[4] = 'm3'
        $resolveArgs[5] = $EntryCode
        $resolveArgs[6] = $anchorSignature
        $resolveArgs[7] = $null
        $resolveArgs[8] = $null
        $resolveArgs[9] = $null
        return [pscustomobject]@{ Ok = [bool]$resolveFormula.Invoke($null, $resolveArgs); Args = $resolveArgs }
    }
    $wrongEntryResult = Invoke-ResolveFormula '0101-01'
    if ($wrongEntryResult.Ok -or [string]$wrongEntryResult.Args[9] -notmatch '当前办法/条目') {
        throw "错误条目公式被复用：$($wrongEntryResult.Args[9])"
    }
    $formulaRuleType.GetField('EntryCode', $flags).SetValue($rule, '')
    $genericResult = Invoke-ResolveFormula '0101-01'
    if (-not $genericResult.Ok) { throw "显式空条目通用公式未命中：$($genericResult.Args[9])" }
    $formulaRuleType.GetField('EntryCode', $flags).SetValue($rule, '0101-01')
    $exactResult = Invoke-ResolveFormula '0101-01'
    if (-not $exactResult.Ok) { throw "当前办法+条目精确公式未命中：$($exactResult.Args[9])" }
    Write-Host 'PASS 公式严格按当前办法/条目或显式通用规则选择，不回退其他条目'

    $candidateScoreType = $type.GetNestedType('SmartMapCandidateScore', $nestedFlags)
    $mapEntryType = $type.GetNestedType('SmartMapEntry', $nestedFlags)
    $scoreListType = [Collections.Generic.List``1].MakeGenericType($candidateScoreType)
    $scoreList = [Activator]::CreateInstance($scoreListType)
    foreach ($weight in @(50, 40)) {
        $score = [Activator]::CreateInstance($candidateScoreType, $true).PSObject.BaseObject
        $mapEntry = [Activator]::CreateInstance($mapEntryType, $true).PSObject.BaseObject
        $mapEntryType.GetField('Weight', $flags).SetValue($mapEntry, $weight)
        $candidateScoreType.GetField('Entry', $flags).SetValue($score, $mapEntry)
        foreach ($field in @('CurrentTargetsValid','HasCurrentMethodMapping','HasEntry','HasCurrentContext')) {
            $candidateScoreType.GetField($field, $flags).SetValue($score, $true)
        }
        [void]$scoreList.Add($score)
    }
    $canAuto = $type.GetMethod('CanAutoSelectSmartMapEntry', $flags)
    function Invoke-CanAutoSelect {
        $invokeArgs = New-Object 'object[]' 1
        $invokeArgs[0] = $scoreList.PSObject.BaseObject
        return [bool]$canAuto.Invoke($null, $invokeArgs)
    }
    if (Invoke-CanAutoSelect) { throw '两个有效组件权重差10时不应静默选择' }
    $secondEntry = $candidateScoreType.GetField('Entry', $flags).GetValue($scoreList[1])
    $mapEntryType.GetField('Weight', $flags).SetValue($secondEntry, 20)
    if (-not (Invoke-CanAutoSelect)) { throw '当前办法/条目唯一且权重差30时应允许自动选择' }
    $candidateScoreType.GetField('HasCurrentMethodMapping', $flags).SetValue($scoreList[0], $false)
    if (Invoke-CanAutoSelect) { throw '空办法兼容关系不应静默压过当前办法关系' }
    Write-Host 'PASS 组件候选小权重差需确认，且空办法关系不能自动采纳'

    $targetRowType = $targetRows[0].GetType()
    $duplicateRadius = [Activator]::CreateInstance($targetRowType, $true).PSObject.BaseObject
    foreach ($pair in @{ Row=5; RawName='桩半径'; DisplayName='桩半径'; NormName='桩半径'; Chapter=$targetRowType.GetField('Chapter', $flags).GetValue($targetRows[0]); Unit='m'; Quantity=[decimal]0.5; QuantityText='0.5' }.GetEnumerator()) {
        $targetRowType.GetField($pair.Key, $flags).SetValue($duplicateRadius, $pair.Value)
    }
    [void]$targetRows.Add($duplicateRadius)
    $ambiguousArgs = [object[]]::new(5); $ambiguousArgs[0] = $rule; $ambiguousArgs[1] = $targetRows; $ambiguousArgs[2] = $targetRows[0]; $ambiguousArgs[3] = 'm3'; $ambiguousArgs[4] = $null
    if ([bool]$evaluateFormula.Invoke($null, $ambiguousArgs)) { throw '同章节邻近行出现重复桩半径时不应自动选择或相加。' }
    Write-Host 'PASS 同名参数重复时拒绝自动计算，不跨行累加'
}
finally {
    if ($book -ne $null) { $book.Close() }
    if (Test-Path -LiteralPath $fixture) { Remove-Item -LiteralPath $fixture -Force }
}

Write-Host 'Test-QuantityFormulaLearning: PASS'
