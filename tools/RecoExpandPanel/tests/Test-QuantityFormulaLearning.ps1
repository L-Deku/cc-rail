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
    if ($compositeOperands.Count -ne 3 -or
        [string]$targetType.GetField('FormulaTemplate', $flags).GetValue($compositeTargets[0]) -ne 'V0*V1*V2*V2*3.14' -or
        [string]$targetType.GetField('EntryCode', $flags).GetValue($compositeTargets[0]) -ne '0101-01') {
        throw 'Composite formula template/operands were not learned correctly'
    }
    if ($scalarOperands.Count -ne 1 -or [string]$targetType.GetField('FormulaTemplate', $flags).GetValue($scalarTargets[0]) -ne 'V0*0.2') {
        throw 'Single-operand cross-unit factor was not stored as a formula'
    }
    $formulaHash = $type.GetMethod('BuildLearningFormulaRuleHash', $flags)
    $targetType.GetField('EntryCode', $flags).SetValue($compositeTargets[0], '0101-01')
    $hashArgs = New-Object 'object[]' 2
    $hashArgs[0] = $composite.PSObject.BaseObject
    $hashArgs[1] = $compositeTargets[0].PSObject.BaseObject
    $hashA = [string]$formulaHash.Invoke($null, $hashArgs)
    $targetType.GetField('EntryCode', $flags).SetValue($compositeTargets[0], '0101-02')
    $hashB = [string]$formulaHash.Invoke($null, $hashArgs)
    $targetType.GetField('EntryCode', $flags).SetValue($compositeTargets[0], '0101-01')
    if ($hashA -eq $hashB) { throw 'Formula rule hash still uses the legacy group entry instead of the target entry' }
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
    $formatOperand = $type.GetMethod('FormatSmartFormulaOperandForInsertion', $flags)
    if ($null -eq $formatOperand) { throw '缺少公式数量的最小括号格式化入口' }
    foreach ($formatCase in @(
        @('V0*0.1', 0, 2, '13', '13'),
        @('V0+V1', 0, 2, '10-2', '10-2'),
        @('V0*2', 0, 2, '10/2', '10/2'),
        @('100/V0', 4, 2, '10/2', '(10/2)'),
        @('2-V0', 2, 2, '10-2', '(10-2)'),
        @('V0*2', 0, 2, '10+2', '(10+2)'),
        @('V0/100', 0, 2, '10+2', '(10+2)')
    )) {
        $formatted = [string]$formatOperand.Invoke($null, @(
            [string]$formatCase[0], [int]$formatCase[1], [int]$formatCase[2], [string]$formatCase[3]))
        if ($formatted -ne [string]$formatCase[4]) {
            throw "公式括号简化错误：$($formatCase[0]) / $($formatCase[3]) => $formatted"
        }
    }
    Write-Host 'PASS 纯加减和纯乘除去掉冗余括号，混合优先级保留必要括号'
    $evalArgs = [object[]]::new(5); $evalArgs[0] = $rule; $evalArgs[1] = $targetRows; $evalArgs[2] = $targetRows[0]; $evalArgs[3] = 'm3'; $evalArgs[4] = $null
    if (-not [bool]$evaluateFormula.Invoke($null, $evalArgs)) { throw 'Composite formula could not be evaluated from current sheet operands' }
    $tryDecimal = $type.GetMethod('TryEvaluateDecimal', $flags, $null, [Type[]]@([string], [decimal].MakeByRefType(), [string].MakeByRefType()), $null)
    $decimalArgs = [object[]]::new(3); $decimalArgs[0] = [string]$evalArgs[4]; $decimalArgs[1] = [decimal]0; $decimalArgs[2] = $null
    if (-not [bool]$tryDecimal.Invoke($null, $decimalArgs) -or [Math]::Abs([decimal]$decimalArgs[1] - [decimal]25.12) -gt [decimal]0.000001) {
        throw "Formula result mismatch: '$($evalArgs[4])' => $($decimalArgs[1])"
    }
    Write-Host "PASS 当前表参数公式计算正确：$($evalArgs[4]) = $($decimalArgs[1])"

    $negativeRule = [Activator]::CreateInstance($formulaRuleType, $true).PSObject.BaseObject
    $formulaRuleType.GetField('TargetUnit', $flags).SetValue($negativeRule, 'm3')
    $negativeOperands = $formulaRuleType.GetField('Operands', $flags).GetValue($negativeRule)
    $negativeOperand = [Activator]::CreateInstance($formulaOperandType, $true).PSObject.BaseObject
    $formulaOperandType.GetField('Index', $flags).SetValue($negativeOperand, 0)
    foreach ($field in @('Signature','Name','Unit')) {
        $formulaOperandType.GetField($field, $flags).SetValue($negativeOperand, $ruleOperands[0].GetType().GetField($field, $flags).GetValue($ruleOperands[0]))
    }
    [void]$negativeOperands.Add($negativeOperand)
    foreach ($invalidTemplate in @('V0*0', 'V0*-1')) {
        $formulaRuleType.GetField('Template', $flags).SetValue($negativeRule, $invalidTemplate)
        $negativeArgs = [object[]]::new(5); $negativeArgs[0] = $negativeRule; $negativeArgs[1] = $targetRows; $negativeArgs[2] = $targetRows[0]; $negativeArgs[3] = 'm3'; $negativeArgs[4] = $null
        if ([bool]$evaluateFormula.Invoke($null, $negativeArgs)) { throw "公式 $invalidTemplate 的计算结果小于等于0时不应进入自动预览。" }
    }
    Write-Host 'PASS 公式计算结果必须大于0'

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
    if ($genericResult.Ok -or [string]$genericResult.Args[9] -notmatch '当前办法/条目') {
        throw '空条目公式不得跨条目作为可信换算公式复用'
    }
    $formulaRuleType.GetField('EntryCode', $flags).SetValue($rule, '0101-01')
    $exactResult = Invoke-ResolveFormula '0101-01'
    if (-not $exactResult.Ok) { throw "当前办法+条目精确公式未命中：$($exactResult.Args[9])" }
    Write-Host 'PASS 公式严格按当前办法和当前条目选择，不回退其他条目或空条目'

    # 已有多参数派生公式时，即使锚点单位与定额单位相同，也必须先完整求值；缺参数不得回退锚点单值。
    $targetRowType = $targetRows[0].GetType()
    $targetRowListType = $targetRows.GetType()
    $sameUnitRows = [Activator]::CreateInstance($targetRowListType)
    foreach ($definition in @(
        [pscustomobject]@{ Row=10; Name='主体混凝土'; Value=[decimal]10 },
        [pscustomobject]@{ Row=11; Name='附加混凝土'; Value=[decimal]2 }
    )) {
        $targetRow = [Activator]::CreateInstance($targetRowType, $true).PSObject.BaseObject
        foreach ($pair in @{ Row=$definition.Row; RawName=$definition.Name; DisplayName=$definition.Name; NormName=$definition.Name; Chapter='测试章节'; Unit='m3'; Quantity=$definition.Value; QuantityText=$definition.Value.ToString([Globalization.CultureInfo]::InvariantCulture) }.GetEnumerator()) {
            $targetRowType.GetField($pair.Key, $flags).SetValue($targetRow, $pair.Value)
        }
        [void]$sameUnitRows.Add($targetRow)
    }

    $sameUnitSnapshot = [Activator]::CreateInstance($snapshotType, $true).PSObject.BaseObject
    $snapshotType.GetField('Method', $flags).SetValue($sameUnitSnapshot, '2024')
    $snapshotType.GetField('SoftwarePartition', $flags).SetValue($sameUnitSnapshot, '2024')
    $snapshotType.GetField('MethodNo', $flags).SetValue($sameUnitSnapshot, 'TB 10801—2024')
    $sameUnitSignature = [string]$signatureMethod.Invoke($null, [object[]]@('主体混凝土', 'm3'))
    $sameUnitKey = [string]$formulaKeyMethod.Invoke($null, [object[]]@($sameUnitSignature, 'quota', 'TEST-DERIVED'))
    $sameUnitFormulaByKey = $snapshotType.GetField('FormulaByKey', $flags).GetValue($sameUnitSnapshot)
    $derivedRule = [Activator]::CreateInstance($formulaRuleType, $true).PSObject.BaseObject
    foreach ($pair in @{ RuleHash='derived-rule'; TargetUnit='m3'; Template='V0+V1'; Method='2024'; EntryCode='0101-01'; SampleCount=1 }.GetEnumerator()) {
        $formulaRuleType.GetField($pair.Key, $flags).SetValue($derivedRule, $pair.Value)
    }
    $derivedOperands = $formulaRuleType.GetField('Operands', $flags).GetValue($derivedRule)
    for ($i = 0; $i -lt 2; $i++) {
        $derivedOperand = [Activator]::CreateInstance($formulaOperandType, $true).PSObject.BaseObject
        $operandName = if ($i -eq 0) { '主体混凝土' } else { '附加混凝土' }
        $formulaOperandType.GetField('Index', $flags).SetValue($derivedOperand, $i)
        $formulaOperandType.GetField('Name', $flags).SetValue($derivedOperand, $operandName)
        $formulaOperandType.GetField('Unit', $flags).SetValue($derivedOperand, 'm3')
        $formulaOperandType.GetField('Signature', $flags).SetValue($derivedOperand, [string]$signatureMethod.Invoke($null, [object[]]@($operandName, 'm3')))
        [void]$derivedOperands.Add($derivedOperand)
    }
    $sameUnitRules = [Activator]::CreateInstance($ruleListType)
    [void]$sameUnitRules.Add($derivedRule)
    $sameUnitFormulaByKey.Add($sameUnitKey, $sameUnitRules)

    $mapEntryType = $type.GetNestedType('SmartMapEntry', $nestedFlags)
    $sameUnitEntry = [Activator]::CreateInstance($mapEntryType, $true).PSObject.BaseObject
    $sameUnitTarget = [Activator]::CreateInstance($smartTargetType, $true).PSObject.BaseObject
    $smartTargetType.GetField('Kind', $flags).SetValue($sameUnitTarget, 'quota')
    $smartTargetType.GetField('Code', $flags).SetValue($sameUnitTarget, 'TEST-DERIVED')
    [void]$mapEntryType.GetField('Targets', $flags).GetValue($sameUnitEntry).Add($sameUnitTarget)
    [void]$mapEntryType.GetField('LocalContextKeys', $flags).GetValue($sameUnitEntry).Add("TB 10801—2024`n0101-01")

    $projectEntries = [Collections.Generic.Dictionary[string,long]]::new([StringComparer]::OrdinalIgnoreCase)
    $projectEntries.Add('0101-01', [long]1)
    $projectQuotaType = $type.GetNestedType('ProjectQuota', $nestedFlags)
    $quotaDictionaryType = [Collections.Generic.Dictionary``2].MakeGenericType([string], $projectQuotaType)
    $currentQuotaByCode = [Activator]::CreateInstance($quotaDictionaryType, [StringComparer]::OrdinalIgnoreCase)
    $currentQuota = [Activator]::CreateInstance($projectQuotaType, $true).PSObject.BaseObject
    foreach ($pair in @{ Code='TEST-DERIVED'; Name='派生公式定额'; Unit='m3'; QuotaSeq=[long]1; IsLibrary=$false }.GetEnumerator()) {
        $projectQuotaType.GetField($pair.Key, $flags).SetValue($currentQuota, $pair.Value)
    }
    $currentQuotaByCode.Add('TEST-DERIVED', $currentQuota)
    $previewType = $type.GetNestedType('FillPreviewItem', $nestedFlags)
    $previewListType = [Collections.Generic.List``1].MakeGenericType($previewType)
    $appendSmartItems = $type.GetMethod('AppendSmartItems', $flags)
    function Invoke-DerivedPreview($Rows) {
        $previewItems = [Activator]::CreateInstance($previewListType)
        $appendArgs = [object[]]::new(13)
        $appendArgs[0] = $previewItems.PSObject.BaseObject
        $appendArgs[1] = $Rows[0]
        $appendArgs[2] = $Rows.PSObject.BaseObject
        $appendArgs[3] = $sameUnitEntry
        $appendArgs[4] = $sameUnitSnapshot
        $appendArgs[5] = $projectEntries.PSObject.BaseObject
        $appendArgs[6] = $currentQuotaByCode.PSObject.BaseObject
        $appendArgs[7] = $null
        $appendArgs[8] = $false
        $appendArgs[9] = 'test'
        $appendArgs[10] = $sameUnitSignature
        $appendArgs[11] = $null
        $appendArgs[12] = $null
        [void]$appendSmartItems.Invoke($null, $appendArgs)
        return ,$previewItems
    }
    $derivedPreview = Invoke-DerivedPreview $sameUnitRows
    $derivedQuantityArgs = [object[]]::new(3); $derivedQuantityArgs[0] = [string]$derivedPreview[0].QuantityText; $derivedQuantityArgs[1] = [decimal]0; $derivedQuantityArgs[2] = $null
    if (-not [bool]$tryDecimal.Invoke($null, $derivedQuantityArgs) -or [decimal]$derivedQuantityArgs[1] -ne [decimal]12 -or [string]$derivedPreview[0].FormulaTemplate -ne 'V0+V1') {
        throw "多参数派生公式被标准同单位换算短路：Quantity='$($derivedPreview[0].QuantityText)' Formula='$($derivedPreview[0].FormulaTemplate)'"
    }

    $missingRows = [Activator]::CreateInstance($targetRowListType)
    [void]$missingRows.Add($sameUnitRows[0])
    $missingPreview = Invoke-DerivedPreview $missingRows
    if ($missingPreview[0].Selected -or [string]$missingPreview[0].Status -notmatch '公式参数缺失或歧义') {
        throw '多参数派生公式缺参数时不得回退锚点行标准换算。'
    }
    Write-Host 'PASS 多参数派生公式优先完整求值，缺参数时整组待确认'

    # 当前分区、办法和条目精确命中的可信 SQL 公式优先于标准同量纲换算。
    $formulaRuleType.GetField('RuleHash', $flags).SetValue($derivedRule, 'linear-rule')
    $formulaRuleType.GetField('Template', $flags).SetValue($derivedRule, 'V0*0.2')
    $derivedOperands.RemoveAt(1)
    $linearPreview = Invoke-DerivedPreview $sameUnitRows
    if ([string]$linearPreview[0].QuantityText -ne '10*0.2' -or [string]$linearPreview[0].FormulaTemplate -ne 'V0*0.2') {
        throw '当前办法和条目精确命中的可信 SQL 公式未优先采用。'
    }
    Write-Host 'PASS 当前办法和条目精确命中的可信 SQL 公式优先采用'

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
    $entryNameField = $candidateScoreType.GetField('EntryName', $flags)
    if ($null -eq $entryNameField) { throw '候选分数缺少条目名称字段' }
    $smartTargetType = $type.GetNestedType('SmartBoxTarget', $nestedFlags)
    $targetIdentity = $type.GetMethod('BuildLearningTargetIdentityKey', $flags)
    $shDisposalIdentity = [string]$targetIdentity.Invoke($null, @('quota', 'SH', '消纳费', 'm³'))
    $shPeIdentity = [string]$targetIdentity.Invoke($null, @('quota', 'SH', 'PE', 'm'))
    if ($shDisposalIdentity -eq $shPeIdentity -or
        [string]$targetIdentity.Invoke($null, @('quota', 'LY-89', '旧名称', '100m3')) -ne 'quota:LY-89') {
        throw '推荐端没有按名称和单位隔离通用辅助代码，或错误拆分普通定额'
    }

    $projectQuotaType = $type.GetNestedType('ProjectQuota', $nestedFlags)
    $quotaDictionaryType = [Collections.Generic.Dictionary``2].MakeGenericType([string], $projectQuotaType)
    $currentMetadata = [Activator]::CreateInstance($quotaDictionaryType)
    $currentKey = $type.GetMethod('BuildSmartCurrentQuotaKey', $flags)
    function Add-CurrentQuota([object]$dictionary, [string]$code, [string]$name, [string]$unit) {
        $quota = [Activator]::CreateInstance($projectQuotaType, $true).PSObject.BaseObject
        $projectQuotaType.GetField('Code', $flags).SetValue($quota, $code)
        $projectQuotaType.GetField('Name', $flags).SetValue($quota, $name)
        $projectQuotaType.GetField('Unit', $flags).SetValue($quota, $unit)
        $key = [string]$currentKey.Invoke($null, @($code, $name, $unit))
        $dictionary[$key] = $quota
    }
    Add-CurrentQuota $currentMetadata 'SH' '消纳费' 'm³'
    Add-CurrentQuota $currentMetadata 'LY-89' '挖掘机挖装石' '100m3'

    $safetyEntry = [Activator]::CreateInstance($mapEntryType, $true).PSObject.BaseObject
    $safetyTargets = $mapEntryType.GetField('Targets', $flags).GetValue($safetyEntry)
    $shTarget = [Activator]::CreateInstance($smartTargetType, $true).PSObject.BaseObject
    foreach ($definition in @(@('Kind', 'quota'), @('Code', 'SH'), @('Name', '消纳费'), @('Unit', 'm3'))) {
        $smartTargetType.GetField($definition[0], $flags).SetValue($shTarget, $definition[1])
    }
    [void]$safetyTargets.Add($shTarget)
    $usableEntry = $type.GetMethod('IsSmartMapEntryUsableInCurrentProject', $flags)
    $usableArgs = New-Object 'object[]' 3
    $usableArgs[0] = $safetyEntry
    $usableArgs[1] = '弃渣外运'
    $usableArgs[2] = $currentMetadata.PSObject.BaseObject
    if ($usableEntry.Invoke($null, $usableArgs)) { throw '纯 SH 辅助组件不得独立进入普通推荐' }

    $missingIdentityEntry = [Activator]::CreateInstance($mapEntryType, $true).PSObject.BaseObject
    $missingIdentityTarget = [Activator]::CreateInstance($smartTargetType, $true).PSObject.BaseObject
    foreach ($definition in @(@('Kind', 'quota'), @('Code', 'SH'), @('Name', ''), @('Unit', 'm3'))) {
        $smartTargetType.GetField($definition[0], $flags).SetValue($missingIdentityTarget, $definition[1])
    }
    [void]$mapEntryType.GetField('Targets', $flags).GetValue($missingIdentityEntry).Add($missingIdentityTarget)
    $usableArgs[0] = $missingIdentityEntry
    if ($usableEntry.Invoke($null, $usableArgs)) { throw '缺少名称或单位的辅助定额不得进入推荐' }

    $usableArgs[0] = $safetyEntry
    $lyTarget = [Activator]::CreateInstance($smartTargetType, $true).PSObject.BaseObject
    foreach ($definition in @(@('Kind', 'quota'), @('Code', 'LY-89'), @('Name', '历史名称可变化'), @('Unit', '100m3'))) {
        $smartTargetType.GetField($definition[0], $flags).SetValue($lyTarget, $definition[1])
    }
    [void]$safetyTargets.Add($lyTarget)
    if (-not $usableEntry.Invoke($null, $usableArgs)) { throw '辅助名称单位精确一致的混合组件应保留完整候选' }

    $mismatchMetadata = [Activator]::CreateInstance($quotaDictionaryType)
    Add-CurrentQuota $mismatchMetadata 'SH' 'PE' 'm'
    Add-CurrentQuota $mismatchMetadata 'LY-89' '挖掘机挖装石' '100m3'
    $usableArgs[2] = $mismatchMetadata.PSObject.BaseObject
    if ($usableEntry.Invoke($null, $usableArgs)) { throw '辅助名称或单位不一致时必须过滤整个混合组件' }

    $sfEntry = [Activator]::CreateInstance($mapEntryType, $true).PSObject.BaseObject
    $sfTarget = [Activator]::CreateInstance($smartTargetType, $true).PSObject.BaseObject
    foreach ($definition in @(@('Kind', 'quota'), @('Code', 'SF'), @('Name', '设备购置费'), @('Unit', '元'))) {
        $smartTargetType.GetField($definition[0], $flags).SetValue($sfTarget, $definition[1])
    }
    [void]$mapEntryType.GetField('Targets', $flags).GetValue($sfEntry).Add($sfTarget)
    $sfMetadata = [Activator]::CreateInstance($quotaDictionaryType)
    Add-CurrentQuota $sfMetadata 'SF' '设备购置费' '元'
    $usableArgs[0] = $sfEntry
    $usableArgs[1] = '设备购置费'
    $usableArgs[2] = $sfMetadata.PSObject.BaseObject
    if (-not $usableEntry.Invoke($null, $usableArgs)) { throw 'SF 设备购置费的既有业务例外被误过滤' }
    Write-Host 'PASS 通用辅助代码按名称单位隔离，纯辅助与混合组件都做原子过滤'

    $firstMapEntry = $candidateScoreType.GetField('Entry', $flags).GetValue($scoreList[0])
    $targetResolutionType = $type.GetNestedType('SmartTargetEntryResolution', $nestedFlags)
    $targetResolutions = $candidateScoreType.GetField('TargetEntries', $flags).GetValue($scoreList[0])
    foreach ($definition in @(
        @('SH', '0401-01', '弃渣工程'),
        @('1009001002*1.224', '0309-01-03-03', '桥涵工程'),
        @('LY-89', '0309-01-03-03', '桥涵工程')
    )) {
        $targetCode = $definition[0]
        $smartTarget = [Activator]::CreateInstance($smartTargetType, $true).PSObject.BaseObject
        $smartTargetType.GetField('Code', $flags).SetValue($smartTarget, $targetCode)
        [void]$mapEntryType.GetField('Targets', $flags).GetValue($firstMapEntry).Add($smartTarget)
        $targetResolution = [Activator]::CreateInstance($targetResolutionType, $true).PSObject.BaseObject
        $targetResolutionType.GetField('Target', $flags).SetValue($targetResolution, $smartTarget)
        $targetResolutionType.GetField('EntryCode', $flags).SetValue($targetResolution, $definition[1])
        $targetResolutionType.GetField('EntryName', $flags).SetValue($targetResolution, $definition[2])
        [void]$targetResolutions.Add($targetResolution)
    }
    $candidateScoreType.GetField('EntryCode', $flags).SetValue($scoreList[0], '0309-01-03-03')
    $entryNameField.SetValue($scoreList[0], '弃渣外运')
    $snapshotType = $type.GetNestedType('SmartLearningSnapshot', $nestedFlags)
    $snapshot = [Activator]::CreateInstance($snapshotType, $true).PSObject.BaseObject
    $candidateLabel = $type.GetMethod('BuildSmartCandidateLabel', $flags).Invoke($null, @($snapshot, $scoreList[0]))
    if ($candidateLabel -ne 'LY-89（桥涵工程 0309-01-03-03） + 1009001002*1.224（桥涵工程 0309-01-03-03） + SH（弃渣工程 0401-01）' -or
        $candidateLabel -match '权重|当前办法') {
        throw "候选下拉应按定额/材料/辅助排序并逐目标显示条目，实际：$candidateLabel"
    }
    $scopeType = $type.GetNestedType('SmartLearningScope', $nestedFlags)
    $classifiedEntryCode = $type.GetMethod('IsSmartClassifiedEntryCode', $flags)
    if (-not $classifiedEntryCode.Invoke($null, @('0309-01-03-03')) -or
        $classifiedEntryCode.Invoke($null, @('ENTRY-test')) -or
        $classifiedEntryCode.Invoke($null, @('EN'))) {
        throw '推荐学习库没有过滤非数字测试条目编号'
    }
    $scopeCodeBuilder = $type.GetMethod('BuildSmartLearningScopeCodes', $flags)
    $learnedEntryCodes = [Collections.Generic.List[string]]::new()
    [void]$learnedEntryCodes.Add('0202-01-01-04-01-02')
    [void]$learnedEntryCodes.Add('0204-01-01-02-01-05')
    [void]$learnedEntryCodes.Add('ENTRY-test')
    $scopeCodeArgs = New-Object 'object[]' 1
    $scopeCodeArgs[0] = $learnedEntryCodes.PSObject.BaseObject
    $scopeCodes = @($scopeCodeBuilder.Invoke($null, $scopeCodeArgs) | Sort-Object)
    if (($scopeCodes -join '|') -ne '02|0202|0204') {
        throw "推荐学习库只应生成专业和分部两级目录，实际：$($scopeCodes -join '|')"
    }
    $scope = [Activator]::CreateInstance($scopeType, $true).PSObject.BaseObject
    $scopeType.GetField('Kind', $flags).SetValue($scope, 'Entry')
    $scopeType.GetField('EntryCode', $flags).SetValue($scope, '03')
    $scopeMatch = $type.GetMethod('SmartEntryCodeMatchesScope', $flags)
    if (-not $scopeMatch.Invoke($null, @('0309-01-03-03', $scope)) -or
        $scopeMatch.Invoke($null, @('0204-01-01', $scope))) {
        throw '前两位专业范围没有按条目边界包含下级'
    }
    $scopeType.GetField('EntryCode', $flags).SetValue($scope, '0309-01-03')
    if (-not $scopeMatch.Invoke($null, @('0309-01-03-01', $scope)) -or
        $scopeMatch.Invoke($null, @('0309-01-04-01', $scope))) {
        throw '完整条目范围没有包含自身下级或错误纳入相邻条目'
    }
    $mapEntryListType = [Collections.Generic.List``1].MakeGenericType($mapEntryType)
    $scopeHits = [Activator]::CreateInstance($mapEntryListType)
    foreach ($definition in @(@('box-bridge', 10), @('box-road', 100), @('box-unclassified', 80))) {
        $entry = [Activator]::CreateInstance($mapEntryType, $true).PSObject.BaseObject
        $mapEntryType.GetField('BoxId', $flags).SetValue($entry, $definition[0])
        $mapEntryType.GetField('Weight', $flags).SetValue($entry, $definition[1])
        [void]$scopeHits.Add($entry)
    }
    $scopeMap = $snapshotType.GetField('ScopeEntriesByBox', $flags).GetValue($snapshot)
    $scopeSetType = $scopeMap.GetType().GetGenericArguments()[1]
    $bridgeCodes = [Activator]::CreateInstance($scopeSetType)
    [void]$bridgeCodes.Add('0309-01-03-01')
    $scopeMap.Add('box-bridge', $bridgeCodes)
    $roadCodes = [Activator]::CreateInstance($scopeSetType)
    [void]$roadCodes.Add('0204-01-01')
    $scopeMap.Add('box-road', $roadCodes)
    $filterByScope = $type.GetMethod('FilterSmartHitsByScope', $flags)
    $scopeType.GetField('EntryCode', $flags).SetValue($scope, '03')
    $filterArgs = New-Object 'object[]' 3
    $filterArgs[0] = $snapshot
    $filterArgs[1] = $scopeHits.PSObject.BaseObject
    $filterArgs[2] = $scope
    $professionalHits = @($filterByScope.Invoke($null, $filterArgs))
    if ($professionalHits.Count -ne 1 -or
        $mapEntryType.GetField('BoxId', $flags).GetValue($professionalHits[0]) -ne 'box-bridge') {
        throw '专业学习库没有优先隔离专业内组件，或被全库高权重组件覆盖'
    }
    $scopeType.GetField('Kind', $flags).SetValue($scope, 'Unclassified')
    $scopeType.GetField('EntryCode', $flags).SetValue($scope, '')
    $unclassifiedHits = @($filterByScope.Invoke($null, $filterArgs))
    if ($unclassifiedHits.Count -ne 1 -or
        $mapEntryType.GetField('BoxId', $flags).GetValue($unclassifiedHits[0]) -ne 'box-unclassified') {
        throw '未归类学习库没有排除已有专业归类的组件'
    }
    Write-Host 'PASS 专业学习库范围按条目边界包含下级'

    $candidateType = $type.GetNestedType('NameQuotaCandidateGroup', $nestedFlags)
    $candidateListType = [System.Collections.Generic.List``1].MakeGenericType($candidateType)
    $candidateList = [Activator]::CreateInstance($candidateListType)
    foreach ($candidateDefinition in @(
        @('box-high', $candidateLabel),
        @('box-low', $candidateLabel),
        @('box-other', 'LY-90（桥涵 0309-01-03-03）')
    )) {
        $candidate = [Activator]::CreateInstance($candidateType)
        $candidateType.GetField('Key', $flags).SetValue($candidate, $candidateDefinition[0])
        $candidateType.GetField('Label', $flags).SetValue($candidate, $candidateDefinition[1])
        [void]$candidateList.Add($candidate)
    }
    $deduplicateMethod = $type.GetMethod('DeduplicateSmartCandidatesByLabel', $flags)
    $deduplicateArgs = New-Object 'object[]' 1
    $deduplicateArgs[0] = $candidateList.PSObject.BaseObject
    $deduplicated = $deduplicateMethod.Invoke($null, $deduplicateArgs)
    if ($deduplicated.Count -ne 2) {
        throw "同显示文案候选未去重: $($deduplicated.Count)"
    }
    if ($candidateType.GetField('Key', $flags).GetValue($deduplicated[0]) -ne 'box-high') {
        throw '同显示文案候选去重未保留排序第一项'
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
