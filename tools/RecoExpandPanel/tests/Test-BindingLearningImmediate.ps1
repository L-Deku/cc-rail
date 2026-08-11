$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..')).Path
$excelLinkPath = Join-Path $repoRoot 'tools\RecoExpandPanel\ExcelLinkFeature.cs'
$autoMatchPath = Join-Path $repoRoot 'tools\RecoExpandPanel\AutoMatchFeature.cs'
$smartFillPath = Join-Path $repoRoot 'tools\RecoExpandPanel\SmartFillFeature.cs'
$templatePanelPath = Join-Path $repoRoot 'tools\RecoExpandPanel\TemplateFillPanel.cs'
$excelLink = [System.IO.File]::ReadAllText($excelLinkPath, [System.Text.Encoding]::UTF8)
$autoMatch = [System.IO.File]::ReadAllText($autoMatchPath, [System.Text.Encoding]::UTF8)
$smartFill = [System.IO.File]::ReadAllText($smartFillPath, [System.Text.Encoding]::UTF8)
$templatePanel = [System.IO.File]::ReadAllText($templatePanelPath, [System.Text.Encoding]::UTF8)

function Assert-Contains([string]$Text, [string]$Expected, [string]$Message) {
    if (-not $Text.Contains($Expected)) { throw $Message }
}

Assert-Contains $excelLink 'ExtractPositiveAdditiveCellAddresses' '缺少正向相加单元格拆分入口。'
Assert-Contains $excelLink 'BuildBindingFeedbackGroups' '缺少按原始表达式构建独立学习组的入口。'
if ($excelLink.Contains('(entryScope ?? "").Trim()')) { throw '绑定组件仍按组级条目拆分，跨条目目标无法组成完整组件。' }
Assert-Contains $excelLink 'string targetEntryCode = LearningPartitionIdentity.NormalizeLearningEntryCode(' 'SQL 学习组件没有规范化目标条目上下文。'
Assert-Contains $excelLink 'GetMappingFeedbackTargetEntryCode(group, target));' 'SQL 学习组件没有按目标保存条目上下文。'
Assert-Contains $excelLink 'EntryCode = !String.IsNullOrWhiteSpace(target.EntryCode) ? target.EntryCode : group.EntryCode ?? ""' '绑定学习组没有优先传递 ExcelQuotaLink 的目标条目或兼容空值旧字段。'
Assert-Contains $excelLink 'group.Targets.Add' '表达式学习组没有保留完整组件框目标。'
Assert-Contains $excelLink 'public string QuotaUnit { get; set; }' 'ExcelQuotaLink 没有持久化定额单位。'
Assert-Contains $excelLink 'public string EntryName { get; set; }' 'ExcelQuotaLink 没有持久化条目名称。'
Assert-Contains $excelLink 'PopulateExcelQuotaLinkLearningContext(conn, link);' '手动/批量绑定没有补齐编制办法和条目上下文。'
Assert-Contains $excelLink 'link.QuotaUnit = GetRowValue(row, "单位", "定额单位", "计量单位");' '手动/批量/快速绑定没有传递定额单位。'
Assert-Contains $excelLink 'item.Link.QuotaUnit = item.QuotaUnit;' '自动匹配预览保存时没有把定额单位写入 ExcelQuotaLink。'
Assert-Contains $autoMatch 'link.QuotaUnit = quotaUnit;' '自动匹配绑定没有传递定额单位。'
Assert-Contains $autoMatch 'link.EntryName = ReadAutoMatchReaderText(reader, 9);' '自动匹配绑定没有传递条目名称。'
Assert-Contains $autoMatch 'link.Method = projectMethod;' '自动匹配绑定没有传递编制办法。'
if ($smartFill.Contains('MergePendingLocalMappingsIntoSmartSnapshot') -or $smartFill.Contains('LoadMappingBoxRows(')) {
    throw '推荐预览仍会叠加或回退本机学习关系。'
}
Assert-Contains $smartFill 'LoadCurrentSmartQuotaMetadata' '推荐预览没有从当前运行版本读取定额元数据。'
if ($smartFill.Contains('BuildNameDrivenQtyText(row.QuantityText, row.Unit, target.Unit)')) { throw '推荐数量仍在使用 SQL 历史 target_unit 换算。' }
Assert-Contains $templatePanel 'FeedbackNameMatches(groupLeader.TemplateName, replacements' '模板铺量右键绑定后没有立即写入学习关系。'

$dll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL)) { $env:RECO_EXPAND_DLL } else { Join-Path $repoRoot 'RecoQuotaRecommend\bin\RecoExpandPanel.dll' }
if (-not (Test-Path -LiteralPath $dll)) { throw "找不到 $dll，先构建" }
$dllDir = Split-Path -Parent $dll
foreach ($dependency in @('NPOI.dll', 'NPOI.OpenXmlFormats.dll', 'NPOI.OpenXml4Net.dll', 'NPOI.OOXML.dll', 'ICSharpCode.SharpZipLib.dll')) {
    $dependencyPath = @($dllDir, (Join-Path $repoRoot 'RecoQuotaRecommend\bin')) |
        ForEach-Object { Join-Path $_ $dependency } |
        Where-Object { Test-Path -LiteralPath $_ } |
        Select-Object -First 1
    if (Test-Path -LiteralPath $dependencyPath) { [void][System.Reflection.Assembly]::LoadFrom($dependencyPath) }
}

$type = [System.Reflection.Assembly]::LoadFrom($dll).GetType('RecoNet.FormPanel', $true)
$flags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$extract = $type.GetMethod('ExtractPositiveAdditiveCellAddresses', $flags)
if ($null -eq $extract) { throw '编译结果缺少 ExtractPositiveAdditiveCellAddresses。' }
$hasRepeatedCellInTerm = $type.GetMethod('HasRepeatedCellReferenceWithinTerm', $flags)
if ($null -eq $hasRepeatedCellInTerm) { throw '编译结果缺少重复单元格非线性判定。' }
$tryScaleFactor = $type.GetMethod('TryExtractPositiveCellScaleFactor', $flags)
if ($null -eq $tryScaleFactor) { throw '编译结果缺少线性系数提取入口。' }

function Assert-Addresses([string]$Expression, [string[]]$Expected) {
    $actual = @($extract.Invoke($null, @($Expression)))
    if (($actual -join ',') -ne ($Expected -join ',')) {
        throw "表达式拆分错误：$Expression -> '$($actual -join ',')'，预期 '$($Expected -join ',')'"
    }
}

Assert-Addresses 'F10/1000+F11/1000' @('F10', 'F11')
Assert-Addresses '(F10/1000)+(F11/1000)' @('F10', 'F11')
Assert-Addresses 'F10+F10' @('F10')
Assert-Addresses 'F10+F11-F12' @('F10', 'F11')
Assert-Addresses 'F10*F11' @()
if (-not [bool]$hasRepeatedCellInTerm.Invoke($null, @('F10*F10*3.14'))) { throw '同一单元格在同项内重复引用未识别为非线性公式。' }
if ([bool]$hasRepeatedCellInTerm.Invoke($null, @('F10+F10'))) { throw '独立正向加项不应因相同单元格被误判为非线性公式。' }
$scaleArgs = [object[]]::new(3); $scaleArgs[0] = 'F10*F10*3.14'; $scaleArgs[1] = 'F10'; $scaleArgs[2] = [decimal]0
if ([bool]$tryScaleFactor.Invoke($null, $scaleArgs.PSObject.BaseObject)) { throw '平方公式不应被简化为 V0*3.14 线性系数。' }
Write-Host 'PASS 正向相加单元格拆分且不跨负项/乘积学习'

$fixturePath = Join-Path ([IO.Path]::GetTempPath()) ('reco-binding-learning-' + [Guid]::NewGuid().ToString('N') + '.xlsx')
$book = $null
try {
    $book = New-Object NPOI.XSSF.UserModel.XSSFWorkbook
    $sheet = $book.CreateSheet('Sheet1')
    foreach ($definition in @(
        [pscustomobject]@{ Row = 0; Name = 'HRB400钢筋'; Unit = 'kg'; Value = 2700.2 },
        [pscustomobject]@{ Row = 1; Name = 'HPB300钢筋'; Unit = 'kg'; Value = 537 },
        [pscustomobject]@{ Row = 2; Name = 'HRB400钢筋'; Unit = 'kg'; Value = 1100 },
        [pscustomobject]@{ Row = 3; Name = 'HPB300钢筋'; Unit = 'kg'; Value = 220 }
    )) {
        $row = $sheet.CreateRow($definition.Row)
        $row.CreateCell(0).SetCellValue($definition.Name)
        $row.CreateCell(1).SetCellValue($definition.Unit)
        $row.CreateCell(5).SetCellValue([double]$definition.Value)
    }
    $partialSheet = $book.CreateSheet('Partial')
    $partialNamedRow = $partialSheet.CreateRow(0)
    $partialNamedRow.CreateCell(0).SetCellValue('有效别名')
    $partialNamedRow.CreateCell(1).SetCellValue('m3')
    $partialNamedRow.CreateCell(5).SetCellValue([double]10)
    $partialUnnamedRow = $partialSheet.CreateRow(1)
    $partialUnnamedRow.CreateCell(1).SetCellValue('m3')
    $partialUnnamedRow.CreateCell(5).SetCellValue([double]20)
    $nonlinearSheet = $book.CreateSheet('Nonlinear')
    $nonlinearRow = $nonlinearSheet.CreateRow(0)
    $nonlinearRow.CreateCell(0).SetCellValue('桩半径')
    $nonlinearRow.CreateCell(1).SetCellValue('m')
    $nonlinearRow.CreateCell(5).SetCellValue([double]2)
    $stream = [IO.File]::Create($fixturePath)
    try { $book.Write($stream) } finally { $stream.Dispose() }

    $readTargetRows = $type.GetMethod('ReadTargetQtyRowsWithChapters', $flags)
    if ($null -eq $readTargetRows) { throw '缺少目标工程量读取入口。' }
    $targetArgs = [object[]]::new(4)
    $targetArgs[0] = $fixturePath.PSObject.BaseObject
    $targetArgs[1] = ([string]'Sheet1').PSObject.BaseObject
    $targetArgs[2] = 6
    $targetRows = @($readTargetRows.Invoke($null, $targetArgs.PSObject.BaseObject))
    if ($targetRows.Count -ne 4) { throw "目标工程量应读取4行，实际 $($targetRows.Count) 行。" }
    $targetRowType = $targetRows[0].GetType()
    if ($targetRowType.GetField('RawName', $flags).GetValue($targetRows[0]) -ne 'HRB400钢筋' -or
        $targetRowType.GetField('Unit', $flags).GetValue($targetRows[0]) -ne 'kg') {
        throw '推荐目标读取没有把工程量名称与单位分开。'
    }
    Write-Host 'PASS 推荐目标工程量名称与单位分开读取'

    $linkType = $type.GetNestedType('ExcelQuotaLink', $flags)
    $dictionaryType = [Collections.Generic.Dictionary``2].MakeGenericType($linkType, [string])
    $links = [Activator]::CreateInstance($dictionaryType)
    foreach ($quota in @(
        [pscustomobject]@{ Code = 'QY-317'; Name = '盖板钢筋'; Unit = 't' },
        [pscustomobject]@{ Code = 'QY-318'; Name = '盖板安装'; Unit = '10m3' }
    )) {
        $link = [Activator]::CreateInstance($linkType)
        $linkType.GetProperty('ExcelPath', $flags).SetValue($link, $fixturePath, $null)
        $linkType.GetProperty('WorksheetName', $flags).SetValue($link, 'Sheet1', $null)
        $linkType.GetProperty('CellAddress', $flags).SetValue($link, 'F1', $null)
        $linkType.GetProperty('Expression', $flags).SetValue($link, 'F1/1000+F2/1000', $null)
        $linkType.GetProperty('QuotaCode', $flags).SetValue($link, $quota.Code, $null)
        $linkType.GetProperty('QuotaName', $flags).SetValue($link, $quota.Name, $null)
        $linkType.GetProperty('QuotaUnit', $flags).SetValue($link, $quota.Unit, $null)
        $linkType.GetProperty('EntryCode', $flags).SetValue($link, '0101-01', $null)
        $linkType.GetProperty('EntryName', $flags).SetValue($link, '沉井工程', $null)
        $linkType.GetProperty('Method', $flags).SetValue($link, '2024', $null)
        $links.Add($link, 'HRB400钢筋 kg')
    }
    $secondLink = [Activator]::CreateInstance($linkType)
    $linkType.GetProperty('ExcelPath', $flags).SetValue($secondLink, $fixturePath, $null)
    $linkType.GetProperty('WorksheetName', $flags).SetValue($secondLink, 'Sheet1', $null)
    $linkType.GetProperty('CellAddress', $flags).SetValue($secondLink, 'F3', $null)
    $linkType.GetProperty('Expression', $flags).SetValue($secondLink, 'F3/1000+F4/1000', $null)
    $linkType.GetProperty('QuotaCode', $flags).SetValue($secondLink, 'QY-317', $null)
    $linkType.GetProperty('QuotaName', $flags).SetValue($secondLink, '另一块盖板钢筋', $null)
    $linkType.GetProperty('QuotaUnit', $flags).SetValue($secondLink, 't', $null)
    $linkType.GetProperty('EntryCode', $flags).SetValue($secondLink, '0201-01', $null)
    $linkType.GetProperty('EntryName', $flags).SetValue($secondLink, '另一处钢筋', $null)
    $linkType.GetProperty('Method', $flags).SetValue($secondLink, '2024', $null)
    $links.Add($secondLink, 'HRB400钢筋 kg')

    # 同一个原始绑定表达式就是一个组件；跨条目目标靠各自 EntryCode/EntryName 保持身份，不能再按组级条目拆散。
    $otherEntryLink = [Activator]::CreateInstance($linkType)
    $linkType.GetProperty('ExcelPath', $flags).SetValue($otherEntryLink, $fixturePath, $null)
    $linkType.GetProperty('WorksheetName', $flags).SetValue($otherEntryLink, 'Sheet1', $null)
    $linkType.GetProperty('CellAddress', $flags).SetValue($otherEntryLink, 'F1', $null)
    $linkType.GetProperty('Expression', $flags).SetValue($otherEntryLink, 'F1/1000+F2/1000', $null)
    $linkType.GetProperty('QuotaCode', $flags).SetValue($otherEntryLink, 'QY-999', $null)
    $linkType.GetProperty('QuotaName', $flags).SetValue($otherEntryLink, '另一条目同源表达式', $null)
    $linkType.GetProperty('QuotaUnit', $flags).SetValue($otherEntryLink, 't', $null)
    $linkType.GetProperty('EntryCode', $flags).SetValue($otherEntryLink, '0301-01', $null)
    $linkType.GetProperty('EntryName', $flags).SetValue($otherEntryLink, '独立条目', $null)
    $linkType.GetProperty('Method', $flags).SetValue($otherEntryLink, '2024', $null)
    $links.Add($otherEntryLink, 'HRB400钢筋 kg')

    $buildGroups = $type.GetMethod('BuildBindingFeedbackGroups', $flags)
    if ($null -eq $buildGroups) { throw '编译结果缺少 BuildBindingFeedbackGroups。' }
    $groups = @($buildGroups.Invoke($null, @($links.PSObject.BaseObject)))
    if ($groups.Count -ne 4) { throw "跨条目目标应按同一原始表达式组成完整组件，预期四个正向来源别名，实际 $($groups.Count) 套。" }
    $actual = @()
    foreach ($group in $groups) {
        $groupType = $group.GetType()
        $name = [string]$groupType.GetField('QuantityName', $flags).GetValue($group)
        $unit = [string]$groupType.GetField('QuantityUnit', $flags).GetValue($group)
        $rowNo = [int]$groupType.GetField('ExcelRow', $flags).GetValue($group)
        $sourceCell = [string]$groupType.GetField('SourceCell', $flags).GetValue($group)
        $entryCode = [string]$groupType.GetField('EntryCode', $flags).GetValue($group)
        $entryName = [string]$groupType.GetField('EntryName', $flags).GetValue($group)
        $method = [string]$groupType.GetField('Method', $flags).GetValue($group)
        $targets = @($groupType.GetField('Targets', $flags).GetValue($group))
        $targetFacts = @($targets | ForEach-Object {
            $targetType = $_.GetType()
            ([string]$targetType.GetField('Code', $flags).GetValue($_)) + ':' +
                ([string]$targetType.GetField('Unit', $flags).GetValue($_)) + ':' +
                ([string]$targetType.GetField('EntryCode', $flags).GetValue($_))
        } | Sort-Object)
        $actual += "$sourceCell|$name|$unit|$rowNo|$method|$entryCode|$entryName|$($targetFacts -join ',')"
    }
    $actual = @($actual | Sort-Object)
    $expected = @(
        'F1|HRB400钢筋|kg|1|2024|0101-01|沉井工程|QY-317:t:0101-01,QY-318:10m3:0101-01,QY-999:t:0301-01',
        'F2|HPB300钢筋|kg|2|2024|0101-01|沉井工程|QY-317:t:0101-01,QY-318:10m3:0101-01,QY-999:t:0301-01',
        'F3|HRB400钢筋|kg|3|2024|0201-01|另一处钢筋|QY-317:t:0201-01',
        'F4|HPB300钢筋|kg|4|2024|0201-01|另一处钢筋|QY-317:t:0201-01'
    ) | Sort-Object
    if (($actual -join ';') -ne ($expected -join ';')) {
        throw "独立别名/单位/行号不正确：'$($actual -join ';')'"
    }
    Write-Host 'PASS 工程量/定额单位分开、组件共用条目且重复 HRB/HPB 不跨表达式合并'

    $partialLinks = [Activator]::CreateInstance($dictionaryType)
    $partialLink = [Activator]::CreateInstance($linkType)
    foreach ($pair in @{
        ExcelPath=$fixturePath; WorksheetName='Partial'; CellAddress='F1'; Expression='F1+F2';
        QuotaCode='TEST-PARTIAL'; QuotaName='部分别名定额'; QuotaUnit='m3'; EntryCode='0101-01'; EntryName='测试条目'; Method='2024'
    }.GetEnumerator()) { $linkType.GetProperty($pair.Key, $flags).SetValue($partialLink, $pair.Value, $null) }
    $partialLinks.Add($partialLink, '有效别名 m3')
    $partialGroups = @($buildGroups.Invoke($null, @($partialLinks.PSObject.BaseObject)))
    $partialSources = @($partialGroups | ForEach-Object { [string]$_.GetType().GetField('SourceCell', $flags).GetValue($_) })
    $partialNames = @($partialGroups | ForEach-Object { [string]$_.GetType().GetField('QuantityName', $flags).GetValue($_) })
    if ($partialGroups.Count -ne 1 -or $partialSources[0] -ne 'F1') {
        throw "普通正向加项的无名行不应拖掉其他有效别名：Count=$($partialGroups.Count), Sources=$($partialSources -join ','), Names=$($partialNames -join ',')"
    }

    $sameDimensionLinks = [Activator]::CreateInstance($dictionaryType)
    $sameDimensionLink = [Activator]::CreateInstance($linkType)
    foreach ($pair in @{
        ExcelPath=$fixturePath; WorksheetName='Sheet1'; CellAddress='F1'; Expression='(F1+F2)*1.05';
        QuotaCode='TEST-105'; QuotaName='同量纲业务系数'; QuotaUnit='t'; EntryCode='0101-01'; EntryName='测试条目'; Method='2024'
    }.GetEnumerator()) { $linkType.GetProperty($pair.Key, $flags).SetValue($sameDimensionLink, $pair.Value, $null) }
    $sameDimensionLinks.Add($sameDimensionLink, 'HRB400钢筋 kg')
    $sameDimensionOtherTarget = [Activator]::CreateInstance($linkType)
    foreach ($pair in @{
        ExcelPath=$fixturePath; WorksheetName='Sheet1'; CellAddress='F1'; Expression='(F1+F2)*1.05';
        QuotaCode='TEST-105-OTHER'; QuotaName='组件内另一目标'; QuotaUnit='m3'; EntryCode='0101-01'; EntryName='测试条目'; Method='2024'
    }.GetEnumerator()) { $linkType.GetProperty($pair.Key, $flags).SetValue($sameDimensionOtherTarget, $pair.Value, $null) }
    $sameDimensionLinks.Add($sameDimensionOtherTarget, 'HRB400钢筋 kg 另一目标')
    if (@($buildGroups.Invoke($null, @($sameDimensionLinks.PSObject.BaseObject))).Count -ne 0) {
        throw '含同量纲业务系数目标的组件框不得拆成残缺关系进入共享学习。'
    }

    $nonlinearLinks = [Activator]::CreateInstance($dictionaryType)
    $nonlinearLink = [Activator]::CreateInstance($linkType)
    foreach ($pair in @{
        ExcelPath=$fixturePath; WorksheetName='Nonlinear'; CellAddress='F1'; Expression='F1*F1*3.14';
        QuotaCode='TEST-AREA'; QuotaName='圆面积'; QuotaUnit='m2'; EntryCode='0101-01'; EntryName='测试条目'; Method='2024'
    }.GetEnumerator()) { $linkType.GetProperty($pair.Key, $flags).SetValue($nonlinearLink, $pair.Value, $null) }
    $nonlinearLinks.Add($nonlinearLink, '桩半径 m')
    $nonlinearGroups = @($buildGroups.Invoke($null, @($nonlinearLinks.PSObject.BaseObject)))
    if ($nonlinearGroups.Count -ne 1) { throw '重复单元格跨量纲公式未学习。' }
    $nonlinearTargets = @($nonlinearGroups[0].GetType().GetField('Targets', $flags).GetValue($nonlinearGroups[0]))
    if ([string]$nonlinearTargets[0].GetType().GetField('FormulaTemplate', $flags).GetValue($nonlinearTargets[0]) -ne 'V0*V0*3.14') {
        throw '重复单元格平方项未保留为 V0*V0*3.14。'
    }
    Write-Host 'PASS 普通别名逐个学习、同量纲业务系数不共享、重复参数保留非线性公式'
}
finally {
    if ($book -ne $null) { $book.Close() }
    if (Test-Path -LiteralPath $fixturePath) { Remove-Item -LiteralPath $fixturePath -Force }
}

$normalizeLearningSignature = $type.GetMethod('NormalizeSmartLearningSignature', $flags)
if ($normalizeLearningSignature.Invoke($null, @('钢筋|KG')) -ne '钢筋|' -or
    $normalizeLearningSignature.Invoke($null, @('钢筋|T')) -ne '钢筋|') {
    throw '旧名称|单位签名没有归并为名称级签名。'
}
if ($normalizeLearningSignature.Invoke($null, @('泥浆　外运（运距10km），Ф560×33.2mm|M3')) -ne
    '泥浆外运(运距10KM),Φ560X33.2MM|') {
    throw '读取存量学习签名时没有应用统一字符归一化。'
}
$buildQuantitySignature = $type.GetMethod('BuildSmartQuantitySignature', $flags)
if ($buildQuantitySignature.Invoke($null, @('HRB400钢筋 kg', 'kg')) -ne 'HRB400钢筋|') {
    throw '旧名称尾部嵌入单位没有借助 QuantityAlias 归并为名称级签名。'
}
$formatBaseline = [string]$buildQuantitySignature.Invoke($null, @('泥浆外运(运距10km),Φ560X33.2mm', 'm3'))
$formatFullWidth = [string]$buildQuantitySignature.Invoke($null, @(" 泥浆　外运（运距10km），Ф560×33.2mm ", 'm3'))
$formatLowerPhi = [string]$buildQuantitySignature.Invoke($null, @('泥浆外运（运距10km），φ560ｘ33.2mm', 'm3'))
if ($formatBaseline -ne '泥浆外运(运距10KM),Φ560X33.2MM|' -or
    $formatFullWidth -ne $formatBaseline -or $formatLowerPhi -ne $formatBaseline) {
    throw "插件未统一空白、全半角标点、Ф/Φ/φ、x/×：$formatBaseline / $formatFullWidth / $formatLowerPhi"
}
if ([string]$buildQuantitySignature.Invoke($null, @('泥浆外运(运距5km),Φ710X33.2mm', 'm3')) -eq $formatBaseline) {
    throw '插件归一化不得删除距离和规格数字等业务参数。'
}
$legacyAliases = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([StringComparer]::OrdinalIgnoreCase)
$legacyAliases['HRB400钢筋KG|'] = 'HRB400钢筋|'
$resolveSqlSignature = $type.GetMethod('ResolveSmartSqlSignature', $flags)
$resolveArgs = [object[]]::new(2); $resolveArgs[0] = 'HRB400钢筋KG|'; $resolveArgs[1] = $legacyAliases.PSObject.BaseObject
if ($resolveSqlSignature.Invoke($null, $resolveArgs.PSObject.BaseObject) -ne 'HRB400钢筋|') {
    throw 'SQL 存量嵌入单位签名没有通过 QuantityAlias 桥接到名称级签名。'
}
$buildQty = $type.GetMethod('BuildNameDrivenQtyText', $flags)
if ($buildQty.Invoke($null, @('2700.2', 'kg', 't')) -ne '2700.2/1000') { throw 'kg 到 t 的当前单位换算错误。' }
if ($buildQty.Invoke($null, @('2', 't', 't')) -ne '2') { throw 't 到 t 不应重复除以1000。' }
if ($buildQty.Invoke($null, @('100', 'm2', 'm3')) -ne '100') { throw 'm2 到 m3 不应静默使用定额单位前缀换算。' }
if ($buildQty.Invoke($null, @('100', '', '10m3')) -ne '100') { throw '缺少当前工程量单位时不应静默换算。' }
if ($buildQty.Invoke($null, @('100', '天然密实方', '压实方')) -ne '100') { throw '不同方态不应静默按 1:1 换算。' }
Write-Host 'PASS 名称级签名兼容旧单位签名且当前单位换算不重复'

Write-Host 'Test-BindingLearningImmediate: PASS'
