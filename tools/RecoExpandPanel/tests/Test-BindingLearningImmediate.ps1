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
Assert-Contains $excelLink 'string entryScope = String.IsNullOrWhiteSpace(link.EntryCode) ? link.ChapterSeq : link.EntryCode;' '绑定组件分组没有使用稳定条目边界。'
Assert-Contains $excelLink 'existing["entry_code"] = group.EntryCode ?? "";' '本机待同步组件没有保存条目上下文。'
Assert-Contains $excelLink 'group.Targets.Add' '表达式学习组没有保留完整组件框目标。'
Assert-Contains $excelLink 'public string QuotaUnit { get; set; }' 'ExcelQuotaLink 没有持久化定额单位。'
Assert-Contains $excelLink 'public string EntryName { get; set; }' 'ExcelQuotaLink 没有持久化条目名称。'
Assert-Contains $excelLink 'PopulateExcelQuotaLinkLearningContext(conn, link);' '手动/批量绑定没有补齐编制办法和条目上下文。'
Assert-Contains $excelLink 'link.QuotaUnit = GetRowValue(row, "单位", "定额单位", "计量单位");' '手动/批量/快速绑定没有传递定额单位。'
Assert-Contains $excelLink 'item.Link.QuotaUnit = item.QuotaUnit;' '自动匹配预览保存时没有把定额单位写入 ExcelQuotaLink。'
Assert-Contains $autoMatch 'link.QuotaUnit = quotaUnit;' '自动匹配绑定没有传递定额单位。'
Assert-Contains $autoMatch 'link.EntryName = ReadAutoMatchReaderText(reader, 9);' '自动匹配绑定没有传递条目名称。'
Assert-Contains $autoMatch 'link.Method = projectMethod;' '自动匹配绑定没有传递编制办法。'
Assert-Contains $smartFill 'MergePendingLocalMappingsIntoSmartSnapshot' 'SQL 快照没有合并本机尚未汇总的新学习关系。'
Assert-Contains $smartFill 'LoadCurrentSmartQuotaMetadata' '推荐预览没有从当前运行版本读取定额元数据。'
if ($smartFill.Contains('BuildNameDrivenQtyText(row.QuantityText, row.Unit, target.Unit)')) { throw '推荐数量仍在使用 SQL 历史 target_unit 换算。' }
Assert-Contains $templatePanel 'FeedbackNameMatches(groupLeader.TemplateName, replacements' '模板铺量右键绑定后没有立即写入学习关系。'

$dll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL)) { $env:RECO_EXPAND_DLL } else { Join-Path $repoRoot 'RecoQuotaRecommend\bin\RecoExpandPanel.dll' }
if (-not (Test-Path -LiteralPath $dll)) { throw "找不到 $dll，先构建" }
$dllDir = Split-Path -Parent $dll
foreach ($dependency in @('NPOI.dll', 'NPOI.OpenXmlFormats.dll', 'NPOI.OpenXml4Net.dll', 'NPOI.OOXML.dll', 'ICSharpCode.SharpZipLib.dll')) {
    $dependencyPath = Join-Path $dllDir $dependency
    if (Test-Path -LiteralPath $dependencyPath) { [void][System.Reflection.Assembly]::LoadFrom($dependencyPath) }
}

$type = [System.Reflection.Assembly]::LoadFrom($dll).GetType('RecoNet.FormPanel', $true)
$flags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$extract = $type.GetMethod('ExtractPositiveAdditiveCellAddresses', $flags)
if ($null -eq $extract) { throw '编译结果缺少 ExtractPositiveAdditiveCellAddresses。' }

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

    # 同一个来源表达式若属于不同稳定条目，必须拆成不同组件框，不能因地址相同而串组。
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
    if ($groups.Count -ne 6) { throw "表达式相同但条目不同也应拆组，预期六个正向来源别名，实际 $($groups.Count) 套。" }
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
            ([string]$targetType.GetField('Code', $flags).GetValue($_)) + ':' + ([string]$targetType.GetField('Unit', $flags).GetValue($_))
        } | Sort-Object)
        $actual += "$sourceCell|$name|$unit|$rowNo|$method|$entryCode|$entryName|$($targetFacts -join ',')"
    }
    $actual = @($actual | Sort-Object)
    $expected = @(
        'F1|HRB400钢筋|kg|1|2024|0101-01|沉井工程|QY-317:t,QY-318:10m3',
        'F2|HPB300钢筋|kg|2|2024|0101-01|沉井工程|QY-317:t,QY-318:10m3',
        'F1|HRB400钢筋|kg|1|2024|0301-01|独立条目|QY-999:t',
        'F2|HPB300钢筋|kg|2|2024|0301-01|独立条目|QY-999:t',
        'F3|HRB400钢筋|kg|3|2024|0201-01|另一处钢筋|QY-317:t',
        'F4|HPB300钢筋|kg|4|2024|0201-01|另一处钢筋|QY-317:t'
    ) | Sort-Object
    if (($actual -join ';') -ne ($expected -join ';')) {
        throw "独立别名/单位/行号不正确：'$($actual -join ';')'"
    }
    Write-Host 'PASS 工程量/定额单位分开、组件共用条目且重复 HRB/HPB 不跨表达式合并'
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
$buildQuantitySignature = $type.GetMethod('BuildSmartQuantitySignature', $flags)
if ($buildQuantitySignature.Invoke($null, @('HRB400钢筋 kg', 'kg')) -ne 'HRB400钢筋|') {
    throw '旧名称尾部嵌入单位没有借助 QuantityAlias 归并为名称级签名。'
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
