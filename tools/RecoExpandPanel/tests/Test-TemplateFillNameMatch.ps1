$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$dll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL)) { $env:RECO_EXPAND_DLL } else { "D:\AI文件\自动预算\RecoQuotaRecommend\bin\RecoExpandPanel.dll" }
if (-not (Test-Path $dll)) { throw "找不到 $dll，先构建" }
$type = [System.Reflection.Assembly]::LoadFrom($dll).GetType('RecoNet.FormPanel')
$dllDir = Split-Path -Parent $dll
foreach ($dependency in @('NPOI.OpenXmlFormats.dll', 'NPOI.OpenXml4Net.dll', 'NPOI.OOXML.dll')) {
    [void][System.Reflection.Assembly]::LoadFrom((Join-Path $dllDir $dependency))
}
$flags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$norm = $type.GetMethod('NormalizeMatchText', $flags)
$score = $type.GetMethod('MatchNameScore', $flags, $null, [Type[]]@([string], [string]), $null)

function N($s) { return $norm.Invoke($null, @($s)) }
function S($a, $b) { return $score.Invoke($null, @((N $a), (N $b))) }

if ((N "铺设 无缝线路（km）") -ne "铺设无缝线路km") { throw "归一化失败: $(N '铺设 无缝线路（km）')" }
if ((N "Ф560×33.2mm") -ne (N "φ560ｘ33.2mm")) { throw "Ф/Φ/φ、x/× 归一化失败" }
Write-Host "PASS 归一化"
if ((S "铺设无缝线路" "铺设无缝线路") -ne 100) { throw "同名应100" }
Write-Host "PASS 同名满分"
if ((S "铺设无缝线路" "无缝线路铺设") -lt 55) { throw "改写措辞应>=55, 实际 $(S '铺设无缝线路' '无缝线路铺设')" }
Write-Host "PASS 改写措辞命中"
if ((S "500m长轨铺设" "25m长轨铺设") -ge 55) { throw "数字不符应<55, 实际 $(S '500m长轨铺设' '25m长轨铺设')" }
Write-Host "PASS 数字不符不误配"
if ((S "土方开挖" "钢筋制作安装") -ge 40) { throw "无关应低分" }
Write-Host "PASS 无关低分"

$chapter = $type.GetMethod('AreMatchChaptersCompatible', $flags)
if (-not $chapter.Invoke($null, @('第一章 路基工程', '路基工程'))) { throw '同章节标题应兼容' }
if ($chapter.Invoke($null, @('第一章 路基工程', '第二章 站场工程'))) { throw '跨章节不应兼容' }
if ($chapter.Invoke($null, @('第一章 路基工程', '第二章 路基工程'))) { throw '章号冲突时标题再相似也不得兼容' }
if ($chapter.Invoke($null, @('', '路基工程'))) { throw '缺章节不应自动兼容' }
Write-Host "PASS 章节兼容守卫"

$templateType = $type.GetNestedType('FillTemplate', $flags)
$templateRowType = $type.GetNestedType('FillTemplateRow', $flags)

function Test-WorkbookReadPerformancePaths {
    $fixturePath = Join-Path ([IO.Path]::GetTempPath()) ("reco-template-fill-" + [Guid]::NewGuid().ToString('N') + '.xlsx')
    $xlsFixturePath = Join-Path ([IO.Path]::GetTempPath()) ("reco-template-fill-" + [Guid]::NewGuid().ToString('N') + '.xls')
    $fixtureBook = $null
    $xlsFixtureBook = $null
    try {
        $fixtureBook = New-Object NPOI.XSSF.UserModel.XSSFWorkbook
        $fixtureSheet = $fixtureBook.CreateSheet('测试表')
        $fixtureRow = $fixtureSheet.CreateRow(0)
        $fixtureRow.CreateCell(0).SetCellValue('工程量名称')
        $fixtureRow.CreateCell(1).SetCellValue([double]123.5)
        $fixtureRow.CreateCell(2).SetCellValue($true)
        [void]$fixtureRow.CreateCell(3)
        $fixtureStream = [IO.File]::Create($fixturePath)
        try { $fixtureBook.Write($fixtureStream) }
        finally { $fixtureStream.Dispose() }

        $addresses = [string[]]@('A1', 'B1', 'C1', 'D1', 'E1')
        $direct = $type.GetMethod('TryReadXlsxTargetCells', $flags)
        $npoi = $type.GetMethod('TryReadSheetTargetCellsByNpoi', $flags)
        $directArgs = [object[]]::new(5)
        $npoiArgs = [object[]]::new(5)
        $directArgs[0] = $fixturePath.PSObject.BaseObject
        $directArgs[1] = ([string]'测试表').PSObject.BaseObject
        $directArgs[2] = $addresses.PSObject.BaseObject
        $npoiArgs[0] = $fixturePath.PSObject.BaseObject
        $npoiArgs[1] = ([string]'测试表').PSObject.BaseObject
        $npoiArgs[2] = $addresses.PSObject.BaseObject
        if (-not $direct.Invoke($null, $directArgs.PSObject.BaseObject)) { throw "定点读取失败: $($directArgs[4])" }
        if (-not $npoi.Invoke($null, $npoiArgs.PSObject.BaseObject)) { throw "NPOI读取失败: $($npoiArgs[4])" }
        foreach ($address in $addresses) {
            if ($directArgs[3][$address] -ne $npoiArgs[3][$address]) {
                throw "两套读取器结果不一致 ${address}: '$($directArgs[3][$address])' / '$($npoiArgs[3][$address])'"
            }
        }
        Write-Host 'PASS xlsx 定点读取与 NPOI 结果一致'

        $linkType = $type.GetNestedType('ExcelQuotaLink', $flags)
        $linkListType = [Collections.Generic.List``1].MakeGenericType($linkType)
        $links = [Activator]::CreateInstance($linkListType)
        foreach ($definition in @(@('A1', 'A1+B1'), @('C1', 'C1'))) {
            $link = [Activator]::CreateInstance($linkType)
            $linkType.GetProperty('ExcelPath', $flags).SetValue($link, $fixturePath, $null)
            $linkType.GetProperty('WorksheetName', $flags).SetValue($link, '测试表', $null)
            $linkType.GetProperty('CellAddress', $flags).SetValue($link, $definition[0], $null)
            $linkType.GetProperty('Expression', $flags).SetValue($link, $definition[1], $null)
            $links.Add($link)
        }
        $contextType = $type.GetNestedType('ExcelSyncReadContext', $flags)
        $contextCtor = @($contextType.GetConstructors($flags))[0]
        $contextCtorArgs = [object[]]::new(1)
        $contextCtorArgs[0] = $links.PSObject.BaseObject
        $context = $contextCtor.Invoke($contextCtorArgs)
        $readCell = $contextType.GetMethod('TryReadWorkbookCellValue', $flags)
        foreach ($address in @('A1', 'B1', 'C1')) {
            $readArgs = [object[]]::new(5)
            $readArgs[0] = $fixturePath.PSObject.BaseObject
            $readArgs[1] = ([string]'测试表').PSObject.BaseObject
            $readArgs[2] = ([string]$address).PSObject.BaseObject
            if (-not $readCell.Invoke($context, $readArgs.PSObject.BaseObject)) { throw "上下文读取失败 ${address}: $($readArgs[4])" }
        }
        if ($contextType.GetField('FileSheetReadCount', $flags).GetValue($context) -ne 1) {
            throw '同一工作表应只加载一次'
        }
        $directCountField = $contextType.GetField('DirectXlsxReadCount', $flags)
        $npoiCountField = $contextType.GetField('NpoiReadCount', $flags)
        if ($null -eq $directCountField -or $null -eq $npoiCountField) {
            throw '缺少读取路径计数字段'
        }
        if ($directCountField.GetValue($context) -ne 1 -or $npoiCountField.GetValue($context) -ne 0) {
            throw 'xlsx 应优先且只使用一次定点读取'
        }
        Write-Host 'PASS xlsx 上下文批量读取只走一次快速路径'

        $xlsFixtureBook = New-Object NPOI.HSSF.UserModel.HSSFWorkbook
        $xlsSheet = $xlsFixtureBook.CreateSheet('测试表')
        $xlsRow = $xlsSheet.CreateRow(0)
        $xlsRow.CreateCell(0).SetCellValue('旧版工程量')
        $xlsStream = [IO.File]::Create($xlsFixturePath)
        try { $xlsFixtureBook.Write($xlsStream) }
        finally { $xlsStream.Dispose() }

        $xlsLinks = [Activator]::CreateInstance($linkListType)
        $xlsLink = [Activator]::CreateInstance($linkType)
        $linkType.GetProperty('ExcelPath', $flags).SetValue($xlsLink, $xlsFixturePath, $null)
        $linkType.GetProperty('WorksheetName', $flags).SetValue($xlsLink, '测试表', $null)
        $linkType.GetProperty('CellAddress', $flags).SetValue($xlsLink, 'A1', $null)
        $linkType.GetProperty('Expression', $flags).SetValue($xlsLink, 'A1', $null)
        $xlsLinks.Add($xlsLink)
        $xlsContextArgs = [object[]]::new(1)
        $xlsContextArgs[0] = $xlsLinks.PSObject.BaseObject
        $xlsContext = $contextCtor.Invoke($xlsContextArgs)
        $xlsReadArgs = [object[]]::new(5)
        $xlsReadArgs[0] = $xlsFixturePath.PSObject.BaseObject
        $xlsReadArgs[1] = ([string]'测试表').PSObject.BaseObject
        $xlsReadArgs[2] = ([string]'A1').PSObject.BaseObject
        if (-not $readCell.Invoke($xlsContext, $xlsReadArgs.PSObject.BaseObject) -or $xlsReadArgs[3] -ne '旧版工程量') {
            throw "xls NPOI 读取失败: '$($xlsReadArgs[3])' / '$($xlsReadArgs[4])'"
        }
        if ($directCountField.GetValue($xlsContext) -ne 0 -or $npoiCountField.GetValue($xlsContext) -ne 1) {
            throw 'xls 应保持只使用一次 NPOI 读取'
        }
        Write-Host 'PASS xls 保持 NPOI 读取路径'

        $templateSheetNames = $type.GetMethod('GetTemplateFillSheetNames', $flags)
        if ($null -eq $templateSheetNames) { throw '缺少模板铺量快速 sheet 名称读取入口' }
        foreach ($fixture in @(
            [pscustomobject]@{ Path = $fixturePath; Expected = '测试表'; Kind = 'xlsx' },
            [pscustomobject]@{ Path = $xlsFixturePath; Expected = '测试表'; Kind = 'xls' }
        )) {
            $sheetArgs = [object[]]::new(2)
            $sheetArgs[0] = ([string]$fixture.Path).PSObject.BaseObject
            $sheetNames = @($templateSheetNames.Invoke($null, $sheetArgs.PSObject.BaseObject))
            if (($sheetNames -join '|') -ne $fixture.Expected) {
                throw "模板铺量 $($fixture.Kind) sheet 名称读取错误: '$($sheetNames -join '|')' / '$($sheetArgs[1])'"
            }
        }
        Write-Host 'PASS 模板铺量 xlsx 快速读取与 xls NPOI sheet 名称兼容'

        $batchTemplate = [Activator]::CreateInstance($templateType)
        $templateType.GetField('WorkbookPath', $flags).SetValue($batchTemplate, $fixturePath)
        foreach ($expr in @('A1+B1', 'C1')) {
            $row = [Activator]::CreateInstance($templateRowType)
            $templateRowType.GetField('SourceSheet', $flags).SetValue($row, '测试表')
            $templateRowType.GetField('SourceExpr', $flags).SetValue($row, $expr)
            $templateType.GetField('Rows', $flags).GetValue($batchTemplate).Add($row)
        }
        $hashSetType = [Collections.Generic.HashSet``1].MakeGenericType([int])
        $hiddenType = [Collections.Generic.Dictionary``2].MakeGenericType([string], $hashSetType)
        $mergedRegionType = $type.GetNestedType('ExcelMergedRegion', $flags)
        $mergedListType = [Collections.Generic.List``1].MakeGenericType($mergedRegionType)
        $mergedType = [Collections.Generic.Dictionary``2].MakeGenericType([string], $mergedListType)
        $hiddenCache = [Activator]::CreateInstance($hiddenType)
        $mergedCache = [Activator]::CreateInstance($mergedType)
        $createBatchContext = $type.GetMethod('CreateNameFillReadContext', $flags)
        if ($null -eq $createBatchContext) { throw '缺少按名称模板批量读取上下文入口' }
        $batchArgs = [object[]]::new(3)
        $batchArgs[0] = $batchTemplate
        $batchArgs[1] = $hiddenCache.PSObject.BaseObject
        $batchArgs[2] = $mergedCache.PSObject.BaseObject
        $batchContext = $createBatchContext.Invoke($null, $batchArgs)
        if ($contextType.GetProperty('FileCount', $flags).GetValue($batchContext, $null) -ne 1 -or
            $contextType.GetProperty('SheetCount', $flags).GetValue($batchContext, $null) -ne 1 -or
            $contextType.GetProperty('CellCount', $flags).GetValue($batchContext, $null) -lt 2) {
            throw '名字模板应把同一工作簿和工作表的名称地址聚合到一个上下文'
        }
        Write-Host 'PASS 名字模板批量聚合一个读取上下文'
    }
    finally {
        if ($null -ne $fixtureBook) {
            try { $fixtureBook.Close() } catch { }
        }
        if ($null -ne $xlsFixtureBook) {
            try { $xlsFixtureBook.Close() } catch { }
        }
        if (Test-Path -LiteralPath $fixturePath) { Remove-Item -LiteralPath $fixturePath -Force }
        if (Test-Path -LiteralPath $xlsFixturePath) { Remove-Item -LiteralPath $xlsFixturePath -Force }
    }
}

Test-WorkbookReadPerformancePaths

$groupBuilder = $type.GetMethod('BuildTemplateNameGroups', $flags)
$conflict = $type.GetMethod('GetExactNameConflict', $flags)
$uniqueBest = $type.GetMethod('FindUniqueBestMatchIndex', $flags)
$resolutionMode = $type.GetMethod('GetExactNameResolutionMode', $flags)
if ($null -eq $groupBuilder -or $null -eq $conflict -or $null -eq $uniqueBest) {
    throw '缺少名字组件分组或歧义判定入口'
}
if ($null -eq $resolutionMode) { throw '缺少精确同名处理模式入口' }
if ($resolutionMode.Invoke($null, @(1, 1)) -ne 'single') { throw '唯一目标和唯一绑定应直接命中' }
if ($resolutionMode.Invoke($null, @(2, 1)) -ne 'reuse') { throw '重复目标和唯一绑定应整组复用' }
if ($resolutionMode.Invoke($null, @(1, 2)) -ne 'choice') { throw '多个同名绑定应下拉选择' }
if ($resolutionMode.Invoke($null, @(2, 2)) -ne 'choice') { throw '重复目标和多个绑定应逐行下拉选择' }

$groupTemplate = [Activator]::CreateInstance($templateType)
$templateType.GetField('WorkbookPath', $flags).SetValue($groupTemplate, 'C:\fixture.xlsx')
function Add-TemplateRow([string]$name, [string]$expr, [string]$chapter, [string]$code) {
    $row = [Activator]::CreateInstance($templateRowType)
    $templateRowType.GetField('MatchName', $flags).SetValue($row, $name)
    $templateRowType.GetField('SourceSheet', $flags).SetValue($row, '测试')
    $templateRowType.GetField('SourceExpr', $flags).SetValue($row, $expr)
    $templateRowType.GetField('MatchChapter', $flags).SetValue($row, $chapter)
    $templateRowType.GetField('QuotaCode', $flags).SetValue($row, $code)
    $templateType.GetField('Rows', $flags).GetValue($groupTemplate).Add($row)
}
Add-TemplateRow '低烟无卤电缆 m' 'D2/100' '' 'DY-519'
Add-TemplateRow '低烟无卤电缆 m' 'D2' '第二章' 'ZLF*1.01'
$groups = $groupBuilder.Invoke($null, @($groupTemplate))
if ($groups.Count -ne 1 -or $groups[0].Indexes.Count -ne 2) { throw '同锚点多定额应组成一个组件组' }
$candidateLabel = $type.GetMethod('BuildTemplateCandidateLabel', $flags)
if ($null -eq $candidateLabel) { throw '缺少绑定组候选标签入口' }
$label = [string]$candidateLabel.Invoke($null, @($groupTemplate, $groups[0]))
if ($label -notmatch 'DY-519' -or $label -notmatch 'ZLF\*1\.01' -or $label -notmatch '组件2条') {
    throw "组件候选标签不得拆组: $label"
}
Add-TemplateRow '低烟无卤电缆 m' 'D30' '' 'DY-520'
$groups = $groupBuilder.Invoke($null, @($groupTemplate))
if ($groups.Count -ne 2) { throw '同名不同来源锚点应形成两个组件组' }
if ($conflict.Invoke($null, @(1, 2)) -ne 'template') { throw '模板同名多来源应判为模板歧义' }
if ($conflict.Invoke($null, @(2, 1)) -ne 'target') { throw '目标多行同名应判为目标歧义' }
if ($conflict.Invoke($null, @(1, 1)) -ne '') { throw '唯一同名不应判为歧义' }

$bestArgs = [object[]]@((N '低烟无卤电缆'), '第七章',
    [string[]]@((N '低烟无卤电缆A'), (N '低烟无卤电缆B')),
    [string[]]@('', ''), $false)
$bestIndex = $uniqueBest.Invoke($null, $bestArgs)
if ($bestIndex -ne -1 -or -not [bool]$bestArgs[4]) { throw '模糊最高分并列应判为黄色歧义' }

$bestArgs = [object[]]@((N '低烟无卤电缆'), '第七章',
    [string[]]@((N '低烟无卤电缆A'), (N '完全无关')),
    [string[]]@('第八章', '第七章'), $false)
$bestIndex = $uniqueBest.Invoke($null, $bestArgs)
if ($bestIndex -ne 0 -or [bool]$bestArgs[4]) { throw '唯一最佳名称不得被章节不同淘汰' }
Write-Host 'PASS 名字组件分组与同名歧义判定'

$targetRank = $type.GetMethod('TemplateTargetRank', $flags)
if ($null -eq $targetRank) { throw '缺少模板目标类别排序入口' }
if ($targetRank.Invoke($null, @('FY-842')) -ne 0) { throw '普通定额应排第一类' }
if ($targetRank.Invoke($null, @('109001003*1.02')) -ne 1) { throw '带系数数字材料应排第二类' }
if ($targetRank.Invoke($null, @('SH')) -ne 2) { throw '辅助代码应排第三类' }

$orderTemplate = [Activator]::CreateInstance($templateType)
function Add-OrderTestRow([string]$code) {
    $row = [Activator]::CreateInstance($templateRowType)
    $templateRowType.GetField('MatchName', $flags).SetValue($row, '圈梁 10m3')
    $templateRowType.GetField('MatchChapter', $flags).SetValue($row, '第八章 房屋')
    $templateRowType.GetField('QuotaCode', $flags).SetValue($row, $code)
    $templateRowType.GetField('OrderInItem', $flags).SetValue($row, 0)
    $templateType.GetField('Rows', $flags).GetValue($orderTemplate).Add($row)
}
Add-OrderTestRow '109001003*1.02'
Add-OrderTestRow 'FY-842'
Add-OrderTestRow 'FY-841'
Add-OrderTestRow 'SH'
$orderGroups = $groupBuilder.Invoke($null, @($orderTemplate))
$orderedGroupIndexes = $type.GetMethod('OrderedTemplateGroupIndexes', $flags)
if ($null -eq $orderedGroupIndexes -or $orderGroups.Count -ne 1) {
    throw '缺少模板组件稳定排序入口'
}
$orderedIndexes = $orderedGroupIndexes.Invoke($null, [object[]]@($orderTemplate, $orderGroups[0]))
$orderedCodes = @($orderedIndexes | ForEach-Object {
    $templateType.GetField('Rows', $flags).GetValue($orderTemplate)[$_].QuotaCode
})
$expectedOrder = @('FY-842', 'FY-841', '109001003*1.02', 'SH')
if (($orderedCodes -join '|') -ne ($expectedOrder -join '|')) {
    throw "模板组件顺序错误: $($orderedCodes -join ' -> ')"
}
Write-Host 'PASS 模板组件按定额、材料、辅助代码稳定排序'

$confirmed = $type.GetMethod('IsInsertGroupFullyConfirmed', $flags)
if (-not $confirmed.Invoke($null, @(3, 3))) { throw '整组写入完成应确认成功' }
if ($confirmed.Invoke($null, @(3, 2))) { throw '部分写入不得确认整组成功' }
Write-Host "PASS 写入成功反馈守卫"

$itemType = $type.GetNestedType('FillPreviewItem', $flags)
$itemListType = [System.Collections.Generic.List``1].MakeGenericType($itemType)
$allItems = [Activator]::CreateInstance($itemListType)
$replacements = [Activator]::CreateInstance($itemListType)
function New-PreviewItem([int]$row, [int]$order, [string]$name) {
    $item = [Activator]::CreateInstance($itemType)
    $itemType.GetField('IsNameDriven', $flags).SetValue($item, $true)
    $itemType.GetField('TargetRow', $flags).SetValue($item, $row)
    $itemType.GetField('GroupOrder', $flags).SetValue($item, $order)
    $itemType.GetField('TargetName', $flags).SetValue($item, $name)
    return $item
}

$candidateType = $type.GetNestedType('NameQuotaCandidateGroup', $flags)
$confirmExact = $type.GetMethod('ConfirmSingleExactNameGroup', $flags)
$confirmCurrent = $type.GetMethod('ConfirmCurrentExactNameGroup', $flags)
$applyCandidate = $type.GetMethod('ApplyExactNameCandidate', $flags)
if ($null -eq $candidateType -or $null -eq $confirmExact -or $null -eq $confirmCurrent -or $null -eq $applyCandidate) {
    throw '缺少重复名称候选或确认入口'
}

$singleItems = [Activator]::CreateInstance($itemListType)
$singleItems.Add((New-PreviewItem 40 0 '钢管 SC20 m'))
$singleItems.Add((New-PreviewItem 40 1 ''))
foreach ($member in $singleItems) {
    $itemType.GetField('NeedExactNameConfirmation', $flags).SetValue($member, $true)
    $itemType.GetField('Selected', $flags).SetValue($member, $false)
    $itemType.GetField('AlignNote', $flags).SetValue($member, '名称学习命中，候选待确认')
}
if (-not $confirmExact.Invoke($null, [object[]]@($singleItems.PSObject.BaseObject, 40))) {
    throw '唯一绑定整组确认应成功'
}
if (@($singleItems | Where-Object { -not $_.Selected -or $_.NeedExactNameConfirmation }).Count -ne 0) {
    throw '唯一绑定确认后应整组勾选并取消红色确认状态'
}
if ($singleItems[0].AlignNote -ne '人工确认重复名称' -or
    $singleItems[1].AlignNote -ne '组件框第 2 条（人工确认重复名称）') {
    throw '唯一绑定确认后应刷新整个组件组状态'
}

$choiceItems = [Activator]::CreateInstance($itemListType)
$choiceLeader = New-PreviewItem 50 0 '重复工程量'
$itemType.GetField('NeedExactNameConfirmation', $flags).SetValue($choiceLeader, $true)
$itemType.GetField('Selected', $flags).SetValue($choiceLeader, $false)
$choiceItems.Add($choiceLeader)
$candidate = [Activator]::CreateInstance($candidateType)
$candidateType.GetField('Key', $flags).SetValue($candidate, 'group-b')
$candidateType.GetField('Label', $flags).SetValue($candidate, 'DY-519 + ZLF*1.01（组件2条）')
$candidateMembers = $candidateType.GetField('Items', $flags).GetValue($candidate)
$candidateMembers.Add((New-PreviewItem 50 0 '重复工程量'))
$candidateMembers.Add((New-PreviewItem 50 1 ''))
$itemType.GetField('QuotaCode', $flags).SetValue($candidateMembers[0], 'DY-519')
$itemType.GetField('QuotaCode', $flags).SetValue($candidateMembers[1], 'ZLF*1.01')
$candidateList = [Activator]::CreateInstance($itemType.GetField('NameQuotaCandidates', $flags).FieldType)
$candidateList.Add($candidate)
$itemType.GetField('NameQuotaCandidates', $flags).SetValue($choiceLeader, $candidateList)
if (-not $applyCandidate.Invoke($null, [object[]]@($choiceItems.PSObject.BaseObject, 50, 'group-b'))) {
    throw '候选组件整组切换应成功'
}
if ($choiceItems.Count -ne 2 -or $choiceItems[0].QuotaCode -ne 'DY-519' -or $choiceItems[1].QuotaCode -ne 'ZLF*1.01') {
    throw '候选切换必须保留完整组件组'
}
if (@($choiceItems | Where-Object { -not $_.Selected -or $_.NeedExactNameConfirmation }).Count -ne 0) {
    throw '候选选择后应整组勾选并取消红色确认状态'
}

$defaultItems = [Activator]::CreateInstance($itemListType)
$defaultLeader = New-PreviewItem 55 0 '默认候选工程量'
$itemType.GetField('QuotaCode', $flags).SetValue($defaultLeader, 'DY-959')
$itemType.GetField('NeedExactNameConfirmation', $flags).SetValue($defaultLeader, $true)
$itemType.GetField('Selected', $flags).SetValue($defaultLeader, $false)
$defaultCandidates = [Activator]::CreateInstance($itemType.GetField('NameQuotaCandidates', $flags).FieldType)
foreach ($definition in @(@('group-a', 'DY-959'), @('group-b', 'DY-942'))) {
    $option = [Activator]::CreateInstance($candidateType)
    $candidateType.GetField('Key', $flags).SetValue($option, $definition[0])
    $candidateType.GetField('Label', $flags).SetValue($option, $definition[1])
    $optionItem = New-PreviewItem 55 0 '默认候选工程量'
    $itemType.GetField('QuotaCode', $flags).SetValue($optionItem, $definition[1])
    $candidateType.GetField('Items', $flags).GetValue($option).Add($optionItem)
    $defaultCandidates.Add($option)
}
$itemType.GetField('NameQuotaCandidates', $flags).SetValue($defaultLeader, $defaultCandidates)
$itemType.GetField('SelectedNameQuotaCandidateKey', $flags).SetValue($defaultLeader, 'group-a')
$itemType.GetField('QuantityText', $flags).SetValue($defaultLeader, '123.45/10')
$defaultItems.Add($defaultLeader)
if (-not $confirmCurrent.Invoke($null, [object[]]@($defaultItems.PSObject.BaseObject, 55))) {
    throw '当前默认候选应允许直接确认'
}
if ($defaultItems.Count -ne 1 -or $defaultItems[0].QuotaCode -ne 'DY-959' -or
    -not $defaultItems[0].Selected -or $defaultItems[0].NeedExactNameConfirmation -or
    $defaultItems[0].QuantityText -ne '123.45/10') {
    throw '直接确认必须接受当前显示候选、保留人工数量并取消确认状态'
}
Write-Host 'PASS 重复名称唯一确认、默认候选确认与候选整组切换'

$panelType = $type.GetNestedType('TemplateFillPanel', $flags)
$panelCtor = @($panelType.GetConstructors($flags) | Where-Object { $_.GetParameters().Count -eq 1 })[0]
$mainForm = New-Object System.Windows.Forms.Form
$panel = $null
try {
    $panel = $panelCtor.Invoke([object[]]@($mainForm.PSObject.BaseObject))
    $reloadCountField = $panelType.GetField('targetWorkbookReloadCount', $flags)
    if ($null -eq $reloadCountField -or $reloadCountField.GetValue($panel) -ne 1) {
        throw '模板铺量窗口首次构造应只刷新一次目标工作簿'
    }
    Write-Host 'PASS 模板铺量窗口首次只刷新一次目标工作簿'
    $templateBox = $panelType.GetField('cmbTemplate', $flags).GetValue($panel)
    $reloadCountBeforeSwitch = $reloadCountField.GetValue($panel)
    $templateBox.Items.Clear()
    [void]$templateBox.Items.Add('临时模板A')
    [void]$templateBox.Items.Add('临时模板B')
    $templateBox.SelectedIndex = 0
    $templateBox.SelectedIndex = 1
    if ($reloadCountField.GetValue($panel) -ne $reloadCountBeforeSwitch) {
        throw '切换模板不得重新枚举目标工作簿'
    }
    Write-Host 'PASS 切换模板不重新枚举目标工作簿'
    if ($null -eq $panelType.GetField('cmbTargetWorkbook', $flags)) {
        throw '模板铺量面板缺少目标 Excel 下拉框'
    }
    $workbookInfoType = $type.GetNestedType('OpenSpreadsheetWorkbookInfo', $flags)
    if ($null -eq $workbookInfoType) { throw '缺少打开工作簿描述类型' }
    $displayNameBuilder = $type.GetMethod('BuildOpenWorkbookDisplayName', $flags)
    $normalDisplay = $displayNameBuilder.Invoke($null, @('C:\目标目录\目标文件.xlsx', $false))
    $sourceDisplay = $displayNameBuilder.Invoke($null, @('C:\目标目录\目标文件.xlsx', $true))
    if ($normalDisplay -ne '目标文件.xlsx' -or $sourceDisplay -ne '目标文件.xlsx') {
        throw "目标 Excel 列表应只显示文件名: '$normalDisplay' / '$sourceDisplay'"
    }
    $workbookInfo = [Activator]::CreateInstance($workbookInfoType)
    $workbookInfoType.GetField('FullName', $flags).SetValue($workbookInfo, 'C:\目标目录\目标文件.xlsx')
    $workbookInfoType.GetField('DisplayName', $flags).SetValue($workbookInfo, '目标文件.xlsx')
    $workbookInfoType.GetField('ActiveSheetName', $flags).SetValue($workbookInfo, '目标表2')
    $workbookSheets = $workbookInfoType.GetField('SheetNames', $flags).GetValue($workbookInfo)
    $workbookSheets.Add('目标表1')
    $workbookSheets.Add('目标表2')
    $targetWorkbookBox = $panelType.GetField('cmbTargetWorkbook', $flags).GetValue($panel)
    $targetWorkbookBox.Items.Clear()
    [void]$targetWorkbookBox.Items.Add($workbookInfo)
    $targetWorkbookBox.SelectedItem = $workbookInfo
    $selectedPath = $panelType.GetMethod('GetSelectedTargetWorkbookPath', $flags).Invoke($panel, $null)
    $targetSheetBox = $panelType.GetField('cmbTargetSheet', $flags).GetValue($panel)
    if ($selectedPath -ne 'C:\目标目录\目标文件.xlsx' -or $targetSheetBox.Text -ne '目标表2') {
        throw "目标 Excel 与目标 sheet 联动失败: '$selectedPath' / '$($targetSheetBox.Text)'"
    }
    Write-Host 'PASS 目标 Excel 与目标 sheet 面板联动'

    $resolveUnit = $panelType.GetMethod('ResolveTemplateFillQuotaUnit', $flags)
    if ($null -eq $resolveUnit) { throw '缺少模板铺量临时绑定单位解析入口' }
    $unitGrid = New-Object System.Windows.Forms.DataGridView
    try {
        [void]$unitGrid.Columns.Add('计量单位', '计量单位')
        [void]$unitGrid.Rows.Add('hm')
        $resolvedUnit = $resolveUnit.Invoke($null,
            [object[]]@($null, $unitGrid.Rows[0].PSObject.BaseObject, [long]0))
        if ($resolvedUnit -ne 'hm') { throw "临时绑定没有读取计量单位: '$resolvedUnit'" }
    }
    finally { $unitGrid.Dispose() }

    $buildQty = $type.GetMethod('BuildNameDrivenQtyText', $flags)
    if ($buildQty.Invoke($null, @('1400', 'm', 'hm')) -ne '1400/100') {
        throw '右键临时绑定 m 到 hm 应生成 1400/100'
    }
    Write-Host 'PASS 临时绑定读取计量单位并按 m 到 hm 生成 1400/100'

    $panelPreview = $panelType.GetField('preview', $flags).GetValue($panel)
    $uiLeader = New-PreviewItem 60 0 '重复工程量'
    $itemType.GetField('QuotaCode', $flags).SetValue($uiLeader, 'DY-959')
    $itemType.GetField('NeedExactNameConfirmation', $flags).SetValue($uiLeader, $true)
    $itemType.GetField('Selected', $flags).SetValue($uiLeader, $false)

    $uiCandidates = [Activator]::CreateInstance($itemType.GetField('NameQuotaCandidates', $flags).FieldType)
    $uiFirst = [Activator]::CreateInstance($candidateType)
    $candidateType.GetField('Key', $flags).SetValue($uiFirst, 'group-a')
    $candidateType.GetField('Label', $flags).SetValue($uiFirst, '来源定额A + 来源材料A（组件2条）')
    $uiFirstItems = $candidateType.GetField('Items', $flags).GetValue($uiFirst)
    $uiFirstItems.Add((New-PreviewItem 60 0 '重复工程量'))
    $uiFirstItems.Add((New-PreviewItem 60 1 ''))
    $itemType.GetField('QuotaCode', $flags).SetValue($uiFirstItems[0], 'DY-959')
    $itemType.GetField('QuotaCode', $flags).SetValue($uiFirstItems[1], 'ZLF*1.01')
    $itemType.GetField('SourceName', $flags).SetValue($uiFirstItems[0], '来源定额A')
    $itemType.GetField('SourceName', $flags).SetValue($uiFirstItems[1], '来源材料A')
    $uiCandidates.Add($uiFirst)

    $uiSecond = [Activator]::CreateInstance($candidateType)
    $candidateType.GetField('Key', $flags).SetValue($uiSecond, 'group-b')
    $candidateType.GetField('Label', $flags).SetValue($uiSecond, '来源定额B + 来源材料B（组件2条）')
    $uiSecondItems = $candidateType.GetField('Items', $flags).GetValue($uiSecond)
    $uiSecondItems.Add((New-PreviewItem 60 0 '重复工程量'))
    $uiSecondItems.Add((New-PreviewItem 60 1 ''))
    $itemType.GetField('QuotaCode', $flags).SetValue($uiSecondItems[0], 'DY-519')
    $itemType.GetField('QuotaCode', $flags).SetValue($uiSecondItems[1], 'ZLF*1.01')
    $itemType.GetField('SourceName', $flags).SetValue($uiSecondItems[0], '来源定额B')
    $itemType.GetField('SourceName', $flags).SetValue($uiSecondItems[1], '来源材料B')
    $itemType.GetField('AlignNote', $flags).SetValue($uiSecondItems[0], '名称学习命中，候选待确认')
    $itemType.GetField('AlignNote', $flags).SetValue($uiSecondItems[1], '名称学习命中，候选待确认')
    $uiCandidates.Add($uiSecond)
    $itemType.GetField('NameQuotaCandidates', $flags).SetValue($uiLeader, $uiCandidates)
    $itemType.GetField('SelectedNameQuotaCandidateKey', $flags).SetValue($uiLeader, 'group-a')
    $panelPreview.Add($uiLeader)
    $uiMember = New-PreviewItem 60 1 ''
    $itemType.GetField('QuotaCode', $flags).SetValue($uiMember, 'ZLF*1.01')
    $itemType.GetField('SourceName', $flags).SetValue($uiLeader, '来源定额A')
    $itemType.GetField('SourceName', $flags).SetValue($uiMember, '来源材料A')
    $itemType.GetField('NeedExactNameConfirmation', $flags).SetValue($uiMember, $true)
    $itemType.GetField('Selected', $flags).SetValue($uiMember, $false)
    $panelPreview.Add($uiMember)

    $panelType.GetMethod('FillGrid', $flags).Invoke($panel, $null)
    $uiGrid = $panelType.GetField('grid', $flags).GetValue($panel)
    if ($uiGrid.Columns['qty'].ReadOnly) {
        throw '推荐定额数量列应允许直接编辑'
    }
    $uiGrid.Rows[0].Cells['qty'].Value = '123.45/10'
    $panelType.GetMethod('FlushGridSelectionsToPreview', $flags).Invoke($panel, $null)
    if ($uiLeader.QuantityText -ne '123.45/10') {
        throw '数量单元格修改后没有同步到写入预览对象'
    }
    $hasUnsafeCandidate = $panelType.GetMethod('HasUnsafeNameQuotaCandidate', $flags)
    $reviewItems = [Activator]::CreateInstance($itemListType)
    $reviewItem = New-PreviewItem 61 0 '待确认工程量'
    $itemType.GetField('AlignNote', $flags).SetValue(
        $reviewItem, '模板存在同名多来源，已带出候选，需下拉确认')
    $reviewItems.Add($reviewItem)
    $reviewArgs = New-Object object[] 1
    $reviewArgs[0] = $reviewItems.PSObject.BaseObject
    if ($hasUnsafeCandidate.Invoke($panel, $reviewArgs)) {
        throw '同名候选确认提示不应被当作组件风险'
    }
    $itemType.GetField('Status', $flags).SetValue($reviewItem, '取数失败：测试风险')
    if (-not $hasUnsafeCandidate.Invoke($panel, $reviewArgs)) {
        throw '真实取数失败状态仍应阻止直接确认'
    }
    if ($uiGrid.Rows.Count -ne 2 -or $uiGrid.Rows[0].Cells['sel'].ReadOnly -or
        -not $uiGrid.Rows[0].Cells['code'].ReadOnly -or $uiGrid.Rows[0].Cells['sname'].ReadOnly -or
        -not ($uiGrid.Rows[1].Cells['sel'] -is [System.Windows.Forms.DataGridViewTextBoxCell])) {
        throw '多候选应在源行定额列下拉，定额编号只读且组件成员不显示复选框'
    }
    if ($uiGrid.Rows[0].DefaultCellStyle.BackColor.ToArgb() -ne [System.Drawing.Color]::MistyRose.ToArgb()) {
        throw '多候选未确认行应标红'
    }
    $uiGrid.Rows[0].Cells['sel'].Value = $true
    $panelType.GetMethod('ApplyNameGroupSelectionFromCheck', $flags).Invoke(
        $panel, @($uiGrid.Rows[0].PSObject.BaseObject))
    $confirmedDefaultGroup = @($panelPreview | Where-Object { $_.TargetRow -eq 60 } | Sort-Object GroupOrder)
    if ($confirmedDefaultGroup[0].QuantityText -ne '123.45/10' -or
        $uiGrid.Rows[0].Cells['qty'].Value -ne '123.45/10' -or
        @($confirmedDefaultGroup | Where-Object { -not $_.Selected -or $_.NeedExactNameConfirmation }).Count -ne 0) {
        throw '默认候选修改数量后直接勾选，必须保留数量并原地确认整组'
    }
    $prepareDropDown = $panelType.GetMethod('PrepareNameQuotaDropDown', $flags)
    if (-not $prepareDropDown.Invoke($panel, @($uiGrid.Rows[0].PSObject.BaseObject))) { throw '源行定额下拉应创建成功' }
    if (-not ($uiGrid.Rows[0].Cells['sname'] -is [System.Windows.Forms.DataGridViewComboBoxCell]) -or
        ($uiGrid.Rows[0].Cells['code'] -is [System.Windows.Forms.DataGridViewComboBoxCell])) {
        throw '源行定额单元格应切换为下拉框，定额编号不得变为下拉框'
    }
    $panelType.GetMethod('ApplyNameQuotaOption', $flags).Invoke($panel,
        @($uiGrid.Rows[0].PSObject.BaseObject, '来源定额B + 来源材料B（组件2条）'))
    if ($uiGrid.Rows.Count -ne 2 -or $uiGrid.Rows[0].Cells['code'].Value -ne 'DY-519' -or
        $uiGrid.Rows[1].Cells['code'].Value -ne 'ZLF*1.01') {
        throw '界面选择组件候选后应展开完整定额组'
    }
    if ($uiGrid.Rows[0].Tag.AlignNote -ne '人工选择同名绑定' -or
        $uiGrid.Rows[1].Tag.AlignNote -ne '组件框第 2 条（人工选择同名绑定）' -or
        $uiGrid.Rows[0].Cells['st'].Value -ne '人工选择同名绑定' -or
        $uiGrid.Rows[1].Cells['st'].Value -ne '组件框第 2 条（人工选择同名绑定）') {
        throw '选择组件候选后应刷新整组状态，不得保留旧候选提示'
    }
    if (-not [bool]$uiGrid.Rows[0].Cells['sel'].Value -or
        @($panelPreview | Where-Object { $_.TargetRow -eq 60 -and -not $_.Selected }).Count -ne 0 -or
        ($uiGrid.Rows[1].Cells['sel'] -is [System.Windows.Forms.DataGridViewCheckBoxCell])) {
        throw '界面选择组件候选后应整组勾选'
    }
    $uiGrid.Rows[0].Cells['sel'].Value = $false
    $panelType.GetMethod('ApplyNameGroupSelectionFromCheck', $flags).Invoke(
        $panel, @($uiGrid.Rows[0].PSObject.BaseObject))
    if (@($panelPreview | Where-Object { $_.TargetRow -eq 60 -and $_.Selected }).Count -ne 0) {
        throw '组件框组首取消勾选应取消整组，隐藏成员不得继续保持选中'
    }
    $uiGrid.Rows[0].Cells['sel'].Value = $true
    $panelType.GetMethod('ApplyNameGroupSelectionFromCheck', $flags).Invoke(
        $panel, @($uiGrid.Rows[0].PSObject.BaseObject))
    if (@($uiGrid.Rows | Where-Object { $_.DefaultCellStyle.BackColor.ToArgb() -eq [System.Drawing.Color]::MistyRose.ToArgb() }).Count -ne 0) {
        throw '界面选择组件候选后应取消红色'
    }
    Write-Host 'PASS 源行定额候选下拉、成员复选框隐藏与组件组确认'

    $scrollPanel = $panelCtor.Invoke([object[]]@($mainForm.PSObject.BaseObject))
    try {
        $scrollPreview = $panelType.GetField('preview', $flags).GetValue($scrollPanel)
        for ($rowNo = 1; $rowNo -le 40; $rowNo++) {
            $filler = New-PreviewItem $rowNo 0 "工程量$rowNo"
            $itemType.GetField('QuotaCode', $flags).SetValue($filler, "Q-$rowNo")
            $scrollPreview.Add($filler)
        }

        $scrollLeader = $scrollPreview[29]
        $itemType.GetField('QuotaCode', $flags).SetValue($scrollLeader, 'Q-A')
        $itemType.GetField('NeedExactNameConfirmation', $flags).SetValue($scrollLeader, $true)
        $itemType.GetField('Selected', $flags).SetValue($scrollLeader, $false)
        $itemType.GetField('AlignNote', $flags).SetValue(
            $scrollLeader, '模板存在同名多来源，已带出候选，需下拉确认')
        $scrollCandidates = [Activator]::CreateInstance($itemType.GetField('NameQuotaCandidates', $flags).FieldType)

        $scrollCandidateA = [Activator]::CreateInstance($candidateType)
        $candidateType.GetField('Key', $flags).SetValue($scrollCandidateA, 'group-a')
        $candidateType.GetField('Label', $flags).SetValue($scrollCandidateA, '来源定额A')
        $scrollCandidateAItem = New-PreviewItem 30 0 '工程量30'
        $itemType.GetField('QuotaCode', $flags).SetValue($scrollCandidateAItem, 'Q-A')
        $candidateType.GetField('Items', $flags).GetValue($scrollCandidateA).Add($scrollCandidateAItem)
        $scrollCandidates.Add($scrollCandidateA)

        $scrollCandidateB = [Activator]::CreateInstance($candidateType)
        $candidateType.GetField('Key', $flags).SetValue($scrollCandidateB, 'group-b')
        $candidateType.GetField('Label', $flags).SetValue($scrollCandidateB, '来源定额B + 来源材料B（组件2条）')
        foreach ($definition in @(@('Q-B', '工程量30'), @('M-B', ''))) {
            $order = $candidateType.GetField('Items', $flags).GetValue($scrollCandidateB).Count
            $member = New-PreviewItem 30 $order $definition[1]
            $itemType.GetField('QuotaCode', $flags).SetValue($member, $definition[0])
            $candidateType.GetField('Items', $flags).GetValue($scrollCandidateB).Add($member)
        }
        $scrollCandidates.Add($scrollCandidateB)
        $itemType.GetField('NameQuotaCandidates', $flags).SetValue($scrollLeader, $scrollCandidates)
        $itemType.GetField('SelectedNameQuotaCandidateKey', $flags).SetValue($scrollLeader, 'group-a')

        $panelType.GetMethod('FillGrid', $flags).Invoke($scrollPanel, $null)
        $scrollPanel.Show()
        [System.Windows.Forms.Application]::DoEvents()
        $scrollGrid = $panelType.GetField('grid', $flags).GetValue($scrollPanel)
        $scrollGrid.CurrentCell = $scrollGrid.Rows[29].Cells['sel']
        $scrollGrid.FirstDisplayedScrollingRowIndex = 15
        [System.Windows.Forms.Application]::DoEvents()
        $topTargetBefore = ($scrollGrid.Rows[$scrollGrid.FirstDisplayedScrollingRowIndex].Tag).TargetRow
        $unaffectedRow = $scrollGrid.Rows[5]

        $updatingField = $panelType.GetField('updatingNameQuotaCell', $flags)
        $updatingField.SetValue($scrollPanel, $true)
        try { $scrollGrid.Rows[29].Cells['sel'].Value = $true }
        finally { $updatingField.SetValue($scrollPanel, $false) }
        $panelType.GetMethod('ApplyNameGroupSelectionFromCheck', $flags).Invoke(
            $scrollPanel, @($scrollGrid.Rows[29].PSObject.BaseObject))

        $confirmedRows = @($scrollGrid.Rows | Where-Object { $_.Tag.TargetRow -eq 30 })
        $topTargetAfterCheck = ($scrollGrid.Rows[$scrollGrid.FirstDisplayedScrollingRowIndex].Tag).TargetRow
        if ($confirmedRows.Count -ne 1 -or $confirmedRows[0].Cells['code'].Value -ne 'Q-A' -or
            -not [bool]$confirmedRows[0].Cells['sel'].Value -or
            -not [String]::IsNullOrWhiteSpace($confirmedRows[0].Tag.Status) -or
            $confirmedRows[0].Tag.NeedExactNameConfirmation -or
            $confirmedRows[0].DefaultCellStyle.BackColor.ToArgb() -eq [System.Drawing.Color]::MistyRose.ToArgb() -or
            $topTargetAfterCheck -ne $topTargetBefore -or
            -not [Object]::ReferenceEquals($unaffectedRow, $scrollGrid.Rows[5])) {
            throw '默认同名候选应能直接勾选确认、取消红色并保持滚动视口'
        }

        $scrollLeaderRow = @($scrollGrid.Rows | Where-Object {
            $_.Tag.TargetRow -eq 30 -and $_.Tag.GroupOrder -eq 0
        })[0]
        $panelType.GetMethod('ApplyNameQuotaOption', $flags).Invoke(
            $scrollPanel, @($scrollLeaderRow.PSObject.BaseObject, '来源定额B + 来源材料B（组件2条）'))
        $scrollTargetRows = @($scrollGrid.Rows | Where-Object { $_.Tag.TargetRow -eq 30 })
        $topTargetAfterChoice = ($scrollGrid.Rows[$scrollGrid.FirstDisplayedScrollingRowIndex].Tag).TargetRow
        if ($scrollTargetRows.Count -ne 2 -or $scrollTargetRows[0].Cells['code'].Value -ne 'Q-B' -or
            $scrollTargetRows[1].Cells['code'].Value -ne 'M-B' -or $topTargetAfterChoice -ne $topTargetBefore -or
            -not [Object]::ReferenceEquals($unaffectedRow, $scrollGrid.Rows[5])) {
            throw '候选切换必须只增删当前组并保持滚动视口'
        }
        Write-Host 'PASS 勾选和候选切换仅局部刷新并保持滚动视口'
    }
    finally {
        $scrollPanel.Close()
        $scrollPanel.Dispose()
    }
}
finally {
    if ($null -ne $panel) { $panel.Dispose() }
    $mainForm.Dispose()
}

$allItems.Add((New-PreviewItem 10 0 '工程量A'))
$allItems.Add((New-PreviewItem 10 1 ''))
$allItems.Add((New-PreviewItem 20 0 '工程量B'))
$replacements.Add((New-PreviewItem 10 9 '工程量A'))
$replacements.Add((New-PreviewItem 10 9 '错误的组员名'))
$replace = $type.GetMethod('ReplacePreviewTargetGroup', $flags)
$replaceArgs = [object[]]@($allItems.PSObject.BaseObject, 10, $replacements.PSObject.BaseObject)
if (-not $replace.Invoke($null, $replaceArgs)) { throw '整组替换应成功' }
if ($allItems.Count -ne 3) { throw "整组替换后总数错误: $($allItems.Count)" }
if ($itemType.GetField('GroupOrder', $flags).GetValue($allItems[0]) -ne 0) { throw '组长顺序应归零' }
if ($itemType.GetField('GroupOrder', $flags).GetValue($allItems[1]) -ne 1) { throw '组员顺序应连续' }
if (-not [String]::IsNullOrEmpty($itemType.GetField('TargetName', $flags).GetValue($allItems[1]))) { throw '组员工程量名应留空' }
Write-Host "PASS 右键原子整组替换"

$groupType = $type.GetNestedType('MappingFeedbackGroup', $flags)
$targetType = $type.GetNestedType('MappingFeedbackTarget', $flags)
$upsert = $type.GetMethod('UpsertMappingBoxGroup', $flags)
$rows = [Activator]::CreateInstance($upsert.GetParameters()[0].ParameterType)
$group = [Activator]::CreateInstance($groupType)
$groupType.GetField('QuantityName', $flags).SetValue($group, '土方外运')
$groupType.GetField('QuantityUnit', $flags).SetValue($group, 'm3')
$targets = $groupType.GetField('Targets', $flags).GetValue($group)
foreach ($code in @('LY-21','LY-34','LY-35')) {
    $target = [Activator]::CreateInstance($targetType)
    $targetType.GetField('Kind', $flags).SetValue($target, 'quota')
    $targetType.GetField('Code', $flags).SetValue($target, $code)
    $targets.Add($target)
}
$upsert.Invoke($null, [object[]]@($rows.PSObject.BaseObject, $group))
if ($rows.Count -ne 3) { throw "组件框应写3条目标记录，实际 $($rows.Count)" }
$boxIds = @($rows | ForEach-Object { $_['box_id'] } | Sort-Object -Unique)
if ($boxIds.Count -ne 1) { throw "组件框目标必须共享一个box_id，实际 $($boxIds.Count)" }
$upsert.Invoke($null, [object[]]@($rows.PSObject.BaseObject, $group))
if ($rows.Count -ne 3) { throw '重复确认不应复制组件目标行' }
if (@($rows | Where-Object { $_['accepted_count'] -ne '2' }).Count -ne 0) { throw '组件框各目标接受次数应一致累加' }
$groupType.GetField('QuantityUnit', $flags).SetValue($group, '10m3')
$upsert.Invoke($null, [object[]]@($rows.PSObject.BaseObject, $group))
if ($rows.Count -ne 3) { throw '同名不同单位不应拆成两套组件关系' }
if (@($rows | Where-Object { $_['accepted_count'] -ne '3' -or $_['quantity_unit'] -ne '10m3' }).Count -ne 0) { throw '名称级关系没有合并单位观察值' }
Write-Host "PASS 组件框整组回写"

$buildBoxIndex = $type.GetMethod('BuildMappingBoxIndex', $flags)
$buildIndexArgs = New-Object 'object[]' 1
$buildIndexArgs[0] = $rows.PSObject.BaseObject
$boxIndex = $buildBoxIndex.Invoke($null, $buildIndexArgs)
$lookupBox = $type.GetMethod('LookupMappingBox', $flags)
$lookupArgs = New-Object 'object[]' 2
$lookupArgs[0] = [string]'土方外运'
$lookupArgs[1] = $boxIndex.PSObject.BaseObject
$boxMatches = $lookupBox.Invoke($null, $lookupArgs)
if ($boxMatches.Count -ne 1) { throw "组件框回读应返回1个整框，实际 $($boxMatches.Count)" }
$boxCandidateType = $type.GetNestedType('BoxCandidate', $flags)
$targetsField = $boxCandidateType.GetField('Targets', $flags)
$readTargets = $targetsField.GetValue($boxMatches[0])
if ($readTargets.Count -ne 3) { throw "组件框回读应保留3个目标，实际 $($readTargets.Count)" }
Write-Host "PASS 组件框整组回读"

$fixturePath = Join-Path $env:TEMP 'reco-template-chapter-test.xlsx'
$targetFixturePath = Join-Path $env:TEMP 'reco-template-target-test.xlsx'
$workbook = New-Object NPOI.XSSF.UserModel.XSSFWorkbook
$targetWorkbook = New-Object NPOI.XSSF.UserModel.XSSFWorkbook
try {
    $sheet = $workbook.CreateSheet('测试')
    $fixtureRows = @(
        @('第一章 路基工程', $null),
        @('挖土方', 10),
        @('第二章 站场工程', $null),
        @('挖土方', 20),
        @('零数量工程量', 0)
    )
    for ($i = 0; $i -lt $fixtureRows.Count; $i++) {
        $excelRow = $sheet.CreateRow($i)
        $excelRow.CreateCell(0).SetCellValue([string]$fixtureRows[$i][0])
        if ($null -ne $fixtureRows[$i][1]) {
            $excelRow.CreateCell(3).SetCellValue([double]$fixtureRows[$i][1])
        }
    }
    $stream = [System.IO.File]::Open($fixturePath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    try { $workbook.Write($stream) } finally { $stream.Dispose() }

    $targetSheet = $targetWorkbook.CreateSheet('测试')
    $targetRows = @(
        @('第一章 路基工程', $null),
        @('挖土方', 22),
        @('第二章 站场工程', $null),
        @('挖土方', 33),
        @('零数量工程量', 0)
    )
    for ($i = 0; $i -lt $targetRows.Count; $i++) {
        $targetExcelRow = $targetSheet.CreateRow($i)
        $targetExcelRow.CreateCell(0).SetCellValue([string]$targetRows[$i][0])
        if ($null -ne $targetRows[$i][1]) {
            $targetExcelRow.CreateCell(3).SetCellValue([double]$targetRows[$i][1])
        }
    }
    $mergedSheet = $targetWorkbook.CreateSheet('组合')
    $mergedRows = @(
        @('回归主工程量A901', 2),
        @('回归辅助工程量B902', 1.5),
        @('回归纯取数C903', 0.5)
    )
    for ($i = 0; $i -lt $mergedRows.Count; $i++) {
        $mergedExcelRow = $mergedSheet.CreateRow($i)
        $mergedExcelRow.CreateCell(0).SetCellValue([string]$mergedRows[$i][0])
        $mergedExcelRow.CreateCell(3).SetCellValue([double]$mergedRows[$i][1])
    }
    $unitSheet = $targetWorkbook.CreateSheet('单位')
    $unitRow = $unitSheet.CreateRow(0)
    $unitRow.CreateCell(1).SetCellValue('热镀锌钢管')
    $unitRow.CreateCell(2).SetCellValue('SC20')
    $unitRow.CreateCell(3).SetCellValue('m')
    $unitRow.CreateCell(4).SetCellValue([double]8000)
    $unitRow.CreateCell(5).SetCellValue([double]1400)
    $singleSheet = $targetWorkbook.CreateSheet('单一')
    $singleRow = $singleSheet.CreateRow(0)
    $singleRow.CreateCell(0).SetCellValue('挖土方')
    $singleRow.CreateCell(3).SetCellValue([double]44)
    $targetStream = [System.IO.File]::Open($targetFixturePath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
    try { $targetWorkbook.Write($targetStream) } finally { $targetStream.Dispose() }

    $templateType = $type.GetNestedType('FillTemplate', $flags)
    $templateRowType = $type.GetNestedType('FillTemplateRow', $flags)
    $template = [Activator]::CreateInstance($templateType)
    $templateRow = [Activator]::CreateInstance($templateRowType)
    $templateType.GetField('WorkbookPath', $flags).SetValue($template, [string]$fixturePath)
    $templateRowType.GetField('SourceSheet', $flags).SetValue($templateRow, [string]'测试')
    $templateRowType.GetField('SourceExpr', $flags).SetValue($templateRow, [string]'D5')
    $templateType.GetField('Rows', $flags).GetValue($template).Add($templateRow)
    $populateChapters = $type.GetMethod('PopulateTemplateMatchChapters', $flags)
    $populateChapters.Invoke($null, @($template))
    if ($templateRowType.GetField('MatchChapter', $flags).GetValue($templateRow) -ne '第二章 站场工程') {
        throw '源模板零数量行仍应记录所属章节'
    }

    $readRows = $type.GetMethod('ReadTargetQtyRows', $flags)
    $actualRows = $readRows.Invoke($null, [object[]]@([string]$fixturePath, [string]'测试', [int]4))
    if ($actualRows.Count -ne 2) { throw "零数量应丢弃，应剩2行，实际 $($actualRows.Count)" }
    if ($actualRows[0].Chapter -ne '第一章 路基工程' -or $actualRows[1].Chapter -ne '第二章 站场工程') {
        throw "章节快照错误: '$($actualRows[0].Chapter)' / '$($actualRows[1].Chapter)'"
    }
    if (@($actualRows | Where-Object { $_.RawName -eq '零数量工程量' }).Count -ne 0) {
        throw '零数量工程量未丢弃'
    }
    Write-Host "PASS 真实工作簿章节快照与零数量丢弃"

    $unitRows = $readRows.Invoke($null,
        [object[]]@([string]$targetFixturePath, [string]'单位', [int]6))
    if ($unitRows.Count -ne 1 -or $unitRows[0].Unit -ne 'm') {
        $actualUnit = if ($unitRows.Count -eq 0) { '' } else { $unitRows[0].Unit }
        throw "数量 F 列应越过数值 E 列读取 D 列单位 m: '$actualUnit'"
    }
    $buildQty = $type.GetMethod('BuildNameDrivenQtyText', $flags)
    if ($buildQty.Invoke($null, @('1400', $unitRows[0].Unit, 'hm')) -ne '1400/100') {
        throw '无模板表达式时 m 到 hm 必须生成 1400/100'
    }
    Write-Host 'PASS 数量列向左越过数值列读取单位并生成 1400/100'

    function New-NamePreviewTemplate([string[]]$codes) {
        $nameTemplate = [Activator]::CreateInstance($templateType)
        $templateType.GetField('Name', $flags).SetValue($nameTemplate, '重复名称夹具')
        $templateType.GetField('MatchBy', $flags).SetValue($nameTemplate, 'name')
        $templateType.GetField('WorkbookPath', $flags).SetValue($nameTemplate, [string]$fixturePath)
        for ($index = 0; $index -lt $codes.Count; $index++) {
            $nameRow = [Activator]::CreateInstance($templateRowType)
            $templateRowType.GetField('MatchName', $flags).SetValue($nameRow, '挖土方 m3')
            $templateRowType.GetField('SourceSheet', $flags).SetValue($nameRow, '测试')
            $templateRowType.GetField('SourceExpr', $flags).SetValue($nameRow, $(if ($index -eq 0) { 'D2' } else { 'D4' }))
            $templateRowType.GetField('QuotaCode', $flags).SetValue($nameRow, $codes[$index])
            $templateRowType.GetField('SourceName', $flags).SetValue($nameRow, "定额$($index + 1)")
            $templateRowType.GetField('Unit', $flags).SetValue($nameRow, 'm3')
            $templateRowType.GetField('ItemNo', $flags).SetValue($nameRow, '01')
            $templateRowType.GetField('SourceQuotaSeq', $flags).SetValue($nameRow, [long](100 + $index))
            $templateRowType.GetField('OrderInItem', $flags).SetValue($nameRow, [int]$index)
            $templateType.GetField('Rows', $flags).GetValue($nameTemplate).Add($nameRow)
        }
        return $nameTemplate
    }

    $buildPreview = $type.GetMethod('BuildPreview_NameDriven', $flags)
    $previewMainForm = New-Object System.Windows.Forms.Form
    try {
        $reuseArgs = [object[]]@($previewMainForm.PSObject.BaseObject, (New-NamePreviewTemplate @('Q-1')), [string]$targetFixturePath, '测试', 'D', $null)
        $reusePreview = $buildPreview.Invoke($null, $reuseArgs)
        if ($reusePreview.Count -ne 2 -or @($reusePreview | Where-Object { $_.QuotaCode -ne 'Q-1' }).Count -ne 0) {
            throw '重复目标应全部带出唯一绑定 Q-1'
        }
        if ($reusePreview[0].QuantityText -ne '22' -or $reusePreview[1].QuantityText -ne '33') {
            throw "名字驱动没有读取目标工作簿数量: '$($reusePreview[0].QuantityText)' / '$($reusePreview[1].QuantityText)'"
        }
        if (@($reusePreview | Where-Object { -not $_.NeedExactNameConfirmation -or $_.Selected }).Count -ne 0) {
            throw '重复目标唯一绑定应标红并默认不勾选'
        }

        $choiceArgs = [object[]]@($previewMainForm.PSObject.BaseObject, (New-NamePreviewTemplate @('Q-1', 'Q-2')), [string]$targetFixturePath, '测试', 'D', $null)
        $choicePreview = $buildPreview.Invoke($null, $choiceArgs)
        if ($choicePreview.Count -ne 2 -or @($choicePreview | Where-Object { $_.QuotaCode -ne 'Q-1' }).Count -ne 0) {
            throw '多个绑定初始应带出稳定的第一候选 Q-1'
        }
        if (@($choicePreview | Where-Object { $null -eq $_.NameQuotaCandidates -or $_.NameQuotaCandidates.Count -ne 2 }).Count -ne 0) {
            throw '每个重复目标行都应保留两个独立下拉候选'
        }
        if (@($choicePreview | Where-Object {
            -not [String]::IsNullOrWhiteSpace($_.Status) -or $_.AlignNote -notmatch '需下拉确认'
        }).Count -ne 0) {
            throw '同名多来源应以待确认提示标红，不得伪装成阻断错误状态'
        }
        $choiceLabels = @($choicePreview[0].NameQuotaCandidates | ForEach-Object { $_.Label })
        if (($choiceLabels -join '|') -ne '定额1|定额2') {
            throw "下拉候选应显示源行定额名称而不是定额编号: '$($choiceLabels -join "|")'"
        }

        $sameArgs = [object[]]@($previewMainForm.PSObject.BaseObject,
            (New-NamePreviewTemplate @('Q-1', 'Q-1')), [string]$targetFixturePath, '单一', 'D', $null)
        $samePreview = $buildPreview.Invoke($null, $sameArgs)
        if ($samePreview.Count -ne 1 -or $samePreview[0].QuotaCode -ne 'Q-1' -or
            $samePreview[0].Selected -or -not $samePreview[0].NeedExactNameConfirmation -or
            ($null -ne $samePreview[0].NameQuotaCandidates -and $samePreview[0].NameQuotaCandidates.Count -gt 1) -or
            -not [String]::IsNullOrWhiteSpace($samePreview[0].Status) -or
            $samePreview[0].AlignNote -notmatch '对应相同绑定') {
            throw '同名多来源对应相同绑定时应默认带出唯一结果、等待勾选且不显示下拉'
        }

        $operandType = $type.GetNestedType('FillOperand', $flags)
        $operandListType = [System.Collections.Generic.List``1].MakeGenericType($operandType)
        function Add-MergedTemplateRow([object]$targetTemplate, [string]$code, [string]$matchName,
            [string]$expr, [int]$order, [string[]]$operandNames) {
            $row = [Activator]::CreateInstance($templateRowType)
            $templateRowType.GetField('MatchName', $flags).SetValue($row, $matchName)
            $templateRowType.GetField('SourceSheet', $flags).SetValue($row, '组合')
            $templateRowType.GetField('SourceExpr', $flags).SetValue($row, $expr)
            $templateRowType.GetField('QuotaCode', $flags).SetValue($row, $code)
            $templateRowType.GetField('SourceName', $flags).SetValue($row, $code)
            $templateRowType.GetField('Unit', $flags).SetValue($row, '10m2')
            $templateRowType.GetField('ItemNo', $flags).SetValue($row, 'REG-01')
            $templateRowType.GetField('SourceQuotaSeq', $flags).SetValue($row, [long](9000 + $order))
            $templateRowType.GetField('OrderInItem', $flags).SetValue($row, $order)
            if ($null -ne $operandNames -and $operandNames.Count -gt 0) {
                $operands = [Activator]::CreateInstance($operandListType)
                foreach ($operandName in $operandNames) {
                    $operand = [Activator]::CreateInstance($operandType)
                    $operandType.GetField('Name', $flags).SetValue($operand, $operandName)
                    $operandType.GetField('Op', $flags).SetValue($operand, '+')
                    [void]$operands.Add($operand)
                }
                $templateRowType.GetField('Operands', $flags).SetValue($row, $operands.PSObject.BaseObject)
            }
            [void]$templateType.GetField('Rows', $flags).GetValue($targetTemplate).Add($row)
        }

        $mergedTemplate = [Activator]::CreateInstance($templateType)
        $templateType.GetField('Name', $flags).SetValue($mergedTemplate, '组合表达式回归夹具')
        $templateType.GetField('MatchBy', $flags).SetValue($mergedTemplate, 'name')
        $templateType.GetField('WorkbookPath', $flags).SetValue($mergedTemplate, [string]$fixturePath)
        Add-MergedTemplateRow $mergedTemplate 'Q-MAIN' '回归主工程量A901' 'D1' 0 $null
        Add-MergedTemplateRow $mergedTemplate 'Q-TRANSPORT' '回归主工程量A901' 'D1+D2+D3' 1 @(
            '回归主工程量A901', '回归辅助工程量B902', '回归纯取数C903')
        Add-MergedTemplateRow $mergedTemplate 'Q-AUX-1' '回归辅助工程量B902' 'D2*10' 2 $null
        Add-MergedTemplateRow $mergedTemplate 'Q-AUX-2' '回归辅助工程量B902' 'D2' 3 $null
        Add-MergedTemplateRow $mergedTemplate 'Q-TRANSPORT' '回归辅助工程量B902' 'D2' 4 $null
        Add-MergedTemplateRow $mergedTemplate '1009001003*1.02' '回归辅助工程量B902' 'D2' 5 $null

        $mergedArgs = [object[]]@($previewMainForm.PSObject.BaseObject, $mergedTemplate.PSObject.BaseObject,
            [string]$targetFixturePath, '组合', 'D', $null)
        $mergedPreview = $buildPreview.Invoke($null, $mergedArgs)
        $codes = @($mergedPreview | Where-Object { -not [String]::IsNullOrWhiteSpace($_.QuotaCode) } |
            ForEach-Object { $_.QuotaCode })
        foreach ($expectedCode in @('Q-MAIN', 'Q-TRANSPORT', 'Q-AUX-1', 'Q-AUX-2', '1009001003*1.02')) {
            if ($codes -notcontains $expectedCode) { throw "组合表达式回归缺少 $expectedCode" }
        }
        if (@($codes | Where-Object { $_ -eq 'Q-TRANSPORT' }).Count -ne 2) {
            throw '辅助工程量存在独立绑定时，即使与合计定额相同也必须保留'
        }
        $transport = @($mergedPreview | Where-Object { $_.QuotaCode -eq 'Q-TRANSPORT' })[0]
        if ($transport.QuantityText -notmatch '2' -or $transport.QuantityText -notmatch '1[\.,]5' -or
            $transport.QuantityText -notmatch '0[\.,]5') {
            throw "组合表达式没有同时代入三行数量: '$($transport.QuantityText)'"
        }
        $auxiliary = @($mergedPreview | Where-Object { $_.TargetRow -eq 2 } | Sort-Object GroupOrder)
        if ($auxiliary.Count -ne 4 -or $auxiliary[0].QuotaCode -ne 'Q-AUX-1' -or
            $auxiliary[1].QuotaCode -ne 'Q-AUX-2' -or $auxiliary[2].QuotaCode -ne 'Q-TRANSPORT' -or
            $auxiliary[3].QuotaCode -ne '1009001003*1.02') {
            throw '参与组合表达式的辅助行没有保留自己的独立定额和材料'
        }
        $auxNote = if ([String]::IsNullOrWhiteSpace($auxiliary[0].Status)) {
            $auxiliary[0].AlignNote
        } else { $auxiliary[0].Status }
        if ($auxNote -notmatch '同时参与第 1 行的表达式取数') {
            throw "辅助行缺少同时参与说明: '$auxNote'"
        }
        $operandOnly = @($mergedPreview | Where-Object { $_.TargetRow -eq 3 })
        if ($operandOnly.Count -ne 1 -or -not [String]::IsNullOrWhiteSpace($operandOnly[0].QuotaCode) -or
            $operandOnly[0].Selected -or $operandOnly[0].NeedManualQuota -or
            $operandOnly[0].AlignNote -notmatch '无独立定额匹配') {
            throw '纯辅助取数行应保留不可写入占位且不得产生错误定额'
        }
        Write-Host 'PASS 组合表达式辅助行保留自身绑定并正确处理纯取数行'

        $buildColumnPreview = $type.GetMethod('BuildPreview_ColumnAnchor', $flags)
        $columnTemplate = New-NamePreviewTemplate @('Q-1')
        $columnPreview = $buildColumnPreview.Invoke($null,
            [object[]]@($columnTemplate, [string]$targetFixturePath, '测试', 'D'))
        if ($columnPreview.Count -ne 1 -or $columnPreview[0].QuantityText -ne '22') {
            throw "列锚点没有读取目标工作簿数量: '$($columnPreview[0].QuantityText)'"
        }
        Write-Host 'PASS 跨工作簿取数、重复目标唯一绑定复用与多绑定候选预览'
    }
    finally { $previewMainForm.Dispose() }
}
finally {
    $workbook.Close()
    $targetWorkbook.Close()
    if (Test-Path -LiteralPath $fixturePath) { Remove-Item -LiteralPath $fixturePath -Force }
    if (Test-Path -LiteralPath $targetFixturePath) { Remove-Item -LiteralPath $targetFixturePath -Force }
}
Write-Host "全部通过"
