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

$selectedItems = [Activator]::CreateInstance($itemListType)
$writtenItems = [Activator]::CreateInstance($itemListType)
$selectedItems.Add((New-PreviewItem 30 0 '工程量C'))
$selectedItems.Add((New-PreviewItem 30 1 ''))
$writtenItems.Add($selectedItems[0])
$filterWritten = $type.GetMethod('FilterFullyWrittenNameGroups', $flags)
$partial = $filterWritten.Invoke($null, [object[]]@($selectedItems.PSObject.BaseObject, $writtenItems.PSObject.BaseObject))
if ($partial.Count -ne 0) { throw '组件框部分写入成功时不得学习' }
$writtenItems.Add($selectedItems[1])
$complete = $filterWritten.Invoke($null, [object[]]@($selectedItems.PSObject.BaseObject, $writtenItems.PSObject.BaseObject))
if ($complete.Count -ne 2) { throw '组件框全部写入成功时应整组学习' }
Write-Host "PASS 组件框部分失败不学习"

$groupType = $type.GetNestedType('MappingFeedbackGroup', $flags)
$targetType = $type.GetNestedType('MappingFeedbackTarget', $flags)
$upsert = $type.GetMethod('UpsertMappingBoxGroup', $flags)
$rows = [Activator]::CreateInstance($upsert.GetParameters()[0].ParameterType)
$group = [Activator]::CreateInstance($groupType)
$groupType.GetField('QuantityName', $flags).SetValue($group, '土方外运')
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
$workbook = New-Object NPOI.XSSF.UserModel.XSSFWorkbook
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
}
finally {
    $workbook.Close()
    if (Test-Path -LiteralPath $fixturePath) { Remove-Item -LiteralPath $fixturePath -Force }
}
Write-Host "全部通过"
