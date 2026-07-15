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

$templateType = $type.GetNestedType('FillTemplate', $flags)
$templateRowType = $type.GetNestedType('FillTemplateRow', $flags)
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
$applyCandidate = $type.GetMethod('ApplyExactNameCandidate', $flags)
if ($null -eq $candidateType -or $null -eq $confirmExact -or $null -eq $applyCandidate) {
    throw '缺少重复名称候选或确认入口'
}

$singleItems = [Activator]::CreateInstance($itemListType)
$singleItems.Add((New-PreviewItem 40 0 '钢管 SC20 m'))
$singleItems.Add((New-PreviewItem 40 1 ''))
foreach ($member in $singleItems) {
    $itemType.GetField('NeedExactNameConfirmation', $flags).SetValue($member, $true)
    $itemType.GetField('Selected', $flags).SetValue($member, $false)
}
if (-not $confirmExact.Invoke($null, [object[]]@($singleItems.PSObject.BaseObject, 40))) {
    throw '唯一绑定整组确认应成功'
}
if (@($singleItems | Where-Object { -not $_.Selected -or $_.NeedExactNameConfirmation }).Count -ne 0) {
    throw '唯一绑定确认后应整组勾选并取消红色确认状态'
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
Write-Host 'PASS 重复名称唯一确认与候选整组切换'

$panelType = $type.GetNestedType('TemplateFillPanel', $flags)
$panelCtor = @($panelType.GetConstructors($flags) | Where-Object { $_.GetParameters().Count -eq 1 })[0]
$mainForm = New-Object System.Windows.Forms.Form
$panel = $null
try {
    $panel = $panelCtor.Invoke([object[]]@($mainForm.PSObject.BaseObject))
    $panelPreview = $panelType.GetField('preview', $flags).GetValue($panel)
    $uiLeader = New-PreviewItem 60 0 '重复工程量'
    $itemType.GetField('QuotaCode', $flags).SetValue($uiLeader, 'DY-959')
    $itemType.GetField('NeedExactNameConfirmation', $flags).SetValue($uiLeader, $true)
    $itemType.GetField('Selected', $flags).SetValue($uiLeader, $false)

    $uiCandidates = [Activator]::CreateInstance($itemType.GetField('NameQuotaCandidates', $flags).FieldType)
    $uiFirst = [Activator]::CreateInstance($candidateType)
    $candidateType.GetField('Key', $flags).SetValue($uiFirst, 'group-a')
    $candidateType.GetField('Label', $flags).SetValue($uiFirst, 'DY-959 明配钢管')
    $candidateType.GetField('Items', $flags).GetValue($uiFirst).Add((New-PreviewItem 60 0 '重复工程量'))
    $itemType.GetField('QuotaCode', $flags).SetValue($candidateType.GetField('Items', $flags).GetValue($uiFirst)[0], 'DY-959')
    $uiCandidates.Add($uiFirst)

    $uiSecond = [Activator]::CreateInstance($candidateType)
    $candidateType.GetField('Key', $flags).SetValue($uiSecond, 'group-b')
    $candidateType.GetField('Label', $flags).SetValue($uiSecond, 'DY-519 + ZLF*1.01（组件2条）')
    $uiSecondItems = $candidateType.GetField('Items', $flags).GetValue($uiSecond)
    $uiSecondItems.Add((New-PreviewItem 60 0 '重复工程量'))
    $uiSecondItems.Add((New-PreviewItem 60 1 ''))
    $itemType.GetField('QuotaCode', $flags).SetValue($uiSecondItems[0], 'DY-519')
    $itemType.GetField('QuotaCode', $flags).SetValue($uiSecondItems[1], 'ZLF*1.01')
    $uiCandidates.Add($uiSecond)
    $itemType.GetField('NameQuotaCandidates', $flags).SetValue($uiLeader, $uiCandidates)
    $itemType.GetField('SelectedNameQuotaCandidateKey', $flags).SetValue($uiLeader, 'group-a')
    $panelPreview.Add($uiLeader)

    $panelType.GetMethod('FillGrid', $flags).Invoke($panel, $null)
    $uiGrid = $panelType.GetField('grid', $flags).GetValue($panel)
    if ($uiGrid.Rows.Count -ne 1 -or -not $uiGrid.Rows[0].Cells['sel'].ReadOnly -or $uiGrid.Rows[0].Cells['code'].ReadOnly) {
        throw '多候选红色行应锁定勾选框并开放定额下拉'
    }
    if ($uiGrid.Rows[0].DefaultCellStyle.BackColor.ToArgb() -ne [System.Drawing.Color]::MistyRose.ToArgb()) {
        throw '多候选未确认行应标红'
    }
    $prepareDropDown = $panelType.GetMethod('PrepareNameQuotaDropDown', $flags)
    if (-not $prepareDropDown.Invoke($panel, @($uiGrid.Rows[0].PSObject.BaseObject))) { throw '定额下拉应创建成功' }
    if (-not ($uiGrid.Rows[0].Cells['code'] -is [System.Windows.Forms.DataGridViewComboBoxCell])) {
        throw '定额编号单元格应切换为下拉框'
    }
    $panelType.GetMethod('ApplyNameQuotaOption', $flags).Invoke($panel,
        @($uiGrid.Rows[0].PSObject.BaseObject, 'DY-519 + ZLF*1.01（组件2条）'))
    if ($uiGrid.Rows.Count -ne 2 -or $uiGrid.Rows[0].Cells['code'].Value -ne 'DY-519' -or
        $uiGrid.Rows[1].Cells['code'].Value -ne 'ZLF*1.01') {
        throw '界面选择组件候选后应展开完整定额组'
    }
    if (@($uiGrid.Rows | Where-Object { -not [bool]$_.Cells['sel'].Value }).Count -ne 0) {
        throw '界面选择组件候选后应整组勾选'
    }
    if (@($uiGrid.Rows | Where-Object { $_.DefaultCellStyle.BackColor.ToArgb() -eq [System.Drawing.Color]::MistyRose.ToArgb() }).Count -ne 0) {
        throw '界面选择组件候选后应取消红色'
    }
    Write-Host 'PASS 定额候选下拉与组件组界面确认'
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

    function New-NamePreviewTemplate([string[]]$codes) {
        $nameTemplate = [Activator]::CreateInstance($templateType)
        $templateType.GetField('Name', $flags).SetValue($nameTemplate, '重复名称夹具')
        $templateType.GetField('MatchBy', $flags).SetValue($nameTemplate, 'name')
        $templateType.GetField('WorkbookPath', $flags).SetValue($nameTemplate, [string]$fixturePath)
        for ($index = 0; $index -lt $codes.Count; $index++) {
            $nameRow = [Activator]::CreateInstance($templateRowType)
            $templateRowType.GetField('MatchName', $flags).SetValue($nameRow, '挖土方')
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
