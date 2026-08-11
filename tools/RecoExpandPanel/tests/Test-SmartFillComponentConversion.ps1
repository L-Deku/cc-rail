$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..'))
$dll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL)) { $env:RECO_EXPAND_DLL } else { Join-Path $repoRoot 'RecoQuotaRecommend\bin\RecoExpandPanel.dll' }
if (-not (Test-Path -LiteralPath $dll)) { throw "Missing DLL: $dll" }

$assembly = [System.Reflection.Assembly]::LoadFrom($dll)
$type = $assembly.GetType('RecoNet.FormPanel', $true)
$flags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$nestedFlags = [System.Reflection.BindingFlags]'Public,NonPublic'

function Require-Method([Type]$Owner, [string]$Name) {
    $method = $Owner.GetMethod($Name, $flags)
    if ($null -eq $method) { throw "Missing regression seam: $($Owner.FullName).$Name" }
    return $method
}

$standardScale = Require-Method $type 'TryBuildExcelLinkUnitScaleSuffix'
$countScale = Require-Method $type 'TryBuildConfirmedCountUnitScaleSuffix'
function Assert-Scale([System.Reflection.MethodInfo]$Method, [string]$From, [string]$To,
    [bool]$ExpectedOk, [string]$ExpectedSuffix) {
    $args = [object[]]::new(3)
    $args[0] = $From
    $args[1] = $To
    $args[2] = $null
    $ok = [bool]$Method.Invoke($null, $args)
    if ($ok -ne $ExpectedOk -or ($ok -and [string]$args[2] -ne $ExpectedSuffix)) {
        throw "Scale mismatch: $From -> $To, ok=$ok suffix='$($args[2])'"
    }
}

foreach ($case in @(
    @('m', 'km', '/1000'), @('m', 'hm', '/100'), @('km', 'm', '*1000'),
    @('kg', 't', '/1000'), @('t', 'kg', '*1000'), @('m', '100m', '/100')
)) { Assert-Scale $standardScale $case[0] $case[1] $true $case[2] }
foreach ($case in @(@('个', '台'), @('项', '10套'), @('m', 'm2'), @('m', 'kg'))) {
    Assert-Scale $standardScale $case[0] $case[1] $false ''
}
Write-Host 'PASS 标准同量纲换算自动完成，跨基础单位不静默按1:1'

foreach ($case in @(
    @('个', '台', ''), @('张', '个', ''), @('项', '10套', '/10'), @('10处', '台', '*10'), @('100个', '10套', '*10')
)) { Assert-Scale $countScale $case[0] $case[1] $true $case[2] }
foreach ($case in @(@('m', 'kg'), @('m', 'm2'), @('亩', '公顷'))) {
    Assert-Scale $countScale $case[0] $case[1] $false ''
}
Write-Host 'PASS 离散计数单位只在明确确认后按倍率换算'

$previewType = $type.GetNestedType('FillPreviewItem', $nestedFlags)
$panelType = $type.GetNestedType('TemplateFillPanel', $nestedFlags)
if ($null -eq $previewType -or $null -eq $panelType) { throw 'Missing preview/panel nested type' }

$targetRowType = $type.GetNestedType('TargetQtyRow', $nestedFlags)
$targetType = $type.GetNestedType('SmartBoxTarget', $nestedFlags)
$mapEntryType = $type.GetNestedType('SmartMapEntry', $nestedFlags)
$resolutionType = $type.GetNestedType('SmartTargetEntryResolution', $nestedFlags)
$snapshotType = $type.GetNestedType('SmartLearningSnapshot', $nestedFlags)
$projectQuotaType = $type.GetNestedType('ProjectQuota', $nestedFlags)
$appendSmartItems = Require-Method $type 'AppendSmartItems'
function Invoke-SmartUnitPreview([string]$SourceUnit, [string]$QuotaUnit) {
    $row = [Activator]::CreateInstance($targetRowType, $true).PSObject.BaseObject
    foreach ($pair in @{ Row=7; RawName='测试工程量'; DisplayName='测试工程量'; NormName='测试工程量';
        Chapter='测试章节'; Unit=$SourceUnit; Quantity=[decimal]100; QuantityText='100' }.GetEnumerator()) {
        $targetRowType.GetField($pair.Key, $flags).SetValue($row, $pair.Value)
    }
    $rowListType = [Collections.Generic.List``1].MakeGenericType($targetRowType)
    $rows = [Activator]::CreateInstance($rowListType); [void]$rows.Add($row)

    $target = [Activator]::CreateInstance($targetType, $true).PSObject.BaseObject
    foreach ($pair in @{ Kind='quota'; Code='TEST-UNIT'; Name='测试定额'; Unit=$QuotaUnit }.GetEnumerator()) {
        $targetType.GetField($pair.Key, $flags).SetValue($target, $pair.Value)
    }
    $entry = [Activator]::CreateInstance($mapEntryType, $true).PSObject.BaseObject
    [void]$mapEntryType.GetField('Targets', $flags).GetValue($entry).Add($target)
    $resolution = [Activator]::CreateInstance($resolutionType, $true).PSObject.BaseObject
    foreach ($pair in @{ Target=$target; EntryCode='0101-01'; EntryName='测试条目'; EntrySeq=[long]1 }.GetEnumerator()) {
        $resolutionType.GetField($pair.Key, $flags).SetValue($resolution, $pair.Value)
    }
    $resolutionListType = [Collections.Generic.List``1].MakeGenericType($resolutionType)
    $resolutions = [Activator]::CreateInstance($resolutionListType); [void]$resolutions.Add($resolution)

    $snapshot = [Activator]::CreateInstance($snapshotType, $true).PSObject.BaseObject
    foreach ($pair in @{ Method='2020'; SoftwarePartition='2020'; MethodNo='30号文' }.GetEnumerator()) {
        $snapshotType.GetField($pair.Key, $flags).SetValue($snapshot, $pair.Value)
    }
    $projectEntries = [Collections.Generic.Dictionary[string,long]]::new([StringComparer]::OrdinalIgnoreCase)
    $projectEntries.Add('0101-01', [long]1)
    $quota = [Activator]::CreateInstance($projectQuotaType, $true).PSObject.BaseObject
    foreach ($pair in @{ Code='TEST-UNIT'; Name='测试定额'; Unit=$QuotaUnit; QuotaSeq=[long]1; IsLibrary=$false }.GetEnumerator()) {
        $projectQuotaType.GetField($pair.Key, $flags).SetValue($quota, $pair.Value)
    }
    $quotaDictionaryType = [Collections.Generic.Dictionary``2].MakeGenericType([string], $projectQuotaType)
    $currentQuotas = [Activator]::CreateInstance($quotaDictionaryType, [StringComparer]::OrdinalIgnoreCase)
    $currentQuotas.Add('TEST-UNIT', $quota)
    $previewListType = [Collections.Generic.List``1].MakeGenericType($previewType)
    $items = [Activator]::CreateInstance($previewListType)
    $args = [object[]]::new(13)
    $args[0]=$items; $args[1]=$row; $args[2]=$rows; $args[3]=$entry; $args[4]=$snapshot
    $args[5]=$projectEntries; $args[6]=$currentQuotas; $args[7]=$resolutions; $args[8]=$false
    $args[9]='SQL精确'; $args[10]='测试签名|'; $args[11]=$null; $args[12]=$null
    [void]$appendSmartItems.Invoke($null, $args)
    if ($items.Count -ne 1) { throw "Expected one preview item, got $($items.Count)" }
    return $items[0]
}

$countPreview = Invoke-SmartUnitPreview '个' '台'
if ([string]$countPreview.Status -ne '待确认计数单位1:1' -or $countPreview.Selected) {
    throw "Count-unit preview must start as soft confirmation: '$($countPreview.Status)'"
}
$sheetPreview = Invoke-SmartUnitPreview '张' '个'
if ([string]$sheetPreview.Status -ne '待确认计数单位1:1' -or $sheetPreview.Selected) {
    throw "Sheet-to-piece preview must start as soft confirmation: '$($sheetPreview.Status)'"
}
$metricPreview = Invoke-SmartUnitPreview 'm' 'km'
if ([string]$metricPreview.QuantityText -ne '100/1000' -or -not [String]::IsNullOrWhiteSpace([string]$metricPreview.Status)) {
    throw 'Standard metric conversion must be automatic and safe'
}
$crossPreview = Invoke-SmartUnitPreview 'm' 'm2'
if ([string]$crossPreview.Status -ne '缺跨量纲换算系数' -or $crossPreview.Selected) {
    throw 'Cross-dimension preview without SQL formula must be a hard block'
}
$areaPreview = Invoke-SmartUnitPreview '亩' '公顷'
if ([string]$areaPreview.Status -ne '缺跨量纲换算系数') { throw 'Area units must not enter count-unit 1:1 confirmation' }
Write-Host 'PASS SmartFill按标准换算、计数软确认和跨量纲硬阻断分流'

function New-PreviewItem([string]$SourceUnit, [string]$QuotaUnit, [string]$Quantity,
    [string]$Status, [bool]$Selected, [int]$Order) {
    $item = [Activator]::CreateInstance($previewType).PSObject.BaseObject
    foreach ($pair in @{
        IsNameDriven=$true; TargetRow=7; GroupOrder=$Order; TargetUnit=$SourceUnit; Unit=$QuotaUnit;
        TargetQuantityText=$Quantity; QuantityText=$Quantity; Status=$Status; Selected=$Selected;
        QuotaCode=('TEST-' + $Order); SourceName=('测试定额' + $Order);
        ChosenQuotaSeq=[long]1; NeighborSourceQuotaSeq=[long]1; ChosenItemSeq=[long]1
    }.GetEnumerator()) { $previewType.GetField($pair.Key, $flags).SetValue($item, $pair.Value) }
    return $item
}

$confirmCount = Require-Method $panelType 'ConfirmPendingCountUnitScale'
$countItem = New-PreviewItem '项' '10套' '20' '待确认计数单位1:1' $false 0
if (-not [bool]$confirmCount.Invoke($null, @($countItem)) -or
    [string]$countItem.QuantityText -ne '20/10' -or
    -not [String]::IsNullOrWhiteSpace([string]$countItem.Status)) {
    throw 'Group-leader confirmation did not apply the count-unit scale and clear only the soft state'
}

$applyEdited = Require-Method $panelType 'ApplyEditedNameQuotaQuantity'
$crossItem = New-PreviewItem 'm' 'm2' '100' '缺跨量纲换算系数；缺条目' $false 2
if ([bool]$applyEdited.Invoke($null, @($crossItem, '100'))) {
    throw 'Unchanged quantity text must not clear a conversion block'
}
if (-not [bool]$applyEdited.Invoke($null, @($crossItem, '原数量*0.35'))) {
    throw 'A valid manual cross-dimension expression should clear its conversion block'
}
if ([string]$crossItem.QuantityText -match '原数量' -or
    [string]$crossItem.Status -ne '缺条目' -or
    [string]$crossItem.AlignNote -notmatch '数量已人工确认') {
    throw "Manual quantity confirmation cleared the wrong state or was not normalized: quantity='$($crossItem.QuantityText)' status='$($crossItem.Status)'"
}
Write-Host 'PASS 数量人工兜底只在文本真实变化且结果大于0时解除换算阻断'

$isHard = Require-Method $panelType 'IsNameQuotaHardStatus'
if ([bool]$isHard.Invoke($null, @('待确认计数单位1:1'))) { throw 'Soft count confirmation was classified as a hard block' }
foreach ($status in @('缺跨量纲换算系数', '缺当前定额单位', '缺条目', '公式参数缺失或歧义', '未知错误')) {
    if (-not [bool]$isHard.Invoke($null, @($status))) { throw "Hard status was not blocked: $status" }
}

$getColor = Require-Method $panelType 'GetNameQuotaRowBackColor'
$softItem = New-PreviewItem '个' '台' '1' '待确认计数单位1:1' $false 0
$hardItem = New-PreviewItem 'm' 'm2' '1' '缺跨量纲换算系数' $false 2
$exactItem = New-PreviewItem '个' '个' '1' '' $false 0
$previewType.GetField('NeedExactNameConfirmation', $flags).SetValue($exactItem, $true)
$manualItem = New-PreviewItem '个' '' '1' '未匹配' $false 0
$previewType.GetField('NeedManualQuota', $flags).SetValue($manualItem, $true)
$normalItem = New-PreviewItem 'm' 'km' '1/1000' '' $true 1
$yellowArgb = [System.Drawing.Color]::FromArgb(255, 246, 196).ToArgb()
if (($getColor.Invoke($null, @($softItem))).ToArgb() -ne $yellowArgb -or
    ($getColor.Invoke($null, @($exactItem))).ToArgb() -ne $yellowArgb -or
    ($getColor.Invoke($null, @($manualItem))).ToArgb() -ne $yellowArgb) {
    throw 'Soft confirmation rows must be yellow'
}
if (($getColor.Invoke($null, @($hardItem))).ToArgb() -ne [System.Drawing.Color]::MistyRose.ToArgb()) {
    throw 'Hard blockers must be MistyRose'
}
if (($getColor.Invoke($null, @($normalItem))).ToArgb() -ne [System.Drawing.Color]::Empty.ToArgb()) {
    throw 'Safe standard conversion must keep the normal background'
}
Write-Host 'PASS 软确认黄色、硬阻断红色、安全换算正常底色'

$listType = [Collections.Generic.List``1].MakeGenericType($previewType)
$group = [Activator]::CreateInstance($listType)
$safe1 = New-PreviewItem '个' '台' '1' '' $true 0
$safe2 = New-PreviewItem 'm' 'km' '100/1000' '' $true 1
$safe3 = New-PreviewItem 'm' 'm2' '100*0.35' '' $true 2
[void]$group.Add($safe1); [void]$group.Add($safe2); [void]$group.Add($safe3)
$groupSafe = Require-Method $type 'IsNameQuotaGroupSafeForWrite'
$groupArgs = [object[]]::new(1); $groupArgs[0] = $group.PSObject.BaseObject
if (-not [bool]$groupSafe.Invoke($null, $groupArgs)) { throw 'Fully confirmed component should be writable' }
$safe3.Status = '缺跨量纲换算系数'
if ([bool]$groupSafe.Invoke($null, $groupArgs)) { throw 'A hard blocker must reject the whole component' }
$safe3.Status = ''; $safe2.Selected = $false
if ([bool]$groupSafe.Invoke($null, $groupArgs)) { throw 'A partially selected component must reject the whole component' }

$pureAux = [Activator]::CreateInstance($listType)
$zlf = New-PreviewItem '个' '个' '1' '' $true 0
$zlf.TemplateName = '推荐定额'; $zlf.QuotaCode = 'ZLF'; $zlf.ChosenItemName = '安装工程费'
[void]$pureAux.Add($zlf)
$pureAuxArgs = [object[]]::new(1); $pureAuxArgs[0] = $pureAux.PSObject.BaseObject
if ([bool]$groupSafe.Invoke($null, $pureAuxArgs)) { throw 'A SmartFill pure-ZLF component must not pass the write gate' }

$mixed = [Activator]::CreateInstance($listType)
$ordinary = New-PreviewItem '个' '个' '1' '' $true 0
$ordinary.TemplateName = '推荐定额'; $ordinary.QuotaCode = 'EY-299'; $ordinary.ChosenItemName = '安装工程费'
$zlfFollower = New-PreviewItem '个' '个' '1' '' $true 1
$zlfFollower.TemplateName = '推荐定额'; $zlfFollower.QuotaCode = 'ZLF'; $zlfFollower.ChosenItemName = '安装工程费'
[void]$mixed.Add($ordinary); [void]$mixed.Add($zlfFollower)
$mixedArgs = [object[]]::new(1); $mixedArgs[0] = $mixed.PSObject.BaseObject
if (-not [bool]$groupSafe.Invoke($null, $mixedArgs)) { throw 'A SmartFill ordinary-plus-ZLF component should remain writable as a whole' }

$pureSf = [Activator]::CreateInstance($listType)
$sf = New-PreviewItem '项' '项' '1' '' $true 0
$sf.TemplateName = '推荐定额'; $sf.QuotaCode = 'SF'; $sf.ChosenItemName = '设备购置费'
[void]$pureSf.Add($sf)
$pureSfArgs = [object[]]::new(1); $pureSfArgs[0] = $pureSf.PSObject.BaseObject
if (-not [bool]$groupSafe.Invoke($null, $pureSfArgs)) { throw 'A pure SF equipment-purchase component should keep its business exception' }

$mixedSf = [Activator]::CreateInstance($listType)
$mixedOrdinary = New-PreviewItem '个' '台' '1' '' $true 0
$mixedOrdinary.TemplateName = '推荐定额'; $mixedOrdinary.QuotaCode = 'EY-299'; $mixedOrdinary.ChosenItemName = '安装工程费'
$mixedEquipment = New-PreviewItem '项' '项' '1' '' $true 1
$mixedEquipment.TemplateName = '推荐定额'; $mixedEquipment.QuotaCode = 'SF'; $mixedEquipment.ChosenItemName = '设备购置费'
[void]$mixedSf.Add($mixedOrdinary); [void]$mixedSf.Add($mixedEquipment)
$mixedSfArgs = [object[]]::new(1); $mixedSfArgs[0] = $mixedSf.PSObject.BaseObject
if (-not [bool]$groupSafe.Invoke($null, $mixedSfArgs)) { throw 'An ordinary-plus-SF component with target-level entries should remain writable' }
$mixedEquipment.ChosenItemName = '安装工程费'
if ([bool]$groupSafe.Invoke($null, $mixedSfArgs)) { throw 'SF inside a mixed component must still target an equipment-purchase entry' }
$mixedEquipment.ChosenItemName = '设备购置费'; $mixedOrdinary.ChosenItemName = '设备购置费'
if ([bool]$groupSafe.Invoke($null, $mixedSfArgs)) { throw 'A non-SF quota inside a mixed component must not target an equipment-purchase entry' }
Write-Host 'PASS 写入前按组件整组全成或整组不成'

$panelSource = [IO.File]::ReadAllText((Join-Path $repoRoot 'tools\RecoExpandPanel\TemplateFillPanel.cs'), [Text.Encoding]::UTF8)
$featureSource = [IO.File]::ReadAllText((Join-Path $repoRoot 'tools\RecoExpandPanel\TemplateFillFeature.cs'), [Text.Encoding]::UTF8)
if ($panelSource -notmatch 'Cells\["qty"\]\.ReadOnly\s*=\s*false') { throw 'Each component quantity cell must be explicitly editable' }
if ($featureSource -notmatch 'IsNameQuotaGroupSafeForWrite') { throw 'ApplyFill is not wired to the component-level safety gate' }

Write-Host 'Test-SmartFillComponentConversion: PASS'
