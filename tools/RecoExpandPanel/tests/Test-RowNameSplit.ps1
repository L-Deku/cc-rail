param(
    [string]$ExpandDll = $env:RECO_EXPAND_DLL
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..')).Path
if ([string]::IsNullOrWhiteSpace($ExpandDll)) {
    $ExpandDll = Join-Path $repoRoot 'RecoQuotaRecommend\bin\RecoExpandPanel.dll'
}
if (-not (Test-Path -LiteralPath $ExpandDll)) {
    throw "Missing RecoExpandPanel.dll: $ExpandDll"
}

$dllDir = Split-Path -Parent (Resolve-Path -LiteralPath $ExpandDll).Path
foreach ($dependencyName in @(
    'ICSharpCode.SharpZipLib.dll',
    'NPOI.dll',
    'NPOI.OpenXmlFormats.dll',
    'NPOI.OpenXml4Net.dll',
    'NPOI.OOXML.dll'
)) {
    $dependencyPath = Join-Path $dllDir $dependencyName
    if (-not (Test-Path -LiteralPath $dependencyPath)) {
        throw "Missing test dependency: $dependencyPath"
    }
    [void][Reflection.Assembly]::LoadFrom($dependencyPath)
}

$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $ExpandDll).Path)
$type = $assembly.GetType('RecoNet.FormPanel', $true)
$flags = [Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$failures = New-Object 'System.Collections.Generic.List[string]'

function Add-Failure([string]$Message) {
    [void]$failures.Add($Message)
}

function Set-LinkProperty($Link, [string]$Name, $Value) {
    $property = $Link.GetType().GetProperty($Name, $flags)
    if ($null -eq $property) { throw "Missing ExcelQuotaLink property: $Name" }
    $property.SetValue($Link.PSObject.BaseObject, $Value.PSObject.BaseObject, $null)
}

function Set-PreviewField($Item, [string]$Name, $Value) {
    $field = $Item.GetType().GetField($Name, $flags)
    if ($null -eq $field) { throw "Missing FillPreviewItem field: $Name" }
    $field.SetValue($Item.PSObject.BaseObject, $Value.PSObject.BaseObject)
}

# Contract: split structured fragments into main name, context label and unit.
$fragmentType = $type.GetNestedType('RowNameFragment', $flags)
$fragmentListType = [Collections.Generic.List``1].MakeGenericType($fragmentType)
$split = $type.GetMethod('SplitRowNameParts', $flags)
$cases = @(
    [pscustomobject]@{ Name='water'; Unit='m3'; Defs=@([pscustomobject]@{C=1;T='Earthwork';V=$true},[pscustomobject]@{C=2;T='Pond';V=$true},[pscustomobject]@{C=3;T='Pumping';V=$false},[pscustomobject]@{C=4;T='m3';V=$false}); Main='Pumping'; Context='Earthwork Pond'; ExpectedUnit='m3' },
    [pscustomobject]@{ Name='rail'; Unit='pcs'; Defs=@([pscustomobject]@{C=1;T='Stock';V=$true},[pscustomobject]@{C=2;T='Rail joints';V=$false},[pscustomobject]@{C=3;T='pcs';V=$false}); Main='Rail joints'; Context='Stock'; ExpectedUnit='pcs' },
    [pscustomobject]@{ Name='steps'; Unit=''; Defs=@([pscustomobject]@{C=1;T='Roadbed';V=$true},[pscustomobject]@{C=2;T='Cut steps';V=$false}); Main='Cut steps'; Context='Roadbed'; ExpectedUnit='' },
    [pscustomobject]@{ Name='turnout'; Unit='set'; Defs=@([pscustomobject]@{C=1;T='Track';V=$true},[pscustomobject]@{C=2;T='Turnout';V=$true},[pscustomobject]@{C=3;T='P50';V=$false},[pscustomobject]@{C=4;T='No.9';V=$false},[pscustomobject]@{C=5;T='set';V=$false}); Main='P50 No.9'; Context='Track Turnout'; ExpectedUnit='set' },
    [pscustomobject]@{ Name='single'; Unit=''; Defs=@([pscustomobject]@{C=1;T='Independent';V=$false}); Main='Independent'; Context=''; ExpectedUnit='' },
    [pscustomobject]@{ Name='all-vertical'; Unit=''; Defs=@([pscustomobject]@{C=1;T='Left';V=$true},[pscustomobject]@{C=3;T='Right';V=$true},[pscustomobject]@{C=2;T='Middle';V=$true}); Main='Right'; Context='Left Middle'; ExpectedUnit='' },
    [pscustomobject]@{ Name='blank-filter'; Unit='kg'; Defs=@([pscustomobject]@{C=1;T='  ';V=$false},[pscustomobject]@{C=2;T='Steel';V=$false},[pscustomobject]@{C=3;T='kg';V=$false}); Main='Steel'; Context=''; ExpectedUnit='kg' },
    [pscustomobject]@{ Name='top-times'; Unit='顶次'; Defs=@([pscustomobject]@{C=1;T='铺设顶进导轨及抱枕';V=$false},[pscustomobject]@{C=2;T='顶次';V=$false}); Main='铺设顶进导轨及抱枕'; Context=''; ExpectedUnit='顶次' },
    [pscustomobject]@{ Name='empty'; Unit='m'; Defs=@(); Main=''; Context=''; ExpectedUnit='' }
)
foreach ($case in $cases) {
    $list = [Activator]::CreateInstance($fragmentListType)
    foreach ($definition in $case.Defs) {
        $fragment = [Activator]::CreateInstance($fragmentType)
        $fragmentType.GetField('Column', $flags).SetValue($fragment, [int]$definition.C)
        $fragmentType.GetField('Text', $flags).SetValue($fragment, [string]$definition.T)
        $fragmentType.GetField('IsVerticalMerge', $flags).SetValue($fragment, [bool]$definition.V)
        [void]$list.Add($fragment)
    }
    $args2 = [object[]]::new(2)
    $args2[0] = $list.PSObject.BaseObject
    $args2[1] = [string]$case.Unit
    $parts = $split.Invoke($null, $args2.PSObject.BaseObject)
    $partsType = $parts.GetType()
    $main = [string]$partsType.GetField('MainName', $flags).GetValue($parts)
    $context = [string]$partsType.GetField('ContextLabel', $flags).GetValue($parts)
    $unit = [string]$partsType.GetField('Unit', $flags).GetValue($parts)
    if ($main -cne $case.Main -or $context -cne $case.Context -or $unit -cne $case.ExpectedUnit) {
        Add-Failure "split/$($case.Name): main='$main' context='$context' unit='$unit'"
    }
}

$fixturePath = Join-Path ([IO.Path]::GetTempPath()) ('reco-row-name-' + [Guid]::NewGuid().ToString('N') + '.xlsx')
$fixtureBook = $null
$fixtureStage = 'create workbook'
try {
    $fixtureBook = New-Object NPOI.XSSF.UserModel.XSSFWorkbook
    $sheet = $fixtureBook.CreateSheet('Sheet1')
    $row1 = $sheet.CreateRow(0)
    $row2 = $sheet.CreateRow(1)
    $row3 = $sheet.CreateRow(2)
    $row4 = $sheet.CreateRow(3)
    $row5 = $sheet.CreateRow(4)
    $row1.CreateCell(0).SetCellValue('Earthwork')
    $row1.CreateCell(1).SetCellValue('Pond')
    $row2.CreateCell(2).SetCellValue('Pumping')
    $row2.CreateCell(3).SetCellValue('m3')
    $row2.CreateCell(4).SetCellValue([double]1)
    $row3.CreateCell(3).SetCellValue('m3')
    $row3.CreateCell(4).SetCellValue([double]1)
    $row4.CreateCell(2).SetCellValue('m3')
    $row5.CreateCell(1).SetCellValue('Filling')
    $row5.CreateCell(9).SetCellValue([double]47200)
    [void]$sheet.AddMergedRegion([NPOI.SS.Util.CellRangeAddress]::new(0, 1, 0, 0))
    [void]$sheet.AddMergedRegion([NPOI.SS.Util.CellRangeAddress]::new(0, 1, 1, 1))
    [void]$sheet.AddMergedRegion([NPOI.SS.Util.CellRangeAddress]::new(3, 4, 2, 2))
    foreach ($hiddenColumn in 5..8) { $sheet.SetColumnHidden($hiddenColumn, $true) }
    $stream = [IO.File]::Create($fixturePath)
    try { $fixtureBook.Write($stream) } finally { $stream.Dispose() }
    $fixtureBook.Close()
    $fixtureBook = $null

    $linkType = $type.GetNestedType('ExcelQuotaLink', $flags)
    $linkListType = [Collections.Generic.List``1].MakeGenericType($linkType)
    $readLinks = [Activator]::CreateInstance($linkListType)
    $link = [Activator]::CreateInstance($linkType)
    Set-LinkProperty $link 'ExcelPath' $fixturePath
    Set-LinkProperty $link 'WorksheetName' 'Sheet1'
    Set-LinkProperty $link 'CellAddress' 'E2'
    Set-LinkProperty $link 'Expression' 'E2'
    Set-LinkProperty $link 'QuotaCode' 'Q-1'
    [void]$readLinks.Add($link)

    $fixtureStage = 'prepare read context'
    $hashSetType = [Collections.Generic.HashSet``1].MakeGenericType([int])
    $hiddenType = [Collections.Generic.Dictionary``2].MakeGenericType([string], $hashSetType)
    $mergedRegionType = $type.GetNestedType('ExcelMergedRegion', $flags)
    $mergedListType = [Collections.Generic.List``1].MakeGenericType($mergedRegionType)
    $mergedType = [Collections.Generic.Dictionary``2].MakeGenericType([string], $mergedListType)
    $hiddenCache = [Activator]::CreateInstance($hiddenType)
    $mergedCache = [Activator]::CreateInstance($mergedType)
    $addLinks = $type.GetMethod('AddQuantityNameReadLinks', $flags)
    $addArgs = [object[]]::new(6)
    $addArgs[0] = $readLinks.PSObject.BaseObject
    $addArgs[1] = $fixturePath.PSObject.BaseObject
    $addArgs[2] = ([string]'Sheet1').PSObject.BaseObject
    $addArgs[3] = ([string]'E2').PSObject.BaseObject
    $addArgs[4] = $hiddenCache.PSObject.BaseObject
    $addArgs[5] = $mergedCache.PSObject.BaseObject
    [void]$addLinks.Invoke($null, $addArgs.PSObject.BaseObject)

    $contextType = $type.GetNestedType('ExcelSyncReadContext', $flags)
    $contextCtor = @($contextType.GetConstructors($flags))[0]
    $contextArgs = [object[]]::new(1)
    $contextArgs[0] = $readLinks.PSObject.BaseObject
    $readContext = $contextCtor.Invoke($contextArgs.PSObject.BaseObject)

    # P0: template MatchName and target RawName must use the same main-name contract.
    $fixtureStage = 'read template name'
    $readFull = $type.GetMethod('ReadFullNameForCell', $flags)
    $fullArgs = [object[]]::new(6)
    $fullArgs[0] = $fixturePath.PSObject.BaseObject
    $fullArgs[1] = ([string]'Sheet1').PSObject.BaseObject
    $fullArgs[2] = ([string]'E2').PSObject.BaseObject
    $fullArgs[3] = $hiddenCache.PSObject.BaseObject
    $fullArgs[4] = $mergedCache.PSObject.BaseObject
    $fullArgs[5] = $readContext.PSObject.BaseObject
    $templateName = [string]$readFull.Invoke($null, $fullArgs.PSObject.BaseObject)

    $fixtureStage = 'read target rows'
    $readRows = $type.GetMethod('ReadTargetQtyRowsWithChapters', $flags)
    $rowArgs = [object[]]::new(4)
    $rowArgs[0] = $fixturePath.PSObject.BaseObject
    $rowArgs[1] = ([string]'Sheet1').PSObject.BaseObject
    $rowArgs[2] = 5
    $rowArgs[3] = $null
    $targetRows = $readRows.Invoke($null, $rowArgs.PSObject.BaseObject)
    if ($targetRows.Count -ne 1) {
        Add-Failure "target rows: expected 1, actual $($targetRows.Count)"
    } else {
        $target = $targetRows[0]
        $targetType = $target.GetType()
        $targetName = [string]$targetType.GetField('RawName', $flags).GetValue($target)
        $targetChapter = [string]$targetType.GetField('Chapter', $flags).GetValue($target)
        if ($templateName -cne $targetName -or $targetName -cne 'Pumping') {
            Add-Failure "P0 template/target mismatch: template='$templateName' target='$targetName'"
        }
        $chapterByRow = $rowArgs[3]
        $mappedChapter = [string]$chapterByRow[2]
        if ($mappedChapter -cne $targetChapter) {
            Add-Failure "P1 chapter mismatch: map='$mappedChapter' row='$targetChapter'"
        }
    }

    # 单位可以物理相距超过6列；中间隐藏列不计入6个可见列，并从纵向合并区域锚点回填。
    $fixtureStage = 'read distant merged unit'
    $wideArgs = [object[]]::new(4)
    $wideArgs[0] = $fixturePath.PSObject.BaseObject
    $wideArgs[1] = ([string]'Sheet1').PSObject.BaseObject
    $wideArgs[2] = 10
    $wideArgs[3] = $null
    $wideRows = $readRows.Invoke($null, $wideArgs.PSObject.BaseObject)
    if ($wideRows.Count -ne 1) {
        Add-Failure "merged unit target rows: expected 1, actual $($wideRows.Count)"
    } else {
        $wideTarget = $wideRows[0]
        $wideType = $wideTarget.GetType()
        $wideName = [string]$wideType.GetField('RawName', $flags).GetValue($wideTarget)
        $wideUnit = [string]$wideType.GetField('Unit', $flags).GetValue($wideTarget)
        if ($wideName -cne 'Filling' -or $wideUnit -cne 'm3') {
            Add-Failure "distant merged unit: name='$wideName' unit='$wideUnit'"
        }
    }

    # P1-3: all normal binding fallbacks must already be main names.
    $fixtureStage = 'finalize binding names'
    $links = [Activator]::CreateInstance($linkListType)
    [void]$links.Add($link)
    $finalize = $type.GetMethod('FinalizeBoundLinkNames', $flags)
    $finalizeArgs = [object[]]::new(2)
    $finalizeArgs[0] = $links.PSObject.BaseObject
    $finalizeArgs[1] = $false
    $finalizedNames = $finalize.Invoke($null, $finalizeArgs.PSObject.BaseObject)
    $fallbackName = [string]$finalizedNames[$link]
    if ($fallbackName -cne 'Pumping') {
        Add-Failure "P1 fallback producer returned '$fallbackName'"
    }

    # Characterization guard: when the only fragment is the unit, fallback naming must retain that unit.
    $fixtureStage = 'preserve fallback unit'
    $unitLink = [Activator]::CreateInstance($linkType)
    Set-LinkProperty $unitLink 'ExcelPath' $fixturePath
    Set-LinkProperty $unitLink 'WorksheetName' 'Sheet1'
    Set-LinkProperty $unitLink 'CellAddress' 'E3'
    Set-LinkProperty $unitLink 'Expression' 'E3'
    Set-LinkProperty $unitLink 'QuotaCode' 'Q-2'
    $nameDictionaryType = [Collections.Generic.Dictionary``2].MakeGenericType($linkType, [string])
    $unitFallbackNames = [Activator]::CreateInstance($nameDictionaryType)
    $unitFallbackNames.Add($unitLink, 'Fallback pumping')
    $buildGroups = $type.GetMethod('BuildBindingFeedbackGroups', $flags)
    $unitGroupArgs = [object[]]::new(1)
    $unitGroupArgs[0] = $unitFallbackNames.PSObject.BaseObject
    $unitGroups = $buildGroups.Invoke($null, $unitGroupArgs.PSObject.BaseObject)
    if ($unitGroups.Count -ne 1) {
        Add-Failure "E1 fallback unit expected 1 group, actual $($unitGroups.Count)"
    } else {
        $unitGroupType = $unitGroups[0].GetType()
        $unitGroupName = [string]$unitGroupType.GetField('QuantityName', $flags).GetValue($unitGroups[0])
        $unitGroupUnit = [string]$unitGroupType.GetField('QuantityUnit', $flags).GetValue($unitGroups[0])
        if ($unitGroupName -cne 'Fallback pumping' -or $unitGroupUnit -cne 'm3') {
            Add-Failure "E1 fallback unit returned name='$unitGroupName' unit='$unitGroupUnit'"
        }
    }

    # E3: learning feedback accepts only the full main-name key and must not strip it again.
    $fixtureStage = 'validate template feedback name key'
    $previewType = $type.GetNestedType('FillPreviewItem', $flags)
    $previewListType = [Collections.Generic.List``1].MakeGenericType($previewType)
    $buildTemplateFeedback = $type.GetMethod('BuildTemplateRightClickFeedbackGroup', $flags)

    $displayOnlyItem = [Activator]::CreateInstance($previewType)
    Set-PreviewField $displayOnlyItem 'IsNameDriven' $true
    Set-PreviewField $displayOnlyItem 'QuotaCode' 'Q-3'
    Set-PreviewField $displayOnlyItem 'TargetName' 'Truncated display name'
    Set-PreviewField $displayOnlyItem 'TargetFullName' ''
    Set-PreviewField $displayOnlyItem 'TargetUnit' 'm3'
    $displayOnlyItems = [Activator]::CreateInstance($previewListType)
    [void]$displayOnlyItems.Add($displayOnlyItem)
    $displayOnlyArgs = [object[]]::new(8)
    $displayOnlyArgs[0] = $displayOnlyItems.PSObject.BaseObject
    $displayOnlyArgs[1] = ''
    $displayOnlyArgs[2] = ''
    $displayOnlyArgs[3] = $null
    $displayOnlyArgs[4] = 1
    $displayOnlyArgs[5] = 0
    $displayOnlyArgs[6] = 0
    $displayOnlyArgs[7] = 'accepted'
    $displayOnlyGroup = $buildTemplateFeedback.Invoke($null, $displayOnlyArgs.PSObject.BaseObject)
    if ($null -ne $displayOnlyGroup) {
        Add-Failure 'E3 template feedback accepted TargetName without TargetFullName.'
    }

    $fullNameItem = [Activator]::CreateInstance($previewType)
    Set-PreviewField $fullNameItem 'IsNameDriven' $true
    Set-PreviewField $fullNameItem 'QuotaCode' 'Q-4'
    Set-PreviewField $fullNameItem 'TargetName' 'Display value'
    Set-PreviewField $fullNameItem 'TargetFullName' 'Pumping m3'
    Set-PreviewField $fullNameItem 'TargetUnit' 'm3'
    $fullNameItems = [Activator]::CreateInstance($previewListType)
    [void]$fullNameItems.Add($fullNameItem)
    $fullNameArgs = [object[]]::new(8)
    $fullNameArgs[0] = $fullNameItems.PSObject.BaseObject
    $fullNameArgs[1] = ''
    $fullNameArgs[2] = ''
    $fullNameArgs[3] = $null
    $fullNameArgs[4] = 1
    $fullNameArgs[5] = 0
    $fullNameArgs[6] = 0
    $fullNameArgs[7] = 'accepted'
    $fullNameGroup = $buildTemplateFeedback.Invoke($null, $fullNameArgs.PSObject.BaseObject)
    if ($null -eq $fullNameGroup) {
        Add-Failure 'E3 template feedback rejected a non-empty TargetFullName.'
    } else {
        $feedbackName = [string]$fullNameGroup.GetType().GetField('QuantityName', $flags).GetValue($fullNameGroup)
        if ($feedbackName -cne 'Pumping m3') {
            Add-Failure "E3 template feedback stripped TargetFullName to '$feedbackName'"
        }
    }

    # Force the consumer down its fallback path by removing the workbook after finalization.
    $fixtureStage = 'consume binding fallback'
    Remove-Item -LiteralPath $fixturePath -Force
    $groupArgs = [object[]]::new(1)
    $groupArgs[0] = $finalizedNames.PSObject.BaseObject
    $groups = $buildGroups.Invoke($null, $groupArgs.PSObject.BaseObject)
    if ($groups.Count -ne 1) {
        Add-Failure "P1 fallback consumer expected 1 group, actual $($groups.Count)"
    } else {
        $groupName = [string]$groups[0].GetType().GetField('QuantityName', $flags).GetValue($groups[0])
        if ($groupName -cne 'Pumping') {
            Add-Failure "P1 fallback consumer returned '$groupName'"
        }
    }
}
catch {
    Add-Failure ("fixture/reflection at $fixtureStage`: " + $_.Exception.Message)
}
finally {
    if ($null -ne $fixtureBook) { try { $fixtureBook.Close() } catch { } }
    if (Test-Path -LiteralPath $fixturePath) { Remove-Item -LiteralPath $fixturePath -Force }
}

# Preserve the legacy display/full reader byte-for-byte on the recorded 31-row workbook baseline.
$baselineChecked = 0
$baselinePath = Join-Path $repoRoot 'artifacts\reco2-nameparts-readrow-baseline-20260828.tsv'
if (Test-Path -LiteralPath $baselinePath) {
    $baselineLines = Get-Content -LiteralPath $baselinePath -Encoding UTF8
    $baselineWorkbook = ([string]$baselineLines[0]).Substring(([string]$baselineLines[0]).IndexOf("`t") + 1)
    $baselineSheet = ([string]$baselineLines[1]).Substring(([string]$baselineLines[1]).IndexOf("`t") + 1)
    if (Test-Path -LiteralPath $baselineWorkbook) {
        try {
            $linkType = $type.GetNestedType('ExcelQuotaLink', $flags)
            $linkListType = [Collections.Generic.List``1].MakeGenericType($linkType)
            $baselineLinks = [Activator]::CreateInstance($linkListType)
            $hashSetType = [Collections.Generic.HashSet``1].MakeGenericType([int])
            $hiddenType = [Collections.Generic.Dictionary``2].MakeGenericType([string], $hashSetType)
            $mergedRegionType = $type.GetNestedType('ExcelMergedRegion', $flags)
            $mergedListType = [Collections.Generic.List``1].MakeGenericType($mergedRegionType)
            $mergedType = [Collections.Generic.Dictionary``2].MakeGenericType([string], $mergedListType)
            $baselineHidden = [Activator]::CreateInstance($hiddenType)
            $baselineMerged = [Activator]::CreateInstance($mergedType)
            $addLinks = $type.GetMethod('AddQuantityNameReadLinks', $flags)
            $baselineRows = @($baselineLines | Select-Object -Skip 4 | ConvertFrom-Csv -Delimiter "`t")
            foreach ($expected in $baselineRows) {
                $address = [string]$expected.address
                $baselineLink = [Activator]::CreateInstance($linkType)
                Set-LinkProperty $baselineLink 'ExcelPath' $baselineWorkbook
                Set-LinkProperty $baselineLink 'WorksheetName' $baselineSheet
                Set-LinkProperty $baselineLink 'CellAddress' $address
                Set-LinkProperty $baselineLink 'Expression' $address
                [void]$baselineLinks.Add($baselineLink)
                $addArgs = [object[]]::new(6)
                $addArgs[0] = $baselineLinks.PSObject.BaseObject
                $addArgs[1] = $baselineWorkbook.PSObject.BaseObject
                $addArgs[2] = $baselineSheet.PSObject.BaseObject
                $addArgs[3] = $address.PSObject.BaseObject
                $addArgs[4] = $baselineHidden.PSObject.BaseObject
                $addArgs[5] = $baselineMerged.PSObject.BaseObject
                [void]$addLinks.Invoke($null, $addArgs.PSObject.BaseObject)
            }
            $contextType = $type.GetNestedType('ExcelSyncReadContext', $flags)
            $contextCtor = @($contextType.GetConstructors($flags))[0]
            $contextArgs = [object[]]::new(1)
            $contextArgs[0] = $baselineLinks.PSObject.BaseObject
            $baselineContext = $contextCtor.Invoke($contextArgs.PSObject.BaseObject)
            $readName = @($type.GetMethods($flags) | Where-Object { $_.Name -eq 'ReadRowNameAt' -and $_.GetParameters().Count -eq 7 })[0]
            foreach ($expected in $baselineRows) {
                $address = [string]$expected.address
                $shortArgs = [object[]]@($baselineWorkbook, $baselineSheet, $address,
                    $baselineHidden.PSObject.BaseObject, $baselineMerged.PSObject.BaseObject,
                    $baselineContext.PSObject.BaseObject, $false)
                $fullArgs = [object[]]@($baselineWorkbook, $baselineSheet, $address,
                    $baselineHidden.PSObject.BaseObject, $baselineMerged.PSObject.BaseObject,
                    $baselineContext.PSObject.BaseObject, $true)
                $actualShort = [string]$readName.Invoke($null, $shortArgs.PSObject.BaseObject)
                $actualFull = [string]$readName.Invoke($null, $fullArgs.PSObject.BaseObject)
                if ($actualShort -cne [string]$expected.short -or $actualFull -cne [string]$expected.full) {
                    Add-Failure "baseline/${address}: short='$actualShort' full='$actualFull'"
                }
            }
            $baselineChecked = $baselineRows.Count
        }
        catch {
            Add-Failure ('31-row baseline: ' + $_.Exception.Message)
        }
    }
}

# P1-2/P1-3 static guards cover the AutoMatch entry that cannot be invoked without a business DB.
$templateSource = [IO.File]::ReadAllText((Join-Path $repoRoot 'tools\RecoExpandPanel\TemplateFillFeature.cs'), [Text.Encoding]::UTF8)
$nameMatchSource = [IO.File]::ReadAllText((Join-Path $repoRoot 'tools\RecoExpandPanel\TemplateFillNameMatch.cs'), [Text.Encoding]::UTF8)
$excelLinkSource = [IO.File]::ReadAllText((Join-Path $repoRoot 'tools\RecoExpandPanel\ExcelLinkFeature.cs'), [Text.Encoding]::UTF8)
if ($templateSource.Contains('public string TargetMainName;') -or $nameMatchSource.Contains('.TargetMainName')) {
    Add-Failure 'P1 duplicate TargetMainName state still exists.'
}
if ($excelLinkSource.Contains('savedQuantityNames[item.Link] = item.QuantityName ?? "";')) {
    Add-Failure 'P1 AutoMatch still sends item.QuantityName directly to learning.'
}
if (-not $excelLinkSource.Contains('FinalizeBoundLinkNames(pendingLinks, false)')) {
    Add-Failure 'P1 AutoMatch does not finalize main names before learning.'
}

if ($failures.Count -ne 0) {
    throw ("Row-name regression failures ($($failures.Count)):`n - " + ($failures -join "`n - "))
}

Write-Host "PASS: 8 split cases, $baselineChecked baseline rows, template/target parity, chapter parity, safe binding fallback, AutoMatch guard." -ForegroundColor Green
