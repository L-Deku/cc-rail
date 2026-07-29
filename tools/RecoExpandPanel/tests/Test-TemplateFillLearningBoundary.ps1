$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..')).Path
$dll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL)) {
    $env:RECO_EXPAND_DLL
} else {
    Join-Path $repoRoot 'RecoQuotaRecommend\bin\RecoExpandPanel.dll'
}
if (-not (Test-Path -LiteralPath $dll)) { throw "Missing test DLL: $dll" }

$dllDir = Split-Path -Parent $dll
foreach ($dependency in @('NPOI.dll', 'NPOI.OpenXmlFormats.dll', 'NPOI.OpenXml4Net.dll', 'NPOI.OOXML.dll')) {
    [void][System.Reflection.Assembly]::LoadFrom((Join-Path $dllDir $dependency))
}

$flags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$type = [System.Reflection.Assembly]::LoadFrom($dll).GetType('RecoNet.FormPanel')
$templateType = $type.GetNestedType('FillTemplate', $flags)
$rowType = $type.GetNestedType('FillTemplateRow', $flags)
$itemType = $type.GetNestedType('FillPreviewItem', $flags)
$itemListType = [Collections.Generic.List``1].MakeGenericType($itemType)

function New-TemplateRow(
    [string]$name,
    [string]$code,
    [string]$itemNo,
    [string]$unit,
    [string]$sourceExpr,
    [string]$origin,
    [string]$chapter
) {
    $row = [Activator]::CreateInstance($rowType)
    $row.MatchName = $name
    $row.SourceName = $name
    $row.QuotaCode = $code
    $row.ItemNo = $itemNo
    $row.Unit = $unit
    $row.SourceExpr = $sourceExpr
    $row.Origin = $origin
    $row.MatchChapter = $chapter
    return $row
}

function New-PreviewItem([string]$name, [string]$code, [string]$itemNo, [string]$unit, [int]$targetRow) {
    $item = [Activator]::CreateInstance($itemType)
    $item.IsNameDriven = $true
    $item.TargetFullName = $name
    $item.TargetName = $name
    $item.TargetUnit = $unit
    $item.QuotaCode = $code
    $item.ItemNo = $itemNo
    $item.Unit = $unit
    $item.TargetRow = $targetRow
    $item.SourceName = $code
    return $item
}

$normalizeLegacy = $type.GetMethod('NormalizeLegacyFillTemplateRows', $flags)
$legacy = [Activator]::CreateInstance($templateType)
$legacy.MatchBy = 'name'
$legacy.Rows.Add((New-TemplateRow 'Mud transport' 'Q-A' '0101' 'm3' 'D2' '' 'Chapter 1'))
$legacy.Rows.Add((New-TemplateRow 'Mud transport' 'Q-A' '0101' 'm3' '' '' 'Chapter 1'))
$legacy.Rows.Add((New-TemplateRow 'Mud transport' 'Q-B' '0101' 'm3' '' '' 'Chapter 1'))
$removed = $normalizeLegacy.Invoke($null, @($legacy))
if ($removed -ne 1 -or $legacy.Rows.Count -ne 2) { throw 'Legacy duplicate cleanup failed.' }
if ($legacy.Rows[0].Origin -ne 'generated' -or $legacy.Rows[1].Origin -ne 'manual') {
    throw 'Legacy origin classification failed.'
}
Write-Host 'PASS legacy template rows are conservatively classified and deduplicated'

$existing = [Activator]::CreateInstance($templateType)
$existing.MatchBy = 'name'
$existing.Rows.Add((New-TemplateRow 'Mud transport' 'Q-B' '0101' 'm3' '' 'manual' 'Chapter 1'))
$regenerated = [Activator]::CreateInstance($templateType)
$regenerated.MatchBy = 'name'
$regenerated.Rows.Add((New-TemplateRow 'Mud transport' 'Q-A' '0101' 'm3' 'D2' 'generated' 'Chapter 1'))
$regenerated.Rows.Add((New-TemplateRow 'Pipe install m' 'Q-C' '0102' 'm' 'D3' 'generated' 'Chapter 1'))
$merge = $type.GetMethod('MergeRegeneratedFillTemplate', $flags)
$merged = $merge.Invoke($null, @($existing, $regenerated))
$mergedCodes = @($merged.Rows | ForEach-Object { $_.QuotaCode })
if (($mergedCodes -join ',') -ne 'Q-C,Q-B') { throw "Manual override did not win regeneration: $($mergedCodes -join ',')" }
Write-Host 'PASS regeneration replaces generated rows and preserves manual overrides'

$oldItems = [Activator]::CreateInstance($itemListType)
$sameItems = [Activator]::CreateInstance($itemListType)
$differentItems = [Activator]::CreateInstance($itemListType)
$oldItems.Add((New-PreviewItem 'Mud transport' 'Q-A' '0101' 'm3' 2))
$sameItems.Add((New-PreviewItem 'Mud transport' 'Q-A' '0101' 'm3' 2))
$differentItems.Add((New-PreviewItem 'Mud transport' 'Q-B' '0101' 'm3' 2))
$equivalent = $type.GetMethod('AreEquivalentNameBindingGroups', $flags)
if (-not $equivalent.Invoke($null, [object[]]@($oldItems.PSObject.BaseObject, $sameItems.PSObject.BaseObject))) {
    throw 'Equivalent right-click binding was not detected.'
}
if ($equivalent.Invoke($null, [object[]]@($oldItems.PSObject.BaseObject, $differentItems.PSObject.BaseObject))) {
    throw 'Changed right-click binding was treated as equivalent.'
}
Write-Host 'PASS identical right-click binding is idempotent and changed target set is detected'

$templateForReplace = [Activator]::CreateInstance($templateType)
$templateForReplace.MatchBy = 'name'
$templateForReplace.Rows.Add((New-TemplateRow 'Mud transport' 'Q-A' '0101' 'm3' 'D2' 'generated' 'Chapter 1'))
$replaceManual = $type.GetMethod('ReplaceTemplateWithManualBinding', $flags)
$replaceManual.Invoke($null, [object[]]@(
    $templateForReplace,
    $differentItems.PSObject.BaseObject,
    $oldItems.PSObject.BaseObject
))
if ($templateForReplace.Rows.Count -ne 1 -or $templateForReplace.Rows[0].QuotaCode -ne 'Q-B' -or
    $templateForReplace.Rows[0].Origin -ne 'manual' -or -not [String]::IsNullOrEmpty($templateForReplace.Rows[0].SourceExpr)) {
    throw 'Manual right-click replacement did not atomically replace the template group.'
}
Write-Host 'PASS manual right-click replacement becomes the authoritative name-only template group'

$fixture = Join-Path ([IO.Path]::GetTempPath()) ("reco-template-boundary-" + [Guid]::NewGuid().ToString('N') + '.xlsx')
$book = New-Object NPOI.XSSF.UserModel.XSSFWorkbook
try {
    $sheet = $book.CreateSheet('Sheet1')
    [void]$sheet.CreateRow(0)
    $excelRow = $sheet.CreateRow(1)
    $excelRow.CreateCell(3).SetCellValue([double]12.5)
    $stream = [IO.File]::Create($fixture)
    try { $book.Write($stream) } finally { $stream.Dispose() }

    $columnTemplate = [Activator]::CreateInstance($templateType)
    $columnTemplate.MatchBy = 'position'
    $columnTemplate.Rows.Add((New-TemplateRow 'Manual only' 'Q-M' '0101' 'm3' '' 'manual' 'Chapter 1'))
    $columnTemplate.Rows.Add((New-TemplateRow 'Generated' 'Q-G' '0101' 'm3' 'D2' 'generated' 'Chapter 1'))
    $buildColumn = $type.GetMethod('BuildPreview_ColumnAnchor', $flags)
    $columnArgs = [object[]]::new(4)
    $columnArgs[0] = $columnTemplate.PSObject.BaseObject
    $columnArgs[1] = ([string]$fixture).PSObject.BaseObject
    $columnArgs[2] = ([string]'Sheet1').PSObject.BaseObject
    $columnArgs[3] = ([string]'D').PSObject.BaseObject
    $columnPreview = $buildColumn.Invoke($null, $columnArgs.PSObject.BaseObject)
    if ($columnPreview.Count -ne 1 -or $columnPreview[0].QuotaCode -ne 'Q-G' -or $columnPreview[0].QuantityText -ne '12.5') {
        throw 'Column-anchor preview did not ignore name-only manual rows.'
    }
    Write-Host 'PASS column-anchor preview ignores name-only manual rows'
}
finally {
    $book.Close()
    if (Test-Path -LiteralPath $fixture) { Remove-Item -LiteralPath $fixture -Force }
}

$groupType = $type.GetNestedType('MappingFeedbackGroup', $flags)
$targetType = $type.GetNestedType('MappingFeedbackTarget', $flags)
$upsert = $type.GetMethod('UpsertMappingBoxGroup', $flags)
$rows = [Activator]::CreateInstance($upsert.GetParameters()[0].ParameterType)
function New-FeedbackGroup([string]$code, [int]$accepted, [int]$corrected, [int]$rejected) {
    $group = [Activator]::CreateInstance($groupType)
    $group.QuantityName = 'Mud transport'
    $group.QuantityUnit = 'm3'
    $group.Method = '2024'
    $group.AcceptedCount = $accepted
    $group.CorrectedCount = $corrected
    $group.RejectedCount = $rejected
    $target = [Activator]::CreateInstance($targetType)
    $target.Kind = 'quota'
    $target.Code = $code
    $group.Targets.Add($target)
    return $group
}
$upsert.Invoke($null, [object[]]@($rows.PSObject.BaseObject, (New-FeedbackGroup 'Q-A' 1 0 0)))
$upsert.Invoke($null, [object[]]@($rows.PSObject.BaseObject, (New-FeedbackGroup 'Q-A' 0 0 1)))
$upsert.Invoke($null, [object[]]@($rows.PSObject.BaseObject, (New-FeedbackGroup 'Q-B' 0 1 0)))
$oldRow = @($rows | Where-Object { $_['target_code'] -eq 'Q-A' })[0]
$newRow = @($rows | Where-Object { $_['target_code'] -eq 'Q-B' })[0]
if ($oldRow['accepted_count'] -ne '1' -or $oldRow['rejected_count'] -ne '1' -or $oldRow['weight'] -ne '0') {
    throw 'Rejected old relation did not lose local weight.'
}
if ($newRow['accepted_count'] -ne '0' -or $newRow['corrected_count'] -ne '1' -or $newRow['weight'] -ne '20') {
    throw 'Corrected new relation did not gain local weight.'
}
$buildBoxIndex = $type.GetMethod('BuildMappingBoxIndex', $flags)
$boxIndexArgs = [object[]]::new(1)
$boxIndexArgs[0] = $rows.PSObject.BaseObject
$boxIndex = $buildBoxIndex.Invoke($null, $boxIndexArgs.PSObject.BaseObject)
$lookupBox = $type.GetMethod('LookupMappingBox', $flags)
$lookupArgs = [object[]]::new(2)
$lookupArgs[0] = ([string]'Mud transport').PSObject.BaseObject
$lookupArgs[1] = $boxIndex.PSObject.BaseObject
$matches = $lookupBox.Invoke($null, $lookupArgs.PSObject.BaseObject)
if ($matches.Count -ne 1 -or $matches[0].Targets.Count -ne 1 -or $matches[0].Targets[0].QuotaCode -ne 'Q-B') {
    throw 'Rejected local relation remained visible to name-driven fallback.'
}
Write-Host 'PASS local learning applies correction and rejection deltas'

$applySource = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\TemplateFillFeature.cs') -Raw
if ($applySource -match 'FeedbackNameMatches\s*\(') {
    throw 'Direct template apply still calls learning feedback.'
}
$excelLinkSource = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\ExcelLinkFeature.cs') -Raw
if ($excelLinkSource -notmatch 'template-right-click') { throw 'Right-click learning source marker is missing.' }
$learningSource = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\LearningDbFeature.cs') -Raw
foreach ($required in @('accepted_count=accepted_count+@accepted', 'corrected_count=corrected_count+@corrected',
    'rejected_count=rejected_count+@rejected')) {
    if ($learningSource -notlike "*$required*") { throw "SQL learning delta is missing: $required" }
}
Write-Host 'PASS direct apply has no learning side effect and SQL deltas are explicit'
Write-Host 'ALL PASS'
