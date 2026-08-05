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

$targetEntryRows = [Activator]::CreateInstance($upsert.GetParameters()[0].ParameterType)
$targetEntryGroup = New-FeedbackGroup 'Q-TARGET' 1 0 0
$groupType.GetField('EntryCode', $flags).SetValue($targetEntryGroup, '0101-LEGACY')
$groupType.GetField('EntryName', $flags).SetValue($targetEntryGroup, 'legacy group entry')
$targetEntryTarget = $groupType.GetField('Targets', $flags).GetValue($targetEntryGroup)[0]
$targetType.GetField('EntryCode', $flags).SetValue($targetEntryTarget, '0202-TARGET')
$targetType.GetField('EntryName', $flags).SetValue($targetEntryTarget, 'target entry')
$upsert.Invoke($null, [object[]]@($targetEntryRows.PSObject.BaseObject, $targetEntryGroup.PSObject.BaseObject))
$targetEntryRow = @($targetEntryRows)[0]
if ($targetEntryRow['entry_code'] -ne '0202-TARGET' -or $targetEntryRow['entry_name'] -ne 'target entry') {
    throw 'mapping-boxes persistence used legacy group entry instead of target-level entry.'
}
Write-Host 'PASS mapping-boxes persists each target entry before the legacy group fallback'

$collectAccepted = $type.GetMethod('CollectFullyWrittenNameDrivenGroups', $flags)
$shouldRecordLocal = $type.GetMethod('ShouldRecordAcceptedFillGroupLocally', $flags)
$shouldRetryRemote = $type.GetMethod('ShouldRetryAcceptedFillGroupRemotely', $flags)
$buildAccepted = $type.GetMethod('BuildTemplateRightClickFeedbackGroup', $flags)
if ($null -eq $collectAccepted -or $null -eq $shouldRecordLocal -or
    $null -eq $shouldRetryRemote -or $null -eq $buildAccepted) {
    throw 'Accepted apply learning behavior entry points are missing.'
}
$acceptedItems = [Activator]::CreateInstance($itemListType)
foreach ($code in @('Q-1', 'Q-2', 'Q-3')) {
    $acceptedItem = New-PreviewItem 'Mud transport accepted' $code '0101' 'm3' 20
    $acceptedItem.Selected = $true
    $acceptedItems.Add($acceptedItem)
}
for ($acceptedIndex = 0; $acceptedIndex -lt $acceptedItems.Count; $acceptedIndex++) {
    $acceptedItems[$acceptedIndex].ChosenItemNo = '02-0' + ($acceptedIndex + 1)
    $acceptedItems[$acceptedIndex].ChosenItemName = 'target entry ' + ($acceptedIndex + 1)
}
$writtenSetType = [Collections.Generic.HashSet``1].MakeGenericType($itemType)
$writtenItems = [Activator]::CreateInstance($writtenSetType)
[void]$writtenItems.Add($acceptedItems[0])
[void]$writtenItems.Add($acceptedItems[1])
$collectArgs = [object[]]::new(2)
$collectArgs[0] = $acceptedItems.PSObject.BaseObject
$collectArgs[1] = $writtenItems.PSObject.BaseObject
$partialGroups = $collectAccepted.Invoke($null, $collectArgs)
if ($partialGroups.Count -ne 0) { throw 'A three-target group with only two successful writes produced accepted feedback.' }
[void]$writtenItems.Add($acceptedItems[2])
$fullGroups = $collectAccepted.Invoke($null, $collectArgs)
if ($fullGroups.Count -ne 1 -or $fullGroups[0].Count -ne 3) {
    throw 'A fully written three-target group did not produce exactly one accepted group.'
}
$acceptedItems[1].Selected = $false
if ($collectAccepted.Invoke($null, $collectArgs).Count -ne 0) {
    throw 'An unselected target inside a fully written group produced accepted feedback.'
}
$acceptedItems[1].Selected = $true
$fullGroups = $collectAccepted.Invoke($null, $collectArgs)
$buildArgs = [object[]]::new(8)
$buildArgs[0] = $fullGroups[0]
$buildArgs[1] = 'accepted.xlsx'
$buildArgs[2] = 'Sheet1'
$buildArgs[3] = $null
$buildArgs[4] = 1
$buildArgs[5] = 0
$buildArgs[6] = 0
$buildArgs[7] = 'accepted'
$acceptedFeedback = $buildAccepted.Invoke($null, $buildArgs)
if ($null -eq $acceptedFeedback -or $acceptedFeedback.AcceptedCount -ne 1 -or
    $acceptedFeedback.Targets.Count -ne 3 -or $acceptedFeedback.Workbook -ne 'accepted.xlsx') {
    throw 'A fully written group did not build one complete accepted mapping group.'
}
for ($acceptedIndex = 0; $acceptedIndex -lt $acceptedFeedback.Targets.Count; $acceptedIndex++) {
    if ($acceptedFeedback.Targets[$acceptedIndex].EntryCode -ne ('02-0' + ($acceptedIndex + 1)) -or
        $acceptedFeedback.Targets[$acceptedIndex].EntryName -ne ('target entry ' + ($acceptedIndex + 1))) {
        throw 'Accepted feedback collapsed target-level entries back to the group entry.'
    }
}
$anchorItem = New-PreviewItem 'Column anchor' 'Q-A' '0101' 'm3' 21
$anchorItem.IsNameDriven = $false
$anchorItems = [Activator]::CreateInstance($itemListType)
[void]$anchorItems.Add($anchorItem)
$anchorWritten = [Activator]::CreateInstance($writtenSetType)
[void]$anchorWritten.Add($anchorItem)
$anchorArgs = [object[]]::new(2)
$anchorArgs[0] = $anchorItems.PSObject.BaseObject
$anchorArgs[1] = $anchorWritten.PSObject.BaseObject
if ($collectAccepted.Invoke($null, $anchorArgs).Count -ne 0) { throw 'Column-anchor apply produced accepted feedback.' }

$gateArgs = [object[]]::new(1)
$gateArgs[0] = $fullGroups[0]
if (-not [bool]$shouldRecordLocal.Invoke($null, $gateArgs)) { throw 'Fresh accepted group did not request local feedback.' }
foreach ($item in $acceptedItems) { $item.LocalFeedbackRecorded = $true; $item.RemoteFeedbackDurable = $false }
if ([bool]$shouldRecordLocal.Invoke($null, $gateArgs)) { throw 'Repeated apply would add a second local vote after remote failure.' }
if (-not [bool]$shouldRetryRemote.Invoke($null, $gateArgs)) { throw 'Remote-only retry was not preserved after local feedback.' }
foreach ($item in $acceptedItems) { $item.RemoteFeedbackDurable = $true }
if ([bool]$shouldRetryRemote.Invoke($null, $gateArgs)) { throw 'Durable accepted feedback was scheduled for another remote retry.' }
Write-Host 'PASS accepted apply learns only complete name-driven groups and keeps local/remote idempotence separate'

$batchType = $type.GetNestedType('LearningDbOutboxBatch', $flags)
$batch = [Activator]::CreateInstance($batchType, $true).PSObject.BaseObject
$batchType.GetField('BatchId', $flags).SetValue($batch, ('a' * 32))
$batchType.GetField('Source', $flags).SetValue($batch, 'apply-accept')
[void]$batchType.GetField('Groups', $flags).GetValue($batch).Add($acceptedFeedback)
$serializeBatch = $type.GetMethod('SerializeLearningDbOutboxBatch', $flags)
$toFlatJson = $type.GetMethod('ToFlatJson', $flags)
$parseBatch = $type.GetMethod('ParseLearningDbOutboxBatch', $flags)
$serializeArgs = New-Object 'object[]' 1; $serializeArgs[0] = $batch
$flatBatch = $serializeBatch.Invoke($null, $serializeArgs)
$jsonArgs = New-Object 'object[]' 1; $jsonArgs[0] = $flatBatch.PSObject.BaseObject
$outboxLine = [string]$toFlatJson.Invoke($null, $jsonArgs)
$parseArgs = New-Object 'object[]' 1; $parseArgs[0] = $outboxLine
$parsedBatch = $parseBatch.Invoke($null, $parseArgs)
$parsedGroup = $batchType.GetField('Groups', $flags).GetValue($parsedBatch)[0]
$parsedTargets = $groupType.GetField('Targets', $flags).GetValue($parsedGroup)
if ($parsedTargets.Count -ne 3 -or $parsedTargets[0].EntryCode -ne '02-01' -or
    $parsedTargets[2].EntryCode -ne '02-03' -or $parsedTargets[2].EntryName -ne 'target entry 3') {
    throw 'Outbox serialize/parse roundtrip lost target-level entries.'
}
Write-Host 'PASS Outbox roundtrip preserves target-level entries without database access'

$applySource = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\TemplateFillFeature.cs') -Raw
if ($applySource -notmatch 'apply-accept' -or $applySource -notmatch 'CollectFullyWrittenNameDrivenGroups') {
    throw 'Direct apply is missing complete-group accepted feedback.'
}
$excelLinkSource = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\ExcelLinkFeature.cs') -Raw
if ($excelLinkSource -notmatch 'template-right-click') { throw 'Right-click learning source marker is missing.' }
$learningSource = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\LearningDbFeature.cs') -Raw
foreach ($required in @('accepted_count=accepted_count+@accepted', 'corrected_count=corrected_count+@corrected',
    'rejected_count=rejected_count+@rejected')) {
    if ($learningSource -notlike "*$required*") { throw "SQL learning delta is missing: $required" }
}
Write-Host 'PASS direct apply records complete-group accepted feedback and SQL deltas are explicit'
Write-Host 'ALL PASS'
