$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..'))
$sourceDll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL)) { $env:RECO_EXPAND_DLL } else { Join-Path $repoRoot 'RecoQuotaRecommend\bin\RecoExpandPanel.dll' }
if (-not (Test-Path -LiteralPath $sourceDll)) { throw "找不到 $sourceDll，先构建" }

# 隔离测试数据，避免读取或覆盖实际运行目录的学习库。
$isolatedTestDir = [System.IO.Path]::Combine($repoRoot, 'obj\smartfill-pending-local-test')
if ([String]::IsNullOrWhiteSpace($isolatedTestDir)) { throw "isolated test directory is empty; repo=$repoRoot" }
[void][System.IO.Directory]::CreateDirectory($isolatedTestDir)
$dll = Join-Path $isolatedTestDir 'RecoExpandPanel.dll'
Copy-Item -LiteralPath $sourceDll -Destination $dll -Force
$sourceDllDir = Split-Path -Parent $sourceDll
foreach ($dependency in @('NPOI.dll', 'NPOI.OpenXmlFormats.dll', 'NPOI.OpenXml4Net.dll', 'NPOI.OOXML.dll', 'ICSharpCode.SharpZipLib.dll')) {
    $dependencyPath = Join-Path $sourceDllDir $dependency
    if (Test-Path -LiteralPath $dependencyPath) {
        $targetDependency = Join-Path $isolatedTestDir $dependency
        Copy-Item -LiteralPath $dependencyPath -Destination $targetDependency -Force
        [void][System.Reflection.Assembly]::LoadFrom($targetDependency)
    }
}
$dataDir = Join-Path $isolatedTestDir 'RecoQuotaData'
[void][System.IO.Directory]::CreateDirectory($dataDir)
$mappingPath = Join-Path $dataDir 'mapping-boxes.jsonl'
$suffix = [Guid]::NewGuid().ToString('N').Substring(0, 12)
$quantityName = 'CODEX名称级关系' + $suffix
$targetCode = 'TEST-Q-' + $suffix
$wrongMethodTargetCode = 'TEST-WRONG-' + $suffix
$formulaFields = ',"formula_template":"V0*0.2","formula_target_unit":"m3","formula_operand_count":"1","formula_operand_0_name":"' + $quantityName + '","formula_operand_0_unit":"m2","formula_operand_0_signature":"' + $quantityName.ToUpperInvariant() + '|"'
$mappingJsonKg = '{"record_type":"mapping_box","box_id":"box-pending-local-test","method":"2024","entry_code":"0101","target_kind":"quota","target_code":"' + $targetCode + '","target_name":"历史名称","target_unit":"历史单位","quantity_name":"' + $quantityName + '","quantity_unit":"kg","weight":"15","accepted_count":"1","corrected_count":"0","rejected_count":"0","last_used_at":"2099-07-21 14:31:18"' + $formulaFields + '}'
$mappingJsonT = '{"record_type":"mapping_box","box_id":"box-pending-local-test","method":"2024","entry_code":"0101","target_kind":"quota","target_code":"' + $targetCode + '","target_name":"历史名称","target_unit":"历史单位","quantity_name":"' + $quantityName + '","quantity_unit":"t","weight":"20","accepted_count":"2","corrected_count":"0","rejected_count":"0","last_used_at":"2099-07-21 14:31:19"' + $formulaFields + '}'
$mappingJsonWrongMethod = '{"record_type":"mapping_box","box_id":"box-pending-local-test","method":"2020","entry_code":"9901","target_kind":"quota","target_code":"' + $wrongMethodTargetCode + '","target_name":"错误办法定额","target_unit":"t","quantity_name":"' + $quantityName + '","quantity_unit":"t","weight":"100"}'
[System.IO.File]::WriteAllText($mappingPath, $mappingJsonKg + [Environment]::NewLine + $mappingJsonT + [Environment]::NewLine + $mappingJsonWrongMethod + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))
$quotaIndex = '{"quota_code":"' + $targetCode + '","quota_name":"当前版本定额","quota_unit":"t","is_current":"1"}'
[System.IO.File]::WriteAllText((Join-Path $dataDir 'quota-index.jsonl'), $quotaIndex + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))

$type = [System.Reflection.Assembly]::LoadFrom($dll).GetType('RecoNet.FormPanel', $true)
$flags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$load = $type.GetMethod('LoadSmartLearningSnapshot', $flags)
$args = New-Object 'object[]' 4
$args[0] = '2024'
$args[1] = '2024'
$args[2] = 'TB 10801—2024'
$snapshot = $load.Invoke($null, $args)
$snapshotType = $snapshot.GetType()
$merge = $type.GetMethod('MergePendingLocalMappingsIntoSmartSnapshot', $flags)
$pendingKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$pendingKeys.Add($quantityName.ToUpperInvariant() + "|`nbox-pending-local-test") | Out-Null
$mergeArgs = New-Object 'object[]' 2
$mergeArgs[0] = $snapshot.PSObject.BaseObject
$mergeArgs[1] = $pendingKeys.PSObject.BaseObject
$merge.Invoke($null, $mergeArgs) | Out-Null
$bySignature = $snapshotType.GetField('BySignature', $flags).GetValue($snapshot)
$key = $quantityName.ToUpperInvariant() + '|'
if (-not $bySignature.ContainsKey($key)) { throw "未叠加本机待汇总签名：$key / $($args[1])" }
if ($bySignature.ContainsKey($quantityName.ToUpperInvariant() + '|KG') -or $bySignature.ContainsKey($quantityName.ToUpperInvariant() + '|T')) {
    throw 'kg/t 本机历史记录仍被拆成两套签名'
}

$targets = @()
$maxWeight = 0
$hasPendingLocal = $false
foreach ($entry in @($bySignature[$key])) {
    $entryType = $entry.GetType()
    $weight = [int]$entryType.GetField('Weight', $flags).GetValue($entry)
    if ($weight -gt $maxWeight) { $maxWeight = $weight }
    if ([bool]$entryType.GetField('PendingLocal', $flags).GetValue($entry)) { $hasPendingLocal = $true }
    foreach ($target in @($entryType.GetField('Targets', $flags).GetValue($entry))) {
        $targets += [string]$target.GetType().GetField('Code', $flags).GetValue($target)
    }
}
if (@($targets | Where-Object { $_ -eq $targetCode }).Count -ne 1) { throw "旧单位记录归并后目标应只出现一次：$($targets -join ',')" }
if (@($targets | Where-Object { $_ -eq $wrongMethodTargetCode }).Count -ne 0) { throw "2020 本机关系泄漏到 2024 快照：$($targets -join ',')" }
if ($maxWeight -ne 20 -or -not $hasPendingLocal) {
    throw "本机待汇总签名未保留真实权重或没有 PendingLocal 优先标记：weight=$maxWeight pending=$hasPendingLocal"
}
$formulaByKey = $snapshotType.GetField('FormulaByKey', $flags).GetValue($snapshot)
$formulaKey = $key + "`nquota:" + $targetCode.ToUpperInvariant()
if (-not $formulaByKey.ContainsKey($formulaKey) -or $formulaByKey[$formulaKey].Count -ne 1) { throw '本机待汇总公式没有立即进入下一次推荐快照' }

$metadataMethod = $type.GetMethod('LoadCurrentSmartQuotaMetadata', $flags)
$metadata = $metadataMethod.Invoke($null, [object[]]@($null, $snapshot.PSObject.BaseObject))
$currentQuota = $metadata[$targetCode]
if ($null -eq $currentQuota) { throw '没有从当前运行目录 quota-index 读取定额元数据' }
$currentUnit = [string]$currentQuota.GetType().GetField('Unit', $flags).GetValue($currentQuota)
if ($currentUnit -ne 't') { throw "当前版本定额单位错误：$currentUnit" }

$signatureMethod = $type.GetMethod('BuildSmartQuantitySignature', $flags)
$legacySignature = [string]$signatureMethod.Invoke($null, [object[]]@('HRB400钢筋 kg', ''))
if ($legacySignature -ne 'HRB400钢筋|') { throw "空 quantity_unit 的旧尾部单位未桥接：$legacySignature" }
$materialNameSignature = [string]$signatureMethod.Invoke($null, [object[]]@('设备线夹', ''))
if ($materialNameSignature -ne '设备线夹|') { throw "普通名称尾字被误剥离：$materialNameSignature" }

$normalizeMethod = $type.GetMethod('NormalizeSmartProjectMethod', $flags)
$methodCases = @{
    'TB 10801-2024' = '2024'
    '国铁科法〔2017〕30号文' = '2020'
    '101号文估算' = '2020'
    '101-estimate' = '2020'
}
foreach ($methodCase in $methodCases.GetEnumerator()) {
    $actual = [string]$normalizeMethod.Invoke($null, [object[]]@([string]$methodCase.Key))
    if ($actual -ne $methodCase.Value) { throw "办法归一化错误：$($methodCase.Key) -> $actual" }
}

# quota-index 缓存必须在 path/length/mtime 变化时失效。
$quotaIndexPath = Join-Path $dataDir 'quota-index.jsonl'
$changedQuotaIndex = '{"quota_code":"' + $targetCode + '","quota_name":"当前版本定额已刷新","quota_unit":"kg","is_current":"1"}'
[System.IO.File]::WriteAllText($quotaIndexPath, $changedQuotaIndex + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))
[System.IO.File]::SetLastWriteTimeUtc($quotaIndexPath, [DateTime]::UtcNow.AddSeconds(5))
$metadata2 = $metadataMethod.Invoke($null, [object[]]@($null, $snapshot.PSObject.BaseObject))
$currentQuota2 = $metadata2[$targetCode]
$currentUnit2 = [string]$currentQuota2.GetType().GetField('Unit', $flags).GetValue($currentQuota2)
if ($currentUnit2 -ne 'kg') { throw "quota-index 缓存未按 path/length/mtime 失效：$currentUnit2" }

Write-Host "PASS pending键叠加、旧单位签名桥接、办法归一化、当前版本单位与缓存失效均生效：$key -> $targetCode / $currentUnit2 / weight=$maxWeight"
