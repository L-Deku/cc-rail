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
$wrongPartitionTargetCode = 'TEST-WRONG-' + $suffix
$mappingJsonKg = '{"record_type":"mapping_box","software_partition":"2024","box_id":"box-pending-local-test","target_kind":"quota","target_code":"' + $targetCode + '","target_name":"历史名称","target_unit":"历史单位","quantity_name":"' + $quantityName + '","quantity_unit":"kg","weight":"15","accepted_count":"1","corrected_count":"0","rejected_count":"0","last_used_at":"2099-07-21 14:31:18"}'
$mappingJsonT = '{"record_type":"mapping_box","software_partition":"2024","box_id":"box-pending-local-test","target_kind":"quota","target_code":"' + $targetCode + '","target_name":"历史名称","target_unit":"历史单位","quantity_name":"' + $quantityName + '","quantity_unit":"t","weight":"20","accepted_count":"2","corrected_count":"0","rejected_count":"0","last_used_at":"2099-07-21 14:31:19"}'
$mappingContext = '{"record_type":"mapping_context","software_partition":"2024","method_no":"TB 10801—2024","box_id":"box-pending-local-test","target_kind":"quota","target_code":"' + $targetCode + '","target_name":"历史名称","target_unit":"历史单位","quantity_name":"' + $quantityName + '","quantity_unit":"t","entry_code":"0101","formula_template":"V0*0.2","formula_target_unit":"m3","formula_method":"2024","formula_software_partition":"2024","formula_method_no":"TB 10801—2024","formula_entry_code":"0101","formula_operand_count":"1","formula_operand_0_name":"' + $quantityName + '","formula_operand_0_unit":"m2","formula_operand_0_signature":"' + $quantityName.ToUpperInvariant() + '|"}'
$mappingJsonWrongPartition = '{"record_type":"mapping_box","software_partition":"2020","box_id":"box-pending-local-test","target_kind":"quota","target_code":"' + $wrongPartitionTargetCode + '","target_name":"错误分区定额","target_unit":"t","quantity_name":"' + $quantityName + '","quantity_unit":"t","weight":"100"}'
[System.IO.File]::WriteAllText($mappingPath, ($mappingJsonKg, $mappingJsonT, $mappingContext, $mappingJsonWrongPartition -join [Environment]::NewLine) + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))
$quotaIndex = '{"quota_code":"' + $targetCode + '","quota_name":"当前版本定额","quota_unit":"t","is_current":"1"}'
[System.IO.File]::WriteAllText((Join-Path $dataDir 'quota-index.jsonl'), $quotaIndex + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))
$quotaIndex2020 = '{"quota_code":"' + $targetCode + '","source_database":"RecoData2020","quota_name":"2020版本定额","quota_unit":"hm","is_current":"1"}'
[System.IO.File]::WriteAllText((Join-Path $dataDir 'quota-index-2020.jsonl'), $quotaIndex2020 + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))

$type = [System.Reflection.Assembly]::LoadFrom($dll).GetType('RecoNet.FormPanel', $true)
$flags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$snapshotType = $type.GetNestedType('SmartLearningSnapshot', $flags)
$snapshot = [Activator]::CreateInstance($snapshotType, $true)
$snapshotBase = $snapshot.PSObject.BaseObject
$snapshotType.GetField('Method', $flags).SetValue($snapshotBase, '2024')
$snapshotType.GetField('SoftwarePartition', $flags).SetValue($snapshotBase, '2024')
$snapshotType.GetField('MethodNo', $flags).SetValue($snapshotBase, 'TB 10801—2024')
$merge = $type.GetMethod('MergePendingLocalMappingsIntoSmartSnapshot', $flags)
if ($null -ne $merge) { throw 'SQL-only 推荐仍保留本机 pending 关系叠加入口' }
$bySignature = $snapshotType.GetField('BySignature', $flags).GetValue($snapshot)
$key = $quantityName.ToUpperInvariant() + '|'
if ($bySignature.Count -ne 0) { throw '仅存在本机 mapping-boxes 时仍产生了推荐关系' }
$entryType = $type.GetNestedType('SmartMapEntry', $flags)
$targetType = $type.GetNestedType('SmartBoxTarget', $flags)
$entry = [Activator]::CreateInstance($entryType, $true).PSObject.BaseObject
$target = [Activator]::CreateInstance($targetType, $true).PSObject.BaseObject
$targetType.GetField('Kind', $flags).SetValue($target, 'quota')
$targetType.GetField('Code', $flags).SetValue($target, $targetCode)
[void]$entryType.GetField('Targets', $flags).GetValue($entry).Add($target)
$entryListType = [Collections.Generic.List``1].MakeGenericType($entryType)
$entryList = [Activator]::CreateInstance($entryListType).PSObject.BaseObject
[void]$entryList.Add($entry)
$bySignature.Add($key, $entryList)

$metadataMethod = $type.GetMethod('LoadCurrentSmartQuotaMetadata', $flags)
$metadata = $metadataMethod.Invoke($null, [object[]]@($null, $snapshot.PSObject.BaseObject))
$currentQuota = $metadata[$targetCode]
if ($null -eq $currentQuota) { throw '没有从当前运行目录 quota-index 读取定额元数据' }
$currentUnit = [string]$currentQuota.GetType().GetField('Unit', $flags).GetValue($currentQuota)
if ($currentUnit -ne 't') { throw "当前版本定额单位错误：$currentUnit" }
$snapshotType.GetField('SoftwarePartition', $flags).SetValue($snapshotBase, '2020')
$metadata2020 = $metadataMethod.Invoke($null, [object[]]@($null, $snapshot.PSObject.BaseObject))
$currentQuota2020 = $metadata2020[$targetCode]
$currentUnit2020 = [string]$currentQuota2020.GetType().GetField('Unit', $flags).GetValue($currentQuota2020)
if ($currentUnit2020 -ne 'hm') { throw "共享目录 2020 分区没有读取 quota-index-2020：$currentUnit2020" }
$snapshotType.GetField('SoftwarePartition', $flags).SetValue($snapshotBase, '2024')

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

Write-Host "PASS 本机 pending 配对已删除；共享目录分区索引只补元数据，旧单位签名、办法归一化与缓存失效均生效：$targetCode / 2020=$currentUnit2020 / 2024=$currentUnit2"
