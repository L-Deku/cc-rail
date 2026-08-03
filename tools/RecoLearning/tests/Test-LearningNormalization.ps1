$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$learningRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $learningRoot 'Common.ps1')

$kgSignature = Get-QuantitySignature 'HRB400钢筋 kg' 'kg'
$tonSignature = Get-QuantitySignature 'HRB400钢筋 t' 't'
if ($kgSignature -ne 'HRB400钢筋|' -or $tonSignature -ne $kgSignature) {
  throw "旧名称尾部单位没有归并到名称级签名: kg=$kgSignature t=$tonSignature"
}

$canonical = Get-CanonicalQuantityName '钢筋笼制作 kg' 'kg'
if ($canonical -ne '钢筋笼制作') { throw "尾部独立单位剥离错误: $canonical" }
if ((Get-CanonicalQuantityName '设备kg' 'kg') -ne '设备kg') { throw '没有空格分隔的名称尾部不应被误删' }

$knownUnits = @('kg','t','m','km','kg/m')
$legacyKgName = Get-CanonicalQuantityName 'HRB400钢筋 kg' '' $knownUnits
if ($legacyKgName -ne 'HRB400钢筋') { throw "空 quantity_unit 的可靠尾部单位未剥离: $legacyKgName" }
if ((Get-QuantitySignature 'HRB400钢筋 kg' '' $knownUnits) -ne 'HRB400钢筋|') { throw '空单位存量数据未归并到名称级签名' }
$legacyAliasHash = Get-Md5Hex (Get-NormalizedPart $legacyKgName)
$currentAliasHash = Get-Md5Hex (Get-NormalizedPart (Get-CanonicalQuantityName 'HRB400钢筋' 'kg' $knownUnits))
if ($legacyAliasHash -ne $currentAliasHash) { throw '存量尾部单位和新数据未使用同一 alias_hash' }
$mapKey2024 = Get-MethodScopedMapKey '2024' 'HRB400钢筋|' 'box-1'
$mapKey2020 = Get-MethodScopedMapKey '2020' 'HRB400钢筋|' 'box-1'
$mapKeyEstimate = Get-MethodScopedMapKey '101-estimate' 'HRB400钢筋|' 'box-1'
$mapKeyEstimateCn = Get-MethodScopedMapKey '101号文估算' 'HRB400钢筋|' 'box-1'
$mapKeyLegacy = Get-MethodScopedMapKey '' 'HRB400钢筋|' 'box-1'
if ($mapKeyEstimate -ne $mapKey2020 -or $mapKeyEstimateCn -ne $mapKey2020) { throw '101 号文估算未归入 2020 学习分区' }
if ($mapKey2024 -eq $mapKey2020 -or $mapKeyLegacy -eq $mapKey2020 -or $mapKeyLegacy -eq $mapKey2024) { throw '2020/2024/空办法兼容分区未正确隔离' }
if ((Get-CanonicalQuantityName '公路里程5km' '' $knownUnits) -ne '公路里程5km') { throw '无空白边界的 km 不应被剥离' }
if ((Get-CanonicalQuantityName '进口材料 lb' '' $knownUnits) -ne '进口材料 lb') { throw '未经验证的尾部词不应被当成单位' }
$specName = Get-CanonicalQuantityName '电缆规格 kg/m km' '' $knownUnits
if ($specName -ne '电缆规格 kg/m') { throw "应只剥离最后的 km，不应破坏规格中的 kg/m: $specName" }

$observedUnits = @(Get-ReliableQuantityUnits -Rows @(
  [pscustomobject]@{ quantity_unit = 'm3'; target_unit = '100m3' },
  [pscustomobject]@{ quantity_unit = ''; target_unit = 'kg/m' },
  [pscustomobject]@{ quantity_unit = 'bad unit'; target_unit = '' }
))
if ($observedUnits -notcontains 'kg/m' -or $observedUnits -contains 'bad unit') { throw '可靠单位集的观测单位筛选错误' }

$normalizedName = Get-NormalizedPart $canonical
if ((Get-Md5Hex $normalizedName) -ne (Get-Md5Hex (Get-NormalizedPart (Get-CanonicalQuantityName '钢筋笼制作 t' 't')))) {
  throw '同名不同单位的 alias_hash 应归并为同一套'
}

$formatBaseline = Get-NormalizedPart '泥浆外运(运距10km),Φ560X33.2mm'
$formatFullWidth = Get-NormalizedPart " 泥浆　外运（运距10km），Ф560×33.2mm "
$formatLowerPhi = Get-NormalizedPart '泥浆外运（运距10km），φ560ｘ33.2mm'
if ($formatBaseline -ne '泥浆外运(运距10KM),Φ560X33.2MM' -or
    $formatFullWidth -ne $formatBaseline -or $formatLowerPhi -ne $formatBaseline) {
  throw "空白、全半角标点、Ф/Φ/φ、x/× 未归并为同一签名: $formatBaseline / $formatFullWidth / $formatLowerPhi"
}
if ((Get-NormalizedPart '泥浆外运(运距5km),Φ710X33.2mm') -eq $formatBaseline) {
  throw '距离和规格数字等业务参数不得被归一化掉'
}

$rebuild = Get-Content -LiteralPath (Join-Path $learningRoot 'Rebuild-Aggregates.ps1') -Raw -Encoding UTF8
if ($rebuild -notmatch 'BeginTransaction\(\[System\.Data\.IsolationLevel\]::Serializable\)') { throw '聚合重建缺少 Serializable 事务' }
if ($rebuild -notmatch 'TABLOCKX,HOLDLOCK') { throw '聚合重建没有锁住流水快照到提交' }
if ($rebuild -notmatch '\$rebuildTransaction\.Rollback\(\)') { throw '聚合演练和失败路径缺少回滚' }
if ($rebuild -notmatch 'Get-Md5Hex \(Get-NormalizedPart \$canonicalName\)') { throw '重建 alias_hash 仍包含原始单位或原始名称' }
if ($rebuild -notmatch 'Get-ReliableQuantityUnits' -or $rebuild -notmatch 'Get-InferredTrailingQuantityUnit') { throw '重建未处理 quantity_unit 为空的存量尾部单位' }
if ($rebuild -notmatch '存量尾部单位推断' -or $rebuild -notmatch '潜在归并冲突') { throw 'DryRun 未报告推断数量和潜在冲突' }
if ($rebuild -notmatch "'signature','box_id','method'" -or $rebuild -notmatch '\$m\.Method') { throw 'SignatureBoxMap 重建数据缺少 method' }
if ($rebuild -notmatch '\$mapKey = Get-MethodScopedMapKey \$method \$sig \$boxId') { throw '同签名组件框未按 method 隔离聚合' }
if ($rebuild -notmatch '\$methodSignature = \(Get-LearningMethodPartition \(\[string\]\$g\.Method\)\) \+ "`n" \+ \$sig') { throw 'DryRun 冲突统计未按 2020/2024 办法隔离' }

$import = Get-Content -LiteralPath (Join-Path $learningRoot 'Import-JsonlLibraries.ps1') -Raw -Encoding UTF8
if ($import -notmatch '\[switch\]\$ImportBindingHistory' -or $import -notmatch '\[string\]\$SourceId') { throw '历史 JSONL 导入未改成显式一次性操作' }
if ($import -notmatch '\$alreadyCommand' -or $import -notmatch 'WHERE source IN') { throw '历史 JSONL 导入缺少来源级防重复保护' }
if ($import -notmatch 'BeginTransaction\(\[System\.Data\.IsolationLevel\]::Serializable\)' -or
    ([regex]::Matches($import, 'Invoke-RecoBulkCopyInTransaction')).Count -lt 2) { throw '两类历史流水没有在同一事务内迁移' }

Write-Host 'Test-LearningNormalization: PASS'
