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
$mappingJsonKg = '{"record_type":"mapping_box","box_id":"box-pending-local-test","target_kind":"quota","target_code":"' + $targetCode + '","target_name":"历史名称","target_unit":"历史单位","quantity_name":"' + $quantityName + '","quantity_unit":"kg","weight":"15","accepted_count":"1","corrected_count":"0","rejected_count":"0","last_used_at":"2099-07-21 14:31:18"}'
$mappingJsonT = '{"record_type":"mapping_box","box_id":"box-pending-local-test","target_kind":"quota","target_code":"' + $targetCode + '","target_name":"历史名称","target_unit":"历史单位","quantity_name":"' + $quantityName + '","quantity_unit":"t","weight":"20","accepted_count":"2","corrected_count":"0","rejected_count":"0","last_used_at":"2099-07-21 14:31:19"}'
[System.IO.File]::WriteAllText($mappingPath, $mappingJsonKg + [Environment]::NewLine + $mappingJsonT + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))
$quotaIndex = '{"quota_code":"' + $targetCode + '","quota_name":"当前版本定额","quota_unit":"t","is_current":"1"}'
[System.IO.File]::WriteAllText((Join-Path $dataDir 'quota-index.jsonl'), $quotaIndex + [Environment]::NewLine, [System.Text.UTF8Encoding]::new($false))

$type = [System.Reflection.Assembly]::LoadFrom($dll).GetType('RecoNet.FormPanel', $true)
$flags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$load = $type.GetMethod('LoadSmartLearningSnapshot', $flags)
$args = New-Object 'object[]' 2
$args[0] = '2024'
$snapshot = $load.Invoke($null, $args)
$snapshotType = $snapshot.GetType()
$bySignature = $snapshotType.GetField('BySignature', $flags).GetValue($snapshot)
$key = $quantityName.ToUpperInvariant() + '|'
if (-not $bySignature.ContainsKey($key)) { throw "未叠加本机待汇总签名：$key / $($args[1])" }
if ($bySignature.ContainsKey($quantityName.ToUpperInvariant() + '|KG') -or $bySignature.ContainsKey($quantityName.ToUpperInvariant() + '|T')) {
    throw 'kg/t 本机历史记录仍被拆成两套签名'
}

$targets = @()
$maxWeight = 0
foreach ($entry in @($bySignature[$key])) {
    $entryType = $entry.GetType()
    $weight = [int]$entryType.GetField('Weight', $flags).GetValue($entry)
    if ($weight -gt $maxWeight) { $maxWeight = $weight }
    foreach ($target in @($entryType.GetField('Targets', $flags).GetValue($entry))) {
        $targets += [string]$target.GetType().GetField('Code', $flags).GetValue($target)
    }
}
if (@($targets | Where-Object { $_ -eq $targetCode }).Count -ne 1) { throw "旧单位记录归并后目标应只出现一次：$($targets -join ',')" }
if ($maxWeight -lt 1000) { throw "本机待汇总签名没有即时优先：$maxWeight" }

$metadataMethod = $type.GetMethod('LoadCurrentSmartQuotaMetadata', $flags)
$metadata = $metadataMethod.Invoke($null, [object[]]@($null, $snapshot.PSObject.BaseObject))
$currentQuota = $metadata[$targetCode]
if ($null -eq $currentQuota) { throw '没有从当前运行目录 quota-index 读取定额元数据' }
$currentUnit = [string]$currentQuota.GetType().GetField('Unit', $flags).GetValue($currentQuota)
if ($currentUnit -ne 't') { throw "当前版本定额单位错误：$currentUnit" }
Write-Host "PASS 旧单位签名归并且当前版本单位生效：$key -> $targetCode / $currentUnit / weight=$maxWeight"
