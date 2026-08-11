$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
$learning = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\LearningDbFeature.cs') -Raw -Encoding UTF8
$excelLink = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\ExcelLinkFeature.cs') -Raw -Encoding UTF8
$templateMatch = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\TemplateFillNameMatch.cs') -Raw -Encoding UTF8
$schema = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoLearning\schema.sql') -Raw -Encoding UTF8
$credentialStore = Get-Content -LiteralPath (Join-Path $repoRoot 'RecoShared\RecoSqlCredentialStore.cs') -Raw -Encoding UTF8

if ($learning -notmatch 'RecordBindingEventsToLearningDb') { throw '缺少 SQL 学习入口 RecordBindingEventsToLearningDb' }
if ($learning -notmatch 'RecoSqlCredentialStore\.BuildConnectionString\("learning", "RecoLearning", 1433, 3\)') { throw '学习库连接没有统一走共享 DPAPI 凭据存储' }
if ($learning -match 'ServerSetting\.xml') { throw '学习库连接不应读取业务库配置' }
if ($learning -match '\b(?:\d{1,3}\.){3}\d{1,3}\b' -or $learning -match 'User ID=|Password=' -or
    $credentialStore -match 'User ID=|Password=') { throw '学习库源码仍含服务器、账号或密码字面量' }
if ($learning -notmatch 'BeginTransaction') { throw '学习流水与推荐核心聚合表没有使用同一事务' }
if ($learning -notmatch 'for \(int attempt = 0; attempt < 2; attempt\+\+\)') { throw '死锁或唯一冲突没有整笔重试一次' }
if ($learning -notmatch '1205' -or $learning -notmatch '2601' -or $learning -notmatch '2627') { throw '短重试范围不完整' }
if (([regex]::Matches($learning, 'UPDLOCK,HOLDLOCK')).Count -lt 6) { throw '首次并发 upsert 的键范围锁不完整' }

$recordStart = $learning.IndexOf('private static bool RecordBindingEventsToLearningDb', [StringComparison]::Ordinal)
$recordEnd = $learning.IndexOf('private static void WriteBindingEventsToLearningDb', $recordStart, [StringComparison]::Ordinal)
$recordBody = if ($recordStart -ge 0 -and $recordEnd -gt $recordStart) { $learning.Substring($recordStart, $recordEnd - $recordStart) } else { '' }
if ([String]::IsNullOrWhiteSpace($recordBody)) { throw '无法定位 SQL 学习入口方法体' }
foreach ($forbidden in @('TryAppendLearningDbOutbox(', 'TryAppendLearningDbDeadLetter(', 'TryReplayPendingLearningDbEvents(')) {
    if ($recordBody.Contains($forbidden)) { throw "SQL-only 学习入口仍写本地队列: $forbidden" }
}
if ($recordBody -notmatch 'Learning was rejected before SQL write because the current software partition is unknown' -or
    $recordBody -notmatch 'IsValidLearningSoftwarePartition\(group\.SoftwarePartition\)' -or
    $recordBody -notmatch 'String\.IsNullOrEmpty\(group\.MethodNo\)' -or
    $recordBody -notmatch 'String\.IsNullOrEmpty\(NormalizeLearningDbMethod\(group\.Method\)\)') {
    throw 'SQL 写入仍可能接受未知分区、空办法号或空 method'
}
if ($recordBody.Contains('IsLearningFeedbackGroupRecommendable')) {
    throw 'BindingLog 审计入口不得用推荐语义门禁丢弃原始流水'
}

if ($learning -notmatch 'UpsertBindingGroupAggregates' -or $learning -notmatch 'dbo\.SignatureBoxMap' -or
    $learning -notmatch 'dbo\.QuotaBoxTarget') { throw 'SQL 写入没有同步维护推荐聚合表' }
if ($learning -notmatch 'WHERE software_partition=@software_partition AND signature=@s AND box_id=@box' -or
    $learning -notmatch 'SignatureBoxMap\(software_partition,signature,method,box_id') { throw 'SignatureBoxMap 没有按软件分区隔离' }
if ($learning -notmatch 'WHERE software_partition=@software_partition AND method_no=@method_no AND signature=@s' -or
    $learning -notmatch 'SignatureEntryMap\(software_partition,method_no,signature,target_code,method,entry_code') { throw '条目聚合没有按分区和办法号隔离' }
if ($learning -notmatch 'GetMappingFeedbackTargetEntryCode\(group, target\)' -or
    $learning -notmatch 'BuildLearningTargetIdentityKey') { throw 'SQL 学习缺少目标级条目或同码异义保护' }
$aggregateStart = $learning.IndexOf('private static void UpsertBindingGroupAggregates', [StringComparison]::Ordinal)
$aggregateBody = if ($aggregateStart -ge 0) { $learning.Substring($aggregateStart) } else { '' }
if ($aggregateBody -notmatch 'IsLearningFeedbackGroupRecommendable\(group\)') {
    throw '推荐聚合入口缺少纯辅助、同码异义或 SF 条目门禁'
}
if ($learning -match 'NormalizeForSignature\(unit\)') { throw 'SQL 聚合签名不应包含工程量单位' }
if ($learning -notmatch 'source_cell') { throw '多单元格别名没有保存独立来源单元格' }
foreach ($indexName in 'IX_BindingLog_recommend_entry','IX_BindingLog_recommend_source','IX_SignatureEntryMap_method','IX_QuantityFormulaRule_method') {
    if ($schema -notmatch [regex]::Escape($indexName)) { throw "推荐读取缺少索引 $indexName" }
}

$storeStart = $excelLink.IndexOf('private static void RecordMappingGroupsToLearningDb', [StringComparison]::Ordinal)
$storeEnd = $excelLink.IndexOf('private static LocalMappingSaveResult SaveMappingGroupsToLocalFile', $storeStart, [StringComparison]::Ordinal)
$storeBody = if ($storeStart -ge 0 -and $storeEnd -gt $storeStart) { $excelLink.Substring($storeStart, $storeEnd - $storeStart) } else { '' }
if ([String]::IsNullOrWhiteSpace($storeBody) -or -not $storeBody.Contains('RecordBindingEventsToLearningDb(source, groups)')) { throw '绑定入口没有调用 SQL 学习库' }
if ($storeBody -match 'SaveMappingGroupsToLocalFile|mapping-boxes\.jsonl') { throw '绑定入口仍写本机 mapping-boxes' }
if ($templateMatch -notmatch 'ConsumeLearningDbDurableResult\(mappingGroups\)' -or
    $templateMatch -notmatch 'SqlFeedbackDurable') { throw 'SQL 失败后的当前预览重试状态未保留' }

Write-Host 'Test-LearningDbDoubleWrite (SQL-only): PASS'
