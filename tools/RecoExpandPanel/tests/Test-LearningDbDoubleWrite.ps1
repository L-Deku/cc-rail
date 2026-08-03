# 双写学习库标记测试:校验 LearningDbFeature.cs 存在且 RecordMappingGroupsToStore 已挂接。
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
$learningPath = Join-Path $repoRoot 'tools\RecoExpandPanel\LearningDbFeature.cs'
if (-not (Test-Path -LiteralPath $learningPath)) { throw '缺少 LearningDbFeature.cs' }
$learning = Get-Content -LiteralPath $learningPath -Raw
$excelLink = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\ExcelLinkFeature.cs') -Raw
$templateMatch = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\TemplateFillNameMatch.cs') -Raw
$schema = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoLearning\schema.sql') -Raw

if ($learning -notmatch 'RecordBindingEventsToLearningDb') { throw '缺少双写入口 RecordBindingEventsToLearningDb' }
if ($learning -notmatch '192\.168\.2\.213') { throw '学习库必须固定连中央服务器,不得跟随 ServerSetting.xml' }
if ($learning -match 'ServerSetting\.xml') { throw '学习库连接不应再读 ServerSetting.xml(2020版指向另一台服务器)' }
if ($learning -match 'learningDbUnavailable') { throw '单次 SQL 失败不应永久停用本进程后续学习' }
if ($learning -notmatch 'Connect Timeout=3') { throw '缺少短超时保护' }
if ($learning -notmatch 'catch') { throw '缺少异常吞噬保护' }
if ($learning -notmatch 'BeginTransaction') { throw '学习流水与推荐核心聚合表没有使用同一事务' }
if ($learning -notmatch 'for \(int attempt = 0; attempt < 2; attempt\+\+\)') { throw '死锁/唯一冲突没有整笔重试一次' }
if ($learning -notmatch '1205' -or $learning -notmatch '2601' -or $learning -notmatch '2627') { throw '重试范围没有覆盖死锁与唯一键冲突' }
if ($learning -notmatch 'BatchId = Guid\.NewGuid\(\)\.ToString\("N"\)' -or
    $learning -notmatch 'cmd\.Parameters\.AddWithValue\("@eh", BuildLearningMd5') { throw '整笔重试没有复用稳定事件键，提交结果不确定时可能重复计数' }
if (([regex]::Matches($learning, 'UPDLOCK,HOLDLOCK')).Count -lt 6) { throw '首次并发 upsert 的键范围锁不完整' }
if ($learning -notmatch 'learning-db-outbox\.jsonl' -or $learning -notmatch 'RecoQuotaData\.learning-db-outbox\.lock') { throw 'SQL 失败没有独立命名互斥的本机 outbox' }
if ($learning -notmatch 'learning-db-outbox\.dead-letter\.jsonl' -or
    $learning -notmatch 'TryMoveLearningDbOutboxBatchToDeadLetter') { throw '永久无效的 outbox 批次没有 dead-letter 隔离通道' }
if ($learning -notmatch 'HasUnsupportedLearningDbMethod' -or
    $learning -notmatch 'was not queued because') { throw '新绑定的空办法关系仍可能进入 active outbox 或缺少可观察日志' }
if ($learning -notmatch 'LearningDbWriteResult\.PermanentFailure' -or
    $learning -notmatch 'LearningDbWriteResult\.RetryableFailure') { throw 'outbox 重放未区分永久无效与瞬时 SQL 故障' }
if ($learning -notmatch 'TryMoveLearningDbOutboxBatchToDeadLetter\(batch\.BatchId' -or
    $learning -notmatch 'continue;') { throw '毒批次隔离后没有继续重放后续批次' }
if ($learning -notmatch 'ReplayPendingLearningDbEvents' -or $learning -notmatch 'LoadPendingLearningMappingKeys') { throw '后续绑定/推荐缺少 outbox 重放与 pending 键入口' }
if ($learning -notmatch 'WriteAllLinesAtomic' -or $learning -notmatch 'RemoveLearningDbOutboxBatch') { throw 'outbox 没有原子保存或成功确认清除' }
if ($learning -notmatch 'IsLearningDbBatchAlreadyCommitted' -or $learning -notmatch 'group_key=@group_key') { throw '重复重放没有沿用稳定事件批次确认，可能重复计数' }
$serializeStart = $learning.IndexOf('SerializeLearningDbOutboxBatch', [StringComparison]::Ordinal)
$serializeEnd = $learning.IndexOf('ParseLearningDbOutboxBatch', $serializeStart, [StringComparison]::Ordinal)
$serializeBody = if ($serializeStart -ge 0 -and $serializeEnd -gt $serializeStart) { $learning.Substring($serializeStart, $serializeEnd - $serializeStart) } else { '' }
if ($serializeBody -match 'Password|connectionString|AgentDbPassword') { throw 'outbox 不得保存数据库密码或连接串' }
if ($templateMatch -notmatch 'ConsumeLearningDbDurableResult\(mappingGroups\)') { throw 'MappingFeedbackRecorded 仍可能在 SQL/outbox 均失败时阻止正式写入重试' }
foreach ($indexName in 'IX_BindingLog_recommend_entry','IX_BindingLog_recommend_source','IX_SignatureEntryMap_method','IX_QuantityFormulaRule_method') {
    if ($schema -notmatch [regex]::Escape($indexName)) { throw "推荐读取缺少索引 $indexName" }
}
if ($schema -notmatch 'PRIMARY KEY \(signature, method, box_id\)' -or $schema -notmatch "COL_LENGTH\('dbo\.SignatureBoxMap','method'\)") { throw 'schema.sql 缺少 SignatureBoxMap.method 的新表定义或存量迁移' }
if ($learning -notmatch 'UpsertBindingGroupAggregates') { throw '绑定流水写入后没有增量维护推荐核心聚合表' }
if ($learning -notmatch 'dbo\.SignatureBoxMap') { throw '增量学习没有更新 SignatureBoxMap' }
if ($learning -notmatch 'WHERE signature=@s AND method=@method AND box_id=@box' -or
    $learning -notmatch 'SignatureBoxMap\(signature,method,box_id') { throw 'SignatureBoxMap 没有按编制办法隔离组件关系' }
if ($learning -notmatch 'NormalizeLearningDbMethod' -or
    $learning -notmatch 'String\.Equals\(normalized, "2020"' -or
    $learning -notmatch 'String\.Equals\(normalized, "2024"') { throw '新增学习数据没有强制归一为 2020/2024 两个主分区' }
if ($learning -notmatch 'missing or unsupported' -or
    $learning -notmatch 'String\.IsNullOrEmpty\(NormalizeLearningDbMethod\(group\.Method\)\)') { throw '新写入仍可能落入空 method 历史兼容分区' }
if (([regex]::Matches($learning, 'string method = NormalizeLearningDbMethod\(group\.Method\);')).Count -lt 3) { throw '组件、公式和条目增量表未全部使用同一规范办法' }
if ($learning -notmatch 'NormalizeLearningDbMethod\(group\.Method\)\) \+ "\|" \+ \(group == null \? "" : group\.EntryCode') { throw '公式规则键没有同时保留规范办法和稳定条目上下文' }
if ($learning -notmatch 'dbo\.QuotaBoxTarget') { throw '增量学习没有更新 QuotaBoxTarget' }
if ($learning -notmatch 'BuildLearningTargetIdentityKey' -or
    $learning -notmatch 'HasConflictingContextSensitiveTargets') { throw '增量聚合没有隔离同码异义的通用辅助代码' }
if ($excelLink -notmatch 'BuildLearningTargetIdentityKey' -or
    $excelLink -notmatch 'HasConflictingContextSensitiveTargets') { throw '本地组件框仍可能把同码异义辅助行合并' }
if ($learning -notmatch 'method, project_id, entry_code, entry_name') { throw 'BindingLog 没有写入真实办法/项目/条目信息' }
if ($learning -notmatch 'NormalizeForSignature\(name\) \+ "\|"') { throw 'SQL 聚合签名仍未改成名称级' }
if ($learning -match 'NormalizeForSignature\(unit\)') { throw 'SQL 聚合签名不应包含工程量单位' }
if ($learning -notmatch 'method=@method AND entry_code=@entry') { throw 'SignatureEntryMap 没有按真实办法和条目更新' }
if ($learning -notmatch "CASE WHEN @unit='' THEN target_unit") { throw '空定额单位会覆盖 SQL 已有元数据' }
if ($learning -notmatch 'source_cell') { throw '多单元格别名没有保存独立来源单元格追溯信息' }
if ($excelLink -notmatch 'RecordBindingEventsToLearningDb\(source, groups\)') { throw 'RecordMappingGroupsToStore 未挂接双写' }
$storeStart = $excelLink.IndexOf('private static void RecordMappingGroupsToStore', [StringComparison]::Ordinal)
$localCatch = $excelLink.IndexOf('catch (Exception ex)', $storeStart, [StringComparison]::Ordinal)
$sqlCall = $excelLink.IndexOf('RecordBindingEventsToLearningDb(source, groups);', $storeStart, [StringComparison]::Ordinal)
if ($storeStart -lt 0 -or $localCatch -lt 0 -or $sqlCall -lt $localCatch) { throw 'SQL 双写仍可能被本机 jsonl 失败短路' }
if ($excelLink -match 'mapping-boxes lock timeout\.\s*"\);\s*return;') { throw '本机 jsonl 锁超时不应阻止 SQL 学习' }
Write-Host 'Test-LearningDbDoubleWrite: PASS'
