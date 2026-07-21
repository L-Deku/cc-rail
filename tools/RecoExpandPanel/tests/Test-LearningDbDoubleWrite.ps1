# 双写学习库标记测试:校验 LearningDbFeature.cs 存在且 RecordMappingGroupsToStore 已挂接。
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
$learningPath = Join-Path $repoRoot 'tools\RecoExpandPanel\LearningDbFeature.cs'
if (-not (Test-Path -LiteralPath $learningPath)) { throw '缺少 LearningDbFeature.cs' }
$learning = Get-Content -LiteralPath $learningPath -Raw
$excelLink = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\ExcelLinkFeature.cs') -Raw

if ($learning -notmatch 'RecordBindingEventsToLearningDb') { throw '缺少双写入口 RecordBindingEventsToLearningDb' }
if ($learning -notmatch '192\.168\.2\.213') { throw '学习库必须固定连中央服务器,不得跟随 ServerSetting.xml' }
if ($learning -match 'ServerSetting\.xml') { throw '学习库连接不应再读 ServerSetting.xml(2020版指向另一台服务器)' }
if ($learning -notmatch 'learningDbUnavailable = true') { throw '缺少失败即停用保护' }
if ($learning -notmatch 'Connect Timeout=3') { throw '缺少短超时保护' }
if ($learning -notmatch 'catch') { throw '缺少异常吞噬保护' }
if ($learning -notmatch 'BeginTransaction') { throw '学习流水与推荐核心聚合表没有使用同一事务' }
if ($learning -notmatch 'UpsertBindingGroupAggregates') { throw '绑定流水写入后没有增量维护推荐核心聚合表' }
if ($learning -notmatch 'dbo\.SignatureBoxMap') { throw '增量学习没有更新 SignatureBoxMap' }
if ($learning -notmatch 'dbo\.QuotaBoxTarget') { throw '增量学习没有更新 QuotaBoxTarget' }
if ($learning -notmatch 'method, project_id, entry_code, entry_name') { throw 'BindingLog 没有写入真实办法/项目/条目信息' }
if ($learning -notmatch 'NormalizeForSignature\(name\) \+ "\|"') { throw 'SQL 聚合签名仍未改成名称级' }
if ($learning -match 'NormalizeForSignature\(unit\)') { throw 'SQL 聚合签名不应包含工程量单位' }
if ($learning -notmatch 'method=@method AND entry_code=@entry') { throw 'SignatureEntryMap 没有按真实办法和条目更新' }
if ($learning -notmatch "CASE WHEN @unit='' THEN target_unit") { throw '空定额单位会覆盖 SQL 已有元数据' }
if ($learning -notmatch 'source_cell') { throw '多单元格别名没有保存独立来源单元格追溯信息' }
if ($excelLink -notmatch 'RecordBindingEventsToLearningDb\(source, groups\)') { throw 'RecordMappingGroupsToStore 未挂接双写' }
if ($excelLink -notmatch 'SQL 是多人共享主学习库') { throw 'SQL 双写仍可能被本机 jsonl 失败短路' }
if ($excelLink -match 'mapping-boxes lock timeout\.\s*"\);\s*return;') { throw '本机 jsonl 锁超时不应阻止 SQL 学习' }
Write-Host 'Test-LearningDbDoubleWrite: PASS'
