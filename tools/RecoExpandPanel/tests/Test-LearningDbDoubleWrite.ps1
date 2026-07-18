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
if ($excelLink -notmatch 'RecordBindingEventsToLearningDb\(source, groups\)') { throw 'RecordMappingGroupsToStore 未挂接双写' }
Write-Host 'Test-LearningDbDoubleWrite: PASS'
