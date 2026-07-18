# 双写学习库标记测试:校验 LearningDbFeature.cs 存在且 RecordMappingGroupsToStore 已挂接。
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
$learningPath = Join-Path $repoRoot 'tools\RecoExpandPanel\LearningDbFeature.cs'
if (-not (Test-Path -LiteralPath $learningPath)) { throw '缺少 LearningDbFeature.cs' }
$learning = Get-Content -LiteralPath $learningPath -Raw
$excelLink = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\ExcelLinkFeature.cs') -Raw

if ($learning -notmatch 'RecordBindingEventsToLearningDb') { throw '缺少双写入口 RecordBindingEventsToLearningDb' }
if ($learning -notmatch 'learningDbUnavailable = true') { throw '缺少失败即停用保护' }
if ($learning -notmatch 'Connect Timeout=3') { throw '缺少短超时保护' }
if ($learning -notmatch 'catch') { throw '缺少异常吞噬保护' }
if ($excelLink -notmatch 'RecordBindingEventsToLearningDb\(source, groups\)') { throw 'RecordMappingGroupsToStore 未挂接双写' }
Write-Host 'Test-LearningDbDoubleWrite: PASS'
