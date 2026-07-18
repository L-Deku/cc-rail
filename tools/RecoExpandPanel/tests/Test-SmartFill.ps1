# 智能铺量标记测试:引擎文件、漏斗关键路径、UI 挂接、菜单入口。
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
$smartPath = Join-Path $repoRoot 'tools\RecoExpandPanel\SmartFillFeature.cs'
if (-not (Test-Path -LiteralPath $smartPath)) { throw '缺少 SmartFillFeature.cs' }
$smart = Get-Content -LiteralPath $smartPath -Raw
$panel = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\TemplateFillPanel.cs') -Raw
$excelLink = Get-Content -LiteralPath (Join-Path $repoRoot 'tools\RecoExpandPanel\ExcelLinkFeature.cs') -Raw

if ($smart -notmatch 'BuildPreview_SmartFill') { throw '缺少 BuildPreview_SmartFill' }
if ($smart -notmatch 'LoadSmartLearningSnapshot') { throw '缺少学习库快照加载 LoadSmartLearningSnapshot' }
if ($smart -notmatch 'BuildMappingBoxIndex|LoadMappingBoxRows') { throw '缺少 jsonl 回退' }
if ($smart -notmatch 'IsLibraryQuota = true') { throw '缺少库内定额原生粘贴路径' }
if ($smart -notmatch 'IsNameDriven = true') { throw '智能铺量项必须 IsNameDriven=true,否则不回流学习库' }
if ($panel -notmatch '三·智能铺量') { throw 'cmbMode 缺少第三模式' }
if ($panel -notmatch 'BuildPreview_SmartFill') { throw 'OnPreview 未挂接智能铺量分支' }
if ($excelLink -notmatch '打开智能铺量面板') { throw '缺少智能铺量菜单入口' }
Write-Host 'Test-SmartFill: PASS'
