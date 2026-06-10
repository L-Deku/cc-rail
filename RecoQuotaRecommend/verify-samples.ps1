# 检索规则回归验证：反射调用 RecoQuotaRecommend.dll 跑 AGENTS.md 规定的代表样例。
# 在临时目录副本上运行，只读，不影响部署文件和用户数据。
$ErrorActionPreference = "Stop"

$soft = "C:\Users\谢刚\Desktop\自动预算\铁路基本建设工程投资控制系统2020网络版V0503021201"
$work = Join-Path $env:TEMP ("reco-verify-" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
New-Item -ItemType Directory -Path (Join-Path $work "RecoQuotaData") -Force | Out-Null
Copy-Item (Join-Path $soft "RecoQuotaRecommend.dll") $work -Force
foreach ($f in @("quota-index.jsonl", "material-index.jsonl", "mapping-boxes.jsonl", "learning.jsonl")) {
  $src = Join-Path $soft "RecoQuotaData\$f"
  if (Test-Path $src) { Copy-Item $src (Join-Path $work "RecoQuotaData") -Force }
}

$asm = [System.Reflection.Assembly]::LoadFrom((Join-Path $work "RecoQuotaRecommend.dll"))
$flags = [System.Reflection.BindingFlags]"Static,Public,NonPublic"

$storeType = $asm.GetType("RecoQuotaRecommend.SearchIndexStore")
$store = $storeType.GetMethod("LoadOrBuild", $flags).Invoke($null, @())
$itemType = $asm.GetType("RecoQuotaRecommend.ExcelQuantityItem")
$rowType = $asm.GetType("RecoQuotaRecommend.RecommendationRow")
$searchMethod = $storeType.GetMethod("Search")

function New-QuantityItem([string]$name, [string]$unit, [string]$value) {
  $i = [Activator]::CreateInstance($itemType)
  $itemType.GetField("Name").SetValue($i, $name)
  $itemType.GetField("Unit").SetValue($i, $unit)
  $itemType.GetField("ValueText").SetValue($i, $value)
  $itemType.GetField("RawRowText").SetValue($i, $name)
  return $i
}

function Show-Rows($label, $rows) {
  Write-Host ("--- " + $label + "  返回 " + $rows.Count + " 条")
  foreach ($r in $rows) {
    $code = $rowType.GetField("QuotaCode").GetValue($r)
    $name = $rowType.GetField("QuotaName").GetValue($r)
    $kind = $rowType.GetField("TargetKind").GetValue($r)
    $score = $rowType.GetField("Score").GetValue($r)
    $src = $rowType.GetField("Source").GetValue($r)
    Write-Host ("    [" + $src + "/" + $kind + "] " + $code + " " + $name + " score=" + $score)
  }
}

Write-Host "===== 本地索引 Search 样例 ====="
foreach ($case in @(
    @("警示带", "m", "100"),
    @("警示桩", "根", "10"),
    @("HPB300钢筋", "t", "5"),
    @("HRB400钢筋", "t", "5"))) {
  $item = New-QuantityItem $case[0] $case[1] $case[2]
  $rows = $searchMethod.Invoke($store, @($item, "预算定额"))
  Show-Rows $case[0] $rows
}

Write-Host "===== 组件框 MappingStore ====="
$lrType = $asm.GetType("RecoQuotaRecommend.LearningRecord")
$listType = [System.Collections.Generic.List``1].MakeGenericType($lrType)
$records = [Activator]::CreateInstance($listType)
$mapType = $asm.GetType("RecoQuotaRecommend.MappingStore")
$mapStore = $mapType.GetMethod("Load", $flags).Invoke($null, @(, $records))

$boxesField = $mapType.GetField("boxes", [System.Reflection.BindingFlags]"Instance,NonPublic")
$boxes = $boxesField.GetValue($mapStore)
$boxType = $asm.GetType("RecoQuotaRecommend.MappingBox")
$badIds = 0
$sampleTotal = 0
foreach ($b in $boxes) {
  $id = $boxType.GetField("BoxId").GetValue($b)
  if ($id -notmatch '^box-[0-9a-f]{16}$') { $badIds++; Write-Host ("    非规范ID: " + $id) }
  $sampleTotal += $boxType.GetField("Samples").GetValue($b).Count
}
$fileBoxIds = (Get-Content (Join-Path $work "RecoQuotaData\mapping-boxes.jsonl") |
  Where-Object { $_ -match '"box_id":"([^"]+)"' } |
  ForEach-Object { ($_ | Select-String '"box_id":"([^"]+)"').Matches[0].Groups[1].Value } |
  Sort-Object -Unique)
Write-Host ("文件中旧 box_id 数: " + $fileBoxIds.Count + "，加载后框数: " + $boxes.Count + "，样本总数: " + $sampleTotal + "，非规范新ID: " + $badIds)

$findMethod = $mapType.GetMethod("Find")
$item = New-QuantityItem "土方外运" "m3" "200"
$rows = $findMethod.Invoke($mapStore, @($item, "预算定额", $store))
Show-Rows "土方外运(组件框)" $rows

Write-Host "===== 完成（临时目录: $work）====="
