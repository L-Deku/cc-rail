# tools/RecoLearning/Get-LearningStats.ps1
# 学习库体检报告:各表规模、流水来源分布、条目解析率、权重榜。
. "$PSScriptRoot\Common.ps1"

Write-Host "=== RecoLearning 学习库统计 ==="
foreach ($t in 'BindingLog','QuantityAlias','QuotaBox','QuotaBoxTarget','SignatureBoxMap','QuantityFormulaRule','QuantityFormulaOperand','SignatureEntryMap','EntryQuota','ChapterEntry','EngineeringTemplate','SheetTemplateRow') {
  $exists = Invoke-RecoScalar -Sql "SELECT COUNT(*) FROM sys.tables WHERE schema_id=SCHEMA_ID('dbo') AND name=@name" -Parameters @{ name = $t }
  if ([int]$exists -eq 0) { Write-Host ("{0,-20} {1,10}" -f $t, '(未安装)'); continue }
  $n = Invoke-RecoScalar -Sql ("SELECT COUNT(*) FROM dbo." + $t)
  Write-Host ("{0,-20} {1,10:N0} 行" -f $t, [long]$n)
}

Write-Host "`n--- 流水来源分布 ---"
foreach ($row in (Invoke-RecoQuery -Sql "SELECT source, COUNT(*) AS n FROM dbo.BindingLog GROUP BY source ORDER BY n DESC").Rows) {
  Write-Host ("{0,-25} {1,8:N0}" -f $row.source, $row.n)
}

$resolved = Invoke-RecoScalar -Sql "SELECT COUNT(*) FROM dbo.BindingLog WHERE entry_code <> ''"
$totalLog = Invoke-RecoScalar -Sql "SELECT COUNT(*) FROM dbo.BindingLog"
if ([long]$totalLog -gt 0) { Write-Host ("`n条目解析率: {0:P1}" -f ([double]$resolved / [double]$totalLog)) }

$legacyRows = Invoke-RecoQuery -Sql "SELECT group_key, quantity_name, quantity_unit, target_unit FROM dbo.BindingLog WHERE quantity_name <> ''"
$knownUnits = @(Get-ReliableQuantityUnits -Rows $legacyRows.Rows)
$inferredRows = 0
$inferredGroups = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$inferredNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($row in $legacyRows.Rows) {
  if (-not [string]::IsNullOrWhiteSpace([string]$row.quantity_unit)) { continue }
  $unit = Get-InferredTrailingQuantityUnit ([string]$row.quantity_name) $knownUnits
  if ($unit -eq '') { continue }
  $inferredRows++
  [void]$inferredGroups.Add([string]$row.group_key)
  [void]$inferredNames.Add([string]$row.quantity_name)
}
Write-Host ("`n可保守推断的存量尾部单位: " + $inferredRows + " 行 / " + $inferredGroups.Count + " 组 / " + $inferredNames.Count + " 个原始名称")

Write-Host "`n--- 权重前 10 的签名映射 ---"
foreach ($row in (Invoke-RecoQuery -Sql "SELECT TOP 10 m.signature, m.box_id, m.weight, m.accepted_count, m.corrected_count FROM dbo.SignatureBoxMap m ORDER BY m.weight DESC, m.accepted_count DESC").Rows) {
  Write-Host ("[{0,3}] {1}  →  {2}" -f $row.weight, $row.signature, $row.box_id)
}
