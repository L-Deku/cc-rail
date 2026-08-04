# tools/RecoLearning/Initialize-RecoLearning.ps1
# 建库+建表,幂等可重跑。
. "$PSScriptRoot\Common.ps1"

$exists = Invoke-RecoScalar -Database master -Sql "SELECT COUNT(*) FROM sys.databases WHERE name = 'RecoLearning'"
if ([int]$exists -eq 0) {
  [void](Invoke-RecoNonQuery -Database master -Sql "CREATE DATABASE RecoLearning")
  Write-Host "已创建数据库 RecoLearning"
} else {
  Write-Host "数据库 RecoLearning 已存在"
}

$schema = Get-Content "$PSScriptRoot\schema.sql" -Raw -Encoding UTF8
[void](Invoke-RecoNonQuery -Sql $schema)

$tables = Invoke-RecoQuery -Sql "SELECT name FROM sys.tables ORDER BY name"
$names = @($tables.Rows | ForEach-Object { $_.name })
Write-Host ("表结构就绪(" + $names.Count + " 张): " + ($names -join ', '))
if ($names.Count -lt 8) { throw "表数量不足 8 张,建表失败" }
