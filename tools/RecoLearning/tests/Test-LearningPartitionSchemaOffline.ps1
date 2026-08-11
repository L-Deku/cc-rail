param()

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$learningRoot = Split-Path -Parent $PSScriptRoot
$schemaPath = Join-Path $learningRoot 'schema.sql'
$schema = [IO.File]::ReadAllText($schemaPath, [Text.Encoding]::UTF8)
$finalizePath = Join-Path $learningRoot 'finalize-partition-schema.sql'
$finalize = [IO.File]::ReadAllText($finalizePath, [Text.Encoding]::UTF8)

function Assert-Match {
  param([string]$Pattern, [string]$Message)
  if ($schema -notmatch $Pattern) { throw $Message }
}

if ($schema -match '(?i)\bTRUNCATE\b' -or
    $schema -match '(?i)DROP\s+CONSTRAINT' -or
    $schema -match '(?i)DROP\s+INDEX') {
  throw 'schema.sql 的 DDL-1 边界中出现了截断或存量约束/索引删除。'
}

foreach ($table in 'BindingLog','SignatureBoxMap','QuantityFormulaRule','SignatureEntryMap','EngineeringTemplate') {
  Assert-Match ("COL_LENGTH\('dbo\." + [regex]::Escape($table) + "','software_partition'\) IS NULL\s+ALTER TABLE dbo\." + [regex]::Escape($table) + " ADD software_partition NVARCHAR\(10\) NULL") ("旧库没有为 " + $table + ' 以可空方式增加 software_partition。')
}
foreach ($table in 'BindingLog','QuantityFormulaRule','SignatureEntryMap','EngineeringTemplate') {
  Assert-Match ("COL_LENGTH\('dbo\." + [regex]::Escape($table) + "','method_no'\) IS NULL\s+ALTER TABLE dbo\." + [regex]::Escape($table) + " ADD method_no NVARCHAR\(100\) NULL") ("旧库没有为 " + $table + ' 以可空方式增加 method_no。')
}

Assert-Match 'PK_SignatureBoxMap PRIMARY KEY \(software_partition, signature, box_id\)' '新库 SignatureBoxMap 目标主键错误。'
Assert-Match 'PK_SignatureEntryMap PRIMARY KEY \(software_partition, method_no, signature, target_code, entry_code\)' '新库 SignatureEntryMap 目标主键错误。'
Assert-Match 'PK_EngineeringTemplate PRIMARY KEY \(software_partition, method_no, engineering_type, entry_code, box_id\)' '新库 EngineeringTemplate 目标主键错误。'
Assert-Match 'IX_QuantityFormulaRule_partition' 'QuantityFormulaRule 缺少分区查询索引。'
Assert-Match 'IX_BindingLog_partition_entry' 'BindingLog 缺少分区/办法/条目索引。'

foreach ($index in 'IX_BindingLog_partition_entry','IX_BindingLog_partition_source','IX_SignatureBoxMap_partition','IX_QuantityFormulaRule_partition','IX_SignatureEntryMap_partition','IX_EngineeringTemplate_partition') {
  if ($finalize -notmatch ('DROP INDEX\s+' + [regex]::Escape($index)) -or $finalize -notmatch ('CREATE INDEX\s+' + [regex]::Escape($index))) {
    throw ('DDL-2 没有在 ALTER COLUMN 前后成对删除并重建依赖索引: ' + $index)
  }
}
if ($finalize.IndexOf('DROP INDEX IX_BindingLog_partition_entry',[StringComparison]::Ordinal) -gt $finalize.IndexOf('ALTER TABLE dbo.BindingLog ALTER COLUMN software_partition',[StringComparison]::Ordinal)) {
  throw 'DDL-2 在 ALTER COLUMN 之后才删除 BindingLog 分区索引。'
}

Write-Host 'PASS B1a schema.sql 安全边界及 DDL-2 依赖索引成对切换'
