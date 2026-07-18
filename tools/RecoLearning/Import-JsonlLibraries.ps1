# tools/RecoLearning/Import-JsonlLibraries.ps1
# 导入 RecoQuotaData 下四个 jsonl:章节树/条目定额库全量重载,mapping-boxes/learning 追加为流水。
param(
  [string]$DataDir = "D:\AI文件\自动预算\2024铁路工程云计价系统网络版V1.0\铁路工程云计价系统网络版V1.0\RecoQuotaData"
)
. "$PSScriptRoot\Common.ps1"

function Read-Jsonl {
  param([string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { throw "找不到文件: $Path" }
  foreach ($line in [System.IO.File]::ReadLines($Path, [System.Text.Encoding]::UTF8)) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $line | ConvertFrom-Json
  }
}

function Get-Prop { param($Obj, [string]$Name) $p = $Obj.PSObject.Properties[$Name]; if ($p) { [string]$p.Value } else { '' } }

# ---------- 1. ChapterEntry:全量重载 ----------
$dt = New-Object System.Data.DataTable
foreach ($c in 'method','method_no','entry_code','entry_name','unit','entry_type') { [void]$dt.Columns.Add($c, [string]) }
[void]$dt.Columns.Add('level', [int])
$seen = @{}; $srcCount = 0
foreach ($r in (Read-Jsonl (Join-Path $DataDir 'chapter-entries.jsonl'))) {
  $srcCount++
  $key = (Get-Prop $r 'method') + '|' + (Get-Prop $r 'entry_code')
  if ($seen.ContainsKey($key)) { continue }
  $seen[$key] = $true
  $lvl = 0; [void][int]::TryParse((Get-Prop $r 'level'), [ref]$lvl)
  [void]$dt.Rows.Add((Get-Prop $r 'method'), (Get-Prop $r 'method_no'), (Get-Prop $r 'entry_code'), (Get-Prop $r 'entry_name'), (Get-Prop $r 'unit'), (Get-Prop $r 'entry_type'), $lvl)
}
[void](Invoke-RecoNonQuery -Sql "TRUNCATE TABLE dbo.ChapterEntry")
Invoke-RecoBulkCopy -Table $dt -TargetTable 'dbo.ChapterEntry'
Write-Host ("ChapterEntry: 源 $srcCount 行,去重导入 " + $dt.Rows.Count + " 行")

# ---------- 2. EntryQuota:全量重载,重复键保留 project_count 最大者 ----------
$best = @{}; $srcCount = 0
foreach ($r in (Read-Jsonl (Join-Path $DataDir 'chapter-quota-library.jsonl'))) {
  $srcCount++
  $key = ((Get-Prop $r 'method') + '|' + (Get-Prop $r 'entry_code') + '|' + (Get-Prop $r 'target_kind') + '|' + (Get-Prop $r 'quota_code'))
  $pc = 0; [void][int]::TryParse((Get-Prop $r 'project_count'), [ref]$pc)
  if ($best.ContainsKey($key) -and $best[$key].ProjectCount -ge $pc) { continue }
  $best[$key] = [pscustomobject]@{ Row = $r; ProjectCount = $pc }
}
$dt2 = New-Object System.Data.DataTable
foreach ($c in 'method','method_no','entry_code','entry_name','target_kind','quota_code','quota_name','quota_unit') { [void]$dt2.Columns.Add($c, [string]) }
[void]$dt2.Columns.Add('project_count', [int])
foreach ($c in 'source','last_seen') { [void]$dt2.Columns.Add($c, [string]) }
foreach ($item in $best.Values) {
  $r = $item.Row
  $tk = Get-Prop $r 'target_kind'; if ($tk -eq '') { $tk = 'quota' }
  [void]$dt2.Rows.Add((Get-Prop $r 'method'), (Get-Prop $r 'method_no'), (Get-Prop $r 'entry_code'), (Get-Prop $r 'entry_name'), $tk, (Get-Prop $r 'quota_code'), (Get-Prop $r 'quota_name'), (Get-Prop $r 'quota_unit'), $item.ProjectCount, (Get-Prop $r 'source'), (Get-Prop $r 'last_seen'))
}
[void](Invoke-RecoNonQuery -Sql "TRUNCATE TABLE dbo.EntryQuota")
Invoke-RecoBulkCopy -Table $dt2 -TargetTable 'dbo.EntryQuota'
Write-Host ("EntryQuota: 源 $srcCount 行,去重导入 " + $dt2.Rows.Count + " 行")

# ---------- 3. mapping-boxes.jsonl → BindingLog(追加,按行哈希幂等) ----------
$existing = @{}
foreach ($row in (Invoke-RecoQuery -Sql "SELECT event_hash FROM dbo.BindingLog").Rows) { $existing[$row.event_hash] = $true }

function Parse-EventTime {
  param([string]$Text)
  $t = [datetime]::MinValue
  if ([datetime]::TryParse($Text, [ref]$t)) { return $t }
  return (Get-Date)
}

$dt3 = New-Object System.Data.DataTable
foreach ($c in 'source','method','project_id','entry_code','entry_name','quantity_name','quantity_unit','target_kind','target_code','target_name','target_unit','group_key','event_hash','extra') { [void]$dt3.Columns.Add($c, [string]) }
[void]$dt3.Columns.Add('occurred_at', [datetime])

$srcCount = 0; $skipped = 0
foreach ($line in [System.IO.File]::ReadLines((Join-Path $DataDir 'mapping-boxes.jsonl'), [System.Text.Encoding]::UTF8)) {
  if ([string]::IsNullOrWhiteSpace($line)) { continue }
  $srcCount++
  $hash = Get-Md5Hex ('mapping-boxes|' + $line)
  if ($existing.ContainsKey($hash)) { $skipped++; continue }
  $r = $line | ConvertFrom-Json
  $entryCodes = Get-Prop $r 'entry_codes'
  $method = ''; $entryCode = ''
  if ($entryCodes -match '^([^:]+):([^,]+)') { $method = $Matches[1]; $entryCode = $Matches[2] }
  $extra = @{ box_id = (Get-Prop $r 'box_id'); weight = (Get-Prop $r 'weight'); accepted_count = (Get-Prop $r 'accepted_count'); corrected_count = (Get-Prop $r 'corrected_count'); rejected_count = (Get-Prop $r 'rejected_count'); entry_codes = $entryCodes } | ConvertTo-Json -Compress
  $groupKey = Get-Md5Hex ((Get-Prop $r 'box_id') + '|' + (Get-QuantitySignature (Get-Prop $r 'quantity_name') (Get-Prop $r 'quantity_unit')))
  [void]$dt3.Rows.Add('import:mapping-boxes', $method, '', $entryCode, '', (Get-Prop $r 'quantity_name'), (Get-Prop $r 'quantity_unit'), (Get-Prop $r 'target_kind'), (Get-Prop $r 'target_code'), (Get-Prop $r 'target_name'), (Get-Prop $r 'target_unit'), $groupKey, $hash, $extra, (Parse-EventTime (Get-Prop $r 'last_used_at')))
}
Invoke-RecoBulkCopy -Table $dt3 -TargetTable 'dbo.BindingLog'
Write-Host ("mapping-boxes: 源 $srcCount 行,新增 " + $dt3.Rows.Count + " 行,跳过已存在 $skipped 行")

# ---------- 4. learning.jsonl → BindingLog(追加,按行哈希幂等) ----------
$dt4 = $dt3.Clone()
$srcCount = 0; $skipped = 0
foreach ($line in [System.IO.File]::ReadLines((Join-Path $DataDir 'learning.jsonl'), [System.Text.Encoding]::UTF8)) {
  if ([string]::IsNullOrWhiteSpace($line)) { continue }
  $srcCount++
  $hash = Get-Md5Hex ('learning|' + $line)
  if ($existing.ContainsKey($hash)) { $skipped++; continue }
  $r = $line | ConvertFrom-Json
  $extra = @{ user_action = (Get-Prop $r 'user_action'); match_reason = (Get-Prop $r 'match_reason'); match_score = (Get-Prop $r 'match_score') } | ConvertTo-Json -Compress
  $groupKey = Get-Md5Hex ((Get-Prop $r 'quantity_signature') + '|' + (Get-Prop $r 'updated_at'))
  [void]$dt4.Rows.Add('import:learning', '', '', '', '', (Get-Prop $r 'quantity_name'), (Get-Prop $r 'quantity_unit'), 'quota', (Get-Prop $r 'quota_code'), (Get-Prop $r 'quota_name'), (Get-Prop $r 'quota_unit'), $groupKey, $hash, $extra, (Parse-EventTime (Get-Prop $r 'updated_at')))
}
Invoke-RecoBulkCopy -Table $dt4 -TargetTable 'dbo.BindingLog'
Write-Host ("learning: 源 $srcCount 行,新增 " + $dt4.Rows.Count + " 行,跳过已存在 $skipped 行")
