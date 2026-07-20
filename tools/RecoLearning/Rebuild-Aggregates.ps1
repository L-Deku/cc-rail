# tools/RecoLearning/Rebuild-Aggregates.ps1
# 由 BindingLog 流水全量重算聚合表(QuantityAlias/QuotaBox/QuotaBoxTarget/SignatureBoxMap)。
# 可随时重跑;这是"定期整理"的入口。
# 聚合规则:同一 group_key 的多目标构成一个定额组样本;
#   weight = max(0, min(100, 10*accepted + 20*corrected - 10*rejected));
#   box_id 优先沿用 mapping-boxes 原始编号(extra.box_id),否则 auto- + 目标集合哈希前 16 位。
. "$PSScriptRoot\Common.ps1"

$log = Invoke-RecoQuery -Sql "SELECT occurred_at, quantity_name, quantity_unit, target_kind, target_code, target_name, target_unit, group_key, extra, method, entry_code, entry_name FROM dbo.BindingLog WHERE quantity_name <> ''"
Write-Host ("流水行数: " + $log.Rows.Count)

# 1) 按 group_key 聚成"绑定事件组"
$groups = @{}
foreach ($row in $log.Rows) {
  $gk = [string]$row.group_key
  if (-not $groups.ContainsKey($gk)) {
    $groups[$gk] = @{ Name = [string]$row.quantity_name; Unit = [string]$row.quantity_unit; At = $row.occurred_at; Extra = [string]$row.extra; Targets = @{}; Method = ''; EntryCode = '' }
  }
  $g = $groups[$gk]
  if ($row.occurred_at -gt $g.At) { $g.At = $row.occurred_at }
  if ($g.EntryCode -eq '' -and [string]$row.entry_code -ne '') { $g.EntryCode = [string]$row.entry_code; $g.Method = [string]$row.method }
  $tk = ([string]$row.target_kind) + ':' + ([string]$row.target_code)
  if (-not $g.Targets.ContainsKey($tk)) { $g.Targets[$tk] = @{ Kind = [string]$row.target_kind; Code = [string]$row.target_code; Name = [string]$row.target_name; Unit = [string]$row.target_unit } }
}

function Get-ExtraInt { param([string]$Extra, [string]$Name, [int]$Fallback)
  if ([string]::IsNullOrWhiteSpace($Extra)) { return $Fallback }
  try { $o = $Extra | ConvertFrom-Json; $p = $o.PSObject.Properties[$Name]; if ($p -and "$($p.Value)" -match '^\d+$') { return [int]$p.Value } } catch {}
  return $Fallback
}
function Get-ExtraText { param([string]$Extra, [string]$Name)
  if ([string]::IsNullOrWhiteSpace($Extra)) { return '' }
  try { $o = $Extra | ConvertFrom-Json; $p = $o.PSObject.Properties[$Name]; if ($p) { return [string]$p.Value } } catch {}
  return ''
}

# 2a) 第一遍:先为每个目标集合定死 box 编号(mapping-boxes 原始编号优先,取字典序最小者保证确定性)
$boxes = @{}    # target_set_hash → @{ Id; Targets }
foreach ($g in $groups.Values) {
  $keys = @($g.Targets.Keys | Sort-Object)
  $setHash = Get-Md5Hex ($keys -join ';')
  $g['SetHash'] = $setHash
  $preferredId = Get-ExtraText $g.Extra 'box_id'
  if (-not $boxes.ContainsKey($setHash)) {
    $boxes[$setHash] = @{ Id = ('auto-' + $setHash.Substring(0, 16)); Targets = $g.Targets }
  }
  if ($preferredId) {
    $current = $boxes[$setHash].Id
    if ($current.StartsWith('auto-') -or [string]::CompareOrdinal($preferredId, $current) -lt 0) {
      $boxes[$setHash].Id = $preferredId
    }
  }
}

# 2b) 第二遍:用定稿的 box 编号归并别名与签名映射
$aliases = @{}  # alias_hash → @{ Raw; Unit; Sig; Count; First; Last }
$maps = @{}     # signature + '|' + box_id → 计数
foreach ($g in $groups.Values) {
  $boxId = $boxes[$g.SetHash].Id

  $sig = Get-QuantitySignature $g.Name $g.Unit
  $aliasHash = Get-Md5Hex ($g.Name + '|' + $g.Unit)
  if (-not $aliases.ContainsKey($aliasHash)) { $aliases[$aliasHash] = @{ Raw = $g.Name; Unit = $g.Unit; Sig = $sig; Count = 0; First = $g.At; Last = $g.At } }
  $a = $aliases[$aliasHash]; $a.Count++
  if ($g.At -lt $a.First) { $a.First = $g.At }
  if ($g.At -gt $a.Last) { $a.Last = $g.At }

  $acc = Get-ExtraInt $g.Extra 'accepted_count' 1
  $cor = Get-ExtraInt $g.Extra 'corrected_count' 0
  if ($cor -eq 0 -and (Get-ExtraText $g.Extra 'user_action') -eq 'correction') { $cor = 1 }
  $rej = Get-ExtraInt $g.Extra 'rejected_count' 0
  $mapKey = $sig + '|' + $boxId
  if (-not $maps.ContainsKey($mapKey)) { $maps[$mapKey] = @{ Sig = $sig; BoxId = $boxId; Acc = 0; Cor = 0; Rej = 0; Last = $g.At } }
  $m = $maps[$mapKey]; $m.Acc += $acc; $m.Cor += $cor; $m.Rej += $rej
  if ($g.At -gt $m.Last) { $m.Last = $g.At }
}

# 3) 全量重载四张聚合表
[void](Invoke-RecoNonQuery -Sql "TRUNCATE TABLE dbo.SignatureBoxMap; TRUNCATE TABLE dbo.QuotaBoxTarget; TRUNCATE TABLE dbo.QuotaBox; TRUNCATE TABLE dbo.QuantityAlias")

$dtBox = New-Object System.Data.DataTable
foreach ($c in 'box_id','target_set_hash') { [void]$dtBox.Columns.Add($c, [string]) }
$dtTarget = New-Object System.Data.DataTable
foreach ($c in 'box_id','target_kind','target_code','target_name','target_unit') { [void]$dtTarget.Columns.Add($c, [string]) }
foreach ($entry in $boxes.GetEnumerator()) {
  [void]$dtBox.Rows.Add($entry.Value.Id, $entry.Key)
  foreach ($t in $entry.Value.Targets.Values) { [void]$dtTarget.Rows.Add($entry.Value.Id, $t.Kind, $t.Code, $t.Name, $t.Unit) }
}
Invoke-RecoBulkCopy -Table $dtBox -TargetTable 'dbo.QuotaBox'
Invoke-RecoBulkCopy -Table $dtTarget -TargetTable 'dbo.QuotaBoxTarget'

$dtAlias = New-Object System.Data.DataTable
foreach ($c in 'alias_hash','raw_name','quantity_unit','signature') { [void]$dtAlias.Columns.Add($c, [string]) }
[void]$dtAlias.Columns.Add('seen_count', [int]); [void]$dtAlias.Columns.Add('first_seen', [datetime]); [void]$dtAlias.Columns.Add('last_seen', [datetime])
foreach ($entry in $aliases.GetEnumerator()) {
  $a = $entry.Value
  [void]$dtAlias.Rows.Add($entry.Key, $a.Raw, $a.Unit, $a.Sig, $a.Count, $a.First, $a.Last)
}
Invoke-RecoBulkCopy -Table $dtAlias -TargetTable 'dbo.QuantityAlias'

$dtMap = New-Object System.Data.DataTable
foreach ($c in 'signature','box_id') { [void]$dtMap.Columns.Add($c, [string]) }
foreach ($c in 'weight','accepted_count','corrected_count','rejected_count') { [void]$dtMap.Columns.Add($c, [int]) }
[void]$dtMap.Columns.Add('last_used_at', [datetime])
foreach ($m in $maps.Values) {
  $weight = [Math]::Max(0, [Math]::Min(100, 10 * $m.Acc + 20 * $m.Cor - 10 * $m.Rej))
  [void]$dtMap.Rows.Add($m.Sig, $m.BoxId, $weight, $m.Acc, $m.Cor, $m.Rej, $m.Last)
}
Invoke-RecoBulkCopy -Table $dtMap -TargetTable 'dbo.SignatureBoxMap'

# 4) 签名级条目证据:某工程量(签名)+某定额 历史上实际放过的条目,按办法分组计数。
$sigEntry = @{}
foreach ($row in $log.Rows) {
  if ([string]$row.entry_code -eq '' -or [string]$row.target_kind -ne 'quota' -or [string]$row.target_code -eq '') { continue }
  $sig = Get-QuantitySignature ([string]$row.quantity_name) ([string]$row.quantity_unit)
  $key = $sig + "`n" + [string]$row.target_code + "`n" + [string]$row.method + "`n" + [string]$row.entry_code
  if (-not $sigEntry.ContainsKey($key)) {
    $sigEntry[$key] = @{ Sig = $sig; Code = [string]$row.target_code; Method = [string]$row.method; Entry = [string]$row.entry_code; EntryName = [string]$row.entry_name; Count = 0; Last = $row.occurred_at }
  }
  $s = $sigEntry[$key]; $s.Count++
  if ($row.occurred_at -gt $s.Last) { $s.Last = $row.occurred_at }
  if ($s.EntryName -eq '' -and [string]$row.entry_name -ne '') { $s.EntryName = [string]$row.entry_name }
}
$dtSig = New-Object System.Data.DataTable
foreach ($c in 'signature','target_code','method','entry_code','entry_name') { [void]$dtSig.Columns.Add($c, [string]) }
[void]$dtSig.Columns.Add('sample_count', [int]); [void]$dtSig.Columns.Add('last_used_at', [datetime])
foreach ($s in $sigEntry.Values) { [void]$dtSig.Rows.Add($s.Sig, $s.Code, $s.Method, $s.Entry, $s.EntryName, $s.Count, $s.Last) }
[void](Invoke-RecoNonQuery -Sql "TRUNCATE TABLE dbo.SignatureEntryMap")
Invoke-RecoBulkCopy -Table $dtSig -TargetTable 'dbo.SignatureEntryMap'

# 5) 工程模板归集:条目前缀(前2位)=工程类型,统计每个工程类型下条目与定额组的共现。
$tmpl = @{}
foreach ($g in $groups.Values) {
  if ($g.EntryCode -eq '' -or $g.EntryCode.Length -lt 2) { continue }
  $prefix = $g.EntryCode.Substring(0, 2)
  $boxId = $boxes[$g.SetHash].Id
  $key = $g.Method + "`n" + $prefix + "`n" + $g.EntryCode + "`n" + $boxId
  if (-not $tmpl.ContainsKey($key)) {
    $tmpl[$key] = @{ Method = $g.Method; Prefix = $prefix; Entry = $g.EntryCode; BoxId = $boxId; Count = 0; Last = $g.At }
  }
  $t = $tmpl[$key]; $t.Count++
  if ($g.At -gt $t.Last) { $t.Last = $g.At }
}
$dtTmpl = New-Object System.Data.DataTable
foreach ($c in 'method','engineering_type','entry_code','box_id') { [void]$dtTmpl.Columns.Add($c, [string]) }
[void]$dtTmpl.Columns.Add('sample_count', [int]); [void]$dtTmpl.Columns.Add('last_seen', [datetime])
foreach ($t in $tmpl.Values) { [void]$dtTmpl.Rows.Add($t.Method, $t.Prefix, $t.Entry, $t.BoxId, $t.Count, $t.Last) }
[void](Invoke-RecoNonQuery -Sql "TRUNCATE TABLE dbo.EngineeringTemplate")
Invoke-RecoBulkCopy -Table $dtTmpl -TargetTable 'dbo.EngineeringTemplate'

# 6) 表模板行归集(一表一模板原料):带工作簿/工作表上下文的事件组,按表内行号有序沉淀。
$sheetRows = @{}
foreach ($g in $groups.Values) {
  $wb = Get-ExtraText $g.Extra 'workbook'
  if ($wb -eq '') { continue }
  $ws = Get-ExtraText $g.Extra 'worksheet'
  $rowNo = Get-ExtraInt $g.Extra 'excel_row' 0
  if ($rowNo -eq 0) {
    $cell = Get-ExtraText $g.Extra 'cell'
    if ($cell -match '(\d+)') { $rowNo = [int]$Matches[1] }
  }
  $boxId = $boxes[$g.SetHash].Id
  $sig = Get-QuantitySignature $g.Name $g.Unit
  $prefix = if ($g.EntryCode.Length -ge 2) { $g.EntryCode.Substring(0, 2) } else { '' }
  $key = $g.Method + "`n" + $wb + "`n" + $ws + "`n" + $rowNo + "`n" + $sig + "`n" + $boxId
  if (-not $sheetRows.ContainsKey($key)) {
    $sheetRows[$key] = @{ Method = $g.Method; Wb = $wb; Ws = $ws; RowNo = $rowNo; Sig = $sig; BoxId = $boxId; Entry = $g.EntryCode; Prefix = $prefix; Count = 0; Last = $g.At }
  }
  $sr = $sheetRows[$key]; $sr.Count++
  if ($g.At -gt $sr.Last) { $sr.Last = $g.At }
  if ($sr.Entry -eq '' -and $g.EntryCode -ne '') { $sr.Entry = $g.EntryCode; $sr.Prefix = $prefix }
}
$dtSheet = New-Object System.Data.DataTable
foreach ($c in 'method','workbook','worksheet') { [void]$dtSheet.Columns.Add($c, [string]) }
[void]$dtSheet.Columns.Add('excel_row', [int])
foreach ($c in 'signature','box_id','entry_code','engineering_type') { [void]$dtSheet.Columns.Add($c, [string]) }
[void]$dtSheet.Columns.Add('sample_count', [int]); [void]$dtSheet.Columns.Add('last_seen', [datetime])
foreach ($sr in $sheetRows.Values) { [void]$dtSheet.Rows.Add($sr.Method, $sr.Wb, $sr.Ws, $sr.RowNo, $sr.Sig, $sr.BoxId, $sr.Entry, $sr.Prefix, $sr.Count, $sr.Last) }
[void](Invoke-RecoNonQuery -Sql "TRUNCATE TABLE dbo.SheetTemplateRow")
Invoke-RecoBulkCopy -Table $dtSheet -TargetTable 'dbo.SheetTemplateRow'

Write-Host ("聚合完成: QuotaBox " + $dtBox.Rows.Count + " / QuotaBoxTarget " + $dtTarget.Rows.Count + " / QuantityAlias " + $dtAlias.Rows.Count + " / SignatureBoxMap " + $dtMap.Rows.Count + " / SignatureEntryMap " + $dtSig.Rows.Count + " / EngineeringTemplate " + $dtTmpl.Rows.Count + " / SheetTemplateRow " + $dtSheet.Rows.Count)
