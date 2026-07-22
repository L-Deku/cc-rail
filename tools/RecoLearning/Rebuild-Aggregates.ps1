# tools/RecoLearning/Rebuild-Aggregates.ps1
# 由 BindingLog 流水全量重算推荐核心、公式、条目和模板聚合表。
# 可随时重跑;这是"定期整理"的入口。
# 聚合规则:同一 group_key 的多目标构成一个定额组样本;
#   weight = max(0, min(100, 10*accepted + 20*corrected - 10*rejected));
#   box_id 优先沿用 mapping-boxes 原始编号(extra.box_id),否则 auto- + 目标集合哈希前 16 位。
param([switch]$DryRun)
. "$PSScriptRoot\Common.ps1"

$rebuildConnection = New-Object System.Data.SqlClient.SqlConnection (Get-RecoConnectionString)
$rebuildTransaction = $null
try {
$rebuildConnection.Open()
$rebuildTransaction = $rebuildConnection.BeginTransaction([System.Data.IsolationLevel]::Serializable)
# 从读取流水到全部聚合替换提交始终持有流水表独占锁，防止增量绑定在快照与替换之间被吞掉。
$log = Invoke-RecoQueryInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Sql "SELECT occurred_at, quantity_name, quantity_unit, target_kind, target_code, target_name, target_unit, group_key, extra, method, entry_code, entry_name FROM dbo.BindingLog WITH (TABLOCKX,HOLDLOCK) WHERE quantity_name <> ''"
Write-Host ("流水行数: " + $log.Rows.Count)
$knownUnits = @(Get-ReliableQuantityUnits -Rows $log.Rows)
$inferredRows = 0
$inferredGroups = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$inferredNames = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
$inferredUnits = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($row in $log.Rows) {
  if (-not [string]::IsNullOrWhiteSpace([string]$row.quantity_unit)) { continue }
  $inferredUnit = Get-InferredTrailingQuantityUnit ([string]$row.quantity_name) $knownUnits
  if ($inferredUnit -eq '') { continue }
  $inferredRows++
  [void]$inferredGroups.Add([string]$row.group_key)
  [void]$inferredNames.Add([string]$row.quantity_name)
  [void]$inferredUnits.Add($inferredUnit)
}

# 1) 按 group_key 聚成"绑定事件组"
$groups = @{}
foreach ($row in $log.Rows) {
  $gk = [string]$row.group_key
  if (-not $groups.ContainsKey($gk)) {
    $groups[$gk] = @{ Name = [string]$row.quantity_name; Unit = [string]$row.quantity_unit; At = $row.occurred_at; Extra = [string]$row.extra; Targets = @{}; Method = ''; EntryCode = '' }
  }
  $g = $groups[$gk]
  if ($row.occurred_at -gt $g.At) { $g.At = $row.occurred_at }
  if ($g.Method -eq '' -and [string]$row.method -ne '') { $g.Method = Get-LearningMethodPartition ([string]$row.method) }
  if ($g.EntryCode -eq '' -and [string]$row.entry_code -ne '') { $g.EntryCode = [string]$row.entry_code }
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
  $g['EffectiveUnit'] = ([string]$g.Unit).Trim()
  if ($g.EffectiveUnit -eq '') { $g.EffectiveUnit = Get-InferredTrailingQuantityUnit $g.Name $knownUnits }
  $g['UnitInferred'] = ([string]$g.Unit).Trim() -eq '' -and $g.EffectiveUnit -ne ''
  $g['CanonicalName'] = Get-CanonicalQuantityName $g.Name $g.Unit $knownUnits
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
$maps = @{}     # method + signature + box_id → 计数，2020/2024 隔离，101 号文估算归入 2020
foreach ($g in $groups.Values) {
  $boxId = $boxes[$g.SetHash].Id

  $canonicalName = $g.CanonicalName
  $sig = Get-QuantitySignature $canonicalName $g.EffectiveUnit $knownUnits
  $aliasHash = Get-Md5Hex (Get-NormalizedPart $canonicalName)
  if (-not $aliases.ContainsKey($aliasHash)) { $aliases[$aliasHash] = @{ Raw = $canonicalName; Unit = $g.EffectiveUnit; Sig = $sig; Count = 0; First = $g.At; Last = $g.At } }
  $a = $aliases[$aliasHash]; $a.Count++
  if ($g.At -lt $a.First) { $a.First = $g.At }
  if ($g.At -gt $a.Last) { $a.Last = $g.At; $a.Raw = $canonicalName; $a.Unit = $g.EffectiveUnit }

  $acc = Get-ExtraInt $g.Extra 'accepted_count' 1
  $cor = Get-ExtraInt $g.Extra 'corrected_count' 0
  if ($cor -eq 0 -and (Get-ExtraText $g.Extra 'user_action') -eq 'correction') { $cor = 1 }
  $rej = Get-ExtraInt $g.Extra 'rejected_count' 0
  $method = Get-LearningMethodPartition ([string]$g.Method)
  $mapKey = Get-MethodScopedMapKey $method $sig $boxId
  if (-not $maps.ContainsKey($mapKey)) { $maps[$mapKey] = @{ Method = $method; Sig = $sig; BoxId = $boxId; Acc = 0; Cor = 0; Rej = 0; Last = $g.At } }
  $m = $maps[$mapKey]; $m.Acc += $acc; $m.Cor += $cor; $m.Rej += $rej
  if ($g.At -gt $m.Last) { $m.Last = $g.At }
}

$inferenceBySignature = @{}
foreach ($g in $groups.Values) {
  $sig = Get-QuantitySignature $g.CanonicalName $g.EffectiveUnit $knownUnits
  $methodSignature = (Get-LearningMethodPartition ([string]$g.Method)) + "`n" + $sig
  if (-not $inferenceBySignature.ContainsKey($methodSignature)) {
    $inferenceBySignature[$methodSignature] = @{
      AllSets = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
      InferredSets = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
      InferredUnits = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
      HasExplicit = $false
      HasInferred = $false
    }
  }
  $audit = $inferenceBySignature[$methodSignature]
  [void]$audit.AllSets.Add($g.SetHash)
  if ($g.UnitInferred) {
    $audit.HasInferred = $true
    [void]$audit.InferredSets.Add($g.SetHash)
    [void]$audit.InferredUnits.Add($g.EffectiveUnit)
  } else {
    $audit.HasExplicit = $true
  }
}
$inferredSignatureCount = 0
$inferredMergeCount = 0
$multiUnitInferenceCount = 0
$potentialMappingConflictCount = 0
foreach ($audit in $inferenceBySignature.Values) {
  if (-not $audit.HasInferred) { continue }
  $inferredSignatureCount++
  if ($audit.HasExplicit) { $inferredMergeCount++ }
  if ($audit.InferredUnits.Count -gt 1) { $multiUnitInferenceCount++ }
  if ($audit.AllSets.Count -gt 1) { $potentialMappingConflictCount++ }
}

# 2c) 已确认跨单位数量公式：单系数和多参数公式统一用 V0/V1... 模板表示。
$formulas = @{}
foreach ($row in $log.Rows) {
  $extra = [string]$row.extra
  $template = Get-ExtraText $extra 'formula_template'
  $operandCount = Get-ExtraInt $extra 'formula_operand_count' 0
  if ($template -eq '' -or $operandCount -le 0) { continue }
  $sig = Get-QuantitySignature ([string]$row.quantity_name) ([string]$row.quantity_unit) $knownUnits
  $kind = [string]$row.target_kind; if ($kind -eq '') { $kind = 'quota' }
  $code = [string]$row.target_code
  $targetUnit = Get-ExtraText $extra 'formula_target_unit'; if ($targetUnit -eq '') { $targetUnit = [string]$row.target_unit }
  $operands = @()
  $formulaMethod = Get-LearningMethodPartition ([string]$row.method)
  $ruleRaw = $sig + '|' + $kind + ':' + $code.ToUpperInvariant() + '|' + $targetUnit.Trim().ToLowerInvariant() + '|' + $template + '|' + $formulaMethod + '|' + [string]$row.entry_code
  $valid = $true
  for ($i = 0; $i -lt $operandCount; $i++) {
    $prefix = 'formula_operand_' + $i + '_'
    $operandName = Get-ExtraText $extra ($prefix + 'name')
    $operandUnit = Get-ExtraText $extra ($prefix + 'unit')
    $operandSig = Get-ExtraText $extra ($prefix + 'signature')
    if ($operandName -ne '') { $operandSig = Get-QuantitySignature $operandName $operandUnit $knownUnits }
    if ($operandSig -eq '') { $valid = $false; break }
    $lastBar = $operandSig.LastIndexOf('|')
    $operandSig = if ($lastBar -ge 0) { $operandSig.Substring(0, $lastBar) + '|' } else { $operandSig + '|' }
    $operands += @{ Index = $i; Sig = $operandSig; Name = $operandName; Unit = $operandUnit }
    $ruleRaw += '|' + $operandSig + '@' + $operandUnit.Trim().ToLowerInvariant()
  }
  if (-not $valid) { continue }
  $ruleHash = Get-ExtraText $extra 'formula_rule_hash'
  if ($ruleHash -eq '' -or $formulaMethod -ne ([string]$row.method).Trim()) { $ruleHash = Get-Md5Hex $ruleRaw }
  if (-not $formulas.ContainsKey($ruleHash)) {
    $formulas[$ruleHash] = @{ Hash = $ruleHash; Sig = $sig; Kind = $kind; Code = $code; Unit = $targetUnit; Template = $template; Method = $formulaMethod; Entry = [string]$row.entry_code; Count = 0; First = $row.occurred_at; Last = $row.occurred_at; Operands = $operands }
  }
  $f = $formulas[$ruleHash]; $f.Count++
  if ($row.occurred_at -lt $f.First) { $f.First = $row.occurred_at }
  if ($row.occurred_at -gt $f.Last) { $f.Last = $row.occurred_at }
}

# 3) 全量重载推荐核心和数量公式聚合表
[void](Invoke-RecoNonQueryInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Sql "TRUNCATE TABLE dbo.SignatureBoxMap; TRUNCATE TABLE dbo.QuotaBoxTarget; TRUNCATE TABLE dbo.QuotaBox; TRUNCATE TABLE dbo.QuantityAlias")

$dtBox = New-Object System.Data.DataTable
foreach ($c in 'box_id','target_set_hash') { [void]$dtBox.Columns.Add($c, [string]) }
$dtTarget = New-Object System.Data.DataTable
foreach ($c in 'box_id','target_kind','target_code','target_name','target_unit') { [void]$dtTarget.Columns.Add($c, [string]) }
foreach ($entry in $boxes.GetEnumerator()) {
  [void]$dtBox.Rows.Add($entry.Value.Id, $entry.Key)
  foreach ($t in $entry.Value.Targets.Values) { [void]$dtTarget.Rows.Add($entry.Value.Id, $t.Kind, $t.Code, $t.Name, $t.Unit) }
}
Invoke-RecoBulkCopyInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Table $dtBox -TargetTable 'dbo.QuotaBox'
Invoke-RecoBulkCopyInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Table $dtTarget -TargetTable 'dbo.QuotaBoxTarget'

$dtAlias = New-Object System.Data.DataTable
foreach ($c in 'alias_hash','raw_name','quantity_unit','signature') { [void]$dtAlias.Columns.Add($c, [string]) }
[void]$dtAlias.Columns.Add('seen_count', [int]); [void]$dtAlias.Columns.Add('first_seen', [datetime]); [void]$dtAlias.Columns.Add('last_seen', [datetime])
foreach ($entry in $aliases.GetEnumerator()) {
  $a = $entry.Value
  [void]$dtAlias.Rows.Add($entry.Key, $a.Raw, $a.Unit, $a.Sig, $a.Count, $a.First, $a.Last)
}
Invoke-RecoBulkCopyInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Table $dtAlias -TargetTable 'dbo.QuantityAlias'

$dtMap = New-Object System.Data.DataTable
foreach ($c in 'signature','box_id','method') { [void]$dtMap.Columns.Add($c, [string]) }
foreach ($c in 'weight','accepted_count','corrected_count','rejected_count') { [void]$dtMap.Columns.Add($c, [int]) }
[void]$dtMap.Columns.Add('last_used_at', [datetime])
foreach ($m in $maps.Values) {
  $weight = [Math]::Max(0, [Math]::Min(100, 10 * $m.Acc + 20 * $m.Cor - 10 * $m.Rej))
  [void]$dtMap.Rows.Add($m.Sig, $m.BoxId, $m.Method, $weight, $m.Acc, $m.Cor, $m.Rej, $m.Last)
}
Invoke-RecoBulkCopyInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Table $dtMap -TargetTable 'dbo.SignatureBoxMap'

$dtFormula = New-Object System.Data.DataTable
foreach ($c in 'rule_hash','anchor_signature','target_kind','target_code','target_unit','formula_template','method','entry_code') { [void]$dtFormula.Columns.Add($c, [string]) }
[void]$dtFormula.Columns.Add('sample_count', [int]); [void]$dtFormula.Columns.Add('first_seen', [datetime]); [void]$dtFormula.Columns.Add('last_seen', [datetime])
$dtOperand = New-Object System.Data.DataTable
[void]$dtOperand.Columns.Add('rule_hash', [string]); [void]$dtOperand.Columns.Add('operand_index', [int])
foreach ($c in 'operand_signature','operand_name','operand_unit') { [void]$dtOperand.Columns.Add($c, [string]) }
foreach ($f in $formulas.Values) {
  [void]$dtFormula.Rows.Add($f.Hash, $f.Sig, $f.Kind, $f.Code, $f.Unit, $f.Template, $f.Method, $f.Entry, $f.Count, $f.First, $f.Last)
  foreach ($o in $f.Operands) { [void]$dtOperand.Rows.Add($f.Hash, $o.Index, $o.Sig, $o.Name, $o.Unit) }
}
[void](Invoke-RecoNonQueryInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Sql "TRUNCATE TABLE dbo.QuantityFormulaOperand; TRUNCATE TABLE dbo.QuantityFormulaRule")
Invoke-RecoBulkCopyInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Table $dtFormula -TargetTable 'dbo.QuantityFormulaRule'
Invoke-RecoBulkCopyInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Table $dtOperand -TargetTable 'dbo.QuantityFormulaOperand'

# 4) 签名级条目证据:某工程量(签名)+某定额 历史上实际放过的条目,按办法分组计数。
$sigEntry = @{}
foreach ($row in $log.Rows) {
  if ([string]$row.entry_code -eq '' -or [string]$row.target_kind -ne 'quota' -or [string]$row.target_code -eq '') { continue }
  $sig = Get-QuantitySignature ([string]$row.quantity_name) ([string]$row.quantity_unit) $knownUnits
  $entryMethod = Get-LearningMethodPartition ([string]$row.method)
  $key = $sig + "`n" + [string]$row.target_code + "`n" + $entryMethod + "`n" + [string]$row.entry_code
  if (-not $sigEntry.ContainsKey($key)) {
    $sigEntry[$key] = @{ Sig = $sig; Code = [string]$row.target_code; Method = $entryMethod; Entry = [string]$row.entry_code; EntryName = [string]$row.entry_name; Count = 0; Last = $row.occurred_at }
  }
  $s = $sigEntry[$key]; $s.Count++
  if ($row.occurred_at -gt $s.Last) { $s.Last = $row.occurred_at }
  if ($s.EntryName -eq '' -and [string]$row.entry_name -ne '') { $s.EntryName = [string]$row.entry_name }
}
$dtSig = New-Object System.Data.DataTable
foreach ($c in 'signature','target_code','method','entry_code','entry_name') { [void]$dtSig.Columns.Add($c, [string]) }
[void]$dtSig.Columns.Add('sample_count', [int]); [void]$dtSig.Columns.Add('last_used_at', [datetime])
foreach ($s in $sigEntry.Values) { [void]$dtSig.Rows.Add($s.Sig, $s.Code, $s.Method, $s.Entry, $s.EntryName, $s.Count, $s.Last) }
[void](Invoke-RecoNonQueryInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Sql "TRUNCATE TABLE dbo.SignatureEntryMap")
Invoke-RecoBulkCopyInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Table $dtSig -TargetTable 'dbo.SignatureEntryMap'

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
[void](Invoke-RecoNonQueryInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Sql "TRUNCATE TABLE dbo.EngineeringTemplate")
Invoke-RecoBulkCopyInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Table $dtTmpl -TargetTable 'dbo.EngineeringTemplate'

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
  $sig = Get-QuantitySignature $g.Name $g.Unit $knownUnits
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
[void](Invoke-RecoNonQueryInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Sql "TRUNCATE TABLE dbo.SheetTemplateRow")
Invoke-RecoBulkCopyInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Table $dtSheet -TargetTable 'dbo.SheetTemplateRow'

$completionPrefix = "聚合完成"
if ($DryRun) {
  Write-Host ("存量尾部单位推断: " + $inferredRows + " 行 / " + $inferredGroups.Count + " 组 / " + $inferredNames.Count + " 个原始名称 / " + $inferredSignatureCount + " 个名称级签名 / " + $inferredUnits.Count + " 种单位")
  Write-Host ("潜在归并冲突(按办法隔离): 与显式单位样本合流 " + $inferredMergeCount + " 个签名/办法 / 同名多推断单位 " + $multiUnitInferenceCount + " 个签名/办法 / 归并后多组件框 " + $potentialMappingConflictCount + " 个签名/办法")
  $rebuildTransaction.Rollback()
  $completionPrefix = "聚合演练完成（已回滚）"
} else {
  $rebuildTransaction.Commit()
}
$rebuildTransaction.Dispose()
$rebuildTransaction = $null
Write-Host ($completionPrefix + ": QuotaBox " + $dtBox.Rows.Count + " / QuotaBoxTarget " + $dtTarget.Rows.Count + " / QuantityAlias " + $dtAlias.Rows.Count + " / SignatureBoxMap " + $dtMap.Rows.Count + " / QuantityFormulaRule " + $dtFormula.Rows.Count + " / SignatureEntryMap " + $dtSig.Rows.Count + " / EngineeringTemplate " + $dtTmpl.Rows.Count + " / SheetTemplateRow " + $dtSheet.Rows.Count)
}
catch {
  if ($null -ne $rebuildTransaction) { try { $rebuildTransaction.Rollback() } catch {} }
  throw
}
finally {
  if ($null -ne $rebuildTransaction) { $rebuildTransaction.Dispose() }
  $rebuildConnection.Dispose()
}
