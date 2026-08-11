# tools/RecoLearning/Rebuild-Aggregates.ps1
# 由 BindingLog 流水全量重算推荐核心、公式、条目和模板聚合表。
# 可随时重跑;这是"定期整理"的入口。
# 聚合规则:同一 group_key 的多目标构成一个定额组样本;
#   weight = max(0, 10*accepted + 20*corrected - 10*rejected);
#   box_id 优先沿用 mapping-boxes 原始编号(extra.box_id),否则 auto- + 目标集合哈希前 16 位。
param([switch]$DryRun)
. "$PSScriptRoot\Common.ps1"

function Test-ClassifiedEntryCode {
  param([string]$EntryCode)
  return (Get-NormalizedLearningEntryCode $EntryCode) -ne ''
}

$rebuildConnection = New-Object System.Data.SqlClient.SqlConnection (Get-RecoConnectionString)
$rebuildTransaction = $null
try {
$rebuildConnection.Open()
$rebuildTransaction = $rebuildConnection.BeginTransaction([System.Data.IsolationLevel]::Serializable)
# 从读取流水到全部聚合替换提交始终持有流水表独占锁，防止增量绑定在快照与替换之间被吞掉。
$log = Invoke-RecoQueryInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Sql "SELECT occurred_at, quantity_name, quantity_unit, target_kind, target_code, target_name, target_unit, group_key, extra, method, software_partition, method_no, entry_code, entry_name FROM dbo.BindingLog WITH (TABLOCKX,HOLDLOCK) WHERE quantity_name <> ''"
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
    $groups[$gk] = @{ Name = [string]$row.quantity_name; Unit = [string]$row.quantity_unit; At = $row.occurred_at; Extra = [string]$row.extra; Targets = @{}; Method = ''; Partition = ''; MethodNo = ''; MissingPartition = $false; UnsafePartition = $false; UnsafeMethodNo = $false; UnsafeContextTargets = $false }
  }
  $g = $groups[$gk]
  if ($row.occurred_at -gt $g.At) { $g.At = $row.occurred_at }
  if ($g.Method -eq '' -and [string]$row.method -ne '') { $g.Method = Get-LearningMethodPartition ([string]$row.method) }
  $rowPartition = Get-NormalizedLearningSoftwarePartition ([string]$row.software_partition)
  if ($rowPartition -eq '') { $g.MissingPartition = $true }
  elseif ($g.Partition -eq '') { $g.Partition = $rowPartition }
  elseif ($g.Partition -ne $rowPartition) { $g.UnsafePartition = $true }
  $rowMethodNo = Get-NormalizedLearningMethodNo ([string]$row.method_no)
  if ($rowMethodNo -ne '' -and $g.MethodNo -eq '') { $g.MethodNo = $rowMethodNo }
  elseif ($rowMethodNo -ne '' -and $g.MethodNo -ne $rowMethodNo) { $g.UnsafeMethodNo = $true }
  $targetKind = ([string]$row.target_kind).Trim()
  $targetCode = ([string]$row.target_code).Trim()
  if ($targetKind -eq '') { $targetKind = if ($targetCode -match '^\d+$') { 'material' } else { 'quota' } }
  $baseKey = $targetKind.ToLowerInvariant() + ':' + $targetCode.ToUpperInvariant()
  $identityKey = Get-LearningTargetIdentityKey $targetKind $targetCode ([string]$row.target_name) ([string]$row.target_unit)
  if ($g.Targets.ContainsKey($baseKey)) {
    if ((Test-ContextSensitiveLearningCode $targetCode) -and
        -not [string]::Equals([string]$g.Targets[$baseKey].Identity, $identityKey, [System.StringComparison]::OrdinalIgnoreCase)) {
      $g.UnsafeContextTargets = $true
    }
  } else {
    $g.Targets[$baseKey] = @{ Kind = $targetKind; Code = $targetCode; Name = [string]$row.target_name; Unit = [string]$row.target_unit; Identity = $identityKey; EntryCode = (Get-NormalizedLearningEntryCode ([string]$row.entry_code)); EntryName = [string]$row.entry_name; Partition = $rowPartition; MethodNo = $rowMethodNo }
  }
  if ($g.Targets.ContainsKey($baseKey)) {
    $storedTarget = $g.Targets[$baseKey]
    $normalizedEntryCode = Get-NormalizedLearningEntryCode ([string]$row.entry_code)
    if ([string]::IsNullOrWhiteSpace([string]$storedTarget.EntryCode) -and $normalizedEntryCode -ne '') {
      $storedTarget.EntryCode = $normalizedEntryCode
      $storedTarget.EntryName = [string]$row.entry_name
      $storedTarget.Partition = $rowPartition
      $storedTarget.MethodNo = $rowMethodNo
    }
  }
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
function Test-PositiveLearningEvidence { param([string]$Extra)
  return ((Get-ExtraInt $Extra 'accepted_count' 1) + (Get-ExtraInt $Extra 'corrected_count' 0)) -gt 0
}

function Test-AggregateLearningGroupRecommendable { param($Group)
  $targets = @($Group.Targets.Values)
  if ($targets.Count -eq 0 -or $Group.UnsafeContextTargets -or $Group.UnsafePartition -or
      $Group.MissingPartition -or (Get-NormalizedLearningSoftwarePartition ([string]$Group.Partition)) -eq '') { return $false }
  if (-not (Test-AggregateLearningTargetSetRecommendable $targets)) { return $false }
  $incompleteContextTargets = @($targets | Where-Object {
    (Test-ContextSensitiveLearningCode ([string]$_.Code)) -and
      ([string]::IsNullOrWhiteSpace([string]$_.Name) -or [string]::IsNullOrWhiteSpace([string]$_.Unit))
  })
  return $incompleteContextTargets.Count -eq 0
}

function Test-AggregateLearningTargetSetRecommendable { param([System.Collections.IEnumerable]$Targets)
  $targets = @($Targets)
  if ($targets.Count -eq 0) { return $false }
  $hasPrimaryTarget = @($targets | Where-Object {
    $kind = [string]$_.Kind; if ($kind -eq '') { $kind = 'quota' }
    $baseCode = Get-LearningBaseTargetCode ([string]$_.Code)
    [string]::Equals($kind, 'quota', [System.StringComparison]::OrdinalIgnoreCase) -and
      @('SF','SH','SQ','ZLF','LF','YF','TLF','GF','JF','XGT1') -notcontains $baseCode
  }).Count -gt 0
  $allSf = @($targets | Where-Object {
    $kind = [string]$_.Kind; if ($kind -eq '') { $kind = 'quota' }
    [string]::Equals($kind, 'quota', [System.StringComparison]::OrdinalIgnoreCase) -and
      (Get-LearningBaseTargetCode ([string]$_.Code)) -eq 'SF'
  }).Count -eq $targets.Count
  return $hasPrimaryTarget -or $allSf
}

function Test-SfEntryConstraint { param($Group)
  foreach ($target in @($Group.Targets.Values)) {
    $isSf = (Get-LearningBaseTargetCode ([string]$target.Code)) -eq 'SF'
    $isEquipmentEntry = ([string]$target.EntryName).IndexOf('设备购置费', [System.StringComparison]::OrdinalIgnoreCase) -ge 0
    if ($isSf -ne $isEquipmentEntry) { return $false }
  }
  return $true
}

function Test-EngineeringScopeTarget { param($Target)
  $kind = [string]$Target.Kind; if ($kind -eq '') { $kind = 'quota' }
  if (-not [string]::Equals($kind, 'quota', [System.StringComparison]::OrdinalIgnoreCase)) { return $false }
  $baseCode = Get-LearningBaseTargetCode ([string]$Target.Code)
  if ($baseCode -eq 'SF') { return $true }
  return @('SH','SQ','ZLF','LF','YF','TLF','GF','JF','XGT1') -notcontains $baseCode
}

function Get-DistinctEngineeringScopeTargets {
  param([System.Collections.IEnumerable]$Targets)
  $seenEntries = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($target in @($Targets)) {
    if (-not (Test-EngineeringScopeTarget $target)) { continue }
    if ($seenEntries.Add([string]$target.EntryCode)) { Write-Output $target }
  }
}

function Get-AvailableAutoBoxId {
  param([string]$SetHash, [System.Collections.Generic.HashSet[string]]$UsedIds)
  for ($length = 16; $length -le $SetHash.Length; $length += 4) {
    $candidate = 'auto-' + $SetHash.Substring(0, $length)
    if (-not $UsedIds.Contains($candidate)) { return $candidate }
  }
  $suffix = 2
  do {
    $candidate = 'auto-' + $SetHash + '-' + $suffix
    $suffix++
  } while ($UsedIds.Contains($candidate))
  return $candidate
}

function Resolve-RebuildBoxIdCollisions {
  param([hashtable]$Boxes)
  $usedIds = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  $ordered = @($Boxes.GetEnumerator() | Sort-Object `
    @{ Expression = { if ([bool]$_.Value.IsPreferred) { 0 } else { 1 } } },
    @{ Expression = { [string]$_.Key } })
  $resolved = 0
  foreach ($pair in $ordered) {
    $currentId = [string]$pair.Value.Id
    if ($usedIds.Add($currentId)) { continue }
    $pair.Value.Id = Get-AvailableAutoBoxId ([string]$pair.Key) $usedIds
    [void]$usedIds.Add([string]$pair.Value.Id)
    $resolved++
  }
  return $resolved
}

$unsafeContextGroupCount = @($groups.Values | Where-Object { $_.UnsafeContextTargets }).Count
$skippedPureAuxiliaryGroupCount = @($groups.Values | Where-Object {
  -not $_.UnsafeContextTargets -and -not (Test-AggregateLearningTargetSetRecommendable @($_.Targets.Values))
}).Count
$skippedIncompleteContextGroupCount = @($groups.Values | Where-Object {
  -not $_.UnsafeContextTargets -and (Test-AggregateLearningTargetSetRecommendable @($_.Targets.Values)) -and
    -not (Test-AggregateLearningGroupRecommendable $_)
}).Count
$skippedSfConstraintGroupCount = @($groups.Values | Where-Object {
  -not $_.UnsafeContextTargets -and (Test-AggregateLearningGroupRecommendable $_) -and -not (Test-SfEntryConstraint $_)
}).Count
$aggregateGroupKeys = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($pair in $groups.GetEnumerator()) {
  if ((Test-AggregateLearningGroupRecommendable $pair.Value) -and (Test-SfEntryConstraint $pair.Value)) {
    [void]$aggregateGroupKeys.Add([string]$pair.Key)
  }
}
$aggregateGroups = @($aggregateGroupKeys | ForEach-Object { $groups[$_] })

# 2a) 第一遍:先为每个目标集合定死 box 编号(mapping-boxes 原始编号优先,取字典序最小者保证确定性)
$preferredHashes = @{}
foreach ($g in $aggregateGroups) {
  $identityKeys = @($g.Targets.Values | ForEach-Object { [string]$_.Identity } | Sort-Object)
  $setHash = Get-Md5Hex ($identityKeys -join ';')
  $g['SetHash'] = $setHash
  $g['ContainsContextSensitiveTarget'] = @($g.Targets.Values | Where-Object { Test-ContextSensitiveLearningCode ([string]$_.Code) }).Count -gt 0
  if ($g.ContainsContextSensitiveTarget) { continue }
  $preferredId = Get-ExtraText $g.Extra 'box_id'
  if (-not $preferredId) { continue }
  if (-not $preferredHashes.ContainsKey($preferredId)) {
    $preferredHashes[$preferredId] = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  }
  [void]$preferredHashes[$preferredId].Add($setHash)
}

$boxes = @{}    # target_set_hash → @{ Id; Targets }
foreach ($g in $aggregateGroups) {
  $setHash = $g.SetHash
  $g['EffectiveUnit'] = ([string]$g.Unit).Trim()
  if ($g.EffectiveUnit -eq '') { $g.EffectiveUnit = Get-InferredTrailingQuantityUnit $g.Name $knownUnits }
  $g['UnitInferred'] = ([string]$g.Unit).Trim() -eq '' -and $g.EffectiveUnit -ne ''
  $g['CanonicalName'] = Get-CanonicalQuantityName $g.Name $g.Unit $knownUnits
  $preferredId = Get-ExtraText $g.Extra 'box_id'
  if (-not $boxes.ContainsKey($setHash)) {
    $boxes[$setHash] = @{ Id = ('auto-' + $setHash.Substring(0, 16)); Targets = $g.Targets; IsPreferred = $false }
  }
  if ($preferredId -and -not $g.ContainsContextSensitiveTarget -and $preferredHashes[$preferredId].Count -eq 1) {
    $current = $boxes[$setHash].Id
    if ($current.StartsWith('auto-') -or [string]::CompareOrdinal($preferredId, $current) -lt 0) {
      $boxes[$setHash].Id = $preferredId
      $boxes[$setHash].IsPreferred = $true
    }
  }
}
$resolvedBoxIdCollisionCount = Resolve-RebuildBoxIdCollisions $boxes

# 2b) 第二遍:用定稿的 box 编号归并别名与签名映射
$aliases = @{}  # alias_hash → @{ Raw; Unit; Sig; Count; First; Last }
$maps = @{}     # software_partition + signature + box_id → 计数；30/101 普通关系在 2020 内共享
foreach ($g in $aggregateGroups) {
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
  $mapKey = $g.Partition + "`n" + $sig + "`n" + $boxId
  if (-not $maps.ContainsKey($mapKey)) { $maps[$mapKey] = @{ Partition = $g.Partition; Method = $method; Sig = $sig; BoxId = $boxId; Acc = 0; Cor = 0; Rej = 0; Last = $g.At } }
  $m = $maps[$mapKey]; $m.Acc += $acc; $m.Cor += $cor; $m.Rej += $rej
  if ($g.At -gt $m.Last) { $m.Last = $g.At }
}

$inferenceBySignature = @{}
foreach ($g in $aggregateGroups) {
  $sig = Get-QuantitySignature $g.CanonicalName $g.EffectiveUnit $knownUnits
  $methodSignature = $g.Partition + "`n" + $sig
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
  if (-not $aggregateGroupKeys.Contains([string]$row.group_key)) { continue }
  $extra = [string]$row.extra
  if (-not (Test-PositiveLearningEvidence $extra)) { continue }
  $template = Get-ExtraText $extra 'formula_template'
  $operandCount = Get-ExtraInt $extra 'formula_operand_count' 0
  if ($template -eq '' -or $operandCount -le 0) { continue }
  $sig = Get-QuantitySignature ([string]$row.quantity_name) ([string]$row.quantity_unit) $knownUnits
  $kind = ([string]$row.target_kind).Trim().ToLowerInvariant(); if ($kind -eq '') { $kind = 'quota' }
  $code = [string]$row.target_code
  $targetUnit = Get-ExtraText $extra 'formula_target_unit'; if ($targetUnit -eq '') { $targetUnit = [string]$row.target_unit }
  $targetUnit = Get-NormalizedLearningFormulaUnit $targetUnit
  $operands = @()
  $formulaMethod = Get-LearningMethodPartition ([string]$row.method)
  $formulaPartition = Get-NormalizedLearningSoftwarePartition ([string]$row.software_partition)
  $formulaMethodNo = Get-NormalizedLearningMethodNo ([string]$row.method_no)
  $formulaEntryCode = Get-NormalizedLearningEntryCode ([string]$row.entry_code)
  if ($formulaPartition -eq '' -or $formulaMethodNo -eq '' -or $formulaEntryCode -eq '') { continue }
  $ruleRaw = $sig + '|' + $kind + ':' + $code.ToUpperInvariant() + '|' + $targetUnit + '|' + $template + '|' + $formulaPartition + '|' + $formulaMethodNo + '|' + $formulaEntryCode
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
    $operandUnit = Get-NormalizedLearningFormulaUnit $operandUnit
    $operands += @{ Index = $i; Sig = $operandSig; Name = $operandName; Unit = $operandUnit }
    $ruleRaw += '|' + $operandSig + '@' + $operandUnit
  }
  if (-not $valid) { continue }
  # 旧 extra.formula_rule_hash 不含分区与办法号，切换时必须无条件按新布局重算。
  $ruleHash = Get-Md5Hex $ruleRaw
  if (-not $formulas.ContainsKey($ruleHash)) {
    $formulas[$ruleHash] = @{ Hash = $ruleHash; Sig = $sig; Kind = $kind; Code = $code; Unit = $targetUnit; Template = $template; Method = $formulaMethod; Partition = $formulaPartition; MethodNo = $formulaMethodNo; Entry = $formulaEntryCode; Count = 0; First = $row.occurred_at; Last = $row.occurred_at; Operands = $operands }
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
foreach ($c in 'software_partition','signature','box_id','method') { [void]$dtMap.Columns.Add($c, [string]) }
foreach ($c in 'weight','accepted_count','corrected_count','rejected_count') { [void]$dtMap.Columns.Add($c, [int]) }
[void]$dtMap.Columns.Add('last_used_at', [datetime])
foreach ($m in $maps.Values) {
  $weight = [Math]::Max(0, 10 * $m.Acc + 20 * $m.Cor - 10 * $m.Rej)
  [void]$dtMap.Rows.Add($m.Partition, $m.Sig, $m.BoxId, $m.Method, $weight, $m.Acc, $m.Cor, $m.Rej, $m.Last)
}
Invoke-RecoBulkCopyInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Table $dtMap -TargetTable 'dbo.SignatureBoxMap'

$dtFormula = New-Object System.Data.DataTable
foreach ($c in 'rule_hash','anchor_signature','target_kind','target_code','target_unit','formula_template','method','software_partition','method_no','entry_code') { [void]$dtFormula.Columns.Add($c, [string]) }
[void]$dtFormula.Columns.Add('sample_count', [int]); [void]$dtFormula.Columns.Add('first_seen', [datetime]); [void]$dtFormula.Columns.Add('last_seen', [datetime])
$dtOperand = New-Object System.Data.DataTable
[void]$dtOperand.Columns.Add('rule_hash', [string]); [void]$dtOperand.Columns.Add('operand_index', [int])
foreach ($c in 'operand_signature','operand_name','operand_unit') { [void]$dtOperand.Columns.Add($c, [string]) }
foreach ($f in $formulas.Values) {
  [void]$dtFormula.Rows.Add($f.Hash, $f.Sig, $f.Kind, $f.Code, $f.Unit, $f.Template, $f.Method, $f.Partition, $f.MethodNo, $f.Entry, $f.Count, $f.First, $f.Last)
  foreach ($o in $f.Operands) { [void]$dtOperand.Rows.Add($f.Hash, $o.Index, $o.Sig, $o.Name, $o.Unit) }
}
[void](Invoke-RecoNonQueryInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Sql "TRUNCATE TABLE dbo.QuantityFormulaOperand; TRUNCATE TABLE dbo.QuantityFormulaRule")
Invoke-RecoBulkCopyInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Table $dtFormula -TargetTable 'dbo.QuantityFormulaRule'
Invoke-RecoBulkCopyInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Table $dtOperand -TargetTable 'dbo.QuantityFormulaOperand'

# 4) 签名级条目证据:某工程量(签名)+某定额 历史上实际放过的条目,按办法分组计数。
$sigEntry = @{}
foreach ($row in $log.Rows) {
  if (-not $aggregateGroupKeys.Contains([string]$row.group_key)) { continue }
  if (-not (Test-PositiveLearningEvidence ([string]$row.extra))) { continue }
  $entryPartition = Get-NormalizedLearningSoftwarePartition ([string]$row.software_partition)
  $entryMethodNo = Get-NormalizedLearningMethodNo ([string]$row.method_no)
  $entryCode = Get-NormalizedLearningEntryCode ([string]$row.entry_code)
  if ($entryPartition -eq '' -or $entryMethodNo -eq '' -or $entryCode -eq '' -or
      [string]$row.target_kind -ne 'quota' -or [string]$row.target_code -eq '') { continue }
  $sig = Get-QuantitySignature ([string]$row.quantity_name) ([string]$row.quantity_unit) $knownUnits
  $entryMethod = Get-LearningMethodPartition ([string]$row.method)
  $key = $entryPartition + "`n" + $entryMethodNo + "`n" + $sig + "`n" + [string]$row.target_code + "`n" + $entryCode
  if (-not $sigEntry.ContainsKey($key)) {
    $sigEntry[$key] = @{ Partition = $entryPartition; MethodNo = $entryMethodNo; Sig = $sig; Code = [string]$row.target_code; Method = $entryMethod; Entry = $entryCode; EntryName = [string]$row.entry_name; Count = 0; Last = $row.occurred_at }
  }
  $s = $sigEntry[$key]; $s.Count++
  if ($row.occurred_at -gt $s.Last) { $s.Last = $row.occurred_at }
  if ($s.EntryName -eq '' -and [string]$row.entry_name -ne '') { $s.EntryName = [string]$row.entry_name }
}
$dtSig = New-Object System.Data.DataTable
foreach ($c in 'software_partition','method_no','signature','target_code','method','entry_code','entry_name') { [void]$dtSig.Columns.Add($c, [string]) }
[void]$dtSig.Columns.Add('sample_count', [int]); [void]$dtSig.Columns.Add('last_used_at', [datetime])
foreach ($s in $sigEntry.Values) { [void]$dtSig.Rows.Add($s.Partition, $s.MethodNo, $s.Sig, $s.Code, $s.Method, $s.Entry, $s.EntryName, $s.Count, $s.Last) }
[void](Invoke-RecoNonQueryInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Sql "TRUNCATE TABLE dbo.SignatureEntryMap")
Invoke-RecoBulkCopyInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Table $dtSig -TargetTable 'dbo.SignatureEntryMap'

# 5) 工程模板归集:条目前缀(前2位)=工程类型,统计每个工程类型下条目与定额组的共现。
$tmpl = @{}
foreach ($g in $aggregateGroups) {
  if (-not (Test-PositiveLearningEvidence ([string]$g.Extra))) { continue }
  if ($g.UnsafeMethodNo -or (Get-NormalizedLearningMethodNo ([string]$g.MethodNo)) -eq '') { continue }
  $boxId = $boxes[$g.SetHash].Id
  $scopeTargets = @(Get-DistinctEngineeringScopeTargets $g.Targets.Values)
  foreach ($scopeTarget in $scopeTargets) {
    $entryCode = Get-NormalizedLearningEntryCode ([string]$scopeTarget.EntryCode)
    if (-not (Test-ClassifiedEntryCode $entryCode)) { continue }
    if ((Get-NormalizedLearningSoftwarePartition ([string]$scopeTarget.Partition)) -ne $g.Partition -or
        (Get-NormalizedLearningMethodNo ([string]$scopeTarget.MethodNo)) -ne $g.MethodNo) { continue }
    $prefix = $entryCode.Substring(0, 2)
    $key = $g.Partition + "`n" + $g.MethodNo + "`n" + $prefix + "`n" + $entryCode + "`n" + $boxId
    if (-not $tmpl.ContainsKey($key)) {
      $tmpl[$key] = @{ Partition = $g.Partition; MethodNo = $g.MethodNo; Method = $g.Method; Prefix = $prefix; Entry = $entryCode; BoxId = $boxId; Count = 0; Last = $g.At }
    }
    $t = $tmpl[$key]; $t.Count++
    if ($g.At -gt $t.Last) { $t.Last = $g.At }
  }
}
$dtTmpl = New-Object System.Data.DataTable
foreach ($c in 'software_partition','method_no','method','engineering_type','entry_code','box_id') { [void]$dtTmpl.Columns.Add($c, [string]) }
[void]$dtTmpl.Columns.Add('sample_count', [int]); [void]$dtTmpl.Columns.Add('last_seen', [datetime])
foreach ($t in $tmpl.Values) { [void]$dtTmpl.Rows.Add($t.Partition, $t.MethodNo, $t.Method, $t.Prefix, $t.Entry, $t.BoxId, $t.Count, $t.Last) }
[void](Invoke-RecoNonQueryInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Sql "TRUNCATE TABLE dbo.EngineeringTemplate")
Invoke-RecoBulkCopyInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Table $dtTmpl -TargetTable 'dbo.EngineeringTemplate'

# 6) SheetTemplateRow 含条目却没有办法号，本次切换只清空且不再回填。
$dtSheet = New-Object System.Data.DataTable
[void](Invoke-RecoNonQueryInTransaction -Connection $rebuildConnection -Transaction $rebuildTransaction -Sql "TRUNCATE TABLE dbo.SheetTemplateRow")

$completionPrefix = "聚合完成"
Write-Host ("跳过同码多义组件: " + $unsafeContextGroupCount + " 组；跳过纯辅助组件: " + $skippedPureAuxiliaryGroupCount + " 组；跳过身份不完整辅助组件: " + $skippedIncompleteContextGroupCount + " 组；跳过违反 SF 双向条目约束组件: " + $skippedSfConstraintGroupCount + " 组；解决 box_id 冲突: " + $resolvedBoxIdCollisionCount + " 个；保留原始 BindingLog")
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
