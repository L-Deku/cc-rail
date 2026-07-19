# tools/RecoLearning/Import-ExcelLinks.ps1
# 收割 ExcelLinks\*.xml 绑定明细 → BindingLog。
# ProjectId 形如 "192.168.2.13,1433|Reco20260622092010550":按其中的服务器地址连对应项目库,
# 解析 办法(项目信息)、条目(章节表 by 条目序号)、定额单位(定额输入 by 定额序号)。
# 幂等:已存在的行若缺 条目/定额单位 而本次解析到了,则补齐(UPDATE);否则跳过。
param(
  [string]$ExcelLinksDir = "D:\AI文件\自动预算\2024铁路工程云计价系统网络版V1.0\铁路工程云计价系统网络版V1.0\ExcelLinks"
)
. "$PSScriptRoot\Common.ps1"

$existing = @{}   # event_hash → @{ Id; EntryCode; TargetUnit }
foreach ($row in (Invoke-RecoQuery -Sql "SELECT id, event_hash, entry_code, target_unit FROM dbo.BindingLog WHERE source = 'import:excel-links'").Rows) {
  $existing[$row.event_hash] = @{ Id = [long]$row.id; EntryCode = [string]$row.entry_code; TargetUnit = [string]$row.target_unit }
}

$projCache = @{}   # "server|库名" → @{ Ok; Method; Entries(条目序号→@{Code;Name}); QuotaUnits(定额序号→单位) }
function Get-ProjectInfo {
  param([string]$Server, [string]$Db)
  $cacheKey = "$Server|$Db"
  if ($projCache.ContainsKey($cacheKey)) { return $projCache[$cacheKey] }
  $info = @{ Ok = $false; Method = ''; Entries = @{}; QuotaUnits = @{} }
  try {
    $methodNo = [string](Invoke-RecoScalar -Server $Server -Database $Db -Sql "SELECT TOP 1 编制办法文号 FROM 项目信息")
    if ($methodNo -match '2024') { $info.Method = '2024' }
    elseif ($methodNo -match '2020|30号文') { $info.Method = '2020' }   # 30号文=2020办法文号
    else { $info.Method = $methodNo }
    $entries = Invoke-RecoQuery -Server $Server -Database $Db -Sql "SELECT 条目序号, 条目编号, 工程或费用项目名称 FROM 章节表"
    foreach ($e in $entries.Rows) { $info.Entries[[string]$e.条目序号] = @{ Code = [string]$e.条目编号; Name = [string]$e.工程或费用项目名称 } }
    $units = Invoke-RecoQuery -Server $Server -Database $Db -Sql "SELECT 定额序号, 单位 FROM 定额输入"
    foreach ($u in $units.Rows) { $info.QuotaUnits[[string]$u.定额序号] = ([string]$u.单位).Trim() }
    $info.Ok = $true
  } catch {
    Write-Warning ("项目库 $Server/$Db 无法访问,条目与单位将留空: " + $_.Exception.Message)
  }
  $projCache[$cacheKey] = $info
  return $info
}

$dt = New-Object System.Data.DataTable
foreach ($c in 'source','method','project_id','entry_code','entry_name','quantity_name','quantity_unit','target_kind','target_code','target_name','target_unit','group_key','event_hash','extra') { [void]$dt.Columns.Add($c, [string]) }
[void]$dt.Columns.Add('occurred_at', [datetime])

$updates = New-Object System.Collections.Generic.List[object]
$total = 0; $noName = 0; $skipped = 0; $entryMiss = 0; $updated = 0
foreach ($file in (Get-ChildItem -LiteralPath $ExcelLinksDir -Filter '*.xml')) {
  [xml]$doc = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
  $links = @($doc.ExcelLinkStore.Links.ExcelQuotaLink)
  foreach ($link in $links) {
    if ($link -eq $null) { continue }
    $total++
    $quantityName = [string]$link.QuantityName
    if ([string]::IsNullOrWhiteSpace($quantityName)) { $noName++; continue }
    $projectId = [string]$link.ProjectId
    $parts = $projectId -split '\|'
    $db = $parts[-1]
    $server = if ($parts.Count -ge 2 -and $parts[0].Trim()) { $parts[0].Trim() } else { $script:RecoServer }
    $hash = Get-Md5Hex ('excel-links|' + $projectId + '|' + $link.QuotaSequence + '|' + $link.QuotaCode + '|' + $link.UpdatedAt)
    $info = Get-ProjectInfo -Server $server -Db $db
    $entryCode = ''; $entryName = ''
    if ($info.Ok -and $info.Entries.ContainsKey([string]$link.ChapterSeq)) {
      $entryCode = $info.Entries[[string]$link.ChapterSeq].Code
      $entryName = $info.Entries[[string]$link.ChapterSeq].Name
    } else { $entryMiss++ }
    $quotaUnit = ''
    if ($info.Ok -and $info.QuotaUnits.ContainsKey([string]$link.QuotaSequence)) { $quotaUnit = $info.QuotaUnits[[string]$link.QuotaSequence] }

    if ($existing.ContainsKey($hash)) {
      $old = $existing[$hash]
      $needEntry = ($old.EntryCode -eq '' -and $entryCode -ne '')
      $needUnit = ($old.TargetUnit -eq '' -and $quotaUnit -ne '')
      if ($needEntry -or $needUnit) {
        $updates.Add(@{ Id = $old.Id; EntryCode = $entryCode; EntryName = $entryName; TargetUnit = $quotaUnit; Method = $info.Method; NeedEntry = $needEntry; NeedUnit = $needUnit })
      } else { $skipped++ }
      continue
    }

    $occurred = [datetime]::MinValue
    if (-not [datetime]::TryParse([string]$link.UpdatedAt, [ref]$occurred)) { $occurred = $file.LastWriteTime }
    $groupKey = Get-Md5Hex ($projectId + '|' + $link.ExcelPath + '|' + $link.WorksheetName + '|' + $link.CellAddress)
    $extra = @{ quota_sequence = [string]$link.QuotaSequence; chapter_seq = [string]$link.ChapterSeq; total_no = [string]$link.TotalNo; cell = [string]$link.CellAddress; worksheet = [string]$link.WorksheetName; workbook = [System.IO.Path]::GetFileName([string]$link.ExcelPath); expression = [string]$link.Expression } | ConvertTo-Json -Compress
    [void]$dt.Rows.Add('import:excel-links', $info.Method, $db, $entryCode, $entryName, $quantityName, '', 'quota', [string]$link.QuotaCode, [string]$link.QuotaName, $quotaUnit, $groupKey, $hash, $extra, $occurred)
  }
}
Invoke-RecoBulkCopy -Table $dt -TargetTable 'dbo.BindingLog'

foreach ($u in $updates) {
  $set = @()
  $params = @{ id = $u.Id }
  if ($u.NeedEntry) { $set += "entry_code = @ec, entry_name = @en"; $params['ec'] = $u.EntryCode; $params['en'] = $u.EntryName; if ($u.Method) { $set += "method = @m"; $params['m'] = $u.Method } }
  if ($u.NeedUnit) { $set += "target_unit = @tu"; $params['tu'] = $u.TargetUnit }
  [void](Invoke-RecoNonQuery -Sql ("UPDATE dbo.BindingLog SET " + ($set -join ', ') + " WHERE id = @id") -Parameters $params)
  $updated++
}
Write-Host ("ExcelLinks[$ExcelLinksDir]: 扫描 $total 条,新增 " + $dt.Rows.Count + " 行,补齐 $updated 行,无工程量名跳过 $noName,已存在跳过 $skipped,条目未解析 $entryMiss")
