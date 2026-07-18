# tools/RecoLearning/Import-ExcelLinks.ps1
# 收割 ExcelLinks\*.xml 绑定明细 → BindingLog。
# 条目号/办法通过回连项目库解析;项目库已删除时保留映射、条目留空并告警。
param(
  [string]$ExcelLinksDir = "D:\AI文件\自动预算\2024铁路工程云计价系统网络版V1.0\铁路工程云计价系统网络版V1.0\ExcelLinks"
)
. "$PSScriptRoot\Common.ps1"

$existing = @{}
foreach ($row in (Invoke-RecoQuery -Sql "SELECT event_hash FROM dbo.BindingLog WHERE source = 'import:excel-links'").Rows) { $existing[$row.event_hash] = $true }

$projCache = @{}   # 库名 → @{ Ok=bool; Method=string; Entries=@{ 条目序号 → @{Code;Name} } }
function Get-ProjectInfo {
  param([string]$Db)
  if ($projCache.ContainsKey($Db)) { return $projCache[$Db] }
  $info = @{ Ok = $false; Method = ''; Entries = @{} }
  try {
    $methodNo = [string](Invoke-RecoScalar -Database $Db -Sql "SELECT 编制办法文号 FROM 项目信息")
    if ($methodNo -match '2024') { $info.Method = '2024' }
    elseif ($methodNo -match '2020') { $info.Method = '2020' }
    else { $info.Method = $methodNo }
    $entries = Invoke-RecoQuery -Database $Db -Sql "SELECT 条目序号, 条目编号, 工程或费用项目名称 FROM 章节表"
    foreach ($e in $entries.Rows) { $info.Entries[[string]$e.条目序号] = @{ Code = [string]$e.条目编号; Name = [string]$e.工程或费用项目名称 } }
    $info.Ok = $true
  } catch {
    Write-Warning ("项目库 $Db 无法访问,条目将留空: " + $_.Exception.Message)
  }
  $projCache[$Db] = $info
  return $info
}

$dt = New-Object System.Data.DataTable
foreach ($c in 'source','method','project_id','entry_code','entry_name','quantity_name','quantity_unit','target_kind','target_code','target_name','target_unit','group_key','event_hash','extra') { [void]$dt.Columns.Add($c, [string]) }
[void]$dt.Columns.Add('occurred_at', [datetime])

$total = 0; $noName = 0; $skipped = 0; $entryMiss = 0
foreach ($file in (Get-ChildItem -LiteralPath $ExcelLinksDir -Filter '*.xml')) {
  [xml]$doc = Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8
  $links = @($doc.ExcelLinkStore.Links.ExcelQuotaLink)
  foreach ($link in $links) {
    if ($link -eq $null) { continue }
    $total++
    $quantityName = [string]$link.QuantityName
    if ([string]::IsNullOrWhiteSpace($quantityName)) { $noName++; continue }
    $projectId = [string]$link.ProjectId
    $db = ($projectId -split '\|')[-1]
    $hash = Get-Md5Hex ('excel-links|' + $projectId + '|' + $link.QuotaSequence + '|' + $link.QuotaCode + '|' + $link.UpdatedAt)
    if ($existing.ContainsKey($hash)) { $skipped++; continue }
    $info = Get-ProjectInfo -Db $db
    $entryCode = ''; $entryName = ''
    if ($info.Ok -and $info.Entries.ContainsKey([string]$link.ChapterSeq)) {
      $entryCode = $info.Entries[[string]$link.ChapterSeq].Code
      $entryName = $info.Entries[[string]$link.ChapterSeq].Name
    } else { $entryMiss++ }
    $occurred = [datetime]::MinValue
    if (-not [datetime]::TryParse([string]$link.UpdatedAt, [ref]$occurred)) { $occurred = $file.LastWriteTime }
    $groupKey = Get-Md5Hex ($projectId + '|' + $link.ExcelPath + '|' + $link.WorksheetName + '|' + $link.CellAddress)
    $extra = @{ quota_sequence = [string]$link.QuotaSequence; chapter_seq = [string]$link.ChapterSeq; total_no = [string]$link.TotalNo; cell = [string]$link.CellAddress; worksheet = [string]$link.WorksheetName; workbook = [System.IO.Path]::GetFileName([string]$link.ExcelPath); expression = [string]$link.Expression } | ConvertTo-Json -Compress
    [void]$dt.Rows.Add('import:excel-links', $info.Method, $db, $entryCode, $entryName, $quantityName, '', 'quota', [string]$link.QuotaCode, [string]$link.QuotaName, '', $groupKey, $hash, $extra, $occurred)
  }
}
Invoke-RecoBulkCopy -Table $dt -TargetTable 'dbo.BindingLog'
Write-Host ("ExcelLinks: 扫描 $total 条绑定,新增 " + $dt.Rows.Count + " 行,无工程量名跳过 $noName,已存在跳过 $skipped,条目未解析 $entryMiss")
