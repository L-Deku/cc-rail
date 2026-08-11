$ErrorActionPreference = 'Stop'

function Assert-True([bool]$Condition, [string]$Message) {
  if (-not $Condition) { throw $Message }
}

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $testDir '..\..\..')).Path
$tool = Join-Path $repoRoot 'tools\RecoLearning\Repair-MappingBoxesIdentityConflicts.ps1'
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
$tempRoot = Join-Path $tempBase ('RecoRepairMapping-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($tempRoot)
$path = Join-Path $tempRoot 'mapping-boxes.jsonl'
$utf8NoBom = New-Object Text.UTF8Encoding($false)

$boxA = '{"record_type":"mapping_box","software_partition":"2020","box_id":"b1","target_kind":"quota","target_code":"DY-1","target_name":"Cable","target_unit":"m","quantity_name":"Cable","quantity_unit":"m","weight":"10","accepted_count":"1","corrected_count":"0","rejected_count":"0","last_used_at":"2026-08-07 10:00:00","custom":"A"}'
$boxB = '{"record_type":"mapping_box","software_partition":"2020","box_id":"b1","target_kind":"quota","target_code":"DY-1","target_name":"Cable","target_unit":"m","quantity_name":"Cable","quantity_unit":"m","weight":"20","accepted_count":"2","corrected_count":"1","rejected_count":"0","last_used_at":"2026-08-07 11:00:00","custom":"B"}'
$context = '{"record_type":"mapping_context","software_partition":"2020","method_no":"2020","box_id":"b1","target_kind":"quota","target_code":"DY-1","target_name":"Cable","target_unit":"m","quantity_name":"Cable","quantity_unit":"m","entry_code":"0101-01","entry_name":"Entry"}'
$unknown = '{"record_type":"future_type","raw":"keep"}'
$malformed = '{not-json'
[IO.File]::WriteAllText($path, ($unknown + "`n" + $boxA + "`n" + $boxB + "`n" + $context + "`n" + $context + "`n" + $malformed + "`n"), $utf8NoBom)

try {
  $before = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
  $analysis = & $tool -Analyze -Path $path
  Assert-True ($analysis.ConflictCount -eq 2) 'Analyze did not find both conflict kinds.'
  Assert-True ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -eq $before) 'Analyze changed the file.'
  Assert-True (Test-Path -LiteralPath $analysis.ReportPath -PathType Leaf) 'Analyze report missing.'
  Assert-True (Test-Path -LiteralPath $analysis.DecisionTemplate -PathType Leaf) 'Decision template missing.'

  $decisions = @(Import-Csv -LiteralPath $analysis.DecisionTemplate)
  foreach ($decision in $decisions) { $decision.keep_line = ([string]$decision.line_numbers).Split(',')[0] }
  $decisionFile = Join-Path $tempRoot 'decisions.csv'
  $decisions | Export-Csv -LiteralPath $decisionFile -NoTypeInformation -Encoding UTF8
  $running = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -in @('RejjNet2020','ReJJGSNet2024','ReJJQDNet2024') })
  if ($running.Count -ne 0) {
    $blocked = $false
    try { [void](& $tool -Repair -Path $path -DecisionFile $decisionFile -ExpectedSha256 $before) } catch { $blocked = $_.Exception.Message -like 'Repair requires all target software processes stopped*' }
    Assert-True $blocked 'Repair was not blocked while target software was running.'
    Assert-True ((Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash -eq $before) 'Blocked repair changed the file.'
    Write-Host 'PASS B17/T26 Analyze read-only and active-process Repair gate preserves hash'
    return
  }

  $repair = & $tool -Repair -Path $path -DecisionFile $decisionFile -ExpectedSha256 $before
  Assert-True ($repair.RepairedConflictCount -eq 2) 'Repair count mismatch.'
  Assert-True (Test-Path -LiteralPath $repair.BackupPath -PathType Leaf) 'Repair backup missing.'
  $saved = [IO.File]::ReadAllText($path, [Text.Encoding]::UTF8)
  Assert-True ($saved.StartsWith($unknown + "`n", [StringComparison]::Ordinal)) 'Unknown line moved or changed.'
  Assert-True ($saved.EndsWith($malformed + "`n", [StringComparison]::Ordinal)) 'Malformed line moved or changed.'
  $afterAnalysis = & $tool -Analyze -Path $path
  Assert-True ($afterAnalysis.ConflictCount -eq 0) 'Repair left conflicts.'
  Write-Host 'PASS B17/T26 analyze is read-only; explicit repair is complete, backed up, and revalidated'
}
finally {
  $resolved = [IO.Path]::GetFullPath($tempRoot)
  if (-not $resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase)) { throw 'Refusing to clean unexpected test directory.' }
  if ([IO.Directory]::Exists($resolved)) { [IO.Directory]::Delete($resolved, $true) }
}
