$ErrorActionPreference = 'Stop'

function Assert-True([bool]$Condition, [string]$Message) {
  if (-not $Condition) { throw $Message }
}

function Write-Utf8Json([string]$Path, $Value) {
  [IO.File]::WriteAllText($Path, ($Value | ConvertTo-Json -Depth 12), (New-Object Text.UTF8Encoding($true)))
}

function Write-JsonLines([string]$Path, [object[]]$Rows) {
  $lines = @($Rows | ForEach-Object { $_ | ConvertTo-Json -Compress })
  [IO.File]::WriteAllText($Path, (($lines -join "`n") + "`n"), (New-Object Text.UTF8Encoding($false)))
}

function Invoke-Exporter([string]$Script, [string]$RunId, [string]$FixtureData) {
  $startInfo = New-Object Diagnostics.ProcessStartInfo
  $startInfo.FileName = 'powershell.exe'
  $startInfo.Arguments = '-NoProfile -ExecutionPolicy Bypass -File "' + $Script + '" -RunId ' + $RunId + ' -TargetDatabase RecoLearning -FixtureDataPath "' + $FixtureData + '"'
  $startInfo.UseShellExecute = $false
  $startInfo.CreateNoWindow = $true
  $startInfo.RedirectStandardOutput = $true
  $startInfo.RedirectStandardError = $true
  $startInfo.EnvironmentVariables['RECO_PARTITION_EXPORT_TEST'] = '1'
  $process = New-Object Diagnostics.Process
  $process.StartInfo = $startInfo
  try {
    Assert-True $process.Start() 'Could not start D4 exporter fixture.'
    $stdout = $process.StandardOutput.ReadToEnd()
    $stderr = $process.StandardError.ReadToEnd()
    $process.WaitForExit()
    Assert-True ($process.ExitCode -eq 0) ('D4 exporter fixture failed: ' + $stdout + $stderr)
  }
  finally { $process.Dispose() }
}

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $testDir '..\..\..')).Path
$sourceScript = Join-Path $repoRoot 'tools\RecoLearning\Export-PartitionMappingBoxes.ps1'
$fixtureRoot = Join-Path $repoRoot ('obj\mapping-export-fixture-' + [Guid]::NewGuid().ToString('N'))
$fixtureBoundary = Join-Path $repoRoot 'obj\mapping-export-fixture-'

try {
  $toolDir = Join-Path $fixtureRoot 'tools\RecoLearning'
  $stateDir = Join-Path $toolDir 'migration-state'
  $runtime2020 = Join-Path $fixtureRoot '铁路基本建设工程投资控制系统2020网络版V0503021201'
  $runtime2024 = Join-Path $fixtureRoot '2024铁路工程云计价系统网络版V1.0\铁路工程云计价系统网络版V1.0'
  foreach ($directory in @($toolDir,$stateDir,$runtime2020,$runtime2024)) { [void][IO.Directory]::CreateDirectory($directory) }
  $fixtureScript = Join-Path $toolDir 'Export-PartitionMappingBoxes.ps1'
  [IO.File]::Copy($sourceScript,$fixtureScript,$false)

  [IO.File]::WriteAllText((Join-Path $runtime2020 'RejjNet2020.exe'),'fixture',[Text.Encoding]::UTF8)
  [IO.File]::WriteAllText((Join-Path $runtime2024 'RejjNet2020.exe'),'fixture',[Text.Encoding]::UTF8)
  [IO.File]::WriteAllText((Join-Path $runtime2024 'ReJJGSNet2024.exe'),'fixture',[Text.Encoding]::UTF8)

  foreach ($runtime in @($runtime2020,$runtime2024)) { [void][IO.Directory]::CreateDirectory((Join-Path $runtime 'RecoQuotaData')) }
  Write-JsonLines (Join-Path $runtime2020 'RecoQuotaData\quota-index.jsonl') @(
    [ordered]@{quota_code='DY-519';quota_name='电力电缆敷设';quota_unit='hm'},
    [ordered]@{quota_code='Q-OK';quota_name='有效定额';quota_unit='个'}
  )
  Write-JsonLines (Join-Path $runtime2024 'RecoQuotaData\quota-index.jsonl') @(
    [ordered]@{quota_code='DY-519';quota_name='10kV及以下电力电缆终端头制作安装';quota_unit='个'},
    [ordered]@{quota_code='Q-OK';quota_name='新版有效定额';quota_unit='套'}
  )
  Write-JsonLines (Join-Path $runtime2020 'RecoQuotaData\material-index.jsonl') @([ordered]@{material_code='109001005';material_name='2020电缆材料';material_unit='m'})
  Write-JsonLines (Join-Path $runtime2024 'RecoQuotaData\material-index.jsonl') @([ordered]@{material_code='109001005';material_name='2024电缆材料';material_unit='km'})

  $runId = [Guid]::NewGuid().ToString('N')
  $fixtureDataPath = Join-Path $fixtureRoot 'fixture-data.json'
  Write-Utf8Json $fixtureDataPath ([ordered]@{
    maps=@(
      [ordered]@{software_partition='2020';signature='SIG-A|';box_id='box-a';weight=30;accepted_count=2;corrected_count=1;rejected_count=0;last_used_at='2026-08-08T01:00:00'},
      [ordered]@{software_partition='2024';signature='SIG-A|';box_id='box-a';weight=20;accepted_count=2;corrected_count=0;rejected_count=0;last_used_at='2026-08-08T02:00:00'},
      [ordered]@{software_partition='2020';signature='SIG-BAD|';box_id='box-bad';weight=10;accepted_count=1;corrected_count=0;rejected_count=0;last_used_at='2026-08-08T03:00:00'},
      [ordered]@{software_partition='2024';signature='SIG-SF|';box_id='box-sf';weight=10;accepted_count=1;corrected_count=0;rejected_count=0;last_used_at='2026-08-08T04:00:00'}
    )
    aliases=@(
      [ordered]@{signature='SIG-A|';raw_name='电缆敷设';quantity_unit='m'},
      [ordered]@{signature='SIG-A|';raw_name='低烟无卤阻燃电缆';quantity_unit='m'},
      [ordered]@{signature='SIG-BAD|';raw_name='残缺组件';quantity_unit='个'},
      [ordered]@{signature='SIG-SF|';raw_name='设备购置费';quantity_unit='元'}
    )
    targets=@(
      [ordered]@{box_id='box-a';target_kind='quota';target_code='DY-519';target_name='跨版本共享旧名称';target_unit='错误单位'},
      [ordered]@{box_id='box-a';target_kind='quota';target_code='109001005*1.02';target_name='跨版本共享旧材料';target_unit='错误单位'},
      [ordered]@{box_id='box-bad';target_kind='quota';target_code='Q-OK';target_name='有效';target_unit='个'},
      [ordered]@{box_id='box-bad';target_kind='quota';target_code='01';target_name='不可回退';target_unit='个'},
      [ordered]@{box_id='box-sf';target_kind='aux';target_code='SF';target_name='设备购置费';target_unit='元'}
    )
  })

  $evidenceDir = Join-Path $stateDir ('partition-' + $runId + '-deployment')
  [void][IO.Directory]::CreateDirectory($evidenceDir)
  $isolationFiles = New-Object System.Collections.Generic.List[object]
  foreach ($runtime in @([pscustomobject]@{id='2020';path=$runtime2020},[pscustomobject]@{id='2024';path=$runtime2024})) {
    $sourcePath = Join-Path $runtime.path 'RecoQuotaData\mapping-boxes.jsonl'
    $archivePath = $sourcePath + '.pre-partition-' + $runId + '.bak'
    [IO.File]::WriteAllText($archivePath,'old-mapping-' + $runtime.id,[Text.Encoding]::UTF8)
    [void]$isolationFiles.Add([pscustomobject]@{runtime_id=$runtime.id;source_path=$sourcePath;archive_path=$archivePath;sha256=(Get-FileHash -LiteralPath $archivePath -Algorithm SHA256).Hash})
  }
  $evidencePath = Join-Path $evidenceDir 'deployment-evidence.json'
  Write-Utf8Json $evidencePath ([ordered]@{run_id=$runId;target_database='RecoLearning'})
  Write-Utf8Json (Join-Path $evidenceDir 'mapping-isolation.json') ([ordered]@{run_id=$runId;state='isolated';files=$isolationFiles.ToArray()})
  Write-Utf8Json (Join-Path $evidenceDir 'dll-deployment.json') ([ordered]@{run_id=$runId;state='deployed'})
  Write-Utf8Json (Join-Path $stateDir ('partition-' + $runId + '.json')) ([ordered]@{run_id=$runId;target_database='RecoLearning';state='consumed';deployment_evidence_path=$evidencePath;deployment_evidence_sha256=(Get-FileHash -LiteralPath $evidencePath -Algorithm SHA256).Hash})

  Invoke-Exporter $fixtureScript $runId $fixtureDataPath
  $output2020 = Join-Path $runtime2020 'RecoQuotaData\mapping-boxes.jsonl'
  $output2024 = Join-Path $runtime2024 'RecoQuotaData\mapping-boxes.jsonl'
  $hash2020 = (Get-FileHash -LiteralPath $output2020 -Algorithm SHA256).Hash
  $hash2024 = (Get-FileHash -LiteralPath $output2024 -Algorithm SHA256).Hash
  $rows2020 = @([IO.File]::ReadAllLines($output2020,[Text.Encoding]::UTF8) | ForEach-Object { $_ | ConvertFrom-Json })
  $rows2024 = @([IO.File]::ReadAllLines($output2024,[Text.Encoding]::UTF8) | ForEach-Object { $_ | ConvertFrom-Json })
  $shared2020 = @($rows2024 | Where-Object {$_.software_partition -eq '2020'})
  $shared2024 = @($rows2024 | Where-Object {$_.software_partition -eq '2024'})
  Assert-True ($rows2020.Count -eq 4 -and $rows2024.Count -eq 9 -and $shared2020.Count -eq 4 -and $shared2024.Count -eq 5) 'D4 shared-runtime row counts are incorrect.'
  Assert-True (@($rows2020 | Where-Object {$_.target_code -eq 'DY-519' -and $_.target_name -eq '电力电缆敷设' -and $_.target_unit -eq 'hm'}).Count -eq 2) '2020 same-code metadata was not resolved from the 2020 index.'
  Assert-True (@($shared2020 | Where-Object {$_.target_code -eq 'DY-519' -and $_.target_name -eq '电力电缆敷设' -and $_.target_unit -eq 'hm'}).Count -eq 2) 'Cloud runtime does not contain the 2020 executable partition.'
  Assert-True (@($shared2024 | Where-Object {$_.target_code -eq 'DY-519' -and $_.target_name -eq '10kV及以下电力电缆终端头制作安装' -and $_.target_unit -eq '个'}).Count -eq 2) '2024 same-code metadata was not resolved from the 2024 index.'
  Assert-True (@($rows2020 | Where-Object {$_.target_code -eq '109001005*1.02' -and $_.target_name -eq '2020电缆材料'}).Count -eq 2) 'Material multiplier code was not preserved or resolved by its base code.'
  Assert-True (@($rows2024 | Where-Object {$_.target_code -eq 'SF' -and $_.target_name -eq '设备购置费' -and $_.target_unit -eq '元'}).Count -eq 1) 'Auxiliary target context was not preserved.'
  Assert-True (@($rows2020 | Where-Object {$_.box_id -eq 'box-bad'}).Count -eq 0) 'Incomplete component box was partially exported.'
  Assert-True (@($rows2020 + $rows2024 | Where-Object {$_.record_type -ne 'mapping_box' -or $_.method_no -ne ''}).Count -eq 0) 'D4 emitted a context row or non-empty method_no.'
  $report = Get-Content -LiteralPath (Join-Path $evidenceDir 'mapping-export-report.json') -Raw -Encoding UTF8 | ConvertFrom-Json
  Assert-True ([int]$report.partitions.'2020'.isolated_box_count -eq 1 -and [int]$report.partitions.'2020'.isolated_line_count -eq 2) ('D4 isolation report counts are incorrect: boxes=' + [string]$report.partitions.'2020'.isolated_box_count + '; lines=' + [string]$report.partitions.'2020'.isolated_line_count + '; detail=' + ($report.partitions.'2020'.isolated_boxes | ConvertTo-Json -Compress))
  Assert-True ([string]::Join('|',@($report.runtime_files.'2024'.included_partitions)) -eq '2020|2024') 'Cloud runtime report does not declare both executable partitions.'

  [IO.File]::Delete($output2020)
  [IO.File]::Delete($output2024)
  [IO.File]::Delete((Join-Path $evidenceDir 'mapping-export-report.json'))
  Invoke-Exporter $fixtureScript $runId $fixtureDataPath
  Assert-True ((Get-FileHash -LiteralPath $output2020 -Algorithm SHA256).Hash -eq $hash2020) '2020 D4 output is not byte-deterministic.'
  Assert-True ((Get-FileHash -LiteralPath $output2024 -Algorithm SHA256).Hash -eq $hash2024) '2024 D4 output is not byte-deterministic.'
  Write-Host 'PASS T15 partition metadata, multiplier material, aux context, whole-box isolation, counts, and byte determinism'
}
finally {
  if ([IO.Directory]::Exists($fixtureRoot)) {
    $resolvedFixture = [IO.Path]::GetFullPath($fixtureRoot)
    if (-not $resolvedFixture.StartsWith($fixtureBoundary,[StringComparison]::OrdinalIgnoreCase)) { throw 'Fixture cleanup escaped the expected workspace boundary.' }
    $fixtureItem = Get-Item -LiteralPath $resolvedFixture -Force
    if (($fixtureItem.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Fixture cleanup target is a reparse point.' }
    [IO.Directory]::Delete($resolvedFixture,$true)
  }
}
