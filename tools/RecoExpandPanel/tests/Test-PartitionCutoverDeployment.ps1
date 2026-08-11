param(
  [Parameter(Mandatory = $true)]
  [string]$OutputDirectory
)

$ErrorActionPreference = 'Stop'

function Assert-True([bool]$Condition, [string]$Message) {
  if (-not $Condition) { throw $Message }
}

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $testDir '..\..\..')).Path
$output = if ([IO.Path]::IsPathRooted($OutputDirectory)) { [IO.Path]::GetFullPath($OutputDirectory) } else { [IO.Path]::GetFullPath((Join-Path $repoRoot $OutputDirectory)) }
Assert-True (-not (Test-Path -LiteralPath $output)) 'OutputDirectory must not exist before the test.'

$mappingBefore = @{}
foreach ($file in @(Get-ChildItem -LiteralPath $repoRoot -Recurse -File -Filter 'mapping-boxes.jsonl' -ErrorAction SilentlyContinue)) {
  $mappingBefore[$file.FullName] = (Get-FileHash -LiteralPath $file.FullName -Algorithm SHA256).Hash
}

& powershell.exe -ExecutionPolicy Bypass -File (Join-Path $repoRoot 'RecoQuotaRecommend\build.ps1') -BuildOnly -OutputDirectory $output
if ($LASTEXITCODE -ne 0) { throw "BuildOnly failed: $LASTEXITCODE" }

$names = @(Get-ChildItem -LiteralPath $output -Force | Select-Object -ExpandProperty Name | Sort-Object)
$expected = @('artifact-manifest.json','RecoExpandPanel.dll','RecoQuotaRecommend.dll') | Sort-Object
Assert-True ($names.Count -eq $expected.Count -and [string]::Join('|',$names) -eq [string]::Join('|',$expected)) 'BuildOnly output whitelist mismatch.'
$manifest = Get-Content -LiteralPath (Join-Path $output 'artifact-manifest.json') -Raw | ConvertFrom-Json
Assert-True (@($manifest.files).Count -eq 2) 'Manifest DLL count mismatch.'
foreach ($file in @($manifest.files)) {
  $actual = (Get-FileHash -LiteralPath ([string]$file.path) -Algorithm SHA256).Hash
  Assert-True ($actual -eq [string]$file.sha256) ("DLL hash mismatch: " + $file.name)
}
Assert-True (@($manifest.source_file_hashes).Count -gt 0) 'Manifest lacks source hashes.'
foreach ($source in @($manifest.source_file_hashes)) {
  $sourcePath = Join-Path $repoRoot ([string]$source.path)
  Assert-True ((Get-FileHash -LiteralPath $sourcePath -Algorithm SHA256).Hash -eq [string]$source.sha256) ("Source hash mismatch: " + $source.path)
}
Assert-True ($manifest.build_source_snapshot.mode -eq 'workspace-clean-copy' -and
  $manifest.build_source_snapshot.hash_source -eq 'snapshot-copy' -and
  $manifest.build_source_snapshot.removed_after_build) 'Clean-source snapshot evidence missing.'
Assert-True ($manifest.source_commit_role -eq 'base-head' -and $null -ne $manifest.source_worktree_dirty) 'Dirty-worktree source provenance is ambiguous.'
$leftovers = @(Get-ChildItem -LiteralPath (Split-Path -Parent $output) -Directory -Filter '.build-source-*' -ErrorAction SilentlyContinue)
Assert-True ($leftovers.Count -eq 0) 'Build source snapshot was not cleaned.'
foreach ($pair in $mappingBefore.GetEnumerator()) {
  Assert-True (Test-Path -LiteralPath $pair.Key -PathType Leaf) ("BuildOnly removed mapping file: " + $pair.Key)
  Assert-True ((Get-FileHash -LiteralPath $pair.Key -Algorithm SHA256).Hash -eq $pair.Value) ("BuildOnly changed mapping file: " + $pair.Key)
}
Write-Host 'PASS B18/T27 BuildOnly clean-source snapshot, source hashes, whitelist, and no mapping-file writes'
