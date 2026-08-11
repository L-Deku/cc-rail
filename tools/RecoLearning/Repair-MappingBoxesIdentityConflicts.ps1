param(
  [Parameter(Mandatory = $true)]
  [string]$Path,
  [switch]$Analyze,
  [switch]$Repair,
  [string]$DecisionFile = '',
  [string]$ExpectedSha256 = ''
)

$ErrorActionPreference = 'Stop'
. "$PSScriptRoot\Common.ps1"

if ($Analyze -and $Repair) { throw 'Choose either -Analyze or -Repair.' }
if (-not $Analyze -and -not $Repair) { $Analyze = $true }
$resolvedPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
if (-not [IO.File]::Exists($resolvedPath)) { throw "File does not exist: $resolvedPath" }

function Get-FileSha256([string]$FilePath) {
  return (Get-FileHash -LiteralPath $FilePath -Algorithm SHA256).Hash
}

function Get-PropertyText($Row, [string]$Name) {
  if ($null -eq $Row) { return '' }
  $property = $Row.PSObject.Properties[$Name]
  if ($null -eq $property -or $null -eq $property.Value) { return '' }
  return [string]$property.Value
}

function Normalize-QuantitySignature([string]$Name) {
  $source = ([string]$Name).Normalize([Text.NormalizationForm]::FormKC).ToUpperInvariant().Replace([char]0x0424, [char]0x03A6).Replace([char]0x00D7, 'X')
  return (($source.ToCharArray() | Where-Object { -not [char]::IsWhiteSpace($_) }) -join '') + '|'
}

function Normalize-Unit([string]$Unit) {
  return ([string]$Unit).Normalize([Text.NormalizationForm]::FormKC).Trim().ToUpperInvariant().Replace(' ', '')
}

function Test-ContextSensitiveCode([string]$Code) {
  $baseCode = ([string]$Code).Trim().ToUpperInvariant()
  $suffix = $baseCode.IndexOfAny([char[]]@('*','/'))
  if ($suffix -ge 0) { $baseCode = $baseCode.Substring(0, $suffix) }
  return $baseCode -in @('SF','SH','SQ','ZLF','LF','YF','TLF','GF','JF','XGT1')
}

function Get-TargetIdentity($Row) {
  $code = (Get-PropertyText $Row 'target_code').Trim().ToUpperInvariant()
  if ($code.Length -eq 0) { return '' }
  $kind = (Get-PropertyText $Row 'target_kind').Trim().ToLowerInvariant()
  if ($kind.Length -eq 0) {
    if ($code -match '^\d+$') { $kind = 'material' } else { $kind = 'quota' }
  }
  $identity = $kind + ':' + $code
  if (Test-ContextSensitiveCode $code) {
    $identity += '|' + (Normalize-QuantitySignature (Get-PropertyText $Row 'target_name')).TrimEnd('|') + '|' + (Normalize-Unit (Get-PropertyText $Row 'target_unit'))
  }
  return $identity
}

function Get-BoxIdentity($Row) {
  if ((Get-PropertyText $Row 'record_type') -ne 'mapping_box') { return '' }
  $partition = Get-NormalizedLearningSoftwarePartition (Get-PropertyText $Row 'software_partition')
  $boxId = (Get-PropertyText $Row 'box_id').Trim().ToUpperInvariant()
  $target = Get-TargetIdentity $Row
  $quantity = Normalize-QuantitySignature (Get-PropertyText $Row 'quantity_name')
  if ($partition.Length -eq 0 -or $boxId.Length -eq 0 -or $target.Length -eq 0 -or $quantity -eq '|') { return '' }
  return $partition + "`n" + $boxId + "`n" + $target + "`n" + $quantity
}

function Get-ContextIdentity($Row) {
  if ((Get-PropertyText $Row 'record_type') -ne 'mapping_context') { return '' }
  $partition = Get-NormalizedLearningSoftwarePartition (Get-PropertyText $Row 'software_partition')
  $methodNo = Get-NormalizedLearningMethodNo (Get-PropertyText $Row 'method_no')
  $entryCode = Get-NormalizedLearningEntryCode (Get-PropertyText $Row 'entry_code')
  $boxId = (Get-PropertyText $Row 'box_id').Trim().ToUpperInvariant()
  $target = Get-TargetIdentity $Row
  $quantity = Normalize-QuantitySignature (Get-PropertyText $Row 'quantity_name')
  if ($partition.Length -eq 0 -or $methodNo.Length -eq 0 -or $entryCode.Length -eq 0 -or
      $boxId.Length -eq 0 -or $target.Length -eq 0 -or $quantity -eq '|') { return '' }
  $identity = $partition + "`n" + $methodNo + "`n" + $boxId + "`n" + $target + "`n" + $quantity + "`n" + $entryCode
  $formulaHash = (Get-PropertyText $Row 'formula_rule_hash').Trim().ToUpperInvariant()
  if ($formulaHash.Length -eq 0) { return $identity }
  return $identity + "`n" + $formulaHash
}

function Read-RawDocument([string]$FilePath) {
  [byte[]]$bytes = [IO.File]::ReadAllBytes($FilePath)
  $preambleLength = 0
  if ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) {
    $encoding = New-Object Text.UTF8Encoding($true, $true); $preambleLength = 3
  } elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) {
    $encoding = New-Object Text.UnicodeEncoding($false, $true, $true); $preambleLength = 2
  } elseif ($bytes.Length -ge 2 -and $bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF) {
    $encoding = New-Object Text.UnicodeEncoding($true, $true, $true); $preambleLength = 2
  } else {
    $encoding = New-Object Text.UTF8Encoding($false, $true)
  }
  $text = $encoding.GetString($bytes, $preambleLength, $bytes.Length - $preambleLength)
  $lines = New-Object System.Collections.Generic.List[object]
  $start = 0
  for ($index = 0; $index -lt $text.Length; $index++) {
    if ($text[$index] -ne "`r" -and $text[$index] -ne "`n") { continue }
    $content = $text.Substring($start, $index - $start)
    $terminator = [string]$text[$index]
    if ($text[$index] -eq "`r" -and $index + 1 -lt $text.Length -and $text[$index + 1] -eq "`n") { $terminator = "`r`n"; $index++ }
    $row = $null
    try { if (-not [string]::IsNullOrWhiteSpace($content)) { $row = $content | ConvertFrom-Json -ErrorAction Stop } } catch { $row = $null }
    $lines.Add([pscustomobject]@{ LineNumber = $lines.Count + 1; Content = $content; Terminator = $terminator; Row = $row; Removed = $false })
    $start = $index + 1
  }
  if ($start -lt $text.Length) {
    $content = $text.Substring($start)
    $row = $null
    try { if (-not [string]::IsNullOrWhiteSpace($content)) { $row = $content | ConvertFrom-Json -ErrorAction Stop } } catch { $row = $null }
    $lines.Add([pscustomobject]@{ LineNumber = $lines.Count + 1; Content = $content; Terminator = ''; Row = $row; Removed = $false })
  }
  return [pscustomobject]@{ Encoding = $encoding; Lines = $lines; Bytes = $bytes }
}

$boxKnown = [Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
foreach ($name in @('record_type','software_partition','box_id','target_kind','target_code','target_name','target_unit','quantity_name','quantity_unit','weight','accepted_count','corrected_count','rejected_count','last_used_at','method','method_no','project_id','entry_codes','entry_code','entry_name')) { [void]$boxKnown.Add($name) }

function Test-BoxKnownField([string]$Name) {
  return $boxKnown.Contains($Name) -or $Name.StartsWith('formula_', [StringComparison]::OrdinalIgnoreCase)
}

function Get-Conflicts($Document) {
  $conflicts = New-Object System.Collections.Generic.List[object]
  $contexts = @($Document.Lines | Where-Object { $null -ne $_.Row } | Group-Object { Get-ContextIdentity $_.Row } | Where-Object { $_.Name.Length -gt 0 -and $_.Count -gt 1 })
  foreach ($group in $contexts) {
    $lineNumbers = @($group.Group | ForEach-Object { $_.LineNumber })
    $identity = $group.Name
    $conflicts.Add([pscustomobject]@{ Kind='DuplicateContextIdentity'; Identity=$identity; Lines=$lineNumbers; Summary='duplicate context identity' })
  }
  $boxes = @($Document.Lines | Where-Object { $null -ne $_.Row } | Group-Object { Get-BoxIdentity $_.Row } | Where-Object { $_.Name.Length -gt 0 -and $_.Count -gt 1 })
  foreach ($group in $boxes) {
    $unknownValues = @{}
    $ambiguous = New-Object System.Collections.Generic.List[string]
    foreach ($line in $group.Group) {
      foreach ($property in $line.Row.PSObject.Properties) {
        if (Test-BoxKnownField $property.Name) { continue }
        if ($unknownValues.ContainsKey($property.Name) -and [string]$unknownValues[$property.Name] -cne [string]$property.Value) { [void]$ambiguous.Add($property.Name) }
        $unknownValues[$property.Name] = [string]$property.Value
      }
    }
    if ($ambiguous.Count -eq 0) { continue }
    $conflicts.Add([pscustomobject]@{ Kind='AmbiguousBoxUnknownFields'; Identity=$group.Name; Lines=@($group.Group | ForEach-Object { $_.LineNumber }); Summary=(($ambiguous | Sort-Object -Unique) -join ',') })
  }
  return $conflicts
}

function Get-ConflictId([string]$Kind, [string]$Identity) {
  $sha = [Security.Cryptography.SHA256]::Create()
  try { return ([BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($Kind + "`n" + $Identity))).Replace('-', '')) }
  finally { $sha.Dispose() }
}

function Write-AnalysisArtifacts([string]$FilePath, [string]$Hash, $Conflicts) {
  $directory = Join-Path (Split-Path -Parent $FilePath) 'diagnostics'
  [void][IO.Directory]::CreateDirectory($directory)
  $prefix = $Hash.Substring(0, [Math]::Min(12, $Hash.Length))
  $reportPath = Join-Path $directory ('mapping-boxes-identity-analysis-' + $prefix + '.json')
  $decisionPath = Join-Path $directory ('mapping-boxes-identity-decisions-' + $prefix + '.csv')
  $report = [ordered]@{ file_path=$FilePath; file_sha256=$Hash; conflict_count=$Conflicts.Count; conflicts=@($Conflicts | ForEach-Object { [ordered]@{ conflict_id=(Get-ConflictId $_.Kind $_.Identity); conflict_kind=$_.Kind; identity=$_.Identity; line_numbers=$_.Lines; summary=$_.Summary } }) }
  [IO.File]::WriteAllText($reportPath, ($report | ConvertTo-Json -Depth 8), (New-Object Text.UTF8Encoding($true)))
  @($Conflicts | ForEach-Object { [pscustomobject]@{ conflict_id=(Get-ConflictId $_.Kind $_.Identity); conflict_kind=$_.Kind; identity=$_.Identity; line_numbers=($_.Lines -join ','); keep_line='' } }) | Export-Csv -LiteralPath $decisionPath -NoTypeInformation -Encoding UTF8
  return [pscustomobject]@{ ReportPath=$reportPath; DecisionPath=$decisionPath }
}

function Assert-SoftwareStopped {
  $running = @(Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.ProcessName -in @('RejjNet2020','ReJJGSNet2024','ReJJQDNet2024') })
  if ($running.Count -ne 0) { throw ('Repair requires all target software processes stopped: ' + (($running | ForEach-Object { $_.ProcessName + ':' + $_.Id }) -join ',')) }
}

function Write-DocumentAtomic($Document, [string]$FilePath) {
  $builder = New-Object Text.StringBuilder
  foreach ($line in $Document.Lines) { if (-not $line.Removed) { [void]$builder.Append($line.Content); [void]$builder.Append($line.Terminator) } }
  [byte[]]$body = $Document.Encoding.GetBytes($builder.ToString())
  [byte[]]$preamble = $Document.Encoding.GetPreamble()
  [byte[]]$output = New-Object byte[] ($preamble.Length + $body.Length)
  [Array]::Copy($preamble, 0, $output, 0, $preamble.Length)
  [Array]::Copy($body, 0, $output, $preamble.Length, $body.Length)
  $directory = Split-Path -Parent $FilePath
  $temp = Join-Path $directory ((Split-Path -Leaf $FilePath) + '.' + [Guid]::NewGuid().ToString('N') + '.tmp')
  $previous = Join-Path $directory ((Split-Path -Leaf $FilePath) + '.' + [Guid]::NewGuid().ToString('N') + '.previous')
  try {
    [IO.File]::WriteAllBytes($temp, $output)
    [IO.File]::Replace($temp, $FilePath, $previous, $true)
  } finally {
    if ([IO.File]::Exists($temp)) { [IO.File]::Delete($temp) }
    if ([IO.File]::Exists($previous)) { [IO.File]::Delete($previous) }
  }
}

$mutex = New-Object Threading.Mutex($false, 'RecoQuotaData.mapping-boxes.lock')
$acquired = $false
try {
  try { $acquired = $mutex.WaitOne(5000) } catch [Threading.AbandonedMutexException] { $acquired = $true }
  if (-not $acquired) { throw 'mapping-boxes lock timeout.' }
  $beforeHash = Get-FileSha256 $resolvedPath
  $document = Read-RawDocument $resolvedPath
  $conflicts = @(Get-Conflicts $document)

  if ($Analyze) {
    $artifacts = Write-AnalysisArtifacts $resolvedPath $beforeHash $conflicts
    $afterHash = Get-FileSha256 $resolvedPath
    if ($afterHash -ne $beforeHash) { throw 'Analyze changed the target file.' }
    [pscustomobject]@{ Mode='Analyze'; FilePath=$resolvedPath; FileSha256=$beforeHash; ConflictCount=$conflicts.Count; ReportPath=$artifacts.ReportPath; DecisionTemplate=$artifacts.DecisionPath }
    return
  }

  Assert-SoftwareStopped
  if ([string]::IsNullOrWhiteSpace($DecisionFile) -or -not (Test-Path -LiteralPath $DecisionFile -PathType Leaf)) { throw '-Repair requires -DecisionFile.' }
  if ($ExpectedSha256 -notmatch '^[A-Fa-f0-9]{64}$' -or $beforeHash -ne $ExpectedSha256.ToUpperInvariant()) { throw 'ExpectedSha256 does not match the locked file.' }
  $decisions = @(Import-Csv -LiteralPath $DecisionFile)
  if ($decisions.Count -ne $conflicts.Count) { throw 'Decision file must contain exactly one row for every current conflict.' }
  foreach ($conflict in $conflicts) {
    $id = Get-ConflictId $conflict.Kind $conflict.Identity
    $matches = @($decisions | Where-Object { $_.conflict_id -eq $id })
    if ($matches.Count -ne 1) { throw "Missing or duplicate decision: $id" }
    $keepLine = 0
    if (-not [int]::TryParse([string]$matches[0].keep_line, [ref]$keepLine) -or $keepLine -notin $conflict.Lines) { throw "Invalid keep_line for conflict: $id" }
    $groupLines = @($document.Lines | Where-Object { $_.LineNumber -in $conflict.Lines })
    $keep = $groupLines | Where-Object { $_.LineNumber -eq $keepLine } | Select-Object -First 1
    if ($conflict.Kind -eq 'AmbiguousBoxUnknownFields') {
      $maxFields = @{}
      foreach ($field in @('weight','accepted_count','corrected_count','rejected_count')) { $maxFields[$field] = @($groupLines | ForEach-Object { [int](Get-PropertyText $_.Row $field) } | Measure-Object -Maximum).Maximum }
      $latest = @($groupLines | ForEach-Object { Get-PropertyText $_.Row 'last_used_at' } | Sort-Object -Descending | Select-Object -First 1)[0]
      foreach ($field in $maxFields.Keys) { $keep.Row.$field = [string]$maxFields[$field] }
      $keep.Row.last_used_at = $latest
      $keep.Content = $keep.Row | ConvertTo-Json -Compress -Depth 20
    }
    foreach ($line in $groupLines) { if ($line.LineNumber -ne $keepLine) { $line.Removed = $true } }
  }

  $backup = $resolvedPath + '.pre-identity-repair-' + (Get-Date -Format 'yyyyMMdd-HHmmss-fff') + '.bak'
  [IO.File]::Copy($resolvedPath, $backup, $false)
  Write-DocumentAtomic $document $resolvedPath
  $afterDocument = Read-RawDocument $resolvedPath
  $remaining = @(Get-Conflicts $afterDocument)
  if ($remaining.Count -ne 0) { throw 'Post-repair analysis still reports identity conflicts.' }
  [pscustomobject]@{ Mode='Repair'; FilePath=$resolvedPath; BeforeSha256=$beforeHash; AfterSha256=(Get-FileSha256 $resolvedPath); BackupPath=$backup; RepairedConflictCount=$conflicts.Count }
}
finally {
  if ($acquired) { $mutex.ReleaseMutex() }
  $mutex.Dispose()
}
