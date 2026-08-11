# RecoLearning 公共函数:连接、SQL 执行、批量导入、签名归一化。
$ErrorActionPreference = 'Stop'
$script:LearningDb   = 'RecoLearning'

$credentialStoreScript = Join-Path (Split-Path -Parent $PSScriptRoot) 'RecoCredential\RecoCredentialStore.ps1'
. $credentialStoreScript

function Get-RecoConnectionString {
  param([string]$Database = $script:LearningDb, [string]$Server = '')

  $credential = Get-RecoSqlCredential -Name Learning
  if ([string]::IsNullOrWhiteSpace($Server)) {
    $Server = $credential.Server
  }
  elseif (-not [string]::Equals($Server, $credential.Server, [StringComparison]::OrdinalIgnoreCase)) {
    $businessCredential = Get-RecoSqlCredential -Name Business
    if (-not [string]::Equals($Server, $businessCredential.Server, [StringComparison]::OrdinalIgnoreCase)) {
      throw 'No local DPAPI SQL credential entry matches the requested server.'
    }
    $credential = $businessCredential
  }

  $builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
  $builder['Data Source'] = $Server
  $builder['Initial Catalog'] = $Database
  $builder['User ID'] = $credential.User
  $builder['Password'] = $credential.Password
  $builder['Connect Timeout'] = 15
  $builder['Persist Security Info'] = $false
  return $builder.ConnectionString
}

function Invoke-RecoNonQuery {
  param([Parameter(Mandatory)][string]$Sql, [string]$Database = $script:LearningDb, [hashtable]$Parameters = @{}, [string]$Server = '')
  $conn = New-Object System.Data.SqlClient.SqlConnection (Get-RecoConnectionString -Database $Database -Server $Server)
  try {
    $conn.Open()
    $cmd = $conn.CreateCommand(); $cmd.CommandText = $Sql; $cmd.CommandTimeout = 300
    foreach ($k in $Parameters.Keys) { [void]$cmd.Parameters.AddWithValue('@' + $k, $Parameters[$k]) }
    return $cmd.ExecuteNonQuery()
  } finally { $conn.Dispose() }
}

function Invoke-RecoScalar {
  param([Parameter(Mandatory)][string]$Sql, [string]$Database = $script:LearningDb, [hashtable]$Parameters = @{}, [string]$Server = '')
  $conn = New-Object System.Data.SqlClient.SqlConnection (Get-RecoConnectionString -Database $Database -Server $Server)
  try {
    $conn.Open()
    $cmd = $conn.CreateCommand(); $cmd.CommandText = $Sql; $cmd.CommandTimeout = 300
    foreach ($k in $Parameters.Keys) { [void]$cmd.Parameters.AddWithValue('@' + $k, $Parameters[$k]) }
    return $cmd.ExecuteScalar()
  } finally { $conn.Dispose() }
}

function Invoke-RecoQuery {
  param([Parameter(Mandatory)][string]$Sql, [string]$Database = $script:LearningDb, [hashtable]$Parameters = @{}, [string]$Server = '')
  $conn = New-Object System.Data.SqlClient.SqlConnection (Get-RecoConnectionString -Database $Database -Server $Server)
  try {
    $conn.Open()
    $cmd = $conn.CreateCommand(); $cmd.CommandText = $Sql; $cmd.CommandTimeout = 300
    foreach ($k in $Parameters.Keys) { [void]$cmd.Parameters.AddWithValue('@' + $k, $Parameters[$k]) }
    $dt = New-Object System.Data.DataTable
    $da = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
    [void]$da.Fill($dt)
    return ,$dt
  } finally { $conn.Dispose() }
}

function Invoke-RecoBulkCopy {
  param([Parameter(Mandatory)][System.Data.DataTable]$Table, [Parameter(Mandatory)][string]$TargetTable)
  if ($Table.Rows.Count -eq 0) { return }
  $conn = New-Object System.Data.SqlClient.SqlConnection (Get-RecoConnectionString)
  try {
    $conn.Open()
    $bulk = New-Object System.Data.SqlClient.SqlBulkCopy($conn)
    try {
      $bulk.DestinationTableName = $TargetTable
      $bulk.BulkCopyTimeout = 600
      foreach ($col in $Table.Columns) { [void]$bulk.ColumnMappings.Add($col.ColumnName, $col.ColumnName) }
      $bulk.WriteToServer($Table)
    } finally { $bulk.Close() }
  } finally { $conn.Dispose() }
}

function Invoke-RecoQueryInTransaction {
  param([Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory)][System.Data.SqlClient.SqlTransaction]$Transaction,
        [Parameter(Mandatory)][string]$Sql)
  $cmd = $Connection.CreateCommand(); $cmd.Transaction = $Transaction; $cmd.CommandText = $Sql; $cmd.CommandTimeout = 300
  $dt = New-Object System.Data.DataTable
  $da = New-Object System.Data.SqlClient.SqlDataAdapter($cmd)
  [void]$da.Fill($dt)
  return ,$dt
}

function Invoke-RecoNonQueryInTransaction {
  param([Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory)][System.Data.SqlClient.SqlTransaction]$Transaction,
        [Parameter(Mandatory)][string]$Sql)
  $cmd = $Connection.CreateCommand(); $cmd.Transaction = $Transaction; $cmd.CommandText = $Sql; $cmd.CommandTimeout = 300
  return $cmd.ExecuteNonQuery()
}

function Invoke-RecoBulkCopyInTransaction {
  param([Parameter(Mandatory)][System.Data.SqlClient.SqlConnection]$Connection,
        [Parameter(Mandatory)][System.Data.SqlClient.SqlTransaction]$Transaction,
        [Parameter(Mandatory)][System.Data.DataTable]$Table,
        [Parameter(Mandatory)][string]$TargetTable)
  if ($Table.Rows.Count -eq 0) { return }
  $bulk = New-Object System.Data.SqlClient.SqlBulkCopy($Connection, [System.Data.SqlClient.SqlBulkCopyOptions]::Default, $Transaction)
  try {
    $bulk.DestinationTableName = $TargetTable; $bulk.BulkCopyTimeout = 600
    foreach ($col in $Table.Columns) { [void]$bulk.ColumnMappings.Add($col.ColumnName, $col.ColumnName) }
    $bulk.WriteToServer($Table)
  } finally { $bulk.Close() }
}

# 与插件 ExcelLinkFeature.NormalizeForSignature 同款的基础归一化。
function Get-NormalizedPart {
  param([string]$Text)
  $source = ([string]$Text).Normalize([System.Text.NormalizationForm]::FormKC).ToUpperInvariant()
  $source = $source.Replace([char]0x0424, [char]0x03A6).Replace([char]0x00D7, [char]0x0058)
  $normalized = New-Object System.Text.StringBuilder $source.Length
  foreach ($raw in $source.ToCharArray()) {
    if ([char]::IsWhiteSpace($raw)) { continue }
    [void]$normalized.Append($raw)
  }
  return $normalized.ToString()
}

function Get-NormalizedLearningUnit {
  param([string]$Text)
  $unit = ([string]$Text).Trim().ToLowerInvariant()
  $unit = $unit.Replace(' ', '').Replace(([string][char]0x3000), '')
  $unit = $unit.Replace([char]0xff08, '(').Replace([char]0xff09, ')')
  $unit = $unit.Replace([char]0xff0e, '.').Replace([char]0x00b7, '.').Replace([char]0xfe52, '.')
  $unit = $unit.Replace([char]0xff4d, 'm').Replace([char]0xff2d, 'm')
  for ($i = 0; $i -le 9; $i++) { $unit = $unit.Replace([char](0xff10 + $i), [char](0x30 + $i)) }
  $unit = $unit.Replace([char]0x00b2, '2').Replace([char]0x00b3, '3')
  $unit = $unit.Replace(([string][char]0x33a1), 'm2').Replace(([string][char]0x33a5), 'm3')
  $unit = $unit.Replace('m^2', 'm2').Replace('m＾2', 'm2')
  $unit = $unit.Replace('m^3', 'm3').Replace('m＾3', 'm3')
  $unit = $unit.Replace('千米', 'km').Replace('公里', 'km').Replace('百米', 'hm')
  $unit = $unit.Replace('千克', 'kg').Replace('公斤', 'kg').Replace('吨', 't')
  $unit = $unit.Replace('立方米', 'm3').Replace('平米', 'm2').Replace('平方米', 'm2').Replace('米', 'm')
  return $unit.ToUpperInvariant()
}

# 公式哈希与 C# ExcelLinkFeature.NormalizeExcelLinkUnit 保持逐字一致（小写）。
function Get-NormalizedLearningFormulaUnit {
  param([string]$Text)
  return (Get-NormalizedLearningUnit $Text).ToLowerInvariant()
}

function Get-NormalizedLearningEntryCode {
  param([string]$Text)
  $code = ([string]$Text).Trim()
  foreach ($dash in 0x2010,0x2011,0x2012,0x2013,0x2014,0x2212,0xff0d) {
    $code = $code.Replace([char]$dash, '-')
  }
  if ($code -notmatch '^[0-9]{2,}(?:-[0-9]+)*$') { return '' }
  return $code
}

function Get-NormalizedLearningMethodNo {
  param([string]$Text)
  $method = ([string]$Text).Trim()
  if ($method -eq '') { return '' }
  $compact = $method.Replace(' ', '').Replace(([string][char]0x3000), '').Replace('-', '').Replace(([string][char]0x2014), '')
  if ($compact.IndexOf('101号文', [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
      $compact.IndexOf('101estimate', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return '101号文估算' }
  if ($compact.IndexOf('TB10801', [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
      $compact.IndexOf('2024', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return 'TB 10801—2024' }
  if ($compact.IndexOf('国铁科法', [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
      $compact.IndexOf('30号文', [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -or
      $compact.IndexOf('2020', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return '30号文' }
  return ''
}

function Get-NormalizedLearningSoftwarePartition {
  param([string]$Text)
  $partition = ([string]$Text).Trim()
  if ($partition -eq '2020' -or $partition -eq '2024') { return $partition }
  return ''
}

function Get-LearningBaseTargetCode {
  param([string]$Code)
  $normalized = ([string]$Code).Trim().ToUpperInvariant()
  $asterisk = $normalized.IndexOf('*')
  $slash = $normalized.IndexOf('/')
  $suffix = -1
  if ($asterisk -ge 0) { $suffix = $asterisk }
  if ($slash -ge 0 -and ($suffix -lt 0 -or $slash -lt $suffix)) { $suffix = $slash }
  if ($suffix -ge 0) { return $normalized.Substring(0, $suffix) }
  return $normalized
}

function Test-ContextSensitiveLearningCode {
  param([string]$Code)
  return @('SF','SH','SQ','ZLF','LF','YF','TLF','GF','JF','XGT1') -contains (Get-LearningBaseTargetCode $Code)
}

function Get-LearningTargetIdentityKey {
  param([string]$Kind, [string]$Code, [string]$Name, [string]$Unit)
  $rawCode = ([string]$Code).Trim()
  $normalizedKind = ([string]$Kind).Trim().ToLowerInvariant()
  if ($normalizedKind -eq '') {
    $normalizedKind = if ($rawCode -match '^\d+$') { 'material' } else { 'quota' }
  }
  $baseKey = $normalizedKind + ':' + $rawCode.ToUpperInvariant()
  if (-not (Test-ContextSensitiveLearningCode $rawCode)) { return $baseKey }
  return $baseKey + '|' + (Get-NormalizedPart $Name) + '|' + (Get-NormalizedLearningUnit $Unit)
}

function Test-ReliableQuantityUnit {
  param([string]$Unit)
  $unitText = ([string]$Unit).Trim()
  if ($unitText -eq '' -or $unitText.Length -gt 20 -or $unitText -match '[\s\u3000]') { return $false }
  return $unitText -match '[\p{L}\u4e00-\u9fff²³㎡㎥%]'
}

function Get-ReliableQuantityUnits {
  param([System.Collections.IEnumerable]$Rows)
  $units = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
  foreach ($unit in 'm','km','hm','m2','m²','㎡','10m2','100m2','m3','m³','㎥','10m3','100m3','kg','t','10t','100t','吨','个','根','套','台','组','张','处','项','座','樘','块','孔','工日','元','系统','%') {
    [void]$units.Add($unit)
  }
  if ($null -ne $Rows) {
    foreach ($row in $Rows) {
      foreach ($propertyName in 'quantity_unit','target_unit') {
        $property = $row.PSObject.Properties[$propertyName]
        if ($null -eq $property) { continue }
        $unit = ([string]$property.Value).Trim()
        if (Test-ReliableQuantityUnit $unit) { [void]$units.Add($unit) }
      }
    }
  }
  return @($units | Sort-Object @{ Expression = { $_.Length }; Descending = $true }, @{ Expression = { $_ }; Descending = $false })
}

function Get-InferredTrailingQuantityUnit {
  param([string]$Name, [System.Collections.IEnumerable]$KnownUnits)
  $nameText = ([string]$Name).Trim()
  if ($nameText -eq '' -or $null -eq $KnownUnits) { return '' }
  foreach ($candidate in $KnownUnits) {
    $unitText = ([string]$candidate).Trim()
    if (-not (Test-ReliableQuantityUnit $unitText) -or $nameText.Length -le $unitText.Length) { continue }
    if (-not $nameText.EndsWith($unitText, [System.StringComparison]::OrdinalIgnoreCase)) { continue }
    $boundary = $nameText[$nameText.Length - $unitText.Length - 1]
    if ($boundary -eq ' ' -or $boundary -eq "`t" -or [int][char]$boundary -eq 0x3000) { return $unitText }
  }
  return ''
}

function Get-CanonicalQuantityName {
  param([string]$Name, [string]$Unit, [System.Collections.IEnumerable]$KnownUnits)
  $nameText = ([string]$Name).Trim()
  $unitText = ([string]$Unit).Trim()
  if ($unitText -eq '') { $unitText = Get-InferredTrailingQuantityUnit $nameText $KnownUnits }
  if ($nameText -eq '' -or $unitText -eq '') { return $nameText }
  $pattern = '[ \t\u3000]+' + [regex]::Escape($unitText) + '[ \t\u3000]*$'
  return ([regex]::Replace($nameText, $pattern, '', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)).Trim()
}

function Get-QuantitySignature {
  param([string]$Name, [string]$Unit, [System.Collections.IEnumerable]$KnownUnits)
  $sig = (Get-NormalizedPart (Get-CanonicalQuantityName $Name $Unit $KnownUnits)) + '|'
  if ($sig.Length -gt 450) { $sig = $sig.Substring(0, 450) }
  return $sig
}

function Get-LearningMethodPartition {
  param([string]$Method)
  $methodText = ([string]$Method).Trim()
  if ($methodText -eq '') { return '' }
  if ($methodText.IndexOf('2024', [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return '2024' }
  # 2020 软件内预算和 101 号文估算可共存，不再拆成第三个学习分区。
  return '2020'
}

function Get-MethodScopedMapKey {
  param([string]$Method, [string]$Signature, [string]$BoxId)
  return (Get-LearningMethodPartition $Method) + "`n" + ([string]$Signature) + "`n" + ([string]$BoxId)
}

function Get-Md5Hex {
  param([string]$Text)
  $md5 = [System.Security.Cryptography.MD5]::Create()
  try { ([System.BitConverter]::ToString($md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes([string]$Text)))).Replace('-','').ToLowerInvariant() }
  finally { $md5.Dispose() }
}
