# RecoLearning 公共函数:连接、SQL 执行、批量导入、签名归一化。
$ErrorActionPreference = 'Stop'

$script:RecoServer   = '192.168.2.213,1433'
$script:RecoUser     = 'reco'
$script:RecoPassword = 'Des_Reco_2006'
$script:LearningDb   = 'RecoLearning'

function Get-RecoConnectionString {
  param([string]$Database = $script:LearningDb, [string]$Server = $script:RecoServer)
  "Server=$Server;Database=$Database;User ID=$($script:RecoUser);Password=$($script:RecoPassword);Connect Timeout=15"
}

function Invoke-RecoNonQuery {
  param([Parameter(Mandatory)][string]$Sql, [string]$Database = $script:LearningDb, [hashtable]$Parameters = @{}, [string]$Server = $script:RecoServer)
  $conn = New-Object System.Data.SqlClient.SqlConnection (Get-RecoConnectionString -Database $Database -Server $Server)
  try {
    $conn.Open()
    $cmd = $conn.CreateCommand(); $cmd.CommandText = $Sql; $cmd.CommandTimeout = 300
    foreach ($k in $Parameters.Keys) { [void]$cmd.Parameters.AddWithValue('@' + $k, $Parameters[$k]) }
    return $cmd.ExecuteNonQuery()
  } finally { $conn.Dispose() }
}

function Invoke-RecoScalar {
  param([Parameter(Mandatory)][string]$Sql, [string]$Database = $script:LearningDb, [hashtable]$Parameters = @{}, [string]$Server = $script:RecoServer)
  $conn = New-Object System.Data.SqlClient.SqlConnection (Get-RecoConnectionString -Database $Database -Server $Server)
  try {
    $conn.Open()
    $cmd = $conn.CreateCommand(); $cmd.CommandText = $Sql; $cmd.CommandTimeout = 300
    foreach ($k in $Parameters.Keys) { [void]$cmd.Parameters.AddWithValue('@' + $k, $Parameters[$k]) }
    return $cmd.ExecuteScalar()
  } finally { $conn.Dispose() }
}

function Invoke-RecoQuery {
  param([Parameter(Mandatory)][string]$Sql, [string]$Database = $script:LearningDb, [hashtable]$Parameters = @{}, [string]$Server = $script:RecoServer)
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
function Get-NormalizedPart { param([string]$Text) ([string]$Text).Trim().ToUpperInvariant().Replace(' ', '') }

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
