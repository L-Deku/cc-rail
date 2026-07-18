# RecoLearning 公共函数:连接、SQL 执行、批量导入、签名归一化。
$ErrorActionPreference = 'Stop'

$script:RecoServer   = '192.168.2.213,1433'
$script:RecoUser     = 'reco'
$script:RecoPassword = 'Des_Reco_2006'
$script:LearningDb   = 'RecoLearning'

function Get-RecoConnectionString {
  param([string]$Database = $script:LearningDb)
  "Server=$($script:RecoServer);Database=$Database;User ID=$($script:RecoUser);Password=$($script:RecoPassword);Connect Timeout=15"
}

function Invoke-RecoNonQuery {
  param([Parameter(Mandatory)][string]$Sql, [string]$Database = $script:LearningDb, [hashtable]$Parameters = @{})
  $conn = New-Object System.Data.SqlClient.SqlConnection (Get-RecoConnectionString -Database $Database)
  try {
    $conn.Open()
    $cmd = $conn.CreateCommand(); $cmd.CommandText = $Sql; $cmd.CommandTimeout = 300
    foreach ($k in $Parameters.Keys) { [void]$cmd.Parameters.AddWithValue('@' + $k, $Parameters[$k]) }
    return $cmd.ExecuteNonQuery()
  } finally { $conn.Dispose() }
}

function Invoke-RecoScalar {
  param([Parameter(Mandatory)][string]$Sql, [string]$Database = $script:LearningDb, [hashtable]$Parameters = @{})
  $conn = New-Object System.Data.SqlClient.SqlConnection (Get-RecoConnectionString -Database $Database)
  try {
    $conn.Open()
    $cmd = $conn.CreateCommand(); $cmd.CommandText = $Sql; $cmd.CommandTimeout = 300
    foreach ($k in $Parameters.Keys) { [void]$cmd.Parameters.AddWithValue('@' + $k, $Parameters[$k]) }
    return $cmd.ExecuteScalar()
  } finally { $conn.Dispose() }
}

function Invoke-RecoQuery {
  param([Parameter(Mandatory)][string]$Sql, [string]$Database = $script:LearningDb, [hashtable]$Parameters = @{})
  $conn = New-Object System.Data.SqlClient.SqlConnection (Get-RecoConnectionString -Database $Database)
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

# 与插件 ExcelLinkFeature.NormalizeForSignature 完全同款,不得改动
function Get-NormalizedPart { param([string]$Text) ([string]$Text).Trim().ToUpperInvariant().Replace(' ', '') }

function Get-QuantitySignature {
  param([string]$Name, [string]$Unit)
  $sig = (Get-NormalizedPart $Name) + '|' + (Get-NormalizedPart $Unit)
  if ($sig.Length -gt 450) { $sig = $sig.Substring(0, 450) }
  return $sig
}

function Get-Md5Hex {
  param([string]$Text)
  $md5 = [System.Security.Cryptography.MD5]::Create()
  try { ([System.BitConverter]::ToString($md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes([string]$Text)))).Replace('-','').ToLowerInvariant() }
  finally { $md5.Dispose() }
}
