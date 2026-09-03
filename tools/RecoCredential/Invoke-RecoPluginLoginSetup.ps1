param(
  [Security.SecureString]$Password,
  [switch]$SkipConfig,
  [switch]$SkipProjectDatabases
)

# 一键建号：用本机 DPAPI 凭据库里的管理员账号（reco，两台服务器上均为 sysadmin）连 .213 与 .13，
# 创建最小权限登录名 reco_plugin：
#   .213：RecoLearning 读写，RecoData2024 只读，model 只读，其余在线用户库（项目库）逐个只读；
#   .13 ：RecoData2020 只读，model 只读，其余在线用户库逐个只读。
# 项目库有几百到几千个且多为 AUTO_CLOSE，逐库授权每库约 0.3~1 秒，整体可能要几十分钟；脚本逐库执行、
# 打印进度、单库失败不中断，可反复运行（已建的会跳过）。最后直接生成 RecoPluginSql.json。全程不显示任何密码。

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'RecoCredentialStore.ps1')

function ConvertFrom-SecureStringPlain {
  param([Security.SecureString]$Secure)
  $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secure)
  try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
  finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
}

function Read-PluginPassword {
  while ($true) {
    $first = Read-Host -Prompt '请为新账号 reco_plugin 设置密码（只用英文字母和数字，至少 12 位）' -AsSecureString
    $second = Read-Host -Prompt '请再输入一次确认' -AsSecureString
    $p1 = ConvertFrom-SecureStringPlain $first
    $p2 = ConvertFrom-SecureStringPlain $second
    if ($p1 -ne $p2) { Write-Host '两次输入不一致，请重来。'; continue }
    if ($p1 -notmatch '^[A-Za-z0-9]{12,}$') { Write-Host '密码只能用英文字母和数字，且至少 12 位，请重来。'; continue }
    return $first
  }
}

function New-AdminConnection {
  param([pscustomobject]$Credential)
  $builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
  $builder['Data Source'] = $Credential.Server
  $builder['Initial Catalog'] = 'master'
  $builder['User ID'] = $Credential.User
  $builder['Password'] = $Credential.Password
  $builder['Connect Timeout'] = 15
  $builder['Encrypt'] = $false
  $builder['TrustServerCertificate'] = $true
  $builder['Persist Security Info'] = $false
  return (New-Object System.Data.SqlClient.SqlConnection $builder.ConnectionString)
}

function Invoke-NonQuery {
  param([System.Data.SqlClient.SqlConnection]$Connection, [string]$Sql, [int]$TimeoutSeconds = 120)
  $cmd = $Connection.CreateCommand()
  $cmd.CommandText = $Sql
  $cmd.CommandTimeout = $TimeoutSeconds
  [void]$cmd.ExecuteNonQuery()
}

function Invoke-Scalar {
  param([System.Data.SqlClient.SqlConnection]$Connection, [string]$Sql, [int]$TimeoutSeconds = 120)
  $cmd = $Connection.CreateCommand()
  $cmd.CommandText = $Sql
  $cmd.CommandTimeout = $TimeoutSeconds
  return $cmd.ExecuteScalar()
}

function Grant-DatabaseRoles {
  param([System.Data.SqlClient.SqlConnection]$Connection, [string]$Database, [string[]]$Roles, [int]$TimeoutSeconds = 120)
  $sql = 'USE ' + ('[' + $Database.Replace(']', ']]') + ']') + ";`n" +
    "IF NOT EXISTS (SELECT 1 FROM sys.database_principals WHERE name = N'reco_plugin') CREATE USER [reco_plugin] FOR LOGIN [reco_plugin];`n"
  foreach ($role in $Roles) {
    # SQL Server 2008 R2：没有 IS_ROLEMEMBER 和 ALTER ROLE … ADD MEMBER，用目录视图判断后 sp_addrolemember。
    $sql += "IF NOT EXISTS (SELECT 1 FROM sys.database_role_members rm JOIN sys.database_principals r ON r.principal_id = rm.role_principal_id JOIN sys.database_principals u ON u.principal_id = rm.member_principal_id WHERE r.name = N'$role' AND u.name = N'reco_plugin') EXEC sp_addrolemember N'$role', N'reco_plugin';`n"
  }
  $sql += 'USE [master];'
  Invoke-NonQuery -Connection $Connection -Sql $sql -TimeoutSeconds $TimeoutSeconds
}

function Invoke-SetupOnServer {
  param(
    [string]$Label,
    [pscustomobject]$Credential,
    [System.Collections.IDictionary]$FixedDatabases,
    [string]$DefaultDatabase,
    [string]$PlainPassword,
    [switch]$SkipProjects
  )
  Write-Host ''
  Write-Host ('==== ' + $Label + '：' + $Credential.Server + '（以 ' + $Credential.User + ' 身份执行）')
  $conn = New-AdminConnection -Credential $Credential
  try {
    $conn.Open()
    if ([int](Invoke-Scalar -Connection $conn -Sql "SELECT IS_SRVROLEMEMBER('sysadmin')") -ne 1) {
      throw ('账号 ' + $Credential.User + ' 在 ' + $Credential.Server + ' 上不是 sysadmin，无法建号。请联系服务器管理员。')
    }

    # 1. 登录名：不存在就建，存在就同步密码；确保启用；不属于任何服务器角色。
    $loginSql = @"
IF NOT EXISTS (SELECT 1 FROM sys.server_principals WHERE name = N'reco_plugin')
    CREATE LOGIN [reco_plugin] WITH PASSWORD = N'$PlainPassword', CHECK_POLICY = OFF, CHECK_EXPIRATION = OFF, DEFAULT_DATABASE = [$DefaultDatabase];
ELSE
    ALTER LOGIN [reco_plugin] WITH PASSWORD = N'$PlainPassword';
ALTER LOGIN [reco_plugin] ENABLE;
"@
    Invoke-NonQuery -Connection $conn -Sql $loginSql
    $roleCount = [int](Invoke-Scalar -Connection $conn -Sql @"
SELECT COUNT(*) FROM sys.server_role_members m
JOIN sys.server_principals r ON r.principal_id = m.role_principal_id
JOIN sys.server_principals p ON p.principal_id = m.member_principal_id
WHERE p.name = N'reco_plugin' AND r.name <> N'public'
"@)
    if ($roleCount -ne 0) { throw ($Label + '：reco_plugin 不应属于任何服务器角色，请人工核对后移除。') }
    Write-Host '  登录名 reco_plugin 已就绪（密码已设置，无服务器角色）。'

    # 2. 固定库。
    foreach ($db in @($FixedDatabases.Keys)) {
      $dbId = Invoke-Scalar -Connection $conn -Sql "SELECT DB_ID(N'$db')"
      if ($null -eq $dbId -or $dbId -is [DBNull]) {
        Write-Host ('  跳过：库不存在 ' + $db)
        continue
      }
      Grant-DatabaseRoles -Connection $conn -Database $db -Roles $FixedDatabases[$db]
      Write-Host ('  已授权 ' + $db + '：' + ($FixedDatabases[$db] -join ','))
    }

    if ($SkipProjects) {
      Write-Host '  跳过项目库逐库授权（-SkipProjectDatabases）。'
      return
    }

    # 3. 其余在线用户库（项目库）：只读，逐库执行、打印进度、失败继续。
    $excluded = @('RecoLearning', 'RecoData2024', 'RecoData2020') + @($FixedDatabases.Keys)
    $listCmd = $conn.CreateCommand()
    $listCmd.CommandText = "SELECT name FROM sys.databases WHERE database_id > 4 AND state_desc = N'ONLINE' AND is_read_only = 0 ORDER BY name"
    $databases = New-Object System.Collections.Generic.List[string]
    $reader = $listCmd.ExecuteReader()
    try { while ($reader.Read()) { $databases.Add([string]$reader['name']) } }
    finally { $reader.Dispose() }
    $targets = @($databases | Where-Object { $excluded -notcontains $_ })
    Write-Host ('  项目库共 ' + $targets.Count + ' 个，开始逐库授权只读（每库约 0.3~1 秒，请耐心等待，不要关闭窗口）…')
    $done = 0; $failed = New-Object System.Collections.Generic.List[string]
    $sw = [Diagnostics.Stopwatch]::StartNew()
    foreach ($db in $targets) {
      try { Grant-DatabaseRoles -Connection $conn -Database $db -Roles @('db_datareader') -TimeoutSeconds 60 }
      catch { $failed.Add($db + '（' + $_.Exception.Message.Split("`n")[0] + '）') }
      $done++
      if ($done % 50 -eq 0 -or $done -eq $targets.Count) {
        $rate = $sw.Elapsed.TotalSeconds / [Math]::Max($done, 1)
        $remain = [int]([Math]::Round($rate * ($targets.Count - $done)))
        Write-Host ('  进度 ' + $done + '/' + $targets.Count + '，失败 ' + $failed.Count + '，预计还需约 ' + [int][Math]::Ceiling($remain / 60) + ' 分钟')
      }
    }
    if ($failed.Count -gt 0) {
      Write-Host ('  以下 ' + $failed.Count + ' 个库授权失败（可稍后重跑本脚本重试）：')
      $failed | ForEach-Object { Write-Host ('    ' + $_) }
      throw ($Label + '：有 ' + $failed.Count + ' 个项目库授权失败。')
    }
    Write-Host ('PASS：' + $Label + ' 全部 ' + $targets.Count + ' 个项目库已授权只读，用时 ' + [int]$sw.Elapsed.TotalMinutes + ' 分钟。')
  }
  finally { $conn.Dispose() }
}

if ($null -eq $Password) { $Password = Read-PluginPassword }
$plainPassword = ConvertFrom-SecureStringPlain $Password
if ($plainPassword -notmatch '^[A-Za-z0-9]{12,}$') { throw '密码只能用英文字母和数字，且至少 12 位。' }

try {
  $learning = Get-RecoSqlCredential -Name Learning
  $business = Get-RecoSqlCredential -Name Business
  Invoke-SetupOnServer -Label '学习库服务器(.213)' -Credential $learning -DefaultDatabase 'RecoLearning' -PlainPassword $plainPassword -SkipProjects:$SkipProjectDatabases `
    -FixedDatabases ([ordered]@{ 'RecoLearning' = @('db_datareader', 'db_datawriter'); 'RecoData2024' = @('db_datareader'); 'RecoData2020' = @('db_datareader'); 'model' = @('db_datareader') })
  # 两台服务器上都可能同时有 RecoData2020 与 RecoData2024（正式目录的 2020 宿主实际连 .213），定额库两边都授只读；不存在的会自动跳过。
  Invoke-SetupOnServer -Label '业务库服务器(.13)' -Credential $business -DefaultDatabase 'RecoData2020' -PlainPassword $plainPassword -SkipProjects:$SkipProjectDatabases `
    -FixedDatabases ([ordered]@{ 'RecoData2020' = @('db_datareader'); 'RecoData2024' = @('db_datareader'); 'model' = @('db_datareader') })

  if (-not $SkipConfig) {
    Write-Host ''
    Write-Host '==== 生成插件 SQL 配置文件 RecoPluginSql.json'
    & (Join-Path $PSScriptRoot 'New-RecoPluginSqlConfig.ps1') -LearningServer $learning.Server -BusinessServer $business.Server -Password $Password -Force
  }
  Write-Host ''
  Write-Host '全部完成。'
}
finally {
  $plainPassword = $null
  $learning = $null
  $business = $null
}
