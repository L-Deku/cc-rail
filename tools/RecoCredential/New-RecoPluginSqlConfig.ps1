param(
  [string]$LearningServer = '192.168.2.213,1433',
  [string]$BusinessServer = '192.168.2.13,1433',
  [string]$User = 'reco_plugin',
  [Security.SecureString]$Password,
  [string]$OutputPath,
  [ValidateRange(10000, 1000000)][int]$Iterations = 20000,
  [switch]$SkipConnectionTest,
  [switch]$Force
)

# 生成随同事发布包分发的插件 SQL 配置文件 RecoPluginSql.json。
# 内容是最小权限登录名（默认 reco_plugin）的两条端点记录，用固定应用口令封装，仅防目视。
# 生成前会连两台服务器核对：RecoLearning 可读可写、RecoData2024/RecoData2020 可读；任一失败不落盘。

$ErrorActionPreference = 'Stop'

. (Join-Path $PSScriptRoot 'RecoCredentialStore.ps1')
. (Join-Path $PSScriptRoot 'RecoCredentialTransfer.ps1')

if ([string]::Equals($User.Trim(), 'reco', [StringComparison]::OrdinalIgnoreCase)) {
  throw '插件配置文件不得使用管理员 SQL 登录名 reco。'
}
if ([string]::IsNullOrWhiteSpace($OutputPath)) {
  $repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
  $OutputPath = Join-Path $repoRoot 'RecoQuotaRecommend\bin\RecoPluginSql.json'
}
$OutputPath = [IO.Path]::GetFullPath($OutputPath)
if ((Test-Path -LiteralPath $OutputPath) -and -not $Force) {
  throw "已存在插件 SQL 配置文件，重新生成请加 -Force：$OutputPath"
}

if ($null -eq $Password) {
  $Password = Read-Host -Prompt ('请输入 SQL 登录名 ' + $User + ' 的密码') -AsSecureString
}
$bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
try { $plainPassword = [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr) }
finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
if ([string]::IsNullOrWhiteSpace($plainPassword)) { throw '密码不能为空。' }

function Test-RecoPluginEndpoint {
  param(
    [Parameter(Mandatory)][string]$Server,
    [Parameter(Mandatory)][string]$Database,
    [Parameter(Mandatory)][string]$User,
    [Parameter(Mandatory)][string]$PlainPassword,
    [switch]$RequireWrite
  )
  $builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
  $builder['Data Source'] = $Server
  $builder['Initial Catalog'] = $Database
  $builder['User ID'] = $User
  $builder['Password'] = $PlainPassword
  $builder['Connect Timeout'] = 15
  $builder['Encrypt'] = $false
  $builder['TrustServerCertificate'] = $true
  $builder['Persist Security Info'] = $false
  $conn = New-Object System.Data.SqlClient.SqlConnection $builder.ConnectionString
  try {
    $conn.Open()
    $cmd = $conn.CreateCommand()
    if ($RequireWrite) {
      $cmd.CommandText = @"
SELECT DB_NAME() AS db, SUSER_SNAME() AS login,
  HAS_PERMS_BY_NAME(N'dbo.BindingLog', N'OBJECT', N'SELECT') AS can_select,
  HAS_PERMS_BY_NAME(N'dbo.BindingLog', N'OBJECT', N'INSERT') AS can_insert,
  HAS_PERMS_BY_NAME(N'dbo.BindingLog', N'OBJECT', N'UPDATE') AS can_update
"@
    }
    else {
      $cmd.CommandText = @"
SELECT DB_NAME() AS db, SUSER_SNAME() AS login,
  (SELECT TOP 1 HAS_PERMS_BY_NAME(QUOTENAME(SCHEMA_NAME(schema_id)) + N'.' + QUOTENAME(name), N'OBJECT', N'SELECT')
   FROM sys.tables ORDER BY name) AS can_select,
  1 AS can_insert, 1 AS can_update
"@
    }
    $reader = $cmd.ExecuteReader()
    try {
      if (-not $reader.Read()) { throw "端点 $Server/$Database 未返回权限信息。" }
      $db = [string]$reader['db']
      $login = [string]$reader['login']
      $canSelect = [int]$reader['can_select']
      $canInsert = [int]$reader['can_insert']
      $canUpdate = [int]$reader['can_update']
    }
    finally { $reader.Dispose() }
    if (-not [string]::Equals($db, $Database, [StringComparison]::OrdinalIgnoreCase)) {
      throw "端点 $Server 实际打开的库是 $db，不是 $Database。"
    }
    if (-not [string]::Equals($login, $User, [StringComparison]::OrdinalIgnoreCase)) {
      throw "端点 $Server 实际登录名是 $login，不是 $User。"
    }
    if ($canSelect -ne 1) { throw "登录名 $User 对 $Server/$Database 没有读权限。" }
    if ($RequireWrite -and ($canInsert -ne 1 -or $canUpdate -ne 1)) {
      throw "登录名 $User 对 $Server/$Database 没有写权限（需要 db_datawriter）。"
    }
    Write-Host ("PASS：{0}/{1} 登录名 {2} 权限核对通过" -f $Server, $Database, $login)
  }
  finally { $conn.Dispose() }
}

try {
  if (-not $SkipConnectionTest) {
    Test-RecoPluginEndpoint -Server $LearningServer -Database 'RecoLearning' -User $User -PlainPassword $plainPassword -RequireWrite
    Test-RecoPluginEndpoint -Server $LearningServer -Database 'RecoData2024' -User $User -PlainPassword $plainPassword
    Test-RecoPluginEndpoint -Server $BusinessServer -Database 'RecoData2020' -User $User -PlainPassword $plainPassword
  }
  else {
    Write-Host '跳过连接核对（-SkipConnectionTest）。'
  }

  $json = New-RecoPluginSqlConfigJson `
    -LearningServer $LearningServer -LearningUser $User -LearningPassword $plainPassword `
    -BusinessServer $BusinessServer -BusinessUser $User -BusinessPassword $plainPassword `
    -Iterations $Iterations
  $parent = Split-Path -Parent $OutputPath
  [void][IO.Directory]::CreateDirectory($parent)
  $temp = Join-Path $parent ([IO.Path]::GetFileName($OutputPath) + '.' + [Guid]::NewGuid().ToString('N') + '.tmp')
  try {
    [IO.File]::WriteAllText($temp, $json, (New-Object Text.UTF8Encoding($false)))
    $readBack = Read-RecoPluginSqlConfig -Path $temp
    if ($readBack.Learning.Server -ne $LearningServer -or $readBack.Learning.User -ne $User -or $readBack.Learning.Password -ne $plainPassword -or
        $readBack.Business.Server -ne $BusinessServer -or $readBack.Business.User -ne $User -or $readBack.Business.Password -ne $plainPassword) {
      throw '插件 SQL 配置文件回读校验失败。'
    }
    if (Test-Path -LiteralPath $OutputPath) { Remove-Item -LiteralPath $OutputPath -Force }
    [IO.File]::Move($temp, $OutputPath)
    $temp = $null
  }
  finally {
    if ($temp -and (Test-Path -LiteralPath $temp)) { Remove-Item -LiteralPath $temp -Force }
  }
  Write-Host ('PASS：插件 SQL 配置文件已生成：' + $OutputPath)
  Write-Host ('包 ID：' + $readBack.PackageId)
  Write-Host ('学习端点：' + $LearningServer + '  业务端点：' + $BusinessServer + '  登录名：' + $User)
  Write-Host '把该文件随发布包复制到软件根目录即可；换密码时重新生成并只覆盖此文件。'
}
finally {
  $plainPassword = $null
}
