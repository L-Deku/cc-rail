$ErrorActionPreference = 'Stop'

function Read-RequiredText {
  param([Parameter(Mandatory)][string]$Prompt)
  $value = (Read-Host $Prompt).Trim()
  if ([string]::IsNullOrWhiteSpace($value)) {
    throw ($Prompt + '不能为空。')
  }
  return $value
}

function Read-RequiredSecret {
  param([Parameter(Mandatory)][string]$Prompt)
  $value = Read-Host $Prompt -AsSecureString
  if ($null -eq $value -or $value.Length -eq 0) {
    throw ($Prompt + '不能为空。')
  }
  return $value
}

function ConvertTo-PlainText {
  param([Parameter(Mandatory)][Security.SecureString]$SecureValue)
  $pointer = [IntPtr]::Zero
  try {
    $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureValue)
    return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)
  }
  finally {
    if ($pointer -ne [IntPtr]::Zero) {
      [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)
    }
  }
}

function Test-RecoLearningCredential {
  param(
    [Parameter(Mandatory)][string]$Server,
    [Parameter(Mandatory)][string]$User,
    [Parameter(Mandatory)][string]$Password
  )

  $dataSource = $Server.Trim()
  if ($dataSource.IndexOf(',') -lt 0) {
    $dataSource += ',1433'
  }
  $builder = New-Object System.Data.SqlClient.SqlConnectionStringBuilder
  $builder['Data Source'] = $dataSource
  $builder['Initial Catalog'] = 'RecoLearning'
  $builder['User ID'] = $User
  $builder['Password'] = $Password
  $builder['Connect Timeout'] = 5
  $builder['Encrypt'] = $false
  $builder['TrustServerCertificate'] = $true
  $builder['Persist Security Info'] = $false

  $connection = New-Object System.Data.SqlClient.SqlConnection $builder.ConnectionString
  try {
    $connection.Open()
    $command = $connection.CreateCommand()
    try {
      $command.CommandText = "SELECT CASE WHEN DB_NAME()=N'RecoLearning' THEN 1 ELSE 0 END"
      $command.CommandTimeout = 5
      if ([int]$command.ExecuteScalar() -ne 1) {
        throw '学习库只读验证连接到了错误的数据库。'
      }
    }
    finally {
      $command.Dispose()
    }
  }
  catch [System.Data.SqlClient.SqlException] {
    [string[]]$numbers = @($_.Exception.Errors | ForEach-Object { $_.Number.ToString() } | Select-Object -Unique)
    throw ('学习库只读连接失败，SQL错误号：' + ($numbers -join ','))
  }
  finally {
    $connection.Dispose()
  }
}

$modulePath = Join-Path $PSScriptRoot 'RecoCredentialStore.ps1'
if (-not (Test-Path -LiteralPath $modulePath)) {
  throw "缺少凭据模块：$modulePath"
}
. $modulePath

foreach ($processName in @('ReJJGSNet2024', 'RejjNet2020')) {
  if (Get-Process -Name $processName -ErrorAction SilentlyContinue) {
    throw "请先关闭 $processName，再初始化SQL凭据。"
  }
}

$storePath = Get-RecoSqlCredentialStorePath
Write-Host '此操作必须由实际运行计价软件的Windows用户执行。'
Write-Host ('凭据保存位置：' + $storePath)
Write-Host '密码仅用于本次内存验证和CurrentUser DPAPI加密，不会显示或写入脚本。'

if (Test-Path -LiteralPath $storePath) {
  $existing = Get-RecoSqlCredential -Name Learning -Path $storePath
  try {
    Test-RecoLearningCredential -Server $existing.Server -User $existing.User -Password $existing.Password
    Write-Host 'PASS：现有SQL凭据可由当前用户解密，RecoLearning只读连接正常。'
    exit 0
  }
  finally {
    $existing = $null
  }
}

$learningPassword = $null
$businessPassword = $null
$learningSecurePassword = $null
$businessSecurePassword = $null
$createdStore = $false
try {
  $learningServer = Read-RequiredText '学习库服务器'
  $learningUser = Read-RequiredText '学习库用户名'
  $learningSecurePassword = Read-RequiredSecret '学习库密码'
  $businessServer = Read-RequiredText '业务库服务器'
  $businessUser = Read-RequiredText '业务库用户名'
  $businessSecurePassword = Read-RequiredSecret '业务库密码'

  $learningPassword = ConvertTo-PlainText $learningSecurePassword
  $businessPassword = ConvertTo-PlainText $businessSecurePassword

  Write-Host '正在执行RecoLearning只读连接验证……'
  Test-RecoLearningCredential -Server $learningServer -User $learningUser -Password $learningPassword

  $writtenPath = Write-RecoSqlCredentialStore `
    -LearningServer $learningServer `
    -LearningUser $learningUser `
    -LearningPassword $learningPassword `
    -BusinessServer $businessServer `
    -BusinessUser $businessUser `
    -BusinessPassword $businessPassword `
    -Path $storePath
  $createdStore = $true

  $roundTrip = Get-RecoSqlCredential -Name Learning -Path $writtenPath
  try {
    if ($roundTrip.Server -ne $learningServer -or $roundTrip.User -ne $learningUser -or
        $roundTrip.Password -ne $learningPassword) {
      throw 'DPAPI凭据写入后的当前用户回读校验失败。'
    }
  }
  finally {
    $roundTrip = $null
  }

  Write-Host 'PASS：同事账户SQL凭据初始化完成，RecoLearning只读连接正常。'
  Write-Host '请重新打开计价软件，再绑定一条工程量进行SQL学习验证。'
}
catch {
  if ($createdStore -and (Test-Path -LiteralPath $storePath)) {
    Remove-Item -LiteralPath $storePath -Force
  }
  throw
}
finally {
  $learningPassword = $null
  $businessPassword = $null
  if ($learningSecurePassword) { $learningSecurePassword.Dispose() }
  if ($businessSecurePassword) { $businessSecurePassword.Dispose() }
}
