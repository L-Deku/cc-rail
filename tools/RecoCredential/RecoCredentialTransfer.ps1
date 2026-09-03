# 插件 SQL 配置文件 RecoPluginSql.json 的加密/解密核心。
# 加密核心：64 字节密钥束 = 前 32 字节 AES-256-CBC 密钥 + 后 32 字节 HMAC-SHA256 密钥，先加密后 MAC。
# 密钥由固定应用口令 + 文件内随机盐经 PBKDF2 派生；口令是常量，仅防目视，不是保密手段。
# 与 RecoShared\RecoSqlCredentialStore.cs 的 ReadFromPluginConfig 逐字节对应，改任何一侧都要同步另一侧并跑 tests\Test-RecoPluginSqlConfig.ps1。

function New-RecoRandomBytes {
  param([Parameter(Mandatory)][int]$Length)
  $bytes = New-Object byte[] $Length
  $rng = [Security.Cryptography.RandomNumberGenerator]::Create()
  try { $rng.GetBytes($bytes) }
  finally { $rng.Dispose() }
  return ,$bytes
}

function Get-RecoCredentialMacInput {
  param(
    [Parameter(Mandatory)][string]$Context,
    [Parameter(Mandatory)][byte[]]$InitializationVector,
    [Parameter(Mandatory)][byte[]]$Ciphertext
  )
  $prefix = [Text.Encoding]::UTF8.GetBytes("RecoBudget.CredentialTransfer.v1`n" + $Context + "`n")
  $result = New-Object byte[] ($prefix.Length + $InitializationVector.Length + $Ciphertext.Length)
  [Array]::Copy($prefix, 0, $result, 0, $prefix.Length)
  [Array]::Copy($InitializationVector, 0, $result, $prefix.Length, $InitializationVector.Length)
  [Array]::Copy($Ciphertext, 0, $result, $prefix.Length + $InitializationVector.Length, $Ciphertext.Length)
  return ,$result
}

function Test-RecoFixedTimeEquals {
  param(
    [Parameter(Mandatory)][byte[]]$Left,
    [Parameter(Mandatory)][byte[]]$Right
  )
  if ($Left.Length -ne $Right.Length) { return $false }
  $difference = 0
  for ($index = 0; $index -lt $Left.Length; $index++) {
    $difference = $difference -bor ($Left[$index] -bxor $Right[$index])
  }
  return $difference -eq 0
}

function Protect-RecoCredentialPayloadWithKeyBundle {
  param(
    [Parameter(Mandatory)][byte[]]$Plaintext,
    [Parameter(Mandatory)][byte[]]$KeyBundle,
    [Parameter(Mandatory)][string]$Context
  )
  if ($KeyBundle.Length -ne 64) { throw '密钥束长度无效。' }
  $encryptionKey = New-Object byte[] 32
  $macKey = New-Object byte[] 32
  [Array]::Copy($KeyBundle, 0, $encryptionKey, 0, 32)
  [Array]::Copy($KeyBundle, 32, $macKey, 0, 32)
  $initializationVector = New-RecoRandomBytes 16
  $ciphertext = $null
  $macInput = $null
  $mac = $null
  try {
    $aes = [Security.Cryptography.Aes]::Create()
    try {
      $aes.KeySize = 256
      $aes.Mode = [Security.Cryptography.CipherMode]::CBC
      $aes.Padding = [Security.Cryptography.PaddingMode]::PKCS7
      $aes.Key = $encryptionKey
      $aes.IV = $initializationVector
      $encryptor = $aes.CreateEncryptor()
      try { $ciphertext = $encryptor.TransformFinalBlock($Plaintext, 0, $Plaintext.Length) }
      finally { $encryptor.Dispose() }
    }
    finally { $aes.Dispose() }

    $macInput = Get-RecoCredentialMacInput -Context $Context `
      -InitializationVector $initializationVector -Ciphertext $ciphertext
    $hmac = New-Object Security.Cryptography.HMACSHA256
    try {
      $hmac.Key = $macKey
      $mac = $hmac.ComputeHash($macInput)
    }
    finally { $hmac.Dispose() }
    return [pscustomobject]@{
      Iv = $initializationVector
      Ciphertext = $ciphertext
      Mac = $mac
    }
  }
  finally {
    [Array]::Clear($encryptionKey, 0, $encryptionKey.Length)
    [Array]::Clear($macKey, 0, $macKey.Length)
    if ($macInput) { [Array]::Clear($macInput, 0, $macInput.Length) }
  }
}

function Unprotect-RecoCredentialPayloadWithKeyBundle {
  param(
    [Parameter(Mandatory)][byte[]]$InitializationVector,
    [Parameter(Mandatory)][byte[]]$Ciphertext,
    [Parameter(Mandatory)][byte[]]$Mac,
    [Parameter(Mandatory)][byte[]]$KeyBundle,
    [Parameter(Mandatory)][string]$Context,
    [string]$IntegrityFailureMessage = '加密载荷完整性校验失败，文件可能已损坏或被修改。'
  )
  if ($KeyBundle.Length -ne 64) { throw '密钥束长度无效。' }
  $encryptionKey = New-Object byte[] 32
  $macKey = New-Object byte[] 32
  [Array]::Copy($KeyBundle, 0, $encryptionKey, 0, 32)
  [Array]::Copy($KeyBundle, 32, $macKey, 0, 32)
  $macInput = $null
  $actualMac = $null
  try {
    $macInput = Get-RecoCredentialMacInput -Context $Context `
      -InitializationVector $InitializationVector -Ciphertext $Ciphertext
    $hmac = New-Object Security.Cryptography.HMACSHA256
    try {
      $hmac.Key = $macKey
      $actualMac = $hmac.ComputeHash($macInput)
    }
    finally { $hmac.Dispose() }
    if (-not (Test-RecoFixedTimeEquals -Left $Mac -Right $actualMac)) {
      throw $IntegrityFailureMessage
    }

    $aes = [Security.Cryptography.Aes]::Create()
    try {
      $aes.KeySize = 256
      $aes.Mode = [Security.Cryptography.CipherMode]::CBC
      $aes.Padding = [Security.Cryptography.PaddingMode]::PKCS7
      $aes.Key = $encryptionKey
      $aes.IV = $InitializationVector
      $decryptor = $aes.CreateDecryptor()
      try { return ,$decryptor.TransformFinalBlock($Ciphertext, 0, $Ciphertext.Length) }
      finally { $decryptor.Dispose() }
    }
    finally { $aes.Dispose() }
  }
  finally {
    [Array]::Clear($encryptionKey, 0, $encryptionKey.Length)
    [Array]::Clear($macKey, 0, $macKey.Length)
    if ($macInput) { [Array]::Clear($macInput, 0, $macInput.Length) }
    if ($actualMac) { [Array]::Clear($actualMac, 0, $actualMac.Length) }
  }
}

# 必须与 RecoShared\RecoSqlCredentialStore.cs 的 PluginConfigPassphrase 逐字一致。
$script:RecoPluginSqlConfigPassphrase = 'RecoBudget.PluginSqlConfig.v1:9f3a6c1d-5b2e-4e8a-a7c4-2d1f0b8e6c53'
$script:RecoPluginSqlConfigPackageType = 'plugin_sql_config'
$script:RecoPluginSqlConfigDefaultIterations = 20000

function Get-RecoPluginSqlConfigKeyBundle {
  param(
    [Parameter(Mandatory)][byte[]]$Salt,
    [Parameter(Mandatory)][int]$Iterations
  )
  if ($Salt.Length -ne 16) { throw '插件配置加密盐长度无效。' }
  if ($Iterations -lt 10000 -or $Iterations -gt 1000000) { throw '插件配置密钥迭代次数无效。' }
  $derive = New-Object Security.Cryptography.Rfc2898DeriveBytes -ArgumentList $script:RecoPluginSqlConfigPassphrase, $Salt, $Iterations
  try { return ,$derive.GetBytes(64) }
  finally { $derive.Dispose() }
}

function Get-RecoPluginSqlConfigContext {
  param(
    [Parameter(Mandatory)][string]$PackageId,
    [Parameter(Mandatory)][int]$Iterations,
    [Parameter(Mandatory)][string]$SaltBase64
  )
  return 'plugin-sql-config|' + $PackageId + '|' + $Iterations.ToString() + '|' + $SaltBase64
}

# 生成 RecoPluginSql.json 的完整 JSON 文本。明文格式与 DPAPI 凭据库相同（version=1 + 六个 Base64 字段）。
function New-RecoPluginSqlConfigJson {
  param(
    [Parameter(Mandatory)][string]$LearningServer,
    [Parameter(Mandatory)][string]$LearningUser,
    [Parameter(Mandatory)][string]$LearningPassword,
    [Parameter(Mandatory)][string]$BusinessServer,
    [Parameter(Mandatory)][string]$BusinessUser,
    [Parameter(Mandatory)][string]$BusinessPassword,
    [int]$Iterations = $script:RecoPluginSqlConfigDefaultIterations,
    [switch]$AllowAdministrativeLogin
  )
  foreach ($user in @($LearningUser, $BusinessUser)) {
    if (-not $AllowAdministrativeLogin -and [string]::Equals($user.Trim(), 'reco', [StringComparison]::OrdinalIgnoreCase)) {
      throw '插件配置文件不得使用管理员 SQL 登录名 reco，请使用最小权限的插件专用登录名。'
    }
  }
  $lines = @(
    'version=1'
    'learning.server=' + (ConvertTo-RecoCredentialField $LearningServer)
    'learning.user=' + (ConvertTo-RecoCredentialField $LearningUser)
    'learning.password=' + (ConvertTo-RecoCredentialField $LearningPassword)
    'business.server=' + (ConvertTo-RecoCredentialField $BusinessServer)
    'business.user=' + (ConvertTo-RecoCredentialField $BusinessUser)
    'business.password=' + (ConvertTo-RecoCredentialField $BusinessPassword)
  )
  $plaintext = [Text.Encoding]::UTF8.GetBytes(($lines -join "`n"))
  $salt = New-RecoRandomBytes 16
  $saltBase64 = [Convert]::ToBase64String($salt)
  $packageId = [Guid]::NewGuid().ToString('N')
  $context = Get-RecoPluginSqlConfigContext -PackageId $packageId -Iterations $Iterations -SaltBase64 $saltBase64
  $keyBundle = $null
  try {
    $keyBundle = Get-RecoPluginSqlConfigKeyBundle -Salt $salt -Iterations $Iterations
    $protected = Protect-RecoCredentialPayloadWithKeyBundle -Plaintext $plaintext -KeyBundle $keyBundle -Context $context
    $package = [ordered]@{
      version = 1
      package_type = $script:RecoPluginSqlConfigPackageType
      package_id = $packageId
      iterations = $Iterations
      salt = $saltBase64
      iv = [Convert]::ToBase64String($protected.Iv)
      ciphertext = [Convert]::ToBase64String($protected.Ciphertext)
      mac = [Convert]::ToBase64String($protected.Mac)
      created_at_utc = [DateTime]::UtcNow.ToString('o')
    }
    return ($package | ConvertTo-Json -Depth 3)
  }
  finally {
    [Array]::Clear($plaintext, 0, $plaintext.Length)
    if ($keyBundle) { [Array]::Clear($keyBundle, 0, $keyBundle.Length) }
  }
}

# 读回 RecoPluginSql.json，返回 learning/business 两条记录（Server/User/Password）。供发布包构建校验与测试使用。
function Read-RecoPluginSqlConfig {
  param([Parameter(Mandatory)][string]$Path)
  $target = [IO.Path]::GetFullPath($Path)
  if (-not [IO.File]::Exists($target)) { throw "插件 SQL 配置文件不存在：$target" }
  if (([IO.File]::GetAttributes($target) -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
    throw '插件 SQL 配置文件不能是重解析点。'
  }
  $package = Get-Content -LiteralPath $target -Raw -Encoding UTF8 | ConvertFrom-Json
  if ([int]$package.version -ne 1) { throw '插件 SQL 配置文件版本不受支持。' }
  if ([string]$package.package_type -ne $script:RecoPluginSqlConfigPackageType) { throw '插件 SQL 配置文件类型不受支持。' }
  $iterations = [int]$package.iterations
  $salt = [Convert]::FromBase64String([string]$package.salt)
  $context = Get-RecoPluginSqlConfigContext -PackageId ([string]$package.package_id) -Iterations $iterations -SaltBase64 ([string]$package.salt)
  $keyBundle = $null
  $plaintext = $null
  try {
    $keyBundle = Get-RecoPluginSqlConfigKeyBundle -Salt $salt -Iterations $iterations
    $plaintext = Unprotect-RecoCredentialPayloadWithKeyBundle `
      -InitializationVector ([Convert]::FromBase64String([string]$package.iv)) `
      -Ciphertext ([Convert]::FromBase64String([string]$package.ciphertext)) `
      -Mac ([Convert]::FromBase64String([string]$package.mac)) `
      -KeyBundle $keyBundle -Context $context `
      -IntegrityFailureMessage '插件 SQL 配置文件完整性校验失败，文件可能已损坏或被修改。'
    $values = @{}
    foreach ($line in ([Text.Encoding]::UTF8.GetString($plaintext) -split "`r?`n")) {
      if ([string]::IsNullOrWhiteSpace($line)) { continue }
      $separator = $line.IndexOf('=')
      if ($separator -le 0) { throw '插件 SQL 配置明文格式无效。' }
      $values[$line.Substring(0, $separator).Trim().ToLowerInvariant()] = $line.Substring($separator + 1).Trim()
    }
    if ($values['version'] -ne '1') { throw '插件 SQL 配置明文版本不受支持。' }
    $result = @{}
    foreach ($entry in @('learning', 'business')) {
      $fields = @{}
      foreach ($field in @('server', 'user', 'password')) {
        $key = $entry + '.' + $field
        if (-not $values.ContainsKey($key)) { throw "插件 SQL 配置缺少 $key。" }
        $value = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String([string]$values[$key]))
        if ([string]::IsNullOrWhiteSpace($value)) { throw "插件 SQL 配置的 $key 为空。" }
        $fields[$field] = $value
      }
      $result[$entry] = [pscustomobject]@{
        Server = [string]$fields['server']
        User = [string]$fields['user']
        Password = [string]$fields['password']
      }
    }
    return [pscustomobject]@{
      PackageId = [string]$package.package_id
      CreatedAtUtc = [string]$package.created_at_utc
      Learning = $result['learning']
      Business = $result['business']
    }
  }
  finally {
    if ($keyBundle) { [Array]::Clear($keyBundle, 0, $keyBundle.Length) }
    if ($plaintext) { [Array]::Clear($plaintext, 0, $plaintext.Length) }
  }
}
