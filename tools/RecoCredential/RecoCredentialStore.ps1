function Get-RecoSqlCredentialStorePath {
  $configured = [Environment]::GetEnvironmentVariable('RECO_SQL_CREDENTIAL_STORE_PATH')
  if (-not [string]::IsNullOrWhiteSpace($configured)) {
    if (-not [IO.Path]::IsPathRooted($configured)) {
      throw 'RECO_SQL_CREDENTIAL_STORE_PATH must be an absolute path.'
    }
    return [IO.Path]::GetFullPath($configured)
  }

  $localAppData = [Environment]::GetFolderPath([Environment+SpecialFolder]::LocalApplicationData)
  return [IO.Path]::Combine($localAppData, 'RecoBudget', 'Secrets', 'sql-credentials.dpapi')
}

function ConvertTo-RecoCredentialField {
  param([Parameter(Mandatory)][string]$Value)
  if ([string]::IsNullOrWhiteSpace($Value)) {
    throw 'SQL credential fields cannot be empty.'
  }
  return [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Value))
}

function Write-RecoSqlCredentialStore {
  param(
    [Parameter(Mandatory)][string]$LearningServer,
    [Parameter(Mandatory)][string]$LearningUser,
    [Parameter(Mandatory)][string]$LearningPassword,
    [Parameter(Mandatory)][string]$BusinessServer,
    [Parameter(Mandatory)][string]$BusinessUser,
    [Parameter(Mandatory)][string]$BusinessPassword,
    [string]$Path = (Get-RecoSqlCredentialStorePath)
  )

  $target = [IO.Path]::GetFullPath($Path)
  if ([IO.File]::Exists($target)) {
    throw "Refusing to overwrite the existing DPAPI SQL credential store: $target"
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
  $entropy = [Text.Encoding]::UTF8.GetBytes('RecoBudget.SqlCredentials.v1')
  $encrypted = $null
  $temp = $null
  try {
    Add-Type -AssemblyName System.Security
    $encrypted = [Security.Cryptography.ProtectedData]::Protect(
      $plaintext,
      $entropy,
      [Security.Cryptography.DataProtectionScope]::CurrentUser)
    $directory = [IO.Path]::GetDirectoryName($target)
    [void][IO.Directory]::CreateDirectory($directory)
    $temp = [IO.Path]::Combine($directory, ([IO.Path]::GetFileName($target) + '.' + [Guid]::NewGuid().ToString('N') + '.tmp'))
    [IO.File]::WriteAllBytes($temp, $encrypted)
    [IO.File]::Move($temp, $target)
    $temp = $null
  }
  finally {
    if ($plaintext) { [Array]::Clear($plaintext, 0, $plaintext.Length) }
    if ($encrypted) { [Array]::Clear($encrypted, 0, $encrypted.Length) }
    if ($temp -and [IO.File]::Exists($temp)) { [IO.File]::Delete($temp) }
  }

  return $target
}

function Get-RecoSqlCredential {
  param(
    [Parameter(Mandatory)][ValidateSet('Learning', 'Business')][string]$Name,
    [string]$Path = (Get-RecoSqlCredentialStorePath)
  )

  $target = [IO.Path]::GetFullPath($Path)
  if (-not [IO.File]::Exists($target)) {
    throw "The local DPAPI SQL credential store is missing: $target"
  }
  if (([IO.File]::GetAttributes($target) -band [IO.FileAttributes]::ReparsePoint) -ne 0) {
    throw 'The local DPAPI SQL credential store cannot be a reparse point.'
  }

  $encrypted = [IO.File]::ReadAllBytes($target)
  $plaintext = $null
  try {
    Add-Type -AssemblyName System.Security
    $entropy = [Text.Encoding]::UTF8.GetBytes('RecoBudget.SqlCredentials.v1')
    $plaintext = [Security.Cryptography.ProtectedData]::Unprotect(
      $encrypted,
      $entropy,
      [Security.Cryptography.DataProtectionScope]::CurrentUser)
    $values = @{}
    foreach ($line in ([Text.Encoding]::UTF8.GetString($plaintext) -split "`r?`n")) {
      if ([string]::IsNullOrWhiteSpace($line)) { continue }
      $separator = $line.IndexOf('=')
      if ($separator -le 0 -or $separator -eq ($line.Length - 1)) {
        throw 'The local DPAPI SQL credential store has an invalid payload.'
      }
      $key = $line.Substring(0, $separator).Trim().ToLowerInvariant()
      if ($values.ContainsKey($key)) {
        throw 'The local DPAPI SQL credential store contains a duplicate key.'
      }
      $values[$key] = $line.Substring($separator + 1).Trim()
    }
    if ($values['version'] -ne '1') {
      throw 'The local DPAPI SQL credential store version is unsupported.'
    }

    $prefix = $Name.ToLowerInvariant() + '.'
    $result = @{}
    foreach ($field in @('server', 'user', 'password')) {
      $key = $prefix + $field
      if (-not $values.ContainsKey($key)) {
        throw "The local DPAPI SQL credential store is missing $key."
      }
      try {
        $value = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String([string]$values[$key]))
      }
      catch {
        throw "The local DPAPI SQL credential store has invalid encoding for $key."
      }
      if ([string]::IsNullOrWhiteSpace($value)) {
        throw "The local DPAPI SQL credential store has an empty $key."
      }
      $result[$field] = $value
    }
    return [pscustomobject]@{
      Server = [string]$result['server']
      User = [string]$result['user']
      Password = [string]$result['password']
    }
  }
  catch [Security.Cryptography.CryptographicException] {
    throw 'The local DPAPI SQL credential store cannot be decrypted by the current Windows user.'
  }
  finally {
    if ($plaintext) { [Array]::Clear($plaintext, 0, $plaintext.Length) }
    if ($encrypted) { [Array]::Clear($encrypted, 0, $encrypted.Length) }
  }
}
