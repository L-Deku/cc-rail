$ErrorActionPreference = 'Stop'

function Assert-Equal {
  param($Actual, $Expected, [string]$Message)
  if ($Actual -ne $Expected) {
    throw $Message
  }
}

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $testDir '..\..\..')).Path
$credentialScript = Join-Path $repoRoot 'tools\RecoCredential\RecoCredentialStore.ps1'
$commonScript = Join-Path $repoRoot 'tools\RecoLearning\Common.ps1'
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
$tempRoot = Join-Path $tempBase ('RecoConnectionRouting-' + [Guid]::NewGuid().ToString('N'))
[IO.Directory]::CreateDirectory($tempRoot) | Out-Null
$storePath = Join-Path $tempRoot 'sql-credentials.dpapi'
$previousOverride = [Environment]::GetEnvironmentVariable('RECO_SQL_CREDENTIAL_STORE_PATH')

try {
  [Environment]::SetEnvironmentVariable('RECO_SQL_CREDENTIAL_STORE_PATH', $storePath)
  . $credentialScript
  [void](Write-RecoSqlCredentialStore `
    -LearningServer 'learning.invalid' `
    -LearningUser 'learning-user' `
    -LearningPassword 'learning-password' `
    -BusinessServer 'business.invalid' `
    -BusinessUser 'business-user' `
    -BusinessPassword 'business-password')
  . $commonScript

  $learning = New-Object System.Data.SqlClient.SqlConnectionStringBuilder (Get-RecoConnectionString)
  Assert-Equal $learning['Data Source'] 'learning.invalid' 'Learning route selected the wrong server.'
  Assert-Equal $learning['Initial Catalog'] 'RecoLearning' 'Learning route selected the wrong database.'
  Assert-Equal $learning['User ID'] 'learning-user' 'Learning route selected the wrong user.'

  $business = New-Object System.Data.SqlClient.SqlConnectionStringBuilder (Get-RecoConnectionString -Server 'business.invalid' -Database 'RecoData2020')
  Assert-Equal $business['Data Source'] 'business.invalid' 'Business route selected the wrong server.'
  Assert-Equal $business['Initial Catalog'] 'RecoData2020' 'Business route selected the wrong database.'
  Assert-Equal $business['User ID'] 'business-user' 'Business route selected the wrong user.'

  $unknownRejected = $false
  try {
    [void](Get-RecoConnectionString -Server 'unknown.invalid' -Database 'RecoData2020')
  }
  catch {
    $unknownRejected = $true
  }
  if (-not $unknownRejected) {
    throw 'Unknown SQL endpoint was not rejected.'
  }

  Write-Host 'PASS DPAPI connection routing: Learning, Business, unknown rejection'
  Write-Host 'PASS no SQL connection was opened'
}
finally {
  [Environment]::SetEnvironmentVariable('RECO_SQL_CREDENTIAL_STORE_PATH', $previousOverride)
  $resolvedTempRoot = [IO.Path]::GetFullPath($tempRoot)
  if (-not $resolvedTempRoot.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Refusing to clean unexpected test directory.'
  }
  if ([IO.Directory]::Exists($resolvedTempRoot)) {
    [IO.Directory]::Delete($resolvedTempRoot, $true)
  }
}
