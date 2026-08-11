$ErrorActionPreference = 'Stop'

$credentialRoot = Split-Path -Parent $PSScriptRoot
$workspaceRoot = Split-Path -Parent (Split-Path -Parent $credentialRoot)
$module = Join-Path $credentialRoot 'RecoCredentialStore.ps1'
$sharedSource = Join-Path $workspaceRoot 'RecoShared\RecoSqlCredentialStore.cs'
. $module

$systemTemp = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testRoot = [IO.Path]::Combine($systemTemp, 'RecoCredentialTest-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)

try {
  $store = Join-Path $testRoot 'credentials.dpapi'
  $written = Write-RecoSqlCredentialStore `
    -LearningServer 'learning.test.local' `
    -LearningUser 'learning-user' `
    -LearningPassword 'learning-password' `
    -BusinessServer 'business.test.local' `
    -BusinessUser 'business-user' `
    -BusinessPassword 'business-password' `
    -Path $store

  if ($written -ne $store -or -not [IO.File]::Exists($store)) {
    throw 'PowerShell writer did not create the expected store.'
  }
  $learning = Get-RecoSqlCredential -Name Learning -Path $store
  $business = Get-RecoSqlCredential -Name Business -Path $store
  if ($learning.Server -ne 'learning.test.local' -or
      $learning.User -ne 'learning-user' -or
      $learning.Password -ne 'learning-password' -or
      $business.Server -ne 'business.test.local' -or
      $business.User -ne 'business-user' -or
      $business.Password -ne 'business-password') {
    throw 'PowerShell credential round-trip failed.'
  }

  $duplicateRejected = $false
  try {
    [void](Write-RecoSqlCredentialStore `
      -LearningServer 'x' -LearningUser 'x' -LearningPassword 'x' `
      -BusinessServer 'x' -BusinessUser 'x' -BusinessPassword 'x' -Path $store)
  }
  catch {
    $duplicateRejected = $true
  }
  if (-not $duplicateRejected) {
    throw 'Existing credential store was overwritten.'
  }

  $harnessSource = Join-Path $testRoot 'CredentialHarness.cs'
  $harnessExe = Join-Path $testRoot 'CredentialHarness.exe'
  $harness = @'
using System;

internal static class CredentialHarness
{
    private static int Main(string[] args)
    {
        Environment.SetEnvironmentVariable("RECO_SQL_CREDENTIAL_STORE_PATH", args[0]);
        RecoSqlCredential learning = RecoSqlCredentialStore.Read("learning");
        RecoSqlCredential business = RecoSqlCredentialStore.Read("business");
        if (learning.Server != "learning.test.local" ||
            learning.User != "learning-user" ||
            learning.Password != "learning-password" ||
            business.Server != "business.test.local" ||
            business.User != "business-user" ||
            business.Password != "business-password")
        {
            return 2;
        }
        string connectionString = RecoSqlCredentialStore.BuildConnectionString("business", "RecoData2020", 1433, 8);
        if (connectionString.IndexOf("business.test.local,1433", StringComparison.Ordinal) < 0 ||
            connectionString.IndexOf("RecoData2020", StringComparison.Ordinal) < 0)
        {
            return 3;
        }
        Console.WriteLine("CSharpRoundTrip=PASS");
        return 0;
    }
}
'@
  [IO.File]::WriteAllText($harnessSource, $harness, (New-Object Text.UTF8Encoding($false)))

  $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
  if (-not [IO.File]::Exists($csc)) {
    $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe'
  }
  & $csc /nologo /target:exe /out:$harnessExe `
    /reference:System.Data.dll `
    /reference:System.Security.dll `
    $sharedSource `
    $harnessSource
  if ($LASTEXITCODE -ne 0) {
    throw "Credential harness build failed with exit code $LASTEXITCODE."
  }
  $harnessOutput = @(& $harnessExe $store)
  if ($LASTEXITCODE -ne 0 -or $harnessOutput -notcontains 'CSharpRoundTrip=PASS') {
    throw 'C# credential round-trip failed.'
  }

  $corruptStore = Join-Path $testRoot 'corrupt.dpapi'
  [IO.File]::WriteAllBytes($corruptStore, [byte[]](1, 2, 3, 4))
  $corruptRejected = $false
  try {
    [void](Get-RecoSqlCredential -Name Learning -Path $corruptStore)
  }
  catch {
    $corruptRejected = $true
  }
  if (-not $corruptRejected) {
    throw 'Corrupt credential store was accepted.'
  }

  'PASS: DPAPI store PowerShell round-trip'
  'PASS: existing store overwrite rejected'
  'PASS: DPAPI store C# round-trip'
  'PASS: corrupt store rejected'
}
finally {
  $fullTestRoot = [IO.Path]::GetFullPath($testRoot)
  $boundary = $systemTemp.TrimEnd('\') + '\'
  if ($fullTestRoot.StartsWith($boundary, [StringComparison]::OrdinalIgnoreCase) -and
      [IO.Directory]::Exists($fullTestRoot) -and
      (([IO.File]::GetAttributes($fullTestRoot) -band [IO.FileAttributes]::ReparsePoint) -eq 0)) {
    [IO.Directory]::Delete($fullTestRoot, $true)
  }
}
