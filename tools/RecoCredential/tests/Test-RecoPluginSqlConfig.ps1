$ErrorActionPreference = 'Stop'

# 验证插件 SQL 配置文件 RecoPluginSql.json：PowerShell 生成 → C# 解密往返、篡改被拒、reco 登录名被拒、
# 无配置文件时回退 DPAPI、按服务器路由。全程不打开任何 SQL 连接。

$credentialRoot = Split-Path -Parent $PSScriptRoot
$workspaceRoot = Split-Path -Parent (Split-Path -Parent $credentialRoot)
$sharedSource = Join-Path $workspaceRoot 'RecoShared\RecoSqlCredentialStore.cs'
. (Join-Path $credentialRoot 'RecoCredentialStore.ps1')
. (Join-Path $credentialRoot 'RecoCredentialTransfer.ps1')
$generator = Join-Path $credentialRoot 'New-RecoPluginSqlConfig.ps1'

$systemTemp = [IO.Path]::GetFullPath([IO.Path]::GetTempPath())
$testRoot = [IO.Path]::Combine($systemTemp, 'RecoPluginSqlConfigTest-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($testRoot)
$previousConfigOverride = [Environment]::GetEnvironmentVariable('RECO_PLUGIN_SQL_CONFIG_PATH')
$previousStoreOverride = [Environment]::GetEnvironmentVariable('RECO_SQL_CREDENTIAL_STORE_PATH')

function Assert-Throws {
  param([scriptblock]$Action, [string]$Message)
  $thrown = $false
  try { & $Action } catch { $thrown = $true }
  if (-not $thrown) { throw $Message }
}

try {
  # 1. 通过正式生成脚本生成（跳过连接核对）
  $configPath = Join-Path $testRoot 'RecoPluginSql.json'
  $password = ConvertTo-SecureString 'plugin-password-1' -AsPlainText -Force
  & $generator -LearningServer 'learning.test.local,1433' -BusinessServer 'business.test.local' -User 'reco_plugin' `
    -Password $password -OutputPath $configPath -Iterations 10000 -SkipConnectionTest | Out-Null
  if (-not (Test-Path -LiteralPath $configPath)) { throw '生成脚本没有产出配置文件。' }

  $readBack = Read-RecoPluginSqlConfig -Path $configPath
  if ($readBack.Learning.Server -ne 'learning.test.local,1433' -or $readBack.Learning.User -ne 'reco_plugin' -or
      $readBack.Learning.Password -ne 'plugin-password-1' -or $readBack.Business.Server -ne 'business.test.local' -or
      $readBack.Business.Password -ne 'plugin-password-1') {
    throw 'PowerShell 回读与输入不一致。'
  }
  'PASS: PowerShell 生成与回读往返'

  # 2. 生成脚本拒绝 reco 登录名
  Assert-Throws { & $generator -User 'reco' -Password $password -OutputPath (Join-Path $testRoot 'x.json') -SkipConnectionTest -Iterations 10000 | Out-Null } '生成脚本接受了 reco 登录名。'
  Assert-Throws { [void](New-RecoPluginSqlConfigJson -LearningServer 'a' -LearningUser 'RECO' -LearningPassword 'p' -BusinessServer 'b' -BusinessUser 'x' -BusinessPassword 'p' -Iterations 10000) } '生成函数接受了 RECO 登录名。'
  'PASS: 生成端拒绝 reco 登录名'

  # 3. 篡改 mac / ciphertext 被拒
  $raw = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
  $tamperedMac = Join-Path $testRoot 'tampered-mac.json'
  $macBytes = [Convert]::FromBase64String($raw.mac); $macBytes[0] = $macBytes[0] -bxor 1
  $raw.mac = [Convert]::ToBase64String($macBytes)
  [IO.File]::WriteAllText($tamperedMac, ($raw | ConvertTo-Json -Depth 3), (New-Object Text.UTF8Encoding($false)))
  Assert-Throws { [void](Read-RecoPluginSqlConfig -Path $tamperedMac) } '篡改 mac 后仍被接受。'
  $raw = Get-Content -LiteralPath $configPath -Raw -Encoding UTF8 | ConvertFrom-Json
  $tamperedCipher = Join-Path $testRoot 'tampered-cipher.json'
  $cipherBytes = [Convert]::FromBase64String($raw.ciphertext); $cipherBytes[3] = $cipherBytes[3] -bxor 1
  $raw.ciphertext = [Convert]::ToBase64String($cipherBytes)
  [IO.File]::WriteAllText($tamperedCipher, ($raw | ConvertTo-Json -Depth 3), (New-Object Text.UTF8Encoding($false)))
  Assert-Throws { [void](Read-RecoPluginSqlConfig -Path $tamperedCipher) } '篡改 ciphertext 后仍被接受。'
  'PASS: PowerShell 端篡改被拒'

  # 4. 一份含 reco 登录名的配置（绕过生成端守卫）供 C# 端拒绝测试
  $adminConfig = Join-Path $testRoot 'admin.json'
  $adminJson = New-RecoPluginSqlConfigJson -LearningServer 'l' -LearningUser 'reco' -LearningPassword 'p' `
    -BusinessServer 'b' -BusinessUser 'reco' -BusinessPassword 'p' -Iterations 10000 -AllowAdministrativeLogin
  [IO.File]::WriteAllText($adminConfig, $adminJson, (New-Object Text.UTF8Encoding($false)))

  # 5. DPAPI 回退用的凭据库
  $dpapiStore = Join-Path $testRoot 'fallback.dpapi'
  [void](Write-RecoSqlCredentialStore -LearningServer 'dpapi-learning.local' -LearningUser 'reco' -LearningPassword 'dpapi-pass' `
    -BusinessServer 'dpapi-business.local' -BusinessUser 'reco' -BusinessPassword 'dpapi-pass' -Path $dpapiStore)

  # 6. C# 端
  $harnessSource = Join-Path $testRoot 'PluginSqlConfigHarness.cs'
  $harnessExe = Join-Path $testRoot 'PluginSqlConfigHarness.exe'
  $harness = @'
using System;
using System.Data.SqlClient;

internal static class PluginSqlConfigHarness
{
    private static int Main(string[] args)
    {
        string configPath = args[0];
        string tamperedMac = args[1];
        string tamperedCipher = args[2];
        string adminConfig = args[3];
        string dpapiStore = args[4];
        string missingConfig = args[5];

        Environment.SetEnvironmentVariable("RECO_PLUGIN_SQL_CONFIG_PATH", configPath);
        Environment.SetEnvironmentVariable("RECO_SQL_CREDENTIAL_STORE_PATH", dpapiStore);
        RecoSqlCredential learning = RecoSqlCredentialStore.Read("learning");
        RecoSqlCredential business = RecoSqlCredentialStore.Read("business");
        if (learning.Server != "learning.test.local,1433" || learning.User != "reco_plugin" || learning.Password != "plugin-password-1" ||
            business.Server != "business.test.local" || business.User != "reco_plugin" || business.Password != "plugin-password-1")
        {
            Console.WriteLine("FAIL: plugin config round-trip");
            return 2;
        }
        Console.WriteLine("PASS: C# 解密插件配置往返");

        SqlConnectionStringBuilder b1 = new SqlConnectionStringBuilder(RecoSqlCredentialStore.BuildConnectionStringForServer("learning.test.local", "RecoData2024", 1433, 8));
        SqlConnectionStringBuilder b2 = new SqlConnectionStringBuilder(RecoSqlCredentialStore.BuildConnectionStringForServer("LEARNING.test.local,1433", "RecoLearning", 1433, 8));
        SqlConnectionStringBuilder b3 = new SqlConnectionStringBuilder(RecoSqlCredentialStore.BuildConnectionStringForServer("business.test.local", "RecoData2020", 1433, 8));
        if (b1.DataSource != "learning.test.local,1433" || b1.InitialCatalog != "RecoData2024" || b1.UserID != "reco_plugin" ||
            b2.DataSource != "learning.test.local,1433" || b2.InitialCatalog != "RecoLearning" ||
            b3.DataSource != "business.test.local,1433" || b3.InitialCatalog != "RecoData2020" || b3.PersistSecurityInfo)
        {
            Console.WriteLine("FAIL: server routing");
            return 3;
        }
        if (!Throws(delegate { RecoSqlCredentialStore.BuildConnectionStringForServer("unknown.local", "RecoData2020", 1433, 8); }))
        {
            Console.WriteLine("FAIL: unknown server accepted");
            return 4;
        }
        Console.WriteLine("PASS: C# 按服务器路由与未知服务器拒绝");

        Environment.SetEnvironmentVariable("RECO_PLUGIN_SQL_CONFIG_PATH", tamperedMac);
        if (!Throws(delegate { RecoSqlCredentialStore.Read("learning"); })) { Console.WriteLine("FAIL: tampered mac accepted"); return 5; }
        Environment.SetEnvironmentVariable("RECO_PLUGIN_SQL_CONFIG_PATH", tamperedCipher);
        if (!Throws(delegate { RecoSqlCredentialStore.Read("learning"); })) { Console.WriteLine("FAIL: tampered ciphertext accepted"); return 6; }
        Console.WriteLine("PASS: C# 端篡改被拒");

        Environment.SetEnvironmentVariable("RECO_PLUGIN_SQL_CONFIG_PATH", adminConfig);
        if (!Throws(delegate { RecoSqlCredentialStore.Read("learning"); })) { Console.WriteLine("FAIL: reco login accepted"); return 7; }
        Console.WriteLine("PASS: C# 端拒绝含 reco 登录名的配置");

        Environment.SetEnvironmentVariable("RECO_PLUGIN_SQL_CONFIG_PATH", missingConfig);
        RecoSqlCredential fallback = RecoSqlCredentialStore.Read("business");
        if (fallback.Server != "dpapi-business.local" || fallback.User != "reco" || fallback.Password != "dpapi-pass")
        {
            Console.WriteLine("FAIL: DPAPI fallback");
            return 8;
        }
        Console.WriteLine("PASS: 无配置文件时回退 DPAPI 凭据库");
        return 0;
    }

    private static bool Throws(Action action)
    {
        try { action(); return false; }
        catch (Exception) { return true; }
    }
}
'@
  [IO.File]::WriteAllText($harnessSource, $harness, (New-Object Text.UTF8Encoding($false)))
  $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework64\v4.0.30319\csc.exe'
  if (-not [IO.File]::Exists($csc)) { $csc = Join-Path $env:WINDIR 'Microsoft.NET\Framework\v4.0.30319\csc.exe' }
  & $csc /nologo /target:exe /out:$harnessExe /reference:System.Data.dll /reference:System.Security.dll $sharedSource $harnessSource
  if ($LASTEXITCODE -ne 0) { throw "C# 测试宿主编译失败：$LASTEXITCODE" }
  $missingConfig = Join-Path $testRoot 'does-not-exist.json'
  $output = @(& $harnessExe $configPath $tamperedMac $tamperedCipher $adminConfig $dpapiStore $missingConfig)
  $output | ForEach-Object { $_ }
  if ($LASTEXITCODE -ne 0) { throw "C# 端测试失败，退出码 $LASTEXITCODE" }
}
finally {
  [Environment]::SetEnvironmentVariable('RECO_PLUGIN_SQL_CONFIG_PATH', $previousConfigOverride)
  [Environment]::SetEnvironmentVariable('RECO_SQL_CREDENTIAL_STORE_PATH', $previousStoreOverride)
  $fullTestRoot = [IO.Path]::GetFullPath($testRoot)
  $boundary = $systemTemp.TrimEnd('\') + '\'
  if ($fullTestRoot.StartsWith($boundary, [StringComparison]::OrdinalIgnoreCase) -and [IO.Directory]::Exists($fullTestRoot)) {
    [IO.Directory]::Delete($fullTestRoot, $true)
  }
}
