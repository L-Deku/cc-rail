$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if ([String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL) -or [String]::IsNullOrWhiteSpace($env:RECO_QUOTA_DLL)) {
    throw 'Set RECO_EXPAND_DLL and RECO_QUOTA_DLL to source-built assemblies.'
}

$expandAssembly = [System.Reflection.Assembly]::LoadFrom($env:RECO_EXPAND_DLL)
$quotaAssembly = [System.Reflection.Assembly]::LoadFrom($env:RECO_QUOTA_DLL)
$flags = [System.Reflection.BindingFlags]'NonPublic,Static,Public,Instance'

Add-Type -TypeDefinition @'
using System;
using System.Threading;

public static class NamedMutexProbe
{
    public static readonly ManualResetEvent Ready = new ManualResetEvent(false);

    public static Thread Hold(string name, int milliseconds)
    {
        Ready.Reset();
        Thread thread = new Thread(delegate()
        {
            using (Mutex mutex = new Mutex(false, name))
            {
                mutex.WaitOne();
                Ready.Set();
                Thread.Sleep(milliseconds);
                mutex.ReleaseMutex();
            }
        });
        thread.IsBackground = true;
        thread.Start();
        return thread;
    }
}
'@

function Assert-SharedLockTimeoutDoesNotWrite($assembly, [string]$label, [string]$path) {
    $type = $assembly.GetType('LocalMappingFileStore', $true)
    $method = $type.GetMethod('Save', $flags)
    if ($null -eq $method) {
        throw "$label shared mapping save helper was not found."
    }

    $before = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    $thread = [NamedMutexProbe]::Hold('RecoQuotaData.mapping-boxes.lock', 400)
    if (-not [NamedMutexProbe]::Ready.WaitOne(2000)) {
        throw "$label mutex holder did not start."
    }

    $arguments = New-Object object[] 5
    $arguments[0] = $path
    $arguments[1] = '2024'
    $arguments[2] = 'storage-safety-test'
    $arguments[3] = 50
    $arguments[4] = $null
    $result = $method.Invoke($null, $arguments)
    $thread.Join()
    if ($null -eq $result) {
        throw "$label shared mapping save returned null: $($method.ToString())."
    }
    $status = [string]$result.GetType().GetField('Status', $flags).GetValue($result)
    $after = (Get-FileHash -LiteralPath $path -Algorithm SHA256).Hash
    if ($status -ne 'LockTimeout' -or $after -ne $before) {
        throw "$label did not preserve the file after a shared-lock timeout: status=$status."
    }
}

$formType = $expandAssembly.GetType('RecoNet.FormPanel', $true)

$work = Join-Path $env:TEMP ('reco-storage-test-' + [Guid]::NewGuid().ToString('N'))
[System.IO.Directory]::CreateDirectory($work) | Out-Null
try {
    $path = Join-Path $work 'mapping-boxes.jsonl'
    [System.IO.File]::WriteAllText($path, 'old', [System.Text.Encoding]::UTF8)
    Assert-SharedLockTimeoutDoesNotWrite $quotaAssembly 'MappingStore' $path
    Assert-SharedLockTimeoutDoesNotWrite $expandAssembly 'RecoExpandPanel' $path

    foreach ($pair in @(
        @($formType, 'RecoExpandPanel'),
        @($quotaAssembly.GetType('RecoQuotaRecommend.QuotaRecommendPanel', $true), 'RecoQuotaRecommend')
    )) {
        $logPath = Join-Path $work ($pair[1] + '.log')
        $stream = [System.IO.File]::Open($logPath, [System.IO.FileMode]::Create, [System.IO.FileAccess]::Write)
        try { $stream.SetLength(5L * 1024L * 1024L) } finally { $stream.Dispose() }
        $rotate = $pair[0].GetMethod('RotateLogIfNeeded', $flags)
        $rotateArgs = New-Object object[] 1
        $rotateArgs[0] = [string]$logPath
        $rotate.Invoke($null, $rotateArgs) | Out-Null
        if ([System.IO.File]::Exists($logPath) -or -not [System.IO.File]::Exists($logPath + '.1')) {
            throw "$($pair[1]) log rotation did not move the full log to .1."
        }
    }
}
finally {
    if ([System.IO.Directory]::Exists($work)) {
        [System.IO.Directory]::Delete($work, $true)
    }
}

Write-Host 'PASS: shared lock timeout preserves mapping file, and log rotation remains active.'
