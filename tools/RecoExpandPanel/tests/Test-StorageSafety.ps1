$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if ([String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL) -or [String]::IsNullOrWhiteSpace($env:RECO_QUOTA_DLL)) {
    throw 'Set RECO_EXPAND_DLL and RECO_QUOTA_DLL to source-built assemblies.'
}

$expandAssembly = [System.Reflection.Assembly]::LoadFrom($env:RECO_EXPAND_DLL)
$quotaAssembly = [System.Reflection.Assembly]::LoadFrom($env:RECO_QUOTA_DLL)
$flags = [System.Reflection.BindingFlags]'NonPublic,Static,Public'

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

function Assert-LockTimeoutDoesNotRun($type, [string]$label) {
    $method = $type.GetMethod('TryWithMappingBoxesLock', $flags)
    if ($null -eq $method) {
        throw "$label lock helper was not found."
    }

    $thread = [NamedMutexProbe]::Hold('RecoQuotaData.mapping-boxes.lock', 400)
    if (-not [NamedMutexProbe]::Ready.WaitOne(2000)) {
        throw "$label mutex holder did not start."
    }

    $script:lockActionRan = $false
    $action = [Action]{ $script:lockActionRan = $true }
    $result = [bool]$method.Invoke($null, @($action, 50))
    $thread.Join()
    if ($result -or $script:lockActionRan) {
        throw "$label executed its write action after a lock timeout."
    }
}

$mappingType = $quotaAssembly.GetType('RecoQuotaRecommend.MappingStore', $true)
$formType = $expandAssembly.GetType('RecoNet.FormPanel', $true)
Assert-LockTimeoutDoesNotRun $mappingType 'MappingStore'
Assert-LockTimeoutDoesNotRun $formType 'RecoExpandPanel'

$work = Join-Path $env:TEMP ('reco-storage-test-' + [Guid]::NewGuid().ToString('N'))
[System.IO.Directory]::CreateDirectory($work) | Out-Null
try {
    $path = Join-Path $work 'mapping-boxes.jsonl'
    [System.IO.File]::WriteAllText($path, 'old', [System.Text.Encoding]::UTF8)
    $atomicWrite = $formType.GetMethod('WriteAllLinesAtomic', $flags)
    $writeArgs = New-Object object[] 3
    $writeArgs[0] = [string]$path
    $writeArgs[1] = [string[]]@('new')
    $writeArgs[2] = [System.Text.Encoding]::UTF8
    $atomicWrite.Invoke($null, $writeArgs) | Out-Null
    if (([System.IO.File]::ReadAllText($path)).Trim() -ne 'new') {
        throw 'Atomic mapping write did not publish the new file.'
    }
    if (-not [System.IO.File]::Exists($path + '.bak') -or ([System.IO.File]::ReadAllText($path + '.bak')).Trim() -ne 'old') {
        throw 'Atomic mapping write did not preserve the previous file as .bak.'
    }

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

Write-Host 'PASS: lock timeout, atomic replacement backup, and log rotation.'
