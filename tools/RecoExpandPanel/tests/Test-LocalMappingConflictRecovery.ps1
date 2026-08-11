param(
    [string]$ChapterQuotaLibraryExe = ""
)

$ErrorActionPreference = "Stop"

function Assert-True {
    param(
        [bool]$Condition,
        [string]$Message
    )
    if (-not $Condition) {
        throw $Message
    }
}

function Quote-ProcessArgument {
    param([string]$Value)
    if ($Value.Contains('"')) {
        throw "Test path contains an unsupported quote character."
    }
    return '"' + $Value + '"'
}

function Invoke-TagMappingBoxes {
    param(
        [string]$ExePath,
        [string]$FilePath,
        [string]$DataDirectory
    )

    $startInfo = New-Object System.Diagnostics.ProcessStartInfo
    $startInfo.FileName = $ExePath
    $startInfo.Arguments = @(
        "TagMappingBoxes",
        "--file",
        (Quote-ProcessArgument $FilePath),
        "--data-dir",
        (Quote-ProcessArgument $DataDirectory)
    ) -join " "
    $startInfo.UseShellExecute = $false
    $startInfo.CreateNoWindow = $true
    $startInfo.RedirectStandardOutput = $true
    $startInfo.RedirectStandardError = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $startInfo
    try {
        Assert-True $process.Start() "Failed to start ChapterQuotaLibrary."
        $stdout = $process.StandardOutput.ReadToEnd()
        $stderr = $process.StandardError.ReadToEnd()
        $process.WaitForExit()
        return [pscustomobject]@{
            ExitCode = $process.ExitCode
            Output = $stdout + $stderr
        }
    }
    finally {
        $process.Dispose()
    }
}

function Get-Sha256 {
    param([string]$Path)
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
}

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $testDir "..\..\..")).Path
if ([string]::IsNullOrWhiteSpace($ChapterQuotaLibraryExe)) {
    $ChapterQuotaLibraryExe = Join-Path $repoRoot "tools\ChapterQuotaLibrary\bin\ChapterQuotaLibrary.exe"
}
$ChapterQuotaLibraryExe = (Resolve-Path -LiteralPath $ChapterQuotaLibraryExe).Path

$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
$tempRoot = Join-Path $tempBase ("RecoLocalMappingConflict-" + [Guid]::NewGuid().ToString("N"))
[IO.Directory]::CreateDirectory($tempRoot) | Out-Null
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

try {
    $cases = @(
        [pscustomobject]@{
            Name = "partitioned"
            Content = '{"record_type":"mapping_box","software_partition":"2024","box_id":"b1"}'
            IsEmpty = $false
            ExpectPartitionWarning = $true
        },
        [pscustomobject]@{
            Name = "empty"
            Content = ""
            IsEmpty = $true
            ExpectPartitionWarning = $false
        },
        [pscustomobject]@{
            Name = "malformed"
            Content = '{not-json'
            IsEmpty = $false
            ExpectPartitionWarning = $false
        },
        [pscustomobject]@{
            Name = "legacy"
            Content = '{"box_id":"legacy-1","target_code":"DY-1"}'
            IsEmpty = $false
            ExpectPartitionWarning = $false
        }
    )

    $partitionWarning = -join @(
        [char]0x5C1A, [char]0x672A, [char]0x5206,
        [char]0x533A, [char]0x611F, [char]0x77E5
    )

    foreach ($case in $cases) {
        $caseDir = Join-Path $tempRoot $case.Name
        [IO.Directory]::CreateDirectory($caseDir) | Out-Null
        $filePath = Join-Path $caseDir "mapping-boxes.jsonl"
        if ($case.IsEmpty) {
            [IO.File]::WriteAllBytes($filePath, [byte[]]@())
        }
        else {
            [IO.File]::WriteAllText($filePath, $case.Content, $utf8NoBom)
        }

        $before = Get-Sha256 $filePath
        $result = Invoke-TagMappingBoxes $ChapterQuotaLibraryExe $filePath $caseDir
        $after = Get-Sha256 $filePath

        Assert-True ($result.ExitCode -ne 0) ("TagMappingBoxes unexpectedly succeeded for " + $case.Name)
        Assert-True ($before -eq $after) ("TagMappingBoxes changed fixture: " + $case.Name)
        if ($case.ExpectPartitionWarning) {
            Assert-True ($result.Output.Contains($partitionWarning)) "Partition-awareness warning was not visible."
        }
        Write-Host ("PASS B10 fixture: " + $case.Name + " hash unchanged")
    }

    $lockedDir = Join-Path $tempRoot "locked"
    [IO.Directory]::CreateDirectory($lockedDir) | Out-Null
    $lockedFile = Join-Path $lockedDir "mapping-boxes.jsonl"
    [IO.File]::WriteAllText($lockedFile, '{"box_id":"locked-1"}', $utf8NoBom)
    $lockedBefore = Get-Sha256 $lockedFile
    $mutex = New-Object System.Threading.Mutex($false, "RecoQuotaData.mapping-boxes.lock")
    $ownsMutex = $false
    try {
        $ownsMutex = $mutex.WaitOne(0)
        Assert-True $ownsMutex "Test could not acquire mapping-boxes mutex."
        $lockedResult = Invoke-TagMappingBoxes $ChapterQuotaLibraryExe $lockedFile $lockedDir
    }
    finally {
        if ($ownsMutex) {
            $mutex.ReleaseMutex()
        }
        $mutex.Dispose()
    }
    $lockedAfter = Get-Sha256 $lockedFile
    Assert-True ($lockedResult.ExitCode -ne 0) "TagMappingBoxes unexpectedly succeeded while mutex was held."
    Assert-True ($lockedBefore -eq $lockedAfter) "TagMappingBoxes changed the locked fixture."
    Write-Host "PASS B10 mutex timeout: hash unchanged"
    Write-Host "PASS T26(e): third mapping-boxes mutex holder is read-only"
}
finally {
    $resolvedTempRoot = [IO.Path]::GetFullPath($tempRoot)
    if (-not $resolvedTempRoot.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clean unexpected test directory."
    }
    if ([IO.Directory]::Exists($resolvedTempRoot)) {
        [IO.Directory]::Delete($resolvedTempRoot, $true)
    }
}
