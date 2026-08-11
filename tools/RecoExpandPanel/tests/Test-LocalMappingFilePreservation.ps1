$ErrorActionPreference = "Stop"

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$repoRoot = (Resolve-Path -LiteralPath (Join-Path $testDir "..\..\..")).Path
$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
$tempRoot = Join-Path $tempBase ("RecoLocalMapping-" + [Guid]::NewGuid().ToString("N"))
[void][IO.Directory]::CreateDirectory($tempRoot)

try {
    $csc = Join-Path $env:WINDIR "Microsoft.NET\Framework64\v4.0.30319\csc.exe"
    $harness = Join-Path $testDir "LocalMappingFileStoreHarness.cs"
    $identity = Join-Path $repoRoot "RecoShared\LearningPartitionIdentity.cs"
    $store = Join-Path $repoRoot "RecoShared\LocalMappingFileStore.cs"
    $exe = Join-Path $tempRoot "LocalMappingFileStoreHarness.exe"
    & $csc /nologo /target:exe "/out:$exe" /reference:System.Web.Extensions.dll $identity $store $harness
    if ($LASTEXITCODE -ne 0) { throw "Local mapping harness compilation failed: $LASTEXITCODE" }
    & $exe (Join-Path $tempRoot "fixtures")
    if ($LASTEXITCODE -ne 0) { throw "Local mapping harness failed: $LASTEXITCODE" }
}
finally {
    $resolved = [IO.Path]::GetFullPath($tempRoot)
    if (-not $resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to clean unexpected test directory."
    }
    if ([IO.Directory]::Exists($resolved)) {
        [IO.Directory]::Delete($resolved, $true)
    }
}
