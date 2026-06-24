$ErrorActionPreference = "Continue"
$bin = "C:\Users\谢刚\Desktop\自动预算\RecoQuotaRecommend\bin"
$root = "C:\Users\谢刚\Desktop\自动预算\RecoQuotaData"
$dirs = @(
  "C:\Users\谢刚\Desktop\自动预算\铁路基本建设工程投资控制系统2020网络版V0503021201",
  "C:\Users\谢刚\Desktop\自动预算\2024铁路工程云计价系统网络版V1.0\铁路工程云计价系统网络版V1.0"
)
$dllNames = @("RecoQuotaRecommend.dll", "RecoExpandPanel.dll", "RecoPluginLoader.dll", "0Harmony.dll")
$dataNames = @("chapter-entries.jsonl", "chapter-quota-library.jsonl")
$done = @{}
$deadline = (Get-Date).AddMinutes(60)

while ($done.Count -lt $dirs.Count -and (Get-Date) -lt $deadline) {
  foreach ($d in $dirs) {
    if ($done.ContainsKey($d)) { continue }
    $target = Join-Path $d "RecoQuotaRecommend.dll"
    $ok = $true
    # probe write lock: exclusive open succeeds only when the program is closed
    try {
      if (Test-Path $target) {
        $fs = [System.IO.File]::Open($target, 'Open', 'Write', 'None')
        $fs.Close()
      }
    } catch {
      $ok = $false
    }
    if (-not $ok) { continue }

    try {
      foreach ($n in $dllNames) {
        $src = Join-Path $bin $n
        if (Test-Path $src) { Copy-Item -LiteralPath $src -Destination $d -Force }
      }
      $dataTarget = Join-Path $d "RecoQuotaData"
      New-Item -ItemType Directory -Path $dataTarget -Force | Out-Null
      foreach ($n in $dataNames) {
        $src = Join-Path $root $n
        if (Test-Path $src) { Copy-Item -LiteralPath $src -Destination $dataTarget -Force }
      }
      $done[$d] = $true
      Write-Output ("DEPLOYED " + $d)
    } catch {
      Write-Output ("RETRY " + $d + " : " + $_.Exception.Message)
    }
  }
  if ($done.Count -lt $dirs.Count) { Start-Sleep -Seconds 3 }
}

if ($done.Count -eq $dirs.Count) { Write-Output "ALL DEPLOYED" } else { Write-Output "TIMEOUT - some dirs still locked" }
