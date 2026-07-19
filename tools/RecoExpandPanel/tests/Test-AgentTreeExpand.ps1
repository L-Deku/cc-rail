# 懒加载树祖先前缀链单元测试:反射调用编译产物中的 BuildAgentItemAncestorPrefixes。
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $PSScriptRoot))
$dll = Join-Path $repoRoot 'RecoQuotaRecommend\bin\RecoExpandPanel.dll'
if (-not (Test-Path -LiteralPath $dll)) { throw "缺少编译产物: $dll" }
Add-Type -Path $dll

$method = [RecoNet.FormPanel].GetMethod('BuildAgentItemAncestorPrefixes', [System.Reflection.BindingFlags]'NonPublic,Public,Static')
if ($method -eq $null) { throw '缺少 BuildAgentItemAncestorPrefixes' }

function Assert-Prefixes([string]$ItemNo, [string[]]$Expected) {
    $actual = @($method.Invoke($null, @($ItemNo)))
    if (($actual -join ',') -ne ($Expected -join ',')) {
        throw "前缀链不符 [$ItemNo]: 期望 $($Expected -join ',') 实际 $($actual -join ',')"
    }
}

Assert-Prefixes '0821-01-04-09-03' @('08','0821','0821-01','0821-01-04','0821-01-04-09')
Assert-Prefixes '0719-01-02' @('07','0719','0719-01')
Assert-Prefixes '0101' @('01')
Assert-Prefixes '07' @()
Write-Host 'Test-AgentTreeExpand: PASS'
