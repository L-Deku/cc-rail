param(
    [string]$ExpandDll = $env:RECO_EXPAND_DLL
)

# 覆盖代码审查 §2.1 / §2.2 / §3.4：批量绑定失败行占位、绑定存储原子写与损坏保护、纯函数替换边界。

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest
[Console]::OutputEncoding = [Text.Encoding]::UTF8

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..')).Path
if ([string]::IsNullOrWhiteSpace($ExpandDll)) {
    $ExpandDll = Join-Path $repoRoot 'RecoQuotaRecommend\bin\RecoExpandPanel.dll'
}
if (-not (Test-Path -LiteralPath $ExpandDll)) {
    throw "Missing RecoExpandPanel.dll: $ExpandDll"
}

$dllDir = Split-Path -Parent (Resolve-Path -LiteralPath $ExpandDll).Path
foreach ($dependencyName in @(
    'ICSharpCode.SharpZipLib.dll',
    'NPOI.dll',
    'NPOI.OpenXmlFormats.dll',
    'NPOI.OpenXml4Net.dll',
    'NPOI.OOXML.dll'
)) {
    $dependencyPath = Join-Path $dllDir $dependencyName
    if (-not (Test-Path -LiteralPath $dependencyPath)) {
        throw "Missing test dependency: $dependencyPath"
    }
    [void][Reflection.Assembly]::LoadFrom($dependencyPath)
}

$assembly = [Reflection.Assembly]::LoadFrom((Resolve-Path -LiteralPath $ExpandDll).Path)
$type = $assembly.GetType('RecoNet.FormPanel', $true)
$flags = [Reflection.BindingFlags]'Public,NonPublic,Static,Instance'

function Get-StaticMethod([string]$Name, [int]$ParameterCount) {
    $methods = @($type.GetMethods($flags) | Where-Object { $_.Name -eq $Name -and $_.GetParameters().Count -eq $ParameterCount })
    if ($methods.Count -ne 1) {
        throw "Expected exactly one FormPanel.$Name with $ParameterCount parameters, found $($methods.Count)"
    }
    return $methods[0]
}

function Assert-Equal([string]$Label, $Expected, $Actual) {
    if ([string]$Expected -cne [string]$Actual) {
        throw "$Label`: expected '$Expected', actual '$Actual'"
    }
}

function Assert-True([string]$Label, [bool]$Condition) {
    if (-not $Condition) {
        throw "$Label`: condition was false"
    }
}

# 反射参数统一拆包：泛型集合、字符串等都传 BaseObject，避免 PowerShell 把集合展开成多个参数。
function New-ArgumentArray([object[]]$Values) {
    $boxed = [object[]]::new($Values.Count)
    for ($i = 0; $i -lt $Values.Count; $i++) {
        $value = $Values[$i]
        if ($null -ne $value) { $value = $value.PSObject.BaseObject }
        $boxed[$i] = $value
    }
    return ,$boxed
}

# ---------------------------------------------------------------------------
# 1. 批量绑定地址：行偏移只由所选行序号决定，失败行也占位。
#    起点 E4，选 3 行，第 2 行创建失败：第 3 行必须得到 E6（旧实现会得到 E5）。
# ---------------------------------------------------------------------------
$buildAddress = Get-StaticMethod 'BuildBatchBindAddress' 3
$startColumn = 5   # E
$startRow = 4
$rowResults = @($true, $false, $true)   # 第 2 行 TryCreateQuotaLink 失败
$assigned = New-Object 'System.Collections.Generic.List[string]'
$rowOrdinal = 0
foreach ($created in $rowResults) {
    $ordinal = $rowOrdinal++
    $address = [string]$buildAddress.Invoke($null, (New-ArgumentArray @([int]$startColumn, [int]$startRow, [int]$ordinal)))
    if (-not $created) { continue }
    [void]$assigned.Add($address)
}
Assert-Equal 'batch address count' 2 $assigned.Count
Assert-Equal 'batch address row1' 'E4' $assigned[0]
Assert-Equal 'batch address row3 (after a skipped row)' 'E6' $assigned[1]
Assert-Equal 'batch address AA10+3' 'AA13' ([string]$buildAddress.Invoke($null, (New-ArgumentArray @([int]27, [int]10, [int]3))))

# 源码守卫：入口循环必须在 TryCreateQuotaLink 之前按行序号算地址，且不再用成功计数 bound 做偏移。
$excelLinkSource = [IO.File]::ReadAllText((Join-Path $repoRoot 'tools\RecoExpandPanel\ExcelLinkFeature.cs'), [Text.Encoding]::UTF8)
$ordinalIndex = $excelLinkSource.IndexOf('string address = BuildBatchBindAddress(startRef.Column, startRef.Row, ordinal);')
$tryCreateIndex = $excelLinkSource.IndexOf('if (!TryCreateQuotaLink(mainForm, conn, row, out link, out error))', [Math]::Max(0, $ordinalIndex))
Assert-True 'batch bind computes the address from the row ordinal before TryCreateQuotaLink' ($ordinalIndex -ge 0 -and $tryCreateIndex -gt $ordinalIndex)
Assert-True 'batch bind no longer offsets by the success counter' (-not $excelLinkSource.Contains('(startRef.Row + bound)'))
Assert-True 'batch bind reports skipped rows in the result message' ($excelLinkSource.Contains('skippedRows.Add(') -and $excelLinkSource.Contains('String.Join(Environment.NewLine, skippedRows.ToArray())'))
Write-Host 'PASS batch-bind address: skipped row keeps its slot (E4, skip, E6); result lists skipped rows.' -ForegroundColor Green

# ---------------------------------------------------------------------------
# 2. TryBuildCellTokenFormula：单次替换，占位符 V{i} 不会被地址表里的 V 列地址二次替换。
# ---------------------------------------------------------------------------
$tokenFormula = Get-StaticMethod 'TryBuildCellTokenFormula' 3
function Invoke-TokenFormula([string]$Expression, [string[]]$Addresses) {
    $list = New-Object 'System.Collections.Generic.List[string]'
    foreach ($address in $Addresses) { [void]$list.Add($address) }
    $invokeArgs = New-ArgumentArray @($Expression, $list, $null)
    $ok = [bool]$tokenFormula.Invoke($null, $invokeArgs)
    return [pscustomobject]@{ Ok = $ok; Template = [string]$invokeArgs[2] }
}

$tokenCases = @(
    [pscustomobject]@{ Expression = '(X5+Y6)*V1'; Addresses = @('X5', 'Y6', 'V1'); Template = '(V0+V1)*V2' },
    [pscustomobject]@{ Expression = '(X5+Y6)*V1'; Addresses = @('V1', 'X5', 'Y6'); Template = '(V1+V2)*V0' },
    [pscustomobject]@{ Expression = 'A1+A1*2';    Addresses = @('A1');             Template = 'V0+V0*2' },
    [pscustomobject]@{ Expression = 'a1+A10';     Addresses = @('A1', 'A10');      Template = 'V0+V1' }
)
foreach ($case in $tokenCases) {
    $result = Invoke-TokenFormula $case.Expression $case.Addresses
    Assert-True "token formula ok for '$($case.Expression)'" $result.Ok
    Assert-Equal "token formula template for '$($case.Expression)'" $case.Template $result.Template
}
Write-Host "PASS TryBuildCellTokenFormula: $($tokenCases.Count) cases, placeholders are never re-replaced." -ForegroundColor Green

# ---------------------------------------------------------------------------
# 3. NormalizeExpressionMergedAnchors：地址替换带边界，B2→B1 不得把 B21 改成 B11。
#    合并区域 B1:B2（锚点 B1，B2 为非锚点格）。
# ---------------------------------------------------------------------------
$mergedRegionType = $type.GetNestedType('ExcelMergedRegion', $flags)
$mergedListType = [Collections.Generic.List``1].MakeGenericType($mergedRegionType)
$regions = [Activator]::CreateInstance($mergedListType)
$region = [Activator]::CreateInstance($mergedRegionType)
$mergedRegionType.GetField('FirstRow', $flags).SetValue($region, [int]1)
$mergedRegionType.GetField('LastRow', $flags).SetValue($region, [int]2)
$mergedRegionType.GetField('FirstColumn', $flags).SetValue($region, [int]2)
$mergedRegionType.GetField('LastColumn', $flags).SetValue($region, [int]2)
[void]$regions.Add($region)

$normalizeAnchors = @($type.GetMethods($flags) | Where-Object {
    $_.Name -eq 'NormalizeExpressionMergedAnchors' -and
    $_.GetParameters().Count -eq 2 -and
    $_.GetParameters()[1].ParameterType.IsGenericType
})
if ($normalizeAnchors.Count -ne 1) { throw "Expected one NormalizeExpressionMergedAnchors(string, List<ExcelMergedRegion>), found $($normalizeAnchors.Count)" }
$normalizeAnchors = $normalizeAnchors[0]

$anchorCases = @(
    [pscustomobject]@{ Expression = 'B2+B21';   Expected = 'B1+B21' },
    [pscustomobject]@{ Expression = 'B21+B2';   Expected = 'B21+B1' },
    [pscustomobject]@{ Expression = '(B2+B21)*2'; Expected = '(B1+B21)*2' },
    [pscustomobject]@{ Expression = 'B1+B2';    Expected = 'B1+B1' }
)
foreach ($case in $anchorCases) {
    $actual = [string]$normalizeAnchors.Invoke($null, (New-ArgumentArray @($case.Expression, $regions)))
    Assert-Equal "merged anchor normalize '$($case.Expression)'" $case.Expected $actual
}
Write-Host "PASS NormalizeExpressionMergedAnchors: $($anchorCases.Count) cases, B2->B1 leaves B21 untouched." -ForegroundColor Green

# ---------------------------------------------------------------------------
# 4. 绑定存储：损坏文件改名备份 + LoadFailed + 拒绝保存；正常保存原子写入且无临时残留。
# ---------------------------------------------------------------------------
$loadStore = Get-StaticMethod 'LoadStoreFromPath' 1
$saveStore = Get-StaticMethod 'SaveStoreToPath' 2
$storeType = $type.GetNestedType('ExcelLinkStore', $flags)
$linkType = $type.GetNestedType('ExcelQuotaLink', $flags)
$loadFailedProperty = $storeType.GetProperty('LoadFailed', $flags)
$backupPathProperty = $storeType.GetProperty('CorruptBackupPath', $flags)
$linksProperty = $storeType.GetProperty('Links', $flags)
if ($null -eq $loadFailedProperty -or $null -eq $backupPathProperty) { throw 'ExcelLinkStore is missing LoadFailed/CorruptBackupPath.' }
Assert-True 'LoadFailed is XmlIgnore' (@($loadFailedProperty.GetCustomAttributes([Xml.Serialization.XmlIgnoreAttribute], $false)).Count -eq 1)
Assert-True 'CorruptBackupPath is XmlIgnore' (@($backupPathProperty.GetCustomAttributes([Xml.Serialization.XmlIgnoreAttribute], $false)).Count -eq 1)

function Get-InnerException([Exception]$Exception, [Type]$Wanted) {
    $current = $Exception
    while ($null -ne $current) {
        if ($Wanted.IsInstanceOfType($current)) { return $current }
        $current = $current.InnerException
    }
    return $null
}

function New-StoreLink([long]$Sequence, [string]$Code, [string]$Address) {
    $link = [Activator]::CreateInstance($linkType)
    $linkType.GetProperty('QuotaSequence', $flags).SetValue($link, [long]$Sequence, $null)
    $linkType.GetProperty('QuotaCode', $flags).SetValue($link, [string]$Code, $null)
    $linkType.GetProperty('CellAddress', $flags).SetValue($link, [string]$Address, $null)
    $linkType.GetProperty('Expression', $flags).SetValue($link, [string]$Address, $null)
    return $link
}

function Get-TempResidue([string]$Directory) {
    return @(Get-ChildItem -LiteralPath $Directory -Force | Where-Object { $_.Name -like '*.tmp' })
}

$workDir = Join-Path ([IO.Path]::GetTempPath()) ('reco-excel-link-store-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($workDir)
try {
    $storePath = Join-Path $workDir 'store.xml'
    $truncatedXml = '<?xml version="1.0"?><ExcelLinkStore><Links><ExcelQuotaLink><QuotaSequence>1</Quota'
    [IO.File]::WriteAllText($storePath, $truncatedXml, [Text.Encoding]::UTF8)

    # 4a. 截断 XML → LoadFailed，损坏文件改名为 store.corrupt-yyyyMMdd-HHmmss.xml 备份，内容原样保留。
    $failedStore = $loadStore.Invoke($null, (New-ArgumentArray @($storePath)))
    Assert-True 'truncated store reports LoadFailed' ([bool]$loadFailedProperty.GetValue($failedStore, $null))
    Assert-Equal 'truncated store exposes an empty link list' 0 $linksProperty.GetValue($failedStore, $null).Count
    Assert-True 'corrupt file was moved away from the store path' (-not (Test-Path -LiteralPath $storePath))
    $backups = @(Get-ChildItem -LiteralPath $workDir -Force | Where-Object { $_.Name -match '^store\.corrupt-\d{8}-\d{6}(-\d+)?\.xml$' })
    Assert-Equal 'exactly one corrupt backup' 1 $backups.Count
    Assert-Equal 'CorruptBackupPath points at the backup' $backups[0].FullName ([string]$backupPathProperty.GetValue($failedStore, $null))
    Assert-Equal 'corrupt backup keeps the original bytes' $truncatedXml ([IO.File]::ReadAllText($backups[0].FullName, [Text.Encoding]::UTF8))

    # 4b. LoadFailed 的库拒绝保存：抛 InvalidOperationException，消息指向备份；备份仍在，不写新文件。
    $refused = $null
    try {
        [void]$saveStore.Invoke($null, (New-ArgumentArray @($storePath, $failedStore)))
    }
    catch {
        $refused = Get-InnerException $_.Exception ([InvalidOperationException])
    }
    if ($null -eq $refused) { throw 'SaveStoreToPath accepted a LoadFailed store.' }
    Assert-True 'refusal message names the corrupt store' ($refused.Message.Contains('绑定存储文件损坏'))
    Assert-True 'refusal message names the backup path' ($refused.Message.Contains($backups[0].FullName))
    Assert-True 'refused save wrote no store file' (-not (Test-Path -LiteralPath $storePath))
    Assert-True 'refused save kept the corrupt backup' (Test-Path -LiteralPath $backups[0].FullName)
    Assert-Equal 'refused save left no temp file' 0 @(Get-TempResidue $workDir).Count

    # 4c. 正常保存：首次写入 + 已有文件替换，均可重新加载且无临时残留。
    $store = [Activator]::CreateInstance($storeType)
    [void]$linksProperty.GetValue($store, $null).Add((New-StoreLink 42 'Q-1' 'E4'))
    [void]$saveStore.Invoke($null, (New-ArgumentArray @($storePath, $store)))
    Assert-True 'first save created the store file' (Test-Path -LiteralPath $storePath)
    Assert-Equal 'first save left no temp file' 0 @(Get-TempResidue $workDir).Count
    $reloaded = $loadStore.Invoke($null, (New-ArgumentArray @($storePath)))
    Assert-True 'reloaded store is healthy' (-not [bool]$loadFailedProperty.GetValue($reloaded, $null))
    Assert-Equal 'reloaded store link count' 1 $linksProperty.GetValue($reloaded, $null).Count
    Assert-Equal 'reloaded store link code' 'Q-1' ([string]$linkType.GetProperty('QuotaCode', $flags).GetValue($linksProperty.GetValue($reloaded, $null)[0], $null))

    [void]$linksProperty.GetValue($reloaded, $null).Add((New-StoreLink 43 'Q-2' 'E5'))
    [void]$saveStore.Invoke($null, (New-ArgumentArray @($storePath, $reloaded)))
    Assert-Equal 'replace save left no temp file' 0 @(Get-TempResidue $workDir).Count
    $reloadedAgain = $loadStore.Invoke($null, (New-ArgumentArray @($storePath)))
    Assert-Equal 'replace save persisted both links' 2 $linksProperty.GetValue($reloadedAgain, $null).Count
    Assert-True 'corrupt backup survives later saves' (Test-Path -LiteralPath $backups[0].FullName)
    Write-Host 'PASS store safety: corrupt file backed up + LoadFailed + save refused; atomic first/replace save with no temp residue.' -ForegroundColor Green
}
finally {
    if (Test-Path -LiteralPath $workDir) { Remove-Item -LiteralPath $workDir -Recurse -Force }
}

Write-Host 'PASS: Excel link data safety (batch address, token formula, merged anchors, store load/save).' -ForegroundColor Green
