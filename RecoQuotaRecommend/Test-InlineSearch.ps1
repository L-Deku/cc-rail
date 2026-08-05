$ErrorActionPreference = "Stop"

$root = Split-Path -Parent $PSScriptRoot
$softwareDir = Join-Path $root "铁路基本建设工程投资控制系统2020网络版V0503021201"
$sourceDll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_QUOTA_DLL)) {
    $env:RECO_QUOTA_DLL
} else {
    Join-Path $PSScriptRoot "bin\RecoQuotaRecommend.dll"
}

if (-not (Test-Path -LiteralPath $sourceDll)) {
    throw "RecoQuotaRecommend.dll not found: $sourceDll"
}

$testDll = (Resolve-Path -LiteralPath $sourceDll).Path
$dataDir = Join-Path (Split-Path -Parent $testDll) "RecoQuotaData"
$createdDataDir = -not (Test-Path -LiteralPath $dataDir)
if ($createdDataDir) {
    [System.IO.Directory]::CreateDirectory($dataDir) | Out-Null
}
$copiedFiles = New-Object System.Collections.Generic.List[string]

try {
    foreach ($name in @("quota-index.jsonl", "material-index.jsonl")) {
        $destination = Join-Path $dataDir $name
        if (-not (Test-Path -LiteralPath $destination)) {
            Copy-Item -LiteralPath (Join-Path $softwareDir "RecoQuotaData\$name") -Destination $destination
            [void]$copiedFiles.Add($destination)
        }
    }

    $assembly = [Reflection.Assembly]::LoadFrom($testDll)
    $flags = [Reflection.BindingFlags]"Static,Instance,Public,NonPublic"
    $featureType = $assembly.GetType("RecoQuotaRecommend.QuotaInlineSearchFeature", $true)
    $parseMethod = $featureType.GetMethod("TryParseInlineQuery", $flags)

    function Assert-Parse([string]$text, [bool]$expectedOk, [string]$expectedQuery, [bool]$expectedMaterial) {
        [object[]]$arguments = @($text, "", $false)
        $actualOk = [bool]$parseMethod.Invoke($null, $arguments)
        if ($actualOk -ne $expectedOk -or [string]$arguments[1] -ne $expectedQuery -or [bool]$arguments[2] -ne $expectedMaterial) {
            throw "Parse mismatch: text=$text ok=$actualOk query=$($arguments[1]) material=$($arguments[2])"
        }
    }

    Assert-Parse "水泥" $true "水泥" $false
    Assert-Parse "水泥/" $true "水泥" $true
    Assert-Parse "水泥/ " $true "水泥" $true
    Assert-Parse "/" $false "" $true
    Assert-Parse "水泥/管" $true "水泥/管" $false
    Write-Host "PASS 定额/材料模式解析"

    $storeType = $assembly.GetType("RecoQuotaRecommend.SearchIndexStore", $true)
    $resolveDatabaseMethod = $storeType.GetMethod("ResolveDatabaseNameForHost", $flags)
    function Assert-Database([string]$processIdentity, [bool]$has2020, [bool]$has2024, [string]$expected) {
        [object[]]$arguments = @("D:\AI文件\铁路工程云计价系统网络版V1.0", $processIdentity, $has2020, $has2024)
        $actual = [string]$resolveDatabaseMethod.Invoke($null, $arguments)
        if ($actual -ne $expected) {
            throw "Database mismatch: process=$processIdentity actual=$actual expected=$expected"
        }
    }
    Assert-Database "RejjNet2020.exe" $true $true "RecoData2020"
    Assert-Database "ReJJGSNet2024.exe" $true $true "RecoData2024"
    Assert-Database "ReJJQDNet2024.exe" $true $true "RecoData2024"

    $resolveMaterialPathMethod = $storeType.GetMethod("ResolveMaterialIndexPath", $flags)
    [object[]]$path2020Arguments = @("D:\test\RecoQuotaData", "RecoData2020")
    [object[]]$path2024Arguments = @("D:\test\RecoQuotaData", "RecoData2024")
    $path2020 = [string]$resolveMaterialPathMethod.Invoke($null, $path2020Arguments)
    $path2024 = [string]$resolveMaterialPathMethod.Invoke($null, $path2024Arguments)
    if ([IO.Path]::GetFileName($path2020) -ne "material-index-2020.jsonl" -or
        [IO.Path]::GetFileName($path2024) -ne "material-index-2024.jsonl") {
        throw "Versioned material cache paths are invalid: 2020=$path2020 2024=$path2024"
    }

    $resolveQuotaPathMethod = $storeType.GetMethod("ResolveQuotaIndexPath", $flags)
    $quotaPath2020 = [string]$resolveQuotaPathMethod.Invoke($null, $path2020Arguments)
    $quotaPath2024 = [string]$resolveQuotaPathMethod.Invoke($null, $path2024Arguments)
    if ([IO.Path]::GetFileName($quotaPath2020) -ne "quota-index-2020.jsonl" -or
        [IO.Path]::GetFileName($quotaPath2024) -ne "quota-index-2024.jsonl") {
        throw "Versioned quota cache paths are invalid: 2020=$quotaPath2020 2024=$quotaPath2024"
    }
    Write-Host "PASS 混合目录按实际进程选库且定额/材料缓存分版本"

    $store = $storeType.GetMethod("LoadOrBuild", $flags).Invoke($null, @())
    $itemType = $assembly.GetType("RecoQuotaRecommend.ExcelQuantityItem", $true)
    $rowType = $assembly.GetType("RecoQuotaRecommend.RecommendationRow", $true)

    function New-QuantityItem([string]$name) {
        $item = [Activator]::CreateInstance($itemType)
        $itemType.GetField("Name", $flags).SetValue($item, $name)
        $itemType.GetField("OriginalName", $flags).SetValue($item, $name)
        $itemType.GetField("Unit", $flags).SetValue($item, "")
        $itemType.GetField("ValueText", $flags).SetValue($item, "")
        $itemType.GetField("ContextText", $flags).SetValue($item, $name)
        $itemType.GetField("RawRowText", $flags).SetValue($item, $name)
        $itemType.GetField("SkipAiNameNormalization", $flags).SetValue($item, $true)
        return $item.PSObject.BaseObject
    }

    $limitedMethod = $storeType.GetMethod("SearchQuotaCandidates", $flags)
    $allMethod = $storeType.GetMethod("SearchAllQuotaCandidates", $flags)
    $quotaItem = New-QuantityItem "水"
    [object[]]$limitedArguments = @($quotaItem, "全部", $null, 100)
    [object[]]$allArguments = @($quotaItem, "全部", $null)
    [object]$limited = $limitedMethod.Invoke($store, $limitedArguments)
    [object]$all = $allMethod.Invoke($store, $allArguments)
    if ($limited.Count -ne 100 -or $all.Count -le 100) {
        throw "Unlimited quota search failed: limited=$($limited.Count) all=$($all.Count)"
    }
    for ($index = 0; $index -lt $limited.Count; $index++) {
        $limitedCode = [string]$rowType.GetField("QuotaCode", $flags).GetValue($limited[$index])
        $allCode = [string]$rowType.GetField("QuotaCode", $flags).GetValue($all[$index])
        if ($limitedCode -ne $allCode) {
            throw "Quota order changed at index ${index}: limited=$limitedCode all=$allCode"
        }
    }
    Write-Host ("PASS 定额不限量且前100条排序不变: " + $all.Count)

    $materialMethod = $storeType.GetMethod("SearchMaterialCandidates", $flags)
    $materialItem = New-QuantityItem "电缆"
    [object[]]$materialArguments = @($materialItem)
    [object]$materials = $materialMethod.Invoke($store, $materialArguments)
    if ($materials.Count -le 100) {
        throw "Unlimited material search returned too few rows: $($materials.Count)"
    }
    foreach ($row in $materials) {
        $name = [string]$rowType.GetField("QuotaName", $flags).GetValue($row)
        $kind = [string]$rowType.GetField("TargetKind", $flags).GetValue($row)
        if ($name.Replace(" ", "").IndexOf("电缆", [StringComparison]::OrdinalIgnoreCase) -lt 0 -or $kind -ne "material") {
            throw "Invalid material candidate: kind=$kind name=$name"
        }
    }

    $deviceMaterialItem = New-QuantityItem "设备线夹"
    [object[]]$deviceMaterialArguments = @($deviceMaterialItem)
    [object]$deviceMaterials = $materialMethod.Invoke($store, $deviceMaterialArguments)
    if ($deviceMaterials.Count -eq 0) {
        throw "Material names containing '设备' were incorrectly filtered."
    }
    Write-Host ("PASS 材料名称连续匹配且不限量: " + $materials.Count)

    $tableMethod = $featureType.GetMethod("BuildCandidateTable", $flags)
    [object[]]$tableArguments = New-Object object[] 1
    $tableArguments[0] = $all.PSObject.BaseObject
    $tableTimer = [Diagnostics.Stopwatch]::StartNew()
    $table = $tableMethod.Invoke($null, $tableArguments)
    $tableTimer.Stop()
    try {
        if ($table.Rows.Count -ne $all.Count -or [int]$table.Rows[$table.Rows.Count - 1]["CandidateIndex"] -ne $all.Count - 1) {
            throw "Bound candidate table lost rows: table=$($table.Rows.Count) all=$($all.Count)"
        }

        $grid = New-Object System.Windows.Forms.DataGridView
        try {
            $grid.AutoGenerateColumns = $false
            $grid.AllowUserToAddRows = $false
            foreach ($name in @("QuotaCode", "QuotaName", "QuotaUnit", "BasePrice", "WorkContent")) {
                $column = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
                $column.Name = $name
                $column.DataPropertyName = $name
                [void]$grid.Columns.Add($column)
            }
            $grid.BindingContext = New-Object System.Windows.Forms.BindingContext
            $bindTimer = [Diagnostics.Stopwatch]::StartNew()
            $grid.DataSource = $table
            $bindTimer.Stop()
            if ($grid.Rows.Count -ne $all.Count) {
                throw "DataGridView binding lost rows: grid=$($grid.Rows.Count) all=$($all.Count)"
            }
        } finally {
            $grid.Dispose()
        }
    } finally {
        $table.Dispose()
    }
    Write-Host ("PASS DataTable绑定保留全部候选: buildMs=" + $tableTimer.ElapsedMilliseconds + " bindMs=" + $bindTimer.ElapsedMilliseconds)
} finally {
    foreach ($path in $copiedFiles) {
        if (Test-Path -LiteralPath $path) {
            Remove-Item -LiteralPath $path -Force
        }
    }
    if ($createdDataDir -and (Test-Path -LiteralPath $dataDir)) {
        Remove-Item -LiteralPath $dataDir -Force
    }
}
