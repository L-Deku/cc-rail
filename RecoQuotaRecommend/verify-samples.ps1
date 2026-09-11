# 检索规则回归验证（断言式）：反射调用 RecoQuotaRecommend.dll 跑 AGENTS.md「代码编辑规则」列出的代表样例。
# - 默认用仓库 bin\RecoQuotaRecommend.dll，可用 $env:RECO_QUOTA_DLL 覆盖。
# - 只读：DLL 与索引文件复制到临时目录后加载，不动部署目录和用户数据；不读取任何本地学习文件。
# - 临时目录里放一个空的 RejjNet2020.exe 占位，使 SearchIndexStore.ResolveDatabaseName() 在无宿主进程时判定为 RecoData2020，
#   并只复制 quota-index-2020.jsonl / material-index-2020.jsonl（正式 2024 软件目录下当前只有 2020 版带 source_database 标记的索引）。
# - 每个样例独立断言并逐一报告 PASS/FAIL/SKIP；全部通过打印 PASS，任一失败最后 throw（退出码非 0）。
# - 结构：外层进程准备临时目录并启动子进程执行断言（DLL 被 LoadFrom 锁定，只有子进程退出后外层才能删除临时目录）。
[CmdletBinding()]
param(
    [switch]$Inner
)

$ErrorActionPreference = "Stop"
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$root = Split-Path -Parent $PSScriptRoot
$softwareDir = Join-Path $root "2024铁路工程云计价系统网络版V1.0\铁路工程云计价系统网络版V1.0"
$indexFiles = @("quota-index-2020.jsonl", "material-index-2020.jsonl")

function Invoke-Assert([bool]$condition, [string]$message) {
    if (-not $condition) {
        throw ("ASSERT FAILED: " + $message)
    }
}

if (-not $Inner) {
    # ===== 外层：准备临时目录，启动子进程，最后清理 =====
    $sourceDll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_QUOTA_DLL)) {
        $env:RECO_QUOTA_DLL
    } else {
        Join-Path $PSScriptRoot "bin\RecoQuotaRecommend.dll"
    }
    if (-not (Test-Path -LiteralPath $sourceDll)) {
        throw "RecoQuotaRecommend.dll not found: $sourceDll"
    }
    foreach ($name in $indexFiles) {
        $src = Join-Path $softwareDir ("RecoQuotaData\" + $name)
        if (-not (Test-Path -LiteralPath $src)) {
            throw "Index file not found: $src"
        }
    }

    $work = Join-Path $env:TEMP ("reco-verify-" + [Guid]::NewGuid().ToString("N").Substring(0, 8))
    $dataDir = Join-Path $work "RecoQuotaData"
    [System.IO.Directory]::CreateDirectory($dataDir) | Out-Null
    try {
        Copy-Item -LiteralPath $sourceDll -Destination (Join-Path $work "RecoQuotaRecommend.dll") -Force
        # 空占位主程序：让无宿主进程时的数据库判定落到 RecoData2020。
        [System.IO.File]::WriteAllBytes((Join-Path $work "RejjNet2020.exe"), [byte[]]@())
        foreach ($name in $indexFiles) {
            Copy-Item -LiteralPath (Join-Path $softwareDir ("RecoQuotaData\" + $name)) -Destination (Join-Path $dataDir $name) -Force
        }

        Write-Host ("DLL:  " + (Resolve-Path -LiteralPath $sourceDll).Path)
        Write-Host ("Work: " + $work)

        $hostExe = (Get-Process -Id $PID).Path
        $env:RECO_VERIFY_WORK = $work
        try {
            & $hostExe -NoProfile -NoLogo -ExecutionPolicy Bypass -File $PSCommandPath -Inner
            $exitCode = $LASTEXITCODE
        } finally {
            Remove-Item Env:RECO_VERIFY_WORK -ErrorAction SilentlyContinue
        }
        if ($exitCode -ne 0) {
            throw ("verify-samples FAILED (exit code " + $exitCode + ")")
        }
        Write-Host "PASS verify-samples 全部样例通过"
    } finally {
        if (Test-Path -LiteralPath $work) {
            Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue
            if (Test-Path -LiteralPath $work) {
                Write-Warning ("临时目录未能完全清理: " + $work)
            }
        }
    }
    return
}

# ===== 子进程：加载 DLL 与索引并断言 =====
$failures = New-Object System.Collections.Generic.List[string]
try {
    $work = $env:RECO_VERIFY_WORK
    if ([String]::IsNullOrWhiteSpace($work) -or -not (Test-Path -LiteralPath $work)) {
        throw "RECO_VERIFY_WORK is not set or does not exist: $work"
    }
    $dataDir = Join-Path $work "RecoQuotaData"
    $testDll = Join-Path $work "RecoQuotaRecommend.dll"

    $assembly = [System.Reflection.Assembly]::LoadFrom($testDll)
    $flags = [System.Reflection.BindingFlags]"Static,Instance,Public,NonPublic"
    $storeType = $assembly.GetType("RecoQuotaRecommend.SearchIndexStore", $true)
    $itemType = $assembly.GetType("RecoQuotaRecommend.ExcelQuantityItem", $true)
    $rowType = $assembly.GetType("RecoQuotaRecommend.RecommendationRow", $true)

    # --- 数据库判定与索引加载核对（前置条件，失败直接中止） ---
    $store = $storeType.GetMethod("LoadOrBuild", $flags).Invoke($null, @())
    $databaseName = [string]$storeType.GetField("materialDatabaseName", $flags).GetValue($store)
    $quotaIndexPath = [string]$storeType.GetField("quotaIndexPath", $flags).GetValue($store)
    $quotaCount = [int]$storeType.GetProperty("QuotaCount", $flags).GetValue($store, $null)
    $materialCount = [int]$storeType.GetProperty("MaterialCount", $flags).GetValue($store, $null)
    Write-Host ("Store: database=" + $databaseName + " quotaIndex=" + [IO.Path]::GetFileName($quotaIndexPath) + " quotas=" + $quotaCount + " materials=" + $materialCount)
    Invoke-Assert ($databaseName -eq "RecoData2020") ("无宿主进程时应按占位 RejjNet2020.exe 判定为 RecoData2020，实际 " + $databaseName)
    Invoke-Assert ([IO.Path]::GetFileName($quotaIndexPath) -eq "quota-index-2020.jsonl") ("应加载 quota-index-2020.jsonl，实际 " + $quotaIndexPath)
    Invoke-Assert ($quotaCount -gt 0) "定额索引为空"
    Invoke-Assert ($materialCount -gt 0) "材料索引为空"

    # --- 从当前索引文件建立 定额编号 -> 书号集合 查找表，用于核对返回行书号与当前索引库一致 ---
    $bookCodesByQuota = New-Object "System.Collections.Generic.Dictionary[string,System.Collections.Generic.HashSet[string]]" ([StringComparer]::OrdinalIgnoreCase)
    $codeRegex = New-Object System.Text.RegularExpressions.Regex('"quota_code":"([^"]*)"')
    $bookRegex = New-Object System.Text.RegularExpressions.Regex('"book_code":"([^"]*)"')
    $dbRegex = New-Object System.Text.RegularExpressions.Regex('"source_database":"([^"]*)"')
    $indexSourceDatabase = ""
    foreach ($line in [System.IO.File]::ReadLines((Join-Path $dataDir "quota-index-2020.jsonl"), [System.Text.Encoding]::UTF8)) {
        if ([String]::IsNullOrWhiteSpace($line)) { continue }
        $codeMatch = $codeRegex.Match($line)
        $bookMatch = $bookRegex.Match($line)
        if (-not $codeMatch.Success -or -not $bookMatch.Success) { continue }
        if ($indexSourceDatabase -eq "") {
            $dbMatch = $dbRegex.Match($line)
            if ($dbMatch.Success) { $indexSourceDatabase = $dbMatch.Groups[1].Value }
        }
        $code = $codeMatch.Groups[1].Value.Trim()
        $set = $null
        if (-not $bookCodesByQuota.TryGetValue($code, [ref]$set)) {
            $set = New-Object "System.Collections.Generic.HashSet[string]" ([StringComparer]::OrdinalIgnoreCase)
            $bookCodesByQuota[$code] = $set
        }
        [void]$set.Add($bookMatch.Groups[1].Value.Trim())
    }
    Invoke-Assert ($indexSourceDatabase -eq "RecoData2020") ("quota-index-2020.jsonl 的 source_database 应为 RecoData2020，实际 " + $indexSourceDatabase)
    Write-Host ("Index: source_database=" + $indexSourceDatabase + " distinctQuotaCodes=" + $bookCodesByQuota.Count)

    # 与 QuotaInlineSearchFeature.BuildItem 一致地构造查询项。
    function New-QuantityItem([string]$name, [string]$unit, [string]$value) {
        $item = [Activator]::CreateInstance($itemType)
        $itemType.GetField("Name", $flags).SetValue($item, $name)
        $itemType.GetField("OriginalName", $flags).SetValue($item, $name)
        $itemType.GetField("Unit", $flags).SetValue($item, $unit)
        $itemType.GetField("ValueText", $flags).SetValue($item, $value)
        $itemType.GetField("ContextText", $flags).SetValue($item, ($name + " " + $unit + " " + $value))
        $itemType.GetField("RawRowText", $flags).SetValue($item, ($name + " " + $unit + " " + $value))
        $itemType.GetField("SkipAiNameNormalization", $flags).SetValue($item, $true)
        return $item.PSObject.BaseObject
    }

    function Get-RowField($row, [string]$name) {
        return $rowType.GetField($name, $flags).GetValue($row)
    }

    function Show-Rows([string]$label, $rows, [int]$max = 8) {
        Write-Host ("--- " + $label + "  返回 " + $rows.Count + " 条")
        $shown = 0
        foreach ($r in $rows) {
            if ($shown -ge $max) { Write-Host ("    ... 其余 " + ($rows.Count - $shown) + " 条略"); break }
            Write-Host ("    [" + (Get-RowField $r "Source") + "/" + (Get-RowField $r "TargetKind") + "] " + (Get-RowField $r "QuotaCode") + " " + (Get-RowField $r "QuotaName") + " 书号=" + (Get-RowField $r "BookCode") + " score=" + (Get-RowField $r "Score"))
            $shown++
        }
    }

    # 每条返回行的书号必须与当前索引库（2020 版）一致：等于索引文件里该定额编号的 book_code，且不是 2024 版书号。
    function Assert-BookCodes([string]$label, $rows) {
        foreach ($r in $rows) {
            if ([string](Get-RowField $r "TargetKind") -ne "quota") { continue }
            $code = [string](Get-RowField $r "QuotaCode")
            $book = [string](Get-RowField $r "BookCode")
            Invoke-Assert (-not [String]::IsNullOrWhiteSpace($book)) ($label + ": " + $code + " 书号为空")
            Invoke-Assert ($book.IndexOf("2024") -lt 0) ($label + ": " + $code + " 书号 " + $book + " 不是 2020 版索引库书号")
            $set = $null
            Invoke-Assert ($bookCodesByQuota.TryGetValue($code, [ref]$set)) ($label + ": " + $code + " 不在 quota-index-2020.jsonl 中")
            Invoke-Assert ($set.Contains($book)) ($label + ": " + $code + " 书号 " + $book + " 与索引文件不一致（索引: " + ($set -join ",") + "）")
        }
    }

    function Get-Codes($rows) {
        [string[]]$codes = @(foreach ($r in $rows) { [string](Get-RowField $r "QuotaCode") })
        return ,$codes
    }

    # 单个样例独立执行：断言失败记入 $failures，继续跑后续样例，最后统一判定。
    function Invoke-Sample([string]$title, [scriptblock]$body) {
        Write-Host ("===== " + $title + " =====")
        try {
            & $body
        } catch {
            $message = $_.Exception.Message
            Write-Host ("FAIL " + $title + ": " + $message)
            if ($_.InvocationInfo -ne $null -and $message -notlike "ASSERT FAILED:*") {
                Write-Host ("  at " + $_.InvocationInfo.PositionMessage)
            }
            $failures.Add($title + ": " + $message)
        }
    }

    # 普通检索的真实入口：QuotaInlineSearchFeature.BuildSearchResult -> SearchIndexStore.SearchAllQuotaCandidates(item, "全部", scope)
    # （定额模式不限量，scope 由宿主当前条目决定，离线为 null）。
    $searchAll = $storeType.GetMethod("SearchAllQuotaCandidates", $flags)
    $searchLimited = $storeType.GetMethod("SearchQuotaCandidates", $flags)
    $allCategories = "全部"

    Invoke-Sample "样例 1：警示带 应命中 PY-738 警示（示踪）带铺设" {
        $item = New-QuantityItem "警示带" "m" "100"
        [object[]]$arguments = @($item, $allCategories, $null)
        [object]$rows = $searchAll.Invoke($store, $arguments)
        Show-Rows "警示带" $rows
        Assert-BookCodes "警示带" $rows
        $codes = Get-Codes $rows
        Invoke-Assert ($codes -contains "PY-738") ("警示带 未命中 PY-738（返回 " + $rows.Count + " 条: " + ($codes -join ",") + "）")
        $py738Set = $null
        [void]$bookCodesByQuota.TryGetValue("PY-738", [ref]$py738Set)
        Invoke-Assert ($py738Set -ne $null -and $py738Set.Contains("PY_2017")) "PY-738 在 2020 索引中的书号应为 PY_2017"
        $py738 = $null
        foreach ($r in $rows) { if ([string](Get-RowField $r "QuotaCode") -eq "PY-738") { $py738 = $r; break } }
        Invoke-Assert (([string](Get-RowField $py738 "QuotaName")).Contains("警示")) "PY-738 定额名应含「警示」"
        Write-Host "PASS 警示带 -> PY-738"
    }

    Invoke-Sample "样例 2：警示桩 不应误命中 ZY-41 现浇检查坑、井 混凝土" {
        $item = New-QuantityItem "警示桩" "根" "10"
        [object[]]$arguments = @($item, $allCategories, $null)
        [object]$rows = $searchAll.Invoke($store, $arguments)
        Show-Rows "警示桩" $rows
        Assert-BookCodes "警示桩" $rows
        $codes = Get-Codes $rows
        Invoke-Assert (-not ($codes -contains "ZY-41")) "警示桩 误命中 ZY-41"
        Write-Host "PASS 警示桩 不含 ZY-41"
    }

    Invoke-Sample "样例 3：HPB300钢筋 不应返回材料" {
        $item = New-QuantityItem "HPB300钢筋" "t" "5"
        [object[]]$arguments = @($item, $allCategories, $null)
        [object]$rows = $searchAll.Invoke($store, $arguments)
        Show-Rows "HPB300钢筋" $rows
        Assert-BookCodes "HPB300钢筋" $rows
        foreach ($r in $rows) {
            $kind = [string](Get-RowField $r "TargetKind")
            $src = [string](Get-RowField $r "Source")
            Invoke-Assert ($kind -eq "quota" -and $src -eq "index") ("HPB300钢筋 返回了非定额行: " + $src + "/" + $kind + " " + (Get-RowField $r "QuotaCode"))
        }
        Write-Host ("PASS HPB300钢筋 无材料行（定额 " + $rows.Count + " 条）")
    }

    Invoke-Sample "样例 4：HRB400钢筋 普通检索 1 条最优定额" {
        # 当前代码里普通检索入口 SearchAllQuotaCandidates 不限量（弹窗展示全部候选），没有 limit=1 的调用方；
        # 因此按“最优定额 = 排序第 1 条”断言：top-1 必须是 HRB400/螺纹钢筋类定额，且 SearchQuotaCandidates(limit=1) 恰返回该 1 条。
        $item = New-QuantityItem "HRB400钢筋" "t" "5"
        [object[]]$arguments = @($item, $allCategories, $null)
        [object]$rows = $searchAll.Invoke($store, $arguments)
        Show-Rows "HRB400钢筋" $rows
        Assert-BookCodes "HRB400钢筋" $rows
        Invoke-Assert ($rows.Count -ge 1) "HRB400钢筋 未返回任何定额"
        $topCode = [string](Get-RowField $rows[0] "QuotaCode")
        $topName = [string](Get-RowField $rows[0] "QuotaName")
        Invoke-Assert ($topName -match "HRB400|螺纹") ("HRB400钢筋 top-1 不是 HRB400/螺纹钢筋定额: " + $topCode + " " + $topName)
        [object[]]$limitedArguments = @($item, $allCategories, $null, 1)
        [object]$best = $searchLimited.Invoke($store, $limitedArguments)
        Invoke-Assert ($best.Count -eq 1) ("SearchQuotaCandidates(limit=1) 应恰返回 1 条，实际 " + $best.Count)
        Invoke-Assert ([string](Get-RowField $best[0] "QuotaCode") -eq $topCode) "limit=1 结果与不限量排序第 1 条不一致"
        Write-Host ("PASS HRB400钢筋 最优定额 " + $topCode + " " + $topName)
    }

    Write-Host "===== 样例 5：土方外运 组件框多条 ====="
    # 当前组件框推荐只从 RecoLearning SQL 学习库（SignatureBoxMap/QuotaBoxTarget）读取，入口在 RecoExpandPanel 的 SmartFillFeature，
    # 需要 SQL 与宿主当前项目上下文；RecoQuotaRecommend.MappingStore 读的是本地 mapping-boxes.jsonl，数据存储规则禁止用它回退。
    Write-Host "SKIP 土方外运 组件框：需要 RecoLearning SQL 与宿主上下文，离线脚本不覆盖（请在实机按「推荐定额」窗口验证）"

    if ($failures.Count -gt 0) {
        Write-Host ("===== 失败 " + $failures.Count + " 项 =====")
        foreach ($f in $failures) { Write-Host ("  - " + $f) }
        exit 1
    }
    Write-Host "PASS 全部样例断言通过"
    exit 0
} catch {
    Write-Host ("FAIL " + $_.Exception.Message)
    if ($_.InvocationInfo -ne $null) { Write-Host ("  at " + $_.InvocationInfo.PositionMessage) }
    exit 1
}
