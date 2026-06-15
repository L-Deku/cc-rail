param(
    [switch]$WhatIfOnly
)

$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$sourcePath = Join-Path $PSScriptRoot 'Migrate2020EstimateTo2024.cs'
$reportDir = Join-Path $PSScriptRoot 'reports'
$decisionPath = Join-Path $reportDir 'applied-manual-decisions.csv'
$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$beforePath = Join-Path $reportDir "cleanup-before-$stamp.csv"
$resourcePath = Join-Path $reportDir "cleanup-resource-candidates-$stamp.csv"
$summaryPath = Join-Path $reportDir "cleanup-summary-$stamp.txt"

function Read-Const([string]$name) {
    $source = Get-Content -LiteralPath $sourcePath -Raw -Encoding UTF8
    $pattern = 'private const string ' + [regex]::Escape($name) + '\s*=\s*"([^"]*)"'
    $match = [regex]::Match($source, $pattern)
    if (-not $match.Success) {
        throw "Cannot find const $name in $sourcePath"
    }
    $match.Groups[1].Value
}

function New-Command($connection, $transaction, [string]$sql) {
    $cmd = $connection.CreateCommand()
    $cmd.CommandTimeout = 300
    $cmd.CommandText = $sql
    if ($transaction) {
        $cmd.Transaction = $transaction
    }
    $cmd
}

function Add-Param($cmd, [string]$name, $value) {
    $param = $cmd.Parameters.Add($name, [System.Data.SqlDbType]::NVarChar)
    $param.Value = $value
}

function Add-IntParam($cmd, [string]$name, [int]$value) {
    $param = $cmd.Parameters.Add($name, [System.Data.SqlDbType]::Int)
    $param.Value = $value
}

function Invoke-Scalar($connection, $transaction, [string]$sql) {
    $cmd = New-Command $connection $transaction $sql
    try {
        $cmd.ExecuteScalar()
    }
    finally {
        $cmd.Dispose()
    }
}

function Invoke-NonQuery($connection, $transaction, [string]$sql) {
    $cmd = New-Command $connection $transaction $sql
    try {
        $cmd.ExecuteNonQuery()
    }
    finally {
        $cmd.Dispose()
    }
}

function Quote-Sql([string]$value) {
    "N'" + $value.Replace("'", "''") + "'"
}

function Book-InSql($books) {
    ($books | ForEach-Object { Quote-Sql $_ }) -join ','
}

$server = Read-Const 'Server'
$targetDb = Read-Const 'TargetDb'
$sqlUser = Read-Const 'SqlUser'
$sqlPassword = Read-Const 'SqlPassword'

$connectionString = "Data Source=$server,1433;Initial Catalog=$targetDb;User ID=$sqlUser;Password=$sqlPassword;TrustServerCertificate=True"

$books = @(
    'LG_2018','QG_2018','SG_2018','GG_2018','TG_2018','XG_2018','EG_2018','DG_2018','HG_2018','FG_2018','PG_2018','JG_2018','ZG_2018',
    'TZ_2020','XZ_2020','EZ_2020','DZ_2020','HZ_2020','FZ_2020','PZ_2020','JZ_2020'
)
$bookSql = Book-InSql $books

if (-not (Test-Path -LiteralPath $decisionPath)) {
    throw "Missing manual decision file: $decisionPath"
}

$resourceCandidates = New-Object System.Collections.Generic.List[object]
Import-Csv -LiteralPath $decisionPath -Encoding UTF8 | ForEach-Object {
    if ($_.manual_rule -eq '0=使用C列补充电算代号') {
        $resourceCandidates.Add([pscustomobject]@{
            kind = $_.kind
            code = [int]$_.new_code
            name = $_.target_name
            source = 'manual-L0'
        })
    }
}

$resourceCandidates.Add([pscustomobject]@{ kind = 'Material'; code = 1990051; name = '钢件防腐处理'; source = 'obsolete-alias' })
$resourceCandidates.Add([pscustomobject]@{ kind = 'Material'; code = 5833051; name = '钢芯铝绞线'; source = 'obsolete-alias' })

$resourceCandidates = $resourceCandidates |
    Sort-Object kind, code, name -Unique

$cn = New-Object System.Data.SqlClient.SqlConnection($connectionString)
$cn.Open()

try {
    $beforeRows = New-Object System.Collections.Generic.List[object]
    foreach ($table in @('定额库索引','定额节索引','定额库','定额库消耗')) {
        $count = Invoke-Scalar $cn $null "select count(*) from $table where 书号 in ($bookSql)"
        $beforeRows.Add([pscustomobject]@{ table = $table; scope = 'migrated-books'; count = [int]$count })
    }
    foreach ($book in @('LY_2024','DY_2024','QY_2024','YY_2024','HT_2024')) {
        $count = Invoke-Scalar $cn $null "select count(*) from 定额库 where 书号=N'$book'"
        $beforeRows.Add([pscustomobject]@{ table = '定额库'; scope = $book; count = [int]$count })
    }
    $beforeRows | Export-Csv -LiteralPath $beforePath -Encoding UTF8 -NoTypeInformation

    $resourceRows = New-Object System.Collections.Generic.List[object]
    foreach ($resource in $resourceCandidates) {
        if ($resource.kind -eq 'Machine') {
            $existsSql = 'select count(*) from 台班定额库 where 电算代号=@code and 机械台班名称=@name'
        }
        else {
            $existsSql = 'select count(*) from 材料单价库 where 电算代号=@code and 材料名称=@name'
        }
        $cmd = New-Command $cn $null $existsSql
        Add-IntParam $cmd '@code' $resource.code
        Add-Param $cmd '@name' $resource.name
        try {
            $exists = [int]$cmd.ExecuteScalar()
        }
        finally {
            $cmd.Dispose()
        }

        $cmd = New-Command $cn $null "select count(*) from 定额库消耗 where 电算代号=@code"
        Add-IntParam $cmd '@code' $resource.code
        try {
            $allRefs = [int]$cmd.ExecuteScalar()
        }
        finally {
            $cmd.Dispose()
        }

        $cmd = New-Command $cn $null "select count(*) from 定额库消耗 where 电算代号=@code and 书号 not in ($bookSql)"
        Add-IntParam $cmd '@code' $resource.code
        try {
            $otherRefs = [int]$cmd.ExecuteScalar()
        }
        finally {
            $cmd.Dispose()
        }

        $resourceRows.Add([pscustomobject]@{
            kind = $resource.kind
            code = $resource.code
            name = $resource.name
            source = $resource.source
            exists_before = $exists
            consume_refs_before = $allRefs
            non_migrated_refs_before = $otherRefs
        })
    }
    $resourceRows | Export-Csv -LiteralPath $resourcePath -Encoding UTF8 -NoTypeInformation

    if ($WhatIfOnly) {
        "WHATIF only. No database changes were made.`r`nBefore: $beforePath`r`nResources: $resourcePath" |
            Set-Content -LiteralPath $summaryPath -Encoding UTF8
        Write-Output "WHATIF complete. Reports: $beforePath ; $resourcePath"
        return
    }

    $tx = $cn.BeginTransaction()
    try {
        $deletedConsume = Invoke-NonQuery $cn $tx "delete from 定额库消耗 where 书号 in ($bookSql)"
        $deletedQuota = Invoke-NonQuery $cn $tx "delete from 定额库 where 书号 in ($bookSql)"
        $deletedSections = Invoke-NonQuery $cn $tx "delete from 定额节索引 where 书号 in ($bookSql)"
        $deletedIndex = Invoke-NonQuery $cn $tx "delete from 定额库索引 where 书号 in ($bookSql)"

        $deletedMaterials = 0
        $deletedMachines = 0
        foreach ($resource in $resourceCandidates) {
            if ($resource.kind -eq 'Machine') {
                $cmd = New-Command $cn $tx 'delete from 台班定额库 where 电算代号=@code and 机械台班名称=@name and not exists(select 1 from 定额库消耗 where 电算代号=@code)'
                Add-IntParam $cmd '@code' $resource.code
                Add-Param $cmd '@name' $resource.name
                try {
                    $deletedMachines += $cmd.ExecuteNonQuery()
                }
                finally {
                    $cmd.Dispose()
                }
            }
            else {
                $cmd = New-Command $cn $tx 'delete from 材料单价库 where 电算代号=@code and 材料名称=@name and not exists(select 1 from 定额库消耗 where 电算代号=@code)'
                Add-IntParam $cmd '@code' $resource.code
                Add-Param $cmd '@name' $resource.name
                try {
                    $deletedMaterials += $cmd.ExecuteNonQuery()
                }
                finally {
                    $cmd.Dispose()
                }
            }
        }

        $remainingMigrated = 0
        foreach ($table in @('定额库索引','定额节索引','定额库','定额库消耗')) {
            $remainingMigrated += [int](Invoke-Scalar $cn $tx "select count(*) from $table where 书号 in ($bookSql)")
        }
        if ($remainingMigrated -ne 0) {
            throw "Cleanup verification failed: migrated rows remain: $remainingMigrated"
        }

        foreach ($row in $beforeRows | Where-Object { $_.scope -like '*_2024' }) {
            $afterCount = [int](Invoke-Scalar $cn $tx "select count(*) from 定额库 where 书号=N'$($row.scope)'")
            if ($afterCount -ne $row.count) {
                throw "Budget book count changed for $($row.scope): before=$($row.count), after=$afterCount"
            }
        }

        $tx.Commit()
    }
    catch {
        $tx.Rollback()
        throw
    }

    $afterRows = New-Object System.Collections.Generic.List[object]
    foreach ($table in @('定额库索引','定额节索引','定额库','定额库消耗')) {
        $count = Invoke-Scalar $cn $null "select count(*) from $table where 书号 in ($bookSql)"
        $afterRows.Add([pscustomobject]@{ table = $table; scope = 'migrated-books'; count = [int]$count })
    }

    $summary = @()
    $summary += "Cleanup completed at $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $summary += "Deleted 定额库消耗: $deletedConsume"
    $summary += "Deleted 定额库: $deletedQuota"
    $summary += "Deleted 定额节索引: $deletedSections"
    $summary += "Deleted 定额库索引: $deletedIndex"
    $summary += "Deleted 材料单价库 supplements: $deletedMaterials"
    $summary += "Deleted 台班定额库 supplements: $deletedMachines"
    $summary += "Remaining migrated rows: " + (($afterRows | Measure-Object count -Sum).Sum)
    $summary += "Before report: $beforePath"
    $summary += "Resource candidate report: $resourcePath"
    $summary | Set-Content -LiteralPath $summaryPath -Encoding UTF8

    Write-Output ($summary -join [Environment]::NewLine)
}
finally {
    $cn.Close()
}

