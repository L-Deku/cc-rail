param(
    [Parameter(Mandatory = $true)]
    [string]$RecoQuotaRecommendDll
)

$ErrorActionPreference = "Stop"

function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}

function Set-Field {
    param([Type]$Type, [object]$Instance, [string]$Name, [object]$Value)
    $field = $Type.GetField($Name, [Reflection.BindingFlags]'Instance,Public,NonPublic')
    if ($null -eq $field) { throw "Missing field: $Name" }
    $rawValue = if ($null -eq $Value) { $null } else { $Value.PSObject.BaseObject }
    $field.SetValue($Instance.PSObject.BaseObject, $rawValue)
}

function New-RecommendationRows {
    param([Reflection.Assembly]$Assembly)
    $itemType = $Assembly.GetType('RecoQuotaRecommend.ExcelQuantityItem', $true)
    $rowType = $Assembly.GetType('RecoQuotaRecommend.RecommendationRow', $true)
    $item = [Activator]::CreateInstance($itemType, $true)
    Set-Field $itemType $item 'Name' 'Cable'
    Set-Field $itemType $item 'Unit' 'm'
    Set-Field $itemType $item 'ValueText' '10'
    $row = [Activator]::CreateInstance($rowType, $true)
    Set-Field $rowType $row 'Item' $item
    Set-Field $rowType $row 'QuotaCode' 'DY-1'
    Set-Field $rowType $row 'QuotaName' 'Cable quota'
    Set-Field $rowType $row 'QuotaUnit' 'm'
    Set-Field $rowType $row 'TargetKind' 'quota'
    $listType = [System.Collections.Generic.List``1].MakeGenericType($rowType)
    $list = [Activator]::CreateInstance($listType)
    [void]$list.GetType().GetMethod('Add').Invoke($list, [object[]]@($row))
    return ,$list
}

function New-Scope {
    param([Reflection.Assembly]$Assembly, [string]$MethodNo, [string]$EntryCode)
    $scopeType = $Assembly.GetType('RecoQuotaRecommend.EntryScope', $true)
    $scope = [Activator]::CreateInstance($scopeType, $true)
    Set-Field $scopeType $scope 'SoftwarePartition' '2020'
    Set-Field $scopeType $scope 'Method' '2020'
    Set-Field $scopeType $scope 'MethodNo' $MethodNo
    Set-Field $scopeType $scope 'MatchedEntryCode' $EntryCode
    Set-Field $scopeType $scope 'ProjectEntryCode' $EntryCode
    Set-Field $scopeType $scope 'EntryName' ('Entry-' + $EntryCode)
    return ,$scope
}

$dll = (Resolve-Path -LiteralPath $RecoQuotaRecommendDll).Path
$assembly = [Reflection.Assembly]::LoadFrom($dll)
$storeType = $assembly.GetType('RecoQuotaRecommend.MappingStore', $true)
$load = $storeType.GetMethod('LoadForTesting', [Reflection.BindingFlags]'Static,NonPublic')
$accept = $storeType.GetMethod('Accept', [Reflection.BindingFlags]'Instance,Public,NonPublic')
Assert-True ($null -ne $load -and $null -ne $accept) 'MappingStore test entry points missing.'

$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
$tempRoot = Join-Path $tempBase ('RecoMappingStore-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($tempRoot)
$path = Join-Path $tempRoot 'mapping-boxes.jsonl'
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
$unknown = '{"record_type":"future_type","raw":"keep"}'
[IO.File]::WriteAllText($path, $unknown + "`n", $utf8NoBom)

try {
    foreach ($case in @(
        [pscustomobject]@{ MethodNo = '2020'; Entry = '0101-01' },
        [pscustomobject]@{ MethodNo = '101-estimate'; Entry = '0101-01' },
        [pscustomobject]@{ MethodNo = '2020'; Entry = '0101-02' }
    )) {
        Write-Host ('RUN MappingStore ' + $case.MethodNo + ' ' + $case.Entry)
        $loadArguments = New-Object 'object[]' 2
        $loadArguments[0] = [string]$path
        $loadArguments[1] = [string]'2020'
        $store = $load.Invoke($null, $loadArguments)
        Write-Host '  loaded'
        $rows = New-RecommendationRows $assembly
        Write-Host '  rows'
        $scope = New-Scope $assembly $case.MethodNo $case.Entry
        Write-Host '  scope'
        $arguments = New-Object 'object[]' 2
        $arguments[0] = $rows.PSObject.BaseObject
        $arguments[1] = $scope.PSObject.BaseObject
        [void]$accept.Invoke($store.PSObject.BaseObject, $arguments)
    }

    $lines = [IO.File]::ReadAllLines($path, [Text.Encoding]::UTF8)
    Assert-True ($lines[0] -eq $unknown) 'MappingStore changed the unknown record.'
    $parsed = @($lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { ConvertFrom-Json $_ })
    $boxes = @($parsed | Where-Object { $_.record_type -eq 'mapping_box' })
    $contexts = @($parsed | Where-Object { $_.record_type -eq 'mapping_context' })
    Assert-True ($boxes.Count -eq 1) 'MappingStore split ordinary relation by method/entry.'
    Assert-True ($contexts.Count -eq 3) 'MappingStore contexts did not coexist.'
    foreach ($forbidden in @('method','method_no','project_id','entry_codes','entry_code','entry_name','formula_method_no')) {
        Assert-True ($null -eq $boxes[0].PSObject.Properties[$forbidden]) ("mapping_box contains context field: " + $forbidden)
    }
    $pairs = @($contexts | ForEach-Object { $_.method_no + '|' + $_.entry_code } | Sort-Object -Unique)
    Assert-True ($pairs.Count -eq 3) 'MappingStore context identity lost method_no or entry_code.'
    Write-Host 'PASS T22 MappingStore Accept dual record type and method/entry isolation'
}
finally {
    $resolved = [IO.Path]::GetFullPath($tempRoot)
    if (-not $resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase)) {
        throw 'Refusing to clean unexpected test directory.'
    }
    if ([IO.Directory]::Exists($resolved)) { [IO.Directory]::Delete($resolved, $true) }
}
