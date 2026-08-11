param(
    [Parameter(Mandatory = $true)]
    [string]$RecoExpandPanelDll
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
    $field.SetValue($Instance, $Value)
}

function Add-ToList {
    param([object]$List, [object]$Value)
    $method = $List.GetType().GetMethod('Add')
    [void]$method.Invoke($List, [object[]]@($Value))
}

function New-FeedbackGroup {
    param(
        [Type]$GroupType,
        [Type]$TargetType,
        [string]$MethodNo,
        [string]$EntryCode
    )
    $group = [Activator]::CreateInstance($GroupType, $true)
    Set-Field $GroupType $group 'QuantityName' 'Cable'
    Set-Field $GroupType $group 'QuantityUnit' 'm'
    Set-Field $GroupType $group 'Method' '2020'
    Set-Field $GroupType $group 'MethodNo' $MethodNo
    Set-Field $GroupType $group 'SoftwarePartition' '2020'
    Set-Field $GroupType $group 'EntryCode' $EntryCode
    Set-Field $GroupType $group 'EntryName' ('Entry-' + $EntryCode)
    $target = [Activator]::CreateInstance($TargetType, $true)
    Set-Field $TargetType $target 'Kind' 'quota'
    Set-Field $TargetType $target 'Code' 'DY-1'
    Set-Field $TargetType $target 'Name' 'Cable quota'
    Set-Field $TargetType $target 'Unit' 'm'
    Set-Field $TargetType $target 'EntryCode' $EntryCode
    Set-Field $TargetType $target 'EntryName' ('Entry-' + $EntryCode)
    $targets = $GroupType.GetField('Targets', [Reflection.BindingFlags]'Instance,Public,NonPublic').GetValue($group)
    Add-ToList $targets $target
    return ,$group
}

$dll = (Resolve-Path -LiteralPath $RecoExpandPanelDll).Path
$assembly = [Reflection.Assembly]::LoadFrom($dll)
$formType = $assembly.GetType('RecoNet.FormPanel', $true)
$nestedFlags = [Reflection.BindingFlags]'NonPublic'
$groupType = $formType.GetNestedType('MappingFeedbackGroup', $nestedFlags)
$targetType = $formType.GetNestedType('MappingFeedbackTarget', $nestedFlags)
$method = $formType.GetMethod('UpsertMappingBoxGroup', [Reflection.BindingFlags]'Static,NonPublic')
$saveMethod = $formType.GetMethod('SaveMappingGroupsToLocalFile', [Reflection.BindingFlags]'Static,NonPublic')
Assert-True ($null -ne $groupType -and $null -ne $targetType -and $null -ne $method -and $null -ne $saveMethod) 'ExcelLink local mutation types were not found.'

$dictionaryType = [System.Collections.Generic.Dictionary[string,string]]
$listType = [System.Collections.Generic.List``1].MakeGenericType($dictionaryType)
$rows = [Activator]::CreateInstance($listType)
$boxes = [Activator]::CreateInstance($listType)
$contexts = [Activator]::CreateInstance($listType)
$groups = @(
    (New-FeedbackGroup $groupType $targetType '2020' '0101-01'),
    (New-FeedbackGroup $groupType $targetType '101-estimate' '0101-01'),
    (New-FeedbackGroup $groupType $targetType '2020' '0101-02')
)

foreach ($group in $groups) {
    $arguments = New-Object 'object[]' 4
    $arguments[0] = $rows
    $arguments[1] = $boxes
    $arguments[2] = $contexts
    $arguments[3] = $group
    [void]$method.Invoke($null, $arguments)
}

Assert-True ($boxes.Count -eq 1) 'Ordinary relation was split by method or entry.'
Assert-True ($contexts.Count -eq 3) 'Method/entry contexts did not coexist.'
$box = $boxes[0]
foreach ($forbidden in @('method','method_no','project_id','entry_codes','entry_code','entry_name','formula_method_no')) {
    Assert-True (-not $box.ContainsKey($forbidden)) ("mapping_box contains context field: " + $forbidden)
}
Assert-True ($box['record_type'] -eq 'mapping_box' -and $box['software_partition'] -eq '2020') 'mapping_box partition/type missing.'

$methodEntryPairs = @($contexts | ForEach-Object { $_['method_no'] + '|' + $_['entry_code'] } | Sort-Object -Unique)
Assert-True ($methodEntryPairs.Count -eq 3) 'Context identity lost method_no or entry_code.'
Assert-True (@($contexts | Where-Object { $_['record_type'] -ne 'mapping_context' }).Count -eq 0) 'Non-context row emitted in context set.'

$tempBase = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
$tempRoot = Join-Path $tempBase ('RecoExcelLinkLocal-' + [Guid]::NewGuid().ToString('N'))
[void][IO.Directory]::CreateDirectory($tempRoot)
$path = Join-Path $tempRoot 'mapping-boxes.jsonl'
$unknown = '{"record_type":"future_type","raw":"keep"}'
[IO.File]::WriteAllText($path, $unknown + "`n", (New-Object System.Text.UTF8Encoding($false)))
try {
    $groupListType = [System.Collections.Generic.List``1].MakeGenericType($groupType)
    $groupList = [Activator]::CreateInstance($groupListType)
    foreach ($group in $groups) { Add-ToList $groupList $group }
    $saveArguments = New-Object 'object[]' 4
    $saveArguments[0] = [string]$path
    $saveArguments[1] = [string]'2020'
    $saveArguments[2] = $groupList
    $saveArguments[3] = [string]'offline-excel-endpoint'
    $saveResult = $saveMethod.Invoke($null, $saveArguments)
    $succeeded = $saveResult.GetType().GetProperty('Succeeded', [Reflection.BindingFlags]'Instance,Public,NonPublic').GetValue($saveResult, $null)
    Assert-True $succeeded 'ExcelLink local-file endpoint returned failure.'
    $lines = [IO.File]::ReadAllLines($path, [Text.Encoding]::UTF8)
    Assert-True ($lines[0] -eq $unknown) 'ExcelLink endpoint changed unknown row.'
    $saved = @($lines | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object { ConvertFrom-Json $_ })
    Assert-True (@($saved | Where-Object { $_.record_type -eq 'mapping_box' }).Count -eq 1) 'ExcelLink endpoint persisted split mapping_box rows.'
    Assert-True (@($saved | Where-Object { $_.record_type -eq 'mapping_context' }).Count -eq 3) 'ExcelLink endpoint did not persist all contexts.'
    Write-Host 'PASS T17/T22 ExcelLink local-file endpoint preservation and method/entry isolation'
}
finally {
    $resolved = [IO.Path]::GetFullPath($tempRoot)
    if (-not $resolved.StartsWith($tempBase, [StringComparison]::OrdinalIgnoreCase)) { throw 'Refusing to clean unexpected test directory.' }
    if ([IO.Directory]::Exists($resolved)) { [IO.Directory]::Delete($resolved, $true) }
}
