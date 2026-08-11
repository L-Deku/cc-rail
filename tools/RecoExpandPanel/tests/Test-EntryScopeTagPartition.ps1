param(
  [Parameter(Mandatory)][string]$QuotaDll
)

$ErrorActionPreference = 'Stop'

function Assert-Equal {
  param($Actual, $Expected, [string]$Message)
  if ($Actual -ne $Expected) { throw $Message }
}

$assembly = [Reflection.Assembly]::Load([IO.File]::ReadAllBytes((Resolve-Path -LiteralPath $QuotaDll).Path))
$scopeType = $assembly.GetType('RecoQuotaRecommend.EntryScope', $true)
$tagProperty = $scopeType.GetProperty('Tag', [Reflection.BindingFlags]'Instance,Public')
if ($tagProperty -eq $null) { throw 'EntryScope.Tag was not found.' }

function New-ScopeTag {
  param([string]$Partition, [string]$MethodNo, [string]$EntryCode)
  $scope = [Activator]::CreateInstance($scopeType, $true)
  $scopeType.GetField('SoftwarePartition').SetValue($scope, $Partition)
  $scopeType.GetField('MethodNo').SetValue($scope, $MethodNo)
  $scopeType.GetField('MatchedEntryCode').SetValue($scope, $EntryCode)
  return [string]$tagProperty.GetValue($scope, $null)
}

$tag30 = New-ScopeTag '2020' '30号文' '0616-02-01'
$tag101 = New-ScopeTag '2020' '101号文估算' '0616-02-01'
$tag2024 = New-ScopeTag '2024' 'TB 10801—2024' '0616-02-01'
Assert-Equal $tag30 '2020:30号文:0616-02-01' '30-method tag is incorrect.'
Assert-Equal $tag101 '2020:101号文估算:0616-02-01' '101-estimate tag is incorrect.'
Assert-Equal $tag2024 '2024:TB 10801—2024:0616-02-01' '2024 tag is incorrect.'
if ($tag30 -eq $tag101 -or $tag30 -eq $tag2024 -or $tag101 -eq $tag2024) {
  throw 'EntryScope tags are not isolated by partition and method number.'
}
Assert-Equal (New-ScopeTag '' '30号文' '0616-02-01') '' 'Unknown partition produced a formal tag.'
Assert-Equal (New-ScopeTag '2020' '' '0616-02-01') '' 'Unknown method produced a formal tag.'
Assert-Equal (New-ScopeTag '2020' '30号文' '2') '' 'Invalid entry code produced a formal tag.'
Assert-Equal (New-ScopeTag '2020' '30号文' ('0616' + [char]0x2010 + '02-01')) '2020:30号文:0616-02-01' 'Unicode dash was not normalized in the tag.'

Write-Host 'PASS T23 EntryScope tags isolate partition and method number'
Write-Host 'PASS unknown identity and invalid entry code do not produce formal tags'
