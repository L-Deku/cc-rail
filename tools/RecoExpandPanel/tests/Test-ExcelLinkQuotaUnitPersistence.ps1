$ErrorActionPreference = 'Stop'

$repoRoot = [System.IO.Path]::GetFullPath((Join-Path $PSScriptRoot '..\..\..'))
$dll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL)) { $env:RECO_EXPAND_DLL } else { Join-Path $repoRoot 'RecoQuotaRecommend\bin\RecoExpandPanel.dll' }
if (-not (Test-Path -LiteralPath $dll)) { throw "Missing DLL: $dll" }

$assembly = [System.Reflection.Assembly]::LoadFrom($dll)
$panelType = $assembly.GetType('RecoNet.FormPanel', $true)
$flags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$linkType = $panelType.GetNestedType('ExcelQuotaLink', $flags)
$storeType = $panelType.GetNestedType('ExcelLinkStore', $flags)
$quotaUnitProperty = $linkType.GetProperty('QuotaUnit', $flags)
if ($null -eq $quotaUnitProperty) { throw 'ExcelQuotaLink.QuotaUnit is missing' }
$entryCodeProperty = $linkType.GetProperty('EntryCode', $flags)
$entryNameProperty = $linkType.GetProperty('EntryName', $flags)
$methodProperty = $linkType.GetProperty('Method', $flags)
if ($null -eq $entryCodeProperty -or $null -eq $entryNameProperty -or $null -eq $methodProperty) { throw 'ExcelQuotaLink learning entry context is missing' }

$link = [Activator]::CreateInstance($linkType).PSObject.BaseObject
$linkType.GetProperty('QuotaCode', $flags).SetValue($link, 'QY-317', $null)
$linkType.GetProperty('QuotaName', $flags).SetValue($link, 'cover rebar', $null)
$quotaUnitProperty.SetValue($link, 't', $null)
$entryCodeProperty.SetValue($link, '0101-01', $null)
$entryNameProperty.SetValue($link, 'entry name', $null)
$methodProperty.SetValue($link, '2024', $null)
$store = [Activator]::CreateInstance($storeType).PSObject.BaseObject
$links = $storeType.GetProperty('Links', $flags).GetValue($store, $null).PSObject.BaseObject
[void]$links.Add($link)

$serializer = New-Object System.Xml.Serialization.XmlSerializer($storeType)
$stream = New-Object System.IO.MemoryStream
$serializer.Serialize($stream, $store)
$xml = [System.Text.Encoding]::UTF8.GetString($stream.ToArray())
$stream.Dispose()
if (-not $xml.Contains('<QuotaUnit>t</QuotaUnit>')) { throw 'New XML did not persist QuotaUnit' }
if (-not $xml.Contains('<EntryCode>0101-01</EntryCode>') -or -not $xml.Contains('<EntryName>entry name</EntryName>') -or -not $xml.Contains('<Method>2024</Method>')) {
    throw 'New XML did not persist method/entry context'
}

$legacyXml = [regex]::Replace($xml, '<QuotaUnit>.*?</QuotaUnit>', '')
$legacyXml = [regex]::Replace($legacyXml, '<EntryCode>.*?</EntryCode>', '')
$legacyXml = [regex]::Replace($legacyXml, '<EntryName>.*?</EntryName>', '')
$legacyXml = [regex]::Replace($legacyXml, '<Method>.*?</Method>', '')
$legacyBytes = [System.Text.Encoding]::UTF8.GetBytes($legacyXml)
$legacyStream = New-Object System.IO.MemoryStream(,$legacyBytes)
$legacyStore = $serializer.Deserialize($legacyStream).PSObject.BaseObject
$legacyStream.Dispose()
$legacyLinks = $storeType.GetProperty('Links', $flags).GetValue($legacyStore, $null).PSObject.BaseObject
$legacyLink = $legacyLinks[0].PSObject.BaseObject
$legacyUnit = [string]$quotaUnitProperty.GetValue($legacyLink, $null)
if (-not [String]::IsNullOrEmpty($legacyUnit)) { throw "Legacy XML should load with an empty QuotaUnit, got '$legacyUnit'" }
foreach ($property in @($entryCodeProperty, $entryNameProperty, $methodProperty)) {
    if (-not [String]::IsNullOrEmpty([string]$property.GetValue($legacyLink, $null))) { throw "Legacy XML should load missing $($property.Name) as empty" }
}

Write-Host 'Test-ExcelLinkQuotaUnitPersistence: PASS (new XML persists unit/method/entry; legacy XML remains compatible)'
