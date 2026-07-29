param(
    [Parameter(Mandatory = $true)][string]$TemplatePath,
    [Parameter(Mandatory = $true)][string]$DllPath
)

$ErrorActionPreference = 'Stop'

$templateFile = (Resolve-Path -LiteralPath $TemplatePath).Path
$assemblyFile = (Resolve-Path -LiteralPath $DllPath).Path
$dllDir = Split-Path -Parent $assemblyFile
foreach ($dependency in @('NPOI.dll', 'NPOI.OpenXmlFormats.dll', 'NPOI.OpenXml4Net.dll', 'NPOI.OOXML.dll')) {
    $dependencyPath = Join-Path $dllDir $dependency
    if (Test-Path -LiteralPath $dependencyPath) {
        [void][System.Reflection.Assembly]::LoadFrom($dependencyPath)
    }
}

$flags = [System.Reflection.BindingFlags]'Public,NonPublic,Static,Instance'
$type = [System.Reflection.Assembly]::LoadFrom($assemblyFile).GetType('RecoNet.FormPanel')
$templateType = $type.GetNestedType('FillTemplate', $flags)
$normalize = $type.GetMethod('NormalizeLegacyFillTemplateRows', $flags)
Add-Type -AssemblyName System.Web.Extensions
$serializer = New-Object System.Web.Script.Serialization.JavaScriptSerializer
$serializer.MaxJsonLength = 16 * 1024 * 1024

$json = [IO.File]::ReadAllText($templateFile, [Text.Encoding]::UTF8)
$template = $serializer.Deserialize($json, $templateType)
$before = $template.Rows.Count
$removed = [int]$normalize.Invoke($null, @($template))
$generated = @($template.Rows | Where-Object { $_.Origin -eq 'generated' }).Count
$manual = @($template.Rows | Where-Object { $_.Origin -eq 'manual' }).Count

$stamp = Get-Date -Format 'yyyyMMdd-HHmmss'
$backup = $templateFile + '.backup-' + $stamp
[IO.File]::Copy($templateFile, $backup, $false)
[IO.File]::WriteAllText($templateFile, $serializer.Serialize($template), [Text.Encoding]::UTF8)

$verify = $serializer.Deserialize([IO.File]::ReadAllText($templateFile, [Text.Encoding]::UTF8), $templateType)
if ($verify.Rows.Count -ne ($before - $removed)) {
    throw 'Template migration verification failed.'
}

[pscustomobject]@{
    Template = $templateFile
    Backup = $backup
    BeforeRows = $before
    RemovedDuplicateRows = $removed
    AfterRows = $verify.Rows.Count
    GeneratedRows = $generated
    ManualRows = $manual
}
