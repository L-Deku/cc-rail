$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..')).Path
$autoMatchPath = Join-Path $repoRoot 'tools\RecoExpandPanel\AutoMatchFeature.cs'
$templateFillPath = Join-Path $repoRoot 'tools\RecoExpandPanel\TemplateFillPanel.cs'
$autoMatch = [System.IO.File]::ReadAllText($autoMatchPath, [System.Text.Encoding]::UTF8)
$templateFill = [System.IO.File]::ReadAllText($templateFillPath, [System.Text.Encoding]::UTF8)
$guard = 'column.SortMode = DataGridViewColumnSortMode.NotSortable;'

if (-not $autoMatch.Contains($guard)) {
    throw 'Auto-match grid column sorting is still enabled.'
}
if (-not $templateFill.Contains($guard)) {
    throw 'Template-fill grid column sorting is still enabled.'
}

Write-Host 'PASS: header sorting is disabled for auto-match and template-fill grids.'
