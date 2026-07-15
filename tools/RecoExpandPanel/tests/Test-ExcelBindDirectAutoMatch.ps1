$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$repoRoot = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot "..\..\..")).Path
$excelLinkPath = Join-Path $repoRoot "tools\RecoExpandPanel\ExcelLinkFeature.cs"
$autoMatchPath = Join-Path $repoRoot "tools\RecoExpandPanel\AutoMatchFeature.cs"
$excelLink = [System.IO.File]::ReadAllText($excelLinkPath, [System.Text.Encoding]::UTF8)
$autoMatch = [System.IO.File]::ReadAllText($autoMatchPath, [System.Text.Encoding]::UTF8)

function Assert-Contains([string]$Text, [string]$Expected, [string]$Message) {
  if (-not $Text.Contains($Expected)) {
    throw $Message
  }
}

function Assert-NotContains([string]$Text, [string]$Unexpected, [string]$Message) {
  if ($Text.Contains($Unexpected)) {
    throw $Message
  }
}

Assert-Contains $excelLink 'bindExcel.Click += delegate { ShowAutoMatchDialog(mainForm); };' 'The context-menu entry does not open AutoMatchDialog directly.'
Assert-Contains $excelLink 'if (e.Control && e.Shift && e.KeyCode == Keys.E)' 'The Ctrl+Shift+E quick-bind entry was removed.'
Assert-Contains $excelLink 'ShowQuickBindPanel(mainForm);' 'The QuickBindPanel entry was removed.'
Assert-NotContains $excelLink 'else if (e.Control && e.KeyCode == Keys.E)' 'The legacy Ctrl+E entry still exists.'
Assert-NotContains $excelLink 'ExcelSmartBindPanel' 'The legacy ExcelSmartBindPanel code still exists.'
Assert-NotContains $excelLink 'SmartBindSelectedQuotasToExcel' 'The legacy smart-bind entry method still exists.'
Assert-Contains $excelLink 'private static void SaveAutoMatchPreviewAccepted(Form mainForm, SqlConnection conn, List<AiMatchPreviewItem> accepted)' 'The auto-match save callback was not moved to the entry layer.'
Assert-Contains $autoMatch 'manualMatchButton.Text = "\u624b\u52a8\u5339\u914d";' 'The manual-match button inside AutoMatchDialog was removed.'
Assert-Contains $autoMatch 'private void PollManualMatchCell()' 'The manual-match polling flow inside AutoMatchDialog was removed.'
Assert-Contains $autoMatch 'private void AcceptCurrentItem()' 'The single-bind flow inside AutoMatchDialog was removed.'
Assert-Contains $autoMatch 'private void AcceptCheckedItems()' 'The bind-all flow inside AutoMatchDialog was removed.'

Write-Host 'PASS: direct auto-match entry, legacy panel removal, and retained AutoMatchDialog flows.'
