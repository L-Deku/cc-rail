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
Assert-Contains $excelLink 'private static int SaveAutoMatchPreviewAccepted(Form mainForm, SqlConnection conn, List<AiMatchPreviewItem> accepted)' 'The auto-match save callback does not report the saved count from the entry layer.'
Assert-Contains $excelLink 'Dictionary<ExcelQuotaLink, string> savedQuantityNames = new Dictionary<ExcelQuotaLink, string>();' 'The auto-match save callback does not collect saved quantity names for mapping feedback.'
Assert-Contains $excelLink 'savedQuantityNames[item.Link] = item.QuantityName ?? "";' 'The auto-match save callback does not associate saved bindings with their quantity names.'
Assert-Contains $excelLink 'RecordBindingsToMappingStore(savedQuantityNames);' 'Accepted auto-match bindings are not written to the recommendation mapping pool.'
Assert-Contains $autoMatch 'manualMatchButton.Text = "\u624b\u52a8\u5339\u914d";' 'The manual-match button inside AutoMatchDialog was removed.'
Assert-Contains $autoMatch 'private void PollManualMatchCell()' 'The manual-match polling flow inside AutoMatchDialog was removed.'
Assert-Contains $autoMatch 'private void AcceptCurrentItem()' 'The single-bind flow inside AutoMatchDialog was removed.'
Assert-Contains $autoMatch 'private void AcceptCheckedItems()' 'The bind-all flow inside AutoMatchDialog was removed.'
Assert-Contains $excelLink 'return SaveAutoMatchPreviewAccepted(mainForm, conn, accepted);' 'The auto-match save callback does not return the saved count.'
Assert-Contains $autoMatch 'public event Func<List<AiMatchPreviewItem>, int> Accepted;' 'AutoMatchDialog does not receive the saved count.'
Assert-Contains $autoMatch 'status.Text = "\u5df2\u5168\u90e8\u7ed1\u5b9a " + saved.ToString(CultureInfo.InvariantCulture) + " \u6761\u3002";' 'Bind-all does not report the saved count in the status label.'
Assert-NotContains $autoMatch 'DialogResult = DialogResult.OK;' 'Bind-all still sets a success dialog result and closes the window.'

Write-Host 'PASS: direct auto-match entry, legacy panel removal, and retained AutoMatchDialog flows.'
