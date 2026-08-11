$ErrorActionPreference = 'Stop'

$testDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$sourceDir = Split-Path -Parent $testDir

function Get-Section([string]$text, [string]$startMarker, [string]$endMarker) {
    $start = $text.IndexOf($startMarker, [StringComparison]::Ordinal)
    if ($start -lt 0) { throw "Missing start marker: $startMarker" }
    $end = $text.IndexOf($endMarker, $start + $startMarker.Length, [StringComparison]::Ordinal)
    if ($end -lt 0) { throw "Missing end marker: $endMarker" }
    return $text.Substring($start, $end - $start)
}

$agentExecutor = [IO.File]::ReadAllText((Join-Path $sourceDir 'AgentExecutor.cs'))
$openProject = Get-Section $agentExecutor 'private static SqlConnection GetOpenProjectConnection' 'private static T WithOpenProjectConnectionOnUi<T>'
if (-not $openProject.Contains('EnsureOpen(conn);')) {
    throw 'Closed host project connections are rejected instead of being safely reopened on the same SqlConnection.'
}
if ($openProject.Contains('if (conn.State != ConnectionState.Open)') -or
    $openProject.Contains('当前项目数据库连接未打开')) {
    throw 'The old state-only rejection path was reintroduced.'
}
if ($openProject.Contains('new SqlConnection(') -or $openProject.Contains('AgentCreateWorkConnection(')) {
    throw 'The host project connection helper must not clone the connection.'
}

$uiConnection = Get-Section $agentExecutor 'private static T WithOpenProjectConnectionOnUi<T>' 'private static AgentSelectionSnapshot CaptureAgentSelection'
if (-not $uiConnection.Contains('mainForm.InvokeRequired') -or
    -not $uiConnection.Contains('mainForm.Invoke(') -or
    -not $uiConnection.Contains('GetOpenProjectConnection(mainForm)') -or
    -not $uiConnection.Contains('expectedConnection') -or
    -not $uiConnection.Contains('expectedConnectionIdentity') -or
    -not $uiConnection.Contains('GetProjectConnectionIdentity(conn)') -or
    -not $uiConnection.Contains('Object.ReferenceEquals') -or
    -not $uiConnection.Contains('mainForm.IsDisposed') -or
    -not $uiConnection.Contains('mainForm.Disposing') -or
    -not $uiConnection.Contains('mainForm.IsHandleCreated')) {
    throw 'Background agent reads are not marshalled onto the host UI thread before borrowing its connection.'
}
if ($uiConnection.Contains('new SqlConnection(') -or
    $uiConnection.Contains('using (SqlConnection conn') -or
    $uiConnection.Contains('using (conn)') -or
    $uiConnection.Contains('conn.Dispose(') -or $uiConnection.Contains('conn.Close(') -or
    $uiConnection.Contains('conn.ChangeDatabase(')) {
    throw 'The UI-thread connection helper clones, owns, closes, or changes the borrowed host connection.'
}

$templateFeature = [IO.File]::ReadAllText((Join-Path $sourceDir 'TemplateFillFeature.cs'))
$chapterNames = Get-Section $templateFeature 'private static Dictionary<string, string> LoadChapterNameMap' 'private static string ChapterTreeDisplayName'
if (-not $chapterNames.Contains('GetOpenProjectConnection(mainForm)') -or $chapterNames.Contains('AgentCreateWorkConnection(mainForm)')) {
    throw 'SmartFill chapter-name loading still clones the host project connection.'
}

$applyFill = Get-Section $templateFeature 'private static string ApplyFill' 'private static List<List<FillPreviewItem>> CollectFullyWrittenNameDrivenGroups'
if (-not $applyFill.Contains('SqlConnection conn = GetOpenProjectConnection(mainForm);')) {
    throw 'ApplyFill does not borrow the safely reopened host project connection.'
}
if ($applyFill.Contains('AgentCreateWorkConnection(mainForm)') -or
    $applyFill.Contains('using (SqlConnection conn') -or
    $applyFill.Contains('conn.Dispose(') -or $applyFill.Contains('conn.Close(') -or
    $applyFill.Contains('conn.ChangeDatabase(')) {
    throw 'ApplyFill still clones, owns, closes, or changes the borrowed host project connection.'
}
if (-not $applyFill.Contains('conn.BeginTransaction()') -or
    -not $applyFill.Contains('transaction.Commit();') -or
    -not $applyFill.Contains('transaction.Rollback();')) {
    throw 'ApplyFill no longer keeps its current-project writes inside the existing transaction.'
}

$templatePanel = [IO.File]::ReadAllText((Join-Path $sourceDir 'TemplateFillPanel.cs'))
$unitList = Get-Section $templatePanel 'private static List<string> ListAgentUnits' 'private void OnDeleteTemplate'
if (-not $unitList.Contains('GetOpenProjectConnection(mainForm)') -or $unitList.Contains('AgentCreateWorkConnection(mainForm)')) {
    throw 'SmartFill target-unit loading still clones the host project connection.'
}

$reloadSheets = Get-Section $templatePanel 'private void ReloadSourceSheets' 'private string GetSelectedTargetWorkbookPath'
$buildTemplate = Get-Section $templatePanel 'private void OnBuild' 'private void OnPreview'
foreach ($section in @($reloadSheets, $buildTemplate)) {
    if (-not $section.Contains('GetOpenProjectConnection(mainForm)') -or
        $section.Contains('AgentCreateWorkConnection(mainForm)') -or
        $section.Contains('using (SqlConnection conn') -or
        $section.Contains('using (conn)') -or
        $section.Contains('conn.Dispose(') -or $section.Contains('conn.Close(') -or
        $section.Contains('conn.ChangeDatabase(')) {
        throw 'A template read/build path still clones, owns, closes, or changes the host project connection.'
    }
}

$multiplier = [IO.File]::ReadAllText((Join-Path $sourceDir 'MultiplierFeature.cs'))
$contextAction = Get-Section $multiplier 'private static void ApplyContextMenuAction' 'private static AgentCommand BuildContextMenuCommand'
if (-not $contextAction.Contains('GetOpenProjectConnection(mainForm)') -or
    $contextAction.Contains('AgentCreateWorkConnection(mainForm)') -or
    $contextAction.Contains('using (SqlConnection conn') -or
    $contextAction.Contains('using (conn)') -or
    $contextAction.Contains('conn.Dispose(') -or $contextAction.Contains('conn.Close(') -or
    $contextAction.Contains('conn.ChangeDatabase(')) {
    throw 'The multiplier preview still clones, owns, closes, or changes the host project connection.'
}

$executePlan = Get-Section $agentExecutor 'private static string ExecuteAgentPlan' 'private static int ExecuteAgentFieldUpdate'
if (-not $executePlan.Contains('GetOpenProjectConnection(mainForm)') -or
    $executePlan.Contains('AgentCreateWorkConnection(mainForm)') -or
    $executePlan.Contains('using (SqlConnection conn') -or
    $executePlan.Contains('using (conn)') -or
    $executePlan.Contains('conn.Dispose(') -or $executePlan.Contains('conn.Close(') -or
    $executePlan.Contains('conn.ChangeDatabase(')) {
    throw 'The user-confirmed agent execution still clones, owns, closes, or changes the host project connection.'
}
if (-not $executePlan.Contains('conn.BeginTransaction()') -or
    -not $executePlan.Contains('transaction.Commit();') -or
    -not $executePlan.Contains('transaction.Rollback();')) {
    throw 'Agent execution no longer preserves its current-project transaction.'
}
if (-not $executePlan.Contains('plan.ProjectConnection') -or
    -not $executePlan.Contains('plan.ProjectConnectionIdentity') -or
    -not $executePlan.Contains('GetProjectConnectionIdentity(conn)') -or
    -not $executePlan.Contains('Object.ReferenceEquals') -or
    -not $executePlan.Contains('undo.ProjectConnection = conn;')) {
    throw 'Agent execution does not reject a preview after the user switches projects.'
}

$agentPlan = Get-Section $agentExecutor 'private sealed class AgentPlan' 'private sealed class AgentUndoRow'
if (-not $agentPlan.Contains('public SqlConnection ProjectConnection;')) {
    throw 'Agent plans do not retain the exact host project connection used to build the preview.'
}
if (-not $agentPlan.Contains('public string ProjectConnectionIdentity;')) {
    throw 'Agent plans do not retain the immutable server/database identity used to build the preview.'
}
$buildPlan = Get-Section $agentExecutor 'private static AgentPlan BuildAgentPlan' 'private static void BuildMultiplyPlan'
if (-not $buildPlan.Contains('plan.ProjectConnection = conn;') -or
    -not $buildPlan.Contains('plan.ProjectConnectionIdentity = GetProjectConnectionIdentity(conn);')) {
    throw 'Agent previews do not record the exact host project connection used to build them.'
}

$undoRecord = Get-Section $agentExecutor 'private sealed class AgentUndoRecord' 'private static List<AgentUndoRecord> GetAgentUndoStack'
if (-not $undoRecord.Contains('public SqlConnection ProjectConnection;') -or
    -not $undoRecord.Contains('public string ProjectConnectionIdentity;')) {
    throw 'Agent undo records are not scoped to the project connection where the action ran.'
}
$undoPlan = Get-Section $agentExecutor 'private static AgentPlan BuildAgentUndoPlan' 'private static string ExecuteAgentUndo'
$redoPlan = Get-Section $agentExecutor 'private static AgentPlan BuildAgentRedoPlan' 'private static string ExecuteAgentRedo'
foreach ($section in @($undoPlan, $redoPlan)) {
    if (-not $section.Contains('GetOpenProjectConnection(mainForm)') -or
        -not $section.Contains('record.ProjectConnection') -or
        -not $section.Contains('record.ProjectConnectionIdentity') -or
        -not $section.Contains('GetProjectConnectionIdentity(conn)') -or
        -not $section.Contains('Object.ReferenceEquals') -or
        -not $section.Contains('plan.ProjectConnection = record.ProjectConnection;') -or
        -not $section.Contains('plan.ProjectConnectionIdentity = record.ProjectConnectionIdentity;')) {
        throw 'Agent undo/redo can cross from the original project into a newly opened project.'
    }
}
foreach ($section in @($chapterNames, $applyFill, $unitList, $reloadSheets, $buildTemplate,
    $contextAction, $executePlan, $undoPlan, $redoPlan)) {
    if (-not $section.Contains('GetOpenProjectConnection(mainForm)') -or
        $section.Contains('new SqlConnection(') -or $section.Contains('.ConnectionString') -or
        $section.Contains('using (SqlConnection') -or $section.Contains('using (conn)') -or
        $section.Contains('conn.Dispose(') -or $section.Contains('conn.Close(') -or
        $section.Contains('conn.ChangeDatabase(')) {
        throw 'A project database consumer can still clone, own, close, or retarget the borrowed host connection.'
    }
}

$smartFill = [IO.File]::ReadAllText((Join-Path $sourceDir 'SmartFillFeature.cs'))
$learningScopes = Get-Section $smartFill 'private static List<SmartLearningScope> LoadSmartLearningScopes' 'private static bool IsSmartClassifiedEntryCode'
$previewStart = $smartFill.IndexOf('private static List<FillPreviewItem> BuildPreview_SmartFill', [StringComparison]::Ordinal)
if ($previewStart -lt 0) { throw 'Missing SmartFill preview method.' }
$smartPreview = $smartFill.Substring($previewStart)
$templateNameMatch = [IO.File]::ReadAllText((Join-Path $sourceDir 'TemplateFillNameMatch.cs'))
$projectQuotas = Get-Section $templateNameMatch 'private static List<ProjectQuota> LoadProjectQuotas' 'private static string BuildNameBindingTargetSetSignature'
$agentChat = [IO.File]::ReadAllText((Join-Path $sourceDir 'AgentChatFeature.cs'))
$chatInput = Get-Section $agentChat 'private void HandleUserInput' 'private void RunAgentPipeline'
foreach ($case in @(
    @($learningScopes, 'projectConn'), @($smartPreview, 'conn'),
    @($projectQuotas, 'conn'), @($chatInput, 'expectedConnection')
)) {
    $section = [string]$case[0]
    $variable = [string]$case[1]
    if (-not $section.Contains('GetOpenProjectConnection(mainForm)') -or
        $section.Contains($variable + '.ConnectionString') -or
        $section.Contains('using (' + $variable + ')') -or
        $section.Contains($variable + '.Dispose(') -or $section.Contains($variable + '.Close(') -or
        $section.Contains($variable + '.ChangeDatabase(')) {
        throw "A project database consumer can still clone, own, close, or retarget borrowed variable $variable."
    }
}

$pipeline = Get-Section $agentChat 'private void RunAgentPipeline' 'private void OnPipelineDone'
if ($pipeline.Contains('AgentCreateWorkConnection(mainForm)') -or
    ([regex]::Matches($pipeline, 'WithOpenProjectConnectionOnUi\s*\(').Count -ne 2) -or
    ([regex]::Matches($pipeline, 'RequestAgentParse\s*\(').Count -ne 1) -or
    ([regex]::Matches($pipeline, 'CollectAgentContext\s*\(').Count -ne 1) -or
    ([regex]::Matches($pipeline, 'BuildAgentPlan\s*\(').Count -ne 1) -or
    -not $pipeline.Contains('SqlConnection expectedConnection') -or
    -not $pipeline.Contains('string expectedConnectionIdentity')) {
    throw 'The background agent pipeline does not marshal both database stages through the UI-thread connection helper.'
}
$contextBeforeHttp = '(?s)AgentContext\s+context\s*=\s*WithOpenProjectConnectionOnUi.*?return\s+CollectAgentContext\(.*?\);\s*\}\);\s*llmResult\s*=\s*RequestAgentParse'
$httpBeforePlan = '(?s)llmResult\s*=\s*RequestAgentParse.*?plan\s*=\s*WithOpenProjectConnectionOnUi.*?return\s+BuildAgentPlan\('
$workerIndex = $pipeline.IndexOf('Thread worker = new Thread(delegate()', [StringComparison]::Ordinal)
$requestIndex = $pipeline.IndexOf('RequestAgentParse', [StringComparison]::Ordinal)
$doneIndex = $pipeline.IndexOf('BeginInvoke((MethodInvoker)delegate', [StringComparison]::Ordinal)
if (-not [regex]::IsMatch($pipeline, $contextBeforeHttp) -or
    -not [regex]::IsMatch($pipeline, $httpBeforePlan) -or
    $workerIndex -lt 0 -or $requestIndex -le $workerIndex -or $doneIndex -le $requestIndex) {
    throw 'The AI network parsing stage was removed instead of remaining in the background pipeline.'
}

$production = New-Object Text.StringBuilder
foreach ($file in @(Get-ChildItem -LiteralPath $sourceDir -Filter '*.cs' -File)) {
    [void]$production.AppendLine([IO.File]::ReadAllText($file.FullName))
}
if ($production.ToString().Contains('AgentCreateWorkConnection')) {
    throw 'A production path still calls AgentCreateWorkConnection instead of safely borrowing the host connection.'
}

Write-Host 'PASS: all project database paths safely borrow the host connection; background agent reads marshal to the UI thread.'
