$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

$dll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_EXPAND_DLL)) {
    $env:RECO_EXPAND_DLL
} else {
    Join-Path (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..\..')).Path 'RecoQuotaRecommend\bin\RecoExpandPanel.dll'
}

Add-Type -Path $dll
Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Text;

public static class NativeMenuTestApi
{
    [DllImport("user32.dll")]
    public static extern IntPtr CreatePopupMenu();

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool AppendMenu(IntPtr menu, uint flags, UIntPtr id, string text);

    [DllImport("user32.dll")]
    public static extern int GetMenuItemCount(IntPtr menu);

    [DllImport("user32.dll")]
    public static extern IntPtr GetSubMenu(IntPtr menu, int position);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool DestroyMenu(IntPtr menu);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    private static extern int GetMenuString(IntPtr menu, uint item, StringBuilder text, int maxCount, uint flags);

    public static string GetText(IntPtr menu, int position)
    {
        StringBuilder text = new StringBuilder(256);
        GetMenuString(menu, (uint)position, text, text.Capacity, 0x0400);
        return text.ToString();
    }
}
'@
$formType = [RecoNet.FormPanel]
$flags = [System.Reflection.BindingFlags]'NonPublic,Static,Instance,Public'

function Invoke-PrivateStatic([string]$name, [object[]]$arguments) {
    $method = $formType.GetMethod($name, $flags)
    if ($null -eq $method) {
        throw "Method not found: $name"
    }
    $invokeArguments = New-Object object[] $arguments.Count
    for ($i = 0; $i -lt $arguments.Count; $i++) {
        $invokeArguments[$i] = if ($null -eq $arguments[$i]) { $null } else { $arguments[$i].PSObject.BaseObject }
    }
    $result = $method.Invoke($null, $invokeArguments)
    for ($i = 0; $i -lt $arguments.Count; $i++) {
        $arguments[$i] = $invokeArguments[$i]
    }
    return $result
}

function Get-ToolChild($parent, [string]$text) {
    $matches = @($parent.DropDownItems | Where-Object { $_ -is [System.Windows.Forms.ToolStripMenuItem] -and $_.Text -eq $text })
    if ($matches.Count -ne 1) {
        throw "Expected one ToolStrip child '$text', got $($matches.Count)."
    }
    return $matches[0]
}

function Get-LegacyChild($parent, [string]$text) {
    $matches = @($parent.MenuItems | Where-Object { $_.Text.Replace('&', '').Trim() -eq $text })
    if ($matches.Count -ne 1) {
        throw "Expected one legacy child '$text', got $($matches.Count)."
    }
    return $matches[0]
}

function Assert-ToolMenu($root) {
    $expectedFields = @('工程数量', '定额编号', '单价', '定额调整')
    $actualFields = @($root.DropDownItems | Where-Object { $_ -is [System.Windows.Forms.ToolStripMenuItem] } | ForEach-Object Text)
    if (($actualFields -join '|') -ne ($expectedFields -join '|')) {
        throw "Unexpected ToolStrip fields: $($actualFields -join '|')"
    }

    foreach ($field in @('工程数量', '定额编号')) {
        $branch = Get-ToolChild $root $field
        $leaves = @($branch.DropDownItems | ForEach-Object Text)
        if (($leaves -join '|') -ne '乘系数|删系数') {
            throw "Unexpected leaves under $field`: $($leaves -join '|')"
        }
    }

    $priceLeaves = @((Get-ToolChild $root '单价').DropDownItems | ForEach-Object Text)
    if (($priceLeaves -join '|') -ne '乘系数') {
        throw "Unexpected leaves under 单价: $($priceLeaves -join '|')"
    }

    $adjustLeaves = @((Get-ToolChild $root '定额调整').DropDownItems | ForEach-Object Text)
    if (($adjustLeaves -join '|') -ne '调整输入|删除调整') {
        throw "Unexpected leaves under 定额调整: $($adjustLeaves -join '|')"
    }
}

function Assert-LegacyMenu($root) {
    $expectedFields = @('工程数量', '定额编号', '单价', '定额调整')
    $actualFields = @($root.MenuItems | ForEach-Object { $_.Text.Replace('&', '').Trim() })
    if (($actualFields -join '|') -ne ($expectedFields -join '|')) {
        throw "Unexpected legacy fields: $($actualFields -join '|')"
    }

    foreach ($field in @('工程数量', '定额编号')) {
        $branch = Get-LegacyChild $root $field
        $leaves = @($branch.MenuItems | ForEach-Object { $_.Text.Replace('&', '').Trim() })
        if (($leaves -join '|') -ne '乘系数|删系数') {
            throw "Unexpected legacy leaves under $field`: $($leaves -join '|')"
        }
    }

    $priceLeaves = @((Get-LegacyChild $root '单价').MenuItems | ForEach-Object { $_.Text.Replace('&', '').Trim() })
    if (($priceLeaves -join '|') -ne '乘系数') {
        throw "Unexpected legacy leaves under 单价: $($priceLeaves -join '|')"
    }

    $adjustLeaves = @((Get-LegacyChild $root '定额调整').MenuItems | ForEach-Object { $_.Text.Replace('&', '').Trim() })
    if (($adjustLeaves -join '|') -ne '调整输入|删除调整') {
        throw "Unexpected legacy leaves under 定额调整: $($adjustLeaves -join '|')"
    }
}

function Parse-Factor([string]$text) {
    $arguments = New-Object object[] 3
    $arguments[0] = $text
    $ok = [bool](Invoke-PrivateStatic 'TryParseFactor' $arguments)
    if (-not $ok) {
        throw "Factor parse failed: $text / $($arguments[2])"
    }
    return $arguments[1]
}

function Read-Field($value, [string]$name) {
    $field = $value.GetType().GetField($name, $flags)
    if ($null -ne $field) {
        Write-Output -NoEnumerate ($field.GetValue($value))
        return
    }
    $property = $value.GetType().GetProperty($name, $flags)
    if ($null -ne $property) {
        Write-Output -NoEnumerate ($property.GetValue($value, $null))
        return
    }
    throw "Member not found: $name"
}

function Build-Command([string]$target, [string]$action, [string]$value, [bool]$tree, [long]$unitId) {
    return Invoke-PrivateStatic 'BuildContextMenuCommand' @($target, $action, $value, $tree, $unitId)
}

function Get-NativeItemIndex([IntPtr]$menu, [string]$text) {
    for ($i = 0; $i -lt [NativeMenuTestApi]::GetMenuItemCount($menu); $i++) {
        if ([NativeMenuTestApi]::GetText($menu, $i) -eq $text) {
            return $i
        }
    }
    return -1
}

function Assert-NativeBranch([IntPtr]$root, [string]$field, [string[]]$expectedLeaves) {
    $fieldIndex = Get-NativeItemIndex $root $field
    if ($fieldIndex -lt 0) {
        throw "Native field not found: $field"
    }
    $branch = [NativeMenuTestApi]::GetSubMenu($root, $fieldIndex)
    if ($branch -eq [IntPtr]::Zero) {
        throw "Native field has no submenu: $field"
    }
    $actualLeaves = @()
    for ($i = 0; $i -lt [NativeMenuTestApi]::GetMenuItemCount($branch); $i++) {
        $actualLeaves += [NativeMenuTestApi]::GetText($branch, $i)
    }
    if (($actualLeaves -join '|') -ne ($expectedLeaves -join '|')) {
        throw "Unexpected native leaves under ${field}: $($actualLeaves -join '|')"
    }
}

$toolRoot = New-Object System.Windows.Forms.ToolStripMenuItem '乘系数'
[void]$toolRoot.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripMenuItem '乘到原来的工程量'))
[void]$toolRoot.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripMenuItem '乘到定额编号'))
[void](Invoke-PrivateStatic 'ConfigureFactorTargetMenu' @($toolRoot, $null, $true, $false))
[void](Invoke-PrivateStatic 'ConfigureFactorTargetMenu' @($toolRoot, $null, $true, $false))
Assert-ToolMenu $toolRoot

$legacyRoot = New-Object System.Windows.Forms.MenuItem '乘系数'
[void]$legacyRoot.MenuItems.Add((New-Object System.Windows.Forms.MenuItem '乘到原来的工程量'))
[void]$legacyRoot.MenuItems.Add((New-Object System.Windows.Forms.MenuItem '乘到定额编号'))
[void](Invoke-PrivateStatic 'ConfigureLegacyFactorTargetMenu' @($legacyRoot, $null))
[void](Invoke-PrivateStatic 'ConfigureLegacyFactorTargetMenu' @($legacyRoot, $null))
Assert-LegacyMenu $legacyRoot

$factor = Parse-Factor '0.9'
if ((Read-Field $factor 'Suffix') -ne '*0.9') { throw 'Bare factor must normalize to *0.9.' }
$factor = Parse-Factor '*0.9'
if ((Read-Field $factor 'Suffix') -ne '*0.9') { throw 'Explicit multiply factor changed unexpectedly.' }
$factor = Parse-Factor '/2'
if ((Read-Field $factor 'Suffix') -ne '/2') { throw 'Divide factor changed unexpectedly.' }

$treeCommand = Build-Command 'quantity' 'multiply' '0.9' $true 42
if ((Read-Field $treeCommand 'Type') -ne 'multiply_quantity') { throw 'Tree multiply command type mismatch.' }
if ((Read-Field $treeCommand 'Operator') -ne '*') { throw 'Tree multiply operator mismatch.' }
if ((Read-Field $treeCommand 'Factor') -ne '0.9') { throw 'Tree multiply factor mismatch.' }
if (-not [bool](Read-Field $treeCommand 'IncludeChildren')) { throw 'Tree command must include child entries.' }
$treeItems = Read-Field $treeCommand 'Items'
$treeUnits = Read-Field $treeCommand 'Units'
if ($treeItems.Count -ne 1 -or [string]$treeItems[0] -ne '@currentitem') { throw 'Tree command must target the current item.' }
if ($treeUnits.Count -ne 1 -or [string]$treeUnits[0] -ne '42') { throw 'Tree command must explicitly carry current unit id 42.' }

$gridCommand = Build-Command 'quota_code' 'remove' '/2' $false 0
if ((Read-Field $gridCommand 'Type') -ne 'remove_text') { throw 'Grid remove command type mismatch.' }
if ((Read-Field $gridCommand 'Target') -ne 'quota_code') { throw 'Grid remove target mismatch.' }
if ((Read-Field $gridCommand 'RemoveText') -ne '/2') { throw 'Grid remove fragment mismatch.' }
$gridItems = Read-Field $gridCommand 'Items'
if ($gridItems.Count -ne 1 -or [string]$gridItems[0] -ne '@selected') { throw 'Grid command must target selected quota rows.' }

$setAdjustment = Build-Command 'adjustment' 'set_adjustment' '/XG1' $true 42
if ((Read-Field $setAdjustment 'Type') -ne 'set_adjustment' -or (Read-Field $setAdjustment 'Mode') -ne 'set') {
    throw 'Adjustment input must replace the complete adjustment string.'
}
$removeAdjustment = Build-Command 'adjustment' 'remove_adjustment' '/XG1' $true 42
if ((Read-Field $removeAdjustment 'Type') -ne 'remove_text' -or (Read-Field $removeAdjustment 'RemoveText') -ne '/XG1') {
    throw 'Adjustment removal command mismatch.'
}

$unitGuarded = $false
try {
    [void](Build-Command 'quantity' 'multiply' '0.9' $true 0)
} catch {
    $unitGuarded = $_.Exception.ToString().IndexOf('无法识别当前单元', [StringComparison]::Ordinal) -ge 0
}
if (-not $unitGuarded) { throw 'Tree command did not reject a missing current unit.' }

$unwrapped = [string](Invoke-PrivateStatic 'RemoveAgentQuantityFragment' @('(100)*0.9', '*0.9'))
if ($unwrapped -ne '100') { throw "Quantity parentheses were not removed: $unwrapped" }

$nativeMenu = [NativeMenuTestApi]::CreatePopupMenu()
if ($nativeMenu -eq [IntPtr]::Zero) { throw 'Could not create native test menu.' }
try {
    [void][NativeMenuTestApi]::AppendMenu($nativeMenu, 0, [UIntPtr]::op_Explicit(1001), '计算参数设置')
    [void][NativeMenuTestApi]::AppendMenu($nativeMenu, 0, [UIntPtr]::op_Explicit(1002), '删除条目')

    $nativeType = $formType.GetNestedType('NativeTreeMenuFilter', $flags)
    if ($null -eq $nativeType) { throw 'NativeTreeMenuFilter type not found.' }
    $constructor = $nativeType.GetConstructor($flags, $null, [Type[]]@([System.Windows.Forms.Form]), $null)
    if ($null -eq $constructor) { throw 'NativeTreeMenuFilter constructor not found.' }
    $constructorArguments = New-Object object[] 1
    $constructorArguments[0] = $null
    $nativeFilter = $constructor.Invoke($constructorArguments)
    $patchNative = $nativeType.GetMethod('TryPatchNativeTreeMenu', $flags)
    $patchArguments = New-Object object[] 1
    $patchArguments[0] = $nativeMenu
    [void]$patchNative.Invoke($nativeFilter, $patchArguments)

    $rootIndex = Get-NativeItemIndex $nativeMenu '乘系数'
    if ($rootIndex -lt 0) { throw 'Native multiplier root was not inserted.' }
    $nativeRoot = [NativeMenuTestApi]::GetSubMenu($nativeMenu, $rootIndex)
    Assert-NativeBranch $nativeRoot '工程数量' @('乘系数', '删系数')
    Assert-NativeBranch $nativeRoot '定额编号' @('乘系数', '删系数')
    Assert-NativeBranch $nativeRoot '单价' @('乘系数')
    Assert-NativeBranch $nativeRoot '定额调整' @('调整输入', '删除调整')
} finally {
    [void][NativeMenuTestApi]::DestroyMenu($nativeMenu)
}

Write-Host 'PASS: managed/native multiplier menus, commands, factor normalization, and current-unit guard.'
