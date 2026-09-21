$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

# 补充材料/补充设备批量删除：用字段名与宿主 FormBcCl/FormBcSb 相同的假窗体做反射冒烟验证，
# 覆盖按钮安装、选中行收集、号段过滤、m_sql 表名解析和本地移除，不连数据库。

$dll = if (-not [String]::IsNullOrWhiteSpace($env:RECO_BULK_DELETE_DLL)) {
    $env:RECO_BULK_DELETE_DLL
} else {
    Join-Path (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '..\..')).Path 'RecoQuotaRecommend\bin\RecoSupplementBulkDelete.dll'
}

Add-Type -Path $dll
Add-Type -ReferencedAssemblies 'System.Windows.Forms', 'System.Data', 'System.Drawing', 'System.Xml' -TypeDefinition @'
using System.Data;
using System.Windows.Forms;

public class FakeBcForm : Form
{
    public ToolStrip toolStrip1 = new ToolStrip();
    public ToolStripButton tsb_add = new ToolStripButton("添加");
    public ToolStripButton Btn_DoError = new ToolStripButton("异常处理");
    public ToolStripButton tsb_out = new ToolStripButton("导出");
    public ToolStripComboBox comboBox = new ToolStripComboBox();
    public DataGridView dataGridView = new DataGridView();
    public bool m_bRead;
    public long m_start = 400000001L;
    public long m_end = 400009999L;
    public string m_sql = " select lock,文号,电算代号,材料名称 from 材料单价库 ";
    public System.Data.SqlClient.SqlConnection m_cnn;
    public DataTable Table;

    public FakeBcForm(string nameColumn)
    {
        toolStrip1.Items.Add(tsb_add);
        toolStrip1.Items.Add(tsb_out);
        toolStrip1.Items.Add(Btn_DoError);
        Controls.Add(toolStrip1);
        Controls.Add(dataGridView);
        dataGridView.SelectionMode = DataGridViewSelectionMode.CellSelect;
        dataGridView.MultiSelect = false;
        Table = new DataTable("bc");
        Table.Columns.Add("电算代号", typeof(long));
        Table.Columns.Add(nameColumn, typeof(string));
        Table.Rows.Add(400040033L, "阻容盒");
        Table.Rows.Add(400040034L, "钽电容 ");
        Table.Rows.Add(123L, "系统条目");
        dataGridView.DataSource = Table;
    }
}
'@

$pluginType = [RecoSupplementBulkDelete.SupplementBulkDeletePlugin]
$flags = [System.Reflection.BindingFlags]'NonPublic,Static,Instance,Public'

function Invoke-PrivateStatic([string]$name, [object[]]$arguments) {
    $method = $pluginType.GetMethod($name, $flags)
    if ($null -eq $method) {
        throw "Method not found: $name"
    }
    $parameters = $method.GetParameters()
    $invokeArguments = New-Object object[] $arguments.Count
    for ($i = 0; $i -lt $arguments.Count; $i++) {
        # 不要用 if/else 语句给变量赋值：语句输出会把集合展开（空 List 会变成 $null，泛型 List 会变成 Object[]）。
        $value = $arguments[$i]
        if ($null -ne $value) { $value = $value.PSObject.BaseObject }
        # PowerShell 会把函数返回的 List 展开成 Object[]，这里按目标参数类型重新装回泛型 List。
        $parameterType = $parameters[$i].ParameterType
        if ($null -ne $value -and $value -is [System.Array] -and $parameterType.IsGenericType -and
            $parameterType.GetGenericTypeDefinition() -eq [System.Collections.Generic.List`1]) {
            $typed = [Activator]::CreateInstance($parameterType)
            foreach ($item in $value) { [void]$typed.Add($item) }
            $value = $typed
        }
        $invokeArguments[$i] = $value
    }
    return $method.Invoke($null, $invokeArguments)
}

function Assert-True([bool]$condition, [string]$message) {
    if (-not $condition) {
        throw "ASSERT FAILED: $message"
    }
}

# ---- 表名解析 ----
Assert-True ((Invoke-PrivateStatic 'ParseSourceTable' @(' select a,b from 材料单价库 where x=1 ')) -eq '材料单价库') 'Should parse 材料单价库 from host sql.'
Assert-True ((Invoke-PrivateStatic 'ParseSourceTable' @('select * from [dbo].[设备单价库] order by 1')) -eq '设备单价库') 'Should parse bracketed schema-qualified table.'
Assert-True ((Invoke-PrivateStatic 'ParseSourceTable' @('select 1')) -eq '') 'No from clause should yield empty.'
Assert-True ((Invoke-PrivateStatic 'ParseSourceTable' @('select * from (select 1) x')) -eq '') 'Subquery source should be rejected.'

# ---- 补充材料窗体 ----
$materialTarget = Invoke-PrivateStatic 'FindTargetByLabel' @('补充材料')
Assert-True ($null -ne $materialTarget) 'Material target should exist.'
$form = New-Object FakeBcForm '材料名称'
[void]$form.Handle
$grid = $form.dataGridView
Assert-True ($grid.Rows.Count -ge 3) "Fake grid should have bound rows, got $($grid.Rows.Count)."

$installed = Invoke-PrivateStatic 'Install' @($form, $materialTarget)
Assert-True ([bool]$installed) 'Install should return true.'
$button = $form.toolStrip1.Items['RecoSupplementBulkDelete']
Assert-True ($null -ne $button) 'Bulk delete button should be added to toolStrip1.'
Assert-True ($button.Text -eq '批量删除') "Button text should be 批量删除, got '$($button.Text)'."
Assert-True ($form.toolStrip1.Items.IndexOf($button) -eq ($form.toolStrip1.Items.IndexOf($form.Btn_DoError) + 1)) 'Material button should sit right after 异常处理.'
Assert-True ($grid.MultiSelect) 'Grid MultiSelect should be enabled.'
Assert-True ($grid.SelectionMode -eq [System.Windows.Forms.DataGridViewSelectionMode]::RowHeaderSelect) 'CellSelect should be switched to RowHeaderSelect.'
[void](Invoke-PrivateStatic 'Install' @($form, $materialTarget))
Assert-True (@($form.toolStrip1.Items | Where-Object { $_.Name -eq 'RecoSupplementBulkDelete' }).Count -eq 1) 'Re-install must not duplicate the button.'

Assert-True ((Invoke-PrivateStatic 'ResolveNameColumn' @($grid)) -eq '材料名称') 'Name column should resolve to 材料名称.'

$grid.ClearSelection()
$grid.Rows[0].Selected = $true
$grid.Rows[2].Cells[1].Selected = $true
$selected = @(Invoke-PrivateStatic 'CollectSelectedRows' @($grid, '电算代号', '材料名称'))
Assert-True ($selected.Count -eq 2) "Expected 2 selected rows, got $($selected.Count)."
$rowType = $selected[0].GetType()
$codeField = $rowType.GetField('CodeText', $flags)
$nameField = $rowType.GetField('Name', $flags)
Assert-True ($codeField.GetValue($selected[0]) -eq '400040033') "First selected code should be 400040033, got $($codeField.GetValue($selected[0]))."
Assert-True ($nameField.GetValue($selected[0]) -eq '阻容盒') 'First selected name should be 阻容盒.'
Assert-True ($codeField.GetValue($selected[1]) -eq '123') "Second selected code should be 123, got $($codeField.GetValue($selected[1]))."

$skipped = New-Object 'System.Collections.Generic.List[string]'
$valid = @(Invoke-PrivateStatic 'FilterRows' @($selected, [long]400000001, [long]499999999, $skipped))
Assert-True ($valid.Count -eq 1) "Expected 1 valid row, got $($valid.Count)."
Assert-True ($codeField.GetValue($valid[0]) -eq '400040033') 'Valid row should be 400040033.'
Assert-True ($skipped.Count -eq 1 -and $skipped[0].Contains('不在补充号段')) "Out-of-range row should be skipped with reason, got: $($skipped -join '; ')"

$removed = Invoke-PrivateStatic 'RemoveRowsLocally' @($grid, '材料名称', $valid)
Assert-True ($removed -eq 1) "Expected 1 row removed locally, got $removed."
Assert-True ($form.Table.Rows.Count -eq 2) "Table should have 2 rows left, got $($form.Table.Rows.Count)."
Assert-True (@($form.Table.Rows | Where-Object { [long]$_['电算代号'] -eq 400040033 }).Count -eq 0) 'Deleted code must be gone from the bound table.'
$form.Dispose()

# ---- 补充设备窗体：没有 Btn_DoError 锚点，按钮追加到末尾，名称列为 设备名称 ----
$equipmentTarget = Invoke-PrivateStatic 'FindTargetByLabel' @('补充设备')
Assert-True ($null -ne $equipmentTarget) 'Equipment target should exist.'
$sbForm = New-Object FakeBcForm '设备名称'
$sbForm.m_sql = 'select 电算代号,设备名称 from 设备单价库'
[void]$sbForm.Handle
Assert-True ([bool](Invoke-PrivateStatic 'Install' @($sbForm, $equipmentTarget))) 'Equipment install should return true.'
$sbButton = $sbForm.toolStrip1.Items['RecoSupplementBulkDelete']
Assert-True ($null -ne $sbButton) 'Equipment bulk delete button should exist.'
Assert-True ($sbForm.toolStrip1.Items.IndexOf($sbButton) -eq ($sbForm.toolStrip1.Items.Count - 1)) 'Equipment button should be appended at the end.'
Assert-True ((Invoke-PrivateStatic 'ResolveNameColumn' @($sbForm.dataGridView)) -eq '设备名称') 'Equipment name column should resolve to 设备名称.'
Assert-True ((Invoke-PrivateStatic 'ParseSourceTable' @($sbForm.m_sql)) -eq '设备单价库') 'Equipment source table should parse.'
$sbForm.Dispose()

Write-Host 'Test-SupplementBulkDelete: PASS'
