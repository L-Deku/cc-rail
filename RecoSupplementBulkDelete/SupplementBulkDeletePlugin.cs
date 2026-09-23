using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace RecoSupplementBulkDelete
{
    // 独立插件：给宿主的“补充材料”(RecoNet.补充单价.FormBcCl) 和“补充设备”(RecoNet.补充单价.FormBcSb)
    // 窗口加“批量删除”按钮。宿主方法体受运行时保护，读不到原生删除逻辑，因此：
    //  1. 只用窗体自身的 m_cnn 连接、只删 m_sql 里 from 后面的那张表；
    //  2. 删前扫描同一个库里所有带“电算代号/定额编号”列的表，任何一张表引用了该代号就跳过；
    //  3. 表清单查一次系统视图后缓存，每张表每批只发一条查询，耗时写进日志。
    // 窗口检测不用定时器：2026-09-23 用户反馈装插件后打开料费方案变卡，500ms 定时器是本插件唯一的
    // 常驻动作（每次唤醒消息循环都会连带触发进程里所有 Idle 处理器），改为挂 Application.Idle，
    // 只按索引遍历 Application.OpenForms（不分配、不扫控件树），程序真正空闲时零唤醒。
    public static class SupplementBulkDeletePlugin
    {
        private const int ReferenceChunkSize = 200;
        private const long MaxLogBytes = 5L * 1024L * 1024L;
        private const string ButtonName = "RecoSupplementBulkDelete";
        private const string ButtonText = "批量删除";
        private static readonly object LogLock = new object();
        private static bool idleHooked;
        private static readonly HashSet<Form> InstalledForms = new HashSet<Form>();
        private static readonly Dictionary<string, List<ReferenceColumn>> ReferenceColumnCache =
            new Dictionary<string, List<ReferenceColumn>>(StringComparer.OrdinalIgnoreCase);
        private static readonly Regex FromTableRegex = new Regex(@"\bfrom\s+(\S+)", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        private static readonly Regex IdentifierRegex = new Regex(@"^[\p{L}\p{N}_]+$", RegexOptions.CultureInvariant);

        private sealed class FormTarget
        {
            public string TypeName;
            public string Label;
            public string AnchorButton;
            public long CodeMin;
            public long CodeMax;
        }

        // 补充材料电算代号固定在 4 开头的 9 位号段（状态栏 m_start/m_end 只是当前设计单位的子区间，不能作守卫）。
        // 补充设备号段暂按 9 位整数守卫，安装日志会记录 m_start/m_end 以便日后收紧。
        private static readonly FormTarget[] Targets = new FormTarget[]
        {
            new FormTarget { TypeName = "RecoNet.补充单价.FormBcCl", Label = "补充材料", AnchorButton = "Btn_DoError", CodeMin = 400000001L, CodeMax = 499999999L },
            new FormTarget { TypeName = "RecoNet.补充单价.FormBcSb", Label = "补充设备", AnchorButton = null, CodeMin = 100000000L, CodeMax = 999999999L }
        };

        private sealed class SupplementRow
        {
            public int RowIndex;
            public string CodeText;
            public long Code;
            public string Name;

            public string Display
            {
                get { return CodeText + "  " + Name; }
            }
        }

        private sealed class ReferenceColumn
        {
            public string Table;
            public string Column;
            public bool Numeric;

            public string Display
            {
                get { return Table + "." + Column; }
            }
        }

        private sealed class InstalledContext
        {
            public Form Form;
            public DataGridView Grid;
            public FormTarget Target;
            public string SourceTable;
            public string NameColumn;
        }

        // 由 RecoPluginLoader 在 AppDomain 初始化时调用（主线程、消息循环尚未启动）；
        // 与 RecoExpandPanel 一样直接挂 Application.Idle，不建任何定时器。
        public static void InstallOnIdle()
        {
            if (idleHooked)
            {
                return;
            }

            idleHooked = true;
            Application.Idle += OnIdle;
            Log("InstallOnIdle registered (Application.Idle, no timer).");
        }

        // 每次消息循环空闲时跑一遍：按索引读 Application.OpenForms（通常不到 10 个窗体），
        // 不分配数组、不枚举控件树；只有遇到目标窗体第一次出现时才做安装。
        private static void OnIdle(object sender, EventArgs e)
        {
            try
            {
                ScanOpenForms();
            }
            catch (Exception ex)
            {
                Log("Scan failed: " + ex.Message);
            }
        }

        private static void ScanOpenForms()
        {
            FormCollection openForms = Application.OpenForms;
            int count = openForms.Count;
            for (int i = 0; i < count && i < openForms.Count; i++)
            {
                Form form = openForms[i];
                if (form == null || form.IsDisposed || InstalledForms.Contains(form))
                {
                    continue;
                }

                FormTarget target = FindTarget(form.GetType().FullName);
                if (target != null)
                {
                    Install(form, target);
                }
            }
        }

        private static FormTarget FindTarget(string typeName)
        {
            foreach (FormTarget target in Targets)
            {
                if (String.Equals(target.TypeName, typeName, StringComparison.Ordinal))
                {
                    return target;
                }
            }

            return null;
        }

        private static FormTarget FindTargetByLabel(string label)
        {
            foreach (FormTarget target in Targets)
            {
                if (String.Equals(target.Label, label, StringComparison.Ordinal))
                {
                    return target;
                }
            }

            return null;
        }

        private static bool Install(Form form, FormTarget target)
        {
            if (form == null || target == null)
            {
                return false;
            }

            InstalledForms.Add(form);
            form.FormClosed += delegate { InstalledForms.Remove(form); };

            ToolStrip strip = GetField<ToolStrip>(form, "toolStrip1");
            DataGridView grid = GetField<DataGridView>(form, "dataGridView");
            if (strip == null || grid == null)
            {
                Log(target.Label + ": toolStrip1/dataGridView not found on " + form.GetType().FullName + ", skipped.");
                return false;
            }

            if (strip.Items[ButtonName] != null)
            {
                return true;
            }

            // 单元格选择模式下无法点行头整行选中，改成行头选择模式；单元格编辑不受影响。
            if (grid.SelectionMode == DataGridViewSelectionMode.CellSelect)
            {
                grid.SelectionMode = DataGridViewSelectionMode.RowHeaderSelect;
            }
            grid.MultiSelect = true;

            InstalledContext context = new InstalledContext();
            context.Form = form;
            context.Grid = grid;
            context.Target = target;

            ToolStripButton button = new ToolStripButton();
            button.Name = ButtonName;
            button.Text = ButtonText;
            button.ToolTipText = "删除表格中选中的多条" + target.Label + "（按住 Ctrl 或 Shift 点行头可多选）";
            button.Image = CreateIcon();
            button.ImageScaling = ToolStripItemImageScaling.SizeToFit;
            ToolStripItem anchor = String.IsNullOrEmpty(target.AnchorButton) ? null : GetField<ToolStripItem>(form, target.AnchorButton);
            ToolStripItem styleSource = anchor ?? GetField<ToolStripItem>(form, "tsb_add");
            button.DisplayStyle = styleSource != null ? styleSource.DisplayStyle : ToolStripItemDisplayStyle.ImageAndText;
            button.Click += delegate { Run(context); };

            int index = anchor != null ? strip.Items.IndexOf(anchor) : -1;
            if (index >= 0)
            {
                strip.Items.Insert(index + 1, button);
            }
            else
            {
                strip.Items.Add(button);
            }

            Log(target.Label + ": bulk delete installed on " + form.GetType().FullName + ".");
            LogDiagnostics(context, strip);
            return true;
        }

        private static Image CreateIcon()
        {
            Bitmap bitmap = new Bitmap(16, 16);
            using (Graphics g = Graphics.FromImage(bitmap))
            {
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                g.Clear(Color.Transparent);
                using (Pen pen = new Pen(Color.FromArgb(200, 40, 40), 2.4f))
                {
                    g.DrawLine(pen, 4, 4, 12, 12);
                    g.DrawLine(pen, 12, 4, 4, 12);
                }
            }

            return bitmap;
        }

        private static void LogDiagnostics(InstalledContext context, ToolStrip strip)
        {
            try
            {
                Form form = context.Form;
                DataGridView grid = context.Grid;
                StringBuilder sb = new StringBuilder();
                sb.Append(context.Target.Label).Append(" form: text=").Append(form.Text);
                sb.Append(" modal=").Append(form.Modal);
                sb.Append(" read=").Append(Convert.ToString(GetFieldObject(form, "m_bRead"), CultureInfo.InvariantCulture));
                sb.Append(" isXm=").Append(Convert.ToString(GetFieldObject(form, "m_bIsXm"), CultureInfo.InvariantCulture));
                sb.Append(" single=").Append(Convert.ToString(GetFieldObject(form, "m_IsSingle"), CultureInfo.InvariantCulture));
                sb.Append(" start=").Append(Convert.ToString(GetFieldObject(form, "m_start"), CultureInfo.InvariantCulture));
                sb.Append(" end=").Append(Convert.ToString(GetFieldObject(form, "m_end"), CultureInfo.InvariantCulture));
                sb.Append(" wh=").Append(Convert.ToString(GetFieldObject(form, "m_strWh"), CultureInfo.InvariantCulture));
                sb.Append(" hcode=").Append(Convert.ToString(GetFieldObject(form, "m_strHCode"), CultureInfo.InvariantCulture));
                SqlConnection conn = GetField<SqlConnection>(form, "m_cnn");
                sb.Append(" db=").Append(conn == null ? "<null>" : conn.Database + "/" + conn.State);
                sb.Append(" sql=").Append(Convert.ToString(GetFieldObject(form, "m_sql"), CultureInfo.InvariantCulture));
                Log(sb.ToString());

                sb.Length = 0;
                sb.Append(context.Target.Label).Append(" toolstrip:");
                foreach (ToolStripItem item in strip.Items)
                {
                    sb.Append(" [").Append(item.Name).Append('=').Append(item.Text)
                        .Append(item.Visible ? "" : ",hidden").Append(item.Enabled ? "" : ",disabled").Append(']');
                }
                Log(sb.ToString());

                sb.Length = 0;
                sb.Append(context.Target.Label).Append(" grid: mode=").Append(grid.SelectionMode)
                    .Append(" multi=").Append(grid.MultiSelect)
                    .Append(" rows=").Append(grid.Rows.Count)
                    .Append(" source=").Append(DescribeSource(grid.DataSource))
                    .Append(" columns:");
                foreach (DataGridViewColumn column in grid.Columns)
                {
                    sb.Append(" [").Append(column.Name).Append('|').Append(column.HeaderText).Append('|').Append(column.DataPropertyName)
                        .Append(column.Visible ? "" : "|hidden").Append(']');
                }
                Log(sb.ToString());
            }
            catch (Exception ex)
            {
                Log("Diagnostics failed: " + ex.Message);
            }
        }

        private static string DescribeSource(object source)
        {
            if (source == null)
            {
                return "<null>";
            }

            DataTable table = ResolveTable(source);
            if (table == null)
            {
                return source.GetType().Name;
            }

            StringBuilder sb = new StringBuilder();
            sb.Append(source.GetType().Name).Append(':').Append(table.TableName).Append('(');
            foreach (DataColumn column in table.Columns)
            {
                sb.Append(column.ColumnName).Append(',');
            }
            sb.Append(')');
            return sb.ToString();
        }

        private static DataTable ResolveTable(object source)
        {
            DataTable table = source as DataTable;
            if (table != null)
            {
                return table;
            }

            DataView view = source as DataView;
            if (view != null)
            {
                return view.Table;
            }

            BindingSource binding = source as BindingSource;
            if (binding != null)
            {
                DataTable inner = ResolveTable(binding.DataSource);
                if (inner != null)
                {
                    return inner;
                }

                DataView listView = binding.List as DataView;
                return listView != null ? listView.Table : null;
            }

            return null;
        }

        // ---- 行收集与过滤 ----

        private static string ResolveNameColumn(DataGridView grid)
        {
            DataTable table = ResolveTable(grid.DataSource);
            if (table != null)
            {
                foreach (DataColumn column in table.Columns)
                {
                    if (column.ColumnName.EndsWith("名称", StringComparison.Ordinal))
                    {
                        return column.ColumnName;
                    }
                }
            }

            foreach (DataGridViewColumn column in grid.Columns)
            {
                if (!String.IsNullOrEmpty(column.DataPropertyName) && column.DataPropertyName.EndsWith("名称", StringComparison.Ordinal))
                {
                    return column.DataPropertyName;
                }
                if (column.HeaderText.EndsWith("名称", StringComparison.Ordinal))
                {
                    return column.HeaderText;
                }
            }

            return "";
        }

        private static string GetCellText(DataGridViewRow row, string columnName)
        {
            if (row == null || String.IsNullOrEmpty(columnName))
            {
                return "";
            }

            DataRowView rowView = row.DataBoundItem as DataRowView;
            if (rowView != null && rowView.DataView.Table.Columns.Contains(columnName))
            {
                object value = rowView[columnName];
                return value == null || value == DBNull.Value ? "" : Convert.ToString(value, CultureInfo.InvariantCulture).Trim();
            }

            if (row.DataGridView != null)
            {
                foreach (DataGridViewColumn column in row.DataGridView.Columns)
                {
                    if (String.Equals(column.DataPropertyName, columnName, StringComparison.Ordinal) ||
                        String.Equals(column.HeaderText, columnName, StringComparison.Ordinal) ||
                        String.Equals(column.Name, columnName, StringComparison.Ordinal))
                    {
                        object value = row.Cells[column.Index].Value;
                        if (value != null && value != DBNull.Value)
                        {
                            return Convert.ToString(value, CultureInfo.InvariantCulture).Trim();
                        }
                    }
                }
            }

            return "";
        }

        private static List<SupplementRow> CollectSelectedRows(DataGridView grid, string codeColumn, string nameColumn)
        {
            List<SupplementRow> rows = new List<SupplementRow>();
            if (grid == null)
            {
                return rows;
            }

            SortedDictionary<int, bool> indices = new SortedDictionary<int, bool>();
            foreach (DataGridViewRow row in grid.SelectedRows)
            {
                indices[row.Index] = true;
            }
            foreach (DataGridViewCell cell in grid.SelectedCells)
            {
                if (cell.RowIndex >= 0)
                {
                    indices[cell.RowIndex] = true;
                }
            }

            foreach (int index in indices.Keys)
            {
                if (index < 0 || index >= grid.Rows.Count || grid.Rows[index].IsNewRow)
                {
                    continue;
                }

                DataGridViewRow row = grid.Rows[index];
                SupplementRow item = new SupplementRow();
                item.RowIndex = index;
                item.CodeText = GetCellText(row, codeColumn);
                item.Name = GetCellText(row, nameColumn);
                long code;
                item.Code = Int64.TryParse(item.CodeText, NumberStyles.Integer, CultureInfo.InvariantCulture, out code) ? code : 0L;
                rows.Add(item);
            }

            return rows;
        }

        private static List<SupplementRow> FilterRows(List<SupplementRow> rows, long min, long max, List<string> skipped)
        {
            List<SupplementRow> valid = new List<SupplementRow>();
            foreach (SupplementRow row in rows)
            {
                if (String.IsNullOrEmpty(row.CodeText))
                {
                    skipped.Add("第 " + (row.RowIndex + 1).ToString(CultureInfo.InvariantCulture) + " 行：电算代号为空，未删除");
                    continue;
                }

                if (row.Code <= 0)
                {
                    skipped.Add(row.Display + "：电算代号不是整数，未删除");
                    continue;
                }

                if (row.Code < min || row.Code > max)
                {
                    skipped.Add(row.Display + "：不在补充号段 " + min.ToString(CultureInfo.InvariantCulture) + "---" + max.ToString(CultureInfo.InvariantCulture) + " 内，未删除");
                    continue;
                }

                valid.Add(row);
            }

            return valid;
        }

        private static string ParseSourceTable(string hostSql)
        {
            if (String.IsNullOrEmpty(hostSql))
            {
                return "";
            }

            Match match = FromTableRegex.Match(hostSql);
            if (!match.Success)
            {
                return "";
            }

            // 去掉方括号和结尾的逗号/分号，再取最后一段（dbo.表名 → 表名）；子查询“(select ...”会因含括号而不通过标识符校验。
            string table = match.Groups[1].Value.Trim().TrimEnd(',', ';').Replace("[", "").Replace("]", "");
            int dot = table.LastIndexOf('.');
            if (dot >= 0)
            {
                table = table.Substring(dot + 1);
            }

            return IdentifierRegex.IsMatch(table) ? table : "";
        }

        // ---- 删除主流程 ----

        private static void Run(InstalledContext context)
        {
            Form form = context.Form;
            DataGridView grid = context.Grid;
            FormTarget target = context.Target;
            string title = "批量删除" + target.Label;
            try
            {
                if (GetFieldBool(form, "m_bRead"))
                {
                    MessageBox.Show(form, "当前" + target.Label + "窗口是只读状态，不能删除。", title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                ToolStripItem addButton = GetField<ToolStripItem>(form, "tsb_add");
                if (addButton != null && (!addButton.Enabled || !addButton.Visible))
                {
                    MessageBox.Show(form, "当前窗口的“添加”按钮不可用（无编辑权限或已被锁定），不能批量删除。", title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string hostSql = Convert.ToString(GetFieldObject(form, "m_sql"), CultureInfo.InvariantCulture) ?? "";
                string sourceTable = ParseSourceTable(hostSql);
                if (String.IsNullOrEmpty(sourceTable))
                {
                    Log(target.Label + ": cannot resolve source table from host sql=" + hostSql);
                    MessageBox.Show(form, "无法从当前窗口确定数据表，为安全起见不执行批量删除。", title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                context.SourceTable = sourceTable;

                string nameColumn = ResolveNameColumn(grid);
                if (String.IsNullOrEmpty(nameColumn))
                {
                    Log(target.Label + ": name column not found.");
                    MessageBox.Show(form, "无法确定名称列，为安全起见不执行批量删除。", title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                context.NameColumn = nameColumn;

                grid.EndEdit();
                List<SupplementRow> selected = CollectSelectedRows(grid, "电算代号", nameColumn);
                if (selected.Count == 0)
                {
                    MessageBox.Show(form, "请先在表格中选中要删除的" + target.Label + "行（按住 Ctrl 或 Shift 点行头可多选）。", title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                List<string> skipped = new List<string>();
                List<SupplementRow> valid = FilterRows(selected, target.CodeMin, target.CodeMax, skipped);
                if (valid.Count == 0)
                {
                    MessageBox.Show(form, "选中的行都不能删除：" + Environment.NewLine + String.Join(Environment.NewLine, skipped.ToArray()), title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                StringBuilder confirm = new StringBuilder();
                confirm.Append("将从表 ").Append(sourceTable).Append(" 中删除以下 ").Append(valid.Count.ToString(CultureInfo.InvariantCulture)).Append(" 条").Append(target.Label).Append("，删除后不可恢复：").AppendLine().AppendLine();
                for (int i = 0; i < valid.Count && i < 20; i++)
                {
                    confirm.AppendLine(valid[i].Display);
                }
                if (valid.Count > 20)
                {
                    confirm.AppendLine("……还有 " + (valid.Count - 20).ToString(CultureInfo.InvariantCulture) + " 条");
                }
                if (skipped.Count > 0)
                {
                    confirm.AppendLine().AppendLine("另有 " + skipped.Count.ToString(CultureInfo.InvariantCulture) + " 行不符合条件，将跳过。");
                }
                confirm.AppendLine().Append("已被本库任何表引用的代号会自动跳过。是否继续？");
                DialogResult answer = MessageBox.Show(form, confirm.ToString(), title, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (answer != DialogResult.Yes)
                {
                    return;
                }

                SqlConnection conn = GetField<SqlConnection>(form, "m_cnn");
                if (conn == null)
                {
                    MessageBox.Show(form, "未找到窗口的数据库连接，无法删除。", title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                List<SupplementRow> deletedRows = new List<SupplementRow>();
                int deletedCount = DeleteRows(conn, context, valid, skipped, deletedRows);
                Log(target.Label + ": db=" + conn.Database + " table=" + sourceTable +
                    " requested=" + valid.Count.ToString(CultureInfo.InvariantCulture) +
                    " deleted=" + deletedCount.ToString(CultureInfo.InvariantCulture) +
                    " skipped=" + skipped.Count.ToString(CultureInfo.InvariantCulture) +
                    " codes=" + String.Join(",", deletedRows.ConvertAll(delegate(SupplementRow r) { return r.CodeText; }).ToArray()));

                RefreshGrid(context, deletedRows);

                StringBuilder summary = new StringBuilder();
                summary.Append("已删除 ").Append(deletedCount.ToString(CultureInfo.InvariantCulture)).Append(" 条").Append(target.Label).Append("。");
                if (skipped.Count > 0)
                {
                    summary.AppendLine().AppendLine().Append("以下 ").Append(skipped.Count.ToString(CultureInfo.InvariantCulture)).Append(" 行未删除：");
                    for (int i = 0; i < skipped.Count && i < 30; i++)
                    {
                        summary.AppendLine().Append(skipped[i]);
                    }
                    if (skipped.Count > 30)
                    {
                        summary.AppendLine().Append("……还有 " + (skipped.Count - 30).ToString(CultureInfo.InvariantCulture) + " 行");
                    }
                }
                MessageBox.Show(form, summary.ToString(), title, MessageBoxButtons.OK, skipped.Count > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log(target.Label + ": bulk delete failed: " + ex);
                MessageBox.Show(form, "批量删除失败：" + ex.Message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static int DeleteRows(SqlConnection conn, InstalledContext context, List<SupplementRow> rows, List<string> skipped, List<SupplementRow> deletedRows)
        {
            EnsureOpen(conn);
            int deleted = 0;
            using (SqlTransaction tx = conn.BeginTransaction())
            {
                Dictionary<string, List<string>> references = FindReferences(conn, tx, context.SourceTable, rows);
                foreach (SupplementRow row in rows)
                {
                    List<string> usedBy;
                    if (references.TryGetValue(row.CodeText, out usedBy) && usedBy.Count > 0)
                    {
                        skipped.Add(row.Display + "：已被 " + String.Join("、", usedBy.ToArray()) + " 引用，未删除");
                        continue;
                    }

                    string nameColumn = context.NameColumn;
                    int affected = ExecuteNonQuery(conn, tx,
                        "delete from [" + context.SourceTable + "] where [电算代号]=@code and ltrim(rtrim(isnull([" + nameColumn + "],'')))=@name", row.Code, row.Name);
                    if (affected == 0)
                    {
                        int sameCode = ExecuteScalar(conn, tx, "select count(*) from [" + context.SourceTable + "] where [电算代号]=@code", row.Code, null);
                        if (sameCode == 1)
                        {
                            affected = ExecuteNonQuery(conn, tx, "delete from [" + context.SourceTable + "] where [电算代号]=@code", row.Code, null);
                        }
                        else if (sameCode > 1)
                        {
                            skipped.Add(row.Display + "：数据库中同一电算代号有 " + sameCode.ToString(CultureInfo.InvariantCulture) + " 条且名称不一致，未删除");
                            continue;
                        }
                    }

                    if (affected == 0)
                    {
                        skipped.Add(row.Display + "：数据库中未找到该记录，未删除");
                        continue;
                    }

                    deleted += affected;
                    deletedRows.Add(row);
                }

                tx.Commit();
            }

            return deleted;
        }

        // ---- 引用扫描 ----

        private static List<ReferenceColumn> GetReferenceColumns(SqlConnection conn, SqlTransaction tx, string sourceTable)
        {
            string key = conn.DataSource + "|" + conn.Database + "|" + sourceTable;
            List<ReferenceColumn> cached;
            if (ReferenceColumnCache.TryGetValue(key, out cached))
            {
                return cached;
            }

            List<ReferenceColumn> columns = new List<ReferenceColumn>();
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText =
                    "select c.table_name, c.column_name, c.data_type " +
                    "from information_schema.columns c " +
                    "join information_schema.tables t on t.table_schema=c.table_schema and t.table_name=c.table_name " +
                    "where t.table_type='BASE TABLE' and c.column_name in (N'电算代号', N'定额编号') and c.table_name<>@source " +
                    "order by c.table_name, c.column_name";
                cmd.Parameters.Add("@source", SqlDbType.NVarChar, 200).Value = sourceTable;
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string dataType = Convert.ToString(reader[2], CultureInfo.InvariantCulture).ToLowerInvariant();
                        bool numeric = dataType == "int" || dataType == "bigint" || dataType == "smallint" || dataType == "tinyint" ||
                                       dataType == "numeric" || dataType == "decimal" || dataType == "float" || dataType == "real" || dataType == "money";
                        bool text = dataType == "char" || dataType == "nchar" || dataType == "varchar" || dataType == "nvarchar";
                        if (!numeric && !text)
                        {
                            continue;
                        }

                        ReferenceColumn column = new ReferenceColumn();
                        column.Table = Convert.ToString(reader[0], CultureInfo.InvariantCulture);
                        column.Column = Convert.ToString(reader[1], CultureInfo.InvariantCulture);
                        column.Numeric = numeric;
                        columns.Add(column);
                    }
                }
            }

            ReferenceColumnCache[key] = columns;
            StringBuilder sb = new StringBuilder();
            foreach (ReferenceColumn column in columns)
            {
                sb.Append(' ').Append(column.Display).Append(column.Numeric ? "(n)" : "(s)");
            }
            Log("Reference columns in " + conn.Database + " (excluding " + sourceTable + "): count=" + columns.Count.ToString(CultureInfo.InvariantCulture) + sb);
            return columns;
        }

        // 每张表每批只发一条 in(...) 查询，返回被引用的代号；耗时记日志。
        private static Dictionary<string, List<string>> FindReferences(SqlConnection conn, SqlTransaction tx, string sourceTable, List<SupplementRow> rows)
        {
            Dictionary<string, List<string>> result = new Dictionary<string, List<string>>(StringComparer.Ordinal);
            Stopwatch watch = Stopwatch.StartNew();
            List<ReferenceColumn> columns = GetReferenceColumns(conn, tx, sourceTable);
            int queries = 0;
            foreach (ReferenceColumn column in columns)
            {
                for (int offset = 0; offset < rows.Count; offset += ReferenceChunkSize)
                {
                    int count = Math.Min(ReferenceChunkSize, rows.Count - offset);
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = tx;
                        StringBuilder sql = new StringBuilder();
                        string expr = column.Numeric ? "[" + column.Column + "]" : "ltrim(rtrim([" + column.Column + "]))";
                        sql.Append("select distinct ").Append(expr).Append(" from [").Append(column.Table).Append("] where ").Append(expr).Append(" in (");
                        for (int i = 0; i < count; i++)
                        {
                            if (i > 0)
                            {
                                sql.Append(',');
                            }
                            string name = "@p" + i.ToString(CultureInfo.InvariantCulture);
                            sql.Append(name);
                            SupplementRow row = rows[offset + i];
                            if (column.Numeric)
                            {
                                cmd.Parameters.Add(name, SqlDbType.BigInt).Value = row.Code;
                            }
                            else
                            {
                                cmd.Parameters.Add(name, SqlDbType.NVarChar, 100).Value = row.CodeText;
                            }
                        }
                        sql.Append(')');
                        cmd.CommandText = sql.ToString();
                        queries++;
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                object value = reader[0];
                                if (value == null || value == DBNull.Value)
                                {
                                    continue;
                                }

                                string codeText = column.Numeric
                                    ? Convert.ToInt64(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture)
                                    : Convert.ToString(value, CultureInfo.InvariantCulture).Trim();
                                List<string> usedBy;
                                if (!result.TryGetValue(codeText, out usedBy))
                                {
                                    usedBy = new List<string>();
                                    result[codeText] = usedBy;
                                }
                                usedBy.Add(column.Display);
                            }
                        }
                    }
                }
            }

            watch.Stop();
            Log("Reference scan: tables=" + columns.Count.ToString(CultureInfo.InvariantCulture) +
                " queries=" + queries.ToString(CultureInfo.InvariantCulture) +
                " codes=" + rows.Count.ToString(CultureInfo.InvariantCulture) +
                " referenced=" + result.Count.ToString(CultureInfo.InvariantCulture) +
                " ms=" + watch.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture));
            return result;
        }

        // ---- SQL 辅助 ----

        private static int ExecuteScalar(SqlConnection conn, SqlTransaction tx, string sql, long code, string name)
        {
            using (SqlCommand cmd = BuildCommand(conn, tx, sql, code, name))
            {
                object value = cmd.ExecuteScalar();
                return value == null || value == DBNull.Value ? 0 : Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
        }

        private static int ExecuteNonQuery(SqlConnection conn, SqlTransaction tx, string sql, long code, string name)
        {
            using (SqlCommand cmd = BuildCommand(conn, tx, sql, code, name))
            {
                return cmd.ExecuteNonQuery();
            }
        }

        private static SqlCommand BuildCommand(SqlConnection conn, SqlTransaction tx, string sql, long code, string name)
        {
            SqlCommand cmd = conn.CreateCommand();
            cmd.Transaction = tx;
            cmd.CommandText = sql;
            cmd.Parameters.Add("@code", SqlDbType.BigInt).Value = code;
            if (name != null)
            {
                cmd.Parameters.Add("@name", SqlDbType.NVarChar, 400).Value = name;
            }
            return cmd;
        }

        private static void EnsureOpen(SqlConnection conn)
        {
            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
        }

        // ---- 删除后刷新 ----

        // 先在窗体已绑定的 DataTable 里移除已删的行（不触发宿主未知逻辑）；
        // 只有本地移除不完整时才回退调用宿主自己的下拉刷新重新查库。
        private static void RefreshGrid(InstalledContext context, List<SupplementRow> deletedRows)
        {
            if (deletedRows.Count == 0)
            {
                return;
            }

            DataGridView grid = context.Grid;
            try
            {
                int removed = RemoveRowsLocally(grid, context.NameColumn, deletedRows);
                Log(context.Target.Label + ": locally removed " + removed.ToString(CultureInfo.InvariantCulture) + "/" + deletedRows.Count.ToString(CultureInfo.InvariantCulture));
                if (removed >= deletedRows.Count)
                {
                    grid.Refresh();
                    return;
                }
            }
            catch (Exception ex)
            {
                Log(context.Target.Label + ": local row removal failed: " + ex.Message);
            }

            try
            {
                MethodInfo refresh = context.Form.GetType().GetMethod("comboBox_SelectedIndexChanged", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                if (refresh != null)
                {
                    refresh.Invoke(context.Form, new object[] { GetFieldObject(context.Form, "comboBox"), EventArgs.Empty });
                    Log(context.Target.Label + ": grid refreshed via host comboBox_SelectedIndexChanged.");
                }
            }
            catch (Exception ex)
            {
                Log(context.Target.Label + ": host refresh failed: " + ex.Message);
            }
        }

        private static int RemoveRowsLocally(DataGridView grid, string nameColumn, List<SupplementRow> deletedRows)
        {
            HashSet<string> keys = new HashSet<string>(StringComparer.Ordinal);
            foreach (SupplementRow row in deletedRows)
            {
                keys.Add(row.CodeText + "|" + row.Name);
            }

            DataTable table = ResolveTable(grid.DataSource);
            int removed = 0;
            if (table != null)
            {
                grid.ClearSelection();
                List<DataRow> targets = new List<DataRow>();
                foreach (DataRow dataRow in table.Rows)
                {
                    if (dataRow.RowState == DataRowState.Deleted || dataRow.RowState == DataRowState.Detached)
                    {
                        continue;
                    }

                    string code = table.Columns.Contains("电算代号") ? Convert.ToString(dataRow["电算代号"], CultureInfo.InvariantCulture).Trim() : "";
                    string name = table.Columns.Contains(nameColumn) ? Convert.ToString(dataRow[nameColumn], CultureInfo.InvariantCulture).Trim() : "";
                    if (keys.Contains(code + "|" + name))
                    {
                        targets.Add(dataRow);
                    }
                }

                foreach (DataRow target in targets)
                {
                    table.Rows.Remove(target);
                    removed++;
                }

                return removed;
            }

            List<int> indices = new List<int>();
            foreach (SupplementRow row in deletedRows)
            {
                indices.Add(row.RowIndex);
            }
            indices.Sort();
            for (int i = indices.Count - 1; i >= 0; i--)
            {
                int index = indices[i];
                if (index >= 0 && index < grid.Rows.Count && !grid.Rows[index].IsNewRow)
                {
                    grid.Rows.RemoveAt(index);
                    removed++;
                }
            }

            return removed;
        }

        // ---- 反射与日志 ----

        private static T GetField<T>(object target, string name) where T : class
        {
            return GetFieldObject(target, name) as T;
        }

        private static object GetFieldObject(object target, string name)
        {
            if (target == null || String.IsNullOrEmpty(name))
            {
                return null;
            }

            FieldInfo field = target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            return field == null ? null : field.GetValue(target);
        }

        private static bool GetFieldBool(object target, string name)
        {
            object value = GetFieldObject(target, name);
            return value is bool && (bool)value;
        }

        public static void Log(string message)
        {
            try
            {
                string dir = Path.GetDirectoryName(typeof(SupplementBulkDeletePlugin).Assembly.Location);
                string path = Path.Combine(dir, "RecoSupplementBulkDelete.log");
                string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " " + message + Environment.NewLine;
                lock (LogLock)
                {
                    if (File.Exists(path) && new FileInfo(path).Length >= MaxLogBytes)
                    {
                        string backup = path + ".1";
                        if (File.Exists(backup))
                        {
                            File.Delete(backup);
                        }
                        File.Move(path, backup);
                    }
                    File.AppendAllText(path, line, Encoding.UTF8);
                }
            }
            catch
            {
            }
        }
    }
}
