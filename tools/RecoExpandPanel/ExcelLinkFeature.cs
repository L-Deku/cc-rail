using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.IO.Compression;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        private static readonly Dictionary<Form, ExcelLinkRuntime> ExcelLinkRuntimes = new Dictionary<Form, ExcelLinkRuntime>();
        private static readonly Dictionary<Form, ExcelLinkPanel> ExcelLinkPanels = new Dictionary<Form, ExcelLinkPanel>();
        private static readonly Dictionary<Form, ExcelSmartBindPanel> ExcelSmartBindPanels = new Dictionary<Form, ExcelSmartBindPanel>();
        private static readonly Dictionary<Form, QuickBindPanel> QuickBindPanels = new Dictionary<Form, QuickBindPanel>();
        private static readonly HashSet<DataGridView> HookedQuotaGrids = new HashSet<DataGridView>();

        private static void EnsureExcelLinkRuntime(Form mainForm)
        {
            if (mainForm == null || ExcelLinkRuntimes.ContainsKey(mainForm))
            {
                return;
            }

            ExcelLinkRuntime runtime = new ExcelLinkRuntime(mainForm);
            ExcelLinkRuntimes[mainForm] = runtime;
            mainForm.FormClosed += delegate
            {
                runtime.Dispose();
                ExcelLinkRuntimes.Remove(mainForm);
                if (ExcelLinkPanels.ContainsKey(mainForm))
                {
                    ExcelLinkPanels.Remove(mainForm);
                }
                if (QuickBindPanels.ContainsKey(mainForm))
                {
                    QuickBindPanels.Remove(mainForm);
                }
                if (ExcelSmartBindPanels.ContainsKey(mainForm))
                {
                    ExcelSmartBindPanels.Remove(mainForm);
                }
            };
        }

        private static void InstallQuotaGridShortcuts(Form mainForm)
        {
            DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
            if (grid == null || HookedQuotaGrids.Contains(grid))
            {
                return;
            }

            HookedQuotaGrids.Add(grid);
            grid.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.Control && e.Shift && e.KeyCode == Keys.E)
                {
                    ShowQuickBindPanel(mainForm);
                    e.Handled = true;
                }
                else if (e.Control && e.KeyCode == Keys.E)
                {
                    SmartBindSelectedQuotasToExcel(mainForm);
                    e.Handled = true;
                }
            };
            Log("Excel quick bind shortcut installed.");
        }


        private static void AddExcelLinkItemsIfMatched(ContextMenuStrip menu)
        {
            if (menu == null || !MenuInfos.ContainsKey(menu))
            {
                return;
            }

            MenuInfo info = MenuInfos[menu];
            bool isDeMenu = info.Name == "contextMenuStripDE" || IsSource(menu, info.MainForm, "dataGridViewDE") || HasAnyItem(menu, "定额输入", "定额调整", "单价分析", "全选(A)");
            if (!isDeMenu)
            {
                return;
            }

            AddExcelLinkItems(menu, info.MainForm);
        }

        private static void AddExcelLinkItems(ToolStrip menu, Form mainForm)
        {
            int baseIndex;
            ToolStripMenuItem multiply = FindMenuItem(menu, "乘系数");
            if (multiply != null)
            {
                baseIndex = menu.Items.IndexOf(multiply) + 1;
            }
            else
            {
                baseIndex = FirstVisibleIndex(menu);
                if (baseIndex < 0)
                {
                    baseIndex = menu.Items.Count;
                }
                else
                {
                    baseIndex++;
                }
            }
                RemoveMenuItem(menu, "批量绑定Excel工程量");
                RemoveMenuItem(menu, "打开Excel快速绑定窗口");

                ToolStripMenuItem bindExcel = FindMenuItem(menu, "绑定Excel工程量");
                if (bindExcel == null)
                {
                    bindExcel = new ToolStripMenuItem("绑定Excel工程量");
                    bindExcel.Visible = true;
                    bindExcel.Available = true;
                    bindExcel.Enabled = true;
                    bindExcel.Click += delegate { SmartBindSelectedQuotasToExcel(mainForm); };
                    menu.Items.Insert(Math.Min(baseIndex, menu.Items.Count), bindExcel);
                    baseIndex++;
                }
                ApplyMenuIcon(bindExcel, "excel_bind.png");
                int bindIndex = menu.Items.IndexOf(bindExcel);
                if (bindIndex >= 0)
                {
                    baseIndex = Math.Max(baseIndex, bindIndex + 1);
                }

                ToolStripMenuItem trainMapping = FindMenuItem(menu, "\u6dfb\u52a0\u5bf9\u5e94\u6846\u5185\u5bb9");
                if (trainMapping == null)
                {
                    trainMapping = new ToolStripMenuItem("\u6dfb\u52a0\u5bf9\u5e94\u6846\u5185\u5bb9");
                    trainMapping.Visible = true;
                    trainMapping.Available = true;
                    trainMapping.Enabled = true;
                    trainMapping.Click += delegate { ShowMappingBoxTrainerPanel(mainForm); };
                    menu.Items.Insert(Math.Min(baseIndex, menu.Items.Count), trainMapping);
                    baseIndex++;
                }
                ApplyMenuIcon(trainMapping, "recommend_quota.png");

                ToolStripMenuItem openPanel = FindMenuItem(menu, "打开Excel联动面板");
                if (openPanel == null)
                {
                    openPanel = new ToolStripMenuItem("打开Excel联动面板");
                    openPanel.Visible = true;
                    openPanel.Available = true;
                    openPanel.Enabled = true;
                    openPanel.Click += delegate { ShowExcelLinkPanel(mainForm); };
                    menu.Items.Insert(Math.Min(baseIndex, menu.Items.Count), openPanel);
                }
                ApplyMenuIcon(openPanel, "excel_panel.png");
        }

        private static void RemoveMenuItem(ToolStrip menu, string text)
        {
            ToolStripMenuItem item = FindMenuItem(menu, text);
            if (item == null)
            {
                return;
            }

            menu.Items.Remove(item);
            item.Dispose();
        }

        private static void SmartBindSelectedQuotasToExcel(Form mainForm)
        {
            ShowExcelSmartBindPanel(mainForm);
        }

        private static void ShowExcelSmartBindPanel(Form mainForm)
        {
            EnsureExcelLinkRuntime(mainForm);
            ExcelSmartBindPanel panel;
            if (!ExcelSmartBindPanels.TryGetValue(mainForm, out panel) || panel == null || panel.IsDisposed)
            {
                panel = new ExcelSmartBindPanel(mainForm);
                ExcelSmartBindPanels[mainForm] = panel;
            }

            panel.Show(mainForm);
            panel.Activate();
            panel.RefreshCurrentContext();
        }

        private static bool PromptExcelBindOptions(IWin32Window owner, ExcelCellAddress cell, int selectedRowCount, out ExcelBindOptions options)
        {
            options = null;
            using (Form dialog = new Form())
            using (Label info = new Label())
            using (RadioButton continuous = new RadioButton())
            using (RadioButton expression = new RadioButton())
            using (Label countLabel = new Label())
            using (NumericUpDown countBox = new NumericUpDown())
            using (Label suffixLabel = new Label())
            using (TextBox suffixText = new TextBox())
            using (Label expressionLabel = new Label())
            using (TextBox expressionText = new TextBox())
            using (Button multiply = new Button())
            using (Button divide = new Button())
            using (Button ok = new Button())
            using (Button cancel = new Button())
            {
                dialog.Text = "绑定Excel工程量";
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.MinimizeBox = false;
                dialog.MaximizeBox = false;
                dialog.ClientSize = new System.Drawing.Size(560, 255);

                info.Left = 12;
                info.Top = 12;
                info.Width = 530;
                info.Height = 40;
                info.Text = "Excel当前选择：" + Path.GetFileName(cell.WorkbookPath) + "!" + cell.WorksheetName + "!" + BuildSelectionDisplay(cell) + "；软件已选定额：" + selectedRowCount.ToString(CultureInfo.InvariantCulture) + " 条。";

                continuous.Left = 16;
                continuous.Top = 60;
                continuous.Width = 250;
                continuous.Text = "连续绑定：从当前格向下对应定额";

                countLabel.Left = 36;
                countLabel.Top = 92;
                countLabel.Width = 90;
                countLabel.Text = "绑定条数";

                countBox.Left = 130;
                countBox.Top = 88;
                countBox.Width = 80;
                countBox.Minimum = 1;
                countBox.Maximum = 500;
                countBox.Value = Math.Max(1, selectedRowCount);

                suffixLabel.Left = 230;
                suffixLabel.Top = 92;
                suffixLabel.Width = 80;
                suffixLabel.Text = "每行附加";

                suffixText.Left = 310;
                suffixText.Top = 88;
                suffixText.Width = 120;
                suffixText.Text = "";

                expression.Left = 16;
                expression.Top = 126;
                expression.Width = 250;
                expression.Text = "表达式绑定：求和、乘除或手写公式";

                expressionLabel.Left = 36;
                expressionLabel.Top = 158;
                expressionLabel.Width = 90;
                expressionLabel.Text = "表达式";

                expressionText.Left = 130;
                expressionText.Top = 154;
                expressionText.Width = 300;
                expressionText.Text = BuildDefaultExpression(cell);

                multiply.Text = "*";
                multiply.Left = 438;
                multiply.Top = 152;
                multiply.Width = 28;
                multiply.Click += delegate
                {
                    expression.Checked = true;
                    expressionText.Text = expressionText.Text + "*";
                    expressionText.SelectionStart = expressionText.Text.Length;
                    expressionText.Focus();
                };

                divide.Text = "/";
                divide.Left = 470;
                divide.Top = 152;
                divide.Width = 28;
                divide.Click += delegate
                {
                    expression.Checked = true;
                    expressionText.Text = expressionText.Text + "/";
                    expressionText.SelectionStart = expressionText.Text.Length;
                    expressionText.Focus();
                };

                ok.Text = "确定";
                ok.Left = 370;
                ok.Top = 210;
                ok.Width = 80;
                ok.DialogResult = DialogResult.OK;

                cancel.Text = "取消";
                cancel.Left = 460;
                cancel.Top = 210;
                cancel.Width = 80;
                cancel.DialogResult = DialogResult.Cancel;

                if (cell.SelectionAddresses != null && cell.SelectionAddresses.Count > 1)
                {
                    expression.Checked = true;
                }
                else
                {
                    continuous.Checked = selectedRowCount > 1;
                    expression.Checked = selectedRowCount <= 1;
                }

                dialog.Controls.Add(info);
                dialog.Controls.Add(continuous);
                dialog.Controls.Add(countLabel);
                dialog.Controls.Add(countBox);
                dialog.Controls.Add(suffixLabel);
                dialog.Controls.Add(suffixText);
                dialog.Controls.Add(expression);
                dialog.Controls.Add(expressionLabel);
                dialog.Controls.Add(expressionText);
                dialog.Controls.Add(multiply);
                dialog.Controls.Add(divide);
                dialog.Controls.Add(ok);
                dialog.Controls.Add(cancel);
                dialog.AcceptButton = ok;
                dialog.CancelButton = cancel;

                while (dialog.ShowDialog(owner) == DialogResult.OK)
                {
                    string suffix = suffixText.Text.Trim();
                    string expr = expressionText.Text.Trim().ToUpperInvariant();
                    if (expression.Checked && String.IsNullOrEmpty(expr))
                    {
                        MessageBox.Show(owner, "请填写表达式，例如 E4+E5 或 E4*0.8。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        expressionText.Focus();
                        continue;
                    }

                    if (!String.IsNullOrEmpty(suffix) && !IsValidRowSuffix(suffix))
                    {
                        MessageBox.Show(owner, "每行附加只支持 *系数 或 /系数，例如 *0.8、/1.05。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        suffixText.Focus();
                        suffixText.SelectAll();
                        continue;
                    }

                    options = new ExcelBindOptions();
                    options.ExpressionMode = expression.Checked;
                    options.RowCount = Convert.ToInt32(countBox.Value, CultureInfo.InvariantCulture);
                    options.RowSuffix = suffix;
                    options.Expression = expr;
                    return true;
                }
            }

            return false;
        }

        private static string BuildSelectionDisplay(ExcelCellAddress cell)
        {
            if (cell.SelectionAddresses == null || cell.SelectionAddresses.Count <= 1)
            {
                return cell.CellAddress;
            }

            return cell.SelectionAddresses[0] + "..." + cell.SelectionAddresses[cell.SelectionAddresses.Count - 1];
        }

        private static string BuildDefaultExpression(ExcelCellAddress cell)
        {
            if (cell.SelectionAddresses == null || cell.SelectionAddresses.Count == 0)
            {
                return cell.CellAddress;
            }

            return String.Join("+", cell.SelectionAddresses.ToArray());
        }

        private static bool IsValidRowSuffix(string suffix)
        {
            if (String.IsNullOrEmpty(suffix))
            {
                return true;
            }

            char op = suffix[0];
            if (op != '*' && op != '/')
            {
                return false;
            }

            decimal value;
            return Decimal.TryParse(suffix.Substring(1).Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                || Decimal.TryParse(suffix.Substring(1).Trim(), NumberStyles.Float, CultureInfo.CurrentCulture, out value);
        }

        private static List<DataGridViewRow> ResolveTargetQuotaRows(DataGridView grid, List<DataGridViewRow> selectedRows, int requestedCount)
        {
            if (selectedRows.Count >= requestedCount)
            {
                return selectedRows.OrderBy(r => r.Index).Take(requestedCount).ToList();
            }

            return GetConsecutiveQuotaRows(grid, requestedCount);
        }

        private static List<DataGridViewRow> GetConsecutiveQuotaRows(DataGridView grid, int requestedCount)
        {
            List<DataGridViewRow> rows = new List<DataGridViewRow>();
            DataGridViewRow startRow = GetCurrentQuotaRow(grid);
            if (grid == null || startRow == null || startRow.Index < 0)
            {
                return rows;
            }

            for (int i = startRow.Index; i < grid.Rows.Count && rows.Count < requestedCount; i++)
            {
                DataGridViewRow row = grid.Rows[i];
                if (row != null && !row.IsNewRow)
                {
                    rows.Add(row);
                }
            }

            return rows;
        }

        private static string BuildRowExpression(string startAddress, int offset, string suffix)
        {
            CellRef start;
            if (!TryParseCellAddress(startAddress, out start))
            {
                return startAddress + (suffix ?? "");
            }

            string address = ColumnNumberToName(start.Column) + (start.Row + offset).ToString(CultureInfo.InvariantCulture);
            return address + (suffix ?? "");
        }

        private static void BindSelectedQuotaToExcel(Form mainForm)
        {
            try
            {
                SqlConnection conn = GetProjectConnection(mainForm);
                if (conn == null)
                {
                    MessageBox.Show(mainForm, "没有找到当前项目数据库连接。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
                if (grid == null)
                {
                    MessageBox.Show(mainForm, "没有找到定额输入表格。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DataGridViewRow row = GetCurrentQuotaRow(grid);
                if (row == null)
                {
                    MessageBox.Show(mainForm, "请先选择一条定额行。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                ExcelQuotaLink link;
                string error;
                if (!TryCreateQuotaLink(mainForm, conn, row, out link, out error))
                {
                    MessageBox.Show(mainForm, error, "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                ExcelCellAddress cell;
                if (!TryGetActiveExcelCell(out cell, out error))
                {
                    if (!PromptExcelCell(mainForm, error, out cell))
                    {
                        return;
                    }
                }

                link.ExcelPath = cell.WorkbookPath;
                link.WorksheetName = cell.WorksheetName;
                link.CellAddress = cell.CellAddress;
                link.Expression = cell.CellAddress;
                link.LastSyncValue = cell.DisplayValue;
                link.LastStatus = "已绑定，等待同步";
                link.UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

                ExcelLinkStore store = LoadStore(conn);
                store.Upsert(link);
                SaveStore(conn, store);

                EnsureExcelLinkRuntime(mainForm);
                if (ExcelLinkRuntimes.ContainsKey(mainForm))
                {
                    ExcelLinkRuntimes[mainForm].Reload();
                }

                RefreshExcelLinkPanel(mainForm);
                MessageBox.Show(mainForm, "已绑定：" + link.QuotaCode + " -> " + Path.GetFileName(link.ExcelPath) + "!" + link.WorksheetName + "!" + link.CellAddress, "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log("BindSelectedQuotaToExcel failed: " + ex);
                MessageBox.Show(mainForm, "绑定失败：" + ex.Message, "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void BatchBindSelectedQuotasToExcel(Form mainForm)
        {
            try
            {
                SqlConnection conn = GetProjectConnection(mainForm);
                if (conn == null)
                {
                    MessageBox.Show(mainForm, "没有找到当前项目数据库连接。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
                List<DataGridViewRow> rows = GetSelectedQuotaRows(grid);
                if (rows.Count < 2)
                {
                    MessageBox.Show(mainForm, "请先在定额输入表中选择两条或更多连续定额行。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                ExcelCellAddress startCell;
                string error;
                if (!TryGetActiveExcelCell(out startCell, out error))
                {
                    if (!PromptExcelCell(mainForm, error, out startCell))
                    {
                        return;
                    }
                }

                CellRef startRef;
                if (!TryParseCellAddress(startCell.CellAddress, out startRef))
                {
                    MessageBox.Show(mainForm, "起始单元格地址不正确，例如 E4。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                ExcelLinkStore store = LoadStore(conn);
                int bound = 0;
                foreach (DataGridViewRow row in rows)
                {
                    ExcelQuotaLink link;
                    if (!TryCreateQuotaLink(mainForm, conn, row, out link, out error))
                    {
                        continue;
                    }

                    string address = ColumnNumberToName(startRef.Column) + (startRef.Row + bound).ToString(CultureInfo.InvariantCulture);
                    string displayValue;
                    string readError;
                    TryReadXlsxCellValue(startCell.WorkbookPath, startCell.WorksheetName, address, out displayValue, out readError);

                    link.ExcelPath = startCell.WorkbookPath;
                    link.WorksheetName = startCell.WorksheetName;
                    link.CellAddress = address;
                    link.Expression = address;
                    link.LastSyncValue = displayValue ?? "";
                    link.LastStatus = "批量绑定，等待同步";
                    link.UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                    store.Upsert(link);
                    bound++;
                }

                SaveStore(conn, store);
                EnsureExcelLinkRuntime(mainForm);
                if (ExcelLinkRuntimes.ContainsKey(mainForm))
                {
                    ExcelLinkRuntimes[mainForm].Reload();
                }

                RefreshExcelLinkPanel(mainForm);
                MessageBox.Show(mainForm, "已批量绑定 " + bound.ToString(CultureInfo.InvariantCulture) + " 条定额，从 " + startCell.WorksheetName + "!" + startCell.CellAddress + " 开始向下对应。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log("BatchBindSelectedQuotasToExcel failed: " + ex);
                MessageBox.Show(mainForm, "批量绑定失败：" + ex.Message, "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void ShowExcelLinkPanel(Form mainForm)
        {
            EnsureExcelLinkRuntime(mainForm);
            ExcelLinkPanel panel;
            if (!ExcelLinkPanels.TryGetValue(mainForm, out panel) || panel == null || panel.IsDisposed)
            {
                panel = new ExcelLinkPanel(mainForm);
                ExcelLinkPanels[mainForm] = panel;
            }

            panel.Reload();
            panel.Show(mainForm);
            panel.Activate();
        }

        private static void ShowQuickBindPanel(Form mainForm)
        {
            EnsureExcelLinkRuntime(mainForm);
            QuickBindPanel panel;
            if (!QuickBindPanels.TryGetValue(mainForm, out panel) || panel == null || panel.IsDisposed)
            {
                panel = new QuickBindPanel(mainForm);
                QuickBindPanels[mainForm] = panel;
            }

            panel.Show(mainForm);
            panel.Activate();
        }

        private static void RefreshExcelLinkPanel(Form mainForm)
        {
            ExcelLinkPanel panel;
            if (ExcelLinkPanels.TryGetValue(mainForm, out panel) && panel != null && !panel.IsDisposed)
            {
                panel.Reload();
            }
        }

        private static DataGridViewRow GetCurrentQuotaRow(DataGridView grid)
        {
            if (grid == null)
            {
                return null;
            }

            if (grid.SelectedRows.Count > 0)
            {
                return grid.SelectedRows[0];
            }

            if (grid.SelectedCells.Count > 0)
            {
                int rowIndex = grid.SelectedCells[0].RowIndex;
                if (rowIndex >= 0 && rowIndex < grid.Rows.Count)
                {
                    return grid.Rows[rowIndex];
                }
            }

            return grid.CurrentRow;
        }

        private static List<DataGridViewRow> GetSelectedQuotaRows(DataGridView grid)
        {
            Dictionary<int, DataGridViewRow> rows = new Dictionary<int, DataGridViewRow>();
            if (grid == null)
            {
                return new List<DataGridViewRow>();
            }

            foreach (DataGridViewRow row in grid.SelectedRows)
            {
                if (row != null && !row.IsNewRow && row.Index >= 0)
                {
                    rows[row.Index] = row;
                }
            }

            foreach (DataGridViewCell cell in grid.SelectedCells)
            {
                if (cell.RowIndex >= 0 && cell.RowIndex < grid.Rows.Count)
                {
                    DataGridViewRow row = grid.Rows[cell.RowIndex];
                    if (row != null && !row.IsNewRow)
                    {
                        rows[row.Index] = row;
                    }
                }
            }

            if (rows.Count == 0 && grid.CurrentRow != null && !grid.CurrentRow.IsNewRow)
            {
                rows[grid.CurrentRow.Index] = grid.CurrentRow;
            }

            return rows.OrderBy(k => k.Key).Select(k => k.Value).ToList();
        }

        private static bool TryCreateQuotaLink(Form mainForm, SqlConnection conn, DataGridViewRow row, out ExcelQuotaLink link, out string error)
        {
            link = null;
            error = null;

            QuotaKey key;
            if (!TryGetQuotaKey(row, out key))
            {
                error = "无法识别当前定额行的总概算序号、条目序号或顺号。";
                return false;
            }

            long quotaSequence;
            string quotaSequenceText = GetRowValue(row, "定额序号", "定额序号DE");
            if (!Int64.TryParse(quotaSequenceText, NumberStyles.Integer, CultureInfo.InvariantCulture, out quotaSequence))
            {
                quotaSequence = ResolveQuotaSequence(conn, key);
            }

            if (quotaSequence <= 0)
            {
                error = "无法识别当前定额行的定额序号。";
                return false;
            }

            string quotaCode = GetRowValue(row, "定额编号", "定额编号DE", "编号");
            string quotaName = GetRowValue(row, "工程或费用项目名称", "名称", "项目名称");

            link = new ExcelQuotaLink();
            link.ProjectId = GetProjectId(conn);
            link.QuotaSequence = quotaSequence;
            link.TotalNo = key.TotalNo;
            link.ChapterSeq = key.ChapterSeq;
            link.OrderNo = key.OrderNo;
            link.QuotaCode = quotaCode;
            link.QuotaName = quotaName;
            return true;
        }

        private static long ResolveQuotaSequence(SqlConnection conn, QuotaKey key)
        {
            EnsureOpen(conn);
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 定额序号 from 定额输入 where 总概算序号=@zgs and 条目序号=@tm and 顺号=@xh";
                cmd.Parameters.AddWithValue("@zgs", key.TotalNo);
                cmd.Parameters.AddWithValue("@tm", key.ChapterSeq);
                cmd.Parameters.AddWithValue("@xh", key.OrderNo);
                object result = cmd.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                {
                    return Convert.ToInt64(result, CultureInfo.InvariantCulture);
                }
            }

            return 0;
        }

        private static bool TryGetActiveExcelCell(out ExcelCellAddress cell, out string error)
        {
            cell = null;
            error = null;
            dynamic excel = null;
            try
            {
                excel = GetActiveSpreadsheetApplication();
                if (excel == null)
                {
                    error = BuildExcelConnectError("没有找到正在运行的 Excel/WPS 表格");
                    return false;
                }
                dynamic workbook = excel.ActiveWorkbook;
                dynamic sheet = excel.ActiveSheet;
                dynamic selection = excel.Selection;
                if (workbook == null || sheet == null || selection == null)
                {
                    error = BuildExcelConnectError("已经连接到 Excel/WPS，但没有读到当前工作簿、工作表或选中单元格");
                    return false;
                }

                dynamic firstCell = selection.Cells[1, 1];
                string workbookPath = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture);
                string worksheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture);
                string address = Convert.ToString(firstCell.Address(false, false), CultureInfo.InvariantCulture);
                object value = firstCell.Value2;
                if (String.IsNullOrEmpty(workbookPath) || String.IsNullOrEmpty(worksheetName) || String.IsNullOrEmpty(address))
                {
                    error = BuildExcelConnectError("无法读取当前 Excel 单元格地址");
                    return false;
                }

                cell = new ExcelCellAddress();
                cell.WorkbookPath = workbookPath;
                cell.WorksheetName = worksheetName;
                cell.CellAddress = address;
                cell.DisplayValue = ExcelValueToText(value);
                cell.SelectionAddresses = ReadSelectedCellAddresses(selection);
                return true;
            }
            catch (COMException ex)
            {
                error = BuildExcelConnectError("读取 Excel/WPS 当前单元格失败：" + ex.Message);
                return false;
            }
            catch (Exception ex)
            {
                error = BuildExcelConnectError("读取 Excel/WPS 当前单元格失败：" + ex.Message);
                return false;
            }
        }

        private static string BuildExcelConnectError(string reason)
        {
            return reason + "。已尝试通过 Excel/WPS 注册接口和 Excel 窗口对象连接。若仍失败，请确认：1. Excel/WPS 已打开工程量表；2. 已单击选中一个工程数量单元格且没有处于编辑状态；3. 表格不在受保护视图；4. Excel/WPS 与本软件使用相同权限运行（都普通运行，或都以管理员运行）；5. 当前 Office/WPS 已正确安装并注册 COM。若现场仍连不上，可在随后弹出的手动窗口里选择文件、工作表并输入单元格地址完成绑定。";
        }

        private static List<string> ReadSelectedCellAddresses(dynamic selection)
        {
            List<string> addresses = new List<string>();
            try
            {
                int count = Convert.ToInt32(selection.Cells.Count, CultureInfo.InvariantCulture);
                int limit = Math.Min(count, 200);
                for (int i = 1; i <= limit; i++)
                {
                    dynamic selectedCell = selection.Cells[i];
                    string address = Convert.ToString(selectedCell.Address(false, false), CultureInfo.InvariantCulture);
                    if (!String.IsNullOrEmpty(address))
                    {
                        addresses.Add(address.ToUpperInvariant());
                    }
                }
            }
            catch (Exception ex)
            {
                Log("ReadSelectedCellAddresses failed: " + ex.Message);
            }

            return addresses;
        }

        private static List<AiQuotaMatchRow> BuildAiQuotaRows(Form mainForm, SqlConnection conn, List<DataGridViewRow> rows)
        {
            List<AiQuotaMatchRow> result = new List<AiQuotaMatchRow>();
            foreach (DataGridViewRow row in rows ?? new List<DataGridViewRow>())
            {
                ExcelQuotaLink link;
                string error;
                if (!TryCreateQuotaLink(mainForm, conn, row, out link, out error))
                {
                    continue;
                }

                string currentQuantity = GetRowValue(row, "工程数量输入", "工程数量");
                decimal currentQuantityValue;
                string currentQuantityError;
                if (TryEvaluateDecimal(currentQuantity, out currentQuantityValue, out currentQuantityError) && currentQuantityValue == 0m)
                {
                    continue;
                }

                result.Add(new AiQuotaMatchRow
                {
                    Link = link,
                    QuotaUnit = GetRowValue(row, "单位", "定额单位"),
                    CurrentQuantityText = currentQuantity
                });
            }

            return result;
        }

        private static bool TryReadActiveExcelSelectionContext(out AiExcelSelectionContext context, out string error)
        {
            context = null;
            error = null;
            try
            {
                dynamic excel = GetActiveSpreadsheetApplication();
                if (excel == null)
                {
                    error = BuildExcelConnectError("没有找到正在运行的 Excel/WPS 表格");
                    return false;
                }

                dynamic workbook = excel.ActiveWorkbook;
                dynamic sheet = excel.ActiveSheet;
                dynamic selection = excel.Selection;
                if (workbook == null || sheet == null || selection == null)
                {
                    error = BuildExcelConnectError("已经连接到 Excel/WPS，但没有读到当前工作簿、工作表或选区");
                    return false;
                }

                context = new AiExcelSelectionContext();
                context.WorkbookPath = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture);
                context.WorksheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture);

                if (TryReadActiveExcelSelectionContextBulk(selection, context))
                {
                    return true;
                }

                int count = Convert.ToInt32(selection.Cells.Count, CultureInfo.InvariantCulture);
                int limit = Math.Min(count, 800);
                for (int i = 1; i <= limit; i++)
                {
                    dynamic selectedCell = selection.Cells[i];
                    string address = NormalizeCellAddress(Convert.ToString(selectedCell.Address(false, false), CultureInfo.InvariantCulture));
                    object rawValue = selectedCell.Value2;
                    string text = ExcelValueToText(rawValue);
                    if (String.IsNullOrWhiteSpace(address) || String.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    CellRef cellRef;
                    if (!TryParseCellAddress(address, out cellRef))
                    {
                        continue;
                    }

                    decimal parsed;
                    string parseError;
                    context.Cells.Add(new AiExcelCell
                    {
                        Address = address,
                        Text = text,
                        Row = cellRef.Row,
                        Column = cellRef.Column,
                        IsNumber = TryEvaluateDecimal(text, out parsed, out parseError)
                    });
                }

                if (context.Cells.Count == 0)
                {
                    error = "当前 Excel 选区没有可读取的内容。请先框选包含工程量名称和数量的区域。";
                    return false;
                }

                if (!context.Cells.Any(c => c.IsNumber))
                {
                    error = "当前 Excel 选区没有可计算的数量单元格。";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                error = BuildExcelConnectError("读取 Excel/WPS 当前选区失败：" + ex.Message);
                return false;
            }
        }

        private static bool TryReadActiveExcelSelectionContextBulk(dynamic selection, AiExcelSelectionContext context)
        {
            try
            {
                int rowCount = Convert.ToInt32(selection.Rows.Count, CultureInfo.InvariantCulture);
                int colCount = Convert.ToInt32(selection.Columns.Count, CultureInfo.InvariantCulture);
                if (rowCount <= 0 || colCount <= 0 || rowCount * colCount > 1200)
                {
                    return false;
                }

                int firstRow = Convert.ToInt32(selection.Row, CultureInfo.InvariantCulture);
                int firstColumn = Convert.ToInt32(selection.Column, CultureInfo.InvariantCulture);
                object rawValues = selection.Value2;
                if (rowCount == 1 && colCount == 1)
                {
                    AddAiExcelCell(context, firstRow, firstColumn, rawValues);
                }
                else
                {
                    Array values = rawValues as Array;
                    if (values == null)
                    {
                        return false;
                    }

                    for (int row = 1; row <= rowCount; row++)
                    {
                        for (int col = 1; col <= colCount; col++)
                        {
                            AddAiExcelCell(context, firstRow + row - 1, firstColumn + col - 1, values.GetValue(row, col));
                        }
                    }
                }

                return context.Cells.Count > 0 && context.Cells.Any(c => c.IsNumber);
            }
            catch (Exception ex)
            {
                Log("Bulk read active Excel selection failed: " + ex.Message);
                context.Cells.Clear();
                return false;
            }
        }

        private static void AddAiExcelCell(AiExcelSelectionContext context, int row, int column, object rawValue)
        {
            string text = ExcelValueToText(rawValue);
            if (String.IsNullOrWhiteSpace(text))
            {
                return;
            }

            decimal parsed;
            string parseError;
            context.Cells.Add(new AiExcelCell
            {
                Address = ColumnNumberToName(column) + row.ToString(CultureInfo.InvariantCulture),
                Text = text,
                Row = row,
                Column = column,
                IsNumber = TryEvaluateDecimal(text, out parsed, out parseError)
            });
        }

        private static DeepSeekExcelMatchSettings LoadDeepSeekExcelMatchSettings()
        {
            DeepSeekExcelMatchSettings settings = new DeepSeekExcelMatchSettings();
            string path = Path.Combine(FindRecoQuotaDataDir(), "deepseek-settings.json");
            if (!File.Exists(path))
            {
                return settings;
            }

            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                Dictionary<string, object> values = serializer.DeserializeObject(File.ReadAllText(path, Encoding.UTF8)) as Dictionary<string, object>;
                if (values == null)
                {
                    return settings;
                }

                settings.Enabled = ReadJsonBool(values, "enabled", false);
                settings.ApiKey = ReadJsonString(values, "api_key", "");
                settings.Model = ReadJsonString(values, "model", settings.Model);
                settings.BaseUrl = ReadJsonString(values, "base_url", settings.BaseUrl).TrimEnd('/');
                settings.TimeoutSeconds = ClampInt(ReadJsonInt(values, "timeout_seconds", settings.TimeoutSeconds), 5, 120);
                settings.MaxRowsPerBatch = ClampInt(ReadJsonInt(values, "max_rows_per_batch", settings.MaxRowsPerBatch), 1, 20);
            }
            catch (Exception ex)
            {
                Log("Load DeepSeek settings for Excel match failed: " + ex.Message);
                settings.Enabled = false;
            }

            return settings;
        }

        private static List<AiMatchResult> RequestAiExcelMatches(DeepSeekExcelMatchSettings settings, List<AiQuotaMatchRow> quotas, AiExcelSelectionContext selection)
        {
            List<AiMatchResult> localMatches = BuildLocalQuantityMatchResults(quotas, selection);
            HashSet<long> localMatchedIds = new HashSet<long>(localMatches.Select(match => match.QuotaSequence));
            List<AiQuotaMatchRow> remaining = (quotas ?? new List<AiQuotaMatchRow>())
                .Where(row => row != null && row.Link != null && !localMatchedIds.Contains(row.Link.QuotaSequence))
                .ToList();
            if (remaining.Count == 0)
            {
                return localMatches;
            }

            JavaScriptSerializer serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = 1024 * 1024 * 4;
            List<AiMatchResult> all = new List<AiMatchResult>(localMatches);
            int batchSize = Math.Max(1, settings.MaxRowsPerBatch);
            Exception lastBatchError = null;
            int successfulBatches = 0;
            for (int i = 0; i < remaining.Count; i += batchSize)
            {
                List<AiQuotaMatchRow> batch = remaining.Skip(i).Take(batchSize).ToList();
                try
                {
                    string requestJson = BuildAiExcelMatchRequestJson(serializer, settings, batch, selection);
                    string responseJson = SendDeepSeekExcelMatchRequest(settings, requestJson);
                    all.AddRange(ParseAiExcelMatchResponse(serializer, responseJson));
                    successfulBatches++;
                }
                catch (Exception ex)
                {
                    lastBatchError = ex;
                    Log("DeepSeek Excel match batch failed: " + ex.Message);
                }
            }

            // 全部批次失败时仍按原行为抛出，让调用方提示并回退本地数量匹配；部分失败则保留成功批次结果。
            if (lastBatchError != null && successfulBatches == 0)
            {
                throw lastBatchError;
            }

            return all;
        }

        private static string BuildAiExcelMatchRequestJson(JavaScriptSerializer serializer, DeepSeekExcelMatchSettings settings, List<AiQuotaMatchRow> quotas, AiExcelSelectionContext selection)
        {
            Dictionary<int, List<AiExcelCell>> rows = selection.Cells
                .GroupBy(c => c.Row)
                .OrderBy(g => g.Key)
                .Take(80)
                .ToDictionary(g => g.Key, g => g.OrderBy(c => c.Column).ToList());

            List<object> quotaObjects = new List<object>();
            foreach (AiQuotaMatchRow row in quotas)
            {
                quotaObjects.Add(new Dictionary<string, object>
                {
                    { "quota_sequence", row.Link.QuotaSequence.ToString(CultureInfo.InvariantCulture) },
                    { "quota_code", row.Link.QuotaCode ?? "" },
                    { "quota_name", row.Link.QuotaName ?? "" },
                    { "quota_unit", row.QuotaUnit ?? "" },
                    { "current_quantity", row.CurrentQuantityText ?? "" }
                });
            }

            List<object> rowObjects = new List<object>();
            foreach (KeyValuePair<int, List<AiExcelCell>> pair in rows)
            {
                List<object> cells = new List<object>();
                foreach (AiExcelCell cell in pair.Value)
                {
                    cells.Add(new Dictionary<string, object>
                    {
                        { "address", cell.Address },
                        { "text", TruncateForPrompt(cell.Text, 120) },
                        { "is_number", cell.IsNumber }
                    });
                }

                rowObjects.Add(new Dictionary<string, object>
                {
                    { "row", pair.Key },
                    { "cells", cells }
                });
            }

            Dictionary<string, object> body = new Dictionary<string, object>();
            body["task"] = "把软件中已选定额逐条匹配到Excel选区里的工程数量单元格。";
            body["rules"] = new string[]
            {
                "只能选择rows里存在的Excel单元格地址，不能编造地址。",
                "优先结合定额名称、定额单位、当前工程数量和Excel整行文字上下文匹配。",
                "如果一个定额数量由多个单元格组成，可以用 A1+B1、A1*2、A1/2 这样的表达式，但表达式里的单元格必须来自选区。",
                "不确定时不要返回该定额。",
                "返回严格JSON：{\"results\":[{\"quota_sequence\":\"123\",\"cell_address\":\"H3\",\"expression\":\"H3\",\"quantity_name\":\"工程量名称\",\"confidence\":80,\"reason\":\"简短理由\"}]}"
            };
            body["quotas"] = quotaObjects;
            body["rows"] = rowObjects;

            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["model"] = settings.Model;
            payload["stream"] = false;
            payload["temperature"] = 0.1;
            payload["max_tokens"] = 3000;
            payload["response_format"] = new Dictionary<string, object> { { "type", "json_object" } };
            payload["messages"] = new object[]
            {
                new Dictionary<string, object>
                {
                    { "role", "system" },
                    { "content", "你是铁路工程预算软件的Excel工程量绑定助手。你只做定额与Excel数量单元格匹配，必须保守，不能编造单元格地址。" }
                },
                new Dictionary<string, object>
                {
                    { "role", "user" },
                    { "content", serializer.Serialize(body) }
                }
            };

            return serializer.Serialize(payload);
        }

        private static string SendDeepSeekExcelMatchRequest(DeepSeekExcelMatchSettings settings, string requestJson)
        {
            ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol | (SecurityProtocolType)3072;
            string endpoint = settings.BaseUrl.TrimEnd('/');
            if (!endpoint.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase))
            {
                endpoint += "/chat/completions";
            }

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(endpoint);
            request.Method = "POST";
            request.ContentType = "application/json";
            request.Accept = "application/json";
            request.Headers["Authorization"] = "Bearer " + settings.ApiKey;
            int timeoutSeconds = Math.Max(5, settings.TimeoutSeconds);
            request.Timeout = timeoutSeconds * 1000;
            request.ReadWriteTimeout = timeoutSeconds * 1000;

            byte[] payload = Encoding.UTF8.GetBytes(requestJson);
            request.ContentLength = payload.Length;
            using (Stream stream = request.GetRequestStream())
            {
                stream.Write(payload, 0, payload.Length);
            }

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
            {
                return reader.ReadToEnd();
            }
        }

        private static List<AiMatchResult> ParseAiExcelMatchResponse(JavaScriptSerializer serializer, string responseJson)
        {
            Dictionary<string, object> root = serializer.DeserializeObject(responseJson) as Dictionary<string, object>;
            List<object> choices = GetJsonList(root, "choices");
            Dictionary<string, object> firstChoice = choices == null || choices.Count == 0 ? null : choices[0] as Dictionary<string, object>;
            Dictionary<string, object> message = firstChoice == null ? null : ReadJsonObject(firstChoice, "message");
            string content = message == null ? "" : ReadJsonString(message, "content", "");
            if (String.IsNullOrWhiteSpace(content))
            {
                return new List<AiMatchResult>();
            }

            Dictionary<string, object> resultRoot = serializer.DeserializeObject(content) as Dictionary<string, object>;
            List<object> results = GetJsonList(resultRoot, "results");
            if (results == null)
            {
                return new List<AiMatchResult>();
            }

            List<AiMatchResult> matches = new List<AiMatchResult>();
            foreach (object item in results)
            {
                Dictionary<string, object> row = item as Dictionary<string, object>;
                if (row == null)
                {
                    continue;
                }

                long quotaSequence;
                if (!Int64.TryParse(ReadJsonString(row, "quota_sequence", ""), NumberStyles.Integer, CultureInfo.InvariantCulture, out quotaSequence))
                {
                    continue;
                }

                matches.Add(new AiMatchResult
                {
                    QuotaSequence = quotaSequence,
                    CellAddress = NormalizeCellAddress(ReadJsonString(row, "cell_address", "")),
                    Expression = NormalizeCellAddress(ReadJsonString(row, "expression", "")),
                    QuantityName = ReadJsonString(row, "quantity_name", ""),
                    Confidence = ReadJsonInt(row, "confidence", 0),
                    Reason = ReadJsonString(row, "reason", "")
                });
            }

            return matches;
        }

        private static List<AiMatchPreviewItem> BuildAiMatchPreviewItems(List<AiQuotaMatchRow> quotas, AiExcelSelectionContext selection, List<AiMatchResult> results)
        {
            Dictionary<long, AiQuotaMatchRow> quotaById = quotas.ToDictionary(q => q.Link.QuotaSequence);
            HashSet<string> selectedAddresses = selection.AddressSet;
            HashSet<long> usedQuotas = new HashSet<long>();
            List<AiMatchPreviewItem> preview = new List<AiMatchPreviewItem>();

            foreach (AiMatchResult result in results ?? new List<AiMatchResult>())
            {
                AiQuotaMatchRow quota;
                if (!quotaById.TryGetValue(result.QuotaSequence, out quota) || usedQuotas.Contains(result.QuotaSequence))
                {
                    continue;
                }

                string expression = String.IsNullOrWhiteSpace(result.Expression) ? result.CellAddress : result.Expression;
                expression = NormalizeExpressionOperators(expression);
                string firstCell = ExtractFirstCellAddress(expression);
                if (String.IsNullOrWhiteSpace(firstCell) || !selectedAddresses.Contains(firstCell))
                {
                    continue;
                }

                List<string> expressionCells = ExtractCellAddressesFromExpression(expression);
                if (expressionCells.Count == 0 || expressionCells.Any(address => !selectedAddresses.Contains(address)))
                {
                    continue;
                }

                string displayValue;
                decimal quantity;
                string readError;
                if (!TryEvaluateWorkbookExpression(selection.WorkbookPath, selection.WorksheetName, expression, out displayValue, out quantity, out readError))
                {
                    continue;
                }

                preview.Add(new AiMatchPreviewItem
                {
                    Checked = result.Confidence >= 65,
                    Link = quota.Link,
                    QuotaUnit = quota.QuotaUnit,
                    WorkbookPath = selection.WorkbookPath,
                    WorksheetName = selection.WorksheetName,
                    Expression = expression,
                    CellAddress = firstCell,
                    DisplayValue = displayValue,
                    QuantityName = String.IsNullOrWhiteSpace(result.QuantityName) ? BuildQuantityNameFromExcelRow(selection, firstCell) : result.QuantityName
                });
                usedQuotas.Add(result.QuotaSequence);
            }

            AddUnmatchedAiPreviewItems(preview, quotas, selection);
            return SortAiPreviewItemsByQuotaOrder(preview, quotas);
        }

        private static List<AiMatchPreviewItem> BuildLocalQuantityMatchPreviewItems(List<AiQuotaMatchRow> quotas, AiExcelSelectionContext selection)
        {
            List<AiMatchPreviewItem> preview = new List<AiMatchPreviewItem>();
            List<AiExcelCell> numberCells = selection.Cells
                .Where(c => c.IsNumber)
                .OrderBy(c => c.Row)
                .ThenBy(c => c.Column)
                .ToList();

            foreach (AiQuotaMatchRow quota in quotas ?? new List<AiQuotaMatchRow>())
            {
                decimal quotaQuantity;
                string quotaError;
                if (quota == null || quota.Link == null || !TryEvaluateDecimal(quota.CurrentQuantityText, out quotaQuantity, out quotaError))
                {
                    continue;
                }

                AiExcelCell bestCell = null;
                string bestExpression = null;
                decimal bestScore = Decimal.MaxValue;
                foreach (AiExcelCell cell in numberCells)
                {
                    string expression;
                    decimal score;
                    if (!TryBuildLocalQuantityExpression(quota.CurrentQuantityText, quotaQuantity, cell, out expression, out score))
                    {
                        continue;
                    }

                    if (score < bestScore)
                    {
                        bestScore = score;
                        bestCell = cell;
                        bestExpression = expression;
                    }
                }

                if (bestCell == null)
                {
                    continue;
                }

                preview.Add(new AiMatchPreviewItem
                {
                    Checked = true,
                    Link = quota.Link,
                    QuotaUnit = quota.QuotaUnit,
                    WorkbookPath = selection.WorkbookPath,
                    WorksheetName = selection.WorksheetName,
                    Expression = bestExpression,
                    CellAddress = bestCell.Address,
                    DisplayValue = FormatAiMatchDecimal(quotaQuantity),
                    QuantityName = BuildQuantityNameFromExcelRow(selection, bestCell.Address)
                });
            }

            AddUnmatchedAiPreviewItems(preview, quotas, selection);
            return SortAiPreviewItemsByQuotaOrder(preview, quotas);
        }

        private static List<AiMatchResult> BuildLocalQuantityMatchResults(List<AiQuotaMatchRow> quotas, AiExcelSelectionContext selection)
        {
            return BuildLocalQuantityMatchPreviewItems(quotas, selection)
                .Where(item => item != null && item.Link != null && !String.IsNullOrWhiteSpace(item.Expression))
                .Select(item => new AiMatchResult
                {
                    QuotaSequence = item.Link.QuotaSequence,
                    CellAddress = item.CellAddress,
                    Expression = item.Expression,
                    QuantityName = item.QuantityName,
                    Confidence = 100,
                    Reason = "本地数量匹配"
                })
                .ToList();
        }

        private static List<AiMatchPreviewItem> SortAiPreviewItemsByQuotaOrder(List<AiMatchPreviewItem> preview, List<AiQuotaMatchRow> quotas)
        {
            Dictionary<long, int> order = new Dictionary<long, int>();
            int index = 0;
            foreach (AiQuotaMatchRow row in quotas ?? new List<AiQuotaMatchRow>())
            {
                if (row != null && row.Link != null && !order.ContainsKey(row.Link.QuotaSequence))
                {
                    order[row.Link.QuotaSequence] = index++;
                }
            }

            return (preview ?? new List<AiMatchPreviewItem>())
                .OrderBy(item =>
                {
                    int value;
                    return item != null && item.Link != null && order.TryGetValue(item.Link.QuotaSequence, out value) ? value : Int32.MaxValue;
                })
                .ToList();
        }

        private static void AddUnmatchedAiPreviewItems(List<AiMatchPreviewItem> preview, List<AiQuotaMatchRow> quotas, AiExcelSelectionContext selection)
        {
            HashSet<long> existing = new HashSet<long>(preview
                .Where(item => item != null && item.Link != null)
                .Select(item => item.Link.QuotaSequence));

            foreach (AiQuotaMatchRow quota in quotas ?? new List<AiQuotaMatchRow>())
            {
                if (quota == null || quota.Link == null || existing.Contains(quota.Link.QuotaSequence))
                {
                    continue;
                }

                preview.Add(new AiMatchPreviewItem
                {
                    Checked = false,
                    Link = quota.Link,
                    QuotaUnit = quota.QuotaUnit,
                    WorkbookPath = selection == null ? null : selection.WorkbookPath,
                    WorksheetName = selection == null ? null : selection.WorksheetName,
                    Expression = "",
                    CellAddress = "",
                    DisplayValue = "",
                    QuantityName = ""
                });
            }
        }

        private static bool TryBuildLocalQuantityExpression(string quotaText, decimal quotaQuantity, AiExcelCell cell, out string expression, out decimal score)
        {
            expression = null;
            score = Decimal.MaxValue;

            decimal cellQuantity;
            string error;
            if (cell == null || !TryEvaluateDecimal(cell.Text, out cellQuantity, out error))
            {
                return false;
            }

            decimal directScore = RelativeDifference(quotaQuantity, cellQuantity);
            if (directScore <= 0.03m)
            {
                expression = cell.Address;
                score = directScore;
                return true;
            }

            string op;
            decimal factor;
            decimal leftValue;
            if (TryParseSimpleScaleExpression(quotaText, out leftValue, out op, out factor) &&
                RelativeDifference(leftValue, cellQuantity) <= 0.03m)
            {
                expression = cell.Address + op + factor.ToString(CultureInfo.InvariantCulture);
                score = RelativeDifference(leftValue, cellQuantity);
                return true;
            }

            return false;
        }

        private static bool TryParseSimpleScaleExpression(string text, out decimal leftValue, out string op, out decimal factor)
        {
            leftValue = 0;
            op = null;
            factor = 0;
            string value = (text ?? "").Trim().Replace("×", "*").Replace("X", "*").Replace("x", "*");
            int index = value.LastIndexOf('/');
            if (index < 0)
            {
                index = value.LastIndexOf('*');
            }

            if (index <= 0 || index >= value.Length - 1)
            {
                return false;
            }

            op = value[index].ToString();
            string left = value.Substring(0, index).Trim();
            string right = value.Substring(index + 1).Trim();
            string parseError;
            return TryEvaluateDecimal(left, out leftValue, out parseError) &&
                Decimal.TryParse(right, NumberStyles.Float, CultureInfo.InvariantCulture, out factor) &&
                factor != 0m;
        }

        private static decimal RelativeDifference(decimal expected, decimal actual)
        {
            decimal diff = Math.Abs(expected - actual);
            decimal baseValue = Math.Max(Math.Abs(expected), Math.Abs(actual));
            if (baseValue < 0.000001m)
            {
                return diff;
            }

            return diff / baseValue;
        }

        private static string FormatAiMatchDecimal(decimal value)
        {
            return value.ToString("0.########", CultureInfo.InvariantCulture);
        }

        private static string NormalizeExpressionOperators(string expression)
        {
            return (expression ?? "").Trim().ToUpperInvariant()
                .Replace("×", "*")
                .Replace("（", "(")
                .Replace("）", ")");
        }

        private static List<string> ExtractCellAddressesFromExpression(string expression)
        {
            List<string> addresses = new List<string>();
            string normalized = NormalizeExpressionOperators(expression);
            for (int i = 0; i < normalized.Length; i++)
            {
                if (normalized[i] >= 'A' && normalized[i] <= 'Z')
                {
                    int start = i;
                    while (i < normalized.Length && normalized[i] >= 'A' && normalized[i] <= 'Z')
                    {
                        i++;
                    }
                    while (i < normalized.Length && Char.IsDigit(normalized[i]))
                    {
                        i++;
                    }

                    string candidate = normalized.Substring(start, i - start);
                    CellRef parsed;
                    if (TryParseCellAddress(candidate, out parsed) && !addresses.Contains(candidate))
                    {
                        addresses.Add(candidate);
                    }
                }
            }

            return addresses;
        }

        private static string BuildQuantityNameFromExcelRow(AiExcelSelectionContext selection, string cellAddress)
        {
            CellRef target;
            if (!TryParseCellAddress(cellAddress, out target))
            {
                return "";
            }

            return String.Join(" ", selection.Cells
                .Where(c => c.Row == target.Row && !c.IsNumber)
                .OrderBy(c => c.Column)
                .Select(c => c.Text)
                .Where(text => !String.IsNullOrWhiteSpace(text))
                .Take(6)
                .ToArray()).Trim();
        }

        private static string BuildQuantityNameNearActiveExcelCell(ExcelCellAddress cell)
        {
            if (cell == null)
            {
                return "";
            }

            try
            {
                CellRef cellRef;
                if (!TryParseCellAddress(cell.CellAddress, out cellRef))
                {
                    return "";
                }

                dynamic excel = GetActiveSpreadsheetApplication();
                if (excel == null)
                {
                    return "";
                }

                dynamic sheet = excel.ActiveSheet;
                List<string> texts = new List<string>();
                int startColumn = Math.Max(1, cellRef.Column - 6);
                for (int col = startColumn; col < cellRef.Column; col++)
                {
                    string address = ColumnNumberToName(col) + cellRef.Row.ToString(CultureInfo.InvariantCulture);
                    dynamic range = sheet.Range[address];
                    string text = ExcelValueToText(range.Value2);
                    decimal parsed;
                    string parseError;
                    if (!String.IsNullOrWhiteSpace(text) && !TryEvaluateDecimal(text, out parsed, out parseError))
                    {
                        texts.Add(text);
                    }
                }

                return String.Join(" ", texts.ToArray()).Trim();
            }
            catch (Exception ex)
            {
                Log("BuildQuantityNameNearActiveExcelCell failed: " + ex.Message);
                return "";
            }
        }

        private static void AcceptAiMatchesToMappingStore(List<AiMatchPreviewItem> items)
        {
            if (items == null || items.Count == 0)
            {
                return;
            }

            try
            {
                WithMappingBoxesLock(delegate
                {
                string path = Path.Combine(FindRecoQuotaDataDir(), "mapping-boxes.jsonl");
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                List<Dictionary<string, string>> rows = new List<Dictionary<string, string>>();
                if (File.Exists(path))
                {
                    foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
                    {
                        Dictionary<string, string> parsed = ParseFlatJson(line);
                        if (parsed.Count > 0)
                        {
                            rows.Add(parsed);
                        }
                    }
                }

                foreach (AiMatchPreviewItem item in items)
                {
                    if (item == null || item.Link == null || String.IsNullOrWhiteSpace(item.Link.QuotaCode) || String.IsNullOrWhiteSpace(item.QuantityName))
                    {
                        continue;
                    }

                    string targetKey = "quota:" + item.Link.QuotaCode.Trim().ToUpperInvariant();
                    string boxId = FindExistingMappingBoxId(rows, targetKey) ?? BuildSingleQuotaBoxId(item.Link.QuotaCode);
                    string signature = NormalizeForSignature(item.QuantityName) + "|";
                    Dictionary<string, string> existing = rows.FirstOrDefault(row =>
                        String.Equals(GetFlat(row, "box_id"), boxId, StringComparison.OrdinalIgnoreCase) &&
                        String.Equals(NormalizeForSignature(GetFlat(row, "quantity_name")) + "|" + NormalizeForSignature(GetFlat(row, "quantity_unit")), signature, StringComparison.OrdinalIgnoreCase));

                    if (existing == null)
                    {
                        existing = new Dictionary<string, string>();
                        rows.Add(existing);
                        existing["record_type"] = "mapping_box";
                        existing["box_id"] = boxId;
                        existing["target_kind"] = "quota";
                        existing["target_code"] = item.Link.QuotaCode ?? "";
                        existing["target_name"] = item.Link.QuotaName ?? "";
                        existing["target_unit"] = item.QuotaUnit ?? "";
                        existing["quantity_name"] = item.QuantityName ?? "";
                        existing["quantity_unit"] = "";
                        existing["weight"] = "10";
                        existing["accepted_count"] = "0";
                        existing["corrected_count"] = "0";
                        existing["rejected_count"] = "0";
                    }

                    existing["weight"] = (ReadFlatInt(existing, "weight", 0) + 5).ToString(CultureInfo.InvariantCulture);
                    existing["accepted_count"] = (ReadFlatInt(existing, "accepted_count", 0) + 1).ToString(CultureInfo.InvariantCulture);
                    existing["last_used_at"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                }

                TrimMappingRows(rows, 30);
                File.WriteAllLines(path, rows.Select(ToFlatJson).ToArray(), Encoding.UTF8);
                });
            }
            catch (Exception ex)
            {
                Log("Accept AI Excel matches to mapping store failed: " + ex.Message);
            }
        }

        private static string FindRecoQuotaDataDir()
        {
            return Path.Combine(Path.GetDirectoryName(typeof(FormPanel).Assembly.Location), "RecoQuotaData");
        }

        private static void TrimMappingRows(List<Dictionary<string, string>> rows, int maxSamplesPerBox)
        {
            foreach (IGrouping<string, Dictionary<string, string>> boxGroup in rows
                .Where(row => !String.IsNullOrWhiteSpace(GetFlat(row, "box_id")))
                .GroupBy(row => GetFlat(row, "box_id"), StringComparer.OrdinalIgnoreCase)
                .ToList())
            {
                List<IGrouping<string, Dictionary<string, string>>> sampleGroups = boxGroup
                    .GroupBy(row => NormalizeForSignature(GetFlat(row, "quantity_name")) + "|" + NormalizeForSignature(GetFlat(row, "quantity_unit")), StringComparer.OrdinalIgnoreCase)
                    .Where(group => !String.IsNullOrWhiteSpace(group.Key.Trim('|')))
                    .ToList();
                if (sampleGroups.Count <= maxSamplesPerBox)
                {
                    continue;
                }

                HashSet<string> removeSamples = new HashSet<string>(
                    sampleGroups
                        .OrderBy(group => group.Min(row => ReadFlatInt(row, "weight", 0)))
                        .ThenBy(group => group.Min(row => GetFlat(row, "last_used_at")))
                        .Take(sampleGroups.Count - maxSamplesPerBox)
                        .Select(group => group.Key),
                    StringComparer.OrdinalIgnoreCase);

                rows.RemoveAll(row =>
                    String.Equals(GetFlat(row, "box_id"), boxGroup.Key, StringComparison.OrdinalIgnoreCase) &&
                    removeSamples.Contains(NormalizeForSignature(GetFlat(row, "quantity_name")) + "|" + NormalizeForSignature(GetFlat(row, "quantity_unit"))));
            }
        }

        private static string BuildSingleQuotaBoxId(string quotaCode)
        {
            return BuildStableMappingBoxId("quota:" + (quotaCode ?? "").Trim().ToUpperInvariant());
        }

        // 与 RecoQuotaRecommend 的 MappingStore 使用同一套规则：对小写化的目标键做 SHA1。
        // String.GetHashCode 在 x86/x64 进程间不一致且会碰撞，不能用于持久化 ID。
        private static string BuildStableMappingBoxId(string raw)
        {
            using (SHA1 sha = SHA1.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes((raw ?? "").ToLowerInvariant()));
                StringBuilder builder = new StringBuilder("box-");
                for (int i = 0; i < 8; i++)
                {
                    builder.Append(hash[i].ToString("x2", CultureInfo.InvariantCulture));
                }

                return builder.ToString();
            }
        }

        // 旧版文件里的 box_id 是 GetHashCode 生成的数字，按目标组合查找可同时兼容新旧 ID。
        private static string FindExistingMappingBoxId(List<Dictionary<string, string>> rows, string targetKey)
        {
            foreach (IGrouping<string, Dictionary<string, string>> group in rows
                .Where(row => !String.IsNullOrWhiteSpace(GetFlat(row, "box_id")))
                .GroupBy(row => GetFlat(row, "box_id"), StringComparer.OrdinalIgnoreCase))
            {
                List<string> keys = group
                    .Select(row =>
                    {
                        string kind = GetFlat(row, "target_kind").Trim();
                        string code = GetFlat(row, "target_code").Trim();
                        return (String.IsNullOrWhiteSpace(kind) ? GuessMappingTargetKind(code) : kind.ToLowerInvariant()) + ":" + code.ToUpperInvariant();
                    })
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (keys.Count == 1 && String.Equals(keys[0], targetKey, StringComparison.OrdinalIgnoreCase))
                {
                    return group.Key;
                }
            }

            return null;
        }

        private const string MappingBoxesMutexName = "RecoQuotaData.mapping-boxes.lock";

        // mapping-boxes.jsonl 有三个写入方（推荐窗口扶正、Excel联动AI匹配、扶正训练器），
        // 都是整文件读改写，必须用跨程序集一致的命名互斥锁串行化，避免互相覆盖。
        private static void WithMappingBoxesLock(Action action)
        {
            System.Threading.Mutex mutex = new System.Threading.Mutex(false, MappingBoxesMutexName);
            bool acquired = false;
            try
            {
                try
                {
                    acquired = mutex.WaitOne(5000);
                }
                catch (System.Threading.AbandonedMutexException)
                {
                    acquired = true;
                }

                action();
            }
            finally
            {
                if (acquired)
                {
                    mutex.ReleaseMutex();
                }
                mutex.Dispose();
            }
        }

        private static Dictionary<string, object> ReadJsonObject(Dictionary<string, object> values, string key)
        {
            object value;
            return values != null && values.TryGetValue(key, out value) ? value as Dictionary<string, object> : null;
        }

        private static List<object> GetJsonList(Dictionary<string, object> values, string key)
        {
            if (values == null)
            {
                return null;
            }

            object value;
            if (!values.TryGetValue(key, out value) || value == null)
            {
                return null;
            }

            ArrayList arrayList = value as ArrayList;
            if (arrayList != null)
            {
                return arrayList.Cast<object>().ToList();
            }

            object[] objectArray = value as object[];
            return objectArray == null ? null : objectArray.ToList();
        }

        private static string ReadJsonString(Dictionary<string, object> values, string key, string fallback)
        {
            object value;
            return values != null && values.TryGetValue(key, out value) && value != null
                ? Convert.ToString(value, CultureInfo.InvariantCulture)
                : fallback;
        }

        private static int ReadJsonInt(Dictionary<string, object> values, string key, int fallback)
        {
            object value;
            if (values == null || !values.TryGetValue(key, out value) || value == null)
            {
                return fallback;
            }

            int parsed;
            return Int32.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) ? parsed : fallback;
        }

        private static bool ReadJsonBool(Dictionary<string, object> values, string key, bool fallback)
        {
            object value;
            if (values == null || !values.TryGetValue(key, out value) || value == null)
            {
                return fallback;
            }

            bool parsed;
            if (Boolean.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), out parsed))
            {
                return parsed;
            }

            int number;
            return Int32.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out number) ? number != 0 : fallback;
        }

        private static int ClampInt(int value, int min, int max)
        {
            return Math.Max(min, Math.Min(max, value));
        }

        private static string TruncateForPrompt(string text, int maxLength)
        {
            string value = (text ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }

        private static string NormalizeForSignature(string text)
        {
            return (text ?? "").Trim().ToUpperInvariant().Replace(" ", "");
        }

        private static string GetFlat(Dictionary<string, string> values, string key)
        {
            string value;
            return values != null && values.TryGetValue(key, out value) ? value ?? "" : "";
        }

        private static int ReadFlatInt(Dictionary<string, string> values, string key, int fallback)
        {
            int parsed;
            return Int32.TryParse(GetFlat(values, key), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) ? parsed : fallback;
        }

        private static Dictionary<string, string> ParseFlatJson(string line)
        {
            Dictionary<string, string> result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (String.IsNullOrWhiteSpace(line))
            {
                return result;
            }

            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                Dictionary<string, object> values = serializer.DeserializeObject(line) as Dictionary<string, object>;
                if (values == null)
                {
                    return result;
                }

                foreach (KeyValuePair<string, object> pair in values)
                {
                    result[pair.Key] = pair.Value == null ? "" : Convert.ToString(pair.Value, CultureInfo.InvariantCulture);
                }
            }
            catch
            {
            }

            return result;
        }

        private static string ToFlatJson(Dictionary<string, string> values)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append('{');
            bool first = true;
            foreach (KeyValuePair<string, string> pair in values)
            {
                if (!first)
                {
                    builder.Append(',');
                }

                first = false;
                builder.Append('"').Append(EscapeJsonText(pair.Key)).Append('"').Append(':')
                    .Append('"').Append(EscapeJsonText(pair.Value)).Append('"');
            }

            builder.Append('}');
            return builder.ToString();
        }

        private static string EscapeJsonText(string value)
        {
            StringBuilder builder = new StringBuilder();
            foreach (char ch in value ?? "")
            {
                switch (ch)
                {
                    case '\\':
                        builder.Append("\\\\");
                        break;
                    case '"':
                        builder.Append("\\\"");
                        break;
                    case '\r':
                        builder.Append("\\r");
                        break;
                    case '\n':
                        builder.Append("\\n");
                        break;
                    case '\t':
                        builder.Append("\\t");
                        break;
                    default:
                        builder.Append(ch);
                        break;
                }
            }

            return builder.ToString();
        }

        private static object GetActiveSpreadsheetApplication()
        {
            List<string> diagnostics = new List<string>();
            object app;
            if (TryGetActiveSpreadsheetApplicationByProgId(out app, diagnostics))
            {
                return app;
            }

            if (TryGetExcelApplicationByWindowObject(out app, diagnostics))
            {
                return app;
            }

            Log("GetActiveSpreadsheetApplication failed: " + String.Join(" | ", diagnostics.ToArray()));
            return null;
        }

        private static bool TryGetActiveSpreadsheetApplicationByProgId(out object app, List<string> diagnostics)
        {
            app = null;
            string[] progIds = new string[]
            {
                "ket.Application",
                "KET.Application",
                "et.Application",
                "Excel.Application"
            };

            foreach (string progId in progIds)
            {
                try
                {
                    app = Marshal.GetActiveObject(progId);
                    if (app != null)
                    {
                        Log("Excel application connected by ProgID: " + progId);
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    diagnostics.Add(progId + ": " + ex.Message);
                }
            }

            return false;
        }

        private static bool TryGetExcelApplicationByWindowObject(out object app, List<string> diagnostics)
        {
            app = null;
            List<IntPtr> excelChildWindows = new List<IntPtr>();
            EnumWindows(delegate(IntPtr hWnd, IntPtr lParam)
            {
                CollectExcelChildWindows(hWnd, excelChildWindows);
                return true;
            }, IntPtr.Zero);

            foreach (IntPtr childWindow in excelChildWindows)
            {
                try
                {
                    Guid dispatchGuid = new Guid("00020400-0000-0000-C000-000000000046");
                    object nativeObject;
                    int hr = AccessibleObjectFromWindow(childWindow, ObjIdNativeOm, ref dispatchGuid, out nativeObject);
                    if (hr != 0 || nativeObject == null)
                    {
                        diagnostics.Add("AccessibleObjectFromWindow " + childWindow.ToString("X") + ": hr=" + hr.ToString(CultureInfo.InvariantCulture));
                        continue;
                    }

                    dynamic window = nativeObject;
                    app = window.Application;
                    if (app != null)
                    {
                        Log("Excel application connected by window object.");
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    diagnostics.Add("WindowObject " + childWindow.ToString("X") + ": " + ex.Message);
                }
            }

            if (excelChildWindows.Count == 0)
            {
                diagnostics.Add("No EXCEL7 child window found.");
            }

            return false;
        }

        private static void CollectExcelChildWindows(IntPtr parent, List<IntPtr> excelChildWindows)
        {
            EnumChildWindows(parent, delegate(IntPtr hWnd, IntPtr lParam)
            {
                string className = GetWindowClassName(hWnd);
                if (String.Equals(className, "EXCEL7", StringComparison.OrdinalIgnoreCase))
                {
                    excelChildWindows.Add(hWnd);
                }

                CollectExcelChildWindows(hWnd, excelChildWindows);
                return true;
            }, IntPtr.Zero);
        }

        private static string GetWindowClassName(IntPtr hWnd)
        {
            StringBuilder builder = new StringBuilder(256);
            int length = GetClassName(hWnd, builder, builder.Capacity);
            return length <= 0 ? String.Empty : builder.ToString();
        }

        private static bool PromptExcelCell(IWin32Window owner, string reason, out ExcelCellAddress cell)
        {
            cell = null;
            using (Form dialog = new Form())
            using (Label info = new Label())
            using (Label fileLabel = new Label())
            using (TextBox fileText = new TextBox())
            using (Button browse = new Button())
            using (Label sheetLabel = new Label())
            using (ComboBox sheetBox = new ComboBox())
            using (Label cellLabel = new Label())
            using (TextBox cellText = new TextBox())
            using (DataGridView preview = new DataGridView())
            using (Button ok = new Button())
            using (Button cancel = new Button())
            {
                dialog.Text = "手动绑定Excel工程量";
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.MinimizeBox = false;
                dialog.MaximizeBox = false;
                dialog.ClientSize = new System.Drawing.Size(820, 520);

                info.Left = 12;
                info.Top = 12;
                info.Width = 590;
                info.Height = 34;
                info.Text = reason + "；也可以在这里手动选择 .xlsx 文件并填写工作表和单元格。";

                fileLabel.Text = "Excel文件";
                fileLabel.Left = 12;
                fileLabel.Top = 58;
                fileLabel.Width = 80;

                fileText.Left = 94;
                fileText.Top = 55;
                fileText.Width = 600;

                browse.Text = "选择";
                browse.Left = 704;
                browse.Top = 53;
                browse.Width = 80;
                browse.Click += delegate
                {
                    using (OpenFileDialog chooser = new OpenFileDialog())
                    {
                        chooser.Title = "选择工程数量Excel文件";
                        chooser.Filter = "Excel工作簿 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
                        if (chooser.ShowDialog(dialog) == DialogResult.OK)
                        {
                            fileText.Text = chooser.FileName;
                            LoadSheetNamesIntoCombo(chooser.FileName, sheetBox);
                            LoadPreviewGrid(chooser.FileName, Convert.ToString(sheetBox.Text, CultureInfo.InvariantCulture), preview);
                        }
                    }
                };

                sheetLabel.Text = "工作表";
                sheetLabel.Left = 12;
                sheetLabel.Top = 96;
                sheetLabel.Width = 80;

                sheetBox.Left = 94;
                sheetBox.Top = 93;
                sheetBox.Width = 360;
                sheetBox.DropDownStyle = ComboBoxStyle.DropDown;
                sheetBox.SelectedIndexChanged += delegate
                {
                    LoadPreviewGrid(fileText.Text.Trim(), Convert.ToString(sheetBox.Text, CultureInfo.InvariantCulture), preview);
                };

                cellLabel.Text = "单元格";
                cellLabel.Left = 475;
                cellLabel.Top = 96;
                cellLabel.Width = 60;

                cellText.Left = 535;
                cellText.Top = 93;
                cellText.Width = 80;
                cellText.Text = "E4";

                preview.Left = 12;
                preview.Top = 132;
                preview.Width = 796;
                preview.Height = 320;
                preview.ReadOnly = true;
                preview.AllowUserToAddRows = false;
                preview.AllowUserToDeleteRows = false;
                preview.SelectionMode = DataGridViewSelectionMode.CellSelect;
                preview.RowHeadersWidth = 54;
                preview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                preview.CellClick += delegate(object sender, DataGridViewCellEventArgs e)
                {
                    if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                    {
                        cellText.Text = ColumnNumberToName(e.ColumnIndex + 1) + (e.RowIndex + 1).ToString(CultureInfo.InvariantCulture);
                        dialog.DialogResult = DialogResult.OK;
                        dialog.Close();
                    }
                };

                ok.Text = "确定";
                ok.Left = 620;
                ok.Top = 468;
                ok.Width = 80;
                ok.DialogResult = DialogResult.OK;

                cancel.Text = "取消";
                cancel.Left = 706;
                cancel.Top = 468;
                cancel.Width = 80;
                cancel.DialogResult = DialogResult.Cancel;

                dialog.Controls.Add(info);
                dialog.Controls.Add(fileLabel);
                dialog.Controls.Add(fileText);
                dialog.Controls.Add(browse);
                dialog.Controls.Add(sheetLabel);
                dialog.Controls.Add(sheetBox);
                dialog.Controls.Add(cellLabel);
                dialog.Controls.Add(cellText);
                dialog.Controls.Add(preview);
                dialog.Controls.Add(ok);
                dialog.Controls.Add(cancel);
                dialog.AcceptButton = ok;
                dialog.CancelButton = cancel;

                while (dialog.ShowDialog(owner) == DialogResult.OK)
                {
                    string path = fileText.Text.Trim();
                    string sheet = Convert.ToString(sheetBox.Text, CultureInfo.InvariantCulture).Trim();
                    string address = cellText.Text.Trim().ToUpperInvariant();
                    if (!File.Exists(path))
                    {
                        MessageBox.Show(owner, "请选择存在的 Excel 文件。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        continue;
                    }

                    if (String.IsNullOrEmpty(sheet) || String.IsNullOrEmpty(address))
                    {
                        MessageBox.Show(owner, "请填写工作表和单元格地址，例如 E4。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        continue;
                    }

                    string displayValue;
                    string readError;
                    if (!TryReadWorkbookCellValue(path, sheet, address, out displayValue, out readError))
                    {
                        DialogResult result = MessageBox.Show(owner, "当前无法读取该单元格：" + readError + Environment.NewLine + "仍然保存绑定吗？", "Excel联动", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result != DialogResult.Yes)
                        {
                            continue;
                        }
                    }

                    cell = new ExcelCellAddress();
                    cell.WorkbookPath = path;
                    cell.WorksheetName = sheet;
                    cell.CellAddress = address;
                    cell.DisplayValue = displayValue ?? "";
                    cell.SelectionAddresses = new List<string> { address };
                    return true;
                }
            }

            return false;
        }

        private static void LoadSheetNamesIntoCombo(string path, ComboBox sheetBox)
        {
            sheetBox.Items.Clear();
            string error;
            foreach (string name in GetSheetNamesByNpoi(path, out error))
            {
                sheetBox.Items.Add(name);
            }

            if (sheetBox.Items.Count == 0)
            {
                foreach (string name in GetXlsxSheetNames(path, out error))
                {
                    sheetBox.Items.Add(name);
                }
            }

            if (sheetBox.Items.Count > 0)
            {
                sheetBox.SelectedIndex = 0;
            }
        }

        private static void LoadPreviewGrid(string path, string sheetName, DataGridView preview)
        {
            LoadPreviewGrid(path, sheetName, preview, "A1", 40, 12);
        }

        private static void LoadPreviewGrid(string path, string sheetName, DataGridView preview, string startAddress, int maxRows, int maxCols)
        {
            preview.Columns.Clear();
            preview.Rows.Clear();
            if (String.IsNullOrEmpty(path) || String.IsNullOrEmpty(sheetName) || !File.Exists(path))
            {
                return;
            }

            CellRef start;
            if (!TryParseCellAddress(startAddress, out start))
            {
                start = new CellRef { Column = 1, Row = 1 };
            }

            for (int col = 1; col <= maxCols; col++)
            {
                string name = ColumnNumberToName(start.Column + col - 1);
                DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
                column.Name = name;
                column.HeaderText = name;
                column.Width = col == 2 || col == 3 ? 180 : 80;
                column.FillWeight = col == 2 || col == 3 ? 180 : 80;
                preview.Columns.Add(column);
            }

            for (int row = 1; row <= maxRows; row++)
            {
                int rowIndex = preview.Rows.Add();
                preview.Rows[rowIndex].HeaderCell.Value = (start.Row + row - 1).ToString(CultureInfo.InvariantCulture);
            }

            int endRow = start.Row + maxRows - 1;
            int endCol = start.Column + maxCols - 1;
            Dictionary<string, string> values = ReadWorkbookSheetCells(path, sheetName, endRow, endCol);
            foreach (KeyValuePair<string, string> pair in values)
            {
                CellRef cell;
                if (TryParseCellAddress(pair.Key, out cell)
                    && cell.Row >= start.Row
                    && cell.Row <= endRow
                    && cell.Column >= start.Column
                    && cell.Column <= endCol)
                {
                    preview.Rows[cell.Row - start.Row].Cells[cell.Column - start.Column].Value = pair.Value;
                }
            }
        }

        private static bool TryReadExcelCellValue(ExcelQuotaLink link, out string valueText, out decimal quantity, out string error)
        {
            valueText = null;
            quantity = 0;
            error = null;

            dynamic excel = null;
            try
            {
                excel = GetActiveSpreadsheetApplication();
                if (excel == null)
                {
                    return TryReadExcelCellValueFromFile(link, out valueText, out quantity, out error);
                }
                dynamic workbooks = excel.Workbooks;
                dynamic targetWorkbook = null;
                int count = Convert.ToInt32(workbooks.Count, CultureInfo.InvariantCulture);
                for (int i = 1; i <= count; i++)
                {
                    dynamic workbook = workbooks.Item(i);
                    string fullName = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture);
                    if (String.Equals(Path.GetFullPath(fullName), Path.GetFullPath(link.ExcelPath), StringComparison.OrdinalIgnoreCase))
                    {
                        targetWorkbook = workbook;
                        break;
                    }
                }

                if (targetWorkbook == null)
                {
                    return TryReadExcelCellValueFromFile(link, out valueText, out quantity, out error);
                }

                dynamic sheet = targetWorkbook.Worksheets[link.WorksheetName];
                string liveExpression = String.IsNullOrWhiteSpace(link.Expression) ? link.CellAddress : link.Expression;
                string normalizedExpression = NormalizeCellAddress(liveExpression).Replace("×", "*").Replace("（", "(").Replace("）", ")");
                CellRef onlyCell;
                if (TryParseCellAddress(normalizedExpression, out onlyCell))
                {
                    dynamic range = sheet.Range[normalizedExpression];
                    object rawValue = range.Value2;
                    valueText = ExcelValueToText(rawValue);
                    if (String.IsNullOrWhiteSpace(valueText))
                    {
                        error = "Excel 单元格为空";
                        return false;
                    }

                    if (!TryEvaluateDecimal(valueText, out quantity, out error))
                    {
                        error = "Excel 单元格值无法计算：" + error;
                        return false;
                    }

                    return true;
                }

                string resolved = ResolveLiveExpression(sheet, normalizedExpression, out error);
                if (resolved == null)
                {
                    return TryReadExcelCellValueFromFile(link, out valueText, out quantity, out error);
                }

                if (!TryEvaluateDecimal(resolved, out quantity, out error))
                {
                    error = "表达式计算失败：" + error;
                    return false;
                }

                valueText = quantity.ToString(CultureInfo.InvariantCulture);
                return true;
            }
            catch (COMException ex)
            {
                if (TryReadExcelCellValueFromFile(link, out valueText, out quantity, out error))
                {
                    return true;
                }

                error = "读取 Excel 失败：" + ex.Message + "；文件读取也失败：" + error;
                return false;
            }
            catch (Exception ex)
            {
                if (TryReadExcelCellValueFromFile(link, out valueText, out quantity, out error))
                {
                    return true;
                }

                error = "数值无法计算：" + ex.Message + "；文件读取也失败：" + error;
                return false;
            }
        }

        private static bool TryReadExcelCellValueFromFile(ExcelQuotaLink link, out string valueText, out decimal quantity, out string error)
        {
            valueText = null;
            quantity = 0;
            error = null;
            if (String.IsNullOrEmpty(link.ExcelPath) || !File.Exists(link.ExcelPath))
            {
                error = "Excel 文件不存在";
                return false;
            }

            string expression = String.IsNullOrWhiteSpace(link.Expression) ? link.CellAddress : link.Expression;
            if (!TryEvaluateWorkbookExpression(link.ExcelPath, link.WorksheetName, expression, out valueText, out quantity, out error))
            {
                return false;
            }

            if (String.IsNullOrWhiteSpace(valueText))
            {
                error = "Excel 单元格为空";
                return false;
            }

            return true;
        }

        private static bool TryEvaluateWorkbookExpression(string path, string sheetName, string expression, out string valueText, out decimal quantity, out string error)
        {
            valueText = null;
            quantity = 0;
            error = null;

            if (String.IsNullOrWhiteSpace(expression))
            {
                error = "Excel 表达式为空";
                return false;
            }

            string normalized = NormalizeCellAddress(expression).Replace("×", "*").Replace("（", "(").Replace("）", ")");
            CellRef onlyCell;
            if (TryParseCellAddress(normalized, out onlyCell))
            {
                string cellValue;
                if (!TryReadWorkbookCellValue(path, sheetName, normalized, out cellValue, out error))
                {
                    return false;
                }

                if (!TryEvaluateDecimal(cellValue, out quantity, out error))
                {
                    error = normalized + " 单元格值无法计算：" + error;
                    return false;
                }

                valueText = cellValue;
                return true;
            }

            string resolved = ResolveWorkbookExpression(path, sheetName, normalized, out error);
            if (resolved == null)
            {
                return false;
            }

            if (!TryEvaluateDecimal(resolved, out quantity, out error))
            {
                error = "表达式计算失败：" + error;
                return false;
            }

            valueText = quantity.ToString(CultureInfo.InvariantCulture);
            return true;
        }

        private static bool TryEvaluateDecimal(string expression, out decimal value, out string error)
        {
            value = 0;
            error = null;

            try
            {
                value = EvaluateDecimal(expression);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static string ResolveWorkbookExpression(string path, string sheetName, string expression, out string error)
        {
            error = null;
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < expression.Length;)
            {
                char ch = expression[i];
                if (Char.IsLetter(ch))
                {
                    int start = i;
                    while (i < expression.Length && Char.IsLetter(expression[i]))
                    {
                        i++;
                    }
                    while (i < expression.Length && Char.IsDigit(expression[i]))
                    {
                        i++;
                    }

                    string token = expression.Substring(start, i - start).ToUpperInvariant();
                    CellRef cell;
                    if (!TryParseCellAddress(token, out cell))
                    {
                        error = "表达式里包含无法识别的单元格：" + token;
                        return null;
                    }

                    string cellValue;
                    if (!TryReadWorkbookCellValue(path, sheetName, token, out cellValue, out error))
                    {
                        error = token + " 读取失败：" + error;
                        return null;
                    }

                    decimal parsed;
                    string parseError;
                    if (!TryEvaluateDecimal(cellValue, out parsed, out parseError))
                    {
                        error = token + " 单元格值无法计算：" + parseError;
                        return null;
                    }

                    builder.Append(parsed.ToString(CultureInfo.InvariantCulture));
                    continue;
                }

                if ("0123456789.+-*/() ".IndexOf(ch) >= 0)
                {
                    builder.Append(ch);
                    i++;
                    continue;
                }

                error = "表达式里包含不支持的字符：" + ch;
                return null;
            }

            return builder.ToString();
        }

        private static string ResolveLiveExpression(dynamic sheet, string expression, out string error)
        {
            error = null;
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < expression.Length;)
            {
                char ch = expression[i];
                if (Char.IsLetter(ch))
                {
                    int start = i;
                    while (i < expression.Length && Char.IsLetter(expression[i]))
                    {
                        i++;
                    }
                    while (i < expression.Length && Char.IsDigit(expression[i]))
                    {
                        i++;
                    }

                    string token = expression.Substring(start, i - start).ToUpperInvariant();
                    CellRef cell;
                    if (!TryParseCellAddress(token, out cell))
                    {
                        error = "表达式里包含无法识别的单元格：" + token;
                        return null;
                    }

                    string cellValue;
                    try
                    {
                        dynamic range = sheet.Range[token];
                        cellValue = ExcelValueToText((object)range.Value2);
                    }
                    catch (Exception ex)
                    {
                        error = token + " 读取失败：" + ex.Message;
                        return null;
                    }

                    decimal parsed;
                    string parseError;
                    if (!TryEvaluateDecimal(cellValue, out parsed, out parseError))
                    {
                        error = token + " 单元格值无法计算：" + parseError;
                        return null;
                    }

                    builder.Append(parsed.ToString(CultureInfo.InvariantCulture));
                    continue;
                }

                if ("0123456789.+-*/() ".IndexOf(ch) >= 0)
                {
                    builder.Append(ch);
                    i++;
                    continue;
                }

                error = "表达式里包含不支持的字符：" + ch;
                return null;
            }

            return builder.ToString();
        }

        private static bool TryReadWorkbookCellValue(string path, string sheetName, string cellAddress, out string valueText, out string error)
        {
            if (TryReadCellValueByNpoi(path, sheetName, cellAddress, out valueText, out error))
            {
                return true;
            }

            string npoiError = error;
            if (TryReadXlsxCellValue(path, sheetName, cellAddress, out valueText, out error))
            {
                return true;
            }

            error = "NPOI读取失败：" + npoiError + "；直接读取失败：" + error;
            return false;
        }

        private static bool TryReadXlsxCellValue(string path, string sheetName, string cellAddress, out string valueText, out string error)
        {
            valueText = null;
            error = null;
            try
            {
                using (ZipArchive archive = OpenZipArchiveShared(path))
                {
                    List<string> sharedStrings = ReadSharedStrings(archive);
                    string sheetPath = ResolveSheetPath(archive, sheetName);
                    if (String.IsNullOrEmpty(sheetPath))
                    {
                        error = "找不到工作表：" + sheetName;
                        return false;
                    }

                    ZipArchiveEntry sheetEntry = archive.GetEntry(sheetPath);
                    if (sheetEntry == null)
                    {
                        error = "找不到工作表数据：" + sheetPath;
                        return false;
                    }

                    using (Stream stream = sheetEntry.Open())
                    using (XmlReader reader = XmlReader.Create(stream))
                    {
                        string target = NormalizeCellAddress(cellAddress);
                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c")
                            {
                                string reference = NormalizeCellAddress(reader.GetAttribute("r"));
                                if (String.Equals(reference, target, StringComparison.OrdinalIgnoreCase))
                                {
                                    valueText = ReadCellValue(reader.ReadSubtree(), reader.GetAttribute("t"), sharedStrings);
                                    return true;
                                }
                            }
                        }
                    }
                }

                error = "找不到单元格：" + cellAddress;
                return false;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static Dictionary<string, string> ReadXlsxSheetCells(string path, string sheetName, int maxRows, int maxCols)
        {
            Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (ZipArchive archive = OpenZipArchiveShared(path))
                {
                    List<string> sharedStrings = ReadSharedStrings(archive);
                    string sheetPath = ResolveSheetPath(archive, sheetName);
                    if (String.IsNullOrEmpty(sheetPath))
                    {
                        return values;
                    }

                    ZipArchiveEntry sheetEntry = archive.GetEntry(sheetPath);
                    if (sheetEntry == null)
                    {
                        return values;
                    }

                    using (Stream stream = sheetEntry.Open())
                    using (XmlReader reader = XmlReader.Create(stream))
                    {
                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c")
                            {
                                string reference = NormalizeCellAddress(reader.GetAttribute("r"));
                                CellRef cell;
                                if (!TryParseCellAddress(reference, out cell) || cell.Row > maxRows || cell.Column > maxCols)
                                {
                                    continue;
                                }

                                values[reference] = ReadCellValue(reader.ReadSubtree(), reader.GetAttribute("t"), sharedStrings);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("ReadXlsxSheetCells failed: " + ex.Message);
            }

            return values;
        }

        private static Dictionary<string, string> ReadWorkbookSheetCells(string path, string sheetName, int maxRows, int maxCols)
        {
            Dictionary<string, string> values = ReadSheetCellsByNpoi(path, sheetName, maxRows, maxCols);
            if (values.Count > 0)
            {
                return values;
            }

            return ReadXlsxSheetCells(path, sheetName, maxRows, maxCols);
        }

        private static IEnumerable<string> GetSheetNamesByNpoi(string path, out string error)
        {
            error = null;
            List<string> names = new List<string>();
            try
            {
                using (Stream stream = OpenWorkbookStreamShared(path))
                {
                    IWorkbook workbook = WorkbookFactory.Create(stream);
                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        names.Add(workbook.GetSheetName(i));
                    }
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
            }

            return names;
        }

        private static bool TryReadCellValueByNpoi(string path, string sheetName, string cellAddress, out string valueText, out string error)
        {
            valueText = null;
            error = null;
            try
            {
                CellRef cellRef;
                if (!TryParseCellAddress(cellAddress, out cellRef))
                {
                    error = "单元格地址不正确：" + cellAddress;
                    return false;
                }

                using (Stream stream = OpenWorkbookStreamShared(path))
                {
                    IWorkbook workbook = WorkbookFactory.Create(stream);
                    ISheet sheet = workbook.GetSheet(sheetName);
                    if (sheet == null)
                    {
                        error = "找不到工作表：" + sheetName;
                        return false;
                    }

                    IRow row = sheet.GetRow(cellRef.Row - 1);
                    ICell cell = row == null ? null : row.GetCell(cellRef.Column - 1);
                    valueText = NpoiCellToText(cell);
                    return true;
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static Dictionary<string, string> ReadSheetCellsByNpoi(string path, string sheetName, int maxRows, int maxCols)
        {
            Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (Stream stream = OpenWorkbookStreamShared(path))
                {
                    IWorkbook workbook = WorkbookFactory.Create(stream);
                    ISheet sheet = workbook.GetSheet(sheetName);
                    if (sheet == null)
                    {
                        return values;
                    }

                    for (int rowIndex = 0; rowIndex < maxRows; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        if (row == null)
                        {
                            continue;
                        }

                        for (int colIndex = 0; colIndex < maxCols; colIndex++)
                        {
                            ICell cell = row.GetCell(colIndex);
                            string text = NpoiCellToText(cell);
                            if (!String.IsNullOrEmpty(text))
                            {
                                values[ColumnNumberToName(colIndex + 1) + (rowIndex + 1).ToString(CultureInfo.InvariantCulture)] = text;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("ReadSheetCellsByNpoi failed: " + ex.Message);
            }

            return values;
        }

        private static string NpoiCellToText(ICell cell)
        {
            if (cell == null)
            {
                return "";
            }

            try
            {
                switch (cell.CellType)
                {
                    case CellType.Numeric:
                        return cell.NumericCellValue.ToString("G15", CultureInfo.InvariantCulture);
                    case CellType.String:
                        return cell.StringCellValue == null ? "" : cell.StringCellValue.Trim();
                    case CellType.Boolean:
                        return cell.BooleanCellValue ? "TRUE" : "FALSE";
                    case CellType.Formula:
                        switch (cell.CachedFormulaResultType)
                        {
                            case CellType.Numeric:
                                return cell.NumericCellValue.ToString("G15", CultureInfo.InvariantCulture);
                            case CellType.String:
                                return cell.StringCellValue == null ? "" : cell.StringCellValue.Trim();
                            case CellType.Boolean:
                                return cell.BooleanCellValue ? "TRUE" : "FALSE";
                            default:
                                return cell.ToString().Trim();
                        }
                    default:
                        return cell.ToString().Trim();
                }
            }
            catch
            {
                return cell.ToString().Trim();
            }
        }

        private static IEnumerable<string> GetSheetNamesByExcelCom(string path, out string error)
        {
            error = null;
            List<string> names = new List<string>();
            object excel = null;
            object workbook = null;
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    error = "本机没有可用的 Microsoft Excel COM。";
                    return names;
                }

                excel = Activator.CreateInstance(excelType);
                dynamic app = excel;
                app.Visible = false;
                app.DisplayAlerts = false;
                workbook = app.Workbooks.Open(path, 0, true);
                dynamic wb = workbook;
                int count = Convert.ToInt32(wb.Worksheets.Count, CultureInfo.InvariantCulture);
                for (int i = 1; i <= count; i++)
                {
                    dynamic sheet = wb.Worksheets.Item(i);
                    names.Add(Convert.ToString(sheet.Name, CultureInfo.InvariantCulture));
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
            }
            finally
            {
                CloseExcelComWorkbook(excel, workbook);
            }

            return names;
        }

        private static bool TryReadCellValueByExcelCom(string path, string sheetName, string cellAddress, out string valueText, out string error)
        {
            valueText = null;
            error = null;
            object excel = null;
            object workbook = null;
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    error = "本机没有可用的 Microsoft Excel COM。";
                    return false;
                }

                excel = Activator.CreateInstance(excelType);
                dynamic app = excel;
                app.Visible = false;
                app.DisplayAlerts = false;
                workbook = app.Workbooks.Open(path, 0, true);
                dynamic wb = workbook;
                dynamic sheet = wb.Worksheets[sheetName];
                dynamic range = sheet.Range[cellAddress];
                valueText = ExcelValueToText(range.Value2);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
            finally
            {
                CloseExcelComWorkbook(excel, workbook);
            }
        }

        private static Dictionary<string, string> ReadSheetCellsByExcelCom(string path, string sheetName, int maxRows, int maxCols)
        {
            Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            object excel = null;
            object workbook = null;
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    return values;
                }

                excel = Activator.CreateInstance(excelType);
                dynamic app = excel;
                app.Visible = false;
                app.DisplayAlerts = false;
                workbook = app.Workbooks.Open(path, 0, true);
                dynamic wb = workbook;
                dynamic sheet = wb.Worksheets[sheetName];
                for (int row = 1; row <= maxRows; row++)
                {
                    for (int col = 1; col <= maxCols; col++)
                    {
                        object value = sheet.Cells[row, col].Value2;
                        if (value != null)
                        {
                            values[ColumnNumberToName(col) + row.ToString(CultureInfo.InvariantCulture)] = ExcelValueToText(value);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("ReadSheetCellsByExcelCom failed: " + ex.Message);
            }
            finally
            {
                CloseExcelComWorkbook(excel, workbook);
            }

            return values;
        }

        private static void CloseExcelComWorkbook(object excel, object workbook)
        {
            try
            {
                if (workbook != null)
                {
                    dynamic wb = workbook;
                    wb.Close(false);
                }
            }
            catch
            {
            }

            try
            {
                if (excel != null)
                {
                    dynamic app = excel;
                    app.Quit();
                }
            }
            catch
            {
            }

            try
            {
                if (workbook != null)
                {
                    Marshal.FinalReleaseComObject(workbook);
                }
            }
            catch
            {
            }

            try
            {
                if (excel != null)
                {
                    Marshal.FinalReleaseComObject(excel);
                }
            }
            catch
            {
            }
        }

        private static IEnumerable<string> GetXlsxSheetNames(string path, out string error)
        {
            error = null;
            List<string> names = new List<string>();
            try
            {
                using (ZipArchive archive = OpenZipArchiveShared(path))
                {
                    ZipArchiveEntry workbook = archive.GetEntry("xl/workbook.xml");
                    if (workbook == null)
                    {
                        error = "不是有效的 .xlsx 文件";
                        return names;
                    }

                    XDocument doc;
                    using (Stream stream = workbook.Open())
                    {
                        doc = XDocument.Load(stream);
                    }

                    XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                    foreach (XElement sheet in doc.Descendants(ns + "sheet"))
                    {
                        XAttribute attr = sheet.Attribute("name");
                        if (attr != null)
                        {
                            names.Add(attr.Value);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
            }

            return names;
        }

        private static string ResolveSheetPath(ZipArchive archive, string sheetName)
        {
            ZipArchiveEntry workbook = archive.GetEntry("xl/workbook.xml");
            ZipArchiveEntry rels = archive.GetEntry("xl/_rels/workbook.xml.rels");
            if (workbook == null || rels == null)
            {
                return null;
            }

            XDocument workbookDoc;
            XDocument relsDoc;
            using (Stream stream = workbook.Open())
            {
                workbookDoc = XDocument.Load(stream);
            }

            using (Stream stream = rels.Open())
            {
                relsDoc = XDocument.Load(stream);
            }

            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XElement targetSheet = workbookDoc.Descendants(ns + "sheet")
                .FirstOrDefault(s => String.Equals((string)s.Attribute("name"), sheetName, StringComparison.OrdinalIgnoreCase));
            if (targetSheet == null)
            {
                return null;
            }

            string relId = (string)targetSheet.Attribute(relNs + "id");
            if (String.IsNullOrEmpty(relId))
            {
                return null;
            }

            XNamespace packageRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
            XElement rel = relsDoc.Descendants(packageRelNs + "Relationship")
                .FirstOrDefault(r => String.Equals((string)r.Attribute("Id"), relId, StringComparison.OrdinalIgnoreCase));
            if (rel == null)
            {
                return null;
            }

            string target = ((string)rel.Attribute("Target") ?? "").Replace('\\', '/');
            if (target.StartsWith("/", StringComparison.Ordinal))
            {
                target = target.TrimStart('/');
            }
            else if (!target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
            {
                target = "xl/" + target;
            }

            return target;
        }

        private static List<string> ReadSharedStrings(ZipArchive archive)
        {
            List<string> values = new List<string>();
            ZipArchiveEntry entry = archive.GetEntry("xl/sharedStrings.xml");
            if (entry == null)
            {
                return values;
            }

            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            using (Stream stream = entry.Open())
            {
                XDocument doc = XDocument.Load(stream);
                foreach (XElement si in doc.Descendants(ns + "si"))
                {
                    values.Add(String.Concat(si.Descendants(ns + "t").Select(t => (string)t)));
                }
            }

            return values;
        }

        private static Stream OpenWorkbookStreamShared(string path)
        {
            try
            {
                return new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            }
            catch
            {
                string temp = Path.Combine(Path.GetTempPath(), "RecoExcelLink_" + Guid.NewGuid().ToString("N") + Path.GetExtension(path));
                File.Copy(path, temp, true);
                return new DeleteOnCloseFileStream(temp);
            }
        }

        private static ZipArchive OpenZipArchiveShared(string path)
        {
            return new ZipArchive(OpenWorkbookStreamShared(path), ZipArchiveMode.Read);
        }

        private sealed class DeleteOnCloseFileStream : FileStream
        {
            private readonly string path;

            public DeleteOnCloseFileStream(string path)
                : base(path, FileMode.Open, FileAccess.Read, FileShare.Read)
            {
                this.path = path;
            }

            protected override void Dispose(bool disposing)
            {
                base.Dispose(disposing);
                try
                {
                    File.Delete(path);
                }
                catch
                {
                }
            }
        }

        private static string ReadCellValue(XmlReader reader, string type, List<string> sharedStrings)
        {
            string value = null;
            using (reader)
            {
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element && (reader.LocalName == "v" || reader.LocalName == "t"))
                    {
                        value = reader.ReadElementContentAsString();
                        if (reader.LocalName == "t")
                        {
                            break;
                        }
                    }
                }
            }

            if (value == null)
            {
                return "";
            }

            if (type == "s")
            {
                int index;
                if (Int32.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out index) && index >= 0 && index < sharedStrings.Count)
                {
                    return sharedStrings[index];
                }
            }

            if (type == "b")
            {
                return value == "1" ? "TRUE" : "FALSE";
            }

            return value.Trim();
        }

        private static string NormalizeCellAddress(string address)
        {
            if (String.IsNullOrEmpty(address))
            {
                return "";
            }

            return address.Replace("$", "").Trim().ToUpperInvariant();
        }

        private static bool TryParseCellAddress(string address, out CellRef cell)
        {
            cell = new CellRef();
            string normalized = NormalizeCellAddress(address);
            if (String.IsNullOrEmpty(normalized))
            {
                return false;
            }

            int index = 0;
            int column = 0;
            while (index < normalized.Length && normalized[index] >= 'A' && normalized[index] <= 'Z')
            {
                column = column * 26 + (normalized[index] - 'A' + 1);
                index++;
            }

            int row;
            if (column <= 0 || index >= normalized.Length || !Int32.TryParse(normalized.Substring(index), NumberStyles.Integer, CultureInfo.InvariantCulture, out row) || row <= 0)
            {
                return false;
            }

            cell.Column = column;
            cell.Row = row;
            return true;
        }

        private static string ExtractFirstCellAddress(string expression)
        {
            if (String.IsNullOrEmpty(expression))
            {
                return null;
            }

            string normalized = NormalizeCellAddress(expression);
            for (int i = 0; i < normalized.Length; i++)
            {
                if (normalized[i] >= 'A' && normalized[i] <= 'Z')
                {
                    int start = i;
                    while (i < normalized.Length && normalized[i] >= 'A' && normalized[i] <= 'Z')
                    {
                        i++;
                    }
                    while (i < normalized.Length && Char.IsDigit(normalized[i]))
                    {
                        i++;
                    }

                    string candidate = normalized.Substring(start, i - start);
                    CellRef parsed;
                    if (TryParseCellAddress(candidate, out parsed))
                    {
                        return candidate;
                    }
                }
            }

            return null;
        }

        private static string ColumnNumberToName(int column)
        {
            if (column <= 0)
            {
                return "";
            }

            StringBuilder builder = new StringBuilder();
            int value = column;
            while (value > 0)
            {
                value--;
                builder.Insert(0, (char)('A' + (value % 26)));
                value /= 26;
            }

            return builder.ToString();
        }

        private struct CellRef
        {
            public int Column;
            public int Row;
        }

        private static string ExcelValueToText(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return "";
            }

            if (value is double || value is float)
            {
                return Convert.ToDouble(value, CultureInfo.InvariantCulture).ToString("G15", CultureInfo.InvariantCulture);
            }

            if (value is decimal)
            {
                return ((decimal)value).ToString(CultureInfo.InvariantCulture);
            }

            if (value is int || value is long || value is short || value is byte)
            {
                return Convert.ToString(value, CultureInfo.InvariantCulture);
            }

            return Convert.ToString(value, CultureInfo.CurrentCulture).Trim();
        }

        private static SyncSummary SyncExcelLinks(Form mainForm, bool manual)
        {
            SyncSummary summary = new SyncSummary();
            SqlConnection conn = GetProjectConnection(mainForm);
            if (conn == null)
            {
                summary.Message = "没有找到当前项目数据库连接。";
                return summary;
            }

            ExcelLinkStore store = LoadStore(conn);
            if (store.Links.Count == 0)
            {
                summary.Message = "当前项目还没有 Excel 联动绑定。";
                return summary;
            }

            List<PendingSync> pending = new List<PendingSync>();
            foreach (ExcelQuotaLink link in store.Links)
            {
                string valueText;
                decimal quantity;
                string error;
                if (!TryReadExcelCellValue(link, out valueText, out quantity, out error))
                {
                    link.LastStatus = error;
                    link.UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                    summary.Skipped++;
                    continue;
                }

                PendingSync item = new PendingSync();
                item.Link = link;
                item.ValueText = valueText;
                item.Quantity = quantity;
                pending.Add(item);
            }

            if (pending.Count == 0)
            {
                SaveStore(conn, store);
                summary.Message = "没有可同步的有效绑定。";
                return summary;
            }

            EnsureOpen(conn);
            using (SqlTransaction transaction = conn.BeginTransaction())
            using (SqlCommand select = conn.CreateCommand())
            using (SqlCommand update = conn.CreateCommand())
            {
                select.Transaction = transaction;
                select.CommandText = "select 工程数量输入 from 定额输入 where 定额序号=@id";
                select.Parameters.Add("@id", SqlDbType.BigInt);

                update.Transaction = transaction;
                update.CommandText = "update 定额输入 set 工程数量输入=@value, 工程数量=@quantity where 定额序号=@id";
                update.Parameters.Add("@value", SqlDbType.NVarChar, 200);
                update.Parameters.Add("@quantity", SqlDbType.Float);
                update.Parameters.Add("@id", SqlDbType.BigInt);

                try
                {
                    foreach (PendingSync item in pending)
                    {
                        select.Parameters["@id"].Value = item.Link.QuotaSequence;
                        object oldValue = select.ExecuteScalar();
                        if (oldValue == null)
                        {
                            item.Link.LastStatus = "定额行不存在";
                            summary.Skipped++;
                            continue;
                        }

                        update.Parameters["@value"].Value = item.ValueText;
                        update.Parameters["@quantity"].Value = Convert.ToDouble(item.Quantity, CultureInfo.InvariantCulture);
                        update.Parameters["@id"].Value = item.Link.QuotaSequence;
                        int changed = update.ExecuteNonQuery();
                        if (changed > 0)
                        {
                            item.Link.LastSyncValue = item.ValueText;
                            item.Link.LastStatus = "已同步";
                            item.Link.UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                            AppendExcelSyncLog(conn, item.Link, Convert.ToString(oldValue, CultureInfo.InvariantCulture), item.ValueText, manual);
                            summary.Changed += changed;
                        }
                        else
                        {
                            item.Link.LastStatus = "定额行不存在";
                            summary.Skipped++;
                        }
                    }

                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
            }

            SaveStore(conn, store);
            RefreshCurrentQuotaGrid(mainForm);
            RefreshExcelLinkPanel(mainForm);
            summary.Message = "已同步 " + summary.Changed.ToString(CultureInfo.InvariantCulture) + " 条，跳过 " + summary.Skipped.ToString(CultureInfo.InvariantCulture) + " 条。";
            return summary;
        }

        private static ExcelLinkStore LoadStore(SqlConnection conn)
        {
            string path = GetStorePath(conn);
            if (!File.Exists(path))
            {
                return new ExcelLinkStore();
            }

            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(ExcelLinkStore));
                using (FileStream stream = File.OpenRead(path))
                {
                    ExcelLinkStore store = serializer.Deserialize(stream) as ExcelLinkStore;
                    return store ?? new ExcelLinkStore();
                }
            }
            catch (Exception ex)
            {
                Log("Load Excel link store failed: " + ex);
                return new ExcelLinkStore();
            }
        }

        private static void SaveStore(SqlConnection conn, ExcelLinkStore store)
        {
            string path = GetStorePath(conn);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            XmlSerializer serializer = new XmlSerializer(typeof(ExcelLinkStore));
            using (FileStream stream = File.Create(path))
            {
                serializer.Serialize(stream, store);
            }
        }

        private static string GetStorePath(SqlConnection conn)
        {
            string dir = Path.Combine(Path.GetDirectoryName(typeof(FormPanel).Assembly.Location), "ExcelLinks");
            string name = SafeHash(GetProjectId(conn)) + ".xml";
            return Path.Combine(dir, name);
        }

        private static string GetProjectId(SqlConnection conn)
        {
            if (conn == null)
            {
                return "unknown";
            }

            return conn.DataSource + "|" + conn.Database;
        }

        private static string SafeHash(string text)
        {
            using (MD5 md5 = MD5.Create())
            {
                byte[] bytes = md5.ComputeHash(Encoding.UTF8.GetBytes(text ?? ""));
                StringBuilder builder = new StringBuilder();
                foreach (byte b in bytes)
                {
                    builder.Append(b.ToString("x2", CultureInfo.InvariantCulture));
                }

                return builder.ToString();
            }
        }

        private static void AppendExcelSyncLog(SqlConnection conn, ExcelQuotaLink link, string oldValue, string newValue, bool manual)
        {
            try
            {
                string dir = Path.Combine(Path.GetDirectoryName(GetStorePath(conn)), "Logs");
                Directory.CreateDirectory(dir);
                string path = Path.Combine(dir, SafeHash(GetProjectId(conn)) + ".log");
                string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
                    + "\t" + (manual ? "Manual" : "Auto")
                    + "\t" + link.QuotaSequence.ToString(CultureInfo.InvariantCulture)
                    + "\t" + link.QuotaCode
                    + "\t" + link.ExcelPath
                    + "\t" + link.WorksheetName + "!" + link.CellAddress
                    + "\t" + oldValue
                    + "\t" + newValue
                    + Environment.NewLine;
                File.AppendAllText(path, line, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Log("AppendExcelSyncLog failed: " + ex.Message);
            }
        }

        private sealed class ExcelCellAddress
        {
            public string WorkbookPath;
            public string WorksheetName;
            public string CellAddress;
            public string DisplayValue;
            public List<string> SelectionAddresses;
        }

        private sealed class ExcelBindOptions
        {
            public bool ExpressionMode;
            public int RowCount;
            public string RowSuffix;
            public string Expression;
        }

        public sealed class ExcelQuotaLink
        {
            public string ProjectId { get; set; }
            public long QuotaSequence { get; set; }
            public string TotalNo { get; set; }
            public string ChapterSeq { get; set; }
            public string OrderNo { get; set; }
            public string QuotaCode { get; set; }
            public string QuotaName { get; set; }
            public string ExcelPath { get; set; }
            public string WorksheetName { get; set; }
            public string CellAddress { get; set; }
            public string Expression { get; set; }
            public string LastSyncValue { get; set; }
            public string LastStatus { get; set; }
            public string UpdatedAt { get; set; }
        }

        public sealed class ExcelLinkStore
        {
            public List<ExcelQuotaLink> Links { get; set; }

            public ExcelLinkStore()
            {
                Links = new List<ExcelQuotaLink>();
            }

            public void Upsert(ExcelQuotaLink link)
            {
                for (int i = Links.Count - 1; i >= 0; i--)
                {
                    if (Links[i].QuotaSequence == link.QuotaSequence)
                    {
                        Links.RemoveAt(i);
                    }
                }

                Links.Add(link);
            }
        }

        private sealed class PendingSync
        {
            public ExcelQuotaLink Link;
            public string ValueText;
            public decimal Quantity;
        }

        private sealed class AiQuotaMatchRow
        {
            public ExcelQuotaLink Link;
            public string QuotaUnit;
            public string CurrentQuantityText;
        }

        private sealed class AiExcelSelectionContext
        {
            public string WorkbookPath;
            public string WorksheetName;
            public readonly List<AiExcelCell> Cells = new List<AiExcelCell>();

            public HashSet<string> AddressSet
            {
                get
                {
                    return new HashSet<string>(Cells.Select(c => c.Address), StringComparer.OrdinalIgnoreCase);
                }
            }
        }

        private sealed class AiExcelCell
        {
            public string Address;
            public string Text;
            public int Row;
            public int Column;
            public bool IsNumber;
        }

        private sealed class AiMatchResult
        {
            public long QuotaSequence;
            public string CellAddress;
            public string Expression;
            public string QuantityName;
            public int Confidence;
            public string Reason;
        }

        private sealed class AiMatchPreviewItem
        {
            public bool Checked;
            public ExcelQuotaLink Link;
            public string QuotaUnit;
            public string WorkbookPath;
            public string WorksheetName;
            public string Expression;
            public string CellAddress;
            public string DisplayValue;
            public string QuantityName;
        }

        private sealed class DeepSeekExcelMatchSettings
        {
            public bool Enabled;
            public string ApiKey;
            public string Model = "deepseek-v4-pro";
            public string BaseUrl = "https://api.deepseek.com";
            public int TimeoutSeconds = 30;
            public int MaxRowsPerBatch = 8;

            public bool IsAvailable
            {
                get { return Enabled && !String.IsNullOrWhiteSpace(ApiKey); }
            }
        }

        private sealed class SyncSummary
        {
            public int Changed;
            public int Skipped;
            public string Message;
        }

        private sealed class ExcelLinkRuntime : IDisposable
        {
            private readonly Form mainForm;
            private readonly Timer timer;
            private readonly Dictionary<string, DateTime> knownWriteTimes = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
            private bool syncing;

            public ExcelLinkRuntime(Form mainForm)
            {
                this.mainForm = mainForm;
                timer = new Timer();
                timer.Interval = 1800;
                timer.Tick += delegate { Tick(); };
                Reload();
                timer.Start();
            }

            public void Reload()
            {
                SqlConnection conn = GetProjectConnection(mainForm);
                if (conn == null)
                {
                    return;
                }

                ExcelLinkStore store = LoadStore(conn);
                foreach (ExcelQuotaLink link in store.Links)
                {
                    if (!String.IsNullOrEmpty(link.ExcelPath) && File.Exists(link.ExcelPath) && !knownWriteTimes.ContainsKey(link.ExcelPath))
                    {
                        knownWriteTimes[link.ExcelPath] = File.GetLastWriteTimeUtc(link.ExcelPath);
                    }
                }
            }

            private void Tick()
            {
                if (syncing)
                {
                    return;
                }

                try
                {
                    SqlConnection conn = GetProjectConnection(mainForm);
                    if (conn == null)
                    {
                        return;
                    }

                    ExcelLinkStore store = LoadStore(conn);
                    bool changed = false;
                    foreach (ExcelQuotaLink link in store.Links)
                    {
                        if (String.IsNullOrEmpty(link.ExcelPath) || !File.Exists(link.ExcelPath))
                        {
                            continue;
                        }

                        DateTime writeTime = File.GetLastWriteTimeUtc(link.ExcelPath);
                        DateTime known;
                        if (!knownWriteTimes.TryGetValue(link.ExcelPath, out known))
                        {
                            knownWriteTimes[link.ExcelPath] = writeTime;
                            continue;
                        }

                        if (writeTime > known)
                        {
                            knownWriteTimes[link.ExcelPath] = writeTime;
                            changed = true;
                        }
                    }

                    if (changed)
                    {
                        syncing = true;
                        SyncExcelLinks(mainForm, false);
                    }
                }
                catch (Exception ex)
                {
                    Log("Excel link auto sync failed: " + ex);
                }
                finally
                {
                    syncing = false;
                }
            }

            public void Dispose()
            {
                timer.Stop();
                timer.Dispose();
            }
        }

        private sealed class ExcelSmartBindPanel : Form
        {
            private readonly Form mainForm;
            private readonly Label currentQuotaLabel;
            private readonly Label currentExcelLabel;
            private readonly RadioButton simpleMode;
            private readonly RadioButton expressionMode;
            private readonly TextBox expressionText;
            private readonly Label status;
            private readonly Timer refreshTimer;
            private ExcelCellAddress lastExcelCell;
            private string lastQuotaKey;
            private string lastExcelKey;
            private string lastExpressionExcelKey;
            private string lastAutoExpressionText;

            public ExcelSmartBindPanel(Form mainForm)
            {
                this.mainForm = mainForm;
                Text = "绑定Excel工程量";
                StartPosition = FormStartPosition.CenterParent;
                Size = new System.Drawing.Size(620, 320);
                MinimumSize = new System.Drawing.Size(560, 280);
                MinimizeBox = false;

                Label tip = new Label();
                tip.Left = 12;
                tip.Top = 10;
                tip.Width = 580;
                tip.Height = 34;
                tip.Text = "本窗口可保持打开：在软件里点定额、在WPS/Excel里点单元格，再回到这里绑定或添加到表达式。";
                tip.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

                currentQuotaLabel = new Label();
                currentQuotaLabel.Left = 12;
                currentQuotaLabel.Top = 52;
                currentQuotaLabel.Width = 580;
                currentQuotaLabel.Height = 24;
                currentQuotaLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

                currentExcelLabel = new Label();
                currentExcelLabel.Left = 12;
                currentExcelLabel.Top = 82;
                currentExcelLabel.Width = 580;
                currentExcelLabel.Height = 24;
                currentExcelLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

                simpleMode = new RadioButton();
                simpleMode.Left = 16;
                simpleMode.Top = 118;
                simpleMode.Width = 160;
                simpleMode.Text = "简单绑定当前格";
                simpleMode.Checked = true;

                expressionMode = new RadioButton();
                expressionMode.Left = 16;
                expressionMode.Top = 152;
                expressionMode.Width = 130;
                expressionMode.Text = "表达式绑定";

                expressionText = new TextBox();
                expressionText.Left = 150;
                expressionText.Top = 150;
                expressionText.Width = 330;
                expressionText.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                expressionText.GotFocus += delegate { expressionMode.Checked = true; };
                expressionMode.CheckedChanged += delegate
                {
                    if (expressionMode.Checked && lastExcelCell != null)
                    {
                        ApplyExcelCellToExpression(lastExcelCell, true);
                    }
                };

                Button multiply = new Button();
                multiply.Text = "*";
                multiply.Left = 150;
                multiply.Top = 182;
                multiply.Width = 28;
                multiply.Click += delegate { AppendExpressionToken("*"); };

                Button divide = new Button();
                divide.Text = "/";
                divide.Left = 184;
                divide.Top = 182;
                divide.Width = 28;
                divide.Click += delegate { AppendExpressionToken("/"); };

                Button plus = new Button();
                plus.Text = "+";
                plus.Left = 218;
                plus.Top = 182;
                plus.Width = 28;
                plus.Click += delegate { AppendExpressionToken("+"); };

                Button minus = new Button();
                minus.Text = "-";
                minus.Left = 252;
                minus.Top = 182;
                minus.Width = 28;
                minus.Click += delegate { AppendExpressionToken("-"); };

                Button aiMatch = new Button();
                aiMatch.Text = "AI智能匹配";
                aiMatch.Left = 12;
                aiMatch.Top = 226;
                aiMatch.Width = 120;
                aiMatch.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
                aiMatch.Click += delegate { MatchSelectedQuotasWithAi(); };

                Button bind = new Button();
                bind.Text = "绑定到当前定额";
                bind.Left = 350;
                bind.Top = 226;
                bind.Width = 115;
                bind.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
                bind.Click += delegate { BindToCurrentQuota(); };

                Button close = new Button();
                close.Text = "关闭";
                close.Left = 475;
                close.Top = 226;
                close.Width = 80;
                close.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
                close.Click += delegate { Hide(); };

                status = new Label();
                status.Left = 12;
                status.Top = 260;
                status.Width = 580;
                status.Height = 24;
                status.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;

                Controls.Add(tip);
                Controls.Add(currentQuotaLabel);
                Controls.Add(currentExcelLabel);
                Controls.Add(simpleMode);
                Controls.Add(expressionMode);
                Controls.Add(expressionText);
                Controls.Add(multiply);
                Controls.Add(divide);
                Controls.Add(plus);
                Controls.Add(minus);
                Controls.Add(aiMatch);
                Controls.Add(bind);
                Controls.Add(close);
                Controls.Add(status);

                refreshTimer = new Timer();
                refreshTimer.Interval = 700;
                refreshTimer.Tick += delegate
                {
                    if (Visible)
                    {
                        RefreshCurrentContext();
                    }
                };
                refreshTimer.Start();
                RefreshCurrentContext();
            }

            protected override void OnFormClosing(FormClosingEventArgs e)
            {
                if (e.CloseReason == CloseReason.UserClosing)
                {
                    e.Cancel = true;
                    Hide();
                    return;
                }

                base.OnFormClosing(e);
            }

            protected override void OnVisibleChanged(EventArgs e)
            {
                base.OnVisibleChanged(e);
                if (refreshTimer == null)
                {
                    return;
                }

                if (Visible)
                {
                    refreshTimer.Start();
                    RefreshCurrentContext();
                }
                else
                {
                    refreshTimer.Stop();
                }
            }

            public void RefreshCurrentContext()
            {
                DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
                DataGridViewRow row = GetCurrentQuotaRow(grid);
                string quotaText = row == null
                    ? "当前定额：未选中"
                    : "当前定额：" + GetRowValue(row, "定额编号", "定额编号DE", "编号") + " " + GetRowValue(row, "工程或费用项目名称", "名称", "项目名称");
                string quotaKey = row == null ? "" : row.Index.ToString(CultureInfo.InvariantCulture) + "|" + quotaText;
                if (!String.Equals(lastQuotaKey, quotaKey, StringComparison.Ordinal))
                {
                    currentQuotaLabel.Text = quotaText;
                    lastQuotaKey = quotaKey;
                }

                ExcelCellAddress cell;
                string error;
                if (TryGetActiveExcelCell(out cell, out error))
                {
                    string excelKey = cell.WorkbookPath + "|" + cell.WorksheetName + "|" + BuildSelectionDisplay(cell);
                    if (!String.Equals(lastExcelKey, excelKey, StringComparison.OrdinalIgnoreCase))
                    {
                        currentExcelLabel.Text = "当前Excel：" + Path.GetFileName(cell.WorkbookPath) + "!" + cell.WorksheetName + "!" + BuildSelectionDisplay(cell);
                        lastExcelKey = excelKey;
                        lastExcelCell = cell;
                        ApplyExcelCellToExpression(cell, false);
                    }
                }
                else
                {
                    if (!String.Equals(lastExcelKey, "", StringComparison.Ordinal))
                    {
                        currentExcelLabel.Text = "当前Excel：未连接或未选中单元格";
                        lastExcelKey = "";
                    }
                }
            }

            private void AppendExpressionToken(string token)
            {
                expressionMode.Checked = true;
                expressionText.Text = expressionText.Text + token;
                expressionText.SelectionStart = expressionText.Text.Length;
                expressionText.Focus();
            }

            private void ApplyExcelCellToExpression(ExcelCellAddress cell, bool force)
            {
                if (!expressionMode.Checked || cell == null)
                {
                    return;
                }

                string excelKey = cell.WorkbookPath + "|" + cell.WorksheetName + "|" + BuildSelectionDisplay(cell);
                if (!force && String.Equals(lastExpressionExcelKey, excelKey, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                string token = BuildDefaultExpression(cell);
                string current = expressionText.Text.Trim().ToUpperInvariant();
                if (String.IsNullOrEmpty(current))
                {
                    expressionText.Text = token;
                    lastAutoExpressionText = token;
                }
                else if (EndsWithExpressionOperator(current))
                {
                    expressionText.Text = expressionText.Text + token;
                    lastAutoExpressionText = expressionText.Text.Trim().ToUpperInvariant();
                }
                else if (String.Equals(current, lastAutoExpressionText, StringComparison.OrdinalIgnoreCase))
                {
                    expressionText.Text = token;
                    lastAutoExpressionText = token;
                }
                else
                {
                    lastExpressionExcelKey = excelKey;
                    return;
                }

                expressionText.SelectionStart = expressionText.Text.Length;
                lastExpressionExcelKey = excelKey;
                status.Text = "已添加：" + token;
            }

            private static bool EndsWithExpressionOperator(string text)
            {
                if (String.IsNullOrWhiteSpace(text))
                {
                    return false;
                }

                char ch = text.Trim()[text.Trim().Length - 1];
                return ch == '+' || ch == '-' || ch == '*' || ch == '/';
            }

            private void MatchSelectedQuotasWithAi()
            {
                try
                {
                    SqlConnection conn = GetProjectConnection(mainForm);
                    if (conn == null)
                    {
                        status.Text = "没有找到当前项目数据库连接。";
                        return;
                    }

                    DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
                    List<AiQuotaMatchRow> quotas = BuildAiQuotaRows(mainForm, conn, GetSelectedQuotaRows(grid));
                    if (quotas.Count == 0)
                    {
                        status.Text = "请先在定额输入表中选择要匹配的定额。";
                        return;
                    }

                    AiExcelSelectionContext selection;
                    string error;
                    if (!TryReadActiveExcelSelectionContext(out selection, out error))
                    {
                        status.Text = "AI匹配失败：" + error;
                        return;
                    }

                    DeepSeekExcelMatchSettings settings = LoadDeepSeekExcelMatchSettings();
                    if (!settings.IsAvailable)
                    {
                        status.Text = "DeepSeek未启用或未配置API Key，请先在推荐定额窗口的AI设置中配置。";
                        return;
                    }

                    status.Text = "正在先按工程数量快速匹配，剩余项再交给DeepSeek...";
                    Application.DoEvents();

                    List<AiMatchPreviewItem> preview = new List<AiMatchPreviewItem>();
                    string fallbackMessage = null;
                    try
                    {
                        List<AiMatchResult> results = RequestAiExcelMatches(settings, quotas, selection);
                        preview = BuildAiMatchPreviewItems(quotas, selection, results);
                    }
                    catch (Exception aiEx)
                    {
                        Log("DeepSeek Excel match request failed, fallback to local quantity match: " + aiEx.Message);
                        fallbackMessage = "DeepSeek超时或失败，已改用本地数量匹配。";
                    }

                    if (preview.Count == 0)
                    {
                        List<AiMatchPreviewItem> fallback = BuildLocalQuantityMatchPreviewItems(quotas, selection);
                        if (fallback.Count > 0)
                        {
                            preview = fallback;
                            if (String.IsNullOrEmpty(fallbackMessage))
                            {
                                fallbackMessage = "AI无可用结果，已改用本地数量匹配。";
                            }
                        }
                    }

                    if (preview.Count == 0)
                    {
                        AddUnmatchedAiPreviewItems(preview, quotas, selection);
                    }

                    if (preview.Count == 0)
                    {
                        status.Text = "AI没有返回可确认的绑定建议。";
                        return;
                    }

                    if (!String.IsNullOrEmpty(fallbackMessage))
                    {
                        status.Text = fallbackMessage;
                    }

                    AiMatchPreviewDialog modelessDialog = new AiMatchPreviewDialog(preview);
                    modelessDialog.Accepted += delegate(List<AiMatchPreviewItem> accepted)
                    {
                        SaveAiMatchPreviewAccepted(conn, selection, accepted);
                    };
                    modelessDialog.Cancelled += delegate
                    {
                        status.Text = "已取消AI匹配绑定。";
                    };
                    modelessDialog.Show(this);
                    return;
                }
                catch (Exception ex)
                {
                    Log("AI Excel match failed: " + ex);
                    status.Text = "AI匹配失败：" + ex.Message;
                }
            }

            private void SaveAiMatchPreviewAccepted(SqlConnection conn, AiExcelSelectionContext selection, List<AiMatchPreviewItem> accepted)
            {
                try
                {
                    if (accepted == null || accepted.Count == 0)
                    {
                        status.Text = "没有勾选任何AI匹配结果。";
                        return;
                    }

                    ExcelLinkStore store = LoadStore(conn);
                    foreach (AiMatchPreviewItem item in accepted)
                    {
                        item.Link.ExcelPath = String.IsNullOrWhiteSpace(item.WorkbookPath) ? selection.WorkbookPath : item.WorkbookPath;
                        item.Link.WorksheetName = String.IsNullOrWhiteSpace(item.WorksheetName) ? selection.WorksheetName : item.WorksheetName;
                        item.Link.CellAddress = item.CellAddress;
                        item.Link.Expression = item.Expression;
                        item.Link.LastSyncValue = item.DisplayValue ?? "";
                        item.Link.LastStatus = "AI智能匹配绑定，等待同步";
                        item.Link.UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                        store.Upsert(item.Link);
                    }

                    SaveStore(conn, store);
                    EnsureExcelLinkRuntime(mainForm);
                    if (ExcelLinkRuntimes.ContainsKey(mainForm))
                    {
                        ExcelLinkRuntimes[mainForm].Reload();
                    }

                    RefreshExcelLinkPanel(mainForm);
                    status.Text = "AI已绑定 " + accepted.Count.ToString(CultureInfo.InvariantCulture) + " 条选中定额。";
                }
                catch (Exception ex)
                {
                    Log("Save AI Excel match preview failed: " + ex);
                    status.Text = "AI匹配保存失败：" + ex.Message;
                }
            }

            private void BindToCurrentQuota()
            {
                try
                {
                    SqlConnection conn = GetProjectConnection(mainForm);
                    if (conn == null)
                    {
                        status.Text = "没有找到当前项目数据库连接。";
                        return;
                    }

                    DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
                    DataGridViewRow row = GetCurrentQuotaRow(grid);
                    if (row == null)
                    {
                        status.Text = "请先在软件里点击要绑定的定额行。";
                        return;
                    }

                    ExcelCellAddress cell;
                    string error;
                    if (!TryGetActiveExcelCell(out cell, out error))
                    {
                        status.Text = "请先在WPS/Excel里点击工程数量单元格。";
                        return;
                    }

                    string expression = simpleMode.Checked ? cell.CellAddress : expressionText.Text.Trim().ToUpperInvariant();
                    if (String.IsNullOrEmpty(expression))
                    {
                        status.Text = "请填写表达式，或在表达式模式下点击WPS/Excel单元格自动加入。";
                        expressionText.Focus();
                        return;
                    }

                    string firstCell = ExtractFirstCellAddress(expression);
                    if (String.IsNullOrEmpty(firstCell))
                    {
                        status.Text = "表达式里至少需要一个单元格地址。";
                        return;
                    }

                    ExcelQuotaLink link;
                    if (!TryCreateQuotaLink(mainForm, conn, row, out link, out error))
                    {
                        status.Text = error;
                        return;
                    }

                    string displayValue = null;
                    string readError = null;
                    decimal quantity = 0;
                    bool evaluated = false;
                    if (simpleMode.Checked &&
                        String.Equals(NormalizeCellAddress(expression), NormalizeCellAddress(cell.CellAddress), StringComparison.OrdinalIgnoreCase) &&
                        !String.IsNullOrWhiteSpace(cell.DisplayValue))
                    {
                        evaluated = TryEvaluateDecimal(cell.DisplayValue, out quantity, out readError);
                        if (evaluated)
                        {
                            displayValue = cell.DisplayValue;
                        }
                    }

                    if (!evaluated && !TryEvaluateWorkbookExpression(cell.WorkbookPath, cell.WorksheetName, expression, out displayValue, out quantity, out readError))
                    {
                        status.Text = "表达式无法读取或计算：" + readError;
                        return;
                    }

                    link.ExcelPath = cell.WorkbookPath;
                    link.WorksheetName = cell.WorksheetName;
                    link.CellAddress = firstCell;
                    link.Expression = expression;
                    link.LastSyncValue = displayValue ?? "";
                    link.LastStatus = simpleMode.Checked ? "简单绑定，等待同步" : "表达式绑定，等待同步";
                    link.UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

                    ExcelLinkStore store = LoadStore(conn);
                    store.Upsert(link);
                    SaveStore(conn, store);

                    EnsureExcelLinkRuntime(mainForm);
                    if (ExcelLinkRuntimes.ContainsKey(mainForm))
                    {
                        ExcelLinkRuntimes[mainForm].Reload();
                    }

                    RefreshExcelLinkPanel(mainForm);
                    RefreshCurrentContext();
                    status.Text = "已绑定：" + link.QuotaCode + " -> " + expression;
                }
                catch (Exception ex)
                {
                    Log("ExcelSmartBindPanel bind failed: " + ex);
                    status.Text = "绑定失败：" + ex.Message;
                }
            }
        }

        private sealed class AiMatchPreviewDialog : Form
        {
            private readonly DataGridView grid;
            private readonly List<AiMatchPreviewItem> items;
            private readonly Label status;
            public event Action<List<AiMatchPreviewItem>> Accepted;
            public event Action Cancelled;

            public AiMatchPreviewDialog(List<AiMatchPreviewItem> previewItems)
            {
                items = previewItems ?? new List<AiMatchPreviewItem>();
                Text = "AI智能匹配确认";
                StartPosition = FormStartPosition.CenterParent;
                Size = new System.Drawing.Size(860, 460);
                MinimumSize = new System.Drawing.Size(760, 360);
                MinimizeBox = false;

                Label tip = new Label();
                tip.Dock = DockStyle.Top;
                tip.Height = 34;
                tip.Padding = new Padding(8, 8, 8, 0);
                tip.Text = "请确认AI匹配结果，勾选后写入Excel联动绑定。";

                grid = new DataGridView();
                grid.Dock = DockStyle.Fill;
                grid.AllowUserToAddRows = false;
                grid.AllowUserToDeleteRows = false;
                grid.RowHeadersVisible = false;
                grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                grid.CellDoubleClick += delegate { ApplyCurrentExcelCellToSelectedRow(); };

                DataGridViewCheckBoxColumn checkedColumn = new DataGridViewCheckBoxColumn();
                checkedColumn.Name = "Checked";
                checkedColumn.HeaderText = "绑定";
                checkedColumn.FillWeight = 45;
                grid.Columns.Add(checkedColumn);
                grid.Columns.Add("QuotaCode", "定额编号");
                grid.Columns.Add("QuotaName", "定额名称");
                grid.Columns.Add("QuotaUnit", "定额单位");
                grid.Columns.Add("Expression", "Excel单元格/表达式");
                grid.Columns.Add("Value", "拟链接工程数量");
                grid.Columns["QuotaCode"].FillWeight = 80;
                grid.Columns["QuotaName"].FillWeight = 230;
                grid.Columns["QuotaUnit"].FillWeight = 65;
                grid.Columns["Expression"].FillWeight = 130;
                grid.Columns["Value"].FillWeight = 95;

                foreach (AiMatchPreviewItem item in items)
                {
                    int index = grid.Rows.Add(
                        item.Checked,
                        item.Link == null ? "" : item.Link.QuotaCode,
                        item.Link == null ? "" : item.Link.QuotaName,
                        item.QuotaUnit,
                        item.Expression,
                        item.DisplayValue);
                    grid.Rows[index].Tag = item;
                }

                Button applyCurrent = new Button();
                applyCurrent.Text = "用当前Excel格匹配";
                applyCurrent.Width = 130;
                applyCurrent.Click += delegate { ApplyCurrentExcelCellToSelectedRow(); };

                Button ok = new Button();
                ok.Text = "确认绑定";
                ok.Width = 90;
                ok.Click += delegate
                {
                    grid.EndEdit();
                    List<AiMatchPreviewItem> accepted = GetAcceptedItems();
                    if (accepted.Count == 0)
                    {
                        status.Text = "请至少勾选一条已匹配Excel单元格的定额。";
                        return;
                    }

                    if (Accepted != null)
                    {
                        Accepted(accepted);
                    }

                    DialogResult = DialogResult.OK;
                    Close();
                };

                Button cancel = new Button();
                cancel.Text = "取消";
                cancel.Width = 75;
                cancel.Click += delegate
                {
                    if (Cancelled != null)
                    {
                        Cancelled();
                    }

                    DialogResult = DialogResult.Cancel;
                    Close();
                };

                FlowLayoutPanel buttons = new FlowLayoutPanel();
                buttons.Dock = DockStyle.Bottom;
                buttons.Height = 44;
                buttons.FlowDirection = FlowDirection.RightToLeft;
                buttons.Padding = new Padding(8);
                buttons.Controls.Add(cancel);
                buttons.Controls.Add(ok);
                buttons.Controls.Add(applyCurrent);

                status = new Label();
                status.Dock = DockStyle.Bottom;
                status.Height = 26;
                status.Padding = new Padding(8, 2, 8, 2);
                status.Text = "未匹配行可选中后，到WPS/Excel点单元格，再点“用当前Excel格匹配”。";

                Controls.Add(grid);
                Controls.Add(buttons);
                Controls.Add(status);
                Controls.Add(tip);
            }

            private void ApplyCurrentExcelCellToSelectedRow()
            {
                grid.EndEdit();
                DataGridViewRow row = grid.CurrentRow;
                if (row == null)
                {
                    status.Text = "请先选中预览表中的一条定额。";
                    return;
                }

                AiMatchPreviewItem item = row.Tag as AiMatchPreviewItem;
                if (item == null || item.Link == null)
                {
                    status.Text = "当前行无法匹配。";
                    return;
                }

                ExcelCellAddress cell;
                string error;
                if (!TryGetActiveExcelCell(out cell, out error))
                {
                    status.Text = "请先在WPS/Excel里点选工程数量单元格。";
                    return;
                }

                string expression = BuildDefaultExpression(cell);
                string displayValue;
                decimal quantity;
                string readError;
                if (!TryEvaluateWorkbookExpression(cell.WorkbookPath, cell.WorksheetName, expression, out displayValue, out quantity, out readError))
                {
                    status.Text = "当前Excel格无法计算：" + readError;
                    return;
                }

                item.Checked = true;
                item.WorkbookPath = cell.WorkbookPath;
                item.WorksheetName = cell.WorksheetName;
                item.Expression = expression;
                item.CellAddress = ExtractFirstCellAddress(expression);
                item.DisplayValue = displayValue ?? "";
                item.QuantityName = "";

                row.Cells["Checked"].Value = true;
                row.Cells["Expression"].Value = item.Expression;
                row.Cells["Value"].Value = item.DisplayValue;
                status.Text = "已匹配：" + (item.Link.QuotaCode ?? "") + " -> " + Path.GetFileName(item.WorkbookPath) + "!" + item.WorksheetName + "!" + item.Expression;
            }

            public List<AiMatchPreviewItem> GetAcceptedItems()
            {
                List<AiMatchPreviewItem> accepted = new List<AiMatchPreviewItem>();
                foreach (DataGridViewRow row in grid.Rows)
                {
                    bool isChecked = row.Cells["Checked"].Value is bool && (bool)row.Cells["Checked"].Value;
                    AiMatchPreviewItem item = row.Tag as AiMatchPreviewItem;
                    if (isChecked && item != null && !String.IsNullOrWhiteSpace(item.Expression) && !String.IsNullOrWhiteSpace(item.CellAddress))
                    {
                        accepted.Add(item);
                    }
                }

                return accepted;
            }
        }

        private sealed class QuickBindPanel : Form
        {
            private const int PreviewRows = 40;
            private const int PreviewColumns = 12;
            private readonly Form mainForm;
            private readonly TextBox fileText;
            private readonly ComboBox sheetBox;
            private readonly DataGridView preview;
            private readonly Label status;
            private readonly CheckBox expressionMode;
            private readonly TextBox expressionText;
            private readonly TextBox previewStartText;
            private readonly TextBox cellAddressText;

            public QuickBindPanel(Form mainForm)
            {
                this.mainForm = mainForm;
                Text = "Excel快速绑定";
                StartPosition = FormStartPosition.CenterParent;
                Size = new System.Drawing.Size(1100, 700);
                MinimumSize = new System.Drawing.Size(820, 560);
                MinimizeBox = false;

                Label tip = new Label();
                tip.Left = 12;
                tip.Top = 10;
                tip.Width = 1060;
                tip.Height = 24;
                tip.Text = "保持本窗口打开：选中定额行后，可点预览格、输入单元格地址，或直接绑定 Excel 当前选中单元格。";
                tip.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

                Label fileLabel = new Label();
                fileLabel.Text = "Excel文件";
                fileLabel.Left = 12;
                fileLabel.Top = 45;
                fileLabel.Width = 70;

                fileText = new TextBox();
                fileText.Left = 86;
                fileText.Top = 42;
                fileText.Width = 810;
                fileText.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

                Button choose = new Button();
                choose.Text = "选择";
                choose.Left = 906;
                choose.Top = 40;
                choose.Width = 75;
                choose.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                choose.Click += delegate
                {
                    using (OpenFileDialog chooser = new OpenFileDialog())
                    {
                        chooser.Title = "选择工程数量Excel文件";
                        chooser.Filter = "Excel文件 (*.xls;*.xlsx)|*.xls;*.xlsx|所有文件 (*.*)|*.*";
                        if (chooser.ShowDialog(this) == DialogResult.OK)
                        {
                            fileText.Text = chooser.FileName;
                            LoadSheetNamesIntoCombo(chooser.FileName, sheetBox);
                            status.Text = "已选择Excel文件。大文件建议直接在Excel里选中单元格后点“绑定Excel当前格”；需要预览时再点“预览”。";
                        }
                    }
                };

                Button refresh = new Button();
                refresh.Text = "刷新";
                refresh.Left = 988;
                refresh.Top = 40;
                refresh.Width = 75;
                refresh.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                refresh.Click += delegate { LoadPreview(); };

                Label sheetLabel = new Label();
                sheetLabel.Text = "工作表";
                sheetLabel.Left = 12;
                sheetLabel.Top = 82;
                sheetLabel.Width = 70;

                sheetBox = new ComboBox();
                sheetBox.Left = 86;
                sheetBox.Top = 79;
                sheetBox.Width = 300;
                sheetBox.DropDownStyle = ComboBoxStyle.DropDownList;
                sheetBox.SelectedIndexChanged += delegate { status.Text = "已切换工作表。需要查看内嵌表格时点“预览”。"; };

                expressionMode = new CheckBox();
                expressionMode.Text = "计算式模式";
                expressionMode.Left = 410;
                expressionMode.Top = 81;
                expressionMode.Width = 95;
                expressionMode.Anchor = AnchorStyles.Top | AnchorStyles.Left;

                expressionText = new TextBox();
                expressionText.Left = 510;
                expressionText.Top = 79;
                expressionText.Width = 360;
                expressionText.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

                Button plus = CreateExpressionButton("+", 880, 77);
                Button minus = CreateExpressionButton("-", 908, 77);
                Button multiply = CreateExpressionButton("*", 936, 77);
                Button divide = CreateExpressionButton("/", 964, 77);

                Label startLabel = new Label();
                startLabel.Text = "预览起点";
                startLabel.Left = 12;
                startLabel.Top = 116;
                startLabel.Width = 70;

                previewStartText = new TextBox();
                previewStartText.Left = 86;
                previewStartText.Top = 113;
                previewStartText.Width = 70;
                previewStartText.Text = "A1";

                Button jumpPreview = new Button();
                jumpPreview.Text = "预览";
                jumpPreview.Left = 164;
                jumpPreview.Top = 111;
                jumpPreview.Width = 62;
                jumpPreview.Click += delegate { LoadPreview(); };

                Label cellLabel = new Label();
                cellLabel.Text = "单元格";
                cellLabel.Left = 244;
                cellLabel.Top = 116;
                cellLabel.Width = 54;

                cellAddressText = new TextBox();
                cellAddressText.Left = 300;
                cellAddressText.Top = 113;
                cellAddressText.Width = 80;

                Button bindCell = new Button();
                bindCell.Text = "绑定单元格";
                bindCell.Left = 388;
                bindCell.Top = 111;
                bindCell.Width = 95;
                bindCell.Click += delegate { BindTypedCell(); };

                Button bindActive = new Button();
                bindActive.Text = "绑定Excel当前格";
                bindActive.Left = 494;
                bindActive.Top = 111;
                bindActive.Width = 125;
                bindActive.Click += delegate { BindActiveExcelCell(); };

                Button bindExpression = new Button();
                bindExpression.Text = "绑定计算式";
                bindExpression.Left = 730;
                bindExpression.Top = 626;
                bindExpression.Width = 110;
                bindExpression.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
                bindExpression.Click += delegate { BindExpression(); };

                Button clearExpression = new Button();
                clearExpression.Text = "清空公式";
                clearExpression.Left = 850;
                clearExpression.Top = 626;
                clearExpression.Width = 90;
                clearExpression.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
                clearExpression.Click += delegate { expressionText.Text = ""; };

                preview = new DataGridView();
                preview.Left = 12;
                preview.Top = 150;
                preview.Width = 1060;
                preview.Height = 460;
                preview.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
                preview.ReadOnly = true;
                preview.AllowUserToAddRows = false;
                preview.AllowUserToDeleteRows = false;
                preview.SelectionMode = DataGridViewSelectionMode.CellSelect;
                preview.RowHeadersWidth = 54;
                preview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                preview.CellClick += delegate(object sender, DataGridViewCellEventArgs e)
                {
                    if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                    {
                        string address = GetPreviewCellAddress(e.RowIndex, e.ColumnIndex);
                        cellAddressText.Text = address;
                        if (expressionMode.Checked || !String.IsNullOrWhiteSpace(expressionText.Text))
                        {
                            AppendExpressionToken(address);
                        }
                        else
                        {
                            BindCurrentQuota(address, address);
                        }
                    }
                };

                status = new Label();
                status.Left = 12;
                status.Top = 626;
                status.Width = 700;
                status.Height = 24;
                status.Text = "请选择Excel文件。";
                status.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;

                Button close = new Button();
                close.Text = "关闭";
                close.Left = 992;
                close.Top = 624;
                close.Width = 80;
                close.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
                close.Click += delegate { Hide(); };

                Controls.Add(tip);
                Controls.Add(fileLabel);
                Controls.Add(fileText);
                Controls.Add(choose);
                Controls.Add(refresh);
                Controls.Add(sheetLabel);
                Controls.Add(sheetBox);
                Controls.Add(expressionMode);
                Controls.Add(expressionText);
                Controls.Add(plus);
                Controls.Add(minus);
                Controls.Add(multiply);
                Controls.Add(divide);
                Controls.Add(startLabel);
                Controls.Add(previewStartText);
                Controls.Add(jumpPreview);
                Controls.Add(cellLabel);
                Controls.Add(cellAddressText);
                Controls.Add(bindCell);
                Controls.Add(bindActive);
                Controls.Add(preview);
                Controls.Add(status);
                Controls.Add(bindExpression);
                Controls.Add(clearExpression);
                Controls.Add(close);
            }

            private Button CreateExpressionButton(string text, int left, int top)
            {
                Button button = new Button();
                button.Text = text;
                button.Left = left;
                button.Top = top;
                button.Width = 24;
                button.Height = 24;
                button.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                button.Click += delegate
                {
                    expressionMode.Checked = true;
                    if (String.IsNullOrEmpty(expressionText.Text.Trim()) && text != "-")
                    {
                        status.Text = "请先点击一个单元格，再选择运算符。";
                        return;
                    }

                    AppendExpressionToken(text);
                };
                return button;
            }

            protected override void OnFormClosing(FormClosingEventArgs e)
            {
                if (e.CloseReason == CloseReason.UserClosing)
                {
                    e.Cancel = true;
                    Hide();
                    return;
                }

                base.OnFormClosing(e);
            }

            private void LoadPreview()
            {
                string path = fileText.Text.Trim();
                string sheet = Convert.ToString(sheetBox.Text, CultureInfo.InvariantCulture);
                string start = NormalizePreviewStart();
                LoadPreviewGrid(path, sheet, preview, start, PreviewRows, PreviewColumns);
                status.Text = preview.Columns.Count == 0
                    ? "未能读取预览，请确认文件没有损坏或被独占打开。"
                    : "正在预览 " + start + " 起的区域；也可直接输入任意单元格地址绑定。";
            }

            private void AppendExpressionToken(string token)
            {
                expressionMode.Checked = true;
                expressionText.Text = expressionText.Text + token;
                expressionText.SelectionStart = expressionText.Text.Length;
                status.Text = "计算式：" + expressionText.Text + "；完成后点“绑定计算式”。";
            }

            private string NormalizePreviewStart()
            {
                string start = previewStartText.Text.Trim();
                CellRef cell;
                if (String.IsNullOrEmpty(start) || !TryParseCellAddress(start, out cell))
                {
                    start = "A1";
                    previewStartText.Text = start;
                }

                return start.ToUpperInvariant();
            }

            private string GetPreviewCellAddress(int rowIndex, int columnIndex)
            {
                CellRef start;
                if (!TryParseCellAddress(NormalizePreviewStart(), out start))
                {
                    start = new CellRef { Column = 1, Row = 1 };
                }

                return ColumnNumberToName(start.Column + columnIndex) + (start.Row + rowIndex).ToString(CultureInfo.InvariantCulture);
            }

            private void BindTypedCell()
            {
                string address = cellAddressText.Text.Trim().ToUpperInvariant();
                CellRef cell;
                if (!TryParseCellAddress(address, out cell))
                {
                    status.Text = "请输入正确的单元格地址，例如 E4、AB125。";
                    cellAddressText.Focus();
                    cellAddressText.SelectAll();
                    return;
                }

                BindCurrentQuota(address, address);
            }

            private void BindActiveExcelCell()
            {
                ExcelCellAddress cell;
                string error;
                if (!TryGetActiveExcelCell(out cell, out error))
                {
                    status.Text = error;
                    return;
                }

                fileText.Text = cell.WorkbookPath;
                LoadSheetNamesIntoCombo(cell.WorkbookPath, sheetBox);
                sheetBox.Text = cell.WorksheetName;
                previewStartText.Text = cell.CellAddress;
                cellAddressText.Text = cell.CellAddress;
                BindCurrentQuota(cell.CellAddress, cell.CellAddress);
            }

            private void BindExpression()
            {
                string expression = expressionText.Text.Trim();
                if (String.IsNullOrEmpty(expression))
                {
                    status.Text = "请先勾选计算式模式并点击单元格组成表达式，例如 E4+E5 或 E4*1.15。";
                    return;
                }

                string firstCell = ExtractFirstCellAddress(expression);
                if (String.IsNullOrEmpty(firstCell))
                {
                    status.Text = "计算式里至少需要一个单元格地址。";
                    return;
                }

                BindCurrentQuota(firstCell, expression);
            }

            private void BindCurrentQuota(string address, string expression)
            {
                try
                {
                    string path = fileText.Text.Trim();
                    string sheet = Convert.ToString(sheetBox.Text, CultureInfo.InvariantCulture);
                    if (!File.Exists(path) || String.IsNullOrEmpty(sheet))
                    {
                        status.Text = "请先选择Excel文件和工作表。";
                        return;
                    }

                    SqlConnection conn = GetProjectConnection(mainForm);
                    if (conn == null)
                    {
                        status.Text = "没有找到当前项目数据库连接。";
                        return;
                    }

                    DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
                    DataGridViewRow row = GetCurrentQuotaRow(grid);
                    if (row == null)
                    {
                        status.Text = "请先在软件定额表中选中一条定额行。";
                        return;
                    }

                    ExcelQuotaLink link;
                    string error;
                    if (!TryCreateQuotaLink(mainForm, conn, row, out link, out error))
                    {
                        status.Text = error;
                        return;
                    }

                    string displayValue;
                    string readError;
                    decimal quantity;
                    if (!TryEvaluateWorkbookExpression(path, sheet, expression, out displayValue, out quantity, out readError))
                    {
                        status.Text = "计算式无法读取或计算：" + readError;
                        return;
                    }

                    link.ExcelPath = path;
                    link.WorksheetName = sheet;
                    link.CellAddress = address;
                    link.Expression = expression;
                    link.LastSyncValue = displayValue ?? "";
                    link.LastStatus = "快速绑定，等待同步";
                    link.UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

                    ExcelLinkStore store = LoadStore(conn);
                    store.Upsert(link);
                    SaveStore(conn, store);

                    EnsureExcelLinkRuntime(mainForm);
                    if (ExcelLinkRuntimes.ContainsKey(mainForm))
                    {
                        ExcelLinkRuntimes[mainForm].Reload();
                    }

                    RefreshExcelLinkPanel(mainForm);
                    status.Text = "已绑定：" + link.QuotaCode + " -> " + Path.GetFileName(path) + "!" + sheet + "!" + expression;
                    if (!String.Equals(address, expression, StringComparison.OrdinalIgnoreCase))
                    {
                        expressionText.Text = "";
                        expressionMode.Checked = false;
                    }
                }
                catch (Exception ex)
                {
                    Log("Quick bind failed: " + ex);
                    status.Text = "绑定失败：" + ex.Message;
                }
            }
        }

        private sealed class ExcelLinkPanel : Form
        {
            private readonly Form mainForm;
            private readonly DataGridView grid;
            private readonly Label status;

            public ExcelLinkPanel(Form mainForm)
            {
                this.mainForm = mainForm;
                Text = "Excel工程量联动";
                StartPosition = FormStartPosition.CenterParent;
                Size = new System.Drawing.Size(980, 520);
                MinimizeBox = false;

                grid = new DataGridView();
                grid.Dock = DockStyle.Fill;
                grid.ReadOnly = true;
                grid.AllowUserToAddRows = false;
                grid.AllowUserToDeleteRows = false;
                grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                grid.MultiSelect = true;
                grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                grid.Columns.Add("QuotaSequence", "定额序号");
                grid.Columns.Add("QuotaCode", "定额编号");
                grid.Columns.Add("QuotaName", "名称");
                grid.Columns.Add("Excel", "Excel单元格");
                grid.Columns.Add("LastValue", "最近值");
                grid.Columns.Add("Status", "状态");
                grid.Columns.Add("UpdatedAt", "更新时间");
                grid.Columns["QuotaSequence"].FillWeight = 70;
                grid.Columns["QuotaCode"].FillWeight = 90;
                grid.Columns["QuotaName"].FillWeight = 180;
                grid.Columns["Excel"].FillWeight = 220;
                grid.Columns["LastValue"].FillWeight = 80;
                grid.Columns["Status"].FillWeight = 100;
                grid.Columns["UpdatedAt"].FillWeight = 110;

                Button sync = new Button();
                sync.Text = "同步一次";
                sync.Width = 90;
                sync.Click += delegate
                {
                    try
                    {
                        SyncSummary result = SyncExcelLinks(mainForm, true);
                        status.Text = result.Message;
                        Reload();
                    }
                    catch (Exception ex)
                    {
                        status.Text = "同步失败：" + ex.Message;
                        Log("Manual Excel sync failed: " + ex);
                    }
                };

                Button refresh = new Button();
                refresh.Text = "刷新";
                refresh.Width = 75;
                refresh.Click += delegate { Reload(); };

                Button delete = new Button();
                delete.Text = "删除选中绑定";
                delete.Width = 110;
                delete.Click += delegate { DeleteSelectedLinks(); };

                Button close = new Button();
                close.Text = "关闭";
                close.Width = 75;
                close.Click += delegate { Hide(); };

                status = new Label();
                status.AutoSize = false;
                status.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
                status.Dock = DockStyle.Fill;

                FlowLayoutPanel buttons = new FlowLayoutPanel();
                buttons.Dock = DockStyle.Top;
                buttons.Height = 42;
                buttons.Padding = new Padding(8);
                buttons.Controls.Add(sync);
                buttons.Controls.Add(refresh);
                buttons.Controls.Add(delete);
                buttons.Controls.Add(close);

                Panel bottom = new Panel();
                bottom.Dock = DockStyle.Bottom;
                bottom.Height = 28;
                bottom.Padding = new Padding(8, 0, 8, 4);
                bottom.Controls.Add(status);

                Controls.Add(grid);
                Controls.Add(bottom);
                Controls.Add(buttons);
            }

            protected override void OnFormClosing(FormClosingEventArgs e)
            {
                if (e.CloseReason == CloseReason.UserClosing)
                {
                    e.Cancel = true;
                    Hide();
                    return;
                }

                base.OnFormClosing(e);
            }

            public void Reload()
            {
                grid.Rows.Clear();
                SqlConnection conn = GetProjectConnection(mainForm);
                if (conn == null)
                {
                    status.Text = "没有找到当前项目数据库连接。";
                    return;
                }

                ExcelLinkStore store = LoadStore(conn);
                foreach (ExcelQuotaLink link in store.Links.OrderBy(l => l.QuotaSequence))
                {
                    string excel = Path.GetFileName(link.ExcelPath) + "!" + link.WorksheetName + "!" + link.CellAddress;
                    grid.Rows.Add(
                        link.QuotaSequence.ToString(CultureInfo.InvariantCulture),
                        link.QuotaCode,
                        link.QuotaName,
                        excel,
                        link.LastSyncValue,
                        link.LastStatus,
                        link.UpdatedAt);
                }

                status.Text = "绑定数量：" + store.Links.Count.ToString(CultureInfo.InvariantCulture);
            }

            private void DeleteSelectedLinks()
            {
                SqlConnection conn = GetProjectConnection(mainForm);
                if (conn == null)
                {
                    status.Text = "没有找到当前项目数据库连接。";
                    return;
                }

                HashSet<long> ids = new HashSet<long>();
                foreach (DataGridViewRow row in grid.SelectedRows)
                {
                    long id;
                    if (row.Cells["QuotaSequence"].Value != null && Int64.TryParse(Convert.ToString(row.Cells["QuotaSequence"].Value, CultureInfo.InvariantCulture), out id))
                    {
                        ids.Add(id);
                    }
                }

                if (ids.Count == 0)
                {
                    status.Text = "请先选择要删除的绑定。";
                    return;
                }

                ExcelLinkStore store = LoadStore(conn);
                int before = store.Links.Count;
                store.Links.RemoveAll(l => ids.Contains(l.QuotaSequence));
                SaveStore(conn, store);
                if (ExcelLinkRuntimes.ContainsKey(mainForm))
                {
                    ExcelLinkRuntimes[mainForm].Reload();
                }

                Reload();
                status.Text = "已删除 " + (before - store.Links.Count).ToString(CultureInfo.InvariantCulture) + " 条绑定。";
            }
        }

        private const uint ObjIdNativeOm = 0xFFFFFFF0;

        private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(
            IntPtr hwnd,
            uint dwObjectId,
            ref Guid riid,
            [MarshalAs(UnmanagedType.IDispatch)] out object ppvObject);
    }
}
