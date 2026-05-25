using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace RecoQuotaRecommend
{
    public sealed class QuotaRecommendPanel : Form
    {
        private static readonly HashSet<Form> InstalledForms = new HashSet<Form>();
        private static readonly Dictionary<ContextMenuStrip, MenuInfo> MenuInfos = new Dictionary<ContextMenuStrip, MenuInfo>();
        private static readonly Dictionary<Form, RecommendDialog> RecommendDialogs = new Dictionary<Form, RecommendDialog>();
        private static bool idleHooked;

        private sealed class MenuInfo
        {
            public Form MainForm;
            public string Name;
        }

        public static void InstallOnIdle()
        {
            if (idleHooked)
            {
                return;
            }

            idleHooked = true;
            Log("InstallOnIdle registered.");
            Application.Idle += delegate
            {
                try
                {
                    Form mainForm = FindMainForm();
                    if (mainForm != null && !InstalledForms.Contains(mainForm))
                    {
                        Install(mainForm);
                    }
                }
                catch (Exception ex)
                {
                    Log("Idle install failed: " + ex);
                }
            };
        }

        private static void Install(Form mainForm)
        {
            int menus = InstallAllContextMenus(mainForm);
            if (menus == 0)
            {
                Log("Context menus not found.");
                return;
            }

            InstalledForms.Add(mainForm);
            mainForm.FormClosed += delegate { InstalledForms.Remove(mainForm); };
            Log("Quota recommend menu installed. menus=" + menus.ToString(CultureInfo.InvariantCulture));
        }

        private static int InstallAllContextMenus(Form mainForm)
        {
            int count = 0;
            foreach (FieldInfo field in mainForm.GetType().GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
            {
                ContextMenuStrip menu = field.GetValue(mainForm) as ContextMenuStrip;
                if (menu == null)
                {
                    continue;
                }

                count++;
                menu.Opening -= ContextMenuOpening;
                menu.Opening += ContextMenuOpening;
                menu.Opened -= ContextMenuOpened;
                menu.Opened += ContextMenuOpened;
                MenuInfos[menu] = new MenuInfo { MainForm = mainForm, Name = field.Name };
                AddRecommendItemIfMatched(menu);
            }

            return count;
        }

        private static void ContextMenuOpening(object sender, CancelEventArgs e)
        {
            ContextMenuStrip menu = sender as ContextMenuStrip;
            AddRecommendItemIfMatched(menu);
            BeginAddRecommendItem(menu);
        }

        private static void ContextMenuOpened(object sender, EventArgs e)
        {
            ContextMenuStrip menu = sender as ContextMenuStrip;
            AddRecommendItemIfMatched(menu);
            BeginAddRecommendItem(menu);
        }

        private static void BeginAddRecommendItem(ContextMenuStrip menu)
        {
            if (menu == null)
            {
                return;
            }

            try
            {
                menu.BeginInvoke((MethodInvoker)delegate { AddRecommendItemIfMatched(menu); });
            }
            catch
            {
            }
        }

        private static void AddRecommendItemIfMatched(ContextMenuStrip menu)
        {
            if (menu == null || !MenuInfos.ContainsKey(menu))
            {
                return;
            }

            MenuInfo info = MenuInfos[menu];
            bool isQuotaMenu = info.Name == "contextMenuStripDE" || IsSource(menu, info.MainForm, "dataGridViewDE");
            if (!isQuotaMenu)
            {
                return;
            }

            AddRecommendItem(menu, info.MainForm);
        }

        private static void AddRecommendItem(ContextMenuStrip menu, Form mainForm)
        {
            ToolStripMenuItem item = FindMenuItem(menu, "\u63a8\u8350\u5b9a\u989d");
            if (item != null)
            {
                item.Visible = true;
                item.Available = true;
                item.Enabled = true;
                return;
            }

            int insertIndex = Math.Min(2, menu.Items.Count);
            item = new ToolStripMenuItem("\u63a8\u8350\u5b9a\u989d");
            item.Visible = true;
            item.Available = true;
            item.Enabled = true;
            item.Click += delegate { ShowRecommendDialog(mainForm); };
            menu.Items.Insert(insertIndex, item);
        }

        private static void ShowRecommendDialog(Form mainForm)
        {
            try
            {
                RecommendDialog dialog;
                if (!RecommendDialogs.TryGetValue(mainForm, out dialog) || dialog == null || dialog.IsDisposed)
                {
                    dialog = new RecommendDialog(mainForm, GetSelectionText(mainForm));
                    RecommendDialogs[mainForm] = dialog;
                    dialog.FormClosed += delegate { RecommendDialogs.Remove(mainForm); };
                    dialog.Show(mainForm);
                }
                else
                {
                    dialog.Show();
                    dialog.Activate();
                }
            }
            catch (Exception ex)
            {
                Log("Show recommend dialog failed: " + ex);
                MessageBox.Show(mainForm, ex.Message, "\u63a8\u8350\u5b9a\u989d", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string GetSelectionText(Form mainForm)
        {
            DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
            if (grid == null)
            {
                return "";
            }

            List<string> parts = new List<string>();
            foreach (DataGridViewCell cell in grid.SelectedCells)
            {
                string text = Convert.ToString(cell.FormattedValue, CultureInfo.CurrentCulture);
                if (!String.IsNullOrWhiteSpace(text))
                {
                    parts.Add(text.Trim());
                }
            }

            if (parts.Count == 0 && grid.CurrentRow != null)
            {
                foreach (DataGridViewCell cell in grid.CurrentRow.Cells)
                {
                    string text = Convert.ToString(cell.FormattedValue, CultureInfo.CurrentCulture);
                    if (!String.IsNullOrWhiteSpace(text))
                    {
                        parts.Add(text.Trim());
                    }
                }
            }

            return String.Join(" ", parts.Distinct().Take(12).ToArray());
        }

        private static ToolStripMenuItem FindMenuItem(ContextMenuStrip menu, string text)
        {
            foreach (ToolStripItem item in menu.Items)
            {
                ToolStripMenuItem menuItem = item as ToolStripMenuItem;
                if (menuItem != null && menuItem.Text == text)
                {
                    return menuItem;
                }
            }

            return null;
        }

        private static bool IsSource(ContextMenuStrip menu, Form mainForm, string fieldName)
        {
            Control source = menu.SourceControl;
            Control expected = GetField<Control>(mainForm, fieldName);
            return source != null && expected != null && Object.ReferenceEquals(source, expected);
        }

        private static Form FindMainForm()
        {
            foreach (Form form in Application.OpenForms)
            {
                if (form != null && form.GetType().FullName == "RecoNet.RecoMainForm")
                {
                    return form;
                }
            }

            return null;
        }

        private static T GetField<T>(object target, string name) where T : class
        {
            if (target == null)
            {
                return null;
            }

            FieldInfo field = target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            return field == null ? null : field.GetValue(target) as T;
        }

        internal static void Log(string message)
        {
            try
            {
                string dir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location);
                string path = Path.Combine(dir, "RecoQuotaRecommend.log");
                File.AppendAllText(path, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " " + message + Environment.NewLine);
            }
            catch
            {
            }
        }
    }

    internal sealed class RecommendDialog : Form
    {
        private readonly Form mainForm;
        private readonly DataGridView resultGrid;
        private readonly Label statusLabel;
        private readonly List<LearningRecord> records;
        private readonly List<RecommendationRow> recommendations = new List<RecommendationRow>();

        public RecommendDialog(Form owner, string initialQuery)
        {
            mainForm = owner;
            Text = "\u6279\u91cf\u63a8\u8350\u5b9a\u989d";
            StartPosition = FormStartPosition.CenterParent;
            Width = 1160;
            Height = 680;
            MinimizeBox = false;

            records = LearningStore.Load();

            Button readButton = new Button();
            readButton.Text = "\u91cd\u65b0\u8bfb\u53d6Excel\u6846\u9009";
            readButton.Left = 12;
            readButton.Top = 10;
            readButton.Width = 150;
            readButton.Click += delegate { ReadExcelSelectionAndRecommend(); };

            Button clipboardButton = new Button();
            clipboardButton.Text = "\u8bfb\u53d6\u526a\u8d34\u677f\u6846\u9009";
            clipboardButton.Left = 170;
            clipboardButton.Top = 10;
            clipboardButton.Width = 140;
            clipboardButton.Click += delegate { ReadClipboardAndRecommend(); };

            Button selectAllButton = new Button();
            selectAllButton.Text = "\u5168\u9009";
            selectAllButton.Left = 318;
            selectAllButton.Top = 10;
            selectAllButton.Width = 70;
            selectAllButton.Click += delegate { SetChecked(true); };

            Button clearButton = new Button();
            clearButton.Text = "\u5168\u4e0d\u9009";
            clearButton.Left = 396;
            clearButton.Top = 10;
            clearButton.Width = 80;
            clearButton.Click += delegate { SetChecked(false); };

            Button pasteButton = new Button();
            pasteButton.Text = "\u590d\u5236\u52fe\u9009\u5185\u5bb9";
            pasteButton.Left = 484;
            pasteButton.Top = 10;
            pasteButton.Width = 140;
            pasteButton.Click += delegate { CopyCheckedForManualPaste(); };

            resultGrid = new DataGridView();
            resultGrid.Left = 12;
            resultGrid.Top = 48;
            resultGrid.Width = 1120;
            resultGrid.Height = 555;
            resultGrid.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            resultGrid.AllowUserToAddRows = false;
            resultGrid.AllowUserToDeleteRows = false;
            resultGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            resultGrid.MultiSelect = true;
            resultGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            resultGrid.RowHeadersVisible = false;
            AddColumns();

            statusLabel = new Label();
            statusLabel.Left = 12;
            statusLabel.Top = 612;
            statusLabel.Width = 1120;
            statusLabel.Height = 36;
            statusLabel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            Controls.Add(readButton);
            Controls.Add(clipboardButton);
            Controls.Add(selectAllButton);
            Controls.Add(clearButton);
            Controls.Add(pasteButton);
            Controls.Add(resultGrid);
            Controls.Add(statusLabel);

            ReadExcelSelectionAndRecommend();
        }

        private void AddColumns()
        {
            DataGridViewCheckBoxColumn check = new DataGridViewCheckBoxColumn();
            check.Name = "Checked";
            check.HeaderText = "\u5199\u5165";
            check.Width = 45;
            resultGrid.Columns.Add(check);
            resultGrid.Columns.Add("QuantityName", "\u5de5\u7a0b\u91cf\u540d\u79f0");
            resultGrid.Columns.Add("QuantityUnit", "\u5355\u4f4d");
            resultGrid.Columns.Add("QuantityValue", "Excel\u5de5\u7a0b\u91cf");
            resultGrid.Columns.Add("QuotaQuantity", "\u5b9a\u989d\u5de5\u7a0b\u91cf");
            resultGrid.Columns.Add("QuotaCode", "\u63a8\u8350\u5b9a\u989d");
            resultGrid.Columns.Add("QuotaName", "\u5b9a\u989d\u540d\u79f0");
            resultGrid.Columns.Add("QuotaUnit", "\u5b9a\u989d\u5355\u4f4d");
            resultGrid.Columns["QuantityName"].FillWeight = 180;
            resultGrid.Columns["QuantityUnit"].FillWeight = 50;
            resultGrid.Columns["QuantityValue"].FillWeight = 80;
            resultGrid.Columns["QuotaQuantity"].FillWeight = 85;
            resultGrid.Columns["QuotaCode"].FillWeight = 80;
            resultGrid.Columns["QuotaName"].FillWeight = 210;
            resultGrid.Columns["QuotaUnit"].FillWeight = 60;

            foreach (DataGridViewColumn column in resultGrid.Columns)
            {
                if (column.Name != "Checked")
                {
                    column.ReadOnly = true;
                }
            }
        }

        private void ReadExcelSelectionAndRecommend()
        {
            recommendations.Clear();
            resultGrid.Rows.Clear();

            ExcelSelection selection;
            string error;
            if (!TryReadActiveExcelSelection(out selection, out error))
            {
                if (TryReadClipboardSelection(out selection, out error))
                {
                    FillRecommendations(selection);
                    statusLabel.Text = "\u672a\u627e\u5230 Excel/WPS COM\uff0c\u5df2\u6539\u4e3a\u8bfb\u53d6\u526a\u8d34\u677f\u6846\u9009\u5185\u5bb9\u3002\u8bf7\u5728 Excel \u6846\u9009\u540e Ctrl+C\uff0c\u518d\u70b9\u201c\u8bfb\u53d6\u526a\u8d34\u677f\u6846\u9009\u201d\u3002";
                    return;
                }

                statusLabel.Text = error + "\u53ef\u5728 Excel \u91cc\u6846\u9009\u540e\u6309 Ctrl+C\uff0c\u518d\u70b9\u201c\u8bfb\u53d6\u526a\u8d34\u677f\u6846\u9009\u201d\u3002";
                return;
            }

            FillRecommendations(selection);
        }

        private void ReadClipboardAndRecommend()
        {
            recommendations.Clear();
            resultGrid.Rows.Clear();

            ExcelSelection selection;
            string error;
            if (!TryReadClipboardSelection(out selection, out error))
            {
                statusLabel.Text = error;
                return;
            }

            FillRecommendations(selection);
        }

        private void FillRecommendations(ExcelSelection selection)
        {
            recommendations.Clear();
            resultGrid.Rows.Clear();

            foreach (ExcelQuantityItem item in selection.Items)
            {
                RecommendationRow recommendation = BuildRecommendation(item);
                recommendations.Add(recommendation);
                resultGrid.Rows.Add(
                    recommendation.Score >= 60,
                    item.Name,
                    item.Unit,
                    item.ValueText,
                    recommendation.ConvertedValueText,
                    recommendation.QuotaCode,
                    recommendation.QuotaName,
                    recommendation.QuotaUnit);
            }

            statusLabel.Text = String.Format(
                CultureInfo.CurrentCulture,
                "\u5df2\u8bfb\u53d6 {0} \u884cExcel\u5de5\u7a0b\u91cf\uff0c\u5b66\u4e60\u5e93 {1} \u6761\u3002\u9ed8\u8ba4\u52fe\u9009\u8f83\u53ef\u9760\u7684\u63a8\u8350\u3002",
                selection.Items.Count,
                records.Count);
        }

        private RecommendationRow BuildRecommendation(ExcelQuantityItem item)
        {
            ScoredRecord best = records
                .Select(r => new ScoredRecord { Record = r, Score = ScoreRecord(r, item) })
                .OrderByDescending(r => r.Score)
                .ThenByDescending(r => r.Record.MatchScore)
                .FirstOrDefault();

            RecommendationRow row = new RecommendationRow();
            row.Item = item;
            row.ConvertedValueText = item.ValueText;
            if (best == null || best.Score <= 0)
            {
                row.Score = 0;
                row.Reason = "\u672a\u627e\u5230\u76f8\u4f3c\u5b66\u4e60\u8bb0\u5f55";
                return row;
            }

            row.Record = best.Record;
            row.QuotaCode = best.Record.QuotaCode;
            row.QuotaName = best.Record.QuotaName;
            row.QuotaUnit = best.Record.QuotaUnit;
            row.ConvertedValueText = ConvertQuantityForQuotaUnit(item.ValueText, item.Unit, row.QuotaUnit);
            row.Score = best.Score;
            row.Reason = BuildReason(best.Record, item);
            return row;
        }

        private static int ScoreRecord(LearningRecord record, ExcelQuantityItem item)
        {
            string queryName = Normalize(item.Name);
            string context = Normalize(item.ContextText);
            string searchable = Normalize(String.Join(" ", new string[]
            {
                record.BudgetGroup,
                record.QuotaCode,
                record.QuotaName,
                record.QuotaUnit,
                record.QuantitySection,
                record.QuantityName,
                record.QuantityUnit,
                record.MatchReason
            }));

            int score = Math.Max(0, record.MatchScore / 3);
            string recordQuantityName = Normalize(record.QuantityName);
            string recordSection = Normalize(record.QuantitySection);
            string recordBudgetGroup = Normalize(record.BudgetGroup);

            if (!String.IsNullOrEmpty(queryName) && recordQuantityName == queryName)
            {
                score += 95;
            }
            else if (!String.IsNullOrEmpty(queryName) && (recordQuantityName.Contains(queryName) || queryName.Contains(recordQuantityName)))
            {
                score += 70;
            }

            if (!String.IsNullOrEmpty(item.SectionName) && (recordSection.Contains(Normalize(item.SectionName)) || recordBudgetGroup.Contains(Normalize(item.SectionName))))
            {
                score += 35;
            }

            if (UnitCompatible(record.QuantityUnit, item.Unit) || UnitCompatible(record.QuotaUnit, item.Unit))
            {
                score += 20;
            }

            foreach (string token in Tokenize(item.ContextText))
            {
                if (token.Length >= 2 && searchable.Contains(token))
                {
                    score += 12;
                }
            }

            return score;
        }

        private static string BuildReason(LearningRecord record, ExcelQuantityItem item)
        {
            List<string> reasons = new List<string>();
            if (UnitCompatible(record.QuantityUnit, item.Unit) || UnitCompatible(record.QuotaUnit, item.Unit))
            {
                reasons.Add("\u5355\u4f4d\u63a5\u8fd1");
            }
            if (Normalize(record.QuantityName).Contains(Normalize(item.Name)) || Normalize(item.Name).Contains(Normalize(record.QuantityName)))
            {
                reasons.Add("\u5de5\u7a0b\u91cf\u540d\u79f0\u76f8\u4f3c");
            }
            if (!String.IsNullOrWhiteSpace(record.MatchReason))
            {
                reasons.Add(record.MatchReason);
            }
            return reasons.Count == 0 ? "\u5b66\u4e60\u5e93\u76f8\u4f3c\u8bb0\u5f55" : String.Join(";", reasons.ToArray());
        }

        private void SetChecked(bool value)
        {
            foreach (DataGridViewRow row in resultGrid.Rows)
            {
                row.Cells["Checked"].Value = value;
            }
        }

        private void CopyCheckedForManualPaste()
        {
            List<RecommendationRow> rows = GetCheckedRecommendations();
            if (rows.Count == 0)
            {
                statusLabel.Text = "\u6ca1\u6709\u52fe\u9009\u4efb\u4f55\u63a8\u8350\u884c\u3002";
                return;
            }

            Clipboard.SetText(BuildTabSeparated(rows));
            statusLabel.Text = "\u5df2\u590d\u5236 " + rows.Count.ToString(CultureInfo.InvariantCulture) + " \u6761\u7c98\u8d34\u7528\u5185\u5bb9\uff1a\u7b2c1\u5217\u5b9a\u989d\u7f16\u53f7\uff0c\u7b2c4\u5217\u5de5\u7a0b\u6570\u91cf\u3002\u8bf7\u5728\u5b9a\u989d\u8868\u7b2c1\u5217\u76ee\u6807\u4f4d\u7f6e Ctrl+V\u3002";
        }

        private List<RecommendationRow> GetCheckedRecommendations()
        {
            List<RecommendationRow> rows = new List<RecommendationRow>();
            for (int i = 0; i < resultGrid.Rows.Count && i < recommendations.Count; i++)
            {
                object value = resultGrid.Rows[i].Cells["Checked"].Value;
                bool isChecked = value is bool && (bool)value;
                if (isChecked && !String.IsNullOrWhiteSpace(recommendations[i].QuotaCode))
                {
                    rows.Add(recommendations[i]);
                }
            }

            return rows;
        }

        private static string BuildTabSeparated(List<RecommendationRow> rows)
        {
            StringBuilder builder = new StringBuilder();
            foreach (RecommendationRow row in rows)
            {
                builder.Append(CleanCell(row.QuotaCode)).Append('\t')
                    .Append('\t')
                    .Append('\t')
                    .Append(CleanCell(row.ConvertedValueText))
                    .AppendLine();
            }

            return builder.ToString();
        }

        private static string CleanCell(string value)
        {
            return (value ?? "").Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim();
        }

        private static string ConvertQuantityForQuotaUnit(string quantityText, string excelUnit, string quotaUnit)
        {
            decimal quantity;
            if (!TryEvaluateQuantity(quantityText, out quantity))
            {
                return quantityText;
            }

            UnitScale excel = ParseUnitScale(excelUnit);
            UnitScale quota = ParseUnitScale(quotaUnit);
            if (String.IsNullOrEmpty(excel.BaseUnit) || String.IsNullOrEmpty(quota.BaseUnit))
            {
                return FormatDecimal(quantity);
            }

            if (!String.Equals(excel.BaseUnit, quota.BaseUnit, StringComparison.Ordinal))
            {
                return FormatDecimal(quantity);
            }

            if (excel.Scale <= 0 || quota.Scale <= 0)
            {
                return FormatDecimal(quantity);
            }

            decimal converted = quantity * excel.Scale / quota.Scale;
            return FormatDecimal(converted);
        }

        private static bool TryEvaluateQuantity(string text, out decimal value)
        {
            value = 0m;
            string expression = (text ?? "").Trim();
            if (expression.StartsWith("=", StringComparison.Ordinal))
            {
                expression = expression.Substring(1);
            }

            if (Decimal.TryParse(expression, NumberStyles.Float, CultureInfo.InvariantCulture, out value) ||
                Decimal.TryParse(expression, NumberStyles.Float, CultureInfo.CurrentCulture, out value))
            {
                return true;
            }

            expression = expression
                .Replace("\u00d7", "*")
                .Replace("x", "*")
                .Replace("X", "*")
                .Replace("\uff08", "(")
                .Replace("\uff09", ")");

            if (expression.Any(ch => !(Char.IsDigit(ch) || ch == '.' || ch == '+' || ch == '-' || ch == '*' || ch == '/' || ch == '(' || ch == ')' || Char.IsWhiteSpace(ch))))
            {
                return false;
            }

            try
            {
                DataTable table = new DataTable();
                object result = table.Compute(expression, String.Empty);
                value = Convert.ToDecimal(result, CultureInfo.InvariantCulture);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string FormatDecimal(decimal value)
        {
            return value.ToString("0.########", CultureInfo.InvariantCulture);
        }

        private static UnitScale ParseUnitScale(string unit)
        {
            string normalized = NormalizeRawUnit(unit);
            decimal scale = 1m;
            int index = 0;
            while (index < normalized.Length && (Char.IsDigit(normalized[index]) || normalized[index] == '.'))
            {
                index++;
            }

            if (index > 0)
            {
                Decimal.TryParse(normalized.Substring(0, index), NumberStyles.Float, CultureInfo.InvariantCulture, out scale);
            }

            string baseUnit = index > 0 ? normalized.Substring(index) : normalized;
            UnitScale result = new UnitScale();
            result.Scale = scale <= 0 ? 1m : scale;
            result.BaseUnit = baseUnit;
            return result;
        }

        private static string NormalizeRawUnit(string unit)
        {
            return (unit ?? "")
                .Trim()
                .ToLowerInvariant()
                .Replace(" ", "")
                .Replace("\u3000", "")
                .Replace("\u00b2", "2")
                .Replace("\u00b3", "3")
                .Replace("\uff4d", "m")
                .Replace("\u33a1", "m2")
                .Replace("\u33a5", "m3")
                .Replace("\u7acb\u65b9\u7c73", "m3")
                .Replace("\u5e73\u65b9\u7c73", "m2")
                .Replace("\u5ef6\u7c73", "m")
                .Replace("\u7c73", "m")
                .Replace("\u5428", "t")
                .Replace("\u5343\u514b", "kg");
        }

        private static bool TryReadActiveExcelSelection(out ExcelSelection selection, out string error)
        {
            selection = null;
            error = "";
            try
            {
                dynamic excel = GetActiveSpreadsheetApplication();
                if (excel == null)
                {
                    error = "\u6ca1\u6709\u627e\u5230\u6b63\u5728\u8fd0\u884c\u7684 Excel/WPS\u3002\u8bf7\u5148\u5728\u5de5\u7a0b\u91cf\u8868\u4e2d\u6846\u9009\u8981\u63a8\u8350\u7684\u884c\u3002";
                    return false;
                }

                dynamic workbook = excel.ActiveWorkbook;
                dynamic sheet = excel.ActiveSheet;
                dynamic range = excel.Selection;
                if (workbook == null || sheet == null || range == null)
                {
                    error = "\u8bf7\u5148\u5728 Excel/WPS \u4e2d\u6846\u9009\u5de5\u7a0b\u91cf\u884c\u3002";
                    return false;
                }

                selection = new ExcelSelection();
                selection.WorkbookPath = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture);
                selection.WorksheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture);

                int rowCount = Convert.ToInt32(range.Rows.Count, CultureInfo.InvariantCulture);
                int columnCount = Convert.ToInt32(range.Columns.Count, CultureInfo.InvariantCulture);
                for (int r = 1; r <= rowCount; r++)
                {
                    ExcelQuantityItem item = BuildQuantityItemFromRangeRow(range, selection.WorksheetName, r, columnCount);
                    if (item != null)
                    {
                        selection.Items.Add(item);
                    }
                }

                if (selection.Items.Count == 0)
                {
                    error = "\u6846\u9009\u533a\u57df\u91cc\u6ca1\u6709\u8bc6\u522b\u5230\u5de5\u7a0b\u91cf\u884c\uff0c\u8bf7\u628a\u5de5\u7a0b\u91cf\u540d\u79f0\u3001\u5355\u4f4d\u3001\u6570\u91cf\u4e00\u8d77\u6846\u9009\u3002";
                    return false;
                }

                return true;
            }
            catch (COMException)
            {
                error = "\u8bfb\u53d6 Excel/WPS \u6846\u9009\u533a\u57df\u5931\u8d25\uff0c\u8bf7\u786e\u8ba4\u8868\u683c\u5df2\u6253\u5f00\u5e76\u5df2\u6846\u9009\u5de5\u7a0b\u91cf\u884c\u3002";
                return false;
            }
            catch (Exception ex)
            {
                error = "\u8bfb\u53d6 Excel/WPS \u6846\u9009\u533a\u57df\u5931\u8d25\uff1a" + ex.Message;
                return false;
            }
        }

        private static bool TryReadClipboardSelection(out ExcelSelection selection, out string error)
        {
            selection = null;
            error = "";
            try
            {
                if (!Clipboard.ContainsText(TextDataFormat.UnicodeText) && !Clipboard.ContainsText())
                {
                    error = "\u526a\u8d34\u677f\u91cc\u6ca1\u6709 Excel \u6846\u9009\u5185\u5bb9\u3002\u8bf7\u5148\u5728 Excel/WPS \u91cc\u6846\u9009\u5de5\u7a0b\u91cf\u884c\u5e76\u6309 Ctrl+C\u3002";
                    return false;
                }

                string text = Clipboard.ContainsText(TextDataFormat.UnicodeText)
                    ? Clipboard.GetText(TextDataFormat.UnicodeText)
                    : Clipboard.GetText();
                if (String.IsNullOrWhiteSpace(text))
                {
                    error = "\u526a\u8d34\u677f\u5185\u5bb9\u4e3a\u7a7a\u3002";
                    return false;
                }

                selection = new ExcelSelection();
                selection.WorksheetName = "\u526a\u8d34\u677f";
                string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < lines.Length; i++)
                {
                    string[] cells = lines[i].Split('\t');
                    ExcelQuantityItem item = BuildQuantityItemFromTextRow(cells, i + 1);
                    if (item != null)
                    {
                        selection.Items.Add(item);
                    }
                }

                if (selection.Items.Count == 0)
                {
                    error = "\u526a\u8d34\u677f\u5185\u5bb9\u91cc\u6ca1\u6709\u8bc6\u522b\u5230\u5de5\u7a0b\u91cf\u884c\uff0c\u8bf7\u628a\u5de5\u7a0b\u91cf\u540d\u79f0\u3001\u5355\u4f4d\u3001\u6570\u91cf\u4e00\u8d77\u590d\u5236\u3002";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "\u8bfb\u53d6\u526a\u8d34\u677f\u5931\u8d25\uff1a" + ex.Message;
                return false;
            }
        }

        private static ExcelQuantityItem BuildQuantityItemFromRangeRow(dynamic range, string worksheetName, int relativeRow, int columnCount)
        {
            ExcelQuantityItem fixedItem = TryBuildThreeColumnRangeItem(range, worksheetName, relativeRow, columnCount);
            if (fixedItem != null)
            {
                return fixedItem;
            }

            List<CellValue> cells = new List<CellValue>();
            for (int c = 1; c <= columnCount; c++)
            {
                dynamic cell = range.Cells[relativeRow, c];
                string text = ExcelValueToText(cell.Value2);
                string formula = "";
                try
                {
                    formula = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture);
                }
                catch
                {
                }

                if (!String.IsNullOrWhiteSpace(text) || !String.IsNullOrWhiteSpace(formula))
                {
                    CellValue value = new CellValue();
                    value.Text = text;
                    value.Formula = formula;
                    value.Address = Convert.ToString(cell.Address(false, false), CultureInfo.InvariantCulture);
                    value.RowNumber = Convert.ToInt32(cell.Row, CultureInfo.InvariantCulture);
                    value.SourceIndex = c;
                    cells.Add(value);
                }
            }

            return BuildQuantityItemFromCells(cells, worksheetName);
        }

        private static ExcelQuantityItem TryBuildThreeColumnRangeItem(dynamic range, string worksheetName, int relativeRow, int columnCount)
        {
            if (columnCount < 3)
            {
                return null;
            }

            dynamic nameCell = range.Cells[relativeRow, 1];
            dynamic unitCell = range.Cells[relativeRow, 2];
            dynamic quantityCell = range.Cells[relativeRow, 3];
            string name = ExcelValueToText(nameCell.Value2);
            string unit = ExcelValueToText(unitCell.Value2);
            string quantity = ExcelValueToText(quantityCell.Value2);
            if (String.IsNullOrWhiteSpace(name) || String.IsNullOrWhiteSpace(quantity))
            {
                return null;
            }

            if (LooksLikeOrderOrHeader(name) || !IsQuantityLike(quantity))
            {
                return null;
            }

            ExcelQuantityItem item = new ExcelQuantityItem();
            item.WorksheetName = worksheetName;
            item.RowNumber = Convert.ToInt32(quantityCell.Row, CultureInfo.InvariantCulture);
            item.CellAddress = Convert.ToString(quantityCell.Address(false, false), CultureInfo.InvariantCulture);
            item.Name = name;
            item.Unit = unit;
            item.ValueText = quantity;
            try
            {
                item.Formula = Convert.ToString(quantityCell.Formula, CultureInfo.InvariantCulture);
            }
            catch
            {
                item.Formula = "";
            }
            item.ContextText = name + " " + unit + " " + quantity;
            item.SectionName = name;
            return item;
        }

        private static ExcelQuantityItem BuildQuantityItemFromTextRow(string[] rawCells, int rowNumber)
        {
            ExcelQuantityItem fixedItem = TryBuildThreeColumnClipboardItem(rawCells, rowNumber);
            if (fixedItem != null)
            {
                return fixedItem;
            }

            List<CellValue> cells = new List<CellValue>();
            for (int i = 0; i < rawCells.Length; i++)
            {
                string text = (rawCells[i] ?? "").Trim();
                if (!String.IsNullOrWhiteSpace(text))
                {
                    CellValue value = new CellValue();
                    value.Text = text;
                    value.Formula = text.StartsWith("=", StringComparison.Ordinal) ? text : "";
                    value.Address = "R" + rowNumber.ToString(CultureInfo.InvariantCulture) + "C" + (i + 1).ToString(CultureInfo.InvariantCulture);
                    value.RowNumber = rowNumber;
                    value.SourceIndex = i + 1;
                    cells.Add(value);
                }
            }

            return BuildQuantityItemFromCells(cells, "\u526a\u8d34\u677f");
        }

        private static ExcelQuantityItem TryBuildThreeColumnClipboardItem(string[] rawCells, int rowNumber)
        {
            if (rawCells == null || rawCells.Length < 3)
            {
                return null;
            }

            string name = (rawCells[0] ?? "").Trim();
            string unit = (rawCells[1] ?? "").Trim();
            string quantity = (rawCells[2] ?? "").Trim();
            if (String.IsNullOrWhiteSpace(name) || String.IsNullOrWhiteSpace(quantity))
            {
                return null;
            }

            if (LooksLikeOrderOrHeader(name) || !IsQuantityLike(quantity))
            {
                return null;
            }

            ExcelQuantityItem item = new ExcelQuantityItem();
            item.WorksheetName = "\u526a\u8d34\u677f";
            item.RowNumber = rowNumber;
            item.CellAddress = "R" + rowNumber.ToString(CultureInfo.InvariantCulture) + "C3";
            item.Name = name;
            item.Unit = unit;
            item.ValueText = quantity;
            item.Formula = quantity.StartsWith("=", StringComparison.Ordinal) ? quantity : "";
            item.ContextText = name + " " + unit + " " + quantity;
            item.SectionName = name;
            return item;
        }

        private static ExcelQuantityItem BuildQuantityItemFromCells(List<CellValue> cells, string worksheetName)
        {
            if (cells.Count == 0)
            {
                return null;
            }

            CellValue nameCell = null;
            CellValue unitCell = null;
            CellValue quantityCell = null;
            if (!TryPickByKnownLayout(cells, out nameCell, out unitCell, out quantityCell))
            {
                quantityCell = cells.LastOrDefault(c => IsQuantityLike(c.Text));
                if (quantityCell == null)
                {
                    return null;
                }

                unitCell = cells.LastOrDefault(c => c != quantityCell && LooksLikeUnit(c.Text));
                string fallbackName = PickQuantityName(cells, quantityCell, unitCell);
                if (!String.IsNullOrWhiteSpace(fallbackName))
                {
                    nameCell = cells.FirstOrDefault(c => c.Text == fallbackName);
                }
            }

            if (quantityCell == null)
            {
                return null;
            }

            string name = nameCell == null ? PickQuantityName(cells, quantityCell, unitCell) : nameCell.Text;
            if (String.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            ExcelQuantityItem item = new ExcelQuantityItem();
            item.WorksheetName = worksheetName;
            item.RowNumber = quantityCell.RowNumber;
            item.CellAddress = quantityCell.Address;
            item.Name = name;
            item.Unit = unitCell == null ? "" : unitCell.Text;
            item.ValueText = quantityCell.Text;
            item.Formula = quantityCell.Formula;
            item.ContextText = String.Join(" ", cells.Select(c => c.Text).Where(t => !String.IsNullOrWhiteSpace(t)).ToArray());
            item.SectionName = GuessSectionName(cells);
            return item;
        }

        private static bool TryPickByKnownLayout(List<CellValue> cells, out CellValue nameCell, out CellValue unitCell, out CellValue quantityCell)
        {
            nameCell = null;
            unitCell = null;
            quantityCell = null;

            int maxIndex = cells.Max(c => c.SourceIndex);
            int[][] layouts = new int[][]
            {
                new int[] { 2, 6, 7 },
                new int[] { 1, 5, 6 },
                new int[] { 1, 2, 3 }
            };

            foreach (int[] layout in layouts)
            {
                if (maxIndex < layout[2])
                {
                    continue;
                }

                CellValue n = cells.FirstOrDefault(c => c.SourceIndex == layout[0]);
                CellValue u = cells.FirstOrDefault(c => c.SourceIndex == layout[1]);
                CellValue q = cells.FirstOrDefault(c => c.SourceIndex == layout[2]);
                if (n != null && q != null && !IsNumberLike(n.Text) && IsQuantityLike(q.Text))
                {
                    nameCell = n;
                    unitCell = u != null && LooksLikeUnit(u.Text) ? u : cells.LastOrDefault(c => c.SourceIndex < q.SourceIndex && LooksLikeUnit(c.Text));
                    quantityCell = q;
                    return true;
                }
            }

            return false;
        }

        private static string PickQuantityName(List<CellValue> cells, CellValue quantityCell, CellValue unitCell)
        {
            IEnumerable<CellValue> candidates = cells.Where(c => c != quantityCell && c != unitCell && !IsNumberLike(c.Text) && !LooksLikeUnit(c.Text));
            CellValue best = candidates
                .Where(c => !LooksLikeOrderOrHeader(c.Text))
                .OrderByDescending(c => CountChineseLikeChars(c.Text))
                .ThenByDescending(c => (c.Text ?? "").Length)
                .FirstOrDefault();
            return best == null ? "" : best.Text;
        }

        private static string GuessSectionName(List<CellValue> cells)
        {
            foreach (CellValue cell in cells)
            {
                string text = cell.Text ?? "";
                if (text.Length >= 2 && text.Length <= 20 && !IsNumberLike(text) && !LooksLikeUnit(text))
                {
                    return text;
                }
            }

            return "";
        }

        private static object GetActiveSpreadsheetApplication()
        {
            string[] progIds = new string[] { "ket.Application", "KET.Application", "et.Application", "Excel.Application" };
            foreach (string progId in progIds)
            {
                try
                {
                    object app = Marshal.GetActiveObject(progId);
                    if (app != null)
                    {
                        return app;
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        private static string ExcelValueToText(object value)
        {
            if (value == null)
            {
                return "";
            }
            if (value is double || value is float)
            {
                return Convert.ToDouble(value, CultureInfo.InvariantCulture).ToString("0.########", CultureInfo.InvariantCulture);
            }
            if (value is decimal)
            {
                return ((decimal)value).ToString("0.########", CultureInfo.InvariantCulture);
            }
            return Convert.ToString(value, CultureInfo.CurrentCulture).Trim();
        }

        private static bool LooksLikeUnit(string text)
        {
            string unit = Normalize(text);
            string[] units = new string[] { "m", "m2", "m3", "kg", "t", "吨", "米", "平方米", "立方米", "处", "个", "座", "项", "根", "孔", "延米" };
            return units.Contains(unit);
        }

        private static bool UnitCompatible(string left, string right)
        {
            string l = NormalizeUnit(left);
            string r = NormalizeUnit(right);
            return !String.IsNullOrEmpty(l) && !String.IsNullOrEmpty(r) && (l == r || l.EndsWith(r, StringComparison.Ordinal) || r.EndsWith(l, StringComparison.Ordinal));
        }

        private static string NormalizeUnit(string unit)
        {
            return Normalize(unit)
                .Replace("100", "")
                .Replace("10", "")
                .Replace("立方米", "m3")
                .Replace("平方米", "m2")
                .Replace("米", "m")
                .Replace("吨", "t")
                .Replace("㎡", "m2")
                .Replace("m²", "m2")
                .Replace("m³", "m3");
        }

        private static bool IsNumberLike(string text)
        {
            decimal value;
            return Decimal.TryParse((text ?? "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                || Decimal.TryParse((text ?? "").Trim(), NumberStyles.Float, CultureInfo.CurrentCulture, out value);
        }

        private static bool IsQuantityLike(string text)
        {
            if (IsNumberLike(text))
            {
                return true;
            }

            string value = (text ?? "").Trim();
            bool hasDigit = value.Any(Char.IsDigit);
            bool hasOperator = value.IndexOfAny(new char[] { '+', '-', '*', '/', '×', '(', ')', '（', '）' }) >= 0;
            return hasDigit && hasOperator;
        }

        private static bool LooksLikeOrderOrHeader(string text)
        {
            string value = Normalize(text);
            return value == "\u5e8f\u53f7"
                || value == "\u7f16\u53f7"
                || value == "\u5355\u4f4d"
                || value == "\u5de5\u7a0b\u91cf"
                || value == "\u5de5\u7a0b\u9879\u76ee"
                || IsNumberLike(value);
        }

        private static int CountChineseLikeChars(string text)
        {
            int count = 0;
            foreach (char ch in text ?? "")
            {
                if (ch >= 0x4e00 && ch <= 0x9fff)
                {
                    count++;
                }
            }
            return count;
        }

        private static IEnumerable<string> Tokenize(string text)
        {
            foreach (string token in Normalize(text).Split(new char[] { ' ', '/', ',', ';', '\t', '(', ')', '[', ']', '（', '）' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (token.Length >= 2 && !IsNumberLike(token))
                {
                    yield return token;
                }
            }
        }

        private static string Normalize(string text)
        {
            return (text ?? "").Replace("\r", " ").Replace("\n", " ").Trim().ToLowerInvariant();
        }

        private static T GetField<T>(object target, string name) where T : class
        {
            if (target == null)
            {
                return null;
            }

            FieldInfo field = target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            return field == null ? null : field.GetValue(target) as T;
        }

        private sealed class ScoredRecord
        {
            public LearningRecord Record;
            public int Score;
        }

        private sealed class CellValue
        {
            public string Text;
            public string Formula;
            public string Address;
            public int RowNumber;
            public int SourceIndex;
        }
    }

    internal sealed class ExcelSelection
    {
        public string WorkbookPath;
        public string WorksheetName;
        public readonly List<ExcelQuantityItem> Items = new List<ExcelQuantityItem>();
    }

    internal sealed class ExcelQuantityItem
    {
        public string WorksheetName;
        public int RowNumber;
        public string CellAddress;
        public string Name;
        public string SectionName;
        public string Unit;
        public string ValueText;
        public string Formula;
        public string ContextText;
    }

    internal sealed class RecommendationRow
    {
        public ExcelQuantityItem Item;
        public LearningRecord Record;
        public string QuotaCode;
        public string QuotaName;
        public string QuotaUnit;
        public string ConvertedValueText;
        public int Score;
        public string Reason;
    }

    internal struct UnitScale
    {
        public string BaseUnit;
        public decimal Scale;
    }

    internal sealed class LearningRecord
    {
        public string ProjectName;
        public string BudgetFile;
        public string BudgetGroup;
        public string QuotaCode;
        public string QuotaName;
        public string QuotaUnit;
        public string QuantitySection;
        public string QuantityName;
        public string QuantityUnit;
        public string MatchReason;
        public int MatchScore;
    }

    internal static class LearningStore
    {
        public static List<LearningRecord> Load()
        {
            string path = FindLearningPath();
            if (String.IsNullOrEmpty(path))
            {
                return new List<LearningRecord>();
            }

            List<LearningRecord> records = new List<LearningRecord>();
            foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
            {
                if (String.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                Dictionary<string, string> item = ParseFlatJson(line);
                LearningRecord record = new LearningRecord();
                record.ProjectName = Get(item, "project_name");
                record.BudgetFile = Get(item, "budget_file");
                record.BudgetGroup = Get(item, "budget_group");
                record.QuotaCode = Get(item, "quota_code");
                record.QuotaName = Get(item, "quota_name");
                record.QuotaUnit = Get(item, "quota_unit");
                record.QuantitySection = Get(item, "quantity_section");
                record.QuantityName = Get(item, "quantity_name");
                record.QuantityUnit = Get(item, "quantity_unit");
                record.MatchReason = Get(item, "match_reason");
                int parsed;
                record.MatchScore = Int32.TryParse(Get(item, "match_score"), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) ? parsed : 0;
                if (!String.IsNullOrWhiteSpace(record.QuotaCode))
                {
                    records.Add(record);
                }
            }

            return records;
        }

        private static string FindLearningPath()
        {
            string baseDir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location);
            string[] candidates = new string[]
            {
                Path.Combine(baseDir, "RecoQuotaData", "learning.jsonl"),
                Path.Combine(Path.GetDirectoryName(baseDir), "RecoQuotaData", "learning.jsonl")
            };

            foreach (string candidate in candidates)
            {
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }

            return "";
        }

        private static string Get(Dictionary<string, string> values, string key)
        {
            string value;
            return values.TryGetValue(key, out value) ? value : "";
        }

        private static Dictionary<string, string> ParseFlatJson(string line)
        {
            Dictionary<string, string> result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            int index = 0;
            SkipWhitespace(line, ref index);
            if (index < line.Length && line[index] == '{')
            {
                index++;
            }

            while (index < line.Length)
            {
                SkipWhitespace(line, ref index);
                if (index >= line.Length || line[index] == '}')
                {
                    break;
                }

                string key = ReadJsonString(line, ref index);
                SkipWhitespace(line, ref index);
                if (index < line.Length && line[index] == ':')
                {
                    index++;
                }
                SkipWhitespace(line, ref index);
                string value = ReadJsonString(line, ref index);
                result[key] = value;
                SkipWhitespace(line, ref index);
                if (index < line.Length && line[index] == ',')
                {
                    index++;
                }
            }

            return result;
        }

        private static string ReadJsonString(string text, ref int index)
        {
            StringBuilder builder = new StringBuilder();
            if (index < text.Length && text[index] == '"')
            {
                index++;
            }

            while (index < text.Length)
            {
                char ch = text[index++];
                if (ch == '"')
                {
                    break;
                }

                if (ch == '\\' && index < text.Length)
                {
                    char escaped = text[index++];
                    if (escaped == 'n')
                    {
                        builder.Append('\n');
                    }
                    else if (escaped == 'r')
                    {
                        builder.Append('\r');
                    }
                    else if (escaped == 't')
                    {
                        builder.Append('\t');
                    }
                    else
                    {
                        builder.Append(escaped);
                    }
                }
                else
                {
                    builder.Append(ch);
                }
            }

            return builder.ToString();
        }

        private static void SkipWhitespace(string text, ref int index)
        {
            while (index < text.Length && Char.IsWhiteSpace(text[index]))
            {
                index++;
            }
        }
    }
}
