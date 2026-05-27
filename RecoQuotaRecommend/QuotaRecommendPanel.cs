using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
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
        private readonly ComboBox quotaCategoryCombo;
        private readonly List<LearningRecord> records;
        private readonly SearchIndexStore searchIndex;
        private readonly MappingStore mappingStore;
        private readonly List<RecommendationRow> recommendations = new List<RecommendationRow>();
        private ExcelSelection currentSelection;

        public RecommendDialog(Form owner, string initialQuery)
        {
            mainForm = owner;
            Text = "\u6279\u91cf\u63a8\u8350\u5b9a\u989d";
            StartPosition = FormStartPosition.CenterParent;
            Width = 1160;
            Height = 680;
            MinimizeBox = false;

            records = LearningStore.Load();
            LearningStore.BackupLearningFileIfNeeded();
            searchIndex = SearchIndexStore.LoadOrBuild();
            mappingStore = MappingStore.Load(records);

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

            Label categoryLabel = new Label();
            categoryLabel.Text = "\u5b9a\u989d\u7c7b\u578b";
            categoryLabel.Left = 632;
            categoryLabel.Top = 15;
            categoryLabel.Width = 58;

            quotaCategoryCombo = new ComboBox();
            quotaCategoryCombo.Left = 690;
            quotaCategoryCombo.Top = 11;
            quotaCategoryCombo.Width = 110;
            quotaCategoryCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            quotaCategoryCombo.Items.AddRange(new object[] { "\u9884\u7b97\u5b9a\u989d", "\u6982\u7b97\u5b9a\u989d", "\u4f30\u7b97\u5b9a\u989d", "\u5168\u90e8" });
            quotaCategoryCombo.SelectedIndex = 0;
            quotaCategoryCombo.SelectedIndexChanged += delegate
            {
                if (currentSelection != null)
                {
                    FillRecommendations(currentSelection);
                }
            };

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
            resultGrid.CellContentClick += ResultGridCellContentClick;
            resultGrid.CellEndEdit += ResultGridCellEndEdit;
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
            Controls.Add(categoryLabel);
            Controls.Add(quotaCategoryCombo);
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
            DataGridViewButtonColumn correct = new DataGridViewButtonColumn();
            correct.Name = "Correct";
            correct.HeaderText = "\u6276\u6b63";
            correct.Text = "\u6276\u6b63";
            correct.UseColumnTextForButtonValue = false;
            correct.Width = 55;
            resultGrid.Columns.Add(correct);
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
                if (column.Name != "Checked" && column.Name != "QuantityName")
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
            Stopwatch stopwatch = Stopwatch.StartNew();
            currentSelection = selection;
            recommendations.Clear();
            resultGrid.Rows.Clear();
            string categoryFilter = SelectedQuotaCategory;
            RecommendationBatchStats stats = new RecommendationBatchStats();
            Dictionary<string, List<RecommendationRow>> batchCache = new Dictionary<string, List<RecommendationRow>>(StringComparer.OrdinalIgnoreCase);

            resultGrid.SuspendLayout();
            try
            {
                foreach (ExcelQuantityItem item in selection.Items)
                {
                    string cacheKey = BuildBatchCacheKey(item, categoryFilter);
                    List<RecommendationRow> itemRecommendations;
                    List<RecommendationRow> cached;
                    if (batchCache.TryGetValue(cacheKey, out cached))
                    {
                        itemRecommendations = CloneRecommendationsForItem(item, cached);
                        stats.CacheHits++;
                    }
                    else
                    {
                        itemRecommendations = BuildRecommendations(item, categoryFilter, stats);
                        batchCache[cacheKey] = CloneRecommendationsForItem(item, itemRecommendations);
                    }

                    int itemRowIndex = 0;
                    foreach (RecommendationRow recommendation in itemRecommendations)
                    {
                        bool isContinuation = itemRowIndex > 0 && String.Equals(recommendation.Source, "mapping", StringComparison.OrdinalIgnoreCase);
                        recommendations.Add(recommendation);
                        int gridRowIndex = resultGrid.Rows.Add(
                            recommendation.Score >= 60,
                            isContinuation ? "" : "\u6276\u6b63",
                            isContinuation ? "" : item.Name,
                            item.Unit,
                            item.ValueText,
                            recommendation.ConvertedValueText,
                            recommendation.QuotaCode,
                            recommendation.QuotaName,
                            recommendation.QuotaUnit);
                        if (isContinuation)
                        {
                            DataGridViewRow gridRow = resultGrid.Rows[gridRowIndex];
                            gridRow.Cells["Correct"] = new DataGridViewTextBoxCell();
                            gridRow.Cells["Correct"].Value = "";
                            gridRow.Cells["Correct"].ReadOnly = true;
                        }
                        itemRowIndex++;
                    }
                }
            }
            finally
            {
                resultGrid.ResumeLayout();
            }

            stopwatch.Stop();

            statusLabel.Text = String.Format(
                CultureInfo.CurrentCulture,
                "\u5df2\u8bfb\u53d6 {0} \u884cExcel\u5de5\u7a0b\u91cf\uff0c\u5b9a\u989d\u7c7b\u578b\uff1a{1}\uff0c\u5bf9\u5e94\u6846\u547d\u4e2d {2} \u884c\uff0c\u7d22\u5f15\u68c0\u7d22 {3} \u884c\uff0c\u7a7a\u7ed3\u679c {4} \u884c\uff0c\u91cd\u590d\u590d\u7528 {5} \u884c\uff0c\u8017\u65f6 {6} ms\u3002",
                selection.Items.Count,
                categoryFilter,
                stats.MappingHits,
                stats.IndexSearches,
                stats.EmptyRows,
                stats.CacheHits,
                stopwatch.ElapsedMilliseconds);
        }

        private string SelectedQuotaCategory
        {
            get
            {
                return quotaCategoryCombo == null || quotaCategoryCombo.SelectedItem == null
                    ? "\u9884\u7b97\u5b9a\u989d"
                    : Convert.ToString(quotaCategoryCombo.SelectedItem, CultureInfo.CurrentCulture);
            }
        }

        private static string BuildBatchCacheKey(ExcelQuantityItem item, string categoryFilter)
        {
            return TextMatcher.Normalize(item == null ? "" : item.Name) + "|" + TextMatcher.Normalize(item == null ? "" : item.Unit) + "|" + TextMatcher.Normalize(categoryFilter);
        }

        private static List<RecommendationRow> CloneRecommendationsForItem(ExcelQuantityItem item, List<RecommendationRow> source)
        {
            List<RecommendationRow> rows = new List<RecommendationRow>();
            foreach (RecommendationRow original in source)
            {
                RecommendationRow row = new RecommendationRow();
                row.Item = item;
                row.Record = original.Record;
                row.QuotaCode = original.QuotaCode;
                row.QuotaName = original.QuotaName;
                row.QuotaUnit = original.QuotaUnit;
                row.ConvertedValueText = String.IsNullOrWhiteSpace(original.QuotaUnit)
                    ? (item == null ? original.ConvertedValueText : item.ValueText)
                    : ConvertQuantityForIndex(item == null ? "" : item.ValueText, item == null ? "" : item.Unit, original.QuotaUnit);
                row.Score = original.Score;
                row.Reason = original.Reason;
                row.Source = original.Source;
                row.TargetKind = original.TargetKind;
                row.BoxId = original.BoxId;
                rows.Add(row);
            }
            return rows;
        }

        private List<RecommendationRow> BuildRecommendations(ExcelQuantityItem item, string categoryFilter, RecommendationBatchStats stats)
        {
            List<RecommendationRow> mapped = mappingStore.Find(item, categoryFilter, searchIndex);
            if (mapped.Count > 0)
            {
                stats.MappingHits++;
                return mapped;
            }

            stats.IndexSearches++;
            List<RecommendationRow> indexed = searchIndex.Search(item, categoryFilter);
            if (indexed.Count > 0)
            {
                return indexed;
            }

            RecommendationRow empty = new RecommendationRow();
            empty.Item = item;
            empty.ConvertedValueText = item.ValueText;
            empty.Score = 0;
            empty.Reason = "\u672a\u5339\u914d\u5230\u5b9a\u989d\uff0c\u8bf7\u4eba\u5de5\u6276\u6b63";
            empty.Source = "empty";
            empty.TargetKind = "quota";
            stats.EmptyRows++;
            return new List<RecommendationRow> { empty };
        }

        private RecommendationRow BuildRecommendation(ExcelQuantityItem item)
        {
            ScoredRecord best = records
                .Where(r => !r.IsCorrection)
                .Where(r => String.Equals(QuotaEntry.GuessKind(r.QuotaCode), "quota", StringComparison.OrdinalIgnoreCase))
                .Select(r => new ScoredRecord { Record = r, Score = ScoreRecord(r, item) })
                .Where(r => r.Score > 0)
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
            row.Source = "learning";
            row.TargetKind = QuotaEntry.GuessKind(row.QuotaCode);
            return row;
        }

        private static RecommendationRow BuildRecommendationFromRecord(ExcelQuantityItem item, LearningRecord record, int score)
        {
            RecommendationRow row = new RecommendationRow();
            row.Item = item;
            row.Record = record;
            row.QuotaCode = record.QuotaCode;
            row.QuotaName = record.QuotaName;
            row.QuotaUnit = record.QuotaUnit;
            row.ConvertedValueText = ConvertQuantityForQuotaUnit(item.ValueText, item.Unit, row.QuotaUnit);
            row.Score = score;
            row.Reason = "\u4eba\u5de5\u6276\u6b63";
            row.Source = "learning";
            row.TargetKind = QuotaEntry.GuessKind(row.QuotaCode);
            return row;
        }

        private static int ScoreRecord(LearningRecord record, ExcelQuantityItem item)
        {
            if (TextMatcher.IsSteelOnlyAgainstConcrete(item.Name, record.QuantityName) || TextMatcher.IsSteelOnlyAgainstConcrete(item.Name, record.QuotaName))
            {
                return 0;
            }

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
            bool textMatched = TextMatcher.HasStrongNamePairMatch(item.Name, record.QuantityName);

            if (!String.IsNullOrEmpty(queryName) && recordQuantityName == queryName)
            {
                score += 95;
            }
            else if (!String.IsNullOrEmpty(queryName) && (recordQuantityName.Contains(queryName) || queryName.Contains(recordQuantityName)))
            {
                score += 70;
            }
            else if (textMatched)
            {
                score += Math.Min(70, TextMatcher.NamePairScore(item.Name, record.QuantityName));
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

            return textMatched ? score : 0;
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

        private void ResultGridCellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
            {
                return;
            }

            if (resultGrid.Columns[e.ColumnIndex].Name == "Correct")
            {
                if (!(resultGrid.Rows[e.RowIndex].Cells[e.ColumnIndex] is DataGridViewButtonCell))
                {
                    return;
                }
                CorrectRecommendation(e.RowIndex);
            }
        }

        private void ResultGridCellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= recommendations.Count || e.ColumnIndex < 0)
            {
                return;
            }

            if (resultGrid.Columns[e.ColumnIndex].Name != "QuantityName")
            {
                return;
            }

            object value = resultGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            string newName = Convert.ToString(value, CultureInfo.CurrentCulture).Trim();
            if (String.IsNullOrWhiteSpace(newName))
            {
                resultGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = recommendations[e.RowIndex].Item.Name;
                return;
            }

            ExcelQuantityItem oldItem = recommendations[e.RowIndex].Item;
            string oldName = oldItem.Name;
            foreach (RecommendationRow row in recommendations.Where(r => Object.ReferenceEquals(r.Item, oldItem)))
            {
                row.Item.Name = newName;
                row.Item.SectionName = newName;
                row.Item.ContextText = ReplaceContextName(row.Item.ContextText, oldName, newName);
            }
        }

        private void CorrectRecommendation(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= recommendations.Count)
            {
                return;
            }

            RecommendationRow recommendation = recommendations[rowIndex];
            List<QuotaEntry> quotas = GetSelectedQuotaEntries(mainForm);
            if (quotas.Count == 0)
            {
                MessageBox.Show(this, "\u8bf7\u5148\u5728\u5b9a\u989d\u8f93\u5165\u8868\u4e2d\u9009\u4e2d\u4e00\u6761\u6216\u591a\u6761\u6b63\u786e\u7684\u5b9a\u989d\uff0c\u518d\u70b9\u51fb\u6276\u6b63\u3002", "\u6276\u6b63\u5b9a\u989d", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                mappingStore.Correct(recommendation.Item, recommendation, quotas);
                LearningStore.ReplaceCorrections(recommendation.Item, quotas);
                records.Clear();
                records.AddRange(LearningStore.Load());
                if (currentSelection != null)
                {
                    FillRecommendations(currentSelection);
                }

                statusLabel.Text = "\u5df2\u6276\u6b63\uff1a" + recommendation.Item.Name + " -> " + String.Join("\uff0c", quotas.Select(q => q.QuotaCode).ToArray()) + "\u3002\u4e0b\u6b21\u63a8\u8350\u5c06\u4f18\u5148\u4f7f\u7528\u8be5\u5bf9\u5e94\u5173\u7cfb\u3002";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "\u6276\u6b63\u5931\u8d25\uff1a" + ex.Message, "\u6276\u6b63\u5b9a\u989d", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string ReplaceContextName(string context, string oldName, string newName)
        {
            if (String.IsNullOrWhiteSpace(context))
            {
                return newName;
            }

            if (!String.IsNullOrWhiteSpace(oldName) && context.IndexOf(oldName, StringComparison.Ordinal) >= 0)
            {
                return context.Replace(oldName, newName);
            }

            return newName + " " + context;
        }

        private static List<QuotaEntry> GetSelectedQuotaEntries(Form mainForm)
        {
            List<QuotaEntry> result = new List<QuotaEntry>();
            DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
            if (grid == null)
            {
                return result;
            }

            SortedDictionary<int, DataGridViewRow> rows = new SortedDictionary<int, DataGridViewRow>();
            foreach (DataGridViewRow row in grid.SelectedRows)
            {
                if (row != null && !row.IsNewRow)
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

            foreach (DataGridViewRow row in rows.Values)
            {
                QuotaEntry entry = new QuotaEntry();
                entry.QuotaCode = GetRowValue(row, "\u5b9a\u989d\u7f16\u53f7", "\u5b9a\u989d\u7f16\u53f7DE", "\u7f16\u53f7");
                entry.QuotaName = GetRowValue(row, "\u5de5\u7a0b\u6216\u8d39\u7528\u9879\u76ee\u540d\u79f0", "\u540d\u79f0", "\u9879\u76ee\u540d\u79f0");
                entry.QuotaUnit = GetRowValue(row, "\u5355\u4f4d");
                if (!String.IsNullOrWhiteSpace(entry.QuotaCode))
                {
                    entry.TargetKind = QuotaEntry.GuessKind(entry.QuotaCode);
                    result.Add(entry);
                }
            }

            return result
                .GroupBy(q => q.QuotaCode, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();
        }

        private static string GetRowValue(DataGridViewRow row, params string[] names)
        {
            DataRowView rowView = row.DataBoundItem as DataRowView;
            if (rowView != null)
            {
                foreach (string name in names)
                {
                    if (rowView.DataView.Table.Columns.Contains(name))
                    {
                        object value = rowView[name];
                        if (value != null && value != DBNull.Value)
                        {
                            return Convert.ToString(value, CultureInfo.CurrentCulture).Trim();
                        }
                    }
                }
            }

            if (row.DataGridView != null)
            {
                foreach (DataGridViewColumn column in row.DataGridView.Columns)
                {
                    foreach (string name in names)
                    {
                        if (String.Equals(column.Name, name, StringComparison.OrdinalIgnoreCase) ||
                            String.Equals(column.HeaderText, name, StringComparison.OrdinalIgnoreCase))
                        {
                            object value = row.Cells[column.Index].Value;
                            if (value != null)
                            {
                                return Convert.ToString(value, CultureInfo.CurrentCulture).Trim();
                            }
                        }
                    }
                }
            }

            return "";
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
            mappingStore.Accept(rows);
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

        internal static string ConvertQuantityForIndex(string quantityText, string excelUnit, string quotaUnit)
        {
            return ConvertQuantityForQuotaUnit(quantityText, excelUnit, quotaUnit);
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

        internal static bool UnitCompatibleForIndex(string left, string right)
        {
            return UnitCompatible(left, right);
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

        private sealed class RecommendationBatchStats
        {
            public int MappingHits;
            public int IndexSearches;
            public int EmptyRows;
            public int CacheHits;
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
        public string Source;
        public string BoxId;
        public string TargetKind;
    }

    internal sealed class QuotaEntry
    {
        public string TargetKind;
        public string QuotaCode;
        public string QuotaName;
        public string QuotaUnit;

        public static string GuessKind(string code)
        {
            string value = (code ?? "").Trim();
            return value.Length > 0 && value.All(Char.IsDigit) ? "material" : "quota";
        }

        public string TargetKey
        {
            get { return (String.IsNullOrWhiteSpace(TargetKind) ? GuessKind(QuotaCode) : TargetKind) + ":" + (QuotaCode ?? "").Trim().ToUpperInvariant(); }
        }
    }

    internal struct UnitScale
    {
        public string BaseUnit;
        public decimal Scale;
    }

    internal sealed class LearningRecord
    {
        public bool IsCorrection;
        public string QuantitySignature;
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

    internal sealed class SearchIndexStore
    {
        private readonly List<IndexQuota> quotas = new List<IndexQuota>();
        private readonly List<IndexMaterial> materials = new List<IndexMaterial>();
        private readonly Dictionary<string, List<IndexQuota>> quotaTokenIndex = new Dictionary<string, List<IndexQuota>>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, IndexQuota> quotasByCode = new Dictionary<string, IndexQuota>(StringComparer.OrdinalIgnoreCase);

        public int QuotaCount { get { return quotas.Count; } }
        public int MaterialCount { get { return materials.Count; } }

        public static SearchIndexStore LoadOrBuild()
        {
            SearchIndexStore store = new SearchIndexStore();
            string dataDir = LearningStore.FindDataDir();
            string quotaPath = Path.Combine(dataDir, "quota-index.jsonl");
            string materialPath = Path.Combine(dataDir, "material-index.jsonl");

            if (!File.Exists(quotaPath) || !File.Exists(materialPath))
            {
                try
                {
                    ExportFromSql(dataDir, quotaPath, materialPath);
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Build search index failed: " + ex.Message);
                }
            }

            store.LoadFiles(quotaPath, materialPath);
            return store;
        }

        public List<RecommendationRow> Search(ExcelQuantityItem item, string categoryFilter)
        {
            List<IndexQuota> candidates = GetQuotaCandidates(item, categoryFilter);
            if (candidates.Count == 0)
            {
                return new List<RecommendationRow>();
            }

            List<ScoredQuota> quotaHits = candidates
                .Select(q => new ScoredQuota { Quota = q, Score = ScoreQuota(item, q) })
                .Where(q => q.Score >= 55)
                .OrderByDescending(q => q.Score)
                .ThenBy(q => q.Quota.SortOrder)
                .Take(1)
                .ToList();

            List<RecommendationRow> rows = new List<RecommendationRow>();
            foreach (ScoredQuota hit in quotaHits)
            {
                rows.Add(hit.Quota.ToRecommendation(item, hit.Score));
            }

            return rows
                .GroupBy(r => (r.TargetKind ?? "") + ":" + (r.QuotaCode ?? ""), StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .OrderByDescending(r => r.Score)
                .Take(4)
                .ToList();
        }

        public bool IsMappingTargetAllowed(string targetKind, string code, string categoryFilter)
        {
            string kind = String.IsNullOrWhiteSpace(targetKind) ? QuotaEntry.GuessKind(code) : targetKind;
            if (!String.Equals(kind, "quota", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            IndexQuota quota;
            if (!quotasByCode.TryGetValue((code ?? "").Trim(), out quota))
            {
                return false;
            }

            return CategoryAllowed(quota.BookCategory, categoryFilter);
        }

        private void LoadFiles(string quotaPath, string materialPath)
        {
            if (File.Exists(quotaPath))
            {
                foreach (string line in File.ReadAllLines(quotaPath, Encoding.UTF8))
                {
                    if (String.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                    IndexQuota quota = new IndexQuota();
                    quota.QuotaCode = LearningStore.Get(values, "quota_code");
                    quota.QuotaName = LearningStore.Get(values, "quota_name");
                    quota.QuotaUnit = LearningStore.Get(values, "quota_unit");
                    quota.BookCode = LearningStore.Get(values, "book_code");
                    quota.BookCategory = LearningStore.Get(values, "book_category");
                    quota.Specialty = LearningStore.Get(values, "specialty");
                    quota.SectionNo = LearningStore.Get(values, "section_no");
                    quota.SectionName = LearningStore.Get(values, "section_name");
                    quota.WorkContent = LearningStore.Get(values, "work_content");
                    quota.SearchText = LearningStore.Get(values, "search_text");
                    int sortOrder;
                    quota.SortOrder = Int32.TryParse(LearningStore.Get(values, "sort_order"), NumberStyles.Integer, CultureInfo.InvariantCulture, out sortOrder) ? sortOrder : Int32.MaxValue;
                    if (!String.IsNullOrWhiteSpace(quota.QuotaCode))
                    {
                        quotas.Add(quota);
                        if (!quotasByCode.ContainsKey(quota.QuotaCode.Trim()))
                        {
                            quotasByCode[quota.QuotaCode.Trim()] = quota;
                        }
                    }
                }

                BuildQuotaTokenIndex();
            }

            if (File.Exists(materialPath))
            {
                foreach (string line in File.ReadAllLines(materialPath, Encoding.UTF8))
                {
                    if (String.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                    IndexMaterial material = new IndexMaterial();
                    material.MaterialCode = LearningStore.Get(values, "material_code");
                    material.MaterialName = LearningStore.Get(values, "material_name");
                    material.MaterialUnit = LearningStore.Get(values, "material_unit");
                    material.DocNo = LearningStore.Get(values, "doc_no");
                    material.IsMainMaterial = LearningStore.Get(values, "is_main_material") == "1";
                    material.TransportCategory = LearningStore.Get(values, "transport_category");
                    material.SearchText = LearningStore.Get(values, "search_text");
                    if (!String.IsNullOrWhiteSpace(material.MaterialCode))
                    {
                        materials.Add(material);
                    }
                }
            }
        }

        private void BuildQuotaTokenIndex()
        {
            quotaTokenIndex.Clear();
            foreach (IndexQuota quota in quotas)
            {
                foreach (string token in QuotaIndexTokens(quota))
                {
                    List<IndexQuota> bucket;
                    if (!quotaTokenIndex.TryGetValue(token, out bucket))
                    {
                        bucket = new List<IndexQuota>();
                        quotaTokenIndex[token] = bucket;
                    }
                    bucket.Add(quota);
                }
            }
        }

        private static IEnumerable<string> QuotaIndexTokens(IndexQuota quota)
        {
            string text = String.Join(" ", new string[] { quota.QuotaName, quota.WorkContent, quota.SectionName, quota.Specialty });
            foreach (string token in TextMatcher.Keywords(text).Distinct())
            {
                if (UseTokenForCandidateLookup(token))
                {
                    yield return token;
                }
            }
        }

        private List<IndexQuota> GetQuotaCandidates(ExcelQuantityItem item, string categoryFilter)
        {
            HashSet<IndexQuota> candidates = new HashSet<IndexQuota>();
            foreach (string token in CandidateLookupTokens(item.Name))
            {
                List<IndexQuota> bucket;
                if (!quotaTokenIndex.TryGetValue(token, out bucket))
                {
                    continue;
                }

                foreach (IndexQuota quota in bucket)
                {
                    if (CategoryAllowed(quota.BookCategory, categoryFilter))
                    {
                        candidates.Add(quota);
                    }
                }
            }

            return candidates.ToList();
        }

        private static IEnumerable<string> CandidateLookupTokens(string quantityName)
        {
            string normalized = TextMatcher.Normalize(quantityName).Replace(" ", "");
            if (UseTokenForCandidateLookup(normalized))
            {
                yield return normalized;
            }

            foreach (string token in TextMatcher.Keywords(quantityName).Distinct())
            {
                if (UseTokenForCandidateLookup(token))
                {
                    yield return token;
                }
            }
        }

        private static bool UseTokenForCandidateLookup(string token)
        {
            if (String.IsNullOrWhiteSpace(token) || token.Length < 2)
            {
                return false;
            }

            return !TextMatcher.IsNumberLikeToken(token);
        }

        private static bool CategoryAllowed(string bookCategory, string categoryFilter)
        {
            string category = (bookCategory ?? "").Trim();
            string filter = (categoryFilter ?? "").Trim();
            if (String.IsNullOrWhiteSpace(filter) || String.Equals(filter, "\u5168\u90e8", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (String.Equals(category, filter, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return IsCommonCategory(category);
        }

        private static bool IsCommonCategory(string category)
        {
            return String.IsNullOrWhiteSpace(category) ||
                String.Equals(category, "\u57fa\u672c\u5b9a\u989d", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(category, "\u8865\u5145\u5b9a\u989d", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(category, "\u8865\u5145\u5355\u4ef7\u5206\u6790", StringComparison.OrdinalIgnoreCase);
        }

        private static int ScoreQuota(ExcelQuantityItem item, IndexQuota quota)
        {
            string query = TextMatcher.Normalize(item.Name);
            string nameText = TextMatcher.Normalize(quota.QuotaName);
            string workText = TextMatcher.Normalize(quota.WorkContent);
            string sectionText = TextMatcher.Normalize(quota.SectionName);
            string specialtyText = TextMatcher.Normalize(quota.Specialty);
            string searchable = TextMatcher.Normalize(quota.SearchText);
            if (String.IsNullOrWhiteSpace(query) || String.IsNullOrWhiteSpace(searchable))
            {
                return 0;
            }

            if (IsSteelQuantity(query) && IsConcreteQuotaName(nameText))
            {
                return 0;
            }

            int score = 0;
            int primaryMatches = 0;
            int coreMatches = 0;
            if (nameText.Contains(query))
            {
                score += 90;
                primaryMatches++;
            }
            else if (workText.Contains(query))
            {
                score += 65;
                primaryMatches++;
            }

            List<string> tokens = TextMatcher.Keywords(item.Name).Distinct().ToList();
            int matched = 0;
            foreach (string token in tokens)
            {
                if (token.Length < 1 || !searchable.Contains(token))
                {
                    continue;
                }

                matched++;
                bool coreToken = TextMatcher.IsPureChinese(token) && token.Length >= 2;
                if (coreToken)
                {
                    coreMatches++;
                }

                bool primaryHit = nameText.Contains(token) || workText.Contains(token);
                if (primaryHit && token.Length >= 2)
                {
                    primaryMatches++;
                }

                if (nameText.Contains(token))
                {
                    score += TokenScore(token, 55, 42, 28);
                }
                else if (workText.Contains(token))
                {
                    score += TokenScore(token, 36, 28, 18);
                }
                else if (sectionText.Contains(token))
                {
                    score += TokenScore(token, 18, 14, 9);
                }
                else if (specialtyText.Contains(token))
                {
                    score += TokenScore(token, 10, 8, 5);
                }
            }

            if (IsSteelQuantity(query))
            {
                score += SteelPreferenceScore(query, nameText, workText, quota.QuotaUnit);
            }

            if (primaryMatches == 0 && coreMatches < 2)
            {
                return 0;
            }

            if (tokens.Count > 0)
            {
                score += (int)Math.Round(20.0 * matched / tokens.Count);
            }

            if (RecommendDialog.UnitCompatibleForIndex(quota.QuotaUnit, item.Unit))
            {
                score += 12;
            }

            return score;
        }

        private static int TokenScore(string token, int shortChinese, int longChinese, int mixed)
        {
            if (TextMatcher.HasAsciiOrDigit(token))
            {
                return mixed;
            }

            if (token.Length == 1)
            {
                return Math.Max(1, shortChinese / 10);
            }

            return token.Length == 2 ? shortChinese : longChinese;
        }

        private static bool IsSteelQuantity(string normalizedQuery)
        {
            return TextMatcher.IsSteelQuantityName(normalizedQuery);
        }

        private static bool IsConcreteQuotaName(string normalizedQuotaName)
        {
            if (!TextMatcher.IsConcreteQuantityName(normalizedQuotaName))
            {
                return false;
            }

            return !normalizedQuotaName.Contains("构件钢筋") &&
                !normalizedQuotaName.Contains("圆钢筋") &&
                !normalizedQuotaName.Contains("螺纹钢筋") &&
                !normalizedQuotaName.Contains("箍筋") &&
                !normalizedQuotaName.Contains("钢筋制作") &&
                !normalizedQuotaName.Contains("钢筋制安") &&
                !normalizedQuotaName.Contains("钢筋绑扎");
        }

        private static int SteelPreferenceScore(string query, string nameText, string workText, string quotaUnit)
        {
            int score = 0;
            bool steelOperation = nameText.Contains("构件钢筋") ||
                nameText.Contains("钢筋制作") ||
                nameText.Contains("钢筋制安") ||
                nameText.Contains("钢筋绑扎") ||
                workText.Contains("钢筋制作") ||
                workText.Contains("钢筋制安") ||
                (workText.Contains("钢筋") && (workText.Contains("制作") || workText.Contains("绑扎") || workText.Contains("安装")));

            if (steelOperation)
            {
                score += 55;
            }

            if ((query.Contains("hpb") || query.Contains("光圆") || query.Contains("圆钢")) && nameText.Contains("圆钢筋"))
            {
                score += 80;
            }

            if ((query.Contains("hrb") || query.Contains("螺纹")) && (nameText.Contains("hrb") || nameText.Contains("螺纹钢筋")))
            {
                score += 80;
            }

            if (nameText.Contains("钢筋混凝土") && !steelOperation)
            {
                score -= 70;
            }

            if (!RecommendDialog.UnitCompatibleForIndex(quotaUnit, "kg") && !RecommendDialog.UnitCompatibleForIndex(quotaUnit, "t"))
            {
                score -= 30;
            }

            return score;
        }

        private static void ExportFromSql(string dataDir, string quotaPath, string materialPath)
        {
            Directory.CreateDirectory(dataDir);
            string server = ReadServer();
            if (String.IsNullOrWhiteSpace(server))
            {
                server = "127.0.0.1";
            }

            string databaseName = ResolveDatabaseName();
            QuotaRecommendPanel.Log("Build search index from database: " + databaseName);
            string connectionString = "Data Source=" + server + ",1433;Initial Catalog=" + databaseName + ";User ID=reco;Password=" + BuildSqlPassword() + ";Connect Timeout=8;Encrypt=False;TrustServerCertificate=True";
            using (System.Data.SqlClient.SqlConnection connection = new System.Data.SqlClient.SqlConnection(connectionString))
            {
                connection.Open();
                WriteQuotaIndex(connection, quotaPath);
                WriteMaterialIndex(connection, materialPath);
            }
        }

        private static string ReadServer()
        {
            try
            {
                string baseDir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location);
                string path = Path.Combine(baseDir, "ServerSetting.xml");
                if (!File.Exists(path))
                {
                    return "";
                }

                string text = File.ReadAllText(path, Encoding.UTF8);
                int start = text.IndexOf("<ServerIP>", StringComparison.OrdinalIgnoreCase);
                int end = text.IndexOf("</ServerIP>", StringComparison.OrdinalIgnoreCase);
                if (start >= 0 && end > start)
                {
                    return text.Substring(start + 10, end - start - 10).Trim();
                }

                start = text.IndexOf("<Server>", StringComparison.OrdinalIgnoreCase);
                end = text.IndexOf("</Server>", StringComparison.OrdinalIgnoreCase);
                return start >= 0 && end > start ? text.Substring(start + 8, end - start - 8).Trim() : "";
            }
            catch
            {
                return "";
            }
        }

        private static string BuildSqlPassword()
        {
            return String.Join("_", new string[] { "Des", "Reco", "2006" });
        }

        private static string ResolveDatabaseName()
        {
            try
            {
                string baseDir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location) ?? "";
                string processPath = "";
                try
                {
                    processPath = Process.GetCurrentProcess().MainModule.FileName ?? "";
                }
                catch
                {
                }

                string probe = (baseDir + " " + processPath).ToLowerInvariant();
                if (probe.Contains("2024") ||
                    File.Exists(Path.Combine(baseDir, "ReJJGSNet2024.exe")) ||
                    File.Exists(Path.Combine(baseDir, "ReJJQDNet2024.exe")))
                {
                    return "RecoData2024";
                }
            }
            catch
            {
            }

            return "RecoData2020";
        }

        private static void WriteQuotaIndex(System.Data.SqlClient.SqlConnection connection, string path)
        {
            string temp = path + ".tmp";
            using (StreamWriter writer = new StreamWriter(temp, false, Encoding.UTF8))
            using (System.Data.SqlClient.SqlCommand command = connection.CreateCommand())
            {
                command.CommandTimeout = 60;
                command.CommandText =
                    "SELECT a.定额编号,a.定额名称,a.单位,a.书号,ISNULL(b.分类,''),ISNULL(b.专业名称,''),ISNULL(a.节号,''),ISNULL(c.节名称,''),ISNULL(CAST(a.工作内容 AS nvarchar(max)),''),ISNULL(a.基价,0),ISNULL(b.现行定额,0),ISNULL(a.流水号,2147483647) " +
                    "FROM dbo.定额库 a LEFT JOIN dbo.定额库索引 b ON a.书号=b.书号 LEFT JOIN dbo.定额节索引 c ON a.书号=c.书号 AND a.节号=c.节号 " +
                    "WHERE ISNULL(a.定额编号,'')<>'' AND ISNULL(a.定额名称,'')<>''";
                using (System.Data.SqlClient.SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        Dictionary<string, string> row = new Dictionary<string, string>();
                        row["quota_code"] = ReadString(reader, 0);
                        row["quota_name"] = ReadString(reader, 1);
                        row["quota_unit"] = ReadString(reader, 2);
                        row["book_code"] = ReadString(reader, 3);
                        row["book_category"] = ReadString(reader, 4);
                        row["specialty"] = ReadString(reader, 5);
                        row["section_no"] = ReadString(reader, 6);
                        row["section_name"] = ReadString(reader, 7);
                        row["work_content"] = ReadString(reader, 8);
                        row["base_price"] = Convert.ToString(reader.GetDouble(9), CultureInfo.InvariantCulture);
                        row["is_current"] = IsTruthy(reader.GetValue(10)) ? "1" : "0";
                        row["sort_order"] = Convert.ToString(reader.GetInt32(11), CultureInfo.InvariantCulture);
                        row["search_text"] = TextMatcher.Normalize(String.Join(" ", new string[] { row["quota_code"], row["quota_name"], row["quota_unit"], row["book_category"], row["specialty"], row["section_name"], row["work_content"] }));
                        writer.WriteLine(LearningStore.ToJson(row));
                    }
                }
            }

            ReplaceFile(temp, path);
        }

        private static void WriteMaterialIndex(System.Data.SqlClient.SqlConnection connection, string path)
        {
            string temp = path + ".tmp";
            using (StreamWriter writer = new StreamWriter(temp, false, Encoding.UTF8))
            using (System.Data.SqlClient.SqlCommand command = connection.CreateCommand())
            {
                command.CommandTimeout = 60;
                command.CommandText =
                    "SELECT 电算代号,材料名称,单位,文号,ISNULL(主材标志,''),ISNULL(材料运输类别,''),ISNULL(基期单价,0),ISNULL(编制期价,0) " +
                    "FROM dbo.材料单价库 WHERE ISNULL(材料名称,'')<>''";
                using (System.Data.SqlClient.SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        Dictionary<string, string> row = new Dictionary<string, string>();
                        row["material_code"] = Convert.ToString(reader.GetValue(0), CultureInfo.InvariantCulture);
                        row["material_name"] = ReadString(reader, 1);
                        row["material_unit"] = ReadString(reader, 2);
                        row["doc_no"] = ReadString(reader, 3);
                        row["is_main_material"] = ReadString(reader, 4) == "1" ? "1" : "0";
                        row["transport_category"] = ReadString(reader, 5);
                        row["base_price"] = Convert.ToString(reader.GetDouble(6), CultureInfo.InvariantCulture);
                        row["current_price"] = Convert.ToString(reader.GetDouble(7), CultureInfo.InvariantCulture);
                        row["search_text"] = TextMatcher.Normalize(String.Join(" ", new string[] { row["material_code"], row["material_name"], row["material_unit"], row["doc_no"], row["transport_category"] }));
                        writer.WriteLine(LearningStore.ToJson(row));
                    }
                }
            }

            ReplaceFile(temp, path);
        }

        private static string ReadString(System.Data.SqlClient.SqlDataReader reader, int index)
        {
            return reader.IsDBNull(index) ? "" : Convert.ToString(reader.GetValue(index), CultureInfo.CurrentCulture).Trim();
        }

        private static bool IsTruthy(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return false;
            }

            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture) != 0;
            }
            catch
            {
                return String.Equals(Convert.ToString(value, CultureInfo.InvariantCulture), "true", StringComparison.OrdinalIgnoreCase);
            }
        }

        private static void ReplaceFile(string temp, string path)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
            File.Move(temp, path);
        }

        private sealed class ScoredQuota
        {
            public IndexQuota Quota;
            public int Score;
        }

    }

    internal sealed class IndexQuota
    {
        public string QuotaCode;
        public string QuotaName;
        public string QuotaUnit;
        public string BookCode;
        public string BookCategory;
        public string Specialty;
        public string SectionNo;
        public string SectionName;
        public string WorkContent;
        public string SearchText;
        public int SortOrder;

        public RecommendationRow ToRecommendation(ExcelQuantityItem item, int score)
        {
            RecommendationRow row = new RecommendationRow();
            row.Item = item;
            row.QuotaCode = QuotaCode;
            row.QuotaName = QuotaName;
            row.QuotaUnit = QuotaUnit;
            row.ConvertedValueText = RecommendDialog.ConvertQuantityForIndex(item.ValueText, item.Unit, QuotaUnit);
            row.Score = score;
            row.Reason = "\u5168\u91cf\u5b9a\u989d\u7d22\u5f15\u5173\u952e\u8bcd\u5339\u914d";
            row.Source = "index";
            row.TargetKind = "quota";
            return row;
        }
    }

    internal sealed class IndexMaterial
    {
        public string MaterialCode;
        public string MaterialName;
        public string MaterialUnit;
        public string DocNo;
        public bool IsMainMaterial;
        public string TransportCategory;
        public string SearchText;

        public RecommendationRow ToRecommendation(ExcelQuantityItem item, int score)
        {
            RecommendationRow row = new RecommendationRow();
            row.Item = item;
            row.QuotaCode = MaterialCode;
            row.QuotaName = MaterialName;
            row.QuotaUnit = MaterialUnit;
            row.ConvertedValueText = RecommendDialog.ConvertQuantityForIndex(item.ValueText, item.Unit, MaterialUnit);
            row.Score = score;
            row.Reason = IsMainMaterial ? "\u4e3b\u8981\u6750\u6599\u7d22\u5f15\u5173\u952e\u8bcd\u5339\u914d" : "\u6750\u6599\u7d22\u5f15\u5173\u952e\u8bcd\u5339\u914d";
            row.Source = "index";
            row.TargetKind = "material";
            return row;
        }
    }

    internal sealed class MappingStore
    {
        private const int MaxSamplesPerBox = 30;
        private readonly string path;
        private readonly List<MappingBox> boxes = new List<MappingBox>();

        private MappingStore(string filePath)
        {
            path = filePath;
        }

        public static MappingStore Load(List<LearningRecord> records)
        {
            string filePath = Path.Combine(LearningStore.FindDataDir(), "mapping-boxes.jsonl");
            MappingStore store = new MappingStore(filePath);
            store.LoadFile();
            if (store.boxes.Count == 0)
            {
                store.ImportCorrections(records);
                if (store.boxes.Count > 0)
                {
                    store.Save();
                }
            }
            return store;
        }

        public List<RecommendationRow> Find(ExcelQuantityItem item, string categoryFilter, SearchIndexStore searchIndex)
        {
            ScoredBox best = null;
            foreach (MappingBox box in boxes)
            {
                List<MappingTarget> allowedTargets = FilterTargetsByCategory(box.Targets, categoryFilter, searchIndex);
                if (allowedTargets.Count == 0)
                {
                    continue;
                }

                int score = box.Score(item);
                if (score >= 70 && (best == null || score > best.Score))
                {
                    best = new ScoredBox { Box = box, Score = score, AllowedTargets = allowedTargets };
                }
            }

            if (best == null)
            {
                return new List<RecommendationRow>();
            }

            return best.AllowedTargets
                .OrderBy(t => TargetSortRank(t.TargetKind, t.Code))
                .ThenBy(t => t.Code ?? "", StringComparer.OrdinalIgnoreCase)
                .Select(t => t.ToRecommendation(item, best.Score, best.Box.BoxId))
                .ToList();
        }

        private static List<MappingTarget> FilterTargetsByCategory(List<MappingTarget> targets, string categoryFilter, SearchIndexStore searchIndex)
        {
            List<MappingTarget> quotaTargets = targets
                .Where(t => String.Equals(String.IsNullOrWhiteSpace(t.TargetKind) ? QuotaEntry.GuessKind(t.Code) : t.TargetKind, "quota", StringComparison.OrdinalIgnoreCase))
                .ToList();
            List<MappingTarget> materialTargets = targets
                .Where(t => !String.Equals(String.IsNullOrWhiteSpace(t.TargetKind) ? QuotaEntry.GuessKind(t.Code) : t.TargetKind, "quota", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (quotaTargets.Count == 0)
            {
                return materialTargets;
            }

            List<MappingTarget> allowedQuotaTargets = quotaTargets
                .Where(t => searchIndex != null && searchIndex.IsMappingTargetAllowed(t.TargetKind, t.Code, categoryFilter))
                .ToList();
            if (allowedQuotaTargets.Count == 0)
            {
                return new List<MappingTarget>();
            }

            allowedQuotaTargets.AddRange(materialTargets);
            return allowedQuotaTargets;
        }

        public void Accept(List<RecommendationRow> rows)
        {
            bool changed = false;
            foreach (IGrouping<string, RecommendationRow> group in rows
                .Where(r => r != null && r.Item != null && !String.IsNullOrWhiteSpace(r.QuotaCode))
                .GroupBy(r => LearningStore.BuildQuantitySignature(r.Item), StringComparer.OrdinalIgnoreCase))
            {
                RecommendationRow first = group.First();
                MappingBox box = null;
                string boxId = group.Select(r => r.BoxId).FirstOrDefault(id => !String.IsNullOrWhiteSpace(id));
                if (!String.IsNullOrWhiteSpace(boxId))
                {
                    box = boxes.FirstOrDefault(b => String.Equals(b.BoxId, boxId, StringComparison.OrdinalIgnoreCase));
                }

                if (box == null)
                {
                    box = FindOrCreateBox(group.Select(row => new QuotaEntry
                    {
                        TargetKind = String.IsNullOrWhiteSpace(row.TargetKind) ? QuotaEntry.GuessKind(row.QuotaCode) : row.TargetKind,
                        QuotaCode = row.QuotaCode,
                        QuotaName = row.QuotaName,
                        QuotaUnit = row.QuotaUnit
                    }).ToList());
                }

                MappingSample sample = box.FindOrCreateSample(first.Item.Name, first.Item.Unit);
                sample.Weight += 5;
                sample.AcceptedCount += 1;
                sample.LastUsedAt = Now();
                box.TrimSamples(MaxSamplesPerBox);
                changed = true;
            }

            if (changed)
            {
                Save();
            }
        }

        public void Correct(ExcelQuantityItem item, RecommendationRow oldRecommendation, List<QuotaEntry> selectedTargets)
        {
            if (item == null || selectedTargets == null || selectedTargets.Count == 0)
            {
                return;
            }

            Penalize(item, oldRecommendation, selectedTargets);
            MappingBox box = FindOrCreateBox(selectedTargets);
            MappingSample sample = box.FindOrCreateSample(item.Name, item.Unit);
            sample.Weight += 20;
            sample.CorrectedCount += 1;
            sample.LastUsedAt = Now();
            box.TrimSamples(MaxSamplesPerBox);
            Save();
        }

        private void Penalize(ExcelQuantityItem item, RecommendationRow oldRecommendation, List<QuotaEntry> selectedTargets)
        {
            if (oldRecommendation == null || String.IsNullOrWhiteSpace(oldRecommendation.QuotaCode))
            {
                return;
            }

            string oldKey = (String.IsNullOrWhiteSpace(oldRecommendation.TargetKind) ? QuotaEntry.GuessKind(oldRecommendation.QuotaCode) : oldRecommendation.TargetKind) + ":" + oldRecommendation.QuotaCode.ToUpperInvariant();
            if (selectedTargets.Any(t => String.Equals(t.TargetKey, oldKey, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            foreach (MappingBox box in boxes)
            {
                if (!box.Targets.Any(t => String.Equals(t.TargetKey, oldKey, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                MappingSample sample = box.FindSample(item.Name, item.Unit);
                if (sample != null)
                {
                    sample.Weight = Math.Max(0, sample.Weight - 10);
                    sample.RejectedCount += 1;
                    sample.LastUsedAt = Now();
                }
            }
        }

        private MappingBox FindOrCreateBox(List<QuotaEntry> targets)
        {
            List<MappingTarget> normalized = targets
                .Where(t => !String.IsNullOrWhiteSpace(t.QuotaCode))
                .Select(t => new MappingTarget
                {
                    TargetKind = String.IsNullOrWhiteSpace(t.TargetKind) ? QuotaEntry.GuessKind(t.QuotaCode) : t.TargetKind,
                    Code = t.QuotaCode.Trim(),
                    Name = t.QuotaName ?? "",
                    Unit = t.QuotaUnit ?? ""
                })
                .GroupBy(t => t.TargetKey, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .OrderBy(t => TargetSortRank(t.TargetKind, t.Code))
                .ThenBy(t => t.Code ?? "", StringComparer.OrdinalIgnoreCase)
                .ToList();

            string boxId = BuildBoxId(normalized);
            MappingBox box = boxes.FirstOrDefault(b => String.Equals(b.BoxId, boxId, StringComparison.OrdinalIgnoreCase));
            if (box == null)
            {
                box = new MappingBox { BoxId = boxId };
                box.Targets.AddRange(normalized);
                boxes.Add(box);
            }
            else
            {
                foreach (MappingTarget target in normalized)
                {
                    if (!box.Targets.Any(t => String.Equals(t.TargetKey, target.TargetKey, StringComparison.OrdinalIgnoreCase)))
                    {
                        box.Targets.Add(target);
                    }
                }
            }

            return box;
        }

        private void ImportCorrections(List<LearningRecord> records)
        {
            foreach (IGrouping<string, LearningRecord> group in records
                .Where(r => r.IsCorrection && !String.IsNullOrWhiteSpace(r.QuotaCode))
                .GroupBy(r => r.QuantitySignature, StringComparer.OrdinalIgnoreCase))
            {
                List<QuotaEntry> targets = group.Select(r => new QuotaEntry
                {
                    TargetKind = QuotaEntry.GuessKind(r.QuotaCode),
                    QuotaCode = r.QuotaCode,
                    QuotaName = r.QuotaName,
                    QuotaUnit = r.QuotaUnit
                }).ToList();
                MappingBox box = FindOrCreateBox(targets);
                LearningRecord first = group.First();
                MappingSample sample = box.FindOrCreateSample(first.QuantityName, first.QuantityUnit);
                sample.Weight = Math.Max(sample.Weight, 30);
                sample.CorrectedCount += 1;
                sample.LastUsedAt = Now();
            }
        }

        private void LoadFile()
        {
            if (!File.Exists(path))
            {
                return;
            }

            Dictionary<string, MappingBox> byId = new Dictionary<string, MappingBox>(StringComparer.OrdinalIgnoreCase);
            foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
            {
                if (String.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                string boxId = LearningStore.Get(values, "box_id");
                if (String.IsNullOrWhiteSpace(boxId))
                {
                    continue;
                }

                MappingBox box;
                if (!byId.TryGetValue(boxId, out box))
                {
                    box = new MappingBox { BoxId = boxId };
                    byId[boxId] = box;
                    boxes.Add(box);
                }

                MappingTarget target = new MappingTarget
                {
                    TargetKind = LearningStore.Get(values, "target_kind"),
                    Code = LearningStore.Get(values, "target_code"),
                    Name = LearningStore.Get(values, "target_name"),
                    Unit = LearningStore.Get(values, "target_unit")
                };
                if (!String.IsNullOrWhiteSpace(target.Code) && !box.Targets.Any(t => String.Equals(t.TargetKey, target.TargetKey, StringComparison.OrdinalIgnoreCase)))
                {
                    box.Targets.Add(target);
                }

                MappingSample sample = new MappingSample();
                sample.QuantityName = LearningStore.Get(values, "quantity_name");
                sample.QuantityUnit = LearningStore.Get(values, "quantity_unit");
                sample.Weight = ParseInt(LearningStore.Get(values, "weight"), 0);
                sample.AcceptedCount = ParseInt(LearningStore.Get(values, "accepted_count"), 0);
                sample.CorrectedCount = ParseInt(LearningStore.Get(values, "corrected_count"), 0);
                sample.RejectedCount = ParseInt(LearningStore.Get(values, "rejected_count"), 0);
                sample.LastUsedAt = LearningStore.Get(values, "last_used_at");
                if (!String.IsNullOrWhiteSpace(sample.QuantityName) && box.FindSample(sample.QuantityName, sample.QuantityUnit) == null)
                {
                    box.Samples.Add(sample);
                }
            }
        }

        private void Save()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            string temp = path + ".tmp";
            using (StreamWriter writer = new StreamWriter(temp, false, Encoding.UTF8))
            {
                foreach (MappingBox box in boxes)
                {
                    box.TrimSamples(MaxSamplesPerBox);
                    foreach (MappingTarget target in box.Targets
                        .OrderBy(t => TargetSortRank(t.TargetKind, t.Code))
                        .ThenBy(t => t.Code ?? "", StringComparer.OrdinalIgnoreCase))
                    {
                        foreach (MappingSample sample in box.Samples)
                        {
                            Dictionary<string, string> row = new Dictionary<string, string>();
                            row["record_type"] = "mapping_box";
                            row["box_id"] = box.BoxId;
                            row["target_kind"] = String.IsNullOrWhiteSpace(target.TargetKind) ? QuotaEntry.GuessKind(target.Code) : target.TargetKind;
                            row["target_code"] = target.Code;
                            row["target_name"] = target.Name;
                            row["target_unit"] = target.Unit;
                            row["quantity_name"] = sample.QuantityName;
                            row["quantity_unit"] = sample.QuantityUnit;
                            row["weight"] = sample.Weight.ToString(CultureInfo.InvariantCulture);
                            row["accepted_count"] = sample.AcceptedCount.ToString(CultureInfo.InvariantCulture);
                            row["corrected_count"] = sample.CorrectedCount.ToString(CultureInfo.InvariantCulture);
                            row["rejected_count"] = sample.RejectedCount.ToString(CultureInfo.InvariantCulture);
                            row["last_used_at"] = sample.LastUsedAt;
                            writer.WriteLine(LearningStore.ToJson(row));
                        }
                    }
                }
            }

            if (File.Exists(path))
            {
                File.Delete(path);
            }
            File.Move(temp, path);
        }

        private static string BuildBoxId(List<MappingTarget> targets)
        {
            string raw = String.Join("|", targets
                .OrderBy(t => t.TargetKey, StringComparer.OrdinalIgnoreCase)
                .Select(t => t.TargetKey)
                .ToArray());
            return "box-" + Math.Abs(raw.ToLowerInvariant().GetHashCode()).ToString(CultureInfo.InvariantCulture);
        }

        private static int TargetSortRank(string targetKind, string code)
        {
            string kind = String.IsNullOrWhiteSpace(targetKind) ? QuotaEntry.GuessKind(code) : targetKind;
            return String.Equals(kind, "quota", StringComparison.OrdinalIgnoreCase) ? 0 : 1;
        }

        private static int ParseInt(string text, int fallback)
        {
            int value;
            return Int32.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value) ? value : fallback;
        }

        private static string Now()
        {
            return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        private sealed class ScoredBox
        {
            public MappingBox Box;
            public int Score;
            public List<MappingTarget> AllowedTargets;
        }
    }

    internal sealed class MappingBox
    {
        public string BoxId;
        public readonly List<MappingTarget> Targets = new List<MappingTarget>();
        public readonly List<MappingSample> Samples = new List<MappingSample>();

        public int Score(ExcelQuantityItem item)
        {
            int best = 0;
            foreach (MappingSample sample in Samples)
            {
                if (!TextMatcher.HasStrongNamePairMatch(item.Name, sample.QuantityName))
                {
                    continue;
                }

                if (TextMatcher.IsSteelOnlyAgainstConcrete(item.Name, sample.QuantityName))
                {
                    continue;
                }

                int similarity = TextMatcher.NamePairScore(item.Name, sample.QuantityName);
                if (similarity <= 0)
                {
                    continue;
                }

                int score = similarity + Math.Min(40, sample.Weight);
                if (RecommendDialog.UnitCompatibleForIndex(sample.QuantityUnit, item.Unit))
                {
                    score += 12;
                }
                best = Math.Max(best, score);
            }

            return best;
        }

        public MappingSample FindOrCreateSample(string name, string unit)
        {
            MappingSample sample = FindSample(name, unit);
            if (sample != null)
            {
                return sample;
            }

            sample = new MappingSample { QuantityName = name ?? "", QuantityUnit = unit ?? "", Weight = 10, LastUsedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) };
            Samples.Add(sample);
            return sample;
        }

        public MappingSample FindSample(string name, string unit)
        {
            string signature = LearningStore.BuildQuantitySignature(name, unit);
            return Samples.FirstOrDefault(s => String.Equals(LearningStore.BuildQuantitySignature(s.QuantityName, s.QuantityUnit), signature, StringComparison.OrdinalIgnoreCase));
        }

        public void TrimSamples(int maxSamples)
        {
            while (Samples.Count > maxSamples)
            {
                MappingSample remove = Samples
                    .OrderBy(s => s.Weight)
                    .ThenBy(s => s.LastUsedAt ?? "")
                    .First();
                Samples.Remove(remove);
            }
        }
    }

    internal sealed class MappingTarget
    {
        public string TargetKind;
        public string Code;
        public string Name;
        public string Unit;

        public string TargetKey
        {
            get { return (String.IsNullOrWhiteSpace(TargetKind) ? QuotaEntry.GuessKind(Code) : TargetKind) + ":" + (Code ?? "").Trim().ToUpperInvariant(); }
        }

        public RecommendationRow ToRecommendation(ExcelQuantityItem item, int score, string boxId)
        {
            RecommendationRow row = new RecommendationRow();
            row.Item = item;
            row.QuotaCode = Code;
            row.QuotaName = Name;
            row.QuotaUnit = Unit;
            row.ConvertedValueText = RecommendDialog.ConvertQuantityForIndex(item.ValueText, item.Unit, Unit);
            row.Score = score;
            row.Reason = "\u5b9a\u989d\u5bf9\u5e94\u6846\u6743\u91cd\u5339\u914d";
            row.Source = "mapping";
            row.BoxId = boxId;
            row.TargetKind = String.IsNullOrWhiteSpace(TargetKind) ? QuotaEntry.GuessKind(Code) : TargetKind;
            return row;
        }
    }

    internal sealed class MappingSample
    {
        public string QuantityName;
        public string QuantityUnit;
        public int Weight;
        public int AcceptedCount;
        public int CorrectedCount;
        public int RejectedCount;
        public string LastUsedAt;
    }

    internal static class TextMatcher
    {
        public static string Normalize(string text)
        {
            return (text ?? "")
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("（", "(")
                .Replace("）", ")")
                .Replace("，", ",")
                .Replace("、", " ")
                .Trim()
                .ToLowerInvariant();
        }

        public static int Similarity(string left, string right)
        {
            return Math.Min(100, NamePairScore(left, right));
        }

        public static int NamePairScore(string left, string right)
        {
            string l = Normalize(left).Replace(" ", "");
            string r = Normalize(right).Replace(" ", "");
            if (String.IsNullOrEmpty(l) || String.IsNullOrEmpty(r))
            {
                return 0;
            }

            if (l == r)
            {
                return 120;
            }
            if (l.Contains(r) || r.Contains(l))
            {
                return 95;
            }

            List<string> leftTokens = Keywords(l).Distinct().ToList();
            if (leftTokens.Count == 0)
            {
                return 0;
            }

            int score = 0;
            int possible = 0;
            int hits = 0;
            foreach (string token in leftTokens)
            {
                int tokenScore = PairTokenScore(token);
                possible += tokenScore;
                if (r.Contains(token))
                {
                    hits++;
                    score += tokenScore;
                }
            }

            if (hits == 0)
            {
                return 0;
            }

            score += (int)Math.Round(18.0 * hits / leftTokens.Count);
            return possible > 0 && score > 115 ? 115 : score;
        }

        public static bool HasStrongNamePairMatch(string left, string right)
        {
            string l = Normalize(left).Replace(" ", "");
            string r = Normalize(right).Replace(" ", "");
            if (String.IsNullOrEmpty(l) || String.IsNullOrEmpty(r))
            {
                return false;
            }

            if (l == r || l.Contains(r) || r.Contains(l))
            {
                return true;
            }

            bool steelConcreteBlocked = IsSteelOnlyAgainstConcrete(l, r);
            foreach (string token in Keywords(l).Distinct())
            {
                if (token.Length < 2 || !r.Contains(token))
                {
                    continue;
                }

                if (steelConcreteBlocked && String.Equals(token, "钢筋", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                return true;
            }

            return false;
        }

        public static bool IsSteelOnlyAgainstConcrete(string left, string right)
        {
            string l = Normalize(left);
            string r = Normalize(right);
            bool leftConcrete = IsConcreteQuantityName(l);
            bool rightConcrete = IsConcreteQuantityName(r);
            return (IsSteelQuantityName(l) && !leftConcrete && rightConcrete) ||
                (IsSteelQuantityName(r) && !rightConcrete && leftConcrete);
        }

        public static bool IsSteelQuantityName(string text)
        {
            string value = Normalize(text);
            return value.Contains("钢筋") ||
                value.Contains("hpb") ||
                value.Contains("hrb") ||
                value.Contains("圆钢") ||
                value.Contains("螺纹");
        }

        public static bool IsConcreteQuantityName(string text)
        {
            string value = Normalize(text);
            return value.Contains("混凝土") ||
                value.Contains("砼") ||
                value.Contains("商品混凝土");
        }

        private static int PairTokenScore(string token)
        {
            if (String.IsNullOrWhiteSpace(token))
            {
                return 0;
            }

            if (HasAsciiOrDigit(token))
            {
                return token.Length >= 3 ? 28 : 8;
            }

            if (token.Length == 1)
            {
                return 4;
            }
            if (token.Length == 2)
            {
                return 24;
            }
            if (token.Length == 3)
            {
                return 38;
            }
            return 50;
        }

        public static IEnumerable<string> Keywords(string text)
        {
            string normalized = Normalize(text);
            foreach (string part in normalized.Split(new char[] { ' ', '/', ',', ';', '\t', '(', ')', '[', ']', '+', '-', '*', '=' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string token = part.Trim();
                if (token.Length < 2 || IsNumberLike(token))
                {
                    continue;
                }

                yield return token;
                foreach (string segment in SplitAlphaNumericAndChinese(token))
                {
                    if (segment.Length >= 2 && !IsNumberLike(segment))
                    {
                        yield return segment;
                    }
                }
                if (ContainsChinese(token))
                {
                    for (int i = 0; i + 2 <= token.Length; i++)
                    {
                        string gram = token.Substring(i, 2);
                        if (!IsNumberLike(gram))
                        {
                            yield return gram;
                        }
                    }
                    for (int i = 0; i + 3 <= token.Length; i++)
                    {
                        string gram = token.Substring(i, 3);
                        if (!IsNumberLike(gram))
                        {
                            yield return gram;
                        }
                    }
                    for (int i = 0; i < token.Length; i++)
                    {
                        string gram = token.Substring(i, 1);
                        if (IsPureChinese(gram))
                        {
                            yield return gram;
                        }
                    }
                }
            }
        }

        public static bool HasAsciiOrDigit(string text)
        {
            foreach (char ch in text ?? "")
            {
                if ((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || Char.IsDigit(ch))
                {
                    return true;
                }
            }
            return false;
        }

        public static bool IsPureChinese(string text)
        {
            bool hasChinese = false;
            foreach (char ch in text ?? "")
            {
                if (ch >= 0x4e00 && ch <= 0x9fff)
                {
                    hasChinese = true;
                    continue;
                }

                return false;
            }
            return hasChinese;
        }

        private static IEnumerable<string> SplitAlphaNumericAndChinese(string token)
        {
            StringBuilder builder = new StringBuilder();
            int lastKind = 0;
            foreach (char ch in token ?? "")
            {
                int kind = Char.IsLetterOrDigit(ch) && !(ch >= 0x4e00 && ch <= 0x9fff) ? 1 : ((ch >= 0x4e00 && ch <= 0x9fff) ? 2 : 0);
                if (kind == 0)
                {
                    if (builder.Length > 0)
                    {
                        yield return builder.ToString();
                        builder.Length = 0;
                    }
                    lastKind = 0;
                    continue;
                }

                if (builder.Length > 0 && kind != lastKind)
                {
                    yield return builder.ToString();
                    builder.Length = 0;
                }

                builder.Append(ch);
                lastKind = kind;
            }

            if (builder.Length > 0)
            {
                yield return builder.ToString();
            }
        }

        private static bool ContainsChinese(string text)
        {
            foreach (char ch in text ?? "")
            {
                if (ch >= 0x4e00 && ch <= 0x9fff)
                {
                    return true;
                }
            }
            return false;
        }

        private static bool IsNumberLike(string text)
        {
            decimal value;
            return Decimal.TryParse((text ?? "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        public static bool IsNumberLikeToken(string text)
        {
            return IsNumberLike(text);
        }
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
                record.IsCorrection = String.Equals(Get(item, "record_type"), "correction", StringComparison.OrdinalIgnoreCase)
                    || String.Equals(Get(item, "user_action"), "correction", StringComparison.OrdinalIgnoreCase);
                record.ProjectName = Get(item, "project_name");
                record.BudgetFile = Get(item, "budget_file");
                record.BudgetGroup = Get(item, "budget_group");
                record.QuotaCode = Get(item, "quota_code");
                record.QuotaName = Get(item, "quota_name");
                record.QuotaUnit = Get(item, "quota_unit");
                record.QuantitySection = Get(item, "quantity_section");
                record.QuantityName = Get(item, "quantity_name");
                record.QuantityUnit = Get(item, "quantity_unit");
                record.QuantitySignature = Get(item, "quantity_signature");
                if (String.IsNullOrWhiteSpace(record.QuantitySignature))
                {
                    record.QuantitySignature = BuildQuantitySignature(record.QuantityName, record.QuantityUnit);
                }
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

        public static void BackupLearningFileIfNeeded()
        {
            try
            {
                string path = FindLearningPath();
                if (String.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    return;
                }

                string directory = Path.GetDirectoryName(path);
                if (Directory.GetFiles(directory, "learning.jsonl.*.bak").Length > 0)
                {
                    return;
                }

                string marker = Path.Combine(directory, "learning.jsonl." + DateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture) + ".bak");
                File.Copy(path, marker, false);
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("Backup learning.jsonl failed: " + ex.Message);
            }
        }

        public static void ReplaceCorrections(ExcelQuantityItem item, List<QuotaEntry> quotas)
        {
            string signature = BuildQuantitySignature(item);
            List<string> paths = FindLearningPaths();
            if (paths.Count == 0)
            {
                string baseDir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location);
                paths.Add(Path.Combine(baseDir, "RecoQuotaData", "learning.jsonl"));
            }

            foreach (string path in paths)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                List<string> existing = File.Exists(path)
                    ? File.ReadAllLines(path, Encoding.UTF8).ToList()
                    : new List<string>();

                List<string> kept = new List<string>();
                foreach (string line in existing)
                {
                    if (String.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    Dictionary<string, string> values = ParseFlatJson(line);
                    bool isCorrection = String.Equals(Get(values, "record_type"), "correction", StringComparison.OrdinalIgnoreCase)
                        || String.Equals(Get(values, "user_action"), "correction", StringComparison.OrdinalIgnoreCase);
                    string lineSignature = Get(values, "quantity_signature");
                    if (String.IsNullOrWhiteSpace(lineSignature))
                    {
                        lineSignature = BuildQuantitySignature(Get(values, "quantity_name"), Get(values, "quantity_unit"));
                    }

                    if (isCorrection && String.Equals(lineSignature, signature, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    kept.Add(line);
                }

                foreach (QuotaEntry quota in quotas)
                {
                    Dictionary<string, string> record = new Dictionary<string, string>();
                    record["record_type"] = "correction";
                    record["user_action"] = "correction";
                    record["quantity_signature"] = signature;
                    record["quantity_name"] = item.Name;
                    record["quantity_unit"] = item.Unit;
                    record["quantity_section"] = item.SectionName;
                    record["quota_code"] = quota.QuotaCode;
                    record["quota_name"] = quota.QuotaName;
                    record["quota_unit"] = quota.QuotaUnit;
                    record["match_score"] = "1000";
                    record["match_reason"] = "\u4eba\u5de5\u6276\u6b63";
                    record["updated_at"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                    kept.Add(ToJson(record));
                }

                File.WriteAllLines(path, kept.ToArray(), Encoding.UTF8);
            }
        }

        public static string BuildQuantitySignature(ExcelQuantityItem item)
        {
            return BuildQuantitySignature(item == null ? "" : item.Name, item == null ? "" : item.Unit);
        }

        public static string BuildQuantitySignature(string name, string unit)
        {
            return NormalizeForSignature(name) + "|" + NormalizeForSignature(unit);
        }

        private static string FindLearningPath()
        {
            List<string> paths = FindLearningPaths();
            return paths.Count == 0 ? "" : paths[0];
        }

        internal static string FindDataDir()
        {
            string baseDir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location);
            return Path.Combine(baseDir, "RecoQuotaData");
        }

        private static List<string> FindLearningPaths()
        {
            List<string> paths = new List<string>();
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
                    paths.Add(candidate);
                }
            }

            return paths;
        }

        internal static string ToJson(Dictionary<string, string> values)
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
                builder.Append('"').Append(EscapeJson(pair.Key)).Append('"').Append(':')
                    .Append('"').Append(EscapeJson(pair.Value)).Append('"');
            }

            builder.Append('}');
            return builder.ToString();
        }

        private static string EscapeJson(string value)
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

        private static string NormalizeForSignature(string text)
        {
            return (text ?? "").Replace("\r", "").Replace("\n", "").Replace(" ", "").Trim().ToLowerInvariant();
        }

        internal static string Get(Dictionary<string, string> values, string key)
        {
            string value;
            return values.TryGetValue(key, out value) ? value : "";
        }

        internal static Dictionary<string, string> ParseFlatJson(string line)
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
