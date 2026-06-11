using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace RecoQuotaRecommend
{
    public sealed class QuotaRecommendPanel : Form
    {
        private static readonly HashSet<Form> InstalledForms = new HashSet<Form>();
        private static readonly Dictionary<ContextMenuStrip, MenuInfo> MenuInfos = new Dictionary<ContextMenuStrip, MenuInfo>();
        private static readonly Dictionary<Form, RecommendDialog> RecommendDialogs = new Dictionary<Form, RecommendDialog>();
        private static Image recommendMenuIcon;
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
                ApplyRecommendMenuIcon(item);
                return;
            }

            int insertIndex = Math.Min(2, menu.Items.Count);
            item = new ToolStripMenuItem("\u63a8\u8350\u5b9a\u989d");
            item.Visible = true;
            item.Available = true;
            item.Enabled = true;
            ApplyRecommendMenuIcon(item);
            item.Click += delegate { ShowRecommendDialog(mainForm); };
            menu.Items.Insert(insertIndex, item);
        }

        private static void ApplyRecommendMenuIcon(ToolStripMenuItem item)
        {
            if (item == null)
            {
                return;
            }

            Image icon = LoadRecommendMenuIcon();
            if (icon != null)
            {
                item.Image = icon;
                item.ImageScaling = ToolStripItemImageScaling.SizeToFit;
            }
        }

        private static Image LoadRecommendMenuIcon()
        {
            if (recommendMenuIcon != null)
            {
                return recommendMenuIcon;
            }

            try
            {
                string dir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location);
                string path = Path.Combine(dir, "RecoExpandPanelIcons", "recommend_quota.png");
                if (File.Exists(path))
                {
                    using (Image image = Image.FromFile(path))
                    {
                        recommendMenuIcon = new Bitmap(image);
                    }

                    return recommendMenuIcon;
                }
            }
            catch (Exception ex)
            {
                Log("Load recommend menu icon failed: " + ex.Message);
            }

            recommendMenuIcon = DrawRecommendMenuIcon();
            return recommendMenuIcon;
        }

        private static Image DrawRecommendMenuIcon()
        {
            Bitmap bitmap = new Bitmap(24, 24);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                graphics.Clear(Color.Transparent);
                using (SolidBrush paper = new SolidBrush(Color.FromArgb(244, 248, 251)))
                using (Pen border = new Pen(Color.FromArgb(104, 126, 144)))
                using (Pen gridPen = new Pen(Color.FromArgb(168, 184, 196)))
                using (SolidBrush excel = new SolidBrush(Color.FromArgb(70, 139, 73)))
                using (Pen excelBorder = new Pen(Color.FromArgb(46, 96, 49)))
                using (Pen xPen = new Pen(Color.White, 1.6f))
                using (SolidBrush mark = new SolidBrush(Color.FromArgb(79, 144, 84)))
                using (Pen markPen = new Pen(Color.FromArgb(57, 102, 61), 1.4f))
                {
                    graphics.FillRectangle(paper, 6, 2, 14, 18);
                    graphics.DrawRectangle(border, 6, 2, 14, 18);
                    graphics.DrawLine(gridPen, 9, 7, 17, 7);
                    graphics.DrawLine(gridPen, 9, 11, 17, 11);
                    graphics.DrawLine(gridPen, 9, 15, 17, 15);
                    graphics.DrawLine(gridPen, 12, 5, 12, 18);
                    graphics.DrawLine(gridPen, 16, 5, 16, 18);
                    graphics.FillRectangle(excel, 1, 8, 9, 10);
                    graphics.DrawRectangle(excelBorder, 1, 8, 9, 10);
                    xPen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                    xPen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                    graphics.DrawLine(xPen, 3, 10, 8, 16);
                    graphics.DrawLine(xPen, 8, 10, 3, 16);
                    Point[] star = new Point[]
                    {
                        new Point(19, 3),
                        new Point(21, 7),
                        new Point(23, 8),
                        new Point(21, 10),
                        new Point(21, 13),
                        new Point(18, 11),
                        new Point(15, 12),
                        new Point(16, 9),
                        new Point(14, 6),
                        new Point(17, 7)
                    };
                    graphics.FillPolygon(mark, star);
                    graphics.DrawPolygon(markPen, star);
                }
            }

            return bitmap;
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
        // 与 SearchIndexStore.Search 的入选阈值保持一致：本地索引达到该分数即可直接展示，不必等 AI。
        private const int LocalIndexDisplayScore = 55;

        private readonly Form mainForm;
        private readonly DataGridView resultGrid;
        private readonly Label statusLabel;
        private readonly ComboBox quotaCategoryCombo;
        private readonly CheckBox aiNameCheckBox;
        private readonly List<LearningRecord> records;
        private readonly SearchIndexStore searchIndex;
        private readonly MappingStore mappingStore;
        private DeepSeekSettings deepSeekSettings;
        private readonly List<RecommendationRow> recommendations = new List<RecommendationRow>();
        private ExcelSelection currentSelection;
        private int aiRequestVersion;

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
            deepSeekSettings = DeepSeekSettings.Load();

            aiNameCheckBox = new CheckBox();
            aiNameCheckBox.Text = "AI\u8bc6\u522b\u5de5\u7a0b\u91cf";
            aiNameCheckBox.Left = 12;
            aiNameCheckBox.Top = 13;
            aiNameCheckBox.Width = 150;
            aiNameCheckBox.Checked = false;
            aiNameCheckBox.CheckedChanged += delegate
            {
                if (currentSelection != null)
                {
                    FillRecommendations(currentSelection);
                }
            };

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

            Button aiSettingsButton = new Button();
            aiSettingsButton.Text = "AI\u8bbe\u7f6e";
            aiSettingsButton.Left = 632;
            aiSettingsButton.Top = 10;
            aiSettingsButton.Width = 78;
            aiSettingsButton.Click += delegate { ShowDeepSeekSettings(); };

            Label categoryLabel = new Label();
            categoryLabel.Text = "\u5b9a\u989d\u7c7b\u578b";
            categoryLabel.Left = 720;
            categoryLabel.Top = 15;
            categoryLabel.Width = 58;

            quotaCategoryCombo = new ComboBox();
            quotaCategoryCombo.Left = 778;
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

            Controls.Add(aiNameCheckBox);
            Controls.Add(clipboardButton);
            Controls.Add(selectAllButton);
            Controls.Add(clearButton);
            Controls.Add(pasteButton);
            Controls.Add(aiSettingsButton);
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
            check.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            check.Width = 48;
            resultGrid.Columns.Add(check);
            DataGridViewButtonColumn correct = new DataGridViewButtonColumn();
            correct.Name = "Correct";
            correct.HeaderText = "\u6276\u6b63";
            correct.Text = "\u6276\u6b63";
            correct.UseColumnTextForButtonValue = false;
            correct.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            correct.Width = 48;
            resultGrid.Columns.Add(correct);
            resultGrid.Columns.Add("QuantityName", "\u5de5\u7a0b\u91cf\u540d\u79f0");
            resultGrid.Columns.Add("QuantityUnit", "\u5355\u4f4d");
            resultGrid.Columns.Add("QuantityValue", "Excel\u5de5\u7a0b\u91cf");
            resultGrid.Columns.Add("QuotaCode", "\u63a8\u8350\u5b9a\u989d");
            resultGrid.Columns.Add("QuotaName", "\u5b9a\u989d\u540d\u79f0");
            resultGrid.Columns.Add("QuotaUnit", "\u5b9a\u989d\u5355\u4f4d");
            resultGrid.Columns.Add("QuotaQuantity", "\u5b9a\u989d\u5de5\u7a0b\u91cf");
            resultGrid.Columns.Add("SourceStatus", "\u6765\u6e90");
            resultGrid.Columns["QuantityName"].FillWeight = 180;
            resultGrid.Columns["QuantityUnit"].FillWeight = 50;
            resultGrid.Columns["QuantityValue"].FillWeight = 80;
            resultGrid.Columns["QuotaCode"].FillWeight = 80;
            resultGrid.Columns["QuotaName"].FillWeight = 210;
            resultGrid.Columns["QuotaUnit"].FillWeight = 60;
            resultGrid.Columns["QuotaQuantity"].FillWeight = 85;
            resultGrid.Columns["SourceStatus"].FillWeight = 70;

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
            FillRecommendations(selection, aiNameCheckBox != null && aiNameCheckBox.Checked);
        }

        private void FillRecommendations(ExcelSelection selection, bool normalizeNames)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            currentSelection = selection;
            recommendations.Clear();
            resultGrid.Rows.Clear();
            if (normalizeNames)
            {
                NormalizeQuantityNamesWithDeepSeek(selection, true);
            }
            else
            {
                RestoreOriginalQuantityNames(selection);
            }
            int requestVersion = ++aiRequestVersion;
            string categoryFilter = SelectedQuotaCategory;
            RecommendationBatchStats stats = new RecommendationBatchStats();
            Dictionary<string, List<RecommendationRow>> batchCache = new Dictionary<string, List<RecommendationRow>>(StringComparer.OrdinalIgnoreCase);
            List<AiPendingRecommendation> aiPending = new List<AiPendingRecommendation>();

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
                            recommendation.QuotaCode,
                            recommendation.QuotaName,
                            recommendation.QuotaUnit,
                            recommendation.ConvertedValueText,
                            RecommendationStatusText(recommendation));
                        recommendation.GridRowIndex = gridRowIndex;
                        resultGrid.Rows[gridRowIndex].Cells["SourceStatus"].ToolTipText = recommendation.Reason ?? "";
                        if (!String.IsNullOrWhiteSpace(item.OriginalName) || !String.IsNullOrWhiteSpace(item.RawRowText) || !String.IsNullOrWhiteSpace(item.AiNameReason))
                        {
                            resultGrid.Rows[gridRowIndex].Cells["QuantityName"].ToolTipText =
                                "\u539f\u59cb\u540d\u79f0\uff1a" + (item.OriginalName ?? "") +
                                "\r\nAI\u7406\u7531\uff1a" + (item.AiNameReason ?? "") +
                                "\r\n\u539f\u59cb\u884c\uff1a" + (item.RawRowText ?? "");
                        }
                        if (isContinuation)
                        {
                            DataGridViewRow gridRow = resultGrid.Rows[gridRowIndex];
                            gridRow.Cells["Correct"] = new DataGridViewTextBoxCell();
                            gridRow.Cells["Correct"].Value = "";
                            gridRow.Cells["Correct"].ReadOnly = true;
                        }
                        else if (ShouldQueueDeepSeek(recommendation))
                        {
                            recommendation.AiRowId = "r" + aiPending.Count.ToString(CultureInfo.InvariantCulture);
                            recommendation.AiPending = true;
                            aiPending.Add(new AiPendingRecommendation
                            {
                                Row = recommendation,
                                GridRowIndex = gridRowIndex,
                                Request = new DeepSeekRequestRow
                                {
                                    RowId = recommendation.AiRowId,
                                    Item = item,
                                    Candidates = recommendation.AiCandidates,
                                    MappingCandidates = recommendation.AiMappingCandidates
                                }
                            });
                            resultGrid.Rows[gridRowIndex].Cells["SourceStatus"].Value = "\u0041\u0049\u8865\u63a8\u4e2d";
                            stats.AiQueued++;
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
                "\u5df2\u8bfb\u53d6 {0} \u884cExcel\u5de5\u7a0b\u91cf\uff0c\u5b9a\u989d\u7c7b\u578b\uff1a{1}\uff0c\u5bf9\u5e94\u6846\u547d\u4e2d {2} \u884c\uff0c\u7d22\u5f15\u68c0\u7d22 {3} \u884c\uff0c\u7a7a\u7ed3\u679c {4} \u884c\uff0c\u91cd\u590d\u590d\u7528 {5} \u884c\uff0cAI\u8865\u63a8 {6} \u884c\uff0c\u8017\u65f6 {7} ms\u3002",
                selection.Items.Count,
                categoryFilter,
                stats.MappingHits,
                stats.IndexSearches,
                stats.EmptyRows,
                stats.CacheHits,
                stats.AiQueued,
                stopwatch.ElapsedMilliseconds);

            StartDeepSeekRecommendations(aiPending, requestVersion);
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

        private static void RestoreOriginalQuantityNames(ExcelSelection selection)
        {
            if (selection == null)
            {
                return;
            }

            foreach (ExcelQuantityItem item in selection.Items)
            {
                if (item == null || String.IsNullOrWhiteSpace(item.OriginalName) || String.Equals(item.Name, item.OriginalName, StringComparison.Ordinal))
                {
                    continue;
                }

                item.Name = item.OriginalName;
                item.SectionName = item.OriginalName;
                item.ContextText = item.Name + " " + item.Unit + " " + item.ValueText + " " + item.RawRowText;
            }
        }

        private void NormalizeQuantityNamesWithDeepSeek(ExcelSelection selection, bool force)
        {
            if (selection == null || selection.Items.Count == 0 || !deepSeekSettings.CanNormalizeNames)
            {
                return;
            }

            List<DeepSeekNameRequestRow> rows = new List<DeepSeekNameRequestRow>();
            for (int i = 0; i < selection.Items.Count; i++)
            {
                ExcelQuantityItem item = selection.Items[i];
                if (item == null)
                {
                    continue;
                }

                if (String.IsNullOrWhiteSpace(item.OriginalName))
                {
                    item.OriginalName = item.Name;
                }
                if (item.SkipAiNameNormalization && !force)
                {
                    continue;
                }
                rows.Add(new DeepSeekNameRequestRow
                {
                    RowId = "n" + i.ToString(CultureInfo.InvariantCulture),
                    Item = item
                });
            }

            if (rows.Count == 0)
            {
                return;
            }

            statusLabel.Text = "DeepSeek\u6b63\u5728\u8bc6\u522b\u5de5\u7a0b\u91cf\u540d\u79f0...";
            statusLabel.Refresh();
            try
            {
                DeepSeekClient client = new DeepSeekClient(deepSeekSettings);
                int batchSize = Math.Max(1, deepSeekSettings.MaxRowsPerBatch);
                int changed = 0;
                for (int i = 0; i < rows.Count; i += batchSize)
                {
                    List<DeepSeekNameRequestRow> batch = rows.Skip(i).Take(batchSize).ToList();
                    Dictionary<string, DeepSeekNameResult> byRow = client.NormalizeQuantityNames(batch)
                        .Where(r => r != null && !String.IsNullOrWhiteSpace(r.RowId))
                        .GroupBy(r => r.RowId, StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
                    foreach (DeepSeekNameRequestRow request in batch)
                    {
                        DeepSeekNameResult result;
                        if (!byRow.TryGetValue(request.RowId, out result) || String.IsNullOrWhiteSpace(result.QuantityName) || result.Confidence < 50)
                        {
                            continue;
                        }

                        ExcelQuantityItem item = request.Item;
                        item.AiName = CleanAiQuantityName(result.QuantityName);
                        if (String.IsNullOrWhiteSpace(item.AiName))
                        {
                            continue;
                        }

                        item.AiNameConfidence = result.Confidence;
                        item.AiNameReason = result.Reason;
                        item.Name = item.AiName;
                        item.SectionName = item.AiName;
                        item.ContextText = item.AiName + " " + item.Unit + " " + item.ValueText + " " + item.RawRowText;
                        changed++;
                    }
                }

                if (changed > 0)
                {
                    QuotaRecommendPanel.Log("DeepSeek normalized quantity names. changed=" + changed.ToString(CultureInfo.InvariantCulture));
                }
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("DeepSeek quantity name normalization failed: " + ex.Message);
            }
        }

        private static string CleanAiQuantityName(string name)
        {
            string value = (name ?? "").Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim();
            return value.Length > 80 ? value.Substring(0, 80).Trim() : value;
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
                row.AiCandidates = original.AiCandidates;
                row.AiMappingCandidates = original.AiMappingCandidates;
                rows.Add(row);
            }
            return rows;
        }

        private List<RecommendationRow> BuildRecommendations(ExcelQuantityItem item, string categoryFilter, RecommendationBatchStats stats)
        {
            List<AiQuotaCandidate> aiCandidates = new List<AiQuotaCandidate>();
            List<RecommendationRow> mapped = mappingStore.Find(item, categoryFilter, searchIndex);
            if (mapped.Count > 0)
            {
                stats.MappingHits++;
                return mapped;
            }

            List<AiMappingCandidate> mappingCandidates = deepSeekSettings.CanDetectMapping
                ? mappingStore.BuildDeepSeekCandidates(item, categoryFilter, searchIndex, deepSeekSettings.MaxCandidatesPerRow)
                : new List<AiMappingCandidate>();
            stats.IndexSearches++;
            foreach (AiQuotaCandidate candidate in searchIndex.BuildDeepSeekCandidates(item, categoryFilter, deepSeekSettings.MaxCandidatesPerRow))
            {
                if (!aiCandidates.Any(c => c != null && c.Quota != null && candidate != null && candidate.Quota != null && String.Equals(c.Quota.QuotaCode, candidate.Quota.QuotaCode, StringComparison.OrdinalIgnoreCase)))
                {
                    aiCandidates.Add(candidate);
                }
            }

            // \u672c\u5730\u7d22\u5f15\u9ad8\u5206\u547d\u4e2d\u65f6\u76f4\u63a5\u5c55\u793a\uff08\u6765\u6e90=\u672c\u5730\u7d22\u5f15\uff09\uff0cDeepSeek \u4e0d\u53ef\u7528\u4e5f\u6709\u7ed3\u679c\uff1b
            // AI \u8fd4\u56de\u4e14\u7f6e\u4fe1\u5ea6\u4e0d\u4f4e\u4e8e\u672c\u5730\u5206\u65f6\u624d\u5728 ApplyDeepSeekResults \u4e2d\u8986\u76d6\u3002
            RecommendationRow row = null;
            AiQuotaCandidate topLocal = aiCandidates
                .Where(c => c != null && c.Quota != null)
                .OrderByDescending(c => c.LocalScore)
                .ThenBy(c => c.Quota.SortOrder)
                .FirstOrDefault();
            if (topLocal != null && topLocal.LocalScore >= LocalIndexDisplayScore)
            {
                row = topLocal.Quota.ToRecommendation(item, topLocal.LocalScore);
            }

            if (row == null)
            {
                row = new RecommendationRow();
                row.Item = item;
                row.ConvertedValueText = item.ValueText;
                row.Score = 0;
                row.Reason = deepSeekSettings.CanRecommendQuota
                    ? "AI\u63a8\u8350\u7b49\u5f85\u8fd4\u56de"
                    : "DeepSeek AI\u672a\u542f\u7528\uff0c\u8bf7\u5728AI\u8bbe\u7f6e\u4e2d\u914d\u7f6e";
                row.Source = "empty";
                row.TargetKind = "quota";
            }

            row.AiMappingCandidates = mappingCandidates;
            row.AiCandidates = aiCandidates
                .Where(c => c != null && c.Quota != null)
                .GroupBy(c => c.Quota.QuotaCode, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.OrderByDescending(c => c.LocalScore).First())
                .OrderByDescending(c => c.LocalScore)
                .Take(deepSeekSettings.MaxCandidatesPerRow)
                .ToList();
            if (String.Equals(row.Source, "empty", StringComparison.OrdinalIgnoreCase) &&
                row.AiCandidates.Count == 0 && row.AiMappingCandidates.Count == 0)
            {
                row.Reason = "AI\u65e0\u6709\u6548\u5019\u9009\uff0c\u8bf7\u4eba\u5de5\u6276\u6b63";
                stats.EmptyRows++;
            }

            return new List<RecommendationRow> { row };
        }

        private List<RecommendationRow> FindMappingWithDeepSeek(ExcelQuantityItem item, string categoryFilter)
        {
            if (item == null || !deepSeekSettings.CanDetectMapping)
            {
                return new List<RecommendationRow>();
            }

            List<AiMappingCandidate> candidates = mappingStore.BuildDeepSeekCandidates(item, categoryFilter, searchIndex, Math.Max(3, deepSeekSettings.MaxCandidatesPerRow));
            if (candidates.Count == 0)
            {
                return new List<RecommendationRow>();
            }

            try
            {
                DeepSeekClient client = new DeepSeekClient(deepSeekSettings);
                DeepSeekMappingSelection selection = client.SelectMappingBox(item, candidates);
                if (selection == null || String.IsNullOrWhiteSpace(selection.BoxId) || selection.Confidence < deepSeekSettings.DisplayConfidence)
                {
                    return new List<RecommendationRow>();
                }

                AiMappingCandidate candidate = candidates.FirstOrDefault(c => String.Equals(c.BoxId, selection.BoxId, StringComparison.OrdinalIgnoreCase));
                return candidate == null ? new List<RecommendationRow>() : candidate.ToRecommendations(item, selection.Confidence, selection.Reason);
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("DeepSeek mapping detection failed: " + ex.Message);
                return new List<RecommendationRow>();
            }
        }

        private bool ShouldQueueDeepSeek(RecommendationRow row)
        {
            return row != null &&
                (deepSeekSettings.CanRecommendQuota || deepSeekSettings.CanDetectMapping) &&
                ((row.AiCandidates != null && row.AiCandidates.Count > 0) ||
                    (row.AiMappingCandidates != null && row.AiMappingCandidates.Count > 0));
        }

        private static string RecommendationStatusText(RecommendationRow row)
        {
            if (row == null)
            {
                return "";
            }

            if (String.Equals(row.Source, "mapping", StringComparison.OrdinalIgnoreCase))
            {
                return "\u5bf9\u5e94\u6846";
            }
            if (String.Equals(row.Source, "deepseek", StringComparison.OrdinalIgnoreCase))
            {
                return "AI\u8865\u63a8";
            }
            if (String.Equals(row.Source, "empty", StringComparison.OrdinalIgnoreCase))
            {
                return "\u672a\u5339\u914d";
            }
            if (String.Equals(row.Source, "index", StringComparison.OrdinalIgnoreCase))
            {
                return "\u672c\u5730\u7d22\u5f15";
            }

            return row.Source ?? "";
        }

        private void StartDeepSeekRecommendations(List<AiPendingRecommendation> pending, int requestVersion)
        {
            if (pending == null || pending.Count == 0 || !deepSeekSettings.IsAvailable)
            {
                return;
            }

            foreach (List<AiPendingRecommendation> batch in BuildDeepSeekBatches(pending))
            {
                ThreadPool.QueueUserWorkItem(delegate
                {
                    List<DeepSeekSelection> selections = new List<DeepSeekSelection>();
                    string error = "";
                    try
                    {
                        DeepSeekClient client = new DeepSeekClient(deepSeekSettings);
                        selections = client.Rank(batch.Select(p => p.Request).ToList());
                    }
                    catch (Exception ex)
                    {
                        error = ex.Message;
                        QuotaRecommendPanel.Log("DeepSeek recommendation failed: " + ex.Message);
                        if (!IsNonRetryableDeepSeekError(error))
                        {
                            selections = RetryDeepSeekOneByOne(batch);
                            if (selections.Count > 0)
                            {
                                error = "";
                            }
                        }
                    }

                    try
                    {
                        if (!IsDisposed && IsHandleCreated)
                        {
                            BeginInvoke((MethodInvoker)delegate
                            {
                                ApplyDeepSeekResults(batch, selections, error, requestVersion);
                            });
                        }
                    }
                    catch
                    {
                    }
                });
            }
        }

        private static List<List<AiPendingRecommendation>> BuildDeepSeekBatches(List<AiPendingRecommendation> pending)
        {
            List<List<AiPendingRecommendation>> batches = new List<List<AiPendingRecommendation>>();
            List<AiPendingRecommendation> current = new List<AiPendingRecommendation>();
            int currentCost = 0;
            foreach (AiPendingRecommendation item in pending ?? new List<AiPendingRecommendation>())
            {
                int cost = EstimateDeepSeekRowCost(item);
                if (current.Count > 0 && (current.Count >= 12 || currentCost + cost > 90))
                {
                    batches.Add(current);
                    current = new List<AiPendingRecommendation>();
                    currentCost = 0;
                }

                current.Add(item);
                currentCost += cost;
            }

            if (current.Count > 0)
            {
                batches.Add(current);
            }

            return batches;
        }

        private static int EstimateDeepSeekRowCost(AiPendingRecommendation pending)
        {
            if (pending == null || pending.Request == null)
            {
                return 1;
            }

            int quotaCount = pending.Request.Candidates == null ? 0 : pending.Request.Candidates.Count;
            int mappingCount = pending.Request.MappingCandidates == null ? 0 : pending.Request.MappingCandidates.Count;
            int textCost = 0;
            if (pending.Request.Item != null)
            {
                textCost = Math.Min(8, ((pending.Request.Item.RawRowText ?? "").Length + (pending.Request.Item.Name ?? "").Length) / 40);
            }

            return 2 + quotaCount + mappingCount * 2 + textCost;
        }

        private void ApplyDeepSeekResults(List<AiPendingRecommendation> batch, List<DeepSeekSelection> selections, string error, int requestVersion)
        {
            if (requestVersion != aiRequestVersion || batch == null)
            {
                return;
            }

            Dictionary<string, DeepSeekSelection> byRow = new Dictionary<string, DeepSeekSelection>(StringComparer.OrdinalIgnoreCase);
            foreach (DeepSeekSelection selection in selections ?? new List<DeepSeekSelection>())
            {
                if (!String.IsNullOrWhiteSpace(selection.RowId) && !byRow.ContainsKey(selection.RowId))
                {
                    byRow[selection.RowId] = selection;
                }
            }

            int applied = 0;
            foreach (AiPendingRecommendation pending in (batch ?? new List<AiPendingRecommendation>()).OrderByDescending(p => p.GridRowIndex))
            {
                RecommendationRow row = pending.Row;
                if (row == null || pending.GridRowIndex < 0 || pending.GridRowIndex >= resultGrid.Rows.Count || pending.GridRowIndex >= recommendations.Count)
                {
                    continue;
                }
                if (!Object.ReferenceEquals(recommendations[pending.GridRowIndex], row))
                {
                    continue;
                }

                row.AiPending = false;
                DeepSeekSelection selection;
                if (!String.IsNullOrWhiteSpace(error))
                {
                    SetRecommendationStatus(pending.GridRowIndex, DeepSeekFailureStatus(error), error);
                    continue;
                }

                if (!byRow.TryGetValue(row.AiRowId, out selection))
                {
                    SetRecommendationStatus(pending.GridRowIndex, selections == null || selections.Count == 0 ? "AI\u8fd4\u56de\u4e3a\u7a7a" : "AI\u65e0\u7ed3\u679c", "");
                    continue;
                }

                if (!String.IsNullOrWhiteSpace(selection.ErrorText))
                {
                    SetRecommendationStatus(pending.GridRowIndex, DeepSeekFailureStatus(selection.ErrorText), selection.ErrorText);
                    continue;
                }

                AiMappingCandidate mappingCandidate = (pending.Request.MappingCandidates ?? new List<AiMappingCandidate>())
                    .FirstOrDefault(c => c != null && String.Equals(c.BoxId, selection.BoxId, StringComparison.OrdinalIgnoreCase));
                if (mappingCandidate != null && selection.Confidence >= 65)
                {
                    List<RecommendationRow> mappedRows = mappingCandidate.ToRecommendations(row.Item, selection.Confidence, selection.Reason);
                    if (mappedRows.Count > 0)
                    {
                        ApplyMappedRowsFromDeepSeek(pending.GridRowIndex, row, mappedRows);
                        applied += mappedRows.Count;
                        continue;
                    }
                }

                AiQuotaCandidate candidate = (pending.Request.Candidates ?? new List<AiQuotaCandidate>())
                    .FirstOrDefault(c => c != null && c.Quota != null && String.Equals(c.Quota.QuotaCode, selection.SelectedCode, StringComparison.OrdinalIgnoreCase));
                if (candidate == null)
                {
                    SetRecommendationStatus(pending.GridRowIndex, String.IsNullOrWhiteSpace(selection.SelectedCode) ? "AI\u65e0\u7ed3\u679c" : "AI\u8fd4\u56de\u65e0\u6548", String.IsNullOrWhiteSpace(selection.SelectedCode) ? selection.Reason : "\u8fd4\u56de\u7f16\u53f7\u4e0d\u5728\u672c\u5730\u5019\u9009\u4e2d");
                    continue;
                }

                int confidence = Math.Max(0, Math.Min(100, selection.Confidence));
                if (confidence < deepSeekSettings.DisplayConfidence ||
                    (!String.Equals(row.Source, "empty", StringComparison.OrdinalIgnoreCase) &&
                        !String.Equals(row.Source, "mapping", StringComparison.OrdinalIgnoreCase) &&
                        confidence < row.Score))
                {
                    SetRecommendationStatus(pending.GridRowIndex, "\u672c\u5730\u4f18\u5148", selection.Reason);
                    continue;
                }

                row.QuotaCode = candidate.Quota.QuotaCode;
                row.QuotaName = candidate.Quota.QuotaName;
                row.QuotaUnit = candidate.Quota.QuotaUnit;
                row.ConvertedValueText = RecommendDialog.ConvertQuantityForIndex(row.Item.ValueText, row.Item.Unit, candidate.Quota.QuotaUnit);
                row.Score = confidence;
                row.Reason = "DeepSeek\u5728\u672c\u5730\u5019\u9009\u4e2d\u9009\u62e9" + (String.IsNullOrWhiteSpace(selection.Reason) ? "" : "\uff1a" + selection.Reason);
                row.Source = "deepseek";
                row.TargetKind = "quota";
                row.BoxId = "";

                DataGridViewRow gridRow = resultGrid.Rows[pending.GridRowIndex];
                gridRow.Cells["QuotaCode"].Value = row.QuotaCode;
                gridRow.Cells["QuotaName"].Value = row.QuotaName;
                gridRow.Cells["QuotaUnit"].Value = row.QuotaUnit;
                gridRow.Cells["QuotaQuantity"].Value = row.ConvertedValueText;
                gridRow.Cells["Checked"].Value = confidence >= deepSeekSettings.AutoCheckConfidence &&
                    RecommendDialog.UnitCompatibleForIndex(row.Item.Unit, row.QuotaUnit);
                SetRecommendationStatus(pending.GridRowIndex, "AI\u5df2\u8865\u63a8", row.Reason);
                applied++;
            }

            if (applied > 0)
            {
                statusLabel.Text = statusLabel.Text + " AI\u5df2\u8865\u63a8 " + applied.ToString(CultureInfo.InvariantCulture) + " \u884c\u3002";
            }
        }

        private List<DeepSeekSelection> RetryDeepSeekOneByOne(List<AiPendingRecommendation> batch)
        {
            List<DeepSeekSelection> selections = new List<DeepSeekSelection>();
            foreach (AiPendingRecommendation pending in batch ?? new List<AiPendingRecommendation>())
            {
                if (pending == null || pending.Request == null)
                {
                    continue;
                }

                try
                {
                    DeepSeekClient client = new DeepSeekClient(deepSeekSettings);
                    List<DeepSeekSelection> result = client.Rank(new List<DeepSeekRequestRow> { pending.Request });
                    if (result != null && result.Count > 0)
                    {
                        selections.AddRange(result);
                    }
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("DeepSeek single-row retry failed: " + ex.Message);
                    selections.Add(new DeepSeekSelection
                    {
                        RowId = pending.Request.RowId,
                        Confidence = 0,
                        ErrorText = ex.Message
                    });
                }
            }

            return selections;
        }

        private static bool IsNonRetryableDeepSeekError(string error)
        {
            string value = (error ?? "").ToLowerInvariant();
            return value.Contains("401") ||
                value.Contains("402") ||
                value.Contains("authentication") ||
                value.Contains("api key") ||
                value.Contains("balance") ||
                value.Contains("insufficient") ||
                value.Contains("422");
        }

        private static string DeepSeekFailureStatus(string error)
        {
            string value = (error ?? "").ToLowerInvariant();
            if (value.Contains("401") || value.Contains("authentication") || value.Contains("api key"))
            {
                return "AI Key\u5f02\u5e38";
            }
            if (value.Contains("402") || value.Contains("balance") || value.Contains("insufficient"))
            {
                return "AI\u4f59\u989d\u4e0d\u8db3";
            }
            if (value.Contains("429") || value.Contains("rate limit"))
            {
                return "AI\u9650\u6d41";
            }
            if (value.Contains("timeout") || value.Contains("timed out") || value.Contains("\u8d85\u65f6"))
            {
                return "AI\u8d85\u65f6";
            }
            if (value.Contains("500") || value.Contains("503") || value.Contains("server") || value.Contains("overload"))
            {
                return "AI\u670d\u52a1\u5f02\u5e38";
            }
            if (value.Contains("json") || value.Contains("invalid") || value.Contains("format") || value.Contains("422"))
            {
                return "AI\u8fd4\u56de\u65e0\u6548";
            }
            if (value.Contains("connect") || value.Contains("network") || value.Contains("name resolution") || value.Contains("remote") || value.Contains("\u7f51\u7edc"))
            {
                return "AI\u7f51\u7edc\u5931\u8d25";
            }

            return "AI\u5931\u8d25";
        }

        private void ApplyMappedRowsFromDeepSeek(int gridRowIndex, RecommendationRow oldRow, List<RecommendationRow> mappedRows)
        {
            if (mappedRows == null || mappedRows.Count == 0 || gridRowIndex < 0 || gridRowIndex >= recommendations.Count || gridRowIndex >= resultGrid.Rows.Count)
            {
                return;
            }

            RecommendationRow first = mappedRows[0];
            recommendations[gridRowIndex] = first;
            first.GridRowIndex = gridRowIndex;
            DataGridViewRow gridRow = resultGrid.Rows[gridRowIndex];
            gridRow.Cells["Checked"].Value = first.Score >= 60;
            gridRow.Cells["QuotaCode"].Value = first.QuotaCode;
            gridRow.Cells["QuotaName"].Value = first.QuotaName;
            gridRow.Cells["QuotaUnit"].Value = first.QuotaUnit;
            gridRow.Cells["QuotaQuantity"].Value = first.ConvertedValueText;
            gridRow.Cells["SourceStatus"].Value = "AI\u5bf9\u5e94\u6846";
            gridRow.Cells["SourceStatus"].ToolTipText = first.Reason ?? "";

            int insertAt = gridRowIndex + 1;
            for (int i = 1; i < mappedRows.Count; i++)
            {
                RecommendationRow mapped = mappedRows[i];
                mapped.GridRowIndex = insertAt;
                recommendations.Insert(insertAt, mapped);
                resultGrid.Rows.Insert(insertAt, mapped.Score >= 60, "", "", mapped.Item.Unit, mapped.Item.ValueText, mapped.QuotaCode, mapped.QuotaName, mapped.QuotaUnit, mapped.ConvertedValueText, "AI\u5bf9\u5e94\u6846");
                DataGridViewRow continuation = resultGrid.Rows[insertAt];
                continuation.Cells["Correct"] = new DataGridViewTextBoxCell();
                continuation.Cells["Correct"].Value = "";
                continuation.Cells["Correct"].ReadOnly = true;
                continuation.Cells["SourceStatus"].ToolTipText = mapped.Reason ?? "";
                insertAt++;
            }

            for (int i = 0; i < recommendations.Count; i++)
            {
                recommendations[i].GridRowIndex = i;
            }
        }

        private void SetRecommendationStatus(int gridRowIndex, string text, string tooltip)
        {
            if (gridRowIndex < 0 || gridRowIndex >= resultGrid.Rows.Count)
            {
                return;
            }

            DataGridViewCell cell = resultGrid.Rows[gridRowIndex].Cells["SourceStatus"];
            cell.Value = text ?? "";
            cell.ToolTipText = tooltip ?? "";
        }

        private void ShowDeepSeekSettings()
        {
            using (DeepSeekSettingsDialog dialog = new DeepSeekSettingsDialog(deepSeekSettings))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                deepSeekSettings = dialog.Settings;
                aiRequestVersion++;
                statusLabel.Text = deepSeekSettings.IsAvailable
                    ? "DeepSeek AI\u8bbe\u7f6e\u5df2\u4fdd\u5b58\uff0c\u540e\u7eed\u91cd\u65b0\u8bfb\u53d6Excel/\u526a\u8d34\u677f\u65f6\u751f\u6548\u3002"
                    : "DeepSeek AI\u8bbe\u7f6e\u5df2\u4fdd\u5b58\uff0c\u5f53\u524d\u672a\u542f\u7528AI\u8865\u63a8\u3002";
            }
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

                selection.Items.AddRange(BuildQuantityItemsFromRange(range, selection.WorksheetName));
                ApplyActiveLeftGroups(selection);
                NormalizeSelectionItems(selection);
                LogSelectionSummary("Excel selection", selection);

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

                string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                ExcelSelection textSelection = BuildSelectionFromClipboardLines(lines);

                ExcelSelection activeSelection;
                if (TryReadActiveExcelSelectionForClipboard(lines.Length, out activeSelection))
                {
                    if (textSelection.Items.Count > 0 && SelectionGroupScore(textSelection) > SelectionGroupScore(activeSelection))
                    {
                        selection = textSelection;
                        ApplyActiveLeftGroups(selection);
                        NormalizeSelectionItems(selection);
                        LogSelectionSummary("Clipboard text selection preferred over active selection", selection);
                        return true;
                    }

                    selection = activeSelection;
                    LogSelectionSummary("Clipboard backed by active Excel selection", selection);
                    return true;
                }

                if (textSelection.Items.Count > 0)
                {
                    selection = textSelection;
                    ApplyActiveLeftGroups(selection);
                    NormalizeSelectionItems(selection);
                    LogSelectionSummary("Clipboard text selection", selection);
                    return true;
                }

                if (TryReadClipboardHtmlSelection(out selection))
                {
                    ApplyActiveLeftGroups(selection);
                    NormalizeSelectionItems(selection);
                    LogSelectionSummary("Clipboard HTML selection", selection);
                    return true;
                }

                selection = textSelection;
                LogSelectionSummary("Clipboard text selection", selection);

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

        private static bool TryReadActiveExcelSelectionForClipboard(int expectedRows, out ExcelSelection selection)
        {
            selection = null;
            try
            {
                dynamic excel = GetActiveSpreadsheetApplication();
                if (excel == null)
                {
                    return false;
                }

                dynamic workbook = excel.ActiveWorkbook;
                dynamic sheet = excel.ActiveSheet;
                dynamic range = excel.Selection;
                if (workbook == null || sheet == null || range == null)
                {
                    return false;
                }

                int rowCount = Convert.ToInt32(range.Rows.Count, CultureInfo.InvariantCulture);
                bool rowCountDiffers = expectedRows > 0 && rowCount != expectedRows;

                selection = new ExcelSelection();
                selection.WorkbookPath = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture);
                selection.WorksheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture);

                selection.Items.AddRange(BuildQuantityItemsFromRange(range, selection.WorksheetName));
                ApplyActiveLeftGroups(selection);
                NormalizeSelectionItems(selection);
                if (selection.Items.Count == 0)
                {
                    return false;
                }

                if (rowCountDiffers)
                {
                    QuotaRecommendPanel.Log("Clipboard line count differs from active selection rows. textLines="
                        + expectedRows.ToString(CultureInfo.InvariantCulture)
                        + ", selectionRows=" + rowCount.ToString(CultureInfo.InvariantCulture)
                        + ", parsedItems=" + selection.Items.Count.ToString(CultureInfo.InvariantCulture));
                    if (rowCount < expectedRows)
                    {
                        return true;
                    }

                    return selection.Items.Count <= expectedRows;
                }

                return true;
            }
            catch
            {
                selection = null;
                return false;
            }
        }

        private static ExcelSelection BuildSelectionFromClipboardLines(string[] lines)
        {
            ExcelSelection selection = new ExcelSelection();
            selection.WorksheetName = "\u526a\u8d34\u677f";
            List<List<string>> textTable = new List<List<string>>();
            for (int i = 0; i < lines.Length; i++)
            {
                textTable.Add(lines[i].Split('\t').ToList());
            }

            selection.Items.AddRange(BuildQuantityItemsFromTextTable(textTable, selection.WorksheetName));
            NormalizeSelectionItems(selection);
            return selection;
        }

        private static int SelectionGroupScore(ExcelSelection selection)
        {
            if (selection == null || selection.Items.Count == 0)
            {
                return 0;
            }

            int score = 0;
            foreach (ExcelQuantityItem item in selection.Items)
            {
                if (item == null)
                {
                    continue;
                }

                string name = (item.Name ?? "").Trim();
                string section = (item.SectionName ?? "").Trim();
                if (LooksLikeGroupText(section)
                    && !String.Equals(section, name, StringComparison.Ordinal)
                    && name.IndexOf(section, StringComparison.Ordinal) >= 0)
                {
                    score += 2;
                }
                else if (name.IndexOf(' ') >= 0)
                {
                    score += 1;
                }
            }

            return score;
        }

        private static bool TryReadClipboardHtmlSelection(out ExcelSelection selection)
        {
            selection = null;
            try
            {
                if (!Clipboard.ContainsText(TextDataFormat.Html))
                {
                    return false;
                }

                string html = Clipboard.GetText(TextDataFormat.Html);
                if (String.IsNullOrWhiteSpace(html))
                {
                    return false;
                }

                List<List<string>> table = ParseHtmlTable(html);
                if (table.Count == 0)
                {
                    return false;
                }

                selection = new ExcelSelection();
                selection.WorksheetName = "\u526a\u8d34\u677fHTML";
                selection.Items.AddRange(BuildQuantityItemsFromTextTable(table, selection.WorksheetName));

                NormalizeSelectionItems(selection);

                if (selection.Items.Count == 0)
                {
                    selection = null;
                    return false;
                }

                QuotaRecommendPanel.Log("Clipboard HTML parsed. htmlRows=" + table.Count.ToString(CultureInfo.InvariantCulture) + ", items=" + selection.Items.Count.ToString(CultureInfo.InvariantCulture));
                return true;
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("Clipboard HTML parse failed: " + ex.Message);
                selection = null;
                return false;
            }
        }

        private static List<List<string>> ParseHtmlTable(string html)
        {
            List<List<string>> result = new List<List<string>>();
            Dictionary<int, CarryCell> carries = new Dictionary<int, CarryCell>();
            MatchCollection rowMatches = Regex.Matches(html, @"<tr\b[\s\S]*?</tr>", RegexOptions.IgnoreCase);
            foreach (Match rowMatch in rowMatches)
            {
                string rowHtml = rowMatch.Value;
                MatchCollection cellMatches = Regex.Matches(rowHtml, @"<t[dh]\b[^>]*>[\s\S]*?</t[dh]>", RegexOptions.IgnoreCase);
                if (cellMatches.Count == 0)
                {
                    continue;
                }

                List<string> row = new List<string>();
                int col = 0;
                foreach (Match cellMatch in cellMatches)
                {
                    FillCarriedCells(row, carries, ref col);

                    string cellHtml = cellMatch.Value;
                    int rowSpan = Math.Max(1, ParseSpan(cellHtml, "rowspan"));
                    int colSpan = Math.Max(1, ParseSpan(cellHtml, "colspan"));
                    string text = HtmlCellText(cellHtml);
                    for (int i = 0; i < colSpan; i++)
                    {
                        SetListValue(row, col + i, text);
                        if (rowSpan > 1)
                        {
                            carries[col + i] = new CarryCell { Text = text, RemainingRows = rowSpan - 1 };
                        }
                    }

                    col += colSpan;
                }

                FillCarriedCells(row, carries, ref col);
                if (row.Any(v => !String.IsNullOrWhiteSpace(v)))
                {
                    result.Add(row);
                }
            }

            return result;
        }

        private static void FillCarriedCells(List<string> row, Dictionary<int, CarryCell> carries, ref int col)
        {
            while (carries.ContainsKey(col))
            {
                CarryCell carry = carries[col];
                SetListValue(row, col, carry.Text);
                carry.RemainingRows--;
                if (carry.RemainingRows <= 0)
                {
                    carries.Remove(col);
                }
                else
                {
                    carries[col] = carry;
                }

                col++;
            }
        }

        private static int ParseSpan(string html, string name)
        {
            Match match = Regex.Match(html, name + "\\s*=\\s*[\"']?(\\d+)", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                return 1;
            }

            int value;
            return Int32.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out value) ? value : 1;
        }

        private static string HtmlCellText(string cellHtml)
        {
            string text = Regex.Replace(cellHtml, @"<br\s*/?>", "\n", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"<[^>]+>", "");
            return NormalizeCellText(WebUtility.HtmlDecode(text));
        }

        private static void SetListValue(List<string> row, int index, string value)
        {
            while (row.Count <= index)
            {
                row.Add("");
            }

            row[index] = value;
        }

        private static List<ExcelQuantityItem> BuildQuantityItemsFromRange(dynamic range, string worksheetName)
        {
            List<List<CellValue>> rows = new List<List<CellValue>>();
            int rowCount = Convert.ToInt32(range.Rows.Count, CultureInfo.InvariantCulture);
            int columnCount = Convert.ToInt32(range.Columns.Count, CultureInfo.InvariantCulture);
            for (int r = 1; r <= rowCount; r++)
            {
                List<CellValue> row = new List<CellValue>();
                string leftGroup = TryReadLeftGroupFromRangeRow(range, r);
                if (LooksLikeGroupText(leftGroup))
                {
                    CellValue value = new CellValue();
                    value.Text = leftGroup;
                    value.Formula = "";
                    value.Address = "LEFT" + r.ToString(CultureInfo.InvariantCulture);
                    value.RowNumber = r;
                    value.SourceIndex = 0;
                    row.Add(value);
                }

                for (int c = 1; c <= columnCount; c++)
                {
                    dynamic cell = range.Cells[r, c];
                    string text = ReadCellTextWithMerge(cell);
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
                        row.Add(value);
                    }
                }

                rows.Add(row);
            }

            return BuildQuantityItemsFromCellRows(rows, worksheetName);
        }

        private static List<ExcelQuantityItem> BuildQuantityItemsFromTextTable(List<List<string>> table, string worksheetName)
        {
            List<List<CellValue>> rows = new List<List<CellValue>>();
            for (int r = 0; r < table.Count; r++)
            {
                List<CellValue> row = new List<CellValue>();
                List<string> source = table[r] ?? new List<string>();
                for (int c = 0; c < source.Count; c++)
                {
                    string text = NormalizeCellText(source[c]);
                    if (String.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    CellValue value = new CellValue();
                    value.Text = text;
                    value.Formula = text.StartsWith("=", StringComparison.Ordinal) ? text : "";
                    value.Address = "R" + (r + 1).ToString(CultureInfo.InvariantCulture) + "C" + (c + 1).ToString(CultureInfo.InvariantCulture);
                    value.RowNumber = r + 1;
                    value.SourceIndex = c + 1;
                    row.Add(value);
                }

                rows.Add(row);
            }

            return BuildQuantityItemsFromCellRows(rows, worksheetName);
        }

        private static List<ExcelQuantityItem> BuildQuantityItemsFromCellRows(List<List<CellValue>> rows, string worksheetName)
        {
            List<ExcelQuantityItem> result = new List<ExcelQuantityItem>();
            if (rows == null || rows.Count == 0)
            {
                return result;
            }

            int quantityColumn = FindQuantityColumn(rows);
            if (quantityColumn < 0)
            {
                return result;
            }

            int unitColumn = FindUnitColumn(rows, quantityColumn);
            int detailColumn = FindDetailColumn(rows, quantityColumn, unitColumn);
            if (detailColumn < 0)
            {
                return result;
            }

            int groupColumn = FindGroupColumn(rows, detailColumn, unitColumn, quantityColumn);
            string[] units = BuildUnitsByRow(rows, unitColumn);
            string currentGroup = "";
            for (int i = 0; i < rows.Count; i++)
            {
                List<CellValue> row = rows[i] ?? new List<CellValue>();
                CellValue quantityCell = GetCell(row, quantityColumn);
                if (quantityCell == null || !IsQuantityLike(quantityCell.Text))
                {
                    continue;
                }

                string detail = GetCellText(row, detailColumn);
                if (String.IsNullOrWhiteSpace(detail) || LooksLikeOrderOrHeader(detail) || LooksLikeUnit(detail))
                {
                    detail = PickQuantityName(row, quantityCell, unitColumn >= 0 ? GetCell(row, unitColumn) : null);
                }

                if (String.IsNullOrWhiteSpace(detail) || LooksLikeOrderOrHeader(detail))
                {
                    continue;
                }

                string group = groupColumn >= 0 ? GetCellText(row, groupColumn) : "";
                if (LooksLikeGroupText(group))
                {
                    currentGroup = group;
                }
                else
                {
                    group = currentGroup;
                }

                if (String.Equals(group, detail, StringComparison.Ordinal))
                {
                    group = "";
                }

                string name = CombineQuantityName(group, detail);
                if (String.IsNullOrWhiteSpace(name))
                {
                    continue;
                }

                ExcelQuantityItem item = new ExcelQuantityItem();
                item.WorksheetName = worksheetName;
                item.RowNumber = quantityCell.RowNumber;
                item.CellAddress = quantityCell.Address;
                item.Name = name;
                item.Unit = i < units.Length ? units[i] : "";
                item.ValueText = quantityCell.Text;
                item.Formula = quantityCell.Formula;
                item.ContextText = name + " " + item.Unit + " " + item.ValueText;
                item.SectionName = String.IsNullOrWhiteSpace(group) ? detail : group;
                item.OriginalName = name;
                item.RawRowText = BuildRawRowText(row);
                result.Add(item);
            }

            QuotaRecommendPanel.Log("Grid parser: rows=" + rows.Count.ToString(CultureInfo.InvariantCulture)
                + ", qtyCol=" + quantityColumn.ToString(CultureInfo.InvariantCulture)
                + ", unitCol=" + unitColumn.ToString(CultureInfo.InvariantCulture)
                + ", detailCol=" + detailColumn.ToString(CultureInfo.InvariantCulture)
                + ", groupCol=" + groupColumn.ToString(CultureInfo.InvariantCulture)
                + ", items=" + result.Count.ToString(CultureInfo.InvariantCulture));
            return result;
        }

        private static int FindQuantityColumn(List<List<CellValue>> rows)
        {
            var candidates = rows.SelectMany(r => r)
                .Where(c => c != null && IsQuantityLike(c.Text) && !LooksLikeUnit(c.Text))
                .GroupBy(c => c.SourceIndex)
                .Select(g => new { Column = g.Key, Count = g.Count() })
                .OrderByDescending(x => x.Count)
                .ThenByDescending(x => x.Column)
                .ToList();
            return candidates.Count == 0 ? -1 : candidates[0].Column;
        }

        private static int FindUnitColumn(List<List<CellValue>> rows, int quantityColumn)
        {
            var candidates = rows.SelectMany(r => r)
                .Where(c => c != null && c.SourceIndex != quantityColumn && LooksLikeUnit(c.Text))
                .GroupBy(c => c.SourceIndex)
                .Select(g => new { Column = g.Key, Count = g.Count() })
                .OrderByDescending(x => x.Column < quantityColumn)
                .ThenByDescending(x => x.Count)
                .ThenBy(x => Math.Abs(quantityColumn - x.Column))
                .ToList();
            return candidates.Count == 0 ? -1 : candidates[0].Column;
        }

        private static int FindDetailColumn(List<List<CellValue>> rows, int quantityColumn, int unitColumn)
        {
            int rightLimit = unitColumn >= 0 && unitColumn < quantityColumn ? unitColumn : quantityColumn;
            var candidates = rows.SelectMany(r => r)
                .Where(c => c != null
                    && c.SourceIndex < rightLimit
                    && c.SourceIndex != unitColumn
                    && LooksLikeGroupText(c.Text))
                .GroupBy(c => c.SourceIndex)
                .Select(g => new { Column = g.Key, Count = g.Count() })
                .OrderByDescending(x => x.Column)
                .ThenByDescending(x => x.Count)
                .ToList();
            return candidates.Count == 0 ? -1 : candidates[0].Column;
        }

        private static int FindGroupColumn(List<List<CellValue>> rows, int detailColumn, int unitColumn, int quantityColumn)
        {
            var candidates = rows.SelectMany(r => r)
                .Where(c => c != null
                    && c.SourceIndex < detailColumn
                    && c.SourceIndex != unitColumn
                    && c.SourceIndex != quantityColumn
                    && !String.IsNullOrWhiteSpace(c.Text)
                    && !LooksLikeUnit(c.Text)
                    && !IsQuantityLike(c.Text)
                    && !LooksLikeOrderOrHeader(c.Text))
                .GroupBy(c => c.SourceIndex)
                .Select(g => new { Column = g.Key, Count = g.Count() })
                .OrderByDescending(x => x.Column)
                .ThenByDescending(x => x.Count)
                .ToList();
            return candidates.Count == 0 ? -1 : candidates[0].Column;
        }

        private static string[] BuildUnitsByRow(List<List<CellValue>> rows, int unitColumn)
        {
            string[] units = new string[rows.Count];
            for (int i = 0; i < rows.Count; i++)
            {
                CellValue cell = unitColumn >= 0 ? GetCell(rows[i], unitColumn) : null;
                if (cell != null && LooksLikeUnit(cell.Text))
                {
                    units[i] = cell.Text.Trim();
                }
            }

            List<string> knownUnits = units.Where(u => !String.IsNullOrWhiteSpace(u)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (knownUnits.Count == 1)
            {
                for (int i = 0; i < units.Length; i++)
                {
                    units[i] = knownUnits[0];
                }

                return units;
            }

            string lastUnit = "";
            for (int i = 0; i < units.Length; i++)
            {
                if (!String.IsNullOrWhiteSpace(units[i]))
                {
                    lastUnit = units[i];
                }
                else if (!String.IsNullOrWhiteSpace(lastUnit))
                {
                    units[i] = lastUnit;
                }
            }

            string nextUnit = "";
            for (int i = units.Length - 1; i >= 0; i--)
            {
                if (!String.IsNullOrWhiteSpace(units[i]))
                {
                    nextUnit = units[i];
                }
                else if (!String.IsNullOrWhiteSpace(nextUnit))
                {
                    units[i] = nextUnit;
                }
            }

            return units;
        }

        private static CellValue GetCell(List<CellValue> row, int sourceIndex)
        {
            if (row == null)
            {
                return null;
            }

            return row.FirstOrDefault(c => c.SourceIndex == sourceIndex);
        }

        private static string GetCellText(List<CellValue> row, int sourceIndex)
        {
            CellValue cell = GetCell(row, sourceIndex);
            return cell == null ? "" : (cell.Text ?? "").Trim();
        }

        private static string BuildRawRowText(List<CellValue> row)
        {
            return String.Join(" ", (row ?? new List<CellValue>())
                .OrderBy(c => c.SourceIndex)
                .Select(c => c == null ? "" : c.Text)
                .Where(t => !String.IsNullOrWhiteSpace(t))
                .ToArray());
        }

        private static ExcelQuantityItem BuildQuantityItemFromHtmlRow(List<string> cells, int rowNumber, ref string lastGroupName, ref string lastUnit)
        {
            string[] raw = cells.ToArray();
            ExcelQuantityItem four = TryBuildFourColumnClipboardItem(raw, rowNumber, ref lastGroupName, ref lastUnit);
            if (four != null)
            {
                four.WorksheetName = "\u526a\u8d34\u677fHTML";
                return four;
            }

            ExcelQuantityItem three = TryBuildThreeColumnClipboardItem(raw, rowNumber, ref lastUnit);
            if (three != null)
            {
                three.WorksheetName = "\u526a\u8d34\u677fHTML";
                return three;
            }

            return null;
        }

        private static ExcelQuantityItem BuildQuantityItemFromRangeRow(dynamic range, string worksheetName, int relativeRow, int columnCount, ref string lastGroupName, ref string lastUnit)
        {
            ExcelQuantityItem fourColumnItem = TryBuildFourColumnRangeItem(range, worksheetName, relativeRow, columnCount, ref lastGroupName, ref lastUnit);
            if (fourColumnItem != null)
            {
                return fourColumnItem;
            }

            ExcelQuantityItem fixedItem = TryBuildThreeColumnRangeItem(range, worksheetName, relativeRow, columnCount, ref lastUnit);
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

        private static ExcelQuantityItem TryBuildFourColumnRangeItem(dynamic range, string worksheetName, int relativeRow, int columnCount, ref string lastGroupName, ref string lastUnit)
        {
            if (columnCount < 4)
            {
                return null;
            }

            dynamic groupCell = range.Cells[relativeRow, 1];
            dynamic detailCell = range.Cells[relativeRow, 2];
            dynamic unitCell = range.Cells[relativeRow, 3];
            dynamic quantityCell = range.Cells[relativeRow, 4];
            string group = ReadCellTextWithMerge(groupCell);
            string detail = ReadCellTextWithMerge(detailCell);
            string unit = ReadCellTextWithMerge(unitCell);
            string quantity = ExcelValueToText(quantityCell.Value2);

            if (!String.IsNullOrWhiteSpace(group) && !LooksLikeOrderOrHeader(group))
            {
                lastGroupName = group;
            }
            else
            {
                group = lastGroupName;
            }

            if (!String.IsNullOrWhiteSpace(unit) && !LooksLikeOrderOrHeader(unit))
            {
                lastUnit = unit;
            }
            else
            {
                unit = lastUnit;
            }

            string name = CombineQuantityName(group, detail);
            if (String.IsNullOrWhiteSpace(name) || String.IsNullOrWhiteSpace(quantity))
            {
                return null;
            }

            if (LooksLikeOrderOrHeader(name) || LooksLikeOrderOrHeader(detail) || !IsQuantityLike(quantity))
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
            item.SectionName = group;
            item.OriginalName = name;
            item.RawRowText = item.ContextText;
            item.SkipAiNameNormalization = true;
            return item;
        }

        private static ExcelQuantityItem TryBuildThreeColumnRangeItem(dynamic range, string worksheetName, int relativeRow, int columnCount, ref string lastUnit)
        {
            if (columnCount < 3)
            {
                return null;
            }

            dynamic nameCell = range.Cells[relativeRow, 1];
            dynamic unitCell = range.Cells[relativeRow, 2];
            dynamic quantityCell = range.Cells[relativeRow, 3];
            string name = ReadCellTextWithMerge(nameCell);
            string unit = ReadCellTextWithMerge(unitCell);
            string quantity = ExcelValueToText(quantityCell.Value2);
            string group = TryReadLeftGroupFromRangeRow(range, relativeRow);
            if (!String.IsNullOrWhiteSpace(group) && !LooksLikeOrderOrHeader(group))
            {
                name = CombineQuantityName(group, name);
            }

            if (!String.IsNullOrWhiteSpace(unit) && !LooksLikeOrderOrHeader(unit))
            {
                lastUnit = unit;
            }
            else
            {
                unit = lastUnit;
            }
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
            item.SectionName = String.IsNullOrWhiteSpace(group) ? name : group;
            item.OriginalName = name;
            item.RawRowText = item.ContextText;
            item.SkipAiNameNormalization = true;
            return item;
        }

        private static List<string> TryReadActiveSelectionLeftGroups(int expectedRows)
        {
            List<string> result = new List<string>();
            try
            {
                dynamic excel = GetActiveSpreadsheetApplication();
                if (excel == null)
                {
                    return result;
                }

                dynamic range = excel.Selection;
                if (range == null)
                {
                    return result;
                }

                int rowCount = Convert.ToInt32(range.Rows.Count, CultureInfo.InvariantCulture);
                int count = expectedRows > 0 ? Math.Min(expectedRows, rowCount) : rowCount;
                string lastGroup = "";
                for (int row = 1; row <= count; row++)
                {
                    string group = TryReadLeftGroupFromRangeRow(range, row);
                    if (LooksLikeGroupText(group))
                    {
                        lastGroup = group;
                    }
                    else
                    {
                        group = lastGroup;
                    }

                    result.Add(group);
                }
            }
            catch
            {
            }

            return result;
        }

        private static string TryReadLeftGroupFromRangeRow(dynamic range, int relativeRow)
        {
            try
            {
                dynamic firstCell = range.Cells[relativeRow, 1];
                int row = TryReadInt(firstCell, "Row");
                int column = TryReadInt(firstCell, "Column");
                dynamic worksheet = null;
                try
                {
                    worksheet = firstCell.Worksheet;
                }
                catch
                {
                }

                int maxOffset = column > 1 ? Math.Min(8, column - 1) : 8;
                for (int offset = 1; offset <= maxOffset; offset++)
                {
                    string text = "";
                    if (worksheet != null && row > 0 && column > offset)
                    {
                        try
                        {
                            dynamic sheetCell = worksheet.Cells[row, column - offset];
                            text = ReadCellTextWithMerge(sheetCell);
                        }
                        catch
                        {
                            text = "";
                        }
                    }

                    if (!LooksLikeGroupText(text))
                    {
                        try
                        {
                            dynamic offsetCell = firstCell.Offset[0, -offset];
                            text = ReadCellTextWithMerge(offsetCell);
                        }
                        catch
                        {
                            text = "";
                        }
                    }

                    if (LooksLikeGroupText(text))
                    {
                        return text;
                    }
                }
            }
            catch
            {
            }

            return "";
        }

        private static int TryReadInt(dynamic source, string propertyName)
        {
            try
            {
                object value = source.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, null, source, null, CultureInfo.InvariantCulture);
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                try
                {
                    object value = propertyName == "Row" ? source.Row : source.Column;
                    return Convert.ToInt32(value, CultureInfo.InvariantCulture);
                }
                catch
                {
                    return 0;
                }
            }
        }

        private static string ReadCellTextWithMerge(dynamic cell)
        {
            try
            {
                string text = ExcelValueToText(cell.Value2);
                if (!String.IsNullOrWhiteSpace(text))
                {
                    return text;
                }

                dynamic mergeArea = cell.MergeArea;
                if (mergeArea != null)
                {
                    dynamic first = mergeArea.Cells[1, 1];
                    return ExcelValueToText(first.Value2);
                }
            }
            catch
            {
            }

            return "";
        }

        private static ExcelQuantityItem BuildQuantityItemFromTextRow(string[] rawCells, int rowNumber, ref string lastGroupName, ref string lastUnit)
        {
            ExcelQuantityItem fourColumnItem = TryBuildFourColumnClipboardItem(rawCells, rowNumber, ref lastGroupName, ref lastUnit);
            if (fourColumnItem != null)
            {
                return fourColumnItem;
            }

            ExcelQuantityItem fixedItem = TryBuildThreeColumnClipboardItem(rawCells, rowNumber, ref lastUnit);
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

        private static ExcelQuantityItem TryBuildFourColumnClipboardItem(string[] rawCells, int rowNumber, ref string lastGroupName, ref string lastUnit)
        {
            if (rawCells == null || rawCells.Length < 4)
            {
                return null;
            }

            string group = NormalizeCellText(rawCells[0]);
            string detail = NormalizeCellText(rawCells[1]);
            string unit = NormalizeCellText(rawCells[2]);
            string quantity = NormalizeCellText(rawCells[3]);

                if (LooksLikeGroupText(group))
            {
                lastGroupName = group;
            }
            else
            {
                group = lastGroupName;
            }

            if (!String.IsNullOrWhiteSpace(unit) && !LooksLikeOrderOrHeader(unit))
            {
                lastUnit = unit;
            }
            else
            {
                unit = lastUnit;
            }

            string name = CombineQuantityName(group, detail);
            if (String.IsNullOrWhiteSpace(name) || String.IsNullOrWhiteSpace(quantity))
            {
                return null;
            }

            if (LooksLikeOrderOrHeader(name) || LooksLikeOrderOrHeader(detail) || !IsQuantityLike(quantity))
            {
                return null;
            }

            ExcelQuantityItem item = new ExcelQuantityItem();
            item.WorksheetName = "\u526a\u8d34\u677f";
            item.RowNumber = rowNumber;
            item.CellAddress = "R" + rowNumber.ToString(CultureInfo.InvariantCulture) + "C4";
            item.Name = name;
            item.Unit = unit;
            item.ValueText = quantity;
            item.Formula = quantity.StartsWith("=", StringComparison.Ordinal) ? quantity : "";
            item.ContextText = name + " " + unit + " " + quantity;
            item.SectionName = group;
            item.OriginalName = name;
            item.RawRowText = item.ContextText;
            item.SkipAiNameNormalization = true;
            return item;
        }

        private static string CombineQuantityName(string group, string detail)
        {
            string left = (group ?? "").Trim();
            string right = (detail ?? "").Trim();
            if (String.IsNullOrWhiteSpace(left))
            {
                return right;
            }

            if (String.IsNullOrWhiteSpace(right))
            {
                return left;
            }

            if (right.IndexOf(left, StringComparison.Ordinal) >= 0)
            {
                return right;
            }

            return left + " " + right;
        }

        private static void ApplyActiveLeftGroups(ExcelSelection selection)
        {
            if (selection == null || selection.Items.Count == 0)
            {
                return;
            }

            List<string> groups = TryReadActiveSelectionLeftGroups(selection.Items.Count);
            if (groups.Count == 0)
            {
                return;
            }

            int groupHits = groups.Count(g => LooksLikeGroupText(g));
            string sampleGroup = groups.FirstOrDefault(g => LooksLikeGroupText(g)) ?? "";
            QuotaRecommendPanel.Log("Active left group scan: rows="
                + groups.Count.ToString(CultureInfo.InvariantCulture)
                + ", groups=" + groupHits.ToString(CultureInfo.InvariantCulture)
                + ", sample=" + sampleGroup);

            int count = Math.Min(selection.Items.Count, groups.Count);
            for (int i = 0; i < count; i++)
            {
                ExcelQuantityItem item = selection.Items[i];
                string group = (groups[i] ?? "").Trim();
                if (item == null || String.IsNullOrWhiteSpace(group) || LooksLikeOrderOrHeader(group))
                {
                    continue;
                }

                string section = (item.SectionName ?? "").Trim();
                bool alreadyHasGroup = (item.Name ?? "").IndexOf(group, StringComparison.Ordinal) >= 0;
                bool appearsUngrouped = String.IsNullOrWhiteSpace(section) || String.Equals(section, item.Name, StringComparison.Ordinal);
                if (alreadyHasGroup || !appearsUngrouped)
                {
                    continue;
                }

                item.Name = CombineQuantityName(group, item.Name);
                item.SectionName = group;
                item.ContextText = item.Name + " " + item.Unit + " " + item.ValueText;
            }
        }

        private static void NormalizeSelectionItems(ExcelSelection selection)
        {
            if (selection == null || selection.Items.Count == 0)
            {
                return;
            }

            string lastUnit = "";
            foreach (ExcelQuantityItem item in selection.Items)
            {
                if (item == null)
                {
                    continue;
                }

                item.Name = (item.Name ?? "").Trim();
                item.Unit = (item.Unit ?? "").Trim();
                item.ValueText = (item.ValueText ?? "").Trim();
                if (!String.IsNullOrWhiteSpace(item.Unit))
                {
                    lastUnit = item.Unit;
                }
                else if (!String.IsNullOrWhiteSpace(lastUnit))
                {
                    item.Unit = lastUnit;
                }
            }

            string nextUnit = "";
            for (int i = selection.Items.Count - 1; i >= 0; i--)
            {
                ExcelQuantityItem item = selection.Items[i];
                if (item == null)
                {
                    continue;
                }

                if (!String.IsNullOrWhiteSpace(item.Unit))
                {
                    nextUnit = item.Unit;
                }
                else if (!String.IsNullOrWhiteSpace(nextUnit))
                {
                    item.Unit = nextUnit;
                }
            }

            List<string> knownUnits = selection.Items
                .Where(i => i != null && !String.IsNullOrWhiteSpace(i.Unit))
                .Select(i => i.Unit.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (knownUnits.Count == 1)
            {
                foreach (ExcelQuantityItem item in selection.Items)
                {
                    if (item != null && String.IsNullOrWhiteSpace(item.Unit))
                    {
                        item.Unit = knownUnits[0];
                    }
                }
            }

            foreach (ExcelQuantityItem item in selection.Items)
            {
                if (item != null)
                {
                    item.Name = (item.Name ?? "").Trim();
                    if (String.IsNullOrWhiteSpace(item.OriginalName))
                    {
                        item.OriginalName = item.Name;
                    }
                    if (String.IsNullOrWhiteSpace(item.RawRowText))
                    {
                        item.RawRowText = item.ContextText;
                    }
                    if (String.IsNullOrWhiteSpace(item.RawRowText))
                    {
                        item.RawRowText = item.Name + " " + item.Unit + " " + item.ValueText;
                    }
                    item.ContextText = item.Name + " " + item.Unit + " " + item.ValueText + " " + item.RawRowText;
                }
            }
        }

        private static void LogSelectionSummary(string source, ExcelSelection selection)
        {
            try
            {
                if (selection == null)
                {
                    QuotaRecommendPanel.Log(source + ": no selection");
                    return;
                }

                StringBuilder builder = new StringBuilder();
                int take = Math.Min(5, selection.Items.Count);
                for (int i = 0; i < take; i++)
                {
                    ExcelQuantityItem item = selection.Items[i];
                    if (item == null)
                    {
                        continue;
                    }

                    if (builder.Length > 0)
                    {
                        builder.Append(" | ");
                    }

                    builder.Append(item.Name);
                    builder.Append("/");
                    builder.Append(item.Unit);
                    builder.Append("/");
                    builder.Append(item.ValueText);
                }

                QuotaRecommendPanel.Log(source + ": items=" + selection.Items.Count.ToString(CultureInfo.InvariantCulture) + ", sample=" + builder.ToString());
            }
            catch
            {
            }
        }

        private static ExcelQuantityItem TryBuildThreeColumnClipboardItem(string[] rawCells, int rowNumber, ref string lastUnit)
        {
            if (rawCells == null || rawCells.Length < 3)
            {
                return null;
            }

            string name = NormalizeCellText(rawCells[0]);
            string unit = NormalizeCellText(rawCells[1]);
            string quantity = NormalizeCellText(rawCells[2]);
            if (!String.IsNullOrWhiteSpace(unit) && !LooksLikeOrderOrHeader(unit))
            {
                lastUnit = unit;
            }
            else
            {
                unit = lastUnit;
            }
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
            item.OriginalName = name;
            item.RawRowText = item.ContextText;
            item.SkipAiNameNormalization = true;
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
            bool knownLayout = TryPickByKnownLayout(cells, out nameCell, out unitCell, out quantityCell);
            if (!knownLayout)
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
            item.OriginalName = name;
            item.RawRowText = item.ContextText;
            item.SkipAiNameNormalization = knownLayout;
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
                return NormalizeCellText(Convert.ToDouble(value, CultureInfo.InvariantCulture).ToString("0.########", CultureInfo.InvariantCulture));
            }
            if (value is decimal)
            {
                return NormalizeCellText(((decimal)value).ToString("0.########", CultureInfo.InvariantCulture));
            }
            return NormalizeCellText(Convert.ToString(value, CultureInfo.CurrentCulture));
        }

        private static string NormalizeCellText(string text)
        {
            return Regex.Replace((text ?? "").Replace("\r", " ").Replace("\n", " "), @"\s+", " ").Trim();
        }

        private static bool LooksLikeUnit(string text)
        {
            string unit = NormalizeRawUnit(text);
            string[] units = new string[] { "m", "m2", "m3", "kg", "t", "处", "个", "座", "项", "根", "孔", "环", "组" };
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
            return NormalizeRawUnit(unit)
                .Replace("100", "")
                .Replace("10", "")
                .Replace("㎡", "m2")
                .Replace("㎥", "m3")
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
            if (CountChineseLikeChars(value) > 0)
            {
                return false;
            }

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

        private static bool LooksLikeGroupText(string text)
        {
            return !String.IsNullOrWhiteSpace(text)
                && !LooksLikeOrderOrHeader(text)
                && !LooksLikeUnit(text)
                && !IsQuantityLike(text);
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
            public int AiQueued;
        }

        private sealed class CellValue
        {
            public string Text;
            public string Formula;
            public string Address;
            public int RowNumber;
            public int SourceIndex;
        }

        private struct CarryCell
        {
            public string Text;
            public int RemainingRows;
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
        public string OriginalName;
        public string AiName;
        public int AiNameConfidence;
        public string AiNameReason;
        public bool SkipAiNameNormalization;
        public string SectionName;
        public string Unit;
        public string ValueText;
        public string Formula;
        public string ContextText;
        public string RawRowText;
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
        public int GridRowIndex;
        public string AiRowId;
        public bool AiPending;
        public List<AiQuotaCandidate> AiCandidates;
        public List<AiMappingCandidate> AiMappingCandidates;
    }

    internal sealed class AiQuotaCandidate
    {
        public IndexQuota Quota;
        public int LocalScore;
    }

    internal sealed class AiMappingCandidate
    {
        public string BoxId;
        public int LocalScore;
        public string SampleNames;
        public List<MappingTarget> Targets;

        public List<RecommendationRow> ToRecommendations(ExcelQuantityItem item, int score, string reason)
        {
            return (Targets ?? new List<MappingTarget>())
                .OrderBy(t => MappingStore.TargetSortRank(t.TargetKind, t.Code))
                .ThenBy(t => t.Code ?? "", StringComparer.OrdinalIgnoreCase)
                .Select(t =>
                {
                    RecommendationRow row = t.ToRecommendation(item, score, BoxId);
                    row.Reason = String.IsNullOrWhiteSpace(reason) ? "DeepSeek\u5bf9\u5e94\u6846\u68c0\u6d4b" : "DeepSeek\u5bf9\u5e94\u6846\u68c0\u6d4b\uff1a" + reason;
                    return row;
                })
                .ToList();
        }
    }

    internal sealed class DeepSeekRequestRow
    {
        public string RowId;
        public ExcelQuantityItem Item;
        public List<AiQuotaCandidate> Candidates;
        public List<AiMappingCandidate> MappingCandidates;
    }

    internal sealed class DeepSeekNameRequestRow
    {
        public string RowId;
        public ExcelQuantityItem Item;
    }

    internal sealed class AiPendingRecommendation
    {
        public RecommendationRow Row;
        public int GridRowIndex;
        public DeepSeekRequestRow Request;
    }

    internal sealed class DeepSeekSelection
    {
        public string RowId;
        public string BoxId;
        public string SelectedCode;
        public int Confidence;
        public string Reason;
        public string ErrorText;
    }

    internal sealed class DeepSeekNameResult
    {
        public string RowId;
        public string QuantityName;
        public int Confidence;
        public string Reason;
    }

    internal sealed class DeepSeekMappingSelection
    {
        public string BoxId;
        public int Confidence;
        public string Reason;
    }

    internal sealed class DeepSeekSettings
    {
        public bool Enabled;
        public string ApiKey;
        public string Model = "deepseek-v4-pro";
        public string BaseUrl = "https://api.deepseek.com";
        public int TimeoutSeconds = 8;
        public int MaxRowsPerBatch = 8;
        public int MaxCandidatesPerRow = 12;
        public int LocalHighScore = 80;
        public int DisplayConfidence = 65;
        public int AutoCheckConfidence = 85;
        public bool EnableNameNormalization = true;
        public bool EnableMappingDetection = true;
        public bool EnableQuotaRecommendation = true;

        public bool IsAvailable
        {
            get { return !String.IsNullOrWhiteSpace(ApiKey); }
        }

        public bool CanNormalizeNames
        {
            get { return IsAvailable && EnableNameNormalization; }
        }

        public bool CanRecommendQuota
        {
            get { return IsAvailable && EnableQuotaRecommendation; }
        }

        public bool CanDetectMapping
        {
            get { return IsAvailable; }
        }

        public DeepSeekSettings Copy()
        {
            return new DeepSeekSettings
            {
                Enabled = Enabled,
                ApiKey = ApiKey,
                Model = Model,
                BaseUrl = BaseUrl,
                TimeoutSeconds = TimeoutSeconds,
                MaxRowsPerBatch = MaxRowsPerBatch,
                MaxCandidatesPerRow = MaxCandidatesPerRow,
                LocalHighScore = LocalHighScore,
                DisplayConfidence = DisplayConfidence,
                AutoCheckConfidence = AutoCheckConfidence,
                EnableNameNormalization = EnableNameNormalization,
                EnableMappingDetection = EnableMappingDetection,
                EnableQuotaRecommendation = EnableQuotaRecommendation
            };
        }

        public static DeepSeekSettings Load()
        {
            DeepSeekSettings settings = new DeepSeekSettings();
            string path = ConfigPath();
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

                settings.Enabled = ReadBool(values, "enabled", false);
                settings.ApiKey = ReadString(values, "api_key", "");
                settings.Model = ReadString(values, "model", settings.Model);
                settings.BaseUrl = ReadString(values, "base_url", settings.BaseUrl).TrimEnd('/');
                settings.TimeoutSeconds = Clamp(ReadInt(values, "timeout_seconds", settings.TimeoutSeconds), 2, 60);
                settings.MaxRowsPerBatch = Clamp(ReadInt(values, "max_rows_per_batch", settings.MaxRowsPerBatch), 1, 20);
                settings.MaxCandidatesPerRow = Clamp(ReadInt(values, "max_candidates_per_row", settings.MaxCandidatesPerRow), 3, 20);
                settings.LocalHighScore = Clamp(ReadInt(values, "local_high_score", settings.LocalHighScore), 60, 120);
                settings.DisplayConfidence = Clamp(ReadInt(values, "display_confidence", settings.DisplayConfidence), 1, 100);
                settings.AutoCheckConfidence = Clamp(ReadInt(values, "auto_check_confidence", settings.AutoCheckConfidence), 1, 100);
                settings.EnableNameNormalization = ReadBool(values, "enable_name_normalization", settings.EnableNameNormalization);
                settings.EnableMappingDetection = ReadBool(values, "enable_mapping_detection", settings.EnableMappingDetection);
                settings.EnableQuotaRecommendation = ReadBool(values, "enable_quota_recommendation", settings.EnableQuotaRecommendation);
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("Load DeepSeek settings failed: " + ex.Message);
                settings.Enabled = false;
            }

            return settings;
        }

        public void Save()
        {
            Directory.CreateDirectory(LearningStore.FindDataDir());
            string path = ConfigPath();
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("{");
            AppendJson(builder, "enabled", Enabled ? "true" : "false", false, true);
            AppendJson(builder, "api_key", ApiKey ?? "", true, true);
            AppendJson(builder, "model", String.IsNullOrWhiteSpace(Model) ? "deepseek-v4-pro" : Model, true, true);
            AppendJson(builder, "base_url", String.IsNullOrWhiteSpace(BaseUrl) ? "https://api.deepseek.com" : BaseUrl, true, true);
            AppendJson(builder, "timeout_seconds", TimeoutSeconds.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "max_rows_per_batch", MaxRowsPerBatch.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "max_candidates_per_row", MaxCandidatesPerRow.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "local_high_score", LocalHighScore.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "display_confidence", DisplayConfidence.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "auto_check_confidence", AutoCheckConfidence.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "enable_name_normalization", EnableNameNormalization ? "true" : "false", false, true);
            AppendJson(builder, "enable_mapping_detection", EnableMappingDetection ? "true" : "false", false, true);
            AppendJson(builder, "enable_quota_recommendation", EnableQuotaRecommendation ? "true" : "false", false, false);
            builder.AppendLine("}");
            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        public static string ConfigPath()
        {
            return Path.Combine(LearningStore.FindDataDir(), "deepseek-settings.json");
        }

        private static void AppendJson(StringBuilder builder, string key, string value, bool quoteValue, bool comma)
        {
            builder.Append("  \"").Append(EscapeJson(key)).Append("\": ");
            if (quoteValue)
            {
                builder.Append("\"").Append(EscapeJson(value)).Append("\"");
            }
            else
            {
                builder.Append(value);
            }
            if (comma)
            {
                builder.Append(",");
            }
            builder.AppendLine();
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

        private static string ReadString(Dictionary<string, object> values, string key, string fallback)
        {
            object value;
            return values.TryGetValue(key, out value) && value != null ? Convert.ToString(value, CultureInfo.InvariantCulture) : fallback;
        }

        private static int ReadInt(Dictionary<string, object> values, string key, int fallback)
        {
            object value;
            if (!values.TryGetValue(key, out value) || value == null)
            {
                return fallback;
            }

            int parsed;
            return Int32.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) ? parsed : fallback;
        }

        private static bool ReadBool(Dictionary<string, object> values, string key, bool fallback)
        {
            object value;
            if (!values.TryGetValue(key, out value) || value == null)
            {
                return fallback;
            }

            if (value is bool)
            {
                return (bool)value;
            }

            bool parsed;
            return Boolean.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), out parsed) ? parsed : fallback;
        }

        private static int Clamp(int value, int min, int max)
        {
            return Math.Max(min, Math.Min(max, value));
        }
    }

    internal sealed class DeepSeekSettingsDialog : Form
    {
        private readonly CheckBox enabledCheck;
        private readonly CheckBox normalizeNameCheck;
        private readonly CheckBox mappingDetectCheck;
        private readonly CheckBox recommendQuotaCheck;
        private readonly TextBox apiKeyText;
        private readonly CheckBox showKeyCheck;
        private readonly TextBox modelText;
        private readonly TextBox baseUrlText;
        private readonly NumericUpDown timeoutInput;
        private readonly NumericUpDown batchInput;
        private readonly NumericUpDown candidatesInput;
        private readonly NumericUpDown localHighScoreInput;
        private readonly NumericUpDown displayConfidenceInput;
        private readonly NumericUpDown autoCheckConfidenceInput;

        public DeepSeekSettings Settings { get; private set; }

        public DeepSeekSettingsDialog(DeepSeekSettings current)
        {
            Settings = (current ?? new DeepSeekSettings()).Copy();
            Text = "DeepSeek AI\u8bbe\u7f6e";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(560, 300);

            enabledCheck = new CheckBox { Checked = true };
            normalizeNameCheck = new CheckBox { Checked = true };
            mappingDetectCheck = new CheckBox { Checked = true };
            recommendQuotaCheck = new CheckBox { Checked = true };
            batchInput = HiddenNumber(Settings.MaxRowsPerBatch, 1, 20);
            localHighScoreInput = HiddenNumber(Settings.LocalHighScore, 60, 120);
            displayConfidenceInput = HiddenNumber(Settings.DisplayConfidence, 1, 100);

            AddLabel("API Key", 24, 28, 150);
            apiKeyText = new TextBox();
            apiKeyText.Left = 180;
            apiKeyText.Top = 24;
            apiKeyText.Width = 280;
            apiKeyText.UseSystemPasswordChar = true;
            apiKeyText.Text = Settings.ApiKey ?? "";
            Controls.Add(apiKeyText);

            showKeyCheck = new CheckBox();
            showKeyCheck.Left = 470;
            showKeyCheck.Top = 26;
            showKeyCheck.Width = 60;
            showKeyCheck.Text = "\u663e\u793a";
            showKeyCheck.CheckedChanged += delegate { apiKeyText.UseSystemPasswordChar = !showKeyCheck.Checked; };
            Controls.Add(showKeyCheck);

            AddLabel("\u6a21\u578b", 24, 64, 150);
            modelText = AddTextBox(Settings.Model, 180, 60, 280);

            AddLabel("\u63a5\u53e3\u5730\u5740", 24, 100, 150);
            baseUrlText = AddTextBox(Settings.BaseUrl, 180, 96, 280);

            AddLabel("\u8d85\u65f6\u79d2\u6570", 24, 136, 150);
            timeoutInput = AddNumber(Settings.TimeoutSeconds, 180, 132, 2, 60);

            AddLabel("\u6bcf\u884c\u5019\u9009\u6570", 24, 172, 150);
            candidatesInput = AddNumber(Settings.MaxCandidatesPerRow, 180, 168, 3, 20);

            AddLabel("AI\u81ea\u52a8\u52fe\u9009\u7f6e\u4fe1\u5ea6", 24, 208, 150);
            autoCheckConfidenceInput = AddNumber(Settings.AutoCheckConfidence, 180, 204, 1, 100);

            Button saveButton = new Button();
            saveButton.Text = "\u4fdd\u5b58";
            saveButton.Left = 366;
            saveButton.Top = 252;
            saveButton.Width = 80;
            saveButton.DialogResult = DialogResult.None;
            saveButton.Click += delegate { SaveSettings(); };
            Controls.Add(saveButton);

            Button cancelButton = new Button();
            cancelButton.Text = "\u53d6\u6d88";
            cancelButton.Left = 456;
            cancelButton.Top = 252;
            cancelButton.Width = 80;
            cancelButton.DialogResult = DialogResult.Cancel;
            Controls.Add(cancelButton);

            AcceptButton = saveButton;
            CancelButton = cancelButton;
        }

        private void SaveSettings()
        {
            string model = modelText.Text.Trim();
            string baseUrl = baseUrlText.Text.Trim();
            string apiKey = apiKeyText.Text.Trim();

            if (String.IsNullOrWhiteSpace(apiKey))
            {
                MessageBox.Show(this, "\u8bf7\u586b\u5199 DeepSeek API Key\u3002", "DeepSeek AI\u8bbe\u7f6e", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (String.IsNullOrWhiteSpace(model))
            {
                model = "deepseek-v4-pro";
            }
            if (String.IsNullOrWhiteSpace(baseUrl))
            {
                baseUrl = "https://api.deepseek.com";
            }

            DeepSeekSettings updated = new DeepSeekSettings();
            updated.Enabled = !String.IsNullOrWhiteSpace(apiKey);
            updated.ApiKey = apiKey;
            updated.Model = model;
            updated.BaseUrl = baseUrl.TrimEnd('/');
            updated.TimeoutSeconds = Convert.ToInt32(timeoutInput.Value);
            updated.MaxRowsPerBatch = Convert.ToInt32(batchInput.Value);
            updated.MaxCandidatesPerRow = Convert.ToInt32(candidatesInput.Value);
            updated.LocalHighScore = Convert.ToInt32(localHighScoreInput.Value);
            updated.DisplayConfidence = Convert.ToInt32(displayConfidenceInput.Value);
            updated.AutoCheckConfidence = Convert.ToInt32(autoCheckConfidenceInput.Value);
            updated.EnableNameNormalization = true;
            updated.EnableMappingDetection = true;
            updated.EnableQuotaRecommendation = true;

            try
            {
                updated.Save();
                Settings = updated;
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "\u4fdd\u5b58 DeepSeek \u8bbe\u7f6e\u5931\u8d25\uff1a" + ex.Message, "DeepSeek AI\u8bbe\u7f6e", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private TextBox AddTextBox(string text, int left, int top, int width)
        {
            TextBox box = new TextBox();
            box.Left = left;
            box.Top = top;
            box.Width = width;
            box.Text = text ?? "";
            Controls.Add(box);
            return box;
        }

        private NumericUpDown AddNumber(int value, int left, int top, int min, int max)
        {
            NumericUpDown input = new NumericUpDown();
            input.Left = left;
            input.Top = top;
            input.Width = 120;
            input.Minimum = min;
            input.Maximum = max;
            input.Value = Math.Max(min, Math.Min(max, value));
            Controls.Add(input);
            return input;
        }

        private NumericUpDown HiddenNumber(int value, int min, int max)
        {
            NumericUpDown input = new NumericUpDown();
            input.Minimum = min;
            input.Maximum = max;
            input.Value = Math.Max(min, Math.Min(max, value));
            return input;
        }

        private void AddLabel(string text, int left, int top, int width)
        {
            Label label = new Label();
            label.Text = text;
            label.Left = left;
            label.Top = top + 4;
            label.Width = width;
            Controls.Add(label);
        }
    }

    internal sealed class DeepSeekClient
    {
        private readonly DeepSeekSettings settings;
        private readonly JavaScriptSerializer serializer = new JavaScriptSerializer();

        public DeepSeekClient(DeepSeekSettings deepSeekSettings)
        {
            settings = deepSeekSettings;
            serializer.MaxJsonLength = 1024 * 1024 * 4;
        }

        public List<DeepSeekSelection> Rank(List<DeepSeekRequestRow> rows)
        {
            if (rows == null || rows.Count == 0)
            {
                return new List<DeepSeekSelection>();
            }

            return ParseResponse(SendRequest(BuildUnifiedRequestJson(rows)));
        }

        public List<DeepSeekNameResult> NormalizeQuantityNames(List<DeepSeekNameRequestRow> rows)
        {
            if (rows == null || rows.Count == 0)
            {
                return new List<DeepSeekNameResult>();
            }

            return ParseNameResponse(SendRequest(BuildNameRequestJson(rows)));
        }

        public DeepSeekMappingSelection SelectMappingBox(ExcelQuantityItem item, List<AiMappingCandidate> candidates)
        {
            if (item == null || candidates == null || candidates.Count == 0)
            {
                return null;
            }

            return ParseMappingResponse(SendRequest(BuildMappingRequestJson(item, candidates)));
        }

        private string SendRequest(string requestJson)
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
            request.Timeout = settings.TimeoutSeconds * 1000;
            request.ReadWriteTimeout = settings.TimeoutSeconds * 1000;

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

        private string BuildRequestJson(List<DeepSeekRequestRow> rows)
        {
            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["model"] = settings.Model;
            payload["stream"] = false;
            payload["temperature"] = 0.1;
            payload["max_tokens"] = 1200;
            payload["response_format"] = new Dictionary<string, object> { { "type", "json_object" } };

            List<object> messages = new List<object>();
            messages.Add(new Dictionary<string, object>
            {
                { "role", "system" },
                { "content", "你是铁路工程定额推荐助手。只能从用户给出的本地候选定额中选择，不能编造编号，不能选择材料。必须输出严格 json，格式为 {\"results\":[{\"row_id\":\"r0\",\"selected_code\":\"PY-738\",\"confidence\":90,\"reason\":\"简短理由\"}]}。如果没有可靠候选，selected_code 置空，confidence 置 0。" }
            });
            messages.Add(new Dictionary<string, object>
            {
                { "role", "user" },
                { "content", BuildUserPrompt(rows) }
            });
            payload["messages"] = messages;
            return serializer.Serialize(payload);
        }

        private string BuildNameRequestJson(List<DeepSeekNameRequestRow> rows)
        {
            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["model"] = settings.Model;
            payload["stream"] = false;
            payload["temperature"] = 0.1;
            payload["max_tokens"] = 1400;
            payload["response_format"] = new Dictionary<string, object> { { "type", "json_object" } };

            List<object> messages = new List<object>();
            messages.Add(new Dictionary<string, object>
            {
                { "role", "system" },
                { "content", "You normalize Chinese railway construction quantity names. Return strict JSON: {\"results\":[{\"row_id\":\"n0\",\"quantity_name\":\"标准工程量名称\",\"confidence\":90,\"reason\":\"short reason\"}]}. The quantity_name must be concise, contain the engineering object and key specification, and must not include serial numbers, units, quantities, formulas, or prices." }
            });
            messages.Add(new Dictionary<string, object>
            {
                { "role", "user" },
                { "content", BuildNameUserPrompt(rows) }
            });
            payload["messages"] = messages;
            return serializer.Serialize(payload);
        }

        private string BuildNameUserPrompt(List<DeepSeekNameRequestRow> rows)
        {
            List<object> requestRows = new List<object>();
            foreach (DeepSeekNameRequestRow row in rows)
            {
                ExcelQuantityItem item = row.Item;
                requestRows.Add(new Dictionary<string, object>
                {
                    { "row_id", row.RowId ?? "" },
                    { "local_name", item == null ? "" : item.OriginalName ?? item.Name ?? "" },
                    { "section_name", item == null ? "" : item.SectionName ?? "" },
                    { "unit", item == null ? "" : item.Unit ?? "" },
                    { "quantity_value", item == null ? "" : item.ValueText ?? "" },
                    { "raw_row_text", item == null ? "" : Truncate(item.RawRowText, 300) },
                    { "context_text", item == null ? "" : Truncate(item.ContextText, 300) }
                });
            }

            Dictionary<string, object> body = new Dictionary<string, object>();
            body["task"] = "For each row, summarize all possible quantity-name columns into one standard Chinese engineering quantity name.";
            body["rules"] = new string[]
            {
                "Use all name-like columns before the unit/quantity as context.",
                "Do not include unit, quantity, formula, serial number, or price.",
                "Keep key material/model/specification when it changes the quota selection.",
                "If the row is ambiguous, return the best concise name with lower confidence."
            };
            body["rows"] = requestRows;
            return serializer.Serialize(body);
        }

        private string BuildMappingRequestJson(ExcelQuantityItem item, List<AiMappingCandidate> candidates)
        {
            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["model"] = settings.Model;
            payload["stream"] = false;
            payload["temperature"] = 0.1;
            payload["max_tokens"] = 1000;
            payload["response_format"] = new Dictionary<string, object> { { "type", "json_object" } };

            List<object> messages = new List<object>();
            messages.Add(new Dictionary<string, object>
            {
                { "role", "system" },
                { "content", "You match a normalized railway engineering quantity name to an existing human-corrected mapping box. Only choose from candidates. Return strict JSON: {\"box_id\":\"box-123\",\"confidence\":90,\"reason\":\"short reason\"}. If no candidate is reliable, return empty box_id and confidence 0." }
            });
            messages.Add(new Dictionary<string, object>
            {
                { "role", "user" },
                { "content", BuildMappingUserPrompt(item, candidates) }
            });
            payload["messages"] = messages;
            return serializer.Serialize(payload);
        }

        private string BuildMappingUserPrompt(ExcelQuantityItem item, List<AiMappingCandidate> candidates)
        {
            List<object> candidateRows = new List<object>();
            foreach (AiMappingCandidate candidate in candidates)
            {
                candidateRows.Add(new Dictionary<string, object>
                {
                    { "box_id", candidate.BoxId ?? "" },
                    { "sample_quantity_names", Truncate(candidate.SampleNames, 260) },
                    { "targets", String.Join(" + ", (candidate.Targets ?? new List<MappingTarget>()).Select(t => (t.Code ?? "") + " " + (t.Name ?? "")).ToArray()) },
                    { "local_score", candidate.LocalScore }
                });
            }

            Dictionary<string, object> body = new Dictionary<string, object>();
            body["quantity_name"] = item.Name ?? "";
            body["original_name"] = item.OriginalName ?? "";
            body["unit"] = item.Unit ?? "";
            body["raw_row_text"] = Truncate(item.RawRowText, 300);
            body["rules"] = new string[]
            {
                "Prefer human-corrected mapping boxes when the engineering meaning is equivalent.",
                "Do not choose a box only because of a single generic word.",
                "Steel and concrete terms must not be confused.",
                "If unsure, return empty box_id."
            };
            body["candidates"] = candidateRows;
            return serializer.Serialize(body);
        }

        private string BuildUserPrompt(List<DeepSeekRequestRow> rows)
        {
            List<object> requestRows = new List<object>();
            foreach (DeepSeekRequestRow row in rows)
            {
                List<object> candidates = new List<object>();
                foreach (AiQuotaCandidate candidate in row.Candidates ?? new List<AiQuotaCandidate>())
                {
                    if (candidate == null || candidate.Quota == null)
                    {
                        continue;
                    }

                    candidates.Add(new Dictionary<string, object>
                    {
                        { "code", candidate.Quota.QuotaCode ?? "" },
                        { "name", candidate.Quota.QuotaName ?? "" },
                        { "unit", candidate.Quota.QuotaUnit ?? "" },
                        { "work_content", Truncate(candidate.Quota.WorkContent, 160) },
                        { "local_score", candidate.LocalScore }
                    });
                }

                requestRows.Add(new Dictionary<string, object>
                {
                    { "row_id", row.RowId ?? "" },
                    { "quantity_name", row.Item == null ? "" : row.Item.Name ?? "" },
                    { "quantity_unit", row.Item == null ? "" : row.Item.Unit ?? "" },
                    { "quantity_value", row.Item == null ? "" : row.Item.ValueText ?? "" },
                    { "candidates", candidates }
                });
            }

            Dictionary<string, object> body = new Dictionary<string, object>();
            body["task"] = "请对每行工程量从 candidates 中选择最合适的一条预算定额，返回 json。";
            body["rules"] = new string[]
            {
                "selected_code 必须完全等于某个候选 code，否则置空。",
                "普通关键词检索只选一条定额。",
                "单位不接近或名称证据不足时置空。",
                "钢筋工程量不要误选钢筋混凝土构件。"
            };
            body["rows"] = requestRows;
            return serializer.Serialize(body);
        }

        private string BuildUnifiedRequestJson(List<DeepSeekRequestRow> rows)
        {
            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["model"] = settings.Model;
            payload["stream"] = false;
            payload["temperature"] = 0.1;
            payload["max_tokens"] = 1400;
            payload["response_format"] = new Dictionary<string, object> { { "type", "json_object" } };

            List<object> messages = new List<object>();
            messages.Add(new Dictionary<string, object>
            {
                { "role", "system" },
                { "content", "You are a railway quota recommendation assistant. For each row, first decide whether an existing human-corrected mapping box matches. If yes, return selected_box_id. If no mapping box is reliable, choose one quota from quota_candidates. Never invent ids. Return strict JSON: {\"results\":[{\"row_id\":\"r0\",\"selected_box_id\":\"box-123\",\"selected_code\":\"\",\"confidence\":90,\"reason\":\"short reason\"}]}." }
            });
            messages.Add(new Dictionary<string, object>
            {
                { "role", "user" },
                { "content", BuildUnifiedUserPrompt(rows) }
            });
            payload["messages"] = messages;
            return serializer.Serialize(payload);
        }

        private string BuildUnifiedUserPrompt(List<DeepSeekRequestRow> rows)
        {
            List<object> requestRows = new List<object>();
            foreach (DeepSeekRequestRow row in rows)
            {
                List<object> quotaCandidates = new List<object>();
                foreach (AiQuotaCandidate candidate in row.Candidates ?? new List<AiQuotaCandidate>())
                {
                    if (candidate == null || candidate.Quota == null)
                    {
                        continue;
                    }

                    quotaCandidates.Add(new Dictionary<string, object>
                    {
                        { "code", candidate.Quota.QuotaCode ?? "" },
                        { "name", candidate.Quota.QuotaName ?? "" },
                        { "unit", candidate.Quota.QuotaUnit ?? "" },
                        { "work_content", Truncate(candidate.Quota.WorkContent, 160) },
                        { "local_score", candidate.LocalScore }
                    });
                }

                List<object> mappingCandidates = new List<object>();
                foreach (AiMappingCandidate candidate in row.MappingCandidates ?? new List<AiMappingCandidate>())
                {
                    if (candidate == null)
                    {
                        continue;
                    }

                    mappingCandidates.Add(new Dictionary<string, object>
                    {
                        { "box_id", candidate.BoxId ?? "" },
                        { "sample_quantity_names", Truncate(candidate.SampleNames, 220) },
                        { "targets", String.Join(" + ", (candidate.Targets ?? new List<MappingTarget>()).Select(t => (t.Code ?? "") + " " + (t.Name ?? "")).ToArray()) },
                        { "local_score", candidate.LocalScore }
                    });
                }

                requestRows.Add(new Dictionary<string, object>
                {
                    { "row_id", row.RowId ?? "" },
                    { "quantity_name", row.Item == null ? "" : row.Item.Name ?? "" },
                    { "original_name", row.Item == null ? "" : row.Item.OriginalName ?? "" },
                    { "quantity_unit", row.Item == null ? "" : row.Item.Unit ?? "" },
                    { "quantity_value", row.Item == null ? "" : row.Item.ValueText ?? "" },
                    { "raw_row_text", row.Item == null ? "" : Truncate(row.Item.RawRowText, 260) },
                    { "mapping_candidates", mappingCandidates },
                    { "quota_candidates", quotaCandidates }
                });
            }

            Dictionary<string, object> body = new Dictionary<string, object>();
            body["task"] = "For each quantity row, first match mapping_candidates; if no reliable mapping box exists, select the best quota from quota_candidates.";
            body["rules"] = new string[]
            {
                "selected_box_id must exactly equal a mapping candidate box_id, otherwise leave it empty.",
                "selected_code must exactly equal a quota candidate code, otherwise leave it empty.",
                "Prefer a reliable human-corrected mapping box over a single quota.",
                "Do not choose a mapping box from generic one-word similarity only.",
                "Do not confuse steel quantities with concrete structure quotas."
            };
            body["rows"] = requestRows;
            return serializer.Serialize(body);
        }

        private List<DeepSeekSelection> ParseResponse(string responseJson)
        {
            object rootObject = serializer.DeserializeObject(responseJson);
            Dictionary<string, object> root = rootObject as Dictionary<string, object>;
            if (root == null)
            {
                return new List<DeepSeekSelection>();
            }

            List<object> choices = GetList(root, "choices");
            if (choices == null || choices.Count == 0)
            {
                return new List<DeepSeekSelection>();
            }

            Dictionary<string, object> firstChoice = choices[0] as Dictionary<string, object>;
            Dictionary<string, object> message = firstChoice == null ? null : firstChoice.ContainsKey("message") ? firstChoice["message"] as Dictionary<string, object> : null;
            string content = message == null || !message.ContainsKey("content") ? "" : Convert.ToString(message["content"], CultureInfo.InvariantCulture);
            if (String.IsNullOrWhiteSpace(content))
            {
                return new List<DeepSeekSelection>();
            }

            Dictionary<string, object> resultRoot = serializer.DeserializeObject(content) as Dictionary<string, object>;
            if (resultRoot == null)
            {
                return new List<DeepSeekSelection>();
            }

            List<object> results = GetList(resultRoot, "results");
            if (results == null && resultRoot.ContainsKey("row_id"))
            {
                results = new List<object> { resultRoot };
            }

            List<DeepSeekSelection> selections = new List<DeepSeekSelection>();
            foreach (object item in results ?? new List<object>())
            {
                Dictionary<string, object> row = item as Dictionary<string, object>;
                if (row == null)
                {
                    continue;
                }

                selections.Add(new DeepSeekSelection
                {
                    RowId = ReadString(row, "row_id"),
                    BoxId = String.IsNullOrWhiteSpace(ReadString(row, "selected_box_id")) ? ReadString(row, "box_id") : ReadString(row, "selected_box_id"),
                    SelectedCode = ReadString(row, "selected_code"),
                    Confidence = ReadInt(row, "confidence"),
                    Reason = ReadString(row, "reason")
                });
            }

            return selections;
        }

        private DeepSeekMappingSelection ParseMappingResponse(string responseJson)
        {
            object rootObject = serializer.DeserializeObject(responseJson);
            Dictionary<string, object> root = rootObject as Dictionary<string, object>;
            if (root == null)
            {
                return null;
            }

            List<object> choices = GetList(root, "choices");
            if (choices == null || choices.Count == 0)
            {
                return null;
            }

            Dictionary<string, object> firstChoice = choices[0] as Dictionary<string, object>;
            Dictionary<string, object> message = firstChoice == null || !firstChoice.ContainsKey("message") ? null : firstChoice["message"] as Dictionary<string, object>;
            string content = message == null || !message.ContainsKey("content") ? "" : Convert.ToString(message["content"], CultureInfo.InvariantCulture);
            if (String.IsNullOrWhiteSpace(content))
            {
                return null;
            }

            Dictionary<string, object> resultRoot = serializer.DeserializeObject(content) as Dictionary<string, object>;
            if (resultRoot == null)
            {
                return null;
            }

            return new DeepSeekMappingSelection
            {
                BoxId = ReadString(resultRoot, "box_id"),
                Confidence = ReadInt(resultRoot, "confidence"),
                Reason = ReadString(resultRoot, "reason")
            };
        }

        private List<DeepSeekNameResult> ParseNameResponse(string responseJson)
        {
            object rootObject = serializer.DeserializeObject(responseJson);
            Dictionary<string, object> root = rootObject as Dictionary<string, object>;
            if (root == null)
            {
                return new List<DeepSeekNameResult>();
            }

            List<object> choices = GetList(root, "choices");
            if (choices == null || choices.Count == 0)
            {
                return new List<DeepSeekNameResult>();
            }

            Dictionary<string, object> firstChoice = choices[0] as Dictionary<string, object>;
            Dictionary<string, object> message = firstChoice == null || !firstChoice.ContainsKey("message") ? null : firstChoice["message"] as Dictionary<string, object>;
            string content = message == null || !message.ContainsKey("content") ? "" : Convert.ToString(message["content"], CultureInfo.InvariantCulture);
            if (String.IsNullOrWhiteSpace(content))
            {
                return new List<DeepSeekNameResult>();
            }

            Dictionary<string, object> resultRoot = serializer.DeserializeObject(content) as Dictionary<string, object>;
            if (resultRoot == null)
            {
                return new List<DeepSeekNameResult>();
            }

            List<object> results = GetList(resultRoot, "results");
            if (results == null && resultRoot.ContainsKey("row_id"))
            {
                results = new List<object> { resultRoot };
            }

            List<DeepSeekNameResult> normalized = new List<DeepSeekNameResult>();
            foreach (object item in results ?? new List<object>())
            {
                Dictionary<string, object> row = item as Dictionary<string, object>;
                if (row == null)
                {
                    continue;
                }

                normalized.Add(new DeepSeekNameResult
                {
                    RowId = ReadString(row, "row_id"),
                    QuantityName = ReadString(row, "quantity_name"),
                    Confidence = ReadInt(row, "confidence"),
                    Reason = ReadString(row, "reason")
                });
            }

            return normalized;
        }

        private static List<object> GetList(Dictionary<string, object> values, string key)
        {
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
            if (objectArray != null)
            {
                return objectArray.ToList();
            }

            return null;
        }

        private static string ReadString(Dictionary<string, object> values, string key)
        {
            object value;
            return values.TryGetValue(key, out value) && value != null ? Convert.ToString(value, CultureInfo.InvariantCulture) : "";
        }

        private static int ReadInt(Dictionary<string, object> values, string key)
        {
            object value;
            if (!values.TryGetValue(key, out value) || value == null)
            {
                return 0;
            }

            int parsed;
            return Int32.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) ? parsed : 0;
        }

        private static string Truncate(string text, int maxLength)
        {
            string value = (text ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }
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

        public List<AiQuotaCandidate> BuildDeepSeekCandidates(ExcelQuantityItem item, string categoryFilter, int limit)
        {
            if (item == null)
            {
                return new List<AiQuotaCandidate>();
            }

            int max = Math.Max(1, limit);
            return GetQuotaCandidates(item, categoryFilter)
                .Select(q => new AiQuotaCandidate { Quota = q, LocalScore = ScoreQuota(item, q) })
                .Where(c => c.LocalScore > 0)
                .OrderByDescending(c => c.LocalScore)
                .ThenBy(c => c.Quota.SortOrder)
                .Take(max)
                .ToList();
        }

        public AiQuotaCandidate BuildDeepSeekCandidateByCode(ExcelQuantityItem item, string code, string categoryFilter)
        {
            if (item == null || String.IsNullOrWhiteSpace(code))
            {
                return null;
            }

            IndexQuota quota;
            if (!quotasByCode.TryGetValue(code.Trim(), out quota) || !CategoryAllowed(quota.BookCategory, categoryFilter))
            {
                return null;
            }

            return new AiQuotaCandidate { Quota = quota, LocalScore = ScoreQuota(item, quota) };
        }

        public List<AiQuotaCandidate> BuildDeepSeekCandidatesFromRecommendations(ExcelQuantityItem item, List<RecommendationRow> rows, string categoryFilter)
        {
            List<AiQuotaCandidate> candidates = new List<AiQuotaCandidate>();
            foreach (RecommendationRow row in rows ?? new List<RecommendationRow>())
            {
                if (row == null || !String.Equals(String.IsNullOrWhiteSpace(row.TargetKind) ? QuotaEntry.GuessKind(row.QuotaCode) : row.TargetKind, "quota", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                AiQuotaCandidate candidate = BuildDeepSeekCandidateByCode(item, row.QuotaCode, categoryFilter);
                if (candidate != null)
                {
                    candidate.LocalScore = Math.Max(candidate.LocalScore, row.Score);
                    candidates.Add(candidate);
                }
            }

            return candidates;
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

        public List<AiMappingCandidate> BuildDeepSeekCandidates(ExcelQuantityItem item, string categoryFilter, SearchIndexStore searchIndex, int limit)
        {
            if (item == null)
            {
                return new List<AiMappingCandidate>();
            }

            return boxes
                .Select(box => new
                {
                    Box = box,
                    Targets = FilterTargetsByCategory(box.Targets, categoryFilter, searchIndex),
                    Score = Math.Max(box.Score(item), box.LooseScore(item))
                })
                .Where(x => x.Targets.Count > 0 && x.Score >= 20)
                .OrderByDescending(x => x.Score)
                .Take(Math.Max(1, limit))
                .Select(x => new AiMappingCandidate
                {
                    BoxId = x.Box.BoxId,
                    LocalScore = x.Score,
                    SampleNames = x.Box.SampleNamesForPrompt(),
                    Targets = x.Targets
                })
                .ToList();
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
            List<MappingBox> parsed = null;
            WithMappingBoxesLock(delegate
            {
                parsed = ParseFile(path);
            });
            boxes.Clear();
            boxes.AddRange(CanonicalizeBoxes(parsed));
        }

        private static List<MappingBox> ParseFile(string filePath)
        {
            List<MappingBox> result = new List<MappingBox>();
            if (!File.Exists(filePath))
            {
                return result;
            }

            Dictionary<string, MappingBox> byId = new Dictionary<string, MappingBox>(StringComparer.OrdinalIgnoreCase);
            foreach (string line in File.ReadAllLines(filePath, Encoding.UTF8))
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
                    result.Add(box);
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

            return result;
        }

        // 旧版 box_id 由 String.GetHashCode 生成，跨进程位数不稳定且可能碰撞；
        // 加载时按目标组合重算稳定 ID，同一组合的旧框自动合并，实现旧文件无感迁移。
        private static List<MappingBox> CanonicalizeBoxes(List<MappingBox> parsed)
        {
            List<MappingBox> result = new List<MappingBox>();
            Dictionary<string, MappingBox> byId = new Dictionary<string, MappingBox>(StringComparer.OrdinalIgnoreCase);
            foreach (MappingBox box in parsed ?? new List<MappingBox>())
            {
                if (box.Targets.Count == 0)
                {
                    continue;
                }

                string canonicalId = BuildBoxId(box.Targets);
                MappingBox existing;
                if (!byId.TryGetValue(canonicalId, out existing))
                {
                    box.BoxId = canonicalId;
                    byId[canonicalId] = box;
                    result.Add(box);
                    continue;
                }

                MergeBox(existing, box);
            }

            return result;
        }

        private static void MergeBox(MappingBox into, MappingBox from)
        {
            foreach (MappingTarget target in from.Targets)
            {
                if (!into.Targets.Any(t => String.Equals(t.TargetKey, target.TargetKey, StringComparison.OrdinalIgnoreCase)))
                {
                    into.Targets.Add(target);
                }
            }

            foreach (MappingSample sample in from.Samples)
            {
                MappingSample existing = into.FindSample(sample.QuantityName, sample.QuantityUnit);
                if (existing == null)
                {
                    into.Samples.Add(sample);
                }
                else if (String.Compare(sample.LastUsedAt ?? "", existing.LastUsedAt ?? "", StringComparison.Ordinal) > 0)
                {
                    existing.Weight = sample.Weight;
                    existing.AcceptedCount = sample.AcceptedCount;
                    existing.CorrectedCount = sample.CorrectedCount;
                    existing.RejectedCount = sample.RejectedCount;
                    existing.LastUsedAt = sample.LastUsedAt;
                }
            }
        }

        private void Save()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            WithMappingBoxesLock(delegate
            {
                MergeFromDisk();
                WriteFile();
            });
        }

        // Excel联动AI匹配和扶正训练器也会写这个文件；整文件重写前先合并磁盘上的新增记录，避免覆盖丢失。
        private void MergeFromDisk()
        {
            foreach (MappingBox diskBox in CanonicalizeBoxes(ParseFile(path)))
            {
                MappingBox memory = boxes.FirstOrDefault(b => String.Equals(b.BoxId, diskBox.BoxId, StringComparison.OrdinalIgnoreCase));
                if (memory == null)
                {
                    boxes.Add(diskBox);
                }
                else
                {
                    MergeBox(memory, diskBox);
                }
            }
        }

        private void WriteFile()
        {
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

        private const string MappingBoxesMutexName = "RecoQuotaData.mapping-boxes.lock";

        // 学习库有多个写入方（本窗口扶正、RecoExpandPanel 的 Excel联动与训练器），
        // 用跨程序集一致的命名互斥锁串行化读改写，名称必须与 RecoExpandPanel 保持一致。
        private static void WithMappingBoxesLock(Action action)
        {
            Mutex mutex = new Mutex(false, MappingBoxesMutexName);
            bool acquired = false;
            try
            {
                try
                {
                    acquired = mutex.WaitOne(5000);
                }
                catch (AbandonedMutexException)
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

        // 与 RecoExpandPanel 的 BuildStableMappingBoxId 使用同一套规则：对小写化目标键做 SHA1。
        private static string BuildBoxId(List<MappingTarget> targets)
        {
            string raw = String.Join("|", targets
                .OrderBy(t => t.TargetKey, StringComparer.OrdinalIgnoreCase)
                .Select(t => t.TargetKey)
                .ToArray());
            using (System.Security.Cryptography.SHA1 sha = System.Security.Cryptography.SHA1.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(raw.ToLowerInvariant()));
                StringBuilder builder = new StringBuilder("box-");
                for (int i = 0; i < 8; i++)
                {
                    builder.Append(hash[i].ToString("x2", CultureInfo.InvariantCulture));
                }

                return builder.ToString();
            }
        }

        public static int TargetSortRank(string targetKind, string code)
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
                if (!CanUseSampleForItem(item, sample))
                {
                    continue;
                }

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

        public int LooseScore(ExcelQuantityItem item)
        {
            int best = 0;
            string raw = item == null ? "" : item.RawRowText;
            foreach (MappingSample sample in Samples)
            {
                if (!CanUseSampleForItem(item, sample))
                {
                    continue;
                }

                int score = Math.Max(
                    TextMatcher.NamePairScore(item == null ? "" : item.Name, sample.QuantityName),
                    TextMatcher.NamePairScore(raw, sample.QuantityName) / 2);
                if (RecommendDialog.UnitCompatibleForIndex(sample.QuantityUnit, item == null ? "" : item.Unit))
                {
                    score += 8;
                }
                best = Math.Max(best, score + Math.Min(20, sample.Weight / 2));
            }

            return best;
        }

        private bool CanUseSampleForItem(ExcelQuantityItem item, MappingSample sample)
        {
            if (item == null || sample == null)
            {
                return false;
            }

            if (!String.IsNullOrWhiteSpace(item.Unit) &&
                !String.IsNullOrWhiteSpace(sample.QuantityUnit) &&
                !RecommendDialog.UnitCompatibleForIndex(sample.QuantityUnit, item.Unit))
            {
                return false;
            }

            string targetText = String.Join(" ", Targets.Select(t => (t == null ? "" : (t.Code ?? "") + " " + (t.Name ?? "") + " " + (t.Unit ?? ""))).ToArray());
            return !HasEngineeringProcessConflict(item.Name + " " + item.RawRowText, sample.QuantityName + " " + targetText);
        }

        private static bool HasEngineeringProcessConflict(string quantityText, string candidateText)
        {
            string q = TextMatcher.Normalize(quantityText);
            string c = TextMatcher.Normalize(candidateText);
            bool qSheetPileRemoval = ContainsAny(q, "\u62c9\u68ee", "\u94a2\u677f\u6869") && ContainsAny(q, "\u62d4\u9664", "\u62c6\u9664", "\u62d4\u51fa");
            bool qHasGrout = ContainsAny(q, "\u6ce8\u6d46", "\u6c34\u6ce5\u6d46", "\u586b\u5145", "\u5145\u586b", "\u56de\u586b");
            bool cHasGrout = ContainsAny(c, "\u6ce8\u6d46", "\u6c34\u6ce5\u6d46", "\u586b\u5145", "\u5145\u586b");
            if (qSheetPileRemoval && !qHasGrout && cHasGrout)
            {
                return true;
            }

            bool qConcreteOnly = TextMatcher.IsConcreteQuantityName(q) && !TextMatcher.IsSteelQuantityName(q);
            bool cSteelOnly = TextMatcher.IsSteelQuantityName(c) && !TextMatcher.IsConcreteQuantityName(c);
            if (qConcreteOnly && cSteelOnly)
            {
                return true;
            }

            return false;
        }

        private static bool ContainsAny(string text, params string[] keywords)
        {
            foreach (string keyword in keywords ?? new string[0])
            {
                if (!String.IsNullOrWhiteSpace(keyword) && (text ?? "").Contains(TextMatcher.Normalize(keyword)))
                {
                    return true;
                }
            }
            return false;
        }

        public string SampleNamesForPrompt()
        {
            return String.Join("；", Samples
                .OrderByDescending(s => s.Weight)
                .ThenByDescending(s => s.LastUsedAt ?? "")
                .Take(8)
                .Select(s => s.QuantityName)
                .Where(n => !String.IsNullOrWhiteSpace(n))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray());
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
