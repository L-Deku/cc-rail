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
        private readonly ChapterLibraryStore chapterLibrary;
        private readonly Label entryScopeLabel;
        private readonly ToolTip entryScopeTip = new ToolTip();
        private DeepSeekSettings deepSeekSettings;
        private readonly List<RecommendationRow> recommendations = new List<RecommendationRow>();
        private ExcelSelection currentSelection;
        private int aiRequestVersion;
        // AI 名称/列结构后台识别的版本号：新一轮刷新会使旧的后台任务结果作废
        private int namePrepVersion;
        private EntryScope currentEntryScope;
        private string lastScopeKeyUsed;
        // 项目数据库 → 是否按本库编制办法编制（QD/其他办法的项目不做条目过滤）
        private readonly Dictionary<string, string> projectMethodNoCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public RecommendDialog(Form owner, string initialQuery)
        {
            mainForm = owner;
            Text = "\u6279\u91cf\u63a8\u8350\u5b9a\u989d";
            StartPosition = FormStartPosition.CenterParent;
            Width = 1280;
            Height = 680;
            MinimizeBox = false;

            records = LearningStore.Load();
            LearningStore.BackupLearningFileIfNeeded();
            searchIndex = SearchIndexStore.LoadOrBuild();
            mappingStore = MappingStore.Load(records);
            chapterLibrary = ChapterLibraryStore.Load();
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

            // 条目信息标签随窗口拉宽而变宽（Anchor 含 Right），并挂 ToolTip 以便窄窗时悬停查看完整内容
            entryScopeLabel = new Label();
            entryScopeLabel.Left = 894;
            entryScopeLabel.Top = 15;
            entryScopeLabel.Width = 162;
            entryScopeLabel.Height = 17;
            entryScopeLabel.AutoEllipsis = true;
            entryScopeLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            entryScopeLabel.Text = "";

            Button refreshEntryButton = new Button();
            refreshEntryButton.Text = "刷新条目";
            refreshEntryButton.Width = 66;
            refreshEntryButton.Left = 1264 - 12 - refreshEntryButton.Width; // 贴右
            refreshEntryButton.Top = 10;
            refreshEntryButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            refreshEntryButton.Click += delegate
            {
                if (currentSelection != null)
                {
                    FillRecommendations(currentSelection);
                }
                else
                {
                    RefreshEntryScope();
                }
            };

            resultGrid = new DataGridView();
            resultGrid.Left = 12;
            resultGrid.Top = 48;
            resultGrid.Width = 1240;
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
            statusLabel.Width = 1240;
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
            Controls.Add(entryScopeLabel);
            Controls.Add(refreshEntryButton);
            Controls.Add(resultGrid);
            Controls.Add(statusLabel);

            RefreshEntryScope();
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

        // 每次推荐前重新识别当前条目：对话框是非模态复用的，用户随时会在主程序里切换条目
        internal void RefreshEntryScope()
        {
            currentEntryScope = null;
            try
            {
                if (chapterLibrary != null && !chapterLibrary.IsEmpty)
                {
                    System.Data.SqlClient.SqlConnection conn = GetField<System.Data.SqlClient.SqlConnection>(mainForm, "m_ProjectConn");
                    if (conn != null)
                    {
                        string methodNo = ProjectMethodNo(conn);
                        string projectEntryName;
                        string projectEntryCode = ResolveCurrentChapterNo(conn, out projectEntryName);
                        if (!String.IsNullOrEmpty(projectEntryCode))
                        {
                            currentEntryScope = chapterLibrary.ResolveScope(methodNo, projectEntryCode, projectEntryName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("RefreshEntryScope failed: " + ex.Message);
                currentEntryScope = null;
            }

            UpdateEntryScopeLabel();
        }

        // 当前项目是否按本库的编制办法编制（QD 清单变体等其他办法 ⇒ 不做条目过滤）
        private string ProjectMethodNo(System.Data.SqlClient.SqlConnection conn)
        {
            string dbName;
            try
            {
                dbName = conn.Database ?? "";
            }
            catch
            {
                return "";
            }

            if (String.IsNullOrWhiteSpace(dbName))
            {
                return "";
            }

            string cached;
            if (projectMethodNoCache.TryGetValue(dbName, out cached))
            {
                return cached;
            }

            string methodNo = "";
            try
            {
                EnsureConnectionOpen(conn);
                using (System.Data.SqlClient.SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "select 编制办法文号 from 项目信息";
                    object result = cmd.ExecuteScalar();
                    methodNo = result == null || result == DBNull.Value ? "" : Convert.ToString(result, CultureInfo.InvariantCulture).Trim();
                }
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("Project methodNo query failed: " + ex.Message);
            }

            projectMethodNoCache[dbName] = methodNo;
            return methodNo;
        }

        private static string NormalizeMethodNo(string text)
        {
            return (text ?? "").Replace('—', '-').Replace('–', '-').Replace('－', '-').Replace(" ", "").Trim();
        }

        // 移植自 RecoExpandPanel MultiplierFeature.ResolveChapterNo：定额行的条目序号优先，再走属性页/树节点。
        // 同时带回条目名称，用于识别用户复制条目的来源。
        private string ResolveCurrentChapterNo(System.Data.SqlClient.SqlConnection conn, out string entryName)
        {
            entryName = "";
            DataGridView quotaGrid = GetField<DataGridView>(mainForm, "dataGridViewDE");
            if (quotaGrid != null && quotaGrid.CurrentRow != null && !quotaGrid.CurrentRow.IsNewRow)
            {
                string seq = GetRowValue(quotaGrid.CurrentRow, "条目序号");
                string fromSeq = LookupChapterNoBySeq(conn, seq, out entryName);
                if (!String.IsNullOrEmpty(fromSeq))
                {
                    return fromSeq;
                }
            }

            string fromPropGrid = ReadPropertyGridValue("条目编号");
            if (!String.IsNullOrEmpty(fromPropGrid))
            {
                entryName = ReadPropertyGridValue("工程或费用项目名称") ?? "";
                return fromPropGrid;
            }

            TreeView tree = GetField<TreeView>(mainForm, "Tv_tree");
            TreeNode node = tree != null ? tree.SelectedNode : GetField<TreeNode>(mainForm, "CurrNode");
            if (node != null)
            {
                string fromTag = TryGetTagValue(node.Tag, "条目编号");
                if (!String.IsNullOrEmpty(fromTag))
                {
                    entryName = node.Text ?? "";
                    return fromTag;
                }

                string seq = TryGetTagValue(node.Tag, "条目序号");
                if (String.IsNullOrEmpty(seq) && IsAllDigits(node.Name))
                {
                    seq = node.Name;
                }

                string fromSeq = LookupChapterNoBySeq(conn, seq, out entryName);
                if (!String.IsNullOrEmpty(fromSeq))
                {
                    if (String.IsNullOrEmpty(entryName))
                    {
                        entryName = node.Text ?? "";
                    }
                    return fromSeq;
                }

                if (!String.IsNullOrEmpty(node.Name) && !IsAllDigits(node.Name))
                {
                    entryName = node.Text ?? "";
                    return node.Name;
                }
            }

            return null;
        }

        private static string LookupChapterNoBySeq(System.Data.SqlClient.SqlConnection conn, string seq, out string entryName)
        {
            entryName = "";
            if (String.IsNullOrWhiteSpace(seq) || conn == null)
            {
                return null;
            }

            EnsureConnectionOpen(conn);
            using (System.Data.SqlClient.SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 条目编号, 工程或费用项目名称 from 章节表 where 条目序号=@id";
                cmd.Parameters.AddWithValue("@id", seq.Trim());
                using (System.Data.SqlClient.SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                    {
                        return null;
                    }

                    string code = reader.IsDBNull(0) ? null : Convert.ToString(reader.GetValue(0), CultureInfo.InvariantCulture).Trim();
                    entryName = reader.IsDBNull(1) ? "" : Convert.ToString(reader.GetValue(1), CultureInfo.InvariantCulture).Trim();
                    return code;
                }
            }
        }

        private static void EnsureConnectionOpen(System.Data.SqlClient.SqlConnection conn)
        {
            if (conn != null && conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
        }

        private string ReadPropertyGridValue(string propertyName)
        {
            DataGridView propGrid = GetField<DataGridView>(mainForm, "dataGridViewProp");
            if (propGrid == null)
            {
                return null;
            }

            foreach (DataGridViewRow row in propGrid.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

                if (String.Equals(GetRowValue(row, "属性名称"), propertyName, StringComparison.Ordinal))
                {
                    return GetRowValue(row, "数据");
                }
                if (row.Cells.Count >= 2)
                {
                    object nameValue = row.Cells[0].Value;
                    if (nameValue != null && String.Equals(Convert.ToString(nameValue).Trim(), propertyName, StringComparison.Ordinal))
                    {
                        object dataValue = row.Cells[1].Value;
                        return dataValue == null ? null : Convert.ToString(dataValue).Trim();
                    }
                }
            }

            return null;
        }

        private static string TryGetTagValue(object source, string name)
        {
            if (source == null)
            {
                return null;
            }

            DataRowView rowView = source as DataRowView;
            if (rowView != null && rowView.DataView.Table.Columns.Contains(name))
            {
                return Convert.ToString(rowView[name], CultureInfo.InvariantCulture);
            }

            DataRow dataRow = source as DataRow;
            if (dataRow != null && dataRow.Table.Columns.Contains(name))
            {
                return Convert.ToString(dataRow[name], CultureInfo.InvariantCulture);
            }

            PropertyInfo prop = source.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
            if (prop != null)
            {
                object value = prop.GetValue(source, null);
                return value == null ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
            }

            return null;
        }

        private static bool IsAllDigits(string text)
        {
            return !String.IsNullOrEmpty(text) && text.All(Char.IsDigit);
        }

        private void UpdateEntryScopeLabel()
        {
            if (entryScopeLabel == null)
            {
                return;
            }

            if (chapterLibrary == null || chapterLibrary.IsEmpty)
            {
                entryScopeLabel.Text = "条目库未启用";
                entryScopeLabel.ForeColor = SystemColors.GrayText;
                entryScopeTip.SetToolTip(entryScopeLabel, "未找到章节条目库（chapter-entries.jsonl），按全库推荐。");
                return;
            }

            if (currentEntryScope != null && currentEntryScope.Strict)
            {
                string text = "条目:" + currentEntryScope.MatchedEntryCode + " " + (currentEntryScope.EntryName ?? "")
                    + "｜池" + currentEntryScope.PoolKeys.Count.ToString(CultureInfo.InvariantCulture) + "条 严格";
                entryScopeLabel.Text = text;
                entryScopeLabel.ForeColor = Color.FromArgb(46, 96, 49);
                string tip = text;
                if (!String.Equals(currentEntryScope.ProjectEntryCode, currentEntryScope.MatchedEntryCode, StringComparison.Ordinal))
                {
                    tip += "\r\n（当前条目 " + currentEntryScope.ProjectEntryCode + " 为新建/复制条目，采用来源条目 " + currentEntryScope.MatchedEntryCode + " 的定额池）";
                }
                entryScopeTip.SetToolTip(entryScopeLabel, tip);
            }
            else
            {
                entryScopeLabel.Text = "条目:未识别（全库推荐）";
                entryScopeLabel.ForeColor = SystemColors.GrayText;
                entryScopeTip.SetToolTip(entryScopeLabel, "未识别到当前定额行所属的小计/指标条目，按全库推荐。可在主程序选中定额行后点“刷新条目”。");
            }
        }

        private Dictionary<string, bool> SnapshotCheckStates()
        {
            Dictionary<string, bool> states = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < resultGrid.Rows.Count && i < recommendations.Count; i++)
            {
                RecommendationRow rec = recommendations[i];
                if (rec == null || String.IsNullOrWhiteSpace(rec.QuotaCode))
                {
                    continue;
                }

                object value = resultGrid.Rows[i].Cells["Checked"].Value;
                states[CheckStateKey(rec)] = value is bool && (bool)value;
            }

            return states;
        }

        private static string CheckStateKey(RecommendationRow row)
        {
            if (row == null || row.Item == null)
            {
                return "";
            }

            string nameForKey = String.IsNullOrWhiteSpace(row.Item.OriginalName) ? row.Item.Name : row.Item.OriginalName;
            string codeKey = (String.IsNullOrWhiteSpace(row.TargetKind) ? QuotaEntry.GuessKind(row.QuotaCode) : row.TargetKind)
                + ":" + (row.QuotaCode ?? "").Trim().ToUpperInvariant();
            return LearningStore.BuildQuantitySignature(nameForKey, row.Item.Unit) + "|" + codeKey;
        }

        private void FillRecommendations(ExcelSelection selection, bool normalizeNames)
        {
            currentSelection = selection;
            int prepVersion = ++namePrepVersion;
            if (!normalizeNames)
            {
                RestoreOriginalQuantityNames(selection);
                FillRecommendationsCore(selection);
                return;
            }

            if (!deepSeekSettings.CanDetectColumns && !deepSeekSettings.CanNormalizeNames)
            {
                FillRecommendationsCore(selection);
                return;
            }

            // AI 名称/列结构识别是同步 HTTP（超时可达数十秒），放后台线程执行，避免冻结主界面；
            // 旧结果先保留在表格里，识别完成后再整体重建。prepVersion 防止等待期间用户再次触发刷新时，
            // 过期的识别结果覆盖新一轮内容。
            statusLabel.Text = "DeepSeek正在后台识别工程量名称/列结构...";
            ThreadPool.QueueUserWorkItem(delegate
            {
                try
                {
                    if (prepVersion != namePrepVersion)
                    {
                        return;
                    }

                    // 先让 AI 判断列结构并重建多列名称；命中的条目已标记跳过润色，其余条目仍走名称润色。
                    bool layoutApplied = ApplyAiColumnLayout(selection);
                    NormalizeQuantityNamesWithDeepSeek(selection, !layoutApplied);
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("AI name preparation failed: " + ex.Message);
                }

                try
                {
                    if (!IsDisposed && IsHandleCreated)
                    {
                        BeginInvoke((MethodInvoker)delegate
                        {
                            if (prepVersion == namePrepVersion)
                            {
                                FillRecommendationsCore(selection);
                            }
                        });
                    }
                }
                catch
                {
                }
            });
        }

        private void FillRecommendationsCore(ExcelSelection selection)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            currentSelection = selection;
            RefreshEntryScope();
            EntryScope scope = currentEntryScope;
            lastScopeKeyUsed = scope != null && scope.Strict ? scope.Tag : "";
            // 重建前快照各行勾选状态，避免用户手动取消的勾选在刷新后又被自动勾上
            Dictionary<string, bool> priorChecks = SnapshotCheckStates();
            recommendations.Clear();
            resultGrid.Rows.Clear();
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
                    string cacheKey = BuildBatchCacheKey(item, categoryFilter) + "|" + lastScopeKeyUsed;
                    List<RecommendationRow> itemRecommendations;
                    List<RecommendationRow> cached;
                    if (batchCache.TryGetValue(cacheKey, out cached))
                    {
                        itemRecommendations = CloneRecommendationsForItem(item, cached);
                        stats.CacheHits++;
                    }
                    else
                    {
                        itemRecommendations = BuildRecommendations(item, categoryFilter, scope, stats);
                        batchCache[cacheKey] = CloneRecommendationsForItem(item, itemRecommendations);
                    }

                    int itemRowIndex = 0;
                    foreach (RecommendationRow recommendation in itemRecommendations)
                    {
                        bool isContinuation = itemRowIndex > 0 && String.Equals(recommendation.Source, "mapping", StringComparison.OrdinalIgnoreCase);
                        recommendations.Add(recommendation);
                        bool defaultChecked = recommendation.Score >= 60 && !String.IsNullOrWhiteSpace(recommendation.QuotaCode);
                        bool priorChecked;
                        bool checkedValue = priorChecks.TryGetValue(CheckStateKey(recommendation), out priorChecked) ? priorChecked : defaultChecked;
                        int gridRowIndex = resultGrid.Rows.Add(
                            checkedValue,
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
                                Scope = scope,
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
            statusLabel.Text += scope != null && scope.Strict
                ? "条目范围：" + scope.MatchedEntryCode + "（池" + scope.PoolKeys.Count.ToString(CultureInfo.InvariantCulture) + "条，严格）。"
                : (chapterLibrary != null && !chapterLibrary.IsEmpty ? "条目范围：全库。" : "");

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

        // AI 列映射兜底：把整张表的原始网格交给 DeepSeek 判断列结构，再用确定性逻辑(BuildItemsFromColumnLayout)
        // 重建名称/单位。仅更新已识别条目（按行号对回），保留 OriginalName 以便关闭 AI 时还原。返回是否有改动。
        private bool ApplyAiColumnLayout(ExcelSelection selection)
        {
            if (selection == null || selection.Items.Count == 0
                || selection.RawRows == null || selection.RawRows.Count == 0
                || !deepSeekSettings.CanDetectColumns)
            {
                return false;
            }

            try
            {
                DeepSeekClient client = new DeepSeekClient(deepSeekSettings);
                DeepSeekColumnLayout layout = client.DetectColumnLayout(selection.RawRows);
                if (layout == null || layout.Confidence < 70 || layout.QuantityColumn <= 0
                    || layout.NameColumns == null || layout.NameColumns.Length == 0)
                {
                    return false;
                }

                List<int> descColumns = layout.NameColumns
                    .Where(c => c > 0 && c != layout.QuantityColumn && c != layout.UnitColumn)
                    .Distinct()
                    .OrderBy(c => c)
                    .ToList();
                if (descColumns.Count == 0)
                {
                    return false;
                }

                List<ExcelQuantityItem> rebuilt = BuildItemsFromColumnLayout(selection.RawRows, descColumns, layout.UnitColumn, layout.QuantityColumn, selection.WorksheetName);
                Dictionary<int, ExcelQuantityItem> byRow = new Dictionary<int, ExcelQuantityItem>();
                foreach (ExcelQuantityItem r in rebuilt)
                {
                    if (r != null)
                    {
                        byRow[r.RowNumber] = r;
                    }
                }

                int changed = 0;
                foreach (ExcelQuantityItem item in selection.Items)
                {
                    ExcelQuantityItem mapped;
                    if (item == null || !byRow.TryGetValue(item.RowNumber, out mapped) || String.IsNullOrWhiteSpace(mapped.Name))
                    {
                        continue;
                    }

                    if (String.IsNullOrWhiteSpace(item.OriginalName))
                    {
                        item.OriginalName = item.Name;
                    }

                    if (String.Equals(item.Name, mapped.Name, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    item.AiName = mapped.Name;
                    item.Name = mapped.Name;
                    if (!String.IsNullOrWhiteSpace(mapped.Unit))
                    {
                        item.Unit = mapped.Unit;
                    }
                    item.SectionName = mapped.SectionName;
                    item.ContextText = item.Name + " " + item.Unit + " " + item.ValueText + " " + item.RawRowText;
                    item.SkipAiNameNormalization = true;
                    changed++;
                }

                QuotaRecommendPanel.Log("AI column layout. nameCols=[" + String.Join(",", descColumns.Select(c => c.ToString(CultureInfo.InvariantCulture)).ToArray()) + "]"
                    + ", unitCol=" + layout.UnitColumn.ToString(CultureInfo.InvariantCulture)
                    + ", qtyCol=" + layout.QuantityColumn.ToString(CultureInfo.InvariantCulture)
                    + ", conf=" + layout.Confidence.ToString(CultureInfo.InvariantCulture)
                    + ", changed=" + changed.ToString(CultureInfo.InvariantCulture));
                return changed > 0;
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("AI column layout detection failed: " + ex.Message);
                return false;
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

        // 工程量名称含“模板/模版”视为模板工程量，不配定额
        private static bool IsFormworkQuantity(string name)
        {
            string text = name ?? "";
            return text.IndexOf("模板", StringComparison.Ordinal) >= 0 || text.IndexOf("模版", StringComparison.Ordinal) >= 0;
        }

        private List<RecommendationRow> BuildRecommendations(ExcelQuantityItem item, string categoryFilter, EntryScope scope, RecommendationBatchStats stats)
        {
            List<AiQuotaCandidate> aiCandidates = new List<AiQuotaCandidate>();
            // 人工扶正过的对应框优先（即便是模板，也按用户指定的定额显示）。
            // Find 内部已要求强名称匹配且加权分 >=70，这里不再与全量索引分数比高低：
            // 短名称（如“钢筋”）的索引分可超过对应框分数上限，旧比较会导致扶正后仍只显示单条索引结果。
            List<RecommendationRow> mapped = mappingStore.Find(item, categoryFilter, searchIndex, scope);
            if (mapped.Count > 0)
            {
                stats.MappingHits++;
                return mapped;
            }

            // 模板类工程量按规则不配定额：保留工程量行，推荐定额留空，且不走 AI 补推
            if (IsFormworkQuantity(item.Name))
            {
                RecommendationRow skip = new RecommendationRow();
                skip.Item = item;
                skip.QuotaCode = "";
                skip.QuotaName = "";
                skip.QuotaUnit = "";
                skip.ConvertedValueText = item.ValueText;
                skip.Score = 0;
                skip.Reason = "模板工程量按规则不推荐定额";
                skip.Source = "skip";
                skip.TargetKind = "quota";
                skip.AiCandidates = new List<AiQuotaCandidate>();
                skip.AiMappingCandidates = new List<AiMappingCandidate>();
                stats.EmptyRows++;
                return new List<RecommendationRow> { skip };
            }

            List<AiMappingCandidate> mappingCandidates = deepSeekSettings.CanDetectMapping
                ? mappingStore.BuildDeepSeekCandidates(item, categoryFilter, searchIndex, deepSeekSettings.MaxCandidatesPerRow, scope)
                : new List<AiMappingCandidate>();
            stats.IndexSearches++;
            foreach (AiQuotaCandidate candidate in searchIndex.BuildDeepSeekCandidates(item, categoryFilter, deepSeekSettings.MaxCandidatesPerRow, scope))
            {
                if (!aiCandidates.Any(c => c != null && c.Quota != null && candidate != null && candidate.Quota != null && String.Equals(c.Quota.QuotaCode, candidate.Quota.QuotaCode, StringComparison.OrdinalIgnoreCase)))
                {
                    aiCandidates.Add(candidate);
                }
            }

            // 严格条目模式：把整条目定额池补进候选，确保池里相关定额都能被本地匹配或 AI 选中（关键词没命中也不漏）
            if (scope != null && scope.Strict)
            {
                foreach (AiQuotaCandidate candidate in searchIndex.BuildScopeCandidates(item, scope, deepSeekSettings.MaxCandidatesPerRow))
                {
                    if (!aiCandidates.Any(c => c != null && c.Quota != null && candidate != null && candidate.Quota != null && String.Equals(c.Quota.QuotaCode, candidate.Quota.QuotaCode, StringComparison.OrdinalIgnoreCase)))
                    {
                        aiCandidates.Add(candidate);
                    }
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
            if (String.Equals(row.Source, "skip", StringComparison.OrdinalIgnoreCase))
            {
                return "\u6a21\u677f\u514d\u63a8";
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
                // 每批控制在较小规模，让单次 AI 请求能在超时时间内返回，减少超时行
                int cost = EstimateDeepSeekRowCost(item);
                if (current.Count > 0 && (current.Count >= 6 || currentCost + cost > 50))
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
                if (row == null)
                {
                    continue;
                }

                // 先返回的批次可能已插入对应框延续行并重排行号，发起请求时快照的 GridRowIndex 会过期；
                // 按行对象在当前列表中的实际位置重新定位，避免结果被静默丢弃（表现为一直“AI补推中”）。
                int gridRowIndex = IndexOfRecommendation(row);
                if (gridRowIndex < 0 || gridRowIndex >= resultGrid.Rows.Count)
                {
                    continue;
                }

                row.AiPending = false;
                DeepSeekSelection selection;
                if (!String.IsNullOrWhiteSpace(error))
                {
                    SetRecommendationStatus(gridRowIndex, DeepSeekFailureStatus(error), error);
                    continue;
                }

                if (!byRow.TryGetValue(row.AiRowId, out selection))
                {
                    SetRecommendationStatus(gridRowIndex, selections == null || selections.Count == 0 ? "AI\u8fd4\u56de\u4e3a\u7a7a" : "AI\u65e0\u7ed3\u679c", "");
                    continue;
                }

                if (!String.IsNullOrWhiteSpace(selection.ErrorText))
                {
                    SetRecommendationStatus(gridRowIndex, DeepSeekFailureStatus(selection.ErrorText), selection.ErrorText);
                    continue;
                }

                AiMappingCandidate mappingCandidate = (pending.Request.MappingCandidates ?? new List<AiMappingCandidate>())
                    .FirstOrDefault(c => c != null && String.Equals(c.BoxId, selection.BoxId, StringComparison.OrdinalIgnoreCase));
                if (mappingCandidate != null && selection.Confidence >= 65)
                {
                    List<RecommendationRow> mappedRows = mappingCandidate.ToRecommendations(row.Item, selection.Confidence, selection.Reason);
                    if (mappedRows.Count > 0)
                    {
                        ApplyMappedRowsFromDeepSeek(gridRowIndex, row, mappedRows);
                        applied += mappedRows.Count;
                        continue;
                    }
                }

                AiQuotaCandidate candidate = (pending.Request.Candidates ?? new List<AiQuotaCandidate>())
                    .FirstOrDefault(c => c != null && c.Quota != null && String.Equals(c.Quota.QuotaCode, selection.SelectedCode, StringComparison.OrdinalIgnoreCase));
                if (candidate == null)
                {
                    SetRecommendationStatus(gridRowIndex, String.IsNullOrWhiteSpace(selection.SelectedCode) ? "AI\u65e0\u7ed3\u679c" : "AI\u8fd4\u56de\u65e0\u6548", String.IsNullOrWhiteSpace(selection.SelectedCode) ? selection.Reason : "\u8fd4\u56de\u7f16\u53f7\u4e0d\u5728\u672c\u5730\u5019\u9009\u4e2d");
                    continue;
                }

                // \u5019\u9009\u751f\u6210\u9636\u6bb5\u5df2\u6309\u6761\u76ee\u8fc7\u6ee4\uff0c\u8fd9\u91cc\u518d\u515c\u5e95\u4e00\u6b21\uff0c\u9632\u6b62 AI \u8fd4\u56de\u6c60\u5916\u5b9a\u989d
                if (pending.Scope != null && pending.Scope.Strict && !pending.Scope.Allows("quota", candidate.Quota.QuotaCode))
                {
                    SetRecommendationStatus(gridRowIndex, "AI\u8d85\u51fa\u6761\u76ee\u8303\u56f4", "AI\u8fd4\u56de\u7684\u5b9a\u989d\u4e0d\u5728\u5f53\u524d\u6761\u76ee\u5b9a\u989d\u6c60\u5185");
                    continue;
                }

                int confidence = Math.Max(0, Math.Min(100, selection.Confidence));
                if (confidence < deepSeekSettings.DisplayConfidence ||
                    (!String.Equals(row.Source, "empty", StringComparison.OrdinalIgnoreCase) &&
                        !String.Equals(row.Source, "mapping", StringComparison.OrdinalIgnoreCase) &&
                        confidence < row.Score))
                {
                    SetRecommendationStatus(gridRowIndex, "\u672c\u5730\u4f18\u5148", selection.Reason);
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

                DataGridViewRow gridRow = resultGrid.Rows[gridRowIndex];
                gridRow.Cells["QuotaCode"].Value = row.QuotaCode;
                gridRow.Cells["QuotaName"].Value = row.QuotaName;
                gridRow.Cells["QuotaUnit"].Value = row.QuotaUnit;
                gridRow.Cells["QuotaQuantity"].Value = row.ConvertedValueText;
                gridRow.Cells["Checked"].Value = confidence >= deepSeekSettings.AutoCheckConfidence &&
                    RecommendDialog.UnitCompatibleForIndex(row.Item.Unit, row.QuotaUnit);
                SetRecommendationStatus(gridRowIndex, "AI\u5df2\u8865\u63a8", row.Reason);
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

        // 按引用定位推荐行的当前网格行号；行不存在（已被刷新或替换）时返回 -1
        private int IndexOfRecommendation(RecommendationRow row)
        {
            for (int i = 0; i < recommendations.Count; i++)
            {
                if (Object.ReferenceEquals(recommendations[i], row))
                {
                    return i;
                }
            }

            return -1;
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
                mappingStore.Correct(recommendation.Item, recommendation, quotas, currentEntryScope);
                if (currentEntryScope != null && currentEntryScope.Strict && chapterLibrary != null)
                {
                    foreach (QuotaEntry quota in quotas)
                    {
                        chapterLibrary.AddUserQuota(currentEntryScope, quota.TargetKind, quota.QuotaCode, quota.QuotaName, quota.QuotaUnit);
                    }
                }
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
            mappingStore.Accept(rows, currentEntryScope);
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

                List<List<CellValue>> rawRows;
                selection.Items.AddRange(BuildQuantityItemsFromRange(range, selection.WorksheetName, out rawRows));
                selection.RawRows = rawRows;
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

                List<List<CellValue>> rawRows;
                selection.Items.AddRange(BuildQuantityItemsFromRange(range, selection.WorksheetName, out rawRows));
                selection.RawRows = rawRows;
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

            List<List<CellValue>> rawRows;
            selection.Items.AddRange(BuildQuantityItemsFromTextTable(textTable, selection.WorksheetName, out rawRows));
            selection.RawRows = rawRows;
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
                List<List<CellValue>> rawRows;
                selection.Items.AddRange(BuildQuantityItemsFromTextTable(table, selection.WorksheetName, out rawRows));
                selection.RawRows = rawRows;

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

        private static List<ExcelQuantityItem> BuildQuantityItemsFromRange(dynamic range, string worksheetName, out List<List<CellValue>> rawRows)
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

            rawRows = rows;
            return BuildQuantityItemsFromCellRows(rows, worksheetName);
        }

        private static List<ExcelQuantityItem> BuildQuantityItemsFromTextTable(List<List<string>> table, string worksheetName, out List<List<CellValue>> rawRows)
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

            rawRows = rows;
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
            int nameBoundary = unitColumn >= 0 && unitColumn < quantityColumn ? unitColumn : quantityColumn;
            List<int> descColumns = FindDescriptionColumns(rows, nameBoundary, unitColumn, quantityColumn);
            if (descColumns.Count == 0)
            {
                return result;
            }

            result = BuildItemsFromColumnLayout(rows, descColumns, unitColumn, quantityColumn, worksheetName);

            QuotaRecommendPanel.Log("Grid parser: rows=" + rows.Count.ToString(CultureInfo.InvariantCulture)
                + ", qtyCol=" + quantityColumn.ToString(CultureInfo.InvariantCulture)
                + ", unitCol=" + unitColumn.ToString(CultureInfo.InvariantCulture)
                + ", descCols=[" + String.Join(",", descColumns.Select(c => c.ToString(CultureInfo.InvariantCulture)).ToArray()) + "]"
                + ", items=" + result.Count.ToString(CultureInfo.InvariantCulture));
            return result;
        }

        // 根据已识别的列布局（多个描述列 + 单位列 + 数量列）逐行构建工程量条目。
        // descColumns 为升序排列的描述列索引，名称由这些列按列序拼接而成，从而支持超过四列的表。
        private static List<ExcelQuantityItem> BuildItemsFromColumnLayout(List<List<CellValue>> rows, List<int> descColumns, int unitColumn, int quantityColumn, string worksheetName)
        {
            List<ExcelQuantityItem> result = new List<ExcelQuantityItem>();
            if (rows == null || descColumns == null || descColumns.Count == 0 || quantityColumn < 0)
            {
                return result;
            }

            int sectionColumn = descColumns[0];
            // 仅当存在两个及以上描述列时，最左侧列才作为"分部/小节"承接（兼容旧的 group 行为）；
            // 只有一个描述列时按行逐条取名、不做承接，与旧的三列表现完全一致。
            bool useCarryDown = descColumns.Count >= 2;
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

                string section = "";
                if (useCarryDown)
                {
                    section = GetCellText(row, sectionColumn);
                    if (LooksLikeGroupText(section))
                    {
                        currentGroup = section;
                    }
                    else
                    {
                        section = currentGroup;
                    }
                }

                List<string> parts = new List<string>();
                foreach (int c in descColumns)
                {
                    string text = useCarryDown && c == sectionColumn ? section : GetCellText(row, c);
                    if (!String.IsNullOrWhiteSpace(text) && !LooksLikeOrderOrHeader(text) && !LooksLikeUnit(text))
                    {
                        parts.Add(text);
                    }
                }

                string name = CombineQuantityNames(parts);
                if (String.IsNullOrWhiteSpace(name) || LooksLikeOrderOrHeader(name))
                {
                    name = PickQuantityName(row, quantityCell, unitColumn >= 0 ? GetCell(row, unitColumn) : null);
                }

                if (String.IsNullOrWhiteSpace(name) || LooksLikeOrderOrHeader(name))
                {
                    continue;
                }

                string group = useCarryDown ? currentGroup : "";
                if (String.Equals(group, name, StringComparison.Ordinal))
                {
                    group = "";
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
                item.SectionName = String.IsNullOrWhiteSpace(group) ? name : group;
                item.OriginalName = name;
                item.RawRowText = BuildRawRowText(row);
                result.Add(item);
            }

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

        // 找出数量列左侧、整列以"组文本"（中文描述、非序号/单位/数量）为主的所有列，作为名称/描述列。
        // 返回升序列索引；序号列因是数字会被 LooksLikeGroupText 自动排除。支持任意数量的描述列。
        private static List<int> FindDescriptionColumns(List<List<CellValue>> rows, int nameBoundary, int unitColumn, int quantityColumn)
        {
            List<int> columns = new List<int>();
            if (rows == null)
            {
                return columns;
            }

            var byColumn = rows.SelectMany(r => r ?? new List<CellValue>())
                .Where(c => c != null
                    && c.SourceIndex < nameBoundary
                    && c.SourceIndex != unitColumn
                    && c.SourceIndex != quantityColumn)
                .GroupBy(c => c.SourceIndex);

            foreach (var columnGroup in byColumn)
            {
                int groupText = 0;
                int nonEmpty = 0;
                foreach (CellValue cell in columnGroup)
                {
                    string text = (cell.Text ?? "").Trim();
                    if (String.IsNullOrEmpty(text))
                    {
                        continue;
                    }

                    nonEmpty++;
                    if (LooksLikeGroupText(text))
                    {
                        groupText++;
                    }
                }

                if (nonEmpty > 0 && groupText > 0 && groupText * 2 >= nonEmpty)
                {
                    columns.Add(columnGroup.Key);
                }
            }

            columns.Sort();
            return columns;
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

        // 把多个描述列文本按列序折叠成一个工程量名称，复用 CombineQuantityName 的子串去重/空格拼接逻辑。
        private static string CombineQuantityNames(IEnumerable<string> parts)
        {
            string name = "";
            foreach (string part in parts ?? new string[0])
            {
                if (String.IsNullOrWhiteSpace(part))
                {
                    continue;
                }

                name = CombineQuantityName(name, part);
            }

            return name;
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

        private sealed class RecommendationBatchStats
        {
            public int MappingHits;
            public int IndexSearches;
            public int EmptyRows;
            public int CacheHits;
            public int AiQueued;
        }

        private struct CarryCell
        {
            public string Text;
            public int RemainingRows;
        }
    }
}
