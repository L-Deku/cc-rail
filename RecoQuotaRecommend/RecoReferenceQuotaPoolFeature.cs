using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace RecoQuotaRecommend
{
    // 把"章节条目 -> 原始定额池"（chapter-quota-library.jsonl）显示到宿主"定额输入界面"
    // 最下面已有的"参考定额"框里：随当前章节条目刷新，双击一条定额即填入定额输入表当前行。
    // 该框控件属于闭源主程序、字段名未知，故安装时先把控件清单写入日志（一次性发现），
    // 再按 TabPage.Text=="参考定额" / 控件 Text 含"参考定额"自动定位容器，叠放自管的只读表格。
    // 约定：本仓库 .cs 无 BOM、按 GBK 编译，中文注释可保留（被忽略），中文字符串字面量一律 \u 转义。
    internal sealed class ReferenceQuotaPoolFeature
    {
        private const int RefreshDelayMs = 200;
        private const string ReferenceTabText = "\u53c2\u8003\u5b9a\u989d"; // 参考定额
        private static readonly Dictionary<Form, Runtime> Runtimes = new Dictionary<Form, Runtime>();

        public static void Install(Form mainForm)
        {
            if (mainForm == null || Runtimes.ContainsKey(mainForm))
            {
                return;
            }

            DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
            if (grid == null)
            {
                QuotaRecommendPanel.Log("Reference quota pool skipped: dataGridViewDE not found.");
                return;
            }

            Runtime runtime = new Runtime(mainForm, grid);
            if (!runtime.Install())
            {
                runtime.Dispose();
                return;
            }

            Runtimes[mainForm] = runtime;
            mainForm.FormClosed += delegate
            {
                Runtime existing;
                if (Runtimes.TryGetValue(mainForm, out existing))
                {
                    existing.Dispose();
                    Runtimes.Remove(mainForm);
                }
            };
            QuotaRecommendPanel.Log("Reference quota pool installed.");
        }

        private sealed class Runtime : IDisposable
        {
            private readonly Form mainForm;
            private readonly DataGridView grid; // 宿主定额输入表 dataGridViewDE
            private readonly Timer timer;
            private readonly ChapterLibraryStore chapterLibrary;
            private readonly Dictionary<string, List<PoolItem>> poolByEntry; // matchedEntryCode -> 富字段定额池
            private readonly Dictionary<string, QuotaInfo> quotaIndex; // 定额编号(大写) -> 名称/单位/基价/工作内容（取自 quota-index.jsonl）
            private readonly Dictionary<string, bool> methodCache = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            private DataGridView refGrid; // 我们自管的只读表格，叠放进"参考定额"框
            private Panel hostPanel;      // 容纳工具条 + refGrid，整体叠放进"参考定额"框
            private DataGridView nativeGrid; // tabPageCommon 里宿主原生 dg_Common，隐藏它以免盖住我们的表
            private Control refContainer;    // "参考定额"页容器，监听其可见性以便切到该页时重新加载数据
            private ToolStripLabel entryLabel; // 工具条右侧显示当前命中的条目编号
            private List<PoolItem> displayedItems = new List<PoolItem>(); // 与 refGrid 行一一对应，供删除取 code/kind
            private TreeView tree;
            private string currentKey;
            private bool disposed;

            public Runtime(Form mainForm, DataGridView grid)
            {
                this.mainForm = mainForm;
                this.grid = grid;
                timer = new Timer();
                timer.Interval = RefreshDelayMs;
                timer.Tick += TimerTick;
                chapterLibrary = ChapterLibraryStore.Load();
                poolByEntry = LoadPool(chapterLibrary == null ? "" : chapterLibrary.MethodKey);
                quotaIndex = LoadQuotaIndex();
            }

            public bool Install()
            {
                if (chapterLibrary == null || chapterLibrary.IsEmpty)
                {
                    QuotaRecommendPanel.Log("Reference quota pool skipped: chapter library empty.");
                    return false;
                }

                // 一次性发现：把窗体内的表格/选项卡清单写日志，便于核对"参考定额"控件
                LogControlInventory(mainForm);

                Control container = FindReferenceContainer(mainForm);
                if (container == null)
                {
                    QuotaRecommendPanel.Log("Reference quota pool skipped: reference container not found (see inventory log).");
                    return false;
                }

                refGrid = BuildRefGrid();

                ToolStrip bar = new ToolStrip();
                bar.GripStyle = ToolStripGripStyle.Hidden;
                bar.Dock = DockStyle.Top;
                ToolStripButton addBtn = new ToolStripButton("\u589e\u52a0\u5b9a\u989d");
                addBtn.DisplayStyle = ToolStripItemDisplayStyle.Text;
                addBtn.ToolTipText = "\u628a\u4e00\u6761\u5b9a\u989d\u52a0\u5165\u5f53\u524d\u6761\u76ee\u7684\u53c2\u8003\u5b9a\u989d\u6c60\uff08\u9ed8\u8ba4\u53d6\u5b9a\u989d\u8f93\u5165\u8868\u5f53\u524d\u884c\u7684\u5b9a\u989d\u7f16\u53f7\uff09";
                addBtn.Click += AddButtonClick;
                ToolStripButton delBtn = new ToolStripButton("\u5220\u9664\u5b9a\u989d");
                delBtn.DisplayStyle = ToolStripItemDisplayStyle.Text;
                delBtn.ToolTipText = "\u628a\u53c2\u8003\u6846\u4e2d\u9009\u4e2d\u7684\u5b9a\u989d\u4ece\u5f53\u524d\u6761\u76ee\u7684\u53c2\u8003\u5b9a\u989d\u6c60\u79fb\u9664";
                delBtn.Click += DeleteButtonClick;
                entryLabel = new ToolStripLabel("");
                entryLabel.Alignment = ToolStripItemAlignment.Right;
                bar.Items.Add(addBtn);
                bar.Items.Add(new ToolStripSeparator());
                bar.Items.Add(delBtn);
                bar.Items.Add(entryLabel);

                hostPanel = new Panel();
                hostPanel.Name = "recoReferenceQuotaHost";
                hostPanel.Dock = DockStyle.Fill;
                hostPanel.Controls.Add(refGrid);
                hostPanel.Controls.Add(bar);
                bar.Dock = DockStyle.Top;
                refGrid.Dock = DockStyle.Fill;
                // z 序：工具条置后(先 Dock 占顶部) + 表格置前(后 Dock 填剩余)，否则工具条会盖住表格列头
                bar.SendToBack();
                refGrid.BringToFront();

                container.Controls.Add(hostPanel);
                HideNativeGrids(container);
                hostPanel.BringToFront();

                // 强制创建表格句柄，并监听"参考定额"页可见性：切到该页时重新加载（非绑定表格在未激活时 Rows.Add 不生效）
                refContainer = container;
                container.VisibleChanged -= OnContainerVisibleChanged;
                container.VisibleChanged += OnContainerVisibleChanged;

                grid.CurrentCellChanged -= OnGridMove;
                grid.CurrentCellChanged += OnGridMove;
                grid.DataSourceChanged -= OnGridMove;
                grid.DataSourceChanged += OnGridMove;
                tree = GetField<TreeView>(mainForm, "Tv_tree");
                if (tree != null)
                {
                    tree.AfterSelect -= OnTreeSelect;
                    tree.AfterSelect += OnTreeSelect;
                }

                ShowEmpty();
                ScheduleRefresh();
                QuotaRecommendPanel.Log("Reference quota pool bound to container: " + container.Name + " pooledEntries=" + poolByEntry.Count.ToString(CultureInfo.InvariantCulture));
                LogRefGridState("after-install");
                return true;
            }

            public void Dispose()
            {
                if (disposed)
                {
                    return;
                }
                disposed = true;
                timer.Stop();
                timer.Dispose();
                try { grid.CurrentCellChanged -= OnGridMove; } catch { }
                try { grid.DataSourceChanged -= OnGridMove; } catch { }
                try { if (tree != null) tree.AfterSelect -= OnTreeSelect; } catch { }
                try { if (refContainer != null) refContainer.VisibleChanged -= OnContainerVisibleChanged; } catch { }
                try { if (nativeGrid != null) { nativeGrid.VisibleChanged -= NativeGridVisibleChanged; nativeGrid.Visible = true; } } catch { }
                try { if (hostPanel != null && hostPanel.Parent != null) hostPanel.Parent.Controls.Remove(hostPanel); } catch { }
                try { if (hostPanel != null) hostPanel.Dispose(); } catch { }
            }

            // 隐藏"参考定额"框里宿主原生的 dg_Common，并钉住隐藏状态，确保始终显示我们自管的表（编号/名称/单位/基期价格/内容）
            private void HideNativeGrids(Control container)
            {
                foreach (Control child in container.Controls)
                {
                    DataGridView dgv = child as DataGridView;
                    if (dgv != null && dgv != refGrid)
                    {
                        nativeGrid = dgv;
                        try { dgv.Visible = false; } catch { }
                        dgv.VisibleChanged -= NativeGridVisibleChanged;
                        dgv.VisibleChanged += NativeGridVisibleChanged;
                        break;
                    }
                }
            }

            private void NativeGridVisibleChanged(object sender, EventArgs e)
            {
                if (disposed || nativeGrid == null)
                {
                    return;
                }
                if (nativeGrid.Visible)
                {
                    try { nativeGrid.Visible = false; } catch { }
                }
            }

            // 宿主刷新可能把原生表重新置前/显示，每次刷新都把我们的面板钉回最上层
            private void KeepOnTop()
            {
                try
                {
                    if (nativeGrid != null && nativeGrid.Visible) { nativeGrid.Visible = false; }
                    if (hostPanel != null) { hostPanel.BringToFront(); }
                }
                catch { }
            }

            // 诊断：把覆盖表的真实列/可见性/层级写日志，用于核对"参考定额"框显示
            private void LogRefGridState(string tag)
            {
                try
                {
                    StringBuilder cols = new StringBuilder();
                    if (refGrid != null)
                    {
                        foreach (DataGridViewColumn c in refGrid.Columns) { cols.Append("[").Append(c.HeaderText).Append("]"); }
                    }
                    int childIndex = hostPanel != null && hostPanel.Parent != null ? hostPanel.Parent.Controls.GetChildIndex(hostPanel) : -99;
                    QuotaRecommendPanel.Log("RefGrid[" + tag + "] gridVisible=" + (refGrid != null && refGrid.Visible)
                        + " cols=" + (refGrid == null ? -1 : refGrid.Columns.Count) + cols
                        + " rows=" + (refGrid == null ? -1 : refGrid.Rows.Count)
                        + " hostVisible=" + (hostPanel != null && hostPanel.Visible)
                        + " hostBounds=" + (hostPanel == null ? "" : hostPanel.Bounds.ToString())
                        + " hostChildIndex=" + childIndex.ToString(CultureInfo.InvariantCulture)
                        + " nativeVisible=" + (nativeGrid == null ? "n/a" : nativeGrid.Visible.ToString()));
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("RefGrid state log failed: " + ex.Message);
                }
            }

            private DataGridView BuildRefGrid()
            {
                DataGridView g = new DataGridView();
                g.Name = "recoReferenceQuotaGrid";
                g.Dock = DockStyle.Fill;
                g.ReadOnly = true;
                g.AllowUserToAddRows = false;
                g.AllowUserToDeleteRows = false;
                g.AllowUserToResizeRows = false;
                g.RowHeadersVisible = false;
                g.ColumnHeadersVisible = true;
                g.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                g.ColumnHeadersHeight = 28; // 固定表头高度，避免 AutoSize 压成 0 导致表头不显示
                g.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                g.MultiSelect = false;
                g.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                g.EditMode = DataGridViewEditMode.EditProgrammatically;
                g.BackgroundColor = SystemColors.Window;
                try { g.Font = grid.Font; } catch { }
                // 表头加粗 + 蓝底，确保一眼可辨、与数据行区分
                g.EnableHeadersVisualStyles = false;
                g.ColumnHeadersDefaultCellStyle.BackColor = SystemColors.GradientActiveCaption;
                g.ColumnHeadersDefaultCellStyle.ForeColor = SystemColors.ControlText;
                try { g.ColumnHeadersDefaultCellStyle.Font = new Font(g.Font, FontStyle.Bold); } catch { }
                // 不绑定 DataSource，按 RecommendDialog 一样的 显式建列 + 手动加行 方式，表头才稳定显示
                g.Columns.Add("c_code", "\u7f16\u53f7");
                g.Columns.Add("c_name", "\u540d\u79f0");
                g.Columns.Add("c_unit", "\u5355\u4f4d");
                g.Columns.Add("c_price", "\u57fa\u671f\u4ef7\u683c");
                g.Columns.Add("c_content", "\u5185\u5bb9");
                g.Columns["c_code"].FillWeight = 16;
                g.Columns["c_name"].FillWeight = 34;
                g.Columns["c_unit"].FillWeight = 9;
                g.Columns["c_price"].FillWeight = 12;
                g.Columns["c_content"].FillWeight = 29;
                g.Columns["c_price"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                foreach (DataGridViewColumn col in g.Columns) { col.SortMode = DataGridViewColumnSortMode.NotSortable; }
                g.CellDoubleClick += RefGridDoubleClick;
                return g;
            }

            private void TimerTick(object sender, EventArgs e)
            {
                timer.Stop();
                RefreshSafe();
            }

            private void ScheduleRefresh()
            {
                timer.Stop();
                timer.Start();
            }

            private void OnGridMove(object sender, EventArgs e) { ScheduleRefresh(); }
            private void OnTreeSelect(object sender, TreeViewEventArgs e) { ScheduleRefresh(); }

            // 切到"参考定额"页（容器变可见）时强制重新加载：此时表格已激活，Rows.Add 才会真正加进去
            private void OnContainerVisibleChanged(object sender, EventArgs e)
            {
                if (disposed || refContainer == null || !refContainer.Visible)
                {
                    return;
                }
                currentKey = null;
                ScheduleRefresh();
            }

            private void RefreshSafe()
            {
                if (disposed || refGrid == null || refGrid.IsDisposed)
                {
                    return;
                }
                try
                {
                    Refresh();
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Reference quota pool refresh failed: " + ex.Message);
                }
            }

            private void Refresh()
            {
                KeepOnTop();
                EntryScope scope = ResolveCurrentScope();
                UpdateEntryLabel(scope);
                if (scope == null || String.IsNullOrEmpty(scope.MatchedEntryCode))
                {
                    SetKey("\0none");
                    return;
                }

                List<PoolItem> items;
                poolByEntry.TryGetValue(scope.MatchedEntryCode, out items);
                // 参考框只显示定额，过滤掉材料（全数字代号等）
                List<PoolItem> quotas = items == null
                    ? null
                    : items.Where(x => !String.Equals(x.Kind, "material", StringComparison.OrdinalIgnoreCase)).ToList();
                if (quotas == null || quotas.Count == 0)
                {
                    SetKey("\0empty:" + scope.Method + ":" + scope.MatchedEntryCode);
                    return;
                }

                string key = "pool:" + scope.Method + ":" + scope.MatchedEntryCode;
                if (key == currentKey)
                {
                    return;
                }
                currentKey = key;
                BindItems(quotas);
            }

            // 解析当前定额输入行/属性/树节点所属的章节条目，并映射到库内条目（含编制办法匹配校验）
            private EntryScope ResolveCurrentScope()
            {
                if (chapterLibrary == null || chapterLibrary.IsEmpty)
                {
                    return null;
                }
                SqlConnection conn = GetField<SqlConnection>(mainForm, "m_ProjectConn");
                if (conn != null && !ProjectUsesLibraryMethod(conn))
                {
                    return null;
                }
                string entryName;
                string entryCode = ResolveCurrentChapterNo(conn, out entryName);
                if (String.IsNullOrWhiteSpace(entryCode))
                {
                    return null;
                }
                return chapterLibrary.ResolveScope(entryCode, entryName);
            }

            private void UpdateEntryLabel(EntryScope scope)
            {
                if (entryLabel == null)
                {
                    return;
                }
                entryLabel.Text = scope == null || String.IsNullOrEmpty(scope.MatchedEntryCode)
                    ? ""
                    : ("\u5f53\u524d\u6761\u76ee: " + scope.MatchedEntryCode);
            }

            // 空态：保留五列表头、清空数据；用 currentKey 去抖避免同条目反复重建
            private void SetKey(string key)
            {
                if (key == currentKey)
                {
                    return;
                }
                currentKey = key;
                ShowEmpty();
            }

            private void ShowEmpty()
            {
                displayedItems = new List<PoolItem>();
                if (refGrid == null || refGrid.IsDisposed)
                {
                    return;
                }
                refGrid.Rows.Clear();
                refGrid.ClearSelection();
            }

            private void BindItems(List<PoolItem> items)
            {
                displayedItems = new List<PoolItem>(items);
                if (refGrid == null || refGrid.IsDisposed)
                {
                    return;
                }
                refGrid.Rows.Clear();
                foreach (PoolItem it in items)
                {
                    QuotaInfo info;
                    quotaIndex.TryGetValue((it.Code ?? "").ToUpperInvariant(), out info);
                    string name = info != null && info.Name.Length > 0 ? info.Name : it.Name;
                    string unit = info != null && info.Unit.Length > 0 ? info.Unit : it.Unit;
                    string price = info != null ? info.Price : "";
                    string content = info != null ? info.Content : "";
                    refGrid.Rows.Add(it.Code, name, unit, price, content);
                }
                refGrid.ClearSelection();
                LogRefGridState("after-bind rows=" + items.Count.ToString(CultureInfo.InvariantCulture));
            }

            private void ApplyColumnStyles()
            {
                if (refGrid == null || refGrid.Columns.Count < 5)
                {
                    return;
                }
                refGrid.Columns[0].FillWeight = 16; // 编号
                refGrid.Columns[1].FillWeight = 34; // 名称
                refGrid.Columns[2].FillWeight = 9;  // 单位
                refGrid.Columns[3].FillWeight = 12; // 基期价格
                refGrid.Columns[4].FillWeight = 29; // 内容
                refGrid.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }

            private void RefGridDoubleClick(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0 || refGrid == null || e.RowIndex >= refGrid.Rows.Count)
                {
                    return;
                }
                object value = refGrid.Rows[e.RowIndex].Cells[0].Value;
                string code = value == null ? "" : Convert.ToString(value, CultureInfo.CurrentCulture).Trim();
                if (code.Length == 0)
                {
                    return;
                }
                // 延迟到双击事件结束后再执行：让焦点从只读参考表干净地转回宿主定额输入表，原生粘贴才会落到正确网格
                try { mainForm.BeginInvoke((MethodInvoker)delegate { ApplyCode(code); }); }
                catch { ApplyCode(code); }
            }

            // 双击参考定额 -> 复用内联检索那套已验证成功的"原生单回车提交"：
            // 设光标到定额编号列 -> BeginEdit -> 往 EditingControl 写编号 -> 回车，宿主自动解析名称/单位/单价/消耗。
            private void ApplyCode(string code)
            {
                if (String.IsNullOrWhiteSpace(code) || grid == null)
                {
                    return;
                }
                try
                {
                    int codeCol = FindColumnIndex(grid, QuotaCodeColumns());
                    if (codeCol < 0)
                    {
                        QuotaRecommendPanel.Log("Reference quota apply: code column not found.");
                        return;
                    }
                    int targetRow = ResolveTargetRow(codeCol);
                    if (targetRow < 0)
                    {
                        QuotaRecommendPanel.Log("Reference quota apply: no target row.");
                        return;
                    }
                    bool filled = NativeEnterCommit(code.Trim(), targetRow, codeCol);
                    QuotaRecommendPanel.Log("Reference quota native-enter: " + code.Trim()
                        + " row=" + targetRow.ToString(CultureInfo.InvariantCulture)
                        + " filled=" + filled.ToString()
                        + " data=" + DescribeInputRow(targetRow));
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Reference quota apply failed: " + ex.Message);
                }
            }

            // 选目标行：当前行若定额编号为空就用当前行；否则找一个空行；都没有用最后一行
            private int ResolveTargetRow(int codeCol)
            {
                int cur = grid.CurrentCell == null ? -1 : grid.CurrentCell.RowIndex;
                if (cur >= 0 && cur < grid.Rows.Count && !grid.Rows[cur].IsNewRow && IsBlankCode(cur, codeCol))
                {
                    return cur;
                }
                for (int r = grid.Rows.Count - 1; r >= 0; r--)
                {
                    if (grid.Rows[r].IsNewRow)
                    {
                        continue;
                    }
                    if (IsBlankCode(r, codeCol))
                    {
                        return r;
                    }
                }
                return cur >= 0 ? cur : grid.Rows.Count - 1;
            }

            private bool IsBlankCode(int row, int codeCol)
            {
                if (row < 0 || row >= grid.Rows.Count)
                {
                    return false;
                }
                object v = grid.Rows[row].Cells[codeCol].Value;
                return v == null || String.IsNullOrWhiteSpace(Convert.ToString(v, CultureInfo.CurrentCulture));
            }

            // 原生单回车提交：与 QuotaInlineSearchFeature.TryApplyViaNativeEnterCommit 同一套已验证成功的写入
            private bool NativeEnterCommit(string code, int targetRow, int codeCol)
            {
                try
                {
                    grid.Focus();
                    grid.ClearSelection();
                    grid.CurrentCell = grid.Rows[targetRow].Cells[codeCol];
                    grid.Rows[targetRow].Selected = true;
                    bool began = grid.BeginEdit(true);
                    TextBoxBase ed = grid.EditingControl as TextBoxBase;
                    if (!began || ed == null)
                    {
                        QuotaRecommendPanel.Log("Reference quota native edit unavailable: began=" + began.ToString());
                        return false;
                    }
                    ed.Text = code;
                    ed.SelectionStart = ed.TextLength;
                    ed.SelectionLength = 0;
                    grid.NotifyCurrentCellDirty(true);
                    SendKeys.SendWait("{ENTER}");
                    Application.DoEvents();
                    return RowLooksFilled(targetRow, codeCol);
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Reference quota native-enter failed: " + ex.Message);
                    return false;
                }
            }

            private bool RowLooksFilled(int row, int codeCol)
            {
                if (row < 0 || row >= grid.Rows.Count)
                {
                    return false;
                }
                string c = Convert.ToString(grid.Rows[row].Cells[codeCol].Value, CultureInfo.CurrentCulture);
                if (String.IsNullOrWhiteSpace(c))
                {
                    return false;
                }
                return !String.IsNullOrWhiteSpace(GetRowValue(grid.Rows[row], QuotaNameColumns()));
            }

            // 把定额输入表当前单元格移到末尾新增行的"定额编号"列（粘贴落点），与 AgentExecutor.MoveAgentGridToNewRow 一致
            private bool TryApplyViaQuotaCodeCell(string code)
            {
                if (String.IsNullOrWhiteSpace(code))
                {
                    return false;
                }

                int targetRow = MoveGridToTargetRow();
                int codeColumn = FindColumnIndex(grid, QuotaCodeColumns());
                if (targetRow < 0 || targetRow >= grid.Rows.Count || codeColumn < 0)
                {
                    return false;
                }

                try
                {
                    grid.Focus();
                    Application.DoEvents();
                    grid.BeginEdit(true);
                    Application.DoEvents();
                    Clipboard.SetText(code.Trim());
                    SendKeys.SendWait("^a");
                    SendKeys.SendWait("^v");
                    SendKeys.SendWait("{ENTER}");
                    Application.DoEvents();

                    bool filled = TargetRowLooksNativeFilled(targetRow, code);
                    QuotaRecommendPanel.Log("Reference quota code-cell submitted: " + code.Trim()
                        + " row=" + targetRow.ToString(CultureInfo.InvariantCulture)
                        + " filled=" + filled.ToString()
                        + " data=" + DescribeInputRow(targetRow));
                    return filled;
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Reference quota code-cell submit failed: " + ex.Message);
                    return false;
                }
            }

            private bool TryApplyDirectCode(string code)
            {
                if (String.IsNullOrWhiteSpace(code))
                {
                    return false;
                }

                QuotaInfo info;
                if (!quotaIndex.TryGetValue(code.Trim().ToUpperInvariant(), out info) || info == null)
                {
                    QuotaRecommendPanel.Log("Reference quota direct fill skipped: quota index missing " + code);
                    return false;
                }

                int targetRow = MoveGridToTargetRow();
                if (targetRow < 0 || targetRow >= grid.Rows.Count)
                {
                    return false;
                }

                try
                {
                    DataGridViewRow row = grid.Rows[targetRow];
                    double quantity;
                    bool hasQuantity = TryParseDouble(GetRowValue(row, new[] { "\u5de5\u7a0b\u6570\u91cf\u8f93\u5165", "\u5de5\u7a0b\u6570\u91cf" }), out quantity);
                    double price;
                    bool hasPrice = TryParseDouble(info.Price, out price);

                    SetRowValue(row, QuotaCodeColumns(), code.Trim());
                    SetRowValue(row, QuotaNameColumns(), info.Name ?? "");
                    SetRowValue(row, QuotaUnitColumns(), info.Unit ?? "");
                    if (hasPrice)
                    {
                        SetRowValue(row, new[] { "\u5355\u4ef7", "\u57fa\u671f\u4ef7\u683c" }, price);
                        if (hasQuantity)
                        {
                            SetRowValue(row, new[] { "\u5408\u4ef7" }, quantity * price);
                        }
                    }

                    grid.CurrentCell = row.Cells[Math.Max(0, FindColumnIndex(grid, QuotaCodeColumns()))];
                    grid.EndEdit();
                    grid.InvalidateRow(targetRow);
                    QuotaRecommendPanel.Log("Reference quota direct-filled: " + code
                        + " row=" + targetRow.ToString(CultureInfo.InvariantCulture)
                        + " data=" + DescribeInputRow(targetRow));
                    return TargetRowLooksNativeFilled(targetRow, code);
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Reference quota direct fill failed: " + ex.Message);
                    return false;
                }
            }

            private bool TargetRowLooksNativeFilled(int targetRowIndex, string code)
            {
                if (targetRowIndex < 0 || targetRowIndex >= grid.Rows.Count)
                {
                    return false;
                }

                DataGridViewRow row = grid.Rows[targetRowIndex];
                if (!String.Equals(NormalizeQuotaCode(GetRowValue(row, QuotaCodeColumns())), NormalizeQuotaCode(code), StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                return !String.IsNullOrWhiteSpace(GetRowValue(row, QuotaNameColumns())) ||
                    !String.IsNullOrWhiteSpace(GetRowValue(row, QuotaUnitColumns())) ||
                    !String.IsNullOrWhiteSpace(GetRowValue(row, new[] { "\u5355\u4ef7", "\u57fa\u671f\u4ef7\u683c" }));
            }

            private int MoveGridToTargetRow()
            {
                int rowIndex = grid.CurrentCell == null ? -1 : grid.CurrentCell.RowIndex;
                if (rowIndex < 0 || rowIndex >= grid.Rows.Count)
                {
                    rowIndex = grid.Rows.Count - 1;
                }

                int columnIndex = FindColumnIndex(grid, QuotaCodeColumns());
                if (rowIndex < 0 || columnIndex < 0)
                {
                    return -1;
                }

                try
                {
                    grid.Focus();
                    grid.ClearSelection();
                    grid.CurrentCell = grid.Rows[rowIndex].Cells[columnIndex];
                    grid.Rows[rowIndex].Selected = true;
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Reference quota move-to-target-row failed: " + ex.Message);
                }
                return rowIndex;
            }

            private void MoveCurrentCellAwayFromCode(int targetRowIndex, int codeColumn)
            {
                if (targetRowIndex < 0 || targetRowIndex >= grid.Rows.Count)
                {
                    return;
                }

                int nextColumn = FindColumnIndex(grid, QuotaNameColumns());
                if (nextColumn < 0 || nextColumn == codeColumn || !grid.Columns[nextColumn].Visible)
                {
                    foreach (DataGridViewColumn column in grid.Columns)
                    {
                        if (column.Visible && column.Index != codeColumn)
                        {
                            nextColumn = column.Index;
                            break;
                        }
                    }
                }
                if (nextColumn >= 0 && nextColumn < grid.Columns.Count)
                {
                    grid.CurrentCell = grid.Rows[targetRowIndex].Cells[nextColumn];
                }
            }

            private void LogPasteProbe(string phase, int targetRowIndex, string code)
            {
                try
                {
                    QuotaRecommendPanel.Log("Reference paste probe[" + phase + "] code=" + (code ?? "")
                        + " row=" + targetRowIndex.ToString(CultureInfo.InvariantCulture)
                        + " rows=" + grid.Rows.Count.ToString(CultureInfo.InvariantCulture)
                        + " allowAdd=" + grid.AllowUserToAddRows
                        + " newRowIdx=" + grid.NewRowIndex.ToString(CultureInfo.InvariantCulture)
                        + " focused=" + grid.Focused
                        + " edit=" + (grid.EditingControl == null ? "null" : grid.EditingControl.GetType().FullName)
                        + " curCell=" + (grid.CurrentCell == null ? "null" : (grid.CurrentCell.RowIndex.ToString(CultureInfo.InvariantCulture) + "," + grid.CurrentCell.ColumnIndex.ToString(CultureInfo.InvariantCulture)))
                        + " rowData=" + DescribeInputRow(targetRowIndex));
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Reference paste probe[" + phase + "] failed: " + ex.Message);
                }
            }

            private void ScheduleDelayedPasteProbe(int actualRowIndex, int pasteRowIndex, string code)
            {
                SchedulePasteProbeOnce("delay-500", 500, actualRowIndex, pasteRowIndex, code);
                SchedulePasteProbeOnce("delay-1500", 1500, actualRowIndex, pasteRowIndex, code);
            }

            private void SchedulePasteProbeOnce(string phase, int delayMs, int actualRowIndex, int pasteRowIndex, string code)
            {
                Timer probeTimer = new Timer();
                probeTimer.Interval = delayMs;
                probeTimer.Tick += delegate
                {
                    probeTimer.Stop();
                    probeTimer.Dispose();
                    LogPasteProbe(phase + "-actual", actualRowIndex, code);
                    if (pasteRowIndex != actualRowIndex)
                    {
                        LogPasteProbe(phase + "-paste", pasteRowIndex, code);
                    }
                };
                probeTimer.Start();
            }

            private string DescribeInputRow(int targetRowIndex)
            {
                if (targetRowIndex < 0 || targetRowIndex >= grid.Rows.Count)
                {
                    return "<out-of-range>";
                }

                DataGridViewRow row = grid.Rows[targetRowIndex];
                return "code=" + GetRowValue(row, QuotaCodeColumns())
                    + "|name=" + GetRowValue(row, QuotaNameColumns())
                    + "|unit=" + GetRowValue(row, QuotaUnitColumns())
                    + "|qty=" + GetRowValue(row, new[] { "\u5de5\u7a0b\u6570\u91cf\u8f93\u5165", "\u5de5\u7a0b\u6570\u91cf" })
                    + "|price=" + GetRowValue(row, new[] { "\u5355\u4ef7", "\u57fa\u671f\u4ef7\u683c" });
            }

            // 调宿主定额输入表右键菜单 contextMenuStripDE 的"粘贴"项，触发原生粘贴解析
            private bool TryInvokePasteMenu()
            {
                ContextMenuStrip menu = GetField<ContextMenuStrip>(mainForm, "contextMenuStripDE");
                if (menu == null)
                {
                    return false;
                }
                foreach (ToolStripItem item in menu.Items)
                {
                    if (item is ToolStripMenuItem && item.Available && item.Text != null &&
                        item.Text.IndexOf("\u7c98\u8d34", StringComparison.Ordinal) >= 0)
                    {
                        try
                        {
                            ((ToolStripMenuItem)item).PerformClick();
                            QuotaRecommendPanel.Log("Reference quota paste via menu: " + item.Text);
                            return true;
                        }
                        catch (Exception ex)
                        {
                            QuotaRecommendPanel.Log("Reference quota paste menu failed: " + ex.Message);
                            return false;
                        }
                    }
                }
                return false;
            }

            // ===== 增加 / 删除 参考定额池定额 =====

            // 增加：把一条定额加入当前条目的参考池（默认取定额输入表当前行的定额编号，弹框可改），追加 source=user 行
            private void AddButtonClick(object sender, EventArgs e)
            {
                try
                {
                    EntryScope scope = ResolveCurrentScope();
                    if (scope == null || !scope.Strict)
                    {
                        MessageBox.Show(mainForm, "\u8bf7\u5148\u5728\u5b9a\u989d\u8f93\u5165\u8868\u5b9a\u4f4d\u5230\u4e00\u4e2a\u6709\u53c2\u8003\u6c60\u7684\u7ae0\u8282\u6761\u76ee\u3002", "\u589e\u52a0\u53c2\u8003\u5b9a\u989d", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    string initial = grid.CurrentRow != null && !grid.CurrentRow.IsNewRow ? GetRowValue(grid.CurrentRow, QuotaCodeColumns()) : "";
                    string code = PromptForCode(initial);
                    if (String.IsNullOrWhiteSpace(code))
                    {
                        return;
                    }
                    code = code.Trim();
                    string name, unit;
                    FindQuotaDisplay(code, out name, out unit);
                    chapterLibrary.AddUserQuota(scope, "", code, name, unit);
                    ReloadPool();
                    currentKey = null;
                    RefreshSafe();
                    QuotaRecommendPanel.Log("Reference quota pool user add: entry=" + scope.MatchedEntryCode + " code=" + code);
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Reference quota pool add failed: " + ex.Message);
                }
            }

            // 删除：把参考框选中的定额从当前条目参考池移除（追加 deleted=1 墓碑行，软删除可恢复）
            private void DeleteButtonClick(object sender, EventArgs e)
            {
                try
                {
                    if (refGrid == null || refGrid.CurrentRow == null)
                    {
                        return;
                    }
                    int idx = refGrid.CurrentRow.Index;
                    if (idx < 0 || idx >= displayedItems.Count)
                    {
                        return;
                    }
                    PoolItem item = displayedItems[idx];
                    EntryScope scope = ResolveCurrentScope();
                    if (scope == null || String.IsNullOrEmpty(scope.MatchedEntryCode))
                    {
                        return;
                    }
                    if (MessageBox.Show(mainForm, "\u786e\u8ba4\u4ece\u5f53\u524d\u6761\u76ee\u7684\u53c2\u8003\u5b9a\u989d\u6c60\u5220\u9664\uff1a" + item.Code + " \uff1f", "\u5220\u9664\u53c2\u8003\u5b9a\u989d", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                    {
                        return;
                    }
                    chapterLibrary.RemoveUserQuota(scope, item.Kind, item.Code);
                    ReloadPool();
                    currentKey = null;
                    RefreshSafe();
                    QuotaRecommendPanel.Log("Reference quota pool user delete: entry=" + scope.MatchedEntryCode + " code=" + item.Code);
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Reference quota pool delete failed: " + ex.Message);
                }
            }

            // 重新从 chapter-quota-library.jsonl 装载富字段池（增删后保持显示与文件一致）
            private void ReloadPool()
            {
                Dictionary<string, List<PoolItem>> fresh = LoadPool(chapterLibrary == null ? "" : chapterLibrary.MethodKey);
                poolByEntry.Clear();
                foreach (KeyValuePair<string, List<PoolItem>> pair in fresh)
                {
                    poolByEntry[pair.Key] = pair.Value;
                }
            }

            // 在定额输入表里按定额编号找一行，取其名称/单位（增加时给参考池补全显示字段）
            private void FindQuotaDisplay(string code, out string name, out string unit)
            {
                name = "";
                unit = "";
                if (String.IsNullOrWhiteSpace(code) || grid == null)
                {
                    return;
                }
                string target = code.Trim();
                foreach (DataGridViewRow row in grid.Rows)
                {
                    if (row.IsNewRow)
                    {
                        continue;
                    }
                    if (String.Equals(GetRowValue(row, QuotaCodeColumns()), target, StringComparison.OrdinalIgnoreCase))
                    {
                        name = GetRowValue(row, QuotaNameColumns());
                        unit = GetRowValue(row, new[] { "\u5355\u4f4d" });
                        return;
                    }
                }
            }

            // 简易输入框：让用户确认/修改要加入参考池的定额编号
            private string PromptForCode(string initial)
            {
                using (Form dlg = new Form())
                {
                    dlg.Text = "\u589e\u52a0\u53c2\u8003\u5b9a\u989d";
                    dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                    dlg.StartPosition = FormStartPosition.CenterParent;
                    dlg.MinimizeBox = false;
                    dlg.MaximizeBox = false;
                    dlg.ShowInTaskbar = false;
                    dlg.ClientSize = new Size(330, 110);
                    Label lbl = new Label();
                    lbl.Text = "\u5b9a\u989d\u7f16\u53f7\uff1a";
                    lbl.SetBounds(12, 18, 80, 20);
                    TextBox box = new TextBox();
                    box.Text = initial ?? "";
                    box.SetBounds(92, 15, 220, 24);
                    box.SelectAll();
                    Button ok = new Button();
                    ok.Text = "\u786e\u5b9a";
                    ok.DialogResult = DialogResult.OK;
                    ok.SetBounds(148, 62, 78, 28);
                    Button cancel = new Button();
                    cancel.Text = "\u53d6\u6d88";
                    cancel.DialogResult = DialogResult.Cancel;
                    cancel.SetBounds(234, 62, 78, 28);
                    dlg.Controls.Add(lbl);
                    dlg.Controls.Add(box);
                    dlg.Controls.Add(ok);
                    dlg.Controls.Add(cancel);
                    dlg.AcceptButton = ok;
                    dlg.CancelButton = cancel;
                    return dlg.ShowDialog(mainForm) == DialogResult.OK ? box.Text : null;
                }
            }

            // ===== 当前章节条目解析（精简复制自 QuotaInlineSearchFeature，保持本文件独立）=====

            private bool ProjectUsesLibraryMethod(SqlConnection conn)
            {
                if (conn == null || String.IsNullOrWhiteSpace(chapterLibrary.MethodNo))
                {
                    return true;
                }
                string dbName;
                try { dbName = conn.Database ?? ""; }
                catch { return true; }
                bool cached;
                if (methodCache.TryGetValue(dbName, out cached))
                {
                    return cached;
                }
                bool matches = true;
                try
                {
                    EnsureConnectionOpen(conn);
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        // select 编制办法文号 from 项目信息
                        cmd.CommandText = "select \u7f16\u5236\u529e\u6cd5\u6587\u53f7 from \u9879\u76ee\u4fe1\u606f";
                        object result = cmd.ExecuteScalar();
                        string methodNo = result == null || result == DBNull.Value ? "" : Convert.ToString(result, CultureInfo.InvariantCulture).Trim();
                        if (!String.IsNullOrEmpty(methodNo))
                        {
                            matches = String.Equals(NormalizeMethodNo(methodNo), NormalizeMethodNo(chapterLibrary.MethodNo), StringComparison.OrdinalIgnoreCase);
                        }
                    }
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Reference quota project method check failed (treat as match): " + ex.Message);
                }
                methodCache[dbName] = matches;
                return matches;
            }

            private string ResolveCurrentChapterNo(SqlConnection conn, out string entryName)
            {
                entryName = "";
                if (grid.CurrentRow != null && !grid.CurrentRow.IsNewRow)
                {
                    // 条目序号
                    string seq = GetRowValue(grid.CurrentRow, new[] { "\u6761\u76ee\u5e8f\u53f7" });
                    string fromSeq = LookupChapterNoBySeq(conn, seq, out entryName);
                    if (!String.IsNullOrEmpty(fromSeq))
                    {
                        return fromSeq;
                    }
                }

                // 条目编号
                string fromProp = ReadPropertyGridValue("\u6761\u76ee\u7f16\u53f7");
                if (!String.IsNullOrEmpty(fromProp))
                {
                    // 工程或费用项目名称
                    entryName = ReadPropertyGridValue("\u5de5\u7a0b\u6216\u8d39\u7528\u9879\u76ee\u540d\u79f0") ?? "";
                    return fromProp;
                }

                TreeView tv = GetField<TreeView>(mainForm, "Tv_tree");
                TreeNode node = tv != null ? tv.SelectedNode : GetField<TreeNode>(mainForm, "CurrNode");
                if (node != null)
                {
                    string fromTag = TryGetTagValue(node.Tag, "\u6761\u76ee\u7f16\u53f7");
                    if (!String.IsNullOrEmpty(fromTag))
                    {
                        entryName = node.Text ?? "";
                        return fromTag;
                    }
                    string seq = TryGetTagValue(node.Tag, "\u6761\u76ee\u5e8f\u53f7");
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

            private static string LookupChapterNoBySeq(SqlConnection conn, string seq, out string entryName)
            {
                entryName = "";
                if (String.IsNullOrWhiteSpace(seq) || conn == null)
                {
                    return null;
                }
                EnsureConnectionOpen(conn);
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    // select 条目编号, 工程或费用项目名称 from 章节表 where 条目序号=@id
                    cmd.CommandText = "select \u6761\u76ee\u7f16\u53f7, \u5de5\u7a0b\u6216\u8d39\u7528\u9879\u76ee\u540d\u79f0 from \u7ae0\u8282\u8868 where \u6761\u76ee\u5e8f\u53f7=@id";
                    cmd.Parameters.AddWithValue("@id", seq.Trim());
                    using (SqlDataReader reader = cmd.ExecuteReader())
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
                    // 属性名称 / 数据
                    if (String.Equals(GetRowValue(row, new[] { "\u5c5e\u6027\u540d\u79f0" }), propertyName, StringComparison.Ordinal))
                    {
                        return GetRowValue(row, new[] { "\u6570\u636e" });
                    }
                }
                return null;
            }
        }

        // ===== 数据加载：自带富字段池（chapter-quota-library.jsonl）=====

        private sealed class PoolItem
        {
            public string Code;
            public string Name;
            public string Unit;
            public string Kind;
            public string Source;
            public int ProjectCount;
        }

        // 定额编号 -> 名称/单位/基期价格(基价)/内容(工作内容)，取自 quota-index.jsonl
        private sealed class QuotaInfo
        {
            public string Name = "";
            public string Unit = "";
            public string Price = "";
            public string Content = "";
        }

        private static Dictionary<string, QuotaInfo> LoadQuotaIndex()
        {
            Dictionary<string, QuotaInfo> map = new Dictionary<string, QuotaInfo>(StringComparer.OrdinalIgnoreCase);
            try
            {
                string path = Path.Combine(LearningStore.FindDataDir(), "quota-index.jsonl");
                if (!File.Exists(path))
                {
                    QuotaRecommendPanel.Log("Reference quota index missing: " + path);
                    return map;
                }
                foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
                {
                    if (String.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }
                    Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                    string code = LearningStore.Get(values, "quota_code").Trim();
                    if (code.Length == 0)
                    {
                        continue;
                    }
                    QuotaInfo info = new QuotaInfo();
                    info.Name = LearningStore.Get(values, "quota_name").Trim();
                    info.Unit = LearningStore.Get(values, "quota_unit").Trim();
                    info.Price = LearningStore.Get(values, "base_price").Trim();
                    info.Content = LearningStore.Get(values, "work_content").Trim();
                    map[code.ToUpperInvariant()] = info; // 同码后写覆盖先写
                }
                QuotaRecommendPanel.Log("Reference quota index loaded. quotas=" + map.Count.ToString(CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("Reference quota index load failed: " + ex.Message);
            }
            return map;
        }

        private static Dictionary<string, List<PoolItem>> LoadPool(string methodKey)
        {
            Dictionary<string, List<PoolItem>> map = new Dictionary<string, List<PoolItem>>(StringComparer.Ordinal);
            // entry -> (kind:CODE -> PoolItem)，按文件顺序后写覆盖先写；deleted=1 即移除该 code（软删除）
            Dictionary<string, Dictionary<string, PoolItem>> tmp = new Dictionary<string, Dictionary<string, PoolItem>>(StringComparer.Ordinal);
            try
            {
                string path = Path.Combine(LearningStore.FindDataDir(), "chapter-quota-library.jsonl");
                if (!File.Exists(path))
                {
                    return map;
                }
                foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
                {
                    if (String.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }
                    Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                    if (!String.Equals(LearningStore.Get(values, "method"), methodKey, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    string entry = LearningStore.Get(values, "entry_code").Trim();
                    string code = LearningStore.Get(values, "quota_code").Trim();
                    if (entry.Length == 0 || code.Length == 0)
                    {
                        continue;
                    }
                    string kind = LearningStore.Get(values, "target_kind").Trim().ToLowerInvariant();
                    if (kind.Length == 0)
                    {
                        kind = "quota";
                    }
                    string itemKey = kind + ":" + code.ToUpperInvariant();

                    Dictionary<string, PoolItem> inner;
                    if (!tmp.TryGetValue(entry, out inner))
                    {
                        inner = new Dictionary<string, PoolItem>(StringComparer.Ordinal);
                        tmp[entry] = inner;
                    }

                    if (LearningStore.Get(values, "deleted").Trim() == "1")
                    {
                        inner.Remove(itemKey);
                        continue;
                    }

                    PoolItem item = new PoolItem();
                    item.Code = code;
                    item.Name = LearningStore.Get(values, "quota_name").Trim();
                    item.Unit = LearningStore.Get(values, "quota_unit").Trim();
                    item.Kind = kind;
                    item.Source = LearningStore.Get(values, "source").Trim();
                    int projectCount;
                    int.TryParse(LearningStore.Get(values, "project_count").Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out projectCount);
                    item.ProjectCount = projectCount;
                    inner[itemKey] = item;
                }
                foreach (KeyValuePair<string, Dictionary<string, PoolItem>> pair in tmp)
                {
                    if (pair.Value.Count == 0)
                    {
                        continue;
                    }
                    List<PoolItem> list = new List<PoolItem>(pair.Value.Values);
                    list.Sort(ComparePoolItem);
                    map[pair.Key] = list;
                }
                QuotaRecommendPanel.Log("Reference quota pool data loaded. method=" + methodKey + " entries=" + map.Count.ToString(CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("Reference quota pool load failed: " + ex.Message);
            }
            return map;
        }

        private static int ComparePoolItem(PoolItem a, PoolItem b)
        {
            bool am = a.Kind == "material";
            bool bm = b.Kind == "material";
            if (am != bm)
            {
                return am ? 1 : -1; // 定额在前，材料在后
            }
            if (a.ProjectCount != b.ProjectCount)
            {
                return b.ProjectCount.CompareTo(a.ProjectCount); // 项目数降序
            }
            return String.Compare(a.Code, b.Code, StringComparison.OrdinalIgnoreCase);
        }

        private static DataTable BuildTable()
        {
            DataTable table = new DataTable();
            table.Columns.Add("\u7f16\u53f7"); // 定额编号
            table.Columns.Add("\u540d\u79f0");             // 名称
            table.Columns.Add("\u5355\u4f4d");             // 单位
            table.Columns.Add("\u57fa\u671f\u4ef7\u683c");       // 项目数
            table.Columns.Add("\u5185\u5bb9");             // 来源
            return table;
        }

        private static string SourceLabel(string source)
        {
            switch ((source ?? "").ToLowerInvariant())
            {
                case "seed": return "\u79cd\u5b50"; // 种子
                case "scan": return "\u626b\u63cf"; // 扫描
                case "user": return "\u91c7\u7eb3"; // 采纳
                default: return source ?? "";
            }
        }

        // ===== 运行时发现"参考定额"容器 + 控件清单日志 =====

        private static Control FindReferenceContainer(Form mainForm)
        {
            // 优先：选项卡页 TabPage.Text 含"参考定额"
            Control byTab = FindControl(mainForm, delegate(Control c)
            {
                TabPage tp = c as TabPage;
                return tp != null && Squeeze(tp.Text).Contains(ReferenceTabText);
            });
            if (byTab != null)
            {
                return byTab;
            }
            // 兜底：任意可容纳子控件、且 Text 含"参考定额"的容器（GroupBox/Panel/TabPage）
            return FindControl(mainForm, delegate(Control c)
            {
                return (c is Panel || c is GroupBox || c is TabPage) && Squeeze(c.Text).Contains(ReferenceTabText);
            });
        }

        private static Control FindControl(Control parent, Predicate<Control> match)
        {
            if (parent == null)
            {
                return null;
            }
            foreach (Control child in parent.Controls)
            {
                if (match(child))
                {
                    return child;
                }
                Control found = FindControl(child, match);
                if (found != null)
                {
                    return found;
                }
            }
            return null;
        }

        private static void LogControlInventory(Form mainForm)
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("Reference quota pool: control inventory:");
                WalkControls(mainForm, 0, sb);
                QuotaRecommendPanel.Log(sb.ToString());
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("Reference quota pool inventory failed: " + ex.Message);
            }
        }

        private static void WalkControls(Control parent, int depth, StringBuilder sb)
        {
            foreach (Control child in parent.Controls)
            {
                string indent = new string(' ', depth * 2);
                TabPage tp = child as TabPage;
                DataGridView dg = child as DataGridView;
                if (tp != null)
                {
                    sb.AppendLine(indent + "TabPage Name=" + child.Name + " Text=" + tp.Text);
                }
                else if (dg != null)
                {
                    StringBuilder cols = new StringBuilder();
                    try
                    {
                        foreach (DataGridViewColumn col in dg.Columns)
                        {
                            cols.Append("[").Append(col.HeaderText).Append("/").Append(col.Name).Append("]");
                        }
                    }
                    catch { }
                    sb.AppendLine(indent + "DataGridView Name=" + child.Name + " parent=" + parent.Name + " cols=" + cols);
                }
                else if (child is TabControl)
                {
                    sb.AppendLine(indent + "TabControl Name=" + child.Name);
                }
                WalkControls(child, depth + 1, sb);
            }
        }

        // 去掉半角/全角空格便于匹配标题；'　' 是全角空格
        private static string Squeeze(string text) { return (text ?? "").Replace(" ", "").Replace("\u3000", "").Trim(); }

        // ===== 通用助手（精简复制，保持本文件独立）=====

        private static int FindColumnIndex(DataGridView grid, string[] names)
        {
            if (grid == null)
            {
                return -1;
            }
            foreach (DataGridViewColumn column in grid.Columns)
            {
                if (ColumnMatches(column, names))
                {
                    return column.Index;
                }
            }
            return -1;
        }

        private static bool ColumnMatches(DataGridViewColumn column, string[] names)
        {
            foreach (string name in names)
            {
                if (String.Equals(column.DataPropertyName, name, StringComparison.OrdinalIgnoreCase) ||
                    String.Equals(column.Name, name, StringComparison.OrdinalIgnoreCase) ||
                    String.Equals(column.HeaderText, name, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }

        private static string GetRowValue(DataGridViewRow row, string[] names)
        {
            if (row == null)
            {
                return "";
            }
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
                    if (ColumnMatches(column, names))
                    {
                        object value = row.Cells[column.Index].Value;
                        if (value != null)
                        {
                            return Convert.ToString(value, CultureInfo.CurrentCulture).Trim();
                        }
                    }
                }
            }
            return "";
        }

        private static void SetRowValue(DataGridViewRow row, int columnIndex, object value)
        {
            if (row == null || row.DataGridView == null || columnIndex < 0 || columnIndex >= row.DataGridView.Columns.Count)
            {
                return;
            }
            DataRowView rowView = row.DataBoundItem as DataRowView;
            DataGridViewColumn column = row.DataGridView.Columns[columnIndex];
            if (rowView != null && !String.IsNullOrWhiteSpace(column.DataPropertyName) && rowView.DataView.Table.Columns.Contains(column.DataPropertyName))
            {
                rowView[column.DataPropertyName] = value ?? DBNull.Value;
                return;
            }
            row.Cells[columnIndex].Value = value;
        }

        private static bool SetRowValue(DataGridViewRow row, string[] names, object value)
        {
            if (row == null || row.DataGridView == null)
            {
                return false;
            }

            foreach (DataGridViewColumn column in row.DataGridView.Columns)
            {
                if (ColumnMatches(column, names))
                {
                    SetRowValue(row, column.Index, value);
                    return true;
                }
            }
            return false;
        }

        private static bool TryParseDouble(string text, out double value)
        {
            value = 0d;
            if (String.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            return Double.TryParse(text.Replace(",", "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out value) ||
                Double.TryParse(text.Replace(",", "").Trim(), NumberStyles.Float, CultureInfo.CurrentCulture, out value);
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

        private static bool IsAllDigits(string text) { return !String.IsNullOrEmpty(text) && text.All(Char.IsDigit); }
        // 归一化编制办法文号：各种破折号统一成 '-'；'–'=–, '—'=—, '－'=－
        private static string NormalizeMethodNo(string text) { return (text ?? "").Replace('\u2013', '-').Replace('\u2014', '-').Replace('\uff0d', '-').Replace(" ", "").Trim(); }
        private static void EnsureConnectionOpen(SqlConnection conn) { if (conn != null && conn.State != ConnectionState.Open) conn.Open(); }

        private static string NormalizeQuotaCode(string code)
        {
            return (code ?? "")
                .Replace('\u2013', '-')
                .Replace('\u2014', '-')
                .Replace('\uff0d', '-')
                .Replace(" ", "")
                .Trim()
                .ToUpperInvariant();
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

        private static string[] QuotaCodeColumns()
        {
            // 定额编号 / 定额编号DE / 编号
            return new[] { "\u5b9a\u989d\u7f16\u53f7", "\u5b9a\u989d\u7f16\u53f7DE", "\u7f16\u53f7" };
        }

        private static string[] QuotaNameColumns()
        {
            // \u5de5\u7a0b\u6216\u8d39\u7528\u9879\u76ee\u540d\u79f0 / \u540d\u79f0 / \u9879\u76ee\u540d\u79f0
            return new[] { "\u5de5\u7a0b\u6216\u8d39\u7528\u9879\u76ee\u540d\u79f0", "\u540d\u79f0", "\u9879\u76ee\u540d\u79f0" };
        }

        private static string[] QuotaUnitColumns()
        {
            return new[] { "\u5355\u4f4d", "\u5355\u4f4dDE" };
        }
    }
}
