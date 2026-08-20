using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        private static IEnumerable<string> GetTemplateFillSheetNames(string path, out string error)
        {
            string extension = Path.GetExtension(path ?? "");
            if (String.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase))
            {
                List<string> names = GetXlsxSheetNames(path, out error).ToList();
                if (names.Count > 0) return names;
            }

            return GetSheetNamesByNpoi(path, out error);
        }

        private sealed class TemplateFillPanel : Form
        {
            private readonly Form mainForm;
            private List<FillPreviewItem> preview = new List<FillPreviewItem>();

            private readonly ComboBox cmbTemplate = new ComboBox();
            private readonly Button btnDeleteTemplate = new Button();
            private readonly TextBox txtUnit = new TextBox();
            private readonly ComboBox cmbSourceSheet = new ComboBox();
            private readonly TextBox txtName = new TextBox();
            private readonly Button btnBuild = new Button();
            private readonly ComboBox cmbMode = new ComboBox();
            private readonly ComboBox cmbTargetWorkbook = new ComboBox();
            private readonly ComboBox cmbTargetSheet = new ComboBox();
            private readonly TextBox txtColumn = new TextBox();
            private readonly ComboBox cmbTargetUnit = new ComboBox();
            private readonly Button btnSmartLearningScope = new Button();
            private readonly ToolStripDropDown smartLearningScopeDropDown = new ToolStripDropDown();
            private readonly TreeView smartLearningScopeTree = new TreeView();
            private ToolStripControlHost smartLearningScopeHost;
            private readonly Button btnPreview = new Button();
            private readonly Button btnApply = new Button();
            private readonly Button btnCheckSelected = new Button();
            private readonly Button btnUncheckSelected = new Button();
            private readonly Label lblCurrentEntry = new Label();
            private readonly Label lblWriteScope = new Label();
            private readonly CheckBox chkNameMode = new CheckBox();
            private readonly ToolTip targetWorkbookToolTip = new ToolTip();
            private readonly SplitContainer split = new SplitContainer();
            private readonly TreeView itemTree = new TreeView();
            private readonly DataGridView grid = new DataGridView();
            private Dictionary<string, string> chapterNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            private string currentTreeScope = "";
            private bool updatingTreeChecks;
            private bool rebuildingTree;
            private bool updatingNameQuotaCell;
            private bool reloadingTargetWorkbooks;
            private bool rebuildingSmartLearningScopeTree;
            private int targetWorkbookReloadCount;
            private string smartActiveWorkbookPath = "";
            private string smartPreviewWorkbookPath = "";
            private SmartLearningScope selectedSmartLearningScope = SmartLearningScope.CreateAll();
            private TreeView hostTree;
            private bool busy;
            private bool smartPreviewReady;
            private SmartScopeLoadStatus smartLoadStatus = SmartScopeLoadStatus.Success;
            private bool currentEntryWritable;
            private int smartPreviewVersion;
            private SmartPreviewContext previewContext;
            private CurrentSmartEntry currentSmartEntry;
            private bool updatingSmartPreviewInputs;
            private string smartScopeNotice = "";

            private sealed class SmartPreviewContext
            {
                public SqlConnection ProjectConnection;
                public string ProjectConnectionIdentity = "";
                public long CurrentUnitId;
                public string CurrentUnitCode = "";
                public string SoftwarePartition = "";
                public string MethodNo = "";
                public string ScopeKind = "";
                public string ScopeEntryCode = "";
                public string WorkbookPath = "";
                public string Worksheet = "";
                public string TargetColumn = "";
                public int PreviewVersion;
            }

            private sealed class CurrentSmartEntry
            {
                public SqlConnection ProjectConnection;
                public string ProjectConnectionIdentity = "";
                public long UnitId;
                public string UnitCode = "";
                public long EntrySequence;
                public string EntryCode = "";
                public string EntryName = "";
                public TreeNode Node;
            }

            private readonly bool smartOnly;

            public TemplateFillPanel(Form owner) : this(owner, false)
            {
            }

            // smartOnlyMode=true:独立"推荐定额"窗口,上部只有目标一行,固定走学习库漏斗。
            public TemplateFillPanel(Form owner, bool smartOnlyMode)
            {
                mainForm = owner;
                smartOnly = smartOnlyMode;
                Text = smartOnly ? "推荐定额" : "模板铺量";
                StartPosition = FormStartPosition.CenterParent;
                ClientSize = new Size(900, 580);
                BuildLayout();
                if (smartOnly) HookSmartHostTree();
                if (!smartOnly)
                {
                    ReloadTemplateList();
                    ReloadSourceSheets();
                }
                if (smartOnly)
                {
                    ReloadSmartLearningScopeTree();
                    string ignoredWorkbook;
                    string ignoredError;
                    TryResolveSmartActiveWorkbook(out ignoredWorkbook, out ignoredError);
                    RefreshCurrentSmartEntry(false);
                }
                else ReloadTargetWorkbooks();
                string cur = GetCurrentUnitNo(mainForm);
                if (!String.IsNullOrEmpty(cur)) txtUnit.Text = cur;
                RefreshApplyEnabled();
            }

            private void BuildLayout()
            {
                int targetTop = smartOnly ? 12 : 79;
                if (!smartOnly)
                {
                // —— 生成模板 ——
                AddLabel("源单元号", 12, 15, 56);
                txtUnit.SetBounds(72, 12, 90, 23); txtUnit.Text = "_ZGS_01";
                AddLabel("源sheet", 175, 15, 48);
                cmbSourceSheet.SetBounds(225, 12, 150, 23);
                cmbSourceSheet.DropDownStyle = ComboBoxStyle.DropDown; // 可选可填
                AddLabel("模板名", 388, 15, 48);
                txtName.SetBounds(438, 12, 120, 23); txtName.Text = "";
                btnBuild.SetBounds(568, 11, 130, 25); btnBuild.Text = "从该单元生成模板";
                btnBuild.Click += delegate { OnBuild(); };
                chkNameMode.SetBounds(568, 38, 130, 20);
                chkNameMode.Text = "按名字生成";
                Controls.Add(chkNameMode);

                // —— 套用配置 ——
                AddLabel("模板", 12, 50, 36);
                cmbTemplate.SetBounds(50, 47, 185, 23); cmbTemplate.DropDownStyle = ComboBoxStyle.DropDownList;
                btnDeleteTemplate.SetBounds(240, 46, 70, 25); btnDeleteTemplate.Text = "删除模板";
                btnDeleteTemplate.Click += delegate { OnDeleteTemplate(); };
                AddLabel("取数模式", 320, 50, 60);
                cmbMode.SetBounds(385, 47, 150, 23); cmbMode.DropDownStyle = ComboBoxStyle.DropDownList;
                }
                // 选项与模式联动放在条件块外:推荐定额窗口虽不显示该下拉,仍靠 SelectedIndex==2 驱动预览分支。
                cmbMode.Items.AddRange(new object[] { "一·列锚点", "二·名字驱动", "三·推荐定额(学习库)" });
                cmbMode.SelectedIndex = smartOnly ? 2 : 0;
                cmbMode.SelectedIndexChanged += delegate
                {
                    bool smart = cmbMode.SelectedIndex == 2;
                    cmbTemplate.Enabled = !smart;
                    btnDeleteTemplate.Enabled = !smart;
                };
                if (smartOnly)
                {
                    AddLabel("推荐学习库", 12, targetTop + 3, 70);
                    btnSmartLearningScope.SetBounds(85, targetTop, 170, 23);
                    btnSmartLearningScope.Text = "全部学习库";
                    btnSmartLearningScope.TextAlign = ContentAlignment.MiddleLeft;
                    btnSmartLearningScope.Click += delegate { ShowSmartLearningScopeDropDown(); };
                    smartLearningScopeTree.BorderStyle = BorderStyle.None;
                    smartLearningScopeTree.HideSelection = false;
                    smartLearningScopeTree.ShowLines = true;
                    smartLearningScopeTree.ShowPlusMinus = true;
                    smartLearningScopeTree.ShowRootLines = true;
                    smartLearningScopeTree.AfterSelect += delegate(object sender, TreeViewEventArgs e)
                    {
                        OnSmartLearningScopeSelected(e.Node);
                    };
                }
                else
                {
                    AddLabel("目标Excel", 12, targetTop + 3, 60);
                    cmbTargetWorkbook.SetBounds(75, targetTop, 190, 23);
                    cmbTargetWorkbook.DropDownWidth = 420;
                    cmbTargetWorkbook.DropDownStyle = ComboBoxStyle.DropDownList;
                    cmbTargetWorkbook.DropDown += delegate { ReloadTargetWorkbooks(); };
                    cmbTargetWorkbook.SelectedIndexChanged += delegate { if (!reloadingTargetWorkbooks) ReloadTargetSheets(); };
                }
                AddLabel("目标sheet", 275, targetTop + 3, 60);
                cmbTargetSheet.SetBounds(340, targetTop, 105, 23); cmbTargetSheet.Text = "";
                cmbTargetSheet.DropDownStyle = ComboBoxStyle.DropDown; // 可选可填
                cmbTargetSheet.DropDown += delegate
                {
                    if (smartOnly)
                    {
                        string ignoredWorkbook;
                        string ignoredError;
                        TryResolveSmartActiveWorkbook(out ignoredWorkbook, out ignoredError);
                    }
                    else ReloadTargetSheets();
                };
                AddLabel("目标列", 455, targetTop + 3, 50);
                txtColumn.SetBounds(505, targetTop, 40, 23); txtColumn.Text = "";
                if (!smartOnly)
                {
                    AddLabel("目标单元", 555, targetTop + 3, 60);
                    cmbTargetUnit.SetBounds(620, targetTop, 80, 23); cmbTargetUnit.Text = "_ZGS_02";
                    cmbTargetUnit.DropDownStyle = ComboBoxStyle.DropDown; // 可选可填
                    cmbTargetUnit.DropDown += delegate { ReloadTargetUnits(); };
                    btnPreview.SetBounds(710, targetTop - 1, 60, 25);
                    btnApply.SetBounds(780, targetTop - 1, 108, 25);
                }
                else
                {
                    btnCheckSelected.SetBounds(555, targetTop - 1, 88, 25); btnCheckSelected.Text = "勾选选中行";
                    btnUncheckSelected.SetBounds(648, targetTop - 1, 78, 25); btnUncheckSelected.Text = "取消勾选";
                    btnPreview.SetBounds(731, targetTop - 1, 55, 25);
                    btnApply.SetBounds(791, targetTop - 1, 97, 25);
                    btnCheckSelected.Click += delegate { SetSelectedSmartGroupsChecked(true); };
                    btnUncheckSelected.Click += delegate { SetSelectedSmartGroupsChecked(false); };
                    lblCurrentEntry.AutoSize = false;
                    lblCurrentEntry.SetBounds(12, targetTop + 29, 520, 18);
                    lblWriteScope.AutoSize = false;
                    lblWriteScope.TextAlign = ContentAlignment.MiddleRight;
                    lblWriteScope.SetBounds(535, targetTop + 29, 353, 18);
                    Controls.Add(btnCheckSelected);
                    Controls.Add(btnUncheckSelected);
                    Controls.Add(lblCurrentEntry);
                    Controls.Add(lblWriteScope);
                }
                btnPreview.Text = "预览";
                btnPreview.Click += delegate { OnPreview(); };
                btnApply.Text = smartOnly ? "写入当前条目" : "写入目标单元";
                btnApply.Click += delegate { OnApply(); };

                Label reminder = new Label
                {
                    Text = smartOnly
                        ? "写入＝把选中且勾选的定额写入左侧章节树当前条目。写入完成即已保存；点“计算”只刷新单价、合价和汇总。"
                        : "写入＝复制定额到“目标单元”的对应条目（条目序号全局共享）。写入完成即已保存；点“计算”只刷新单价与汇总。",
                    ForeColor = Color.Firebrick, AutoSize = false
                };
                reminder.SetBounds(12, targetTop + (smartOnly ? 51 : 29), 876, 18);
                reminder.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

                // —— 左侧条目树 + 右侧预览表：SplitContainer 分栏，可拖动调整宽度 ——
                int contentTop = targetTop + (smartOnly ? 75 : 53);
                split.SetBounds(12, contentTop, 876, ClientSize.Height - contentTop - 12);
                split.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
                split.Orientation = Orientation.Vertical;
                split.Panel1MinSize = 100;
                split.Panel2MinSize = 200;
                split.SplitterDistance = 250;

                itemTree.Dock = DockStyle.Fill;
                itemTree.CheckBoxes = true;
                itemTree.HideSelection = false;
                itemTree.AfterSelect += delegate { OnTreeScopeChanged(); };
                itemTree.AfterCheck += delegate(object sender, TreeViewEventArgs e) { OnTreeNodeChecked(e.Node); };

                grid.Dock = DockStyle.Fill;
                grid.ReadOnly = false; grid.AllowUserToAddRows = false;
                grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                grid.MultiSelect = true;
                grid.RowHeadersVisible = false;
                grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "选", Name = "sel", FillWeight = 6 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "定额编号", Name = "code", ReadOnly = true, FillWeight = 16 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "源行定额", Name = "sname", ReadOnly = false, FillWeight = 20 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "单位", Name = "unit", ReadOnly = true, FillWeight = 8 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "目标行工程量名", Name = "tname", ReadOnly = true, FillWeight = 20 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "数量", Name = "qty", ReadOnly = false, FillWeight = 10 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "状态", Name = "st", ReadOnly = true, FillWeight = 14 });
                foreach (DataGridViewColumn column in grid.Columns)
                {
                    column.SortMode = DataGridViewColumnSortMode.NotSortable;
                }
                grid.CurrentCellDirtyStateChanged += delegate
                {
                    if (grid.IsCurrentCellDirty)
                    {
                        grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
                    }
                };
                grid.CellClick += delegate(object sender, DataGridViewCellEventArgs e)
                {
                    if (e.RowIndex < 0 || e.ColumnIndex < 0 ||
                        !String.Equals(grid.Columns[e.ColumnIndex].Name, "sname", StringComparison.Ordinal)) return;
                    DataGridViewRow row = grid.Rows[e.RowIndex];
                    if (!PrepareNameQuotaDropDown(row)) return;
                    grid.CurrentCell = row.Cells[e.ColumnIndex];
                    grid.BeginEdit(true);
                    ComboBox combo = grid.EditingControl as ComboBox;
                    if (combo != null) combo.DroppedDown = true;
                };
                grid.EditingControlShowing += delegate(object sender, DataGridViewEditingControlShowingEventArgs e)
                {
                    ComboBox combo = e.Control as ComboBox;
                    if (combo == null) return;
                    combo.SelectionChangeCommitted -= OnNameQuotaSelectionCommitted;
                    combo.SelectionChangeCommitted += OnNameQuotaSelectionCommitted;
                };
                grid.CellValueChanged += delegate(object sender, DataGridViewCellEventArgs e)
                {
                    if (updatingNameQuotaCell || e.RowIndex < 0 || e.ColumnIndex < 0) return;
                    if (String.Equals(grid.Columns[e.ColumnIndex].Name, "sel", StringComparison.Ordinal))
                    {
                        ApplyNameGroupSelectionFromCheck(grid.Rows[e.RowIndex]);
                        UpdateSmartWriteScope();
                    }
                    else if (String.Equals(grid.Columns[e.ColumnIndex].Name, "qty", StringComparison.Ordinal))
                    {
                        FillPreviewItem item = grid.Rows[e.RowIndex].Tag as FillPreviewItem;
                        if (item != null && ApplyEditedNameQuotaQuantity(item,
                            Convert.ToString(grid.Rows[e.RowIndex].Cells["qty"].Value).Trim()))
                            RefreshNameQuotaRiskStateInGrid(item.TargetRow);
                    }
                };
                grid.DataError += delegate(object sender, DataGridViewDataErrorEventArgs e)
                {
                    Log("Template fill grid data error: " + (e.Exception == null ? "unknown" : e.Exception.Message));
                    e.ThrowException = false;
                };
                // 一量对多定额:工程量名列按合并单元格绘制——组内行间不画横线,文字在整个合并区域内水平+垂直居中。
                grid.CellPainting += delegate(object sender, DataGridViewCellPaintingEventArgs e)
                {
                    if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                    if (!String.Equals(grid.Columns[e.ColumnIndex].Name, "tname", StringComparison.Ordinal)) return;
                    FillPreviewItem cur = grid.Rows[e.RowIndex].Tag as FillPreviewItem;
                    if (cur == null || !cur.IsNameDriven) return;
                    int start = e.RowIndex, end = e.RowIndex;
                    while (start > 0 && IsSameQuantityGroup(grid.Rows[start - 1].Tag as FillPreviewItem, cur)) start--;
                    while (end < grid.Rows.Count - 1 && IsSameQuantityGroup(grid.Rows[end + 1].Tag as FillPreviewItem, cur)) end++;
                    if (start == end) return; // 非一对多,走默认绘制

                    // 组内横线全部压掉:非首行不画上边线,非末行不画下边线。
                    if (e.RowIndex > start) e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None;
                    if (e.RowIndex < end) e.AdvancedBorderStyle.Bottom = DataGridViewAdvancedCellBorderStyle.None;
                    e.PaintBackground(e.ClipBounds, true);
                    Rectangle union = GetVisibleMergedTargetNameBounds(grid, e.ColumnIndex, start, end);
                    string text = Convert.ToString(grid.Rows[start].Cells[e.ColumnIndex].Value);
                    if (!union.IsEmpty && !String.IsNullOrEmpty(text))
                    {
                        bool selected = (e.State & DataGridViewElementStates.Selected) != 0;
                        Color foreColor = selected ? e.CellStyle.SelectionForeColor : e.CellStyle.ForeColor;
                        // 每个成员行都重绘自己与合并区域相交的部分。点击只使首行失效时，
                        // 不再依赖末行顺带重绘整块，避免合并文字在点击中或点击后消失。
                        DrawMergedTargetNameTextForCell(e.Graphics, text, e.CellStyle.Font,
                            union, e.CellBounds, foreColor);
                    }
                    e.Handled = true;
                };
                grid.SelectionChanged += delegate
                {
                    // 选择变化时 DataGridView 往往只重绘新旧选中行；合并工程量名跨多行，
                    // 必须让整列同时失效，否则未被重绘的成员行会留下被擦除的文字片段。
                    if (grid.Columns.Contains("tname"))
                        grid.InvalidateColumn(grid.Columns["tname"].Index);
                    if (!updatingNameQuotaCell) UpdateSmartWriteScope();
                };

                ContextMenuStrip gridMenu = new ContextMenuStrip();
                ToolStripMenuItem miBindSelected = new ToolStripMenuItem("绑定软件选中的定额到此行");
                gridMenu.Items.Add(miBindSelected);
                grid.ContextMenuStrip = gridMenu;
                grid.MouseDown += delegate(object sender, MouseEventArgs e)
                {
                    if (e.Button != MouseButtons.Right) return;
                    DataGridView.HitTestInfo hit = grid.HitTest(e.X, e.Y);
                    if (hit.RowIndex >= 0) { grid.ClearSelection(); grid.Rows[hit.RowIndex].Selected = true; grid.CurrentCell = grid.Rows[hit.RowIndex].Cells[0]; }
                };
                gridMenu.Opening += delegate(object sender, System.ComponentModel.CancelEventArgs e)
                {
                    FillPreviewItem cur = grid.SelectedRows.Count > 0 ? grid.SelectedRows[0].Tag as FillPreviewItem : null;
                    miBindSelected.Enabled = cur != null && cur.IsNameDriven;
                };
                miBindSelected.Click += delegate { OnBindSelectedQuotaToRow(); };

                split.Panel1.Controls.Add(itemTree);
                split.Panel2.Controls.Add(grid);
                if (smartOnly) split.Panel1Collapsed = true;

                if (!smartOnly)
                {
                    Controls.Add(txtUnit); Controls.Add(cmbSourceSheet); Controls.Add(txtName); Controls.Add(btnBuild);
                    Controls.Add(cmbTemplate); Controls.Add(btnDeleteTemplate); Controls.Add(cmbMode);
                }
                if (smartOnly) Controls.Add(btnSmartLearningScope);
                else Controls.Add(cmbTargetWorkbook);
                Controls.Add(cmbTargetSheet); Controls.Add(txtColumn);
                if (!smartOnly) Controls.Add(cmbTargetUnit);
                Controls.Add(btnPreview); Controls.Add(btnApply); Controls.Add(reminder); Controls.Add(split);
                if (smartOnly)
                {
                    cmbTargetSheet.TextChanged += delegate { InvalidateSmartPreview(); };
                    txtColumn.TextChanged += delegate { InvalidateSmartPreview(); };
                    Activated += delegate { RefreshCurrentSmartEntry(false); };
                    FormClosed += OnSmartPanelClosed;
                }
            }

            private void AddLabel(string text, int x, int y, int w)
            {
                Label l = new Label { Text = text, AutoSize = false };
                l.SetBounds(x, y, w, 18); Controls.Add(l);
            }

            private void HookSmartHostTree()
            {
                if (!smartOnly) return;
                hostTree = GetField<TreeView>(mainForm, "Tv_tree");
                if (hostTree == null) return;
                hostTree.AfterSelect += OnHostTreeSelected;
            }

            private void OnHostTreeSelected(object sender, TreeViewEventArgs e)
            {
                if (!smartOnly || IsDisposed || Disposing || !IsHandleCreated) return;
                try
                {
                    BeginInvoke(new MethodInvoker(delegate { RefreshCurrentSmartEntry(false); }));
                }
                catch { }
            }

            private void OnSmartPanelClosed(object sender, FormClosedEventArgs e)
            {
                try
                {
                    if (hostTree != null) hostTree.AfterSelect -= OnHostTreeSelected;
                }
                catch { }
                hostTree = null;
                previewContext = null;
            }

            private static bool IsEditableAgentQuotaGrid(DataGridView agentGrid)
            {
                if (agentGrid == null || !agentGrid.Visible || !agentGrid.Enabled || !agentGrid.AllowUserToAddRows ||
                    agentGrid.NewRowIndex < 0) return false;
                return agentGrid.Columns.Cast<DataGridViewColumn>().Any(column => column.Visible && !column.ReadOnly &&
                    ((column.Name ?? "").IndexOf("定额编号", StringComparison.OrdinalIgnoreCase) >= 0 ||
                     (column.HeaderText ?? "").IndexOf("定额编号", StringComparison.OrdinalIgnoreCase) >= 0));
            }

            private bool TryResolveCurrentSmartEntry(out CurrentSmartEntry result, out string error)
            {
                result = null;
                error = "请在左侧章节树选择可写入的具体条目";
                if (mainForm == null || mainForm.IsDisposed) return false;
                TreeView tree = GetField<TreeView>(mainForm, "Tv_tree");
                TreeNode currNode = GetField<TreeNode>(mainForm, "CurrNode");
                TreeNode node = ResolveSmartHostTreeNode(tree, currNode);
                if (node == null) return false;
                if (node.Nodes.Count != 0)
                {
                    error = "当前树节点不是可写入的叶条目";
                    return false;
                }
                DataGridView agentGrid = GetField<DataGridView>(mainForm, "dataGridViewDE");
                if (!IsEditableAgentQuotaGrid(agentGrid))
                {
                    error = "当前条目未处于可输入定额状态";
                    return false;
                }

                string seqText = TryGetValue(node.Tag, "条目序号");
                if (String.IsNullOrWhiteSpace(seqText) && IsNumeric(node.Name)) seqText = node.Name;
                string code = (TryGetValue(node.Tag, "条目编号") ?? "").Trim();
                if (code.Length == 0 && !String.IsNullOrWhiteSpace(node.Name) && !IsNumeric(node.Name))
                    code = node.Name.Trim();
                long sequence;
                if (!Int64.TryParse(seqText ?? "", NumberStyles.Integer, CultureInfo.InvariantCulture, out sequence) || sequence <= 0)
                {
                    error = "当前树节点缺少可核对的条目序号";
                    return false;
                }
                if (code.Length == 0)
                {
                    error = "当前树节点缺少可核对的条目编号";
                    return false;
                }

                SqlConnection conn = GetOpenProjectConnection(mainForm);
                string dbCode = "";
                string dbName = "";
                int matchCount = 0;
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "select 条目编号,工程或费用项目名称 from 章节表 where 条目序号=@seq";
                    cmd.Parameters.AddWithValue("@seq", sequence);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            matchCount++;
                            dbCode = reader.IsDBNull(0) ? "" : Convert.ToString(reader.GetValue(0)).Trim();
                            dbName = reader.IsDBNull(1) ? "" : Convert.ToString(reader.GetValue(1)).Trim();
                        }
                    }
                }
                if (matchCount != 1 || dbCode.Length == 0 || !String.Equals(code, dbCode, StringComparison.OrdinalIgnoreCase))
                {
                    error = "当前树节点与项目章节表身份不一致";
                    return false;
                }
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "select count(*) from 章节表 where 条目编号=@code and 条目序号=@seq";
                    cmd.Parameters.AddWithValue("@code", dbCode);
                    cmd.Parameters.AddWithValue("@seq", sequence);
                    if (Convert.ToInt32(cmd.ExecuteScalar(), CultureInfo.InvariantCulture) != 1)
                    {
                        error = "当前树节点不是项目中的唯一可写条目";
                        return false;
                    }
                }

                AgentSelectionSnapshot selection = CaptureAgentSelection(mainForm);
                if (selection == null || selection.CurrentUnitId <= 0)
                {
                    error = "无法识别当前单元";
                    return false;
                }
                result = new CurrentSmartEntry
                {
                    ProjectConnection = conn,
                    ProjectConnectionIdentity = GetProjectConnectionIdentity(conn),
                    UnitId = selection.CurrentUnitId,
                    UnitCode = selection.CurrentUnitCode ?? "",
                    EntrySequence = sequence,
                    EntryCode = dbCode,
                    EntryName = dbName,
                    Node = node
                };
                return true;
            }

            private static TreeNode ResolveSmartHostTreeNode(TreeView tree, TreeNode currNode)
            {
                // SelectedNode 有值时始终以界面真实选中项为准，不要求与宿主 CurrNode 引用相同；
                // 宿主切换焦点期间 SelectedNode 可能暂时为空，此时回退 CurrNode，再由后续项目章节表校验身份。
                return tree == null || tree.SelectedNode == null ? currNode : tree.SelectedNode;
            }

            private void RefreshCurrentSmartEntry(bool forApply)
            {
                if (!smartOnly) return;
                CurrentSmartEntry entry;
                string error;
                currentEntryWritable = TryResolveCurrentSmartEntry(out entry, out error);
                currentSmartEntry = currentEntryWritable ? entry : null;
                lblCurrentEntry.Text = currentEntryWritable
                    ? "当前条目：" + entry.EntryCode + "  " + entry.EntryName
                    : "当前条目：（" + error + "）";
                RefreshSmartSfEntryState();
                RefreshApplyEnabled();
                UpdateSmartWriteScope();
            }

            private HashSet<int> GetSelectedSmartTargetRows()
            {
                HashSet<int> selected = new HashSet<int>();
                foreach (DataGridViewRow row in grid.SelectedRows)
                {
                    FillPreviewItem item = row.Tag as FillPreviewItem;
                    if (item != null) selected.Add(item.TargetRow);
                }
                return selected;
            }

            private HashSet<int> GetCheckedSmartTargetRows()
            {
                return new HashSet<int>(preview.Where(item => item != null && item.IsNameDriven &&
                    item.GroupOrder == 0 && item.Selected).Select(item => item.TargetRow));
            }

            private void SetSelectedSmartGroupsChecked(bool value)
            {
                if (!smartOnly) return;
                HashSet<int> selectedRows = GetSelectedSmartTargetRows();
                foreach (FillPreviewItem item in preview.Where(item => item != null && selectedRows.Contains(item.TargetRow)))
                {
                    item.Selected = value;
                }
                bool old = updatingNameQuotaCell;
                updatingNameQuotaCell = true;
                try
                {
                    foreach (DataGridViewRow row in grid.Rows)
                    {
                        FillPreviewItem item = row.Tag as FillPreviewItem;
                        if (item != null && item.GroupOrder == 0 && selectedRows.Contains(item.TargetRow))
                            row.Cells["sel"].Value = value;
                    }
                }
                finally { updatingNameQuotaCell = old; }
                UpdateSmartWriteScope();
                RefreshApplyEnabled();
            }

            private void UpdateSmartWriteScope()
            {
                if (!smartOnly) return;
                HashSet<int> selected = GetSelectedSmartTargetRows();
                HashSet<int> checkedRows = GetCheckedSmartTargetRows();
                HashSet<int> intersection = new HashSet<int>(selected);
                intersection.IntersectWith(checkedRows);
                int itemCount = preview.Count(item => item != null && intersection.Contains(item.TargetRow));
                lblWriteScope.Text = "将写入 " + itemCount.ToString(CultureInfo.InvariantCulture) + " 条（选中 " +
                    selected.Count.ToString(CultureInfo.InvariantCulture) + " 组 ∩ 已勾选 " +
                    checkedRows.Count.ToString(CultureInfo.InvariantCulture) + " 组）";
            }

            private void InvalidateSmartPreview()
            {
                if (!smartOnly || updatingSmartPreviewInputs) return;
                smartPreviewVersion++;
                previewContext = null;
                smartPreviewReady = false;
                smartPreviewWorkbookPath = "";
                preview = new List<FillPreviewItem>();
                bool old = updatingNameQuotaCell;
                updatingNameQuotaCell = true;
                try { grid.Rows.Clear(); }
                finally { updatingNameQuotaCell = old; }
                UpdateSmartWriteScope();
                RefreshApplyEnabled();
            }

            private void RefreshApplyEnabled()
            {
                if (!smartOnly)
                {
                    btnApply.Enabled = !busy;
                    btnPreview.Enabled = !busy;
                    return;
                }
                bool hasWritableGroup = preview.GroupBy(item => item.TargetRow)
                    .Any(group => group.All(item => item != null && String.IsNullOrEmpty(item.Status) && !item.SfEntryBlocked));
                btnPreview.Enabled = !busy && smartLoadStatus == SmartScopeLoadStatus.Success;
                btnApply.Enabled = !busy && smartPreviewReady && previewContext != null && currentEntryWritable && hasWritableGroup;
            }

            // SF state is recomputed from the current host entry; it never mutates the persistent Status text.
            private void RefreshSmartSfEntryState()
            {
                if (!smartOnly || preview == null) return;
                foreach (FillPreviewItem item in preview)
                {
                    item.SfEntryBlocked = false;
                    item.SfEntryBlockReason = "";
                }
                if (!currentEntryWritable || currentSmartEntry == null) return;
                bool currentIsEquipment = (currentSmartEntry.EntryName ?? "").IndexOf("设备购置费", StringComparison.OrdinalIgnoreCase) >= 0;
                foreach (IGrouping<int, FillPreviewItem> group in preview.Where(item => item != null).GroupBy(item => item.TargetRow))
                {
                    bool hasSf = group.Any(item => String.Equals((item.QuotaCode ?? "").Trim(), "SF", StringComparison.OrdinalIgnoreCase));
                    bool hasNonSf = group.Any(item => !String.Equals((item.QuotaCode ?? "").Trim(), "SF", StringComparison.OrdinalIgnoreCase));
                    string reason = "";
                    if (currentIsEquipment && hasNonSf) reason = "设备购置费条目只接受 SF";
                    else if (hasSf && !currentIsEquipment)
                    {
                        long sfSeq;
                        string sfCode;
                        string sfName;
                        if (!TryResolveSiblingEquipmentEntry(currentSmartEntry.ProjectConnection, currentSmartEntry.EntryCode,
                            out sfSeq, out sfCode, out sfName, out reason)) { }
                    }
                    if (reason.Length == 0) continue;
                    foreach (FillPreviewItem item in group)
                    {
                        item.SfEntryBlocked = true;
                        item.SfEntryBlockReason = reason;
                    }
                }
            }

            private static bool TryResolveSiblingEquipmentEntry(SqlConnection conn, string currentEntryCode,
                out long entrySequence, out string entryCode, out string entryName, out string error)
            {
                entrySequence = 0;
                entryCode = "";
                entryName = "";
                error = "";
                string[] currentParts = (currentEntryCode ?? "").Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
                if (currentParts.Length < 2)
                {
                    error = "当前条目无法确定同级设备购置费条目";
                    return false;
                }
                List<CurrentSmartEntry> candidates = new List<CurrentSmartEntry>();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "select 条目序号,条目编号,工程或费用项目名称 from 章节表 where 工程或费用项目名称 like @name";
                    cmd.Parameters.AddWithValue("@name", "%设备购置费%");
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string candidateCode = reader.IsDBNull(1) ? "" : Convert.ToString(reader.GetValue(1)).Trim();
                            string[] parts = candidateCode.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
                            if (parts.Length != currentParts.Length) continue;
                            bool sameParent = true;
                            for (int i = 0; i < parts.Length - 1; i++)
                            {
                                if (!String.Equals(parts[i], currentParts[i], StringComparison.OrdinalIgnoreCase))
                                {
                                    sameParent = false;
                                    break;
                                }
                            }
                            if (!sameParent) continue;
                            candidates.Add(new CurrentSmartEntry
                            {
                                EntrySequence = Convert.ToInt64(reader.GetValue(0), CultureInfo.InvariantCulture),
                                EntryCode = candidateCode,
                                EntryName = reader.IsDBNull(2) ? "" : Convert.ToString(reader.GetValue(2)).Trim()
                            });
                        }
                    }
                }
                if (candidates.Count != 1)
                {
                    error = candidates.Count == 0
                        ? "未找到唯一的同级设备购置费条目"
                        : "同级设备购置费条目不唯一";
                    return false;
                }
                entrySequence = candidates[0].EntrySequence;
                entryCode = candidates[0].EntryCode;
                entryName = candidates[0].EntryName;
                return true;
            }

            // 同一 Excel 工程量行产生的多条定额 = 一个合并显示组。
            private static bool IsSameQuantityGroup(FillPreviewItem other, FillPreviewItem current)
            {
                return other != null && current != null && other.IsNameDriven && current.IsNameDriven &&
                    other.TargetRow == current.TargetRow;
            }

            private void ReloadTemplateList()
            {
                cmbTemplate.Items.Clear();
                foreach (string n in ListFillTemplateNames()) cmbTemplate.Items.Add(n);
                if (cmbTemplate.Items.Count > 0) cmbTemplate.SelectedIndex = 0;
            }

            // 源sheet 下拉：列出绑定库里记录过的 Excel 工作表名。
            private void ReloadSourceSheets()
            {
                try
                {
                    string keep = cmbSourceSheet.Text;
                    cmbSourceSheet.Items.Clear();
                    SqlConnection conn = GetOpenProjectConnection(mainForm);
                    foreach (string s in ListBoundSheetNames(conn)) cmbSourceSheet.Items.Add(s);
                    if (!String.IsNullOrEmpty(keep)) cmbSourceSheet.Text = keep;
                    else if (cmbSourceSheet.Items.Count > 0) cmbSourceSheet.SelectedIndex = 0;
                }
                catch { /* 取不到绑定时留空，用户可手填 */ }
            }

            private string GetSelectedTargetWorkbookPath()
            {
                if (smartOnly)
                {
                    return String.IsNullOrWhiteSpace(smartPreviewWorkbookPath)
                        ? smartActiveWorkbookPath
                        : smartPreviewWorkbookPath;
                }
                OpenSpreadsheetWorkbookInfo selected = cmbTargetWorkbook.SelectedItem as OpenSpreadsheetWorkbookInfo;
                return selected == null ? "" : selected.FullName ?? "";
            }

            private bool TryResolveSmartActiveWorkbook(out string workbookPath, out string error)
            {
                workbookPath = "";
                error = "";
                List<OpenSpreadsheetWorkbookInfo> workbooks;
                string listError;
                if (!TryListOpenSpreadsheetWorkbooks(out workbooks, out listError))
                {
                    error = String.IsNullOrWhiteSpace(listError) ? "无法读取当前活动的 Excel/WPS 工作簿。" : listError;
                    return false;
                }

                List<OpenSpreadsheetWorkbookInfo> active = (workbooks ?? new List<OpenSpreadsheetWorkbookInfo>())
                    .Where(item => item != null && item.IsActive)
                    .ToList();
                if (active.Count == 0)
                {
                    smartActiveWorkbookPath = "";
                    cmbTargetSheet.Items.Clear();
                    cmbTargetSheet.Text = "";
                    error = "没有检测到当前活动的 Excel/WPS 工作簿，请先切换到已保存的工程量工作簿。";
                    return false;
                }
                if (active.Count > 1)
                {
                    smartActiveWorkbookPath = "";
                    cmbTargetSheet.Items.Clear();
                    cmbTargetSheet.Text = "";
                    error = "同时检测到多个活动的 Excel/WPS 工作簿，无法确定推荐取数来源，请只保留一个活动工作簿后重试。";
                    return false;
                }

                OpenSpreadsheetWorkbookInfo selected = active[0];
                string fullName = (selected.FullName ?? "").Trim();
                if (fullName.Length == 0 || !File.Exists(fullName))
                {
                    smartActiveWorkbookPath = "";
                    cmbTargetSheet.Items.Clear();
                    cmbTargetSheet.Text = "";
                    error = "当前活动工作簿尚未保存到磁盘，请保存后再预览推荐定额。";
                    return false;
                }

                string keepSheet = cmbTargetSheet.Text;
                if (selected.SheetNames == null || selected.SheetNames.Count == 0)
                {
                    string sheetError;
                    selected.SheetNames = GetTemplateFillSheetNames(fullName, out sheetError).ToList();
                    if (selected.SheetNames.Count == 0 && !String.IsNullOrWhiteSpace(sheetError))
                    {
                        error = "读取当前活动工作簿的 sheet 失败：" + sheetError;
                        return false;
                    }
                }

                cmbTargetSheet.Items.Clear();
                foreach (string sheetName in selected.SheetNames) cmbTargetSheet.Items.Add(sheetName);
                string targetSheet = selected.SheetNames.FirstOrDefault(sheetName =>
                    String.Equals(sheetName, keepSheet, StringComparison.OrdinalIgnoreCase));
                if (String.IsNullOrWhiteSpace(targetSheet) && !String.IsNullOrWhiteSpace(selected.ActiveSheetName))
                {
                    targetSheet = selected.SheetNames.FirstOrDefault(sheetName =>
                        String.Equals(sheetName, selected.ActiveSheetName, StringComparison.OrdinalIgnoreCase));
                }
                if (String.IsNullOrWhiteSpace(targetSheet)) targetSheet = selected.SheetNames.FirstOrDefault();
                cmbTargetSheet.Text = targetSheet ?? "";
                smartActiveWorkbookPath = fullName;
                workbookPath = fullName;
                return true;
            }

            private static string BuildSmartLearningScopeText(SmartLearningScope scope)
            {
                if (scope == null || String.Equals(scope.Kind, "All", StringComparison.OrdinalIgnoreCase)) return "全部学习库";
                if (String.Equals(scope.Kind, "Unclassified", StringComparison.OrdinalIgnoreCase)) return "未归类";
                string code = (scope.EntryCode ?? "").Trim();
                string name = (scope.DisplayName ?? "").Trim();
                if (code.Length == 2 && code.All(Char.IsDigit))
                {
                    int value;
                    return Int32.TryParse(code, out value)
                        ? ToChineseOrdinal(value) + "、" + (name.Length == 0 ? code : name)
                        : code + (name.Length == 0 ? "" : " " + name);
                }
                if (code.Length == 4 && code.All(Char.IsDigit))
                {
                    return code.Substring(2, 2) + "." + (name.Length == 0 ? code : name);
                }
                return code + (name.Length == 0 ? "" : " " + name);
            }

            private void ReloadSmartLearningScopeTree()
            {
                if (!smartOnly) return;
                string keepKind = selectedSmartLearningScope == null ? "All" : selectedSmartLearningScope.Kind ?? "All";
                string keepCode = selectedSmartLearningScope == null ? "" : selectedSmartLearningScope.EntryCode ?? "";
                rebuildingSmartLearningScopeTree = true;
                try
                {
                    smartLearningScopeTree.BeginUpdate();
                    smartLearningScopeTree.Nodes.Clear();
                    SmartLearningScope allScope = SmartLearningScope.CreateAll();
                    TreeNode allNode = new TreeNode("全部学习库") { Tag = allScope };
                    smartLearningScopeTree.Nodes.Add(allNode);

                    List<SmartLearningScope> scopes;
                    try
                    {
                        SmartScopeLoadResult scopeLoad = LoadSmartLearningScopes(mainForm);
                        scopes = scopeLoad.Scopes;
                        smartLoadStatus = scopeLoad.Status;
                        smartScopeNotice = scopeLoad.Message ?? "";
                        if (scopeLoad.Status != SmartScopeLoadStatus.Success)
                        {
                            btnSmartLearningScope.Text = scopeLoad.Message ?? "学习库不可用";
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("Load smart learning scopes failed: " + ex.Message);
                        smartLoadStatus = SmartScopeLoadStatus.Error;
                        smartScopeNotice = "学习库不可用，未生成推荐。";
                        scopes = new List<SmartLearningScope>();
                    }

                    SmartLearningScope unclassified = scopes.FirstOrDefault(scope => scope != null &&
                        String.Equals(scope.Kind, "Unclassified", StringComparison.OrdinalIgnoreCase));
                    List<SmartLearningScope> entries = scopes
                        .Where(scope => scope != null && String.Equals(scope.Kind, "Entry", StringComparison.OrdinalIgnoreCase) &&
                            !String.IsNullOrWhiteSpace(scope.EntryCode))
                        .GroupBy(scope => scope.EntryCode, StringComparer.OrdinalIgnoreCase)
                        .Select(group => group.First())
                        .OrderBy(scope => scope.EntryCode, StringComparer.OrdinalIgnoreCase)
                        .ToList();
                    Dictionary<string, TreeNode> professionNodes = new Dictionary<string, TreeNode>(StringComparer.OrdinalIgnoreCase);
                    Dictionary<string, TreeNode> divisionNodes = new Dictionary<string, TreeNode>(StringComparer.OrdinalIgnoreCase);
                    foreach (SmartLearningScope scope in entries.Where(item => item.EntryCode.Length == 2 && item.EntryCode.All(Char.IsDigit)))
                    {
                        TreeNode node = new TreeNode(BuildSmartLearningScopeText(scope)) { Tag = scope };
                        smartLearningScopeTree.Nodes.Add(node);
                        professionNodes[scope.EntryCode] = node;
                    }
                    foreach (SmartLearningScope scope in entries.Where(item => item.EntryCode.Length == 4 && item.EntryCode.All(Char.IsDigit)))
                    {
                        string professionCode = scope.EntryCode.Substring(0, 2);
                        TreeNode parent;
                        if (!professionNodes.TryGetValue(professionCode, out parent)) continue;
                        TreeNode node = new TreeNode(BuildSmartLearningScopeText(scope)) { Tag = scope };
                        parent.Nodes.Add(node);
                        divisionNodes[scope.EntryCode] = node;
                    }
                    if (unclassified != null)
                    {
                        smartLearningScopeTree.Nodes.Add(new TreeNode(BuildSmartLearningScopeText(unclassified)) { Tag = unclassified });
                    }

                    TreeNode selectedNode = null;
                    foreach (TreeNode rootNode in smartLearningScopeTree.Nodes)
                    {
                        selectedNode = FindSmartLearningScopeNode(rootNode, keepKind, keepCode);
                        if (selectedNode != null) break;
                    }
                    if (selectedNode == null) selectedNode = allNode;
                    selectedSmartLearningScope = selectedNode.Tag as SmartLearningScope ?? allScope;
                    smartLearningScopeTree.SelectedNode = selectedNode;
                    btnSmartLearningScope.Text = smartLoadStatus == SmartScopeLoadStatus.Success
                        ? BuildSmartLearningScopeText(selectedSmartLearningScope)
                        : (String.IsNullOrWhiteSpace(smartScopeNotice) ? "学习库不可用" : smartScopeNotice);
                }
                finally
                {
                    smartLearningScopeTree.EndUpdate();
                    rebuildingSmartLearningScopeTree = false;
                    RefreshApplyEnabled();
                }
            }

            private static TreeNode FindSmartLearningScopeNode(TreeNode node, string kind, string entryCode)
            {
                if (node == null) return null;
                SmartLearningScope scope = node.Tag as SmartLearningScope;
                if (scope != null &&
                    String.Equals(scope.Kind ?? "", kind ?? "", StringComparison.OrdinalIgnoreCase) &&
                    String.Equals(scope.EntryCode ?? "", entryCode ?? "", StringComparison.OrdinalIgnoreCase))
                {
                    return node;
                }
                foreach (TreeNode child in node.Nodes)
                {
                    TreeNode found = FindSmartLearningScopeNode(child, kind, entryCode);
                    if (found != null) return found;
                }
                return null;
            }

            private void ShowSmartLearningScopeDropDown()
            {
                ReloadSmartLearningScopeTree();
                if (smartLearningScopeHost == null)
                {
                    smartLearningScopeTree.Size = new Size(420, 360);
                    smartLearningScopeHost = new ToolStripControlHost(smartLearningScopeTree)
                    {
                        AutoSize = false,
                        Margin = Padding.Empty,
                        Padding = Padding.Empty,
                        Size = smartLearningScopeTree.Size
                    };
                    smartLearningScopeDropDown.Padding = Padding.Empty;
                    smartLearningScopeDropDown.Items.Add(smartLearningScopeHost);
                }
                smartLearningScopeDropDown.Show(btnSmartLearningScope, new Point(0, btnSmartLearningScope.Height));
            }

            private void OnSmartLearningScopeSelected(TreeNode node)
            {
                if (rebuildingSmartLearningScopeTree || node == null) return;
                SmartLearningScope scope = node.Tag as SmartLearningScope;
                if (scope == null) return;
                bool changed = selectedSmartLearningScope == null ||
                    !String.Equals(selectedSmartLearningScope.Kind ?? "", scope.Kind ?? "", StringComparison.OrdinalIgnoreCase) ||
                    !String.Equals(selectedSmartLearningScope.EntryCode ?? "", scope.EntryCode ?? "", StringComparison.OrdinalIgnoreCase);
                selectedSmartLearningScope = scope;
                btnSmartLearningScope.Text = BuildSmartLearningScopeText(scope);
                smartLearningScopeDropDown.Close();
                if (changed) InvalidateSmartPreview();
            }

            private void AddTemplateSourceWorkbook(List<OpenSpreadsheetWorkbookInfo> workbooks)
            {
                if (cmbTemplate.SelectedItem == null) return;
                FillTemplate template;
                try { template = LoadFillTemplate(Convert.ToString(cmbTemplate.SelectedItem)); }
                catch { return; }
                if (template == null || String.IsNullOrWhiteSpace(template.WorkbookPath) || !File.Exists(template.WorkbookPath)) return;

                string fullName;
                try { fullName = Path.GetFullPath(template.WorkbookPath); }
                catch { return; }
                OpenSpreadsheetWorkbookInfo info = workbooks.FirstOrDefault(item =>
                    String.Equals(item.FullName, fullName, StringComparison.OrdinalIgnoreCase));
                if (info == null)
                {
                    info = new OpenSpreadsheetWorkbookInfo();
                    info.FullName = fullName;
                    string sheetError;
                    info.SheetNames = GetTemplateFillSheetNames(fullName, out sheetError).ToList();
                    workbooks.Add(info);
                }
                info.IsTemplateSource = true;
                info.DisplayName = BuildOpenWorkbookDisplayName(fullName, true);
            }

            private void ReloadTargetWorkbooks()
            {
                if (reloadingTargetWorkbooks) return;
                targetWorkbookReloadCount++;
                string keepSheet = cmbTargetSheet.Text;
                reloadingTargetWorkbooks = true;
                try
                {
                    string keepPath = GetSelectedTargetWorkbookPath();
                    List<OpenSpreadsheetWorkbookInfo> workbooks;
                    string error;
                    TryListOpenSpreadsheetWorkbooks(out workbooks, out error);
                    AddTemplateSourceWorkbook(workbooks);

                    cmbTargetWorkbook.Items.Clear();
                    foreach (OpenSpreadsheetWorkbookInfo workbook in workbooks)
                    {
                        cmbTargetWorkbook.Items.Add(workbook);
                    }

                    OpenSpreadsheetWorkbookInfo selected = workbooks.FirstOrDefault(item =>
                        String.Equals(item.FullName, keepPath, StringComparison.OrdinalIgnoreCase));
                    if (selected == null) selected = workbooks.FirstOrDefault(item => item.IsActive);
                    if (selected == null) selected = workbooks.FirstOrDefault(item => item.IsTemplateSource);
                    if (selected == null) selected = workbooks.FirstOrDefault();
                    if (selected != null) cmbTargetWorkbook.SelectedItem = selected;
                    else
                    {
                        cmbTargetSheet.Items.Clear();
                        cmbTargetSheet.Text = "";
                        targetWorkbookToolTip.SetToolTip(cmbTargetWorkbook, error ?? "");
                    }
                }
                catch (Exception ex)
                {
                    Log("Reload target workbooks failed: " + ex);
                    targetWorkbookToolTip.SetToolTip(cmbTargetWorkbook, "读取 Excel/WPS 工作簿失败：" + ex.Message);
                }
                finally
                {
                    reloadingTargetWorkbooks = false;
                    ReloadTargetSheets(keepSheet);
                }
            }

            // 目标 sheet 只跟随当前明确选中的目标 Excel，不再根据活动工作簿猜测。
            private void ReloadTargetSheets()
            {
                ReloadTargetSheets(cmbTargetSheet.Text);
            }

            private void ReloadTargetSheets(string keep)
            {
                try
                {
                    OpenSpreadsheetWorkbookInfo selected = cmbTargetWorkbook.SelectedItem as OpenSpreadsheetWorkbookInfo;
                    cmbTargetSheet.Items.Clear();
                    if (selected == null)
                    {
                        cmbTargetSheet.Text = "";
                        return;
                    }

                    if (selected.SheetNames == null || selected.SheetNames.Count == 0)
                    {
                        string sheetError;
                        selected.SheetNames = GetTemplateFillSheetNames(selected.FullName, out sheetError).ToList();
                    }
                    foreach (string sheetName in selected.SheetNames) cmbTargetSheet.Items.Add(sheetName);
                    targetWorkbookToolTip.SetToolTip(cmbTargetWorkbook, selected.FullName ?? "");

                    string target = selected.SheetNames.FirstOrDefault(sheetName =>
                        String.Equals(sheetName, keep, StringComparison.OrdinalIgnoreCase));
                    if (String.IsNullOrWhiteSpace(target) && !String.IsNullOrWhiteSpace(selected.ActiveSheetName))
                    {
                        target = selected.SheetNames.FirstOrDefault(sheetName =>
                            String.Equals(sheetName, selected.ActiveSheetName, StringComparison.OrdinalIgnoreCase));
                    }
                    if (String.IsNullOrWhiteSpace(target)) target = selected.SheetNames.FirstOrDefault();
                    cmbTargetSheet.Text = target ?? "";
                }
                catch (Exception ex)
                {
                    Log("Reload target sheets failed: " + ex);
                    targetWorkbookToolTip.SetToolTip(cmbTargetWorkbook, "读取 Excel/WPS 工作表失败：" + ex.Message);
                }
            }

            // 目标单元 下拉展开时刷新：列出项目全部单元（_ZGS_ 编号，纯编号，供 ResolveAgentUnitIdSimple 精确匹配）。
            private void ReloadTargetUnits()
            {
                try
                {
                    string keep = cmbTargetUnit.Text;
                    List<string> units = ListAgentUnits(mainForm);
                    if (units.Count > 0)
                    {
                        cmbTargetUnit.Items.Clear();
                        foreach (string u in units) cmbTargetUnit.Items.Add(u);
                    }
                    cmbTargetUnit.Text = keep;
                }
                catch { /* 取不到时留空，用户可手填 */ }
            }

            // 列出项目全部单元(编号 如 _ZGS_01)。
            private static List<string> ListAgentUnits(Form mainForm)
            {
                List<string> result = new List<string>();
                try
                {
                    SqlConnection conn = GetOpenProjectConnection(mainForm);
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "select 总概算编号 from 总概算信息 order by 总概算序号";
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string code = reader.IsDBNull(0) ? "" : Convert.ToString(reader.GetValue(0)).Trim();
                                if (code.Length > 0) result.Add(code);
                            }
                        }
                    }
                }
                catch { /* 取不到时留空，用户可手填 */ }
                return result;
            }

            private void OnDeleteTemplate()
            {
                try
                {
                    if (cmbTemplate.SelectedItem == null) { MessageBox.Show(this, "请先选择要删除的模板。", "模板铺量"); return; }
                    string name = Convert.ToString(cmbTemplate.SelectedItem);
                    if (MessageBox.Show(this, "确认删除模板「" + name + "」？此操作不可撤销。",
                        "模板铺量", MessageBoxButtons.OKCancel) != DialogResult.OK) return;
                    DeleteFillTemplate(name);
                    ReloadTemplateList();
                }
                catch (Exception ex) { MessageBox.Show(this, "删除失败：" + ex.Message, "模板铺量"); }
            }

            private void OnBuild()
            {
                try
                {
                    // 借用宿主当前项目连接，不得在插件中关闭或释放。
                    int count;
                    List<string> warnings;
                    SqlConnection conn = GetOpenProjectConnection(mainForm);
                    FillTemplate existing = LoadFillTemplate(txtName.Text.Trim());
                    FillTemplate t = chkNameMode.Checked
                        ? BuildNameFillTemplateFromBindings(mainForm, conn, txtName.Text.Trim(), txtUnit.Text.Trim(), cmbSourceSheet.Text.Trim())
                        : BuildFillTemplateFromBindings(mainForm, conn, txtName.Text.Trim(), txtUnit.Text.Trim(), cmbSourceSheet.Text.Trim());
                    t = MergeRegeneratedFillTemplate(existing, t);
                    count = t.Rows.Count;
                    warnings = t.BuildWarnings;
                    SaveFillTemplate(t);
                    ReloadTemplateList();
                    string msg = count > 0
                        ? ("模板已生成并保存：" + count + " 条定额。")
                        : ("模板已生成，但收到 0 条定额。\n请确认该单元的定额已用“绑定Excel工程量”绑到 sheet「" + cmbSourceSheet.Text.Trim() + "」。");
                    if (warnings != null && warnings.Count > 0)
                    {
                        msg += "\n\n以下绑定被跳过（不属于源单元 " + txtUnit.Text.Trim() + "）：\n" + String.Join("\n", warnings.ToArray());
                    }
                    MessageBox.Show(this, msg, "模板铺量");
                }
                catch (Exception ex) { MessageBox.Show(this, "生成失败：" + ex.Message, "模板铺量"); }
            }

            private void OnPreview()
            {
                if (smartOnly) InvalidateSmartPreview();
                SetBusy(true, "预览中...");
                try
                {
                    // 模式三·推荐定额:不需要模板,直接走学习库漏斗。
                    if (cmbMode.SelectedIndex == 2)
                    {
                        string smartWorkbook;
                        string smartWorkbookError;
                        updatingSmartPreviewInputs = true;
                        bool workbookReady;
                        try { workbookReady = TryResolveSmartActiveWorkbook(out smartWorkbook, out smartWorkbookError); }
                        finally { updatingSmartPreviewInputs = false; }
                        if (!workbookReady)
                        {
                            MessageBox.Show(this, smartWorkbookError, "推荐定额");
                            return;
                        }
                        smartPreviewWorkbookPath = smartWorkbook;
                        string smartWarning = null;
                        string softwarePartition;
                        string normalizedMethodNo;
                        preview = BuildPreview_SmartFill(mainForm, smartWorkbook, cmbTargetSheet.Text.Trim(), txtColumn.Text.Trim(),
                            selectedSmartLearningScope, out smartWarning, out smartLoadStatus, out softwarePartition, out normalizedMethodNo);
                        if (!String.IsNullOrWhiteSpace(smartScopeNotice) && smartLoadStatus == SmartScopeLoadStatus.Success)
                            smartWarning = String.IsNullOrWhiteSpace(smartWarning) ? smartScopeNotice : smartWarning + "\n" + smartScopeNotice;
                        if (!String.IsNullOrEmpty(smartWarning)) MessageBox.Show(this, smartWarning, "推荐定额");
                        RefreshCurrentSmartEntry(false);
                        smartPreviewReady = smartLoadStatus == SmartScopeLoadStatus.Success && preview.Count > 0;
                        if (smartPreviewReady)
                        {
                            SqlConnection conn = GetOpenProjectConnection(mainForm);
                            AgentSelectionSnapshot selection = CaptureAgentSelection(mainForm);
                            smartPreviewVersion++;
                            previewContext = new SmartPreviewContext
                            {
                                ProjectConnection = conn,
                                ProjectConnectionIdentity = GetProjectConnectionIdentity(conn),
                                CurrentUnitId = selection == null ? 0 : selection.CurrentUnitId,
                                CurrentUnitCode = selection == null ? "" : selection.CurrentUnitCode ?? "",
                                SoftwarePartition = softwarePartition ?? "",
                                MethodNo = normalizedMethodNo ?? "",
                                ScopeKind = selectedSmartLearningScope == null ? "All" : selectedSmartLearningScope.Kind ?? "All",
                                ScopeEntryCode = selectedSmartLearningScope == null ? "" : selectedSmartLearningScope.EntryCode ?? "",
                                WorkbookPath = Path.GetFullPath(smartWorkbook),
                                Worksheet = cmbTargetSheet.Text.Trim(),
                                TargetColumn = txtColumn.Text.Trim().ToUpperInvariant(),
                                PreviewVersion = smartPreviewVersion
                            };
                        }
                        FillGrid();
                        return;
                    }

                    if (cmbTemplate.SelectedItem == null) { MessageBox.Show(this, "请先选择模板。", "模板铺量"); return; }
                    FillTemplate t = LoadFillTemplate(Convert.ToString(cmbTemplate.SelectedItem));
                    if (t == null) { MessageBox.Show(this, "模板加载失败。", "模板铺量"); return; }
                    string targetWorkbook = GetSelectedTargetWorkbookPath();
                    if (String.IsNullOrWhiteSpace(targetWorkbook) || !File.Exists(targetWorkbook))
                    {
                        MessageBox.Show(this, "没有可用的目标 Excel，请先打开并保存工作簿。", "模板铺量");
                        return;
                    }
                    string ndWarning = null;
                    preview = cmbMode.SelectedIndex == 0
                        ? BuildPreview_ColumnAnchor(t, targetWorkbook, cmbTargetSheet.Text.Trim(), txtColumn.Text.Trim())
                        : BuildPreview_NameDriven(mainForm, t, targetWorkbook, cmbTargetSheet.Text.Trim(), txtColumn.Text.Trim(), out ndWarning);
                    if (!String.IsNullOrEmpty(ndWarning)) MessageBox.Show(this, ndWarning, "模板铺量");
                    RebuildItemTree();
                    FillGrid();
                    if (preview.Count == 0)
                        MessageBox.Show(this, "预览为空：该模板里没有定额。请回到上一步重新“从该单元生成模板”，并确认收到的定额条数大于 0。", "模板铺量");
                }
                catch (Exception ex) { MessageBox.Show(this, "预览失败：" + ex.Message, "模板铺量"); }
                finally { SetBusy(false, ""); }
            }

            // 右侧表格：当前树节点范围内的定额平铺（“调整”列不显示，写入时仍随源行完整复制）。
            private void FillGrid()
            {
                bool previousUpdating = updatingNameQuotaCell;
                updatingNameQuotaCell = true;
                try
                {
                    grid.Rows.Clear();
                    string scope = currentTreeScope ?? "";
                    bool nameDriven = preview.Any(p => p.IsNameDriven);
                    Dictionary<int, FillPreviewItem> nameLeaders = preview
                        .Where(item => item != null && item.IsNameDriven && item.GroupOrder == 0)
                        .GroupBy(item => item.TargetRow)
                        .ToDictionary(group => group.Key, group => group.First());
                    HashSet<int> visibleNameGroups = new HashSet<int>(preview
                        .Where(item => item != null && item.IsNameDriven)
                        .GroupBy(item => item.TargetRow)
                        .Where(group => String.IsNullOrEmpty(scope) ||
                            group.Any(item => IsItemNoUnderChapter(item.ItemNo ?? "", scope)))
                        .Select(group => group.Key));
                    IEnumerable<FillPreviewItem> ordered = preview.Where(item => item != null &&
                        (item.IsNameDriven
                            ? visibleNameGroups.Contains(item.TargetRow)
                            : String.IsNullOrEmpty(scope) || IsItemNoUnderChapter(item.ItemNo ?? "", scope)));
                    ordered = nameDriven
                        ? ordered.OrderBy(item => item.TargetRow).ThenBy(item => item.GroupOrder)
                        : ordered.OrderBy(item => item.ItemNo ?? "", StringComparer.OrdinalIgnoreCase).ThenBy(item => item.OrderInItem);
                    foreach (FillPreviewItem item in ordered)
                    {
                        int index = grid.Rows.Add();
                        FillPreviewItem leader;
                        nameLeaders.TryGetValue(item.TargetRow, out leader);
                        SetGridRow(grid.Rows[index], item, leader);
                    }
                    grid.ClearSelection();
                }
                finally
                {
                    updatingNameQuotaCell = previousUpdating;
                }
                UpdateSmartWriteScope();
                RefreshApplyEnabled();
            }

            private void SetGridRow(DataGridViewRow row, FillPreviewItem item, FillPreviewItem leader)
            {
                int selectIndex = grid.Columns["sel"].Index;
                int codeIndex = grid.Columns["code"].Index;
                int sourceNameIndex = grid.Columns["sname"].Index;
                bool isGroupMember = item.IsNameDriven && item.GroupOrder > 0;
                if (isGroupMember && !(row.Cells[selectIndex] is DataGridViewTextBoxCell))
                {
                    row.Cells[selectIndex] = new DataGridViewTextBoxCell();
                }
                else if (!isGroupMember && !(row.Cells[selectIndex] is DataGridViewCheckBoxCell))
                {
                    row.Cells[selectIndex] = new DataGridViewCheckBoxCell();
                }
                if (row.Cells[codeIndex] is DataGridViewComboBoxCell)
                {
                    row.Cells[codeIndex] = new DataGridViewTextBoxCell();
                }
                if (row.Cells[sourceNameIndex] is DataGridViewComboBoxCell)
                {
                    row.Cells[sourceNameIndex] = new DataGridViewTextBoxCell();
                }

                string statusText = String.IsNullOrEmpty(item.Status) ? (item.AlignNote ?? "") : item.Status;
                row.SetValues(isGroupMember ? (object)"" : item.Selected, item.QuotaCode, item.SourceName, item.Unit ?? "",
                    item.TargetName, item.QuantityText, statusText);
                row.Tag = item;
                row.Cells["sel"].ReadOnly = isGroupMember;
                row.Cells["sel"].ToolTipText = "";
                row.Cells["code"].ToolTipText = "";
                row.Cells["sname"].ToolTipText = "";
                row.DefaultCellStyle.BackColor = Color.Empty;

                bool hasCandidates = item.GroupOrder == 0 && item.NameQuotaCandidates != null &&
                    item.NameQuotaCandidates.Count > 1;
                bool requiresChoice = leader != null && leader.NeedExactNameConfirmation &&
                    leader.NameQuotaCandidates != null && leader.NameQuotaCandidates.Count > 1;
                string blockingReason = item.IsNameDriven
                    ? GetUnsafeNameQuotaCandidateReason(preview.Where(candidate => candidate != null &&
                        candidate.IsNameDriven && candidate.TargetRow == item.TargetRow))
                    : "";
                row.Cells["code"].ReadOnly = true;
                row.Cells["sname"].ReadOnly = !hasCandidates;
                row.Cells["qty"].ReadOnly = false;
                row.Cells["sel"].ReadOnly = isGroupMember;
                if (hasCandidates)
                {
                    row.Cells["sname"].ToolTipText = "点击选择该工程量名称绑定的源行定额或组件组";
                }
                if (requiresChoice)
                {
                    if (isGroupMember)
                    {
                        row.Cells["sel"].ToolTipText = "请在组首勾选接受当前候选，或在源行定额列切换绑定组";
                    }
                    else
                    {
                        row.Cells["sel"].ToolTipText = "勾选接受当前候选；点击源行定额可切换绑定组";
                    }
                }
                if (!String.IsNullOrWhiteSpace(blockingReason))
                {
                    row.Cells["sel"].ToolTipText = "该组件存在风险：" + blockingReason + "。请先修改数量或处理单位、条目问题。";
                }
                row.DefaultCellStyle.BackColor = GetNameQuotaRowBackColor(item);
            }

            private sealed class TemplateFillGridViewState
            {
                public int FirstDisplayedIndex = -1;
                public bool HasTopKey;
                public int TopTargetRow;
                public int TopGroupOrder;
                public bool HasCurrentKey;
                public int CurrentTargetRow;
                public int CurrentGroupOrder;
                public int CurrentColumnIndex = -1;
                public int HorizontalOffset;
            }

            private TemplateFillGridViewState CaptureGridViewState()
            {
                TemplateFillGridViewState state = new TemplateFillGridViewState();
                state.HorizontalOffset = grid.HorizontalScrollingOffset;
                if (grid.Rows.Count == 0) return state;

                try { state.FirstDisplayedIndex = grid.FirstDisplayedScrollingRowIndex; }
                catch { state.FirstDisplayedIndex = -1; }
                if (state.FirstDisplayedIndex >= 0 && state.FirstDisplayedIndex < grid.Rows.Count)
                {
                    FillPreviewItem top = grid.Rows[state.FirstDisplayedIndex].Tag as FillPreviewItem;
                    if (top != null)
                    {
                        state.HasTopKey = true;
                        state.TopTargetRow = top.TargetRow;
                        state.TopGroupOrder = top.GroupOrder;
                    }
                }
                if (grid.CurrentCell != null && grid.CurrentCell.RowIndex >= 0)
                {
                    FillPreviewItem current = grid.Rows[grid.CurrentCell.RowIndex].Tag as FillPreviewItem;
                    if (current != null)
                    {
                        state.HasCurrentKey = true;
                        state.CurrentTargetRow = current.TargetRow;
                        state.CurrentGroupOrder = current.GroupOrder;
                        state.CurrentColumnIndex = grid.CurrentCell.ColumnIndex;
                    }
                }
                return state;
            }

            private int FindGridRowIndex(int targetRow, int groupOrder)
            {
                foreach (DataGridViewRow row in grid.Rows)
                {
                    FillPreviewItem item = row.Tag as FillPreviewItem;
                    if (item != null && item.TargetRow == targetRow && item.GroupOrder == groupOrder)
                        return row.Index;
                }
                return -1;
            }

            private void RestoreGridViewState(TemplateFillGridViewState state)
            {
                if (state == null || grid.Rows.Count == 0) return;

                int currentIndex = state.HasCurrentKey
                    ? FindGridRowIndex(state.CurrentTargetRow, state.CurrentGroupOrder) : -1;
                if (currentIndex < 0 && state.HasCurrentKey)
                    currentIndex = FindGridRowIndex(state.CurrentTargetRow, 0);
                if (currentIndex >= 0 && state.CurrentColumnIndex >= 0 &&
                    state.CurrentColumnIndex < grid.Columns.Count)
                    grid.CurrentCell = grid.Rows[currentIndex].Cells[state.CurrentColumnIndex];

                int topIndex = state.HasTopKey ? FindGridRowIndex(state.TopTargetRow, state.TopGroupOrder) : -1;
                if (topIndex < 0 && state.FirstDisplayedIndex >= 0)
                    topIndex = Math.Min(state.FirstDisplayedIndex, grid.Rows.Count - 1);
                if (topIndex >= 0)
                {
                    try { grid.FirstDisplayedScrollingRowIndex = topIndex; }
                    catch { }
                }
                try { grid.HorizontalScrollingOffset = Math.Max(0, state.HorizontalOffset); }
                catch { }
            }

            private bool RefreshTargetGroupInGrid(int targetRow)
            {
                List<DataGridViewRow> existing = grid.Rows.Cast<DataGridViewRow>()
                    .Where(row =>
                    {
                        FillPreviewItem item = row.Tag as FillPreviewItem;
                        return item != null && item.IsNameDriven && item.TargetRow == targetRow;
                    })
                    .OrderBy(row => row.Index)
                    .ToList();
                List<FillPreviewItem> desired = preview
                    .Where(item => item != null && item.IsNameDriven && item.TargetRow == targetRow)
                    .OrderBy(item => item.GroupOrder)
                    .ToList();
                if (existing.Count == 0 || desired.Count == 0) return false;

                TemplateFillGridViewState state = CaptureGridViewState();
                int insertAt = existing[0].Index;
                bool previousUpdating = updatingNameQuotaCell;
                updatingNameQuotaCell = true;
                try
                {
                    grid.EndEdit();
                    for (int i = existing.Count - 1; i >= desired.Count; i--)
                        grid.Rows.RemoveAt(insertAt + i);
                    for (int i = existing.Count; i < desired.Count; i++)
                        grid.Rows.Insert(insertAt + i, false, "", "", "", "", "", "");

                    FillPreviewItem leader = desired.FirstOrDefault(item => item.GroupOrder == 0) ?? desired[0];
                    for (int i = 0; i < desired.Count; i++)
                        SetGridRow(grid.Rows[insertAt + i], desired[i], leader);
                }
                finally
                {
                    updatingNameQuotaCell = previousUpdating;
                }
                RestoreGridViewState(state);
                return true;
            }

            private bool PrepareNameQuotaDropDown(DataGridViewRow row)
            {
                FillPreviewItem item = row == null ? null : row.Tag as FillPreviewItem;
                if (item == null || item.GroupOrder != 0 || item.NameQuotaCandidates == null ||
                    item.NameQuotaCandidates.Count <= 1) return false;
                if (row.Cells["sname"] is DataGridViewComboBoxCell) return true;

                DataGridViewComboBoxCell combo = new DataGridViewComboBoxCell();
                combo.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton;
                combo.DropDownWidth = 500;
                foreach (NameQuotaCandidateGroup option in item.NameQuotaCandidates)
                {
                    combo.Items.Add(option.Label);
                }
                NameQuotaCandidateGroup current = item.NameQuotaCandidates.FirstOrDefault(option =>
                    String.Equals(option.Key, item.SelectedNameQuotaCandidateKey, StringComparison.Ordinal));
                combo.Value = (current ?? item.NameQuotaCandidates[0]).Label;

                updatingNameQuotaCell = true;
                try { row.Cells[grid.Columns["sname"].Index] = combo; }
                finally { updatingNameQuotaCell = false; }
                return true;
            }

            private void OnNameQuotaSelectionCommitted(object sender, EventArgs e)
            {
                ComboBox combo = sender as ComboBox;
                DataGridViewRow row = grid.CurrentCell == null ? grid.CurrentRow : grid.CurrentCell.OwningRow;
                if (combo == null || row == null) return;
                ApplyNameQuotaOption(row, Convert.ToString(combo.SelectedItem));
            }

            private void ApplyNameQuotaOption(DataGridViewRow row, string label)
            {
                if (updatingNameQuotaCell) return;
                FillPreviewItem item = row == null ? null : row.Tag as FillPreviewItem;
                if (item == null || item.NameQuotaCandidates == null || item.NameQuotaCandidates.Count <= 1) return;
                NameQuotaCandidateGroup option = item.NameQuotaCandidates.FirstOrDefault(candidate =>
                    String.Equals(candidate.Label, label, StringComparison.Ordinal));
                if (option == null) return;
                updatingNameQuotaCell = true;
                try
                {
                    int targetRow = item.TargetRow;
                    if (ApplyExactNameCandidate(preview, targetRow, option.Key))
                        RefreshTargetGroupInGrid(targetRow);
                }
                finally { updatingNameQuotaCell = false; }
            }

            private void ConfirmExactNameFromCheck(DataGridViewRow row)
            {
                FillPreviewItem item = row == null ? null : row.Tag as FillPreviewItem;
                bool value = row != null && Convert.ToBoolean(row.Cells["sel"].Value ?? false);
                if (item == null || !value) return;
                List<FillPreviewItem> currentGroup = preview.Where(candidate => candidate != null &&
                    candidate.IsNameDriven && candidate.TargetRow == item.TargetRow).ToList();
                foreach (FillPreviewItem candidate in currentGroup) ConfirmPendingCountUnitScale(candidate);
                if (HasUnsafeNameQuotaCandidate(currentGroup))
                {
                    foreach (FillPreviewItem candidate in currentGroup) candidate.Selected = false;
                    updatingNameQuotaCell = true;
                    try
                    {
                        row.Cells["sel"].Value = false;
                        RefreshTargetGroupInGrid(item.TargetRow);
                        DataGridViewRow firstBlockingRow = grid.Rows.Cast<DataGridViewRow>().FirstOrDefault(candidateRow =>
                        {
                            FillPreviewItem candidate = candidateRow.Tag as FillPreviewItem;
                            return candidate != null && candidate.IsNameDriven && candidate.TargetRow == item.TargetRow &&
                                IsNameQuotaHardStatus(candidate.Status);
                        });
                        if (firstBlockingRow != null)
                        {
                            firstBlockingRow.Cells["qty"].ToolTipText = "请先处理：" + firstBlockingRow.Cells["st"].Value;
                            grid.CurrentCell = firstBlockingRow.Cells["qty"];
                        }
                    }
                    finally { updatingNameQuotaCell = false; }
                    return;
                }

                updatingNameQuotaCell = true;
                try
                {
                    int targetRow = item.TargetRow;
                    if (currentGroup.Any(candidate => candidate.NeedExactNameConfirmation))
                        ConfirmCurrentExactNameGroup(preview, targetRow);
                    else
                        foreach (FillPreviewItem candidate in currentGroup) candidate.Selected = true;
                    RefreshTargetGroupInGrid(targetRow);
                }
                finally { updatingNameQuotaCell = false; }
            }

            private void ApplyNameGroupSelectionFromCheck(DataGridViewRow row)
            {
                FillPreviewItem item = row == null ? null : row.Tag as FillPreviewItem;
                if (item == null || !item.IsNameDriven || item.GroupOrder != 0) return;
                bool value = Convert.ToBoolean(row.Cells["sel"].Value ?? false);
                if (value)
                {
                    ConfirmExactNameFromCheck(row);
                    return;
                }
                foreach (FillPreviewItem member in preview.Where(candidate => candidate != null &&
                    candidate.IsNameDriven && candidate.TargetRow == item.TargetRow))
                {
                    member.Selected = value;
                }
            }

            private static bool HasUnsafeNameQuotaCandidate(IEnumerable<FillPreviewItem> items)
            {
                return items == null || items.Any(item => item == null || IsNameQuotaHardStatus(item.Status) || item.SfEntryBlocked);
            }

            private static string GetUnsafeNameQuotaCandidateReason(IEnumerable<FillPreviewItem> items)
            {
                if (items == null) return "未找到组件数据";
                return String.Join("；", items.Where(item => item != null)
                    .SelectMany(item => SplitNameQuotaStatus(item.Status).Where(part => !IsSoftNameQuotaStatusPart(part))
                        .Concat(item.SfEntryBlocked && !String.IsNullOrWhiteSpace(item.SfEntryBlockReason)
                            ? new[] { item.SfEntryBlockReason }
                            : new string[0]))
                    .Distinct(StringComparer.Ordinal).ToArray());
            }

            private static List<string> SplitNameQuotaStatus(string status)
            {
                return (status ?? "").Split(new[] { '；' }, StringSplitOptions.RemoveEmptyEntries)
                    .Select(part => part.Trim()).Where(part => part.Length > 0).ToList();
            }

            private static bool IsSoftNameQuotaStatusPart(string statusPart)
            {
                return String.Equals((statusPart ?? "").Trim(), "待确认计数单位1:1", StringComparison.Ordinal);
            }

            private static bool IsNameQuotaHardStatus(string status)
            {
                List<string> parts = SplitNameQuotaStatus(status);
                return parts.Any(part => !IsSoftNameQuotaStatusPart(part));
            }

            private static Color GetNameQuotaRowBackColor(FillPreviewItem item)
            {
                if (item == null) return Color.Empty;
                if (item.NeedManualQuota &&
                    String.Equals((item.Status ?? "").Trim(), "未匹配", StringComparison.Ordinal))
                    return Color.FromArgb(255, 246, 196);
                if (IsNameQuotaHardStatus(item.Status) || item.SfEntryBlocked) return Color.MistyRose;
                if (item.NeedExactNameConfirmation || item.NeedManualQuota ||
                    SplitNameQuotaStatus(item.Status).Any(IsSoftNameQuotaStatusPart))
                    return Color.FromArgb(255, 246, 196);
                return Color.Empty;
            }

            private static bool ConfirmPendingCountUnitScale(FillPreviewItem item)
            {
                if (item == null) return false;
                List<string> parts = SplitNameQuotaStatus(item.Status);
                if (!parts.Any(IsSoftNameQuotaStatusPart)) return false;
                string suffix;
                if (!TryBuildConfirmedCountUnitScaleSuffix(item.TargetUnit, item.Unit, out suffix)) return false;

                string quantityBase = String.IsNullOrWhiteSpace(item.TargetQuantityText)
                    ? item.QuantityText
                    : item.TargetQuantityText;
                item.QuantityText = (quantityBase ?? "") + suffix;
                item.Status = String.Join("；", parts.Where(part => !IsSoftNameQuotaStatusPart(part)).ToArray());
                item.AlignNote = AppendPreviewNote(item.AlignNote,
                    "计数单位1:1已确认" + (String.IsNullOrEmpty(suffix) ? "" : "，数量" + suffix));
                return true;
            }

            private static bool ApplyEditedNameQuotaQuantity(FillPreviewItem item, string quantityText)
            {
                if (item == null) return false;
                string editedText = (quantityText ?? "").Trim();
                if (String.Equals(editedText, (item.QuantityText ?? "").Trim(), StringComparison.Ordinal)) return false;
                string quantityBase = (item.TargetQuantityText ?? "").Trim();
                if (editedText.IndexOf("原数量", StringComparison.Ordinal) >= 0)
                {
                    if (String.IsNullOrWhiteSpace(quantityBase)) return false;
                    editedText = editedText.Replace("原数量", "(" + quantityBase + ")");
                }
                item.QuantityText = editedText;
                if (String.IsNullOrWhiteSpace(item.Status)) return false;

                List<string> statusParts = SplitNameQuotaStatus(item.Status);
                if (!statusParts.Any(part => String.Equals(part, "缺跨量纲换算系数", StringComparison.Ordinal) ||
                    String.Equals(part, "公式参数缺失或歧义", StringComparison.Ordinal))) return false;
                decimal quantity;
                string error;
                if (!TryEvaluateDecimal(item.QuantityText, out quantity, out error) || quantity <= 0m) return false;

                item.Status = String.Join("；", statusParts.Where(part =>
                    !String.Equals(part, "缺跨量纲换算系数", StringComparison.Ordinal) &&
                    !String.Equals(part, "公式参数缺失或歧义", StringComparison.Ordinal)).ToArray());
                item.AlignNote = AppendPreviewNote(item.AlignNote, "数量已人工确认");
                return true;
            }

            private void RefreshNameQuotaRiskStateInGrid(int targetRow)
            {
                List<FillPreviewItem> group = preview.Where(item => item != null && item.IsNameDriven &&
                    item.TargetRow == targetRow).ToList();
                string blockingReason = GetUnsafeNameQuotaCandidateReason(group);
                foreach (DataGridViewRow row in grid.Rows)
                {
                    FillPreviewItem item = row.Tag as FillPreviewItem;
                    if (item == null || !item.IsNameDriven || item.TargetRow != targetRow) continue;
                    bool isGroupMember = item.GroupOrder > 0;
                    row.Cells["qty"].ReadOnly = false;
                    row.Cells["sel"].ReadOnly = isGroupMember;
                    row.Cells["sel"].ToolTipText = String.IsNullOrWhiteSpace(blockingReason)
                        ? (item.NeedExactNameConfirmation ? "勾选接受当前候选" : "")
                        : "该组件存在风险：" + blockingReason + "。请先修改数量或处理单位、条目问题。";
                    row.Cells["st"].Value = String.IsNullOrEmpty(item.Status) ? (item.AlignNote ?? "") : item.Status;
                    row.DefaultCellStyle.BackColor = GetNameQuotaRowBackColor(item);
                }
            }

            private static Rectangle GetVisibleMergedTargetNameBounds(DataGridView ownerGrid, int columnIndex,
                int startRow, int endRow)
            {
                if (ownerGrid == null || columnIndex < 0 || columnIndex >= ownerGrid.Columns.Count ||
                    startRow < 0 || endRow < startRow || endRow >= ownerGrid.Rows.Count) return Rectangle.Empty;
                Rectangle result = Rectangle.Empty;
                for (int i = startRow; i <= endRow; i++)
                {
                    if (!ownerGrid.Rows[i].Displayed) continue;
                    Rectangle cellBounds = ownerGrid.GetCellDisplayRectangle(columnIndex, i, true);
                    if (cellBounds.Width <= 0 || cellBounds.Height <= 0) continue;
                    result = result.IsEmpty ? cellBounds : Rectangle.Union(result, cellBounds);
                }
                if (result.IsEmpty) return result;

                Rectangle dataBounds = ownerGrid.ClientRectangle;
                if (ownerGrid.ColumnHeadersVisible)
                {
                    dataBounds.Y += ownerGrid.ColumnHeadersHeight;
                    dataBounds.Height = Math.Max(0, ownerGrid.ClientRectangle.Bottom - dataBounds.Y);
                }
                result.Intersect(dataBounds);
                return result;
            }

            private static void DrawMergedTargetNameTextForCell(Graphics graphics, string text, Font font,
                Rectangle mergedBounds, Rectangle cellBounds, Color foreColor)
            {
                if (graphics == null || font == null || String.IsNullOrEmpty(text) ||
                    mergedBounds.IsEmpty || cellBounds.IsEmpty) return;
                Rectangle clip = Rectangle.Intersect(mergedBounds, cellBounds);
                if (clip.IsEmpty) return;

                System.Drawing.Drawing2D.GraphicsState state = graphics.Save();
                try
                {
                    graphics.IntersectClip(clip);
                    // 靠左 + 垂直居中,与普通行的左对齐保持一致。
                    TextRenderer.DrawText(graphics, text, font, mergedBounds, foreColor,
                        TextFormatFlags.Left | TextFormatFlags.VerticalCenter |
                        TextFormatFlags.WordBreak | TextFormatFlags.EndEllipsis);
                }
                finally
                {
                    graphics.Restore(state);
                }
            }

            // —— 条目树 ——
            // 用章节表的 条目编号+名称 构建层级：只显示编号为两位数字（01、02…）这级及以下，
            // “第一部分”这类更高层级不进树。节点 Tag 为条目编号，根节点 Tag 为空串表示全部。
            private void RebuildItemTree()
            {
                if (smartOnly) return;
                rebuildingTree = true;
                itemTree.BeginUpdate();
                try
                {
                    itemTree.Nodes.Clear();
                    currentTreeScope = "";
                    chapterNames = LoadChapterNameMap(mainForm);

                    TreeNode root = new TreeNode("全部条目");
                    root.Tag = "";
                    itemTree.Nodes.Add(root);

                    Dictionary<string, TreeNode> nodesByCode = new Dictionary<string, TreeNode>(StringComparer.OrdinalIgnoreCase);
                    foreach (string itemNo in preview
                        .Select(it => (it.ItemNo ?? "").Trim())
                        .Where(no => no.Length > 0)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(no => no, StringComparer.OrdinalIgnoreCase))
                    {
                        TreeNode parent = root;
                        foreach (string code in BuildChapterChain(chapterNames, itemNo))
                        {
                            TreeNode node;
                            if (!nodesByCode.TryGetValue(code, out node))
                            {
                                node = new TreeNode(ChapterTreeDisplayName(chapterNames, code));
                                node.Tag = code;
                                node.ToolTipText = code;
                                parent.Nodes.Add(node);
                                nodesByCode[code] = node;
                            }

                            parent = node;
                        }
                    }

                    root.Expand();
                    itemTree.SelectedNode = root;
                }
                finally
                {
                    itemTree.EndUpdate();
                    rebuildingTree = false;
                }
            }

            private void OnTreeScopeChanged()
            {
                if (rebuildingTree)
                {
                    return;
                }

                TreeNode node = itemTree.SelectedNode;
                FlushGridSelectionsToPreview();
                currentTreeScope = node == null ? "" : Convert.ToString(node.Tag);
                FillGrid();
            }

            // 勾选树节点＝整枝勾选/取消该条目（含下级）的全部定额。
            private void OnTreeNodeChecked(TreeNode node)
            {
                if (rebuildingTree || updatingTreeChecks || node == null)
                {
                    return;
                }

                updatingTreeChecks = true;
                try
                {
                    bool value = node.Checked;
                    SetTreeChildrenChecked(node, value);
                    FlushGridSelectionsToPreview();
                    string scope = Convert.ToString(node.Tag);
                    foreach (FillPreviewItem it in preview)
                    {
                        if (String.IsNullOrEmpty(scope) || IsItemNoUnderChapter(it.ItemNo ?? "", scope))
                        {
                            it.Selected = value;
                        }
                    }

                    FillGrid();
                }
                finally
                {
                    updatingTreeChecks = false;
                }
            }

            private void FlushGridSelectionsToPreview()
            {
                grid.EndEdit();
                foreach (DataGridViewRow row in grid.Rows)
                {
                    FillPreviewItem it = row.Tag as FillPreviewItem;
                    if (it == null) continue;
                    ApplyEditedNameQuotaQuantity(it, Convert.ToString(row.Cells["qty"].Value).Trim());
                    if (row.Cells["sel"] is DataGridViewCheckBoxCell)
                    {
                        bool selected = Convert.ToBoolean(row.Cells["sel"].Value ?? false);
                        if (it.IsNameDriven && it.GroupOrder == 0)
                        {
                            foreach (FillPreviewItem member in preview.Where(candidate => candidate != null &&
                                candidate.IsNameDriven && candidate.TargetRow == it.TargetRow))
                            {
                                member.Selected = selected;
                            }
                        }
                        else
                        {
                            it.Selected = selected;
                        }
                    }
                }
            }

            private static string ResolveTemplateFillQuotaUnit(SqlConnection conn, DataGridViewRow row, long quotaSequence)
            {
                string unit = GetRowValue(row, "单位", "定额单位", "计量单位").Trim();
                if (!String.IsNullOrWhiteSpace(unit)) return unit;
                if (conn == null || quotaSequence <= 0) return "";

                try
                {
                    EnsureOpen(conn);
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "select top 1 单位 from 定额输入 where 定额序号=@seq";
                        cmd.Parameters.AddWithValue("@seq", quotaSequence);
                        object value = cmd.ExecuteScalar();
                        return value == null || value == DBNull.Value ? "" : Convert.ToString(value).Trim();
                    }
                }
                catch (Exception ex)
                {
                    Log("ResolveTemplateFillQuotaUnit failed: " + ex.Message);
                    return "";
                }
            }

            private string PromptTemplateCrossUnitFactor(string sourceUnit, string targetUnit, string quotaCode)
            {
                while (true)
                {
                    FactorInfo input = PromptFactor(this, "跨单位换算 " + (quotaCode ?? "") + " " + sourceUnit + " → " + targetUnit);
                    if (input == null) return null;
                    decimal value;
                    if (!Decimal.TryParse(input.Factor, NumberStyles.Float, CultureInfo.InvariantCulture, out value) || value <= 0m)
                    {
                        MessageBox.Show(this, "换算系数必须大于 0。", "跨单位换算", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        continue;
                    }
                    decimal multiplier = input.Operator == "/" ? 1m / value : value;
                    return multiplier.ToString("0.############", CultureInfo.InvariantCulture);
                }
            }

            // 右键：把软件定额输入表当前选中的一行，绑定为该预览行的复制来源（含所在条目）。
            // 注意：与"绑定Excel工程量"同款用主程序共享连接（克隆连接在部分环境登录失败，
            // 会导致 ResolveQuotaSequence 查不到序号）；共享连接不得 using 释放。
            private void OnBindSelectedQuotaToRow()
            {
                try
                {
                    if (grid.SelectedRows.Count == 0) return;
                    FillPreviewItem it = grid.SelectedRows[0].Tag as FillPreviewItem;
                    if (it == null || !it.IsNameDriven) return;
                    List<FillPreviewItem> oldGroup = preview
                        .Where(p => p != null && p.IsNameDriven && p.TargetRow == it.TargetRow)
                        .OrderBy(p => p.GroupOrder)
                        .ToList();
                    FillPreviewItem groupLeader = oldGroup.FirstOrDefault() ?? it;

                    DataGridView de = GetField<DataGridView>(mainForm, "dataGridViewDE");
                    List<DataGridViewRow> rows = GetSelectedQuotaRows(de);
                    if (rows.Count == 0)
                    {
                        DataGridViewRow cur = GetCurrentQuotaRow(de);
                        if (cur != null) rows.Add(cur);
                    }
                    if (rows.Count == 0) { MessageBox.Show(this, "请先在软件定额输入表中选中一条或多条定额行。", "模板铺量"); return; }
                    SqlConnection conn = GetProjectConnection(mainForm);
                    if (conn == null) { MessageBox.Show(this, "没有找到当前项目数据库连接。", "模板铺量"); return; }

                    // 先在临时列表完整验证；任何失败都不修改原组件组。
                    List<FillPreviewItem> replacements = new List<FillPreviewItem>();
                    List<string> errors = new List<string>();
                    Dictionary<string, string> itemNoCache = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    Dictionary<string, string> conversionFactors = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (DataGridViewRow row in rows)
                    {
                        ExcelQuotaLink link; string err;
                        if (!TryCreateQuotaLink(mainForm, conn, row, out link, out err))
                        {
                            errors.Add((link == null ? "" : link.QuotaCode + "：") + err);
                            continue;
                        }

                        long itemSeq;
                        string itemNo = "";
                        if (Int64.TryParse((link.ChapterSeq ?? "").Trim(), out itemSeq) && itemSeq > 0)
                        {
                            itemNo = ResolveChapterItemNo(conn, link.ChapterSeq, itemNoCache);
                        }
                        if (itemSeq <= 0 || String.IsNullOrWhiteSpace(itemNo))
                        {
                            errors.Add((link.QuotaCode ?? "") + "：无法确认所在条目。");
                            continue;
                        }

                        int order = replacements.Count;
                        FillPreviewItem target = new FillPreviewItem();
                        target.IsNameDriven = true;
                        target.TemplateName = groupLeader.TemplateName;
                        target.TargetRow = groupLeader.TargetRow;
                        target.TargetChapter = groupLeader.TargetChapter;
                        target.TargetName = order == 0 ? groupLeader.TargetName : "";
                        target.TargetFullName = groupLeader.TargetFullName;
                        target.TargetUnit = groupLeader.TargetUnit;
                        target.TargetQuantityText = groupLeader.TargetQuantityText;
                        target.OrderInItem = groupLeader.OrderInItem;
                        target.NeighborSourceQuotaSeq = groupLeader.NeighborSourceQuotaSeq;
                        target.GroupOrder = order;
                        target.ChosenQuotaSeq = link.QuotaSequence;
                        target.QuotaCode = link.QuotaCode;
                        target.SourceName = link.QuotaName;
                        target.IsLibraryQuota = false;
                        target.ChosenItemSeq = itemSeq;
                        target.ChosenItemNo = itemNo;
                        target.ItemNo = itemNo;
                        target.Unit = ResolveTemplateFillQuotaUnit(conn, row, link.QuotaSequence);
                        string qtyBase = String.IsNullOrEmpty(groupLeader.TargetQuantityText) ? groupLeader.QuantityText : groupLeader.TargetQuantityText;
                        if (String.IsNullOrWhiteSpace(groupLeader.TargetUnit) || String.IsNullOrWhiteSpace(target.Unit))
                        {
                            errors.Add((link.QuotaCode ?? "") + "：无法确认 Excel/定额单位，不能安全学习数量关系。");
                            continue;
                        }
                        string standardSuffix;
                        if (TryBuildExcelLinkUnitScaleSuffix(groupLeader.TargetUnit, target.Unit, out standardSuffix))
                        {
                            target.QuantityText = (qtyBase ?? "") + standardSuffix;
                        }
                        else
                        {
                            string pairKey = NormalizeExcelLinkUnit(groupLeader.TargetUnit) + "\n" + NormalizeExcelLinkUnit(target.Unit) + "\n" + (link.QuotaCode ?? "");
                            string conversionFactor;
                            if (!conversionFactors.TryGetValue(pairKey, out conversionFactor))
                            {
                                conversionFactor = PromptTemplateCrossUnitFactor(groupLeader.TargetUnit, target.Unit, link.QuotaCode);
                                if (conversionFactor == null) return;
                                conversionFactors[pairKey] = conversionFactor;
                            }
                            target.FormulaTemplate = conversionFactor == "1" ? "V0" : "V0*" + conversionFactor;
                            string formulaName = StripTrailingQuantityUnit(
                                String.IsNullOrWhiteSpace(groupLeader.TargetFullName) ? groupLeader.TargetName : groupLeader.TargetFullName,
                                groupLeader.TargetUnit);
                            string formulaSignature = NormalizeForSignature(formulaName) + "|";
                            if (formulaSignature.Length > 450) formulaSignature = formulaSignature.Substring(0, 450);
                            target.FormulaOperands = new List<QuantityFormulaOperandInfo>
                            {
                                new QuantityFormulaOperandInfo { Name = formulaName, Unit = groupLeader.TargetUnit, Signature = formulaSignature }
                            };
                            target.QuantityText = (qtyBase ?? "") + "*" + conversionFactor;
                        }
                        target.NeedManualQuota = false;
                        target.Selected = true;
                        target.Status = "";
                        target.AlignNote = order == 0
                            ? ("已绑定 " + (link.QuotaCode ?? "") + (rows.Count > 1 ? "（组 " + rows.Count.ToString() + " 条）" : "（软件选中行，含条目）"))
                            : ("组件框第 " + (order + 1).ToString(CultureInfo.InvariantCulture) + " 条（软件选中行）");
                        replacements.Add(target);
                    }

                    if (errors.Count > 0)
                    {
                        MessageBox.Show(this, "整组重绑未执行，原组件保持不变：\n" + String.Join("\n", errors.ToArray()), "模板铺量");
                        return;
                    }
                    if (replacements.Count == 0) return;
                    if (smartOnly)
                    {
                        replacements = replacements
                            .OrderBy(target => TemplateTargetRank(target.QuotaCode))
                            .ThenBy(target => target.QuotaCode ?? "", StringComparer.OrdinalIgnoreCase)
                            .ToList();
                        for (int i = 0; i < replacements.Count; i++)
                        {
                            replacements[i].GroupOrder = i;
                            replacements[i].TargetName = i == 0 ? groupLeader.TargetName : "";
                            replacements[i].AlignNote = i == 0
                                ? ("已绑定 " + (replacements[i].QuotaCode ?? "") +
                                    (replacements.Count > 1 ? "（组 " + replacements.Count.ToString(CultureInfo.InvariantCulture) + " 条）" : "（软件选中行，含条目）"))
                                : ("组件框第 " + (i + 1).ToString(CultureInfo.InvariantCulture) + " 条（软件选中行）");
                        }
                    }

                    bool bindingChanged = !AreEquivalentNameBindingGroups(oldGroup, replacements);
                    if (!ReplacePreviewTargetGroup(preview, groupLeader.TargetRow, replacements)) return;
                    if (bindingChanged)
                    {
                        FeedbackNameMatches(groupLeader.TemplateName, replacements,
                            System.IO.Path.GetFileName(GetSelectedTargetWorkbookPath() ?? ""), cmbTargetSheet.Text.Trim(), conn, oldGroup);
                    }
                    else
                    {
                        replacements[0].AlignNote = "绑定关系未变化，未重复学习";
                    }
                    RefreshTargetGroupInGrid(groupLeader.TargetRow);
                }
                catch (Exception ex) { MessageBox.Show(this, "绑定失败：" + ex.Message, "模板铺量"); }
            }

            private bool IsSmartPreviewContextCurrent(CurrentSmartEntry entry, out string error)
            {
                error = "";
                SmartPreviewContext context = previewContext;
                if (context == null || entry == null)
                {
                    error = "预览已失效，请重新预览";
                    return false;
                }
                if (!Object.ReferenceEquals(context.ProjectConnection, entry.ProjectConnection) ||
                    !String.Equals(context.ProjectConnectionIdentity, entry.ProjectConnectionIdentity, StringComparison.OrdinalIgnoreCase) ||
                    context.CurrentUnitId != entry.UnitId ||
                    (!String.IsNullOrWhiteSpace(context.CurrentUnitCode) &&
                     !String.Equals(context.CurrentUnitCode, entry.UnitCode, StringComparison.OrdinalIgnoreCase)))
                {
                    error = "项目或当前单元已切换，请重新预览";
                    return false;
                }
                if (context.PreviewVersion != smartPreviewVersion ||
                    !String.Equals(context.ScopeKind, selectedSmartLearningScope == null ? "All" : selectedSmartLearningScope.Kind ?? "All", StringComparison.OrdinalIgnoreCase) ||
                    !String.Equals(context.ScopeEntryCode, selectedSmartLearningScope == null ? "" : selectedSmartLearningScope.EntryCode ?? "", StringComparison.OrdinalIgnoreCase) ||
                    !String.Equals(context.Worksheet, cmbTargetSheet.Text.Trim(), StringComparison.OrdinalIgnoreCase) ||
                    !String.Equals(context.TargetColumn, txtColumn.Text.Trim().ToUpperInvariant(), StringComparison.OrdinalIgnoreCase))
                {
                    error = "推荐范围或 Excel 预览输入已变化，请重新预览";
                    return false;
                }
                string activeWorkbook;
                string workbookError;
                updatingSmartPreviewInputs = true;
                bool workbookReady;
                try { workbookReady = TryResolveSmartActiveWorkbook(out activeWorkbook, out workbookError); }
                finally { updatingSmartPreviewInputs = false; }
                if (!workbookReady || !String.Equals(Path.GetFullPath(activeWorkbook), context.WorkbookPath, StringComparison.OrdinalIgnoreCase))
                {
                    error = workbookReady ? "当前活动工作簿已变化，请重新预览" : workbookError;
                    return false;
                }
                SmartMethodRoute route = ResolveSmartMethodRoute(SmartResolveProjectMethod(entry.ProjectConnection));
                if (!String.Equals(context.SoftwarePartition, ResolveLearningSoftwarePartition(), StringComparison.OrdinalIgnoreCase) ||
                    !String.Equals(context.MethodNo, NormalizeLearningMethodNo(route.MethodNo), StringComparison.OrdinalIgnoreCase))
                {
                    error = "当前软件分区或编制办法已变化，请重新预览";
                    return false;
                }
                return true;
            }

            private bool StampSelectedSmartEntries(List<FillPreviewItem> selectedItems, CurrentSmartEntry entry, out string error)
            {
                error = "";
                bool currentIsEquipment = (entry.EntryName ?? "").IndexOf("设备购置费", StringComparison.OrdinalIgnoreCase) >= 0;
                long sfSequence = entry.EntrySequence;
                string sfCode = entry.EntryCode;
                string sfName = entry.EntryName;
                bool needsSfRedirect = selectedItems.Any(item => String.Equals((item.QuotaCode ?? "").Trim(), "SF", StringComparison.OrdinalIgnoreCase)) &&
                    !currentIsEquipment;
                if (needsSfRedirect && !TryResolveSiblingEquipmentEntry(entry.ProjectConnection, entry.EntryCode,
                    out sfSequence, out sfCode, out sfName, out error)) return false;

                foreach (IGrouping<int, FillPreviewItem> group in selectedItems.GroupBy(item => item.TargetRow))
                {
                    bool hasSf = group.Any(item => String.Equals((item.QuotaCode ?? "").Trim(), "SF", StringComparison.OrdinalIgnoreCase));
                    bool hasNonSf = group.Any(item => !String.Equals((item.QuotaCode ?? "").Trim(), "SF", StringComparison.OrdinalIgnoreCase));
                    if (currentIsEquipment && hasNonSf)
                    {
                        error = "设备购置费条目只接受 SF，所选组件整组未写入";
                        return false;
                    }
                    if (hasSf && !currentIsEquipment && sfSequence <= 0)
                    {
                        error = "未找到唯一同级设备购置费条目，所选组件整组未写入";
                        return false;
                    }
                    foreach (FillPreviewItem item in group)
                    {
                        bool sf = String.Equals((item.QuotaCode ?? "").Trim(), "SF", StringComparison.OrdinalIgnoreCase);
                        item.ChosenItemSeq = sf ? sfSequence : entry.EntrySequence;
                        item.ChosenItemNo = sf ? sfCode : entry.EntryCode;
                        item.ChosenItemName = sf ? sfName : entry.EntryName;
                        item.SfRedirect = sf && !currentIsEquipment;
                        item.EntrySource = item.SfRedirect ? "sf-sibling-redirect" : "user-selected";
                        item.Selected = true;
                    }
                }
                return true;
            }

            private void OnApply()
            {
                try
                {
                    FlushGridSelectionsToPreview();
                    if (smartOnly)
                    {
                        RefreshCurrentSmartEntry(true);
                        string contextError = "当前树节点不是可写入的条目，请选中具体条目后重试。";
                        if (!currentEntryWritable || currentSmartEntry == null ||
                            !IsSmartPreviewContextCurrent(currentSmartEntry, out contextError))
                        {
                            MessageBox.Show(this, String.IsNullOrWhiteSpace(contextError)
                                ? "当前树节点不是可写入的条目，请选中具体条目后重试。"
                                : contextError, "推荐定额");
                            return;
                        }
                        HashSet<int> selectedRows = GetSelectedSmartTargetRows();
                        HashSet<int> checkedRows = GetCheckedSmartTargetRows();
                        selectedRows.IntersectWith(checkedRows);
                        if (selectedRows.Count == 0)
                        {
                            MessageBox.Show(this, "没有同时被选中且勾选的定额。请先在表格中选中要写的行（可 Ctrl/Shift 多选），再勾选左侧复选框。", "推荐定额");
                            return;
                        }
                        List<FillPreviewItem> selectedItems = preview.Where(item => item != null && selectedRows.Contains(item.TargetRow))
                            .OrderBy(item => item.TargetRow).ThenBy(item => item.GroupOrder).ToList();
                        string stampError;
                        if (!StampSelectedSmartEntries(selectedItems, currentSmartEntry, out stampError))
                        {
                            MessageBox.Show(this, stampError, "推荐定额");
                            return;
                        }
                        long approvedEntrySequence = currentSmartEntry.EntrySequence;
                        string approvedEntryCode = currentSmartEntry.EntryCode;
                        if (MessageBox.Show(this, "确认把选中且勾选的 " + selectedItems.Count.ToString(CultureInfo.InvariantCulture) +
                            " 条定额写入当前条目【" + currentSmartEntry.EntryCode + " " + currentSmartEntry.EntryName + "】？\n写入动作本身会保存；“计算”只刷新价格和汇总。",
                            "推荐定额", MessageBoxButtons.OKCancel) != DialogResult.OK) return;
                        // 用户确认框的嵌套消息循环期间仍可能切换项目/条目；真正写入前再读一次并与确认文案中的目标对比。
                        RefreshCurrentSmartEntry(true);
                        contextError = "确认后项目或条目已变化，请重新选择并写入。";
                        if (!currentEntryWritable || currentSmartEntry == null ||
                            !IsSmartPreviewContextCurrent(currentSmartEntry, out contextError) ||
                            currentSmartEntry.EntrySequence != approvedEntrySequence ||
                            !String.Equals(currentSmartEntry.EntryCode, approvedEntryCode, StringComparison.OrdinalIgnoreCase))
                        {
                            MessageBox.Show(this, String.IsNullOrWhiteSpace(contextError)
                                ? "确认后项目或条目已变化，请重新选择并写入。"
                                : contextError, "推荐定额");
                            return;
                        }
                        if (!StampSelectedSmartEntries(selectedItems, currentSmartEntry, out stampError))
                        {
                            MessageBox.Show(this, stampError, "推荐定额");
                            return;
                        }
                        SetBusy(true, "写入中...");
                        string sourceWorkbook = Path.GetFileName(previewContext.WorkbookPath);
                        bool smartSucceeded;
                        string smartResult = ApplyFillToSelectedEntry(mainForm, currentSmartEntry.UnitId, currentSmartEntry.UnitCode,
                            currentSmartEntry.EntrySequence, currentSmartEntry.EntryCode, currentSmartEntry.EntryName,
                            selectedItems, sourceWorkbook, previewContext.Worksheet, out smartSucceeded);
                        MessageBox.Show(this, smartResult, "推荐定额");
                        if (smartSucceeded) InvalidateSmartPreview();
                        return;
                    }
                    int selectedCount = preview.Count(it => it.Selected);
                    string targetUnit = cmbTargetUnit.Text.Trim();
                    if (MessageBox.Show(this, "确认把勾选的 " + selectedCount.ToString() + " 条定额（含树筛选后未显示条目中已勾选的行）复制到目标单元【" + targetUnit + "】的对应条目？",
                        "模板铺量", MessageBoxButtons.OKCancel) != DialogResult.OK) return;

                    SetBusy(true, "写入中...");
                    string result = ApplyFill(mainForm, targetUnit, preview,
                        System.IO.Path.GetFileName(GetSelectedTargetWorkbookPath() ?? ""), cmbTargetSheet.Text.Trim());
                    MessageBox.Show(this, result, "模板铺量");
                }
                catch (Exception ex) { MessageBox.Show(this, "写入失败：" + ex.Message, smartOnly ? "推荐定额" : "模板铺量"); }
                finally { SetBusy(false, ""); }
            }

            private void SetBusy(bool busy, string action)
            {
                this.busy = busy;
                UseWaitCursor = busy;
                Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
                string baseTitle = smartOnly ? "推荐定额" : "模板铺量";
                Text = busy && !String.IsNullOrEmpty(action) ? baseTitle + " - " + action : baseTitle;
                RefreshApplyEnabled();
                Refresh();
            }
        }

        private static readonly Dictionary<Form, TemplateFillPanel> TemplateFillPanels = new Dictionary<Form, TemplateFillPanel>();
        private static void ShowTemplateFillPanel(Form mainForm)
        {
            TemplateFillPanel panel;
            if (!TemplateFillPanels.TryGetValue(mainForm, out panel) || panel == null || panel.IsDisposed)
            {
                panel = new TemplateFillPanel(mainForm);
                TemplateFillPanels[mainForm] = panel;
            }
            panel.Show(mainForm); panel.Activate();
        }

        private static readonly Dictionary<Form, TemplateFillPanel> SmartFillPanels = new Dictionary<Form, TemplateFillPanel>();

        // 推荐定额入口:学习库智能铺量独立窗口(上部只有目标一行),与模板铺量窗口互不干扰。
        private static void ShowSmartFillPanel(Form mainForm)
        {
            TemplateFillPanel panel;
            if (!SmartFillPanels.TryGetValue(mainForm, out panel) || panel == null || panel.IsDisposed)
            {
                panel = new TemplateFillPanel(mainForm, true);
                SmartFillPanels[mainForm] = panel;
            }
            panel.Show(mainForm); panel.Activate();
        }
    }
}
