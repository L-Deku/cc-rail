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
            private readonly Button btnPreview = new Button();
            private readonly Button btnApply = new Button();
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
            private int targetWorkbookReloadCount;

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
                if (!smartOnly)
                {
                    ReloadTemplateList();
                    ReloadSourceSheets();
                }
                ReloadTargetWorkbooks();
                string cur = GetCurrentUnitNo(mainForm);
                if (!String.IsNullOrEmpty(cur)) txtUnit.Text = cur;
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
                AddLabel("目标Excel", 12, targetTop + 3, 60);
                cmbTargetWorkbook.SetBounds(75, targetTop, 190, 23);
                cmbTargetWorkbook.DropDownWidth = 420;
                cmbTargetWorkbook.DropDownStyle = ComboBoxStyle.DropDownList;
                cmbTargetWorkbook.DropDown += delegate { ReloadTargetWorkbooks(); };
                cmbTargetWorkbook.SelectedIndexChanged += delegate { if (!reloadingTargetWorkbooks) ReloadTargetSheets(); };
                AddLabel("目标sheet", 275, targetTop + 3, 60);
                cmbTargetSheet.SetBounds(340, targetTop, 105, 23); cmbTargetSheet.Text = "";
                cmbTargetSheet.DropDownStyle = ComboBoxStyle.DropDown; // 可选可填
                cmbTargetSheet.DropDown += delegate { ReloadTargetSheets(); };
                AddLabel("目标列", 455, targetTop + 3, 50);
                txtColumn.SetBounds(505, targetTop, 40, 23); txtColumn.Text = "";
                AddLabel("目标单元", 555, targetTop + 3, 60);
                cmbTargetUnit.SetBounds(620, targetTop, 80, 23); cmbTargetUnit.Text = "_ZGS_02";
                cmbTargetUnit.DropDownStyle = ComboBoxStyle.DropDown; // 可选可填
                cmbTargetUnit.DropDown += delegate { ReloadTargetUnits(); };
                btnPreview.SetBounds(710, targetTop - 1, 60, 25); btnPreview.Text = "预览";
                btnPreview.Click += delegate { OnPreview(); };
                btnApply.SetBounds(780, targetTop - 1, 108, 25); btnApply.Text = "写入目标单元";
                btnApply.Click += delegate { OnApply(); };

                Label reminder = new Label
                {
                    Text = "写入＝复制定额到“目标单元”的对应条目（条目序号全局共享）。写入后请在软件点一次“计算”刷新单价与汇总。",
                    ForeColor = Color.Firebrick, AutoSize = false
                };
                reminder.SetBounds(12, targetTop + 29, 876, 18);
                reminder.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

                // —— 左侧条目树 + 右侧预览表：SplitContainer 分栏，可拖动调整宽度 ——
                split.SetBounds(12, targetTop + 53, 876, ClientSize.Height - (targetTop + 53) - 12);
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
                grid.RowHeadersVisible = false;
                grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "选", Name = "sel", FillWeight = 6 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "定额编号", Name = "code", ReadOnly = false, FillWeight = 16 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "源行定额", Name = "sname", ReadOnly = true, FillWeight = 20 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "单位", Name = "unit", ReadOnly = true, FillWeight = 8 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "目标行工程量名", Name = "tname", ReadOnly = true, FillWeight = 20 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "数量", Name = "qty", ReadOnly = true, FillWeight = 10 });
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
                        !String.Equals(grid.Columns[e.ColumnIndex].Name, "code", StringComparison.Ordinal)) return;
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
                        ConfirmExactNameFromCheck(grid.Rows[e.RowIndex]);
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

                    int above = 0, total = 0;
                    for (int i = start; i <= end; i++)
                    {
                        if (i < e.RowIndex) above += grid.Rows[i].Height;
                        total += grid.Rows[i].Height;
                    }
                    Rectangle union = new Rectangle(e.CellBounds.Left, e.CellBounds.Top - above, e.CellBounds.Width, total);

                    if (e.RowIndex > start) e.AdvancedBorderStyle.Top = DataGridViewAdvancedCellBorderStyle.None;
                    e.PaintBackground(e.ClipBounds, true);
                    string text = Convert.ToString(grid.Rows[start].Cells[e.ColumnIndex].Value);
                    if (!String.IsNullOrEmpty(text))
                    {
                        bool selected = (e.State & DataGridViewElementStates.Selected) != 0;
                        Color foreColor = selected ? e.CellStyle.SelectionForeColor : e.CellStyle.ForeColor;
                        TextRenderer.DrawText(e.Graphics, text, e.CellStyle.Font, union, foreColor,
                            TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.WordBreak | TextFormatFlags.EndEllipsis);
                    }
                    e.Handled = true;
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

                if (!smartOnly)
                {
                    Controls.Add(txtUnit); Controls.Add(cmbSourceSheet); Controls.Add(txtName); Controls.Add(btnBuild);
                    Controls.Add(cmbTemplate); Controls.Add(btnDeleteTemplate); Controls.Add(cmbMode);
                }
                Controls.Add(cmbTargetWorkbook); Controls.Add(cmbTargetSheet); Controls.Add(txtColumn);
                Controls.Add(cmbTargetUnit);
                Controls.Add(btnPreview); Controls.Add(btnApply); Controls.Add(reminder); Controls.Add(split);
            }

            private void AddLabel(string text, int x, int y, int w)
            {
                Label l = new Label { Text = text, AutoSize = false };
                l.SetBounds(x, y, w, 18); Controls.Add(l);
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
                    using (SqlConnection conn = AgentCreateWorkConnection(mainForm))
                    {
                        foreach (string s in ListBoundSheetNames(conn)) cmbSourceSheet.Items.Add(s);
                    }
                    if (!String.IsNullOrEmpty(keep)) cmbSourceSheet.Text = keep;
                    else if (cmbSourceSheet.Items.Count > 0) cmbSourceSheet.SelectedIndex = 0;
                }
                catch { /* 取不到绑定时留空，用户可手填 */ }
            }

            private string GetSelectedTargetWorkbookPath()
            {
                OpenSpreadsheetWorkbookInfo selected = cmbTargetWorkbook.SelectedItem as OpenSpreadsheetWorkbookInfo;
                return selected == null ? "" : selected.FullName ?? "";
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
                catch { /* 取不到时留空，预览时给出明确提示 */ }
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
                catch { /* 取不到时留空，预览时给出明确提示 */ }
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
                    using (SqlConnection conn = AgentCreateWorkConnection(mainForm))
                    {
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
                    // 用克隆的独立连接（与 ApplyFill/智能助手一致），不要直接用主程序共享连接，
                    // 否则可能拿到未初始化连接串、或被 using 误释放主程序连接。
                    int count;
                    List<string> warnings;
                    using (SqlConnection conn = AgentCreateWorkConnection(mainForm))
                    {
                        FillTemplate t = chkNameMode.Checked
                            ? BuildNameFillTemplateFromBindings(mainForm, conn, txtName.Text.Trim(), txtUnit.Text.Trim(), cmbSourceSheet.Text.Trim())
                            : BuildFillTemplateFromBindings(mainForm, conn, txtName.Text.Trim(), txtUnit.Text.Trim(), cmbSourceSheet.Text.Trim());
                        count = t.Rows.Count;
                        warnings = t.BuildWarnings;
                        SaveFillTemplate(t);
                    }
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
                SetBusy(true, "预览中...");
                try
                {
                    // 模式三·推荐定额:不需要模板,直接走学习库漏斗。
                    if (cmbMode.SelectedIndex == 2)
                    {
                        string smartWorkbook = GetSelectedTargetWorkbookPath();
                        if (String.IsNullOrWhiteSpace(smartWorkbook) || !File.Exists(smartWorkbook))
                        {
                            MessageBox.Show(this, "没有可用的目标 Excel，请先打开并保存工作簿。", "推荐定额");
                            return;
                        }
                        string smartWarning = null;
                        preview = BuildPreview_SmartFill(mainForm, smartWorkbook, cmbTargetSheet.Text.Trim(), txtColumn.Text.Trim(), out smartWarning);
                        if (!String.IsNullOrEmpty(smartWarning)) MessageBox.Show(this, smartWarning, "推荐定额");
                        RebuildItemTree();
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
                    IEnumerable<FillPreviewItem> ordered = preview
                        .Where(item => String.IsNullOrEmpty(scope) || IsItemNoUnderChapter(item.ItemNo ?? "", scope));
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
                }
                finally
                {
                    updatingNameQuotaCell = previousUpdating;
                }
            }

            private void SetGridRow(DataGridViewRow row, FillPreviewItem item, FillPreviewItem leader)
            {
                int codeIndex = grid.Columns["code"].Index;
                if (row.Cells[codeIndex] is DataGridViewComboBoxCell)
                {
                    row.Cells[codeIndex] = new DataGridViewTextBoxCell();
                }

                string statusText = String.IsNullOrEmpty(item.Status) ? (item.AlignNote ?? "") : item.Status;
                row.SetValues(item.Selected, item.QuotaCode, item.SourceName, item.Unit ?? "",
                    item.TargetName, item.QuantityText, statusText);
                row.Tag = item;
                row.Cells["sel"].ReadOnly = false;
                row.Cells["sel"].ToolTipText = "";
                row.Cells["code"].ToolTipText = "";
                row.DefaultCellStyle.BackColor = Color.Empty;

                bool hasCandidates = item.GroupOrder == 0 && item.NameQuotaCandidates != null &&
                    item.NameQuotaCandidates.Count > 1;
                bool requiresChoice = leader != null && leader.NeedExactNameConfirmation &&
                    leader.NameQuotaCandidates != null && leader.NameQuotaCandidates.Count > 1;
                row.Cells["code"].ReadOnly = !hasCandidates;
                if (hasCandidates)
                {
                    row.Cells["code"].ToolTipText = "点击选择该工程量名称绑定的定额或组件组";
                }
                if (requiresChoice)
                {
                    if (item.GroupOrder > 0)
                    {
                        row.Cells["sel"].ReadOnly = true;
                        row.Cells["sel"].ToolTipText = "请在组首勾选接受当前候选，或在定额编号列切换绑定组";
                    }
                    else
                    {
                        row.Cells["sel"].ToolTipText = "勾选接受当前候选；点击定额编号可切换绑定组";
                    }
                }
                if (item.NeedExactNameConfirmation || !String.IsNullOrEmpty(item.Status))
                    row.DefaultCellStyle.BackColor = Color.MistyRose;
                else if (item.NeedManualQuota)
                    row.DefaultCellStyle.BackColor = Color.FromArgb(255, 246, 196);
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
                    .Where(item => item != null && item.IsNameDriven && item.TargetRow == targetRow &&
                        (String.IsNullOrEmpty(currentTreeScope) || IsItemNoUnderChapter(item.ItemNo ?? "", currentTreeScope)))
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
                if (row.Cells["code"] is DataGridViewComboBoxCell) return true;

                DataGridViewComboBoxCell combo = new DataGridViewComboBoxCell();
                combo.DisplayStyle = DataGridViewComboBoxDisplayStyle.DropDownButton;
                foreach (NameQuotaCandidateGroup option in item.NameQuotaCandidates)
                {
                    combo.Items.Add(option.Label);
                }
                NameQuotaCandidateGroup current = item.NameQuotaCandidates.FirstOrDefault(option =>
                    String.Equals(option.Key, item.SelectedNameQuotaCandidateKey, StringComparison.Ordinal));
                combo.Value = (current ?? item.NameQuotaCandidates[0]).Label;

                updatingNameQuotaCell = true;
                try { row.Cells[grid.Columns["code"].Index] = combo; }
                finally { updatingNameQuotaCell = false; }
                return true;
            }

            private void OnNameQuotaSelectionCommitted(object sender, EventArgs e)
            {
                ComboBox combo = sender as ComboBox;
                if (combo == null || grid.CurrentRow == null) return;
                ApplyNameQuotaOption(grid.CurrentRow, Convert.ToString(combo.SelectedItem));
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
                if (item == null || !value || !item.NeedExactNameConfirmation) return;

                updatingNameQuotaCell = true;
                try
                {
                    int targetRow = item.TargetRow;
                    if (ConfirmCurrentExactNameGroup(preview, targetRow))
                        RefreshTargetGroupInGrid(targetRow);
                }
                finally { updatingNameQuotaCell = false; }
            }

            // —— 条目树 ——
            // 用章节表的 条目编号+名称 构建层级：只显示编号为两位数字（01、02…）这级及以下，
            // “第一部分”这类更高层级不进树。节点 Tag 为条目编号，根节点 Tag 为空串表示全部。
            private void RebuildItemTree()
            {
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
                    if (it != null)
                    {
                        it.Selected = Convert.ToBoolean(row.Cells["sel"].Value ?? false);
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
                        target.QuantityText = BuildNameDrivenQtyText(qtyBase, groupLeader.TargetUnit, target.Unit);
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

                    if (!ReplacePreviewTargetGroup(preview, groupLeader.TargetRow, replacements)) return;
                    RefreshTargetGroupInGrid(groupLeader.TargetRow);
                }
                catch (Exception ex) { MessageBox.Show(this, "绑定失败：" + ex.Message, "模板铺量"); }
            }

            private void OnApply()
            {
                try
                {
                    FlushGridSelectionsToPreview();
                    int selectedCount = preview.Count(it => it.Selected);
                    string targetUnit = cmbTargetUnit.Text.Trim();
                    if (MessageBox.Show(this, "确认把勾选的 " + selectedCount.ToString() + " 条定额（含树筛选后未显示条目中已勾选的行）复制到目标单元【" + targetUnit + "】的对应条目？",
                        "模板铺量", MessageBoxButtons.OKCancel) != DialogResult.OK) return;

                    SetBusy(true, "写入中...");
                    string result = ApplyFill(mainForm, targetUnit, preview);
                    MessageBox.Show(this, result, "模板铺量");
                }
                catch (Exception ex) { MessageBox.Show(this, "写入失败：" + ex.Message, "模板铺量"); }
                finally { SetBusy(false, ""); }
            }

            private void SetBusy(bool busy, string action)
            {
                UseWaitCursor = busy;
                Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
                btnPreview.Enabled = !busy;
                btnApply.Enabled = !busy;
                Text = busy && !String.IsNullOrEmpty(action) ? "模板铺量 - " + action : "模板铺量";
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
