using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
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
            private readonly ComboBox cmbTargetSheet = new ComboBox();
            private readonly TextBox txtColumn = new TextBox();
            private readonly ComboBox cmbTargetUnit = new ComboBox();
            private readonly Button btnPreview = new Button();
            private readonly Button btnApply = new Button();
            private readonly CheckBox chkNameMode = new CheckBox();
            private readonly SplitContainer split = new SplitContainer();
            private readonly TreeView itemTree = new TreeView();
            private readonly DataGridView grid = new DataGridView();
            private Dictionary<string, string> chapterNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            private string currentTreeScope = "";
            private bool updatingTreeChecks;
            private bool rebuildingTree;

            public TemplateFillPanel(Form owner)
            {
                mainForm = owner;
                Text = "模板铺量";
                StartPosition = FormStartPosition.CenterParent;
                ClientSize = new Size(900, 580);
                BuildLayout();
                ReloadTemplateList();
                ReloadSourceSheets();
                string cur = GetCurrentUnitNo(mainForm);
                if (!String.IsNullOrEmpty(cur)) txtUnit.Text = cur;
            }

            private void BuildLayout()
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
                cmbMode.Items.AddRange(new object[] { "一·列锚点", "二·名字驱动" });
                cmbMode.SelectedIndex = 0;
                AddLabel("目标sheet", 12, 82, 60);
                cmbTargetSheet.SetBounds(75, 79, 120, 23); cmbTargetSheet.Text = "";
                cmbTargetSheet.DropDownStyle = ComboBoxStyle.DropDown; // 可选可填
                cmbTargetSheet.DropDown += delegate { ReloadTargetSheets(); };
                AddLabel("目标列", 205, 82, 50);
                txtColumn.SetBounds(255, 79, 50, 23); txtColumn.Text = "";
                AddLabel("目标单元", 315, 82, 60);
                cmbTargetUnit.SetBounds(380, 79, 90, 23); cmbTargetUnit.Text = "_ZGS_02";
                cmbTargetUnit.DropDownStyle = ComboBoxStyle.DropDown; // 可选可填
                cmbTargetUnit.DropDown += delegate { ReloadTargetUnits(); };
                btnPreview.SetBounds(480, 78, 70, 25); btnPreview.Text = "预览";
                btnPreview.Click += delegate { OnPreview(); };
                btnApply.SetBounds(560, 78, 150, 25); btnApply.Text = "写入目标单元";
                btnApply.Click += delegate { OnApply(); };

                Label reminder = new Label
                {
                    Text = "写入＝复制定额到“目标单元”的对应条目（条目序号全局共享）。写入后请在软件点一次“计算”刷新单价与汇总。",
                    ForeColor = Color.Firebrick, AutoSize = false
                };
                reminder.SetBounds(12, 108, 876, 18);
                reminder.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

                // —— 左侧条目树 + 右侧预览表：SplitContainer 分栏，可拖动调整宽度 ——
                split.SetBounds(12, 132, 876, 436);
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
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "定额编号", Name = "code", ReadOnly = true, FillWeight = 16 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "源行定额", Name = "sname", ReadOnly = true, FillWeight = 20 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "单位", Name = "unit", ReadOnly = true, FillWeight = 8 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "目标行工程量名", Name = "tname", ReadOnly = true, FillWeight = 20 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "数量", Name = "qty", ReadOnly = true, FillWeight = 10 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "状态", Name = "st", ReadOnly = true, FillWeight = 14 });
                grid.CurrentCellDirtyStateChanged += delegate
                {
                    if (grid.IsCurrentCellDirty)
                    {
                        grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
                    }
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

                Controls.Add(txtUnit); Controls.Add(cmbSourceSheet); Controls.Add(txtName); Controls.Add(btnBuild);
                Controls.Add(cmbTemplate); Controls.Add(btnDeleteTemplate); Controls.Add(cmbMode); Controls.Add(cmbTargetSheet); Controls.Add(txtColumn);
                Controls.Add(cmbTargetUnit);
                Controls.Add(btnPreview); Controls.Add(btnApply); Controls.Add(reminder); Controls.Add(split);
            }

            private void AddLabel(string text, int x, int y, int w)
            {
                Label l = new Label { Text = text, AutoSize = false };
                l.SetBounds(x, y, w, 18); Controls.Add(l);
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

            // 目标sheet 下拉展开时刷新：优先读当前打开的 Excel/WPS 工作簿的全部工作表名；
            // 读不到（Excel 未开）时回退绑定库里出现过的 sheet 名。
            private void ReloadTargetSheets()
            {
                try
                {
                    string keep = cmbTargetSheet.Text;
                    List<string> sheetNames; string activeSheetName; string error;
                    if (TryListActiveWorkbookSheets(out sheetNames, out activeSheetName, out error) && sheetNames.Count > 0)
                    {
                        cmbTargetSheet.Items.Clear();
                        foreach (string s in sheetNames) cmbTargetSheet.Items.Add(s);
                    }
                    else
                    {
                        cmbTargetSheet.Items.Clear();
                        using (SqlConnection conn = AgentCreateWorkConnection(mainForm))
                        {
                            foreach (string s in ListBoundSheetNames(conn)) cmbTargetSheet.Items.Add(s);
                        }
                    }
                    cmbTargetSheet.Text = keep;
                }
                catch { /* 取不到时留空，用户可手填 */ }
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
                    if (cmbTemplate.SelectedItem == null) { MessageBox.Show(this, "请先选择模板。", "模板铺量"); return; }
                    FillTemplate t = LoadFillTemplate(Convert.ToString(cmbTemplate.SelectedItem));
                    if (t == null) { MessageBox.Show(this, "模板加载失败。", "模板铺量"); return; }
                    string ndWarning = null;
                    preview = cmbMode.SelectedIndex == 0
                        ? BuildPreview_ColumnAnchor(t, cmbTargetSheet.Text.Trim(), txtColumn.Text.Trim())
                        : BuildPreview_NameDriven(mainForm, t, cmbTargetSheet.Text.Trim(), txtColumn.Text.Trim(), out ndWarning);
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
                grid.Rows.Clear();
                string scope = currentTreeScope ?? "";
                bool nameDriven = preview.Any(p => p.IsNameDriven);
                IEnumerable<FillPreviewItem> ordered = preview
                    .Where(item => String.IsNullOrEmpty(scope) || IsItemNoUnderChapter(item.ItemNo ?? "", scope));
                ordered = nameDriven
                    ? ordered.OrderBy(item => item.TargetRow).ThenBy(item => item.GroupOrder)
                    : ordered.OrderBy(item => item.ItemNo ?? "", StringComparer.OrdinalIgnoreCase).ThenBy(item => item.OrderInItem);
                foreach (FillPreviewItem it in ordered)
                {
                    string statusText = String.IsNullOrEmpty(it.Status) ? (it.AlignNote ?? "") : it.Status;
                    int idx = grid.Rows.Add(it.Selected, it.QuotaCode,
                        it.SourceName, it.Unit ?? "", it.TargetName, it.QuantityText, statusText);
                    grid.Rows[idx].Tag = it;
                    if (!String.IsNullOrEmpty(it.Status))
                        grid.Rows[idx].DefaultCellStyle.BackColor = Color.MistyRose;
                    else if (it.NeedManualQuota)
                        grid.Rows[idx].DefaultCellStyle.BackColor = Color.FromArgb(255, 246, 196);
                }
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
                    DataGridView de = GetField<DataGridView>(mainForm, "dataGridViewDE");
                    DataGridViewRow row = GetCurrentQuotaRow(de);
                    if (row == null) { MessageBox.Show(this, "请先在软件定额输入表中选中一条定额行。", "模板铺量"); return; }
                    SqlConnection conn = GetProjectConnection(mainForm);
                    if (conn == null) { MessageBox.Show(this, "没有找到当前项目数据库连接。", "模板铺量"); return; }
                    ExcelQuotaLink link; string err;
                    if (!TryCreateQuotaLink(mainForm, conn, row, out link, out err)) { MessageBox.Show(this, err, "模板铺量"); return; }
                    it.ChosenQuotaSeq = link.QuotaSequence;
                    it.QuotaCode = link.QuotaCode;
                    it.SourceName = link.QuotaName;
                    it.IsLibraryQuota = false;
                    long itemSeq;
                    if (Int64.TryParse((link.ChapterSeq ?? "").Trim(), out itemSeq) && itemSeq > 0)
                    {
                        it.ChosenItemSeq = itemSeq;
                        it.ChosenItemNo = ResolveChapterItemNo(conn, link.ChapterSeq, null);
                        it.ItemNo = it.ChosenItemNo;
                    }
                    it.Unit = GetRowValue(row, "单位", "定额单位");
                    it.NeedManualQuota = false;
                    it.Selected = true;
                    it.Status = "";
                    it.AlignNote = "已绑定 " + (link.QuotaCode ?? "") + "（软件选中行，含条目）";
                    FillGrid();
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
    }
}
