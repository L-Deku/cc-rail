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
            private readonly TextBox txtSheet = new TextBox();
            private readonly TextBox txtColumn = new TextBox();
            private readonly TextBox txtTargetUnit = new TextBox();
            private readonly Button btnPreview = new Button();
            private readonly Button btnApply = new Button();
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

                // —— 套用配置 ——
                AddLabel("模板", 12, 50, 36);
                cmbTemplate.SetBounds(50, 47, 185, 23); cmbTemplate.DropDownStyle = ComboBoxStyle.DropDownList;
                btnDeleteTemplate.SetBounds(240, 46, 70, 25); btnDeleteTemplate.Text = "删除模板";
                btnDeleteTemplate.Click += delegate { OnDeleteTemplate(); };
                AddLabel("取数模式", 320, 50, 60);
                cmbMode.SetBounds(385, 47, 150, 23); cmbMode.DropDownStyle = ComboBoxStyle.DropDownList;
                cmbMode.Items.AddRange(new object[] { "一·列锚点", "二·固定绑定列" });
                cmbMode.SelectedIndex = 0;
                AddLabel("目标sheet", 12, 82, 60);
                txtSheet.SetBounds(75, 79, 120, 23); txtSheet.Text = "";
                AddLabel("目标列", 205, 82, 50);
                txtColumn.SetBounds(255, 79, 50, 23); txtColumn.Text = "";
                AddLabel("目标单元", 315, 82, 60);
                txtTargetUnit.SetBounds(380, 79, 90, 23); txtTargetUnit.Text = "_ZGS_02";
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
                grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "选", Name = "sel", FillWeight = 6 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "条目", Name = "item", ReadOnly = true, FillWeight = 14, Visible = false });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "定额编号", Name = "code", ReadOnly = true, FillWeight = 16 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "源行定额", Name = "sname", ReadOnly = true, FillWeight = 20 });
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

                split.Panel1.Controls.Add(itemTree);
                split.Panel2.Controls.Add(grid);

                Controls.Add(txtUnit); Controls.Add(cmbSourceSheet); Controls.Add(txtName); Controls.Add(btnBuild);
                Controls.Add(cmbTemplate); Controls.Add(btnDeleteTemplate); Controls.Add(cmbMode); Controls.Add(txtSheet); Controls.Add(txtColumn);
                Controls.Add(txtTargetUnit);
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
                    using (SqlConnection conn = AgentCreateWorkConnection(mainForm))
                    {
                        FillTemplate t = BuildFillTemplateFromBindings(mainForm, conn, txtName.Text.Trim(),
                            txtUnit.Text.Trim(), cmbSourceSheet.Text.Trim());
                        count = t.Rows.Count;
                        SaveFillTemplate(t);
                    }
                    ReloadTemplateList();
                    MessageBox.Show(this, count > 0
                        ? ("模板已生成并保存：" + count + " 条定额。")
                        : ("模板已生成，但收到 0 条定额。\n请确认该单元的定额已用“绑定Excel工程量”绑到 sheet「" + cmbSourceSheet.Text.Trim() + "」。"),
                        "模板铺量");
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
                    preview = cmbMode.SelectedIndex == 0
                        ? BuildPreview_ColumnAnchor(t, txtSheet.Text.Trim(), txtColumn.Text.Trim())
                        : BuildPreview_FixedColumn(t);
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
                foreach (FillPreviewItem it in preview
                    .Where(item => String.IsNullOrEmpty(scope) || IsItemNoUnderChapter(item.ItemNo ?? "", scope))
                    .OrderBy(item => item.ItemNo ?? "", StringComparer.OrdinalIgnoreCase)
                    .ThenBy(item => item.OrderInItem))
                {
                    int idx = grid.Rows.Add(it.Selected, it.ItemNo, it.QuotaCode,
                        it.SourceName, it.TargetName, it.QuantityText, it.Status);
                    grid.Rows[idx].Tag = it;
                    if (!String.IsNullOrEmpty(it.Status))
                        grid.Rows[idx].DefaultCellStyle.BackColor = Color.MistyRose;
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
                    LoadChapterNames();

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
                        foreach (string code in BuildChapterChain(itemNo))
                        {
                            TreeNode node;
                            if (!nodesByCode.TryGetValue(code, out node))
                            {
                                node = new TreeNode(GetChapterDisplayName(code));
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

            private void LoadChapterNames()
            {
                chapterNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                try
                {
                    using (SqlConnection conn = AgentCreateWorkConnection(mainForm))
                    {
                        EnsureOpen(conn);
                        using (SqlCommand cmd = conn.CreateCommand())
                        {
                            cmd.CommandText = "select 条目编号, 工程或费用项目名称 from 章节表";
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    if (reader.IsDBNull(0))
                                    {
                                        continue;
                                    }

                                    string code = Convert.ToString(reader.GetValue(0)).Trim();
                                    if (code.Length == 0 || chapterNames.ContainsKey(code))
                                    {
                                        continue;
                                    }

                                    chapterNames[code] = reader.IsDBNull(1) ? "" : Convert.ToString(reader.GetValue(1)).Trim();
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log("Template fill load chapter names failed: " + ex.Message);
                }
            }

            // 条目编号的祖先链：章节表里存在、是 itemNo 前缀、且编号以两位数字开头的编号，按长度升序。
            private List<string> BuildChapterChain(string itemNo)
            {
                List<string> chain = chapterNames.Keys
                    .Where(code => IsChapterTreeCode(code) && IsItemNoUnderChapter(itemNo, code))
                    .OrderBy(code => code.Length)
                    .ThenBy(code => code, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (!chain.Any(code => String.Equals(code, itemNo, StringComparison.OrdinalIgnoreCase)))
                {
                    chain.Add(itemNo);
                }

                return chain;
            }

            private static bool IsChapterTreeCode(string code)
            {
                return !String.IsNullOrEmpty(code) && code.Length >= 2 && Char.IsDigit(code[0]) && Char.IsDigit(code[1]);
            }

            // itemNo 是否属于编号 code 的条目（本身或下级）。
            // 下级判定：带横杠段（0101 -> 0101-04），或纯数字编号续位（01 -> 0101）。
            private static bool IsItemNoUnderChapter(string itemNo, string code)
            {
                itemNo = (itemNo ?? "").Trim();
                code = (code ?? "").Trim();
                if (itemNo.Length == 0 || code.Length == 0)
                {
                    return false;
                }

                if (String.Equals(itemNo, code, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }

                if (!itemNo.StartsWith(code, StringComparison.OrdinalIgnoreCase) || itemNo.Length <= code.Length)
                {
                    return false;
                }

                char next = itemNo[code.Length];
                if (next == '-')
                {
                    return true;
                }

                return Char.IsDigit(next) && code.All(Char.IsDigit);
            }

            private string GetChapterDisplayName(string code)
            {
                string name;
                if (!chapterNames.TryGetValue(code, out name) || String.IsNullOrEmpty(name))
                {
                    return code;
                }

                // 与软件左侧章节树一致：两位纯数字章（01）显示"一、"，四位及以上（0101）显示"01."；
                // 带横杠的下级条目名称自带"一、/(一)/1."等序号，不再重复。
                if (code.All(Char.IsDigit))
                {
                    if (code.Length == 2)
                    {
                        int value;
                        if (Int32.TryParse(code, out value))
                        {
                            return ToChineseOrdinal(value) + "、" + name;
                        }
                    }
                    else if (code.Length >= 4)
                    {
                        return code.Substring(code.Length - 2) + "." + name;
                    }
                }

                return name;
            }

            private static string ToChineseOrdinal(int value)
            {
                string[] digits = { "零", "一", "二", "三", "四", "五", "六", "七", "八", "九" };
                if (value <= 0 || value > 99)
                {
                    return value.ToString();
                }

                if (value < 10)
                {
                    return digits[value];
                }

                int tens = value / 10;
                int ones = value % 10;
                return (tens == 1 ? "" : digits[tens]) + "十" + (ones == 0 ? "" : digits[ones]);
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

            private static void SetTreeChildrenChecked(TreeNode node, bool value)
            {
                foreach (TreeNode child in node.Nodes)
                {
                    child.Checked = value;
                    SetTreeChildrenChecked(child, value);
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

            private void OnApply()
            {
                try
                {
                    FlushGridSelectionsToPreview();
                    int selectedCount = preview.Count(it => it.Selected);
                    string targetUnit = txtTargetUnit.Text.Trim();
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
