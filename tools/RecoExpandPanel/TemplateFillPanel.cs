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
            private readonly DataGridView grid = new DataGridView();

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
                txtName.SetBounds(438, 12, 120, 23); txtName.Text = "轨道-模板";
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
                txtSheet.SetBounds(75, 79, 120, 23); txtSheet.Text = "方案二";
                AddLabel("目标列", 205, 82, 50);
                txtColumn.SetBounds(255, 79, 50, 23); txtColumn.Text = "E";
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

                grid.SetBounds(12, 132, 876, 436);
                grid.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
                grid.ReadOnly = false; grid.AllowUserToAddRows = false;
                grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                grid.Columns.Add(new DataGridViewCheckBoxColumn { HeaderText = "选", Name = "sel", FillWeight = 6 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "条目", Name = "item", ReadOnly = true, FillWeight = 14 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "定额编号", Name = "code", ReadOnly = true, FillWeight = 16 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "调整", Name = "adj", ReadOnly = true, FillWeight = 14 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "源行项目名", Name = "sname", ReadOnly = true, FillWeight = 18 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "目标行项目名", Name = "tname", ReadOnly = true, FillWeight = 18 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "数量", Name = "qty", ReadOnly = true, FillWeight = 10 });
                grid.Columns.Add(new DataGridViewTextBoxColumn { HeaderText = "状态", Name = "st", ReadOnly = true, FillWeight = 14 });

                Controls.Add(txtUnit); Controls.Add(cmbSourceSheet); Controls.Add(txtName); Controls.Add(btnBuild);
                Controls.Add(cmbTemplate); Controls.Add(btnDeleteTemplate); Controls.Add(cmbMode); Controls.Add(txtSheet); Controls.Add(txtColumn);
                Controls.Add(txtTargetUnit);
                Controls.Add(btnPreview); Controls.Add(btnApply); Controls.Add(reminder); Controls.Add(grid);
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
                try
                {
                    if (cmbTemplate.SelectedItem == null) { MessageBox.Show(this, "请先选择模板。", "模板铺量"); return; }
                    FillTemplate t = LoadFillTemplate(Convert.ToString(cmbTemplate.SelectedItem));
                    if (t == null) { MessageBox.Show(this, "模板加载失败。", "模板铺量"); return; }
                    preview = cmbMode.SelectedIndex == 0
                        ? BuildPreview_ColumnAnchor(t, txtSheet.Text.Trim(), txtColumn.Text.Trim())
                        : BuildPreview_FixedColumn(t);
                    FillGrid();
                    if (preview.Count == 0)
                        MessageBox.Show(this, "预览为空：该模板里没有定额。请回到上一步重新“从该单元生成模板”，并确认收到的定额条数大于 0。", "模板铺量");
                }
                catch (Exception ex) { MessageBox.Show(this, "预览失败：" + ex.Message, "模板铺量"); }
            }

            private void FillGrid()
            {
                grid.Rows.Clear();
                foreach (FillPreviewItem it in preview)
                {
                    int idx = grid.Rows.Add(it.Selected, it.ItemNo, it.QuotaCode, it.Adjust,
                        it.SourceName, it.TargetName, it.QuantityText, it.Status);
                    if (!String.IsNullOrEmpty(it.Status))
                        grid.Rows[idx].DefaultCellStyle.BackColor = Color.MistyRose;
                }
            }

            private void OnApply()
            {
                try
                {
                    for (int i = 0; i < preview.Count && i < grid.Rows.Count; i++)
                        preview[i].Selected = Convert.ToBoolean(grid.Rows[i].Cells["sel"].Value ?? false);

                    string targetUnit = txtTargetUnit.Text.Trim();
                    if (MessageBox.Show(this, "确认把勾选定额复制到目标单元【" + targetUnit + "】的对应条目？",
                        "模板铺量", MessageBoxButtons.OKCancel) != DialogResult.OK) return;

                    string result = ApplyFill(mainForm, targetUnit, preview);
                    MessageBox.Show(this, result, "模板铺量");
                }
                catch (Exception ex) { MessageBox.Show(this, "写入失败：" + ex.Message, "模板铺量"); }
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
