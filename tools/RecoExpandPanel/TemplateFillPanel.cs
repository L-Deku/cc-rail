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
            private readonly TextBox txtUnit = new TextBox();
            private readonly TextBox txtItemPrefix = new TextBox();
            private readonly TextBox txtName = new TextBox();
            private readonly Button btnBuild = new Button();
            private readonly ComboBox cmbMode = new ComboBox();
            private readonly TextBox txtSheet = new TextBox();
            private readonly TextBox txtColumn = new TextBox();
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
            }

            private void BuildLayout()
            {
                // —— 生成模板 ——
                AddLabel("源单元号", 12, 15, 60);
                txtUnit.SetBounds(80, 12, 90, 23); txtUnit.Text = "_ZGS_01";
                AddLabel("专业条目前缀", 180, 15, 80);
                txtItemPrefix.SetBounds(265, 12, 70, 23); txtItemPrefix.Text = "04";
                AddLabel("模板名", 345, 15, 50);
                txtName.SetBounds(400, 12, 130, 23); txtName.Text = "轨道-模板";
                btnBuild.SetBounds(540, 11, 130, 25); btnBuild.Text = "从该单元生成模板";
                btnBuild.Click += delegate { OnBuild(); };

                // —— 套用配置 ——
                AddLabel("模板", 12, 50, 40);
                cmbTemplate.SetBounds(55, 47, 200, 23); cmbTemplate.DropDownStyle = ComboBoxStyle.DropDownList;
                AddLabel("取数模式", 270, 50, 60);
                cmbMode.SetBounds(335, 47, 150, 23); cmbMode.DropDownStyle = ComboBoxStyle.DropDownList;
                cmbMode.Items.AddRange(new object[] { "一·列锚点", "二·固定绑定列" });
                cmbMode.SelectedIndex = 0;
                AddLabel("目标sheet", 12, 82, 60);
                txtSheet.SetBounds(75, 79, 120, 23); txtSheet.Text = "方案二";
                AddLabel("目标列", 205, 82, 50);
                txtColumn.SetBounds(255, 79, 50, 23); txtColumn.Text = "E";
                btnPreview.SetBounds(320, 78, 80, 25); btnPreview.Text = "预览";
                btnPreview.Click += delegate { OnPreview(); };
                btnApply.SetBounds(410, 78, 150, 25); btnApply.Text = "写入当前单元";
                btnApply.Click += delegate { OnApply(); };

                Label reminder = new Label
                {
                    Text = "注意：写入的是软件【当前单元】。请先在软件顶部总概算下拉切到目标单元，再点“写入当前单元”。",
                    ForeColor = Color.Firebrick, AutoSize = false
                };
                reminder.SetBounds(12, 108, 876, 18);

                grid.SetBounds(12, 132, 876, 436);
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

                Controls.Add(txtUnit); Controls.Add(txtItemPrefix); Controls.Add(txtName); Controls.Add(btnBuild);
                Controls.Add(cmbTemplate); Controls.Add(cmbMode); Controls.Add(txtSheet); Controls.Add(txtColumn);
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

            private void OnBuild()
            {
                try
                {
                    using (SqlConnection conn = GetProjectConnection(mainForm))
                    {
                        FillTemplate t = BuildFillTemplateFromUnit(mainForm, conn, txtName.Text.Trim(),
                            txtItemPrefix.Text.Trim(), txtUnit.Text.Trim(), txtItemPrefix.Text.Trim());
                        SaveFillTemplate(t);
                    }
                    ReloadTemplateList();
                    MessageBox.Show(this, "模板已生成并保存。", "模板铺量");
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

                    if (MessageBox.Show(this, "确认把勾选定额追加到软件【当前单元】？\n请确认软件已切到目标单元。",
                        "模板铺量", MessageBoxButtons.OKCancel) != DialogResult.OK) return;

                    string result = ApplyFill(mainForm, preview);
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
