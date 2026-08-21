using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        private static readonly Dictionary<Form, AgentPanelWindow> AgentPanelWindows = new Dictionary<Form, AgentPanelWindow>();

        // 手填单价的补充定额代码；连同"超过4位纯数字=材料编号"一起，用来在候选列表里标出类型。
        private static readonly string[] AgentSupplementCodes = new string[] { "SF", "SH", "SQ", "ZLF", "LF", "TLF" };

        private static void ShowAgentPanelWindow(Form mainForm)
        {
            AgentPanelWindow window;
            if (!AgentPanelWindows.TryGetValue(mainForm, out window) || window.IsDisposed)
            {
                window = new AgentPanelWindow(mainForm);
                AgentPanelWindows[mainForm] = window;
                mainForm.FormClosed += delegate
                {
                    AgentPanelWindows.Remove(mainForm);
                };
            }

            if (!window.Visible)
            {
                window.Show(mainForm);
            }

            window.BringToFront();
            window.OnActivatedFromHost();
        }

        // 定额编号属于哪一类：材料编号（超过4位纯数字）、补充定额（SF/SH/SQ/ZLF/LF/TLF，手填单价）、普通定额。
        private static string AgentQuotaCodeKind(string code)
        {
            string text = (code ?? "").Trim();
            if (text.Length == 0)
            {
                return "";
            }

            string baseCode = text;
            int cut = baseCode.IndexOfAny(new char[] { '*', '/' });
            if (cut > 0)
            {
                baseCode = baseCode.Substring(0, cut);
            }

            bool allDigits = baseCode.Length > 0;
            for (int i = 0; i < baseCode.Length; i++)
            {
                if (!Char.IsDigit(baseCode[i]))
                {
                    allDigits = false;
                    break;
                }
            }

            if (allDigits)
            {
                return baseCode.Length > 4 ? "材料" : "定额";
            }

            foreach (string supplement in AgentSupplementCodes)
            {
                if (String.Equals(baseCode, supplement, StringComparison.OrdinalIgnoreCase))
                {
                    return "补充";
                }
            }

            return "定额";
        }

        // 一个可选目标：范围内出现过的 定额编号 / 名称 组合。
        private sealed class AgentQuotaCandidate
        {
            public string Code = "";
            public string Name = "";
        }

        // 范围（条目 × 单元）内出现过的定额编号与名称，供"按编号/按名称"两种挑选模式用。
        private static List<AgentQuotaCandidate> LoadAgentQuotaCandidates(SqlConnection conn, List<string> itemNos,
            bool includeChildren, List<long> unitIds)
        {
            List<AgentQuotaCandidate> candidates = new List<AgentQuotaCandidate>();
            HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
            List<string> scope = (itemNos == null || itemNos.Count == 0) ? new List<string> { null } : itemNos;

            foreach (string itemNo in scope)
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    string where = String.IsNullOrEmpty(itemNo)
                        ? "1=1"
                        : BuildAgentItemCondition(cmd, itemNo, includeChildren, 0);
                    cmd.CommandText = "select distinct DE.定额编号, DE.工程或费用项目名称 " +
                        "from 定额输入 DE inner join 章节表 ZJ on DE.条目序号=ZJ.条目序号 " +
                        "where " + where + BuildAgentUnitCondition(unitIds) +
                        " order by DE.定额编号, DE.工程或费用项目名称";
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            AgentQuotaCandidate candidate = new AgentQuotaCandidate();
                            candidate.Code = reader.IsDBNull(0) ? "" : Convert.ToString(reader.GetValue(0)).Trim();
                            candidate.Name = reader.IsDBNull(1) ? "" : Convert.ToString(reader.GetValue(1)).Trim();
                            if (candidate.Code.Length == 0 && candidate.Name.Length == 0)
                            {
                                continue;
                            }

                            if (seen.Add(candidate.Code + "" + candidate.Name))
                            {
                                candidates.Add(candidate);
                            }
                        }
                    }
                }
            }

            return candidates;
        }

        // 选择对话框里的一项：Key 是真正参与命令的值，Display 是给人看的。
        private sealed class AgentPickItem
        {
            public string Key = "";
            public string Display = "";

            public override string ToString()
            {
                return Display;
            }
        }

        // 通用挑选对话框：上面一个过滤框，中间列表，可单选或多选。
        private sealed class AgentPickerDialog : Form
        {
            private readonly List<AgentPickItem> allItems;
            private readonly bool multiSelect;
            private readonly TextBox filterBox;
            private readonly CheckedListBox checkedList;
            private readonly ListBox singleList;
            private readonly Label countLabel;
            private readonly HashSet<string> checkedKeys = new HashSet<string>(StringComparer.Ordinal);

            public List<string> SelectedKeys = new List<string>();

            public AgentPickerDialog(string title, List<AgentPickItem> items, bool multiSelect, List<string> preselected)
            {
                this.allItems = items;
                this.multiSelect = multiSelect;
                if (preselected != null)
                {
                    foreach (string key in preselected)
                    {
                        checkedKeys.Add(key);
                    }
                }

                Text = title;
                StartPosition = FormStartPosition.CenterParent;
                ClientSize = new Size(620, 500);
                MinimizeBox = false;
                MaximizeBox = false;
                ShowInTaskbar = false;
                FormBorderStyle = FormBorderStyle.SizableToolWindow;

                Panel top = new Panel();
                top.Dock = DockStyle.Top;
                top.Height = 34;
                top.Padding = new Padding(8, 4, 8, 4);

                filterBox = new TextBox();
                filterBox.Dock = DockStyle.Fill;
                filterBox.TextChanged += delegate { RefillList(); };
                top.Controls.Add(filterBox);

                Label filterHint = new Label();
                filterHint.Text = "筛选：";
                filterHint.Dock = DockStyle.Left;
                filterHint.Width = 44;
                filterHint.TextAlign = ContentAlignment.MiddleLeft;
                top.Controls.Add(filterHint);

                Panel bottom = new Panel();
                bottom.Dock = DockStyle.Bottom;
                bottom.Height = 42;
                bottom.Padding = new Padding(8, 6, 8, 6);

                Button ok = new Button();
                ok.Text = "确定";
                ok.Width = 84;
                ok.Dock = DockStyle.Right;
                ok.DialogResult = DialogResult.OK;
                ok.Click += delegate { CollectResult(); };

                Button cancel = new Button();
                cancel.Text = "取消";
                cancel.Width = 84;
                cancel.Dock = DockStyle.Right;
                cancel.DialogResult = DialogResult.Cancel;

                Panel spacer = new Panel();
                spacer.Dock = DockStyle.Right;
                spacer.Width = 8;

                countLabel = new Label();
                countLabel.Dock = DockStyle.Fill;
                countLabel.TextAlign = ContentAlignment.MiddleLeft;
                countLabel.ForeColor = Color.Gray;

                bottom.Controls.Add(countLabel);
                bottom.Controls.Add(cancel);
                bottom.Controls.Add(spacer);
                bottom.Controls.Add(ok);

                if (multiSelect)
                {
                    checkedList = new CheckedListBox();
                    checkedList.Dock = DockStyle.Fill;
                    checkedList.CheckOnClick = true;
                    checkedList.IntegralHeight = false;
                    checkedList.ItemCheck += delegate(object sender, ItemCheckEventArgs e)
                    {
                        AgentPickItem item = checkedList.Items[e.Index] as AgentPickItem;
                        if (item == null)
                        {
                            return;
                        }

                        if (e.NewValue == CheckState.Checked)
                        {
                            checkedKeys.Add(item.Key);
                        }
                        else
                        {
                            checkedKeys.Remove(item.Key);
                        }

                        BeginInvoke((MethodInvoker)delegate { UpdateCount(); });
                    };
                    Controls.Add(checkedList);
                }
                else
                {
                    singleList = new ListBox();
                    singleList.Dock = DockStyle.Fill;
                    singleList.IntegralHeight = false;
                    singleList.DoubleClick += delegate
                    {
                        if (singleList.SelectedItem != null)
                        {
                            CollectResult();
                            DialogResult = DialogResult.OK;
                        }
                    };
                    Controls.Add(singleList);
                }

                Controls.Add(top);
                Controls.Add(bottom);
                AcceptButton = ok;
                CancelButton = cancel;

                RefillList();
                UpdateCount();
            }

            private void RefillList()
            {
                string filter = (filterBox.Text ?? "").Trim();
                List<AgentPickItem> shown = new List<AgentPickItem>();
                foreach (AgentPickItem item in allItems)
                {
                    if (filter.Length == 0 || item.Display.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        shown.Add(item);
                    }
                }

                if (multiSelect)
                {
                    checkedList.BeginUpdate();
                    checkedList.Items.Clear();
                    foreach (AgentPickItem item in shown)
                    {
                        checkedList.Items.Add(item, checkedKeys.Contains(item.Key));
                    }

                    checkedList.EndUpdate();
                }
                else
                {
                    singleList.BeginUpdate();
                    singleList.Items.Clear();
                    foreach (AgentPickItem item in shown)
                    {
                        singleList.Items.Add(item);
                    }

                    singleList.EndUpdate();
                    if (singleList.Items.Count > 0)
                    {
                        singleList.SelectedIndex = 0;
                    }
                }
            }

            private void UpdateCount()
            {
                if (multiSelect)
                {
                    countLabel.Text = "已选 " + checkedKeys.Count.ToString(CultureInfo.InvariantCulture) +
                        " 项 / 共 " + allItems.Count.ToString(CultureInfo.InvariantCulture) + " 项";
                }
                else
                {
                    countLabel.Text = "共 " + allItems.Count.ToString(CultureInfo.InvariantCulture) + " 项";
                }
            }

            private void CollectResult()
            {
                SelectedKeys = new List<string>();
                if (multiSelect)
                {
                    foreach (AgentPickItem item in allItems)
                    {
                        if (checkedKeys.Contains(item.Key))
                        {
                            SelectedKeys.Add(item.Key);
                        }
                    }
                }
                else
                {
                    AgentPickItem picked = singleList.SelectedItem as AgentPickItem;
                    if (picked != null)
                    {
                        SelectedKeys.Add(picked.Key);
                    }
                }
            }
        }

        // 点选式智能指令窗口。定位是跨条目、跨单元的批量操作：
        // 单个条目内改数量/编号/单价/调整，主程序右键"乘系数"已经能做，这里不重复。
        private sealed class AgentPanelWindow : Form
        {
            private const string TabValue = "改数值";
            private const string TabRows = "加减定额";
            private const string TabCross = "跨条目";
            private const string TabText = "说一句话";

            private readonly Form mainForm;

            // 第一区：作用范围
            private readonly TextBox unitBox;
            private readonly CheckBox allUnitsBox;
            private readonly ListBox itemListBox;
            private readonly CheckBox includeChildrenBox;
            private readonly RadioButton byCodeRadio;
            private readonly RadioButton byNameRadio;
            private readonly Button pickTargetButton;
            private readonly Button takeFromGridButton;
            private readonly ListBox targetListBox;
            private readonly Label targetSummaryLabel;

            // 第二区：操作
            private readonly TabControl tabs;
            private readonly ComboBox fieldBox;
            private readonly ComboBox valueActionBox;
            private readonly Label valueLabel;
            private readonly TextBox valueBox;
            private readonly Label valueHintLabel;

            private readonly ComboBox rowActionBox;
            private readonly Label quotaGridLabel;
            private readonly DataGridView quotaGrid;
            private readonly Label rowHintLabel;

            private readonly ComboBox crossActionBox;
            private readonly TextBox crossTargetBox;
            private readonly Label crossHintLabel;

            private readonly TextBox textBox;
            private readonly Button textSendButton;
            private readonly Label textHintLabel;

            // 底部
            private readonly Label sentenceLabel;
            private readonly Button undoButton;
            private readonly Button redoButton;
            private readonly Button previewButton;
            private readonly Panel previewPanel;
            private readonly Label summaryLabel;
            private readonly DataGridView previewGrid;
            private readonly Label statusLabel;
            private readonly ToolTip toolTip = new ToolTip();

            private readonly List<AgentUnitOption> unitOptions = new List<AgentUnitOption>();
            private readonly List<AgentItemOption> itemOptions = new List<AgentItemOption>();
            private readonly List<string> selectedUnitKeys = new List<string>();
            private readonly List<string> selectedItemNos = new List<string>();
            private readonly List<string> selectedTargetCodes = new List<string>();
            private readonly List<string> selectedTargetNames = new List<string>();
            private readonly List<string> crossTargetItems = new List<string>();

            private AgentPlan pendingPlan;
            private bool busy;
            private bool suppressSentence;
            private string scopeProjectIdentity = "";

            public AgentPanelWindow(Form mainForm)
            {
                this.mainForm = mainForm;
                Text = "智能指令助手 (Ctrl+Q)　—　跨条目 / 跨单元批量操作";
                StartPosition = FormStartPosition.Manual;
                Size = new Size(960, 850);
                MinimumSize = new Size(820, 680);
                ShowInTaskbar = false;
                try
                {
                    Location = new Point(
                        Math.Max(0, mainForm.Right - Width - 16),
                        Math.Max(0, mainForm.Top + 30));
                }
                catch
                {
                    StartPosition = FormStartPosition.CenterParent;
                }

                // ===== 第一区：作用范围 =====
                Panel scopePanel = new Panel();
                scopePanel.Dock = DockStyle.Top;
                scopePanel.Height = 248;
                scopePanel.Padding = new Padding(10, 6, 10, 4);

                Label scopeTitle = new Label();
                scopeTitle.Text = "第一步  作用范围（这一区配一次，下面所有操作共用）";
                scopeTitle.Dock = DockStyle.Top;
                scopeTitle.Height = 22;
                scopeTitle.TextAlign = ContentAlignment.MiddleLeft;
                scopeTitle.ForeColor = Color.FromArgb(20, 60, 160);

                Panel unitRow = new Panel();
                unitRow.Dock = DockStyle.Top;
                unitRow.Height = 30;

                Label unitLabel = new Label();
                unitLabel.Text = "单元";
                unitLabel.SetBounds(0, 6, 34, 20);

                unitBox = new TextBox();
                unitBox.ReadOnly = true;
                unitBox.BackColor = SystemColors.Window;
                unitBox.SetBounds(38, 3, 500, 24);
                unitBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

                Button unitButton = new Button();
                unitButton.Text = "选单元";
                unitButton.SetBounds(544, 2, 76, 26);
                unitButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                unitButton.Click += delegate { PickUnits(); };

                allUnitsBox = new CheckBox();
                allUnitsBox.Text = "所有单元";
                allUnitsBox.SetBounds(626, 5, 84, 22);
                allUnitsBox.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                allUnitsBox.CheckedChanged += delegate
                {
                    unitButton.Enabled = !allUnitsBox.Checked;
                    UpdateUnitBox();
                    ClearTargetSelection("单元变了，目标定额已清空，请重新选。");
                    UpdateSentence();
                };

                unitRow.Controls.Add(unitBox);
                unitRow.Controls.Add(unitLabel);
                unitRow.Controls.Add(unitButton);
                unitRow.Controls.Add(allUnitsBox);

                // --- 条目：在左侧树点一个条目，点"加入"累积；每个条目含其下全部子条目 ---
                Panel itemPanel = new Panel();
                itemPanel.Dock = DockStyle.Top;
                itemPanel.Height = 86;

                Label itemLabel = new Label();
                itemLabel.Text = "条目";
                itemLabel.SetBounds(0, 4, 34, 20);

                itemListBox = new ListBox();
                itemListBox.SetBounds(38, 2, 500, 80);
                itemListBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
                itemListBox.IntegralHeight = false;
                itemListBox.SelectionMode = SelectionMode.MultiExtended;

                Button addItemButton = new Button();
                addItemButton.Text = "加入树上选中";
                addItemButton.SetBounds(544, 2, 104, 26);
                addItemButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                addItemButton.Click += delegate { AddItemFromTree(false); };

                Button removeItemButton = new Button();
                removeItemButton.Text = "移除";
                removeItemButton.SetBounds(544, 30, 50, 26);
                removeItemButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                removeItemButton.Click += delegate { RemoveSelectedItems(); };

                Button clearItemButton = new Button();
                clearItemButton.Text = "清空";
                clearItemButton.SetBounds(598, 30, 50, 26);
                clearItemButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                clearItemButton.Click += delegate { ClearItems(); };

                includeChildrenBox = new CheckBox();
                includeChildrenBox.Text = "含所有子条目";
                includeChildrenBox.Checked = true;
                includeChildrenBox.SetBounds(544, 58, 110, 22);
                includeChildrenBox.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                includeChildrenBox.CheckedChanged += delegate
                {
                    ClearTargetSelection("范围变了，目标定额已清空，请重新选。");
                    UpdateSentence();
                };

                itemPanel.Controls.Add(itemListBox);
                itemPanel.Controls.Add(itemLabel);
                itemPanel.Controls.Add(addItemButton);
                itemPanel.Controls.Add(removeItemButton);
                itemPanel.Controls.Add(clearItemButton);
                itemPanel.Controls.Add(includeChildrenBox);

                // --- 目标定额/材料/SF：按编号 或 按名称 ---
                Panel targetHeaderRow = new Panel();
                targetHeaderRow.Dock = DockStyle.Top;
                targetHeaderRow.Height = 28;

                Label targetLabel = new Label();
                targetLabel.Text = "目标";
                targetLabel.SetBounds(0, 5, 34, 20);

                byCodeRadio = new RadioButton();
                byCodeRadio.Text = "按编号";
                byCodeRadio.Checked = true;
                byCodeRadio.SetBounds(38, 4, 74, 20);
                byCodeRadio.CheckedChanged += delegate
                {
                    if (byCodeRadio.Checked)
                    {
                        ClearTargetSelection(null);
                        UpdateSentence();
                    }
                };

                byNameRadio = new RadioButton();
                byNameRadio.Text = "按名称";
                byNameRadio.SetBounds(118, 4, 74, 20);
                byNameRadio.CheckedChanged += delegate
                {
                    if (byNameRadio.Checked)
                    {
                        ClearTargetSelection(null);
                        UpdateSentence();
                    }
                };

                pickTargetButton = new Button();
                pickTargetButton.Text = "从范围内列表选…";
                pickTargetButton.SetBounds(198, 1, 122, 24);
                pickTargetButton.Click += delegate { PickTargets(); };

                takeFromGridButton = new Button();
                takeFromGridButton.Text = "取定额表选中行";
                takeFromGridButton.SetBounds(326, 1, 116, 24);
                takeFromGridButton.Click += delegate { TakeTargetsFromHostGrid(); };

                Button clearTargetButton = new Button();
                clearTargetButton.Text = "清空";
                clearTargetButton.SetBounds(448, 1, 56, 24);
                clearTargetButton.Click += delegate
                {
                    ClearTargetSelection("目标已清空，将作用于范围内全部行。");
                    UpdateSentence();
                };

                targetSummaryLabel = new Label();
                targetSummaryLabel.SetBounds(512, 5, 300, 20);
                targetSummaryLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                targetSummaryLabel.ForeColor = Color.Gray;

                targetHeaderRow.Controls.Add(targetLabel);
                targetHeaderRow.Controls.Add(byCodeRadio);
                targetHeaderRow.Controls.Add(byNameRadio);
                targetHeaderRow.Controls.Add(pickTargetButton);
                targetHeaderRow.Controls.Add(takeFromGridButton);
                targetHeaderRow.Controls.Add(clearTargetButton);
                targetHeaderRow.Controls.Add(targetSummaryLabel);

                targetListBox = new ListBox();
                targetListBox.Dock = DockStyle.Fill;
                targetListBox.IntegralHeight = false;
                targetListBox.SelectionMode = SelectionMode.None;
                targetListBox.BackColor = SystemColors.Control;
                targetListBox.BorderStyle = BorderStyle.FixedSingle;

                scopePanel.Controls.Add(targetListBox);
                scopePanel.Controls.Add(targetHeaderRow);
                scopePanel.Controls.Add(itemPanel);
                scopePanel.Controls.Add(unitRow);
                scopePanel.Controls.Add(scopeTitle);

                // ===== 第二区：做什么 =====
                tabs = new TabControl();
                tabs.Dock = DockStyle.Fill;
                tabs.Padding = new Point(14, 4);

                // --- 页签 1：改数值 ---
                TabPage valuePage = new TabPage(TabValue);
                valuePage.Padding = new Padding(12, 14, 12, 8);
                valuePage.UseVisualStyleBackColor = true;

                Label fieldLabel = new Label();
                fieldLabel.Text = "字段";
                fieldLabel.SetBounds(4, 10, 34, 20);

                fieldBox = new ComboBox();
                fieldBox.DropDownStyle = ComboBoxStyle.DropDownList;
                fieldBox.SetBounds(42, 6, 120, 24);
                fieldBox.Items.AddRange(new object[] { "工程数量", "单价", "定额编号", "定额调整" });
                fieldBox.SelectedIndex = 0;
                fieldBox.SelectedIndexChanged += delegate { RefreshValueActions(); };

                Label actionLabel = new Label();
                actionLabel.Text = "操作";
                actionLabel.SetBounds(180, 10, 34, 20);

                valueActionBox = new ComboBox();
                valueActionBox.DropDownStyle = ComboBoxStyle.DropDownList;
                valueActionBox.SetBounds(218, 6, 120, 24);
                valueActionBox.SelectedIndexChanged += delegate { RefreshValueInput(); };

                valueLabel = new Label();
                valueLabel.Text = "系数";
                valueLabel.SetBounds(356, 10, 62, 20);

                valueBox = new TextBox();
                valueBox.SetBounds(420, 6, 140, 24);
                valueBox.TextChanged += delegate { UpdateSentence(); };

                valueHintLabel = new Label();
                valueHintLabel.SetBounds(4, 46, 700, 76);
                valueHintLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                valueHintLabel.ForeColor = Color.Gray;

                valuePage.Controls.Add(fieldLabel);
                valuePage.Controls.Add(fieldBox);
                valuePage.Controls.Add(actionLabel);
                valuePage.Controls.Add(valueActionBox);
                valuePage.Controls.Add(valueLabel);
                valuePage.Controls.Add(valueBox);
                valuePage.Controls.Add(valueHintLabel);

                // --- 页签 2：加减定额 ---
                TabPage rowPage = new TabPage(TabRows);
                rowPage.Padding = new Padding(12, 14, 12, 8);
                rowPage.UseVisualStyleBackColor = true;

                Label rowActionLabel = new Label();
                rowActionLabel.Text = "动作";
                rowActionLabel.SetBounds(4, 10, 34, 20);

                rowActionBox = new ComboBox();
                rowActionBox.DropDownStyle = ComboBoxStyle.DropDownList;
                rowActionBox.SetBounds(42, 6, 160, 24);
                rowActionBox.Items.AddRange(new object[] { "替换目标定额", "新增定额", "删除目标定额" });
                rowActionBox.SelectedIndex = 0;
                rowActionBox.SelectedIndexChanged += delegate { RefreshRowAction(); };

                rowHintLabel = new Label();
                rowHintLabel.SetBounds(214, 10, 500, 20);
                rowHintLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                rowHintLabel.ForeColor = Color.Gray;

                quotaGridLabel = new Label();
                quotaGridLabel.SetBounds(4, 40, 400, 18);

                quotaGrid = new DataGridView();
                quotaGrid.SetBounds(4, 60, 700, 150);
                quotaGrid.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
                quotaGrid.AllowUserToAddRows = true;
                quotaGrid.AllowUserToDeleteRows = true;
                quotaGrid.RowHeadersVisible = false;
                quotaGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                quotaGrid.Columns.Add("Code", "定额编号");
                quotaGrid.Columns.Add("Quantity", "工程数量");
                quotaGrid.Columns["Code"].FillWeight = 62;
                quotaGrid.Columns["Quantity"].FillWeight = 38;
                quotaGrid.CellEndEdit += delegate { UpdateSentence(); };
                quotaGrid.UserDeletedRow += delegate { UpdateSentence(); };

                rowPage.Controls.Add(rowActionLabel);
                rowPage.Controls.Add(rowActionBox);
                rowPage.Controls.Add(rowHintLabel);
                rowPage.Controls.Add(quotaGridLabel);
                rowPage.Controls.Add(quotaGrid);

                // --- 页签 3：跨条目 ---
                TabPage crossPage = new TabPage(TabCross);
                crossPage.Padding = new Padding(12, 14, 12, 8);
                crossPage.UseVisualStyleBackColor = true;

                Label crossActionLabel = new Label();
                crossActionLabel.Text = "动作";
                crossActionLabel.SetBounds(4, 10, 34, 20);

                crossActionBox = new ComboBox();
                crossActionBox.DropDownStyle = ComboBoxStyle.DropDownList;
                crossActionBox.SetBounds(42, 6, 120, 24);
                crossActionBox.Items.AddRange(new object[] { "复制到", "移动到" });
                crossActionBox.SelectedIndex = 0;
                crossActionBox.SelectedIndexChanged += delegate { RefreshCrossAction(); };

                Label crossTargetLabel = new Label();
                crossTargetLabel.Text = "目标条目";
                crossTargetLabel.SetBounds(180, 10, 62, 20);

                crossTargetBox = new TextBox();
                crossTargetBox.ReadOnly = true;
                crossTargetBox.BackColor = SystemColors.Window;
                crossTargetBox.SetBounds(246, 6, 356, 24);
                crossTargetBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

                Button crossTargetButton = new Button();
                crossTargetButton.Text = "选条目";
                crossTargetButton.SetBounds(608, 5, 78, 26);
                crossTargetButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                crossTargetButton.Click += delegate { PickCrossTargets(); };

                crossHintLabel = new Label();
                crossHintLabel.SetBounds(4, 46, 700, 76);
                crossHintLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                crossHintLabel.ForeColor = Color.Gray;

                crossPage.Controls.Add(crossActionLabel);
                crossPage.Controls.Add(crossActionBox);
                crossPage.Controls.Add(crossTargetLabel);
                crossPage.Controls.Add(crossTargetBox);
                crossPage.Controls.Add(crossTargetButton);
                crossPage.Controls.Add(crossHintLabel);

                // --- 页签 4：说一句话 ---
                TabPage textPage = new TabPage(TabText);
                textPage.Padding = new Padding(12, 14, 12, 8);
                textPage.UseVisualStyleBackColor = true;

                textBox = new TextBox();
                textBox.Multiline = true;
                textBox.ScrollBars = ScrollBars.Vertical;
                textBox.SetBounds(4, 6, 590, 62);
                textBox.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                textBox.Font = new Font(Font.FontFamily, 10.5f);

                textSendButton = new Button();
                textSendButton.Text = "交给 AI";
                textSendButton.SetBounds(602, 6, 88, 30);
                textSendButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                textSendButton.Click += delegate { SubmitText(); };

                Button helpButton = new Button();
                helpButton.Text = "指令帮助";
                helpButton.SetBounds(602, 40, 88, 28);
                helpButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                helpButton.Click += delegate { ShowHelpDialog(); };

                textHintLabel = new Label();
                textHintLabel.SetBounds(4, 74, 690, 90);
                textHintLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                textHintLabel.ForeColor = Color.Gray;
                textHintLabel.Text = "兜底通道：上面的按钮拼不出来时才用（新建单元、运输方案、材料价方案只能从这里走）。\r\n" +
                    "也可以直接输入 撤销 / 重做 / 帮助 / 探查 关键词。\r\n" +
                    "自然语言需要已配置 RecoQuotaData/deepseek-settings.json；确定性文本语法（如 工程数量 0101-01 *0.85）不需要 AI。";

                textPage.Controls.Add(textBox);
                textPage.Controls.Add(textSendButton);
                textPage.Controls.Add(helpButton);
                textPage.Controls.Add(textHintLabel);

                tabs.TabPages.Add(valuePage);
                tabs.TabPages.Add(rowPage);
                tabs.TabPages.Add(crossPage);
                tabs.TabPages.Add(textPage);
                tabs.SelectedIndexChanged += delegate
                {
                    RefreshScopeAvailability();
                    UpdateSentence();
                };

                // ===== 操作条 =====
                Panel actionBar = new Panel();
                actionBar.Dock = DockStyle.Bottom;
                actionBar.Height = 66;
                actionBar.Padding = new Padding(10, 4, 10, 4);

                sentenceLabel = new Label();
                sentenceLabel.Dock = DockStyle.Top;
                sentenceLabel.Height = 30;
                sentenceLabel.TextAlign = ContentAlignment.MiddleLeft;
                sentenceLabel.ForeColor = Color.FromArgb(20, 60, 160);

                Panel buttonRow = new Panel();
                buttonRow.Dock = DockStyle.Fill;

                undoButton = new Button();
                undoButton.Text = "撤销上一步";
                undoButton.SetBounds(0, 1, 108, 28);
                undoButton.Click += delegate { PreviewUndo(); };

                redoButton = new Button();
                redoButton.Text = "重做";
                redoButton.SetBounds(114, 1, 68, 28);
                redoButton.Click += delegate { PreviewRedo(); };

                previewButton = new Button();
                previewButton.Text = "生成预览";
                previewButton.Width = 110;
                previewButton.Height = 28;
                previewButton.Top = 1;
                previewButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                previewButton.BackColor = Color.FromArgb(222, 236, 252);
                previewButton.Click += delegate { SubmitPanelCommand(); };

                buttonRow.Controls.Add(undoButton);
                buttonRow.Controls.Add(redoButton);
                buttonRow.Controls.Add(previewButton);
                buttonRow.Resize += delegate
                {
                    previewButton.Left = Math.Max(190, buttonRow.Width - previewButton.Width - 4);
                };

                actionBar.Controls.Add(buttonRow);
                actionBar.Controls.Add(sentenceLabel);

                // ===== 第三区：确认执行 =====
                previewPanel = new Panel();
                previewPanel.Dock = DockStyle.Bottom;
                previewPanel.Height = 250;
                previewPanel.Visible = false;
                previewPanel.Padding = new Padding(10, 4, 10, 4);

                summaryLabel = new Label();
                summaryLabel.Dock = DockStyle.Top;
                summaryLabel.Height = 30;
                summaryLabel.TextAlign = ContentAlignment.MiddleLeft;
                summaryLabel.ForeColor = Color.FromArgb(160, 80, 0);

                previewGrid = new DataGridView();
                previewGrid.Dock = DockStyle.Fill;
                previewGrid.ReadOnly = true;
                previewGrid.AllowUserToAddRows = false;
                previewGrid.AllowUserToDeleteRows = false;
                previewGrid.RowHeadersVisible = false;
                previewGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                previewGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                previewGrid.Columns.Add("Action", "操作");
                previewGrid.Columns.Add("Unit", "单元");
                previewGrid.Columns.Add("Item", "工程或费用项目名称");
                previewGrid.Columns.Add("Code", "定额编号");
                previewGrid.Columns.Add("Old", "原值");
                previewGrid.Columns.Add("New", "新值");
                previewGrid.Columns["Action"].FillWeight = 13;
                previewGrid.Columns["Unit"].FillWeight = 9;
                previewGrid.Columns["Item"].FillWeight = 28;
                previewGrid.Columns["Code"].FillWeight = 14;
                previewGrid.Columns["Old"].FillWeight = 18;
                previewGrid.Columns["New"].FillWeight = 18;

                Panel confirmRow = new Panel();
                confirmRow.Dock = DockStyle.Bottom;
                confirmRow.Height = 40;

                Button confirmButton = new Button();
                confirmButton.Text = "确认执行";
                confirmButton.Width = 110;
                confirmButton.Height = 30;
                confirmButton.Top = 5;
                confirmButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                confirmButton.BackColor = Color.FromArgb(220, 240, 220);
                confirmButton.Click += delegate { ConfirmPlan(); };

                Button cancelButton = new Button();
                cancelButton.Text = "取消";
                cancelButton.Width = 80;
                cancelButton.Height = 30;
                cancelButton.Top = 5;
                cancelButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                cancelButton.Click += delegate { CancelPlan("已取消，未执行任何修改。"); };

                confirmRow.Controls.Add(confirmButton);
                confirmRow.Controls.Add(cancelButton);
                confirmRow.Resize += delegate
                {
                    confirmButton.Left = confirmRow.Width - confirmButton.Width - cancelButton.Width - 20;
                    cancelButton.Left = confirmRow.Width - cancelButton.Width - 8;
                };

                previewPanel.Controls.Add(previewGrid);
                previewPanel.Controls.Add(summaryLabel);
                previewPanel.Controls.Add(confirmRow);

                statusLabel = new Label();
                statusLabel.Dock = DockStyle.Bottom;
                statusLabel.Height = 24;
                statusLabel.TextAlign = ContentAlignment.MiddleLeft;
                statusLabel.BorderStyle = BorderStyle.FixedSingle;
                statusLabel.Padding = new Padding(6, 0, 0, 0);

                Controls.Add(tabs);
                Controls.Add(actionBar);
                Controls.Add(previewPanel);
                Controls.Add(statusLabel);
                Controls.Add(scopePanel);

                FormClosing += delegate(object sender, FormClosingEventArgs e)
                {
                    if (e.CloseReason == CloseReason.UserClosing)
                    {
                        e.Cancel = true;
                        Hide();
                    }
                };

                RefreshValueActions();
                RefreshRowAction();
                RefreshCrossAction();
                LoadScopeOptions();
                UpdateTargetDisplay();
                RefreshScopeAvailability();
                RefreshUndoRedoButtons();
                UpdateSentence();
                SetStatus("用法：左侧树点中条目 → \"加入树上选中\"（可反复加多个）；选单元；" +
                    "目标可以从定额表多选几行后点\"取定额表选中行\"，按编号或精确名称在整个范围内查找。", false);
            }

            // ===== 打开/激活 =====

            public void OnActivatedFromHost()
            {
                ReloadScopeOptionsIfProjectChanged();
                if (selectedItemNos.Count == 0)
                {
                    AddItemFromTree(true);
                }

                RefreshUndoRedoButtons();
                UpdateSentence();
            }

            // 主程序换项目后，单元/条目候选必须重取，否则还是上一个项目的数据。
            private void ReloadScopeOptionsIfProjectChanged()
            {
                string identity;
                try
                {
                    identity = GetProjectConnectionIdentity(GetOpenProjectConnection(mainForm));
                }
                catch (Exception)
                {
                    return;
                }

                if (String.Equals(identity, scopeProjectIdentity, StringComparison.OrdinalIgnoreCase) && unitOptions.Count > 0)
                {
                    return;
                }

                selectedUnitKeys.Clear();
                selectedItemNos.Clear();
                crossTargetItems.Clear();
                crossTargetBox.Text = "";
                ClearTargetSelection(null);
                UpdateItemList();
                LoadScopeOptions();
            }

            // ===== 作用范围 =====

            private void LoadScopeOptions()
            {
                unitOptions.Clear();
                itemOptions.Clear();
                scopeProjectIdentity = "";
                try
                {
                    SqlConnection conn = GetOpenProjectConnection(mainForm);
                    unitOptions.AddRange(LoadAgentUnitOptions(conn));
                    itemOptions.AddRange(LoadAgentItemOptions(conn));
                    scopeProjectIdentity = GetProjectConnectionIdentity(conn);
                }
                catch (AgentPlanException ex)
                {
                    SetStatus(ex.Message, true);
                    return;
                }
                catch (Exception ex)
                {
                    SetStatus("读取项目单元/条目失败：" + ex.Message, true);
                    Log("Agent panel load scope failed: " + ex);
                    return;
                }

                AgentSelectionSnapshot snapshot = CaptureAgentSelectionForPanel();
                selectedUnitKeys.Clear();
                if (snapshot.CurrentUnitId > 0)
                {
                    foreach (AgentUnitOption unit in unitOptions)
                    {
                        if (unit.UnitId == snapshot.CurrentUnitId)
                        {
                            selectedUnitKeys.Add(UnitKey(unit));
                            break;
                        }
                    }
                }

                UpdateUnitBox();
                AddItemFromTree(true);
            }

            private static string UnitKey(AgentUnitOption unit)
            {
                // 优先用 _ZGS_ 编号；缺编号时退回总概算序号，两者 ResolveAgentUnitIds 都认。
                string code = (unit.Code ?? "").Trim();
                return code.Length > 0 ? code : unit.UnitId.ToString(CultureInfo.InvariantCulture);
            }

            private string UnitDisplayOf(string key)
            {
                foreach (AgentUnitOption unit in unitOptions)
                {
                    if (String.Equals(UnitKey(unit), key, StringComparison.OrdinalIgnoreCase))
                    {
                        return unit.Display;
                    }
                }

                return key;
            }

            private string ItemDisplayOf(string itemNo)
            {
                foreach (AgentItemOption item in itemOptions)
                {
                    if (String.Equals(item.ItemNo, itemNo, StringComparison.OrdinalIgnoreCase))
                    {
                        return item.Display;
                    }
                }

                return itemNo;
            }

            private void UpdateUnitBox()
            {
                if (allUnitsBox.Checked)
                {
                    unitBox.Text = "所有单元";
                    return;
                }

                if (selectedUnitKeys.Count == 0)
                {
                    unitBox.Text = "（未识别当前单元，请点\"选单元\"）";
                    return;
                }

                List<string> parts = new List<string>();
                foreach (string key in selectedUnitKeys)
                {
                    parts.Add(UnitDisplayOf(key));
                }

                unitBox.Text = String.Join("、", parts.ToArray());
            }

            private void UpdateItemList()
            {
                itemListBox.Items.Clear();
                foreach (string itemNo in selectedItemNos)
                {
                    itemListBox.Items.Add(ItemDisplayOf(itemNo));
                }
            }

            private void PickUnits()
            {
                if (unitOptions.Count == 0)
                {
                    LoadScopeOptions();
                    if (unitOptions.Count == 0)
                    {
                        return;
                    }
                }

                List<AgentPickItem> items = new List<AgentPickItem>();
                foreach (AgentUnitOption unit in unitOptions)
                {
                    AgentPickItem item = new AgentPickItem();
                    item.Key = UnitKey(unit);
                    item.Display = unit.Display;
                    items.Add(item);
                }

                using (AgentPickerDialog dialog = new AgentPickerDialog("选择单元（可多选）", items, true, selectedUnitKeys))
                {
                    if (dialog.ShowDialog(this) != DialogResult.OK)
                    {
                        return;
                    }

                    selectedUnitKeys.Clear();
                    selectedUnitKeys.AddRange(dialog.SelectedKeys);
                }

                UpdateUnitBox();
                ClearTargetSelection("单元变了，目标定额已清空，请重新选。");
                UpdateSentence();
            }

            // 从左侧树取当前条目并累积进列表；同一个条目不重复加。
            private void AddItemFromTree(bool quiet)
            {
                try
                {
                    TreeView tree = GetField<TreeView>(mainForm, "Tv_tree");
                    TreeNode node = tree != null ? tree.SelectedNode : GetField<TreeNode>(mainForm, "CurrNode");
                    if (node == null)
                    {
                        if (!quiet)
                        {
                            SetStatus("当前没有选中树节点，请先在左侧树上点一个条目。", true);
                        }

                        return;
                    }

                    SqlConnection hostConn = GetProjectConnection(mainForm);
                    string itemNo = hostConn == null ? null : ResolveChapterNo(mainForm, hostConn, node);
                    if (String.IsNullOrEmpty(itemNo))
                    {
                        if (!quiet)
                        {
                            SetStatus("无法识别当前条目编号。", true);
                        }

                        return;
                    }

                    if (selectedItemNos.Contains(itemNo))
                    {
                        if (!quiet)
                        {
                            SetStatus("条目 " + ItemDisplayOf(itemNo) + " 已经在列表里了。", false);
                        }

                        return;
                    }

                    selectedItemNos.Add(itemNo);
                    UpdateItemList();
                    ClearTargetSelection(null);
                    UpdateSentence();
                    if (!quiet)
                    {
                        SetStatus("已加入条目 " + ItemDisplayOf(itemNo) + "，共 " +
                            selectedItemNos.Count.ToString(CultureInfo.InvariantCulture) + " 个。可以继续在树上点别的条目再加。", false);
                    }
                }
                catch (Exception ex)
                {
                    if (!quiet)
                    {
                        SetStatus("读取当前条目失败：" + ex.Message, true);
                    }

                    Log("Agent panel add item failed: " + ex);
                }
            }

            private void RemoveSelectedItems()
            {
                List<int> indexes = new List<int>();
                foreach (int index in itemListBox.SelectedIndices)
                {
                    indexes.Add(index);
                }

                if (indexes.Count == 0)
                {
                    SetStatus("请先在条目列表里选中要移除的行。", true);
                    return;
                }

                indexes.Sort();
                for (int i = indexes.Count - 1; i >= 0; i--)
                {
                    if (indexes[i] >= 0 && indexes[i] < selectedItemNos.Count)
                    {
                        selectedItemNos.RemoveAt(indexes[i]);
                    }
                }

                UpdateItemList();
                ClearTargetSelection(null);
                UpdateSentence();
            }

            private void ClearItems()
            {
                selectedItemNos.Clear();
                UpdateItemList();
                ClearTargetSelection(null);
                UpdateSentence();
                SetStatus("条目列表已清空。", false);
            }

            // ===== 目标定额 / 材料 / SF =====

            private void ClearTargetSelection(string message)
            {
                selectedTargetCodes.Clear();
                selectedTargetNames.Clear();
                UpdateTargetDisplay();
                if (!String.IsNullOrEmpty(message))
                {
                    SetStatus(message, false);
                }
            }

            private void UpdateTargetDisplay()
            {
                targetListBox.Items.Clear();
                List<string> selected = byNameRadio.Checked ? selectedTargetNames : selectedTargetCodes;
                foreach (string value in selected)
                {
                    targetListBox.Items.Add(value);
                }

                if (selected.Count == 0)
                {
                    targetSummaryLabel.ForeColor = Color.FromArgb(160, 80, 0);
                    targetSummaryLabel.Text = "未选目标 = 作用于范围内全部行";
                }
                else
                {
                    targetSummaryLabel.ForeColor = Color.FromArgb(0, 120, 0);
                    targetSummaryLabel.Text = "按" + (byNameRadio.Checked ? "名称" : "编号") + "查找这 " +
                        selected.Count.ToString(CultureInfo.InvariantCulture) + " 项";
                }
            }

            // 在主程序定额输入表里多选几行，把这些行的编号或精确名称取过来当查找依据。
            // 注意语义：取的是"找什么"，不是"只改这几行"——真正作用的是上面条目×单元范围内所有匹配的行。
            private void TakeTargetsFromHostGrid()
            {
                List<string> picked = new List<string>();
                HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
                int rowCount = 0;
                try
                {
                    DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
                    if (grid == null)
                    {
                        SetStatus("没有找到主程序的定额输入表格。", true);
                        return;
                    }

                    bool byName = byNameRadio.Checked;
                    foreach (DataGridViewRow row in GetSelectedQuotaRows(grid))
                    {
                        rowCount++;
                        string value = byName
                            ? GetRowValue(row, "工程或费用项目名称", "名称", "项目名称")
                            : GetRowValue(row, "定额编号DE", "定额编号");
                        value = (value ?? "").Trim();
                        if (value.Length > 0 && seen.Add(value))
                        {
                            picked.Add(value);
                        }
                    }
                }
                catch (Exception ex)
                {
                    SetStatus("读取定额表选中行失败：" + ex.Message, true);
                    Log("Agent panel take targets from grid failed: " + ex);
                    return;
                }

                if (rowCount == 0)
                {
                    SetStatus("主程序定额表里没有选中行。请先在定额输入表里按住 Ctrl 或 Shift 多选几行。", true);
                    return;
                }

                if (picked.Count == 0)
                {
                    SetStatus(byNameRadio.Checked
                        ? "选中的行读不到项目名称，换\"按编号\"再试。"
                        : "选中的行读不到定额编号，换\"按名称\"再试。", true);
                    return;
                }

                List<string> current = byNameRadio.Checked ? selectedTargetNames : selectedTargetCodes;
                int added = 0;
                foreach (string value in picked)
                {
                    if (!current.Contains(value))
                    {
                        current.Add(value);
                        added++;
                    }
                }

                UpdateTargetDisplay();
                UpdateSentence();
                SetStatus("从定额表 " + rowCount.ToString(CultureInfo.InvariantCulture) + " 个选中行取到 " +
                    added.ToString(CultureInfo.InvariantCulture) + " 个新" + (byNameRadio.Checked ? "名称" : "编号") +
                    "，共 " + current.Count.ToString(CultureInfo.InvariantCulture) + " 个。将在上面的条目×单元范围内查找。", false);
            }

            private void PickTargets()
            {
                if (selectedItemNos.Count == 0)
                {
                    SetStatus("请先加入至少一个条目，才能列出范围内的定额。", true);
                    return;
                }

                List<AgentQuotaCandidate> candidates;
                try
                {
                    SqlConnection conn = GetOpenProjectConnection(mainForm);
                    List<long> unitIds = ResolveScopeUnitIds(conn);
                    candidates = LoadAgentQuotaCandidates(conn, selectedItemNos, includeChildrenBox.Checked, unitIds);
                }
                catch (AgentPlanException ex)
                {
                    SetStatus(ex.Message, true);
                    return;
                }
                catch (Exception ex)
                {
                    SetStatus("读取范围内定额失败：" + ex.Message, true);
                    Log("Agent panel load candidates failed: " + ex);
                    return;
                }

                if (candidates.Count == 0)
                {
                    SetStatus("所选条目和单元范围内没有任何定额行。", true);
                    return;
                }

                bool byName = byNameRadio.Checked;
                List<AgentPickItem> items = new List<AgentPickItem>();
                HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
                foreach (AgentQuotaCandidate candidate in candidates)
                {
                    string key = byName ? candidate.Name : candidate.Code;
                    if (key.Length == 0 || !seen.Add(key))
                    {
                        continue;
                    }

                    AgentPickItem item = new AgentPickItem();
                    item.Key = key;
                    if (byName)
                    {
                        item.Display = candidate.Name + "　［" + candidate.Code + "］";
                    }
                    else
                    {
                        string kind = AgentQuotaCodeKind(candidate.Code);
                        item.Display = candidate.Code + "　［" + kind + "］　" + candidate.Name;
                    }

                    items.Add(item);
                }

                List<string> current = byName ? selectedTargetNames : selectedTargetCodes;
                string title = byName ? "按名称选择目标（可多选，精确名称）" : "按编号选择目标：定额 / 材料 / SF（可多选）";
                using (AgentPickerDialog dialog = new AgentPickerDialog(title, items, true, current))
                {
                    if (dialog.ShowDialog(this) != DialogResult.OK)
                    {
                        return;
                    }

                    current.Clear();
                    current.AddRange(dialog.SelectedKeys);
                }

                UpdateTargetDisplay();
                UpdateSentence();
            }

            // 目标候选查询要用真实的单元序号，这里做一次和执行层一致的解析。
            private List<long> ResolveScopeUnitIds(SqlConnection conn)
            {
                AgentCommand probe = new AgentCommand();
                probe.Units = BuildUnitTokens();
                return ResolveAgentUnitIds(conn, probe, CaptureAgentSelectionForPanel(), new List<string>());
            }

            private void PickCrossTargets()
            {
                if (itemOptions.Count == 0)
                {
                    LoadScopeOptions();
                }

                List<AgentPickItem> items = new List<AgentPickItem>();
                foreach (AgentItemOption option in itemOptions)
                {
                    AgentPickItem item = new AgentPickItem();
                    item.Key = option.ItemNo;
                    item.Display = option.Display;
                    items.Add(item);
                }

                bool multi = crossActionBox.Text == "复制到";
                string title = multi ? "选择目标条目（可多选）" : "选择目标条目（移动只能选一个）";
                using (AgentPickerDialog dialog = new AgentPickerDialog(title, items, multi, crossTargetItems))
                {
                    if (dialog.ShowDialog(this) != DialogResult.OK)
                    {
                        return;
                    }

                    crossTargetItems.Clear();
                    crossTargetItems.AddRange(dialog.SelectedKeys);
                }

                if (!multi && crossTargetItems.Count > 1)
                {
                    crossTargetItems.RemoveRange(1, crossTargetItems.Count - 1);
                }

                List<string> parts = new List<string>();
                foreach (string itemNo in crossTargetItems)
                {
                    parts.Add(ItemDisplayOf(itemNo));
                }

                crossTargetBox.Text = parts.Count == 0 ? "" : String.Join("、", parts.ToArray());
                UpdateSentence();
            }

            // 页签切换时，不适用的范围控件置灰，避免用户白配一通再报错。
            private void RefreshScopeAvailability()
            {
                string tab = tabs.SelectedTab == null ? "" : tabs.SelectedTab.Text;
                bool scopeUsed = tab != TabText;
                bool targetUsed = scopeUsed && !(tab == TabRows && rowActionBox.Text == "新增定额");

                unitBox.Enabled = scopeUsed;
                allUnitsBox.Enabled = scopeUsed;
                itemListBox.Enabled = scopeUsed;
                includeChildrenBox.Enabled = scopeUsed;
                byCodeRadio.Enabled = targetUsed;
                byNameRadio.Enabled = targetUsed;
                pickTargetButton.Enabled = targetUsed;
                takeFromGridButton.Enabled = targetUsed;
                targetListBox.Enabled = targetUsed;
                previewButton.Enabled = tab != TabText;
            }

            // ===== 页签内联动 =====

            private void RefreshValueActions()
            {
                suppressSentence = true;
                string field = fieldBox.Text;
                valueActionBox.Items.Clear();
                if (field == "工程数量")
                {
                    valueActionBox.Items.AddRange(new object[] { "乘以", "除以", "设为", "清空", "删除片段" });
                }
                else if (field == "单价")
                {
                    valueActionBox.Items.AddRange(new object[] { "乘以", "除以", "设为" });
                }
                else if (field == "定额编号")
                {
                    valueActionBox.Items.AddRange(new object[] { "乘以", "除以", "删除片段" });
                }
                else
                {
                    valueActionBox.Items.AddRange(new object[] { "设为", "追加", "删除内容" });
                }

                valueActionBox.SelectedIndex = 0;
                suppressSentence = false;
                RefreshValueInput();
            }

            private void RefreshValueInput()
            {
                string field = fieldBox.Text;
                string action = valueActionBox.Text;
                bool needsValue = action != "清空";
                valueBox.Enabled = needsValue;
                if (!needsValue)
                {
                    valueBox.Text = "";
                }

                if (action == "乘以" || action == "除以")
                {
                    valueLabel.Text = "系数";
                }
                else if (action == "删除片段" || action == "删除内容")
                {
                    valueLabel.Text = "要删除的";
                }
                else if (field == "定额调整")
                {
                    valueLabel.Text = "调整内容";
                }
                else
                {
                    valueLabel.Text = "数值";
                }

                if (field == "定额编号")
                {
                    valueHintLabel.Text = "在定额编号后面追加乘除系数（如 LY-21 变成 LY-21*9），适合软件原生缩放定额；改完通常要在主程序里手工重算。";
                }
                else if (field == "定额调整")
                {
                    valueHintLabel.Text = "调整内容按整串写入或从原串删除，例如 /XG1、/1294861,,1。\"追加\"会拼在原有调整后面。";
                }
                else if (field == "单价")
                {
                    valueHintLabel.Text = "SF / SH / SQ / ZLF / LF / TLF 这类手填单价的行改了才稳定保留；普通定额的单价是计算值，会被主程序重算覆盖。";
                }
                else if (action == "删除片段")
                {
                    valueHintLabel.Text = "从工程数量输入里去掉指定子串，例如原来是 100*0.85，删除 *0.85 后剩 100。字段里没有该片段的行会跳过。";
                }
                else if (action == "清空")
                {
                    valueHintLabel.Text = "把工程数量输入和计算数量都置空。要直接改成某个数字请用\"设为\"。";
                }
                else
                {
                    valueHintLabel.Text = "\"设为\"直接写死数值；\"乘以/除以\"在原数量后面追加系数。\r\n" +
                        "单个条目内的乘系数/删系数，主程序右键\"乘系数\"也能做；这里的价值是一次覆盖多个条目、多个单元。";
                }

                UpdateSentence();
            }

            private void RefreshRowAction()
            {
                string action = rowActionBox.Text;
                bool isReplace = action == "替换目标定额";
                bool isInsert = action == "新增定额";

                quotaGridLabel.Visible = isReplace || isInsert;
                quotaGrid.Visible = isReplace || isInsert;

                if (isReplace)
                {
                    quotaGridLabel.Text = "替换成（数量留空＝不改数量）　—　被替换的就是上面选中的目标";
                    rowHintLabel.Text = "1对1、1对多、多对1都支持。";
                }
                else if (isInsert)
                {
                    quotaGridLabel.Text = "要新增的定额（数量留空＝不填）";
                    rowHintLabel.Text = "在上面列出的每个条目下各新增一份，追加到条目末尾。";
                }
                else
                {
                    rowHintLabel.Text = "删除上面选中的目标行；不选目标则删除范围内全部定额行。";
                }

                RefreshScopeAvailability();
                UpdateSentence();
            }

            private void RefreshCrossAction()
            {
                bool move = crossActionBox.Text == "移动到";
                crossHintLabel.Text = move
                    ? "移动不保留来源副本，目标只能有一个条目。来源条目只能是上面列表里的第一个（且只能有一个）。"
                    : "复制会在目标条目下新增同样的定额，来源保持不变。目标条目可以选多个。";
                if (move && crossTargetItems.Count > 1)
                {
                    crossTargetItems.RemoveRange(1, crossTargetItems.Count - 1);
                    crossTargetBox.Text = ItemDisplayOf(crossTargetItems[0]);
                }

                UpdateSentence();
            }

            // ===== 组装命令 =====

            private List<string> BuildUnitTokens()
            {
                if (allUnitsBox.Checked)
                {
                    return new List<string> { "所有" };
                }

                return new List<string>(selectedUnitKeys);
            }

            private List<string> BuildItemTokens()
            {
                if (selectedItemNos.Count == 0)
                {
                    throw new AgentPlanException("请先在左侧树点中条目，再点\"加入树上选中\"。可以反复加多个条目。");
                }

                return new List<string>(selectedItemNos);
            }

            // 目标：按编号写进 QuotaFilter，按名称写进 QuotaName（顿号分隔的精确名称串）。
            private void ApplyTarget(AgentCommand command)
            {
                if (byNameRadio.Checked)
                {
                    if (selectedTargetNames.Count > 0)
                    {
                        command.QuotaName = String.Join("、", selectedTargetNames.ToArray());
                    }
                }
                else if (selectedTargetCodes.Count > 0)
                {
                    command.QuotaFilter = new List<string>(selectedTargetCodes);
                }
            }

            private void ApplyScope(AgentCommand command)
            {
                command.IncludeChildren = includeChildrenBox.Checked;
                command.Units = BuildUnitTokens();
                command.Items = BuildItemTokens();
                ApplyTarget(command);
            }

            private static string RequireText(string value, string what)
            {
                string text = (value ?? "").Trim();
                if (text.Length == 0)
                {
                    throw new AgentPlanException("请填写" + what + "。");
                }

                return text;
            }

            private List<AgentQuotaInput> ReadQuotaGrid()
            {
                List<AgentQuotaInput> quotas = new List<AgentQuotaInput>();
                foreach (DataGridViewRow row in quotaGrid.Rows)
                {
                    if (row.IsNewRow)
                    {
                        continue;
                    }

                    string code = Convert.ToString(row.Cells["Code"].Value ?? "").Trim();
                    string quantity = Convert.ToString(row.Cells["Quantity"].Value ?? "").Trim();
                    if (code.Length == 0 && quantity.Length == 0)
                    {
                        continue;
                    }

                    if (code.Length == 0)
                    {
                        throw new AgentPlanException("有一行只填了数量没填定额编号，请补齐或删掉该行。");
                    }

                    AgentQuotaInput quota = new AgentQuotaInput();
                    quota.Code = code;
                    quota.Quantity = quantity;
                    quotas.Add(quota);
                }

                return quotas;
            }

            private AgentCommand BuildValueCommand()
            {
                string field = fieldBox.Text;
                string action = valueActionBox.Text;
                string value = (valueBox.Text ?? "").Trim();

                AgentCommand command = new AgentCommand();
                ApplyScope(command);

                if (field == "工程数量")
                {
                    if (action == "设为")
                    {
                        command.Type = "set_quantity";
                        command.Value = RequireText(value, "数值");
                    }
                    else if (action == "清空")
                    {
                        command.Type = "clear_quantity";
                    }
                    else if (action == "删除片段")
                    {
                        command.Type = "remove_text";
                        command.Target = "quantity";
                        command.RemoveText = RequireText(value, "要删除的片段");
                    }
                    else
                    {
                        command.Type = "multiply_quantity";
                        command.Target = "quantity";
                        command.Operator = action == "除以" ? "/" : "*";
                        command.Factor = RequireText(value, "系数");
                    }
                }
                else if (field == "单价")
                {
                    if (action == "设为")
                    {
                        command.Type = "set_unit_price";
                        command.Value = RequireText(value, "数值");
                    }
                    else
                    {
                        command.Type = "multiply_quantity";
                        command.Target = "unit_price";
                        command.Operator = action == "除以" ? "/" : "*";
                        command.Factor = RequireText(value, "系数");
                    }
                }
                else if (field == "定额编号")
                {
                    if (action == "删除片段")
                    {
                        command.Type = "remove_text";
                        command.Target = "quota_code";
                        command.RemoveText = RequireText(value, "要删除的片段");
                    }
                    else
                    {
                        command.Type = "multiply_quantity";
                        command.Target = "quota_code";
                        command.Operator = action == "除以" ? "/" : "*";
                        command.Factor = RequireText(value, "系数");
                    }
                }
                else
                {
                    if (action == "删除内容")
                    {
                        command.Type = "remove_text";
                        command.Target = "adjustment";
                        command.RemoveText = RequireText(value, "要删除的调整内容");
                    }
                    else
                    {
                        command.Type = "set_adjustment";
                        command.Mode = action == "追加" ? "append" : "set";
                        command.Value = RequireText(value, "调整内容");
                    }
                }

                return command;
            }

            private AgentCommand BuildRowCommand()
            {
                string action = rowActionBox.Text;
                AgentCommand command = new AgentCommand();

                if (action == "新增定额")
                {
                    command.Type = "insert_quotas";
                    command.IncludeChildren = includeChildrenBox.Checked;
                    command.Units = BuildUnitTokens();
                    command.Items = BuildItemTokens();
                    command.Quotas = ReadQuotaGrid();
                    if (command.Quotas.Count == 0)
                    {
                        throw new AgentPlanException("请填至少一条要新增的定额。");
                    }

                    return command;
                }

                if (action == "删除目标定额")
                {
                    command.Type = "delete_quotas";
                    ApplyScope(command);
                    return command;
                }

                command.Type = "replace_quotas";
                ApplyScope(command);
                if (byNameRadio.Checked)
                {
                    if (selectedTargetNames.Count == 0)
                    {
                        throw new AgentPlanException("请先在上面按名称选出要被替换的目标。");
                    }
                }
                else
                {
                    if (selectedTargetCodes.Count == 0)
                    {
                        throw new AgentPlanException("请先在上面按编号选出要被替换的目标。");
                    }

                    command.FromCodes = new List<string>(selectedTargetCodes);
                    command.QuotaFilter = new List<string>();
                }

                command.ToQuotas = ReadQuotaGrid();
                if (command.ToQuotas.Count == 0)
                {
                    throw new AgentPlanException("请填至少一条替换成的定额。");
                }

                HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (AgentQuotaInput quota in command.ToQuotas)
                {
                    if (!seen.Add(quota.Code))
                    {
                        throw new AgentPlanException("替换成的定额编号重复了：" + quota.Code);
                    }
                }

                if (command.ToQuotas.Count > 1 && allUnitsBox.Checked)
                {
                    throw new AgentPlanException("一条定额拆成多条时不能选\"所有单元\"，请限定到一个单元。");
                }

                return command;
            }

            private AgentCommand BuildCrossCommand()
            {
                if (crossTargetItems.Count == 0)
                {
                    throw new AgentPlanException("请先选择目标条目。");
                }

                bool move = crossActionBox.Text == "移动到";
                if (move && crossTargetItems.Count != 1)
                {
                    throw new AgentPlanException("移动定额只能有一个目标条目。");
                }

                List<string> sources = BuildItemTokens();
                if (sources.Count > 1)
                {
                    throw new AgentPlanException("复制/移动的来源只能是一个条目，请把上面的条目列表减到一个。");
                }

                AgentCommand command = new AgentCommand();
                command.Type = move ? "move_quotas" : "copy_quotas";
                command.IncludeChildren = includeChildrenBox.Checked;
                command.Units = BuildUnitTokens();
                command.SourceItem = sources[0];
                command.TargetItems = new List<string>(crossTargetItems);
                ApplyTarget(command);

                foreach (string target in crossTargetItems)
                {
                    if (String.Equals(target, command.SourceItem, StringComparison.OrdinalIgnoreCase))
                    {
                        throw new AgentPlanException("来源条目和目标条目不能相同。");
                    }
                }

                return command;
            }

            private List<AgentCommand> BuildCommands()
            {
                string tab = tabs.SelectedTab == null ? "" : tabs.SelectedTab.Text;
                List<AgentCommand> commands = new List<AgentCommand>();
                if (tab == TabValue)
                {
                    commands.Add(BuildValueCommand());
                }
                else if (tab == TabRows)
                {
                    commands.Add(BuildRowCommand());
                }
                else if (tab == TabCross)
                {
                    commands.Add(BuildCrossCommand());
                }

                return commands;
            }

            private void UpdateSentence()
            {
                if (suppressSentence)
                {
                    return;
                }

                string tab = tabs.SelectedTab == null ? "" : tabs.SelectedTab.Text;
                if (tab == TabText)
                {
                    sentenceLabel.ForeColor = Color.Gray;
                    sentenceLabel.Text = "文本/自然语言通道：写好后点\"交给 AI\"。";
                    return;
                }

                try
                {
                    List<AgentCommand> commands = BuildCommands();
                    if (commands.Count == 0)
                    {
                        sentenceLabel.Text = "";
                        return;
                    }

                    List<string> parts = new List<string>();
                    foreach (AgentCommand command in commands)
                    {
                        parts.Add(command.Describe());
                    }

                    sentenceLabel.ForeColor = Color.FromArgb(20, 60, 160);
                    sentenceLabel.Text = "将要执行：" + String.Join("；", parts.ToArray());
                }
                catch (Exception ex)
                {
                    sentenceLabel.ForeColor = Color.Gray;
                    sentenceLabel.Text = "还差一步：" + ex.Message;
                }
            }

            // ===== 提交与预览 =====

            private AgentSelectionSnapshot CaptureAgentSelectionForPanel()
            {
                // 本面板不使用"主程序选中行"作为范围，快照里的行选中一律丢弃，
                // 避免宿主表格里顺带选中的当前行影响批量范围。
                s_agentInvokeFromTree = true;
                try
                {
                    return CaptureAgentSelection(mainForm);
                }
                finally
                {
                    s_agentInvokeFromTree = false;
                }
            }

            private void SubmitPanelCommand()
            {
                if (busy)
                {
                    SetStatus("上一条还在处理中，请稍候。", true);
                    return;
                }

                if (pendingPlan != null)
                {
                    CancelPlan("已放弃之前未确认的预览。");
                }

                List<AgentCommand> commands;
                try
                {
                    commands = BuildCommands();
                }
                catch (AgentPlanException ex)
                {
                    SetStatus(ex.Message, true);
                    return;
                }
                catch (Exception ex)
                {
                    SetStatus("配置有误：" + ex.Message, true);
                    return;
                }

                if (commands.Count == 0)
                {
                    SetStatus("当前页签没有可执行的操作。", true);
                    return;
                }

                RunPipeline(null, commands, null);
            }

            private void SubmitText()
            {
                if (busy)
                {
                    SetStatus("上一条还在处理中，请稍候。", true);
                    return;
                }

                string text = (textBox.Text ?? "").Trim();
                if (text.Length == 0)
                {
                    return;
                }

                if (pendingPlan != null)
                {
                    CancelPlan("已放弃之前未确认的预览。");
                }

                string normalized = NormalizeAgentInput(text).TrimStart('/');
                if (normalized == "帮助" || String.Equals(normalized, "help", StringComparison.OrdinalIgnoreCase) || normalized == "?")
                {
                    ShowHelpDialog();
                    return;
                }

                if (normalized == "撤销" || normalized == "撤回" || normalized == "撤掉上一步" || normalized == "撤掉")
                {
                    PreviewUndo();
                    return;
                }

                if (normalized == "重做" || normalized == "恢复")
                {
                    PreviewRedo();
                    return;
                }

                if (normalized.StartsWith("探查", StringComparison.Ordinal))
                {
                    // 探查会输出很多行，状态栏放不下，收集完整体显示。
                    StringBuilder diagnostics = new StringBuilder();
                    RunAgentDiagnostics(mainForm, normalized.Substring(2), delegate(string line)
                    {
                        diagnostics.AppendLine(line);
                    });
                    ShowTextDialog("探查结果", diagnostics.ToString());
                    SetStatus("探查完成。", false);
                    return;
                }

                List<AgentCommand> commands = null;
                AgentParseResult fallback;
                if (TryParseAgentChain(text, out fallback))
                {
                    if (!String.IsNullOrEmpty(fallback.Error))
                    {
                        SetStatus(fallback.Error, true);
                        return;
                    }

                    commands = fallback.Commands;
                }

                DeepSeekExcelMatchSettings settings = null;
                if (commands == null)
                {
                    settings = LoadDeepSeekExcelMatchSettings();
                    if (!settings.IsAvailable)
                    {
                        SetStatus("没有可用的 AI 配置（RecoQuotaData/deepseek-settings.json 需启用并填写 api_key）。" +
                            "可以改用上面的按钮页签，或输入确定性语法（点\"指令帮助\"看格式）。", true);
                        return;
                    }
                }

                RunPipeline(text, commands, settings);
            }

            private void RunPipeline(string text, List<AgentCommand> preParsed, DeepSeekExcelMatchSettings settings)
            {
                SqlConnection expectedConnection;
                string expectedConnectionIdentity;
                try
                {
                    expectedConnection = GetOpenProjectConnection(mainForm);
                    expectedConnectionIdentity = GetProjectConnectionIdentity(expectedConnection);
                }
                catch (AgentPlanException ex)
                {
                    SetStatus(ex.Message, true);
                    return;
                }

                AgentSelectionSnapshot snapshot = CaptureAgentSelectionForPanel();
                busy = true;
                previewButton.Enabled = false;
                textSendButton.Enabled = false;
                SetStatus(preParsed != null ? "正在生成预览…" : "AI 解析中…", false);
                Stopwatch watch = Stopwatch.StartNew();
                Thread worker = new Thread(delegate()
                {
                    List<AgentCommand> commands = preParsed;
                    AgentParseResult llmResult = null;
                    AgentPlan plan = null;
                    string error = null;
                    try
                    {
                        if (commands == null)
                        {
                            AgentContext context = WithOpenProjectConnectionOnUi(mainForm, expectedConnection,
                                expectedConnectionIdentity, delegate(SqlConnection conn)
                            {
                                return CollectAgentContext(conn, snapshot, text);
                            });
                            llmResult = RequestAgentParse(settings, context, text);
                            if (!String.IsNullOrEmpty(llmResult.Error))
                            {
                                error = llmResult.Error;
                            }
                            else if (llmResult.Commands.Count > 0)
                            {
                                commands = llmResult.Commands;
                            }
                        }

                        if (error == null && commands != null && commands.Count > 0)
                        {
                            plan = WithOpenProjectConnectionOnUi(mainForm, expectedConnection,
                                expectedConnectionIdentity, delegate(SqlConnection conn)
                            {
                                return BuildAgentPlan(conn, snapshot, commands);
                            });
                        }
                    }
                    catch (AgentPlanException ex)
                    {
                        error = ex.Message;
                    }
                    catch (Exception ex)
                    {
                        error = "处理失败：" + ex.Message;
                        Log("Agent panel pipeline failed: " + ex);
                    }

                    watch.Stop();
                    AgentParseResult resultForUi = llmResult;
                    AgentPlan planForUi = plan;
                    string errorForUi = error;
                    double seconds = watch.Elapsed.TotalSeconds;
                    try
                    {
                        BeginInvoke((MethodInvoker)delegate
                        {
                            OnPipelineDone(resultForUi, planForUi, errorForUi, seconds);
                        });
                    }
                    catch
                    {
                        busy = false;
                    }
                });
                worker.IsBackground = true;
                worker.Start();
            }

            private void OnPipelineDone(AgentParseResult llmResult, AgentPlan plan, string error, double seconds)
            {
                busy = false;
                textSendButton.Enabled = true;
                RefreshScopeAvailability();
                string elapsed = "（耗时 " + seconds.ToString("0.0", CultureInfo.InvariantCulture) + " 秒）";
                if (!String.IsNullOrEmpty(error))
                {
                    SetStatus(error + " " + elapsed, true);
                    return;
                }

                if (llmResult != null && llmResult.Commands.Count == 0)
                {
                    SetStatus((String.IsNullOrEmpty(llmResult.Clarification)
                        ? "没有解析出可执行的命令，可以补充条目编号后重试。"
                        : llmResult.Clarification) + " " + elapsed, true);
                    return;
                }

                if (plan == null)
                {
                    SetStatus("没有生成执行计划。" + elapsed, true);
                    return;
                }

                if (plan.PreviewRows.Count == 0)
                {
                    SetStatus("没有匹配到任何数据行，未生成执行计划。" + elapsed, true);
                    return;
                }

                if (llmResult != null)
                {
                    StringBuilder description = new StringBuilder("AI 理解为：");
                    for (int i = 0; i < llmResult.Commands.Count; i++)
                    {
                        if (i > 0)
                        {
                            description.Append("；");
                        }

                        description.Append(llmResult.Commands[i].Describe());
                    }

                    sentenceLabel.ForeColor = Color.FromArgb(20, 60, 160);
                    sentenceLabel.Text = description.ToString();
                }

                ShowPlanPreview(plan);
                StringBuilder note = new StringBuilder("请核对下方预览，点\"确认执行\"生效。");
                foreach (string warning in plan.Warnings)
                {
                    note.Append("  注意：").Append(warning);
                }

                SetStatus(note.ToString() + " " + elapsed, false);
            }

            private void ShowPlanPreview(AgentPlan plan)
            {
                pendingPlan = plan;
                previewGrid.Rows.Clear();
                foreach (AgentPlanRow row in plan.PreviewRows.Take(2000))
                {
                    previewGrid.Rows.Add(
                        row.Action,
                        AgentUnitDisplay(plan.UnitCodes, row.UnitId),
                        !String.IsNullOrEmpty(row.ItemName) ? row.ItemName : (row.ItemNo ?? ""),
                        row.QuotaCode ?? "",
                        row.OldValue ?? "",
                        row.NewValue ?? "");
                }

                string extra = plan.PreviewRows.Count > 2000 ? "（预览表只显示前2000行）" : "";
                summaryLabel.Text = "第三步  确认执行：" + plan.Summary + extra;
                previewPanel.Visible = true;
            }

            private void CancelPlan(string message)
            {
                pendingPlan = null;
                previewPanel.Visible = false;
                previewGrid.Rows.Clear();
                if (!String.IsNullOrEmpty(message))
                {
                    SetStatus(message, false);
                }
            }

            private void ConfirmPlan()
            {
                AgentPlan plan = pendingPlan;
                if (plan == null)
                {
                    previewPanel.Visible = false;
                    return;
                }

                if (plan.PreviewRows.Count > 200 && !ConfirmLargePlan(plan.PreviewRows.Count))
                {
                    SetStatus("已取消大批量执行。", false);
                    return;
                }

                pendingPlan = null;
                previewPanel.Visible = false;
                Enabled = false;
                try
                {
                    string message = ExecuteAgentPlan(mainForm, plan, delegate(string line) { });
                    SetStatus(message.Replace("\r\n", "  ").Replace("\n", "  "), false);
                }
                catch (AgentPlanException ex)
                {
                    SetStatus(ex.Message, true);
                }
                catch (Exception ex)
                {
                    SetStatus("执行失败：" + ex.Message, true);
                    Log("Agent panel execute failed: " + ex);
                }
                finally
                {
                    Enabled = true;
                    previewGrid.Rows.Clear();
                    RefreshUndoRedoButtons();
                    UpdateSentence();
                }
            }

            private bool ConfirmLargePlan(int rowCount)
            {
                using (Form dialog = new Form())
                using (Label label = new Label())
                using (TextBox box = new TextBox())
                using (Button ok = new Button())
                using (Button cancel = new Button())
                {
                    dialog.Text = "大批量操作确认";
                    dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                    dialog.StartPosition = FormStartPosition.CenterParent;
                    dialog.MinimizeBox = false;
                    dialog.MaximizeBox = false;
                    dialog.ClientSize = new Size(360, 130);

                    label.Text = "本次将影响 " + rowCount.ToString(CultureInfo.InvariantCulture) + " 行数据。\r\n请输入\"确认\"两字后继续：";
                    label.SetBounds(12, 12, 330, 40);
                    box.SetBounds(12, 58, 330, 24);
                    ok.Text = "继续";
                    ok.SetBounds(180, 92, 75, 28);
                    ok.DialogResult = DialogResult.OK;
                    cancel.Text = "取消";
                    cancel.SetBounds(265, 92, 75, 28);
                    cancel.DialogResult = DialogResult.Cancel;

                    dialog.Controls.Add(label);
                    dialog.Controls.Add(box);
                    dialog.Controls.Add(ok);
                    dialog.Controls.Add(cancel);
                    dialog.AcceptButton = ok;
                    dialog.CancelButton = cancel;

                    while (dialog.ShowDialog(this) == DialogResult.OK)
                    {
                        if ((box.Text ?? "").Trim() == "确认")
                        {
                            return true;
                        }

                        box.SelectAll();
                        box.Focus();
                    }

                    return false;
                }
            }

            // ===== 撤销 / 重做 =====

            private void RefreshUndoRedoButtons()
            {
                try
                {
                    List<AgentUndoRecord> undoStack = GetAgentUndoStack(mainForm);
                    undoButton.Enabled = undoStack.Count > 0;
                    if (undoStack.Count > 0)
                    {
                        AgentUndoRecord last = undoStack[undoStack.Count - 1];
                        SetToolTip(undoButton, "撤销：" + last.Summary + "（" + last.Time.ToString("HH:mm:ss") + "）");
                    }

                    List<AgentUndoRecord> redoStack = GetAgentRedoStack(mainForm);
                    redoButton.Enabled = redoStack.Count > 0;
                    if (redoStack.Count > 0)
                    {
                        SetToolTip(redoButton, "重做：" + redoStack[redoStack.Count - 1].Summary);
                    }
                }
                catch (Exception ex)
                {
                    Log("Agent panel refresh undo buttons failed: " + ex.Message);
                }
            }

            private void SetToolTip(Control control, string text)
            {
                try
                {
                    toolTip.SetToolTip(control, text);
                }
                catch
                {
                }
            }

            private void PreviewUndo()
            {
                try
                {
                    ShowPlanPreview(BuildAgentUndoPlan(mainForm));
                    SetStatus("这是撤销预览，点\"确认执行\"才会真正回滚。", false);
                }
                catch (AgentPlanException ex)
                {
                    SetStatus(ex.Message, true);
                }
                catch (Exception ex)
                {
                    SetStatus("撤销准备失败：" + ex.Message, true);
                    Log("Agent panel undo preview failed: " + ex);
                }
            }

            private void PreviewRedo()
            {
                try
                {
                    ShowPlanPreview(BuildAgentRedoPlan(mainForm));
                    SetStatus("这是重做预览，点\"确认执行\"才会重新应用。", false);
                }
                catch (AgentPlanException ex)
                {
                    SetStatus(ex.Message, true);
                }
                catch (Exception ex)
                {
                    SetStatus("重做准备失败：" + ex.Message, true);
                    Log("Agent panel redo preview failed: " + ex);
                }
            }

            // ===== 状态栏与帮助 =====

            private void SetStatus(string text, bool isError)
            {
                statusLabel.ForeColor = isError ? Color.FromArgb(190, 30, 30) : Color.FromArgb(50, 50, 50);
                statusLabel.Text = text ?? "";
            }

            private void ShowTextDialog(string title, string text)
            {
                using (Form dialog = new Form())
                {
                    dialog.Text = title;
                    dialog.StartPosition = FormStartPosition.CenterParent;
                    dialog.ClientSize = new Size(720, 480);
                    dialog.ShowInTaskbar = false;
                    dialog.MinimizeBox = false;

                    TextBox box = new TextBox();
                    box.Dock = DockStyle.Fill;
                    box.Multiline = true;
                    box.ReadOnly = true;
                    box.ScrollBars = ScrollBars.Both;
                    box.WordWrap = false;
                    box.Font = new Font(FontFamily.GenericMonospace, 9f);
                    box.Text = text ?? "";

                    dialog.Controls.Add(box);
                    dialog.ShowDialog(this);
                }
            }

            private void ShowHelpDialog()
            {
                using (Form dialog = new Form())
                {
                    dialog.Text = "文本指令帮助（\"说一句话\"页签用）";
                    dialog.StartPosition = FormStartPosition.CenterParent;
                    dialog.ClientSize = new Size(820, 560);
                    dialog.ShowInTaskbar = false;
                    dialog.MinimizeBox = false;

                    DataGridView grid = new DataGridView();
                    grid.Dock = DockStyle.Fill;
                    grid.ReadOnly = true;
                    grid.AllowUserToAddRows = false;
                    grid.AllowUserToDeleteRows = false;
                    grid.RowHeadersVisible = false;
                    grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    grid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                    grid.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                    grid.DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft;
                    grid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
                    grid.Columns.Add("Scene", "场景");
                    grid.Columns.Add("Format", "写法");
                    grid.Columns.Add("Example", "实例");
                    grid.Columns.Add("Note", "说明");
                    grid.Columns["Scene"].FillWeight = 13;
                    grid.Columns["Format"].FillWeight = 25;
                    grid.Columns["Example"].FillWeight = 31;
                    grid.Columns["Note"].FillWeight = 31;
                    PopulateAgentHelpGrid(grid);

                    dialog.Controls.Add(grid);
                    dialog.ShowDialog(this);
                }
            }

            private static void PopulateAgentHelpGrid(DataGridView grid)
            {
                grid.Rows.Clear();
                AddHelpRow(grid, "分工",
                    "单个条目内改数量/编号/单价/调整 → 主程序右键\"乘系数\"；跨条目跨单元批量 → 本窗口上面的按钮页签。",
                    "本页只说明\"说一句话\"页签的文本写法。",
                    "按钮页签拼不出来的场景（新建单元、运输方案、材料价方案）才需要文本或 AI。");
                AddHelpRow(grid, "自然语言（需AI）",
                    "直接用一句话描述要改什么，助手会先生成预览，确认后才执行。",
                    "把0101-01条目的定额数量乘0.85\r\n把南江路泵房单元0308-01的运输方案设为3\r\n照着_ZGS_02再建一个单元，叫测算二版",
                    "需要已配置 RecoQuotaData/deepseek-settings.json。AI 无法唯一判断条目或单元时，会要求补充。");
                AddHelpRow(grid, "作用范围",
                    "不写条目编号=当前选中的条目或定额；不写定额过滤=该条目下全部定额；单元=xxx 可限定单元。",
                    "工程数量 *0.85\r\n工程数量 0101-01 *0.85\r\n工程数量 0101-01、0102-01 *0.85\r\n删除数量 0308、0309 *0 单元=所有",
                    "多个条目或单元用顿号\"、\"隔开（英文逗号也兼容）；分号\"；\"只用于分隔多条完整命令。");
                AddHelpRow(grid, "工程数量",
                    "工程数量 [条目编号] [定额编号或精确定额名称] 数字、*系数或/系数",
                    "工程数量 LY-21 100\r\n工程数量 0101-01 LY-21 *0.85\r\n工程数量 0305 0",
                    "末尾普通数字=直接设值，带*或/=乘除原数量。定额名称完全匹配。");
                AddHelpRow(grid, "单价 / 定额编号",
                    "单价 [条目编号] [定额编号] 数字或*系数；定额编号 [条目编号] [定额编号] *系数",
                    "单价 SH 3500\r\n单价 0101-01 SH *1.05\r\n定额编号 0101-01 LY-21 *9",
                    "SF/SH/SQ/ZLF/LF/TLF 是手填单价的补充定额；普通定额单价是计算值，会被主程序重算覆盖。");
                AddHelpRow(grid, "定额调整",
                    "定额调整 [条目编号] [定额编号] 调整内容；删除时在调整内容前写 删除。",
                    "定额调整 0101-01 LY-21 /XG1\r\n定额调整 0101-01 LY-21 删除 /XG1",
                    "调整内容按整串写入或从原串删除。");
                AddHelpRow(grid, "增删定额",
                    "输入定额 [条目编号] 编号=数量；删除定额 [条目编号] [定额编号]；替换定额 [条目编号] 原定额 新定额",
                    "输入定额 0101-01 LY-21=100\r\n删除定额 0101-01 LY-21\r\n替换定额 0101-01 LY-21 QY-100",
                    "文本通道的\"替换定额\"只支持一对一；一对多、多对一请用\"加减定额\"页签。");
                AddHelpRow(grid, "复制 / 移动",
                    "复制定额 来源条目 到 目标条目；移动定额 来源条目 [定额编号] 到 目标条目",
                    "复制定额 0101-01 到 0102-01\r\n移动定额 0305 LY-21、LY-22 到 0306 单元=04",
                    "移动不保留来源副本，目标只能有一个条目。");
                AddHelpRow(grid, "运输 / 材料价方案",
                    "设运输方案 [条目编号] 方案序号 [运输参数] 单元=单元名；改材料价 [材料|机械|设备|工费] 方案名称 单元=单元名",
                    "设运输方案 0101-01 4 PH0 单元=南江路泵房\r\n改材料价 材料 部颁25年4季度 单元=南江路泵房",
                    "按钮界面不提供这两项，只能从这里走。必须明确单元，改完需要在软件里手工触发重算。");
                AddHelpRow(grid, "新建单元",
                    "新建单元 新名称 从 源单元名称；源也可写 _ZGS_编号 或总概算序号。",
                    "新建单元 测算二版 从 _ZGS_02",
                    "按钮界面不提供，只能从这里走。会复制源单元的总概算条目、单项概算信息、定额输入等数据。");
                AddHelpRow(grid, "撤销 / 重做 / 探查",
                    "撤销；重做；探查 关键词",
                    "撤销\r\n重做\r\n探查 当前选择",
                    "撤销只记录本次软件运行期间由本工具执行的操作；删除、插入、新建单元不支持重做。");
            }

            private static void AddHelpRow(DataGridView grid, string scene, string format, string example, string note)
            {
                grid.Rows.Add(scene, format, example, note);
            }
        }
    }
}
