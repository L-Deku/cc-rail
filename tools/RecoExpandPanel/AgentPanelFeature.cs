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

        // 界面配色，集中放一处好统一调。
        private static readonly Color AgentPanelSectionBack = Color.FromArgb(246, 248, 251);
        private static readonly Color AgentPanelTitleFore = Color.FromArgb(23, 78, 166);
        private static readonly Color AgentPanelHintFore = Color.FromArgb(122, 128, 138);
        private static readonly Color AgentPanelOkFore = Color.FromArgb(22, 121, 60);
        private static readonly Color AgentPanelWarnFore = Color.FromArgb(176, 96, 12);
        private static readonly Color AgentPanelErrorFore = Color.FromArgb(190, 40, 40);
        private static readonly Color AgentPanelLine = Color.FromArgb(219, 225, 234);

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

        // 用户从左侧树点选"添加条目"时记下的节点，按条目编号索引。
        // 插入定额要先在树上定位条目才能粘贴，而宿主树节点不一定把条目序号/编号放进 Name 或 Tag，
        // 反查经常失败；用户既然是从树上点进来的，直接把那个节点留着复用最稳。
        private static readonly Dictionary<string, TreeNode> AgentKnownItemNodes =
            new Dictionary<string, TreeNode>(StringComparer.OrdinalIgnoreCase);

        private static void RememberAgentItemNode(string itemNo, TreeNode node)
        {
            if (String.IsNullOrEmpty(itemNo) || node == null)
            {
                return;
            }

            AgentKnownItemNodes[itemNo] = node;
        }

        // 取回记住的节点；节点可能因为树重建而失效，这里顺带校验它还挂在树上。
        private static TreeNode TryGetKnownAgentItemNode(TreeView tree, string itemNo)
        {
            TreeNode node;
            if (tree == null || String.IsNullOrEmpty(itemNo) ||
                !AgentKnownItemNodes.TryGetValue(itemNo, out node) || node == null)
            {
                return null;
            }

            try
            {
                if (node.TreeView == null || !Object.ReferenceEquals(node.TreeView, tree))
                {
                    AgentKnownItemNodes.Remove(itemNo);
                    return null;
                }

                return node;
            }
            catch (Exception)
            {
                AgentKnownItemNodes.Remove(itemNo);
                return null;
            }
        }

        // 目标清单里的一条：按编号找，还是按精确名称找。两种可以混在同一个清单里。
        private sealed class AgentTargetEntry
        {
            public bool ByName;
            public string Value = "";

            public override string ToString()
            {
                if (ByName)
                {
                    return "【名称】" + Value;
                }

                string kind = AgentQuotaCodeKind(Value);
                return "【编号】" + Value + (kind.Length > 0 ? "　（" + kind + "）" : "");
            }
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
            private bool refilling;

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
                top.Height = 36;
                top.Padding = new Padding(10, 6, 10, 6);

                filterBox = new TextBox();
                filterBox.Dock = DockStyle.Fill;
                filterBox.TextChanged += delegate { RefillList(); };

                Label filterHint = new Label();
                filterHint.Text = "筛选";
                filterHint.Dock = DockStyle.Left;
                filterHint.Width = 40;
                filterHint.TextAlign = ContentAlignment.MiddleLeft;

                top.Controls.Add(filterBox);
                top.Controls.Add(filterHint);

                Panel bottom = new Panel();
                bottom.Dock = DockStyle.Bottom;
                bottom.Height = 46;
                bottom.Padding = new Padding(10, 8, 10, 8);

                Button ok = new Button();
                ok.Text = "确定";
                ok.Width = 88;
                ok.Dock = DockStyle.Right;
                ok.DialogResult = DialogResult.OK;
                ok.Click += delegate { CollectResult(); };

                Panel spacer = new Panel();
                spacer.Dock = DockStyle.Right;
                spacer.Width = 8;

                Button cancel = new Button();
                cancel.Text = "取消";
                cancel.Width = 88;
                cancel.Dock = DockStyle.Right;
                cancel.DialogResult = DialogResult.Cancel;

                countLabel = new Label();
                countLabel.Dock = DockStyle.Fill;
                countLabel.TextAlign = ContentAlignment.MiddleLeft;
                countLabel.ForeColor = AgentPanelHintFore;

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
                    checkedList.BorderStyle = BorderStyle.None;
                    checkedList.ItemCheck += delegate(object sender, ItemCheckEventArgs e)
                    {
                        // 重填列表时 Items.Add(item, checked) 也会触发本事件，这时不必处理。
                        if (refilling)
                        {
                            return;
                        }

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

                        // 计数来自 checkedKeys，上面已经同步更新，直接刷新即可。
                        // 这里不能用 BeginInvoke：构造期间窗口句柄还没建，会直接抛异常。
                        UpdateCount();
                    };
                    Controls.Add(checkedList);
                }
                else
                {
                    singleList = new ListBox();
                    singleList.Dock = DockStyle.Fill;
                    singleList.IntegralHeight = false;
                    singleList.BorderStyle = BorderStyle.None;
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
                    refilling = true;
                    try
                    {
                        checkedList.BeginUpdate();
                        checkedList.Items.Clear();
                        foreach (AgentPickItem item in shown)
                        {
                            checkedList.Items.Add(item, checkedKeys.Contains(item.Key));
                        }

                        checkedList.EndUpdate();
                    }
                    finally
                    {
                        refilling = false;
                    }
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
                    countLabel.Text = "已勾选 " + checkedKeys.Count.ToString(CultureInfo.InvariantCulture) +
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
            private const string TabReplace = "替换定额";
            private const string TabRows = "增删定额";
            private const string TabCross = "跨条目复制";
            private const string TabText = "指令";

            private readonly Form mainForm;

            // 第一区：作用范围
            private readonly TextBox unitBox;
            private readonly CheckBox allUnitsBox;
            private readonly ListBox itemListBox;
            private readonly CheckBox includeChildrenBox;
            private readonly Button takeCodeButton;
            private readonly Button takeNameButton;
            private readonly ListBox targetListBox;
            private readonly Label targetHintLabel;

            // 第二区：操作
            private readonly TabControl tabs;
            private readonly ComboBox fieldBox;
            private readonly ComboBox valueActionBox;
            private readonly Label valueLabel;
            private readonly TextBox valueBox;
            private readonly Label valueHintLabel;

            private readonly DataGridView replaceGrid;
            private readonly ComboBox rowActionBox;
            private readonly Label insertGridLabel;
            private Panel insertGridHeader;
            private readonly DataGridView insertGrid;
            private readonly Label rowHintLabel;

            private readonly ComboBox crossActionBox;
            private readonly ListBox crossTargetList;
            private readonly Label crossHintLabel;

            private readonly TextBox textBox;
            private readonly Button textSendButton;

            // 底部
            private readonly Label sentenceLabel;
            private readonly Button undoButton;
            private readonly Button redoButton;
            private readonly Button previewButton;
            private readonly Label statusLabel;
            private readonly ToolTip toolTip = new ToolTip();

            private readonly List<AgentUnitOption> unitOptions = new List<AgentUnitOption>();
            private readonly List<AgentItemOption> itemOptions = new List<AgentItemOption>();
            private readonly List<string> selectedUnitKeys = new List<string>();
            private readonly List<string> selectedItemNos = new List<string>();
            private readonly List<AgentTargetEntry> targetEntries = new List<AgentTargetEntry>();
            private readonly List<string> crossTargetItems = new List<string>();

            private bool busy;
            private bool suppressSentence;
            private string scopeProjectIdentity = "";

            public AgentPanelWindow(Form mainForm)
            {
                this.mainForm = mainForm;
                Text = "智能指令助手 (Ctrl+Q)　—　跨条目 / 跨单元批量操作";
                StartPosition = FormStartPosition.Manual;
                Size = new Size(920, 830);
                MinimumSize = new Size(840, 700);
                ShowInTaskbar = false;
                Font = new Font("Microsoft YaHei UI", 9f, FontStyle.Regular, GraphicsUnit.Point);
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

                Font titleFont = new Font(Font, FontStyle.Bold);

                // ===================== 第一区：作用范围 =====================
                Panel scopePanel = new Panel();
                scopePanel.Dock = DockStyle.Top;
                scopePanel.Height = 202;
                scopePanel.BackColor = AgentPanelSectionBack;
                scopePanel.Padding = new Padding(14, 8, 14, 10);


                // --- 单元行 ---
                Panel unitRow = new Panel();
                unitRow.Dock = DockStyle.Top;
                unitRow.Height = 28;

                unitBox = new TextBox();
                unitBox.Dock = DockStyle.Fill;
                unitBox.ReadOnly = true;
                unitBox.BackColor = Color.White;

                Label unitLabel = new Label();
                unitLabel.Dock = DockStyle.Left;
                unitLabel.Width = 44;
                unitLabel.TextAlign = ContentAlignment.MiddleLeft;
                unitLabel.Text = "单元";

                Button unitPickButton = new Button();
                unitPickButton.Dock = DockStyle.Right;
                unitPickButton.Width = 92;
                unitPickButton.Text = "选单元…";
                unitPickButton.Click += delegate { PickUnits(); };

                allUnitsBox = new CheckBox();
                allUnitsBox.Dock = DockStyle.Right;
                allUnitsBox.Width = 92;
                allUnitsBox.Text = "所有单元";
                allUnitsBox.TextAlign = ContentAlignment.MiddleLeft;
                allUnitsBox.CheckedChanged += delegate
                {
                    unitPickButton.Enabled = !allUnitsBox.Checked;
                    UpdateUnitBox();
                    UpdateSentence();
                };

                unitRow.Controls.Add(unitBox);
                unitRow.Controls.Add(unitLabel);
                unitRow.Controls.Add(unitPickButton);
                unitRow.Controls.Add(allUnitsBox);

                Panel gapUnderUnit = new Panel();
                gapUnderUnit.Dock = DockStyle.Top;
                gapUnderUnit.Height = 6;

                // --- 条目行 ---
                Panel itemRow = new Panel();
                itemRow.Dock = DockStyle.Top;
                itemRow.Height = 82;

                itemListBox = new ListBox();
                itemListBox.Dock = DockStyle.Fill;
                itemListBox.IntegralHeight = false;
                itemListBox.SelectionMode = SelectionMode.MultiExtended;
                itemListBox.BorderStyle = BorderStyle.FixedSingle;

                Label itemLabel = new Label();
                itemLabel.Dock = DockStyle.Left;
                itemLabel.Width = 44;
                itemLabel.TextAlign = ContentAlignment.TopLeft;
                itemLabel.Padding = new Padding(0, 4, 0, 0);
                itemLabel.Text = "条目";

                Panel itemSideBar = new Panel();
                itemSideBar.Dock = DockStyle.Right;
                itemSideBar.Width = 190;
                itemSideBar.Padding = new Padding(8, 0, 0, 0);

                Button addItemButton = new Button();
                addItemButton.Dock = DockStyle.Fill;
                addItemButton.Text = "◀ 添加条目";
                addItemButton.TextAlign = ContentAlignment.MiddleLeft;
                addItemButton.Padding = new Padding(6, 0, 0, 0);
                addItemButton.Click += delegate { AddItemFromTree(false); };

                Panel itemSmallRow = new Panel();
                itemSmallRow.Dock = DockStyle.Bottom;
                itemSmallRow.Height = 26;
                itemSmallRow.Padding = new Padding(0, 2, 0, 2);

                Button clearItemButton = new Button();
                clearItemButton.Dock = DockStyle.Left;
                clearItemButton.Width = 88;
                clearItemButton.Text = "全部清空";
                clearItemButton.Click += delegate { ClearItems(); };

                Button removeItemButton = new Button();
                removeItemButton.Dock = DockStyle.Left;
                removeItemButton.Width = 88;
                removeItemButton.Text = "移除选中";
                removeItemButton.Click += delegate { RemoveSelectedItems(); };

                itemSmallRow.Controls.Add(clearItemButton);
                itemSmallRow.Controls.Add(removeItemButton);

                includeChildrenBox = new CheckBox();
                includeChildrenBox.Dock = DockStyle.Bottom;
                includeChildrenBox.Height = 22;
                includeChildrenBox.Checked = true;
                includeChildrenBox.Text = "含所有子条目";
                includeChildrenBox.CheckedChanged += delegate { UpdateSentence(); };

                itemSideBar.Controls.Add(addItemButton);
                itemSideBar.Controls.Add(itemSmallRow);
                itemSideBar.Controls.Add(includeChildrenBox);

                itemRow.Controls.Add(itemListBox);
                itemRow.Controls.Add(itemLabel);
                itemRow.Controls.Add(itemSideBar);

                Panel gapUnderItem = new Panel();
                gapUnderItem.Dock = DockStyle.Top;
                gapUnderItem.Height = 6;

                // --- 目标行 ---
                Panel targetArea = new Panel();
                targetArea.Dock = DockStyle.Fill;

                Panel targetHeader = new Panel();
                targetHeader.Dock = DockStyle.Top;
                targetHeader.Height = 28;

                targetHintLabel = new Label();
                targetHintLabel.Dock = DockStyle.Fill;
                targetHintLabel.TextAlign = ContentAlignment.MiddleLeft;
                targetHintLabel.ForeColor = AgentPanelWarnFore;

                Label targetLabel = new Label();
                targetLabel.Dock = DockStyle.Left;
                targetLabel.Width = 44;
                targetLabel.TextAlign = ContentAlignment.MiddleLeft;
                targetLabel.Text = "目标";

                takeCodeButton = new Button();
                takeCodeButton.Dock = DockStyle.Right;
                takeCodeButton.Width = 132;
                takeCodeButton.Text = "取选中行的编号";
                takeCodeButton.Click += delegate { TakeTargetsFromHostGrid(false); };

                takeNameButton = new Button();
                takeNameButton.Dock = DockStyle.Right;
                takeNameButton.Width = 132;
                takeNameButton.Text = "取选中行的名称";
                takeNameButton.Click += delegate { TakeTargetsFromHostGrid(true); };

                Button removeTargetButton = new Button();
                removeTargetButton.Dock = DockStyle.Right;
                removeTargetButton.Width = 88;
                removeTargetButton.Text = "移除选中";
                removeTargetButton.Click += delegate { RemoveSelectedTargets(); };

                Button clearTargetButton = new Button();
                clearTargetButton.Dock = DockStyle.Right;
                clearTargetButton.Width = 88;
                clearTargetButton.Text = "全部清空";
                clearTargetButton.Click += delegate
                {
                    targetEntries.Clear();
                    UpdateTargetDisplay();
                    UpdateSentence();
                    SetStatus("目标已清空，将作用于范围内全部行。", false);
                };

                targetHeader.Controls.Add(targetHintLabel);
                targetHeader.Controls.Add(targetLabel);
                targetHeader.Controls.Add(takeCodeButton);
                targetHeader.Controls.Add(takeNameButton);
                targetHeader.Controls.Add(removeTargetButton);
                targetHeader.Controls.Add(clearTargetButton);

                targetListBox = new ListBox();
                targetListBox.Dock = DockStyle.Fill;
                targetListBox.IntegralHeight = false;
                targetListBox.SelectionMode = SelectionMode.MultiExtended;
                targetListBox.BorderStyle = BorderStyle.FixedSingle;
                // 目标条目都很短，横向排开，别让右边一大片空着。
                targetListBox.MultiColumn = true;
                targetListBox.ColumnWidth = 250;

                targetArea.Controls.Add(targetListBox);
                targetArea.Controls.Add(targetHeader);

                scopePanel.Controls.Add(targetArea);
                scopePanel.Controls.Add(gapUnderItem);
                scopePanel.Controls.Add(itemRow);
                scopePanel.Controls.Add(gapUnderUnit);
                scopePanel.Controls.Add(unitRow);

                // ===================== 第二区：选操作 =====================
                Panel tabsArea = new Panel();
                tabsArea.Dock = DockStyle.Fill;
                tabsArea.Padding = new Padding(14, 8, 14, 4);

                tabs = new TabControl();
                tabs.Dock = DockStyle.Fill;
                tabs.Padding = new Point(16, 5);

                // --- 页签：改数值 ---
                TabPage valuePage = new TabPage(TabValue);
                valuePage.UseVisualStyleBackColor = true;
                valuePage.Padding = new Padding(16, 16, 16, 10);

                Label fieldLabel = new Label();
                fieldLabel.SetBounds(6, 12, 64, 22);
                fieldLabel.Text = "操作对象";

                fieldBox = new ComboBox();
                fieldBox.DropDownStyle = ComboBoxStyle.DropDownList;
                fieldBox.SetBounds(74, 8, 120, 25);
                fieldBox.Items.AddRange(new object[] { "工程数量", "单价", "定额编号", "定额调整" });
                fieldBox.SelectedIndex = 0;
                fieldBox.SelectedIndexChanged += delegate { RefreshValueActions(); };

                Label actionLabel = new Label();
                actionLabel.SetBounds(210, 12, 40, 22);
                actionLabel.Text = "操作";

                valueActionBox = new ComboBox();
                valueActionBox.DropDownStyle = ComboBoxStyle.DropDownList;
                valueActionBox.SetBounds(254, 8, 132, 25);
                valueActionBox.SelectedIndexChanged += delegate { RefreshValueInput(); };

                valueLabel = new Label();
                valueLabel.SetBounds(402, 12, 60, 22);
                valueLabel.Text = "系数";

                valueBox = new TextBox();
                valueBox.SetBounds(466, 8, 112, 25);
                valueBox.TextChanged += delegate { UpdateSentence(); };

                valueHintLabel = new Label();
                valueHintLabel.SetBounds(6, 48, 760, 90);
                valueHintLabel.ForeColor = AgentPanelHintFore;

                valuePage.Controls.Add(fieldLabel);
                valuePage.Controls.Add(fieldBox);
                valuePage.Controls.Add(actionLabel);
                valuePage.Controls.Add(valueActionBox);
                valuePage.Controls.Add(valueLabel);
                valuePage.Controls.Add(valueBox);
                valuePage.Controls.Add(valueHintLabel);

                // --- 页签：替换定额 ---
                TabPage replacePage = new TabPage(TabReplace);
                replacePage.UseVisualStyleBackColor = true;
                replacePage.Padding = new Padding(16, 12, 16, 12);

                replaceGrid = BuildQuotaGrid();
                replaceGrid.Dock = DockStyle.Fill;

                Panel replaceGridHeader = new Panel();
                replaceGridHeader.Dock = DockStyle.Top;
                replaceGridHeader.Height = 28;

                Label replaceGridLabel = new Label();
                replaceGridLabel.Dock = DockStyle.Fill;
                replaceGridLabel.TextAlign = ContentAlignment.MiddleLeft;
                replaceGridLabel.Text = "替换成（数量留空 = 不改数量）";

                Button replaceFillButton = new Button();
                replaceFillButton.Dock = DockStyle.Right;
                replaceFillButton.Width = 132;
                replaceFillButton.Text = "取选中行的编号";
                replaceFillButton.Click += delegate { FillGridCodesFromHostGrid(replaceGrid); };

                replaceGridHeader.Controls.Add(replaceGridLabel);
                replaceGridHeader.Controls.Add(replaceFillButton);

                Label replaceHint = new Label();
                replaceHint.Dock = DockStyle.Top;
                replaceHint.Height = 52;
                replaceHint.ForeColor = AgentPanelHintFore;
                replaceHint.Text = "被替换的就是上面「目标」里列的那些行，不用在这里再选一遍。\r\n" +
                    "一对一、一对多（拆成几条）、多对一（合并成一条）都支持。拆成多条时只能作用于一个单元。";

                replacePage.Controls.Add(replaceGrid);
                replacePage.Controls.Add(replaceGridHeader);
                replacePage.Controls.Add(replaceHint);

                // --- 页签：增删定额 ---
                TabPage rowPage = new TabPage(TabRows);
                rowPage.UseVisualStyleBackColor = true;
                rowPage.Padding = new Padding(16, 12, 16, 12);

                insertGrid = BuildQuotaGrid();
                insertGrid.Dock = DockStyle.Fill;

                insertGridHeader = new Panel();
                insertGridHeader.Dock = DockStyle.Top;
                insertGridHeader.Height = 28;

                insertGridLabel = new Label();
                insertGridLabel.Dock = DockStyle.Fill;
                insertGridLabel.TextAlign = ContentAlignment.MiddleLeft;
                insertGridLabel.Text = "要新增的定额（数量留空 = 不填）";

                Button insertFillButton = new Button();
                insertFillButton.Dock = DockStyle.Right;
                insertFillButton.Width = 132;
                insertFillButton.Text = "取选中行的编号";
                insertFillButton.Click += delegate { FillGridCodesFromHostGrid(insertGrid); };

                insertGridHeader.Controls.Add(insertGridLabel);
                insertGridHeader.Controls.Add(insertFillButton);

                rowHintLabel = new Label();
                rowHintLabel.Dock = DockStyle.Top;
                rowHintLabel.Height = 44;
                rowHintLabel.ForeColor = AgentPanelHintFore;

                Panel rowActionRow = new Panel();
                rowActionRow.Dock = DockStyle.Top;
                rowActionRow.Height = 34;

                Label rowActionLabel = new Label();
                rowActionLabel.SetBounds(0, 6, 40, 22);
                rowActionLabel.Text = "动作";

                rowActionBox = new ComboBox();
                rowActionBox.DropDownStyle = ComboBoxStyle.DropDownList;
                rowActionBox.SetBounds(48, 2, 180, 25);
                rowActionBox.Items.AddRange(new object[] { "新增定额", "删除目标定额" });
                rowActionBox.SelectedIndex = 0;
                rowActionBox.SelectedIndexChanged += delegate { RefreshRowAction(); };

                rowActionRow.Controls.Add(rowActionLabel);
                rowActionRow.Controls.Add(rowActionBox);

                rowPage.Controls.Add(insertGrid);
                rowPage.Controls.Add(insertGridHeader);
                rowPage.Controls.Add(rowHintLabel);
                rowPage.Controls.Add(rowActionRow);

                // --- 页签：跨条目 ---
                TabPage crossPage = new TabPage(TabCross);
                crossPage.UseVisualStyleBackColor = true;
                crossPage.Padding = new Padding(16, 16, 16, 10);

                Label crossActionLabel = new Label();
                crossActionLabel.SetBounds(6, 8, 40, 22);
                crossActionLabel.Text = "动作";

                crossActionBox = new ComboBox();
                crossActionBox.DropDownStyle = ComboBoxStyle.DropDownList;
                crossActionBox.SetBounds(50, 4, 124, 25);
                crossActionBox.Items.AddRange(new object[] { "复制到", "移动到" });
                crossActionBox.SelectedIndex = 0;
                crossActionBox.SelectedIndexChanged += delegate { RefreshCrossAction(); };

                Panel crossActionRow = new Panel();
                crossActionRow.Dock = DockStyle.Top;
                crossActionRow.Height = 34;
                crossActionRow.Controls.Add(crossActionLabel);
                crossActionRow.Controls.Add(crossActionBox);

                crossHintLabel = new Label();
                crossHintLabel.Dock = DockStyle.Top;
                crossHintLabel.Height = 44;
                crossHintLabel.ForeColor = AgentPanelHintFore;

                Panel crossTargetHeader = new Panel();
                crossTargetHeader.Dock = DockStyle.Top;
                crossTargetHeader.Height = 28;

                Label crossTargetLabel = new Label();
                crossTargetLabel.Dock = DockStyle.Fill;
                crossTargetLabel.TextAlign = ContentAlignment.MiddleLeft;
                crossTargetLabel.Text = "目标条目（在左侧树点中后按「添加条目」，可反复加）";

                Button crossAddButton = new Button();
                crossAddButton.Dock = DockStyle.Right;
                crossAddButton.Width = 110;
                crossAddButton.Text = "◀ 添加条目";
                crossAddButton.Click += delegate { AddCrossTargetFromTree(); };

                Button crossRemoveButton = new Button();
                crossRemoveButton.Dock = DockStyle.Right;
                crossRemoveButton.Width = 88;
                crossRemoveButton.Text = "移除选中";
                crossRemoveButton.Click += delegate { RemoveSelectedCrossTargets(); };

                Button crossClearButton = new Button();
                crossClearButton.Dock = DockStyle.Right;
                crossClearButton.Width = 88;
                crossClearButton.Text = "全部清空";
                crossClearButton.Click += delegate { ClearCrossTargets(); };

                crossTargetHeader.Controls.Add(crossTargetLabel);
                crossTargetHeader.Controls.Add(crossAddButton);
                crossTargetHeader.Controls.Add(crossRemoveButton);
                crossTargetHeader.Controls.Add(crossClearButton);

                crossTargetList = new ListBox();
                crossTargetList.Dock = DockStyle.Fill;
                crossTargetList.IntegralHeight = false;
                crossTargetList.SelectionMode = SelectionMode.MultiExtended;
                crossTargetList.BorderStyle = BorderStyle.FixedSingle;

                crossPage.Controls.Add(crossTargetList);
                crossPage.Controls.Add(crossTargetHeader);
                crossPage.Controls.Add(crossHintLabel);
                crossPage.Controls.Add(crossActionRow);

                // --- 页签：说一句话 ---
                TabPage textPage = new TabPage(TabText);
                textPage.UseVisualStyleBackColor = true;
                textPage.Padding = new Padding(16, 12, 16, 12);

                textBox = new TextBox();
                textBox.Dock = DockStyle.Fill;
                textBox.Multiline = true;
                textBox.ScrollBars = ScrollBars.Vertical;
                textBox.Font = new Font(Font.FontFamily, 10.5f);

                Panel textSideBar = new Panel();
                textSideBar.Dock = DockStyle.Right;
                textSideBar.Width = 116;
                textSideBar.Padding = new Padding(8, 0, 0, 0);

                textSendButton = new Button();
                textSendButton.Dock = DockStyle.Top;
                textSendButton.Height = 34;
                textSendButton.Text = "交给 AI";
                textSendButton.Click += delegate { SubmitText(); };

                Panel textButtonGap = new Panel();
                textButtonGap.Dock = DockStyle.Top;
                textButtonGap.Height = 6;

                Button helpButton = new Button();
                helpButton.Dock = DockStyle.Top;
                helpButton.Height = 30;
                helpButton.Text = "指令帮助";
                helpButton.Click += delegate { ShowHelpDialog(); };

                textSideBar.Controls.Add(helpButton);
                textSideBar.Controls.Add(textButtonGap);
                textSideBar.Controls.Add(textSendButton);

                Label textHintLabel = new Label();
                textHintLabel.Dock = DockStyle.Bottom;
                textHintLabel.Height = 74;
                textHintLabel.ForeColor = AgentPanelHintFore;
                textHintLabel.Text = "兜底通道：上面的按钮拼不出来时才用。新建单元、运输方案、材料价方案只能从这里走。\r\n" +
                    "也可以直接输入 撤销 / 重做 / 帮助 / 探查 关键词。\r\n" +
                    "自然语言需要已配置 RecoQuotaData/deepseek-settings.json；确定性写法（如 工程数量 0101-01 *0.85）不需要 AI。";

                textPage.Controls.Add(textBox);
                textPage.Controls.Add(textSideBar);
                textPage.Controls.Add(textHintLabel);

                tabs.TabPages.Add(valuePage);
                tabs.TabPages.Add(replacePage);
                tabs.TabPages.Add(rowPage);
                tabs.TabPages.Add(crossPage);
                tabs.TabPages.Add(textPage);
                tabs.SelectedIndexChanged += delegate
                {
                    RefreshScopeAvailability();
                    UpdateSentence();
                };

                tabsArea.Controls.Add(tabs);

                // ===================== 操作条 =====================
                Panel actionBar = new Panel();
                actionBar.Dock = DockStyle.Bottom;
                actionBar.Height = 70;
                actionBar.Padding = new Padding(14, 6, 14, 8);

                sentenceLabel = new Label();
                sentenceLabel.Dock = DockStyle.Top;
                sentenceLabel.Height = 28;
                sentenceLabel.TextAlign = ContentAlignment.MiddleLeft;
                sentenceLabel.ForeColor = AgentPanelTitleFore;

                Panel buttonRow = new Panel();
                buttonRow.Dock = DockStyle.Fill;

                previewButton = new Button();
                previewButton.Dock = DockStyle.Right;
                previewButton.Width = 132;
                previewButton.Text = "生成预览 ▸";
                previewButton.Font = titleFont;
                previewButton.BackColor = Color.FromArgb(222, 236, 252);
                previewButton.Click += delegate { SubmitPanelCommand(); };

                Panel buttonSpacer = new Panel();
                buttonSpacer.Dock = DockStyle.Fill;

                undoButton = new Button();
                undoButton.Dock = DockStyle.Left;
                undoButton.Width = 116;
                undoButton.Text = "撤销上一步";
                undoButton.Click += delegate { PreviewUndo(); };

                Panel undoGap = new Panel();
                undoGap.Dock = DockStyle.Left;
                undoGap.Width = 6;

                redoButton = new Button();
                redoButton.Dock = DockStyle.Left;
                redoButton.Width = 76;
                redoButton.Text = "重做";
                redoButton.Click += delegate { PreviewRedo(); };

                buttonRow.Controls.Add(buttonSpacer);
                buttonRow.Controls.Add(redoButton);
                buttonRow.Controls.Add(undoGap);
                buttonRow.Controls.Add(undoButton);
                buttonRow.Controls.Add(previewButton);

                actionBar.Controls.Add(buttonRow);
                actionBar.Controls.Add(sentenceLabel);

                // ===================== 第三区：确认执行 =====================
                statusLabel = new Label();
                statusLabel.Dock = DockStyle.Bottom;
                statusLabel.Height = 26;
                statusLabel.TextAlign = ContentAlignment.MiddleLeft;
                statusLabel.BackColor = Color.FromArgb(240, 242, 246);
                statusLabel.Padding = new Padding(10, 0, 6, 0);

                Controls.Add(tabsArea);
                Controls.Add(actionBar);
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
                SetStatus("用法：左侧树点中条目 → 点「添加条目」，可反复加；再选单元；" +
                    "然后在主程序定额表多选几行，点「取选中行的编号」或「取选中行的名称」。", false);
            }

            private static DataGridView BuildQuotaGrid()
            {
                DataGridView grid = new DataGridView();
                grid.AllowUserToAddRows = true;
                grid.AllowUserToDeleteRows = true;
                grid.RowHeadersVisible = false;
                grid.BackgroundColor = Color.White;
                grid.BorderStyle = BorderStyle.FixedSingle;
                grid.EnableHeadersVisualStyles = false;
                grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(238, 242, 248);
                grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                grid.Columns.Add("Code", "定额编号");
                grid.Columns.Add("Quantity", "工程数量");
                grid.Columns["Code"].FillWeight = 62;
                grid.Columns["Quantity"].FillWeight = 38;
                return grid;
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
                UpdateCrossTargetList();
                targetEntries.Clear();
                UpdateTargetDisplay();
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
                    unitBox.Text = "（未识别当前单元，请点右边「选单元…」）";
                    return;
                }

                List<string> parts = new List<string>();
                foreach (string key in selectedUnitKeys)
                {
                    parts.Add(UnitDisplayOf(key));
                }

                unitBox.Text = "共 " + selectedUnitKeys.Count.ToString(CultureInfo.InvariantCulture) + " 个：" +
                    String.Join("、", parts.ToArray());
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

                using (AgentPickerDialog dialog = new AgentPickerDialog("勾选要操作的单元（可多选）", items, true, selectedUnitKeys))
                {
                    if (dialog.ShowDialog(this) != DialogResult.OK)
                    {
                        return;
                    }

                    selectedUnitKeys.Clear();
                    selectedUnitKeys.AddRange(dialog.SelectedKeys);
                }

                UpdateUnitBox();
                UpdateSentence();
                SetStatus("已选 " + selectedUnitKeys.Count.ToString(CultureInfo.InvariantCulture) + " 个单元。", false);
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
                            SetStatus("左侧树上还没有选中节点，请先在树上点一个条目。", true);
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
                    RememberAgentItemNode(itemNo, node);
                    UpdateItemList();
                    UpdateSentence();
                    if (!quiet)
                    {
                        SetStatus("已加入 " + ItemDisplayOf(itemNo) + "，条目共 " +
                            selectedItemNos.Count.ToString(CultureInfo.InvariantCulture) +
                            " 个。可以继续在左侧树点别的条目再加。", false);
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
                UpdateSentence();
            }

            private void ClearItems()
            {
                selectedItemNos.Clear();
                UpdateItemList();
                UpdateSentence();
                SetStatus("条目列表已清空。", false);
            }

            // ===== 目标定额 / 材料 / SF =====

            private void UpdateTargetDisplay()
            {
                targetListBox.Items.Clear();
                foreach (AgentTargetEntry entry in targetEntries)
                {
                    targetListBox.Items.Add(entry);
                }

                int codeCount = 0;
                int nameCount = 0;
                foreach (AgentTargetEntry entry in targetEntries)
                {
                    if (entry.ByName)
                    {
                        nameCount++;
                    }
                    else
                    {
                        codeCount++;
                    }
                }

                if (targetEntries.Count == 0)
                {
                    targetHintLabel.ForeColor = AgentPanelWarnFore;
                    targetHintLabel.Text = "未选目标 → 作用于范围内全部行";
                }
                else
                {
                    targetHintLabel.ForeColor = AgentPanelOkFore;
                    List<string> parts = new List<string>();
                    if (codeCount > 0)
                    {
                        parts.Add("编号 " + codeCount.ToString(CultureInfo.InvariantCulture) + " 项");
                    }

                    if (nameCount > 0)
                    {
                        parts.Add("名称 " + nameCount.ToString(CultureInfo.InvariantCulture) + " 项");
                    }

                    targetHintLabel.Text = "在范围内查找：" + String.Join(" + ", parts.ToArray());
                }
            }

            // 在主程序定额输入表里多选几行，把这些行的编号或精确名称取过来当查找依据。
            // 注意语义：取的是"找什么"，不是"只改这几行"——真正作用的是上面条目×单元范围内所有匹配的行。
            private void TakeTargetsFromHostGrid(bool byName)
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
                    SetStatus(byName ? "选中的行读不到项目名称。" : "选中的行读不到定额编号。", true);
                    return;
                }

                int added = 0;
                foreach (string value in picked)
                {
                    bool exists = false;
                    foreach (AgentTargetEntry entry in targetEntries)
                    {
                        if (entry.ByName == byName && String.Equals(entry.Value, value, StringComparison.Ordinal))
                        {
                            exists = true;
                            break;
                        }
                    }

                    if (exists)
                    {
                        continue;
                    }

                    AgentTargetEntry fresh = new AgentTargetEntry();
                    fresh.ByName = byName;
                    fresh.Value = value;
                    targetEntries.Add(fresh);
                    added++;
                }

                UpdateTargetDisplay();
                UpdateSentence();
                SetStatus("从定额表 " + rowCount.ToString(CultureInfo.InvariantCulture) + " 个选中行取到 " +
                    added.ToString(CultureInfo.InvariantCulture) + " 个新" + (byName ? "名称" : "编号") +
                    "。这些是查找依据，会在上面的条目 × 单元范围内逐一查找。", false);
            }

            private void RemoveSelectedTargets()
            {
                List<int> indexes = new List<int>();
                foreach (int index in targetListBox.SelectedIndices)
                {
                    indexes.Add(index);
                }

                if (indexes.Count == 0)
                {
                    SetStatus("请先在目标列表里选中要移除的行。", true);
                    return;
                }

                indexes.Sort();
                for (int i = indexes.Count - 1; i >= 0; i--)
                {
                    if (indexes[i] >= 0 && indexes[i] < targetEntries.Count)
                    {
                        targetEntries.RemoveAt(indexes[i]);
                    }
                }

                UpdateTargetDisplay();
                UpdateSentence();
            }

            private List<string> TargetCodes()
            {
                List<string> codes = new List<string>();
                foreach (AgentTargetEntry entry in targetEntries)
                {
                    if (!entry.ByName)
                    {
                        codes.Add(entry.Value);
                    }
                }

                return codes;
            }

            private List<string> TargetNames()
            {
                List<string> names = new List<string>();
                foreach (AgentTargetEntry entry in targetEntries)
                {
                    if (entry.ByName)
                    {
                        names.Add(entry.Value);
                    }
                }

                return names;
            }

            private void AddCrossTargetFromTree()
            {
                TreeView tree = GetField<TreeView>(mainForm, "Tv_tree");
                TreeNode node = tree != null ? tree.SelectedNode : GetField<TreeNode>(mainForm, "CurrNode");
                if (node == null)
                {
                    SetStatus("左侧树上还没有选中节点，请先在树上点一个目标条目。", true);
                    return;
                }

                string itemNo;
                try
                {
                    SqlConnection hostConn = GetProjectConnection(mainForm);
                    itemNo = hostConn == null ? null : ResolveChapterNo(mainForm, hostConn, node);
                }
                catch (Exception ex)
                {
                    SetStatus("读取当前条目失败：" + ex.Message, true);
                    Log("Agent panel add cross target failed: " + ex);
                    return;
                }

                if (String.IsNullOrEmpty(itemNo))
                {
                    SetStatus("无法识别当前条目编号。", true);
                    return;
                }

                if (crossTargetItems.Contains(itemNo))
                {
                    SetStatus("目标条目 " + ItemDisplayOf(itemNo) + " 已经在列表里了。", false);
                    return;
                }

                bool move = crossActionBox.Text == "移动到";
                if (move && crossTargetItems.Count > 0)
                {
                    // 移动只能有一个目标，直接换成新点的这个。
                    crossTargetItems.Clear();
                }

                crossTargetItems.Add(itemNo);
                RememberAgentItemNode(itemNo, node);
                UpdateCrossTargetList();
                UpdateSentence();
                SetStatus("已加入目标条目 " + ItemDisplayOf(itemNo) +
                    (move ? "。移动只能有一个目标。" : "，共 " +
                        crossTargetItems.Count.ToString(CultureInfo.InvariantCulture) + " 个。"), false);
            }

            private void RemoveSelectedCrossTargets()
            {
                List<int> indexes = new List<int>();
                foreach (int index in crossTargetList.SelectedIndices)
                {
                    indexes.Add(index);
                }

                if (indexes.Count == 0)
                {
                    SetStatus("请先在目标条目列表里选中要移除的行。", true);
                    return;
                }

                indexes.Sort();
                for (int i = indexes.Count - 1; i >= 0; i--)
                {
                    if (indexes[i] >= 0 && indexes[i] < crossTargetItems.Count)
                    {
                        crossTargetItems.RemoveAt(indexes[i]);
                    }
                }

                UpdateCrossTargetList();
                UpdateSentence();
            }

            private void ClearCrossTargets()
            {
                crossTargetItems.Clear();
                UpdateCrossTargetList();
                UpdateSentence();
                SetStatus("目标条目已清空。", false);
            }

            private void UpdateCrossTargetList()
            {
                crossTargetList.Items.Clear();
                foreach (string itemNo in crossTargetItems)
                {
                    crossTargetList.Items.Add(ItemDisplayOf(itemNo));
                }
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
                takeCodeButton.Enabled = targetUsed;
                takeNameButton.Enabled = targetUsed;
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
                    valueActionBox.Items.AddRange(new object[] { "乘以", "除以", "改成", "清空数量", "删除系数" });
                }
                else if (field == "单价")
                {
                    valueActionBox.Items.AddRange(new object[] { "乘以", "除以", "改成" });
                }
                else if (field == "定额编号")
                {
                    valueActionBox.Items.AddRange(new object[] { "乘以", "除以", "删除系数" });
                }
                else
                {
                    valueActionBox.Items.AddRange(new object[] { "改成", "追加", "删除调整内容" });
                }

                valueActionBox.SelectedIndex = 0;
                suppressSentence = false;
                RefreshValueInput();
            }

            private void RefreshValueInput()
            {
                string field = fieldBox.Text;
                string action = valueActionBox.Text;
                bool needsValue = action != "清空数量";
                valueBox.Enabled = needsValue;
                if (!needsValue)
                {
                    valueBox.Text = "";
                }

                if (action == "乘以" || action == "除以")
                {
                    valueLabel.Text = "系数";
                }
                else if ((action == "删除系数" || action == "删除调整内容"))
                {
                    valueLabel.Text = "要删的";
                }
                else if (field == "定额调整")
                {
                    valueLabel.Text = "调整内容";
                }
                else if (!needsValue)
                {
                    valueLabel.Text = "";
                }
                else
                {
                    valueLabel.Text = "改成";
                }

                if (action == "清空数量")
                {
                    valueHintLabel.Text = "把工程数量输入和计算数量都置成空白（不是 0）。\r\n要改成某个具体数字，请选「改成」。";
                }
                else if (action == "删除调整内容")
                {
                    valueHintLabel.Text = "从定额调整里去掉指定的一段，例如填 /XG1 就把原串里的 /XG1 抠掉。\r\n" +
                        "原串里没有这一段的行会自动跳过。";
                }
                else if (action == "删除系数")
                {
                    valueHintLabel.Text = "撤掉之前乘上去的系数。填要删的那一段，例如　*0.85　。\r\n" +
                        "工程数量：原来是　(100)*0.85　，删掉后会连同外层括号一起还原成　100　。\r\n" +
                        "定额编号：原来是　LY-21*9　，删掉后还原成　LY-21　。字段里没有这一段的行会跳过。";
                }
                else if (field == "定额编号")
                {
                    valueHintLabel.Text = "在定额编号后面直接追加乘除系数（不加括号），例如 LY-21 变成 LY-21*9，" +
                        "这是软件原生的缩放定额写法。\r\n改完通常要在主程序里手工触发一次重算。";
                }
                else if (field == "定额调整")
                {
                    valueHintLabel.Text = "调整内容按整串写入，例如 /XG1 、 /1294861,,1 。\r\n「追加」是拼在原有调整后面，「改成」是整串替换。";
                }
                else if (field == "单价")
                {
                    valueHintLabel.Text = "SF / SH / SQ / ZLF / LF / TLF 这类手填单价的行改了才稳定保留。\r\n" +
                        "普通定额的单价是计算值，会被主程序重算覆盖，那种情况建议改「定额编号 × 系数」。";
                }
                else
                {
                    valueHintLabel.Text = "「改成」直接写死数值；「乘以 / 除以」把原式括起来再乘，例如 100 变成 (100)*0.85 。\r\n" +
                        "单个条目内的乘系数，主程序右键「乘系数」也能做；这里的价值是一次覆盖多个条目、多个单元。";
                }

                UpdateSentence();
            }

            private void RefreshRowAction()
            {
                bool isInsert = rowActionBox.Text == "新增定额";
                insertGridHeader.Visible = isInsert;
                insertGrid.Visible = isInsert;
                rowHintLabel.Text = isInsert
                    ? "在上面列出的每个条目下各新增一份，追加到条目末尾。\r\n新增不看「目标」那一栏。"
                    : "删除上面「目标」里列的那些行。\r\n目标为空时，会删掉范围内的全部定额行，注意别误删。";

                RefreshScopeAvailability();
                UpdateSentence();
            }

            private void RefreshCrossAction()
            {
                bool move = crossActionBox.Text == "移动到";
                crossHintLabel.Text = move
                    ? "移动不保留来源副本，目标只能有一个条目。\r\n来源条目只能有一个，请把上面的条目列表减到一条。"
                    : "复制会在目标条目下新增同样的定额，来源保持不变，目标条目可以选多个。\r\n来源条目只能有一个。";
                if (move && crossTargetItems.Count > 1)
                {
                    crossTargetItems.RemoveRange(1, crossTargetItems.Count - 1);
                    UpdateCrossTargetList();
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
                    throw new AgentPlanException("还没选条目：在左侧树点中条目后，点「添加条目」。");
                }

                return new List<string>(selectedItemNos);
            }

            // 目标：编号写进 QuotaFilter，精确名称写进 QuotaName，两者可以同时给，执行层取并集。
            private void ApplyTarget(AgentCommand command)
            {
                List<string> codes = TargetCodes();
                List<string> names = TargetNames();
                if (codes.Count > 0)
                {
                    command.QuotaFilter = codes;
                }

                if (names.Count > 0)
                {
                    command.QuotaName = String.Join("、", names.ToArray());
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

            // 从主程序定额表的选中行取定额编号，追加到网格里（数量留空，自己填）。
            private void FillGridCodesFromHostGrid(DataGridView grid)
            {
                List<string> codes = new List<string>();
                HashSet<string> seen = new HashSet<string>(StringComparer.Ordinal);
                int rowCount = 0;
                try
                {
                    DataGridView host = GetField<DataGridView>(mainForm, "dataGridViewDE");
                    if (host == null)
                    {
                        SetStatus("没有找到主程序的定额输入表格。", true);
                        return;
                    }

                    foreach (DataGridViewRow row in GetSelectedQuotaRows(host))
                    {
                        rowCount++;
                        string code = (GetRowValue(row, "定额编号DE", "定额编号") ?? "").Trim();
                        if (code.Length > 0 && seen.Add(code))
                        {
                            codes.Add(code);
                        }
                    }
                }
                catch (Exception ex)
                {
                    SetStatus("读取定额表选中行失败：" + ex.Message, true);
                    Log("Agent panel fill grid from host failed: " + ex);
                    return;
                }

                if (rowCount == 0)
                {
                    SetStatus("主程序定额表里没有选中行。请先在定额输入表里按住 Ctrl 或 Shift 多选几行。", true);
                    return;
                }

                if (codes.Count == 0)
                {
                    SetStatus("选中的行读不到定额编号。", true);
                    return;
                }

                HashSet<string> already = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (DataGridViewRow row in grid.Rows)
                {
                    if (row.IsNewRow)
                    {
                        continue;
                    }

                    string existing = Convert.ToString(row.Cells["Code"].Value ?? "").Trim();
                    if (existing.Length > 0)
                    {
                        already.Add(existing);
                    }
                }

                int added = 0;
                foreach (string code in codes)
                {
                    if (already.Contains(code))
                    {
                        continue;
                    }

                    int index = grid.Rows.Add();
                    grid.Rows[index].Cells["Code"].Value = code;
                    added++;
                }

                UpdateSentence();
                SetStatus("从定额表 " + rowCount.ToString(CultureInfo.InvariantCulture) + " 个选中行填入 " +
                    added.ToString(CultureInfo.InvariantCulture) + " 个编号，数量需要自己填。", false);
            }

            private static List<AgentQuotaInput> ReadQuotaGrid(DataGridView grid)
            {
                List<AgentQuotaInput> quotas = new List<AgentQuotaInput>();
                foreach (DataGridViewRow row in grid.Rows)
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
                    if (action == "改成")
                    {
                        command.Type = "set_quantity";
                        command.Value = RequireText(value, "数值");
                    }
                    else if (action == "清空数量")
                    {
                        command.Type = "clear_quantity";
                    }
                    else if ((action == "删除系数" || action == "删除调整内容"))
                    {
                        command.Type = "remove_text";
                        command.Target = "quantity";
                        command.RemoveText = RequireText(value, "要删掉的那段文字");
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
                    if (action == "改成")
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
                    if ((action == "删除系数" || action == "删除调整内容"))
                    {
                        command.Type = "remove_text";
                        command.Target = "quota_code";
                        command.RemoveText = RequireText(value, "要删掉的那段文字");
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
                    if ((action == "删除系数" || action == "删除调整内容"))
                    {
                        command.Type = "remove_text";
                        command.Target = "adjustment";
                        command.RemoveText = RequireText(value, "要删掉的那段调整内容");
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

            private AgentCommand BuildReplaceCommand()
            {
                AgentCommand command = new AgentCommand();
                ApplyScope(command);
                if (targetEntries.Count == 0)
                {
                    throw new AgentPlanException("替换定额必须先在上面「目标」里指定要被替换的是哪些，不能对整个范围替换。");
                }

                command.Type = "replace_quotas";
                command.FromCodes = TargetCodes();
                command.QuotaFilter = new List<string>();
                command.ToQuotas = ReadQuotaGrid(replaceGrid);
                if (command.ToQuotas.Count == 0)
                {
                    throw new AgentPlanException("请在「替换成」表格里填至少一条新定额。");
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
                    throw new AgentPlanException("一条拆成多条时不能选「所有单元」，请限定到一个单元。");
                }

                return command;
            }

            private AgentCommand BuildRowCommand()
            {
                AgentCommand command = new AgentCommand();
                if (rowActionBox.Text == "新增定额")
                {
                    command.Type = "insert_quotas";
                    command.IncludeChildren = includeChildrenBox.Checked;
                    command.Units = BuildUnitTokens();
                    command.Items = BuildItemTokens();
                    command.Quotas = ReadQuotaGrid(insertGrid);
                    if (command.Quotas.Count == 0)
                    {
                        throw new AgentPlanException("请填至少一条要新增的定额。");
                    }

                    return command;
                }

                command.Type = "delete_quotas";
                ApplyScope(command);
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
                    throw new AgentPlanException("复制 / 移动的来源只能是一个条目，请把上面的条目列表减到一条。");
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
                else if (tab == TabReplace)
                {
                    commands.Add(BuildReplaceCommand());
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
                    sentenceLabel.ForeColor = AgentPanelHintFore;
                    sentenceLabel.Text = "指令通道：写好后点「交给 AI」。";
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

                    sentenceLabel.ForeColor = AgentPanelTitleFore;
                    sentenceLabel.Text = "将要执行：" + String.Join("；", parts.ToArray());
                }
                catch (Exception ex)
                {
                    sentenceLabel.ForeColor = AgentPanelHintFore;
                    sentenceLabel.Text = "还差一步：" + ex.Message;
                }
            }

            // ===== 提交与预览 =====

            private AgentSelectionSnapshot CaptureAgentSelectionForPanel()
            {
                // 本面板不用"主程序选中行"当作用范围，快照里的行选中一律丢弃，
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
                            "可以改用上面的按钮页签，或输入确定性写法（点「指令帮助」看格式）。", true);
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

                    sentenceLabel.ForeColor = AgentPanelTitleFore;
                    sentenceLabel.Text = description.ToString();
                }

                SetStatus("已生成预览。" + elapsed, false);
                ShowPlanPreview(plan);
            }

            // 预览独立成窗口：以前内嵌在下方，一显示就把页签区挤扁，
            // 上一步填定额编号的表格会被压没且没法滚动。
            private void ShowPlanPreview(AgentPlan plan)
            {
                if (!ShowPlanPreviewDialog(plan))
                {
                    SetStatus("已取消，未执行任何修改。", false);
                    return;
                }

                if (plan.PreviewRows.Count > 200 && !ConfirmLargePlan(plan.PreviewRows.Count))
                {
                    SetStatus("已取消大批量执行。", false);
                    return;
                }

                ExecutePlanNow(plan);
            }

            private bool ShowPlanPreviewDialog(AgentPlan plan)
            {
                using (Form dialog = new Form())
                {
                    dialog.Text = "确认执行　—　" + plan.Summary;
                    dialog.StartPosition = FormStartPosition.CenterParent;
                    dialog.ClientSize = new Size(960, 560);
                    dialog.MinimumSize = new Size(700, 360);
                    dialog.ShowInTaskbar = false;
                    dialog.MinimizeBox = false;
                    dialog.Font = Font;

                    DataGridView grid = new DataGridView();
                    grid.Dock = DockStyle.Fill;
                    grid.ReadOnly = true;
                    grid.AllowUserToAddRows = false;
                    grid.AllowUserToDeleteRows = false;
                    grid.RowHeadersVisible = false;
                    grid.BackgroundColor = Color.White;
                    grid.BorderStyle = BorderStyle.None;
                    grid.EnableHeadersVisualStyles = false;
                    grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(238, 242, 248);
                    grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                    grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    grid.Columns.Add("Action", "操作");
                    grid.Columns.Add("Unit", "单元");
                    grid.Columns.Add("Item", "工程或费用项目名称");
                    grid.Columns.Add("Code", "定额编号");
                    grid.Columns.Add("Old", "原值");
                    grid.Columns.Add("New", "新值");
                    grid.Columns["Action"].FillWeight = 13;
                    grid.Columns["Unit"].FillWeight = 9;
                    grid.Columns["Item"].FillWeight = 28;
                    grid.Columns["Code"].FillWeight = 14;
                    grid.Columns["Old"].FillWeight = 18;
                    grid.Columns["New"].FillWeight = 18;

                    foreach (AgentPlanRow row in plan.PreviewRows.Take(2000))
                    {
                        grid.Rows.Add(
                            row.Action,
                            AgentUnitDisplay(plan.UnitCodes, row.UnitId),
                            !String.IsNullOrEmpty(row.ItemName) ? row.ItemName : (row.ItemNo ?? ""),
                            row.QuotaCode ?? "",
                            row.OldValue ?? "",
                            row.NewValue ?? "");
                    }

                    Label summary = new Label();
                    summary.Dock = DockStyle.Top;
                    summary.Height = 30;
                    summary.TextAlign = ContentAlignment.MiddleLeft;
                    summary.ForeColor = AgentPanelWarnFore;
                    summary.Font = new Font(Font, FontStyle.Bold);
                    summary.Padding = new Padding(10, 0, 0, 0);
                    summary.Text = plan.Summary +
                        (plan.PreviewRows.Count > 2000 ? "（只显示前 2000 行）" : "");

                    Label warnings = new Label();
                    warnings.Dock = DockStyle.Top;
                    warnings.ForeColor = AgentPanelHintFore;
                    warnings.Padding = new Padding(10, 0, 10, 4);
                    if (plan.Warnings.Count > 0)
                    {
                        warnings.Height = Math.Min(66, 18 * plan.Warnings.Count + 6);
                        warnings.Text = "注意：" + String.Join("　", plan.Warnings.ToArray());
                    }
                    else
                    {
                        warnings.Height = 0;
                        warnings.Visible = false;
                    }

                    Panel buttons = new Panel();
                    buttons.Dock = DockStyle.Bottom;
                    buttons.Height = 48;
                    buttons.Padding = new Padding(10, 9, 10, 9);

                    Button confirm = new Button();
                    confirm.Dock = DockStyle.Right;
                    confirm.Width = 124;
                    confirm.Text = "确认执行";
                    confirm.Font = new Font(Font, FontStyle.Bold);
                    confirm.BackColor = Color.FromArgb(216, 240, 220);
                    confirm.DialogResult = DialogResult.OK;

                    Panel gap = new Panel();
                    gap.Dock = DockStyle.Right;
                    gap.Width = 8;

                    Button cancel = new Button();
                    cancel.Dock = DockStyle.Right;
                    cancel.Width = 92;
                    cancel.Text = "取消";
                    cancel.DialogResult = DialogResult.Cancel;

                    Label rowCount = new Label();
                    rowCount.Dock = DockStyle.Fill;
                    rowCount.TextAlign = ContentAlignment.MiddleLeft;
                    rowCount.ForeColor = AgentPanelHintFore;
                    rowCount.Text = "共 " + plan.PreviewRows.Count.ToString(CultureInfo.InvariantCulture) + " 行";

                    buttons.Controls.Add(rowCount);
                    buttons.Controls.Add(cancel);
                    buttons.Controls.Add(gap);
                    buttons.Controls.Add(confirm);

                    dialog.Controls.Add(grid);
                    dialog.Controls.Add(warnings);
                    dialog.Controls.Add(summary);
                    dialog.Controls.Add(buttons);
                    dialog.AcceptButton = confirm;
                    dialog.CancelButton = cancel;

                    return dialog.ShowDialog(this) == DialogResult.OK;
                }
            }

            private void ExecutePlanNow(AgentPlan plan)
            {
                Enabled = false;
                try
                {
                    string message = ExecuteAgentPlan(mainForm, plan, delegate(string line) { });
                    SetStatus(message.Replace("\r\n", "　").Replace("\n", "　"), false);
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
                    dialog.ClientSize = new Size(380, 136);

                    label.Text = "本次将影响 " + rowCount.ToString(CultureInfo.InvariantCulture) + " 行数据。\r\n请输入「确认」两字后继续：";
                    label.SetBounds(14, 14, 350, 40);
                    box.SetBounds(14, 60, 350, 25);
                    ok.Text = "继续";
                    ok.SetBounds(196, 96, 80, 28);
                    ok.DialogResult = DialogResult.OK;
                    cancel.Text = "取消";
                    cancel.SetBounds(284, 96, 80, 28);
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
                statusLabel.ForeColor = isError ? AgentPanelErrorFore : Color.FromArgb(52, 58, 68);
                statusLabel.Text = text ?? "";
            }

            private void ShowTextDialog(string title, string text)
            {
                using (Form dialog = new Form())
                {
                    dialog.Text = title;
                    dialog.StartPosition = FormStartPosition.CenterParent;
                    dialog.ClientSize = new Size(760, 500);
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
                    dialog.Text = "文本指令帮助（「指令」页签用）";
                    dialog.StartPosition = FormStartPosition.CenterParent;
                    dialog.ClientSize = new Size(860, 580);
                    dialog.ShowInTaskbar = false;
                    dialog.MinimizeBox = false;

                    DataGridView grid = new DataGridView();
                    grid.Dock = DockStyle.Fill;
                    grid.ReadOnly = true;
                    grid.AllowUserToAddRows = false;
                    grid.AllowUserToDeleteRows = false;
                    grid.RowHeadersVisible = false;
                    grid.BackgroundColor = Color.White;
                    grid.EnableHeadersVisualStyles = false;
                    grid.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(238, 242, 248);
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
                    "单个条目内改数量/编号/单价/调整 → 主程序右键「乘系数」；跨条目跨单元批量 → 本窗口上面的按钮页签。",
                    "本页只说明「指令」页签的文本写法。",
                    "按钮页签拼不出来的场景（新建单元、运输方案、材料价方案）才需要文本或 AI。");
                AddHelpRow(grid, "自然语言（需AI）",
                    "直接用一句话描述要改什么，助手会先生成预览，确认后才执行。",
                    "把0101-01条目的定额数量乘0.85\r\n把南江路泵房单元0308-01的运输方案设为3\r\n照着_ZGS_02再建一个单元，叫测算二版",
                    "需要已配置 RecoQuotaData/deepseek-settings.json。AI 无法唯一判断条目或单元时，会要求补充。");
                AddHelpRow(grid, "作用范围",
                    "不写条目编号=当前选中的条目或定额；不写定额过滤=该条目下全部定额；单元=xxx 可限定单元。",
                    "工程数量 *0.85\r\n工程数量 0101-01 *0.85\r\n工程数量 0101-01、0102-01 *0.85\r\n删除数量 0308、0309 *0 单元=所有",
                    "多个条目或单元用顿号「、」隔开（英文逗号也兼容）；分号「；」只用于分隔多条完整命令。");
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
                    "文本通道的「替换定额」只支持一对一；一对多、多对一请用「替换定额」页签。");
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
