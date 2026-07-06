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
        private static readonly Dictionary<Form, AgentChatWindow> AgentChatWindows = new Dictionary<Form, AgentChatWindow>();
        private static readonly HashSet<Control> AgentShortcutHookedControls = new HashSet<Control>();
        // 记录最近一次 Ctrl+Q 是从章节树还是定额表进入的：树=意图整个条目（忽略定额表里顺带的当前行）。
        private static bool s_agentInvokeFromTree;

        private static void EnsureAgentChatRuntime(Form mainForm)
        {
            HookAgentShortcut(mainForm, GetField<DataGridView>(mainForm, "dataGridViewDE"));
            HookAgentShortcut(mainForm, GetField<TreeView>(mainForm, "Tv_tree"));
        }

        private static void HookAgentShortcut(Form mainForm, Control control)
        {
            if (control == null || AgentShortcutHookedControls.Contains(control))
            {
                return;
            }

            AgentShortcutHookedControls.Add(control);
            control.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.Control && !e.Shift && e.KeyCode == Keys.Q)
                {
                    s_agentInvokeFromTree = control is TreeView;
                    ShowAgentChatWindow(mainForm);
                    e.Handled = true;
                }
            };
        }

        // 智能指令改为仅 Ctrl+Q 进入，不再加右键菜单项。
        // 此方法只负责清掉历史版本可能已加入的菜单项。
        private static void AddAgentChatItemIfMatched(ContextMenuStrip menu)
        {
            if (menu == null)
            {
                return;
            }

            ToolStripMenuItem existing = FindMenuItem(menu, "智能指令(聊天)");
            if (existing != null)
            {
                menu.Items.Remove(existing);
                existing.Dispose();
            }
        }

        private static void ShowAgentChatWindow(Form mainForm)
        {
            AgentChatWindow window;
            if (!AgentChatWindows.TryGetValue(mainForm, out window) || window.IsDisposed)
            {
                window = new AgentChatWindow(mainForm);
                AgentChatWindows[mainForm] = window;
                mainForm.FormClosed += delegate
                {
                    AgentChatWindows.Remove(mainForm);
                };
            }

            if (!window.Visible)
            {
                window.Show(mainForm);
            }

            window.BringToFront();
            window.FocusInput();
        }

        private sealed class AgentChatWindow : Form
        {
            private readonly Form mainForm;
            private readonly RichTextBox transcript;
            private readonly Panel helpPanel;
            private readonly DataGridView helpGrid;
            private readonly Panel previewPanel;
            private readonly Label summaryLabel;
            private readonly DataGridView previewGrid;
            private readonly Button confirmButton;
            private readonly Button cancelButton;
            private readonly TextBox inputBox;
            private readonly Button sendButton;
            private AgentPlan pendingPlan;
            private bool parsing;

            public AgentChatWindow(Form mainForm)
            {
                this.mainForm = mainForm;
                Text = "智能指令助手 (Ctrl+Q)";
                StartPosition = FormStartPosition.Manual;
                Size = new Size(700, 720);
                MinimumSize = new Size(540, 480);
                ShowInTaskbar = false;
                try
                {
                    Location = new Point(
                        Math.Max(0, mainForm.Right - Width - 24),
                        Math.Max(0, mainForm.Top + 80));
                }
                catch
                {
                    StartPosition = FormStartPosition.CenterParent;
                }

                Panel inputPanel = new Panel();
                inputPanel.Dock = DockStyle.Bottom;
                inputPanel.Height = 92;
                inputPanel.Padding = new Padding(8, 4, 8, 8);

                Panel toolRow = new Panel();
                toolRow.Dock = DockStyle.Top;
                toolRow.Height = 30;

                Button insertItemButton = new Button();
                insertItemButton.Text = "插入当前条目";
                insertItemButton.Width = 105;
                insertItemButton.Height = 26;
                insertItemButton.Top = 2;
                insertItemButton.Left = 0;
                insertItemButton.Click += delegate { InsertCurrentTreeItem(); };

                Label pickHint = new Label();
                pickHint.Text = "（先在左侧树点中条目，再点此按钮把编号填入）";
                pickHint.AutoSize = true;
                pickHint.Top = 7;
                pickHint.Left = 112;
                pickHint.ForeColor = Color.Gray;

                Button helpButton = new Button();
                helpButton.Text = "帮助";
                helpButton.Width = 60;
                helpButton.Height = 26;
                helpButton.Top = 2;
                helpButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                helpButton.Click += delegate { ShowHelp(); };
                toolRow.Resize += delegate { helpButton.Left = toolRow.Width - helpButton.Width - 4; };

                toolRow.Controls.Add(insertItemButton);
                toolRow.Controls.Add(pickHint);
                toolRow.Controls.Add(helpButton);

                Panel sendRow = new Panel();
                sendRow.Dock = DockStyle.Fill;

                sendButton = new Button();
                sendButton.Text = "发送";
                sendButton.Width = 76;
                sendButton.Dock = DockStyle.Right;
                sendButton.Click += delegate { SubmitInput(); };

                inputBox = new TextBox();
                inputBox.Dock = DockStyle.Fill;
                inputBox.Font = new Font(Font.FontFamily, 10.5f);
                inputBox.KeyDown += delegate(object sender, KeyEventArgs e)
                {
                    if (e.KeyCode == Keys.Enter)
                    {
                        SubmitInput();
                        e.Handled = true;
                        e.SuppressKeyPress = true;
                    }
                };

                sendRow.Controls.Add(inputBox);
                sendRow.Controls.Add(sendButton);

                inputPanel.Controls.Add(sendRow);
                inputPanel.Controls.Add(toolRow);

                previewPanel = new Panel();
                previewPanel.Dock = DockStyle.Bottom;
                previewPanel.Height = 260;
                previewPanel.Visible = false;
                previewPanel.Padding = new Padding(8, 4, 8, 4);

                summaryLabel = new Label();
                summaryLabel.Dock = DockStyle.Top;
                summaryLabel.Height = 34;
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

                Panel buttonPanel = new Panel();
                buttonPanel.Dock = DockStyle.Bottom;
                buttonPanel.Height = 40;

                confirmButton = new Button();
                confirmButton.Text = "确认执行";
                confirmButton.Width = 110;
                confirmButton.Height = 30;
                confirmButton.Top = 5;
                confirmButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                confirmButton.BackColor = Color.FromArgb(220, 240, 220);
                confirmButton.Click += delegate { ConfirmPlan(); };

                cancelButton = new Button();
                cancelButton.Text = "取消";
                cancelButton.Width = 80;
                cancelButton.Height = 30;
                cancelButton.Top = 5;
                cancelButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                cancelButton.Click += delegate { CancelPlan("已取消，未执行任何修改。"); };

                buttonPanel.Controls.Add(confirmButton);
                buttonPanel.Controls.Add(cancelButton);
                buttonPanel.Resize += delegate
                {
                    confirmButton.Left = buttonPanel.Width - confirmButton.Width - cancelButton.Width - 24;
                    cancelButton.Left = buttonPanel.Width - cancelButton.Width - 12;
                };

                previewPanel.Controls.Add(previewGrid);
                previewPanel.Controls.Add(summaryLabel);
                previewPanel.Controls.Add(buttonPanel);

                transcript = new RichTextBox();
                transcript.Dock = DockStyle.Fill;
                transcript.ReadOnly = true;
                transcript.BackColor = Color.White;
                transcript.BorderStyle = BorderStyle.None;
                transcript.Font = new Font(Font.FontFamily, 10f);

                helpPanel = new Panel();
                helpPanel.Dock = DockStyle.Fill;
                helpPanel.Visible = false;
                helpPanel.Padding = new Padding(8, 4, 8, 4);

                Panel helpTop = new Panel();
                helpTop.Dock = DockStyle.Top;
                helpTop.Height = 34;

                Label helpTitle = new Label();
                helpTitle.Text = "帮助内容";
                helpTitle.Dock = DockStyle.Fill;
                helpTitle.TextAlign = ContentAlignment.MiddleLeft;
                helpTitle.ForeColor = Color.FromArgb(60, 60, 60);

                Button closeHelpButton = new Button();
                closeHelpButton.Text = "关闭帮助";
                closeHelpButton.Width = 82;
                closeHelpButton.Dock = DockStyle.Right;
                closeHelpButton.Click += delegate { HideHelp(); };

                helpTop.Controls.Add(closeHelpButton);
                helpTop.Controls.Add(helpTitle);

                helpGrid = new DataGridView();
                helpGrid.Dock = DockStyle.Fill;
                helpGrid.ReadOnly = true;
                helpGrid.AllowUserToAddRows = false;
                helpGrid.AllowUserToDeleteRows = false;
                helpGrid.AllowUserToResizeRows = true;
                helpGrid.RowHeadersVisible = false;
                helpGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                helpGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                helpGrid.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                helpGrid.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                helpGrid.DefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft;
                helpGrid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                helpGrid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
                helpGrid.Columns.Add("Scene", "场景");
                helpGrid.Columns.Add("Format", "写法");
                helpGrid.Columns.Add("Example", "实例");
                helpGrid.Columns.Add("Note", "说明");
                helpGrid.Columns["Scene"].FillWeight = 13;
                helpGrid.Columns["Format"].FillWeight = 25;
                helpGrid.Columns["Example"].FillWeight = 31;
                helpGrid.Columns["Note"].FillWeight = 31;

                helpPanel.Controls.Add(helpGrid);
                helpPanel.Controls.Add(helpTop);

                Controls.Add(transcript);
                Controls.Add(helpPanel);
                Controls.Add(previewPanel);
                Controls.Add(inputPanel);

                FormClosing += delegate(object sender, FormClosingEventArgs e)
                {
                    if (e.CloseReason == CloseReason.UserClosing)
                    {
                        e.Cancel = true;
                        Hide();
                    }
                };

                AppendSystem("我是聊天指令助手：用一句话描述要做的操作，我会先列出受影响的数据让你确认。输入\"帮助\"看示例。");
                AppendSystem("提示：在左侧树点中条目后，点\"插入当前条目\"按钮可把编号填进输入框；同名条目存在于每个单元，默认只改当前单元。");
            }

            public void FocusInput()
            {
                try
                {
                    inputBox.Focus();
                }
                catch
                {
                }
            }

            private void InsertCurrentTreeItem()
            {
                try
                {
                    TreeView tree = GetField<TreeView>(mainForm, "Tv_tree");
                    TreeNode node = tree != null ? tree.SelectedNode : GetField<TreeNode>(mainForm, "CurrNode");
                    if (node == null)
                    {
                        AppendError("当前没有选中树节点。");
                        return;
                    }

                    SqlConnection hostConn = GetProjectConnection(mainForm);
                    string itemNo = hostConn == null ? null : ResolveChapterNo(mainForm, hostConn, node);
                    if (String.IsNullOrEmpty(itemNo))
                    {
                        AppendError("无法识别当前条目编号。");
                        return;
                    }

                    InsertIntoInput(itemNo);
                }
                catch (Exception ex)
                {
                    AppendError("读取当前条目失败：" + ex.Message);
                }
            }

            private void InsertIntoInput(string itemNo)
            {
                string current = inputBox.Text ?? "";
                int pos = inputBox.SelectionStart;
                if (pos < 0 || pos > current.Length)
                {
                    pos = current.Length;
                }

                string prefix = current.Substring(0, pos);
                string insert = itemNo;
                if (prefix.Length > 0 && !prefix.EndsWith(" ", StringComparison.Ordinal) &&
                    !prefix.EndsWith(",", StringComparison.Ordinal) && !prefix.EndsWith("，", StringComparison.Ordinal))
                {
                    char last = prefix[prefix.Length - 1];
                    bool lastIsNumberish = Char.IsDigit(last) || last == '-' || last == '.';
                    insert = (lastIsNumberish ? "," : "") + itemNo;
                }

                inputBox.Text = prefix + insert + current.Substring(pos);
                inputBox.SelectionStart = (prefix + insert).Length;
                FocusInput();
            }

            private void AppendLine(string prefix, string text, Color color)
            {
                transcript.SelectionStart = transcript.TextLength;
                transcript.SelectionLength = 0;
                transcript.SelectionColor = color;
                transcript.AppendText(prefix + text + Environment.NewLine);
                transcript.SelectionStart = transcript.TextLength;
                transcript.ScrollToCaret();
            }

            private void AppendUser(string text)
            {
                AppendLine("你> ", text, Color.FromArgb(20, 60, 160));
            }

            private void AppendSystem(string text)
            {
                AppendLine("助手> ", text, Color.FromArgb(60, 60, 60));
            }

            private void AppendSuccess(string text)
            {
                AppendLine("助手> ", text, Color.FromArgb(0, 130, 0));
            }

            private void AppendError(string text)
            {
                AppendLine("助手> ", text, Color.FromArgb(190, 30, 30));
            }

            private void SubmitInput()
            {
                string text = (inputBox.Text ?? "").Trim();
                if (text.Length == 0)
                {
                    return;
                }

                inputBox.Text = "";
                HandleUserInput(text);
            }

            private void HandleUserInput(string text)
            {
                AppendUser(text);
                if (parsing)
                {
                    AppendError("上一条指令还在处理中，请稍候。");
                    return;
                }

                if (pendingPlan != null)
                {
                    CancelPlan("已放弃之前未确认的计划。");
                }

                string normalized = NormalizeAgentInput(text).TrimStart('/');
                if (normalized == "帮助" || String.Equals(normalized, "help", StringComparison.OrdinalIgnoreCase) || normalized == "?")
                {
                    ShowHelp();
                    return;
                }

                HideHelp();

                if (normalized == "撤销" || normalized == "撤回" || normalized == "撤掉上一步" || normalized == "撤掉")
                {
                    try
                    {
                        ShowPlanPreview(BuildAgentUndoPlan(mainForm));
                    }
                    catch (AgentPlanException ex)
                    {
                        AppendError(ex.Message);
                    }
                    catch (Exception ex)
                    {
                        AppendError("撤销准备失败：" + ex.Message);
                        Log("Agent undo preview failed: " + ex);
                    }

                    return;
                }

                if (normalized == "重做" || normalized == "恢复")
                {
                    try
                    {
                        ShowPlanPreview(BuildAgentRedoPlan(mainForm));
                    }
                    catch (AgentPlanException ex)
                    {
                        AppendError(ex.Message);
                    }
                    catch (Exception ex)
                    {
                        AppendError("重做准备失败：" + ex.Message);
                        Log("Agent redo preview failed: " + ex);
                    }

                    return;
                }

                if (normalized.StartsWith("探查", StringComparison.Ordinal))
                {
                    RunAgentDiagnostics(mainForm, normalized.Substring(2), AppendSystem);
                    return;
                }

                List<AgentCommand> commands = null;
                AgentParseResult fallback;
                if (TryParseAgentChain(text, out fallback))
                {
                    if (!String.IsNullOrEmpty(fallback.Error))
                    {
                        AppendError(fallback.Error);
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
                        AppendError("没有可用的 AI 配置（RecoQuotaData/deepseek-settings.json 需启用并填写 api_key），" +
                            "自然语言指令暂不可用。可以用确定性语法，输入\"帮助\"查看格式。");
                        return;
                    }
                }

                AgentSelectionSnapshot snapshot = CaptureAgentSelection(mainForm);
                RunAgentPipeline(text, commands, settings, snapshot);
            }

            // 后台流水线：独立连接 -> (可选)LLM解析 -> 生成预览计划 -> 回UI线程展示。
            private void RunAgentPipeline(string text, List<AgentCommand> preParsed, DeepSeekExcelMatchSettings settings, AgentSelectionSnapshot snapshot)
            {
                parsing = true;
                AppendSystem(preParsed != null ? "正在生成预览…" : "AI 解析中…");
                Stopwatch watch = Stopwatch.StartNew();
                Thread worker = new Thread(delegate()
                {
                    List<AgentCommand> commands = preParsed;
                    AgentParseResult llmResult = null;
                    AgentPlan plan = null;
                    string error = null;
                    try
                    {
                        using (SqlConnection conn = AgentCreateWorkConnection(mainForm))
                        {
                            if (commands == null)
                            {
                                AgentContext context = CollectAgentContext(conn, snapshot, text);
                                llmResult = RequestAgentParse(settings, context, text);
                                if (!String.IsNullOrEmpty(llmResult.Error))
                                {
                                    error = llmResult.Error;
                                }
                                else if (llmResult.Commands.Count == 0)
                                {
                                    // 仅澄清，无命令
                                }
                                else
                                {
                                    commands = llmResult.Commands;
                                }
                            }

                            if (error == null && commands != null && commands.Count > 0)
                            {
                                plan = BuildAgentPlan(conn, snapshot, commands);
                            }
                        }
                    }
                    catch (AgentPlanException ex)
                    {
                        error = ex.Message;
                    }
                    catch (Exception ex)
                    {
                        error = "处理失败：" + ex.Message;
                        Log("Agent pipeline failed: " + ex);
                    }

                    watch.Stop();
                    try
                    {
                        BeginInvoke((MethodInvoker)delegate
                        {
                            parsing = false;
                            OnPipelineDone(llmResult, plan, error, watch.Elapsed.TotalSeconds);
                        });
                    }
                    catch
                    {
                        parsing = false;
                    }
                });
                worker.IsBackground = true;
                worker.Start();
            }

            private void OnPipelineDone(AgentParseResult llmResult, AgentPlan plan, string error, double seconds)
            {
                string elapsed = "（耗时 " + seconds.ToString("0.0", CultureInfo.InvariantCulture) + " 秒）";
                if (!String.IsNullOrEmpty(error))
                {
                    AppendError(error + " " + elapsed);
                    return;
                }

                if (llmResult != null && llmResult.Commands.Count == 0)
                {
                    AppendSystem((String.IsNullOrEmpty(llmResult.Clarification)
                        ? "没有解析出可执行的命令。可以补充条目编号后重试，或输入\"帮助\"。"
                        : llmResult.Clarification) + " " + elapsed);
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

                    AppendSystem(description.ToString());
                }

                if (plan == null)
                {
                    AppendSystem("没有生成执行计划。" + elapsed);
                    return;
                }

                foreach (string warning in plan.Warnings)
                {
                    AppendSystem("注意：" + warning);
                }

                if (plan.PreviewRows.Count == 0)
                {
                    AppendSystem("没有匹配到任何数据行，未生成执行计划。" + elapsed);
                    return;
                }

                ShowPlanPreview(plan);
                AppendSystem(elapsed);
            }

            private void ShowPlanPreview(AgentPlan plan)
            {
                pendingPlan = plan;
                HideHelp();
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
                summaryLabel.Text = plan.Summary + extra;
                previewPanel.Visible = true;
                AppendSystem("请核对上方预览（" + plan.Summary + "），点\"确认执行\"生效，或\"取消\"。");
            }

            private void CancelPlan(string message)
            {
                pendingPlan = null;
                previewPanel.Visible = false;
                previewGrid.Rows.Clear();
                if (!String.IsNullOrEmpty(message))
                {
                    AppendSystem(message);
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
                    AppendSystem("已取消大批量执行。");
                    return;
                }

                pendingPlan = null;
                previewPanel.Visible = false;
                Enabled = false;
                try
                {
                    string message = ExecuteAgentPlan(mainForm, plan, AppendSystem);
                    AppendSuccess(message);
                }
                catch (AgentPlanException ex)
                {
                    AppendError(ex.Message);
                }
                catch (Exception ex)
                {
                    AppendError("执行失败：" + ex.Message);
                    Log("Agent execute failed: " + ex);
                }
                finally
                {
                    Enabled = true;
                    previewGrid.Rows.Clear();
                    FocusInput();
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

            private void ShowHelp()
            {
                if (helpGrid.Rows.Count == 0)
                {
                    PopulateHelpGrid(helpGrid);
                }

                transcript.Visible = false;
                helpPanel.Visible = true;
                helpGrid.Focus();
            }

            private void HideHelp()
            {
                helpPanel.Visible = false;
                transcript.Visible = true;
                FocusInput();
            }

            private static void PopulateHelpGrid(DataGridView grid)
            {
                grid.Rows.Clear();
                AddHelpRow(grid, "自然语言（需AI）",
                    "直接用一句话描述要改什么，助手会先生成预览，确认后才执行。",
                    "把0101-01条目的定额数量乘0.85\r\n把南江路泵房单元0308-01的运输方案设为3\r\n照着_ZGS_02再建一个单元，叫测算二版",
                    "需要已配置 RecoQuotaData/deepseek-settings.json。AI 无法唯一判断条目或单元时，会要求补充。");
                AddHelpRow(grid, "作用范围",
                    "不写条目编号=当前选中的条目或定额；不写定额过滤=该条目下全部定额；单元=xxx 可限定单元。",
                    "工程数量 *0.85\r\n工程数量 0101-01 *0.85\r\n工程数量 0101-01 LY-21 *0.85 单元=南江路泵房",
                    "同一个条目编号在每个单元都存在。不写单元时默认只改当前单元；要跨单元请写 所有单元 或 单元=具体名称。");
                AddHelpRow(grid, "编号判别",
                    "带横杠且纯数字的编号当作条目；含字母或 SH/SQ/ZLF/TLF/LF 等当作定额或费用代码。",
                    "0101-01 是条目\r\nLY-21 是定额\r\nSH 是填单价代码",
                    "左侧树选中条目后，可点“插入当前条目”把条目编号填入输入框。");
                AddHelpRow(grid, "乘除工程数量",
                    "工程数量 [条目编号] [定额编号] *系数 或 /系数",
                    "工程数量 0101-01 LY-21 *0.85\r\n工程数量 0101-01 /2\r\n工程数量 *1.1",
                    "给定额编号时只改该定额；省略定额时改条目下符合当前选择规则的行。数量表达式会写入工程数量输入。");
                AddHelpRow(grid, "乘定额编号",
                    "定额编号 [条目编号] [定额编号] *系数 或 /系数",
                    "定额编号 0101-01 LY-21 *9\r\n定额编号 LY-21 /1.1",
                    "是在定额编号后追加乘除系数，适合软件原生缩放定额；改后通常需要在软件里手工重算。");
                AddHelpRow(grid, "改单价",
                    "单价 [条目编号] [定额编号] *系数 或 /系数",
                    "单价 0101-01 SH *1.05\r\n单价 SH /1.1",
                    "单价按两位小数预览；仅改当前计算值，长期保留通常不如用“定额编号 *系数”。");
                AddHelpRow(grid, "删除乘除片段",
                    "工程数量/定额编号 [条目编号] [定额编号] 删除*系数 或 删除/系数",
                    "工程数量 0101-01 LY-21 删除*0.85\r\n定额编号 LY-21 删除/1.1",
                    "只删除字段中已有的相同片段；字段里不存在该片段的行会跳过。单价不支持删除片段。");
                AddHelpRow(grid, "定额调整",
                    "定额调整 [条目编号] [定额编号] 调整内容；删除时在调整内容前写 删除。",
                    "定额调整 0101-01 LY-21 /XG1\r\n定额调整 0101-01 LY-21 删除 /XG1\r\n定额调整 LY-21 /1294861,,1",
                    "调整内容按整串写入或从原串删除。省略条目时作用于当前条目。");
                AddHelpRow(grid, "设数量/清空数量",
                    "设数量 [条目编号] [定额编号] 数量；清空数量 [条目编号] [定额编号]",
                    "设数量 0101-01 LY-21 100\r\n清空数量 0101-01 LY-21\r\n设数量 LY-21 25.5",
                    "设数量会同时更新工程数量输入和计算数量；清空数量会把数量置空。");
                AddHelpRow(grid, "替换/删除定额",
                    "替换定额 [条目编号] 原定额 新定额；删除定额 [条目编号] [定额编号]",
                    "替换定额 0101-01 LY-21 QY-100\r\n删除定额 0101-01 LY-21\r\n删除定额 LY-21",
                    "替换会保留原定额编号后面的乘除系数后缀；删除会连同相关计算缓存一起清理。");
                AddHelpRow(grid, "输入/复制定额",
                    "输入定额 [条目编号] 编号=数量；复制定额 来源条目 到 目标条目",
                    "输入定额 0101-01 LY-21=100\r\n输入定额 LY-21=100,QY-100=5\r\n复制定额 0101-01 到 0102-01",
                    "输入定额省略条目时写入当前选中条目；复制定额会把来源条目中命中的定额复制到目标条目。");
                AddHelpRow(grid, "运输方案",
                    "设运输方案 [条目编号] 方案序号 [运输参数] 单元=单元名",
                    "设运输方案 0101-01 4 PH0 单元=南江路泵房\r\n设运输方案 0308-01 3 单元=_ZGS_03",
                    "方案序号必须是数字。工具只设置方案号和可选参数，不现场生成材料运输方案定义。");
                AddHelpRow(grid, "材料/机械/设备/工费方案",
                    "改材料价 [材料|机械|设备|工费] 方案名称 单元=单元名",
                    "改材料价 材料 部颁25年4季度 单元=南江路泵房\r\n改材料价 机械 机械费方案A 单元=_ZGS_03",
                    "必须明确单元。修改方案后需要在软件里手工触发重算，相关费用才会更新。");
                AddHelpRow(grid, "新建单元",
                    "新建单元 新名称 从 源单元名称；源也可写 _ZGS_编号 或总概算序号。",
                    "新建单元 测算二版 从 _ZGS_02\r\n复制单元 比选方案 从 南江路泵房",
                    "会复制源单元的总概算条目、单项概算信息、定额输入等相关数据，并生成新的总概算编号。");
                AddHelpRow(grid, "多条链式执行",
                    "多条确定性指令用英文分号 ; 隔开，一次生成预览。",
                    "工程数量 0101-01 *0.9 ; 定额编号 0102-01 *0.9 ; 设数量 0103-01 100",
                    "各段独立解析，不依赖前一段结果。任一段格式错误会中止整条链式指令。");
                AddHelpRow(grid, "撤销/重做/探查",
                    "撤销；重做；探查 关键词；帮助",
                    "撤销\r\n重做\r\n探查 当前选择",
                    "撤销只记录本次软件运行期间由智能指令助手执行的操作；删除、插入、新建单元不支持重做。");
            }

            private static void AddHelpRow(DataGridView grid, string scene, string format, string example, string note)
            {
                grid.Rows.Add(scene, format, example, note);
            }
        }
    }
}
