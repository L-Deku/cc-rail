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
                previewGrid.Columns.Add("Item", "条目编号");
                previewGrid.Columns.Add("Code", "定额编号");
                previewGrid.Columns.Add("Old", "原值");
                previewGrid.Columns.Add("New", "新值");
                previewGrid.Columns["Action"].FillWeight = 13;
                previewGrid.Columns["Unit"].FillWeight = 9;
                previewGrid.Columns["Item"].FillWeight = 20;
                previewGrid.Columns["Code"].FillWeight = 16;
                previewGrid.Columns["Old"].FillWeight = 21;
                previewGrid.Columns["New"].FillWeight = 21;

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

                Controls.Add(transcript);
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
                previewGrid.Rows.Clear();
                foreach (AgentPlanRow row in plan.PreviewRows.Take(2000))
                {
                    previewGrid.Rows.Add(
                        row.Action,
                        row.UnitId > 0 ? row.UnitId.ToString(CultureInfo.InvariantCulture) : "",
                        row.ItemNo ?? "",
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
                AppendSystem("直接用自然语言描述操作（需配置AI），例如：");
                AppendSystem("  把0101-01条目的定额数量乘0.85 / 把定额编号乘0.9 / 把单价乘1.1");
                AppendSystem("  把0101-01的数量设为100 / 把0101-01里的LY-21全换成QY-100");
                AppendSystem("  删除当前条目里LY-21 / 在0101-01下输入LY-21数量100");
                AppendSystem("  把南江路泵房单元0308-01的运输方案设为3 / 把该单元材料费方案换成部颁25年4季度");
                AppendSystem("  照着_ZGS_02再建一个单元，叫测算二版");
                AppendSystem("不用AI的确定性语法（单元=xxx 可选）。条目/定额规则：给定额编号=该条目下全部该定额；不给=当前选中的定额行；不给条目编号=当前条目：");
                AppendSystem("  工程数量 / 定额编号 / 单价 / 定额调整  [条目编号] [定额编号] 操作 —— 单个编号自动判别：带-纯数字(0101-01)=条目；材料编号无-；含字母(LY-21/ZLF/TLF/SH/SQ)=定额");
                AppendSystem("    工程数量/定额编号/单价 的操作为 *系数 / /系数（乘除）或 删除*系数 / 删除/系数（去掉）");
                AppendSystem("    例：工程数量 0101-01 LY-21 *0.85 / 定额编号 0101-01 LY-21 *9 / 单价 LY-21 /1.1 / 工程数量 删除*0.85（当前选中定额）");
                AppendSystem("    （定额编号字段不管装定额/材料编号/ZLF，操作都只在其后追加或删除乘除；单价为计算值不支持删除）");
                AppendSystem("  定额调整 0101-01 LY-21 /XG1（整串写入，原样）/ 定额调整 0101-01 LY-21 删除 /XG1（去掉该调整）/ 定额调整 LY-21 /1294861,,1（当前条目该定额）");
                AppendSystem("  设数量 0101-01 LY-21 100 / 清空数量 0101-01 LY-21 / 删除定额 0101-01 LY-21（条目/定额规则同上：省略定额=选中定额，省略条目=当前条目）");
                AppendSystem("  替换定额 0101-01 LY-21 QY-100 / 复制定额 0101-01 到 0102-01 / 输入定额 [0101-01] LY-21=100（省略条目编号则输入到当前选中条目）");
                AppendSystem("  设运输方案 0101-01 4 PH0 单元=南江路泵房（PH0为可选运输参数）/ 改材料价 材料 部颁25年4季度 单元=南江路泵房");
                AppendSystem("  新建单元 新名称 从 源单元名称（源也可用 _ZGS_02 或总概算序号）");
                AppendSystem("  分号链式：多条指令用 ; 隔开一次预览执行，例：工程数量 0101-01 *0.9 ; 定额编号 0102-01 *0.9 ; 设数量 0103-01 100（各段独立，互不依赖前一段结果）");
                AppendSystem("其他命令：撤销(撤回/撤掉上一步) / 重做(恢复刚撤销的) / 探查(诊断) / 帮助");
                AppendSystem("提示：用\"定额编号 ×系数\"(追加*系数,软件原生缩放,重算后有效)优于直接\"单价 ×系数\"(仅补充定额长期有效)。");
                AppendSystem("省事：不写条目编号=作用于当前选中的条目/定额；不写[定额过滤]=该条目下全部定额(若有选中定额行则只作用于选中的)。");
                AppendSystem("  例：选中几行定额后直接发 \"工程数量 0.85\" 就只改那几行；什么都不选则改当前条目全部。");
                AppendSystem("重要：同一条目编号在每个单元里都存在！不点名单元时默认只改当前单元；要全部单元请明说\"所有单元\"。");
                AppendSystem("条目编号：在左侧树点中后按\"插入当前条目\"填入。所有修改先预览确认，可\"撤销\"。改编号/单价/方案后需在软件里手工重算。");
            }
        }
    }
}
