using System;
using System.Collections.Generic;
using System.Windows.Forms;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        private static readonly HashSet<Control> AgentShortcutHookedControls = new HashSet<Control>();
        // 记录最近一次采集选中快照时是否要忽略定额表里顺带选中的当前行。
        // 点选式面板会在采集前按自己的"定额来源"单选项设置这个标志，不再依赖进入方式。
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

        // 入口保持不变，实际打开的是点选式面板（AgentPanelFeature.cs）。
        private static void ShowAgentChatWindow(Form mainForm)
        {
            ShowAgentPanelWindow(mainForm);
        }
    }
}
