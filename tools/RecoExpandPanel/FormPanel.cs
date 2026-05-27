using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.IO.Compression;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Serialization;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace RecoNet
{
    public class FormPanel : Form
    {
        private static readonly HashSet<Form> InstalledForms = new HashSet<Form>();
        private static readonly HashSet<TreeView> HookedTrees = new HashSet<TreeView>();
        private static readonly Dictionary<ContextMenuStrip, MenuInfo> MenuInfos = new Dictionary<ContextMenuStrip, MenuInfo>();
        private static readonly Dictionary<ContextMenu, MenuInfo> LegacyMenuInfos = new Dictionary<ContextMenu, MenuInfo>();
        private static readonly Dictionary<Form, NativeTreeMenuFilter> NativeTreeMenuFilters = new Dictionary<Form, NativeTreeMenuFilter>();
        private static readonly Dictionary<Form, ExcelLinkRuntime> ExcelLinkRuntimes = new Dictionary<Form, ExcelLinkRuntime>();
        private static readonly Dictionary<Form, ExcelLinkPanel> ExcelLinkPanels = new Dictionary<Form, ExcelLinkPanel>();
        private static readonly Dictionary<Form, QuickBindPanel> QuickBindPanels = new Dictionary<Form, QuickBindPanel>();
        private static readonly Dictionary<string, System.Drawing.Image> MenuIconCache = new Dictionary<string, System.Drawing.Image>(StringComparer.OrdinalIgnoreCase);
        private static readonly HashSet<DataGridView> HookedQuotaGrids = new HashSet<DataGridView>();
        private static bool idleHooked;
        private readonly Timer installTimer;
        private readonly object owner;
        private Form installedMainForm;

        private sealed class MenuInfo
        {
            public Form MainForm;
            public string Name;
        }

        private sealed class QuotaKey
        {
            public string TotalNo;
            public string ChapterSeq;
            public string OrderNo;

            public string Key
            {
                get { return TotalNo + "|" + ChapterSeq + "|" + OrderNo; }
            }
        }

        private sealed class FactorInfo
        {
            public string Operator;
            public string Factor;

            public string Suffix
            {
                get { return Operator + Factor; }
            }
        }

        public static void InstallOnIdle()
        {
            if (idleHooked)
            {
                return;
            }

            idleHooked = true;
            Log("InstallOnIdle registered.");
            Application.Idle += delegate
            {
                try
                {
                    Form mainForm = FindMainForm();
                    if (mainForm != null && !InstalledForms.Contains(mainForm))
                    {
                        new FormPanel(mainForm);
                    }
                }
                catch (Exception ex)
                {
                    Log("Idle install failed: " + ex);
                }
            };
        }

        public FormPanel()
            : this(null)
        {
        }

        public FormPanel(Form owner)
            : this((object)owner)
        {
        }

        public FormPanel(object owner)
        {
            this.owner = owner;
            Log("FormPanel constructed. Owner=" + (owner == null ? "<null>" : owner.GetType().FullName));
            ShowInTaskbar = false;
            WindowState = FormWindowState.Minimized;
            Opacity = 0;

            installTimer = new Timer();
            installTimer.Interval = 800;
            installTimer.Tick += delegate { TryInstall(); };
            installTimer.Start();
            TryInstall();
        }

        protected override void SetVisibleCore(bool value)
        {
            base.SetVisibleCore(false);
        }

        private void TryInstall()
        {
            try
            {
                Form mainForm = owner as Form ?? FindMainForm();
                if (mainForm == null || mainForm.GetType().FullName != "RecoNet.RecoMainForm")
                {
                    Log("Main form not ready.");
                    return;
                }

                if (InstalledForms.Contains(mainForm))
                {
                    InstallAllContextMenus(mainForm);
                    InstallTreeMouseHook(mainForm);
                    InstallNativeTreeMenuFilter(mainForm);
                    return;
                }

                int menus = InstallAllContextMenus(mainForm);
                InstallTreeMouseHook(mainForm);
                InstallNativeTreeMenuFilter(mainForm);
                InstallQuotaGridShortcuts(mainForm);
                EnsureExcelLinkRuntime(mainForm);
                if (menus == 0)
                {
                    Log("Context menus not found.");
                    return;
                }

                InstalledForms.Add(mainForm);
                installedMainForm = mainForm;
                Log("Multiplier menu installed. menus=" + menus.ToString(CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                Log("Install failed: " + ex);
                // Keep the host application quiet; the next timer tick can retry.
            }
        }

        private static void InstallTreeMouseHook(Form mainForm)
        {
            TreeView tree = GetField<TreeView>(mainForm, "Tv_tree");
            if (tree == null || HookedTrees.Contains(tree))
            {
                return;
            }

            HookedTrees.Add(tree);
            if (tree.ContextMenu != null)
            {
                tree.ContextMenu.Popup -= LegacyContextMenuPopup;
                tree.ContextMenu.Popup += LegacyContextMenuPopup;
                LegacyMenuInfos[tree.ContextMenu] = new MenuInfo { MainForm = mainForm, Name = "MenuTree" };
                AddLegacyTreeMultiplierItem(tree.ContextMenu, mainForm);
            }
            if (tree.ContextMenuStrip != null)
            {
                tree.ContextMenuStrip.Opening -= ContextMenuOpening;
                tree.ContextMenuStrip.Opening += ContextMenuOpening;
                MenuInfos[tree.ContextMenuStrip] = new MenuInfo { MainForm = mainForm, Name = "MenuTree" };
                AddMultiplierItem(tree.ContextMenuStrip, mainForm, false, true);
            }
            tree.MouseUp += delegate(object sender, MouseEventArgs e)
            {
                if (e.Button != MouseButtons.Right)
                {
                    return;
                }

                TreeView tv = sender as TreeView;
                if (tv == null)
                {
                    return;
                }

                TreeNode node = tv.GetNodeAt(e.Location);
                if (node != null)
                {
                    tv.SelectedNode = node;
                }

                PatchVisibleToolStripTreeMenus(mainForm);
            };
            Log("Tree mouse hook installed.");
        }

        private static void TryPatchVisibleTreeMenu(Form mainForm, System.Drawing.Point screenPoint)
        {
            bool patched = false;
            foreach (ContextMenuStrip menu in MenuInfos.Keys.ToArray())
            {
                if (menu == null || !menu.Visible)
                {
                    continue;
                }

                if (!LooksLikeTreeMenu(menu))
                {
                    continue;
                }

                AddMultiplierItem(menu, mainForm, false, true);
                ForceMenuRelayout(menu, screenPoint);
                patched = true;
            }

            if (PatchVisibleToolStripTreeMenus(mainForm))
            {
                patched = true;
            }

            if (!patched)
            {
                Log("No visible tree menu found to patch.");
            }
        }

        private static void ForceMenuRelayout(ContextMenuStrip menu, System.Drawing.Point screenPoint)
        {
            try
            {
                menu.SuspendLayout();
                menu.ResumeLayout(true);
                menu.PerformLayout();
                menu.AutoSize = true;
            }
            catch (Exception ex)
            {
                Log("Menu relayout failed: " + ex.Message);
            }
        }

        private static int InstallAllContextMenus(Form mainForm)
        {
            int count = 0;
            foreach (FieldInfo field in mainForm.GetType().GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
            {
                ContextMenuStrip menu = field.GetValue(mainForm) as ContextMenuStrip;
                if (menu == null)
                {
                    continue;
                }

                count++;
                menu.Opening -= ContextMenuOpening;
                menu.Opening += ContextMenuOpening;
                MenuInfos[menu] = new MenuInfo { MainForm = mainForm, Name = field.Name };
                AddMultiplierItemIfMatched(menu);
            }

            foreach (FieldInfo field in mainForm.GetType().GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
            {
                ContextMenu menu = field.GetValue(mainForm) as ContextMenu;
                if (menu == null)
                {
                    continue;
                }

                count++;
                menu.Popup -= LegacyContextMenuPopup;
                menu.Popup += LegacyContextMenuPopup;
                LegacyMenuInfos[menu] = new MenuInfo { MainForm = mainForm, Name = field.Name };
                AddLegacyTreeMultiplierItemIfMatched(menu);
            }

            return count;
        }

        private static void LegacyContextMenuPopup(object sender, EventArgs e)
        {
            AddLegacyTreeMultiplierItemIfMatched(sender as ContextMenu);
        }

        private static void ContextMenuOpening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ContextMenuStrip menu = sender as ContextMenuStrip;
            AddMultiplierItemIfMatched(menu);
        }

        private static void AddMultiplierItemIfMatched(ContextMenuStrip menu)
        {
            if (menu == null || !MenuInfos.ContainsKey(menu))
            {
                return;
            }

            MenuInfo info = MenuInfos[menu];
            bool isDeMenu = info.Name == "contextMenuStripDE" || IsSource(menu, info.MainForm, "dataGridViewDE") || HasAnyItem(menu, "定额输入", "定额调整", "单价分析", "全选(A)");
            bool isTreeMenu = info.Name == "MenuTree" || IsSource(menu, info.MainForm, "Tv_tree") || HasAnyItem(menu, "计算参数设置", "删除条目", "整理清单编码", "复制数据(D)", "计算结果统计", "插入子级", "恢复章节", "删除单项概算标识");
            if (!isDeMenu && !isTreeMenu)
            {
                return;
            }

            AddMultiplierItem(menu, info.MainForm, isDeMenu, isTreeMenu);
        }

        private static void AddLegacyTreeMultiplierItemIfMatched(ContextMenu menu)
        {
            if (menu == null || !LegacyMenuInfos.ContainsKey(menu))
            {
                return;
            }

            MenuInfo info = LegacyMenuInfos[menu];
            bool isTreeMenu = info.Name == "MenuTree" || LooksLikeLegacyTreeMenu(menu);
            if (!isTreeMenu)
            {
                return;
            }

            AddLegacyTreeMultiplierItem(menu, info.MainForm);
        }

        private static void InstallNativeTreeMenuFilter(Form mainForm)
        {
            if (mainForm == null || NativeTreeMenuFilters.ContainsKey(mainForm))
            {
                return;
            }

            NativeTreeMenuFilter filter = new NativeTreeMenuFilter(mainForm);
            NativeTreeMenuFilters[mainForm] = filter;
            Application.AddMessageFilter(filter);
            mainForm.FormClosed += delegate
            {
                Application.RemoveMessageFilter(filter);
                NativeTreeMenuFilters.Remove(mainForm);
            };
            Log("Native tree menu filter installed.");
        }

        private static bool PatchVisibleToolStripTreeMenus(Form mainForm)
        {
            bool patched = false;
            EnumThreadWindows(GetCurrentThreadId(), delegate(IntPtr hWnd, IntPtr lParam)
            {
                if (!IsWindowVisible(hWnd))
                {
                    return true;
                }

                ToolStrip menu = Control.FromHandle(hWnd) as ToolStrip;
                if (menu == null || !LooksLikeTreeMenu(menu))
                {
                    return true;
                }

                AddMultiplierItem(menu, mainForm, false, true);
                try
                {
                    menu.PerformLayout();
                    menu.Refresh();
                }
                catch
                {
                }

                patched = true;
                Log("Visible ToolStrip tree menu patched. visible=" + VisibleItemText(menu));
                return true;
            }, IntPtr.Zero);

            return patched;
        }

        private static void EnsureExcelLinkRuntime(Form mainForm)
        {
            if (mainForm == null || ExcelLinkRuntimes.ContainsKey(mainForm))
            {
                return;
            }

            ExcelLinkRuntime runtime = new ExcelLinkRuntime(mainForm);
            ExcelLinkRuntimes[mainForm] = runtime;
            mainForm.FormClosed += delegate
            {
                runtime.Dispose();
                ExcelLinkRuntimes.Remove(mainForm);
                if (ExcelLinkPanels.ContainsKey(mainForm))
                {
                    ExcelLinkPanels.Remove(mainForm);
                }
                if (QuickBindPanels.ContainsKey(mainForm))
                {
                    QuickBindPanels.Remove(mainForm);
                }
            };
        }

        private static void InstallQuotaGridShortcuts(Form mainForm)
        {
            DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
            if (grid == null || HookedQuotaGrids.Contains(grid))
            {
                return;
            }

            HookedQuotaGrids.Add(grid);
            grid.KeyDown += delegate(object sender, KeyEventArgs e)
            {
                if (e.Control && e.KeyCode == Keys.E)
                {
                    ShowQuickBindPanel(mainForm);
                    e.Handled = true;
                }
            };
            Log("Excel quick bind shortcut installed.");
        }

        private static bool IsSource(ContextMenuStrip menu, Form mainForm, string fieldName)
        {
            Control source = menu.SourceControl;
            Control expected = GetField<Control>(mainForm, fieldName);
            return source != null && expected != null && Object.ReferenceEquals(source, expected);
        }

        private static bool HasAnyItem(ToolStrip menu, params string[] texts)
        {
            foreach (ToolStripItem item in menu.Items)
            {
                foreach (string text in texts)
                {
                    if (item.Text == text)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static Form FindMainForm()
        {
            foreach (Form form in Application.OpenForms)
            {
                if (form != null && form.GetType().FullName == "RecoNet.RecoMainForm")
                {
                    return form;
                }
            }

            return null;
        }

        private static void AddMultiplierItem(ToolStrip menu, Form mainForm, bool isDeMenu, bool isTreeMenu)
        {
            int insertIndex = FindInsertIndex(menu, isTreeMenu);
            if (insertIndex < 0)
            {
                insertIndex = FirstVisibleIndex(menu);
                if (insertIndex < 0)
                {
                    insertIndex = menu.Items.Count;
                }
                else
                {
                    insertIndex++;
                }
            }

            ToolStripMenuItem multiply = FindMenuItem(menu, "乘系数");
            if (multiply != null && multiply.DropDownItems.Count == 0)
            {
                int existingIndex = menu.Items.IndexOf(multiply);
                menu.Items.Remove(multiply);
                multiply.Dispose();
                multiply = null;
                if (existingIndex >= 0)
                {
                    insertIndex = existingIndex;
                }
            }

            if (multiply == null)
            {
                multiply = new ToolStripMenuItem("乘系数");
                multiply.Visible = true;
                multiply.Available = true;
                multiply.Enabled = true;
                menu.Items.Insert(Math.Min(insertIndex, menu.Items.Count), multiply);
            }
            multiply.Visible = true;
            multiply.Available = true;
            multiply.Enabled = true;
            ApplyMenuIcon(multiply, "multiply.png");
            ConfigureFactorTargetMenu(multiply, mainForm, isDeMenu, isTreeMenu);

            if (isDeMenu)
            {
                int baseIndex = menu.Items.IndexOf(multiply) + 1;
                ToolStripMenuItem bindExcel = FindMenuItem(menu, "绑定Excel工程量");
                if (bindExcel == null)
                {
                    bindExcel = new ToolStripMenuItem("绑定Excel工程量");
                    bindExcel.Visible = true;
                    bindExcel.Available = true;
                    bindExcel.Enabled = true;
                    bindExcel.Click += delegate { ShowQuickBindPanel(mainForm); };
                    menu.Items.Insert(Math.Min(baseIndex, menu.Items.Count), bindExcel);
                    baseIndex++;
                }
                ApplyMenuIcon(bindExcel, "excel_bind.png");

                ToolStripMenuItem openPanel = FindMenuItem(menu, "打开Excel联动面板");
                if (openPanel == null)
                {
                    openPanel = new ToolStripMenuItem("打开Excel联动面板");
                    openPanel.Visible = true;
                    openPanel.Available = true;
                    openPanel.Enabled = true;
                    openPanel.Click += delegate { ShowExcelLinkPanel(mainForm); };
                    menu.Items.Insert(Math.Min(baseIndex, menu.Items.Count), openPanel);
                }
                ApplyMenuIcon(openPanel, "excel_panel.png");
            }
        }

        private static void ConfigureFactorTargetMenu(ToolStripMenuItem root, Form mainForm, bool isDeMenu, bool isTreeMenu)
        {
            if (root == null)
            {
                return;
            }

            EnsureFactorTargetItem(root, mainForm, isDeMenu, isTreeMenu, "乘到原来的工程量", "quantity");
            EnsureFactorTargetItem(root, mainForm, isDeMenu, isTreeMenu, "乘到定额编号", "quotaCode");
        }

        private static void EnsureFactorTargetItem(ToolStripMenuItem parent, Form mainForm, bool isDeMenu, bool isTreeMenu, string text, string target)
        {
            if (FindDropDownMenuItem(parent, text) != null)
            {
                return;
            }

            AddFactorTargetItem(parent, mainForm, isDeMenu, isTreeMenu, text, target);
        }

        private static void AddFactorTargetItem(ToolStripMenuItem parent, Form mainForm, bool isDeMenu, bool isTreeMenu, string text, string target)
        {
            ToolStripMenuItem item = new ToolStripMenuItem(text);
            item.Click += delegate
            {
                if (isDeMenu)
                {
                    ApplyToSelectedQuotaRows(mainForm, target);
                }
                else if (isTreeMenu)
                {
                    ApplyToTree(mainForm, target);
                }
            };
            parent.DropDownItems.Add(item);
        }

        private static ToolStripMenuItem FindDropDownMenuItem(ToolStripMenuItem menu, string text)
        {
            foreach (ToolStripItem item in menu.DropDownItems)
            {
                ToolStripMenuItem menuItem = item as ToolStripMenuItem;
                if (menuItem != null && menuItem.Text == text)
                {
                    return menuItem;
                }
            }

            return null;
        }

        private static ToolStripMenuItem FindMenuItem(ToolStrip menu, string text)
        {
            foreach (ToolStripItem item in menu.Items)
            {
                ToolStripMenuItem menuItem = item as ToolStripMenuItem;
                if (menuItem != null && menuItem.Text == text)
                {
                    return menuItem;
                }
            }

            return null;
        }

        private static void AddLegacyTreeMultiplierItem(ContextMenu menu, Form mainForm)
        {
            if (menu == null)
            {
                return;
            }

            MenuItem multiply = FindLegacyMenuItem(menu, "乘系数");
            if (multiply != null && multiply.MenuItems.Count == 0)
            {
                int existingIndex = menu.MenuItems.IndexOf(multiply);
                menu.MenuItems.Remove(multiply);
                multiply.Dispose();
                multiply = null;
                AddLegacyTreeMultiplierItem(menu, mainForm, existingIndex);
                return;
            }

            if (multiply == null)
            {
                AddLegacyTreeMultiplierItem(menu, mainForm, FindLegacyInsertIndex(menu));
                return;
            }

            EnsureLegacyTargetItem(multiply, "乘到原来的工程量", mainForm, "quantity");
            EnsureLegacyTargetItem(multiply, "乘到定额编号", mainForm, "quotaCode");
        }

        private static void AddLegacyTreeMultiplierItem(ContextMenu menu, Form mainForm, int insertIndex)
        {
            MenuItem multiply = new MenuItem("乘系数");
            EnsureLegacyTargetItem(multiply, "乘到原来的工程量", mainForm, "quantity");
            EnsureLegacyTargetItem(multiply, "乘到定额编号", mainForm, "quotaCode");
            int index = Math.Max(0, Math.Min(insertIndex, menu.MenuItems.Count));
            menu.MenuItems.Add(index, multiply);
            Log("Legacy tree multiplier inserted. index=" + index.ToString(CultureInfo.InvariantCulture));
        }

        private static void EnsureLegacyTargetItem(MenuItem parent, string text, Form mainForm, string target)
        {
            if (FindLegacyMenuItem(parent, text) != null)
            {
                return;
            }

            MenuItem item = new MenuItem(text);
            item.Click += delegate { ApplyToTree(mainForm, target); };
            parent.MenuItems.Add(item);
        }

        private static MenuItem FindLegacyMenuItem(Menu menu, string text)
        {
            if (menu == null)
            {
                return null;
            }

            foreach (MenuItem item in menu.MenuItems)
            {
                if (NormalizeMenuText(item.Text) == NormalizeMenuText(text))
                {
                    return item;
                }
            }

            return null;
        }

        private static bool LooksLikeLegacyTreeMenu(Menu menu)
        {
            return HasAnyLegacyItem(menu, "计算参数设置", "删除条目", "恢复章节", "整理清单编码", "复制数据(D)", "清空数据", "计算结果统计", "删除单项概算标识");
        }

        private static bool HasAnyLegacyItem(Menu menu, params string[] texts)
        {
            if (menu == null)
            {
                return false;
            }

            foreach (MenuItem item in menu.MenuItems)
            {
                string itemText = NormalizeMenuText(item.Text);
                foreach (string text in texts)
                {
                    if (itemText == NormalizeMenuText(text))
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool LooksLikeTreeMenu(ToolStrip menu)
        {
            int matches = 0;
            if (HasAnyItem(menu, "计算参数设置")) matches++;
            if (HasAnyItem(menu, "删除条目")) matches++;
            if (HasAnyItem(menu, "整理清单编码")) matches++;
            if (HasAnyItem(menu, "复制数据(D)", "复制数据(&D)")) matches++;
            if (HasAnyItem(menu, "清空数据")) matches++;
            if (HasAnyItem(menu, "计算结果统计")) matches++;
            if (HasAnyItem(menu, "插入子级", "恢复章节", "删除单项概算标识")) matches++;
            return matches >= 2;
        }

        private static int FindLegacyInsertIndex(ContextMenu menu)
        {
            string[] anchors = new string[] { "计算参数设置", "删除单项概算标识", "删除条目", "恢复章节", "整理清单编码", "复制数据(D)", "计算结果统计" };
            for (int i = 0; i < menu.MenuItems.Count; i++)
            {
                string text = NormalizeMenuText(menu.MenuItems[i].Text);
                foreach (string anchor in anchors)
                {
                    if (text == NormalizeMenuText(anchor))
                    {
                        return i + 1;
                    }
                }
            }

            return menu.MenuItems.Count;
        }

        private static string NormalizeMenuText(string text)
        {
            if (String.IsNullOrEmpty(text))
            {
                return String.Empty;
            }

            return text.Replace("&", String.Empty).Trim();
        }

        private sealed class NativeTreeMenuFilter : IMessageFilter
        {
            private const int WM_CONTEXTMENU = 0x007B;
            private const int WM_COMMAND = 0x0111;
            private const int WM_INITMENUPOPUP = 0x0117;
            private const int WM_RBUTTONDOWN = 0x0204;
            private const int WM_RBUTTONUP = 0x0205;
            private const uint MF_STRING = 0x0000;
            private const uint MF_POPUP = 0x0010;
            private const uint MF_BYPOSITION = 0x0400;
            private const int NativeQuantityCommand = 28433;
            private const int NativeQuotaCodeCommand = 28434;

            private readonly Form mainForm;
            private bool scanScheduled;

            public NativeTreeMenuFilter(Form mainForm)
            {
                this.mainForm = mainForm;
            }

            public bool PreFilterMessage(ref Message m)
            {
                if (m.Msg == WM_INITMENUPOPUP)
                {
                    TryPatchNativeTreeMenu(m.WParam);
                    return false;
                }

                if (m.Msg == WM_CONTEXTMENU || m.Msg == WM_RBUTTONDOWN || m.Msg == WM_RBUTTONUP)
                {
                    ScheduleVisibleToolStripScan();
                    return false;
                }

                if (m.Msg == WM_COMMAND)
                {
                    int command = unchecked((int)((long)m.WParam & 0xFFFF));
                    if (command == NativeQuantityCommand)
                    {
                        ApplyToTree(mainForm, "quantity");
                        return true;
                    }
                    if (command == NativeQuotaCodeCommand)
                    {
                        ApplyToTree(mainForm, "quotaCode");
                        return true;
                    }
                }

                return false;
            }

            private void ScheduleVisibleToolStripScan()
            {
                if (scanScheduled)
                {
                    return;
                }

                scanScheduled = true;
                Timer timer = new Timer();
                timer.Interval = 60;
                timer.Tick += delegate
                {
                    timer.Stop();
                    timer.Dispose();
                    scanScheduled = false;
                    PatchVisibleToolStripTreeMenus(mainForm);
                };
                timer.Start();
            }

            private void TryPatchNativeTreeMenu(IntPtr menu)
            {
                if (menu == IntPtr.Zero || HasNativeMenuItem(menu, "乘系数") || !LooksLikeNativeTreeMenu(menu))
                {
                    return;
                }

                IntPtr submenu = CreatePopupMenu();
                if (submenu == IntPtr.Zero)
                {
                    return;
                }

                AppendMenu(submenu, MF_STRING, new UIntPtr((uint)NativeQuantityCommand), "乘到原来的工程量");
                AppendMenu(submenu, MF_STRING, new UIntPtr((uint)NativeQuotaCodeCommand), "乘到定额编号");

                UIntPtr popupId = UIntPtr.Size == 8
                    ? new UIntPtr((ulong)submenu.ToInt64())
                    : new UIntPtr((uint)submenu.ToInt32());
                int index = Math.Max(0, FindNativeInsertIndex(menu));
                InsertMenu(menu, (uint)index, MF_BYPOSITION | MF_POPUP, popupId, "乘系数");
                Log("Native tree multiplier inserted. index=" + index.ToString(CultureInfo.InvariantCulture) + ", visible=" + NativeMenuText(menu));
            }

            private static bool PatchVisibleToolStripTreeMenus(Form mainForm)
            {
                return FormPanel.PatchVisibleToolStripTreeMenus(mainForm);
            }

            private static bool LooksLikeNativeTreeMenu(IntPtr menu)
            {
                int matches = 0;
                if (HasNativeMenuItem(menu, "计算参数设置")) matches++;
                if (HasNativeMenuItem(menu, "删除条目")) matches++;
                if (HasNativeMenuItem(menu, "整理清单编码")) matches++;
                if (HasNativeMenuItem(menu, "复制数据(D)")) matches++;
                if (HasNativeMenuItem(menu, "清空数据")) matches++;
                if (HasNativeMenuItem(menu, "计算结果统计")) matches++;
                if (HasNativeMenuItem(menu, "恢复章节") || HasNativeMenuItem(menu, "删除单项概算标识")) matches++;
                return matches >= 2;
            }

            private static int FindNativeInsertIndex(IntPtr menu)
            {
                string[] anchors = new string[] { "计算参数设置", "删除单项概算标识", "删除条目", "恢复章节", "整理清单编码", "复制数据(D)", "计算结果统计" };
                int count = GetMenuItemCount(menu);
                for (int i = 0; i < count; i++)
                {
                    string text = NormalizeMenuText(GetNativeMenuItemText(menu, i));
                    foreach (string anchor in anchors)
                    {
                        if (text == NormalizeMenuText(anchor))
                        {
                            return i + 1;
                        }
                    }
                }

                return count;
            }

            private static bool HasNativeMenuItem(IntPtr menu, string text)
            {
                int count = GetMenuItemCount(menu);
                string expected = NormalizeMenuText(text);
                for (int i = 0; i < count; i++)
                {
                    if (NormalizeMenuText(GetNativeMenuItemText(menu, i)) == expected)
                    {
                        return true;
                    }
                }

                return false;
            }

            private static string NativeMenuText(IntPtr menu)
            {
                List<string> texts = new List<string>();
                int count = GetMenuItemCount(menu);
                for (int i = 0; i < count; i++)
                {
                    texts.Add(GetNativeMenuItemText(menu, i));
                }

                return String.Join("|", texts.ToArray());
            }

            private static string GetNativeMenuItemText(IntPtr menu, int index)
            {
                StringBuilder buffer = new StringBuilder(256);
                int length = GetMenuString(menu, (uint)index, buffer, buffer.Capacity, MF_BYPOSITION);
                return length > 0 ? buffer.ToString() : String.Empty;
            }

            [DllImport("user32.dll")]
            private static extern int GetMenuItemCount(IntPtr hMenu);

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            private static extern int GetMenuString(IntPtr hMenu, uint uIDItem, StringBuilder lpString, int nMaxCount, uint uFlag);

            [DllImport("user32.dll")]
            private static extern IntPtr CreatePopupMenu();

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool AppendMenu(IntPtr hMenu, uint uFlags, UIntPtr uIDNewItem, string lpNewItem);

            [DllImport("user32.dll", CharSet = CharSet.Auto)]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool InsertMenu(IntPtr hMenu, uint uPosition, uint uFlags, UIntPtr uIDNewItem, string lpNewItem);
        }

        private delegate bool EnumThreadWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumThreadWindows(uint dwThreadId, EnumThreadWindowsProc lpfn, IntPtr lParam);

        [DllImport("kernel32.dll")]
        private static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        private static void ApplyMenuIcon(ToolStripMenuItem item, string iconName)
        {
            if (item == null)
            {
                return;
            }

            System.Drawing.Image image = LoadMenuIcon(iconName);
            if (image != null)
            {
                item.Image = image;
                item.ImageScaling = ToolStripItemImageScaling.SizeToFit;
            }
        }

        private static System.Drawing.Image LoadMenuIcon(string iconName)
        {
            if (String.IsNullOrEmpty(iconName))
            {
                return null;
            }

            System.Drawing.Image cached;
            if (MenuIconCache.TryGetValue(iconName, out cached))
            {
                return cached;
            }

            try
            {
                string dir = Path.GetDirectoryName(typeof(FormPanel).Assembly.Location);
                string path = Path.Combine(dir, "RecoExpandPanelIcons", iconName);
                if (!File.Exists(path))
                {
                    return null;
                }

                using (System.Drawing.Image image = System.Drawing.Image.FromFile(path))
                {
                    cached = new System.Drawing.Bitmap(image);
                }

                MenuIconCache[iconName] = cached;
                return cached;
            }
            catch (Exception ex)
            {
                Log("LoadMenuIcon failed: " + iconName + " " + ex.Message);
                return null;
            }
        }

        private static int FindInsertIndex(ToolStrip menu, bool isTreeMenu)
        {
            string[] anchors = isTreeMenu
                ? new string[] { "计算参数设置", "删除单项概算标识", "删除条目", "恢复章节", "整理清单编码", "复制数据(D)", "计算结果统计" }
                : new string[] { "单价分析", "全选(A)", "定额调整", "复制(C)" };

            for (int i = 0; i < menu.Items.Count; i++)
            {
                string text = menu.Items[i].Text;
                foreach (string anchor in anchors)
                {
                    if (text == anchor)
                    {
                        return i + 1;
                    }
                }
            }

            return -1;
        }

        private static int FirstVisibleIndex(ToolStrip menu)
        {
            for (int i = 0; i < menu.Items.Count; i++)
            {
                if (menu.Items[i].Available && menu.Items[i].Visible)
                {
                    return i;
                }
            }

            return -1;
        }

        private static string VisibleItemText(ToolStrip menu)
        {
            List<string> texts = new List<string>();
            foreach (ToolStripItem item in menu.Items)
            {
                if (item.Available && item.Visible)
                {
                    texts.Add(String.IsNullOrEmpty(item.Text) ? "<sep>" : item.Text);
                }
            }

            return String.Join("|", texts.ToArray());
        }

        private static void ApplyToTree(Form mainForm)
        {
            ApplyToTree(mainForm, "quantity");
        }

        private static void ApplyToTree(Form mainForm, string target)
        {
            FactorInfo factor = PromptFactor(mainForm);
            if (factor == null)
            {
                return;
            }

            TreeView tree = GetField<TreeView>(mainForm, "Tv_tree");
            TreeNode node = tree != null ? tree.SelectedNode : GetField<TreeNode>(mainForm, "CurrNode");
            if (node == null)
            {
                MessageBox.Show(mainForm, "请先选择左侧条目。", "乘系数", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            SqlConnection conn = GetProjectConnection(mainForm);
            if (conn == null)
            {
                MessageBox.Show(mainForm, "没有找到当前项目数据库连接。", "乘系数", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string itemNo = ResolveChapterNo(mainForm, conn, node);
            if (String.IsNullOrEmpty(itemNo))
            {
                MessageBox.Show(mainForm, "无法识别当前条目的编号，请切换条目后再试。", "乘系数", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int changed = ApplyFactorByChapterNo(conn, itemNo, factor, target);
            RefreshCurrentQuotaGrid(mainForm);
            Log("Tree factor applied. ChapterNo=" + itemNo + ", Target=" + target + ", Factor=" + factor.Suffix + ", Changed=" + changed.ToString(CultureInfo.InvariantCulture));
            MessageBox.Show(mainForm, "已处理 " + changed.ToString(CultureInfo.InvariantCulture) + " 条定额。", "乘系数", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static void ApplyToSelectedQuotaRows(Form mainForm)
        {
            ApplyToSelectedQuotaRows(mainForm, "quantity");
        }

        private static void ApplyToSelectedQuotaRows(Form mainForm, string target)
        {
            FactorInfo factor = PromptFactor(mainForm);
            if (factor == null)
            {
                return;
            }

            DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
            if (grid == null)
            {
                MessageBox.Show(mainForm, "没有找到定额输入表格。", "乘系数", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            List<QuotaKey> keys = GetSelectedQuotaKeys(grid);
            if (keys.Count == 0)
            {
                MessageBox.Show(mainForm, "请先选择需要调整的定额行。", "乘系数", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            SqlConnection conn = GetProjectConnection(mainForm);
            if (conn == null)
            {
                MessageBox.Show(mainForm, "没有找到当前项目数据库连接。", "乘系数", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int changed = ApplyFactorByQuotaKeys(conn, keys, factor, target);
            RefreshCurrentQuotaGrid(mainForm);
            Log("Grid factor applied. Target=" + target + ", Factor=" + factor.Suffix + ", Changed=" + changed.ToString(CultureInfo.InvariantCulture));
            MessageBox.Show(mainForm, "已处理 " + changed.ToString(CultureInfo.InvariantCulture) + " 条定额。", "乘系数", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static FactorInfo PromptFactor(IWin32Window owner)
        {
            using (Form dialog = new Form())
            using (Label label = new Label())
            using (TextBox textBox = new TextBox())
            using (Button ok = new Button())
            using (Button cancel = new Button())
            {
                dialog.Text = "乘系数";
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.MinimizeBox = false;
                dialog.MaximizeBox = false;
                dialog.ClientSize = new System.Drawing.Size(300, 110);

                label.Text = "请输入系数：";
                label.Left = 12;
                label.Top = 15;
                label.Width = 90;

                textBox.Left = 100;
                textBox.Top = 12;
                textBox.Width = 180;

                ok.Text = "确定";
                ok.Left = 124;
                ok.Top = 62;
                ok.Width = 75;
                ok.DialogResult = DialogResult.OK;

                cancel.Text = "取消";
                cancel.Left = 205;
                cancel.Top = 62;
                cancel.Width = 75;
                cancel.DialogResult = DialogResult.Cancel;

                dialog.Controls.Add(label);
                dialog.Controls.Add(textBox);
                dialog.Controls.Add(ok);
                dialog.Controls.Add(cancel);
                dialog.AcceptButton = ok;
                dialog.CancelButton = cancel;

                while (dialog.ShowDialog(owner) == DialogResult.OK)
                {
                    FactorInfo factor;
                    string error;
                    if (TryParseFactor(textBox.Text, out factor, out error))
                    {
                        return factor;
                    }

                    MessageBox.Show(owner, error, "乘系数", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    textBox.SelectAll();
                    textBox.Focus();
                }
            }

            return null;
        }

        private static bool TryParseFactor(string input, out FactorInfo factor, out string error)
        {
            factor = null;
            error = null;
            string text = input == null ? String.Empty : input.Trim();
            if (String.IsNullOrEmpty(text))
            {
                error = "请输入系数。";
                return false;
            }

            string op = "*";
            char first = text[0];
            if (first == '*' || first == '×' || first == 'x' || first == 'X' || first == '＊')
            {
                op = "*";
                text = text.Substring(1).Trim();
            }
            else if (first == '/' || first == '÷' || first == '／')
            {
                op = "/";
                text = text.Substring(1).Trim();
            }

            decimal parsed;
            if (!Decimal.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed) &&
                !Decimal.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out parsed))
            {
                error = "系数格式不正确，请输入数字，或输入 *系数、/系数。";
                return false;
            }

            if (op == "/" && parsed == 0m)
            {
                error = "除系数不能为 0。";
                return false;
            }

            factor = new FactorInfo();
            factor.Operator = op;
            factor.Factor = parsed.ToString(CultureInfo.InvariantCulture);
            return true;
        }

        private static int ApplyFactorByChapterNo(SqlConnection conn, string chapterNo, FactorInfo factor, string target)
        {
            EnsureOpen(conn);
            DataTable table = new DataTable();
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = String.Equals(target, "quotaCode", StringComparison.OrdinalIgnoreCase)
                    ? "select 定额序号, 定额编号 from 定额输入 where 条目序号 in (select 条目序号 from 章节表 where 条目编号 like @bh) order by 定额序号"
                    : "select 定额序号, 工程数量输入 from 定额输入 where 条目序号 in (select 条目序号 from 章节表 where 条目编号 like @bh) order by 定额序号";
                cmd.Parameters.AddWithValue("@bh", chapterNo + "%");
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    adapter.Fill(table);
                }
            }

            return UpdateRows(conn, table, factor, target);
        }

        private static int ApplyFactorByQuotaKeys(SqlConnection conn, IEnumerable<QuotaKey> quotaKeys, FactorInfo factor, string target)
        {
            EnsureOpen(conn);
            DataTable table = new DataTable();
            foreach (QuotaKey key in quotaKeys.GroupBy(k => k.Key).Select(g => g.First()))
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = String.Equals(target, "quotaCode", StringComparison.OrdinalIgnoreCase)
                        ? "select 定额序号, 定额编号 from 定额输入 where 总概算序号=@zgs and 条目序号=@tm and 顺号=@xh"
                        : "select 定额序号, 工程数量输入 from 定额输入 where 总概算序号=@zgs and 条目序号=@tm and 顺号=@xh";
                    cmd.Parameters.AddWithValue("@zgs", key.TotalNo);
                    cmd.Parameters.AddWithValue("@tm", key.ChapterSeq);
                    cmd.Parameters.AddWithValue("@xh", key.OrderNo);
                    using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                    {
                        adapter.Fill(table);
                    }
                }
            }

            return UpdateRows(conn, table, factor, target);
        }

        private static int UpdateRows(SqlConnection conn, DataTable table, FactorInfo factor, string target)
        {
            int changed = 0;
            using (SqlTransaction transaction = conn.BeginTransaction())
            using (SqlCommand quantityUpdate = conn.CreateCommand())
            using (SqlCommand codeUpdate = conn.CreateCommand())
            {
                quantityUpdate.Transaction = transaction;
                quantityUpdate.CommandText = "update 定额输入 set 工程数量输入=@value, 工程数量=@quantity where 定额序号=@id";
                quantityUpdate.Parameters.Add("@value", SqlDbType.NVarChar, 200);
                quantityUpdate.Parameters.Add("@quantity", SqlDbType.Float);
                quantityUpdate.Parameters.Add("@id", SqlDbType.BigInt);

                codeUpdate.Transaction = transaction;
                codeUpdate.CommandText = "update 定额输入 set 定额编号=@value where 定额序号=@id";
                codeUpdate.Parameters.Add("@value", SqlDbType.NVarChar, 200);
                codeUpdate.Parameters.Add("@id", SqlDbType.BigInt);

                try
                {
                    foreach (DataRow row in table.Rows)
                    {
                        string oldValue = String.Equals(target, "quotaCode", StringComparison.OrdinalIgnoreCase)
                            ? Convert.ToString(row["定额编号"]).Trim()
                            : Convert.ToString(row["工程数量输入"]).Trim();
                        if (String.IsNullOrEmpty(oldValue))
                        {
                            continue;
                        }

                        string expression = String.Equals(target, "quotaCode", StringComparison.OrdinalIgnoreCase)
                            ? BuildExpression(oldValue, factor)
                            : BuildExpression("(" + oldValue + ")", factor);
                        long quotaSequence = Convert.ToInt64(row["定额序号"], CultureInfo.InvariantCulture);
                        if (String.Equals(target, "quotaCode", StringComparison.OrdinalIgnoreCase))
                        {
                            codeUpdate.Parameters["@value"].Value = expression;
                            codeUpdate.Parameters["@id"].Value = quotaSequence;
                            changed += codeUpdate.ExecuteNonQuery();
                        }
                        else
                        {
                            quantityUpdate.Parameters["@value"].Value = expression;
                            quantityUpdate.Parameters["@quantity"].Value = Convert.ToDouble(EvaluateDecimal(expression), CultureInfo.InvariantCulture);
                            quantityUpdate.Parameters["@id"].Value = quotaSequence;
                            changed += quantityUpdate.ExecuteNonQuery();
                        }
                    }

                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
            }

            return changed;
        }

        private static string BuildExpression(string oldValue, FactorInfo factor)
        {
            string cleanOld = oldValue.Trim();
            return cleanOld + factor.Suffix;
        }

        private static decimal EvaluateDecimal(string expression)
        {
            string normalized = expression.Trim()
                .Replace("×", "*")
                .Replace("X", "*")
                .Replace("x", "*")
                .Replace("（", "(")
                .Replace("）", ")");

            decimal direct;
            if (Decimal.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out direct) ||
                Decimal.TryParse(normalized, NumberStyles.Float, CultureInfo.CurrentCulture, out direct))
            {
                return direct;
            }

            DataTable computeTable = new DataTable();
            object value = computeTable.Compute(normalized, String.Empty);
            return Convert.ToDecimal(value, CultureInfo.InvariantCulture);
        }

        private static void RefreshCurrentQuotaGrid(Form mainForm)
        {
            DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
            if (grid == null)
            {
                return;
            }

            CurrencyManager manager = grid.DataSource != null && grid.BindingContext != null
                ? grid.BindingContext[grid.DataSource, grid.DataMember] as CurrencyManager
                : null;
            if (manager != null)
            {
                manager.Refresh();
            }

            grid.Refresh();
        }

        private static List<QuotaKey> GetSelectedQuotaKeys(DataGridView grid)
        {
            Dictionary<string, QuotaKey> keys = new Dictionary<string, QuotaKey>();
            foreach (DataGridViewRow row in grid.SelectedRows)
            {
                AddQuotaKey(keys, row);
            }

            foreach (DataGridViewCell cell in grid.SelectedCells)
            {
                if (cell.RowIndex >= 0 && cell.RowIndex < grid.Rows.Count)
                {
                    AddQuotaKey(keys, grid.Rows[cell.RowIndex]);
                }
            }

            if (keys.Count == 0 && grid.CurrentRow != null)
            {
                AddQuotaKey(keys, grid.CurrentRow);
            }

            return keys.Values.ToList();
        }

        private static void AddQuotaKey(Dictionary<string, QuotaKey> keys, DataGridViewRow row)
        {
            QuotaKey key;
            if (!TryGetQuotaKey(row, out key))
            {
                return;
            }

            if (!keys.ContainsKey(key.Key))
            {
                keys.Add(key.Key, key);
            }
        }

        private static bool TryGetQuotaKey(DataGridViewRow row, out QuotaKey key)
        {
            key = null;
            if (row == null)
            {
                return false;
            }

            string totalNo = GetRowValue(row, "总概算序号de", "总概算序号");
            string chapterSeq = GetRowValue(row, "条目序号");
            string orderNo = GetRowValue(row, "顺号DE", "顺号");

            if (String.IsNullOrEmpty(totalNo) || String.IsNullOrEmpty(chapterSeq) || String.IsNullOrEmpty(orderNo))
            {
                Log("Selected row lacks quota key. total=" + totalNo + ", chapter=" + chapterSeq + ", order=" + orderNo);
                return false;
            }

            key = new QuotaKey { TotalNo = totalNo, ChapterSeq = chapterSeq, OrderNo = orderNo };
            return true;
        }

        private static string GetRowValue(DataGridViewRow row, params string[] names)
        {
            DataRowView rowView = row.DataBoundItem as DataRowView;
            if (rowView != null)
            {
                foreach (string name in names)
                {
                    if (rowView.DataView.Table.Columns.Contains(name))
                    {
                        object value = rowView[name];
                        if (value != null && value != DBNull.Value)
                        {
                            return Convert.ToString(value).Trim();
                        }
                    }
                }
            }

            if (row.DataGridView != null)
            {
                foreach (string name in names)
                {
                    if (row.DataGridView.Columns.Contains(name))
                    {
                        object value = row.Cells[name].Value;
                        if (value != null)
                        {
                            return Convert.ToString(value).Trim();
                        }
                    }
                }
            }

            return null;
        }

        private static string ResolveChapterNo(Form mainForm, SqlConnection conn, TreeNode node)
        {
            string fromPropGrid = ReadPropertyGridValue(mainForm, "条目编号");
            if (!String.IsNullOrEmpty(fromPropGrid))
            {
                return fromPropGrid;
            }

            string fromResultGrid = ReadCurrentResultGridValue(mainForm, "条目编号");
            if (!String.IsNullOrEmpty(fromResultGrid))
            {
                return fromResultGrid;
            }

            string fromTag = TryGetValue(node.Tag, "条目编号");
            if (!String.IsNullOrEmpty(fromTag))
            {
                return fromTag;
            }

            string seq = TryGetValue(node.Tag, "条目序号");
            if (String.IsNullOrEmpty(seq) && IsNumeric(node.Name))
            {
                seq = node.Name;
            }

            if (!String.IsNullOrEmpty(seq))
            {
                EnsureOpen(conn);
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "select 条目编号 from 章节表 where 条目序号=@id";
                    cmd.Parameters.AddWithValue("@id", seq);
                    object result = cmd.ExecuteScalar();
                    if (result != null && result != DBNull.Value)
                    {
                        return Convert.ToString(result);
                    }
                }
            }

            if (!String.IsNullOrEmpty(node.Name) && !IsNumeric(node.Name))
            {
                return node.Name;
            }

            return null;
        }

        private static string ReadPropertyGridValue(Form mainForm, string propertyName)
        {
            DataGridView propGrid = GetField<DataGridView>(mainForm, "dataGridViewProp");
            if (propGrid == null)
            {
                return null;
            }

            foreach (DataGridViewRow row in propGrid.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

                string name = GetCellText(row, "属性名称");
                if (name == propertyName)
                {
                    return GetCellText(row, "数据");
                }
            }

            return null;
        }

        private static string ReadCurrentResultGridValue(Form mainForm, string columnName)
        {
            DataGridView grid = GetField<DataGridView>(mainForm, "Grid");
            if (grid == null || !grid.Columns.Contains(columnName))
            {
                return null;
            }

            DataGridViewRow row = grid.CurrentRow;
            if (row == null)
            {
                return null;
            }

            return GetCellText(row, columnName);
        }

        private static string GetCellText(DataGridViewRow row, string columnName)
        {
            if (row == null || row.DataGridView == null || !row.DataGridView.Columns.Contains(columnName))
            {
                return null;
            }

            object value = row.Cells[columnName].Value;
            return value == null ? null : Convert.ToString(value).Trim();
        }

        private static string TryGetValue(object source, string name)
        {
            if (source == null)
            {
                return null;
            }

            DataRowView rowView = source as DataRowView;
            if (rowView != null && rowView.DataView.Table.Columns.Contains(name))
            {
                return Convert.ToString(rowView[name]);
            }

            DataRow row = source as DataRow;
            if (row != null && row.Table.Columns.Contains(name))
            {
                return Convert.ToString(row[name]);
            }

            PropertyInfo prop = source.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
            if (prop != null)
            {
                object value = prop.GetValue(source, null);
                return value == null ? null : Convert.ToString(value);
            }

            return null;
        }

        private static void BindSelectedQuotaToExcel(Form mainForm)
        {
            try
            {
                SqlConnection conn = GetProjectConnection(mainForm);
                if (conn == null)
                {
                    MessageBox.Show(mainForm, "没有找到当前项目数据库连接。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
                if (grid == null)
                {
                    MessageBox.Show(mainForm, "没有找到定额输入表格。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DataGridViewRow row = GetCurrentQuotaRow(grid);
                if (row == null)
                {
                    MessageBox.Show(mainForm, "请先选择一条定额行。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                ExcelQuotaLink link;
                string error;
                if (!TryCreateQuotaLink(mainForm, conn, row, out link, out error))
                {
                    MessageBox.Show(mainForm, error, "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                ExcelCellAddress cell;
                if (!TryGetActiveExcelCell(out cell, out error))
                {
                    if (!PromptExcelCell(mainForm, error, out cell))
                    {
                        return;
                    }
                }

                link.ExcelPath = cell.WorkbookPath;
                link.WorksheetName = cell.WorksheetName;
                link.CellAddress = cell.CellAddress;
                link.Expression = cell.CellAddress;
                link.LastSyncValue = cell.DisplayValue;
                link.LastStatus = "已绑定，等待同步";
                link.UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

                ExcelLinkStore store = LoadStore(conn);
                store.Upsert(link);
                SaveStore(conn, store);

                EnsureExcelLinkRuntime(mainForm);
                if (ExcelLinkRuntimes.ContainsKey(mainForm))
                {
                    ExcelLinkRuntimes[mainForm].Reload();
                }

                RefreshExcelLinkPanel(mainForm);
                MessageBox.Show(mainForm, "已绑定：" + link.QuotaCode + " -> " + Path.GetFileName(link.ExcelPath) + "!" + link.WorksheetName + "!" + link.CellAddress, "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log("BindSelectedQuotaToExcel failed: " + ex);
                MessageBox.Show(mainForm, "绑定失败：" + ex.Message, "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void BatchBindSelectedQuotasToExcel(Form mainForm)
        {
            try
            {
                SqlConnection conn = GetProjectConnection(mainForm);
                if (conn == null)
                {
                    MessageBox.Show(mainForm, "没有找到当前项目数据库连接。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
                List<DataGridViewRow> rows = GetSelectedQuotaRows(grid);
                if (rows.Count < 2)
                {
                    MessageBox.Show(mainForm, "请先在定额输入表中选择两条或更多连续定额行。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                ExcelCellAddress startCell;
                string error;
                if (!TryGetActiveExcelCell(out startCell, out error))
                {
                    if (!PromptExcelCell(mainForm, error, out startCell))
                    {
                        return;
                    }
                }

                CellRef startRef;
                if (!TryParseCellAddress(startCell.CellAddress, out startRef))
                {
                    MessageBox.Show(mainForm, "起始单元格地址不正确，例如 E4。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                ExcelLinkStore store = LoadStore(conn);
                int bound = 0;
                foreach (DataGridViewRow row in rows)
                {
                    ExcelQuotaLink link;
                    if (!TryCreateQuotaLink(mainForm, conn, row, out link, out error))
                    {
                        continue;
                    }

                    string address = ColumnNumberToName(startRef.Column) + (startRef.Row + bound).ToString(CultureInfo.InvariantCulture);
                    string displayValue;
                    string readError;
                    TryReadXlsxCellValue(startCell.WorkbookPath, startCell.WorksheetName, address, out displayValue, out readError);

                    link.ExcelPath = startCell.WorkbookPath;
                    link.WorksheetName = startCell.WorksheetName;
                    link.CellAddress = address;
                    link.Expression = address;
                    link.LastSyncValue = displayValue ?? "";
                    link.LastStatus = "批量绑定，等待同步";
                    link.UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                    store.Upsert(link);
                    bound++;
                }

                SaveStore(conn, store);
                EnsureExcelLinkRuntime(mainForm);
                if (ExcelLinkRuntimes.ContainsKey(mainForm))
                {
                    ExcelLinkRuntimes[mainForm].Reload();
                }

                RefreshExcelLinkPanel(mainForm);
                MessageBox.Show(mainForm, "已批量绑定 " + bound.ToString(CultureInfo.InvariantCulture) + " 条定额，从 " + startCell.WorksheetName + "!" + startCell.CellAddress + " 开始向下对应。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log("BatchBindSelectedQuotasToExcel failed: " + ex);
                MessageBox.Show(mainForm, "批量绑定失败：" + ex.Message, "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static void ShowExcelLinkPanel(Form mainForm)
        {
            EnsureExcelLinkRuntime(mainForm);
            ExcelLinkPanel panel;
            if (!ExcelLinkPanels.TryGetValue(mainForm, out panel) || panel == null || panel.IsDisposed)
            {
                panel = new ExcelLinkPanel(mainForm);
                ExcelLinkPanels[mainForm] = panel;
            }

            panel.Reload();
            panel.Show(mainForm);
            panel.Activate();
        }

        private static void ShowQuickBindPanel(Form mainForm)
        {
            EnsureExcelLinkRuntime(mainForm);
            QuickBindPanel panel;
            if (!QuickBindPanels.TryGetValue(mainForm, out panel) || panel == null || panel.IsDisposed)
            {
                panel = new QuickBindPanel(mainForm);
                QuickBindPanels[mainForm] = panel;
            }

            panel.Show(mainForm);
            panel.Activate();
        }

        private static void RefreshExcelLinkPanel(Form mainForm)
        {
            ExcelLinkPanel panel;
            if (ExcelLinkPanels.TryGetValue(mainForm, out panel) && panel != null && !panel.IsDisposed)
            {
                panel.Reload();
            }
        }

        private static DataGridViewRow GetCurrentQuotaRow(DataGridView grid)
        {
            if (grid == null)
            {
                return null;
            }

            if (grid.SelectedRows.Count > 0)
            {
                return grid.SelectedRows[0];
            }

            if (grid.SelectedCells.Count > 0)
            {
                int rowIndex = grid.SelectedCells[0].RowIndex;
                if (rowIndex >= 0 && rowIndex < grid.Rows.Count)
                {
                    return grid.Rows[rowIndex];
                }
            }

            return grid.CurrentRow;
        }

        private static List<DataGridViewRow> GetSelectedQuotaRows(DataGridView grid)
        {
            Dictionary<int, DataGridViewRow> rows = new Dictionary<int, DataGridViewRow>();
            if (grid == null)
            {
                return new List<DataGridViewRow>();
            }

            foreach (DataGridViewRow row in grid.SelectedRows)
            {
                if (row != null && !row.IsNewRow && row.Index >= 0)
                {
                    rows[row.Index] = row;
                }
            }

            foreach (DataGridViewCell cell in grid.SelectedCells)
            {
                if (cell.RowIndex >= 0 && cell.RowIndex < grid.Rows.Count)
                {
                    DataGridViewRow row = grid.Rows[cell.RowIndex];
                    if (row != null && !row.IsNewRow)
                    {
                        rows[row.Index] = row;
                    }
                }
            }

            if (rows.Count == 0 && grid.CurrentRow != null && !grid.CurrentRow.IsNewRow)
            {
                rows[grid.CurrentRow.Index] = grid.CurrentRow;
            }

            return rows.OrderBy(k => k.Key).Select(k => k.Value).ToList();
        }

        private static bool TryCreateQuotaLink(Form mainForm, SqlConnection conn, DataGridViewRow row, out ExcelQuotaLink link, out string error)
        {
            link = null;
            error = null;

            QuotaKey key;
            if (!TryGetQuotaKey(row, out key))
            {
                error = "无法识别当前定额行的总概算序号、条目序号或顺号。";
                return false;
            }

            long quotaSequence;
            string quotaSequenceText = GetRowValue(row, "定额序号", "定额序号DE");
            if (!Int64.TryParse(quotaSequenceText, NumberStyles.Integer, CultureInfo.InvariantCulture, out quotaSequence))
            {
                quotaSequence = ResolveQuotaSequence(conn, key);
            }

            if (quotaSequence <= 0)
            {
                error = "无法识别当前定额行的定额序号。";
                return false;
            }

            string quotaCode = GetRowValue(row, "定额编号", "定额编号DE", "编号");
            string quotaName = GetRowValue(row, "工程或费用项目名称", "名称", "项目名称");

            link = new ExcelQuotaLink();
            link.ProjectId = GetProjectId(conn);
            link.QuotaSequence = quotaSequence;
            link.TotalNo = key.TotalNo;
            link.ChapterSeq = key.ChapterSeq;
            link.OrderNo = key.OrderNo;
            link.QuotaCode = quotaCode;
            link.QuotaName = quotaName;
            return true;
        }

        private static long ResolveQuotaSequence(SqlConnection conn, QuotaKey key)
        {
            EnsureOpen(conn);
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 定额序号 from 定额输入 where 总概算序号=@zgs and 条目序号=@tm and 顺号=@xh";
                cmd.Parameters.AddWithValue("@zgs", key.TotalNo);
                cmd.Parameters.AddWithValue("@tm", key.ChapterSeq);
                cmd.Parameters.AddWithValue("@xh", key.OrderNo);
                object result = cmd.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                {
                    return Convert.ToInt64(result, CultureInfo.InvariantCulture);
                }
            }

            return 0;
        }

        private static bool TryGetActiveExcelCell(out ExcelCellAddress cell, out string error)
        {
            cell = null;
            error = null;
            dynamic excel = null;
            try
            {
                excel = GetActiveSpreadsheetApplication();
                if (excel == null)
                {
                    error = "没有找到正在运行的 Excel/WPS 表格。请先打开表格，并选中要绑定的工程数量单元格。";
                    return false;
                }
                dynamic workbook = excel.ActiveWorkbook;
                dynamic sheet = excel.ActiveSheet;
                dynamic selection = excel.Selection;
                if (workbook == null || sheet == null || selection == null)
                {
                    error = "请先打开 Excel，并选中要绑定的工程数量单元格。";
                    return false;
                }

                dynamic firstCell = selection.Cells[1, 1];
                string workbookPath = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture);
                string worksheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture);
                string address = Convert.ToString(firstCell.Address(false, false), CultureInfo.InvariantCulture);
                object value = firstCell.Value2;
                if (String.IsNullOrEmpty(workbookPath) || String.IsNullOrEmpty(worksheetName) || String.IsNullOrEmpty(address))
                {
                    error = "无法读取当前 Excel 单元格地址。";
                    return false;
                }

                cell = new ExcelCellAddress();
                cell.WorkbookPath = workbookPath;
                cell.WorksheetName = worksheetName;
                cell.CellAddress = address;
                cell.DisplayValue = ExcelValueToText(value);
                return true;
            }
            catch (COMException)
            {
                error = "没有找到正在运行的 Excel/WPS 表格。请先打开表格，并选中要绑定的工程数量单元格。";
                return false;
            }
            catch (Exception ex)
            {
                error = "读取 Excel 当前单元格失败：" + ex.Message;
                return false;
            }
        }

        private static object GetActiveSpreadsheetApplication()
        {
            string[] progIds = new string[]
            {
                "ket.Application",
                "KET.Application",
                "et.Application",
                "Excel.Application"
            };

            foreach (string progId in progIds)
            {
                try
                {
                    object app = Marshal.GetActiveObject(progId);
                    if (app != null)
                    {
                        return app;
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        private static bool PromptExcelCell(IWin32Window owner, string reason, out ExcelCellAddress cell)
        {
            cell = null;
            using (Form dialog = new Form())
            using (Label info = new Label())
            using (Label fileLabel = new Label())
            using (TextBox fileText = new TextBox())
            using (Button browse = new Button())
            using (Label sheetLabel = new Label())
            using (ComboBox sheetBox = new ComboBox())
            using (Label cellLabel = new Label())
            using (TextBox cellText = new TextBox())
            using (DataGridView preview = new DataGridView())
            using (Button ok = new Button())
            using (Button cancel = new Button())
            {
                dialog.Text = "手动绑定Excel工程量";
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.MinimizeBox = false;
                dialog.MaximizeBox = false;
                dialog.ClientSize = new System.Drawing.Size(820, 520);

                info.Left = 12;
                info.Top = 12;
                info.Width = 590;
                info.Height = 34;
                info.Text = reason + "；也可以在这里手动选择 .xlsx 文件并填写工作表和单元格。";

                fileLabel.Text = "Excel文件";
                fileLabel.Left = 12;
                fileLabel.Top = 58;
                fileLabel.Width = 80;

                fileText.Left = 94;
                fileText.Top = 55;
                fileText.Width = 600;

                browse.Text = "选择";
                browse.Left = 704;
                browse.Top = 53;
                browse.Width = 80;
                browse.Click += delegate
                {
                    using (OpenFileDialog chooser = new OpenFileDialog())
                    {
                        chooser.Title = "选择工程数量Excel文件";
                        chooser.Filter = "Excel工作簿 (*.xlsx)|*.xlsx|所有文件 (*.*)|*.*";
                        if (chooser.ShowDialog(dialog) == DialogResult.OK)
                        {
                            fileText.Text = chooser.FileName;
                            LoadSheetNamesIntoCombo(chooser.FileName, sheetBox);
                            LoadPreviewGrid(chooser.FileName, Convert.ToString(sheetBox.Text, CultureInfo.InvariantCulture), preview);
                        }
                    }
                };

                sheetLabel.Text = "工作表";
                sheetLabel.Left = 12;
                sheetLabel.Top = 96;
                sheetLabel.Width = 80;

                sheetBox.Left = 94;
                sheetBox.Top = 93;
                sheetBox.Width = 360;
                sheetBox.DropDownStyle = ComboBoxStyle.DropDown;
                sheetBox.SelectedIndexChanged += delegate
                {
                    LoadPreviewGrid(fileText.Text.Trim(), Convert.ToString(sheetBox.Text, CultureInfo.InvariantCulture), preview);
                };

                cellLabel.Text = "单元格";
                cellLabel.Left = 475;
                cellLabel.Top = 96;
                cellLabel.Width = 60;

                cellText.Left = 535;
                cellText.Top = 93;
                cellText.Width = 80;
                cellText.Text = "E4";

                preview.Left = 12;
                preview.Top = 132;
                preview.Width = 796;
                preview.Height = 320;
                preview.ReadOnly = true;
                preview.AllowUserToAddRows = false;
                preview.AllowUserToDeleteRows = false;
                preview.SelectionMode = DataGridViewSelectionMode.CellSelect;
                preview.RowHeadersWidth = 54;
                preview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                preview.CellClick += delegate(object sender, DataGridViewCellEventArgs e)
                {
                    if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                    {
                        cellText.Text = ColumnNumberToName(e.ColumnIndex + 1) + (e.RowIndex + 1).ToString(CultureInfo.InvariantCulture);
                        dialog.DialogResult = DialogResult.OK;
                        dialog.Close();
                    }
                };

                ok.Text = "确定";
                ok.Left = 620;
                ok.Top = 468;
                ok.Width = 80;
                ok.DialogResult = DialogResult.OK;

                cancel.Text = "取消";
                cancel.Left = 706;
                cancel.Top = 468;
                cancel.Width = 80;
                cancel.DialogResult = DialogResult.Cancel;

                dialog.Controls.Add(info);
                dialog.Controls.Add(fileLabel);
                dialog.Controls.Add(fileText);
                dialog.Controls.Add(browse);
                dialog.Controls.Add(sheetLabel);
                dialog.Controls.Add(sheetBox);
                dialog.Controls.Add(cellLabel);
                dialog.Controls.Add(cellText);
                dialog.Controls.Add(preview);
                dialog.Controls.Add(ok);
                dialog.Controls.Add(cancel);
                dialog.AcceptButton = ok;
                dialog.CancelButton = cancel;

                while (dialog.ShowDialog(owner) == DialogResult.OK)
                {
                    string path = fileText.Text.Trim();
                    string sheet = Convert.ToString(sheetBox.Text, CultureInfo.InvariantCulture).Trim();
                    string address = cellText.Text.Trim().ToUpperInvariant();
                    if (!File.Exists(path))
                    {
                        MessageBox.Show(owner, "请选择存在的 Excel 文件。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        continue;
                    }

                    if (String.IsNullOrEmpty(sheet) || String.IsNullOrEmpty(address))
                    {
                        MessageBox.Show(owner, "请填写工作表和单元格地址，例如 E4。", "Excel联动", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        continue;
                    }

                    string displayValue;
                    string readError;
                    if (!TryReadWorkbookCellValue(path, sheet, address, out displayValue, out readError))
                    {
                        DialogResult result = MessageBox.Show(owner, "当前无法读取该单元格：" + readError + Environment.NewLine + "仍然保存绑定吗？", "Excel联动", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (result != DialogResult.Yes)
                        {
                            continue;
                        }
                    }

                    cell = new ExcelCellAddress();
                    cell.WorkbookPath = path;
                    cell.WorksheetName = sheet;
                    cell.CellAddress = address;
                    cell.DisplayValue = displayValue ?? "";
                    return true;
                }
            }

            return false;
        }

        private static void LoadSheetNamesIntoCombo(string path, ComboBox sheetBox)
        {
            sheetBox.Items.Clear();
            string error;
            foreach (string name in GetSheetNamesByNpoi(path, out error))
            {
                sheetBox.Items.Add(name);
            }

            if (sheetBox.Items.Count == 0)
            {
                foreach (string name in GetXlsxSheetNames(path, out error))
                {
                    sheetBox.Items.Add(name);
                }
            }

            if (sheetBox.Items.Count > 0)
            {
                sheetBox.SelectedIndex = 0;
            }
        }

        private static void LoadPreviewGrid(string path, string sheetName, DataGridView preview)
        {
            preview.Columns.Clear();
            preview.Rows.Clear();
            if (String.IsNullOrEmpty(path) || String.IsNullOrEmpty(sheetName) || !File.Exists(path))
            {
                return;
            }

            const int maxRows = 40;
            const int maxCols = 12;
            for (int col = 1; col <= maxCols; col++)
            {
                string name = ColumnNumberToName(col);
                DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
                column.Name = name;
                column.HeaderText = name;
                column.Width = col == 2 || col == 3 ? 180 : 80;
                column.FillWeight = col == 2 || col == 3 ? 180 : 80;
                preview.Columns.Add(column);
            }

            for (int row = 1; row <= maxRows; row++)
            {
                int rowIndex = preview.Rows.Add();
                preview.Rows[rowIndex].HeaderCell.Value = row.ToString(CultureInfo.InvariantCulture);
            }

            Dictionary<string, string> values = ReadWorkbookSheetCells(path, sheetName, maxRows, maxCols);
            foreach (KeyValuePair<string, string> pair in values)
            {
                CellRef cell;
                if (TryParseCellAddress(pair.Key, out cell) && cell.Row >= 1 && cell.Row <= maxRows && cell.Column >= 1 && cell.Column <= maxCols)
                {
                    preview.Rows[cell.Row - 1].Cells[cell.Column - 1].Value = pair.Value;
                }
            }
        }

        private static bool TryReadExcelCellValue(ExcelQuotaLink link, out string valueText, out decimal quantity, out string error)
        {
            valueText = null;
            quantity = 0;
            error = null;

            dynamic excel = null;
            try
            {
                excel = GetActiveSpreadsheetApplication();
                if (excel == null)
                {
                    return TryReadExcelCellValueFromFile(link, out valueText, out quantity, out error);
                }
                dynamic workbooks = excel.Workbooks;
                dynamic targetWorkbook = null;
                int count = Convert.ToInt32(workbooks.Count, CultureInfo.InvariantCulture);
                for (int i = 1; i <= count; i++)
                {
                    dynamic workbook = workbooks.Item(i);
                    string fullName = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture);
                    if (String.Equals(Path.GetFullPath(fullName), Path.GetFullPath(link.ExcelPath), StringComparison.OrdinalIgnoreCase))
                    {
                        targetWorkbook = workbook;
                        break;
                    }
                }

                if (targetWorkbook == null)
                {
                    return TryReadExcelCellValueFromFile(link, out valueText, out quantity, out error);
                }

                dynamic sheet = targetWorkbook.Worksheets[link.WorksheetName];
                dynamic range = sheet.Range[link.CellAddress];
                object rawValue = range.Value2;
                valueText = ExcelValueToText(rawValue);
                if (String.IsNullOrWhiteSpace(valueText))
                {
                    error = "Excel 单元格为空";
                    return false;
                }

                quantity = EvaluateDecimal(valueText);
                return true;
            }
            catch (COMException ex)
            {
                if (TryReadExcelCellValueFromFile(link, out valueText, out quantity, out error))
                {
                    return true;
                }

                error = "读取 Excel 失败：" + ex.Message + "；文件读取也失败：" + error;
                return false;
            }
            catch (Exception ex)
            {
                if (TryReadExcelCellValueFromFile(link, out valueText, out quantity, out error))
                {
                    return true;
                }

                error = "数值无法计算：" + ex.Message + "；文件读取也失败：" + error;
                return false;
            }
        }

        private static bool TryReadExcelCellValueFromFile(ExcelQuotaLink link, out string valueText, out decimal quantity, out string error)
        {
            valueText = null;
            quantity = 0;
            error = null;
            if (String.IsNullOrEmpty(link.ExcelPath) || !File.Exists(link.ExcelPath))
            {
                error = "Excel 文件不存在";
                return false;
            }

            string expression = String.IsNullOrWhiteSpace(link.Expression) ? link.CellAddress : link.Expression;
            if (!TryEvaluateWorkbookExpression(link.ExcelPath, link.WorksheetName, expression, out valueText, out quantity, out error))
            {
                return false;
            }

            if (String.IsNullOrWhiteSpace(valueText))
            {
                error = "Excel 单元格为空";
                return false;
            }

            return true;
        }

        private static bool TryEvaluateWorkbookExpression(string path, string sheetName, string expression, out string valueText, out decimal quantity, out string error)
        {
            valueText = null;
            quantity = 0;
            error = null;

            if (String.IsNullOrWhiteSpace(expression))
            {
                error = "Excel 表达式为空";
                return false;
            }

            string normalized = NormalizeCellAddress(expression).Replace("×", "*").Replace("（", "(").Replace("）", ")");
            CellRef onlyCell;
            if (TryParseCellAddress(normalized, out onlyCell))
            {
                string cellValue;
                if (!TryReadWorkbookCellValue(path, sheetName, normalized, out cellValue, out error))
                {
                    return false;
                }

                quantity = EvaluateDecimal(cellValue);
                valueText = cellValue;
                return true;
            }

            string resolved = ResolveWorkbookExpression(path, sheetName, normalized, out error);
            if (resolved == null)
            {
                return false;
            }

            quantity = EvaluateDecimal(resolved);
            valueText = quantity.ToString(CultureInfo.InvariantCulture);
            return true;
        }

        private static string ResolveWorkbookExpression(string path, string sheetName, string expression, out string error)
        {
            error = null;
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < expression.Length;)
            {
                char ch = expression[i];
                if (Char.IsLetter(ch))
                {
                    int start = i;
                    while (i < expression.Length && Char.IsLetter(expression[i]))
                    {
                        i++;
                    }
                    while (i < expression.Length && Char.IsDigit(expression[i]))
                    {
                        i++;
                    }

                    string token = expression.Substring(start, i - start).ToUpperInvariant();
                    CellRef cell;
                    if (!TryParseCellAddress(token, out cell))
                    {
                        error = "表达式里包含无法识别的单元格：" + token;
                        return null;
                    }

                    string cellValue;
                    if (!TryReadWorkbookCellValue(path, sheetName, token, out cellValue, out error))
                    {
                        error = token + " 读取失败：" + error;
                        return null;
                    }

                    decimal parsed = EvaluateDecimal(cellValue);
                    builder.Append(parsed.ToString(CultureInfo.InvariantCulture));
                    continue;
                }

                if ("0123456789.+-*/() ".IndexOf(ch) >= 0)
                {
                    builder.Append(ch);
                    i++;
                    continue;
                }

                error = "表达式里包含不支持的字符：" + ch;
                return null;
            }

            return builder.ToString();
        }

        private static bool TryReadWorkbookCellValue(string path, string sheetName, string cellAddress, out string valueText, out string error)
        {
            if (TryReadCellValueByNpoi(path, sheetName, cellAddress, out valueText, out error))
            {
                return true;
            }

            string npoiError = error;
            if (TryReadXlsxCellValue(path, sheetName, cellAddress, out valueText, out error))
            {
                return true;
            }

            error = "NPOI读取失败：" + npoiError + "；直接读取失败：" + error;
            return false;
        }

        private static bool TryReadXlsxCellValue(string path, string sheetName, string cellAddress, out string valueText, out string error)
        {
            valueText = null;
            error = null;
            try
            {
                using (ZipArchive archive = OpenZipArchiveShared(path))
                {
                    List<string> sharedStrings = ReadSharedStrings(archive);
                    string sheetPath = ResolveSheetPath(archive, sheetName);
                    if (String.IsNullOrEmpty(sheetPath))
                    {
                        error = "找不到工作表：" + sheetName;
                        return false;
                    }

                    ZipArchiveEntry sheetEntry = archive.GetEntry(sheetPath);
                    if (sheetEntry == null)
                    {
                        error = "找不到工作表数据：" + sheetPath;
                        return false;
                    }

                    using (Stream stream = sheetEntry.Open())
                    using (XmlReader reader = XmlReader.Create(stream))
                    {
                        string target = NormalizeCellAddress(cellAddress);
                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c")
                            {
                                string reference = NormalizeCellAddress(reader.GetAttribute("r"));
                                if (String.Equals(reference, target, StringComparison.OrdinalIgnoreCase))
                                {
                                    valueText = ReadCellValue(reader.ReadSubtree(), reader.GetAttribute("t"), sharedStrings);
                                    return true;
                                }
                            }
                        }
                    }
                }

                error = "找不到单元格：" + cellAddress;
                return false;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static Dictionary<string, string> ReadXlsxSheetCells(string path, string sheetName, int maxRows, int maxCols)
        {
            Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (ZipArchive archive = OpenZipArchiveShared(path))
                {
                    List<string> sharedStrings = ReadSharedStrings(archive);
                    string sheetPath = ResolveSheetPath(archive, sheetName);
                    if (String.IsNullOrEmpty(sheetPath))
                    {
                        return values;
                    }

                    ZipArchiveEntry sheetEntry = archive.GetEntry(sheetPath);
                    if (sheetEntry == null)
                    {
                        return values;
                    }

                    using (Stream stream = sheetEntry.Open())
                    using (XmlReader reader = XmlReader.Create(stream))
                    {
                        while (reader.Read())
                        {
                            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "c")
                            {
                                string reference = NormalizeCellAddress(reader.GetAttribute("r"));
                                CellRef cell;
                                if (!TryParseCellAddress(reference, out cell) || cell.Row > maxRows || cell.Column > maxCols)
                                {
                                    continue;
                                }

                                values[reference] = ReadCellValue(reader.ReadSubtree(), reader.GetAttribute("t"), sharedStrings);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("ReadXlsxSheetCells failed: " + ex.Message);
            }

            return values;
        }

        private static Dictionary<string, string> ReadWorkbookSheetCells(string path, string sheetName, int maxRows, int maxCols)
        {
            Dictionary<string, string> values = ReadSheetCellsByNpoi(path, sheetName, maxRows, maxCols);
            if (values.Count > 0)
            {
                return values;
            }

            return ReadXlsxSheetCells(path, sheetName, maxRows, maxCols);
        }

        private static IEnumerable<string> GetSheetNamesByNpoi(string path, out string error)
        {
            error = null;
            List<string> names = new List<string>();
            try
            {
                using (Stream stream = OpenWorkbookStreamShared(path))
                {
                    IWorkbook workbook = WorkbookFactory.Create(stream);
                    for (int i = 0; i < workbook.NumberOfSheets; i++)
                    {
                        names.Add(workbook.GetSheetName(i));
                    }
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
            }

            return names;
        }

        private static bool TryReadCellValueByNpoi(string path, string sheetName, string cellAddress, out string valueText, out string error)
        {
            valueText = null;
            error = null;
            try
            {
                CellRef cellRef;
                if (!TryParseCellAddress(cellAddress, out cellRef))
                {
                    error = "单元格地址不正确：" + cellAddress;
                    return false;
                }

                using (Stream stream = OpenWorkbookStreamShared(path))
                {
                    IWorkbook workbook = WorkbookFactory.Create(stream);
                    ISheet sheet = workbook.GetSheet(sheetName);
                    if (sheet == null)
                    {
                        error = "找不到工作表：" + sheetName;
                        return false;
                    }

                    IRow row = sheet.GetRow(cellRef.Row - 1);
                    ICell cell = row == null ? null : row.GetCell(cellRef.Column - 1);
                    valueText = NpoiCellToText(cell);
                    return true;
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        private static Dictionary<string, string> ReadSheetCellsByNpoi(string path, string sheetName, int maxRows, int maxCols)
        {
            Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (Stream stream = OpenWorkbookStreamShared(path))
                {
                    IWorkbook workbook = WorkbookFactory.Create(stream);
                    ISheet sheet = workbook.GetSheet(sheetName);
                    if (sheet == null)
                    {
                        return values;
                    }

                    for (int rowIndex = 0; rowIndex < maxRows; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        if (row == null)
                        {
                            continue;
                        }

                        for (int colIndex = 0; colIndex < maxCols; colIndex++)
                        {
                            ICell cell = row.GetCell(colIndex);
                            string text = NpoiCellToText(cell);
                            if (!String.IsNullOrEmpty(text))
                            {
                                values[ColumnNumberToName(colIndex + 1) + (rowIndex + 1).ToString(CultureInfo.InvariantCulture)] = text;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("ReadSheetCellsByNpoi failed: " + ex.Message);
            }

            return values;
        }

        private static string NpoiCellToText(ICell cell)
        {
            if (cell == null)
            {
                return "";
            }

            try
            {
                switch (cell.CellType)
                {
                    case CellType.Numeric:
                        return cell.NumericCellValue.ToString("G15", CultureInfo.InvariantCulture);
                    case CellType.String:
                        return cell.StringCellValue == null ? "" : cell.StringCellValue.Trim();
                    case CellType.Boolean:
                        return cell.BooleanCellValue ? "TRUE" : "FALSE";
                    case CellType.Formula:
                        switch (cell.CachedFormulaResultType)
                        {
                            case CellType.Numeric:
                                return cell.NumericCellValue.ToString("G15", CultureInfo.InvariantCulture);
                            case CellType.String:
                                return cell.StringCellValue == null ? "" : cell.StringCellValue.Trim();
                            case CellType.Boolean:
                                return cell.BooleanCellValue ? "TRUE" : "FALSE";
                            default:
                                return cell.ToString().Trim();
                        }
                    default:
                        return cell.ToString().Trim();
                }
            }
            catch
            {
                return cell.ToString().Trim();
            }
        }

        private static IEnumerable<string> GetSheetNamesByExcelCom(string path, out string error)
        {
            error = null;
            List<string> names = new List<string>();
            object excel = null;
            object workbook = null;
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    error = "本机没有可用的 Microsoft Excel COM。";
                    return names;
                }

                excel = Activator.CreateInstance(excelType);
                dynamic app = excel;
                app.Visible = false;
                app.DisplayAlerts = false;
                workbook = app.Workbooks.Open(path, 0, true);
                dynamic wb = workbook;
                int count = Convert.ToInt32(wb.Worksheets.Count, CultureInfo.InvariantCulture);
                for (int i = 1; i <= count; i++)
                {
                    dynamic sheet = wb.Worksheets.Item(i);
                    names.Add(Convert.ToString(sheet.Name, CultureInfo.InvariantCulture));
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
            }
            finally
            {
                CloseExcelComWorkbook(excel, workbook);
            }

            return names;
        }

        private static bool TryReadCellValueByExcelCom(string path, string sheetName, string cellAddress, out string valueText, out string error)
        {
            valueText = null;
            error = null;
            object excel = null;
            object workbook = null;
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    error = "本机没有可用的 Microsoft Excel COM。";
                    return false;
                }

                excel = Activator.CreateInstance(excelType);
                dynamic app = excel;
                app.Visible = false;
                app.DisplayAlerts = false;
                workbook = app.Workbooks.Open(path, 0, true);
                dynamic wb = workbook;
                dynamic sheet = wb.Worksheets[sheetName];
                dynamic range = sheet.Range[cellAddress];
                valueText = ExcelValueToText(range.Value2);
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
            finally
            {
                CloseExcelComWorkbook(excel, workbook);
            }
        }

        private static Dictionary<string, string> ReadSheetCellsByExcelCom(string path, string sheetName, int maxRows, int maxCols)
        {
            Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            object excel = null;
            object workbook = null;
            try
            {
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                if (excelType == null)
                {
                    return values;
                }

                excel = Activator.CreateInstance(excelType);
                dynamic app = excel;
                app.Visible = false;
                app.DisplayAlerts = false;
                workbook = app.Workbooks.Open(path, 0, true);
                dynamic wb = workbook;
                dynamic sheet = wb.Worksheets[sheetName];
                for (int row = 1; row <= maxRows; row++)
                {
                    for (int col = 1; col <= maxCols; col++)
                    {
                        object value = sheet.Cells[row, col].Value2;
                        if (value != null)
                        {
                            values[ColumnNumberToName(col) + row.ToString(CultureInfo.InvariantCulture)] = ExcelValueToText(value);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("ReadSheetCellsByExcelCom failed: " + ex.Message);
            }
            finally
            {
                CloseExcelComWorkbook(excel, workbook);
            }

            return values;
        }

        private static void CloseExcelComWorkbook(object excel, object workbook)
        {
            try
            {
                if (workbook != null)
                {
                    dynamic wb = workbook;
                    wb.Close(false);
                }
            }
            catch
            {
            }

            try
            {
                if (excel != null)
                {
                    dynamic app = excel;
                    app.Quit();
                }
            }
            catch
            {
            }

            try
            {
                if (workbook != null)
                {
                    Marshal.FinalReleaseComObject(workbook);
                }
            }
            catch
            {
            }

            try
            {
                if (excel != null)
                {
                    Marshal.FinalReleaseComObject(excel);
                }
            }
            catch
            {
            }
        }

        private static IEnumerable<string> GetXlsxSheetNames(string path, out string error)
        {
            error = null;
            List<string> names = new List<string>();
            try
            {
                using (ZipArchive archive = OpenZipArchiveShared(path))
                {
                    ZipArchiveEntry workbook = archive.GetEntry("xl/workbook.xml");
                    if (workbook == null)
                    {
                        error = "不是有效的 .xlsx 文件";
                        return names;
                    }

                    XDocument doc;
                    using (Stream stream = workbook.Open())
                    {
                        doc = XDocument.Load(stream);
                    }

                    XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                    foreach (XElement sheet in doc.Descendants(ns + "sheet"))
                    {
                        XAttribute attr = sheet.Attribute("name");
                        if (attr != null)
                        {
                            names.Add(attr.Value);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                error = ex.Message;
            }

            return names;
        }

        private static string ResolveSheetPath(ZipArchive archive, string sheetName)
        {
            ZipArchiveEntry workbook = archive.GetEntry("xl/workbook.xml");
            ZipArchiveEntry rels = archive.GetEntry("xl/_rels/workbook.xml.rels");
            if (workbook == null || rels == null)
            {
                return null;
            }

            XDocument workbookDoc;
            XDocument relsDoc;
            using (Stream stream = workbook.Open())
            {
                workbookDoc = XDocument.Load(stream);
            }

            using (Stream stream = rels.Open())
            {
                relsDoc = XDocument.Load(stream);
            }

            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XNamespace relNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
            XElement targetSheet = workbookDoc.Descendants(ns + "sheet")
                .FirstOrDefault(s => String.Equals((string)s.Attribute("name"), sheetName, StringComparison.OrdinalIgnoreCase));
            if (targetSheet == null)
            {
                return null;
            }

            string relId = (string)targetSheet.Attribute(relNs + "id");
            if (String.IsNullOrEmpty(relId))
            {
                return null;
            }

            XNamespace packageRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
            XElement rel = relsDoc.Descendants(packageRelNs + "Relationship")
                .FirstOrDefault(r => String.Equals((string)r.Attribute("Id"), relId, StringComparison.OrdinalIgnoreCase));
            if (rel == null)
            {
                return null;
            }

            string target = ((string)rel.Attribute("Target") ?? "").Replace('\\', '/');
            if (target.StartsWith("/", StringComparison.Ordinal))
            {
                target = target.TrimStart('/');
            }
            else if (!target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
            {
                target = "xl/" + target;
            }

            return target;
        }

        private static List<string> ReadSharedStrings(ZipArchive archive)
        {
            List<string> values = new List<string>();
            ZipArchiveEntry entry = archive.GetEntry("xl/sharedStrings.xml");
            if (entry == null)
            {
                return values;
            }

            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            using (Stream stream = entry.Open())
            {
                XDocument doc = XDocument.Load(stream);
                foreach (XElement si in doc.Descendants(ns + "si"))
                {
                    values.Add(String.Concat(si.Descendants(ns + "t").Select(t => (string)t)));
                }
            }

            return values;
        }

        private static Stream OpenWorkbookStreamShared(string path)
        {
            try
            {
                return new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            }
            catch
            {
                string temp = Path.Combine(Path.GetTempPath(), "RecoExcelLink_" + Guid.NewGuid().ToString("N") + Path.GetExtension(path));
                File.Copy(path, temp, true);
                return new DeleteOnCloseFileStream(temp);
            }
        }

        private static ZipArchive OpenZipArchiveShared(string path)
        {
            return new ZipArchive(OpenWorkbookStreamShared(path), ZipArchiveMode.Read);
        }

        private sealed class DeleteOnCloseFileStream : FileStream
        {
            private readonly string path;

            public DeleteOnCloseFileStream(string path)
                : base(path, FileMode.Open, FileAccess.Read, FileShare.Read)
            {
                this.path = path;
            }

            protected override void Dispose(bool disposing)
            {
                base.Dispose(disposing);
                try
                {
                    File.Delete(path);
                }
                catch
                {
                }
            }
        }

        private static string ReadCellValue(XmlReader reader, string type, List<string> sharedStrings)
        {
            string value = null;
            using (reader)
            {
                while (reader.Read())
                {
                    if (reader.NodeType == XmlNodeType.Element && (reader.LocalName == "v" || reader.LocalName == "t"))
                    {
                        value = reader.ReadElementContentAsString();
                        if (reader.LocalName == "t")
                        {
                            break;
                        }
                    }
                }
            }

            if (value == null)
            {
                return "";
            }

            if (type == "s")
            {
                int index;
                if (Int32.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out index) && index >= 0 && index < sharedStrings.Count)
                {
                    return sharedStrings[index];
                }
            }

            if (type == "b")
            {
                return value == "1" ? "TRUE" : "FALSE";
            }

            return value.Trim();
        }

        private static string NormalizeCellAddress(string address)
        {
            if (String.IsNullOrEmpty(address))
            {
                return "";
            }

            return address.Replace("$", "").Trim().ToUpperInvariant();
        }

        private static bool TryParseCellAddress(string address, out CellRef cell)
        {
            cell = new CellRef();
            string normalized = NormalizeCellAddress(address);
            if (String.IsNullOrEmpty(normalized))
            {
                return false;
            }

            int index = 0;
            int column = 0;
            while (index < normalized.Length && normalized[index] >= 'A' && normalized[index] <= 'Z')
            {
                column = column * 26 + (normalized[index] - 'A' + 1);
                index++;
            }

            int row;
            if (column <= 0 || index >= normalized.Length || !Int32.TryParse(normalized.Substring(index), NumberStyles.Integer, CultureInfo.InvariantCulture, out row) || row <= 0)
            {
                return false;
            }

            cell.Column = column;
            cell.Row = row;
            return true;
        }

        private static string ExtractFirstCellAddress(string expression)
        {
            if (String.IsNullOrEmpty(expression))
            {
                return null;
            }

            string normalized = NormalizeCellAddress(expression);
            for (int i = 0; i < normalized.Length; i++)
            {
                if (normalized[i] >= 'A' && normalized[i] <= 'Z')
                {
                    int start = i;
                    while (i < normalized.Length && normalized[i] >= 'A' && normalized[i] <= 'Z')
                    {
                        i++;
                    }
                    while (i < normalized.Length && Char.IsDigit(normalized[i]))
                    {
                        i++;
                    }

                    string candidate = normalized.Substring(start, i - start);
                    CellRef parsed;
                    if (TryParseCellAddress(candidate, out parsed))
                    {
                        return candidate;
                    }
                }
            }

            return null;
        }

        private static string ColumnNumberToName(int column)
        {
            if (column <= 0)
            {
                return "";
            }

            StringBuilder builder = new StringBuilder();
            int value = column;
            while (value > 0)
            {
                value--;
                builder.Insert(0, (char)('A' + (value % 26)));
                value /= 26;
            }

            return builder.ToString();
        }

        private struct CellRef
        {
            public int Column;
            public int Row;
        }

        private static string ExcelValueToText(object value)
        {
            if (value == null)
            {
                return "";
            }

            if (value is double || value is float)
            {
                return Convert.ToDouble(value, CultureInfo.InvariantCulture).ToString("G15", CultureInfo.InvariantCulture);
            }

            if (value is decimal)
            {
                return ((decimal)value).ToString(CultureInfo.InvariantCulture);
            }

            if (value is int || value is long || value is short || value is byte)
            {
                return Convert.ToString(value, CultureInfo.InvariantCulture);
            }

            return Convert.ToString(value, CultureInfo.CurrentCulture).Trim();
        }

        private static SyncSummary SyncExcelLinks(Form mainForm, bool manual)
        {
            SyncSummary summary = new SyncSummary();
            SqlConnection conn = GetProjectConnection(mainForm);
            if (conn == null)
            {
                summary.Message = "没有找到当前项目数据库连接。";
                return summary;
            }

            ExcelLinkStore store = LoadStore(conn);
            if (store.Links.Count == 0)
            {
                summary.Message = "当前项目还没有 Excel 联动绑定。";
                return summary;
            }

            List<PendingSync> pending = new List<PendingSync>();
            foreach (ExcelQuotaLink link in store.Links)
            {
                string valueText;
                decimal quantity;
                string error;
                if (!TryReadExcelCellValue(link, out valueText, out quantity, out error))
                {
                    link.LastStatus = error;
                    link.UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                    summary.Skipped++;
                    continue;
                }

                PendingSync item = new PendingSync();
                item.Link = link;
                item.ValueText = valueText;
                item.Quantity = quantity;
                pending.Add(item);
            }

            if (pending.Count == 0)
            {
                SaveStore(conn, store);
                summary.Message = "没有可同步的有效绑定。";
                return summary;
            }

            EnsureOpen(conn);
            using (SqlTransaction transaction = conn.BeginTransaction())
            using (SqlCommand select = conn.CreateCommand())
            using (SqlCommand update = conn.CreateCommand())
            {
                select.Transaction = transaction;
                select.CommandText = "select 工程数量输入 from 定额输入 where 定额序号=@id";
                select.Parameters.Add("@id", SqlDbType.BigInt);

                update.Transaction = transaction;
                update.CommandText = "update 定额输入 set 工程数量输入=@value, 工程数量=@quantity where 定额序号=@id";
                update.Parameters.Add("@value", SqlDbType.NVarChar, 200);
                update.Parameters.Add("@quantity", SqlDbType.Float);
                update.Parameters.Add("@id", SqlDbType.BigInt);

                try
                {
                    foreach (PendingSync item in pending)
                    {
                        select.Parameters["@id"].Value = item.Link.QuotaSequence;
                        object oldValue = select.ExecuteScalar();
                        if (oldValue == null)
                        {
                            item.Link.LastStatus = "定额行不存在";
                            summary.Skipped++;
                            continue;
                        }

                        update.Parameters["@value"].Value = item.ValueText;
                        update.Parameters["@quantity"].Value = Convert.ToDouble(item.Quantity, CultureInfo.InvariantCulture);
                        update.Parameters["@id"].Value = item.Link.QuotaSequence;
                        int changed = update.ExecuteNonQuery();
                        if (changed > 0)
                        {
                            item.Link.LastSyncValue = item.ValueText;
                            item.Link.LastStatus = "已同步";
                            item.Link.UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                            AppendExcelSyncLog(conn, item.Link, Convert.ToString(oldValue, CultureInfo.InvariantCulture), item.ValueText, manual);
                            summary.Changed += changed;
                        }
                        else
                        {
                            item.Link.LastStatus = "定额行不存在";
                            summary.Skipped++;
                        }
                    }

                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
            }

            SaveStore(conn, store);
            RefreshCurrentQuotaGrid(mainForm);
            RefreshExcelLinkPanel(mainForm);
            summary.Message = "已同步 " + summary.Changed.ToString(CultureInfo.InvariantCulture) + " 条，跳过 " + summary.Skipped.ToString(CultureInfo.InvariantCulture) + " 条。";
            return summary;
        }

        private static ExcelLinkStore LoadStore(SqlConnection conn)
        {
            string path = GetStorePath(conn);
            if (!File.Exists(path))
            {
                return new ExcelLinkStore();
            }

            try
            {
                XmlSerializer serializer = new XmlSerializer(typeof(ExcelLinkStore));
                using (FileStream stream = File.OpenRead(path))
                {
                    ExcelLinkStore store = serializer.Deserialize(stream) as ExcelLinkStore;
                    return store ?? new ExcelLinkStore();
                }
            }
            catch (Exception ex)
            {
                Log("Load Excel link store failed: " + ex);
                return new ExcelLinkStore();
            }
        }

        private static void SaveStore(SqlConnection conn, ExcelLinkStore store)
        {
            string path = GetStorePath(conn);
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            XmlSerializer serializer = new XmlSerializer(typeof(ExcelLinkStore));
            using (FileStream stream = File.Create(path))
            {
                serializer.Serialize(stream, store);
            }
        }

        private static string GetStorePath(SqlConnection conn)
        {
            string dir = Path.Combine(Path.GetDirectoryName(typeof(FormPanel).Assembly.Location), "ExcelLinks");
            string name = SafeHash(GetProjectId(conn)) + ".xml";
            return Path.Combine(dir, name);
        }

        private static string GetProjectId(SqlConnection conn)
        {
            if (conn == null)
            {
                return "unknown";
            }

            return conn.DataSource + "|" + conn.Database;
        }

        private static string SafeHash(string text)
        {
            using (MD5 md5 = MD5.Create())
            {
                byte[] bytes = md5.ComputeHash(Encoding.UTF8.GetBytes(text ?? ""));
                StringBuilder builder = new StringBuilder();
                foreach (byte b in bytes)
                {
                    builder.Append(b.ToString("x2", CultureInfo.InvariantCulture));
                }

                return builder.ToString();
            }
        }

        private static void AppendExcelSyncLog(SqlConnection conn, ExcelQuotaLink link, string oldValue, string newValue, bool manual)
        {
            try
            {
                string dir = Path.Combine(Path.GetDirectoryName(GetStorePath(conn)), "Logs");
                Directory.CreateDirectory(dir);
                string path = Path.Combine(dir, SafeHash(GetProjectId(conn)) + ".log");
                string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)
                    + "\t" + (manual ? "Manual" : "Auto")
                    + "\t" + link.QuotaSequence.ToString(CultureInfo.InvariantCulture)
                    + "\t" + link.QuotaCode
                    + "\t" + link.ExcelPath
                    + "\t" + link.WorksheetName + "!" + link.CellAddress
                    + "\t" + oldValue
                    + "\t" + newValue
                    + Environment.NewLine;
                File.AppendAllText(path, line, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Log("AppendExcelSyncLog failed: " + ex.Message);
            }
        }

        private sealed class ExcelCellAddress
        {
            public string WorkbookPath;
            public string WorksheetName;
            public string CellAddress;
            public string DisplayValue;
        }

        public sealed class ExcelQuotaLink
        {
            public string ProjectId { get; set; }
            public long QuotaSequence { get; set; }
            public string TotalNo { get; set; }
            public string ChapterSeq { get; set; }
            public string OrderNo { get; set; }
            public string QuotaCode { get; set; }
            public string QuotaName { get; set; }
            public string ExcelPath { get; set; }
            public string WorksheetName { get; set; }
            public string CellAddress { get; set; }
            public string Expression { get; set; }
            public string LastSyncValue { get; set; }
            public string LastStatus { get; set; }
            public string UpdatedAt { get; set; }
        }

        public sealed class ExcelLinkStore
        {
            public List<ExcelQuotaLink> Links { get; set; }

            public ExcelLinkStore()
            {
                Links = new List<ExcelQuotaLink>();
            }

            public void Upsert(ExcelQuotaLink link)
            {
                for (int i = Links.Count - 1; i >= 0; i--)
                {
                    if (Links[i].QuotaSequence == link.QuotaSequence)
                    {
                        Links.RemoveAt(i);
                    }
                }

                Links.Add(link);
            }
        }

        private sealed class PendingSync
        {
            public ExcelQuotaLink Link;
            public string ValueText;
            public decimal Quantity;
        }

        private sealed class SyncSummary
        {
            public int Changed;
            public int Skipped;
            public string Message;
        }

        private sealed class ExcelLinkRuntime : IDisposable
        {
            private readonly Form mainForm;
            private readonly Timer timer;
            private readonly Dictionary<string, DateTime> knownWriteTimes = new Dictionary<string, DateTime>(StringComparer.OrdinalIgnoreCase);
            private bool syncing;

            public ExcelLinkRuntime(Form mainForm)
            {
                this.mainForm = mainForm;
                timer = new Timer();
                timer.Interval = 1800;
                timer.Tick += delegate { Tick(); };
                Reload();
                timer.Start();
            }

            public void Reload()
            {
                SqlConnection conn = GetProjectConnection(mainForm);
                if (conn == null)
                {
                    return;
                }

                ExcelLinkStore store = LoadStore(conn);
                foreach (ExcelQuotaLink link in store.Links)
                {
                    if (!String.IsNullOrEmpty(link.ExcelPath) && File.Exists(link.ExcelPath) && !knownWriteTimes.ContainsKey(link.ExcelPath))
                    {
                        knownWriteTimes[link.ExcelPath] = File.GetLastWriteTimeUtc(link.ExcelPath);
                    }
                }
            }

            private void Tick()
            {
                if (syncing)
                {
                    return;
                }

                try
                {
                    SqlConnection conn = GetProjectConnection(mainForm);
                    if (conn == null)
                    {
                        return;
                    }

                    ExcelLinkStore store = LoadStore(conn);
                    bool changed = false;
                    foreach (ExcelQuotaLink link in store.Links)
                    {
                        if (String.IsNullOrEmpty(link.ExcelPath) || !File.Exists(link.ExcelPath))
                        {
                            continue;
                        }

                        DateTime writeTime = File.GetLastWriteTimeUtc(link.ExcelPath);
                        DateTime known;
                        if (!knownWriteTimes.TryGetValue(link.ExcelPath, out known))
                        {
                            knownWriteTimes[link.ExcelPath] = writeTime;
                            continue;
                        }

                        if (writeTime > known)
                        {
                            knownWriteTimes[link.ExcelPath] = writeTime;
                            changed = true;
                        }
                    }

                    if (changed)
                    {
                        syncing = true;
                        SyncExcelLinks(mainForm, false);
                    }
                }
                catch (Exception ex)
                {
                    Log("Excel link auto sync failed: " + ex);
                }
                finally
                {
                    syncing = false;
                }
            }

            public void Dispose()
            {
                timer.Stop();
                timer.Dispose();
            }
        }

        private sealed class QuickBindPanel : Form
        {
            private readonly Form mainForm;
            private readonly TextBox fileText;
            private readonly ComboBox sheetBox;
            private readonly DataGridView preview;
            private readonly Label status;
            private readonly CheckBox expressionMode;
            private readonly TextBox expressionText;

            public QuickBindPanel(Form mainForm)
            {
                this.mainForm = mainForm;
                Text = "Excel快速绑定";
                StartPosition = FormStartPosition.CenterParent;
                Size = new System.Drawing.Size(920, 580);
                MinimizeBox = false;

                Label tip = new Label();
                tip.Left = 12;
                tip.Top = 10;
                tip.Width = 880;
                tip.Height = 24;
                tip.Text = "保持本窗口打开：在软件中选中定额行，再单击下方Excel预览单元格即可绑定。定额表内按 Ctrl+E 可打开本窗口。";
                tip.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

                Label fileLabel = new Label();
                fileLabel.Text = "Excel文件";
                fileLabel.Left = 12;
                fileLabel.Top = 45;
                fileLabel.Width = 70;

                fileText = new TextBox();
                fileText.Left = 86;
                fileText.Top = 42;
                fileText.Width = 650;
                fileText.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

                Button choose = new Button();
                choose.Text = "选择";
                choose.Left = 746;
                choose.Top = 40;
                choose.Width = 75;
                choose.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                choose.Click += delegate
                {
                    using (OpenFileDialog chooser = new OpenFileDialog())
                    {
                        chooser.Title = "选择工程数量Excel文件";
                        chooser.Filter = "Excel文件 (*.xls;*.xlsx)|*.xls;*.xlsx|所有文件 (*.*)|*.*";
                        if (chooser.ShowDialog(this) == DialogResult.OK)
                        {
                            fileText.Text = chooser.FileName;
                            LoadSheetNamesIntoCombo(chooser.FileName, sheetBox);
                            LoadPreview();
                        }
                    }
                };

                Button refresh = new Button();
                refresh.Text = "刷新";
                refresh.Left = 828;
                refresh.Top = 40;
                refresh.Width = 75;
                refresh.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                refresh.Click += delegate { LoadPreview(); };

                Label sheetLabel = new Label();
                sheetLabel.Text = "工作表";
                sheetLabel.Left = 12;
                sheetLabel.Top = 82;
                sheetLabel.Width = 70;

                sheetBox = new ComboBox();
                sheetBox.Left = 86;
                sheetBox.Top = 79;
                sheetBox.Width = 360;
                sheetBox.DropDownStyle = ComboBoxStyle.DropDownList;
                sheetBox.SelectedIndexChanged += delegate { LoadPreview(); };

                expressionMode = new CheckBox();
                expressionMode.Text = "计算式模式";
                expressionMode.Left = 466;
                expressionMode.Top = 81;
                expressionMode.Width = 95;
                expressionMode.Anchor = AnchorStyles.Top | AnchorStyles.Left;

                expressionText = new TextBox();
                expressionText.Left = 566;
                expressionText.Top = 79;
                expressionText.Width = 210;
                expressionText.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

                Button plus = CreateExpressionButton("+", 785, 77);
                Button minus = CreateExpressionButton("-", 813, 77);
                Button multiply = CreateExpressionButton("*", 841, 77);
                Button divide = CreateExpressionButton("/", 869, 77);

                Button bindExpression = new Button();
                bindExpression.Text = "绑定计算式";
                bindExpression.Left = 600;
                bindExpression.Top = 506;
                bindExpression.Width = 110;
                bindExpression.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
                bindExpression.Click += delegate { BindExpression(); };

                Button clearExpression = new Button();
                clearExpression.Text = "清空公式";
                clearExpression.Left = 720;
                clearExpression.Top = 506;
                clearExpression.Width = 90;
                clearExpression.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
                clearExpression.Click += delegate { expressionText.Text = ""; };

                preview = new DataGridView();
                preview.Left = 12;
                preview.Top = 116;
                preview.Width = 890;
                preview.Height = 380;
                preview.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
                preview.ReadOnly = true;
                preview.AllowUserToAddRows = false;
                preview.AllowUserToDeleteRows = false;
                preview.SelectionMode = DataGridViewSelectionMode.CellSelect;
                preview.RowHeadersWidth = 54;
                preview.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                preview.CellClick += delegate(object sender, DataGridViewCellEventArgs e)
                {
                    if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                    {
                        string address = ColumnNumberToName(e.ColumnIndex + 1) + (e.RowIndex + 1).ToString(CultureInfo.InvariantCulture);
                        if (expressionMode.Checked || !String.IsNullOrWhiteSpace(expressionText.Text))
                        {
                            AppendExpressionToken(address);
                        }
                        else
                        {
                            BindCurrentQuota(address, address);
                        }
                    }
                };

                status = new Label();
                status.Left = 12;
                status.Top = 506;
                status.Width = 570;
                status.Height = 24;
                status.Text = "请选择Excel文件。";
                status.Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;

                Button close = new Button();
                close.Text = "关闭";
                close.Left = 820;
                close.Top = 504;
                close.Width = 80;
                close.Anchor = AnchorStyles.Right | AnchorStyles.Bottom;
                close.Click += delegate { Hide(); };

                Controls.Add(tip);
                Controls.Add(fileLabel);
                Controls.Add(fileText);
                Controls.Add(choose);
                Controls.Add(refresh);
                Controls.Add(sheetLabel);
                Controls.Add(sheetBox);
                Controls.Add(expressionMode);
                Controls.Add(expressionText);
                Controls.Add(plus);
                Controls.Add(minus);
                Controls.Add(multiply);
                Controls.Add(divide);
                Controls.Add(preview);
                Controls.Add(status);
                Controls.Add(bindExpression);
                Controls.Add(clearExpression);
                Controls.Add(close);
            }

            private Button CreateExpressionButton(string text, int left, int top)
            {
                Button button = new Button();
                button.Text = text;
                button.Left = left;
                button.Top = top;
                button.Width = 24;
                button.Height = 24;
                button.Anchor = AnchorStyles.Top | AnchorStyles.Right;
                button.Click += delegate
                {
                    expressionMode.Checked = true;
                    if (String.IsNullOrEmpty(expressionText.Text.Trim()) && text != "-")
                    {
                        status.Text = "请先点击一个单元格，再选择运算符。";
                        return;
                    }

                    AppendExpressionToken(text);
                };
                return button;
            }

            protected override void OnFormClosing(FormClosingEventArgs e)
            {
                if (e.CloseReason == CloseReason.UserClosing)
                {
                    e.Cancel = true;
                    Hide();
                    return;
                }

                base.OnFormClosing(e);
            }

            private void LoadPreview()
            {
                string path = fileText.Text.Trim();
                string sheet = Convert.ToString(sheetBox.Text, CultureInfo.InvariantCulture);
                LoadPreviewGrid(path, sheet, preview);
                status.Text = preview.Columns.Count == 0 ? "未能读取预览，请确认文件没有损坏或被独占打开。" : "选择软件定额行后，单击一个Excel单元格即可绑定。";
            }

            private void AppendExpressionToken(string token)
            {
                expressionMode.Checked = true;
                expressionText.Text = expressionText.Text + token;
                expressionText.SelectionStart = expressionText.Text.Length;
                status.Text = "计算式：" + expressionText.Text + "；完成后点“绑定计算式”。";
            }

            private void BindExpression()
            {
                string expression = expressionText.Text.Trim();
                if (String.IsNullOrEmpty(expression))
                {
                    status.Text = "请先勾选计算式模式并点击单元格组成表达式，例如 E4+E5 或 E4*1.15。";
                    return;
                }

                string firstCell = ExtractFirstCellAddress(expression);
                if (String.IsNullOrEmpty(firstCell))
                {
                    status.Text = "计算式里至少需要一个单元格地址。";
                    return;
                }

                BindCurrentQuota(firstCell, expression);
            }

            private void BindCurrentQuota(string address, string expression)
            {
                try
                {
                    string path = fileText.Text.Trim();
                    string sheet = Convert.ToString(sheetBox.Text, CultureInfo.InvariantCulture);
                    if (!File.Exists(path) || String.IsNullOrEmpty(sheet))
                    {
                        status.Text = "请先选择Excel文件和工作表。";
                        return;
                    }

                    SqlConnection conn = GetProjectConnection(mainForm);
                    if (conn == null)
                    {
                        status.Text = "没有找到当前项目数据库连接。";
                        return;
                    }

                    DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
                    DataGridViewRow row = GetCurrentQuotaRow(grid);
                    if (row == null)
                    {
                        status.Text = "请先在软件定额表中选中一条定额行。";
                        return;
                    }

                    ExcelQuotaLink link;
                    string error;
                    if (!TryCreateQuotaLink(mainForm, conn, row, out link, out error))
                    {
                        status.Text = error;
                        return;
                    }

                    string displayValue;
                    string readError;
                    decimal quantity;
                    if (!TryEvaluateWorkbookExpression(path, sheet, expression, out displayValue, out quantity, out readError))
                    {
                        status.Text = "计算式无法读取或计算：" + readError;
                        return;
                    }

                    link.ExcelPath = path;
                    link.WorksheetName = sheet;
                    link.CellAddress = address;
                    link.Expression = expression;
                    link.LastSyncValue = displayValue ?? "";
                    link.LastStatus = "快速绑定，等待同步";
                    link.UpdatedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);

                    ExcelLinkStore store = LoadStore(conn);
                    store.Upsert(link);
                    SaveStore(conn, store);

                    EnsureExcelLinkRuntime(mainForm);
                    if (ExcelLinkRuntimes.ContainsKey(mainForm))
                    {
                        ExcelLinkRuntimes[mainForm].Reload();
                    }

                    RefreshExcelLinkPanel(mainForm);
                    status.Text = "已绑定：" + link.QuotaCode + " -> " + Path.GetFileName(path) + "!" + sheet + "!" + expression;
                    if (!String.Equals(address, expression, StringComparison.OrdinalIgnoreCase))
                    {
                        expressionText.Text = "";
                        expressionMode.Checked = false;
                    }
                }
                catch (Exception ex)
                {
                    Log("Quick bind failed: " + ex);
                    status.Text = "绑定失败：" + ex.Message;
                }
            }
        }

        private sealed class ExcelLinkPanel : Form
        {
            private readonly Form mainForm;
            private readonly DataGridView grid;
            private readonly Label status;

            public ExcelLinkPanel(Form mainForm)
            {
                this.mainForm = mainForm;
                Text = "Excel工程量联动";
                StartPosition = FormStartPosition.CenterParent;
                Size = new System.Drawing.Size(980, 520);
                MinimizeBox = false;

                grid = new DataGridView();
                grid.Dock = DockStyle.Fill;
                grid.ReadOnly = true;
                grid.AllowUserToAddRows = false;
                grid.AllowUserToDeleteRows = false;
                grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                grid.MultiSelect = true;
                grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                grid.Columns.Add("QuotaSequence", "定额序号");
                grid.Columns.Add("QuotaCode", "定额编号");
                grid.Columns.Add("QuotaName", "名称");
                grid.Columns.Add("Excel", "Excel单元格");
                grid.Columns.Add("LastValue", "最近值");
                grid.Columns.Add("Status", "状态");
                grid.Columns.Add("UpdatedAt", "更新时间");
                grid.Columns["QuotaSequence"].FillWeight = 70;
                grid.Columns["QuotaCode"].FillWeight = 90;
                grid.Columns["QuotaName"].FillWeight = 180;
                grid.Columns["Excel"].FillWeight = 220;
                grid.Columns["LastValue"].FillWeight = 80;
                grid.Columns["Status"].FillWeight = 100;
                grid.Columns["UpdatedAt"].FillWeight = 110;

                Button sync = new Button();
                sync.Text = "同步一次";
                sync.Width = 90;
                sync.Click += delegate
                {
                    try
                    {
                        SyncSummary result = SyncExcelLinks(mainForm, true);
                        status.Text = result.Message;
                        Reload();
                    }
                    catch (Exception ex)
                    {
                        status.Text = "同步失败：" + ex.Message;
                        Log("Manual Excel sync failed: " + ex);
                    }
                };

                Button refresh = new Button();
                refresh.Text = "刷新";
                refresh.Width = 75;
                refresh.Click += delegate { Reload(); };

                Button delete = new Button();
                delete.Text = "删除选中绑定";
                delete.Width = 110;
                delete.Click += delegate { DeleteSelectedLinks(); };

                Button close = new Button();
                close.Text = "关闭";
                close.Width = 75;
                close.Click += delegate { Hide(); };

                status = new Label();
                status.AutoSize = false;
                status.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
                status.Dock = DockStyle.Fill;

                FlowLayoutPanel buttons = new FlowLayoutPanel();
                buttons.Dock = DockStyle.Top;
                buttons.Height = 42;
                buttons.Padding = new Padding(8);
                buttons.Controls.Add(sync);
                buttons.Controls.Add(refresh);
                buttons.Controls.Add(delete);
                buttons.Controls.Add(close);

                Panel bottom = new Panel();
                bottom.Dock = DockStyle.Bottom;
                bottom.Height = 28;
                bottom.Padding = new Padding(8, 0, 8, 4);
                bottom.Controls.Add(status);

                Controls.Add(grid);
                Controls.Add(bottom);
                Controls.Add(buttons);
            }

            protected override void OnFormClosing(FormClosingEventArgs e)
            {
                if (e.CloseReason == CloseReason.UserClosing)
                {
                    e.Cancel = true;
                    Hide();
                    return;
                }

                base.OnFormClosing(e);
            }

            public void Reload()
            {
                grid.Rows.Clear();
                SqlConnection conn = GetProjectConnection(mainForm);
                if (conn == null)
                {
                    status.Text = "没有找到当前项目数据库连接。";
                    return;
                }

                ExcelLinkStore store = LoadStore(conn);
                foreach (ExcelQuotaLink link in store.Links.OrderBy(l => l.QuotaSequence))
                {
                    string excel = Path.GetFileName(link.ExcelPath) + "!" + link.WorksheetName + "!" + link.CellAddress;
                    grid.Rows.Add(
                        link.QuotaSequence.ToString(CultureInfo.InvariantCulture),
                        link.QuotaCode,
                        link.QuotaName,
                        excel,
                        link.LastSyncValue,
                        link.LastStatus,
                        link.UpdatedAt);
                }

                status.Text = "绑定数量：" + store.Links.Count.ToString(CultureInfo.InvariantCulture);
            }

            private void DeleteSelectedLinks()
            {
                SqlConnection conn = GetProjectConnection(mainForm);
                if (conn == null)
                {
                    status.Text = "没有找到当前项目数据库连接。";
                    return;
                }

                HashSet<long> ids = new HashSet<long>();
                foreach (DataGridViewRow row in grid.SelectedRows)
                {
                    long id;
                    if (row.Cells["QuotaSequence"].Value != null && Int64.TryParse(Convert.ToString(row.Cells["QuotaSequence"].Value, CultureInfo.InvariantCulture), out id))
                    {
                        ids.Add(id);
                    }
                }

                if (ids.Count == 0)
                {
                    status.Text = "请先选择要删除的绑定。";
                    return;
                }

                ExcelLinkStore store = LoadStore(conn);
                int before = store.Links.Count;
                store.Links.RemoveAll(l => ids.Contains(l.QuotaSequence));
                SaveStore(conn, store);
                if (ExcelLinkRuntimes.ContainsKey(mainForm))
                {
                    ExcelLinkRuntimes[mainForm].Reload();
                }

                Reload();
                status.Text = "已删除 " + (before - store.Links.Count).ToString(CultureInfo.InvariantCulture) + " 条绑定。";
            }
        }

        private static bool IsNumeric(string text)
        {
            long value;
            return !String.IsNullOrEmpty(text) && Int64.TryParse(text, out value);
        }

        private static SqlConnection GetProjectConnection(Form mainForm)
        {
            SqlConnection conn = GetField<SqlConnection>(mainForm, "m_ProjectConn");
            if (conn != null)
            {
                return conn;
            }

            object login = mainForm.GetType().Assembly.GetType("RecoNet.RecoLogin");
            return null;
        }

        public static void Log(string message)
        {
            try
            {
                string dir = Path.GetDirectoryName(typeof(FormPanel).Assembly.Location);
                string path = Path.Combine(dir, "RecoExpandPanel.log");
                File.AppendAllText(path, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " " + message + Environment.NewLine);
            }
            catch
            {
            }
        }

        private static T GetField<T>(object target, string name) where T : class
        {
            if (target == null)
            {
                return null;
            }

            FieldInfo field = target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            return field == null ? null : field.GetValue(target) as T;
        }

        private static void EnsureOpen(SqlConnection conn)
        {
            if (conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
        }
    }

    public class AutoLoadDomainManager : AppDomainManager
    {
        public override void InitializeNewDomain(AppDomainSetup appDomainInfo)
        {
            base.InitializeNewDomain(appDomainInfo);
            try
            {
                FormPanel.InstallOnIdle();
                FormPanel.Log("AutoLoadDomainManager initialized.");
            }
            catch (Exception ex)
            {
                FormPanel.Log("AutoLoadDomainManager failed: " + ex);
            }
        }
    }
}
