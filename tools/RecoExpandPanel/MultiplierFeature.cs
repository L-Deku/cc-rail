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
    public partial class FormPanel : Form
    {
        private static readonly HashSet<TreeView> HookedTrees = new HashSet<TreeView>();
        private static readonly Dictionary<Form, NativeTreeMenuFilter> NativeTreeMenuFilters = new Dictionary<Form, NativeTreeMenuFilter>();

        private sealed class FactorInfo
        {
            public string Operator;
            public string Factor;

            public string Suffix
            {
                get { return Operator + Factor; }
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

        }

        private static void ConfigureFactorTargetMenu(ToolStripMenuItem root, Form mainForm, bool isDeMenu, bool isTreeMenu)
        {
            if (root == null)
            {
                return;
            }

            RemoveDropDownMenuItem(root, "乘到原来的工程量");
            RemoveDropDownMenuItem(root, "乘到定额编号");

            ToolStripMenuItem quantity = EnsureFactorFieldItem(root, "工程数量");
            EnsureFactorActionItem(quantity, mainForm, isDeMenu, isTreeMenu, "乘系数", "quantity", "multiply");
            EnsureFactorActionItem(quantity, mainForm, isDeMenu, isTreeMenu, "删系数", "quantity", "remove");

            ToolStripMenuItem quotaCode = EnsureFactorFieldItem(root, "定额编号");
            EnsureFactorActionItem(quotaCode, mainForm, isDeMenu, isTreeMenu, "乘系数", "quota_code", "multiply");
            EnsureFactorActionItem(quotaCode, mainForm, isDeMenu, isTreeMenu, "删系数", "quota_code", "remove");

            ToolStripMenuItem unitPrice = EnsureFactorFieldItem(root, "单价");
            EnsureFactorActionItem(unitPrice, mainForm, isDeMenu, isTreeMenu, "乘系数", "unit_price", "multiply");
            RemoveDropDownMenuItem(unitPrice, "删系数");

            ToolStripMenuItem adjustment = EnsureFactorFieldItem(root, "定额调整");
            EnsureFactorActionItem(adjustment, mainForm, isDeMenu, isTreeMenu, "调整输入", "adjustment", "set_adjustment");
            EnsureFactorActionItem(adjustment, mainForm, isDeMenu, isTreeMenu, "删除调整", "adjustment", "remove_adjustment");
        }

        private static ToolStripMenuItem EnsureFactorFieldItem(ToolStripMenuItem parent, string text)
        {
            ToolStripMenuItem item = FindDropDownMenuItem(parent, text);
            if (item == null)
            {
                item = new ToolStripMenuItem(text);
                parent.DropDownItems.Add(item);
            }

            return item;
        }

        private static void EnsureFactorActionItem(ToolStripMenuItem parent, Form mainForm, bool isDeMenu, bool isTreeMenu, string text, string target, string action)
        {
            if (FindDropDownMenuItem(parent, text) != null)
            {
                return;
            }

            ToolStripMenuItem item = new ToolStripMenuItem(text);
            item.Click += delegate { ApplyContextMenuAction(mainForm, isDeMenu, isTreeMenu, target, action); };
            parent.DropDownItems.Add(item);
        }

        private static void RemoveDropDownMenuItem(ToolStripMenuItem parent, string text)
        {
            ToolStripMenuItem item = FindDropDownMenuItem(parent, text);
            if (item == null)
            {
                return;
            }

            parent.DropDownItems.Remove(item);
            item.Dispose();
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

            ConfigureLegacyFactorTargetMenu(multiply, mainForm);
        }

        private static void AddLegacyTreeMultiplierItem(ContextMenu menu, Form mainForm, int insertIndex)
        {
            MenuItem multiply = new MenuItem("乘系数");
            ConfigureLegacyFactorTargetMenu(multiply, mainForm);
            int index = Math.Max(0, Math.Min(insertIndex, menu.MenuItems.Count));
            menu.MenuItems.Add(index, multiply);
            Log("Legacy tree multiplier inserted. index=" + index.ToString(CultureInfo.InvariantCulture));
        }

        private static void ConfigureLegacyFactorTargetMenu(MenuItem root, Form mainForm)
        {
            RemoveLegacyMenuItem(root, "乘到原来的工程量");
            RemoveLegacyMenuItem(root, "乘到定额编号");

            MenuItem quantity = EnsureLegacyFieldItem(root, "工程数量");
            EnsureLegacyActionItem(quantity, "乘系数", mainForm, "quantity", "multiply");
            EnsureLegacyActionItem(quantity, "删系数", mainForm, "quantity", "remove");

            MenuItem quotaCode = EnsureLegacyFieldItem(root, "定额编号");
            EnsureLegacyActionItem(quotaCode, "乘系数", mainForm, "quota_code", "multiply");
            EnsureLegacyActionItem(quotaCode, "删系数", mainForm, "quota_code", "remove");

            MenuItem unitPrice = EnsureLegacyFieldItem(root, "单价");
            EnsureLegacyActionItem(unitPrice, "乘系数", mainForm, "unit_price", "multiply");
            RemoveLegacyMenuItem(unitPrice, "删系数");

            MenuItem adjustment = EnsureLegacyFieldItem(root, "定额调整");
            EnsureLegacyActionItem(adjustment, "调整输入", mainForm, "adjustment", "set_adjustment");
            EnsureLegacyActionItem(adjustment, "删除调整", mainForm, "adjustment", "remove_adjustment");
        }

        private static MenuItem EnsureLegacyFieldItem(MenuItem parent, string text)
        {
            MenuItem item = FindLegacyMenuItem(parent, text);
            if (item == null)
            {
                item = new MenuItem(text);
                parent.MenuItems.Add(item);
            }

            return item;
        }

        private static void EnsureLegacyActionItem(MenuItem parent, string text, Form mainForm, string target, string action)
        {
            if (FindLegacyMenuItem(parent, text) != null)
            {
                return;
            }

            MenuItem item = new MenuItem(text);
            item.Click += delegate { ApplyContextMenuAction(mainForm, false, true, target, action); };
            parent.MenuItems.Add(item);
        }

        private static void RemoveLegacyMenuItem(Menu parent, string text)
        {
            MenuItem item = FindLegacyMenuItem(parent, text);
            if (item == null)
            {
                return;
            }

            parent.MenuItems.Remove(item);
            item.Dispose();
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
            private const int NativeQuantityMultiplyCommand = 28433;
            private const int NativeQuotaCodeMultiplyCommand = 28434;
            private const int NativeQuantityRemoveCommand = 28435;
            private const int NativeQuotaCodeRemoveCommand = 28436;
            private const int NativeUnitPriceMultiplyCommand = 28437;
            private const int NativeAdjustmentSetCommand = 28438;
            private const int NativeAdjustmentRemoveCommand = 28439;

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
                    if (command == NativeQuantityMultiplyCommand)
                    {
                        ApplyContextMenuAction(mainForm, false, true, "quantity", "multiply");
                        return true;
                    }
                    if (command == NativeQuantityRemoveCommand)
                    {
                        ApplyContextMenuAction(mainForm, false, true, "quantity", "remove");
                        return true;
                    }
                    if (command == NativeQuotaCodeMultiplyCommand)
                    {
                        ApplyContextMenuAction(mainForm, false, true, "quota_code", "multiply");
                        return true;
                    }
                    if (command == NativeQuotaCodeRemoveCommand)
                    {
                        ApplyContextMenuAction(mainForm, false, true, "quota_code", "remove");
                        return true;
                    }
                    if (command == NativeUnitPriceMultiplyCommand)
                    {
                        ApplyContextMenuAction(mainForm, false, true, "unit_price", "multiply");
                        return true;
                    }
                    if (command == NativeAdjustmentSetCommand)
                    {
                        ApplyContextMenuAction(mainForm, false, true, "adjustment", "set_adjustment");
                        return true;
                    }
                    if (command == NativeAdjustmentRemoveCommand)
                    {
                        ApplyContextMenuAction(mainForm, false, true, "adjustment", "remove_adjustment");
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
                IntPtr quantity = CreatePopupMenu();
                IntPtr quotaCode = CreatePopupMenu();
                IntPtr unitPrice = CreatePopupMenu();
                IntPtr adjustment = CreatePopupMenu();
                if (submenu == IntPtr.Zero || quantity == IntPtr.Zero || quotaCode == IntPtr.Zero ||
                    unitPrice == IntPtr.Zero || adjustment == IntPtr.Zero)
                {
                    return;
                }

                AppendMenu(quantity, MF_STRING, new UIntPtr((uint)NativeQuantityMultiplyCommand), "乘系数");
                AppendMenu(quantity, MF_STRING, new UIntPtr((uint)NativeQuantityRemoveCommand), "删系数");
                AppendPopupMenu(submenu, quantity, "工程数量");

                AppendMenu(quotaCode, MF_STRING, new UIntPtr((uint)NativeQuotaCodeMultiplyCommand), "乘系数");
                AppendMenu(quotaCode, MF_STRING, new UIntPtr((uint)NativeQuotaCodeRemoveCommand), "删系数");
                AppendPopupMenu(submenu, quotaCode, "定额编号");

                AppendMenu(unitPrice, MF_STRING, new UIntPtr((uint)NativeUnitPriceMultiplyCommand), "乘系数");
                AppendPopupMenu(submenu, unitPrice, "单价");

                AppendMenu(adjustment, MF_STRING, new UIntPtr((uint)NativeAdjustmentSetCommand), "调整输入");
                AppendMenu(adjustment, MF_STRING, new UIntPtr((uint)NativeAdjustmentRemoveCommand), "删除调整");
                AppendPopupMenu(submenu, adjustment, "定额调整");

                UIntPtr popupId = UIntPtr.Size == 8
                    ? new UIntPtr((ulong)submenu.ToInt64())
                    : new UIntPtr((uint)submenu.ToInt32());
                int index = Math.Max(0, FindNativeInsertIndex(menu));
                InsertMenu(menu, (uint)index, MF_BYPOSITION | MF_POPUP, popupId, "乘系数");
                Log("Native tree multiplier inserted. index=" + index.ToString(CultureInfo.InvariantCulture) + ", visible=" + NativeMenuText(menu));
            }

            private static void AppendPopupMenu(IntPtr parent, IntPtr child, string text)
            {
                UIntPtr popupId = UIntPtr.Size == 8
                    ? new UIntPtr((ulong)child.ToInt64())
                    : new UIntPtr((uint)child.ToInt32());
                AppendMenu(parent, MF_POPUP, popupId, text);
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

        private static void ApplyContextMenuAction(Form mainForm, bool isDeMenu, bool isTreeMenu, string target, string action)
        {
            if (mainForm == null || (!isDeMenu && !isTreeMenu))
            {
                return;
            }

            string title = ContextMenuActionTitle(action);
            string value;
            if (action == "multiply" || action == "remove")
            {
                FactorInfo factor = PromptFactor(mainForm, title);
                if (factor == null)
                {
                    return;
                }

                value = factor.Suffix;
            }
            else
            {
                value = PromptContextText(mainForm, title, action == "set_adjustment" ? "请输入调整：" : "请输入要删除的调整：");
                if (value == null)
                {
                    return;
                }
            }

            try
            {
                AgentSelectionSnapshot selection = CaptureAgentSelection(mainForm);
                AgentCommand command = BuildContextMenuCommand(target, action, value, isTreeMenu, selection.CurrentUnitId);
                AgentPlan plan;
                SqlConnection conn = GetOpenProjectConnection(mainForm);
                plan = BuildAgentPlan(conn, selection, new List<AgentCommand> { command });

                string warnings = plan.Warnings.Count == 0 ? "" : Environment.NewLine + Environment.NewLine + String.Join(Environment.NewLine, plan.Warnings.ToArray());
                if (plan.AffectedCount == 0)
                {
                    MessageBox.Show(mainForm, "没有需要执行的改动。" + warnings, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string message = ExecuteAgentPlan(mainForm, plan, null);
                MessageBox.Show(mainForm, message + warnings, title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                Log("Context menu action executed. Scope=" + (isTreeMenu ? "tree" : "quota") +
                    ", UnitId=" + selection.CurrentUnitId.ToString(CultureInfo.InvariantCulture) +
                    ", Target=" + target + ", Action=" + action + ", Affected=" + plan.AffectedCount.ToString(CultureInfo.InvariantCulture));
            }
            catch (AgentPlanException ex)
            {
                MessageBox.Show(mainForm, ex.Message, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                Log("Context menu action failed. Target=" + target + ", Action=" + action + " " + ex);
                MessageBox.Show(mainForm, "执行失败：" + ex.Message, title, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static AgentCommand BuildContextMenuCommand(string target, string action, string value, bool isTreeMenu, long currentUnitId)
        {
            AgentCommand command = new AgentCommand();
            command.Items = new List<string> { isTreeMenu ? AgentCurrentItemToken : AgentSelectedToken };
            command.IncludeChildren = isTreeMenu;
            command.Target = target;
            if (isTreeMenu)
            {
                if (currentUnitId <= 0)
                {
                    throw new AgentPlanException("无法识别当前单元。为避免误改其他单元，已中止操作。");
                }

                command.Units = new List<string> { currentUnitId.ToString(CultureInfo.InvariantCulture) };
            }

            if (action == "multiply")
            {
                FactorInfo factor;
                string error;
                if (!TryParseFactor(value, out factor, out error))
                {
                    throw new AgentPlanException(error);
                }

                command.Type = "multiply_quantity";
                command.Operator = factor.Operator;
                command.Factor = factor.Factor;
            }
            else if (action == "remove")
            {
                if (target == "unit_price")
                {
                    throw new AgentPlanException("单价不支持删系数。");
                }

                FactorInfo factor;
                string error;
                if (!TryParseFactor(value, out factor, out error))
                {
                    throw new AgentPlanException(error);
                }

                command.Type = "remove_text";
                command.RemoveText = factor.Suffix;
            }
            else if (action == "set_adjustment" && target == "adjustment")
            {
                command.Type = "set_adjustment";
                command.Mode = "set";
                command.Value = value;
            }
            else if (action == "remove_adjustment" && target == "adjustment")
            {
                command.Type = "remove_text";
                command.RemoveText = value;
            }
            else
            {
                throw new AgentPlanException("不支持的右键操作。");
            }

            return command;
        }

        private static string ContextMenuActionTitle(string action)
        {
            if (action == "remove") { return "删系数"; }
            if (action == "set_adjustment") { return "调整输入"; }
            if (action == "remove_adjustment") { return "删除调整"; }
            return "乘系数";
        }

        private static FactorInfo PromptFactor(IWin32Window owner, string title)
        {
            using (Form dialog = new Form())
            using (Label label = new Label())
            using (TextBox textBox = new TextBox())
            using (Button ok = new Button())
            using (Button cancel = new Button())
            {
                dialog.Text = title;
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

                    MessageBox.Show(owner, error, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    textBox.SelectAll();
                    textBox.Focus();
                }
            }

            return null;
        }

        private static string PromptContextText(IWin32Window owner, string title, string prompt)
        {
            using (Form dialog = new Form())
            using (Label label = new Label())
            using (TextBox textBox = new TextBox())
            using (Button ok = new Button())
            using (Button cancel = new Button())
            {
                dialog.Text = title;
                dialog.FormBorderStyle = FormBorderStyle.FixedDialog;
                dialog.StartPosition = FormStartPosition.CenterParent;
                dialog.MinimizeBox = false;
                dialog.MaximizeBox = false;
                dialog.ClientSize = new System.Drawing.Size(380, 110);

                label.Text = prompt;
                label.Left = 12;
                label.Top = 15;
                label.Width = 150;

                textBox.Left = 160;
                textBox.Top = 12;
                textBox.Width = 200;

                ok.Text = "确定";
                ok.Left = 204;
                ok.Top = 62;
                ok.Width = 75;
                ok.DialogResult = DialogResult.OK;

                cancel.Text = "取消";
                cancel.Left = 285;
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
                    string value = textBox.Text == null ? "" : textBox.Text.Trim();
                    if (value.Length > 0)
                    {
                        return value;
                    }

                    MessageBox.Show(owner, "输入内容不能为空。", title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
    }
}
