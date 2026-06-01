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
        private static readonly HashSet<Form> InstalledForms = new HashSet<Form>();
        private static readonly Dictionary<ContextMenuStrip, MenuInfo> MenuInfos = new Dictionary<ContextMenuStrip, MenuInfo>();
        private static readonly Dictionary<ContextMenu, MenuInfo> LegacyMenuInfos = new Dictionary<ContextMenu, MenuInfo>();
        private static readonly Dictionary<string, System.Drawing.Image> MenuIconCache = new Dictionary<string, System.Drawing.Image>(StringComparer.OrdinalIgnoreCase);
        private static bool idleHooked;
        private readonly Timer installTimer;
        private readonly object owner;
        private Form installedMainForm;

        private sealed class MenuInfo
        {
            public Form MainForm;
            public string Name;
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
                AddExcelLinkItemsIfMatched(menu);
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
            AddExcelLinkItemsIfMatched(menu);
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

        private static decimal EvaluateDecimal(string expression)
        {
            if (String.IsNullOrWhiteSpace(expression))
            {
                throw new InvalidOperationException("表达式为空。");
            }

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
            if (value == null || value == DBNull.Value)
            {
                throw new InvalidOperationException("表达式没有计算结果。");
            }

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
