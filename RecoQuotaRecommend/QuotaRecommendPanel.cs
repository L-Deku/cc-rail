using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace RecoQuotaRecommend
{
    public sealed class QuotaRecommendPanel : Form
    {
        private const long MaxLogBytes = 5L * 1024L * 1024L;
        private const int LogBackupCount = 3;
        private static readonly object LogLock = new object();
        private static readonly HashSet<Form> InstalledForms = new HashSet<Form>();
        private static readonly Dictionary<ContextMenuStrip, MenuInfo> MenuInfos = new Dictionary<ContextMenuStrip, MenuInfo>();
        private static readonly Dictionary<Form, RecommendDialog> RecommendDialogs = new Dictionary<Form, RecommendDialog>();
        private static Image recommendMenuIcon;
        private static bool idleHooked;
        private static bool consumeCryptoBridgeStarted;
        private static System.Windows.Forms.Timer consumeCryptoBridgeTimer;
        private static string consumeCryptoBridgeLastRequest;

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
            StartConsumeCryptoBridge();
            // 2024 迁移定额兼容层已整体移除：项目设置勾选迁移书号后，
            // 查询/输入/计算全部由主程序原生流程承担（见 AGENTS.md 迁移经验）。
            Application.Idle += delegate
            {
                try
                {
                    Form mainForm = FindMainForm();
                    if (mainForm != null && !InstalledForms.Contains(mainForm))
                    {
                        Install(mainForm);
                    }
                }
                catch (Exception ex)
                {
                    Log("Idle install failed: " + ex);
                }
            };
        }

        private static void StartConsumeCryptoBridge()
        {
            if (consumeCryptoBridgeStarted)
            {
                return;
            }

            consumeCryptoBridgeStarted = true;
            consumeCryptoBridgeTimer = new System.Windows.Forms.Timer();
            consumeCryptoBridgeTimer.Interval = 3000;
            consumeCryptoBridgeTimer.Tick += delegate { ProcessConsumeCryptoBridgeRequests(); };
            consumeCryptoBridgeTimer.Start();
            Log("Consume crypto bridge started.");
        }

        private static void ProcessConsumeCryptoBridgeRequests()
        {
            try
            {
                string dataDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "RecoQuotaData");
                string requestPath = Path.Combine(dataDir, "consume-encrypt-requests.tsv");
                if (!File.Exists(requestPath))
                {
                    return;
                }

                FileInfo requestInfo = new FileInfo(requestPath);
                string requestSignature = requestInfo.LastWriteTimeUtc.Ticks.ToString(CultureInfo.InvariantCulture) + ":" + requestInfo.Length.ToString(CultureInfo.InvariantCulture);
                if (String.Equals(requestSignature, consumeCryptoBridgeLastRequest, StringComparison.Ordinal))
                {
                    return;
                }

                consumeCryptoBridgeLastRequest = requestSignature;
                Directory.CreateDirectory(dataDir);
                string responsePath = Path.Combine(dataDir, "consume-encrypt-responses.tsv");
                string tempPath = responsePath + ".tmp";
                string[] lines = File.ReadAllLines(requestPath, Encoding.UTF8);
                using (StreamWriter writer = new StreamWriter(tempPath, false, Encoding.UTF8))
                {
                    writer.WriteLine("# ConsumeCryptoBridgeV1");
                    int ok = 0;
                    int err = 0;
                    foreach (string line in lines)
                    {
                        if (String.IsNullOrWhiteSpace(line) || line.StartsWith("#", StringComparison.Ordinal))
                        {
                            continue;
                        }

                        string[] parts = line.Split('\t');
                        if (parts.Length < 3)
                        {
                            continue;
                        }

                        string book = parts[0];
                        string code = parts[1];
                        try
                        {
                            string plain = Encoding.UTF8.GetString(Convert.FromBase64String(parts[2]));
                            string encrypted = EncryptConsumeForCurrentApp(plain);
                            writer.WriteLine(book + "\t" + code + "\tOK\t" + Convert.ToBase64String(Encoding.UTF8.GetBytes(encrypted)));
                            ok++;
                        }
                        catch (Exception ex)
                        {
                            writer.WriteLine(book + "\t" + code + "\tERR\t" + Convert.ToBase64String(Encoding.UTF8.GetBytes(ex.GetBaseException().Message)));
                            err++;
                        }
                    }

                    writer.WriteLine("# ok=" + ok.ToString(CultureInfo.InvariantCulture) + " err=" + err.ToString(CultureInfo.InvariantCulture));
                }

                if (File.Exists(responsePath))
                {
                    File.Delete(responsePath);
                }

                File.Move(tempPath, responsePath);
                Log("Consume crypto bridge processed request: " + requestSignature);
            }
            catch (Exception ex)
            {
                Log("Consume crypto bridge failed: " + ex);
            }
        }

        private static string EncryptConsumeForCurrentApp(string plain)
        {
            Type securityType = null;
            foreach (Assembly assembly in AppDomain.CurrentDomain.GetAssemblies())
            {
                securityType = assembly.GetType("RecoNet.Security", false);
                if (securityType != null)
                {
                    break;
                }
            }

            if (securityType == null)
            {
                throw new InvalidOperationException("RecoNet.Security not found in current AppDomain.");
            }

            object security = Activator.CreateInstance(securityType, true);
            MethodInfo method = securityType.GetMethod("Encrypto", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, new[] { typeof(string) }, null);
            if (method == null)
            {
                throw new MissingMethodException("RecoNet.Security.Encrypto(string) not found.");
            }

            return Convert.ToString(method.Invoke(security, new object[] { plain }), CultureInfo.InvariantCulture);
        }

        private static void Install(Form mainForm)
        {
            QuotaInlineSearchFeature.Install(mainForm);
            ReferenceQuotaPoolFeature.Install(mainForm);
            int menus = InstallAllContextMenus(mainForm);
            if (menus == 0)
            {
                Log("Context menus not found.");
                return;
            }

            InstalledForms.Add(mainForm);
            mainForm.FormClosed += delegate { InstalledForms.Remove(mainForm); };
            Log("Quota recommend menu installed. menus=" + menus.ToString(CultureInfo.InvariantCulture));
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
                menu.Opened -= ContextMenuOpened;
                menu.Opened += ContextMenuOpened;
                MenuInfos[menu] = new MenuInfo { MainForm = mainForm, Name = field.Name };
                AddRecommendItemIfMatched(menu);
            }

            return count;
        }

        private static void ContextMenuOpening(object sender, CancelEventArgs e)
        {
            ContextMenuStrip menu = sender as ContextMenuStrip;
            AddRecommendItemIfMatched(menu);
            BeginAddRecommendItem(menu);
        }

        private static void ContextMenuOpened(object sender, EventArgs e)
        {
            ContextMenuStrip menu = sender as ContextMenuStrip;
            AddRecommendItemIfMatched(menu);
            BeginAddRecommendItem(menu);
        }

        private static void BeginAddRecommendItem(ContextMenuStrip menu)
        {
            if (menu == null)
            {
                return;
            }

            try
            {
                menu.BeginInvoke((MethodInvoker)delegate { AddRecommendItemIfMatched(menu); });
            }
            catch
            {
            }
        }

        private static void AddRecommendItemIfMatched(ContextMenuStrip menu)
        {
            if (menu == null || !MenuInfos.ContainsKey(menu))
            {
                return;
            }

            MenuInfo info = MenuInfos[menu];
            bool isQuotaMenu = info.Name == "contextMenuStripDE" || IsSource(menu, info.MainForm, "dataGridViewDE");
            if (!isQuotaMenu)
            {
                return;
            }

            AddRecommendItem(menu, info.MainForm);
        }

        private static void AddRecommendItem(ContextMenuStrip menu, Form mainForm)
        {
            ToolStripMenuItem item = FindMenuItem(menu, "\u63a8\u8350\u5b9a\u989d");
            if (item != null)
            {
                item.Visible = true;
                item.Available = true;
                item.Enabled = true;
                ApplyRecommendMenuIcon(item);
                return;
            }

            int insertIndex = Math.Min(2, menu.Items.Count);
            item = new ToolStripMenuItem("\u63a8\u8350\u5b9a\u989d");
            item.Visible = true;
            item.Available = true;
            item.Enabled = true;
            ApplyRecommendMenuIcon(item);
            item.Click += delegate { ShowRecommendDialog(mainForm); };
            menu.Items.Insert(insertIndex, item);
        }

        private static void ApplyRecommendMenuIcon(ToolStripMenuItem item)
        {
            if (item == null)
            {
                return;
            }

            Image icon = LoadRecommendMenuIcon();
            if (icon != null)
            {
                item.Image = icon;
                item.ImageScaling = ToolStripItemImageScaling.SizeToFit;
            }
        }

        private static Image LoadRecommendMenuIcon()
        {
            if (recommendMenuIcon != null)
            {
                return recommendMenuIcon;
            }

            try
            {
                string dir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location);
                string path = Path.Combine(dir, "RecoExpandPanelIcons", "recommend_quota.png");
                if (File.Exists(path))
                {
                    using (Image image = Image.FromFile(path))
                    {
                        recommendMenuIcon = new Bitmap(image);
                    }

                    return recommendMenuIcon;
                }
            }
            catch (Exception ex)
            {
                Log("Load recommend menu icon failed: " + ex.Message);
            }

            recommendMenuIcon = DrawRecommendMenuIcon();
            return recommendMenuIcon;
        }

        private static Image DrawRecommendMenuIcon()
        {
            Bitmap bitmap = new Bitmap(24, 24);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                graphics.Clear(Color.Transparent);
                using (SolidBrush paper = new SolidBrush(Color.FromArgb(244, 248, 251)))
                using (Pen border = new Pen(Color.FromArgb(104, 126, 144)))
                using (Pen gridPen = new Pen(Color.FromArgb(168, 184, 196)))
                using (SolidBrush excel = new SolidBrush(Color.FromArgb(70, 139, 73)))
                using (Pen excelBorder = new Pen(Color.FromArgb(46, 96, 49)))
                using (Pen xPen = new Pen(Color.White, 1.6f))
                using (SolidBrush mark = new SolidBrush(Color.FromArgb(79, 144, 84)))
                using (Pen markPen = new Pen(Color.FromArgb(57, 102, 61), 1.4f))
                {
                    graphics.FillRectangle(paper, 6, 2, 14, 18);
                    graphics.DrawRectangle(border, 6, 2, 14, 18);
                    graphics.DrawLine(gridPen, 9, 7, 17, 7);
                    graphics.DrawLine(gridPen, 9, 11, 17, 11);
                    graphics.DrawLine(gridPen, 9, 15, 17, 15);
                    graphics.DrawLine(gridPen, 12, 5, 12, 18);
                    graphics.DrawLine(gridPen, 16, 5, 16, 18);
                    graphics.FillRectangle(excel, 1, 8, 9, 10);
                    graphics.DrawRectangle(excelBorder, 1, 8, 9, 10);
                    xPen.StartCap = System.Drawing.Drawing2D.LineCap.Round;
                    xPen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                    graphics.DrawLine(xPen, 3, 10, 8, 16);
                    graphics.DrawLine(xPen, 8, 10, 3, 16);
                    Point[] star = new Point[]
                    {
                        new Point(19, 3),
                        new Point(21, 7),
                        new Point(23, 8),
                        new Point(21, 10),
                        new Point(21, 13),
                        new Point(18, 11),
                        new Point(15, 12),
                        new Point(16, 9),
                        new Point(14, 6),
                        new Point(17, 7)
                    };
                    graphics.FillPolygon(mark, star);
                    graphics.DrawPolygon(markPen, star);
                }
            }

            return bitmap;
        }

        private static void ShowRecommendDialog(Form mainForm)
        {
            try
            {
                RecommendDialog dialog;
                if (!RecommendDialogs.TryGetValue(mainForm, out dialog) || dialog == null || dialog.IsDisposed)
                {
                    dialog = new RecommendDialog(mainForm, GetSelectionText(mainForm));
                    RecommendDialogs[mainForm] = dialog;
                    dialog.FormClosed += delegate { RecommendDialogs.Remove(mainForm); };
                    dialog.Show(mainForm);
                }
                else
                {
                    dialog.Show();
                    dialog.Activate();
                    dialog.RefreshEntryScope();
                }
            }
            catch (Exception ex)
            {
                Log("Show recommend dialog failed: " + ex);
                MessageBox.Show(mainForm, ex.Message, "\u63a8\u8350\u5b9a\u989d", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string GetSelectionText(Form mainForm)
        {
            DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
            if (grid == null)
            {
                return "";
            }

            List<string> parts = new List<string>();
            foreach (DataGridViewCell cell in grid.SelectedCells)
            {
                string text = Convert.ToString(cell.FormattedValue, CultureInfo.CurrentCulture);
                if (!String.IsNullOrWhiteSpace(text))
                {
                    parts.Add(text.Trim());
                }
            }

            if (parts.Count == 0 && grid.CurrentRow != null)
            {
                foreach (DataGridViewCell cell in grid.CurrentRow.Cells)
                {
                    string text = Convert.ToString(cell.FormattedValue, CultureInfo.CurrentCulture);
                    if (!String.IsNullOrWhiteSpace(text))
                    {
                        parts.Add(text.Trim());
                    }
                }
            }

            return String.Join(" ", parts.Distinct().Take(12).ToArray());
        }

        private static ToolStripMenuItem FindMenuItem(ContextMenuStrip menu, string text)
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

        private static bool IsSource(ContextMenuStrip menu, Form mainForm, string fieldName)
        {
            Control source = menu.SourceControl;
            Control expected = GetField<Control>(mainForm, fieldName);
            return source != null && expected != null && Object.ReferenceEquals(source, expected);
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

        private static T GetField<T>(object target, string name) where T : class
        {
            if (target == null)
            {
                return null;
            }

            FieldInfo field = target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            return field == null ? null : field.GetValue(target) as T;
        }

        internal static void Log(string message)
        {
            try
            {
                string dir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location);
                string path = Path.Combine(dir, "RecoQuotaRecommend.log");
                string line = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " " + message + Environment.NewLine;
                lock (LogLock)
                {
                    RotateLogIfNeeded(path);
                    File.AppendAllText(path, line, Encoding.UTF8);
                }
            }
            catch
            {
            }
        }

        private static void RotateLogIfNeeded(string path)
        {
            if (!File.Exists(path) || new FileInfo(path).Length < MaxLogBytes)
            {
                return;
            }

            for (int i = LogBackupCount; i >= 1; i--)
            {
                string source = i == 1 ? path : path + "." + (i - 1).ToString(CultureInfo.InvariantCulture);
                string target = path + "." + i.ToString(CultureInfo.InvariantCulture);
                if (!File.Exists(source))
                {
                    continue;
                }
                if (File.Exists(target))
                {
                    File.Delete(target);
                }
                File.Move(source, target);
            }
        }
    }
}
