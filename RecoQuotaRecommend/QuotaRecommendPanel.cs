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
        private static bool idleHooked;
        private static bool consumeCryptoBridgeStarted;
        private static System.Windows.Forms.Timer consumeCryptoBridgeTimer;
        private static string consumeCryptoBridgeLastRequest;


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
            InstalledForms.Add(mainForm);
            mainForm.FormClosed += delegate { InstalledForms.Remove(mainForm); };
            Log("Quota features installed (recommend dialog removed).");
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
