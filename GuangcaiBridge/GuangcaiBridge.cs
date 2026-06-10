using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;
using Microsoft.Win32;

internal static class GuangcaiBridgeProgram
{
    private const int SwHide = 0;
    private const int SwRestore = 9;
    private const uint MouseEventLeftDown = 0x0002;
    private const uint MouseEventLeftUp = 0x0004;
    private const uint KeyEventKeyUp = 0x0002;
    private const ushort VkControl = 0x11;
    private const ushort VkReturn = 0x0D;
    private const ushort VkA = 0x41;
    private const ushort VkV = 0x56;
    private const ushort VkBack = 0x08;
    private const uint KeyEventUnicode = 0x0004;
    private const uint SwpNoActivate = 0x0010;
    private static readonly object CommandLock = new object();
    private static readonly AutoResetEvent CommandReady = new AutoResetEvent(false);
    private static BridgeCommand pendingCommand;
    private static long latestCommandSequence;
    private static bool shuttingDown;
    private static int parentProcessId;
    private static bool nativeAdapterLogged;

    private delegate bool EnumWindowsProc(IntPtr window, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc callback, IntPtr parameter);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr window);

    [DllImport("user32.dll")]
    private static extern bool GetWindowRect(IntPtr window, out NativeRect rect);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr window);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr window, StringBuilder className, int maxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr window, StringBuilder text, int maxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowTextLength(IntPtr window);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr window, int command);

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(
        IntPtr window,
        IntPtr insertAfter,
        int x,
        int y,
        int width,
        int height,
        uint flags);

    [DllImport("user32.dll")]
    private static extern uint SendInput(uint count, NativeInput[] inputs, int size);

    [DllImport("user32.dll")]
    private static extern bool SetCursorPos(int x, int y);

    [DllImport("user32.dll")]
    private static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extraInfo);

    [STAThread]
    private static int Main(string[] args)
    {
        try
        {
            parentProcessId = ParseParentProcessId(args);
            Thread reader = new Thread(ReadCommands);
            reader.IsBackground = true;
            reader.Name = "GuangcaiBridge.CommandReader";
            reader.Start();

            Log("Bridge started. parent=" + parentProcessId.ToString(CultureInfo.InvariantCulture));
            while (!shuttingDown && IsParentAlive())
            {
                CommandReady.WaitOne(500);
                BridgeCommand command;
                lock (CommandLock)
                {
                    command = pendingCommand;
                    pendingCommand = null;
                }

                if (command != null)
                {
                    ExecuteCommand(command);
                }
            }
        }
        catch (Exception ex)
        {
            Log("Bridge fatal error: " + ex);
            return 1;
        }

        Log("Bridge stopped.");
        return 0;
    }

    private static void ReadCommands()
    {
        try
        {
            string line;
            while ((line = Console.ReadLine()) != null)
            {
                BridgeCommand command = ParseCommand(line);
                if (command == null)
                {
                    continue;
                }

                lock (CommandLock)
                {
                    command.Sequence = ++latestCommandSequence;
                    pendingCommand = command;
                }
                CommandReady.Set();
            }
        }
        catch (Exception ex)
        {
            Log("Read command failed: " + ex.Message);
        }
        finally
        {
            shuttingDown = true;
            CommandReady.Set();
        }
    }

    private static BridgeCommand ParseCommand(string line)
    {
        string[] parts = (line ?? String.Empty).Split('\t');
        if (parts.Length == 0)
        {
            return null;
        }

        BridgeCommand command = new BridgeCommand();
        command.Name = parts[0];
        if (parts.Length > 1 && !String.IsNullOrEmpty(parts[1]))
        {
            try
            {
                command.MaterialName = Encoding.UTF8.GetString(Convert.FromBase64String(parts[1]));
            }
            catch
            {
                return null;
            }
        }

        long ownerValue;
        if (parts.Length > 2 &&
            Int64.TryParse(parts[2], NumberStyles.Integer, CultureInfo.InvariantCulture, out ownerValue))
        {
            command.OwnerWindow = new IntPtr(ownerValue);
        }

        return command;
    }

    private static void ExecuteCommand(BridgeCommand command)
    {
        try
        {
            IntPtr window = EnsureGuangcaiWindow(command);
            if (IsSuperseded(command))
            {
                Log("Skipped superseded Guangcai command.");
                return;
            }

            if (window == IntPtr.Zero)
            {
                if (!String.IsNullOrWhiteSpace(command.MaterialName))
                {
                    CopyMaterialName(command.MaterialName);
                }
                Log("Guangcai window unavailable; material copied to clipboard.");
                return;
            }

            ActivateAndDock(window, command.OwnerWindow);
            if (String.Equals(command.Name, "QUERY", StringComparison.OrdinalIgnoreCase) &&
                !String.IsNullOrWhiteSpace(command.MaterialName))
            {
                LogNativeAdapterStatus();
                if (!TryQueryWithAutomation(window, command.MaterialName) &&
                    !TryQueryWithLayoutFallback(window, command.MaterialName))
                {
                    CopyMaterialName(command.MaterialName);
                    ActivateAndDock(window, command.OwnerWindow);
                    Log("UI Automation query unavailable; material copied to clipboard.");
                }
            }
        }
        catch (Exception ex)
        {
            Log("Execute command failed: " + ex);
            if (!String.IsNullOrWhiteSpace(command.MaterialName))
            {
                CopyMaterialName(command.MaterialName);
            }
        }
    }

    private static IntPtr EnsureGuangcaiWindow(BridgeCommand command)
    {
        IntPtr window = FindGuangcaiWindow();
        if (window != IntPtr.Zero)
        {
            return window;
        }

        string clientPath = FindClientPath();
        if (String.IsNullOrEmpty(clientPath))
        {
            Log("JGBClient.exe was not found.");
            return IntPtr.Zero;
        }

        if (!LaunchGuangcaiProcess(clientPath, "-showmw"))
        {
            return IntPtr.Zero;
        }

        Log("JGBClient show-window request sent.");
        for (int i = 0; i < 10; i++)
        {
            if (IsSuperseded(command))
            {
                return IntPtr.Zero;
            }
            Thread.Sleep(500);
            window = FindGuangcaiWindow();
            if (window != IntPtr.Zero)
            {
                return window;
            }
        }

        string viewPath = Path.Combine(Path.GetDirectoryName(clientPath), "JGBView.exe");
        if (File.Exists(viewPath))
        {
            LaunchGuangcaiProcess(viewPath, "-showmw");
            Log("JGBView show-window request sent.");
        }

        for (int i = 0; i < 30; i++)
        {
            if (IsSuperseded(command))
            {
                return IntPtr.Zero;
            }
            Thread.Sleep(500);
            window = FindGuangcaiWindow();
            if (window != IntPtr.Zero)
            {
                return window;
            }
        }

        if (TryOpenFromNotificationArea())
        {
            Log("Guangcai notification icon invoked.");
            for (int i = 0; i < 20; i++)
            {
                if (IsSuperseded(command))
                {
                    return IntPtr.Zero;
                }
                Thread.Sleep(500);
                window = FindGuangcaiWindow();
                if (window != IntPtr.Zero)
                {
                    return window;
                }
            }
        }

        return IntPtr.Zero;
    }

    private static bool IsSuperseded(BridgeCommand command)
    {
        lock (CommandLock)
        {
            return command != null && command.Sequence < latestCommandSequence;
        }
    }

    private static bool LaunchGuangcaiProcess(string path, string arguments)
    {
        try
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = path;
            startInfo.Arguments = arguments;
            startInfo.WorkingDirectory = Path.GetDirectoryName(path);
            startInfo.UseShellExecute = true;
            Process.Start(startInfo);
            return true;
        }
        catch (Exception ex)
        {
            Log("Launch Guangcai process failed: " + ex.Message);
            return false;
        }
    }

    private static IntPtr FindGuangcaiWindow()
    {
        HashSet<int> processIds = new HashSet<int>();
        AddProcessIds(processIds, "JGBClient");
        AddProcessIds(processIds, "JGBView");
        AddProcessIds(processIds, "JGBLoading");
        if (processIds.Count == 0)
        {
            return IntPtr.Zero;
        }

        IntPtr bestWindow = IntPtr.Zero;
        long bestArea = 0;
        string bestWindowDiagnostic = null;
        List<string> diagnostics = new List<string>();
        EnumWindows(delegate(IntPtr window, IntPtr parameter)
        {
            uint processId;
            GetWindowThreadProcessId(window, out processId);
            if (!processIds.Contains((int)processId))
            {
                return true;
            }

            NativeRect rect;
            if (!GetWindowRect(window, out rect))
            {
                return true;
            }

            long width = Math.Max(0, rect.Right - rect.Left);
            long height = Math.Max(0, rect.Bottom - rect.Top);
            long area = width * height;
            bool visible = IsWindowVisible(window);
            string className = GetWindowClassName(window);
            string title = GetWindowTitle(window);
            if (IsTrayMessageWindow(className, title))
            {
                if (visible)
                {
                    ShowWindow(window, SwHide);
                    Log("Hidden Guangcai tray message window. class=" + className + ", title=" + title);
                }
                return true;
            }

            if (!IsLikelyGuangcaiMainWindow(className, title, area, visible))
            {
                if (diagnostics.Count < 20)
                {
                    diagnostics.Add(BuildWindowDiagnostic(processId, visible, area, className, title));
                }
                return true;
            }

            if (diagnostics.Count < 20)
            {
                diagnostics.Add(BuildWindowDiagnostic(processId, visible, area, className, title));
            }

            if (visible)
            {
                if (area > bestArea)
                {
                    bestArea = area;
                    bestWindow = window;
                    bestWindowDiagnostic = BuildWindowDiagnostic(processId, visible, area, className, title);
                }
            }
            return true;
        }, IntPtr.Zero);

        if (bestWindow != IntPtr.Zero)
        {
            Log("Selected visible Guangcai window: " + bestWindowDiagnostic);
            return bestWindow;
        }

        if (diagnostics.Count > 0)
        {
            Log("Guangcai window candidates: " + String.Join(" | ", diagnostics.ToArray()));
        }
        return IntPtr.Zero;
    }

    private static string BuildWindowDiagnostic(uint processId, bool visible, long area, string className, string title)
    {
        return "pid=" + processId.ToString(CultureInfo.InvariantCulture) +
            ", visible=" + visible.ToString() +
            ", area=" + area.ToString(CultureInfo.InvariantCulture) +
            ", class=" + className +
            ", title=" + title;
    }

    private static string GetWindowClassName(IntPtr window)
    {
        StringBuilder className = new StringBuilder(256);
        GetClassName(window, className, className.Capacity);
        return className.ToString();
    }

    private static string GetWindowTitle(IntPtr window)
    {
        int length = Math.Min(Math.Max(256, GetWindowTextLength(window) + 1), 1024);
        StringBuilder title = new StringBuilder(length);
        GetWindowText(window, title, title.Capacity);
        return title.ToString();
    }

    private static bool IsTrayMessageWindow(string className, string title)
    {
        string text = ((className ?? String.Empty) + " " + (title ?? String.Empty)).ToLowerInvariant();
        return text.IndexOf("qtrayiconmessagewindow", StringComparison.Ordinal) >= 0 ||
            text.IndexOf("trayiconmessage", StringComparison.Ordinal) >= 0;
    }

    private static bool IsLikelyGuangcaiMainWindow(string className, string title, long area, bool visible)
    {
        if (area < 100000)
        {
            return false;
        }

        if (ContainsAny(className ?? String.Empty, "QTrayIconMessageWindow"))
        {
            return false;
        }

        if (ContainsAny(title ?? String.Empty, "\u5e7f\u6750\u52a9\u624b", "\u5e7f\u6750"))
        {
            return true;
        }

        if (visible)
        {
            return true;
        }

        return false;
    }

    private static void AddProcessIds(HashSet<int> ids, string processName)
    {
        try
        {
            foreach (Process process in Process.GetProcessesByName(processName))
            {
                ids.Add(process.Id);
                process.Dispose();
            }
        }
        catch
        {
        }
    }

    private static string FindClientPath()
    {
        string runningPath = FindRunningClientPath();
        if (!String.IsNullOrEmpty(runningPath))
        {
            return runningPath;
        }

        string registryPath = FindInstallLocation(Registry.CurrentUser);
        if (String.IsNullOrEmpty(registryPath))
        {
            registryPath = FindInstallLocation(Registry.LocalMachine);
        }
        if (!String.IsNullOrEmpty(registryPath))
        {
            string client = Path.Combine(registryPath, "JGBClient.exe");
            if (File.Exists(client))
            {
                return client;
            }
        }

        string[] candidates =
        {
            "D:\\\u8f6f\u4ef6\u4e0b\u8f7d\\\u5e7f\u6750\u52a9\u624b\\JGBClient.exe",
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Glodon", "JGBClient.exe"),
            Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Glodon", "JGBClient.exe")
        };
        foreach (string candidate in candidates)
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return null;
    }

    private static string FindRunningClientPath()
    {
        try
        {
            foreach (Process process in Process.GetProcessesByName("JGBClient"))
            {
                try
                {
                    string path = process.MainModule.FileName;
                    if (File.Exists(path))
                    {
                        return path;
                    }
                }
                catch
                {
                }
                finally
                {
                    process.Dispose();
                }
            }
        }
        catch
        {
        }

        return null;
    }

    private static string FindInstallLocation(RegistryKey root)
    {
        string[] uninstallRoots =
        {
            @"Software\Microsoft\Windows\CurrentVersion\Uninstall",
            @"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
        };

        foreach (string uninstallRoot in uninstallRoots)
        {
            try
            {
                using (RegistryKey rootKey = root.OpenSubKey(uninstallRoot))
                {
                    if (rootKey == null)
                    {
                        continue;
                    }

                    foreach (string subKeyName in rootKey.GetSubKeyNames())
                    {
                        using (RegistryKey subKey = rootKey.OpenSubKey(subKeyName))
                        {
                            string displayName = Convert.ToString(subKey.GetValue("DisplayName"), CultureInfo.InvariantCulture);
                            if (displayName.IndexOf("\u5e7f\u6750\u52a9\u624b", StringComparison.OrdinalIgnoreCase) < 0)
                            {
                                continue;
                            }

                            string location = Convert.ToString(subKey.GetValue("InstallLocation"), CultureInfo.InvariantCulture);
                            if (!String.IsNullOrWhiteSpace(location) && Directory.Exists(location))
                            {
                                return location.Trim();
                            }
                        }
                    }
                }
            }
            catch
            {
            }
        }

        return null;
    }

    private static void ActivateAndDock(IntPtr window, IntPtr ownerWindow)
    {
        ShowWindow(window, SwRestore);
        if (ownerWindow != IntPtr.Zero)
        {
            NativeRect ownerRect;
            NativeRect targetRect;
            if (GetWindowRect(ownerWindow, out ownerRect) && GetWindowRect(window, out targetRect))
            {
                int width = Math.Max(720, targetRect.Right - targetRect.Left);
                int height = Math.Max(500, targetRect.Bottom - targetRect.Top);
                Screen screen = Screen.FromRectangle(new System.Drawing.Rectangle(
                    ownerRect.Left,
                    ownerRect.Top,
                    Math.Max(1, ownerRect.Right - ownerRect.Left),
                    Math.Max(1, ownerRect.Bottom - ownerRect.Top)));
                System.Drawing.Rectangle work = screen.WorkingArea;
                int x = ownerRect.Right + 6;
                int y = Math.Max(work.Top, ownerRect.Top);
                if (x + width > work.Right)
                {
                    x = Math.Max(work.Left, ownerRect.Left - width - 6);
                }
                if (x + width > work.Right || x < work.Left)
                {
                    x = Math.Max(work.Left, work.Right - width);
                    y = Math.Max(work.Top, work.Bottom - height);
                }

                width = Math.Min(width, work.Width);
                height = Math.Min(height, work.Height);
                SetWindowPos(window, IntPtr.Zero, x, y, width, height, SwpNoActivate);
            }
        }

        SetForegroundWindow(window);
    }

    private static bool TryQueryWithAutomation(IntPtr window, string materialName)
    {
        try
        {
            AutomationElement root = AutomationElement.FromHandle(window);
            if (root == null)
            {
                return false;
            }

            AutomationElement input = FindSearchInput(root);
            if (input == null)
            {
                return false;
            }

            object patternObject;
            if (!input.TryGetCurrentPattern(ValuePattern.Pattern, out patternObject))
            {
                return false;
            }

            ValuePattern valuePattern = patternObject as ValuePattern;
            if (valuePattern == null || valuePattern.Current.IsReadOnly)
            {
                return false;
            }

            valuePattern.SetValue(materialName);
            AutomationElement button = FindSearchButton(root);
            if (button != null && button.TryGetCurrentPattern(InvokePattern.Pattern, out patternObject))
            {
                InvokePattern invokePattern = patternObject as InvokePattern;
                if (invokePattern != null)
                {
                    invokePattern.Invoke();
                    Log("Material query submitted through UI Automation.");
                    return true;
                }
            }

            input.SetFocus();
            AutomationElement focused = AutomationElement.FocusedElement;
            if (focused != null && SameElement(input, focused) && SendEnter())
            {
                Log("Material query submitted with Enter.");
                return true;
            }

            // Some Guangcai versions refresh results immediately after ValuePattern.SetValue.
            Log("Material name set; no explicit search control was exposed.");
            return true;
        }
        catch (Exception ex)
        {
            Log("UI Automation query failed: " + ex.Message);
            return false;
        }
    }

    private static bool TryQueryWithLayoutFallback(IntPtr window, string materialName)
    {
        try
        {
            NativeRect rect;
            if (!GetWindowRect(window, out rect))
            {
                return false;
            }

            int width = rect.Right - rect.Left;
            int height = rect.Bottom - rect.Top;
            if (width < 600 || height < 350)
            {
                Log("Layout fallback skipped because window is too small. width=" +
                    width.ToString(CultureInfo.InvariantCulture) +
                    ", height=" + height.ToString(CultureInfo.InvariantCulture));
                return false;
            }

            ShowWindow(window, SwRestore);
            SetForegroundWindow(window);
            Thread.Sleep(200);
            CopyMaterialName(materialName);

            int searchY = rect.Top + Math.Max(86, Math.Min(112, height / 7));
            int inputX = rect.Right - Math.Max(110, Math.Min(185, width / 6));
            int buttonX = rect.Right - 24;
            Log("Layout fallback search point. inputX=" +
                inputX.ToString(CultureInfo.InvariantCulture) +
                ", buttonX=" + buttonX.ToString(CultureInfo.InvariantCulture) +
                ", y=" + searchY.ToString(CultureInfo.InvariantCulture));

            if (!ClickPoint(inputX, searchY))
            {
                Log("Layout fallback input click failed.");
                return false;
            }

            Thread.Sleep(120);
            SendCtrlKey(VkA);
            Thread.Sleep(80);
            SendKey(VkBack);
            Thread.Sleep(80);
            if (!SendUnicodeText(materialName))
            {
                SendCtrlKey(VkV);
            }
            Thread.Sleep(120);
            SendKey(VkReturn);
            Thread.Sleep(120);
            ClickPoint(buttonX, searchY);
            Log("Material query submitted with layout fallback.");
            return true;
        }
        catch (Exception ex)
        {
            Log("Layout fallback query failed: " + ex.Message);
            return false;
        }
    }

    private static bool TryOpenFromNotificationArea()
    {
        try
        {
            AutomationElement root = AutomationElement.RootElement;
            if (TryInvokeNamedButton(root, "\u5e7f\u6750\u52a9\u624b", "jgb"))
            {
                return true;
            }

            if (TryInvokeNamedButton(root, "\u663e\u793a\u9690\u85cf\u7684\u56fe\u6807", "show hidden icons"))
            {
                Thread.Sleep(500);
                return TryInvokeNamedButton(AutomationElement.RootElement, "\u5e7f\u6750\u52a9\u624b", "jgb");
            }
        }
        catch (Exception ex)
        {
            Log("Notification icon automation failed: " + ex.Message);
        }

        return false;
    }

    private static bool TryInvokeNamedButton(AutomationElement root, params string[] names)
    {
        AutomationElementCollection buttons = root.FindAll(
            TreeScope.Descendants,
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
        foreach (AutomationElement button in buttons)
        {
            try
            {
                string text = (button.Current.Name + " " + button.Current.AutomationId).ToLowerInvariant();
                if (!ContainsAny(text, names))
                {
                    continue;
                }

                object patternObject;
                if (button.TryGetCurrentPattern(InvokePattern.Pattern, out patternObject))
                {
                    InvokePattern invokePattern = patternObject as InvokePattern;
                    if (invokePattern != null)
                    {
                        invokePattern.Invoke();
                        return true;
                    }
                }

            }
            catch
            {
            }
        }

        return false;
    }

    private static AutomationElement FindSearchInput(AutomationElement root)
    {
        AutomationElementCollection edits = root.FindAll(
            TreeScope.Descendants,
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit));
        AutomationElement best = null;
        int bestScore = Int32.MinValue;
        foreach (AutomationElement edit in edits)
        {
            int score = ScoreSearchInput(edit);
            if (score > bestScore)
            {
                bestScore = score;
                best = edit;
            }
        }

        return bestScore >= 60 ? best : null;
    }

    private static int ScoreSearchInput(AutomationElement element)
    {
        try
        {
            if (!element.Current.IsEnabled || element.Current.IsOffscreen)
            {
                return Int32.MinValue;
            }

            string text = (element.Current.Name + " " +
                element.Current.AutomationId + " " +
                element.Current.ClassName).ToLowerInvariant();
            if (ContainsAny(text, "\u5bc6\u7801", "\u9a8c\u8bc1\u7801", "\u624b\u673a", "\u8d26\u53f7", "password", "login"))
            {
                return Int32.MinValue;
            }

            int score = 0;
            if (ContainsAny(text, "\u6750\u6599\u540d\u79f0", "\u6750\u6599", "material"))
            {
                score += 120;
            }
            if (ContainsAny(text, "\u641c\u7d22", "\u67e5\u8be2", "search", "query"))
            {
                score += 80;
            }
            if (text.IndexOf("qlineedit", StringComparison.Ordinal) >= 0)
            {
                score += 20;
            }

            System.Windows.Rect rect = element.Current.BoundingRectangle;
            if (rect.Width >= 160 && rect.Height >= 18 && rect.Height <= 80)
            {
                score += 20;
            }

            object pattern;
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out pattern))
            {
                score += 20;
            }
            return score;
        }
        catch
        {
            return Int32.MinValue;
        }
    }

    private static AutomationElement FindSearchButton(AutomationElement root)
    {
        AutomationElementCollection buttons = root.FindAll(
            TreeScope.Descendants,
            new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
        foreach (AutomationElement button in buttons)
        {
            try
            {
                string text = (button.Current.Name + " " + button.Current.AutomationId).ToLowerInvariant();
                if (ContainsAny(text, "\u641c\u7d22", "\u67e5\u8be2", "search", "query") &&
                    button.Current.IsEnabled && !button.Current.IsOffscreen)
                {
                    return button;
                }
            }
            catch
            {
            }
        }

        return null;
    }

    private static bool SameElement(AutomationElement left, AutomationElement right)
    {
        try
        {
            int[] leftId = left.GetRuntimeId();
            int[] rightId = right.GetRuntimeId();
            if (leftId.Length != rightId.Length)
            {
                return false;
            }
            for (int i = 0; i < leftId.Length; i++)
            {
                if (leftId[i] != rightId[i])
                {
                    return false;
                }
            }
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool SendEnter()
    {
        NativeInput down = NativeInput.Keyboard(VkReturn, 0);
        NativeInput up = NativeInput.Keyboard(VkReturn, KeyEventKeyUp);
        return SendInput(2, new[] { down, up }, Marshal.SizeOf(typeof(NativeInput))) == 2;
    }

    private static bool ClickPoint(int x, int y)
    {
        if (!SetCursorPos(x, y))
        {
            Log("SetCursorPos failed.");
            return false;
        }

        NativeInput down = NativeInput.Mouse(MouseEventLeftDown);
        NativeInput up = NativeInput.Mouse(MouseEventLeftUp);
        if (SendInput(2, new[] { down, up }, Marshal.SizeOf(typeof(NativeInput))) == 2)
        {
            return true;
        }

        mouse_event(MouseEventLeftDown, 0, 0, 0, UIntPtr.Zero);
        Thread.Sleep(30);
        mouse_event(MouseEventLeftUp, 0, 0, 0, UIntPtr.Zero);
        return true;
    }

    private static bool SendKey(ushort virtualKey)
    {
        NativeInput down = NativeInput.Keyboard(virtualKey, 0);
        NativeInput up = NativeInput.Keyboard(virtualKey, KeyEventKeyUp);
        return SendInput(2, new[] { down, up }, Marshal.SizeOf(typeof(NativeInput))) == 2;
    }

    private static bool SendCtrlKey(ushort virtualKey)
    {
        NativeInput ctrlDown = NativeInput.Keyboard(VkControl, 0);
        NativeInput keyDown = NativeInput.Keyboard(virtualKey, 0);
        NativeInput keyUp = NativeInput.Keyboard(virtualKey, KeyEventKeyUp);
        NativeInput ctrlUp = NativeInput.Keyboard(VkControl, KeyEventKeyUp);
        return SendInput(4, new[] { ctrlDown, keyDown, keyUp, ctrlUp }, Marshal.SizeOf(typeof(NativeInput))) == 4;
    }

    private static bool SendUnicodeText(string text)
    {
        if (String.IsNullOrEmpty(text))
        {
            return true;
        }

        List<NativeInput> inputs = new List<NativeInput>();
        foreach (char ch in text)
        {
            inputs.Add(NativeInput.Unicode(ch, 0));
            inputs.Add(NativeInput.Unicode(ch, KeyEventKeyUp));
        }

        return SendInput((uint)inputs.Count, inputs.ToArray(), Marshal.SizeOf(typeof(NativeInput))) == inputs.Count;
    }

    private static bool ContainsAny(string value, params string[] candidates)
    {
        foreach (string candidate in candidates)
        {
            if (value.IndexOf(candidate, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }
        }
        return false;
    }

    private static void CopyMaterialName(string materialName)
    {
        for (int attempt = 0; attempt < 3; attempt++)
        {
            try
            {
                Clipboard.SetText(materialName);
                return;
            }
            catch
            {
                Thread.Sleep(100);
            }
        }
    }

    private static void LogNativeAdapterStatus()
    {
        if (nativeAdapterLogged)
        {
            return;
        }

        nativeAdapterLogged = true;
        string clientPath = FindClientPath();
        string baseDir = String.IsNullOrEmpty(clientPath) ? null : Path.GetDirectoryName(clientPath);
        string pluginPath = String.IsNullOrEmpty(baseDir) ? null : Path.Combine(baseDir, "Bin", "JGBPlugin2.dll");
        if (!String.IsNullOrEmpty(pluginPath) && File.Exists(pluginPath))
        {
            Log("Native JGBPlugin2.dll detected; undocumented ABI is not invoked. Using isolated UI Automation mode.");
        }
        else
        {
            Log("Native Guangcai host plugin not found. Using isolated UI Automation mode.");
        }
    }

    private static int ParseParentProcessId(string[] args)
    {
        for (int i = 0; i + 1 < args.Length; i++)
        {
            int value;
            if (args[i] == "--parent" &&
                Int32.TryParse(args[i + 1], NumberStyles.Integer, CultureInfo.InvariantCulture, out value))
            {
                return value;
            }
        }
        return 0;
    }

    private static bool IsParentAlive()
    {
        if (parentProcessId <= 0)
        {
            return true;
        }
        try
        {
            Process parent = Process.GetProcessById(parentProcessId);
            bool alive = !parent.HasExited;
            parent.Dispose();
            return alive;
        }
        catch
        {
            return false;
        }
    }

    private static void Log(string message)
    {
        try
        {
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "GuangcaiBridge.log");
            File.AppendAllText(
                path,
                DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) +
                " " + message + Environment.NewLine);
        }
        catch
        {
        }
    }

    private sealed class BridgeCommand
    {
        public long Sequence;
        public string Name;
        public string MaterialName;
        public IntPtr OwnerWindow;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeRect
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeInput
    {
        public uint Type;
        public NativeInputUnion Union;

        public static NativeInput Keyboard(ushort virtualKey, uint flags)
        {
            NativeInput input = new NativeInput();
            input.Type = 1;
            input.Union.Keyboard = new NativeKeyboardInput();
            input.Union.Keyboard.VirtualKey = virtualKey;
            input.Union.Keyboard.Flags = flags;
            return input;
        }

        public static NativeInput Unicode(char value, uint flags)
        {
            NativeInput input = new NativeInput();
            input.Type = 1;
            input.Union.Keyboard = new NativeKeyboardInput();
            input.Union.Keyboard.ScanCode = value;
            input.Union.Keyboard.Flags = flags | KeyEventUnicode;
            return input;
        }

        public static NativeInput Mouse(uint flags)
        {
            NativeInput input = new NativeInput();
            input.Type = 0;
            input.Union.Mouse = new NativeMouseInput();
            input.Union.Mouse.Flags = flags;
            return input;
        }
    }

    [StructLayout(LayoutKind.Explicit)]
    private struct NativeInputUnion
    {
        [FieldOffset(0)]
        public NativeMouseInput Mouse;

        [FieldOffset(0)]
        public NativeKeyboardInput Keyboard;

        [FieldOffset(0)]
        public NativeHardwareInput Hardware;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeMouseInput
    {
        public int X;
        public int Y;
        public uint MouseData;
        public uint Flags;
        public uint Time;
        public IntPtr ExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeKeyboardInput
    {
        public ushort VirtualKey;
        public ushort ScanCode;
        public uint Flags;
        public uint Time;
        public IntPtr ExtraInfo;
    }

    [StructLayout(LayoutKind.Sequential)]
    private struct NativeHardwareInput
    {
        public uint Message;
        public ushort ParameterLow;
        public ushort ParameterHigh;
    }
}
