using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace RecoNet
{
    internal enum GuangcaiBridgeState
    {
        NotStarted,
        Starting,
        Connected,
        Fallback,
        Unavailable
    }

    internal sealed class MaterialSelection
    {
        public string MaterialName { get; set; }
        public int RowIndex { get; set; }
    }

    internal sealed class GuangcaiPriceSelection
    {
        public decimal Price { get; set; }
        public bool IsTaxIncluded { get; set; }
        public string Unit { get; set; }
    }

    internal interface IGuangcaiBridge : IDisposable
    {
        GuangcaiBridgeState State { get; }
        bool EnsureStarted();
        void QueryMaterial(string materialName, IntPtr ownerWindow);
        bool TryReceiveSelectedPrice(out GuangcaiPriceSelection selection);
        void ActivateWindow(IntPtr ownerWindow);
    }

    public partial class FormPanel : Form
    {
        private static readonly Dictionary<Form, GuangcaiMaterialRuntime> GuangcaiRuntimes =
            new Dictionary<Form, GuangcaiMaterialRuntime>();
        private static IGuangcaiBridge sharedGuangcaiBridge;
        private static bool guangcaiExitHooked;

        private static void EnsureGuangcaiLinkRuntime(Form mainForm)
        {
            if (mainForm == null)
            {
                return;
            }

            List<Form> forms = new List<Form>();
            foreach (Form openForm in Application.OpenForms)
            {
                forms.Add(openForm);
            }

            foreach (Form form in forms)
            {
                if (form == null || form.IsDisposed || form.GetType().FullName != "RecoNet.FormLFFA")
                {
                    continue;
                }

                if (GuangcaiRuntimes.ContainsKey(form))
                {
                    continue;
                }

                DataGridView grid = GetField<DataGridView>(form, "Grid");
                if (grid == null)
                {
                    continue;
                }

                IGuangcaiBridge bridge = GetSharedGuangcaiBridge();
                GuangcaiMaterialRuntime runtime = new GuangcaiMaterialRuntime(form, grid, bridge);
                GuangcaiRuntimes[form] = runtime;
                form.FormClosed += delegate
                {
                    runtime.Dispose();
                    GuangcaiRuntimes.Remove(form);
                };
                Log("Guangcai material link installed on FormLFFA.Grid.");
            }
        }

        private static IGuangcaiBridge GetSharedGuangcaiBridge()
        {
            if (sharedGuangcaiBridge == null)
            {
                sharedGuangcaiBridge = new GuangcaiBridgeClient();
            }

            if (!guangcaiExitHooked)
            {
                guangcaiExitHooked = true;
                Application.ApplicationExit += delegate
                {
                    if (sharedGuangcaiBridge != null)
                    {
                        sharedGuangcaiBridge.Dispose();
                        sharedGuangcaiBridge = null;
                    }
                };
            }

            return sharedGuangcaiBridge;
        }

        private sealed class GuangcaiMaterialRuntime : IDisposable
        {
            private const string MaterialNameColumn = "\u6750\u6599\u540d\u79f0";
            private readonly Form form;
            private readonly DataGridView grid;
            private readonly IGuangcaiBridge bridge;
            private string lastMaterialName;
            private int lastRowIndex = -1;
            private bool disposed;

            public GuangcaiMaterialRuntime(Form owner, DataGridView materialGrid, IGuangcaiBridge materialBridge)
            {
                form = owner;
                grid = materialGrid;
                bridge = materialBridge;
                grid.CellClick += GridCellClick;
            }

            private void GridCellClick(object sender, DataGridViewCellEventArgs e)
            {
                try
                {
                    if (disposed || e.RowIndex < 0 || e.ColumnIndex < 0 || e.RowIndex >= grid.Rows.Count)
                    {
                        return;
                    }

                    DataGridViewColumn column = grid.Columns[e.ColumnIndex];
                    if (!IsMaterialNameColumn(column))
                    {
                        return;
                    }

                    string materialName = ReadMaterialName(grid.Rows[e.RowIndex], e.ColumnIndex);
                    if (String.IsNullOrWhiteSpace(materialName))
                    {
                        return;
                    }

                    materialName = materialName.Trim();
                    if (e.RowIndex == lastRowIndex &&
                        String.Equals(materialName, lastMaterialName, StringComparison.Ordinal))
                    {
                        bridge.ActivateWindow(form.Handle);
                        return;
                    }

                    MaterialSelection selection = new MaterialSelection
                    {
                        MaterialName = materialName,
                        RowIndex = e.RowIndex
                    };
                    lastRowIndex = selection.RowIndex;
                    lastMaterialName = selection.MaterialName;
                    if (!bridge.EnsureStarted())
                    {
                        FormPanel.Log("Guangcai bridge is unavailable.");
                        return;
                    }

                    bridge.QueryMaterial(selection.MaterialName, form.Handle);
                    FormPanel.Log("Guangcai query queued. row=" +
                        selection.RowIndex.ToString(CultureInfo.InvariantCulture) +
                        ", nameLength=" + selection.MaterialName.Length.ToString(CultureInfo.InvariantCulture));
                }
                catch (Exception ex)
                {
                    FormPanel.Log("Guangcai material click failed: " + ex);
                }
            }

            private static bool IsMaterialNameColumn(DataGridViewColumn column)
            {
                if (column == null)
                {
                    return false;
                }

                return String.Equals(column.Name, MaterialNameColumn, StringComparison.Ordinal) ||
                    String.Equals(column.HeaderText, MaterialNameColumn, StringComparison.Ordinal) ||
                    String.Equals(column.DataPropertyName, MaterialNameColumn, StringComparison.Ordinal);
            }

            private static string ReadMaterialName(DataGridViewRow row, int columnIndex)
            {
                DataRowView rowView = row.DataBoundItem as DataRowView;
                if (rowView != null && rowView.DataView.Table.Columns.Contains(MaterialNameColumn))
                {
                    object boundValue = rowView[MaterialNameColumn];
                    if (boundValue != null && boundValue != DBNull.Value)
                    {
                        return Convert.ToString(boundValue, CultureInfo.CurrentCulture);
                    }
                }

                object value = row.Cells[columnIndex].Value;
                if (value == null)
                {
                    value = row.Cells[columnIndex].FormattedValue;
                }

                return value == null ? null : Convert.ToString(value, CultureInfo.CurrentCulture);
            }

            public void Dispose()
            {
                if (disposed)
                {
                    return;
                }

                disposed = true;
                grid.CellClick -= GridCellClick;
            }
        }

        private sealed class GuangcaiBridgeClient : IGuangcaiBridge
        {
            private readonly object syncRoot = new object();
            private Process process;
            private StreamWriter input;
            private GuangcaiBridgeState state = GuangcaiBridgeState.NotStarted;
            private bool disposed;

            public GuangcaiBridgeState State
            {
                get
                {
                    lock (syncRoot)
                    {
                        return state;
                    }
                }
            }

            public bool EnsureStarted()
            {
                lock (syncRoot)
                {
                    if (disposed)
                    {
                        return false;
                    }

                    if (process != null && !process.HasExited && input != null)
                    {
                        state = GuangcaiBridgeState.Connected;
                        return true;
                    }

                    DisposeProcess();
                    state = GuangcaiBridgeState.Starting;
                    string baseDir = Path.GetDirectoryName(typeof(FormPanel).Assembly.Location);
                    string bridgePath = Path.Combine(baseDir, "GuangcaiBridge.exe");
                    if (!File.Exists(bridgePath))
                    {
                        state = GuangcaiBridgeState.Unavailable;
                        FormPanel.Log("GuangcaiBridge.exe not found: " + bridgePath);
                        return false;
                    }

                    try
                    {
                        ProcessStartInfo startInfo = new ProcessStartInfo();
                        startInfo.FileName = bridgePath;
                        startInfo.Arguments = "--parent " +
                            Process.GetCurrentProcess().Id.ToString(CultureInfo.InvariantCulture);
                        startInfo.WorkingDirectory = baseDir;
                        startInfo.UseShellExecute = false;
                        startInfo.CreateNoWindow = true;
                        startInfo.RedirectStandardInput = true;
                        process = Process.Start(startInfo);
                        input = process.StandardInput;
                        input.AutoFlush = true;
                        state = GuangcaiBridgeState.Connected;
                        FormPanel.Log("Guangcai bridge process started. pid=" +
                            process.Id.ToString(CultureInfo.InvariantCulture));
                        return true;
                    }
                    catch (Exception ex)
                    {
                        DisposeProcess();
                        state = GuangcaiBridgeState.Unavailable;
                        FormPanel.Log("Start Guangcai bridge failed: " + ex);
                        return false;
                    }
                }
            }

            public void QueryMaterial(string materialName, IntPtr ownerWindow)
            {
                if (String.IsNullOrWhiteSpace(materialName) || !EnsureStarted())
                {
                    return;
                }

                string encoded = Convert.ToBase64String(Encoding.UTF8.GetBytes(materialName.Trim()));
                SendCommand("QUERY\t" + encoded + "\t" +
                    ownerWindow.ToInt64().ToString(CultureInfo.InvariantCulture));
            }

            public bool TryReceiveSelectedPrice(out GuangcaiPriceSelection selection)
            {
                selection = null;
                return false;
            }

            public void ActivateWindow(IntPtr ownerWindow)
            {
                if (!EnsureStarted())
                {
                    return;
                }

                SendCommand("ACTIVATE\t\t" +
                    ownerWindow.ToInt64().ToString(CultureInfo.InvariantCulture));
            }

            private void SendCommand(string command)
            {
                lock (syncRoot)
                {
                    if (disposed || input == null || process == null || process.HasExited)
                    {
                        state = GuangcaiBridgeState.Unavailable;
                        return;
                    }

                    try
                    {
                        input.WriteLine(command);
                    }
                    catch (Exception ex)
                    {
                        state = GuangcaiBridgeState.Unavailable;
                        FormPanel.Log("Send Guangcai bridge command failed: " + ex.Message);
                        DisposeProcess();
                    }
                }
            }

            public void Dispose()
            {
                lock (syncRoot)
                {
                    if (disposed)
                    {
                        return;
                    }

                    disposed = true;
                    DisposeProcess();
                    state = GuangcaiBridgeState.NotStarted;
                }
            }

            private void DisposeProcess()
            {
                if (input != null)
                {
                    try
                    {
                        input.Dispose();
                    }
                    catch
                    {
                    }
                    input = null;
                }

                if (process != null)
                {
                    try
                    {
                        process.Dispose();
                    }
                    catch
                    {
                    }
                    process = null;
                }
            }
        }
    }
}
