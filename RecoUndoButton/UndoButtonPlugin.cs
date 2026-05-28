using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace RecoUndoButton
{
    public sealed class UndoButtonPlugin
    {
        private static readonly Dictionary<Form, UndoRuntime> Runtimes = new Dictionary<Form, UndoRuntime>();
        private static bool idleHooked;

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
                    if (mainForm != null && !Runtimes.ContainsKey(mainForm))
                    {
                        UndoRuntime runtime = new UndoRuntime(mainForm);
                        if (runtime.Install())
                        {
                            Runtimes[mainForm] = runtime;
                            mainForm.FormClosed += delegate { Runtimes.Remove(mainForm); };
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log("Idle install failed: " + ex);
                }
            };
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

        internal static T GetField<T>(object target, string name) where T : class
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
                string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "RecoUndoButton.log");
                File.AppendAllText(path, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " " + message + Environment.NewLine);
            }
            catch
            {
            }
        }

        private sealed class UndoRuntime
        {
            private const int MaxUndoSteps = 20;
            private readonly Form mainForm;
            private readonly List<GridSnapshot> undoStack = new List<GridSnapshot>();
            private readonly Timer settleTimer;
            private readonly List<ToolStripItem> undoItems = new List<ToolStripItem>();
            private DataGridView grid;
            private GridSnapshot lastStableSnapshot;
            private GridSnapshot pendingBeforeSnapshot;
            private bool restoring;

            public UndoRuntime(Form mainForm)
            {
                this.mainForm = mainForm;
                settleTimer = new Timer();
                settleTimer.Interval = 350;
                settleTimer.Tick += delegate { FlushPendingChange(); };
            }

            public bool Install()
            {
                grid = UndoButtonPlugin.GetField<DataGridView>(mainForm, "dataGridViewDE");
                if (grid == null)
                {
                    UndoButtonPlugin.Log("dataGridViewDE not found.");
                    return false;
                }

                lastStableSnapshot = CaptureSnapshot();
                HookGrid();
                bool menuAdded = InstallTopMenuItem();
                if (!menuAdded)
                {
                    menuAdded = InstallContextMenuItem();
                }

                UpdateUndoItems();
                UndoButtonPlugin.Log("Undo runtime installed. topMenu=" + menuAdded.ToString(CultureInfo.InvariantCulture));
                return true;
            }

            private void HookGrid()
            {
                grid.CellValueChanged -= GridChanged;
                grid.CellValueChanged += GridChanged;
                grid.RowsAdded -= GridChanged;
                grid.RowsAdded += GridChanged;
                grid.RowsRemoved -= GridChanged;
                grid.RowsRemoved += GridChanged;
                grid.DataBindingComplete -= GridDataBindingComplete;
                grid.DataBindingComplete += GridDataBindingComplete;
                grid.KeyDown -= GridKeyDown;
                grid.KeyDown += GridKeyDown;
            }

            private void GridChanged(object sender, EventArgs e)
            {
                ScheduleChangeObserved();
            }

            private void GridDataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
            {
                ScheduleChangeObserved();
            }

            private void GridKeyDown(object sender, KeyEventArgs e)
            {
                if (e.Control && (e.KeyCode == Keys.V || e.KeyCode == Keys.X || e.KeyCode == Keys.Delete))
                {
                    ScheduleChangeObserved();
                }
            }

            private bool InstallTopMenuItem()
            {
                MenuStrip menuStrip = FindMenuStrip(mainForm);
                if (menuStrip != null)
                {
                    ToolStripMenuItem item = CreateUndoMenuItem();
                    ToolStripMenuItem editMenu = FindMenuItem(menuStrip.Items, "编辑");
                    if (editMenu != null)
                    {
                        if (FindMenuItem(editMenu.DropDownItems, "撤回上一步") == null)
                        {
                            editMenu.DropDownItems.Insert(Math.Min(1, editMenu.DropDownItems.Count), item);
                            undoItems.Add(item);
                        }
                    }
                    else if (FindMenuItem(menuStrip.Items, "撤回上一步") == null)
                    {
                        menuStrip.Items.Add(item);
                        undoItems.Add(item);
                    }

                    return undoItems.Count > 0;
                }

                ToolStrip toolStrip = FindToolStrip(mainForm);
                if (toolStrip != null)
                {
                    if (FindToolStripItem(toolStrip.Items, "撤回上一步") == null)
                    {
                        ToolStripButton button = new ToolStripButton("撤回上一步");
                        button.DisplayStyle = ToolStripItemDisplayStyle.Text;
                        button.Click += delegate { UndoLastStep(); };
                        toolStrip.Items.Add(button);
                        undoItems.Add(button);
                    }

                    return undoItems.Count > 0;
                }

                UndoButtonPlugin.Log("Top menu/toolbar not found, will try context menu fallback.");
                return false;
            }

            private bool InstallContextMenuItem()
            {
                ContextMenuStrip menu = UndoButtonPlugin.GetField<ContextMenuStrip>(mainForm, "contextMenuStripDE");
                if (menu == null && grid != null)
                {
                    menu = grid.ContextMenuStrip;
                }

                if (menu == null)
                {
                    UndoButtonPlugin.Log("contextMenuStripDE not found.");
                    return false;
                }

                if (FindMenuItem(menu.Items, "撤回上一步") != null)
                {
                    return true;
                }

                ToolStripMenuItem item = CreateUndoMenuItem();
                menu.Items.Insert(Math.Min(1, menu.Items.Count), item);
                undoItems.Add(item);
                return true;
            }

            private ToolStripMenuItem CreateUndoMenuItem()
            {
                ToolStripMenuItem item = new ToolStripMenuItem("撤回上一步");
                item.Click += delegate { UndoLastStep(); };
                return item;
            }

            private void ScheduleChangeObserved()
            {
                if (restoring || grid == null || grid.IsDisposed)
                {
                    return;
                }

                if (pendingBeforeSnapshot == null)
                {
                    pendingBeforeSnapshot = lastStableSnapshot ?? CaptureSnapshot();
                }

                settleTimer.Stop();
                settleTimer.Start();
            }

            private void FlushPendingChange()
            {
                settleTimer.Stop();
                if (restoring || pendingBeforeSnapshot == null)
                {
                    return;
                }

                GridSnapshot current = CaptureSnapshot();
                if (!GridSnapshot.AreEquivalent(pendingBeforeSnapshot, current))
                {
                    PushUndo(pendingBeforeSnapshot);
                    lastStableSnapshot = current;
                }

                pendingBeforeSnapshot = null;
                UpdateUndoItems();
            }

            private void PushUndo(GridSnapshot snapshot)
            {
                if (snapshot == null)
                {
                    return;
                }

                if (undoStack.Count > 0 && GridSnapshot.AreEquivalent(undoStack[undoStack.Count - 1], snapshot))
                {
                    return;
                }

                undoStack.Add(snapshot);
                while (undoStack.Count > MaxUndoSteps)
                {
                    undoStack.RemoveAt(0);
                }
            }

            private void UndoLastStep()
            {
                try
                {
                    EndGridEdit();
                    FlushPendingChange();
                    if (undoStack.Count == 0)
                    {
                        MessageBox.Show(mainForm, "没有可撤回的操作。", "撤回上一步", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdateUndoItems();
                        return;
                    }

                    GridSnapshot target = undoStack[undoStack.Count - 1];
                    undoStack.RemoveAt(undoStack.Count - 1);
                    RestoreSnapshot(target);
                    lastStableSnapshot = CaptureSnapshot();
                    pendingBeforeSnapshot = null;
                    UpdateUndoItems();
                }
                catch (Exception ex)
                {
                    UndoButtonPlugin.Log("Undo failed: " + ex);
                    MessageBox.Show(mainForm, "撤回失败：" + ex.Message, "撤回上一步", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            private void EndGridEdit()
            {
                try
                {
                    grid.EndEdit();
                    BindingSource bindingSource = grid.DataSource as BindingSource;
                    if (bindingSource != null)
                    {
                        bindingSource.EndEdit();
                    }
                }
                catch
                {
                }
            }

            private GridSnapshot CaptureSnapshot()
            {
                GridSnapshot snapshot = new GridSnapshot();
                snapshot.CurrentRowIndex = grid.CurrentCell == null ? -1 : grid.CurrentCell.RowIndex;
                snapshot.CurrentColumnIndex = grid.CurrentCell == null ? -1 : grid.CurrentCell.ColumnIndex;

                DataView view = ResolveDataView(grid.DataSource);
                DataTable table = view == null ? ResolveDataTable(grid.DataSource) : view.Table;
                if (table != null)
                {
                    snapshot.Mode = "table";
                    foreach (DataColumn column in table.Columns)
                    {
                        snapshot.Columns.Add(column.ColumnName);
                    }

                    if (view != null)
                    {
                        foreach (DataRowView rowView in view)
                        {
                            snapshot.Rows.Add(rowView.Row.ItemArray.ToArray());
                        }
                    }
                    else
                    {
                        foreach (DataRow row in table.Rows)
                        {
                            snapshot.Rows.Add(row.ItemArray.ToArray());
                        }
                    }

                    snapshot.Signature = snapshot.BuildSignature();
                    return snapshot;
                }

                snapshot.Mode = "grid";
                foreach (DataGridViewColumn column in grid.Columns)
                {
                    snapshot.Columns.Add(column.Name);
                }

                foreach (DataGridViewRow row in grid.Rows)
                {
                    if (row.IsNewRow)
                    {
                        continue;
                    }

                    object[] values = new object[grid.Columns.Count];
                    for (int i = 0; i < grid.Columns.Count; i++)
                    {
                        values[i] = row.Cells[i].Value;
                    }

                    snapshot.Rows.Add(values);
                }

                snapshot.Signature = snapshot.BuildSignature();
                return snapshot;
            }

            private void RestoreSnapshot(GridSnapshot snapshot)
            {
                restoring = true;
                try
                {
                    DataView view = ResolveDataView(grid.DataSource);
                    DataTable table = view == null ? ResolveDataTable(grid.DataSource) : view.Table;
                    if (snapshot.Mode == "table" && table != null)
                    {
                        RestoreTableSnapshot(view, table, snapshot);
                    }
                    else
                    {
                        RestoreGridSnapshot(snapshot);
                    }

                    RestoreCurrentCell(snapshot);
                }
                finally
                {
                    restoring = false;
                }
            }

            private void RestoreTableSnapshot(DataView view, DataTable table, GridSnapshot snapshot)
            {
                if (table.Columns.Count != snapshot.Columns.Count)
                {
                    throw new InvalidOperationException("当前表格列结构已经变化，不能安全撤回。");
                }

                if (view != null && view.Count == snapshot.Rows.Count)
                {
                    for (int rowIndex = 0; rowIndex < snapshot.Rows.Count; rowIndex++)
                    {
                        ApplyValuesToDataRow(view[rowIndex].Row, snapshot.Rows[rowIndex]);
                    }
                }
                else if (table.Rows.Count == snapshot.Rows.Count)
                {
                    for (int rowIndex = 0; rowIndex < snapshot.Rows.Count; rowIndex++)
                    {
                        ApplyValuesToDataRow(table.Rows[rowIndex], snapshot.Rows[rowIndex]);
                    }
                }
                else
                {
                    table.Rows.Clear();
                    foreach (object[] values in snapshot.Rows)
                    {
                        DataRow row = table.NewRow();
                        ApplyValuesToDataRow(row, values);
                        table.Rows.Add(row);
                    }
                }

                CurrencyManager manager = mainForm.BindingContext[grid.DataSource] as CurrencyManager;
                if (manager != null)
                {
                    manager.Refresh();
                }

                grid.Refresh();
            }

            private void ApplyValuesToDataRow(DataRow row, object[] values)
            {
                int count = Math.Min(row.Table.Columns.Count, values.Length);
                for (int i = 0; i < count; i++)
                {
                    object value = values[i];
                    row[i] = value == null ? DBNull.Value : value;
                }
            }

            private void RestoreGridSnapshot(GridSnapshot snapshot)
            {
                if (grid.DataSource != null)
                {
                    throw new InvalidOperationException("当前表格绑定方式暂不支持直接恢复。");
                }

                while (grid.Rows.Count > snapshot.Rows.Count)
                {
                    grid.Rows.RemoveAt(grid.Rows.Count - 1);
                }

                while (grid.Rows.Count < snapshot.Rows.Count)
                {
                    grid.Rows.Add();
                }

                for (int rowIndex = 0; rowIndex < snapshot.Rows.Count; rowIndex++)
                {
                    object[] values = snapshot.Rows[rowIndex];
                    for (int columnIndex = 0; columnIndex < values.Length && columnIndex < grid.Columns.Count; columnIndex++)
                    {
                        grid.Rows[rowIndex].Cells[columnIndex].Value = values[columnIndex];
                    }
                }
            }

            private void RestoreCurrentCell(GridSnapshot snapshot)
            {
                if (snapshot.CurrentRowIndex < 0 || snapshot.CurrentColumnIndex < 0)
                {
                    return;
                }

                if (snapshot.CurrentRowIndex >= grid.Rows.Count || snapshot.CurrentColumnIndex >= grid.Columns.Count)
                {
                    return;
                }

                try
                {
                    grid.CurrentCell = grid.Rows[snapshot.CurrentRowIndex].Cells[snapshot.CurrentColumnIndex];
                }
                catch
                {
                }
            }

            private void UpdateUndoItems()
            {
                bool enabled = undoStack.Count > 0;
                foreach (ToolStripItem item in undoItems.ToArray())
                {
                    if (item == null || item.IsDisposed)
                    {
                        undoItems.Remove(item);
                        continue;
                    }

                    item.Enabled = enabled;
                }
            }

            private static MenuStrip FindMenuStrip(Control root)
            {
                foreach (FieldInfo field in root.GetType().GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
                {
                    MenuStrip menu = field.GetValue(root) as MenuStrip;
                    if (menu != null)
                    {
                        return menu;
                    }
                }

                return FindControl<MenuStrip>(root);
            }

            private static ToolStrip FindToolStrip(Control root)
            {
                foreach (FieldInfo field in root.GetType().GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
                {
                    ToolStrip toolStrip = field.GetValue(root) as ToolStrip;
                    if (toolStrip != null && !(toolStrip is MenuStrip) && !(toolStrip is StatusStrip))
                    {
                        return toolStrip;
                    }
                }

                foreach (ToolStrip toolStrip in FindControls<ToolStrip>(root))
                {
                    if (!(toolStrip is MenuStrip) && !(toolStrip is StatusStrip))
                    {
                        return toolStrip;
                    }
                }

                return null;
            }

            private static T FindControl<T>(Control root) where T : Control
            {
                foreach (T control in FindControls<T>(root))
                {
                    return control;
                }

                return null;
            }

            private static IEnumerable<T> FindControls<T>(Control root) where T : Control
            {
                if (root == null)
                {
                    yield break;
                }

                T matched = root as T;
                if (matched != null)
                {
                    yield return matched;
                }

                foreach (Control child in root.Controls)
                {
                    foreach (T nested in FindControls<T>(child))
                    {
                        yield return nested;
                    }
                }
            }

            private static ToolStripMenuItem FindMenuItem(ToolStripItemCollection items, string text)
            {
                foreach (ToolStripItem item in items)
                {
                    ToolStripMenuItem menuItem = item as ToolStripMenuItem;
                    if (menuItem != null && NormalizeMenuText(menuItem.Text) == text)
                    {
                        return menuItem;
                    }
                }

                return null;
            }

            private static ToolStripItem FindToolStripItem(ToolStripItemCollection items, string text)
            {
                foreach (ToolStripItem item in items)
                {
                    if (NormalizeMenuText(item.Text) == text)
                    {
                        return item;
                    }
                }

                return null;
            }

            private static string NormalizeMenuText(string text)
            {
                return String.IsNullOrEmpty(text) ? "" : text.Replace("&", "").Trim();
            }

            private static DataView ResolveDataView(object source)
            {
                BindingSource bindingSource = source as BindingSource;
                if (bindingSource != null)
                {
                    DataView bindingView = bindingSource.List as DataView;
                    if (bindingView != null)
                    {
                        return bindingView;
                    }

                    return ResolveDataView(bindingSource.DataSource);
                }

                return source as DataView;
            }

            private static DataTable ResolveDataTable(object source)
            {
                BindingSource bindingSource = source as BindingSource;
                if (bindingSource != null)
                {
                    DataTable bindingTable = bindingSource.List as DataTable;
                    if (bindingTable != null)
                    {
                        return bindingTable;
                    }

                    return ResolveDataTable(bindingSource.DataSource);
                }

                DataView view = source as DataView;
                if (view != null)
                {
                    return view.Table;
                }

                return source as DataTable;
            }
        }

        private sealed class GridSnapshot
        {
            public string Mode;
            public readonly List<string> Columns = new List<string>();
            public readonly List<object[]> Rows = new List<object[]>();
            public int CurrentRowIndex;
            public int CurrentColumnIndex;
            public string Signature;

            public static bool AreEquivalent(GridSnapshot left, GridSnapshot right)
            {
                if (left == null || right == null)
                {
                    return false;
                }

                return String.Equals(left.Signature, right.Signature, StringComparison.Ordinal);
            }

            public string BuildSignature()
            {
                StringBuilder builder = new StringBuilder();
                builder.Append(Mode).Append("|").Append(Columns.Count.ToString(CultureInfo.InvariantCulture)).Append("|").Append(Rows.Count.ToString(CultureInfo.InvariantCulture)).AppendLine();
                foreach (string column in Columns)
                {
                    builder.Append(column).Append('\t');
                }

                builder.AppendLine();
                foreach (object[] row in Rows)
                {
                    foreach (object value in row)
                    {
                        builder.Append(ValueToSignatureText(value)).Append('\t');
                    }

                    builder.AppendLine();
                }

                return builder.ToString();
            }

            private static string ValueToSignatureText(object value)
            {
                if (value == null || value == DBNull.Value)
                {
                    return "";
                }

                IFormattable formattable = value as IFormattable;
                if (formattable != null)
                {
                    return formattable.ToString(null, CultureInfo.InvariantCulture);
                }

                return Convert.ToString(value, CultureInfo.InvariantCulture);
            }
        }
    }
}
