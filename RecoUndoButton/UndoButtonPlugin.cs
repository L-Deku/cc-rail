using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
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
            private readonly List<GridSnapshot> redoStack = new List<GridSnapshot>();
            private readonly Timer settleTimer;
            private readonly List<ToolStripItem> undoItems = new List<ToolStripItem>();
            private readonly List<ToolStripItem> redoItems = new List<ToolStripItem>();
            private DataGridView grid;
            private GridSnapshot lastStableSnapshot;
            private GridSnapshot pendingBeforeSnapshot;
            private Image undoIcon;
            private Image redoIcon;
            private ToolStrip floatingToolbar;
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
                bool buttonsAdded = InstallToolbarItems();
                if (!buttonsAdded)
                {
                    buttonsAdded = InstallContextMenuItem();
                }

                UpdateUndoItems();
                UndoButtonPlugin.Log("Undo runtime installed. buttons=" + buttonsAdded.ToString(CultureInfo.InvariantCulture));
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

            private bool InstallToolbarItems()
            {
                RemoveExistingToolbarButtons(mainForm);
                ToolStrip floating = CreateFloatingChapterToolbar();
                if (floating != null)
                {
                    floatingToolbar = floating;
                    return undoItems.Count > 0 && redoItems.Count > 0;
                }

                ToolStrip toolStrip = FindChapterToolStrip(mainForm);
                if (toolStrip != null)
                {
                    if (FindNamedItem(toolStrip.Items, "RecoUndoButton_Undo") == null)
                    {
                        ToolStripButton undo = CreateUndoToolButton();
                        ToolStripButton redo = CreateRedoToolButton();
                        int insertIndex = FindToolbarInsertIndex(toolStrip);
                        toolStrip.Items.Insert(Math.Min(insertIndex, toolStrip.Items.Count), redo);
                        toolStrip.Items.Insert(Math.Min(insertIndex, toolStrip.Items.Count), undo);
                        undoItems.Add(undo);
                        redoItems.Add(redo);
                    }

                    return undoItems.Count > 0 && redoItems.Count > 0;
                }

                MenuStrip menuStrip = FindMenuStrip(mainForm);
                if (menuStrip != null)
                {
                    RemoveOldUndoTextItem(menuStrip.Items);
                    if (FindNamedItem(menuStrip.Items, "RecoUndoButton_Undo") == null)
                    {
                        ToolStripButton undo = CreateUndoToolButton();
                        ToolStripButton redo = CreateRedoToolButton();
                        int index = FindEditMenuIndex(menuStrip.Items);
                        if (index < 0)
                        {
                            index = menuStrip.Items.Count;
                        }
                        else
                        {
                            index++;
                        }

                        menuStrip.Items.Insert(Math.Min(index, menuStrip.Items.Count), redo);
                        menuStrip.Items.Insert(Math.Min(index, menuStrip.Items.Count), undo);
                        undoItems.Add(undo);
                        redoItems.Add(redo);
                    }

                    return undoItems.Count > 0 && redoItems.Count > 0;
                }

                UndoButtonPlugin.Log("Chapter toolbar/menu not found, will try context menu fallback.");
                return false;
            }

            private ToolStrip CreateFloatingChapterToolbar()
            {
                TreeView tree = UndoButtonPlugin.GetField<TreeView>(mainForm, "Tv_tree");
                if (tree == null || tree.IsDisposed)
                {
                    UndoButtonPlugin.Log("Tv_tree not found, cannot create floating toolbar.");
                    return null;
                }

                ToolStrip toolbar = new ToolStrip();
                toolbar.Name = "RecoUndoButton_FloatingToolbar";
                toolbar.AutoSize = false;
                toolbar.CanOverflow = false;
                toolbar.GripStyle = ToolStripGripStyle.Hidden;
                toolbar.RenderMode = ToolStripRenderMode.System;
                toolbar.Dock = DockStyle.None;
                toolbar.Padding = new Padding(0);
                toolbar.Margin = new Padding(0);
                toolbar.BackColor = SystemColors.Control;
                toolbar.ImageScalingSize = new Size(28, 28);
                toolbar.Size = new Size(64, 30);

                ToolStripButton undo = CreateUndoToolButton();
                ToolStripButton redo = CreateRedoToolButton();
                toolbar.Items.Add(undo);
                toolbar.Items.Add(redo);
                undoItems.Add(undo);
                redoItems.Add(redo);

                mainForm.Controls.Add(toolbar);
                PositionFloatingToolbar(tree, toolbar);
                toolbar.BringToFront();
                mainForm.Resize += delegate { PositionFloatingToolbar(tree, toolbar); };
                mainForm.Layout += delegate { PositionFloatingToolbar(tree, toolbar); };
                tree.ParentChanged += delegate { PositionFloatingToolbar(tree, toolbar); };
                tree.VisibleChanged += delegate { PositionFloatingToolbar(tree, toolbar); };
                try
                {
                    mainForm.BeginInvoke((MethodInvoker)delegate { PositionFloatingToolbar(tree, toolbar); });
                }
                catch
                {
                }

                UndoButtonPlugin.Log("Floating chapter toolbar created.");
                return toolbar;
            }

            private void PositionFloatingToolbar(TreeView tree, ToolStrip toolbar)
            {
                if (tree == null || toolbar == null || tree.IsDisposed || toolbar.IsDisposed)
                {
                    return;
                }

                try
                {
                    Rectangle treeBounds = ControlBoundsOnForm(mainForm, tree);
                    Rectangle anchorBounds = FindRightIconAnchor(mainForm, tree, toolbar);
                    int x;
                    int y;
                    if (!anchorBounds.IsEmpty)
                    {
                        x = anchorBounds.Left - toolbar.Width - 8;
                        y = anchorBounds.Top + Math.Max(0, (anchorBounds.Height - toolbar.Height) / 2);
                    }
                    else
                    {
                        x = treeBounds.Left + 165;
                        y = treeBounds.Top - 34;
                    }

                    int minimumX = treeBounds.Left + 130;
                    if (x < minimumX)
                    {
                        x = minimumX;
                    }
                    if (x < 0)
                    {
                        x = 0;
                    }
                    if (y < 0)
                    {
                        y = 0;
                    }

                    Point targetLocation = new Point(x, y);
                    if (toolbar.Location != targetLocation)
                    {
                        toolbar.Location = targetLocation;
                    }
                    toolbar.Visible = tree.Visible;
                    toolbar.BringToFront();
                    UndoButtonPlugin.Log("Floating toolbar positioned: x=" + x.ToString(CultureInfo.InvariantCulture) + " y=" + y.ToString(CultureInfo.InvariantCulture) + " anchor=" + RectangleToLog(anchorBounds));
                }
                catch (Exception ex)
                {
                    UndoButtonPlugin.Log("Position floating toolbar failed: " + ex.Message);
                }
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

                RemoveOldUndoTextItem(menu.Items);
                if (FindNamedItem(menu.Items, "RecoUndoButton_Undo") != null)
                {
                    return true;
                }

                ToolStripMenuItem undo = CreateUndoMenuButton();
                ToolStripMenuItem redo = CreateRedoMenuButton();
                int insertIndex = Math.Min(1, menu.Items.Count);
                menu.Items.Insert(insertIndex, redo);
                menu.Items.Insert(insertIndex, undo);
                undoItems.Add(undo);
                redoItems.Add(redo);
                return true;
            }

            private ToolStripMenuItem CreateUndoMenuButton()
            {
                ToolStripMenuItem item = new ToolStripMenuItem("");
                item.Name = "RecoUndoButton_Undo";
                item.ToolTipText = "撤回上一步";
                ConfigureIconItem(item);
                item.Image = GetArrowIcon(true);
                item.Click += delegate { UndoLastStep(); };
                return item;
            }

            private ToolStripMenuItem CreateRedoMenuButton()
            {
                ToolStripMenuItem item = new ToolStripMenuItem("");
                item.Name = "RecoUndoButton_Redo";
                item.ToolTipText = "恢复下一步";
                ConfigureIconItem(item);
                item.Image = GetArrowIcon(false);
                item.Click += delegate { RedoLastStep(); };
                return item;
            }

            private ToolStripButton CreateUndoToolButton()
            {
                ToolStripButton button = new ToolStripButton();
                button.Name = "RecoUndoButton_Undo";
                button.ToolTipText = "撤回上一步";
                ConfigureIconItem(button);
                button.Image = GetArrowIcon(true);
                button.Click += delegate { UndoLastStep(); };
                return button;
            }

            private ToolStripButton CreateRedoToolButton()
            {
                ToolStripButton button = new ToolStripButton();
                button.Name = "RecoUndoButton_Redo";
                button.ToolTipText = "恢复下一步";
                ConfigureIconItem(button);
                button.Image = GetArrowIcon(false);
                button.Click += delegate { RedoLastStep(); };
                return button;
            }

            private static void ConfigureIconItem(ToolStripItem item)
            {
                item.AutoSize = false;
                item.DisplayStyle = ToolStripItemDisplayStyle.Image;
                item.ImageScaling = ToolStripItemImageScaling.SizeToFit;
                item.Margin = new Padding(1, 0, 1, 0);
                item.Padding = new Padding(0);
                item.Size = new Size(30, 30);
            }

            private Image GetArrowIcon(bool undo)
            {
                if (undo)
                {
                    return undoIcon ?? (undoIcon = LoadArrowIcon("undo.png") ?? CreateArrowIcon(true));
                }

                return redoIcon ?? (redoIcon = LoadArrowIcon("redo.png") ?? CreateArrowIcon(false));
            }

            private static Image LoadArrowIcon(string fileName)
            {
                try
                {
                    string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "RecoUndoButtonIcons", fileName);
                    if (!File.Exists(path))
                    {
                        return null;
                    }

                    using (Image image = Image.FromFile(path))
                    {
                        return new Bitmap(image);
                    }
                }
                catch (Exception ex)
                {
                    UndoButtonPlugin.Log("Load arrow icon failed: " + fileName + " " + ex.Message);
                    return null;
                }
            }

            private static Bitmap CreateArrowIcon(bool undo)
            {
                Bitmap bitmap = new Bitmap(18, 18);
                using (Graphics graphics = Graphics.FromImage(bitmap))
                {
                    graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
                    graphics.Clear(Color.Transparent);
                    using (Pen pen = new Pen(Color.FromArgb(22, 25, 29), 1.9f))
                    using (GraphicsPath path = new GraphicsPath())
                    {
                        pen.StartCap = LineCap.Round;
                        pen.EndCap = LineCap.Round;
                        pen.LineJoin = LineJoin.Round;
                        if (undo)
                        {
                            graphics.DrawLine(pen, 5.9f, 5.4f, 2.9f, 5.4f);
                            graphics.DrawLine(pen, 2.9f, 5.4f, 2.9f, 8.4f);
                            path.StartFigure();
                            path.AddBezier(3.2f, 5.8f, 6.3f, 3.0f, 13.8f, 4.8f, 13.8f, 10.0f);
                            path.AddBezier(13.8f, 13.7f, 10.6f, 15.0f, 7.4f, 13.2f, 6.3f, 11.6f);
                            graphics.DrawPath(pen, path);
                        }
                        else
                        {
                            graphics.DrawLine(pen, 12.1f, 5.4f, 15.1f, 5.4f);
                            graphics.DrawLine(pen, 15.1f, 5.4f, 15.1f, 8.4f);
                            path.StartFigure();
                            path.AddBezier(14.8f, 5.8f, 11.7f, 3.0f, 4.2f, 4.8f, 4.2f, 10.0f);
                            path.AddBezier(4.2f, 13.7f, 7.4f, 15.0f, 10.6f, 13.2f, 11.7f, 11.6f);
                            graphics.DrawPath(pen, path);
                        }
                    }
                }

                return bitmap;
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
                    redoStack.Clear();
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

            private void PushRedo(GridSnapshot snapshot)
            {
                if (snapshot == null)
                {
                    return;
                }

                if (redoStack.Count > 0 && GridSnapshot.AreEquivalent(redoStack[redoStack.Count - 1], snapshot))
                {
                    return;
                }

                redoStack.Add(snapshot);
                while (redoStack.Count > MaxUndoSteps)
                {
                    redoStack.RemoveAt(0);
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

                    GridSnapshot current = CaptureSnapshot();
                    GridSnapshot target = undoStack[undoStack.Count - 1];
                    undoStack.RemoveAt(undoStack.Count - 1);
                    PushRedo(current);
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

            private void RedoLastStep()
            {
                try
                {
                    EndGridEdit();
                    FlushPendingChange();
                    if (redoStack.Count == 0)
                    {
                        MessageBox.Show(mainForm, "没有可恢复的操作。", "恢复下一步", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdateUndoItems();
                        return;
                    }

                    GridSnapshot current = CaptureSnapshot();
                    GridSnapshot target = redoStack[redoStack.Count - 1];
                    redoStack.RemoveAt(redoStack.Count - 1);
                    PushUndo(current);
                    RestoreSnapshot(target);
                    lastStableSnapshot = CaptureSnapshot();
                    pendingBeforeSnapshot = null;
                    UpdateUndoItems();
                }
                catch (Exception ex)
                {
                    UndoButtonPlugin.Log("Redo failed: " + ex);
                    MessageBox.Show(mainForm, "恢复失败：" + ex.Message, "恢复下一步", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                bool undoEnabled = undoStack.Count > 0;
                foreach (ToolStripItem item in undoItems.ToArray())
                {
                    if (item == null || item.IsDisposed)
                    {
                        undoItems.Remove(item);
                        continue;
                    }

                    item.Enabled = true;
                    item.ToolTipText = undoEnabled ? "撤回上一步" : "没有可撤回的操作";
                }

                bool redoEnabled = redoStack.Count > 0;
                foreach (ToolStripItem item in redoItems.ToArray())
                {
                    if (item == null || item.IsDisposed)
                    {
                        redoItems.Remove(item);
                        continue;
                    }

                    item.Enabled = true;
                    item.ToolTipText = redoEnabled ? "恢复下一步" : "没有可恢复的操作";
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

            private static void RemoveExistingToolbarButtons(Control root)
            {
                foreach (ToolStrip toolStrip in FindControls<ToolStrip>(root))
                {
                    if (toolStrip == null || toolStrip.IsDisposed)
                    {
                        continue;
                    }

                    RemoveNamedToolbarButton(toolStrip.Items, "RecoUndoButton_Undo");
                    RemoveNamedToolbarButton(toolStrip.Items, "RecoUndoButton_Redo");
                    RemoveOldUndoTextItem(toolStrip.Items);
                }
            }

            private static void RemoveNamedToolbarButton(ToolStripItemCollection items, string name)
            {
                ToolStripItem item = FindNamedItem(items, name);
                if (item == null)
                {
                    return;
                }

                items.Remove(item);
                item.Dispose();
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

            private static ToolStrip FindChapterToolStrip(Form mainForm)
            {
                TreeView tree = UndoButtonPlugin.GetField<TreeView>(mainForm, "Tv_tree");
                ToolStrip best = null;
                int bestScore = Int32.MinValue;
                foreach (ToolStrip toolStrip in FindControls<ToolStrip>(mainForm))
                {
                    if (toolStrip == null || toolStrip.IsDisposed || toolStrip is MenuStrip || toolStrip is StatusStrip)
                    {
                        continue;
                    }

                    int score = ScoreToolStrip(toolStrip, tree, mainForm);
                    if (score > bestScore)
                    {
                        bestScore = score;
                        best = toolStrip;
                    }
                }

                if (best != null)
                {
                    UndoButtonPlugin.Log("Chapter toolbar selected: name=" + best.Name + " score=" + bestScore.ToString(CultureInfo.InvariantCulture) + " items=" + best.Items.Count.ToString(CultureInfo.InvariantCulture));
                }

                return best;
            }

            private static Rectangle FindRightIconAnchor(Form mainForm, TreeView tree, ToolStrip floatingToolbar)
            {
                if (mainForm == null || tree == null)
                {
                    return Rectangle.Empty;
                }

                Rectangle treeBounds = ControlBoundsOnForm(mainForm, tree);
                int desiredTop = treeBounds.Top - 34;
                Rectangle best = Rectangle.Empty;
                int bestScore = Int32.MinValue;

                foreach (ToolStrip toolStrip in FindControls<ToolStrip>(mainForm))
                {
                    if (toolStrip == null || toolStrip.IsDisposed || toolStrip == floatingToolbar || toolStrip is MenuStrip || toolStrip is StatusStrip)
                    {
                        continue;
                    }

                    Rectangle stripBounds = ControlBoundsOnForm(mainForm, toolStrip);
                    for (int i = 0; i < toolStrip.Items.Count; i++)
                    {
                        ToolStripItem item = toolStrip.Items[i];
                        if (item == null || !item.Visible || item.Image == null)
                        {
                            continue;
                        }

                        Rectangle itemBounds = new Rectangle(stripBounds.Left + item.Bounds.Left, stripBounds.Top + item.Bounds.Top, item.Bounds.Width, item.Bounds.Height);
                        int score = 0;
                        score += Math.Max(0, 180 - Math.Abs(itemBounds.Top - desiredTop) * 8);
                        if (itemBounds.Left >= treeBounds.Left + 220)
                        {
                            score += 120;
                        }
                        else if (itemBounds.Left >= treeBounds.Left + 170)
                        {
                            score += 70;
                        }
                        else if (itemBounds.Left < treeBounds.Left + 90)
                        {
                            score -= 150;
                        }

                        if (itemBounds.Left < treeBounds.Right + 40)
                        {
                            score += 30;
                        }
                        if (itemBounds.Top < treeBounds.Top - 65 || itemBounds.Top > treeBounds.Top + 5)
                        {
                            score -= 120;
                        }

                        if (score > bestScore || (score == bestScore && !best.IsEmpty && itemBounds.Left < best.Left))
                        {
                            bestScore = score;
                            best = itemBounds;
                        }
                    }
                }

                if (!best.IsEmpty)
                {
                    UndoButtonPlugin.Log("Right icon anchor selected: score=" + bestScore.ToString(CultureInfo.InvariantCulture) + " bounds=" + RectangleToLog(best));
                }

                return best;
            }

            private static int FindToolbarInsertIndex(ToolStrip toolStrip)
            {
                if (toolStrip == null)
                {
                    return 0;
                }

                for (int i = 0; i < toolStrip.Items.Count; i++)
                {
                    ToolStripItem item = toolStrip.Items[i];
                    if (item == null || !item.Visible)
                    {
                        continue;
                    }

                    string text = (item.Text ?? "").Trim();
                    if (text.IndexOf("章节", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        int indexAfterChapterLabel = Math.Min(i + 1, toolStrip.Items.Count);
                        UndoButtonPlugin.Log("Toolbar insert after chapter label: index=" + indexAfterChapterLabel.ToString(CultureInfo.InvariantCulture) + " text=" + text);
                        return indexAfterChapterLabel;
                    }
                }

                for (int i = 0; i < toolStrip.Items.Count; i++)
                {
                    ToolStripItem item = toolStrip.Items[i];
                    if (item != null && item.Visible && item.Image != null)
                    {
                        UndoButtonPlugin.Log("Toolbar insert before first image item: index=" + i.ToString(CultureInfo.InvariantCulture) + " name=" + item.Name);
                        return i;
                    }
                }

                for (int i = 0; i < toolStrip.Items.Count; i++)
                {
                    if (toolStrip.Items[i].Visible)
                    {
                        return i;
                    }
                }

                return 0;
            }

            private static string RectangleToLog(Rectangle rectangle)
            {
                if (rectangle.IsEmpty)
                {
                    return "empty";
                }

                return rectangle.Left.ToString(CultureInfo.InvariantCulture) + "," + rectangle.Top.ToString(CultureInfo.InvariantCulture) + "," + rectangle.Width.ToString(CultureInfo.InvariantCulture) + "," + rectangle.Height.ToString(CultureInfo.InvariantCulture);
            }

            private static int ScoreToolStrip(ToolStrip toolStrip, TreeView tree, Form mainForm)
            {
                int score = 0;
                if (toolStrip.Items.Count > 0)
                {
                    score += Math.Min(20, toolStrip.Items.Count * 2);
                }
                foreach (ToolStripItem item in toolStrip.Items)
                {
                    if (item is ToolStripButton)
                    {
                        score += 4;
                    }
                    if (item.Image != null)
                    {
                        score += 4;
                    }
                }

                if (tree != null)
                {
                    Rectangle stripBounds = ControlBoundsOnForm(mainForm, toolStrip);
                    Rectangle treeBounds = ControlBoundsOnForm(mainForm, tree);

                    int desiredTop = treeBounds.Top - 34;
                    int verticalDistance = Math.Abs(stripBounds.Top - desiredTop);
                    score += Math.Max(0, 140 - verticalDistance * 5);

                    if (stripBounds.Top < treeBounds.Top - 65)
                    {
                        score -= 120;
                    }
                    if (stripBounds.Top > treeBounds.Top + 5)
                    {
                        score -= 80;
                    }

                    bool horizontallyNearRightIconRow = stripBounds.Right > treeBounds.Left + 220 && stripBounds.Left < treeBounds.Right + 80;
                    if (horizontallyNearRightIconRow)
                    {
                        score += 55;
                    }
                    if (stripBounds.Left >= treeBounds.Left + 120 && stripBounds.Left <= treeBounds.Right)
                    {
                        score += 45;
                    }
                    if (stripBounds.Width >= 120)
                    {
                        score += 10;
                    }
                }

                string name = (toolStrip.Name ?? "").ToLowerInvariant();
                if (name.Contains("tree") || name.Contains("chapter") || name.Contains("tool"))
                {
                    score += 12;
                }

                return score;
            }

            private static Rectangle ControlBoundsOnForm(Form mainForm, Control control)
            {
                if (mainForm == null || control == null)
                {
                    return Rectangle.Empty;
                }

                try
                {
                    Rectangle screenBounds = control.RectangleToScreen(control.ClientRectangle);
                    Point location = mainForm.PointToClient(screenBounds.Location);
                    return new Rectangle(location, screenBounds.Size);
                }
                catch
                {
                    return control.Bounds;
                }
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

            private static ToolStripItem FindNamedItem(ToolStripItemCollection items, string name)
            {
                foreach (ToolStripItem item in items)
                {
                    if (String.Equals(item.Name, name, StringComparison.Ordinal))
                    {
                        return item;
                    }
                }

                return null;
            }

            private static int FindEditMenuIndex(ToolStripItemCollection items)
            {
                for (int i = 0; i < items.Count; i++)
                {
                    if (NormalizeMenuText(items[i].Text) == "编辑")
                    {
                        return i;
                    }
                }

                return -1;
            }

            private static void RemoveOldUndoTextItem(ToolStripItemCollection items)
            {
                ToolStripItem oldItem = FindToolStripItem(items, "撤回上一步");
                if (oldItem != null && String.IsNullOrEmpty(oldItem.Name))
                {
                    items.Remove(oldItem);
                    oldItem.Dispose();
                }
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
