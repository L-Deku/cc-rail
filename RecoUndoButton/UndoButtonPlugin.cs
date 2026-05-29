using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
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
            private const int MaxUndoSteps = 10;
            private readonly Form mainForm;
            private readonly Dictionary<string, ContextHistory> histories = new Dictionary<string, ContextHistory>();
            private readonly Timer settleTimer;
            private readonly List<ToolStripItem> undoItems = new List<ToolStripItem>();
            private readonly List<ToolStripItem> redoItems = new List<ToolStripItem>();
            private DataGridView grid;
            private string pendingContextKey;
            private BindingSource observedBindingSource;
            private DataTable observedTable;
            private TreeView observedTree;
            private Image undoIcon;
            private Image redoIcon;
            private DateTime suppressChangesUntilUtc;
            private GridSnapshot lastRestoredSnapshot;
            private string lastRestoredChapterKey;
            private bool replayingRestoredSnapshot;
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

                ResetCurrentContextSnapshot();
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
                grid.DataSourceChanged -= GridDataSourceChanged;
                grid.DataSourceChanged += GridDataSourceChanged;
                grid.KeyDown -= GridKeyDown;
                grid.KeyDown += GridKeyDown;
                HookDataSourceEvents();
                HookTreeEvents();
            }

            private void GridChanged(object sender, EventArgs e)
            {
                ScheduleChangeObserved();
            }

            private void GridDataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
            {
                HookDataSourceEvents();
                ScheduleRestoredSnapshotReplay("DataBindingComplete");
                EnsureCurrentContextSnapshot();
                UpdateUndoItems();
            }

            private void GridDataSourceChanged(object sender, EventArgs e)
            {
                HookDataSourceEvents();
                ScheduleRestoredSnapshotReplay("DataSourceChanged");
                ResetCurrentContextSnapshot();
            }

            private void TreeAfterSelect(object sender, TreeViewEventArgs e)
            {
                ScheduleRestoredSnapshotReplay("TreeAfterSelect");
                EnsureCurrentContextSnapshot();
                UpdateUndoItems();
            }

            private void GridKeyDown(object sender, KeyEventArgs e)
            {
                if (e.Control && (e.KeyCode == Keys.V || e.KeyCode == Keys.X || e.KeyCode == Keys.Delete))
                {
                    ScheduleChangeObserved();
                }
            }

            private void HookTreeEvents()
            {
                TreeView tree = UndoButtonPlugin.GetField<TreeView>(mainForm, "Tv_tree");
                if (observedTree == tree)
                {
                    return;
                }

                if (observedTree != null)
                {
                    observedTree.AfterSelect -= TreeAfterSelect;
                }

                observedTree = tree;
                if (observedTree != null)
                {
                    observedTree.AfterSelect -= TreeAfterSelect;
                    observedTree.AfterSelect += TreeAfterSelect;
                }
            }

            private void HookDataSourceEvents()
            {
                BindingSource bindingSource = grid == null ? null : grid.DataSource as BindingSource;
                DataView view = grid == null ? null : ResolveDataView(grid.DataSource);
                DataTable table = view == null ? ResolveDataTable(grid == null ? null : grid.DataSource) : view.Table;

                if (observedBindingSource != bindingSource)
                {
                    if (observedBindingSource != null)
                    {
                        observedBindingSource.ListChanged -= BindingSourceListChanged;
                    }

                    observedBindingSource = bindingSource;
                    if (observedBindingSource != null)
                    {
                        observedBindingSource.ListChanged -= BindingSourceListChanged;
                        observedBindingSource.ListChanged += BindingSourceListChanged;
                    }
                }

                if (observedTable != table)
                {
                    UnhookTableEvents(observedTable);
                    observedTable = table;
                    HookTableEvents(observedTable);
                }
            }

            private void HookTableEvents(DataTable table)
            {
                if (table == null)
                {
                    return;
                }

                table.ColumnChanging -= TableColumnChanging;
                table.ColumnChanging += TableColumnChanging;
                table.ColumnChanged -= TableColumnChanged;
                table.ColumnChanged += TableColumnChanged;
                table.RowChanging -= TableRowChanging;
                table.RowChanging += TableRowChanging;
                table.RowChanged -= TableRowChanged;
                table.RowChanged += TableRowChanged;
                table.RowDeleting -= TableRowDeleting;
                table.RowDeleting += TableRowDeleting;
                table.RowDeleted -= TableRowDeleted;
                table.RowDeleted += TableRowDeleted;
            }

            private void UnhookTableEvents(DataTable table)
            {
                if (table == null)
                {
                    return;
                }

                table.ColumnChanging -= TableColumnChanging;
                table.ColumnChanged -= TableColumnChanged;
                table.RowChanging -= TableRowChanging;
                table.RowChanged -= TableRowChanged;
                table.RowDeleting -= TableRowDeleting;
                table.RowDeleted -= TableRowDeleted;
            }

            private void TableColumnChanging(object sender, DataColumnChangeEventArgs e)
            {
                ScheduleChangeObserved();
            }

            private void TableColumnChanged(object sender, DataColumnChangeEventArgs e)
            {
                ScheduleChangeObserved();
            }

            private void TableRowChanging(object sender, DataRowChangeEventArgs e)
            {
                ScheduleChangeObserved();
            }

            private void TableRowChanged(object sender, DataRowChangeEventArgs e)
            {
                ScheduleChangeObserved();
            }

            private void TableRowDeleting(object sender, DataRowChangeEventArgs e)
            {
                ScheduleChangeObserved();
            }

            private void TableRowDeleted(object sender, DataRowChangeEventArgs e)
            {
                ScheduleChangeObserved();
            }

            private void BindingSourceListChanged(object sender, ListChangedEventArgs e)
            {
                HookDataSourceEvents();
                ScheduleChangeObserved();
            }

            private bool InstallToolbarItems()
            {
                RemoveExistingToolbarButtons(mainForm);
                RemoveFloatingToolbar(mainForm);

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
                if (restoring || DateTime.UtcNow < suppressChangesUntilUtc || grid == null || grid.IsDisposed)
                {
                    return;
                }

                string contextKey = GetCurrentContextKey();
                ContextHistory history = GetHistory(contextKey, true);
                if (history.LastStableSnapshot == null)
                {
                    history.LastStableSnapshot = CaptureSnapshot(contextKey);
                }

                if (history.PendingBeforeSnapshot == null)
                {
                    history.PendingBeforeSnapshot = history.LastStableSnapshot ?? CaptureSnapshot(contextKey);
                }

                pendingContextKey = contextKey;
                settleTimer.Stop();
                settleTimer.Start();
            }

            private void FlushPendingChange()
            {
                settleTimer.Stop();
                if (restoring || String.IsNullOrEmpty(pendingContextKey))
                {
                    return;
                }

                string currentContextKey = GetCurrentContextKey();
                ContextHistory history = GetHistory(pendingContextKey, false);
                if (history == null || history.PendingBeforeSnapshot == null)
                {
                    pendingContextKey = null;
                    return;
                }

                if (!String.Equals(currentContextKey, pendingContextKey, StringComparison.Ordinal))
                {
                    history.PendingBeforeSnapshot = null;
                    pendingContextKey = null;
                    UpdateUndoItems();
                    return;
                }

                GridSnapshot current = CaptureSnapshot(currentContextKey);
                if (!GridSnapshot.AreEquivalent(history.PendingBeforeSnapshot, current))
                {
                    PushUndo(history, history.PendingBeforeSnapshot);
                    history.RedoStack.Clear();
                    history.LastStableSnapshot = current;
                    ClearLastRestoredSnapshot();
                }

                history.PendingBeforeSnapshot = null;
                pendingContextKey = null;
                UpdateUndoItems();
            }

            private void PushUndo(ContextHistory history, GridSnapshot snapshot)
            {
                if (history == null || snapshot == null)
                {
                    return;
                }

                if (history.UndoStack.Count > 0 && GridSnapshot.AreEquivalent(history.UndoStack[history.UndoStack.Count - 1], snapshot))
                {
                    return;
                }

                history.UndoStack.Add(snapshot);
                while (history.UndoStack.Count > MaxUndoSteps)
                {
                    history.UndoStack.RemoveAt(0);
                }
            }

            private void PushRedo(ContextHistory history, GridSnapshot snapshot)
            {
                if (history == null || snapshot == null)
                {
                    return;
                }

                if (history.RedoStack.Count > 0 && GridSnapshot.AreEquivalent(history.RedoStack[history.RedoStack.Count - 1], snapshot))
                {
                    return;
                }

                history.RedoStack.Add(snapshot);
                while (history.RedoStack.Count > MaxUndoSteps)
                {
                    history.RedoStack.RemoveAt(0);
                }
            }

            private void ResetCurrentContextSnapshot()
            {
                if (restoring || grid == null || grid.IsDisposed)
                {
                    return;
                }

                string contextKey = GetCurrentContextKey();
                ContextHistory history = GetHistory(contextKey, true);
                history.LastStableSnapshot = CaptureSnapshot(contextKey);
                history.PendingBeforeSnapshot = null;
                if (String.Equals(pendingContextKey, contextKey, StringComparison.Ordinal))
                {
                    pendingContextKey = null;
                    settleTimer.Stop();
                }

                UpdateUndoItems();
            }

            private void EnsureCurrentContextSnapshot()
            {
                if (restoring || grid == null || grid.IsDisposed)
                {
                    return;
                }

                string contextKey = GetCurrentContextKey();
                ContextHistory history = GetHistory(contextKey, true);
                if (history.LastStableSnapshot == null)
                {
                    history.LastStableSnapshot = CaptureSnapshot(contextKey);
                }
            }

            private ContextHistory GetHistory(string contextKey, bool create)
            {
                if (String.IsNullOrEmpty(contextKey))
                {
                    contextKey = "unknown";
                }

                ContextHistory history;
                if (!histories.TryGetValue(contextKey, out history) && create)
                {
                    history = new ContextHistory();
                    histories[contextKey] = history;
                }

                return history;
            }

            private string GetCurrentContextKey()
            {
                string chapterKey = GetCurrentChapterKey();
                string sourceKey = GetCurrentDataSourceKey();
                return chapterKey + "|" + sourceKey;
            }

            private string GetCurrentChapterKey()
            {
                TreeView tree = UndoButtonPlugin.GetField<TreeView>(mainForm, "Tv_tree");
                if (tree == null || tree.SelectedNode == null)
                {
                    return "chapter:none";
                }

                return "chapter:" + tree.SelectedNode.FullPath;
            }

            private string GetCurrentDataSourceKey()
            {
                if (grid == null)
                {
                    return "source:none";
                }

                object source = grid.DataSource;
                DataView view = ResolveDataView(source);
                DataTable table = view == null ? ResolveDataTable(source) : view.Table;
                if (table != null)
                {
                    return "table:" + RuntimeHelpers.GetHashCode(table).ToString(CultureInfo.InvariantCulture) + ":" + table.TableName;
                }

                if (source != null)
                {
                    return "source:" + RuntimeHelpers.GetHashCode(source).ToString(CultureInfo.InvariantCulture) + ":" + source.GetType().FullName;
                }

                return "grid:" + RuntimeHelpers.GetHashCode(grid).ToString(CultureInfo.InvariantCulture);
            }

            private void UndoLastStep()
            {
                try
                {
                    EndGridEdit();
                    FlushPendingChange();
                    string contextKey = GetCurrentContextKey();
                    ContextHistory history = GetHistory(contextKey, true);
                    if (history.UndoStack.Count == 0)
                    {
                        MessageBox.Show(mainForm, "没有可撤回的操作。", "撤回上一步", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdateUndoItems();
                        return;
                    }

                    GridSnapshot current = CaptureSnapshot(contextKey);
                    GridSnapshot target = history.UndoStack[history.UndoStack.Count - 1];
                    if (!String.Equals(target.ContextKey, contextKey, StringComparison.Ordinal))
                    {
                        history.UndoStack.Clear();
                        UpdateUndoItems();
                        MessageBox.Show(mainForm, "当前条目与撤回记录不一致，已阻止跨条目撤回。", "撤回上一步", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    history.UndoStack.RemoveAt(history.UndoStack.Count - 1);
                    PushRedo(history, current);
                    RestoreSnapshot(target);
                    history.LastStableSnapshot = CaptureSnapshot(contextKey);
                    history.PendingBeforeSnapshot = null;
                    pendingContextKey = null;
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
                    string contextKey = GetCurrentContextKey();
                    ContextHistory history = GetHistory(contextKey, true);
                    if (history.RedoStack.Count == 0)
                    {
                        MessageBox.Show(mainForm, "没有可恢复的操作。", "恢复下一步", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        UpdateUndoItems();
                        return;
                    }

                    GridSnapshot current = CaptureSnapshot(contextKey);
                    GridSnapshot target = history.RedoStack[history.RedoStack.Count - 1];
                    if (!String.Equals(target.ContextKey, contextKey, StringComparison.Ordinal))
                    {
                        history.RedoStack.Clear();
                        UpdateUndoItems();
                        MessageBox.Show(mainForm, "当前条目与恢复记录不一致，已阻止跨条目恢复。", "恢复下一步", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }

                    history.RedoStack.RemoveAt(history.RedoStack.Count - 1);
                    PushUndo(history, current);
                    RestoreSnapshot(target);
                    history.LastStableSnapshot = CaptureSnapshot(contextKey);
                    history.PendingBeforeSnapshot = null;
                    pendingContextKey = null;
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
                return CaptureSnapshot(GetCurrentContextKey());
            }

            private GridSnapshot CaptureSnapshot(string contextKey)
            {
                GridSnapshot snapshot = new GridSnapshot();
                snapshot.ContextKey = contextKey;
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
                            object[] values;
                            if (!TryCopyDataRowValues(rowView.Row, out values))
                            {
                                continue;
                            }

                            snapshot.Rows.Add(values);
                        }
                    }
                    else
                    {
                        foreach (DataRow row in table.Rows)
                        {
                            object[] values;
                            if (!TryCopyDataRowValues(row, out values))
                            {
                                continue;
                            }

                            snapshot.Rows.Add(values);
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
                BeginRestoreEventSuppression();
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
                    RememberRestoredSnapshot(snapshot);
                    ClearPendingChange();
                }
                finally
                {
                    EndRestoreEventSuppression();
                }
            }

            private void BeginRestoreEventSuppression()
            {
                restoring = true;
                suppressChangesUntilUtc = DateTime.UtcNow.AddMilliseconds(900);
                ClearPendingChange();
            }

            private void EndRestoreEventSuppression()
            {
                ClearPendingChange();
                try
                {
                    mainForm.BeginInvoke((MethodInvoker)delegate
                    {
                        ClearPendingChange();
                        suppressChangesUntilUtc = DateTime.UtcNow.AddMilliseconds(250);
                        try
                        {
                            mainForm.BeginInvoke((MethodInvoker)delegate
                            {
                                ClearPendingChange();
                                restoring = false;
                            });
                        }
                        catch
                        {
                            restoring = false;
                        }
                    });
                }
                catch
                {
                    restoring = false;
                }
            }

            private void ClearPendingChange()
            {
                settleTimer.Stop();
                if (!String.IsNullOrEmpty(pendingContextKey))
                {
                    ContextHistory pendingHistory = GetHistory(pendingContextKey, false);
                    if (pendingHistory != null)
                    {
                        pendingHistory.PendingBeforeSnapshot = null;
                    }
                }

                pendingContextKey = null;
            }

            private void RememberRestoredSnapshot(GridSnapshot snapshot)
            {
                if (snapshot == null)
                {
                    return;
                }

                lastRestoredSnapshot = snapshot.Clone();
                lastRestoredChapterKey = GetCurrentChapterKey();
            }

            private void ClearLastRestoredSnapshot()
            {
                lastRestoredSnapshot = null;
                lastRestoredChapterKey = null;
            }

            private void ScheduleRestoredSnapshotReplay(string reason)
            {
                if (lastRestoredSnapshot == null || mainForm == null || mainForm.IsDisposed)
                {
                    return;
                }

                try
                {
                    mainForm.BeginInvoke((MethodInvoker)delegate { ReplayRestoredSnapshotIfNeeded(reason); });
                }
                catch
                {
                }
            }

            private void ReplayRestoredSnapshotIfNeeded(string reason)
            {
                if (lastRestoredSnapshot == null || restoring || replayingRestoredSnapshot || grid == null || grid.IsDisposed)
                {
                    return;
                }

                if (!String.Equals(lastRestoredChapterKey, GetCurrentChapterKey(), StringComparison.Ordinal))
                {
                    return;
                }

                GridSnapshot current = CaptureSnapshot(GetCurrentContextKey());
                if (GridSnapshot.AreEquivalent(lastRestoredSnapshot, current))
                {
                    return;
                }

                replayingRestoredSnapshot = true;
                try
                {
                    UndoButtonPlugin.Log("Replay restored snapshot after " + reason + ".");
                    RestoreSnapshot(lastRestoredSnapshot);
                    ContextHistory history = GetHistory(GetCurrentContextKey(), true);
                    history.LastStableSnapshot = CaptureSnapshot(GetCurrentContextKey());
                    history.PendingBeforeSnapshot = null;
                }
                finally
                {
                    replayingRestoredSnapshot = false;
                }
            }

            private void RestoreTableSnapshot(DataView view, DataTable table, GridSnapshot snapshot)
            {
                if (!IsSchemaCompatible(table, snapshot))
                {
                    throw new InvalidOperationException("当前表格列结构已经变化，不能安全撤回。");
                }

                RestoreRowsToTable(view, table, snapshot, grid.DataSource as BindingSource);
                SyncHostDeInputTable(snapshot, table);
                StabilizeRestoredTable(table);
                RefreshGridBinding();
            }

            private void RestoreRowsToTable(DataView view, DataTable table, GridSnapshot snapshot, BindingSource bindingSource)
            {
                bool restoreListChanged = false;
                bool oldRaiseListChanged = true;
                if (bindingSource != null)
                {
                    oldRaiseListChanged = bindingSource.RaiseListChangedEvents;
                    bindingSource.RaiseListChangedEvents = false;
                    restoreListChanged = true;
                }

                try
                {
                    table.BeginLoadData();
                    CommitDeletedRows(table);
                    List<DataRow> liveRows = GetLiveRows(view, table);
                    if (liveRows.Count == snapshot.Rows.Count)
                    {
                        for (int rowIndex = 0; rowIndex < snapshot.Rows.Count; rowIndex++)
                        {
                            ApplyValuesToDataRow(liveRows[rowIndex], snapshot.Rows[rowIndex]);
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
                }
                finally
                {
                    try
                    {
                        table.EndLoadData();
                    }
                    catch
                    {
                    }

                    if (restoreListChanged)
                    {
                        bindingSource.RaiseListChangedEvents = oldRaiseListChanged;
                    }
                }
            }

            private void SyncHostDeInputTable(GridSnapshot snapshot, DataTable restoredTable)
            {
                if (snapshot == null)
                {
                    return;
                }

                try
                {
                    FieldInfo field = mainForm.GetType().GetField("m_dtDeInput", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                    if (field == null)
                    {
                        return;
                    }

                    DataTable hostTable = field.GetValue(mainForm) as DataTable;
                    if (hostTable == null || Object.ReferenceEquals(hostTable, restoredTable))
                    {
                        return;
                    }

                    if (!IsSchemaCompatible(hostTable, snapshot))
                    {
                        UndoButtonPlugin.Log("Skip host m_dtDeInput sync: schema differs.");
                        return;
                    }

                    RestoreRowsToTable(null, hostTable, snapshot, null);
                    StabilizeRestoredTable(hostTable);
                    UndoButtonPlugin.Log("Synced host m_dtDeInput rows from restored snapshot.");
                }
                catch (Exception ex)
                {
                    UndoButtonPlugin.Log("Sync host m_dtDeInput failed: " + ex.Message);
                }
            }

            private static bool IsSchemaCompatible(DataTable table, GridSnapshot snapshot)
            {
                if (table == null || snapshot == null || table.Columns.Count != snapshot.Columns.Count)
                {
                    return false;
                }

                for (int i = 0; i < snapshot.Columns.Count; i++)
                {
                    if (!String.Equals(table.Columns[i].ColumnName, snapshot.Columns[i], StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }
                }

                return true;
            }

            private void StabilizeRestoredTable(DataTable table)
            {
                if (table == null)
                {
                    return;
                }

                try
                {
                    table.AcceptChanges();
                }
                catch (Exception ex)
                {
                    UndoButtonPlugin.Log("Accept restored table failed: " + ex.Message);
                }
            }

            private void RefreshGridBinding()
            {
                CurrencyManager manager = mainForm.BindingContext[grid.DataSource] as CurrencyManager;
                if (manager != null)
                {
                    try
                    {
                        manager.EndCurrentEdit();
                        manager.Refresh();
                    }
                    catch (Exception ex)
                    {
                        UndoButtonPlugin.Log("Refresh currency manager failed: " + ex.Message);
                    }
                }

                BindingSource bindingSource = grid.DataSource as BindingSource;
                if (bindingSource != null)
                {
                    try
                    {
                        bindingSource.EndEdit();
                        bindingSource.ResetBindings(false);
                    }
                    catch (Exception ex)
                    {
                        UndoButtonPlugin.Log("Refresh binding source failed: " + ex.Message);
                    }
                }

                grid.Refresh();
            }

            private static bool TryCopyDataRowValues(DataRow row, out object[] values)
            {
                values = null;
                if (!IsLiveDataRow(row))
                {
                    return false;
                }

                try
                {
                    values = row.ItemArray.ToArray();
                    return true;
                }
                catch (DeletedRowInaccessibleException)
                {
                    return false;
                }
                catch (RowNotInTableException)
                {
                    return false;
                }
            }

            private static bool IsLiveDataRow(DataRow row)
            {
                return row != null && row.RowState != DataRowState.Deleted && row.RowState != DataRowState.Detached;
            }

            private static List<DataRow> GetLiveRows(DataView view, DataTable table)
            {
                List<DataRow> rows = new List<DataRow>();
                if (view != null)
                {
                    foreach (DataRowView rowView in view)
                    {
                        try
                        {
                            if (IsLiveDataRow(rowView.Row))
                            {
                                rows.Add(rowView.Row);
                            }
                        }
                        catch (DeletedRowInaccessibleException)
                        {
                        }
                        catch (RowNotInTableException)
                        {
                        }
                    }

                    return rows;
                }

                foreach (DataRow row in table.Rows)
                {
                    if (IsLiveDataRow(row))
                    {
                        rows.Add(row);
                    }
                }

                return rows;
            }

            private static void CommitDeletedRows(DataTable table)
            {
                if (table == null)
                {
                    return;
                }

                foreach (DataRow row in table.Rows.Cast<DataRow>().ToArray())
                {
                    if (row.RowState == DataRowState.Deleted)
                    {
                        row.AcceptChanges();
                    }
                }
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
                    DataGridViewRow row = grid.Rows[snapshot.CurrentRowIndex];
                    DataGridViewColumn column = grid.Columns[snapshot.CurrentColumnIndex];
                    if (row == null || column == null || !row.Visible || !column.Visible)
                    {
                        return;
                    }

                    grid.CurrentCell = row.Cells[snapshot.CurrentColumnIndex];
                }
                catch
                {
                }
            }

            private void UpdateUndoItems()
            {
                ContextHistory history = GetHistory(GetCurrentContextKey(), true);
                bool undoEnabled = history.UndoStack.Count > 0;
                foreach (ToolStripItem item in undoItems.ToArray())
                {
                    if (item == null || item.IsDisposed)
                    {
                        undoItems.Remove(item);
                        continue;
                    }

                    item.Enabled = undoEnabled;
                    item.ToolTipText = undoEnabled ? "撤回上一步" : "没有可撤回的操作";
                }

                bool redoEnabled = history.RedoStack.Count > 0;
                foreach (ToolStripItem item in redoItems.ToArray())
                {
                    if (item == null || item.IsDisposed)
                    {
                        redoItems.Remove(item);
                        continue;
                    }

                    item.Enabled = redoEnabled;
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

            private static void RemoveFloatingToolbar(Control root)
            {
                foreach (ToolStrip toolStrip in FindControls<ToolStrip>(root).ToArray())
                {
                    if (toolStrip == null || toolStrip.IsDisposed)
                    {
                        continue;
                    }

                    if (!String.Equals(toolStrip.Name, "RecoUndoButton_FloatingToolbar", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    Control parent = toolStrip.Parent;
                    if (parent != null)
                    {
                        parent.Controls.Remove(toolStrip);
                    }

                    toolStrip.Dispose();
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

        private sealed class ContextHistory
        {
            public readonly List<GridSnapshot> UndoStack = new List<GridSnapshot>();
            public readonly List<GridSnapshot> RedoStack = new List<GridSnapshot>();
            public GridSnapshot LastStableSnapshot;
            public GridSnapshot PendingBeforeSnapshot;
        }

        private sealed class GridSnapshot
        {
            public string ContextKey;
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

            public GridSnapshot Clone()
            {
                GridSnapshot clone = new GridSnapshot();
                clone.ContextKey = ContextKey;
                clone.Mode = Mode;
                clone.CurrentRowIndex = CurrentRowIndex;
                clone.CurrentColumnIndex = CurrentColumnIndex;
                clone.Signature = Signature;
                clone.Columns.AddRange(Columns);
                foreach (object[] row in Rows)
                {
                    clone.Rows.Add(row == null ? null : row.ToArray());
                }

                return clone;
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
