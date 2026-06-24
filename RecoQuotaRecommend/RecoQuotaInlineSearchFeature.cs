using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace RecoQuotaRecommend
{
    internal sealed class QuotaInlineSearchFeature
    {
        private const int SearchDelayMs = 200;
        private const int VisibleRows = 8;
        private const int MaxCandidates = 100;
        private const int PreferredPopupWidth = 820;
        private const int MinimumPopupWidth = 420;
        private const int ResizeGripSize = 6;
        private const string AllCategories = "\u5168\u90e8";
        private static readonly Dictionary<Form, Runtime> Runtimes = new Dictionary<Form, Runtime>();

        public static void Install(Form mainForm)
        {
            if (mainForm == null || Runtimes.ContainsKey(mainForm)) return;
            DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
            if (grid == null)
            {
                QuotaRecommendPanel.Log("Inline quota search skipped: dataGridViewDE not found.");
                return;
            }
            Runtime runtime = new Runtime(mainForm, grid);
            if (!runtime.Install()) { runtime.Dispose(); return; }
            Runtimes[mainForm] = runtime;
            mainForm.FormClosed += delegate
            {
                Runtime existing;
                if (Runtimes.TryGetValue(mainForm, out existing))
                {
                    existing.Dispose();
                    Runtimes.Remove(mainForm);
                }
            };
            QuotaRecommendPanel.Log("Inline quota search installed.");
        }

        private sealed class BufferedCandidateGrid : DataGridView
        {
            public BufferedCandidateGrid()
            {
                DoubleBuffered = true;
                SetStyle(ControlStyles.OptimizedDoubleBuffer |
                    ControlStyles.AllPaintingInWmPaint |
                    ControlStyles.ResizeRedraw, true);
                UpdateStyles();
            }
        }

        private sealed class CandidatePopupForm : Form
        {
            private const int WsExNoActivate = 0x08000000;

            public CandidatePopupForm(DataGridView grid)
            {
                if (grid == null) throw new ArgumentNullException("grid");
                Text = "\u5019\u9009\u5b9a\u989d";
                StartPosition = FormStartPosition.Manual;
                FormBorderStyle = FormBorderStyle.SizableToolWindow;
                ShowInTaskbar = false;
                MinimizeBox = false;
                MaximizeBox = false;
                ControlBox = false;
                BackColor = SystemColors.Control;
                Padding = Padding.Empty;
                SizeGripStyle = SizeGripStyle.Show;
                grid.Dock = DockStyle.Fill;
                Controls.Add(grid);
            }

            protected override bool ShowWithoutActivation
            {
                get { return true; }
            }

            protected override CreateParams CreateParams
            {
                get
                {
                    CreateParams parameters = base.CreateParams;
                    parameters.ExStyle |= WsExNoActivate;
                    return parameters;
                }
            }

            public Size GetWindowSizeForClient(Size clientSize)
            {
                return SizeFromClientSize(clientSize);
            }
        }
        private sealed class Runtime : IDisposable
        {
            private readonly Form mainForm;
            private readonly DataGridView grid;
            private readonly Timer timer;
            private readonly DataGridView candidateGrid;
            private readonly CandidatePopupForm popup;
            private readonly SearchIndexStore searchIndex;
            private readonly ChapterLibraryStore chapterLibrary;
            private readonly Dictionary<string, bool> methodCache = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            private Size rememberedPopupSize = Size.Empty;
            private bool settingPopupSize;
            private TextBox editor;
            private int rowIndex = -1;
            private int colIndex = -1;
            private bool applying;
            private bool disposed;

            public Runtime(Form mainForm, DataGridView grid)
            {
                this.mainForm = mainForm;
                this.grid = grid;
                timer = new Timer();
                timer.Interval = SearchDelayMs;
                timer.Tick += TimerTick;
                candidateGrid = CreateCandidateGrid();
                popup = new CandidatePopupForm(candidateGrid);
                popup.SizeChanged += PopupSizeChanged;
                popup.Deactivate += PopupDeactivate;
                searchIndex = SearchIndexStore.LoadOrBuild();
                chapterLibrary = ChapterLibraryStore.Load();
            }

            private DataGridView CreateCandidateGrid()
            {
                DataGridView candidate = new BufferedCandidateGrid();
                candidate.BorderStyle = BorderStyle.FixedSingle;
                candidate.BackgroundColor = SystemColors.Window;
                candidate.Font = grid.Font;
                candidate.ReadOnly = true;
                candidate.AllowUserToAddRows = false;
                candidate.AllowUserToDeleteRows = false;
                candidate.AllowUserToResizeRows = false;
                candidate.AllowUserToOrderColumns = false;
                candidate.AutoGenerateColumns = false;
                candidate.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                candidate.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
                candidate.RowHeadersVisible = false;
                candidate.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                candidate.MultiSelect = false;
                candidate.ScrollBars = ScrollBars.Vertical;
                candidate.EditMode = DataGridViewEditMode.EditProgrammatically;
                candidate.CellBorderStyle = DataGridViewCellBorderStyle.Single;
                candidate.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
                candidate.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
                candidate.ColumnHeadersHeight = Math.Max(candidate.ColumnHeadersHeight, 22);
                candidate.RowTemplate.Height = Math.Max(candidate.RowTemplate.Height, 21);
                candidate.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                candidate.CellDoubleClick += CandidateGridCellDoubleClick;
                candidate.KeyDown += CandidateGridKeyDown;

                AddCandidateColumn(candidate, "QuotaCode", "\u7f16\u53f7", 88, 0f);
                AddCandidateColumn(candidate, "QuotaName", "\u540d\u79f0", 0, 58f);
                AddCandidateColumn(candidate, "QuotaUnit", "\u5355\u4f4d", 56, 0f);
                AddCandidateColumn(candidate, "BasePrice", "\u57fa\u671f\u4ef7\u683c", 82, 0f);
                AddCandidateColumn(candidate, "WorkContent", "\u5185\u5bb9", 0, 42f);
                candidate.Columns["QuotaUnit"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                candidate.Columns["BasePrice"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                candidate.Columns["BasePrice"].DefaultCellStyle.Format = "0.##";
                return candidate;
            }

            private static void AddCandidateColumn(
                DataGridView candidate,
                string name,
                string header,
                int width,
                float fillWeight)
            {
                DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
                column.Name = name;
                column.HeaderText = header;
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
                if (fillWeight > 0f)
                {
                    column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                    column.FillWeight = fillWeight;
                    column.MinimumWidth = name == "QuotaName" ? 220 : 180;
                }
                else
                {
                    column.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                    column.Width = width;
                }
                candidate.Columns.Add(column);
            }

            public bool Install()
            {
                if (!HasColumn(QuotaNameColumns()) || !HasColumn(QuotaCodeColumns()))
                {
                    QuotaRecommendPanel.Log("Inline quota search skipped: required columns not found.");
                    return false;
                }
                grid.EditingControlShowing -= EditingControlShowing;
                grid.EditingControlShowing += EditingControlShowing;
                grid.CurrentCellChanged -= CloseOnGridMove;
                grid.CurrentCellChanged += CloseOnGridMove;
                grid.CellEndEdit -= CellEndEdit;
                grid.CellEndEdit += CellEndEdit;
                grid.Scroll -= GridScroll;
                grid.Scroll += GridScroll;
                grid.KeyDown -= GridKeyDown;
                grid.KeyDown += GridKeyDown;
                grid.DataSourceChanged -= CloseOnGridMove;
                grid.DataSourceChanged += CloseOnGridMove;
                mainForm.Deactivate -= MainFormDeactivate;
                mainForm.Deactivate += MainFormDeactivate;
                return true;
            }

            public void Dispose()
            {
                if (disposed) return;
                disposed = true;
                DetachEditor();
                HidePopup();
                grid.EditingControlShowing -= EditingControlShowing;
                grid.CurrentCellChanged -= CloseOnGridMove;
                grid.CellEndEdit -= CellEndEdit;
                grid.Scroll -= GridScroll;
                grid.KeyDown -= GridKeyDown;
                grid.DataSourceChanged -= CloseOnGridMove;
                mainForm.Deactivate -= MainFormDeactivate;
                timer.Stop();
                timer.Dispose();
                popup.SizeChanged -= PopupSizeChanged;
                popup.Deactivate -= PopupDeactivate;
                popup.Dispose();
            }

            private void EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
            {
                DetachEditor();
                if (applying || grid.CurrentCell == null || !IsNameColumn(grid.CurrentCell.ColumnIndex) || !IsBlankRow(grid.CurrentCell.RowIndex))
                {
                    HidePopup();
                    return;
                }
                editor = e.Control as TextBox;
                if (editor == null) { HidePopup(); return; }
                rowIndex = grid.CurrentCell.RowIndex;
                colIndex = grid.CurrentCell.ColumnIndex;
                editor.TextChanged += EditorTextChanged;
                editor.KeyDown += EditorKeyDown;
                ScheduleSearch();
            }

            private void CloseOnGridMove(object sender, EventArgs e)
            {
                if (applying || IsInteractingWithCandidateGrid()) return;
                DetachEditor();
                HidePopup();
            }

            private void MainFormDeactivate(object sender, EventArgs e)
            {
                if (applying || IsInteractingWithCandidateGrid()) return;
                DetachEditor();
                HidePopup();
            }

            private void CellEndEdit(object sender, DataGridViewCellEventArgs e)
            {
                if (applying || IsInteractingWithCandidateGrid()) return;
                DetachEditor();
                HidePopup();
            }

            private void GridScroll(object sender, ScrollEventArgs e) { HidePopup(); }
            private void EditorTextChanged(object sender, EventArgs e) { ScheduleSearch(); }
            private void EditorKeyDown(object sender, KeyEventArgs e) { HandleKeys(e); }
            private void GridKeyDown(object sender, KeyEventArgs e) { HandleKeys(e); }
            private void CandidateGridKeyDown(object sender, KeyEventArgs e) { HandleKeys(e); }

            private bool IsInteractingWithCandidateGrid()
            {
                if (!popup.Visible || candidateGrid.IsDisposed) return false;
                if (popup.ContainsFocus || candidateGrid.Focused || candidateGrid.ContainsFocus) return true;
                return popup.Bounds.Contains(Control.MousePosition);
            }

            private void HandleKeys(KeyEventArgs e)
            {
                if (e == null || !popup.Visible) return;
                if (e.KeyCode == Keys.Escape)
                {
                    HidePopup();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    return;
                }
                if (e.KeyCode == Keys.Down || e.KeyCode == Keys.Up)
                {
                    MoveSelection(e.KeyCode == Keys.Down ? 1 : -1);
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    return;
                }
                if (e.KeyCode == Keys.Enter)
                {
                    ApplySelected();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            }

            private void MoveSelection(int delta)
            {
                if (candidateGrid.Rows.Count == 0) return;
                int current = candidateGrid.CurrentCell == null ? 0 : candidateGrid.CurrentCell.RowIndex;
                int index = Math.Max(0, Math.Min(candidateGrid.Rows.Count - 1, current + delta));
                candidateGrid.ClearSelection();
                candidateGrid.CurrentCell = candidateGrid.Rows[index].Cells[0];
                candidateGrid.Rows[index].Selected = true;
            }

            private void CandidateGridCellDoubleClick(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0 || e.RowIndex >= candidateGrid.Rows.Count) return;
                candidateGrid.CurrentCell = candidateGrid.Rows[e.RowIndex].Cells[0];
                candidateGrid.Rows[e.RowIndex].Selected = true;
                ApplySelected();
            }

            private void ScheduleSearch()
            {
                timer.Stop();
                timer.Start();
            }

            private void TimerTick(object sender, EventArgs e)
            {
                timer.Stop();
                RefreshCandidates();
            }

            private void RefreshCandidates()
            {
                if (disposed || applying || editor == null || grid.CurrentCell == null || rowIndex < 0 || rowIndex >= grid.Rows.Count)
                {
                    HidePopup();
                    return;
                }
                if (!IsNameColumn(colIndex) || !IsBlankRow(rowIndex))
                {
                    HidePopup();
                    return;
                }
                string query = NormalizeText(editor.Text);
                if (!LooksSearchable(query))
                {
                    HidePopup();
                    return;
                }
                try
                {
                    ExcelQuantityItem item = BuildItem(query, rowIndex);
                    EntryScope scope = ResolveCurrentScope();
                    List<RecommendationRow> rows = searchIndex.SearchQuotaCandidates(item, AllCategories, scope, MaxCandidates);
                    if (rows.Count == 0) { HidePopup(); return; }
                    ShowRows(rows);
                }
                catch (Exception ex)
                {
                    HidePopup();
                    QuotaRecommendPanel.Log("Inline quota search failed: " + ex.Message);
                }
            }

            private ExcelQuantityItem BuildItem(string query, int sourceRow)
            {
                DataGridViewRow row = sourceRow >= 0 && sourceRow < grid.Rows.Count ? grid.Rows[sourceRow] : null;
                ExcelQuantityItem item = new ExcelQuantityItem();
                item.Name = query;
                item.OriginalName = query;
                item.Unit = row == null ? "" : GetRowValue(row, QuotaUnitColumns());
                item.ValueText = row == null ? "" : GetRowValue(row, QuantityColumns());
                item.ContextText = query + " " + item.Unit + " " + item.ValueText;
                item.RawRowText = item.ContextText;
                item.SkipAiNameNormalization = true;
                return item;
            }

            private void ShowRows(List<RecommendationRow> rows)
            {
                candidateGrid.SuspendLayout();
                try
                {
                    candidateGrid.Rows.Clear();
                    foreach (RecommendationRow row in rows)
                    {
                        int index = candidateGrid.Rows.Add(
                            row.QuotaCode ?? "",
                            row.QuotaName ?? "",
                            row.QuotaUnit ?? "",
                            row.BasePrice,
                            row.WorkContent ?? "");
                        candidateGrid.Rows[index].Tag = row;
                    }
                    if (candidateGrid.Rows.Count > 0)
                    {
                        candidateGrid.ClearSelection();
                        candidateGrid.CurrentCell = candidateGrid.Rows[0].Cells[0];
                        candidateGrid.Rows[0].Selected = true;
                    }
                }
                finally { candidateGrid.ResumeLayout(); }

                Rectangle rect = grid.GetCellDisplayRectangle(colIndex, rowIndex, true);
                if (rect.Width <= 0 || rect.Height <= 0) { HidePopup(); return; }
                Rectangle workingArea = Screen.FromControl(grid).WorkingArea;
                int width = Math.Min(
                    Math.Max(rect.Width, PreferredPopupWidth),
                    Math.Max(420, workingArea.Width - 16));
                int rowHeight = Math.Max(candidateGrid.RowTemplate.Height, 21);
                int visibleCount = Math.Min(VisibleRows, candidateGrid.Rows.Count);
                int height = candidateGrid.ColumnHeadersHeight + visibleCount * rowHeight + 3;
                int minimumHeight = candidateGrid.ColumnHeadersHeight + rowHeight + 3;
                popup.MinimumSize = popup.GetWindowSizeForClient(new Size(MinimumPopupWidth, minimumHeight));
                Size defaultPopupSize = popup.GetWindowSizeForClient(new Size(width, height));
                Size popupSize = ResolvePopupSize(defaultPopupSize, workingArea);
                if (popup.Visible) popup.Hide();
                Point screenLocation = grid.PointToScreen(new Point(rect.Left, rect.Bottom));
                int overflow = screenLocation.X + popupSize.Width - workingArea.Right;
                int popupLeft = overflow > 0 ? Math.Max(workingArea.Left, screenLocation.X - overflow) : screenLocation.X;
                int popupTop = Math.Min(screenLocation.Y, Math.Max(workingArea.Top, workingArea.Bottom - popupSize.Height));
                settingPopupSize = true;
                try
                {
                    popup.Size = popupSize;
                    popup.Location = new Point(popupLeft, popupTop);
                    ResizePopupContent();
                }
                finally
                {
                    settingPopupSize = false;
                }
                if (!popup.Visible) popup.Show(mainForm);
                popup.Location = new Point(popupLeft, popupTop);
                ResizePopupContent();
                editor.Focus();
            }

            private Size ResolvePopupSize(Size defaultSize, Rectangle workingArea)
            {
                Size requested = rememberedPopupSize.IsEmpty ? defaultSize : rememberedPopupSize;
                int minimumWidth = Math.Max(1, popup.MinimumSize.Width);
                int minimumHeight = Math.Max(1, popup.MinimumSize.Height);
                int maximumWidth = Math.Max(minimumWidth, workingArea.Width - 16);
                int maximumHeight = Math.Max(minimumHeight, workingArea.Height - 16);
                return new Size(
                    Math.Min(maximumWidth, Math.Max(minimumWidth, requested.Width)),
                    Math.Min(maximumHeight, Math.Max(minimumHeight, requested.Height)));
            }

            private void PopupSizeChanged(object sender, EventArgs e)
            {
                ResizePopupContent();
                if (popup.Visible && !settingPopupSize)
                {
                    RememberCurrentPopupSize();
                }
            }

            private void PopupDeactivate(object sender, EventArgs e)
            {
                if (applying || IsInteractingWithCandidateGrid()) return;
                HidePopup();
            }

            private void ResizePopupContent()
            {
                if (candidateGrid.Dock != DockStyle.Fill) candidateGrid.Dock = DockStyle.Fill;
                popup.PerformLayout();
                candidateGrid.Invalidate();
                popup.Invalidate();
            }

            private void RememberCurrentPopupSize()
            {
                if (!popup.IsDisposed && popup.Size.Width > 0 && popup.Size.Height > 0)
                {
                    rememberedPopupSize = popup.Size;
                }
            }

            private void ApplySelected()
            {
                if (candidateGrid.CurrentCell == null) return;
                int selectedIndex = candidateGrid.CurrentCell.RowIndex;
                if (selectedIndex < 0 || selectedIndex >= candidateGrid.Rows.Count) return;
                RecommendationRow row = candidateGrid.Rows[selectedIndex].Tag as RecommendationRow;
                if (row == null || String.IsNullOrWhiteSpace(row.QuotaCode)) return;
                try { mainForm.BeginInvoke((MethodInvoker)delegate { ApplyRecommendation(row); }); }
                catch { ApplyRecommendation(row); }
            }

            private void ApplyRecommendation(RecommendationRow recommendation)
            {
                if (recommendation == null || String.IsNullOrWhiteSpace(recommendation.QuotaCode) || rowIndex < 0 || rowIndex >= grid.Rows.Count)
                {
                    HidePopup();
                    return;
                }

                int codeColumn = FindColumnIndex(QuotaCodeColumns());
                if (codeColumn < 0)
                {
                    HidePopup();
                    QuotaRecommendPanel.Log("Inline quota native paste failed: quota code column not found.");
                    return;
                }

                applying = true;
                Stopwatch applyTimer = Stopwatch.StartNew();
                try
                {
                    HidePopup();
                    DetachEditor();
                    try { grid.CancelEdit(); }
                    catch { }

                    if (TryApplyViaNativeEnterCommit(recommendation))
                    {
                        QuotaRecommendPanel.Log("Inline quota apply completed via native single-enter: " + recommendation.QuotaCode.Trim()
                            + " totalMs=" + applyTimer.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture));
                        return;
                    }

                    if (TryApplyViaQuotaCodeCell(recommendation))
                    {
                        QuotaRecommendPanel.Log("Inline quota apply completed: " + recommendation.QuotaCode.Trim()
                            + " totalMs=" + applyTimer.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture));
                        return;
                    }

                    QuotaRecommendPanel.Log("Inline quota apply failed: native single-enter and keyboard fallback both failed for " + recommendation.QuotaCode);
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Inline quota apply failed: " + ex.Message);
                }
                finally
                {
                    applying = false;
                    rowIndex = -1;
                    colIndex = -1;
                }
            }

            private bool TryApplyViaNativeEnterCommit(RecommendationRow recommendation)
            {
                if (recommendation == null || String.IsNullOrWhiteSpace(recommendation.QuotaCode))
                {
                    return false;
                }

                int targetRowIndex = rowIndex >= 0 && rowIndex < grid.Rows.Count ? rowIndex : -1;
                int codeColumn = FindColumnIndex(QuotaCodeColumns());
                if (targetRowIndex < 0 || codeColumn < 0 || codeColumn >= grid.Columns.Count)
                {
                    return false;
                }

                Stopwatch timer = Stopwatch.StartNew();
                try
                {
                    bool targetWasLastRow = targetRowIndex == grid.Rows.Count - 1;
                    grid.Focus();
                    grid.ClearSelection();
                    grid.CurrentCell = grid.Rows[targetRowIndex].Cells[codeColumn];
                    grid.Rows[targetRowIndex].Selected = true;
                    bool beganEdit = grid.BeginEdit(true);
                    long beginEditMs = timer.ElapsedMilliseconds;
                    TextBoxBase editControl = grid.EditingControl as TextBoxBase;
                    if (!beganEdit || editControl == null)
                    {
                        QuotaRecommendPanel.Log("Inline quota native edit unavailable: " + recommendation.QuotaCode.Trim()
                            + " row=" + targetRowIndex.ToString(CultureInfo.InvariantCulture)
                            + " beganEdit=" + beganEdit.ToString()
                            + " elapsedMs=" + timer.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture));
                        return false;
                    }

                    editControl.Text = recommendation.QuotaCode.Trim();
                    editControl.SelectionStart = editControl.TextLength;
                    editControl.SelectionLength = 0;
                    grid.NotifyCurrentCellDirty(true);
                    SendKeys.SendWait("{ENTER}");
                    long enterMs = timer.ElapsedMilliseconds - beginEditMs;
                    Application.DoEvents();

                    bool filled = TargetRowLooksFullyNativeCommitted(targetRowIndex, recommendation, targetWasLastRow);
                    QuotaRecommendPanel.Log("Inline quota native single-enter submitted: " + recommendation.QuotaCode.Trim()
                        + " row=" + targetRowIndex.ToString(CultureInfo.InvariantCulture)
                        + " filled=" + filled.ToString()
                        + " beginEditMs=" + beginEditMs.ToString(CultureInfo.InvariantCulture)
                        + " enterMs=" + enterMs.ToString(CultureInfo.InvariantCulture)
                        + " elapsedMs=" + timer.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture)
                        + " data=" + DescribeInputRow(targetRowIndex));
                    return filled;
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Inline quota native single-enter failed: " + ex.Message);
                    return false;
                }
            }

            private bool TryApplyViaQuotaCodeCell(RecommendationRow recommendation)
            {
                if (recommendation == null || String.IsNullOrWhiteSpace(recommendation.QuotaCode))
                {
                    return false;
                }

                int targetRowIndex = rowIndex >= 0 && rowIndex < grid.Rows.Count ? rowIndex : -1;
                int codeColumn = FindColumnIndex(QuotaCodeColumns());
                if (targetRowIndex < 0 || codeColumn < 0 || codeColumn >= grid.Columns.Count)
                {
                    return false;
                }

                try
                {
                    bool targetWasLastRow = targetRowIndex == grid.Rows.Count - 1;
                    Stopwatch stageTimer = Stopwatch.StartNew();
                    grid.Focus();
                    grid.ClearSelection();
                    grid.CurrentCell = grid.Rows[targetRowIndex].Cells[codeColumn];
                    grid.Rows[targetRowIndex].Selected = true;
                    Application.DoEvents();
                    long focusMs = stageTimer.ElapsedMilliseconds;

                    grid.BeginEdit(true);
                    Application.DoEvents();
                    long editMs = stageTimer.ElapsedMilliseconds - focusMs;

                    Clipboard.SetText(recommendation.QuotaCode.Trim());
                    long clipboardMs = stageTimer.ElapsedMilliseconds - focusMs - editMs;

                    SendKeys.SendWait("^a^v{ENTER}");
                    Application.DoEvents();
                    long submitMs = stageTimer.ElapsedMilliseconds - focusMs - editMs - clipboardMs;

                    bool filled = TargetRowLooksFullyNativeCommitted(targetRowIndex, recommendation, targetWasLastRow);
                    long verifyMs = stageTimer.ElapsedMilliseconds - focusMs - editMs - clipboardMs - submitMs;
                    QuotaRecommendPanel.Log("Inline quota code-cell submitted: " + recommendation.QuotaCode.Trim()
                        + " row=" + targetRowIndex.ToString(CultureInfo.InvariantCulture)
                        + " filled=" + filled.ToString()
                        + " focusMs=" + focusMs.ToString(CultureInfo.InvariantCulture)
                        + " editMs=" + editMs.ToString(CultureInfo.InvariantCulture)
                        + " clipboardMs=" + clipboardMs.ToString(CultureInfo.InvariantCulture)
                        + " submitMs=" + submitMs.ToString(CultureInfo.InvariantCulture)
                        + " verifyMs=" + verifyMs.ToString(CultureInfo.InvariantCulture)
                        + " nativeTotalMs=" + stageTimer.ElapsedMilliseconds.ToString(CultureInfo.InvariantCulture)
                        + " data=" + DescribeInputRow(targetRowIndex));
                    return filled;
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Inline quota code-cell submit failed: " + ex.Message);
                    return false;
                }
            }

            private string DescribeInputRow(int targetRowIndex)
            {
                if (targetRowIndex < 0 || targetRowIndex >= grid.Rows.Count)
                {
                    return "<out-of-range>";
                }

                DataGridViewRow row = grid.Rows[targetRowIndex];
                return "code=" + GetRowValue(row, QuotaCodeColumns())
                    + "|name=" + GetRowValue(row, QuotaNameColumns())
                    + "|unit=" + GetRowValue(row, QuotaUnitColumns())
                    + "|qty=" + GetRowValue(row, QuantityColumns())
                    + "|price=" + GetRowValue(row, UnitPriceColumns())
                    + "|weight=" + GetRowValue(row, SingleWeightColumns())
                    + "|compiler=" + GetRowValue(row, CompilerColumns())
                    + "|modified=" + GetRowValue(row, ModifiedDateColumns())
                    + "|rows=" + grid.Rows.Count.ToString(CultureInfo.InvariantCulture)
                    + "|nextBlank=" + IsBlankRow(targetRowIndex + 1).ToString();
            }

            private bool TargetRowLooksFullyNativeCommitted(int targetRowIndex, RecommendationRow recommendation, bool targetWasLastRow)
            {
                if (!TargetRowLooksNativeFilled(targetRowIndex, recommendation))
                {
                    return false;
                }

                DataGridViewRow targetRow = grid.Rows[targetRowIndex];
                if (!ColumnValuePresentWhenAvailable(targetRow, CompilerColumns()) ||
                    !ColumnValuePresentWhenAvailable(targetRow, ModifiedDateColumns()))
                {
                    return false;
                }

                return !targetWasLastRow || IsBlankRow(targetRowIndex + 1);
            }

            private bool ColumnValuePresentWhenAvailable(DataGridViewRow row, string[] columns)
            {
                return FindColumnIndex(columns) < 0 || !String.IsNullOrWhiteSpace(GetRowValue(row, columns));
            }

            private bool TargetRowLooksNativeFilled(int targetRowIndex, RecommendationRow recommendation)
            {
                if (targetRowIndex < 0 || targetRowIndex >= grid.Rows.Count || recommendation == null)
                {
                    return false;
                }

                DataGridViewRow targetRow = grid.Rows[targetRowIndex];
                string actualCode = GetRowValue(targetRow, QuotaCodeColumns());
                if (!String.Equals(QuotaEntry.NormalizeCode(actualCode), QuotaEntry.NormalizeCode(recommendation.QuotaCode), StringComparison.OrdinalIgnoreCase))
                {
                    return false;
                }

                string actualName = GetRowValue(targetRow, QuotaNameColumns());
                if (String.IsNullOrWhiteSpace(actualName))
                {
                    return false;
                }

                string expectedName = TextMatcher.Normalize(recommendation.QuotaName).Replace(" ", "");
                string normalizedActualName = TextMatcher.Normalize(actualName).Replace(" ", "");
                if (!String.IsNullOrWhiteSpace(expectedName) &&
                    !normalizedActualName.Contains(expectedName) &&
                    !expectedName.Contains(normalizedActualName))
                {
                    return false;
                }

                if (!String.IsNullOrWhiteSpace(recommendation.QuotaUnit) &&
                    String.IsNullOrWhiteSpace(GetRowValue(targetRow, QuotaUnitColumns())))
                {
                    return false;
                }

                if (recommendation.BasePrice > 0.000001d &&
                    String.IsNullOrWhiteSpace(GetRowValue(targetRow, UnitPriceColumns())))
                {
                    return false;
                }

                return true;
            }

            private void HidePopup()
            {
                timer.Stop();
                if (popup.Visible) popup.Hide();
                candidateGrid.Rows.Clear();
            }

            private void DetachEditor()
            {
                if (editor == null) return;
                editor.TextChanged -= EditorTextChanged;
                editor.KeyDown -= EditorKeyDown;
                editor = null;
            }

            private bool IsBlankRow(int index)
            {
                if (index < 0 || index >= grid.Rows.Count) return false;
                DataGridViewRow row = grid.Rows[index];
                if (row == null || row.IsNewRow) return false;
                return String.IsNullOrWhiteSpace(GetRowValue(row, QuotaCodeColumns()));
            }

            private bool IsNameColumn(int index)
            {
                return index >= 0 && index < grid.Columns.Count && ColumnMatches(grid.Columns[index], QuotaNameColumns());
            }

            private bool HasColumn(string[] names) { return FindColumnIndex(names) >= 0; }

            private int FindColumnIndex(string[] names)
            {
                foreach (DataGridViewColumn column in grid.Columns)
                {
                    if (ColumnMatches(column, names)) return column.Index;
                }
                return -1;
            }

            private EntryScope ResolveCurrentScope()
            {
                if (chapterLibrary == null || chapterLibrary.IsEmpty) return null;
                try
                {
                    System.Data.SqlClient.SqlConnection conn = GetField<System.Data.SqlClient.SqlConnection>(mainForm, "m_ProjectConn");
                    if (conn != null && !ProjectUsesLibraryMethod(conn)) return null;
                    string entryName;
                    string entryCode = ResolveCurrentChapterNo(conn, out entryName);
                    return String.IsNullOrWhiteSpace(entryCode) ? null : chapterLibrary.ResolveScope(entryCode, entryName);
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Inline quota scope resolve failed: " + ex.Message);
                    return null;
                }
            }

            private bool ProjectUsesLibraryMethod(System.Data.SqlClient.SqlConnection conn)
            {
                if (conn == null || String.IsNullOrWhiteSpace(chapterLibrary.MethodNo)) return true;
                string dbName;
                try { dbName = conn.Database ?? ""; }
                catch { return true; }
                bool cached;
                if (methodCache.TryGetValue(dbName, out cached)) return cached;
                bool matches = true;
                try
                {
                    EnsureConnectionOpen(conn);
                    using (System.Data.SqlClient.SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "select \u7f16\u5236\u529e\u6cd5\u6587\u53f7 from \u9879\u76ee\u4fe1\u606f";
                        object result = cmd.ExecuteScalar();
                        string methodNo = result == null || result == DBNull.Value ? "" : Convert.ToString(result, CultureInfo.InvariantCulture).Trim();
                        if (!String.IsNullOrEmpty(methodNo)) matches = String.Equals(NormalizeMethodNo(methodNo), NormalizeMethodNo(chapterLibrary.MethodNo), StringComparison.OrdinalIgnoreCase);
                    }
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Inline quota project method check failed (treat as match): " + ex.Message);
                }
                methodCache[dbName] = matches;
                return matches;
            }

            private string ResolveCurrentChapterNo(System.Data.SqlClient.SqlConnection conn, out string entryName)
            {
                entryName = "";
                if (grid.CurrentRow != null && !grid.CurrentRow.IsNewRow)
                {
                    string seq = GetRowValue(grid.CurrentRow, new[] { "\u6761\u76ee\u5e8f\u53f7" });
                    string fromSeq = LookupChapterNoBySeq(conn, seq, out entryName);
                    if (!String.IsNullOrEmpty(fromSeq)) return fromSeq;
                }
                string fromProp = ReadPropertyGridValue("\u6761\u76ee\u7f16\u53f7");
                if (!String.IsNullOrEmpty(fromProp))
                {
                    entryName = ReadPropertyGridValue("\u5de5\u7a0b\u6216\u8d39\u7528\u9879\u76ee\u540d\u79f0") ?? "";
                    return fromProp;
                }
                TreeView tree = GetField<TreeView>(mainForm, "Tv_tree");
                TreeNode node = tree != null ? tree.SelectedNode : GetField<TreeNode>(mainForm, "CurrNode");
                if (node != null)
                {
                    string fromTag = TryGetTagValue(node.Tag, "\u6761\u76ee\u7f16\u53f7");
                    if (!String.IsNullOrEmpty(fromTag)) { entryName = node.Text ?? ""; return fromTag; }
                    string seq = TryGetTagValue(node.Tag, "\u6761\u76ee\u5e8f\u53f7");
                    if (String.IsNullOrEmpty(seq) && IsAllDigits(node.Name)) seq = node.Name;
                    string fromSeq = LookupChapterNoBySeq(conn, seq, out entryName);
                    if (!String.IsNullOrEmpty(fromSeq))
                    {
                        if (String.IsNullOrEmpty(entryName)) entryName = node.Text ?? "";
                        return fromSeq;
                    }
                    if (!String.IsNullOrEmpty(node.Name) && !IsAllDigits(node.Name))
                    {
                        entryName = node.Text ?? "";
                        return node.Name;
                    }
                }
                return null;
            }

            private static string LookupChapterNoBySeq(System.Data.SqlClient.SqlConnection conn, string seq, out string entryName)
            {
                entryName = "";
                if (String.IsNullOrWhiteSpace(seq) || conn == null) return null;
                EnsureConnectionOpen(conn);
                using (System.Data.SqlClient.SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "select \u6761\u76ee\u7f16\u53f7, \u5de5\u7a0b\u6216\u8d39\u7528\u9879\u76ee\u540d\u79f0 from \u7ae0\u8282\u8868 where \u6761\u76ee\u5e8f\u53f7=@id";
                    cmd.Parameters.AddWithValue("@id", seq.Trim());
                    using (System.Data.SqlClient.SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (!reader.Read()) return null;
                        string code = reader.IsDBNull(0) ? null : Convert.ToString(reader.GetValue(0), CultureInfo.InvariantCulture).Trim();
                        entryName = reader.IsDBNull(1) ? "" : Convert.ToString(reader.GetValue(1), CultureInfo.InvariantCulture).Trim();
                        return code;
                    }
                }
            }

            private string ReadPropertyGridValue(string propertyName)
            {
                DataGridView propGrid = GetField<DataGridView>(mainForm, "dataGridViewProp");
                if (propGrid == null) return null;
                foreach (DataGridViewRow row in propGrid.Rows)
                {
                    if (row.IsNewRow) continue;
                    if (String.Equals(GetRowValue(row, new[] { "\u5c5e\u6027\u540d\u79f0" }), propertyName, StringComparison.Ordinal))
                    {
                        return GetRowValue(row, new[] { "\u6570\u636e" });
                    }
                }
                return null;
            }
        }

        private static bool LooksSearchable(string text)
        {
            string value = NormalizeText(text);
            if (String.IsNullOrWhiteSpace(value)) return false;
            if (ContainsChinese(value)) return true;
            return value.Length >= 2 && value.Any(Char.IsLetterOrDigit);
        }

        private static bool ContainsChinese(string text)
        {
            foreach (char ch in text ?? "") if (ch >= 0x4e00 && ch <= 0x9fff) return true;
            return false;
        }

        private static string NormalizeText(string text) { return (text ?? "").Replace("\r", " ").Replace("\n", " ").Trim(); }
        private static string NormalizeMethodNo(string text) { return (text ?? "").Replace('\u2013', '-').Replace('\u2014', '-').Replace('\uff0d', '-').Replace(" ", "").Trim(); }
        private static void EnsureConnectionOpen(System.Data.SqlClient.SqlConnection conn) { if (conn != null && conn.State != ConnectionState.Open) conn.Open(); }
        private static bool IsAllDigits(string text) { return !String.IsNullOrEmpty(text) && text.All(Char.IsDigit); }

        private static string TryGetTagValue(object source, string name)
        {
            if (source == null) return null;
            DataRowView rowView = source as DataRowView;
            if (rowView != null && rowView.DataView.Table.Columns.Contains(name)) return Convert.ToString(rowView[name], CultureInfo.InvariantCulture);
            DataRow dataRow = source as DataRow;
            if (dataRow != null && dataRow.Table.Columns.Contains(name)) return Convert.ToString(dataRow[name], CultureInfo.InvariantCulture);
            PropertyInfo prop = source.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
            if (prop != null)
            {
                object value = prop.GetValue(source, null);
                return value == null ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
            }
            return null;
        }

        private static bool ColumnMatches(DataGridViewColumn column, string[] names)
        {
            foreach (string name in names)
            {
                if (String.Equals(column.DataPropertyName, name, StringComparison.OrdinalIgnoreCase) ||
                    String.Equals(column.Name, name, StringComparison.OrdinalIgnoreCase) ||
                    String.Equals(column.HeaderText, name, StringComparison.OrdinalIgnoreCase)) return true;
            }
            return false;
        }

        private static string GetRowValue(DataGridViewRow row, string[] names)
        {
            if (row == null) return "";
            DataRowView rowView = row.DataBoundItem as DataRowView;
            if (rowView != null)
            {
                foreach (string name in names)
                {
                    if (rowView.DataView.Table.Columns.Contains(name))
                    {
                        object value = rowView[name];
                        if (value != null && value != DBNull.Value) return Convert.ToString(value, CultureInfo.CurrentCulture).Trim();
                    }
                }
            }
            if (row.DataGridView != null)
            {
                foreach (DataGridViewColumn column in row.DataGridView.Columns)
                {
                    if (ColumnMatches(column, names))
                    {
                        object value = row.Cells[column.Index].Value;
                        if (value != null) return Convert.ToString(value, CultureInfo.CurrentCulture).Trim();
                    }
                }
            }
            return "";
        }

        private static T GetField<T>(object target, string name) where T : class
        {
            if (target == null) return null;
            FieldInfo field = target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            return field == null ? null : field.GetValue(target) as T;
        }

        private static string[] QuotaNameColumns()
        {
            return new[] { "\u5de5\u7a0b\u6216\u8d39\u7528\u9879\u76ee\u540d\u79f0", "\u540d\u79f0", "\u9879\u76ee\u540d\u79f0" };
        }

        private static string[] QuotaCodeColumns()
        {
            return new[] { "\u5b9a\u989d\u7f16\u53f7", "\u5b9a\u989d\u7f16\u53f7DE", "\u7f16\u53f7" };
        }

        private static string[] QuotaUnitColumns() { return new[] { "\u5355\u4f4d" }; }
        private static string[] QuantityColumns() { return new[] { "\u5de5\u7a0b\u6570\u91cf\u8f93\u5165", "\u5de5\u7a0b\u6570\u91cf" }; }
        private static string[] UnitPriceColumns() { return new[] { "\u5355\u4ef7", "\u57fa\u671f\u4ef7\u683c" }; }
        private static string[] SingleWeightColumns() { return new[] { "\u5355\u91cd(t)", "\u5355\u91cd" }; }
        private static string[] CompilerColumns() { return new[] { "\u7f16\u5236\u4eba" }; }
        private static string[] ModifiedDateColumns() { return new[] { "\u4fee\u6539\u65e5\u671f" }; }
    }
}
