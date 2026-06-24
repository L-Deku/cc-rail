using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        private static readonly Dictionary<Form, ExcelInstantQuantityInputRuntime> ExcelInstantQuantityInputRuntimes =
            new Dictionary<Form, ExcelInstantQuantityInputRuntime>();

        private static void EnsureExcelInstantQuantityInputRuntime(Form mainForm)
        {
            if (mainForm == null || ExcelInstantQuantityInputRuntimes.ContainsKey(mainForm))
            {
                return;
            }

            ExcelInstantQuantityInputRuntime runtime = new ExcelInstantQuantityInputRuntime(mainForm);
            if (runtime.Install())
            {
                ExcelInstantQuantityInputRuntimes[mainForm] = runtime;
                mainForm.FormClosed += delegate
                {
                    runtime.Dispose();
                    ExcelInstantQuantityInputRuntimes.Remove(mainForm);
                };
            }
        }

        private sealed class ExcelInstantQuantityInputRuntime : IDisposable
        {
            private const int PollIntervalMs = 100;
            private const int ReconnectDelayMs = 2000;
            private readonly Form mainForm;
            private readonly Timer pollTimer;
            private readonly ToolTip statusTip;
            private DataGridView grid;
            private object spreadsheetApplication;
            private bool enabled;
            private readonly List<InstantQuantityTarget> quantityTargets = new List<InstantQuantityTarget>();
            private string lastExcelKey = "";
            private bool wasSpreadsheetForeground;
            private bool waitingForSpreadsheetClick;
            private bool awaitingSpreadsheetActivation;
            private bool waitingForSoftwareBlur;
            private bool hasReusableSpreadsheetSelection;
            private bool applyingQuantity;
            private string lastStatusMessage = "";
            private DateTime lastStatusUtc = DateTime.MinValue;
            private DateTime nextConnectionAttemptUtc = DateTime.MinValue;

            public ExcelInstantQuantityInputRuntime(Form mainForm)
            {
                this.mainForm = mainForm;
                pollTimer = new Timer();
                pollTimer.Interval = PollIntervalMs;
                pollTimer.Tick += delegate { PollActiveSpreadsheetCell(); };
                statusTip = new ToolTip();
                statusTip.ShowAlways = true;
            }

            public bool Install()
            {
                grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
                if (grid == null)
                {
                    Log("Excel instant quantity input skipped: dataGridViewDE not found.");
                    return false;
                }

                HookGrid();
                CaptureCurrentQuantityTarget(false);
                Log("Excel instant quantity input installed. Toggle=Ctrl+Shift+Q.");
                return true;
            }

            public void Dispose()
            {
                pollTimer.Stop();
                pollTimer.Dispose();
                statusTip.Dispose();
                spreadsheetApplication = null;
                ClearQuantityTargets();
                if (grid != null)
                {
                    grid.KeyDown -= GridKeyDown;
                    grid.CellEnter -= GridCellEnter;
                    grid.CurrentCellChanged -= GridCurrentCellChanged;
                    grid.DataSourceChanged -= GridDataSourceChanged;
                }
            }

            private void HookGrid()
            {
                grid.KeyDown -= GridKeyDown;
                grid.KeyDown += GridKeyDown;
                grid.CellEnter -= GridCellEnter;
                grid.CellEnter += GridCellEnter;
                grid.CurrentCellChanged -= GridCurrentCellChanged;
                grid.CurrentCellChanged += GridCurrentCellChanged;
                grid.DataSourceChanged -= GridDataSourceChanged;
                grid.DataSourceChanged += GridDataSourceChanged;
            }

            private void GridKeyDown(object sender, KeyEventArgs e)
            {
                if (e.Control && e.Shift && e.KeyCode == Keys.Q)
                {
                    ToggleEnabled();
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            }

            private void GridCellEnter(object sender, DataGridViewCellEventArgs e)
            {
                CaptureCurrentQuantityTarget(enabled);
            }

            private void GridCurrentCellChanged(object sender, EventArgs e)
            {
                CaptureCurrentQuantityTarget(enabled);
            }

            private void GridDataSourceChanged(object sender, EventArgs e)
            {
                ClearQuantityTargets();
                CaptureCurrentQuantityTarget(false);
            }

            private void ToggleEnabled()
            {
                enabled = !enabled;
                wasSpreadsheetForeground = false;
                ClearQuantityTargets();
                if (enabled)
                {
                    lastExcelKey = TryReadCurrentSpreadsheetKey();
                    CaptureCurrentQuantityTarget(false);
                    pollTimer.Start();
                    ShowStatus("\u5df2\u5f00\u542fExcel\u70b9\u9009\u5373\u586b\uff1a\u5148\u70b9\u8f6f\u4ef6\u5de5\u7a0b\u6570\u91cf\u683c\uff0c\u518d\u70b9Excel\u6570\u91cf\u683c\u3002");
                }
                else
                {
                    lastExcelKey = "";
                    hasReusableSpreadsheetSelection = false;
                    pollTimer.Stop();
                    ShowStatus("\u5df2\u5173\u95edExcel\u70b9\u9009\u5373\u586b\u3002");
                }
            }

            private void CaptureCurrentQuantityTarget(bool notify)
            {
                if (applyingQuantity)
                {
                    return;
                }

                if (grid == null || grid.CurrentCell == null)
                {
                    return;
                }

                DataGridViewCell cell = grid.CurrentCell;
                if (!IsQuantityColumn(cell.ColumnIndex))
                {
                    ClearQuantityTargets();
                    return;
                }

                List<InstantQuantityTarget> selectedTargets = CollectSelectedQuantityTargets(cell);
                bool changed = !SameTargets(quantityTargets, selectedTargets);
                quantityTargets.Clear();
                quantityTargets.AddRange(selectedTargets);
                if (changed)
                {
                    wasSpreadsheetForeground = false;
                    if (enabled)
                    {
                        lastExcelKey = TryReadCurrentSpreadsheetKey();
                        waitingForSpreadsheetClick = true;
                        awaitingSpreadsheetActivation = true;
                        waitingForSoftwareBlur = !hasReusableSpreadsheetSelection;
                    }
                    else
                    {
                        lastExcelKey = "";
                        waitingForSpreadsheetClick = false;
                        awaitingSpreadsheetActivation = false;
                        waitingForSoftwareBlur = false;
                    }

                    if (notify)
                    {
                        ShowStatus(BuildTargetStatusMessage());
                    }
                }
                else if (enabled && !waitingForSpreadsheetClick)
                {
                    wasSpreadsheetForeground = false;
                    lastExcelKey = TryReadCurrentSpreadsheetKey();
                    waitingForSpreadsheetClick = true;
                    awaitingSpreadsheetActivation = true;
                    waitingForSoftwareBlur = !hasReusableSpreadsheetSelection;
                }
            }

            private void ClearQuantityTargets()
            {
                quantityTargets.Clear();
                waitingForSpreadsheetClick = false;
                awaitingSpreadsheetActivation = false;
                waitingForSoftwareBlur = false;
                wasSpreadsheetForeground = false;
                lastExcelKey = "";
            }

            private List<InstantQuantityTarget> CollectSelectedQuantityTargets(DataGridViewCell currentCell)
            {
                List<InstantQuantityTarget> targets = new List<InstantQuantityTarget>();
                if (grid == null || currentCell == null)
                {
                    return targets;
                }

                AddQuantityTarget(targets, currentCell.RowIndex, currentCell.ColumnIndex);

                foreach (DataGridViewCell selectedCell in grid.SelectedCells)
                {
                    AddQuantityTarget(targets, selectedCell.RowIndex, selectedCell.ColumnIndex);
                }

                if (IsQuantityColumn(currentCell.ColumnIndex))
                {
                    foreach (DataGridViewRow selectedRow in grid.SelectedRows)
                    {
                        AddQuantityTarget(targets, selectedRow.Index, currentCell.ColumnIndex);
                    }
                }

                targets.Sort(delegate(InstantQuantityTarget left, InstantQuantityTarget right)
                {
                    int rowCompare = left.RowIndex.CompareTo(right.RowIndex);
                    return rowCompare != 0 ? rowCompare : left.ColumnIndex.CompareTo(right.ColumnIndex);
                });
                return targets;
            }

            private void AddQuantityTarget(List<InstantQuantityTarget> targets, int rowIndex, int columnIndex)
            {
                if (grid == null ||
                    rowIndex < 0 ||
                    rowIndex >= grid.Rows.Count ||
                    columnIndex < 0 ||
                    columnIndex >= grid.Columns.Count ||
                    !IsQuantityColumn(columnIndex))
                {
                    return;
                }

                DataGridViewRow row = grid.Rows[rowIndex];
                if (row == null)
                {
                    return;
                }

                foreach (InstantQuantityTarget target in targets)
                {
                    if (target.RowIndex == rowIndex && target.ColumnIndex == columnIndex)
                    {
                        return;
                    }
                }

                targets.Add(new InstantQuantityTarget(rowIndex, columnIndex));
            }

            private static bool SameTargets(List<InstantQuantityTarget> left, List<InstantQuantityTarget> right)
            {
                if (left == null || right == null || left.Count != right.Count)
                {
                    return false;
                }

                for (int i = 0; i < left.Count; i++)
                {
                    if (left[i].RowIndex != right[i].RowIndex || left[i].ColumnIndex != right[i].ColumnIndex)
                    {
                        return false;
                    }
                }

                return true;
            }

            private bool IsQuantityColumn(int columnIndex)
            {
                return IsGridColumn(columnIndex, "\u5de5\u7a0b\u6570\u91cf\u8f93\u5165") ||
                    IsGridColumn(columnIndex, "\u5de5\u7a0b\u6570\u91cf");
            }

            private bool IsGridColumn(int columnIndex, string expected)
            {
                if (grid == null || columnIndex < 0 || columnIndex >= grid.Columns.Count)
                {
                    return false;
                }

                DataGridViewColumn column = grid.Columns[columnIndex];
                return String.Equals(column.DataPropertyName, expected, StringComparison.Ordinal) ||
                    String.Equals(column.Name, expected, StringComparison.Ordinal) ||
                    String.Equals(column.HeaderText, expected, StringComparison.Ordinal);
            }

            private void PollActiveSpreadsheetCell()
            {
                if (!enabled)
                {
                    return;
                }

                try
                {
                    if (!HasValidTarget())
                    {
                        waitingForSpreadsheetClick = false;
                        CaptureCurrentQuantityTarget(false);
                        return;
                    }

                    if (!waitingForSpreadsheetClick)
                    {
                        return;
                    }

                    if (hasReusableSpreadsheetSelection &&
                        !awaitingSpreadsheetActivation &&
                        IsMainFormForeground())
                    {
                        wasSpreadsheetForeground = false;
                        return;
                    }

                    if (waitingForSoftwareBlur)
                    {
                        if (IsMainFormForeground())
                        {
                            return;
                        }

                        waitingForSoftwareBlur = false;
                    }

                    InstantExcelCell cell;
                    string error;
                    if (!TryReadActiveSpreadsheetCell(out cell, out error))
                    {
                        wasSpreadsheetForeground = false;
                        return;
                    }

                    bool selectionChanged = !String.Equals(lastExcelKey, cell.Key, StringComparison.OrdinalIgnoreCase);
                    bool activationClick = awaitingSpreadsheetActivation && hasReusableSpreadsheetSelection;
                    if (!cell.IsForeground && !selectionChanged && !activationClick)
                    {
                        wasSpreadsheetForeground = false;
                        return;
                    }

                    bool foregroundEntered = cell.IsForeground && !wasSpreadsheetForeground;
                    wasSpreadsheetForeground = cell.IsForeground;
                    if (!foregroundEntered &&
                        !activationClick &&
                        String.Equals(lastExcelKey, cell.Key, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }

                    lastExcelKey = cell.Key;
                    awaitingSpreadsheetActivation = false;
                    if (ApplySpreadsheetCell(cell))
                    {
                        MarkSpreadsheetCellApplied();
                    }
                }
                catch (Exception ex)
                {
                    Log("Excel instant quantity input poll failed: " + ex);
                }
            }

            private void MarkSpreadsheetCellApplied()
            {
                hasReusableSpreadsheetSelection = true;
                waitingForSpreadsheetClick = true;
            }

            private bool HasValidTarget()
            {
                if (grid == null || quantityTargets.Count == 0)
                {
                    return false;
                }

                foreach (InstantQuantityTarget target in quantityTargets)
                {
                    if (target.RowIndex >= 0 &&
                        target.RowIndex < grid.Rows.Count &&
                        target.ColumnIndex >= 0 &&
                        target.ColumnIndex < grid.Columns.Count &&
                        IsQuantityColumn(target.ColumnIndex))
                    {
                        return true;
                    }
                }

                return false;
            }

            private bool ApplySpreadsheetCell(InstantExcelCell excelCell)
            {
                decimal quantity;
                if (!TryEvaluateQuantity(excelCell.ValueText, out quantity))
                {
                    ShowStatus("\u5df2\u8df3\u8fc7\uff1aExcel\u5f53\u524d\u5355\u5143\u683c\u4e0d\u662f\u6709\u6548\u6570\u503c\u3002");
                    Log("Excel instant quantity skipped nonnumeric: " + excelCell.Key + " value=" + excelCell.ValueText);
                    return false;
                }

                List<InstantQuantityTarget> targets = new List<InstantQuantityTarget>(quantityTargets);
                int applied = 0;
                int skipped = 0;
                foreach (InstantQuantityTarget target in targets)
                {
                    if (target.RowIndex < 0 || target.RowIndex >= grid.Rows.Count ||
                        target.ColumnIndex < 0 || target.ColumnIndex >= grid.Columns.Count ||
                        !IsQuantityColumn(target.ColumnIndex))
                    {
                        skipped++;
                        continue;
                    }

                    DataGridViewRow row = grid.Rows[target.RowIndex];
                    if (row == null)
                    {
                        skipped++;
                        continue;
                    }

                    string quotaUnit = ReadSoftwareUnit(row, target.ColumnIndex);
                    decimal converted = quantity;
                    bool blankQuotaWithoutUnit = String.IsNullOrWhiteSpace(quotaUnit) && IsBlankQuotaRow(row);
                    if (!blankQuotaWithoutUnit && !TryConvertQuantity(quantity, excelCell.UnitText, quotaUnit, out converted))
                    {
                        skipped++;
                        Log("Excel instant quantity blocked unit mismatch: row=" + target.RowIndex.ToString(CultureInfo.InvariantCulture)
                            + " col=" + target.ColumnIndex.ToString(CultureInfo.InvariantCulture)
                            + " excel=" + excelCell.Key
                            + " value=" + excelCell.ValueText
                            + " excelUnit=" + excelCell.UnitText
                            + " quotaUnit=" + quotaUnit);
                        continue;
                    }

                    string valueText = FormatDecimal(converted);
                    decimal writeValue;
                    if (!TryEvaluateQuantity(valueText, out writeValue))
                    {
                        writeValue = converted;
                    }

                    WriteQuantity(row, target.ColumnIndex, valueText, writeValue);
                    applied++;
                    Log("Excel instant quantity applied: row=" + target.RowIndex.ToString(CultureInfo.InvariantCulture)
                        + " col=" + target.ColumnIndex.ToString(CultureInfo.InvariantCulture)
                        + " excel=" + excelCell.Key
                        + " value=" + excelCell.ValueText
                        + " excelUnit=" + excelCell.UnitText
                        + " quotaUnit=" + quotaUnit
                        + " result=" + valueText
                        + " converted=True");
                }

                if (applied > 0)
                {
                    ShowStatus("\u5df2\u5199\u5165" + applied.ToString(CultureInfo.InvariantCulture) + "\u884c" +
                        (skipped > 0 ? "\uff0c\u8df3\u8fc7" + skipped.ToString(CultureInfo.InvariantCulture) + "\u884c" : ""));
                    return true;
                }

                ShowStatus("WPS\u5355\u4f4d " + CleanDisplayUnit(excelCell.UnitText) + " \u4e0e\u5df2\u9009\u5b9a\u5b9a\u989d\u5355\u4f4d\u4e0d\u5339\u914d\uff0c\u672a\u5199\u5165\u3002");
                return false;
            }

            private string ReadSoftwareUnit(DataGridViewRow row, int quantityColumnIndex)
            {
                string unit = GetRowValue(row, "\u5355\u4f4d", "\u5b9a\u989d\u5355\u4f4d");
                if (!String.IsNullOrWhiteSpace(unit))
                {
                    return unit;
                }

                for (int col = quantityColumnIndex - 1; col >= 0 && col >= quantityColumnIndex - 5; col--)
                {
                    object value = row.Cells[col].Value;
                    string text = value == null ? "" : Convert.ToString(value, CultureInfo.CurrentCulture);
                    if (LooksLikeInstantUnit(text))
                    {
                        return text;
                    }
                }

                return "";
            }

            private string BuildTargetStatusMessage()
            {
                string countText = quantityTargets.Count.ToString(CultureInfo.InvariantCulture);
                if (hasReusableSpreadsheetSelection)
                {
                    return "\u5df2\u9009\u5b9a" + countText + "\u4e2a\u8f6f\u4ef6\u6570\u91cf\u683c\uff0c\u5c06\u590d\u7528\u5f53\u524dExcel\u5de5\u7a0b\u6570\u91cf\u3002";
                }

                return "\u5df2\u9009\u5b9a" + countText + "\u4e2a\u8f6f\u4ef6\u6570\u91cf\u683c\uff0c\u8bf7\u70b9\u51fbExcel\u5de5\u7a0b\u6570\u91cf\u5355\u5143\u683c\u3002";
            }

            private static bool IsBlankQuotaRow(DataGridViewRow row)
            {
                if (row == null)
                {
                    return false;
                }

                string code = GetRowValue(row, "\u7f16\u53f7", "\u5b9a\u989d\u7f16\u53f7", "\u5b9a\u989d\u7f16\u53f7DE");
                string name = GetRowValue(row, "\u540d\u79f0", "\u5de5\u7a0b\u6216\u8d39\u7528\u9879\u76ee\u540d\u79f0", "\u9879\u76ee\u540d\u79f0");
                return String.IsNullOrWhiteSpace(code) && String.IsNullOrWhiteSpace(name);
            }

            private void WriteQuantity(DataGridViewRow row, int visibleColumnIndex, string valueText, decimal numericValue)
            {
                DataRowView rowView = row.DataBoundItem as DataRowView;
                DataRow dataRow = rowView == null ? null : rowView.Row;
                bool wroteInput = false;
                DataGridViewCell previousCell = grid.CurrentCell;
                int firstDisplayedRowIndex = -1;
                int horizontalOffset = 0;
                try
                {
                    firstDisplayedRowIndex = grid.FirstDisplayedScrollingRowIndex;
                    horizontalOffset = grid.HorizontalScrollingOffset;
                }
                catch
                {
                }

                if (dataRow != null && dataRow.Table != null)
                {
                    wroteInput = TrySetDataRowValue(dataRow, "\u5de5\u7a0b\u6570\u91cf\u8f93\u5165", valueText);
                    if (!wroteInput)
                    {
                        DataGridViewColumn visibleColumn = grid.Columns[visibleColumnIndex];
                        string boundName = String.IsNullOrEmpty(visibleColumn.DataPropertyName)
                            ? visibleColumn.Name
                            : visibleColumn.DataPropertyName;
                        wroteInput = TrySetDataRowValue(dataRow, boundName, valueText);
                    }

                    TrySetDataRowValue(dataRow, "\u5de5\u7a0b\u6570\u91cf", numericValue);
                }

                if (!wroteInput)
                {
                    row.Cells[visibleColumnIndex].Value = valueText;
                }

                try
                {
                    applyingQuantity = true;
                    grid.CurrentCell = row.Cells[visibleColumnIndex];
                    if (grid.IsCurrentCellDirty)
                    {
                        grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
                    }

                    grid.EndEdit();
                    BindingSource bindingSource = grid.DataSource as BindingSource;
                    if (bindingSource != null)
                    {
                        bindingSource.EndEdit();
                    }

                    CurrencyManager manager = grid.BindingContext == null
                        ? null
                        : grid.BindingContext[grid.DataSource] as CurrencyManager;
                    if (manager != null)
                    {
                        manager.EndCurrentEdit();
                    }

                    grid.InvalidateCell(visibleColumnIndex, row.Index);
                    RestoreGridViewPosition(previousCell, firstDisplayedRowIndex, horizontalOffset);
                }
                catch (Exception ex)
                {
                    Log("Excel instant quantity commit failed: " + ex.Message);
                }
                finally
                {
                    applyingQuantity = false;
                }
            }

            private void RestoreGridViewPosition(DataGridViewCell previousCell, int firstDisplayedRowIndex, int horizontalOffset)
            {
                if (grid == null || grid.IsDisposed)
                {
                    return;
                }

                try
                {
                    if (previousCell != null &&
                        previousCell.RowIndex >= 0 &&
                        previousCell.RowIndex < grid.Rows.Count &&
                        previousCell.ColumnIndex >= 0 &&
                        previousCell.ColumnIndex < grid.Columns.Count)
                    {
                        grid.CurrentCell = grid.Rows[previousCell.RowIndex].Cells[previousCell.ColumnIndex];
                    }

                    if (firstDisplayedRowIndex >= 0 &&
                        firstDisplayedRowIndex < grid.Rows.Count &&
                        !grid.Rows[firstDisplayedRowIndex].IsNewRow)
                    {
                        grid.FirstDisplayedScrollingRowIndex = firstDisplayedRowIndex;
                    }

                    if (horizontalOffset >= 0)
                    {
                        grid.HorizontalScrollingOffset = horizontalOffset;
                    }
                }
                catch (Exception ex)
                {
                    Log("Excel instant quantity restore view failed: " + ex.Message);
                }
            }

            private bool TrySetDataRowValue(DataRow row, string columnName, object value)
            {
                if (row == null || row.Table == null || String.IsNullOrEmpty(columnName) || !row.Table.Columns.Contains(columnName))
                {
                    return false;
                }

                DataColumn column = row.Table.Columns[columnName];
                try
                {
                    if (value == null)
                    {
                        row[column] = DBNull.Value;
                    }
                    else if (column.DataType == typeof(string))
                    {
                        row[column] = Convert.ToString(value, CultureInfo.CurrentCulture);
                    }
                    else
                    {
                        Type targetType = Nullable.GetUnderlyingType(column.DataType) ?? column.DataType;
                        row[column] = Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture);
                    }

                    return true;
                }
                catch (Exception ex)
                {
                    Log("Excel instant quantity set row value failed: column=" + columnName + " " + ex.Message);
                    return false;
                }
            }

            private string TryReadCurrentSpreadsheetKey()
            {
                InstantExcelCell cell;
                string error;
                return TryReadActiveSpreadsheetCell(out cell, out error) ? cell.Key : "";
            }

            private bool TryReadActiveSpreadsheetCell(out InstantExcelCell cell, out string error)
            {
                cell = null;
                error = null;
                dynamic spreadsheet = null;
                try
                {
                    spreadsheet = GetCachedSpreadsheetApplication();
                    if (spreadsheet == null)
                    {
                        error = "No active spreadsheet application.";
                        return false;
                    }

                    dynamic workbook = spreadsheet.ActiveWorkbook;
                    dynamic sheet = spreadsheet.ActiveSheet;
                    dynamic selection = spreadsheet.Selection;
                    if (workbook == null || sheet == null || selection == null)
                    {
                        error = "No active workbook, sheet, or selection.";
                        return false;
                    }

                    dynamic firstCell = selection.Cells[1, 1];
                    int row = Convert.ToInt32(firstCell.Row, CultureInfo.InvariantCulture);
                    int column = Convert.ToInt32(firstCell.Column, CultureInfo.InvariantCulture);
                    string address = Convert.ToString(firstCell.Address(false, false), CultureInfo.InvariantCulture);
                    string workbookPath = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture);
                    string worksheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture);
                    string valueText = InstantExcelValueToText(firstCell.Value2);
                    string unitText = "";
                    if (column > 1)
                    {
                        dynamic unitCell = sheet.Cells[row, column - 1];
                        unitText = InstantExcelValueToText(unitCell.Value2);
                    }

                    cell = new InstantExcelCell();
                    cell.WorkbookPath = workbookPath;
                    cell.WorksheetName = worksheetName;
                    cell.Address = address;
                    cell.ValueText = valueText;
                    cell.UnitText = unitText;
                    cell.IsForeground = IsSpreadsheetForeground(spreadsheet);
                    cell.Key = workbookPath + "|" + worksheetName + "|" + address + "|" + valueText + "|" + unitText;
                    return true;
                }
                catch (COMException ex)
                {
                    ResetSpreadsheetApplication(ex.Message);
                    error = ex.Message;
                    return false;
                }
                catch (Exception ex)
                {
                    ResetSpreadsheetApplication(ex.Message);
                    error = ex.Message;
                    return false;
                }
            }

            private object GetCachedSpreadsheetApplication()
            {
                if (spreadsheetApplication != null)
                {
                    return spreadsheetApplication;
                }

                DateTime now = DateTime.UtcNow;
                if (now < nextConnectionAttemptUtc)
                {
                    return null;
                }

                nextConnectionAttemptUtc = now.AddMilliseconds(ReconnectDelayMs);
                spreadsheetApplication = GetActiveSpreadsheetApplication();
                return spreadsheetApplication;
            }

            private void ResetSpreadsheetApplication(string reason)
            {
                spreadsheetApplication = null;
                nextConnectionAttemptUtc = DateTime.UtcNow.AddMilliseconds(ReconnectDelayMs);
                if (!String.IsNullOrWhiteSpace(reason))
                {
                    Log("Excel instant quantity input disconnected: " + reason);
                }
            }

            private static bool IsSpreadsheetForeground(dynamic spreadsheet)
            {
                try
                {
                    IntPtr foreground = InstantGetForegroundWindow();
                    if (foreground == IntPtr.Zero)
                    {
                        return false;
                    }

                    if (SpreadsheetHwndMatchesForegroundProcess(spreadsheet, foreground))
                    {
                        return true;
                    }

                    return WindowContainsExcelGrid(foreground);
                }
                catch
                {
                    return false;
                }
            }

            private bool IsMainFormForeground()
            {
                try
                {
                    if (mainForm == null || mainForm.IsDisposed || !mainForm.IsHandleCreated)
                    {
                        return false;
                    }

                    IntPtr foreground = InstantGetForegroundWindow();
                    if (foreground == IntPtr.Zero)
                    {
                        return false;
                    }

                    IntPtr mainHandle = mainForm.Handle;
                    return foreground == mainHandle || InstantIsChild(mainHandle, foreground);
                }
                catch
                {
                    return false;
                }
            }

            private static bool SpreadsheetHwndMatchesForegroundProcess(dynamic spreadsheet, IntPtr foreground)
            {
                try
                {
                    uint foregroundPid;
                    InstantGetWindowThreadProcessId(foreground, out foregroundPid);
                    IntPtr spreadsheetHandle = new IntPtr(Convert.ToInt64(spreadsheet.Hwnd, CultureInfo.InvariantCulture));
                    uint spreadsheetPid;
                    InstantGetWindowThreadProcessId(spreadsheetHandle, out spreadsheetPid);
                    return foregroundPid != 0 && foregroundPid == spreadsheetPid;
                }
                catch
                {
                    return false;
                }
            }

            private static bool WindowContainsExcelGrid(IntPtr windowHandle)
            {
                if (windowHandle == IntPtr.Zero)
                {
                    return false;
                }

                try
                {
                    if (String.Equals(GetWindowClassName(windowHandle), "EXCEL7", StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }

                    List<IntPtr> excelWindows = new List<IntPtr>();
                    CollectExcelChildWindows(windowHandle, excelWindows);
                    return excelWindows.Count > 0;
                }
                catch
                {
                    return false;
                }
            }

            private static string InstantExcelValueToText(object value)
            {
                if (value == null || value == DBNull.Value)
                {
                    return "";
                }

                if (value is double || value is float)
                {
                    return Convert.ToDouble(value, CultureInfo.InvariantCulture).ToString("G15", CultureInfo.InvariantCulture);
                }

                if (value is decimal)
                {
                    return ((decimal)value).ToString(CultureInfo.InvariantCulture);
                }

                if (value is int || value is long || value is short || value is byte)
                {
                    return Convert.ToString(value, CultureInfo.InvariantCulture);
                }

                return Convert.ToString(value, CultureInfo.CurrentCulture).Trim();
            }

            private static bool TryEvaluateQuantity(string text, out decimal value)
            {
                value = 0m;
                string normalized = (text ?? "").Trim();
                if (normalized.StartsWith("=", StringComparison.Ordinal))
                {
                    normalized = normalized.Substring(1);
                }

                return Decimal.TryParse(normalized, NumberStyles.Float, CultureInfo.InvariantCulture, out value) ||
                    Decimal.TryParse(normalized, NumberStyles.Float, CultureInfo.CurrentCulture, out value);
            }

            private static bool TryConvertQuantity(decimal quantity, string excelUnitText, string quotaUnitText, out decimal converted)
            {
                converted = quantity;
                InstantUnitScale excelUnit = ParseInstantUnitScale(excelUnitText);
                InstantUnitScale quotaUnit = ParseInstantUnitScale(quotaUnitText);
                if (String.IsNullOrEmpty(excelUnit.BaseUnit) ||
                    String.IsNullOrEmpty(quotaUnit.BaseUnit) ||
                    excelUnit.Scale <= 0m ||
                    quotaUnit.Scale <= 0m ||
                    !String.Equals(excelUnit.BaseUnit, quotaUnit.BaseUnit, StringComparison.Ordinal))
                {
                    return false;
                }

                converted = quantity * excelUnit.Scale / quotaUnit.Scale;
                return true;
            }

            private static InstantUnitScale ParseInstantUnitScale(string text)
            {
                string unit = NormalizeInstantUnit(text);
                if (String.IsNullOrEmpty(unit))
                {
                    return new InstantUnitScale();
                }

                Match match = Regex.Match(unit, @"^(?<scale>\d+(?:\.\d+)?)(?<unit>.+)$");
                decimal scale = 1m;
                string baseUnit = unit;
                if (match.Success)
                {
                    Decimal.TryParse(match.Groups["scale"].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out scale);
                    baseUnit = match.Groups["unit"].Value;
                }

                baseUnit = NormalizeInstantBaseUnit(baseUnit);
                if (String.IsNullOrEmpty(baseUnit))
                {
                    return new InstantUnitScale();
                }

                InstantUnitScale result = new InstantUnitScale();
                result.Scale = scale <= 0m ? 1m : scale;
                result.BaseUnit = baseUnit;
                return result;
            }

            private static string NormalizeInstantUnit(string text)
            {
                string unit = (text ?? "").Trim().ToLowerInvariant();
                unit = unit.Replace(" ", "").Replace("\u3000", "");
                unit = unit.Replace("\uff10", "0").Replace("\uff11", "1").Replace("\uff12", "2")
                    .Replace("\uff13", "3").Replace("\uff14", "4").Replace("\uff15", "5")
                    .Replace("\uff16", "6").Replace("\uff17", "7").Replace("\uff18", "8")
                    .Replace("\uff19", "9");
                unit = unit.Replace("\u00b2", "2").Replace("\u00b3", "3");
                unit = unit.Replace("\u33a1", "m2").Replace("\u33a5", "m3");
                unit = unit.Replace("m^2", "m2").Replace("m\uff3e2", "m2");
                unit = unit.Replace("m^3", "m3").Replace("m\uff3e3", "m3");
                unit = unit.Replace("\u7acb\u65b9\u7c73", "m3");
                unit = unit.Replace("\u5e73\u65b9\u7c73", "m2");
                unit = unit.Replace("\u7c73", "m");
                return unit;
            }

            private static string NormalizeInstantBaseUnit(string unit)
            {
                unit = NormalizeInstantUnit(unit);
                if (unit == "m" || unit == "m2" || unit == "m3" || unit == "kg" || unit == "t")
                {
                    return unit;
                }

                return unit;
            }

            private static bool LooksLikeInstantUnit(string text)
            {
                InstantUnitScale unit = ParseInstantUnitScale(text);
                return !String.IsNullOrEmpty(unit.BaseUnit);
            }

            private static string FormatDecimal(decimal value)
            {
                decimal rounded = Decimal.Round(value, 2, MidpointRounding.AwayFromZero);
                return value == Decimal.Truncate(value)
                    ? rounded.ToString("0", CultureInfo.InvariantCulture)
                    : rounded.ToString("0.00", CultureInfo.InvariantCulture);
            }

            private static string CleanDisplayUnit(string text)
            {
                return String.IsNullOrWhiteSpace(text) ? "\u7a7a\u5355\u4f4d" : text.Trim();
            }

            private void ShowStatus(string message)
            {
                DateTime now = DateTime.UtcNow;
                if (String.Equals(lastStatusMessage, message, StringComparison.Ordinal) &&
                    (now - lastStatusUtc).TotalMilliseconds < 600)
                {
                    return;
                }

                lastStatusMessage = message;
                lastStatusUtc = now;
                Log("Excel instant quantity status: " + message);
                if (grid == null || grid.IsDisposed)
                {
                    return;
                }

                try
                {
                    statusTip.Show(message, grid, 12, 12, 1800);
                }
                catch
                {
                }
            }

            [DllImport("user32.dll")]
            private static extern IntPtr InstantGetForegroundWindow();

            [DllImport("user32.dll")]
            private static extern uint InstantGetWindowThreadProcessId(IntPtr hWnd, out uint processId);

            [DllImport("user32.dll", EntryPoint = "IsChild")]
            private static extern bool InstantIsChild(IntPtr hWndParent, IntPtr hWnd);

            private sealed class InstantExcelCell
            {
                public string WorkbookPath;
                public string WorksheetName;
                public string Address;
                public string ValueText;
                public string UnitText;
                public string Key;
                public bool IsForeground;
            }

            private struct InstantQuantityTarget
            {
                public readonly int RowIndex;
                public readonly int ColumnIndex;

                public InstantQuantityTarget(int rowIndex, int columnIndex)
                {
                    RowIndex = rowIndex;
                    ColumnIndex = columnIndex;
                }
            }

            private struct InstantUnitScale
            {
                public string BaseUnit;
                public decimal Scale;
            }
        }
    }
}
