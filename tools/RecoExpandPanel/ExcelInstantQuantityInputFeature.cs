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
            private int targetRowIndex = -1;
            private int targetColumnIndex = -1;
            private string lastExcelKey = "";
            private bool wasSpreadsheetForeground;
            private bool waitingForSpreadsheetClick;
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
                waitingForSpreadsheetClick = false;
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
                targetRowIndex = -1;
                targetColumnIndex = -1;
                waitingForSpreadsheetClick = false;
                CaptureCurrentQuantityTarget(false);
            }

            private void ToggleEnabled()
            {
                enabled = !enabled;
                wasSpreadsheetForeground = false;
                waitingForSpreadsheetClick = false;
                lastExcelKey = TryReadCurrentSpreadsheetKey();
                if (enabled)
                {
                    CaptureCurrentQuantityTarget(false);
                    pollTimer.Start();
                    ShowStatus("\u5df2\u5f00\u542fExcel\u70b9\u9009\u5373\u586b\uff1a\u5148\u70b9\u8f6f\u4ef6\u5de5\u7a0b\u6570\u91cf\u683c\uff0c\u518d\u70b9Excel\u6570\u91cf\u683c\u3002");
                }
                else
                {
                    pollTimer.Stop();
                    ShowStatus("\u5df2\u5173\u95edExcel\u70b9\u9009\u5373\u586b\u3002");
                }
            }

            private void CaptureCurrentQuantityTarget(bool notify)
            {
                if (grid == null || grid.CurrentCell == null)
                {
                    return;
                }

                DataGridViewCell cell = grid.CurrentCell;
                if (!IsQuantityColumn(cell.ColumnIndex))
                {
                    return;
                }

                bool changed = targetRowIndex != cell.RowIndex || targetColumnIndex != cell.ColumnIndex;
                targetRowIndex = cell.RowIndex;
                targetColumnIndex = cell.ColumnIndex;
                if (changed)
                {
                    wasSpreadsheetForeground = false;
                    lastExcelKey = TryReadCurrentSpreadsheetKey();
                    waitingForSpreadsheetClick = enabled;
                    if (notify)
                    {
                        ShowStatus("\u5df2\u9009\u5b9a\u8f6f\u4ef6\u6570\u91cf\u683c\uff0c\u8bf7\u70b9\u51fbExcel\u5de5\u7a0b\u6570\u91cf\u5355\u5143\u683c\u3002");
                    }
                }
                else if (enabled && !waitingForSpreadsheetClick)
                {
                    wasSpreadsheetForeground = false;
                    lastExcelKey = TryReadCurrentSpreadsheetKey();
                    waitingForSpreadsheetClick = true;
                }
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

                    InstantExcelCell cell;
                    string error;
                    if (!TryReadActiveSpreadsheetCell(out cell, out error))
                    {
                        wasSpreadsheetForeground = false;
                        return;
                    }

                    bool selectionChanged = !String.Equals(lastExcelKey, cell.Key, StringComparison.OrdinalIgnoreCase);
                    if (!cell.IsForeground && !selectionChanged)
                    {
                        wasSpreadsheetForeground = false;
                        return;
                    }

                    bool foregroundEntered = cell.IsForeground && !wasSpreadsheetForeground;
                    wasSpreadsheetForeground = cell.IsForeground;
                    if (!foregroundEntered && String.Equals(lastExcelKey, cell.Key, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }

                    lastExcelKey = cell.Key;
                    if (ApplySpreadsheetCell(cell))
                    {
                        waitingForSpreadsheetClick = false;
                    }
                }
                catch (Exception ex)
                {
                    Log("Excel instant quantity input poll failed: " + ex);
                }
            }

            private bool HasValidTarget()
            {
                return grid != null &&
                    targetRowIndex >= 0 &&
                    targetRowIndex < grid.Rows.Count &&
                    targetColumnIndex >= 0 &&
                    targetColumnIndex < grid.Columns.Count &&
                    IsQuantityColumn(targetColumnIndex);
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

                DataGridViewRow row = grid.Rows[targetRowIndex];
                if (row == null || row.IsNewRow)
                {
                    return false;
                }

                string quotaUnit = ReadSoftwareUnit(row);
                decimal converted = quantity;
                if (!TryConvertQuantity(quantity, excelCell.UnitText, quotaUnit, out converted))
                {
                    ShowStatus("WPS\u5355\u4f4d " + CleanDisplayUnit(excelCell.UnitText) + " \u4e0e\u5b9a\u989d\u5355\u4f4d " + CleanDisplayUnit(quotaUnit) + " \u4e0d\u5339\u914d\uff0c\u672a\u5199\u5165\u3002");
                    Log("Excel instant quantity blocked unit mismatch: row=" + targetRowIndex.ToString(CultureInfo.InvariantCulture)
                        + " col=" + targetColumnIndex.ToString(CultureInfo.InvariantCulture)
                        + " excel=" + excelCell.Key
                        + " value=" + excelCell.ValueText
                        + " excelUnit=" + excelCell.UnitText
                        + " quotaUnit=" + quotaUnit);
                    return false;
                }

                string valueText = FormatDecimal(converted);

                WriteQuantity(row, targetColumnIndex, valueText, converted);
                string status = "\u5df2\u5199\u5165\uff1a" + valueText + " (" + CleanDisplayUnit(excelCell.UnitText) + " -> " + CleanDisplayUnit(quotaUnit) + ")";
                ShowStatus(status);
                Log("Excel instant quantity applied: row=" + targetRowIndex.ToString(CultureInfo.InvariantCulture)
                    + " col=" + targetColumnIndex.ToString(CultureInfo.InvariantCulture)
                    + " excel=" + excelCell.Key
                    + " value=" + excelCell.ValueText
                    + " excelUnit=" + excelCell.UnitText
                    + " quotaUnit=" + quotaUnit
                    + " result=" + valueText
                    + " converted=True");
                return true;
            }

            private string ReadSoftwareUnit(DataGridViewRow row)
            {
                string unit = GetRowValue(row, "\u5355\u4f4d", "\u5b9a\u989d\u5355\u4f4d");
                if (!String.IsNullOrWhiteSpace(unit))
                {
                    return unit;
                }

                for (int col = targetColumnIndex - 1; col >= 0 && col >= targetColumnIndex - 5; col--)
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

            private void WriteQuantity(DataGridViewRow row, int visibleColumnIndex, string valueText, decimal numericValue)
            {
                DataRowView rowView = row.DataBoundItem as DataRowView;
                DataRow dataRow = rowView == null ? null : rowView.Row;
                bool wroteInput = false;

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
                        manager.Refresh();
                    }

                    grid.InvalidateCell(visibleColumnIndex, row.Index);
                }
                catch (Exception ex)
                {
                    Log("Excel instant quantity commit failed: " + ex.Message);
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
                Log("Excel instant quantity status: " + message);
                if (grid == null || grid.IsDisposed)
                {
                    return;
                }

                DateTime now = DateTime.UtcNow;
                if ((now - lastStatusUtc).TotalMilliseconds < 600)
                {
                    return;
                }

                lastStatusUtc = now;
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

            private struct InstantUnitScale
            {
                public string BaseUnit;
                public decimal Scale;
            }
        }
    }
}
