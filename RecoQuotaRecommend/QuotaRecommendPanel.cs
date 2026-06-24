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
using Harmony = HarmonyLib.Harmony;
using HarmonyMethod = HarmonyLib.HarmonyMethod;

namespace RecoQuotaRecommend
{
    public sealed class QuotaRecommendPanel : Form
    {
        private static readonly HashSet<Form> InstalledForms = new HashSet<Form>();
        private static readonly Dictionary<ContextMenuStrip, MenuInfo> MenuInfos = new Dictionary<ContextMenuStrip, MenuInfo>();
        private static readonly Dictionary<Form, RecommendDialog> RecommendDialogs = new Dictionary<Form, RecommendDialog>();
        private static Image recommendMenuIcon;
        private static bool idleHooked;
        private static bool consumeCryptoBridgeStarted;
        private static System.Windows.Forms.Timer consumeCryptoBridgeTimer;
        private static string consumeCryptoBridgeLastRequest;
        private static System.Windows.Forms.Timer readOnlyCalcLearningTimer;
        private static System.Windows.Forms.Timer migratedQuotaLookupTimer;
        private static readonly Dictionary<string, string> ReadOnlyCalcLearningSignatures = new Dictionary<string, string>();
        private static readonly HashSet<string> ReadOnlyCalcLearningColumnLogs = new HashSet<string>();
        private static readonly HashSet<string> ReadOnlyCalcLearningDataSetSchemaLogs = new HashSet<string>();
        private static System.Windows.Forms.Timer inputCompatibilityTimer;
        private static readonly HashSet<Form> LoggedValidationDialogs = new HashSet<Form>();
        private static readonly HashSet<Form> DiagnosedInputForms = new HashSet<Form>();
        private static readonly HashSet<Form> DiagnosedDataSetForms = new HashSet<Form>();
        private static readonly HashSet<DataGridView> WrappedInputGrids = new HashSet<DataGridView>();
        private static readonly HashSet<DataGridView> MigratedQuotaCompatibilityGrids = new HashSet<DataGridView>();
        private static readonly Dictionary<DataGridView, Delegate> MigratedQuotaValidatingWrappers =
            new Dictionary<DataGridView, Delegate>();
        private static readonly Dictionary<DataGridView, Delegate> MigratedQuotaValueChangedWrappers =
            new Dictionary<DataGridView, Delegate>();
        private static readonly object MigratedQuotaCodesLock = new object();
        private static HashSet<string> migratedQuotaCodes;
        private static Dictionary<string, MigratedQuotaRecord> migratedQuotaRecords;
        private static bool migratedQuotaCodesLoadAttempted;
        private static Harmony findDeCompatibilityHarmony;
        private static bool findDeCompatibilityHookInstalled;
        private static bool findDeCompatibilityTypeMissingLogged;
        private static bool repairingMigratedQuotaRow;
        private static bool updatingMigratedQuotaCompatibilityRow;
        private static readonly HashSet<DataTable> HydratedMigratedQuotaCaches = new HashSet<DataTable>();
        private static readonly HashSet<Form> ProbedNativeQuotaLookups = new HashSet<Form>();
        private const int ReadOnlyCalcLearningMaxRowsPerGrid = 12;
        private const int ReadOnlyCalcLearningMaxColumnsPerRow = 120;
        private const int ReadOnlyCalcLearningMaxDataSetRows = 8;
        private static readonly string[] MigratedQuotaBooks =
        {
            "LG_2018", "QG_2018", "SG_2018", "GG_2018", "TG_2018", "XG_2018",
            "EG_2018", "DG_2018", "HG_2018", "FG_2018", "PG_2018", "JG_2018",
            "ZG_2018", "TZ_2020", "XZ_2020", "EZ_2020", "DZ_2020", "HZ_2020",
            "FZ_2020", "PZ_2020", "JZ_2020"
        };

        private sealed class MenuInfo
        {
            public Form MainForm;
            public string Name;
        }

        private sealed class MigratedQuotaRecord
        {
            public string Code;
            public string Book;
            public string Name;
            public string Unit;
            public string WorkContent;
            public string EncryptedConsume;
            public string PlainConsume;
            public string BasicQuota;
            public double BasePrice;
            public double Weight;
            public string SectionNo;
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
            // 2024 migration compatibility is intentionally disabled here.
            // After project settings include the migrated quota books, native
            // lookup/input/calculation should own the full workflow.
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

        private static void StartReadOnlyCalcLearningDiagnostics()
        {
            if (readOnlyCalcLearningTimer != null)
            {
                return;
            }

            readOnlyCalcLearningTimer = new System.Windows.Forms.Timer();
            readOnlyCalcLearningTimer.Interval = 1500;
            readOnlyCalcLearningTimer.Tick += delegate
            {
                try
                {
                    Form mainForm = FindMainForm();
                    if (mainForm != null)
                    {
                        LogReadOnlyCalcLearningSnapshot(mainForm);
                    }
                }
                catch (Exception ex)
                {
                    Log("ReadOnly calc learning failed: " + ex);
                }
            };
            readOnlyCalcLearningTimer.Start();
            Log("ReadOnly calc learning diagnostics started. no hooks, no event wrapping, no row writes.");
        }

        private static void StartMigratedQuotaLookupCompatibility()
        {
            if (migratedQuotaLookupTimer != null)
            {
                return;
            }

            migratedQuotaLookupTimer = new System.Windows.Forms.Timer();
            migratedQuotaLookupTimer.Interval = 500;
            migratedQuotaLookupTimer.Tick += delegate
            {
                try
                {
                    Form mainForm = FindMainForm();
                    if (mainForm != null)
                    {
                        InstallMigratedQuotaGridCompatibility(mainForm);
                        HydrateMigratedQuotaNativeCache(mainForm);
                    }
                }
                catch (Exception ex)
                {
                    Log("Migrated quota lookup compatibility failed: " + ex);
                }
            };
            migratedQuotaLookupTimer.Start();
            ThreadPool.QueueUserWorkItem(delegate { EnsureMigratedQuotaCodesLoaded(); });
            Log("Migrated quota compatibility started. input-event path only; no project row scan.");
        }

        private static void HydrateMigratedQuotaNativeCache(Form mainForm)
        {
            Dictionary<string, MigratedQuotaRecord> records = migratedQuotaRecords;
            if (mainForm == null || records == null || records.Count == 0)
            {
                return;
            }

            DataTable cache = GetField<DataTable>(mainForm, "dtCommonDe");
            if (!IsMigratedQuotaNativeCache(cache))
            {
                DataGridView grid = GetField<DataGridView>(mainForm, "dg_input");
                if (grid != null)
                {
                    cache = grid.DataSource as DataTable;
                    BindingSource source = grid.DataSource as BindingSource;
                    if (cache == null && source != null)
                    {
                        cache = source.DataSource as DataTable;
                    }
                }
            }

            if (cache == null || HydratedMigratedQuotaCaches.Contains(cache))
            {
                return;
            }

            DataColumn codeColumn = FindDataColumn(cache, "\u7535\u7b97\u4ee3\u53f7");
            DataColumn nameColumn = FindDataColumn(cache, "\u540d\u79f0");
            DataColumn unitColumn = FindDataColumn(cache, "\u5355\u4f4d");
            DataColumn contentColumn = FindDataColumn(cache, "\u73b0\u4ef7");
            DataColumn basePriceColumn = FindDataColumn(cache, "\u57fa\u4ef7");
            DataColumn sectionColumn = FindDataColumn(cache, "\u8282\u53f7");
            DataColumn actionColumn = FindDataColumn(cache, "\u64cd\u4f5c");
            if (!IsMigratedQuotaNativeCache(cache))
            {
                Log("Migrated quota native cache skipped: incompatible schema columns="
                    + String.Join("|", cache.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToArray()));
                return;
            }

            HashSet<string> existingCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (DataRow row in cache.Rows)
            {
                string code = Convert.ToString(row[codeColumn], CultureInfo.InvariantCulture);
                if (!String.IsNullOrWhiteSpace(code))
                {
                    existingCodes.Add(code.Trim());
                }
            }

            int added = 0;
            cache.BeginLoadData();
            try
            {
                foreach (MigratedQuotaRecord record in records.Values)
                {
                    if (record == null || String.IsNullOrWhiteSpace(record.Code) ||
                        existingCodes.Contains(record.Code))
                    {
                        continue;
                    }

                    DataRow row = cache.NewRow();
                    SetDataRowColumn(row, codeColumn, record.Code);
                    SetDataRowColumn(row, nameColumn, record.Name);
                    SetDataRowColumn(row, unitColumn, record.Unit);
                    SetDataRowColumn(row, contentColumn, record.WorkContent);
                    SetDataRowColumn(row, basePriceColumn, record.BasePrice);
                    SetDataRowColumn(row, sectionColumn, record.SectionNo);
                    SetDataRowColumn(row, actionColumn, "\u663e\u793a\u6d88\u8017");
                    cache.Rows.Add(row);
                    row.AcceptChanges();
                    existingCodes.Add(record.Code);
                    added++;
                }
            }
            finally
            {
                cache.EndLoadData();
            }

            HydratedMigratedQuotaCaches.Add(cache);
            Log("Migrated quota native cache hydrated: added="
                + added.ToString(CultureInfo.InvariantCulture)
                + " total=" + cache.Rows.Count.ToString(CultureInfo.InvariantCulture)
                + " original-preserved=" + (cache.Rows.Count - added).ToString(CultureInfo.InvariantCulture));
            if (!ProbedNativeQuotaLookups.Contains(mainForm))
            {
                ProbedNativeQuotaLookups.Add(mainForm);
                ProbeMigratedQuotaLookup(mainForm, "DY-1");
                ProbeMigratedQuotaLookup(mainForm, "LG-12");
            }
        }

        private static bool IsMigratedQuotaNativeCache(DataTable table)
        {
            return table != null &&
                FindDataColumn(table, "\u7535\u7b97\u4ee3\u53f7") != null &&
                FindDataColumn(table, "\u540d\u79f0") != null &&
                FindDataColumn(table, "\u5355\u4f4d") != null &&
                FindDataColumn(table, "\u73b0\u4ef7") != null &&
                FindDataColumn(table, "\u57fa\u4ef7") != null &&
                FindDataColumn(table, "\u8282\u53f7") != null &&
                FindDataColumn(table, "\u64cd\u4f5c") != null;
        }

        private static DataColumn FindDataColumn(DataTable table, string name)
        {
            if (table == null || String.IsNullOrEmpty(name))
            {
                return null;
            }

            foreach (DataColumn column in table.Columns)
            {
                if (String.Equals(column.ColumnName, name, StringComparison.OrdinalIgnoreCase))
                {
                    return column;
                }
            }
            return null;
        }

        private static void SetDataRowColumn(DataRow row, DataColumn column, object value)
        {
            if (row == null || column == null)
            {
                return;
            }

            if (value == null)
            {
                row[column] = DBNull.Value;
                return;
            }

            Type targetType = Nullable.GetUnderlyingType(column.DataType) ?? column.DataType;
            if (targetType == typeof(string))
            {
                row[column] = Convert.ToString(value, CultureInfo.CurrentCulture);
                return;
            }

            row[column] = Convert.ChangeType(value, targetType, CultureInfo.InvariantCulture);
        }

        private static void InstallMigratedQuotaGridCompatibility(Form mainForm)
        {
            DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
            if (grid == null)
            {
                return;
            }

            PropertyInfo eventsProperty = typeof(Component).GetProperty(
                "Events",
                BindingFlags.Instance | BindingFlags.NonPublic);
            EventHandlerList events = eventsProperty == null
                ? null
                : eventsProperty.GetValue(grid, null) as EventHandlerList;
            if (events == null)
            {
                return;
            }

            bool validatingWrapped = false;
            bool valueChangedWrapped = false;
            for (Type type = grid.GetType(); type != null; type = type.BaseType)
            {
                foreach (FieldInfo field in type.GetFields(
                    BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.DeclaredOnly))
                {
                    if (field.FieldType != typeof(object) ||
                        field.Name.IndexOf("EVENT", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        continue;
                    }

                    object key;
                    try
                    {
                        key = field.GetValue(null);
                    }
                    catch
                    {
                        continue;
                    }

                    Delegate handler = events[key];
                    if (handler == null)
                    {
                        continue;
                    }

                    if (!validatingWrapped &&
                        String.Equals(field.Name, "EVENT_DATAGRIDVIEWCELLVALIDATING", StringComparison.OrdinalIgnoreCase) &&
                        handler is DataGridViewCellValidatingEventHandler)
                    {
                        Delegate installedWrapper;
                        if (MigratedQuotaValidatingWrappers.TryGetValue(grid, out installedWrapper) &&
                            Object.ReferenceEquals(handler, installedWrapper))
                        {
                            validatingWrapped = true;
                            continue;
                        }

                        DataGridViewCellValidatingEventHandler original =
                            (DataGridViewCellValidatingEventHandler)handler;
                        DataGridViewCellValidatingEventHandler wrapper = delegate(
                            object sender,
                            DataGridViewCellValidatingEventArgs e)
                        {
                            string code = NormalizeMigratedQuotaCode(
                                Convert.ToString(e.FormattedValue, CultureInfo.CurrentCulture));
                            MigratedQuotaRecord record;
                            bool isCodeColumn = IsQuotaCodeColumn(grid, e.ColumnIndex);
                            if (HasMigratedQuotaPrefix(code))
                            {
                                Log("Migrated quota validating seen: code=" + code
                                    + " row=" + e.RowIndex.ToString(CultureInfo.InvariantCulture)
                                    + " col=" + e.ColumnIndex.ToString(CultureInfo.InvariantCulture)
                                    + " isCodeColumn=" + isCodeColumn.ToString());
                            }
                            if (isCodeColumn && TryGetMigratedQuotaRecord(code, out record))
                            {
                                e.Cancel = false;
                                ApplyMigratedQuotaRecord(mainForm, grid, e.RowIndex, record, "validate");
                                return;
                            }
                            if (isCodeColumn && HasMigratedQuotaPrefix(code))
                            {
                                Log("Migrated quota validating passed to native: code=" + code
                                    + " cacheLoaded=" + (migratedQuotaRecords != null).ToString());
                            }

                            original(sender, e);
                        };
                        events.RemoveHandler(key, handler);
                        events.AddHandler(key, wrapper);
                        MigratedQuotaValidatingWrappers[grid] = wrapper;
                        validatingWrapped = true;
                        Log("Migrated quota validating handler wrapped: handlers="
                            + handler.GetInvocationList().Length.ToString(CultureInfo.InvariantCulture));
                    }
                    else if (!valueChangedWrapped &&
                        String.Equals(field.Name, "EVENT_DATAGRIDVIEWCELLVALUECHANGED", StringComparison.OrdinalIgnoreCase) &&
                        handler is DataGridViewCellEventHandler)
                    {
                        Delegate installedWrapper;
                        if (MigratedQuotaValueChangedWrappers.TryGetValue(grid, out installedWrapper) &&
                            Object.ReferenceEquals(handler, installedWrapper))
                        {
                            valueChangedWrapped = true;
                            continue;
                        }

                        DataGridViewCellEventHandler original = (DataGridViewCellEventHandler)handler;
                        DataGridViewCellEventHandler wrapper = delegate(object sender, DataGridViewCellEventArgs e)
                        {
                            if (updatingMigratedQuotaCompatibilityRow)
                            {
                                return;
                            }

                            MigratedQuotaRecord record;
                            if (TryGetMigratedQuotaRecordFromRow(grid, e.RowIndex, out record) &&
                                (IsQuotaCodeColumn(grid, e.ColumnIndex) ||
                                 IsQuotaQuantityColumn(grid, e.ColumnIndex)))
                            {
                                ApplyMigratedQuotaRecord(mainForm, grid, e.RowIndex, record, "value-changed");
                                return;
                            }

                            original(sender, e);
                        };
                        events.RemoveHandler(key, handler);
                        events.AddHandler(key, wrapper);
                        MigratedQuotaValueChangedWrappers[grid] = wrapper;
                        valueChangedWrapped = true;
                        Log("Migrated quota value-changed handler wrapped: handlers="
                            + handler.GetInvocationList().Length.ToString(CultureInfo.InvariantCulture));
                    }
                }
            }

            if (validatingWrapped && valueChangedWrapped)
            {
                if (!MigratedQuotaCompatibilityGrids.Contains(grid))
                {
                    MigratedQuotaCompatibilityGrids.Add(grid);
                    Log("Migrated quota grid compatibility installed.");
                }
            }
        }

        private static bool IsQuotaCodeColumn(DataGridView grid, int columnIndex)
        {
            return IsGridColumn(grid, columnIndex, "\u5b9a\u989d\u7f16\u53f7");
        }

        private static bool IsQuotaQuantityColumn(DataGridView grid, int columnIndex)
        {
            return IsGridColumn(grid, columnIndex, "\u5de5\u7a0b\u6570\u91cf\u8f93\u5165") ||
                IsGridColumn(grid, columnIndex, "\u5de5\u7a0b\u6570\u91cf");
        }

        private static bool IsGridColumn(DataGridView grid, int columnIndex, string expected)
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

        private static bool TryGetMigratedQuotaRecordFromRow(
            DataGridView grid,
            int rowIndex,
            out MigratedQuotaRecord record)
        {
            record = null;
            if (grid == null || rowIndex < 0 || rowIndex >= grid.Rows.Count)
            {
                return false;
            }

            DataGridViewRow gridRow = grid.Rows[rowIndex];
            DataRowView rowView = gridRow.DataBoundItem as DataRowView;
            string code = ReadQuotaRowText(
                rowView == null ? null : rowView.Row,
                gridRow,
                "\u5b9a\u989d\u7f16\u53f7");
            return TryGetMigratedQuotaRecord(code, out record);
        }

        private static void RepairAllMigratedQuotaRows(Form mainForm, string reason)
        {
            if (!IsMigratedQuotaRowRepairEnabled())
            {
                return;
            }

            DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
            if (grid == null || updatingMigratedQuotaCompatibilityRow)
            {
                return;
            }

            for (int rowIndex = 0; rowIndex < grid.Rows.Count; rowIndex++)
            {
                MigratedQuotaRecord record;
                if (TryGetMigratedQuotaRecordFromRow(grid, rowIndex, out record))
                {
                    ApplyMigratedQuotaRecord(mainForm, grid, rowIndex, record, reason);
                }
            }
        }

        private static void ApplyMigratedQuotaRecord(
            Form mainForm,
            DataGridView grid,
            int rowIndex,
            MigratedQuotaRecord record,
            string reason)
        {
            if (grid == null || record == null || rowIndex < 0 || rowIndex >= grid.Rows.Count)
            {
                return;
            }

            try
            {
                updatingMigratedQuotaCompatibilityRow = true;
                DataGridViewRow gridRow = grid.Rows[rowIndex];
                DataRowView rowView = gridRow.DataBoundItem as DataRowView;
                DataRow dataRow = rowView == null ? null : rowView.Row;

                if (String.IsNullOrEmpty(record.PlainConsume))
                {
                    record.PlainConsume = DecryptConsumeForCurrentApp(record.EncryptedConsume);
                }

                double quantity;
                bool hasInputQuantity = TryReadQuotaRowDouble(
                    dataRow,
                    gridRow,
                    "\u5de5\u7a0b\u6570\u91cf\u8f93\u5165",
                    out quantity);
                if (!hasInputQuantity)
                {
                    TryReadQuotaRowDouble(
                        dataRow,
                        gridRow,
                        "\u5de5\u7a0b\u6570\u91cf",
                        out quantity);
                }

                double price = record.BasePrice;
                try
                {
                    price = CalculateMigratedQuotaCurrentPrice(mainForm, record);
                }
                catch (Exception ex)
                {
                    Log("Migrated quota price calculation failed; using base price: code="
                        + record.Code + " " + ex);
                }
                double total = quantity * price;

                SetQuotaRowValue(dataRow, gridRow, "\u5b9a\u989d\u7f16\u53f7", record.Code);
                SetQuotaRowValue(dataRow, gridRow, "\u5de5\u7a0b\u6216\u8d39\u7528\u9879\u76ee\u540d\u79f0", record.Name);
                SetQuotaRowValue(dataRow, gridRow, "\u5355\u4f4d", record.Unit);
                if (hasInputQuantity)
                {
                    SetQuotaRowValue(dataRow, gridRow, "\u5de5\u7a0b\u6570\u91cf", quantity);
                }
                SetQuotaRowValue(dataRow, gridRow, "\u5355\u4ef7", price);
                SetQuotaRowValue(dataRow, gridRow, "\u5408\u4ef7", total);
                SetQuotaRowValue(dataRow, gridRow, "\u5355\u91cd", record.Weight);
                SetQuotaRowValue(dataRow, gridRow, "\u57fa\u672c\u5b9a\u989d", record.BasicQuota ?? "");
                SetQuotaRowValue(dataRow, gridRow, "\u5b9a\u989d\u6d88\u8017", record.PlainConsume);

                string signature = "compat:" + record.Code + ":"
                    + quantity.ToString("0.########", CultureInfo.InvariantCulture) + ":"
                    + price.ToString("0.########", CultureInfo.InvariantCulture);
                string oldSignature;
                if (!ReadOnlyCalcLearningSignatures.TryGetValue(signature, out oldSignature))
                {
                    ReadOnlyCalcLearningSignatures[signature] = signature;
                    Log("Migrated quota row updated: code=" + record.Code
                        + " reason=" + reason
                        + " quantity=" + quantity.ToString(CultureInfo.InvariantCulture)
                        + " price=" + price.ToString(CultureInfo.InvariantCulture));
                }
            }
            catch (Exception ex)
            {
                Log("Migrated quota row update failed: code=" + record.Code + " " + ex);
            }
            finally
            {
                updatingMigratedQuotaCompatibilityRow = false;
            }
        }

        private static double CalculateMigratedQuotaCurrentPrice(
            Form mainForm,
            MigratedQuotaRecord record)
        {
            DataSet commonDataSet = GetReadOnlyCommonDataSet(mainForm);
            if (commonDataSet == null || String.IsNullOrEmpty(record.PlainConsume))
            {
                return record.BasePrice;
            }

            double total = 0d;
            string[] parts = record.PlainConsume.Split(new[] { '@' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts)
            {
                if (part.Length <= 10)
                {
                    return record.BasePrice;
                }

                int code;
                double quantity;
                if (!Int32.TryParse(part.Substring(0, 10), NumberStyles.Integer, CultureInfo.InvariantCulture, out code) ||
                    !Double.TryParse(part.Substring(10), NumberStyles.Float, CultureInfo.InvariantCulture, out quantity))
                {
                    return record.BasePrice;
                }

                double unitPrice;
                if (!TryGetCurrentResourcePrice(commonDataSet, code, out unitPrice))
                {
                    Log("Migrated quota price fallback: code=" + record.Code
                        + " missingResource=" + code.ToString(CultureInfo.InvariantCulture));
                    return record.BasePrice;
                }

                total += quantity * unitPrice;
            }

            return Math.Round(total, 2, MidpointRounding.AwayFromZero);
        }

        private static bool TryGetCurrentResourcePrice(DataSet dataSet, int code, out double price)
        {
            if (code >= 1 && code <= 7)
            {
                return TryGetDataSetPrice(
                    dataSet,
                    "\u5de5\u8d39\u65b9\u6848",
                    code,
                    new[] { "\u7f16\u5236\u671f\u6807\u51c6", "\u57fa\u671f\u6807\u51c6" },
                    out price);
            }

            if (TryGetDataSetPrice(
                dataSet,
                "\u6599\u8d39\u65b9\u6848",
                code,
                new[] { "\u7f16\u5236\u671f\u4ef7", "\u57fa\u671f\u5355\u4ef7" },
                out price) ||
                TryGetDataSetPrice(
                    dataSet,
                    "\u6750\u6599\u5355\u4ef7",
                    code,
                    new[] { "\u7f16\u5236\u671f\u4ef7", "\u57fa\u671f\u5355\u4ef7" },
                    out price))
            {
                return true;
            }

            DataRow machineRow;
            if (TryGetDataSetRow(dataSet, "\u673a\u8d39\u65b9\u6848", code, out machineRow) ||
                TryGetDataSetRow(dataSet, "\u53f0\u73ed\u5355\u4ef7", code, out machineRow))
            {
                return TryCalculateMachinePrice(dataSet, machineRow, out price);
            }

            price = 0d;
            return false;
        }

        private static bool TryCalculateMachinePrice(DataSet dataSet, DataRow row, out double price)
        {
            double fixedCost =
                ReadDataRowDouble(row, "\u6298\u65e7\u8d39") +
                ReadDataRowDouble(row, "\u68c0\u4fee\u8d39") +
                ReadDataRowDouble(row, "\u7ef4\u62a4\u8d39") +
                ReadDataRowDouble(row, "\u5b89\u88c5\u62c6\u5378\u8d39") +
                ReadDataRowDouble(row, "\u5176\u4ed6\u8d39\u7528");
            double laborPrice;
            if (!TryGetDataSetPrice(
                dataSet,
                "\u5de5\u8d39\u65b9\u6848",
                3,
                new[] { "\u7f16\u5236\u671f\u6807\u51c6", "\u57fa\u671f\u6807\u51c6" },
                out laborPrice))
            {
                laborPrice = 0d;
            }

            price = fixedCost + ReadDataRowDouble(row, "\u4eba\u5de5") * laborPrice;
            string[] quantityColumns =
            {
                "\u6c7d\u6cb9", "\u67f4\u6cb9", "\u7164", "\u5929\u7136\u6c14", "\u7535", "\u6c34"
            };
            int[] materialCodes = { 32, 33, 35, 37, 31, 50 };
            for (int i = 0; i < quantityColumns.Length; i++)
            {
                double resourcePrice;
                double resourceQuantity = ReadDataRowDouble(row, quantityColumns[i]);
                if (resourceQuantity != 0d &&
                    TryGetDataSetPrice(
                        dataSet,
                        "\u6599\u8d39\u65b9\u6848",
                        materialCodes[i],
                        new[] { "\u7f16\u5236\u671f\u4ef7", "\u57fa\u671f\u5355\u4ef7" },
                        out resourcePrice))
                {
                    price += resourceQuantity * resourcePrice;
                }
            }

            if (price == 0d)
            {
                price = ReadDataRowDouble(row, "\u57fa\u4ef7");
            }
            return price != 0d;
        }

        private static bool TryGetDataSetPrice(
            DataSet dataSet,
            string tableName,
            int code,
            string[] priceColumns,
            out double price)
        {
            DataRow row;
            if (TryGetDataSetRow(dataSet, tableName, code, out row))
            {
                foreach (string priceColumn in priceColumns)
                {
                    if (row.Table.Columns.Contains(priceColumn) &&
                        row[priceColumn] != DBNull.Value)
                    {
                        price = Convert.ToDouble(row[priceColumn], CultureInfo.InvariantCulture);
                        return true;
                    }
                }
            }

            price = 0d;
            return false;
        }

        private static bool TryGetDataSetRow(
            DataSet dataSet,
            string tableName,
            int code,
            out DataRow row)
        {
            row = null;
            if (dataSet == null || !dataSet.Tables.Contains(tableName))
            {
                return false;
            }

            DataTable table = dataSet.Tables[tableName];
            string codeColumn = "\u7535\u7b97\u4ee3\u53f7";
            if (!table.Columns.Contains(codeColumn))
            {
                return false;
            }

            foreach (DataRow candidate in table.Rows)
            {
                long candidateCode;
                if (candidate[codeColumn] != DBNull.Value &&
                    Int64.TryParse(
                        Convert.ToString(candidate[codeColumn], CultureInfo.InvariantCulture),
                        NumberStyles.Integer,
                        CultureInfo.InvariantCulture,
                        out candidateCode) &&
                    candidateCode == code)
                {
                    row = candidate;
                    return true;
                }
            }
            return false;
        }

        private static double ReadDataRowDouble(DataRow row, string columnName)
        {
            if (row == null || row.Table == null || !row.Table.Columns.Contains(columnName) ||
                row[columnName] == DBNull.Value)
            {
                return 0d;
            }
            return Convert.ToDouble(row[columnName], CultureInfo.InvariantCulture);
        }

        private static void ClearMigratedQuotaConsumeErrors(Form mainForm)
        {
            DataGridView errorGrid = GetField<DataGridView>(mainForm, "dg_err");
            if (errorGrid == null)
            {
                return;
            }

            DataTable table = errorGrid.DataSource as DataTable;
            BindingSource source = errorGrid.DataSource as BindingSource;
            if (table == null && source != null)
            {
                table = source.DataSource as DataTable;
            }
            if (table == null || !table.Columns.Contains("\u9519\u8bef\u63cf\u8ff0"))
            {
                return;
            }

            for (int i = table.Rows.Count - 1; i >= 0; i--)
            {
                string message = Convert.ToString(
                    table.Rows[i]["\u9519\u8bef\u63cf\u8ff0"],
                    CultureInfo.CurrentCulture);
                if (message.IndexOf("\u65e0\u6cd5\u627e\u5230\u5b9a\u989d'", StringComparison.Ordinal) < 0 ||
                    message.IndexOf("\u7684\u6d88\u8017\u6570\u636e", StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                foreach (string prefix in new[]
                {
                    "LG-", "QG-", "SG-", "GG-", "TG-", "XG-", "EG-", "DG-", "HG-", "FG-",
                    "PG-", "JG-", "ZG-", "TZ-", "XZ-", "EZ-", "DZ-", "HZ-", "FZ-", "PZ-", "JZ-"
                })
                {
                    if (message.IndexOf(prefix, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        table.Rows.RemoveAt(i);
                        break;
                    }
                }
            }
        }

        private static void LogReadOnlyCalcLearningSnapshot(Form mainForm)
        {
            LogReadOnlyDataSetDeepSnapshot(mainForm);

            List<DataGridView> grids = new List<DataGridView>();
            AddReadOnlyGrid(grids, GetField<DataGridView>(mainForm, "dataGridViewDE"));
            AddReadOnlyGrid(grids, GetField<DataGridView>(mainForm, "dataGridViewProp"));
            CollectReadOnlyDataGrids(mainForm, grids);

            HashSet<string> quotaCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < grids.Count; i++)
            {
                DataGridView grid = grids[i];
                if (grid == null || grid.IsDisposed || !grid.IsHandleCreated)
                {
                    continue;
                }

                if (!ReadOnlyGridHasContent(grid))
                {
                    continue;
                }

                string source = "grid:" + grid.GetHashCode().ToString(CultureInfo.InvariantCulture) + ":" + GetReadOnlyControlPath(grid);
                LogReadOnlyGridIfChanged(source, grid, quotaCodes);
            }

            LogReadOnlyProjectRowsForCodes(mainForm, quotaCodes);
        }

        private static void LogReadOnlyDataSetDeepSnapshot(Form mainForm)
        {
            try
            {
                DataSet commonDataSet = GetReadOnlyCommonDataSet(mainForm);
                DataSet projectDataSet = GetField<DataSet>(mainForm, "m_DataSet");
                LogReadOnlyDataSetDeepSnapshot("common", commonDataSet);
                LogReadOnlyDataSetDeepSnapshot("project", projectDataSet);
            }
            catch (Exception ex)
            {
                Log("ReadOnly dataset snapshot failed: " + ex.Message);
            }
        }

        private static DataSet GetReadOnlyCommonDataSet(Form mainForm)
        {
            if (mainForm == null)
            {
                return null;
            }

            FieldInfo commonField = mainForm.GetType().GetField(
                "m_CmmonDataSet",
                BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
            object common = commonField == null ? null : commonField.GetValue(null);
            FieldInfo dataSetField = common == null ? null : common.GetType().GetField(
                "m_DataSet",
                BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
            return dataSetField == null ? null : dataSetField.GetValue(common) as DataSet;
        }

        private static void LogReadOnlyDataSetDeepSnapshot(string label, DataSet dataSet)
        {
            if (dataSet == null)
            {
                return;
            }

            string dataSetKey = "dataset:" + label + ":" + dataSet.GetHashCode().ToString(CultureInfo.InvariantCulture);
            if (!ReadOnlyCalcLearningDataSetSchemaLogs.Contains(dataSetKey))
            {
                ReadOnlyCalcLearningDataSetSchemaLogs.Add(dataSetKey);
                Log("ReadOnly dataset snapshot: label=" + label
                    + " name=" + dataSet.DataSetName
                    + " tables=" + dataSet.Tables.Count.ToString(CultureInfo.InvariantCulture));
            }

            foreach (DataTable table in dataSet.Tables)
            {
                LogReadOnlyDataTableDeepSnapshot(label, table);
            }
        }

        private static void LogReadOnlyDataTableDeepSnapshot(string label, DataTable table)
        {
            if (table == null)
            {
                return;
            }

            string tableKey = "dataset-table:" + label + ":" + table.GetHashCode().ToString(CultureInfo.InvariantCulture);
            string[] columnNames = table.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToArray();
            string schema = table.TableName + "|" + table.Rows.Count.ToString(CultureInfo.InvariantCulture)
                + "|" + String.Join("|", columnNames);
            if (!ReadOnlyCalcLearningDataSetSchemaLogs.Contains(tableKey))
            {
                ReadOnlyCalcLearningDataSetSchemaLogs.Add(tableKey);
                Log("ReadOnly dataset table schema: label=" + label
                    + " table=" + table.TableName
                    + " rows=" + table.Rows.Count.ToString(CultureInfo.InvariantCulture)
                    + " cols=" + String.Join("|", columnNames));
            }

            int lg2 = CountTableMatches(table, "LG-2");
            int lg12 = CountTableMatches(table, "LG-12");
            int lg14 = CountTableMatches(table, "LG-14");
            int dy1 = CountTableMatches(table, "DY-1");
            int lgBook = CountTableMatches(table, "LG_2018");
            int dyBook = CountTableMatches(table, "DY_2024");
            bool interesting = IsReadOnlyInterestingDataTable(table) ||
                lg2 > 0 || lg12 > 0 || lg14 > 0 || dy1 > 0 || lgBook > 0 || dyBook > 0;
            if (!interesting)
            {
                return;
            }

            string summary = schema
                + "|LG-2=" + lg2.ToString(CultureInfo.InvariantCulture)
                + "|LG-12=" + lg12.ToString(CultureInfo.InvariantCulture)
                + "|LG-14=" + lg14.ToString(CultureInfo.InvariantCulture)
                + "|DY-1=" + dy1.ToString(CultureInfo.InvariantCulture)
                + "|LG_2018=" + lgBook.ToString(CultureInfo.InvariantCulture)
                + "|DY_2024=" + dyBook.ToString(CultureInfo.InvariantCulture);
            string signatureKey = "dataset-table-summary:" + label + ":" + table.GetHashCode().ToString(CultureInfo.InvariantCulture);
            string oldSummary;
            if (ReadOnlyCalcLearningSignatures.TryGetValue(signatureKey, out oldSummary) &&
                String.Equals(oldSummary, summary, StringComparison.Ordinal))
            {
                return;
            }

            ReadOnlyCalcLearningSignatures[signatureKey] = summary;
            Log("ReadOnly dataset table summary: label=" + label
                + " table=" + table.TableName
                + " rows=" + table.Rows.Count.ToString(CultureInfo.InvariantCulture)
                + " LG-2=" + lg2.ToString(CultureInfo.InvariantCulture)
                + " LG-12=" + lg12.ToString(CultureInfo.InvariantCulture)
                + " LG-14=" + lg14.ToString(CultureInfo.InvariantCulture)
                + " DY-1=" + dy1.ToString(CultureInfo.InvariantCulture)
                + " LG_2018=" + lgBook.ToString(CultureInfo.InvariantCulture)
                + " DY_2024=" + dyBook.ToString(CultureInfo.InvariantCulture));

            LogReadOnlyDataTableSampleRows(label, table, lg2 + lg12 + lg14 + dy1 + lgBook + dyBook > 0);
        }

        private static bool IsReadOnlyInterestingDataTable(DataTable table)
        {
            string name = table.TableName ?? "";
            if (IsReadOnlyInterestingColumnName(name))
            {
                return true;
            }

            foreach (DataColumn column in table.Columns)
            {
                if (IsReadOnlyInterestingColumnName(column.ColumnName))
                {
                    return true;
                }
            }

            return false;
        }

        private static bool IsReadOnlyInterestingColumnName(string name)
        {
            string text = name ?? "";
            return text.IndexOf("定额", StringComparison.Ordinal) >= 0 ||
                text.IndexOf("消耗", StringComparison.Ordinal) >= 0 ||
                text.IndexOf("电算", StringComparison.Ordinal) >= 0 ||
                text.IndexOf("单价", StringComparison.Ordinal) >= 0 ||
                text.IndexOf("合价", StringComparison.Ordinal) >= 0 ||
                text.IndexOf("工费", StringComparison.Ordinal) >= 0 ||
                text.IndexOf("料费", StringComparison.Ordinal) >= 0 ||
                text.IndexOf("机费", StringComparison.Ordinal) >= 0;
        }

        private static void LogReadOnlyDataTableSampleRows(string label, DataTable table, bool preferMatches)
        {
            int logged = 0;
            for (int i = 0; i < table.Rows.Count && logged < ReadOnlyCalcLearningMaxDataSetRows; i++)
            {
                DataRow row = table.Rows[i];
                bool matched = ReadOnlyDataRowTextContains(row, "LG-2") ||
                    ReadOnlyDataRowTextContains(row, "LG-12") ||
                    ReadOnlyDataRowTextContains(row, "LG-14") ||
                    ReadOnlyDataRowTextContains(row, "DY-1") ||
                    ReadOnlyDataRowTextContains(row, "LG_2018") ||
                    ReadOnlyDataRowTextContains(row, "DY_2024");
                if (preferMatches && !matched)
                {
                    continue;
                }

                Log("ReadOnly dataset table row: label=" + label
                    + " table=" + table.TableName
                    + " row=" + i.ToString(CultureInfo.InvariantCulture)
                    + " matched=" + matched.ToString()
                    + " {" + BuildReadOnlyDataRowText(row, null) + "}");
                logged++;
            }
        }

        private static bool ReadOnlyDataRowTextContains(DataRow row, string expected)
        {
            if (row == null || row.Table == null)
            {
                return false;
            }

            foreach (DataColumn column in row.Table.Columns)
            {
                string text = "";
                try
                {
                    text = Convert.ToString(row[column], CultureInfo.CurrentCulture);
                }
                catch
                {
                }

                if (text.IndexOf(expected, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static void AddReadOnlyGrid(List<DataGridView> grids, DataGridView grid)
        {
            if (grid == null)
            {
                return;
            }

            foreach (DataGridView existing in grids)
            {
                if (Object.ReferenceEquals(existing, grid))
                {
                    return;
                }
            }

            grids.Add(grid);
        }

        private static void CollectReadOnlyDataGrids(Control parent, List<DataGridView> grids)
        {
            if (parent == null)
            {
                return;
            }

            foreach (Control child in parent.Controls)
            {
                DataGridView grid = child as DataGridView;
                if (grid != null)
                {
                    AddReadOnlyGrid(grids, grid);
                }

                CollectReadOnlyDataGrids(child, grids);
            }
        }

        private static bool ReadOnlyGridHasContent(DataGridView grid)
        {
            if (grid == null || grid.Rows.Count == 0)
            {
                return false;
            }

            int checkedRows = 0;
            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row == null || row.IsNewRow)
                {
                    continue;
                }

                checkedRows++;
                if (ReadOnlyRowHasContent(row))
                {
                    return true;
                }

                if (checkedRows >= ReadOnlyCalcLearningMaxRowsPerGrid)
                {
                    break;
                }
            }

            return false;
        }

        private static bool ReadOnlyRowHasContent(DataGridViewRow row)
        {
            DataRowView rowView = row.DataBoundItem as DataRowView;
            if (rowView != null && rowView.Row != null && rowView.Row.Table != null)
            {
                foreach (DataColumn column in rowView.Row.Table.Columns)
                {
                    if (!String.IsNullOrWhiteSpace(Convert.ToString(rowView.Row[column], CultureInfo.CurrentCulture)))
                    {
                        return true;
                    }
                }
            }

            DataGridView grid = row.DataGridView;
            if (grid == null)
            {
                return false;
            }

            foreach (DataGridViewColumn column in grid.Columns)
            {
                object raw = row.Cells[column.Index].Value;
                if (!String.IsNullOrWhiteSpace(Convert.ToString(raw, CultureInfo.CurrentCulture)))
                {
                    return true;
                }
            }

            return false;
        }

        private static void LogReadOnlyGridIfChanged(string source, DataGridView grid, HashSet<string> quotaCodes)
        {
            List<string> rowLogs = new List<string>();
            string signature = BuildReadOnlyGridSignature(source, grid, quotaCodes, rowLogs);
            string oldSignature;
            if (ReadOnlyCalcLearningSignatures.TryGetValue(source, out oldSignature) &&
                String.Equals(oldSignature, signature, StringComparison.Ordinal))
            {
                return;
            }

            ReadOnlyCalcLearningSignatures[source] = signature;
            Log("ReadOnly calc grid snapshot: source=" + source
                + " rows=" + grid.Rows.Count.ToString(CultureInfo.InvariantCulture)
                + " cols=" + grid.Columns.Count.ToString(CultureInfo.InvariantCulture)
                + " current=" + DescribeReadOnlyCurrentCell(grid)
                + " editing=" + (grid.EditingControl == null ? "" : SafeReadOnlyText(grid.EditingControl.Text, 80))
                + " dataSource=" + DescribeReadOnlyDataSource(grid.DataSource));

            if (!ReadOnlyCalcLearningColumnLogs.Contains(source))
            {
                ReadOnlyCalcLearningColumnLogs.Add(source);
                Log("ReadOnly calc grid columns: source=" + source + " " + DescribeReadOnlyColumns(grid));
            }

            foreach (string rowLog in rowLogs)
            {
                Log(rowLog);
            }
        }

        private static string BuildReadOnlyGridSignature(
            string source,
            DataGridView grid,
            HashSet<string> quotaCodes,
            List<string> rowLogs)
        {
            StringBuilder signature = new StringBuilder();
            signature.Append("rows=").Append(grid.Rows.Count.ToString(CultureInfo.InvariantCulture));
            signature.Append(";cols=").Append(grid.Columns.Count.ToString(CultureInfo.InvariantCulture));
            signature.Append(";current=").Append(DescribeReadOnlyCurrentCell(grid));
            signature.Append(";editing=").Append(grid.EditingControl == null ? "" : grid.EditingControl.Text);

            int loggedRows = 0;
            foreach (DataGridViewRow row in grid.Rows)
            {
                if (row == null || row.IsNewRow)
                {
                    continue;
                }

                string visible = DescribeReadOnlyVisibleRow(row, quotaCodes);
                string bound = DescribeReadOnlyBoundRow(row, quotaCodes);
                if (String.IsNullOrEmpty(visible) && String.IsNullOrEmpty(bound))
                {
                    continue;
                }

                signature.Append("|row").Append(row.Index.ToString(CultureInfo.InvariantCulture))
                    .Append(":").Append(visible).Append(":").Append(bound);
                rowLogs.Add("ReadOnly calc grid row: source=" + source
                    + " row=" + row.Index.ToString(CultureInfo.InvariantCulture)
                    + " visible={" + visible + "}"
                    + " bound={" + bound + "}");
                loggedRows++;
                if (loggedRows >= ReadOnlyCalcLearningMaxRowsPerGrid)
                {
                    break;
                }
            }

            return signature.ToString();
        }

        private static string DescribeReadOnlyCurrentCell(DataGridView grid)
        {
            DataGridViewCell cell = grid.CurrentCell;
            if (cell == null)
            {
                return "";
            }

            string column = cell.OwningColumn == null ? "" : ReadOnlyColumnLabel(cell.OwningColumn);
            return cell.RowIndex.ToString(CultureInfo.InvariantCulture)
                + "," + cell.ColumnIndex.ToString(CultureInfo.InvariantCulture)
                + ":" + column + "=" + SafeReadOnlyText(Convert.ToString(cell.Value, CultureInfo.CurrentCulture), 80);
        }

        private static string DescribeReadOnlyColumns(DataGridView grid)
        {
            List<string> parts = new List<string>();
            foreach (DataGridViewColumn column in grid.Columns)
            {
                parts.Add(column.Index.ToString(CultureInfo.InvariantCulture)
                    + ":name=" + column.Name
                    + ",header=" + column.HeaderText
                    + ",property=" + column.DataPropertyName);
            }

            return String.Join(" | ", parts.ToArray());
        }

        private static string DescribeReadOnlyVisibleRow(DataGridViewRow row, HashSet<string> quotaCodes)
        {
            DataGridView grid = row.DataGridView;
            if (grid == null)
            {
                return "";
            }

            List<string> parts = new List<string>();
            foreach (DataGridViewColumn column in grid.Columns)
            {
                object raw = row.Cells[column.Index].Value;
                string text = Convert.ToString(raw, CultureInfo.CurrentCulture);
                if (String.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                AddReadOnlyQuotaCodes(quotaCodes, text);
                parts.Add(ReadOnlyColumnLabel(column) + "=" + SafeReadOnlyText(text, 160));
                if (parts.Count >= ReadOnlyCalcLearningMaxColumnsPerRow)
                {
                    break;
                }
            }

            return String.Join(";", parts.ToArray());
        }

        private static string DescribeReadOnlyBoundRow(DataGridViewRow row, HashSet<string> quotaCodes)
        {
            DataRowView rowView = row.DataBoundItem as DataRowView;
            if (rowView == null || rowView.Row == null || rowView.Row.Table == null)
            {
                return row.DataBoundItem == null ? "" : "boundType=" + row.DataBoundItem.GetType().FullName;
            }

            return BuildReadOnlyDataRowText(rowView.Row, quotaCodes);
        }

        private static string BuildReadOnlyDataRowText(DataRow row, HashSet<string> quotaCodes)
        {
            List<string> parts = new List<string>();
            parts.Add("table=" + row.Table.TableName);
            parts.Add("state=" + row.RowState.ToString());

            foreach (DataColumn column in row.Table.Columns)
            {
                string text = "";
                try
                {
                    text = Convert.ToString(row[column], CultureInfo.CurrentCulture);
                }
                catch
                {
                }

                if (String.IsNullOrWhiteSpace(text))
                {
                    continue;
                }

                AddReadOnlyQuotaCodes(quotaCodes, text);
                parts.Add(column.ColumnName + "=" + SafeReadOnlyText(text, 240));
                if (parts.Count >= ReadOnlyCalcLearningMaxColumnsPerRow)
                {
                    parts.Add("truncated=1");
                    break;
                }
            }

            return String.Join(";", parts.ToArray());
        }

        private static string ReadOnlyColumnLabel(DataGridViewColumn column)
        {
            if (!String.IsNullOrEmpty(column.DataPropertyName))
            {
                return column.DataPropertyName;
            }
            if (!String.IsNullOrEmpty(column.HeaderText))
            {
                return column.HeaderText;
            }
            return column.Name;
        }

        private static string DescribeReadOnlyDataSource(object dataSource)
        {
            if (dataSource == null)
            {
                return "";
            }

            DataTable table = dataSource as DataTable;
            if (table != null)
            {
                return "DataTable:" + table.TableName;
            }

            BindingSource bindingSource = dataSource as BindingSource;
            if (bindingSource != null)
            {
                return "BindingSource:" + DescribeReadOnlyDataSource(bindingSource.DataSource);
            }

            return dataSource.GetType().FullName;
        }

        private static string GetReadOnlyControlPath(Control control)
        {
            List<string> parts = new List<string>();
            for (Control current = control; current != null; current = current.Parent)
            {
                string name = String.IsNullOrEmpty(current.Name) ? current.GetType().Name : current.Name;
                parts.Add(name);
                if (current is Form)
                {
                    break;
                }
            }

            parts.Reverse();
            return String.Join("/", parts.ToArray());
        }

        private static void AddReadOnlyQuotaCodes(HashSet<string> quotaCodes, string text)
        {
            if (quotaCodes == null || String.IsNullOrWhiteSpace(text))
            {
                return;
            }

            foreach (Match match in Regex.Matches(text.ToUpperInvariant(), @"\b[A-Z]{1,3}-\d+[A-Z]?\b"))
            {
                quotaCodes.Add(match.Value);
            }
        }

        private static void LogReadOnlyProjectRowsForCodes(Form mainForm, HashSet<string> quotaCodes)
        {
            if (quotaCodes == null || quotaCodes.Count == 0)
            {
                return;
            }

            DataSet projectDataSet = GetField<DataSet>(mainForm, "m_DataSet");
            if (projectDataSet == null)
            {
                return;
            }

            foreach (DataTable table in projectDataSet.Tables)
            {
                for (int i = 0; i < table.Rows.Count; i++)
                {
                    DataRow row = table.Rows[i];
                    string matchedCode;
                    if (!ReadOnlyDataRowContainsQuotaCode(row, quotaCodes, out matchedCode))
                    {
                        continue;
                    }

                    string rowText = BuildReadOnlyDataRowText(row, quotaCodes);
                    string key = "project:" + table.TableName + ":" + i.ToString(CultureInfo.InvariantCulture) + ":" + matchedCode;
                    string oldSignature;
                    if (ReadOnlyCalcLearningSignatures.TryGetValue(key, out oldSignature) &&
                        String.Equals(oldSignature, rowText, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    ReadOnlyCalcLearningSignatures[key] = rowText;
                    Log("ReadOnly calc project row: key=" + key + " {" + rowText + "}");
                }
            }
        }

        private static bool ReadOnlyDataRowContainsQuotaCode(
            DataRow row,
            HashSet<string> quotaCodes,
            out string matchedCode)
        {
            matchedCode = "";
            if (row == null || row.Table == null)
            {
                return false;
            }

            foreach (DataColumn column in row.Table.Columns)
            {
                string text = "";
                try
                {
                    text = Convert.ToString(row[column], CultureInfo.CurrentCulture);
                }
                catch
                {
                }

                if (String.IsNullOrEmpty(text))
                {
                    continue;
                }

                foreach (string code in quotaCodes)
                {
                    if (text.IndexOf(code, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        matchedCode = code;
                        return true;
                    }
                }
            }

            return false;
        }

        private static string SafeReadOnlyText(string text, int maxLength)
        {
            if (String.IsNullOrEmpty(text))
            {
                return "";
            }

            string normalized = text.Replace("\r", " ").Replace("\n", " ").Replace("\t", " ");
            if (normalized.Length <= maxLength)
            {
                return normalized;
            }

            return normalized.Substring(0, maxLength) + "...";
        }

        private static void StartInputCompatibilityDiagnostics()
        {
            if (inputCompatibilityTimer != null)
            {
                return;
            }

            inputCompatibilityTimer = new System.Windows.Forms.Timer();
            inputCompatibilityTimer.Interval = 1000;
            inputCompatibilityTimer.Tick += delegate
            {
                try
                {
                    InstallFindDeCompatibilityHook();
                    Form mainForm = FindMainForm();
                    if (mainForm != null && !DiagnosedInputForms.Contains(mainForm))
                    {
                        DiagnoseQuotaInputHandlers(mainForm);
                        DiagnosedInputForms.Add(mainForm);
                    }
                    if (mainForm != null && !DiagnosedDataSetForms.Contains(mainForm) &&
                        LogQuotaDataSets(mainForm))
                    {
                        DiagnosedDataSetForms.Add(mainForm);
                    }

                    LogNewValidationDialogs(mainForm);
                }
                catch (Exception ex)
                {
                    Log("Input compatibility diagnostics failed: " + ex);
                }
            };
            inputCompatibilityTimer.Start();
            Log("Input compatibility diagnostics started.");
            if (Is2024Process())
            {
                ThreadPool.QueueUserWorkItem(delegate { EnsureMigratedQuotaCodesLoaded(); });
            }
        }

        private static void DiagnoseQuotaInputHandlers(Form mainForm)
        {
            DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
            if (grid == null)
            {
                Log("Input diagnostics: dataGridViewDE not found.");
                return;
            }

            Log("Input diagnostics: form=" + mainForm.GetType().FullName
                + " grid=" + grid.GetType().FullName
                + " columns=" + grid.Columns.Count.ToString(CultureInfo.InvariantCulture));
            foreach (DataGridViewColumn column in grid.Columns)
            {
                Log("Input column: index=" + column.Index.ToString(CultureInfo.InvariantCulture)
                    + " name=" + column.Name
                    + " header=" + column.HeaderText
                    + " property=" + column.DataPropertyName);
            }

            PropertyInfo eventsProperty = typeof(Component).GetProperty("Events", BindingFlags.Instance | BindingFlags.NonPublic);
            EventHandlerList events = eventsProperty == null ? null : eventsProperty.GetValue(grid, null) as EventHandlerList;
            if (events == null)
            {
                Log("Input diagnostics: EventHandlerList unavailable.");
                return;
            }

            bool wrapHandlers = !WrappedInputGrids.Contains(grid);
            for (Type type = grid.GetType(); type != null; type = type.BaseType)
            {
                foreach (FieldInfo field in type.GetFields(BindingFlags.Static | BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.DeclaredOnly))
                {
                    if (field.FieldType != typeof(object) ||
                        field.Name.IndexOf("EVENT", StringComparison.OrdinalIgnoreCase) < 0)
                    {
                        continue;
                    }

                    object key;
                    try
                    {
                        key = field.GetValue(null);
                    }
                    catch
                    {
                        continue;
                    }

                    Delegate handler = events[key];
                    if (handler == null)
                    {
                        continue;
                    }

                    foreach (Delegate item in handler.GetInvocationList())
                    {
                        Log("Input event: " + type.FullName + "." + field.Name
                            + " -> " + item.Method.DeclaringType.FullName + "." + item.Method.Name
                            + " target=" + (item.Target == null ? "(static)" : item.Target.GetType().FullName));
                    }

                    if (wrapHandlers)
                    {
                        WrapInputEventHandler(events, key, field.Name, handler, grid);
                    }
                }
            }

            if (wrapHandlers)
            {
                WrappedInputGrids.Add(grid);
                Log("Input diagnostics: key handlers wrapped.");
            }
        }

        private static bool LogQuotaDataSets(Form mainForm)
        {
            try
            {
                FieldInfo commonField = mainForm.GetType().GetField(
                    "m_CmmonDataSet",
                    BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
                object common = commonField == null ? null : commonField.GetValue(null);
                FieldInfo commonDataSetField = common == null ? null : common.GetType().GetField(
                    "m_DataSet",
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                DataSet commonDataSet = commonDataSetField == null
                    ? null
                    : commonDataSetField.GetValue(common) as DataSet;
                FieldInfo projectField = mainForm.GetType().GetField(
                    "m_DataSet",
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                DataSet projectDataSet = projectField == null
                    ? null
                    : projectField.GetValue(mainForm) as DataSet;
                if (commonDataSet == null || projectDataSet == null)
                {
                    return false;
                }

                LogDataSetSummary("common", commonDataSet);
                LogDataSetSummary("project", projectDataSet);
                return true;
            }
            catch (Exception ex)
            {
                Log("Input dataset diagnostics failed: " + ex);
                return false;
            }
        }

        private static void LogDataSetSummary(string label, DataSet dataSet)
        {
            if (dataSet == null)
            {
                Log("Input dataset: " + label + " unavailable.");
                return;
            }

            Log("Input dataset: " + label
                + " name=" + dataSet.DataSetName
                + " tables=" + dataSet.Tables.Count.ToString(CultureInfo.InvariantCulture));
            foreach (DataTable table in dataSet.Tables)
            {
                string[] columns = table.Columns.Cast<DataColumn>()
                    .Select(x => x.ColumnName)
                    .ToArray();
                int lgBookMatches = CountTableMatches(table, "LG_2018");
                int lgCodeMatches = CountTableMatches(table, "LG-12");
                Log("Input dataset table: " + label
                    + " name=" + table.TableName
                    + " rows=" + table.Rows.Count.ToString(CultureInfo.InvariantCulture)
                    + " cols=" + String.Join("|", columns)
                    + " LG_2018=" + lgBookMatches.ToString(CultureInfo.InvariantCulture)
                    + " LG-12=" + lgCodeMatches.ToString(CultureInfo.InvariantCulture));
                if (lgBookMatches > 0 || String.Equals(table.TableName, "定额库索引", StringComparison.Ordinal))
                {
                    LogMatchingRows(label, table, new[] { "LG_2018", "DY_2024", "LY_2024" });
                }
            }
        }

        private static void LogMatchingRows(string label, DataTable table, string[] expectedValues)
        {
            foreach (DataRow row in table.Rows)
            {
                bool matched = false;
                foreach (DataColumn column in table.Columns)
                {
                    string value = Convert.ToString(row[column], CultureInfo.CurrentCulture).Trim();
                    if (expectedValues.Any(x => String.Equals(x, value, StringComparison.OrdinalIgnoreCase)))
                    {
                        matched = true;
                        break;
                    }
                }
                if (!matched)
                {
                    continue;
                }

                List<string> values = new List<string>();
                foreach (DataColumn column in table.Columns)
                {
                    values.Add(column.ColumnName + "="
                        + Convert.ToString(row[column], CultureInfo.CurrentCulture));
                }
                Log("Input dataset row: " + label
                    + " table=" + table.TableName
                    + " " + String.Join(";", values.ToArray()));
            }
        }

        private static int CountTableMatches(DataTable table, string expected)
        {
            int count = 0;
            foreach (DataRow row in table.Rows)
            {
                foreach (DataColumn column in table.Columns)
                {
                    if (String.Equals(
                        Convert.ToString(row[column], CultureInfo.CurrentCulture).Trim(),
                        expected,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        count++;
                        break;
                    }
                }
            }
            return count;
        }

        private static void WrapInputEventHandler(
            EventHandlerList events,
            object key,
            string eventKeyName,
            Delegate handler,
            DataGridView grid)
        {
            if (String.Equals(eventKeyName, "EVENT_DATAGRIDVIEWCELLVALIDATING", StringComparison.OrdinalIgnoreCase) &&
                handler is DataGridViewCellValidatingEventHandler)
            {
                DataGridViewCellValidatingEventHandler original =
                    (DataGridViewCellValidatingEventHandler)handler;
                DataGridViewCellValidatingEventHandler wrapper = delegate(object sender, DataGridViewCellValidatingEventArgs e)
                {
                    string formattedValue = Convert.ToString(e.FormattedValue, CultureInfo.CurrentCulture);
                    LogInputEventState(grid, eventKeyName + " before", e.RowIndex, e.ColumnIndex,
                        "formatted=" + formattedValue
                        + " cancel=" + e.Cancel.ToString());
                    if (HasMigratedQuotaPrefix(formattedValue))
                    {
                        ProbeMigratedQuotaLookup(grid.FindForm(), formattedValue);
                    }
                    original(sender, e);
                    LogInputEventState(grid, eventKeyName + " after", e.RowIndex, e.ColumnIndex,
                        "formatted=" + formattedValue
                        + " cancel=" + e.Cancel.ToString());
                };
                events.RemoveHandler(key, handler);
                events.AddHandler(key, wrapper);
                return;
            }

            if (String.Equals(eventKeyName, "EVENT_DATAGRIDVIEWCELLVALUECHANGED", StringComparison.OrdinalIgnoreCase) &&
                handler is DataGridViewCellEventHandler)
            {
                DataGridViewCellEventHandler original = (DataGridViewCellEventHandler)handler;
                DataGridViewCellEventHandler wrapper = delegate(object sender, DataGridViewCellEventArgs e)
                {
                    LogInputEventState(grid, eventKeyName + " before", e.RowIndex, e.ColumnIndex, null);
                    original(sender, e);
                    LogInputEventState(grid, eventKeyName + " after", e.RowIndex, e.ColumnIndex, null);
                };
                events.RemoveHandler(key, handler);
                events.AddHandler(key, wrapper);
                return;
            }

            if (String.Equals(eventKeyName, "EVENT_DATAGRIDVIEWDATAERROR", StringComparison.OrdinalIgnoreCase) &&
                handler is DataGridViewDataErrorEventHandler)
            {
                DataGridViewDataErrorEventHandler original = (DataGridViewDataErrorEventHandler)handler;
                DataGridViewDataErrorEventHandler wrapper = delegate(object sender, DataGridViewDataErrorEventArgs e)
                {
                    LogInputEventState(grid, eventKeyName + " before", e.RowIndex, e.ColumnIndex,
                        "context=" + e.Context.ToString()
                        + " cancel=" + e.Cancel.ToString()
                        + " throw=" + e.ThrowException.ToString()
                        + " exception=" + (e.Exception == null ? "" : e.Exception.Message));
                    original(sender, e);
                    LogInputEventState(grid, eventKeyName + " after", e.RowIndex, e.ColumnIndex,
                        "cancel=" + e.Cancel.ToString()
                        + " throw=" + e.ThrowException.ToString());
                };
                events.RemoveHandler(key, handler);
                events.AddHandler(key, wrapper);
                return;
            }

            if (String.Equals(eventKeyName, "EventKeyDown", StringComparison.OrdinalIgnoreCase) &&
                handler is KeyEventHandler)
            {
                KeyEventHandler original = (KeyEventHandler)handler;
                KeyEventHandler wrapper = delegate(object sender, KeyEventArgs e)
                {
                    LogInputEventState(grid, eventKeyName + " before", -1, -1,
                        "key=" + e.KeyCode.ToString()
                        + " handled=" + e.Handled.ToString()
                        + " suppress=" + e.SuppressKeyPress.ToString());
                    original(sender, e);
                    LogInputEventState(grid, eventKeyName + " after", -1, -1,
                        "key=" + e.KeyCode.ToString()
                        + " handled=" + e.Handled.ToString()
                        + " suppress=" + e.SuppressKeyPress.ToString());
                };
                events.RemoveHandler(key, handler);
                events.AddHandler(key, wrapper);
                return;
            }

            if (String.Equals(eventKeyName, "EventValidated", StringComparison.OrdinalIgnoreCase) &&
                handler is EventHandler)
            {
                EventHandler original = (EventHandler)handler;
                EventHandler wrapper = delegate(object sender, EventArgs e)
                {
                    LogInputEventState(grid, eventKeyName + " before", -1, -1, null);
                    original(sender, e);
                    LogInputEventState(grid, eventKeyName + " after", -1, -1, null);
                };
                events.RemoveHandler(key, handler);
                events.AddHandler(key, wrapper);
            }
        }

        private static void LogInputEventState(
            DataGridView grid,
            string stage,
            int rowIndex,
            int columnIndex,
            string extra)
        {
            try
            {
                DataGridViewCell cell = null;
                if (rowIndex >= 0 && rowIndex < grid.Rows.Count &&
                    columnIndex >= 0 && columnIndex < grid.Columns.Count)
                {
                    cell = grid.Rows[rowIndex].Cells[columnIndex];
                }
                else
                {
                    cell = grid.CurrentCell;
                }

                object dataBoundItem = null;
                if (cell != null && cell.RowIndex >= 0 && cell.RowIndex < grid.Rows.Count)
                {
                    dataBoundItem = grid.Rows[cell.RowIndex].DataBoundItem;
                }

                Log("Input trace: " + stage
                    + " cell=" + (cell == null ? "(null)" :
                        cell.RowIndex.ToString(CultureInfo.InvariantCulture) + ","
                        + cell.ColumnIndex.ToString(CultureInfo.InvariantCulture))
                    + " column=" + (cell == null ? "" : cell.OwningColumn.Name)
                    + " value=" + (cell == null ? "" : Convert.ToString(cell.Value, CultureInfo.CurrentCulture))
                    + " edited=" + (grid.EditingControl == null ? "" : grid.EditingControl.Text)
                    + " dirty=" + grid.IsCurrentCellDirty.ToString()
                    + " boundType=" + (dataBoundItem == null ? "(null)" : dataBoundItem.GetType().FullName)
                    + (String.IsNullOrEmpty(extra) ? "" : " " + extra));
            }
            catch (Exception ex)
            {
                Log("Input trace failed at " + stage + ": " + ex.Message);
            }
        }

        private static void ScheduleMigratedQuotaRowRepair(DataGridView grid, int rowIndex, string reason)
        {
            if (!IsMigratedQuotaRowRepairEnabled())
            {
                return;
            }

            if (repairingMigratedQuotaRow || !Is2024Process() || grid == null ||
                rowIndex < 0 || rowIndex >= grid.Rows.Count || !grid.IsHandleCreated ||
                grid.IsDisposed)
            {
                return;
            }

            try
            {
                grid.BeginInvoke((MethodInvoker)delegate
                {
                    RepairMigratedQuotaRow(grid, rowIndex, reason);
                });
            }
            catch (Exception ex)
            {
                Log("Input compatibility row repair schedule failed: " + ex.Message);
            }
        }

        private static void RepairMigratedQuotaRow(DataGridView grid, int rowIndex, string reason)
        {
            if (!IsMigratedQuotaRowRepairEnabled())
            {
                return;
            }

            if (repairingMigratedQuotaRow || !Is2024Process() || grid == null ||
                rowIndex < 0 || rowIndex >= grid.Rows.Count)
            {
                return;
            }

            DataGridViewRow gridRow = grid.Rows[rowIndex];
            DataRow dataRow = null;
            DataRowView rowView = gridRow.DataBoundItem as DataRowView;
            if (rowView != null)
            {
                dataRow = rowView.Row;
            }

            string code = ReadQuotaRowText(dataRow, gridRow, "\u5b9a\u989d\u7f16\u53f7");
            MigratedQuotaRecord record;
            if (!TryGetMigratedQuotaRecord(code, out record))
            {
                return;
            }

            try
            {
                repairingMigratedQuotaRow = true;

                double quantity;
                bool hasQuantity = TryReadQuotaRowDouble(dataRow, gridRow, "\u5de5\u7a0b\u6570\u91cf", out quantity);
                if ((!hasQuantity || quantity == 0d) &&
                    TryReadQuotaRowDouble(dataRow, gridRow, "\u5de5\u7a0b\u6570\u91cf\u8f93\u5165", out quantity))
                {
                    hasQuantity = true;
                }
                if (!hasQuantity)
                {
                    quantity = 0d;
                }

                double total = hasQuantity ? quantity * record.BasePrice : 0d;
                string consume = record.PlainConsume;
                if (String.IsNullOrEmpty(consume) && !String.IsNullOrEmpty(record.EncryptedConsume))
                {
                    consume = DecryptConsumeForCurrentApp(record.EncryptedConsume);
                    record.PlainConsume = consume;
                }

                SetQuotaRowValue(dataRow, gridRow, "\u5b9a\u989d\u7f16\u53f7", record.Code);
                SetQuotaRowValue(dataRow, gridRow, "\u5de5\u7a0b\u6216\u8d39\u7528\u9879\u76ee\u540d\u79f0", record.Name);
                SetQuotaRowValue(dataRow, gridRow, "\u5355\u4f4d", record.Unit);
                if (hasQuantity)
                {
                    SetQuotaRowValue(dataRow, gridRow, "\u5de5\u7a0b\u6570\u91cf", quantity);
                }
                SetQuotaRowValue(dataRow, gridRow, "\u5355\u4ef7", record.BasePrice);
                if (hasQuantity)
                {
                    SetQuotaRowValue(dataRow, gridRow, "\u5408\u4ef7", total);
                }
                SetQuotaRowValue(dataRow, gridRow, "\u5355\u91cd", record.Weight);
                SetQuotaRowValue(dataRow, gridRow, "\u5b9a\u989d\u7efc\u5408\u5355\u4ef7", record.BasePrice);
                if (hasQuantity)
                {
                    SetQuotaRowValue(dataRow, gridRow, "\u5b9a\u989d\u7efc\u5408\u5408\u4ef7", total);
                }
                SetQuotaRowValue(dataRow, gridRow, "\u57fa\u672c\u5b9a\u989d", record.BasicQuota ?? "");
                if (!String.IsNullOrEmpty(consume))
                {
                    SetQuotaRowValue(dataRow, gridRow, "\u5b9a\u989d\u6d88\u8017", consume);
                }

                Log("Input compatibility row repaired: code=" + record.Code
                    + " reason=" + reason
                    + " quantity=" + quantity.ToString(CultureInfo.InvariantCulture)
                    + " price=" + record.BasePrice.ToString(CultureInfo.InvariantCulture)
                    + " total=" + total.ToString(CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                Log("Input compatibility row repair failed: " + ex);
            }
            finally
            {
                repairingMigratedQuotaRow = false;
            }
        }

        private static bool IsMigratedQuotaRowRepairEnabled()
        {
            return false;
        }

        private static string ReadQuotaRowText(DataRow dataRow, DataGridViewRow gridRow, string columnName)
        {
            object value;
            if (TryGetQuotaRowValue(dataRow, gridRow, columnName, out value))
            {
                return Convert.ToString(value, CultureInfo.CurrentCulture);
            }
            return "";
        }

        private static bool TryReadQuotaRowDouble(
            DataRow dataRow,
            DataGridViewRow gridRow,
            string columnName,
            out double value)
        {
            object raw;
            if (!TryGetQuotaRowValue(dataRow, gridRow, columnName, out raw))
            {
                value = 0d;
                return false;
            }

            if (raw is double)
            {
                value = (double)raw;
                return true;
            }
            if (raw is decimal)
            {
                value = Convert.ToDouble((decimal)raw, CultureInfo.InvariantCulture);
                return true;
            }

            string text = Convert.ToString(raw, CultureInfo.CurrentCulture).Trim();
            return Double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value) ||
                Double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        private static bool TryGetQuotaRowValue(
            DataRow dataRow,
            DataGridViewRow gridRow,
            string columnName,
            out object value)
        {
            if (dataRow != null && dataRow.Table != null &&
                dataRow.Table.Columns.Contains(columnName))
            {
                value = dataRow[columnName];
                return true;
            }

            if (gridRow != null && gridRow.DataGridView != null)
            {
                foreach (DataGridViewColumn column in gridRow.DataGridView.Columns)
                {
                    if (String.Equals(column.DataPropertyName, columnName, StringComparison.Ordinal) ||
                        String.Equals(column.Name, columnName, StringComparison.Ordinal) ||
                        String.Equals(column.HeaderText, columnName, StringComparison.Ordinal))
                    {
                        value = gridRow.Cells[column.Index].Value;
                        return true;
                    }
                }
            }

            value = null;
            return false;
        }

        private static void SetQuotaRowValue(
            DataRow dataRow,
            DataGridViewRow gridRow,
            string columnName,
            object value)
        {
            try
            {
                if (dataRow != null && dataRow.Table != null &&
                    dataRow.Table.Columns.Contains(columnName))
                {
                    if (QuotaRowValuesEqual(dataRow[columnName], value))
                    {
                        return;
                    }
                    dataRow[columnName] = value ?? DBNull.Value;
                    return;
                }

                if (gridRow != null && gridRow.DataGridView != null)
                {
                    foreach (DataGridViewColumn column in gridRow.DataGridView.Columns)
                    {
                        if (String.Equals(column.DataPropertyName, columnName, StringComparison.Ordinal) ||
                            String.Equals(column.Name, columnName, StringComparison.Ordinal) ||
                            String.Equals(column.HeaderText, columnName, StringComparison.Ordinal))
                        {
                            if (QuotaRowValuesEqual(gridRow.Cells[column.Index].Value, value))
                            {
                                return;
                            }
                            gridRow.Cells[column.Index].Value = value;
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("Input compatibility row set failed: column=" + columnName + " " + ex.Message);
            }
        }

        private static bool QuotaRowValuesEqual(object current, object next)
        {
            if (current == null || current == DBNull.Value)
            {
                return next == null || next == DBNull.Value ||
                    String.IsNullOrEmpty(Convert.ToString(next, CultureInfo.CurrentCulture));
            }
            if (next == null || next == DBNull.Value)
            {
                return String.IsNullOrEmpty(Convert.ToString(current, CultureInfo.CurrentCulture));
            }

            double currentDouble;
            double nextDouble;
            if (TryConvertToDouble(current, out currentDouble) &&
                TryConvertToDouble(next, out nextDouble))
            {
                return Math.Abs(currentDouble - nextDouble) < 0.000001d;
            }

            return String.Equals(
                Convert.ToString(current, CultureInfo.CurrentCulture),
                Convert.ToString(next, CultureInfo.CurrentCulture),
                StringComparison.Ordinal);
        }

        private static bool TryConvertToDouble(object value, out double result)
        {
            if (value is double)
            {
                result = (double)value;
                return true;
            }
            if (value is float || value is decimal || value is int || value is long ||
                value is short || value is byte)
            {
                result = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                return true;
            }

            string text = Convert.ToString(value, CultureInfo.CurrentCulture);
            return Double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out result) ||
                Double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out result);
        }

        private static void ProbeMigratedQuotaLookup(Form mainForm, string code)
        {
            try
            {
                if (mainForm == null)
                {
                    Log("Input lookup probe: main form unavailable.");
                    return;
                }

                FieldInfo deField = mainForm.GetType().GetField(
                    "de",
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                object current = deField == null ? null : deField.GetValue(mainForm);
                if (current == null)
                {
                    Log("Input lookup probe: DEBase instance unavailable.");
                    return;
                }

                Type type = current.GetType();
                MethodInfo method = type.GetMethod(
                    "FindDe",
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                    null,
                    new[] { typeof(string), typeof(bool) },
                    null);
                if (method == null)
                {
                    Log("Input lookup probe: FindDe(string,bool) not found.");
                    return;
                }

                foreach (bool flag in new[] { false, true })
                {
                    object probe = Activator.CreateInstance(type, true);
                    CopyInstanceFields(current, probe);
                    bool result = Convert.ToBoolean(
                        method.Invoke(probe, new object[] { code, flag }),
                        CultureInfo.InvariantCulture);
                    Log("Input lookup probe: overload=FindDe(string,bool)"
                        + " flag=" + flag.ToString()
                        + " result=" + result.ToString()
                        + " state=" + DescribeSimpleFields(probe));
                }

                List<string> candidates = new List<string>();
                foreach (string name in new[] { "gslb", "strSelectTMBH", "strSelTMBH" })
                {
                    FieldInfo field = mainForm.GetType().GetField(
                        name,
                        BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
                    string value = field == null ? "" : Convert.ToString(
                        field.GetValue(field.IsStatic ? null : mainForm),
                        CultureInfo.CurrentCulture);
                    if (!String.IsNullOrWhiteSpace(value) && !candidates.Contains(value))
                    {
                        candidates.Add(value);
                    }
                    Log("Input lookup context: " + name + "=" + value);
                }

                MethodInfo overload = type.GetMethod(
                    "FindDe",
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                    null,
                    new[] { typeof(string), typeof(string), typeof(bool) },
                    null);
                if (overload != null)
                {
                    foreach (string candidate in candidates)
                    {
                        foreach (bool flag in new[] { false, true })
                        {
                            object probe = Activator.CreateInstance(type, true);
                            CopyInstanceFields(current, probe);
                            bool result = Convert.ToBoolean(
                                overload.Invoke(probe, new object[] { code, candidate, flag }),
                                CultureInfo.InvariantCulture);
                            Log("Input lookup probe: overload=FindDe(string,string,bool)"
                                + " value=" + candidate
                                + " flag=" + flag.ToString()
                                + " result=" + result.ToString()
                                + " state=" + DescribeSimpleFields(probe));
                        }
                    }
                }
            }
            catch (TargetInvocationException ex)
            {
                Log("Input lookup probe failed: "
                    + (ex.InnerException == null ? ex.ToString() : ex.InnerException.ToString()));
            }
            catch (Exception ex)
            {
                Log("Input lookup probe failed: " + ex);
            }
        }

        private static void CopyInstanceFields(object source, object target)
        {
            for (Type type = source.GetType(); type != null; type = type.BaseType)
            {
                foreach (FieldInfo field in type.GetFields(
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.DeclaredOnly))
                {
                    if (!field.IsInitOnly)
                    {
                        field.SetValue(target, field.GetValue(source));
                    }
                }
            }
        }

        private static string DescribeSimpleFields(object value)
        {
            List<string> parts = new List<string>();
            for (Type type = value.GetType(); type != null; type = type.BaseType)
            {
                foreach (FieldInfo field in type.GetFields(
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.DeclaredOnly))
                {
                    Type fieldType = field.FieldType;
                    if (fieldType != typeof(string) && fieldType != typeof(bool) &&
                        fieldType != typeof(int) && fieldType != typeof(long) &&
                        fieldType != typeof(double) && fieldType != typeof(decimal))
                    {
                        continue;
                    }

                    parts.Add(field.Name + "="
                        + Convert.ToString(field.GetValue(value), CultureInfo.CurrentCulture));
                }
            }
            return String.Join(";", parts.ToArray());
        }

        private static void InstallFindDeCompatibilityHook()
        {
            if (!Is2024Process() || findDeCompatibilityHookInstalled)
            {
                return;
            }

            try
            {
                Type deBaseType = null;
                foreach (Assembly assembly in AppDomain.CurrentDomain.GetAssemblies())
                {
                    deBaseType = assembly.GetType("RecoNet.DEBase", false);
                    if (deBaseType != null)
                    {
                        break;
                    }
                }
                if (deBaseType == null)
                {
                    if (!findDeCompatibilityTypeMissingLogged)
                    {
                        findDeCompatibilityTypeMissingLogged = true;
                        Log("Input compatibility hook: RecoNet.DEBase not found; waiting for runtime load.");
                    }
                    return;
                }

                MethodInfo original = deBaseType.GetMethod(
                    "FindDe",
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                    null,
                    new[] { typeof(string), typeof(bool) },
                    null);
                MethodInfo prefix = typeof(QuotaRecommendPanel).GetMethod(
                    "FindDeCompatibilityPrefix",
                    BindingFlags.Static | BindingFlags.NonPublic);
                MethodInfo originalWithContext = deBaseType.GetMethod(
                    "FindDe",
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                    null,
                    new[] { typeof(string), typeof(string), typeof(bool) },
                    null);
                MethodInfo prefixWithContext = typeof(QuotaRecommendPanel).GetMethod(
                    "FindDeCompatibilityWithContextPrefix",
                    BindingFlags.Static | BindingFlags.NonPublic);
                if (original == null || prefix == null)
                {
                    Log("Input compatibility hook: method not found.");
                    return;
                }

                findDeCompatibilityHarmony = new Harmony("reco.quota.compatibility.2020-estimate");
                findDeCompatibilityHarmony.Patch(original, new HarmonyMethod(prefix));
                if (originalWithContext != null && prefixWithContext != null)
                {
                    findDeCompatibilityHarmony.Patch(originalWithContext, new HarmonyMethod(prefixWithContext));
                }
                findDeCompatibilityHookInstalled = true;
                Log("Input compatibility hook installed: RecoNet.DEBase.FindDe migrated quota lookup only."
                    + " withContext=" + (originalWithContext != null && prefixWithContext != null).ToString());
            }
            catch (Exception ex)
            {
                Log("Input compatibility hook failed: " + ex);
            }
        }

        private static bool FindDeCompatibilityPrefix(
            object __instance,
            string __0,
            bool __1,
            ref bool __result)
        {
            return ApplyFindDeCompatibility(__instance, __0, __1, ref __result);
        }

        private static bool FindDeCompatibilityWithContextPrefix(
            object __instance,
            string __0,
            string __1,
            bool __2,
            ref bool __result)
        {
            return ApplyFindDeCompatibility(__instance, __0, __2, ref __result);
        }

        private static bool ApplyFindDeCompatibility(
            object instance,
            string code,
            bool decryptConsume,
            ref bool result)
        {
            try
            {
                MigratedQuotaRecord record;
                if (!TryGetMigratedQuotaRecord(code, out record))
                {
                    return true;
                }

                string consume = record.EncryptedConsume ?? "";
                if (decryptConsume)
                {
                    if (record.PlainConsume == null)
                    {
                        record.PlainConsume = DecryptConsumeForCurrentApp(consume);
                    }
                    consume = record.PlainConsume;
                }

                SetInstanceField(instance, "xh", consume);
                SetInstanceField(instance, "jbde", record.BasicQuota ?? "");
                SetInstanceField(instance, "tableName", record.Book ?? "");
                SetInstanceField(instance, "jNo", record.SectionNo ?? "");
                SetInstanceField(instance, "bh", record.Code ?? "");
                SetInstanceField(instance, "name", record.Name ?? "");
                SetInstanceField(instance, "dw", record.Unit ?? "");
                SetInstanceField(instance, "sl", 0d);
                SetInstanceField(instance, "tz", "");
                SetInstanceField(instance, "dj", record.BasePrice);
                SetInstanceField(instance, "dz", record.Weight);
                SetInstanceField(instance, "tmbh", "");
                SetInstanceField(instance, "mem_xh", "");
                SetInstanceField(instance, "mem_jbde", "");

                result = true;
                Log("Input compatibility FindDe: code=" + record.Code
                    + " book=" + record.Book
                    + " decrypted=" + decryptConsume.ToString());
                return false;
            }
            catch (Exception ex)
            {
                Log("Input compatibility FindDe failed: " + ex);
                return true;
            }
        }

        private static bool TryGetMigratedQuotaRecord(string code, out MigratedQuotaRecord record)
        {
            record = null;
            string normalized = NormalizeMigratedQuotaCode(code);
            if (!HasMigratedQuotaPrefix(normalized))
            {
                return false;
            }

            EnsureMigratedQuotaCodesLoaded();
            Dictionary<string, MigratedQuotaRecord> records = migratedQuotaRecords;
            return records != null && records.TryGetValue(normalized, out record);
        }

        private static void SetInstanceField(object instance, string fieldName, object value)
        {
            for (Type type = instance.GetType(); type != null; type = type.BaseType)
            {
                FieldInfo field = type.GetField(
                    fieldName,
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.DeclaredOnly);
                if (field == null)
                {
                    continue;
                }

                field.SetValue(instance, value);
                return;
            }
        }

        private static string DecryptConsumeForCurrentApp(string encrypted)
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
            MethodInfo method = securityType.GetMethod(
                "Decrypto",
                BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance,
                null,
                new[] { typeof(string) },
                null);
            if (method == null)
            {
                throw new MissingMethodException("RecoNet.Security.Decrypto(string) not found.");
            }

            return Convert.ToString(
                method.Invoke(security, new object[] { encrypted }),
                CultureInfo.InvariantCulture);
        }

        private static bool IsMigratedQuotaCode(string value)
        {
            string code = NormalizeMigratedQuotaCode(value);
            if (code.Length == 0)
            {
                return false;
            }

            EnsureMigratedQuotaCodesLoaded();
            HashSet<string> codes = migratedQuotaCodes;
            return codes != null && codes.Contains(code);
        }

        private static bool HasMigratedQuotaPrefix(string value)
        {
            string code = NormalizeMigratedQuotaCode(value);
            int separator = code.IndexOf('-');
            if (separator <= 0)
            {
                return false;
            }

            string prefix = code.Substring(0, separator);
            return new[]
            {
                "LG", "QG", "SG", "GG", "TG", "XG", "EG", "DG", "HG", "FG", "PG",
                "JG", "ZG", "TZ", "XZ", "EZ", "DZ", "HZ", "FZ", "PZ", "JZ"
            }.Contains(prefix);
        }

        private static string NormalizeMigratedQuotaCode(string value)
        {
            string code = (value ?? "").Trim().ToUpperInvariant();
            if (code.Length == 0)
            {
                return "";
            }

            code = code
                .Replace('\u2010', '-')
                .Replace('\u2011', '-')
                .Replace('\u2012', '-')
                .Replace('\u2013', '-')
                .Replace('\u2014', '-')
                .Replace('\u2212', '-')
                .Replace('\uff0d', '-')
                .Replace('\ufe63', '-');
            return Regex.Replace(code, @"\s+", "");
        }

        private static bool Is2024Process()
        {
            try
            {
                return (Process.GetCurrentProcess().ProcessName ?? "")
                    .IndexOf("2024", StringComparison.OrdinalIgnoreCase) >= 0;
            }
            catch
            {
                return false;
            }
        }

        private static void EnsureMigratedQuotaCodesLoaded()
        {
            lock (MigratedQuotaCodesLock)
            {
                if (migratedQuotaCodes != null || migratedQuotaCodesLoadAttempted)
                {
                    return;
                }

                migratedQuotaCodesLoadAttempted = true;
                try
                {
                    string server = ReadCompatibilityServer();
                    if (String.IsNullOrWhiteSpace(server))
                    {
                        server = "127.0.0.1";
                    }

                    string connectionString = "Data Source=" + server
                        + ",1433;Initial Catalog=" + ResolveCompatibilityDatabaseName()
                        + ";User ID=reco;Password=" + BuildCompatibilitySqlPassword()
                        + ";Connect Timeout=8;Encrypt=False;TrustServerCertificate=True";
                    HashSet<string> codes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    Dictionary<string, MigratedQuotaRecord> records =
                        new Dictionary<string, MigratedQuotaRecord>(StringComparer.OrdinalIgnoreCase);
                    using (System.Data.SqlClient.SqlConnection connection =
                        new System.Data.SqlClient.SqlConnection(connectionString))
                    {
                        connection.Open();
                        using (System.Data.SqlClient.SqlCommand command = connection.CreateCommand())
                        {
                            command.CommandTimeout = 30;
                            StringBuilder sql = new StringBuilder();
                            sql.Append("SELECT a.[定额编号],a.[书号],a.[定额名称],a.[单位],");
                            sql.Append("CONVERT(nvarchar(max),a.[消耗]),CONVERT(nvarchar(max),a.[基本定额]),");
                            sql.Append("ISNULL(a.[基价],0),ISNULL(a.[单重],0),ISNULL(a.[节号],''),ISNULL(a.[工作内容],'') ");
                            sql.Append("FROM dbo.[定额库] a WHERE a.[书号] IN (");
                            for (int i = 0; i < MigratedQuotaBooks.Length; i++)
                            {
                                if (i > 0)
                                {
                                    sql.Append(",");
                                }
                                string parameterName = "@book" + i.ToString(CultureInfo.InvariantCulture);
                                sql.Append(parameterName);
                                command.Parameters.AddWithValue(parameterName, MigratedQuotaBooks[i]);
                            }
                            sql.Append(")");
                            command.CommandText = sql.ToString();
                            using (System.Data.SqlClient.SqlDataReader reader = command.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    string code = Convert.ToString(reader.GetValue(0), CultureInfo.InvariantCulture);
                                    if (!String.IsNullOrWhiteSpace(code))
                                    {
                                        code = NormalizeMigratedQuotaCode(code);
                                        codes.Add(code);
                                        records[code] = new MigratedQuotaRecord
                                        {
                                            Code = code,
                                            Book = Convert.ToString(reader.GetValue(1), CultureInfo.InvariantCulture),
                                            Name = Convert.ToString(reader.GetValue(2), CultureInfo.CurrentCulture),
                                            Unit = Convert.ToString(reader.GetValue(3), CultureInfo.CurrentCulture),
                                            EncryptedConsume = Convert.ToString(reader.GetValue(4), CultureInfo.InvariantCulture),
                                            BasicQuota = Convert.ToString(reader.GetValue(5), CultureInfo.InvariantCulture),
                                            BasePrice = Convert.ToDouble(reader.GetValue(6), CultureInfo.InvariantCulture),
                                            Weight = Convert.ToDouble(reader.GetValue(7), CultureInfo.InvariantCulture),
                                            SectionNo = Convert.ToString(reader.GetValue(8), CultureInfo.InvariantCulture),
                                            WorkContent = Convert.ToString(reader.GetValue(9), CultureInfo.CurrentCulture)
                                        };
                                    }
                                }
                            }
                        }
                    }

                    migratedQuotaCodes = codes;
                    migratedQuotaRecords = records;
                    Log("Input compatibility quota cache loaded: books="
                        + MigratedQuotaBooks.Length.ToString(CultureInfo.InvariantCulture)
                        + " codes=" + codes.Count.ToString(CultureInfo.InvariantCulture));
                }
                catch (Exception ex)
                {
                    migratedQuotaCodesLoadAttempted = false;
                    Log("Input compatibility quota cache failed: " + ex);
                }
            }
        }

        private static string ReadCompatibilityServer()
        {
            try
            {
                string baseDir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location) ?? "";
                string path = Path.Combine(baseDir, "ServerSetting.xml");
                if (!File.Exists(path))
                {
                    return "";
                }

                string text = File.ReadAllText(path, Encoding.UTF8);
                Match match = Regex.Match(
                    text,
                    "<(?:ServerIP|Server)>(.*?)</(?:ServerIP|Server)>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);
                return match.Success ? match.Groups[1].Value.Trim() : "";
            }
            catch
            {
                return "";
            }
        }

        private static string ResolveCompatibilityDatabaseName()
        {
            try
            {
                string processName = Process.GetCurrentProcess().ProcessName ?? "";
                if (processName.IndexOf("2024", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return "RecoData2024";
                }
            }
            catch
            {
            }

            return "RecoData2020";
        }

        private static string BuildCompatibilitySqlPassword()
        {
            return String.Join("_", new string[] { "Des", "Reco", "2006" });
        }

        private static void LogNewValidationDialogs(Form mainForm)
        {
            foreach (Form form in Application.OpenForms)
            {
                if (form == null || LoggedValidationDialogs.Contains(form))
                {
                    continue;
                }

                string text = CollectControlText(form);
                if (text.IndexOf("定额编号无效或费用类型不匹配", StringComparison.Ordinal) < 0)
                {
                    continue;
                }

                LoggedValidationDialogs.Add(form);
                Log("Quota validation dialog: type=" + form.GetType().FullName + " text=" + text);
                DataGridView grid = mainForm == null ? null : GetField<DataGridView>(mainForm, "dataGridViewDE");
                if (grid != null)
                {
                    Log("Quota validation grid: currentCell="
                        + (grid.CurrentCell == null ? "(null)" : grid.CurrentCell.RowIndex.ToString(CultureInfo.InvariantCulture)
                            + "," + grid.CurrentCell.ColumnIndex.ToString(CultureInfo.InvariantCulture))
                        + " value=" + (grid.CurrentCell == null ? "" : Convert.ToString(grid.CurrentCell.Value, CultureInfo.CurrentCulture))
                        + " editing=" + (grid.EditingControl == null ? "(null)" : grid.EditingControl.GetType().FullName));
                }
            }
        }

        private static string CollectControlText(Control root)
        {
            List<string> values = new List<string>();
            CollectControlText(root, values);
            return String.Join(" | ", values.Where(x => !String.IsNullOrWhiteSpace(x)).Distinct().ToArray());
        }

        private static void CollectControlText(Control root, List<string> values)
        {
            if (root == null)
            {
                return;
            }

            if (!String.IsNullOrWhiteSpace(root.Text))
            {
                values.Add(root.Text.Trim());
            }

            foreach (Control child in root.Controls)
            {
                CollectControlText(child, values);
            }
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
                File.AppendAllText(path, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " " + message + Environment.NewLine);
            }
            catch
            {
            }
        }
    }

    internal sealed class RecommendDialog : Form
    {
        // 与 SearchIndexStore.Search 的入选阈值保持一致：本地索引达到该分数即可直接展示，不必等 AI。
        private const int LocalIndexDisplayScore = 55;

        private readonly Form mainForm;
        private readonly DataGridView resultGrid;
        private readonly Label statusLabel;
        private readonly ComboBox quotaCategoryCombo;
        private readonly CheckBox aiNameCheckBox;
        private readonly List<LearningRecord> records;
        private readonly SearchIndexStore searchIndex;
        private readonly MappingStore mappingStore;
        private readonly ChapterLibraryStore chapterLibrary;
        private readonly Label entryScopeLabel;
        private readonly ToolTip entryScopeTip = new ToolTip();
        private DeepSeekSettings deepSeekSettings;
        private readonly List<RecommendationRow> recommendations = new List<RecommendationRow>();
        private ExcelSelection currentSelection;
        private int aiRequestVersion;
        private EntryScope currentEntryScope;
        private string lastScopeKeyUsed;
        // 项目数据库 → 是否按本库编制办法编制（QD/其他办法的项目不做条目过滤）
        private readonly Dictionary<string, bool> projectMethodCache = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

        public RecommendDialog(Form owner, string initialQuery)
        {
            mainForm = owner;
            Text = "\u6279\u91cf\u63a8\u8350\u5b9a\u989d";
            StartPosition = FormStartPosition.CenterParent;
            Width = 1280;
            Height = 680;
            MinimizeBox = false;

            records = LearningStore.Load();
            LearningStore.BackupLearningFileIfNeeded();
            searchIndex = SearchIndexStore.LoadOrBuild();
            mappingStore = MappingStore.Load(records);
            chapterLibrary = ChapterLibraryStore.Load();
            deepSeekSettings = DeepSeekSettings.Load();

            aiNameCheckBox = new CheckBox();
            aiNameCheckBox.Text = "AI\u8bc6\u522b\u5de5\u7a0b\u91cf";
            aiNameCheckBox.Left = 12;
            aiNameCheckBox.Top = 13;
            aiNameCheckBox.Width = 150;
            aiNameCheckBox.Checked = false;
            aiNameCheckBox.CheckedChanged += delegate
            {
                if (currentSelection != null)
                {
                    FillRecommendations(currentSelection);
                }
            };

            Button clipboardButton = new Button();
            clipboardButton.Text = "\u8bfb\u53d6\u526a\u8d34\u677f\u6846\u9009";
            clipboardButton.Left = 170;
            clipboardButton.Top = 10;
            clipboardButton.Width = 140;
            clipboardButton.Click += delegate { ReadClipboardAndRecommend(); };

            Button selectAllButton = new Button();
            selectAllButton.Text = "\u5168\u9009";
            selectAllButton.Left = 318;
            selectAllButton.Top = 10;
            selectAllButton.Width = 70;
            selectAllButton.Click += delegate { SetChecked(true); };

            Button clearButton = new Button();
            clearButton.Text = "\u5168\u4e0d\u9009";
            clearButton.Left = 396;
            clearButton.Top = 10;
            clearButton.Width = 80;
            clearButton.Click += delegate { SetChecked(false); };

            Button pasteButton = new Button();
            pasteButton.Text = "\u590d\u5236\u52fe\u9009\u5185\u5bb9";
            pasteButton.Left = 484;
            pasteButton.Top = 10;
            pasteButton.Width = 140;
            pasteButton.Click += delegate { CopyCheckedForManualPaste(); };

            Button aiSettingsButton = new Button();
            aiSettingsButton.Text = "AI\u8bbe\u7f6e";
            aiSettingsButton.Left = 632;
            aiSettingsButton.Top = 10;
            aiSettingsButton.Width = 78;
            aiSettingsButton.Click += delegate { ShowDeepSeekSettings(); };

            Label categoryLabel = new Label();
            categoryLabel.Text = "\u5b9a\u989d\u7c7b\u578b";
            categoryLabel.Left = 720;
            categoryLabel.Top = 15;
            categoryLabel.Width = 58;

            quotaCategoryCombo = new ComboBox();
            quotaCategoryCombo.Left = 778;
            quotaCategoryCombo.Top = 11;
            quotaCategoryCombo.Width = 110;
            quotaCategoryCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            quotaCategoryCombo.Items.AddRange(new object[] { "\u9884\u7b97\u5b9a\u989d", "\u6982\u7b97\u5b9a\u989d", "\u4f30\u7b97\u5b9a\u989d", "\u5168\u90e8" });
            quotaCategoryCombo.SelectedIndex = 0;
            quotaCategoryCombo.SelectedIndexChanged += delegate
            {
                if (currentSelection != null)
                {
                    FillRecommendations(currentSelection);
                }
            };

            // 条目信息标签随窗口拉宽而变宽（Anchor 含 Right），并挂 ToolTip 以便窄窗时悬停查看完整内容
            entryScopeLabel = new Label();
            entryScopeLabel.Left = 894;
            entryScopeLabel.Top = 15;
            entryScopeLabel.Width = 162;
            entryScopeLabel.Height = 17;
            entryScopeLabel.AutoEllipsis = true;
            entryScopeLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            entryScopeLabel.Text = "";

            Button refreshEntryButton = new Button();
            refreshEntryButton.Text = "刷新条目";
            refreshEntryButton.Width = 66;
            refreshEntryButton.Left = 1264 - 12 - refreshEntryButton.Width; // 贴右
            refreshEntryButton.Top = 10;
            refreshEntryButton.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            refreshEntryButton.Click += delegate
            {
                if (currentSelection != null)
                {
                    FillRecommendations(currentSelection);
                }
                else
                {
                    RefreshEntryScope();
                }
            };

            resultGrid = new DataGridView();
            resultGrid.Left = 12;
            resultGrid.Top = 48;
            resultGrid.Width = 1240;
            resultGrid.Height = 555;
            resultGrid.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            resultGrid.AllowUserToAddRows = false;
            resultGrid.AllowUserToDeleteRows = false;
            resultGrid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            resultGrid.MultiSelect = true;
            resultGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            resultGrid.RowHeadersVisible = false;
            resultGrid.CellContentClick += ResultGridCellContentClick;
            resultGrid.CellEndEdit += ResultGridCellEndEdit;
            AddColumns();

            statusLabel = new Label();
            statusLabel.Left = 12;
            statusLabel.Top = 612;
            statusLabel.Width = 1240;
            statusLabel.Height = 36;
            statusLabel.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;

            Controls.Add(aiNameCheckBox);
            Controls.Add(clipboardButton);
            Controls.Add(selectAllButton);
            Controls.Add(clearButton);
            Controls.Add(pasteButton);
            Controls.Add(aiSettingsButton);
            Controls.Add(categoryLabel);
            Controls.Add(quotaCategoryCombo);
            Controls.Add(entryScopeLabel);
            Controls.Add(refreshEntryButton);
            Controls.Add(resultGrid);
            Controls.Add(statusLabel);

            RefreshEntryScope();
            ReadExcelSelectionAndRecommend();
        }

        private void AddColumns()
        {
            DataGridViewCheckBoxColumn check = new DataGridViewCheckBoxColumn();
            check.Name = "Checked";
            check.HeaderText = "\u5199\u5165";
            check.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            check.Width = 48;
            resultGrid.Columns.Add(check);
            DataGridViewButtonColumn correct = new DataGridViewButtonColumn();
            correct.Name = "Correct";
            correct.HeaderText = "\u6276\u6b63";
            correct.Text = "\u6276\u6b63";
            correct.UseColumnTextForButtonValue = false;
            correct.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
            correct.Width = 48;
            resultGrid.Columns.Add(correct);
            resultGrid.Columns.Add("QuantityName", "\u5de5\u7a0b\u91cf\u540d\u79f0");
            resultGrid.Columns.Add("QuantityUnit", "\u5355\u4f4d");
            resultGrid.Columns.Add("QuantityValue", "Excel\u5de5\u7a0b\u91cf");
            resultGrid.Columns.Add("QuotaCode", "\u63a8\u8350\u5b9a\u989d");
            resultGrid.Columns.Add("QuotaName", "\u5b9a\u989d\u540d\u79f0");
            resultGrid.Columns.Add("QuotaUnit", "\u5b9a\u989d\u5355\u4f4d");
            resultGrid.Columns.Add("QuotaQuantity", "\u5b9a\u989d\u5de5\u7a0b\u91cf");
            resultGrid.Columns.Add("SourceStatus", "\u6765\u6e90");
            resultGrid.Columns["QuantityName"].FillWeight = 180;
            resultGrid.Columns["QuantityUnit"].FillWeight = 50;
            resultGrid.Columns["QuantityValue"].FillWeight = 80;
            resultGrid.Columns["QuotaCode"].FillWeight = 80;
            resultGrid.Columns["QuotaName"].FillWeight = 210;
            resultGrid.Columns["QuotaUnit"].FillWeight = 60;
            resultGrid.Columns["QuotaQuantity"].FillWeight = 85;
            resultGrid.Columns["SourceStatus"].FillWeight = 70;

            foreach (DataGridViewColumn column in resultGrid.Columns)
            {
                if (column.Name != "Checked" && column.Name != "QuantityName")
                {
                    column.ReadOnly = true;
                }
            }
        }

        private void ReadExcelSelectionAndRecommend()
        {
            recommendations.Clear();
            resultGrid.Rows.Clear();

            ExcelSelection selection;
            string error;
            if (!TryReadActiveExcelSelection(out selection, out error))
            {
                if (TryReadClipboardSelection(out selection, out error))
                {
                    FillRecommendations(selection);
                    statusLabel.Text = "\u672a\u627e\u5230 Excel/WPS COM\uff0c\u5df2\u6539\u4e3a\u8bfb\u53d6\u526a\u8d34\u677f\u6846\u9009\u5185\u5bb9\u3002\u8bf7\u5728 Excel \u6846\u9009\u540e Ctrl+C\uff0c\u518d\u70b9\u201c\u8bfb\u53d6\u526a\u8d34\u677f\u6846\u9009\u201d\u3002";
                    return;
                }

                statusLabel.Text = error + "\u53ef\u5728 Excel \u91cc\u6846\u9009\u540e\u6309 Ctrl+C\uff0c\u518d\u70b9\u201c\u8bfb\u53d6\u526a\u8d34\u677f\u6846\u9009\u201d\u3002";
                return;
            }

            FillRecommendations(selection);
        }

        private void ReadClipboardAndRecommend()
        {
            recommendations.Clear();
            resultGrid.Rows.Clear();

            ExcelSelection selection;
            string error;
            if (!TryReadClipboardSelection(out selection, out error))
            {
                statusLabel.Text = error;
                return;
            }

            FillRecommendations(selection);
        }

        private void FillRecommendations(ExcelSelection selection)
        {
            FillRecommendations(selection, aiNameCheckBox != null && aiNameCheckBox.Checked);
        }

        // 每次推荐前重新识别当前条目：对话框是非模态复用的，用户随时会在主程序里切换条目
        internal void RefreshEntryScope()
        {
            currentEntryScope = null;
            try
            {
                if (chapterLibrary != null && !chapterLibrary.IsEmpty)
                {
                    System.Data.SqlClient.SqlConnection conn = GetField<System.Data.SqlClient.SqlConnection>(mainForm, "m_ProjectConn");
                    if (conn != null && ProjectUsesLibraryMethod(conn))
                    {
                        string projectEntryName;
                        string projectEntryCode = ResolveCurrentChapterNo(conn, out projectEntryName);
                        if (!String.IsNullOrEmpty(projectEntryCode))
                        {
                            currentEntryScope = chapterLibrary.ResolveScope(projectEntryCode, projectEntryName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("RefreshEntryScope failed: " + ex.Message);
                currentEntryScope = null;
            }

            UpdateEntryScopeLabel();
        }

        // 当前项目是否按本库的编制办法编制（QD 清单变体等其他办法 ⇒ 不做条目过滤）
        private bool ProjectUsesLibraryMethod(System.Data.SqlClient.SqlConnection conn)
        {
            string dbName;
            try
            {
                dbName = conn.Database ?? "";
            }
            catch
            {
                return false;
            }

            if (String.IsNullOrWhiteSpace(dbName))
            {
                return false;
            }

            bool cached;
            if (projectMethodCache.TryGetValue(dbName, out cached))
            {
                return cached;
            }

            bool matches = true;
            try
            {
                EnsureConnectionOpen(conn);
                using (System.Data.SqlClient.SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "select 编制办法文号 from 项目信息";
                    object result = cmd.ExecuteScalar();
                    string methodNo = result == null || result == DBNull.Value ? "" : Convert.ToString(result, CultureInfo.InvariantCulture).Trim();
                    if (!String.IsNullOrEmpty(methodNo) && !String.IsNullOrEmpty(chapterLibrary.MethodNo))
                    {
                        matches = String.Equals(NormalizeMethodNo(methodNo), NormalizeMethodNo(chapterLibrary.MethodNo), StringComparison.OrdinalIgnoreCase);
                    }
                }
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("Project method check failed (treat as match): " + ex.Message);
            }

            projectMethodCache[dbName] = matches;
            return matches;
        }

        private static string NormalizeMethodNo(string text)
        {
            return (text ?? "").Replace('—', '-').Replace('–', '-').Replace('－', '-').Replace(" ", "").Trim();
        }

        // 移植自 RecoExpandPanel MultiplierFeature.ResolveChapterNo：定额行的条目序号优先，再走属性页/树节点。
        // 同时带回条目名称，用于识别用户复制条目的来源。
        private string ResolveCurrentChapterNo(System.Data.SqlClient.SqlConnection conn, out string entryName)
        {
            entryName = "";
            DataGridView quotaGrid = GetField<DataGridView>(mainForm, "dataGridViewDE");
            if (quotaGrid != null && quotaGrid.CurrentRow != null && !quotaGrid.CurrentRow.IsNewRow)
            {
                string seq = GetRowValue(quotaGrid.CurrentRow, "条目序号");
                string fromSeq = LookupChapterNoBySeq(conn, seq, out entryName);
                if (!String.IsNullOrEmpty(fromSeq))
                {
                    return fromSeq;
                }
            }

            string fromPropGrid = ReadPropertyGridValue("条目编号");
            if (!String.IsNullOrEmpty(fromPropGrid))
            {
                entryName = ReadPropertyGridValue("工程或费用项目名称") ?? "";
                return fromPropGrid;
            }

            TreeView tree = GetField<TreeView>(mainForm, "Tv_tree");
            TreeNode node = tree != null ? tree.SelectedNode : GetField<TreeNode>(mainForm, "CurrNode");
            if (node != null)
            {
                string fromTag = TryGetTagValue(node.Tag, "条目编号");
                if (!String.IsNullOrEmpty(fromTag))
                {
                    entryName = node.Text ?? "";
                    return fromTag;
                }

                string seq = TryGetTagValue(node.Tag, "条目序号");
                if (String.IsNullOrEmpty(seq) && IsAllDigits(node.Name))
                {
                    seq = node.Name;
                }

                string fromSeq = LookupChapterNoBySeq(conn, seq, out entryName);
                if (!String.IsNullOrEmpty(fromSeq))
                {
                    if (String.IsNullOrEmpty(entryName))
                    {
                        entryName = node.Text ?? "";
                    }
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
            if (String.IsNullOrWhiteSpace(seq) || conn == null)
            {
                return null;
            }

            EnsureConnectionOpen(conn);
            using (System.Data.SqlClient.SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 条目编号, 工程或费用项目名称 from 章节表 where 条目序号=@id";
                cmd.Parameters.AddWithValue("@id", seq.Trim());
                using (System.Data.SqlClient.SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                    {
                        return null;
                    }

                    string code = reader.IsDBNull(0) ? null : Convert.ToString(reader.GetValue(0), CultureInfo.InvariantCulture).Trim();
                    entryName = reader.IsDBNull(1) ? "" : Convert.ToString(reader.GetValue(1), CultureInfo.InvariantCulture).Trim();
                    return code;
                }
            }
        }

        private static void EnsureConnectionOpen(System.Data.SqlClient.SqlConnection conn)
        {
            if (conn != null && conn.State != ConnectionState.Open)
            {
                conn.Open();
            }
        }

        private string ReadPropertyGridValue(string propertyName)
        {
            DataGridView propGrid = GetField<DataGridView>(mainForm, "dataGridViewProp");
            if (propGrid == null)
            {
                return null;
            }

            foreach (DataGridViewRow row in propGrid.Rows)
            {
                if (row.IsNewRow)
                {
                    continue;
                }

                if (String.Equals(GetRowValue(row, "属性名称"), propertyName, StringComparison.Ordinal))
                {
                    return GetRowValue(row, "数据");
                }
            }

            return null;
        }

        private static string TryGetTagValue(object source, string name)
        {
            if (source == null)
            {
                return null;
            }

            DataRowView rowView = source as DataRowView;
            if (rowView != null && rowView.DataView.Table.Columns.Contains(name))
            {
                return Convert.ToString(rowView[name], CultureInfo.InvariantCulture);
            }

            DataRow dataRow = source as DataRow;
            if (dataRow != null && dataRow.Table.Columns.Contains(name))
            {
                return Convert.ToString(dataRow[name], CultureInfo.InvariantCulture);
            }

            PropertyInfo prop = source.GetType().GetProperty(name, BindingFlags.Public | BindingFlags.Instance);
            if (prop != null)
            {
                object value = prop.GetValue(source, null);
                return value == null ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
            }

            return null;
        }

        private static bool IsAllDigits(string text)
        {
            return !String.IsNullOrEmpty(text) && text.All(Char.IsDigit);
        }

        private void UpdateEntryScopeLabel()
        {
            if (entryScopeLabel == null)
            {
                return;
            }

            if (chapterLibrary == null || chapterLibrary.IsEmpty)
            {
                entryScopeLabel.Text = "条目库未启用";
                entryScopeLabel.ForeColor = SystemColors.GrayText;
                entryScopeTip.SetToolTip(entryScopeLabel, "未找到章节条目库（chapter-entries.jsonl），按全库推荐。");
                return;
            }

            if (currentEntryScope != null && currentEntryScope.Strict)
            {
                string text = "条目:" + currentEntryScope.MatchedEntryCode + " " + (currentEntryScope.EntryName ?? "")
                    + "｜池" + currentEntryScope.PoolKeys.Count.ToString(CultureInfo.InvariantCulture) + "条 严格";
                entryScopeLabel.Text = text;
                entryScopeLabel.ForeColor = Color.FromArgb(46, 96, 49);
                string tip = text;
                if (!String.Equals(currentEntryScope.ProjectEntryCode, currentEntryScope.MatchedEntryCode, StringComparison.Ordinal))
                {
                    tip += "\r\n（当前条目 " + currentEntryScope.ProjectEntryCode + " 为新建/复制条目，采用来源条目 " + currentEntryScope.MatchedEntryCode + " 的定额池）";
                }
                entryScopeTip.SetToolTip(entryScopeLabel, tip);
            }
            else
            {
                entryScopeLabel.Text = "条目:未识别（全库推荐）";
                entryScopeLabel.ForeColor = SystemColors.GrayText;
                entryScopeTip.SetToolTip(entryScopeLabel, "未识别到当前定额行所属的小计/指标条目，按全库推荐。可在主程序选中定额行后点“刷新条目”。");
            }
        }

        private Dictionary<string, bool> SnapshotCheckStates()
        {
            Dictionary<string, bool> states = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < resultGrid.Rows.Count && i < recommendations.Count; i++)
            {
                RecommendationRow rec = recommendations[i];
                if (rec == null || String.IsNullOrWhiteSpace(rec.QuotaCode))
                {
                    continue;
                }

                object value = resultGrid.Rows[i].Cells["Checked"].Value;
                states[CheckStateKey(rec)] = value is bool && (bool)value;
            }

            return states;
        }

        private static string CheckStateKey(RecommendationRow row)
        {
            if (row == null || row.Item == null)
            {
                return "";
            }

            string nameForKey = String.IsNullOrWhiteSpace(row.Item.OriginalName) ? row.Item.Name : row.Item.OriginalName;
            string codeKey = (String.IsNullOrWhiteSpace(row.TargetKind) ? QuotaEntry.GuessKind(row.QuotaCode) : row.TargetKind)
                + ":" + (row.QuotaCode ?? "").Trim().ToUpperInvariant();
            return LearningStore.BuildQuantitySignature(nameForKey, row.Item.Unit) + "|" + codeKey;
        }

        private void FillRecommendations(ExcelSelection selection, bool normalizeNames)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();
            currentSelection = selection;
            RefreshEntryScope();
            EntryScope scope = currentEntryScope;
            lastScopeKeyUsed = scope != null && scope.Strict ? scope.Tag : "";
            // 重建前快照各行勾选状态，避免用户手动取消的勾选在刷新后又被自动勾上
            Dictionary<string, bool> priorChecks = SnapshotCheckStates();
            recommendations.Clear();
            resultGrid.Rows.Clear();
            if (normalizeNames)
            {
                // 先让 AI 判断列结构并重建多列名称；命中的条目已标记跳过润色，其余条目仍走名称润色。
                bool layoutApplied = ApplyAiColumnLayout(selection);
                NormalizeQuantityNamesWithDeepSeek(selection, !layoutApplied);
            }
            else
            {
                RestoreOriginalQuantityNames(selection);
            }
            int requestVersion = ++aiRequestVersion;
            string categoryFilter = SelectedQuotaCategory;
            RecommendationBatchStats stats = new RecommendationBatchStats();
            Dictionary<string, List<RecommendationRow>> batchCache = new Dictionary<string, List<RecommendationRow>>(StringComparer.OrdinalIgnoreCase);
            List<AiPendingRecommendation> aiPending = new List<AiPendingRecommendation>();

            resultGrid.SuspendLayout();
            try
            {
                foreach (ExcelQuantityItem item in selection.Items)
                {
                    string cacheKey = BuildBatchCacheKey(item, categoryFilter) + "|" + lastScopeKeyUsed;
                    List<RecommendationRow> itemRecommendations;
                    List<RecommendationRow> cached;
                    if (batchCache.TryGetValue(cacheKey, out cached))
                    {
                        itemRecommendations = CloneRecommendationsForItem(item, cached);
                        stats.CacheHits++;
                    }
                    else
                    {
                        itemRecommendations = BuildRecommendations(item, categoryFilter, scope, stats);
                        batchCache[cacheKey] = CloneRecommendationsForItem(item, itemRecommendations);
                    }

                    int itemRowIndex = 0;
                    foreach (RecommendationRow recommendation in itemRecommendations)
                    {
                        bool isContinuation = itemRowIndex > 0 && String.Equals(recommendation.Source, "mapping", StringComparison.OrdinalIgnoreCase);
                        recommendations.Add(recommendation);
                        bool defaultChecked = recommendation.Score >= 60 && !String.IsNullOrWhiteSpace(recommendation.QuotaCode);
                        bool priorChecked;
                        bool checkedValue = priorChecks.TryGetValue(CheckStateKey(recommendation), out priorChecked) ? priorChecked : defaultChecked;
                        int gridRowIndex = resultGrid.Rows.Add(
                            checkedValue,
                            isContinuation ? "" : "\u6276\u6b63",
                            isContinuation ? "" : item.Name,
                            item.Unit,
                            item.ValueText,
                            recommendation.QuotaCode,
                            recommendation.QuotaName,
                            recommendation.QuotaUnit,
                            recommendation.ConvertedValueText,
                            RecommendationStatusText(recommendation));
                        recommendation.GridRowIndex = gridRowIndex;
                        resultGrid.Rows[gridRowIndex].Cells["SourceStatus"].ToolTipText = recommendation.Reason ?? "";
                        if (!String.IsNullOrWhiteSpace(item.OriginalName) || !String.IsNullOrWhiteSpace(item.RawRowText) || !String.IsNullOrWhiteSpace(item.AiNameReason))
                        {
                            resultGrid.Rows[gridRowIndex].Cells["QuantityName"].ToolTipText =
                                "\u539f\u59cb\u540d\u79f0\uff1a" + (item.OriginalName ?? "") +
                                "\r\nAI\u7406\u7531\uff1a" + (item.AiNameReason ?? "") +
                                "\r\n\u539f\u59cb\u884c\uff1a" + (item.RawRowText ?? "");
                        }
                        if (isContinuation)
                        {
                            DataGridViewRow gridRow = resultGrid.Rows[gridRowIndex];
                            gridRow.Cells["Correct"] = new DataGridViewTextBoxCell();
                            gridRow.Cells["Correct"].Value = "";
                            gridRow.Cells["Correct"].ReadOnly = true;
                        }
                        else if (ShouldQueueDeepSeek(recommendation))
                        {
                            recommendation.AiRowId = "r" + aiPending.Count.ToString(CultureInfo.InvariantCulture);
                            recommendation.AiPending = true;
                            aiPending.Add(new AiPendingRecommendation
                            {
                                Row = recommendation,
                                GridRowIndex = gridRowIndex,
                                Scope = scope,
                                Request = new DeepSeekRequestRow
                                {
                                    RowId = recommendation.AiRowId,
                                    Item = item,
                                    Candidates = recommendation.AiCandidates,
                                    MappingCandidates = recommendation.AiMappingCandidates
                                }
                            });
                            resultGrid.Rows[gridRowIndex].Cells["SourceStatus"].Value = "\u0041\u0049\u8865\u63a8\u4e2d";
                            stats.AiQueued++;
                        }
                        itemRowIndex++;
                    }
                }
            }
            finally
            {
                resultGrid.ResumeLayout();
            }

            stopwatch.Stop();

            statusLabel.Text = String.Format(
                CultureInfo.CurrentCulture,
                "\u5df2\u8bfb\u53d6 {0} \u884cExcel\u5de5\u7a0b\u91cf\uff0c\u5b9a\u989d\u7c7b\u578b\uff1a{1}\uff0c\u5bf9\u5e94\u6846\u547d\u4e2d {2} \u884c\uff0c\u7d22\u5f15\u68c0\u7d22 {3} \u884c\uff0c\u7a7a\u7ed3\u679c {4} \u884c\uff0c\u91cd\u590d\u590d\u7528 {5} \u884c\uff0cAI\u8865\u63a8 {6} \u884c\uff0c\u8017\u65f6 {7} ms\u3002",
                selection.Items.Count,
                categoryFilter,
                stats.MappingHits,
                stats.IndexSearches,
                stats.EmptyRows,
                stats.CacheHits,
                stats.AiQueued,
                stopwatch.ElapsedMilliseconds);
            statusLabel.Text += scope != null && scope.Strict
                ? "条目范围：" + scope.MatchedEntryCode + "（池" + scope.PoolKeys.Count.ToString(CultureInfo.InvariantCulture) + "条，严格）。"
                : (chapterLibrary != null && !chapterLibrary.IsEmpty ? "条目范围：全库。" : "");

            StartDeepSeekRecommendations(aiPending, requestVersion);
        }

        private string SelectedQuotaCategory
        {
            get
            {
                return quotaCategoryCombo == null || quotaCategoryCombo.SelectedItem == null
                    ? "\u9884\u7b97\u5b9a\u989d"
                    : Convert.ToString(quotaCategoryCombo.SelectedItem, CultureInfo.CurrentCulture);
            }
        }

        private static string BuildBatchCacheKey(ExcelQuantityItem item, string categoryFilter)
        {
            return TextMatcher.Normalize(item == null ? "" : item.Name) + "|" + TextMatcher.Normalize(item == null ? "" : item.Unit) + "|" + TextMatcher.Normalize(categoryFilter);
        }

        private static void RestoreOriginalQuantityNames(ExcelSelection selection)
        {
            if (selection == null)
            {
                return;
            }

            foreach (ExcelQuantityItem item in selection.Items)
            {
                if (item == null || String.IsNullOrWhiteSpace(item.OriginalName) || String.Equals(item.Name, item.OriginalName, StringComparison.Ordinal))
                {
                    continue;
                }

                item.Name = item.OriginalName;
                item.SectionName = item.OriginalName;
                item.ContextText = item.Name + " " + item.Unit + " " + item.ValueText + " " + item.RawRowText;
            }
        }

        // AI 列映射兜底：把整张表的原始网格交给 DeepSeek 判断列结构，再用确定性逻辑(BuildItemsFromColumnLayout)
        // 重建名称/单位。仅更新已识别条目（按行号对回），保留 OriginalName 以便关闭 AI 时还原。返回是否有改动。
        private bool ApplyAiColumnLayout(ExcelSelection selection)
        {
            if (selection == null || selection.Items.Count == 0
                || selection.RawRows == null || selection.RawRows.Count == 0
                || !deepSeekSettings.CanDetectColumns)
            {
                return false;
            }

            statusLabel.Text = "DeepSeek正在识别工程量表列结构...";
            statusLabel.Refresh();
            try
            {
                DeepSeekClient client = new DeepSeekClient(deepSeekSettings);
                DeepSeekColumnLayout layout = client.DetectColumnLayout(selection.RawRows);
                if (layout == null || layout.Confidence < 70 || layout.QuantityColumn <= 0
                    || layout.NameColumns == null || layout.NameColumns.Length == 0)
                {
                    return false;
                }

                List<int> descColumns = layout.NameColumns
                    .Where(c => c > 0 && c != layout.QuantityColumn && c != layout.UnitColumn)
                    .Distinct()
                    .OrderBy(c => c)
                    .ToList();
                if (descColumns.Count == 0)
                {
                    return false;
                }

                List<ExcelQuantityItem> rebuilt = BuildItemsFromColumnLayout(selection.RawRows, descColumns, layout.UnitColumn, layout.QuantityColumn, selection.WorksheetName);
                Dictionary<int, ExcelQuantityItem> byRow = new Dictionary<int, ExcelQuantityItem>();
                foreach (ExcelQuantityItem r in rebuilt)
                {
                    if (r != null)
                    {
                        byRow[r.RowNumber] = r;
                    }
                }

                int changed = 0;
                foreach (ExcelQuantityItem item in selection.Items)
                {
                    ExcelQuantityItem mapped;
                    if (item == null || !byRow.TryGetValue(item.RowNumber, out mapped) || String.IsNullOrWhiteSpace(mapped.Name))
                    {
                        continue;
                    }

                    if (String.IsNullOrWhiteSpace(item.OriginalName))
                    {
                        item.OriginalName = item.Name;
                    }

                    if (String.Equals(item.Name, mapped.Name, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    item.AiName = mapped.Name;
                    item.Name = mapped.Name;
                    if (!String.IsNullOrWhiteSpace(mapped.Unit))
                    {
                        item.Unit = mapped.Unit;
                    }
                    item.SectionName = mapped.SectionName;
                    item.ContextText = item.Name + " " + item.Unit + " " + item.ValueText + " " + item.RawRowText;
                    item.SkipAiNameNormalization = true;
                    changed++;
                }

                QuotaRecommendPanel.Log("AI column layout. nameCols=[" + String.Join(",", descColumns.Select(c => c.ToString(CultureInfo.InvariantCulture)).ToArray()) + "]"
                    + ", unitCol=" + layout.UnitColumn.ToString(CultureInfo.InvariantCulture)
                    + ", qtyCol=" + layout.QuantityColumn.ToString(CultureInfo.InvariantCulture)
                    + ", conf=" + layout.Confidence.ToString(CultureInfo.InvariantCulture)
                    + ", changed=" + changed.ToString(CultureInfo.InvariantCulture));
                return changed > 0;
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("AI column layout detection failed: " + ex.Message);
                return false;
            }
        }

        private void NormalizeQuantityNamesWithDeepSeek(ExcelSelection selection, bool force)
        {
            if (selection == null || selection.Items.Count == 0 || !deepSeekSettings.CanNormalizeNames)
            {
                return;
            }

            List<DeepSeekNameRequestRow> rows = new List<DeepSeekNameRequestRow>();
            for (int i = 0; i < selection.Items.Count; i++)
            {
                ExcelQuantityItem item = selection.Items[i];
                if (item == null)
                {
                    continue;
                }

                if (String.IsNullOrWhiteSpace(item.OriginalName))
                {
                    item.OriginalName = item.Name;
                }
                if (item.SkipAiNameNormalization && !force)
                {
                    continue;
                }
                rows.Add(new DeepSeekNameRequestRow
                {
                    RowId = "n" + i.ToString(CultureInfo.InvariantCulture),
                    Item = item
                });
            }

            if (rows.Count == 0)
            {
                return;
            }

            statusLabel.Text = "DeepSeek\u6b63\u5728\u8bc6\u522b\u5de5\u7a0b\u91cf\u540d\u79f0...";
            statusLabel.Refresh();
            try
            {
                DeepSeekClient client = new DeepSeekClient(deepSeekSettings);
                int batchSize = Math.Max(1, deepSeekSettings.MaxRowsPerBatch);
                int changed = 0;
                for (int i = 0; i < rows.Count; i += batchSize)
                {
                    List<DeepSeekNameRequestRow> batch = rows.Skip(i).Take(batchSize).ToList();
                    Dictionary<string, DeepSeekNameResult> byRow = client.NormalizeQuantityNames(batch)
                        .Where(r => r != null && !String.IsNullOrWhiteSpace(r.RowId))
                        .GroupBy(r => r.RowId, StringComparer.OrdinalIgnoreCase)
                        .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
                    foreach (DeepSeekNameRequestRow request in batch)
                    {
                        DeepSeekNameResult result;
                        if (!byRow.TryGetValue(request.RowId, out result) || String.IsNullOrWhiteSpace(result.QuantityName) || result.Confidence < 50)
                        {
                            continue;
                        }

                        ExcelQuantityItem item = request.Item;
                        item.AiName = CleanAiQuantityName(result.QuantityName);
                        if (String.IsNullOrWhiteSpace(item.AiName))
                        {
                            continue;
                        }

                        item.AiNameConfidence = result.Confidence;
                        item.AiNameReason = result.Reason;
                        item.Name = item.AiName;
                        item.SectionName = item.AiName;
                        item.ContextText = item.AiName + " " + item.Unit + " " + item.ValueText + " " + item.RawRowText;
                        changed++;
                    }
                }

                if (changed > 0)
                {
                    QuotaRecommendPanel.Log("DeepSeek normalized quantity names. changed=" + changed.ToString(CultureInfo.InvariantCulture));
                }
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("DeepSeek quantity name normalization failed: " + ex.Message);
            }
        }

        private static string CleanAiQuantityName(string name)
        {
            string value = (name ?? "").Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim();
            return value.Length > 80 ? value.Substring(0, 80).Trim() : value;
        }

        private static List<RecommendationRow> CloneRecommendationsForItem(ExcelQuantityItem item, List<RecommendationRow> source)
        {
            List<RecommendationRow> rows = new List<RecommendationRow>();
            foreach (RecommendationRow original in source)
            {
                RecommendationRow row = new RecommendationRow();
                row.Item = item;
                row.Record = original.Record;
                row.QuotaCode = original.QuotaCode;
                row.QuotaName = original.QuotaName;
                row.QuotaUnit = original.QuotaUnit;
                row.ConvertedValueText = String.IsNullOrWhiteSpace(original.QuotaUnit)
                    ? (item == null ? original.ConvertedValueText : item.ValueText)
                    : ConvertQuantityForIndex(item == null ? "" : item.ValueText, item == null ? "" : item.Unit, original.QuotaUnit);
                row.Score = original.Score;
                row.Reason = original.Reason;
                row.Source = original.Source;
                row.TargetKind = original.TargetKind;
                row.BoxId = original.BoxId;
                row.AiCandidates = original.AiCandidates;
                row.AiMappingCandidates = original.AiMappingCandidates;
                rows.Add(row);
            }
            return rows;
        }

        // 工程量名称含“模板/模版”视为模板工程量，不配定额
        private static bool IsFormworkQuantity(string name)
        {
            string text = name ?? "";
            return text.IndexOf("模板", StringComparison.Ordinal) >= 0 || text.IndexOf("模版", StringComparison.Ordinal) >= 0;
        }

        private List<RecommendationRow> BuildRecommendations(ExcelQuantityItem item, string categoryFilter, EntryScope scope, RecommendationBatchStats stats)
        {
            List<AiQuotaCandidate> aiCandidates = new List<AiQuotaCandidate>();
            // 人工扶正过的对应框优先（即便是模板，也按用户指定的定额显示）
            List<RecommendationRow> mapped = mappingStore.Find(item, categoryFilter, searchIndex, scope);
            if (mapped.Count > 0)
            {
                stats.MappingHits++;
                return mapped;
            }

            // 模板类工程量按规则不配定额：保留工程量行，推荐定额留空，且不走 AI 补推
            if (IsFormworkQuantity(item.Name))
            {
                RecommendationRow skip = new RecommendationRow();
                skip.Item = item;
                skip.QuotaCode = "";
                skip.QuotaName = "";
                skip.QuotaUnit = "";
                skip.ConvertedValueText = item.ValueText;
                skip.Score = 0;
                skip.Reason = "模板工程量按规则不推荐定额";
                skip.Source = "skip";
                skip.TargetKind = "quota";
                skip.AiCandidates = new List<AiQuotaCandidate>();
                skip.AiMappingCandidates = new List<AiMappingCandidate>();
                stats.EmptyRows++;
                return new List<RecommendationRow> { skip };
            }

            List<AiMappingCandidate> mappingCandidates = deepSeekSettings.CanDetectMapping
                ? mappingStore.BuildDeepSeekCandidates(item, categoryFilter, searchIndex, deepSeekSettings.MaxCandidatesPerRow, scope)
                : new List<AiMappingCandidate>();
            stats.IndexSearches++;
            foreach (AiQuotaCandidate candidate in searchIndex.BuildDeepSeekCandidates(item, categoryFilter, deepSeekSettings.MaxCandidatesPerRow, scope))
            {
                if (!aiCandidates.Any(c => c != null && c.Quota != null && candidate != null && candidate.Quota != null && String.Equals(c.Quota.QuotaCode, candidate.Quota.QuotaCode, StringComparison.OrdinalIgnoreCase)))
                {
                    aiCandidates.Add(candidate);
                }
            }

            // 严格条目模式：把整条目定额池补进候选，确保池里相关定额都能被本地匹配或 AI 选中（关键词没命中也不漏）
            if (scope != null && scope.Strict)
            {
                foreach (AiQuotaCandidate candidate in searchIndex.BuildScopeCandidates(item, scope, deepSeekSettings.MaxCandidatesPerRow))
                {
                    if (!aiCandidates.Any(c => c != null && c.Quota != null && candidate != null && candidate.Quota != null && String.Equals(c.Quota.QuotaCode, candidate.Quota.QuotaCode, StringComparison.OrdinalIgnoreCase)))
                    {
                        aiCandidates.Add(candidate);
                    }
                }
            }

            // \u672c\u5730\u7d22\u5f15\u9ad8\u5206\u547d\u4e2d\u65f6\u76f4\u63a5\u5c55\u793a\uff08\u6765\u6e90=\u672c\u5730\u7d22\u5f15\uff09\uff0cDeepSeek \u4e0d\u53ef\u7528\u4e5f\u6709\u7ed3\u679c\uff1b
            // AI \u8fd4\u56de\u4e14\u7f6e\u4fe1\u5ea6\u4e0d\u4f4e\u4e8e\u672c\u5730\u5206\u65f6\u624d\u5728 ApplyDeepSeekResults \u4e2d\u8986\u76d6\u3002
            RecommendationRow row = null;
            AiQuotaCandidate topLocal = aiCandidates
                .Where(c => c != null && c.Quota != null)
                .OrderByDescending(c => c.LocalScore)
                .ThenBy(c => c.Quota.SortOrder)
                .FirstOrDefault();
            if (topLocal != null && topLocal.LocalScore >= LocalIndexDisplayScore)
            {
                row = topLocal.Quota.ToRecommendation(item, topLocal.LocalScore);
            }

            if (row == null)
            {
                row = new RecommendationRow();
                row.Item = item;
                row.ConvertedValueText = item.ValueText;
                row.Score = 0;
                row.Reason = deepSeekSettings.CanRecommendQuota
                    ? "AI\u63a8\u8350\u7b49\u5f85\u8fd4\u56de"
                    : "DeepSeek AI\u672a\u542f\u7528\uff0c\u8bf7\u5728AI\u8bbe\u7f6e\u4e2d\u914d\u7f6e";
                row.Source = "empty";
                row.TargetKind = "quota";
            }

            row.AiMappingCandidates = mappingCandidates;
            row.AiCandidates = aiCandidates
                .Where(c => c != null && c.Quota != null)
                .GroupBy(c => c.Quota.QuotaCode, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.OrderByDescending(c => c.LocalScore).First())
                .OrderByDescending(c => c.LocalScore)
                .Take(deepSeekSettings.MaxCandidatesPerRow)
                .ToList();
            if (String.Equals(row.Source, "empty", StringComparison.OrdinalIgnoreCase) &&
                row.AiCandidates.Count == 0 && row.AiMappingCandidates.Count == 0)
            {
                row.Reason = "AI\u65e0\u6709\u6548\u5019\u9009\uff0c\u8bf7\u4eba\u5de5\u6276\u6b63";
                stats.EmptyRows++;
            }

            return new List<RecommendationRow> { row };
        }

        private List<RecommendationRow> FindMappingWithDeepSeek(ExcelQuantityItem item, string categoryFilter)
        {
            if (item == null || !deepSeekSettings.CanDetectMapping)
            {
                return new List<RecommendationRow>();
            }

            List<AiMappingCandidate> candidates = mappingStore.BuildDeepSeekCandidates(item, categoryFilter, searchIndex, Math.Max(3, deepSeekSettings.MaxCandidatesPerRow), null);
            if (candidates.Count == 0)
            {
                return new List<RecommendationRow>();
            }

            try
            {
                DeepSeekClient client = new DeepSeekClient(deepSeekSettings);
                DeepSeekMappingSelection selection = client.SelectMappingBox(item, candidates);
                if (selection == null || String.IsNullOrWhiteSpace(selection.BoxId) || selection.Confidence < deepSeekSettings.DisplayConfidence)
                {
                    return new List<RecommendationRow>();
                }

                AiMappingCandidate candidate = candidates.FirstOrDefault(c => String.Equals(c.BoxId, selection.BoxId, StringComparison.OrdinalIgnoreCase));
                return candidate == null ? new List<RecommendationRow>() : candidate.ToRecommendations(item, selection.Confidence, selection.Reason);
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("DeepSeek mapping detection failed: " + ex.Message);
                return new List<RecommendationRow>();
            }
        }

        private bool ShouldQueueDeepSeek(RecommendationRow row)
        {
            return row != null &&
                (deepSeekSettings.CanRecommendQuota || deepSeekSettings.CanDetectMapping) &&
                ((row.AiCandidates != null && row.AiCandidates.Count > 0) ||
                    (row.AiMappingCandidates != null && row.AiMappingCandidates.Count > 0));
        }

        private static string RecommendationStatusText(RecommendationRow row)
        {
            if (row == null)
            {
                return "";
            }

            if (String.Equals(row.Source, "mapping", StringComparison.OrdinalIgnoreCase))
            {
                return "\u5bf9\u5e94\u6846";
            }
            if (String.Equals(row.Source, "deepseek", StringComparison.OrdinalIgnoreCase))
            {
                return "AI\u8865\u63a8";
            }
            if (String.Equals(row.Source, "empty", StringComparison.OrdinalIgnoreCase))
            {
                return "\u672a\u5339\u914d";
            }
            if (String.Equals(row.Source, "skip", StringComparison.OrdinalIgnoreCase))
            {
                return "\u6a21\u677f\u514d\u63a8";
            }
            if (String.Equals(row.Source, "index", StringComparison.OrdinalIgnoreCase))
            {
                return "\u672c\u5730\u7d22\u5f15";
            }

            return row.Source ?? "";
        }

        private void StartDeepSeekRecommendations(List<AiPendingRecommendation> pending, int requestVersion)
        {
            if (pending == null || pending.Count == 0 || !deepSeekSettings.IsAvailable)
            {
                return;
            }

            foreach (List<AiPendingRecommendation> batch in BuildDeepSeekBatches(pending))
            {
                ThreadPool.QueueUserWorkItem(delegate
                {
                    List<DeepSeekSelection> selections = new List<DeepSeekSelection>();
                    string error = "";
                    try
                    {
                        DeepSeekClient client = new DeepSeekClient(deepSeekSettings);
                        selections = client.Rank(batch.Select(p => p.Request).ToList());
                    }
                    catch (Exception ex)
                    {
                        error = ex.Message;
                        QuotaRecommendPanel.Log("DeepSeek recommendation failed: " + ex.Message);
                        if (!IsNonRetryableDeepSeekError(error))
                        {
                            selections = RetryDeepSeekOneByOne(batch);
                            if (selections.Count > 0)
                            {
                                error = "";
                            }
                        }
                    }

                    try
                    {
                        if (!IsDisposed && IsHandleCreated)
                        {
                            BeginInvoke((MethodInvoker)delegate
                            {
                                ApplyDeepSeekResults(batch, selections, error, requestVersion);
                            });
                        }
                    }
                    catch
                    {
                    }
                });
            }
        }

        private static List<List<AiPendingRecommendation>> BuildDeepSeekBatches(List<AiPendingRecommendation> pending)
        {
            List<List<AiPendingRecommendation>> batches = new List<List<AiPendingRecommendation>>();
            List<AiPendingRecommendation> current = new List<AiPendingRecommendation>();
            int currentCost = 0;
            foreach (AiPendingRecommendation item in pending ?? new List<AiPendingRecommendation>())
            {
                // 每批控制在较小规模，让单次 AI 请求能在超时时间内返回，减少超时行
                int cost = EstimateDeepSeekRowCost(item);
                if (current.Count > 0 && (current.Count >= 6 || currentCost + cost > 50))
                {
                    batches.Add(current);
                    current = new List<AiPendingRecommendation>();
                    currentCost = 0;
                }

                current.Add(item);
                currentCost += cost;
            }

            if (current.Count > 0)
            {
                batches.Add(current);
            }

            return batches;
        }

        private static int EstimateDeepSeekRowCost(AiPendingRecommendation pending)
        {
            if (pending == null || pending.Request == null)
            {
                return 1;
            }

            int quotaCount = pending.Request.Candidates == null ? 0 : pending.Request.Candidates.Count;
            int mappingCount = pending.Request.MappingCandidates == null ? 0 : pending.Request.MappingCandidates.Count;
            int textCost = 0;
            if (pending.Request.Item != null)
            {
                textCost = Math.Min(8, ((pending.Request.Item.RawRowText ?? "").Length + (pending.Request.Item.Name ?? "").Length) / 40);
            }

            return 2 + quotaCount + mappingCount * 2 + textCost;
        }

        private void ApplyDeepSeekResults(List<AiPendingRecommendation> batch, List<DeepSeekSelection> selections, string error, int requestVersion)
        {
            if (requestVersion != aiRequestVersion || batch == null)
            {
                return;
            }

            Dictionary<string, DeepSeekSelection> byRow = new Dictionary<string, DeepSeekSelection>(StringComparer.OrdinalIgnoreCase);
            foreach (DeepSeekSelection selection in selections ?? new List<DeepSeekSelection>())
            {
                if (!String.IsNullOrWhiteSpace(selection.RowId) && !byRow.ContainsKey(selection.RowId))
                {
                    byRow[selection.RowId] = selection;
                }
            }

            int applied = 0;
            foreach (AiPendingRecommendation pending in (batch ?? new List<AiPendingRecommendation>()).OrderByDescending(p => p.GridRowIndex))
            {
                RecommendationRow row = pending.Row;
                if (row == null || pending.GridRowIndex < 0 || pending.GridRowIndex >= resultGrid.Rows.Count || pending.GridRowIndex >= recommendations.Count)
                {
                    continue;
                }
                if (!Object.ReferenceEquals(recommendations[pending.GridRowIndex], row))
                {
                    continue;
                }

                row.AiPending = false;
                DeepSeekSelection selection;
                if (!String.IsNullOrWhiteSpace(error))
                {
                    SetRecommendationStatus(pending.GridRowIndex, DeepSeekFailureStatus(error), error);
                    continue;
                }

                if (!byRow.TryGetValue(row.AiRowId, out selection))
                {
                    SetRecommendationStatus(pending.GridRowIndex, selections == null || selections.Count == 0 ? "AI\u8fd4\u56de\u4e3a\u7a7a" : "AI\u65e0\u7ed3\u679c", "");
                    continue;
                }

                if (!String.IsNullOrWhiteSpace(selection.ErrorText))
                {
                    SetRecommendationStatus(pending.GridRowIndex, DeepSeekFailureStatus(selection.ErrorText), selection.ErrorText);
                    continue;
                }

                AiMappingCandidate mappingCandidate = (pending.Request.MappingCandidates ?? new List<AiMappingCandidate>())
                    .FirstOrDefault(c => c != null && String.Equals(c.BoxId, selection.BoxId, StringComparison.OrdinalIgnoreCase));
                if (mappingCandidate != null && selection.Confidence >= 65)
                {
                    List<RecommendationRow> mappedRows = mappingCandidate.ToRecommendations(row.Item, selection.Confidence, selection.Reason);
                    if (mappedRows.Count > 0)
                    {
                        ApplyMappedRowsFromDeepSeek(pending.GridRowIndex, row, mappedRows);
                        applied += mappedRows.Count;
                        continue;
                    }
                }

                AiQuotaCandidate candidate = (pending.Request.Candidates ?? new List<AiQuotaCandidate>())
                    .FirstOrDefault(c => c != null && c.Quota != null && String.Equals(c.Quota.QuotaCode, selection.SelectedCode, StringComparison.OrdinalIgnoreCase));
                if (candidate == null)
                {
                    SetRecommendationStatus(pending.GridRowIndex, String.IsNullOrWhiteSpace(selection.SelectedCode) ? "AI\u65e0\u7ed3\u679c" : "AI\u8fd4\u56de\u65e0\u6548", String.IsNullOrWhiteSpace(selection.SelectedCode) ? selection.Reason : "\u8fd4\u56de\u7f16\u53f7\u4e0d\u5728\u672c\u5730\u5019\u9009\u4e2d");
                    continue;
                }

                // \u5019\u9009\u751f\u6210\u9636\u6bb5\u5df2\u6309\u6761\u76ee\u8fc7\u6ee4\uff0c\u8fd9\u91cc\u518d\u515c\u5e95\u4e00\u6b21\uff0c\u9632\u6b62 AI \u8fd4\u56de\u6c60\u5916\u5b9a\u989d
                if (pending.Scope != null && pending.Scope.Strict && !pending.Scope.Allows("quota", candidate.Quota.QuotaCode))
                {
                    SetRecommendationStatus(pending.GridRowIndex, "AI\u8d85\u51fa\u6761\u76ee\u8303\u56f4", "AI\u8fd4\u56de\u7684\u5b9a\u989d\u4e0d\u5728\u5f53\u524d\u6761\u76ee\u5b9a\u989d\u6c60\u5185");
                    continue;
                }

                int confidence = Math.Max(0, Math.Min(100, selection.Confidence));
                if (confidence < deepSeekSettings.DisplayConfidence ||
                    (!String.Equals(row.Source, "empty", StringComparison.OrdinalIgnoreCase) &&
                        !String.Equals(row.Source, "mapping", StringComparison.OrdinalIgnoreCase) &&
                        confidence < row.Score))
                {
                    SetRecommendationStatus(pending.GridRowIndex, "\u672c\u5730\u4f18\u5148", selection.Reason);
                    continue;
                }

                row.QuotaCode = candidate.Quota.QuotaCode;
                row.QuotaName = candidate.Quota.QuotaName;
                row.QuotaUnit = candidate.Quota.QuotaUnit;
                row.ConvertedValueText = RecommendDialog.ConvertQuantityForIndex(row.Item.ValueText, row.Item.Unit, candidate.Quota.QuotaUnit);
                row.Score = confidence;
                row.Reason = "DeepSeek\u5728\u672c\u5730\u5019\u9009\u4e2d\u9009\u62e9" + (String.IsNullOrWhiteSpace(selection.Reason) ? "" : "\uff1a" + selection.Reason);
                row.Source = "deepseek";
                row.TargetKind = "quota";
                row.BoxId = "";

                DataGridViewRow gridRow = resultGrid.Rows[pending.GridRowIndex];
                gridRow.Cells["QuotaCode"].Value = row.QuotaCode;
                gridRow.Cells["QuotaName"].Value = row.QuotaName;
                gridRow.Cells["QuotaUnit"].Value = row.QuotaUnit;
                gridRow.Cells["QuotaQuantity"].Value = row.ConvertedValueText;
                gridRow.Cells["Checked"].Value = confidence >= deepSeekSettings.AutoCheckConfidence &&
                    RecommendDialog.UnitCompatibleForIndex(row.Item.Unit, row.QuotaUnit);
                SetRecommendationStatus(pending.GridRowIndex, "AI\u5df2\u8865\u63a8", row.Reason);
                applied++;
            }

            if (applied > 0)
            {
                statusLabel.Text = statusLabel.Text + " AI\u5df2\u8865\u63a8 " + applied.ToString(CultureInfo.InvariantCulture) + " \u884c\u3002";
            }
        }

        private List<DeepSeekSelection> RetryDeepSeekOneByOne(List<AiPendingRecommendation> batch)
        {
            List<DeepSeekSelection> selections = new List<DeepSeekSelection>();
            foreach (AiPendingRecommendation pending in batch ?? new List<AiPendingRecommendation>())
            {
                if (pending == null || pending.Request == null)
                {
                    continue;
                }

                try
                {
                    DeepSeekClient client = new DeepSeekClient(deepSeekSettings);
                    List<DeepSeekSelection> result = client.Rank(new List<DeepSeekRequestRow> { pending.Request });
                    if (result != null && result.Count > 0)
                    {
                        selections.AddRange(result);
                    }
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("DeepSeek single-row retry failed: " + ex.Message);
                    selections.Add(new DeepSeekSelection
                    {
                        RowId = pending.Request.RowId,
                        Confidence = 0,
                        ErrorText = ex.Message
                    });
                }
            }

            return selections;
        }

        private static bool IsNonRetryableDeepSeekError(string error)
        {
            string value = (error ?? "").ToLowerInvariant();
            return value.Contains("401") ||
                value.Contains("402") ||
                value.Contains("authentication") ||
                value.Contains("api key") ||
                value.Contains("balance") ||
                value.Contains("insufficient") ||
                value.Contains("422");
        }

        private static string DeepSeekFailureStatus(string error)
        {
            string value = (error ?? "").ToLowerInvariant();
            if (value.Contains("401") || value.Contains("authentication") || value.Contains("api key"))
            {
                return "AI Key\u5f02\u5e38";
            }
            if (value.Contains("402") || value.Contains("balance") || value.Contains("insufficient"))
            {
                return "AI\u4f59\u989d\u4e0d\u8db3";
            }
            if (value.Contains("429") || value.Contains("rate limit"))
            {
                return "AI\u9650\u6d41";
            }
            if (value.Contains("timeout") || value.Contains("timed out") || value.Contains("\u8d85\u65f6"))
            {
                return "AI\u8d85\u65f6";
            }
            if (value.Contains("500") || value.Contains("503") || value.Contains("server") || value.Contains("overload"))
            {
                return "AI\u670d\u52a1\u5f02\u5e38";
            }
            if (value.Contains("json") || value.Contains("invalid") || value.Contains("format") || value.Contains("422"))
            {
                return "AI\u8fd4\u56de\u65e0\u6548";
            }
            if (value.Contains("connect") || value.Contains("network") || value.Contains("name resolution") || value.Contains("remote") || value.Contains("\u7f51\u7edc"))
            {
                return "AI\u7f51\u7edc\u5931\u8d25";
            }

            return "AI\u5931\u8d25";
        }

        private void ApplyMappedRowsFromDeepSeek(int gridRowIndex, RecommendationRow oldRow, List<RecommendationRow> mappedRows)
        {
            if (mappedRows == null || mappedRows.Count == 0 || gridRowIndex < 0 || gridRowIndex >= recommendations.Count || gridRowIndex >= resultGrid.Rows.Count)
            {
                return;
            }

            RecommendationRow first = mappedRows[0];
            recommendations[gridRowIndex] = first;
            first.GridRowIndex = gridRowIndex;
            DataGridViewRow gridRow = resultGrid.Rows[gridRowIndex];
            gridRow.Cells["Checked"].Value = first.Score >= 60;
            gridRow.Cells["QuotaCode"].Value = first.QuotaCode;
            gridRow.Cells["QuotaName"].Value = first.QuotaName;
            gridRow.Cells["QuotaUnit"].Value = first.QuotaUnit;
            gridRow.Cells["QuotaQuantity"].Value = first.ConvertedValueText;
            gridRow.Cells["SourceStatus"].Value = "AI\u5bf9\u5e94\u6846";
            gridRow.Cells["SourceStatus"].ToolTipText = first.Reason ?? "";

            int insertAt = gridRowIndex + 1;
            for (int i = 1; i < mappedRows.Count; i++)
            {
                RecommendationRow mapped = mappedRows[i];
                mapped.GridRowIndex = insertAt;
                recommendations.Insert(insertAt, mapped);
                resultGrid.Rows.Insert(insertAt, mapped.Score >= 60, "", "", mapped.Item.Unit, mapped.Item.ValueText, mapped.QuotaCode, mapped.QuotaName, mapped.QuotaUnit, mapped.ConvertedValueText, "AI\u5bf9\u5e94\u6846");
                DataGridViewRow continuation = resultGrid.Rows[insertAt];
                continuation.Cells["Correct"] = new DataGridViewTextBoxCell();
                continuation.Cells["Correct"].Value = "";
                continuation.Cells["Correct"].ReadOnly = true;
                continuation.Cells["SourceStatus"].ToolTipText = mapped.Reason ?? "";
                insertAt++;
            }

            for (int i = 0; i < recommendations.Count; i++)
            {
                recommendations[i].GridRowIndex = i;
            }
        }

        private void SetRecommendationStatus(int gridRowIndex, string text, string tooltip)
        {
            if (gridRowIndex < 0 || gridRowIndex >= resultGrid.Rows.Count)
            {
                return;
            }

            DataGridViewCell cell = resultGrid.Rows[gridRowIndex].Cells["SourceStatus"];
            cell.Value = text ?? "";
            cell.ToolTipText = tooltip ?? "";
        }

        private void ShowDeepSeekSettings()
        {
            using (DeepSeekSettingsDialog dialog = new DeepSeekSettingsDialog(deepSeekSettings))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                deepSeekSettings = dialog.Settings;
                aiRequestVersion++;
                statusLabel.Text = deepSeekSettings.IsAvailable
                    ? "DeepSeek AI\u8bbe\u7f6e\u5df2\u4fdd\u5b58\uff0c\u540e\u7eed\u91cd\u65b0\u8bfb\u53d6Excel/\u526a\u8d34\u677f\u65f6\u751f\u6548\u3002"
                    : "DeepSeek AI\u8bbe\u7f6e\u5df2\u4fdd\u5b58\uff0c\u5f53\u524d\u672a\u542f\u7528AI\u8865\u63a8\u3002";
            }
        }

        private RecommendationRow BuildRecommendation(ExcelQuantityItem item)
        {
            ScoredRecord best = records
                .Where(r => !r.IsCorrection)
                .Where(r => String.Equals(QuotaEntry.GuessKind(r.QuotaCode), "quota", StringComparison.OrdinalIgnoreCase))
                .Select(r => new ScoredRecord { Record = r, Score = ScoreRecord(r, item) })
                .Where(r => r.Score > 0)
                .OrderByDescending(r => r.Score)
                .ThenByDescending(r => r.Record.MatchScore)
                .FirstOrDefault();

            RecommendationRow row = new RecommendationRow();
            row.Item = item;
            row.ConvertedValueText = item.ValueText;
            if (best == null || best.Score <= 0)
            {
                row.Score = 0;
                row.Reason = "\u672a\u627e\u5230\u76f8\u4f3c\u5b66\u4e60\u8bb0\u5f55";
                return row;
            }

            row.Record = best.Record;
            row.QuotaCode = best.Record.QuotaCode;
            row.QuotaName = best.Record.QuotaName;
            row.QuotaUnit = best.Record.QuotaUnit;
            row.ConvertedValueText = ConvertQuantityForQuotaUnit(item.ValueText, item.Unit, row.QuotaUnit);
            row.Score = best.Score;
            row.Reason = BuildReason(best.Record, item);
            row.Source = "learning";
            row.TargetKind = QuotaEntry.GuessKind(row.QuotaCode);
            return row;
        }

        private static RecommendationRow BuildRecommendationFromRecord(ExcelQuantityItem item, LearningRecord record, int score)
        {
            RecommendationRow row = new RecommendationRow();
            row.Item = item;
            row.Record = record;
            row.QuotaCode = record.QuotaCode;
            row.QuotaName = record.QuotaName;
            row.QuotaUnit = record.QuotaUnit;
            row.ConvertedValueText = ConvertQuantityForQuotaUnit(item.ValueText, item.Unit, row.QuotaUnit);
            row.Score = score;
            row.Reason = "\u4eba\u5de5\u6276\u6b63";
            row.Source = "learning";
            row.TargetKind = QuotaEntry.GuessKind(row.QuotaCode);
            return row;
        }

        private static int ScoreRecord(LearningRecord record, ExcelQuantityItem item)
        {
            if (TextMatcher.IsSteelOnlyAgainstConcrete(item.Name, record.QuantityName) || TextMatcher.IsSteelOnlyAgainstConcrete(item.Name, record.QuotaName))
            {
                return 0;
            }

            string queryName = Normalize(item.Name);
            string context = Normalize(item.ContextText);
            string searchable = Normalize(String.Join(" ", new string[]
            {
                record.BudgetGroup,
                record.QuotaCode,
                record.QuotaName,
                record.QuotaUnit,
                record.QuantitySection,
                record.QuantityName,
                record.QuantityUnit,
                record.MatchReason
            }));

            int score = Math.Max(0, record.MatchScore / 3);
            string recordQuantityName = Normalize(record.QuantityName);
            string recordSection = Normalize(record.QuantitySection);
            string recordBudgetGroup = Normalize(record.BudgetGroup);
            bool textMatched = TextMatcher.HasStrongNamePairMatch(item.Name, record.QuantityName);

            if (!String.IsNullOrEmpty(queryName) && recordQuantityName == queryName)
            {
                score += 95;
            }
            else if (!String.IsNullOrEmpty(queryName) && (recordQuantityName.Contains(queryName) || queryName.Contains(recordQuantityName)))
            {
                score += 70;
            }
            else if (textMatched)
            {
                score += Math.Min(70, TextMatcher.NamePairScore(item.Name, record.QuantityName));
            }

            if (!String.IsNullOrEmpty(item.SectionName) && (recordSection.Contains(Normalize(item.SectionName)) || recordBudgetGroup.Contains(Normalize(item.SectionName))))
            {
                score += 35;
            }

            if (UnitCompatible(record.QuantityUnit, item.Unit) || UnitCompatible(record.QuotaUnit, item.Unit))
            {
                score += 20;
            }

            foreach (string token in Tokenize(item.ContextText))
            {
                if (token.Length >= 2 && searchable.Contains(token))
                {
                    score += 12;
                }
            }

            return textMatched ? score : 0;
        }

        private static string BuildReason(LearningRecord record, ExcelQuantityItem item)
        {
            List<string> reasons = new List<string>();
            if (UnitCompatible(record.QuantityUnit, item.Unit) || UnitCompatible(record.QuotaUnit, item.Unit))
            {
                reasons.Add("\u5355\u4f4d\u63a5\u8fd1");
            }
            if (Normalize(record.QuantityName).Contains(Normalize(item.Name)) || Normalize(item.Name).Contains(Normalize(record.QuantityName)))
            {
                reasons.Add("\u5de5\u7a0b\u91cf\u540d\u79f0\u76f8\u4f3c");
            }
            if (!String.IsNullOrWhiteSpace(record.MatchReason))
            {
                reasons.Add(record.MatchReason);
            }
            return reasons.Count == 0 ? "\u5b66\u4e60\u5e93\u76f8\u4f3c\u8bb0\u5f55" : String.Join(";", reasons.ToArray());
        }

        private void SetChecked(bool value)
        {
            foreach (DataGridViewRow row in resultGrid.Rows)
            {
                row.Cells["Checked"].Value = value;
            }
        }

        private void ResultGridCellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0)
            {
                return;
            }

            if (resultGrid.Columns[e.ColumnIndex].Name == "Correct")
            {
                if (!(resultGrid.Rows[e.RowIndex].Cells[e.ColumnIndex] is DataGridViewButtonCell))
                {
                    return;
                }
                CorrectRecommendation(e.RowIndex);
            }
        }

        private void ResultGridCellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= recommendations.Count || e.ColumnIndex < 0)
            {
                return;
            }

            if (resultGrid.Columns[e.ColumnIndex].Name != "QuantityName")
            {
                return;
            }

            object value = resultGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            string newName = Convert.ToString(value, CultureInfo.CurrentCulture).Trim();
            if (String.IsNullOrWhiteSpace(newName))
            {
                resultGrid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = recommendations[e.RowIndex].Item.Name;
                return;
            }

            ExcelQuantityItem oldItem = recommendations[e.RowIndex].Item;
            string oldName = oldItem.Name;
            foreach (RecommendationRow row in recommendations.Where(r => Object.ReferenceEquals(r.Item, oldItem)))
            {
                row.Item.Name = newName;
                row.Item.SectionName = newName;
                row.Item.ContextText = ReplaceContextName(row.Item.ContextText, oldName, newName);
            }
        }

        private void CorrectRecommendation(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= recommendations.Count)
            {
                return;
            }

            RecommendationRow recommendation = recommendations[rowIndex];
            List<QuotaEntry> quotas = GetSelectedQuotaEntries(mainForm);
            if (quotas.Count == 0)
            {
                MessageBox.Show(this, "\u8bf7\u5148\u5728\u5b9a\u989d\u8f93\u5165\u8868\u4e2d\u9009\u4e2d\u4e00\u6761\u6216\u591a\u6761\u6b63\u786e\u7684\u5b9a\u989d\uff0c\u518d\u70b9\u51fb\u6276\u6b63\u3002", "\u6276\u6b63\u5b9a\u989d", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            try
            {
                mappingStore.Correct(recommendation.Item, recommendation, quotas, currentEntryScope);
                if (currentEntryScope != null && currentEntryScope.Strict && chapterLibrary != null)
                {
                    foreach (QuotaEntry quota in quotas)
                    {
                        chapterLibrary.AddUserQuota(currentEntryScope, quota.TargetKind, quota.QuotaCode, quota.QuotaName, quota.QuotaUnit);
                    }
                }
                LearningStore.ReplaceCorrections(recommendation.Item, quotas);
                records.Clear();
                records.AddRange(LearningStore.Load());
                if (currentSelection != null)
                {
                    FillRecommendations(currentSelection);
                }

                statusLabel.Text = "\u5df2\u6276\u6b63\uff1a" + recommendation.Item.Name + " -> " + String.Join("\uff0c", quotas.Select(q => q.QuotaCode).ToArray()) + "\u3002\u4e0b\u6b21\u63a8\u8350\u5c06\u4f18\u5148\u4f7f\u7528\u8be5\u5bf9\u5e94\u5173\u7cfb\u3002";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "\u6276\u6b63\u5931\u8d25\uff1a" + ex.Message, "\u6276\u6b63\u5b9a\u989d", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static string ReplaceContextName(string context, string oldName, string newName)
        {
            if (String.IsNullOrWhiteSpace(context))
            {
                return newName;
            }

            if (!String.IsNullOrWhiteSpace(oldName) && context.IndexOf(oldName, StringComparison.Ordinal) >= 0)
            {
                return context.Replace(oldName, newName);
            }

            return newName + " " + context;
        }

        private static List<QuotaEntry> GetSelectedQuotaEntries(Form mainForm)
        {
            List<QuotaEntry> result = new List<QuotaEntry>();
            DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
            if (grid == null)
            {
                return result;
            }

            SortedDictionary<int, DataGridViewRow> rows = new SortedDictionary<int, DataGridViewRow>();
            foreach (DataGridViewRow row in grid.SelectedRows)
            {
                if (row != null && !row.IsNewRow)
                {
                    rows[row.Index] = row;
                }
            }

            foreach (DataGridViewCell cell in grid.SelectedCells)
            {
                if (cell.RowIndex >= 0 && cell.RowIndex < grid.Rows.Count)
                {
                    DataGridViewRow row = grid.Rows[cell.RowIndex];
                    if (row != null && !row.IsNewRow)
                    {
                        rows[row.Index] = row;
                    }
                }
            }

            if (rows.Count == 0 && grid.CurrentRow != null && !grid.CurrentRow.IsNewRow)
            {
                rows[grid.CurrentRow.Index] = grid.CurrentRow;
            }

            foreach (DataGridViewRow row in rows.Values)
            {
                QuotaEntry entry = new QuotaEntry();
                entry.QuotaCode = GetRowValue(row, "\u5b9a\u989d\u7f16\u53f7", "\u5b9a\u989d\u7f16\u53f7DE", "\u7f16\u53f7");
                entry.QuotaName = GetRowValue(row, "\u5de5\u7a0b\u6216\u8d39\u7528\u9879\u76ee\u540d\u79f0", "\u540d\u79f0", "\u9879\u76ee\u540d\u79f0");
                entry.QuotaUnit = GetRowValue(row, "\u5355\u4f4d");
                if (!String.IsNullOrWhiteSpace(entry.QuotaCode))
                {
                    entry.TargetKind = QuotaEntry.GuessKind(entry.QuotaCode);
                    result.Add(entry);
                }
            }

            return result
                .GroupBy(q => q.QuotaCode, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();
        }

        private static string GetRowValue(DataGridViewRow row, params string[] names)
        {
            DataRowView rowView = row.DataBoundItem as DataRowView;
            if (rowView != null)
            {
                foreach (string name in names)
                {
                    if (rowView.DataView.Table.Columns.Contains(name))
                    {
                        object value = rowView[name];
                        if (value != null && value != DBNull.Value)
                        {
                            return Convert.ToString(value, CultureInfo.CurrentCulture).Trim();
                        }
                    }
                }
            }

            if (row.DataGridView != null)
            {
                foreach (DataGridViewColumn column in row.DataGridView.Columns)
                {
                    foreach (string name in names)
                    {
                        if (String.Equals(column.Name, name, StringComparison.OrdinalIgnoreCase) ||
                            String.Equals(column.HeaderText, name, StringComparison.OrdinalIgnoreCase))
                        {
                            object value = row.Cells[column.Index].Value;
                            if (value != null)
                            {
                                return Convert.ToString(value, CultureInfo.CurrentCulture).Trim();
                            }
                        }
                    }
                }
            }

            return "";
        }

        private void CopyCheckedForManualPaste()
        {
            List<RecommendationRow> rows = GetCheckedRecommendations();
            if (rows.Count == 0)
            {
                statusLabel.Text = "\u6ca1\u6709\u52fe\u9009\u4efb\u4f55\u63a8\u8350\u884c\u3002";
                return;
            }

            Clipboard.SetText(BuildTabSeparated(rows));
            mappingStore.Accept(rows, currentEntryScope);
            statusLabel.Text = "\u5df2\u590d\u5236 " + rows.Count.ToString(CultureInfo.InvariantCulture) + " \u6761\u7c98\u8d34\u7528\u5185\u5bb9\uff1a\u7b2c1\u5217\u5b9a\u989d\u7f16\u53f7\uff0c\u7b2c4\u5217\u5de5\u7a0b\u6570\u91cf\u3002\u8bf7\u5728\u5b9a\u989d\u8868\u7b2c1\u5217\u76ee\u6807\u4f4d\u7f6e Ctrl+V\u3002";
        }

        private List<RecommendationRow> GetCheckedRecommendations()
        {
            List<RecommendationRow> rows = new List<RecommendationRow>();
            for (int i = 0; i < resultGrid.Rows.Count && i < recommendations.Count; i++)
            {
                object value = resultGrid.Rows[i].Cells["Checked"].Value;
                bool isChecked = value is bool && (bool)value;
                if (isChecked && !String.IsNullOrWhiteSpace(recommendations[i].QuotaCode))
                {
                    rows.Add(recommendations[i]);
                }
            }

            return rows;
        }

        private static string BuildTabSeparated(List<RecommendationRow> rows)
        {
            StringBuilder builder = new StringBuilder();
            foreach (RecommendationRow row in rows)
            {
                builder.Append(CleanCell(row.QuotaCode)).Append('\t')
                    .Append('\t')
                    .Append('\t')
                    .Append(CleanCell(row.ConvertedValueText))
                    .AppendLine();
            }

            return builder.ToString();
        }

        private static string CleanCell(string value)
        {
            return (value ?? "").Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim();
        }

        private static string ConvertQuantityForQuotaUnit(string quantityText, string excelUnit, string quotaUnit)
        {
            decimal quantity;
            if (!TryEvaluateQuantity(quantityText, out quantity))
            {
                return quantityText;
            }

            UnitScale excel = ParseUnitScale(excelUnit);
            UnitScale quota = ParseUnitScale(quotaUnit);
            if (String.IsNullOrEmpty(excel.BaseUnit) || String.IsNullOrEmpty(quota.BaseUnit))
            {
                return FormatDecimal(quantity);
            }

            if (!String.Equals(excel.BaseUnit, quota.BaseUnit, StringComparison.Ordinal))
            {
                return FormatDecimal(quantity);
            }

            if (excel.Scale <= 0 || quota.Scale <= 0)
            {
                return FormatDecimal(quantity);
            }

            decimal converted = quantity * excel.Scale / quota.Scale;
            return FormatDecimal(converted);
        }

        internal static string ConvertQuantityForIndex(string quantityText, string excelUnit, string quotaUnit)
        {
            return ConvertQuantityForQuotaUnit(quantityText, excelUnit, quotaUnit);
        }

        private static bool TryEvaluateQuantity(string text, out decimal value)
        {
            value = 0m;
            string expression = (text ?? "").Trim();
            if (expression.StartsWith("=", StringComparison.Ordinal))
            {
                expression = expression.Substring(1);
            }

            if (Decimal.TryParse(expression, NumberStyles.Float, CultureInfo.InvariantCulture, out value) ||
                Decimal.TryParse(expression, NumberStyles.Float, CultureInfo.CurrentCulture, out value))
            {
                return true;
            }

            expression = expression
                .Replace("\u00d7", "*")
                .Replace("x", "*")
                .Replace("X", "*")
                .Replace("\uff08", "(")
                .Replace("\uff09", ")");

            if (expression.Any(ch => !(Char.IsDigit(ch) || ch == '.' || ch == '+' || ch == '-' || ch == '*' || ch == '/' || ch == '(' || ch == ')' || Char.IsWhiteSpace(ch))))
            {
                return false;
            }

            try
            {
                DataTable table = new DataTable();
                object result = table.Compute(expression, String.Empty);
                value = Convert.ToDecimal(result, CultureInfo.InvariantCulture);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string FormatDecimal(decimal value)
        {
            return value.ToString("0.########", CultureInfo.InvariantCulture);
        }

        private static UnitScale ParseUnitScale(string unit)
        {
            string normalized = NormalizeRawUnit(unit);
            decimal scale = 1m;
            int index = 0;
            while (index < normalized.Length && (Char.IsDigit(normalized[index]) || normalized[index] == '.'))
            {
                index++;
            }

            if (index > 0)
            {
                Decimal.TryParse(normalized.Substring(0, index), NumberStyles.Float, CultureInfo.InvariantCulture, out scale);
            }

            string baseUnit = index > 0 ? normalized.Substring(index) : normalized;
            UnitScale result = new UnitScale();
            result.Scale = scale <= 0 ? 1m : scale;
            result.BaseUnit = baseUnit;
            return result;
        }

        private static string NormalizeRawUnit(string unit)
        {
            return (unit ?? "")
                .Trim()
                .ToLowerInvariant()
                .Replace(" ", "")
                .Replace("\u3000", "")
                .Replace("\u00b2", "2")
                .Replace("\u00b3", "3")
                .Replace("\uff4d", "m")
                .Replace("\u33a1", "m2")
                .Replace("\u33a5", "m3")
                .Replace("\u7acb\u65b9\u7c73", "m3")
                .Replace("\u5e73\u65b9\u7c73", "m2")
                .Replace("\u5ef6\u7c73", "m")
                .Replace("\u7c73", "m")
                .Replace("\u5428", "t")
                .Replace("\u5343\u514b", "kg");
        }

        private static bool TryReadActiveExcelSelection(out ExcelSelection selection, out string error)
        {
            selection = null;
            error = "";
            try
            {
                dynamic excel = GetActiveSpreadsheetApplication();
                if (excel == null)
                {
                    error = "\u6ca1\u6709\u627e\u5230\u6b63\u5728\u8fd0\u884c\u7684 Excel/WPS\u3002\u8bf7\u5148\u5728\u5de5\u7a0b\u91cf\u8868\u4e2d\u6846\u9009\u8981\u63a8\u8350\u7684\u884c\u3002";
                    return false;
                }

                dynamic workbook = excel.ActiveWorkbook;
                dynamic sheet = excel.ActiveSheet;
                dynamic range = excel.Selection;
                if (workbook == null || sheet == null || range == null)
                {
                    error = "\u8bf7\u5148\u5728 Excel/WPS \u4e2d\u6846\u9009\u5de5\u7a0b\u91cf\u884c\u3002";
                    return false;
                }

                selection = new ExcelSelection();
                selection.WorkbookPath = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture);
                selection.WorksheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture);

                List<List<CellValue>> rawRows;
                selection.Items.AddRange(BuildQuantityItemsFromRange(range, selection.WorksheetName, out rawRows));
                selection.RawRows = rawRows;
                ApplyActiveLeftGroups(selection);
                NormalizeSelectionItems(selection);
                LogSelectionSummary("Excel selection", selection);

                if (selection.Items.Count == 0)
                {
                    error = "\u6846\u9009\u533a\u57df\u91cc\u6ca1\u6709\u8bc6\u522b\u5230\u5de5\u7a0b\u91cf\u884c\uff0c\u8bf7\u628a\u5de5\u7a0b\u91cf\u540d\u79f0\u3001\u5355\u4f4d\u3001\u6570\u91cf\u4e00\u8d77\u6846\u9009\u3002";
                    return false;
                }

                return true;
            }
            catch (COMException)
            {
                error = "\u8bfb\u53d6 Excel/WPS \u6846\u9009\u533a\u57df\u5931\u8d25\uff0c\u8bf7\u786e\u8ba4\u8868\u683c\u5df2\u6253\u5f00\u5e76\u5df2\u6846\u9009\u5de5\u7a0b\u91cf\u884c\u3002";
                return false;
            }
            catch (Exception ex)
            {
                error = "\u8bfb\u53d6 Excel/WPS \u6846\u9009\u533a\u57df\u5931\u8d25\uff1a" + ex.Message;
                return false;
            }
        }

        private static bool TryReadClipboardSelection(out ExcelSelection selection, out string error)
        {
            selection = null;
            error = "";
            try
            {
                if (!Clipboard.ContainsText(TextDataFormat.UnicodeText) && !Clipboard.ContainsText())
                {
                    error = "\u526a\u8d34\u677f\u91cc\u6ca1\u6709 Excel \u6846\u9009\u5185\u5bb9\u3002\u8bf7\u5148\u5728 Excel/WPS \u91cc\u6846\u9009\u5de5\u7a0b\u91cf\u884c\u5e76\u6309 Ctrl+C\u3002";
                    return false;
                }

                string text = Clipboard.ContainsText(TextDataFormat.UnicodeText)
                    ? Clipboard.GetText(TextDataFormat.UnicodeText)
                    : Clipboard.GetText();
                if (String.IsNullOrWhiteSpace(text))
                {
                    error = "\u526a\u8d34\u677f\u5185\u5bb9\u4e3a\u7a7a\u3002";
                    return false;
                }

                string[] lines = text.Replace("\r\n", "\n").Replace('\r', '\n').Split(new char[] { '\n' }, StringSplitOptions.RemoveEmptyEntries);
                ExcelSelection textSelection = BuildSelectionFromClipboardLines(lines);

                ExcelSelection activeSelection;
                if (TryReadActiveExcelSelectionForClipboard(lines.Length, out activeSelection))
                {
                    if (textSelection.Items.Count > 0 && SelectionGroupScore(textSelection) > SelectionGroupScore(activeSelection))
                    {
                        selection = textSelection;
                        ApplyActiveLeftGroups(selection);
                        NormalizeSelectionItems(selection);
                        LogSelectionSummary("Clipboard text selection preferred over active selection", selection);
                        return true;
                    }

                    selection = activeSelection;
                    LogSelectionSummary("Clipboard backed by active Excel selection", selection);
                    return true;
                }

                if (textSelection.Items.Count > 0)
                {
                    selection = textSelection;
                    ApplyActiveLeftGroups(selection);
                    NormalizeSelectionItems(selection);
                    LogSelectionSummary("Clipboard text selection", selection);
                    return true;
                }

                if (TryReadClipboardHtmlSelection(out selection))
                {
                    ApplyActiveLeftGroups(selection);
                    NormalizeSelectionItems(selection);
                    LogSelectionSummary("Clipboard HTML selection", selection);
                    return true;
                }

                selection = textSelection;
                LogSelectionSummary("Clipboard text selection", selection);

                if (selection.Items.Count == 0)
                {
                    error = "\u526a\u8d34\u677f\u5185\u5bb9\u91cc\u6ca1\u6709\u8bc6\u522b\u5230\u5de5\u7a0b\u91cf\u884c\uff0c\u8bf7\u628a\u5de5\u7a0b\u91cf\u540d\u79f0\u3001\u5355\u4f4d\u3001\u6570\u91cf\u4e00\u8d77\u590d\u5236\u3002";
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                error = "\u8bfb\u53d6\u526a\u8d34\u677f\u5931\u8d25\uff1a" + ex.Message;
                return false;
            }
        }

        private static bool TryReadActiveExcelSelectionForClipboard(int expectedRows, out ExcelSelection selection)
        {
            selection = null;
            try
            {
                dynamic excel = GetActiveSpreadsheetApplication();
                if (excel == null)
                {
                    return false;
                }

                dynamic workbook = excel.ActiveWorkbook;
                dynamic sheet = excel.ActiveSheet;
                dynamic range = excel.Selection;
                if (workbook == null || sheet == null || range == null)
                {
                    return false;
                }

                int rowCount = Convert.ToInt32(range.Rows.Count, CultureInfo.InvariantCulture);
                bool rowCountDiffers = expectedRows > 0 && rowCount != expectedRows;

                selection = new ExcelSelection();
                selection.WorkbookPath = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture);
                selection.WorksheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture);

                List<List<CellValue>> rawRows;
                selection.Items.AddRange(BuildQuantityItemsFromRange(range, selection.WorksheetName, out rawRows));
                selection.RawRows = rawRows;
                ApplyActiveLeftGroups(selection);
                NormalizeSelectionItems(selection);
                if (selection.Items.Count == 0)
                {
                    return false;
                }

                if (rowCountDiffers)
                {
                    QuotaRecommendPanel.Log("Clipboard line count differs from active selection rows. textLines="
                        + expectedRows.ToString(CultureInfo.InvariantCulture)
                        + ", selectionRows=" + rowCount.ToString(CultureInfo.InvariantCulture)
                        + ", parsedItems=" + selection.Items.Count.ToString(CultureInfo.InvariantCulture));
                    if (rowCount < expectedRows)
                    {
                        return true;
                    }

                    return selection.Items.Count <= expectedRows;
                }

                return true;
            }
            catch
            {
                selection = null;
                return false;
            }
        }

        private static ExcelSelection BuildSelectionFromClipboardLines(string[] lines)
        {
            ExcelSelection selection = new ExcelSelection();
            selection.WorksheetName = "\u526a\u8d34\u677f";
            List<List<string>> textTable = new List<List<string>>();
            for (int i = 0; i < lines.Length; i++)
            {
                textTable.Add(lines[i].Split('\t').ToList());
            }

            List<List<CellValue>> rawRows;
            selection.Items.AddRange(BuildQuantityItemsFromTextTable(textTable, selection.WorksheetName, out rawRows));
            selection.RawRows = rawRows;
            NormalizeSelectionItems(selection);
            return selection;
        }

        private static int SelectionGroupScore(ExcelSelection selection)
        {
            if (selection == null || selection.Items.Count == 0)
            {
                return 0;
            }

            int score = 0;
            foreach (ExcelQuantityItem item in selection.Items)
            {
                if (item == null)
                {
                    continue;
                }

                string name = (item.Name ?? "").Trim();
                string section = (item.SectionName ?? "").Trim();
                if (LooksLikeGroupText(section)
                    && !String.Equals(section, name, StringComparison.Ordinal)
                    && name.IndexOf(section, StringComparison.Ordinal) >= 0)
                {
                    score += 2;
                }
                else if (name.IndexOf(' ') >= 0)
                {
                    score += 1;
                }
            }

            return score;
        }

        private static bool TryReadClipboardHtmlSelection(out ExcelSelection selection)
        {
            selection = null;
            try
            {
                if (!Clipboard.ContainsText(TextDataFormat.Html))
                {
                    return false;
                }

                string html = Clipboard.GetText(TextDataFormat.Html);
                if (String.IsNullOrWhiteSpace(html))
                {
                    return false;
                }

                List<List<string>> table = ParseHtmlTable(html);
                if (table.Count == 0)
                {
                    return false;
                }

                selection = new ExcelSelection();
                selection.WorksheetName = "\u526a\u8d34\u677fHTML";
                List<List<CellValue>> rawRows;
                selection.Items.AddRange(BuildQuantityItemsFromTextTable(table, selection.WorksheetName, out rawRows));
                selection.RawRows = rawRows;

                NormalizeSelectionItems(selection);

                if (selection.Items.Count == 0)
                {
                    selection = null;
                    return false;
                }

                QuotaRecommendPanel.Log("Clipboard HTML parsed. htmlRows=" + table.Count.ToString(CultureInfo.InvariantCulture) + ", items=" + selection.Items.Count.ToString(CultureInfo.InvariantCulture));
                return true;
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("Clipboard HTML parse failed: " + ex.Message);
                selection = null;
                return false;
            }
        }

        private static List<List<string>> ParseHtmlTable(string html)
        {
            List<List<string>> result = new List<List<string>>();
            Dictionary<int, CarryCell> carries = new Dictionary<int, CarryCell>();
            MatchCollection rowMatches = Regex.Matches(html, @"<tr\b[\s\S]*?</tr>", RegexOptions.IgnoreCase);
            foreach (Match rowMatch in rowMatches)
            {
                string rowHtml = rowMatch.Value;
                MatchCollection cellMatches = Regex.Matches(rowHtml, @"<t[dh]\b[^>]*>[\s\S]*?</t[dh]>", RegexOptions.IgnoreCase);
                if (cellMatches.Count == 0)
                {
                    continue;
                }

                List<string> row = new List<string>();
                int col = 0;
                foreach (Match cellMatch in cellMatches)
                {
                    FillCarriedCells(row, carries, ref col);

                    string cellHtml = cellMatch.Value;
                    int rowSpan = Math.Max(1, ParseSpan(cellHtml, "rowspan"));
                    int colSpan = Math.Max(1, ParseSpan(cellHtml, "colspan"));
                    string text = HtmlCellText(cellHtml);
                    for (int i = 0; i < colSpan; i++)
                    {
                        SetListValue(row, col + i, text);
                        if (rowSpan > 1)
                        {
                            carries[col + i] = new CarryCell { Text = text, RemainingRows = rowSpan - 1 };
                        }
                    }

                    col += colSpan;
                }

                FillCarriedCells(row, carries, ref col);
                if (row.Any(v => !String.IsNullOrWhiteSpace(v)))
                {
                    result.Add(row);
                }
            }

            return result;
        }

        private static void FillCarriedCells(List<string> row, Dictionary<int, CarryCell> carries, ref int col)
        {
            while (carries.ContainsKey(col))
            {
                CarryCell carry = carries[col];
                SetListValue(row, col, carry.Text);
                carry.RemainingRows--;
                if (carry.RemainingRows <= 0)
                {
                    carries.Remove(col);
                }
                else
                {
                    carries[col] = carry;
                }

                col++;
            }
        }

        private static int ParseSpan(string html, string name)
        {
            Match match = Regex.Match(html, name + "\\s*=\\s*[\"']?(\\d+)", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                return 1;
            }

            int value;
            return Int32.TryParse(match.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out value) ? value : 1;
        }

        private static string HtmlCellText(string cellHtml)
        {
            string text = Regex.Replace(cellHtml, @"<br\s*/?>", "\n", RegexOptions.IgnoreCase);
            text = Regex.Replace(text, @"<[^>]+>", "");
            return NormalizeCellText(WebUtility.HtmlDecode(text));
        }

        private static void SetListValue(List<string> row, int index, string value)
        {
            while (row.Count <= index)
            {
                row.Add("");
            }

            row[index] = value;
        }

        private static List<ExcelQuantityItem> BuildQuantityItemsFromRange(dynamic range, string worksheetName, out List<List<CellValue>> rawRows)
        {
            List<List<CellValue>> rows = new List<List<CellValue>>();
            int rowCount = Convert.ToInt32(range.Rows.Count, CultureInfo.InvariantCulture);
            int columnCount = Convert.ToInt32(range.Columns.Count, CultureInfo.InvariantCulture);
            for (int r = 1; r <= rowCount; r++)
            {
                List<CellValue> row = new List<CellValue>();
                string leftGroup = TryReadLeftGroupFromRangeRow(range, r);
                if (LooksLikeGroupText(leftGroup))
                {
                    CellValue value = new CellValue();
                    value.Text = leftGroup;
                    value.Formula = "";
                    value.Address = "LEFT" + r.ToString(CultureInfo.InvariantCulture);
                    value.RowNumber = r;
                    value.SourceIndex = 0;
                    row.Add(value);
                }

                for (int c = 1; c <= columnCount; c++)
                {
                    dynamic cell = range.Cells[r, c];
                    string text = ReadCellTextWithMerge(cell);
                    string formula = "";
                    try
                    {
                        formula = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture);
                    }
                    catch
                    {
                    }

                    if (!String.IsNullOrWhiteSpace(text) || !String.IsNullOrWhiteSpace(formula))
                    {
                        CellValue value = new CellValue();
                        value.Text = text;
                        value.Formula = formula;
                        value.Address = Convert.ToString(cell.Address(false, false), CultureInfo.InvariantCulture);
                        value.RowNumber = Convert.ToInt32(cell.Row, CultureInfo.InvariantCulture);
                        value.SourceIndex = c;
                        row.Add(value);
                    }
                }

                rows.Add(row);
            }

            rawRows = rows;
            return BuildQuantityItemsFromCellRows(rows, worksheetName);
        }

        private static List<ExcelQuantityItem> BuildQuantityItemsFromTextTable(List<List<string>> table, string worksheetName, out List<List<CellValue>> rawRows)
        {
            List<List<CellValue>> rows = new List<List<CellValue>>();
            for (int r = 0; r < table.Count; r++)
            {
                List<CellValue> row = new List<CellValue>();
                List<string> source = table[r] ?? new List<string>();
                for (int c = 0; c < source.Count; c++)
                {
                    string text = NormalizeCellText(source[c]);
                    if (String.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    CellValue value = new CellValue();
                    value.Text = text;
                    value.Formula = text.StartsWith("=", StringComparison.Ordinal) ? text : "";
                    value.Address = "R" + (r + 1).ToString(CultureInfo.InvariantCulture) + "C" + (c + 1).ToString(CultureInfo.InvariantCulture);
                    value.RowNumber = r + 1;
                    value.SourceIndex = c + 1;
                    row.Add(value);
                }

                rows.Add(row);
            }

            rawRows = rows;
            return BuildQuantityItemsFromCellRows(rows, worksheetName);
        }

        private static List<ExcelQuantityItem> BuildQuantityItemsFromCellRows(List<List<CellValue>> rows, string worksheetName)
        {
            List<ExcelQuantityItem> result = new List<ExcelQuantityItem>();
            if (rows == null || rows.Count == 0)
            {
                return result;
            }

            int quantityColumn = FindQuantityColumn(rows);
            if (quantityColumn < 0)
            {
                return result;
            }

            int unitColumn = FindUnitColumn(rows, quantityColumn);
            int nameBoundary = unitColumn >= 0 && unitColumn < quantityColumn ? unitColumn : quantityColumn;
            List<int> descColumns = FindDescriptionColumns(rows, nameBoundary, unitColumn, quantityColumn);
            if (descColumns.Count == 0)
            {
                return result;
            }

            result = BuildItemsFromColumnLayout(rows, descColumns, unitColumn, quantityColumn, worksheetName);

            QuotaRecommendPanel.Log("Grid parser: rows=" + rows.Count.ToString(CultureInfo.InvariantCulture)
                + ", qtyCol=" + quantityColumn.ToString(CultureInfo.InvariantCulture)
                + ", unitCol=" + unitColumn.ToString(CultureInfo.InvariantCulture)
                + ", descCols=[" + String.Join(",", descColumns.Select(c => c.ToString(CultureInfo.InvariantCulture)).ToArray()) + "]"
                + ", items=" + result.Count.ToString(CultureInfo.InvariantCulture));
            return result;
        }

        // 根据已识别的列布局（多个描述列 + 单位列 + 数量列）逐行构建工程量条目。
        // descColumns 为升序排列的描述列索引，名称由这些列按列序拼接而成，从而支持超过四列的表。
        private static List<ExcelQuantityItem> BuildItemsFromColumnLayout(List<List<CellValue>> rows, List<int> descColumns, int unitColumn, int quantityColumn, string worksheetName)
        {
            List<ExcelQuantityItem> result = new List<ExcelQuantityItem>();
            if (rows == null || descColumns == null || descColumns.Count == 0 || quantityColumn < 0)
            {
                return result;
            }

            int sectionColumn = descColumns[0];
            // 仅当存在两个及以上描述列时，最左侧列才作为"分部/小节"承接（兼容旧的 group 行为）；
            // 只有一个描述列时按行逐条取名、不做承接，与旧的三列表现完全一致。
            bool useCarryDown = descColumns.Count >= 2;
            string[] units = BuildUnitsByRow(rows, unitColumn);
            string currentGroup = "";
            for (int i = 0; i < rows.Count; i++)
            {
                List<CellValue> row = rows[i] ?? new List<CellValue>();
                CellValue quantityCell = GetCell(row, quantityColumn);
                if (quantityCell == null || !IsQuantityLike(quantityCell.Text))
                {
                    continue;
                }

                string section = "";
                if (useCarryDown)
                {
                    section = GetCellText(row, sectionColumn);
                    if (LooksLikeGroupText(section))
                    {
                        currentGroup = section;
                    }
                    else
                    {
                        section = currentGroup;
                    }
                }

                List<string> parts = new List<string>();
                foreach (int c in descColumns)
                {
                    string text = useCarryDown && c == sectionColumn ? section : GetCellText(row, c);
                    if (!String.IsNullOrWhiteSpace(text) && !LooksLikeOrderOrHeader(text) && !LooksLikeUnit(text))
                    {
                        parts.Add(text);
                    }
                }

                string name = CombineQuantityNames(parts);
                if (String.IsNullOrWhiteSpace(name) || LooksLikeOrderOrHeader(name))
                {
                    name = PickQuantityName(row, quantityCell, unitColumn >= 0 ? GetCell(row, unitColumn) : null);
                }

                if (String.IsNullOrWhiteSpace(name) || LooksLikeOrderOrHeader(name))
                {
                    continue;
                }

                string group = useCarryDown ? currentGroup : "";
                if (String.Equals(group, name, StringComparison.Ordinal))
                {
                    group = "";
                }

                ExcelQuantityItem item = new ExcelQuantityItem();
                item.WorksheetName = worksheetName;
                item.RowNumber = quantityCell.RowNumber;
                item.CellAddress = quantityCell.Address;
                item.Name = name;
                item.Unit = i < units.Length ? units[i] : "";
                item.ValueText = quantityCell.Text;
                item.Formula = quantityCell.Formula;
                item.ContextText = name + " " + item.Unit + " " + item.ValueText;
                item.SectionName = String.IsNullOrWhiteSpace(group) ? name : group;
                item.OriginalName = name;
                item.RawRowText = BuildRawRowText(row);
                result.Add(item);
            }

            return result;
        }

        private static int FindQuantityColumn(List<List<CellValue>> rows)
        {
            var candidates = rows.SelectMany(r => r)
                .Where(c => c != null && IsQuantityLike(c.Text) && !LooksLikeUnit(c.Text))
                .GroupBy(c => c.SourceIndex)
                .Select(g => new { Column = g.Key, Count = g.Count() })
                .OrderByDescending(x => x.Count)
                .ThenByDescending(x => x.Column)
                .ToList();
            return candidates.Count == 0 ? -1 : candidates[0].Column;
        }

        private static int FindUnitColumn(List<List<CellValue>> rows, int quantityColumn)
        {
            var candidates = rows.SelectMany(r => r)
                .Where(c => c != null && c.SourceIndex != quantityColumn && LooksLikeUnit(c.Text))
                .GroupBy(c => c.SourceIndex)
                .Select(g => new { Column = g.Key, Count = g.Count() })
                .OrderByDescending(x => x.Column < quantityColumn)
                .ThenByDescending(x => x.Count)
                .ThenBy(x => Math.Abs(quantityColumn - x.Column))
                .ToList();
            return candidates.Count == 0 ? -1 : candidates[0].Column;
        }

        // 找出数量列左侧、整列以"组文本"（中文描述、非序号/单位/数量）为主的所有列，作为名称/描述列。
        // 返回升序列索引；序号列因是数字会被 LooksLikeGroupText 自动排除。支持任意数量的描述列。
        private static List<int> FindDescriptionColumns(List<List<CellValue>> rows, int nameBoundary, int unitColumn, int quantityColumn)
        {
            List<int> columns = new List<int>();
            if (rows == null)
            {
                return columns;
            }

            var byColumn = rows.SelectMany(r => r ?? new List<CellValue>())
                .Where(c => c != null
                    && c.SourceIndex < nameBoundary
                    && c.SourceIndex != unitColumn
                    && c.SourceIndex != quantityColumn)
                .GroupBy(c => c.SourceIndex);

            foreach (var columnGroup in byColumn)
            {
                int groupText = 0;
                int nonEmpty = 0;
                foreach (CellValue cell in columnGroup)
                {
                    string text = (cell.Text ?? "").Trim();
                    if (String.IsNullOrEmpty(text))
                    {
                        continue;
                    }

                    nonEmpty++;
                    if (LooksLikeGroupText(text))
                    {
                        groupText++;
                    }
                }

                if (nonEmpty > 0 && groupText > 0 && groupText * 2 >= nonEmpty)
                {
                    columns.Add(columnGroup.Key);
                }
            }

            columns.Sort();
            return columns;
        }

        private static string[] BuildUnitsByRow(List<List<CellValue>> rows, int unitColumn)
        {
            string[] units = new string[rows.Count];
            for (int i = 0; i < rows.Count; i++)
            {
                CellValue cell = unitColumn >= 0 ? GetCell(rows[i], unitColumn) : null;
                if (cell != null && LooksLikeUnit(cell.Text))
                {
                    units[i] = cell.Text.Trim();
                }
            }

            List<string> knownUnits = units.Where(u => !String.IsNullOrWhiteSpace(u)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (knownUnits.Count == 1)
            {
                for (int i = 0; i < units.Length; i++)
                {
                    units[i] = knownUnits[0];
                }

                return units;
            }

            string lastUnit = "";
            for (int i = 0; i < units.Length; i++)
            {
                if (!String.IsNullOrWhiteSpace(units[i]))
                {
                    lastUnit = units[i];
                }
                else if (!String.IsNullOrWhiteSpace(lastUnit))
                {
                    units[i] = lastUnit;
                }
            }

            string nextUnit = "";
            for (int i = units.Length - 1; i >= 0; i--)
            {
                if (!String.IsNullOrWhiteSpace(units[i]))
                {
                    nextUnit = units[i];
                }
                else if (!String.IsNullOrWhiteSpace(nextUnit))
                {
                    units[i] = nextUnit;
                }
            }

            return units;
        }

        private static CellValue GetCell(List<CellValue> row, int sourceIndex)
        {
            if (row == null)
            {
                return null;
            }

            return row.FirstOrDefault(c => c.SourceIndex == sourceIndex);
        }

        private static string GetCellText(List<CellValue> row, int sourceIndex)
        {
            CellValue cell = GetCell(row, sourceIndex);
            return cell == null ? "" : (cell.Text ?? "").Trim();
        }

        private static string BuildRawRowText(List<CellValue> row)
        {
            return String.Join(" ", (row ?? new List<CellValue>())
                .OrderBy(c => c.SourceIndex)
                .Select(c => c == null ? "" : c.Text)
                .Where(t => !String.IsNullOrWhiteSpace(t))
                .ToArray());
        }

        private static ExcelQuantityItem BuildQuantityItemFromHtmlRow(List<string> cells, int rowNumber, ref string lastGroupName, ref string lastUnit)
        {
            string[] raw = cells.ToArray();
            ExcelQuantityItem four = TryBuildFourColumnClipboardItem(raw, rowNumber, ref lastGroupName, ref lastUnit);
            if (four != null)
            {
                four.WorksheetName = "\u526a\u8d34\u677fHTML";
                return four;
            }

            ExcelQuantityItem three = TryBuildThreeColumnClipboardItem(raw, rowNumber, ref lastUnit);
            if (three != null)
            {
                three.WorksheetName = "\u526a\u8d34\u677fHTML";
                return three;
            }

            return null;
        }

        private static ExcelQuantityItem BuildQuantityItemFromRangeRow(dynamic range, string worksheetName, int relativeRow, int columnCount, ref string lastGroupName, ref string lastUnit)
        {
            ExcelQuantityItem fourColumnItem = TryBuildFourColumnRangeItem(range, worksheetName, relativeRow, columnCount, ref lastGroupName, ref lastUnit);
            if (fourColumnItem != null)
            {
                return fourColumnItem;
            }

            ExcelQuantityItem fixedItem = TryBuildThreeColumnRangeItem(range, worksheetName, relativeRow, columnCount, ref lastUnit);
            if (fixedItem != null)
            {
                return fixedItem;
            }

            List<CellValue> cells = new List<CellValue>();
            for (int c = 1; c <= columnCount; c++)
            {
                dynamic cell = range.Cells[relativeRow, c];
                string text = ExcelValueToText(cell.Value2);
                string formula = "";
                try
                {
                    formula = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture);
                }
                catch
                {
                }

                if (!String.IsNullOrWhiteSpace(text) || !String.IsNullOrWhiteSpace(formula))
                {
                    CellValue value = new CellValue();
                    value.Text = text;
                    value.Formula = formula;
                    value.Address = Convert.ToString(cell.Address(false, false), CultureInfo.InvariantCulture);
                    value.RowNumber = Convert.ToInt32(cell.Row, CultureInfo.InvariantCulture);
                    value.SourceIndex = c;
                    cells.Add(value);
                }
            }

            return BuildQuantityItemFromCells(cells, worksheetName);
        }

        private static ExcelQuantityItem TryBuildFourColumnRangeItem(dynamic range, string worksheetName, int relativeRow, int columnCount, ref string lastGroupName, ref string lastUnit)
        {
            if (columnCount < 4)
            {
                return null;
            }

            dynamic groupCell = range.Cells[relativeRow, 1];
            dynamic detailCell = range.Cells[relativeRow, 2];
            dynamic unitCell = range.Cells[relativeRow, 3];
            dynamic quantityCell = range.Cells[relativeRow, 4];
            string group = ReadCellTextWithMerge(groupCell);
            string detail = ReadCellTextWithMerge(detailCell);
            string unit = ReadCellTextWithMerge(unitCell);
            string quantity = ExcelValueToText(quantityCell.Value2);

            if (!String.IsNullOrWhiteSpace(group) && !LooksLikeOrderOrHeader(group))
            {
                lastGroupName = group;
            }
            else
            {
                group = lastGroupName;
            }

            if (!String.IsNullOrWhiteSpace(unit) && !LooksLikeOrderOrHeader(unit))
            {
                lastUnit = unit;
            }
            else
            {
                unit = lastUnit;
            }

            string name = CombineQuantityName(group, detail);
            if (String.IsNullOrWhiteSpace(name) || String.IsNullOrWhiteSpace(quantity))
            {
                return null;
            }

            if (LooksLikeOrderOrHeader(name) || LooksLikeOrderOrHeader(detail) || !IsQuantityLike(quantity))
            {
                return null;
            }

            ExcelQuantityItem item = new ExcelQuantityItem();
            item.WorksheetName = worksheetName;
            item.RowNumber = Convert.ToInt32(quantityCell.Row, CultureInfo.InvariantCulture);
            item.CellAddress = Convert.ToString(quantityCell.Address(false, false), CultureInfo.InvariantCulture);
            item.Name = name;
            item.Unit = unit;
            item.ValueText = quantity;
            try
            {
                item.Formula = Convert.ToString(quantityCell.Formula, CultureInfo.InvariantCulture);
            }
            catch
            {
                item.Formula = "";
            }
            item.ContextText = name + " " + unit + " " + quantity;
            item.SectionName = group;
            item.OriginalName = name;
            item.RawRowText = item.ContextText;
            item.SkipAiNameNormalization = true;
            return item;
        }

        private static ExcelQuantityItem TryBuildThreeColumnRangeItem(dynamic range, string worksheetName, int relativeRow, int columnCount, ref string lastUnit)
        {
            if (columnCount < 3)
            {
                return null;
            }

            dynamic nameCell = range.Cells[relativeRow, 1];
            dynamic unitCell = range.Cells[relativeRow, 2];
            dynamic quantityCell = range.Cells[relativeRow, 3];
            string name = ReadCellTextWithMerge(nameCell);
            string unit = ReadCellTextWithMerge(unitCell);
            string quantity = ExcelValueToText(quantityCell.Value2);
            string group = TryReadLeftGroupFromRangeRow(range, relativeRow);
            if (!String.IsNullOrWhiteSpace(group) && !LooksLikeOrderOrHeader(group))
            {
                name = CombineQuantityName(group, name);
            }

            if (!String.IsNullOrWhiteSpace(unit) && !LooksLikeOrderOrHeader(unit))
            {
                lastUnit = unit;
            }
            else
            {
                unit = lastUnit;
            }
            if (String.IsNullOrWhiteSpace(name) || String.IsNullOrWhiteSpace(quantity))
            {
                return null;
            }

            if (LooksLikeOrderOrHeader(name) || !IsQuantityLike(quantity))
            {
                return null;
            }

            ExcelQuantityItem item = new ExcelQuantityItem();
            item.WorksheetName = worksheetName;
            item.RowNumber = Convert.ToInt32(quantityCell.Row, CultureInfo.InvariantCulture);
            item.CellAddress = Convert.ToString(quantityCell.Address(false, false), CultureInfo.InvariantCulture);
            item.Name = name;
            item.Unit = unit;
            item.ValueText = quantity;
            try
            {
                item.Formula = Convert.ToString(quantityCell.Formula, CultureInfo.InvariantCulture);
            }
            catch
            {
                item.Formula = "";
            }
            item.ContextText = name + " " + unit + " " + quantity;
            item.SectionName = String.IsNullOrWhiteSpace(group) ? name : group;
            item.OriginalName = name;
            item.RawRowText = item.ContextText;
            item.SkipAiNameNormalization = true;
            return item;
        }

        private static List<string> TryReadActiveSelectionLeftGroups(int expectedRows)
        {
            List<string> result = new List<string>();
            try
            {
                dynamic excel = GetActiveSpreadsheetApplication();
                if (excel == null)
                {
                    return result;
                }

                dynamic range = excel.Selection;
                if (range == null)
                {
                    return result;
                }

                int rowCount = Convert.ToInt32(range.Rows.Count, CultureInfo.InvariantCulture);
                int count = expectedRows > 0 ? Math.Min(expectedRows, rowCount) : rowCount;
                string lastGroup = "";
                for (int row = 1; row <= count; row++)
                {
                    string group = TryReadLeftGroupFromRangeRow(range, row);
                    if (LooksLikeGroupText(group))
                    {
                        lastGroup = group;
                    }
                    else
                    {
                        group = lastGroup;
                    }

                    result.Add(group);
                }
            }
            catch
            {
            }

            return result;
        }

        private static string TryReadLeftGroupFromRangeRow(dynamic range, int relativeRow)
        {
            try
            {
                dynamic firstCell = range.Cells[relativeRow, 1];
                int row = TryReadInt(firstCell, "Row");
                int column = TryReadInt(firstCell, "Column");
                dynamic worksheet = null;
                try
                {
                    worksheet = firstCell.Worksheet;
                }
                catch
                {
                }

                int maxOffset = column > 1 ? Math.Min(8, column - 1) : 8;
                for (int offset = 1; offset <= maxOffset; offset++)
                {
                    string text = "";
                    if (worksheet != null && row > 0 && column > offset)
                    {
                        try
                        {
                            dynamic sheetCell = worksheet.Cells[row, column - offset];
                            text = ReadCellTextWithMerge(sheetCell);
                        }
                        catch
                        {
                            text = "";
                        }
                    }

                    if (!LooksLikeGroupText(text))
                    {
                        try
                        {
                            dynamic offsetCell = firstCell.Offset[0, -offset];
                            text = ReadCellTextWithMerge(offsetCell);
                        }
                        catch
                        {
                            text = "";
                        }
                    }

                    if (LooksLikeGroupText(text))
                    {
                        return text;
                    }
                }
            }
            catch
            {
            }

            return "";
        }

        private static int TryReadInt(dynamic source, string propertyName)
        {
            try
            {
                object value = source.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, null, source, null, CultureInfo.InvariantCulture);
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                try
                {
                    object value = propertyName == "Row" ? source.Row : source.Column;
                    return Convert.ToInt32(value, CultureInfo.InvariantCulture);
                }
                catch
                {
                    return 0;
                }
            }
        }

        private static string ReadCellTextWithMerge(dynamic cell)
        {
            try
            {
                string text = ExcelValueToText(cell.Value2);
                if (!String.IsNullOrWhiteSpace(text))
                {
                    return text;
                }

                dynamic mergeArea = cell.MergeArea;
                if (mergeArea != null)
                {
                    dynamic first = mergeArea.Cells[1, 1];
                    return ExcelValueToText(first.Value2);
                }
            }
            catch
            {
            }

            return "";
        }

        private static ExcelQuantityItem BuildQuantityItemFromTextRow(string[] rawCells, int rowNumber, ref string lastGroupName, ref string lastUnit)
        {
            ExcelQuantityItem fourColumnItem = TryBuildFourColumnClipboardItem(rawCells, rowNumber, ref lastGroupName, ref lastUnit);
            if (fourColumnItem != null)
            {
                return fourColumnItem;
            }

            ExcelQuantityItem fixedItem = TryBuildThreeColumnClipboardItem(rawCells, rowNumber, ref lastUnit);
            if (fixedItem != null)
            {
                return fixedItem;
            }

            List<CellValue> cells = new List<CellValue>();
            for (int i = 0; i < rawCells.Length; i++)
            {
                string text = (rawCells[i] ?? "").Trim();
                if (!String.IsNullOrWhiteSpace(text))
                {
                    CellValue value = new CellValue();
                    value.Text = text;
                    value.Formula = text.StartsWith("=", StringComparison.Ordinal) ? text : "";
                    value.Address = "R" + rowNumber.ToString(CultureInfo.InvariantCulture) + "C" + (i + 1).ToString(CultureInfo.InvariantCulture);
                    value.RowNumber = rowNumber;
                    value.SourceIndex = i + 1;
                    cells.Add(value);
                }
            }

            return BuildQuantityItemFromCells(cells, "\u526a\u8d34\u677f");
        }

        private static ExcelQuantityItem TryBuildFourColumnClipboardItem(string[] rawCells, int rowNumber, ref string lastGroupName, ref string lastUnit)
        {
            if (rawCells == null || rawCells.Length < 4)
            {
                return null;
            }

            string group = NormalizeCellText(rawCells[0]);
            string detail = NormalizeCellText(rawCells[1]);
            string unit = NormalizeCellText(rawCells[2]);
            string quantity = NormalizeCellText(rawCells[3]);

                if (LooksLikeGroupText(group))
            {
                lastGroupName = group;
            }
            else
            {
                group = lastGroupName;
            }

            if (!String.IsNullOrWhiteSpace(unit) && !LooksLikeOrderOrHeader(unit))
            {
                lastUnit = unit;
            }
            else
            {
                unit = lastUnit;
            }

            string name = CombineQuantityName(group, detail);
            if (String.IsNullOrWhiteSpace(name) || String.IsNullOrWhiteSpace(quantity))
            {
                return null;
            }

            if (LooksLikeOrderOrHeader(name) || LooksLikeOrderOrHeader(detail) || !IsQuantityLike(quantity))
            {
                return null;
            }

            ExcelQuantityItem item = new ExcelQuantityItem();
            item.WorksheetName = "\u526a\u8d34\u677f";
            item.RowNumber = rowNumber;
            item.CellAddress = "R" + rowNumber.ToString(CultureInfo.InvariantCulture) + "C4";
            item.Name = name;
            item.Unit = unit;
            item.ValueText = quantity;
            item.Formula = quantity.StartsWith("=", StringComparison.Ordinal) ? quantity : "";
            item.ContextText = name + " " + unit + " " + quantity;
            item.SectionName = group;
            item.OriginalName = name;
            item.RawRowText = item.ContextText;
            item.SkipAiNameNormalization = true;
            return item;
        }

        private static string CombineQuantityName(string group, string detail)
        {
            string left = (group ?? "").Trim();
            string right = (detail ?? "").Trim();
            if (String.IsNullOrWhiteSpace(left))
            {
                return right;
            }

            if (String.IsNullOrWhiteSpace(right))
            {
                return left;
            }

            if (right.IndexOf(left, StringComparison.Ordinal) >= 0)
            {
                return right;
            }

            return left + " " + right;
        }

        // 把多个描述列文本按列序折叠成一个工程量名称，复用 CombineQuantityName 的子串去重/空格拼接逻辑。
        private static string CombineQuantityNames(IEnumerable<string> parts)
        {
            string name = "";
            foreach (string part in parts ?? new string[0])
            {
                if (String.IsNullOrWhiteSpace(part))
                {
                    continue;
                }

                name = CombineQuantityName(name, part);
            }

            return name;
        }

        private static void ApplyActiveLeftGroups(ExcelSelection selection)
        {
            if (selection == null || selection.Items.Count == 0)
            {
                return;
            }

            List<string> groups = TryReadActiveSelectionLeftGroups(selection.Items.Count);
            if (groups.Count == 0)
            {
                return;
            }

            int groupHits = groups.Count(g => LooksLikeGroupText(g));
            string sampleGroup = groups.FirstOrDefault(g => LooksLikeGroupText(g)) ?? "";
            QuotaRecommendPanel.Log("Active left group scan: rows="
                + groups.Count.ToString(CultureInfo.InvariantCulture)
                + ", groups=" + groupHits.ToString(CultureInfo.InvariantCulture)
                + ", sample=" + sampleGroup);

            int count = Math.Min(selection.Items.Count, groups.Count);
            for (int i = 0; i < count; i++)
            {
                ExcelQuantityItem item = selection.Items[i];
                string group = (groups[i] ?? "").Trim();
                if (item == null || String.IsNullOrWhiteSpace(group) || LooksLikeOrderOrHeader(group))
                {
                    continue;
                }

                string section = (item.SectionName ?? "").Trim();
                bool alreadyHasGroup = (item.Name ?? "").IndexOf(group, StringComparison.Ordinal) >= 0;
                bool appearsUngrouped = String.IsNullOrWhiteSpace(section) || String.Equals(section, item.Name, StringComparison.Ordinal);
                if (alreadyHasGroup || !appearsUngrouped)
                {
                    continue;
                }

                item.Name = CombineQuantityName(group, item.Name);
                item.SectionName = group;
                item.ContextText = item.Name + " " + item.Unit + " " + item.ValueText;
            }
        }

        private static void NormalizeSelectionItems(ExcelSelection selection)
        {
            if (selection == null || selection.Items.Count == 0)
            {
                return;
            }

            string lastUnit = "";
            foreach (ExcelQuantityItem item in selection.Items)
            {
                if (item == null)
                {
                    continue;
                }

                item.Name = (item.Name ?? "").Trim();
                item.Unit = (item.Unit ?? "").Trim();
                item.ValueText = (item.ValueText ?? "").Trim();
                if (!String.IsNullOrWhiteSpace(item.Unit))
                {
                    lastUnit = item.Unit;
                }
                else if (!String.IsNullOrWhiteSpace(lastUnit))
                {
                    item.Unit = lastUnit;
                }
            }

            string nextUnit = "";
            for (int i = selection.Items.Count - 1; i >= 0; i--)
            {
                ExcelQuantityItem item = selection.Items[i];
                if (item == null)
                {
                    continue;
                }

                if (!String.IsNullOrWhiteSpace(item.Unit))
                {
                    nextUnit = item.Unit;
                }
                else if (!String.IsNullOrWhiteSpace(nextUnit))
                {
                    item.Unit = nextUnit;
                }
            }

            List<string> knownUnits = selection.Items
                .Where(i => i != null && !String.IsNullOrWhiteSpace(i.Unit))
                .Select(i => i.Unit.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (knownUnits.Count == 1)
            {
                foreach (ExcelQuantityItem item in selection.Items)
                {
                    if (item != null && String.IsNullOrWhiteSpace(item.Unit))
                    {
                        item.Unit = knownUnits[0];
                    }
                }
            }

            foreach (ExcelQuantityItem item in selection.Items)
            {
                if (item != null)
                {
                    item.Name = (item.Name ?? "").Trim();
                    if (String.IsNullOrWhiteSpace(item.OriginalName))
                    {
                        item.OriginalName = item.Name;
                    }
                    if (String.IsNullOrWhiteSpace(item.RawRowText))
                    {
                        item.RawRowText = item.ContextText;
                    }
                    if (String.IsNullOrWhiteSpace(item.RawRowText))
                    {
                        item.RawRowText = item.Name + " " + item.Unit + " " + item.ValueText;
                    }
                    item.ContextText = item.Name + " " + item.Unit + " " + item.ValueText + " " + item.RawRowText;
                }
            }
        }

        private static void LogSelectionSummary(string source, ExcelSelection selection)
        {
            try
            {
                if (selection == null)
                {
                    QuotaRecommendPanel.Log(source + ": no selection");
                    return;
                }

                StringBuilder builder = new StringBuilder();
                int take = Math.Min(5, selection.Items.Count);
                for (int i = 0; i < take; i++)
                {
                    ExcelQuantityItem item = selection.Items[i];
                    if (item == null)
                    {
                        continue;
                    }

                    if (builder.Length > 0)
                    {
                        builder.Append(" | ");
                    }

                    builder.Append(item.Name);
                    builder.Append("/");
                    builder.Append(item.Unit);
                    builder.Append("/");
                    builder.Append(item.ValueText);
                }

                QuotaRecommendPanel.Log(source + ": items=" + selection.Items.Count.ToString(CultureInfo.InvariantCulture) + ", sample=" + builder.ToString());
            }
            catch
            {
            }
        }

        private static ExcelQuantityItem TryBuildThreeColumnClipboardItem(string[] rawCells, int rowNumber, ref string lastUnit)
        {
            if (rawCells == null || rawCells.Length < 3)
            {
                return null;
            }

            string name = NormalizeCellText(rawCells[0]);
            string unit = NormalizeCellText(rawCells[1]);
            string quantity = NormalizeCellText(rawCells[2]);
            if (!String.IsNullOrWhiteSpace(unit) && !LooksLikeOrderOrHeader(unit))
            {
                lastUnit = unit;
            }
            else
            {
                unit = lastUnit;
            }
            if (String.IsNullOrWhiteSpace(name) || String.IsNullOrWhiteSpace(quantity))
            {
                return null;
            }

            if (LooksLikeOrderOrHeader(name) || !IsQuantityLike(quantity))
            {
                return null;
            }

            ExcelQuantityItem item = new ExcelQuantityItem();
            item.WorksheetName = "\u526a\u8d34\u677f";
            item.RowNumber = rowNumber;
            item.CellAddress = "R" + rowNumber.ToString(CultureInfo.InvariantCulture) + "C3";
            item.Name = name;
            item.Unit = unit;
            item.ValueText = quantity;
            item.Formula = quantity.StartsWith("=", StringComparison.Ordinal) ? quantity : "";
            item.ContextText = name + " " + unit + " " + quantity;
            item.SectionName = name;
            item.OriginalName = name;
            item.RawRowText = item.ContextText;
            item.SkipAiNameNormalization = true;
            return item;
        }

        private static ExcelQuantityItem BuildQuantityItemFromCells(List<CellValue> cells, string worksheetName)
        {
            if (cells.Count == 0)
            {
                return null;
            }

            CellValue nameCell = null;
            CellValue unitCell = null;
            CellValue quantityCell = null;
            bool knownLayout = TryPickByKnownLayout(cells, out nameCell, out unitCell, out quantityCell);
            if (!knownLayout)
            {
                quantityCell = cells.LastOrDefault(c => IsQuantityLike(c.Text));
                if (quantityCell == null)
                {
                    return null;
                }

                unitCell = cells.LastOrDefault(c => c != quantityCell && LooksLikeUnit(c.Text));
                string fallbackName = PickQuantityName(cells, quantityCell, unitCell);
                if (!String.IsNullOrWhiteSpace(fallbackName))
                {
                    nameCell = cells.FirstOrDefault(c => c.Text == fallbackName);
                }
            }

            if (quantityCell == null)
            {
                return null;
            }

            string name = nameCell == null ? PickQuantityName(cells, quantityCell, unitCell) : nameCell.Text;
            if (String.IsNullOrWhiteSpace(name))
            {
                return null;
            }

            ExcelQuantityItem item = new ExcelQuantityItem();
            item.WorksheetName = worksheetName;
            item.RowNumber = quantityCell.RowNumber;
            item.CellAddress = quantityCell.Address;
            item.Name = name;
            item.Unit = unitCell == null ? "" : unitCell.Text;
            item.ValueText = quantityCell.Text;
            item.Formula = quantityCell.Formula;
            item.ContextText = String.Join(" ", cells.Select(c => c.Text).Where(t => !String.IsNullOrWhiteSpace(t)).ToArray());
            item.SectionName = GuessSectionName(cells);
            item.OriginalName = name;
            item.RawRowText = item.ContextText;
            item.SkipAiNameNormalization = knownLayout;
            return item;
        }

        private static bool TryPickByKnownLayout(List<CellValue> cells, out CellValue nameCell, out CellValue unitCell, out CellValue quantityCell)
        {
            nameCell = null;
            unitCell = null;
            quantityCell = null;

            int maxIndex = cells.Max(c => c.SourceIndex);
            int[][] layouts = new int[][]
            {
                new int[] { 2, 6, 7 },
                new int[] { 1, 5, 6 },
                new int[] { 1, 2, 3 }
            };

            foreach (int[] layout in layouts)
            {
                if (maxIndex < layout[2])
                {
                    continue;
                }

                CellValue n = cells.FirstOrDefault(c => c.SourceIndex == layout[0]);
                CellValue u = cells.FirstOrDefault(c => c.SourceIndex == layout[1]);
                CellValue q = cells.FirstOrDefault(c => c.SourceIndex == layout[2]);
                if (n != null && q != null && !IsNumberLike(n.Text) && IsQuantityLike(q.Text))
                {
                    nameCell = n;
                    unitCell = u != null && LooksLikeUnit(u.Text) ? u : cells.LastOrDefault(c => c.SourceIndex < q.SourceIndex && LooksLikeUnit(c.Text));
                    quantityCell = q;
                    return true;
                }
            }

            return false;
        }

        private static string PickQuantityName(List<CellValue> cells, CellValue quantityCell, CellValue unitCell)
        {
            IEnumerable<CellValue> candidates = cells.Where(c => c != quantityCell && c != unitCell && !IsNumberLike(c.Text) && !LooksLikeUnit(c.Text));
            CellValue best = candidates
                .Where(c => !LooksLikeOrderOrHeader(c.Text))
                .OrderByDescending(c => CountChineseLikeChars(c.Text))
                .ThenByDescending(c => (c.Text ?? "").Length)
                .FirstOrDefault();
            return best == null ? "" : best.Text;
        }

        private static string GuessSectionName(List<CellValue> cells)
        {
            foreach (CellValue cell in cells)
            {
                string text = cell.Text ?? "";
                if (text.Length >= 2 && text.Length <= 20 && !IsNumberLike(text) && !LooksLikeUnit(text))
                {
                    return text;
                }
            }

            return "";
        }

        private static object GetActiveSpreadsheetApplication()
        {
            string[] progIds = new string[] { "ket.Application", "KET.Application", "et.Application", "Excel.Application" };
            foreach (string progId in progIds)
            {
                try
                {
                    object app = Marshal.GetActiveObject(progId);
                    if (app != null)
                    {
                        return app;
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        private static string ExcelValueToText(object value)
        {
            if (value == null)
            {
                return "";
            }
            if (value is double || value is float)
            {
                return NormalizeCellText(Convert.ToDouble(value, CultureInfo.InvariantCulture).ToString("0.########", CultureInfo.InvariantCulture));
            }
            if (value is decimal)
            {
                return NormalizeCellText(((decimal)value).ToString("0.########", CultureInfo.InvariantCulture));
            }
            return NormalizeCellText(Convert.ToString(value, CultureInfo.CurrentCulture));
        }

        private static string NormalizeCellText(string text)
        {
            return Regex.Replace((text ?? "").Replace("\r", " ").Replace("\n", " "), @"\s+", " ").Trim();
        }

        private static bool LooksLikeUnit(string text)
        {
            string unit = NormalizeRawUnit(text);
            string[] units = new string[] { "m", "m2", "m3", "kg", "t", "处", "个", "座", "项", "根", "孔", "环", "组" };
            return units.Contains(unit);
        }

        private static bool UnitCompatible(string left, string right)
        {
            string l = NormalizeUnit(left);
            string r = NormalizeUnit(right);
            return !String.IsNullOrEmpty(l) && !String.IsNullOrEmpty(r) && (l == r || l.EndsWith(r, StringComparison.Ordinal) || r.EndsWith(l, StringComparison.Ordinal));
        }

        internal static bool UnitCompatibleForIndex(string left, string right)
        {
            return UnitCompatible(left, right);
        }

        private static string NormalizeUnit(string unit)
        {
            return NormalizeRawUnit(unit)
                .Replace("100", "")
                .Replace("10", "")
                .Replace("㎡", "m2")
                .Replace("㎥", "m3")
                .Replace("m²", "m2")
                .Replace("m³", "m3");
        }

        private static bool IsNumberLike(string text)
        {
            decimal value;
            return Decimal.TryParse((text ?? "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out value)
                || Decimal.TryParse((text ?? "").Trim(), NumberStyles.Float, CultureInfo.CurrentCulture, out value);
        }

        private static bool IsQuantityLike(string text)
        {
            if (IsNumberLike(text))
            {
                return true;
            }

            string value = (text ?? "").Trim();
            if (CountChineseLikeChars(value) > 0)
            {
                return false;
            }

            bool hasDigit = value.Any(Char.IsDigit);
            bool hasOperator = value.IndexOfAny(new char[] { '+', '-', '*', '/', '×', '(', ')', '（', '）' }) >= 0;
            return hasDigit && hasOperator;
        }

        private static bool LooksLikeOrderOrHeader(string text)
        {
            string value = Normalize(text);
            return value == "\u5e8f\u53f7"
                || value == "\u7f16\u53f7"
                || value == "\u5355\u4f4d"
                || value == "\u5de5\u7a0b\u91cf"
                || value == "\u5de5\u7a0b\u9879\u76ee"
                || IsNumberLike(value);
        }

        private static bool LooksLikeGroupText(string text)
        {
            return !String.IsNullOrWhiteSpace(text)
                && !LooksLikeOrderOrHeader(text)
                && !LooksLikeUnit(text)
                && !IsQuantityLike(text);
        }

        private static int CountChineseLikeChars(string text)
        {
            int count = 0;
            foreach (char ch in text ?? "")
            {
                if (ch >= 0x4e00 && ch <= 0x9fff)
                {
                    count++;
                }
            }
            return count;
        }

        private static IEnumerable<string> Tokenize(string text)
        {
            foreach (string token in Normalize(text).Split(new char[] { ' ', '/', ',', ';', '\t', '(', ')', '[', ']', '（', '）' }, StringSplitOptions.RemoveEmptyEntries))
            {
                if (token.Length >= 2 && !IsNumberLike(token))
                {
                    yield return token;
                }
            }
        }

        private static string Normalize(string text)
        {
            return (text ?? "").Replace("\r", " ").Replace("\n", " ").Trim().ToLowerInvariant();
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

        private sealed class ScoredRecord
        {
            public LearningRecord Record;
            public int Score;
        }

        private sealed class RecommendationBatchStats
        {
            public int MappingHits;
            public int IndexSearches;
            public int EmptyRows;
            public int CacheHits;
            public int AiQueued;
        }

        private struct CarryCell
        {
            public string Text;
            public int RemainingRows;
        }
    }

    // 单元格快照：保留文本/公式/地址/行号/列索引（列索引从 1 开始，0 为合并表头注入的左侧分组列）。
    internal sealed class CellValue
    {
        public string Text;
        public string Formula;
        public string Address;
        public int RowNumber;
        public int SourceIndex;
    }

    internal sealed class ExcelSelection
    {
        public string WorkbookPath;
        public string WorksheetName;
        public readonly List<ExcelQuantityItem> Items = new List<ExcelQuantityItem>();
        // 解析时的原始单元格网格（按行、按列保留结构），供 AI 列映射兜底重新判断列角色。
        public List<List<CellValue>> RawRows;
    }

    internal sealed class ExcelQuantityItem
    {
        public string WorksheetName;
        public int RowNumber;
        public string CellAddress;
        public string Name;
        public string OriginalName;
        public string AiName;
        public int AiNameConfidence;
        public string AiNameReason;
        public bool SkipAiNameNormalization;
        public string SectionName;
        public string Unit;
        public string ValueText;
        public string Formula;
        public string ContextText;
        public string RawRowText;
    }

    internal sealed class RecommendationRow
    {
        public ExcelQuantityItem Item;
        public LearningRecord Record;
        public string QuotaCode;
        public string QuotaName;
        public string QuotaUnit;
        public string BookCode;
        public string Specialty;
        public double BasePrice;
        public string WorkContent;
        public string ConvertedValueText;
        public int Score;
        public string Reason;
        public string Source;
        public string BoxId;
        public string TargetKind;
        public int GridRowIndex;
        public string AiRowId;
        public bool AiPending;
        public List<AiQuotaCandidate> AiCandidates;
        public List<AiMappingCandidate> AiMappingCandidates;
    }

    internal sealed class AiQuotaCandidate
    {
        public IndexQuota Quota;
        public int LocalScore;
    }

    internal sealed class AiMappingCandidate
    {
        public string BoxId;
        public int LocalScore;
        public string SampleNames;
        public List<MappingTarget> Targets;

        public List<RecommendationRow> ToRecommendations(ExcelQuantityItem item, int score, string reason)
        {
            return (Targets ?? new List<MappingTarget>())
                .OrderBy(t => MappingStore.TargetSortRank(t.TargetKind, t.Code))
                .ThenBy(t => t.Code ?? "", StringComparer.OrdinalIgnoreCase)
                .Select(t =>
                {
                    RecommendationRow row = t.ToRecommendation(item, score, BoxId);
                    row.Reason = String.IsNullOrWhiteSpace(reason) ? "DeepSeek\u5bf9\u5e94\u6846\u68c0\u6d4b" : "DeepSeek\u5bf9\u5e94\u6846\u68c0\u6d4b\uff1a" + reason;
                    return row;
                })
                .ToList();
        }
    }

    internal sealed class DeepSeekRequestRow
    {
        public string RowId;
        public ExcelQuantityItem Item;
        public List<AiQuotaCandidate> Candidates;
        public List<AiMappingCandidate> MappingCandidates;
    }

    internal sealed class DeepSeekNameRequestRow
    {
        public string RowId;
        public ExcelQuantityItem Item;
    }

    internal sealed class AiPendingRecommendation
    {
        public RecommendationRow Row;
        public int GridRowIndex;
        public DeepSeekRequestRow Request;
        public EntryScope Scope;
    }

    internal sealed class DeepSeekSelection
    {
        public string RowId;
        public string BoxId;
        public string SelectedCode;
        public int Confidence;
        public string Reason;
        public string ErrorText;
    }

    internal sealed class DeepSeekNameResult
    {
        public string RowId;
        public string QuantityName;
        public int Confidence;
        public string Reason;
    }

    internal sealed class DeepSeekMappingSelection
    {
        public string BoxId;
        public int Confidence;
        public string Reason;
    }

    // AI 识别出的工程量表列布局：哪些列组成名称、哪列单位、哪列数量（列索引从 1 开始）。
    internal sealed class DeepSeekColumnLayout
    {
        public int[] NameColumns;
        public int UnitColumn;
        public int QuantityColumn;
        public int Confidence;
    }

    internal sealed class DeepSeekSettings
    {
        public bool Enabled;
        public string ApiKey;
        public string Model = "deepseek-v4-pro";
        public string BaseUrl = "https://api.deepseek.com";
        public int TimeoutSeconds = 20;
        public int MaxRowsPerBatch = 8;
        public int MaxCandidatesPerRow = 12;
        public int LocalHighScore = 80;
        public int DisplayConfidence = 65;
        public int AutoCheckConfidence = 85;
        public bool EnableNameNormalization = true;
        public bool EnableMappingDetection = true;
        public bool EnableQuotaRecommendation = true;
        public bool EnableColumnDetection = true;

        public bool IsAvailable
        {
            get { return !String.IsNullOrWhiteSpace(ApiKey); }
        }

        public bool CanNormalizeNames
        {
            get { return IsAvailable && EnableNameNormalization; }
        }

        public bool CanDetectColumns
        {
            get { return IsAvailable && EnableColumnDetection; }
        }

        public bool CanRecommendQuota
        {
            get { return IsAvailable && EnableQuotaRecommendation; }
        }

        public bool CanDetectMapping
        {
            get { return IsAvailable; }
        }

        public DeepSeekSettings Copy()
        {
            return new DeepSeekSettings
            {
                Enabled = Enabled,
                ApiKey = ApiKey,
                Model = Model,
                BaseUrl = BaseUrl,
                TimeoutSeconds = TimeoutSeconds,
                MaxRowsPerBatch = MaxRowsPerBatch,
                MaxCandidatesPerRow = MaxCandidatesPerRow,
                LocalHighScore = LocalHighScore,
                DisplayConfidence = DisplayConfidence,
                AutoCheckConfidence = AutoCheckConfidence,
                EnableNameNormalization = EnableNameNormalization,
                EnableMappingDetection = EnableMappingDetection,
                EnableQuotaRecommendation = EnableQuotaRecommendation,
                EnableColumnDetection = EnableColumnDetection
            };
        }

        public static DeepSeekSettings Load()
        {
            DeepSeekSettings settings = new DeepSeekSettings();
            string path = ConfigPath();
            if (!File.Exists(path))
            {
                return settings;
            }

            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                Dictionary<string, object> values = serializer.DeserializeObject(File.ReadAllText(path, Encoding.UTF8)) as Dictionary<string, object>;
                if (values == null)
                {
                    return settings;
                }

                settings.Enabled = ReadBool(values, "enabled", false);
                settings.ApiKey = ReadString(values, "api_key", "");
                settings.Model = ReadString(values, "model", settings.Model);
                settings.BaseUrl = ReadString(values, "base_url", settings.BaseUrl).TrimEnd('/');
                settings.TimeoutSeconds = Clamp(ReadInt(values, "timeout_seconds", settings.TimeoutSeconds), 2, 120);
                settings.MaxRowsPerBatch = Clamp(ReadInt(values, "max_rows_per_batch", settings.MaxRowsPerBatch), 1, 20);
                settings.MaxCandidatesPerRow = Clamp(ReadInt(values, "max_candidates_per_row", settings.MaxCandidatesPerRow), 3, 20);
                settings.LocalHighScore = Clamp(ReadInt(values, "local_high_score", settings.LocalHighScore), 60, 120);
                settings.DisplayConfidence = Clamp(ReadInt(values, "display_confidence", settings.DisplayConfidence), 1, 100);
                settings.AutoCheckConfidence = Clamp(ReadInt(values, "auto_check_confidence", settings.AutoCheckConfidence), 1, 100);
                settings.EnableNameNormalization = ReadBool(values, "enable_name_normalization", settings.EnableNameNormalization);
                settings.EnableMappingDetection = ReadBool(values, "enable_mapping_detection", settings.EnableMappingDetection);
                settings.EnableQuotaRecommendation = ReadBool(values, "enable_quota_recommendation", settings.EnableQuotaRecommendation);
                settings.EnableColumnDetection = ReadBool(values, "enable_column_detection", settings.EnableColumnDetection);
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("Load DeepSeek settings failed: " + ex.Message);
                settings.Enabled = false;
            }

            return settings;
        }

        public void Save()
        {
            Directory.CreateDirectory(LearningStore.FindDataDir());
            string path = ConfigPath();
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("{");
            AppendJson(builder, "enabled", Enabled ? "true" : "false", false, true);
            AppendJson(builder, "api_key", ApiKey ?? "", true, true);
            AppendJson(builder, "model", String.IsNullOrWhiteSpace(Model) ? "deepseek-v4-pro" : Model, true, true);
            AppendJson(builder, "base_url", String.IsNullOrWhiteSpace(BaseUrl) ? "https://api.deepseek.com" : BaseUrl, true, true);
            AppendJson(builder, "timeout_seconds", TimeoutSeconds.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "max_rows_per_batch", MaxRowsPerBatch.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "max_candidates_per_row", MaxCandidatesPerRow.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "local_high_score", LocalHighScore.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "display_confidence", DisplayConfidence.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "auto_check_confidence", AutoCheckConfidence.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "enable_name_normalization", EnableNameNormalization ? "true" : "false", false, true);
            AppendJson(builder, "enable_mapping_detection", EnableMappingDetection ? "true" : "false", false, true);
            AppendJson(builder, "enable_quota_recommendation", EnableQuotaRecommendation ? "true" : "false", false, true);
            AppendJson(builder, "enable_column_detection", EnableColumnDetection ? "true" : "false", false, false);
            builder.AppendLine("}");
            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        public static string ConfigPath()
        {
            return Path.Combine(LearningStore.FindDataDir(), "deepseek-settings.json");
        }

        private static void AppendJson(StringBuilder builder, string key, string value, bool quoteValue, bool comma)
        {
            builder.Append("  \"").Append(EscapeJson(key)).Append("\": ");
            if (quoteValue)
            {
                builder.Append("\"").Append(EscapeJson(value)).Append("\"");
            }
            else
            {
                builder.Append(value);
            }
            if (comma)
            {
                builder.Append(",");
            }
            builder.AppendLine();
        }

        private static string EscapeJson(string value)
        {
            StringBuilder builder = new StringBuilder();
            foreach (char ch in value ?? "")
            {
                switch (ch)
                {
                    case '\\':
                        builder.Append("\\\\");
                        break;
                    case '"':
                        builder.Append("\\\"");
                        break;
                    case '\r':
                        builder.Append("\\r");
                        break;
                    case '\n':
                        builder.Append("\\n");
                        break;
                    case '\t':
                        builder.Append("\\t");
                        break;
                    default:
                        builder.Append(ch);
                        break;
                }
            }
            return builder.ToString();
        }

        private static string ReadString(Dictionary<string, object> values, string key, string fallback)
        {
            object value;
            return values.TryGetValue(key, out value) && value != null ? Convert.ToString(value, CultureInfo.InvariantCulture) : fallback;
        }

        private static int ReadInt(Dictionary<string, object> values, string key, int fallback)
        {
            object value;
            if (!values.TryGetValue(key, out value) || value == null)
            {
                return fallback;
            }

            int parsed;
            return Int32.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) ? parsed : fallback;
        }

        private static bool ReadBool(Dictionary<string, object> values, string key, bool fallback)
        {
            object value;
            if (!values.TryGetValue(key, out value) || value == null)
            {
                return fallback;
            }

            if (value is bool)
            {
                return (bool)value;
            }

            bool parsed;
            return Boolean.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), out parsed) ? parsed : fallback;
        }

        private static int Clamp(int value, int min, int max)
        {
            return Math.Max(min, Math.Min(max, value));
        }
    }

    internal sealed class DeepSeekSettingsDialog : Form
    {
        private readonly CheckBox enabledCheck;
        private readonly CheckBox normalizeNameCheck;
        private readonly CheckBox mappingDetectCheck;
        private readonly CheckBox recommendQuotaCheck;
        private readonly TextBox apiKeyText;
        private readonly CheckBox showKeyCheck;
        private readonly TextBox modelText;
        private readonly TextBox baseUrlText;
        private readonly NumericUpDown timeoutInput;
        private readonly NumericUpDown batchInput;
        private readonly NumericUpDown candidatesInput;
        private readonly NumericUpDown localHighScoreInput;
        private readonly NumericUpDown displayConfidenceInput;
        private readonly NumericUpDown autoCheckConfidenceInput;

        public DeepSeekSettings Settings { get; private set; }

        public DeepSeekSettingsDialog(DeepSeekSettings current)
        {
            Settings = (current ?? new DeepSeekSettings()).Copy();
            Text = "DeepSeek AI\u8bbe\u7f6e";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(560, 300);

            enabledCheck = new CheckBox { Checked = true };
            normalizeNameCheck = new CheckBox { Checked = true };
            mappingDetectCheck = new CheckBox { Checked = true };
            recommendQuotaCheck = new CheckBox { Checked = true };
            batchInput = HiddenNumber(Settings.MaxRowsPerBatch, 1, 20);
            localHighScoreInput = HiddenNumber(Settings.LocalHighScore, 60, 120);
            displayConfidenceInput = HiddenNumber(Settings.DisplayConfidence, 1, 100);

            AddLabel("API Key", 24, 28, 150);
            apiKeyText = new TextBox();
            apiKeyText.Left = 180;
            apiKeyText.Top = 24;
            apiKeyText.Width = 280;
            apiKeyText.UseSystemPasswordChar = true;
            apiKeyText.Text = Settings.ApiKey ?? "";
            Controls.Add(apiKeyText);

            showKeyCheck = new CheckBox();
            showKeyCheck.Left = 470;
            showKeyCheck.Top = 26;
            showKeyCheck.Width = 60;
            showKeyCheck.Text = "\u663e\u793a";
            showKeyCheck.CheckedChanged += delegate { apiKeyText.UseSystemPasswordChar = !showKeyCheck.Checked; };
            Controls.Add(showKeyCheck);

            AddLabel("\u6a21\u578b", 24, 64, 150);
            modelText = AddTextBox(Settings.Model, 180, 60, 280);

            AddLabel("\u63a5\u53e3\u5730\u5740", 24, 100, 150);
            baseUrlText = AddTextBox(Settings.BaseUrl, 180, 96, 280);

            AddLabel("\u8d85\u65f6\u79d2\u6570", 24, 136, 150);
            timeoutInput = AddNumber(Settings.TimeoutSeconds, 180, 132, 2, 120);

            AddLabel("\u6bcf\u884c\u5019\u9009\u6570", 24, 172, 150);
            candidatesInput = AddNumber(Settings.MaxCandidatesPerRow, 180, 168, 3, 20);

            AddLabel("AI\u81ea\u52a8\u52fe\u9009\u7f6e\u4fe1\u5ea6", 24, 208, 150);
            autoCheckConfidenceInput = AddNumber(Settings.AutoCheckConfidence, 180, 204, 1, 100);

            Button saveButton = new Button();
            saveButton.Text = "\u4fdd\u5b58";
            saveButton.Left = 366;
            saveButton.Top = 252;
            saveButton.Width = 80;
            saveButton.DialogResult = DialogResult.None;
            saveButton.Click += delegate { SaveSettings(); };
            Controls.Add(saveButton);

            Button cancelButton = new Button();
            cancelButton.Text = "\u53d6\u6d88";
            cancelButton.Left = 456;
            cancelButton.Top = 252;
            cancelButton.Width = 80;
            cancelButton.DialogResult = DialogResult.Cancel;
            Controls.Add(cancelButton);

            AcceptButton = saveButton;
            CancelButton = cancelButton;
        }

        private void SaveSettings()
        {
            string model = modelText.Text.Trim();
            string baseUrl = baseUrlText.Text.Trim();
            string apiKey = apiKeyText.Text.Trim();

            if (String.IsNullOrWhiteSpace(apiKey))
            {
                MessageBox.Show(this, "\u8bf7\u586b\u5199 DeepSeek API Key\u3002", "DeepSeek AI\u8bbe\u7f6e", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (String.IsNullOrWhiteSpace(model))
            {
                model = "deepseek-v4-pro";
            }
            if (String.IsNullOrWhiteSpace(baseUrl))
            {
                baseUrl = "https://api.deepseek.com";
            }

            DeepSeekSettings updated = new DeepSeekSettings();
            updated.Enabled = !String.IsNullOrWhiteSpace(apiKey);
            updated.ApiKey = apiKey;
            updated.Model = model;
            updated.BaseUrl = baseUrl.TrimEnd('/');
            updated.TimeoutSeconds = Convert.ToInt32(timeoutInput.Value);
            updated.MaxRowsPerBatch = Convert.ToInt32(batchInput.Value);
            updated.MaxCandidatesPerRow = Convert.ToInt32(candidatesInput.Value);
            updated.LocalHighScore = Convert.ToInt32(localHighScoreInput.Value);
            updated.DisplayConfidence = Convert.ToInt32(displayConfidenceInput.Value);
            updated.AutoCheckConfidence = Convert.ToInt32(autoCheckConfidenceInput.Value);
            updated.EnableNameNormalization = true;
            updated.EnableMappingDetection = true;
            updated.EnableQuotaRecommendation = true;

            try
            {
                updated.Save();
                Settings = updated;
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "\u4fdd\u5b58 DeepSeek \u8bbe\u7f6e\u5931\u8d25\uff1a" + ex.Message, "DeepSeek AI\u8bbe\u7f6e", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private TextBox AddTextBox(string text, int left, int top, int width)
        {
            TextBox box = new TextBox();
            box.Left = left;
            box.Top = top;
            box.Width = width;
            box.Text = text ?? "";
            Controls.Add(box);
            return box;
        }

        private NumericUpDown AddNumber(int value, int left, int top, int min, int max)
        {
            NumericUpDown input = new NumericUpDown();
            input.Left = left;
            input.Top = top;
            input.Width = 120;
            input.Minimum = min;
            input.Maximum = max;
            input.Value = Math.Max(min, Math.Min(max, value));
            Controls.Add(input);
            return input;
        }

        private NumericUpDown HiddenNumber(int value, int min, int max)
        {
            NumericUpDown input = new NumericUpDown();
            input.Minimum = min;
            input.Maximum = max;
            input.Value = Math.Max(min, Math.Min(max, value));
            return input;
        }

        private void AddLabel(string text, int left, int top, int width)
        {
            Label label = new Label();
            label.Text = text;
            label.Left = left;
            label.Top = top + 4;
            label.Width = width;
            Controls.Add(label);
        }
    }

    internal sealed class DeepSeekClient
    {
        private readonly DeepSeekSettings settings;
        private readonly JavaScriptSerializer serializer = new JavaScriptSerializer();

        public DeepSeekClient(DeepSeekSettings deepSeekSettings)
        {
            settings = deepSeekSettings;
            serializer.MaxJsonLength = 1024 * 1024 * 4;
        }

        public List<DeepSeekSelection> Rank(List<DeepSeekRequestRow> rows)
        {
            if (rows == null || rows.Count == 0)
            {
                return new List<DeepSeekSelection>();
            }

            return ParseResponse(SendRequest(BuildUnifiedRequestJson(rows)));
        }

        public List<DeepSeekNameResult> NormalizeQuantityNames(List<DeepSeekNameRequestRow> rows)
        {
            if (rows == null || rows.Count == 0)
            {
                return new List<DeepSeekNameResult>();
            }

            return ParseNameResponse(SendRequest(BuildNameRequestJson(rows)));
        }

        // 对整张工程量表只调用一次，让模型判断列结构（哪些列组成名称、哪列单位、哪列数量）。
        public DeepSeekColumnLayout DetectColumnLayout(List<List<CellValue>> rows)
        {
            if (rows == null || rows.Count == 0)
            {
                return null;
            }

            return ParseColumnLayoutResponse(SendRequest(BuildColumnLayoutRequestJson(rows)));
        }

        public DeepSeekMappingSelection SelectMappingBox(ExcelQuantityItem item, List<AiMappingCandidate> candidates)
        {
            if (item == null || candidates == null || candidates.Count == 0)
            {
                return null;
            }

            return ParseMappingResponse(SendRequest(BuildMappingRequestJson(item, candidates)));
        }

        private string SendRequest(string requestJson)
        {
            ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol | (SecurityProtocolType)3072;
            string endpoint = settings.BaseUrl.TrimEnd('/');
            if (!endpoint.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase))
            {
                endpoint += "/chat/completions";
            }

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(endpoint);
            request.Method = "POST";
            request.ContentType = "application/json";
            request.Accept = "application/json";
            request.Headers["Authorization"] = "Bearer " + settings.ApiKey;
            request.Timeout = settings.TimeoutSeconds * 1000;
            request.ReadWriteTimeout = settings.TimeoutSeconds * 1000;

            byte[] payload = Encoding.UTF8.GetBytes(requestJson);
            request.ContentLength = payload.Length;
            using (Stream stream = request.GetRequestStream())
            {
                stream.Write(payload, 0, payload.Length);
            }

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
            {
                return reader.ReadToEnd();
            }
        }

        private string BuildRequestJson(List<DeepSeekRequestRow> rows)
        {
            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["model"] = settings.Model;
            payload["stream"] = false;
            payload["temperature"] = 0.1;
            payload["max_tokens"] = 1200;
            payload["response_format"] = new Dictionary<string, object> { { "type", "json_object" } };

            List<object> messages = new List<object>();
            messages.Add(new Dictionary<string, object>
            {
                { "role", "system" },
                { "content", "你是铁路工程定额推荐助手。只能从用户给出的本地候选定额中选择，不能编造编号，不能选择材料。必须输出严格 json，格式为 {\"results\":[{\"row_id\":\"r0\",\"selected_code\":\"PY-738\",\"confidence\":90,\"reason\":\"简短理由\"}]}。如果没有可靠候选，selected_code 置空，confidence 置 0。" }
            });
            messages.Add(new Dictionary<string, object>
            {
                { "role", "user" },
                { "content", BuildUserPrompt(rows) }
            });
            payload["messages"] = messages;
            return serializer.Serialize(payload);
        }

        private string BuildNameRequestJson(List<DeepSeekNameRequestRow> rows)
        {
            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["model"] = settings.Model;
            payload["stream"] = false;
            payload["temperature"] = 0.1;
            payload["max_tokens"] = 1400;
            payload["response_format"] = new Dictionary<string, object> { { "type", "json_object" } };

            List<object> messages = new List<object>();
            messages.Add(new Dictionary<string, object>
            {
                { "role", "system" },
                { "content", "You normalize Chinese railway construction quantity names. Return strict JSON: {\"results\":[{\"row_id\":\"n0\",\"quantity_name\":\"标准工程量名称\",\"confidence\":90,\"reason\":\"short reason\"}]}. The quantity_name must be concise, contain the engineering object and key specification, and must not include serial numbers, units, quantities, formulas, or prices." }
            });
            messages.Add(new Dictionary<string, object>
            {
                { "role", "user" },
                { "content", BuildNameUserPrompt(rows) }
            });
            payload["messages"] = messages;
            return serializer.Serialize(payload);
        }

        private string BuildNameUserPrompt(List<DeepSeekNameRequestRow> rows)
        {
            List<object> requestRows = new List<object>();
            foreach (DeepSeekNameRequestRow row in rows)
            {
                ExcelQuantityItem item = row.Item;
                requestRows.Add(new Dictionary<string, object>
                {
                    { "row_id", row.RowId ?? "" },
                    { "local_name", item == null ? "" : item.OriginalName ?? item.Name ?? "" },
                    { "section_name", item == null ? "" : item.SectionName ?? "" },
                    { "unit", item == null ? "" : item.Unit ?? "" },
                    { "quantity_value", item == null ? "" : item.ValueText ?? "" },
                    { "raw_row_text", item == null ? "" : Truncate(item.RawRowText, 300) },
                    { "context_text", item == null ? "" : Truncate(item.ContextText, 300) }
                });
            }

            Dictionary<string, object> body = new Dictionary<string, object>();
            body["task"] = "For each row, summarize all possible quantity-name columns into one standard Chinese engineering quantity name.";
            body["rules"] = new string[]
            {
                "Use all name-like columns before the unit/quantity as context.",
                "Do not include unit, quantity, formula, serial number, or price.",
                "Keep key material/model/specification when it changes the quota selection.",
                "If the row is ambiguous, return the best concise name with lower confidence."
            };
            body["rows"] = requestRows;
            return serializer.Serialize(body);
        }

        private string BuildMappingRequestJson(ExcelQuantityItem item, List<AiMappingCandidate> candidates)
        {
            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["model"] = settings.Model;
            payload["stream"] = false;
            payload["temperature"] = 0.1;
            payload["max_tokens"] = 1000;
            payload["response_format"] = new Dictionary<string, object> { { "type", "json_object" } };

            List<object> messages = new List<object>();
            messages.Add(new Dictionary<string, object>
            {
                { "role", "system" },
                { "content", "You match a normalized railway engineering quantity name to an existing human-corrected mapping box. Only choose from candidates. Return strict JSON: {\"box_id\":\"box-123\",\"confidence\":90,\"reason\":\"short reason\"}. If no candidate is reliable, return empty box_id and confidence 0." }
            });
            messages.Add(new Dictionary<string, object>
            {
                { "role", "user" },
                { "content", BuildMappingUserPrompt(item, candidates) }
            });
            payload["messages"] = messages;
            return serializer.Serialize(payload);
        }

        private string BuildMappingUserPrompt(ExcelQuantityItem item, List<AiMappingCandidate> candidates)
        {
            List<object> candidateRows = new List<object>();
            foreach (AiMappingCandidate candidate in candidates)
            {
                candidateRows.Add(new Dictionary<string, object>
                {
                    { "box_id", candidate.BoxId ?? "" },
                    { "sample_quantity_names", Truncate(candidate.SampleNames, 260) },
                    { "targets", String.Join(" + ", (candidate.Targets ?? new List<MappingTarget>()).Select(t => (t.Code ?? "") + " " + (t.Name ?? "")).ToArray()) },
                    { "local_score", candidate.LocalScore }
                });
            }

            Dictionary<string, object> body = new Dictionary<string, object>();
            body["quantity_name"] = item.Name ?? "";
            body["original_name"] = item.OriginalName ?? "";
            body["unit"] = item.Unit ?? "";
            body["raw_row_text"] = Truncate(item.RawRowText, 300);
            body["rules"] = new string[]
            {
                "Prefer human-corrected mapping boxes when the engineering meaning is equivalent.",
                "Do not choose a box only because of a single generic word.",
                "Steel and concrete terms must not be confused.",
                "If unsure, return empty box_id."
            };
            body["candidates"] = candidateRows;
            return serializer.Serialize(body);
        }

        private string BuildUserPrompt(List<DeepSeekRequestRow> rows)
        {
            List<object> requestRows = new List<object>();
            foreach (DeepSeekRequestRow row in rows)
            {
                List<object> candidates = new List<object>();
                foreach (AiQuotaCandidate candidate in row.Candidates ?? new List<AiQuotaCandidate>())
                {
                    if (candidate == null || candidate.Quota == null)
                    {
                        continue;
                    }

                    candidates.Add(new Dictionary<string, object>
                    {
                        { "code", candidate.Quota.QuotaCode ?? "" },
                        { "name", candidate.Quota.QuotaName ?? "" },
                        { "unit", candidate.Quota.QuotaUnit ?? "" },
                        { "work_content", Truncate(candidate.Quota.WorkContent, 160) },
                        { "local_score", candidate.LocalScore }
                    });
                }

                requestRows.Add(new Dictionary<string, object>
                {
                    { "row_id", row.RowId ?? "" },
                    { "quantity_name", row.Item == null ? "" : row.Item.Name ?? "" },
                    { "quantity_unit", row.Item == null ? "" : row.Item.Unit ?? "" },
                    { "quantity_value", row.Item == null ? "" : row.Item.ValueText ?? "" },
                    { "candidates", candidates }
                });
            }

            Dictionary<string, object> body = new Dictionary<string, object>();
            body["task"] = "请对每行工程量从 candidates 中选择最合适的一条预算定额，返回 json。";
            body["rules"] = new string[]
            {
                "selected_code 必须完全等于某个候选 code，否则置空。",
                "普通关键词检索只选一条定额。",
                "单位不接近或名称证据不足时置空。",
                "钢筋工程量不要误选钢筋混凝土构件。"
            };
            body["rows"] = requestRows;
            return serializer.Serialize(body);
        }

        private string BuildUnifiedRequestJson(List<DeepSeekRequestRow> rows)
        {
            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["model"] = settings.Model;
            payload["stream"] = false;
            payload["temperature"] = 0.1;
            payload["max_tokens"] = 1400;
            payload["response_format"] = new Dictionary<string, object> { { "type", "json_object" } };

            List<object> messages = new List<object>();
            messages.Add(new Dictionary<string, object>
            {
                { "role", "system" },
                { "content", "You are a railway quota recommendation assistant. For each row, first decide whether an existing human-corrected mapping box matches. If yes, return selected_box_id. If no mapping box is reliable, choose one quota from quota_candidates. Never invent ids. Return strict JSON: {\"results\":[{\"row_id\":\"r0\",\"selected_box_id\":\"box-123\",\"selected_code\":\"\",\"confidence\":90,\"reason\":\"short reason\"}]}." }
            });
            messages.Add(new Dictionary<string, object>
            {
                { "role", "user" },
                { "content", BuildUnifiedUserPrompt(rows) }
            });
            payload["messages"] = messages;
            return serializer.Serialize(payload);
        }

        private string BuildUnifiedUserPrompt(List<DeepSeekRequestRow> rows)
        {
            List<object> requestRows = new List<object>();
            foreach (DeepSeekRequestRow row in rows)
            {
                List<object> quotaCandidates = new List<object>();
                foreach (AiQuotaCandidate candidate in row.Candidates ?? new List<AiQuotaCandidate>())
                {
                    if (candidate == null || candidate.Quota == null)
                    {
                        continue;
                    }

                    quotaCandidates.Add(new Dictionary<string, object>
                    {
                        { "code", candidate.Quota.QuotaCode ?? "" },
                        { "name", candidate.Quota.QuotaName ?? "" },
                        { "unit", candidate.Quota.QuotaUnit ?? "" },
                        { "work_content", Truncate(candidate.Quota.WorkContent, 160) },
                        { "local_score", candidate.LocalScore }
                    });
                }

                List<object> mappingCandidates = new List<object>();
                foreach (AiMappingCandidate candidate in row.MappingCandidates ?? new List<AiMappingCandidate>())
                {
                    if (candidate == null)
                    {
                        continue;
                    }

                    mappingCandidates.Add(new Dictionary<string, object>
                    {
                        { "box_id", candidate.BoxId ?? "" },
                        { "sample_quantity_names", Truncate(candidate.SampleNames, 220) },
                        { "targets", String.Join(" + ", (candidate.Targets ?? new List<MappingTarget>()).Select(t => (t.Code ?? "") + " " + (t.Name ?? "")).ToArray()) },
                        { "local_score", candidate.LocalScore }
                    });
                }

                requestRows.Add(new Dictionary<string, object>
                {
                    { "row_id", row.RowId ?? "" },
                    { "quantity_name", row.Item == null ? "" : row.Item.Name ?? "" },
                    { "original_name", row.Item == null ? "" : row.Item.OriginalName ?? "" },
                    { "quantity_unit", row.Item == null ? "" : row.Item.Unit ?? "" },
                    { "quantity_value", row.Item == null ? "" : row.Item.ValueText ?? "" },
                    { "raw_row_text", row.Item == null ? "" : Truncate(row.Item.RawRowText, 260) },
                    { "mapping_candidates", mappingCandidates },
                    { "quota_candidates", quotaCandidates }
                });
            }

            Dictionary<string, object> body = new Dictionary<string, object>();
            body["task"] = "For each quantity row, first match mapping_candidates; if no reliable mapping box exists, select the best quota from quota_candidates.";
            body["rules"] = new string[]
            {
                "selected_box_id must exactly equal a mapping candidate box_id, otherwise leave it empty.",
                "selected_code must exactly equal a quota candidate code, otherwise leave it empty.",
                "Prefer a reliable human-corrected mapping box over a single quota.",
                "Do not choose a mapping box from generic one-word similarity only.",
                "Do not confuse steel quantities with concrete structure quotas."
            };
            body["rows"] = requestRows;
            return serializer.Serialize(body);
        }

        private List<DeepSeekSelection> ParseResponse(string responseJson)
        {
            object rootObject = serializer.DeserializeObject(responseJson);
            Dictionary<string, object> root = rootObject as Dictionary<string, object>;
            if (root == null)
            {
                return new List<DeepSeekSelection>();
            }

            List<object> choices = GetList(root, "choices");
            if (choices == null || choices.Count == 0)
            {
                return new List<DeepSeekSelection>();
            }

            Dictionary<string, object> firstChoice = choices[0] as Dictionary<string, object>;
            Dictionary<string, object> message = firstChoice == null ? null : firstChoice.ContainsKey("message") ? firstChoice["message"] as Dictionary<string, object> : null;
            string content = message == null || !message.ContainsKey("content") ? "" : Convert.ToString(message["content"], CultureInfo.InvariantCulture);
            if (String.IsNullOrWhiteSpace(content))
            {
                return new List<DeepSeekSelection>();
            }

            Dictionary<string, object> resultRoot = serializer.DeserializeObject(content) as Dictionary<string, object>;
            if (resultRoot == null)
            {
                return new List<DeepSeekSelection>();
            }

            List<object> results = GetList(resultRoot, "results");
            if (results == null && resultRoot.ContainsKey("row_id"))
            {
                results = new List<object> { resultRoot };
            }

            List<DeepSeekSelection> selections = new List<DeepSeekSelection>();
            foreach (object item in results ?? new List<object>())
            {
                Dictionary<string, object> row = item as Dictionary<string, object>;
                if (row == null)
                {
                    continue;
                }

                selections.Add(new DeepSeekSelection
                {
                    RowId = ReadString(row, "row_id"),
                    BoxId = String.IsNullOrWhiteSpace(ReadString(row, "selected_box_id")) ? ReadString(row, "box_id") : ReadString(row, "selected_box_id"),
                    SelectedCode = ReadString(row, "selected_code"),
                    Confidence = ReadInt(row, "confidence"),
                    Reason = ReadString(row, "reason")
                });
            }

            return selections;
        }

        private DeepSeekMappingSelection ParseMappingResponse(string responseJson)
        {
            object rootObject = serializer.DeserializeObject(responseJson);
            Dictionary<string, object> root = rootObject as Dictionary<string, object>;
            if (root == null)
            {
                return null;
            }

            List<object> choices = GetList(root, "choices");
            if (choices == null || choices.Count == 0)
            {
                return null;
            }

            Dictionary<string, object> firstChoice = choices[0] as Dictionary<string, object>;
            Dictionary<string, object> message = firstChoice == null || !firstChoice.ContainsKey("message") ? null : firstChoice["message"] as Dictionary<string, object>;
            string content = message == null || !message.ContainsKey("content") ? "" : Convert.ToString(message["content"], CultureInfo.InvariantCulture);
            if (String.IsNullOrWhiteSpace(content))
            {
                return null;
            }

            Dictionary<string, object> resultRoot = serializer.DeserializeObject(content) as Dictionary<string, object>;
            if (resultRoot == null)
            {
                return null;
            }

            return new DeepSeekMappingSelection
            {
                BoxId = ReadString(resultRoot, "box_id"),
                Confidence = ReadInt(resultRoot, "confidence"),
                Reason = ReadString(resultRoot, "reason")
            };
        }

        private List<DeepSeekNameResult> ParseNameResponse(string responseJson)
        {
            object rootObject = serializer.DeserializeObject(responseJson);
            Dictionary<string, object> root = rootObject as Dictionary<string, object>;
            if (root == null)
            {
                return new List<DeepSeekNameResult>();
            }

            List<object> choices = GetList(root, "choices");
            if (choices == null || choices.Count == 0)
            {
                return new List<DeepSeekNameResult>();
            }

            Dictionary<string, object> firstChoice = choices[0] as Dictionary<string, object>;
            Dictionary<string, object> message = firstChoice == null || !firstChoice.ContainsKey("message") ? null : firstChoice["message"] as Dictionary<string, object>;
            string content = message == null || !message.ContainsKey("content") ? "" : Convert.ToString(message["content"], CultureInfo.InvariantCulture);
            if (String.IsNullOrWhiteSpace(content))
            {
                return new List<DeepSeekNameResult>();
            }

            Dictionary<string, object> resultRoot = serializer.DeserializeObject(content) as Dictionary<string, object>;
            if (resultRoot == null)
            {
                return new List<DeepSeekNameResult>();
            }

            List<object> results = GetList(resultRoot, "results");
            if (results == null && resultRoot.ContainsKey("row_id"))
            {
                results = new List<object> { resultRoot };
            }

            List<DeepSeekNameResult> normalized = new List<DeepSeekNameResult>();
            foreach (object item in results ?? new List<object>())
            {
                Dictionary<string, object> row = item as Dictionary<string, object>;
                if (row == null)
                {
                    continue;
                }

                normalized.Add(new DeepSeekNameResult
                {
                    RowId = ReadString(row, "row_id"),
                    QuantityName = ReadString(row, "quantity_name"),
                    Confidence = ReadInt(row, "confidence"),
                    Reason = ReadString(row, "reason")
                });
            }

            return normalized;
        }

        private string BuildColumnLayoutRequestJson(List<List<CellValue>> rows)
        {
            int columnCount = 0;
            foreach (List<CellValue> row in rows)
            {
                if (row == null)
                {
                    continue;
                }

                foreach (CellValue cell in row)
                {
                    if (cell != null && cell.SourceIndex > columnCount)
                    {
                        columnCount = cell.SourceIndex;
                    }
                }
            }

            List<object> sampleRows = new List<object>();
            int taken = 0;
            foreach (List<CellValue> row in rows)
            {
                if (row == null || row.Count == 0)
                {
                    continue;
                }

                Dictionary<string, object> cells = new Dictionary<string, object>();
                foreach (CellValue cell in row)
                {
                    if (cell == null)
                    {
                        continue;
                    }

                    string text = Truncate(cell.Text, 40);
                    if (String.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    cells[cell.SourceIndex.ToString(CultureInfo.InvariantCulture)] = text;
                }

                if (cells.Count == 0)
                {
                    continue;
                }

                sampleRows.Add(cells);
                taken++;
                if (taken >= 12)
                {
                    break;
                }
            }

            Dictionary<string, object> body = new Dictionary<string, object>();
            body["column_count"] = columnCount;
            body["task"] = "判断该工程量表的列结构：哪些列共同组成工程量名称(name_columns，按从左到右顺序)，哪一列是单位(unit_column)，哪一列是数量(quantity_column)。";
            body["rules"] = new string[]
            {
                "列索引从1开始；name_columns 应包含数量列左侧所有描述/名称类列（如部位、项目名称、规格型号等），不含序号列。",
                "unit_column 为单位列（m、m2、m3、kg、t、处、个、座、项等）；若无单位列填0。",
                "quantity_column 为数量/工程量数值列。",
                "若无法判断填 confidence 0。"
            };
            body["rows"] = sampleRows;

            List<object> messages = new List<object>();
            messages.Add(new Dictionary<string, object>
            {
                { "role", "system" },
                { "content", "你是工程量表结构识别助手。根据给定的数据行（按 列索引->文本 给出，可能含表头行），判断每列的角色。必须严格输出 JSON：{\"name_columns\":[2,3,4],\"unit_column\":5,\"quantity_column\":6,\"confidence\":90}。列索引从1开始。" }
            });
            messages.Add(new Dictionary<string, object>
            {
                { "role", "user" },
                { "content", serializer.Serialize(body) }
            });

            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["model"] = settings.Model;
            payload["stream"] = false;
            payload["temperature"] = 0.1;
            payload["max_tokens"] = 400;
            payload["response_format"] = new Dictionary<string, object> { { "type", "json_object" } };
            payload["messages"] = messages;
            return serializer.Serialize(payload);
        }

        private DeepSeekColumnLayout ParseColumnLayoutResponse(string responseJson)
        {
            Dictionary<string, object> root = serializer.DeserializeObject(responseJson) as Dictionary<string, object>;
            if (root == null)
            {
                return null;
            }

            List<object> choices = GetList(root, "choices");
            if (choices == null || choices.Count == 0)
            {
                return null;
            }

            Dictionary<string, object> firstChoice = choices[0] as Dictionary<string, object>;
            Dictionary<string, object> message = firstChoice == null || !firstChoice.ContainsKey("message") ? null : firstChoice["message"] as Dictionary<string, object>;
            string content = message == null || !message.ContainsKey("content") ? "" : Convert.ToString(message["content"], CultureInfo.InvariantCulture);
            if (String.IsNullOrWhiteSpace(content))
            {
                return null;
            }

            Dictionary<string, object> resultRoot = serializer.DeserializeObject(content) as Dictionary<string, object>;
            if (resultRoot == null)
            {
                return null;
            }

            List<int> names = new List<int>();
            foreach (object o in GetList(resultRoot, "name_columns") ?? new List<object>())
            {
                int v;
                if (Int32.TryParse(Convert.ToString(o, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out v) && v > 0)
                {
                    names.Add(v);
                }
            }

            return new DeepSeekColumnLayout
            {
                NameColumns = names.ToArray(),
                UnitColumn = ReadInt(resultRoot, "unit_column"),
                QuantityColumn = ReadInt(resultRoot, "quantity_column"),
                Confidence = ReadInt(resultRoot, "confidence")
            };
        }

        private static List<object> GetList(Dictionary<string, object> values, string key)
        {
            object value;
            if (!values.TryGetValue(key, out value) || value == null)
            {
                return null;
            }

            ArrayList arrayList = value as ArrayList;
            if (arrayList != null)
            {
                return arrayList.Cast<object>().ToList();
            }

            object[] objectArray = value as object[];
            if (objectArray != null)
            {
                return objectArray.ToList();
            }

            return null;
        }

        private static string ReadString(Dictionary<string, object> values, string key)
        {
            object value;
            return values.TryGetValue(key, out value) && value != null ? Convert.ToString(value, CultureInfo.InvariantCulture) : "";
        }

        private static int ReadInt(Dictionary<string, object> values, string key)
        {
            object value;
            if (!values.TryGetValue(key, out value) || value == null)
            {
                return 0;
            }

            int parsed;
            return Int32.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) ? parsed : 0;
        }

        private static string Truncate(string text, int maxLength)
        {
            string value = (text ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }
    }

    internal sealed class QuotaEntry
    {
        public string TargetKind;
        public string QuotaCode;
        public string QuotaName;
        public string QuotaUnit;

        // 去掉 *乘数 / ×乘数 与 参/换/借 调整后缀，取原始编号（用于判类与索引查找）
        public static string NormalizeCode(string code)
        {
            string value = (code ?? "").Trim();
            if (value.Length == 0)
            {
                return "";
            }

            int cut = value.IndexOfAny(new[] { '*', '×' });
            if (cut >= 0)
            {
                value = value.Substring(0, cut);
            }
            value = value.Replace("参", "").Replace("换", "").Replace("借", "");
            return value.Trim();
        }

        // 三类：纯数字=材料；含横杠=定额（所有真实定额编号都带横杠，已核对全库无例外）；
        // 其余字母代号（GF/ZLF/LF/JF/SF/YF/TLF/XGT1…）=辅助代号，按材料一样与定额配套使用。
        public static string GuessKind(string code)
        {
            string value = NormalizeCode(code);
            if (value.Length == 0)
            {
                return "quota";
            }
            if (value.All(Char.IsDigit))
            {
                return "material";
            }
            return value.IndexOf('-') >= 0 ? "quota" : "aux";
        }

        public static bool IsQuotaKind(string kind)
        {
            return String.Equals(kind, "quota", StringComparison.OrdinalIgnoreCase);
        }

        public string TargetKey
        {
            get { return (String.IsNullOrWhiteSpace(TargetKind) ? GuessKind(QuotaCode) : TargetKind) + ":" + (QuotaCode ?? "").Trim().ToUpperInvariant(); }
        }
    }

    internal struct UnitScale
    {
        public string BaseUnit;
        public decimal Scale;
    }

    internal sealed class LearningRecord
    {
        public bool IsCorrection;
        public string QuantitySignature;
        public string ProjectName;
        public string BudgetFile;
        public string BudgetGroup;
        public string QuotaCode;
        public string QuotaName;
        public string QuotaUnit;
        public string QuantitySection;
        public string QuantityName;
        public string QuantityUnit;
        public string MatchReason;
        public int MatchScore;
    }

    internal sealed class SearchIndexStore
    {
        private readonly List<IndexQuota> quotas = new List<IndexQuota>();
        private readonly List<IndexMaterial> materials = new List<IndexMaterial>();
        private readonly Dictionary<string, List<IndexQuota>> quotaTokenIndex = new Dictionary<string, List<IndexQuota>>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, IndexQuota> quotasByCode = new Dictionary<string, IndexQuota>(StringComparer.OrdinalIgnoreCase);

        public int QuotaCount { get { return quotas.Count; } }
        public int MaterialCount { get { return materials.Count; } }

        public static SearchIndexStore LoadOrBuild()
        {
            SearchIndexStore store = new SearchIndexStore();
            string dataDir = LearningStore.FindDataDir();
            string quotaPath = Path.Combine(dataDir, "quota-index.jsonl");
            string materialPath = Path.Combine(dataDir, "material-index.jsonl");

            if (!File.Exists(quotaPath) || !File.Exists(materialPath))
            {
                try
                {
                    ExportFromSql(dataDir, quotaPath, materialPath);
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Build search index failed: " + ex.Message);
                }
            }

            store.LoadFiles(quotaPath, materialPath);
            return store;
        }

        public List<RecommendationRow> Search(ExcelQuantityItem item, string categoryFilter)
        {
            List<IndexQuota> candidates = GetQuotaCandidates(item, categoryFilter, null);
            if (candidates.Count == 0)
            {
                return new List<RecommendationRow>();
            }

            List<ScoredQuota> quotaHits = candidates
                .Select(q => new ScoredQuota { Quota = q, Score = ScoreQuota(item, q) })
                .Where(q => q.Score >= 55)
                .OrderByDescending(q => q.Score)
                .ThenBy(q => q.Quota.SortOrder)
                .Take(1)
                .ToList();

            List<RecommendationRow> rows = new List<RecommendationRow>();
            foreach (ScoredQuota hit in quotaHits)
            {
                rows.Add(hit.Quota.ToRecommendation(item, hit.Score));
            }

            return rows
                .GroupBy(r => (r.TargetKind ?? "") + ":" + (r.QuotaCode ?? ""), StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .OrderByDescending(r => r.Score)
                .Take(4)
                .ToList();
        }

        public List<RecommendationRow> SearchQuotaCandidates(ExcelQuantityItem item, string categoryFilter, EntryScope scope, int limit)
        {
            if (item == null)
            {
                return new List<RecommendationRow>();
            }

            int max = Math.Max(1, limit);
            string chinesePhrase = ExtractChinesePhrase(item.Name);
            string majorChapter = GetMajorChapterCode(scope);
            List<ScoredQuota> scopedHits = new List<ScoredQuota>();
            if (scope != null && scope.Strict)
            {
                foreach (string code in scope.QuotaPoolCodes)
                {
                    IndexQuota quota;
                    if (!quotasByCode.TryGetValue((code ?? "").Trim(), out quota) ||
                        !CategoryAllowed(quota.BookCategory, categoryFilter))
                    {
                        continue;
                    }

                    if (!QuotaMatchesRequiredPhrase(quota, chinesePhrase))
                    {
                        continue;
                    }

                    int score = ScoreQuota(item, quota);
                    if (score > 0)
                    {
                        scopedHits.Add(CreateScoredQuota(quota, score, scope, majorChapter));
                    }
                }
            }

            IEnumerable<ScoredQuota> allHits = GetQuotaCandidates(item, categoryFilter, null)
                .Select(q => new ScoredQuota { Quota = q, Score = ScoreQuota(item, q) })
                .Where(q => q.Score > 0)
                .Where(q => QuotaMatchesRequiredPhrase(q.Quota, chinesePhrase))
                .Select(q => CreateScoredQuota(q.Quota, q.Score, scope, majorChapter))
                .Concat(scopedHits);

            return allHits
                .GroupBy(q => q.Quota.QuotaCode ?? "", StringComparer.OrdinalIgnoreCase)
                .Select(g => g.OrderBy(x => x.MajorRank).ThenBy(x => x.PoolRank).ThenByDescending(x => x.Score).ThenBy(x => x.Quota.SortOrder).First())
                .OrderBy(q => q.MajorRank)
                .ThenBy(q => q.PoolRank)
                .ThenByDescending(q => q.Score)
                .ThenBy(q => q.Quota.SortOrder)
                .Take(max)
                .Select(q => q.Quota.ToRecommendation(item, q.Score))
                .ToList();
        }

        private static ScoredQuota CreateScoredQuota(IndexQuota quota, int score, EntryScope scope, string majorChapter)
        {
            ScoredQuota result = new ScoredQuota();
            result.Quota = quota;
            result.Score = score;
            result.PoolRank = scope != null && scope.Strict && scope.Allows("quota", quota == null ? null : quota.QuotaCode) ? 0 : 1;
            result.MajorRank = SpecialtyMatchesMajorChapter(majorChapter, quota == null ? null : quota.Specialty) ? 0 : 1;
            return result;
        }

        private static bool QuotaMatchesRequiredPhrase(IndexQuota quota, string chinesePhrase)
        {
            if (String.IsNullOrWhiteSpace(chinesePhrase))
            {
                return true;
            }

            if (quota == null)
            {
                return false;
            }

            string name = ExtractChinesePhrase(quota.QuotaName);
            string work = ExtractChinesePhrase(quota.WorkContent);
            return name.Contains(chinesePhrase) || work.Contains(chinesePhrase);
        }

        private static string ExtractChinesePhrase(string text)
        {
            StringBuilder builder = new StringBuilder();
            foreach (char ch in text ?? "")
            {
                if (ch >= 0x4e00 && ch <= 0x9fff)
                {
                    builder.Append(ch);
                }
            }
            return builder.ToString();
        }

        private static string GetMajorChapterCode(EntryScope scope)
        {
            if (scope == null)
            {
                return "";
            }

            string code = !String.IsNullOrWhiteSpace(scope.ProjectEntryCode) ? scope.ProjectEntryCode : scope.MatchedEntryCode;
            if (String.IsNullOrWhiteSpace(code))
            {
                return "";
            }

            Match match = Regex.Match(code.Trim(), "\\d+");
            if (!match.Success)
            {
                return "";
            }

            string digits = match.Value;
            if (digits.Length == 1)
            {
                return "0" + digits;
            }
            return digits.Substring(0, 2);
        }

        private static bool SpecialtyMatchesMajorChapter(string majorChapter, string specialty)
        {
            string text = TextMatcher.Normalize(specialty);
            if (String.IsNullOrWhiteSpace(majorChapter) || String.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            switch (majorChapter)
            {
                case "02":
                    return text.Contains("\u8def\u57fa");
                case "03":
                    return text.Contains("\u6865\u6db5");
                case "04":
                    return text.Contains("\u96a7\u9053");
                case "05":
                    return text.Contains("\u8f68\u9053");
                case "06":
                    return text.Contains("\u901a\u4fe1") ||
                        text.Contains("\u4fe1\u53f7") ||
                        text.Contains("\u4fe1\u606f") ||
                        text.Contains("\u707e\u5bb3\u76d1\u6d4b");
                case "07":
                    return text.Contains("\u7535\u529b") ||
                        text.Contains("\u7535\u529b\u7275\u5f15\u4f9b\u7535");
                case "08":
                    return text.Contains("\u623f\u5c4b");
                case "09":
                    return text.Contains("\u7ad9\u573a") ||
                        text.Contains("\u7ed9\u6392\u6c34") ||
                        text.Contains("\u673a\u52a1") ||
                        text.Contains("\u8f66\u8f86") ||
                        text.Contains("\u673a\u68b0") ||
                        text.Contains("\u8fd0\u8425\u751f\u4ea7\u8bbe\u5907") ||
                        text.Contains("\u5efa\u7b51\u7269");
                default:
                    return false;
            }
        }
        public bool IsMappingTargetAllowed(string targetKind, string code, string categoryFilter)
        {
            string kind = String.IsNullOrWhiteSpace(targetKind) ? QuotaEntry.GuessKind(code) : targetKind;
            if (!QuotaEntry.IsQuotaKind(kind))
            {
                return true; // 材料与辅助代号随定额一起带出，不受定额类别过滤
            }

            // 扶正时可能带 *乘数（如 LY-25*9），查索引前先归一化取原始编号
            IndexQuota quota;
            if (!quotasByCode.TryGetValue(QuotaEntry.NormalizeCode(code), out quota))
            {
                return false;
            }

            return CategoryAllowed(quota.BookCategory, categoryFilter);
        }

        public List<AiQuotaCandidate> BuildDeepSeekCandidates(ExcelQuantityItem item, string categoryFilter, int limit, EntryScope scope)
        {
            if (item == null)
            {
                return new List<AiQuotaCandidate>();
            }

            int max = Math.Max(1, limit);
            return GetQuotaCandidates(item, categoryFilter, scope)
                .Select(q => new AiQuotaCandidate { Quota = q, LocalScore = ScoreQuota(item, q) })
                .Where(c => c.LocalScore > 0)
                .OrderByDescending(c => c.LocalScore)
                .ThenBy(c => c.Quota.SortOrder)
                .Take(max)
                .ToList();
        }

        // 严格条目模式：把整条目定额池作为候选（不依赖关键词命中），按名称相似度打分。
        // 这样池里相关定额即使名称不完全匹配也能进候选，既提高本地直接命中率，也让 AI 拿到聚焦的小候选集，更快更准。
        public List<AiQuotaCandidate> BuildScopeCandidates(ExcelQuantityItem item, EntryScope scope, int limit)
        {
            if (item == null || scope == null || !scope.Strict)
            {
                return new List<AiQuotaCandidate>();
            }

            List<AiQuotaCandidate> candidates = new List<AiQuotaCandidate>();
            foreach (string code in scope.QuotaPoolCodes)
            {
                IndexQuota quota;
                if (quotasByCode.TryGetValue((code ?? "").Trim(), out quota))
                {
                    candidates.Add(new AiQuotaCandidate { Quota = quota, LocalScore = ScoreQuota(item, quota) });
                }
            }

            return candidates
                .OrderByDescending(c => c.LocalScore)
                .ThenBy(c => c.Quota.SortOrder)
                .Take(Math.Max(1, limit))
                .ToList();
        }

        public AiQuotaCandidate BuildDeepSeekCandidateByCode(ExcelQuantityItem item, string code, string categoryFilter)
        {
            if (item == null || String.IsNullOrWhiteSpace(code))
            {
                return null;
            }

            IndexQuota quota;
            if (!quotasByCode.TryGetValue(code.Trim(), out quota) || !CategoryAllowed(quota.BookCategory, categoryFilter))
            {
                return null;
            }

            return new AiQuotaCandidate { Quota = quota, LocalScore = ScoreQuota(item, quota) };
        }

        public List<AiQuotaCandidate> BuildDeepSeekCandidatesFromRecommendations(ExcelQuantityItem item, List<RecommendationRow> rows, string categoryFilter)
        {
            List<AiQuotaCandidate> candidates = new List<AiQuotaCandidate>();
            foreach (RecommendationRow row in rows ?? new List<RecommendationRow>())
            {
                if (row == null || !String.Equals(String.IsNullOrWhiteSpace(row.TargetKind) ? QuotaEntry.GuessKind(row.QuotaCode) : row.TargetKind, "quota", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                AiQuotaCandidate candidate = BuildDeepSeekCandidateByCode(item, row.QuotaCode, categoryFilter);
                if (candidate != null)
                {
                    candidate.LocalScore = Math.Max(candidate.LocalScore, row.Score);
                    candidates.Add(candidate);
                }
            }

            return candidates;
        }

        private void LoadFiles(string quotaPath, string materialPath)
        {
            if (File.Exists(quotaPath))
            {
                foreach (string line in File.ReadAllLines(quotaPath, Encoding.UTF8))
                {
                    if (String.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                    IndexQuota quota = new IndexQuota();
                    quota.QuotaCode = LearningStore.Get(values, "quota_code");
                    quota.QuotaName = LearningStore.Get(values, "quota_name");
                    quota.QuotaUnit = LearningStore.Get(values, "quota_unit");
                    quota.BookCode = LearningStore.Get(values, "book_code");
                    quota.BookCategory = LearningStore.Get(values, "book_category");
                    quota.Specialty = LearningStore.Get(values, "specialty");
                    quota.SectionNo = LearningStore.Get(values, "section_no");
                    quota.SectionName = LearningStore.Get(values, "section_name");
                    quota.WorkContent = LearningStore.Get(values, "work_content");
                    double basePrice;
                    quota.BasePrice = Double.TryParse(
                        LearningStore.Get(values, "base_price"),
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        out basePrice) ? basePrice : 0d;
                    quota.SearchText = LearningStore.Get(values, "search_text");
                    int sortOrder;
                    quota.SortOrder = Int32.TryParse(LearningStore.Get(values, "sort_order"), NumberStyles.Integer, CultureInfo.InvariantCulture, out sortOrder) ? sortOrder : Int32.MaxValue;
                    if (!String.IsNullOrWhiteSpace(quota.QuotaCode))
                    {
                        quotas.Add(quota);
                        if (!quotasByCode.ContainsKey(quota.QuotaCode.Trim()))
                        {
                            quotasByCode[quota.QuotaCode.Trim()] = quota;
                        }
                    }
                }

                BuildQuotaTokenIndex();
            }

            if (File.Exists(materialPath))
            {
                foreach (string line in File.ReadAllLines(materialPath, Encoding.UTF8))
                {
                    if (String.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                    IndexMaterial material = new IndexMaterial();
                    material.MaterialCode = LearningStore.Get(values, "material_code");
                    material.MaterialName = LearningStore.Get(values, "material_name");
                    material.MaterialUnit = LearningStore.Get(values, "material_unit");
                    material.DocNo = LearningStore.Get(values, "doc_no");
                    material.IsMainMaterial = LearningStore.Get(values, "is_main_material") == "1";
                    material.TransportCategory = LearningStore.Get(values, "transport_category");
                    material.SearchText = LearningStore.Get(values, "search_text");
                    if (!String.IsNullOrWhiteSpace(material.MaterialCode))
                    {
                        materials.Add(material);
                    }
                }
            }
        }

        private void BuildQuotaTokenIndex()
        {
            quotaTokenIndex.Clear();
            foreach (IndexQuota quota in quotas)
            {
                foreach (string token in QuotaIndexTokens(quota))
                {
                    List<IndexQuota> bucket;
                    if (!quotaTokenIndex.TryGetValue(token, out bucket))
                    {
                        bucket = new List<IndexQuota>();
                        quotaTokenIndex[token] = bucket;
                    }
                    bucket.Add(quota);
                }
            }
        }

        private static IEnumerable<string> QuotaIndexTokens(IndexQuota quota)
        {
            string text = String.Join(" ", new string[] { quota.QuotaName, quota.WorkContent, quota.SectionName, quota.Specialty });
            foreach (string token in TextMatcher.Keywords(text).Distinct())
            {
                if (UseTokenForCandidateLookup(token))
                {
                    yield return token;
                }
            }
        }

        private List<IndexQuota> GetQuotaCandidates(ExcelQuantityItem item, string categoryFilter, EntryScope scope)
        {
            HashSet<IndexQuota> candidates = new HashSet<IndexQuota>();
            foreach (string token in CandidateLookupTokens(item.Name))
            {
                List<IndexQuota> bucket;
                if (!quotaTokenIndex.TryGetValue(token, out bucket))
                {
                    continue;
                }

                foreach (IndexQuota quota in bucket)
                {
                    if (!CategoryAllowed(quota.BookCategory, categoryFilter))
                    {
                        continue;
                    }

                    // 严格条目模式：候选只取当前条目定额池内的定额
                    if (scope != null && scope.Strict && !scope.Allows("quota", quota.QuotaCode))
                    {
                        continue;
                    }

                    candidates.Add(quota);
                }
            }

            return candidates.ToList();
        }

        private static IEnumerable<string> CandidateLookupTokens(string quantityName)
        {
            string normalized = TextMatcher.Normalize(quantityName).Replace(" ", "");
            if (UseTokenForCandidateLookup(normalized))
            {
                yield return normalized;
            }

            foreach (string token in TextMatcher.Keywords(quantityName).Distinct())
            {
                if (UseTokenForCandidateLookup(token))
                {
                    yield return token;
                }
            }
        }

        private static bool UseTokenForCandidateLookup(string token)
        {
            if (String.IsNullOrWhiteSpace(token))
            {
                return false;
            }

            if (token.Length == 1)
            {
                return TextMatcher.IsPureChinese(token);
            }

            return !TextMatcher.IsNumberLikeToken(token);
        }

        private static bool CategoryAllowed(string bookCategory, string categoryFilter)
        {
            string category = (bookCategory ?? "").Trim();
            string filter = (categoryFilter ?? "").Trim();
            if (String.IsNullOrWhiteSpace(filter) || String.Equals(filter, "\u5168\u90e8", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (String.Equals(category, filter, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return IsCommonCategory(category);
        }

        private static bool IsCommonCategory(string category)
        {
            return String.IsNullOrWhiteSpace(category) ||
                String.Equals(category, "\u57fa\u672c\u5b9a\u989d", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(category, "\u8865\u5145\u5b9a\u989d", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(category, "\u8865\u5145\u5355\u4ef7\u5206\u6790", StringComparison.OrdinalIgnoreCase);
        }

        private static int ScoreQuota(ExcelQuantityItem item, IndexQuota quota)
        {
            string query = TextMatcher.Normalize(item.Name);
            string nameText = TextMatcher.Normalize(quota.QuotaName);
            string workText = TextMatcher.Normalize(quota.WorkContent);
            string sectionText = TextMatcher.Normalize(quota.SectionName);
            string specialtyText = TextMatcher.Normalize(quota.Specialty);
            string searchable = TextMatcher.Normalize(quota.SearchText);
            if (String.IsNullOrWhiteSpace(query) || String.IsNullOrWhiteSpace(searchable))
            {
                return 0;
            }

            if (IsSteelQuantity(query) && IsConcreteQuotaName(nameText))
            {
                return 0;
            }

            int score = 0;
            int primaryMatches = 0;
            int coreMatches = 0;
            if (nameText.Contains(query))
            {
                score += 90;
                primaryMatches++;
            }
            else if (workText.Contains(query))
            {
                score += 65;
                primaryMatches++;
            }

            List<string> tokens = TextMatcher.Keywords(item.Name).Distinct().ToList();
            int matched = 0;
            foreach (string token in tokens)
            {
                if (token.Length < 1 || !searchable.Contains(token))
                {
                    continue;
                }

                matched++;
                bool coreToken = TextMatcher.IsPureChinese(token) && token.Length >= 2;
                if (coreToken)
                {
                    coreMatches++;
                }

                bool primaryHit = nameText.Contains(token) || workText.Contains(token);
                if (primaryHit && token.Length >= 2)
                {
                    primaryMatches++;
                }

                if (nameText.Contains(token))
                {
                    score += TokenScore(token, 55, 42, 28);
                }
                else if (workText.Contains(token))
                {
                    score += TokenScore(token, 36, 28, 18);
                }
                else if (sectionText.Contains(token))
                {
                    score += TokenScore(token, 18, 14, 9);
                }
                else if (specialtyText.Contains(token))
                {
                    score += TokenScore(token, 10, 8, 5);
                }
            }

            if (IsSteelQuantity(query))
            {
                score += SteelPreferenceScore(query, nameText, workText, quota.QuotaUnit);
            }

            if (primaryMatches == 0 && coreMatches < 2)
            {
                return 0;
            }

            if (tokens.Count > 0)
            {
                score += (int)Math.Round(20.0 * matched / tokens.Count);
            }

            if (RecommendDialog.UnitCompatibleForIndex(quota.QuotaUnit, item.Unit))
            {
                score += 12;
            }

            return score;
        }

        private static int TokenScore(string token, int shortChinese, int longChinese, int mixed)
        {
            if (TextMatcher.HasAsciiOrDigit(token))
            {
                return mixed;
            }

            if (token.Length == 1)
            {
                return Math.Max(1, shortChinese / 10);
            }

            return token.Length == 2 ? shortChinese : longChinese;
        }

        private static bool IsSteelQuantity(string normalizedQuery)
        {
            return TextMatcher.IsSteelQuantityName(normalizedQuery);
        }

        private static bool IsConcreteQuotaName(string normalizedQuotaName)
        {
            if (!TextMatcher.IsConcreteQuantityName(normalizedQuotaName))
            {
                return false;
            }

            return !normalizedQuotaName.Contains("构件钢筋") &&
                !normalizedQuotaName.Contains("圆钢筋") &&
                !normalizedQuotaName.Contains("螺纹钢筋") &&
                !normalizedQuotaName.Contains("箍筋") &&
                !normalizedQuotaName.Contains("钢筋制作") &&
                !normalizedQuotaName.Contains("钢筋制安") &&
                !normalizedQuotaName.Contains("钢筋绑扎");
        }

        private static int SteelPreferenceScore(string query, string nameText, string workText, string quotaUnit)
        {
            int score = 0;
            bool steelOperation = nameText.Contains("构件钢筋") ||
                nameText.Contains("钢筋制作") ||
                nameText.Contains("钢筋制安") ||
                nameText.Contains("钢筋绑扎") ||
                workText.Contains("钢筋制作") ||
                workText.Contains("钢筋制安") ||
                (workText.Contains("钢筋") && (workText.Contains("制作") || workText.Contains("绑扎") || workText.Contains("安装")));

            if (steelOperation)
            {
                score += 55;
            }

            if ((query.Contains("hpb") || query.Contains("光圆") || query.Contains("圆钢")) && nameText.Contains("圆钢筋"))
            {
                score += 80;
            }

            if ((query.Contains("hrb") || query.Contains("螺纹")) && (nameText.Contains("hrb") || nameText.Contains("螺纹钢筋")))
            {
                score += 80;
            }

            if (nameText.Contains("钢筋混凝土") && !steelOperation)
            {
                score -= 70;
            }

            if (!RecommendDialog.UnitCompatibleForIndex(quotaUnit, "kg") && !RecommendDialog.UnitCompatibleForIndex(quotaUnit, "t"))
            {
                score -= 30;
            }

            return score;
        }

        private static void ExportFromSql(string dataDir, string quotaPath, string materialPath)
        {
            Directory.CreateDirectory(dataDir);
            string server = ReadServer();
            if (String.IsNullOrWhiteSpace(server))
            {
                server = "127.0.0.1";
            }

            string databaseName = ResolveDatabaseName();
            QuotaRecommendPanel.Log("Build search index from database: " + databaseName);
            string connectionString = "Data Source=" + server + ",1433;Initial Catalog=" + databaseName + ";User ID=reco;Password=" + BuildSqlPassword() + ";Connect Timeout=8;Encrypt=False;TrustServerCertificate=True";
            using (System.Data.SqlClient.SqlConnection connection = new System.Data.SqlClient.SqlConnection(connectionString))
            {
                connection.Open();
                WriteQuotaIndex(connection, quotaPath);
                WriteMaterialIndex(connection, materialPath);
            }
        }

        private static string ReadServer()
        {
            try
            {
                string baseDir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location);
                string path = Path.Combine(baseDir, "ServerSetting.xml");
                if (!File.Exists(path))
                {
                    return "";
                }

                string text = File.ReadAllText(path, Encoding.UTF8);
                int start = text.IndexOf("<ServerIP>", StringComparison.OrdinalIgnoreCase);
                int end = text.IndexOf("</ServerIP>", StringComparison.OrdinalIgnoreCase);
                if (start >= 0 && end > start)
                {
                    return text.Substring(start + 10, end - start - 10).Trim();
                }

                start = text.IndexOf("<Server>", StringComparison.OrdinalIgnoreCase);
                end = text.IndexOf("</Server>", StringComparison.OrdinalIgnoreCase);
                return start >= 0 && end > start ? text.Substring(start + 8, end - start - 8).Trim() : "";
            }
            catch
            {
                return "";
            }
        }

        private static string BuildSqlPassword()
        {
            return String.Join("_", new string[] { "Des", "Reco", "2006" });
        }

        private static string ResolveDatabaseName()
        {
            try
            {
                string baseDir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location) ?? "";
                string processPath = "";
                try
                {
                    processPath = Process.GetCurrentProcess().MainModule.FileName ?? "";
                }
                catch
                {
                }

                string probe = (baseDir + " " + processPath).ToLowerInvariant();
                if (probe.Contains("2024") ||
                    File.Exists(Path.Combine(baseDir, "ReJJGSNet2024.exe")) ||
                    File.Exists(Path.Combine(baseDir, "ReJJQDNet2024.exe")))
                {
                    return "RecoData2024";
                }
            }
            catch
            {
            }

            return "RecoData2020";
        }

        private static void WriteQuotaIndex(System.Data.SqlClient.SqlConnection connection, string path)
        {
            string temp = path + ".tmp";
            using (StreamWriter writer = new StreamWriter(temp, false, Encoding.UTF8))
            using (System.Data.SqlClient.SqlCommand command = connection.CreateCommand())
            {
                command.CommandTimeout = 60;
                command.CommandText =
                    "SELECT a.定额编号,a.定额名称,a.单位,a.书号,ISNULL(b.分类,''),ISNULL(b.专业名称,''),ISNULL(a.节号,''),ISNULL(c.节名称,''),ISNULL(CAST(a.工作内容 AS nvarchar(max)),''),ISNULL(a.基价,0),ISNULL(b.现行定额,0),ISNULL(a.流水号,2147483647) " +
                    "FROM dbo.定额库 a LEFT JOIN dbo.定额库索引 b ON a.书号=b.书号 LEFT JOIN dbo.定额节索引 c ON a.书号=c.书号 AND a.节号=c.节号 " +
                    "WHERE ISNULL(a.定额编号,'')<>'' AND ISNULL(a.定额名称,'')<>''";
                using (System.Data.SqlClient.SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        Dictionary<string, string> row = new Dictionary<string, string>();
                        row["quota_code"] = ReadString(reader, 0);
                        row["quota_name"] = ReadString(reader, 1);
                        row["quota_unit"] = ReadString(reader, 2);
                        row["book_code"] = ReadString(reader, 3);
                        row["book_category"] = ReadString(reader, 4);
                        row["specialty"] = ReadString(reader, 5);
                        row["section_no"] = ReadString(reader, 6);
                        row["section_name"] = ReadString(reader, 7);
                        row["work_content"] = ReadString(reader, 8);
                        row["base_price"] = Convert.ToString(reader.GetDouble(9), CultureInfo.InvariantCulture);
                        row["is_current"] = IsTruthy(reader.GetValue(10)) ? "1" : "0";
                        row["sort_order"] = Convert.ToString(reader.GetInt32(11), CultureInfo.InvariantCulture);
                        row["search_text"] = TextMatcher.Normalize(String.Join(" ", new string[] { row["quota_code"], row["quota_name"], row["quota_unit"], row["book_category"], row["specialty"], row["section_name"], row["work_content"] }));
                        writer.WriteLine(LearningStore.ToJson(row));
                    }
                }
            }

            ReplaceFile(temp, path);
        }

        private static void WriteMaterialIndex(System.Data.SqlClient.SqlConnection connection, string path)
        {
            string temp = path + ".tmp";
            using (StreamWriter writer = new StreamWriter(temp, false, Encoding.UTF8))
            using (System.Data.SqlClient.SqlCommand command = connection.CreateCommand())
            {
                command.CommandTimeout = 60;
                command.CommandText =
                    "SELECT 电算代号,材料名称,单位,文号,ISNULL(主材标志,''),ISNULL(材料运输类别,''),ISNULL(基期单价,0),ISNULL(编制期价,0) " +
                    "FROM dbo.材料单价库 WHERE ISNULL(材料名称,'')<>''";
                using (System.Data.SqlClient.SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        Dictionary<string, string> row = new Dictionary<string, string>();
                        row["material_code"] = Convert.ToString(reader.GetValue(0), CultureInfo.InvariantCulture);
                        row["material_name"] = ReadString(reader, 1);
                        row["material_unit"] = ReadString(reader, 2);
                        row["doc_no"] = ReadString(reader, 3);
                        row["is_main_material"] = ReadString(reader, 4) == "1" ? "1" : "0";
                        row["transport_category"] = ReadString(reader, 5);
                        row["base_price"] = Convert.ToString(reader.GetDouble(6), CultureInfo.InvariantCulture);
                        row["current_price"] = Convert.ToString(reader.GetDouble(7), CultureInfo.InvariantCulture);
                        row["search_text"] = TextMatcher.Normalize(String.Join(" ", new string[] { row["material_code"], row["material_name"], row["material_unit"], row["doc_no"], row["transport_category"] }));
                        writer.WriteLine(LearningStore.ToJson(row));
                    }
                }
            }

            ReplaceFile(temp, path);
        }

        private static string ReadString(System.Data.SqlClient.SqlDataReader reader, int index)
        {
            return reader.IsDBNull(index) ? "" : Convert.ToString(reader.GetValue(index), CultureInfo.CurrentCulture).Trim();
        }

        private static bool IsTruthy(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return false;
            }

            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture) != 0;
            }
            catch
            {
                return String.Equals(Convert.ToString(value, CultureInfo.InvariantCulture), "true", StringComparison.OrdinalIgnoreCase);
            }
        }

        private static void ReplaceFile(string temp, string path)
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
            File.Move(temp, path);
        }

        private sealed class ScoredQuota
        {
            public IndexQuota Quota;
            public int Score;
            public int MajorRank;
            public int PoolRank;
        }

    }

    internal sealed class IndexQuota
    {
        public string QuotaCode;
        public string QuotaName;
        public string QuotaUnit;
        public string BookCode;
        public string BookCategory;
        public string Specialty;
        public string SectionNo;
        public string SectionName;
        public string WorkContent;
        public double BasePrice;
        public string SearchText;
        public int SortOrder;

        public RecommendationRow ToRecommendation(ExcelQuantityItem item, int score)
        {
            RecommendationRow row = new RecommendationRow();
            row.Item = item;
            row.QuotaCode = QuotaCode;
            row.QuotaName = QuotaName;
            row.QuotaUnit = QuotaUnit;
            row.BookCode = BookCode;
            row.Specialty = Specialty;
            row.BasePrice = BasePrice;
            row.WorkContent = WorkContent;
            row.ConvertedValueText = RecommendDialog.ConvertQuantityForIndex(item.ValueText, item.Unit, QuotaUnit);
            row.Score = score;
            row.Reason = "\u5168\u91cf\u5b9a\u989d\u7d22\u5f15\u5173\u952e\u8bcd\u5339\u914d";
            row.Source = "index";
            row.TargetKind = "quota";
            return row;
        }
    }

    internal sealed class IndexMaterial
    {
        public string MaterialCode;
        public string MaterialName;
        public string MaterialUnit;
        public string DocNo;
        public bool IsMainMaterial;
        public string TransportCategory;
        public string SearchText;

        public RecommendationRow ToRecommendation(ExcelQuantityItem item, int score)
        {
            RecommendationRow row = new RecommendationRow();
            row.Item = item;
            row.QuotaCode = MaterialCode;
            row.QuotaName = MaterialName;
            row.QuotaUnit = MaterialUnit;
            row.ConvertedValueText = RecommendDialog.ConvertQuantityForIndex(item.ValueText, item.Unit, MaterialUnit);
            row.Score = score;
            row.Reason = IsMainMaterial ? "\u4e3b\u8981\u6750\u6599\u7d22\u5f15\u5173\u952e\u8bcd\u5339\u914d" : "\u6750\u6599\u7d22\u5f15\u5173\u952e\u8bcd\u5339\u914d";
            row.Source = "index";
            row.TargetKind = "material";
            return row;
        }
    }

    // 当前定额行所在章节条目的推荐范围；Strict=true 时三个阶段的候选都严格限制在该条目定额池内
    internal sealed class EntryScope
    {
        public string ProjectEntryCode;   // 项目章节表里的原始条目编号
        public string MatchedEntryCode;   // 前缀上溯后命中的库内条目编号
        public string EntryName;
        public string Method;             // "2020" / "2024"
        public HashSet<string> PoolKeys;  // "kind:CODE"（大写）

        public bool Strict
        {
            get { return !String.IsNullOrEmpty(MatchedEntryCode) && PoolKeys != null && PoolKeys.Count > 0; }
        }

        // 对应框 entry_codes 标签格式：method:条目编号
        public string Tag
        {
            get { return Method + ":" + MatchedEntryCode; }
        }

        public bool Allows(string targetKind, string code)
        {
            if (!Strict)
            {
                return true;
            }

            string kind = String.IsNullOrWhiteSpace(targetKind) ? QuotaEntry.GuessKind(code) : targetKind;
            return PoolKeys.Contains(kind.ToLowerInvariant() + ":" + QuotaEntry.NormalizeCode(code).ToUpperInvariant());
        }

        // 条目定额池里所有 quota 类定额编号（去掉 "quota:" 前缀），用于把整池作为候选喂给本地匹配/AI
        public IEnumerable<string> QuotaPoolCodes
        {
            get
            {
                if (PoolKeys == null)
                {
                    yield break;
                }

                foreach (string key in PoolKeys)
                {
                    if (key.StartsWith("quota:", StringComparison.OrdinalIgnoreCase))
                    {
                        yield return key.Substring("quota:".Length);
                    }
                }
            }
        }
    }

    // 章节条目定额库：chapter-entries.jsonl（删减后的条目树）+ chapter-quota-library.jsonl（条目定额池）。
    // chapter-entries.jsonl 不存在时 IsEmpty=true，推荐行为与历史版本完全一致。
    internal sealed class ChapterLibraryStore
    {
        private readonly Dictionary<string, string> entryNames = new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly Dictionary<string, string> entryTypes = new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly Dictionary<string, HashSet<string>> pools = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        // 规范化条目名称 → 小计/指标条目编号列表（识别用户复制条目的来源）
        private readonly Dictionary<string, List<string>> nameIndex = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        public string MethodKey = "";
        public string MethodNo = "";

        public bool IsEmpty
        {
            get { return entryNames.Count == 0; }
        }

        public static ChapterLibraryStore Load()
        {
            ChapterLibraryStore store = new ChapterLibraryStore();
            try
            {
                store.MethodKey = ResolveMethodKey();
                string dataDir = LearningStore.FindDataDir();
                string entriesPath = Path.Combine(dataDir, "chapter-entries.jsonl");
                string libraryPath = Path.Combine(dataDir, "chapter-quota-library.jsonl");
                if (!File.Exists(entriesPath) || !File.Exists(libraryPath))
                {
                    return store;
                }

                foreach (string line in File.ReadAllLines(entriesPath, Encoding.UTF8))
                {
                    if (String.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                    if (!String.Equals(LearningStore.Get(values, "method"), store.MethodKey, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    string code = LearningStore.Get(values, "entry_code").Trim();
                    if (!String.IsNullOrEmpty(code) && !store.entryNames.ContainsKey(code))
                    {
                        store.entryNames[code] = LearningStore.Get(values, "entry_name").Trim();
                        store.entryTypes[code] = LearningStore.Get(values, "entry_type").Trim();
                        if (String.IsNullOrEmpty(store.MethodNo))
                        {
                            store.MethodNo = LearningStore.Get(values, "method_no").Trim();
                        }
                    }
                }

                if (store.entryNames.Count == 0)
                {
                    return store;
                }

                foreach (string line in File.ReadAllLines(libraryPath, Encoding.UTF8))
                {
                    if (String.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                    if (!String.Equals(LearningStore.Get(values, "method"), store.MethodKey, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    string entryCode = LearningStore.Get(values, "entry_code").Trim();
                    string code = LearningStore.Get(values, "quota_code").Trim();
                    if (String.IsNullOrEmpty(entryCode) || String.IsNullOrEmpty(code) || !store.entryNames.ContainsKey(entryCode))
                    {
                        continue;
                    }

                    if (LearningStore.Get(values, "deleted").Trim() == "1")
                    {
                        store.RemovePoolKey(entryCode, LearningStore.Get(values, "target_kind"), code);
                    }
                    else
                    {
                        store.AddPoolKey(entryCode, LearningStore.Get(values, "target_kind"), code);
                    }
                }

                store.BuildNameIndex();
                QuotaRecommendPanel.Log("ChapterLibraryStore loaded. method=" + store.MethodKey + " entries=" + store.entryNames.Count.ToString(CultureInfo.InvariantCulture) + " pooledEntries=" + store.pools.Count.ToString(CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("ChapterLibraryStore load failed: " + ex.Message);
            }

            return store;
        }

        // 与 SearchIndexStore.ResolveDatabaseName 同一套判断：按运行目录/进程判定 2020 还是 2024
        private static string ResolveMethodKey()
        {
            try
            {
                string baseDir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location) ?? "";
                string processPath = "";
                try
                {
                    processPath = Process.GetCurrentProcess().MainModule.FileName ?? "";
                }
                catch
                {
                }

                string probe = (baseDir + " " + processPath).ToLowerInvariant();
                if (probe.Contains("2024") ||
                    File.Exists(Path.Combine(baseDir, "ReJJGSNet2024.exe")) ||
                    File.Exists(Path.Combine(baseDir, "ReJJQDNet2024.exe")))
                {
                    return "2024";
                }
            }
            catch
            {
            }

            return "2020";
        }

        private void AddPoolKey(string entryCode, string kind, string code)
        {
            HashSet<string> pool;
            if (!pools.TryGetValue(entryCode, out pool))
            {
                pool = new HashSet<string>(StringComparer.Ordinal);
                pools[entryCode] = pool;
            }

            string normalizedKind = String.IsNullOrWhiteSpace(kind) ? QuotaEntry.GuessKind(code) : kind.Trim().ToLowerInvariant();
            pool.Add(normalizedKind + ":" + code.ToUpperInvariant());
        }

        private void RemovePoolKey(string entryCode, string kind, string code)
        {
            HashSet<string> pool;
            if (String.IsNullOrEmpty(entryCode) || String.IsNullOrEmpty(code) || !pools.TryGetValue(entryCode, out pool))
            {
                return;
            }

            string normalizedKind = String.IsNullOrWhiteSpace(kind) ? QuotaEntry.GuessKind(code) : kind.Trim().ToLowerInvariant();
            pool.Remove(normalizedKind + ":" + code.ToUpperInvariant());
        }

        private static bool IsQuotaInputEntryType(string entryType)
        {
            return entryType == "小计" || entryType == "指标";
        }

        private static string NormalizeEntryName(string name)
        {
            return Regex.Replace(name ?? "", "[\\s　]+", "").ToLowerInvariant();
        }

        // 名称索引只收"小计/指标"且有池的条目——只有这两类条目有定额输入框
        private void BuildNameIndex()
        {
            nameIndex.Clear();
            foreach (KeyValuePair<string, string> pair in entryNames)
            {
                string entryType;
                HashSet<string> pool;
                if (!entryTypes.TryGetValue(pair.Key, out entryType) || !IsQuotaInputEntryType(entryType))
                {
                    continue;
                }
                if (!pools.TryGetValue(pair.Key, out pool) || pool.Count == 0)
                {
                    continue;
                }

                string nameKey = NormalizeEntryName(pair.Value);
                if (String.IsNullOrEmpty(nameKey))
                {
                    continue;
                }

                List<string> codes;
                if (!nameIndex.TryGetValue(nameKey, out codes))
                {
                    codes = new List<string>();
                    nameIndex[nameKey] = codes;
                }
                codes.Add(pair.Key);
            }
        }

        // 项目条目编号 → 库内保留条目。
        // 顺序：编号精确命中 → 按名称识别"复制条目"的来源（同祖先链优先，再全局唯一）→ 逐级前缀上溯。
        public EntryScope ResolveScope(string projectEntryCode, string projectEntryName)
        {
            string current = (projectEntryCode ?? "").Trim();
            if (String.IsNullOrEmpty(current) || IsEmpty)
            {
                return null;
            }

            EntryScope exact = BuildScope(current, current);
            if (exact != null)
            {
                return exact;
            }

            // 编号不在库内（用户新建/复制的条目）⇒ 按名称找复制来源条目，用它的定额池
            if (!entryNames.ContainsKey(current))
            {
                string nameKey = NormalizeEntryName(projectEntryName);
                List<string> sameName;
                if (!String.IsNullOrEmpty(nameKey) && nameIndex.TryGetValue(nameKey, out sameName) && sameName.Count > 0)
                {
                    string prefix = current;
                    while (true)
                    {
                        int dash = prefix.LastIndexOf('-');
                        if (dash <= 0)
                        {
                            break;
                        }
                        prefix = prefix.Substring(0, dash);
                        string withDash = prefix + "-";
                        foreach (string candidate in sameName)
                        {
                            if (candidate.StartsWith(withDash, StringComparison.Ordinal) || candidate == prefix)
                            {
                                EntryScope copied = BuildScope(current, candidate);
                                if (copied != null)
                                {
                                    return copied;
                                }
                            }
                        }
                    }

                    if (sameName.Count == 1)
                    {
                        EntryScope unique = BuildScope(current, sameName[0]);
                        if (unique != null)
                        {
                            return unique;
                        }
                    }
                }
            }

            // 逐级前缀上溯到最近的"保留且池非空"条目
            string probe = current;
            while (!String.IsNullOrEmpty(probe) && probe != "0")
            {
                EntryScope scope = BuildScope(current, probe);
                if (scope != null)
                {
                    return scope;
                }

                int dash2 = probe.LastIndexOf('-');
                if (dash2 > 0)
                {
                    probe = probe.Substring(0, dash2);
                    continue;
                }

                if (probe.Length > 2)
                {
                    probe = probe.Substring(0, 2);
                    continue;
                }

                break;
            }

            return null;
        }

        private EntryScope BuildScope(string projectEntryCode, string matchedCode)
        {
            HashSet<string> pool;
            if (!entryNames.ContainsKey(matchedCode) || !pools.TryGetValue(matchedCode, out pool) || pool.Count == 0)
            {
                return null;
            }

            EntryScope scope = new EntryScope();
            scope.ProjectEntryCode = projectEntryCode;
            scope.MatchedEntryCode = matchedCode;
            scope.EntryName = entryNames[matchedCode];
            scope.Method = MethodKey;
            scope.PoolKeys = pool;
            return scope;
        }

        // 用户扶正/采纳了池外定额时补进池子，严格模式不与用户作对；追加 source=user 行持久化
        public void AddUserQuota(EntryScope scope, string targetKind, string code, string name, string unit)
        {
            if (scope == null || !scope.Strict || String.IsNullOrWhiteSpace(code))
            {
                return;
            }

            string kind = String.IsNullOrWhiteSpace(targetKind) ? QuotaEntry.GuessKind(code) : targetKind.Trim().ToLowerInvariant();
            string key = kind + ":" + code.Trim().ToUpperInvariant();
            if (scope.PoolKeys.Contains(key))
            {
                return;
            }

            AddPoolKey(scope.MatchedEntryCode, kind, code.Trim());
            try
            {
                Dictionary<string, string> record = new Dictionary<string, string>();
                record["record_type"] = "entry_quota";
                record["method"] = MethodKey;
                record["method_no"] = MethodNo;
                record["entry_code"] = scope.MatchedEntryCode;
                record["entry_name"] = scope.EntryName ?? "";
                record["target_kind"] = kind;
                record["quota_code"] = code.Trim();
                record["quota_name"] = name ?? "";
                record["quota_unit"] = unit ?? "";
                record["project_count"] = "0";
                record["source"] = "user";
                record["last_seen"] = DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                string path = Path.Combine(LearningStore.FindDataDir(), "chapter-quota-library.jsonl");
                File.AppendAllText(path, LearningStore.ToJson(record) + Environment.NewLine, Encoding.UTF8);
                QuotaRecommendPanel.Log("ChapterLibrary user quota added. entry=" + scope.MatchedEntryCode + " code=" + code);
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("ChapterLibrary user quota append failed: " + ex.Message);
            }
        }

        // 用户从参考池删除定额：从内存池移除并追加 deleted=1 墓碑行（软删除，可被后续 add 覆盖恢复）
        public void RemoveUserQuota(EntryScope scope, string targetKind, string code)
        {
            if (scope == null || String.IsNullOrEmpty(scope.MatchedEntryCode) || String.IsNullOrWhiteSpace(code))
            {
                return;
            }

            string kind = String.IsNullOrWhiteSpace(targetKind) ? QuotaEntry.GuessKind(code) : targetKind.Trim().ToLowerInvariant();
            RemovePoolKey(scope.MatchedEntryCode, kind, code.Trim());

            try
            {
                Dictionary<string, string> record = new Dictionary<string, string>();
                record["record_type"] = "entry_quota";
                record["method"] = MethodKey;
                record["method_no"] = MethodNo;
                record["entry_code"] = scope.MatchedEntryCode;
                record["entry_name"] = scope.EntryName ?? "";
                record["target_kind"] = kind;
                record["quota_code"] = code.Trim();
                record["quota_name"] = "";
                record["quota_unit"] = "";
                record["project_count"] = "0";
                record["source"] = "user";
                record["deleted"] = "1";
                record["last_seen"] = DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                string path = Path.Combine(LearningStore.FindDataDir(), "chapter-quota-library.jsonl");
                File.AppendAllText(path, LearningStore.ToJson(record) + Environment.NewLine, Encoding.UTF8);
                QuotaRecommendPanel.Log("ChapterLibrary user quota removed. entry=" + scope.MatchedEntryCode + " code=" + code);
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("ChapterLibrary user quota remove append failed: " + ex.Message);
            }
        }
    }

    internal sealed class MappingStore
    {
        private const int MaxSamplesPerBox = 30;
        private readonly string path;
        private readonly List<MappingBox> boxes = new List<MappingBox>();

        private MappingStore(string filePath)
        {
            path = filePath;
        }

        public static MappingStore Load(List<LearningRecord> records)
        {
            string filePath = Path.Combine(LearningStore.FindDataDir(), "mapping-boxes.jsonl");
            MappingStore store = new MappingStore(filePath);
            store.LoadFile();
            if (store.boxes.Count == 0)
            {
                store.ImportCorrections(records);
                if (store.boxes.Count > 0)
                {
                    store.Save();
                }
            }
            return store;
        }

        public List<RecommendationRow> Find(ExcelQuantityItem item, string categoryFilter, SearchIndexStore searchIndex, EntryScope scope)
        {
            ScoredBox best = null;
            foreach (MappingBox box in boxes)
            {
                if (!BoxAllowedByScope(box, scope))
                {
                    continue;
                }

                List<MappingTarget> allowedTargets = FilterTargetsByCategory(box.Targets, categoryFilter, searchIndex);
                if (allowedTargets.Count == 0)
                {
                    continue;
                }

                int score = box.Score(item);
                if (score >= 70 && (best == null || score > best.Score))
                {
                    best = new ScoredBox { Box = box, Score = score, AllowedTargets = allowedTargets };
                }
            }

            if (best == null)
            {
                return new List<RecommendationRow>();
            }

            return best.AllowedTargets
                .OrderBy(t => TargetSortRank(t.TargetKind, t.Code))
                .ThenBy(t => t.Code ?? "", StringComparer.OrdinalIgnoreCase)
                .Select(t => t.ToRecommendation(item, best.Score, best.Box.BoxId))
                .ToList();
        }

        private static List<MappingTarget> FilterTargetsByCategory(List<MappingTarget> targets, string categoryFilter, SearchIndexStore searchIndex)
        {
            List<MappingTarget> quotaTargets = targets
                .Where(t => String.Equals(String.IsNullOrWhiteSpace(t.TargetKind) ? QuotaEntry.GuessKind(t.Code) : t.TargetKind, "quota", StringComparison.OrdinalIgnoreCase))
                .ToList();
            List<MappingTarget> materialTargets = targets
                .Where(t => !String.Equals(String.IsNullOrWhiteSpace(t.TargetKind) ? QuotaEntry.GuessKind(t.Code) : t.TargetKind, "quota", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (quotaTargets.Count == 0)
            {
                return materialTargets;
            }

            List<MappingTarget> allowedQuotaTargets = quotaTargets
                .Where(t => searchIndex != null && searchIndex.IsMappingTargetAllowed(t.TargetKind, t.Code, categoryFilter))
                .ToList();
            if (allowedQuotaTargets.Count == 0)
            {
                return new List<MappingTarget>();
            }

            allowedQuotaTargets.AddRange(materialTargets);
            return allowedQuotaTargets;
        }

        // 严格条目模式下框可用的条件：已带当前条目标签，或全部 quota 类目标都在条目定额池内
        private static bool BoxAllowedByScope(MappingBox box, EntryScope scope)
        {
            if (scope == null || !scope.Strict)
            {
                return true;
            }

            if (box.EntryCodes.Contains(scope.Tag))
            {
                return true;
            }

            bool hasQuotaTarget = false;
            foreach (MappingTarget target in box.Targets)
            {
                string kind = String.IsNullOrWhiteSpace(target.TargetKind) ? QuotaEntry.GuessKind(target.Code) : target.TargetKind;
                if (!String.Equals(kind, "quota", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                hasQuotaTarget = true;
                if (!scope.Allows(kind, target.Code))
                {
                    return false;
                }
            }

            if (hasQuotaTarget)
            {
                return true;
            }

            // 纯材料框按材料码判断
            return box.Targets.Count > 0 && box.Targets.All(t => scope.Allows(String.IsNullOrWhiteSpace(t.TargetKind) ? QuotaEntry.GuessKind(t.Code) : t.TargetKind, t.Code));
        }

        public List<AiMappingCandidate> BuildDeepSeekCandidates(ExcelQuantityItem item, string categoryFilter, SearchIndexStore searchIndex, int limit, EntryScope scope)
        {
            if (item == null)
            {
                return new List<AiMappingCandidate>();
            }

            return boxes
                .Where(box => BoxAllowedByScope(box, scope))
                .Select(box => new
                {
                    Box = box,
                    Targets = FilterTargetsByCategory(box.Targets, categoryFilter, searchIndex),
                    Score = Math.Max(box.Score(item), box.LooseScore(item))
                })
                .Where(x => x.Targets.Count > 0 && x.Score >= 20)
                .OrderByDescending(x => x.Score)
                .Take(Math.Max(1, limit))
                .Select(x => new AiMappingCandidate
                {
                    BoxId = x.Box.BoxId,
                    LocalScore = x.Score,
                    SampleNames = x.Box.SampleNamesForPrompt(),
                    Targets = x.Targets
                })
                .ToList();
        }

        public void Accept(List<RecommendationRow> rows, EntryScope scope)
        {
            bool changed = false;
            foreach (IGrouping<string, RecommendationRow> group in rows
                .Where(r => r != null && r.Item != null && !String.IsNullOrWhiteSpace(r.QuotaCode))
                .GroupBy(r => LearningStore.BuildQuantitySignature(r.Item), StringComparer.OrdinalIgnoreCase))
            {
                RecommendationRow first = group.First();
                MappingBox box = null;
                string boxId = group.Select(r => r.BoxId).FirstOrDefault(id => !String.IsNullOrWhiteSpace(id));
                if (!String.IsNullOrWhiteSpace(boxId))
                {
                    box = boxes.FirstOrDefault(b => String.Equals(b.BoxId, boxId, StringComparison.OrdinalIgnoreCase));
                }

                if (box == null)
                {
                    box = FindOrCreateBox(group.Select(row => new QuotaEntry
                    {
                        TargetKind = String.IsNullOrWhiteSpace(row.TargetKind) ? QuotaEntry.GuessKind(row.QuotaCode) : row.TargetKind,
                        QuotaCode = row.QuotaCode,
                        QuotaName = row.QuotaName,
                        QuotaUnit = row.QuotaUnit
                    }).ToList());
                }

                MappingSample sample = box.FindOrCreateSample(first.Item.Name, first.Item.Unit);
                sample.Weight += 5;
                sample.AcceptedCount += 1;
                sample.LastUsedAt = Now();
                if (scope != null && scope.Strict)
                {
                    box.EntryCodes.Add(scope.Tag);
                }
                box.TrimSamples(MaxSamplesPerBox);
                changed = true;
            }

            if (changed)
            {
                Save();
            }
        }

        public void Correct(ExcelQuantityItem item, RecommendationRow oldRecommendation, List<QuotaEntry> selectedTargets, EntryScope scope)
        {
            if (item == null || selectedTargets == null || selectedTargets.Count == 0)
            {
                return;
            }

            Penalize(item, oldRecommendation, selectedTargets);
            MappingBox box = FindOrCreateBox(selectedTargets);
            MappingSample sample = box.FindOrCreateSample(item.Name, item.Unit);
            sample.Weight += 20;
            sample.CorrectedCount += 1;
            sample.LastUsedAt = Now();
            if (scope != null && scope.Strict)
            {
                box.EntryCodes.Add(scope.Tag);
            }
            box.TrimSamples(MaxSamplesPerBox);
            Save();
        }

        private void Penalize(ExcelQuantityItem item, RecommendationRow oldRecommendation, List<QuotaEntry> selectedTargets)
        {
            if (oldRecommendation == null || String.IsNullOrWhiteSpace(oldRecommendation.QuotaCode))
            {
                return;
            }

            string oldKey = (String.IsNullOrWhiteSpace(oldRecommendation.TargetKind) ? QuotaEntry.GuessKind(oldRecommendation.QuotaCode) : oldRecommendation.TargetKind) + ":" + oldRecommendation.QuotaCode.ToUpperInvariant();
            if (selectedTargets.Any(t => String.Equals(t.TargetKey, oldKey, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            foreach (MappingBox box in boxes)
            {
                if (!box.Targets.Any(t => String.Equals(t.TargetKey, oldKey, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                MappingSample sample = box.FindSample(item.Name, item.Unit);
                if (sample != null)
                {
                    sample.Weight = Math.Max(0, sample.Weight - 10);
                    sample.RejectedCount += 1;
                    sample.LastUsedAt = Now();
                }
            }
        }

        private MappingBox FindOrCreateBox(List<QuotaEntry> targets)
        {
            List<MappingTarget> normalized = targets
                .Where(t => !String.IsNullOrWhiteSpace(t.QuotaCode))
                .Select(t => new MappingTarget
                {
                    TargetKind = String.IsNullOrWhiteSpace(t.TargetKind) ? QuotaEntry.GuessKind(t.QuotaCode) : t.TargetKind,
                    Code = t.QuotaCode.Trim(),
                    Name = t.QuotaName ?? "",
                    Unit = t.QuotaUnit ?? ""
                })
                .GroupBy(t => t.TargetKey, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .OrderBy(t => TargetSortRank(t.TargetKind, t.Code))
                .ThenBy(t => t.Code ?? "", StringComparer.OrdinalIgnoreCase)
                .ToList();

            string boxId = BuildBoxId(normalized);
            MappingBox box = boxes.FirstOrDefault(b => String.Equals(b.BoxId, boxId, StringComparison.OrdinalIgnoreCase));
            if (box == null)
            {
                box = new MappingBox { BoxId = boxId };
                box.Targets.AddRange(normalized);
                boxes.Add(box);
            }
            else
            {
                foreach (MappingTarget target in normalized)
                {
                    if (!box.Targets.Any(t => String.Equals(t.TargetKey, target.TargetKey, StringComparison.OrdinalIgnoreCase)))
                    {
                        box.Targets.Add(target);
                    }
                }
            }

            return box;
        }

        private void ImportCorrections(List<LearningRecord> records)
        {
            foreach (IGrouping<string, LearningRecord> group in records
                .Where(r => r.IsCorrection && !String.IsNullOrWhiteSpace(r.QuotaCode))
                .GroupBy(r => r.QuantitySignature, StringComparer.OrdinalIgnoreCase))
            {
                List<QuotaEntry> targets = group.Select(r => new QuotaEntry
                {
                    TargetKind = QuotaEntry.GuessKind(r.QuotaCode),
                    QuotaCode = r.QuotaCode,
                    QuotaName = r.QuotaName,
                    QuotaUnit = r.QuotaUnit
                }).ToList();
                MappingBox box = FindOrCreateBox(targets);
                LearningRecord first = group.First();
                MappingSample sample = box.FindOrCreateSample(first.QuantityName, first.QuantityUnit);
                sample.Weight = Math.Max(sample.Weight, 30);
                sample.CorrectedCount += 1;
                sample.LastUsedAt = Now();
            }
        }

        private void LoadFile()
        {
            List<MappingBox> parsed = null;
            WithMappingBoxesLock(delegate
            {
                parsed = ParseFile(path);
            });
            boxes.Clear();
            boxes.AddRange(CanonicalizeBoxes(parsed));
        }

        private static List<MappingBox> ParseFile(string filePath)
        {
            List<MappingBox> result = new List<MappingBox>();
            if (!File.Exists(filePath))
            {
                return result;
            }

            Dictionary<string, MappingBox> byId = new Dictionary<string, MappingBox>(StringComparer.OrdinalIgnoreCase);
            foreach (string line in File.ReadAllLines(filePath, Encoding.UTF8))
            {
                if (String.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                string boxId = LearningStore.Get(values, "box_id");
                if (String.IsNullOrWhiteSpace(boxId))
                {
                    continue;
                }

                MappingBox box;
                if (!byId.TryGetValue(boxId, out box))
                {
                    box = new MappingBox { BoxId = boxId };
                    byId[boxId] = box;
                    result.Add(box);
                }

                foreach (string entryTag in LearningStore.Get(values, "entry_codes").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    box.EntryCodes.Add(entryTag.Trim());
                }

                // 类别由编号确定地推导（覆盖旧文件里把 ZLF/TLF 等辅助代号误存成 quota 的记录）
                string parsedCode = LearningStore.Get(values, "target_code");
                MappingTarget target = new MappingTarget
                {
                    TargetKind = QuotaEntry.GuessKind(parsedCode),
                    Code = parsedCode,
                    Name = LearningStore.Get(values, "target_name"),
                    Unit = LearningStore.Get(values, "target_unit")
                };
                if (!String.IsNullOrWhiteSpace(target.Code) && !box.Targets.Any(t => String.Equals(t.TargetKey, target.TargetKey, StringComparison.OrdinalIgnoreCase)))
                {
                    box.Targets.Add(target);
                }

                MappingSample sample = new MappingSample();
                sample.QuantityName = LearningStore.Get(values, "quantity_name");
                sample.QuantityUnit = LearningStore.Get(values, "quantity_unit");
                sample.Weight = ParseInt(LearningStore.Get(values, "weight"), 0);
                sample.AcceptedCount = ParseInt(LearningStore.Get(values, "accepted_count"), 0);
                sample.CorrectedCount = ParseInt(LearningStore.Get(values, "corrected_count"), 0);
                sample.RejectedCount = ParseInt(LearningStore.Get(values, "rejected_count"), 0);
                sample.LastUsedAt = LearningStore.Get(values, "last_used_at");
                if (!String.IsNullOrWhiteSpace(sample.QuantityName) && box.FindSample(sample.QuantityName, sample.QuantityUnit) == null)
                {
                    box.Samples.Add(sample);
                }
            }

            return result;
        }

        // 旧版 box_id 由 String.GetHashCode 生成，跨进程位数不稳定且可能碰撞；
        // 加载时按目标组合重算稳定 ID，同一组合的旧框自动合并，实现旧文件无感迁移。
        private static List<MappingBox> CanonicalizeBoxes(List<MappingBox> parsed)
        {
            List<MappingBox> result = new List<MappingBox>();
            Dictionary<string, MappingBox> byId = new Dictionary<string, MappingBox>(StringComparer.OrdinalIgnoreCase);
            foreach (MappingBox box in parsed ?? new List<MappingBox>())
            {
                if (box.Targets.Count == 0)
                {
                    continue;
                }

                string canonicalId = BuildBoxId(box.Targets);
                MappingBox existing;
                if (!byId.TryGetValue(canonicalId, out existing))
                {
                    box.BoxId = canonicalId;
                    byId[canonicalId] = box;
                    result.Add(box);
                    continue;
                }

                MergeBox(existing, box);
            }

            return result;
        }

        private static void MergeBox(MappingBox into, MappingBox from)
        {
            into.EntryCodes.UnionWith(from.EntryCodes);
            foreach (MappingTarget target in from.Targets)
            {
                if (!into.Targets.Any(t => String.Equals(t.TargetKey, target.TargetKey, StringComparison.OrdinalIgnoreCase)))
                {
                    into.Targets.Add(target);
                }
            }

            foreach (MappingSample sample in from.Samples)
            {
                MappingSample existing = into.FindSample(sample.QuantityName, sample.QuantityUnit);
                if (existing == null)
                {
                    into.Samples.Add(sample);
                }
                else if (String.Compare(sample.LastUsedAt ?? "", existing.LastUsedAt ?? "", StringComparison.Ordinal) > 0)
                {
                    existing.Weight = sample.Weight;
                    existing.AcceptedCount = sample.AcceptedCount;
                    existing.CorrectedCount = sample.CorrectedCount;
                    existing.RejectedCount = sample.RejectedCount;
                    existing.LastUsedAt = sample.LastUsedAt;
                }
            }
        }

        private void Save()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            WithMappingBoxesLock(delegate
            {
                MergeFromDisk();
                WriteFile();
            });
        }

        // Excel联动AI匹配和扶正训练器也会写这个文件；整文件重写前先合并磁盘上的新增记录，避免覆盖丢失。
        private void MergeFromDisk()
        {
            foreach (MappingBox diskBox in CanonicalizeBoxes(ParseFile(path)))
            {
                MappingBox memory = boxes.FirstOrDefault(b => String.Equals(b.BoxId, diskBox.BoxId, StringComparison.OrdinalIgnoreCase));
                if (memory == null)
                {
                    boxes.Add(diskBox);
                }
                else
                {
                    MergeBox(memory, diskBox);
                }
            }
        }

        private void WriteFile()
        {
            string temp = path + ".tmp";
            using (StreamWriter writer = new StreamWriter(temp, false, Encoding.UTF8))
            {
                foreach (MappingBox box in boxes)
                {
                    box.TrimSamples(MaxSamplesPerBox);
                    foreach (MappingTarget target in box.Targets
                        .OrderBy(t => TargetSortRank(t.TargetKind, t.Code))
                        .ThenBy(t => t.Code ?? "", StringComparer.OrdinalIgnoreCase))
                    {
                        foreach (MappingSample sample in box.Samples)
                        {
                            Dictionary<string, string> row = new Dictionary<string, string>();
                            row["record_type"] = "mapping_box";
                            row["box_id"] = box.BoxId;
                            row["target_kind"] = String.IsNullOrWhiteSpace(target.TargetKind) ? QuotaEntry.GuessKind(target.Code) : target.TargetKind;
                            row["target_code"] = target.Code;
                            row["target_name"] = target.Name;
                            row["target_unit"] = target.Unit;
                            row["quantity_name"] = sample.QuantityName;
                            row["quantity_unit"] = sample.QuantityUnit;
                            row["weight"] = sample.Weight.ToString(CultureInfo.InvariantCulture);
                            row["accepted_count"] = sample.AcceptedCount.ToString(CultureInfo.InvariantCulture);
                            row["corrected_count"] = sample.CorrectedCount.ToString(CultureInfo.InvariantCulture);
                            row["rejected_count"] = sample.RejectedCount.ToString(CultureInfo.InvariantCulture);
                            row["last_used_at"] = sample.LastUsedAt;
                            if (box.EntryCodes.Count > 0)
                            {
                                row["entry_codes"] = String.Join(",", box.EntryCodes.OrderBy(t => t, StringComparer.OrdinalIgnoreCase).ToArray());
                            }
                            writer.WriteLine(LearningStore.ToJson(row));
                        }
                    }
                }
            }

            if (File.Exists(path))
            {
                File.Delete(path);
            }
            File.Move(temp, path);
        }

        private const string MappingBoxesMutexName = "RecoQuotaData.mapping-boxes.lock";

        // 学习库有多个写入方（本窗口扶正、RecoExpandPanel 的 Excel联动与训练器），
        // 用跨程序集一致的命名互斥锁串行化读改写，名称必须与 RecoExpandPanel 保持一致。
        private static void WithMappingBoxesLock(Action action)
        {
            Mutex mutex = new Mutex(false, MappingBoxesMutexName);
            bool acquired = false;
            try
            {
                try
                {
                    acquired = mutex.WaitOne(5000);
                }
                catch (AbandonedMutexException)
                {
                    acquired = true;
                }

                action();
            }
            finally
            {
                if (acquired)
                {
                    mutex.ReleaseMutex();
                }
                mutex.Dispose();
            }
        }

        // 与 RecoExpandPanel 的 BuildStableMappingBoxId 使用同一套规则：对小写化目标键做 SHA1。
        private static string BuildBoxId(List<MappingTarget> targets)
        {
            string raw = String.Join("|", targets
                .OrderBy(t => t.TargetKey, StringComparer.OrdinalIgnoreCase)
                .Select(t => t.TargetKey)
                .ToArray());
            using (System.Security.Cryptography.SHA1 sha = System.Security.Cryptography.SHA1.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(raw.ToLowerInvariant()));
                StringBuilder builder = new StringBuilder("box-");
                for (int i = 0; i < 8; i++)
                {
                    builder.Append(hash[i].ToString("x2", CultureInfo.InvariantCulture));
                }

                return builder.ToString();
            }
        }

        public static int TargetSortRank(string targetKind, string code)
        {
            string kind = String.IsNullOrWhiteSpace(targetKind) ? QuotaEntry.GuessKind(code) : targetKind;
            return String.Equals(kind, "quota", StringComparison.OrdinalIgnoreCase) ? 0 : 1;
        }

        private static int ParseInt(string text, int fallback)
        {
            int value;
            return Int32.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value) ? value : fallback;
        }

        private static string Now()
        {
            return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        private sealed class ScoredBox
        {
            public MappingBox Box;
            public int Score;
            public List<MappingTarget> AllowedTargets;
        }
    }

    internal sealed class MappingBox
    {
        public string BoxId;
        public readonly List<MappingTarget> Targets = new List<MappingTarget>();
        public readonly List<MappingSample> Samples = new List<MappingSample>();
        // 章节条目标签（"2020:0101-01" 形式，method:条目编号），用于按条目分类对应框
        public readonly HashSet<string> EntryCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public int Score(ExcelQuantityItem item)
        {
            int best = 0;
            foreach (MappingSample sample in Samples)
            {
                if (!CanUseSampleForItem(item, sample))
                {
                    continue;
                }

                if (!TextMatcher.HasStrongNamePairMatch(item.Name, sample.QuantityName))
                {
                    continue;
                }

                if (TextMatcher.IsSteelOnlyAgainstConcrete(item.Name, sample.QuantityName))
                {
                    continue;
                }

                int similarity = TextMatcher.NamePairScore(item.Name, sample.QuantityName);
                if (similarity <= 0)
                {
                    continue;
                }

                int score = similarity + Math.Min(40, sample.Weight);
                if (RecommendDialog.UnitCompatibleForIndex(sample.QuantityUnit, item.Unit))
                {
                    score += 12;
                }
                best = Math.Max(best, score);
            }

            return best;
        }

        public int LooseScore(ExcelQuantityItem item)
        {
            int best = 0;
            string raw = item == null ? "" : item.RawRowText;
            foreach (MappingSample sample in Samples)
            {
                if (!CanUseSampleForItem(item, sample))
                {
                    continue;
                }

                int score = Math.Max(
                    TextMatcher.NamePairScore(item == null ? "" : item.Name, sample.QuantityName),
                    TextMatcher.NamePairScore(raw, sample.QuantityName) / 2);
                if (RecommendDialog.UnitCompatibleForIndex(sample.QuantityUnit, item == null ? "" : item.Unit))
                {
                    score += 8;
                }
                best = Math.Max(best, score + Math.Min(20, sample.Weight / 2));
            }

            return best;
        }

        private bool CanUseSampleForItem(ExcelQuantityItem item, MappingSample sample)
        {
            if (item == null || sample == null)
            {
                return false;
            }

            if (!String.IsNullOrWhiteSpace(item.Unit) &&
                !String.IsNullOrWhiteSpace(sample.QuantityUnit) &&
                !RecommendDialog.UnitCompatibleForIndex(sample.QuantityUnit, item.Unit))
            {
                return false;
            }

            string targetText = String.Join(" ", Targets.Select(t => (t == null ? "" : (t.Code ?? "") + " " + (t.Name ?? "") + " " + (t.Unit ?? ""))).ToArray());
            return !HasEngineeringProcessConflict(item.Name + " " + item.RawRowText, sample.QuantityName + " " + targetText);
        }

        private static bool HasEngineeringProcessConflict(string quantityText, string candidateText)
        {
            string q = TextMatcher.Normalize(quantityText);
            string c = TextMatcher.Normalize(candidateText);
            bool qSheetPileRemoval = ContainsAny(q, "\u62c9\u68ee", "\u94a2\u677f\u6869") && ContainsAny(q, "\u62d4\u9664", "\u62c6\u9664", "\u62d4\u51fa");
            bool qHasGrout = ContainsAny(q, "\u6ce8\u6d46", "\u6c34\u6ce5\u6d46", "\u586b\u5145", "\u5145\u586b", "\u56de\u586b");
            bool cHasGrout = ContainsAny(c, "\u6ce8\u6d46", "\u6c34\u6ce5\u6d46", "\u586b\u5145", "\u5145\u586b");
            if (qSheetPileRemoval && !qHasGrout && cHasGrout)
            {
                return true;
            }

            bool qConcreteOnly = TextMatcher.IsConcreteQuantityName(q) && !TextMatcher.IsSteelQuantityName(q);
            bool cSteelOnly = TextMatcher.IsSteelQuantityName(c) && !TextMatcher.IsConcreteQuantityName(c);
            if (qConcreteOnly && cSteelOnly)
            {
                return true;
            }

            return false;
        }

        private static bool ContainsAny(string text, params string[] keywords)
        {
            foreach (string keyword in keywords ?? new string[0])
            {
                if (!String.IsNullOrWhiteSpace(keyword) && (text ?? "").Contains(TextMatcher.Normalize(keyword)))
                {
                    return true;
                }
            }
            return false;
        }

        public string SampleNamesForPrompt()
        {
            return String.Join("；", Samples
                .OrderByDescending(s => s.Weight)
                .ThenByDescending(s => s.LastUsedAt ?? "")
                .Take(8)
                .Select(s => s.QuantityName)
                .Where(n => !String.IsNullOrWhiteSpace(n))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray());
        }

        public MappingSample FindOrCreateSample(string name, string unit)
        {
            MappingSample sample = FindSample(name, unit);
            if (sample != null)
            {
                return sample;
            }

            sample = new MappingSample { QuantityName = name ?? "", QuantityUnit = unit ?? "", Weight = 10, LastUsedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) };
            Samples.Add(sample);
            return sample;
        }

        public MappingSample FindSample(string name, string unit)
        {
            string signature = LearningStore.BuildQuantitySignature(name, unit);
            return Samples.FirstOrDefault(s => String.Equals(LearningStore.BuildQuantitySignature(s.QuantityName, s.QuantityUnit), signature, StringComparison.OrdinalIgnoreCase));
        }

        public void TrimSamples(int maxSamples)
        {
            while (Samples.Count > maxSamples)
            {
                MappingSample remove = Samples
                    .OrderBy(s => s.Weight)
                    .ThenBy(s => s.LastUsedAt ?? "")
                    .First();
                Samples.Remove(remove);
            }
        }
    }

    internal sealed class MappingTarget
    {
        public string TargetKind;
        public string Code;
        public string Name;
        public string Unit;

        public string TargetKey
        {
            get { return (String.IsNullOrWhiteSpace(TargetKind) ? QuotaEntry.GuessKind(Code) : TargetKind) + ":" + (Code ?? "").Trim().ToUpperInvariant(); }
        }

        public RecommendationRow ToRecommendation(ExcelQuantityItem item, int score, string boxId)
        {
            RecommendationRow row = new RecommendationRow();
            row.Item = item;
            row.QuotaCode = Code;
            row.QuotaName = Name;
            row.QuotaUnit = Unit;
            row.ConvertedValueText = RecommendDialog.ConvertQuantityForIndex(item.ValueText, item.Unit, Unit);
            row.Score = score;
            row.Reason = "\u5b9a\u989d\u5bf9\u5e94\u6846\u6743\u91cd\u5339\u914d";
            row.Source = "mapping";
            row.BoxId = boxId;
            row.TargetKind = String.IsNullOrWhiteSpace(TargetKind) ? QuotaEntry.GuessKind(Code) : TargetKind;
            return row;
        }
    }

    internal sealed class MappingSample
    {
        public string QuantityName;
        public string QuantityUnit;
        public int Weight;
        public int AcceptedCount;
        public int CorrectedCount;
        public int RejectedCount;
        public string LastUsedAt;
    }

    internal static class TextMatcher
    {
        public static string Normalize(string text)
        {
            return (text ?? "")
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("（", "(")
                .Replace("）", ")")
                .Replace("，", ",")
                .Replace("、", " ")
                .Trim()
                .ToLowerInvariant();
        }

        public static int Similarity(string left, string right)
        {
            return Math.Min(100, NamePairScore(left, right));
        }

        public static int NamePairScore(string left, string right)
        {
            string l = Normalize(left).Replace(" ", "");
            string r = Normalize(right).Replace(" ", "");
            if (String.IsNullOrEmpty(l) || String.IsNullOrEmpty(r))
            {
                return 0;
            }

            if (l == r)
            {
                return 120;
            }
            if (l.Contains(r) || r.Contains(l))
            {
                return 95;
            }

            List<string> leftTokens = Keywords(l).Distinct().ToList();
            if (leftTokens.Count == 0)
            {
                return 0;
            }

            int score = 0;
            int possible = 0;
            int hits = 0;
            foreach (string token in leftTokens)
            {
                int tokenScore = PairTokenScore(token);
                possible += tokenScore;
                if (r.Contains(token))
                {
                    hits++;
                    score += tokenScore;
                }
            }

            if (hits == 0)
            {
                return 0;
            }

            score += (int)Math.Round(18.0 * hits / leftTokens.Count);
            return possible > 0 && score > 115 ? 115 : score;
        }

        public static bool HasStrongNamePairMatch(string left, string right)
        {
            string l = Normalize(left).Replace(" ", "");
            string r = Normalize(right).Replace(" ", "");
            if (String.IsNullOrEmpty(l) || String.IsNullOrEmpty(r))
            {
                return false;
            }

            if (l == r || l.Contains(r) || r.Contains(l))
            {
                return true;
            }

            bool steelConcreteBlocked = IsSteelOnlyAgainstConcrete(l, r);
            foreach (string token in Keywords(l).Distinct())
            {
                if (token.Length < 2 || !r.Contains(token))
                {
                    continue;
                }

                if (steelConcreteBlocked && String.Equals(token, "钢筋", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                return true;
            }

            return false;
        }

        public static bool IsSteelOnlyAgainstConcrete(string left, string right)
        {
            string l = Normalize(left);
            string r = Normalize(right);
            bool leftConcrete = IsConcreteQuantityName(l);
            bool rightConcrete = IsConcreteQuantityName(r);
            return (IsSteelQuantityName(l) && !leftConcrete && rightConcrete) ||
                (IsSteelQuantityName(r) && !rightConcrete && leftConcrete);
        }

        public static bool IsSteelQuantityName(string text)
        {
            string value = Normalize(text);
            return value.Contains("钢筋") ||
                value.Contains("hpb") ||
                value.Contains("hrb") ||
                value.Contains("圆钢") ||
                value.Contains("螺纹");
        }

        public static bool IsConcreteQuantityName(string text)
        {
            string value = Normalize(text);
            return value.Contains("混凝土") ||
                value.Contains("砼") ||
                value.Contains("商品混凝土");
        }

        private static int PairTokenScore(string token)
        {
            if (String.IsNullOrWhiteSpace(token))
            {
                return 0;
            }

            if (HasAsciiOrDigit(token))
            {
                return token.Length >= 3 ? 28 : 8;
            }

            if (token.Length == 1)
            {
                return 4;
            }
            if (token.Length == 2)
            {
                return 24;
            }
            if (token.Length == 3)
            {
                return 38;
            }
            return 50;
        }

        public static IEnumerable<string> Keywords(string text)
        {
            string normalized = Normalize(text);
            foreach (string part in normalized.Split(new char[] { ' ', '/', ',', ';', '\t', '(', ')', '[', ']', '+', '-', '*', '=' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string token = part.Trim();
                if (token.Length < 2 || IsNumberLike(token))
                {
                    continue;
                }

                yield return token;
                foreach (string segment in SplitAlphaNumericAndChinese(token))
                {
                    if (segment.Length >= 2 && !IsNumberLike(segment))
                    {
                        yield return segment;
                    }
                }
                if (ContainsChinese(token))
                {
                    for (int i = 0; i + 2 <= token.Length; i++)
                    {
                        string gram = token.Substring(i, 2);
                        if (!IsNumberLike(gram))
                        {
                            yield return gram;
                        }
                    }
                    for (int i = 0; i + 3 <= token.Length; i++)
                    {
                        string gram = token.Substring(i, 3);
                        if (!IsNumberLike(gram))
                        {
                            yield return gram;
                        }
                    }
                    for (int i = 0; i < token.Length; i++)
                    {
                        string gram = token.Substring(i, 1);
                        if (IsPureChinese(gram))
                        {
                            yield return gram;
                        }
                    }
                }
            }
        }

        public static bool HasAsciiOrDigit(string text)
        {
            foreach (char ch in text ?? "")
            {
                if ((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || Char.IsDigit(ch))
                {
                    return true;
                }
            }
            return false;
        }

        public static bool IsPureChinese(string text)
        {
            bool hasChinese = false;
            foreach (char ch in text ?? "")
            {
                if (ch >= 0x4e00 && ch <= 0x9fff)
                {
                    hasChinese = true;
                    continue;
                }

                return false;
            }
            return hasChinese;
        }

        private static IEnumerable<string> SplitAlphaNumericAndChinese(string token)
        {
            StringBuilder builder = new StringBuilder();
            int lastKind = 0;
            foreach (char ch in token ?? "")
            {
                int kind = Char.IsLetterOrDigit(ch) && !(ch >= 0x4e00 && ch <= 0x9fff) ? 1 : ((ch >= 0x4e00 && ch <= 0x9fff) ? 2 : 0);
                if (kind == 0)
                {
                    if (builder.Length > 0)
                    {
                        yield return builder.ToString();
                        builder.Length = 0;
                    }
                    lastKind = 0;
                    continue;
                }

                if (builder.Length > 0 && kind != lastKind)
                {
                    yield return builder.ToString();
                    builder.Length = 0;
                }

                builder.Append(ch);
                lastKind = kind;
            }

            if (builder.Length > 0)
            {
                yield return builder.ToString();
            }
        }

        private static bool ContainsChinese(string text)
        {
            foreach (char ch in text ?? "")
            {
                if (ch >= 0x4e00 && ch <= 0x9fff)
                {
                    return true;
                }
            }
            return false;
        }

        private static bool IsNumberLike(string text)
        {
            decimal value;
            return Decimal.TryParse((text ?? "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        public static bool IsNumberLikeToken(string text)
        {
            return IsNumberLike(text);
        }
    }

    internal static class LearningStore
    {
        public static List<LearningRecord> Load()
        {
            string path = FindLearningPath();
            if (String.IsNullOrEmpty(path))
            {
                return new List<LearningRecord>();
            }

            List<LearningRecord> records = new List<LearningRecord>();
            foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
            {
                if (String.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                Dictionary<string, string> item = ParseFlatJson(line);
                LearningRecord record = new LearningRecord();
                record.IsCorrection = String.Equals(Get(item, "record_type"), "correction", StringComparison.OrdinalIgnoreCase)
                    || String.Equals(Get(item, "user_action"), "correction", StringComparison.OrdinalIgnoreCase);
                record.ProjectName = Get(item, "project_name");
                record.BudgetFile = Get(item, "budget_file");
                record.BudgetGroup = Get(item, "budget_group");
                record.QuotaCode = Get(item, "quota_code");
                record.QuotaName = Get(item, "quota_name");
                record.QuotaUnit = Get(item, "quota_unit");
                record.QuantitySection = Get(item, "quantity_section");
                record.QuantityName = Get(item, "quantity_name");
                record.QuantityUnit = Get(item, "quantity_unit");
                record.QuantitySignature = Get(item, "quantity_signature");
                if (String.IsNullOrWhiteSpace(record.QuantitySignature))
                {
                    record.QuantitySignature = BuildQuantitySignature(record.QuantityName, record.QuantityUnit);
                }
                record.MatchReason = Get(item, "match_reason");
                int parsed;
                record.MatchScore = Int32.TryParse(Get(item, "match_score"), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) ? parsed : 0;
                if (!String.IsNullOrWhiteSpace(record.QuotaCode))
                {
                    records.Add(record);
                }
            }

            return records;
        }

        public static void BackupLearningFileIfNeeded()
        {
            try
            {
                string path = FindLearningPath();
                if (String.IsNullOrWhiteSpace(path) || !File.Exists(path))
                {
                    return;
                }

                string directory = Path.GetDirectoryName(path);
                if (Directory.GetFiles(directory, "learning.jsonl.*.bak").Length > 0)
                {
                    return;
                }

                string marker = Path.Combine(directory, "learning.jsonl." + DateTime.Now.ToString("yyyyMMddHHmmss", CultureInfo.InvariantCulture) + ".bak");
                File.Copy(path, marker, false);
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("Backup learning.jsonl failed: " + ex.Message);
            }
        }

        public static void ReplaceCorrections(ExcelQuantityItem item, List<QuotaEntry> quotas)
        {
            string signature = BuildQuantitySignature(item);
            List<string> paths = FindLearningPaths();
            if (paths.Count == 0)
            {
                string baseDir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location);
                paths.Add(Path.Combine(baseDir, "RecoQuotaData", "learning.jsonl"));
            }

            foreach (string path in paths)
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path));
                List<string> existing = File.Exists(path)
                    ? File.ReadAllLines(path, Encoding.UTF8).ToList()
                    : new List<string>();

                List<string> kept = new List<string>();
                foreach (string line in existing)
                {
                    if (String.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    Dictionary<string, string> values = ParseFlatJson(line);
                    bool isCorrection = String.Equals(Get(values, "record_type"), "correction", StringComparison.OrdinalIgnoreCase)
                        || String.Equals(Get(values, "user_action"), "correction", StringComparison.OrdinalIgnoreCase);
                    string lineSignature = Get(values, "quantity_signature");
                    if (String.IsNullOrWhiteSpace(lineSignature))
                    {
                        lineSignature = BuildQuantitySignature(Get(values, "quantity_name"), Get(values, "quantity_unit"));
                    }

                    if (isCorrection && String.Equals(lineSignature, signature, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    kept.Add(line);
                }

                foreach (QuotaEntry quota in quotas)
                {
                    Dictionary<string, string> record = new Dictionary<string, string>();
                    record["record_type"] = "correction";
                    record["user_action"] = "correction";
                    record["quantity_signature"] = signature;
                    record["quantity_name"] = item.Name;
                    record["quantity_unit"] = item.Unit;
                    record["quantity_section"] = item.SectionName;
                    record["quota_code"] = quota.QuotaCode;
                    record["quota_name"] = quota.QuotaName;
                    record["quota_unit"] = quota.QuotaUnit;
                    record["match_score"] = "1000";
                    record["match_reason"] = "\u4eba\u5de5\u6276\u6b63";
                    record["updated_at"] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                    kept.Add(ToJson(record));
                }

                File.WriteAllLines(path, kept.ToArray(), Encoding.UTF8);
            }
        }

        public static string BuildQuantitySignature(ExcelQuantityItem item)
        {
            return BuildQuantitySignature(item == null ? "" : item.Name, item == null ? "" : item.Unit);
        }

        public static string BuildQuantitySignature(string name, string unit)
        {
            return NormalizeForSignature(name) + "|" + NormalizeForSignature(unit);
        }

        private static string FindLearningPath()
        {
            List<string> paths = FindLearningPaths();
            return paths.Count == 0 ? "" : paths[0];
        }

        internal static string FindDataDir()
        {
            string baseDir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location);
            return Path.Combine(baseDir, "RecoQuotaData");
        }

        private static List<string> FindLearningPaths()
        {
            List<string> paths = new List<string>();
            string baseDir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location);
            string[] candidates = new string[]
            {
                Path.Combine(baseDir, "RecoQuotaData", "learning.jsonl"),
                Path.Combine(Path.GetDirectoryName(baseDir), "RecoQuotaData", "learning.jsonl")
            };

            foreach (string candidate in candidates)
            {
                if (File.Exists(candidate))
                {
                    paths.Add(candidate);
                }
            }

            return paths;
        }

        internal static string ToJson(Dictionary<string, string> values)
        {
            StringBuilder builder = new StringBuilder();
            builder.Append('{');
            bool first = true;
            foreach (KeyValuePair<string, string> pair in values)
            {
                if (!first)
                {
                    builder.Append(',');
                }

                first = false;
                builder.Append('"').Append(EscapeJson(pair.Key)).Append('"').Append(':')
                    .Append('"').Append(EscapeJson(pair.Value)).Append('"');
            }

            builder.Append('}');
            return builder.ToString();
        }

        private static string EscapeJson(string value)
        {
            StringBuilder builder = new StringBuilder();
            foreach (char ch in value ?? "")
            {
                switch (ch)
                {
                    case '\\':
                        builder.Append("\\\\");
                        break;
                    case '"':
                        builder.Append("\\\"");
                        break;
                    case '\r':
                        builder.Append("\\r");
                        break;
                    case '\n':
                        builder.Append("\\n");
                        break;
                    case '\t':
                        builder.Append("\\t");
                        break;
                    default:
                        builder.Append(ch);
                        break;
                }
            }

            return builder.ToString();
        }

        private static string NormalizeForSignature(string text)
        {
            return (text ?? "").Replace("\r", "").Replace("\n", "").Replace(" ", "").Trim().ToLowerInvariant();
        }

        internal static string Get(Dictionary<string, string> values, string key)
        {
            string value;
            return values.TryGetValue(key, out value) ? value : "";
        }

        internal static Dictionary<string, string> ParseFlatJson(string line)
        {
            Dictionary<string, string> result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            int index = 0;
            SkipWhitespace(line, ref index);
            if (index < line.Length && line[index] == '{')
            {
                index++;
            }

            while (index < line.Length)
            {
                SkipWhitespace(line, ref index);
                if (index >= line.Length || line[index] == '}')
                {
                    break;
                }

                string key = ReadJsonString(line, ref index);
                SkipWhitespace(line, ref index);
                if (index < line.Length && line[index] == ':')
                {
                    index++;
                }
                SkipWhitespace(line, ref index);
                string value = ReadJsonString(line, ref index);
                result[key] = value;
                SkipWhitespace(line, ref index);
                if (index < line.Length && line[index] == ',')
                {
                    index++;
                }
            }

            return result;
        }

        private static string ReadJsonString(string text, ref int index)
        {
            StringBuilder builder = new StringBuilder();
            if (index < text.Length && text[index] == '"')
            {
                index++;
            }

            while (index < text.Length)
            {
                char ch = text[index++];
                if (ch == '"')
                {
                    break;
                }

                if (ch == '\\' && index < text.Length)
                {
                    char escaped = text[index++];
                    if (escaped == 'n')
                    {
                        builder.Append('\n');
                    }
                    else if (escaped == 'r')
                    {
                        builder.Append('\r');
                    }
                    else if (escaped == 't')
                    {
                        builder.Append('\t');
                    }
                    else
                    {
                        builder.Append(escaped);
                    }
                }
                else
                {
                    builder.Append(ch);
                }
            }

            return builder.ToString();
        }

        private static void SkipWhitespace(string text, ref int index)
        {
            while (index < text.Length && Char.IsWhiteSpace(text[index]))
            {
                index++;
            }
        }
    }
}
