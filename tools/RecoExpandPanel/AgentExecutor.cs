using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        private static readonly Dictionary<Form, List<AgentUndoRecord>> AgentUndoStacks = new Dictionary<Form, List<AgentUndoRecord>>();
        private static readonly Dictionary<Form, List<AgentUndoRecord>> AgentRedoStacks = new Dictionary<Form, List<AgentUndoRecord>>();

        // 与迁移工具相同的诊断账号，仅当主程序连接串克隆失败时兜底使用。
        private const string AgentDbUser = "reco";
        private const string AgentDbPassword = "Des_Reco_2006";

        // 删除定额行时需要同事务清理的从属表（均含 定额序号 列，主程序的计算缓存）。
        private static readonly string[] AgentQuotaDependentTables = new string[]
        {
            "定额唯一值", "清单_定额单价分析", "定额小计结尾计算结果"
        };

        // 新建单元（复制总概算）时需要按 总概算序号 复制的表（基线diff实测主程序复制行为）。
        private static readonly string[] AgentUnitCopyTables = new string[]
        {
            "总概算条目", "单项概算信息", "定额输入", "施工监理费率", "贷款利息"
        };

        private sealed class AgentPlanException : Exception
        {
            public AgentPlanException(string message)
                : base(message)
            {
            }
        }

        // 在 UI 线程采集的选中状态快照，供后台线程使用（后台不许碰窗体控件）。
        private sealed class AgentSelectionSnapshot
        {
            public string ItemNo = "";
            public string ItemName = "";
            public List<QuotaKey> QuotaKeys = new List<QuotaKey>();
            public List<string> QuotaCodes = new List<string>();
            public long CurrentUnitId;   // 0 = 未识别
            public string CurrentUnitCode = "";   // 当前单元总概算编号（如 ZGS_05）
            public Dictionary<long, string> UnitCodeMap = new Dictionary<long, string>();   // 全部单元 总概算序号->总概算编号
        }

        // 解析出的一条目标定额行（定额输入表），带 multiply/set 需要的原始字段。
        private sealed class AgentTargetRow
        {
            public long QuotaSequence;
            public string ItemNo;
            public string QuotaCode;
            public string QuantityInput;
            public object Quantity;      // double 或 DBNull
            public object UnitPrice;     // double 或 DBNull
            public string AdjustText;    // 定额调整(ntext)
            public long UnitId;
        }

        // 通用字段更新：一条 UPDATE，支持任意表/任意列。
        private sealed class AgentFieldUpdate
        {
            public string Table;
            public string KeyClause;     // 如 "定额序号=@k0"
            public Dictionary<string, object> KeyValues = new Dictionary<string, object>();
            public long QuotaSequence;   // 定额输入行用于主程序内存同步，0=不适用
            public Dictionary<string, object> NewValues = new Dictionary<string, object>();
            public Dictionary<string, object> OldValues = new Dictionary<string, object>();
            // 展示
            public string Action;
            public long UnitId;
            public string ItemNo;
            public string QuotaCode;
            public string OldDisplay;
            public string NewDisplay;
        }

        private sealed class AgentDeleteRow
        {
            public long QuotaSequence;
            public long UnitId;
            public string ItemNo;
            public string QuotaCode;
        }

        private sealed class AgentInsertGroup
        {
            public string ItemNo;
            public List<AgentQuotaInput> Quotas = new List<AgentQuotaInput>();
        }

        private sealed class AgentUnitCopySpec
        {
            public long SourceUnitId;
            public string SourceZgsNo;
            public string SourceName;
            public string NewName;
            public string NewZgsNo;
        }

        // 纯展示行（预览表格）。
        private sealed class AgentPlanRow
        {
            public string Action;
            public long UnitId;
            public string ItemNo;
            public string QuotaCode;
            public string OldValue;
            public string NewValue;
        }

        private sealed class AgentPlan
        {
            public List<AgentCommand> Commands = new List<AgentCommand>();
            public List<AgentFieldUpdate> FieldUpdates = new List<AgentFieldUpdate>();
            public List<AgentDeleteRow> Deletes = new List<AgentDeleteRow>();
            public List<AgentInsertGroup> Inserts = new List<AgentInsertGroup>();
            public List<AgentUnitCopySpec> UnitCopies = new List<AgentUnitCopySpec>();
            public List<AgentPlanRow> PreviewRows = new List<AgentPlanRow>();
            public Dictionary<long, string> UnitCodes = new Dictionary<long, string>();   // 总概算序号 -> 总概算编号
            public List<string> Warnings = new List<string>();
            public bool NeedsRecalc;
            public string Summary = "";
            public bool IsUndo;
            public bool IsRedo;
            public AgentUndoRecord UndoRecord;

            public int AffectedCount
            {
                get { return PreviewRows.Count; }
            }
        }

        private sealed class AgentUndoRow
        {
            public string Kind;   // F=字段还原 D=重插整行 I=删除新插入行 CU=删除新建单元
            public long QuotaSequence;
            public long UnitId;
            public string Table;
            public string KeyClause;
            public Dictionary<string, object> KeyValues;
            public Dictionary<string, object> OldValues;
            public Dictionary<string, object> NewValues;   // 重做用：F 类的正向值
            public Dictionary<string, object> FullRow;
        }

        private sealed class AgentUndoRecord
        {
            public string Summary;
            public DateTime Time;
            public List<AgentUndoRow> Rows = new List<AgentUndoRow>();
        }

        private static List<AgentUndoRecord> GetAgentUndoStack(Form mainForm)
        {
            List<AgentUndoRecord> stack;
            if (!AgentUndoStacks.TryGetValue(mainForm, out stack))
            {
                stack = new List<AgentUndoRecord>();
                AgentUndoStacks[mainForm] = stack;
            }

            return stack;
        }

        private static List<AgentUndoRecord> GetAgentRedoStack(Form mainForm)
        {
            List<AgentUndoRecord> stack;
            if (!AgentRedoStacks.TryGetValue(mainForm, out stack))
            {
                stack = new List<AgentUndoRecord>();
                AgentRedoStacks[mainForm] = stack;
            }

            return stack;
        }

        // 重做只支持纯字段更新(F)的记录；含删除/插入/新建单元的不入重做栈。
        private static bool IsAgentRecordRedoable(AgentUndoRecord record)
        {
            foreach (AgentUndoRow row in record.Rows)
            {
                if (row.Kind != "F" || row.NewValues == null)
                {
                    return false;
                }
            }

            return record.Rows.Count > 0;
        }

        // ===== 连接与选中快照 =====

        private static SqlConnection AgentCreateWorkConnection(Form mainForm)
        {
            SqlConnection host = GetProjectConnection(mainForm);
            if (host == null)
            {
                throw new AgentPlanException("没有找到当前项目数据库连接，请先打开一个项目。");
            }

            string hostString = host.ConnectionString;
            string database = host.Database;
            try
            {
                SqlConnection conn = new SqlConnection(hostString);
                conn.Open();
                if (!String.IsNullOrEmpty(database) && !String.Equals(conn.Database, database, StringComparison.OrdinalIgnoreCase))
                {
                    conn.ChangeDatabase(database);
                }

                return conn;
            }
            catch (Exception ex)
            {
                Log("Agent clone connection failed, fallback to reco: " + ex.Message);
            }

            SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder(hostString);
            string fallback = "Server=" + builder.DataSource + ";Database=" + database +
                ";User ID=" + AgentDbUser + ";Password=" + AgentDbPassword + ";Connect Timeout=8";
            SqlConnection fallbackConn = new SqlConnection(fallback);
            fallbackConn.Open();
            return fallbackConn;
        }

        private static AgentSelectionSnapshot CaptureAgentSelection(Form mainForm)
        {
            AgentSelectionSnapshot snapshot = new AgentSelectionSnapshot();
            try
            {
                DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
                if (grid != null)
                {
                    snapshot.QuotaKeys = GetSelectedQuotaKeys(grid);
                    foreach (DataGridViewRow row in grid.SelectedRows.Cast<DataGridViewRow>()
                        .Concat(grid.SelectedCells.Cast<DataGridViewCell>()
                            .Where(c => c.RowIndex >= 0 && c.RowIndex < grid.Rows.Count)
                            .Select(c => grid.Rows[c.RowIndex]))
                        .Distinct()
                        .Take(20))
                    {
                        string code = GetRowValue(row, "定额编号DE", "定额编号");
                        if (!String.IsNullOrEmpty(code) && !snapshot.QuotaCodes.Contains(code))
                        {
                            snapshot.QuotaCodes.Add(code);
                        }
                    }

                    foreach (DataGridViewRow row in grid.Rows)
                    {
                        if (row.IsNewRow)
                        {
                            continue;
                        }

                        string unitText = GetRowValue(row, "总概算序号de", "总概算序号");
                        long unitId;
                        if (Int64.TryParse(unitText, NumberStyles.Integer, CultureInfo.InvariantCulture, out unitId) && unitId > 0)
                        {
                            snapshot.CurrentUnitId = unitId;
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("Agent capture grid selection failed: " + ex.Message);
            }

            try
            {
                TreeView tree = GetField<TreeView>(mainForm, "Tv_tree");
                TreeNode node = tree != null ? tree.SelectedNode : GetField<TreeNode>(mainForm, "CurrNode");
                if (node != null)
                {
                    SqlConnection hostConn = GetProjectConnection(mainForm);
                    if (hostConn != null)
                    {
                        snapshot.ItemNo = ResolveChapterNo(mainForm, hostConn, node) ?? "";
                    }

                    snapshot.ItemName = node.Text ?? "";
                    if (snapshot.CurrentUnitId == 0)
                    {
                        long unitId;
                        string unitText = TryGetValue(node.Tag, "总概算序号");
                        if (Int64.TryParse(unitText ?? "", NumberStyles.Integer, CultureInfo.InvariantCulture, out unitId))
                        {
                            snapshot.CurrentUnitId = unitId;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("Agent capture tree selection failed: " + ex.Message);
            }

            // 从章节树按 Ctrl+Q 进入时意图是整个条目，丢弃定额表里"当前行"这类顺带选中，
            // 让"未明确选定额"回落到当前条目下全部定额（修复只改第一条的问题）。
            if (s_agentInvokeFromTree)
            {
                snapshot.QuotaKeys = new List<QuotaKey>();
                snapshot.QuotaCodes = new List<string>();
            }

            snapshot.UnitCodeMap = ReadAgentUnitCodeMap(mainForm);
            string curCode;
            if (snapshot.CurrentUnitId > 0 && snapshot.UnitCodeMap.TryGetValue(snapshot.CurrentUnitId, out curCode))
            {
                snapshot.CurrentUnitCode = curCode;
            }
            else
            {
                snapshot.CurrentUnitCode = ReadAgentCurrentUnitText(mainForm);   // 兜底：直接读当前下拉框文本
            }

            return snapshot;
        }

        // 从"总概算"下拉框的绑定数据里读全部单元 总概算序号->总概算编号 映射（与界面同源，最可靠）。
        // 找不到时把各下拉框的列名写进日志，便于定位。
        private static Dictionary<long, string> ReadAgentUnitCodeMap(Form mainForm)
        {
            Dictionary<long, string> map = new Dictionary<long, string>();
            try
            {
                List<string> diag = new List<string>();
                foreach (ComboBox cb in EnumerateAgentControls<ComboBox>(mainForm))
                {
                    DataTable table = GetAgentComboTable(cb);
                    if (table == null)
                    {
                        continue;
                    }

                    diag.Add("[" + String.Join(",", table.Columns.Cast<DataColumn>().Select(c => c.ColumnName).ToArray()) + "]");

                    if (!table.Columns.Contains("总概算序号") || !table.Columns.Contains("总概算编号"))
                    {
                        continue;
                    }

                    foreach (DataRow row in table.Rows)
                    {
                        long sn;
                        if (!Int64.TryParse(Convert.ToString(row["总概算序号"]), NumberStyles.Integer, CultureInfo.InvariantCulture, out sn) || sn <= 0)
                        {
                            continue;
                        }

                        string code = Convert.ToString(row["总概算编号"]).Trim();
                        int p = code.IndexOfAny(new char[] { '(', '（' });
                        if (p > 0)
                        {
                            code = code.Substring(0, p).Trim();
                        }

                        if (code.Length > 0)
                        {
                            map[sn] = code;
                        }
                    }
                }

                if (map.Count == 0)
                {
                    Log("Agent unit map: no usable 总概算 combo. comboTables=" + String.Join(" ", diag.ToArray()));
                }
            }
            catch (Exception ex)
            {
                Log("ReadAgentUnitCodeMap failed: " + ex.Message);
            }

            return map;
        }

        private static DataTable GetAgentComboTable(ComboBox cb)
        {
            object ds = cb.DataSource;
            DataView dv = ds as DataView;
            if (dv != null) { return dv.Table; }
            DataTable dt = ds as DataTable;
            if (dt != null) { return dt; }
            BindingSource bs = ds as BindingSource;
            if (bs != null)
            {
                DataView dv2 = bs.List as DataView;
                if (dv2 != null) { return dv2.Table; }
                DataTable dt2 = bs.DataSource as DataTable;
                if (dt2 != null) { return dt2; }
                DataView dv3 = bs.DataSource as DataView;
                if (dv3 != null) { return dv3.Table; }
            }
            return null;
        }

        // 兜底：直接读含 ZGS 的下拉框当前文本，取 "(" 前的编号（仅当前单元，可能不精确）。
        private static string ReadAgentCurrentUnitText(Form mainForm)
        {
            try
            {
                foreach (ComboBox cb in EnumerateAgentControls<ComboBox>(mainForm))
                {
                    string code = ExtractAgentUnitCode(cb.Text);
                    if (code.Length > 0) { return code; }
                }
            }
            catch (Exception ex)
            {
                Log("ReadAgentCurrentUnitText failed: " + ex.Message);
            }
            return "";
        }

        // 从下拉文本里抽总概算编号：必须含 "ZGS"（单元编号特征），取 "(" 前部分，如 "ZGS_03(沪昆…)" -> "ZGS_03"。
        private static string ExtractAgentUnitCode(string text)
        {
            if (String.IsNullOrEmpty(text))
            {
                return "";
            }

            string t = text.Trim();
            if (t.IndexOf("ZGS", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return "";
            }

            int paren = t.IndexOfAny(new char[] { '(', '（' });
            return (paren > 0 ? t.Substring(0, paren) : t).Trim();
        }

        private static IEnumerable<T> EnumerateAgentControls<T>(Control root) where T : Control
        {
            foreach (Control c in root.Controls)
            {
                T match = c as T;
                if (match != null)
                {
                    yield return match;
                }

                foreach (T sub in EnumerateAgentControls<T>(c))
                {
                    yield return sub;
                }
            }
        }

        // 单元显示：优先总概算编号，没有就回退到序号。unitId<=0 返回空串。
        private static string AgentUnitDisplay(Dictionary<long, string> codes, long unitId)
        {
            if (unitId <= 0)
            {
                return "";
            }

            string code;
            if (codes != null && codes.TryGetValue(unitId, out code) && !String.IsNullOrEmpty(code))
            {
                return code;
            }

            return unitId.ToString(CultureInfo.InvariantCulture);
        }

        // ===== 条目/单元解析 =====

        private static void ValidateAgentItemExists(SqlConnection conn, string itemNo)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select count(*) from 章节表 where 条目编号=@bh";
                cmd.Parameters.AddWithValue("@bh", itemNo);
                int count = Convert.ToInt32(cmd.ExecuteScalar(), CultureInfo.InvariantCulture);
                if (count == 0)
                {
                    throw new AgentPlanException("条目编号 \"" + itemNo + "\" 在本项目章节表中不存在，已中止。请核对编号（可输入条目名称让AI帮忙定位）。");
                }
            }
        }

        private static List<long> ResolveAgentUnitIds(SqlConnection conn, AgentCommand command, AgentSelectionSnapshot selection, List<string> warnings)
        {
            List<string> tokens = command.Units ?? new List<string>();
            if (tokens.Count == 1 && (tokens[0] == "所有" || tokens[0] == "全部" || tokens[0] == "*"))
            {
                warnings.Add("本次作用于所有单元。");
                return null;
            }

            if (tokens.Count == 0)
            {
                if (selection.CurrentUnitId > 0)
                {
                    string curDisplay = String.IsNullOrEmpty(selection.CurrentUnitCode)
                        ? "总概算序号 " + selection.CurrentUnitId.ToString(CultureInfo.InvariantCulture)
                        : selection.CurrentUnitCode;
                    warnings.Add("默认只作用于当前单元（" + curDisplay + "）。要跨单元请在指令里写明单元，或写\"所有单元\"。");
                    return new List<long> { selection.CurrentUnitId };
                }

                warnings.Add("未能识别当前单元，本次会匹配所有单元里的同名条目，请在预览里核对\"单元\"列！");
                return null;
            }

            List<long> ids = new List<long>();
            foreach (string token in tokens)
            {
                string t = token.Trim();
                if (t.Length == 0)
                {
                    continue;
                }

                if (t == AgentSelectedToken || t == "当前")
                {
                    if (selection.CurrentUnitId <= 0)
                    {
                        throw new AgentPlanException("指令里提到了\"当前单元\"，但没能识别当前单元，请先打开某个单元的条目。");
                    }

                    if (!ids.Contains(selection.CurrentUnitId))
                    {
                        ids.Add(selection.CurrentUnitId);
                    }

                    continue;
                }

                List<long> matched = new List<long>();
                List<string> labels = new List<string>();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    long numeric;
                    if (Int64.TryParse(t, NumberStyles.Integer, CultureInfo.InvariantCulture, out numeric))
                    {
                        cmd.CommandText = "select 总概算序号, 总概算编号, 编制范围 from 总概算信息 where 总概算序号=@v";
                        cmd.Parameters.AddWithValue("@v", numeric);
                    }
                    else
                    {
                        cmd.CommandText = "select 总概算序号, 总概算编号, 编制范围 from 总概算信息 where 总概算编号=@v or 编制范围=@v";
                        cmd.Parameters.AddWithValue("@v", t);
                    }

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            matched.Add(Convert.ToInt64(reader.GetValue(0), CultureInfo.InvariantCulture));
                            labels.Add(Convert.ToString(reader.GetValue(0)) + "/" + Convert.ToString(reader.GetValue(1)) + "/" + Convert.ToString(reader.GetValue(2)));
                        }
                    }
                }

                if (matched.Count == 0)
                {
                    throw new AgentPlanException("没有找到单元 \"" + t + "\"（可用单元名称、_ZGS_编号或总概算序号）。");
                }

                if (matched.Count > 1)
                {
                    throw new AgentPlanException("单元 \"" + t + "\" 有多个匹配：" + String.Join("；", labels.ToArray()) +
                        "。请改用 _ZGS_编号 或 总概算序号 精确指定。");
                }

                if (!ids.Contains(matched[0]))
                {
                    ids.Add(matched[0]);
                }
            }

            return ids.Count == 0 ? null : ids;
        }

        private static string BuildAgentItemCondition(SqlCommand cmd, string itemNo, bool includeChildren, int index)
        {
            string exact = "@bh" + index.ToString(CultureInfo.InvariantCulture);
            cmd.Parameters.AddWithValue(exact, itemNo);
            if (!includeChildren)
            {
                return "ZJ.条目编号=" + exact;
            }

            if (itemNo.IndexOf('.') >= 0)
            {
                string dot = "@bhdot" + index.ToString(CultureInfo.InvariantCulture);
                string dash = "@bhdash" + index.ToString(CultureInfo.InvariantCulture);
                cmd.Parameters.AddWithValue(dot, itemNo + ".%");
                cmd.Parameters.AddWithValue(dash, itemNo + "-%");
                return "(ZJ.条目编号=" + exact + " or ZJ.条目编号 like " + dot + " or ZJ.条目编号 like " + dash + ")";
            }

            string prefix = "@bhpre" + index.ToString(CultureInfo.InvariantCulture);
            cmd.Parameters.AddWithValue(prefix, itemNo + "%");
            return "(ZJ.条目编号=" + exact + " or ZJ.条目编号 like " + prefix + ")";
        }

        private static string BuildAgentQuotaFilterCondition(SqlCommand cmd, List<string> filter)
        {
            if (filter == null || filter.Count == 0)
            {
                return "";
            }

            List<string> parts = new List<string>();
            for (int i = 0; i < filter.Count; i++)
            {
                string code = filter[i].Trim();
                string p = "@q" + i.ToString(CultureInfo.InvariantCulture);
                cmd.Parameters.AddWithValue(p, code);
                cmd.Parameters.AddWithValue(p + "m", code + "*%");
                cmd.Parameters.AddWithValue(p + "d", code + "/%");
                parts.Add("DE.定额编号=" + p + " or DE.定额编号 like " + p + "m or DE.定额编号 like " + p + "d");
            }

            return " and (" + String.Join(" or ", parts.ToArray()) + ")";
        }

        private static string BuildAgentUnitCondition(List<long> unitIds)
        {
            if (unitIds == null || unitIds.Count == 0)
            {
                return "";
            }

            return " and DE.总概算序号 in (" +
                String.Join(",", unitIds.Select(id => id.ToString(CultureInfo.InvariantCulture)).ToArray()) + ")";
        }

        private static bool AgentQuotaCodeMatches(string code, List<string> filter)
        {
            if (filter == null || filter.Count == 0)
            {
                return true;
            }

            string actual = (code ?? "").Trim();
            foreach (string want in filter)
            {
                string w = want.Trim();
                if (String.Equals(actual, w, StringComparison.OrdinalIgnoreCase) ||
                    actual.StartsWith(w + "*", StringComparison.OrdinalIgnoreCase) ||
                    actual.StartsWith(w + "/", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }

            return false;
        }

        private static AgentTargetRow ReadAgentTargetRow(IDataRecord record)
        {
            AgentTargetRow row = new AgentTargetRow();
            row.QuotaSequence = Convert.ToInt64(record.GetValue(0), CultureInfo.InvariantCulture);
            row.ItemNo = record.IsDBNull(1) ? "" : Convert.ToString(record.GetValue(1)).Trim();
            row.QuotaCode = record.IsDBNull(2) ? "" : Convert.ToString(record.GetValue(2)).Trim();
            row.QuantityInput = record.IsDBNull(3) ? "" : Convert.ToString(record.GetValue(3)).Trim();
            row.Quantity = record.IsDBNull(4) ? (object)DBNull.Value : record.GetValue(4);
            row.UnitPrice = record.IsDBNull(5) ? (object)DBNull.Value : record.GetValue(5);
            row.UnitId = record.IsDBNull(6) ? 0 : Convert.ToInt64(record.GetValue(6), CultureInfo.InvariantCulture);
            row.AdjustText = record.IsDBNull(7) ? "" : Convert.ToString(record.GetValue(7)).Trim();
            return row;
        }

        private const string AgentTargetRowSelect =
            "select DE.定额序号, ZJ.条目编号, DE.定额编号, DE.工程数量输入, DE.工程数量, DE.单价, DE.总概算序号, cast(DE.定额调整 as nvarchar(max)) " +
            "from 定额输入 DE inner join 章节表 ZJ on DE.条目序号=ZJ.条目序号 ";

        private static List<AgentTargetRow> ResolveAgentTargetRows(SqlConnection conn, AgentSelectionSnapshot selection,
            List<string> items, bool includeChildren, List<string> quotaFilter, List<long> unitIds)
        {
            Dictionary<long, AgentTargetRow> rows = new Dictionary<long, AgentTargetRow>();
            foreach (string rawItem in items)
            {
                string itemNo = rawItem;
                if (itemNo == AgentCurrentItemToken)
                {
                    // 给了定额编号但没给条目编号：作用于当前条目下全部该定额（忽略行选中）。
                    if (String.IsNullOrEmpty(selection.ItemNo))
                    {
                        throw new AgentPlanException("指令里用到了\"当前条目\"，但当前没有选中条目，请先在左侧树点中一个条目。");
                    }

                    itemNo = selection.ItemNo;
                }
                else if (itemNo == AgentSelectedToken)
                {
                    if (selection.QuotaKeys.Count > 0)
                    {
                        foreach (AgentTargetRow row in ResolveAgentRowsByKeys(conn, selection.QuotaKeys, quotaFilter))
                        {
                            rows[row.QuotaSequence] = row;
                        }

                        continue;
                    }

                    if (String.IsNullOrEmpty(selection.ItemNo))
                    {
                        throw new AgentPlanException("指令里提到了\"选中的\"，但当前没有选中任何定额行或条目，请先在软件里选中再试。");
                    }

                    itemNo = selection.ItemNo;
                }

                ValidateAgentItemExists(conn, itemNo);
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    string condition = BuildAgentItemCondition(cmd, itemNo, includeChildren, 0);
                    string filter = BuildAgentQuotaFilterCondition(cmd, quotaFilter);
                    string unitCondition = BuildAgentUnitCondition(unitIds);
                    cmd.CommandText = AgentTargetRowSelect + "where " + condition + filter + unitCondition +
                        " order by DE.总概算序号, ZJ.条目编号, DE.顺号";
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            AgentTargetRow row = ReadAgentTargetRow(reader);
                            rows[row.QuotaSequence] = row;
                        }
                    }
                }
            }

            return rows.Values.ToList();
        }

        private static List<AgentTargetRow> ResolveAgentRowsByKeys(SqlConnection conn, List<QuotaKey> keys, List<string> quotaFilter)
        {
            List<AgentTargetRow> rows = new List<AgentTargetRow>();
            foreach (QuotaKey key in keys)
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = AgentTargetRowSelect + "where DE.总概算序号=@zgs and DE.条目序号=@tm and DE.顺号=@xh";
                    cmd.Parameters.AddWithValue("@zgs", key.TotalNo);
                    cmd.Parameters.AddWithValue("@tm", key.ChapterSeq);
                    cmd.Parameters.AddWithValue("@xh", key.OrderNo);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            AgentTargetRow row = ReadAgentTargetRow(reader);
                            if (AgentQuotaCodeMatches(row.QuotaCode, quotaFilter))
                            {
                                rows.Add(row);
                            }
                        }
                    }
                }
            }

            return rows;
        }

        // 解析条目编号 -> 条目序号集合（含子树），用于总概算条目等按 条目序号 定位的表。
        private static List<long> ResolveAgentChapterSeqs(SqlConnection conn, AgentSelectionSnapshot selection, List<string> items, bool includeChildren)
        {
            HashSet<long> seqs = new HashSet<long>();
            foreach (string rawItem in items)
            {
                string itemNo = (rawItem == AgentSelectedToken || rawItem == AgentCurrentItemToken) ? selection.ItemNo : rawItem;
                if (String.IsNullOrEmpty(itemNo))
                {
                    throw new AgentPlanException("指令里提到了\"选中的\"条目，但当前没有选中条目。");
                }

                ValidateAgentItemExists(conn, itemNo);
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    string condition = BuildAgentItemCondition(cmd, itemNo, includeChildren, 0);
                    cmd.CommandText = "select 条目序号 from 章节表 ZJ where " + condition;
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            seqs.Add(Convert.ToInt64(reader.GetValue(0), CultureInfo.InvariantCulture));
                        }
                    }
                }
            }

            return seqs.ToList();
        }

        // ===== 字段更新构造（针对定额输入行） =====

        private static AgentFieldUpdate BuildQuotaFieldUpdate(AgentTargetRow row, string action, Dictionary<string, object> newValues, Dictionary<string, object> oldValues, string oldDisplay, string newDisplay)
        {
            AgentFieldUpdate update = new AgentFieldUpdate();
            update.Table = "定额输入";
            update.KeyClause = "定额序号=@k0";
            update.KeyValues["@k0"] = row.QuotaSequence;
            update.QuotaSequence = row.QuotaSequence;
            update.NewValues = newValues;
            update.OldValues = oldValues;
            update.Action = action;
            update.UnitId = row.UnitId;
            update.ItemNo = row.ItemNo;
            update.QuotaCode = row.QuotaCode;
            update.OldDisplay = oldDisplay;
            update.NewDisplay = newDisplay;
            return update;
        }

        private static double AgentToDouble(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return 0;
            }

            return Convert.ToDouble(value, CultureInfo.InvariantCulture);
        }

        // ===== 计划构建（只读） =====

        private static AgentPlan BuildAgentPlan(SqlConnection conn, AgentSelectionSnapshot selection, List<AgentCommand> commands)
        {
            AgentPlan plan = new AgentPlan();
            plan.Commands = commands;

            foreach (AgentCommand command in commands)
            {
                List<long> unitIds = (command.Type == "create_unit" || command.Type == "set_material_scheme")
                    ? null
                    : ResolveAgentUnitIds(conn, command, selection, plan.Warnings);

                switch (command.Type)
                {
                    case "multiply_quantity":
                        BuildMultiplyPlan(conn, selection, command, unitIds, plan);
                        break;
                    case "set_quantity":
                        BuildSetQuantityPlan(conn, selection, command, unitIds, plan);
                        break;
                    case "remove_text":
                        BuildRemoveTextPlan(conn, selection, command, unitIds, plan);
                        break;
                    case "set_adjustment":
                        BuildSetAdjustmentPlan(conn, selection, command, unitIds, plan);
                        break;
                    case "clear_quantity":
                        BuildClearQuantityPlan(conn, selection, command, unitIds, plan);
                        break;
                    case "replace_quota_code":
                        BuildReplaceCodePlan(conn, selection, command, unitIds, plan);
                        break;
                    case "delete_quotas":
                        BuildDeletePlan(conn, selection, command, unitIds, plan);
                        break;
                    case "copy_quotas":
                        BuildCopyPlan(conn, selection, command, unitIds, plan);
                        break;
                    case "insert_quotas":
                        BuildInsertPlan(conn, selection, command, plan);
                        break;
                    case "set_transport_scheme":
                        BuildTransportSchemePlan(conn, selection, command, unitIds, plan);
                        break;
                    case "set_material_scheme":
                        BuildMaterialSchemePlan(conn, selection, command, plan);
                        break;
                    case "create_unit":
                        plan.UnitCopies.Add(BuildAgentUnitCopySpec(conn, selection, command, plan));
                        break;
                    default:
                        throw new AgentPlanException("不支持的命令类型：" + command.Type);
                }
            }

            // 全部单元 序号->编号，预览/汇总按编号显示（覆盖跨单元、未识别当前单元等情况）。
            if (selection.UnitCodeMap != null && selection.UnitCodeMap.Count > 0)
            {
                plan.UnitCodes = selection.UnitCodeMap;
            }
            if (selection.CurrentUnitId > 0 && !String.IsNullOrEmpty(selection.CurrentUnitCode) && !plan.UnitCodes.ContainsKey(selection.CurrentUnitId))
            {
                plan.UnitCodes[selection.CurrentUnitId] = selection.CurrentUnitCode;
            }

            FinalizeAgentPlan(plan);
            return plan;
        }

        private static void BuildMultiplyPlan(SqlConnection conn, AgentSelectionSnapshot selection, AgentCommand command, List<long> unitIds, AgentPlan plan)
        {
            List<AgentTargetRow> targets = ResolveAgentTargetRows(conn, selection, command.Items, command.IncludeChildren, command.QuotaFilter, unitIds);
            int skipped = 0;
            foreach (AgentTargetRow row in targets)
            {
                if (command.Target == "quota_code")
                {
                    if (String.IsNullOrEmpty(row.QuotaCode))
                    {
                        skipped++;
                        continue;
                    }

                    string newCode = row.QuotaCode + command.Operator + command.Factor;
                    plan.FieldUpdates.Add(BuildQuotaFieldUpdate(row, "乘定额编号",
                        new Dictionary<string, object> { { "定额编号", newCode } },
                        new Dictionary<string, object> { { "定额编号", row.QuotaCode } },
                        row.QuotaCode, newCode));
                    plan.NeedsRecalc = true;
                }
                else if (command.Target == "unit_price")
                {
                    double oldPrice = AgentToDouble(row.UnitPrice);
                    if (oldPrice == 0)
                    {
                        skipped++;
                        continue;
                    }

                    double factor = Convert.ToDouble(command.Factor, CultureInfo.InvariantCulture);
                    double newPrice = command.Operator == "/" ? oldPrice / factor : oldPrice * factor;
                    newPrice = Math.Round(newPrice, 2, MidpointRounding.AwayFromZero);   // 单价保留两位小数
                    double newTotal = Math.Round(newPrice * AgentToDouble(row.Quantity), 2, MidpointRounding.AwayFromZero);
                    plan.FieldUpdates.Add(BuildQuotaFieldUpdate(row, "乘单价",
                        new Dictionary<string, object> { { "单价", newPrice }, { "合价", newTotal } },
                        new Dictionary<string, object> { { "单价", row.UnitPrice }, { "合价", DBNull.Value } },
                        oldPrice.ToString("0.00", CultureInfo.InvariantCulture), newPrice.ToString("0.00", CultureInfo.InvariantCulture)));
                }
                else
                {
                    if (String.IsNullOrEmpty(row.QuantityInput))
                    {
                        skipped++;
                        continue;
                    }

                    string newExpr = "(" + row.QuantityInput + ")" + command.Operator + command.Factor;
                    object newQuantity;
                    try
                    {
                        newQuantity = Convert.ToDouble(EvaluateDecimal(newExpr), CultureInfo.InvariantCulture);
                    }
                    catch (Exception)
                    {
                        skipped++;
                        continue;
                    }

                    plan.FieldUpdates.Add(BuildQuotaFieldUpdate(row, "乘数量",
                        new Dictionary<string, object> { { "工程数量输入", newExpr }, { "工程数量", newQuantity } },
                        new Dictionary<string, object> { { "工程数量输入", row.QuantityInput }, { "工程数量", row.Quantity } },
                        row.QuantityInput, newExpr));
                }
            }

            if (command.Target == "unit_price")
            {
                plan.Warnings.Add("单价直接改单价列，保留两位小数：仅对补充定额(SH/SQ/ZLF/LF/SF/TLF 等手填单价)长期有效；普通定额单价为计算值(常为0会被跳过)，会被主程序重算覆盖，建议改用\"定额编号 ×系数\"。");
            }

            if (skipped > 0)
            {
                plan.Warnings.Add("乘系数：有 " + skipped.ToString(CultureInfo.InvariantCulture) + " 行目标为空或为0/无法计算，已跳过。");
            }
        }

        private static void BuildSetQuantityPlan(SqlConnection conn, AgentSelectionSnapshot selection, AgentCommand command, List<long> unitIds, AgentPlan plan)
        {
            object newQuantity;
            try
            {
                newQuantity = Convert.ToDouble(EvaluateDecimal(command.Value), CultureInfo.InvariantCulture);
            }
            catch (Exception)
            {
                throw new AgentPlanException("设数量的数值无法计算：" + command.Value);
            }

            List<AgentTargetRow> targets = ResolveAgentTargetRows(conn, selection, command.Items, command.IncludeChildren, command.QuotaFilter, unitIds);
            foreach (AgentTargetRow row in targets)
            {
                plan.FieldUpdates.Add(BuildQuotaFieldUpdate(row, "设数量",
                    new Dictionary<string, object> { { "工程数量输入", command.Value }, { "工程数量", newQuantity } },
                    new Dictionary<string, object> { { "工程数量输入", row.QuantityInput }, { "工程数量", row.Quantity } },
                    row.QuantityInput, command.Value));
            }
        }

        // 设置/追加定额调整字段（如 /XG1、/1294861,,1）。mode=append 追加到现有串后，否则替换整串。
        private static void BuildSetAdjustmentPlan(SqlConnection conn, AgentSelectionSnapshot selection, AgentCommand command, List<long> unitIds, AgentPlan plan)
        {
            bool append = command.Mode == "append";
            string value = command.Value ?? "";
            List<AgentTargetRow> targets = ResolveAgentTargetRows(conn, selection, command.Items, command.IncludeChildren, command.QuotaFilter, unitIds);
            int skipped = 0;
            foreach (AgentTargetRow row in targets)
            {
                string oldAdj = row.AdjustText ?? "";
                string newAdj;
                if (append)
                {
                    if (oldAdj.IndexOf(value, StringComparison.Ordinal) >= 0)
                    {
                        skipped++;   // 已含该调整，避免重复追加
                        continue;
                    }

                    newAdj = oldAdj + value;
                }
                else
                {
                    newAdj = value;
                    if (oldAdj == newAdj)
                    {
                        skipped++;
                        continue;
                    }
                }

                plan.FieldUpdates.Add(BuildQuotaFieldUpdate(row, append ? "定额调整" : "设调整",
                    new Dictionary<string, object> { { "定额调整", (object)newAdj } },
                    new Dictionary<string, object> { { "定额调整", (object)oldAdj } },
                    String.IsNullOrEmpty(oldAdj) ? "(空)" : oldAdj, newAdj));
            }

            plan.NeedsRecalc = true;
            if (skipped > 0)
            {
                plan.Warnings.Add("有 " + skipped.ToString(CultureInfo.InvariantCulture) + " 行已含该调整或与目标相同，已跳过。");
            }
        }

        // 从 工程数量输入 / 定额编号 / 定额调整 里去掉指定子串（如 /100、*9、/XG1），逐行字面替换。
        private static void BuildRemoveTextPlan(SqlConnection conn, AgentSelectionSnapshot selection, AgentCommand command, List<long> unitIds, AgentPlan plan)
        {
            string fragment = command.RemoveText;
            List<AgentTargetRow> targets = ResolveAgentTargetRows(conn, selection, command.Items, command.IncludeChildren, command.QuotaFilter, unitIds);
            int skipped = 0;
            int emptyCode = 0;
            foreach (AgentTargetRow row in targets)
            {
                if (command.Target == "adjustment")
                {
                    string oldAdj = row.AdjustText;
                    if (String.IsNullOrEmpty(oldAdj) || oldAdj.IndexOf(fragment, StringComparison.Ordinal) < 0)
                    {
                        skipped++;
                        continue;
                    }

                    string newAdj = oldAdj.Replace(fragment, "");
                    plan.FieldUpdates.Add(BuildQuotaFieldUpdate(row, "去掉调整",
                        new Dictionary<string, object> { { "定额调整", (object)newAdj } },
                        new Dictionary<string, object> { { "定额调整", (object)oldAdj } },
                        oldAdj, String.IsNullOrEmpty(newAdj) ? "(空)" : newAdj));
                    plan.NeedsRecalc = true;
                }
                else if (command.Target == "quota_code")
                {
                    string oldCode = row.QuotaCode;
                    if (String.IsNullOrEmpty(oldCode) || oldCode.IndexOf(fragment, StringComparison.Ordinal) < 0)
                    {
                        skipped++;
                        continue;
                    }

                    string newCode = oldCode.Replace(fragment, "");
                    if (String.IsNullOrEmpty(newCode))
                    {
                        emptyCode++;
                        continue;
                    }

                    plan.FieldUpdates.Add(BuildQuotaFieldUpdate(row, "去掉编号",
                        new Dictionary<string, object> { { "定额编号", newCode } },
                        new Dictionary<string, object> { { "定额编号", oldCode } },
                        oldCode, newCode));
                    plan.NeedsRecalc = true;
                }
                else
                {
                    string oldExpr = row.QuantityInput;
                    if (String.IsNullOrEmpty(oldExpr) || oldExpr.IndexOf(fragment, StringComparison.Ordinal) < 0)
                    {
                        skipped++;
                        continue;
                    }

                    string newExpr = oldExpr.Replace(fragment, "");
                    object newQuantity;
                    if (String.IsNullOrEmpty(newExpr))
                    {
                        newQuantity = DBNull.Value;
                    }
                    else
                    {
                        try
                        {
                            newQuantity = Convert.ToDouble(EvaluateDecimal(newExpr), CultureInfo.InvariantCulture);
                        }
                        catch (Exception)
                        {
                            skipped++;
                            continue;
                        }
                    }

                    plan.FieldUpdates.Add(BuildQuotaFieldUpdate(row, "删除数量",
                        new Dictionary<string, object> { { "工程数量输入", newExpr }, { "工程数量", newQuantity } },
                        new Dictionary<string, object> { { "工程数量输入", oldExpr }, { "工程数量", row.Quantity } },
                        oldExpr, newExpr));
                }
            }

            if (emptyCode > 0)
            {
                plan.Warnings.Add("有 " + emptyCode.ToString(CultureInfo.InvariantCulture) + " 行去掉后定额编号会变空，已跳过。");
            }

            if (skipped > 0)
            {
                plan.Warnings.Add("有 " + skipped.ToString(CultureInfo.InvariantCulture) + " 行不含\"" + fragment + "\"或去掉后无法计算，已跳过。");
            }
        }

        private static void BuildClearQuantityPlan(SqlConnection conn, AgentSelectionSnapshot selection, AgentCommand command, List<long> unitIds, AgentPlan plan)
        {
            List<AgentTargetRow> targets = ResolveAgentTargetRows(conn, selection, command.Items, command.IncludeChildren, command.QuotaFilter, unitIds);
            foreach (AgentTargetRow row in targets)
            {
                if (String.IsNullOrEmpty(row.QuantityInput) && row.Quantity == DBNull.Value)
                {
                    continue;
                }

                plan.FieldUpdates.Add(BuildQuotaFieldUpdate(row, "清空数量",
                    new Dictionary<string, object> { { "工程数量输入", "" }, { "工程数量", DBNull.Value } },
                    new Dictionary<string, object> { { "工程数量输入", row.QuantityInput }, { "工程数量", row.Quantity } },
                    row.QuantityInput, "(清空)"));
            }
        }

        private static void BuildReplaceCodePlan(SqlConnection conn, AgentSelectionSnapshot selection, AgentCommand command, List<long> unitIds, AgentPlan plan)
        {
            // 用 from_code 作为过滤匹配（含 *系数 后缀的同号），逐行把 base 替换成 to_code，保留系数后缀。
            List<string> filter = new List<string> { command.FromCode };
            List<AgentTargetRow> targets = ResolveAgentTargetRows(conn, selection, command.Items, command.IncludeChildren, filter, unitIds);
            foreach (AgentTargetRow row in targets)
            {
                string oldCode = row.QuotaCode;
                string suffix = "";
                if (oldCode.Length > command.FromCode.Length &&
                    oldCode.StartsWith(command.FromCode, StringComparison.OrdinalIgnoreCase))
                {
                    char next = oldCode[command.FromCode.Length];
                    if (next == '*' || next == '/')
                    {
                        suffix = oldCode.Substring(command.FromCode.Length);
                    }
                    else
                    {
                        continue; // 不是同号，跳过（前缀但非系数分隔）
                    }
                }
                else if (!String.Equals(oldCode, command.FromCode, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string newCode = command.ToCode + suffix;
                plan.FieldUpdates.Add(BuildQuotaFieldUpdate(row, "替换定额",
                    new Dictionary<string, object> { { "定额编号", newCode } },
                    new Dictionary<string, object> { { "定额编号", oldCode } },
                    oldCode, newCode));
            }

            plan.NeedsRecalc = true;
        }

        private static void BuildDeletePlan(SqlConnection conn, AgentSelectionSnapshot selection, AgentCommand command, List<long> unitIds, AgentPlan plan)
        {
            List<AgentTargetRow> targets = ResolveAgentTargetRows(conn, selection, command.Items, command.IncludeChildren, command.QuotaFilter, unitIds);
            foreach (AgentTargetRow row in targets)
            {
                plan.Deletes.Add(new AgentDeleteRow
                {
                    QuotaSequence = row.QuotaSequence,
                    UnitId = row.UnitId,
                    ItemNo = row.ItemNo,
                    QuotaCode = row.QuotaCode
                });
            }
        }

        private static void BuildCopyPlan(SqlConnection conn, AgentSelectionSnapshot selection, AgentCommand command, List<long> unitIds, AgentPlan plan)
        {
            List<AgentTargetRow> source = ResolveAgentTargetRows(conn, selection, new List<string> { command.SourceItem }, false, command.QuotaFilter, unitIds);
            if (source.Count == 0)
            {
                throw new AgentPlanException("来源条目 " + command.SourceItem + " 下没有可复制的定额。");
            }

            foreach (string target in command.TargetItems)
            {
                ValidateAgentItemExists(conn, target);
                AgentInsertGroup group = new AgentInsertGroup();
                group.ItemNo = target;
                foreach (AgentTargetRow src in source)
                {
                    AgentQuotaInput quota = new AgentQuotaInput();
                    quota.Code = src.QuotaCode;
                    quota.Quantity = !String.IsNullOrEmpty(src.QuantityInput)
                        ? src.QuantityInput
                        : (src.Quantity == DBNull.Value ? "" : Convert.ToString(src.Quantity, CultureInfo.InvariantCulture));
                    group.Quotas.Add(quota);
                }

                plan.Inserts.Add(group);
            }
        }

        private static void BuildInsertPlan(SqlConnection conn, AgentSelectionSnapshot selection, AgentCommand command, AgentPlan plan)
        {
            foreach (string rawItem in command.Items)
            {
                string itemNo = (rawItem == AgentSelectedToken || rawItem == AgentCurrentItemToken) ? selection.ItemNo : rawItem;
                if (String.IsNullOrEmpty(itemNo))
                {
                    throw new AgentPlanException("指令里提到了\"当前条目\"，但当前没有选中条目。");
                }

                ValidateAgentItemExists(conn, itemNo);
                AgentInsertGroup group = new AgentInsertGroup();
                group.ItemNo = itemNo;
                group.Quotas = command.Quotas;
                plan.Inserts.Add(group);
            }
        }

        private static void BuildTransportSchemePlan(SqlConnection conn, AgentSelectionSnapshot selection, AgentCommand command, List<long> unitIds, AgentPlan plan)
        {
            if (unitIds == null)
            {
                throw new AgentPlanException("设运输方案必须指定单元（不能对所有单元一次设，请逐个单元指定）。");
            }

            // 校验方案号是否已有定义（材料运输方案里有该方案序号的行）。
            int schemeDefs = 0;
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select count(*) from 材料运输方案 where 方案序号=@n";
                cmd.Parameters.AddWithValue("@n", command.Scheme);
                schemeDefs = Convert.ToInt32(cmd.ExecuteScalar(), CultureInfo.InvariantCulture);
            }

            if (schemeDefs == 0)
            {
                plan.Warnings.Add("运输方案号 " + command.Scheme + " 在本项目还没有材料运输方案定义。本工具只设方案号、不会现场生成方案定义，" +
                    "建议先在软件里把该方案建好（或改用已有方案号），否则运杂费可能算不出来。");
            }

            bool setParam = !String.IsNullOrEmpty(command.TransportParam);
            List<long> seqs = ResolveAgentChapterSeqs(conn, selection, command.Items, command.IncludeChildren);
            if (seqs.Count == 0)
            {
                return;
            }

            string seqIn = String.Join(",", seqs.Select(s => s.ToString(CultureInfo.InvariantCulture)).ToArray());
            string unitIn = String.Join(",", unitIds.Select(u => u.ToString(CultureInfo.InvariantCulture)).ToArray());
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select TGT.总概算序号, TGT.条目序号, TGT.运输方案, ZJ.条目编号, cast(TGT.参数调整 as nvarchar(200)) " +
                    "from 总概算条目 TGT inner join 章节表 ZJ on TGT.条目序号=ZJ.条目序号 " +
                    "where TGT.总概算序号 in (" + unitIn + ") and TGT.条目序号 in (" + seqIn + ")";
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        long unitId = Convert.ToInt64(reader.GetValue(0), CultureInfo.InvariantCulture);
                        long seq = Convert.ToInt64(reader.GetValue(1), CultureInfo.InvariantCulture);
                        string oldScheme = reader.IsDBNull(2) ? "" : Convert.ToString(reader.GetValue(2)).Trim();
                        string itemNo = reader.IsDBNull(3) ? "" : Convert.ToString(reader.GetValue(3)).Trim();
                        string oldParam = reader.IsDBNull(4) ? "" : Convert.ToString(reader.GetValue(4)).Trim();
                        bool schemeSame = oldScheme == command.Scheme;
                        bool paramSame = !setParam || oldParam == command.TransportParam;
                        if (schemeSame && paramSame)
                        {
                            continue;
                        }

                        AgentFieldUpdate update = new AgentFieldUpdate();
                        update.Table = "总概算条目";
                        update.KeyClause = "总概算序号=@k0 and 条目序号=@k1";
                        update.KeyValues["@k0"] = unitId;
                        update.KeyValues["@k1"] = seq;
                        update.NewValues["运输方案"] = command.Scheme;
                        update.OldValues["运输方案"] = oldScheme;
                        if (setParam)
                        {
                            update.NewValues["参数调整"] = command.TransportParam;
                            update.OldValues["参数调整"] = oldParam;
                        }

                        update.Action = "设运输方案";
                        update.UnitId = unitId;
                        update.ItemNo = itemNo;
                        update.OldDisplay = oldScheme + (setParam ? "/" + oldParam : "");
                        update.NewDisplay = command.Scheme + (setParam ? "/" + command.TransportParam : "");
                        plan.FieldUpdates.Add(update);
                    }
                }
            }

            plan.NeedsRecalc = true;
            plan.Warnings.Add("运输方案改的是方案号" + (setParam ? "和运输参数" : "") + "；运杂费需要在软件里手工触发\"重算\"才会更新（手工设方案也一样不会立即重算）。");
        }

        private static readonly Dictionary<string, string> AgentMaterialSchemeColumns = new Dictionary<string, string>(StringComparer.Ordinal)
        {
            { "材料", "材料费方案" }, { "机械", "机械费方案" }, { "设备", "设备费方案" }, { "工费", "工费方案" }
        };

        private static void BuildMaterialSchemePlan(SqlConnection conn, AgentSelectionSnapshot selection, AgentCommand command, AgentPlan plan)
        {
            List<long> unitIds = ResolveAgentUnitIds(conn, command, selection, plan.Warnings);
            if (unitIds == null)
            {
                throw new AgentPlanException("改材料价方案必须明确指定单元。");
            }

            string column;
            if (!AgentMaterialSchemeColumns.TryGetValue(command.SchemeKind, out column))
            {
                column = "材料费方案";
            }

            string unitIn = String.Join(",", unitIds.Select(u => u.ToString(CultureInfo.InvariantCulture)).ToArray());
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 总概算序号, 总概算编号, [" + column + "] from 总概算信息 where 总概算序号 in (" + unitIn + ")";
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        long unitId = Convert.ToInt64(reader.GetValue(0), CultureInfo.InvariantCulture);
                        string zgs = Convert.ToString(reader.GetValue(1)).Trim();
                        string oldName = reader.IsDBNull(2) ? "" : Convert.ToString(reader.GetValue(2)).Trim();
                        if (oldName == command.SchemeName)
                        {
                            continue;
                        }

                        AgentFieldUpdate update = new AgentFieldUpdate();
                        update.Table = "总概算信息";
                        update.KeyClause = "总概算序号=@k0";
                        update.KeyValues["@k0"] = unitId;
                        update.NewValues[column] = command.SchemeName;
                        update.OldValues[column] = oldName;
                        update.Action = "改" + command.SchemeKind + "费方案";
                        update.UnitId = unitId;
                        update.ItemNo = zgs;
                        update.OldDisplay = oldName;
                        update.NewDisplay = command.SchemeName;
                        plan.FieldUpdates.Add(update);
                    }
                }
            }

            plan.NeedsRecalc = true;
            plan.Warnings.Add("改" + command.SchemeKind + "费方案后，需要在软件里手工触发\"重算\"，相关费用才会按新方案更新。");
        }

        private static void FinalizeAgentPlan(AgentPlan plan)
        {
            // 同一行同时被字段更新和删除时，删除优先。
            HashSet<long> deleteIds = new HashSet<long>(plan.Deletes.Select(d => d.QuotaSequence));
            int overlap = plan.FieldUpdates.RemoveAll(u => u.Table == "定额输入" && deleteIds.Contains(u.QuotaSequence));
            if (overlap > 0)
            {
                plan.Warnings.Add("有 " + overlap.ToString(CultureInfo.InvariantCulture) + " 行同时命中修改和删除，按删除处理。");
            }

            // 预览展示行
            foreach (AgentFieldUpdate update in plan.FieldUpdates)
            {
                plan.PreviewRows.Add(new AgentPlanRow
                {
                    Action = update.Action,
                    UnitId = update.UnitId,
                    ItemNo = update.ItemNo,
                    QuotaCode = update.QuotaCode,
                    OldValue = update.OldDisplay,
                    NewValue = update.NewDisplay
                });
            }

            foreach (AgentDeleteRow del in plan.Deletes)
            {
                plan.PreviewRows.Add(new AgentPlanRow
                {
                    Action = "删除",
                    UnitId = del.UnitId,
                    ItemNo = del.ItemNo,
                    QuotaCode = del.QuotaCode,
                    OldValue = "(整行删除)",
                    NewValue = ""
                });
            }

            foreach (AgentInsertGroup group in plan.Inserts)
            {
                foreach (AgentQuotaInput quota in group.Quotas)
                {
                    plan.PreviewRows.Add(new AgentPlanRow
                    {
                        Action = "插入",
                        ItemNo = group.ItemNo,
                        QuotaCode = quota.Code,
                        OldValue = "(新增)",
                        NewValue = quota.Quantity
                    });
                }
            }

            foreach (AgentUnitCopySpec copy in plan.UnitCopies)
            {
                plan.PreviewRows.Add(new AgentPlanRow
                {
                    Action = "新建单元",
                    UnitId = copy.SourceUnitId,
                    ItemNo = copy.NewZgsNo,
                    OldValue = "(复制自 " + copy.SourceZgsNo + " " + copy.SourceName + ")",
                    NewValue = copy.NewName
                });
            }

            List<string> parts = new List<string>();
            int updates = plan.FieldUpdates.Count;
            if (updates > 0) { parts.Add("修改 " + updates.ToString(CultureInfo.InvariantCulture) + " 处"); }
            if (plan.Deletes.Count > 0) { parts.Add("删除 " + plan.Deletes.Count.ToString(CultureInfo.InvariantCulture) + " 行"); }
            int insertCount = plan.Inserts.Sum(g => g.Quotas.Count);
            if (insertCount > 0) { parts.Add("插入 " + insertCount.ToString(CultureInfo.InvariantCulture) + " 条"); }
            foreach (AgentUnitCopySpec copy in plan.UnitCopies)
            {
                parts.Add("新建单元\"" + copy.NewName + "\"");
            }

            List<long> unitList = plan.PreviewRows.Where(r => r.UnitId > 0).Select(r => r.UnitId).Distinct().ToList();
            string unitSummary = unitList.Count == 0 ? "" : "；涉及单元 " + String.Join(",", unitList.Select(u => AgentUnitDisplay(plan.UnitCodes, u)).ToArray());
            plan.Summary = (parts.Count == 0 ? "无可执行项" : String.Join("，", parts.ToArray())) + unitSummary;
        }

        private static AgentUnitCopySpec BuildAgentUnitCopySpec(SqlConnection conn, AgentSelectionSnapshot selection, AgentCommand command, AgentPlan plan)
        {
            AgentCommand probe = new AgentCommand();
            probe.Units = new List<string> { command.SourceItem };
            List<long> sourceIds = ResolveAgentUnitIds(conn, probe, selection, new List<string>());
            if (sourceIds == null || sourceIds.Count != 1)
            {
                throw new AgentPlanException("新建单元需要一个明确的源单元：" + command.SourceItem);
            }

            AgentUnitCopySpec spec = new AgentUnitCopySpec();
            spec.SourceUnitId = sourceIds[0];
            spec.NewName = command.NewName;
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 总概算编号, 编制范围 from 总概算信息 where 总概算序号=@id";
                cmd.Parameters.AddWithValue("@id", spec.SourceUnitId);
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                    {
                        throw new AgentPlanException("源单元不存在：" + command.SourceItem);
                    }

                    spec.SourceZgsNo = Convert.ToString(reader.GetValue(0)).Trim();
                    spec.SourceName = Convert.ToString(reader.GetValue(1)).Trim();
                }
            }

            int maxOrdinal = 0;
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 总概算编号 from 总概算信息";
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string no = Convert.ToString(reader.GetValue(0)) ?? "";
                        if (no.StartsWith("_ZGS_", StringComparison.OrdinalIgnoreCase))
                        {
                            int ordinal;
                            if (Int32.TryParse(no.Substring(5), NumberStyles.Integer, CultureInfo.InvariantCulture, out ordinal) && ordinal > maxOrdinal)
                            {
                                maxOrdinal = ordinal;
                            }
                        }
                    }
                }
            }

            spec.NewZgsNo = "_ZGS_" + (maxOrdinal + 1).ToString("00", CultureInfo.InvariantCulture);
            return spec;
        }

        // ===== 执行 =====

        private static string ExecuteAgentPlan(Form mainForm, AgentPlan plan, Action<string> progress)
        {
            using (SqlConnection conn = AgentCreateWorkConnection(mainForm))
            {
                if (plan.IsUndo)
                {
                    return ExecuteAgentUndo(mainForm, conn, plan.UndoRecord);
                }

                if (plan.IsRedo)
                {
                    return ExecuteAgentRedo(mainForm, conn, plan.UndoRecord);
                }

                // 新操作会让"重做"失效（标准 undo/redo 语义）。
                GetAgentRedoStack(mainForm).Clear();

                // 复核定额输入目标行仍存在。
                List<long> ids = plan.FieldUpdates.Where(u => u.Table == "定额输入").Select(u => u.QuotaSequence)
                    .Concat(plan.Deletes.Select(d => d.QuotaSequence)).Distinct().ToList();
                if (ids.Count > 0 && CountAgentExistingRows(conn, ids) != ids.Count)
                {
                    throw new AgentPlanException("从预览到现在数据已发生变化，请重新发起指令并预览。");
                }

                AgentUndoRecord undo = new AgentUndoRecord();
                undo.Summary = plan.Summary;
                undo.Time = DateTime.Now;

                foreach (AgentDeleteRow del in plan.Deletes)
                {
                    undo.Rows.Add(new AgentUndoRow { Kind = "D", QuotaSequence = del.QuotaSequence, FullRow = LoadAgentFullRow(conn, del.QuotaSequence) });
                }

                int updated = 0;
                int deleted = 0;
                using (SqlTransaction transaction = conn.BeginTransaction())
                {
                    try
                    {
                        foreach (AgentFieldUpdate update in plan.FieldUpdates)
                        {
                            undo.Rows.Add(new AgentUndoRow
                            {
                                Kind = "F",
                                QuotaSequence = update.QuotaSequence,
                                Table = update.Table,
                                KeyClause = update.KeyClause,
                                KeyValues = new Dictionary<string, object>(update.KeyValues),
                                OldValues = new Dictionary<string, object>(update.OldValues),
                                NewValues = new Dictionary<string, object>(update.NewValues)
                            });
                            updated += ExecuteAgentFieldUpdate(conn, transaction, update.Table, update.KeyClause, update.KeyValues, update.NewValues);
                        }

                        foreach (AgentDeleteRow del in plan.Deletes)
                        {
                            using (SqlCommand cmd = conn.CreateCommand())
                            {
                                cmd.Transaction = transaction;
                                cmd.CommandText = "delete from 定额输入 where 定额序号=@id";
                                cmd.Parameters.AddWithValue("@id", del.QuotaSequence);
                                deleted += cmd.ExecuteNonQuery();
                            }
                        }

                        if (plan.Deletes.Count > 0)
                        {
                            DeleteAgentDependentRows(conn, transaction, plan.Deletes.Select(d => d.QuotaSequence).ToList());
                        }

                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }

                StringBuilder message = new StringBuilder();
                if (updated > 0 || deleted > 0)
                {
                    message.Append("已修改 ").Append(updated.ToString(CultureInfo.InvariantCulture))
                        .Append(" 处，删除 ").Append(deleted.ToString(CultureInfo.InvariantCulture)).Append(" 行。");
                }

                // 同步主程序内存表（仅定额输入字段更新/删除）。
                SyncAgentHostQuotaTable(mainForm,
                    plan.FieldUpdates.Where(u => u.Table == "定额输入").ToList(),
                    plan.Deletes.Select(d => d.QuotaSequence).ToList());

                foreach (AgentUnitCopySpec copy in plan.UnitCopies)
                {
                    if (progress != null)
                    {
                        progress("正在复制单元 " + copy.SourceZgsNo + " -> " + copy.NewZgsNo + "…");
                    }

                    long newUnitId = ExecuteAgentUnitCopy(conn, copy);
                    undo.Rows.Add(new AgentUndoRow { Kind = "CU", UnitId = newUnitId });
                    if (message.Length > 0) { message.AppendLine(); }
                    message.Append("已新建单元 ").Append(copy.NewZgsNo).Append(" \"").Append(copy.NewName)
                        .Append("\"（总概算序号 ").Append(newUnitId.ToString(CultureInfo.InvariantCulture))
                        .Append("）。注意：需要关闭并重新打开本项目，左侧树上才会显示新单元。");
                }

                foreach (AgentInsertGroup group in plan.Inserts)
                {
                    if (progress != null)
                    {
                        progress("正在向条目 " + group.ItemNo + " 粘贴 " + group.Quotas.Count.ToString(CultureInfo.InvariantCulture) + " 条定额…");
                    }

                    string groupMessage = ExecuteAgentInsertGroup(mainForm, conn, group, undo);
                    if (message.Length > 0) { message.AppendLine(); }
                    message.Append(groupMessage);
                }

                RefreshCurrentQuotaGrid(mainForm);
                if (plan.NeedsRecalc)
                {
                    if (message.Length > 0) { message.AppendLine(); }
                    message.Append("⚠ 本次改动涉及定额编号/单价/方案，主程序需要手工触发\"重算\"后造价才会更新。");
                }

                if (undo.Rows.Count > 0)
                {
                    GetAgentUndoStack(mainForm).Add(undo);
                    AppendAgentAuditLog(undo);
                    message.AppendLine().Append("（可输入\"撤销\"回滚本次操作）");
                }

                Log("Agent plan executed. " + plan.Summary);
                return message.Length == 0 ? "没有需要执行的改动。" : message.ToString();
            }
        }

        private static int ExecuteAgentFieldUpdate(SqlConnection conn, SqlTransaction transaction, string table, string keyClause, Dictionary<string, object> keyValues, Dictionary<string, object> newValues)
        {
            List<string> setParts = new List<string>();
            int i = 0;
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = transaction;
                foreach (KeyValuePair<string, object> pair in newValues)
                {
                    string p = "@v" + i.ToString(CultureInfo.InvariantCulture);
                    setParts.Add("[" + pair.Key + "]=" + p);
                    cmd.Parameters.AddWithValue(p, pair.Value ?? DBNull.Value);
                    i++;
                }

                foreach (KeyValuePair<string, object> pair in keyValues)
                {
                    cmd.Parameters.AddWithValue(pair.Key, pair.Value ?? DBNull.Value);
                }

                cmd.CommandText = "update [" + table + "] set " + String.Join(", ", setParts.ToArray()) + " where " + keyClause;
                return cmd.ExecuteNonQuery();
            }
        }

        private static void DeleteAgentDependentRows(SqlConnection conn, SqlTransaction transaction, List<long> quotaIds)
        {
            foreach (string table in AgentQuotaDependentTables)
            {
                // 不同项目模板的从属表不一定都有；先确认表与列都存在再删，
                // 避免在事务内对不存在的表执行 DELETE 抛错（可能让事务中毒）。
                if (!AgentDependentTableUsable(conn, transaction, table))
                {
                    continue;
                }

                for (int i = 0; i < quotaIds.Count; i += 500)
                {
                    string inList = String.Join(",", quotaIds.Skip(i).Take(500)
                        .Select(id => id.ToString(CultureInfo.InvariantCulture)).ToArray());
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandText = "delete from [" + table + "] where 定额序号 in (" + inList + ")";
                        int removed = cmd.ExecuteNonQuery();
                        if (removed > 0)
                        {
                            Log("Agent dependent cleanup: " + table + " removed " + removed.ToString(CultureInfo.InvariantCulture));
                        }
                    }
                }
            }
        }

        private static bool AgentDependentTableUsable(SqlConnection conn, SqlTransaction transaction, string table)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = "select count(*) from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME=@t and COLUMN_NAME=N'定额序号'";
                cmd.Parameters.AddWithValue("@t", table);
                return Convert.ToInt32(cmd.ExecuteScalar(), CultureInfo.InvariantCulture) > 0;
            }
        }

        private static void SyncAgentHostQuotaTable(Form mainForm, List<AgentFieldUpdate> updates, List<long> deletedIds)
        {
            try
            {
                DataTable hostTable = GetField<DataTable>(mainForm, "m_dtDeInput");
                if (hostTable == null || !hostTable.Columns.Contains("定额序号"))
                {
                    return;
                }

                Dictionary<long, AgentFieldUpdate> updateById = new Dictionary<long, AgentFieldUpdate>();
                foreach (AgentFieldUpdate update in updates)
                {
                    if (update.QuotaSequence > 0)
                    {
                        updateById[update.QuotaSequence] = update;
                    }
                }

                HashSet<long> deletes = new HashSet<long>(deletedIds);
                List<DataRow> toRemove = new List<DataRow>();
                foreach (DataRow row in hostTable.Rows)
                {
                    if (row.RowState == DataRowState.Deleted || row.RowState == DataRowState.Detached)
                    {
                        continue;
                    }

                    object idValue = row["定额序号"];
                    if (idValue == null || idValue == DBNull.Value)
                    {
                        continue;
                    }

                    long id = Convert.ToInt64(idValue, CultureInfo.InvariantCulture);
                    if (deletes.Contains(id))
                    {
                        toRemove.Add(row);
                        continue;
                    }

                    AgentFieldUpdate update;
                    if (updateById.TryGetValue(id, out update))
                    {
                        foreach (KeyValuePair<string, object> pair in update.NewValues)
                        {
                            if (hostTable.Columns.Contains(pair.Key))
                            {
                                row[pair.Key] = pair.Value ?? DBNull.Value;
                            }
                        }

                        row.AcceptChanges();
                    }
                }

                foreach (DataRow row in toRemove)
                {
                    hostTable.Rows.Remove(row);
                }

                if (toRemove.Count > 0 || updateById.Count > 0)
                {
                    Log("Agent host table synced. updated<=" + updateById.Count.ToString(CultureInfo.InvariantCulture) +
                        ", removed=" + toRemove.Count.ToString(CultureInfo.InvariantCulture));
                }
            }
            catch (Exception ex)
            {
                Log("Agent host table sync failed: " + ex.Message);
            }
        }

        private static long ExecuteAgentUnitCopy(SqlConnection conn, AgentUnitCopySpec spec)
        {
            using (SqlTransaction transaction = conn.BeginTransaction())
            {
                try
                {
                    List<string> columns = LoadAgentTableColumns(conn, transaction, "总概算信息", true);
                    columns.Remove("总概算序号");
                    List<string> selectParts = new List<string>();
                    foreach (string column in columns)
                    {
                        if (column == "总概算编号") { selectParts.Add("@newZgs"); }
                        else if (column == "编制范围") { selectParts.Add("@newName"); }
                        else { selectParts.Add("[" + column + "]"); }
                    }

                    long newUnitId;
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandText = "insert into 总概算信息 (" +
                            String.Join(", ", columns.Select(c => "[" + c + "]").ToArray()) +
                            ") select " + String.Join(", ", selectParts.ToArray()) +
                            " from 总概算信息 where 总概算序号=@src; select cast(scope_identity() as bigint)";
                        cmd.Parameters.AddWithValue("@newZgs", spec.NewZgsNo);
                        cmd.Parameters.AddWithValue("@newName", spec.NewName);
                        cmd.Parameters.AddWithValue("@src", spec.SourceUnitId);
                        object result = cmd.ExecuteScalar();
                        if (result == null || result == DBNull.Value)
                        {
                            throw new AgentPlanException("新建总概算信息失败，未取得新序号。");
                        }

                        newUnitId = Convert.ToInt64(result, CultureInfo.InvariantCulture);
                    }

                    foreach (string table in AgentUnitCopyTables)
                    {
                        List<string> tableColumns;
                        try
                        {
                            tableColumns = LoadAgentTableColumns(conn, transaction, table, false);
                        }
                        catch (Exception ex)
                        {
                            Log("Agent unit copy skip table " + table + ": " + ex.Message);
                            continue;
                        }

                        if (!tableColumns.Contains("总概算序号"))
                        {
                            continue;
                        }

                        List<string> copySelect = new List<string>();
                        foreach (string column in tableColumns)
                        {
                            if (column == "总概算序号") { copySelect.Add(newUnitId.ToString(CultureInfo.InvariantCulture)); }
                            else if (table == "单项概算信息" && column == "单项概算编号") { copySelect.Add("replace([单项概算编号], @oldPrefix, @newPrefix)"); }
                            else { copySelect.Add("[" + column + "]"); }
                        }

                        using (SqlCommand cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = transaction;
                            cmd.CommandText = "insert into [" + table + "] (" +
                                String.Join(", ", tableColumns.Select(c => "[" + c + "]").ToArray()) +
                                ") select " + String.Join(", ", copySelect.ToArray()) +
                                " from [" + table + "] where 总概算序号=@src";
                            if (table == "单项概算信息")
                            {
                                cmd.Parameters.AddWithValue("@oldPrefix", spec.SourceZgsNo + "-");
                                cmd.Parameters.AddWithValue("@newPrefix", spec.NewZgsNo + "-");
                            }

                            cmd.Parameters.AddWithValue("@src", spec.SourceUnitId);
                            cmd.CommandTimeout = 120;
                            int copied = cmd.ExecuteNonQuery();
                            Log("Agent unit copy: " + table + " copied " + copied.ToString(CultureInfo.InvariantCulture));
                        }
                    }

                    transaction.Commit();
                    return newUnitId;
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
            }
        }

        private static List<string> LoadAgentTableColumns(SqlConnection conn, SqlTransaction transaction, string table, bool excludeIdentity)
        {
            List<string> columns = new List<string>();
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText =
                    "select COLUMN_NAME, COLUMNPROPERTY(OBJECT_ID(@t), COLUMN_NAME, 'IsIdentity') " +
                    "from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME=@t order by ORDINAL_POSITION";
                cmd.Parameters.AddWithValue("@t", table);
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        bool isIdentity = !reader.IsDBNull(1) && Convert.ToInt32(reader.GetValue(1), CultureInfo.InvariantCulture) == 1;
                        if (isIdentity)
                        {
                            continue;
                        }

                        columns.Add(Convert.ToString(reader.GetValue(0)));
                    }
                }
            }

            if (columns.Count == 0)
            {
                throw new AgentPlanException("表 " + table + " 不存在或没有可复制的列。");
            }

            return columns;
        }

        private static int CountAgentExistingRows(SqlConnection conn, List<long> ids)
        {
            int total = 0;
            for (int i = 0; i < ids.Count; i += 500)
            {
                List<long> chunk = ids.Skip(i).Take(500).ToList();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "select count(*) from 定额输入 where 定额序号 in (" +
                        String.Join(",", chunk.Select(id => id.ToString(CultureInfo.InvariantCulture)).ToArray()) + ")";
                    total += Convert.ToInt32(cmd.ExecuteScalar(), CultureInfo.InvariantCulture);
                }
            }

            return total;
        }

        private static Dictionary<string, object> LoadAgentFullRow(SqlConnection conn, long quotaSequence)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select * from 定额输入 where 定额序号=@id";
                cmd.Parameters.AddWithValue("@id", quotaSequence);
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    DataTable table = new DataTable();
                    adapter.Fill(table);
                    if (table.Rows.Count == 0)
                    {
                        return null;
                    }

                    Dictionary<string, object> values = new Dictionary<string, object>();
                    foreach (DataColumn column in table.Columns)
                    {
                        values[column.ColumnName] = table.Rows[0][column];
                    }

                    return values;
                }
            }
        }

        // ===== 插入路径：复用主程序自身的粘贴管线 =====

        private static string ExecuteAgentInsertGroup(Form mainForm, SqlConnection conn, AgentInsertGroup group, AgentUndoRecord undo)
        {
            HashSet<long> before = LoadAgentItemQuotaIds(conn, group.ItemNo);
            if (!TryNavigateToAgentItem(mainForm, conn, group.ItemNo))
            {
                return "条目 " + group.ItemNo + "：未能在左侧树上定位该条目，已跳过插入。请手动打开该条目后重试。";
            }

            WaitAgentUiIdle(800);
            DataGridView grid = GetField<DataGridView>(mainForm, "dataGridViewDE");
            if (grid == null)
            {
                return "条目 " + group.ItemNo + "：没有找到定额输入表格，已跳过插入。";
            }

            StringBuilder text = new StringBuilder();
            foreach (AgentQuotaInput quota in group.Quotas)
            {
                text.Append(CleanAgentCell(quota.Code)).Append('\t').Append('\t').Append('\t')
                    .Append(CleanAgentCell(quota.Quantity)).Append("\r\n");
            }

            try
            {
                grid.Focus();
                MoveAgentGridToNewRow(grid);
                Clipboard.SetText(text.ToString());
                if (!TryInvokeAgentPasteMenu(mainForm))
                {
                    SendKeys.SendWait("^v");
                }
            }
            catch (Exception ex)
            {
                Log("Agent paste failed: " + ex);
                return "条目 " + group.ItemNo + "：粘贴失败（" + ex.Message + "）。";
            }

            WaitAgentUiIdle(1200);
            HashSet<long> after = LoadAgentItemQuotaIds(conn, group.ItemNo);
            List<long> added = after.Where(id => !before.Contains(id)).ToList();
            foreach (long id in added)
            {
                undo.Rows.Add(new AgentUndoRow { Kind = "I", QuotaSequence = id });
            }

            if (added.Count == 0)
            {
                return "条目 " + group.ItemNo + "：已发送粘贴，但数据库中暂未检测到新行（主程序可能尚未落库）。请在界面上核对，该部分插入暂不支持撤销。";
            }

            return "条目 " + group.ItemNo + "：已插入 " + added.Count.ToString(CultureInfo.InvariantCulture) + " 条定额。";
        }

        private static string CleanAgentCell(string value)
        {
            return (value ?? "").Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim();
        }

        private static HashSet<long> LoadAgentItemQuotaIds(SqlConnection conn, string itemNo)
        {
            HashSet<long> ids = new HashSet<long>();
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select DE.定额序号 from 定额输入 DE inner join 章节表 ZJ on DE.条目序号=ZJ.条目序号 where ZJ.条目编号=@bh";
                cmd.Parameters.AddWithValue("@bh", itemNo);
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        ids.Add(Convert.ToInt64(reader.GetValue(0), CultureInfo.InvariantCulture));
                    }
                }
            }

            return ids;
        }

        private static void WaitAgentUiIdle(int milliseconds)
        {
            DateTime deadline = DateTime.Now.AddMilliseconds(milliseconds);
            while (DateTime.Now < deadline)
            {
                Application.DoEvents();
                Thread.Sleep(50);
            }
        }

        private static bool TryNavigateToAgentItem(Form mainForm, SqlConnection conn, string itemNo)
        {
            string seq = null;
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select top 1 条目序号 from 章节表 where 条目编号=@bh";
                cmd.Parameters.AddWithValue("@bh", itemNo);
                object result = cmd.ExecuteScalar();
                if (result == null || result == DBNull.Value)
                {
                    return false;
                }

                seq = Convert.ToString(result, CultureInfo.InvariantCulture).Trim();
            }

            TreeView tree = GetField<TreeView>(mainForm, "Tv_tree");
            if (tree == null)
            {
                return false;
            }

            TreeNode node = FindAgentTreeNode(tree.Nodes, seq, itemNo);
            if (node == null)
            {
                return false;
            }

            try
            {
                tree.SelectedNode = node;
                node.EnsureVisible();
                return true;
            }
            catch (Exception ex)
            {
                Log("Agent navigate tree failed: " + ex.Message);
                return false;
            }
        }

        private static TreeNode FindAgentTreeNode(TreeNodeCollection nodes, string seq, string itemNo)
        {
            foreach (TreeNode node in nodes)
            {
                if (node.Name == seq ||
                    TryGetValue(node.Tag, "条目序号") == seq ||
                    TryGetValue(node.Tag, "条目编号") == itemNo)
                {
                    return node;
                }

                TreeNode child = FindAgentTreeNode(node.Nodes, seq, itemNo);
                if (child != null)
                {
                    return child;
                }
            }

            return null;
        }

        private static void MoveAgentGridToNewRow(DataGridView grid)
        {
            int rowIndex = grid.AllowUserToAddRows && grid.NewRowIndex >= 0 ? grid.NewRowIndex : grid.Rows.Count - 1;
            if (rowIndex < 0)
            {
                return;
            }

            int columnIndex = -1;
            foreach (DataGridViewColumn column in grid.Columns)
            {
                if (column.Visible &&
                    ((column.Name != null && column.Name.IndexOf("定额编号", StringComparison.Ordinal) >= 0) ||
                     (column.HeaderText != null && column.HeaderText.IndexOf("定额编号", StringComparison.Ordinal) >= 0)))
                {
                    columnIndex = column.Index;
                    break;
                }
            }

            if (columnIndex < 0)
            {
                foreach (DataGridViewColumn column in grid.Columns)
                {
                    if (column.Visible)
                    {
                        columnIndex = column.Index;
                        break;
                    }
                }
            }

            if (columnIndex < 0)
            {
                return;
            }

            try
            {
                grid.CurrentCell = grid.Rows[rowIndex].Cells[columnIndex];
            }
            catch (Exception ex)
            {
                Log("Agent move to new row failed: " + ex.Message);
            }
        }

        private static bool TryInvokeAgentPasteMenu(Form mainForm)
        {
            ContextMenuStrip menu = GetField<ContextMenuStrip>(mainForm, "contextMenuStripDE");
            if (menu == null)
            {
                return false;
            }

            foreach (ToolStripItem item in menu.Items)
            {
                if (item is ToolStripMenuItem && item.Available && item.Text != null &&
                    item.Text.IndexOf("粘贴", StringComparison.Ordinal) >= 0)
                {
                    try
                    {
                        ((ToolStripMenuItem)item).PerformClick();
                        Log("Agent paste via menu item: " + item.Text);
                        return true;
                    }
                    catch (Exception ex)
                    {
                        Log("Agent paste menu failed: " + ex.Message);
                        return false;
                    }
                }
            }

            return false;
        }

        // ===== 撤销 =====

        private static AgentPlan BuildAgentUndoPlan(Form mainForm)
        {
            List<AgentUndoRecord> stack = GetAgentUndoStack(mainForm);
            if (stack.Count == 0)
            {
                throw new AgentPlanException("没有可撤销的智能体操作（撤销栈只记录本次软件运行期间由聊天指令执行的修改）。");
            }

            AgentUndoRecord record = stack[stack.Count - 1];
            AgentPlan plan = new AgentPlan();
            plan.IsUndo = true;
            plan.UndoRecord = record;
            plan.Summary = "撤销：" + record.Summary + "（" + record.Time.ToString("HH:mm:ss", CultureInfo.InvariantCulture) + "）";
            foreach (AgentUndoRow row in record.Rows)
            {
                AgentPlanRow display = new AgentPlanRow();
                if (row.Kind == "F")
                {
                    display.Action = "还原字段";
                    display.NewValue = String.Join("; ", row.OldValues.Select(p => p.Key + "=" + FormatAgentUndoValue(p.Value)).ToArray());
                }
                else if (row.Kind == "D")
                {
                    display.Action = "恢复整行";
                    display.QuotaCode = row.FullRow != null && row.FullRow.ContainsKey("定额编号") ? Convert.ToString(row.FullRow["定额编号"]) : "";
                    display.NewValue = "(重插被删除的行)";
                }
                else if (row.Kind == "CU")
                {
                    display.Action = "删除新建单元";
                    display.UnitId = row.UnitId;
                    display.NewValue = "(删除总概算序号 " + row.UnitId.ToString(CultureInfo.InvariantCulture) + " 及其全部数据)";
                }
                else
                {
                    display.Action = "删除插入行";
                    display.NewValue = "(删除当时插入的行)";
                }

                plan.PreviewRows.Add(display);
            }

            return plan;
        }

        private static string FormatAgentUndoValue(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return "(空)";
            }

            string text = Convert.ToString(value, CultureInfo.InvariantCulture);
            return text.Length > 40 ? text.Substring(0, 40) + "…" : text;
        }

        private static string ExecuteAgentUndo(Form mainForm, SqlConnection conn, AgentUndoRecord record)
        {
            int restored = 0;
            int reinserted = 0;
            int removed = 0;
            int unitsRemoved = 0;
            List<AgentFieldUpdate> hostUpdates = new List<AgentFieldUpdate>();
            List<long> hostDeletes = new List<long>();
            using (SqlTransaction transaction = conn.BeginTransaction())
            {
                try
                {
                    for (int i = record.Rows.Count - 1; i >= 0; i--)
                    {
                        AgentUndoRow row = record.Rows[i];
                        if (row.Kind == "F")
                        {
                            restored += ExecuteAgentFieldUpdate(conn, transaction, row.Table, row.KeyClause, row.KeyValues, row.OldValues);
                            if (row.Table == "定额输入" && row.QuotaSequence > 0)
                            {
                                hostUpdates.Add(new AgentFieldUpdate { QuotaSequence = row.QuotaSequence, NewValues = row.OldValues });
                            }
                        }
                        else if (row.Kind == "I")
                        {
                            using (SqlCommand cmd = conn.CreateCommand())
                            {
                                cmd.Transaction = transaction;
                                cmd.CommandText = "delete from 定额输入 where 定额序号=@id";
                                cmd.Parameters.AddWithValue("@id", row.QuotaSequence);
                                removed += cmd.ExecuteNonQuery();
                            }

                            hostDeletes.Add(row.QuotaSequence);
                        }
                        else if (row.Kind == "D" && row.FullRow != null)
                        {
                            reinserted += ReinsertAgentRow(conn, transaction, row.FullRow);
                        }
                        else if (row.Kind == "CU")
                        {
                            foreach (string table in AgentUnitCopyTables)
                            {
                                using (SqlCommand cmd = conn.CreateCommand())
                                {
                                    cmd.Transaction = transaction;
                                    cmd.CommandText = "delete from [" + table + "] where 总概算序号=@id";
                                    cmd.Parameters.AddWithValue("@id", row.UnitId);
                                    cmd.CommandTimeout = 120;
                                    try { cmd.ExecuteNonQuery(); }
                                    catch (Exception ex) { Log("Agent undo unit table " + table + " failed: " + ex.Message); }
                                }
                            }

                            using (SqlCommand cmd = conn.CreateCommand())
                            {
                                cmd.Transaction = transaction;
                                cmd.CommandText = "delete from 总概算信息 where 总概算序号=@id";
                                cmd.Parameters.AddWithValue("@id", row.UnitId);
                                unitsRemoved += cmd.ExecuteNonQuery();
                            }
                        }
                    }

                    if (hostDeletes.Count > 0)
                    {
                        DeleteAgentDependentRows(conn, transaction, hostDeletes);
                    }

                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
            }

            SyncAgentHostQuotaTable(mainForm, hostUpdates, hostDeletes);
            GetAgentUndoStack(mainForm).Remove(record);
            RefreshCurrentQuotaGrid(mainForm);
            Log("Agent undo executed. " + record.Summary);
            StringBuilder message = new StringBuilder();
            message.Append("撤销完成：还原 ").Append(restored.ToString(CultureInfo.InvariantCulture))
                .Append(" 处，恢复被删行 ").Append(reinserted.ToString(CultureInfo.InvariantCulture))
                .Append(" 行，删除插入行 ").Append(removed.ToString(CultureInfo.InvariantCulture)).Append(" 行");
            if (unitsRemoved > 0)
            {
                message.Append("，删除新建单元 ").Append(unitsRemoved.ToString(CultureInfo.InvariantCulture)).Append(" 个");
            }

            message.Append("。");
            if (IsAgentRecordRedoable(record))
            {
                GetAgentRedoStack(mainForm).Add(record);
                message.Append("（可输入\"重做\"恢复）");
            }
            else
            {
                message.Append("（含删除/插入/新建单元，不支持重做）");
            }

            return message.ToString();
        }

        // ===== 重做（撤回上一次撤销） =====

        private static AgentPlan BuildAgentRedoPlan(Form mainForm)
        {
            List<AgentUndoRecord> stack = GetAgentRedoStack(mainForm);
            if (stack.Count == 0)
            {
                throw new AgentPlanException("没有可重做的操作（先\"撤销\"过一次纯字段修改，才能\"重做\"）。");
            }

            AgentUndoRecord record = stack[stack.Count - 1];
            AgentPlan plan = new AgentPlan();
            plan.IsRedo = true;
            plan.UndoRecord = record;
            plan.Summary = "重做：" + record.Summary;
            foreach (AgentUndoRow row in record.Rows)
            {
                plan.PreviewRows.Add(new AgentPlanRow
                {
                    Action = "重做字段",
                    NewValue = String.Join("; ", row.NewValues.Select(p => p.Key + "=" + FormatAgentUndoValue(p.Value)).ToArray())
                });
            }

            return plan;
        }

        private static string ExecuteAgentRedo(Form mainForm, SqlConnection conn, AgentUndoRecord record)
        {
            int applied = 0;
            List<AgentFieldUpdate> hostUpdates = new List<AgentFieldUpdate>();
            using (SqlTransaction transaction = conn.BeginTransaction())
            {
                try
                {
                    foreach (AgentUndoRow row in record.Rows)
                    {
                        applied += ExecuteAgentFieldUpdate(conn, transaction, row.Table, row.KeyClause, row.KeyValues, row.NewValues);
                        if (row.Table == "定额输入" && row.QuotaSequence > 0)
                        {
                            hostUpdates.Add(new AgentFieldUpdate { QuotaSequence = row.QuotaSequence, NewValues = row.NewValues });
                        }
                    }

                    transaction.Commit();
                }
                catch
                {
                    transaction.Rollback();
                    throw;
                }
            }

            SyncAgentHostQuotaTable(mainForm, hostUpdates, new List<long>());
            GetAgentRedoStack(mainForm).Remove(record);
            GetAgentUndoStack(mainForm).Add(record);   // 重做后又可以再撤销
            RefreshCurrentQuotaGrid(mainForm);
            Log("Agent redo executed. " + record.Summary);
            return "重做完成：重新应用 " + applied.ToString(CultureInfo.InvariantCulture) + " 处修改。（可再\"撤销\"）";
        }

        private static int ReinsertAgentRow(SqlConnection conn, SqlTransaction transaction, Dictionary<string, object> fullRow)
        {
            try
            {
                return ReinsertAgentRowCore(conn, transaction, fullRow, true);
            }
            catch (Exception ex)
            {
                Log("Agent reinsert with identity failed, retry without identity: " + ex.Message);
                Dictionary<string, object> withoutId = fullRow
                    .Where(pair => pair.Key != "定额序号")
                    .ToDictionary(pair => pair.Key, pair => pair.Value);
                return ReinsertAgentRowCore(conn, transaction, withoutId, false);
            }
        }

        private static int ReinsertAgentRowCore(SqlConnection conn, SqlTransaction transaction, Dictionary<string, object> values, bool identityInsert)
        {
            List<string> columns = values.Keys.ToList();
            StringBuilder sql = new StringBuilder();
            if (identityInsert)
            {
                sql.Append("set identity_insert 定额输入 on; ");
            }

            sql.Append("insert into 定额输入 (")
                .Append(String.Join(", ", columns.Select(c => "[" + c + "]").ToArray()))
                .Append(") values (")
                .Append(String.Join(", ", columns.Select((c, i) => "@p" + i.ToString(CultureInfo.InvariantCulture)).ToArray()))
                .Append(")");
            if (identityInsert)
            {
                sql.Append("; set identity_insert 定额输入 off;");
            }

            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = sql.ToString();
                for (int i = 0; i < columns.Count; i++)
                {
                    cmd.Parameters.AddWithValue("@p" + i.ToString(CultureInfo.InvariantCulture), values[columns[i]] ?? DBNull.Value);
                }

                return cmd.ExecuteNonQuery();
            }
        }

        private static void AppendAgentAuditLog(AgentUndoRecord record)
        {
            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                serializer.MaxJsonLength = 1024 * 1024 * 16;
                Dictionary<string, object> entry = new Dictionary<string, object>();
                entry["time"] = record.Time.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
                entry["summary"] = record.Summary;
                List<object> rows = new List<object>();
                foreach (AgentUndoRow row in record.Rows)
                {
                    Dictionary<string, object> item = new Dictionary<string, object>();
                    item["kind"] = row.Kind;
                    item["quota_sequence"] = row.QuotaSequence;
                    if (row.Kind == "F")
                    {
                        item["table"] = row.Table;
                        item["old_values"] = AgentSanitizeForJson(row.OldValues);
                    }
                    else if (row.Kind == "D" && row.FullRow != null)
                    {
                        item["full_row"] = AgentSanitizeForJson(row.FullRow);
                    }
                    else if (row.Kind == "CU")
                    {
                        item["unit_id"] = row.UnitId;
                    }

                    rows.Add(item);
                }

                entry["rows"] = rows;
                string path = Path.Combine(FindRecoQuotaDataDir(), "agent-undo.jsonl");
                File.AppendAllText(path, serializer.Serialize(entry) + Environment.NewLine, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Log("Agent audit log failed: " + ex.Message);
            }
        }

        private static Dictionary<string, object> AgentSanitizeForJson(Dictionary<string, object> values)
        {
            Dictionary<string, object> clean = new Dictionary<string, object>();
            foreach (KeyValuePair<string, object> pair in values)
            {
                clean[pair.Key] = pair.Value == DBNull.Value ? null : pair.Value;
            }

            return clean;
        }
    }
}
