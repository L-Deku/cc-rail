using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        private sealed class AgentBaseline
        {
            public DateTime Time;
            public string Database;
            public long UnitId;   // >0 表示按单元(总概算序号)限定的基线
            public Dictionary<string, long> RowCounts = new Dictionary<string, long>(StringComparer.Ordinal);
            public Dictionary<string, string> WatchedSpecs = new Dictionary<string, string>(StringComparer.Ordinal);  // 表名 -> 键列spec
            public Dictionary<string, Dictionary<string, Dictionary<string, object>>> WatchedRows =
                new Dictionary<string, Dictionary<string, Dictionary<string, object>>>(StringComparer.Ordinal);
        }

        private static readonly Dictionary<Form, AgentBaseline> AgentBaselines = new Dictionary<Form, AgentBaseline>();

        // 基线/对比时整行快照的表（表名 -> 主键列，复合主键用 | 分隔）
        private static readonly string[][] AgentWatchedTables = new string[][]
        {
            new string[] { "定额输入", "定额序号" },
            new string[] { "章节表", "条目序号" },
            new string[] { "单项概算信息", "单项概算序号" },
            new string[] { "总概算条目", "总概算序号|条目序号" },
            new string[] { "材料单价", "方案名称|电算代号" }
        };

        // 按单元(总概算序号)限定基线时观察的级联候选表。含 总概算序号 列的会按单元过滤，否则整表快照。
        // 键列留空时运行时自动取主键。用于摸清"改运输方案/材料方案"会连带改哪些计算结果表。
        private static readonly string[][] AgentUnitWatchedTables = new string[][]
        {
            new string[] { "总概算条目", "总概算序号|条目序号" },
            new string[] { "定额输入", "定额序号" },
            new string[] { "总概算信息", "总概算序号" },
            new string[] { "单项概算费用汇总", null },
            new string[] { "工料机计算表", null },
            new string[] { "劳材统计表", null },
            new string[] { "主劳材统计表", null },
            new string[] { "运杂费计算明细", null },
            new string[] { "材料运输方案", "材料分类序号|方案序号" }
        };

        private const int AgentWatchedRowCap = 80000;

        private static bool AgentTableHasColumn(SqlConnection conn, string table, string column)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select count(*) from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME=@t and COLUMN_NAME=@c";
                cmd.Parameters.AddWithValue("@t", table);
                cmd.Parameters.AddWithValue("@c", column);
                return Convert.ToInt32(cmd.ExecuteScalar(), CultureInfo.InvariantCulture) > 0;
            }
        }

        private static void AgentDiagLog(string message)
        {
            try
            {
                string path = Path.Combine(FindRecoQuotaDataDir(), "agent-diagnostics.log");
                File.AppendAllText(path, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + " " + message + Environment.NewLine, Encoding.UTF8);
            }
            catch
            {
            }
        }

        // 入口：处理 “探查 …” 指令。output 把行写进聊天窗，所有内容同时落 agent-diagnostics.log。
        private static void RunAgentDiagnostics(Form mainForm, string args, Action<string> output)
        {
            Action<string> emit = delegate(string line)
            {
                output(line);
                AgentDiagLog(line);
            };

            string trimmed = (args ?? "").Trim();
            try
            {
                if (trimmed == "表")
                {
                    AgentDiagTables(mainForm, emit);
                }
                else if (trimmed.StartsWith("列", StringComparison.Ordinal))
                {
                    string table = trimmed.Substring(1).Trim();
                    if (table.Length == 0)
                    {
                        emit("用法：探查 列 表名（例如：探查 列 定额输入）");
                    }
                    else
                    {
                        AgentDiagColumns(mainForm, table, emit);
                    }
                }
                else if (trimmed.StartsWith("基线单元", StringComparison.Ordinal))
                {
                    AgentDiagBaselineUnit(mainForm, trimmed.Substring(4).Trim(), emit);
                }
                else if (trimmed.StartsWith("基线", StringComparison.Ordinal))
                {
                    AgentDiagBaseline(mainForm, trimmed.Substring(2).Trim(), emit);
                }
                else if (trimmed == "对比")
                {
                    AgentDiagCompare(mainForm, emit);
                }
                else if (trimmed == "窗体")
                {
                    AgentDiagMainForm(mainForm, emit);
                }
                else
                {
                    emit("探查子命令：");
                    emit("  探查 表        —— 当前项目库全部表及行数");
                    emit("  探查 列 表名   —— 某个表的列结构（含自增标记）");
                    emit("  探查 基线      —— 记录当前数据快照（之后在软件里手工操作）");
                    emit("  探查 基线 表名,表名 —— 额外指定要做行级快照的表（无主键的表用 表名:键列|键列）");
                    emit("  探查 基线单元 单元 —— 只快照该单元相关的大表，用于摸清改运输方案/材料方案的级联改动");
                    emit("  探查 对比      —— 与基线对比，列出行数变化和关键表的行级差异");
                    emit("  探查 窗体      —— dump 主程序窗体字段和疑似可用方法（详情看 agent-diagnostics.log）");
                }
            }
            catch (Exception ex)
            {
                emit("探查出错：" + ex.Message);
                Log("Agent diagnostics failed: " + ex);
            }
        }

        private static SqlConnection RequireAgentConnection(Form mainForm)
        {
            SqlConnection conn = GetProjectConnection(mainForm);
            if (conn == null)
            {
                throw new AgentPlanException("没有找到当前项目数据库连接，请先打开一个项目。");
            }

            EnsureOpen(conn);
            return conn;
        }

        private static Dictionary<string, long> LoadAgentTableRowCounts(SqlConnection conn)
        {
            Dictionary<string, long> counts = new Dictionary<string, long>(StringComparer.Ordinal);
            try
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText =
                        "select t.name, sum(p.rows) from sys.tables t " +
                        "inner join sys.partitions p on p.object_id=t.object_id and p.index_id in (0,1) " +
                        "group by t.name order by t.name";
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            counts[Convert.ToString(reader.GetValue(0))] = Convert.ToInt64(reader.GetValue(1), CultureInfo.InvariantCulture);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("Agent row counts via sys.partitions failed, fallback: " + ex.Message);
                List<string> tables = new List<string>();
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "select TABLE_NAME from INFORMATION_SCHEMA.TABLES where TABLE_TYPE='BASE TABLE' order by TABLE_NAME";
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            tables.Add(Convert.ToString(reader.GetValue(0)));
                        }
                    }
                }

                foreach (string table in tables)
                {
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "select count(*) from [" + table.Replace("]", "]]") + "]";
                        counts[table] = Convert.ToInt64(cmd.ExecuteScalar(), CultureInfo.InvariantCulture);
                    }
                }
            }

            return counts;
        }

        private static void AgentDiagTables(Form mainForm, Action<string> emit)
        {
            SqlConnection conn = RequireAgentConnection(mainForm);
            emit("当前项目库：" + conn.Database);
            Dictionary<string, long> counts = LoadAgentTableRowCounts(conn);
            foreach (KeyValuePair<string, long> pair in counts.OrderBy(p => p.Key, StringComparer.Ordinal))
            {
                emit("  " + pair.Key + " = " + pair.Value.ToString(CultureInfo.InvariantCulture) + " 行");
            }

            emit("共 " + counts.Count.ToString(CultureInfo.InvariantCulture) + " 张表。");
        }

        private static void AgentDiagColumns(Form mainForm, string table, Action<string> emit)
        {
            SqlConnection conn = RequireAgentConnection(mainForm);
            int count = 0;
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    "select COLUMN_NAME, DATA_TYPE, IS_NULLABLE, " +
                    "COLUMNPROPERTY(OBJECT_ID(@t), COLUMN_NAME, 'IsIdentity') " +
                    "from INFORMATION_SCHEMA.COLUMNS where TABLE_NAME=@t order by ORDINAL_POSITION";
                cmd.Parameters.AddWithValue("@t", table);
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        count++;
                        string identity = !reader.IsDBNull(3) && Convert.ToInt32(reader.GetValue(3), CultureInfo.InvariantCulture) == 1 ? " [自增]" : "";
                        string nullable = Convert.ToString(reader.GetValue(2)) == "YES" ? " 可空" : "";
                        emit("  " + Convert.ToString(reader.GetValue(0)) + " : " + Convert.ToString(reader.GetValue(1)) + nullable + identity);
                    }
                }
            }

            emit(count == 0 ? "表 " + table + " 不存在或没有列。" : "表 " + table + " 共 " + count.ToString(CultureInfo.InvariantCulture) + " 列。");
        }

        private static string BuildAgentRowKey(IDataRecord record, int[] keyOrdinals)
        {
            StringBuilder key = new StringBuilder();
            foreach (int ordinal in keyOrdinals)
            {
                if (key.Length > 0)
                {
                    key.Append('|');
                }

                key.Append(record.IsDBNull(ordinal) ? "<null>" : Convert.ToString(record.GetValue(ordinal), CultureInfo.InvariantCulture));
            }

            return key.ToString();
        }

        // unitId>0 且表含 总概算序号 列时，只快照该单元的行（让大表也能做行级diff）。
        private static Dictionary<string, Dictionary<string, object>> LoadAgentWatchedRows(SqlConnection conn, string table, string keySpec, long rowCount, long unitId)
        {
            Dictionary<string, Dictionary<string, object>> rows = new Dictionary<string, Dictionary<string, object>>(StringComparer.Ordinal);
            bool filtered = unitId > 0 && AgentTableHasColumn(conn, table, "总概算序号");
            if (!filtered && rowCount > AgentWatchedRowCap)
            {
                return null;
            }

            string[] keyColumns = keySpec.Split('|');
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select * from [" + table.Replace("]", "]]") + "]" +
                    (filtered ? " where 总概算序号=@u" : "");
                if (filtered)
                {
                    cmd.Parameters.AddWithValue("@u", unitId);
                }

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    int[] keyOrdinals = keyColumns.Select(c => reader.GetOrdinal(c)).ToArray();
                    while (reader.Read())
                    {
                        if (rows.Count > AgentWatchedRowCap)
                        {
                            return null;   // 即便过滤后仍过大
                        }

                        Dictionary<string, object> values = new Dictionary<string, object>(StringComparer.Ordinal);
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            values[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                        }

                        rows[BuildAgentRowKey(reader, keyOrdinals)] = values;
                    }
                }
            }

            return rows;
        }

        // 查询表的主键列（复合主键按顺序 | 连接），没有主键返回 null。
        private static string LoadAgentPrimaryKeySpec(SqlConnection conn, string table)
        {
            List<string> columns = new List<string>();
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    "select c.COLUMN_NAME from INFORMATION_SCHEMA.TABLE_CONSTRAINTS t " +
                    "inner join INFORMATION_SCHEMA.KEY_COLUMN_USAGE c on c.CONSTRAINT_NAME=t.CONSTRAINT_NAME " +
                    "where t.TABLE_NAME=@t and t.CONSTRAINT_TYPE='PRIMARY KEY' order by c.ORDINAL_POSITION";
                cmd.Parameters.AddWithValue("@t", table);
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        columns.Add(Convert.ToString(reader.GetValue(0)));
                    }
                }
            }

            return columns.Count == 0 ? null : String.Join("|", columns.ToArray());
        }

        private static void AgentDiagBaseline(Form mainForm, string extraTables, Action<string> emit)
        {
            SqlConnection conn = RequireAgentConnection(mainForm);
            AgentBaseline baseline = new AgentBaseline();
            baseline.Time = DateTime.Now;
            baseline.Database = conn.Database;
            baseline.RowCounts = LoadAgentTableRowCounts(conn);

            Dictionary<string, string> specs = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (string[] watched in AgentWatchedTables)
            {
                specs[watched[0]] = watched[1];
            }

            // 额外指定的表：表名 或 表名:键列|键列；未给键列时自动查主键。
            foreach (string raw in (extraTables ?? "").Split(new char[] { ',', '，', ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string spec = raw.Trim();
                if (spec.Length == 0)
                {
                    continue;
                }

                string table = spec;
                string keySpec = null;
                int colon = spec.IndexOf(':');
                if (colon > 0)
                {
                    table = spec.Substring(0, colon).Trim();
                    keySpec = spec.Substring(colon + 1).Trim();
                }

                if (!baseline.RowCounts.ContainsKey(table))
                {
                    emit("  表 " + table + " 不存在，跳过。");
                    continue;
                }

                if (String.IsNullOrEmpty(keySpec))
                {
                    keySpec = LoadAgentPrimaryKeySpec(conn, table);
                }

                if (String.IsNullOrEmpty(keySpec))
                {
                    emit("  表 " + table + " 没有主键，请用 表名:键列|键列 的格式指定，已跳过。");
                    continue;
                }

                specs[table] = keySpec;
            }

            baseline.WatchedSpecs = specs;
            foreach (KeyValuePair<string, string> watched in specs)
            {
                long count;
                if (!baseline.RowCounts.TryGetValue(watched.Key, out count))
                {
                    continue;
                }

                Dictionary<string, Dictionary<string, object>> rows;
                try
                {
                    rows = LoadAgentWatchedRows(conn, watched.Key, watched.Value, count, 0);
                }
                catch (Exception ex)
                {
                    emit("  表 " + watched.Key + " 快照失败（" + ex.Message + "），跳过。");
                    continue;
                }

                if (rows == null)
                {
                    emit("  注意：表 " + watched.Key + " 行数超过 " + AgentWatchedRowCap.ToString(CultureInfo.InvariantCulture) + "，只对比行数不做行级diff。");
                }
                else
                {
                    baseline.WatchedRows[watched.Key] = rows;
                }
            }

            AgentBaselines[mainForm] = baseline;
            emit("基线已记录（库 " + baseline.Database + "，" + baseline.RowCounts.Count.ToString(CultureInfo.InvariantCulture) +
                " 张表，行级快照 " + baseline.WatchedRows.Count.ToString(CultureInfo.InvariantCulture) + " 张）。");
            emit("现在请在软件里手工完成目标操作（例如新建单项概算/设置材料价），完成后输入：探查 对比");
        }

        // 解析单元 token（名称/_ZGS_编号/总概算序号）-> 总概算序号，0=未找到/不唯一。
        private static long ResolveAgentUnitIdSimple(SqlConnection conn, string token, out string label)
        {
            label = "";
            List<long> matched = new List<long>();
            List<string> labels = new List<string>();
            using (SqlCommand cmd = conn.CreateCommand())
            {
                long numeric;
                if (Int64.TryParse(token, NumberStyles.Integer, CultureInfo.InvariantCulture, out numeric))
                {
                    cmd.CommandText = "select 总概算序号, 总概算编号, 编制范围 from 总概算信息 where 总概算序号=@v";
                    cmd.Parameters.AddWithValue("@v", numeric);
                }
                else
                {
                    cmd.CommandText = "select 总概算序号, 总概算编号, 编制范围 from 总概算信息 where 总概算编号=@v or 编制范围=@v";
                    cmd.Parameters.AddWithValue("@v", token);
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

            if (matched.Count == 1)
            {
                label = labels[0];
                return matched[0];
            }

            if (matched.Count > 1)
            {
                label = "多个匹配：" + String.Join("；", labels.ToArray());
            }

            return 0;
        }

        private static void AgentDiagBaselineUnit(Form mainForm, string unitToken, Action<string> emit)
        {
            SqlConnection conn = RequireAgentConnection(mainForm);
            if (String.IsNullOrEmpty(unitToken))
            {
                emit("用法：探查 基线单元 单元名/_ZGS_编号/总概算序号，例如：探查 基线单元 南江路泵房");
                return;
            }

            string label;
            long unitId = ResolveAgentUnitIdSimple(conn, unitToken, out label);
            if (unitId == 0)
            {
                emit(label.Length > 0 ? "单元不唯一，" + label + "，请用 _ZGS_编号 或 总概算序号。" : "没找到单元：" + unitToken);
                return;
            }

            AgentBaseline baseline = new AgentBaseline();
            baseline.Time = DateTime.Now;
            baseline.Database = conn.Database;
            baseline.UnitId = unitId;
            baseline.RowCounts = LoadAgentTableRowCounts(conn);

            Dictionary<string, string> specs = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (string[] watched in AgentUnitWatchedTables)
            {
                if (!baseline.RowCounts.ContainsKey(watched[0]))
                {
                    continue;
                }

                string keySpec = watched[1];
                if (String.IsNullOrEmpty(keySpec))
                {
                    keySpec = LoadAgentPrimaryKeySpec(conn, watched[0]);
                }

                if (String.IsNullOrEmpty(keySpec))
                {
                    emit("  表 " + watched[0] + " 没有主键，跳过。");
                    continue;
                }

                specs[watched[0]] = keySpec;
            }

            baseline.WatchedSpecs = specs;
            foreach (KeyValuePair<string, string> watched in specs)
            {
                long count;
                baseline.RowCounts.TryGetValue(watched.Key, out count);
                Dictionary<string, Dictionary<string, object>> rows;
                try
                {
                    rows = LoadAgentWatchedRows(conn, watched.Key, watched.Value, count, unitId);
                }
                catch (Exception ex)
                {
                    emit("  表 " + watched.Key + " 快照失败（" + ex.Message + "），跳过。");
                    continue;
                }

                if (rows == null)
                {
                    emit("  注意：表 " + watched.Key + " 即使按单元过滤后仍超过 " + AgentWatchedRowCap.ToString(CultureInfo.InvariantCulture) + " 行，跳过。");
                }
                else
                {
                    baseline.WatchedRows[watched.Key] = rows;
                }
            }

            AgentBaselines[mainForm] = baseline;
            emit("已按单元 " + label + " 记录基线（行级快照 " + baseline.WatchedRows.Count.ToString(CultureInfo.InvariantCulture) +
                " 张表，含 " + String.Join("、", baseline.WatchedRows.Keys.ToArray()) + "）。");
            emit("现在请在软件里对【这个单元】的条目手工设置运输方案/材料方案（并触发重算），完成后输入：探查 对比");
        }

        private static string FormatAgentValue(object value)
        {
            if (value == null)
            {
                return "<null>";
            }

            string text = Convert.ToString(value, CultureInfo.InvariantCulture);
            return text.Length > 80 ? text.Substring(0, 80) + "…" : text;
        }

        private static void AgentDiagCompare(Form mainForm, Action<string> emit)
        {
            AgentBaseline baseline;
            if (!AgentBaselines.TryGetValue(mainForm, out baseline))
            {
                emit("还没有记录基线，请先输入：探查 基线");
                return;
            }

            SqlConnection conn = RequireAgentConnection(mainForm);
            if (!String.Equals(conn.Database, baseline.Database, StringComparison.OrdinalIgnoreCase))
            {
                emit("当前项目库(" + conn.Database + ")与基线库(" + baseline.Database + ")不一致，对比无意义。请重新做基线。");
                return;
            }

            emit("与基线（" + baseline.Time.ToString("HH:mm:ss", CultureInfo.InvariantCulture) + "）对比：");
            Dictionary<string, long> current = LoadAgentTableRowCounts(conn);
            int changedTables = 0;
            foreach (KeyValuePair<string, long> pair in current.OrderBy(p => p.Key, StringComparer.Ordinal))
            {
                long before;
                baseline.RowCounts.TryGetValue(pair.Key, out before);
                if (pair.Value != before)
                {
                    changedTables++;
                    emit("  [行数变化] " + pair.Key + ": " + before.ToString(CultureInfo.InvariantCulture) + " -> " + pair.Value.ToString(CultureInfo.InvariantCulture));
                }
            }

            if (changedTables == 0)
            {
                emit("  没有任何表的行数发生变化。");
            }

            foreach (KeyValuePair<string, Dictionary<string, Dictionary<string, object>>> watched in baseline.WatchedRows)
            {
                string table = watched.Key;
                long count;
                current.TryGetValue(table, out count);
                string keySpec;
                if (!baseline.WatchedSpecs.TryGetValue(table, out keySpec))
                {
                    keySpec = AgentWatchedTables.First(w => w[0] == table)[1];
                }

                Dictionary<string, Dictionary<string, object>> nowRows = LoadAgentWatchedRows(conn, table, keySpec, count, baseline.UnitId);
                if (nowRows == null)
                {
                    emit("  表 " + table + " 当前行数过大，跳过行级diff。");
                    continue;
                }

                List<string> added = nowRows.Keys.Where(k => !watched.Value.ContainsKey(k)).ToList();
                List<string> removed = watched.Value.Keys.Where(k => !nowRows.ContainsKey(k)).ToList();
                int changedRows = 0;
                int emitted = 0;
                foreach (string key in added)
                {
                    if (emitted < 30)
                    {
                        Dictionary<string, object> row = nowRows[key];
                        string detail = String.Join("; ", row
                            .Where(p => p.Value != null && Convert.ToString(p.Value, CultureInfo.InvariantCulture).Trim().Length > 0)
                            .Select(p => p.Key + "=" + FormatAgentValue(p.Value))
                            .ToArray());
                        emit("  [" + table + " 新增 " + key + "] " + detail);
                        emitted++;
                    }
                    else
                    {
                        AgentDiagLog("  [" + table + " 新增 " + key + "] " + String.Join("; ", nowRows[key].Select(p => p.Key + "=" + FormatAgentValue(p.Value)).ToArray()));
                    }
                }

                foreach (string key in removed)
                {
                    emit("  [" + table + " 删除] 主键 " + key);
                }

                foreach (KeyValuePair<string, Dictionary<string, object>> pair in watched.Value)
                {
                    Dictionary<string, object> nowRow;
                    if (!nowRows.TryGetValue(pair.Key, out nowRow))
                    {
                        continue;
                    }

                    List<string> diffs = new List<string>();
                    foreach (KeyValuePair<string, object> col in pair.Value)
                    {
                        object nowValue;
                        nowRow.TryGetValue(col.Key, out nowValue);
                        string oldText = col.Value == null ? "<null>" : Convert.ToString(col.Value, CultureInfo.InvariantCulture);
                        string newText = nowValue == null ? "<null>" : Convert.ToString(nowValue, CultureInfo.InvariantCulture);
                        if (!String.Equals(oldText, newText, StringComparison.Ordinal))
                        {
                            diffs.Add(col.Key + ": " + FormatAgentValue(col.Value) + " -> " + FormatAgentValue(nowValue));
                        }
                    }

                    if (diffs.Count > 0)
                    {
                        changedRows++;
                        if (changedRows <= 30)
                        {
                            emit("  [" + table + " 修改 " + pair.Key + "] " + String.Join("; ", diffs.ToArray()));
                        }
                        else
                        {
                            AgentDiagLog("  [" + table + " 修改 " + pair.Key + "] " + String.Join("; ", diffs.ToArray()));
                        }
                    }
                }

                if (added.Count > 0 || removed.Count > 0 || changedRows > 0)
                {
                    emit("  表 " + table + " 合计：新增 " + added.Count.ToString(CultureInfo.InvariantCulture) +
                        "，删除 " + removed.Count.ToString(CultureInfo.InvariantCulture) +
                        "，修改 " + changedRows.ToString(CultureInfo.InvariantCulture) + "（超出30条的明细在 agent-diagnostics.log）");
                }
            }

            emit("对比完成。完整输出见 RecoQuotaData/agent-diagnostics.log");
        }

        private static void AgentDiagMainForm(Form mainForm, Action<string> emit)
        {
            Type type = mainForm.GetType();
            Regex interesting = new Regex("粘贴|Paste|新建|单项|概算|材料|价差|价格|保存|Save|插入|Insert|删除|Delete|刷新|Refresh", RegexOptions.IgnoreCase);

            AgentDiagLog("===== 窗体字段 dump：" + type.FullName + " =====");
            int fieldCount = 0;
            foreach (FieldInfo field in type.GetFields(BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic))
            {
                fieldCount++;
                AgentDiagLog("FIELD " + field.Name + " : " + field.FieldType.FullName);
            }

            emit("窗体字段共 " + fieldCount.ToString(CultureInfo.InvariantCulture) + " 个，已全部写入 agent-diagnostics.log。");

            emit("名称疑似相关的方法：");
            int matched = 0;
            foreach (MethodInfo method in type.GetMethods(BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic))
            {
                if (!interesting.IsMatch(method.Name))
                {
                    continue;
                }

                matched++;
                string signature = method.Name + "(" + String.Join(", ", method.GetParameters().Select(p => p.ParameterType.Name).ToArray()) + ")";
                AgentDiagLog("METHOD " + signature);
                if (matched <= 40)
                {
                    emit("  " + signature);
                }
            }

            if (matched > 40)
            {
                emit("  …共 " + matched.ToString(CultureInfo.InvariantCulture) + " 个，其余见日志。");
            }

            ContextMenuStrip deMenu = GetField<ContextMenuStrip>(mainForm, "contextMenuStripDE");
            if (deMenu != null)
            {
                emit("定额表右键菜单项：");
                foreach (ToolStripItem item in deMenu.Items)
                {
                    string text = String.IsNullOrEmpty(item.Text) ? "<分隔线>" : item.Text;
                    emit("  " + text);
                    AgentDiagLog("DEMENU " + text);
                }
            }
        }
    }
}
