using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        // 学习库固定在中央服务器,不跟随各软件的项目服务器(2020版连192.168.2.13,学习库统一在.213)。
        private const string LearningDbServer = "192.168.2.213,1433";
        private const string LearningDbOutboxFileName = "learning-db-outbox.jsonl";
        private const string LearningDbOutboxMutexName = "RecoQuotaData.learning-db-outbox.lock";

        private static string learningDbConnectionString;
        [ThreadStatic]
        private static List<MappingFeedbackGroup> lastLearningDbResultGroups;
        [ThreadStatic]
        private static bool lastLearningDbResultDurable;

        private sealed class LearningDbOutboxBatch
        {
            public string BatchId;
            public string Source;
            public List<MappingFeedbackGroup> Groups = new List<MappingFeedbackGroup>();
        }

        // RecoLearning 是多人共享主学习库：流水与推荐核心聚合在同一事务提交；
        // 瞬时并发错误整笔重试一次；其他失败只记日志，不影响后续绑定或本机 jsonl 备份。
        private static bool RecordBindingEventsToLearningDb(string source, List<MappingFeedbackGroup> groups)
        {
            RememberLearningDbDurableResult(source, groups, false);
            if (groups == null || groups.Count == 0)
            {
                return false;
            }

            PrepareLearningGroupsForOutbox(groups);
            if (!TryReplayPendingLearningDbEvents())
            {
                LearningDbOutboxBatch pending = new LearningDbOutboxBatch
                {
                    BatchId = Guid.NewGuid().ToString("N"),
                    Source = source ?? "",
                    Groups = groups
                };
                bool queued = TryAppendLearningDbOutbox(pending);
                RememberLearningDbDurableResult(source, groups, queued);
                return queued;
            }

            LearningDbOutboxBatch batch = new LearningDbOutboxBatch
            {
                BatchId = Guid.NewGuid().ToString("N"),
                Source = source ?? "",
                Groups = groups
            };
            bool written = TryWriteLearningDbBatch(batch, true);
            bool durable = written || TryAppendLearningDbOutbox(batch);
            RememberLearningDbDurableResult(source, groups, durable);
            return durable;
        }

        private static void WriteBindingEventsToLearningDb(string connectionString, string source,
            List<MappingFeedbackGroup> groups, string batchId)
        {
            if (String.IsNullOrEmpty(connectionString)) return;
            if ((groups ?? new List<MappingFeedbackGroup>()).Any(group => group != null &&
                !String.IsNullOrWhiteSpace(group.QuantityName) && String.IsNullOrEmpty(NormalizeLearningDbMethod(group.Method))))
            {
                throw new InvalidOperationException("Learning DB method is missing or unsupported; binding remains queued locally.");
            }

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                using (SqlTransaction transaction = conn.BeginTransaction())
                {
                    try
                    {
                        for (int groupIndex = 0; groupIndex < groups.Count; groupIndex++)
                        {
                            MappingFeedbackGroup group = groups[groupIndex];
                            if (group == null || String.IsNullOrWhiteSpace(group.QuantityName))
                            {
                                continue;
                            }

                            string groupKey = BuildLearningMd5((batchId ?? "") + "|group|" + groupIndex.ToString(CultureInfo.InvariantCulture));
                            for (int targetIndex = 0; targetIndex < group.Targets.Count; targetIndex++)
                            {
                                MappingFeedbackTarget target = group.Targets[targetIndex];
                                if (target == null || String.IsNullOrWhiteSpace(target.Code))
                                {
                                    continue;
                                }

                                using (SqlCommand cmd = conn.CreateCommand())
                                {
                                    cmd.Transaction = transaction;
                                    cmd.CommandTimeout = 5;
                                    cmd.CommandText = "INSERT INTO dbo.BindingLog (occurred_at, source, method, project_id, entry_code, entry_name, quantity_name, quantity_unit, target_kind, target_code, target_name, target_unit, group_key, event_hash, extra) " +
                                        "VALUES (SYSDATETIME(), @source, @method, @project, @ec, @en, @qn, @qu, @tk, @tc, @tn, @tu, @gk, @eh, @ex)";
                                    cmd.Parameters.AddWithValue("@source", "plugin:" + (source ?? ""));
                                    cmd.Parameters.AddWithValue("@method", NormalizeLearningDbMethod(group.Method));
                                    cmd.Parameters.AddWithValue("@project", group.ProjectId ?? "");
                                    cmd.Parameters.AddWithValue("@ec", group.EntryCode ?? "");
                                    cmd.Parameters.AddWithValue("@en", group.EntryName ?? "");
                                    cmd.Parameters.AddWithValue("@qn", group.QuantityName ?? "");
                                    cmd.Parameters.AddWithValue("@qu", group.QuantityUnit ?? "");
                                    cmd.Parameters.AddWithValue("@tk", String.IsNullOrEmpty(target.Kind) ? "quota" : target.Kind);
                                    cmd.Parameters.AddWithValue("@tc", target.Code ?? "");
                                    cmd.Parameters.AddWithValue("@tn", target.Name ?? "");
                                    cmd.Parameters.AddWithValue("@tu", target.Unit ?? "");
                                    cmd.Parameters.AddWithValue("@gk", groupKey);
                                    cmd.Parameters.AddWithValue("@eh", BuildLearningMd5((batchId ?? "") + "|group|" +
                                        groupIndex.ToString(CultureInfo.InvariantCulture) + "|target|" + targetIndex.ToString(CultureInfo.InvariantCulture)));
                                    Dictionary<string, string> flat = new Dictionary<string, string>();
                                    if (!String.IsNullOrEmpty(group.Workbook)) flat["workbook"] = group.Workbook;
                                    if (!String.IsNullOrEmpty(group.Worksheet)) flat["worksheet"] = group.Worksheet;
                                    if (group.ExcelRow > 0) flat["excel_row"] = group.ExcelRow.ToString();
                                    if (!String.IsNullOrEmpty(group.BoxId)) flat["box_id"] = group.BoxId;
                                    if (!String.IsNullOrEmpty(group.Expression)) flat["expression"] = group.Expression;
                                    if (!String.IsNullOrEmpty(group.SourceCell)) flat["source_cell"] = group.SourceCell;
                                    if (!String.IsNullOrWhiteSpace(target.FormulaTemplate) && group.FormulaOperands.Count > 0)
                                    {
                                        flat["formula_rule_hash"] = BuildLearningFormulaRuleHash(group, target);
                                        flat["formula_template"] = target.FormulaTemplate;
                                        flat["formula_target_unit"] = target.Unit ?? "";
                                        flat["formula_operand_count"] = group.FormulaOperands.Count.ToString(CultureInfo.InvariantCulture);
                                        for (int operandIndex = 0; operandIndex < group.FormulaOperands.Count; operandIndex++)
                                        {
                                            QuantityFormulaOperandInfo operand = group.FormulaOperands[operandIndex];
                                            string prefix = "formula_operand_" + operandIndex.ToString(CultureInfo.InvariantCulture) + "_";
                                            flat[prefix + "name"] = operand.Name ?? "";
                                            flat[prefix + "unit"] = operand.Unit ?? "";
                                            flat[prefix + "signature"] = operand.Signature ?? "";
                                        }
                                    }
                                    string extra = flat.Count == 0 ? "" : ToFlatJson(flat);
                                    if (extra.Length > 0) cmd.Parameters.AddWithValue("@ex", extra);
                                    else cmd.Parameters.AddWithValue("@ex", DBNull.Value);
                                    cmd.ExecuteNonQuery();
                                }
                            }

                            UpsertBindingGroupAggregates(conn, transaction, group);
                        }
                        transaction.Commit();
                    }
                    catch
                    {
                        try { transaction.Rollback(); } catch { }
                        throw;
                    }
                }
            }
        }

        private static bool TryWriteLearningDbBatch(LearningDbOutboxBatch batch, bool logFailure)
        {
            if (batch == null || batch.Groups == null || batch.Groups.Count == 0) return true;
            string connectionString = GetLearningDbConnectionString();
            if (String.IsNullOrEmpty(connectionString)) return false;

            for (int attempt = 0; attempt < 2; attempt++)
            {
                try
                {
                    WriteBindingEventsToLearningDb(connectionString, batch.Source, batch.Groups, batch.BatchId);
                    return true;
                }
                catch (SqlException ex)
                {
                    if (IsUniqueLearningDbConflict(ex) && IsLearningDbBatchAlreadyCommitted(connectionString, batch)) return true;
                    if (attempt == 0 && IsRetryableLearningDbException(ex))
                    {
                        Log("Learning DB transient write failure; retrying once: " + ex.Message);
                        Thread.Sleep(100);
                        continue;
                    }
                    if (logFailure) Log("Learning DB write queued for retry: " + ex.Message);
                    return false;
                }
                catch (Exception ex)
                {
                    if (logFailure) Log("Learning DB write queued for retry: " + ex.Message);
                    return false;
                }
            }
            return false;
        }

        private static bool IsUniqueLearningDbConflict(SqlException ex)
        {
            return ex != null && ex.Errors.Cast<SqlError>().Any(error => error.Number == 2601 || error.Number == 2627);
        }

        private static bool IsLearningDbBatchAlreadyCommitted(string connectionString, LearningDbOutboxBatch batch)
        {
            if (batch == null || batch.Groups == null || String.IsNullOrWhiteSpace(batch.BatchId)) return false;
            try
            {
                bool hasEvents = false;
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    for (int groupIndex = 0; groupIndex < batch.Groups.Count; groupIndex++)
                    {
                        MappingFeedbackGroup group = batch.Groups[groupIndex];
                        int expected = group == null || group.Targets == null ? 0 :
                            group.Targets.Count(target => target != null && !String.IsNullOrWhiteSpace(target.Code));
                        if (expected == 0) continue;
                        hasEvents = true;
                        using (SqlCommand cmd = conn.CreateCommand())
                        {
                            cmd.CommandTimeout = 5;
                            cmd.CommandText = "SELECT COUNT(*) FROM dbo.BindingLog WHERE group_key=@group_key";
                            cmd.Parameters.AddWithValue("@group_key", BuildLearningMd5(batch.BatchId + "|group|" +
                                groupIndex.ToString(CultureInfo.InvariantCulture)));
                            if (Convert.ToInt32(cmd.ExecuteScalar(), CultureInfo.InvariantCulture) != expected) return false;
                        }
                    }
                }
                return hasEvents;
            }
            catch
            {
                return false;
            }
        }

        private static void PrepareLearningGroupsForOutbox(List<MappingFeedbackGroup> groups)
        {
            foreach (MappingFeedbackGroup group in groups ?? new List<MappingFeedbackGroup>())
            {
                if (group == null) continue;
                group.Method = NormalizeLearningDbMethod(group.Method);
                if (!String.IsNullOrWhiteSpace(group.BoxId)) continue;
                List<string> targetKeys = (group.Targets ?? new List<MappingFeedbackTarget>())
                    .Where(target => target != null && !String.IsNullOrWhiteSpace(target.Code))
                    .Select(target => BuildMappingTargetKey(target.Kind, target.Code))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(key => key, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (targetKeys.Count > 0) group.BoxId = BuildStableMappingBoxId(String.Join("|", targetKeys.ToArray()));
            }
        }

        private static bool TryAppendLearningDbOutbox(LearningDbOutboxBatch batch)
        {
            if (batch == null || String.IsNullOrWhiteSpace(batch.BatchId) || batch.Groups == null || batch.Groups.Count == 0) return false;
            try
            {
                bool durable = false;
                bool locked = TryWithLearningDbOutboxLock(delegate
                {
                    string path = GetLearningDbOutboxPath();
                    Directory.CreateDirectory(Path.GetDirectoryName(path));
                    List<string> lines = File.Exists(path)
                        ? File.ReadAllLines(path, Encoding.UTF8).ToList()
                        : new List<string>();
                    bool exists = lines.Any(line => String.Equals(GetFlat(ParseFlatJson(line), "batch_id"), batch.BatchId,
                        StringComparison.OrdinalIgnoreCase));
                    if (!exists)
                    {
                        lines.Add(ToFlatJson(SerializeLearningDbOutboxBatch(batch)));
                        WriteAllLinesAtomic(path, lines.ToArray(), Encoding.UTF8);
                    }
                    durable = true;
                }, 5000);
                if (!locked) Log("Learning DB outbox lock timeout; pending SQL event was not persisted.");
                return locked && durable;
            }
            catch (Exception ex)
            {
                Log("Learning DB outbox append failed: " + ex.Message);
                return false;
            }
        }

        private static Dictionary<string, string> SerializeLearningDbOutboxBatch(LearningDbOutboxBatch batch)
        {
            Dictionary<string, string> row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            row["record_type"] = "learning_db_outbox";
            row["batch_id"] = batch.BatchId ?? "";
            row["source"] = batch.Source ?? "";
            row["group_count"] = batch.Groups.Count.ToString(CultureInfo.InvariantCulture);
            for (int groupIndex = 0; groupIndex < batch.Groups.Count; groupIndex++)
            {
                MappingFeedbackGroup group = batch.Groups[groupIndex] ?? new MappingFeedbackGroup();
                string prefix = "g" + groupIndex.ToString(CultureInfo.InvariantCulture) + "_";
                row[prefix + "quantity_name"] = group.QuantityName ?? "";
                row[prefix + "quantity_unit"] = group.QuantityUnit ?? "";
                row[prefix + "method"] = NormalizeLearningDbMethod(group.Method);
                row[prefix + "project_id"] = group.ProjectId ?? "";
                row[prefix + "entry_code"] = group.EntryCode ?? "";
                row[prefix + "entry_name"] = group.EntryName ?? "";
                row[prefix + "workbook"] = group.Workbook ?? "";
                row[prefix + "worksheet"] = group.Worksheet ?? "";
                row[prefix + "excel_row"] = group.ExcelRow.ToString(CultureInfo.InvariantCulture);
                row[prefix + "box_id"] = group.BoxId ?? "";
                row[prefix + "expression"] = group.Expression ?? "";
                row[prefix + "source_cell"] = group.SourceCell ?? "";
                row[prefix + "target_count"] = (group.Targets == null ? 0 : group.Targets.Count).ToString(CultureInfo.InvariantCulture);
                for (int targetIndex = 0; targetIndex < (group.Targets == null ? 0 : group.Targets.Count); targetIndex++)
                {
                    MappingFeedbackTarget target = group.Targets[targetIndex] ?? new MappingFeedbackTarget();
                    string targetPrefix = prefix + "t" + targetIndex.ToString(CultureInfo.InvariantCulture) + "_";
                    row[targetPrefix + "kind"] = target.Kind ?? "";
                    row[targetPrefix + "code"] = target.Code ?? "";
                    row[targetPrefix + "name"] = target.Name ?? "";
                    row[targetPrefix + "unit"] = target.Unit ?? "";
                    row[targetPrefix + "formula"] = target.FormulaTemplate ?? "";
                }
                row[prefix + "operand_count"] = (group.FormulaOperands == null ? 0 : group.FormulaOperands.Count).ToString(CultureInfo.InvariantCulture);
                for (int operandIndex = 0; operandIndex < (group.FormulaOperands == null ? 0 : group.FormulaOperands.Count); operandIndex++)
                {
                    QuantityFormulaOperandInfo operand = group.FormulaOperands[operandIndex] ?? new QuantityFormulaOperandInfo();
                    string operandPrefix = prefix + "o" + operandIndex.ToString(CultureInfo.InvariantCulture) + "_";
                    row[operandPrefix + "name"] = operand.Name ?? "";
                    row[operandPrefix + "unit"] = operand.Unit ?? "";
                    row[operandPrefix + "signature"] = operand.Signature ?? "";
                }
            }
            return row;
        }

        private static LearningDbOutboxBatch ParseLearningDbOutboxBatch(string line)
        {
            Dictionary<string, string> row = ParseFlatJson(line);
            string batchId = GetFlat(row, "batch_id").Trim();
            int groupCount = ReadFlatInt(row, "group_count", 0);
            if (batchId.Length != 32 || !batchId.All(Uri.IsHexDigit) || groupCount <= 0 || groupCount > 10000) return null;
            LearningDbOutboxBatch batch = new LearningDbOutboxBatch { BatchId = batchId, Source = GetFlat(row, "source") };
            for (int groupIndex = 0; groupIndex < groupCount; groupIndex++)
            {
                string prefix = "g" + groupIndex.ToString(CultureInfo.InvariantCulture) + "_";
                MappingFeedbackGroup group = new MappingFeedbackGroup
                {
                    QuantityName = GetFlat(row, prefix + "quantity_name"),
                    QuantityUnit = GetFlat(row, prefix + "quantity_unit"),
                    Method = NormalizeLearningDbMethod(GetFlat(row, prefix + "method")),
                    ProjectId = GetFlat(row, prefix + "project_id"),
                    EntryCode = GetFlat(row, prefix + "entry_code"),
                    EntryName = GetFlat(row, prefix + "entry_name"),
                    Workbook = GetFlat(row, prefix + "workbook"),
                    Worksheet = GetFlat(row, prefix + "worksheet"),
                    ExcelRow = ReadFlatInt(row, prefix + "excel_row", 0),
                    BoxId = GetFlat(row, prefix + "box_id"),
                    Expression = GetFlat(row, prefix + "expression"),
                    SourceCell = GetFlat(row, prefix + "source_cell")
                };
                int targetCount = ReadFlatInt(row, prefix + "target_count", 0);
                int operandCount = ReadFlatInt(row, prefix + "operand_count", 0);
                if (targetCount < 0 || targetCount > 10000 || operandCount < 0 || operandCount > 10000) return null;
                for (int targetIndex = 0; targetIndex < targetCount; targetIndex++)
                {
                    string targetPrefix = prefix + "t" + targetIndex.ToString(CultureInfo.InvariantCulture) + "_";
                    group.Targets.Add(new MappingFeedbackTarget
                    {
                        Kind = GetFlat(row, targetPrefix + "kind"),
                        Code = GetFlat(row, targetPrefix + "code"),
                        Name = GetFlat(row, targetPrefix + "name"),
                        Unit = GetFlat(row, targetPrefix + "unit"),
                        FormulaTemplate = GetFlat(row, targetPrefix + "formula")
                    });
                }
                for (int operandIndex = 0; operandIndex < operandCount; operandIndex++)
                {
                    string operandPrefix = prefix + "o" + operandIndex.ToString(CultureInfo.InvariantCulture) + "_";
                    group.FormulaOperands.Add(new QuantityFormulaOperandInfo
                    {
                        Name = GetFlat(row, operandPrefix + "name"),
                        Unit = GetFlat(row, operandPrefix + "unit"),
                        Signature = GetFlat(row, operandPrefix + "signature")
                    });
                }
                batch.Groups.Add(group);
            }
            return batch;
        }

        private static bool TryReplayPendingLearningDbEvents()
        {
            List<LearningDbOutboxBatch> batches = LoadLearningDbOutboxBatches();
            foreach (LearningDbOutboxBatch batch in batches.Take(20))
            {
                if (!TryWriteLearningDbBatch(batch, false)) return false;
                RemoveLearningDbOutboxBatch(batch.BatchId);
            }
            return true;
        }

        // 推荐预览读取 SQL 前调用，确保此前离线绑定优先补传到共享学习库。
        private static void ReplayPendingLearningDbEvents()
        {
            TryReplayPendingLearningDbEvents();
        }

        // 返回“名称级签名\n组件框”键，供推荐层只叠加尚未确认写入 SQL 的本机关系。
        private static HashSet<string> LoadPendingLearningMappingKeys()
        {
            HashSet<string> keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (LearningDbOutboxBatch batch in LoadLearningDbOutboxBatches())
            {
                foreach (MappingFeedbackGroup group in batch.Groups)
                {
                    if (group == null || String.IsNullOrWhiteSpace(group.QuantityName) || String.IsNullOrWhiteSpace(group.BoxId)) continue;
                    string signature = NormalizeForSignature(group.QuantityName) + "|";
                    if (signature.Length > 450) signature = signature.Substring(0, 450);
                    keys.Add(signature + "\n" + group.BoxId);
                }
            }
            return keys;
        }

        private static List<LearningDbOutboxBatch> LoadLearningDbOutboxBatches()
        {
            List<LearningDbOutboxBatch> batches = new List<LearningDbOutboxBatch>();
            try
            {
                TryWithLearningDbOutboxLock(delegate
                {
                    string path = GetLearningDbOutboxPath();
                    if (!File.Exists(path)) return;
                    HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
                    {
                        LearningDbOutboxBatch batch = ParseLearningDbOutboxBatch(line);
                        if (batch != null && seen.Add(batch.BatchId)) batches.Add(batch);
                    }
                }, 1000);
            }
            catch (Exception ex)
            {
                Log("Learning DB outbox load failed: " + ex.Message);
            }
            return batches;
        }

        private static void RemoveLearningDbOutboxBatch(string batchId)
        {
            if (String.IsNullOrWhiteSpace(batchId)) return;
            try
            {
                if (!TryWithLearningDbOutboxLock(delegate
                {
                    string path = GetLearningDbOutboxPath();
                    if (!File.Exists(path)) return;
                    List<string> remaining = File.ReadAllLines(path, Encoding.UTF8)
                        .Where(line => !String.Equals(GetFlat(ParseFlatJson(line), "batch_id"), batchId, StringComparison.OrdinalIgnoreCase))
                        .ToList();
                    WriteAllLinesAtomic(path, remaining.ToArray(), Encoding.UTF8);
                }, 5000))
                {
                    Log("Learning DB outbox cleanup lock timeout; committed event remains for idempotent replay.");
                }
            }
            catch (Exception ex)
            {
                Log("Learning DB outbox cleanup failed; committed event remains for idempotent replay: " + ex.Message);
            }
        }

        private static bool TryWithLearningDbOutboxLock(Action action, int timeoutMilliseconds)
        {
            Mutex mutex = new Mutex(false, LearningDbOutboxMutexName);
            bool acquired = false;
            try
            {
                try { acquired = mutex.WaitOne(timeoutMilliseconds); }
                catch (AbandonedMutexException) { acquired = true; }
                if (!acquired) return false;
                action();
                return true;
            }
            finally
            {
                if (acquired) mutex.ReleaseMutex();
                mutex.Dispose();
            }
        }

        private static string GetLearningDbOutboxPath()
        {
            return Path.Combine(FindRecoQuotaDataDir(), LearningDbOutboxFileName);
        }

        private static void RememberLearningDbDurableResult(string source, List<MappingFeedbackGroup> groups, bool durable)
        {
            if (!String.Equals(source ?? "", "name-match", StringComparison.OrdinalIgnoreCase)) return;
            lastLearningDbResultGroups = groups;
            lastLearningDbResultDurable = durable;
        }

        private static bool ConsumeLearningDbDurableResult(List<MappingFeedbackGroup> groups)
        {
            bool durable = Object.ReferenceEquals(lastLearningDbResultGroups, groups) && lastLearningDbResultDurable;
            lastLearningDbResultGroups = null;
            lastLearningDbResultDurable = false;
            return durable;
        }

        private static bool IsRetryableLearningDbException(SqlException ex)
        {
            if (ex == null) return false;
            foreach (SqlError error in ex.Errors)
            {
                if (error.Number == -2 || error.Number == 53 || error.Number == 64 || error.Number == 233 ||
                    error.Number == 1205 || error.Number == 2601 || error.Number == 2627 ||
                    error.Number == 10053 || error.Number == 10054 || error.Number == 10060)
                {
                    return true;
                }
            }
            return false;
        }

        // 绑定流水是审计源；同时增量维护推荐预览直接读取的核心聚合表，避免等待全量收割。
        private static void UpsertBindingGroupAggregates(SqlConnection conn, SqlTransaction transaction, MappingFeedbackGroup group)
        {
            List<MappingFeedbackTarget> targets = (group.Targets ?? new List<MappingFeedbackTarget>())
                .Where(target => target != null && !String.IsNullOrWhiteSpace(target.Code))
                .GroupBy(target => BuildMappingTargetKey(target.Kind, target.Code), StringComparer.OrdinalIgnoreCase)
                .Select(items => items.First())
                .OrderBy(target => BuildMappingTargetKey(target.Kind, target.Code), StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (targets.Count == 0) return;

            string name = TrimLearningText(group.QuantityName, 1000);
            string unit = TrimLearningText(group.QuantityUnit, 50);
            string signature = NormalizeForSignature(name) + "|";
            if (signature.Length > 450) signature = signature.Substring(0, 450);
            string aliasHash = BuildLearningMd5(NormalizeForSignature(name));
            List<string> targetKeys = targets.Select(target => BuildMappingTargetKey(target.Kind, target.Code)).ToList();
            string targetSetHash = BuildLearningMd5(String.Join(";", targetKeys.ToArray()));
            string boxId = TrimLearningText(group.BoxId, 64);
            if (String.IsNullOrWhiteSpace(boxId)) boxId = BuildStableMappingBoxId(String.Join("|", targetKeys.ToArray()));

            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandTimeout = 5;
                cmd.CommandText =
                    "UPDATE dbo.QuantityAlias WITH (UPDLOCK,HOLDLOCK) SET raw_name=@n, quantity_unit=@u, signature=@s, seen_count=seen_count+1, last_seen=SYSDATETIME() WHERE alias_hash=@h; " +
                    "IF @@ROWCOUNT=0 INSERT INTO dbo.QuantityAlias(alias_hash,raw_name,quantity_unit,signature,seen_count,first_seen,last_seen) VALUES(@h,@n,@u,@s,1,SYSDATETIME(),SYSDATETIME());";
                cmd.Parameters.AddWithValue("@n", name);
                cmd.Parameters.AddWithValue("@u", unit);
                cmd.Parameters.AddWithValue("@s", signature);
                cmd.Parameters.AddWithValue("@h", aliasHash);
                cmd.ExecuteNonQuery();
            }

            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandTimeout = 5;
                cmd.CommandText = "SELECT box_id FROM dbo.QuotaBox WITH (UPDLOCK,HOLDLOCK) WHERE target_set_hash=@hash";
                cmd.Parameters.AddWithValue("@hash", targetSetHash);
                object existing = cmd.ExecuteScalar();
                if (existing != null && existing != DBNull.Value)
                {
                    boxId = Convert.ToString(existing);
                }
            }
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandTimeout = 5;
                cmd.CommandText = "IF NOT EXISTS(SELECT 1 FROM dbo.QuotaBox WITH (UPDLOCK,HOLDLOCK) WHERE box_id=@box) INSERT INTO dbo.QuotaBox(box_id,target_set_hash,status) VALUES(@box,@hash,'active')";
                cmd.Parameters.AddWithValue("@box", boxId);
                cmd.Parameters.AddWithValue("@hash", targetSetHash);
                cmd.ExecuteNonQuery();
            }

            foreach (MappingFeedbackTarget target in targets)
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    cmd.CommandTimeout = 5;
                    cmd.CommandText =
                        "UPDATE dbo.QuotaBoxTarget WITH (UPDLOCK,HOLDLOCK) SET " +
                        "target_name=CASE WHEN @name='' THEN target_name ELSE @name END, " +
                        "target_unit=CASE WHEN @unit='' THEN target_unit ELSE @unit END " +
                        "WHERE box_id=@box AND target_kind=@kind AND target_code=@code; " +
                        "IF @@ROWCOUNT=0 INSERT INTO dbo.QuotaBoxTarget(box_id,target_kind,target_code,target_name,target_unit) VALUES(@box,@kind,@code,@name,@unit);";
                    cmd.Parameters.AddWithValue("@box", boxId);
                    cmd.Parameters.AddWithValue("@kind", TrimLearningText(String.IsNullOrWhiteSpace(target.Kind) ? "quota" : target.Kind, 20));
                    cmd.Parameters.AddWithValue("@code", TrimLearningText(target.Code, 100));
                    cmd.Parameters.AddWithValue("@name", TrimLearningText(target.Name, 500));
                    cmd.Parameters.AddWithValue("@unit", TrimLearningText(target.Unit, 50));
                    cmd.ExecuteNonQuery();
                }
            }

            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandTimeout = 5;
                string method = NormalizeLearningDbMethod(group.Method);
                cmd.CommandText =
                    "UPDATE dbo.SignatureBoxMap WITH (UPDLOCK,HOLDLOCK) SET accepted_count=accepted_count+1, " +
                    "weight=CASE WHEN 10*(accepted_count+1)+20*corrected_count-10*rejected_count>100 THEN 100 " +
                    "WHEN 10*(accepted_count+1)+20*corrected_count-10*rejected_count<0 THEN 0 " +
                    "ELSE 10*(accepted_count+1)+20*corrected_count-10*rejected_count END, last_used_at=SYSDATETIME() " +
                    "WHERE signature=@s AND method=@method AND box_id=@box; " +
                    "IF @@ROWCOUNT=0 INSERT INTO dbo.SignatureBoxMap(signature,method,box_id,weight,accepted_count,corrected_count,rejected_count,last_used_at) VALUES(@s,@method,@box,10,1,0,0,SYSDATETIME());";
                cmd.Parameters.AddWithValue("@s", signature);
                cmd.Parameters.AddWithValue("@method", method);
                cmd.Parameters.AddWithValue("@box", boxId);
                cmd.ExecuteNonQuery();
            }

            bool hasFormulaTables = false;
            if (group.FormulaOperands.Count > 0 && targets.Any(target => !String.IsNullOrWhiteSpace(target.FormulaTemplate)))
            {
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.Transaction = transaction;
                    cmd.CommandTimeout = 5;
                    cmd.CommandText = "SELECT CASE WHEN OBJECT_ID('dbo.QuantityFormulaRule','U') IS NULL OR OBJECT_ID('dbo.QuantityFormulaOperand','U') IS NULL THEN 0 ELSE 1 END";
                    hasFormulaTables = Convert.ToInt32(cmd.ExecuteScalar(), CultureInfo.InvariantCulture) != 0;
                }
            }
            if (hasFormulaTables)
            {
                foreach (MappingFeedbackTarget target in targets)
                {
                    string formulaTemplate = TrimLearningText(target.FormulaTemplate, 2000);
                    if (formulaTemplate.Length == 0) continue;
                    string targetKind = TrimLearningText(String.IsNullOrWhiteSpace(target.Kind) ? "quota" : target.Kind, 20);
                    string targetCode = TrimLearningText(target.Code, 100);
                    string targetUnit = TrimLearningText(NormalizeExcelLinkUnit(target.Unit), 50);
                    string method = NormalizeLearningDbMethod(group.Method);
                    string formulaEntry = TrimLearningText(group.EntryCode, 100);
                    string ruleHash = BuildLearningFormulaRuleHash(group, target);
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandTimeout = 5;
                        cmd.CommandText =
                            "UPDATE dbo.QuantityFormulaRule WITH (UPDLOCK,HOLDLOCK) SET sample_count=sample_count+1,last_seen=SYSDATETIME() WHERE rule_hash=@hash; " +
                            "IF @@ROWCOUNT=0 INSERT INTO dbo.QuantityFormulaRule(rule_hash,anchor_signature,target_kind,target_code,target_unit,formula_template,method,entry_code,sample_count,first_seen,last_seen) " +
                            "VALUES(@hash,@s,@kind,@code,@unit,@formula,@method,@entry,1,SYSDATETIME(),SYSDATETIME());";
                        cmd.Parameters.AddWithValue("@hash", ruleHash);
                        cmd.Parameters.AddWithValue("@s", signature);
                        cmd.Parameters.AddWithValue("@kind", targetKind);
                        cmd.Parameters.AddWithValue("@code", targetCode);
                        cmd.Parameters.AddWithValue("@unit", targetUnit);
                        cmd.Parameters.AddWithValue("@formula", formulaTemplate);
                        cmd.Parameters.AddWithValue("@method", method);
                        cmd.Parameters.AddWithValue("@entry", formulaEntry);
                        cmd.ExecuteNonQuery();
                    }
                    for (int operandIndex = 0; operandIndex < group.FormulaOperands.Count; operandIndex++)
                    {
                        QuantityFormulaOperandInfo operand = group.FormulaOperands[operandIndex];
                        using (SqlCommand cmd = conn.CreateCommand())
                        {
                            cmd.Transaction = transaction;
                            cmd.CommandTimeout = 5;
                            cmd.CommandText =
                                "UPDATE dbo.QuantityFormulaOperand WITH (UPDLOCK,HOLDLOCK) SET " +
                                "operand_name=CASE WHEN @name='' THEN operand_name ELSE @name END,operand_unit=CASE WHEN @unit='' THEN operand_unit ELSE @unit END " +
                                "WHERE rule_hash=@hash AND operand_index=@idx; " +
                                "IF @@ROWCOUNT=0 INSERT INTO dbo.QuantityFormulaOperand(rule_hash,operand_index,operand_signature,operand_name,operand_unit) VALUES(@hash,@idx,@sig,@name,@unit);";
                            cmd.Parameters.AddWithValue("@hash", ruleHash);
                            cmd.Parameters.AddWithValue("@idx", operandIndex);
                            cmd.Parameters.AddWithValue("@sig", TrimLearningText(operand.Signature, 450));
                            cmd.Parameters.AddWithValue("@name", TrimLearningText(operand.Name, 1000));
                            cmd.Parameters.AddWithValue("@unit", TrimLearningText(NormalizeExcelLinkUnit(operand.Unit), 50));
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
            }

            string entryCode = TrimLearningText(group.EntryCode, 100);
            if (entryCode.Length > 0)
            {
                string method = NormalizeLearningDbMethod(group.Method);
                string entryName = TrimLearningText(group.EntryName, 500);
                foreach (MappingFeedbackTarget target in targets.Where(item => String.Equals(item.Kind ?? "quota", "quota", StringComparison.OrdinalIgnoreCase)))
                {
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandTimeout = 5;
                        cmd.CommandText =
                            "UPDATE dbo.SignatureEntryMap WITH (UPDLOCK,HOLDLOCK) SET entry_name=CASE WHEN @entry_name='' THEN entry_name ELSE @entry_name END, " +
                            "sample_count=sample_count+1,last_used_at=SYSDATETIME() WHERE signature=@s AND target_code=@code AND method=@method AND entry_code=@entry; " +
                            "IF @@ROWCOUNT=0 INSERT INTO dbo.SignatureEntryMap(signature,target_code,method,entry_code,entry_name,sample_count,last_used_at) VALUES(@s,@code,@method,@entry,@entry_name,1,SYSDATETIME());";
                        cmd.Parameters.AddWithValue("@s", signature);
                        cmd.Parameters.AddWithValue("@code", TrimLearningText(target.Code, 100));
                        cmd.Parameters.AddWithValue("@method", method);
                        cmd.Parameters.AddWithValue("@entry", entryCode);
                        cmd.Parameters.AddWithValue("@entry_name", entryName);
                        cmd.ExecuteNonQuery();
                    }
                }
            }
        }

        private static string TrimLearningText(string value, int maxLength)
        {
            string text = (value ?? "").Trim();
            return text.Length <= maxLength ? text : text.Substring(0, maxLength);
        }

        // 共享学习关系只按运行定额世代隔离。同一 2020 项目内的 30 号文预算和 101 号文估算可并存，
        // 其中公式规则再通过目标编号和稳定条目编号区分。空办法只供读取历史兼容，不允许新增写入。
        private static string NormalizeLearningDbMethod(string method)
        {
            string normalized = NormalizeSmartProjectMethod(method);
            if (String.Equals(normalized, "2024", StringComparison.OrdinalIgnoreCase)) return "2024";
            if (String.Equals(normalized, "2020", StringComparison.OrdinalIgnoreCase)) return "2020";
            return "";
        }

        private static string BuildLearningFormulaRuleHash(MappingFeedbackGroup group, MappingFeedbackTarget target)
        {
            string signature = NormalizeForSignature(group == null ? "" : group.QuantityName) + "|";
            if (signature.Length > 450) signature = signature.Substring(0, 450);
            string kind = target == null || String.IsNullOrWhiteSpace(target.Kind) ? "quota" : target.Kind.Trim().ToLowerInvariant();
            string raw = signature + "|" + kind + ":" + (target == null ? "" : target.Code ?? "").Trim().ToUpperInvariant() + "|" +
                NormalizeExcelLinkUnit(target == null ? "" : target.Unit) + "|" + (target == null ? "" : target.FormulaTemplate ?? "") + "|" +
                (group == null ? "" : NormalizeLearningDbMethod(group.Method)) + "|" + (group == null ? "" : group.EntryCode ?? "");
            if (group != null)
            {
                foreach (QuantityFormulaOperandInfo operand in group.FormulaOperands)
                {
                    raw += "|" + (operand.Signature ?? "") + "@" + NormalizeExcelLinkUnit(operand.Unit);
                }
            }
            return BuildLearningMd5(raw);
        }

        private static string BuildLearningMd5(string value)
        {
            using (MD5 md5 = MD5.Create())
            {
                byte[] hash = md5.ComputeHash(Encoding.UTF8.GetBytes(value ?? ""));
                StringBuilder builder = new StringBuilder(32);
                foreach (byte item in hash) builder.Append(item.ToString("x2"));
                return builder.ToString();
            }
        }

        private static string GetLearningDbConnectionString()
        {
            if (learningDbConnectionString == null)
            {
                learningDbConnectionString = "Server=" + LearningDbServer + ";Database=RecoLearning;User ID=" + AgentDbUser + ";Password=" + AgentDbPassword + ";Connect Timeout=3";
            }

            return learningDbConnectionString;
        }
    }
}
