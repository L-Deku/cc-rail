using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        // 学习库连接只从当前 Windows 用户的 DPAPI 凭据库读取。
        private const string LearningDbOutboxFileName = "learning-db-outbox.jsonl";
        private const string LearningDbDeadLetterFileName = "learning-db-outbox.dead-letter.jsonl";
        private const string LearningDbOutboxMutexName = "RecoQuotaData.learning-db-outbox.lock";
        private const long LearningDbCircuitWindowTicks = TimeSpan.TicksPerSecond * 60;

        private static string learningDbConnectionString;
        private static long learningDbCircuitOpenUntilTicks;
        private static long learningDbConnectionOpenAttempts;
        private static long learningDbRetrySleepCount;
        [ThreadStatic]
        private static List<MappingFeedbackGroup> lastLearningDbResultGroups;
        [ThreadStatic]
        private static bool lastLearningDbResultDurable;

        private sealed class LearningDbOutboxBatch
        {
            public string BatchId;
            public string Source;
            public string SoftwarePartition;
            public string ProcessName;
            public string ModuleFileName;
            public List<MappingFeedbackGroup> Groups = new List<MappingFeedbackGroup>();
        }

        private enum LearningDbWriteResult
        {
            Succeeded,
            RetryableFailure,
            PermanentFailure
        }

        private static bool IsLearningDbCircuitOpen()
        {
            return Interlocked.Read(ref learningDbCircuitOpenUntilTicks) > DateTime.UtcNow.Ticks;
        }

        private static void OpenLearningDbCircuit()
        {
            Interlocked.Exchange(ref learningDbCircuitOpenUntilTicks, DateTime.UtcNow.Ticks + LearningDbCircuitWindowTicks);
        }

        private static void ClearLearningDbCircuit()
        {
            Interlocked.Exchange(ref learningDbCircuitOpenUntilTicks, 0L);
        }

        private static bool IsLearningDbCircuitBreakingSqlErrorNumber(int number)
        {
            return number == -2 || number == 53 || number == 121 || number == 258 ||
                number == 10053 || number == 10054 || number == 10060 || number == 10061 ||
                number == 10065 || number == 11001;
        }

        private static bool IsLearningDbCircuitBreakingException(Exception ex)
        {
            if (ex == null) return false;
            SqlException sql = ex as SqlException;
            if (sql != null) return sql.Errors.Cast<SqlError>().Any(error => IsLearningDbCircuitBreakingSqlErrorNumber(error.Number));
            if (ex is TimeoutException || ex is IOException || ex is SocketException) return true;
            return ex.InnerException != null && IsLearningDbCircuitBreakingException(ex.InnerException);
        }

        private static void ObserveLearningDbFailure(Exception ex)
        {
            if (IsLearningDbCircuitBreakingException(ex)) OpenLearningDbCircuit();
        }

        // RecoLearning 是唯一学习存储：流水与推荐核心聚合在同一事务提交。
        // SQL 失败只返回未持久化状态，不写本机学习文件、outbox 或 dead-letter。
        private static bool RecordBindingEventsToLearningDb(string source, List<MappingFeedbackGroup> groups)
        {
            RememberLearningDbDurableResult(source, groups, false);
            if (groups == null || groups.Count == 0)
            {
                return false;
            }

            PrepareLearningGroupsForSql(groups);
            List<MappingFeedbackGroup> unknownPartitionGroups = groups
                .Where(group => group != null && !String.IsNullOrWhiteSpace(group.QuantityName) &&
                    !IsValidLearningSoftwarePartition(group.SoftwarePartition))
                .ToList();
            if (unknownPartitionGroups.Count > 0)
            {
                Log("Learning was rejected before SQL write because the current software partition is unknown. groups=" +
                    unknownPartitionGroups.Count.ToString(CultureInfo.InvariantCulture));
            }
            List<MappingFeedbackGroup> validGroups = groups
                .Where(group => group != null && !String.IsNullOrWhiteSpace(group.QuantityName) &&
                    IsValidLearningSoftwarePartition(group.SoftwarePartition) &&
                    !String.IsNullOrEmpty(group.MethodNo) &&
                    !String.IsNullOrEmpty(NormalizeLearningDbMethod(group.Method)) &&
                    group.Targets != null && group.Targets.Any(target => target != null &&
                        !String.IsNullOrWhiteSpace(target.Code)))
                .ToList();
            int unsupportedMethodCount = groups.Count(group => group != null && !String.IsNullOrWhiteSpace(group.QuantityName) &&
                IsValidLearningSoftwarePartition(group.SoftwarePartition) &&
                (String.IsNullOrEmpty(group.MethodNo) || String.IsNullOrEmpty(NormalizeLearningDbMethod(group.Method))));
            int invalidStructureCount = groups.Count(group => group != null && !String.IsNullOrWhiteSpace(group.QuantityName) &&
                IsValidLearningSoftwarePartition(group.SoftwarePartition) &&
                !String.IsNullOrEmpty(group.MethodNo) &&
                !String.IsNullOrEmpty(NormalizeLearningDbMethod(group.Method)) &&
                (group.Targets == null || !group.Targets.Any(target => target != null &&
                    !String.IsNullOrWhiteSpace(target.Code))));
            int unsupportedCount = unknownPartitionGroups.Count + unsupportedMethodCount + invalidStructureCount;
            if (unsupportedCount > 0)
            {
                Log("Learning DB binding was not queued because " + unsupportedCount.ToString(CultureInfo.InvariantCulture) +
                    " learning group(s) have an unsupported partition/method or no auditable target.");
            }
            if (validGroups.Count == 0)
            {
                return false;
            }

            LearningDbOutboxBatch batch = CreateLearningDbOutboxBatch(source, validGroups);
            string failureReason;
            LearningDbWriteResult writeResult = TryWriteLearningDbBatch(batch, out failureReason);
            bool durable = writeResult == LearningDbWriteResult.Succeeded;
            if (!durable)
            {
                Log("Learning DB binding was not persisted; local learning is disabled. result=" +
                    writeResult.ToString() + " reason=" + NormalizeLearningDbDeadLetterReason(failureReason));
            }
            bool fullyWritten = durable && unsupportedCount == 0;
            RememberLearningDbDurableResult(source, groups, fullyWritten);
            return fullyWritten;
        }

        private static void WriteBindingEventsToLearningDb(string connectionString, string source,
            List<MappingFeedbackGroup> groups, string batchId)
        {
            if (String.IsNullOrEmpty(connectionString)) return;
            if ((groups ?? new List<MappingFeedbackGroup>()).Any(group => group != null &&
                !String.IsNullOrWhiteSpace(group.QuantityName) &&
                (!IsValidLearningSoftwarePartition(group.SoftwarePartition) ||
                    String.IsNullOrEmpty(group.MethodNo) ||
                    String.IsNullOrEmpty(NormalizeLearningDbMethod(group.Method)))))
            {
                throw new InvalidOperationException("Learning DB partition or method identity is missing; binding was isolated.");
            }

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                Interlocked.Increment(ref learningDbConnectionOpenAttempts);
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
                                    cmd.CommandText = "INSERT INTO dbo.BindingLog (occurred_at, source, method, software_partition, method_no, project_id, entry_code, entry_name, quantity_name, quantity_unit, target_kind, target_code, target_name, target_unit, group_key, event_hash, extra) " +
                                        "VALUES (SYSDATETIME(), @source, @method, @software_partition, @method_no, @project, @ec, @en, @qn, @qu, @tk, @tc, @tn, @tu, @gk, @eh, @ex)";
                                    cmd.Parameters.AddWithValue("@source", "plugin:" + (source ?? ""));
                                    cmd.Parameters.AddWithValue("@method", NormalizeLearningDbMethod(group.Method));
                                    cmd.Parameters.AddWithValue("@software_partition", group.SoftwarePartition);
                                    cmd.Parameters.AddWithValue("@method_no", group.MethodNo);
                                    cmd.Parameters.AddWithValue("@project", group.ProjectId ?? "");
                                    cmd.Parameters.AddWithValue("@ec", GetMappingFeedbackTargetEntryCode(group, target));
                                    cmd.Parameters.AddWithValue("@en", GetMappingFeedbackTargetEntryName(group, target));
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
                                    flat["accepted_count"] = Math.Max(0, group.AcceptedCount).ToString(CultureInfo.InvariantCulture);
                                    flat["corrected_count"] = Math.Max(0, group.CorrectedCount).ToString(CultureInfo.InvariantCulture);
                                    flat["rejected_count"] = Math.Max(0, group.RejectedCount).ToString(CultureInfo.InvariantCulture);
                                    if (!String.IsNullOrWhiteSpace(group.UserAction)) flat["user_action"] = group.UserAction;
                                    string formulaEntryCode = LearningPartitionIdentity.NormalizeLearningEntryCode(
                                        GetMappingFeedbackTargetEntryCode(group, target));
                                    if (!String.IsNullOrWhiteSpace(target.FormulaTemplate) && group.FormulaOperands.Count > 0 &&
                                        !String.IsNullOrEmpty(formulaEntryCode))
                                    {
                                        flat["formula_rule_hash"] = BuildLearningFormulaRuleHash(group, target);
                                        flat["formula_template"] = target.FormulaTemplate;
                                        flat["formula_target_unit"] = target.Unit ?? "";
                                        flat["formula_software_partition"] = group.SoftwarePartition ?? "";
                                        flat["formula_method_no"] = group.MethodNo ?? "";
                                        flat["formula_entry_code"] = formulaEntryCode;
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

        private static LearningDbWriteResult TryWriteLearningDbBatch(LearningDbOutboxBatch batch, out string failureReason)
        {
            failureReason = "";
            if (batch == null || batch.Groups == null || batch.Groups.Count == 0) return LearningDbWriteResult.Succeeded;
            if (HasUnsupportedLearningDbMethod(batch, out failureReason))
            {
                Log("Learning DB batch " + (batch.BatchId ?? "") + " is permanently invalid: " + failureReason);
                return LearningDbWriteResult.PermanentFailure;
            }
            if (IsLearningDbCircuitOpen())
            {
                failureReason = "learning_db_circuit_open";
                return LearningDbWriteResult.RetryableFailure;
            }
            string connectionString = GetLearningDbConnectionString();
            if (String.IsNullOrEmpty(connectionString))
            {
                failureReason = "learning_db_connection_unavailable";
                Log("Learning DB connection is unavailable; batch " + (batch.BatchId ?? "") + " remains queued for retry.");
                return LearningDbWriteResult.RetryableFailure;
            }

            for (int attempt = 0; attempt < 2; attempt++)
            {
                try
                {
                    WriteBindingEventsToLearningDb(connectionString, batch.Source, batch.Groups, batch.BatchId);
                    return LearningDbWriteResult.Succeeded;
                }
                catch (SqlException ex)
                {
                    if (IsUniqueLearningDbConflict(ex) && IsLearningDbBatchAlreadyCommitted(connectionString, batch))
                    {
                        return LearningDbWriteResult.Succeeded;
                    }
                    if (IsLearningDbCircuitBreakingException(ex))
                    {
                        OpenLearningDbCircuit();
                        failureReason = "connection_sql_" + GetLearningDbSqlErrorNumbers(ex);
                        Log("Learning DB connection failure for batch " + (batch.BatchId ?? "") +
                            " (SQL " + GetLearningDbSqlErrorNumbers(ex) + "); circuit opened and active outbox replay will retry later.");
                        return LearningDbWriteResult.RetryableFailure;
                    }
                    if (attempt == 0 && IsRetryableLearningDbException(ex))
                    {
                        Log("Learning DB transient write failure for batch " + (batch.BatchId ?? "") +
                            "; retrying once (SQL " + GetLearningDbSqlErrorNumbers(ex) + ").");
                        Interlocked.Increment(ref learningDbRetrySleepCount);
                        Thread.Sleep(100);
                        continue;
                    }
                    if (IsRetryableLearningDbException(ex))
                    {
                        failureReason = "transient_sql_" + GetLearningDbSqlErrorNumbers(ex);
                        Log("Learning DB transient write failure for batch " + (batch.BatchId ?? "") +
                            " (SQL " + GetLearningDbSqlErrorNumbers(ex) + "); active outbox replay will retry later.");
                        return LearningDbWriteResult.RetryableFailure;
                    }
                    failureReason = "non_retryable_sql_" + GetLearningDbSqlErrorNumbers(ex);
                    Log("Learning DB non-retryable write failure for batch " + (batch.BatchId ?? "") +
                        " (SQL " + GetLearningDbSqlErrorNumbers(ex) + "); moving it to dead-letter storage.");
                    return LearningDbWriteResult.PermanentFailure;
                }
                catch (Exception ex)
                {
                    if (IsLearningDbCircuitBreakingException(ex))
                    {
                        OpenLearningDbCircuit();
                        failureReason = "connection_" + ex.GetType().Name;
                        Log("Learning DB connection failure for batch " + (batch.BatchId ?? "") +
                            " (" + ex.GetType().Name + "); circuit opened and active outbox replay will retry later.");
                        return LearningDbWriteResult.RetryableFailure;
                    }
                    failureReason = "non_retryable_" + ex.GetType().Name;
                    Log("Learning DB non-retryable write failure for batch " + (batch.BatchId ?? "") +
                        " (" + ex.GetType().Name + "); moving it to dead-letter storage.");
                    return LearningDbWriteResult.PermanentFailure;
                }
            }
            failureReason = "transient_sql_unknown";
            return LearningDbWriteResult.RetryableFailure;
        }

        private static string GetLearningDbSqlErrorNumbers(SqlException ex)
        {
            if (ex == null) return "unknown";
            string value = String.Join(",", ex.Errors.Cast<SqlError>()
                .Select(error => error.Number.ToString(CultureInfo.InvariantCulture))
                .Distinct()
                .ToArray());
            return String.IsNullOrEmpty(value) ? "unknown" : value;
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

        private static LearningDbOutboxBatch CreateLearningDbOutboxBatch(string source, List<MappingFeedbackGroup> groups)
        {
            string processName;
            string moduleFileName;
            ReadLearningProcessIdentity(out processName, out moduleFileName);
            List<MappingFeedbackGroup> safeGroups = groups ?? new List<MappingFeedbackGroup>();
            List<string> partitions = safeGroups
                .Where(group => group != null && IsValidLearningSoftwarePartition(group.SoftwarePartition))
                .Select(group => group.SoftwarePartition)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            return new LearningDbOutboxBatch
            {
                BatchId = Guid.NewGuid().ToString("N"),
                Source = source ?? "",
                SoftwarePartition = partitions.Count == 1 ? partitions[0] : "",
                ProcessName = processName,
                ModuleFileName = moduleFileName,
                Groups = safeGroups
            };
        }

        private static bool IsValidLearningSoftwarePartition(string value)
        {
            return String.Equals(value, "2020", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(value, "2024", StringComparison.OrdinalIgnoreCase);
        }

        private static void ReadLearningProcessIdentity(out string processName, out string moduleFileName)
        {
            processName = "";
            moduleFileName = "";
            try
            {
                System.Diagnostics.Process process = System.Diagnostics.Process.GetCurrentProcess();
                processName = process.ProcessName ?? "";
                try
                {
                    string path = process.MainModule == null ? "" : (process.MainModule.FileName ?? "");
                    moduleFileName = Path.GetFileName(path) ?? "";
                }
                catch
                {
                    moduleFileName = "";
                }
            }
            catch
            {
                processName = "";
                moduleFileName = "";
            }
        }

        private static void PrepareLearningGroupsForSql(List<MappingFeedbackGroup> groups)
        {
            foreach (MappingFeedbackGroup group in groups ?? new List<MappingFeedbackGroup>())
            {
                if (group == null) continue;
                string rawMethod = group.Method;
                if (String.IsNullOrWhiteSpace(group.MethodNo))
                {
                    group.MethodNo = NormalizeLearningMethodNo(rawMethod);
                }
                else
                {
                    group.MethodNo = NormalizeLearningMethodNo(group.MethodNo);
                }
                if (String.IsNullOrWhiteSpace(group.SoftwarePartition))
                {
                    group.SoftwarePartition = ResolveLearningSoftwarePartition();
                }
                group.Method = NormalizeLearningDbMethod(rawMethod);
                List<MappingFeedbackTarget> targets = (group.Targets ?? new List<MappingFeedbackTarget>())
                    .Where(target => target != null && !String.IsNullOrWhiteSpace(target.Code))
                    .ToList();
                bool containsContextSensitiveTarget = targets.Any(target => IsContextSensitiveLearningCode(target.Code));
                if (!containsContextSensitiveTarget && !String.IsNullOrWhiteSpace(group.BoxId)) continue;
                List<string> targetKeys = targets
                    .Select(target => BuildLearningTargetIdentityKey(target.Kind, target.Code, target.Name, target.Unit))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .OrderBy(key => key, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                if (targetKeys.Count > 0)
                {
                    if (containsContextSensitiveTarget)
                    {
                        string targetSetHash = BuildLearningMd5(String.Join(";", targetKeys.ToArray()));
                        group.BoxId = "auto-" + targetSetHash.Substring(0, 16);
                    }
                    else
                    {
                        group.BoxId = BuildStableMappingBoxId(String.Join("|", targetKeys.ToArray()));
                    }
                }
            }
        }

        private static bool HasUnsupportedLearningDbMethod(LearningDbOutboxBatch batch, out string reason)
        {
            reason = "";
            if (batch == null || batch.Groups == null) return false;
            int invalidPartitionCount = batch.Groups.Count(group => group != null && !String.IsNullOrWhiteSpace(group.QuantityName) &&
                !IsValidLearningSoftwarePartition(group.SoftwarePartition));
            int partitionMismatchCount = batch.Groups.Count(group => group != null && !String.IsNullOrWhiteSpace(group.QuantityName) &&
                IsValidLearningSoftwarePartition(group.SoftwarePartition) &&
                (!IsValidLearningSoftwarePartition(batch.SoftwarePartition) ||
                    !String.Equals(group.SoftwarePartition, batch.SoftwarePartition, StringComparison.OrdinalIgnoreCase)));
            int invalidMethodCount = batch.Groups.Count(group => group != null && !String.IsNullOrWhiteSpace(group.QuantityName) &&
                (String.IsNullOrEmpty(group.MethodNo) || String.IsNullOrEmpty(NormalizeLearningDbMethod(group.Method))));
            int invalidEvidenceCount = batch.Groups.Count(group => group != null && !String.IsNullOrWhiteSpace(group.QuantityName) &&
                IsValidLearningSoftwarePartition(group.SoftwarePartition) &&
                !String.IsNullOrEmpty(group.MethodNo) &&
                !String.IsNullOrEmpty(NormalizeLearningDbMethod(group.Method)) && !IsLearningFeedbackGroupRecommendable(group));
            if (invalidPartitionCount == 0 && partitionMismatchCount == 0 && invalidMethodCount == 0 && invalidEvidenceCount == 0) return false;
            reason = "unsupported_learning_group_partition_" + invalidPartitionCount.ToString(CultureInfo.InvariantCulture) +
                "_mismatch_" + partitionMismatchCount.ToString(CultureInfo.InvariantCulture) +
                "_method_" + invalidMethodCount.ToString(CultureInfo.InvariantCulture) +
                "_evidence_" + invalidEvidenceCount.ToString(CultureInfo.InvariantCulture);
            return true;
        }

        private static bool TryAppendLearningDbOutbox(LearningDbOutboxBatch batch)
        {
            if (batch == null || String.IsNullOrWhiteSpace(batch.BatchId) || batch.Groups == null || batch.Groups.Count == 0) return false;
            string unsupportedReason;
            if (HasUnsupportedLearningDbMethod(batch, out unsupportedReason))
            {
                Log("Learning DB batch " + (batch.BatchId ?? "") + " was not queued because " + unsupportedReason +
                    "; local mapping remains available.");
                return false;
            }
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

        private static bool TryAppendLearningDbDeadLetter(LearningDbOutboxBatch batch, string reason)
        {
            if (batch == null || String.IsNullOrWhiteSpace(batch.BatchId)) return false;
            try
            {
                bool saved = false;
                bool locked = TryWithLearningDbOutboxLock(delegate
                {
                    string path = GetLearningDbDeadLetterPath();
                    Directory.CreateDirectory(Path.GetDirectoryName(path));
                    List<string> lines = File.Exists(path)
                        ? File.ReadAllLines(path, Encoding.UTF8).ToList()
                        : new List<string>();
                    bool exists = lines.Any(line => String.Equals(GetFlat(ParseFlatJson(line), "batch_id"), batch.BatchId,
                        StringComparison.OrdinalIgnoreCase));
                    if (!exists)
                    {
                        Dictionary<string, string> row = SerializeLearningDbOutboxBatch(batch);
                        row["dead_letter_reason"] = NormalizeLearningDbDeadLetterReason(reason);
                        row["dead_letter_at"] = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
                        lines.Add(ToFlatJson(row));
                        WriteAllLinesAtomic(path, lines.ToArray(), Encoding.UTF8);
                    }
                    saved = true;
                }, 5000);
                if (!locked) Log("Learning DB dead-letter lock timeout; rejected batch was not persisted.");
                return locked && saved;
            }
            catch (Exception ex)
            {
                Log("Learning DB dead-letter append failed (" + ex.GetType().Name + ").");
                return false;
            }
        }

        private static bool TryMoveLearningDbOutboxBatchToDeadLetter(string batchId, string reason)
        {
            if (String.IsNullOrWhiteSpace(batchId)) return false;
            try
            {
                bool moved = false;
                bool locked = TryWithLearningDbOutboxLock(delegate
                {
                    string activePath = GetLearningDbOutboxPath();
                    if (!File.Exists(activePath))
                    {
                        moved = true;
                        return;
                    }

                    List<string> activeLines = File.ReadAllLines(activePath, Encoding.UTF8).ToList();
                    List<string> rejectedLines = activeLines
                        .Where(line => String.Equals(GetFlat(ParseFlatJson(line), "batch_id"), batchId, StringComparison.OrdinalIgnoreCase))
                        .ToList();
                    if (rejectedLines.Count == 0)
                    {
                        moved = true;
                        return;
                    }

                    string deadLetterPath = GetLearningDbDeadLetterPath();
                    Directory.CreateDirectory(Path.GetDirectoryName(deadLetterPath));
                    List<string> deadLetterLines = File.Exists(deadLetterPath)
                        ? File.ReadAllLines(deadLetterPath, Encoding.UTF8).ToList()
                        : new List<string>();
                    bool alreadySaved = deadLetterLines.Any(line => String.Equals(GetFlat(ParseFlatJson(line), "batch_id"), batchId,
                        StringComparison.OrdinalIgnoreCase));
                    if (!alreadySaved)
                    {
                        Dictionary<string, string> row = ParseFlatJson(rejectedLines[0]);
                        row["dead_letter_reason"] = NormalizeLearningDbDeadLetterReason(reason);
                        row["dead_letter_at"] = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture);
                        deadLetterLines.Add(ToFlatJson(row));
                        WriteAllLinesAtomic(deadLetterPath, deadLetterLines.ToArray(), Encoding.UTF8);
                    }

                    List<string> remaining = activeLines
                        .Where(line => !String.Equals(GetFlat(ParseFlatJson(line), "batch_id"), batchId, StringComparison.OrdinalIgnoreCase))
                        .ToList();
                    WriteAllLinesAtomic(activePath, remaining.ToArray(), Encoding.UTF8);
                    moved = true;
                }, 5000);
                if (!locked) Log("Learning DB outbox dead-letter move lock timeout; rejected batch remains active.");
                return locked && moved;
            }
            catch (Exception ex)
            {
                Log("Learning DB outbox dead-letter move failed (" + ex.GetType().Name + "); rejected batch remains active.");
                return false;
            }
        }

        private static string NormalizeLearningDbDeadLetterReason(string reason)
        {
            string value = (reason ?? "unknown").Trim();
            if (value.Length == 0) value = "unknown";
            StringBuilder safe = new StringBuilder();
            foreach (char ch in value)
            {
                if (Char.IsLetterOrDigit(ch) || ch == '_' || ch == '-' || ch == ',' || ch == '(' || ch == ')') safe.Append(ch);
            }
            if (safe.Length == 0) return "unknown";
            return safe.Length <= 200 ? safe.ToString() : safe.ToString(0, 200);
        }

        private static Dictionary<string, string> SerializeLearningDbOutboxBatch(LearningDbOutboxBatch batch)
        {
            Dictionary<string, string> row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            row["record_type"] = "learning_db_outbox";
            row["batch_id"] = batch.BatchId ?? "";
            row["source"] = batch.Source ?? "";
            row["software_partition"] = batch.SoftwarePartition ?? "";
            row["process_name"] = batch.ProcessName ?? "";
            row["module_file_name"] = batch.ModuleFileName ?? "";
            row["group_count"] = batch.Groups.Count.ToString(CultureInfo.InvariantCulture);
            for (int groupIndex = 0; groupIndex < batch.Groups.Count; groupIndex++)
            {
                MappingFeedbackGroup group = batch.Groups[groupIndex] ?? new MappingFeedbackGroup();
                string prefix = "g" + groupIndex.ToString(CultureInfo.InvariantCulture) + "_";
                row[prefix + "quantity_name"] = group.QuantityName ?? "";
                row[prefix + "quantity_unit"] = group.QuantityUnit ?? "";
                row[prefix + "method"] = NormalizeLearningDbMethod(group.Method);
                row[prefix + "software_partition"] = group.SoftwarePartition ?? "";
                row[prefix + "method_no"] = group.MethodNo ?? "";
                row[prefix + "project_id"] = group.ProjectId ?? "";
                row[prefix + "entry_code"] = group.EntryCode ?? "";
                row[prefix + "entry_name"] = group.EntryName ?? "";
                row[prefix + "workbook"] = group.Workbook ?? "";
                row[prefix + "worksheet"] = group.Worksheet ?? "";
                row[prefix + "excel_row"] = group.ExcelRow.ToString(CultureInfo.InvariantCulture);
                row[prefix + "box_id"] = group.BoxId ?? "";
                row[prefix + "expression"] = group.Expression ?? "";
                row[prefix + "source_cell"] = group.SourceCell ?? "";
                row[prefix + "accepted_count"] = group.AcceptedCount.ToString(CultureInfo.InvariantCulture);
                row[prefix + "corrected_count"] = group.CorrectedCount.ToString(CultureInfo.InvariantCulture);
                row[prefix + "rejected_count"] = group.RejectedCount.ToString(CultureInfo.InvariantCulture);
                row[prefix + "user_action"] = group.UserAction ?? "";
                row[prefix + "target_count"] = (group.Targets == null ? 0 : group.Targets.Count).ToString(CultureInfo.InvariantCulture);
                for (int targetIndex = 0; targetIndex < (group.Targets == null ? 0 : group.Targets.Count); targetIndex++)
                {
                    MappingFeedbackTarget target = group.Targets[targetIndex] ?? new MappingFeedbackTarget();
                    string targetPrefix = prefix + "t" + targetIndex.ToString(CultureInfo.InvariantCulture) + "_";
                    row[targetPrefix + "kind"] = target.Kind ?? "";
                    row[targetPrefix + "code"] = target.Code ?? "";
                    row[targetPrefix + "name"] = target.Name ?? "";
                    row[targetPrefix + "unit"] = target.Unit ?? "";
                    row[targetPrefix + "entry_code"] = target.EntryCode ?? "";
                    row[targetPrefix + "entry_name"] = target.EntryName ?? "";
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
            LearningDbOutboxBatch batch = new LearningDbOutboxBatch
            {
                BatchId = batchId,
                Source = GetFlat(row, "source"),
                SoftwarePartition = GetFlat(row, "software_partition").Trim(),
                ProcessName = GetFlat(row, "process_name"),
                ModuleFileName = GetFlat(row, "module_file_name")
            };
            for (int groupIndex = 0; groupIndex < groupCount; groupIndex++)
            {
                string prefix = "g" + groupIndex.ToString(CultureInfo.InvariantCulture) + "_";
                MappingFeedbackGroup group = new MappingFeedbackGroup
                {
                    QuantityName = GetFlat(row, prefix + "quantity_name"),
                    QuantityUnit = GetFlat(row, prefix + "quantity_unit"),
                    Method = NormalizeLearningDbMethod(GetFlat(row, prefix + "method")),
                    SoftwarePartition = GetFlat(row, prefix + "software_partition").Trim(),
                    MethodNo = NormalizeLearningMethodNo(GetFlat(row, prefix + "method_no")),
                    ProjectId = GetFlat(row, prefix + "project_id"),
                    EntryCode = GetFlat(row, prefix + "entry_code"),
                    EntryName = GetFlat(row, prefix + "entry_name"),
                    Workbook = GetFlat(row, prefix + "workbook"),
                    Worksheet = GetFlat(row, prefix + "worksheet"),
                    ExcelRow = ReadFlatInt(row, prefix + "excel_row", 0),
                    BoxId = GetFlat(row, prefix + "box_id"),
                    Expression = GetFlat(row, prefix + "expression"),
                    SourceCell = GetFlat(row, prefix + "source_cell"),
                    AcceptedCount = ReadFlatInt(row, prefix + "accepted_count", 1),
                    CorrectedCount = ReadFlatInt(row, prefix + "corrected_count", 0),
                    RejectedCount = ReadFlatInt(row, prefix + "rejected_count", 0),
                    UserAction = GetFlat(row, prefix + "user_action")
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
                        EntryCode = GetFlat(row, targetPrefix + "entry_code"),
                        EntryName = GetFlat(row, targetPrefix + "entry_name"),
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
            if (IsLearningDbCircuitOpen()) return false;
            List<LearningDbOutboxBatch> batches = LoadLearningDbOutboxBatches();
            foreach (LearningDbOutboxBatch batch in batches.Take(20))
            {
                string failureReason;
                LearningDbWriteResult writeResult = TryWriteLearningDbBatch(batch, out failureReason);
                if (writeResult == LearningDbWriteResult.Succeeded)
                {
                    RemoveLearningDbOutboxBatch(batch.BatchId);
                    continue;
                }
                if (writeResult == LearningDbWriteResult.PermanentFailure)
                {
                    Log("Learning DB outbox batch " + (batch.BatchId ?? "") +
                        " is permanently invalid and will be isolated; subsequent batches will continue.");
                    if (!TryMoveLearningDbOutboxBatchToDeadLetter(batch.BatchId, failureReason)) return false;
                    continue;
                }
                return false;
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
                    if (String.IsNullOrEmpty(NormalizeLearningDbMethod(group.Method))) continue;
                    if (Math.Max(0, group.AcceptedCount) + Math.Max(0, group.CorrectedCount) <= 0) continue;
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

        private static string GetLearningDbDeadLetterPath()
        {
            return Path.Combine(FindRecoQuotaDataDir(), LearningDbDeadLetterFileName);
        }

        private static void RememberLearningDbDurableResult(string source, List<MappingFeedbackGroup> groups, bool durable)
        {
            if (!String.Equals(source ?? "", "template-right-click", StringComparison.OrdinalIgnoreCase) &&
                !String.Equals(source ?? "", "name-match", StringComparison.OrdinalIgnoreCase) &&
                !String.Equals(source ?? "", "apply-accept", StringComparison.OrdinalIgnoreCase)) return;
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
                if (error.Number == -2 || error.Number == -1 || error.Number == 0 || error.Number == 2 || error.Number == 20 ||
                    error.Number == 53 || error.Number == 64 || error.Number == 121 || error.Number == 233 || error.Number == 258 ||
                    error.Number == 4060 ||
                    error.Number == 1205 || error.Number == 2601 || error.Number == 2627 ||
                    error.Number == 10053 || error.Number == 10054 || error.Number == 10060 || error.Number == 10061 || error.Number == 10065 ||
                    error.Number == 11001)
                {
                    return true;
                }
            }
            return false;
        }

        // 绑定流水是审计源；同时增量维护推荐预览直接读取的核心聚合表，避免等待全量收割。
        private static void UpsertBindingGroupAggregates(SqlConnection conn, SqlTransaction transaction, MappingFeedbackGroup group)
        {
            List<MappingFeedbackTarget> eligibleTargets = (group.Targets ?? new List<MappingFeedbackTarget>())
                .Where(target => target != null && !String.IsNullOrWhiteSpace(target.Code))
                .ToList();
            if (!IsLearningFeedbackGroupRecommendable(group)) return;

            List<MappingFeedbackTarget> targets = eligibleTargets
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
            List<string> targetKeys = targets.Select(target =>
                BuildLearningTargetIdentityKey(target.Kind, target.Code, target.Name, target.Unit)).ToList();
            string targetSetHash = BuildLearningMd5(String.Join(";", targetKeys.ToArray()));
            bool containsContextSensitiveTarget = targets.Any(target => IsContextSensitiveLearningCode(target.Code));
            string stableAggregateBoxId = containsContextSensitiveTarget
                ? "auto-" + targetSetHash.Substring(0, 16)
                : BuildStableMappingBoxId(String.Join("|", targetKeys.ToArray()));
            string boxId = TrimLearningText(group.BoxId, 64);
            if (containsContextSensitiveTarget || String.IsNullOrWhiteSpace(boxId)) boxId = stableAggregateBoxId;

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
                cmd.CommandText = "SELECT target_set_hash FROM dbo.QuotaBox WITH (UPDLOCK,HOLDLOCK) WHERE box_id=@box";
                cmd.Parameters.AddWithValue("@box", boxId);
                object existingHash = cmd.ExecuteScalar();
                if (existingHash != null && existingHash != DBNull.Value &&
                    !String.Equals(Convert.ToString(existingHash), targetSetHash, StringComparison.OrdinalIgnoreCase))
                {
                    boxId = stableAggregateBoxId;
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
                int acceptedDelta = Math.Max(0, group.AcceptedCount);
                int correctedDelta = Math.Max(0, group.CorrectedCount);
                int rejectedDelta = Math.Max(0, group.RejectedCount);
                if (acceptedDelta + correctedDelta + rejectedDelta == 0) acceptedDelta = 1;
                int initialWeight = Math.Max(0,
                    10 * acceptedDelta + 20 * correctedDelta - 10 * rejectedDelta);
                cmd.CommandText =
                    "UPDATE dbo.SignatureBoxMap WITH (UPDLOCK,HOLDLOCK) SET accepted_count=accepted_count+@accepted, " +
                    "corrected_count=corrected_count+@corrected, rejected_count=rejected_count+@rejected, " +
                    "weight=CASE WHEN 10*(accepted_count+@accepted)+20*(corrected_count+@corrected)-10*(rejected_count+@rejected)<0 THEN 0 " +
                    "ELSE 10*(accepted_count+@accepted)+20*(corrected_count+@corrected)-10*(rejected_count+@rejected) END, last_used_at=SYSDATETIME() " +
                    "WHERE software_partition=@software_partition AND signature=@s AND box_id=@box; " +
                    "IF @@ROWCOUNT=0 INSERT INTO dbo.SignatureBoxMap(software_partition,signature,method,box_id,weight,accepted_count,corrected_count,rejected_count,last_used_at) " +
                    "VALUES(@software_partition,@s,@method,@box,@weight,@accepted,@corrected,@rejected,SYSDATETIME());";
                cmd.Parameters.AddWithValue("@software_partition", group.SoftwarePartition);
                cmd.Parameters.AddWithValue("@s", signature);
                cmd.Parameters.AddWithValue("@method", method);
                cmd.Parameters.AddWithValue("@box", boxId);
                cmd.Parameters.AddWithValue("@accepted", acceptedDelta);
                cmd.Parameters.AddWithValue("@corrected", correctedDelta);
                cmd.Parameters.AddWithValue("@rejected", rejectedDelta);
                cmd.Parameters.AddWithValue("@weight", initialWeight);
                cmd.ExecuteNonQuery();
            }

            bool hasFormulaTables = false;
            bool positiveEvidence = Math.Max(0, group.AcceptedCount) + Math.Max(0, group.CorrectedCount) > 0;
            if (positiveEvidence && group.FormulaOperands.Count > 0 && targets.Any(target => !String.IsNullOrWhiteSpace(target.FormulaTemplate)))
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
                    string formulaEntry = TrimLearningText(LearningPartitionIdentity.NormalizeLearningEntryCode(
                        GetMappingFeedbackTargetEntryCode(group, target)), 100);
                    if (formulaEntry.Length == 0) continue;
                    string ruleHash = BuildLearningFormulaRuleHash(group, target);
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandTimeout = 5;
                        cmd.CommandText =
                            "UPDATE dbo.QuantityFormulaRule WITH (UPDLOCK,HOLDLOCK) SET sample_count=sample_count+1,last_seen=SYSDATETIME() WHERE rule_hash=@hash; " +
                            "IF @@ROWCOUNT=0 INSERT INTO dbo.QuantityFormulaRule(rule_hash,anchor_signature,target_kind,target_code,target_unit,formula_template,method,software_partition,method_no,entry_code,sample_count,first_seen,last_seen) " +
                            "VALUES(@hash,@s,@kind,@code,@unit,@formula,@method,@software_partition,@method_no,@entry,1,SYSDATETIME(),SYSDATETIME());";
                        cmd.Parameters.AddWithValue("@hash", ruleHash);
                        cmd.Parameters.AddWithValue("@s", signature);
                        cmd.Parameters.AddWithValue("@kind", targetKind);
                        cmd.Parameters.AddWithValue("@code", targetCode);
                        cmd.Parameters.AddWithValue("@unit", targetUnit);
                        cmd.Parameters.AddWithValue("@formula", formulaTemplate);
                        cmd.Parameters.AddWithValue("@method", method);
                        cmd.Parameters.AddWithValue("@software_partition", group.SoftwarePartition);
                        cmd.Parameters.AddWithValue("@method_no", group.MethodNo);
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

            if (positiveEvidence)
            {
                string method = NormalizeLearningDbMethod(group.Method);
                foreach (MappingFeedbackTarget scopeTarget in targets.Where(target =>
                    IsEngineeringScopeLearningTarget(target.Kind, target.Code))
                    .GroupBy(target => GetMappingFeedbackTargetEntryCode(group, target), StringComparer.OrdinalIgnoreCase)
                    .Select(items => items.First()))
                {
                    string entryCode = TrimLearningText(LearningPartitionIdentity.NormalizeLearningEntryCode(
                        GetMappingFeedbackTargetEntryCode(group, scopeTarget)), 100);
                    if (!IsSmartClassifiedEntryCode(entryCode)) continue;
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandTimeout = 5;
                        cmd.CommandText =
                            "UPDATE dbo.EngineeringTemplate WITH (UPDLOCK,HOLDLOCK) SET sample_count=sample_count+1,last_seen=SYSDATETIME() " +
                            "WHERE software_partition=@software_partition AND method_no=@method_no AND engineering_type=@type AND entry_code=@entry AND box_id=@box; " +
                            "IF @@ROWCOUNT=0 INSERT INTO dbo.EngineeringTemplate(software_partition,method_no,method,engineering_type,entry_code,box_id,sample_count,last_seen) " +
                            "VALUES(@software_partition,@method_no,@method,@type,@entry,@box,1,SYSDATETIME());";
                        cmd.Parameters.AddWithValue("@software_partition", group.SoftwarePartition);
                        cmd.Parameters.AddWithValue("@method_no", group.MethodNo);
                        cmd.Parameters.AddWithValue("@method", method);
                        cmd.Parameters.AddWithValue("@type", entryCode.Substring(0, 2));
                        cmd.Parameters.AddWithValue("@entry", entryCode);
                        cmd.Parameters.AddWithValue("@box", boxId);
                        cmd.ExecuteNonQuery();
                    }
                }
                foreach (MappingFeedbackTarget target in targets.Where(item => String.Equals(item.Kind ?? "quota", "quota", StringComparison.OrdinalIgnoreCase)))
                {
                    string entryCode = TrimLearningText(LearningPartitionIdentity.NormalizeLearningEntryCode(
                        GetMappingFeedbackTargetEntryCode(group, target)), 100);
                    if (entryCode.Length == 0) continue;
                    string entryName = TrimLearningText(GetMappingFeedbackTargetEntryName(group, target), 500);
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandTimeout = 5;
                        cmd.CommandText =
                            "UPDATE dbo.SignatureEntryMap WITH (UPDLOCK,HOLDLOCK) SET entry_name=CASE WHEN @entry_name='' THEN entry_name ELSE @entry_name END, " +
                            "sample_count=sample_count+1,last_used_at=SYSDATETIME() WHERE software_partition=@software_partition AND method_no=@method_no AND signature=@s AND target_code=@code AND entry_code=@entry; " +
                            "IF @@ROWCOUNT=0 INSERT INTO dbo.SignatureEntryMap(software_partition,method_no,signature,target_code,method,entry_code,entry_name,sample_count,last_used_at) VALUES(@software_partition,@method_no,@s,@code,@method,@entry,@entry_name,1,SYSDATETIME());";
                        cmd.Parameters.AddWithValue("@software_partition", group.SoftwarePartition);
                        cmd.Parameters.AddWithValue("@method_no", group.MethodNo);
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

        private static string NormalizeLearningMethodNo(string rawMethod)
        {
            return LearningPartitionIdentity.NormalizeLearningMethodNo(rawMethod);
        }

        private static string ResolveLearningSoftwarePartition()
        {
            string processName;
            string moduleFileName;
            ReadLearningProcessIdentity(out processName, out moduleFileName);
            return LearningPartitionIdentity.ResolveFromProcessIdentity(processName, moduleFileName);
        }

        private static string BuildLearningFormulaRuleHash(MappingFeedbackGroup group, MappingFeedbackTarget target)
        {
            string signature = NormalizeForSignature(group == null ? "" : group.QuantityName) + "|";
            if (signature.Length > 450) signature = signature.Substring(0, 450);
            string kind = target == null || String.IsNullOrWhiteSpace(target.Kind) ? "quota" : target.Kind.Trim().ToLowerInvariant();
            string entryCode = LearningPartitionIdentity.NormalizeLearningEntryCode(
                GetMappingFeedbackTargetEntryCode(group, target));
            string raw = signature + "|" + kind + ":" + (target == null ? "" : target.Code ?? "").Trim().ToUpperInvariant() + "|" +
                NormalizeExcelLinkUnit(target == null ? "" : target.Unit) + "|" + (target == null ? "" : target.FormulaTemplate ?? "") + "|" +
                (group == null ? "" : group.SoftwarePartition ?? "") + "|" +
                (group == null ? "" : NormalizeLearningMethodNo(group.MethodNo)) + "|" + entryCode;
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
                learningDbConnectionString = RecoSqlCredentialStore.BuildConnectionString("learning", "RecoLearning", 1433, 3);
            }

            return learningDbConnectionString;
        }
    }
}
