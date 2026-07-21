using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        // 学习库固定在中央服务器,不跟随各软件的项目服务器(2020版连192.168.2.13,学习库统一在.213)。
        private const string LearningDbServer = "192.168.2.213,1433";

        private static string learningDbConnectionString;
        private static bool learningDbUnavailable;

        // RecoLearning 是多人共享主学习库：流水与推荐核心聚合在同一事务提交；
        // 任何失败只记日志并在本进程内停用，不影响绑定主流程或本机 jsonl 备份。
        private static void RecordBindingEventsToLearningDb(string source, List<MappingFeedbackGroup> groups)
        {
            if (learningDbUnavailable || groups == null || groups.Count == 0)
            {
                return;
            }

            try
            {
                string connectionString = GetLearningDbConnectionString();
                if (String.IsNullOrEmpty(connectionString))
                {
                    return;
                }

                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (SqlTransaction transaction = conn.BeginTransaction())
                    {
                        try
                        {
                            foreach (MappingFeedbackGroup group in groups)
                            {
                                if (group == null || String.IsNullOrWhiteSpace(group.QuantityName))
                                {
                                    continue;
                                }

                                string groupKey = Guid.NewGuid().ToString("N");
                                foreach (MappingFeedbackTarget target in group.Targets)
                                {
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
                                        cmd.Parameters.AddWithValue("@method", group.Method ?? "");
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
                                        cmd.Parameters.AddWithValue("@eh", Guid.NewGuid().ToString("N"));
                                        Dictionary<string, string> flat = new Dictionary<string, string>();
                                        if (!String.IsNullOrEmpty(group.Workbook)) flat["workbook"] = group.Workbook;
                                        if (!String.IsNullOrEmpty(group.Worksheet)) flat["worksheet"] = group.Worksheet;
                                        if (group.ExcelRow > 0) flat["excel_row"] = group.ExcelRow.ToString();
                                        if (!String.IsNullOrEmpty(group.BoxId)) flat["box_id"] = group.BoxId;
                                        if (!String.IsNullOrEmpty(group.Expression)) flat["expression"] = group.Expression;
                                        if (!String.IsNullOrEmpty(group.SourceCell)) flat["source_cell"] = group.SourceCell;
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
            catch (Exception ex)
            {
                learningDbUnavailable = true;
                Log("Learning DB double-write disabled: " + ex.Message);
            }
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
                    "UPDATE dbo.QuantityAlias SET raw_name=@n, quantity_unit=@u, signature=@s, seen_count=seen_count+1, last_seen=SYSDATETIME() WHERE alias_hash=@h; " +
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
                cmd.CommandText = "IF NOT EXISTS(SELECT 1 FROM dbo.QuotaBox WHERE box_id=@box) INSERT INTO dbo.QuotaBox(box_id,target_set_hash,status) VALUES(@box,@hash,'active')";
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
                        "UPDATE dbo.QuotaBoxTarget SET " +
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
                cmd.CommandText =
                    "UPDATE dbo.SignatureBoxMap SET accepted_count=accepted_count+1, " +
                    "weight=CASE WHEN 10*(accepted_count+1)+20*corrected_count-10*rejected_count>100 THEN 100 " +
                    "WHEN 10*(accepted_count+1)+20*corrected_count-10*rejected_count<0 THEN 0 " +
                    "ELSE 10*(accepted_count+1)+20*corrected_count-10*rejected_count END, last_used_at=SYSDATETIME() " +
                    "WHERE signature=@s AND box_id=@box; " +
                    "IF @@ROWCOUNT=0 INSERT INTO dbo.SignatureBoxMap(signature,box_id,weight,accepted_count,corrected_count,rejected_count,last_used_at) VALUES(@s,@box,10,1,0,0,SYSDATETIME());";
                cmd.Parameters.AddWithValue("@s", signature);
                cmd.Parameters.AddWithValue("@box", boxId);
                cmd.ExecuteNonQuery();
            }

            string entryCode = TrimLearningText(group.EntryCode, 100);
            if (entryCode.Length > 0)
            {
                string method = TrimLearningText(group.Method, 50);
                string entryName = TrimLearningText(group.EntryName, 500);
                foreach (MappingFeedbackTarget target in targets.Where(item => String.Equals(item.Kind ?? "quota", "quota", StringComparison.OrdinalIgnoreCase)))
                {
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.Transaction = transaction;
                        cmd.CommandTimeout = 5;
                        cmd.CommandText =
                            "UPDATE dbo.SignatureEntryMap SET entry_name=CASE WHEN @entry_name='' THEN entry_name ELSE @entry_name END, " +
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
