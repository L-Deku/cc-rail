using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        // 学习库固定在中央服务器,不跟随各软件的项目服务器(2020版连192.168.2.13,学习库统一在.213)。
        private const string LearningDbServer = "192.168.2.213,1433";

        private static string learningDbConnectionString;
        private static bool learningDbUnavailable;

        // RecoLearning 双写:只追加 BindingLog 流水;任何失败只记日志并在本进程内停用,绝不影响绑定主流程。
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
                                cmd.CommandTimeout = 5;
                                cmd.CommandText = "INSERT INTO dbo.BindingLog (occurred_at, source, quantity_name, quantity_unit, entry_code, target_kind, target_code, target_name, target_unit, group_key, event_hash) " +
                                    "VALUES (SYSDATETIME(), @source, @qn, @qu, @ec, @tk, @tc, @tn, @tu, @gk, @eh)";
                                cmd.Parameters.AddWithValue("@source", "plugin:" + (source ?? ""));
                                cmd.Parameters.AddWithValue("@qn", group.QuantityName ?? "");
                                cmd.Parameters.AddWithValue("@qu", group.QuantityUnit ?? "");
                                cmd.Parameters.AddWithValue("@ec", group.EntryCode ?? "");
                                cmd.Parameters.AddWithValue("@tk", String.IsNullOrEmpty(target.Kind) ? "quota" : target.Kind);
                                cmd.Parameters.AddWithValue("@tc", target.Code ?? "");
                                cmd.Parameters.AddWithValue("@tn", target.Name ?? "");
                                cmd.Parameters.AddWithValue("@tu", target.Unit ?? "");
                                cmd.Parameters.AddWithValue("@gk", groupKey);
                                cmd.Parameters.AddWithValue("@eh", Guid.NewGuid().ToString("N"));
                                cmd.ExecuteNonQuery();
                            }
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
