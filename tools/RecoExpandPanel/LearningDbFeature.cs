using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Windows.Forms;
using System.Xml;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
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
                                cmd.CommandText = "INSERT INTO dbo.BindingLog (occurred_at, source, quantity_name, quantity_unit, target_kind, target_code, target_name, target_unit, group_key, event_hash) " +
                                    "VALUES (SYSDATETIME(), @source, @qn, @qu, @tk, @tc, @tn, @tu, @gk, @eh)";
                                cmd.Parameters.AddWithValue("@source", "plugin:" + (source ?? ""));
                                cmd.Parameters.AddWithValue("@qn", group.QuantityName ?? "");
                                cmd.Parameters.AddWithValue("@qu", "");
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
            if (learningDbConnectionString != null)
            {
                return learningDbConnectionString;
            }

            try
            {
                string dir = Path.GetDirectoryName(typeof(FormPanel).Assembly.Location);
                string settingPath = Path.Combine(dir, "ServerSetting.xml");
                if (!File.Exists(settingPath))
                {
                    learningDbConnectionString = "";
                    return "";
                }

                XmlDocument doc = new XmlDocument();
                doc.Load(settingPath);
                XmlNode ipNode = doc.SelectSingleNode("//ServerIP");
                XmlNode portNode = doc.SelectSingleNode("//SqlPort");
                string ip = ipNode != null ? (ipNode.InnerText ?? "").Trim() : "";
                string port = portNode != null ? (portNode.InnerText ?? "").Trim() : "1433";
                if (String.IsNullOrEmpty(ip))
                {
                    learningDbConnectionString = "";
                    return "";
                }

                learningDbConnectionString = "Server=" + ip + "," + port + ";Database=RecoLearning;User ID=" + AgentDbUser + ";Password=" + AgentDbPassword + ";Connect Timeout=3";
            }
            catch (Exception ex)
            {
                Log("Learning DB connection setup failed: " + ex.Message);
                learningDbConnectionString = "";
            }

            return learningDbConnectionString;
        }
    }
}
