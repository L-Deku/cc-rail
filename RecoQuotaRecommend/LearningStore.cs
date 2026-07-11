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

namespace RecoQuotaRecommend
{
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
