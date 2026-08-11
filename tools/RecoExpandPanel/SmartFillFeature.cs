using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        private const string SmartMethodNo30 = "30号文";
        private const string SmartMethodNo101Estimate = "101号文估算";
        private const string SmartMethodNo2024 = "TB 10801—2024";

        // ============ 智能铺量(学习库)漏斗匹配引擎 ============
        // 漏斗:①名称级签名 ②模糊候选(仅进下拉,不自动采纳) ③手挂。
        // 数据源只允许 RecoLearning(SQL)；SQL 不可用时停止配对并返回明确提示。

        private sealed class SmartBoxTarget
        {
            public string Kind; public string Code; public string Name; public string Unit;
        }

        private sealed class SmartMapEntry
        {
            public string BoxId;
            public int Weight;
            public int AcceptedCount;
            public int CorrectedCount;
            public int RejectedCount;
            public DateTime LastUsedAt;
            public List<SmartBoxTarget> Targets = new List<SmartBoxTarget>();
            public bool CurrentMethodMapping;
            // 本机 mapping-boxes 可携带办法/条目；保留成配对键，避免把不同样本的办法和条目交叉组合。
            public HashSet<string> LocalContextKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, HashSet<string>> LocalContextKeysByTarget =
                new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        }

        private sealed class SmartEntryStat
        {
            public string EntryCode; public string EntryName; public int ProjectCount;
            public bool CurrentMethodEvidence;
        }

        private sealed class SmartMapCandidateScore
        {
            public SmartMapEntry Entry;
            public string EntryCode;
            public string EntryName;
            public long EntrySeq;
            public bool HasEntry;
            public bool HasCurrentContext;
            public bool HasCurrentMethodMapping;
            public bool PrefixMatch;
            public bool CurrentTargetsValid;
            public bool EntryCandidatesTruncated;
            public List<SmartTargetEntryResolution> TargetEntries = new List<SmartTargetEntryResolution>();
        }

        private sealed class SmartTargetEntryResolution
        {
            public SmartBoxTarget Target;
            public string EntryCode;
            public string EntryName;
            public long EntrySeq;
            public bool FromCurrentContext;
            public string Issue;
            public int EvidenceScore;
        }

        private sealed class SmartLearningScope
        {
            public string Kind;
            public string EntryCode;
            public string DisplayName;

            public static SmartLearningScope CreateAll()
            {
                return new SmartLearningScope { Kind = "All", EntryCode = "", DisplayName = "全部学习库" };
            }
        }

        private sealed class SmartMethodRoute
        {
            public string RawMethod;
            public string LearningMethod;
            public string LibraryMethod;
            public string MethodNo;
        }

        private sealed class SmartQuotaSource
        {
            public string Db; public long QuotaSeq;
        }

        private sealed class SmartFormulaOperand
        {
            public int Index;
            public string Signature;
            public string Name;
            public string Unit;
        }

        private sealed class SmartFormulaRule
        {
            public string RuleHash;
            public string TargetUnit;
            public string Template;
            public string Method;
            public string EntryCode;
            public int SampleCount;
            public DateTime LastSeen;
            public List<SmartFormulaOperand> Operands = new List<SmartFormulaOperand>();
        }

        private sealed class SmartFormulaEvaluation
        {
            public SmartFormulaRule Rule;
            public string QuantityText;
        }

        private sealed class SmartLearningSnapshot
        {
            public bool FromSql;
            public string Method = "";
            public string SoftwarePartition = "";
            public string MethodNo = "";
            public SmartLearningScope SelectedScope = SmartLearningScope.CreateAll();
            public Dictionary<string, List<SmartMapEntry>> BySignature =
                new Dictionary<string, List<SmartMapEntry>>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, List<SmartMapEntry>> ByNameOnly =
                new Dictionary<string, List<SmartMapEntry>>(StringComparer.OrdinalIgnoreCase);
            public List<KeyValuePair<string, MatchTextFeatures>> NameFeatures =
                new List<KeyValuePair<string, MatchTextFeatures>>();
            public Dictionary<string, List<SmartEntryStat>> EntryByQuota =
                new Dictionary<string, List<SmartEntryStat>>(StringComparer.OrdinalIgnoreCase);
            // 定额编号 -> 同办法历史项目里的来源行(跨库整行复制用,取最新)
            public Dictionary<string, SmartQuotaSource> CrossSourceByQuota =
                new Dictionary<string, SmartQuotaSource>(StringComparer.OrdinalIgnoreCase);
            // 签名+"\n"+定额编号 -> 该工程量配该定额时历史实际放过的条目(最强条目证据,按样本数降序)
            public Dictionary<string, List<SmartEntryStat>> EntryBySignatureQuota =
                new Dictionary<string, List<SmartEntryStat>>(StringComparer.OrdinalIgnoreCase);
            // 名称签名+目标编号 -> 已确认的单系数或多参数数量公式。
            public Dictionary<string, List<SmartFormulaRule>> FormulaByKey =
                new Dictionary<string, List<SmartFormulaRule>>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> ProjectEntryNameByCode =
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, string> LearningEntryNameByCode =
                new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            public Dictionary<string, HashSet<string>> ScopeEntriesByBox =
                new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        }

        private static readonly object SmartQuotaIndexCacheLock = new object();
        private static string smartQuotaIndexCachePath = "";
        private static long smartQuotaIndexCacheLength = -1;
        private static DateTime smartQuotaIndexCacheWriteTimeUtc = DateTime.MinValue;
        private static Dictionary<string, ProjectQuota> smartQuotaIndexCache =
            new Dictionary<string, ProjectQuota>(StringComparer.OrdinalIgnoreCase);

        private static string SmartNameSegment(string signature)
        {
            int idx = (signature ?? "").LastIndexOf('|');
            return idx >= 0 ? signature.Substring(0, idx) : (signature ?? "");
        }

        private static string NormalizeSmartProjectMethod(string method)
        {
            string text = (method ?? "").Trim();
            if (text.IndexOf("TB10801", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("2024", StringComparison.OrdinalIgnoreCase) >= 0) return "2024";
            if (text.IndexOf("国铁科法", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("30号文", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("101号文", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("101-estimate", StringComparison.OrdinalIgnoreCase) >= 0 ||
                text.IndexOf("2020", StringComparison.OrdinalIgnoreCase) >= 0) return "2020";
            return text;
        }

        private static SmartMethodRoute ResolveSmartMethodRoute(string rawMethod)
        {
            string raw = (rawMethod ?? "").Trim();
            string learningMethod = NormalizeSmartProjectMethod(raw);
            if (raw.IndexOf("101号文", StringComparison.OrdinalIgnoreCase) >= 0 ||
                raw.IndexOf("101-estimate", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return new SmartMethodRoute
                {
                    RawMethod = raw,
                    LearningMethod = "2020",
                    LibraryMethod = "101-estimate",
                    MethodNo = SmartMethodNo101Estimate
                };
            }
            if (String.Equals(learningMethod, "2024", StringComparison.OrdinalIgnoreCase))
            {
                return new SmartMethodRoute
                {
                    RawMethod = raw,
                    LearningMethod = "2024",
                    LibraryMethod = "2024",
                    MethodNo = SmartMethodNo2024
                };
            }
            if (String.Equals(learningMethod, "2020", StringComparison.OrdinalIgnoreCase))
            {
                return new SmartMethodRoute
                {
                    RawMethod = raw,
                    LearningMethod = "2020",
                    LibraryMethod = "2020",
                    MethodNo = SmartMethodNo30
                };
            }
            return new SmartMethodRoute
            {
                RawMethod = raw,
                LearningMethod = learningMethod,
                LibraryMethod = learningMethod,
                MethodNo = ""
            };
        }

        private static string NormalizeSmartLearningSignature(string signature)
        {
            string raw = signature ?? "";
            int idx = raw.LastIndexOf('|');
            string name = idx >= 0 ? raw.Substring(0, idx) : raw;
            string unit = idx >= 0 ? raw.Substring(idx + 1) : "";
            if (unit.Length > 0 && name.Length > unit.Length && name.EndsWith(unit, StringComparison.OrdinalIgnoreCase))
            {
                name = name.Substring(0, name.Length - unit.Length);
            }
            return NormalizeForSignature(name) + "|";
        }

        private static string BuildSmartFormulaKey(string signature, string kind, string code)
        {
            return NormalizeSmartLearningSignature(signature) + "\n" +
                (String.IsNullOrWhiteSpace(kind) ? "quota" : kind.Trim().ToLowerInvariant()) + ":" + (code ?? "").Trim().ToUpperInvariant();
        }

        private static SmartFormulaRule AddSmartFormulaRule(SmartLearningSnapshot snapshot, string ruleHash, string signature,
            string kind, string code, string targetUnit, string formulaTemplate, string method, string entryCode, int sampleCount, DateTime lastSeen)
        {
            if (snapshot == null || String.IsNullOrWhiteSpace(code) || String.IsNullOrWhiteSpace(formulaTemplate)) return null;
            string key = BuildSmartFormulaKey(signature, kind, code);
            List<SmartFormulaRule> rules;
            if (!snapshot.FormulaByKey.TryGetValue(key, out rules))
            {
                rules = new List<SmartFormulaRule>();
                snapshot.FormulaByKey[key] = rules;
            }
            SmartFormulaRule existing = rules.FirstOrDefault(rule => String.Equals(rule.RuleHash, ruleHash, StringComparison.OrdinalIgnoreCase));
            if (existing == null)
            {
                existing = new SmartFormulaRule
                {
                    RuleHash = ruleHash ?? "",
                    TargetUnit = targetUnit ?? "",
                    Template = formulaTemplate ?? "",
                    Method = NormalizeSmartProjectMethod(method),
                    EntryCode = entryCode ?? "",
                    SampleCount = Math.Max(1, sampleCount),
                    LastSeen = lastSeen
                };
                rules.Add(existing);
                return existing;
            }
            existing.SampleCount = Math.Max(existing.SampleCount, Math.Max(1, sampleCount));
            if (lastSeen > existing.LastSeen) existing.LastSeen = lastSeen;
            return existing;
        }

        private static void UpsertSmartBoxTarget(SmartMapEntry entry, string kind, string code, string name, string unit)
        {
            if (entry == null || String.IsNullOrWhiteSpace(code)) return;
            SmartBoxTarget existing = entry.Targets.FirstOrDefault(target =>
                String.Equals(target.Kind ?? "", kind ?? "", StringComparison.OrdinalIgnoreCase) &&
                String.Equals(target.Code ?? "", code ?? "", StringComparison.OrdinalIgnoreCase));
            if (existing == null)
            {
                entry.Targets.Add(new SmartBoxTarget { Kind = kind, Code = code, Name = name, Unit = unit });
                return;
            }
            if (!String.IsNullOrWhiteSpace(name)) existing.Name = name;
            if (!String.IsNullOrWhiteSpace(unit)) existing.Unit = unit;
        }

        private static string BuildSmartQuantitySignature(string quantityName, string quantityUnit)
        {
            string name = StripTrailingQuantityUnit(quantityName, quantityUnit);
            if (String.IsNullOrWhiteSpace(quantityUnit)) name = StripTrailingKnownSmartQuantityUnit(name);
            string signature = NormalizeForSignature(name) + "|";
            return signature.Length <= 450 ? signature : signature.Substring(0, 450);
        }

        // 兼容旧流水 quantity_unit 为空、单位仍留在 raw_name 尾部的情况。
        // 只在“空白边界 + 已知单位”时剥离，不能把设备线夹等名称尾字误当单位。
        private static string StripTrailingKnownSmartQuantityUnit(string quantityName)
        {
            string name = (quantityName ?? "").TrimEnd();
            int boundary = name.Length - 1;
            while (boundary >= 0 && !Char.IsWhiteSpace(name[boundary])) boundary--;
            if (boundary < 0) return name;
            string suffix = name.Substring(boundary + 1).Trim();
            string prefix = name.Substring(0, boundary).TrimEnd();
            return prefix.Length > 0 && suffix.Length > 0 && LooksLikeExcelLinkUnit(suffix) ? prefix : name;
        }

        private static string ResolveSmartSqlSignature(string storedSignature, Dictionary<string, string> legacyAliases)
        {
            string mapped;
            if (legacyAliases != null && legacyAliases.TryGetValue(storedSignature ?? "", out mapped)) return mapped;
            string normalized = NormalizeSmartLearningSignature(storedSignature);
            return legacyAliases != null && legacyAliases.TryGetValue(normalized, out mapped) ? mapped : normalized;
        }

        // 保留项目信息.编制办法文号的原始办法身份；学习与参考库分区由 ResolveSmartMethodRoute 分别解析。
        private static string SmartResolveProjectMethod(SqlConnection projectConn)
        {
            try
            {
                using (SqlCommand cmd = projectConn.CreateCommand())
                {
                    cmd.CommandText = "SELECT TOP 1 编制办法文号 FROM 项目信息";
                    object raw = cmd.ExecuteScalar();
                    string text = raw != null ? raw.ToString() : "";
                    return text.Trim();
                }
            }
            catch (Exception ex)
            {
                Log("Smart fill resolve method failed: " + ex.Message);
                return "";
            }
        }

        // 目标项目条目表:条目编号 -> 条目序号(写入定位用)。
        private static Dictionary<string, long> LoadSmartProjectEntries(SqlConnection projectConn,
            out Dictionary<string, string> entryNames)
        {
            Dictionary<string, long> map = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
            entryNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (SqlCommand cmd = projectConn.CreateCommand())
                {
                    cmd.CommandText = "SELECT 条目编号, 条目序号, 工程或费用项目名称 FROM 章节表 WHERE 条目编号 IS NOT NULL";
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string code = (reader.IsDBNull(0) ? "" : reader.GetValue(0).ToString()).Trim();
                            if (code.Length == 0 || map.ContainsKey(code)) continue;
                            long seq;
                            if (Int64.TryParse(reader.GetValue(1).ToString(), out seq))
                            {
                                map[code] = seq;
                                entryNames[code] = (reader.IsDBNull(2) ? "" : reader.GetValue(2).ToString()).Trim();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("Smart fill load project entries failed: " + ex.Message);
            }
            return map;
        }

        private static List<SmartLearningScope> LoadSmartLearningScopes(Form mainForm)
        {
            List<SmartLearningScope> result = new List<SmartLearningScope> { SmartLearningScope.CreateAll() };
            if (IsLearningDbCircuitOpen())
            {
                result.Add(new SmartLearningScope { Kind = "Unclassified", EntryCode = "", DisplayName = "未归类" });
                return result;
            }
            string method = "";
            SmartMethodRoute route = ResolveSmartMethodRoute("");
            Dictionary<string, string> projectNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            bool learningDbAccessStarted = false;
            try
            {
                SqlConnection projectConn = GetOpenProjectConnection(mainForm);
                method = SmartResolveProjectMethod(projectConn);
                route = ResolveSmartMethodRoute(method);
                Dictionary<string, long> ignored = LoadSmartProjectEntries(projectConn, out projectNames);
                learningDbAccessStarted = true;
                using (SqlConnection conn = new SqlConnection(GetLearningDbConnectionString()))
                {
                    conn.Open();
                    string softwarePartition = ResolveLearningSoftwarePartition();
                    string normalizedMethodNo = NormalizeLearningMethodNo(route.MethodNo);
                    if (!IsValidLearningSoftwarePartition(softwarePartition) || String.IsNullOrEmpty(normalizedMethodNo))
                        throw new InvalidOperationException("学习分区或办法号无法识别");
                    HashSet<string> entryCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandTimeout = 15;
                        cmd.CommandText = "SELECT DISTINCT entry_code FROM dbo.EngineeringTemplate WHERE software_partition=@software_partition AND method_no=@method_no AND entry_code<>''";
                        cmd.Parameters.AddWithValue("@software_partition", softwarePartition);
                        cmd.Parameters.AddWithValue("@method_no", normalizedMethodNo);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string entryCode = reader.GetString(0).Trim();
                                if (IsSmartClassifiedEntryCode(entryCode)) entryCodes.Add(entryCode);
                            }
                        }
                    }
                    Dictionary<string, string> learningNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandTimeout = 15;
                        cmd.CommandText = "SELECT entry_code,entry_name FROM dbo.ChapterEntry WHERE method=@library_method AND method_no=@method_no";
                        cmd.Parameters.AddWithValue("@library_method", route.LibraryMethod);
                        cmd.Parameters.AddWithValue("@method_no", route.MethodNo);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read()) learningNames[reader.GetString(0).Trim()] = reader.IsDBNull(1) ? "" : reader.GetString(1).Trim();
                        }
                    }
                    result.Add(new SmartLearningScope { Kind = "Unclassified", EntryCode = "", DisplayName = "未归类" });
                    HashSet<string> scopeCodes = BuildSmartLearningScopeCodes(entryCodes);
                    foreach (string code in scopeCodes.OrderBy(value => value, StringComparer.OrdinalIgnoreCase))
                    {
                        string name;
                        if (!projectNames.TryGetValue(code, out name) || String.IsNullOrWhiteSpace(name))
                            learningNames.TryGetValue(code, out name);
                        result.Add(new SmartLearningScope
                        {
                            Kind = "Entry",
                            EntryCode = code,
                            DisplayName = String.IsNullOrWhiteSpace(name) ? code : name.Trim()
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                if (learningDbAccessStarted) ObserveLearningDbFailure(ex);
                Log("Smart fill load learning scopes failed: " + ex.Message);
                if (!result.Any(scope => String.Equals(scope.Kind, "Unclassified", StringComparison.OrdinalIgnoreCase)))
                    result.Add(new SmartLearningScope { Kind = "Unclassified", EntryCode = "", DisplayName = "未归类" });
            }
            return result;
        }

        private static bool IsSmartClassifiedEntryCode(string entryCode)
        {
            string code = (entryCode ?? "").Trim();
            return code.Length >= 2 && Char.IsDigit(code[0]) && Char.IsDigit(code[1]) &&
                code.All(value => Char.IsDigit(value) || value == '-');
        }

        private static HashSet<string> BuildSmartLearningScopeCodes(IEnumerable<string> entryCodes)
        {
            HashSet<string> result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string value in entryCodes ?? Enumerable.Empty<string>())
            {
                string entryCode = (value ?? "").Trim();
                if (!IsSmartClassifiedEntryCode(entryCode)) continue;
                result.Add(entryCode.Substring(0, 2));
                if (entryCode.Length >= 4) result.Add(entryCode.Substring(0, 4));
            }
            return result;
        }

        // 从 RecoLearning 加载快照;失败回退 jsonl(仅签名映射,无条目知识)。
        private static SmartLearningSnapshot LoadSmartLearningSnapshot(string learningMethod, string libraryMethod,
            string methodNo, out string note)
        {
            note = null;
            SmartLearningSnapshot snapshot = new SmartLearningSnapshot
            {
                Method = NormalizeSmartProjectMethod(learningMethod),
                SoftwarePartition = ResolveLearningSoftwarePartition(),
                MethodNo = NormalizeLearningMethodNo(methodNo)
            };
            if (!IsValidLearningSoftwarePartition(snapshot.SoftwarePartition) || String.IsNullOrEmpty(snapshot.MethodNo))
            {
                note = "当前软件分区或编制办法无法识别，已禁止读取学习关系。";
                return snapshot;
            }
            try
            {
                if (IsLearningDbCircuitOpen()) throw new InvalidOperationException("学习库熔断中");
                string connectionString = GetLearningDbConnectionString();
                if (String.IsNullOrEmpty(connectionString)) throw new InvalidOperationException("学习库连接串为空");
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    // 老聚合签名可能把行尾单位嵌进名称（如 HRB400钢筋KG|）。
                    // QuantityAlias 保留了原始名称和独立单位，可在不改历史数据时桥接到名称级签名。
                    Dictionary<string, string> legacySignatureAliases = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandTimeout = 15;
                        cmd.CommandText = "SELECT signature,raw_name,quantity_unit FROM dbo.QuantityAlias";
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string stored = reader.IsDBNull(0) ? "" : reader.GetString(0);
                                string canonical = BuildSmartQuantitySignature(reader.IsDBNull(1) ? "" : reader.GetString(1),
                                    reader.IsDBNull(2) ? "" : reader.GetString(2));
                                if (!String.IsNullOrWhiteSpace(stored) && !String.IsNullOrWhiteSpace(canonical.Trim('|')))
                                {
                                    legacySignatureAliases[stored] = canonical;
                                }
                            }
                        }
                    }
                    Dictionary<string, SmartMapEntry> byKey = new Dictionary<string, SmartMapEntry>(StringComparer.OrdinalIgnoreCase);
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandTimeout = 15;
                        cmd.CommandText =
                            "SELECT m.signature, m.box_id, m.method, m.weight, m.accepted_count, m.corrected_count, m.rejected_count, " +
                            "t.target_kind, t.target_code, t.target_name, t.target_unit, m.last_used_at " +
                            "FROM dbo.SignatureBoxMap m JOIN dbo.QuotaBoxTarget t ON t.box_id = m.box_id " +
                            "WHERE m.weight > 0 AND m.software_partition=@software_partition " +
                            "ORDER BY m.weight DESC, m.box_id";
                        cmd.Parameters.AddWithValue("@software_partition", snapshot.SoftwarePartition);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string signature = ResolveSmartSqlSignature(reader.GetString(0), legacySignatureAliases);
                                string boxId = reader.GetString(1);
                                string mappingMethod = reader.IsDBNull(2) ? "" : NormalizeSmartProjectMethod(reader.GetString(2));
                                bool currentMethodMapping = true;
                                string key = signature + "\n" + boxId;
                                SmartMapEntry entry;
                                if (!byKey.TryGetValue(key, out entry))
                                {
                                    entry = new SmartMapEntry
                                    {
                                        BoxId = boxId,
                                        Weight = reader.GetInt32(3),
                                        AcceptedCount = reader.GetInt32(4),
                                        CorrectedCount = reader.GetInt32(5),
                                        RejectedCount = reader.GetInt32(6),
                                        CurrentMethodMapping = currentMethodMapping,
                                        LastUsedAt = reader.IsDBNull(11) ? DateTime.MinValue : reader.GetDateTime(11)
                                    };
                                    byKey[key] = entry;
                                    List<SmartMapEntry> list;
                                    if (!snapshot.BySignature.TryGetValue(signature, out list))
                                    {
                                        list = new List<SmartMapEntry>();
                                        snapshot.BySignature[signature] = list;
                                    }
                                    list.Add(entry);
                                    string nameSeg = SmartNameSegment(signature);
                                    if (!snapshot.ByNameOnly.TryGetValue(nameSeg, out list))
                                    {
                                        list = new List<SmartMapEntry>();
                                        snapshot.ByNameOnly[nameSeg] = list;
                                    }
                                    list.Add(entry);
                                }
                                else
                                {
                                    if (currentMethodMapping && !entry.CurrentMethodMapping)
                                    {
                                        entry.Weight = reader.GetInt32(3);
                                        entry.AcceptedCount = reader.GetInt32(4);
                                        entry.CorrectedCount = reader.GetInt32(5);
                                        entry.RejectedCount = reader.GetInt32(6);
                                    }
                                    else if (currentMethodMapping == entry.CurrentMethodMapping) entry.Weight = Math.Max(entry.Weight, reader.GetInt32(3));
                                    if (currentMethodMapping) entry.CurrentMethodMapping = true;
                                    DateTime lastUsed = reader.IsDBNull(11) ? DateTime.MinValue : reader.GetDateTime(11);
                                    if (lastUsed > entry.LastUsedAt) entry.LastUsedAt = lastUsed;
                                }
                                UpsertSmartBoxTarget(entry, reader.GetString(7), reader.GetString(8), reader.GetString(9), reader.GetString(10));
                            }
                        }
                    }

                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandTimeout = 15;
                        cmd.CommandText =
                            "SELECT box_id,entry_code FROM dbo.EngineeringTemplate WHERE software_partition=@software_partition AND method_no=@method_no";
                        cmd.Parameters.AddWithValue("@software_partition", snapshot.SoftwarePartition);
                        cmd.Parameters.AddWithValue("@method_no", snapshot.MethodNo);
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string boxId = reader.GetString(0).Trim();
                                string entryCode = reader.GetString(1).Trim();
                                if (!IsSmartClassifiedEntryCode(entryCode)) continue;
                                HashSet<string> codes;
                                if (!snapshot.ScopeEntriesByBox.TryGetValue(boxId, out codes))
                                {
                                    codes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                                    snapshot.ScopeEntriesByBox[boxId] = codes;
                                }
                                if (entryCode.Length > 0) codes.Add(entryCode);
                            }
                        }
                    }
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandTimeout = 15;
                        cmd.CommandText = "SELECT entry_code,entry_name FROM dbo.ChapterEntry WHERE method=@library_method AND method_no=@method_no";
                        cmd.Parameters.AddWithValue("@library_method", libraryMethod ?? "");
                        cmd.Parameters.AddWithValue("@method_no", methodNo ?? "");
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string entryCode = reader.GetString(0).Trim();
                                if (entryCode.Length > 0) snapshot.LearningEntryNameByCode[entryCode] =
                                    reader.IsDBNull(1) ? "" : reader.GetString(1).Trim();
                            }
                        }
                    }

                    using (SqlCommand exists = conn.CreateCommand())
                    {
                        exists.CommandTimeout = 15;
                        exists.CommandText = "SELECT CASE WHEN OBJECT_ID('dbo.QuantityFormulaRule','U') IS NULL OR OBJECT_ID('dbo.QuantityFormulaOperand','U') IS NULL THEN 0 ELSE 1 END";
                        if (Convert.ToInt32(exists.ExecuteScalar(), CultureInfo.InvariantCulture) != 0)
                        {
                            Dictionary<string, SmartFormulaRule> formulaByHash = new Dictionary<string, SmartFormulaRule>(StringComparer.OrdinalIgnoreCase);
                            using (SqlCommand cmd = conn.CreateCommand())
                            {
                                cmd.CommandTimeout = 15;
                                cmd.CommandText =
                                    "SELECT rule_hash,anchor_signature,target_kind,target_code,target_unit,formula_template,method,entry_code,sample_count,last_seen " +
                                    "FROM dbo.QuantityFormulaRule WHERE software_partition=@software_partition AND method_no=@method_no";
                                cmd.Parameters.AddWithValue("@software_partition", snapshot.SoftwarePartition);
                                cmd.Parameters.AddWithValue("@method_no", snapshot.MethodNo);
                                using (SqlDataReader reader = cmd.ExecuteReader())
                                {
                                    while (reader.Read())
                                    {
                                        SmartFormulaRule rule = AddSmartFormulaRule(snapshot, reader.GetString(0), ResolveSmartSqlSignature(reader.GetString(1), legacySignatureAliases), reader.GetString(2),
                                            reader.GetString(3), reader.GetString(4), reader.GetString(5), reader.GetString(6), reader.GetString(7),
                                            reader.GetInt32(8), reader.IsDBNull(9) ? DateTime.MinValue : reader.GetDateTime(9));
                                        if (rule != null) formulaByHash[reader.GetString(0)] = rule;
                                    }
                                }
                            }
                            using (SqlCommand cmd = conn.CreateCommand())
                            {
                                cmd.CommandTimeout = 15;
                                cmd.CommandText =
                                    "SELECT o.rule_hash,o.operand_index,o.operand_signature,o.operand_name,o.operand_unit " +
                                    "FROM dbo.QuantityFormulaOperand o JOIN dbo.QuantityFormulaRule r ON r.rule_hash=o.rule_hash " +
                                    "WHERE r.software_partition=@software_partition AND r.method_no=@method_no ORDER BY o.rule_hash,o.operand_index";
                                cmd.Parameters.AddWithValue("@software_partition", snapshot.SoftwarePartition);
                                cmd.Parameters.AddWithValue("@method_no", snapshot.MethodNo);
                                using (SqlDataReader reader = cmd.ExecuteReader())
                                {
                                    while (reader.Read())
                                    {
                                        SmartFormulaRule rule;
                                        if (!formulaByHash.TryGetValue(reader.GetString(0), out rule)) continue;
                                        rule.Operands.Add(new SmartFormulaOperand
                                        {
                                            Index = reader.GetInt32(1),
                                            Signature = BuildSmartQuantitySignature(reader.GetString(3), reader.GetString(4)),
                                            Name = reader.GetString(3),
                                            Unit = reader.GetString(4)
                                        });
                                    }
                                }
                            }
                        }
                    }

                    if (!String.IsNullOrEmpty(snapshot.Method))
                    {
                        int basePartitionCount;
                        using (SqlCommand count = conn.CreateCommand())
                        {
                            count.CommandTimeout = 15;
                            count.CommandText =
                                "SELECT COUNT(*) FROM dbo.EntryQuota " +
                                "WHERE method=@library_method AND method_no=@method_no AND target_kind='quota'";
                            count.Parameters.AddWithValue("@library_method", libraryMethod ?? "");
                            count.Parameters.AddWithValue("@method_no", methodNo ?? "");
                            basePartitionCount = Convert.ToInt32(count.ExecuteScalar(), CultureInfo.InvariantCulture);
                        }
                        if (ShouldWarnSmartLibraryPartitionMissing(basePartitionCount))
                        {
                            Log("Smart fill exact EntryQuota partition is empty: method=" + (libraryMethod ?? "") +
                                ", method_no=" + (methodNo ?? "") + "; continuing with empty library evidence and no fallback.");
                        }
                        using (SqlCommand cmd = conn.CreateCommand())
                        {
                            cmd.CommandTimeout = 15;
                            cmd.CommandText =
                                "SELECT quota_code, entry_code, entry_name, project_count FROM dbo.EntryQuota q " +
                                "WHERE q.method=@library_method AND q.method_no=@method_no AND q.target_kind='quota' AND EXISTS " +
                                "(SELECT 1 FROM dbo.QuotaBoxTarget t WHERE t.target_code = q.quota_code AND t.target_kind = 'quota')";
                            cmd.Parameters.AddWithValue("@library_method", libraryMethod ?? "");
                            cmd.Parameters.AddWithValue("@method_no", methodNo ?? "");
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    string quotaCode = reader.GetString(0);
                                    List<SmartEntryStat> stats;
                                    if (!snapshot.EntryByQuota.TryGetValue(quotaCode, out stats))
                                    {
                                        stats = new List<SmartEntryStat>();
                                        snapshot.EntryByQuota[quotaCode] = stats;
                                    }
                                    stats.Add(new SmartEntryStat
                                    {
                                        EntryCode = reader.GetString(1),
                                        EntryName = reader.GetString(2),
                                        ProjectCount = reader.GetInt32(3)
                                    });
                                }
                            }
                        }
                        // 签名级条目证据:该工程量配该定额历史上实际放过的条目。
                        using (SqlCommand cmd = conn.CreateCommand())
                        {
                            cmd.CommandTimeout = 15;
                            cmd.CommandText =
                                "SELECT signature, target_code, method, entry_code, entry_name, sample_count FROM dbo.SignatureEntryMap " +
                                "WHERE software_partition=@software_partition AND method_no=@method_no";
                            cmd.Parameters.AddWithValue("@software_partition", snapshot.SoftwarePartition);
                            cmd.Parameters.AddWithValue("@method_no", snapshot.MethodNo);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    string key = ResolveSmartSqlSignature(reader.GetString(0), legacySignatureAliases) + "\n" + reader.GetString(1);
                                    List<SmartEntryStat> stats;
                                    if (!snapshot.EntryBySignatureQuota.TryGetValue(key, out stats))
                                    {
                                        stats = new List<SmartEntryStat>();
                                        snapshot.EntryBySignatureQuota[key] = stats;
                                    }
                                    stats.Add(new SmartEntryStat
                                    {
                                        EntryCode = reader.GetString(3),
                                        EntryName = reader.GetString(4),
                                        ProjectCount = 10000 + reader.GetInt32(5),
                                        CurrentMethodEvidence = true
                                    });
                                }
                            }
                        }
                        foreach (List<SmartEntryStat> stats in snapshot.EntryBySignatureQuota.Values)
                        {
                            stats.Sort(delegate(SmartEntryStat a, SmartEntryStat b)
                            {
                                int methodCompare = b.CurrentMethodEvidence.CompareTo(a.CurrentMethodEvidence);
                                return methodCompare != 0 ? methodCompare : b.ProjectCount.CompareTo(a.ProjectCount);
                            });
                        }

                        // 跨库复制溯源:同办法历史绑定的来源库与源定额行(ORDER BY id,后者覆盖=取最新)。
                        using (SqlCommand cmd = conn.CreateCommand())
                        {
                            cmd.CommandTimeout = 15;
                            cmd.CommandText =
                                "SELECT target_code, project_id, extra FROM dbo.BindingLog " +
                                "WHERE source = 'import:excel-links' AND project_id <> '' AND target_kind = 'quota' " +
                                "AND software_partition=@software_partition AND method_no=@method_no ORDER BY id";
                            cmd.Parameters.AddWithValue("@software_partition", snapshot.SoftwarePartition);
                            cmd.Parameters.AddWithValue("@method_no", snapshot.MethodNo);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    string quotaCode = reader.GetString(0);
                                    string sourceDb = reader.GetString(1);
                                    string extra = reader.IsDBNull(2) ? "" : reader.GetString(2);
                                    long quotaSeq;
                                    if (!Int64.TryParse(GetFlat(ParseFlatJson(extra), "quota_sequence"), out quotaSeq) || quotaSeq <= 0) continue;
                                    snapshot.CrossSourceByQuota[quotaCode] = new SmartQuotaSource { Db = sourceDb, QuotaSeq = quotaSeq };
                                }
                            }
                        }

                        // 绑定流水里的条目证据:同定额在历史绑定中实际放过的条目,权重高于扫描共现。
                        using (SqlCommand cmd = conn.CreateCommand())
                        {
                            cmd.CommandTimeout = 15;
                            cmd.CommandText =
                                "SELECT target_code, entry_code, MAX(entry_name) AS entry_name, COUNT(*) AS n FROM dbo.BindingLog " +
                                "WHERE entry_code <> '' AND target_kind = 'quota' AND software_partition=@software_partition AND method_no=@method_no " +
                                "GROUP BY target_code, entry_code";
                            cmd.Parameters.AddWithValue("@software_partition", snapshot.SoftwarePartition);
                            cmd.Parameters.AddWithValue("@method_no", snapshot.MethodNo);
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    string quotaCode = reader.GetString(0);
                                    List<SmartEntryStat> stats;
                                    if (!snapshot.EntryByQuota.TryGetValue(quotaCode, out stats))
                                    {
                                        stats = new List<SmartEntryStat>();
                                        snapshot.EntryByQuota[quotaCode] = stats;
                                    }
                                    stats.Add(new SmartEntryStat
                                    {
                                        EntryCode = reader.GetString(1),
                                        EntryName = reader.GetString(2),
                                        ProjectCount = 1000 + reader.GetInt32(3)
                                    });
                                }
                            }
                        }

                        foreach (List<SmartEntryStat> stats in snapshot.EntryByQuota.Values)
                        {
                            stats.Sort(delegate(SmartEntryStat a, SmartEntryStat b) { return b.ProjectCount.CompareTo(a.ProjectCount); });
                        }
                    }
                }

                foreach (List<SmartMapEntry> list in snapshot.BySignature.Values)
                {
                    list.Sort(delegate(SmartMapEntry a, SmartMapEntry b) { return b.Weight.CompareTo(a.Weight); });
                }
                foreach (List<SmartMapEntry> list in snapshot.ByNameOnly.Values)
                {
                    list.Sort(delegate(SmartMapEntry a, SmartMapEntry b) { return b.Weight.CompareTo(a.Weight); });
                }
                foreach (string nameSeg in snapshot.ByNameOnly.Keys)
                {
                    snapshot.NameFeatures.Add(new KeyValuePair<string, MatchTextFeatures>(nameSeg, BuildMatchTextFeatures(nameSeg)));
                }
                snapshot.FromSql = true;
                ClearLearningDbCircuit();
                return snapshot;
            }
            catch (Exception ex)
            {
                ObserveLearningDbFailure(ex);
                Log("Smart fill SQL snapshot failed; local learning is disabled: " + ex.Message);
                note = "学习库(SQL)不可用，已禁止本地配对：" + ex.Message;
                return snapshot;
            }
        }

        private static bool ShouldWarnSmartLibraryPartitionMissing(int basePartitionCount)
        {
            return basePartitionCount == 0;
        }

        // 推荐数量只使用当前 Excel 单位和当前运行版本的定额元数据；SQL target_unit 仅作历史描述。
        // 定额索引按 path+length+mtime 缓存；项目库仍每次实时读取，避免把运行中修改的项目数据缓存住。
        private static Dictionary<string, ProjectQuota> LoadCachedSmartQuotaIndex(string path)
        {
            if (String.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return new Dictionary<string, ProjectQuota>(StringComparer.OrdinalIgnoreCase);
            FileInfo info = new FileInfo(path);
            string fullPath = info.FullName;
            long length = info.Length;
            DateTime writeTimeUtc = info.LastWriteTimeUtc;
            lock (SmartQuotaIndexCacheLock)
            {
                if (String.Equals(smartQuotaIndexCachePath, fullPath, StringComparison.OrdinalIgnoreCase) &&
                    smartQuotaIndexCacheLength == length && smartQuotaIndexCacheWriteTimeUtc == writeTimeUtc)
                {
                    return smartQuotaIndexCache;
                }

                Dictionary<string, ProjectQuota> loaded = new Dictionary<string, ProjectQuota>(StringComparer.OrdinalIgnoreCase);
                foreach (string line in File.ReadLines(fullPath, System.Text.Encoding.UTF8))
                {
                    Dictionary<string, string> row = ParseFlatJson(line);
                    string code = GetFlat(row, "quota_code").Trim();
                    if (code.Length == 0 || GetFlat(row, "is_current") == "0" || loaded.ContainsKey(code)) continue;
                    ProjectQuota quota = new ProjectQuota
                    {
                        Code = code,
                        Name = GetFlat(row, "quota_name").Trim(),
                        Unit = GetFlat(row, "quota_unit").Trim(),
                        IsLibrary = true
                    };
                    quota.NormCode = NormalizeMatchText(quota.Code);
                    quota.NormName = NormalizeMatchText(quota.Name);
                    loaded[code] = quota;
                }
                smartQuotaIndexCachePath = fullPath;
                smartQuotaIndexCacheLength = length;
                smartQuotaIndexCacheWriteTimeUtc = writeTimeUtc;
                smartQuotaIndexCache = loaded;
                return smartQuotaIndexCache;
            }
        }

        private static Dictionary<string, ProjectQuota> LoadCurrentSmartQuotaMetadata(Form mainForm, SmartLearningSnapshot snapshot)
        {
            HashSet<string> requiredCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (List<SmartMapEntry> entries in snapshot.BySignature.Values)
            {
                foreach (SmartMapEntry entry in entries)
                {
                    foreach (SmartBoxTarget target in entry.Targets)
                    {
                        if (target != null && String.Equals(target.Kind ?? "quota", "quota", StringComparison.OrdinalIgnoreCase) &&
                            !String.IsNullOrWhiteSpace(target.Code)) requiredCodes.Add(target.Code.Trim());
                    }
                }
            }

            Dictionary<string, ProjectQuota> result = new Dictionary<string, ProjectQuota>(StringComparer.OrdinalIgnoreCase);
            foreach (ProjectQuota quota in LoadProjectQuotas(mainForm))
            {
                if (quota == null || !requiredCodes.Contains(quota.Code ?? "")) continue;
                string currentKey = BuildSmartCurrentQuotaKey(quota.Code, quota.Name, quota.Unit);
                if (!result.ContainsKey(currentKey)) result[currentKey] = quota;
            }

            try
            {
                string dataDir = FindRecoQuotaDataDir();
                string partitionPath = Path.Combine(dataDir,
                    String.Equals(snapshot.SoftwarePartition, "2020", StringComparison.OrdinalIgnoreCase)
                        ? "quota-index-2020.jsonl"
                        : "quota-index-2024.jsonl");
                string path = File.Exists(partitionPath) ? partitionPath : Path.Combine(dataDir, "quota-index.jsonl");
                foreach (KeyValuePair<string, ProjectQuota> pair in LoadCachedSmartQuotaIndex(path))
                {
                    string code = pair.Key;
                    ProjectQuota indexed = pair.Value;
                    if (!requiredCodes.Contains(code) || indexed == null || IsContextSensitiveLearningCode(code)) continue;
                    ProjectQuota existing;
                    if (result.TryGetValue(code, out existing))
                    {
                        if (String.IsNullOrWhiteSpace(existing.Name)) existing.Name = indexed.Name;
                        if (String.IsNullOrWhiteSpace(existing.Unit)) existing.Unit = indexed.Unit;
                        existing.NormCode = NormalizeMatchText(existing.Code);
                        existing.NormName = NormalizeMatchText(existing.Name);
                        continue;
                    }
                    result[code] = indexed;
                }
            }
            catch (Exception ex)
            {
                Log("Smart fill current quota metadata failed: " + ex.Message);
            }
            return result;
        }

        private static string BuildSmartCurrentQuotaKey(string code, string name, string unit)
        {
            return IsContextSensitiveLearningCode(code)
                ? BuildLearningTargetIdentityKey("quota", code, name, unit)
                : (code ?? "").Trim().ToUpperInvariant();
        }

        private static bool TryGetCurrentSmartQuota(Dictionary<string, ProjectQuota> currentQuotaByCode,
            SmartBoxTarget target, out ProjectQuota currentQuota)
        {
            currentQuota = null;
            if (currentQuotaByCode == null || target == null || String.IsNullOrWhiteSpace(target.Code)) return false;
            string key = BuildSmartCurrentQuotaKey(target.Code, target.Name, target.Unit);
            return currentQuotaByCode.TryGetValue(key, out currentQuota) && currentQuota != null;
        }

        private static bool IsSmartMapEntryUsableInCurrentProject(SmartMapEntry entry, string entryName,
            Dictionary<string, ProjectQuota> currentQuotaByCode)
        {
            List<SmartBoxTarget> targets = entry == null
                ? new List<SmartBoxTarget>()
                : entry.Targets.Where(target => target != null && !String.IsNullOrWhiteSpace(target.Code)).ToList();
            if (!IsSmartTargetSetRecommendable(targets)) return false;
            foreach (SmartBoxTarget target in targets)
            {
                if (!String.Equals(target.Kind ?? "quota", "quota", StringComparison.OrdinalIgnoreCase)) return false;
                bool contextSensitive = IsContextSensitiveLearningCode(target.Code);
                if (contextSensitive && (String.IsNullOrWhiteSpace(target.Name) || String.IsNullOrWhiteSpace(target.Unit))) return false;
                ProjectQuota currentQuota;
                if (!TryGetCurrentSmartQuota(currentQuotaByCode, target, out currentQuota) ||
                    String.IsNullOrWhiteSpace(currentQuota.Unit)) return false;
            }
            return true;
        }

        private static bool IsSmartTargetSetRecommendable(IEnumerable<SmartBoxTarget> targets)
        {
            List<SmartBoxTarget> list = (targets ?? Enumerable.Empty<SmartBoxTarget>())
                .Where(target => target != null && !String.IsNullOrWhiteSpace(target.Code)).ToList();
            if (list.Count == 0) return false;
            return list.Any(target => IsPrimaryLearningTarget(target.Kind, target.Code)) ||
                list.All(target =>
                    String.Equals(String.IsNullOrWhiteSpace(target.Kind) ? "quota" : target.Kind.Trim(), "quota",
                        StringComparison.OrdinalIgnoreCase) &&
                    GetLearningBaseTargetCode(target.Code) == "SF");
        }

        // 条目证据必须按目标解析；组件框只负责在全部目标解析完成后做整组判定。
        private static List<SmartTargetEntryResolution> ResolveSmartTargetEntries(SmartLearningSnapshot snapshot,
            Dictionary<string, long> projectEntries, SmartMapEntry mappingEntry, string signature,
            HashSet<string> preferredPrefixes)
        {
            List<SmartTargetEntryResolution> result = new List<SmartTargetEntryResolution>();
            foreach (SmartBoxTarget target in OrderSmartTargets(mappingEntry == null ? null : mappingEntry.Targets))
            {
                result.Add(ResolveSmartTargetEntry(snapshot, projectEntries, mappingEntry, target, signature, preferredPrefixes));
            }

            SmartTargetEntryResolution primary = result.FirstOrDefault(item => item != null && item.Target != null &&
                IsPrimaryLearningTarget(item.Target.Kind, item.Target.Code) && !String.IsNullOrWhiteSpace(item.EntryCode));
            if (primary != null)
            {
                foreach (SmartTargetEntryResolution item in result.Where(item => item != null && item.Target != null &&
                    String.IsNullOrWhiteSpace(item.EntryCode) && IsSmartFollowerTarget(item.Target)))
                {
                    item.EntryCode = primary.EntryCode;
                    item.EntryName = primary.EntryName;
                    item.EntrySeq = primary.EntrySeq;
                    item.FromCurrentContext = false;
                    item.Issue = "";
                }
            }
            return result;
        }

        private static SmartTargetEntryResolution ResolveSmartTargetEntry(SmartLearningSnapshot snapshot,
            Dictionary<string, long> projectEntries, SmartMapEntry mappingEntry, SmartBoxTarget target,
            string signature, HashSet<string> preferredPrefixes)
        {
            SmartTargetEntryResolution result = new SmartTargetEntryResolution { Target = target, EntryCode = "", EntryName = "", Issue = "" };
            if (target == null || String.IsNullOrWhiteSpace(target.Code))
            {
                result.Issue = "目标编号为空";
                return result;
            }

            string[] sigKeys = new string[] { signature ?? "", SmartNameSegment(signature ?? "") + "|" };
            List<SmartEntryStat> currentStats = new List<SmartEntryStat>();
            List<SmartEntryStat> genericStats = new List<SmartEntryStat>();
            bool sawEquipmentEntryForNonSf = false;
            foreach (string sigKey in sigKeys)
            {
                List<SmartEntryStat> stats;
                if (!snapshot.EntryBySignatureQuota.TryGetValue(sigKey + "\n" + (target.Code ?? ""), out stats)) continue;
                foreach (SmartEntryStat stat in stats)
                {
                    if (stat == null || !projectEntries.ContainsKey(stat.EntryCode) ||
                        !SmartEntryCodeMatchesScope(stat.EntryCode, snapshot.SelectedScope)) continue;
                    string resolvedName = ResolveSmartEntryName(snapshot, stat.EntryCode, stat.EntryName);
                    if (GetLearningBaseTargetCode(target.Code) != "SF" &&
                        resolvedName.IndexOf("设备购置费", StringComparison.OrdinalIgnoreCase) >= 0)
                        sawEquipmentEntryForNonSf = true;
                    if (!IsSmartTargetEntryCompatible(target, resolvedName)) continue;
                    if (stat.CurrentMethodEvidence) currentStats.Add(stat);
                    else genericStats.Add(stat);
                }
            }

            bool usePreferredPrefixes = GetLearningBaseTargetCode(target.Code) != "SH";
            currentStats = DistinctAndFilterSmartEntryStats(currentStats, preferredPrefixes, usePreferredPrefixes);
            if (currentStats.Count == 1)
            {
                return CreateSmartTargetEntryResolution(snapshot, projectEntries, target, currentStats[0], true, "");
            }

            string targetIdentity = BuildLearningTargetIdentityKey(target.Kind, target.Code, target.Name, target.Unit);
            HashSet<string> localContexts = null;
            if (mappingEntry != null && mappingEntry.LocalContextKeysByTarget != null)
                mappingEntry.LocalContextKeysByTarget.TryGetValue(targetIdentity, out localContexts);
            if ((localContexts == null || localContexts.Count == 0) && mappingEntry != null &&
                mappingEntry.Targets != null && mappingEntry.Targets.Count == 1)
                localContexts = mappingEntry.LocalContextKeys;
            if (GetLearningBaseTargetCode(target.Code) != "SF" && localContexts != null)
            {
                string methodPrefix = (snapshot.MethodNo ?? "").Trim() + "\n";
                sawEquipmentEntryForNonSf = sawEquipmentEntryForNonSf || localContexts.Any(key =>
                    key.StartsWith(methodPrefix, StringComparison.OrdinalIgnoreCase) &&
                    projectEntries.ContainsKey(key.Substring(methodPrefix.Length)) &&
                    ResolveSmartEntryName(snapshot, key.Substring(methodPrefix.Length), "")
                        .IndexOf("设备购置费", StringComparison.OrdinalIgnoreCase) >= 0);
            }
            List<string> localEntries = SelectSmartLocalEntries(snapshot, projectEntries, target, localContexts,
                preferredPrefixes, usePreferredPrefixes);
            if (localEntries.Count == 1)
            {
                result.EntryCode = localEntries[0];
                result.EntryName = ResolveSmartEntryName(snapshot, result.EntryCode, "");
                result.EntrySeq = projectEntries[result.EntryCode];
                result.FromCurrentContext = true;
                return result;
            }

            if (currentStats.Count > 1)
            {
                return CreateSmartTargetEntryResolution(snapshot, projectEntries, target,
                    currentStats.OrderByDescending(stat => stat.ProjectCount).First(), false, "目标条目证据不唯一");
            }
            if (localEntries.Count > 1)
            {
                result.EntryCode = localEntries[0];
                result.EntryName = ResolveSmartEntryName(snapshot, result.EntryCode, "");
                result.EntrySeq = projectEntries[result.EntryCode];
                result.Issue = "目标条目证据不唯一";
                return result;
            }

            genericStats = DistinctAndFilterSmartEntryStats(genericStats, preferredPrefixes, usePreferredPrefixes);
            SmartEntryStat generic = genericStats.OrderByDescending(stat => stat.ProjectCount).FirstOrDefault();
            if (generic != null)
                return CreateSmartTargetEntryResolution(snapshot, projectEntries, target, generic, false, "");

            List<SmartEntryStat> quotaStats;
            if (String.Equals(target.Kind ?? "quota", "quota", StringComparison.OrdinalIgnoreCase) &&
                snapshot.EntryByQuota.TryGetValue(target.Code ?? "", out quotaStats))
            {
                List<SmartEntryStat> compatible = new List<SmartEntryStat>();
                foreach (SmartEntryStat stat in quotaStats)
                {
                    if (stat == null || !projectEntries.ContainsKey(stat.EntryCode) ||
                        !SmartEntryCodeMatchesScope(stat.EntryCode, snapshot.SelectedScope)) continue;
                    string resolvedName = ResolveSmartEntryName(snapshot, stat.EntryCode, stat.EntryName);
                    if (GetLearningBaseTargetCode(target.Code) != "SF" &&
                        resolvedName.IndexOf("设备购置费", StringComparison.OrdinalIgnoreCase) >= 0)
                        sawEquipmentEntryForNonSf = true;
                    if (IsSmartTargetEntryCompatible(target, resolvedName)) compatible.Add(stat);
                }
                compatible = DistinctAndFilterSmartEntryStats(compatible, preferredPrefixes, usePreferredPrefixes);
                SmartEntryStat best = compatible.OrderByDescending(stat => stat.ProjectCount).FirstOrDefault();
                if (best != null) return CreateSmartTargetEntryResolution(snapshot, projectEntries, target, best, false, "");
            }

            result.Issue = GetLearningBaseTargetCode(target.Code) == "SF"
                ? "SF 必须写入设备购置费条目，当前项目未找到该条目"
                : (sawEquipmentEntryForNonSf
                    ? "设备购置费条目只能写入 SF，当前目标禁止写入"
                    : "学习库未定位到目标项目里的条目");
            return result;
        }

        private static void AddSmartTargetEntryCandidate(Dictionary<string, SmartTargetEntryResolution> candidates,
            SmartLearningSnapshot snapshot, Dictionary<string, long> projectEntries, SmartBoxTarget target,
            string entryCode, string learnedName, bool fromCurrentContext, int evidenceScore)
        {
            string normalizedEntry = LearningPartitionIdentity.NormalizeLearningEntryCode(entryCode);
            if (normalizedEntry.Length == 0 || !projectEntries.ContainsKey(normalizedEntry) ||
                !SmartEntryCodeMatchesScope(normalizedEntry, snapshot.SelectedScope)) return;
            string resolvedName = ResolveSmartEntryName(snapshot, normalizedEntry, learnedName);
            if (!IsSmartTargetEntryCompatible(target, resolvedName)) return;
            SmartTargetEntryResolution existing;
            if (candidates.TryGetValue(normalizedEntry, out existing) && existing.EvidenceScore >= evidenceScore) return;
            candidates[normalizedEntry] = new SmartTargetEntryResolution
            {
                Target = target,
                EntryCode = normalizedEntry,
                EntryName = resolvedName,
                EntrySeq = projectEntries[normalizedEntry],
                FromCurrentContext = fromCurrentContext,
                Issue = "",
                EvidenceScore = evidenceScore
            };
        }

        private static List<SmartTargetEntryResolution> ResolveSmartTargetEntryCandidates(
            SmartLearningSnapshot snapshot, Dictionary<string, long> projectEntries, SmartMapEntry mappingEntry,
            SmartBoxTarget target, string signature, HashSet<string> preferredPrefixes)
        {
            Dictionary<string, SmartTargetEntryResolution> candidates =
                new Dictionary<string, SmartTargetEntryResolution>(StringComparer.OrdinalIgnoreCase);
            if (target == null || String.IsNullOrWhiteSpace(target.Code))
            {
                return new List<SmartTargetEntryResolution> { new SmartTargetEntryResolution
                {
                    Target = target, EntryCode = "", EntryName = "", Issue = "目标编号为空"
                } };
            }

            string[] signatureKeys = new[] { signature ?? "", SmartNameSegment(signature ?? "") + "|" };
            foreach (string signatureKey in signatureKeys)
            {
                List<SmartEntryStat> stats;
                if (!snapshot.EntryBySignatureQuota.TryGetValue(signatureKey + "\n" + (target.Code ?? ""), out stats)) continue;
                foreach (SmartEntryStat stat in stats ?? new List<SmartEntryStat>())
                {
                    if (stat == null) continue;
                    AddSmartTargetEntryCandidate(candidates, snapshot, projectEntries, target, stat.EntryCode,
                        stat.EntryName, stat.CurrentMethodEvidence,
                        (stat.CurrentMethodEvidence ? 40000 : 20000) + Math.Max(0, stat.ProjectCount));
                }
            }

            string targetIdentity = BuildLearningTargetIdentityKey(target.Kind, target.Code, target.Name, target.Unit);
            HashSet<string> localContexts = null;
            if (mappingEntry != null && mappingEntry.LocalContextKeysByTarget != null)
                mappingEntry.LocalContextKeysByTarget.TryGetValue(targetIdentity, out localContexts);
            if ((localContexts == null || localContexts.Count == 0) && mappingEntry != null &&
                mappingEntry.Targets != null && mappingEntry.Targets.Count == 1)
                localContexts = mappingEntry.LocalContextKeys;
            string methodPrefix = (snapshot.MethodNo ?? "").Trim() + "\n";
            foreach (string context in localContexts ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase))
            {
                if (!context.StartsWith(methodPrefix, StringComparison.OrdinalIgnoreCase)) continue;
                AddSmartTargetEntryCandidate(candidates, snapshot, projectEntries, target,
                    context.Substring(methodPrefix.Length), "", true, 60000);
            }

            List<SmartEntryStat> quotaStats;
            if (String.Equals(target.Kind ?? "quota", "quota", StringComparison.OrdinalIgnoreCase) &&
                snapshot.EntryByQuota.TryGetValue(target.Code ?? "", out quotaStats))
            {
                foreach (SmartEntryStat stat in quotaStats ?? new List<SmartEntryStat>())
                {
                    if (stat == null) continue;
                    AddSmartTargetEntryCandidate(candidates, snapshot, projectEntries, target, stat.EntryCode,
                        stat.EntryName, false, 1000 + Math.Max(0, stat.ProjectCount));
                }
            }

            List<SmartTargetEntryResolution> ordered = candidates.Values
                .OrderByDescending(item => item.EvidenceScore)
                .ThenBy(item => item.EntryCode, StringComparer.OrdinalIgnoreCase)
                .ToList();
            bool usePreferredPrefixes = GetLearningBaseTargetCode(target.Code) != "SH";
            if (usePreferredPrefixes && preferredPrefixes != null && preferredPrefixes.Count > 0)
            {
                List<SmartTargetEntryResolution> preferred = ordered.Where(item => item.EntryCode.Length >= 2 &&
                    preferredPrefixes.Contains(item.EntryCode.Substring(0, 2))).ToList();
                if (preferred.Count > 0) ordered = preferred;
            }
            if (ordered.Count > 1)
            {
                foreach (SmartTargetEntryResolution item in ordered) item.Issue = "目标条目证据不唯一，请选择";
            }
            if (ordered.Count > 0) return ordered;
            return new List<SmartTargetEntryResolution> { new SmartTargetEntryResolution
            {
                Target = target,
                EntryCode = "",
                EntryName = "",
                Issue = GetLearningBaseTargetCode(target.Code) == "SF"
                    ? "SF 必须写入设备购置费条目，当前项目未找到该条目"
                    : "学习库未定位到目标项目里的条目"
            } };
        }

        private static SmartTargetEntryResolution CloneSmartTargetEntryResolution(SmartTargetEntryResolution source)
        {
            return source == null ? null : new SmartTargetEntryResolution
            {
                Target = source.Target,
                EntryCode = source.EntryCode,
                EntryName = source.EntryName,
                EntrySeq = source.EntrySeq,
                FromCurrentContext = source.FromCurrentContext,
                Issue = source.Issue,
                EvidenceScore = source.EvidenceScore
            };
        }

        private static List<List<SmartTargetEntryResolution>> ResolveSmartTargetEntryCombinations(
            SmartLearningSnapshot snapshot, Dictionary<string, long> projectEntries, SmartMapEntry mappingEntry,
            string signature, HashSet<string> preferredPrefixes, out bool truncated)
        {
            const int MaxEntryCombinations = 16;
            truncated = false;
            List<List<SmartTargetEntryResolution>> combinations =
                new List<List<SmartTargetEntryResolution>> { new List<SmartTargetEntryResolution>() };
            foreach (SmartBoxTarget target in OrderSmartTargets(mappingEntry == null ? null : mappingEntry.Targets))
            {
                List<SmartTargetEntryResolution> options = ResolveSmartTargetEntryCandidates(snapshot, projectEntries,
                    mappingEntry, target, signature, preferredPrefixes);
                List<List<SmartTargetEntryResolution>> next = new List<List<SmartTargetEntryResolution>>();
                foreach (List<SmartTargetEntryResolution> combination in combinations)
                {
                    foreach (SmartTargetEntryResolution option in options)
                    {
                        if (next.Count >= MaxEntryCombinations) { truncated = true; break; }
                        List<SmartTargetEntryResolution> expanded = combination.Select(CloneSmartTargetEntryResolution).ToList();
                        expanded.Add(CloneSmartTargetEntryResolution(option));
                        next.Add(expanded);
                    }
                    if (next.Count >= MaxEntryCombinations && options.Count > 1) break;
                }
                combinations = next;
                if (combinations.Count == 0) break;
            }

            foreach (List<SmartTargetEntryResolution> combination in combinations)
            {
                SmartTargetEntryResolution primary = combination.FirstOrDefault(item => item != null && item.Target != null &&
                    IsPrimaryLearningTarget(item.Target.Kind, item.Target.Code) && !String.IsNullOrWhiteSpace(item.EntryCode));
                if (primary == null) continue;
                foreach (SmartTargetEntryResolution follower in combination.Where(item => item != null && item.Target != null &&
                    String.IsNullOrWhiteSpace(item.EntryCode) && IsSmartFollowerTarget(item.Target)))
                {
                    follower.EntryCode = primary.EntryCode;
                    follower.EntryName = primary.EntryName;
                    follower.EntrySeq = primary.EntrySeq;
                    follower.Issue = "";
                }
            }
            return combinations;
        }

        private static SmartTargetEntryResolution CreateSmartTargetEntryResolution(SmartLearningSnapshot snapshot,
            Dictionary<string, long> projectEntries, SmartBoxTarget target, SmartEntryStat stat,
            bool fromCurrentContext, string issue)
        {
            string entryCode = stat == null ? "" : stat.EntryCode ?? "";
            return new SmartTargetEntryResolution
            {
                Target = target,
                EntryCode = entryCode,
                EntryName = ResolveSmartEntryName(snapshot, entryCode, stat == null ? "" : stat.EntryName),
                EntrySeq = entryCode.Length > 0 && projectEntries.ContainsKey(entryCode) ? projectEntries[entryCode] : 0,
                FromCurrentContext = fromCurrentContext,
                Issue = issue ?? ""
            };
        }

        private static List<SmartEntryStat> DistinctAndFilterSmartEntryStats(IEnumerable<SmartEntryStat> stats,
            HashSet<string> preferredPrefixes, bool usePreferredPrefixes)
        {
            List<SmartEntryStat> result = (stats ?? Enumerable.Empty<SmartEntryStat>())
                .GroupBy(stat => stat.EntryCode ?? "", StringComparer.OrdinalIgnoreCase)
                .Select(group => group.OrderByDescending(stat => stat.ProjectCount).First())
                .Where(stat => !String.IsNullOrWhiteSpace(stat.EntryCode)).ToList();
            if (!usePreferredPrefixes || preferredPrefixes == null || preferredPrefixes.Count == 0) return result;
            List<SmartEntryStat> prefixed = result.Where(stat => stat.EntryCode.Length >= 2 &&
                preferredPrefixes.Contains(stat.EntryCode.Substring(0, 2))).ToList();
            return prefixed.Count > 0 ? prefixed : result;
        }

        private static List<string> SelectSmartLocalEntries(SmartLearningSnapshot snapshot,
            Dictionary<string, long> projectEntries, SmartBoxTarget target, IEnumerable<string> contexts,
            HashSet<string> preferredPrefixes, bool usePreferredPrefixes)
        {
            string methodPrefix = (snapshot.MethodNo ?? "").Trim() + "\n";
            List<string> result = (contexts ?? Enumerable.Empty<string>())
                .Where(key => key.StartsWith(methodPrefix, StringComparison.OrdinalIgnoreCase))
                .Select(key => key.Substring(methodPrefix.Length))
                .Where(code => projectEntries.ContainsKey(code) &&
                    SmartEntryCodeMatchesScope(code, snapshot.SelectedScope) &&
                    IsSmartTargetEntryCompatible(target, ResolveSmartEntryName(snapshot, code, "")))
                .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (!usePreferredPrefixes || preferredPrefixes == null || preferredPrefixes.Count == 0) return result;
            List<string> prefixed = result.Where(code => code.Length >= 2 && preferredPrefixes.Contains(code.Substring(0, 2))).ToList();
            return prefixed.Count > 0 ? prefixed : result;
        }

        private static bool IsSmartFollowerTarget(SmartBoxTarget target)
        {
            if (target == null) return false;
            if (String.Equals(target.Kind ?? "quota", "material", StringComparison.OrdinalIgnoreCase)) return true;
            string baseCode = GetLearningBaseTargetCode(target.Code);
            return baseCode == "ZLF" || baseCode == "LF";
        }

        private static bool IsSmartTargetEntryCompatible(SmartBoxTarget target, string entryName)
        {
            bool equipmentEntry = (entryName ?? "").IndexOf("设备购置费", StringComparison.OrdinalIgnoreCase) >= 0;
            bool sf = target != null && GetLearningBaseTargetCode(target.Code) == "SF";
            return sf ? equipmentEntry : !equipmentEntry;
        }

        private static string ResolveSmartEntryName(SmartLearningSnapshot snapshot, string entryCode, string learnedName)
        {
            string projectName;
            if (snapshot != null && snapshot.ProjectEntryNameByCode != null &&
                snapshot.ProjectEntryNameByCode.TryGetValue(entryCode ?? "", out projectName) &&
                !String.IsNullOrWhiteSpace(projectName))
            {
                return projectName.Trim();
            }
            string learningName;
            if (snapshot != null && snapshot.LearningEntryNameByCode != null &&
                snapshot.LearningEntryNameByCode.TryGetValue(entryCode ?? "", out learningName) &&
                !String.IsNullOrWhiteSpace(learningName))
            {
                return learningName.Trim();
            }
            return (learnedName ?? "").Trim();
        }

        // 目标项目 定额输入 的列集合(跨库复制时过滤源行里目标库没有的列)。
        private static HashSet<string> LoadQuotaInputColumns(SqlConnection conn, SqlTransaction transaction)
        {
            HashSet<string> columns = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            using (SqlCommand cmd = conn.CreateCommand())
            {
                if (transaction != null) cmd.Transaction = transaction;
                cmd.CommandText = "select top 0 * from 定额输入";
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    for (int i = 0; i < reader.FieldCount; i++) columns.Add(reader.GetName(i));
                }
            }
            return columns;
        }

        // 跨库整行复制:只允许使用 DPAPI 凭据库中明确登记的业务/学习端点做只读溯源。
        private static Dictionary<string, object> LoadCrossDbQuotaRow(SqlConnection targetConn, FillPreviewItem item, HashSet<string> targetColumns)
        {
            foreach (string credentialName in new[] { "business", "learning" })
            {
                try
                {
                    string connectionString = RecoSqlCredentialStore.BuildConnectionString(credentialName, item.SourceDb, 1433, 3);
                    using (SqlConnection src = new SqlConnection(connectionString))
                    {
                        src.Open();
                        using (SqlCommand cmd = src.CreateCommand())
                        {
                            cmd.CommandText = "select * from 定额输入 where 定额序号=@id";
                            cmd.Parameters.AddWithValue("@id", item.SourceDbQuotaSeq);
                            using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                            {
                                DataTable table = new DataTable();
                                adapter.Fill(table);
                                if (table.Rows.Count == 0) return null;   // 库连上了但源行已删,不再试其他服务器
                                Dictionary<string, object> values = new Dictionary<string, object>();
                                foreach (DataColumn column in table.Columns)
                                {
                                    if (targetColumns.Contains(column.ColumnName))
                                    {
                                        values[column.ColumnName] = table.Rows[0][column];
                                    }
                                }
                                return values;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log("Cross-db quota row load failed. endpoint=" + credentialName +
                        " database=" + item.SourceDb + " error=" + ex.GetType().Name);
                }
            }
            return null;
        }

        private static List<SmartMapCandidateScore> RankSmartMapEntries(SmartLearningSnapshot snapshot, List<SmartMapEntry> hits,
            string signature, HashSet<string> preferredPrefixes, Dictionary<string, long> projectEntries,
            Dictionary<string, ProjectQuota> currentQuotaByCode)
        {
            List<SmartMapCandidateScore> scores = new List<SmartMapCandidateScore>();
            foreach (SmartMapEntry hit in hits ?? new List<SmartMapEntry>())
            {
                bool targetsValid = IsSmartMapEntryUsableInCurrentProject(hit, "", currentQuotaByCode);
                if (!targetsValid) continue;
                bool truncated;
                List<List<SmartTargetEntryResolution>> combinations = ResolveSmartTargetEntryCombinations(snapshot,
                    projectEntries, hit, signature, preferredPrefixes, out truncated);
                foreach (List<SmartTargetEntryResolution> targetEntries in combinations)
                {
                    SmartTargetEntryResolution representative = targetEntries.FirstOrDefault(item => item != null && item.Target != null &&
                        IsPrimaryLearningTarget(item.Target.Kind, item.Target.Code) && !String.IsNullOrWhiteSpace(item.EntryCode)) ??
                        targetEntries.FirstOrDefault(item => item != null && !String.IsNullOrWhiteSpace(item.EntryCode));
                    string entryCode = representative == null ? "" : representative.EntryCode ?? "";
                    string entryName = representative == null ? "" : representative.EntryName ?? "";
                    long entrySeq = representative == null ? 0 : representative.EntrySeq;
                    bool hasEntry = targetEntries.Count > 0 && targetEntries.All(item => item != null && !String.IsNullOrWhiteSpace(item.EntryCode));
                    bool fromContext = hasEntry && targetEntries.All(item => item.FromCurrentContext);
                    bool prefixMatch = hasEntry && entryCode.Length >= 2 && preferredPrefixes != null && preferredPrefixes.Count > 0 &&
                        preferredPrefixes.Contains(entryCode.Substring(0, 2));
                    string currentContextPrefix = (snapshot.MethodNo ?? "").Trim() + "\n";
                    bool currentMethodMapping = hit.CurrentMethodMapping || (hit.LocalContextKeys != null &&
                        !String.IsNullOrWhiteSpace(snapshot.MethodNo) && hit.LocalContextKeys.Any(key =>
                            key.StartsWith(currentContextPrefix, StringComparison.OrdinalIgnoreCase)));
                    scores.Add(new SmartMapCandidateScore
                    {
                        Entry = hit,
                        EntryCode = entryCode,
                        EntryName = entryName,
                        EntrySeq = entrySeq,
                        HasEntry = hasEntry,
                        HasCurrentContext = fromContext,
                        HasCurrentMethodMapping = currentMethodMapping,
                        PrefixMatch = prefixMatch,
                        CurrentTargetsValid = targetsValid,
                        EntryCandidatesTruncated = truncated,
                        TargetEntries = targetEntries
                    });
                }
            }
            return OrderSmartMapCandidateScores(scores);
        }

        private static List<SmartMapCandidateScore> OrderSmartMapCandidateScores(IEnumerable<SmartMapCandidateScore> scores)
        {
            return (scores ?? Enumerable.Empty<SmartMapCandidateScore>()).OrderByDescending(score => score.CurrentTargetsValid)
                .ThenByDescending(score => score.HasCurrentMethodMapping)
                .ThenByDescending(score => score.HasCurrentContext)
                .ThenByDescending(score => score.PrefixMatch)
                .ThenByDescending(score => score.HasEntry)
                .ThenByDescending(score => score.Entry == null ? 0 : score.Entry.Weight)
                .ThenByDescending(score => score.Entry == null || score.Entry.Targets == null ? 0 : score.Entry.Targets.Count)
                .ThenByDescending(score => score.Entry == null ? DateTime.MinValue : score.Entry.LastUsedAt)
                .ThenBy(score => score.Entry == null ? "" : score.Entry.BoxId, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static bool CanAutoSelectSmartMapEntry(List<SmartMapCandidateScore> scores)
        {
            if (scores == null || scores.Count == 0) return false;
            SmartMapCandidateScore top = scores[0];
            if (!top.CurrentTargetsValid || !top.HasEntry || !top.HasCurrentContext || top.EntryCandidatesTruncated) return false;
            if (!top.HasCurrentMethodMapping && !IsSingleQuotaTargetBox(top.Entry)) return false;
            if (scores.Count == 1) return true;

            SmartMapCandidateScore second = scores[1];
            if (!second.CurrentTargetsValid) return true;
            if (top.HasCurrentMethodMapping && !second.HasCurrentMethodMapping &&
                HasMinimumPositiveEvidence(top)) return true;
            if (top.HasCurrentContext && !second.HasCurrentContext &&
                HasMinimumPositiveEvidence(top)) return true;
            int topWeight = top.Entry == null ? 0 : top.Entry.Weight;
            int secondWeight = second.Entry == null ? 0 : second.Entry.Weight;
            if (topWeight - secondWeight >= 20) return true;
            // 小样本时也只比较净证据权重，不把不同分值的反馈简单按次数等同。
            return topWeight >= 2 * secondWeight && topWeight - secondWeight >= 10;
        }

        private static bool HasMinimumPositiveEvidence(SmartMapCandidateScore score)
        {
            return score != null && score.Entry != null &&
                (score.Entry.AcceptedCount >= 2 || score.Entry.CorrectedCount >= 1);
        }

        private static bool IsSingleQuotaTargetBox(SmartMapEntry entry)
        {
            return entry != null && entry.Targets != null && entry.Targets.Count == 1 &&
                IsPrimaryLearningTarget(entry.Targets[0].Kind, entry.Targets[0].Code);
        }

        private static IEnumerable<SmartBoxTarget> OrderSmartTargets(IEnumerable<SmartBoxTarget> targets)
        {
            return (targets ?? Enumerable.Empty<SmartBoxTarget>())
                .Where(target => target != null)
                .OrderBy(target => TemplateTargetRank(target.Code))
                .ThenBy(target => target.Code ?? "", StringComparer.OrdinalIgnoreCase);
        }

        private static string ResolveSmartProfessionName(SmartLearningSnapshot snapshot, string entryCode)
        {
            if (String.IsNullOrWhiteSpace(entryCode) || entryCode.Trim().Length < 2) return "";
            string prefix = entryCode.Trim().Substring(0, 2);
            string name;
            if (snapshot != null && snapshot.ProjectEntryNameByCode.TryGetValue(prefix, out name) && !String.IsNullOrWhiteSpace(name))
                return name.Trim();
            if (snapshot != null && snapshot.LearningEntryNameByCode.TryGetValue(prefix, out name) && !String.IsNullOrWhiteSpace(name))
                return name.Trim();
            return prefix;
        }

        private static bool SmartEntryCodeMatchesScope(string entryCode, SmartLearningScope scope)
        {
            if (scope == null || String.Equals(scope.Kind, "All", StringComparison.OrdinalIgnoreCase)) return true;
            if (!String.Equals(scope.Kind, "Entry", StringComparison.OrdinalIgnoreCase)) return false;
            string code = (entryCode ?? "").Trim();
            string selected = (scope.EntryCode ?? "").Trim();
            return IsSmartClassifiedEntryCode(code) && IsSmartClassifiedEntryCode(selected) &&
                IsItemNoUnderChapter(code, selected);
        }

        private static List<SmartMapEntry> FilterSmartHitsByScope(SmartLearningSnapshot snapshot,
            IEnumerable<SmartMapEntry> hits, SmartLearningScope scope)
        {
            List<SmartMapEntry> source = (hits ?? Enumerable.Empty<SmartMapEntry>()).Where(entry => entry != null).ToList();
            if (scope == null || String.Equals(scope.Kind, "All", StringComparison.OrdinalIgnoreCase)) return source;
            return source.Where(entry =>
            {
                HashSet<string> scopeCodes;
                HashSet<string> merged = snapshot != null &&
                    snapshot.ScopeEntriesByBox.TryGetValue(entry.BoxId ?? "", out scopeCodes)
                    ? new HashSet<string>(scopeCodes, StringComparer.OrdinalIgnoreCase)
                    : new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                bool hasPersistedScope = merged.Count > 0;
                if (entry.LocalContextKeys != null)
                {
                    foreach (string key in entry.LocalContextKeys)
                    {
                        int separator = (key ?? "").IndexOf('\n');
                        string localMethodNo = separator >= 0 ? key.Substring(0, separator) : "";
                        if (!String.Equals(localMethodNo, snapshot == null ? "" : snapshot.MethodNo ?? "",
                            StringComparison.OrdinalIgnoreCase)) continue;
                        string code = separator >= 0 ? key.Substring(separator + 1) : "";
                        if (!IsSmartClassifiedEntryCode(code)) continue;
                        merged.Add(code);
                    }
                }
                if (String.Equals(scope.Kind, "Unclassified", StringComparison.OrdinalIgnoreCase)) return !hasPersistedScope;
                return merged.Any(code => SmartEntryCodeMatchesScope(code, scope));
            }).ToList();
        }

        private static string BuildSmartCandidateLabel(SmartLearningSnapshot snapshot, SmartMapCandidateScore score)
        {
            if (score == null || score.Entry == null) return "空组件";
            Dictionary<SmartBoxTarget, SmartTargetEntryResolution> resolutions = (score.TargetEntries ?? new List<SmartTargetEntryResolution>())
                .Where(item => item != null && item.Target != null).ToDictionary(item => item.Target, item => item);
            List<string> parts = new List<string>();
            foreach (SmartBoxTarget target in OrderSmartTargets(score.Entry.Targets))
            {
                string code = (target.Code ?? "").Trim();
                if (code.Length == 0) continue;
                SmartTargetEntryResolution resolution;
                if (!resolutions.TryGetValue(target, out resolution) || !IsSmartClassifiedEntryCode(resolution.EntryCode))
                {
                    parts.Add(code + "（缺条目）");
                    continue;
                }
                string entryName = String.IsNullOrWhiteSpace(resolution.EntryName)
                    ? ResolveSmartProfessionName(snapshot, resolution.EntryCode)
                    : resolution.EntryName.Trim();
                parts.Add(code + "（" + entryName + " " + resolution.EntryCode.Trim() + "）");
            }
            return parts.Count == 0 ? "空组件" : String.Join(" + ", parts.ToArray());
        }

        private static string BuildSmartEntryCombinationKey(SmartMapCandidateScore score)
        {
            string raw = String.Join("|", (score == null ? new List<SmartTargetEntryResolution>() : score.TargetEntries)
                .Where(item => item != null && item.Target != null)
                .Select(item => BuildLearningTargetIdentityKey(item.Target.Kind, item.Target.Code, item.Target.Name, item.Target.Unit) +
                    "=" + (item.EntryCode ?? ""))
                .ToArray());
            return BuildLearningMd5(raw).Substring(0, 12);
        }

        private static List<NameQuotaCandidateGroup> DeduplicateSmartCandidatesByLabel(
            IEnumerable<NameQuotaCandidateGroup> candidates)
        {
            return (candidates ?? Enumerable.Empty<NameQuotaCandidateGroup>())
                .Where(candidate => candidate != null)
                .GroupBy(candidate => candidate.Label ?? "", StringComparer.Ordinal)
                .Select(group => group.First())
                .ToList();
        }

        private static bool AppendRankedSmartMatch(List<FillPreviewItem> items, TargetQtyRow row, List<TargetQtyRow> targetRows,
            List<SmartMapEntry> hits, SmartLearningSnapshot snapshot, Dictionary<string, long> projectEntries,
            Dictionary<string, ProjectQuota> currentQuotaByCode, string baseNote, string signature,
            HashSet<string> preferredPrefixes, Dictionary<string, int> prefixVotes)
        {
            List<SmartMapCandidateScore> scores = RankSmartMapEntries(snapshot, hits, signature, preferredPrefixes,
                projectEntries, currentQuotaByCode);
            if (scores.Count == 0) return false;
            if (CanAutoSelectSmartMapEntry(scores))
            {
                AppendSmartItems(items, row, targetRows, scores[0].Entry, snapshot, projectEntries, currentQuotaByCode,
                    scores[0].TargetEntries, false, baseNote + "，" + BuildSmartCandidateLabel(snapshot, scores[0]),
                    signature, preferredPrefixes, prefixVotes);
                return true;
            }

            List<NameQuotaCandidateGroup> candidates = new List<NameQuotaCandidateGroup>();
            foreach (SmartMapCandidateScore score in scores)
            {
                List<FillPreviewItem> candidateItems = new List<FillPreviewItem>();
                string label = BuildSmartCandidateLabel(snapshot, score);
                AppendSmartItems(candidateItems, row, targetRows, score.Entry, snapshot, projectEntries, currentQuotaByCode,
                    score.TargetEntries, true, baseNote + "，候选：" + label, signature, preferredPrefixes, null);
                if (candidateItems.Count == 0) continue;
                if (score.EntryCandidatesTruncated)
                {
                    foreach (FillPreviewItem candidateItem in candidateItems)
                        candidateItem.AlignNote = AppendPreviewNote(candidateItem.AlignNote, "候选已截断为前16组");
                }
                candidates.Add(new NameQuotaCandidateGroup
                {
                    Key = "smart:" + (score.Entry.BoxId ?? "") + "#" + BuildSmartEntryCombinationKey(score),
                    Label = label,
                    Items = candidateItems
                });
            }
            candidates = DeduplicateSmartCandidatesByLabel(candidates);
            if (candidates.Count == 0) return false;

            List<FillPreviewItem> active = candidates[0].Items.Select(item => item.CloneForNameCandidate()).ToList();
            foreach (FillPreviewItem item in active)
            {
                item.Selected = false;
                item.NeedExactNameConfirmation = true;
            }
            active[0].NameQuotaCandidates = candidates;
            active[0].SelectedNameQuotaCandidateKey = candidates[0].Key;
            active[0].AlignNote = AppendPreviewNote(active[0].AlignNote,
                candidates.Any(candidate => candidate != null && candidate.Items != null &&
                    candidate.Items.Any(item => (item.AlignNote ?? "").IndexOf("候选已截断", StringComparison.OrdinalIgnoreCase) >= 0))
                    ? "条目候选已截断为前16组，请确认完整组件"
                    : (candidates.Count > 1 ? "组件候选接近或上下文不唯一，请确认完整组件" : "缺少唯一的当前办法/条目证据，请确认组件"));
            items.AddRange(active);
            return true;
        }

        private static List<KeyValuePair<int, string>> BuildSmartFuzzyScoresIfUnmatched(bool matched, string nameSignature,
            List<KeyValuePair<string, MatchTextFeatures>> nameFeatures)
        {
            List<KeyValuePair<int, string>> scored = new List<KeyValuePair<int, string>>();
            if (matched) return scored;

            MatchTextFeatures rowFeatures = BuildMatchTextFeatures(nameSignature);
            foreach (KeyValuePair<string, MatchTextFeatures> pair in nameFeatures)
            {
                if (!HaveCompatibleSmartSpecificationNumbers(nameSignature, pair.Key)) continue;
                int score = MatchNameScore(rowFeatures, pair.Value);
                if (score > 0) scored.Add(new KeyValuePair<int, string>(score, pair.Key));
            }
            scored.Sort(delegate(KeyValuePair<int, string> a, KeyValuePair<int, string> b) { return b.Key.CompareTo(a.Key); });
            return scored;
        }

        private static bool HaveCompatibleSmartSpecificationNumbers(string leftSignature, string rightSignature)
        {
            List<string> left = ExtractMatchNumbers(leftSignature ?? "");
            List<string> right = ExtractMatchNumbers(rightSignature ?? "");
            return left.Count == 0 || right.Count == 0 || left.SequenceEqual(right, StringComparer.OrdinalIgnoreCase);
        }

        private static List<NameQuotaCandidateGroup> BuildSmartFuzzyCandidateGroups(
            IEnumerable<KeyValuePair<int, string>> scored, SmartLearningScope scope, string sourceNote,
            SmartLearningSnapshot snapshot, string signature, HashSet<string> preferredPrefixes,
            Dictionary<string, long> projectEntries, Dictionary<string, ProjectQuota> currentQuotaByCode,
            TargetQtyRow row, List<TargetQtyRow> targetRows)
        {
            List<NameQuotaCandidateGroup> result = new List<NameQuotaCandidateGroup>();
            int matchedNames = 0;
            foreach (KeyValuePair<int, string> cand in (scored ?? Enumerable.Empty<KeyValuePair<int, string>>()))
            {
                List<SmartMapEntry> allHits;
                if (!snapshot.ByNameOnly.TryGetValue(cand.Value, out allHits) || allHits.Count == 0) continue;
                List<SmartMapEntry> candHits = FilterSmartHitsByScope(snapshot, allHits, scope);
                if (candHits.Count == 0) continue;
                matchedNames++;
                List<SmartMapCandidateScore> candidateScores = RankSmartMapEntries(snapshot, candHits, signature,
                    preferredPrefixes, projectEntries, currentQuotaByCode);
                foreach (SmartMapCandidateScore candidateScore in candidateScores.Take(16))
                {
                    NameQuotaCandidateGroup group = new NameQuotaCandidateGroup
                    {
                        Key = cand.Value + "|" + candidateScore.Entry.BoxId + "#" + BuildSmartEntryCombinationKey(candidateScore),
                        Label = "≈" + cand.Value + "(" + cand.Key.ToString(CultureInfo.InvariantCulture) + "分)：" +
                            BuildSmartCandidateLabel(snapshot, candidateScore)
                    };
                    List<FillPreviewItem> groupItems = new List<FillPreviewItem>();
                    AppendSmartItems(groupItems, row, targetRows, candidateScore.Entry, snapshot, projectEntries,
                        currentQuotaByCode, candidateScore.TargetEntries, true, "模糊候选，" + sourceNote + ":" + group.Label,
                        signature, preferredPrefixes, null);
                    if (candidateScore.EntryCandidatesTruncated)
                    {
                        foreach (FillPreviewItem groupItem in groupItems)
                            groupItem.AlignNote = AppendPreviewNote(groupItem.AlignNote, "候选已截断为前16组");
                    }
                    group.Items = groupItems;
                    if (groupItems.Count > 0) result.Add(group);
                }
                if (matchedNames >= 3) break;
            }
            return DeduplicateSmartCandidatesByLabel(result);
        }

        private static bool TryEvaluateSmartFormula(SmartFormulaRule rule, List<TargetQtyRow> targetRows,
            TargetQtyRow anchorRow, string currentTargetUnit, out string quantityText)
        {
            quantityText = "";
            if (rule == null || rule.Operands.Count == 0 || String.IsNullOrWhiteSpace(rule.Template) ||
                String.IsNullOrWhiteSpace(currentTargetUnit)) return false;
            string outputSuffix;
            if (!TryBuildExcelLinkUnitScaleSuffix(rule.TargetUnit, currentTargetUnit, out outputSuffix)) return false;

            string expression = rule.Template;
            foreach (SmartFormulaOperand operand in rule.Operands.OrderByDescending(item => item.Index))
            {
                TargetQtyRow row = null;
                if (operand.Index == 0 && String.Equals(NormalizeForSignature(anchorRow.RawName) + "|", operand.Signature, StringComparison.OrdinalIgnoreCase))
                {
                    row = anchorRow;
                }
                else
                {
                    List<TargetQtyRow> candidates = targetRows.Where(item =>
                        String.Equals(NormalizeForSignature(item.RawName) + "|", operand.Signature, StringComparison.OrdinalIgnoreCase) &&
                        String.Equals(item.Chapter ?? "", anchorRow.Chapter ?? "", StringComparison.OrdinalIgnoreCase) &&
                        Math.Abs(item.Row - anchorRow.Row) <= 20).ToList();
                    if (candidates.Count != 1) return false;
                    row = candidates[0];
                }
                string operandSuffix;
                if (!TryBuildExcelLinkUnitScaleSuffix(row.Unit, operand.Unit, out operandSuffix)) return false;
                string operandText = (row.QuantityText ?? "") + operandSuffix;
                decimal operandValue;
                string operandError;
                if (!TryEvaluateDecimal(operandText, out operandValue, out operandError)) return false;
                string pattern = "V" + operand.Index.ToString(CultureInfo.InvariantCulture) + "(?![0-9])";
                string currentExpression = expression;
                expression = System.Text.RegularExpressions.Regex.Replace(expression, pattern,
                    delegate(System.Text.RegularExpressions.Match match)
                    {
                        return FormatSmartFormulaOperandForInsertion(currentExpression, match.Index, match.Length, operandText);
                    });
            }
            if (System.Text.RegularExpressions.Regex.IsMatch(expression, "V[0-9]+")) return false;
            string finalExpression = expression;
            if (!String.IsNullOrEmpty(outputSuffix))
                finalExpression = FormatSmartFormulaOperandForInsertion("V0" + outputSuffix, 0, 2, expression) + outputSuffix;
            decimal result;
            string error;
            if (!TryEvaluateDecimal(finalExpression, out result, out error)) return false;
            if (result <= 0m) return false;
            quantityText = finalExpression;
            return true;
        }

        private static string FormatSmartFormulaOperandForInsertion(string template, int tokenStart, int tokenLength,
            string operandText)
        {
            string value = TrimWholeFormulaParentheses(operandText);
            if (value.Length == 0) return value;

            bool hasAdditive = HasTopLevelSmartFormulaOperator(value, true);
            bool hasMultiplicative = HasTopLevelSmartFormulaOperator(value, false);
            char left = FindSmartFormulaNeighbor(template, tokenStart - 1, -1);
            char right = FindSmartFormulaNeighbor(template, tokenStart + tokenLength, 1);
            bool needsParentheses = hasAdditive && (left == '-' || left == '*' || left == '/' || right == '*' || right == '/') ||
                hasMultiplicative && left == '/';
            return needsParentheses ? "(" + value + ")" : value;
        }

        private static string TrimWholeFormulaParentheses(string value)
        {
            string result = (value ?? "").Trim();
            while (result.Length >= 2 && result[0] == '(' && result[result.Length - 1] == ')')
            {
                int depth = 0;
                bool wrapsWholeValue = true;
                for (int i = 0; i < result.Length; i++)
                {
                    if (result[i] == '(') depth++;
                    else if (result[i] == ')') depth--;
                    if (depth == 0 && i < result.Length - 1)
                    {
                        wrapsWholeValue = false;
                        break;
                    }
                    if (depth < 0) return result;
                }
                if (!wrapsWholeValue || depth != 0) break;
                result = result.Substring(1, result.Length - 2).Trim();
            }
            return result;
        }

        private static bool HasTopLevelSmartFormulaOperator(string value, bool additive)
        {
            int depth = 0;
            char previous = '\0';
            for (int i = 0; i < (value ?? "").Length; i++)
            {
                char ch = value[i];
                if (Char.IsWhiteSpace(ch)) continue;
                if (ch == '(') { depth++; previous = ch; continue; }
                if (ch == ')') { if (depth > 0) depth--; previous = ch; continue; }
                if (depth == 0)
                {
                    if (additive && (ch == '+' || ch == '-') &&
                        previous != '\0' && previous != '(' && previous != '+' && previous != '-' && previous != '*' && previous != '/')
                        return true;
                    if (!additive && (ch == '*' || ch == '/')) return true;
                }
                previous = ch;
            }
            return false;
        }

        private static char FindSmartFormulaNeighbor(string template, int start, int direction)
        {
            for (int i = start; i >= 0 && i < (template ?? "").Length; i += direction)
            {
                if (!Char.IsWhiteSpace(template[i])) return template[i];
            }
            return '\0';
        }

        private static bool IsDerivedSmartFormula(SmartFormulaRule rule)
        {
            if (rule == null || String.IsNullOrWhiteSpace(rule.Template)) return false;
            System.Text.RegularExpressions.MatchCollection variables =
                System.Text.RegularExpressions.Regex.Matches(rule.Template, "V[0-9]+");
            if (variables.Count == 0) return false;
            if (rule.Operands.Count > 1 || variables.Count > 1) return true;

            // 单参数纯比例 V0*k 在同量纲下仍以当前标准单位换算为准。
            string compact = System.Text.RegularExpressions.Regex.Replace(rule.Template, "\\s+", "");
            const string number = "(?:[0-9]+(?:\\.[0-9]*)?|\\.[0-9]+)";
            return !System.Text.RegularExpressions.Regex.IsMatch(compact,
                "^V[0-9]+(?:(?:\\*|/)" + number + ")*$");
        }

        private static List<SmartFormulaRule> SelectContextualSmartFormulaRules(SmartLearningSnapshot snapshot,
            List<SmartFormulaRule> rules, string entryCode)
        {
            if (snapshot == null || rules == null || String.IsNullOrWhiteSpace(entryCode))
                return new List<SmartFormulaRule>();
            return rules.Where(rule =>
                String.Equals(rule.Method ?? "", snapshot.Method ?? "", StringComparison.OrdinalIgnoreCase) &&
                !String.IsNullOrWhiteSpace(rule.EntryCode) &&
                String.Equals(rule.EntryCode ?? "", entryCode ?? "", StringComparison.OrdinalIgnoreCase)).ToList();
        }

        private static bool HasContextualDerivedSmartFormula(SmartLearningSnapshot snapshot, string formulaKey, string entryCode)
        {
            List<SmartFormulaRule> rules;
            return snapshot.FormulaByKey.TryGetValue(formulaKey, out rules) &&
                SelectContextualSmartFormulaRules(snapshot, rules, entryCode).Any(IsDerivedSmartFormula);
        }

        private static bool TryResolveSmartFormula(SmartLearningSnapshot snapshot, List<TargetQtyRow> targetRows,
            TargetQtyRow anchorRow, SmartBoxTarget target, string currentTargetUnit, string entryCode,
            string signature, out SmartFormulaRule selectedRule, out string quantityText, out string issue)
        {
            return TryResolveSmartFormulaCore(snapshot, targetRows, anchorRow, target, currentTargetUnit, entryCode,
                signature, false, out selectedRule, out quantityText, out issue);
        }

        private static bool TryResolveDerivedSmartFormula(SmartLearningSnapshot snapshot, List<TargetQtyRow> targetRows,
            TargetQtyRow anchorRow, SmartBoxTarget target, string currentTargetUnit, string entryCode,
            string signature, out SmartFormulaRule selectedRule, out string quantityText, out string issue)
        {
            return TryResolveSmartFormulaCore(snapshot, targetRows, anchorRow, target, currentTargetUnit, entryCode,
                signature, true, out selectedRule, out quantityText, out issue);
        }

        private static bool TryResolveSmartFormulaCore(SmartLearningSnapshot snapshot, List<TargetQtyRow> targetRows,
            TargetQtyRow anchorRow, SmartBoxTarget target, string currentTargetUnit, string entryCode,
            string signature, bool derivedOnly, out SmartFormulaRule selectedRule, out string quantityText, out string issue)
        {
            selectedRule = null;
            quantityText = "";
            issue = "";
            List<SmartFormulaRule> rules;
            if (!snapshot.FormulaByKey.TryGetValue(BuildSmartFormulaKey(signature, target.Kind, target.Code), out rules) || rules.Count == 0)
            {
                issue = "单位 " + anchorRow.Unit + "→" + currentTargetUnit + " 无可靠换算公式";
                return false;
            }
            List<SmartFormulaRule> contextual = SelectContextualSmartFormulaRules(snapshot, rules, entryCode);
            if (contextual.Count == 0)
            {
                issue = "当前办法/条目没有可复用的换算公式";
                return false;
            }
            if (derivedOnly) contextual = contextual.Where(IsDerivedSmartFormula).ToList();
            if (contextual.Count == 0)
            {
                issue = "当前办法/条目没有可复用的派生换算公式";
                return false;
            }

            List<SmartFormulaEvaluation> valid = new List<SmartFormulaEvaluation>();
            foreach (SmartFormulaRule rule in contextual)
            {
                string evaluated;
                if (TryEvaluateSmartFormula(rule, targetRows, anchorRow, currentTargetUnit, out evaluated))
                {
                    valid.Add(new SmartFormulaEvaluation { Rule = rule, QuantityText = evaluated });
                }
            }
            valid = valid.OrderByDescending(item => String.Equals(item.Rule.Method ?? "", snapshot.Method ?? "", StringComparison.OrdinalIgnoreCase))
                .ThenByDescending(item => item.Rule.SampleCount).ThenByDescending(item => item.Rule.LastSeen).ToList();
            if (valid.Count == 0)
            {
                issue = "换算公式参数未找齐、单位不符或存在重名歧义";
                return false;
            }
            if (valid.Count > 1 &&
                String.Equals(valid[0].Rule.Method ?? "", valid[1].Rule.Method ?? "", StringComparison.OrdinalIgnoreCase) &&
                valid[0].Rule.SampleCount == valid[1].Rule.SampleCount &&
                !String.Equals(valid[0].Rule.RuleHash, valid[1].Rule.RuleHash, StringComparison.OrdinalIgnoreCase))
            {
                issue = "存在多套同权重换算公式，需确认";
                return false;
            }
            selectedRule = valid[0].Rule;
            quantityText = valid[0].QuantityText;
            return true;
        }

        // 由一个映射命中构建预览项(每个定额目标一行,首行承载工程量名)。
        private static void AppendSmartItems(List<FillPreviewItem> items, TargetQtyRow row, List<TargetQtyRow> targetRows, SmartMapEntry entry,
            SmartLearningSnapshot snapshot, Dictionary<string, long> projectEntries,
            Dictionary<string, ProjectQuota> currentQuotaByCode, List<SmartTargetEntryResolution> targetEntries,
            bool needConfirm, string note,
            string signature, HashSet<string> preferredPrefixes, Dictionary<string, int> prefixVotes)
        {
            List<SmartTargetEntryResolution> resolutions = targetEntries ??
                ResolveSmartTargetEntries(snapshot, projectEntries, entry, signature, preferredPrefixes);
            SmartTargetEntryResolution primary = resolutions.FirstOrDefault(item => item != null && item.Target != null &&
                IsPrimaryLearningTarget(item.Target.Kind, item.Target.Code));
            if (primary != null && primary.FromCurrentContext && primary.EntryCode != null &&
                primary.EntryCode.Length >= 2 && prefixVotes != null)
            {
                string prefix = primary.EntryCode.Substring(0, 2);
                int votes;
                prefixVotes.TryGetValue(prefix, out votes);
                prefixVotes[prefix] = votes + 1;
            }
            int order = 0;
            List<FillPreviewItem> groupItems = new List<FillPreviewItem>();
            foreach (SmartTargetEntryResolution resolution in resolutions)
            {
                SmartBoxTarget target = resolution == null ? null : resolution.Target;
                if (target == null || String.IsNullOrEmpty(target.Code)) continue;
                string entryCode = resolution.EntryCode ?? "";
                string entryName = resolution.EntryName ?? "";
                long entrySeq = resolution.EntrySeq;
                bool hasEntry = !String.IsNullOrWhiteSpace(entryCode);
                ProjectQuota currentQuota;
                TryGetCurrentSmartQuota(currentQuotaByCode, target, out currentQuota);
                string currentQuotaUnit = currentQuota == null ? "" : currentQuota.Unit;
                FillPreviewItem item = new FillPreviewItem
                {
                    TemplateName = "推荐定额",
                    IsNameDriven = true,
                    TargetRow = row.Row,
                    TargetName = row.DisplayName,
                    TargetFullName = row.RawName,
                    TargetChapter = row.Chapter,
                    TargetUnit = row.Unit,
                    TargetQuantityText = row.QuantityText,
                    QuotaCode = target.Code,
                    SourceName = currentQuota == null || String.IsNullOrWhiteSpace(currentQuota.Name) ? target.Name : currentQuota.Name,
                    Unit = currentQuotaUnit,
                    ChosenItemName = entryName,
                    GroupOrder = order,
                    OrderInItem = row.Row * 10 + order,
                    NeedExactNameConfirmation = needConfirm,
                    AlignNote = note
                };
                if (String.IsNullOrWhiteSpace(currentQuotaUnit))
                {
                    item.QuantityText = row.QuantityText;
                    item.Selected = false;
                    item.NeedExactNameConfirmation = true;
                    item.Status = "缺当前定额单位";
                    item.AlignNote = AppendPreviewNote(item.AlignNote, "当前版本未找到 " + target.Code + " 的定额单位，禁止自动写入");
                }
                else
                {
                    string formulaKey = BuildSmartFormulaKey(signature, target.Kind, target.Code);
                    List<SmartFormulaRule> formulaRules;
                    bool hasFormula = snapshot.FormulaByKey.TryGetValue(formulaKey, out formulaRules) &&
                        SelectContextualSmartFormulaRules(snapshot, formulaRules, entryCode).Count > 0;
                    SmartFormulaRule formulaRule = null;
                    string formulaQuantity = "";
                    string formulaIssue = "";
                    string standardSuffix;
                    bool formulaResolved = false;
                    bool standardResolved = false;
                    if (hasFormula)
                    {
                        formulaResolved = TryResolveSmartFormula(snapshot, targetRows, row, target, currentQuotaUnit,
                            entryCode, signature, out formulaRule, out formulaQuantity, out formulaIssue);
                    }
                    else if (TryBuildExcelLinkUnitScaleSuffix(row.Unit, currentQuotaUnit, out standardSuffix))
                    {
                        item.QuantityText = (row.QuantityText ?? "") + standardSuffix;
                        item.AlignNote = AppendPreviewNote(item.AlignNote,
                            "\u6807\u51c6\u6362\u7b97 " + (String.IsNullOrEmpty(standardSuffix) ? "1:1" : standardSuffix));
                        standardResolved = true;
                    }

                    if (formulaResolved)
                    {
                        item.QuantityText = formulaQuantity;
                        item.FormulaTemplate = formulaRule.Template;
                        item.FormulaOperands = formulaRule.Operands.OrderBy(operand => operand.Index).Select(operand => new QuantityFormulaOperandInfo
                        {
                            Name = operand.Name,
                            Unit = operand.Unit,
                            Signature = operand.Signature
                        }).ToList();
                        item.AlignNote = AppendPreviewNote(item.AlignNote, "换算公式命中(样本" + formulaRule.SampleCount.ToString(CultureInfo.InvariantCulture) + ")");
                    }
                    else if (!standardResolved)
                    {
                        item.QuantityText = row.QuantityText;
                        item.Selected = false;
                        item.NeedExactNameConfirmation = true;
                        string confirmedCountSuffix;
                        if (!hasFormula && TryBuildConfirmedCountUnitScaleSuffix(row.Unit, currentQuotaUnit, out confirmedCountSuffix))
                        {
                            item.Status = "\u5f85\u786e\u8ba4\u8ba1\u6570\u5355\u4f4d1:1";
                            item.AlignNote = AppendPreviewNote(item.AlignNote,
                                "\u8ba1\u6570\u57fa\u7840\u5355\u4f4d\u9700\u4eba\u5de5\u786e\u8ba4，\u786e\u8ba4\u540e\u5e94\u7528\u6570\u91cf\u7ea7" +
                                (String.IsNullOrEmpty(confirmedCountSuffix) ? " 1:1" : " " + confirmedCountSuffix));
                        }
                        else
                        {
                            item.Status = hasFormula ? "\u516c\u5f0f\u53c2\u6570\u7f3a\u5931\u6216\u6b67\u4e49" : "\u7f3a\u8de8\u91cf\u7eb2\u6362\u7b97\u7cfb\u6570";
                            item.AlignNote = AppendPreviewNote(item.AlignNote,
                                String.IsNullOrWhiteSpace(formulaIssue) ? ("\u5355\u4f4d " + row.Unit + "→" + currentQuotaUnit + " \u65e0\u53ef\u9760\u6362\u7b97\u516c\u5f0f") : formulaIssue);
                        }
                    }
                }

                if (currentQuota != null && !currentQuota.IsLibrary)
                {
                    item.ChosenQuotaSeq = currentQuota.QuotaSeq;   // 项目内已有该定额:整行复制,单价随项目
                    if (hasEntry) { item.ChosenItemSeq = entrySeq; item.ChosenItemNo = entryCode; }
                }
                else
                {
                    item.IsLibraryQuota = true;                 // 项目内没有:优先跨库整行复制,无溯源再原生粘贴
                    item.ChosenItemNo = hasEntry ? entryCode : "";
                    if (hasEntry) item.ChosenItemSeq = entrySeq;
                    SmartQuotaSource crossSource;
                    if (snapshot.CrossSourceByQuota.TryGetValue(target.Code, out crossSource))
                    {
                        item.SourceDb = crossSource.Db;
                        item.SourceDbQuotaSeq = crossSource.QuotaSeq;
                    }
                }
                item.ItemNo = hasEntry ? entryCode : "";
                if (!hasEntry)
                {
                    item.Status = AppendPreviewNote(item.Status, "缺条目");
                    item.AlignNote = AppendPreviewNote(item.AlignNote,
                        String.IsNullOrWhiteSpace(resolution.Issue) ? "学习库未定位到目标项目里的条目,请手选" : resolution.Issue);
                }
                else if (!String.IsNullOrWhiteSpace(resolution.Issue))
                {
                    item.Selected = false;
                    item.NeedExactNameConfirmation = true;
                    item.AlignNote = AppendPreviewNote(item.AlignNote, resolution.Issue);
                }
                groupItems.Add(item);
                order++;
            }
            if (groupItems.Any(item => !String.IsNullOrWhiteSpace(item.Status)))
            {
                foreach (FillPreviewItem groupItem in groupItems)
                {
                    groupItem.Selected = false;
                    groupItem.NeedExactNameConfirmation = true;
                    groupItem.AlignNote = AppendPreviewNote(groupItem.AlignNote, "组件组需要确认或存在阻断，确认前整组不写入");
                }
            }
            items.AddRange(groupItems);
        }

        // 智能铺量预览:漏斗匹配整张目标 sheet。
        private static List<FillPreviewItem> BuildPreview_SmartFill(Form mainForm,
            string targetWorkbook, string targetSheet, string targetColumn, SmartLearningScope scope, out string warning)
        {
            warning = null;
            SmartLearningScope selectedScope = scope ?? SmartLearningScope.CreateAll();
            CellRef colRef;
            if (!TryParseCellAddress((targetColumn ?? "").Trim().ToUpperInvariant() + "1", out colRef))
            {
                warning = "目标列无效。";
                return new List<FillPreviewItem>();
            }
            string workbook = (targetWorkbook ?? "").Trim();
            if (String.IsNullOrWhiteSpace(workbook) || !File.Exists(workbook))
            {
                warning = "目标 Excel 未选择或文件不存在,请先保存后重试。";
                return new List<FillPreviewItem>();
            }

            List<TargetQtyRow> targetRows = ReadTargetQtyRows(workbook, targetSheet, colRef.Column);
            if (targetRows.Count == 0)
            {
                warning = "目标 sheet 未读到工程量行(检查目标列是否为数量列,Excel 是否已保存)。";
                return new List<FillPreviewItem>();
            }

            string method = "";
            SmartMethodRoute route = ResolveSmartMethodRoute("");
            Dictionary<string, long> projectEntries = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> projectEntryNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                SqlConnection conn = GetOpenProjectConnection(mainForm);
                method = SmartResolveProjectMethod(conn);
                route = ResolveSmartMethodRoute(method);
                projectEntries = LoadSmartProjectEntries(conn, out projectEntryNames);
            }
            catch (Exception ex)
            {
                Log("Smart fill project context failed: " + ex.Message);
                warning = "无法读取当前项目编制办法，已停止推荐：" + ex.Message;
                return new List<FillPreviewItem>();
            }

            string snapshotNote;
            SmartLearningSnapshot snapshot = LoadSmartLearningSnapshot(route.LearningMethod,
                route.LibraryMethod, route.MethodNo, out snapshotNote);
            snapshot.SelectedScope = selectedScope;
            snapshot.ProjectEntryNameByCode = projectEntryNames;
            if (snapshot.BySignature.Count == 0)
            {
                warning = snapshotNote ?? "学习库为空,请先积累绑定或运行收割。";
                return new List<FillPreviewItem>();
            }

            Dictionary<string, ProjectQuota> currentQuotaByCode = LoadCurrentSmartQuotaMetadata(mainForm, snapshot);

            List<FillPreviewItem> items = new List<FillPreviewItem>();
            int hitExact = 0, hitNameOnly = 0, fuzzyRows = 0, manualRows = 0;
            HashSet<string> preferredPrefixes = new HashSet<string>(StringComparer.Ordinal);
            // 两遍扫描:第一遍用签名级证据对整表做工程前缀投票;第二遍用投票前缀消歧条目候选与多组歧义。
            for (int pass = 0; pass < 2; pass++)
            {
            items = new List<FillPreviewItem>();
            hitExact = 0; hitNameOnly = 0; fuzzyRows = 0; manualRows = 0;
            Dictionary<string, int> prefixVotes = new Dictionary<string, int>(StringComparer.Ordinal);
            foreach (TargetQtyRow row in targetRows)
            {
                string nameSig = NormalizeForSignature(row.RawName);
                string unitSig = NormalizeForSignature(row.Unit);
                string fullSig = nameSig + "|" + unitSig;
                if (fullSig.Length > 450) fullSig = fullSig.Substring(0, 450);
                string nameLevelSig = nameSig + "|";
                if (nameLevelSig.Length > 450) nameLevelSig = nameLevelSig.Substring(0, 450);

                List<SmartMapEntry> exactHits;
                snapshot.BySignature.TryGetValue(nameLevelSig, out exactHits);
                List<SmartMapEntry> nameHits;
                snapshot.ByNameOnly.TryGetValue(nameSig, out nameHits);
                bool scoped = !String.Equals(selectedScope.Kind, "All", StringComparison.OrdinalIgnoreCase);
                string scopedSource = scoped ? "专业学习库命中" : "学习库命中";
                bool matched = false;
                List<SmartMapEntry> filtered = FilterSmartHitsByScope(snapshot, exactHits, selectedScope);
                if (filtered.Count > 0)
                {
                    matched = AppendRankedSmartMatch(items, row, targetRows, filtered, snapshot, projectEntries, currentQuotaByCode,
                        "名称学习命中，" + scopedSource, nameLevelSig, preferredPrefixes, prefixVotes);
                    if (matched) hitExact++;
                }
                if (!matched)
                {
                    filtered = FilterSmartHitsByScope(snapshot, nameHits, selectedScope);
                    if (filtered.Count > 0)
                    {
                        matched = AppendRankedSmartMatch(items, row, targetRows, filtered, snapshot, projectEntries, currentQuotaByCode,
                            "名称兼容命中，" + scopedSource, nameLevelSig, preferredPrefixes, prefixVotes);
                        if (matched) hitNameOnly++;
                    }
                }

                // 模糊层只在当前选择范围内打分；范围内未命中直接转人工。
                List<KeyValuePair<int, string>> scored = BuildSmartFuzzyScoresIfUnmatched(matched, nameSig, snapshot.NameFeatures);
                List<NameQuotaCandidateGroup> fuzzyCandidates = new List<NameQuotaCandidateGroup>();
                string fuzzySourceNote = "";
                if (!matched)
                {
                    fuzzyCandidates = BuildSmartFuzzyCandidateGroups(scored, selectedScope, scopedSource, snapshot,
                        nameLevelSig, preferredPrefixes, projectEntries, currentQuotaByCode, row, targetRows);
                    if (fuzzyCandidates.Count > 0) fuzzySourceNote = scopedSource;
                }
                if (matched) continue;

                FillPreviewItem manual = new FillPreviewItem
                {
                    TemplateName = "推荐定额",
                    IsNameDriven = true,
                    NeedManualQuota = true,
                    TargetRow = row.Row,
                    TargetName = row.DisplayName,
                    TargetFullName = row.RawName,
                    TargetChapter = row.Chapter,
                    TargetUnit = row.Unit,
                    TargetQuantityText = row.QuantityText,
                    QuantityText = row.QuantityText,
                    OrderInItem = row.Row * 10,
                    Status = "未匹配",
                    Selected = false
                };
                if (fuzzyCandidates.Count > 0)
                {
                    manual.NameQuotaCandidates = fuzzyCandidates;
                    if (manual.NameQuotaCandidates.Count > 0)
                    {
                        manual.AlignNote = (String.IsNullOrWhiteSpace(fuzzySourceNote) ? "" : fuzzySourceNote + "，") +
                            "有 " + manual.NameQuotaCandidates.Count.ToString(CultureInfo.InvariantCulture) + " 个模糊候选,双击选择";
                        fuzzyRows++;
                    }
                    else manualRows++;
                }
                else manualRows++;
                items.Add(manual);
            }

            if (pass == 0)
            {
                preferredPrefixes = new HashSet<string>(prefixVotes.Keys, StringComparer.Ordinal);
                if (preferredPrefixes.Count == 0) break;   // 无签名级证据,一遍结果即最终结果
            }
            }

            warning = "学习库(SQL):精确 " + hitExact.ToString(CultureInfo.InvariantCulture) +
                " 行,同名 " + hitNameOnly.ToString(CultureInfo.InvariantCulture) +
                " 行,模糊候选 " + fuzzyRows.ToString(CultureInfo.InvariantCulture) +
                " 行,待手挂 " + manualRows.ToString(CultureInfo.InvariantCulture) + " 行。" +
                (preferredPrefixes.Count > 0 ? " 工程前缀:" + String.Join("/", preferredPrefixes.ToArray()) + "。" : "") +
                (snapshotNote != null ? " " + snapshotNote : "");
            return items;
        }
    }
}
