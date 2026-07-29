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
        // ============ 智能铺量(学习库)漏斗匹配引擎 ============
        // 漏斗:①名称级签名 ②模糊候选(仅进下拉,不自动采纳) ③手挂。
        // 数据源:RecoLearning(SQL);连不上时回退 mapping-boxes.jsonl(无条目知识)。

        private sealed class SmartBoxTarget
        {
            public string Kind; public string Code; public string Name; public string Unit;
        }

        private sealed class SmartMapEntry
        {
            public string BoxId;
            public int Weight;
            public DateTime LastUsedAt;
            public List<SmartBoxTarget> Targets = new List<SmartBoxTarget>();
            public bool CurrentMethodMapping;
            // 本机 mapping-boxes 可携带办法/条目；保留成配对键，避免把不同样本的办法和条目交叉组合。
            public HashSet<string> LocalContextKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
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
            public bool PendingLocal;
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
            return name + "|";
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

        private static string BuildSmartLocalSignature(Dictionary<string, string> row)
        {
            return BuildSmartQuantitySignature(GetFlat(row, "quantity_name"), GetFlat(row, "quantity_unit"));
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

        private static string BuildSmartLocalContextKey(string method, string entryCode)
        {
            string normalizedMethod = NormalizeSmartProjectMethod(method);
            string normalizedEntry = (entryCode ?? "").Trim();
            return normalizedMethod.Length == 0 || normalizedEntry.Length == 0 ? "" : normalizedMethod + "\n" + normalizedEntry;
        }

        private static void AddSmartLocalContext(SmartMapEntry entry, Dictionary<string, string> row)
        {
            if (entry == null || row == null) return;
            string method = GetSmartLocalMethod(row);
            string entryCode = GetFlat(row, "entry_code");
            if (String.IsNullOrWhiteSpace(entryCode)) entryCode = GetFlat(row, "formula_entry_code");
            string key = BuildSmartLocalContextKey(method, entryCode);
            if (key.Length > 0) entry.LocalContextKeys.Add(key);
        }

        private static string GetSmartLocalMethod(Dictionary<string, string> row)
        {
            if (row == null) return "";
            string method = GetFlat(row, "method");
            if (String.IsNullOrWhiteSpace(method)) method = GetFlat(row, "formula_method");
            return NormalizeSmartProjectMethod(method);
        }

        private static List<Dictionary<string, string>> SelectSmartLocalRowsForMethod(
            IEnumerable<Dictionary<string, string>> rows, string method)
        {
            string currentMethod = NormalizeSmartProjectMethod(method);
            List<Dictionary<string, string>> compatible = (rows ?? Enumerable.Empty<Dictionary<string, string>>())
                .Where(row =>
                {
                    if (ReadFlatInt(row, "weight", 1) <= 0) return false;
                    string rowMethod = GetSmartLocalMethod(row);
                    return rowMethod.Length == 0 || String.Equals(rowMethod, currentMethod, StringComparison.OrdinalIgnoreCase);
                }).ToList();
            List<Dictionary<string, string>> exact = compatible.Where(row =>
                !String.IsNullOrWhiteSpace(currentMethod) &&
                String.Equals(GetSmartLocalMethod(row), currentMethod, StringComparison.OrdinalIgnoreCase)).ToList();
            return exact.Count > 0 ? exact : compatible.Where(row => GetSmartLocalMethod(row).Length == 0).ToList();
        }

        private static string ResolveSmartSqlSignature(string storedSignature, Dictionary<string, string> legacyAliases)
        {
            string mapped;
            if (legacyAliases != null && legacyAliases.TryGetValue(storedSignature ?? "", out mapped)) return mapped;
            string normalized = NormalizeSmartLearningSignature(storedSignature);
            return legacyAliases != null && legacyAliases.TryGetValue(normalized, out mapped) ? mapped : normalized;
        }

        private static void AddSmartLocalFormulaRule(SmartLearningSnapshot snapshot, Dictionary<string, string> row, string signature, bool pendingLocal)
        {
            string template = GetFlat(row, "formula_template");
            int operandCount = ReadFlatInt(row, "formula_operand_count", 0);
            if (String.IsNullOrWhiteSpace(template) || operandCount <= 0) return;
            int sampleCount = Math.Max(1, ReadFlatInt(row, "accepted_count", 1));
            DateTime lastSeen;
            if (!DateTime.TryParse(GetFlat(row, "last_used_at"), out lastSeen)) lastSeen = DateTime.MinValue;
            string kind = GetFlat(row, "target_kind");
            string code = GetFlat(row, "target_code");
            string targetUnit = GetFlat(row, "formula_target_unit");
            string method = GetFlat(row, "formula_method");
            string entryCode = GetFlat(row, "formula_entry_code");
            string ruleRaw = signature + "|" + (String.IsNullOrWhiteSpace(kind) ? "quota" : kind) + ":" + code.ToUpperInvariant() + "|" +
                NormalizeExcelLinkUnit(targetUnit) + "|" + template + "|" + method + "|" + entryCode;
            List<SmartFormulaOperand> operands = new List<SmartFormulaOperand>();
            for (int operandIndex = 0; operandIndex < operandCount; operandIndex++)
            {
                string prefix = "formula_operand_" + operandIndex.ToString(CultureInfo.InvariantCulture) + "_";
                SmartFormulaOperand operand = new SmartFormulaOperand
                {
                    Index = operandIndex,
                    Signature = NormalizeSmartLearningSignature(GetFlat(row, prefix + "signature")),
                    Name = GetFlat(row, prefix + "name"),
                    Unit = NormalizeExcelLinkUnit(GetFlat(row, prefix + "unit"))
                };
                if (String.IsNullOrWhiteSpace(operand.Signature.Trim('|'))) return;
                operands.Add(operand);
                ruleRaw += "|" + operand.Signature + "@" + operand.Unit;
            }
            string ruleHash = GetFlat(row, "formula_rule_hash");
            if (String.IsNullOrWhiteSpace(ruleHash)) ruleHash = BuildLearningMd5(ruleRaw);
            SmartFormulaRule rule = AddSmartFormulaRule(snapshot, ruleHash, signature, kind, code, targetUnit,
                template, method, entryCode, sampleCount, lastSeen);
            if (rule != null)
            {
                if (rule.Operands.Count == 0) rule.Operands.AddRange(operands);
                if (pendingLocal) rule.PendingLocal = true;
            }
        }

        // 当前项目办法:项目信息.编制办法文号 含2024->2024,含2020->2020,否则原文。
        private static string SmartResolveProjectMethod(SqlConnection projectConn)
        {
            try
            {
                using (SqlCommand cmd = projectConn.CreateCommand())
                {
                    cmd.CommandText = "SELECT TOP 1 编制办法文号 FROM 项目信息";
                    object raw = cmd.ExecuteScalar();
                    string text = raw != null ? raw.ToString() : "";
                    return NormalizeSmartProjectMethod(text);
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

        // 从 RecoLearning 加载快照;失败回退 jsonl(仅签名映射,无条目知识)。
        private static SmartLearningSnapshot LoadSmartLearningSnapshot(string method, out string note)
        {
            note = null;
            SmartLearningSnapshot snapshot = new SmartLearningSnapshot { Method = NormalizeSmartProjectMethod(method) };
            try
            {
                // 推荐页每次打开都先尝试重放本机待同步批次；SQL 快照随后实时读取最新聚合。
                ReplayPendingLearningDbEvents();
            }
            catch (Exception ex)
            {
                Log("Smart fill replay pending learning events failed: " + ex.Message);
            }
            try
            {
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
                            "SELECT m.signature, m.box_id, m.method, m.weight, t.target_kind, t.target_code, t.target_name, t.target_unit, m.last_used_at " +
                            "FROM dbo.SignatureBoxMap m JOIN dbo.QuotaBoxTarget t ON t.box_id = m.box_id " +
                            "WHERE m.weight > 0 AND (m.method=@map_method OR m.method='') " +
                            "ORDER BY CASE WHEN m.method=@map_method THEN 0 ELSE 1 END";
                        cmd.Parameters.AddWithValue("@map_method", snapshot.Method ?? "");
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string signature = ResolveSmartSqlSignature(reader.GetString(0), legacySignatureAliases);
                                string boxId = reader.GetString(1);
                                string mappingMethod = reader.IsDBNull(2) ? "" : NormalizeSmartProjectMethod(reader.GetString(2));
                                bool currentMethodMapping = !String.IsNullOrWhiteSpace(snapshot.Method) &&
                                    String.Equals(mappingMethod, snapshot.Method, StringComparison.OrdinalIgnoreCase);
                                string key = signature + "\n" + boxId;
                                SmartMapEntry entry;
                                if (!byKey.TryGetValue(key, out entry))
                                {
                                    entry = new SmartMapEntry
                                    {
                                        BoxId = boxId,
                                        Weight = reader.GetInt32(3),
                                        CurrentMethodMapping = currentMethodMapping,
                                        LastUsedAt = reader.IsDBNull(8) ? DateTime.MinValue : reader.GetDateTime(8)
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
                                    if (currentMethodMapping && !entry.CurrentMethodMapping) entry.Weight = reader.GetInt32(3);
                                    else if (currentMethodMapping == entry.CurrentMethodMapping) entry.Weight = Math.Max(entry.Weight, reader.GetInt32(3));
                                    if (currentMethodMapping) entry.CurrentMethodMapping = true;
                                    DateTime lastUsed = reader.IsDBNull(8) ? DateTime.MinValue : reader.GetDateTime(8);
                                    if (lastUsed > entry.LastUsedAt) entry.LastUsedAt = lastUsed;
                                }
                                UpsertSmartBoxTarget(entry, reader.GetString(4), reader.GetString(5), reader.GetString(6), reader.GetString(7));
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
                                    "FROM dbo.QuantityFormulaRule WHERE method=@method OR method=''";
                                cmd.Parameters.AddWithValue("@method", snapshot.Method ?? "");
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
                                    "WHERE r.method=@method OR r.method='' ORDER BY o.rule_hash,o.operand_index";
                                cmd.Parameters.AddWithValue("@method", snapshot.Method ?? "");
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
                        using (SqlCommand cmd = conn.CreateCommand())
                        {
                            cmd.CommandTimeout = 15;
                            cmd.CommandText =
                                "SELECT quota_code, entry_code, entry_name, project_count FROM dbo.EntryQuota " +
                                "WHERE method = @method AND target_kind = 'quota'";
                            cmd.Parameters.AddWithValue("@method", snapshot.Method);
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
                                "WHERE method = @m3 OR method = ''";
                            cmd.Parameters.AddWithValue("@m3", snapshot.Method);
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
                                        CurrentMethodEvidence = String.Equals(reader.GetString(2), snapshot.Method, StringComparison.OrdinalIgnoreCase)
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
                                "WHERE source = 'import:excel-links' AND project_id <> '' AND target_kind = 'quota' AND method = @m2 ORDER BY id";
                            cmd.Parameters.AddWithValue("@m2", snapshot.Method);
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
                                "WHERE entry_code <> '' AND target_kind = 'quota' AND (method = @method OR method = '') " +
                                "GROUP BY target_code, entry_code";
                            cmd.Parameters.AddWithValue("@method", snapshot.Method);
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

                MergePendingLocalMappingsIntoSmartSnapshot(snapshot, LoadPendingLearningMappingKeys());

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
                return snapshot;
            }
            catch (Exception ex)
            {
                Log("Smart fill SQL snapshot failed, fallback to jsonl: " + ex.Message);
            }

            // —— jsonl 回退:借用 mapping-boxes,转成同构快照(无条目知识)。 ——
            try
            {
                List<Dictionary<string, string>> boxRows = LoadMappingBoxRows();
                foreach (IGrouping<string, Dictionary<string, string>> boxGroup in boxRows
                    .Where(row => !String.IsNullOrWhiteSpace(GetFlat(row, "box_id")))
                    .GroupBy(row => GetFlat(row, "box_id"), StringComparer.OrdinalIgnoreCase))
                {
                    foreach (IGrouping<string, Dictionary<string, string>> signatureGroup in boxGroup
                        .GroupBy(row => BuildSmartLocalSignature(row), StringComparer.OrdinalIgnoreCase))
                    {
                        List<Dictionary<string, string>> sample = SelectSmartLocalRowsForMethod(signatureGroup, snapshot.Method);
                        if (sample.Count == 0) continue;
                        string signature = signatureGroup.Key;
                        SmartMapEntry entry = new SmartMapEntry
                        {
                            BoxId = boxGroup.Key,
                            Weight = sample.Max(row => ReadFlatInt(row, "weight", 0)),
                            CurrentMethodMapping = sample.Any(row => String.Equals(GetSmartLocalMethod(row), snapshot.Method, StringComparison.OrdinalIgnoreCase) &&
                                !String.IsNullOrWhiteSpace(snapshot.Method))
                        };
                        foreach (Dictionary<string, string> row in sample)
                        {
                            UpsertSmartBoxTarget(entry, GetFlat(row, "target_kind"), GetFlat(row, "target_code"),
                                GetFlat(row, "target_name"), GetFlat(row, "target_unit"));
                            AddSmartLocalContext(entry, row);
                            AddSmartLocalFormulaRule(snapshot, row, signature, false);
                        }
                        List<SmartMapEntry> list;
                        if (!snapshot.BySignature.TryGetValue(signature, out list)) { list = new List<SmartMapEntry>(); snapshot.BySignature[signature] = list; }
                        list.Add(entry);
                        string nameSeg = SmartNameSegment(signature);
                        if (!snapshot.ByNameOnly.TryGetValue(nameSeg, out list)) { list = new List<SmartMapEntry>(); snapshot.ByNameOnly[nameSeg] = list; }
                        list.Add(entry);
                    }
                }
                foreach (string nameSeg in snapshot.ByNameOnly.Keys)
                {
                    snapshot.NameFeatures.Add(new KeyValuePair<string, MatchTextFeatures>(nameSeg, BuildMatchTextFeatures(nameSeg)));
                }
                snapshot.FromSql = false;
                note = "学习库(SQL)不可用,已回退本地 mapping-boxes；仅带当前办法/条目字段的新关系可自动定位。";
                return snapshot;
            }
            catch (Exception ex)
            {
                Log("Smart fill jsonl fallback failed: " + ex.Message);
                note = "学习库与本地映射均不可用:" + ex.Message;
                return snapshot;
            }
        }

        // SQL 主库优先；本机只叠加 outbox 中仍待同步的“名称签名+组件框”，不比较跨机器时间。
        private static void MergePendingLocalMappingsIntoSmartSnapshot(SmartLearningSnapshot snapshot, HashSet<string> pendingKeys)
        {
            if (snapshot == null || pendingKeys == null || pendingKeys.Count == 0) return;
            try
            {
                List<Dictionary<string, string>> boxRows = LoadMappingBoxRows();
                foreach (IGrouping<string, Dictionary<string, string>> boxGroup in boxRows
                    .Where(row => !String.IsNullOrWhiteSpace(GetFlat(row, "box_id")))
                    .GroupBy(row => GetFlat(row, "box_id"), StringComparer.OrdinalIgnoreCase))
                {
                    foreach (IGrouping<string, Dictionary<string, string>> signatureGroup in boxGroup
                        .GroupBy(row => BuildSmartLocalSignature(row), StringComparer.OrdinalIgnoreCase))
                    {
                        List<Dictionary<string, string>> sample = SelectSmartLocalRowsForMethod(signatureGroup, snapshot.Method);
                        if (sample.Count == 0) continue;
                        string signature = signatureGroup.Key;
                        string pendingKey = signature + "\n" + boxGroup.Key;
                        if (!pendingKeys.Contains(pendingKey)) continue;

                        List<SmartMapEntry> list;
                        if (!snapshot.BySignature.TryGetValue(signature, out list))
                        {
                            list = new List<SmartMapEntry>();
                            snapshot.BySignature[signature] = list;
                        }
                        SmartMapEntry existing = list.FirstOrDefault(entry => String.Equals(entry.BoxId, boxGroup.Key, StringComparison.OrdinalIgnoreCase));
                        SmartMapEntry localEntry = existing ?? new SmartMapEntry { BoxId = boxGroup.Key };
                        localEntry.Weight = Math.Max(sample.Max(row => ReadFlatInt(row, "weight", 0)), 1000);
                        localEntry.CurrentMethodMapping = sample.Any(row =>
                            String.Equals(GetSmartLocalMethod(row), snapshot.Method, StringComparison.OrdinalIgnoreCase) &&
                            !String.IsNullOrWhiteSpace(snapshot.Method));
                        localEntry.Targets = new List<SmartBoxTarget>();
                        foreach (Dictionary<string, string> row in sample)
                        {
                            UpsertSmartBoxTarget(localEntry, GetFlat(row, "target_kind"), GetFlat(row, "target_code"),
                                GetFlat(row, "target_name"), GetFlat(row, "target_unit"));
                            AddSmartLocalContext(localEntry, row);
                            AddSmartLocalFormulaRule(snapshot, row, signature, true);
                        }
                        if (localEntry.Targets.Count == 0) continue;
                        if (existing == null) list.Add(localEntry);

                        string nameSeg = SmartNameSegment(signature);
                        List<SmartMapEntry> nameList;
                        if (!snapshot.ByNameOnly.TryGetValue(nameSeg, out nameList))
                        {
                            nameList = new List<SmartMapEntry>();
                            snapshot.ByNameOnly[nameSeg] = nameList;
                        }
                        if (!nameList.Contains(localEntry)) nameList.Add(localEntry);
                    }
                }
            }
            catch (Exception ex)
            {
                Log("Smart fill pending local overlay failed: " + ex.Message);
            }
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
                if (quota != null && requiredCodes.Contains(quota.Code ?? "") && !result.ContainsKey(quota.Code)) result[quota.Code] = quota;
            }

            try
            {
                string path = Path.Combine(FindRecoQuotaDataDir(), "quota-index.jsonl");
                foreach (KeyValuePair<string, ProjectQuota> pair in LoadCachedSmartQuotaIndex(path))
                {
                    string code = pair.Key;
                    ProjectQuota indexed = pair.Value;
                    if (!requiredCodes.Contains(code) || indexed == null) continue;
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

        // 为一组定额目标解析放置条目。证据优先级:①签名级(该工程量配该定额历史放过的条目)
        // ②定额级(EntryByQuota),带工程前缀过滤(过滤后为空则放开)。
        private static bool TryResolveSmartEntry(SmartLearningSnapshot snapshot, Dictionary<string, long> projectEntries,
            SmartMapEntry mappingEntry, List<SmartBoxTarget> targets, string signature, HashSet<string> preferredPrefixes,
            out string entryCode, out string entryName, out long entrySeq, out bool fromSignature)
        {
            entryCode = ""; entryName = ""; entrySeq = 0; fromSignature = false;

            // ① 签名级证据:先试完整签名,再试去单位签名(存量老数据单位为空)。
            string[] sigKeys = new string[] { signature ?? "", SmartNameSegment(signature ?? "") + "|" };
            List<SmartEntryStat> genericSignatureStats = new List<SmartEntryStat>();
            List<SmartEntryStat> currentSignatureStats = new List<SmartEntryStat>();
            foreach (SmartBoxTarget target in targets)
            {
                if (target == null || !String.Equals(target.Kind, "quota", StringComparison.OrdinalIgnoreCase)) continue;
                foreach (string sigKey in sigKeys)
                {
                    List<SmartEntryStat> stats;
                    if (!snapshot.EntryBySignatureQuota.TryGetValue(sigKey + "\n" + (target.Code ?? ""), out stats)) continue;
                    foreach (SmartEntryStat stat in stats)
                    {
                        if (!projectEntries.ContainsKey(stat.EntryCode)) continue;
                        if (stat.CurrentMethodEvidence)
                        {
                            currentSignatureStats.Add(stat);
                            continue;
                        }
                        genericSignatureStats.Add(stat);
                    }
                }
            }

            List<SmartEntryStat> distinctCurrentStats = currentSignatureStats
                .GroupBy(stat => stat.EntryCode, StringComparer.OrdinalIgnoreCase)
                .Select(group => group.OrderByDescending(stat => stat.ProjectCount).First()).ToList();
            if (preferredPrefixes != null && preferredPrefixes.Count > 0)
            {
                List<SmartEntryStat> prefixed = distinctCurrentStats.Where(stat => stat.EntryCode != null && stat.EntryCode.Length >= 2 &&
                    preferredPrefixes.Contains(stat.EntryCode.Substring(0, 2))).ToList();
                if (prefixed.Count > 0) distinctCurrentStats = prefixed;
            }
            if (distinctCurrentStats.Count == 1)
            {
                entryCode = distinctCurrentStats[0].EntryCode;
                entryName = ResolveSmartEntryName(snapshot, entryCode, distinctCurrentStats[0].EntryName);
                entrySeq = projectEntries[entryCode];
                fromSignature = true;
                return true;
            }

            // 本机待同步/SQL不可用回退关系可直接携带当前办法+稳定条目编号。
            if (mappingEntry != null && mappingEntry.LocalContextKeys != null)
            {
                string methodPrefix = (snapshot.Method ?? "").Trim() + "\n";
                List<string> localEntries = mappingEntry.LocalContextKeys
                    .Where(key => key.StartsWith(methodPrefix, StringComparison.OrdinalIgnoreCase))
                    .Select(key => key.Substring(methodPrefix.Length))
                    .Where(code => projectEntries.ContainsKey(code))
                    .Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                if (localEntries.Count == 1)
                {
                    entryCode = localEntries[0];
                    entryName = ResolveSmartEntryName(snapshot, entryCode, "");
                    entrySeq = projectEntries[entryCode];
                    fromSignature = true;
                    return true;
                }
            }

            // 当前办法有多条并列条目证据时仅提供一个预览落点，不把它标记为可自动采纳的唯一上下文。
            if (distinctCurrentStats.Count > 1)
            {
                SmartEntryStat ambiguous = distinctCurrentStats.OrderByDescending(stat => stat.ProjectCount).First();
                entryCode = ambiguous.EntryCode;
                entryName = ResolveSmartEntryName(snapshot, entryCode, ambiguous.EntryName);
                entrySeq = projectEntries[entryCode];
                return true;
            }

            // 空办法的历史签名只用于定位候选条目，不算当前办法的唯一上下文，调用方必须要求确认。
            SmartEntryStat generic = genericSignatureStats.OrderByDescending(stat => stat.ProjectCount).FirstOrDefault();
            if (generic != null)
            {
                entryCode = generic.EntryCode;
                entryName = ResolveSmartEntryName(snapshot, entryCode, generic.EntryName);
                entrySeq = projectEntries[entryCode];
                return true;
            }

            // ② 定额级证据,先带前缀过滤,过滤后为空再放开。
            SmartEntryStat best = FindBestQuotaEntry(snapshot, projectEntries, targets, preferredPrefixes);
            if (best == null && preferredPrefixes != null && preferredPrefixes.Count > 0)
            {
                best = FindBestQuotaEntry(snapshot, projectEntries, targets, null);
            }
            if (best == null) return false;
            entryCode = best.EntryCode;
            entryName = ResolveSmartEntryName(snapshot, entryCode, best.EntryName);
            entrySeq = projectEntries[best.EntryCode];
            return true;
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
            return (learnedName ?? "").Trim();
        }

        private static SmartEntryStat FindBestQuotaEntry(SmartLearningSnapshot snapshot, Dictionary<string, long> projectEntries,
            List<SmartBoxTarget> targets, HashSet<string> preferredPrefixes)
        {
            SmartEntryStat best = null;
            foreach (SmartBoxTarget target in targets)
            {
                if (target == null || !String.Equals(target.Kind, "quota", StringComparison.OrdinalIgnoreCase)) continue;
                List<SmartEntryStat> stats;
                if (!snapshot.EntryByQuota.TryGetValue(target.Code ?? "", out stats)) continue;
                foreach (SmartEntryStat stat in stats)
                {
                    if (!projectEntries.ContainsKey(stat.EntryCode)) continue;
                    if (preferredPrefixes != null && preferredPrefixes.Count > 0 &&
                        (stat.EntryCode.Length < 2 || !preferredPrefixes.Contains(stat.EntryCode.Substring(0, 2)))) continue;
                    if (best == null || stat.ProjectCount > best.ProjectCount) best = stat;
                    break; // stats 已按证据强度降序,该定额取第一个满足条件的即可
                }
            }
            return best;
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

        // 跨库整行复制:按学习库溯源从来源项目库拉取源定额行(依次尝试本项目服务器与两台已知服务器)。
        private static Dictionary<string, object> LoadCrossDbQuotaRow(SqlConnection targetConn, FillPreviewItem item, HashSet<string> targetColumns)
        {
            List<string> servers = new List<string>();
            try { if (!String.IsNullOrEmpty(targetConn.DataSource)) servers.Add(targetConn.DataSource); } catch { }
            if (!servers.Contains(LearningDbServer)) servers.Add(LearningDbServer);
            if (!servers.Contains("192.168.2.13,1433")) servers.Add("192.168.2.13,1433");

            foreach (string server in servers)
            {
                try
                {
                    string connectionString = "Server=" + server + ";Database=" + item.SourceDb + ";User ID=" + AgentDbUser + ";Password=" + AgentDbPassword + ";Connect Timeout=3";
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
                    Log("Cross-db quota row load failed on " + server + "/" + item.SourceDb + ": " + ex.Message);
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
                string entryCode; string entryName; long entrySeq; bool fromContext;
                bool hasEntry = TryResolveSmartEntry(snapshot, projectEntries, hit, hit.Targets, signature, preferredPrefixes,
                    out entryCode, out entryName, out entrySeq, out fromContext);
                bool prefixMatch = hasEntry && entryCode.Length >= 2 && preferredPrefixes != null && preferredPrefixes.Count > 0 &&
                    preferredPrefixes.Contains(entryCode.Substring(0, 2));
                string currentContextPrefix = (snapshot.Method ?? "").Trim() + "\n";
                bool currentMethodMapping = hit.CurrentMethodMapping || (hit.LocalContextKeys != null &&
                    !String.IsNullOrWhiteSpace(snapshot.Method) && hit.LocalContextKeys.Any(key =>
                        key.StartsWith(currentContextPrefix, StringComparison.OrdinalIgnoreCase)));
                bool targetsValid = hit.Targets.Count > 0 && hit.Targets.All(target =>
                {
                    ProjectQuota current;
                    return target != null && String.Equals(target.Kind ?? "quota", "quota", StringComparison.OrdinalIgnoreCase) &&
                        !String.IsNullOrWhiteSpace(target.Code) && currentQuotaByCode.TryGetValue(target.Code, out current) &&
                        current != null && !String.IsNullOrWhiteSpace(current.Unit);
                });
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
                    CurrentTargetsValid = targetsValid
                });
            }
            return scores.OrderByDescending(score => score.CurrentTargetsValid)
                .ThenByDescending(score => score.HasCurrentMethodMapping)
                .ThenByDescending(score => score.HasCurrentContext)
                .ThenByDescending(score => score.PrefixMatch)
                .ThenByDescending(score => score.HasEntry)
                .ThenByDescending(score => score.Entry == null ? 0 : score.Entry.Weight)
                .ThenByDescending(score => score.Entry == null ? DateTime.MinValue : score.Entry.LastUsedAt)
                .ThenBy(score => score.Entry == null ? "" : score.Entry.BoxId, StringComparer.OrdinalIgnoreCase).ToList();
        }

        private static bool CanAutoSelectSmartMapEntry(List<SmartMapCandidateScore> scores)
        {
            if (scores == null || scores.Count == 0) return false;
            SmartMapCandidateScore top = scores[0];
            if (!top.CurrentTargetsValid || !top.HasCurrentMethodMapping || !top.HasEntry || !top.HasCurrentContext) return false;
            if (scores.Count == 1) return true;

            SmartMapCandidateScore second = scores[1];
            if (!second.CurrentTargetsValid) return true;
            int topWeight = top.Entry == null ? 0 : top.Entry.Weight;
            int secondWeight = second.Entry == null ? 0 : second.Entry.Weight;
            // 多个当前版本均有效的组件，权重差不足两个接受样本时不静默决定。
            return topWeight - secondWeight >= 20;
        }

        private static string BuildSmartCandidateLabel(SmartMapCandidateScore score)
        {
            if (score == null || score.Entry == null) return "空组件";
            string targets = String.Join(" + ", score.Entry.Targets.Where(target => target != null)
                .Select(target => (target.Code ?? "").Trim()).Where(code => code.Length > 0).ToArray());
            if (String.IsNullOrWhiteSpace(score.EntryCode)) return targets + "（缺条目）";
            return targets + "（条目 " + score.EntryCode.Trim() +
                (String.IsNullOrWhiteSpace(score.EntryName) ? "" : " " + score.EntryName.Trim()) + "）";
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

        private static void AppendRankedSmartMatch(List<FillPreviewItem> items, TargetQtyRow row, List<TargetQtyRow> targetRows,
            List<SmartMapEntry> hits, SmartLearningSnapshot snapshot, Dictionary<string, long> projectEntries,
            Dictionary<string, ProjectQuota> currentQuotaByCode, string baseNote, string signature,
            HashSet<string> preferredPrefixes, Dictionary<string, int> prefixVotes)
        {
            List<SmartMapCandidateScore> scores = RankSmartMapEntries(snapshot, hits, signature, preferredPrefixes,
                projectEntries, currentQuotaByCode);
            if (scores.Count == 0) return;
            if (CanAutoSelectSmartMapEntry(scores))
            {
                AppendSmartItems(items, row, targetRows, scores[0].Entry, snapshot, projectEntries, currentQuotaByCode,
                    false, baseNote + "，" + BuildSmartCandidateLabel(scores[0]), signature, preferredPrefixes, prefixVotes);
                return;
            }

            List<NameQuotaCandidateGroup> candidates = new List<NameQuotaCandidateGroup>();
            foreach (SmartMapCandidateScore score in scores)
            {
                List<FillPreviewItem> candidateItems = new List<FillPreviewItem>();
                string label = BuildSmartCandidateLabel(score);
                AppendSmartItems(candidateItems, row, targetRows, score.Entry, snapshot, projectEntries, currentQuotaByCode,
                    true, baseNote + "，候选：" + label, signature, preferredPrefixes, null);
                if (candidateItems.Count == 0) continue;
                candidates.Add(new NameQuotaCandidateGroup
                {
                    Key = "smart:" + (score.Entry.BoxId ?? ""),
                    Label = label,
                    Items = candidateItems
                });
            }
            candidates = DeduplicateSmartCandidatesByLabel(candidates);
            if (candidates.Count == 0) return;

            List<FillPreviewItem> active = candidates[0].Items.Select(item => item.CloneForNameCandidate()).ToList();
            foreach (FillPreviewItem item in active)
            {
                item.Selected = false;
                item.NeedExactNameConfirmation = true;
            }
            active[0].NameQuotaCandidates = candidates;
            active[0].SelectedNameQuotaCandidateKey = candidates[0].Key;
            active[0].AlignNote = AppendPreviewNote(active[0].AlignNote,
                candidates.Count > 1 ? "组件候选接近或上下文不唯一，请确认完整组件" : "缺少唯一的当前办法/条目证据，请确认组件");
            items.AddRange(active);
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
                expression = System.Text.RegularExpressions.Regex.Replace(expression, pattern, "(" + operandText + ")");
            }
            if (System.Text.RegularExpressions.Regex.IsMatch(expression, "V[0-9]+")) return false;
            decimal result;
            string error;
            if (!TryEvaluateDecimal(expression + outputSuffix, out result, out error)) return false;
            if (result <= 0m) return false;
            quantityText = expression + outputSuffix;
            return true;
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
            List<SmartFormulaRule> contextual = rules.Where(rule =>
                String.Equals(rule.Method ?? "", snapshot.Method ?? "", StringComparison.OrdinalIgnoreCase) &&
                !String.IsNullOrWhiteSpace(rule.EntryCode) &&
                String.Equals(rule.EntryCode ?? "", entryCode ?? "", StringComparison.OrdinalIgnoreCase)).ToList();
            if (contextual.Count == 0)
            {
                contextual = rules.Where(rule => String.IsNullOrWhiteSpace(rule.EntryCode) &&
                    (String.IsNullOrWhiteSpace(rule.Method) ||
                     String.Equals(rule.Method ?? "", snapshot.Method ?? "", StringComparison.OrdinalIgnoreCase))).ToList();
            }
            return contextual;
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
                .ThenByDescending(item => item.Rule.PendingLocal)
                .ThenByDescending(item => item.Rule.SampleCount).ThenByDescending(item => item.Rule.LastSeen).ToList();
            if (valid.Count == 0)
            {
                issue = "换算公式参数未找齐、单位不符或存在重名歧义";
                return false;
            }
            if (valid.Count > 1 &&
                String.Equals(valid[0].Rule.Method ?? "", valid[1].Rule.Method ?? "", StringComparison.OrdinalIgnoreCase) &&
                valid[0].Rule.PendingLocal == valid[1].Rule.PendingLocal &&
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
            Dictionary<string, ProjectQuota> currentQuotaByCode, bool needConfirm, string note,
            string signature, HashSet<string> preferredPrefixes, Dictionary<string, int> prefixVotes)
        {
            string entryCode; string entryName; long entrySeq; bool fromSignature;
            bool hasEntry = TryResolveSmartEntry(snapshot, projectEntries, entry, entry.Targets, signature, preferredPrefixes,
                out entryCode, out entryName, out entrySeq, out fromSignature);
            if (hasEntry && fromSignature && entryCode.Length >= 2 && prefixVotes != null)
            {
                string prefix = entryCode.Substring(0, 2);
                int votes;
                prefixVotes.TryGetValue(prefix, out votes);
                prefixVotes[prefix] = votes + 1;
            }
            int order = 0;
            List<FillPreviewItem> groupItems = new List<FillPreviewItem>();
            foreach (SmartBoxTarget target in entry.Targets)
            {
                if (target == null || String.IsNullOrEmpty(target.Code)) continue;
                ProjectQuota currentQuota;
                currentQuotaByCode.TryGetValue(target.Code, out currentQuota);
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
                    bool hasFormula = snapshot.FormulaByKey.ContainsKey(formulaKey);
                    SmartFormulaRule formulaRule = null;
                    string formulaQuantity = "";
                    string formulaIssue = "";
                    string standardSuffix;
                    bool requiresDerivedFormula = hasFormula &&
                        HasContextualDerivedSmartFormula(snapshot, formulaKey, entryCode);
                    bool formulaResolved = false;
                    bool standardResolved = false;
                    if (requiresDerivedFormula)
                    {
                        formulaResolved = TryResolveDerivedSmartFormula(snapshot, targetRows, row, target, currentQuotaUnit,
                            entryCode, signature, out formulaRule, out formulaQuantity, out formulaIssue);
                    }
                    else if (TryBuildExcelLinkUnitScaleSuffix(row.Unit, currentQuotaUnit, out standardSuffix))
                    {
                        item.QuantityText = (row.QuantityText ?? "") + standardSuffix;
                        standardResolved = true;
                    }
                    else if (hasFormula)
                    {
                        formulaResolved = TryResolveSmartFormula(snapshot, targetRows, row, target, currentQuotaUnit, entryCode,
                            signature, out formulaRule, out formulaQuantity, out formulaIssue);
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
                        item.Status = "待确认换算";
                        item.AlignNote = AppendPreviewNote(item.AlignNote,
                            String.IsNullOrWhiteSpace(formulaIssue) ? ("单位 " + row.Unit + "→" + currentQuotaUnit + " 无可靠换算公式") : formulaIssue);
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
                    item.AlignNote = AppendPreviewNote(item.AlignNote, "学习库未定位到目标项目里的条目,请手选");
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
                    groupItem.AlignNote = AppendPreviewNote(groupItem.AlignNote, "组件组存在单位、条目或公式风险，整组禁止自动写入");
                }
            }
            items.AddRange(groupItems);
        }

        // 智能铺量预览:漏斗匹配整张目标 sheet。
        private static List<FillPreviewItem> BuildPreview_SmartFill(Form mainForm,
            string targetWorkbook, string targetSheet, string targetColumn, out string warning)
        {
            warning = null;
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
            Dictionary<string, long> projectEntries = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> projectEntryNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (SqlConnection conn = AgentCreateWorkConnection(mainForm))
                {
                    method = SmartResolveProjectMethod(conn);
                    projectEntries = LoadSmartProjectEntries(conn, out projectEntryNames);
                }
            }
            catch (Exception ex)
            {
                Log("Smart fill project context failed: " + ex.Message);
            }

            string snapshotNote;
            SmartLearningSnapshot snapshot = LoadSmartLearningSnapshot(method, out snapshotNote);
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

                List<SmartMapEntry> hits;
                if (snapshot.BySignature.TryGetValue(nameLevelSig, out hits) && hits.Count > 0)
                {
                    AppendRankedSmartMatch(items, row, targetRows, hits, snapshot, projectEntries, currentQuotaByCode,
                        "名称学习命中", nameLevelSig, preferredPrefixes, prefixVotes);
                    hitExact++;
                    continue;
                }
                if (snapshot.ByNameOnly.TryGetValue(nameSig, out hits) && hits.Count > 0)
                {
                    AppendRankedSmartMatch(items, row, targetRows, hits, snapshot, projectEntries, currentQuotaByCode,
                        "名称兼容命中", nameLevelSig, preferredPrefixes, prefixVotes);
                    hitNameOnly++;
                    continue;
                }

                // 模糊层:取相似度最高的至多 3 个名称段,只作候选供人工确认。
                MatchTextFeatures rowFeatures = BuildMatchTextFeatures(nameSig);
                List<KeyValuePair<int, string>> scored = new List<KeyValuePair<int, string>>();
                foreach (KeyValuePair<string, MatchTextFeatures> pair in snapshot.NameFeatures)
                {
                    int score = MatchNameScore(rowFeatures, pair.Value);
                    if (score > 0) scored.Add(new KeyValuePair<int, string>(score, pair.Key));
                }
                scored.Sort(delegate(KeyValuePair<int, string> a, KeyValuePair<int, string> b) { return b.Key.CompareTo(a.Key); });

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
                    Status = "未匹配"
                };
                if (scored.Count > 0)
                {
                    manual.NameQuotaCandidates = new List<NameQuotaCandidateGroup>();
                    foreach (KeyValuePair<int, string> cand in scored.Take(3))
                    {
                        List<SmartMapEntry> candHits;
                        if (!snapshot.ByNameOnly.TryGetValue(cand.Value, out candHits) || candHits.Count == 0) continue;
                        List<SmartMapCandidateScore> candidateScores = RankSmartMapEntries(snapshot, candHits, nameLevelSig,
                            preferredPrefixes, projectEntries, currentQuotaByCode);
                        foreach (SmartMapCandidateScore candidateScore in candidateScores.Take(3))
                        {
                            NameQuotaCandidateGroup group = new NameQuotaCandidateGroup
                            {
                                Key = cand.Value + "|" + candidateScore.Entry.BoxId,
                                Label = "≈" + cand.Value + "(" + cand.Key.ToString(CultureInfo.InvariantCulture) + "分)：" +
                                    BuildSmartCandidateLabel(candidateScore)
                            };
                            List<FillPreviewItem> groupItems = new List<FillPreviewItem>();
                            AppendSmartItems(groupItems, row, targetRows, candidateScore.Entry, snapshot, projectEntries, currentQuotaByCode,
                                true, "模糊候选:" + group.Label,
                                nameLevelSig, preferredPrefixes, null);
                            group.Items = groupItems;
                            if (groupItems.Count > 0) manual.NameQuotaCandidates.Add(group);
                        }
                    }
                    if (manual.NameQuotaCandidates.Count > 0)
                    {
                        manual.AlignNote = "有 " + manual.NameQuotaCandidates.Count.ToString(CultureInfo.InvariantCulture) + " 个模糊候选,双击选择";
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

            string sourceLabel = snapshot.FromSql ? "学习库(SQL)" : "本地映射(jsonl回退)";
            warning = sourceLabel + ":精确 " + hitExact.ToString(CultureInfo.InvariantCulture) +
                " 行,同名 " + hitNameOnly.ToString(CultureInfo.InvariantCulture) +
                " 行,模糊候选 " + fuzzyRows.ToString(CultureInfo.InvariantCulture) +
                " 行,待手挂 " + manualRows.ToString(CultureInfo.InvariantCulture) + " 行。" +
                (preferredPrefixes.Count > 0 ? " 工程前缀:" + String.Join("/", preferredPrefixes.ToArray()) + "。" : "") +
                (snapshotNote != null ? " " + snapshotNote : "");
            return items;
        }
    }
}
