using System;
using System.Collections.Generic;
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
        // 漏斗:①精确签名(名称|单位) ②去单位签名(名称|*) ③模糊候选(仅进下拉,不自动采纳) ④手挂。
        // 数据源:RecoLearning(SQL);连不上时回退 mapping-boxes.jsonl(仅①②,无条目知识)。

        private sealed class SmartBoxTarget
        {
            public string Kind; public string Code; public string Name; public string Unit;
        }

        private sealed class SmartMapEntry
        {
            public string BoxId;
            public int Weight;
            public List<SmartBoxTarget> Targets = new List<SmartBoxTarget>();
        }

        private sealed class SmartEntryStat
        {
            public string EntryCode; public string EntryName; public int ProjectCount;
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
        }

        private static string SmartNameSegment(string signature)
        {
            int idx = (signature ?? "").LastIndexOf('|');
            return idx >= 0 ? signature.Substring(0, idx) : (signature ?? "");
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
                    if (text.IndexOf("2024", StringComparison.Ordinal) >= 0) return "2024";
                    if (text.IndexOf("2020", StringComparison.Ordinal) >= 0) return "2020";
                    if (text.IndexOf("30号文", StringComparison.Ordinal) >= 0) return "2020"; // 30号文=2020办法文号
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
        private static Dictionary<string, long> LoadSmartProjectEntries(SqlConnection projectConn)
        {
            Dictionary<string, long> map = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (SqlCommand cmd = projectConn.CreateCommand())
                {
                    cmd.CommandText = "SELECT 条目编号, 条目序号 FROM 章节表 WHERE 条目编号 IS NOT NULL";
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string code = (reader.IsDBNull(0) ? "" : reader.GetValue(0).ToString()).Trim();
                            if (code.Length == 0 || map.ContainsKey(code)) continue;
                            long seq;
                            if (Int64.TryParse(reader.GetValue(1).ToString(), out seq)) map[code] = seq;
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
            SmartLearningSnapshot snapshot = new SmartLearningSnapshot { Method = method ?? "" };
            try
            {
                string connectionString = GetLearningDbConnectionString();
                if (String.IsNullOrEmpty(connectionString)) throw new InvalidOperationException("学习库连接串为空");
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    Dictionary<string, SmartMapEntry> byKey = new Dictionary<string, SmartMapEntry>(StringComparer.OrdinalIgnoreCase);
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandTimeout = 15;
                        cmd.CommandText =
                            "SELECT m.signature, m.box_id, m.weight, t.target_kind, t.target_code, t.target_name, t.target_unit " +
                            "FROM dbo.SignatureBoxMap m JOIN dbo.QuotaBoxTarget t ON t.box_id = m.box_id " +
                            "WHERE m.weight > 0";
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string signature = reader.GetString(0);
                                string boxId = reader.GetString(1);
                                string key = signature + "\n" + boxId;
                                SmartMapEntry entry;
                                if (!byKey.TryGetValue(key, out entry))
                                {
                                    entry = new SmartMapEntry { BoxId = boxId, Weight = reader.GetInt32(2) };
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
                                entry.Targets.Add(new SmartBoxTarget
                                {
                                    Kind = reader.GetString(3),
                                    Code = reader.GetString(4),
                                    Name = reader.GetString(5),
                                    Unit = reader.GetString(6)
                                });
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
                    foreach (IGrouping<string, Dictionary<string, string>> sample in boxGroup
                        .GroupBy(row => NormalizeForSignature(GetFlat(row, "quantity_name")) + "|" + NormalizeForSignature(GetFlat(row, "quantity_unit")), StringComparer.OrdinalIgnoreCase))
                    {
                        SmartMapEntry entry = new SmartMapEntry { BoxId = boxGroup.Key, Weight = sample.Max(row => ReadFlatInt(row, "weight", 0)) };
                        foreach (Dictionary<string, string> row in sample)
                        {
                            entry.Targets.Add(new SmartBoxTarget
                            {
                                Kind = GetFlat(row, "target_kind"),
                                Code = GetFlat(row, "target_code"),
                                Name = GetFlat(row, "target_name"),
                                Unit = GetFlat(row, "target_unit")
                            });
                        }
                        List<SmartMapEntry> list;
                        if (!snapshot.BySignature.TryGetValue(sample.Key, out list)) { list = new List<SmartMapEntry>(); snapshot.BySignature[sample.Key] = list; }
                        list.Add(entry);
                        string nameSeg = SmartNameSegment(sample.Key);
                        if (!snapshot.ByNameOnly.TryGetValue(nameSeg, out list)) { list = new List<SmartMapEntry>(); snapshot.ByNameOnly[nameSeg] = list; }
                        list.Add(entry);
                    }
                }
                foreach (string nameSeg in snapshot.ByNameOnly.Keys)
                {
                    snapshot.NameFeatures.Add(new KeyValuePair<string, MatchTextFeatures>(nameSeg, BuildMatchTextFeatures(nameSeg)));
                }
                snapshot.FromSql = false;
                note = "学习库(SQL)不可用,已回退本地 mapping-boxes(无条目定位知识,条目请手选)。";
                return snapshot;
            }
            catch (Exception ex)
            {
                Log("Smart fill jsonl fallback failed: " + ex.Message);
                note = "学习库与本地映射均不可用:" + ex.Message;
                return snapshot;
            }
        }

        // 为一组定额目标解析放置条目:取组内各定额在该办法下 project_count 最高、且目标项目存在的条目。
        private static bool TryResolveSmartEntry(SmartLearningSnapshot snapshot, Dictionary<string, long> projectEntries,
            List<SmartBoxTarget> targets, out string entryCode, out long entrySeq)
        {
            entryCode = ""; entrySeq = 0;
            SmartEntryStat best = null;
            foreach (SmartBoxTarget target in targets)
            {
                if (target == null || !String.Equals(target.Kind, "quota", StringComparison.OrdinalIgnoreCase)) continue;
                List<SmartEntryStat> stats;
                if (!snapshot.EntryByQuota.TryGetValue(target.Code ?? "", out stats)) continue;
                foreach (SmartEntryStat stat in stats)
                {
                    if (!projectEntries.ContainsKey(stat.EntryCode)) continue;
                    if (best == null || stat.ProjectCount > best.ProjectCount) best = stat;
                    break; // stats 已按 project_count 降序,该定额取第一个存在的即可
                }
            }
            if (best == null) return false;
            entryCode = best.EntryCode;
            entrySeq = projectEntries[best.EntryCode];
            return true;
        }

        // 由一个映射命中构建预览项(每个定额目标一行,首行承载工程量名)。
        private static void AppendSmartItems(List<FillPreviewItem> items, TargetQtyRow row, SmartMapEntry entry,
            SmartLearningSnapshot snapshot, Dictionary<string, long> projectEntries,
            Dictionary<string, ProjectQuota> projectQuotaByCode, bool needConfirm, string note)
        {
            string entryCode; long entrySeq;
            bool hasEntry = TryResolveSmartEntry(snapshot, projectEntries, entry.Targets, out entryCode, out entrySeq);
            int order = 0;
            foreach (SmartBoxTarget target in entry.Targets)
            {
                if (target == null || String.IsNullOrEmpty(target.Code)) continue;
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
                    SourceName = target.Name,
                    Unit = target.Unit,
                    GroupOrder = order,
                    OrderInItem = row.Row * 10 + order,
                    NeedExactNameConfirmation = needConfirm,
                    AlignNote = note
                };
                item.QuantityText = BuildNameDrivenQtyText(row.QuantityText, row.Unit, target.Unit);

                ProjectQuota inProject;
                if (projectQuotaByCode.TryGetValue(target.Code, out inProject) && !inProject.IsLibrary)
                {
                    item.ChosenQuotaSeq = inProject.QuotaSeq;   // 项目内已有该定额:整行复制,单价随项目
                    if (hasEntry) { item.ChosenItemSeq = entrySeq; item.ChosenItemNo = entryCode; }
                }
                else
                {
                    item.IsLibraryQuota = true;                 // 项目内没有:原生粘贴(软件自算单价)
                    item.ChosenItemNo = hasEntry ? entryCode : "";
                    if (hasEntry) item.ChosenItemSeq = entrySeq;
                }
                item.ItemNo = hasEntry ? entryCode : "";
                if (!hasEntry)
                {
                    item.Status = "缺条目";
                    item.AlignNote = AppendPreviewNote(item.AlignNote, "学习库未定位到目标项目里的条目,请手选");
                }
                items.Add(item);
                order++;
            }
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
            try
            {
                using (SqlConnection conn = AgentCreateWorkConnection(mainForm))
                {
                    method = SmartResolveProjectMethod(conn);
                    projectEntries = LoadSmartProjectEntries(conn);
                }
            }
            catch (Exception ex)
            {
                Log("Smart fill project context failed: " + ex.Message);
            }

            string snapshotNote;
            SmartLearningSnapshot snapshot = LoadSmartLearningSnapshot(method, out snapshotNote);
            if (snapshot.BySignature.Count == 0)
            {
                warning = snapshotNote ?? "学习库为空,请先积累绑定或运行收割。";
                return new List<FillPreviewItem>();
            }

            List<ProjectQuota> projectQuotas = LoadProjectQuotas(mainForm);
            Dictionary<string, ProjectQuota> projectQuotaByCode = projectQuotas
                .GroupBy(q => q.Code, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            List<FillPreviewItem> items = new List<FillPreviewItem>();
            int hitExact = 0, hitNameOnly = 0, fuzzyRows = 0, manualRows = 0;
            foreach (TargetQtyRow row in targetRows)
            {
                string nameSig = NormalizeForSignature(row.RawName);
                string unitSig = NormalizeForSignature(row.Unit);
                string fullSig = nameSig + "|" + unitSig;
                if (fullSig.Length > 450) fullSig = fullSig.Substring(0, 450);

                List<SmartMapEntry> hits;
                if (snapshot.BySignature.TryGetValue(fullSig, out hits) && hits.Count > 0)
                {
                    AppendSmartItems(items, row, hits[0], snapshot, projectEntries, projectQuotaByCode,
                        false, "签名精确命中(权重" + hits[0].Weight.ToString(CultureInfo.InvariantCulture) + ")");
                    hitExact++;
                    continue;
                }
                if (snapshot.ByNameOnly.TryGetValue(nameSig, out hits) && hits.Count > 0)
                {
                    AppendSmartItems(items, row, hits[0], snapshot, projectEntries, projectQuotaByCode,
                        true, "同名不同单位命中,请确认单位换算");
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
                        NameQuotaCandidateGroup group = new NameQuotaCandidateGroup
                        {
                            Key = cand.Value + "|" + candHits[0].BoxId,
                            Label = "≈" + cand.Value + "(" + cand.Key.ToString(CultureInfo.InvariantCulture) + "分)"
                        };
                        List<FillPreviewItem> groupItems = new List<FillPreviewItem>();
                        AppendSmartItems(groupItems, row, candHits[0], snapshot, projectEntries, projectQuotaByCode,
                            true, "模糊候选:" + group.Label);
                        group.Items = groupItems;
                        manual.NameQuotaCandidates.Add(group);
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

            string sourceLabel = snapshot.FromSql ? "学习库(SQL)" : "本地映射(jsonl回退)";
            warning = sourceLabel + ":精确 " + hitExact.ToString(CultureInfo.InvariantCulture) +
                " 行,同名 " + hitNameOnly.ToString(CultureInfo.InvariantCulture) +
                " 行,模糊候选 " + fuzzyRows.ToString(CultureInfo.InvariantCulture) +
                " 行,待手挂 " + manualRows.ToString(CultureInfo.InvariantCulture) + " 行。" +
                (snapshotNote != null ? " " + snapshotNote : "");
            return items;
        }
    }
}
