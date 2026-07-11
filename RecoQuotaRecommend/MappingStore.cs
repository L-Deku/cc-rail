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
    internal sealed class MappingStore
    {
        private const int MaxSamplesPerBox = 30;
        private readonly string path;
        private readonly List<MappingBox> boxes = new List<MappingBox>();

        private MappingStore(string filePath)
        {
            path = filePath;
        }

        public static MappingStore Load(List<LearningRecord> records)
        {
            string filePath = Path.Combine(LearningStore.FindDataDir(), "mapping-boxes.jsonl");
            MappingStore store = new MappingStore(filePath);
            store.LoadFile();
            if (store.boxes.Count == 0)
            {
                store.ImportCorrections(records);
                if (store.boxes.Count > 0)
                {
                    store.Save();
                }
            }
            return store;
        }

        public List<RecommendationRow> Find(ExcelQuantityItem item, string categoryFilter, SearchIndexStore searchIndex, EntryScope scope)
        {
            ScoredBox best = null;
            foreach (MappingBox box in boxes)
            {
                if (!BoxAllowedByScope(box, scope))
                {
                    continue;
                }

                List<MappingTarget> allowedTargets = FilterTargetsByCategory(box.Targets, categoryFilter, searchIndex);
                if (allowedTargets.Count == 0)
                {
                    continue;
                }

                int score = box.Score(item);
                if (score >= 70 && (best == null || score > best.Score))
                {
                    best = new ScoredBox { Box = box, Score = score, AllowedTargets = allowedTargets };
                }
            }

            if (best == null)
            {
                return new List<RecommendationRow>();
            }

            return best.AllowedTargets
                .OrderBy(t => TargetSortRank(t.TargetKind, t.Code))
                .ThenBy(t => t.Code ?? "", StringComparer.OrdinalIgnoreCase)
                .Select(t => t.ToRecommendation(item, best.Score, best.Box.BoxId))
                .ToList();
        }

        private static List<MappingTarget> FilterTargetsByCategory(List<MappingTarget> targets, string categoryFilter, SearchIndexStore searchIndex)
        {
            List<MappingTarget> quotaTargets = targets
                .Where(t => String.Equals(String.IsNullOrWhiteSpace(t.TargetKind) ? QuotaEntry.GuessKind(t.Code) : t.TargetKind, "quota", StringComparison.OrdinalIgnoreCase))
                .ToList();
            List<MappingTarget> materialTargets = targets
                .Where(t => !String.Equals(String.IsNullOrWhiteSpace(t.TargetKind) ? QuotaEntry.GuessKind(t.Code) : t.TargetKind, "quota", StringComparison.OrdinalIgnoreCase))
                .ToList();

            if (quotaTargets.Count == 0)
            {
                return materialTargets;
            }

            List<MappingTarget> allowedQuotaTargets = quotaTargets
                .Where(t => searchIndex != null && searchIndex.IsMappingTargetAllowed(t.TargetKind, t.Code, categoryFilter))
                .ToList();
            if (allowedQuotaTargets.Count == 0)
            {
                return new List<MappingTarget>();
            }

            allowedQuotaTargets.AddRange(materialTargets);
            return allowedQuotaTargets;
        }

        // 严格条目模式下框可用的条件：已带当前条目标签，或全部 quota 类目标都在条目定额池内
        private static bool BoxAllowedByScope(MappingBox box, EntryScope scope)
        {
            if (scope == null || !scope.Strict)
            {
                return true;
            }

            if (box.EntryCodes.Contains(scope.Tag))
            {
                return true;
            }

            bool hasQuotaTarget = false;
            foreach (MappingTarget target in box.Targets)
            {
                string kind = String.IsNullOrWhiteSpace(target.TargetKind) ? QuotaEntry.GuessKind(target.Code) : target.TargetKind;
                if (!String.Equals(kind, "quota", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                hasQuotaTarget = true;
                if (!scope.Allows(kind, target.Code))
                {
                    return false;
                }
            }

            if (hasQuotaTarget)
            {
                return true;
            }

            // 纯材料框按材料码判断
            return box.Targets.Count > 0 && box.Targets.All(t => scope.Allows(String.IsNullOrWhiteSpace(t.TargetKind) ? QuotaEntry.GuessKind(t.Code) : t.TargetKind, t.Code));
        }

        public List<AiMappingCandidate> BuildDeepSeekCandidates(ExcelQuantityItem item, string categoryFilter, SearchIndexStore searchIndex, int limit, EntryScope scope)
        {
            if (item == null)
            {
                return new List<AiMappingCandidate>();
            }

            return boxes
                .Where(box => BoxAllowedByScope(box, scope))
                .Select(box => new
                {
                    Box = box,
                    Targets = FilterTargetsByCategory(box.Targets, categoryFilter, searchIndex),
                    Score = Math.Max(box.Score(item), box.LooseScore(item))
                })
                .Where(x => x.Targets.Count > 0 && x.Score >= 20)
                .OrderByDescending(x => x.Score)
                .Take(Math.Max(1, limit))
                .Select(x => new AiMappingCandidate
                {
                    BoxId = x.Box.BoxId,
                    LocalScore = x.Score,
                    SampleNames = x.Box.SampleNamesForPrompt(),
                    Targets = x.Targets
                })
                .ToList();
        }

        public void Accept(List<RecommendationRow> rows, EntryScope scope)
        {
            bool changed = false;
            foreach (IGrouping<string, RecommendationRow> group in rows
                .Where(r => r != null && r.Item != null && !String.IsNullOrWhiteSpace(r.QuotaCode))
                .GroupBy(r => LearningStore.BuildQuantitySignature(r.Item), StringComparer.OrdinalIgnoreCase))
            {
                RecommendationRow first = group.First();
                MappingBox box = null;
                string boxId = group.Select(r => r.BoxId).FirstOrDefault(id => !String.IsNullOrWhiteSpace(id));
                if (!String.IsNullOrWhiteSpace(boxId))
                {
                    box = boxes.FirstOrDefault(b => String.Equals(b.BoxId, boxId, StringComparison.OrdinalIgnoreCase));
                }

                if (box == null)
                {
                    box = FindOrCreateBox(group.Select(row => new QuotaEntry
                    {
                        TargetKind = String.IsNullOrWhiteSpace(row.TargetKind) ? QuotaEntry.GuessKind(row.QuotaCode) : row.TargetKind,
                        QuotaCode = row.QuotaCode,
                        QuotaName = row.QuotaName,
                        QuotaUnit = row.QuotaUnit
                    }).ToList());
                }

                MappingSample sample = box.FindOrCreateSample(first.Item.Name, first.Item.Unit);
                sample.Weight += 5;
                sample.AcceptedCount += 1;
                sample.LastUsedAt = Now();
                if (scope != null && scope.Strict)
                {
                    box.EntryCodes.Add(scope.Tag);
                }
                box.TrimSamples(MaxSamplesPerBox);
                changed = true;
            }

            if (changed)
            {
                Save();
            }
        }

        public void Correct(ExcelQuantityItem item, RecommendationRow oldRecommendation, List<QuotaEntry> selectedTargets, EntryScope scope)
        {
            if (item == null || selectedTargets == null || selectedTargets.Count == 0)
            {
                return;
            }

            Penalize(item, oldRecommendation, selectedTargets);
            MappingBox box = FindOrCreateBox(selectedTargets);
            MappingSample sample = box.FindOrCreateSample(item.Name, item.Unit);
            sample.Weight += 20;
            sample.CorrectedCount += 1;
            sample.LastUsedAt = Now();
            if (scope != null && scope.Strict)
            {
                box.EntryCodes.Add(scope.Tag);
            }
            box.TrimSamples(MaxSamplesPerBox);
            Save();
        }

        private void Penalize(ExcelQuantityItem item, RecommendationRow oldRecommendation, List<QuotaEntry> selectedTargets)
        {
            if (oldRecommendation == null || String.IsNullOrWhiteSpace(oldRecommendation.QuotaCode))
            {
                return;
            }

            string oldKey = (String.IsNullOrWhiteSpace(oldRecommendation.TargetKind) ? QuotaEntry.GuessKind(oldRecommendation.QuotaCode) : oldRecommendation.TargetKind) + ":" + oldRecommendation.QuotaCode.ToUpperInvariant();
            if (selectedTargets.Any(t => String.Equals(t.TargetKey, oldKey, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            foreach (MappingBox box in boxes)
            {
                if (!box.Targets.Any(t => String.Equals(t.TargetKey, oldKey, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                MappingSample sample = box.FindSample(item.Name, item.Unit);
                if (sample != null)
                {
                    sample.Weight = Math.Max(0, sample.Weight - 10);
                    sample.RejectedCount += 1;
                    sample.LastUsedAt = Now();
                }
            }
        }

        private MappingBox FindOrCreateBox(List<QuotaEntry> targets)
        {
            List<MappingTarget> normalized = targets
                .Where(t => !String.IsNullOrWhiteSpace(t.QuotaCode))
                .Select(t => new MappingTarget
                {
                    TargetKind = String.IsNullOrWhiteSpace(t.TargetKind) ? QuotaEntry.GuessKind(t.QuotaCode) : t.TargetKind,
                    Code = t.QuotaCode.Trim(),
                    Name = t.QuotaName ?? "",
                    Unit = t.QuotaUnit ?? ""
                })
                .GroupBy(t => t.TargetKey, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .OrderBy(t => TargetSortRank(t.TargetKind, t.Code))
                .ThenBy(t => t.Code ?? "", StringComparer.OrdinalIgnoreCase)
                .ToList();

            string boxId = BuildBoxId(normalized);
            MappingBox box = boxes.FirstOrDefault(b => String.Equals(b.BoxId, boxId, StringComparison.OrdinalIgnoreCase));
            if (box == null)
            {
                box = new MappingBox { BoxId = boxId };
                box.Targets.AddRange(normalized);
                boxes.Add(box);
            }
            else
            {
                foreach (MappingTarget target in normalized)
                {
                    if (!box.Targets.Any(t => String.Equals(t.TargetKey, target.TargetKey, StringComparison.OrdinalIgnoreCase)))
                    {
                        box.Targets.Add(target);
                    }
                }
            }

            return box;
        }

        private void ImportCorrections(List<LearningRecord> records)
        {
            foreach (IGrouping<string, LearningRecord> group in records
                .Where(r => r.IsCorrection && !String.IsNullOrWhiteSpace(r.QuotaCode))
                .GroupBy(r => r.QuantitySignature, StringComparer.OrdinalIgnoreCase))
            {
                List<QuotaEntry> targets = group.Select(r => new QuotaEntry
                {
                    TargetKind = QuotaEntry.GuessKind(r.QuotaCode),
                    QuotaCode = r.QuotaCode,
                    QuotaName = r.QuotaName,
                    QuotaUnit = r.QuotaUnit
                }).ToList();
                MappingBox box = FindOrCreateBox(targets);
                LearningRecord first = group.First();
                MappingSample sample = box.FindOrCreateSample(first.QuantityName, first.QuantityUnit);
                sample.Weight = Math.Max(sample.Weight, 30);
                sample.CorrectedCount += 1;
                sample.LastUsedAt = Now();
            }
        }

        private void LoadFile()
        {
            List<MappingBox> parsed = null;
            WithMappingBoxesLock(delegate
            {
                parsed = ParseFile(path);
            });
            boxes.Clear();
            boxes.AddRange(CanonicalizeBoxes(parsed));
        }

        private static List<MappingBox> ParseFile(string filePath)
        {
            List<MappingBox> result = new List<MappingBox>();
            if (!File.Exists(filePath))
            {
                return result;
            }

            Dictionary<string, MappingBox> byId = new Dictionary<string, MappingBox>(StringComparer.OrdinalIgnoreCase);
            foreach (string line in File.ReadAllLines(filePath, Encoding.UTF8))
            {
                if (String.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                string boxId = LearningStore.Get(values, "box_id");
                if (String.IsNullOrWhiteSpace(boxId))
                {
                    continue;
                }

                MappingBox box;
                if (!byId.TryGetValue(boxId, out box))
                {
                    box = new MappingBox { BoxId = boxId };
                    byId[boxId] = box;
                    result.Add(box);
                }

                foreach (string entryTag in LearningStore.Get(values, "entry_codes").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    box.EntryCodes.Add(entryTag.Trim());
                }

                // 类别由编号确定地推导（覆盖旧文件里把 ZLF/TLF 等辅助代号误存成 quota 的记录）
                string parsedCode = LearningStore.Get(values, "target_code");
                MappingTarget target = new MappingTarget
                {
                    TargetKind = QuotaEntry.GuessKind(parsedCode),
                    Code = parsedCode,
                    Name = LearningStore.Get(values, "target_name"),
                    Unit = LearningStore.Get(values, "target_unit")
                };
                if (!String.IsNullOrWhiteSpace(target.Code) && !box.Targets.Any(t => String.Equals(t.TargetKey, target.TargetKey, StringComparison.OrdinalIgnoreCase)))
                {
                    box.Targets.Add(target);
                }

                MappingSample sample = new MappingSample();
                sample.QuantityName = LearningStore.Get(values, "quantity_name");
                sample.QuantityUnit = LearningStore.Get(values, "quantity_unit");
                sample.Weight = ParseInt(LearningStore.Get(values, "weight"), 0);
                sample.AcceptedCount = ParseInt(LearningStore.Get(values, "accepted_count"), 0);
                sample.CorrectedCount = ParseInt(LearningStore.Get(values, "corrected_count"), 0);
                sample.RejectedCount = ParseInt(LearningStore.Get(values, "rejected_count"), 0);
                sample.LastUsedAt = LearningStore.Get(values, "last_used_at");
                if (!String.IsNullOrWhiteSpace(sample.QuantityName) && box.FindSample(sample.QuantityName, sample.QuantityUnit) == null)
                {
                    box.Samples.Add(sample);
                }
            }

            return result;
        }

        // 旧版 box_id 由 String.GetHashCode 生成，跨进程位数不稳定且可能碰撞；
        // 加载时按目标组合重算稳定 ID，同一组合的旧框自动合并，实现旧文件无感迁移。
        private static List<MappingBox> CanonicalizeBoxes(List<MappingBox> parsed)
        {
            List<MappingBox> result = new List<MappingBox>();
            Dictionary<string, MappingBox> byId = new Dictionary<string, MappingBox>(StringComparer.OrdinalIgnoreCase);
            foreach (MappingBox box in parsed ?? new List<MappingBox>())
            {
                if (box.Targets.Count == 0)
                {
                    continue;
                }

                string canonicalId = BuildBoxId(box.Targets);
                MappingBox existing;
                if (!byId.TryGetValue(canonicalId, out existing))
                {
                    box.BoxId = canonicalId;
                    byId[canonicalId] = box;
                    result.Add(box);
                    continue;
                }

                MergeBox(existing, box);
            }

            return result;
        }

        private static void MergeBox(MappingBox into, MappingBox from)
        {
            into.EntryCodes.UnionWith(from.EntryCodes);
            foreach (MappingTarget target in from.Targets)
            {
                if (!into.Targets.Any(t => String.Equals(t.TargetKey, target.TargetKey, StringComparison.OrdinalIgnoreCase)))
                {
                    into.Targets.Add(target);
                }
            }

            foreach (MappingSample sample in from.Samples)
            {
                MappingSample existing = into.FindSample(sample.QuantityName, sample.QuantityUnit);
                if (existing == null)
                {
                    into.Samples.Add(sample);
                }
                else if (String.Compare(sample.LastUsedAt ?? "", existing.LastUsedAt ?? "", StringComparison.Ordinal) > 0)
                {
                    existing.Weight = sample.Weight;
                    existing.AcceptedCount = sample.AcceptedCount;
                    existing.CorrectedCount = sample.CorrectedCount;
                    existing.RejectedCount = sample.RejectedCount;
                    existing.LastUsedAt = sample.LastUsedAt;
                }
            }
        }

        private void Save()
        {
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            WithMappingBoxesLock(delegate
            {
                MergeFromDisk();
                WriteFile();
            });
        }

        // Excel联动AI匹配和扶正训练器也会写这个文件；整文件重写前先合并磁盘上的新增记录，避免覆盖丢失。
        private void MergeFromDisk()
        {
            foreach (MappingBox diskBox in CanonicalizeBoxes(ParseFile(path)))
            {
                MappingBox memory = boxes.FirstOrDefault(b => String.Equals(b.BoxId, diskBox.BoxId, StringComparison.OrdinalIgnoreCase));
                if (memory == null)
                {
                    boxes.Add(diskBox);
                }
                else
                {
                    MergeBox(memory, diskBox);
                }
            }
        }

        private void WriteFile()
        {
            string temp = path + ".tmp";
            using (StreamWriter writer = new StreamWriter(temp, false, Encoding.UTF8))
            {
                foreach (MappingBox box in boxes)
                {
                    box.TrimSamples(MaxSamplesPerBox);
                    foreach (MappingTarget target in box.Targets
                        .OrderBy(t => TargetSortRank(t.TargetKind, t.Code))
                        .ThenBy(t => t.Code ?? "", StringComparer.OrdinalIgnoreCase))
                    {
                        foreach (MappingSample sample in box.Samples)
                        {
                            Dictionary<string, string> row = new Dictionary<string, string>();
                            row["record_type"] = "mapping_box";
                            row["box_id"] = box.BoxId;
                            row["target_kind"] = String.IsNullOrWhiteSpace(target.TargetKind) ? QuotaEntry.GuessKind(target.Code) : target.TargetKind;
                            row["target_code"] = target.Code;
                            row["target_name"] = target.Name;
                            row["target_unit"] = target.Unit;
                            row["quantity_name"] = sample.QuantityName;
                            row["quantity_unit"] = sample.QuantityUnit;
                            row["weight"] = sample.Weight.ToString(CultureInfo.InvariantCulture);
                            row["accepted_count"] = sample.AcceptedCount.ToString(CultureInfo.InvariantCulture);
                            row["corrected_count"] = sample.CorrectedCount.ToString(CultureInfo.InvariantCulture);
                            row["rejected_count"] = sample.RejectedCount.ToString(CultureInfo.InvariantCulture);
                            row["last_used_at"] = sample.LastUsedAt;
                            if (box.EntryCodes.Count > 0)
                            {
                                row["entry_codes"] = String.Join(",", box.EntryCodes.OrderBy(t => t, StringComparer.OrdinalIgnoreCase).ToArray());
                            }
                            writer.WriteLine(LearningStore.ToJson(row));
                        }
                    }
                }
            }

            if (File.Exists(path))
            {
                File.Delete(path);
            }
            File.Move(temp, path);
        }

        private const string MappingBoxesMutexName = "RecoQuotaData.mapping-boxes.lock";

        // 学习库有多个写入方（本窗口扶正、RecoExpandPanel 的 Excel联动与训练器），
        // 用跨程序集一致的命名互斥锁串行化读改写，名称必须与 RecoExpandPanel 保持一致。
        private static void WithMappingBoxesLock(Action action)
        {
            Mutex mutex = new Mutex(false, MappingBoxesMutexName);
            bool acquired = false;
            try
            {
                try
                {
                    acquired = mutex.WaitOne(5000);
                }
                catch (AbandonedMutexException)
                {
                    acquired = true;
                }

                action();
            }
            finally
            {
                if (acquired)
                {
                    mutex.ReleaseMutex();
                }
                mutex.Dispose();
            }
        }

        // 与 RecoExpandPanel 的 BuildStableMappingBoxId 使用同一套规则：对小写化目标键做 SHA1。
        private static string BuildBoxId(List<MappingTarget> targets)
        {
            string raw = String.Join("|", targets
                .OrderBy(t => t.TargetKey, StringComparer.OrdinalIgnoreCase)
                .Select(t => t.TargetKey)
                .ToArray());
            using (System.Security.Cryptography.SHA1 sha = System.Security.Cryptography.SHA1.Create())
            {
                byte[] hash = sha.ComputeHash(Encoding.UTF8.GetBytes(raw.ToLowerInvariant()));
                StringBuilder builder = new StringBuilder("box-");
                for (int i = 0; i < 8; i++)
                {
                    builder.Append(hash[i].ToString("x2", CultureInfo.InvariantCulture));
                }

                return builder.ToString();
            }
        }

        public static int TargetSortRank(string targetKind, string code)
        {
            string kind = String.IsNullOrWhiteSpace(targetKind) ? QuotaEntry.GuessKind(code) : targetKind;
            return String.Equals(kind, "quota", StringComparison.OrdinalIgnoreCase) ? 0 : 1;
        }

        private static int ParseInt(string text, int fallback)
        {
            int value;
            return Int32.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value) ? value : fallback;
        }

        private static string Now()
        {
            return DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
        }

        private sealed class ScoredBox
        {
            public MappingBox Box;
            public int Score;
            public List<MappingTarget> AllowedTargets;
        }
    }

    internal sealed class MappingBox
    {
        public string BoxId;
        public readonly List<MappingTarget> Targets = new List<MappingTarget>();
        public readonly List<MappingSample> Samples = new List<MappingSample>();
        // 章节条目标签（"2020:0101-01" 形式，method:条目编号），用于按条目分类对应框
        public readonly HashSet<string> EntryCodes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        public int Score(ExcelQuantityItem item)
        {
            int best = 0;
            foreach (MappingSample sample in Samples)
            {
                if (!CanUseSampleForItem(item, sample))
                {
                    continue;
                }

                if (!TextMatcher.HasStrongNamePairMatch(item.Name, sample.QuantityName))
                {
                    continue;
                }

                if (TextMatcher.IsSteelOnlyAgainstConcrete(item.Name, sample.QuantityName))
                {
                    continue;
                }

                int similarity = TextMatcher.NamePairScore(item.Name, sample.QuantityName);
                if (similarity <= 0)
                {
                    continue;
                }

                int score = similarity + Math.Min(40, sample.Weight);
                if (RecommendDialog.UnitCompatibleForIndex(sample.QuantityUnit, item.Unit))
                {
                    score += 12;
                }
                best = Math.Max(best, score);
            }

            return best;
        }

        public int LooseScore(ExcelQuantityItem item)
        {
            int best = 0;
            string raw = item == null ? "" : item.RawRowText;
            foreach (MappingSample sample in Samples)
            {
                if (!CanUseSampleForItem(item, sample))
                {
                    continue;
                }

                int score = Math.Max(
                    TextMatcher.NamePairScore(item == null ? "" : item.Name, sample.QuantityName),
                    TextMatcher.NamePairScore(raw, sample.QuantityName) / 2);
                if (RecommendDialog.UnitCompatibleForIndex(sample.QuantityUnit, item == null ? "" : item.Unit))
                {
                    score += 8;
                }
                best = Math.Max(best, score + Math.Min(20, sample.Weight / 2));
            }

            return best;
        }

        private bool CanUseSampleForItem(ExcelQuantityItem item, MappingSample sample)
        {
            if (item == null || sample == null)
            {
                return false;
            }

            if (!String.IsNullOrWhiteSpace(item.Unit) &&
                !String.IsNullOrWhiteSpace(sample.QuantityUnit) &&
                !RecommendDialog.UnitCompatibleForIndex(sample.QuantityUnit, item.Unit))
            {
                return false;
            }

            string targetText = String.Join(" ", Targets.Select(t => (t == null ? "" : (t.Code ?? "") + " " + (t.Name ?? "") + " " + (t.Unit ?? ""))).ToArray());
            return !HasEngineeringProcessConflict(item.Name + " " + item.RawRowText, sample.QuantityName + " " + targetText);
        }

        private static bool HasEngineeringProcessConflict(string quantityText, string candidateText)
        {
            string q = TextMatcher.Normalize(quantityText);
            string c = TextMatcher.Normalize(candidateText);
            bool qSheetPileRemoval = ContainsAny(q, "\u62c9\u68ee", "\u94a2\u677f\u6869") && ContainsAny(q, "\u62d4\u9664", "\u62c6\u9664", "\u62d4\u51fa");
            bool qHasGrout = ContainsAny(q, "\u6ce8\u6d46", "\u6c34\u6ce5\u6d46", "\u586b\u5145", "\u5145\u586b", "\u56de\u586b");
            bool cHasGrout = ContainsAny(c, "\u6ce8\u6d46", "\u6c34\u6ce5\u6d46", "\u586b\u5145", "\u5145\u586b");
            if (qSheetPileRemoval && !qHasGrout && cHasGrout)
            {
                return true;
            }

            bool qConcreteOnly = TextMatcher.IsConcreteQuantityName(q) && !TextMatcher.IsSteelQuantityName(q);
            bool cSteelOnly = TextMatcher.IsSteelQuantityName(c) && !TextMatcher.IsConcreteQuantityName(c);
            if (qConcreteOnly && cSteelOnly)
            {
                return true;
            }

            return false;
        }

        private static bool ContainsAny(string text, params string[] keywords)
        {
            foreach (string keyword in keywords ?? new string[0])
            {
                if (!String.IsNullOrWhiteSpace(keyword) && (text ?? "").Contains(TextMatcher.Normalize(keyword)))
                {
                    return true;
                }
            }
            return false;
        }

        public string SampleNamesForPrompt()
        {
            return String.Join("；", Samples
                .OrderByDescending(s => s.Weight)
                .ThenByDescending(s => s.LastUsedAt ?? "")
                .Take(8)
                .Select(s => s.QuantityName)
                .Where(n => !String.IsNullOrWhiteSpace(n))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray());
        }

        public MappingSample FindOrCreateSample(string name, string unit)
        {
            MappingSample sample = FindSample(name, unit);
            if (sample != null)
            {
                return sample;
            }

            sample = new MappingSample { QuantityName = name ?? "", QuantityUnit = unit ?? "", Weight = 10, LastUsedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) };
            Samples.Add(sample);
            return sample;
        }

        public MappingSample FindSample(string name, string unit)
        {
            string signature = LearningStore.BuildQuantitySignature(name, unit);
            return Samples.FirstOrDefault(s => String.Equals(LearningStore.BuildQuantitySignature(s.QuantityName, s.QuantityUnit), signature, StringComparison.OrdinalIgnoreCase));
        }

        public void TrimSamples(int maxSamples)
        {
            while (Samples.Count > maxSamples)
            {
                MappingSample remove = Samples
                    .OrderBy(s => s.Weight)
                    .ThenBy(s => s.LastUsedAt ?? "")
                    .First();
                Samples.Remove(remove);
            }
        }
    }

    internal sealed class MappingTarget
    {
        public string TargetKind;
        public string Code;
        public string Name;
        public string Unit;

        public string TargetKey
        {
            get { return (String.IsNullOrWhiteSpace(TargetKind) ? QuotaEntry.GuessKind(Code) : TargetKind) + ":" + (Code ?? "").Trim().ToUpperInvariant(); }
        }

        public RecommendationRow ToRecommendation(ExcelQuantityItem item, int score, string boxId)
        {
            RecommendationRow row = new RecommendationRow();
            row.Item = item;
            row.QuotaCode = Code;
            row.QuotaName = Name;
            row.QuotaUnit = Unit;
            row.ConvertedValueText = RecommendDialog.ConvertQuantityForIndex(item.ValueText, item.Unit, Unit);
            row.Score = score;
            row.Reason = "\u5b9a\u989d\u5bf9\u5e94\u6846\u6743\u91cd\u5339\u914d";
            row.Source = "mapping";
            row.BoxId = boxId;
            row.TargetKind = String.IsNullOrWhiteSpace(TargetKind) ? QuotaEntry.GuessKind(Code) : TargetKind;
            return row;
        }
    }

    internal sealed class MappingSample
    {
        public string QuantityName;
        public string QuantityUnit;
        public int Weight;
        public int AcceptedCount;
        public int CorrectedCount;
        public int RejectedCount;
        public string LastUsedAt;
    }
}
