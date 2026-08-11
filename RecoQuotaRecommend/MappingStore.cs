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
        private readonly string softwarePartition;
        private readonly List<MappingBox> boxes = new List<MappingBox>();
        private readonly List<Dictionary<string, string>> pendingContexts = new List<Dictionary<string, string>>();

        private MappingStore(string filePath, string partition)
        {
            path = filePath;
            softwarePartition = (partition ?? "").Trim();
        }

        public static MappingStore Load(List<LearningRecord> records)
        {
            string filePath = Path.Combine(LearningStore.FindDataDir(), "mapping-boxes.jsonl");
            MappingStore store = new MappingStore(filePath, ResolveCurrentSoftwarePartition());
            store.LoadFile();
            // 扶正训练器已停写库：不再从旧 learning 纠错记录自动播种 mapping-boxes，
            // 定额对应框只由“推荐窗口扶正”和“绑定Excel工程量”两条高信号路径写入。
            return store;
        }

        internal static MappingStore LoadForTesting(string filePath, string partition)
        {
            MappingStore store = new MappingStore(filePath, partition);
            store.LoadFile();
            return store;
        }

        public List<RecommendationRow> Find(ExcelQuantityItem item, string categoryFilter, SearchIndexStore searchIndex, EntryScope scope)
        {
            ScoredBox best = null;
            foreach (MappingBox box in boxes)
            {
                if (!BoxAllowedByScope(box, item, scope))
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
        private static bool BoxAllowedByScope(MappingBox box, ExcelQuantityItem item, EntryScope scope)
        {
            if (scope == null || !scope.Strict)
            {
                return true;
            }

            if (HasCompleteEntryContext(box, item, scope))
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

        private static bool HasCompleteEntryContext(MappingBox box, ExcelQuantityItem item, EntryScope scope)
        {
            if (box == null || item == null || scope == null || String.IsNullOrEmpty(scope.Tag)) return false;
            string partition = (scope.SoftwarePartition ?? "").Trim();
            string methodNo = LearningPartitionIdentity.NormalizeLearningMethodNo(scope.MethodNo);
            string entryCode = LearningPartitionIdentity.NormalizeLearningEntryCode(scope.MatchedEntryCode);
            string quantitySignature = LocalMappingFileStore.BuildQuantitySignature(item.Name);
            List<MappingTarget> quotaTargets = box.Targets.Where(target =>
                String.Equals(String.IsNullOrWhiteSpace(target.TargetKind) ? QuotaEntry.GuessKind(target.Code) : target.TargetKind,
                    "quota", StringComparison.OrdinalIgnoreCase)).ToList();
            if (quotaTargets.Count == 0) return false;
            foreach (MappingTarget target in quotaTargets)
            {
                string targetIdentity = BuildTargetIdentity(target);
                if (!box.Contexts.Any(context =>
                    String.Equals(context.SoftwarePartition, partition, StringComparison.OrdinalIgnoreCase) &&
                    String.Equals(context.MethodNo, methodNo, StringComparison.OrdinalIgnoreCase) &&
                    String.Equals(context.EntryCode, entryCode, StringComparison.OrdinalIgnoreCase) &&
                    String.Equals(context.QuantitySignature, quantitySignature, StringComparison.OrdinalIgnoreCase) &&
                    String.Equals(context.TargetIdentity, targetIdentity, StringComparison.OrdinalIgnoreCase))) return false;
            }
            return true;
        }

        public List<AiMappingCandidate> BuildDeepSeekCandidates(ExcelQuantityItem item, string categoryFilter, SearchIndexStore searchIndex, int limit, EntryScope scope)
        {
            if (item == null)
            {
                return new List<AiMappingCandidate>();
            }

            return boxes
                .Where(box => BoxAllowedByScope(box, item, scope))
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
                QueueContexts(box, sample, scope);
                box.TrimSamples(MaxSamplesPerBox);
                changed = true;
            }

            if (changed)
            {
                Save("MappingStore.Accept");
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
            QueueContexts(box, sample, scope);
            box.TrimSamples(MaxSamplesPerBox);
            Save("MappingStore.Correct");
        }

        private void QueueContexts(MappingBox box, MappingSample sample, EntryScope scope)
        {
            if (box == null || sample == null || scope == null || String.IsNullOrEmpty(scope.Tag)) return;
            string partition = (scope.SoftwarePartition ?? "").Trim();
            string methodNo = LearningPartitionIdentity.NormalizeLearningMethodNo(scope.MethodNo);
            string entryCode = LearningPartitionIdentity.NormalizeLearningEntryCode(scope.MatchedEntryCode);
            if (!String.Equals(partition, softwarePartition, StringComparison.OrdinalIgnoreCase) ||
                String.IsNullOrEmpty(methodNo) || String.IsNullOrEmpty(entryCode)) return;
            foreach (MappingTarget target in box.Targets)
            {
                Dictionary<string, string> row = BuildRelationshipRow(box, target, sample);
                row["record_type"] = "mapping_context";
                row["software_partition"] = softwarePartition;
                row["method_no"] = methodNo;
                row["entry_code"] = entryCode;
                row["entry_name"] = scope.EntryName ?? "";
                string identity = LocalMappingFileStore.BuildMappingContextIdentity(row);
                if (identity.Length == 0) continue;
                pendingContexts.RemoveAll(existing => String.Equals(
                    LocalMappingFileStore.BuildMappingContextIdentity(existing), identity, StringComparison.OrdinalIgnoreCase));
                pendingContexts.Add(row);
            }
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
                box = new MappingBox { BoxId = boxId, SoftwarePartition = softwarePartition };
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
            // 写入方使用同目录原子替换；读取旧文件或新文件都完整，不必占用写锁阻塞窗口打开。
            List<MappingBox> parsed = ParseFile(path, softwarePartition);
            boxes.Clear();
            boxes.AddRange(CanonicalizeBoxes(parsed));
        }

        private static List<MappingBox> ParseFile(string filePath, string partition)
        {
            List<MappingBox> result = new List<MappingBox>();
            if (!File.Exists(filePath) || String.IsNullOrWhiteSpace(partition))
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
                if (!String.Equals(LearningStore.Get(values, "record_type"), "mapping_box", StringComparison.OrdinalIgnoreCase) ||
                    !String.Equals(LearningStore.Get(values, "software_partition").Trim(), partition, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
                string boxId = LearningStore.Get(values, "box_id");
                if (String.IsNullOrWhiteSpace(boxId))
                {
                    continue;
                }

                MappingBox box;
                if (!byId.TryGetValue(boxId, out box))
                {
                    box = new MappingBox { BoxId = boxId, SoftwarePartition = partition };
                    byId[boxId] = box;
                    result.Add(box);
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

            foreach (string line in File.ReadAllLines(filePath, Encoding.UTF8))
            {
                Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                if (!String.Equals(LearningStore.Get(values, "record_type"), "mapping_context", StringComparison.OrdinalIgnoreCase) ||
                    !String.Equals(LearningStore.Get(values, "software_partition").Trim(), partition, StringComparison.OrdinalIgnoreCase)) continue;
                string boxId = LearningStore.Get(values, "box_id").Trim();
                MappingBox box;
                if (!byId.TryGetValue(boxId, out box))
                {
                    QuotaRecommendPanel.Log("MappingStore ignored orphan mapping_context: box_id=" + boxId);
                    continue;
                }
                string targetIdentity = LocalMappingFileStore.BuildTargetIdentity(values);
                string quantitySignature = LocalMappingFileStore.BuildQuantitySignature(LearningStore.Get(values, "quantity_name"));
                bool relationExists = box.Targets.Any(target => String.Equals(BuildTargetIdentity(target), targetIdentity, StringComparison.OrdinalIgnoreCase)) &&
                    box.Samples.Any(sample => String.Equals(LocalMappingFileStore.BuildQuantitySignature(sample.QuantityName), quantitySignature, StringComparison.OrdinalIgnoreCase));
                if (!relationExists)
                {
                    QuotaRecommendPanel.Log("MappingStore ignored orphan mapping_context relation: box_id=" + boxId);
                    continue;
                }
                string methodNo = LearningPartitionIdentity.NormalizeLearningMethodNo(LearningStore.Get(values, "method_no"));
                string entryCode = LearningPartitionIdentity.NormalizeLearningEntryCode(LearningStore.Get(values, "entry_code"));
                if (methodNo.Length == 0 || entryCode.Length == 0) continue;
                box.Contexts.Add(new MappingContextEvidence
                {
                    SoftwarePartition = partition,
                    MethodNo = methodNo,
                    EntryCode = entryCode,
                    EntryName = LearningStore.Get(values, "entry_name"),
                    QuantitySignature = quantitySignature,
                    TargetIdentity = targetIdentity
                });
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
            foreach (MappingContextEvidence context in from.Contexts)
            {
                if (!into.Contexts.Any(existing => existing.Identity == context.Identity)) into.Contexts.Add(context);
            }
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

        private LocalMappingSaveResult Save(string sourceOperation)
        {
            LocalMappingSaveResult result = LocalMappingFileStore.Save(path, softwarePartition, sourceOperation, 5000,
                delegate(List<Dictionary<string, string>> snapshot)
                {
                    LocalMappingMutation mutation = new LocalMappingMutation
                    {
                        SoftwarePartition = softwarePartition,
                        SourceOperation = sourceOperation,
                        TrimSamples = true,
                        MaxSamplesPerBox = MaxSamplesPerBox
                    };
                    foreach (MappingBox box in boxes)
                    {
                        box.TrimSamples(MaxSamplesPerBox);
                        foreach (MappingTarget target in box.Targets)
                        {
                            foreach (MappingSample sample in box.Samples)
                            {
                                mutation.MappingBoxes.Add(BuildRelationshipRow(box, target, sample));
                            }
                        }
                    }
                    foreach (Dictionary<string, string> context in pendingContexts)
                    {
                        mutation.MappingContexts.Add(new Dictionary<string, string>(context, StringComparer.OrdinalIgnoreCase));
                    }
                    return mutation;
                });
            if (result.Succeeded) pendingContexts.Clear();
            ReportLocalMappingSaveResult(result);
            return result;
        }

        private static Dictionary<string, string> BuildRelationshipRow(MappingBox box, MappingTarget target, MappingSample sample)
        {
            Dictionary<string, string> row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            row["record_type"] = "mapping_box";
            row["software_partition"] = box.SoftwarePartition ?? "";
            row["box_id"] = box.BoxId ?? "";
            row["target_kind"] = String.IsNullOrWhiteSpace(target.TargetKind) ? QuotaEntry.GuessKind(target.Code) : target.TargetKind;
            row["target_code"] = target.Code ?? "";
            row["target_name"] = target.Name ?? "";
            row["target_unit"] = target.Unit ?? "";
            row["quantity_name"] = sample.QuantityName ?? "";
            row["quantity_unit"] = sample.QuantityUnit ?? "";
            row["weight"] = sample.Weight.ToString(CultureInfo.InvariantCulture);
            row["accepted_count"] = sample.AcceptedCount.ToString(CultureInfo.InvariantCulture);
            row["corrected_count"] = sample.CorrectedCount.ToString(CultureInfo.InvariantCulture);
            row["rejected_count"] = sample.RejectedCount.ToString(CultureInfo.InvariantCulture);
            row["last_used_at"] = sample.LastUsedAt ?? "";
            return row;
        }

        private static readonly object LocalMappingWarningLock = new object();
        private static readonly HashSet<string> LocalMappingWarningFingerprints = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        private static void ReportLocalMappingSaveResult(LocalMappingSaveResult result)
        {
            if (result == null || result.Succeeded) return;
            QuotaRecommendPanel.Log("MappingStore local save: " + result.Status + "; " + result.ErrorMessage);
            if (result.Status != LocalMappingSaveStatus.DuplicateContextIdentity &&
                result.Status != LocalMappingSaveStatus.AmbiguousBoxUnknownFields) return;
            string fingerprint = (result.FilePath ?? "") + "\n" + (result.FileSha256 ?? "") + "\n" + (result.ConflictIdentity ?? "");
            lock (LocalMappingWarningLock)
            {
                if (!LocalMappingWarningFingerprints.Add(fingerprint)) return;
            }
            string lines = result.LineNumbers.Count == 0 ? "" : String.Join(",", result.LineNumbers.Select(value => value.ToString(CultureInfo.InvariantCulture)).ToArray());
            string message = "\u672c\u5730\u5b66\u4e60\u5df2\u6682\u505c\uff0c\u5f53\u524d\u6587\u4ef6\u672a\u4fee\u6539\u3002\r\n" +
                "\u51b2\u7a81\u8eab\u4efd\u952e\uff1a" + result.ConflictIdentity.Replace("\n", " / ") + "\r\n" +
                "\u547d\u4e2d\u884c\u53f7\uff1a" + lines + "\r\n" +
                "\u6765\u6e90\u64cd\u4f5c\uff1a" + result.SourceOperation + "\r\n" +
                "\u4fee\u590d\u62a5\u544a\uff1a" + (String.IsNullOrWhiteSpace(result.DiagnosticReportPath) ? "(write failed; see log)" : result.DiagnosticReportPath);
            MessageBox.Show(message, "\u672c\u5730\u5b66\u4e60\u51b2\u7a81", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        private static string ResolveCurrentSoftwarePartition()
        {
            string processName = "";
            string moduleFileName = "";
            try
            {
                Process process = Process.GetCurrentProcess();
                processName = process.ProcessName ?? "";
                try { moduleFileName = process.MainModule == null ? "" : process.MainModule.FileName ?? ""; }
                catch { moduleFileName = ""; }
            }
            catch { }
            return LearningPartitionIdentity.ResolveFromProcessIdentity(processName, moduleFileName);
        }

        private static string BuildTargetIdentity(MappingTarget target)
        {
            Dictionary<string, string> row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            row["target_kind"] = target == null ? "" : target.TargetKind ?? "";
            row["target_code"] = target == null ? "" : target.Code ?? "";
            row["target_name"] = target == null ? "" : target.Name ?? "";
            row["target_unit"] = target == null ? "" : target.Unit ?? "";
            return LocalMappingFileStore.BuildTargetIdentity(row);
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
        public string SoftwarePartition;
        public readonly List<MappingTarget> Targets = new List<MappingTarget>();
        public readonly List<MappingSample> Samples = new List<MappingSample>();
        public readonly List<MappingContextEvidence> Contexts = new List<MappingContextEvidence>();

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

    internal sealed class MappingContextEvidence
    {
        public string SoftwarePartition;
        public string MethodNo;
        public string EntryCode;
        public string EntryName;
        public string QuantitySignature;
        public string TargetIdentity;

        public string Identity
        {
            get
            {
                return (SoftwarePartition ?? "") + "\n" + (MethodNo ?? "") + "\n" +
                    (EntryCode ?? "") + "\n" + (QuantitySignature ?? "") + "\n" + (TargetIdentity ?? "");
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
