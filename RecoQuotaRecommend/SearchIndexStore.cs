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
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace RecoQuotaRecommend
{
    internal sealed class SearchIndexStore
    {
        private static readonly object CacheLock = new object();
        private static readonly Dictionary<string, SearchIndexCacheEntry> StoreCache = new Dictionary<string, SearchIndexCacheEntry>(StringComparer.OrdinalIgnoreCase);
        private readonly List<IndexQuota> quotas = new List<IndexQuota>();
        private readonly object materialSnapshotLock = new object();
        private readonly object materialRefreshLock = new object();
        private List<IndexMaterial> materials = new List<IndexMaterial>();
        private readonly Dictionary<string, List<IndexQuota>> quotaTokenIndex = new Dictionary<string, List<IndexQuota>>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, IndexQuota> quotasByCode = new Dictionary<string, IndexQuota>(StringComparer.OrdinalIgnoreCase);
        private string cacheKey;
        private string quotaIndexPath;
        private string materialIndexPath;
        private string materialDatabaseName;
        private Task materialRefreshTask;

        public int QuotaCount { get { return quotas.Count; } }
        public int MaterialCount
        {
            get
            {
                lock (materialSnapshotLock) return materials.Count;
            }
        }

        private sealed class SearchIndexCacheEntry
        {
            public string Fingerprint;
            public SearchIndexStore Store;
        }

        public static SearchIndexStore LoadOrBuild()
        {
            string dataDir = LearningStore.FindDataDir();
            string materialDatabaseName = ResolveDatabaseName();
            string quotaPath = ResolveQuotaIndexPath(dataDir, materialDatabaseName);
            string legacyQuotaPath = Path.Combine(dataDir, "quota-index.jsonl");
            string materialPath = ResolveMaterialIndexPath(dataDir, materialDatabaseName);
            string legacyMaterialPath = Path.Combine(dataDir, "material-index.jsonl");

            string quotaLoadPath = "";
            if (File.Exists(quotaPath) && QuotaFileMatchesDatabase(quotaPath, materialDatabaseName))
            {
                quotaLoadPath = quotaPath;
            }
            else if (File.Exists(legacyQuotaPath) && QuotaFileMatchesDatabase(legacyQuotaPath, materialDatabaseName))
            {
                quotaLoadPath = legacyQuotaPath;
            }

            if (String.IsNullOrWhiteSpace(quotaLoadPath))
            {
                try
                {
                    ExportQuotaFromSql(dataDir, quotaPath, materialDatabaseName);
                    quotaLoadPath = quotaPath;
                }
                catch (Exception ex)
                {
                    QuotaRecommendPanel.Log("Build quota search index failed: " + ex.Message);
                }
            }

            string materialLoadPath = "";
            if (File.Exists(materialPath) && MaterialFileMatchesDatabase(materialPath, materialDatabaseName))
            {
                materialLoadPath = materialPath;
            }
            else if (File.Exists(legacyMaterialPath) && MaterialFileMatchesDatabase(legacyMaterialPath, materialDatabaseName))
            {
                materialLoadPath = legacyMaterialPath;
            }

            string cacheKey = Path.GetFullPath(dataDir).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + "|" + materialDatabaseName;
            lock (CacheLock)
            {
                string fingerprint = BuildFileFingerprint(quotaLoadPath) + "|" + BuildFileFingerprint(materialLoadPath);
                SearchIndexCacheEntry cached;
                if (StoreCache.TryGetValue(cacheKey, out cached) &&
                    String.Equals(cached.Fingerprint, fingerprint, StringComparison.Ordinal))
                {
                    return cached.Store;
                }

                SearchIndexStore store = new SearchIndexStore();
                store.cacheKey = cacheKey;
                store.quotaIndexPath = quotaLoadPath;
                store.materialIndexPath = materialPath;
                store.materialDatabaseName = materialDatabaseName;
                store.LoadFiles(quotaLoadPath, materialLoadPath);
                StoreCache[cacheKey] = new SearchIndexCacheEntry
                {
                    Fingerprint = BuildFileFingerprint(quotaLoadPath) + "|" + BuildFileFingerprint(materialLoadPath),
                    Store = store
                };
                return store;
            }
        }

        private static string BuildFileFingerprint(string path)
        {
            if (!File.Exists(path))
            {
                return "missing";
            }

            FileInfo info = new FileInfo(path);
            return info.Length.ToString(CultureInfo.InvariantCulture) + ":" + info.LastWriteTimeUtc.Ticks.ToString(CultureInfo.InvariantCulture);
        }

        internal static string ResolveMaterialIndexPath(string dataDir, string databaseName)
        {
            string suffix = String.Equals(databaseName, "RecoData2024", StringComparison.OrdinalIgnoreCase) ? "2024" : "2020";
            return Path.Combine(dataDir ?? "", "material-index-" + suffix + ".jsonl");
        }

        internal static string ResolveQuotaIndexPath(string dataDir, string databaseName)
        {
            string suffix = String.Equals(databaseName, "RecoData2024", StringComparison.OrdinalIgnoreCase) ? "2024" : "2020";
            return Path.Combine(dataDir ?? "", "quota-index-" + suffix + ".jsonl");
        }

        private static bool QuotaFileMatchesDatabase(string path, string databaseName)
        {
            try
            {
                foreach (string line in File.ReadLines(path, Encoding.UTF8))
                {
                    if (String.IsNullOrWhiteSpace(line)) continue;
                    Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                    string sourceDatabase = LearningStore.Get(values, "source_database");
                    if (!String.IsNullOrWhiteSpace(sourceDatabase))
                    {
                        return String.Equals(sourceDatabase, databaseName, StringComparison.OrdinalIgnoreCase);
                    }

                    string bookCode = LearningStore.Get(values, "book_code");
                    string specialty = LearningStore.Get(values, "specialty");
                    bool is2024 = (bookCode ?? "").IndexOf("2024", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        (specialty ?? "").IndexOf("2024", StringComparison.OrdinalIgnoreCase) >= 0;
                    return String.Equals(databaseName, "RecoData2024", StringComparison.OrdinalIgnoreCase) ? is2024 : !is2024;
                }
            }
            catch
            {
            }

            return false;
        }

        private static bool MaterialFileMatchesDatabase(string path, string databaseName)
        {
            try
            {
                foreach (string line in File.ReadLines(path, Encoding.UTF8))
                {
                    if (String.IsNullOrWhiteSpace(line)) continue;
                    Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                    string sourceDatabase = LearningStore.Get(values, "source_database");
                    if (!String.IsNullOrWhiteSpace(sourceDatabase))
                    {
                        return String.Equals(sourceDatabase, databaseName, StringComparison.OrdinalIgnoreCase);
                    }

                    string documentNo = LearningStore.Get(values, "doc_no");
                    if (String.IsNullOrWhiteSpace(documentNo) || String.Equals(documentNo.Trim(), "\u8865\u5145\u6750\u6599", StringComparison.Ordinal))
                    {
                        continue;
                    }

                    bool is2024 = documentNo.IndexOf("2024", StringComparison.OrdinalIgnoreCase) >= 0;
                    return String.Equals(databaseName, "RecoData2024", StringComparison.OrdinalIgnoreCase) ? is2024 : !is2024;
                }
            }
            catch
            {
            }

            return false;
        }

        public List<RecommendationRow> SearchQuotaCandidates(ExcelQuantityItem item, string categoryFilter, EntryScope scope, int limit)
        {
            return SearchQuotaCandidatesCore(item, categoryFilter, scope, Math.Max(1, limit));
        }

        public List<RecommendationRow> SearchAllQuotaCandidates(ExcelQuantityItem item, string categoryFilter, EntryScope scope)
        {
            return SearchQuotaCandidatesCore(item, categoryFilter, scope, null);
        }

        private List<RecommendationRow> SearchQuotaCandidatesCore(ExcelQuantityItem item, string categoryFilter, EntryScope scope, int? limit)
        {
            if (item == null)
            {
                return new List<RecommendationRow>();
            }

            string chinesePhrase = ExtractChinesePhrase(item.Name);
            string majorChapter = GetMajorChapterCode(scope);
            List<ScoredQuota> scopedHits = new List<ScoredQuota>();
            if (scope != null && scope.Strict)
            {
                foreach (string code in scope.QuotaPoolCodes)
                {
                    IndexQuota quota;
                    if (!quotasByCode.TryGetValue((code ?? "").Trim(), out quota) ||
                        !CategoryAllowed(quota.BookCategory, categoryFilter))
                    {
                        continue;
                    }

                    if (!QuotaMatchesRequiredPhrase(quota, chinesePhrase))
                    {
                        continue;
                    }

                    int score = ScoreQuota(item, quota);
                    if (score > 0)
                    {
                        scopedHits.Add(CreateScoredQuota(quota, score, scope, majorChapter));
                    }
                }
            }

            IEnumerable<ScoredQuota> allHits = GetQuotaCandidates(item, categoryFilter, null)
                .Select(q => new ScoredQuota { Quota = q, Score = ScoreQuota(item, q) })
                .Where(q => q.Score > 0)
                .Where(q => QuotaMatchesRequiredPhrase(q.Quota, chinesePhrase))
                .Select(q => CreateScoredQuota(q.Quota, q.Score, scope, majorChapter))
                .Concat(scopedHits);

            IEnumerable<ScoredQuota> ordered = allHits
                .GroupBy(q => q.Quota.QuotaCode ?? "", StringComparer.OrdinalIgnoreCase)
                .Select(g => g.OrderBy(x => x.MajorRank).ThenBy(x => x.PoolRank).ThenByDescending(x => x.Score).ThenBy(x => x.Quota.SortOrder).First())
                .OrderBy(q => q.MajorRank)
                .ThenBy(q => q.PoolRank)
                .ThenByDescending(q => q.Score)
                .ThenBy(q => q.Quota.SortOrder);

            if (limit.HasValue)
            {
                ordered = ordered.Take(limit.Value);
            }

            return ordered
                .Select(q => q.Quota.ToRecommendation(item, q.Score))
                .ToList();
        }

        public List<RecommendationRow> SearchMaterialCandidates(ExcelQuantityItem item)
        {
            if (item == null)
            {
                return new List<RecommendationRow>();
            }

            string query = TextMatcher.Normalize(item.Name).Replace(" ", "");
            if (String.IsNullOrWhiteSpace(query))
            {
                return new List<RecommendationRow>();
            }

            List<IndexMaterial> snapshot;
            lock (materialSnapshotLock) snapshot = materials;

            return snapshot
                .Select(material => new ScoredMaterial
                {
                    Material = material,
                    NormalizedName = TextMatcher.Normalize(material.MaterialName).Replace(" ", "")
                })
                .Where(candidate => candidate.NormalizedName.Contains(query))
                .Select(candidate =>
                {
                    candidate.MatchIndex = candidate.NormalizedName.IndexOf(query, StringComparison.Ordinal);
                    candidate.MatchRank = String.Equals(candidate.NormalizedName, query, StringComparison.Ordinal) ? 0 :
                        (candidate.MatchIndex == 0 ? 1 : 2);
                    return candidate;
                })
                .OrderBy(candidate => candidate.MatchRank)
                .ThenBy(candidate => candidate.MatchIndex)
                .ThenBy(candidate => candidate.NormalizedName.Length)
                .ThenBy(candidate => candidate.Material.MaterialCode ?? "", StringComparer.OrdinalIgnoreCase)
                .Select(candidate => candidate.Material.ToRecommendation(item, 300 - candidate.MatchRank * 50 - candidate.MatchIndex))
                .ToList();
        }

        private static ScoredQuota CreateScoredQuota(IndexQuota quota, int score, EntryScope scope, string majorChapter)
        {
            ScoredQuota result = new ScoredQuota();
            result.Quota = quota;
            result.Score = score;
            result.PoolRank = scope != null && scope.Strict && scope.Allows("quota", quota == null ? null : quota.QuotaCode) ? 0 : 1;
            result.MajorRank = SpecialtyMatchesMajorChapter(majorChapter, quota == null ? null : quota.Specialty) ? 0 : 1;
            return result;
        }

        private static bool QuotaMatchesRequiredPhrase(IndexQuota quota, string chinesePhrase)
        {
            if (String.IsNullOrWhiteSpace(chinesePhrase))
            {
                return true;
            }

            if (quota == null)
            {
                return false;
            }

            string name = ExtractChinesePhrase(quota.QuotaName);
            string work = ExtractChinesePhrase(quota.WorkContent);
            return name.Contains(chinesePhrase) || work.Contains(chinesePhrase);
        }

        private static string ExtractChinesePhrase(string text)
        {
            StringBuilder builder = new StringBuilder();
            foreach (char ch in text ?? "")
            {
                if (ch >= 0x4e00 && ch <= 0x9fff)
                {
                    builder.Append(ch);
                }
            }
            return builder.ToString();
        }

        private static string GetMajorChapterCode(EntryScope scope)
        {
            if (scope == null)
            {
                return "";
            }

            string code = !String.IsNullOrWhiteSpace(scope.ProjectEntryCode) ? scope.ProjectEntryCode : scope.MatchedEntryCode;
            if (String.IsNullOrWhiteSpace(code))
            {
                return "";
            }

            Match match = Regex.Match(code.Trim(), "\\d+");
            if (!match.Success)
            {
                return "";
            }

            string digits = match.Value;
            if (digits.Length == 1)
            {
                return "0" + digits;
            }
            return digits.Substring(0, 2);
        }

        private static bool SpecialtyMatchesMajorChapter(string majorChapter, string specialty)
        {
            string text = TextMatcher.Normalize(specialty);
            if (String.IsNullOrWhiteSpace(majorChapter) || String.IsNullOrWhiteSpace(text))
            {
                return false;
            }

            switch (majorChapter)
            {
                case "02":
                    return text.Contains("\u8def\u57fa");
                case "03":
                    return text.Contains("\u6865\u6db5");
                case "04":
                    return text.Contains("\u96a7\u9053");
                case "05":
                    return text.Contains("\u8f68\u9053");
                case "06":
                    return text.Contains("\u901a\u4fe1") ||
                        text.Contains("\u4fe1\u53f7") ||
                        text.Contains("\u4fe1\u606f") ||
                        text.Contains("\u707e\u5bb3\u76d1\u6d4b");
                case "07":
                    return text.Contains("\u7535\u529b") ||
                        text.Contains("\u7535\u529b\u7275\u5f15\u4f9b\u7535");
                case "08":
                    return text.Contains("\u623f\u5c4b");
                case "09":
                    return text.Contains("\u7ad9\u573a") ||
                        text.Contains("\u7ed9\u6392\u6c34") ||
                        text.Contains("\u673a\u52a1") ||
                        text.Contains("\u8f66\u8f86") ||
                        text.Contains("\u673a\u68b0") ||
                        text.Contains("\u8fd0\u8425\u751f\u4ea7\u8bbe\u5907") ||
                        text.Contains("\u5efa\u7b51\u7269");
                default:
                    return false;
            }
        }
        public bool IsMappingTargetAllowed(string targetKind, string code, string categoryFilter)
        {
            string kind = String.IsNullOrWhiteSpace(targetKind) ? QuotaEntry.GuessKind(code) : targetKind;
            if (!QuotaEntry.IsQuotaKind(kind))
            {
                return true; // 材料与辅助代号随定额一起带出，不受定额类别过滤
            }

            // 扶正时可能带 *乘数（如 LY-25*9），查索引前先归一化取原始编号
            IndexQuota quota;
            if (!quotasByCode.TryGetValue(QuotaEntry.NormalizeCode(code), out quota))
            {
                return false;
            }

            return CategoryAllowed(quota.BookCategory, categoryFilter);
        }

        public List<AiQuotaCandidate> BuildDeepSeekCandidates(ExcelQuantityItem item, string categoryFilter, int limit, EntryScope scope)
        {
            if (item == null)
            {
                return new List<AiQuotaCandidate>();
            }

            int max = Math.Max(1, limit);
            return GetQuotaCandidates(item, categoryFilter, scope)
                .Select(q => new AiQuotaCandidate { Quota = q, LocalScore = ScoreQuota(item, q) })
                .Where(c => c.LocalScore > 0)
                .OrderByDescending(c => c.LocalScore)
                .ThenBy(c => c.Quota.SortOrder)
                .Take(max)
                .ToList();
        }

        // 严格条目模式：把整条目定额池作为候选（不依赖关键词命中），按名称相似度打分。
        // 这样池里相关定额即使名称不完全匹配也能进候选，既提高本地直接命中率，也让 AI 拿到聚焦的小候选集，更快更准。
        public List<AiQuotaCandidate> BuildScopeCandidates(ExcelQuantityItem item, EntryScope scope, int limit)
        {
            if (item == null || scope == null || !scope.Strict)
            {
                return new List<AiQuotaCandidate>();
            }

            List<AiQuotaCandidate> candidates = new List<AiQuotaCandidate>();
            foreach (string code in scope.QuotaPoolCodes)
            {
                IndexQuota quota;
                if (quotasByCode.TryGetValue((code ?? "").Trim(), out quota))
                {
                    candidates.Add(new AiQuotaCandidate { Quota = quota, LocalScore = ScoreQuota(item, quota) });
                }
            }

            return candidates
                .OrderByDescending(c => c.LocalScore)
                .ThenBy(c => c.Quota.SortOrder)
                .Take(Math.Max(1, limit))
                .ToList();
        }

        private void LoadFiles(string quotaPath, string materialPath)
        {
            if (File.Exists(quotaPath))
            {
                foreach (string line in File.ReadLines(quotaPath, Encoding.UTF8))
                {
                    if (String.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                    IndexQuota quota = new IndexQuota();
                    quota.QuotaCode = LearningStore.Get(values, "quota_code");
                    quota.QuotaName = LearningStore.Get(values, "quota_name");
                    quota.QuotaUnit = LearningStore.Get(values, "quota_unit");
                    quota.BookCode = LearningStore.Get(values, "book_code");
                    quota.BookCategory = LearningStore.Get(values, "book_category");
                    quota.Specialty = LearningStore.Get(values, "specialty");
                    quota.SectionNo = LearningStore.Get(values, "section_no");
                    quota.SectionName = LearningStore.Get(values, "section_name");
                    quota.WorkContent = LearningStore.Get(values, "work_content");
                    double basePrice;
                    quota.BasePrice = Double.TryParse(
                        LearningStore.Get(values, "base_price"),
                        NumberStyles.Float,
                        CultureInfo.InvariantCulture,
                        out basePrice) ? basePrice : 0d;
                    quota.SearchText = LearningStore.Get(values, "search_text");
                    int sortOrder;
                    quota.SortOrder = Int32.TryParse(LearningStore.Get(values, "sort_order"), NumberStyles.Integer, CultureInfo.InvariantCulture, out sortOrder) ? sortOrder : Int32.MaxValue;
                    if (!String.IsNullOrWhiteSpace(quota.QuotaCode))
                    {
                        quotas.Add(quota);
                        if (!quotasByCode.ContainsKey(quota.QuotaCode.Trim()))
                        {
                            quotasByCode[quota.QuotaCode.Trim()] = quota;
                        }
                    }
                }

                BuildQuotaTokenIndex();
            }

            if (File.Exists(materialPath))
            {
                lock (materialSnapshotLock)
                {
                    materials = LoadMaterialFile(materialPath);
                }
            }
        }

        private static List<IndexMaterial> LoadMaterialFile(string materialPath)
        {
            List<IndexMaterial> loaded = new List<IndexMaterial>();
            if (!File.Exists(materialPath)) return loaded;

            foreach (string line in File.ReadLines(materialPath, Encoding.UTF8))
            {
                if (String.IsNullOrWhiteSpace(line)) continue;

                Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                IndexMaterial material = new IndexMaterial();
                material.MaterialCode = LearningStore.Get(values, "material_code");
                material.MaterialName = LearningStore.Get(values, "material_name");
                material.MaterialUnit = LearningStore.Get(values, "material_unit");
                material.DocNo = LearningStore.Get(values, "doc_no");
                material.IsMainMaterial = LearningStore.Get(values, "is_main_material") == "1";
                material.TransportCategory = LearningStore.Get(values, "transport_category");
                double basePrice;
                material.BasePrice = Double.TryParse(
                    LearningStore.Get(values, "base_price"),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out basePrice) ? basePrice : 0d;
                double currentPrice;
                material.CurrentPrice = Double.TryParse(
                    LearningStore.Get(values, "current_price"),
                    NumberStyles.Float,
                    CultureInfo.InvariantCulture,
                    out currentPrice) ? currentPrice : 0d;
                material.SearchText = LearningStore.Get(values, "search_text");
                if (!String.IsNullOrWhiteSpace(material.MaterialCode)) loaded.Add(material);
            }

            return loaded;
        }

        public Task RefreshMaterialsFromSourceAsync()
        {
            lock (materialRefreshLock)
            {
                if (materialRefreshTask == null)
                {
                    materialRefreshTask = Task.Factory.StartNew(
                        RefreshMaterialsFromSource,
                        CancellationToken.None,
                        TaskCreationOptions.None,
                        TaskScheduler.Default);
                }
                return materialRefreshTask;
            }
        }

        private void RefreshMaterialsFromSource()
        {
            try
            {
                string server = ReadServer();
                if (String.IsNullOrWhiteSpace(server)) server = "127.0.0.1";
                string databaseName = String.IsNullOrWhiteSpace(materialDatabaseName) ? ResolveDatabaseName() : materialDatabaseName;
                string connectionString = "Data Source=" + server + ",1433;Initial Catalog=" + databaseName + ";User ID=reco;Password=" + BuildSqlPassword() + ";Connect Timeout=8;Encrypt=False;TrustServerCertificate=True";
                List<IndexMaterial> refreshed;
                using (System.Data.SqlClient.SqlConnection connection = new System.Data.SqlClient.SqlConnection(connectionString))
                {
                    connection.Open();
                    refreshed = ReadMaterials(connection);
                }

                if (refreshed.Count == 0)
                {
                    throw new InvalidOperationException("Source material library returned no rows.");
                }

                WriteMaterialIndexFile(refreshed, materialIndexPath, databaseName);
                lock (materialSnapshotLock) materials = refreshed;
                lock (CacheLock)
                {
                    SearchIndexCacheEntry cached;
                    if (!String.IsNullOrWhiteSpace(cacheKey) && StoreCache.TryGetValue(cacheKey, out cached) && Object.ReferenceEquals(cached.Store, this))
                    {
                        cached.Fingerprint = BuildFileFingerprint(quotaIndexPath) + "|" + BuildFileFingerprint(materialIndexPath);
                    }
                }
                QuotaRecommendPanel.Log("Source material library refreshed: database=" + databaseName + " rows=" + refreshed.Count.ToString(CultureInfo.InvariantCulture));
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("Source material library refresh failed; cached materials retained: " + ex.Message);
            }
        }

        private void BuildQuotaTokenIndex()
        {
            quotaTokenIndex.Clear();
            foreach (IndexQuota quota in quotas)
            {
                foreach (string token in QuotaIndexTokens(quota))
                {
                    List<IndexQuota> bucket;
                    if (!quotaTokenIndex.TryGetValue(token, out bucket))
                    {
                        bucket = new List<IndexQuota>();
                        quotaTokenIndex[token] = bucket;
                    }
                    bucket.Add(quota);
                }
            }
        }

        private static IEnumerable<string> QuotaIndexTokens(IndexQuota quota)
        {
            string text = String.Join(" ", new string[] { quota.QuotaName, quota.WorkContent, quota.SectionName, quota.Specialty });
            foreach (string token in TextMatcher.Keywords(text).Distinct())
            {
                if (UseTokenForCandidateLookup(token))
                {
                    yield return token;
                }
            }
        }

        private List<IndexQuota> GetQuotaCandidates(ExcelQuantityItem item, string categoryFilter, EntryScope scope)
        {
            HashSet<IndexQuota> candidates = new HashSet<IndexQuota>();
            foreach (string token in CandidateLookupTokens(item.Name))
            {
                List<IndexQuota> bucket;
                if (!quotaTokenIndex.TryGetValue(token, out bucket))
                {
                    continue;
                }

                foreach (IndexQuota quota in bucket)
                {
                    if (!CategoryAllowed(quota.BookCategory, categoryFilter))
                    {
                        continue;
                    }

                    // 严格条目模式：候选只取当前条目定额池内的定额
                    if (scope != null && scope.Strict && !scope.Allows("quota", quota.QuotaCode))
                    {
                        continue;
                    }

                    candidates.Add(quota);
                }
            }

            return candidates.ToList();
        }

        private static IEnumerable<string> CandidateLookupTokens(string quantityName)
        {
            string normalized = TextMatcher.Normalize(quantityName).Replace(" ", "");
            if (UseTokenForCandidateLookup(normalized))
            {
                yield return normalized;
            }

            foreach (string token in TextMatcher.Keywords(quantityName).Distinct())
            {
                if (UseTokenForCandidateLookup(token))
                {
                    yield return token;
                }
            }
        }

        private static bool UseTokenForCandidateLookup(string token)
        {
            if (String.IsNullOrWhiteSpace(token))
            {
                return false;
            }

            if (token.Length == 1)
            {
                return TextMatcher.IsPureChinese(token);
            }

            return !TextMatcher.IsNumberLikeToken(token);
        }

        private static bool CategoryAllowed(string bookCategory, string categoryFilter)
        {
            string category = (bookCategory ?? "").Trim();
            string filter = (categoryFilter ?? "").Trim();
            if (String.IsNullOrWhiteSpace(filter) || String.Equals(filter, "\u5168\u90e8", StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (String.Equals(category, filter, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            return IsCommonCategory(category);
        }

        private static bool IsCommonCategory(string category)
        {
            return String.IsNullOrWhiteSpace(category) ||
                String.Equals(category, "\u57fa\u672c\u5b9a\u989d", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(category, "\u8865\u5145\u5b9a\u989d", StringComparison.OrdinalIgnoreCase) ||
                String.Equals(category, "\u8865\u5145\u5355\u4ef7\u5206\u6790", StringComparison.OrdinalIgnoreCase);
        }

        private static int ScoreQuota(ExcelQuantityItem item, IndexQuota quota)
        {
            string query = TextMatcher.Normalize(item.Name);
            string nameText = TextMatcher.Normalize(quota.QuotaName);
            string workText = TextMatcher.Normalize(quota.WorkContent);
            string sectionText = TextMatcher.Normalize(quota.SectionName);
            string specialtyText = TextMatcher.Normalize(quota.Specialty);
            string searchable = TextMatcher.Normalize(quota.SearchText);
            if (String.IsNullOrWhiteSpace(query) || String.IsNullOrWhiteSpace(searchable))
            {
                return 0;
            }

            if (IsSteelQuantity(query) && IsConcreteQuotaName(nameText))
            {
                return 0;
            }

            int score = 0;
            int primaryMatches = 0;
            int coreMatches = 0;
            if (nameText.Contains(query))
            {
                score += 90;
                primaryMatches++;
            }
            else if (workText.Contains(query))
            {
                score += 65;
                primaryMatches++;
            }

            List<string> tokens = TextMatcher.Keywords(item.Name).Distinct().ToList();
            int matched = 0;
            foreach (string token in tokens)
            {
                if (token.Length < 1 || !searchable.Contains(token))
                {
                    continue;
                }

                matched++;
                bool coreToken = TextMatcher.IsPureChinese(token) && token.Length >= 2;
                if (coreToken)
                {
                    coreMatches++;
                }

                bool primaryHit = nameText.Contains(token) || workText.Contains(token);
                if (primaryHit && token.Length >= 2)
                {
                    primaryMatches++;
                }

                if (nameText.Contains(token))
                {
                    score += TokenScore(token, 55, 42, 28);
                }
                else if (workText.Contains(token))
                {
                    score += TokenScore(token, 36, 28, 18);
                }
                else if (sectionText.Contains(token))
                {
                    score += TokenScore(token, 18, 14, 9);
                }
                else if (specialtyText.Contains(token))
                {
                    score += TokenScore(token, 10, 8, 5);
                }
            }

            if (IsSteelQuantity(query))
            {
                score += SteelPreferenceScore(query, nameText, workText, quota.QuotaUnit);
            }

            if (primaryMatches == 0 && coreMatches < 2)
            {
                return 0;
            }

            if (tokens.Count > 0)
            {
                score += (int)Math.Round(20.0 * matched / tokens.Count);
            }

            if (RecommendDialog.UnitCompatibleForIndex(quota.QuotaUnit, item.Unit))
            {
                score += 12;
            }

            return score;
        }

        private static int TokenScore(string token, int shortChinese, int longChinese, int mixed)
        {
            if (TextMatcher.HasAsciiOrDigit(token))
            {
                return mixed;
            }

            if (token.Length == 1)
            {
                return Math.Max(1, shortChinese / 10);
            }

            return token.Length == 2 ? shortChinese : longChinese;
        }

        private static bool IsSteelQuantity(string normalizedQuery)
        {
            return TextMatcher.IsSteelQuantityName(normalizedQuery);
        }

        private static bool IsConcreteQuotaName(string normalizedQuotaName)
        {
            if (!TextMatcher.IsConcreteQuantityName(normalizedQuotaName))
            {
                return false;
            }

            return !normalizedQuotaName.Contains("构件钢筋") &&
                !normalizedQuotaName.Contains("圆钢筋") &&
                !normalizedQuotaName.Contains("螺纹钢筋") &&
                !normalizedQuotaName.Contains("箍筋") &&
                !normalizedQuotaName.Contains("钢筋制作") &&
                !normalizedQuotaName.Contains("钢筋制安") &&
                !normalizedQuotaName.Contains("钢筋绑扎");
        }

        private static int SteelPreferenceScore(string query, string nameText, string workText, string quotaUnit)
        {
            int score = 0;
            bool steelOperation = nameText.Contains("构件钢筋") ||
                nameText.Contains("钢筋制作") ||
                nameText.Contains("钢筋制安") ||
                nameText.Contains("钢筋绑扎") ||
                workText.Contains("钢筋制作") ||
                workText.Contains("钢筋制安") ||
                (workText.Contains("钢筋") && (workText.Contains("制作") || workText.Contains("绑扎") || workText.Contains("安装")));

            if (steelOperation)
            {
                score += 55;
            }

            if ((query.Contains("hpb") || query.Contains("光圆") || query.Contains("圆钢")) && nameText.Contains("圆钢筋"))
            {
                score += 80;
            }

            if ((query.Contains("hrb") || query.Contains("螺纹")) && (nameText.Contains("hrb") || nameText.Contains("螺纹钢筋")))
            {
                score += 80;
            }

            if (nameText.Contains("钢筋混凝土") && !steelOperation)
            {
                score -= 70;
            }

            if (!RecommendDialog.UnitCompatibleForIndex(quotaUnit, "kg") && !RecommendDialog.UnitCompatibleForIndex(quotaUnit, "t"))
            {
                score -= 30;
            }

            return score;
        }

        private static void ExportQuotaFromSql(string dataDir, string quotaPath, string databaseName)
        {
            Directory.CreateDirectory(dataDir);
            string server = ReadServer();
            if (String.IsNullOrWhiteSpace(server))
            {
                server = "127.0.0.1";
            }

            QuotaRecommendPanel.Log("Build quota search index from database: " + databaseName);
            string connectionString = "Data Source=" + server + ",1433;Initial Catalog=" + databaseName + ";User ID=reco;Password=" + BuildSqlPassword() + ";Connect Timeout=8;Encrypt=False;TrustServerCertificate=True";
            using (System.Data.SqlClient.SqlConnection connection = new System.Data.SqlClient.SqlConnection(connectionString))
            {
                connection.Open();
                WriteQuotaIndex(connection, quotaPath, databaseName);
            }
        }

        private static string ReadServer()
        {
            try
            {
                string baseDir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location);
                string path = Path.Combine(baseDir, "ServerSetting.xml");
                if (!File.Exists(path))
                {
                    return "";
                }

                string text = File.ReadAllText(path, Encoding.UTF8);
                int start = text.IndexOf("<ServerIP>", StringComparison.OrdinalIgnoreCase);
                int end = text.IndexOf("</ServerIP>", StringComparison.OrdinalIgnoreCase);
                if (start >= 0 && end > start)
                {
                    return text.Substring(start + 10, end - start - 10).Trim();
                }

                start = text.IndexOf("<Server>", StringComparison.OrdinalIgnoreCase);
                end = text.IndexOf("</Server>", StringComparison.OrdinalIgnoreCase);
                return start >= 0 && end > start ? text.Substring(start + 8, end - start - 8).Trim() : "";
            }
            catch
            {
                return "";
            }
        }

        private static string BuildSqlPassword()
        {
            return String.Join("_", new string[] { "Des", "Reco", "2006" });
        }

        private static string ResolveDatabaseName()
        {
            try
            {
                string baseDir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location) ?? "";
                string processIdentity = "";
                try
                {
                    Process process = Process.GetCurrentProcess();
                    processIdentity = process.ProcessName ?? "";
                    try
                    {
                        processIdentity += " " + (process.MainModule.FileName ?? "");
                    }
                    catch
                    {
                    }
                }
                catch
                {
                }

                return ResolveDatabaseNameForHost(
                    baseDir,
                    processIdentity,
                    File.Exists(Path.Combine(baseDir, "RejjNet2020.exe")),
                    File.Exists(Path.Combine(baseDir, "ReJJGSNet2024.exe")) || File.Exists(Path.Combine(baseDir, "ReJJQDNet2024.exe")));
            }
            catch
            {
            }

            return "RecoData2020";
        }

        internal static string ResolveDatabaseNameForHost(string baseDir, string processIdentity, bool has2020Executable, bool has2024Executable)
        {
            string processProbe = (processIdentity ?? "").ToLowerInvariant();
            if (processProbe.Contains("rejjnet2020")) return "RecoData2020";
            if (processProbe.Contains("rejjgsnet2024") || processProbe.Contains("rejjqdnet2024")) return "RecoData2024";
            if (has2020Executable && !has2024Executable) return "RecoData2020";
            if (has2024Executable && !has2020Executable) return "RecoData2024";

            string directoryProbe = (baseDir ?? "").ToLowerInvariant();
            return directoryProbe.Contains("2024") ? "RecoData2024" : "RecoData2020";
        }

        private static void WriteQuotaIndex(System.Data.SqlClient.SqlConnection connection, string path, string databaseName)
        {
            string temp = path + ".tmp";
            using (StreamWriter writer = new StreamWriter(temp, false, Encoding.UTF8))
            using (System.Data.SqlClient.SqlCommand command = connection.CreateCommand())
            {
                command.CommandTimeout = 60;
                command.CommandText =
                    "SELECT a.定额编号,a.定额名称,a.单位,a.书号,ISNULL(b.分类,''),ISNULL(b.专业名称,''),ISNULL(a.节号,''),ISNULL(c.节名称,''),ISNULL(CAST(a.工作内容 AS nvarchar(max)),''),ISNULL(a.基价,0),ISNULL(b.现行定额,0),ISNULL(a.流水号,2147483647) " +
                    "FROM dbo.定额库 a LEFT JOIN dbo.定额库索引 b ON a.书号=b.书号 LEFT JOIN dbo.定额节索引 c ON a.书号=c.书号 AND a.节号=c.节号 " +
                    "WHERE ISNULL(a.定额编号,'')<>'' AND ISNULL(a.定额名称,'')<>''";
                using (System.Data.SqlClient.SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        Dictionary<string, string> row = new Dictionary<string, string>();
                        row["quota_code"] = ReadString(reader, 0);
                        row["source_database"] = databaseName ?? "";
                        row["quota_name"] = ReadString(reader, 1);
                        row["quota_unit"] = ReadString(reader, 2);
                        row["book_code"] = ReadString(reader, 3);
                        row["book_category"] = ReadString(reader, 4);
                        row["specialty"] = ReadString(reader, 5);
                        row["section_no"] = ReadString(reader, 6);
                        row["section_name"] = ReadString(reader, 7);
                        row["work_content"] = ReadString(reader, 8);
                        row["base_price"] = Convert.ToString(reader.GetDouble(9), CultureInfo.InvariantCulture);
                        row["is_current"] = IsTruthy(reader.GetValue(10)) ? "1" : "0";
                        row["sort_order"] = Convert.ToString(reader.GetInt32(11), CultureInfo.InvariantCulture);
                        row["search_text"] = TextMatcher.Normalize(String.Join(" ", new string[] { row["quota_code"], row["quota_name"], row["quota_unit"], row["book_category"], row["specialty"], row["section_name"], row["work_content"] }));
                        writer.WriteLine(LearningStore.ToJson(row));
                    }
                }
            }

            ReplaceFile(temp, path);
        }

        private static List<IndexMaterial> ReadMaterials(System.Data.SqlClient.SqlConnection connection)
        {
            List<IndexMaterial> result = new List<IndexMaterial>();
            using (System.Data.SqlClient.SqlCommand command = connection.CreateCommand())
            {
                command.CommandTimeout = 60;
                command.CommandText =
                    "SELECT 电算代号,材料名称,单位,文号,ISNULL(主材标志,''),ISNULL(材料运输类别,''),ISNULL(基期单价,0),ISNULL(编制期价,0) " +
                    "FROM dbo.材料单价库 WHERE ISNULL(材料名称,'')<>''";
                using (System.Data.SqlClient.SqlDataReader reader = command.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        IndexMaterial material = new IndexMaterial();
                        material.MaterialCode = Convert.ToString(reader.GetValue(0), CultureInfo.InvariantCulture);
                        material.MaterialName = ReadString(reader, 1);
                        material.MaterialUnit = ReadString(reader, 2);
                        material.DocNo = ReadString(reader, 3);
                        material.IsMainMaterial = ReadString(reader, 4) == "1";
                        material.TransportCategory = ReadString(reader, 5);
                        material.BasePrice = reader.GetDouble(6);
                        material.CurrentPrice = reader.GetDouble(7);
                        material.SearchText = TextMatcher.Normalize(String.Join(" ", new string[]
                        {
                            material.MaterialCode,
                            material.MaterialName,
                            material.MaterialUnit,
                            material.DocNo,
                            material.TransportCategory
                        }));
                        if (!String.IsNullOrWhiteSpace(material.MaterialCode)) result.Add(material);
                    }
                }
            }
            return result;
        }

        private static void WriteMaterialIndexFile(IEnumerable<IndexMaterial> materialsToWrite, string path, string databaseName)
        {
            if (String.IsNullOrWhiteSpace(path)) return;
            string temp = path + ".tmp";
            using (StreamWriter writer = new StreamWriter(temp, false, Encoding.UTF8))
            {
                foreach (IndexMaterial material in materialsToWrite ?? new List<IndexMaterial>())
                {
                    Dictionary<string, string> row = new Dictionary<string, string>();
                    row["material_code"] = material.MaterialCode ?? "";
                    row["source_database"] = databaseName ?? "";
                    row["material_name"] = material.MaterialName ?? "";
                    row["material_unit"] = material.MaterialUnit ?? "";
                    row["doc_no"] = material.DocNo ?? "";
                    row["is_main_material"] = material.IsMainMaterial ? "1" : "0";
                    row["transport_category"] = material.TransportCategory ?? "";
                    row["base_price"] = Convert.ToString(material.BasePrice, CultureInfo.InvariantCulture);
                    row["current_price"] = Convert.ToString(material.CurrentPrice, CultureInfo.InvariantCulture);
                    row["search_text"] = material.SearchText ?? "";
                    writer.WriteLine(LearningStore.ToJson(row));
                }
            }

            ReplaceFile(temp, path);
        }

        private static string ReadString(System.Data.SqlClient.SqlDataReader reader, int index)
        {
            return reader.IsDBNull(index) ? "" : Convert.ToString(reader.GetValue(index), CultureInfo.CurrentCulture).Trim();
        }

        private static bool IsTruthy(object value)
        {
            if (value == null || value == DBNull.Value)
            {
                return false;
            }

            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture) != 0;
            }
            catch
            {
                return String.Equals(Convert.ToString(value, CultureInfo.InvariantCulture), "true", StringComparison.OrdinalIgnoreCase);
            }
        }

        private static void ReplaceFile(string temp, string path)
        {
            if (File.Exists(path))
            {
                File.Replace(temp, path, null);
                return;
            }
            File.Move(temp, path);
        }

        private sealed class ScoredQuota
        {
            public IndexQuota Quota;
            public int Score;
            public int MajorRank;
            public int PoolRank;
        }

        private sealed class ScoredMaterial
        {
            public IndexMaterial Material;
            public string NormalizedName;
            public int MatchRank;
            public int MatchIndex;
        }

    }

    internal sealed class IndexQuota
    {
        public string QuotaCode;
        public string QuotaName;
        public string QuotaUnit;
        public string BookCode;
        public string BookCategory;
        public string Specialty;
        public string SectionNo;
        public string SectionName;
        public string WorkContent;
        public double BasePrice;
        public string SearchText;
        public int SortOrder;

        public RecommendationRow ToRecommendation(ExcelQuantityItem item, int score)
        {
            RecommendationRow row = new RecommendationRow();
            row.Item = item;
            row.QuotaCode = QuotaCode;
            row.QuotaName = QuotaName;
            row.QuotaUnit = QuotaUnit;
            row.BookCode = BookCode;
            row.Specialty = Specialty;
            row.BasePrice = BasePrice;
            row.WorkContent = WorkContent;
            row.ConvertedValueText = RecommendDialog.ConvertQuantityForIndex(item.ValueText, item.Unit, QuotaUnit);
            row.Score = score;
            row.Reason = "\u5168\u91cf\u5b9a\u989d\u7d22\u5f15\u5173\u952e\u8bcd\u5339\u914d";
            row.Source = "index";
            row.TargetKind = "quota";
            return row;
        }
    }

    internal sealed class IndexMaterial
    {
        public string MaterialCode;
        public string MaterialName;
        public string MaterialUnit;
        public string DocNo;
        public bool IsMainMaterial;
        public string TransportCategory;
        public double BasePrice;
        public double CurrentPrice;
        public string SearchText;

        public RecommendationRow ToRecommendation(ExcelQuantityItem item, int score)
        {
            RecommendationRow row = new RecommendationRow();
            row.Item = item;
            row.QuotaCode = MaterialCode;
            row.QuotaName = MaterialName;
            row.QuotaUnit = MaterialUnit;
            row.BookCode = DocNo;
            row.BasePrice = BasePrice;
            row.WorkContent = DocNo;
            row.ConvertedValueText = RecommendDialog.ConvertQuantityForIndex(item.ValueText, item.Unit, MaterialUnit);
            row.Score = score;
            row.Reason = IsMainMaterial ? "\u4e3b\u8981\u6750\u6599\u7d22\u5f15\u5173\u952e\u8bcd\u5339\u914d" : "\u6750\u6599\u7d22\u5f15\u5173\u952e\u8bcd\u5339\u914d";
            row.Source = "index";
            row.TargetKind = "material";
            return row;
        }
    }
}
