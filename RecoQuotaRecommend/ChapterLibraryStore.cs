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
    // 当前定额行所在章节条目的推荐范围；Strict=true 时三个阶段的候选都严格限制在该条目定额池内
    internal sealed class EntryScope
    {
        public string ProjectEntryCode;   // 项目章节表里的原始条目编号
        public string MatchedEntryCode;   // 前缀上溯后命中的库内条目编号
        public string EntryName;
        public string Method;             // "2020" / "2024"
        public string MethodNo;           // 项目编制办法文号，用于区分 30号文/101号文/2024 等条目池
        public HashSet<string> PoolKeys;  // "kind:CODE"（大写）

        public bool Strict
        {
            get { return !String.IsNullOrEmpty(MatchedEntryCode) && PoolKeys != null && PoolKeys.Count > 0; }
        }

        // 对应框 entry_codes 标签格式：method:条目编号
        public string Tag
        {
            get { return Method + ":" + MatchedEntryCode; }
        }

        public bool Allows(string targetKind, string code)
        {
            if (!Strict)
            {
                return true;
            }

            string kind = String.IsNullOrWhiteSpace(targetKind) ? QuotaEntry.GuessKind(code) : targetKind;
            string key = kind.ToLowerInvariant() + ":" + QuotaEntry.NormalizeCode(code).ToUpperInvariant();
            if (PoolKeys.Contains(key))
            {
                return true;
            }
            return PoolKeys.Any(k => k.StartsWith(key + "|", StringComparison.Ordinal));
        }

        // 条目定额池里所有 quota 类定额编号（去掉 "quota:" 前缀），用于把整池作为候选喂给本地匹配/AI
        public IEnumerable<string> QuotaPoolCodes
        {
            get
            {
                if (PoolKeys == null)
                {
                    yield break;
                }

                foreach (string key in PoolKeys)
                {
                    if (key.StartsWith("quota:", StringComparison.OrdinalIgnoreCase))
                    {
                        string code = key.Substring("quota:".Length);
                        int detailSep = code.IndexOf('|');
                        yield return detailSep >= 0 ? code.Substring(0, detailSep) : code;
                    }
                }
            }
        }
    }

    // 章节条目定额库：chapter-entries.jsonl（删减后的条目树）+ chapter-quota-library.jsonl（条目定额池）。
    // chapter-entries.jsonl 不存在时 IsEmpty=true，推荐行为与历史版本完全一致。
    internal sealed class ChapterLibraryStore
    {
        private readonly Dictionary<string, string> entryNames = new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly Dictionary<string, string> entryTypes = new Dictionary<string, string>(StringComparer.Ordinal);
        private readonly Dictionary<string, HashSet<string>> pools = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
        // 规范化条目名称 → 小计/指标条目编号列表（识别用户复制条目的来源）
        private readonly Dictionary<string, List<string>> nameIndex = new Dictionary<string, List<string>>(StringComparer.Ordinal);
        private static readonly object StoreCacheLock = new object();
        private static readonly Dictionary<string, ChapterLibraryCacheEntry> StoreCache = new Dictionary<string, ChapterLibraryCacheEntry>(StringComparer.OrdinalIgnoreCase);
        public string MethodKey = "";
        public string MethodNo = "";

        private sealed class ChapterLibraryCacheEntry
        {
            public string Fingerprint;
            public ChapterLibraryStore Store;
        }

        public bool IsEmpty
        {
            get { return entryNames.Count == 0 && pools.Count == 0; }
        }

        public static ChapterLibraryStore Load()
        {
            string methodKey = ResolveMethodKey();
            string dataDir = LearningStore.FindDataDir();
            string entriesPath = Path.Combine(dataDir, "chapter-entries.jsonl");
            string libraryPath = Path.Combine(dataDir, "chapter-quota-library.jsonl");
            string quotaPath = Path.Combine(dataDir, "quota-index.jsonl");
            string cacheKey = Path.GetFullPath(dataDir).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + "|" + methodKey;
            lock (StoreCacheLock)
            {
                string fingerprint = BuildFileFingerprint(entriesPath) + "|" + BuildFileFingerprint(libraryPath) + "|" + BuildFileFingerprint(quotaPath);
                ChapterLibraryCacheEntry cached;
                if (StoreCache.TryGetValue(cacheKey, out cached) &&
                    String.Equals(cached.Fingerprint, fingerprint, StringComparison.Ordinal))
                {
                    return cached.Store;
                }

                ChapterLibraryStore store = new ChapterLibraryStore();
                store.MethodKey = methodKey;
                if (!File.Exists(entriesPath) || !File.Exists(libraryPath))
                {
                    if (cached != null && cached.Store != null)
                    {
                        cached.Store.ReplaceSnapshot(store);
                        cached.Fingerprint = fingerprint;
                        return cached.Store;
                    }
                    StoreCache[cacheKey] = new ChapterLibraryCacheEntry { Fingerprint = fingerprint, Store = store };
                    return store;
                }

                bool loadSucceeded = true;
                try
                {
                foreach (string line in File.ReadLines(entriesPath, Encoding.UTF8))
                {
                    if (String.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                    if (!String.Equals(LearningStore.Get(values, "method"), store.MethodKey, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    string code = LearningStore.Get(values, "entry_code").Trim();
                    if (!String.IsNullOrEmpty(code) && !store.entryNames.ContainsKey(code))
                    {
                        store.entryNames[code] = LearningStore.Get(values, "entry_name").Trim();
                        store.entryTypes[code] = LearningStore.Get(values, "entry_type").Trim();
                        if (String.IsNullOrEmpty(store.MethodNo))
                        {
                            store.MethodNo = LearningStore.Get(values, "method_no").Trim();
                        }
                    }
                }

                HashSet<string> validQuotaCodes = LoadReferenceQuotaCodes(dataDir);

                foreach (string line in File.ReadLines(libraryPath, Encoding.UTF8))
                {
                    if (String.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                    if (!String.Equals(LearningStore.Get(values, "method"), store.MethodKey, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    string entryCode = LearningStore.Get(values, "entry_code").Trim();
                    string code = QuotaEntry.NormalizeCode(LearningStore.Get(values, "quota_code").Trim());
                    string methodNo = LearningStore.Get(values, "method_no").Trim();
                    string targetKind = LearningStore.Get(values, "target_kind").Trim();
                    if (String.IsNullOrEmpty(entryCode) || String.IsNullOrEmpty(code))
                    {
                        continue;
                    }
                    if (String.IsNullOrEmpty(store.MethodNo) && !String.IsNullOrEmpty(methodNo))
                    {
                        store.MethodNo = methodNo;
                    }

                    string entryName = LearningStore.Get(values, "entry_name").Trim();
                    string knownEntryName;
                    if (String.IsNullOrEmpty(entryName) && store.entryNames.TryGetValue(entryCode, out knownEntryName))
                    {
                        entryName = knownEntryName;
                    }
                    string quotaName = LearningStore.Get(values, "quota_name").Trim();
                    string quotaUnit = LearningStore.Get(values, "quota_unit").Trim();
                    string quotaPrice = FirstNonEmpty(values, "base_price", "quota_price", "price");
                    if (!IsAllowedReferencePoolCode(entryName, targetKind, code, validQuotaCodes))
                    {
                        continue;
                    }

                    if (LearningStore.Get(values, "deleted").Trim() == "1")
                    {
                        store.RemovePoolKey(methodNo, entryCode, ReferencePoolKind(entryName, targetKind, code), code, quotaName, quotaUnit, quotaPrice);
                    }
                    else
                    {
                        store.AddPoolKey(methodNo, entryCode, ReferencePoolKind(entryName, targetKind, code), code, quotaName, quotaUnit, quotaPrice);
                    }
                }

                store.BuildNameIndex();
                QuotaRecommendPanel.Log("ChapterLibraryStore loaded. method=" + store.MethodKey + " entries=" + store.entryNames.Count.ToString(CultureInfo.InvariantCulture) + " pooledEntries=" + store.pools.Count.ToString(CultureInfo.InvariantCulture));
                }
                catch (Exception ex)
                {
                    loadSucceeded = false;
                    QuotaRecommendPanel.Log("ChapterLibraryStore load failed: " + ex.Message);
                }

                if (!loadSucceeded)
                {
                    return cached != null && cached.Store != null ? cached.Store : store;
                }

                string loadedFingerprint = BuildFileFingerprint(entriesPath) + "|" + BuildFileFingerprint(libraryPath) + "|" + BuildFileFingerprint(quotaPath);
                if (cached != null && cached.Store != null)
                {
                    cached.Store.ReplaceSnapshot(store);
                    cached.Fingerprint = loadedFingerprint;
                    return cached.Store;
                }

                StoreCache[cacheKey] = new ChapterLibraryCacheEntry
                {
                    Fingerprint = loadedFingerprint,
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

        private void ReplaceSnapshot(ChapterLibraryStore source)
        {
            entryNames.Clear();
            foreach (KeyValuePair<string, string> pair in source.entryNames)
            {
                entryNames[pair.Key] = pair.Value;
            }
            entryTypes.Clear();
            foreach (KeyValuePair<string, string> pair in source.entryTypes)
            {
                entryTypes[pair.Key] = pair.Value;
            }
            pools.Clear();
            foreach (KeyValuePair<string, HashSet<string>> pair in source.pools)
            {
                pools[pair.Key] = new HashSet<string>(pair.Value, StringComparer.Ordinal);
            }
            MethodKey = source.MethodKey;
            MethodNo = source.MethodNo;
            BuildNameIndex();
        }

        private void RefreshCachedFingerprint()
        {
            string dataDir = LearningStore.FindDataDir();
            string fingerprint = BuildFileFingerprint(Path.Combine(dataDir, "chapter-entries.jsonl")) + "|"
                + BuildFileFingerprint(Path.Combine(dataDir, "chapter-quota-library.jsonl")) + "|"
                + BuildFileFingerprint(Path.Combine(dataDir, "quota-index.jsonl"));
            lock (StoreCacheLock)
            {
                foreach (ChapterLibraryCacheEntry cached in StoreCache.Values)
                {
                    if (cached != null && Object.ReferenceEquals(cached.Store, this))
                    {
                        cached.Fingerprint = fingerprint;
                    }
                }
            }
        }

        // 与 SearchIndexStore.ResolveDatabaseName 同一套判断：按运行目录/进程判定 2020 还是 2024
        private static string ResolveMethodKey()
        {
            try
            {
                string baseDir = Path.GetDirectoryName(typeof(QuotaRecommendPanel).Assembly.Location) ?? "";
                string processPath = "";
                try
                {
                    processPath = Process.GetCurrentProcess().MainModule.FileName ?? "";
                }
                catch
                {
                }

                string probe = (baseDir + " " + processPath).ToLowerInvariant();
                if (probe.Contains("2024") ||
                    File.Exists(Path.Combine(baseDir, "ReJJGSNet2024.exe")) ||
                    File.Exists(Path.Combine(baseDir, "ReJJQDNet2024.exe")))
                {
                    return "2024";
                }
            }
            catch
            {
            }

            return "2020";
        }

        private static string NormalizePoolMethodNo(string text)
        {
            return (text ?? "").Replace('\u2013', '-').Replace('\u2014', '-').Replace('\uff0d', '-').Replace(" ", "").Trim();
        }

        private string EffectiveMethodNo(string methodNo)
        {
            return String.IsNullOrWhiteSpace(methodNo) ? (MethodNo ?? "") : methodNo.Trim();
        }

        private static string PoolKey(string methodNo, string entryCode)
        {
            return NormalizePoolMethodNo(methodNo) + "|" + ((entryCode ?? "").Trim());
        }

        private static HashSet<string> LoadReferenceQuotaCodes(string dataDir)
        {
            HashSet<string> result = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                string path = Path.Combine(dataDir, "quota-index.jsonl");
                if (!File.Exists(path))
                {
                    return result;
                }

                foreach (string line in File.ReadLines(path, Encoding.UTF8))
                {
                    if (String.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    Dictionary<string, string> values = LearningStore.ParseFlatJson(line);
                    string code = QuotaEntry.NormalizeCode(LearningStore.Get(values, "quota_code").Trim());
                    if (!String.IsNullOrEmpty(code))
                    {
                        result.Add(code.ToUpperInvariant());
                    }
                }
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("Reference quota code whitelist load failed: " + ex.Message);
            }
            return result;
        }

        private static bool IsEquipmentPurchaseEntry(string entryName)
        {
            return NormalizeEntryName(entryName).IndexOf("\u8bbe\u5907\u8d2d\u7f6e\u8d39", StringComparison.Ordinal) >= 0;
        }

        private static bool IsSfCode(string code)
        {
            return String.Equals(QuotaEntry.NormalizeCode(code), "SF", StringComparison.OrdinalIgnoreCase);
        }

        private static bool IsAllowedReferencePoolCode(string entryName, string targetKind, string code, HashSet<string> validQuotaCodes)
        {
            string normalizedCode = QuotaEntry.NormalizeCode(code);
            if (String.IsNullOrEmpty(normalizedCode))
            {
                return false;
            }

            if (IsEquipmentPurchaseEntry(entryName))
            {
                return IsSfCode(normalizedCode);
            }

            string kind = String.IsNullOrWhiteSpace(targetKind) ? QuotaEntry.GuessKind(normalizedCode) : targetKind.Trim().ToLowerInvariant();
            if (!String.Equals(kind, "quota", StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            return validQuotaCodes != null && validQuotaCodes.Contains(normalizedCode.ToUpperInvariant());
        }

        private static string ReferencePoolKind(string entryName, string targetKind, string code)
        {
            return "quota";
        }

        private static string FirstNonEmpty(Dictionary<string, string> values, params string[] keys)
        {
            foreach (string key in keys)
            {
                string value = LearningStore.Get(values, key).Trim();
                if (!String.IsNullOrWhiteSpace(value))
                {
                    return value;
                }
            }
            return "";
        }

        private static string ReferencePoolItemKey(string targetKind, string code, string name, string unit, string price)
        {
            string kind = String.IsNullOrWhiteSpace(targetKind) ? QuotaEntry.GuessKind(code) : targetKind.Trim().ToLowerInvariant();
            string normalizedCode = QuotaEntry.NormalizeCode(code).ToUpperInvariant();
            string key = kind + ":" + normalizedCode;
            if (IsSfCode(normalizedCode))
            {
                key += "|" + NormalizeReferencePoolIdentity(name)
                    + "|" + NormalizeReferencePoolIdentity(unit)
                    + "|" + NormalizeReferencePoolIdentity(price);
            }
            return key;
        }

        private static string NormalizeReferencePoolIdentity(string value)
        {
            return (value ?? "").Trim().ToUpperInvariant();
        }

        private static bool TrySplitPoolKey(string poolKey, out string methodNo, out string entryCode)
        {
            methodNo = "";
            entryCode = "";
            if (String.IsNullOrEmpty(poolKey))
            {
                return false;
            }
            int sep = poolKey.IndexOf('|');
            if (sep < 0)
            {
                entryCode = poolKey;
                return true;
            }
            methodNo = poolKey.Substring(0, sep);
            entryCode = poolKey.Substring(sep + 1);
            return !String.IsNullOrEmpty(entryCode);
        }

        private void AddPoolKey(string methodNo, string entryCode, string kind, string code)
        {
            AddPoolKey(methodNo, entryCode, kind, code, "", "", "");
        }

        private void AddPoolKey(string methodNo, string entryCode, string kind, string code, string name, string unit, string price)
        {
            HashSet<string> pool;
            string poolKey = PoolKey(methodNo, entryCode);
            if (!pools.TryGetValue(poolKey, out pool))
            {
                pool = new HashSet<string>(StringComparer.Ordinal);
                pools[poolKey] = pool;
            }

            string normalizedKind = String.IsNullOrWhiteSpace(kind) ? QuotaEntry.GuessKind(code) : kind.Trim().ToLowerInvariant();
            pool.Add(ReferencePoolItemKey(normalizedKind, code, name, unit, price));
        }

        private void RemovePoolKey(string methodNo, string entryCode, string kind, string code)
        {
            RemovePoolKey(methodNo, entryCode, kind, code, "", "", "");
        }

        private void RemovePoolKey(string methodNo, string entryCode, string kind, string code, string name, string unit, string price)
        {
            HashSet<string> pool;
            string poolKey = PoolKey(methodNo, entryCode);
            if (String.IsNullOrEmpty(entryCode) || String.IsNullOrEmpty(code) || !pools.TryGetValue(poolKey, out pool))
            {
                return;
            }

            string normalizedKind = String.IsNullOrWhiteSpace(kind) ? QuotaEntry.GuessKind(code) : kind.Trim().ToLowerInvariant();
            string key = ReferencePoolItemKey(normalizedKind, code, name, unit, price);
            pool.Remove(key);
            if (IsSfCode(code) && String.IsNullOrWhiteSpace(name) && String.IsNullOrWhiteSpace(unit) && String.IsNullOrWhiteSpace(price))
            {
                string prefix = normalizedKind + ":" + QuotaEntry.NormalizeCode(code).ToUpperInvariant() + "|";
                foreach (string existing in pool.Where(x => x.StartsWith(prefix, StringComparison.Ordinal)).ToList())
                {
                    pool.Remove(existing);
                }
            }
            if (pool.Count == 0)
            {
                pools.Remove(poolKey);
            }
        }

        private static bool IsQuotaInputEntryType(string entryType)
        {
            return entryType == "小计" || entryType == "指标";
        }

        private static string NormalizeEntryName(string name)
        {
            return Regex.Replace(name ?? "", "[\\s　]+", "").ToLowerInvariant();
        }

        private static string NameIndexKey(string methodNo, string normalizedName)
        {
            if (String.IsNullOrEmpty(normalizedName))
            {
                return "";
            }
            return NormalizePoolMethodNo(methodNo) + "|" + normalizedName;
        }

        // 名称索引只收"小计/指标"且有池的条目——只有这两类条目有定额输入框
        private void BuildNameIndex()
        {
            nameIndex.Clear();
            foreach (KeyValuePair<string, HashSet<string>> poolPair in pools)
            {
                if (poolPair.Value == null || poolPair.Value.Count == 0)
                {
                    continue;
                }

                string methodNo;
                string entryCode;
                if (!TrySplitPoolKey(poolPair.Key, out methodNo, out entryCode))
                {
                    continue;
                }

                string entryType;
                if (!entryTypes.TryGetValue(entryCode, out entryType) || !IsQuotaInputEntryType(entryType))
                {
                    continue;
                }

                string entryName;
                if (!entryNames.TryGetValue(entryCode, out entryName))
                {
                    continue;
                }

                string nameKey = NameIndexKey(methodNo, NormalizeEntryName(entryName));
                if (String.IsNullOrEmpty(nameKey))
                {
                    continue;
                }

                List<string> codes;
                if (!nameIndex.TryGetValue(nameKey, out codes))
                {
                    codes = new List<string>();
                    nameIndex[nameKey] = codes;
                }
                if (!codes.Contains(entryCode))
                {
                    codes.Add(entryCode);
                }
            }
        }

        // 项目条目编号 → 库内保留条目。
        // 顺序：编号精确命中 → 按名称识别"复制条目"的来源（同祖先链优先，再全局唯一）→ 逐级前缀上溯。
        public EntryScope ResolveScope(string projectEntryCode, string projectEntryName)
        {
            return ResolveScope(MethodNo, projectEntryCode, projectEntryName);
        }

        public EntryScope ResolveScopeForUserEdit(string projectEntryCode, string projectEntryName)
        {
            return ResolveScopeForUserEdit(MethodNo, projectEntryCode, projectEntryName);
        }

        public EntryScope ResolveScopeForUserEdit(string methodNo, string projectEntryCode, string projectEntryName)
        {
            EntryScope scope = ResolveScope(methodNo, projectEntryCode, projectEntryName);
            if (scope != null)
            {
                return scope;
            }

            methodNo = EffectiveMethodNo(methodNo);
            string current = (projectEntryCode ?? "").Trim();
            if (String.IsNullOrEmpty(current))
            {
                return null;
            }

            return BuildEditableScope(methodNo, current, current, projectEntryName);
        }

        public EntryScope ResolveScope(string methodNo, string projectEntryCode, string projectEntryName)
        {
            methodNo = EffectiveMethodNo(methodNo);
            string current = (projectEntryCode ?? "").Trim();
            if (String.IsNullOrEmpty(current) || IsEmpty)
            {
                return null;
            }

            EntryScope exact = BuildScope(methodNo, current, current, projectEntryName);
            if (exact != null)
            {
                return exact;
            }

            // 编号不在库内（用户新建/复制的条目）⇒ 按名称找复制来源条目，用它的定额池
            if (!entryNames.ContainsKey(current))
            {
                string nameKey = NameIndexKey(methodNo, NormalizeEntryName(projectEntryName));
                List<string> sameName;
                if (!String.IsNullOrEmpty(nameKey) && nameIndex.TryGetValue(nameKey, out sameName) && sameName.Count > 0)
                {
                    string prefix = current;
                    while (true)
                    {
                        int dash = prefix.LastIndexOf('-');
                        if (dash <= 0)
                        {
                            break;
                        }
                        prefix = prefix.Substring(0, dash);
                        string withDash = prefix + "-";
                        foreach (string candidate in sameName)
                        {
                            if (candidate.StartsWith(withDash, StringComparison.Ordinal) || candidate == prefix)
                            {
                                EntryScope copied = BuildScope(methodNo, current, candidate, projectEntryName);
                                if (copied != null)
                                {
                                    return copied;
                                }
                            }
                        }
                    }

                    if (sameName.Count == 1)
                    {
                        EntryScope unique = BuildScope(methodNo, current, sameName[0], projectEntryName);
                        if (unique != null)
                        {
                            return unique;
                        }
                    }
                }
            }

            // 逐级前缀上溯到最近的"保留且池非空"条目
            string probe = current;
            while (!String.IsNullOrEmpty(probe) && probe != "0")
            {
                EntryScope scope = BuildScope(methodNo, current, probe, projectEntryName);
                if (scope != null)
                {
                    return scope;
                }

                int dash2 = probe.LastIndexOf('-');
                if (dash2 > 0)
                {
                    probe = probe.Substring(0, dash2);
                    continue;
                }

                if (probe.Length > 2)
                {
                    probe = probe.Substring(0, 2);
                    continue;
                }

                break;
            }

            return null;
        }

        private EntryScope BuildScope(string methodNo, string projectEntryCode, string matchedCode, string fallbackEntryName)
        {
            HashSet<string> pool;
            if (!pools.TryGetValue(PoolKey(methodNo, matchedCode), out pool) || pool.Count == 0)
            {
                return null;
            }

            EntryScope scope = new EntryScope();
            scope.ProjectEntryCode = projectEntryCode;
            scope.MatchedEntryCode = matchedCode;
            string entryName;
            scope.EntryName = entryNames.TryGetValue(matchedCode, out entryName) && !String.IsNullOrEmpty(entryName) ? entryName : (fallbackEntryName ?? "");
            scope.Method = MethodKey;
            scope.MethodNo = methodNo;
            scope.PoolKeys = pool;
            return scope;
        }

        private EntryScope BuildEditableScope(string methodNo, string projectEntryCode, string matchedCode, string fallbackEntryName)
        {
            EntryScope scope = new EntryScope();
            scope.ProjectEntryCode = projectEntryCode;
            scope.MatchedEntryCode = matchedCode;
            string entryName;
            scope.EntryName = entryNames.TryGetValue(matchedCode, out entryName) && !String.IsNullOrEmpty(entryName) ? entryName : (fallbackEntryName ?? "");
            scope.Method = MethodKey;
            scope.MethodNo = methodNo;
            HashSet<string> pool;
            scope.PoolKeys = pools.TryGetValue(PoolKey(methodNo, matchedCode), out pool) ? pool : new HashSet<string>(StringComparer.Ordinal);
            return scope;
        }

        // 用户扶正/采纳了池外定额时补进池子，严格模式不与用户作对；追加 source=user 行持久化
        public void AddUserQuota(EntryScope scope, string targetKind, string code, string name, string unit)
        {
            AddUserQuota(scope, targetKind, code, name, unit, "");
        }

        public void AddUserQuota(EntryScope scope, string targetKind, string code, string name, string unit, string price)
        {
            if (scope == null || String.IsNullOrEmpty(scope.MatchedEntryCode) || String.IsNullOrWhiteSpace(code))
            {
                return;
            }

            AddUserQuota(scope.MethodNo, scope.MatchedEntryCode, scope.EntryName, targetKind, code, name, unit, price);
        }

        public void AddUserQuota(string methodNo, string entryCode, string entryName, string targetKind, string code, string name, string unit)
        {
            AddUserQuota(methodNo, entryCode, entryName, targetKind, code, name, unit, "");
        }

        public void AddUserQuota(string methodNo, string entryCode, string entryName, string targetKind, string code, string name, string unit, string price)
        {
            code = QuotaEntry.NormalizeCode(code);
            if (String.IsNullOrEmpty(entryCode) || String.IsNullOrWhiteSpace(code))
            {
                return;
            }

            methodNo = EffectiveMethodNo(methodNo);
            string kind = String.IsNullOrWhiteSpace(targetKind) ? QuotaEntry.GuessKind(code) : targetKind.Trim().ToLowerInvariant();
            if (!IsAllowedReferencePoolCode(entryName, kind, code, LoadReferenceQuotaCodes(LearningStore.FindDataDir())))
            {
                QuotaRecommendPanel.Log("ChapterLibrary user quota ignored by reference-pool rule. entry=" + entryCode + " code=" + code + " kind=" + kind);
                return;
            }
            kind = ReferencePoolKind(entryName, kind, code);
            string key = ReferencePoolItemKey(kind, code, name, unit, price);
            HashSet<string> pool;
            if (pools.TryGetValue(PoolKey(methodNo, entryCode), out pool) && pool.Contains(key))
            {
                return;
            }

            try
            {
                Dictionary<string, string> record = new Dictionary<string, string>();
                record["record_type"] = "entry_quota";
                record["method"] = MethodKey;
                record["method_no"] = methodNo;
                record["entry_code"] = entryCode;
                record["entry_name"] = entryName ?? "";
                record["target_kind"] = kind;
                record["quota_code"] = code.Trim();
                record["quota_name"] = name ?? "";
                record["quota_unit"] = unit ?? "";
                if (!String.IsNullOrWhiteSpace(price))
                {
                    record["base_price"] = price.Trim();
                }
                record["project_count"] = "0";
                record["source"] = "user";
                record["last_seen"] = DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                string path = Path.Combine(LearningStore.FindDataDir(), "chapter-quota-library.jsonl");
                File.AppendAllText(path, LearningStore.ToJson(record) + Environment.NewLine, Encoding.UTF8);
                AddPoolKey(methodNo, entryCode, kind, code.Trim(), name, unit, price);
                BuildNameIndex();
                RefreshCachedFingerprint();
                QuotaRecommendPanel.Log("ChapterLibrary user quota added. methodNo=" + methodNo + " entry=" + entryCode + " code=" + code);
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("ChapterLibrary user quota append failed: " + ex.Message);
            }
        }

        // 用户从参考池删除定额：从内存池移除并追加 deleted=1 墓碑行（软删除，可被后续 add 覆盖恢复）
        public void RemoveUserQuota(EntryScope scope, string targetKind, string code)
        {
            if (scope == null || String.IsNullOrEmpty(scope.MatchedEntryCode) || String.IsNullOrWhiteSpace(code))
            {
                return;
            }

            RemoveUserQuota(scope.MethodNo, scope.MatchedEntryCode, scope.EntryName, targetKind, code);
        }

        public void RemoveUserQuota(string methodNo, string entryCode, string entryName, string targetKind, string code)
        {
            RemoveUserQuota(methodNo, entryCode, entryName, targetKind, code, "", "", "");
        }

        public void RemoveUserQuota(string methodNo, string entryCode, string entryName, string targetKind, string code, string name, string unit, string price)
        {
            code = QuotaEntry.NormalizeCode(code);
            if (String.IsNullOrEmpty(entryCode) || String.IsNullOrWhiteSpace(code))
            {
                return;
            }

            methodNo = EffectiveMethodNo(methodNo);
            string kind = String.IsNullOrWhiteSpace(targetKind) ? QuotaEntry.GuessKind(code) : targetKind.Trim().ToLowerInvariant();

            try
            {
                Dictionary<string, string> record = new Dictionary<string, string>();
                record["record_type"] = "entry_quota";
                record["method"] = MethodKey;
                record["method_no"] = methodNo;
                record["entry_code"] = entryCode;
                record["entry_name"] = entryName ?? "";
                record["target_kind"] = kind;
                record["quota_code"] = code.Trim();
                record["quota_name"] = name ?? "";
                record["quota_unit"] = unit ?? "";
                if (!String.IsNullOrWhiteSpace(price))
                {
                    record["base_price"] = price.Trim();
                }
                record["project_count"] = "0";
                record["source"] = "user";
                record["deleted"] = "1";
                record["last_seen"] = DateTime.Now.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
                string path = Path.Combine(LearningStore.FindDataDir(), "chapter-quota-library.jsonl");
                File.AppendAllText(path, LearningStore.ToJson(record) + Environment.NewLine, Encoding.UTF8);
                RemovePoolKey(methodNo, entryCode, kind, code.Trim(), name, unit, price);
                BuildNameIndex();
                RefreshCachedFingerprint();
                QuotaRecommendPanel.Log("ChapterLibrary user quota removed. methodNo=" + methodNo + " entry=" + entryCode + " code=" + code);
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("ChapterLibrary user quota remove append failed: " + ex.Message);
            }
        }
    }
}
