using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ChapterQuotaLibrary
{
    internal sealed class MethodSpec
    {
        public string Key;          // "2020" / "2024"
        public string LibraryDb;    // RecoData2020 / RecoData2024
        public string MethodNo;     // 编制办法文号（库内存储的精确值）
        public string SeedDb;       // 种子项目数据库名
        public string SheetName;    // Excel sheet 名
    }

    internal sealed class TemplateEntry
    {
        public string Code;
        public string Name;
        public string Unit;
        public string EntryType;
        public int Level;
    }

    internal sealed class ProjectInfo
    {
        public string DbName;
        public string ProjectName;
        public string NormalizedName;
        public DateTime Created;
        public bool IsSeed;
    }

    internal sealed class PoolCode
    {
        public string Kind;
        public string Code;
        public bool Seed;
        public readonly HashSet<string> Projects = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        public readonly Dictionary<string, int> ObservedNames = new Dictionary<string, int>(StringComparer.Ordinal);
        public readonly Dictionary<string, int> ObservedUnits = new Dictionary<string, int>(StringComparer.Ordinal);
        public string LastSeen = "";
        public int AddOrder;
    }

    internal static class Program
    {
        private const string Server = "192.168.2.13";
        private const string SqlUser = "reco";
        private const string SqlPassword = "Des_Reco_2006";

        private static readonly MethodSpec Spec2020 = new MethodSpec
        {
            Key = "2020",
            LibraryDb = "RecoData2020",
            MethodNo = "30号文",
            SeedDb = "Reco20260511134731660",
            SheetName = "30号文"
        };

        private static readonly MethodSpec Spec2024 = new MethodSpec
        {
            Key = "2024",
            LibraryDb = "RecoData2024",
            MethodNo = "TB 10801—2024",
            SeedDb = "Reco20250506093156577",
            SheetName = "TB 10801—2024"
        };

        private static string toolDir;
        private static string reportDir;
        private static string dataDir;

        private static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            try
            {
                if (args.Length < 1)
                {
                    Usage();
                    return 2;
                }

                string root = FindWorkspaceRoot(AppDomain.CurrentDomain.BaseDirectory);
                toolDir = Path.Combine(root, "tools", "ChapterQuotaLibrary");
                reportDir = Path.Combine(toolDir, "reports");
                dataDir = GetArg(args, "--data-dir");
                if (String.IsNullOrWhiteSpace(dataDir))
                {
                    dataDir = Path.Combine(root, "RecoQuotaData");
                }
                Directory.CreateDirectory(reportDir);
                Directory.CreateDirectory(dataDir);

                string command = args[0].Trim();
                if (String.Equals(command, "BuildLibrary", StringComparison.OrdinalIgnoreCase))
                {
                    int maxPool = GetIntArg(args, "--max-pool", 50);
                    int staleStop = GetIntArg(args, "--stale-stop", 30);
                    int limit = GetIntArg(args, "--limit", 0);
                    foreach (MethodSpec spec in SelectMethods(GetArg(args, "--method")))
                    {
                        BuildLibrary(spec, maxPool, staleStop, limit);
                    }
                    return 0;
                }

                if (String.Equals(command, "ExportTemplate", StringComparison.OrdinalIgnoreCase))
                {
                    string outPath = GetArg(args, "--out");
                    if (String.IsNullOrWhiteSpace(outPath))
                    {
                        outPath = Path.Combine(reportDir, "章节条目模板.xlsx");
                    }
                    ExportTemplate(outPath);
                    return 0;
                }

                if (String.Equals(command, "ImportTrimmed", StringComparison.OrdinalIgnoreCase))
                {
                    string inPath = GetArg(args, "--in");
                    if (String.IsNullOrWhiteSpace(inPath) || !File.Exists(inPath))
                    {
                        Console.Error.WriteLine("ImportTrimmed requires --in <xlsx>.");
                        return 2;
                    }
                    ImportTrimmed(inPath);
                    return 0;
                }

                if (String.Equals(command, "TagMappingBoxes", StringComparison.OrdinalIgnoreCase))
                {
                    string filePath = GetArg(args, "--file");
                    if (String.IsNullOrWhiteSpace(filePath))
                    {
                        filePath = Path.Combine(dataDir, "mapping-boxes.jsonl");
                    }
                    TagMappingBoxes(filePath);
                    return 0;
                }

                if (String.Equals(command, "Stats", StringComparison.OrdinalIgnoreCase))
                {
                    Stats();
                    return 0;
                }

                Usage();
                return 2;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine(ex);
                return 1;
            }
        }

        private static void Usage()
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  ChapterQuotaLibrary.exe BuildLibrary [--method 2020|2024|all] [--max-pool 50] [--stale-stop 30] [--limit N]");
            Console.WriteLine("  ChapterQuotaLibrary.exe ExportTemplate [--out <xlsx>]");
            Console.WriteLine("  ChapterQuotaLibrary.exe ImportTrimmed --in <xlsx>");
            Console.WriteLine("  ChapterQuotaLibrary.exe TagMappingBoxes [--file <mapping-boxes.jsonl>]");
            Console.WriteLine("  ChapterQuotaLibrary.exe Stats");
        }

        private static List<MethodSpec> SelectMethods(string method)
        {
            if (String.Equals(method, "2020", StringComparison.OrdinalIgnoreCase))
            {
                return new List<MethodSpec> { Spec2020 };
            }
            if (String.Equals(method, "2024", StringComparison.OrdinalIgnoreCase))
            {
                return new List<MethodSpec> { Spec2024 };
            }
            return new List<MethodSpec> { Spec2020, Spec2024 };
        }

        private static string GetArg(string[] args, string name)
        {
            for (int i = 0; i < args.Length - 1; i++)
            {
                if (String.Equals(args[i], name, StringComparison.OrdinalIgnoreCase))
                {
                    return args[i + 1];
                }
            }
            return "";
        }

        private static int GetIntArg(string[] args, string name, int fallback)
        {
            int value;
            return Int32.TryParse(GetArg(args, name), NumberStyles.Integer, CultureInfo.InvariantCulture, out value) ? value : fallback;
        }

        private static string FindWorkspaceRoot(string start)
        {
            DirectoryInfo dir = new DirectoryInfo(start);
            while (dir != null)
            {
                if (File.Exists(Path.Combine(dir.FullName, "AGENTS.md")) &&
                    Directory.Exists(Path.Combine(dir.FullName, "tools")))
                {
                    return dir.FullName;
                }
                dir = dir.Parent;
            }
            throw new DirectoryNotFoundException("Could not find workspace root from " + start);
        }

        private static SqlConnection OpenDb(string database)
        {
            string connectionString = "Data Source=" + Server + ",1433;Initial Catalog=" + database
                + ";User ID=" + SqlUser + ";Password=" + SqlPassword
                + ";Connect Timeout=8;Encrypt=False;TrustServerCertificate=True";
            SqlConnection conn = new SqlConnection(connectionString);
            conn.Open();
            return conn;
        }

        private static bool IsSafeProjectDbName(string name)
        {
            return Regex.IsMatch(name ?? "", "^Reco[0-9]{17}$");
        }

        // ===================== 归一化 =====================

        private static string NormalizeQuotaCode(string raw)
        {
            string code = (raw ?? "").Trim();
            if (String.IsNullOrEmpty(code) || code == "-")
            {
                return "";
            }

            code = code.Replace("参", "").Replace("换", "").Replace("借", ""); // 参/换/借
            int star = code.IndexOf('*');
            if (star >= 0)
            {
                code = code.Substring(0, star);
            }
            int times = code.IndexOf('×'); // ×
            if (times >= 0)
            {
                code = code.Substring(0, times);
            }

            return code.Trim();
        }

        private static string CodeKind(string code)
        {
            return code.All(Char.IsDigit) ? "material" : "quota";
        }

        private static string CodeKey(string kind, string code)
        {
            return kind + ":" + (code ?? "").ToUpperInvariant();
        }

        private static string NormalizeProjectName(string name)
        {
            return Regex.Replace((name ?? "").Trim(), "\\s+", "");
        }

        private static int EntryLevel(string code)
        {
            if (String.IsNullOrEmpty(code) || code == "0")
            {
                return 0;
            }
            string[] segments = code.Split('-');
            int level = segments.Length;
            if (segments[0].Length > 2)
            {
                level++; // 首段 4 位（如 0101）视为 章(01)+节(0101) 两级
            }
            return level;
        }

        private static bool IsQuotaInputEntryType(string entryType)
        {
            return entryType == "小计" || entryType == "指标";
        }

        private static string NormalizeEntryName(string name)
        {
            return Regex.Replace((name ?? ""), "[\\s　]+", "").ToLowerInvariant();
        }

        // 名称索引：规范化条目名称 → 小计/指标类型的模板条目（用于识别"复制条目"的来源）
        private static Dictionary<string, List<TemplateEntry>> BuildNameIndex(Dictionary<string, TemplateEntry> template)
        {
            Dictionary<string, List<TemplateEntry>> index = new Dictionary<string, List<TemplateEntry>>(StringComparer.Ordinal);
            foreach (TemplateEntry entry in template.Values)
            {
                if (!IsQuotaInputEntryType(entry.EntryType))
                {
                    continue;
                }
                string key = NormalizeEntryName(entry.Name);
                if (String.IsNullOrEmpty(key))
                {
                    continue;
                }
                List<TemplateEntry> list;
                if (!index.TryGetValue(key, out list))
                {
                    list = new List<TemplateEntry>();
                    index[key] = list;
                }
                list.Add(entry);
            }
            return index;
        }

        // 项目条目编号 → 模板条目编号。
        // 顺序：编号精确命中 → 按名称识别复制来源（同祖先链优先，再全局唯一）→ 逐级前缀上溯。
        private static string MapToTemplate(string code, string entryName, Dictionary<string, TemplateEntry> template, Dictionary<string, List<TemplateEntry>> nameIndex)
        {
            string current = (code ?? "").Trim();
            if (String.IsNullOrEmpty(current))
            {
                return null;
            }

            if (template.ContainsKey(current))
            {
                return current;
            }

            // 编号不在模板里 ⇒ 用户新建/复制的条目：先按名称找复制来源
            string nameKey = NormalizeEntryName(entryName);
            List<TemplateEntry> sameName;
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
                    TemplateEntry scoped = sameName.FirstOrDefault(e => e.Code.StartsWith(withDash, StringComparison.Ordinal) || e.Code == prefix);
                    if (scoped != null)
                    {
                        return scoped.Code;
                    }
                }

                if (sameName.Count == 1)
                {
                    return sameName[0].Code;
                }
            }

            while (!String.IsNullOrEmpty(current) && current != "0")
            {
                if (template.ContainsKey(current))
                {
                    return current;
                }
                int dash2 = current.LastIndexOf('-');
                if (dash2 > 0)
                {
                    current = current.Substring(0, dash2);
                    continue;
                }
                if (current.Length > 2)
                {
                    current = current.Substring(0, 2);
                    continue;
                }
                break;
            }
            return template.ContainsKey(current ?? "") ? current : null;
        }

        // ===================== 数据加载 =====================

        private static Dictionary<string, TemplateEntry> LoadTemplate(SqlConnection conn, MethodSpec spec)
        {
            Dictionary<string, TemplateEntry> result = new Dictionary<string, TemplateEntry>(StringComparer.Ordinal);
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 条目编号, 工程或费用项目名称, 单位, 条目类型 from [" + spec.LibraryDb + "].dbo.章节表 where 编制办法文号 = @m order by 条目编号";
                cmd.Parameters.AddWithValue("@m", spec.MethodNo);
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        TemplateEntry entry = new TemplateEntry();
                        entry.Code = S(reader, 0).Trim();
                        entry.Name = S(reader, 1).Trim();
                        entry.Unit = S(reader, 2).Trim();
                        entry.EntryType = S(reader, 3).Trim();
                        entry.Level = EntryLevel(entry.Code);
                        if (!String.IsNullOrEmpty(entry.Code) && !result.ContainsKey(entry.Code))
                        {
                            result[entry.Code] = entry;
                        }
                    }
                }
            }
            return result;
        }

        private sealed class ResolvedNameUnit
        {
            public string Name;
            public string Unit;
        }

        private static Dictionary<string, ResolvedNameUnit> LoadResolveDict(SqlConnection conn, MethodSpec spec)
        {
            Dictionary<string, ResolvedNameUnit> dict = new Dictionary<string, ResolvedNameUnit>(StringComparer.OrdinalIgnoreCase);
            LoadResolveRows(conn, "select 定额编号, 定额名称, 单位 from [" + spec.LibraryDb + "].dbo.定额库 where isnull(定额编号,'')<>''", "quota", dict);
            LoadResolveRows(conn, "select cast(电算代号 as nvarchar(50)), 材料名称, 单位 from [" + spec.LibraryDb + "].dbo.材料单价库 where isnull(材料名称,'')<>''", "material", dict);
            LoadResolveRows(conn, "select cast(电算代号 as nvarchar(50)), 机械台班名称, '台班' from [" + spec.LibraryDb + "].dbo.台班定额库 where isnull(机械台班名称,'')<>''", "material", dict);
            return dict;
        }

        private static void LoadResolveRows(SqlConnection conn, string sql, string kind, Dictionary<string, ResolvedNameUnit> dict)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = sql;
                cmd.CommandTimeout = 60;
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string code = S(reader, 0).Trim();
                        if (String.IsNullOrEmpty(code))
                        {
                            continue;
                        }
                        string key = CodeKey(kind, code);
                        if (!dict.ContainsKey(key))
                        {
                            ResolvedNameUnit value = new ResolvedNameUnit();
                            value.Name = S(reader, 1).Trim();
                            value.Unit = S(reader, 2).Trim();
                            dict[key] = value;
                        }
                    }
                }
            }
        }

        private static List<ProjectInfo> LoadRegistry(SqlConnection conn, MethodSpec spec)
        {
            List<ProjectInfo> projects = new List<ProjectInfo>();
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 项目编号, 建设项目名称, 创建时间 from [" + spec.LibraryDb + "].dbo.项目信息 where 编制办法文号 = @m";
                cmd.Parameters.AddWithValue("@m", spec.MethodNo);
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        ProjectInfo project = new ProjectInfo();
                        project.DbName = S(reader, 0).Trim();
                        project.ProjectName = S(reader, 1).Trim();
                        project.NormalizedName = NormalizeProjectName(project.ProjectName);
                        DateTime created;
                        if (!DateTime.TryParse(S(reader, 2), out created))
                        {
                            created = DateTime.MinValue;
                        }
                        project.Created = created;
                        project.IsSeed = String.Equals(project.DbName, spec.SeedDb, StringComparison.OrdinalIgnoreCase);
                        if (IsSafeProjectDbName(project.DbName))
                        {
                            projects.Add(project);
                        }
                    }
                }
            }

            if (!projects.Any(p => p.IsSeed))
            {
                Console.WriteLine("[" + spec.Key + "] WARN: seed project " + spec.SeedDb + " not found in registry; scanning it anyway.");
                ProjectInfo seed = new ProjectInfo();
                seed.DbName = spec.SeedDb;
                seed.ProjectName = spec.SeedDb;
                seed.NormalizedName = spec.SeedDb;
                seed.Created = DateTime.MaxValue;
                seed.IsSeed = true;
                projects.Add(seed);
            }

            return projects
                .OrderByDescending(p => p.IsSeed)
                .ThenByDescending(p => p.Created)
                .ThenBy(p => p.DbName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // ===================== BuildLibrary =====================

        private sealed class ScanStateRow
        {
            public string DbName;
            public string Status;
            public int Hits;
            public int NewCodes;
            public string Error;
        }

        private static void BuildLibrary(MethodSpec spec, int maxPool, int staleStop, int limit)
        {
            Console.WriteLine("[" + spec.Key + "] BuildLibrary method=" + spec.MethodNo + " max-pool=" + maxPool + " stale-stop=" + staleStop + (limit > 0 ? " limit=" + limit : ""));
            string statePath = Path.Combine(reportDir, "scan-state-" + spec.Key + ".csv");
            string rawPath = Path.Combine(reportDir, "raw-entry-quotas-" + spec.Key + ".jsonl");

            using (SqlConnection conn = OpenDb(spec.LibraryDb))
            {
                Dictionary<string, TemplateEntry> template = LoadTemplate(conn, spec);
                Dictionary<string, List<TemplateEntry>> nameIndex = BuildNameIndex(template);
                Console.WriteLine("[" + spec.Key + "] template entries: " + template.Count);
                Dictionary<string, ResolvedNameUnit> resolveDict = LoadResolveDict(conn, spec);
                List<ProjectInfo> projects = LoadRegistry(conn, spec);
                Console.WriteLine("[" + spec.Key + "] registry projects: " + projects.Count);
                Dictionary<string, ProjectInfo> projectByDb = projects
                    .GroupBy(p => p.DbName, StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

                // 池：entry -> codeKey -> PoolCode；先重放 raw 文件恢复进度
                Dictionary<string, Dictionary<string, PoolCode>> pools = new Dictionary<string, Dictionary<string, PoolCode>>(StringComparer.Ordinal);
                Dictionary<string, HashSet<string>> entryProjects = new Dictionary<string, HashSet<string>>(StringComparer.Ordinal);
                int addOrder = 0;
                if (File.Exists(rawPath))
                {
                    foreach (string line in File.ReadAllLines(rawPath, Encoding.UTF8))
                    {
                        if (String.IsNullOrWhiteSpace(line))
                        {
                            continue;
                        }
                        Dictionary<string, string> values = FlatJson.Parse(line);
                        ApplyHit(pools, entryProjects, maxPool, ref addOrder,
                            FlatJson.Get(values, "entry_code"),
                            FlatJson.Get(values, "kind"),
                            FlatJson.Get(values, "code"),
                            FlatJson.Get(values, "name"),
                            FlatJson.Get(values, "unit"),
                            FlatJson.Get(values, "project"),
                            FlatJson.Get(values, "created"),
                            FlatJson.Get(values, "seed") == "1");
                    }
                }

                List<ScanStateRow> state = LoadScanState(statePath);
                HashSet<string> scanned = new HashSet<string>(state.Select(s => s.DbName), StringComparer.OrdinalIgnoreCase);
                int staleStreak = 0;
                for (int i = state.Count - 1; i >= 0; i--)
                {
                    if (state[i].Status != "ok")
                    {
                        continue;
                    }
                    ProjectInfo prior;
                    bool priorSeed = projectByDb.TryGetValue(state[i].DbName, out prior) && prior.IsSeed;
                    if (priorSeed)
                    {
                        break;
                    }
                    if (state[i].NewCodes > 0)
                    {
                        break;
                    }
                    staleStreak++;
                }

                int scannedThisRun = 0;
                using (StreamWriter rawWriter = new StreamWriter(rawPath, true, Encoding.UTF8))
                {
                    foreach (ProjectInfo project in projects)
                    {
                        if (scanned.Contains(project.DbName))
                        {
                            continue;
                        }
                        if (limit > 0 && scannedThisRun >= limit)
                        {
                            Console.WriteLine("[" + spec.Key + "] --limit " + limit + " reached, stopping scan.");
                            break;
                        }
                        if (!project.IsSeed && staleStreak >= staleStop)
                        {
                            Console.WriteLine("[" + spec.Key + "] saturation reached: " + staleStreak + " consecutive projects added nothing, stopping scan.");
                            break;
                        }

                        ScanStateRow row = new ScanStateRow();
                        row.DbName = project.DbName;
                        try
                        {
                            int newCodes;
                            int hits = ScanProject(conn, project, template, nameIndex, pools, entryProjects, maxPool, ref addOrder, rawWriter, out newCodes);
                            row.Status = "ok";
                            row.Hits = hits;
                            row.NewCodes = newCodes;
                            row.Error = "";
                            if (!project.IsSeed)
                            {
                                staleStreak = newCodes > 0 ? 0 : staleStreak + 1;
                            }
                            Console.WriteLine("[" + spec.Key + "] " + project.DbName + (project.IsSeed ? " (seed)" : "") + " hits=" + hits + " new=" + newCodes + " stale=" + staleStreak + " | " + project.ProjectName);
                        }
                        catch (Exception ex)
                        {
                            row.Status = "error";
                            row.Hits = 0;
                            row.NewCodes = 0;
                            row.Error = ex.Message.Replace("\r", " ").Replace("\n", " ").Replace("\t", " ");
                            Console.WriteLine("[" + spec.Key + "] " + project.DbName + " ERROR " + row.Error);
                        }

                        state.Add(row);
                        scanned.Add(project.DbName);
                        scannedThisRun++;
                        rawWriter.Flush();
                        AppendScanState(statePath, row);
                    }
                }

                WriteLibraryFile(spec, template, pools, entryProjects, resolveDict, projectByDb, maxPool);
                WriteEntrySummary(spec, template, pools, entryProjects);
                WriteBuildSummary(spec, template, pools, entryProjects, state, maxPool, staleStop);
            }
        }

        private static List<ScanStateRow> LoadScanState(string path)
        {
            List<ScanStateRow> rows = new List<ScanStateRow>();
            if (!File.Exists(path))
            {
                return rows;
            }
            foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
            {
                string[] parts = line.Split('\t');
                if (parts.Length < 4)
                {
                    continue;
                }
                ScanStateRow row = new ScanStateRow();
                row.DbName = parts[0];
                row.Status = parts[1];
                Int32.TryParse(parts[2], out row.Hits);
                Int32.TryParse(parts[3], out row.NewCodes);
                row.Error = parts.Length > 4 ? parts[4] : "";
                rows.Add(row);
            }
            return rows;
        }

        private static void AppendScanState(string path, ScanStateRow row)
        {
            File.AppendAllText(path, row.DbName + "\t" + row.Status + "\t"
                + row.Hits.ToString(CultureInfo.InvariantCulture) + "\t"
                + row.NewCodes.ToString(CultureInfo.InvariantCulture) + "\t"
                + (row.Error ?? "") + Environment.NewLine, Encoding.UTF8);
        }

        private static int ScanProject(SqlConnection conn, ProjectInfo project, Dictionary<string, TemplateEntry> template,
            Dictionary<string, List<TemplateEntry>> nameIndex,
            Dictionary<string, Dictionary<string, PoolCode>> pools, Dictionary<string, HashSet<string>> entryProjects,
            int maxPool, ref int addOrder, StreamWriter rawWriter, out int newCodes)
        {
            newCodes = 0;
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            List<Dictionary<string, string>> projectHits = new List<Dictionary<string, string>>();
            string createdText = project.Created == DateTime.MinValue || project.Created == DateTime.MaxValue
                ? ""
                : project.Created.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select z.条目编号, d.定额编号, d.工程或费用项目名称, d.单位, z.工程或费用项目名称"
                    + " from [" + project.DbName + "].dbo.定额输入 d"
                    + " join [" + project.DbName + "].dbo.章节表 z on d.条目序号 = z.条目序号"
                    + " where isnull(d.定额编号,'') <> '' and ltrim(rtrim(d.定额编号)) <> '-'";
                cmd.CommandTimeout = 120;
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string projectEntry = S(reader, 0).Trim();
                        string normCode = NormalizeQuotaCode(S(reader, 1));
                        if (String.IsNullOrEmpty(normCode) || String.IsNullOrEmpty(projectEntry))
                        {
                            continue;
                        }

                        string entryCode = MapToTemplate(projectEntry, S(reader, 4), template, nameIndex);
                        if (entryCode == null)
                        {
                            continue;
                        }

                        string kind = CodeKind(normCode);
                        string dedupeKey = entryCode + "|" + CodeKey(kind, normCode);
                        if (!seen.Add(dedupeKey))
                        {
                            continue;
                        }

                        Dictionary<string, string> hit = new Dictionary<string, string>();
                        hit["record_type"] = "raw_hit";
                        hit["entry_code"] = entryCode;
                        hit["kind"] = kind;
                        hit["code"] = normCode;
                        hit["raw_code"] = S(reader, 1).Trim();
                        hit["name"] = S(reader, 2).Trim();
                        hit["unit"] = S(reader, 3).Trim();
                        hit["db"] = project.DbName;
                        hit["project"] = project.NormalizedName;
                        hit["created"] = createdText;
                        hit["seed"] = project.IsSeed ? "1" : "0";
                        projectHits.Add(hit);
                    }
                }
            }

            foreach (Dictionary<string, string> hit in projectHits)
            {
                rawWriter.WriteLine(FlatJson.ToJson(hit));
                bool added = ApplyHit(pools, entryProjects, maxPool, ref addOrder,
                    hit["entry_code"], hit["kind"], hit["code"], hit["name"], hit["unit"], hit["project"], hit["created"], project.IsSeed);
                if (added)
                {
                    newCodes++;
                }
            }

            return projectHits.Count;
        }

        private static bool ApplyHit(Dictionary<string, Dictionary<string, PoolCode>> pools, Dictionary<string, HashSet<string>> entryProjects,
            int maxPool, ref int addOrder, string entryCode, string kind, string code, string name, string unit, string project, string created, bool seed)
        {
            if (String.IsNullOrEmpty(entryCode) || String.IsNullOrEmpty(code))
            {
                return false;
            }

            HashSet<string> contributors;
            if (!entryProjects.TryGetValue(entryCode, out contributors))
            {
                contributors = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                entryProjects[entryCode] = contributors;
            }
            if (!String.IsNullOrEmpty(project))
            {
                contributors.Add(project);
            }

            Dictionary<string, PoolCode> pool;
            if (!pools.TryGetValue(entryCode, out pool))
            {
                pool = new Dictionary<string, PoolCode>(StringComparer.OrdinalIgnoreCase);
                pools[entryCode] = pool;
            }

            string key = CodeKey(kind, code);
            PoolCode existing;
            bool added = false;
            if (pool.TryGetValue(key, out existing))
            {
                if (seed && !existing.Seed)
                {
                    existing.Seed = true;
                }
            }
            else
            {
                // 种子定额无条件入池；扫描定额受饱和上限约束
                if (!seed && pool.Count >= maxPool)
                {
                    return false;
                }
                existing = new PoolCode();
                existing.Kind = kind;
                existing.Code = code;
                existing.Seed = seed;
                existing.AddOrder = ++addOrder;
                pool[key] = existing;
                added = true;
            }

            if (!String.IsNullOrEmpty(project))
            {
                existing.Projects.Add(project);
            }
            if (!String.IsNullOrEmpty(name))
            {
                int count;
                existing.ObservedNames.TryGetValue(name, out count);
                existing.ObservedNames[name] = count + 1;
            }
            if (!String.IsNullOrEmpty(unit))
            {
                int count;
                existing.ObservedUnits.TryGetValue(unit, out count);
                existing.ObservedUnits[unit] = count + 1;
            }
            if (String.Compare(created ?? "", existing.LastSeen ?? "", StringComparison.Ordinal) > 0)
            {
                existing.LastSeen = created;
            }
            return added;
        }

        private static void WriteLibraryFile(MethodSpec spec, Dictionary<string, TemplateEntry> template,
            Dictionary<string, Dictionary<string, PoolCode>> pools, Dictionary<string, HashSet<string>> entryProjects,
            Dictionary<string, ResolvedNameUnit> resolveDict, Dictionary<string, ProjectInfo> projectByDb, int maxPool)
        {
            string libraryPath = Path.Combine(dataDir, "chapter-quota-library.jsonl");
            List<string> kept = new List<string>();
            if (File.Exists(libraryPath))
            {
                foreach (string line in File.ReadAllLines(libraryPath, Encoding.UTF8))
                {
                    if (String.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }
                    Dictionary<string, string> values = FlatJson.Parse(line);
                    string method = FlatJson.Get(values, "method");
                    string source = FlatJson.Get(values, "source");
                    // 重建当前 method 的 seed/scan 行；其他 method 的行与所有 user 行保留
                    if (!String.Equals(method, spec.Key, StringComparison.OrdinalIgnoreCase)
                        || String.Equals(source, "user", StringComparison.OrdinalIgnoreCase))
                    {
                        kept.Add(line);
                    }
                }
            }

            List<string> unresolved = new List<string>();
            int written = 0;
            foreach (KeyValuePair<string, Dictionary<string, PoolCode>> pair in pools.OrderBy(p => p.Key, StringComparer.Ordinal))
            {
                TemplateEntry entry;
                if (!template.TryGetValue(pair.Key, out entry))
                {
                    continue;
                }

                List<PoolCode> codes = pair.Value.Values
                    .OrderByDescending(c => c.Seed)
                    .ThenByDescending(c => c.Projects.Count)
                    .ThenByDescending(c => c.LastSeen ?? "", StringComparer.Ordinal)
                    .ThenBy(c => c.AddOrder)
                    .ToList();

                int nonSeedBudget = Math.Max(0, maxPool - codes.Count(c => c.Seed));
                foreach (PoolCode code in codes)
                {
                    if (!code.Seed)
                    {
                        if (nonSeedBudget <= 0)
                        {
                            continue;
                        }
                        nonSeedBudget--;
                    }

                    ResolvedNameUnit resolvedValue;
                    string resolvedName;
                    string resolvedUnit;
                    if (resolveDict.TryGetValue(CodeKey(code.Kind, code.Code), out resolvedValue))
                    {
                        resolvedName = resolvedValue.Name;
                        resolvedUnit = resolvedValue.Unit;
                    }
                    else
                    {
                        resolvedName = MostFrequent(code.ObservedNames);
                        resolvedUnit = MostFrequent(code.ObservedUnits);
                        unresolved.Add(code.Kind + "," + code.Code + "," + CsvEscape(resolvedName) + "," + pair.Key);
                    }

                    Dictionary<string, string> record = new Dictionary<string, string>();
                    record["record_type"] = "entry_quota";
                    record["method"] = spec.Key;
                    record["method_no"] = spec.MethodNo;
                    record["entry_code"] = pair.Key;
                    record["entry_name"] = entry.Name;
                    record["target_kind"] = code.Kind;
                    record["quota_code"] = code.Code;
                    record["quota_name"] = resolvedName;
                    record["quota_unit"] = resolvedUnit;
                    record["project_count"] = code.Projects.Count.ToString(CultureInfo.InvariantCulture);
                    record["source"] = code.Seed ? "seed" : "scan";
                    record["last_seen"] = code.LastSeen ?? "";
                    kept.Add(FlatJson.ToJson(record));
                    written++;
                }
            }

            WriteAllLinesAtomic(libraryPath, kept);
            File.WriteAllLines(Path.Combine(reportDir, "unresolved-codes-" + spec.Key + ".csv"),
                new[] { "kind,code,observed_name,entry_code" }.Concat(unresolved.Distinct()).ToArray(), Encoding.UTF8);
            Console.WriteLine("[" + spec.Key + "] library lines written: " + written + " -> " + libraryPath);
        }

        private static void WriteEntrySummary(MethodSpec spec, Dictionary<string, TemplateEntry> template,
            Dictionary<string, Dictionary<string, PoolCode>> pools, Dictionary<string, HashSet<string>> entryProjects)
        {
            List<string> lines = new List<string>();
            lines.Add("entry_code,pool_size,distinct_projects,seed_codes");
            foreach (KeyValuePair<string, Dictionary<string, PoolCode>> pair in pools.OrderBy(p => p.Key, StringComparer.Ordinal))
            {
                if (!template.ContainsKey(pair.Key))
                {
                    continue;
                }
                HashSet<string> contributors;
                entryProjects.TryGetValue(pair.Key, out contributors);
                lines.Add(pair.Key + "," + pair.Value.Count.ToString(CultureInfo.InvariantCulture)
                    + "," + (contributors == null ? 0 : contributors.Count).ToString(CultureInfo.InvariantCulture)
                    + "," + pair.Value.Values.Count(c => c.Seed).ToString(CultureInfo.InvariantCulture));
            }
            File.WriteAllLines(Path.Combine(reportDir, "entry-summary-" + spec.Key + ".csv"), lines.ToArray(), Encoding.UTF8);
        }

        private static void WriteBuildSummary(MethodSpec spec, Dictionary<string, TemplateEntry> template,
            Dictionary<string, Dictionary<string, PoolCode>> pools, Dictionary<string, HashSet<string>> entryProjects,
            List<ScanStateRow> state, int maxPool, int staleStop)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("BuildLibrary summary [" + spec.Key + "] " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
            sb.AppendLine("method_no: " + spec.MethodNo + "  max-pool: " + maxPool + "  stale-stop: " + staleStop);
            sb.AppendLine("projects scanned ok: " + state.Count(s => s.Status == "ok"));
            sb.AppendLine("projects errored:    " + state.Count(s => s.Status == "error"));
            int entriesWithPool = pools.Count(p => template.ContainsKey(p.Key) && p.Value.Count > 0);
            sb.AppendLine("template entries:    " + template.Count);
            sb.AppendLine("entries with pool:   " + entriesWithPool);
            int[] buckets = new int[4];
            foreach (KeyValuePair<string, Dictionary<string, PoolCode>> pair in pools)
            {
                if (!template.ContainsKey(pair.Key))
                {
                    continue;
                }
                int size = pair.Value.Count;
                if (size >= maxPool) buckets[3]++;
                else if (size >= 25) buckets[2]++;
                else if (size >= 10) buckets[1]++;
                else buckets[0]++;
            }
            sb.AppendLine("pool size 1-9:   " + buckets[0]);
            sb.AppendLine("pool size 10-24: " + buckets[1]);
            sb.AppendLine("pool size 25-49: " + buckets[2]);
            sb.AppendLine("pool size " + maxPool + "+ (saturated): " + buckets[3]);
            File.AppendAllText(Path.Combine(reportDir, "build-summary.txt"), sb.ToString() + Environment.NewLine, Encoding.UTF8);
            Console.Write(sb.ToString());
        }

        private static string MostFrequent(Dictionary<string, int> counts)
        {
            string best = "";
            int bestCount = 0;
            foreach (KeyValuePair<string, int> pair in counts)
            {
                if (pair.Value > bestCount)
                {
                    best = pair.Key;
                    bestCount = pair.Value;
                }
            }
            return best;
        }

        private static string CsvEscape(string value)
        {
            string text = (value ?? "").Replace("\r", " ").Replace("\n", " ");
            if (text.Contains(",") || text.Contains("\""))
            {
                return "\"" + text.Replace("\"", "\"\"") + "\"";
            }
            return text;
        }

        // ===================== ExportTemplate =====================

        private static void ExportTemplate(string outPath)
        {
            // 读已建库的条目统计（如果有），让用户删减时能看到每个条目学到了多少定额
            Dictionary<string, int[]> stats2020 = LoadEntrySummary(Spec2020);
            Dictionary<string, int[]> stats2024 = LoadEntrySummary(Spec2024);

            XSSFWorkbook workbook = new XSSFWorkbook();
            ICellStyle headerStyle = workbook.CreateCellStyle();
            IFont headerFont = workbook.CreateFont();
            headerFont.IsBold = true;
            headerStyle.SetFont(headerFont);
            ICellStyle textStyle = workbook.CreateCellStyle();
            textStyle.DataFormat = workbook.CreateDataFormat().GetFormat("@");

            ISheet readme = workbook.CreateSheet("说明");
            string[] notes = new string[]
            {
                "章节条目模板：请删除不需要的条目所在的整行，保留的行不要修改【条目编号】列。",
                "删减完成后把文件发回，用 ImportTrimmed --in <文件> 导入，推荐定额的条目范围随即生效。",
                "「学习到的定额数」是从已完成项目中学到的该条目下不重复定额数量；",
                "「覆盖项目数」是有多少个项目在该条目下填过定额。",
                "删掉父条目不影响保留的子条目；运行时会自动向上归类到最近的保留条目。",
                "sheet：30号文 = 国铁科法[2017]30号（2020版）；TB 10801—2024 = 2024版。"
            };
            for (int i = 0; i < notes.Length; i++)
            {
                readme.CreateRow(i).CreateCell(0).SetCellValue(notes[i]);
            }
            readme.SetColumnWidth(0, 120 * 256);

            foreach (MethodSpec spec in new[] { Spec2020, Spec2024 })
            {
                Dictionary<string, int[]> stats = spec.Key == "2020" ? stats2020 : stats2024;
                using (SqlConnection conn = OpenDb(spec.LibraryDb))
                {
                    Dictionary<string, TemplateEntry> template = LoadTemplate(conn, spec);
                    ISheet sheet = workbook.CreateSheet(spec.SheetName);
                    IRow header = sheet.CreateRow(0);
                    string[] headers = new string[] { "条目编号", "层级", "工程或费用项目名称", "单位", "条目类型", "学习到的定额数", "覆盖项目数" };
                    for (int i = 0; i < headers.Length; i++)
                    {
                        ICell cell = header.CreateCell(i);
                        cell.SetCellValue(headers[i]);
                        cell.CellStyle = headerStyle;
                    }

                    int rowIndex = 1;
                    foreach (TemplateEntry entry in template.Values.OrderBy(e => e.Code, StringComparer.Ordinal))
                    {
                        if (entry.Code == "0")
                        {
                            continue;
                        }
                        IRow row = sheet.CreateRow(rowIndex++);
                        ICell codeCell = row.CreateCell(0);
                        codeCell.SetCellValue(entry.Code);
                        codeCell.CellStyle = textStyle;
                        row.CreateCell(1).SetCellValue(entry.Level);
                        row.CreateCell(2).SetCellValue(new string('　', Math.Max(0, entry.Level - 1)) + entry.Name);
                        row.CreateCell(3).SetCellValue(entry.Unit);
                        row.CreateCell(4).SetCellValue(entry.EntryType);
                        int[] counts;
                        if (stats.TryGetValue(entry.Code, out counts))
                        {
                            row.CreateCell(5).SetCellValue(counts[0]);
                            row.CreateCell(6).SetCellValue(counts[1]);
                        }
                    }

                    sheet.CreateFreezePane(0, 1);
                    sheet.SetColumnWidth(0, 28 * 256);
                    sheet.SetColumnWidth(1, 6 * 256);
                    sheet.SetColumnWidth(2, 60 * 256);
                    sheet.SetColumnWidth(3, 10 * 256);
                    sheet.SetColumnWidth(4, 10 * 256);
                    sheet.SetColumnWidth(5, 16 * 256);
                    sheet.SetColumnWidth(6, 14 * 256);
                    Console.WriteLine("[" + spec.Key + "] sheet " + spec.SheetName + " rows: " + (rowIndex - 1));
                }
            }

            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outPath)));
            using (FileStream stream = new FileStream(outPath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream);
            }
            Console.WriteLine("Template written: " + Path.GetFullPath(outPath));
        }

        private static Dictionary<string, int[]> LoadEntrySummary(MethodSpec spec)
        {
            Dictionary<string, int[]> stats = new Dictionary<string, int[]>(StringComparer.Ordinal);
            string path = Path.Combine(reportDir, "entry-summary-" + spec.Key + ".csv");
            if (!File.Exists(path))
            {
                return stats;
            }
            foreach (string line in File.ReadAllLines(path, Encoding.UTF8).Skip(1))
            {
                string[] parts = line.Split(',');
                if (parts.Length < 3)
                {
                    continue;
                }
                int poolSize;
                int projects;
                Int32.TryParse(parts[1], out poolSize);
                Int32.TryParse(parts[2], out projects);
                stats[parts[0]] = new int[] { poolSize, projects };
            }
            return stats;
        }

        // ===================== ImportTrimmed =====================

        private static void ImportTrimmed(string inPath)
        {
            List<string> warnings = new List<string>();
            List<string> entryLines = new List<string>();
            HashSet<string> keptKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            IWorkbook workbook;
            using (FileStream stream = new FileStream(inPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                workbook = WorkbookFactory.Create(stream);
            }

            foreach (MethodSpec spec in new[] { Spec2020, Spec2024 })
            {
                ISheet sheet = FindSheet(workbook, spec.SheetName);
                if (sheet == null)
                {
                    Console.WriteLine("[" + spec.Key + "] sheet not found in workbook, skipped.");
                    continue;
                }

                using (SqlConnection conn = OpenDb(spec.LibraryDb))
                {
                    Dictionary<string, TemplateEntry> template = LoadTemplate(conn, spec);
                    int codeColumn = FindHeaderColumn(sheet, "条目编号");
                    if (codeColumn < 0)
                    {
                        warnings.Add(spec.Key + ",,sheet missing 条目编号 header");
                        continue;
                    }

                    int kept = 0;
                    for (int rowIndex = 1; rowIndex <= sheet.LastRowNum; rowIndex++)
                    {
                        IRow row = sheet.GetRow(rowIndex);
                        if (row == null)
                        {
                            continue;
                        }
                        string code = CellText(row.GetCell(codeColumn)).Trim();
                        if (String.IsNullOrEmpty(code))
                        {
                            continue;
                        }

                        TemplateEntry entry;
                        if (!template.TryGetValue(code, out entry))
                        {
                            warnings.Add(spec.Key + "," + code + ",unknown entry code");
                            continue;
                        }

                        if (!keptKeys.Add(spec.Key + ":" + code))
                        {
                            continue;
                        }
                        Dictionary<string, string> record = new Dictionary<string, string>();
                        record["record_type"] = "chapter_entry";
                        record["method"] = spec.Key;
                        record["method_no"] = spec.MethodNo;
                        record["entry_code"] = entry.Code;
                        record["entry_name"] = entry.Name;
                        record["unit"] = entry.Unit;
                        record["entry_type"] = entry.EntryType;
                        record["level"] = entry.Level.ToString(CultureInfo.InvariantCulture);
                        entryLines.Add(FlatJson.ToJson(record));
                        kept++;
                    }
                    Console.WriteLine("[" + spec.Key + "] kept entries: " + kept);
                }
            }

            if (entryLines.Count == 0)
            {
                Console.Error.WriteLine("No entries imported; chapter-entries.jsonl NOT written.");
            }
            else
            {
                string entriesPath = Path.Combine(dataDir, "chapter-entries.jsonl");
                WriteAllLinesAtomic(entriesPath, entryLines);
                Console.WriteLine("Entries written: " + entriesPath);

                // 按保留条目修剪定额池
                string libraryPath = Path.Combine(dataDir, "chapter-quota-library.jsonl");
                if (File.Exists(libraryPath))
                {
                    List<string> keptLibrary = new List<string>();
                    int dropped = 0;
                    foreach (string line in File.ReadAllLines(libraryPath, Encoding.UTF8))
                    {
                        if (String.IsNullOrWhiteSpace(line))
                        {
                            continue;
                        }
                        Dictionary<string, string> values = FlatJson.Parse(line);
                        string key = FlatJson.Get(values, "method") + ":" + FlatJson.Get(values, "entry_code");
                        if (keptKeys.Contains(key))
                        {
                            keptLibrary.Add(line);
                        }
                        else
                        {
                            dropped++;
                        }
                    }
                    WriteAllLinesAtomic(libraryPath, keptLibrary);
                    Console.WriteLine("Library pruned: kept " + keptLibrary.Count + ", dropped " + dropped);
                }
            }

            string warningsPath = Path.Combine(reportDir, "import-warnings.csv");
            File.WriteAllLines(warningsPath, new[] { "method,entry_code,message" }.Concat(warnings).ToArray(), Encoding.UTF8);
            if (warnings.Count > 0)
            {
                Console.WriteLine("Warnings: " + warnings.Count + " -> " + warningsPath);
            }
        }

        private static ISheet FindSheet(IWorkbook workbook, string name)
        {
            string normalized = NormalizeDashes(name);
            for (int i = 0; i < workbook.NumberOfSheets; i++)
            {
                if (String.Equals(NormalizeDashes(workbook.GetSheetName(i)), normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return workbook.GetSheetAt(i);
                }
            }
            return null;
        }

        private static string NormalizeDashes(string text)
        {
            return (text ?? "").Replace('—', '-').Replace('–', '-').Replace('－', '-').Trim();
        }

        private static int FindHeaderColumn(ISheet sheet, string headerText)
        {
            IRow header = sheet.GetRow(0);
            if (header == null)
            {
                return -1;
            }
            for (int i = 0; i < header.LastCellNum; i++)
            {
                if (String.Equals(CellText(header.GetCell(i)).Trim(), headerText, StringComparison.Ordinal))
                {
                    return i;
                }
            }
            return -1;
        }

        private static string CellText(ICell cell)
        {
            if (cell == null)
            {
                return "";
            }
            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue ?? "";
                case CellType.Numeric:
                    return cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                case CellType.Formula:
                    try
                    {
                        return cell.StringCellValue ?? "";
                    }
                    catch (Exception)
                    {
                        try
                        {
                            return cell.NumericCellValue.ToString(CultureInfo.InvariantCulture);
                        }
                        catch (Exception)
                        {
                            return "";
                        }
                    }
                default:
                    return "";
            }
        }

        // ===================== TagMappingBoxes =====================

        private const string MappingBoxesMutexName = "RecoQuotaData.mapping-boxes.lock";
        private const int MaxEntryTagsPerBox = 50;

        private static void TagMappingBoxes(string filePath)
        {
            if (!File.Exists(filePath))
            {
                Console.Error.WriteLine("mapping-boxes file not found: " + filePath);
                return;
            }

            // 条目池：methodKey:entry_code -> codeKey 集合；若已有删减后的条目表则只对保留条目打标
            string libraryPath = Path.Combine(dataDir, "chapter-quota-library.jsonl");
            if (!File.Exists(libraryPath))
            {
                Console.Error.WriteLine("chapter-quota-library.jsonl not found, run BuildLibrary first.");
                return;
            }

            // 只对"小计/指标"条目打标——只有这两类条目有定额输入框，其他类型不参与定额匹配
            HashSet<string> keptKeys = null;
            string entriesPath = Path.Combine(dataDir, "chapter-entries.jsonl");
            if (File.Exists(entriesPath))
            {
                keptKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (string line in File.ReadAllLines(entriesPath, Encoding.UTF8))
                {
                    if (String.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }
                    Dictionary<string, string> values = FlatJson.Parse(line);
                    if (!IsQuotaInputEntryType(FlatJson.Get(values, "entry_type").Trim()))
                    {
                        continue;
                    }
                    keptKeys.Add(FlatJson.Get(values, "method") + ":" + FlatJson.Get(values, "entry_code"));
                }
            }

            Dictionary<string, HashSet<string>> entryPools = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (string line in File.ReadAllLines(libraryPath, Encoding.UTF8))
            {
                if (String.IsNullOrWhiteSpace(line))
                {
                    continue;
                }
                Dictionary<string, string> values = FlatJson.Parse(line);
                string entryKey = FlatJson.Get(values, "method") + ":" + FlatJson.Get(values, "entry_code");
                if (keptKeys != null && !keptKeys.Contains(entryKey))
                {
                    continue;
                }
                HashSet<string> pool;
                if (!entryPools.TryGetValue(entryKey, out pool))
                {
                    pool = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    entryPools[entryKey] = pool;
                }
                pool.Add(CodeKey(FlatJson.Get(values, "target_kind"), FlatJson.Get(values, "quota_code")));
            }

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

                List<Dictionary<string, string>> lines = new List<Dictionary<string, string>>();
                Dictionary<string, HashSet<string>> boxTargets = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
                Dictionary<string, HashSet<string>> boxTags = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
                foreach (string line in File.ReadAllLines(filePath, Encoding.UTF8))
                {
                    if (String.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }
                    Dictionary<string, string> values = FlatJson.Parse(line);
                    lines.Add(values);
                    string boxId = FlatJson.Get(values, "box_id");
                    if (String.IsNullOrWhiteSpace(boxId))
                    {
                        continue;
                    }
                    HashSet<string> targets;
                    if (!boxTargets.TryGetValue(boxId, out targets))
                    {
                        targets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        boxTargets[boxId] = targets;
                    }
                    string code = NormalizeQuotaCode(FlatJson.Get(values, "target_code"));
                    if (!String.IsNullOrEmpty(code))
                    {
                        string kind = FlatJson.Get(values, "target_kind");
                        if (String.IsNullOrWhiteSpace(kind))
                        {
                            kind = CodeKind(code);
                        }
                        targets.Add(CodeKey(kind.ToLowerInvariant(), code));
                    }
                    HashSet<string> tags;
                    if (!boxTags.TryGetValue(boxId, out tags))
                    {
                        tags = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        boxTags[boxId] = tags;
                    }
                    foreach (string tag in FlatJson.Get(values, "entry_codes").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                    {
                        tags.Add(tag.Trim());
                    }
                }

                int tagged = 0;
                foreach (KeyValuePair<string, HashSet<string>> pair in boxTargets)
                {
                    // 全部 quota 类目标都落在条目池内才打该条目标签；纯材料框按材料码判断
                    HashSet<string> quotaTargets = new HashSet<string>(pair.Value.Where(t => t.StartsWith("quota:", StringComparison.OrdinalIgnoreCase)), StringComparer.OrdinalIgnoreCase);
                    HashSet<string> check = quotaTargets.Count > 0 ? quotaTargets : pair.Value;
                    if (check.Count == 0)
                    {
                        continue;
                    }
                    HashSet<string> tags = boxTags[pair.Key];
                    foreach (KeyValuePair<string, HashSet<string>> poolPair in entryPools)
                    {
                        if (tags.Count >= MaxEntryTagsPerBox)
                        {
                            break;
                        }
                        if (check.All(t => poolPair.Value.Contains(t)))
                        {
                            if (tags.Add(poolPair.Key))
                            {
                                tagged++;
                            }
                        }
                    }
                }

                List<string> output = new List<string>();
                foreach (Dictionary<string, string> values in lines)
                {
                    string boxId = FlatJson.Get(values, "box_id");
                    HashSet<string> tags;
                    if (!String.IsNullOrWhiteSpace(boxId) && boxTags.TryGetValue(boxId, out tags) && tags.Count > 0)
                    {
                        values["entry_codes"] = String.Join(",", tags.OrderBy(t => t, StringComparer.OrdinalIgnoreCase).ToArray());
                    }
                    output.Add(FlatJson.ToJson(values));
                }

                WriteAllLinesAtomic(filePath, output);
                Console.WriteLine("Boxes: " + boxTargets.Count + ", new entry tags added: " + tagged + " -> " + filePath);
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

        // ===================== Stats =====================

        private static void Stats()
        {
            string[] files = new string[] { "chapter-quota-library.jsonl", "chapter-entries.jsonl", "mapping-boxes.jsonl" };
            foreach (string name in files)
            {
                string path = Path.Combine(dataDir, name);
                if (File.Exists(path))
                {
                    Console.WriteLine(name + ": " + File.ReadAllLines(path, Encoding.UTF8).Count(l => !String.IsNullOrWhiteSpace(l)) + " lines");
                }
                else
                {
                    Console.WriteLine(name + ": (missing)");
                }
            }
        }

        // ===================== 公共 =====================

        private static void WriteAllLinesAtomic(string path, List<string> lines)
        {
            string temp = path + ".tmp";
            File.WriteAllLines(temp, lines.ToArray(), Encoding.UTF8);
            if (File.Exists(path))
            {
                File.Delete(path);
            }
            File.Move(temp, path);
        }

        private static string S(SqlDataReader reader, int index)
        {
            return reader.IsDBNull(index) ? "" : Convert.ToString(reader.GetValue(index), CultureInfo.InvariantCulture) ?? "";
        }
    }

    // 与 RecoQuotaRecommend 的 LearningStore.ParseFlatJson/ToJson 保持同一格式（全字符串平面 JSON）
    internal static class FlatJson
    {
        public static string Get(Dictionary<string, string> values, string key)
        {
            string value;
            return values.TryGetValue(key, out value) ? value : "";
        }

        public static string ToJson(Dictionary<string, string> values)
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
                builder.Append('"').Append(Escape(pair.Key)).Append('"').Append(':')
                    .Append('"').Append(Escape(pair.Value)).Append('"');
            }
            builder.Append('}');
            return builder.ToString();
        }

        private static string Escape(string value)
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

        public static Dictionary<string, string> Parse(string line)
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

                string key = ReadString(line, ref index);
                SkipWhitespace(line, ref index);
                if (index < line.Length && line[index] == ':')
                {
                    index++;
                }
                SkipWhitespace(line, ref index);
                string value = ReadString(line, ref index);
                result[key] = value;
                SkipWhitespace(line, ref index);
                if (index < line.Length && line[index] == ',')
                {
                    index++;
                }
            }

            return result;
        }

        private static string ReadString(string text, ref int index)
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
