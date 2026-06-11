using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

namespace Migrate2020EstimateTo2024
{
    internal static class Program
    {
        private const string Server = "192.168.2.13";
        private const string SourceDb = "RecoData2020";
        private const string TargetDb = "RecoData2024";
        private const string SqlUser = "reco";
        private const string SqlPassword = "Des_Reco_2006";
        private const string ConfirmText = "APPLY RecoData2020 estimate-to-2024 migration";

        private static readonly HashSet<string> SourceCategories = new HashSet<string>(StringComparer.Ordinal)
        {
            "概算定额",
            "估算定额"
        };

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

                string command = args[0].Trim();
                string confirmPath = GetArg(args, "--confirm");
                string root = FindWorkspaceRoot(AppDomain.CurrentDomain.BaseDirectory);
                string softwareDir = Find2024SoftwareDir(root);
                string exePath = Path.Combine(softwareDir, "ReJJGSNet2024.exe");
                if (!File.Exists(exePath))
                {
                    exePath = Path.Combine(softwareDir, "ReJJQDNet2024.exe");
                }

                string reportDir = Path.Combine(root, "tools", "Migrate2020EstimateTo2024", "reports");
                Directory.CreateDirectory(reportDir);

                if (String.Equals(command, "Precheck", StringComparison.OrdinalIgnoreCase))
                {
                    MigrationPlan plan = BuildPlan(exePath, reportDir);
                    WriteReports(plan, reportDir);
                    PrintSummary(plan);
                    Console.WriteLine("Reports: " + reportDir);
                    return plan.HasBlockingConflicts ? 4 : 0;
                }

                if (String.Equals(command, "Apply", StringComparison.OrdinalIgnoreCase))
                {
                    RequireConfirm(confirmPath);
                    MigrationPlan plan = BuildPlan(exePath, reportDir);
                    WriteReports(plan, reportDir);
                    if (plan.HasBlockingConflicts)
                    {
                        PrintSummary(plan);
                        Console.Error.WriteLine("Apply aborted: blocking conflicts found. See conflicts.csv.");
                        return 5;
                    }

                    Apply(plan);
                    Console.WriteLine("Apply completed.");
                    Verify(reportDir);
                    return 0;
                }

                if (String.Equals(command, "Verify", StringComparison.OrdinalIgnoreCase))
                {
                    Verify(reportDir);
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
            Console.WriteLine("  Migrate2020EstimateTo2024.exe Precheck");
            Console.WriteLine("  Migrate2020EstimateTo2024.exe Apply --confirm <confirm-file>");
            Console.WriteLine("  Migrate2020EstimateTo2024.exe Verify");
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

        private static void RequireConfirm(string confirmPath)
        {
            if (String.IsNullOrWhiteSpace(confirmPath) || !File.Exists(confirmPath))
            {
                throw new InvalidOperationException("Apply requires --confirm <file>.");
            }

            string text = File.ReadAllText(confirmPath, Encoding.UTF8).Trim();
            if (!String.Equals(text, ConfirmText, StringComparison.Ordinal))
            {
                throw new InvalidOperationException("Confirm file content mismatch. Expected: " + ConfirmText);
            }
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

        private static string Find2024SoftwareDir(string root)
        {
            string direct = Path.Combine(root, "2024铁路工程云计价系统网络版V1.0", "铁路工程云计价系统网络版V1.0");
            if (Directory.Exists(direct))
            {
                return direct;
            }

            foreach (string dir in Directory.GetDirectories(root, "*", SearchOption.AllDirectories))
            {
                if (File.Exists(Path.Combine(dir, "ReJJGSNet2024.exe")) ||
                    File.Exists(Path.Combine(dir, "ReJJQDNet2024.exe")))
                {
                    return dir;
                }
            }

            throw new DirectoryNotFoundException("Could not find 2024 software directory.");
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

        private static MigrationPlan BuildPlan(string exePath, string reportDir)
        {
            Crypto crypto = new Crypto(exePath);
            if (!crypto.CanRoundTripOld())
            {
                throw new InvalidOperationException("Security.EncryptoOld/DecryptoOld roundtrip failed.");
            }

            using (SqlConnection source = OpenDb(SourceDb))
            using (SqlConnection target = OpenDb(TargetDb))
            {
                MigrationPlan plan = new MigrationPlan();
                plan.BuiltAt = DateTime.Now;
                plan.ReportDir = reportDir;
                plan.SourceIndexes = LoadSourceIndexes(source);
                plan.SourceSections = LoadSourceSections(source, plan.SourceIndexes);
                plan.SourceQuotas = LoadSourceQuotas(source, plan.SourceIndexes);
                plan.TargetBooks = LoadTargetBookKeys(target);
                plan.TargetQuotaKeys = LoadTargetQuotaKeys(target);
                plan.TargetSectionKeys = LoadTargetSectionKeys(target);
                plan.TargetMaterials = LoadTargetMaterials(target);
                plan.TargetMachines = LoadTargetMachines(target);
                plan.SourceMaterials = LoadSourceMaterials(source);
                plan.SourceMachines = LoadSourceMachines(source);
                LoadRanges(target, plan);
                DetectConflicts(plan);
                BuildResourceMap(plan, crypto);
                BuildMigratedConsumptions(plan, crypto);
                return plan;
            }
        }

        private static List<BookIndex> LoadSourceIndexes(SqlConnection conn)
        {
            List<BookIndex> rows = new List<BookIndex>();
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 书号,head,内容,字头,互斥表,分类,专业名称,顺号,现行定额,convert(nvarchar(max),说明) from 定额库索引 where 分类 in ('概算定额','估算定额') order by 分类,顺号";
                using (SqlDataReader r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        rows.Add(new BookIndex
                        {
                            Book = S(r, 0),
                            Head = S(r, 1),
                            Content = S(r, 2),
                            Prefix = S(r, 3),
                            ExclusiveTable = S(r, 4),
                            Category = S(r, 5),
                            Specialty = S(r, 6),
                            Order = I(r, 7),
                            IsCurrent = B(r, 8),
                            Note = S(r, 9)
                        });
                    }
                }
            }

            return rows;
        }

        private static List<SectionRow> LoadSourceSections(SqlConnection conn, List<BookIndex> indexes)
        {
            HashSet<string> books = new HashSet<string>(StringComparer.Ordinal);
            foreach (BookIndex index in indexes)
            {
                books.Add(index.Book);
            }

            List<SectionRow> rows = new List<SectionRow>();
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 书号,节号,节名称,convert(nvarchar(max),描述),起始号,终止号 from 定额节索引 order by 书号,节号";
                using (SqlDataReader r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        string book = S(r, 0);
                        if (!books.Contains(book))
                        {
                            continue;
                        }

                        rows.Add(new SectionRow
                        {
                            Book = book,
                            SectionNo = S(r, 1),
                            SectionName = S(r, 2),
                            Description = S(r, 3),
                            StartNo = S(r, 4),
                            EndNo = S(r, 5)
                        });
                    }
                }
            }

            return rows;
        }

        private static List<QuotaRow> LoadSourceQuotas(SqlConnection conn, List<BookIndex> indexes)
        {
            HashSet<string> books = new HashSet<string>(StringComparer.Ordinal);
            Dictionary<string, string> categories = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (BookIndex index in indexes)
            {
                books.Add(index.Book);
                categories[index.Book] = index.Category;
            }

            List<QuotaRow> rows = new List<QuotaRow>();
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandTimeout = 120;
                cmd.CommandText = "select 书号,定额编号,定额名称,单位,convert(nvarchar(max),消耗),convert(nvarchar(max),工作内容),基本定额,基价,工费,料费,机费,单重,节号,流水号,LOCK,使用,条目序号,导入时间 from 定额库 order by 书号,流水号,定额编号";
                using (SqlDataReader r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        string book = S(r, 0);
                        if (!books.Contains(book))
                        {
                            continue;
                        }

                        rows.Add(new QuotaRow
                        {
                            Book = book,
                            Code = S(r, 1),
                            Name = S(r, 2),
                            Unit = S(r, 3),
                            ConsumeEncrypted = S(r, 4),
                            WorkContent = S(r, 5),
                            BaseQuota = S(r, 6),
                            BasePrice = D(r, 7),
                            LaborFee = D(r, 8),
                            MaterialFee = D(r, 9),
                            MachineFee = D(r, 10),
                            UnitWeight = D(r, 11),
                            SectionNo = S(r, 12),
                            SortOrder = I(r, 13),
                            Locked = B(r, 14),
                            Use = B(r, 15),
                            ItemSequence = I(r, 16),
                            ImportedAt = r.IsDBNull(17) ? (DateTime?)null : r.GetDateTime(17),
                            Category = categories[book]
                        });
                    }
                }
            }

            return rows;
        }

        private static Dictionary<int, Resource> LoadSourceMaterials(SqlConnection conn)
        {
            Dictionary<int, Resource> rows = new Dictionary<int, Resource>();
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 文号,电算代号,材料名称,单位,基期单价,编制期价,编制期含税价,单重,换算系数,汇总标志,主材标志,材料运输类别,LOCK,甲供方式,导入时间 from 材料单价库";
                using (SqlDataReader r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        int code = Convert.ToInt32(r.GetValue(1), CultureInfo.InvariantCulture);
                        rows[code] = new Resource
                        {
                            Kind = ResourceKind.Material,
                            DocNo = S(r, 0),
                            OldCode = code,
                            NewCode = code,
                            Name = S(r, 2),
                            Unit = S(r, 3),
                            BasePrice = D(r, 4),
                            CurrentPrice = D(r, 5),
                            CurrentTaxPrice = D(r, 6),
                            UnitWeight = D(r, 7),
                            ConvertFactor = D(r, 8),
                            SummaryFlag = S(r, 9),
                            MainMaterialFlag = S(r, 10),
                            TransportCategory = S(r, 11),
                            Locked = B(r, 12),
                            SupplyMode = S(r, 13),
                            ImportedAt = r.IsDBNull(14) ? (DateTime?)null : r.GetDateTime(14)
                        };
                    }
                }
            }

            return rows;
        }

        private static Dictionary<int, Resource> LoadSourceMachines(SqlConnection conn)
        {
            Dictionary<int, Resource> rows = new Dictionary<int, Resource>();
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 文号,电算代号,机械台班名称,折旧费,检修费,维护费,安装拆卸费,人工,汽油,柴油,煤,电,水,汇总标志,基价,接触网封锁线路标志,养路费系数,其他费用,LOCK,导入时间 from 台班定额库";
                using (SqlDataReader r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        int code = Convert.ToInt32(r.GetValue(1), CultureInfo.InvariantCulture);
                        rows[code] = new Resource
                        {
                            Kind = ResourceKind.Machine,
                            DocNo = S(r, 0),
                            OldCode = code,
                            NewCode = code,
                            Name = S(r, 2),
                            Unit = "台班",
                            DepreciationFee = D(r, 3),
                            RepairFee = D(r, 4),
                            MaintenanceFee = D(r, 5),
                            InstallUninstallFee = D(r, 6),
                            Labor = D(r, 7),
                            Gasoline = D(r, 8),
                            Diesel = D(r, 9),
                            Coal = D(r, 10),
                            Electricity = D(r, 11),
                            Water = D(r, 12),
                            NaturalGas = 0,
                            SummaryFlag = S(r, 13),
                            BasePrice = D(r, 14),
                            ContactLineBlock = B(r, 15),
                            RoadFeeFactor = D(r, 16),
                            OtherFee = D(r, 17),
                            Locked = B(r, 18),
                            ImportedAt = r.IsDBNull(19) ? (DateTime?)null : r.GetDateTime(19)
                        };
                    }
                }
            }

            return rows;
        }

        private static Dictionary<string, List<Resource>> LoadTargetMaterials(SqlConnection conn)
        {
            Dictionary<string, List<Resource>> rows = new Dictionary<string, List<Resource>>(StringComparer.Ordinal);
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 文号,电算代号,材料名称,单位,基期单价,编制期价,编制期含税价,单重,换算系数,汇总标志,主材标志,材料运输类别,LOCK,甲供方式,导入时间 from 材料单价库";
                using (SqlDataReader r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        Resource res = new Resource
                        {
                            Kind = ResourceKind.Material,
                            DocNo = S(r, 0),
                            NewCode = Convert.ToInt32(r.GetValue(1), CultureInfo.InvariantCulture),
                            Name = S(r, 2),
                            Unit = S(r, 3),
                            BasePrice = D(r, 4),
                            CurrentPrice = D(r, 5),
                            CurrentTaxPrice = D(r, 6),
                            UnitWeight = D(r, 7),
                            ConvertFactor = D(r, 8),
                            SummaryFlag = S(r, 9),
                            MainMaterialFlag = S(r, 10),
                            TransportCategory = S(r, 11),
                            Locked = B(r, 12),
                            SupplyMode = S(r, 13),
                            ImportedAt = r.IsDBNull(14) ? (DateTime?)null : r.GetDateTime(14)
                        };
                        AddByName(rows, res);
                    }
                }
            }

            return rows;
        }

        private static Dictionary<string, List<Resource>> LoadTargetMachines(SqlConnection conn)
        {
            Dictionary<string, List<Resource>> rows = new Dictionary<string, List<Resource>>(StringComparer.Ordinal);
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 文号,电算代号,机械台班名称,折旧费,检修费,维护费,安装拆卸费,人工,汽油,柴油,煤,电,水,天然气,汇总标志,基价,接触网封锁线路标志,养路费系数,其他费用,LOCK,导入时间 from 台班定额库";
                using (SqlDataReader r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        Resource res = new Resource
                        {
                            Kind = ResourceKind.Machine,
                            DocNo = S(r, 0),
                            NewCode = Convert.ToInt32(r.GetValue(1), CultureInfo.InvariantCulture),
                            Name = S(r, 2),
                            Unit = "台班",
                            DepreciationFee = D(r, 3),
                            RepairFee = D(r, 4),
                            MaintenanceFee = D(r, 5),
                            InstallUninstallFee = D(r, 6),
                            Labor = D(r, 7),
                            Gasoline = D(r, 8),
                            Diesel = D(r, 9),
                            Coal = D(r, 10),
                            Electricity = D(r, 11),
                            Water = D(r, 12),
                            NaturalGas = D(r, 13),
                            SummaryFlag = S(r, 14),
                            BasePrice = D(r, 15),
                            ContactLineBlock = B(r, 16),
                            RoadFeeFactor = D(r, 17),
                            OtherFee = D(r, 18),
                            Locked = B(r, 19),
                            ImportedAt = r.IsDBNull(20) ? (DateTime?)null : r.GetDateTime(20)
                        };
                        AddByName(rows, res);
                    }
                }
            }

            return rows;
        }

        private static void AddByName(Dictionary<string, List<Resource>> map, Resource res)
        {
            List<Resource> list;
            if (!map.TryGetValue(res.Name, out list))
            {
                list = new List<Resource>();
                map.Add(res.Name, list);
            }

            list.Add(res);
        }

        private static HashSet<string> LoadTargetBookKeys(SqlConnection conn)
        {
            HashSet<string> rows = new HashSet<string>(StringComparer.Ordinal);
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 书号 from 定额库索引";
                using (SqlDataReader r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        rows.Add(S(r, 0));
                    }
                }
            }
            return rows;
        }

        private static HashSet<string> LoadTargetQuotaKeys(SqlConnection conn)
        {
            HashSet<string> rows = new HashSet<string>(StringComparer.Ordinal);
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 书号,定额编号 from 定额库";
                using (SqlDataReader r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        rows.Add(S(r, 0) + "\t" + S(r, 1));
                    }
                }
            }
            return rows;
        }

        private static HashSet<string> LoadTargetSectionKeys(SqlConnection conn)
        {
            HashSet<string> rows = new HashSet<string>(StringComparer.Ordinal);
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 书号,节号 from 定额节索引";
                using (SqlDataReader r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        rows.Add(S(r, 0) + "\t" + S(r, 1));
                    }
                }
            }
            return rows;
        }

        private static void LoadRanges(SqlConnection conn, MigrationPlan plan)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select isnull(min(材料起始号),0),isnull(max(材料终止号),0),isnull(min(机械起始号),0),isnull(max(机械终止号),0) from 补充电算代号分配表";
                using (SqlDataReader r = cmd.ExecuteReader())
                {
                    if (r.Read())
                    {
                        plan.MaterialRangeStart = Convert.ToInt64(r.GetValue(0), CultureInfo.InvariantCulture);
                        plan.MaterialRangeEnd = Convert.ToInt64(r.GetValue(1), CultureInfo.InvariantCulture);
                        plan.MachineRangeStart = Convert.ToInt64(r.GetValue(2), CultureInfo.InvariantCulture);
                        plan.MachineRangeEnd = Convert.ToInt64(r.GetValue(3), CultureInfo.InvariantCulture);
                    }
                }
            }

            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select isnull(max(电算代号),0) from 材料单价库";
                plan.NextMaterialCode = Math.Max(plan.MaterialRangeStart, Convert.ToInt64(cmd.ExecuteScalar(), CultureInfo.InvariantCulture) + 1);
            }

            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select isnull(max(电算代号),0) from 台班定额库";
                plan.NextMachineCode = Math.Max(plan.MachineRangeStart, Convert.ToInt64(cmd.ExecuteScalar(), CultureInfo.InvariantCulture) + 1);
            }
        }

        private static void DetectConflicts(MigrationPlan plan)
        {
            foreach (BookIndex index in plan.SourceIndexes)
            {
                if (plan.TargetBooks.Contains(index.Book))
                {
                    plan.Conflicts.Add("BOOK," + index.Book + ",目标定额库索引已存在");
                }
            }

            foreach (SectionRow row in plan.SourceSections)
            {
                if (plan.TargetSectionKeys.Contains(row.Book + "\t" + row.SectionNo))
                {
                    plan.Conflicts.Add("SECTION," + row.Book + "/" + row.SectionNo + ",目标定额节索引已存在");
                }
            }

            foreach (QuotaRow row in plan.SourceQuotas)
            {
                if (plan.TargetQuotaKeys.Contains(row.Book + "\t" + row.Code))
                {
                    plan.Conflicts.Add("QUOTA," + row.Book + "/" + row.Code + ",目标定额库已存在");
                }
            }
        }

        private static void BuildResourceMap(MigrationPlan plan, Crypto crypto)
        {
            HashSet<int> usedOldCodes = new HashSet<int>();
            foreach (QuotaRow quota in plan.SourceQuotas)
            {
                string plain = crypto.DecryptOld(quota.ConsumeEncrypted);
                quota.ConsumePlain = plain;
                foreach (ConsumePart part in ParseConsume(plain))
                {
                    usedOldCodes.Add(part.OldCode);
                }
            }

            foreach (int oldCode in usedOldCodes)
            {
                ResourceMapping mapping = new ResourceMapping { OldCode = oldCode };
                if ((oldCode >= 1 && oldCode <= 7) || oldCode == 10 || oldCode == 11)
                {
                    mapping.Kind = ResourceKind.Labor;
                    mapping.NewCode = oldCode;
                    mapping.MatchStatus = oldCode <= 7 ? "人工保留" : "特殊人工保留";
                    plan.ResourceMap[oldCode] = mapping;
                    continue;
                }

                Resource source;
                if (plan.SourceMaterials.TryGetValue(oldCode, out source))
                {
                    mapping.Kind = ResourceKind.Material;
                    mapping.Source = source;
                    ResolveResource(plan, mapping, plan.TargetMaterials);
                    plan.ResourceMap[oldCode] = mapping;
                    continue;
                }

                if (plan.SourceMachines.TryGetValue(oldCode, out source))
                {
                    mapping.Kind = ResourceKind.Machine;
                    mapping.Source = source;
                    ResolveResource(plan, mapping, plan.TargetMachines);
                    plan.ResourceMap[oldCode] = mapping;
                    continue;
                }

                mapping.Kind = ResourceKind.Unknown;
                mapping.MatchStatus = "无法识别";
                mapping.NewCode = oldCode;
                plan.ResourceMap[oldCode] = mapping;
                plan.UnknownCodes.Add(mapping);
            }
        }

        private static void ResolveResource(MigrationPlan plan, ResourceMapping mapping, Dictionary<string, List<Resource>> targetByName)
        {
            List<Resource> targets;
            if (targetByName.TryGetValue(mapping.Source.Name, out targets) && targets.Count > 0)
            {
                ApplyChosenTarget(plan, mapping, ChooseByUnit(targets, mapping.Source.Unit), "名称单位匹配", "名称匹配单位不同", 100);
                return;
            }

            string sourceNormalized = NormalizeResourceName(mapping.Source.Name);
            Resource normalizedTarget = FindNormalizedTarget(targetByName, sourceNormalized, mapping.Source.Unit);
            if (normalizedTarget != null)
            {
                ApplyChosenTarget(plan, mapping, normalizedTarget, "规范化名称匹配", "规范化名称匹配单位不同", 98);
                plan.NameReviewMappings.Add(mapping);
                return;
            }

            ResourceCandidate candidate = FindFuzzyTarget(targetByName, mapping.Source);
            if (candidate.Target != null &&
                candidate.Score >= 92 &&
                candidate.Score - candidate.SecondScore >= 8 &&
                LengthRatioAcceptable(sourceNormalized, NormalizeResourceName(candidate.Target.Name)) &&
                NumbersEqual(mapping.Source.Name, candidate.Target.Name))
            {
                ApplyChosenTarget(plan, mapping, candidate.Target, "相似名称匹配", "相似名称匹配单位不同", candidate.Score);
                mapping.MatchReason = "自动采用唯一高相似候选";
                plan.NameReviewMappings.Add(mapping);
                return;
            }

            Resource supplement = mapping.Source.Clone();
            supplement.DocNo = mapping.Kind == ResourceKind.Material ? "2020迁移补充材料" : "2020迁移补充机械";
            supplement.ImportedAt = DateTime.Now;
            if (mapping.Kind == ResourceKind.Material)
            {
                supplement.NewCode = AllocateMaterialCode(plan);
            }
            else
            {
                supplement.NewCode = AllocateMachineCode(plan);
            }

            mapping.NewCode = supplement.NewCode;
            mapping.Target = supplement;
            mapping.MatchStatus = "2024缺失，补充";
            ResourceCandidate reviewCandidate = candidate.Target != null && candidate.Score >= 60
                ? candidate
                : FindReviewCandidate(targetByName, mapping.Source);
            if (reviewCandidate.Target != null)
            {
                mapping.CandidateName = reviewCandidate.Target.Name;
                mapping.CandidateCode = reviewCandidate.Target.NewCode;
                mapping.CandidateUnit = reviewCandidate.Target.Unit;
                mapping.CandidateScore = reviewCandidate.Score;
                mapping.MatchReason = reviewCandidate.Score >= 60
                    ? "无可靠唯一相似候选，按补充处理"
                    : "低置信兜底候选，仅供人工审核，按补充处理";
            }
            else
            {
                mapping.MatchReason = "无相似候选，按补充处理";
            }
            plan.MissingResources.Add(mapping);
        }

        private static Resource ChooseByUnit(List<Resource> targets, string unit)
        {
            foreach (Resource target in targets)
            {
                if (String.Equals(target.Unit, unit, StringComparison.Ordinal))
                {
                    return target;
                }
            }

            return targets[0];
        }

        private static void ApplyChosenTarget(MigrationPlan plan, ResourceMapping mapping, Resource chosen, string sameUnitStatus, string diffUnitStatus, int score)
        {
            mapping.NewCode = chosen.NewCode;
            mapping.Target = chosen;
            mapping.MatchScore = score;
            mapping.UnitDifference = !String.Equals(mapping.Source.Unit, chosen.Unit, StringComparison.Ordinal);
            mapping.MatchStatus = mapping.UnitDifference ? diffUnitStatus : sameUnitStatus;
            if (mapping.UnitDifference)
            {
                plan.UnitDifferences.Add(mapping);
            }
        }

        private static Resource FindNormalizedTarget(Dictionary<string, List<Resource>> targetByName, string sourceNormalized, string sourceUnit)
        {
            List<Resource> matches = new List<Resource>();
            foreach (KeyValuePair<string, List<Resource>> pair in targetByName)
            {
                if (String.Equals(NormalizeResourceName(pair.Key), sourceNormalized, StringComparison.Ordinal))
                {
                    matches.AddRange(pair.Value);
                }
            }

            if (matches.Count == 0)
            {
                return null;
            }

            return ChooseByUnit(matches, sourceUnit);
        }

        private static ResourceCandidate FindFuzzyTarget(Dictionary<string, List<Resource>> targetByName, Resource source)
        {
            string sourceNormalized = NormalizeResourceName(source.Name);
            ResourceCandidate best = new ResourceCandidate();
            foreach (KeyValuePair<string, List<Resource>> pair in targetByName)
            {
                string targetNormalized = NormalizeResourceName(pair.Key);
                if (!NumbersCompatibleNormalized(sourceNormalized, targetNormalized))
                {
                    continue;
                }

                if (source.Kind == ResourceKind.Machine && !MachineNameCompatible(source.Name, pair.Key))
                {
                    continue;
                }

                int score = SimilarityScore(sourceNormalized, targetNormalized);
                if (source.Kind == ResourceKind.Machine)
                {
                    score = Math.Min(100, score + MachineKeywordBonus(source.Name, pair.Key));
                }

                foreach (Resource target in pair.Value)
                {
                    int unitAdjusted = String.Equals(source.Unit, target.Unit, StringComparison.Ordinal) ? Math.Min(100, score + 3) : score;
                    if (unitAdjusted > best.Score)
                    {
                        best.SecondScore = best.Score;
                        best.Score = unitAdjusted;
                        best.Target = target;
                    }
                    else if (unitAdjusted > best.SecondScore)
                    {
                        best.SecondScore = unitAdjusted;
                    }
                }
            }

            return best;
        }

        private static ResourceCandidate FindReviewCandidate(Dictionary<string, List<Resource>> targetByName, Resource source)
        {
            string sourceNormalized = NormalizeResourceName(source.Name);
            ResourceCandidate best = new ResourceCandidate();
            foreach (KeyValuePair<string, List<Resource>> pair in targetByName)
            {
                string targetNormalized = NormalizeResourceName(pair.Key);
                int score = ReviewSimilarityScore(source.Name, pair.Key, sourceNormalized, targetNormalized, source.Kind);
                foreach (Resource target in pair.Value)
                {
                    int adjusted = score;
                    if (String.Equals(source.Unit, target.Unit, StringComparison.Ordinal))
                    {
                        adjusted += 5;
                    }
                    else
                    {
                        adjusted -= 6;
                    }

                    if (adjusted > best.Score)
                    {
                        best.SecondScore = best.Score;
                        best.Score = Math.Max(1, Math.Min(100, adjusted));
                        best.Target = target;
                    }
                    else if (adjusted > best.SecondScore)
                    {
                        best.SecondScore = adjusted;
                    }
                }
            }

            return best;
        }

        private static int ReviewSimilarityScore(string sourceName, string targetName, string sourceNormalized, string targetNormalized, ResourceKind kind)
        {
            int score = SimilarityScore(sourceNormalized, targetNormalized);
            score = Math.Max(score, TokenOverlapScore(sourceNormalized, targetNormalized));

            List<string> sourceNumbers = ExtractNumbers(sourceNormalized);
            List<string> targetNumbers = ExtractNumbers(targetNormalized);
            int sharedNumbers = CountShared(sourceNumbers, targetNumbers);
            if (sourceNumbers.Count > 0)
            {
                score += sharedNumbers * 4;
                score -= Math.Max(0, sourceNumbers.Count - sharedNumbers) * 6;
            }

            if (kind == ResourceKind.Machine)
            {
                score += MachineKeywordBonus(sourceName, targetName);
                if (!MachineNameCompatible(sourceName, targetName))
                {
                    score -= 28;
                }
            }
            else if (kind == ResourceKind.Material)
            {
                score += MaterialKeywordBonus(sourceName, targetName);
            }

            return Math.Max(1, Math.Min(100, score));
        }

        private static int TokenOverlapScore(string sourceNormalized, string targetNormalized)
        {
            HashSet<string> sourceTokens = BuildNgrams(sourceNormalized);
            HashSet<string> targetTokens = BuildNgrams(targetNormalized);
            if (sourceTokens.Count == 0 || targetTokens.Count == 0)
            {
                return 0;
            }

            int shared = 0;
            foreach (string token in sourceTokens)
            {
                if (targetTokens.Contains(token))
                {
                    shared++;
                }
            }

            return (int)Math.Round((200.0 * shared) / (sourceTokens.Count + targetTokens.Count), MidpointRounding.AwayFromZero);
        }

        private static HashSet<string> BuildNgrams(string text)
        {
            HashSet<string> tokens = new HashSet<string>(StringComparer.Ordinal);
            if (String.IsNullOrEmpty(text))
            {
                return tokens;
            }

            if (text.Length <= 2)
            {
                tokens.Add(text);
                return tokens;
            }

            for (int i = 0; i < text.Length - 1; i++)
            {
                tokens.Add(text.Substring(i, 2));
            }

            return tokens;
        }

        private static int CountShared(List<string> left, List<string> right)
        {
            int count = 0;
            foreach (string value in left)
            {
                if (right.Contains(value))
                {
                    count++;
                }
            }

            return count;
        }

        private static int MaterialKeywordBonus(string sourceName, string targetName)
        {
            int bonus = 0;
            string[] tokens = new[]
            {
                "钢筋", "钢丝", "水泥", "砂", "碎石", "粉煤灰", "减水剂",
                "混凝土", "沥青", "绝缘子", "电缆", "横担", "抱箍", "螺栓",
                "螺母", "垫圈", "支架", "管", "板", "砖", "线"
            };

            foreach (string token in tokens)
            {
                if (sourceName.IndexOf(token, StringComparison.Ordinal) >= 0 && targetName.IndexOf(token, StringComparison.Ordinal) >= 0)
                {
                    bonus += 3;
                }
            }

            return Math.Min(12, bonus);
        }

        private static bool MachineNameCompatible(string sourceName, string targetName)
        {
            if (HasAny(sourceName, "内燃") && HasAny(targetName, "电动", "电力"))
            {
                return false;
            }

            if (HasAny(sourceName, "电动", "电力") && HasAny(targetName, "内燃"))
            {
                return false;
            }

            string[] exclusiveTypes = new[]
            {
                "卷扬机", "空气压缩机", "清水泵", "污水泵", "泥浆泵", "发电机",
                "起重机", "挖掘机", "装载机", "推土机", "压路机", "摊铺机",
                "搅拌机", "搅拌站", "搅拌船", "输送泵车", "输送泵",
                "运输车", "汽车", "叉车", "钻机", "流量计", "示波器",
                "分析仪", "测试仪", "信号发生器", "平台车"
            };

            string sourceType = FirstKeyword(sourceName, exclusiveTypes);
            string targetType = FirstKeyword(targetName, exclusiveTypes);
            if (!String.IsNullOrEmpty(sourceType) && !String.IsNullOrEmpty(targetType) && !String.Equals(sourceType, targetType, StringComparison.Ordinal))
            {
                if (!(sourceType == "运输车" && targetType == "汽车") && !(sourceType == "汽车" && targetType == "运输车"))
                {
                    return false;
                }
            }

            if (HasAny(sourceName, "单筒") && HasAny(targetName, "双筒"))
            {
                return false;
            }

            if (HasAny(sourceName, "双筒") && HasAny(targetName, "单筒"))
            {
                return false;
            }

            if (HasAny(sourceName, "慢速") && HasAny(targetName, "快速"))
            {
                return false;
            }

            if (HasAny(sourceName, "快速") && HasAny(targetName, "慢速"))
            {
                return false;
            }

            return true;
        }

        private static int MachineKeywordBonus(string sourceName, string targetName)
        {
            int bonus = 0;
            string[] tokens = new[]
            {
                "内燃", "电动", "液压", "履带式", "轮胎式", "单筒", "双筒",
                "慢速", "快速", "活塞式", "离心", "清水", "污水", "泥浆",
                "卷扬机", "空气压缩机", "起重机", "挖掘机", "装载机", "推土机",
                "搅拌站", "输送泵", "流量计", "示波器", "分析仪", "信号发生器"
            };

            foreach (string token in tokens)
            {
                if (sourceName.IndexOf(token, StringComparison.Ordinal) >= 0 && targetName.IndexOf(token, StringComparison.Ordinal) >= 0)
                {
                    bonus += 2;
                }
            }

            return Math.Min(8, bonus);
        }

        private static string FirstKeyword(string text, string[] keywords)
        {
            foreach (string keyword in keywords)
            {
                if (text.IndexOf(keyword, StringComparison.Ordinal) >= 0)
                {
                    return keyword;
                }
            }

            return "";
        }

        private static bool HasAny(string text, params string[] keywords)
        {
            foreach (string keyword in keywords)
            {
                if (text.IndexOf(keyword, StringComparison.Ordinal) >= 0)
                {
                    return true;
                }
            }

            return false;
        }

        private static string NormalizeResourceName(string text)
        {
            if (String.IsNullOrWhiteSpace(text))
            {
                return "";
            }

            string value = text.Trim().ToLowerInvariant();
            value = value.Replace('（', '(').Replace('）', ')');
            value = value.Replace('，', ',').Replace('。', '.').Replace('：', ':');
            value = value.Replace('～', '~').Replace('－', '-').Replace('—', '-').Replace('–', '-');
            value = value.Replace('＜', '<').Replace('≤', '<').Replace('≦', '<');
            value = value.Replace('＞', '>').Replace('≥', '>').Replace('≧', '>');
            value = value.Replace('φ', 'f').Replace('Φ', 'f');
            value = value.Replace("活塞式", "");
            value = value.Replace("式", "");
            value = value.Replace("级", "");
            value = value.Replace("mm", "");
            value = value.Replace(" ", "");
            value = value.Replace("　", "");
            value = value.Replace("=", "");
            value = value.Replace("(", "").Replace(")", "");
            value = value.Replace(",", "");
            return value;
        }

        private static bool NumbersCompatible(string sourceName, string targetName)
        {
            return NumbersCompatibleNormalized(NormalizeResourceName(sourceName), NormalizeResourceName(targetName));
        }

        private static bool NumbersEqual(string sourceName, string targetName)
        {
            List<string> sourceNumbers = ExtractNumbers(NormalizeResourceName(sourceName));
            List<string> targetNumbers = ExtractNumbers(NormalizeResourceName(targetName));
            if (sourceNumbers.Count != targetNumbers.Count)
            {
                return false;
            }

            foreach (string number in sourceNumbers)
            {
                if (!targetNumbers.Contains(number))
                {
                    return false;
                }
            }

            return true;
        }

        private static bool LengthRatioAcceptable(string sourceNormalized, string targetNormalized)
        {
            int min = Math.Min(sourceNormalized.Length, targetNormalized.Length);
            int max = Math.Max(sourceNormalized.Length, targetNormalized.Length);
            if (max == 0)
            {
                return false;
            }

            return ((double)min / (double)max) >= 0.75;
        }

        private static bool NumbersCompatibleNormalized(string sourceNormalized, string targetNormalized)
        {
            List<string> sourceNumbers = ExtractNumbers(sourceNormalized);
            List<string> targetNumbers = ExtractNumbers(targetNormalized);
            foreach (string number in sourceNumbers)
            {
                if (!targetNumbers.Contains(number))
                {
                    return false;
                }
            }

            return true;
        }

        private static List<string> ExtractNumbers(string text)
        {
            List<string> numbers = new List<string>();
            foreach (Match match in Regex.Matches(text ?? "", @"\d+(?:\.\d+)?"))
            {
                if (!numbers.Contains(match.Value))
                {
                    numbers.Add(match.Value);
                }
            }

            return numbers;
        }

        private static int SimilarityScore(string left, string right)
        {
            if (String.Equals(left, right, StringComparison.Ordinal))
            {
                return 100;
            }

            if (left.Length == 0 || right.Length == 0)
            {
                return 0;
            }

            int lcs = LongestCommonSubsequence(left, right);
            double coverage = (2.0 * lcs) / (left.Length + right.Length);
            int containsBonus = (left.Contains(right) || right.Contains(left)) ? 8 : 0;
            int score = (int)Math.Round(coverage * 100.0, MidpointRounding.AwayFromZero) + containsBonus;
            return Math.Max(0, Math.Min(100, score));
        }

        private static int LongestCommonSubsequence(string left, string right)
        {
            int[,] dp = new int[left.Length + 1, right.Length + 1];
            for (int i = 1; i <= left.Length; i++)
            {
                for (int j = 1; j <= right.Length; j++)
                {
                    if (left[i - 1] == right[j - 1])
                    {
                        dp[i, j] = dp[i - 1, j - 1] + 1;
                    }
                    else
                    {
                        dp[i, j] = Math.Max(dp[i - 1, j], dp[i, j - 1]);
                    }
                }
            }

            return dp[left.Length, right.Length];
        }

        private static int AllocateMaterialCode(MigrationPlan plan)
        {
            if (plan.MaterialRangeEnd > 0 && plan.NextMaterialCode > plan.MaterialRangeEnd)
            {
                throw new InvalidOperationException("No material supplement code left in configured range.");
            }

            return checked((int)plan.NextMaterialCode++);
        }

        private static int AllocateMachineCode(MigrationPlan plan)
        {
            if (plan.MachineRangeEnd > 0 && plan.NextMachineCode > plan.MachineRangeEnd)
            {
                throw new InvalidOperationException("No machine supplement code left in configured range.");
            }

            return checked((int)plan.NextMachineCode++);
        }

        private static void BuildMigratedConsumptions(MigrationPlan plan, Crypto crypto)
        {
            foreach (QuotaRow quota in plan.SourceQuotas)
            {
                List<ConsumePart> parts = ParseConsume(quota.ConsumePlain);
                StringBuilder plain = new StringBuilder();
                foreach (ConsumePart part in parts)
                {
                    ResourceMapping mapping;
                    if (!plan.ResourceMap.TryGetValue(part.OldCode, out mapping) || mapping.Kind == ResourceKind.Unknown)
                    {
                        plan.Conflicts.Add("CONSUME," + quota.Book + "/" + quota.Code + ",无法映射电算代号 " + part.OldCode.ToString(CultureInfo.InvariantCulture));
                        continue;
                    }

                    if (plain.Length > 0)
                    {
                        plain.Append('@');
                    }

                    plain.Append(mapping.NewCode.ToString("D9", CultureInfo.InvariantCulture));
                    plain.Append(part.QuantityText);
                    quota.ConsumeRows.Add(new ConsumeRow
                    {
                        Book = quota.Book,
                        QuotaCode = quota.Code,
                        Code = mapping.NewCode,
                        Quantity = part.QuantityText
                    });
                }

                quota.MigratedConsumePlain = plain.ToString();
                quota.MigratedConsumeEncrypted = crypto.EncryptOld(quota.MigratedConsumePlain);
            }
        }

        private static List<ConsumePart> ParseConsume(string plain)
        {
            List<ConsumePart> rows = new List<ConsumePart>();
            if (String.IsNullOrWhiteSpace(plain))
            {
                return rows;
            }

            string[] parts = plain.Split(new[] { '@' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string part in parts)
            {
                if (part.Length <= 9)
                {
                    throw new FormatException("Invalid consume item: " + part);
                }

                int code;
                if (!Int32.TryParse(part.Substring(0, 9), NumberStyles.Integer, CultureInfo.InvariantCulture, out code))
                {
                    throw new FormatException("Invalid consume code: " + part);
                }

                string qty = part.Substring(9);
                decimal parsed;
                if (!Decimal.TryParse(qty, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed))
                {
                    throw new FormatException("Invalid consume quantity: " + part);
                }

                rows.Add(new ConsumePart { OldCode = code, QuantityText = qty });
            }

            return rows;
        }

        private static void WriteReports(MigrationPlan plan, string reportDir)
        {
            WriteSummary(plan, Path.Combine(reportDir, "precheck-summary.txt"));
            WriteResourceMap(plan, Path.Combine(reportDir, "resource-map.csv"));
            WriteMissingResources(plan, Path.Combine(reportDir, "missing-resources.csv"));
            WriteUnitDifferences(plan, Path.Combine(reportDir, "unit-differences.csv"));
            WriteNameReview(plan, Path.Combine(reportDir, "name-review.csv"));
            WriteConflicts(plan, Path.Combine(reportDir, "conflicts.csv"));
            WriteRollback(plan, Path.Combine(reportDir, "rollback.sql"));
        }

        private static void WriteSummary(MigrationPlan plan, string path)
        {
            StringBuilder text = new StringBuilder();
            text.AppendLine("2020 概算/估算定额迁移预检");
            text.AppendLine("BuiltAt: " + plan.BuiltAt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture));
            text.AppendLine("Source: " + SourceDb);
            text.AppendLine("Target: " + TargetDb);
            text.AppendLine("Books: " + plan.SourceIndexes.Count.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("Sections: " + plan.SourceSections.Count.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("Quotas: " + plan.SourceQuotas.Count.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("ConsumeRows: " + plan.TotalConsumeRows.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("ResourceCodes: " + plan.ResourceMap.Count.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("DirectOrUnitDiffMappings: " + plan.DirectMappings.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("UnitDifferences: " + plan.UnitDifferences.Count.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("SupplementResources: " + plan.MissingResources.Count.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("SupplementResourcesWithoutCandidate: " + plan.MissingResourcesWithoutCandidate.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("LowConfidenceSupplementCandidates: " + plan.LowConfidenceSupplementCandidates.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("UnknownCodes: " + plan.UnknownCodes.Count.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("Conflicts: " + plan.Conflicts.Count.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("MaterialSupplementRange: " + plan.MaterialRangeStart.ToString(CultureInfo.InvariantCulture) + "-" + plan.MaterialRangeEnd.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("MachineSupplementRange: " + plan.MachineRangeStart.ToString(CultureInfo.InvariantCulture) + "-" + plan.MachineRangeEnd.ToString(CultureInfo.InvariantCulture));
            File.WriteAllText(path, text.ToString(), Encoding.UTF8);
        }

        private static void WriteResourceMap(MigrationPlan plan, string path)
        {
            using (StreamWriter writer = new StreamWriter(path, false, Encoding.UTF8))
            {
                writer.WriteLine("kind,old_code,new_code,status,source_name,target_name,source_unit,target_unit,match_score,reason");
                foreach (ResourceMapping row in plan.SortedResourceMap)
                {
                    writer.WriteLine(Csv(row.Kind.ToString(), row.OldCode, row.NewCode, row.MatchStatus, row.SourceName, row.TargetName, row.SourceUnit, row.TargetUnit, row.MatchScore, row.MatchReason));
                }
            }
        }

        private static void WriteMissingResources(MigrationPlan plan, string path)
        {
            using (StreamWriter writer = new StreamWriter(path, false, Encoding.UTF8))
            {
                writer.WriteLine("kind,old_code,new_supplement_code,source_name,source_unit,base_price,best_candidate_code,best_candidate_name,best_candidate_unit,best_candidate_score,reason");
                foreach (ResourceMapping row in plan.MissingResources)
                {
                    writer.WriteLine(Csv(row.Kind.ToString(), row.OldCode, row.NewCode, row.SourceName, row.SourceUnit, row.Source == null ? 0 : row.Source.BasePrice, row.CandidateCode, row.CandidateName, row.CandidateUnit, row.CandidateScore, row.MatchReason));
                }
            }
        }

        private static void WriteUnitDifferences(MigrationPlan plan, string path)
        {
            using (StreamWriter writer = new StreamWriter(path, false, Encoding.UTF8))
            {
                writer.WriteLine("kind,old_code,new_code,source_name,target_name,source_unit,target_unit,status,match_score");
                foreach (ResourceMapping row in plan.UnitDifferences)
                {
                    writer.WriteLine(Csv(row.Kind.ToString(), row.OldCode, row.NewCode, row.SourceName, row.TargetName, row.SourceUnit, row.TargetUnit, row.MatchStatus, row.MatchScore));
                }
            }
        }

        private static void WriteNameReview(MigrationPlan plan, string path)
        {
            using (StreamWriter writer = new StreamWriter(path, false, Encoding.UTF8))
            {
                writer.WriteLine("kind,old_code,new_code,status,source_name,target_name,source_unit,target_unit,match_score,reason");
                foreach (ResourceMapping row in plan.NameReviewMappings)
                {
                    writer.WriteLine(Csv(row.Kind.ToString(), row.OldCode, row.NewCode, row.MatchStatus, row.SourceName, row.TargetName, row.SourceUnit, row.TargetUnit, row.MatchScore, row.MatchReason));
                }
            }
        }

        private static void WriteConflicts(MigrationPlan plan, string path)
        {
            using (StreamWriter writer = new StreamWriter(path, false, Encoding.UTF8))
            {
                writer.WriteLine("type,key,message");
                foreach (string conflict in plan.Conflicts)
                {
                    writer.WriteLine(conflict);
                }
            }
        }

        private static void WriteRollback(MigrationPlan plan, string path)
        {
            using (StreamWriter writer = new StreamWriter(path, false, Encoding.UTF8))
            {
                writer.WriteLine("-- Generated rollback for Migrate2020EstimateTo2024");
                writer.WriteLine("begin transaction;");
                foreach (BookIndex index in plan.SourceIndexes)
                {
                    writer.WriteLine("delete from 定额库消耗 where 书号=N'" + SqlLiteral(index.Book) + "';");
                    writer.WriteLine("delete from 定额库 where 书号=N'" + SqlLiteral(index.Book) + "';");
                    writer.WriteLine("delete from 定额节索引 where 书号=N'" + SqlLiteral(index.Book) + "';");
                    writer.WriteLine("delete from 定额库索引 where 书号=N'" + SqlLiteral(index.Book) + "';");
                }

                foreach (ResourceMapping mapping in plan.MissingResources)
                {
                    if (mapping.Kind == ResourceKind.Material)
                    {
                        writer.WriteLine("delete from 材料单价库 where 文号=N'2020迁移补充材料' and 电算代号=" + mapping.NewCode.ToString(CultureInfo.InvariantCulture) + ";");
                    }
                    else if (mapping.Kind == ResourceKind.Machine)
                    {
                        writer.WriteLine("delete from 台班定额库 where 文号=N'2020迁移补充机械' and 电算代号=" + mapping.NewCode.ToString(CultureInfo.InvariantCulture) + ";");
                    }
                }

                writer.WriteLine("commit transaction;");
            }
        }

        private static void PrintSummary(MigrationPlan plan)
        {
            Console.WriteLine("Books: " + plan.SourceIndexes.Count.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine("Quotas: " + plan.SourceQuotas.Count.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine("ConsumeRows: " + plan.TotalConsumeRows.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine("ResourceCodes: " + plan.ResourceMap.Count.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine("UnitDifferences: " + plan.UnitDifferences.Count.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine("SupplementResources: " + plan.MissingResources.Count.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine("UnknownCodes: " + plan.UnknownCodes.Count.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine("Conflicts: " + plan.Conflicts.Count.ToString(CultureInfo.InvariantCulture));
        }

        private static void Apply(MigrationPlan plan)
        {
            using (SqlConnection target = OpenDb(TargetDb))
            using (SqlTransaction tx = target.BeginTransaction())
            {
                try
                {
                    foreach (ResourceMapping mapping in plan.MissingResources)
                    {
                        if (mapping.Kind == ResourceKind.Material)
                        {
                            InsertMaterial(target, tx, mapping.Target);
                        }
                        else if (mapping.Kind == ResourceKind.Machine)
                        {
                            InsertMachine(target, tx, mapping.Target);
                        }
                    }

                    foreach (BookIndex index in plan.SourceIndexes)
                    {
                        InsertBookIndex(target, tx, index);
                    }

                    foreach (SectionRow row in plan.SourceSections)
                    {
                        InsertSection(target, tx, row);
                    }

                    foreach (QuotaRow row in plan.SourceQuotas)
                    {
                        InsertQuota(target, tx, row);
                        foreach (ConsumeRow consume in row.ConsumeRows)
                        {
                            InsertConsume(target, tx, consume);
                        }
                    }

                    tx.Commit();
                }
                catch
                {
                    tx.Rollback();
                    throw;
                }
            }
        }

        private static void Verify(string reportDir)
        {
            using (SqlConnection target = OpenDb(TargetDb))
            {
                int bookCount = Count(target, "select count(*) from 定额库索引 where 分类 in ('概算定额','估算定额')");
                int quotaCount = Count(target, "select count(*) from 定额库 a join 定额库索引 b on a.书号=b.书号 where b.分类 in ('概算定额','估算定额')");
                int consumeCount = Count(target, "select count(*) from 定额库消耗 c join 定额库索引 b on c.书号=b.书号 where b.分类 in ('概算定额','估算定额')");
                int budgetBooks = Count(target, "select count(*) from 定额库索引 where 分类='预算定额' and 书号 like '%_2024'");
                int lyCount = Count(target, "select count(*) from 定额库 where 书号='LY_2024'");

                string text = "VerifyAt: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + Environment.NewLine
                    + "MigratedBooksInTarget: " + bookCount.ToString(CultureInfo.InvariantCulture) + Environment.NewLine
                    + "MigratedQuotasInTarget: " + quotaCount.ToString(CultureInfo.InvariantCulture) + Environment.NewLine
                    + "MigratedConsumeRowsInTarget: " + consumeCount.ToString(CultureInfo.InvariantCulture) + Environment.NewLine
                    + "Original2024BudgetBooks: " + budgetBooks.ToString(CultureInfo.InvariantCulture) + Environment.NewLine
                    + "LY_2024QuotaCount: " + lyCount.ToString(CultureInfo.InvariantCulture) + Environment.NewLine;
                File.WriteAllText(Path.Combine(reportDir, "verify-summary.txt"), text, Encoding.UTF8);
                Console.Write(text);
            }
        }

        private static int Count(SqlConnection conn, string sql)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = sql;
                return Convert.ToInt32(cmd.ExecuteScalar(), CultureInfo.InvariantCulture);
            }
        }

        private static void InsertBookIndex(SqlConnection conn, SqlTransaction tx, BookIndex row)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "insert into 定额库索引(书号,head,内容,字头,互斥表,分类,专业名称,顺号,现行定额,说明) values(@book,@head,@content,@prefix,@exclusive,@category,@specialty,@order,@current,@note)";
                Add(cmd, "@book", row.Book);
                Add(cmd, "@head", row.Head);
                Add(cmd, "@content", row.Content);
                Add(cmd, "@prefix", row.Prefix);
                Add(cmd, "@exclusive", row.ExclusiveTable);
                Add(cmd, "@category", row.Category);
                Add(cmd, "@specialty", row.Specialty);
                Add(cmd, "@order", row.Order);
                Add(cmd, "@current", row.IsCurrent ? 1 : 0);
                Add(cmd, "@note", row.Note);
                cmd.ExecuteNonQuery();
            }
        }

        private static void InsertSection(SqlConnection conn, SqlTransaction tx, SectionRow row)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "insert into 定额节索引(书号,节号,节名称,描述,起始号,终止号) values(@book,@no,@name,@desc,@start,@end)";
                Add(cmd, "@book", row.Book);
                Add(cmd, "@no", row.SectionNo);
                Add(cmd, "@name", row.SectionName);
                Add(cmd, "@desc", row.Description);
                Add(cmd, "@start", row.StartNo);
                Add(cmd, "@end", row.EndNo);
                cmd.ExecuteNonQuery();
            }
        }

        private static void InsertQuota(SqlConnection conn, SqlTransaction tx, QuotaRow row)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "insert into 定额库(书号,定额编号,定额名称,单位,消耗,工作内容,基本定额,基价,工费,料费,机费,单重,节号,流水号,LOCK,使用,条目序号,导入时间) values(@book,@code,@name,@unit,@consume,@work,@baseQuota,@basePrice,@laborFee,@materialFee,@machineFee,@weight,@section,@sort,@lock,@use,@item,@time)";
                Add(cmd, "@book", row.Book);
                Add(cmd, "@code", row.Code);
                Add(cmd, "@name", row.Name);
                Add(cmd, "@unit", row.Unit);
                Add(cmd, "@consume", row.MigratedConsumeEncrypted);
                Add(cmd, "@work", row.WorkContent);
                Add(cmd, "@baseQuota", row.BaseQuota);
                Add(cmd, "@basePrice", row.BasePrice);
                Add(cmd, "@laborFee", row.LaborFee);
                Add(cmd, "@materialFee", row.MaterialFee);
                Add(cmd, "@machineFee", row.MachineFee);
                Add(cmd, "@weight", row.UnitWeight);
                Add(cmd, "@section", row.SectionNo);
                Add(cmd, "@sort", row.SortOrder);
                Add(cmd, "@lock", row.Locked);
                Add(cmd, "@use", row.Use);
                Add(cmd, "@item", row.ItemSequence);
                Add(cmd, "@time", row.ImportedAt.HasValue ? (object)row.ImportedAt.Value : DateTime.Now);
                cmd.ExecuteNonQuery();
            }
        }

        private static void InsertConsume(SqlConnection conn, SqlTransaction tx, ConsumeRow row)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "insert into 定额库消耗(书号,定额编号,电算代号,消耗) values(@book,@quota,@code,@qty)";
                Add(cmd, "@book", row.Book);
                Add(cmd, "@quota", row.QuotaCode);
                Add(cmd, "@code", row.Code);
                Add(cmd, "@qty", row.Quantity);
                cmd.ExecuteNonQuery();
            }
        }

        private static void InsertMaterial(SqlConnection conn, SqlTransaction tx, Resource row)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "insert into 材料单价库(文号,电算代号,材料名称,单位,基期单价,编制期价,编制期含税价,单重,换算系数,汇总标志,主材标志,材料运输类别,LOCK,甲供方式,导入时间) values(@doc,@code,@name,@unit,@base,@current,@tax,@weight,@factor,@summary,@main,@transport,@lock,@supply,@time)";
                Add(cmd, "@doc", row.DocNo);
                Add(cmd, "@code", row.NewCode);
                Add(cmd, "@name", row.Name);
                Add(cmd, "@unit", row.Unit);
                Add(cmd, "@base", row.BasePrice);
                Add(cmd, "@current", row.CurrentPrice);
                Add(cmd, "@tax", row.CurrentTaxPrice);
                Add(cmd, "@weight", row.UnitWeight);
                Add(cmd, "@factor", row.ConvertFactor);
                Add(cmd, "@summary", row.SummaryFlag);
                Add(cmd, "@main", row.MainMaterialFlag);
                Add(cmd, "@transport", row.TransportCategory);
                Add(cmd, "@lock", row.Locked);
                Add(cmd, "@supply", row.SupplyMode);
                Add(cmd, "@time", row.ImportedAt.HasValue ? (object)row.ImportedAt.Value : DateTime.Now);
                cmd.ExecuteNonQuery();
            }
        }

        private static void InsertMachine(SqlConnection conn, SqlTransaction tx, Resource row)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "insert into 台班定额库(文号,电算代号,机械台班名称,折旧费,检修费,维护费,安装拆卸费,人工,汽油,柴油,煤,电,水,天然气,汇总标志,基价,接触网封锁线路标志,养路费系数,其他费用,LOCK,导入时间) values(@doc,@code,@name,@dep,@repair,@maint,@install,@labor,@gas,@diesel,@coal,@electric,@water,@ng,@summary,@base,@block,@road,@other,@lock,@time)";
                Add(cmd, "@doc", row.DocNo);
                Add(cmd, "@code", row.NewCode);
                Add(cmd, "@name", row.Name);
                Add(cmd, "@dep", row.DepreciationFee);
                Add(cmd, "@repair", row.RepairFee);
                Add(cmd, "@maint", row.MaintenanceFee);
                Add(cmd, "@install", row.InstallUninstallFee);
                Add(cmd, "@labor", row.Labor);
                Add(cmd, "@gas", row.Gasoline);
                Add(cmd, "@diesel", row.Diesel);
                Add(cmd, "@coal", row.Coal);
                Add(cmd, "@electric", row.Electricity);
                Add(cmd, "@water", row.Water);
                Add(cmd, "@ng", row.NaturalGas);
                Add(cmd, "@summary", row.SummaryFlag);
                Add(cmd, "@base", row.BasePrice);
                Add(cmd, "@block", row.ContactLineBlock);
                Add(cmd, "@road", row.RoadFeeFactor);
                Add(cmd, "@other", row.OtherFee);
                Add(cmd, "@lock", row.Locked);
                Add(cmd, "@time", row.ImportedAt.HasValue ? (object)row.ImportedAt.Value : DateTime.Now);
                cmd.ExecuteNonQuery();
            }
        }

        private static void Add(SqlCommand cmd, string name, object value)
        {
            cmd.Parameters.AddWithValue(name, value ?? DBNull.Value);
        }

        private static string S(SqlDataReader r, int index)
        {
            return r.IsDBNull(index) ? "" : Convert.ToString(r.GetValue(index), CultureInfo.CurrentCulture).Trim();
        }

        private static int I(SqlDataReader r, int index)
        {
            return r.IsDBNull(index) ? 0 : Convert.ToInt32(r.GetValue(index), CultureInfo.InvariantCulture);
        }

        private static double D(SqlDataReader r, int index)
        {
            return r.IsDBNull(index) ? 0 : Convert.ToDouble(r.GetValue(index), CultureInfo.InvariantCulture);
        }

        private static bool B(SqlDataReader r, int index)
        {
            if (r.IsDBNull(index))
            {
                return false;
            }

            object value = r.GetValue(index);
            if (value is bool)
            {
                return (bool)value;
            }

            return Convert.ToInt32(value, CultureInfo.InvariantCulture) != 0;
        }

        private static string Csv(params object[] values)
        {
            StringBuilder line = new StringBuilder();
            for (int i = 0; i < values.Length; i++)
            {
                if (i > 0)
                {
                    line.Append(',');
                }

                string text = values[i] == null ? "" : Convert.ToString(values[i], CultureInfo.InvariantCulture);
                if (text.IndexOfAny(new[] { ',', '"', '\r', '\n' }) >= 0)
                {
                    line.Append('"');
                    line.Append(text.Replace("\"", "\"\""));
                    line.Append('"');
                }
                else
                {
                    line.Append(text);
                }
            }

            return line.ToString();
        }

        private static string SqlLiteral(string text)
        {
            return (text ?? "").Replace("'", "''");
        }
    }

    internal sealed class Crypto
    {
        private readonly object instance;
        private readonly MethodInfo decryptOld;
        private readonly MethodInfo encryptOld;

        public Crypto(string exePath)
        {
            Assembly asm = Assembly.LoadFrom(exePath);
            Type security = asm.GetType("RecoNet.Security", true);
            instance = Activator.CreateInstance(security, true);
            decryptOld = security.GetMethod("DecryptoOld", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, new[] { typeof(string) }, null);
            encryptOld = security.GetMethod("EncryptoOld", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, new[] { typeof(string) }, null);
            if (decryptOld == null || encryptOld == null)
            {
                throw new MissingMethodException("RecoNet.Security DecryptoOld/EncryptoOld not found.");
            }
        }

        public bool CanRoundTripOld()
        {
            string plain = "0000000030.388@0089990021.3";
            return String.Equals(DecryptOld(EncryptOld(plain)), plain, StringComparison.Ordinal);
        }

        public string DecryptOld(string text)
        {
            if (String.IsNullOrWhiteSpace(text))
            {
                return "";
            }

            return Convert.ToString(decryptOld.Invoke(instance, new object[] { text }), CultureInfo.InvariantCulture);
        }

        public string EncryptOld(string text)
        {
            if (String.IsNullOrWhiteSpace(text))
            {
                return "";
            }

            return Convert.ToString(encryptOld.Invoke(instance, new object[] { text }), CultureInfo.InvariantCulture);
        }
    }

    internal sealed class MigrationPlan
    {
        public DateTime BuiltAt;
        public string ReportDir;
        public List<BookIndex> SourceIndexes = new List<BookIndex>();
        public List<SectionRow> SourceSections = new List<SectionRow>();
        public List<QuotaRow> SourceQuotas = new List<QuotaRow>();
        public Dictionary<int, ResourceMapping> ResourceMap = new Dictionary<int, ResourceMapping>();
        public List<ResourceMapping> MissingResources = new List<ResourceMapping>();
        public List<ResourceMapping> UnitDifferences = new List<ResourceMapping>();
        public List<ResourceMapping> NameReviewMappings = new List<ResourceMapping>();
        public List<ResourceMapping> UnknownCodes = new List<ResourceMapping>();
        public List<string> Conflicts = new List<string>();
        public HashSet<string> TargetBooks = new HashSet<string>(StringComparer.Ordinal);
        public HashSet<string> TargetQuotaKeys = new HashSet<string>(StringComparer.Ordinal);
        public HashSet<string> TargetSectionKeys = new HashSet<string>(StringComparer.Ordinal);
        public Dictionary<int, Resource> SourceMaterials = new Dictionary<int, Resource>();
        public Dictionary<int, Resource> SourceMachines = new Dictionary<int, Resource>();
        public Dictionary<string, List<Resource>> TargetMaterials = new Dictionary<string, List<Resource>>();
        public Dictionary<string, List<Resource>> TargetMachines = new Dictionary<string, List<Resource>>();
        public long MaterialRangeStart;
        public long MaterialRangeEnd;
        public long MachineRangeStart;
        public long MachineRangeEnd;
        public long NextMaterialCode;
        public long NextMachineCode;

        public bool HasBlockingConflicts
        {
            get { return Conflicts.Count > 0 || UnknownCodes.Count > 0; }
        }

        public int TotalConsumeRows
        {
            get
            {
                int total = 0;
                foreach (QuotaRow row in SourceQuotas)
                {
                    total += row.ConsumeRows.Count;
                }
                return total;
            }
        }

        public int DirectMappings
        {
            get
            {
                int total = 0;
                foreach (ResourceMapping row in ResourceMap.Values)
                {
                    if (row.MatchStatus == "名称单位匹配" ||
                        row.MatchStatus == "名称匹配单位不同" ||
                        row.MatchStatus == "规范化名称匹配" ||
                        row.MatchStatus == "规范化名称匹配单位不同" ||
                        row.MatchStatus == "相似名称匹配" ||
                        row.MatchStatus == "相似名称匹配单位不同" ||
                        row.MatchStatus == "人工保留" ||
                        row.MatchStatus == "特殊人工保留")
                    {
                        total++;
                    }
                }
                return total;
            }
        }

        public int MissingResourcesWithoutCandidate
        {
            get
            {
                int total = 0;
                foreach (ResourceMapping row in MissingResources)
                {
                    if (row.CandidateCode == 0)
                    {
                        total++;
                    }
                }

                return total;
            }
        }

        public int LowConfidenceSupplementCandidates
        {
            get
            {
                int total = 0;
                foreach (ResourceMapping row in MissingResources)
                {
                    if (row.CandidateCode != 0 && row.CandidateScore < 60)
                    {
                        total++;
                    }
                }

                return total;
            }
        }

        public List<ResourceMapping> SortedResourceMap
        {
            get
            {
                List<ResourceMapping> rows = new List<ResourceMapping>(ResourceMap.Values);
                rows.Sort(delegate(ResourceMapping a, ResourceMapping b) { return a.OldCode.CompareTo(b.OldCode); });
                return rows;
            }
        }
    }

    internal sealed class BookIndex
    {
        public string Book;
        public string Head;
        public string Content;
        public string Prefix;
        public string ExclusiveTable;
        public string Category;
        public string Specialty;
        public int Order;
        public bool IsCurrent;
        public string Note;
    }

    internal sealed class SectionRow
    {
        public string Book;
        public string SectionNo;
        public string SectionName;
        public string Description;
        public string StartNo;
        public string EndNo;
    }

    internal sealed class QuotaRow
    {
        public string Book;
        public string Code;
        public string Name;
        public string Unit;
        public string ConsumeEncrypted;
        public string ConsumePlain;
        public string MigratedConsumePlain;
        public string MigratedConsumeEncrypted;
        public string WorkContent;
        public string BaseQuota;
        public double BasePrice;
        public double LaborFee;
        public double MaterialFee;
        public double MachineFee;
        public double UnitWeight;
        public string SectionNo;
        public int SortOrder;
        public bool Locked;
        public bool Use;
        public int ItemSequence;
        public DateTime? ImportedAt;
        public string Category;
        public List<ConsumeRow> ConsumeRows = new List<ConsumeRow>();
    }

    internal sealed class ConsumePart
    {
        public int OldCode;
        public string QuantityText;
    }

    internal sealed class ConsumeRow
    {
        public string Book;
        public string QuotaCode;
        public int Code;
        public string Quantity;
    }

    internal enum ResourceKind
    {
        Labor,
        Material,
        Machine,
        Unknown
    }

    internal sealed class Resource
    {
        public ResourceKind Kind;
        public string DocNo;
        public int OldCode;
        public int NewCode;
        public string Name;
        public string Unit;
        public double BasePrice;
        public double CurrentPrice;
        public double CurrentTaxPrice;
        public double UnitWeight;
        public double ConvertFactor;
        public string SummaryFlag;
        public string MainMaterialFlag;
        public string TransportCategory;
        public bool Locked;
        public string SupplyMode;
        public DateTime? ImportedAt;
        public double DepreciationFee;
        public double RepairFee;
        public double MaintenanceFee;
        public double InstallUninstallFee;
        public double Labor;
        public double Gasoline;
        public double Diesel;
        public double Coal;
        public double Electricity;
        public double Water;
        public double NaturalGas;
        public bool ContactLineBlock;
        public double RoadFeeFactor;
        public double OtherFee;

        public Resource Clone()
        {
            return (Resource)MemberwiseClone();
        }
    }

    internal sealed class ResourceMapping
    {
        public ResourceKind Kind;
        public int OldCode;
        public int NewCode;
        public Resource Source;
        public Resource Target;
        public string MatchStatus;
        public string MatchReason;
        public int MatchScore;
        public bool UnitDifference;
        public int CandidateCode;
        public string CandidateName;
        public string CandidateUnit;
        public int CandidateScore;

        public string Name
        {
            get { return Source != null ? Source.Name : (Target != null ? Target.Name : ""); }
        }

        public string SourceName
        {
            get { return Source == null ? "" : Source.Name; }
        }

        public string TargetName
        {
            get { return Target == null ? "" : Target.Name; }
        }

        public string SourceUnit
        {
            get { return Source == null ? "" : Source.Unit; }
        }

        public string TargetUnit
        {
            get { return Target == null ? SourceUnit : Target.Unit; }
        }
    }

    internal sealed class ResourceCandidate
    {
        public Resource Target;
        public int Score;
        public int SecondScore;
    }
}
