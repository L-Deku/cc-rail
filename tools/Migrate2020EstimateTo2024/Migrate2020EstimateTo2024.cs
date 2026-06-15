using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

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

        private static readonly Dictionary<int, int> ObsoleteAliasCodes = new Dictionary<int, int>
        {
            { 1990051, 1009003002 },
            { 5833051, 1009090001 }
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
                string exeOverride = GetArg(args, "--exe");
                if (!String.IsNullOrWhiteSpace(exeOverride))
                {
                    exePath = exeOverride;
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

                if (String.Equals(command, "VerifyRebuild", StringComparison.OrdinalIgnoreCase))
                {
                    MigrationPlan plan = BuildPlan(exePath, reportDir);
                    VerifyRebuild(plan, reportDir);
                    return 0;
                }

                if (String.Equals(command, "ExportEncryptRequests", StringComparison.OrdinalIgnoreCase))
                {
                    MigrationPlan plan = BuildPlan(exePath, reportDir);
                    string manualConflict = plan.Conflicts.FirstOrDefault(x => x.StartsWith("MANUAL,", StringComparison.Ordinal));
                    if (manualConflict != null)
                    {
                        throw new InvalidOperationException("Manual resource decision conflict: " + manualConflict);
                    }

                    string path = WriteEncryptRequests(plan, softwareDir);
                    Console.WriteLine("EncryptRequests: " + path);
                    Console.WriteLine("Quotas: " + plan.SourceQuotas.Count.ToString(CultureInfo.InvariantCulture));
                    Console.WriteLine("ConsumeRows: " + plan.TotalConsumeRows.ToString(CultureInfo.InvariantCulture));
                    Console.WriteLine("ManualDecisions: " + plan.ManualOverrideMappings.Count.ToString(CultureInfo.InvariantCulture));
                    Console.WriteLine("SupplementResources: " + plan.MissingResources.Count.ToString(CultureInfo.InvariantCulture));
                    return 0;
                }

                if (String.Equals(command, "ApplyEncryptResponses", StringComparison.OrdinalIgnoreCase))
                {
                    MigrationPlan plan = BuildPlan(exePath, reportDir);
                    ApplyEncryptResponses(softwareDir, plan);
                    Verify(reportDir);
                    return 0;
                }

                if (String.Equals(command, "DumpQuota", StringComparison.OrdinalIgnoreCase))
                {
                    string book = GetArg(args, "--book");
                    string code = GetArg(args, "--code");
                    if (String.IsNullOrWhiteSpace(book) || String.IsNullOrWhiteSpace(code))
                    {
                        throw new ArgumentException("DumpQuota requires --book and --code.");
                    }

                    DumpQuota(exePath, book, code);
                    return 0;
                }

                if (String.Equals(command, "InspectQuotaRows", StringComparison.OrdinalIgnoreCase))
                {
                    string book = GetArg(args, "--book");
                    string code = GetArg(args, "--code");
                    if (String.IsNullOrWhiteSpace(book) || String.IsNullOrWhiteSpace(code))
                    {
                        throw new ArgumentException("InspectQuotaRows requires --book and --code.");
                    }

                    InspectQuotaRows(book, code);
                    return 0;
                }

                if (String.Equals(command, "ListSecurity", StringComparison.OrdinalIgnoreCase))
                {
                    Crypto.ListSecurityMethods(exePath);
                    Crypto.ListSecurityFields(exePath);
                    return 0;
                }

                if (String.Equals(command, "TestCrypto", StringComparison.OrdinalIgnoreCase))
                {
                    string text = GetArg(args, "--text");
                    if (String.IsNullOrWhiteSpace(text))
                    {
                        throw new ArgumentException("TestCrypto requires --text.");
                    }

                    Crypto.TestSecurityMethods(exePath, text);
                    return 0;
                }

                if (String.Equals(command, "TestEncrypt", StringComparison.OrdinalIgnoreCase))
                {
                    string text = GetArg(args, "--text");
                    string ctorKey = GetArg(args, "--ctor");
                    if (String.IsNullOrWhiteSpace(text))
                    {
                        throw new ArgumentException("TestEncrypt requires --text.");
                    }

                    Crypto.TestEncryptMethods(exePath, text, ctorKey);
                    return 0;
                }

                if (String.Equals(command, "DumpIL", StringComparison.OrdinalIgnoreCase))
                {
                    string methodName = GetArg(args, "--method");
                    if (String.IsNullOrWhiteSpace(methodName))
                    {
                        throw new ArgumentException("DumpIL requires --method.");
                    }

                    Crypto.DumpSecurityIL(exePath, methodName);
                    return 0;
                }

                if (String.Equals(command, "KeyInfo", StringComparison.OrdinalIgnoreCase))
                {
                    string methodName = GetArg(args, "--method");
                    if (String.IsNullOrWhiteSpace(methodName))
                    {
                        throw new ArgumentException("KeyInfo requires --method.");
                    }

                    Crypto.PrintByteMethod(exePath, methodName);
                    return 0;
                }

                if (String.Equals(command, "FindSecurityCalls", StringComparison.OrdinalIgnoreCase))
                {
                    Crypto.FindSecurityCalls(exePath);
                    return 0;
                }

                if (String.Equals(command, "MatchEncrypt", StringComparison.OrdinalIgnoreCase))
                {
                    string plain = GetArg(args, "--plain");
                    string cipher = GetArg(args, "--cipher");
                    if (String.IsNullOrEmpty(plain) || String.IsNullOrEmpty(cipher))
                    {
                        throw new ArgumentException("MatchEncrypt requires --plain and --cipher.");
                    }

                    Crypto.MatchEncryptCandidates(exePath, plain, cipher);
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
            Console.WriteLine("  Migrate2020EstimateTo2024.exe VerifyRebuild");
            Console.WriteLine("  Migrate2020EstimateTo2024.exe ExportEncryptRequests");
            Console.WriteLine("  Migrate2020EstimateTo2024.exe ApplyEncryptResponses");
            Console.WriteLine("  Migrate2020EstimateTo2024.exe DumpQuota --book <book> --code <quota-code>");
            Console.WriteLine("  Migrate2020EstimateTo2024.exe InspectQuotaRows --book <book> --code <quota-code>");
            Console.WriteLine("  Migrate2020EstimateTo2024.exe ListSecurity");
            Console.WriteLine("  Migrate2020EstimateTo2024.exe TestCrypto --text <encrypted-text>");
            Console.WriteLine("  Migrate2020EstimateTo2024.exe TestEncrypt --text <plain-text> [--ctor <security-constructor-key>]");
            Console.WriteLine("  Migrate2020EstimateTo2024.exe DumpIL --method <security-method-name>");
            Console.WriteLine("  Migrate2020EstimateTo2024.exe KeyInfo --method <security-byte-method-name>");
            Console.WriteLine("  Migrate2020EstimateTo2024.exe FindSecurityCalls");
            Console.WriteLine("  Migrate2020EstimateTo2024.exe MatchEncrypt --plain <plain-text> --cipher <expected-cipher>");
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
                BuildTargetCodeLookups(plan);
                plan.SourceMaterials = LoadSourceMaterials(source);
                plan.SourceMachines = LoadSourceMachines(source);
                LoadRanges(target, plan);
                DetectConflicts(plan);
                BuildResourceMap(plan, crypto);
                ApplyManualMissingResourceOverrides(plan, reportDir);
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

        private static void BuildTargetCodeLookups(MigrationPlan plan)
        {
            foreach (List<Resource> list in plan.TargetMaterials.Values)
            {
                foreach (Resource resource in list)
                {
                    plan.TargetMaterialByCode[resource.NewCode] = resource;
                }
            }

            foreach (List<Resource> list in plan.TargetMachines.Values)
            {
                foreach (Resource resource in list)
                {
                    plan.TargetMachineByCode[resource.NewCode] = resource;
                }
            }
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
            if (mapping.Kind == ResourceKind.Material)
            {
                supplement.NewCode = AllocateMaterialCode(plan, mapping.OldCode);
                plan.TargetMaterialByCode[supplement.NewCode] = supplement;
            }
            else
            {
                supplement.NewCode = AllocateMachineCode(plan, mapping.OldCode);
                plan.TargetMachineByCode[supplement.NewCode] = supplement;
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
            plan.ResourceInsertMappings.Add(mapping);
        }

        private static void ApplyManualMissingResourceOverrides(MigrationPlan plan, string reportDir)
        {
            Dictionary<int, ManualResourceDecision> decisions = LoadManualDecisions(reportDir);
            if (decisions.Count == 0)
            {
                return;
            }

            foreach (ResourceMapping mapping in plan.ResourceMap.Values)
            {
                ManualResourceDecision decision;
                if (!decisions.TryGetValue(mapping.OldCode, out decision) || String.IsNullOrWhiteSpace(decision.ManualCode))
                {
                    continue;
                }

                long selectedCode;
                if (!TryParseManualCode(decision.ManualCode, out selectedCode))
                {
                    plan.Conflicts.Add("MANUAL," + mapping.OldCode.ToString(CultureInfo.InvariantCulture) + ",L列不是有效电算代号: " + decision.ManualCode);
                    continue;
                }

                plan.ResourceInsertMappings.Remove(mapping);
                plan.MissingResources.Remove(mapping);
                plan.UnitDifferences.Remove(mapping);
                plan.ManualOverrideMappings.Remove(mapping);

                int targetCode;
                if (selectedCode == 0)
                {
                    if (!TryConvertResourceCode(decision.SupplementCode, out targetCode))
                    {
                        plan.Conflicts.Add("MANUAL," + mapping.OldCode.ToString(CultureInfo.InvariantCulture) + ",L列为0但C列补充电算代号无效: " + decision.SupplementCode);
                        continue;
                    }

                    Resource supplement = ResolveTargetByCode(plan, mapping.Kind, targetCode);
                    if (supplement == null)
                    {
                        supplement = mapping.Source.Clone();
                        supplement.NewCode = targetCode;
                        if (!RegisterSupplementResource(plan, mapping.Kind, supplement))
                        {
                            plan.Conflicts.Add("MANUAL," + mapping.OldCode.ToString(CultureInfo.InvariantCulture) + ",C列补充电算代号已被其他资源占用: " + targetCode.ToString(CultureInfo.InvariantCulture));
                            continue;
                        }

                        mapping.Target = supplement;
                        mapping.NewCode = targetCode;
                        plan.ResourceInsertMappings.Add(mapping);
                    }
                    else if (!String.Equals(supplement.Name, mapping.Source.Name, StringComparison.Ordinal))
                    {
                        plan.Conflicts.Add("MANUAL," + mapping.OldCode.ToString(CultureInfo.InvariantCulture) + ",C列补充电算代号对应其他资源: " + targetCode.ToString(CultureInfo.InvariantCulture));
                        continue;
                    }
                    else
                    {
                        mapping.Target = supplement;
                        mapping.NewCode = targetCode;
                    }

                    mapping.ManualOverride = "0=使用C列补充电算代号";
                    mapping.MatchStatus = "人工审核补充资源";
                    mapping.MatchReason = mapping.ManualOverride;
                    mapping.MatchScore = 100;
                    mapping.UnitDifference = false;
                    plan.MissingResources.Add(mapping);
                    plan.ManualOverrideMappings.Add(mapping);
                    continue;
                }

                if (selectedCode == 1)
                {
                    if (!TryConvertResourceCode(decision.CandidateCode, out targetCode))
                    {
                        plan.Conflicts.Add("MANUAL," + mapping.OldCode.ToString(CultureInfo.InvariantCulture) + ",L列为1但G列best_candidate_code无效: " + decision.CandidateCode);
                        continue;
                    }

                    mapping.ManualOverride = "1=使用best_candidate";
                }
                else
                {
                    if (selectedCode > Int32.MaxValue || selectedCode < 0)
                    {
                        plan.Conflicts.Add("MANUAL," + mapping.OldCode.ToString(CultureInfo.InvariantCulture) + ",L列电算代号超出int范围: " + decision.ManualCode);
                        continue;
                    }

                    targetCode = Convert.ToInt32(selectedCode, CultureInfo.InvariantCulture);
                    mapping.ManualOverride = "L列指定电算代号";
                }

                Resource target = ResolveTargetByCode(plan, mapping.Kind, targetCode);
                if (target == null)
                {
                    plan.Conflicts.Add("MANUAL," + mapping.OldCode.ToString(CultureInfo.InvariantCulture) + ",L列电算代号在2024目标库未找到: " + targetCode.ToString(CultureInfo.InvariantCulture));
                    continue;
                }

                target = EnsureConsumeSafeTarget(plan, mapping, target);
                mapping.NewCode = target.NewCode;
                mapping.Target = target;
                mapping.MatchStatus = selectedCode == 1 ? "人工审核使用best_candidate" : "人工审核指定电算代号";
                mapping.MatchReason = mapping.ManualOverride;
                mapping.MatchScore = 100;
                mapping.UnitDifference = !String.Equals(mapping.SourceUnit, mapping.TargetUnit, StringComparison.Ordinal);
                if (mapping.UnitDifference)
                {
                    plan.UnitDifferences.Add(mapping);
                }
                plan.ManualOverrideMappings.Add(mapping);
            }
        }

        private static bool TryConvertResourceCode(string text, out int code)
        {
            long parsed;
            code = 0;
            if (!TryParseManualCode(text, out parsed) || parsed < 0 || parsed > Int32.MaxValue)
            {
                return false;
            }

            code = Convert.ToInt32(parsed, CultureInfo.InvariantCulture);
            return true;
        }

        private static bool RegisterSupplementResource(MigrationPlan plan, ResourceKind kind, Resource resource)
        {
            if (kind == ResourceKind.Material)
            {
                Resource existing;
                if (plan.TargetMaterialByCode.TryGetValue(resource.NewCode, out existing))
                {
                    return String.Equals(existing.Name, resource.Name, StringComparison.Ordinal);
                }

                plan.TargetMaterialByCode[resource.NewCode] = resource;
                return true;
            }

            if (kind == ResourceKind.Machine)
            {
                Resource existing;
                if (plan.TargetMachineByCode.TryGetValue(resource.NewCode, out existing))
                {
                    return String.Equals(existing.Name, resource.Name, StringComparison.Ordinal);
                }

                plan.TargetMachineByCode[resource.NewCode] = resource;
                return true;
            }

            return false;
        }

        private static Resource ResolveTargetByCode(MigrationPlan plan, ResourceKind kind, int code)
        {
            Resource target;
            if (kind == ResourceKind.Material && plan.TargetMaterialByCode.TryGetValue(code, out target))
            {
                return target;
            }

            if (kind == ResourceKind.Machine && plan.TargetMachineByCode.TryGetValue(code, out target))
            {
                return target;
            }

            return null;
        }

        private static bool TryParseManualCode(string raw, out long code)
        {
            code = 0;
            string text = (raw ?? "").Trim();
            if (String.IsNullOrEmpty(text))
            {
                return false;
            }

            decimal decimalValue;
            if (Decimal.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out decimalValue) ||
                Decimal.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out decimalValue))
            {
                code = Decimal.ToInt64(Decimal.Truncate(decimalValue));
                return true;
            }

            return false;
        }

        private static Dictionary<int, ManualResourceDecision> LoadManualDecisions(string reportDir)
        {
            string xlsx = Path.Combine(reportDir, "missing-resources.xlsx");
            if (File.Exists(xlsx))
            {
                return LoadManualDecisionsFromXlsx(xlsx);
            }

            string csv = Path.Combine(reportDir, "missing-resources.csv");
            if (File.Exists(csv))
            {
                return LoadManualDecisionsFromCsv(csv);
            }

            return new Dictionary<int, ManualResourceDecision>();
        }

        private static Dictionary<int, ManualResourceDecision> LoadManualDecisionsFromCsv(string path)
        {
            Dictionary<int, ManualResourceDecision> result = new Dictionary<int, ManualResourceDecision>();
            string[] lines = File.ReadAllLines(path, Encoding.UTF8);
            if (lines.Length < 2)
            {
                return result;
            }

            for (int i = 1; i < lines.Length; i++)
            {
                List<string> cells = SplitCsvLine(lines[i]);
                if (cells.Count < 11)
                {
                    continue;
                }

                int oldCode;
                if (Int32.TryParse(cells[1], NumberStyles.Integer, CultureInfo.InvariantCulture, out oldCode))
                {
                    result[oldCode] = new ManualResourceDecision
                    {
                        OldCode = oldCode,
                        SupplementCode = cells[2],
                        CandidateCode = cells[6],
                        ManualCode = cells.Count > 11 ? cells[11] : ""
                    };
                }
            }

            return result;
        }

        private static List<string> SplitCsvLine(string line)
        {
            List<string> cells = new List<string>();
            StringBuilder current = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < line.Length; i++)
            {
                char ch = line[i];
                if (quoted)
                {
                    if (ch == '"' && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        current.Append('"');
                        i++;
                    }
                    else if (ch == '"')
                    {
                        quoted = false;
                    }
                    else
                    {
                        current.Append(ch);
                    }
                }
                else if (ch == ',')
                {
                    cells.Add(current.ToString());
                    current.Length = 0;
                }
                else if (ch == '"')
                {
                    quoted = true;
                }
                else
                {
                    current.Append(ch);
                }
            }

            cells.Add(current.ToString());
            return cells;
        }

        private static Dictionary<int, ManualResourceDecision> LoadManualDecisionsFromXlsx(string path)
        {
            Dictionary<int, ManualResourceDecision> result = new Dictionary<int, ManualResourceDecision>();
            using (ZipArchive zip = ZipFile.OpenRead(path))
            {
                List<string> sharedStrings = LoadSharedStrings(zip);
                ZipArchiveEntry sheet = null;
                foreach (ZipArchiveEntry entry in zip.Entries)
                {
                    if (entry.FullName.StartsWith("xl/worksheets/sheet", StringComparison.OrdinalIgnoreCase) &&
                        entry.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                    {
                        sheet = entry;
                        break;
                    }
                }

                if (sheet == null)
                {
                    return result;
                }

                XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
                XDocument doc;
                using (Stream stream = sheet.Open())
                {
                    doc = XDocument.Load(stream);
                }

                foreach (XElement row in doc.Descendants(ns + "row"))
                {
                    int rowNumber = GetIntAttribute(row, "r");
                    if (rowNumber <= 1)
                    {
                        continue;
                    }

                    Dictionary<string, string> cells = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (XElement cell in row.Elements(ns + "c"))
                    {
                        XAttribute refAttr = cell.Attribute("r");
                        if (refAttr == null)
                        {
                            continue;
                        }

                        string column = GetColumnName(refAttr.Value);
                        cells[column] = ReadCellValue(cell, sharedStrings, ns);
                    }

                    int oldCode;
                    string oldCodeText;
                    string supplement;
                    string candidate;
                    string manual;
                    if (cells.TryGetValue("B", out oldCodeText) &&
                        cells.TryGetValue("C", out supplement) &&
                        cells.TryGetValue("G", out candidate) &&
                        cells.TryGetValue("L", out manual) &&
                        Int32.TryParse(NormalizeNumericText(oldCodeText), NumberStyles.Integer, CultureInfo.InvariantCulture, out oldCode))
                    {
                        result[oldCode] = new ManualResourceDecision
                        {
                            OldCode = oldCode,
                            SupplementCode = NormalizeNumericText(supplement),
                            CandidateCode = NormalizeNumericText(candidate),
                            ManualCode = NormalizeNumericText(manual)
                        };
                    }
                }
            }

            return result;
        }

        private static List<string> LoadSharedStrings(ZipArchive zip)
        {
            List<string> strings = new List<string>();
            ZipArchiveEntry entry = zip.GetEntry("xl/sharedStrings.xml");
            if (entry == null)
            {
                return strings;
            }

            XNamespace ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            XDocument doc;
            using (Stream stream = entry.Open())
            {
                doc = XDocument.Load(stream);
            }

            foreach (XElement item in doc.Descendants(ns + "si"))
            {
                StringBuilder text = new StringBuilder();
                foreach (XElement t in item.Descendants(ns + "t"))
                {
                    text.Append(t.Value);
                }

                strings.Add(text.ToString());
            }

            return strings;
        }

        private static int GetIntAttribute(XElement element, string name)
        {
            XAttribute attr = element.Attribute(name);
            int value;
            return attr != null && Int32.TryParse(attr.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out value) ? value : 0;
        }

        private static string GetColumnName(string cellRef)
        {
            StringBuilder column = new StringBuilder();
            foreach (char ch in cellRef)
            {
                if (Char.IsLetter(ch))
                {
                    column.Append(ch);
                }
                else
                {
                    break;
                }
            }

            return column.ToString();
        }

        private static string ReadCellValue(XElement cell, List<string> sharedStrings, XNamespace ns)
        {
            XAttribute type = cell.Attribute("t");
            XElement value = cell.Element(ns + "v");
            if (type != null && type.Value == "inlineStr")
            {
                XElement text = cell.Descendants(ns + "t").FirstOrDefault();
                return text == null ? "" : text.Value;
            }

            if (value == null)
            {
                return "";
            }

            string raw = value.Value;
            if (type != null && type.Value == "s")
            {
                int index;
                if (Int32.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out index) &&
                    index >= 0 && index < sharedStrings.Count)
                {
                    return sharedStrings[index];
                }

                return "";
            }

            return raw;
        }

        private static string NormalizeNumericText(string text)
        {
            decimal value;
            if (Decimal.TryParse((text ?? "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out value))
            {
                return Decimal.ToInt64(Decimal.Truncate(value)).ToString(CultureInfo.InvariantCulture);
            }

            return (text ?? "").Trim();
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

        private static Resource EnsureConsumeSafeTarget(MigrationPlan plan, ResourceMapping mapping, Resource chosen)
        {
            int canonicalCode;
            if (ObsoleteAliasCodes.TryGetValue(chosen.NewCode, out canonicalCode))
            {
                Resource canonical = ResolveTargetByCode(plan, mapping.Kind, canonicalCode);
                if (canonical != null)
                {
                    return canonical;
                }
            }

            return chosen;
        }

        private static void ApplyChosenTarget(MigrationPlan plan, ResourceMapping mapping, Resource chosen, string sameUnitStatus, string diffUnitStatus, int score)
        {
            chosen = EnsureConsumeSafeTarget(plan, mapping, chosen);
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

        private static int AllocateMaterialCode(MigrationPlan plan, int preferredCode)
        {
            if (preferredCode > 0 && preferredCode <= 999999999 && !plan.TargetMaterialByCode.ContainsKey(preferredCode))
            {
                return preferredCode;
            }

            while (plan.TargetMaterialByCode.ContainsKey(checked((int)plan.NextMaterialNineCode)))
            {
                plan.NextMaterialNineCode++;
            }

            if (plan.NextMaterialNineCode > 999999999)
            {
                throw new InvalidOperationException("No 9-digit material supplement code left.");
            }

            return checked((int)plan.NextMaterialNineCode++);
        }

        private static int AllocateMachineCode(MigrationPlan plan, int preferredCode)
        {
            if (preferredCode > 0 && preferredCode <= 999999999 && !plan.TargetMachineByCode.ContainsKey(preferredCode))
            {
                return preferredCode;
            }

            while (plan.TargetMachineByCode.ContainsKey(checked((int)plan.NextMachineNineCode)))
            {
                plan.NextMachineNineCode++;
            }

            if (plan.NextMachineNineCode > 999999999)
            {
                throw new InvalidOperationException("No 9-digit machine supplement code left.");
            }

            return checked((int)plan.NextMachineNineCode++);
        }

        private static void BuildMigratedConsumptions(MigrationPlan plan, Crypto crypto)
        {
            foreach (QuotaRow quota in plan.SourceQuotas)
            {
                List<ConsumePart> parts = ParseConsume(quota.ConsumePlain);
                Dictionary<int, decimal> quantities = new Dictionary<int, decimal>();
                List<int> order = new List<int>();
                StringBuilder plain = new StringBuilder();
                foreach (ConsumePart part in parts)
                {
                    ResourceMapping mapping;
                    if (!plan.ResourceMap.TryGetValue(part.OldCode, out mapping) || mapping.Kind == ResourceKind.Unknown)
                    {
                        plan.Conflicts.Add("CONSUME," + quota.Book + "/" + quota.Code + ",无法映射电算代号 " + part.OldCode.ToString(CultureInfo.InvariantCulture));
                        continue;
                    }

                    if (!quantities.ContainsKey(mapping.NewCode))
                    {
                        quantities[mapping.NewCode] = 0m;
                        order.Add(mapping.NewCode);
                    }

                    quantities[mapping.NewCode] += part.QuantityValue;
                }

                foreach (int code in order)
                {
                    string quantityText = FormatQuantity(quantities[code]);
                    if (plain.Length > 0)
                    {
                        plain.Append('@');
                    }

                    plain.Append(code.ToString("D10", CultureInfo.InvariantCulture));
                    plain.Append(quantityText);
                    quota.ConsumeRows.Add(new ConsumeRow
                    {
                        Book = quota.Book,
                        QuotaCode = quota.Code,
                        Code = code,
                        Quantity = quantityText
                    });
                }

                quota.MigratedConsumePlain = plain.ToString();
                quota.MigratedConsumeEncrypted = crypto.EncryptOld(quota.MigratedConsumePlain);
            }
        }

        private static string FormatQuantity(decimal value)
        {
            string text = value.ToString("0.#############################", CultureInfo.InvariantCulture);
            return text == "-0" ? "0" : text;
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

                rows.Add(new ConsumePart { OldCode = code, QuantityText = qty, QuantityValue = parsed });
            }

            return rows;
        }

        private static void WriteReports(MigrationPlan plan, string reportDir)
        {
            WriteSummary(plan, Path.Combine(reportDir, "precheck-summary.txt"));
            WriteResourceMap(plan, Path.Combine(reportDir, "resource-map.csv"));
            WriteMissingResources(plan, Path.Combine(reportDir, "missing-resources.csv"));
            WriteManualOverrides(plan, Path.Combine(reportDir, "manual-overrides.csv"));
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
            text.AppendLine("AliasResourcesForOver9DigitTargets: " + plan.AliasResourceMappings.Count.ToString(CultureInfo.InvariantCulture));
            text.AppendLine("ManualOverrideResources: " + plan.ManualOverrideMappings.Count.ToString(CultureInfo.InvariantCulture));
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

        private static void WriteManualOverrides(MigrationPlan plan, string path)
        {
            using (StreamWriter writer = new StreamWriter(path, false, Encoding.UTF8))
            {
                writer.WriteLine("kind,old_code,new_code,source_name,target_name,source_unit,target_unit,manual_rule");
                foreach (ResourceMapping row in plan.ManualOverrideMappings)
                {
                    writer.WriteLine(Csv(row.Kind.ToString(), row.OldCode, row.NewCode, row.SourceName, row.TargetName, row.SourceUnit, row.TargetUnit, row.ManualOverride));
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

                foreach (ResourceMapping mapping in plan.ResourceInsertMappings)
                {
                    if (mapping.Kind == ResourceKind.Material)
                    {
                        writer.WriteLine("delete from 材料单价库 where 电算代号=" + mapping.NewCode.ToString(CultureInfo.InvariantCulture) + " and 材料名称=N'" + SqlLiteral(mapping.TargetName) + "';");
                    }
                    else if (mapping.Kind == ResourceKind.Machine)
                    {
                        writer.WriteLine("delete from 台班定额库 where 电算代号=" + mapping.NewCode.ToString(CultureInfo.InvariantCulture) + " and 机械台班名称=N'" + SqlLiteral(mapping.TargetName) + "';");
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
            Console.WriteLine("AliasResourcesForOver9DigitTargets: " + plan.AliasResourceMappings.Count.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine("ManualOverrideResources: " + plan.ManualOverrideMappings.Count.ToString(CultureInfo.InvariantCulture));
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
                    foreach (ResourceMapping mapping in plan.ResourceInsertMappings)
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

        private static void VerifyRebuild(MigrationPlan plan, string reportDir)
        {
            Dictionary<string, decimal> expected = new Dictionary<string, decimal>(StringComparer.Ordinal);
            foreach (QuotaRow quota in plan.SourceQuotas)
            {
                foreach (ConsumeRow row in quota.ConsumeRows)
                {
                    string key = row.Book + "\t" + row.QuotaCode + "\t" + row.Code.ToString(CultureInfo.InvariantCulture);
                    expected[key] = Decimal.Parse(row.Quantity, NumberStyles.Float, CultureInfo.InvariantCulture);
                }
            }

            Dictionary<string, decimal> actual = new Dictionary<string, decimal>(StringComparer.Ordinal);
            int duplicateRows = 0;
            using (SqlConnection target = OpenDb(TargetDb))
            using (SqlCommand cmd = target.CreateCommand())
            {
                cmd.CommandTimeout = 300;
                cmd.CommandText = "select c.书号,c.定额编号,c.电算代号,c.消耗 from 定额库消耗 c join 定额库索引 b on c.书号=b.书号 where b.分类 in (N'概算定额',N'估算定额')";
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string key = S(reader, 0) + "\t" + S(reader, 1) + "\t" + I(reader, 2).ToString(CultureInfo.InvariantCulture);
                        decimal quantity = Convert.ToDecimal(reader.GetValue(3), CultureInfo.InvariantCulture);
                        if (actual.ContainsKey(key))
                        {
                            duplicateRows++;
                        }
                        else
                        {
                            actual[key] = quantity;
                        }
                    }
                }
            }

            List<string> mismatches = new List<string>();
            foreach (KeyValuePair<string, decimal> pair in expected)
            {
                decimal actualQuantity;
                if (!actual.TryGetValue(pair.Key, out actualQuantity))
                {
                    mismatches.Add("missing " + pair.Key);
                }
                else if (actualQuantity != pair.Value)
                {
                    mismatches.Add("quantity " + pair.Key + " expected=" + pair.Value.ToString(CultureInfo.InvariantCulture) + " actual=" + actualQuantity.ToString(CultureInfo.InvariantCulture));
                }
            }

            foreach (string key in actual.Keys)
            {
                if (!expected.ContainsKey(key))
                {
                    mismatches.Add("unexpected " + key);
                }
            }

            int missingResources;
            int obsoleteAliases;
            using (SqlConnection target = OpenDb(TargetDb))
            {
                missingResources = Count(target, "select count(*) from 定额库消耗 c join 定额库索引 b on c.书号=b.书号 where b.分类 in (N'概算定额',N'估算定额') and c.电算代号 not in (1,2,3,4,5,6,7,10,11) and not exists(select 1 from 材料单价库 m where m.电算代号=c.电算代号) and not exists(select 1 from 台班定额库 t where t.电算代号=c.电算代号)");
                obsoleteAliases = Count(target, "select (select count(*) from 材料单价库 where 电算代号 in (1990051,5833051))");
            }

            string verification = "VerifiedAt: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + Environment.NewLine
                + "ExpectedRows: " + expected.Count.ToString(CultureInfo.InvariantCulture) + Environment.NewLine
                + "ActualRows: " + actual.Count.ToString(CultureInfo.InvariantCulture) + Environment.NewLine
                + "ManualDecisions: " + plan.ManualOverrideMappings.Count.ToString(CultureInfo.InvariantCulture) + Environment.NewLine
                + "SupplementResources: " + plan.MissingResources.Count.ToString(CultureInfo.InvariantCulture) + Environment.NewLine
                + "DuplicateRows: " + duplicateRows.ToString(CultureInfo.InvariantCulture) + Environment.NewLine
                + "MismatchRows: " + mismatches.Count.ToString(CultureInfo.InvariantCulture) + Environment.NewLine
                + "MissingResourceRows: " + missingResources.ToString(CultureInfo.InvariantCulture) + Environment.NewLine
                + "ObsoleteAliasResources: " + obsoleteAliases.ToString(CultureInfo.InvariantCulture) + Environment.NewLine;
            Console.Write(verification);
            foreach (string mismatch in mismatches.Take(20))
            {
                Console.WriteLine("Mismatch: " + mismatch);
            }

            if (duplicateRows != 0 || mismatches.Count != 0 || missingResources != 0 || obsoleteAliases != 0)
            {
                throw new InvalidOperationException("Rebuild verification failed.");
            }

            WriteManualOverrides(plan, Path.Combine(reportDir, "applied-manual-decisions.csv"));
            File.WriteAllText(Path.Combine(reportDir, "rebuild-verification.txt"), verification, Encoding.UTF8);
        }

        private static string WriteEncryptRequests(MigrationPlan plan, string softwareDir)
        {
            string dataDir = Path.Combine(softwareDir, "RecoQuotaData");
            Directory.CreateDirectory(dataDir);
            string responsePath = Path.Combine(dataDir, "consume-encrypt-responses.tsv");
            if (File.Exists(responsePath))
            {
                File.Delete(responsePath);
            }

            string requestPath = Path.Combine(dataDir, "consume-encrypt-requests.tsv");
            string tempPath = requestPath + ".tmp";
            using (StreamWriter writer = new StreamWriter(tempPath, false, Encoding.UTF8))
            {
                writer.WriteLine("# ConsumeCryptoBridgeV1");
                foreach (QuotaRow row in plan.SourceQuotas)
                {
                    writer.WriteLine(row.Book + "\t" + row.Code + "\t" + Convert.ToBase64String(Encoding.UTF8.GetBytes(row.MigratedConsumePlain ?? "")));
                }
            }

            if (File.Exists(requestPath))
            {
                File.Delete(requestPath);
            }

            File.Move(tempPath, requestPath);
            return requestPath;
        }

        private static void ApplyEncryptResponses(string softwareDir, MigrationPlan plan)
        {
            string responsePath = Path.Combine(softwareDir, "RecoQuotaData", "consume-encrypt-responses.tsv");
            if (!File.Exists(responsePath))
            {
                throw new FileNotFoundException("Encrypt response file not found.", responsePath);
            }

            List<Tuple<string, string, string>> rows = new List<Tuple<string, string, string>>();
            List<string> errors = new List<string>();
            foreach (string line in File.ReadAllLines(responsePath, Encoding.UTF8))
            {
                if (String.IsNullOrWhiteSpace(line) || line.StartsWith("#", StringComparison.Ordinal))
                {
                    continue;
                }

                string[] parts = line.Split('\t');
                if (parts.Length < 4)
                {
                    continue;
                }

                string book = parts[0];
                string code = parts[1];
                string status = parts[2];
                string payload = Encoding.UTF8.GetString(Convert.FromBase64String(parts[3]));
                if (!String.Equals(status, "OK", StringComparison.Ordinal))
                {
                    errors.Add(book + "/" + code + ": " + payload);
                    continue;
                }

                rows.Add(Tuple.Create(book, code, payload));
            }

            if (errors.Count > 0)
            {
                throw new InvalidOperationException("Encrypt bridge returned errors. First: " + errors[0]);
            }

            if (rows.Count != 7385)
            {
                throw new InvalidOperationException("Expected 7385 encrypted consume rows, got " + rows.Count.ToString(CultureInfo.InvariantCulture));
            }

            HashSet<string> responseKeys = new HashSet<string>(StringComparer.Ordinal);
            foreach (Tuple<string, string, string> row in rows)
            {
                responseKeys.Add(row.Item1 + "\t" + row.Item2);
            }

            foreach (QuotaRow quota in plan.SourceQuotas)
            {
                if (!responseKeys.Contains(quota.Book + "\t" + quota.Code))
                {
                    throw new InvalidOperationException("Encrypt response missing quota " + quota.Book + "/" + quota.Code);
                }
            }

            using (SqlConnection target = OpenDb(TargetDb))
            using (SqlTransaction tx = target.BeginTransaction())
            {
                try
                {
                    foreach (ResourceMapping mapping in plan.ResourceInsertMappings)
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
                        using (SqlCommand delete = target.CreateCommand())
                        {
                            delete.Transaction = tx;
                            delete.CommandText = "delete from 定额库消耗 where 书号=@book";
                            Add(delete, "@book", index.Book);
                            delete.ExecuteNonQuery();
                        }
                    }

                    foreach (QuotaRow quota in plan.SourceQuotas)
                    {
                        foreach (ConsumeRow consume in quota.ConsumeRows)
                        {
                            InsertConsume(target, tx, consume);
                        }
                    }

                    foreach (Tuple<string, string, string> row in rows)
                    {
                        using (SqlCommand cmd = target.CreateCommand())
                        {
                            cmd.Transaction = tx;
                            cmd.CommandText = "update 定额库 set 消耗=@consume where 书号=@book and 定额编号=@code";
                            Add(cmd, "@consume", row.Item3);
                            Add(cmd, "@book", row.Item1);
                            Add(cmd, "@code", row.Item2);
                            int affected = cmd.ExecuteNonQuery();
                            if (affected != 1)
                            {
                                throw new InvalidOperationException("Unexpected update count " + affected.ToString(CultureInfo.InvariantCulture) + " for " + row.Item1 + "/" + row.Item2);
                            }
                        }
                    }

                    DeleteObsoleteMigrationResources(target, tx, plan);
                    tx.Commit();
                }
                catch
                {
                    tx.Rollback();
                    throw;
                }
            }

            Console.WriteLine("Applied encrypted consume responses: " + rows.Count.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine("Rebuilt consume detail rows: " + plan.TotalConsumeRows.ToString(CultureInfo.InvariantCulture));
        }

        private static void DeleteObsoleteMigrationResources(SqlConnection target, SqlTransaction tx, MigrationPlan plan)
        {
            DeleteMaterialIfUnused(target, tx, 1990051, "钢件防腐处理");
            DeleteMaterialIfUnused(target, tx, 5833051, "钢芯铝绞线");

            foreach (ResourceMapping mapping in plan.ManualOverrideMappings)
            {
                if (!String.Equals(mapping.ManualOverride, "0=使用C列补充电算代号", StringComparison.Ordinal) ||
                    mapping.OldCode == mapping.NewCode)
                {
                    continue;
                }

                if (mapping.Kind == ResourceKind.Material)
                {
                    DeleteMaterialIfUnused(target, tx, mapping.OldCode, mapping.SourceName);
                }
                else if (mapping.Kind == ResourceKind.Machine)
                {
                    DeleteMachineIfUnused(target, tx, mapping.OldCode, mapping.SourceName);
                }
            }
        }

        private static void DeleteMaterialIfUnused(SqlConnection target, SqlTransaction tx, int code, string name)
        {
            using (SqlCommand cmd = target.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "delete from 材料单价库 where 电算代号=@code and 材料名称=@name and not exists(select 1 from 定额库消耗 where 电算代号=@code)";
                Add(cmd, "@code", code);
                Add(cmd, "@name", name);
                cmd.ExecuteNonQuery();
            }
        }

        private static void DeleteMachineIfUnused(SqlConnection target, SqlTransaction tx, int code, string name)
        {
            using (SqlCommand cmd = target.CreateCommand())
            {
                cmd.Transaction = tx;
                cmd.CommandText = "delete from 台班定额库 where 电算代号=@code and 机械台班名称=@name and not exists(select 1 from 定额库消耗 where 电算代号=@code)";
                Add(cmd, "@code", code);
                Add(cmd, "@name", name);
                cmd.ExecuteNonQuery();
            }
        }

        private static void DumpQuota(string exePath, string book, string code)
        {
            Crypto crypto = new Crypto(exePath);
            using (SqlConnection target = OpenDb(TargetDb))
            using (SqlCommand cmd = target.CreateCommand())
            {
                cmd.CommandText = "select convert(nvarchar(max),消耗) from 定额库 where 书号=@book and 定额编号=@code";
                Add(cmd, "@book", book);
                Add(cmd, "@code", code);
                object value = cmd.ExecuteScalar();
                if (value == null || value == DBNull.Value)
                {
                    Console.WriteLine("Quota not found: " + book + "/" + code);
                    return;
                }

            string encrypted = Convert.ToString(value, CultureInfo.InvariantCulture);
                string plain;
                string decryptMode = "2024";
                try
                {
                    plain = crypto.Decrypt(encrypted);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Decrypt2024Error: " + ex.GetBaseException().Message);
                    decryptMode = "Old";
                    plain = crypto.DecryptOld(encrypted);
                }

                Console.WriteLine("Book: " + book);
                Console.WriteLine("Code: " + code);
                Console.WriteLine("DecryptMode: " + decryptMode);
                Console.WriteLine("EncryptedLength: " + encrypted.Length.ToString(CultureInfo.InvariantCulture));
                Console.WriteLine("PlainLength: " + plain.Length.ToString(CultureInfo.InvariantCulture));
                Console.WriteLine("PlainHead: " + (plain.Length > 500 ? plain.Substring(0, 500) : plain));

                string[] items = plain.Split(new[] { '@' }, StringSplitOptions.RemoveEmptyEntries);
                Console.WriteLine("PlainItems: " + items.Length.ToString(CultureInfo.InvariantCulture));
                for (int i = 0; i < Math.Min(items.Length, 15); i++)
                {
                    string item = items[i];
                    Console.WriteLine("Item" + i.ToString(CultureInfo.InvariantCulture)
                        + ": len=" + item.Length.ToString(CultureInfo.InvariantCulture)
                        + ", first9=" + (item.Length >= 9 ? item.Substring(0, 9) : item)
                        + ", first10=" + (item.Length >= 10 ? item.Substring(0, 10) : item)
                        + ", raw=" + item);
                }

                try
                {
                    List<ConsumePart> parsed = ParseConsume(plain);
                    Console.WriteLine("ParseAs9DigitCount: " + parsed.Count.ToString(CultureInfo.InvariantCulture));
                    int over9 = 0;
                    foreach (string item in items)
                    {
                        if (item.Length >= 10 && Char.IsDigit(item[9]))
                        {
                            over9++;
                        }
                    }

                    Console.WriteLine("ItemsWithDigitAt10thChar: " + over9.ToString(CultureInfo.InvariantCulture));
                }
                catch (Exception ex)
                {
                    Console.WriteLine("ParseAs9DigitError: " + ex.Message);
                }
            }
        }

        private static void InspectQuotaRows(string book, string code)
        {
            using (SqlConnection target = OpenDb(TargetDb))
            {
                Console.WriteLine("Book: " + book);
                Console.WriteLine("Code: " + code);
                Console.WriteLine("DetailRows: " + CountForQuota(target, "select count(*) from 定额库消耗 where 书号=@book and 定额编号=@code", book, code).ToString(CultureInfo.InvariantCulture));
                Console.WriteLine("TenDigitRows: " + CountForQuota(target, "select count(*) from 定额库消耗 where 书号=@book and 定额编号=@code and 电算代号>=1000000000", book, code).ToString(CultureInfo.InvariantCulture));
                Console.WriteLine("DuplicateCodeGroups: " + CountForQuota(target, "select count(*) from (select 电算代号 from 定额库消耗 where 书号=@book and 定额编号=@code group by 电算代号 having count(*)>1) x", book, code).ToString(CultureInfo.InvariantCulture));
                Console.WriteLine("MissingResourceRows: " + CountForQuota(target, "select count(*) from 定额库消耗 c where c.书号=@book and c.定额编号=@code and c.电算代号 not in (1,2,3,4,5,6,7,10,11) and not exists(select 1 from 材料单价库 m where m.电算代号=c.电算代号) and not exists(select 1 from 台班定额库 t where t.电算代号=c.电算代号)", book, code).ToString(CultureInfo.InvariantCulture));

                using (SqlCommand cmd = target.CreateCommand())
                {
                    cmd.CommandText = "select top 1 len(convert(nvarchar(max),消耗)) as ConsumeLength,left(convert(nvarchar(max),消耗),32) as ConsumeHead from 定额库 where 书号=@book and 定额编号=@code";
                    Add(cmd, "@book", book);
                    Add(cmd, "@code", code);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        if (reader.Read())
                        {
                            Console.WriteLine("EncryptedConsumeLength: " + Convert.ToString(reader["ConsumeLength"], CultureInfo.InvariantCulture));
                            Console.WriteLine("EncryptedConsumeHead: " + Convert.ToString(reader["ConsumeHead"], CultureInfo.InvariantCulture));
                        }
                    }
                }

                using (SqlCommand cmd = target.CreateCommand())
                {
                    cmd.CommandText = @"
select top 30
  c.电算代号,
  c.消耗,
  case
    when c.电算代号 between 1 and 7 then N'人工'
    when m.电算代号 is not null then m.材料名称
    when t.电算代号 is not null then t.机械台班名称
    else N'(未找到资源)'
  end as 资源名称,
  case
    when c.电算代号 between 1 and 7 then N'工日'
    when m.电算代号 is not null then m.单位
    when t.电算代号 is not null then N'台班'
    else N''
  end as 单位
from 定额库消耗 c
left join 材料单价库 m on m.电算代号=c.电算代号
left join 台班定额库 t on t.电算代号=c.电算代号
where c.书号=@book and c.定额编号=@code
order by
  case
    when c.电算代号 between 1 and 7 then 0
    when m.电算代号 is not null then 1
    when t.电算代号 is not null then 2
    else 3
  end,
  c.电算代号";
                    Add(cmd, "@book", book);
                    Add(cmd, "@code", code);
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        int index = 0;
                        while (reader.Read())
                        {
                            Console.WriteLine(
                                "Row" + index.ToString(CultureInfo.InvariantCulture)
                                + ": code=" + Convert.ToString(reader["电算代号"], CultureInfo.InvariantCulture)
                                + ", qty=" + Convert.ToString(reader["消耗"], CultureInfo.InvariantCulture)
                                + ", name=" + Convert.ToString(reader["资源名称"], CultureInfo.InvariantCulture)
                                + ", unit=" + Convert.ToString(reader["单位"], CultureInfo.InvariantCulture));
                            index++;
                        }
                    }
                }
            }
        }

        private static int CountForQuota(SqlConnection conn, string sql, string book, string code)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = sql;
                Add(cmd, "@book", book);
                Add(cmd, "@code", code);
                return Convert.ToInt32(cmd.ExecuteScalar(), CultureInfo.InvariantCulture);
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
                Add(cmd, "@time", row.ImportedAt.HasValue ? (object)row.ImportedAt.Value : DBNull.Value);
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
                Add(cmd, "@time", row.ImportedAt.HasValue ? (object)row.ImportedAt.Value : DBNull.Value);
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
                Add(cmd, "@time", row.ImportedAt.HasValue ? (object)row.ImportedAt.Value : DBNull.Value);
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
        private readonly MethodInfo decrypt;
        private readonly MethodInfo encrypt;
        private readonly MethodInfo decryptOld;
        private readonly MethodInfo encryptOld;

        public static void ListSecurityMethods(string exePath)
        {
            Assembly asm = Assembly.LoadFrom(exePath);
            Type security = asm.GetType("RecoNet.Security", true);
            foreach (MethodInfo method in security.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static))
            {
                ParameterInfo[] parameters = method.GetParameters();
                StringBuilder signature = new StringBuilder();
                signature.Append(method.IsStatic ? "static " : "instance ");
                signature.Append(method.ReturnType.FullName);
                signature.Append(' ');
                signature.Append(method.Name);
                signature.Append('(');
                for (int i = 0; i < parameters.Length; i++)
                {
                    if (i > 0)
                    {
                        signature.Append(", ");
                    }

                    signature.Append(parameters[i].ParameterType.FullName);
                    signature.Append(' ');
                    signature.Append(parameters[i].Name);
                }

                signature.Append(')');
                Console.WriteLine(signature.ToString());
            }
        }

        public static void ListSecurityFields(string exePath)
        {
            Assembly asm = Assembly.LoadFrom(exePath);
            Type security = asm.GetType("RecoNet.Security", true);
            Console.WriteLine("Fields:");
            foreach (FieldInfo field in security.GetFields(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static))
            {
                Console.WriteLine((field.IsStatic ? "static " : "instance ") + field.FieldType.FullName + " " + field.Name);
            }

            Console.WriteLine("Constructors:");
            foreach (ConstructorInfo ctor in security.GetConstructors(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance))
            {
                Console.WriteLine(Convert.ToString(ctor, CultureInfo.InvariantCulture));
            }
        }

        public static void TestSecurityMethods(string exePath, string text)
        {
            Assembly asm = Assembly.LoadFrom(exePath);
            Type security = asm.GetType("RecoNet.Security", true);
            object obj = Activator.CreateInstance(security, true);
            foreach (string name in new[] { "Decrypto", "DecryptoFile", "DecryptoOld", "DecryptoFiletest" })
            {
                MethodInfo method = security.GetMethod(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, new[] { typeof(string) }, null);
                if (method == null)
                {
                    Console.WriteLine(name + ": missing");
                    continue;
                }

                try
                {
                    string value = Convert.ToString(method.Invoke(obj, new object[] { text }), CultureInfo.InvariantCulture);
                    Console.WriteLine(name + ": OK len=" + value.Length.ToString(CultureInfo.InvariantCulture)
                        + " head=" + (value.Length > 300 ? value.Substring(0, 300) : value));
                }
                catch (Exception ex)
                {
                    Console.WriteLine(name + ": ERROR " + ex.GetBaseException().Message);
                }
            }

            Dictionary<string, byte[]> keys = new Dictionary<string, byte[]>();
            foreach (string name in new[] { "GetLegalKey", "GetLegalKeyFile", "GetLegalKeytest", "GetLegalKeyOld" })
            {
                byte[] bytes = ReadByteMethod(security, obj, name);
                if (bytes != null)
                {
                    keys[name] = bytes;
                    Console.WriteLine(name + ": bytes=" + bytes.Length.ToString(CultureInfo.InvariantCulture)
                        + " base64=" + Convert.ToBase64String(bytes));
                }
            }

            Dictionary<string, byte[]> ivs = new Dictionary<string, byte[]>();
            foreach (string name in new[] { "GetLegalIV", "GetLegalIVFile", "GetLegalIVtest", "GetLegalIVOld" })
            {
                byte[] bytes = ReadByteMethod(security, obj, name);
                if (bytes != null)
                {
                    ivs[name] = bytes;
                    Console.WriteLine(name + ": bytes=" + bytes.Length.ToString(CultureInfo.InvariantCulture)
                        + " base64=" + Convert.ToBase64String(bytes));
                }
            }

            foreach (KeyValuePair<string, byte[]> key in keys)
            {
                foreach (KeyValuePair<string, byte[]> iv in ivs)
                {
                    TryManualDecrypt(text, key.Key, key.Value, iv.Key, iv.Value);
                }
            }
        }

        private static byte[] ReadByteMethod(Type security, object obj, string name)
        {
            MethodInfo method = security.GetMethod(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, Type.EmptyTypes, null);
            if (method == null)
            {
                return null;
            }

            try
            {
                return method.Invoke(obj, null) as byte[];
            }
            catch (Exception ex)
            {
                Console.WriteLine(name + ": ERROR " + ex.GetBaseException().Message);
                return null;
            }
        }

        private static void TryManualDecrypt(string text, string keyName, byte[] key, string ivName, byte[] iv)
        {
            try
            {
                byte[] cipher = Convert.FromBase64String(text);
                using (System.Security.Cryptography.RijndaelManaged rijndael = new System.Security.Cryptography.RijndaelManaged())
                {
                    rijndael.Key = key;
                    rijndael.IV = iv;
                    rijndael.Mode = System.Security.Cryptography.CipherMode.CBC;
                    rijndael.Padding = System.Security.Cryptography.PaddingMode.PKCS7;
                    using (System.Security.Cryptography.ICryptoTransform transform = rijndael.CreateDecryptor())
                    {
                        byte[] plainBytes = transform.TransformFinalBlock(cipher, 0, cipher.Length);
                        string value = Encoding.UTF8.GetString(plainBytes);
                        Console.WriteLine("Manual " + keyName + "/" + ivName + ": OK len="
                            + value.Length.ToString(CultureInfo.InvariantCulture)
                            + " head=" + (value.Length > 300 ? value.Substring(0, 300) : value));
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Manual " + keyName + "/" + ivName + ": ERROR " + ex.GetBaseException().Message);
            }
        }

        public static void TestEncryptMethods(string exePath, string text, string ctorKey)
        {
            Assembly asm = Assembly.LoadFrom(exePath);
            Type security = asm.GetType("RecoNet.Security", true);
            object obj;
            if (String.IsNullOrEmpty(ctorKey))
            {
                obj = Activator.CreateInstance(security, true);
            }
            else
            {
                obj = Activator.CreateInstance(security, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, new object[] { ctorKey }, CultureInfo.InvariantCulture);
                Console.WriteLine("ctor: " + ctorKey);
            }
            foreach (string fieldName in new[] { "m_strKey", "isOk" })
            {
                FieldInfo field = security.GetField(fieldName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance);
                if (field != null)
                {
                    object fieldValue = field.GetValue(obj);
                    Console.WriteLine("field " + fieldName + ": " + Convert.ToString(fieldValue, CultureInfo.InvariantCulture));
                }
            }
            MethodInfo getId = security.GetMethod("getID", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, new[] { typeof(string) }, null);
            if (getId != null)
            {
                try
                {
                    object id = getId.Invoke(obj, new object[] { text });
                    Console.WriteLine("getID: " + Convert.ToString(id, CultureInfo.InvariantCulture));
                }
                catch (Exception ex)
                {
                    Console.WriteLine("getID: ERROR " + ex.GetBaseException().Message);
                }
            }

            List<string> keyTexts = new List<string>();
            MethodInfo getKey = security.GetMethod("getKey", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, new[] { typeof(bool) }, null);
            if (getKey != null)
            {
                foreach (bool flag in new[] { true, false })
                {
                    try
                    {
                        string keyValue = Convert.ToString(getKey.Invoke(obj, new object[] { flag }), CultureInfo.InvariantCulture);
                        keyTexts.Add(keyValue);
                        Console.WriteLine("getKey(" + flag.ToString(CultureInfo.InvariantCulture) + "): " + keyValue);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("getKey(" + flag.ToString(CultureInfo.InvariantCulture) + "): ERROR " + ex.GetBaseException().Message);
                    }
                }
            }

            foreach (string name in new[] { "Encrypto", "EncryptoFile", "EncryptoOld" })
            {
                MethodInfo method = security.GetMethod(name, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, new[] { typeof(string) }, null);
                if (method == null)
                {
                    Console.WriteLine(name + ": missing");
                    continue;
                }

                try
                {
                    string value = Convert.ToString(method.Invoke(obj, new object[] { text }), CultureInfo.InvariantCulture);
                    Console.WriteLine(name + ": OK len=" + value.Length.ToString(CultureInfo.InvariantCulture)
                        + " value=" + value);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(name + ": ERROR " + ex.GetBaseException().Message);
                }
            }

            MethodInfo encrypt3 = security.GetMethod("Encrypto", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, new[] { typeof(string), typeof(string), typeof(string) }, null);
            MethodInfo decrypt3 = security.GetMethod("Decrypto", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, new[] { typeof(string), typeof(string), typeof(string) }, null);
            string[] candidates = keyTexts.Count == 0 ? new[] { "", " ", "1234567890123456" } : keyTexts.ToArray();
            foreach (string key in candidates)
            {
                foreach (string iv in candidates)
                {
                    if (encrypt3 != null)
                    {
                        try
                        {
                            string value = Convert.ToString(encrypt3.Invoke(obj, new object[] { text, key, iv }), CultureInfo.InvariantCulture);
                            Console.WriteLine("Encrypto3 key=" + key + " iv=" + iv + ": OK len=" + value.Length.ToString(CultureInfo.InvariantCulture) + " value=" + value);
                            if (decrypt3 != null)
                            {
                                string roundtrip = Convert.ToString(decrypt3.Invoke(obj, new object[] { value, key, iv }), CultureInfo.InvariantCulture);
                                Console.WriteLine("Decrypto3 roundtrip: " + roundtrip);
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Encrypto3 key=" + key + " iv=" + iv + ": ERROR " + ex.GetBaseException().Message);
                        }
                    }
                }
            }
        }

        public static void DumpSecurityIL(string exePath, string methodName)
        {
            Assembly asm = Assembly.LoadFrom(exePath);
            Type security = asm.GetType("RecoNet.Security", true);
            Dictionary<ushort, OpCode> opcodes = new Dictionary<ushort, OpCode>();
            foreach (FieldInfo field in typeof(OpCodes).GetFields(BindingFlags.Public | BindingFlags.Static))
            {
                if (field.FieldType == typeof(OpCode))
                {
                    OpCode op = (OpCode)field.GetValue(null);
                    opcodes[(ushort)op.Value] = op;
                }
            }

            foreach (MethodInfo method in security.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static))
            {
                if (!String.Equals(method.Name, methodName, StringComparison.Ordinal))
                {
                    continue;
                }

                Console.WriteLine("Method: " + method);
                MethodBody body = method.GetMethodBody();
                if (body == null)
                {
                    Console.WriteLine("No method body.");
                    continue;
                }

                byte[] il = body.GetILAsByteArray();
                int pos = 0;
                while (pos < il.Length)
                {
                    int offset = pos;
                    ushort value = il[pos++];
                    if (value == 0xfe)
                    {
                        value = (ushort)(0xfe00 | il[pos++]);
                    }

                    OpCode op;
                    if (!opcodes.TryGetValue(value, out op))
                    {
                        Console.WriteLine(offset.ToString("X4", CultureInfo.InvariantCulture) + ": <unknown>");
                        continue;
                    }

                    string operand = ReadOperand(method.Module, il, ref pos, op, offset);
                    Console.WriteLine(offset.ToString("X4", CultureInfo.InvariantCulture) + ": " + op.Name + (operand.Length == 0 ? "" : " " + operand));
                }
            }
        }

        public static void PrintByteMethod(string exePath, string methodName)
        {
            Assembly asm = Assembly.LoadFrom(exePath);
            Type security = asm.GetType("RecoNet.Security", true);
            object obj = Activator.CreateInstance(security, true);
            MethodInfo method = security.GetMethod(methodName, BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, Type.EmptyTypes, null);
            if (method == null)
            {
                Console.WriteLine(methodName + ": missing");
                return;
            }

            object value = method.Invoke(obj, null);
            byte[] bytes = value as byte[];
            if (bytes == null)
            {
                Console.WriteLine(methodName + ": non-byte result " + Convert.ToString(value, CultureInfo.InvariantCulture));
                return;
            }

            Console.WriteLine(methodName + ": bytes=" + bytes.Length.ToString(CultureInfo.InvariantCulture)
                + " base64=" + Convert.ToBase64String(bytes)
                + " hex=" + BitConverter.ToString(bytes).Replace("-", ""));
        }

        public static void FindSecurityCalls(string exePath)
        {
            string dir = Path.GetDirectoryName(exePath);
            foreach (string path in Directory.GetFiles(dir, "*.exe").Concat(Directory.GetFiles(dir, "*.dll")))
            {
                string name = Path.GetFileName(path);
                if (!(name.StartsWith("Reco", StringComparison.OrdinalIgnoreCase) ||
                      name.StartsWith("ReJJ", StringComparison.OrdinalIgnoreCase) ||
                      name.StartsWith("Rejj", StringComparison.OrdinalIgnoreCase) ||
                      name.StartsWith("Rejg", StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                FindSecurityCallsInAssembly(path);
            }
        }

        public static void MatchEncryptCandidates(string exePath, string plain, string expectedCipher)
        {
            string raw = Environment.GetEnvironmentVariable("MIGRATE_CRYPTO_CANDIDATES") ?? "";
            string[] candidates = raw.Split(new[] { '|' }, StringSplitOptions.RemoveEmptyEntries);
            Console.WriteLine("Candidates: " + candidates.Length.ToString(CultureInfo.InvariantCulture));
            if (candidates.Length == 0)
            {
                return;
            }

            Assembly asm = Assembly.LoadFrom(exePath);
            Type security = asm.GetType("RecoNet.Security", true);
            object obj = Activator.CreateInstance(security, true);
            MethodInfo encrypt3 = security.GetMethod("Encrypto", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, new[] { typeof(string), typeof(string), typeof(string) }, null);
            if (encrypt3 == null)
            {
                Console.WriteLine("Encrypto3 missing");
                return;
            }

            int attempts = 0;
            bool matched = false;
            List<Tuple<string, string, int, int>> pairs = new List<Tuple<string, string, int, int>>();
            for (int i = 0; i < candidates.Length; i++)
            {
                string c = candidates[i];
                int sep = c.IndexOf("+++", StringComparison.Ordinal);
                if (sep >= 0)
                {
                    pairs.Add(Tuple.Create(c.Substring(0, sep), c.Substring(sep + 3), i, i));
                }

                pairs.Add(Tuple.Create(c, c, i, i));
                pairs.Add(Tuple.Create(c, "", i, -1));
                pairs.Add(Tuple.Create("", c, -1, i));
            }

            for (int i = 0; i < candidates.Length; i++)
            {
                for (int j = 0; j < candidates.Length; j++)
                {
                    pairs.Add(Tuple.Create(candidates[i], candidates[j], i, j));
                }
            }

            foreach (Tuple<string, string, int, int> pair in pairs)
            {
                attempts++;
                try
                {
                    string value = Convert.ToString(encrypt3.Invoke(obj, new object[] { plain, pair.Item1, pair.Item2 }), CultureInfo.InvariantCulture);
                    if (String.Equals(value, expectedCipher, StringComparison.Ordinal))
                    {
                        matched = true;
                        Console.WriteLine("MATCH keyIndex=" + pair.Item3.ToString(CultureInfo.InvariantCulture)
                            + " ivIndex=" + pair.Item4.ToString(CultureInfo.InvariantCulture)
                            + " keyLength=" + pair.Item1.Length.ToString(CultureInfo.InvariantCulture)
                            + " ivLength=" + pair.Item2.Length.ToString(CultureInfo.InvariantCulture));
                    }
                }
                catch
                {
                }
            }

            Console.WriteLine("Attempts: " + attempts.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine("Matched: " + (matched ? "1" : "0"));
        }

        private static void FindSecurityCallsInAssembly(string path)
        {
            Assembly asm;
            try
            {
                asm = Assembly.LoadFrom(path);
            }
            catch
            {
                return;
            }

            Type[] types;
            try
            {
                types = asm.GetTypes();
            }
            catch
            {
                return;
            }

            foreach (Type type in types)
            {
                foreach (MethodBase method in type.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static).Cast<MethodBase>()
                    .Concat(type.GetConstructors(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static)))
                {
                    MethodBody body = null;
                    try
                    {
                        body = method.GetMethodBody();
                    }
                    catch
                    {
                        continue;
                    }

                    if (body == null)
                    {
                        continue;
                    }

                    byte[] il = body.GetILAsByteArray();
                    int pos = 0;
                    while (pos < il.Length)
                    {
                        int offset = pos;
                        ushort value = il[pos++];
                        if (value == 0xfe)
                        {
                            value = (ushort)(0xfe00 | il[pos++]);
                        }

                        OpCode op = ResolveOpCode(value);
                        if (op.OperandType == OperandType.InlineMethod || op.OperandType == OperandType.InlineTok)
                        {
                            int token = BitConverter.ToInt32(il, pos);
                            pos += 4;
                            try
                            {
                                MemberInfo member = method.Module.ResolveMember(token);
                                string target = Convert.ToString(member, CultureInfo.InvariantCulture);
                                if (target.IndexOf("RecoNet.Security", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    target.IndexOf("Encrypto", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                    target.IndexOf("Decrypto", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    Console.WriteLine(Path.GetFileName(path) + ": " + method.DeclaringType.FullName + "." + method.Name + " IL_" + offset.ToString("X4", CultureInfo.InvariantCulture) + " -> " + target);
                                }
                            }
                            catch
                            {
                            }
                        }
                        else
                        {
                            SkipOperand(il, ref pos, op);
                        }
                    }
                }
            }
        }

        private static OpCode ResolveOpCode(ushort value)
        {
            foreach (FieldInfo field in typeof(OpCodes).GetFields(BindingFlags.Public | BindingFlags.Static))
            {
                if (field.FieldType == typeof(OpCode))
                {
                    OpCode op = (OpCode)field.GetValue(null);
                    if ((ushort)op.Value == value)
                    {
                        return op;
                    }
                }
            }

            return OpCodes.Nop;
        }

        private static void SkipOperand(byte[] il, ref int pos, OpCode op)
        {
            switch (op.OperandType)
            {
                case OperandType.InlineNone:
                    return;
                case OperandType.ShortInlineI:
                case OperandType.ShortInlineVar:
                case OperandType.ShortInlineBrTarget:
                    pos += 1;
                    return;
                case OperandType.InlineVar:
                    pos += 2;
                    return;
                case OperandType.InlineI:
                case OperandType.ShortInlineR:
                case OperandType.InlineBrTarget:
                case OperandType.InlineString:
                case OperandType.InlineField:
                case OperandType.InlineType:
                case OperandType.InlineSig:
                case OperandType.InlineTok:
                    pos += 4;
                    return;
                case OperandType.InlineI8:
                case OperandType.InlineR:
                    pos += 8;
                    return;
                case OperandType.InlineSwitch:
                    int count = BitConverter.ToInt32(il, pos);
                    pos += 4 + count * 4;
                    return;
            }
        }

        private static string ReadOperand(Module module, byte[] il, ref int pos, OpCode op, int offset)
        {
            try
            {
                switch (op.OperandType)
                {
                    case OperandType.InlineNone:
                        return "";
                    case OperandType.ShortInlineI:
                        return il[pos++].ToString(CultureInfo.InvariantCulture);
                    case OperandType.InlineI:
                        int i = BitConverter.ToInt32(il, pos);
                        pos += 4;
                        return i.ToString(CultureInfo.InvariantCulture);
                    case OperandType.InlineI8:
                        long l = BitConverter.ToInt64(il, pos);
                        pos += 8;
                        return l.ToString(CultureInfo.InvariantCulture);
                    case OperandType.ShortInlineR:
                        float f = BitConverter.ToSingle(il, pos);
                        pos += 4;
                        return f.ToString(CultureInfo.InvariantCulture);
                    case OperandType.InlineR:
                        double d = BitConverter.ToDouble(il, pos);
                        pos += 8;
                        return d.ToString(CultureInfo.InvariantCulture);
                    case OperandType.ShortInlineVar:
                        return "V_" + il[pos++].ToString(CultureInfo.InvariantCulture);
                    case OperandType.InlineVar:
                        short s = BitConverter.ToInt16(il, pos);
                        pos += 2;
                        return "V_" + s.ToString(CultureInfo.InvariantCulture);
                    case OperandType.ShortInlineBrTarget:
                        sbyte delta = unchecked((sbyte)il[pos++]);
                        return "IL_" + (pos + delta).ToString("X4", CultureInfo.InvariantCulture);
                    case OperandType.InlineBrTarget:
                        int deltaLong = BitConverter.ToInt32(il, pos);
                        pos += 4;
                        return "IL_" + (pos + deltaLong).ToString("X4", CultureInfo.InvariantCulture);
                    case OperandType.InlineSwitch:
                        int count = BitConverter.ToInt32(il, pos);
                        pos += 4;
                        int basePos = pos + count * 4;
                        List<string> targets = new List<string>();
                        for (int n = 0; n < count; n++)
                        {
                            int switchDelta = BitConverter.ToInt32(il, pos);
                            pos += 4;
                            targets.Add("IL_" + (basePos + switchDelta).ToString("X4", CultureInfo.InvariantCulture));
                        }
                        return String.Join(",", targets.ToArray());
                    case OperandType.InlineString:
                        int stringToken = BitConverter.ToInt32(il, pos);
                        pos += 4;
                        return "\"" + module.ResolveString(stringToken) + "\"";
                    case OperandType.InlineMethod:
                        int methodToken = BitConverter.ToInt32(il, pos);
                        pos += 4;
                        return Convert.ToString(module.ResolveMethod(methodToken), CultureInfo.InvariantCulture);
                    case OperandType.InlineField:
                        int fieldToken = BitConverter.ToInt32(il, pos);
                        pos += 4;
                        return Convert.ToString(module.ResolveField(fieldToken), CultureInfo.InvariantCulture);
                    case OperandType.InlineType:
                    case OperandType.InlineTok:
                        int memberToken = BitConverter.ToInt32(il, pos);
                        pos += 4;
                        return Convert.ToString(module.ResolveMember(memberToken), CultureInfo.InvariantCulture);
                    case OperandType.InlineSig:
                        int sigToken = BitConverter.ToInt32(il, pos);
                        pos += 4;
                        return "sig:" + sigToken.ToString(CultureInfo.InvariantCulture);
                    default:
                        return "";
                }
            }
            catch (Exception ex)
            {
                return "operand-error@" + offset.ToString("X4", CultureInfo.InvariantCulture) + ":" + ex.Message;
            }
        }

        public Crypto(string exePath)
        {
            Assembly asm = Assembly.LoadFrom(exePath);
            Type security = asm.GetType("RecoNet.Security", true);
            instance = Activator.CreateInstance(security, true);
            decrypt = security.GetMethod("Decrypto", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, new[] { typeof(string) }, null);
            encrypt = security.GetMethod("Encrypto", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, new[] { typeof(string) }, null);
            decryptOld = security.GetMethod("DecryptoOld", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, new[] { typeof(string) }, null);
            encryptOld = security.GetMethod("EncryptoOld", BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Instance, null, new[] { typeof(string) }, null);
            if (decrypt == null || encrypt == null || decryptOld == null || encryptOld == null)
            {
                throw new MissingMethodException("RecoNet.Security Encrypto/Decrypto or EncryptoOld/DecryptoOld not found.");
            }
        }

        public bool CanRoundTripOld()
        {
            string plain = "0000000030.388@0089990021.3";
            return String.Equals(DecryptOld(EncryptOld(plain)), plain, StringComparison.Ordinal);
        }

        public bool CanRoundTrip()
        {
            string plain = "0000000030.388@0089990021.3";
            return String.Equals(Decrypt(Encrypt(plain)), plain, StringComparison.Ordinal);
        }

        public string Decrypt(string text)
        {
            if (String.IsNullOrWhiteSpace(text))
            {
                return "";
            }

            return Convert.ToString(decrypt.Invoke(instance, new object[] { text }), CultureInfo.InvariantCulture);
        }

        public string Encrypt(string text)
        {
            if (String.IsNullOrWhiteSpace(text))
            {
                return "";
            }

            return Convert.ToString(encrypt.Invoke(instance, new object[] { text }), CultureInfo.InvariantCulture);
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
        public List<ResourceMapping> ResourceInsertMappings = new List<ResourceMapping>();
        public List<ResourceMapping> AliasResourceMappings = new List<ResourceMapping>();
        public List<ResourceMapping> UnitDifferences = new List<ResourceMapping>();
        public List<ResourceMapping> NameReviewMappings = new List<ResourceMapping>();
        public List<ResourceMapping> ManualOverrideMappings = new List<ResourceMapping>();
        public List<ResourceMapping> UnknownCodes = new List<ResourceMapping>();
        public List<string> Conflicts = new List<string>();
        public HashSet<string> TargetBooks = new HashSet<string>(StringComparer.Ordinal);
        public HashSet<string> TargetQuotaKeys = new HashSet<string>(StringComparer.Ordinal);
        public HashSet<string> TargetSectionKeys = new HashSet<string>(StringComparer.Ordinal);
        public Dictionary<int, Resource> SourceMaterials = new Dictionary<int, Resource>();
        public Dictionary<int, Resource> SourceMachines = new Dictionary<int, Resource>();
        public Dictionary<string, List<Resource>> TargetMaterials = new Dictionary<string, List<Resource>>();
        public Dictionary<string, List<Resource>> TargetMachines = new Dictionary<string, List<Resource>>();
        public Dictionary<int, Resource> TargetMaterialByCode = new Dictionary<int, Resource>();
        public Dictionary<int, Resource> TargetMachineByCode = new Dictionary<int, Resource>();
        public long MaterialRangeStart;
        public long MaterialRangeEnd;
        public long MachineRangeStart;
        public long MachineRangeEnd;
        public long NextMaterialCode;
        public long NextMachineCode;
        public long NextMaterialNineCode = 899000001;
        public long NextMachineNineCode = 919000001;

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
        public decimal QuantityValue;
    }

    internal sealed class ConsumeRow
    {
        public string Book;
        public string QuotaCode;
        public int Code;
        public string Quantity;
    }

    internal sealed class ManualResourceDecision
    {
        public int OldCode;
        public string SupplementCode;
        public string CandidateCode;
        public string ManualCode;
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
        public string ManualOverride;

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
