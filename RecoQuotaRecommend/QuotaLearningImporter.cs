using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;

namespace RecoQuotaRecommend
{
    internal static class QuotaLearningImporter
    {
        private sealed class BudgetQuota
        {
            public string FilePath;
            public string SheetName;
            public int RowNumber;
            public string GroupName;
            public string RawCode;
            public string NormalizedCode;
            public bool IsMaterial;
            public string Name;
            public string Unit;
            public string Quantity;
        }

        private sealed class QuantityItem
        {
            public string FilePath;
            public string SheetName;
            public int RowNumber;
            public string SectionName;
            public string Name;
            public string Unit;
            public string Expression;
        }

        private sealed class MatchResult
        {
            public QuantityItem Item;
            public int Score;
            public string Reason;
        }

        private static int Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            if (args.Length < 1 || args.Length > 2)
            {
                Console.Error.WriteLine("Usage: QuotaLearningImporter <project-folder> [output-folder]");
                return 2;
            }

            string projectFolder = Path.GetFullPath(args[0]);
            if (!Directory.Exists(projectFolder))
            {
                Console.Error.WriteLine("Project folder not found: " + projectFolder);
                return 3;
            }

            string outputFolder = args.Length >= 2
                ? Path.GetFullPath(args[1])
                : Path.Combine(Path.GetDirectoryName(projectFolder), "RecoQuotaData");
            Directory.CreateDirectory(outputFolder);

            List<string> excelFiles = Directory.GetFiles(projectFolder, "*.xls*", SearchOption.AllDirectories)
                .Where(p => !Path.GetFileName(p).StartsWith("~$", StringComparison.Ordinal))
                .ToList();

            List<BudgetQuota> quotas = new List<BudgetQuota>();
            List<QuantityItem> quantities = new List<QuantityItem>();
            foreach (string file in excelFiles)
            {
                if (LooksLikeBudgetFile(file))
                {
                    quotas.AddRange(ReadBudgetQuotas(file));
                }
                else
                {
                    quantities.AddRange(ReadQuantityItems(file));
                }
            }

            string jsonlPath = Path.Combine(outputFolder, "learning.jsonl");
            string csvPath = Path.Combine(outputFolder, "learning.csv");
            string summaryPath = Path.Combine(outputFolder, "learning-summary.txt");
            int records = WriteLearningFiles(projectFolder, quotas, quantities, jsonlPath, csvPath);

            File.WriteAllText(
                summaryPath,
                "Project: " + projectFolder + Environment.NewLine
                + "ImportedAt: " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture) + Environment.NewLine
                + "ExcelFiles: " + excelFiles.Count.ToString(CultureInfo.InvariantCulture) + Environment.NewLine
                + "BudgetQuotas: " + quotas.Count.ToString(CultureInfo.InvariantCulture) + Environment.NewLine
                + "QuantityItems: " + quantities.Count.ToString(CultureInfo.InvariantCulture) + Environment.NewLine
                + "LearningRecords: " + records.ToString(CultureInfo.InvariantCulture) + Environment.NewLine
                + "Jsonl: " + jsonlPath + Environment.NewLine
                + "Csv: " + csvPath + Environment.NewLine,
                Encoding.UTF8);

            Console.WriteLine("Imported budget quotas: " + quotas.Count.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine("Imported quantity items: " + quantities.Count.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine("Learning records: " + records.ToString(CultureInfo.InvariantCulture));
            Console.WriteLine("Output: " + outputFolder);
            return 0;
        }

        private static bool LooksLikeBudgetFile(string path)
        {
            string name = Path.GetFileName(path);
            return name.IndexOf("估算", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("概算", StringComparison.OrdinalIgnoreCase) >= 0
                || name.IndexOf("预算", StringComparison.OrdinalIgnoreCase) >= 0;
        }

        private static List<BudgetQuota> ReadBudgetQuotas(string path)
        {
            List<BudgetQuota> result = new List<BudgetQuota>();
            using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                IWorkbook workbook = WorkbookFactory.Create(stream);
                for (int s = 0; s < workbook.NumberOfSheets; s++)
                {
                    ISheet sheet = workbook.GetSheetAt(s);
                    int headerRow = FindRowContaining(sheet, "单价编号");
                    if (headerRow < 0)
                    {
                        continue;
                    }

                    string currentGroup = "";
                    for (int r = headerRow + 1; r <= sheet.LastRowNum; r++)
                    {
                        IRow row = sheet.GetRow(r);
                        if (row == null)
                        {
                            continue;
                        }

                        string code = CellText(row, 0);
                        string name = CellText(row, 1);
                        if (String.IsNullOrWhiteSpace(code) && String.IsNullOrWhiteSpace(name))
                        {
                            continue;
                        }

                        if (code == "-" && !String.IsNullOrWhiteSpace(name))
                        {
                            currentGroup = CleanText(name);
                            continue;
                        }

                        string normalized = NormalizeQuotaCode(code);
                        if (String.IsNullOrEmpty(normalized))
                        {
                            continue;
                        }

                        result.Add(new BudgetQuota
                        {
                            FilePath = path,
                            SheetName = sheet.SheetName,
                            RowNumber = r + 1,
                            GroupName = currentGroup,
                            RawCode = CleanText(code),
                            NormalizedCode = normalized,
                            IsMaterial = IsMaterialCode(normalized),
                            Name = CleanText(name),
                            Unit = CleanText(CellText(row, 2)),
                            Quantity = CleanText(CellText(row, 3))
                        });
                    }
                }
            }

            return result;
        }

        private static List<QuantityItem> ReadQuantityItems(string path)
        {
            List<QuantityItem> result = new List<QuantityItem>();
            using (FileStream stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                IWorkbook workbook = WorkbookFactory.Create(stream);
                for (int s = 0; s < workbook.NumberOfSheets; s++)
                {
                    ISheet sheet = workbook.GetSheetAt(s);
                    string currentSection = "";
                    for (int r = sheet.FirstRowNum; r <= sheet.LastRowNum; r++)
                    {
                        IRow row = sheet.GetRow(r);
                        if (row == null)
                        {
                            continue;
                        }

                        string order = CleanText(CellText(row, 0));
                        string name = CleanText(CellText(row, 1));
                        string unit = CleanText(CellText(row, 5));
                        string expression = CleanText(CellText(row, 6));

                        if (IsChineseSection(order) && !String.IsNullOrWhiteSpace(name))
                        {
                            currentSection = name;
                            continue;
                        }

                        if (!IsNumericOrder(order) || String.IsNullOrWhiteSpace(name))
                        {
                            continue;
                        }

                        result.Add(new QuantityItem
                        {
                            FilePath = path,
                            SheetName = sheet.SheetName,
                            RowNumber = r + 1,
                            SectionName = currentSection,
                            Name = name,
                            Unit = unit,
                            Expression = expression
                        });
                    }
                }
            }

            return result;
        }

        private static int WriteLearningFiles(string projectFolder, List<BudgetQuota> quotas, List<QuantityItem> quantities, string jsonlPath, string csvPath)
        {
            string importedAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            List<string> jsonLines = new List<string>();
            List<string> csvLines = new List<string>();
            csvLines.Add("project_name,budget_file,budget_sheet,budget_row,budget_group,quota_code,quota_name,quota_unit,quantity_file,quantity_sheet,quantity_row,quantity_section,quantity_name,quantity_unit,quantity_expression,match_score,match_reason,imported_at");

            foreach (BudgetQuota quota in quotas)
            {
                MatchResult match = FindBestQuantityMatch(quota, quantities);
                QuantityItem item = match.Item;
                Dictionary<string, string> record = new Dictionary<string, string>();
                record["project_name"] = Path.GetFileName(projectFolder);
                record["budget_file"] = Path.GetFileName(quota.FilePath);
                record["budget_sheet"] = quota.SheetName;
                record["budget_row"] = quota.RowNumber.ToString(CultureInfo.InvariantCulture);
                record["budget_group"] = quota.GroupName;
                record["quota_code"] = quota.NormalizedCode;
                record["quota_name"] = quota.Name;
                record["quota_unit"] = quota.Unit;
                record["quantity_file"] = item == null ? "" : Path.GetFileName(item.FilePath);
                record["quantity_sheet"] = item == null ? "" : item.SheetName;
                record["quantity_row"] = item == null ? "" : item.RowNumber.ToString(CultureInfo.InvariantCulture);
                record["quantity_section"] = item == null ? "" : item.SectionName;
                record["quantity_name"] = item == null ? "" : item.Name;
                record["quantity_unit"] = item == null ? "" : item.Unit;
                record["quantity_expression"] = item == null ? "" : item.Expression;
                record["match_score"] = match.Score.ToString(CultureInfo.InvariantCulture);
                record["match_reason"] = match.Reason;
                record["imported_at"] = importedAt;

                jsonLines.Add(ToJson(record));
                csvLines.Add(ToCsv(record.Values.ToList()));
            }

            File.WriteAllLines(jsonlPath, jsonLines.ToArray(), Encoding.UTF8);
            File.WriteAllLines(csvPath, csvLines.ToArray(), Encoding.UTF8);
            return jsonLines.Count;
        }

        private static MatchResult FindBestQuantityMatch(BudgetQuota quota, List<QuantityItem> quantities)
        {
            MatchResult best = new MatchResult { Score = 0, Reason = "" };
            foreach (QuantityItem item in quantities)
            {
                int score = 0;
                List<string> reasons = new List<string>();
                if (!String.IsNullOrWhiteSpace(quota.GroupName) && TextSimilar(quota.GroupName, item.SectionName))
                {
                    score += 40;
                    reasons.Add("分组匹配");
                }

                if (UnitCompatible(quota.Unit, item.Unit))
                {
                    score += 10;
                    reasons.Add("单位相近");
                }

                string rule = "";
                int ruleScore = RuleScore(quota.Name, item.Name, out rule);
                if (ruleScore > 0)
                {
                    score += ruleScore;
                    reasons.Add(rule);
                }

                int overlap = KeywordOverlap(quota.Name, item.Name);
                if (overlap > 0)
                {
                    score += overlap * 4;
                    reasons.Add("关键词重合" + overlap.ToString(CultureInfo.InvariantCulture));
                }

                if (score > best.Score)
                {
                    best.Item = item;
                    best.Score = score;
                    best.Reason = String.Join(";", reasons.ToArray());
                }
            }

            return best;
        }

        private static int RuleScore(string quotaName, string quantityName, out string reason)
        {
            string q = NormalizeText(quotaName);
            string n = NormalizeText(quantityName);
            reason = "";

            if ((q.Contains("挖掘机") || q.Contains("挖土") || q.Contains("装车")) && (n.Contains("挖方") || n.Contains("挖土")))
            {
                reason = "规则:挖方";
                return 35;
            }

            if (q.Contains("回填") && n.Contains("回填"))
            {
                reason = "规则:回填";
                return 35;
            }

            if ((q.Contains("运土") || q.Contains("自卸汽车")) && (n.Contains("外运") || n.Contains("运土") || n.Contains("土方")))
            {
                reason = "规则:运土";
                return 35;
            }

            if (q.Contains("钢筋") && n.Contains("钢筋"))
            {
                reason = "规则:钢筋";
                return 35;
            }

            if (q.Contains("c30") && n.Contains("c30"))
            {
                reason = "规则:C30";
                return q.Contains("混凝土") && n.Contains("混凝土") ? 35 : 25;
            }

            if (q.Contains("c20") && n.Contains("c20"))
            {
                reason = "规则:C20";
                return q.Contains("垫层") || n.Contains("垫层") ? 40 : 30;
            }

            if (q.Contains("垫层") && n.Contains("垫层"))
            {
                reason = "规则:垫层";
                return 30;
            }

            if (q.Contains("模板") && n.Contains("模板"))
            {
                reason = "规则:模板";
                return 25;
            }

            return 0;
        }

        private static int KeywordOverlap(string left, string right)
        {
            HashSet<string> leftTokens = new HashSet<string>(Tokenize(left));
            int count = 0;
            foreach (string token in Tokenize(right))
            {
                if (leftTokens.Contains(token))
                {
                    count++;
                }
            }

            return count;
        }

        private static IEnumerable<string> Tokenize(string text)
        {
            string normalized = NormalizeText(text);
            string[] separators = new string[] { " ", "/", "，", ",", "、", "(", ")", "（", "）", "[", "]", "【", "】" };
            foreach (string token in normalized.Split(separators, StringSplitOptions.RemoveEmptyEntries))
            {
                if (token.Length >= 2 && !IsNumberLike(token))
                {
                    yield return token;
                }
            }
        }

        private static string NormalizeQuotaCode(string raw)
        {
            string code = CleanText(raw);
            if (String.IsNullOrWhiteSpace(code) || code == "-")
            {
                return "";
            }

            code = code.Replace("参", "");
            int star = code.IndexOf('*');
            if (star >= 0)
            {
                code = code.Substring(0, star);
            }

            return code.Trim();
        }

        private static bool IsMaterialCode(string code)
        {
            return code.All(Char.IsDigit);
        }

        private static bool UnitCompatible(string quotaUnit, string quantityUnit)
        {
            string q = NormalizeUnit(quotaUnit);
            string n = NormalizeUnit(quantityUnit);
            if (String.IsNullOrEmpty(q) || String.IsNullOrEmpty(n))
            {
                return false;
            }

            return q == n || q.EndsWith(n, StringComparison.Ordinal) || n.EndsWith(q, StringComparison.Ordinal);
        }

        private static string NormalizeUnit(string unit)
        {
            return CleanText(unit)
                .Replace("100", "")
                .Replace("10", "")
                .Replace("立方米", "m3")
                .Replace("平方米", "m2")
                .Replace("米", "m")
                .Replace("吨", "t")
                .Replace("㎡", "m2")
                .Replace("m²", "m2")
                .Replace("m³", "m3")
                .ToLowerInvariant();
        }

        private static bool TextSimilar(string left, string right)
        {
            string l = NormalizeText(left);
            string r = NormalizeText(right);
            return !String.IsNullOrEmpty(l) && !String.IsNullOrEmpty(r) && (l.Contains(r) || r.Contains(l));
        }

        private static string NormalizeText(string text)
        {
            return CleanText(text)
                .Replace("≤", "")
                .Replace("≥", "")
                .Replace("Φ", "φ")
                .Replace("　", "")
                .Replace(" ", "")
                .ToLowerInvariant();
        }

        private static string CleanText(string text)
        {
            return (text ?? "").Replace("\r", "").Replace("\n", "").Trim();
        }

        private static bool IsChineseSection(string text)
        {
            return new HashSet<string> { "一", "二", "三", "四", "五", "六", "七", "八", "九", "十" }.Contains(CleanText(text));
        }

        private static bool IsNumericOrder(string text)
        {
            int ignored;
            return Int32.TryParse(CleanText(text), NumberStyles.Integer, CultureInfo.InvariantCulture, out ignored);
        }

        private static bool IsNumberLike(string text)
        {
            decimal ignored;
            return Decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out ignored);
        }

        private static int FindRowContaining(ISheet sheet, string text)
        {
            for (int r = sheet.FirstRowNum; r <= sheet.LastRowNum; r++)
            {
                IRow row = sheet.GetRow(r);
                if (row == null)
                {
                    continue;
                }

                int lastCell = row.LastCellNum < 0 ? 0 : (int)row.LastCellNum;
                for (int c = 0; c < lastCell; c++)
                {
                    if (CellText(row, c).IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        return r;
                    }
                }
            }

            return -1;
        }

        private static string CellText(IRow row, int index)
        {
            ICell cell = row.GetCell(index);
            if (cell == null)
            {
                return "";
            }

            try
            {
                if (cell.CellType == CellType.Numeric)
                {
                    return cell.NumericCellValue.ToString("G15", CultureInfo.InvariantCulture);
                }

                if (cell.CellType == CellType.Formula)
                {
                    if (cell.CachedFormulaResultType == CellType.Numeric)
                    {
                        return cell.NumericCellValue.ToString("G15", CultureInfo.InvariantCulture);
                    }
                }

                return cell.ToString().Trim();
            }
            catch
            {
                return cell.ToString().Trim();
            }
        }

        private static string ToJson(Dictionary<string, string> record)
        {
            return "{" + String.Join(",", record.Select(kv => "\"" + JsonEscape(kv.Key) + "\":\"" + JsonEscape(kv.Value) + "\"").ToArray()) + "}";
        }

        private static string JsonEscape(string value)
        {
            return (value ?? "").Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\t", "\\t");
        }

        private static string ToCsv(List<string> values)
        {
            return String.Join(",", values.Select(CsvEscape).ToArray());
        }

        private static string CsvEscape(string value)
        {
            string text = value ?? "";
            if (text.IndexOfAny(new char[] { ',', '"', '\r', '\n' }) >= 0)
            {
                return "\"" + text.Replace("\"", "\"\"") + "\"";
            }

            return text;
        }
    }
}
