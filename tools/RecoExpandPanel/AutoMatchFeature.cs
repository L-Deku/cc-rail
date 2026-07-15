using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Data.SqlClient;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        private const int AutoMatchCellLimit = 20000;
        private const int AutoMatchReadBlockRows = 2000;
        private const int AutoMatchFallbackCellLimit = 800;
        private const int AutoMatchCombinationLimit = 500;
        private const int AutoMatchCandidatePerTermLimit = 60;

        private sealed class QuantityExpressionTerm
        {
            public bool Negative;
            public decimal Literal;
            public int LiteralDecimals;
            public string Suffix;
        }

        private sealed class AutoMatchCellValue
        {
            public AiExcelCell Cell;
            public decimal Value;
        }

        private sealed class AutoMatchExpressionCandidate
        {
            public string Expression;
            public string CellAddress;
            public string DisplayValue;
            public string ExcelQuantityText;
            public decimal Quantity;
            public string QuantityName;
            public string Status;
            public bool Checked;
            public int FirstRow;
            public int FirstColumn;
            public decimal Score;
            public readonly List<string> Addresses = new List<string>();
            public List<AutoMatchCandidateOption> Options = new List<AutoMatchCandidateOption>();
        }

        private sealed class AutoMatchCandidateOption
        {
            public string Expression;
            public string CellAddress;
            public string DisplayValue;
            public string ExcelQuantityText;
            public string QuantityName;
            public decimal Quantity;
            public int FirstRow;
            public int FirstColumn;
            public string Label;
        }

        private sealed class AutoMatchNumberIndex
        {
            private readonly List<AutoMatchCellValue> cells;
            private readonly Dictionary<int, Dictionary<decimal, List<AutoMatchCellValue>>> maps =
                new Dictionary<int, Dictionary<decimal, List<AutoMatchCellValue>>>();

            public AutoMatchNumberIndex(List<AutoMatchCellValue> sourceCells)
            {
                cells = sourceCells ?? new List<AutoMatchCellValue>();
            }

            public List<AutoMatchCellValue> GetCandidates(QuantityExpressionTerm term)
            {
                int decimals = Math.Max(0, Math.Min(8, term.LiteralDecimals));
                Dictionary<decimal, List<AutoMatchCellValue>> map;
                if (!maps.TryGetValue(decimals, out map))
                {
                    map = new Dictionary<decimal, List<AutoMatchCellValue>>();
                    foreach (AutoMatchCellValue cell in cells)
                    {
                        decimal key = RoundAutoMatchValue(cell.Value, decimals);
                        List<AutoMatchCellValue> bucket;
                        if (!map.TryGetValue(key, out bucket))
                        {
                            bucket = new List<AutoMatchCellValue>();
                            map[key] = bucket;
                        }

                        bucket.Add(cell);
                    }

                    maps[decimals] = map;
                }

                List<AutoMatchCellValue> result;
                decimal target = RoundAutoMatchValue(term.Literal, decimals);
                if (!map.TryGetValue(target, out result))
                {
                    return new List<AutoMatchCellValue>();
                }

                return result
                    .Where(cell => RelativeDifference(term.Literal, cell.Value) <= 0.03m)
                    .OrderBy(cell => RelativeDifference(term.Literal, cell.Value))
                    .ThenBy(cell => cell.Cell.Row)
                    .ThenBy(cell => cell.Cell.Column)
                    .Take(AutoMatchCandidatePerTermLimit)
                    .ToList();
            }
        }

        private static decimal RoundAutoMatchValue(decimal value, int decimals)
        {
            return Decimal.Round(value, Math.Max(0, Math.Min(8, decimals)), MidpointRounding.AwayFromZero);
        }

        private static bool TryParseQuantityExpressionTerms(string expression, out List<QuantityExpressionTerm> terms)
        {
            terms = new List<QuantityExpressionTerm>();
            string normalized = NormalizeAutoMatchExpression(expression);
            if (String.IsNullOrWhiteSpace(normalized) ||
                normalized.IndexOf('(') >= 0 ||
                normalized.IndexOf(')') >= 0 ||
                Regex.IsMatch(normalized, "[A-Z]"))
            {
                return false;
            }

            int start = 0;
            bool negative = false;
            if (normalized[0] == '+' || normalized[0] == '-')
            {
                negative = normalized[0] == '-';
                start = 1;
            }

            for (int i = start; i <= normalized.Length; i++)
            {
                if (i == normalized.Length || normalized[i] == '+' || normalized[i] == '-')
                {
                    string token = normalized.Substring(start, i - start);
                    QuantityExpressionTerm term;
                    if (!TryBuildQuantityExpressionTerm(negative, token, out term))
                    {
                        terms.Clear();
                        return false;
                    }

                    terms.Add(term);
                    if (i < normalized.Length)
                    {
                        negative = normalized[i] == '-';
                        start = i + 1;
                    }
                }
            }

            return terms.Count > 0;
        }

        // 支持“(内部和式)*系数”/“(内部和式)/除数”形式（如 (49521.626+494012.33+13500)*1.04）：
        // 内部按普通项解析参与匹配，外层系数原样附加到生成的表达式上。outerScale 为空表示普通表达式。
        private static bool TryParseQuantityExpressionTermsWithScale(string expression, out List<QuantityExpressionTerm> terms, out string outerScale)
        {
            outerScale = "";
            if (TryParseQuantityExpressionTerms(expression, out terms))
            {
                return true;
            }

            string normalized = NormalizeAutoMatchExpression(expression);
            if (String.IsNullOrEmpty(normalized) || normalized[0] != '(')
            {
                return false;
            }

            int depth = 0;
            int close = -1;
            for (int i = 0; i < normalized.Length; i++)
            {
                if (normalized[i] == '(')
                {
                    depth++;
                }
                else if (normalized[i] == ')')
                {
                    depth--;
                    if (depth == 0)
                    {
                        close = i;
                        break;
                    }
                }
            }

            if (close <= 1)
            {
                return false;
            }

            string inner = normalized.Substring(1, close - 1);
            string tail = normalized.Substring(close + 1);
            if (tail.Length == 0 || !Regex.IsMatch(tail, @"^([*/]\d+(?:\.\d+)?)+$"))
            {
                return false;
            }

            if (!TryParseQuantityExpressionTerms(inner, out terms))
            {
                return false;
            }

            outerScale = tail;
            return true;
        }

        private static bool TryBuildQuantityExpressionTerm(bool negative, string token, out QuantityExpressionTerm term)
        {
            term = null;
            token = (token ?? "").Trim();
            Match match = Regex.Match(token, @"^(\d+(?:\.\d+)?)([*/]\d+(?:\.\d+)?)*$");
            if (!match.Success)
            {
                return false;
            }

            string literalText = match.Groups[1].Value;
            decimal literal;
            if (!Decimal.TryParse(literalText, NumberStyles.Float, CultureInfo.InvariantCulture, out literal))
            {
                return false;
            }

            int dot = literalText.IndexOf('.');
            term = new QuantityExpressionTerm();
            term.Negative = negative;
            term.Literal = literal;
            term.LiteralDecimals = dot >= 0 ? literalText.Length - dot - 1 : 0;
            term.Suffix = token.Substring(literalText.Length);
            return true;
        }

        private static string NormalizeAutoMatchExpression(string expression)
        {
            string text = NormalizeExpressionOperators(expression)
                .Replace(" ", "")
                .Replace("\u3000", "")
                .Replace("\u00f7", "/")
                .Replace("\uff0b", "+")
                .Replace("\uff0d", "-")
                .Replace("\uff0a", "*")
                .Replace("\uff0f", "/")
                .Replace("\uff08", "(")
                .Replace("\uff09", ")");

            StringBuilder builder = new StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (ch >= '\uff10' && ch <= '\uff19')
                {
                    builder.Append((char)('0' + (ch - '\uff10')));
                }
                else if (ch == '\uff0e')
                {
                    builder.Append('.');
                }
                else
                {
                    builder.Append(ch);
                }
            }

            return builder.ToString().ToUpperInvariant();
        }

        private static int ComputeAutoMatchTextSimilarity(string left, string right)
        {
            string a = NormalizeAutoMatchText(left);
            string b = NormalizeAutoMatchText(right);
            if (String.IsNullOrEmpty(a) || String.IsNullOrEmpty(b))
            {
                return 0;
            }

            if (String.Equals(a, b, StringComparison.OrdinalIgnoreCase))
            {
                return 100;
            }

            if (a.IndexOf(b, StringComparison.OrdinalIgnoreCase) >= 0 ||
                b.IndexOf(a, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return 90;
            }

            double dice = ComputeBigramDice(a, b);
            double charOverlap = ComputeCharacterOverlap(a, b);
            int score = Convert.ToInt32(Math.Round(dice * 80.0 + charOverlap * 20.0));
            return Math.Max(0, Math.Min(100, score));
        }

        private static string NormalizeAutoMatchText(string text)
        {
            if (String.IsNullOrWhiteSpace(text))
            {
                return "";
            }

            StringBuilder builder = new StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (Char.IsLetterOrDigit(ch) || ch >= 0x4e00)
                {
                    builder.Append(Char.ToUpperInvariant(ch));
                }
            }

            return builder.ToString();
        }

        private static double ComputeBigramDice(string a, string b)
        {
            List<string> left = BuildBigrams(a);
            List<string> right = BuildBigrams(b);
            if (left.Count == 0 || right.Count == 0)
            {
                return 0.0;
            }

            Dictionary<string, int> counts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (string value in left)
            {
                int count;
                counts.TryGetValue(value, out count);
                counts[value] = count + 1;
            }

            int intersection = 0;
            foreach (string value in right)
            {
                int count;
                if (counts.TryGetValue(value, out count) && count > 0)
                {
                    intersection++;
                    counts[value] = count - 1;
                }
            }

            return (2.0 * intersection) / (left.Count + right.Count);
        }

        private static List<string> BuildBigrams(string text)
        {
            List<string> result = new List<string>();
            if (String.IsNullOrEmpty(text))
            {
                return result;
            }

            if (text.Length == 1)
            {
                result.Add(text);
                return result;
            }

            for (int i = 0; i < text.Length - 1; i++)
            {
                result.Add(text.Substring(i, 2));
            }

            return result;
        }

        private static double ComputeCharacterOverlap(string a, string b)
        {
            HashSet<char> left = new HashSet<char>(a.ToCharArray());
            HashSet<char> right = new HashSet<char>(b.ToCharArray());
            if (left.Count == 0 || right.Count == 0)
            {
                return 0.0;
            }

            int overlap = left.Count(ch => right.Contains(ch));
            return overlap / (double)Math.Max(left.Count, right.Count);
        }

        private static bool TryParseTargetColumns(string text, out List<int> columns, out string error)
        {
            columns = new List<int>();
            error = null;
            if (String.IsNullOrWhiteSpace(text))
            {
                return true;
            }

            string[] parts = text.Split(new char[] { ',', ';', '\u3001', '\uff0c', '\uff1b', ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string raw in parts)
            {
                string token = raw.Trim().ToUpperInvariant();
                if (String.IsNullOrEmpty(token) || token.Any(ch => ch < 'A' || ch > 'Z'))
                {
                    error = "\u76ee\u6807\u5217\u683c\u5f0f\u4e0d\u6b63\u786e\uff1a" + raw + "\u3002\u8bf7\u586b\u5199 E \u6216 E,F\u3002";
                    return false;
                }

                int column = ColumnNameToNumber(token);
                if (column <= 0)
                {
                    error = "\u76ee\u6807\u5217\u683c\u5f0f\u4e0d\u6b63\u786e\uff1a" + raw + "\u3002\u8bf7\u586b\u5199 E \u6216 E,F\u3002";
                    return false;
                }

                if (!columns.Contains(column))
                {
                    columns.Add(column);
                }
            }

            columns.Sort();
            return true;
        }

        private static int ColumnNameToNumber(string name)
        {
            int value = 0;
            for (int i = 0; i < name.Length; i++)
            {
                char ch = name[i];
                if (ch < 'A' || ch > 'Z')
                {
                    return 0;
                }

                value = value * 26 + (ch - 'A' + 1);
            }

            return value;
        }

        private static List<AiMatchPreviewItem> BuildAutoMatchPreviewItems(List<AiQuotaMatchRow> quotas, AiExcelSelectionContext context, List<int> targetColumns, HashSet<long> alreadyBoundSequences)
        {
            List<AiMatchPreviewItem> preview = new List<AiMatchPreviewItem>();
            List<AutoMatchCellValue> numberCells = BuildAutoMatchNumberCells(context, targetColumns);
            AutoMatchNumberIndex numberIndex = new AutoMatchNumberIndex(numberCells);
            HashSet<string> usedAddresses = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            int previousRow = 0;
            int previousColumn = 0;

            foreach (AiQuotaMatchRow quota in quotas ?? new List<AiQuotaMatchRow>())
            {
                if (quota == null || quota.Link == null)
                {
                    continue;
                }

                if (!quota.Bindable)
                {
                    preview.Add(BuildAutoMatchPreviewItem(quota, context, null, "\u4e0d\u53c2\u4e0e\u7ed1\u5b9a"));
                    continue;
                }

                if (alreadyBoundSequences != null && alreadyBoundSequences.Contains(quota.Link.QuotaSequence))
                {
                    preview.Add(BuildAutoMatchPreviewItem(quota, context, null, "\u5df2\u7ed1\u5b9a"));
                    continue;
                }

                decimal quotaQuantity;
                string quotaError;
                if (!TryEvaluateDecimal(quota.CurrentQuantityText, out quotaQuantity, out quotaError))
                {
                    if (String.IsNullOrWhiteSpace(quota.CurrentQuantityText))
                    {
                        preview.Add(BuildAutoMatchPreviewItem(quota, context, null, "数量为0待手动绑定"));
                    }

                    continue;
                }

                // 数量为0的定额不参与自动匹配，由用户手动绑定到对应的（可为空值的）Excel单元格。
                if (quotaQuantity == 0m)
                {
                    preview.Add(BuildAutoMatchPreviewItem(quota, context, null, "数量为0待手动绑定"));
                    continue;
                }

                AutoMatchExpressionCandidate candidate;
                if (!TryBuildOperandAutoMatch(quota, quotaQuantity, context, numberIndex, usedAddresses, previousRow, previousColumn, out candidate))
                {
                    TryBuildWholeValueAutoMatch(quota, quotaQuantity, context, numberCells, usedAddresses, previousRow, previousColumn, out candidate);
                }

                if (candidate == null)
                {
                    continue;
                }

                preview.Add(BuildAutoMatchPreviewItem(quota, context, candidate, candidate.Status));
                if (candidate.Checked)
                {
                    foreach (string address in candidate.Addresses)
                    {
                        usedAddresses.Add(address);
                    }

                    previousRow = candidate.FirstRow;
                    previousColumn = candidate.FirstColumn;
                }
            }

            AddUnmatchedAiPreviewItems(preview, quotas, context);
            ApplyAutoMatchPreviewDefaults(preview, quotas, context);
            return SortAiPreviewItemsByQuotaOrder(preview, quotas);
        }

        private static AiMatchPreviewItem BuildAutoMatchPreviewItem(AiQuotaMatchRow quota, AiExcelSelectionContext context, AutoMatchExpressionCandidate candidate, string status)
        {
            AiMatchPreviewItem item = new AiMatchPreviewItem();
            item.Checked = candidate != null && candidate.Checked;
            item.Link = quota.Link;
            item.QuotaUnit = quota.QuotaUnit;
            item.WorkbookPath = context == null ? null : context.WorkbookPath;
            item.WorksheetName = context == null ? null : context.WorksheetName;
            item.Expression = candidate == null ? "" : candidate.Expression;
            item.CellAddress = candidate == null ? "" : candidate.CellAddress;
            item.DisplayValue = candidate == null ? "" : candidate.DisplayValue;
            item.ExcelQuantityText = candidate == null ? "" : candidate.ExcelQuantityText;
            item.QuantityName = candidate == null ? "" : candidate.QuantityName;
            item.CurrentQuantityText = quota.CurrentQuantityText;
            item.Bindable = quota.Bindable;
            item.ItemNo = quota.ItemNo ?? "";
            item.MatchStatus = status;
            item.MatchOptions = candidate == null ? new List<AutoMatchCandidateOption>() : (candidate.Options ?? new List<AutoMatchCandidateOption>());
            return item;
        }

        // 过滤模板铺量推送的区段：从"AI推送"标记行（编号为-）开始，到下一个正常标题行
        // （编号为-且名称不是AI推送）为止，区段内的行（含标记行）不进自动匹配预览；
        // 正常标题行本身保留作分组参照。条目边界重置状态。
        private static List<AiQuotaMatchRow> FilterAutoMatchPushedRows(List<AiQuotaMatchRow> rows)
        {
            List<AiQuotaMatchRow> result = new List<AiQuotaMatchRow>();
            bool skippingPushedBlock = false;
            string currentChapter = null;
            foreach (AiQuotaMatchRow row in rows ?? new List<AiQuotaMatchRow>())
            {
                if (row == null || row.Link == null)
                {
                    continue;
                }

                string chapter = row.Link.ChapterSeq ?? "";
                if (!String.Equals(chapter, currentChapter, StringComparison.OrdinalIgnoreCase))
                {
                    currentChapter = chapter;
                    skippingPushedBlock = false;
                }

                string code = (row.Link.QuotaCode ?? "").Trim();
                if (code == "-")
                {
                    string name = (row.Link.QuotaName ?? "").Trim();
                    if (name.StartsWith("AI推送", StringComparison.Ordinal))
                    {
                        skippingPushedBlock = true;
                        continue;
                    }

                    skippingPushedBlock = false;
                    result.Add(row);
                    continue;
                }

                if (skippingPushedBlock)
                {
                    continue;
                }

                result.Add(row);
            }

            return result;
        }

        private static void ApplyAutoMatchPreviewDefaults(List<AiMatchPreviewItem> preview, List<AiQuotaMatchRow> quotas, AiExcelSelectionContext context)
        {
            Dictionary<long, AiQuotaMatchRow> quotaBySequence = (quotas ?? new List<AiQuotaMatchRow>())
                .Where(q => q != null && q.Link != null)
                .GroupBy(q => q.Link.QuotaSequence)
                .ToDictionary(g => g.Key, g => g.First());

            foreach (AiMatchPreviewItem item in preview ?? new List<AiMatchPreviewItem>())
            {
                AiQuotaMatchRow quota;
                if (item == null || item.Link == null || !quotaBySequence.TryGetValue(item.Link.QuotaSequence, out quota))
                {
                    continue;
                }

                if (String.IsNullOrWhiteSpace(item.CurrentQuantityText))
                {
                    item.CurrentQuantityText = quota.CurrentQuantityText;
                }
                item.Bindable = quota.Bindable;
                if (String.IsNullOrWhiteSpace(item.ItemNo))
                {
                    item.ItemNo = quota.ItemNo ?? "";
                }

                if (String.IsNullOrWhiteSpace(item.MatchStatus))
                {
                    item.MatchStatus = String.IsNullOrWhiteSpace(item.Expression) ? "\u672a\u5339\u914d" : "\u5df2\u5339\u914d";
                }

                if (String.IsNullOrWhiteSpace(item.ExcelQuantityText) && !String.IsNullOrWhiteSpace(item.Expression) && context != null)
                {
                    item.ExcelQuantityText = BuildExcelQuantityTextForExpression(context, item.Expression);
                }

                if (String.IsNullOrWhiteSpace(item.WorkbookPath) && context != null)
                {
                    item.WorkbookPath = context.WorkbookPath;
                }

                if (String.IsNullOrWhiteSpace(item.WorksheetName) && context != null)
                {
                    item.WorksheetName = context.WorksheetName;
                }
            }
        }

        private static List<AutoMatchCellValue> BuildAutoMatchNumberCells(AiExcelSelectionContext context, List<int> targetColumns)
        {
            HashSet<int> targetSet = targetColumns != null && targetColumns.Count > 0
                ? new HashSet<int>(targetColumns)
                : null;
            Dictionary<string, AutoMatchCellValue> byAddress = new Dictionary<string, AutoMatchCellValue>(StringComparer.OrdinalIgnoreCase);
            foreach (AiExcelCell cell in context == null ? new List<AiExcelCell>() : context.Cells)
            {
                if (cell == null || !cell.IsNumber || !context.IsCellInTargetColumns(cell, targetSet))
                {
                    continue;
                }

                decimal value;
                string error;
                if (!TryEvaluateDecimal(cell.Text, out value, out error))
                {
                    continue;
                }

                string address = context.NormalizeMergedCellAddress(cell.Address);
                CellRef normalizedCellRef;
                AiExcelCell normalizedCell = cell;
                if (!String.Equals(address, cell.Address, StringComparison.OrdinalIgnoreCase) &&
                    TryParseCellAddress(address, out normalizedCellRef))
                {
                    normalizedCell = new AiExcelCell
                    {
                        Address = address,
                        Text = cell.Text,
                        Row = normalizedCellRef.Row,
                        Column = normalizedCellRef.Column,
                        IsNumber = cell.IsNumber
                    };
                }

                byAddress[address] = new AutoMatchCellValue { Cell = normalizedCell, Value = value };
            }

            return byAddress.Values.OrderBy(cell => cell.Cell.Row).ThenBy(cell => cell.Cell.Column).ToList();
        }

        private static bool TryBuildOperandAutoMatch(AiQuotaMatchRow quota, decimal quotaQuantity, AiExcelSelectionContext context, AutoMatchNumberIndex numberIndex, HashSet<string> usedAddresses, int previousRow, int previousColumn, out AutoMatchExpressionCandidate candidate)
        {
            candidate = null;
            List<AutoMatchExpressionCandidate> matches = new List<AutoMatchExpressionCandidate>();
            List<QuantityExpressionTerm> terms;
            string outerScale;
            if (!TryParseQuantityExpressionTermsWithScale(quota.CurrentQuantityText, out terms, out outerScale))
            {
                return false;
            }

            List<List<AutoMatchCellValue>> termCandidates = new List<List<AutoMatchCellValue>>();
            int combinationCount = 1;
            foreach (QuantityExpressionTerm term in terms)
            {
                List<AutoMatchCellValue> candidates = numberIndex.GetCandidates(term);
                if (candidates.Count == 0)
                {
                    return false;
                }

                termCandidates.Add(candidates);
                combinationCount = CapAutoMatchCount(combinationCount, candidates.Count);
            }

            List<List<AutoMatchCellValue>> combinations = BuildAutoMatchCombinations(termCandidates, AutoMatchCombinationLimit);
            foreach (List<AutoMatchCellValue> cells in combinations)
            {
                if (cells.Select(cell => cell.Cell.Address).Distinct(StringComparer.OrdinalIgnoreCase).Count() != cells.Count)
                {
                    continue;
                }

                string expression = BuildAutoMatchExpression(terms, cells);
                if (outerScale.Length > 0)
                {
                    expression = "(" + expression + ")" + outerScale;
                }

                string displayValue;
                decimal quantity;
                string error;
                if (!TryEvaluateAutoMatchExpression(context, expression, false, out displayValue, out quantity, out error))
                {
                    continue;
                }

                decimal diff = RelativeDifference(quotaQuantity, quantity);
                AutoMatchExpressionCandidate current = BuildAutoMatchCandidate(quota, context, terms, cells, expression, displayValue, quantity, diff, combinationCount, usedAddresses, previousRow, previousColumn);
                matches.Add(current);
                if (candidate == null || current.Score > candidate.Score)
                {
                    candidate = current;
                }
            }

            if (candidate != null)
            {
                candidate.Options = BuildAutoMatchCandidateOptions(matches);
            }

            return candidate != null;
        }

        private static bool TryBuildWholeValueAutoMatch(AiQuotaMatchRow quota, decimal quotaQuantity, AiExcelSelectionContext context, List<AutoMatchCellValue> numberCells, HashSet<string> usedAddresses, int previousRow, int previousColumn, out AutoMatchExpressionCandidate candidate)
        {
            candidate = null;
            List<AutoMatchExpressionCandidate> matches = new List<AutoMatchExpressionCandidate>();
            int matchCount = 0;
            // 定额数量文本的“数值*系数/除数”解析只与定额有关，提到循环外做一次。
            string scaleOp;
            decimal scaleFactor;
            decimal scaleLeftValue;
            bool hasScaleExpression = TryParseSimpleScaleExpression(quota.CurrentQuantityText, out scaleLeftValue, out scaleOp, out scaleFactor);
            foreach (AutoMatchCellValue cell in numberCells ?? new List<AutoMatchCellValue>())
            {
                string expression;
                if (RelativeDifference(quotaQuantity, cell.Value) <= 0.03m)
                {
                    expression = cell.Cell.Address;
                }
                else if (hasScaleExpression && RelativeDifference(scaleLeftValue, cell.Value) <= 0.03m)
                {
                    expression = cell.Cell.Address + scaleOp + scaleFactor.ToString(CultureInfo.InvariantCulture);
                }
                else
                {
                    continue;
                }

                string displayValue;
                decimal quantity;
                string error;
                if (!TryEvaluateAutoMatchExpression(context, expression, false, out displayValue, out quantity, out error))
                {
                    continue;
                }

                matchCount++;
                List<QuantityExpressionTerm> terms = new List<QuantityExpressionTerm>
                {
                    new QuantityExpressionTerm { Negative = false, Literal = cell.Value, LiteralDecimals = CountDecimalPlaces(cell.Cell.Text), Suffix = "" }
                };
                List<AutoMatchCellValue> cells = new List<AutoMatchCellValue> { cell };
                AutoMatchExpressionCandidate current = BuildAutoMatchCandidate(quota, context, terms, cells, expression, displayValue, quantity, RelativeDifference(quotaQuantity, quantity), matchCount, usedAddresses, previousRow, previousColumn);
                matches.Add(current);
                if (candidate == null || current.Score > candidate.Score)
                {
                    candidate = current;
                }
            }

            if (candidate != null && matchCount > 1 && candidate.Status == "\u5df2\u5339\u914d")
            {
                candidate.Checked = false;
                candidate.Status = "\u591a\u5904\u5339\u914d(" + matchCount.ToString(CultureInfo.InvariantCulture) + "\u5904)";
            }
            if (candidate != null)
            {
                candidate.Options = BuildAutoMatchCandidateOptions(matches);
            }

            return candidate != null;
        }

        private static int CountDecimalPlaces(string text)
        {
            text = (text ?? "").Trim();
            int dot = text.IndexOf('.');
            if (dot < 0)
            {
                return 0;
            }

            int count = 0;
            for (int i = dot + 1; i < text.Length && Char.IsDigit(text[i]); i++)
            {
                count++;
            }

            return count;
        }

        private static AutoMatchExpressionCandidate BuildAutoMatchCandidate(AiQuotaMatchRow quota, AiExcelSelectionContext context, List<QuantityExpressionTerm> terms, List<AutoMatchCellValue> cells, string expression, string displayValue, decimal quantity, decimal quantityDiff, int combinationCount, HashSet<string> usedAddresses, int previousRow, int previousColumn)
        {
            AutoMatchCellValue first = cells[0];
            string quantityName = BuildQuantityNameFromExcelRow(context, first.Cell.Address);
            AutoMatchExpressionCandidate candidate = new AutoMatchExpressionCandidate();
            candidate.Expression = expression;
            candidate.CellAddress = first.Cell.Address;
            candidate.DisplayValue = displayValue;
            candidate.ExcelQuantityText = BuildExcelQuantityTextForExpression(context, expression);
            candidate.Quantity = quantity;
            candidate.QuantityName = quantityName;
            candidate.FirstRow = first.Cell.Row;
            candidate.FirstColumn = first.Cell.Column;
            candidate.Checked = quantityDiff <= 0.03m && combinationCount <= 1;
            candidate.Status = quantityDiff > 0.03m
                ? "\u9a8c\u7b97\u4e0d\u7b26"
                : (combinationCount > 1 ? "\u591a\u5904\u5339\u914d(" + combinationCount.ToString(CultureInfo.InvariantCulture) + "\u5904)" : "\u5df2\u5339\u914d");
            foreach (AutoMatchCellValue cell in cells)
            {
                candidate.Addresses.Add(cell.Cell.Address);
            }

            candidate.Score = ComputeAutoMatchCandidateScore(quota, terms, cells, quantityName, quantityDiff, usedAddresses, previousRow, previousColumn);
            return candidate;
        }

        private static List<AutoMatchCandidateOption> BuildAutoMatchCandidateOptions(List<AutoMatchExpressionCandidate> candidates)
        {
            List<AutoMatchCandidateOption> result = new List<AutoMatchCandidateOption>();
            HashSet<string> keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (AutoMatchExpressionCandidate candidate in (candidates ?? new List<AutoMatchExpressionCandidate>())
                .OrderByDescending(item => item.Score))
            {
                if (candidate == null || String.IsNullOrWhiteSpace(candidate.Expression))
                {
                    continue;
                }

                string key = (candidate.Expression ?? "") + "|" + (candidate.QuantityName ?? "");
                if (!keys.Add(key))
                {
                    continue;
                }

                result.Add(new AutoMatchCandidateOption
                {
                    Expression = candidate.Expression,
                    CellAddress = candidate.CellAddress,
                    DisplayValue = candidate.DisplayValue,
                    ExcelQuantityText = candidate.ExcelQuantityText,
                    QuantityName = candidate.QuantityName,
                    Quantity = candidate.Quantity,
                    FirstRow = candidate.FirstRow,
                    FirstColumn = candidate.FirstColumn
                });
                if (result.Count >= 30)
                {
                    break;
                }
            }

            return result;
        }

        private static decimal ComputeAutoMatchCandidateScore(AiQuotaMatchRow quota, List<QuantityExpressionTerm> terms, List<AutoMatchCellValue> cells, string quantityName, decimal quantityDiff, HashSet<string> usedAddresses, int previousRow, int previousColumn)
        {
            decimal score = 0m;
            score += ComputeAutoMatchTextSimilarity(quantityName, quota.Link == null ? "" : quota.Link.QuotaName) * 2m;
            score += Math.Max(0m, 100m - quantityDiff * 1000m);

            for (int i = 0; i < cells.Count && i < terms.Count; i++)
            {
                decimal termDiff = RelativeDifference(terms[i].Literal, cells[i].Value);
                score += Math.Max(0m, 50m - termDiff * 500m);
            }

            bool sameColumn = cells.Select(cell => cell.Cell.Column).Distinct().Count() == 1;
            if (sameColumn)
            {
                score += 30m;
            }

            bool increasingRows = true;
            for (int i = 1; i < cells.Count; i++)
            {
                if (cells[i].Cell.Row <= cells[i - 1].Cell.Row)
                {
                    increasingRows = false;
                    break;
                }
            }

            if (cells.Count > 1 && increasingRows)
            {
                score += 25m;
                if (cells[cells.Count - 1].Cell.Row - cells[0].Cell.Row <= cells.Count + 2)
                {
                    score += 15m;
                }
            }

            if (previousColumn > 0 && FirstColumn(cells) == previousColumn)
            {
                score += 15m;
            }

            if (previousRow > 0 && cells[0].Cell.Row >= previousRow && cells[0].Cell.Row <= previousRow + 8)
            {
                score += 15m;
            }

            foreach (AutoMatchCellValue cell in cells)
            {
                score += usedAddresses != null && usedAddresses.Contains(cell.Cell.Address) ? -40m : 10m;
            }

            return score;
        }

        private static int FirstColumn(List<AutoMatchCellValue> cells)
        {
            return cells == null || cells.Count == 0 ? 0 : cells[0].Cell.Column;
        }

        private static int CapAutoMatchCount(int current, int factor)
        {
            if (current <= 0 || factor <= 0)
            {
                return 0;
            }

            if (current > 10000 / factor)
            {
                return 10000;
            }

            return Math.Min(10000, current * factor);
        }

        private static List<List<AutoMatchCellValue>> BuildAutoMatchCombinations(List<List<AutoMatchCellValue>> termCandidates, int limit)
        {
            List<List<AutoMatchCellValue>> combinations = new List<List<AutoMatchCellValue>>();
            combinations.Add(new List<AutoMatchCellValue>());
            foreach (List<AutoMatchCellValue> candidates in termCandidates)
            {
                List<List<AutoMatchCellValue>> next = new List<List<AutoMatchCellValue>>();
                foreach (List<AutoMatchCellValue> existing in combinations)
                {
                    foreach (AutoMatchCellValue candidate in candidates)
                    {
                        List<AutoMatchCellValue> combo = new List<AutoMatchCellValue>(existing);
                        combo.Add(candidate);
                        next.Add(combo);
                        if (next.Count >= limit)
                        {
                            break;
                        }
                    }

                    if (next.Count >= limit)
                    {
                        break;
                    }
                }

                combinations = next;
                if (combinations.Count == 0)
                {
                    break;
                }
            }

            return combinations;
        }

        private static string BuildAutoMatchExpression(List<QuantityExpressionTerm> terms, List<AutoMatchCellValue> cells)
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < terms.Count; i++)
            {
                QuantityExpressionTerm term = terms[i];
                if (i == 0)
                {
                    if (term.Negative)
                    {
                        builder.Append("-");
                    }
                }
                else
                {
                    builder.Append(term.Negative ? "-" : "+");
                }

                builder.Append(cells[i].Cell.Address);
                builder.Append(term.Suffix ?? "");
            }

            return builder.ToString();
        }

        private static bool TryEvaluateAutoMatchExpression(AiExcelSelectionContext context, string expression, out string displayValue, out decimal quantity, out string error)
        {
            return TryEvaluateAutoMatchExpression(context, expression, true, out displayValue, out quantity, out error);
        }

        private static bool TryEvaluateAutoMatchExpression(AiExcelSelectionContext context, string expression, bool allowWorkbookFallback, out string displayValue, out decimal quantity, out string error)
        {
            displayValue = null;
            quantity = 0m;
            error = null;
            Dictionary<string, AiExcelCell> cells = context == null
                ? new Dictionary<string, AiExcelCell>(StringComparer.OrdinalIgnoreCase)
                : context.CellByAddress;

            string resolved = context == null ? NormalizeExpressionOperators(expression) : context.NormalizeMergedExpression(expression);
            bool allResolved = true;
            foreach (string address in ExtractCellAddressesFromExpression(resolved).OrderByDescending(value => value.Length))
            {
                AiExcelCell cell;
                if (!cells.TryGetValue(address, out cell))
                {
                    allResolved = false;
                    break;
                }

                // 快照内的空值单元格按 0 参与计算（与模板铺量 emptyCellAsZero 语义一致），
                // 支持把数量为0的定额手动绑定到尚未填数的 Excel 格。
                decimal cellValue = 0m;
                string parseError;
                if (!String.IsNullOrWhiteSpace(cell.Text) && !TryEvaluateDecimal(cell.Text, out cellValue, out parseError))
                {
                    allResolved = false;
                    break;
                }

                resolved = resolved.Replace(address, FormatAiMatchDecimal(cellValue));
            }

            if (allResolved && TryEvaluateDecimal(resolved, out quantity, out error))
            {
                displayValue = FormatAiMatchDecimal(quantity);
                return true;
            }

            // 批量匹配循环里禁用实盘回退，避免快照未命中时反复读 Excel 拖慢整体匹配。
            if (allowWorkbookFallback && context != null && TryEvaluateWorkbookExpression(context.WorkbookPath, context.WorksheetName, expression, out displayValue, out quantity, out error))
            {
                displayValue = FormatAiMatchDecimal(quantity);
                return true;
            }

            return false;
        }

        private static string BuildExcelQuantityTextForExpression(AiExcelSelectionContext context, string expression)
        {
            if (context == null || String.IsNullOrWhiteSpace(expression))
            {
                return "";
            }

            Dictionary<string, AiExcelCell> cells = context.CellByAddress;
            string normalized = context.NormalizeMergedExpression(expression);
            MatchCollection matches = Regex.Matches(normalized, @"[A-Z]+\d+");
            if (matches.Count == 0)
            {
                return "";
            }

            StringBuilder builder = new StringBuilder();
            int previousEnd = 0;
            for (int i = 0; i < matches.Count; i++)
            {
                Match match = matches[i];
                string address = NormalizeCellAddress(match.Value);
                CellRef parsed;
                if (!TryParseCellAddress(address, out parsed))
                {
                    previousEnd = match.Index + match.Length;
                    continue;
                }

                builder.Append(normalized.Substring(previousEnd, match.Index - previousEnd));
                string text = "";
                bool read = false;
                AiExcelCell cell;
                if (cells.TryGetValue(address, out cell))
                {
                    text = cell.Text ?? "";
                    read = true;
                }
                else
                {
                    string readError;
                    read = TryReadWorkbookCellValue(context.WorkbookPath, context.WorksheetName, address, out text, out readError);
                }

                builder.Append(!read ? address : (String.IsNullOrWhiteSpace(text) ? "0" : text.Trim()));
                previousEnd = match.Index + match.Length;
            }

            builder.Append(normalized.Substring(previousEnd));
            return builder.ToString();
        }

        private sealed class OpenSpreadsheetWorkbookInfo
        {
            public string FullName;
            public string DisplayName;
            public List<string> SheetNames = new List<string>();
            public string ActiveSheetName;
            public bool IsActive;
            public bool IsTemplateSource;

            public override string ToString()
            {
                return DisplayName ?? FullName ?? "";
            }
        }

        private static string BuildOpenWorkbookDisplayName(string fullName, bool isTemplateSource)
        {
            return Path.GetFileName(fullName ?? "");
        }

        private static bool TryListOpenSpreadsheetWorkbooks(
            out List<OpenSpreadsheetWorkbookInfo> workbooks, out string error)
        {
            workbooks = new List<OpenSpreadsheetWorkbookInfo>();
            error = null;
            Dictionary<string, OpenSpreadsheetWorkbookInfo> byPath =
                new Dictionary<string, OpenSpreadsheetWorkbookInfo>(StringComparer.OrdinalIgnoreCase);
            List<string> diagnostics = new List<string>();

            object activeApplication = GetActiveSpreadsheetApplication();
            if (activeApplication != null)
            {
                CollectOpenWorkbooksFromApplication(activeApplication, true, byPath, diagnostics);
            }

            List<IntPtr> spreadsheetWindows = new List<IntPtr>();
            try
            {
                EnumWindows(delegate(IntPtr hWnd, IntPtr lParam)
                {
                    CollectExcelChildWindows(hWnd, spreadsheetWindows);
                    return true;
                }, IntPtr.Zero);
            }
            catch (Exception ex)
            {
                diagnostics.Add("EnumWindows: " + ex.Message);
            }

            foreach (IntPtr childWindow in spreadsheetWindows)
            {
                try
                {
                    Guid dispatchGuid = new Guid("00020400-0000-0000-C000-000000000046");
                    object nativeObject;
                    int hr = AccessibleObjectFromWindow(childWindow, ObjIdNativeOm, ref dispatchGuid, out nativeObject);
                    if (hr != 0 || nativeObject == null) continue;
                    dynamic window = nativeObject;
                    object application = window.Application;
                    if (application != null)
                    {
                        CollectOpenWorkbooksFromApplication(application, false, byPath, diagnostics);
                    }
                }
                catch (Exception ex)
                {
                    diagnostics.Add("WindowObject " + childWindow.ToString("X") + ": " + ex.Message);
                }
            }

            workbooks = byPath.Values
                .OrderByDescending(item => item.IsActive)
                .ThenBy(item => item.DisplayName, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
            if (workbooks.Count == 0)
            {
                error = diagnostics.Count == 0
                    ? "没有找到已打开并保存的 Excel/WPS 工作簿。"
                    : "没有找到已打开并保存的 Excel/WPS 工作簿：" + diagnostics[0];
                return false;
            }
            return true;
        }

        private static void CollectOpenWorkbooksFromApplication(object application, bool markActiveWorkbook,
            Dictionary<string, OpenSpreadsheetWorkbookInfo> byPath, List<string> diagnostics)
        {
            if (application == null) return;
            try
            {
                dynamic excel = application;
                string activeWorkbookPath = "";
                try
                {
                    dynamic activeWorkbook = excel.ActiveWorkbook;
                    if (activeWorkbook != null)
                    {
                        activeWorkbookPath = NormalizeTemplateWorkbookPath(
                            Convert.ToString(activeWorkbook.FullName, CultureInfo.InvariantCulture));
                    }
                }
                catch
                {
                }

                dynamic books = excel.Workbooks;
                int count = Convert.ToInt32(books.Count, CultureInfo.InvariantCulture);
                for (int i = 1; i <= count; i++)
                {
                    try
                    {
                        dynamic workbook = books.Item(i);
                        string fullName = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture);
                        if (String.IsNullOrWhiteSpace(fullName) || !File.Exists(fullName)) continue;
                        fullName = Path.GetFullPath(fullName);

                        OpenSpreadsheetWorkbookInfo info;
                        if (!byPath.TryGetValue(fullName, out info))
                        {
                            info = new OpenSpreadsheetWorkbookInfo();
                            info.FullName = fullName;
                            info.DisplayName = BuildOpenWorkbookDisplayName(fullName, false);
                            try
                            {
                                dynamic sheets = workbook.Worksheets;
                                int sheetCount = Convert.ToInt32(sheets.Count, CultureInfo.InvariantCulture);
                                for (int sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
                                {
                                    dynamic sheet = sheets.Item(sheetIndex);
                                    info.SheetNames.Add(Convert.ToString(sheet.Name, CultureInfo.InvariantCulture));
                                }
                                dynamic activeSheet = workbook.ActiveSheet;
                                if (activeSheet != null)
                                {
                                    info.ActiveSheetName = Convert.ToString(activeSheet.Name, CultureInfo.InvariantCulture);
                                }
                            }
                            catch (Exception ex)
                            {
                                diagnostics.Add(Path.GetFileName(fullName) + " sheets: " + ex.Message);
                            }
                            byPath[fullName] = info;
                        }

                        if (markActiveWorkbook && String.Equals(NormalizeTemplateWorkbookPath(fullName), activeWorkbookPath,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            info.IsActive = true;
                        }
                    }
                    catch (Exception ex)
                    {
                        diagnostics.Add("Workbook " + i.ToString(CultureInfo.InvariantCulture) + ": " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                diagnostics.Add("Workbooks: " + ex.Message);
            }
        }

        private static bool TryListActiveWorkbookSheets(out List<string> sheetNames, out string activeSheetName, out string error)
        {
            sheetNames = new List<string>();
            activeSheetName = "";
            error = null;
            dynamic excel = null;
            try
            {
                excel = GetActiveSpreadsheetApplication();
                if (excel == null)
                {
                    error = BuildExcelConnectError("\u6ca1\u6709\u627e\u5230\u6b63\u5728\u8fd0\u884c\u7684 Excel/WPS \u8868\u683c");
                    return false;
                }

                dynamic workbook = excel.ActiveWorkbook;
                dynamic activeSheet = excel.ActiveSheet;
                if (workbook == null || activeSheet == null)
                {
                    error = BuildExcelConnectError("\u5df2\u7ecf\u8fde\u63a5\u5230 Excel/WPS\uff0c\u4f46\u6ca1\u6709\u8bfb\u5230\u5f53\u524d\u5de5\u4f5c\u7c3f\u6216\u5de5\u4f5c\u8868");
                    return false;
                }

                activeSheetName = Convert.ToString(activeSheet.Name, CultureInfo.InvariantCulture);
                int count = Convert.ToInt32(workbook.Worksheets.Count, CultureInfo.InvariantCulture);
                for (int i = 1; i <= count; i++)
                {
                    dynamic sheet = workbook.Worksheets[i];
                    sheetNames.Add(Convert.ToString(sheet.Name, CultureInfo.InvariantCulture));
                }

                return sheetNames.Count > 0;
            }
            catch (COMException ex)
            {
                ClearCachedSpreadsheetApplication(excel);
                error = BuildExcelConnectError("\u8bfb\u53d6 Excel/WPS \u5de5\u4f5c\u8868\u5217\u8868\u5931\u8d25\uff1a" + ex.Message);
                return false;
            }
            catch (Exception ex)
            {
                ClearCachedSpreadsheetApplication(excel);
                error = BuildExcelConnectError("\u8bfb\u53d6 Excel/WPS \u5de5\u4f5c\u8868\u5217\u8868\u5931\u8d25\uff1a" + ex.Message);
                return false;
            }
        }

        private static bool TryReadWorksheetCellsForAutoMatch(string sheetName, List<int> targetColumns, out AiExcelSelectionContext context, out string error)
        {
            context = null;
            error = null;
            dynamic excel = null;
            try
            {
                excel = GetActiveSpreadsheetApplication();
                if (excel == null)
                {
                    error = BuildExcelConnectError("\u6ca1\u6709\u627e\u5230\u6b63\u5728\u8fd0\u884c\u7684 Excel/WPS \u8868\u683c");
                    return false;
                }

                dynamic workbook = excel.ActiveWorkbook;
                if (workbook == null)
                {
                    error = BuildExcelConnectError("\u5df2\u7ecf\u8fde\u63a5\u5230 Excel/WPS\uff0c\u4f46\u6ca1\u6709\u8bfb\u5230\u5f53\u524d\u5de5\u4f5c\u7c3f");
                    return false;
                }

                dynamic sheet = String.IsNullOrWhiteSpace(sheetName) ? excel.ActiveSheet : workbook.Worksheets[sheetName];
                dynamic usedRange = sheet.UsedRange;
                if (sheet == null || usedRange == null)
                {
                    error = "\u6ca1\u6709\u8bfb\u5230\u76ee\u6807\u5de5\u4f5c\u8868\u6216 UsedRange\u3002";
                    return false;
                }

                int rowCount = Convert.ToInt32(usedRange.Rows.Count, CultureInfo.InvariantCulture);
                int colCount = Convert.ToInt32(usedRange.Columns.Count, CultureInfo.InvariantCulture);
                int firstRow = Convert.ToInt32(usedRange.Row, CultureInfo.InvariantCulture);
                int firstColumn = Convert.ToInt32(usedRange.Column, CultureInfo.InvariantCulture);
                if (rowCount <= 0 || colCount <= 0)
                {
                    error = "\u76ee\u6807\u5de5\u4f5c\u8868\u6ca1\u6709\u53ef\u8bfb\u53d6\u7684\u5185\u5bb9\u3002";
                    return false;
                }

                int lastColumn = firstColumn + colCount - 1;
                bool hasTargetColumns = targetColumns != null && targetColumns.Count > 0;
                if (!hasTargetColumns && rowCount * colCount > AutoMatchCellLimit)
                {
                    error = "\u5f53\u524d\u5de5\u4f5c\u8868\u8303\u56f4\u8f83\u5927\uff0c\u8bf7\u5148\u586b\u5199\u76ee\u6807\u5217\uff08\u5982 E \u6216 E,F\uff09\u540e\u518d\u5339\u914d\u3002";
                    return false;
                }

                int readLastColumn = hasTargetColumns ? Math.Min(lastColumn, Math.Max(firstColumn, targetColumns.Max())) : lastColumn;
                if (readLastColumn < firstColumn)
                {
                    error = "\u76ee\u6807\u5217\u4e0d\u5728\u5f53\u524d\u5de5\u4f5c\u8868 UsedRange \u5185\u3002";
                    return false;
                }

                context = new AiExcelSelectionContext();
                context.WorkbookPath = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture);
                context.WorksheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture);
                HashSet<int> visibleColumns = BuildAutoMatchVisibleColumns(sheet, context.WorkbookPath, context.WorksheetName, firstColumn, readLastColumn);
                if (visibleColumns.Count == 0)
                {
                    error = "\u76ee\u6807\u5de5\u4f5c\u8868\u8303\u56f4\u5185\u6ca1\u6709\u53ef\u89c1\u5217\u3002";
                    return false;
                }

                int readColCount = readLastColumn - firstColumn + 1;
                LoadAutoMatchMergedRanges(context, firstRow, firstColumn, rowCount, readColCount);
                bool ok;
                if (rowCount * readColCount <= AutoMatchCellLimit)
                {
                    ok = TryReadWorksheetRangeCells(sheet, context, firstRow, firstColumn, rowCount, readColCount);
                }
                else
                {
                    ok = TryReadWorksheetRangeCellsInBlocks(sheet, context, firstRow, firstColumn, rowCount, readColCount);
                }

                if (!ok)
                {
                    context.Cells.Clear();
                    ok = TryReadWorksheetCellsOneByOne(sheet, context, firstRow, firstColumn, rowCount, readColCount, AutoMatchFallbackCellLimit);
                }

                context.Cells.RemoveAll(cell => cell == null || !visibleColumns.Contains(cell.Column));
                if (!ok || context.Cells.Count == 0)
                {
                    error = "\u76ee\u6807\u5de5\u4f5c\u8868\u6ca1\u6709\u53ef\u8bfb\u53d6\u7684\u5185\u5bb9\u3002";
                    return false;
                }

                if (!HasAutoMatchNumberCell(context, targetColumns))
                {
                    // \u76ee\u6807\u5217\u6682\u65f6\u5168\u4e3a\u7a7a\u65f6\u4e5f\u5141\u8bb8\u751f\u6210\u5feb\u7167\uff1a\u6570\u91cf\u4e3a0\u7684\u5b9a\u989d\u9700\u8981\u624b\u52a8\u7ed1\u5b9a\u5230\u7a7a\u503c\u5355\u5143\u683c\u3002
                    Log("Auto match snapshot has no numeric cells in target columns.");
                }

                return true;
            }
            catch (COMException ex)
            {
                ClearCachedSpreadsheetApplication(excel);
                error = BuildExcelConnectError("\u8bfb\u53d6 Excel/WPS \u5de5\u4f5c\u8868\u5931\u8d25\uff1a" + ex.Message);
                return false;
            }
            catch (Exception ex)
            {
                ClearCachedSpreadsheetApplication(excel);
                error = BuildExcelConnectError("\u8bfb\u53d6 Excel/WPS \u5de5\u4f5c\u8868\u5931\u8d25\uff1a" + ex.Message);
                return false;
            }
        }

        private static void LoadAutoMatchMergedRanges(AiExcelSelectionContext context, int firstRow, int firstColumn, int rowCount, int colCount)
        {
            if (context == null || String.IsNullOrWhiteSpace(context.WorkbookPath) || !File.Exists(context.WorkbookPath))
            {
                return;
            }

            int lastRow = firstRow + rowCount - 1;
            int lastColumn = firstColumn + colCount - 1;
            foreach (ExcelMergedRegion region in ReadExcelMergedRegions(context.WorkbookPath, context.WorksheetName))
            {
                if (region.LastRow < firstRow || region.FirstRow > lastRow ||
                    region.LastColumn < firstColumn || region.FirstColumn > lastColumn)
                {
                    continue;
                }

                context.AddMergedRange(
                    Math.Max(region.FirstRow, firstRow),
                    Math.Min(region.LastRow, lastRow),
                    Math.Max(region.FirstColumn, firstColumn),
                    Math.Min(region.LastColumn, lastColumn),
                    region.FirstRow,
                    region.FirstColumn);
            }
        }

        private static HashSet<int> BuildAutoMatchVisibleColumns(dynamic sheet, string workbookPath, string worksheetName, int firstColumn, int lastColumn)
        {
            HashSet<int> columns = new HashSet<int>();
            for (int column = firstColumn; column <= lastColumn; column++)
            {
                if (!IsExcelLinkColumnHidden(sheet, workbookPath, worksheetName, column))
                {
                    columns.Add(column);
                }
            }

            return columns;
        }

        private static bool HasAutoMatchNumberCell(AiExcelSelectionContext context, List<int> targetColumns)
        {
            HashSet<int> targetSet = targetColumns != null && targetColumns.Count > 0
                ? new HashSet<int>(targetColumns)
                : null;
            return context.Cells.Any(cell => cell.IsNumber && context.IsCellInTargetColumns(cell, targetSet));
        }

        private static bool TryReadWorksheetRangeCells(dynamic sheet, AiExcelSelectionContext context, int firstRow, int firstColumn, int rowCount, int colCount)
        {
            try
            {
                string startAddress = ColumnNumberToName(firstColumn) + firstRow.ToString(CultureInfo.InvariantCulture);
                string endAddress = ColumnNumberToName(firstColumn + colCount - 1) + (firstRow + rowCount - 1).ToString(CultureInfo.InvariantCulture);
                dynamic range = sheet.Range[startAddress + ":" + endAddress];
                AddWorksheetRangeValues(context, firstRow, firstColumn, rowCount, colCount, range.Value2);
                return context.Cells.Count > 0;
            }
            catch (Exception ex)
            {
                Log("Auto match bulk worksheet read failed: " + ex.Message);
                return false;
            }
        }

        private static bool TryReadWorksheetRangeCellsInBlocks(dynamic sheet, AiExcelSelectionContext context, int firstRow, int firstColumn, int rowCount, int colCount)
        {
            try
            {
                for (int offset = 0; offset < rowCount; offset += AutoMatchReadBlockRows)
                {
                    int rows = Math.Min(AutoMatchReadBlockRows, rowCount - offset);
                    if (!TryReadWorksheetRangeCells(sheet, context, firstRow + offset, firstColumn, rows, colCount))
                    {
                        return false;
                    }
                }

                return context.Cells.Count > 0;
            }
            catch (Exception ex)
            {
                Log("Auto match block worksheet read failed: " + ex.Message);
                return false;
            }
        }

        private static bool TryReadWorksheetCellsOneByOne(dynamic sheet, AiExcelSelectionContext context, int firstRow, int firstColumn, int rowCount, int colCount, int limit)
        {
            try
            {
                int read = 0;
                for (int row = 0; row < rowCount && read < limit; row++)
                {
                    for (int col = 0; col < colCount && read < limit; col++)
                    {
                        int actualRow = firstRow + row;
                        int actualColumn = firstColumn + col;
                        string address = ColumnNumberToName(actualColumn) + actualRow.ToString(CultureInfo.InvariantCulture);
                        dynamic range = sheet.Range[address];
                        AddAiExcelCell(context, actualRow, actualColumn, range.Value2);
                        read++;
                    }
                }

                return context.Cells.Count > 0;
            }
            catch (Exception ex)
            {
                Log("Auto match cell-by-cell worksheet read failed: " + ex.Message);
                return false;
            }
        }

        private static void AddWorksheetRangeValues(AiExcelSelectionContext context, int firstRow, int firstColumn, int rowCount, int colCount, object rawValues)
        {
            if (rowCount == 1 && colCount == 1)
            {
                AddAiExcelCell(context, firstRow, firstColumn, rawValues);
                return;
            }

            Array values = rawValues as Array;
            if (values == null)
            {
                return;
            }

            for (int row = 1; row <= rowCount; row++)
            {
                for (int col = 1; col <= colCount; col++)
                {
                    object value = GetWorksheetRangeArrayValue(values, row, col, rowCount, colCount);
                    AddAiExcelCell(context, firstRow + row - 1, firstColumn + col - 1, value);
                }
            }
        }

        private static object GetWorksheetRangeArrayValue(Array values, int row, int col, int rowCount, int colCount)
        {
            try
            {
                if (values.Rank == 2)
                {
                    return values.GetValue(row, col);
                }

                if (values.Rank == 1)
                {
                    return values.GetValue(rowCount == 1 ? col : row);
                }
            }
            catch
            {
            }

            return null;
        }

        private sealed class AutoMatchDialog : Form
        {
            private readonly Form mainForm;
            private readonly SqlConnection conn;
            private readonly ComboBox sheetBox;
            private readonly TextBox targetColumnText;
            private readonly DataGridView grid;
            private readonly Label status;
            private readonly CheckBox manualMatchButton;
            private readonly Button startButton;
            private readonly System.Windows.Forms.Timer manualMatchTimer;
            private readonly SplitContainer split;
            private readonly TreeView itemTree;
            private Dictionary<string, string> chapterNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            private string currentTreeScope = "";
            private bool rebuildingTree;
            private bool updatingTreeChecks;
            private AiExcelSelectionContext currentContext;
            private string currentContextKey;
            private string lastManualCellKey;
            private bool manualBaselinePending;
            private bool matchingInProgress;
            private bool updatingQuantityNameCell;
            private bool suppressManualClosedStatus;
            private List<int> pendingBatchCheckRows;
            private bool applyingBatchCheck;
            private List<AiMatchPreviewItem> items = new List<AiMatchPreviewItem>();

            public event Action<List<AiMatchPreviewItem>> Accepted;
            public event Action Cancelled;

            public AutoMatchDialog(Form ownerForm, SqlConnection projectConnection)
            {
                mainForm = ownerForm;
                conn = projectConnection;
                Text = "\u81ea\u52a8\u5339\u914dExcel\u5de5\u7a0b\u91cf";
                StartPosition = FormStartPosition.CenterParent;
                Size = new System.Drawing.Size(1180, 560);
                MinimumSize = new System.Drawing.Size(1000, 420);
                MinimizeBox = false;

                Panel top = new Panel();
                top.Dock = DockStyle.Top;
                top.Height = 46;
                top.Padding = new Padding(8);

                Label sheetLabel = new Label();
                sheetLabel.Text = "\u5de5\u4f5c\u8868";
                sheetLabel.Left = 8;
                sheetLabel.Top = 13;
                sheetLabel.Width = 54;

                sheetBox = new ComboBox();
                sheetBox.DropDownStyle = ComboBoxStyle.DropDownList;
                sheetBox.Left = 66;
                sheetBox.Top = 9;
                sheetBox.Width = 180;

                Label columnLabel = new Label();
                columnLabel.Text = "\u76ee\u6807\u5217";
                columnLabel.Left = 260;
                columnLabel.Top = 13;
                columnLabel.Width = 54;

                targetColumnText = new TextBox();
                targetColumnText.Left = 318;
                targetColumnText.Top = 9;
                targetColumnText.Width = 160;

                startButton = new Button();
                startButton.Text = "\u5f00\u59cb\u5339\u914d";
                startButton.Left = 492;
                startButton.Top = 8;
                startButton.Width = 90;
                startButton.Click += delegate { StartMatch(); };

                manualMatchButton = new CheckBox();
                manualMatchButton.Appearance = Appearance.Button;
                manualMatchButton.Text = "\u624b\u52a8\u5339\u914d";
                manualMatchButton.Left = 590;
                manualMatchButton.Top = 8;
                manualMatchButton.Width = 90;
                manualMatchButton.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
                manualMatchButton.CheckedChanged += delegate { ToggleManualMatch(); };

                top.Controls.Add(sheetLabel);
                top.Controls.Add(sheetBox);
                top.Controls.Add(columnLabel);
                top.Controls.Add(targetColumnText);
                top.Controls.Add(startButton);
                top.Controls.Add(manualMatchButton);

                grid = new DataGridView();
                grid.Dock = DockStyle.Fill;
                grid.AllowUserToAddRows = false;
                grid.AllowUserToDeleteRows = false;
                grid.RowHeadersVisible = false;
                grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                grid.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                grid.CellEndEdit += delegate(object sender, DataGridViewCellEventArgs e)
                {
                    if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && String.Equals(grid.Columns[e.ColumnIndex].Name, "Expression", StringComparison.Ordinal))
                    {
                        UpdateExpressionFromRow(grid.Rows[e.RowIndex]);
                    }
                };
                grid.CellClick += delegate(object sender, DataGridViewCellEventArgs e)
                {
                    if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && String.Equals(grid.Columns[e.ColumnIndex].Name, "QuantityName", StringComparison.Ordinal))
                    {
                        PrepareQuantityNameDropDown(grid.Rows[e.RowIndex]);
                        grid.CurrentCell = grid.Rows[e.RowIndex].Cells[e.ColumnIndex];
                        grid.BeginEdit(true);
                        ComboBox combo = grid.EditingControl as ComboBox;
                        if (combo != null)
                        {
                            combo.DroppedDown = true;
                        }
                    }
                };
                grid.CurrentCellDirtyStateChanged += delegate
                {
                    if (grid.IsCurrentCellDirty)
                    {
                        grid.CommitEdit(DataGridViewDataErrorContexts.Commit);
                    }
                };
                grid.CellMouseDown += delegate(object sender, DataGridViewCellMouseEventArgs e)
                {
                    // 点击勾选框前记录多选行：随后的点击会把选择收拢成单行，先存下来供批量勾选用。
                    if (e.RowIndex >= 0 && e.ColumnIndex >= 0 &&
                        String.Equals(grid.Columns[e.ColumnIndex].Name, "Checked", StringComparison.Ordinal) &&
                        grid.SelectedRows.Count > 1 &&
                        grid.Rows[e.RowIndex].Selected)
                    {
                        pendingBatchCheckRows = grid.SelectedRows.Cast<DataGridViewRow>().Select(row => row.Index).ToList();
                    }
                    else
                    {
                        pendingBatchCheckRows = null;
                    }
                };
                grid.CellValueChanged += delegate(object sender, DataGridViewCellEventArgs e)
                {
                    if (e.RowIndex < 0 || e.ColumnIndex < 0)
                    {
                        return;
                    }

                    if (String.Equals(grid.Columns[e.ColumnIndex].Name, "QuantityName", StringComparison.Ordinal))
                    {
                        ApplyQuantityNameOption(grid.Rows[e.RowIndex], Convert.ToString(grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value, CultureInfo.InvariantCulture));
                    }
                    else if (String.Equals(grid.Columns[e.ColumnIndex].Name, "Checked", StringComparison.Ordinal))
                    {
                        ApplyBatchCheck(e.RowIndex);
                    }
                };
                grid.DataError += delegate(object sender, DataGridViewDataErrorEventArgs e)
                {
                    Log("Auto match grid data error: " + (e.Exception == null ? "unknown" : e.Exception.Message));
                    e.ThrowException = false;
                };
                grid.SelectionChanged += delegate
                {
                    lastManualCellKey = "";
                    manualBaselinePending = true;
                };
                BuildGridColumns();

                Button singleBind = new Button();
                singleBind.Text = "\u5355\u4e2a\u7ed1\u5b9a";
                singleBind.Width = 90;
                singleBind.Click += delegate { AcceptCurrentItem(); };

                Button ok = new Button();
                ok.Text = "\u5168\u90e8\u7ed1\u5b9a";
                ok.Width = 90;
                ok.Click += delegate { AcceptCheckedItems(); };

                Button cancel = new Button();
                cancel.Text = "\u53d6\u6d88";
                cancel.Width = 75;
                cancel.Click += delegate
                {
                    if (Cancelled != null)
                    {
                        Cancelled();
                    }

                    DialogResult = DialogResult.Cancel;
                    Close();
                };

                FlowLayoutPanel buttons = new FlowLayoutPanel();
                buttons.Dock = DockStyle.Bottom;
                buttons.Height = 44;
                buttons.FlowDirection = FlowDirection.RightToLeft;
                buttons.Padding = new Padding(8);
                buttons.Controls.Add(cancel);
                buttons.Controls.Add(ok);
                buttons.Controls.Add(singleBind);

                status = new Label();
                status.Dock = DockStyle.Bottom;
                status.Height = 26;
                status.Padding = new Padding(8, 2, 8, 2);

                // —— 左侧条目树：点节点过滤右侧预览，勾选节点整枝勾选/取消，可拖动分栏 ——
                itemTree = new TreeView();
                itemTree.Dock = DockStyle.Fill;
                itemTree.CheckBoxes = true;
                itemTree.HideSelection = false;
                itemTree.AfterSelect += delegate { OnItemTreeScopeChanged(); };
                itemTree.AfterCheck += delegate(object sender, TreeViewEventArgs e) { OnItemTreeChecked(e.Node); };

                split = new SplitContainer();
                split.Dock = DockStyle.Fill;
                split.Orientation = Orientation.Vertical;
                split.Panel1.Controls.Add(itemTree);
                split.Panel2.Controls.Add(grid);

                Controls.Add(split);
                Controls.Add(buttons);
                Controls.Add(status);
                Controls.Add(top);

                manualMatchTimer = new System.Windows.Forms.Timer();
                manualMatchTimer.Interval = 600;
                manualMatchTimer.Tick += delegate { PollManualMatchCell(); };
                FormClosed += delegate
                {
                    manualMatchTimer.Stop();
                    manualMatchTimer.Dispose();
                };
                Shown += delegate { SetInitialSplitterDistance(); };

                LoadSheets();
            }

            private void SetInitialSplitterDistance()
            {
                if (split.Width <= 0)
                {
                    return;
                }

                split.Panel1MinSize = 80;
                split.Panel2MinSize = Math.Min(400, Math.Max(25, split.Width - split.Panel1MinSize));
                int max = split.Width - split.Panel2MinSize;
                if (max < split.Panel1MinSize)
                {
                    return;
                }

                split.SplitterDistance = Math.Min(Math.Max(220, split.Panel1MinSize), max);
            }

            private void BuildGridColumns()
            {
                DataGridViewCheckBoxColumn checkedColumn = new DataGridViewCheckBoxColumn();
                checkedColumn.Name = "Checked";
                checkedColumn.HeaderText = "\u7ed1\u5b9a";
                checkedColumn.FillWeight = 42;
                grid.Columns.Add(checkedColumn);
                grid.Columns.Add("QuotaCode", "\u5b9a\u989d\u7f16\u53f7");
                grid.Columns.Add("QuotaName", "\u5b9a\u989d\u540d\u79f0");
                grid.Columns.Add("QuotaUnit", "\u5355\u4f4d");
                grid.Columns.Add("CurrentQuantity", "\u5f53\u524d\u5de5\u7a0b\u91cf");
                grid.Columns.Add("Expression", "\u5339\u914d\u8868\u8fbe\u5f0f");
                grid.Columns.Add("Value", "Excel\u5de5\u7a0b\u91cf");
                grid.Columns.Add("QuantityName", "\u5de5\u7a0b\u91cf\u540d\u79f0");
                grid.Columns.Add("Status", "\u72b6\u6001");
                grid.Columns["QuotaCode"].FillWeight = 70;
                grid.Columns["QuotaName"].FillWeight = 190;
                grid.Columns["QuotaUnit"].FillWeight = 52;
                grid.Columns["CurrentQuantity"].FillWeight = 80;
                grid.Columns["Expression"].FillWeight = 115;
                grid.Columns["Value"].FillWeight = 70;
                grid.Columns["QuantityName"].FillWeight = 160;
                grid.Columns["Status"].FillWeight = 90;

                foreach (DataGridViewColumn column in grid.Columns)
                {
                    column.ReadOnly = true;
                }

                grid.Columns["Checked"].ReadOnly = false;
                grid.Columns["Expression"].ReadOnly = false;
                grid.Columns["QuantityName"].ReadOnly = false;
            }

            private void LoadSheets()
            {
                List<string> sheetNames;
                string activeSheetName;
                string error;
                if (!TryListActiveWorkbookSheets(out sheetNames, out activeSheetName, out error))
                {
                    status.Text = error;
                    return;
                }

                sheetBox.Items.Clear();
                foreach (string name in sheetNames)
                {
                    sheetBox.Items.Add(name);
                }

                if (!String.IsNullOrWhiteSpace(activeSheetName) && sheetBox.Items.Contains(activeSheetName))
                {
                    sheetBox.SelectedItem = activeSheetName;
                }
                else if (sheetBox.Items.Count > 0)
                {
                    sheetBox.SelectedIndex = 0;
                }

                status.Text = "\u5df2\u8bfb\u53d6\u5de5\u4f5c\u8868\u5217\u8868\uff0c\u8bf7\u9009\u62e9\u76ee\u6807\u5de5\u4f5c\u8868\u540e\u5f00\u59cb\u5339\u914d\u3002";
            }

            private void StartMatch()
            {
                if (matchingInProgress)
                {
                    return;
                }

                grid.EndEdit();
                List<int> targetColumns;
                string error;

                List<AiQuotaMatchRow> quotas = LoadCurrentSelectedQuotas();
                if (quotas.Count == 0)
                {
                    status.Text = "\u8bf7\u5148\u5728\u5b9a\u989d\u8f93\u5165\u8868\u4e2d\u6846\u9009\u5b9a\u989d\uff0c\u6216\u5728\u5de6\u4fa7\u70b9\u9009\u8981\u5339\u914d\u7684\u7ae0\u8282\u6761\u76ee\u3002";
                    return;
                }

                bool reuseContext;
                status.Text = "\u6b63\u5728\u51c6\u5907Excel\u5feb\u7167...";
                UseWaitCursor = true;
                Application.DoEvents();
                bool snapshotReady = false;
                try
                {
                    if (!EnsureCurrentAutoMatchSnapshot(out targetColumns, out reuseContext, out error))
                    {
                        status.Text = error;
                        return;
                    }

                    snapshotReady = true;
                }
                finally
                {
                    if (!snapshotReady)
                    {
                        UseWaitCursor = false;
                    }
                }

                status.Text = reuseContext ? "\u6b63\u5728\u4f7f\u7528\u5df2\u8bfb\u53d6\u7684Excel\u5feb\u7167\u5339\u914d..." : "\u6b63\u5728\u8bfb\u53d6Excel\u5e76\u672c\u5730\u5339\u914d...";
                HashSet<long> alreadyBoundSequences = LoadAlreadyBoundSequences();
                AiExcelSelectionContext matchContext = currentContext;
                List<int> matchColumns = targetColumns;
                bool reusedSnapshot = reuseContext;
                matchingInProgress = true;
                startButton.Enabled = false;
                System.Threading.ThreadPool.QueueUserWorkItem(delegate
                {
                    List<AiMatchPreviewItem> matched = null;
                    string matchError = null;
                    try
                    {
                        matched = BuildAutoMatchPreviewItems(quotas, matchContext, matchColumns, alreadyBoundSequences);
                    }
                    catch (Exception ex)
                    {
                        matchError = ex.Message;
                        Log("Auto match background build failed: " + ex);
                    }

                    try
                    {
                        if (IsDisposed || !IsHandleCreated)
                        {
                            return;
                        }

                        BeginInvoke(new Action(delegate
                        {
                            matchingInProgress = false;
                            startButton.Enabled = true;
                            UseWaitCursor = false;
                            if (matchError != null)
                            {
                                status.Text = "\u5339\u914d\u5931\u8d25\uff1a" + matchError;
                                return;
                            }

                            items = matched ?? new List<AiMatchPreviewItem>();
                            RebuildItemTree();
                            FillGrid();
                            int checkedCount = items.Count(item => item.Checked);
                            status.Text = "\u5339\u914d\u5b8c\u6210\uff1a\u5171 " + items.Count.ToString(CultureInfo.InvariantCulture) + " \u6761\uff0c\u9ed8\u8ba4\u52fe\u9009 " + checkedCount.ToString(CultureInfo.InvariantCulture) + " \u6761\u3002" + (reusedSnapshot ? "\u5df2\u590d\u7528Excel\u5feb\u7167\u3002" : "");
                        }));
                    }
                    catch (Exception ex)
                    {
                        Log("Auto match UI callback failed: " + ex.Message);
                    }
                });
            }

            private bool EnsureCurrentAutoMatchSnapshot(out List<int> targetColumns, out bool reuseContext, out string error)
            {
                targetColumns = null;
                reuseContext = false;
                error = null;
                if (!TryParseTargetColumns(targetColumnText.Text, out targetColumns, out error))
                {
                    return false;
                }

                if (sheetBox.SelectedItem == null)
                {
                    error = "\u8bf7\u5148\u9009\u62e9\u5de5\u4f5c\u8868\u3002";
                    return false;
                }

                string sheetName = Convert.ToString(sheetBox.SelectedItem, CultureInfo.InvariantCulture);
                string contextKey = BuildAutoMatchSnapshotKey(sheetName, targetColumns);
                reuseContext = currentContext != null && String.Equals(currentContextKey, contextKey, StringComparison.OrdinalIgnoreCase);
                if (reuseContext)
                {
                    return true;
                }

                AiExcelSelectionContext context;
                if (!TryReadWorksheetCellsForAutoMatch(sheetName, targetColumns, out context, out error))
                {
                    return false;
                }

                currentContext = context;
                currentContextKey = contextKey;
                return true;
            }

            private static string BuildAutoMatchSnapshotKey(string sheetName, List<int> targetColumns)
            {
                string columnKey = targetColumns == null || targetColumns.Count == 0
                    ? "*"
                    : String.Join(",", targetColumns.OrderBy(column => column).Select(column => column.ToString(CultureInfo.InvariantCulture)).ToArray());
                return (sheetName ?? "").Trim().ToUpperInvariant() + "|" + columnKey;
            }

            private List<AiQuotaMatchRow> LoadCurrentSelectedQuotas()
            {
                DataGridView quotaGrid = GetField<DataGridView>(mainForm, "dataGridViewDE");
                List<DataGridViewRow> explicitRows = GetExplicitSelectedQuotaRows(quotaGrid);
                if (explicitRows.Count > 1)
                {
                    return BuildAiQuotaRows(mainForm, conn, explicitRows);
                }

                List<AiQuotaMatchRow> chapterRows = LoadCurrentChapterQuotaRows();
                if (chapterRows.Count > 0)
                {
                    return chapterRows;
                }

                List<AiQuotaMatchRow> fallbackRows = explicitRows.Count > 0
                    ? BuildAiQuotaRows(mainForm, conn, explicitRows)
                    : BuildAiQuotaRows(mainForm, conn, GetSelectedQuotaRows(quotaGrid));
                Log("Auto match quotas fallback: explicitRows=" + explicitRows.Count.ToString(CultureInfo.InvariantCulture) + ", fallbackRows=" + fallbackRows.Count.ToString(CultureInfo.InvariantCulture));
                return fallbackRows;
            }

            private List<DataGridViewRow> GetExplicitSelectedQuotaRows(DataGridView quotaGrid)
            {
                Dictionary<int, DataGridViewRow> rows = new Dictionary<int, DataGridViewRow>();
                if (quotaGrid == null)
                {
                    return new List<DataGridViewRow>();
                }

                foreach (DataGridViewRow row in quotaGrid.SelectedRows)
                {
                    if (row != null && !row.IsNewRow && row.Index >= 0)
                    {
                        rows[row.Index] = row;
                    }
                }

                foreach (DataGridViewCell cell in quotaGrid.SelectedCells)
                {
                    if (cell.RowIndex >= 0 && cell.RowIndex < quotaGrid.Rows.Count)
                    {
                        DataGridViewRow row = quotaGrid.Rows[cell.RowIndex];
                        if (row != null && !row.IsNewRow)
                        {
                            rows[row.Index] = row;
                        }
                    }
                }

                return rows.OrderBy(pair => pair.Key).Select(pair => pair.Value).ToList();
            }

            private List<AiQuotaMatchRow> LoadCurrentChapterQuotaRows()
            {
                List<AiQuotaMatchRow> result = new List<AiQuotaMatchRow>();
                TreeView tree = GetField<TreeView>(mainForm, "Tv_tree");
                TreeNode node = tree != null ? tree.SelectedNode : GetField<TreeNode>(mainForm, "CurrNode");
                if (node == null)
                {
                    Log("Auto match chapter: no selected tree node. treeFound=" + (tree != null ? "1" : "0"));
                    return result;
                }

                List<string> chapterSeqs = CollectAutoMatchChapterSeqs(node);
                int tagSeqCount = chapterSeqs.Count;
                string chapterNo = "";
                string totalNo = "";

                try
                {
                    EnsureOpen(conn);
                    string projectId = GetProjectId(conn);
                    totalNo = ResolveAutoMatchTotalNo(node);
                    // 树节点 Tag 里经常没有条目序号，且子节点可能懒加载未展开；
                    // 按当前条目编号从章节表补全本条目及全部子条目的序号。
                    chapterNo = ResolveAutoMatchChapterNo(node);
                    AppendAutoMatchChapterSeqsByItemNo(chapterSeqs, chapterNo);
                    if (chapterSeqs.Count == 0)
                    {
                        Log("Auto match chapter: no chapter seqs. node=" + node.Text + ", itemNo=" + (chapterNo ?? ""));
                        return result;
                    }
                    // SQL Server 单命令参数上限约 2100，条目序号分批查询后在内存里统一排序。
                    const int batchSize = 500;
                    for (int offset = 0; offset < chapterSeqs.Count; offset += batchSize)
                    {
                        int batchCount = Math.Min(batchSize, chapterSeqs.Count - offset);
                        using (SqlCommand cmd = conn.CreateCommand())
                        {
                            List<string> parameterNames = new List<string>();
                            for (int i = 0; i < batchCount; i++)
                            {
                                string parameterName = "@seq" + i.ToString(CultureInfo.InvariantCulture);
                                parameterNames.Add(parameterName);
                                cmd.Parameters.AddWithValue(parameterName, chapterSeqs[offset + i]);
                            }

                            string totalFilter = "";
                            if (!String.IsNullOrWhiteSpace(totalNo))
                            {
                                cmd.Parameters.AddWithValue("@zgs", totalNo);
                                totalFilter = " and DE.\u603b\u6982\u7b97\u5e8f\u53f7=@zgs";
                            }

                            cmd.CommandText = "select DE.\u5b9a\u989d\u5e8f\u53f7, DE.\u603b\u6982\u7b97\u5e8f\u53f7, DE.\u6761\u76ee\u5e8f\u53f7, DE.\u987a\u53f7, DE.\u5b9a\u989d\u7f16\u53f7, DE.\u5de5\u7a0b\u6216\u8d39\u7528\u9879\u76ee\u540d\u79f0, DE.\u5355\u4f4d, DE.\u5de5\u7a0b\u6570\u91cf\u8f93\u5165, ZJ.\u6761\u76ee\u7f16\u53f7 from \u5b9a\u989d\u8f93\u5165 DE left join \u7ae0\u8282\u8868 ZJ on DE.\u6761\u76ee\u5e8f\u53f7=ZJ.\u6761\u76ee\u5e8f\u53f7 where DE.\u6761\u76ee\u5e8f\u53f7 in (" + String.Join(",", parameterNames.ToArray()) + ")" + totalFilter;
                            using (SqlDataReader reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    string currentQuantity = ReadAutoMatchReaderText(reader, 7);
                                    string quotaCode = ReadAutoMatchReaderText(reader, 4);
                                    bool bindable = IsAutoMatchQuotaCode(quotaCode);
                                    long quotaSequence = 0;
                                    object sequenceValue = reader.GetValue(0);
                                    if (sequenceValue != null && sequenceValue != DBNull.Value)
                                    {
                                        quotaSequence = Convert.ToInt64(sequenceValue, CultureInfo.InvariantCulture);
                                    }

                                    if (quotaSequence <= 0)
                                    {
                                        continue;
                                    }

                                    ExcelQuotaLink link = new ExcelQuotaLink();
                                    link.ProjectId = projectId;
                                    link.QuotaSequence = quotaSequence;
                                    link.TotalNo = ReadAutoMatchReaderText(reader, 1);
                                    link.ChapterSeq = ReadAutoMatchReaderText(reader, 2);
                                    link.OrderNo = ReadAutoMatchReaderText(reader, 3);
                                    link.QuotaCode = quotaCode;
                                    link.QuotaName = ReadAutoMatchReaderText(reader, 5);

                                    result.Add(new AiQuotaMatchRow
                                    {
                                        Link = link,
                                        QuotaUnit = ReadAutoMatchReaderText(reader, 6),
                                        CurrentQuantityText = currentQuantity,
                                        Bindable = bindable,
                                        ItemNo = ReadAutoMatchReaderText(reader, 8)
                                    });
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log("Load auto match chapter quota rows failed: " + ex.Message);
                    return new List<AiQuotaMatchRow>();
                }

                Log("Auto match chapter quotas: node=" + node.Text + ", tagSeqs=" + tagSeqCount.ToString(CultureInfo.InvariantCulture) + ", itemNo=" + (chapterNo ?? "") + ", totalSeqs=" + chapterSeqs.Count.ToString(CultureInfo.InvariantCulture) + ", zgs=" + totalNo + ", rows=" + result.Count.ToString(CultureInfo.InvariantCulture));
                // 预览按条目编号由小到大（与左侧章节树一致），同条目内按顺号。
                return FilterAutoMatchPushedRows(result
                    .OrderBy(rowItem => BuildAutoMatchItemNoSortKey(rowItem.ItemNo), StringComparer.Ordinal)
                    .ThenBy(rowItem => ParseAutoMatchSortKey(rowItem.Link.ChapterSeq))
                    .ThenBy(rowItem => ParseAutoMatchSortKey(rowItem.Link.OrderNo))
                    .ThenBy(rowItem => rowItem.Link.QuotaSequence)
                    .ToList());
            }

            private string ResolveAutoMatchChapterNo(TreeNode node)
            {
                if (node == null)
                {
                    return "";
                }

                string fromTag = TryGetValue(node.Tag, "\u6761\u76ee\u7f16\u53f7");
                if (!String.IsNullOrWhiteSpace(fromTag))
                {
                    return fromTag.Trim();
                }

                string nodeName = (node.Name ?? "").Trim();
                string byName = ResolveAutoMatchChapterNoByExactCode(nodeName);
                if (!String.IsNullOrWhiteSpace(byName))
                {
                    return byName;
                }

                string seq = TryGetValue(node.Tag, "\u6761\u76ee\u5e8f\u53f7");
                if (String.IsNullOrWhiteSpace(seq) && IsNumeric(nodeName))
                {
                    seq = nodeName;
                }

                string bySeq = ResolveAutoMatchChapterNoBySeq(seq);
                if (!String.IsNullOrWhiteSpace(bySeq))
                {
                    return bySeq;
                }

                if (!String.IsNullOrWhiteSpace(nodeName) && !IsNumeric(nodeName))
                {
                    return nodeName;
                }

                string fallback = ResolveChapterNo(mainForm, conn, node);
                return fallback == null ? "" : fallback.Trim();
            }

            private string ResolveAutoMatchChapterNoByExactCode(string itemNo)
            {
                itemNo = (itemNo ?? "").Trim();
                if (String.IsNullOrWhiteSpace(itemNo))
                {
                    return "";
                }

                EnsureOpen(conn);
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "select top 1 \u6761\u76ee\u7f16\u53f7 from \u7ae0\u8282\u8868 where \u6761\u76ee\u7f16\u53f7=@no";
                    cmd.Parameters.AddWithValue("@no", itemNo);
                    object result = cmd.ExecuteScalar();
                    return result == null || result == DBNull.Value ? "" : Convert.ToString(result, CultureInfo.InvariantCulture).Trim();
                }
            }

            private string ResolveAutoMatchChapterNoBySeq(string seq)
            {
                seq = (seq ?? "").Trim();
                if (String.IsNullOrWhiteSpace(seq))
                {
                    return "";
                }

                EnsureOpen(conn);
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "select top 1 \u6761\u76ee\u7f16\u53f7 from \u7ae0\u8282\u8868 where \u6761\u76ee\u5e8f\u53f7=@id";
                    cmd.Parameters.AddWithValue("@id", seq);
                    object result = cmd.ExecuteScalar();
                    return result == null || result == DBNull.Value ? "" : Convert.ToString(result, CultureInfo.InvariantCulture).Trim();
                }
            }

            // 条目编号分段左补零做排序键（0101-4 与 0101-04-01 等长比较），空编号排最后。
            private static string BuildAutoMatchItemNoSortKey(string itemNo)
            {
                itemNo = (itemNo ?? "").Trim();
                if (itemNo.Length == 0)
                {
                    return "~";
                }

                string[] parts = itemNo.Split('-');
                StringBuilder builder = new StringBuilder(parts.Length * 13);
                foreach (string part in parts)
                {
                    builder.Append(part.Trim().PadLeft(12, '0'));
                    builder.Append('-');
                }

                return builder.ToString();
            }

            private void AppendAutoMatchChapterSeqsByItemNo(List<string> chapterSeqs, string itemNo)
            {
                itemNo = (itemNo ?? "").Trim();
                if (String.IsNullOrEmpty(itemNo))
                {
                    return;
                }

                HashSet<string> seen = new HashSet<string>(chapterSeqs, StringComparer.OrdinalIgnoreCase);
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "select \u6761\u76ee\u5e8f\u53f7, \u6761\u76ee\u7f16\u53f7 from \u7ae0\u8282\u8868 where \u6761\u76ee\u7f16\u53f7=@no or \u6761\u76ee\u7f16\u53f7 like @prefix";
                    cmd.Parameters.AddWithValue("@no", itemNo);
                    cmd.Parameters.AddWithValue("@prefix", itemNo + "%");
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            if (reader.IsDBNull(0))
                            {
                                continue;
                            }

                            string childNo = reader.IsDBNull(1) ? "" : Convert.ToString(reader.GetValue(1), CultureInfo.InvariantCulture).Trim();
                            if (!IsItemNoUnderChapter(childNo, itemNo))
                            {
                                continue;
                            }

                            string seq = Convert.ToString(reader.GetValue(0), CultureInfo.InvariantCulture).Trim();
                            if (!String.IsNullOrEmpty(seq) && seen.Add(seq))
                            {
                                chapterSeqs.Add(seq);
                            }
                        }
                    }
                }
            }

            private static long ParseAutoMatchSortKey(string text)
            {
                long value;
                return Int64.TryParse((text ?? "").Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out value) ? value : Int64.MaxValue;
            }

            private static List<string> CollectAutoMatchChapterSeqs(TreeNode root)
            {
                List<string> result = new List<string>();
                HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                CollectAutoMatchChapterSeqs(root, result, seen);
                return result;
            }

            private static void CollectAutoMatchChapterSeqs(TreeNode node, List<string> result, HashSet<string> seen)
            {
                if (node == null)
                {
                    return;
                }

                string seq = TryGetValue(node.Tag, "\u6761\u76ee\u5e8f\u53f7");
                if (String.IsNullOrWhiteSpace(seq) && IsNumeric(node.Name))
                {
                    seq = node.Name;
                }

                if (!String.IsNullOrWhiteSpace(seq) && seen.Add(seq))
                {
                    result.Add(seq);
                }

                foreach (TreeNode child in node.Nodes)
                {
                    CollectAutoMatchChapterSeqs(child, result, seen);
                }
            }

            private string ResolveAutoMatchTotalNo(TreeNode node)
            {
                for (TreeNode current = node; current != null; current = current.Parent)
                {
                    string fromTag = TryGetValue(current.Tag, "\u603b\u6982\u7b97\u5e8f\u53f7");
                    if (!String.IsNullOrWhiteSpace(fromTag))
                    {
                        return fromTag;
                    }
                }

                DataGridView quotaGrid = GetField<DataGridView>(mainForm, "dataGridViewDE");
                DataGridViewRow row = GetCurrentQuotaRow(quotaGrid);
                QuotaKey key;
                if (TryGetQuotaKey(row, out key))
                {
                    return key.TotalNo;
                }

                return "";
            }

            private static string ReadAutoMatchReaderText(SqlDataReader reader, int ordinal)
            {
                if (reader == null || reader.IsDBNull(ordinal))
                {
                    return "";
                }

                return Convert.ToString(reader.GetValue(ordinal), CultureInfo.InvariantCulture).Trim();
            }

            private HashSet<long> LoadAlreadyBoundSequences()
            {
                ExcelLinkStore store = LoadStore(conn);
                return new HashSet<long>(
                    (store.Links ?? new List<ExcelQuotaLink>())
                        .Where(link => link != null && !String.IsNullOrWhiteSpace(link.Expression))
                        .Select(link => link.QuotaSequence));
            }

            private void FillGrid()
            {
                grid.Rows.Clear();
                string scope = currentTreeScope ?? "";
                foreach (AiMatchPreviewItem item in items)
                {
                    if (!String.IsNullOrEmpty(scope) && !IsItemNoUnderChapter(item.ItemNo ?? "", scope))
                    {
                        continue;
                    }

                    int index = grid.Rows.Add(
                        item.Checked,
                        item.Link == null ? "" : item.Link.QuotaCode,
                        item.Link == null ? "" : item.Link.QuotaName,
                        item.QuotaUnit,
                        item.CurrentQuantityText,
                        item.Expression,
                        item.ExcelQuantityText,
                        item.QuantityName,
                        item.MatchStatus);
                    grid.Rows[index].Tag = item;
                    if (!item.Bindable)
                    {
                        grid.Rows[index].Cells["Checked"].ReadOnly = true;
                        grid.Rows[index].Cells["Expression"].ReadOnly = true;
                        grid.Rows[index].Cells["QuantityName"].ReadOnly = true;
                        grid.Rows[index].Cells["Checked"].ToolTipText = "\u6807\u9898\u6216\u5206\u7ec4\u884c\u4e0d\u53c2\u4e0e\u7ed1\u5b9a";
                    }
                    if (item.MatchOptions != null && item.MatchOptions.Count > 1)
                    {
                        grid.Rows[index].Cells["QuantityName"].ToolTipText = "\u70b9\u51fb\u9009\u62e9\u591a\u5904\u5339\u914d\u7684\u5de5\u7a0b\u91cf\u540d\u79f0";
                    }

                    ApplyAutoMatchRowStyle(grid.Rows[index]);
                }
            }

            private static bool IsAutoMatchWarningStatus(string matchStatus)
            {
                if (String.IsNullOrWhiteSpace(matchStatus))
                {
                    return false;
                }

                return matchStatus.IndexOf("\u6570\u91cf\u4e3a0", StringComparison.Ordinal) >= 0 ||
                    matchStatus.IndexOf("\u591a\u5904\u5339\u914d", StringComparison.Ordinal) >= 0 ||
                    matchStatus.IndexOf("\u672a\u5339\u914d", StringComparison.Ordinal) >= 0 ||
                    matchStatus.IndexOf("\u9a8c\u7b97\u4e0d\u7b26", StringComparison.Ordinal) >= 0 ||
                    matchStatus.IndexOf("\u8868\u8fbe\u5f0f\u65e0\u6cd5\u8ba1\u7b97", StringComparison.Ordinal) >= 0;
            }

            private void ApplyAutoMatchRowStyle(DataGridViewRow row)
            {
                if (row == null)
                {
                    return;
                }

                AiMatchPreviewItem item = row.Tag as AiMatchPreviewItem;
                bool highlight = item != null && item.Bindable && !item.Checked && IsAutoMatchWarningStatus(item.MatchStatus);
                row.DefaultCellStyle.BackColor = highlight ? System.Drawing.Color.MistyRose : System.Drawing.Color.Empty;
            }

            private void UpdateExpressionFromRow(DataGridViewRow row)
            {
                if (row == null)
                {
                    return;
                }

                AiMatchPreviewItem item = row.Tag as AiMatchPreviewItem;
                if (item == null || item.Link == null)
                {
                    return;
                }

                string expression = Convert.ToString(row.Cells["Expression"].Value, CultureInfo.InvariantCulture);
                expression = currentContext == null ? NormalizeExpressionOperators(expression) : currentContext.NormalizeMergedExpression(expression);
                bool expressionChanged = !String.Equals(item.Expression ?? "", expression ?? "", StringComparison.OrdinalIgnoreCase);
                item.Expression = expression;
                item.CellAddress = ExtractFirstCellAddress(expression);
                if (expressionChanged && item.MatchOptions != null && item.MatchOptions.Count > 0)
                {
                    item.MatchOptions = new List<AutoMatchCandidateOption>();
                    if (row.Cells["QuantityName"] is DataGridViewComboBoxCell)
                    {
                        DataGridViewTextBoxCell textCell = new DataGridViewTextBoxCell();
                        textCell.Value = item.QuantityName ?? "";
                        row.Cells[grid.Columns["QuantityName"].Index] = textCell;
                    }
                }

                if (String.IsNullOrWhiteSpace(expression) || String.IsNullOrWhiteSpace(item.CellAddress))
                {
                    item.Checked = false;
                    item.DisplayValue = "";
                    item.ExcelQuantityText = "";
                    item.QuantityName = "";
                    item.MatchStatus = "\u672a\u5339\u914d";
                    row.Cells["Checked"].Value = false;
                    row.Cells["Value"].Value = "";
                    row.Cells["QuantityName"].Value = "";
                    row.Cells["Status"].Value = item.MatchStatus;
                    ApplyAutoMatchRowStyle(row);
                    return;
                }

                if (currentContext != null)
                {
                    item.WorkbookPath = currentContext.WorkbookPath;
                    item.WorksheetName = currentContext.WorksheetName;
                }

                string displayValue;
                decimal quantity;
                string readError = "";
                if (currentContext == null || !TryEvaluateAutoMatchExpression(currentContext, expression, out displayValue, out quantity, out readError))
                {
                    item.Checked = false;
                    item.DisplayValue = "";
                    item.ExcelQuantityText = "";
                    item.QuantityName = "";
                    item.MatchStatus = "\u8868\u8fbe\u5f0f\u65e0\u6cd5\u8ba1\u7b97";
                    row.Cells["Checked"].Value = false;
                    row.Cells["Value"].Value = "";
                    row.Cells["QuantityName"].Value = "";
                    row.Cells["Status"].Value = item.MatchStatus;
                    status.Text = item.MatchStatus + "\uff1a" + readError;
                    ApplyAutoMatchRowStyle(row);
                    return;
                }

                item.DisplayValue = displayValue ?? "";
                item.ExcelQuantityText = BuildExcelQuantityTextForExpression(currentContext, expression);
                if (String.IsNullOrWhiteSpace(item.ExcelQuantityText))
                {
                    item.ExcelQuantityText = item.DisplayValue;
                }

                item.QuantityName = BuildQuantityNameFromExcelRow(currentContext, item.CellAddress);
                decimal quotaQuantity;
                string quotaError;
                if (TryEvaluateDecimal(item.CurrentQuantityText, out quotaQuantity, out quotaError) &&
                    RelativeDifference(quotaQuantity, quantity) > 0.03m)
                {
                    item.Checked = false;
                    item.MatchStatus = "\u9a8c\u7b97\u4e0d\u7b26";
                    row.Cells["Checked"].Value = false;
                }
                else
                {
                    item.Checked = true;
                    item.MatchStatus = "\u624b\u52a8\u4fee\u6539";
                    row.Cells["Checked"].Value = true;
                }

                row.Cells["Expression"].Value = item.Expression;
                row.Cells["Value"].Value = item.ExcelQuantityText;
                SetQuantityNameCellValue(row, item);
                row.Cells["Status"].Value = item.MatchStatus;
                ApplyAutoMatchRowStyle(row);
            }

            private void RefreshEditedExpressions()
            {
                foreach (DataGridViewRow row in grid.Rows)
                {
                    AiMatchPreviewItem item = row.Tag as AiMatchPreviewItem;
                    string expression = Convert.ToString(row.Cells["Expression"].Value, CultureInfo.InvariantCulture);
                    if (item != null && !String.Equals((item.Expression ?? ""), (expression ?? ""), StringComparison.OrdinalIgnoreCase))
                    {
                        UpdateExpressionFromRow(row);
                    }
                }
            }

            private void PrepareQuantityNameDropDown(DataGridViewRow row)
            {
                if (row == null)
                {
                    return;
                }

                AiMatchPreviewItem item = row.Tag as AiMatchPreviewItem;
                if (item == null || item.MatchOptions == null || item.MatchOptions.Count <= 1)
                {
                    return;
                }

                if (row.Cells["QuantityName"] is DataGridViewComboBoxCell)
                {
                    return;
                }

                Dictionary<string, int> nameCounts = item.MatchOptions
                    .Select(option => String.IsNullOrWhiteSpace(option.QuantityName) ? option.CellAddress ?? "" : option.QuantityName)
                    .GroupBy(name => name)
                    .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);
                DataGridViewComboBoxCell combo = new DataGridViewComboBoxCell();
                foreach (AutoMatchCandidateOption option in item.MatchOptions)
                {
                    option.Label = BuildQuantityNameOptionLabel(option, nameCounts);
                    if (!combo.Items.Contains(option.Label))
                    {
                        combo.Items.Add(option.Label);
                    }
                }

                AutoMatchCandidateOption current = FindCurrentOption(item);
                string value = current == null ? BuildQuantityNameOptionLabel(item.MatchOptions[0], nameCounts) : current.Label;
                combo.Value = value;
                row.Cells[grid.Columns["QuantityName"].Index] = combo;
            }

            private static string BuildQuantityNameOptionLabel(AutoMatchCandidateOption option, Dictionary<string, int> nameCounts)
            {
                if (option == null)
                {
                    return "";
                }

                string name = String.IsNullOrWhiteSpace(option.QuantityName) ? option.CellAddress ?? "" : option.QuantityName.Trim();
                int count = 0;
                if (nameCounts != null)
                {
                    nameCounts.TryGetValue(name, out count);
                }

                if (count > 1 && !String.IsNullOrWhiteSpace(option.CellAddress))
                {
                    return name + " (" + option.CellAddress + ")";
                }

                return name;
            }

            private static AutoMatchCandidateOption FindCurrentOption(AiMatchPreviewItem item)
            {
                if (item == null || item.MatchOptions == null)
                {
                    return null;
                }

                foreach (AutoMatchCandidateOption option in item.MatchOptions)
                {
                    if (String.Equals(option.Expression ?? "", item.Expression ?? "", StringComparison.OrdinalIgnoreCase))
                    {
                        return option;
                    }
                }

                return item.MatchOptions.Count > 0 ? item.MatchOptions[0] : null;
            }

            private void SetQuantityNameCellValue(DataGridViewRow row, AiMatchPreviewItem item)
            {
                if (row == null || item == null)
                {
                    return;
                }

                object value = item.QuantityName ?? "";
                DataGridViewComboBoxCell combo = row.Cells["QuantityName"] as DataGridViewComboBoxCell;
                if (combo != null)
                {
                    AutoMatchCandidateOption option = FindCurrentOption(item);
                    if (option != null)
                    {
                        if (String.IsNullOrWhiteSpace(option.Label))
                        {
                            option.Label = String.IsNullOrWhiteSpace(option.QuantityName) ? option.CellAddress ?? "" : option.QuantityName;
                        }

                        value = option.Label;
                        if (!combo.Items.Contains(value))
                        {
                            combo.Items.Add(value);
                        }
                    }
                }

                row.Cells["QuantityName"].Value = value;
            }

            private void ApplyBatchCheck(int rowIndex)
            {
                if (applyingBatchCheck || rowIndex < 0 || rowIndex >= grid.Rows.Count)
                {
                    return;
                }

                object rawValue = grid.Rows[rowIndex].Cells["Checked"].Value;
                bool value = rawValue is bool && (bool)rawValue;
                AiMatchPreviewItem currentItem = grid.Rows[rowIndex].Tag as AiMatchPreviewItem;
                if (currentItem != null)
                {
                    currentItem.Checked = value;
                    ApplyAutoMatchRowStyle(grid.Rows[rowIndex]);
                }

                if (pendingBatchCheckRows == null)
                {
                    return;
                }

                List<int> targets = pendingBatchCheckRows;
                pendingBatchCheckRows = null;
                if (!targets.Contains(rowIndex))
                {
                    return;
                }

                applyingBatchCheck = true;
                try
                {
                    int applied = 0;
                    foreach (int index in targets)
                    {
                        if (index == rowIndex || index < 0 || index >= grid.Rows.Count)
                        {
                            continue;
                        }

                        DataGridViewRow row = grid.Rows[index];
                        AiMatchPreviewItem item = row.Tag as AiMatchPreviewItem;
                        if (item == null || !item.Bindable)
                        {
                            continue;
                        }

                        item.Checked = value;
                        row.Cells["Checked"].Value = value;
                        ApplyAutoMatchRowStyle(row);
                        applied++;
                    }

                    if (applied > 0)
                    {
                        status.Text = "已批量" + (value ? "勾选" : "取消勾选") + " " + (applied + 1).ToString(CultureInfo.InvariantCulture) + " 行。";
                    }
                }
                finally
                {
                    applyingBatchCheck = false;
                }
            }

            private void ApplyQuantityNameOption(DataGridViewRow row, string selectedLabel)
            {
                if (updatingQuantityNameCell || row == null || String.IsNullOrWhiteSpace(selectedLabel))
                {
                    return;
                }

                AiMatchPreviewItem item = row.Tag as AiMatchPreviewItem;
                if (item == null || item.MatchOptions == null || item.MatchOptions.Count <= 1)
                {
                    return;
                }

                AutoMatchCandidateOption selected = item.MatchOptions
                    .FirstOrDefault(option => String.Equals(option.Label ?? "", selectedLabel, StringComparison.OrdinalIgnoreCase))
                    ?? item.MatchOptions.FirstOrDefault(option => String.Equals(option.QuantityName ?? "", selectedLabel, StringComparison.OrdinalIgnoreCase));
                if (selected == null || String.Equals(selected.Expression ?? "", item.Expression ?? "", StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                updatingQuantityNameCell = true;
                try
                {
                    item.Checked = true;
                    item.Expression = selected.Expression;
                    item.CellAddress = selected.CellAddress;
                    item.DisplayValue = selected.DisplayValue ?? "";
                    item.ExcelQuantityText = selected.ExcelQuantityText ?? "";
                    item.QuantityName = selected.QuantityName ?? "";
                    item.MatchStatus = "\u624b\u52a8\u9009\u62e9";

                    row.Cells["Checked"].Value = true;
                    row.Cells["Expression"].Value = item.Expression;
                    row.Cells["Value"].Value = item.ExcelQuantityText;
                    row.Cells["QuantityName"].Value = selected.Label ?? selected.QuantityName ?? "";
                    row.Cells["Status"].Value = item.MatchStatus;
                    ApplyAutoMatchRowStyle(row);
                    status.Text = "\u5df2\u9009\u62e9\u591a\u5904\u5339\u914d\u5019\u9009\uff1a" + (item.Link == null ? "" : item.Link.QuotaCode) + " -> " + item.Expression;
                }
                finally
                {
                    updatingQuantityNameCell = false;
                }
            }

            private void ToggleManualMatch()
            {
                if (!manualMatchButton.Checked)
                {
                    manualMatchTimer.Stop();
                    if (!suppressManualClosedStatus)
                    {
                        status.Text = "\u624b\u52a8\u5339\u914d\u5df2\u5173\u95ed\u3002";
                    }
                    return;
                }

                if (matchingInProgress)
                {
                    status.Text = "\u81ea\u52a8\u5339\u914d\u8fdb\u884c\u4e2d\uff0c\u8bf7\u7a0d\u5019\u518d\u5f00\u542f\u624b\u52a8\u5339\u914d\u3002";
                    suppressManualClosedStatus = true;
                    manualMatchButton.Checked = false;
                    suppressManualClosedStatus = false;
                    return;
                }

                List<int> targetColumns;
                bool reused;
                string error;
                status.Text = "\u6b63\u5728\u51c6\u5907Excel\u5feb\u7167...";
                Application.DoEvents();
                if (!EnsureCurrentAutoMatchSnapshot(out targetColumns, out reused, out error))
                {
                    status.Text = error;
                    suppressManualClosedStatus = true;
                    manualMatchButton.Checked = false;
                    suppressManualClosedStatus = false;
                    return;
                }

                if (items.Count == 0)
                {
                    List<AiQuotaMatchRow> quotas = LoadCurrentSelectedQuotas();
                    if (quotas.Count == 0)
                    {
                        status.Text = "\u8bf7\u5148\u5728\u5b9a\u989d\u8f93\u5165\u8868\u4e2d\u6846\u9009\u5b9a\u989d\uff0c\u6216\u5728\u5de6\u4fa7\u70b9\u9009\u8981\u5339\u914d\u7684\u7ae0\u8282\u6761\u76ee\u3002";
                        suppressManualClosedStatus = true;
                        manualMatchButton.Checked = false;
                        suppressManualClosedStatus = false;
                        return;
                    }

                    items = quotas
                        .Select(quota => BuildAutoMatchPreviewItem(quota, currentContext, null, "\u672a\u5339\u914d"))
                        .ToList();
                    RebuildItemTree();
                    FillGrid();
                }

                lastManualCellKey = "";
                manualBaselinePending = true;
                manualMatchTimer.Start();
                status.Text = "\u624b\u52a8\u5339\u914d\u5df2\u5f00\u542f\uff1a\u9009\u4e2d\u9884\u89c8\u8868\u5b9a\u989d\u884c\uff0c\u518d\u70b9\u51fbExcel\u5355\u5143\u683c\uff0c\u5c06\u4f7f\u7528\u5f53\u524d\u5feb\u7167\u751f\u6210\u5730\u5740\u8868\u8fbe\u5f0f\u3002";
            }

            private void PollManualMatchCell()
            {
                if (!manualMatchButton.Checked || matchingInProgress || currentContext == null || grid.IsCurrentCellInEditMode)
                {
                    return;
                }

                DataGridViewRow row = grid.CurrentRow;
                if (row == null)
                {
                    return;
                }

                string workbookPath;
                string worksheetName;
                string rawAddress;
                if (!TryGetActiveExcelCellLite(out workbookPath, out worksheetName, out rawAddress))
                {
                    return;
                }

                string address = currentContext.NormalizeMergedCellAddress(rawAddress);
                string key = (workbookPath ?? "") + "|" + (worksheetName ?? "") + "|" + address + "|" + row.Index.ToString(CultureInfo.InvariantCulture);
                if (String.Equals(lastManualCellKey, key, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }

                lastManualCellKey = key;
                if (manualBaselinePending)
                {
                    // 切换目标行或刚开启手动匹配时，先把当前活动格记为基线，
                    // 等用户再点击其它Excel格才执行绑定，避免把上一次的活动格误绑到新行。
                    manualBaselinePending = false;
                    status.Text = "\u5df2\u9009\u4e2d\u5b9a\u989d\u884c\uff0c\u8bf7\u5728Excel\u4e2d\u70b9\u51fb\u8981\u7ed1\u5b9a\u7684\u5355\u5143\u683c\u3002";
                    return;
                }

                if (!String.Equals(workbookPath ?? "", currentContext.WorkbookPath ?? "", StringComparison.OrdinalIgnoreCase) ||
                    !String.Equals(worksheetName ?? "", currentContext.WorksheetName ?? "", StringComparison.OrdinalIgnoreCase))
                {
                    status.Text = "\u5f53\u524dExcel\u5355\u5143\u683c\u4e0d\u5728\u5df2\u8bfb\u53d6\u5feb\u7167\u7684\u5de5\u4f5c\u8868\u4e2d\u3002";
                    return;
                }

                AiExcelCell snapshotCell;
                if (!currentContext.CellByAddress.TryGetValue(address, out snapshotCell))
                {
                    status.Text = "\u5f53\u524dExcel\u683c " + address + " \u4e0d\u5728\u5df2\u8bfb\u53d6\u5feb\u7167\u8303\u56f4\u5185\uff0c\u8bf7\u8c03\u6574\u76ee\u6807\u5217\u540e\u91cd\u65b0\u5f00\u59cb\u5339\u914d\u3002";
                    return;
                }

                ApplySnapshotCellToSelectedRow(row, snapshotCell);
            }

            private static bool TryGetActiveExcelCellLite(out string workbookPath, out string worksheetName, out string cellAddress)
            {
                workbookPath = null;
                worksheetName = null;
                cellAddress = null;
                dynamic excel = null;
                try
                {
                    excel = GetActiveSpreadsheetApplication();
                    if (excel == null)
                    {
                        return false;
                    }

                    dynamic workbook = excel.ActiveWorkbook;
                    dynamic sheet = excel.ActiveSheet;
                    dynamic selection = excel.Selection;
                    if (workbook == null || sheet == null || selection == null)
                    {
                        return false;
                    }

                    dynamic firstCell = selection.Cells[1, 1];
                    workbookPath = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture);
                    worksheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture);
                    cellAddress = Convert.ToString(firstCell.Address(false, false), CultureInfo.InvariantCulture);
                    return !String.IsNullOrEmpty(workbookPath) && !String.IsNullOrEmpty(worksheetName) && !String.IsNullOrEmpty(cellAddress);
                }
                catch (COMException)
                {
                    ClearCachedSpreadsheetApplication((object)excel);
                    return false;
                }
                catch (Exception)
                {
                    ClearCachedSpreadsheetApplication((object)excel);
                    return false;
                }
            }

            private void ApplySnapshotCellToSelectedRow(DataGridViewRow row, AiExcelCell snapshotCell)
            {
                AiMatchPreviewItem item = row == null ? null : row.Tag as AiMatchPreviewItem;
                if (item == null || item.Link == null || snapshotCell == null)
                {
                    status.Text = "\u5f53\u524d\u884c\u65e0\u6cd5\u624b\u52a8\u5339\u914d\u3002";
                    return;
                }
                if (!item.Bindable)
                {
                    status.Text = "\u6807\u9898\u6216\u5206\u7ec4\u884c\u4e0d\u53c2\u4e0e\u7ed1\u5b9a\u3002";
                    return;
                }

                string unitText = currentContext.GetUnitNearCell(snapshotCell, item.QuotaUnit);
                string suffix;
                if (!TryBuildManualSnapshotSuffix(unitText, item.QuotaUnit, out suffix))
                {
                    status.Text = BuildSimpleBindingUnitMismatchMessage(unitText, item.QuotaUnit);
                    return;
                }

                string expression = snapshotCell.Address + suffix;
                string displayValue;
                decimal quantity;
                string readError;
                if (!TryEvaluateAutoMatchExpression(currentContext, expression, out displayValue, out quantity, out readError))
                {
                    status.Text = "\u5feb\u7167\u4e2d\u7684Excel\u683c\u65e0\u6cd5\u8ba1\u7b97\uff1a" + readError;
                    return;
                }

                item.Checked = true;
                item.WorkbookPath = currentContext.WorkbookPath;
                item.WorksheetName = currentContext.WorksheetName;
                item.Expression = expression;
                item.CellAddress = snapshotCell.Address;
                item.DisplayValue = displayValue ?? "";
                item.ExcelQuantityText = String.IsNullOrWhiteSpace(snapshotCell.Text) ? "0" : snapshotCell.Text;
                item.QuantityName = BuildQuantityNameFromExcelRow(currentContext, snapshotCell.Address);
                item.MatchStatus = "\u624b\u52a8\u5339\u914d";
                item.MatchOptions = new List<AutoMatchCandidateOption>();
                if (row.Cells["QuantityName"] is DataGridViewComboBoxCell)
                {
                    DataGridViewTextBoxCell textCell = new DataGridViewTextBoxCell();
                    textCell.Value = item.QuantityName ?? "";
                    row.Cells[grid.Columns["QuantityName"].Index] = textCell;
                }

                row.Cells["Checked"].Value = true;
                row.Cells["Expression"].Value = item.Expression;
                row.Cells["Value"].Value = item.ExcelQuantityText;
                SetQuantityNameCellValue(row, item);
                row.Cells["Status"].Value = item.MatchStatus;
                ApplyAutoMatchRowStyle(row);
                status.Text = "\u5df2\u624b\u52a8\u5339\u914d\uff1a" + (item.Link.QuotaCode ?? "") + " -> " + item.WorksheetName + "!" + item.Expression;
            }

            private static bool TryBuildManualSnapshotSuffix(string excelUnit, string quotaUnit, out string suffix)
            {
                suffix = "";
                if (String.IsNullOrWhiteSpace(excelUnit) || String.IsNullOrWhiteSpace(quotaUnit))
                {
                    return false;
                }

                // 两侧单位同名（如 亩=亩）时按 1:1 绑定，不要求进单位白名单。
                if (String.Equals(NormalizeExcelLinkUnit(excelUnit), NormalizeExcelLinkUnit(quotaUnit), StringComparison.Ordinal))
                {
                    return true;
                }

                if (!LooksLikeExcelLinkUnit(excelUnit) || !LooksLikeExcelLinkUnit(quotaUnit))
                {
                    return false;
                }

                return TryBuildExcelLinkUnitScaleSuffix(excelUnit, quotaUnit, out suffix);
            }

            private void AcceptCurrentItem()
            {
                grid.EndEdit();
                DataGridViewRow row = grid.CurrentRow;
                if (row == null)
                {
                    status.Text = "\u8bf7\u5148\u9009\u4e2d\u8981\u7ed1\u5b9a\u7684\u4e00\u884c\u3002";
                    return;
                }

                UpdateExpressionFromRow(row);
                AiMatchPreviewItem item = row.Tag as AiMatchPreviewItem;
                if (item != null && !item.Bindable)
                {
                    status.Text = "\u6807\u9898\u6216\u5206\u7ec4\u884c\u4e0d\u53c2\u4e0e\u7ed1\u5b9a\u3002";
                    return;
                }
                if (item == null || item.Link == null || String.IsNullOrWhiteSpace(item.Expression) || String.IsNullOrWhiteSpace(item.CellAddress))
                {
                    status.Text = "\u5f53\u524d\u884c\u8fd8\u6ca1\u6709\u53ef\u7ed1\u5b9a\u7684\u5339\u914d\u8868\u8fbe\u5f0f\u3002";
                    return;
                }

                item.Checked = true;
                row.Cells["Checked"].Value = true;
                ApplyAutoMatchRowStyle(row);
                if (Accepted != null)
                {
                    Accepted(new List<AiMatchPreviewItem> { item });
                }

                status.Text = "\u5df2\u5355\u4e2a\u7ed1\u5b9a\uff1a" + (item.Link.QuotaCode ?? "") + " -> " + item.Expression;
            }

            private void AcceptCheckedItems()
            {
                grid.EndEdit();
                RefreshEditedExpressions();
                List<AiMatchPreviewItem> accepted = GetAcceptedItems();
                if (accepted.Count == 0)
                {
                    status.Text = "\u8bf7\u81f3\u5c11\u52fe\u9009\u4e00\u6761\u5df2\u5339\u914dExcel\u5355\u5143\u683c\u7684\u5b9a\u989d\u3002";
                    return;
                }

                if (Accepted != null)
                {
                    Accepted(accepted);
                }

                DialogResult = DialogResult.OK;
                Close();
            }

            public List<AiMatchPreviewItem> GetAcceptedItems()
            {
                // 从 items 收集而不是表格行：树过滤后未显示的已勾选行也要参与全部绑定。
                grid.EndEdit();
                FlushGridCheckedToItems();
                List<AiMatchPreviewItem> accepted = new List<AiMatchPreviewItem>();
                foreach (AiMatchPreviewItem item in items)
                {
                    if (item != null && item.Checked && item.Bindable && !String.IsNullOrWhiteSpace(item.Expression) && !String.IsNullOrWhiteSpace(item.CellAddress))
                    {
                        accepted.Add(item);
                    }
                }

                return accepted;
            }

            // 把当前可见表格行的勾选状态回写到 items（单个勾选不会经过 ApplyBatchCheck 同步）。
            private void FlushGridCheckedToItems()
            {
                foreach (DataGridViewRow row in grid.Rows)
                {
                    AiMatchPreviewItem item = row.Tag as AiMatchPreviewItem;
                    if (item == null)
                    {
                        continue;
                    }

                    object value = row.Cells["Checked"].Value;
                    item.Checked = value is bool && (bool)value;
                    ApplyAutoMatchRowStyle(row);
                }
            }

            private void RebuildItemTree()
            {
                rebuildingTree = true;
                itemTree.BeginUpdate();
                try
                {
                    itemTree.Nodes.Clear();
                    currentTreeScope = "";
                    chapterNames = LoadChapterNameMap(mainForm);

                    TreeNode root = new TreeNode("全部条目");
                    root.Tag = "";
                    itemTree.Nodes.Add(root);

                    Dictionary<string, TreeNode> nodesByCode = new Dictionary<string, TreeNode>(StringComparer.OrdinalIgnoreCase);
                    foreach (string itemNo in items
                        .Select(it => (it.ItemNo ?? "").Trim())
                        .Where(no => no.Length > 0)
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .OrderBy(no => no, StringComparer.OrdinalIgnoreCase))
                    {
                        TreeNode parent = root;
                        foreach (string code in BuildChapterChain(chapterNames, itemNo))
                        {
                            TreeNode node;
                            if (!nodesByCode.TryGetValue(code, out node))
                            {
                                node = new TreeNode(ChapterTreeDisplayName(chapterNames, code));
                                node.Tag = code;
                                node.ToolTipText = code;
                                parent.Nodes.Add(node);
                                nodesByCode[code] = node;
                            }

                            parent = node;
                        }
                    }

                    root.Expand();
                    itemTree.SelectedNode = root;
                }
                finally
                {
                    itemTree.EndUpdate();
                    rebuildingTree = false;
                }
            }

            private void OnItemTreeScopeChanged()
            {
                if (rebuildingTree || matchingInProgress)
                {
                    return;
                }

                grid.EndEdit();
                FlushGridCheckedToItems();
                TreeNode node = itemTree.SelectedNode;
                currentTreeScope = node == null ? "" : Convert.ToString(node.Tag);
                FillGrid();
            }

            // 勾选树节点＝整枝勾选/取消该条目（含下级）下已有匹配表达式的定额。
            private void OnItemTreeChecked(TreeNode node)
            {
                if (rebuildingTree || updatingTreeChecks || matchingInProgress || node == null)
                {
                    return;
                }

                updatingTreeChecks = true;
                try
                {
                    bool value = node.Checked;
                    SetTreeChildrenChecked(node, value);
                    grid.EndEdit();
                    FlushGridCheckedToItems();
                    string scope = Convert.ToString(node.Tag);
                    int applied = 0;
                    foreach (AiMatchPreviewItem item in items)
                    {
                        if (item == null || !item.Bindable)
                        {
                            continue;
                        }

                        if (!String.IsNullOrEmpty(scope) && !IsItemNoUnderChapter(item.ItemNo ?? "", scope))
                        {
                            continue;
                        }

                        bool newValue = value && !String.IsNullOrWhiteSpace(item.Expression);
                        if (item.Checked != newValue)
                        {
                            applied++;
                        }

                        item.Checked = newValue;
                    }

                    FillGrid();
                    status.Text = "已按条目树批量" + (value ? "勾选" : "取消勾选") + " " + applied.ToString(CultureInfo.InvariantCulture) + " 行。";
                }
                finally
                {
                    updatingTreeChecks = false;
                }
            }
        }
    }
}
