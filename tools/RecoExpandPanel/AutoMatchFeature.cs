using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
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
            public decimal Quantity;
            public string QuantityName;
            public string Status;
            public bool Checked;
            public int FirstRow;
            public int FirstColumn;
            public decimal Score;
            public readonly List<string> Addresses = new List<string>();
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

                if (alreadyBoundSequences != null && alreadyBoundSequences.Contains(quota.Link.QuotaSequence))
                {
                    preview.Add(BuildAutoMatchPreviewItem(quota, context, null, "\u5df2\u7ed1\u5b9a"));
                    continue;
                }

                decimal quotaQuantity;
                string quotaError;
                if (!TryEvaluateDecimal(quota.CurrentQuantityText, out quotaQuantity, out quotaError))
                {
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
            item.QuantityName = candidate == null ? "" : candidate.QuantityName;
            item.CurrentQuantityText = quota.CurrentQuantityText;
            item.MatchStatus = status;
            return item;
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

                if (String.IsNullOrWhiteSpace(item.MatchStatus))
                {
                    item.MatchStatus = String.IsNullOrWhiteSpace(item.Expression) ? "\u672a\u5339\u914d" : "\u5df2\u5339\u914d";
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
                if (cell == null || !cell.IsNumber || (targetSet != null && !targetSet.Contains(cell.Column)))
                {
                    continue;
                }

                decimal value;
                string error;
                if (!TryEvaluateDecimal(cell.Text, out value, out error))
                {
                    continue;
                }

                byAddress[cell.Address] = new AutoMatchCellValue { Cell = cell, Value = value };
            }

            return byAddress.Values.OrderBy(cell => cell.Cell.Row).ThenBy(cell => cell.Cell.Column).ToList();
        }

        private static bool TryBuildOperandAutoMatch(AiQuotaMatchRow quota, decimal quotaQuantity, AiExcelSelectionContext context, AutoMatchNumberIndex numberIndex, HashSet<string> usedAddresses, int previousRow, int previousColumn, out AutoMatchExpressionCandidate candidate)
        {
            candidate = null;
            List<QuantityExpressionTerm> terms;
            if (!TryParseQuantityExpressionTerms(quota.CurrentQuantityText, out terms))
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
                string displayValue;
                decimal quantity;
                string error;
                if (!TryEvaluateAutoMatchExpression(context, expression, out displayValue, out quantity, out error))
                {
                    continue;
                }

                decimal diff = RelativeDifference(quotaQuantity, quantity);
                AutoMatchExpressionCandidate current = BuildAutoMatchCandidate(quota, context, terms, cells, expression, displayValue, quantity, diff, combinationCount, usedAddresses, previousRow, previousColumn);
                if (candidate == null || current.Score > candidate.Score)
                {
                    candidate = current;
                }
            }

            return candidate != null;
        }

        private static bool TryBuildWholeValueAutoMatch(AiQuotaMatchRow quota, decimal quotaQuantity, AiExcelSelectionContext context, List<AutoMatchCellValue> numberCells, HashSet<string> usedAddresses, int previousRow, int previousColumn, out AutoMatchExpressionCandidate candidate)
        {
            candidate = null;
            int matchCount = 0;
            foreach (AutoMatchCellValue cell in numberCells ?? new List<AutoMatchCellValue>())
            {
                string expression;
                decimal score;
                if (!TryBuildLocalQuantityExpression(quota.CurrentQuantityText, quotaQuantity, cell.Cell, out expression, out score))
                {
                    continue;
                }

                string displayValue;
                decimal quantity;
                string error;
                if (!TryEvaluateAutoMatchExpression(context, expression, out displayValue, out quantity, out error))
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
            displayValue = null;
            quantity = 0m;
            error = null;
            Dictionary<string, AiExcelCell> cells = context == null
                ? new Dictionary<string, AiExcelCell>(StringComparer.OrdinalIgnoreCase)
                : context.Cells.GroupBy(cell => cell.Address, StringComparer.OrdinalIgnoreCase).ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

            string resolved = NormalizeExpressionOperators(expression);
            bool allResolved = true;
            foreach (string address in ExtractCellAddressesFromExpression(resolved).OrderByDescending(value => value.Length))
            {
                AiExcelCell cell;
                decimal cellValue;
                string parseError;
                if (!cells.TryGetValue(address, out cell) || !TryEvaluateDecimal(cell.Text, out cellValue, out parseError))
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

            if (context != null && TryEvaluateWorkbookExpression(context.WorkbookPath, context.WorksheetName, expression, out displayValue, out quantity, out error))
            {
                displayValue = FormatAiMatchDecimal(quantity);
                return true;
            }

            return false;
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
                int readColCount = readLastColumn - firstColumn + 1;
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

                if (!ok || context.Cells.Count == 0)
                {
                    error = "\u76ee\u6807\u5de5\u4f5c\u8868\u6ca1\u6709\u53ef\u8bfb\u53d6\u7684\u5185\u5bb9\u3002";
                    return false;
                }

                if (!HasAutoMatchNumberCell(context, targetColumns))
                {
                    error = "\u6307\u5b9a\u8303\u56f4\u5185\u6ca1\u6709\u53ef\u8ba1\u7b97\u7684\u6570\u91cf\u5355\u5143\u683c\u3002";
                    return false;
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

        private static bool HasAutoMatchNumberCell(AiExcelSelectionContext context, List<int> targetColumns)
        {
            HashSet<int> targetSet = targetColumns != null && targetColumns.Count > 0
                ? new HashSet<int>(targetColumns)
                : null;
            return context.Cells.Any(cell => cell.IsNumber && (targetSet == null || targetSet.Contains(cell.Column)));
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
            private readonly List<AiQuotaMatchRow> quotas;
            private readonly HashSet<long> alreadyBoundSequences;
            private readonly ComboBox sheetBox;
            private readonly TextBox targetColumnText;
            private readonly DataGridView grid;
            private readonly Label status;
            private List<AiMatchPreviewItem> items = new List<AiMatchPreviewItem>();

            public event Action<List<AiMatchPreviewItem>> Accepted;
            public event Action Cancelled;

            public AutoMatchDialog(List<AiQuotaMatchRow> quotaRows, HashSet<long> alreadyBound)
            {
                quotas = quotaRows ?? new List<AiQuotaMatchRow>();
                alreadyBoundSequences = alreadyBound ?? new HashSet<long>();
                Text = "\u81ea\u52a8\u5339\u914dExcel\u5de5\u7a0b\u91cf";
                StartPosition = FormStartPosition.CenterParent;
                Size = new System.Drawing.Size(980, 560);
                MinimumSize = new System.Drawing.Size(860, 420);
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

                Button start = new Button();
                start.Text = "\u5f00\u59cb\u5339\u914d";
                start.Left = 492;
                start.Top = 8;
                start.Width = 90;
                start.Click += delegate { StartMatch(); };

                top.Controls.Add(sheetLabel);
                top.Controls.Add(sheetBox);
                top.Controls.Add(columnLabel);
                top.Controls.Add(targetColumnText);
                top.Controls.Add(start);

                grid = new DataGridView();
                grid.Dock = DockStyle.Fill;
                grid.AllowUserToAddRows = false;
                grid.AllowUserToDeleteRows = false;
                grid.RowHeadersVisible = false;
                grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                grid.CellDoubleClick += delegate { ApplyCurrentExcelCellToSelectedRow(); };
                BuildGridColumns();

                Button applyCurrent = new Button();
                applyCurrent.Text = "\u7528\u5f53\u524dExcel\u683c\u5339\u914d";
                applyCurrent.Width = 130;
                applyCurrent.Click += delegate { ApplyCurrentExcelCellToSelectedRow(); };

                Button ok = new Button();
                ok.Text = "\u786e\u8ba4\u7ed1\u5b9a";
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
                buttons.Controls.Add(applyCurrent);

                status = new Label();
                status.Dock = DockStyle.Bottom;
                status.Height = 26;
                status.Padding = new Padding(8, 2, 8, 2);

                Controls.Add(grid);
                Controls.Add(buttons);
                Controls.Add(status);
                Controls.Add(top);

                LoadSheets();
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
                grid.Columns.Add("Value", "Excel\u53d6\u503c");
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
                grid.EndEdit();
                List<int> targetColumns;
                string error;
                if (!TryParseTargetColumns(targetColumnText.Text, out targetColumns, out error))
                {
                    status.Text = error;
                    return;
                }

                if (sheetBox.SelectedItem == null)
                {
                    status.Text = "\u8bf7\u5148\u9009\u62e9\u5de5\u4f5c\u8868\u3002";
                    return;
                }

                status.Text = "\u6b63\u5728\u8bfb\u53d6Excel\u5e76\u672c\u5730\u5339\u914d...";
                UseWaitCursor = true;
                Application.DoEvents();
                try
                {
                    AiExcelSelectionContext context;
                    if (!TryReadWorksheetCellsForAutoMatch(Convert.ToString(sheetBox.SelectedItem, CultureInfo.InvariantCulture), targetColumns, out context, out error))
                    {
                        status.Text = error;
                        return;
                    }

                    items = BuildAutoMatchPreviewItems(quotas, context, targetColumns, alreadyBoundSequences);
                    FillGrid();
                    int checkedCount = items.Count(item => item.Checked);
                    status.Text = "\u5339\u914d\u5b8c\u6210\uff1a\u5171 " + items.Count.ToString(CultureInfo.InvariantCulture) + " \u6761\uff0c\u9ed8\u8ba4\u52fe\u9009 " + checkedCount.ToString(CultureInfo.InvariantCulture) + " \u6761\u3002";
                }
                finally
                {
                    UseWaitCursor = false;
                }
            }

            private void FillGrid()
            {
                grid.Rows.Clear();
                foreach (AiMatchPreviewItem item in items)
                {
                    int index = grid.Rows.Add(
                        item.Checked,
                        item.Link == null ? "" : item.Link.QuotaCode,
                        item.Link == null ? "" : item.Link.QuotaName,
                        item.QuotaUnit,
                        item.CurrentQuantityText,
                        item.Expression,
                        item.DisplayValue,
                        item.QuantityName,
                        item.MatchStatus);
                    grid.Rows[index].Tag = item;
                }
            }

            private void ApplyCurrentExcelCellToSelectedRow()
            {
                grid.EndEdit();
                DataGridViewRow row = grid.CurrentRow;
                if (row == null)
                {
                    status.Text = "\u8bf7\u5148\u9009\u4e2d\u9884\u89c8\u8868\u4e2d\u7684\u4e00\u6761\u5b9a\u989d\u3002";
                    return;
                }

                AiMatchPreviewItem item = row.Tag as AiMatchPreviewItem;
                if (item == null || item.Link == null)
                {
                    status.Text = "\u5f53\u524d\u884c\u65e0\u6cd5\u5339\u914d\u3002";
                    return;
                }

                ExcelCellAddress cell;
                string error;
                if (!TryGetActiveExcelCell(out cell, out error, true))
                {
                    status.Text = "\u8bf7\u5148\u5728WPS/Excel\u91cc\u70b9\u9009\u5de5\u7a0b\u6570\u91cf\u5355\u5143\u683c\u3002";
                    return;
                }

                string expression = BuildDefaultExpression(cell);
                string displayValue;
                decimal quantity;
                string readError;
                if (!TryEvaluateWorkbookExpression(cell.WorkbookPath, cell.WorksheetName, expression, out displayValue, out quantity, out readError))
                {
                    status.Text = "\u5f53\u524dExcel\u683c\u65e0\u6cd5\u8ba1\u7b97\uff1a" + readError;
                    return;
                }

                item.Checked = true;
                item.WorkbookPath = cell.WorkbookPath;
                item.WorksheetName = cell.WorksheetName;
                item.Expression = expression;
                item.CellAddress = ExtractFirstCellAddress(expression);
                item.DisplayValue = displayValue ?? "";
                item.QuantityName = BuildQuantityNameNearActiveExcelCell(cell);
                item.MatchStatus = "\u624b\u52a8\u5339\u914d";

                row.Cells["Checked"].Value = true;
                row.Cells["Expression"].Value = item.Expression;
                row.Cells["Value"].Value = item.DisplayValue;
                row.Cells["QuantityName"].Value = item.QuantityName;
                row.Cells["Status"].Value = item.MatchStatus;
                status.Text = "\u5df2\u5339\u914d\uff1a" + (item.Link.QuotaCode ?? "") + " -> " + System.IO.Path.GetFileName(item.WorkbookPath) + "!" + item.WorksheetName + "!" + item.Expression;
            }

            private void AcceptCheckedItems()
            {
                grid.EndEdit();
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
                List<AiMatchPreviewItem> accepted = new List<AiMatchPreviewItem>();
                foreach (DataGridViewRow row in grid.Rows)
                {
                    bool isChecked = row.Cells["Checked"].Value is bool && (bool)row.Cells["Checked"].Value;
                    AiMatchPreviewItem item = row.Tag as AiMatchPreviewItem;
                    if (isChecked && item != null && !String.IsNullOrWhiteSpace(item.Expression) && !String.IsNullOrWhiteSpace(item.CellAddress))
                    {
                        accepted.Add(item);
                    }
                }

                return accepted;
            }
        }
    }
}
