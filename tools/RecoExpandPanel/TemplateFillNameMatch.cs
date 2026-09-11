using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using NPOI.SS.UserModel;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        // ===== 模板铺量·名字驱动匹配 =====
        // 设计：docs/superpowers/specs/2026-07-06-模板铺量-名字驱动-design.md
        // 与序列对齐引擎(铺量plus)同源思路，一期在本仓库自建这一小块；若 铺量plus 已有可移植替换。

        private const int NameMatchMinScore = 55;   // 名字匹配成立的最低分

        // 归一化：全角转半角、去空白与标点、小写；文字内嵌数字保留（500m 的 500 是身份）。
        internal static string NormalizeMatchText(string text)
        {
            if (String.IsNullOrEmpty(text)) return "";
            string source = NormalizeForSignature(text);
            StringBuilder sb = new StringBuilder(source.Length);
            for (int i = 0; i < source.Length; i++)
            {
                char c = source[i];
                if (Char.IsWhiteSpace(c) || !Char.IsLetterOrDigit(c)) continue;
                sb.Append(Char.ToLowerInvariant(c));
            }
            return sb.ToString();
        }

        private static string NormalizeQuantityMatchName(string text)
        {
            string name = (text ?? "").Trim();
            // 尾部“（单位）”/“(单位)”：括号内经 LooksLikeExcelLinkUnit 判定为单位才剥，
            // 否则 C30混凝土(m3) 的 3 会与对方相交、击穿数字规格惩罚。
            if (name.Length > 0 && (name[name.Length - 1] == ')' || name[name.Length - 1] == '）'))
            {
                int open = name.LastIndexOfAny(new[] { '(', '（' });
                if (open > 0)
                {
                    string inner = name.Substring(open + 1, name.Length - open - 2).Trim();
                    if (LooksLikeExcelLinkUnit(inner))
                    {
                        name = name.Substring(0, open).Trim();
                    }
                }
            }
            int separator = name.LastIndexOf(' ');
            if (separator > 0)
            {
                string suffix = name.Substring(separator + 1).Trim();
                if (LooksLikeExcelLinkUnit(suffix))
                {
                    name = name.Substring(0, separator).Trim();
                }
            }
            return NormalizeMatchText(name);
        }

        // 紧贴单位字母的数字不算规格数字：m2/m3 的幂次、串尾 100m/10km 这类“数量+单位”。
        // 500m长轨 的 500 仍是身份（单位字母后还有正文，不在串尾）。
        internal static List<string> ExtractMatchNumbers(string normText)
        {
            List<string> result = new List<string>();
            if (String.IsNullOrEmpty(normText)) return result;
            int i = 0;
            while (i < normText.Length)
            {
                if (Char.IsDigit(normText[i]))
                {
                    int start = i;
                    while (i < normText.Length && (Char.IsDigit(normText[i]) || normText[i] == '.')) i++;
                    string number = normText.Substring(start, i - start).TrimEnd('.');
                    if (IsMatchUnitExponentDigit(normText, start, number)) continue;
                    int unitEnd;
                    if (IsMatchTrailingUnitNumber(normText, start, i, out unitEnd)) { i = unitEnd; continue; }
                    result.Add(number);
                }
                else i++;
            }
            return result;
        }

        private static bool IsMatchAsciiLetter(char c)
        {
            return (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z');
        }

        // m2/m3/km2/cm3 等：数字是长度单位后的幂次。
        private static bool IsMatchUnitExponentDigit(string text, int digitStart, string number)
        {
            if (number != "2" && number != "3") return false;
            int letterStart = digitStart;
            while (letterStart > 0 && IsMatchAsciiLetter(text[letterStart - 1])) letterStart--;
            if (letterStart == digitStart) return false;
            string letters = text.Substring(letterStart, digitStart - letterStart).ToLowerInvariant();
            return letters == "m" || letters == "km" || letters == "hm" || letters == "dm" || letters == "cm" || letters == "mm";
        }

        // 串尾“数字+单位字母(+幂次)”，如 100m、10km、100m2：整体经 LooksLikeExcelLinkUnit 判定为单位才跳过。
        private static bool IsMatchTrailingUnitNumber(string text, int digitStart, int digitEnd, out int unitEnd)
        {
            unitEnd = digitEnd;
            int j = digitEnd;
            while (j < text.Length && IsMatchAsciiLetter(text[j])) j++;
            if (j == digitEnd) return false;
            if (j < text.Length && (text[j] == '2' || text[j] == '3')) j++;
            if (j != text.Length) return false;
            if (!LooksLikeExcelLinkUnit(text.Substring(digitStart, j - digitStart))) return false;
            unitEnd = j;
            return true;
        }

        private static HashSet<string> BuildMatchBigrams(string text)
        {
            HashSet<string> grams = new HashSet<string>(StringComparer.Ordinal);
            if (text.Length == 1) { grams.Add(text); return grams; }
            for (int i = 0; i + 1 < text.Length; i++) grams.Add(text.Substring(i, 2));
            return grams;
        }

        private sealed class MatchTextFeatures
        {
            public string Norm;
            public HashSet<string> Bigrams;
            public List<string> Numbers;
        }

        private sealed class TemplateNameGroup
        {
            public string NormName;
            public string Chapter;
            public string SourceAnchor;
            public MatchTextFeatures Features;
            public List<int> Indexes = new List<int>();
        }

        private static MatchTextFeatures BuildMatchTextFeatures(string normalizedText)
        {
            string norm = normalizedText ?? "";
            return new MatchTextFeatures
            {
                Norm = norm,
                Bigrams = BuildMatchBigrams(norm),
                Numbers = ExtractMatchNumbers(norm)
            };
        }

        // 相似度 0-100：字符 bigram Dice；双方都含数字且数字集不相交时重罚(/3)。
        internal static int MatchNameScore(string leftNorm, string rightNorm)
        {
            return MatchNameScore(BuildMatchTextFeatures(leftNorm), BuildMatchTextFeatures(rightNorm));
        }

        private static int MatchNameScore(MatchTextFeatures left, MatchTextFeatures right)
        {
            if (left == null || right == null || String.IsNullOrEmpty(left.Norm) || String.IsNullOrEmpty(right.Norm)) return 0;
            if (String.Equals(left.Norm, right.Norm, StringComparison.Ordinal)) return 100;
            if (left.Bigrams.Count == 0 || right.Bigrams.Count == 0) return 0;
            int common = left.Bigrams.Count(g => right.Bigrams.Contains(g));
            int score = (int)Math.Round(200.0 * common / (left.Bigrams.Count + right.Bigrams.Count));
            if (left.Numbers.Count > 0 && right.Numbers.Count > 0 && !left.Numbers.Any(n => right.Numbers.Contains(n))) score /= 3;
            return score > 100 ? 100 : score;
        }

        // 章节缺失或不一致时不得自动认领。兼容同义章节标题，但保持保守阈值。
        internal static bool AreMatchChaptersCompatible(string left, string right)
        {
            string l = NormalizeMatchText(left);
            string r = NormalizeMatchText(right);
            if (l.Length == 0 || r.Length == 0) return false;
            int leftOrdinal, rightOrdinal;
            string leftTitle, rightTitle;
            bool leftHasOrdinal = TryGetChapterOrdinal(left, out leftOrdinal, out leftTitle);
            bool rightHasOrdinal = TryGetChapterOrdinal(right, out rightOrdinal, out rightTitle);
            if (leftHasOrdinal && rightHasOrdinal)
            {
                if (leftOrdinal != rightOrdinal) return false;
                // 序号相等时剥掉序号前缀只比标题：“第1章 路基工程”与“第一章 路基工程”应兼容。
                string leftTitleNorm = NormalizeMatchText(leftTitle);
                string rightTitleNorm = NormalizeMatchText(rightTitle);
                if (leftTitleNorm.Length > 0 && rightTitleNorm.Length > 0)
                {
                    l = leftTitleNorm;
                    r = rightTitleNorm;
                }
            }
            if (String.Equals(l, r, StringComparison.Ordinal)) return true;
            if (Math.Min(l.Length, r.Length) >= 3 && (l.Contains(r) || r.Contains(l))) return true;
            return MatchNameScore(l, r) >= 70;
        }

        // 章节序号统一规则（IsChapterAnchorRaw 同用）：第?1章/节/部分、1、、第?一章/节/部分、一、、(一)。
        // 无“第”前缀时，汉字/阿拉伯序号 + 章/节/部分 之后必须是空白、“、”或串尾，
        // 所以“四节段预制梁”“二分之一预制块”不是章节；“1.1 土方开挖”“12 混凝土”“(1)”也不是。
        // title 为剥掉序号前缀后的标题。
        private static bool TryGetChapterOrdinal(string text, out int ordinal, out string title)
        {
            ordinal = 0;
            title = "";
            if (String.IsNullOrWhiteSpace(text)) return false;
            string trimmed = text.Trim();
            System.Text.RegularExpressions.Match match = System.Text.RegularExpressions.Regex.Match(trimmed,
                @"^(?:(?<di>第)\s*)?(?<num>[0-9]+|[一二三四五六七八九十百]+)\s*(?:(?<kind>章|节|部分)|、)|^[\(（]\s*(?<num>[一二三四五六七八九十百]+)\s*[\)）]");
            if (!match.Success) return false;
            if (match.Groups["kind"].Success && !match.Groups["di"].Success)
            {
                int end = match.Index + match.Length;
                if (end < trimmed.Length && !Char.IsWhiteSpace(trimmed[end]) && trimmed[end] != '、') return false;
            }
            string raw = match.Groups["num"].Value;
            if (!Int32.TryParse(raw, NumberStyles.Integer, CultureInfo.InvariantCulture, out ordinal))
            {
                ordinal = ParseChineseChapterNumber(raw);
            }
            if (ordinal <= 0) { ordinal = 0; return false; }
            title = trimmed.Substring(match.Index + match.Length).Trim().TrimStart('、').Trim();
            return true;
        }

        private static int ParseChineseChapterNumber(string raw)
        {
            if (String.IsNullOrEmpty(raw)) return 0;
            const string digits = "零一二三四五六七八九";
            int total = 0;
            int current = 0;
            foreach (char c in raw)
            {
                int digit = digits.IndexOf(c);
                if (digit >= 0) { current = digit; continue; }
                if (c == '十') { total += (current == 0 ? 1 : current) * 10; current = 0; continue; }
                if (c == '百') { total += (current == 0 ? 1 : current) * 100; current = 0; continue; }
                return 0;
            }
            return total + current;
        }

        private static bool SameTemplateChapter(string left, string right)
        {
            return String.Equals(NormalizeMatchText(left), NormalizeMatchText(right), StringComparison.Ordinal);
        }

        private static string BuildTemplateSourceAnchor(FillTemplate template, FillTemplateRow row)
        {
            string first = ExtractFirstCellAddress(row == null ? "" : row.SourceExpr);
            if (!String.IsNullOrEmpty(first))
            {
                return NormalizeTemplateWorkbookPath(GetTemplateRowWorkbookPath(template, row)) + "|" +
                    ((row.SourceSheet ?? "").Trim().ToUpperInvariant()) + "|" + first.ToUpperInvariant();
            }
            return "legacy|" + NormalizeMatchText(row == null ? "" : row.MatchChapter);
        }

        private static List<TemplateNameGroup> BuildTemplateNameGroups(FillTemplate template)
        {
            List<TemplateNameGroup> result = new List<TemplateNameGroup>();
            if (template == null || template.Rows == null) return result;
            Dictionary<string, TemplateNameGroup> byKey = new Dictionary<string, TemplateNameGroup>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < template.Rows.Count; i++)
            {
                FillTemplateRow row = template.Rows[i];
                string norm = NormalizeQuantityMatchName(row == null ? "" : (row.MatchName ?? row.SourceName ?? ""));
                if (norm.Length == 0) continue;
                string anchor = BuildTemplateSourceAnchor(template, row);
                string key = norm + "\u001f" + anchor;
                TemplateNameGroup group;
                if (!byKey.TryGetValue(key, out group))
                {
                    group = new TemplateNameGroup
                    {
                        NormName = norm,
                        Chapter = row.MatchChapter ?? "",
                        SourceAnchor = anchor,
                        Features = BuildMatchTextFeatures(norm)
                    };
                    byKey[key] = group;
                    result.Add(group);
                }
                group.Indexes.Add(i);
                if (String.IsNullOrWhiteSpace(group.Chapter) && !String.IsNullOrWhiteSpace(row.MatchChapter))
                {
                    group.Chapter = row.MatchChapter;
                }
            }
            return result;
        }

        internal static string GetExactNameConflict(int targetNameCount, int templateGroupCount)
        {
            if (targetNameCount > 1) return "target";
            if (templateGroupCount > 1) return "template";
            return "";
        }

        internal static string GetExactNameResolutionMode(int targetNameCount, int templateGroupCount)
        {
            if (templateGroupCount > 1) return "choice";
            if (templateGroupCount == 1 && targetNameCount > 1) return "reuse";
            if (templateGroupCount == 1) return "single";
            return "";
        }

        internal static int FindUniqueBestMatchIndex(string queryNorm, string queryChapter,
            IList<string> candidateNorms, IList<string> candidateChapters, out bool ambiguous)
        {
            List<MatchTextFeatures> features = candidateNorms.Select(BuildMatchTextFeatures).ToList();
            return FindUniqueBestMatchIndexCached(BuildMatchTextFeatures(queryNorm), queryChapter,
                features, candidateChapters, out ambiguous);
        }

        private static int FindUniqueBestMatchIndexCached(MatchTextFeatures queryFeatures, string queryChapter,
            IList<MatchTextFeatures> candidateFeatures, IList<string> candidateChapters, out bool ambiguous)
        {
            ambiguous = false;
            int best = -1;
            int bestScore = NameMatchMinScore - 1;
            int bestChapterRank = -1;
            for (int i = 0; i < candidateFeatures.Count; i++)
            {
                int score = MatchNameScore(queryFeatures, candidateFeatures[i]);
                if (score < NameMatchMinScore) continue;
                int chapterRank = AreMatchChaptersCompatible(queryChapter, candidateChapters[i]) ? 1 : 0;
                if (score > bestScore || (score == bestScore && chapterRank > bestChapterRank))
                {
                    best = i;
                    bestScore = score;
                    bestChapterRank = chapterRank;
                    ambiguous = false;
                }
                else if (score == bestScore && chapterRank == bestChapterRank)
                {
                    ambiguous = true;
                }
            }
            return ambiguous ? -1 : best;
        }

        // 从候选名字列表里挑与 query 最匹配的下标；低于阈值返回 -1。sameChapterOnly 交由调用方先过滤候选。
        internal static int BestMatchIndex(string queryNorm, IList<string> candidateNorms)
        {
            int best = -1, bestScore = NameMatchMinScore - 1;
            for (int i = 0; i < candidateNorms.Count; i++)
            {
                int s = MatchNameScore(queryNorm, candidateNorms[i]);
                if (s > bestScore) { bestScore = s; best = i; }
            }
            return best;
        }

        private static ExcelSyncReadContext CreateNameFillReadContext(
            FillTemplate template,
            Dictionary<string, HashSet<int>> hiddenCache,
            Dictionary<string, List<ExcelMergedRegion>> mergedCache)
        {
            List<ExcelQuotaLink> readLinks = new List<ExcelQuotaLink>();
            if (template == null || template.Rows == null) return new ExcelSyncReadContext(readLinks);

            foreach (FillTemplateRow row in template.Rows)
            {
                if (row == null) continue;
                string workbook = GetTemplateRowWorkbookPath(template, row);
                if (String.IsNullOrWhiteSpace(workbook) || String.IsNullOrWhiteSpace(row.SourceSheet) ||
                    String.IsNullOrWhiteSpace(row.SourceExpr)) continue;

                List<string> cells = ExtractCellAddressesFromExpression(row.SourceExpr);
                if (cells.Count <= 1)
                {
                    AddQuantityNameReadLinks(readLinks, workbook, row.SourceSheet, row.SourceExpr, hiddenCache, mergedCache);
                }
                else
                {
                    foreach (string cell in cells)
                    {
                        AddQuantityNameReadLinks(readLinks, workbook, row.SourceSheet, cell, hiddenCache, mergedCache);
                    }
                }
            }
            return new ExcelSyncReadContext(readLinks);
        }

        // 名字模式模版生成：与 BuildFillTemplateFromBindings 同源，额外为每行读 Excel 工程量主名称；
        // 表达式(E1+E2)拆操作数各读主名称存 Operands，套用时按名字定位、不再绑坐标。
        private static FillTemplate BuildNameFillTemplateFromBindings(
            Form mainForm, SqlConnection conn, string templateName, string unitNo, string sourceSheet)
        {
            FillTemplate template = BuildFillTemplateFromBindings(mainForm, conn, templateName, unitNo, sourceSheet);
            template.MatchBy = "name";

            Dictionary<string, HashSet<int>> hiddenCache = new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, List<ExcelMergedRegion>> mergedCache = new Dictionary<string, List<ExcelMergedRegion>>(StringComparer.OrdinalIgnoreCase);
            ExcelSyncReadContext readContext = CreateNameFillReadContext(template, hiddenCache, mergedCache);

            foreach (FillTemplateRow row in template.Rows)
            {
                string workbook = GetTemplateRowWorkbookPath(template, row);
                if (String.IsNullOrWhiteSpace(workbook) || String.IsNullOrWhiteSpace(row.SourceSheet) || String.IsNullOrWhiteSpace(row.SourceExpr))
                {
                    continue;
                }

                List<string> cells = ExtractCellAddressesFromExpression(row.SourceExpr);
                if (cells.Count <= 1)
                {
                    row.MatchName = ReadFullNameForCell(workbook, row.SourceSheet, row.SourceExpr,
                        hiddenCache, mergedCache, readContext);
                }
                else
                {
                    row.Operands = new List<FillOperand>();
                    foreach (string cell in cells)
                    {
                        FillOperand op = new FillOperand();
                        op.Op = "+";
                        op.Name = ReadFullNameForCell(workbook, row.SourceSheet, cell,
                            hiddenCache, mergedCache, readContext);
                        row.Operands.Add(op);
                    }
                    row.MatchName = row.Operands.Count > 0 ? row.Operands[0].Name : "";
                }
            }
            PopulateTemplateMatchChapters(template);
            return template;
        }

        // 名字模板记录源 Excel 章节。旧模板预览时也调用本方法临时补读，但不静默保存模板。
        private static void PopulateTemplateMatchChapters(FillTemplate template)
        {
            if (template == null || template.Rows == null) return;
            Dictionary<string, Dictionary<int, string>> cache = new Dictionary<string, Dictionary<int, string>>(StringComparer.OrdinalIgnoreCase);
            foreach (FillTemplateRow row in template.Rows)
            {
                if (row == null || !String.IsNullOrWhiteSpace(row.MatchChapter)) continue;
                string workbook = GetTemplateRowWorkbookPath(template, row);
                string first = ExtractFirstCellAddress(row.SourceExpr);
                CellRef cr;
                if (String.IsNullOrWhiteSpace(workbook) || String.IsNullOrWhiteSpace(row.SourceSheet) ||
                    String.IsNullOrEmpty(first) || !TryParseCellAddress(first, out cr))
                {
                    continue;
                }

                string fullWorkbook;
                try { fullWorkbook = Path.GetFullPath(workbook); }
                catch { continue; }
                if (!File.Exists(fullWorkbook)) continue;
                string key = fullWorkbook + "|" + row.SourceSheet + "|" + cr.Column.ToString(CultureInfo.InvariantCulture);
                Dictionary<int, string> chapters;
                if (!cache.TryGetValue(key, out chapters))
                {
                    try
                    {
                        Dictionary<int, string> chapterSnapshot;
                        ReadTargetQtyRowsWithChapters(fullWorkbook, row.SourceSheet, cr.Column, out chapterSnapshot);
                        chapters = chapterSnapshot;
                    }
                    catch (Exception ex)
                    {
                        Log("Populate template chapters failed: " + ex.Message);
                        chapters = new Dictionary<int, string>();
                    }
                    cache[key] = chapters;
                }
                string chapter;
                if (chapters.TryGetValue(cr.Row, out chapter)) row.MatchChapter = chapter;
            }
        }

        // 读某表达式首格所在行的工程量主名称；模板侧与目标侧共用同一拆字段口径。
        private static string ReadFullNameForCell(string workbook, string sheet, string expr,
            Dictionary<string, HashSet<int>> hiddenCache, Dictionary<string, List<ExcelMergedRegion>> mergedCache,
            ExcelSyncReadContext readContext)
        {
            RowNameParts parts;
            return TryReadRowNamePartsAt(workbook, sheet, expr, hiddenCache, mergedCache, readContext, out parts)
                ? parts.MainName
                : "";
        }

        private sealed class TargetQtyRow
        {
            public int Row;
            public string RawName;    // 数量列左侧主名称(不截断，供匹配)
            public string DisplayName;// 数量列左侧截断显示名(3段/15字)，供 UI 显示
            public string NormName;   // 归一化
            public string Chapter;    // 预留：二期章节内就近约束
            public string ContextLabel; // 本行拆出的章节/段落上文标签，只作审计
            public string Unit;       // 数量列左邻格的单位文本(供单位换算)，读不到为空
            public decimal Quantity;
            public string QuantityText;
        }

        // 读目标 sheet：数量列(qtyColumn) 有数字的行=工程量行；主名称由左侧结构化片段拆出；
        // 章节锚点行用于给每个工程量行标 Chapter(取其上方最近锚点)。
        private static List<TargetQtyRow> ReadTargetQtyRows(string workbook, string sheet, int qtyColumn)
        {
            Dictionary<int, string> ignored;
            return ReadTargetQtyRowsWithChapters(workbook, sheet, qtyColumn, out ignored);
        }

        private static string ReadTargetUnitNearQuantity(string workbook, string sheet, int row, int qtyColumn,
            Dictionary<string, HashSet<int>> hiddenCache, ExcelSyncReadContext ctx,
            List<ExcelMergedRegion> mergedRegions)
        {
            if (ctx == null || row <= 0 || qtyColumn <= 1) return "";

            HashSet<int> hiddenColumns = GetSavedHiddenColumns(workbook, sheet, hiddenCache);
            int visibleChecked = 0;
            for (int col = qtyColumn - 1; col >= 1 && visibleChecked < 6; col--)
            {
                if (hiddenColumns.Contains(col)) continue;
                visibleChecked++;

                string sourceKey;
                int fragmentColumn;
                bool isVerticalMerge;
                string text = ReadRowNameCellText(workbook, sheet, col, row, ctx, mergedRegions,
                    out sourceKey, out fragmentColumn, out isVerticalMerge);
                if (LooksLikeExcelLinkUnit(text)) return (text ?? "").Trim();
            }
            return "";
        }

        private static List<TargetQtyRow> ReadTargetQtyRowsWithChapters(string workbook, string sheet, int qtyColumn,
            out Dictionary<int, string> chapterByRow)
        {
            string ignoredError;
            return ReadTargetQtyRowsWithChaptersDetailed(workbook, sheet, qtyColumn, out chapterByRow, out ignoredError);
        }

        // readError：读不到任何工程量行时回传首个读取失败原因（行范围失败或首个单元格取值错误），
        // 供预览 warning 拼接；读到行时为空。（不与上面的同名方法重载：回归脚本按名字反射取唯一方法。）
        private static List<TargetQtyRow> ReadTargetQtyRowsWithChaptersDetailed(string workbook, string sheet, int qtyColumn,
            out Dictionary<int, string> chapterByRow, out string readError)
        {
            List<TargetQtyRow> result = new List<TargetQtyRow>();
            chapterByRow = new Dictionary<int, string>();
            readError = "";
            Dictionary<string, HashSet<int>> hiddenCache = new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, List<ExcelMergedRegion>> mergedCache = new Dictionary<string, List<ExcelMergedRegion>>(StringComparer.OrdinalIgnoreCase);

            List<ExcelQuotaLink> readLinks = new List<ExcelQuotaLink>();
            int firstRow, lastRow;
            string rangeError;
            if (!TryGetSheetRowRange(workbook, sheet, out firstRow, out lastRow, out rangeError))
            {
                readError = String.IsNullOrWhiteSpace(rangeError) ? "无法读取工作表行范围" : rangeError;
                return result;
            }
            string firstCellError = "";
            string qtyColName = ColumnNumberToName(qtyColumn);
            for (int r = firstRow; r <= lastRow; r++)
            {
                string qtyAddr = qtyColName + r.ToString(CultureInfo.InvariantCulture);
                readLinks.Add(new ExcelQuotaLink { ExcelPath = workbook, WorksheetName = sheet, CellAddress = qtyAddr, Expression = qtyAddr });
                AddQuantityNameReadLinks(readLinks, workbook, sheet, qtyAddr, hiddenCache, mergedCache);
            }
            ExcelSyncReadContext ctx = new ExcelSyncReadContext(readLinks);

            string currentChapter = "";
            for (int r = firstRow; r <= lastRow; r++)
            {
                string qtyAddr = qtyColName + r.ToString(CultureInfo.InvariantCulture);
                string disp; decimal qty; string err;
                bool evaluated = TryEvaluateWorkbookExpression(ctx, workbook, sheet, qtyAddr, out disp, out qty, out err, true);
                if (!evaluated && firstCellError.Length == 0 && !String.IsNullOrWhiteSpace(err))
                    firstCellError = qtyAddr + " " + err.Trim();
                bool hasQty = evaluated && qty != 0m;
                if (!hasQty)
                {
                    string name = ReadRowNameAt(workbook, sheet, qtyAddr, hiddenCache, mergedCache, ctx, true);
                    if (IsChapterAnchorRaw(name)) { currentChapter = name; continue; }
                    chapterByRow[r] = currentChapter;
                    continue;
                }

                RowNameParts parts;
                string displayName;
                if (!TryReadRowNamePartsAt(workbook, sheet, qtyAddr, hiddenCache, mergedCache, ctx,
                    out parts, out displayName))
                {
                    chapterByRow[r] = currentChapter;
                    continue;
                }
                string rowChapter = String.IsNullOrWhiteSpace(parts.ContextLabel) ? currentChapter : parts.ContextLabel;
                chapterByRow[r] = rowChapter;
                if (String.IsNullOrWhiteSpace(parts.MainName)) continue;
                displayName = StripTrailingQuantityUnit(displayName, parts.Unit);
                TargetQtyRow row = new TargetQtyRow();
                row.Row = r;
                row.RawName = parts.MainName;
                row.DisplayName = displayName;
                row.NormName = NormalizeQuantityMatchName(parts.MainName);
                row.Chapter = rowChapter;
                row.ContextLabel = parts.ContextLabel;
                row.Quantity = qty;
                row.QuantityText = disp;
                row.Unit = parts.Unit;
                result.Add(row);
            }
            if (result.Count == 0 && firstCellError.Length > 0)
                readError = "行 " + firstRow.ToString(CultureInfo.InvariantCulture) + "-" +
                    lastRow.ToString(CultureInfo.InvariantCulture) + " 内首个取值错误 " + firstCellError;
            return result;
        }

        // 名字驱动数量文本：只在 Excel 单位与定额单位可靠同量纲时生成换算后缀。
        // Excel 单位缺失或跨基础单位时不猜测，由上层已确认公式或人工确认处理。
        private static string BuildNameDrivenQtyText(string baseQtyText, string excelUnit, string quotaUnit)
        {
            string suffix;
            if (!String.IsNullOrWhiteSpace(excelUnit))
            {
                return TryBuildExcelLinkUnitScaleSuffix(excelUnit, quotaUnit, out suffix)
                    ? (baseQtyText ?? "") + suffix
                    : baseQtyText;
            }
            return baseQtyText;
        }

        // “一、/(一)/第X章/第X部分/1、/第1章” 视为章节锚点；规则与 TryGetChapterOrdinal 合一。
        private static bool IsChapterAnchorRaw(string raw)
        {
            int ordinal;
            string title;
            return TryGetChapterOrdinal(raw, out ordinal, out title);
        }

        // 取 sheet 的 UsedRange 行范围（首末非空行号）；失败时 error 说明原因（供预览 warning 展示）。
        private static bool TryGetSheetRowRange(string workbook, string sheet, out int firstRow, out int lastRow,
            out string error)
        {
            firstRow = 0; lastRow = 0; error = "";
            try
            {
                if (!TryGetXlsxUsedRowRange(workbook, sheet, out firstRow, out lastRow, out error))
                {
                    if (String.IsNullOrWhiteSpace(error)) error = "无法读取工作表行范围";
                    return false;
                }
                if (lastRow < firstRow)
                {
                    error = "工作表没有已使用的行";
                    return false;
                }
                if (lastRow - firstRow >= 20000)
                {
                    error = "工作表已使用行范围过大（第 " + firstRow.ToString(CultureInfo.InvariantCulture) + "-" +
                        lastRow.ToString(CultureInfo.InvariantCulture) + " 行，超过 20000 行上限），请先清理表尾空行";
                    return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
        }

        // 取该 sheet 的首末已用行号(1基)。xlsx/xlsm 走流式读 sheet XML：扫描 sheetData 内 <row r="N"> 取首末
        // （与原 NPOI FirstRowNum/LastRowNum 同口径），sheetData 没有任何 row 时才退回 <dimension ref="A1:H500">
        // ——NPOI 等库生成的文件 dimension 可能停留在默认 "A1"，不能优先信任；只有 .xls 回退 NPOI 整簿加载。
        // 打不开工作簿或找不到该 sheet 返回 false 并回传 error。
        private static bool TryGetXlsxUsedRowRange(string workbook, string sheet, out int firstRow, out int lastRow,
            out string error)
        {
            firstRow = 0; lastRow = 0; error = "";
            if (String.IsNullOrWhiteSpace(workbook) || String.IsNullOrWhiteSpace(sheet))
            {
                error = "未指定目标 Excel 或工作表";
                return false;
            }
            if (!File.Exists(workbook))
            {
                error = "目标 Excel 不存在：" + workbook;
                return false;
            }

            try
            {
                if (String.Equals(Path.GetExtension(workbook), ".xls", StringComparison.OrdinalIgnoreCase))
                    return TryGetUsedRowRangeByNpoi(workbook, sheet, out firstRow, out lastRow, out error);
                return TryGetXlsxUsedRowRangeStreaming(workbook, sheet, out firstRow, out lastRow, out error);
            }
            catch (Exception ex)
            {
                Log("TryGetXlsxUsedRowRange 失败: " + ex.Message);
                error = "读取工作表行范围失败：" + ex.Message;
                return false;
            }
        }

        private static bool TryGetXlsxUsedRowRangeStreaming(string workbook, string sheet, out int firstRow, out int lastRow,
            out string error)
        {
            firstRow = 0; lastRow = 0; error = "";
            using (System.IO.Compression.ZipArchive archive = OpenZipArchiveShared(workbook))
            {
                string sheetPath = ResolveSheetPath(archive, sheet);
                if (String.IsNullOrEmpty(sheetPath))
                {
                    error = "工作簿中找不到工作表「" + sheet + "」";
                    return false;
                }
                System.IO.Compression.ZipArchiveEntry sheetEntry = archive.GetEntry(sheetPath);
                if (sheetEntry == null)
                {
                    error = "工作表数据缺失：" + sheetPath;
                    return false;
                }

                int scanFirst = 0, scanLast = 0;
                int dimensionFirst = 0, dimensionLast = 0;
                bool hasDimension = false;
                using (Stream stream = sheetEntry.Open())
                using (System.Xml.XmlReader reader = System.Xml.XmlReader.Create(stream))
                {
                    while (reader.Read())
                    {
                        if (reader.NodeType == System.Xml.XmlNodeType.EndElement && reader.LocalName == "sheetData") break;
                        if (reader.NodeType != System.Xml.XmlNodeType.Element) continue;
                        if (reader.LocalName == "dimension")
                        {
                            hasDimension = TryParseDimensionRowRange(reader.GetAttribute("ref"), out dimensionFirst, out dimensionLast);
                            continue;
                        }
                        if (reader.LocalName != "row") continue;
                        int rowNumber;
                        if (!Int32.TryParse(reader.GetAttribute("r"), NumberStyles.Integer, CultureInfo.InvariantCulture,
                            out rowNumber) || rowNumber <= 0) continue;
                        if (scanFirst == 0 || rowNumber < scanFirst) scanFirst = rowNumber;
                        if (rowNumber > scanLast) scanLast = rowNumber;
                    }
                }
                if (scanLast > 0)
                {
                    firstRow = scanFirst;
                    lastRow = scanLast;
                    return true;
                }
                if (hasDimension)
                {
                    firstRow = dimensionFirst;
                    lastRow = dimensionLast;
                    return true;
                }
                error = "工作表「" + sheet + "」没有已使用的行";
                return false;
            }
        }

        // <dimension ref="A1:H500"> 或单格 "A1"；解析失败返回 false（调用方继续扫描 row）。
        private static bool TryParseDimensionRowRange(string refText, out int firstRow, out int lastRow)
        {
            firstRow = 0; lastRow = 0;
            string[] parts = (refText ?? "").Split(':');
            if (parts.Length < 1 || parts.Length > 2) return false;
            CellRef start, end;
            if (!TryParseCellAddress(parts[0], out start)) return false;
            if (parts.Length == 1) end = start;
            else if (!TryParseCellAddress(parts[1], out end)) return false;
            firstRow = Math.Min(start.Row, end.Row);
            lastRow = Math.Max(start.Row, end.Row);
            return true;
        }

        // .xls 没有可流式读取的 sheet XML，只能用 NPOI 打开工作簿(只读共享流)取首末行号(1基)。
        private static bool TryGetUsedRowRangeByNpoi(string workbook, string sheet, out int firstRow, out int lastRow,
            out string error)
        {
            firstRow = 0; lastRow = 0; error = "";
            using (Stream stream = OpenWorkbookStreamShared(workbook))
            {
                IWorkbook wb = WorkbookFactory.Create(stream);
                ISheet sh = wb.GetSheet(sheet);
                if (sh == null)
                {
                    error = "工作簿中找不到工作表「" + sheet + "」";
                    return false;
                }

                int first0 = sh.FirstRowNum;
                int last0 = sh.LastRowNum;
                if (last0 < first0)
                {
                    error = "工作表「" + sheet + "」没有已使用的行";
                    return false;
                }

                firstRow = first0 + 1;
                lastRow = last0 + 1;
                return true;
            }
        }

        // 模板/对应框组内顺序：真定额、数字材料、费用或辅助类伪代码。
        // 数字材料允许带 *系数 后缀；五位以下纯数字不在这里判为材料。
        private static int TemplateTargetRank(string code)
        {
            string c = (code ?? "").Trim().ToUpperInvariant();
            int star = c.IndexOf('*');
            if (star > 0) c = c.Substring(0, star);
            switch (c)
            {
                case "SF": case "SH": case "SQ": case "ZLF": case "LF":
                case "YF": case "TLF": case "GF": case "JF": case "XGT1":
                    return 2;
            }
            if (c.Length >= 5 && c.All(Char.IsDigit)) return 1;
            return 0;
        }

        // 多操作数表达式(如 E4+E5)：把每个操作数按其工程量名在目标表定位,以其数量代入原表达式。
        // 全部命中才返回 true；exprText=代入后的数字表达式(软件工程数量输入格式)；
        // operandTargetIdx=各操作数命中的目标行下标(供主循环标注“已并入”)。
        private static bool TrySubstituteOperandQuantities(FillTemplateRow trow, List<TargetQtyRow> targetRows,
            List<string> targetNorms, out string exprText, out List<int> operandTargetIdx)
        {
            exprText = null;
            operandTargetIdx = new List<int>();
            if (trow == null || trow.Operands == null || trow.Operands.Count < 2 || String.IsNullOrWhiteSpace(trow.SourceExpr))
            {
                return false;
            }

            List<string> cells = ExtractCellAddressesFromExpression(trow.SourceExpr);
            if (cells.Count != trow.Operands.Count)
            {
                return false;
            }

            List<string> values = new List<string>();
            foreach (FillOperand op in trow.Operands)
            {
                int idx = BestMatchIndex(NormalizeQuantityMatchName(op.Name), targetNorms);
                if (idx < 0)
                {
                    return false;
                }

                operandTargetIdx.Add(idx);
                values.Add("(" + targetRows[idx].QuantityText + ")");
            }

            int next = 0;
            string result = System.Text.RegularExpressions.Regex.Replace(
                trow.SourceExpr.ToUpperInvariant(), "\\$?[A-Z]{1,3}\\$?\\d+",
                delegate(System.Text.RegularExpressions.Match m)
                {
                    string v = next < values.Count ? values[next] : m.Value;
                    next++;
                    return v;
                });

            decimal parsed;
            string err;
            if (!TryEvaluateDecimal(result, out parsed, out err))
            {
                return false;
            }

            exprText = result;
            return true;
        }

        private static List<int> OrderedTemplateGroupIndexes(FillTemplate template, TemplateNameGroup group)
        {
            List<int> indexes = group == null ? new List<int>() : group.Indexes.ToList();
            indexes.Sort(delegate(int a, int b)
            {
                int ra = TemplateTargetRank(template.Rows[a].QuotaCode);
                int rb = TemplateTargetRank(template.Rows[b].QuotaCode);
                if (ra != rb) return ra.CompareTo(rb);
                int order = template.Rows[a].OrderInItem.CompareTo(template.Rows[b].OrderInItem);
                if (order != 0) return order;
                return a.CompareTo(b);
            });
            return indexes;
        }

        private static string BuildTemplateCandidateLabel(FillTemplate template, TemplateNameGroup group)
        {
            List<FillTemplateRow> rows = OrderedTemplateGroupIndexes(template, group)
                .Select(index => template.Rows[index])
                .ToList();
            if (rows.Count == 0) return "";
            string[] names = rows
                .Select(row => String.IsNullOrWhiteSpace(row.SourceName) ? (row.QuotaCode ?? "") : row.SourceName)
                .ToArray();
            if (rows.Count == 1)
            {
                return names[0];
            }
            return String.Join(" + ", names) +
                "（组件" + rows.Count.ToString(CultureInfo.InvariantCulture) + "条）";
        }

        private static string BuildNameQuotaCandidateSignature(NameQuotaCandidateGroup candidate)
        {
            if (candidate == null || candidate.Items == null) return "";
            return String.Join("\u001e", candidate.Items
                .Where(item => item != null)
                .OrderBy(item => item.GroupOrder)
                .Select(item => String.Join("\u001f", new[]
                {
                    (item.QuotaCode ?? "").Trim().ToUpperInvariant(),
                    (item.Unit ?? "").Trim().ToUpperInvariant(),
                    (item.Adjust ?? "").Trim(),
                    (item.QuantityText ?? "").Trim()
                }))
                .ToArray());
        }

        private static List<FillPreviewItem> BuildTemplatePreviewGroup(FillTemplate template, TemplateNameGroup group,
            TargetQtyRow target, List<TargetQtyRow> targetRows, List<string> targetNorms, int targetIndex,
            Dictionary<int, string> mergedIntoByTargetIdx)
        {
            List<FillPreviewItem> result = new List<FillPreviewItem>();
            int groupOrder = 0;
            foreach (int groupIndex in OrderedTemplateGroupIndexes(template, group))
            {
                FillTemplateRow trow = template.Rows[groupIndex];
                FillPreviewItem item = new FillPreviewItem();
                item.IsNameDriven = true;
                item.TemplateName = template.Name;
                item.TargetRow = target.Row;
                item.ItemNo = trow.ItemNo;
                item.QuotaCode = trow.QuotaCode;
                item.Adjust = trow.Adjust;
                item.OrderInItem = trow.OrderInItem;
                item.ChosenQuotaSeq = trow.SourceQuotaSeq;
                item.NeighborSourceQuotaSeq = trow.SourceQuotaSeq;
                item.GroupOrder = groupOrder;
                item.SourceName = trow.SourceName;
                item.Unit = trow.Unit;
                item.TargetName = groupOrder == 0 ? target.DisplayName : "";
                item.TargetFullName = target.RawName;
                item.TargetContextLabel = target.ContextLabel;
                item.TargetChapter = target.Chapter;
                item.TargetUnit = target.Unit;
                item.TargetQuantityText = target.QuantityText;
                item.Selected = true;
                item.NeedManualQuota = false;
                item.AlignNote = groupOrder == 0
                    ? "模版命中"
                    : "组件框第 " + (groupOrder + 1).ToString(CultureInfo.InvariantCulture) + " 条";

                string exprText;
                List<int> operandIndexes;
                if (trow.Operands != null && trow.Operands.Count > 1 &&
                    TrySubstituteOperandQuantities(trow, targetRows, targetNorms, out exprText, out operandIndexes))
                {
                    item.QuantityText = exprText;
                    if (mergedIntoByTargetIdx != null)
                    {
                        foreach (int operandIndex in operandIndexes)
                        {
                            if (operandIndex != targetIndex && !mergedIntoByTargetIdx.ContainsKey(operandIndex))
                            {
                                mergedIntoByTargetIdx[operandIndex] = "同时参与第 " +
                                    target.Row.ToString(CultureInfo.InvariantCulture) + " 行的表达式取数";
                            }
                        }
                    }
                }
                else
                {
                    string display;
                    decimal quantity;
                    string error;
                    string firstCell = ExtractFirstCellAddress(trow.SourceExpr);
                    if (!String.IsNullOrEmpty(firstCell) &&
                        TryEvaluateExpressionWithKnownCell(trow.SourceExpr, firstCell, target.QuantityText,
                            out display, out quantity, out error))
                    {
                        item.QuantityText = display;
                    }
                    else
                    {
                        item.QuantityText = BuildNameDrivenQtyText(target.QuantityText, target.Unit, trow.Unit);
                    }
                }

                result.Add(item);
                groupOrder++;
            }
            return result;
        }

        private static List<NameQuotaCandidateGroup> BuildNameQuotaCandidates(FillTemplate template,
            List<TemplateNameGroup> groups, TargetQtyRow target, List<TargetQtyRow> targetRows,
            List<string> targetNorms, int targetIndex)
        {
            List<NameQuotaCandidateGroup> result = new List<NameQuotaCandidateGroup>();
            Dictionary<NameQuotaCandidateGroup, string> sourceExpressions = new Dictionary<NameQuotaCandidateGroup, string>();
            foreach (TemplateNameGroup group in groups ?? new List<TemplateNameGroup>())
            {
                NameQuotaCandidateGroup candidate = new NameQuotaCandidateGroup();
                candidate.Key = group.SourceAnchor ?? "";
                candidate.Label = BuildTemplateCandidateLabel(template, group);
                candidate.Items = BuildTemplatePreviewGroup(template, group, target, targetRows, targetNorms, targetIndex,
                    null);
                result.Add(candidate);
                int firstIndex = OrderedTemplateGroupIndexes(template, group).FirstOrDefault();
                sourceExpressions[candidate] = group.Indexes.Count == 0 ? "" : (template.Rows[firstIndex].SourceExpr ?? "");
            }

            HashSet<string> effectiveSignatures = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            result = result.Where(candidate => effectiveSignatures.Add(BuildNameQuotaCandidateSignature(candidate))).ToList();
            Dictionary<string, int> labelCounts = result
                .GroupBy(candidate => candidate.Label ?? "", StringComparer.OrdinalIgnoreCase)
                .ToDictionary(group => group.Key, group => group.Count(), StringComparer.OrdinalIgnoreCase);
            foreach (NameQuotaCandidateGroup candidate in result)
            {
                int count;
                if (labelCounts.TryGetValue(candidate.Label ?? "", out count) && count > 1)
                {
                    candidate.Label += "（来源 " + sourceExpressions[candidate] + "）";
                }
            }
            return result;
        }

        private static string AppendPreviewNote(string current, string note)
        {
            if (String.IsNullOrWhiteSpace(note)) return current ?? "";
            if (String.IsNullOrWhiteSpace(current)) return note;
            if (current.IndexOf(note, StringComparison.Ordinal) >= 0) return current;
            return current + "；" + note;
        }

        private static void ApplyMergedExpressionNotes(List<FillPreviewItem> items,
            List<TargetQtyRow> targetRows, Dictionary<int, string> notesByTargetIndex,
            HashSet<int> independentTargetIndexes)
        {
            if (items == null || targetRows == null || notesByTargetIndex == null) return;
            foreach (KeyValuePair<int, string> pair in notesByTargetIndex)
            {
                if (pair.Key < 0 || pair.Key >= targetRows.Count) continue;
                int targetRow = targetRows[pair.Key].Row;
                int insertAt = items.FindIndex(item => item != null && item.IsNameDriven && item.TargetRow == targetRow);
                if (independentTargetIndexes == null || !independentTargetIndexes.Contains(pair.Key))
                {
                    items.RemoveAll(item => item != null && item.IsNameDriven && item.TargetRow == targetRow);
                }
                List<FillPreviewItem> group = items
                    .Where(item => item != null && item.IsNameDriven && item.TargetRow == targetRow)
                    .OrderBy(item => item.GroupOrder)
                    .ToList();
                if (group.Count == 0)
                {
                    TargetQtyRow target = targetRows[pair.Key];
                    FillPreviewItem operandOnly = new FillPreviewItem();
                    operandOnly.IsNameDriven = true;
                    operandOnly.TargetRow = targetRow;
                    operandOnly.TargetName = target.DisplayName;
                    operandOnly.TargetFullName = target.RawName;
                    operandOnly.TargetContextLabel = target.ContextLabel;
                    operandOnly.TargetChapter = target.Chapter;
                    operandOnly.TargetUnit = target.Unit;
                    operandOnly.TargetQuantityText = target.QuantityText;
                    operandOnly.QuantityText = target.QuantityText;
                    operandOnly.AlignNote = pair.Value + "，无独立定额匹配";
                    operandOnly.Selected = false;
                    operandOnly.NeedManualQuota = false;
                    items.Insert(insertAt < 0 ? items.Count : Math.Min(insertAt, items.Count), operandOnly);
                    continue;
                }

                FillPreviewItem leader = group[0];
                for (int order = 0; order < group.Count; order++)
                {
                    group[order].GroupOrder = order;
                    group[order].TargetName = order == 0 ? targetRows[pair.Key].DisplayName : "";
                }
                bool hasIndependentQuota = group.Any(item => !String.IsNullOrWhiteSpace(item.QuotaCode));
                if (hasIndependentQuota)
                {
                    if (!String.IsNullOrWhiteSpace(leader.Status))
                    {
                        leader.Status = AppendPreviewNote(leader.Status, pair.Value);
                    }
                    else
                    {
                        leader.AlignNote = AppendPreviewNote(leader.AlignNote, pair.Value);
                    }
                    continue;
                }

                leader.Status = "";
                leader.AlignNote = pair.Value + "，无独立定额匹配";
                leader.Selected = false;
                leader.NeedManualQuota = false;
            }
        }

        // 名字驱动套用：以目标 Excel 工程量行为主序，逐行匹配定额。返回 items 已按 Excel 行序。
        private static List<FillPreviewItem> BuildPreview_NameDriven(Form mainForm, FillTemplate template,
            string targetWorkbook, string targetSheet, string targetColumn, out string warning)
        {
            warning = null;
            if (!String.Equals(template.MatchBy, "name", StringComparison.OrdinalIgnoreCase))
            {
                warning = "该模板不是“按名字生成”的模板，名字驱动无法使用。请勾选“按名字生成”重新生成模板。";
                return new List<FillPreviewItem>();
            }

            CellRef colRef;
            if (!TryParseCellAddress((targetColumn ?? "").Trim().ToUpperInvariant() + "1", out colRef))
            {
                warning = "目标列无效。";
                return new List<FillPreviewItem>();
            }
            string workbook = (targetWorkbook ?? "").Trim();
            if (String.IsNullOrWhiteSpace(workbook))
            {
                warning = "未选择目标 Excel。";
                return new List<FillPreviewItem>();
            }
            if (!File.Exists(workbook))
            {
                warning = "目标 Excel 未保存或文件不存在，请先保存后重试。";
                return new List<FillPreviewItem>();
            }

            Dictionary<int, string> ignoredChapters;
            string targetReadError;
            List<TargetQtyRow> targetRows = ReadTargetQtyRowsWithChaptersDetailed(workbook, targetSheet, colRef.Column,
                out ignoredChapters, out targetReadError);
            if (targetRows.Count == 0)
            {
                warning = "目标 Excel「" + Path.GetFileName(workbook) + "」的目标 sheet 未读到工程量行（检查目标列是否为数量列，Excel 是否已保存）。" +
                    (String.IsNullOrWhiteSpace(targetReadError) ? "" : "读取详情：" + targetReadError.Trim());
                return new List<FillPreviewItem>();
            }

            List<FillPreviewItem> items = new List<FillPreviewItem>();
            List<TemplateNameGroup> templateGroups = BuildTemplateNameGroups(template);
            Dictionary<string, List<TemplateNameGroup>> groupsByNorm = templateGroups
                .GroupBy(g => g.NormName, StringComparer.Ordinal)
                .ToDictionary(g => g.Key, g => g.ToList(), StringComparer.Ordinal);
            List<string> targetNorms = targetRows.Select(x => x.NormName).ToList();
            List<MatchTextFeatures> targetFeatures = targetNorms.Select(BuildMatchTextFeatures).ToList();
            Dictionary<string, int> targetNameCounts = targetRows
                .Where(r => !String.IsNullOrEmpty(r.NormName))
                .GroupBy(r => r.NormName, StringComparer.Ordinal)
                .ToDictionary(g => g.Key, g => g.Count(), StringComparer.Ordinal);

            // 有精确目标名称的模板组只为该名称保留，避免被前面的近似名行抢走；保留不等于已消费。
            HashSet<TemplateNameGroup> exactReservedGroups = new HashSet<TemplateNameGroup>();
            foreach (TargetQtyRow targetRow in targetRows)
            {
                List<TemplateNameGroup> exactGroups;
                if (groupsByNorm.TryGetValue(targetRow.NormName, out exactGroups))
                {
                    foreach (TemplateNameGroup group in exactGroups) exactReservedGroups.Add(group);
                }
            }

            HashSet<TemplateNameGroup> usedGroups = new HashSet<TemplateNameGroup>();
            Dictionary<int, string> mergedIntoByTargetIdx = new Dictionary<int, string>();
            HashSet<int> independentTargetIndexes = new HashSet<int>();

            FillPreviewItem lastMatched = null;
            for (int trIdx = 0; trIdx < targetRows.Count; trIdx++)
            {
                TargetQtyRow tr = targetRows[trIdx];

                FillPreviewItem item = new FillPreviewItem();
                item.IsNameDriven = true;
                item.TemplateName = template.Name;
                item.TargetRow = tr.Row;
                item.SourceName = "";
                item.TargetName = tr.DisplayName;
                item.TargetFullName = tr.RawName;
                item.TargetContextLabel = tr.ContextLabel;
                item.TargetChapter = tr.Chapter;
                item.TargetUnit = tr.Unit;
                item.TargetQuantityText = tr.QuantityText;
                item.QuantityText = tr.QuantityText;

                List<TemplateNameGroup> exactGroups;
                groupsByNorm.TryGetValue(tr.NormName, out exactGroups);
                int targetNameCount;
                targetNameCounts.TryGetValue(tr.NormName, out targetNameCount);
                List<NameQuotaCandidateGroup> exactCandidates = exactGroups == null
                    ? new List<NameQuotaCandidateGroup>()
                    : BuildNameQuotaCandidates(template, exactGroups, tr, targetRows, targetNorms, trIdx);
                int rawExactGroupCount = exactGroups == null ? 0 : exactGroups.Count;
                bool sameBindingFromMultipleSources = rawExactGroupCount > 1 && exactCandidates.Count == 1;
                string exactMode = GetExactNameResolutionMode(targetNameCount, exactCandidates.Count);
                if (exactMode.Length > 0)
                {
                    bool needsReview = exactMode == "reuse" || exactMode == "choice" || sameBindingFromMultipleSources;
                    List<FillPreviewItem> activeGroup = BuildTemplatePreviewGroup(template, exactGroups[0], tr,
                        targetRows, targetNorms, trIdx, needsReview ? null : mergedIntoByTargetIdx);
                    if (exactGroups.Any(group => group.Indexes.Any(index =>
                        template.Rows[index].Operands == null || template.Rows[index].Operands.Count < 2)))
                    {
                        independentTargetIndexes.Add(trIdx);
                    }
                    foreach (FillPreviewItem member in activeGroup)
                    {
                        member.Selected = !needsReview;
                        member.NeedExactNameConfirmation = needsReview;
                    }
                    if (needsReview && activeGroup.Count > 0)
                    {
                        activeGroup[0].AlignNote = exactMode == "choice"
                            ? "模板存在同名多来源，已带出候选，需下拉确认"
                            : (sameBindingFromMultipleSources
                                ? "模板同名多来源对应相同绑定，已默认带出，需确认"
                                : "目标表存在重复工程量名称，已带出唯一绑定，需确认");
                    }
                    if (exactMode == "choice" && activeGroup.Count > 0)
                    {
                        activeGroup[0].NameQuotaCandidates = exactCandidates;
                        activeGroup[0].SelectedNameQuotaCandidateKey = exactCandidates[0].Key;
                    }
                    if (!needsReview && activeGroup.Count > 0)
                    {
                        usedGroups.Add(exactGroups[0]);
                        lastMatched = activeGroup[0];
                    }
                    items.AddRange(activeGroup);
                    continue;
                }

                TemplateNameGroup matchedGroup = null;
                if (exactGroups == null || exactGroups.Count == 0)
                {
                    List<TemplateNameGroup> candidates = templateGroups
                        .Where(g => !usedGroups.Contains(g) && !exactReservedGroups.Contains(g))
                        .ToList();
                    bool fuzzyAmbiguous;
                    int bestGroupIndex = FindUniqueBestMatchIndexCached(targetFeatures[trIdx], tr.Chapter,
                        candidates.Select(g => g.Features).ToList(),
                        candidates.Select(g => g.Chapter ?? "").ToList(), out fuzzyAmbiguous);
                    if (fuzzyAmbiguous)
                    {
                        item.Selected = false;
                        item.NeedManualQuota = true;
                        item.AlignNote = "名称候选不唯一，需人工确认";
                        items.Add(item);
                        continue;
                    }
                    if (bestGroupIndex >= 0) matchedGroup = candidates[bestGroupIndex];
                }

                if (matchedGroup != null)
                {
                    usedGroups.Add(matchedGroup);
                    List<FillPreviewItem> fuzzyGroup = BuildTemplatePreviewGroup(template, matchedGroup, tr,
                        targetRows, targetNorms, trIdx, mergedIntoByTargetIdx);
                    if (fuzzyGroup.Count > 0) lastMatched = fuzzyGroup[0];
                    items.AddRange(fuzzyGroup);
                    continue;
                }

                item.ItemNo = lastMatched == null ? "" : lastMatched.ItemNo;
                item.NeighborSourceQuotaSeq = lastMatched == null ? 0 : lastMatched.ChosenQuotaSeq;
                item.AlignNote = "无对应定额，右键绑定软件选中定额";
                item.NeedManualQuota = true;
                item.Selected = false;
                if (lastMatched == null)
                {
                    item.Status = "无条目锚点（上方无模版命中行），不可写入";
                    item.NeedManualQuota = false;
                }
                items.Add(item);
            }
            ApplyMergedExpressionNotes(items, targetRows, mergedIntoByTargetIdx, independentTargetIndexes);
            return items;
        }

        private sealed class ProjectQuota
        {
            public string Code; public string Name; public string Unit; public long QuotaSeq;
            public decimal UnitPrice;
            public string NormCode; public string NormName; // 预计算的归一化文本，避免每次打分重复归一化
            public bool IsLibrary;  // true=来自全库 quota-index.jsonl，项目里(尚)无此编号，写入走结构模板构造
            public override string ToString()
            {
                string tail = String.IsNullOrEmpty(Unit) ? "" : "  [" + Unit + "]";
                return (IsLibrary ? "〔库〕" : "") + (Code ?? "") + "  " + (Name ?? "") + tail;
            }
        }

        private static List<ProjectQuota> LoadProjectQuotas(Form mainForm)
        {
            List<ProjectQuota> list = new List<ProjectQuota>();
            try
            {
                SqlConnection conn = GetOpenProjectConnection(mainForm);
                using (SqlCommand cmd = conn.CreateCommand())
                {
                    cmd.CommandText = "with RankedQuota as (select 定额编号, 工程或费用项目名称, 单位, 定额序号, 单价, " +
                        "row_number() over(partition by 定额编号, 工程或费用项目名称, 单位 order by " +
                        "case when 单价 is not null and 单价<>0 then 0 else 1 end, 定额序号) as rn " +
                        "from 定额输入 where 定额编号 is not null and ltrim(rtrim(定额编号))<>'' and 定额编号<>'-') " +
                        "select 定额编号, 工程或费用项目名称, 单位, 定额序号, 单价 from RankedQuota where rn=1";
                    using (SqlDataReader r = cmd.ExecuteReader())
                    {
                        while (r.Read())
                        {
                            ProjectQuota q = new ProjectQuota();
                            q.Code = r.IsDBNull(0) ? "" : Convert.ToString(r.GetValue(0)).Trim();
                            q.Name = r.IsDBNull(1) ? "" : Convert.ToString(r.GetValue(1)).Trim();
                            q.Unit = r.IsDBNull(2) ? "" : Convert.ToString(r.GetValue(2)).Trim();
                            q.QuotaSeq = r.IsDBNull(3) ? 0L : Convert.ToInt64(r.GetValue(3), CultureInfo.InvariantCulture);
                            q.UnitPrice = r.IsDBNull(4) ? 0m : Convert.ToDecimal(r.GetValue(4), CultureInfo.InvariantCulture);
                            q.NormCode = NormalizeMatchText(q.Code);
                            q.NormName = NormalizeMatchText(q.Name);
                            if (q.Code.Length > 0 && q.QuotaSeq > 0) list.Add(q);
                        }
                    }
                }
            }
            catch (Exception ex) { Log("LoadProjectQuotas failed: " + ex.Message); }
            return list;
        }

        // 右键重绑最终落位只走这一处，保证旧组先完整保留、验证成功后才原子替换。
        internal static bool ReplacePreviewTargetGroup(List<FillPreviewItem> all, int targetRow, List<FillPreviewItem> replacements)
        {
            if (all == null || replacements == null || replacements.Count == 0) return false;
            int insertAt = all.FindIndex(item => item != null && item.IsNameDriven && item.TargetRow == targetRow);
            if (insertAt < 0) return false;
            all.RemoveAll(item => item != null && item.IsNameDriven && item.TargetRow == targetRow);
            for (int i = 0; i < replacements.Count; i++)
            {
                replacements[i].GroupOrder = i;
                if (i > 0) replacements[i].TargetName = "";
            }
            all.InsertRange(Math.Min(insertAt, all.Count), replacements);
            return true;
        }

        internal static List<FillPreviewItem> MergePreviewTargetGroup(List<FillPreviewItem> existing,
            List<FillPreviewItem> additions)
        {
            List<FillPreviewItem> merged = new List<FillPreviewItem>();
            HashSet<string> identities = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (FillPreviewItem item in (existing ?? new List<FillPreviewItem>())
                .Concat(additions ?? new List<FillPreviewItem>()))
            {
                if (item == null || String.IsNullOrWhiteSpace(item.QuotaCode)) continue;
                string entryCode = !String.IsNullOrWhiteSpace(item.ChosenItemNo) ? item.ChosenItemNo : item.ItemNo;
                string identity = String.Join("\u001f", new[]
                {
                    ResolveLearningTargetKind(item.TargetKind, item.QuotaCode),
                    (item.QuotaCode ?? "").Trim().ToUpperInvariant(),
                    NormalizeForSignature(item.SourceName),
                    NormalizeForSignature(item.Unit),
                    (entryCode ?? "").Trim().ToUpperInvariant(),
                    (item.Adjust ?? "").Trim(),
                    (item.FormulaTemplate ?? "").Trim()
                });
                if (identities.Add(identity)) merged.Add(item);
            }
            return merged;
        }

        private static string BuildNameBindingTargetSetSignature(IEnumerable<FillPreviewItem> items, bool includeFormula = true)
        {
            return String.Join("\u001e", (items ?? Enumerable.Empty<FillPreviewItem>())
                .Where(item => item != null && !String.IsNullOrWhiteSpace(item.QuotaCode))
                .Select(item => String.Join("\u001f", new[]
                {
                    (item.ItemNo ?? "").Trim().ToUpperInvariant(),
                    (item.QuotaCode ?? "").Trim().ToUpperInvariant(),
                    (item.Unit ?? "").Trim().ToUpperInvariant(),
                    (item.Adjust ?? "").Trim(),
                    includeFormula ? (item.FormulaTemplate ?? "").Trim() : ""
                }))
                .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                .ToArray());
        }

        internal static bool AreEquivalentNameBindingGroups(IEnumerable<FillPreviewItem> left, IEnumerable<FillPreviewItem> right)
        {
            List<FillPreviewItem> leftItems = (left ?? Enumerable.Empty<FillPreviewItem>()).Where(item => item != null).ToList();
            List<FillPreviewItem> rightItems = (right ?? Enumerable.Empty<FillPreviewItem>()).Where(item => item != null).ToList();
            string leftName = leftItems.Select(item => String.IsNullOrWhiteSpace(item.TargetFullName) ? item.TargetName : item.TargetFullName)
                .FirstOrDefault(name => !String.IsNullOrWhiteSpace(name)) ?? "";
            string rightName = rightItems.Select(item => String.IsNullOrWhiteSpace(item.TargetFullName) ? item.TargetName : item.TargetFullName)
                .FirstOrDefault(name => !String.IsNullOrWhiteSpace(name)) ?? "";
            string leftUnit = leftItems.Select(item => item.TargetUnit).FirstOrDefault(unit => !String.IsNullOrWhiteSpace(unit)) ?? "";
            string rightUnit = rightItems.Select(item => item.TargetUnit).FirstOrDefault(unit => !String.IsNullOrWhiteSpace(unit)) ?? "";
            string leftTargets = BuildNameBindingTargetSetSignature(leftItems);
            string rightTargets = BuildNameBindingTargetSetSignature(rightItems);
            return leftTargets.Length > 0 &&
                String.Equals(NormalizeQuantityMatchName(leftName), NormalizeQuantityMatchName(rightName), StringComparison.Ordinal) &&
                String.Equals(NormalizeExcelLinkUnit(leftUnit), NormalizeExcelLinkUnit(rightUnit), StringComparison.OrdinalIgnoreCase) &&
                String.Equals(leftTargets, rightTargets, StringComparison.OrdinalIgnoreCase);
        }

        internal static bool ConfirmSingleExactNameGroup(List<FillPreviewItem> all, int targetRow)
        {
            List<FillPreviewItem> group = (all ?? new List<FillPreviewItem>())
                .Where(item => item != null && item.IsNameDriven && item.TargetRow == targetRow)
                .OrderBy(item => item.GroupOrder)
                .ToList();
            if (group.Count == 0 || (group[0].NameQuotaCandidates != null && group[0].NameQuotaCandidates.Count > 1))
            {
                return false;
            }

            ConfirmExactNameGroupInPlace(group, "人工确认重复名称");
            return true;
        }

        private static void ConfirmExactNameGroupInPlace(List<FillPreviewItem> group, string leaderNote)
        {
            for (int i = 0; i < group.Count; i++)
            {
                FillPreviewItem item = group[i];
                item.Selected = true;
                item.NeedExactNameConfirmation = false;
                item.Status = "";
                item.AlignNote = i == 0
                    ? leaderNote
                    : "组件框第 " + (i + 1).ToString(CultureInfo.InvariantCulture) + " 条（" + leaderNote + "）";
            }
        }

        internal static bool ConfirmCurrentExactNameGroup(List<FillPreviewItem> all, int targetRow)
        {
            FillPreviewItem leader = (all ?? new List<FillPreviewItem>())
                .Where(item => item != null && item.IsNameDriven && item.TargetRow == targetRow)
                .OrderBy(item => item.GroupOrder)
                .FirstOrDefault();
            if (leader == null) return false;
            if (leader.NameQuotaCandidates != null && leader.NameQuotaCandidates.Count > 1)
            {
                if (String.IsNullOrWhiteSpace(leader.SelectedNameQuotaCandidateKey)) return false;
                List<FillPreviewItem> group = (all ?? new List<FillPreviewItem>())
                    .Where(item => item != null && item.IsNameDriven && item.TargetRow == targetRow)
                    .OrderBy(item => item.GroupOrder)
                    .ToList();
                ConfirmExactNameGroupInPlace(group, "人工选择同名绑定");
                return group.Count > 0;
            }
            return ConfirmSingleExactNameGroup(all, targetRow);
        }

        internal static bool ApplyExactNameCandidate(List<FillPreviewItem> all, int targetRow, string candidateKey)
        {
            FillPreviewItem leader = (all ?? new List<FillPreviewItem>())
                .Where(item => item != null && item.IsNameDriven && item.TargetRow == targetRow)
                .OrderBy(item => item.GroupOrder)
                .FirstOrDefault();
            if (leader == null || leader.NameQuotaCandidates == null) return false;

            NameQuotaCandidateGroup candidate = leader.NameQuotaCandidates
                .FirstOrDefault(option => String.Equals(option.Key, candidateKey, StringComparison.Ordinal));
            if (candidate == null || candidate.Items == null || candidate.Items.Count == 0) return false;

            List<FillPreviewItem> replacements = candidate.Items
                .Where(item => item != null)
                .Select(item => item.CloneForNameCandidate())
                .ToList();
            if (replacements.Count == 0) return false;
            bool hasRisk = replacements.Any(item => !String.IsNullOrWhiteSpace(item.Status));
            if (!hasRisk)
            {
                ConfirmExactNameGroupInPlace(replacements, "人工选择同名绑定");
            }
            else
            {
                foreach (FillPreviewItem item in replacements)
                {
                    item.Selected = false;
                    item.NeedExactNameConfirmation = true;
                    item.AlignNote = AppendPreviewNote(item.AlignNote, "人工选择候选，仍需处理组件风险");
                }
            }
            replacements[0].NameQuotaCandidates = leader.NameQuotaCandidates;
            replacements[0].SelectedNameQuotaCandidateKey = candidate.Key;
            return ReplacePreviewTargetGroup(all, targetRow, replacements);
        }

        private static MappingFeedbackGroup BuildTemplateRightClickFeedbackGroup(List<FillPreviewItem> items,
            string sourceWorkbook, string sourceSheet, SqlConnection projectConn,
            int acceptedCount, int correctedCount, int rejectedCount, string userAction)
        {
            List<FillPreviewItem> valid = (items ?? new List<FillPreviewItem>())
                .Where(item => item != null && item.IsNameDriven && !String.IsNullOrWhiteSpace(item.QuotaCode))
                .ToList();
            if (valid.Count == 0) return null;
            FillPreviewItem quantityItem = valid.FirstOrDefault(item => !String.IsNullOrWhiteSpace(item.TargetFullName));
            string name = quantityItem == null ? "" : quantityItem.TargetFullName;
            if (String.IsNullOrWhiteSpace(name)) return null;
            string quantityUnit = valid.Select(item => item.TargetUnit)
                .FirstOrDefault(unit => !String.IsNullOrWhiteSpace(unit)) ?? "";
            string contextLabel = valid.Select(item => item.TargetContextLabel)
                .FirstOrDefault(value => !String.IsNullOrWhiteSpace(value)) ?? "";
            if (String.IsNullOrWhiteSpace(contextLabel) && quantityItem != null)
            {
                contextLabel = quantityItem.TargetChapter ?? "";
            }
            MappingFeedbackGroup mappingGroup = new MappingFeedbackGroup
            {
                QuantityName = name,
                QuantityUnit = quantityUnit,
                QuantityContext = contextLabel,
                EntryCode = valid.Select(item => item.ItemNo).FirstOrDefault(no => !String.IsNullOrWhiteSpace(no)) ?? "",
                Workbook = sourceWorkbook ?? "",
                Worksheet = sourceSheet ?? "",
                ExcelRow = valid[0].TargetRow,
                AcceptedCount = acceptedCount,
                CorrectedCount = correctedCount,
                RejectedCount = rejectedCount,
                UserAction = userAction ?? ""
            };
            PopulateMappingFeedbackGroupProjectContext(projectConn, mappingGroup);
            if (acceptedCount + correctedCount > 0)
            {
                FillPreviewItem formulaSource = valid.FirstOrDefault(item =>
                    !String.IsNullOrWhiteSpace(item.FormulaTemplate) && item.FormulaOperands != null);
                if (formulaSource != null)
                {
                    mappingGroup.FormulaOperands.AddRange(formulaSource.FormulaOperands.Select(operand => new QuantityFormulaOperandInfo
                    {
                        Name = operand.Name,
                        Unit = operand.Unit,
                        Signature = operand.Signature
                    }));
                }
            }
            foreach (FillPreviewItem item in valid)
            {
                mappingGroup.Targets.Add(new MappingFeedbackTarget
                {
                    Kind = ResolveLearningTargetKind(item.TargetKind, item.QuotaCode),
                    Code = item.QuotaCode,
                    Name = item.SourceName,
                    Unit = item.Unit,
                    EntryCode = !String.IsNullOrWhiteSpace(item.ChosenItemNo) ? item.ChosenItemNo : item.ItemNo,
                    EntryName = !String.IsNullOrWhiteSpace(item.ChosenItemName) ? item.ChosenItemName : mappingGroup.EntryName,
                    FormulaTemplate = acceptedCount + correctedCount > 0 ? item.FormulaTemplate : "",
                    ManualFormulaOverride = acceptedCount + correctedCount > 0 && item.ManualFormulaOverride,
                    QuotaSequence = item.WrittenQuotaSequence > 0 ? item.WrittenQuotaSequence : item.ChosenQuotaSeq,
                    SourceEndpointIdentity = item.WrittenQuotaSequence > 0
                        ? GetProjectConnectionIdentity(projectConn)
                        : item.SourceEndpointIdentity,
                    UnitPrice = FilterLearningTargetUnitPrice(item.QuotaCode, item.LearnedUnitPrice),
                    EntrySource = item.EntrySource ?? ""
                });
            }
            return mappingGroup;
        }

        private static void ReplaceTemplateWithManualBinding(FillTemplate template, List<FillPreviewItem> replacements,
            List<FillPreviewItem> replaced)
        {
            if (template == null || template.Rows == null || replacements == null || replacements.Count == 0) return;
            FillPreviewItem leader = replacements.FirstOrDefault(item => item != null);
            if (leader == null) return;
            string fullName = String.IsNullOrWhiteSpace(leader.TargetFullName) ? (leader.TargetName ?? "") : leader.TargetFullName;
            string normName = NormalizeQuantityMatchName(fullName);
            string chapter = leader.TargetChapter ?? "";
            string replacedTargets = BuildNameBindingTargetSetSignature(replaced, false);
            HashSet<int> removeIndexes = new HashSet<int>();
            foreach (TemplateNameGroup group in BuildTemplateNameGroups(template).Where(group =>
                String.Equals(group.NormName, normName, StringComparison.Ordinal)))
            {
                bool sameScope = !String.IsNullOrWhiteSpace(chapter)
                    ? SameTemplateChapter(group.Chapter, chapter)
                    : String.Equals(BuildNameBindingTargetSetSignature(group.Indexes.Select(index =>
                        new FillPreviewItem
                        {
                            ItemNo = template.Rows[index].ItemNo,
                            QuotaCode = template.Rows[index].QuotaCode,
                            Unit = template.Rows[index].Unit,
                            Adjust = template.Rows[index].Adjust
                        }), false), replacedTargets, StringComparison.OrdinalIgnoreCase);
                if (!sameScope) continue;
                foreach (int index in group.Indexes) removeIndexes.Add(index);
            }
            template.Rows = template.Rows.Where((row, index) => !removeIndexes.Contains(index)).ToList();
            int order = 0;
            foreach (FillPreviewItem item in replacements.Where(item => item != null && !String.IsNullOrWhiteSpace(item.QuotaCode)))
            {
                template.Rows.Add(new FillTemplateRow
                {
                    ItemNo = item.ItemNo,
                    ItemName = item.ItemNo,
                    QuotaCode = item.QuotaCode,
                    MatchName = fullName,
                    SourceName = String.IsNullOrEmpty(item.SourceName) ? fullName : item.SourceName,
                    MatchChapter = chapter,
                    Unit = item.Unit,
                    Adjust = item.Adjust,
                    SourceQuotaSeq = item.ChosenQuotaSeq,
                    OrderInItem = order++,
                    Origin = FillTemplateOriginManual
                });
            }
        }

        // 仅名字驱动右键新增/实质性重绑调用：新组扶正，旧组否定，并更新当前模板。
        private static void FeedbackNameMatches(string templateName, List<FillPreviewItem> written,
            string sourceWorkbook = "", string sourceSheet = "", SqlConnection projectConn = null,
            List<FillPreviewItem> replaced = null)
        {
            List<MappingFeedbackGroup> mappingGroups = new List<MappingFeedbackGroup>();
            foreach (IGrouping<int, FillPreviewItem> group in (written ?? new List<FillPreviewItem>())
                .Where(item => item != null && item.IsNameDriven && !item.LearningFeedbackAttempted &&
                    !String.IsNullOrWhiteSpace(item.QuotaCode))
                .GroupBy(item => item.TargetRow))
            {
                MappingFeedbackGroup corrected = BuildTemplateRightClickFeedbackGroup(group.ToList(),
                    sourceWorkbook, sourceSheet, projectConn, 0, 1, 0, "correction");
                if (corrected != null) mappingGroups.Add(corrected);
            }
            foreach (IGrouping<int, FillPreviewItem> group in (replaced ?? new List<FillPreviewItem>())
                .Where(item => item != null && item.IsNameDriven && !String.IsNullOrWhiteSpace(item.QuotaCode))
                .GroupBy(item => item.TargetRow))
            {
                MappingFeedbackGroup rejected = BuildTemplateRightClickFeedbackGroup(group.ToList(),
                    sourceWorkbook, sourceSheet, projectConn, 0, 0, 1, "rejection");
                if (rejected != null) mappingGroups.Add(rejected);
            }
            try
            {
                FillTemplate t = LoadFillTemplate(templateName);
                // 推荐定额模式没有本地模板；模板回写可跳过，但 SQL 学习仍必须继续。
                if (t != null && String.Equals(t.MatchBy, "name", StringComparison.OrdinalIgnoreCase))
                {
                    foreach (IGrouping<int, FillPreviewItem> group in (written ?? new List<FillPreviewItem>())
                        .Where(item => item != null && item.IsNameDriven && !String.IsNullOrWhiteSpace(item.QuotaCode))
                        .GroupBy(item => item.TargetRow))
                    {
                        List<FillPreviewItem> oldGroup = (replaced ?? new List<FillPreviewItem>())
                            .Where(item => item != null && item.TargetRow == group.Key)
                            .ToList();
                        ReplaceTemplateWithManualBinding(t, group.ToList(), oldGroup);
                    }
                    SaveFillTemplate(t);
                }
            }
            catch (Exception ex)
            {
                Log("FeedbackNameMatches template writeback failed: " + ex.Message);
            }

            RecordNameMatchesToLearningDb(mappingGroups);
            bool learningDurable = ConsumeLearningDbDurableResult(mappingGroups);
            foreach (FillPreviewItem item in mappingGroups.Count == 0
                ? Enumerable.Empty<FillPreviewItem>()
                : (written ?? new List<FillPreviewItem>())
                    .Where(item => item != null && item.IsNameDriven && !item.LearningFeedbackAttempted))
            {
                item.LearningFeedbackAttempted = true;
                item.SqlFeedbackDurable = learningDurable;
            }
        }
    }
}
