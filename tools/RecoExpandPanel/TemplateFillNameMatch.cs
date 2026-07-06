using System;
using System.Collections.Generic;
using System.Data.SqlClient;
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
            StringBuilder sb = new StringBuilder(text.Length);
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];
                if (c == '　') continue;
                if (c >= '！' && c <= '～') c = (char)(c - 0xFEE0);
                if (Char.IsWhiteSpace(c) || !Char.IsLetterOrDigit(c)) continue;
                sb.Append(Char.ToLowerInvariant(c));
            }
            return sb.ToString();
        }

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
                    result.Add(normText.Substring(start, i - start).TrimEnd('.'));
                }
                else i++;
            }
            return result;
        }

        private static HashSet<string> BuildMatchBigrams(string text)
        {
            HashSet<string> grams = new HashSet<string>(StringComparer.Ordinal);
            if (text.Length == 1) { grams.Add(text); return grams; }
            for (int i = 0; i + 1 < text.Length; i++) grams.Add(text.Substring(i, 2));
            return grams;
        }

        // 相似度 0-100：字符 bigram Dice；双方都含数字且数字集不相交时重罚(/3)。
        internal static int MatchNameScore(string leftNorm, string rightNorm)
        {
            if (String.IsNullOrEmpty(leftNorm) || String.IsNullOrEmpty(rightNorm)) return 0;
            if (String.Equals(leftNorm, rightNorm, StringComparison.Ordinal)) return 100;
            HashSet<string> l = BuildMatchBigrams(leftNorm);
            HashSet<string> r = BuildMatchBigrams(rightNorm);
            if (l.Count == 0 || r.Count == 0) return 0;
            int common = l.Count(g => r.Contains(g));
            int score = (int)Math.Round(200.0 * common / (l.Count + r.Count));
            List<string> ln = ExtractMatchNumbers(leftNorm);
            List<string> rn = ExtractMatchNumbers(rightNorm);
            if (ln.Count > 0 && rn.Count > 0 && !ln.Any(n => rn.Contains(n))) score /= 3;
            return score > 100 ? 100 : score;
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

        // 名字模式模版生成：与 BuildFillTemplateFromBindings 同源，额外为每行读 Excel 工程量全名；
        // 表达式(E1+E2)拆操作数各读全名存 Operands，套用时按名字定位、不再绑坐标。
        private static FillTemplate BuildNameFillTemplateFromBindings(
            Form mainForm, SqlConnection conn, string templateName, string unitNo, string sourceSheet)
        {
            FillTemplate template = BuildFillTemplateFromBindings(mainForm, conn, templateName, unitNo, sourceSheet);
            template.MatchBy = "name";

            Dictionary<string, HashSet<int>> hiddenCache = new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, List<ExcelMergedRegion>> mergedCache = new Dictionary<string, List<ExcelMergedRegion>>(StringComparer.OrdinalIgnoreCase);

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
                    row.MatchName = ReadFullNameForCell(workbook, row.SourceSheet, row.SourceExpr, hiddenCache, mergedCache);
                }
                else
                {
                    row.Operands = new List<FillOperand>();
                    foreach (string cell in cells)
                    {
                        FillOperand op = new FillOperand();
                        op.Op = "+";
                        op.Name = ReadFullNameForCell(workbook, row.SourceSheet, cell, hiddenCache, mergedCache);
                        row.Operands.Add(op);
                    }
                    row.MatchName = row.Operands.Count > 0 ? row.Operands[0].Name : "";
                }
            }
            return template;
        }

        // 读某表达式首格所在行的【全名】(不截断)。复用绑定阶段的不截断 ReadRowNameAt 重载。
        private static string ReadFullNameForCell(string workbook, string sheet, string expr,
            Dictionary<string, HashSet<int>> hiddenCache, Dictionary<string, List<ExcelMergedRegion>> mergedCache)
        {
            List<ExcelQuotaLink> readLinks = new List<ExcelQuotaLink>();
            AddQuantityNameReadLinks(readLinks, workbook, sheet, expr, hiddenCache, mergedCache);
            ExcelSyncReadContext ctx = new ExcelSyncReadContext(readLinks);
            return ReadRowNameAt(workbook, sheet, expr, hiddenCache, mergedCache, ctx, true);
        }

        private sealed class TargetQtyRow
        {
            public int Row;
            public string RawName;    // 数量列左侧全名(不截断)
            public string NormName;   // 归一化
            public string Chapter;    // 所属 Excel 章节锚点行文本(供章节内就近)；空=未分段
            public decimal Quantity;
            public string QuantityText;
            public bool IsAnchor;     // 章节锚点行(“一、/(一)/第X章”)
        }

        // 读目标 sheet：数量列(qtyColumn) 有数字的行=工程量行；行全名取数量列左侧不截断文本；
        // 章节锚点行用于给每个工程量行标 Chapter(取其上方最近锚点)。
        private static List<TargetQtyRow> ReadTargetQtyRows(string workbook, string sheet, int qtyColumn)
        {
            List<TargetQtyRow> result = new List<TargetQtyRow>();
            Dictionary<string, HashSet<int>> hiddenCache = new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, List<ExcelMergedRegion>> mergedCache = new Dictionary<string, List<ExcelMergedRegion>>(StringComparer.OrdinalIgnoreCase);

            List<ExcelQuotaLink> readLinks = new List<ExcelQuotaLink>();
            int firstRow, lastRow;
            if (!TryGetSheetRowRange(workbook, sheet, out firstRow, out lastRow)) return result;
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
                string name = ReadRowNameAt(workbook, sheet, qtyAddr, hiddenCache, mergedCache, ctx, true);
                if (IsChapterAnchorRaw(name)) { currentChapter = name; continue; }

                string disp; decimal qty; string err;
                bool hasQty = TryEvaluateWorkbookExpression(ctx, workbook, sheet, qtyAddr, out disp, out qty, out err, true) && qty != 0m;
                if (!hasQty || String.IsNullOrWhiteSpace(name)) continue;

                TargetQtyRow row = new TargetQtyRow();
                row.Row = r;
                row.RawName = name;
                row.NormName = NormalizeMatchText(name);
                row.Chapter = currentChapter;
                row.Quantity = qty;
                row.QuantityText = disp;
                result.Add(row);
            }
            return result;
        }

        // “一、/(一)/第X章/第X部分” 视为章节锚点。
        private static bool IsChapterAnchorRaw(string raw)
        {
            if (String.IsNullOrWhiteSpace(raw)) return false;
            string t = raw.Trim();
            return System.Text.RegularExpressions.Regex.IsMatch(t,
                "^(第?[一二三四五六七八九十百]+[、\\.．章节部分]|[(（][一二三四五六七八九十]+[)）])");
        }

        // 取 sheet 的 UsedRange 行范围（首末非空行号）。
        private static bool TryGetSheetRowRange(string workbook, string sheet, out int firstRow, out int lastRow)
        {
            firstRow = 0; lastRow = 0;
            try
            {
                if (!TryGetXlsxUsedRowRange(workbook, sheet, out firstRow, out lastRow)) return false;
                return lastRow >= firstRow && lastRow - firstRow < 20000;
            }
            catch { return false; }
        }

        // 用 NPOI 打开工作簿(只读共享流)，取该 sheet 的首末非空行号(1基)。
        // 打不开工作簿或找不到该 sheet 返回 false。参照 ExcelLinkFeature.cs 里
        // TryReadSheetTargetCellsByNpoi/ReadSheetCellsByNpoi 等既有 NPOI 只读用法。
        private static bool TryGetXlsxUsedRowRange(string workbook, string sheet, out int firstRow, out int lastRow)
        {
            firstRow = 0; lastRow = 0;
            if (String.IsNullOrWhiteSpace(workbook) || String.IsNullOrWhiteSpace(sheet) || !File.Exists(workbook))
            {
                return false;
            }

            try
            {
                using (Stream stream = OpenWorkbookStreamShared(workbook))
                {
                    IWorkbook wb = WorkbookFactory.Create(stream);
                    ISheet sh = wb.GetSheet(sheet);
                    if (sh == null) return false;

                    int first0 = sh.FirstRowNum;
                    int last0 = sh.LastRowNum;
                    if (last0 < first0) return false;

                    firstRow = first0 + 1;
                    lastRow = last0 + 1;
                    return true;
                }
            }
            catch (Exception ex)
            {
                Log("TryGetXlsxUsedRowRange 失败: " + ex.Message);
                return false;
            }
        }
    }
}
