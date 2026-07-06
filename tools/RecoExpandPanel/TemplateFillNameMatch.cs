using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;

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
    }
}
