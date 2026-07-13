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
            public string RawName;    // 数量列左侧全名(不截断，供匹配)
            public string DisplayName;// 数量列左侧截断显示名(3段/15字)，供 UI 显示
            public string NormName;   // 归一化
            public string Chapter;    // 预留：二期章节内就近约束
            public decimal Quantity;
            public string QuantityText;
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
                row.DisplayName = ReadRowNameAt(workbook, sheet, qtyAddr, hiddenCache, mergedCache, ctx, false);
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

        // 费用/辅助类伪代码(排组件框末尾；真定额排前)。含 *系数 后缀先剥掉再判。
        private static int PseudoQuotaRank(string code)
        {
            string c = (code ?? "").Trim().ToUpperInvariant();
            int star = c.IndexOf('*');
            if (star > 0) c = c.Substring(0, star);
            switch (c)
            {
                case "SF": case "SH": case "SQ": case "ZLF": case "LF":
                case "YF": case "TLF": case "GF": case "JF": case "XGT1":
                    return 1;
                default:
                    return 0;
            }
        }

        private sealed class BoxCandidate
        {
            public string QuotaCode;
            public string QuotaName;
            public int Score;
        }

        // 读 mapping-boxes.jsonl 全量行；一次性读入供本次预览多次查询，避免逐行重复打开文件。
        private static List<Dictionary<string, string>> LoadMappingBoxRows()
        {
            List<Dictionary<string, string>> rows = new List<Dictionary<string, string>>();
            try
            {
                string path = System.IO.Path.Combine(FindRecoQuotaDataDir(), "mapping-boxes.jsonl");
                if (!System.IO.File.Exists(path)) return rows;
                foreach (string line in System.IO.File.ReadAllLines(path, Encoding.UTF8))
                {
                    Dictionary<string, string> row = ParseFlatJson(line);
                    if (row.Count > 0) rows.Add(row);
                }
            }
            catch (Exception ex) { Log("LoadMappingBoxRows failed: " + ex.Message); }
            return rows;
        }

        // 为一个工程量全名返回候选定额(按名字相似度)。一期只吃对应框(绑定飞轮)。boxRows 由调用方预读一次。
        private static List<BoxCandidate> LookupMappingBox(string queryFullName, List<Dictionary<string, string>> boxRows)
        {
            List<BoxCandidate> result = new List<BoxCandidate>();
            string norm = NormalizeMatchText(queryFullName);
            if (norm.Length == 0 || boxRows == null) return result;
            foreach (Dictionary<string, string> row in boxRows)
            {
                string qn = GetFlat(row, "quantity_name");
                string code = GetFlat(row, "target_code");
                if (String.IsNullOrWhiteSpace(qn) || String.IsNullOrWhiteSpace(code)) continue;
                int s = MatchNameScore(norm, NormalizeMatchText(qn));
                if (s < NameMatchMinScore) continue;
                result.Add(new BoxCandidate { QuotaCode = code, QuotaName = GetFlat(row, "target_name"), Score = s });
            }
            return result
                .GroupBy(c => c.QuotaCode, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.OrderByDescending(c => c.Score).First())
                .OrderByDescending(c => c.Score)
                .ToList();
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
                int idx = BestMatchIndex(NormalizeMatchText(op.Name), targetNorms);
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

        // 名字驱动套用：以目标 Excel 工程量行为主序，逐行匹配定额。返回 items 已按 Excel 行序。
        private static List<FillPreviewItem> BuildPreview_NameDriven(Form mainForm, FillTemplate template,
            string targetSheet, string targetColumn, out string warning)
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
            string workbook = template.WorkbookPath;
            if (String.IsNullOrWhiteSpace(workbook))
            {
                warning = "模版未记录 Excel 文件。";
                return new List<FillPreviewItem>();
            }

            List<TargetQtyRow> targetRows = ReadTargetQtyRows(workbook, targetSheet, colRef.Column);
            if (targetRows.Count == 0)
            {
                warning = "目标 sheet 未读到工程量行（检查目标列是否为数量列，Excel 是否已保存）。";
                return new List<FillPreviewItem>();
            }

            List<ProjectQuota> projectQuotas = LoadProjectQuotas(mainForm);
            List<Dictionary<string, string>> boxRows = LoadMappingBoxRows();

            List<FillPreviewItem> items = new List<FillPreviewItem>();
            List<string> tmplNorms = template.Rows.Select(r => NormalizeMatchText(r.MatchName ?? r.SourceName ?? "")).ToList();
            List<string> targetNorms = targetRows.Select(x => x.NormName).ToList();

            HashSet<int> usedTmplIdx = new HashSet<int>();
            Dictionary<int, string> mergedIntoByTargetIdx = new Dictionary<int, string>();

            // 两遍匹配：先“归一化完全相等”精确认领，再模糊。防止近似名(如 -2*1.5 vs -4*1.5)抢走精确归属的模板组。
            Dictionary<int, int> exactTmplByTarget = new Dictionary<int, int>();
            HashSet<int> exactClaimedTmpl = new HashSet<int>();
            for (int t = 0; t < targetRows.Count; t++)
            {
                for (int gi = 0; gi < tmplNorms.Count; gi++)
                {
                    if (exactClaimedTmpl.Contains(gi) || tmplNorms[gi].Length == 0)
                    {
                        continue;
                    }
                    if (String.Equals(tmplNorms[gi], targetRows[t].NormName, StringComparison.Ordinal))
                    {
                        exactTmplByTarget[t] = gi;
                        for (int gj = 0; gj < tmplNorms.Count; gj++)
                        {
                            if (String.Equals(tmplNorms[gj], tmplNorms[gi], StringComparison.Ordinal))
                            {
                                exactClaimedTmpl.Add(gj);
                            }
                        }
                        break;
                    }
                }
            }

            FillPreviewItem lastMatched = null;
            for (int trIdx = 0; trIdx < targetRows.Count; trIdx++)
            {
                TargetQtyRow tr = targetRows[trIdx];

                string mergedNote;
                if (mergedIntoByTargetIdx.TryGetValue(trIdx, out mergedNote))
                {
                    FillPreviewItem mergedItem = new FillPreviewItem();
                    mergedItem.IsNameDriven = true;
                    mergedItem.TemplateName = template.Name;
                    mergedItem.TargetRow = tr.Row;
                    mergedItem.TargetName = tr.DisplayName;
                    mergedItem.TargetFullName = tr.RawName;
                    mergedItem.QuantityText = tr.QuantityText;
                    mergedItem.AlignNote = mergedNote;
                    mergedItem.Selected = false;
                    mergedItem.NeedManualQuota = false;
                    items.Add(mergedItem);
                    continue;
                }

                FillPreviewItem item = new FillPreviewItem();
                item.IsNameDriven = true;
                item.TemplateName = template.Name;
                item.TargetRow = tr.Row;
                item.SourceName = "";
                item.TargetName = tr.DisplayName;
                item.TargetFullName = tr.RawName;
                item.QuantityText = tr.QuantityText;

                int ti;
                if (!exactTmplByTarget.TryGetValue(trIdx, out ti))
                {
                    ti = -1;
                    int bestScore = NameMatchMinScore - 1;
                    for (int gi = 0; gi < tmplNorms.Count; gi++)
                    {
                        if (usedTmplIdx.Contains(gi) || exactClaimedTmpl.Contains(gi))
                        {
                            continue;
                        }
                        int s = MatchNameScore(tr.NormName, tmplNorms[gi]);
                        if (s > bestScore) { bestScore = s; ti = gi; }
                    }
                }
                else if (usedTmplIdx.Contains(ti))
                {
                    ti = -1;
                }

                if (ti >= 0)
                {
                    // 组件框：目标一行工程量 -> 模版里所有同名(归一化相等)且未被消费过的定额，按条目内序全展开。
                    string bestNorm = tmplNorms[ti];
                    List<int> groupIdx = new List<int>();
                    for (int gi = 0; gi < tmplNorms.Count; gi++)
                    {
                        if (usedTmplIdx.Contains(gi)) continue;
                        if (String.Equals(tmplNorms[gi], bestNorm, StringComparison.Ordinal)) groupIdx.Add(gi);
                    }
                    groupIdx.Sort(delegate(int a, int b)
                    {
                        int ra = PseudoQuotaRank(template.Rows[a].QuotaCode);
                        int rb = PseudoQuotaRank(template.Rows[b].QuotaCode);
                        if (ra != rb) return ra.CompareTo(rb);
                        return template.Rows[a].OrderInItem.CompareTo(template.Rows[b].OrderInItem);
                    });
                    foreach (int gi in groupIdx) usedTmplIdx.Add(gi);

                    int go = 0;
                    foreach (int gi in groupIdx)
                    {
                        FillTemplateRow trow = template.Rows[gi];
                        FillPreviewItem gitem = (go == 0) ? item : new FillPreviewItem();
                        gitem.IsNameDriven = true;
                        gitem.TemplateName = template.Name;
                        gitem.TargetRow = tr.Row;
                        gitem.ItemNo = trow.ItemNo;
                        gitem.QuotaCode = trow.QuotaCode;
                        gitem.Adjust = trow.Adjust;
                        gitem.OrderInItem = trow.OrderInItem;
                        gitem.ChosenQuotaSeq = trow.SourceQuotaSeq;
                        gitem.NeighborSourceQuotaSeq = trow.SourceQuotaSeq; // 自锚点：模版命中行不依赖上方行也可写入
                        gitem.GroupOrder = go;
                        gitem.SourceName = trow.SourceName;
                        gitem.Unit = trow.Unit;
                        gitem.TargetName = (go == 0) ? tr.DisplayName : "";
                        gitem.TargetFullName = tr.RawName;
                        gitem.AlignNote = (go == 0) ? "模版命中" : ("组件框第 " + (go + 1).ToString(CultureInfo.InvariantCulture) + " 条");

                        // 数量：多操作数表达式(如 E4+E5)优先按名字把各操作数代入源表达式求值；
                        // 否则套用源绑定表达式的换算系数(如 I19/100)；否则直接用目标数量。
                        string exprText;
                        List<int> operandIdx;
                        if (trow.Operands != null && trow.Operands.Count > 1 &&
                            TrySubstituteOperandQuantities(trow, targetRows, targetNorms, out exprText, out operandIdx))
                        {
                            gitem.QuantityText = exprText;
                            for (int oi = 0; oi < operandIdx.Count; oi++)
                            {
                                int oIdx = operandIdx[oi];
                                if (oIdx != trIdx && !mergedIntoByTargetIdx.ContainsKey(oIdx))
                                {
                                    mergedIntoByTargetIdx[oIdx] = "已并入第 " + tr.Row.ToString(CultureInfo.InvariantCulture) + " 行的表达式取数";
                                }
                            }
                        }
                        else
                        {
                            string fdisp; decimal fqty; string ferr;
                            string fcell = ExtractFirstCellAddress(trow.SourceExpr);
                            if (!String.IsNullOrEmpty(fcell) && TryEvaluateExpressionWithKnownCell(trow.SourceExpr, fcell, tr.QuantityText, out fdisp, out fqty, out ferr))
                            {
                                gitem.QuantityText = fdisp;
                            }
                            else
                            {
                                gitem.QuantityText = tr.QuantityText;
                            }
                        }

                        if (go == 0) lastMatched = gitem;
                        items.Add(gitem);
                        go++;
                    }
                    continue;
                }

                List<BoxCandidate> box = LookupMappingBox(tr.RawName, boxRows);
                if (box.Count > 0 && box[0].Score >= 70)
                {
                    item.QuotaCode = box[0].QuotaCode;
                    item.ChosenQuotaSeq = LoadProjectQuotaSeqByCode(projectQuotas, box[0].QuotaCode);
                    ProjectQuota boxQuota = projectQuotas.FirstOrDefault(x => String.Equals(x.Code, box[0].QuotaCode, StringComparison.OrdinalIgnoreCase));
                    item.Unit = boxQuota == null ? "" : boxQuota.Unit;
                    item.ItemNo = lastMatched == null ? "" : lastMatched.ItemNo;
                    item.NeighborSourceQuotaSeq = lastMatched == null ? 0 : lastMatched.ChosenQuotaSeq;
                    item.AlignNote = item.ChosenQuotaSeq > 0
                        ? ("对应框建议 " + box[0].QuotaCode + "，可直接勾选写入，双击可改")
                        : ("对应框建议 " + box[0].QuotaCode + "（项目内无此定额），双击手挂");
                    item.NeedManualQuota = true;
                    item.Selected = false;
                    if (lastMatched == null)
                    {
                        item.Status = "无条目锚点（上方无模版命中行），不可写入";
                        item.NeedManualQuota = false;
                    }
                }
                else
                {
                    item.ItemNo = lastMatched == null ? "" : lastMatched.ItemNo;
                    item.NeighborSourceQuotaSeq = lastMatched == null ? 0 : lastMatched.ChosenQuotaSeq;
                    item.AlignNote = "无对应定额，双击手挂";
                    item.NeedManualQuota = true;
                    item.Selected = false;
                    if (lastMatched == null)
                    {
                        item.Status = "无条目锚点（上方无模版命中行），不可写入";
                        item.NeedManualQuota = false;
                    }
                }
                items.Add(item);
            }
            return items;
        }

        private sealed class ProjectQuota
        {
            public string Code; public string Name; public string Unit; public long QuotaSeq;
            public string NormCode; public string NormName; // 预计算的归一化文本，避免每次打分重复归一化
            public override string ToString()
            { return (Code ?? "") + "  " + (Name ?? "") + (String.IsNullOrEmpty(Unit) ? "" : "  [" + Unit + "]"); }
        }

        private static List<ProjectQuota> LoadProjectQuotas(Form mainForm)
        {
            List<ProjectQuota> list = new List<ProjectQuota>();
            try
            {
                using (SqlConnection conn = AgentCreateWorkConnection(mainForm))
                {
                    EnsureOpen(conn);
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "select 定额编号, 工程或费用项目名称, 单位, min(定额序号) from 定额输入 " +
                            "where 定额编号 is not null and ltrim(rtrim(定额编号))<>'' and 定额编号<>'-' " +
                            "group by 定额编号, 工程或费用项目名称, 单位";
                        using (SqlDataReader r = cmd.ExecuteReader())
                        {
                            while (r.Read())
                            {
                                ProjectQuota q = new ProjectQuota();
                                q.Code = r.IsDBNull(0) ? "" : Convert.ToString(r.GetValue(0)).Trim();
                                q.Name = r.IsDBNull(1) ? "" : Convert.ToString(r.GetValue(1)).Trim();
                                q.Unit = r.IsDBNull(2) ? "" : Convert.ToString(r.GetValue(2)).Trim();
                                q.QuotaSeq = r.IsDBNull(3) ? 0L : Convert.ToInt64(r.GetValue(3), CultureInfo.InvariantCulture);
                                q.NormCode = NormalizeMatchText(q.Code);
                                q.NormName = NormalizeMatchText(q.Name);
                                if (q.Code.Length > 0 && q.QuotaSeq > 0) list.Add(q);
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { Log("LoadProjectQuotas failed: " + ex.Message); }
            return list;
        }

        private static long LoadProjectQuotaSeqByCode(List<ProjectQuota> all, string code)
        {
            if (String.IsNullOrWhiteSpace(code)) return 0;
            ProjectQuota q = all.FirstOrDefault(x => String.Equals(x.Code, code.Trim(), StringComparison.OrdinalIgnoreCase));
            return q == null ? 0 : q.QuotaSeq;
        }

        private static List<ProjectQuota> RankProjectQuotas(List<ProjectQuota> all, string name)
        {
            string q = NormalizeMatchText(name ?? "");
            return all
                .Select(o => new KeyValuePair<int, ProjectQuota>(MatchNameScore(q, o.NormName), o))
                .Where(p => p.Key > 0)
                .OrderByDescending(p => p.Key)
                .Take(8)
                .Select(p => p.Value)
                .ToList();
        }

        // 把手挂结果落到预览：单条=补当前行；多条=当前行为组第一行，其余克隆插入(组内序 GroupOrder)。
        private static void ApplyManualQuotaPick(List<FillPreviewItem> preview, FillPreviewItem target, List<ProjectQuota> picked)
        {
            if (preview == null || target == null || picked == null || picked.Count == 0) return;
            int idx = preview.IndexOf(target);
            target.NeedManualQuota = false;
            target.Selected = true;
            target.QuotaCode = picked[0].Code;
            target.ChosenQuotaSeq = picked[0].QuotaSeq;
            target.SourceName = picked[0].Name;
            target.Unit = picked[0].Unit;
            target.GroupOrder = 0;
            target.AlignNote = "已手挂 " + picked[0].Code + (picked.Count > 1 ? ("（组 " + picked.Count.ToString(CultureInfo.InvariantCulture) + " 条）") : "");
            for (int k = 1; k < picked.Count; k++)
            {
                FillPreviewItem extra = new FillPreviewItem();
                extra.IsNameDriven = true; extra.Selected = true;
                extra.TemplateName = target.TemplateName; extra.TargetRow = target.TargetRow;
                extra.ItemNo = target.ItemNo; extra.OrderInItem = target.OrderInItem;
                extra.NeighborSourceQuotaSeq = target.NeighborSourceQuotaSeq;
                extra.QuotaCode = picked[k].Code; extra.ChosenQuotaSeq = picked[k].QuotaSeq;
                extra.GroupOrder = k; extra.SourceName = picked[k].Name; extra.Unit = picked[k].Unit; extra.TargetName = "";
                extra.QuantityText = target.QuantityText;
                extra.AlignNote = "组件框第 " + (k + 1).ToString(CultureInfo.InvariantCulture) + " 条";
                if (idx >= 0) preview.Insert(idx + k, extra); else preview.Add(extra);
            }
        }

        private sealed class QuotaPickerDialog : Form
        {
            private readonly TextBox txt = new TextBox();
            private readonly CheckedListBox lst = new CheckedListBox();
            private readonly Button ok = new Button();
            private readonly Label lblEmpty = new Label();
            private readonly List<ProjectQuota> all;
            private readonly List<ProjectQuota> suggested;
            public List<ProjectQuota> Picked = new List<ProjectQuota>();

            public QuotaPickerDialog(List<ProjectQuota> allq, List<ProjectQuota> sug)
            {
                all = allq ?? new List<ProjectQuota>();
                suggested = sug ?? new List<ProjectQuota>();
                Text = "选择定额（可多选=一量对多）"; StartPosition = FormStartPosition.CenterParent; ClientSize = new Size(520, 430);
                Label tip = new Label { Text = "搜索定额编号/名称；留空显示推荐；勾选多条=该工程量对应多条定额", AutoSize = false };
                tip.SetBounds(12, 10, 496, 18);
                txt.SetBounds(12, 32, 496, 23);
                lst.SetBounds(12, 62, 496, 320); lst.CheckOnClick = true;
                ok.Text = "确定"; ok.SetBounds(428, 392, 80, 27);
                lblEmpty.Text = "项目内没有匹配的定额。一期仅支持挂“项目内已有”的定额；\r\n请先在软件定额输入中录入一次该定额，再回来手挂。";
                lblEmpty.ForeColor = Color.Firebrick;
                lblEmpty.SetBounds(24, 150, 470, 60);
                lblEmpty.Visible = false;
                Controls.Add(tip); Controls.Add(txt); Controls.Add(lst); Controls.Add(ok);
                Controls.Add(lblEmpty);
                lblEmpty.BringToFront();
                txt.TextChanged += delegate { RefreshList(); };
                ok.Click += delegate {
                    Picked = lst.CheckedItems.Cast<ProjectQuota>().ToList();
                    if (Picked.Count == 0 && lst.SelectedItem is ProjectQuota) Picked.Add((ProjectQuota)lst.SelectedItem);
                    DialogResult = DialogResult.OK; Close();
                };
                RefreshList();
            }

            private void RefreshList()
            {
                string q = NormalizeMatchText(txt.Text);
                lst.Items.Clear();
                IEnumerable<ProjectQuota> src;
                if (q.Length == 0) src = suggested;
                else src = all.Select(o => new KeyValuePair<int, ProjectQuota>(
                        o.NormCode.IndexOf(q, StringComparison.Ordinal) >= 0 ? 1000 : MatchNameScore(q, o.NormName), o))
                    .Where(p => p.Key > 0).OrderByDescending(p => p.Key).Take(60).Select(p => p.Value);
                foreach (ProjectQuota o in src) lst.Items.Add(o);
                lblEmpty.Visible = lst.Items.Count == 0;
            }
        }

        // 回写：把本次名字驱动确认的"工程量名 -> 定额(可多条)"写进对应框 + 当前模版。
        private static void FeedbackNameMatches(string templateName, List<FillPreviewItem> written)
        {
            List<KeyValuePair<string, string>> pairs = new List<KeyValuePair<string, string>>();
            foreach (IGrouping<int, FillPreviewItem> g in written
                .Where(i => i.IsNameDriven && !String.IsNullOrWhiteSpace(i.QuotaCode))
                .GroupBy(i => i.TargetRow))
            {
                string name = g.Select(x => x.TargetName).FirstOrDefault(n => !String.IsNullOrWhiteSpace(n));
                if (String.IsNullOrWhiteSpace(name)) continue;
                foreach (FillPreviewItem it in g)
                {
                    pairs.Add(new KeyValuePair<string, string>(it.QuotaCode, name));
                }
            }
            RecordNameMatchesToMappingStore(pairs);

            try
            {
                FillTemplate t = LoadFillTemplate(templateName);
                if (t == null || !String.Equals(t.MatchBy, "name", StringComparison.OrdinalIgnoreCase)) return;
                bool changed = false;
                foreach (FillPreviewItem it in written.Where(i => i.IsNameDriven && !String.IsNullOrWhiteSpace(i.QuotaCode) && i.GroupOrder == 0))
                {
                    string nm = it.TargetName ?? "";
                    bool exists = t.Rows.Any(r =>
                        String.Equals(NormalizeMatchText(r.MatchName ?? ""), NormalizeMatchText(nm), StringComparison.Ordinal) &&
                        String.Equals(r.QuotaCode ?? "", it.QuotaCode ?? "", StringComparison.OrdinalIgnoreCase) &&
                        String.Equals(r.ItemNo ?? "", it.ItemNo ?? "", StringComparison.OrdinalIgnoreCase));
                    if (exists) continue;
                    t.Rows.Add(new FillTemplateRow { ItemNo = it.ItemNo, ItemName = it.ItemNo, QuotaCode = it.QuotaCode,
                        MatchName = nm, SourceName = nm, SourceQuotaSeq = it.ChosenQuotaSeq, OrderInItem = it.OrderInItem });
                    changed = true;
                }
                if (changed) SaveFillTemplate(t);
            }
            catch (Exception ex) { Log("FeedbackNameMatches template writeback failed: " + ex.Message); }
        }
    }
}
