using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;
using System.Xml;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        // 取数模式
        public enum FillMode { ColumnAnchor = 1, FixedColumn = 2 }

        // 模板里的一条定额
        public sealed class FillTemplateRow
        {
            public string ItemNo;        // 条目编号，如 0401-01
            public string ItemName;      // 条目名称（显示/核对）
            public string QuotaCode;     // 定额编号，含 *系数 后缀，原样
            public string Adjust;        // 定额调整整串（可空）
            public int OrderInItem;      // 条目内序号，保持插入先后
            public string SourceWorkbookPath;
            public string SourceSheet;   // 绑定时所在 sheet
            public string SourceExpr;    // 绑定表达式，如 "E5" 或 "E4+E5"
            public string SourceName;    // 源行项目名（供预览核对）
            public long SourceQuotaSeq;  // 源定额序号（写入时直接复制该行）
        }

        // 一份模板
        public sealed class FillTemplate
        {
            public string Name;
            public string Profession;
            public string SourceUnitNo;
            public string WorkbookPath;
            public List<FillTemplateRow> Rows = new List<FillTemplateRow>();
        }

        // 预览/写入用的一条结果
        public sealed class FillPreviewItem
        {
            public bool Selected = true;
            public string TemplateName;
            public string ItemNo;
            public string QuotaCode;
            public string Adjust;
            public string SourceName;
            public string TargetName;
            public string QuantityText;
            public string Status;
            public int OrderInItem;
            public long SourceQuotaSeq;  // 源定额序号（写入时直接复制该行）
        }

        private sealed class PreparedFillPreviewItem
        {
            public FillPreviewItem Item;
            public string WorkbookPath;
            public string SheetName;
            public string Expression;
        }

        private static string TemplateFillDir()
        {
            string dir = Path.Combine(FindSharedRecoQuotaDataDir(), "fill-templates");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            return dir;
        }

        private static string FindSharedRecoQuotaDataDir()
        {
            string localDataDir = FindRecoQuotaDataDir();
            string aiRoot = FindAiRoot(localDataDir);
            if (!String.IsNullOrEmpty(aiRoot))
            {
                string autoBudgetRoot = Path.Combine(aiRoot, "\u81ea\u52a8\u9884\u7b97");
                if (Directory.Exists(autoBudgetRoot))
                {
                    return Path.Combine(autoBudgetRoot, "RecoQuotaData");
                }
            }

            return localDataDir;
        }

        private static string FindAiRoot(string startPath)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(startPath);
                while (dir != null)
                {
                    if (Directory.Exists(Path.Combine(dir.FullName, "\u81ea\u52a8\u9884\u7b97")))
                    {
                        return dir.FullName;
                    }
                    dir = dir.Parent;
                }
            }
            catch { }

            return "";
        }

        private static List<string> templateFillDirsCache;
        private static DateTime templateFillDirsCacheUtc = DateTime.MinValue;
        private const int TemplateFillDirsCacheMs = 300000;

        private static List<string> TemplateFillDirs()
        {
            // 递归扫描 RecoQuotaData 目录较慢（尤其网络盘），结果缓存约 5 分钟。
            if (templateFillDirsCache != null &&
                (DateTime.UtcNow - templateFillDirsCacheUtc).TotalMilliseconds <= TemplateFillDirsCacheMs)
            {
                return templateFillDirsCache;
            }

            List<string> dirs = new List<string>();
            AddTemplateFillDir(dirs, TemplateFillDir(), true);
            AddTemplateFillDir(dirs, Path.Combine(FindRecoQuotaDataDir(), "fill-templates"), false);

            string aiRoot = FindAiRoot(FindRecoQuotaDataDir());
            if (!String.IsNullOrEmpty(aiRoot))
            {
                string[] roots = new[]
                {
                    Path.Combine(aiRoot, "\u81ea\u52a8\u9884\u7b97"),
                    Path.Combine(aiRoot, "\u81ea\u52a8\u9884\u7b97\u4e13\u7528\u7ebf"),
                    Path.Combine(aiRoot, "\u94c1\u8def\u5de5\u7a0b\u4e91\u8ba1\u4ef7\u7cfb\u7edf\u7f51\u7edc\u7248V1.0-\u5f90\u603b")
                };

                foreach (string root in roots)
                {
                    if (!Directory.Exists(root))
                    {
                        continue;
                    }
                    try
                    {
                        foreach (string dataDir in Directory.GetDirectories(root, "RecoQuotaData", SearchOption.AllDirectories))
                        {
                            AddTemplateFillDir(dirs, Path.Combine(dataDir, "fill-templates"), false);
                        }
                    }
                    catch { }
                }
            }

            templateFillDirsCache = dirs;
            templateFillDirsCacheUtc = DateTime.UtcNow;
            return dirs;
        }

        private static void AddTemplateFillDir(List<string> dirs, string dir, bool create)
        {
            if (String.IsNullOrWhiteSpace(dir))
            {
                return;
            }
            if (create && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            if (!Directory.Exists(dir))
            {
                return;
            }
            string full = Path.GetFullPath(dir).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            if (!dirs.Any(x => String.Equals(x, full, StringComparison.OrdinalIgnoreCase)))
            {
                dirs.Add(full);
            }
        }

        private static void SaveFillTemplate(FillTemplate template)
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = 1024 * 1024 * 16;
            string safe = String.Join("_", template.Name.Split(Path.GetInvalidFileNameChars()));
            string path = Path.Combine(TemplateFillDir(), safe + ".json");
            File.WriteAllText(path, serializer.Serialize(template), Encoding.UTF8);
        }

        private static List<string> ListFillTemplateNames()
        {
            return TemplateFillDirs()
                .SelectMany(dir => Directory.GetFiles(dir, "*.json"))
                .Select(p => Path.GetFileNameWithoutExtension(p))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static FillTemplate LoadFillTemplate(string name)
        {
            string path = TemplateFillDirs()
                .Select(dir => Path.Combine(dir, name + ".json"))
                .Where(File.Exists)
                .OrderByDescending(File.GetLastWriteTimeUtc)
                .FirstOrDefault();
            if (String.IsNullOrEmpty(path)) return null;
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = 1024 * 1024 * 16;
            return serializer.Deserialize<FillTemplate>(File.ReadAllText(path, Encoding.UTF8));
        }

        private static void DeleteFillTemplate(string name)
        {
            if (String.IsNullOrWhiteSpace(name)) return;
            foreach (string dir in TemplateFillDirs())
            {
                string path = Path.Combine(dir, name + ".json");
                if (File.Exists(path)) File.Delete(path);
            }
        }

        // 绑定库里出现过的 Excel 工作表名（去重），供"源sheet"下拉。
        private static List<string> ListBoundSheetNames(SqlConnection conn)
        {
            ExcelLinkStore store = LoadStore(conn);
            return store.Links
                .Where(l => l != null && !String.IsNullOrWhiteSpace(l.WorksheetName))
                .Select(l => l.WorksheetName.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(s => s, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        // 从章节树顶层节点文字解析当前单元号（如 "总概算---[_ZGS_02(南通西)]" -> "_ZGS_02"）。
        private static string GetCurrentUnitNo(Form mainForm)
        {
            try
            {
                TreeView tree = GetField<TreeView>(mainForm, "Tv_tree");
                if (tree == null || tree.Nodes.Count == 0) return "";
                string text = tree.Nodes[0].Text ?? "";
                int i = text.IndexOf("_ZGS_", StringComparison.OrdinalIgnoreCase);
                if (i < 0) return "";
                int j = i + 5;
                while (j < text.Length && (Char.IsLetterOrDigit(text[j]) || text[j] == '_')) j++;
                return text.Substring(i, j - i);
            }
            catch { return ""; }
        }

        // 生成模板（跟着绑定走）：收集【源单元 unitNo】里、绑定到【源 sheet sourceSheet】的定额，
        // 自动跨各条目。按 sheet 过滤可排除绑定库里其它专业(其它 sheet，如站场)的历史绑定。
        private static FillTemplate BuildFillTemplateFromBindings(
            Form mainForm, SqlConnection conn, string templateName, string unitNo, string sourceSheet)
        {
            FillTemplate template = new FillTemplate { Name = templateName, SourceUnitNo = unitNo };

            // 1) 选出本专业的绑定：只按【源 sheet】过滤（按专业隔离）。
            //    单元范围由下面 (总概算序号=@zgs) 的定额查询 + byId 自动收口：
            //    只有本单元的定额会进 byId，别的单元绑定到的定额序号查不到、被跳过。
            ExcelLinkStore store = LoadStore(conn);
            string sheet = (sourceSheet ?? "").Trim();
            List<ExcelQuotaLink> picked = store.Links
                .Where(l => l != null
                    && String.Equals((l.WorksheetName ?? "").Trim(), sheet, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (picked.Count == 0) return template;

            // 2) 解析 源单元 -> 总概算序号(数字)。定额输入表用 总概算序号 关联单元，不是 _ZGS_编号。
            string srcLabel;
            long zgsSeq = ResolveAgentUnitIdSimple(conn, unitNo, out srcLabel);
            if (zgsSeq <= 0) return template; // 找不到该单元

            // 一次查出本单元全部定额的 条目/编号/调整/项目名/顺号，按 定额序号 建索引。
            Dictionary<long, FillTemplateRow> byId = new Dictionary<long, FillTemplateRow>();
            Dictionary<long, int> shunById = new Dictionary<long, int>();
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    "select DE.定额序号, ZJ.条目编号, DE.定额编号, " +
                    "cast(DE.定额调整 as nvarchar(max)), DE.顺号, DE.工程或费用项目名称 " +
                    "from 定额输入 DE inner join 章节表 ZJ on DE.条目序号=ZJ.条目序号 " +
                    "where DE.总概算序号=@zgs";
                cmd.Parameters.AddWithValue("@zgs", zgsSeq);
                using (SqlDataReader r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        long id = Convert.ToInt64(r.GetValue(0), CultureInfo.InvariantCulture);
                        string itemNo = Convert.ToString(r.GetValue(1)).Trim();
                        byId[id] = new FillTemplateRow
                        {
                            ItemNo = itemNo,
                            ItemName = itemNo,
                            QuotaCode = Convert.ToString(r.GetValue(2)).Trim(),
                            Adjust = r.IsDBNull(3) ? "" : Convert.ToString(r.GetValue(3)).Trim(),
                            SourceName = r.IsDBNull(5) ? "" : Convert.ToString(r.GetValue(5)).Trim(),
                            SourceQuotaSeq = id
                        };
                        int shun;
                        shunById[id] = Int32.TryParse(Convert.ToString(r.GetValue(4)), NumberStyles.Integer, CultureInfo.InvariantCulture, out shun) ? shun : 0;
                    }
                }
            }

            // 3) 对选中的绑定，取出对应定额行，填取数引用（保留 顺号 供排序）。
            List<KeyValuePair<int, FillTemplateRow>> collected = new List<KeyValuePair<int, FillTemplateRow>>();
            foreach (ExcelQuotaLink link in picked)
            {
                FillTemplateRow row;
                if (!byId.TryGetValue(link.QuotaSequence, out row)) continue; // 定额已删/不在本单元
                row.SourceWorkbookPath = link.ExcelPath;
                row.SourceSheet = link.WorksheetName;
                row.SourceExpr = String.IsNullOrEmpty(link.Expression) ? link.CellAddress : link.Expression;
                int shun; shunById.TryGetValue(link.QuotaSequence, out shun);
                collected.Add(new KeyValuePair<int, FillTemplateRow>(shun, row));
            }

            string workbookPath = PickTemplateWorkbookPath(collected.Select(p => p.Value.SourceWorkbookPath));
            if (!String.IsNullOrWhiteSpace(workbookPath))
            {
                int before = collected.Count;
                collected = collected
                    .Where(p => SameTemplateWorkbookPath(p.Value.SourceWorkbookPath, workbookPath))
                    .ToList();
                template.WorkbookPath = workbookPath;
                if (collected.Count < before)
                {
                    Log("Template fill skipped bindings from other workbooks for sheet " + sheet + ".");
                }
            }

            // 4) 按 条目编号 分组，组内按 顺号 排序，分配 OrderInItem。
            foreach (IGrouping<string, KeyValuePair<int, FillTemplateRow>> grp in
                     collected.GroupBy(p => p.Value.ItemNo).OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase))
            {
                int order = 0;
                foreach (KeyValuePair<int, FillTemplateRow> pair in grp.OrderBy(p => p.Key))
                {
                    pair.Value.OrderInItem = order++;
                    template.Rows.Add(pair.Value);
                }
            }
            return template;
        }

        private static string PickTemplateWorkbookPath(IEnumerable<string> paths)
        {
            return (paths ?? new string[0])
                .Where(p => !String.IsNullOrWhiteSpace(p))
                .GroupBy(NormalizeTemplateWorkbookPath, StringComparer.OrdinalIgnoreCase)
                .OrderByDescending(g => g.Count())
                .Select(g => g.First())
                .FirstOrDefault() ?? "";
        }

        private static bool SameTemplateWorkbookPath(string left, string right)
        {
            return String.Equals(NormalizeTemplateWorkbookPath(left), NormalizeTemplateWorkbookPath(right), StringComparison.OrdinalIgnoreCase);
        }

        private static string NormalizeTemplateWorkbookPath(string path)
        {
            try
            {
                return Path.GetFullPath(path ?? "");
            }
            catch
            {
                return path ?? "";
            }
        }

        private static string GetTemplateRowWorkbookPath(FillTemplate template, FillTemplateRow row)
        {
            string path = row == null ? "" : row.SourceWorkbookPath;
            if (String.IsNullOrWhiteSpace(path) && template != null)
            {
                path = template.WorkbookPath;
            }

            return path ?? "";
        }

        // 把表达式里每个单元格的列字母替换为 targetColumn（行号不变）。
        // 例 "E5" + targetCol "F" -> "F5"；"E4+E5" -> "F4+F5"。
        private static string RetargetExprColumn(string expr, string targetColumn)
        {
            if (String.IsNullOrWhiteSpace(expr)) return expr;
            List<string> cells = ExtractCellAddressesFromExpression(expr);
            string result = expr.ToUpperInvariant();
            foreach (string cell in cells.OrderByDescending(c => c.Length))
            {
                CellRef cr;
                if (!TryParseCellAddress(cell, out cr)) continue;
                string replaced = targetColumn.ToUpperInvariant() + (cr.Row).ToString(CultureInfo.InvariantCulture);
                result = result.Replace(cell, replaced);
            }
            return result;
        }

        // 模式一：列锚点。行号不变、列换成目标列，从目标 sheet 取数。
        private static List<FillPreviewItem> BuildPreview_ColumnAnchor(
            FillTemplate template, string targetSheet, string targetColumn)
        {
            return BuildPreview(template, row =>
                new KeyValuePair<string, string>(targetSheet, RetargetExprColumn(row.SourceExpr, targetColumn)));
        }

        // 模式二：固定绑定列。直接读模板记录的原 sheet/单元格（用户已把目标单元数量粘进该列）。
        private static List<FillPreviewItem> BuildPreview_FixedColumn(FillTemplate template)
        {
            return BuildPreview(template, row =>
                new KeyValuePair<string, string>(row.SourceSheet, row.SourceExpr));
        }

        // 两种取数模式共用：resolver 给出每行的 (目标sheet, 取数表达式)，其余流程一致。
        private static List<FillPreviewItem> BuildPreview(
            FillTemplate template, Func<FillTemplateRow, KeyValuePair<string, string>> resolver)
        {
            List<FillPreviewItem> items = new List<FillPreviewItem>();
            List<PreparedFillPreviewItem> prepared = new List<PreparedFillPreviewItem>();
            List<ExcelQuotaLink> readLinks = new List<ExcelQuotaLink>();
            Dictionary<string, HashSet<int>> hiddenColumnCache = new Dictionary<string, HashSet<int>>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, List<ExcelMergedRegion>> mergedRegionCache = new Dictionary<string, List<ExcelMergedRegion>>(StringComparer.OrdinalIgnoreCase);
            foreach (FillTemplateRow row in template.Rows)
            {
                FillPreviewItem item = new FillPreviewItem
                {
                    TemplateName = template.Name,
                    ItemNo = row.ItemNo, QuotaCode = row.QuotaCode, Adjust = row.Adjust,
                    SourceName = row.SourceName, OrderInItem = row.OrderInItem,
                    SourceQuotaSeq = row.SourceQuotaSeq
                };

                if (String.IsNullOrWhiteSpace(row.SourceExpr))
                {
                    item.Status = "模板未记录取数位置"; item.Selected = false; items.Add(item); continue;
                }

                KeyValuePair<string, string> tgt = resolver(row);
                string sheet = tgt.Key, expr = tgt.Value;
                string workbook = GetTemplateRowWorkbookPath(template, row);
                if (String.IsNullOrWhiteSpace(workbook))
                {
                    item.Status = "模板未记录 Excel 文件"; item.Selected = false; items.Add(item); continue;
                }

                List<ExcelMergedRegion> rowMergedRegions = GetSavedMergedRegionsCached(workbook, sheet, mergedRegionCache);
                expr = NormalizeExpressionMergedAnchors(expr, rowMergedRegions);

                PreparedFillPreviewItem preparedItem = new PreparedFillPreviewItem();
                preparedItem.Item = item;
                preparedItem.WorkbookPath = workbook;
                preparedItem.SheetName = sheet;
                preparedItem.Expression = expr;
                string first = ExtractFirstCellAddress(expr);
                CellRef firstCell;
                if (!String.IsNullOrEmpty(first) && TryParseCellAddress(first, out firstCell))
                {
                    readLinks.Add(new ExcelQuotaLink { ExcelPath = workbook, WorksheetName = sheet, CellAddress = first, Expression = expr });

                    HashSet<int> hiddenColumns = GetSavedHiddenColumns(workbook, sheet, hiddenColumnCache);
                    for (int col = 1; col < firstCell.Column; col++)
                    {
                        if (hiddenColumns.Contains(col))
                        {
                            continue;
                        }

                        string address = ColumnNumberToName(col) + firstCell.Row.ToString(CultureInfo.InvariantCulture);
                        readLinks.Add(new ExcelQuotaLink { ExcelPath = workbook, WorksheetName = sheet, CellAddress = address, Expression = address });

                        // 合并单元格只有锚点格有值：把所在区域的锚点也加入批量读取，供名称回填。
                        ExcelMergedRegion region = FindExcelMergedRegionAt(rowMergedRegions, col, firstCell.Row);
                        if (region != null)
                        {
                            string anchorAddress = ColumnNumberToName(region.FirstColumn) + region.FirstRow.ToString(CultureInfo.InvariantCulture);
                            readLinks.Add(new ExcelQuotaLink { ExcelPath = workbook, WorksheetName = sheet, CellAddress = anchorAddress, Expression = anchorAddress });
                        }
                    }
                }
                else
                {
                    readLinks.Add(new ExcelQuotaLink { ExcelPath = workbook, WorksheetName = sheet, Expression = expr });
                }

                prepared.Add(preparedItem);
                items.Add(item);
            }

            ExcelSyncReadContext readContext = new ExcelSyncReadContext(readLinks);
            foreach (PreparedFillPreviewItem preparedItem in prepared)
            {
                FillPreviewItem item = preparedItem.Item;
                string display; decimal qty; string err;
                if (!TryEvaluateWorkbookExpression(readContext, preparedItem.WorkbookPath, preparedItem.SheetName, preparedItem.Expression, out display, out qty, out err, true))
                {
                    item.Status = "取数失败：" + err; item.Selected = false; continue;
                }

                item.QuantityText = display;
                item.TargetName = ReadRowNameAt(preparedItem.WorkbookPath, preparedItem.SheetName, preparedItem.Expression, hiddenColumnCache, mergedRegionCache, readContext);
                if (qty == 0m) { item.Status = "数量为0"; item.Selected = false; }
            }

            return items;
        }

        private const int RowNameMaxLength = 20;
        private const int RowNameMaxFragments = 3;

        // 读某表达式首个单元格所在行的名称（该格列前的非数字文本拼接），仅供人工核对。
        // 空格只从包含它的真实合并区域锚点回填；就近优先，10字/3段封顶，同一合并区域只拼一次。
        private static string ReadRowNameAt(string workbook, string sheet, string expr, Dictionary<string, HashSet<int>> hiddenColumnCache, Dictionary<string, List<ExcelMergedRegion>> mergedRegionCache, ExcelSyncReadContext readContext)
        {
            try
            {
                HashSet<int> hiddenColumns = GetSavedHiddenColumns(workbook, sheet, hiddenColumnCache);
                List<ExcelMergedRegion> mergedRegions = GetSavedMergedRegionsCached(workbook, sheet, mergedRegionCache);
                expr = NormalizeExpressionMergedAnchors(expr, mergedRegions);
                string first = ExtractFirstCellAddress(expr);
                CellRef cr;
                if (String.IsNullOrEmpty(first) || !TryParseCellAddress(first, out cr)) return "";
                List<KeyValuePair<int, string>> fragments = new List<KeyValuePair<int, string>>();
                HashSet<string> sourceKeys = new HashSet<string>(StringComparer.Ordinal);
                for (int col = 1; col < cr.Column; col++)
                {
                    if (hiddenColumns.Contains(col))
                    {
                        continue;
                    }

                    string sourceKey;
                    int fragmentColumn;
                    string text = ReadRowNameCellText(workbook, sheet, col, cr.Row, readContext, mergedRegions, out sourceKey, out fragmentColumn);
                    if (String.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    if (sourceKeys.Add(sourceKey))
                    {
                        fragments.Add(new KeyValuePair<int, string>(fragmentColumn, text.Trim()));
                    }
                }

                List<KeyValuePair<int, string>> kept = new List<KeyValuePair<int, string>>();
                int totalLength = 0;
                foreach (KeyValuePair<int, string> fragment in fragments
                    .OrderBy(f => Math.Abs(f.Key - cr.Column))
                    .ThenBy(f => f.Key))
                {
                    if (kept.Count >= RowNameMaxFragments)
                    {
                        break;
                    }

                    if (kept.Count > 0 && totalLength + fragment.Value.Length > RowNameMaxLength)
                    {
                        break;
                    }

                    kept.Add(fragment);
                    totalLength += fragment.Value.Length;
                }

                return String.Join(" ", kept.OrderBy(f => f.Key).Select(f => f.Value).ToArray()).Trim();
            }
            catch { return ""; }
        }

        // 本格文字；空格时只从包含它的真实合并区域锚点回填，数字不算名称。
        // fragmentColumn 统一取合并区域锚点列，保证名称取舍顺序稳定。
        private static string ReadRowNameCellText(string workbook, string sheet, int col, int row, ExcelSyncReadContext readContext, List<ExcelMergedRegion> mergedRegions, out string sourceKey, out int fragmentColumn)
        {
            sourceKey = col.ToString(CultureInfo.InvariantCulture) + "|" + row.ToString(CultureInfo.InvariantCulture);
            fragmentColumn = col;
            if (readContext == null)
            {
                return "";
            }

            string val;
            string readError;
            string addr = ColumnNumberToName(col) + row.ToString(CultureInfo.InvariantCulture);
            ExcelMergedRegion region = FindExcelMergedRegionAt(mergedRegions, col, row);
            if (readContext.TryReadWorkbookCellValue(workbook, sheet, addr, out val, out readError) && !String.IsNullOrWhiteSpace(val))
            {
                decimal parsed;
                string parseError;
                if (TryEvaluateDecimal(val, out parsed, out parseError))
                {
                    return "";
                }

                if (region != null)
                {
                    sourceKey = BuildExcelMergedRegionKey(region);
                    fragmentColumn = region.FirstColumn;
                }

                return val;
            }

            if (region == null)
            {
                return "";
            }

            string anchorAddr = ColumnNumberToName(region.FirstColumn) + region.FirstRow.ToString(CultureInfo.InvariantCulture);
            if (!readContext.TryReadWorkbookCellValue(workbook, sheet, anchorAddr, out val, out readError) || String.IsNullOrWhiteSpace(val))
            {
                return "";
            }

            decimal anchorParsed;
            string anchorParseError;
            if (TryEvaluateDecimal(val, out anchorParsed, out anchorParseError))
            {
                return "";
            }

            sourceKey = BuildExcelMergedRegionKey(region);
            fragmentColumn = region.FirstColumn;
            return val;
        }

        private static List<ExcelMergedRegion> GetSavedMergedRegionsCached(string workbook, string sheet, Dictionary<string, List<ExcelMergedRegion>> cache)
        {
            string key = (workbook ?? "") + "|" + (sheet ?? "");
            List<ExcelMergedRegion> regions;
            if (cache != null && cache.TryGetValue(key, out regions))
            {
                return regions;
            }

            regions = ReadExcelMergedRegions(workbook, sheet);
            if (cache != null)
            {
                cache[key] = regions;
            }

            return regions;
        }

        // —— 条目树共用助手（模板铺量 / 自动匹配窗口）——
        private static Dictionary<string, string> LoadChapterNameMap(Form mainForm)
        {
            Dictionary<string, string> names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using (SqlConnection conn = AgentCreateWorkConnection(mainForm))
                {
                    EnsureOpen(conn);
                    using (SqlCommand cmd = conn.CreateCommand())
                    {
                        cmd.CommandText = "select 条目编号, 工程或费用项目名称 from 章节表";
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                if (reader.IsDBNull(0))
                                {
                                    continue;
                                }

                                string code = Convert.ToString(reader.GetValue(0)).Trim();
                                if (code.Length == 0 || names.ContainsKey(code))
                                {
                                    continue;
                                }

                                names[code] = reader.IsDBNull(1) ? "" : Convert.ToString(reader.GetValue(1)).Trim();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("Load chapter names failed: " + ex.Message);
            }

            return names;
        }

        // 与软件左侧章节树一致：两位纯数字章（01）显示"一、"，四位及以上（0101）显示"01."；
        // 带横杠的下级条目名称自带序号，不再重复。
        private static string ChapterTreeDisplayName(Dictionary<string, string> chapterNames, string code)
        {
            string name;
            if (chapterNames == null || !chapterNames.TryGetValue(code, out name) || String.IsNullOrEmpty(name))
            {
                return code;
            }

            if (code.All(Char.IsDigit))
            {
                if (code.Length == 2)
                {
                    int value;
                    if (Int32.TryParse(code, out value))
                    {
                        return ToChineseOrdinal(value) + "、" + name;
                    }
                }
                else if (code.Length >= 4)
                {
                    return code.Substring(code.Length - 2) + "." + name;
                }
            }

            return name;
        }

        private static string ToChineseOrdinal(int value)
        {
            string[] digits = { "零", "一", "二", "三", "四", "五", "六", "七", "八", "九" };
            if (value <= 0 || value > 99)
            {
                return value.ToString();
            }

            if (value < 10)
            {
                return digits[value];
            }

            int tens = value / 10;
            int ones = value % 10;
            return (tens == 1 ? "" : digits[tens]) + "十" + (ones == 0 ? "" : digits[ones]);
        }

        private static bool IsChapterTreeCode(string code)
        {
            return !String.IsNullOrEmpty(code) && code.Length >= 2 && Char.IsDigit(code[0]) && Char.IsDigit(code[1]);
        }

        // itemNo 是否属于编号 code 的条目（本身或下级）。
        // 下级判定：带横杠段（0101 -> 0101-04），或纯数字编号续位（01 -> 0101）。
        private static bool IsItemNoUnderChapter(string itemNo, string code)
        {
            itemNo = (itemNo ?? "").Trim();
            code = (code ?? "").Trim();
            if (itemNo.Length == 0 || code.Length == 0)
            {
                return false;
            }

            if (String.Equals(itemNo, code, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (!itemNo.StartsWith(code, StringComparison.OrdinalIgnoreCase) || itemNo.Length <= code.Length)
            {
                return false;
            }

            char next = itemNo[code.Length];
            if (next == '-')
            {
                return true;
            }

            return Char.IsDigit(next) && code.All(Char.IsDigit);
        }

        // 条目编号的祖先链：章节表里存在、是 itemNo 前缀、且编号以两位数字开头的编号，按长度升序。
        private static List<string> BuildChapterChain(Dictionary<string, string> chapterNames, string itemNo)
        {
            List<string> chain = (chapterNames == null ? new List<string>() : chapterNames.Keys
                .Where(code => IsChapterTreeCode(code) && IsItemNoUnderChapter(itemNo, code))
                .OrderBy(code => code.Length)
                .ThenBy(code => code, StringComparer.OrdinalIgnoreCase)
                .ToList());
            if (!chain.Any(code => String.Equals(code, itemNo, StringComparison.OrdinalIgnoreCase)))
            {
                chain.Add(itemNo);
            }

            return chain;
        }

        private static void SetTreeChildrenChecked(TreeNode node, bool value)
        {
            foreach (TreeNode child in node.Nodes)
            {
                child.Checked = value;
                SetTreeChildrenChecked(child, value);
            }
        }

        private static HashSet<int> GetSavedHiddenColumns(string workbook, string sheet, Dictionary<string, HashSet<int>> cache)
        {
            string key = (workbook ?? "") + "|" + (sheet ?? "");
            HashSet<int> hidden;
            if (cache != null && cache.TryGetValue(key, out hidden))
            {
                return hidden;
            }

            hidden = new HashSet<int>();
            try
            {
                using (ZipArchive archive = OpenZipArchiveShared(workbook))
                {
                    string sheetPath = ResolveSheetPath(archive, sheet);
                    if (!String.IsNullOrEmpty(sheetPath))
                    {
                        ZipArchiveEntry sheetEntry = archive.GetEntry(sheetPath);
                        if (sheetEntry != null)
                        {
                            using (Stream stream = sheetEntry.Open())
                            using (XmlReader reader = XmlReader.Create(stream))
                            {
                                while (reader.Read())
                                {
                                    if (reader.NodeType != XmlNodeType.Element || reader.LocalName != "col")
                                    {
                                        continue;
                                    }

                                    string hiddenText = reader.GetAttribute("hidden");
                                    if (!String.Equals(hiddenText, "1", StringComparison.OrdinalIgnoreCase) &&
                                        !String.Equals(hiddenText, "true", StringComparison.OrdinalIgnoreCase))
                                    {
                                        continue;
                                    }

                                    int min;
                                    int max;
                                    if (!Int32.TryParse(reader.GetAttribute("min"), NumberStyles.Integer, CultureInfo.InvariantCulture, out min) ||
                                        !Int32.TryParse(reader.GetAttribute("max"), NumberStyles.Integer, CultureInfo.InvariantCulture, out max))
                                    {
                                        continue;
                                    }

                                    if (min > max)
                                    {
                                        int tmp = min;
                                        min = max;
                                        max = tmp;
                                    }

                                    max = Math.Min(max, 16384);
                                    for (int col = Math.Max(1, min); col <= max; col++)
                                    {
                                        hidden.Add(col);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("Read saved hidden columns failed: " + ex.Message);
            }

            if (cache != null)
            {
                cache[key] = hidden;
            }

            return hidden;
        }

        // 写入：把选中预览项对应的源定额行，直接复制到【目标单元】的对应条目（条目序号全局共享，原样保留），
        // 改 总概算序号/顺号/工程数量、丢弃旧 定额序号(新建标识)。不走界面树。
        private static string ApplyFill(Form mainForm, string targetUnitNo, List<FillPreviewItem> items)
        {
            List<FillPreviewItem> selected = items
                .Where(i => i.Selected &&
                    (String.IsNullOrEmpty(i.Status) || String.Equals(i.Status, "\u6570\u91cf\u4e3a0", StringComparison.Ordinal)) &&
                    i.SourceQuotaSeq > 0)
                .OrderBy(i => i.ItemNo, StringComparer.OrdinalIgnoreCase)
                .ThenBy(i => i.OrderInItem)
                .ToList();
            if (selected.Count == 0) return "没有可写入的行。";

            using (SqlConnection conn = AgentCreateWorkConnection(mainForm))
            {
                string targetLabel;
                long targetSeq = ResolveAgentUnitIdSimple(conn, targetUnitNo, out targetLabel);
                if (targetSeq <= 0) return "找不到目标单元：" + targetUnitNo + "（请填 _ZGS_编号 或单元名称）。";

                AgentUndoRecord undo = new AgentUndoRecord { Summary = "模板铺量 -> 单元 " + targetUnitNo, Time = DateTime.Now };
                StringBuilder msg = new StringBuilder();
                int inserted = 0, markerRows = 0, skipped = 0;
                // 每个 (条目序号) 的下一个顺号，写入时递增。
                Dictionary<long, int> nextShun = new Dictionary<long, int>();
                HashSet<long> markerInserted = new HashSet<long>();

                using (SqlTransaction transaction = conn.BeginTransaction())
                {
                    try
                    {
                        foreach (FillPreviewItem item in selected)
                        {
                            Dictionary<string, object> row = LoadTemplateFullRow(conn, transaction, item.SourceQuotaSeq);
                            if (row == null) { skipped++; continue; }

                            // 条目序号(全局)保持不变 -> 落到目标单元的同一条目。
                            long itemSeq = Convert.ToInt64(row["条目序号"], CultureInfo.InvariantCulture);

                            int shun;
                            if (!nextShun.TryGetValue(itemSeq, out shun))
                            {
                                shun = GetMaxShun(conn, transaction, targetSeq, itemSeq) + 1;
                            }

                            if (!markerInserted.Contains(itemSeq))
                            {
                                Dictionary<string, object> marker = new Dictionary<string, object>(row);
                                ApplyTemplateFillMarkerFields(marker, targetSeq, shun, item.TemplateName);
                                marker.Remove("定额序号"); // 让数据库分配新标识

                                long markerId = InsertQuotaRowReturnId(conn, transaction, marker);
                                if (markerId > 0)
                                {
                                    undo.Rows.Add(new AgentUndoRow { Kind = "I", QuotaSequence = markerId });
                                    inserted++;
                                    markerRows++;
                                    markerInserted.Add(itemSeq);
                                    shun++;
                                }
                                else
                                {
                                    skipped++;
                                }
                            }

                            // 数量：用预览取到的目标工程量。
                            decimal qty; string qErr;
                            bool okQty = TryEvaluateDecimal(item.QuantityText, out qty, out qErr);

                            row["总概算序号"] = targetSeq;
                            row["顺号"] = shun;
                            row["工程数量输入"] = (object)(item.QuantityText ?? "");
                            row["工程数量"] = okQty ? (object)qty : DBNull.Value;
                            row.Remove("定额序号"); // 让数据库分配新标识

                            long newId = InsertQuotaRowReturnId(conn, transaction, row);
                            if (newId > 0)
                            {
                                undo.Rows.Add(new AgentUndoRow { Kind = "I", QuotaSequence = newId });
                                inserted++;
                                nextShun[itemSeq] = shun + 1;
                            }
                            else { skipped++; }
                        }

                        transaction.Commit();
                    }
                    catch
                    {
                        transaction.Rollback();
                        throw;
                    }
                }

                if (undo.Rows.Count > 0)
                {
                    GetAgentUndoStack(mainForm).Add(undo);
                    GetAgentRedoStack(mainForm).Clear();
                }
                RefreshCurrentQuotaGrid(mainForm);

                msg.Append("已向单元 ").Append(targetUnitNo).Append(" 追加定额 ")
                   .Append((inserted - markerRows).ToString(CultureInfo.InvariantCulture)).Append(" 条");
                if (markerRows > 0) msg.Append("，标记 ").Append(markerRows.ToString(CultureInfo.InvariantCulture)).Append(" 条");
                if (skipped > 0) msg.Append("，跳过 ").Append(skipped.ToString(CultureInfo.InvariantCulture)).Append(" 条");
                msg.Append("。请在软件点一次“计算”刷新单价/合价与汇总。");
                return msg.ToString();
            }
        }

        private static void ApplyTemplateFillMarkerFields(Dictionary<string, object> row, long targetSeq, int shun, string templateName)
        {
            row["总概算序号"] = targetSeq;
            row["顺号"] = shun;
            SetIfPresent(row, "定额编号", "-");
            SetIfPresent(row, "工程或费用项目名称", BuildTemplateFillMarkerText(templateName));
            SetIfPresent(row, "单位", "");
            SetIfPresent(row, "工程数量输入", "");
            SetIfPresent(row, "工程数量", 0m);
            SetIfPresent(row, "定额调整", "");

            foreach (string name in new string[] { "单价", "合价", "基价", "工费", "料费", "机费", "单重", "合重", "人工费", "材料费", "机械费", "设备费", "主材费", "价差" })
            {
                SetIfPresent(row, name, 0m);
            }
        }

        private static string BuildTemplateFillMarkerText(string templateName)
        {
            string name = (templateName ?? "").Trim();
            return String.IsNullOrEmpty(name) ? "AI推送" : "AI推送：" + name;
        }

        private static void SetIfPresent(Dictionary<string, object> row, string key, object value)
        {
            if (row.ContainsKey(key))
            {
                row[key] = value;
            }
        }

        private static Dictionary<string, object> LoadTemplateFullRow(SqlConnection conn, SqlTransaction transaction, long quotaSequence)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = "select * from 定额输入 where 定额序号=@id";
                cmd.Parameters.AddWithValue("@id", quotaSequence);
                using (SqlDataAdapter adapter = new SqlDataAdapter(cmd))
                {
                    DataTable table = new DataTable();
                    adapter.Fill(table);
                    if (table.Rows.Count == 0)
                    {
                        return null;
                    }

                    Dictionary<string, object> values = new Dictionary<string, object>();
                    foreach (DataColumn column in table.Columns)
                    {
                        values[column.ColumnName] = table.Rows[0][column];
                    }

                    return values;
                }
            }
        }

        // 目标单元某条目下当前最大顺号(无则0)。
        private static int GetMaxShun(SqlConnection conn, long zgsSeq, long itemSeq)
        {
            return GetMaxShun(conn, null, zgsSeq, itemSeq);
        }

        private static int GetMaxShun(SqlConnection conn, SqlTransaction transaction, long zgsSeq, long itemSeq)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = "select isnull(max(顺号),0) from 定额输入 where 总概算序号=@z and 条目序号=@t";
                cmd.Parameters.AddWithValue("@z", zgsSeq);
                cmd.Parameters.AddWithValue("@t", itemSeq);
                object o = cmd.ExecuteScalar();
                return (o == null || o == DBNull.Value) ? 0 : Convert.ToInt32(o, CultureInfo.InvariantCulture);
            }
        }

        // 插入一行 定额输入(不含 定额序号)，返回新分配的 定额序号。
        private static long InsertQuotaRowReturnId(SqlConnection conn, Dictionary<string, object> values)
        {
            return InsertQuotaRowReturnId(conn, null, values);
        }

        private static long InsertQuotaRowReturnId(SqlConnection conn, SqlTransaction transaction, Dictionary<string, object> values)
        {
            List<string> cols = values.Keys.ToList();
            StringBuilder sql = new StringBuilder();
            sql.Append("insert into 定额输入 (")
               .Append(String.Join(", ", cols.Select(c => "[" + c + "]").ToArray()))
               .Append(") values (")
               .Append(String.Join(", ", cols.Select((c, i) => "@p" + i.ToString(CultureInfo.InvariantCulture)).ToArray()))
               .Append("); select cast(scope_identity() as bigint);");
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.Transaction = transaction;
                cmd.CommandText = sql.ToString();
                for (int i = 0; i < cols.Count; i++)
                    cmd.Parameters.AddWithValue("@p" + i.ToString(CultureInfo.InvariantCulture), values[cols[i]] ?? DBNull.Value);
                object o = cmd.ExecuteScalar();
                return (o == null || o == DBNull.Value) ? 0 : Convert.ToInt64(o, CultureInfo.InvariantCulture);
            }
        }
    }
}
