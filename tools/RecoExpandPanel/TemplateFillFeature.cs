using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;

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

        private static string TemplateFillDir()
        {
            string dir = Path.Combine(FindRecoQuotaDataDir(), "fill-templates");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
            return dir;
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
            return Directory.GetFiles(TemplateFillDir(), "*.json")
                .Select(p => Path.GetFileNameWithoutExtension(p))
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        private static FillTemplate LoadFillTemplate(string name)
        {
            string path = Path.Combine(TemplateFillDir(), name + ".json");
            if (!File.Exists(path)) return null;
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = 1024 * 1024 * 16;
            return serializer.Deserialize<FillTemplate>(File.ReadAllText(path, Encoding.UTF8));
        }

        private static void DeleteFillTemplate(string name)
        {
            if (String.IsNullOrWhiteSpace(name)) return;
            string path = Path.Combine(TemplateFillDir(), name + ".json");
            if (File.Exists(path)) File.Delete(path);
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
            template.WorkbookPath = picked[0].ExcelPath;

            // 2) 解析 源单元 -> 总概算序号(数字)。定额输入表用 总概算序号 关联单元，不是 _ZGS_编号。
            long zgsSeq = ResolveUnitSeq(conn, unitNo);
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
                row.SourceSheet = link.WorksheetName;
                row.SourceExpr = String.IsNullOrEmpty(link.Expression) ? link.CellAddress : link.Expression;
                int shun; shunById.TryGetValue(link.QuotaSequence, out shun);
                collected.Add(new KeyValuePair<int, FillTemplateRow>(shun, row));
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

        // 解析 源单元 输入 -> 总概算序号(数字)。支持：直接数字 / _ZGS_编号(总概算编号) / 编制范围名称。
        private static long ResolveUnitSeq(SqlConnection conn, string input)
        {
            string v = (input ?? "").Trim();
            if (v.Length == 0) return 0;
            using (SqlCommand cmd = conn.CreateCommand())
            {
                long n;
                if (Int64.TryParse(v, NumberStyles.Integer, CultureInfo.InvariantCulture, out n))
                {
                    cmd.CommandText = "select top 1 总概算序号 from 总概算信息 where 总概算序号=@v";
                    cmd.Parameters.AddWithValue("@v", n);
                }
                else
                {
                    cmd.CommandText = "select top 1 总概算序号 from 总概算信息 where 总概算编号=@v or 编制范围=@v";
                    cmd.Parameters.AddWithValue("@v", v);
                }
                object o = cmd.ExecuteScalar();
                return (o == null || o == DBNull.Value) ? 0 : Convert.ToInt64(o, CultureInfo.InvariantCulture);
            }
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

        // 模式一：按目标 sheet+列，对每条模板行求工程量。
        private static List<FillPreviewItem> BuildPreview_ColumnAnchor(
            FillTemplate template, string targetSheet, string targetColumn)
        {
            List<FillPreviewItem> items = new List<FillPreviewItem>();
            foreach (FillTemplateRow row in template.Rows)
            {
                FillPreviewItem item = new FillPreviewItem
                {
                    ItemNo = row.ItemNo, QuotaCode = row.QuotaCode, Adjust = row.Adjust,
                    SourceName = row.SourceName, OrderInItem = row.OrderInItem,
                    SourceQuotaSeq = row.SourceQuotaSeq
                };

                if (String.IsNullOrWhiteSpace(row.SourceExpr))
                {
                    item.Status = "模板未记录取数位置"; item.Selected = false; items.Add(item); continue;
                }

                string expr = RetargetExprColumn(row.SourceExpr, targetColumn);
                string display; decimal qty; string err;
                if (!TryEvaluateWorkbookExpression(template.WorkbookPath, targetSheet, expr, out display, out qty, out err))
                {
                    item.Status = "取数失败：" + err; item.Selected = false; items.Add(item); continue;
                }

                item.QuantityText = display;
                item.TargetName = ReadRowNameAt(template.WorkbookPath, targetSheet, expr);
                if (qty == 0m) { item.Status = "数量为0"; }
                items.Add(item);
            }
            return items;
        }

        // 读某表达式首个单元格所在行的名称（A 到该格列前的非数字文本拼接），仅供人工核对。
        private static string ReadRowNameAt(string workbook, string sheet, string expr)
        {
            try
            {
                string first = ExtractFirstCellAddress(expr);
                CellRef cr;
                if (String.IsNullOrEmpty(first) || !TryParseCellAddress(first, out cr)) return "";
                List<string> parts = new List<string>();
                for (int col = 1; col < cr.Column; col++)
                {
                    string addr = ColumnNumberToName(col) + cr.Row.ToString(CultureInfo.InvariantCulture);
                    string val; string e;
                    if (TryReadXlsxCellValue(workbook, sheet, addr, out val, out e) && !String.IsNullOrWhiteSpace(val))
                    {
                        decimal d; string pe;
                        if (!TryEvaluateDecimal(val, out d, out pe)) parts.Add(val.Trim());
                    }
                }
                return String.Join(" ", parts.Take(6).ToArray()).Trim();
            }
            catch { return ""; }
        }

        // 模式二：固定绑定列。直接读模板记录的原单元格（用户已把目标单元数量粘进该列）。
        private static List<FillPreviewItem> BuildPreview_FixedColumn(FillTemplate template)
        {
            List<FillPreviewItem> items = new List<FillPreviewItem>();
            foreach (FillTemplateRow row in template.Rows)
            {
                FillPreviewItem item = new FillPreviewItem
                {
                    ItemNo = row.ItemNo, QuotaCode = row.QuotaCode, Adjust = row.Adjust,
                    SourceName = row.SourceName, OrderInItem = row.OrderInItem,
                    SourceQuotaSeq = row.SourceQuotaSeq
                };

                if (String.IsNullOrWhiteSpace(row.SourceExpr))
                {
                    item.Status = "模板未记录取数位置"; item.Selected = false; items.Add(item); continue;
                }

                string display; decimal qty; string err;
                if (!TryEvaluateWorkbookExpression(template.WorkbookPath, row.SourceSheet, row.SourceExpr, out display, out qty, out err))
                {
                    item.Status = "取数失败：" + err; item.Selected = false; items.Add(item); continue;
                }

                item.QuantityText = display;
                item.TargetName = ReadRowNameAt(template.WorkbookPath, row.SourceSheet, row.SourceExpr);
                if (qty == 0m) { item.Status = "数量为0"; }
                items.Add(item);
            }
            return items;
        }

        // 写入：把选中预览项对应的源定额行，直接复制到【目标单元】的对应条目（条目序号全局共享，原样保留），
        // 改 总概算序号/顺号/工程数量、丢弃旧 定额序号(新建标识)。不走界面树。
        private static string ApplyFill(Form mainForm, string targetUnitNo, List<FillPreviewItem> items)
        {
            List<FillPreviewItem> selected = items
                .Where(i => i.Selected && String.IsNullOrEmpty(i.Status) && i.SourceQuotaSeq > 0)
                .OrderBy(i => i.ItemNo, StringComparer.OrdinalIgnoreCase)
                .ThenBy(i => i.OrderInItem)
                .ToList();
            if (selected.Count == 0) return "没有可写入的行。";

            using (SqlConnection conn = AgentCreateWorkConnection(mainForm))
            {
                long targetSeq = ResolveUnitSeq(conn, targetUnitNo);
                if (targetSeq <= 0) return "找不到目标单元：" + targetUnitNo + "（请填 _ZGS_编号 或单元名称）。";

                AgentUndoRecord undo = new AgentUndoRecord { Summary = "模板铺量 -> 单元 " + targetUnitNo, Time = DateTime.Now };
                StringBuilder msg = new StringBuilder();
                int inserted = 0, skipped = 0;
                // 每个 (条目序号) 的下一个顺号，写入时递增。
                Dictionary<long, int> nextShun = new Dictionary<long, int>();

                foreach (FillPreviewItem item in selected)
                {
                    Dictionary<string, object> row = LoadAgentFullRow(conn, item.SourceQuotaSeq);
                    if (row == null) { skipped++; continue; }

                    // 条目序号(全局)保持不变 -> 落到目标单元的同一条目。
                    long itemSeq = Convert.ToInt64(row["条目序号"], CultureInfo.InvariantCulture);

                    int shun;
                    if (!nextShun.TryGetValue(itemSeq, out shun))
                    {
                        shun = GetMaxShun(conn, targetSeq, itemSeq) + 1;
                    }

                    // 数量：用预览取到的目标工程量。
                    decimal qty; string qErr;
                    bool okQty = TryEvaluateDecimal(item.QuantityText, out qty, out qErr);

                    row["总概算序号"] = targetSeq;
                    row["顺号"] = shun;
                    row["工程数量输入"] = (object)(item.QuantityText ?? "") ;
                    row["工程数量"] = okQty ? (object)qty : DBNull.Value;
                    row.Remove("定额序号"); // 让数据库分配新标识

                    long newId = InsertQuotaRowReturnId(conn, row);
                    if (newId > 0)
                    {
                        undo.Rows.Add(new AgentUndoRow { Kind = "I", QuotaSequence = newId });
                        inserted++;
                        nextShun[itemSeq] = shun + 1;
                    }
                    else { skipped++; }
                }

                if (undo.Rows.Count > 0)
                {
                    GetAgentUndoStack(mainForm).Add(undo);
                    GetAgentRedoStack(mainForm).Clear();
                }
                RefreshCurrentQuotaGrid(mainForm);

                msg.Append("已向单元 ").Append(targetUnitNo).Append(" 追加定额 ")
                   .Append(inserted.ToString(CultureInfo.InvariantCulture)).Append(" 条");
                if (skipped > 0) msg.Append("，跳过 ").Append(skipped.ToString(CultureInfo.InvariantCulture)).Append(" 条");
                msg.Append("。请在软件点一次“计算”刷新单价/合价与汇总。");
                return msg.ToString();
            }
        }

        // 目标单元某条目下当前最大顺号(无则0)。
        private static int GetMaxShun(SqlConnection conn, long zgsSeq, long itemSeq)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
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
            List<string> cols = values.Keys.ToList();
            StringBuilder sql = new StringBuilder();
            sql.Append("insert into 定额输入 (")
               .Append(String.Join(", ", cols.Select(c => "[" + c + "]").ToArray()))
               .Append(") values (")
               .Append(String.Join(", ", cols.Select((c, i) => "@p" + i.ToString(CultureInfo.InvariantCulture)).ToArray()))
               .Append("); select cast(scope_identity() as bigint);");
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = sql.ToString();
                for (int i = 0; i < cols.Count; i++)
                    cmd.Parameters.AddWithValue("@p" + i.ToString(CultureInfo.InvariantCulture), values[cols[i]] ?? DBNull.Value);
                object o = cmd.ExecuteScalar();
                return (o == null || o == DBNull.Value) ? 0 : Convert.ToInt64(o, CultureInfo.InvariantCulture);
            }
        }
    }
}
