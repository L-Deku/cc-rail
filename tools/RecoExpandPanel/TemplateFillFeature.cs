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

        // 生成模板（跟着绑定走）：收集【源单元 unitNo】里、绑定到【源 sheet sourceSheet】的定额，
        // 自动跨各条目。按 sheet 过滤可排除绑定库里其它专业(其它 sheet，如站场)的历史绑定。
        private static FillTemplate BuildFillTemplateFromBindings(
            Form mainForm, SqlConnection conn, string templateName, string unitNo, string sourceSheet)
        {
            FillTemplate template = new FillTemplate { Name = templateName, SourceUnitNo = unitNo };

            // 1) 选出本专业的绑定：同单元 + 同 sheet。
            ExcelLinkStore store = LoadStore(conn);
            string sheet = (sourceSheet ?? "").Trim();
            List<ExcelQuotaLink> picked = store.Links
                .Where(l => l != null
                    && (String.IsNullOrEmpty(unitNo) || String.Equals((l.TotalNo ?? "").Trim(), unitNo.Trim(), StringComparison.OrdinalIgnoreCase))
                    && String.Equals((l.WorksheetName ?? "").Trim(), sheet, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (picked.Count == 0) return template;
            template.WorkbookPath = picked[0].ExcelPath;

            // 2) 一次查出本单元全部定额的 条目/编号/调整/项目名/顺号，按 定额序号 建索引。
            Dictionary<long, FillTemplateRow> byId = new Dictionary<long, FillTemplateRow>();
            Dictionary<long, int> shunById = new Dictionary<long, int>();
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    "select DE.定额序号, ZJ.条目编号, ZJ.条目名称, DE.定额编号, " +
                    "cast(DE.定额调整 as nvarchar(max)), DE.顺号, DE.工程或费用项目名称 " +
                    "from 定额输入 DE inner join 章节表 ZJ on DE.条目序号=ZJ.条目序号 " +
                    "where DE.总概算编号=@unit";
                cmd.Parameters.AddWithValue("@unit", (object)(unitNo ?? "") );
                using (SqlDataReader r = cmd.ExecuteReader())
                {
                    while (r.Read())
                    {
                        long id = Convert.ToInt64(r.GetValue(0), CultureInfo.InvariantCulture);
                        byId[id] = new FillTemplateRow
                        {
                            ItemNo = Convert.ToString(r.GetValue(1)).Trim(),
                            ItemName = Convert.ToString(r.GetValue(2)).Trim(),
                            QuotaCode = Convert.ToString(r.GetValue(3)).Trim(),
                            Adjust = r.IsDBNull(4) ? "" : Convert.ToString(r.GetValue(4)).Trim(),
                            SourceName = r.IsDBNull(6) ? "" : Convert.ToString(r.GetValue(6)).Trim()
                        };
                        int shun;
                        shunById[id] = Int32.TryParse(Convert.ToString(r.GetValue(5)), NumberStyles.Integer, CultureInfo.InvariantCulture, out shun) ? shun : 0;
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
                    SourceName = row.SourceName, OrderInItem = row.OrderInItem
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
                    SourceName = row.SourceName, OrderInItem = row.OrderInItem
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

        // 套用：把选中的预览项按条目分组，追加到【软件当前单元】，并套定额调整，登记撤销。
        // 目标单元 = 软件当前单元（用户须先在软件里切到目标单元）。
        private static string ApplyFill(Form mainForm, List<FillPreviewItem> items)
        {
            List<FillPreviewItem> selected = items
                .Where(i => i.Selected && String.IsNullOrEmpty(i.Status))
                .OrderBy(i => i.ItemNo, StringComparer.OrdinalIgnoreCase)
                .ThenBy(i => i.OrderInItem)
                .ToList();
            if (selected.Count == 0) return "没有可写入的行。";

            using (SqlConnection conn = AgentCreateWorkConnection(mainForm))
            {
                AgentUndoRecord undo = new AgentUndoRecord { Summary = "模板铺量（当前单元）", Time = DateTime.Now };
                StringBuilder msg = new StringBuilder();
                int totalInserted = 0, totalAdjusted = 0;

                foreach (IGrouping<string, FillPreviewItem> g in selected.GroupBy(i => i.ItemNo))
                {
                    if (!ItemNoExists(conn, g.Key))
                    {
                        if (msg.Length > 0) msg.AppendLine();
                        msg.Append("条目 ").Append(g.Key).Append("：项目章节表中不存在该条目编号，已跳过。");
                        continue;
                    }

                    List<FillPreviewItem> rows = g.OrderBy(i => i.OrderInItem).ToList();
                    AgentInsertGroup group = new AgentInsertGroup { ItemNo = g.Key };
                    foreach (FillPreviewItem r in rows)
                    {
                        group.Quotas.Add(new AgentQuotaInput { Code = r.QuotaCode, Quantity = r.QuantityText });
                    }

                    HashSet<long> before = LoadAgentItemQuotaIds(conn, g.Key);
                    ExecuteAgentInsertGroup(mainForm, conn, group, undo);
                    HashSet<long> after = LoadAgentItemQuotaIds(conn, g.Key);
                    List<long> added = after.Where(id => !before.Contains(id)).OrderBy(id => id).ToList();
                    totalInserted += added.Count;

                    // 阶段2：套定额调整（按追加顺序与 rows 配对；定额序号自增，故 id 升序≈插入先后）
                    int n = Math.Min(added.Count, rows.Count);
                    for (int k = 0; k < n; k++)
                    {
                        string adj = rows[k].Adjust;
                        if (String.IsNullOrWhiteSpace(adj)) continue;
                        ApplyAdjustToQuota(conn, added[k], adj, undo);
                        totalAdjusted++;
                    }
                }

                if (undo.Rows.Count > 0)
                {
                    GetAgentUndoStack(mainForm).Add(undo);
                    GetAgentRedoStack(mainForm).Clear();
                }
                RefreshCurrentQuotaGrid(mainForm);

                msg.Insert(0, "已向当前单元追加定额 " + totalInserted + " 条，套用调整 " + totalAdjusted + " 条。"
                    + (msg.Length > 0 ? Environment.NewLine : ""));
                return msg.ToString();
            }
        }

        // 条目编号是否存在于项目全局章节表。
        private static bool ItemNoExists(SqlConnection conn, string itemNo)
        {
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select count(*) from 章节表 where 条目编号=@bh";
                cmd.Parameters.AddWithValue("@bh", itemNo);
                return Convert.ToInt32(cmd.ExecuteScalar(), CultureInfo.InvariantCulture) > 0;
            }
        }

        // 把定额调整整串写入某定额，并登记到撤销记录（F 类，可撤销恢复原值）。
        private static void ApplyAdjustToQuota(SqlConnection conn, long quotaSeq, string adjust, AgentUndoRecord undo)
        {
            string oldAdj;
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select cast(定额调整 as nvarchar(max)) from 定额输入 where 定额序号=@id";
                cmd.Parameters.AddWithValue("@id", quotaSeq);
                object o = cmd.ExecuteScalar();
                oldAdj = (o == null || o == DBNull.Value) ? "" : Convert.ToString(o);
            }
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "update 定额输入 set 定额调整=@adj where 定额序号=@id";
                cmd.Parameters.AddWithValue("@adj", (object)adjust ?? DBNull.Value);
                cmd.Parameters.AddWithValue("@id", quotaSeq);
                cmd.ExecuteNonQuery();
            }
            undo.Rows.Add(new AgentUndoRow
            {
                Kind = "F", QuotaSequence = quotaSeq, Table = "定额输入",
                KeyClause = "定额序号=@k0",
                KeyValues = new Dictionary<string, object> { { "@k0", quotaSeq } },
                OldValues = new Dictionary<string, object> { { "定额调整", (object)oldAdj } },
                NewValues = new Dictionary<string, object> { { "定额调整", (object)adjust } }
            });
        }
    }
}
