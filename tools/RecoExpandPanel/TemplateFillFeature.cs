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

        // 从某单元、某条目编号前缀（专业子树）生成模板。
        private static FillTemplate BuildFillTemplateFromUnit(
            Form mainForm, SqlConnection conn, string templateName, string profession,
            string unitNo, string itemNoPrefix)
        {
            FillTemplate template = new FillTemplate
            {
                Name = templateName, Profession = profession, SourceUnitNo = unitNo
            };

            ExcelLinkStore store = LoadStore(conn);
            Dictionary<long, ExcelQuotaLink> linkBySeq = new Dictionary<long, ExcelQuotaLink>();
            foreach (ExcelQuotaLink link in store.Links)
            {
                linkBySeq[link.QuotaSequence] = link;
            }

            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText =
                    "select DE.定额序号, ZJ.条目编号, ZJ.条目名称, DE.定额编号, DE.工程数量输入, " +
                    "cast(DE.定额调整 as nvarchar(max)), DE.顺号, " +
                    "(select top 1 工程或费用项目名称 from 定额输入 d2 where d2.定额序号=DE.定额序号) as nm " +
                    "from 定额输入 DE inner join 章节表 ZJ on DE.条目序号=ZJ.条目序号 " +
                    "where DE.总概算编号=@unit and ZJ.条目编号 like @pfx " +
                    "order by ZJ.条目编号, DE.顺号";
                cmd.Parameters.AddWithValue("@unit", unitNo);
                cmd.Parameters.AddWithValue("@pfx", (itemNoPrefix ?? "") + "%");
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    int order = 0;
                    string lastItem = null;
                    while (reader.Read())
                    {
                        long seq = Convert.ToInt64(reader.GetValue(0), CultureInfo.InvariantCulture);
                        string itemNo = Convert.ToString(reader.GetValue(1)).Trim();
                        if (itemNo != lastItem) { order = 0; lastItem = itemNo; }

                        FillTemplateRow row = new FillTemplateRow
                        {
                            ItemNo = itemNo,
                            ItemName = Convert.ToString(reader.GetValue(2)).Trim(),
                            QuotaCode = Convert.ToString(reader.GetValue(3)).Trim(),
                            Adjust = reader.IsDBNull(5) ? "" : Convert.ToString(reader.GetValue(5)).Trim(),
                            OrderInItem = order++,
                            SourceName = reader.IsDBNull(7) ? "" : Convert.ToString(reader.GetValue(7)).Trim()
                        };

                        ExcelQuotaLink link;
                        if (linkBySeq.TryGetValue(seq, out link))
                        {
                            row.SourceSheet = link.WorksheetName;
                            row.SourceExpr = String.IsNullOrEmpty(link.Expression) ? link.CellAddress : link.Expression;
                            if (template.WorkbookPath == null) template.WorkbookPath = link.ExcelPath;
                        }

                        template.Rows.Add(row);
                    }
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
    }
}
