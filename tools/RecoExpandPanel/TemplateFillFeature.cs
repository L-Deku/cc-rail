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
    }
}
