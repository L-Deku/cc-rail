using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        private const int AgentTreeLineBudget = 400;

        private sealed class AgentContext
        {
            public string SelectedItemNo = "";
            public string SelectedItemName = "";
            public long SelectedUnitId;
            public List<string> SelectedQuotaCodes = new List<string>();
            public List<string> TreeLines = new List<string>();
            public List<string> UnitLines = new List<string>();
            public bool TreeTruncated;
        }

        // 采集项目上下文（纯数据库 + 快照，无 UI 访问，可在后台线程调用）。
        private static AgentContext CollectAgentContext(SqlConnection conn, AgentSelectionSnapshot selection, string userText)
        {
            AgentContext context = new AgentContext();
            context.SelectedItemNo = selection.ItemNo ?? "";
            context.SelectedItemName = selection.ItemName ?? "";
            context.SelectedUnitId = selection.CurrentUnitId;
            context.SelectedQuotaCodes = selection.QuotaCodes;

            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 总概算序号, 总概算编号, 编制范围 from 总概算信息 order by 总概算序号";
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        context.UnitLines.Add(
                            Convert.ToString(reader.GetValue(0), CultureInfo.InvariantCulture) + " " +
                            Convert.ToString(reader.GetValue(1)).Trim() + " " +
                            Convert.ToString(reader.GetValue(2)).Trim());
                    }
                }
            }

            List<string[]> allItems = new List<string[]>();
            using (SqlCommand cmd = conn.CreateCommand())
            {
                cmd.CommandText = "select 条目编号, 工程或费用项目名称 from 章节表 order by 条目编号";
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string no = reader.IsDBNull(0) ? "" : Convert.ToString(reader.GetValue(0)).Trim();
                        string name = reader.IsDBNull(1) ? "" : Convert.ToString(reader.GetValue(1)).Trim();
                        if (no.Length > 0)
                        {
                            allItems.Add(new string[] { no, name });
                        }
                    }
                }
            }

            if (allItems.Count <= AgentTreeLineBudget)
            {
                foreach (string[] item in allItems)
                {
                    context.TreeLines.Add(item[0] + " " + item[1]);
                }
            }
            else
            {
                context.TreeTruncated = true;
                context.TreeLines = PruneAgentTree(allItems, userText);
            }

            return context;
        }

        // 条目树超出预算时裁剪：优先保留用户原话命中的条目（编号或名称）、其祖先，再按编号长度补满顶层。
        private static List<string> PruneAgentTree(List<string[]> allItems, string userText)
        {
            string text = userText ?? "";
            HashSet<string> picked = new HashSet<string>(StringComparer.Ordinal);
            List<string[]> ordered = new List<string[]>();

            Action<string[]> add = delegate(string[] item)
            {
                if (!picked.Contains(item[0]))
                {
                    picked.Add(item[0]);
                    ordered.Add(item);
                }
            };

            List<string> chunks = ExtractAgentTextChunks(text);
            foreach (string[] item in allItems)
            {
                bool hit = text.IndexOf(item[0], StringComparison.Ordinal) >= 0;
                if (!hit && item[1].Length >= 2)
                {
                    foreach (string chunk in chunks)
                    {
                        if (item[1].IndexOf(chunk, StringComparison.Ordinal) >= 0)
                        {
                            hit = true;
                            break;
                        }
                    }
                }

                if (hit)
                {
                    foreach (string[] ancestor in allItems)
                    {
                        if (ancestor[0].Length < item[0].Length && item[0].StartsWith(ancestor[0], StringComparison.Ordinal))
                        {
                            add(ancestor);
                        }
                    }

                    add(item);
                }
            }

            foreach (string[] item in allItems.OrderBy(i => i[0].Length).ThenBy(i => i[0], StringComparer.Ordinal))
            {
                if (ordered.Count >= AgentTreeLineBudget)
                {
                    break;
                }

                add(item);
            }

            return ordered
                .OrderBy(i => i[0], StringComparer.Ordinal)
                .Take(AgentTreeLineBudget)
                .Select(i => i[0] + " " + i[1])
                .ToList();
        }

        // 把用户原话切成连续中文块（>=2字），用于按名称命中条目。
        private static List<string> ExtractAgentTextChunks(string text)
        {
            List<string> chunks = new List<string>();
            StringBuilder current = new StringBuilder();
            foreach (char c in text ?? "")
            {
                if (c >= 0x4E00 && c <= 0x9FFF)
                {
                    current.Append(c);
                }
                else
                {
                    if (current.Length >= 2)
                    {
                        chunks.Add(current.ToString());
                    }

                    current.Length = 0;
                }
            }

            if (current.Length >= 2)
            {
                chunks.Add(current.ToString());
            }

            return chunks.Distinct().Take(12).ToList();
        }

        private static string BuildAgentParseRequestJson(JavaScriptSerializer serializer, DeepSeekExcelMatchSettings settings, AgentContext context, string userText)
        {
            Dictionary<string, object> body = new Dictionary<string, object>();
            body["task"] = "把用户对铁路概预算软件的中文操作指令解析成结构化命令JSON。";
            body["commands_schema"] = new string[]
            {
                "multiply_quantity：{\"type\":\"multiply_quantity\",\"items\":[\"条目编号\"],\"include_children\":true,\"quota_filter\":[],\"units\":[],\"target\":\"quantity\",\"operator\":\"*\",\"factor\":0.85} 把条目下定额按系数缩放。target=\"quantity\"乘工程数量(默认)，=\"quota_code\"乘定额编号(追加*系数,软件原生缩放,推荐)，=\"unit_price\"直接乘单价列。operator=\"/\"表示除。",
                "set_quantity：{\"type\":\"set_quantity\",\"items\":[\"条目编号\"],\"include_children\":true,\"quota_filter\":[],\"units\":[],\"value\":\"100\"} 把条目下定额的工程数量直接设为某值(可为表达式)。",
                "remove_text：{\"type\":\"remove_text\",\"items\":[\"条目编号\"],\"include_children\":true,\"quota_filter\":[],\"units\":[],\"target\":\"quantity\",\"remove_text\":\"/100\"} 从字段里去掉指定子串。target=\"quantity\"去工程数量里的(如/100,会重算)，=\"quota_code\"去定额编号里的(如*9)，=\"adjustment\"去定额调整里的(如/XG1)。",
                "set_adjustment：{\"type\":\"set_adjustment\",\"items\":[\"条目编号\"],\"include_children\":true,\"quota_filter\":[],\"units\":[],\"mode\":\"append\",\"value\":\"/XG1\"} 设定额调整字段(如/XG1、/1294861,,1)。mode=\"append\"追加到现有调整后，=\"set\"替换整个调整串。",
                "clear_quantity：{\"type\":\"clear_quantity\",\"items\":[\"条目编号\"],\"include_children\":true,\"quota_filter\":[],\"units\":[]} 清空条目下定额的工程数量。",
                "replace_quota_code：{\"type\":\"replace_quota_code\",\"items\":[\"条目编号\"],\"include_children\":false,\"units\":[],\"from_code\":\"LY-21\",\"to_code\":\"QY-100\"} 把条目下所有定额from_code换成to_code。",
                "delete_quotas：{\"type\":\"delete_quotas\",\"items\":[\"条目编号\"],\"include_children\":false,\"quota_filter\":[\"LY-21\"],\"units\":[]} 删除条目下的定额行；quota_filter为空表示该条目全部定额。",
                "copy_quotas：{\"type\":\"copy_quotas\",\"source_item\":\"条目编号\",\"target_items\":[\"条目编号\"],\"quota_filter\":[],\"units\":[]} 把来源条目的定额复制到目标条目。",
                "insert_quotas：{\"type\":\"insert_quotas\",\"items\":[\"条目编号\"],\"quotas\":[{\"code\":\"LY-21\",\"quantity\":\"100\"}]} 在条目下新输入定额。",
                "set_transport_scheme：{\"type\":\"set_transport_scheme\",\"items\":[\"条目编号\"],\"include_children\":true,\"units\":[\"单元\"],\"scheme\":\"4\",\"transport_param\":\"\"} 把条目在某单元的运输方案设为方案序号(整数)。units必填。用户若同时指定运输参数(如PH0/XG0)放进transport_param，没提就留空。",
                "set_material_scheme：{\"type\":\"set_material_scheme\",\"units\":[\"单元\"],\"scheme_kind\":\"材料\",\"scheme_name\":\"方案名称\"} 把单元的材料/机械/设备/工费费方案换成某方案名。scheme_kind∈材料/机械/设备/工费，units必填。",
                "create_unit：{\"type\":\"create_unit\",\"source_item\":\"源单元名称或_ZGS_编号或序号\",\"new_name\":\"新单元名称\"} 复制一个现有单元(总概算)生成新单元。"
            };
            body["rules"] = new string[]
            {
                "只输出严格JSON：{\"commands\":[...],\"need_clarification\":\"\"}，不要输出其他文字。",
                "items、source_item、target_items 里的条目编号必须从 item_tree 中原样选取，绝对不能编造或改写编号。",
                "重要：同一个条目编号在每个单元(总概算)里都存在。用户点名了某个/某些单元时，必须把单元写进 units（用 unit_list 里的序号、_ZGS_编号或名称原文）；用户说\"所有单元/全部单元\"时 units 填 [\"所有\"]；没提单元就留空数组（系统默认只作用于当前单元）。",
                "用户说\"这个/这些/当前/选中的条目(定额)\"时，items 用特殊值 \"@selected\"；说\"当前单元\"时 units 用 [\"@selected\"]。",
                "用户没有指明任何条目编号/条目名称、也没说单元时，默认 items 用 [\"@selected\"]（作用于当前选中的条目或定额），不要因为缺少条目就反问。",
                "用户用条目名称指代时，在 item_tree 里找到名称匹配的行并使用其编号；多个候选无法确定时不要猜，放进 need_clarification 提问。",
                "quota_filter 填用户点名的定额编号（如 LY-21）；用户没限定就留空数组。",
                "factor 必须是数字。用户说\"除以2\"时 operator 用 \"/\"，factor 用 2。",
                "create_unit 的 source_item 用 unit_list 里的名称/_ZGS_编号/序号；new_name 是用户给的新名称，用户没给名称时放进 need_clarification 问。",
                "用户意图不明确、或超出以上命令能力（例如修改材料价、改运输方案、改取费）时，commands 留空数组，并在 need_clarification 里用中文简短说明或反问。",
                "一句话包含多个操作时按顺序输出多条命令。"
            };
            body["examples"] = new object[]
            {
                new Dictionary<string, object>
                {
                    { "user", "把0101-01和0101-02条目的定额数量都乘0.85" },
                    { "assistant", "{\"commands\":[{\"type\":\"multiply_quantity\",\"items\":[\"0101-01\",\"0101-02\"],\"include_children\":true,\"quota_filter\":[],\"units\":[],\"operator\":\"*\",\"factor\":0.85}],\"need_clarification\":\"\"}" }
                },
                new Dictionary<string, object>
                {
                    { "user", "把南江路泵房单元里0308-01条目的数量都除以2" },
                    { "assistant", "{\"commands\":[{\"type\":\"multiply_quantity\",\"items\":[\"0308-01\"],\"include_children\":true,\"quota_filter\":[],\"units\":[\"南江路泵房\"],\"operator\":\"/\",\"factor\":2}],\"need_clarification\":\"\"}" }
                },
                new Dictionary<string, object>
                {
                    { "user", "删除当前条目里LY-21这条定额" },
                    { "assistant", "{\"commands\":[{\"type\":\"delete_quotas\",\"items\":[\"@selected\"],\"include_children\":false,\"quota_filter\":[\"LY-21\"],\"units\":[]}],\"need_clarification\":\"\"}" }
                },
                new Dictionary<string, object>
                {
                    { "user", "把0101-01里的定额编号都乘0.9" },
                    { "assistant", "{\"commands\":[{\"type\":\"multiply_quantity\",\"items\":[\"0101-01\"],\"include_children\":true,\"quota_filter\":[],\"units\":[],\"target\":\"quota_code\",\"operator\":\"*\",\"factor\":0.9}],\"need_clarification\":\"\"}" }
                },
                new Dictionary<string, object>
                {
                    { "user", "把0101-01里LY-37的工程数量里的/100去掉" },
                    { "assistant", "{\"commands\":[{\"type\":\"remove_text\",\"items\":[\"0101-01\"],\"include_children\":true,\"quota_filter\":[\"LY-37\"],\"units\":[],\"target\":\"quantity\",\"remove_text\":\"/100\"}],\"need_clarification\":\"\"}" }
                },
                new Dictionary<string, object>
                {
                    { "user", "给0101-01的LY-444定额加个调整/XG1" },
                    { "assistant", "{\"commands\":[{\"type\":\"set_adjustment\",\"items\":[\"0101-01\"],\"include_children\":true,\"quota_filter\":[\"LY-444\"],\"units\":[],\"mode\":\"append\",\"value\":\"/XG1\"}],\"need_clarification\":\"\"}" }
                },
                new Dictionary<string, object>
                {
                    { "user", "把南江路泵房单元0308-01的运输方案改成3" },
                    { "assistant", "{\"commands\":[{\"type\":\"set_transport_scheme\",\"items\":[\"0308-01\"],\"include_children\":true,\"units\":[\"南江路泵房\"],\"scheme\":\"3\"}],\"need_clarification\":\"\"}" }
                },
                new Dictionary<string, object>
                {
                    { "user", "照着_ZGS_02再建一个单元，叫测算二版" },
                    { "assistant", "{\"commands\":[{\"type\":\"create_unit\",\"source_item\":\"_ZGS_02\",\"new_name\":\"测算二版\"}],\"need_clarification\":\"\"}" }
                }
            };
            body["unit_list"] = context.UnitLines;
            body["item_tree"] = context.TreeLines;
            body["item_tree_truncated"] = context.TreeTruncated;
            body["selected_item"] = String.IsNullOrEmpty(context.SelectedItemNo)
                ? ""
                : context.SelectedItemNo + " " + context.SelectedItemName;
            body["selected_unit"] = context.SelectedUnitId > 0
                ? context.SelectedUnitId.ToString(CultureInfo.InvariantCulture)
                : "";
            body["selected_quota_codes"] = context.SelectedQuotaCodes;
            body["user_message"] = userText ?? "";

            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["model"] = settings.Model;
            payload["stream"] = false;
            payload["temperature"] = 0.1;
            payload["max_tokens"] = 1200;
            payload["response_format"] = new Dictionary<string, object> { { "type", "json_object" } };
            payload["messages"] = new object[]
            {
                new Dictionary<string, object>
                {
                    { "role", "system" },
                    { "content", "你是铁路工程概预算软件的操作指令解析器。你只把用户指令翻译成给定schema的JSON命令，必须保守：编号只能来自item_tree和unit_list，拿不准就用need_clarification反问，绝不编造。" }
                },
                new Dictionary<string, object>
                {
                    { "role", "user" },
                    { "content", serializer.Serialize(body) }
                }
            };

            return serializer.Serialize(payload);
        }

        // 同步调用 DeepSeek 解析指令（调用方负责放到后台线程）。
        private static AgentParseResult RequestAgentParse(DeepSeekExcelMatchSettings settings, AgentContext context, string userText)
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = 1024 * 1024 * 4;
            string requestJson = BuildAgentParseRequestJson(serializer, settings, context, userText);
            string responseJson = SendDeepSeekExcelMatchRequest(settings, requestJson);

            Dictionary<string, object> root = serializer.DeserializeObject(responseJson) as Dictionary<string, object>;
            List<object> choices = GetJsonList(root, "choices");
            Dictionary<string, object> firstChoice = choices == null || choices.Count == 0 ? null : choices[0] as Dictionary<string, object>;
            Dictionary<string, object> message = firstChoice == null ? null : ReadJsonObject(firstChoice, "message");
            string content = message == null ? "" : ReadJsonString(message, "content", "");
            return ParseAgentLlmContent(content);
        }
    }
}
