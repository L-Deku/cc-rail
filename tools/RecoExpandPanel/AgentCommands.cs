using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        private const string AgentSelectedToken = "@selected";
        private const string AgentCurrentItemToken = "@currentitem";

        private sealed class AgentQuotaInput
        {
            public string Code;
            public string Quantity;
        }

        private sealed class AgentCommand
        {
            public string Type;
            public List<string> Items = new List<string>();
            public bool IncludeChildren = true;
            public List<string> QuotaFilter = new List<string>();
            public List<string> Units = new List<string>();   // 单元(总概算)过滤：名称/_ZGS_编号/序号；空=当前单元
            public string Operator = "*";
            public string Factor;
            public string SourceItem;
            public List<string> TargetItems = new List<string>();
            public List<AgentQuotaInput> Quotas = new List<AgentQuotaInput>();
            public string NewName;                            // create_unit 用
            public string Target = "quantity";                // multiply_quantity/remove_text: quantity / quota_code / unit_price
            public string Value;                              // set_quantity 用
            public string RemoveText;                         // remove_text 用：要从字段里去掉的子串(如 /100、*9)
            public string Mode = "set";                       // set_adjustment: set(替换) / append(追加)
            public string FromCode;                           // replace_quota_code 用
            public string ToCode;                             // replace_quota_code 用
            public string Scheme;                             // set_transport_scheme 用（方案序号）
            public string TransportParam;                     // set_transport_scheme 可选：同时设参数调整(运输参数,如PH0)
            public string SchemeKind = "材料";                // set_material_scheme: 材料/机械/设备/工费
            public string SchemeName;                         // set_material_scheme 用

            private string TargetLabel
            {
                get
                {
                    if (Target == "quota_code") { return "定额编号"; }
                    if (Target == "unit_price") { return "单价"; }
                    if (Target == "adjustment") { return "定额调整"; }
                    return "工程数量";
                }
            }

            public string Describe()
            {
                string filter = QuotaFilter.Count == 0 ? "" : "，限定额 " + String.Join(",", QuotaFilter.ToArray());
                if (Units.Count > 0)
                {
                    filter += "，限单元 " + String.Join(",", Units.ToArray());
                }

                switch (Type)
                {
                    case "create_unit":
                        return "复制单元 " + SourceItem + " 新建为 \"" + NewName + "\"";
                    case "set_quantity":
                        return "把条目 " + String.Join(",", Items.ToArray()) + " 的工程数量设为 " + Value + filter;
                    case "remove_text":
                        return "从条目 " + String.Join(",", Items.ToArray()) + " 的" + TargetLabel + "里去掉 \"" + RemoveText + "\"" + filter;
                    case "set_adjustment":
                        return (Mode == "append" ? "给条目 " : "把条目 ") + String.Join(",", Items.ToArray()) +
                            (Mode == "append" ? " 的定额追加调整 \"" : " 的定额调整设为 \"") + Value + "\"" + filter;
                    case "replace_quota_code":
                        return "把条目 " + String.Join(",", Items.ToArray()) + " 里的定额 " + FromCode + " 全部换成 " + ToCode + filter;
                    case "set_transport_scheme":
                        return "把条目 " + String.Join(",", Items.ToArray()) + " 的运输方案设为 " + Scheme +
                            (String.IsNullOrEmpty(TransportParam) ? "" : "（参数 " + TransportParam + "）") + filter;
                    case "set_material_scheme":
                        return "把单元 " + String.Join(",", Units.ToArray()) + " 的" + SchemeKind + "费方案设为 \"" + SchemeName + "\"";
                    case "multiply_quantity":
                        return "条目 " + String.Join(",", Items.ToArray()) + " 的" + TargetLabel + " " + Operator + Factor + filter;
                    case "clear_quantity":
                        return "清空条目 " + String.Join(",", Items.ToArray()) + " 的工程数量" + filter;
                    case "delete_quotas":
                        return "删除条目 " + String.Join(",", Items.ToArray()) + " 的定额" + filter;
                    case "copy_quotas":
                        return "把条目 " + SourceItem + " 的定额复制到 " + String.Join(",", TargetItems.ToArray()) + filter;
                    case "insert_quotas":
                        StringBuilder sb = new StringBuilder();
                        foreach (AgentQuotaInput q in Quotas)
                        {
                            if (sb.Length > 0)
                            {
                                sb.Append("、");
                            }
                            sb.Append(q.Code);
                            if (!String.IsNullOrEmpty(q.Quantity))
                            {
                                sb.Append("=").Append(q.Quantity);
                            }
                        }
                        return "在条目 " + String.Join(",", Items.ToArray()) + " 下输入定额 " + sb;
                    default:
                        return Type;
                }
            }
        }

        private sealed class AgentParseResult
        {
            public List<AgentCommand> Commands = new List<AgentCommand>();
            public string Clarification;
            public string Error;
            public string Source;
        }

        private static string NormalizeAgentInput(string input)
        {
            if (input == null)
            {
                return "";
            }

            return input
                .Replace('，', ',')
                .Replace('、', ',')
                .Replace('；', ';')
                .Replace('【', '[')
                .Replace('】', ']')
                .Replace('＝', '=')
                .Replace('　', ' ')
                .Trim();
        }

        private static List<string> SplitAgentList(string text)
        {
            List<string> values = new List<string>();
            foreach (string part in (text ?? "").Split(new char[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string trimmed = part.Trim();
                if (trimmed.Length == 0)
                {
                    continue;
                }

                if (trimmed == "选中" || trimmed == "当前" || trimmed == "这些" ||
                    String.Equals(trimmed, AgentSelectedToken, StringComparison.OrdinalIgnoreCase))
                {
                    trimmed = AgentSelectedToken;
                }

                if (!values.Contains(trimmed))
                {
                    values.Add(trimmed);
                }
            }

            return values;
        }

        // 从 token 流里摘出 单元=xxx 过滤段，返回剩余 token。
        private static List<string> ExtractAgentUnitFilter(List<string> tokens, AgentCommand command)
        {
            List<string> rest = new List<string>();
            foreach (string token in tokens)
            {
                if (token.StartsWith("单元=", StringComparison.Ordinal))
                {
                    command.Units.AddRange(SplitAgentList(token.Substring(3)));
                }
                else
                {
                    rest.Add(token);
                }
            }

            return rest;
        }

        // 从 token 流里摘出形如 [LY-21,LY-34] 的定额过滤段，返回剩余 token。
        private static List<string> ExtractAgentQuotaFilter(List<string> tokens, AgentCommand command)
        {
            List<string> rest = new List<string>();
            StringBuilder filter = null;
            foreach (string token in tokens)
            {
                if (filter == null && token.StartsWith("[", StringComparison.Ordinal))
                {
                    filter = new StringBuilder(token);
                }
                else if (filter != null)
                {
                    filter.Append(",").Append(token);
                }
                else
                {
                    rest.Add(token);
                }

                if (filter != null && filter.ToString().EndsWith("]", StringComparison.Ordinal))
                {
                    string inner = filter.ToString().Trim('[', ']');
                    command.QuotaFilter = SplitAgentList(inner);
                    filter = null;
                }
            }

            return rest;
        }

        // 省略条目编号时的默认目标：当前选中（先用选中的定额行，没有就用当前条目）。
        private static List<string> AgentSelectedItems()
        {
            return new List<string> { AgentSelectedToken };
        }

        // 像条目编号吗：纯数字且带"-"（如 0101-01）。
        // 材料编号没有"-"（如 1294861）；定额编号含字母（含 LY-21 及 ZLF/TLF/SH/SQ 等特殊编号）。
        // 两类都不是条目编号，单写时按定额过滤（作用于当前条目下该编号）。
        private static bool LooksLikeAgentItemNo(string token)
        {
            if (String.IsNullOrEmpty(token))
            {
                return false;
            }

            bool hasDash = false;
            bool hasDigit = false;
            foreach (char c in token)
            {
                if (Char.IsLetter(c))
                {
                    return false;   // 含字母 -> 定额编号（含 ZLF/TLF/SH/SQ 等）
                }

                if (c == '-')
                {
                    hasDash = true;
                }
                else if (Char.IsDigit(c))
                {
                    hasDigit = true;
                }
            }

            return hasDigit && hasDash;   // 纯数字且带"-" -> 条目编号；纯数字无"-" -> 材料编号（当定额过滤）
        }

        // 把"操作/内容"之前的位置参数解析成 条目/定额 过滤。四个字段命令与定额调整共用同一规则：
        //   2 个 -> 条目编号 + 定额编号；
        //   1 个 -> 像条目编号就当条目(该条目下全部定额)，否则当定额编号(当前条目下该定额)；
        //   0 个 -> 当前选中的定额行(没选中则当前条目全部)。
        private static void ApplyAgentItemQuotaLead(AgentCommand command, List<string> lead)
        {
            if (lead.Count >= 2)
            {
                command.Items = SplitAgentList(lead[0]);
                if (command.QuotaFilter.Count == 0)
                {
                    command.QuotaFilter = SplitAgentList(String.Join(",", lead.Skip(1).ToArray()));
                }
            }
            else if (lead.Count == 1)
            {
                if (command.QuotaFilter.Count > 0 || LooksLikeAgentItemNo(lead[0]))
                {
                    command.Items = SplitAgentList(lead[0]);
                }
                else
                {
                    command.Items = new List<string> { AgentCurrentItemToken };
                    command.QuotaFilter = SplitAgentList(lead[0]);
                }
            }
            else
            {
                command.Items = AgentSelectedItems();
            }
        }

        private static bool IsAgentDigits(string token)
        {
            if (String.IsNullOrEmpty(token))
            {
                return false;
            }

            foreach (char c in token)
            {
                if (!Char.IsDigit(c))
                {
                    return false;
                }
            }

            return true;
        }

        // 分号链式：把 "指令A ; 指令B" 拆成多条确定性命令合并到一个计划里。
        // 含分号即按链式处理；每段都必须是确定性命令（链式不支持AI），任一段失败则整体报错。
        // 不含分号时等价于直接调用 TryParseAgentFallback。
        private static bool TryParseAgentChain(string rawInput, out AgentParseResult result)
        {
            string normalized = NormalizeAgentInput(rawInput);
            string[] segments = normalized.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            if (segments.Length <= 1)
            {
                return TryParseAgentFallback(rawInput, out result);
            }

            AgentParseResult combined = new AgentParseResult();
            combined.Source = "fallback";
            int index = 0;
            foreach (string seg in segments)
            {
                index++;
                string segTrim = seg.Trim();
                if (segTrim.Length == 0)
                {
                    continue;
                }

                AgentParseResult segResult;
                if (!TryParseAgentFallback(segTrim, out segResult))
                {
                    // 含分号但某段不是确定性命令 -> 整体交给 AI 处理。
                    result = combined;
                    return false;
                }

                if (!String.IsNullOrEmpty(segResult.Error))
                {
                    combined.Error = "第 " + index.ToString(CultureInfo.InvariantCulture) + " 段：" + segResult.Error;
                    combined.Commands.Clear();
                    result = combined;
                    return true;
                }

                combined.Commands.AddRange(segResult.Commands);
            }

            result = combined;
            return combined.Commands.Count > 0;
        }

        // 兜底确定性语法。返回 true 表示首词命中关键字（解析失败也算命中，错误放在 result.Error 里给出用法）。
        private static bool TryParseAgentFallback(string rawInput, out AgentParseResult result)
        {
            result = new AgentParseResult();
            result.Source = "fallback";
            string input = NormalizeAgentInput(rawInput);
            List<string> tokens = input.Split(new char[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            if (tokens.Count == 0)
            {
                return false;
            }

            string keyword = tokens[0];
            AgentCommand command = new AgentCommand();
            tokens.RemoveAt(0);
            tokens = ExtractAgentUnitFilter(tokens, command);

            if (keyword == "新建单元" || keyword == "复制单元")
            {
                command.Type = "create_unit";
                int fromIndex = tokens.IndexOf("从");
                if (fromIndex < 0)
                {
                    for (int i = 0; i < tokens.Count; i++)
                    {
                        if (tokens[i].StartsWith("从", StringComparison.Ordinal) && tokens[i].Length > 1)
                        {
                            tokens[i] = tokens[i].Substring(1);
                            tokens.Insert(i, "从");
                            fromIndex = i;
                            break;
                        }
                    }
                }

                if (fromIndex < 1 || fromIndex >= tokens.Count - 1)
                {
                    result.Error = "用法：新建单元 新名称 从 源单元名称（源也可以用 _ZGS_02 或总概算序号）";
                    return true;
                }

                command.NewName = String.Join(" ", tokens.Take(fromIndex).ToArray()).Trim();
                command.SourceItem = String.Join(" ", tokens.Skip(fromIndex + 1).ToArray()).Trim();
                if (command.NewName.Length == 0 || command.SourceItem.Length == 0)
                {
                    result.Error = "用法：新建单元 新名称 从 源单元名称";
                    return true;
                }

                result.Commands.Add(command);
                return true;
            }

            if (keyword == "工程数量" || keyword == "定额编号" || keyword == "单价" ||
                keyword == "乘系数" || keyword == "乘数量" || keyword == "乘定额编号" || keyword == "乘单价")
            {
                // 语法：字段 [条目编号] [定额编号] 操作。操作为 *系数 / /系数（乘除）或 删除*系数 / 删除/系数（去掉该乘除）。
                // 省略条目编号和定额编号时，默认当前条目 + 选中定额行。定额编号字段不论装定额/材料编号/ZLF，只在其后追加或删除乘除。
                string target = (keyword == "定额编号" || keyword == "乘定额编号") ? "quota_code"
                    : ((keyword == "单价" || keyword == "乘单价") ? "unit_price" : "quantity");
                command.Target = target;
                tokens = ExtractAgentQuotaFilter(tokens, command);   // 兼容老的 [定额编号] 中括号写法
                if (tokens.Count < 1)
                {
                    result.Error = "用法：" + keyword + " [条目编号] [定额编号] 操作，操作为 *系数 / /系数 或 删除*系数 / 删除/系数；" +
                        "例如：" + keyword + " 0101-01 LY-21 *0.85，或省略条目和定额 \"" + keyword + " *0.85\" 作用于当前条目的选中定额。";
                    return true;
                }

                // 末尾 token 是操作；其前可有"删除"标记；再前依次是 [条目编号] [定额编号]。
                string opToken = tokens[tokens.Count - 1].Trim();
                tokens.RemoveAt(tokens.Count - 1);
                bool isRemove = false;
                if (opToken.StartsWith("删除", StringComparison.Ordinal)) { isRemove = true; opToken = opToken.Substring(2).Trim(); }
                else if (opToken.StartsWith("删", StringComparison.Ordinal)) { isRemove = true; opToken = opToken.Substring(1).Trim(); }
                else if (tokens.Count >= 1 && (tokens[tokens.Count - 1] == "删除" || tokens[tokens.Count - 1] == "删"))
                {
                    isRemove = true;
                    tokens.RemoveAt(tokens.Count - 1);
                }

                ApplyAgentItemQuotaLead(command, tokens);

                if (isRemove)
                {
                    if (target == "unit_price")
                    {
                        result.Error = "单价不支持\"删除\"（单价是计算值，没有可删的乘除表达式）。如需调整单价请用 单价 */系数。";
                        return true;
                    }

                    if (opToken.Length == 0)
                    {
                        result.Error = "删除内容为空。用法：" + keyword + " [条目编号] [定额编号] 删除*系数（或 删除/系数）。";
                        return true;
                    }

                    command.Type = "remove_text";
                    command.RemoveText = opToken;
                    result.Commands.Add(command);
                    return true;
                }

                command.Type = "multiply_quantity";
                string factorText = opToken;
                if (factorText.StartsWith("*", StringComparison.Ordinal) || factorText.StartsWith("×", StringComparison.Ordinal))
                {
                    factorText = factorText.Substring(1);
                }
                else if (factorText.StartsWith("/", StringComparison.Ordinal) || factorText.StartsWith("÷", StringComparison.Ordinal))
                {
                    command.Operator = "/";
                    factorText = factorText.Substring(1);
                }

                decimal factor;
                if (!Decimal.TryParse(factorText, NumberStyles.Float, CultureInfo.InvariantCulture, out factor) || (command.Operator == "/" && factor == 0m))
                {
                    result.Error = "操作格式不正确（应为 *系数 / /系数 或 删除*系数）：" + opToken;
                    return true;
                }

                command.Factor = factor.ToString(CultureInfo.InvariantCulture);
                result.Commands.Add(command);
                return true;
            }

            if (keyword == "清空数量")
            {
                command.Type = "clear_quantity";
                tokens = ExtractAgentQuotaFilter(tokens, command);
                command.Items = tokens.Count < 1 ? AgentSelectedItems() : SplitAgentList(tokens[0]);
                result.Commands.Add(command);
                return true;
            }

            if (keyword == "删除定额")
            {
                command.Type = "delete_quotas";
                tokens = ExtractAgentQuotaFilter(tokens, command);
                command.Items = tokens.Count < 1 ? AgentSelectedItems() : SplitAgentList(tokens[0]);
                result.Commands.Add(command);
                return true;
            }

            if (keyword == "复制定额")
            {
                command.Type = "copy_quotas";
                tokens = ExtractAgentQuotaFilter(tokens, command);
                int toIndex = tokens.IndexOf("到");
                if (toIndex < 0)
                {
                    // 支持 “复制定额 0101-01 到0102-01” 连写
                    for (int i = 0; i < tokens.Count; i++)
                    {
                        if (tokens[i].StartsWith("到", StringComparison.Ordinal) && tokens[i].Length > 1)
                        {
                            tokens[i] = tokens[i].Substring(1);
                            tokens.Insert(i, "到");
                            toIndex = i;
                            break;
                        }
                    }
                }

                if (toIndex < 1 || toIndex >= tokens.Count - 1)
                {
                    result.Error = "用法：复制定额 来源条目编号 到 目标条目编号[,目标条目编号]";
                    return true;
                }

                command.SourceItem = tokens[toIndex - 1].Trim();
                command.TargetItems = SplitAgentList(String.Join(",", tokens.Skip(toIndex + 1).ToArray()));
                if (String.IsNullOrEmpty(command.SourceItem) || command.TargetItems.Count == 0)
                {
                    result.Error = "用法：复制定额 来源条目编号 到 目标条目编号[,目标条目编号]";
                    return true;
                }

                result.Commands.Add(command);
                return true;
            }

            if (keyword == "输入定额" || keyword == "添加定额")
            {
                command.Type = "insert_quotas";
                if (tokens.Count < 1)
                {
                    result.Error = "用法：输入定额 [条目编号] 定额编号=数量 [定额编号=数量 ...]，例如：输入定额 0101-01 LY-21=100 LY-34=35.5；省略条目编号则输入到当前选中条目（定额编号必填）。";
                    return true;
                }

                // 省略条目编号：第一个参数就是 定额编号=数量（含'='）时，默认输入到当前选中条目。
                int insertPairStart;
                if (tokens[0].IndexOf('=') >= 0)
                {
                    command.Items = AgentSelectedItems();
                    insertPairStart = 0;
                }
                else
                {
                    command.Items = SplitAgentList(tokens[0]);
                    insertPairStart = 1;
                }

                for (int i = insertPairStart; i < tokens.Count; i++)
                {
                    string pair = tokens[i].Trim();
                    if (pair.Length == 0)
                    {
                        continue;
                    }

                    int eq = pair.IndexOf('=');
                    AgentQuotaInput quota = new AgentQuotaInput();
                    if (eq > 0)
                    {
                        quota.Code = pair.Substring(0, eq).Trim();
                        quota.Quantity = pair.Substring(eq + 1).Trim();
                    }
                    else
                    {
                        quota.Code = pair;
                        quota.Quantity = "";
                    }

                    if (quota.Code.Length > 0)
                    {
                        command.Quotas.Add(quota);
                    }
                }

                if (command.Quotas.Count == 0)
                {
                    result.Error = "没有解析到任何定额编号。用法：输入定额 [条目编号] 定额编号=数量 ...（省略条目编号时第一个参数就写 定额编号=数量）";
                    return true;
                }

                result.Commands.Add(command);
                return true;
            }

            if (keyword == "设数量")
            {
                command.Type = "set_quantity";
                tokens = ExtractAgentQuotaFilter(tokens, command);
                if (tokens.Count < 1)
                {
                    result.Error = "用法：设数量 [条目编号] 数值 [定额编号]，例如：设数量 0101-01 100；省略条目编号则作用于当前选中。";
                    return true;
                }

                if (tokens.Count == 1)
                {
                    command.Items = AgentSelectedItems();
                    command.Value = tokens[0].Trim();
                }
                else
                {
                    command.Items = SplitAgentList(tokens[0]);
                    command.Value = tokens[1].Trim();
                }

                result.Commands.Add(command);
                return true;
            }

            if (keyword == "删除数量" || keyword == "去掉数量" || keyword == "去掉编号" || keyword == "去掉调整" || keyword == "删数量字段" || keyword == "删编号字段")
            {
                command.Type = "remove_text";
                command.Target = keyword == "去掉调整" ? "adjustment"
                    : ((keyword == "去掉编号" || keyword == "删编号字段") ? "quota_code" : "quantity");
                tokens = ExtractAgentQuotaFilter(tokens, command);
                if (tokens.Count < 1)
                {
                    result.Error = "用法：" + keyword + " 要删除的内容 [条目编号] [定额编号]，例如：删除数量 /100 0101-01；省略条目编号则作用于当前选中。";
                    return true;
                }

                command.RemoveText = tokens[0];
                command.Items = tokens.Count < 2 ? AgentSelectedItems() : SplitAgentList(tokens[1]);
                result.Commands.Add(command);
                return true;
            }

            if (keyword == "设调整" || keyword == "加调整" || keyword == "定额调整")
            {
                // 语法：定额调整 [条目编号] [定额编号] 调整内容（原样写入定额调整字段，整串替换，如 /XG1、/1294861,,1）。
                // 在调整内容前加"删除"则从定额调整里去掉该内容。条目/定额过滤规则与工程数量等字段命令完全一致。
                tokens = ExtractAgentQuotaFilter(tokens, command);   // 兼容老的 [定额编号] 中括号写法
                if (tokens.Count < 1)
                {
                    result.Error = "用法：" + keyword + " [条目编号] [定额编号] 调整内容，例如：" + keyword + " 0101-01 LY-21 /XG1，或 " + keyword + " 0101-01 LY-21 删除 /XG1；省略定额编号=当前选中定额。";
                    return true;
                }

                // 末尾是调整内容；其前可有"删除"标记；再前依次是 [条目编号] 定额编号。
                string content = tokens[tokens.Count - 1].Trim();
                tokens.RemoveAt(tokens.Count - 1);
                bool removeAdj = false;
                if (content.StartsWith("删除", StringComparison.Ordinal)) { removeAdj = true; content = content.Substring(2).Trim(); }
                else if (content.StartsWith("删", StringComparison.Ordinal)) { removeAdj = true; content = content.Substring(1).Trim(); }
                else if (tokens.Count >= 1 && (tokens[tokens.Count - 1] == "删除" || tokens[tokens.Count - 1] == "删"))
                {
                    removeAdj = true;
                    tokens.RemoveAt(tokens.Count - 1);
                }

                if (content.Length == 0)
                {
                    result.Error = "调整内容为空。用法：" + keyword + " [条目编号] 定额编号 [删除] 调整内容，例如：" + keyword + " 0101-01 LY-21 /XG1。";
                    return true;
                }

                ApplyAgentItemQuotaLead(command, tokens);

                if (removeAdj)
                {
                    command.Type = "remove_text";
                    command.Target = "adjustment";
                    command.RemoveText = content;
                }
                else
                {
                    command.Type = "set_adjustment";
                    command.Mode = "set";   // 定额调整（含别名 设调整/加调整）统一为整串写入
                    command.Value = content;
                }

                result.Commands.Add(command);
                return true;
            }

            if (keyword == "替换定额" || keyword == "换定额")
            {
                command.Type = "replace_quota_code";
                if (tokens.Count < 2)
                {
                    result.Error = "用法：替换定额 [条目编号] 原定额编号 新定额编号，例如：替换定额 0101-01 LY-21 QY-100；省略条目编号则作用于当前选中条目。";
                    return true;
                }

                if (tokens.Count == 2)
                {
                    command.Items = AgentSelectedItems();
                    command.FromCode = tokens[0].Trim();
                    command.ToCode = tokens[1].Trim();
                }
                else
                {
                    command.Items = SplitAgentList(tokens[0]);
                    command.FromCode = tokens[1].Trim();
                    command.ToCode = tokens[2].Trim();
                }

                result.Commands.Add(command);
                return true;
            }

            if (keyword == "设运输方案")
            {
                command.Type = "set_transport_scheme";
                tokens = ExtractAgentQuotaFilter(tokens, command);
                if (tokens.Count < 1)
                {
                    result.Error = "用法：设运输方案 [条目编号] 方案序号 [运输参数] 单元=单元名，例如：设运输方案 0101-01 4 PH0 单元=南江路泵房；省略条目编号则作用于当前选中条目。";
                    return true;
                }

                // 方案序号是纯数字；据此判断第一个 token 是条目还是方案号，从而支持省略条目。
                if (IsAgentDigits(tokens[0]))
                {
                    command.Items = AgentSelectedItems();
                    command.Scheme = tokens[0].Trim();
                    if (tokens.Count >= 2)
                    {
                        command.TransportParam = tokens[1].Trim();
                    }
                }
                else
                {
                    command.Items = SplitAgentList(tokens[0]);
                    if (tokens.Count < 2)
                    {
                        result.Error = "缺少方案序号。用法：设运输方案 条目编号 方案序号 [运输参数] 单元=单元名";
                        return true;
                    }

                    command.Scheme = tokens[1].Trim();
                    if (tokens.Count >= 3)
                    {
                        command.TransportParam = tokens[2].Trim();
                    }
                }

                result.Commands.Add(command);
                return true;
            }

            if (keyword == "改材料价" || keyword == "换材料价方案")
            {
                command.Type = "set_material_scheme";
                if (tokens.Count < 1)
                {
                    result.Error = "用法：改材料价 [材料|机械|设备|工费] 方案名称 单元=单元名，例如：改材料价 材料 部颁25年4季度信息价 单元=南江路泵房";
                    return true;
                }

                int nameStart = 0;
                if (tokens[0] == "材料" || tokens[0] == "机械" || tokens[0] == "设备" || tokens[0] == "工费")
                {
                    command.SchemeKind = tokens[0];
                    nameStart = 1;
                }

                command.SchemeName = String.Join(" ", tokens.Skip(nameStart).ToArray()).Trim();
                if (command.SchemeName.Length == 0)
                {
                    result.Error = "用法：改材料价 [材料|机械|设备|工费] 方案名称 单元=单元名";
                    return true;
                }

                result.Commands.Add(command);
                return true;
            }

            return false;
        }

        // 解析 LLM 返回的 {"commands":[...],"need_clarification":""} 信封。
        private static AgentParseResult ParseAgentLlmContent(string content)
        {
            AgentParseResult result = new AgentParseResult();
            result.Source = "llm";
            if (String.IsNullOrWhiteSpace(content))
            {
                result.Error = "AI 返回为空。";
                return result;
            }

            string json = content.Trim();
            if (json.StartsWith("```", StringComparison.Ordinal))
            {
                int firstBrace = json.IndexOf('{');
                int lastBrace = json.LastIndexOf('}');
                if (firstBrace >= 0 && lastBrace > firstBrace)
                {
                    json = json.Substring(firstBrace, lastBrace - firstBrace + 1);
                }
            }

            JavaScriptSerializer serializer = new JavaScriptSerializer();
            serializer.MaxJsonLength = 1024 * 1024 * 4;
            Dictionary<string, object> root;
            try
            {
                root = serializer.DeserializeObject(json) as Dictionary<string, object>;
            }
            catch (Exception ex)
            {
                result.Error = "AI 返回不是有效 JSON：" + ex.Message;
                return result;
            }

            if (root == null)
            {
                result.Error = "AI 返回不是有效 JSON 对象。";
                return result;
            }

            result.Clarification = ReadJsonString(root, "need_clarification", "");
            List<object> commands = GetJsonList(root, "commands");
            if (commands == null)
            {
                return result;
            }

            foreach (object item in commands)
            {
                Dictionary<string, object> dict = item as Dictionary<string, object>;
                if (dict == null)
                {
                    continue;
                }

                AgentCommand command = new AgentCommand();
                command.Type = ReadJsonString(dict, "type", "").Trim().ToLowerInvariant();
                command.Items = ReadAgentJsonStringList(dict, "items");
                command.IncludeChildren = ReadJsonBool(dict, "include_children", true);
                command.QuotaFilter = ReadAgentJsonStringList(dict, "quota_filter");
                command.SourceItem = ReadJsonString(dict, "source_item", "").Trim();
                command.TargetItems = ReadAgentJsonStringList(dict, "target_items");
                command.Units = ReadAgentJsonStringList(dict, "units");
                command.NewName = ReadJsonString(dict, "new_name", "").Trim();
                string op = ReadJsonString(dict, "operator", "*").Trim();
                command.Operator = op == "/" ? "/" : "*";
                command.Factor = ReadAgentJsonNumberText(dict, "factor");
                string target = ReadJsonString(dict, "target", "quantity").Trim().ToLowerInvariant();
                command.Target = (target == "quota_code" || target == "unit_price" || target == "adjustment") ? target : "quantity";
                command.Value = ReadAgentJsonNumberText(dict, "value");
                if (command.Value.Length == 0)
                {
                    command.Value = ReadJsonString(dict, "value", "").Trim();
                }

                command.RemoveText = ReadJsonString(dict, "remove_text", "");
                string mode = ReadJsonString(dict, "mode", "set").Trim().ToLowerInvariant();
                command.Mode = mode == "append" ? "append" : "set";
                command.FromCode = ReadJsonString(dict, "from_code", "").Trim();
                command.ToCode = ReadJsonString(dict, "to_code", "").Trim();
                command.Scheme = ReadAgentJsonNumberText(dict, "scheme");
                if (command.Scheme.Length == 0)
                {
                    command.Scheme = ReadJsonString(dict, "scheme", "").Trim();
                }

                command.TransportParam = ReadJsonString(dict, "transport_param", "").Trim();

                string kind = ReadJsonString(dict, "scheme_kind", "材料").Trim();
                command.SchemeKind = (kind == "机械" || kind == "设备" || kind == "工费") ? kind : "材料";
                command.SchemeName = ReadJsonString(dict, "scheme_name", "").Trim();
                command.Items = command.Items.Select(NormalizeAgentItemToken).Where(v => v.Length > 0).ToList();
                command.TargetItems = command.TargetItems.Select(NormalizeAgentItemToken).Where(v => v.Length > 0).ToList();
                command.SourceItem = NormalizeAgentItemToken(command.SourceItem);

                List<object> quotas = GetJsonList(dict, "quotas");
                if (quotas != null)
                {
                    foreach (object quotaItem in quotas)
                    {
                        Dictionary<string, object> quotaDict = quotaItem as Dictionary<string, object>;
                        if (quotaDict == null)
                        {
                            continue;
                        }

                        AgentQuotaInput quota = new AgentQuotaInput();
                        quota.Code = ReadJsonString(quotaDict, "code", "").Trim();
                        quota.Quantity = ReadAgentJsonNumberText(quotaDict, "quantity");
                        if (quota.Quantity.Length == 0)
                        {
                            quota.Quantity = ReadJsonString(quotaDict, "quantity", "").Trim();
                        }

                        if (quota.Code.Length > 0)
                        {
                            command.Quotas.Add(quota);
                        }
                    }
                }

                string error = ValidateAgentCommandShape(command);
                if (error != null)
                {
                    result.Error = error;
                    result.Commands.Clear();
                    return result;
                }

                result.Commands.Add(command);
            }

            return result;
        }

        private static string NormalizeAgentItemToken(string value)
        {
            string trimmed = (value ?? "").Trim();
            if (String.Equals(trimmed, AgentSelectedToken, StringComparison.OrdinalIgnoreCase))
            {
                return AgentSelectedToken;
            }

            return trimmed;
        }

        // 校验命令结构完整性（不查数据库，只查字段形状）。返回 null 表示通过。
        private static string ValidateAgentCommandShape(AgentCommand command)
        {
            switch (command.Type)
            {
                case "multiply_quantity":
                    if (command.Items.Count == 0)
                    {
                        return "乘系数命令缺少条目。";
                    }

                    decimal factor;
                    if (String.IsNullOrEmpty(command.Factor) ||
                        !Decimal.TryParse(command.Factor, NumberStyles.Float, CultureInfo.InvariantCulture, out factor))
                    {
                        return "乘系数命令缺少有效系数。";
                    }

                    if (command.Operator == "/" && factor == 0m)
                    {
                        return "除系数不能为 0。";
                    }

                    return null;
                case "clear_quantity":
                case "delete_quotas":
                    return command.Items.Count == 0 ? "命令缺少条目。" : null;
                case "copy_quotas":
                    if (String.IsNullOrEmpty(command.SourceItem))
                    {
                        return "复制定额命令缺少来源条目。";
                    }

                    return command.TargetItems.Count == 0 ? "复制定额命令缺少目标条目。" : null;
                case "insert_quotas":
                    if (command.Items.Count == 0)
                    {
                        return "输入定额命令缺少条目。";
                    }

                    return command.Quotas.Count == 0 ? "输入定额命令缺少定额列表。" : null;
                case "create_unit":
                    if (String.IsNullOrEmpty(command.SourceItem))
                    {
                        return "新建单元命令缺少源单元。";
                    }

                    return String.IsNullOrEmpty(command.NewName) ? "新建单元命令缺少新名称。" : null;
                case "set_quantity":
                    if (command.Items.Count == 0)
                    {
                        return "设数量命令缺少条目。";
                    }

                    return String.IsNullOrEmpty(command.Value) ? "设数量命令缺少数值。" : null;
                case "remove_text":
                    if (command.Items.Count == 0)
                    {
                        return "去掉字段命令缺少条目。";
                    }

                    return String.IsNullOrEmpty(command.RemoveText) ? "去掉字段命令缺少要去掉的内容。" : null;
                case "set_adjustment":
                    if (command.Items.Count == 0)
                    {
                        return "设定额调整命令缺少条目。";
                    }

                    return String.IsNullOrEmpty(command.Value) ? "设定额调整命令缺少调整内容。" : null;
                case "replace_quota_code":
                    if (command.Items.Count == 0)
                    {
                        return "替换定额命令缺少条目。";
                    }

                    if (String.IsNullOrEmpty(command.FromCode))
                    {
                        return "替换定额命令缺少原定额编号。";
                    }

                    return String.IsNullOrEmpty(command.ToCode) ? "替换定额命令缺少新定额编号。" : null;
                case "set_transport_scheme":
                {
                    if (command.Items.Count == 0)
                    {
                        return "设运输方案命令缺少条目。";
                    }

                    int scheme;
                    if (String.IsNullOrEmpty(command.Scheme) ||
                        !Int32.TryParse(command.Scheme, NumberStyles.Integer, CultureInfo.InvariantCulture, out scheme))
                    {
                        return "设运输方案命令缺少有效方案序号。";
                    }

                    return null;
                }
                case "set_material_scheme":
                    if (command.Units.Count == 0)
                    {
                        return "改材料价方案命令必须指定单元。";
                    }

                    return String.IsNullOrEmpty(command.SchemeName) ? "改材料价方案命令缺少方案名称。" : null;
                default:
                    return "不支持的命令类型：" + command.Type;
            }
        }

        private static List<string> ReadAgentJsonStringList(Dictionary<string, object> values, string key)
        {
            List<string> list = new List<string>();
            List<object> items = GetJsonList(values, key);
            if (items == null)
            {
                string single = ReadJsonString(values, key, "");
                if (!String.IsNullOrEmpty(single))
                {
                    list.AddRange(SplitAgentList(single));
                }

                return list;
            }

            foreach (object item in items)
            {
                string text = item == null ? "" : Convert.ToString(item, CultureInfo.InvariantCulture).Trim();
                if (text.Length > 0 && !list.Contains(text))
                {
                    list.Add(text);
                }
            }

            return list;
        }

        private static string ReadAgentJsonNumberText(Dictionary<string, object> values, string key)
        {
            object value;
            if (values == null || !values.TryGetValue(key, out value) || value == null)
            {
                return "";
            }

            if (value is int || value is long || value is double || value is decimal)
            {
                return Convert.ToDecimal(value, CultureInfo.InvariantCulture).ToString(CultureInfo.InvariantCulture);
            }

            string text = Convert.ToString(value, CultureInfo.InvariantCulture).Trim();
            decimal parsed;
            if (Decimal.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed))
            {
                return parsed.ToString(CultureInfo.InvariantCulture);
            }

            return text;
        }
    }
}
