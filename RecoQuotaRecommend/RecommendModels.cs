using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace RecoQuotaRecommend
{
    // 单元格快照：保留文本/公式/地址/行号/列索引（列索引从 1 开始，0 为合并表头注入的左侧分组列）。
    internal sealed class CellValue
    {
        public string Text;
        public string Formula;
        public string Address;
        public int RowNumber;
        public int SourceIndex;
    }

    internal sealed class ExcelSelection
    {
        public string WorkbookPath;
        public string WorksheetName;
        public readonly List<ExcelQuantityItem> Items = new List<ExcelQuantityItem>();
        // 解析时的原始单元格网格（按行、按列保留结构），供 AI 列映射兜底重新判断列角色。
        public List<List<CellValue>> RawRows;
    }

    internal sealed class ExcelQuantityItem
    {
        public string WorksheetName;
        public int RowNumber;
        public string CellAddress;
        public string Name;
        public string OriginalName;
        public string AiName;
        public int AiNameConfidence;
        public string AiNameReason;
        public bool SkipAiNameNormalization;
        public string SectionName;
        public string Unit;
        public string ValueText;
        public string Formula;
        public string ContextText;
        public string RawRowText;
    }

    internal sealed class RecommendationRow
    {
        public ExcelQuantityItem Item;
        public LearningRecord Record;
        public string QuotaCode;
        public string QuotaName;
        public string QuotaUnit;
        public string BookCode;
        public string Specialty;
        public double BasePrice;
        public string WorkContent;
        public string ConvertedValueText;
        public int Score;
        public string Reason;
        public string Source;
        public string BoxId;
        public string TargetKind;
        public int GridRowIndex;
        public string AiRowId;
        public bool AiPending;
        public List<AiQuotaCandidate> AiCandidates;
        public List<AiMappingCandidate> AiMappingCandidates;
    }

    internal sealed class AiQuotaCandidate
    {
        public IndexQuota Quota;
        public int LocalScore;
    }

    internal sealed class AiMappingCandidate
    {
        public string BoxId;
        public int LocalScore;
        public string SampleNames;
        public List<MappingTarget> Targets;

        public List<RecommendationRow> ToRecommendations(ExcelQuantityItem item, int score, string reason)
        {
            return (Targets ?? new List<MappingTarget>())
                .OrderBy(t => MappingStore.TargetSortRank(t.TargetKind, t.Code))
                .ThenBy(t => t.Code ?? "", StringComparer.OrdinalIgnoreCase)
                .Select(t =>
                {
                    RecommendationRow row = t.ToRecommendation(item, score, BoxId);
                    row.Reason = String.IsNullOrWhiteSpace(reason) ? "DeepSeek\u5bf9\u5e94\u6846\u68c0\u6d4b" : "DeepSeek\u5bf9\u5e94\u6846\u68c0\u6d4b\uff1a" + reason;
                    return row;
                })
                .ToList();
        }
    }

    internal sealed class DeepSeekRequestRow
    {
        public string RowId;
        public ExcelQuantityItem Item;
        public List<AiQuotaCandidate> Candidates;
        public List<AiMappingCandidate> MappingCandidates;
    }

    internal sealed class DeepSeekNameRequestRow
    {
        public string RowId;
        public ExcelQuantityItem Item;
    }

    internal sealed class AiPendingRecommendation
    {
        public RecommendationRow Row;
        public int GridRowIndex;
        public DeepSeekRequestRow Request;
        public EntryScope Scope;
    }

    internal sealed class DeepSeekSelection
    {
        public string RowId;
        public string BoxId;
        public string SelectedCode;
        public int Confidence;
        public string Reason;
        public string ErrorText;
    }

    internal sealed class DeepSeekNameResult
    {
        public string RowId;
        public string QuantityName;
        public int Confidence;
        public string Reason;
    }

    // AI 识别出的工程量表列布局：哪些列组成名称、哪列单位、哪列数量（列索引从 1 开始）。
    internal sealed class DeepSeekColumnLayout
    {
        public int[] NameColumns;
        public int UnitColumn;
        public int QuantityColumn;
        public int Confidence;
    }

    internal sealed class QuotaEntry
    {
        public string TargetKind;
        public string QuotaCode;
        public string QuotaName;
        public string QuotaUnit;

        // 去掉 *乘数 / ×乘数 与 参/换/借 调整后缀，取原始编号（用于判类与索引查找）
        public static string NormalizeCode(string code)
        {
            string value = (code ?? "").Trim();
            if (value.Length == 0)
            {
                return "";
            }

            int cut = value.IndexOfAny(new[] { '*', '×' });
            if (cut >= 0)
            {
                value = value.Substring(0, cut);
            }
            int slash = value.LastIndexOf('/');
            if (slash > 0 && slash < value.Length - 1 && value.Substring(slash + 1).All(Char.IsDigit))
            {
                value = value.Substring(0, slash);
            }
            value = value.Replace("参", "").Replace("换", "").Replace("借", "");
            return value.Trim();
        }

        // 三类：纯数字=材料；含横杠=定额（所有真实定额编号都带横杠，已核对全库无例外）；
        // 其余字母代号（GF/ZLF/LF/JF/SF/YF/TLF/XGT1…）=辅助代号，按材料一样与定额配套使用。
        public static string GuessKind(string code)
        {
            string value = NormalizeCode(code);
            if (value.Length == 0)
            {
                return "quota";
            }
            if (value.All(Char.IsDigit))
            {
                return "material";
            }
            return value.IndexOf('-') >= 0 ? "quota" : "aux";
        }

        public static bool IsQuotaKind(string kind)
        {
            return String.Equals(kind, "quota", StringComparison.OrdinalIgnoreCase);
        }

        public string TargetKey
        {
            get { return (String.IsNullOrWhiteSpace(TargetKind) ? GuessKind(QuotaCode) : TargetKind) + ":" + (QuotaCode ?? "").Trim().ToUpperInvariant(); }
        }
    }

    internal struct UnitScale
    {
        public string BaseUnit;
        public decimal Scale;
    }

    internal sealed class LearningRecord
    {
        public bool IsCorrection;
        public string QuantitySignature;
        public string ProjectName;
        public string BudgetFile;
        public string BudgetGroup;
        public string QuotaCode;
        public string QuotaName;
        public string QuotaUnit;
        public string QuantitySection;
        public string QuantityName;
        public string QuantityUnit;
        public string MatchReason;
        public int MatchScore;
    }
}
