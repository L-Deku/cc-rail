using System;
using System.Data;
using System.Globalization;
using System.Linq;

namespace RecoQuotaRecommend
{
    // 原推荐定额窗口已于 2026-07-19 删除,由 RecoExpandPanel 的"推荐定额"(学习库智能铺量)取代。
    // 本类仅保留全库共用的单位换算/兼容工具;类名保持不变,SearchIndexStore/MappingStore 调用方无需改动。
    internal static class RecommendDialog
    {
        internal static string ConvertQuantityForIndex(string quantityText, string excelUnit, string quotaUnit)
        {
            decimal quantity;
            if (!TryEvaluateQuantity(quantityText, out quantity))
            {
                return quantityText;
            }

            UnitScale excel = ParseUnitScale(excelUnit);
            UnitScale quota = ParseUnitScale(quotaUnit);
            if (String.IsNullOrEmpty(excel.BaseUnit) || String.IsNullOrEmpty(quota.BaseUnit))
            {
                return FormatDecimal(quantity);
            }

            if (!String.Equals(excel.BaseUnit, quota.BaseUnit, StringComparison.Ordinal))
            {
                return FormatDecimal(quantity);
            }

            if (excel.Scale <= 0 || quota.Scale <= 0)
            {
                return FormatDecimal(quantity);
            }

            decimal converted = quantity * excel.Scale / quota.Scale;
            return FormatDecimal(converted);
        }

        internal static bool UnitCompatibleForIndex(string left, string right)
        {
            string l = NormalizeUnit(left);
            string r = NormalizeUnit(right);
            return !String.IsNullOrEmpty(l) && !String.IsNullOrEmpty(r) && (l == r || l.EndsWith(r, StringComparison.Ordinal) || r.EndsWith(l, StringComparison.Ordinal));
        }

        private static bool TryEvaluateQuantity(string text, out decimal value)
        {
            value = 0m;
            string expression = (text ?? "").Trim();
            if (expression.StartsWith("=", StringComparison.Ordinal))
            {
                expression = expression.Substring(1);
            }

            if (Decimal.TryParse(expression, NumberStyles.Float, CultureInfo.InvariantCulture, out value) ||
                Decimal.TryParse(expression, NumberStyles.Float, CultureInfo.CurrentCulture, out value))
            {
                return true;
            }

            expression = expression
                .Replace("×", "*")
                .Replace("x", "*")
                .Replace("X", "*")
                .Replace("（", "(")
                .Replace("）", ")");

            if (expression.Any(ch => !(Char.IsDigit(ch) || ch == '.' || ch == '+' || ch == '-' || ch == '*' || ch == '/' || ch == '(' || ch == ')' || Char.IsWhiteSpace(ch))))
            {
                return false;
            }

            try
            {
                DataTable table = new DataTable();
                object result = table.Compute(expression, String.Empty);
                value = Convert.ToDecimal(result, CultureInfo.InvariantCulture);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string FormatDecimal(decimal value)
        {
            return value.ToString("0.########", CultureInfo.InvariantCulture);
        }

        private static UnitScale ParseUnitScale(string unit)
        {
            string normalized = NormalizeRawUnit(unit);
            decimal scale = 1m;
            int index = 0;
            while (index < normalized.Length && (Char.IsDigit(normalized[index]) || normalized[index] == '.'))
            {
                index++;
            }

            if (index > 0)
            {
                Decimal.TryParse(normalized.Substring(0, index), NumberStyles.Float, CultureInfo.InvariantCulture, out scale);
            }

            string baseUnit = index > 0 ? normalized.Substring(index) : normalized;
            UnitScale result = new UnitScale();
            result.Scale = scale <= 0 ? 1m : scale;
            result.BaseUnit = baseUnit;
            return result;
        }

        private static string NormalizeRawUnit(string unit)
        {
            return (unit ?? "")
                .Trim()
                .ToLowerInvariant()
                .Replace(" ", "")
                .Replace("　", "")
                .Replace("²", "2")
                .Replace("³", "3")
                .Replace("ｍ", "m")
                .Replace("㎡", "m2")
                .Replace("㎥", "m3")
                .Replace("立方米", "m3")
                .Replace("平方米", "m2")
                .Replace("延米", "m")
                .Replace("米", "m")
                .Replace("吨", "t")
                .Replace("千克", "kg");
        }

        private static string NormalizeUnit(string unit)
        {
            return NormalizeRawUnit(unit)
                .Replace("100", "")
                .Replace("10", "")
                .Replace("㎡", "m2")
                .Replace("㎥", "m3")
                .Replace("m²", "m2")
                .Replace("m³", "m3");
        }
    }
}
