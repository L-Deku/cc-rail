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
    internal static class TextMatcher
    {
        public static string Normalize(string text)
        {
            return (text ?? "")
                .Replace("\r", " ")
                .Replace("\n", " ")
                .Replace("（", "(")
                .Replace("）", ")")
                .Replace("，", ",")
                .Replace("、", " ")
                .Trim()
                .ToLowerInvariant();
        }

        public static int NamePairScore(string left, string right)
        {
            string l = Normalize(left).Replace(" ", "");
            string r = Normalize(right).Replace(" ", "");
            if (String.IsNullOrEmpty(l) || String.IsNullOrEmpty(r))
            {
                return 0;
            }

            if (l == r)
            {
                return 120;
            }
            if (l.Contains(r) || r.Contains(l))
            {
                return 95;
            }

            List<string> leftTokens = Keywords(l).Distinct().ToList();
            if (leftTokens.Count == 0)
            {
                return 0;
            }

            int score = 0;
            int possible = 0;
            int hits = 0;
            foreach (string token in leftTokens)
            {
                int tokenScore = PairTokenScore(token);
                possible += tokenScore;
                if (r.Contains(token))
                {
                    hits++;
                    score += tokenScore;
                }
            }

            if (hits == 0)
            {
                return 0;
            }

            score += (int)Math.Round(18.0 * hits / leftTokens.Count);
            return possible > 0 && score > 115 ? 115 : score;
        }

        public static bool HasStrongNamePairMatch(string left, string right)
        {
            string l = Normalize(left).Replace(" ", "");
            string r = Normalize(right).Replace(" ", "");
            if (String.IsNullOrEmpty(l) || String.IsNullOrEmpty(r))
            {
                return false;
            }

            if (l == r || l.Contains(r) || r.Contains(l))
            {
                return true;
            }

            bool steelConcreteBlocked = IsSteelOnlyAgainstConcrete(l, r);
            foreach (string token in Keywords(l).Distinct())
            {
                if (token.Length < 2 || !r.Contains(token))
                {
                    continue;
                }

                if (steelConcreteBlocked && String.Equals(token, "钢筋", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                return true;
            }

            return false;
        }

        public static bool IsSteelOnlyAgainstConcrete(string left, string right)
        {
            string l = Normalize(left);
            string r = Normalize(right);
            bool leftConcrete = IsConcreteQuantityName(l);
            bool rightConcrete = IsConcreteQuantityName(r);
            return (IsSteelQuantityName(l) && !leftConcrete && rightConcrete) ||
                (IsSteelQuantityName(r) && !rightConcrete && leftConcrete);
        }

        public static bool IsSteelQuantityName(string text)
        {
            string value = Normalize(text);
            return value.Contains("钢筋") ||
                value.Contains("hpb") ||
                value.Contains("hrb") ||
                value.Contains("圆钢") ||
                value.Contains("螺纹");
        }

        public static bool IsConcreteQuantityName(string text)
        {
            string value = Normalize(text);
            return value.Contains("混凝土") ||
                value.Contains("砼") ||
                value.Contains("商品混凝土");
        }

        private static int PairTokenScore(string token)
        {
            if (String.IsNullOrWhiteSpace(token))
            {
                return 0;
            }

            if (HasAsciiOrDigit(token))
            {
                return token.Length >= 3 ? 28 : 8;
            }

            if (token.Length == 1)
            {
                return 4;
            }
            if (token.Length == 2)
            {
                return 24;
            }
            if (token.Length == 3)
            {
                return 38;
            }
            return 50;
        }

        public static IEnumerable<string> Keywords(string text)
        {
            string normalized = Normalize(text);
            foreach (string part in normalized.Split(new char[] { ' ', '/', ',', ';', '\t', '(', ')', '[', ']', '+', '-', '*', '=' }, StringSplitOptions.RemoveEmptyEntries))
            {
                string token = part.Trim();
                if (token.Length < 2 || IsNumberLike(token))
                {
                    continue;
                }

                yield return token;
                foreach (string segment in SplitAlphaNumericAndChinese(token))
                {
                    if (segment.Length >= 2 && !IsNumberLike(segment))
                    {
                        yield return segment;
                    }
                }
                if (ContainsChinese(token))
                {
                    for (int i = 0; i + 2 <= token.Length; i++)
                    {
                        string gram = token.Substring(i, 2);
                        if (!IsNumberLike(gram))
                        {
                            yield return gram;
                        }
                    }
                    for (int i = 0; i + 3 <= token.Length; i++)
                    {
                        string gram = token.Substring(i, 3);
                        if (!IsNumberLike(gram))
                        {
                            yield return gram;
                        }
                    }
                    for (int i = 0; i < token.Length; i++)
                    {
                        string gram = token.Substring(i, 1);
                        if (IsPureChinese(gram))
                        {
                            yield return gram;
                        }
                    }
                }
            }
        }

        public static bool HasAsciiOrDigit(string text)
        {
            foreach (char ch in text ?? "")
            {
                if ((ch >= 'a' && ch <= 'z') || (ch >= 'A' && ch <= 'Z') || Char.IsDigit(ch))
                {
                    return true;
                }
            }
            return false;
        }

        public static bool IsPureChinese(string text)
        {
            bool hasChinese = false;
            foreach (char ch in text ?? "")
            {
                if (ch >= 0x4e00 && ch <= 0x9fff)
                {
                    hasChinese = true;
                    continue;
                }

                return false;
            }
            return hasChinese;
        }

        private static IEnumerable<string> SplitAlphaNumericAndChinese(string token)
        {
            StringBuilder builder = new StringBuilder();
            int lastKind = 0;
            foreach (char ch in token ?? "")
            {
                int kind = Char.IsLetterOrDigit(ch) && !(ch >= 0x4e00 && ch <= 0x9fff) ? 1 : ((ch >= 0x4e00 && ch <= 0x9fff) ? 2 : 0);
                if (kind == 0)
                {
                    if (builder.Length > 0)
                    {
                        yield return builder.ToString();
                        builder.Length = 0;
                    }
                    lastKind = 0;
                    continue;
                }

                if (builder.Length > 0 && kind != lastKind)
                {
                    yield return builder.ToString();
                    builder.Length = 0;
                }

                builder.Append(ch);
                lastKind = kind;
            }

            if (builder.Length > 0)
            {
                yield return builder.ToString();
            }
        }

        private static bool ContainsChinese(string text)
        {
            foreach (char ch in text ?? "")
            {
                if (ch >= 0x4e00 && ch <= 0x9fff)
                {
                    return true;
                }
            }
            return false;
        }

        private static bool IsNumberLike(string text)
        {
            decimal value;
            return Decimal.TryParse((text ?? "").Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out value);
        }

        public static bool IsNumberLikeToken(string text)
        {
            return IsNumberLike(text);
        }
    }
}
