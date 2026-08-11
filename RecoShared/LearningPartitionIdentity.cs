using System;
using System.Text.RegularExpressions;

internal static class LearningPartitionIdentity
{
    private static readonly Regex EntryCodePattern = new Regex(
        "^[0-9]{2,}(?:-[0-9]+)*$",
        RegexOptions.CultureInvariant);

    internal static string ResolveFromProcessIdentity(string processName, string moduleFileName)
    {
        string probe = ((processName ?? "") + " " + (moduleFileName ?? "")).ToLowerInvariant();
        bool is2020 = probe.Contains("rejjnet2020");
        bool is2024 = probe.Contains("rejjgsnet2024") || probe.Contains("rejjqdnet2024");
        if (is2020 == is2024)
        {
            return "";
        }
        return is2020 ? "2020" : "2024";
    }

    internal static string NormalizeLearningEntryCode(string raw)
    {
        string value = (raw ?? "").Trim()
            .Replace('\u2010', '-')
            .Replace('\u2011', '-')
            .Replace('\u2012', '-')
            .Replace('\u2013', '-')
            .Replace('\u2014', '-')
            .Replace('\u2212', '-')
            .Replace('\uff0d', '-');
        return EntryCodePattern.IsMatch(value) ? value : "";
    }

    internal static string NormalizeLearningMethodNo(string rawMethod)
    {
        string probe = (rawMethod ?? "").Trim().ToLowerInvariant().Replace(" ", "");
        if (probe.Contains("101号文") || probe.Contains("101-estimate")) return "101号文估算";
        if (probe.Contains("tb10801") || probe.Contains("2024")) return "TB 10801—2024";
        if (probe.Contains("国铁科法") || probe.Contains("30号文") || probe.Contains("2020")) return "30号文";
        return "";
    }
}
