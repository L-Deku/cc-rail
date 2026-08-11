using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

internal static class LocalMappingFileStoreHarness
{
    private static void Assert(bool condition, string message)
    {
        if (!condition) throw new InvalidOperationException(message);
    }

    private static Dictionary<string, string> Box(string partition, string boxId, string code, string quantity, int weight)
    {
        return new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            { "record_type", "mapping_box" }, { "software_partition", partition }, { "box_id", boxId },
            { "target_kind", "quota" }, { "target_code", code }, { "target_name", "name-" + code },
            { "target_unit", "m" }, { "quantity_name", quantity }, { "quantity_unit", "m" },
            { "weight", weight.ToString() }, { "accepted_count", "2" }, { "corrected_count", "0" },
            { "rejected_count", "0" }, { "last_used_at", "2026-08-07 10:00:00" }
        };
    }

    private static Dictionary<string, string> Context(string methodNo, string entryCode, string custom)
    {
        Dictionary<string, string> row = Box("2020", "b1", "DY-1", "Cable", 1);
        row["record_type"] = "mapping_context";
        row["method_no"] = methodNo;
        row["entry_code"] = entryCode;
        row["entry_name"] = "entry-" + entryCode;
        row.Remove("weight"); row.Remove("accepted_count"); row.Remove("corrected_count");
        row.Remove("rejected_count"); row.Remove("last_used_at");
        row["custom_context"] = custom;
        return row;
    }

    private static string Json(Dictionary<string, string> row)
    {
        return new System.Web.Script.Serialization.JavaScriptSerializer().Serialize(row);
    }

    private static string Hash(string path)
    {
        using (SHA256 sha = SHA256.Create())
        using (FileStream stream = File.OpenRead(path))
            return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "");
    }

    private static LocalMappingSaveResult Save(string path, Dictionary<string, string> box, params Dictionary<string, string>[] contexts)
    {
        return LocalMappingFileStore.Save(path, "2020", "offline-harness", 1000,
            delegate(List<Dictionary<string, string>> snapshot)
            {
                LocalMappingMutation mutation = new LocalMappingMutation
                {
                    SoftwarePartition = "2020",
                    SourceOperation = "offline-harness",
                    TrimSamples = true,
                    MaxSamplesPerBox = 30
                };
                if (box != null) mutation.MappingBoxes.Add(box);
                foreach (Dictionary<string, string> context in contexts ?? new Dictionary<string, string>[0])
                    mutation.MappingContexts.Add(context);
                return mutation;
            });
    }

    private static void TestPreservation(string root)
    {
        string path = Path.Combine(root, "preserve.jsonl");
        Dictionary<string, string> oldBox = Box("2020", "b1", "DY-1", "Cable", 10);
        oldBox["custom_box"] = "keep";
        oldBox["entry_codes"] = "legacy-must-leave-box";
        string unknown = "{\"record_type\":\"future_type\",\"raw\":\"A B\"}";
        string otherPartition = Json(Box("2024", "b-other", "DY-2", "Other", 5));
        string context30 = Json(Context("30号文", "0101-01", "thirty"));
        Dictionary<string, string> formula101 = Context("101号文估算", "0202-01", "estimate");
        formula101["formula_rule_hash"] = "ABC123";
        formula101["formula_template"] = "V0/1000";
        string context101 = Json(formula101);
        string malformed = "{not-json";
        string text = unknown + "\r\n\r\n" + Json(oldBox) + "\r\n" + context30 + "\r\n" +
            context101 + "\r\n" + malformed + "\r\n" + otherPartition + "\r\n";
        File.WriteAllText(path, text, new UTF8Encoding(true));
        byte[] before = File.ReadAllBytes(path);

        Dictionary<string, string> updatedBox = Box("2020", "b1", "DY-1", "Cable", 30);
        Dictionary<string, string> updated30 = Context("30号文", "0101-01", "ignored-new-value");
        updated30.Remove("custom_context");
        updated30["entry_name"] = "updated-entry";
        LocalMappingSaveResult result = Save(path, updatedBox, updated30);
        Assert(result.Succeeded, "preservation save failed: " + result.Status);
        byte[] after = File.ReadAllBytes(path);
        Assert(after.Length >= 3 && after[0] == 0xEF && after[1] == 0xBB && after[2] == 0xBF, "UTF-8 BOM lost");
        string saved = File.ReadAllText(path, Encoding.UTF8);
        Assert(saved.Contains(unknown + "\r\n\r\n"), "unknown row or blank line changed");
        Assert(saved.Contains(context101 + "\r\n" + malformed + "\r\n" + otherPartition), "untouched context/malformed/other partition changed");
        Assert(saved.Contains("\"custom_box\":\"keep\""), "box unknown field lost");
        Assert(saved.Contains("\"custom_context\":\"thirty\""), "context unknown field lost");
        Assert(saved.Contains("\"entry_name\":\"updated-entry\""), "context owned field was not updated");
        string savedBox = saved.Split(new[] { "\r\n" }, StringSplitOptions.None)
            .First(line => line.Contains("\"record_type\":\"mapping_box\"") && line.Contains("\"box_id\":\"b1\""));
        Assert(!savedBox.Contains("entry_codes") && !savedBox.Contains("method_no") && !savedBox.Contains("entry_code"),
            "mapping_box still contains context fields");
        Assert(!before.SequenceEqual(after), "fixture did not change");
    }

    private static void TestDuplicateContextRefuses(string root)
    {
        string path = Path.Combine(root, "duplicate-context.jsonl");
        string line = Json(Context("30号文", "0101-01", "one"));
        File.WriteAllText(path, line + "\n" + line + "\n", new UTF8Encoding(false));
        string before = Hash(path);
        LocalMappingSaveResult result = Save(path, Box("2020", "b1", "DY-1", "Cable", 20));
        Assert(result.Status == LocalMappingSaveStatus.DuplicateContextIdentity, "duplicate context not rejected");
        Assert(result.LineNumbers.SequenceEqual(new[] { 1, 2 }), "duplicate context line numbers missing");
        Assert(Hash(path) == before, "duplicate context changed original file");
        Assert(File.Exists(result.DiagnosticReportPath), "duplicate context diagnostic missing");
    }

    private static void TestAmbiguousBoxRefuses(string root)
    {
        string path = Path.Combine(root, "ambiguous-box.jsonl");
        Dictionary<string, string> one = Box("2020", "b1", "DY-1", "Cable", 10);
        Dictionary<string, string> two = Box("2020", "b1", "DY-1", "Cable", 20);
        one["custom"] = "A"; two["custom"] = "B";
        File.WriteAllText(path, Json(one) + "\n" + Json(two) + "\n", new UTF8Encoding(false));
        string before = Hash(path);
        LocalMappingSaveResult result = Save(path, Box("2020", "b1", "DY-1", "Cable", 30));
        Assert(result.Status == LocalMappingSaveStatus.AmbiguousBoxUnknownFields, "ambiguous box not rejected");
        Assert(Hash(path) == before, "ambiguous box changed original file");
    }

    private static void TestNewFileAndLf(string root)
    {
        string newPath = Path.Combine(root, "new.jsonl");
        Assert(Save(newPath, Box("2020", "b1", "DY-1", "Cable", 10)).Succeeded, "new save failed");
        byte[] bytes = File.ReadAllBytes(newPath);
        Assert(bytes.Length > 5 && bytes[0] == 0xEF && bytes[1] == 0xBB && bytes[2] == 0xBF, "new file lacks BOM");
        Assert(Encoding.UTF8.GetString(bytes).Contains("\r\n"), "new file lacks CRLF");

        string lfPath = Path.Combine(root, "lf.jsonl");
        string untouched = "{\"record_type\":\"future_type\",\"raw\":\"lf\"}";
        File.WriteAllText(lfPath, untouched + "\n" + Json(Box("2020", "b1", "DY-1", "Cable", 5)) + "\n", new UTF8Encoding(false));
        Assert(Save(lfPath, Box("2020", "b1", "DY-1", "Cable", 15)).Succeeded, "LF save failed");
        byte[] lfBytes = File.ReadAllBytes(lfPath);
        Assert(!(lfBytes.Length >= 3 && lfBytes[0] == 0xEF && lfBytes[1] == 0xBB && lfBytes[2] == 0xBF), "no-BOM file gained BOM");
        string lf = Encoding.UTF8.GetString(lfBytes);
        Assert(lf.StartsWith(untouched + "\n", StringComparison.Ordinal), "LF untouched row changed");
        Assert(!lf.Contains("\r\n"), "LF file gained CRLF");
    }

    private static void TestTrimIsolation(string root)
    {
        string path = Path.Combine(root, "trim.jsonl");
        List<string> lines = new List<string>();
        for (int index = 0; index < 31; index++)
        {
            Dictionary<string, string> box = Box("2020", "trim-box", "DY-1", "Q" + index.ToString("00"), index + 1);
            lines.Add(Json(box));
            Dictionary<string, string> context = Context("30号文", "0101-01", "c" + index);
            context["box_id"] = "trim-box";
            context["quantity_name"] = "Q" + index.ToString("00");
            lines.Add(Json(context));
        }
        Dictionary<string, string> other = Box("2024", "trim-box", "DY-1", "Q00", 1);
        lines.Add(Json(other));
        File.WriteAllText(path, String.Join("\n", lines.ToArray()) + "\n", new UTF8Encoding(false));
        LocalMappingSaveResult result = Save(path, null);
        Assert(result.Succeeded, "trim save failed");
        string saved = File.ReadAllText(path);
        List<Dictionary<string, object>> parsed = File.ReadAllLines(path)
            .Where(line => !String.IsNullOrWhiteSpace(line))
            .Select(line => new System.Web.Script.Serialization.JavaScriptSerializer().DeserializeObject(line) as Dictionary<string, object>)
            .Where(row => row != null).ToList();
        Assert(!parsed.Any(row => Convert.ToString(row["record_type"]) == "mapping_box" &&
            Convert.ToString(row["software_partition"]) == "2020" && Convert.ToString(row["quantity_name"]) == "Q00"),
            "weak current-partition sample remained");
        Assert(!saved.Contains("\"custom_context\":\"c0\""), "trimmed sample context remained");
        Assert(saved.Contains(Json(other)), "other partition sample was trimmed");
        Assert(saved.Contains("\"custom_context\":\"c1\""), "unrelated current context was removed");
    }

    public static int Main(string[] args)
    {
        try
        {
            string root = args.Length == 0 ? Path.Combine(Path.GetTempPath(), "RecoLocalMappingHarness") : args[0];
            Directory.CreateDirectory(root);
            Console.WriteLine("RUN preservation");
            TestPreservation(root);
            Console.WriteLine("RUN duplicate-context");
            TestDuplicateContextRefuses(root);
            Console.WriteLine("RUN ambiguous-box");
            TestAmbiguousBoxRefuses(root);
            Console.WriteLine("RUN encoding");
            TestNewFileAndLf(root);
            Console.WriteLine("RUN trim");
            TestTrimIsolation(root);
            Console.WriteLine("PASS T17 shared local mapping preservation/conflict/encoding/splice/trim");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine(ex.ToString());
            return 1;
        }
    }
}
