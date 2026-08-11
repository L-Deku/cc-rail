using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Web.Script.Serialization;

internal enum LocalMappingSaveStatus
{
    Success,
    DuplicateContextIdentity,
    AmbiguousBoxUnknownFields,
    LockTimeout,
    InvalidPartition,
    IoFailure
}

internal sealed class LocalMappingSaveResult
{
    internal LocalMappingSaveStatus Status;
    internal string FilePath = "";
    internal string FileSha256 = "";
    internal string ConflictIdentity = "";
    internal readonly List<int> LineNumbers = new List<int>();
    internal string SourceOperation = "";
    internal string DiagnosticReportPath = "";
    internal string ErrorMessage = "";

    internal bool Succeeded
    {
        get { return Status == LocalMappingSaveStatus.Success; }
    }
}

internal sealed class LocalMappingMutation
{
    internal string SoftwarePartition = "";
    internal string SourceOperation = "";
    internal readonly List<Dictionary<string, string>> MappingBoxes = new List<Dictionary<string, string>>();
    internal readonly List<Dictionary<string, string>> MappingContexts = new List<Dictionary<string, string>>();
    internal bool TrimSamples = true;
    internal int MaxSamplesPerBox = 30;
}

// Shared by RecoQuotaRecommend.dll and RecoExpandPanel.dll. All read/modify/write work stays
// inside the same named mutex so neither caller can overwrite a lock-before snapshot.
internal static class LocalMappingFileStore
{
    internal const string MutexName = "RecoQuotaData.mapping-boxes.lock";

    private static readonly HashSet<string> BoxOwnedFields = new HashSet<string>(new[]
    {
        "record_type", "software_partition", "box_id", "target_kind", "target_code",
        "target_name", "target_unit", "quantity_name", "quantity_unit", "weight",
        "accepted_count", "corrected_count", "rejected_count", "last_used_at"
    }, StringComparer.OrdinalIgnoreCase);

    private static readonly HashSet<string> ContextOwnedFields = new HashSet<string>(new[]
    {
        "record_type", "software_partition", "method_no", "box_id", "target_kind",
        "target_code", "target_name", "target_unit", "quantity_name", "quantity_unit",
        "entry_code", "entry_name", "formula_rule_hash", "formula_template",
        "formula_target_unit", "formula_method", "formula_software_partition",
        "formula_method_no", "formula_entry_code", "formula_operand_count"
    }, StringComparer.OrdinalIgnoreCase);

    internal static LocalMappingSaveResult Save(
        string path,
        string softwarePartition,
        string sourceOperation,
        int timeoutMilliseconds,
        Func<List<Dictionary<string, string>>, LocalMappingMutation> mutationFactory)
    {
        LocalMappingSaveResult result = NewResult(path, sourceOperation);
        string partition = NormalizePartition(softwarePartition);
        if (partition.Length == 0)
        {
            result.Status = LocalMappingSaveStatus.InvalidPartition;
            result.ErrorMessage = "Unknown learning software partition.";
            return result;
        }

        Mutex mutex = new Mutex(false, MutexName);
        bool acquired = false;
        try
        {
            try
            {
                acquired = mutex.WaitOne(timeoutMilliseconds);
            }
            catch (AbandonedMutexException)
            {
                acquired = true;
            }
            if (!acquired)
            {
                result.Status = LocalMappingSaveStatus.LockTimeout;
                result.FileSha256 = ComputeFileSha256(path);
                result.ErrorMessage = "mapping-boxes lock timeout.";
                return result;
            }

            RawDocument document = RawDocument.Read(path);
            result.FileSha256 = document.Sha256;
            List<Dictionary<string, string>> snapshot = document.Lines
                .Where(line => line.Values != null)
                .Select(line => Clone(line.Values))
                .ToList();
            LocalMappingMutation mutation = mutationFactory == null ? null : mutationFactory(snapshot);
            if (mutation == null)
            {
                mutation = new LocalMappingMutation();
            }
            mutation.SoftwarePartition = NormalizePartition(
                String.IsNullOrWhiteSpace(mutation.SoftwarePartition) ? partition : mutation.SoftwarePartition);
            mutation.SourceOperation = String.IsNullOrWhiteSpace(mutation.SourceOperation)
                ? (sourceOperation ?? "")
                : mutation.SourceOperation;
            if (!String.Equals(mutation.SoftwarePartition, partition, StringComparison.OrdinalIgnoreCase))
            {
                result.Status = LocalMappingSaveStatus.InvalidPartition;
                result.ErrorMessage = "Mutation partition differs from the current process partition.";
                return result;
            }

            LocalMappingSaveResult conflict = FindContextConflict(document, partition, mutation.SourceOperation);
            if (conflict != null)
            {
                conflict.FilePath = path ?? "";
                conflict.FileSha256 = document.Sha256;
                conflict.DiagnosticReportPath = TryWriteDiagnostic(conflict);
                return conflict;
            }

            ApplyResult applied = Apply(document, mutation);
            if (applied.Conflict != null)
            {
                applied.Conflict.FilePath = path ?? "";
                applied.Conflict.FileSha256 = document.Sha256;
                applied.Conflict.SourceOperation = mutation.SourceOperation;
                applied.Conflict.DiagnosticReportPath = TryWriteDiagnostic(applied.Conflict);
                return applied.Conflict;
            }

            if (!applied.Changed)
            {
                result.Status = LocalMappingSaveStatus.Success;
                return result;
            }

            string directory = Path.GetDirectoryName(path);
            if (String.IsNullOrWhiteSpace(directory)) directory = Directory.GetCurrentDirectory();
            Directory.CreateDirectory(directory);
            string temp = Path.Combine(directory, Path.GetFileName(path) + "." + Guid.NewGuid().ToString("N") + ".tmp");
            try
            {
                File.WriteAllBytes(temp, applied.Bytes);
                if (File.Exists(path))
                {
                    File.Replace(temp, path, null, true);
                }
                else
                {
                    File.Move(temp, path);
                }
            }
            finally
            {
                if (File.Exists(temp)) File.Delete(temp);
            }
            result.Status = LocalMappingSaveStatus.Success;
            result.FileSha256 = ComputeFileSha256(path);
            return result;
        }
        catch (Exception ex)
        {
            result.Status = LocalMappingSaveStatus.IoFailure;
            result.ErrorMessage = ex.Message;
            result.FileSha256 = ComputeFileSha256(path);
            return result;
        }
        finally
        {
            if (acquired) mutex.ReleaseMutex();
            mutex.Dispose();
        }
    }

    internal static string BuildMappingBoxIdentity(Dictionary<string, string> row)
    {
        if (!IsRecordType(row, "mapping_box")) return "";
        string partition = NormalizePartition(Get(row, "software_partition"));
        string boxId = Get(row, "box_id").Trim();
        string target = BuildTargetIdentity(row);
        string quantity = BuildQuantitySignature(Get(row, "quantity_name"));
        if (partition.Length == 0 || boxId.Length == 0 || target.Length == 0 || quantity == "|") return "";
        return partition + "\n" + boxId.ToUpperInvariant() + "\n" + target + "\n" + quantity;
    }

    internal static string BuildMappingContextIdentity(Dictionary<string, string> row)
    {
        if (!IsRecordType(row, "mapping_context")) return "";
        string partition = NormalizePartition(Get(row, "software_partition"));
        string methodNo = LearningPartitionIdentity.NormalizeLearningMethodNo(Get(row, "method_no"));
        string boxId = Get(row, "box_id").Trim();
        string target = BuildTargetIdentity(row);
        string quantity = BuildQuantitySignature(Get(row, "quantity_name"));
        string entry = LearningPartitionIdentity.NormalizeLearningEntryCode(Get(row, "entry_code"));
        if (partition.Length == 0 || methodNo.Length == 0 || boxId.Length == 0 ||
            target.Length == 0 || quantity == "|" || entry.Length == 0) return "";
        string identity = partition + "\n" + methodNo + "\n" + boxId.ToUpperInvariant() + "\n" +
            target + "\n" + quantity + "\n" + entry;
        string formulaHash = Get(row, "formula_rule_hash").Trim().ToUpperInvariant();
        return formulaHash.Length == 0 ? identity : identity + "\n" + formulaHash;
    }

    internal static string BuildQuantitySignature(string name)
    {
        string source = (name ?? "").Normalize(NormalizationForm.FormKC).ToUpperInvariant()
            .Replace('\u0424', '\u03A6').Replace('\u00D7', 'X');
        StringBuilder normalized = new StringBuilder(source.Length + 1);
        foreach (char value in source)
        {
            if (!Char.IsWhiteSpace(value)) normalized.Append(value);
        }
        return normalized.ToString() + "|";
    }

    internal static string BuildTargetIdentity(Dictionary<string, string> row)
    {
        string code = Get(row, "target_code").Trim().ToUpperInvariant();
        if (code.Length == 0) return "";
        string kind = Get(row, "target_kind").Trim().ToLowerInvariant();
        if (kind.Length == 0) kind = code.All(Char.IsDigit) ? "material" : "quota";
        string identity = kind + ":" + code;
        if (IsContextSensitiveCode(code))
        {
            identity += "|" + BuildQuantitySignature(Get(row, "target_name")).TrimEnd('|') +
                "|" + NormalizeUnit(Get(row, "target_unit"));
        }
        return identity;
    }

    private static ApplyResult Apply(RawDocument document, LocalMappingMutation mutation)
    {
        Dictionary<int, byte[]> replacements = new Dictionary<int, byte[]>();
        HashSet<int> removals = new HashSet<int>();
        List<Dictionary<string, string>> newBoxes = new List<Dictionary<string, string>>();
        List<Dictionary<string, string>> newContexts = new List<Dictionary<string, string>>();
        Dictionary<string, List<RawLine>> boxIndex = Index(document, "mapping_box", mutation.SoftwarePartition, BuildMappingBoxIdentity);
        Dictionary<string, List<RawLine>> contextIndex = Index(document, "mapping_context", mutation.SoftwarePartition, BuildMappingContextIdentity);

        foreach (IGrouping<string, Dictionary<string, string>> desiredGroup in mutation.MappingBoxes
            .Where(row => row != null)
            .Select(NormalizeDesiredBox)
            .Where(row => BuildMappingBoxIdentity(row).Length > 0)
            .GroupBy(BuildMappingBoxIdentity, StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase))
        {
            Dictionary<string, string> desired = desiredGroup.Last();
            string identity = desiredGroup.Key;
            List<RawLine> existing;
            if (!boxIndex.TryGetValue(identity, out existing) || existing.Count == 0)
            {
                newBoxes.Add(desired);
                continue;
            }
            LocalMappingSaveResult ambiguity = FindBoxUnknownFieldConflict(existing, identity, mutation.SourceOperation);
            if (ambiguity != null) return new ApplyResult { Conflict = ambiguity };
            RawLine earliest = existing.OrderBy(line => line.Index).First();
            Dictionary<string, string> merged = MergeDuplicateBoxes(existing);
            OverlayOwned(merged, desired, true);
            StripBoxContextFields(merged);
            replacements[earliest.Index] = document.EncodeJson(merged);
            foreach (RawLine duplicate in existing.Where(line => line.Index != earliest.Index)) removals.Add(duplicate.Index);
        }

        foreach (Dictionary<string, string> desired in mutation.MappingContexts
            .Where(row => row != null)
            .Select(NormalizeDesiredContext)
            .Where(row => BuildMappingContextIdentity(row).Length > 0)
            .OrderBy(BuildMappingContextIdentity, StringComparer.OrdinalIgnoreCase))
        {
            string identity = BuildMappingContextIdentity(desired);
            List<RawLine> existing;
            if (!contextIndex.TryGetValue(identity, out existing) || existing.Count == 0)
            {
                newContexts.Add(desired);
                continue;
            }
            Dictionary<string, string> merged = Clone(existing[0].Values);
            OverlayOwned(merged, desired, false);
            replacements[existing[0].Index] = document.EncodeJson(merged);
        }

        VirtualRows virtualRows = BuildVirtualRows(document, replacements, removals, newBoxes, newContexts);
        if (mutation.TrimSamples && mutation.MaxSamplesPerBox > 0)
        {
            TrimSamples(virtualRows, mutation.SoftwarePartition, mutation.MaxSamplesPerBox);
        }
        byte[] output = document.Splice(virtualRows);
        return new ApplyResult
        {
            Changed = !document.OriginalBytes.SequenceEqual(output),
            Bytes = output
        };
    }

    private static VirtualRows BuildVirtualRows(
        RawDocument document,
        Dictionary<int, byte[]> replacements,
        HashSet<int> removals,
        List<Dictionary<string, string>> newBoxes,
        List<Dictionary<string, string>> newContexts)
    {
        VirtualRows result = new VirtualRows();
        int lastPartitionBoxSlot = -1;
        string partition = newBoxes.Select(row => NormalizePartition(Get(row, "software_partition"))).FirstOrDefault(value => value.Length > 0) ??
            newContexts.Select(row => NormalizePartition(Get(row, "software_partition"))).FirstOrDefault(value => value.Length > 0) ?? "";
        for (int i = 0; i < document.Lines.Count; i++)
        {
            RawLine raw = document.Lines[i];
            if (raw.Values != null && IsRecordType(raw.Values, "mapping_box") &&
                String.Equals(NormalizePartition(Get(raw.Values, "software_partition")), partition, StringComparison.OrdinalIgnoreCase))
            {
                lastPartitionBoxSlot = i;
            }
            if (removals.Contains(i)) continue;
            byte[] content;
            Dictionary<string, string> values = raw.Values;
            if (replacements.TryGetValue(i, out content))
            {
                values = ParseJson(document.Encoding.GetString(content));
            }
            else
            {
                content = raw.Content;
            }
            result.Rows.Add(new VirtualRow
            {
                OriginalIndex = i,
                Content = content,
                Terminator = raw.Terminator,
                Values = values == null ? null : Clone(values)
            });
        }
        int insertPosition = lastPartitionBoxSlot < 0
            ? result.Rows.Count
            : result.Rows.TakeWhile(row => row.OriginalIndex <= lastPartitionBoxSlot).Count();
        foreach (Dictionary<string, string> row in newBoxes.OrderBy(BuildMappingBoxIdentity, StringComparer.OrdinalIgnoreCase))
        {
            result.Rows.Insert(insertPosition++, VirtualRow.New(document.EncodeJson(row), Clone(row)));
        }
        foreach (Dictionary<string, string> row in newContexts.OrderBy(BuildMappingContextIdentity, StringComparer.OrdinalIgnoreCase))
        {
            result.Rows.Add(VirtualRow.New(document.EncodeJson(row), Clone(row)));
        }
        return result;
    }

    private static void TrimSamples(VirtualRows rows, string partition, int maxSamples)
    {
        List<VirtualRow> boxes = rows.Rows.Where(row => row.Values != null && IsRecordType(row.Values, "mapping_box") &&
            String.Equals(NormalizePartition(Get(row.Values, "software_partition")), partition, StringComparison.OrdinalIgnoreCase)).ToList();
        HashSet<string> removedRelations = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (IGrouping<string, VirtualRow> box in boxes.GroupBy(row => Get(row.Values, "box_id"), StringComparer.OrdinalIgnoreCase))
        {
            List<IGrouping<string, VirtualRow>> samples = box.GroupBy(row => BuildQuantitySignature(Get(row.Values, "quantity_name")), StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Key != "|").ToList();
            if (samples.Count <= maxSamples) continue;
            foreach (IGrouping<string, VirtualRow> weak in samples
                .OrderBy(group => group.Min(row => ReadInt(row.Values, "weight")))
                .ThenBy(group => group.Max(row => Get(row.Values, "last_used_at")), StringComparer.Ordinal)
                .Take(samples.Count - maxSamples))
            {
                foreach (VirtualRow row in weak)
                {
                    row.Removed = true;
                    removedRelations.Add(BuildRelationshipIdentity(row.Values));
                }
            }
        }
        foreach (VirtualRow context in rows.Rows.Where(row => row.Values != null && IsRecordType(row.Values, "mapping_context") &&
            String.Equals(NormalizePartition(Get(row.Values, "software_partition")), partition, StringComparison.OrdinalIgnoreCase)))
        {
            if (removedRelations.Contains(BuildRelationshipIdentity(context.Values))) context.Removed = true;
        }
    }

    private static string BuildRelationshipIdentity(Dictionary<string, string> row)
    {
        return NormalizePartition(Get(row, "software_partition")) + "\n" + Get(row, "box_id").Trim().ToUpperInvariant() + "\n" +
            BuildTargetIdentity(row) + "\n" + BuildQuantitySignature(Get(row, "quantity_name"));
    }

    private static LocalMappingSaveResult FindContextConflict(RawDocument document, string partition, string sourceOperation)
    {
        foreach (IGrouping<string, RawLine> group in document.Lines
            .Where(line => line.Values != null && IsRecordType(line.Values, "mapping_context") &&
                String.Equals(NormalizePartition(Get(line.Values, "software_partition")), partition, StringComparison.OrdinalIgnoreCase))
            .Where(line => BuildMappingContextIdentity(line.Values).Length > 0)
            .GroupBy(line => BuildMappingContextIdentity(line.Values), StringComparer.OrdinalIgnoreCase))
        {
            if (group.Count() <= 1) continue;
            LocalMappingSaveResult result = NewResult(document.Path, sourceOperation);
            result.Status = LocalMappingSaveStatus.DuplicateContextIdentity;
            result.ConflictIdentity = group.Key;
            result.LineNumbers.AddRange(group.Select(line => line.Index + 1));
            result.ErrorMessage = "Duplicate mapping_context identity.";
            return result;
        }
        return null;
    }

    private static LocalMappingSaveResult FindBoxUnknownFieldConflict(List<RawLine> lines, string identity, string sourceOperation)
    {
        Dictionary<string, string> seen = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (RawLine line in lines)
        {
            foreach (KeyValuePair<string, string> pair in line.Values)
            {
                if (IsBoxKnownField(pair.Key)) continue;
                string value;
                if (seen.TryGetValue(pair.Key, out value) && !String.Equals(value, pair.Value, StringComparison.Ordinal))
                {
                    LocalMappingSaveResult result = NewResult("", sourceOperation);
                    result.Status = LocalMappingSaveStatus.AmbiguousBoxUnknownFields;
                    result.ConflictIdentity = identity;
                    result.LineNumbers.AddRange(lines.Select(item => item.Index + 1));
                    result.ErrorMessage = "Ambiguous mapping_box unknown field: " + pair.Key;
                    return result;
                }
                seen[pair.Key] = pair.Value;
            }
        }
        return null;
    }

    private static Dictionary<string, string> MergeDuplicateBoxes(List<RawLine> lines)
    {
        Dictionary<string, string> merged = Clone(lines.OrderBy(line => line.Index).First().Values);
        foreach (string field in new[] { "weight", "accepted_count", "corrected_count", "rejected_count" })
        {
            merged[field] = lines.Max(line => ReadInt(line.Values, field)).ToString(CultureInfo.InvariantCulture);
        }
        merged["last_used_at"] = lines.Max(line => Get(line.Values, "last_used_at"));
        foreach (RawLine line in lines)
        {
            foreach (KeyValuePair<string, string> pair in line.Values)
            {
                if (!merged.ContainsKey(pair.Key)) merged[pair.Key] = pair.Value;
            }
        }
        return merged;
    }

    private static void OverlayOwned(Dictionary<string, string> destination, Dictionary<string, string> source, bool box)
    {
        foreach (KeyValuePair<string, string> pair in source)
        {
            if (box ? BoxOwnedFields.Contains(pair.Key) : IsContextOwnedField(pair.Key)) destination[pair.Key] = pair.Value ?? "";
        }
    }

    private static Dictionary<string, string> NormalizeDesiredBox(Dictionary<string, string> source)
    {
        Dictionary<string, string> row = Clone(source);
        row["record_type"] = "mapping_box";
        row["software_partition"] = NormalizePartition(Get(row, "software_partition"));
        StripBoxContextFields(row);
        return row;
    }

    private static Dictionary<string, string> NormalizeDesiredContext(Dictionary<string, string> source)
    {
        Dictionary<string, string> row = Clone(source);
        row["record_type"] = "mapping_context";
        row["software_partition"] = NormalizePartition(Get(row, "software_partition"));
        row["method_no"] = LearningPartitionIdentity.NormalizeLearningMethodNo(Get(row, "method_no"));
        row["entry_code"] = LearningPartitionIdentity.NormalizeLearningEntryCode(Get(row, "entry_code"));
        return row;
    }

    private static void StripBoxContextFields(Dictionary<string, string> row)
    {
        foreach (string key in row.Keys.Where(key => IsContextOwnedField(key) ||
            String.Equals(key, "method", StringComparison.OrdinalIgnoreCase) ||
            String.Equals(key, "project_id", StringComparison.OrdinalIgnoreCase) ||
            String.Equals(key, "entry_codes", StringComparison.OrdinalIgnoreCase)).ToList())
        {
            if (!BoxOwnedFields.Contains(key)) row.Remove(key);
        }
    }

    private static bool IsBoxKnownField(string key)
    {
        return BoxOwnedFields.Contains(key) || IsContextOwnedField(key) ||
            String.Equals(key, "method", StringComparison.OrdinalIgnoreCase) ||
            String.Equals(key, "project_id", StringComparison.OrdinalIgnoreCase) ||
            String.Equals(key, "entry_codes", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsContextOwnedField(string key)
    {
        return ContextOwnedFields.Contains(key) || (key ?? "").StartsWith("formula_operand_", StringComparison.OrdinalIgnoreCase);
    }

    private static Dictionary<string, List<RawLine>> Index(
        RawDocument document,
        string recordType,
        string partition,
        Func<Dictionary<string, string>, string> identityBuilder)
    {
        Dictionary<string, List<RawLine>> index = new Dictionary<string, List<RawLine>>(StringComparer.OrdinalIgnoreCase);
        foreach (RawLine line in document.Lines.Where(line => line.Values != null && IsRecordType(line.Values, recordType) &&
            String.Equals(NormalizePartition(Get(line.Values, "software_partition")), partition, StringComparison.OrdinalIgnoreCase)))
        {
            string identity = identityBuilder(line.Values);
            if (identity.Length == 0) continue;
            List<RawLine> slots;
            if (!index.TryGetValue(identity, out slots))
            {
                slots = new List<RawLine>();
                index[identity] = slots;
            }
            slots.Add(line);
        }
        return index;
    }

    private static string TryWriteDiagnostic(LocalMappingSaveResult result)
    {
        try
        {
            string directory = Path.Combine(Path.GetDirectoryName(result.FilePath), "diagnostics");
            Directory.CreateDirectory(directory);
            string prefix = String.IsNullOrWhiteSpace(result.FileSha256) ? "unknown" : result.FileSha256.Substring(0, Math.Min(12, result.FileSha256.Length));
            string path = Path.Combine(directory, "mapping-boxes-identity-conflict-" +
                DateTime.Now.ToString("yyyyMMdd-HHmmss-fff", CultureInfo.InvariantCulture) + "-" + prefix + ".json");
            Dictionary<string, object> report = new Dictionary<string, object>();
            report["file_path"] = result.FilePath;
            report["file_sha256"] = result.FileSha256;
            report["conflict_kind"] = result.Status.ToString();
            report["identity"] = result.ConflictIdentity;
            report["line_numbers"] = result.LineNumbers.ToArray();
            report["source_operation"] = result.SourceOperation;
            report["summary"] = result.ErrorMessage;
            File.WriteAllText(path, new JavaScriptSerializer().Serialize(report), new UTF8Encoding(true));
            return path;
        }
        catch
        {
            return "";
        }
    }

    private static LocalMappingSaveResult NewResult(string path, string sourceOperation)
    {
        return new LocalMappingSaveResult
        {
            FilePath = path ?? "",
            SourceOperation = sourceOperation ?? "",
            Status = LocalMappingSaveStatus.Success
        };
    }

    private static string NormalizePartition(string value)
    {
        string partition = (value ?? "").Trim();
        return partition == "2020" || partition == "2024" ? partition : "";
    }

    private static string NormalizeUnit(string value)
    {
        return (value ?? "").Normalize(NormalizationForm.FormKC).Trim().ToUpperInvariant().Replace(" ", "");
    }

    private static bool IsContextSensitiveCode(string code)
    {
        string baseCode = (code ?? "").Trim().ToUpperInvariant();
        int suffix = baseCode.IndexOfAny(new[] { '*', '/' });
        if (suffix >= 0) baseCode = baseCode.Substring(0, suffix);
        return baseCode == "SF" || baseCode == "SH" || baseCode == "SQ" || baseCode == "ZLF" ||
            baseCode == "LF" || baseCode == "YF" || baseCode == "TLF" || baseCode == "GF" ||
            baseCode == "JF" || baseCode == "XGT1";
    }

    private static bool IsRecordType(Dictionary<string, string> row, string type)
    {
        return String.Equals(Get(row, "record_type").Trim(), type, StringComparison.OrdinalIgnoreCase);
    }

    private static string Get(Dictionary<string, string> row, string key)
    {
        string value;
        return row != null && row.TryGetValue(key, out value) ? value ?? "" : "";
    }

    private static int ReadInt(Dictionary<string, string> row, string key)
    {
        int value;
        return Int32.TryParse(Get(row, key), NumberStyles.Integer, CultureInfo.InvariantCulture, out value) ? value : 0;
    }

    private static Dictionary<string, string> Clone(Dictionary<string, string> source)
    {
        Dictionary<string, string> result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (source != null)
        {
            foreach (KeyValuePair<string, string> pair in source) result[pair.Key] = pair.Value ?? "";
        }
        return result;
    }

    private static Dictionary<string, string> ParseJson(string text)
    {
        if (String.IsNullOrWhiteSpace(text)) return null;
        try
        {
            Dictionary<string, object> source = new JavaScriptSerializer().DeserializeObject(text) as Dictionary<string, object>;
            if (source == null) return null;
            Dictionary<string, string> result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (KeyValuePair<string, object> pair in source)
            {
                if (pair.Value == null || pair.Value is string || pair.Value is bool || pair.Value is char ||
                    pair.Value is byte || pair.Value is short || pair.Value is int || pair.Value is long ||
                    pair.Value is float || pair.Value is double || pair.Value is decimal)
                {
                    result[pair.Key] = pair.Value == null ? "" : Convert.ToString(pair.Value, CultureInfo.InvariantCulture);
                }
                else
                {
                    return null;
                }
            }
            return result;
        }
        catch
        {
            return null;
        }
    }

    private static string ComputeFileSha256(string path)
    {
        try
        {
            if (!File.Exists(path)) return "";
            using (SHA256 sha = SHA256.Create())
            using (FileStream stream = File.OpenRead(path))
            {
                return String.Concat(sha.ComputeHash(stream).Select(value => value.ToString("X2", CultureInfo.InvariantCulture)).ToArray());
            }
        }
        catch
        {
            return "";
        }
    }

    private sealed class ApplyResult
    {
        internal bool Changed;
        internal byte[] Bytes = new byte[0];
        internal LocalMappingSaveResult Conflict;
    }

    private sealed class VirtualRows
    {
        internal readonly List<VirtualRow> Rows = new List<VirtualRow>();
    }

    private sealed class VirtualRow
    {
        internal int OriginalIndex = -1;
        internal byte[] Content = new byte[0];
        internal byte[] Terminator = new byte[0];
        internal Dictionary<string, string> Values;
        internal bool Removed;

        internal static VirtualRow New(byte[] content, Dictionary<string, string> values)
        {
            return new VirtualRow { Content = content, Values = values, Terminator = null };
        }
    }

    private sealed class RawLine
    {
        internal int Index;
        internal byte[] Content = new byte[0];
        internal byte[] Terminator = new byte[0];
        internal Dictionary<string, string> Values;
    }

    private sealed class RawDocument
    {
        internal string Path = "";
        internal byte[] OriginalBytes = new byte[0];
        internal byte[] Bom = new byte[0];
        internal Encoding Encoding = new UTF8Encoding(false, true);
        internal byte[] MainTerminator = new byte[] { 13, 10 };
        internal string Sha256 = "";
        internal readonly List<RawLine> Lines = new List<RawLine>();

        internal static RawDocument Read(string path)
        {
            RawDocument document = new RawDocument();
            document.Path = path ?? "";
            bool exists = File.Exists(path);
            document.OriginalBytes = exists ? File.ReadAllBytes(path) : new byte[0];
            document.Sha256 = ComputeBytesSha256(document.OriginalBytes);
            int offset = 0;
            if (!exists)
            {
                document.Bom = new byte[] { 0xEF, 0xBB, 0xBF };
                document.Encoding = new UTF8Encoding(false, true);
            }
            else if (StartsWith(document.OriginalBytes, new byte[] { 0xEF, 0xBB, 0xBF }))
            {
                document.Bom = new byte[] { 0xEF, 0xBB, 0xBF };
                document.Encoding = new UTF8Encoding(false, true);
                offset = 3;
            }
            else if (StartsWith(document.OriginalBytes, new byte[] { 0xFF, 0xFE }))
            {
                document.Bom = new byte[] { 0xFF, 0xFE };
                document.Encoding = new UnicodeEncoding(false, false, true);
                offset = 2;
            }
            else if (StartsWith(document.OriginalBytes, new byte[] { 0xFE, 0xFF }))
            {
                document.Bom = new byte[] { 0xFE, 0xFF };
                document.Encoding = new UnicodeEncoding(true, false, true);
                offset = 2;
            }
            document.SplitLines(offset);
            return document;
        }

        internal byte[] EncodeJson(Dictionary<string, string> values)
        {
            return Encoding.GetBytes(new JavaScriptSerializer().Serialize(values));
        }

        internal byte[] Splice(VirtualRows rows)
        {
            List<VirtualRow> kept = rows.Rows.Where(row => !row.Removed).ToList();
            using (MemoryStream stream = new MemoryStream())
            {
                stream.Write(Bom, 0, Bom.Length);
                for (int index = 0; index < kept.Count; index++)
                {
                    VirtualRow row = kept[index];
                    stream.Write(row.Content, 0, row.Content.Length);
                    byte[] terminator = row.Terminator;
                    bool isNew = terminator == null;
                    if (isNew) terminator = MainTerminator;
                    if (terminator.Length == 0 && index < kept.Count - 1) terminator = MainTerminator;
                    stream.Write(terminator, 0, terminator.Length);
                }
                return stream.ToArray();
            }
        }

        private void SplitLines(int offset)
        {
            int unit = Encoding is UnicodeEncoding ? 2 : 1;
            byte[] lf = Encoding.GetBytes("\n");
            byte[] cr = Encoding.GetBytes("\r");
            int start = offset;
            int index = offset;
            while (index <= OriginalBytes.Length - lf.Length)
            {
                if (!Matches(OriginalBytes, index, lf))
                {
                    index += unit;
                    continue;
                }
                int contentEnd = index;
                int terminatorStart = index;
                if (index >= cr.Length + offset && Matches(OriginalBytes, index - cr.Length, cr))
                {
                    contentEnd -= cr.Length;
                    terminatorStart -= cr.Length;
                }
                byte[] content = Slice(OriginalBytes, start, contentEnd - start);
                byte[] terminator = Slice(OriginalBytes, terminatorStart, index + lf.Length - terminatorStart);
                AddLine(content, terminator);
                if (Lines.Count == 1) MainTerminator = terminator;
                start = index + lf.Length;
                index = start;
            }
            if (start < OriginalBytes.Length) AddLine(Slice(OriginalBytes, start, OriginalBytes.Length - start), new byte[0]);
            if (Lines.Count == 0 && OriginalBytes.Length > offset) AddLine(Slice(OriginalBytes, offset, OriginalBytes.Length - offset), new byte[0]);
            if (MainTerminator == null || MainTerminator.Length == 0) MainTerminator = Encoding.GetBytes("\r\n");
        }

        private void AddLine(byte[] content, byte[] terminator)
        {
            Dictionary<string, string> values = null;
            try { values = ParseJson(Encoding.GetString(content)); } catch { values = null; }
            Lines.Add(new RawLine { Index = Lines.Count, Content = content, Terminator = terminator, Values = values });
        }

        private static bool StartsWith(byte[] source, byte[] prefix)
        {
            return source.Length >= prefix.Length && Matches(source, 0, prefix);
        }

        private static bool Matches(byte[] source, int offset, byte[] value)
        {
            if (offset < 0 || offset + value.Length > source.Length) return false;
            for (int i = 0; i < value.Length; i++) if (source[offset + i] != value[i]) return false;
            return true;
        }

        private static byte[] Slice(byte[] source, int offset, int count)
        {
            byte[] result = new byte[count];
            Buffer.BlockCopy(source, offset, result, 0, count);
            return result;
        }

        private static string ComputeBytesSha256(byte[] bytes)
        {
            using (SHA256 sha = SHA256.Create())
            {
                return String.Concat(sha.ComputeHash(bytes ?? new byte[0]).Select(value => value.ToString("X2", CultureInfo.InvariantCulture)).ToArray());
            }
        }
    }
}
