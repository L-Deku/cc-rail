using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;

internal sealed class RecoSqlCredential
{
    internal string Server;
    internal string User;
    internal string Password;
}

// SQL 凭据来源（按顺序）：
// 1. 程序集同目录的 RecoPluginSql.json：随同事发布包分发，内容为最小权限登录名（reco_plugin），
//    用固定应用口令做 PBKDF2 + AES-256-CBC + HMAC-SHA256 封装，仅防目视，不是保密手段。
// 2. 当前 Windows 用户的 DPAPI 凭据库：管理员本机与离线工具使用，登录名为高权限的 reco。
// 两条来源的明文都是 version=1 加六个 Base64 字段的行文本，与 tools\RecoCredential\RecoCredentialStore.ps1 一致。
internal static class RecoSqlCredentialStore
{
    private const string StoreEnvironmentVariable = "RECO_SQL_CREDENTIAL_STORE_PATH";
    private const string PluginConfigEnvironmentVariable = "RECO_PLUGIN_SQL_CONFIG_PATH";
    internal const string PluginConfigFileName = "RecoPluginSql.json";
    internal const string PluginConfigPackageType = "plugin_sql_config";
    // 仅防目视，非保密。必须与 tools\RecoCredential\RecoCredentialTransfer.ps1 中 Get-RecoPluginSqlConfigKeyBundle 逐字一致。
    private const string PluginConfigPassphrase = "RecoBudget.PluginSqlConfig.v1:9f3a6c1d-5b2e-4e8a-a7c4-2d1f0b8e6c53";
    private const string MacPrefix = "RecoBudget.CredentialTransfer.v1";
    private const string ForbiddenPluginLogin = "reco";
    private const int MinPluginConfigIterations = 10000;
    private const int MaxPluginConfigIterations = 1000000;
    private static readonly byte[] Entropy = Encoding.UTF8.GetBytes("RecoBudget.SqlCredentials.v1");

    internal static string GetStorePath()
    {
        string configured = Environment.GetEnvironmentVariable(StoreEnvironmentVariable);
        if (!String.IsNullOrWhiteSpace(configured))
        {
            if (!Path.IsPathRooted(configured))
            {
                throw new InvalidOperationException(StoreEnvironmentVariable + " must be an absolute path.");
            }
            return Path.GetFullPath(configured);
        }

        string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        return Path.Combine(localAppData, "RecoBudget", "Secrets", "sql-credentials.dpapi");
    }

    internal static string GetPluginConfigPath()
    {
        string configured = Environment.GetEnvironmentVariable(PluginConfigEnvironmentVariable);
        if (!String.IsNullOrWhiteSpace(configured))
        {
            if (!Path.IsPathRooted(configured))
            {
                throw new InvalidOperationException(PluginConfigEnvironmentVariable + " must be an absolute path.");
            }
            return Path.GetFullPath(configured);
        }

        string assemblyDir = Path.GetDirectoryName(typeof(RecoSqlCredentialStore).Assembly.Location) ?? "";
        return Path.Combine(assemblyDir, PluginConfigFileName);
    }

    internal static bool PluginConfigExists()
    {
        return File.Exists(GetPluginConfigPath());
    }

    internal static RecoSqlCredential Read(string name)
    {
        string normalizedName = NormalizeEntryName(name);
        string pluginConfigPath = GetPluginConfigPath();
        if (File.Exists(pluginConfigPath))
        {
            return ReadFromPluginConfig(pluginConfigPath, normalizedName);
        }
        return ReadFromDpapiStore(normalizedName);
    }

    internal static string BuildConnectionString(string name, string database, int port, int timeoutSeconds)
    {
        return BuildConnectionString(Read(name), database, port, timeoutSeconds);
    }

    // 按服务器主机部分（忽略端口）在 learning/business 两条记录里路由，与 tools\RecoLearning\Common.ps1 的 Get-RecoConnectionString 同规则。
    internal static string BuildConnectionStringForServer(string server, string database, int port, int timeoutSeconds)
    {
        string host = NormalizeHost(server);
        if (host.Length == 0)
        {
            throw new ArgumentException("Server is required.", "server");
        }

        foreach (string name in new[] { "learning", "business" })
        {
            RecoSqlCredential credential = Read(name);
            if (String.Equals(NormalizeHost(credential.Server), host, StringComparison.OrdinalIgnoreCase))
            {
                return BuildConnectionString(credential, database, port, timeoutSeconds);
            }
        }

        throw new InvalidOperationException("No SQL credential entry matches server " + host + ".");
    }

    private static string BuildConnectionString(RecoSqlCredential credential, string database, int port, int timeoutSeconds)
    {
        if (String.IsNullOrWhiteSpace(database))
        {
            throw new ArgumentException("Database is required.", "database");
        }

        string dataSource = credential.Server;
        if (port > 0 && dataSource.IndexOf(',') < 0)
        {
            dataSource += "," + port;
        }

        SqlConnectionStringBuilder builder = new SqlConnectionStringBuilder();
        builder.DataSource = dataSource;
        builder.InitialCatalog = database;
        builder.UserID = credential.User;
        builder.Password = credential.Password;
        builder.ConnectTimeout = timeoutSeconds;
        builder.Encrypt = false;
        builder.TrustServerCertificate = true;
        builder.PersistSecurityInfo = false;
        return builder.ConnectionString;
    }

    private static string NormalizeEntryName(string name)
    {
        string normalizedName = (name ?? "").Trim().ToLowerInvariant();
        if (normalizedName != "learning" && normalizedName != "business")
        {
            throw new ArgumentException("Unknown SQL credential entry.", "name");
        }
        return normalizedName;
    }

    private static string NormalizeHost(string server)
    {
        string text = (server ?? "").Trim();
        int comma = text.IndexOf(',');
        if (comma >= 0)
        {
            text = text.Substring(0, comma);
        }
        return text.Trim().ToLowerInvariant();
    }

    private static void RejectReparsePoint(string path, string label)
    {
        if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0)
        {
            throw new InvalidOperationException(label + " cannot be a reparse point.");
        }
    }

    private static RecoSqlCredential ReadFromDpapiStore(string normalizedName)
    {
        string path = GetStorePath();
        if (!File.Exists(path))
        {
            throw new FileNotFoundException(
                "No SQL credential source found: " + PluginConfigFileName + " is not next to the plugin and the local DPAPI SQL credential store is missing.",
                path);
        }
        RejectReparsePoint(path, "The local DPAPI SQL credential store");

        byte[] encrypted = File.ReadAllBytes(path);
        byte[] plaintext = null;
        try
        {
            plaintext = ProtectedData.Unprotect(encrypted, Entropy, DataProtectionScope.CurrentUser);
            return ExtractCredential(ParsePayload(Encoding.UTF8.GetString(plaintext)), normalizedName);
        }
        catch (CryptographicException ex)
        {
            throw new InvalidOperationException(
                "The local DPAPI SQL credential store cannot be decrypted by the current Windows user.", ex);
        }
        finally
        {
            if (plaintext != null)
            {
                Array.Clear(plaintext, 0, plaintext.Length);
            }
            Array.Clear(encrypted, 0, encrypted.Length);
        }
    }

    private static RecoSqlCredential ReadFromPluginConfig(string path, string normalizedName)
    {
        RejectReparsePoint(path, PluginConfigFileName);
        string json = File.ReadAllText(path, Encoding.UTF8);

        if (ExtractJsonNumber(json, "version") != "1")
        {
            throw new InvalidDataException(PluginConfigFileName + " version is unsupported.");
        }
        if (ExtractJsonString(json, "package_type") != PluginConfigPackageType)
        {
            throw new InvalidDataException(PluginConfigFileName + " package type is unsupported.");
        }
        string packageId = ExtractJsonString(json, "package_id");
        int iterations;
        if (!Int32.TryParse(ExtractJsonNumber(json, "iterations"), out iterations) ||
            iterations < MinPluginConfigIterations || iterations > MaxPluginConfigIterations)
        {
            throw new InvalidDataException(PluginConfigFileName + " has an invalid iteration count.");
        }
        string saltText = ExtractJsonString(json, "salt");
        byte[] salt = DecodeBase64(saltText, "salt", 16);
        byte[] iv = DecodeBase64(ExtractJsonString(json, "iv"), "iv", 16);
        byte[] mac = DecodeBase64(ExtractJsonString(json, "mac"), "mac", 32);
        byte[] ciphertext = DecodeBase64(ExtractJsonString(json, "ciphertext"), "ciphertext", 0);
        if (ciphertext.Length == 0 || ciphertext.Length % 16 != 0)
        {
            throw new InvalidDataException(PluginConfigFileName + " has an invalid ciphertext length.");
        }

        string context = "plugin-sql-config|" + packageId + "|" + iterations.ToString() + "|" + saltText;
        byte[] keyBundle = null;
        byte[] encryptionKey = new byte[32];
        byte[] macKey = new byte[32];
        byte[] plaintext = null;
        try
        {
            using (Rfc2898DeriveBytes derive = new Rfc2898DeriveBytes(PluginConfigPassphrase, salt, iterations))
            {
                keyBundle = derive.GetBytes(64);
            }
            Array.Copy(keyBundle, 0, encryptionKey, 0, 32);
            Array.Copy(keyBundle, 32, macKey, 0, 32);

            byte[] prefix = Encoding.UTF8.GetBytes(MacPrefix + "\n" + context + "\n");
            byte[] macInput = new byte[prefix.Length + iv.Length + ciphertext.Length];
            Array.Copy(prefix, 0, macInput, 0, prefix.Length);
            Array.Copy(iv, 0, macInput, prefix.Length, iv.Length);
            Array.Copy(ciphertext, 0, macInput, prefix.Length + iv.Length, ciphertext.Length);
            byte[] actualMac;
            using (HMACSHA256 hmac = new HMACSHA256(macKey))
            {
                actualMac = hmac.ComputeHash(macInput);
            }
            if (!FixedTimeEquals(mac, actualMac))
            {
                throw new InvalidDataException(PluginConfigFileName + " failed integrity verification; the file is corrupt or was modified.");
            }

            using (Aes aes = Aes.Create())
            {
                aes.KeySize = 256;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;
                aes.Key = encryptionKey;
                aes.IV = iv;
                using (ICryptoTransform decryptor = aes.CreateDecryptor())
                {
                    plaintext = decryptor.TransformFinalBlock(ciphertext, 0, ciphertext.Length);
                }
            }

            Dictionary<string, string> values = ParsePayload(Encoding.UTF8.GetString(plaintext));
            foreach (string entry in new[] { "learning", "business" })
            {
                if (String.Equals(Decode(values, entry + ".user"), ForbiddenPluginLogin, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidOperationException(PluginConfigFileName + " must not contain the administrative SQL login.");
                }
            }
            return ExtractCredential(values, normalizedName);
        }
        finally
        {
            if (keyBundle != null) Array.Clear(keyBundle, 0, keyBundle.Length);
            Array.Clear(encryptionKey, 0, encryptionKey.Length);
            Array.Clear(macKey, 0, macKey.Length);
            if (plaintext != null) Array.Clear(plaintext, 0, plaintext.Length);
        }
    }

    private static RecoSqlCredential ExtractCredential(Dictionary<string, string> values, string normalizedName)
    {
        string prefix = normalizedName + ".";
        return new RecoSqlCredential
        {
            Server = Decode(values, prefix + "server"),
            User = Decode(values, prefix + "user"),
            Password = Decode(values, prefix + "password")
        };
    }

    private static bool FixedTimeEquals(byte[] left, byte[] right)
    {
        if (left.Length != right.Length)
        {
            return false;
        }
        int difference = 0;
        for (int index = 0; index < left.Length; index++)
        {
            difference |= left[index] ^ right[index];
        }
        return difference == 0;
    }

    private static byte[] DecodeBase64(string text, string field, int expectedLength)
    {
        byte[] bytes;
        try
        {
            bytes = Convert.FromBase64String(text);
        }
        catch (FormatException ex)
        {
            throw new InvalidDataException(PluginConfigFileName + " has invalid encoding for " + field + ".", ex);
        }
        if (expectedLength > 0 && bytes.Length != expectedLength)
        {
            throw new InvalidDataException(PluginConfigFileName + " has an invalid " + field + " length.");
        }
        return bytes;
    }

    // 配置文件由 New-RecoPluginSqlConfig.ps1 生成，字段值只含 Base64/GUID/数字，不含转义，故用正则提取而不引入 JSON 库。
    private static string ExtractJsonString(string json, string key)
    {
        Match match = Regex.Match(json, "\"" + Regex.Escape(key) + "\"\\s*:\\s*\"([^\"\\\\]*)\"");
        if (!match.Success)
        {
            throw new InvalidDataException(PluginConfigFileName + " is missing " + key + ".");
        }
        return match.Groups[1].Value;
    }

    private static string ExtractJsonNumber(string json, string key)
    {
        Match match = Regex.Match(json, "\"" + Regex.Escape(key) + "\"\\s*:\\s*(\\d+)");
        if (!match.Success)
        {
            throw new InvalidDataException(PluginConfigFileName + " is missing " + key + ".");
        }
        return match.Groups[1].Value;
    }

    private static Dictionary<string, string> ParsePayload(string payload)
    {
        Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        string[] lines = payload.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n');
        foreach (string rawLine in lines)
        {
            string line = rawLine.Trim();
            if (line.Length == 0)
            {
                continue;
            }
            int separator = line.IndexOf('=');
            if (separator <= 0 || separator == line.Length - 1)
            {
                throw new InvalidDataException("The SQL credential payload is invalid.");
            }
            string key = line.Substring(0, separator).Trim();
            string value = line.Substring(separator + 1).Trim();
            if (values.ContainsKey(key))
            {
                throw new InvalidDataException("The SQL credential payload contains a duplicate key.");
            }
            values.Add(key, value);
        }

        string version;
        if (!values.TryGetValue("version", out version) || version != "1")
        {
            throw new InvalidDataException("The SQL credential payload version is unsupported.");
        }
        return values;
    }

    private static string Decode(Dictionary<string, string> values, string key)
    {
        string encoded;
        if (!values.TryGetValue(key, out encoded))
        {
            throw new InvalidDataException("The SQL credential payload is missing " + key + ".");
        }
        string value;
        try
        {
            value = Encoding.UTF8.GetString(Convert.FromBase64String(encoded));
        }
        catch (FormatException ex)
        {
            throw new InvalidDataException("The SQL credential payload has invalid encoding for " + key + ".", ex);
        }
        if (String.IsNullOrWhiteSpace(value))
        {
            throw new InvalidDataException("The SQL credential payload has an empty " + key + ".");
        }
        return value;
    }
}
