using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Security.Cryptography;
using System.Text;

internal sealed class RecoSqlCredential
{
    internal string Server;
    internal string User;
    internal string Password;
}

internal static class RecoSqlCredentialStore
{
    private const string StoreEnvironmentVariable = "RECO_SQL_CREDENTIAL_STORE_PATH";
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

    internal static RecoSqlCredential Read(string name)
    {
        string normalizedName = (name ?? "").Trim().ToLowerInvariant();
        if (normalizedName != "learning" && normalizedName != "business")
        {
            throw new ArgumentException("Unknown SQL credential entry.", "name");
        }

        string path = GetStorePath();
        if (!File.Exists(path))
        {
            throw new FileNotFoundException("The local DPAPI SQL credential store is missing.", path);
        }
        if ((File.GetAttributes(path) & FileAttributes.ReparsePoint) != 0)
        {
            throw new InvalidOperationException("The local DPAPI SQL credential store cannot be a reparse point.");
        }

        byte[] encrypted = File.ReadAllBytes(path);
        byte[] plaintext = null;
        try
        {
            plaintext = ProtectedData.Unprotect(encrypted, Entropy, DataProtectionScope.CurrentUser);
            Dictionary<string, string> values = ParsePayload(Encoding.UTF8.GetString(plaintext));
            string prefix = normalizedName + ".";
            return new RecoSqlCredential
            {
                Server = Decode(values, prefix + "server"),
                User = Decode(values, prefix + "user"),
                Password = Decode(values, prefix + "password")
            };
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

    internal static string BuildConnectionString(string name, string database, int port, int timeoutSeconds)
    {
        if (String.IsNullOrWhiteSpace(database))
        {
            throw new ArgumentException("Database is required.", "database");
        }

        RecoSqlCredential credential = Read(name);
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
                throw new InvalidDataException("The local DPAPI SQL credential store has an invalid payload.");
            }
            string key = line.Substring(0, separator).Trim();
            string value = line.Substring(separator + 1).Trim();
            if (values.ContainsKey(key))
            {
                throw new InvalidDataException("The local DPAPI SQL credential store contains a duplicate key.");
            }
            values.Add(key, value);
        }

        string version;
        if (!values.TryGetValue("version", out version) || version != "1")
        {
            throw new InvalidDataException("The local DPAPI SQL credential store version is unsupported.");
        }
        return values;
    }

    private static string Decode(Dictionary<string, string> values, string key)
    {
        string encoded;
        if (!values.TryGetValue(key, out encoded))
        {
            throw new InvalidDataException("The local DPAPI SQL credential store is missing " + key + ".");
        }
        string value;
        try
        {
            value = Encoding.UTF8.GetString(Convert.FromBase64String(encoded));
        }
        catch (FormatException ex)
        {
            throw new InvalidDataException("The local DPAPI SQL credential store has invalid encoding for " + key + ".", ex);
        }
        if (String.IsNullOrWhiteSpace(value))
        {
            throw new InvalidDataException("The local DPAPI SQL credential store has an empty " + key + ".");
        }
        return value;
    }
}
