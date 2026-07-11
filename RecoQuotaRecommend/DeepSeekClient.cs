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
    internal sealed class DeepSeekSettings
    {
        public bool Enabled;
        public string ApiKey;
        public string Model = "deepseek-v4-pro";
        public string BaseUrl = "https://api.deepseek.com";
        public int TimeoutSeconds = 20;
        public int MaxRowsPerBatch = 8;
        public int MaxCandidatesPerRow = 12;
        public int LocalHighScore = 80;
        public int DisplayConfidence = 65;
        public int AutoCheckConfidence = 85;
        public bool EnableNameNormalization = true;
        public bool EnableMappingDetection = true;
        public bool EnableQuotaRecommendation = true;
        public bool EnableColumnDetection = true;

        public bool IsAvailable
        {
            get { return !String.IsNullOrWhiteSpace(ApiKey); }
        }

        public bool CanNormalizeNames
        {
            get { return IsAvailable && EnableNameNormalization; }
        }

        public bool CanDetectColumns
        {
            get { return IsAvailable && EnableColumnDetection; }
        }

        public bool CanRecommendQuota
        {
            get { return IsAvailable && EnableQuotaRecommendation; }
        }

        public bool CanDetectMapping
        {
            get { return IsAvailable; }
        }

        public DeepSeekSettings Copy()
        {
            return new DeepSeekSettings
            {
                Enabled = Enabled,
                ApiKey = ApiKey,
                Model = Model,
                BaseUrl = BaseUrl,
                TimeoutSeconds = TimeoutSeconds,
                MaxRowsPerBatch = MaxRowsPerBatch,
                MaxCandidatesPerRow = MaxCandidatesPerRow,
                LocalHighScore = LocalHighScore,
                DisplayConfidence = DisplayConfidence,
                AutoCheckConfidence = AutoCheckConfidence,
                EnableNameNormalization = EnableNameNormalization,
                EnableMappingDetection = EnableMappingDetection,
                EnableQuotaRecommendation = EnableQuotaRecommendation,
                EnableColumnDetection = EnableColumnDetection
            };
        }

        public static DeepSeekSettings Load()
        {
            DeepSeekSettings settings = new DeepSeekSettings();
            string path = ConfigPath();
            if (!File.Exists(path))
            {
                return settings;
            }

            try
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                Dictionary<string, object> values = serializer.DeserializeObject(File.ReadAllText(path, Encoding.UTF8)) as Dictionary<string, object>;
                if (values == null)
                {
                    return settings;
                }

                settings.Enabled = ReadBool(values, "enabled", false);
                settings.ApiKey = ReadString(values, "api_key", "");
                settings.Model = ReadString(values, "model", settings.Model);
                settings.BaseUrl = ReadString(values, "base_url", settings.BaseUrl).TrimEnd('/');
                settings.TimeoutSeconds = Clamp(ReadInt(values, "timeout_seconds", settings.TimeoutSeconds), 2, 120);
                settings.MaxRowsPerBatch = Clamp(ReadInt(values, "max_rows_per_batch", settings.MaxRowsPerBatch), 1, 20);
                settings.MaxCandidatesPerRow = Clamp(ReadInt(values, "max_candidates_per_row", settings.MaxCandidatesPerRow), 3, 20);
                settings.LocalHighScore = Clamp(ReadInt(values, "local_high_score", settings.LocalHighScore), 60, 120);
                settings.DisplayConfidence = Clamp(ReadInt(values, "display_confidence", settings.DisplayConfidence), 1, 100);
                settings.AutoCheckConfidence = Clamp(ReadInt(values, "auto_check_confidence", settings.AutoCheckConfidence), 1, 100);
                settings.EnableNameNormalization = ReadBool(values, "enable_name_normalization", settings.EnableNameNormalization);
                settings.EnableMappingDetection = ReadBool(values, "enable_mapping_detection", settings.EnableMappingDetection);
                settings.EnableQuotaRecommendation = ReadBool(values, "enable_quota_recommendation", settings.EnableQuotaRecommendation);
                settings.EnableColumnDetection = ReadBool(values, "enable_column_detection", settings.EnableColumnDetection);
            }
            catch (Exception ex)
            {
                QuotaRecommendPanel.Log("Load DeepSeek settings failed: " + ex.Message);
                settings.Enabled = false;
            }

            return settings;
        }

        public void Save()
        {
            Directory.CreateDirectory(LearningStore.FindDataDir());
            string path = ConfigPath();
            StringBuilder builder = new StringBuilder();
            builder.AppendLine("{");
            AppendJson(builder, "enabled", Enabled ? "true" : "false", false, true);
            AppendJson(builder, "api_key", ApiKey ?? "", true, true);
            AppendJson(builder, "model", String.IsNullOrWhiteSpace(Model) ? "deepseek-v4-pro" : Model, true, true);
            AppendJson(builder, "base_url", String.IsNullOrWhiteSpace(BaseUrl) ? "https://api.deepseek.com" : BaseUrl, true, true);
            AppendJson(builder, "timeout_seconds", TimeoutSeconds.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "max_rows_per_batch", MaxRowsPerBatch.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "max_candidates_per_row", MaxCandidatesPerRow.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "local_high_score", LocalHighScore.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "display_confidence", DisplayConfidence.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "auto_check_confidence", AutoCheckConfidence.ToString(CultureInfo.InvariantCulture), false, true);
            AppendJson(builder, "enable_name_normalization", EnableNameNormalization ? "true" : "false", false, true);
            AppendJson(builder, "enable_mapping_detection", EnableMappingDetection ? "true" : "false", false, true);
            AppendJson(builder, "enable_quota_recommendation", EnableQuotaRecommendation ? "true" : "false", false, true);
            AppendJson(builder, "enable_column_detection", EnableColumnDetection ? "true" : "false", false, false);
            builder.AppendLine("}");
            File.WriteAllText(path, builder.ToString(), Encoding.UTF8);
        }

        public static string ConfigPath()
        {
            return Path.Combine(LearningStore.FindDataDir(), "deepseek-settings.json");
        }

        private static void AppendJson(StringBuilder builder, string key, string value, bool quoteValue, bool comma)
        {
            builder.Append("  \"").Append(EscapeJson(key)).Append("\": ");
            if (quoteValue)
            {
                builder.Append("\"").Append(EscapeJson(value)).Append("\"");
            }
            else
            {
                builder.Append(value);
            }
            if (comma)
            {
                builder.Append(",");
            }
            builder.AppendLine();
        }

        private static string EscapeJson(string value)
        {
            StringBuilder builder = new StringBuilder();
            foreach (char ch in value ?? "")
            {
                switch (ch)
                {
                    case '\\':
                        builder.Append("\\\\");
                        break;
                    case '"':
                        builder.Append("\\\"");
                        break;
                    case '\r':
                        builder.Append("\\r");
                        break;
                    case '\n':
                        builder.Append("\\n");
                        break;
                    case '\t':
                        builder.Append("\\t");
                        break;
                    default:
                        builder.Append(ch);
                        break;
                }
            }
            return builder.ToString();
        }

        private static string ReadString(Dictionary<string, object> values, string key, string fallback)
        {
            object value;
            return values.TryGetValue(key, out value) && value != null ? Convert.ToString(value, CultureInfo.InvariantCulture) : fallback;
        }

        private static int ReadInt(Dictionary<string, object> values, string key, int fallback)
        {
            object value;
            if (!values.TryGetValue(key, out value) || value == null)
            {
                return fallback;
            }

            int parsed;
            return Int32.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) ? parsed : fallback;
        }

        private static bool ReadBool(Dictionary<string, object> values, string key, bool fallback)
        {
            object value;
            if (!values.TryGetValue(key, out value) || value == null)
            {
                return fallback;
            }

            if (value is bool)
            {
                return (bool)value;
            }

            bool parsed;
            return Boolean.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), out parsed) ? parsed : fallback;
        }

        private static int Clamp(int value, int min, int max)
        {
            return Math.Max(min, Math.Min(max, value));
        }
    }

    internal sealed class DeepSeekSettingsDialog : Form
    {
        private readonly CheckBox enabledCheck;
        private readonly CheckBox normalizeNameCheck;
        private readonly CheckBox mappingDetectCheck;
        private readonly CheckBox recommendQuotaCheck;
        private readonly TextBox apiKeyText;
        private readonly CheckBox showKeyCheck;
        private readonly TextBox modelText;
        private readonly TextBox baseUrlText;
        private readonly NumericUpDown timeoutInput;
        private readonly NumericUpDown batchInput;
        private readonly NumericUpDown candidatesInput;
        private readonly NumericUpDown localHighScoreInput;
        private readonly NumericUpDown displayConfidenceInput;
        private readonly NumericUpDown autoCheckConfidenceInput;

        public DeepSeekSettings Settings { get; private set; }

        public DeepSeekSettingsDialog(DeepSeekSettings current)
        {
            Settings = (current ?? new DeepSeekSettings()).Copy();
            Text = "DeepSeek AI\u8bbe\u7f6e";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ClientSize = new Size(560, 300);

            enabledCheck = new CheckBox { Checked = true };
            normalizeNameCheck = new CheckBox { Checked = true };
            mappingDetectCheck = new CheckBox { Checked = true };
            recommendQuotaCheck = new CheckBox { Checked = true };
            batchInput = HiddenNumber(Settings.MaxRowsPerBatch, 1, 20);
            localHighScoreInput = HiddenNumber(Settings.LocalHighScore, 60, 120);
            displayConfidenceInput = HiddenNumber(Settings.DisplayConfidence, 1, 100);

            AddLabel("API Key", 24, 28, 150);
            apiKeyText = new TextBox();
            apiKeyText.Left = 180;
            apiKeyText.Top = 24;
            apiKeyText.Width = 280;
            apiKeyText.UseSystemPasswordChar = true;
            apiKeyText.Text = Settings.ApiKey ?? "";
            Controls.Add(apiKeyText);

            showKeyCheck = new CheckBox();
            showKeyCheck.Left = 470;
            showKeyCheck.Top = 26;
            showKeyCheck.Width = 60;
            showKeyCheck.Text = "\u663e\u793a";
            showKeyCheck.CheckedChanged += delegate { apiKeyText.UseSystemPasswordChar = !showKeyCheck.Checked; };
            Controls.Add(showKeyCheck);

            AddLabel("\u6a21\u578b", 24, 64, 150);
            modelText = AddTextBox(Settings.Model, 180, 60, 280);

            AddLabel("\u63a5\u53e3\u5730\u5740", 24, 100, 150);
            baseUrlText = AddTextBox(Settings.BaseUrl, 180, 96, 280);

            AddLabel("\u8d85\u65f6\u79d2\u6570", 24, 136, 150);
            timeoutInput = AddNumber(Settings.TimeoutSeconds, 180, 132, 2, 120);

            AddLabel("\u6bcf\u884c\u5019\u9009\u6570", 24, 172, 150);
            candidatesInput = AddNumber(Settings.MaxCandidatesPerRow, 180, 168, 3, 20);

            AddLabel("AI\u81ea\u52a8\u52fe\u9009\u7f6e\u4fe1\u5ea6", 24, 208, 150);
            autoCheckConfidenceInput = AddNumber(Settings.AutoCheckConfidence, 180, 204, 1, 100);

            Button saveButton = new Button();
            saveButton.Text = "\u4fdd\u5b58";
            saveButton.Left = 366;
            saveButton.Top = 252;
            saveButton.Width = 80;
            saveButton.DialogResult = DialogResult.None;
            saveButton.Click += delegate { SaveSettings(); };
            Controls.Add(saveButton);

            Button cancelButton = new Button();
            cancelButton.Text = "\u53d6\u6d88";
            cancelButton.Left = 456;
            cancelButton.Top = 252;
            cancelButton.Width = 80;
            cancelButton.DialogResult = DialogResult.Cancel;
            Controls.Add(cancelButton);

            AcceptButton = saveButton;
            CancelButton = cancelButton;
        }

        private void SaveSettings()
        {
            string model = modelText.Text.Trim();
            string baseUrl = baseUrlText.Text.Trim();
            string apiKey = apiKeyText.Text.Trim();

            if (String.IsNullOrWhiteSpace(apiKey))
            {
                MessageBox.Show(this, "\u8bf7\u586b\u5199 DeepSeek API Key\u3002", "DeepSeek AI\u8bbe\u7f6e", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (String.IsNullOrWhiteSpace(model))
            {
                model = "deepseek-v4-pro";
            }
            if (String.IsNullOrWhiteSpace(baseUrl))
            {
                baseUrl = "https://api.deepseek.com";
            }

            DeepSeekSettings updated = new DeepSeekSettings();
            updated.Enabled = !String.IsNullOrWhiteSpace(apiKey);
            updated.ApiKey = apiKey;
            updated.Model = model;
            updated.BaseUrl = baseUrl.TrimEnd('/');
            updated.TimeoutSeconds = Convert.ToInt32(timeoutInput.Value);
            updated.MaxRowsPerBatch = Convert.ToInt32(batchInput.Value);
            updated.MaxCandidatesPerRow = Convert.ToInt32(candidatesInput.Value);
            updated.LocalHighScore = Convert.ToInt32(localHighScoreInput.Value);
            updated.DisplayConfidence = Convert.ToInt32(displayConfidenceInput.Value);
            updated.AutoCheckConfidence = Convert.ToInt32(autoCheckConfidenceInput.Value);
            updated.EnableNameNormalization = true;
            updated.EnableMappingDetection = true;
            updated.EnableQuotaRecommendation = true;

            try
            {
                updated.Save();
                Settings = updated;
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "\u4fdd\u5b58 DeepSeek \u8bbe\u7f6e\u5931\u8d25\uff1a" + ex.Message, "DeepSeek AI\u8bbe\u7f6e", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private TextBox AddTextBox(string text, int left, int top, int width)
        {
            TextBox box = new TextBox();
            box.Left = left;
            box.Top = top;
            box.Width = width;
            box.Text = text ?? "";
            Controls.Add(box);
            return box;
        }

        private NumericUpDown AddNumber(int value, int left, int top, int min, int max)
        {
            NumericUpDown input = new NumericUpDown();
            input.Left = left;
            input.Top = top;
            input.Width = 120;
            input.Minimum = min;
            input.Maximum = max;
            input.Value = Math.Max(min, Math.Min(max, value));
            Controls.Add(input);
            return input;
        }

        private NumericUpDown HiddenNumber(int value, int min, int max)
        {
            NumericUpDown input = new NumericUpDown();
            input.Minimum = min;
            input.Maximum = max;
            input.Value = Math.Max(min, Math.Min(max, value));
            return input;
        }

        private void AddLabel(string text, int left, int top, int width)
        {
            Label label = new Label();
            label.Text = text;
            label.Left = left;
            label.Top = top + 4;
            label.Width = width;
            Controls.Add(label);
        }
    }

    internal sealed class DeepSeekClient
    {
        private readonly DeepSeekSettings settings;
        private readonly JavaScriptSerializer serializer = new JavaScriptSerializer();

        public DeepSeekClient(DeepSeekSettings deepSeekSettings)
        {
            settings = deepSeekSettings;
            serializer.MaxJsonLength = 1024 * 1024 * 4;
        }

        public List<DeepSeekSelection> Rank(List<DeepSeekRequestRow> rows)
        {
            if (rows == null || rows.Count == 0)
            {
                return new List<DeepSeekSelection>();
            }

            return ParseResponse(SendRequest(BuildUnifiedRequestJson(rows)));
        }

        public List<DeepSeekNameResult> NormalizeQuantityNames(List<DeepSeekNameRequestRow> rows)
        {
            if (rows == null || rows.Count == 0)
            {
                return new List<DeepSeekNameResult>();
            }

            return ParseNameResponse(SendRequest(BuildNameRequestJson(rows)));
        }

        // 对整张工程量表只调用一次，让模型判断列结构（哪些列组成名称、哪列单位、哪列数量）。
        public DeepSeekColumnLayout DetectColumnLayout(List<List<CellValue>> rows)
        {
            if (rows == null || rows.Count == 0)
            {
                return null;
            }

            return ParseColumnLayoutResponse(SendRequest(BuildColumnLayoutRequestJson(rows)));
        }

        private string SendRequest(string requestJson)
        {
            ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol | (SecurityProtocolType)3072;
            string endpoint = settings.BaseUrl.TrimEnd('/');
            if (!endpoint.EndsWith("/chat/completions", StringComparison.OrdinalIgnoreCase))
            {
                endpoint += "/chat/completions";
            }

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(endpoint);
            request.Method = "POST";
            request.ContentType = "application/json";
            request.Accept = "application/json";
            request.Headers["Authorization"] = "Bearer " + settings.ApiKey;
            request.Timeout = settings.TimeoutSeconds * 1000;
            request.ReadWriteTimeout = settings.TimeoutSeconds * 1000;

            byte[] payload = Encoding.UTF8.GetBytes(requestJson);
            request.ContentLength = payload.Length;
            using (Stream stream = request.GetRequestStream())
            {
                stream.Write(payload, 0, payload.Length);
            }

            using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
            using (Stream stream = response.GetResponseStream())
            using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
            {
                return reader.ReadToEnd();
            }
        }

        private string BuildNameRequestJson(List<DeepSeekNameRequestRow> rows)
        {
            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["model"] = settings.Model;
            payload["stream"] = false;
            payload["temperature"] = 0.1;
            payload["max_tokens"] = 1400;
            payload["response_format"] = new Dictionary<string, object> { { "type", "json_object" } };

            List<object> messages = new List<object>();
            messages.Add(new Dictionary<string, object>
            {
                { "role", "system" },
                { "content", "You normalize Chinese railway construction quantity names. Return strict JSON: {\"results\":[{\"row_id\":\"n0\",\"quantity_name\":\"标准工程量名称\",\"confidence\":90,\"reason\":\"short reason\"}]}. The quantity_name must be concise, contain the engineering object and key specification, and must not include serial numbers, units, quantities, formulas, or prices." }
            });
            messages.Add(new Dictionary<string, object>
            {
                { "role", "user" },
                { "content", BuildNameUserPrompt(rows) }
            });
            payload["messages"] = messages;
            return serializer.Serialize(payload);
        }

        private string BuildNameUserPrompt(List<DeepSeekNameRequestRow> rows)
        {
            List<object> requestRows = new List<object>();
            foreach (DeepSeekNameRequestRow row in rows)
            {
                ExcelQuantityItem item = row.Item;
                requestRows.Add(new Dictionary<string, object>
                {
                    { "row_id", row.RowId ?? "" },
                    { "local_name", item == null ? "" : item.OriginalName ?? item.Name ?? "" },
                    { "section_name", item == null ? "" : item.SectionName ?? "" },
                    { "unit", item == null ? "" : item.Unit ?? "" },
                    { "quantity_value", item == null ? "" : item.ValueText ?? "" },
                    { "raw_row_text", item == null ? "" : Truncate(item.RawRowText, 300) },
                    { "context_text", item == null ? "" : Truncate(item.ContextText, 300) }
                });
            }

            Dictionary<string, object> body = new Dictionary<string, object>();
            body["task"] = "For each row, summarize all possible quantity-name columns into one standard Chinese engineering quantity name.";
            body["rules"] = new string[]
            {
                "Use all name-like columns before the unit/quantity as context.",
                "Do not include unit, quantity, formula, serial number, or price.",
                "Keep key material/model/specification when it changes the quota selection.",
                "If the row is ambiguous, return the best concise name with lower confidence."
            };
            body["rows"] = requestRows;
            return serializer.Serialize(body);
        }

        private string BuildUnifiedRequestJson(List<DeepSeekRequestRow> rows)
        {
            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["model"] = settings.Model;
            payload["stream"] = false;
            payload["temperature"] = 0.1;
            payload["max_tokens"] = 1400;
            payload["response_format"] = new Dictionary<string, object> { { "type", "json_object" } };

            List<object> messages = new List<object>();
            messages.Add(new Dictionary<string, object>
            {
                { "role", "system" },
                { "content", "You are a railway quota recommendation assistant. For each row, first decide whether an existing human-corrected mapping box matches. If yes, return selected_box_id. If no mapping box is reliable, choose one quota from quota_candidates. Never invent ids. Return strict JSON: {\"results\":[{\"row_id\":\"r0\",\"selected_box_id\":\"box-123\",\"selected_code\":\"\",\"confidence\":90,\"reason\":\"short reason\"}]}." }
            });
            messages.Add(new Dictionary<string, object>
            {
                { "role", "user" },
                { "content", BuildUnifiedUserPrompt(rows) }
            });
            payload["messages"] = messages;
            return serializer.Serialize(payload);
        }

        private string BuildUnifiedUserPrompt(List<DeepSeekRequestRow> rows)
        {
            List<object> requestRows = new List<object>();
            foreach (DeepSeekRequestRow row in rows)
            {
                List<object> quotaCandidates = new List<object>();
                foreach (AiQuotaCandidate candidate in row.Candidates ?? new List<AiQuotaCandidate>())
                {
                    if (candidate == null || candidate.Quota == null)
                    {
                        continue;
                    }

                    quotaCandidates.Add(new Dictionary<string, object>
                    {
                        { "code", candidate.Quota.QuotaCode ?? "" },
                        { "name", candidate.Quota.QuotaName ?? "" },
                        { "unit", candidate.Quota.QuotaUnit ?? "" },
                        { "work_content", Truncate(candidate.Quota.WorkContent, 160) },
                        { "local_score", candidate.LocalScore }
                    });
                }

                List<object> mappingCandidates = new List<object>();
                foreach (AiMappingCandidate candidate in row.MappingCandidates ?? new List<AiMappingCandidate>())
                {
                    if (candidate == null)
                    {
                        continue;
                    }

                    mappingCandidates.Add(new Dictionary<string, object>
                    {
                        { "box_id", candidate.BoxId ?? "" },
                        { "sample_quantity_names", Truncate(candidate.SampleNames, 220) },
                        { "targets", String.Join(" + ", (candidate.Targets ?? new List<MappingTarget>()).Select(t => (t.Code ?? "") + " " + (t.Name ?? "")).ToArray()) },
                        { "local_score", candidate.LocalScore }
                    });
                }

                requestRows.Add(new Dictionary<string, object>
                {
                    { "row_id", row.RowId ?? "" },
                    { "quantity_name", row.Item == null ? "" : row.Item.Name ?? "" },
                    { "original_name", row.Item == null ? "" : row.Item.OriginalName ?? "" },
                    { "quantity_unit", row.Item == null ? "" : row.Item.Unit ?? "" },
                    { "quantity_value", row.Item == null ? "" : row.Item.ValueText ?? "" },
                    { "raw_row_text", row.Item == null ? "" : Truncate(row.Item.RawRowText, 260) },
                    { "mapping_candidates", mappingCandidates },
                    { "quota_candidates", quotaCandidates }
                });
            }

            Dictionary<string, object> body = new Dictionary<string, object>();
            body["task"] = "For each quantity row, first match mapping_candidates; if no reliable mapping box exists, select the best quota from quota_candidates.";
            body["rules"] = new string[]
            {
                "selected_box_id must exactly equal a mapping candidate box_id, otherwise leave it empty.",
                "selected_code must exactly equal a quota candidate code, otherwise leave it empty.",
                "Prefer a reliable human-corrected mapping box over a single quota.",
                "Do not choose a mapping box from generic one-word similarity only.",
                "Do not confuse steel quantities with concrete structure quotas."
            };
            body["rows"] = requestRows;
            return serializer.Serialize(body);
        }

        private List<DeepSeekSelection> ParseResponse(string responseJson)
        {
            object rootObject = serializer.DeserializeObject(responseJson);
            Dictionary<string, object> root = rootObject as Dictionary<string, object>;
            if (root == null)
            {
                return new List<DeepSeekSelection>();
            }

            List<object> choices = GetList(root, "choices");
            if (choices == null || choices.Count == 0)
            {
                return new List<DeepSeekSelection>();
            }

            Dictionary<string, object> firstChoice = choices[0] as Dictionary<string, object>;
            Dictionary<string, object> message = firstChoice == null ? null : firstChoice.ContainsKey("message") ? firstChoice["message"] as Dictionary<string, object> : null;
            string content = message == null || !message.ContainsKey("content") ? "" : Convert.ToString(message["content"], CultureInfo.InvariantCulture);
            if (String.IsNullOrWhiteSpace(content))
            {
                return new List<DeepSeekSelection>();
            }

            Dictionary<string, object> resultRoot = serializer.DeserializeObject(content) as Dictionary<string, object>;
            if (resultRoot == null)
            {
                return new List<DeepSeekSelection>();
            }

            List<object> results = GetList(resultRoot, "results");
            if (results == null && resultRoot.ContainsKey("row_id"))
            {
                results = new List<object> { resultRoot };
            }

            List<DeepSeekSelection> selections = new List<DeepSeekSelection>();
            foreach (object item in results ?? new List<object>())
            {
                Dictionary<string, object> row = item as Dictionary<string, object>;
                if (row == null)
                {
                    continue;
                }

                selections.Add(new DeepSeekSelection
                {
                    RowId = ReadString(row, "row_id"),
                    BoxId = String.IsNullOrWhiteSpace(ReadString(row, "selected_box_id")) ? ReadString(row, "box_id") : ReadString(row, "selected_box_id"),
                    SelectedCode = ReadString(row, "selected_code"),
                    Confidence = ReadInt(row, "confidence"),
                    Reason = ReadString(row, "reason")
                });
            }

            return selections;
        }

        private List<DeepSeekNameResult> ParseNameResponse(string responseJson)
        {
            object rootObject = serializer.DeserializeObject(responseJson);
            Dictionary<string, object> root = rootObject as Dictionary<string, object>;
            if (root == null)
            {
                return new List<DeepSeekNameResult>();
            }

            List<object> choices = GetList(root, "choices");
            if (choices == null || choices.Count == 0)
            {
                return new List<DeepSeekNameResult>();
            }

            Dictionary<string, object> firstChoice = choices[0] as Dictionary<string, object>;
            Dictionary<string, object> message = firstChoice == null || !firstChoice.ContainsKey("message") ? null : firstChoice["message"] as Dictionary<string, object>;
            string content = message == null || !message.ContainsKey("content") ? "" : Convert.ToString(message["content"], CultureInfo.InvariantCulture);
            if (String.IsNullOrWhiteSpace(content))
            {
                return new List<DeepSeekNameResult>();
            }

            Dictionary<string, object> resultRoot = serializer.DeserializeObject(content) as Dictionary<string, object>;
            if (resultRoot == null)
            {
                return new List<DeepSeekNameResult>();
            }

            List<object> results = GetList(resultRoot, "results");
            if (results == null && resultRoot.ContainsKey("row_id"))
            {
                results = new List<object> { resultRoot };
            }

            List<DeepSeekNameResult> normalized = new List<DeepSeekNameResult>();
            foreach (object item in results ?? new List<object>())
            {
                Dictionary<string, object> row = item as Dictionary<string, object>;
                if (row == null)
                {
                    continue;
                }

                normalized.Add(new DeepSeekNameResult
                {
                    RowId = ReadString(row, "row_id"),
                    QuantityName = ReadString(row, "quantity_name"),
                    Confidence = ReadInt(row, "confidence"),
                    Reason = ReadString(row, "reason")
                });
            }

            return normalized;
        }

        private string BuildColumnLayoutRequestJson(List<List<CellValue>> rows)
        {
            int columnCount = 0;
            foreach (List<CellValue> row in rows)
            {
                if (row == null)
                {
                    continue;
                }

                foreach (CellValue cell in row)
                {
                    if (cell != null && cell.SourceIndex > columnCount)
                    {
                        columnCount = cell.SourceIndex;
                    }
                }
            }

            List<object> sampleRows = new List<object>();
            int taken = 0;
            foreach (List<CellValue> row in rows)
            {
                if (row == null || row.Count == 0)
                {
                    continue;
                }

                Dictionary<string, object> cells = new Dictionary<string, object>();
                foreach (CellValue cell in row)
                {
                    if (cell == null)
                    {
                        continue;
                    }

                    string text = Truncate(cell.Text, 40);
                    if (String.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    cells[cell.SourceIndex.ToString(CultureInfo.InvariantCulture)] = text;
                }

                if (cells.Count == 0)
                {
                    continue;
                }

                sampleRows.Add(cells);
                taken++;
                if (taken >= 12)
                {
                    break;
                }
            }

            Dictionary<string, object> body = new Dictionary<string, object>();
            body["column_count"] = columnCount;
            body["task"] = "判断该工程量表的列结构：哪些列共同组成工程量名称(name_columns，按从左到右顺序)，哪一列是单位(unit_column)，哪一列是数量(quantity_column)。";
            body["rules"] = new string[]
            {
                "列索引从1开始；name_columns 应包含数量列左侧所有描述/名称类列（如部位、项目名称、规格型号等），不含序号列。",
                "unit_column 为单位列（m、m2、m3、kg、t、处、个、座、项等）；若无单位列填0。",
                "quantity_column 为数量/工程量数值列。",
                "若无法判断填 confidence 0。"
            };
            body["rows"] = sampleRows;

            List<object> messages = new List<object>();
            messages.Add(new Dictionary<string, object>
            {
                { "role", "system" },
                { "content", "你是工程量表结构识别助手。根据给定的数据行（按 列索引->文本 给出，可能含表头行），判断每列的角色。必须严格输出 JSON：{\"name_columns\":[2,3,4],\"unit_column\":5,\"quantity_column\":6,\"confidence\":90}。列索引从1开始。" }
            });
            messages.Add(new Dictionary<string, object>
            {
                { "role", "user" },
                { "content", serializer.Serialize(body) }
            });

            Dictionary<string, object> payload = new Dictionary<string, object>();
            payload["model"] = settings.Model;
            payload["stream"] = false;
            payload["temperature"] = 0.1;
            payload["max_tokens"] = 400;
            payload["response_format"] = new Dictionary<string, object> { { "type", "json_object" } };
            payload["messages"] = messages;
            return serializer.Serialize(payload);
        }

        private DeepSeekColumnLayout ParseColumnLayoutResponse(string responseJson)
        {
            Dictionary<string, object> root = serializer.DeserializeObject(responseJson) as Dictionary<string, object>;
            if (root == null)
            {
                return null;
            }

            List<object> choices = GetList(root, "choices");
            if (choices == null || choices.Count == 0)
            {
                return null;
            }

            Dictionary<string, object> firstChoice = choices[0] as Dictionary<string, object>;
            Dictionary<string, object> message = firstChoice == null || !firstChoice.ContainsKey("message") ? null : firstChoice["message"] as Dictionary<string, object>;
            string content = message == null || !message.ContainsKey("content") ? "" : Convert.ToString(message["content"], CultureInfo.InvariantCulture);
            if (String.IsNullOrWhiteSpace(content))
            {
                return null;
            }

            Dictionary<string, object> resultRoot = serializer.DeserializeObject(content) as Dictionary<string, object>;
            if (resultRoot == null)
            {
                return null;
            }

            List<int> names = new List<int>();
            foreach (object o in GetList(resultRoot, "name_columns") ?? new List<object>())
            {
                int v;
                if (Int32.TryParse(Convert.ToString(o, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out v) && v > 0)
                {
                    names.Add(v);
                }
            }

            return new DeepSeekColumnLayout
            {
                NameColumns = names.ToArray(),
                UnitColumn = ReadInt(resultRoot, "unit_column"),
                QuantityColumn = ReadInt(resultRoot, "quantity_column"),
                Confidence = ReadInt(resultRoot, "confidence")
            };
        }

        private static List<object> GetList(Dictionary<string, object> values, string key)
        {
            object value;
            if (!values.TryGetValue(key, out value) || value == null)
            {
                return null;
            }

            ArrayList arrayList = value as ArrayList;
            if (arrayList != null)
            {
                return arrayList.Cast<object>().ToList();
            }

            object[] objectArray = value as object[];
            if (objectArray != null)
            {
                return objectArray.ToList();
            }

            return null;
        }

        private static string ReadString(Dictionary<string, object> values, string key)
        {
            object value;
            return values.TryGetValue(key, out value) && value != null ? Convert.ToString(value, CultureInfo.InvariantCulture) : "";
        }

        private static int ReadInt(Dictionary<string, object> values, string key)
        {
            object value;
            if (!values.TryGetValue(key, out value) || value == null)
            {
                return 0;
            }

            int parsed;
            return Int32.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture), NumberStyles.Integer, CultureInfo.InvariantCulture, out parsed) ? parsed : 0;
        }

        private static string Truncate(string text, int maxLength)
        {
            string value = (text ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
            return value.Length <= maxLength ? value : value.Substring(0, maxLength);
        }
    }
}
