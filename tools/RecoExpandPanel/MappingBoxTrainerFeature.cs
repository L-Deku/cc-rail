using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace RecoNet
{
    public partial class FormPanel : Form
    {
        private static readonly Dictionary<Form, MappingBoxTrainerDialog> MappingBoxTrainerDialogs = new Dictionary<Form, MappingBoxTrainerDialog>();

        private static void ShowMappingBoxTrainerPanel(Form mainForm)
        {
            try
            {
                MappingBoxTrainerDialog dialog;
                if (!MappingBoxTrainerDialogs.TryGetValue(mainForm, out dialog) || dialog == null || dialog.IsDisposed)
                {
                    dialog = new MappingBoxTrainerDialog(mainForm);
                    MappingBoxTrainerDialogs[mainForm] = dialog;
                    dialog.FormClosed += delegate { MappingBoxTrainerDialogs.Remove(mainForm); };
                    dialog.Show(mainForm);
                }
                else
                {
                    dialog.Show();
                }

                dialog.Activate();
            }
            catch (Exception ex)
            {
                Log("Show mapping box trainer failed: " + ex);
                MessageBox.Show(mainForm, "\u6253\u5f00\u5bf9\u5e94\u6846\u8bad\u7ec3\u7a97\u53e3\u5931\u8d25\uff1a" + ex.Message, "\u6dfb\u52a0\u5bf9\u5e94\u6846\u5185\u5bb9", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static List<MappingTrainerTarget> BuildMappingTrainerTargets(DataGridView grid)
        {
            List<MappingTrainerTarget> targets = new List<MappingTrainerTarget>();
            int index = 0;
            foreach (DataGridViewRow row in GetSelectedQuotaRows(grid))
            {
                string code = GetRowValue(row, "\u5b9a\u989d\u7f16\u53f7", "\u5b9a\u989d\u7f16\u53f7DE", "\u7f16\u53f7");
                if (String.IsNullOrWhiteSpace(code))
                {
                    continue;
                }

                MappingTrainerTarget target = new MappingTrainerTarget();
                target.TargetId = "t" + index.ToString(CultureInfo.InvariantCulture);
                target.TargetKind = GuessMappingTargetKind(code);
                target.Code = code.Trim();
                target.Name = GetRowValue(row, "\u5de5\u7a0b\u6216\u8d39\u7528\u9879\u76ee\u540d\u79f0", "\u540d\u79f0", "\u9879\u76ee\u540d\u79f0");
                target.Unit = GetRowValue(row, "\u5355\u4f4d", "\u5b9a\u989d\u5355\u4f4d");
                target.QuotaQuantity = GetRowValue(row, "\u5de5\u7a0b\u6570\u91cf\u8f93\u5165", "\u5de5\u7a0b\u6570\u91cf");
                targets.Add(target);
                index++;
            }

            return targets;
        }

        private static List<MappingTrainerQuantity> BuildMappingTrainerQuantities(AiExcelSelectionContext selection)
        {
            List<MappingTrainerQuantity> result = new List<MappingTrainerQuantity>();
            if (selection == null)
            {
                return result;
            }

            int index = 0;
            string currentGroup = "";
            foreach (IGrouping<int, AiExcelCell> group in selection.Cells.GroupBy(c => c.Row).OrderBy(g => g.Key))
            {
                List<AiExcelCell> cells = group.OrderBy(c => c.Column).ToList();
                AiExcelCell quantityCell = cells.Where(c => c.IsNumber).OrderByDescending(c => c.Column).FirstOrDefault();
                if (quantityCell == null)
                {
                    string groupText = BuildTrainerQuantityNameFromCells(cells, Int32.MaxValue);
                    if (!String.IsNullOrWhiteSpace(groupText))
                    {
                        currentGroup = groupText;
                    }
                    continue;
                }

                string name = BuildTrainerQuantityNameFromCells(cells, quantityCell.Column);
                if (!String.IsNullOrWhiteSpace(currentGroup) &&
                    !String.IsNullOrWhiteSpace(name) &&
                    name.IndexOf(currentGroup, StringComparison.Ordinal) < 0)
                {
                    name = currentGroup + " " + name;
                }

                if (String.IsNullOrWhiteSpace(name) && !cells.Any(c => !c.IsNumber))
                {
                    continue;
                }

                MappingTrainerQuantity quantity = new MappingTrainerQuantity();
                quantity.QuantityId = "q" + index.ToString(CultureInfo.InvariantCulture);
                quantity.ExcelRow = group.Key;
                quantity.CellAddress = quantityCell.Address;
                quantity.QuantityName = CleanTrainerText(name);
                quantity.OriginalName = quantity.QuantityName;
                quantity.Unit = PickTrainerUnit(cells, quantityCell.Column);
                quantity.QuantityValue = quantityCell.Text;
                quantity.RawText = (currentGroup + " " + String.Join(" ", cells.Select(c => c.Text).Where(t => !String.IsNullOrWhiteSpace(t)).ToArray())).Trim();
                result.Add(quantity);
                index++;
            }

            return result;
        }

        private static string BuildTrainerQuantityNameFromCells(List<AiExcelCell> rowCells, int quantityColumn)
        {
            List<string> parts = new List<string>();
            foreach (AiExcelCell cell in rowCells.OrderBy(c => c.Column))
            {
                if (cell.IsNumber || cell.Column == quantityColumn)
                {
                    continue;
                }

                string text = CleanTrainerText(cell.Text);
                if (String.IsNullOrWhiteSpace(text) || LooksLikeTrainerUnit(text) || LooksLikeTrainerHeader(text))
                {
                    continue;
                }

                parts.Add(text);
            }

            return CleanTrainerText(String.Join(" ", parts.Take(8).ToArray()));
        }

        private static string PickTrainerUnit(List<AiExcelCell> rowCells, int quantityColumn)
        {
            foreach (AiExcelCell cell in rowCells
                .Where(c => !c.IsNumber && Math.Abs(c.Column - quantityColumn) <= 3)
                .OrderBy(c => Math.Abs(c.Column - quantityColumn)))
            {
                string text = CleanTrainerText(cell.Text);
                if (LooksLikeTrainerUnit(text))
                {
                    return text;
                }
            }

            return "";
        }

        private static bool LooksLikeTrainerHeader(string text)
        {
            string value = NormalizeForSignature(text);
            return value == "\u5e8f\u53f7" ||
                value == "\u7f16\u53f7" ||
                value == "\u5355\u4f4d" ||
                value == "\u6570\u91cf" ||
                value == "\u5de5\u7a0b\u91cf" ||
                value == "\u5de5\u7a0b\u6570\u91cf";
        }

        private static bool LooksLikeTrainerUnit(string text)
        {
            string value = (text ?? "").Trim().ToLowerInvariant().Replace(" ", "");
            return value == "m" ||
                value == "m2" ||
                value == "m3" ||
                value == "kg" ||
                value == "t" ||
                value == "\u4e2a" ||
                value == "\u5ea7" ||
                value == "\u5904" ||
                value == "\u6839" ||
                value == "\u5957" ||
                value == "\u5757" ||
                value == "\u7c73" ||
                value == "\u5e73\u65b9\u7c73" ||
                value == "\u7acb\u65b9\u7c73" ||
                value == "\u5428";
        }

        private static string CleanTrainerText(string text)
        {
            string value = (text ?? "").Replace("\r", " ").Replace("\n", " ").Replace("\t", " ").Trim();
            while (value.Contains("  "))
            {
                value = value.Replace("  ", " ");
            }
            return value.Length > 120 ? value.Substring(0, 120).Trim() : value;
        }

        private static string GuessMappingTargetKind(string code)
        {
            string value = (code ?? "").Trim();
            return value.Length > 0 && value.All(Char.IsDigit) ? "material" : "quota";
        }

        private static int MappingTrainerTargetSortRank(string targetKind, string code)
        {
            string kind = String.IsNullOrWhiteSpace(targetKind) ? GuessMappingTargetKind(code) : targetKind;
            return String.Equals(kind, "quota", StringComparison.OrdinalIgnoreCase) ? 0 : 1;
        }

        private static string BuildMappingTrainerBoxId(MappingTrainerTarget target)
        {
            return BuildStableMappingBoxId(target == null ? "" : target.TargetKey);
        }

        private static void SaveMappingTrainerPairs(List<MappingTrainerPairRow> acceptedRows)
        {
            if (acceptedRows == null || acceptedRows.Count == 0)
            {
                return;
            }

            string dataDir = FindRecoQuotaDataDir();
            Directory.CreateDirectory(dataDir);
            string path = Path.Combine(dataDir, "mapping-boxes.jsonl");

            WithMappingBoxesLock(delegate
            {
            BackupMappingTrainerFile(path);

            List<Dictionary<string, string>> rows = new List<Dictionary<string, string>>();
            if (File.Exists(path))
            {
                foreach (string line in File.ReadAllLines(path, Encoding.UTF8))
                {
                    Dictionary<string, string> parsed = ParseFlatJson(line);
                    if (parsed.Count > 0)
                    {
                        rows.Add(parsed);
                    }
                }
            }

            string now = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            foreach (MappingTrainerPairRow pair in acceptedRows)
            {
                if (pair == null || pair.Target == null || pair.Quantity == null)
                {
                    continue;
                }

                string quantityName = CleanTrainerText(pair.Quantity.QuantityName);
                if (String.IsNullOrWhiteSpace(quantityName) || String.IsNullOrWhiteSpace(pair.Target.Code))
                {
                    continue;
                }

                string quantityUnit = CleanTrainerText(pair.Quantity.Unit);
                string sampleKey = NormalizeForSignature(quantityName) + "|" + NormalizeForSignature(quantityUnit);
                string boxId = FindExistingMappingBoxId(rows, pair.Target.TargetKey) ?? BuildMappingTrainerBoxId(pair.Target);
                Dictionary<string, string> existing = rows.FirstOrDefault(row =>
                    String.Equals(GetFlat(row, "box_id"), boxId, StringComparison.OrdinalIgnoreCase) &&
                    String.Equals(GetFlat(row, "target_kind") + ":" + GetFlat(row, "target_code").Trim().ToUpperInvariant(), pair.Target.TargetKey, StringComparison.OrdinalIgnoreCase) &&
                    String.Equals(NormalizeForSignature(GetFlat(row, "quantity_name")) + "|" + NormalizeForSignature(GetFlat(row, "quantity_unit")), sampleKey, StringComparison.OrdinalIgnoreCase));

                if (existing == null)
                {
                    existing = new Dictionary<string, string>();
                    rows.Add(existing);
                    existing["record_type"] = "mapping_box";
                    existing["box_id"] = boxId;
                    existing["target_kind"] = pair.Target.TargetKind;
                    existing["target_code"] = pair.Target.Code ?? "";
                    existing["target_name"] = pair.Target.Name ?? "";
                    existing["target_unit"] = pair.Target.Unit ?? "";
                    existing["quantity_name"] = quantityName;
                    existing["quantity_unit"] = quantityUnit;
                    existing["weight"] = "20";
                    existing["accepted_count"] = "0";
                    existing["corrected_count"] = "0";
                    existing["rejected_count"] = "0";
                }
                else
                {
                    existing["target_kind"] = pair.Target.TargetKind;
                    existing["target_code"] = pair.Target.Code ?? "";
                    existing["target_name"] = pair.Target.Name ?? "";
                    existing["target_unit"] = pair.Target.Unit ?? "";
                    existing["quantity_name"] = quantityName;
                    existing["quantity_unit"] = quantityUnit;
                }

                existing["weight"] = (ReadFlatInt(existing, "weight", 0) + 20).ToString(CultureInfo.InvariantCulture);
                existing["corrected_count"] = (ReadFlatInt(existing, "corrected_count", 0) + 1).ToString(CultureInfo.InvariantCulture);
                existing["last_used_at"] = now;
            }

            TrimMappingRows(rows, 30);
            File.WriteAllLines(path, rows.Select(ToFlatJson).ToArray(), Encoding.UTF8);
            });
        }

        private static void BackupMappingTrainerFile(string path)
        {
            if (!File.Exists(path))
            {
                return;
            }

            string backupPath = path + ".pre-mapping-trainer.bak";
            if (!File.Exists(backupPath))
            {
                File.Copy(path, backupPath, false);
            }
        }

        private sealed class MappingTrainerTarget
        {
            public string TargetId;
            public string TargetKind;
            public string Code;
            public string Name;
            public string Unit;
            public string QuotaQuantity;

            public string TargetKey
            {
                get { return (String.IsNullOrWhiteSpace(TargetKind) ? GuessMappingTargetKind(Code) : TargetKind) + ":" + (Code ?? "").Trim().ToUpperInvariant(); }
            }
        }

        private sealed class MappingTrainerQuantity
        {
            public string QuantityId;
            public int ExcelRow;
            public string CellAddress;
            public string QuantityName;
            public string OriginalName;
            public string Unit;
            public string QuantityValue;
            public string RawText;
        }

        private sealed class MappingTrainerPairRow
        {
            public bool Checked;
            public MappingTrainerTarget Target;
            public MappingTrainerQuantity Quantity;
            public string Source;
            public int Confidence;
            public string Reason;
        }

        private sealed class MappingTrainerAiPair
        {
            public string TargetId;
            public string QuantityId;
            public string QuantityName;
            public int Confidence;
            public string Reason;
        }

        private sealed class MappingBoxTrainerDialog : Form
        {
            private readonly Form mainForm;
            private List<MappingTrainerTarget> targets = new List<MappingTrainerTarget>();
            private List<MappingTrainerQuantity> quantities = new List<MappingTrainerQuantity>();
            private readonly List<MappingTrainerPairRow> pairs = new List<MappingTrainerPairRow>();
            private readonly DataGridView grid;
            private readonly Label summaryLabel;
            private readonly Label status;
            private readonly CheckBox aiNameCheckBox;

            public MappingBoxTrainerDialog(Form mainForm)
            {
                this.mainForm = mainForm;

                Text = "\u6dfb\u52a0\u5bf9\u5e94\u6846\u5185\u5bb9";
                StartPosition = FormStartPosition.CenterParent;
                Size = new System.Drawing.Size(1080, 620);
                MinimumSize = new System.Drawing.Size(920, 460);
                MinimizeBox = false;

                Panel top = new Panel();
                top.Dock = DockStyle.Top;
                top.Height = 48;

                aiNameCheckBox = new CheckBox();
                aiNameCheckBox.Left = 10;
                aiNameCheckBox.Top = 14;
                aiNameCheckBox.Width = 130;
                aiNameCheckBox.Text = "AI\u8bc6\u522b\u5de5\u7a0b\u91cf";
                aiNameCheckBox.Checked = false;

                summaryLabel = new Label();
                summaryLabel.Left = 150;
                summaryLabel.Top = 12;
                summaryLabel.Width = 880;
                summaryLabel.Height = 28;
                summaryLabel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                summaryLabel.Text = "\u6253\u5f00\u672c\u7a97\u53e3\u540e\uff0c\u53ef\u5728\u5b9a\u989d\u8f93\u5165\u8868\u6846\u9009\u5b9a\u989d\uff0c\u518d\u5728Excel/WPS\u6846\u9009\u5de5\u7a0b\u91cf\uff0c\u7136\u540e\u70b9AI\u914d\u5bf9\u3002";
                top.Controls.Add(aiNameCheckBox);
                top.Controls.Add(summaryLabel);

                grid = new DataGridView();
                grid.Dock = DockStyle.Fill;
                grid.AllowUserToAddRows = false;
                grid.AllowUserToDeleteRows = false;
                grid.RowHeadersVisible = false;
                grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                grid.MultiSelect = true;
                grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                grid.CellContentClick += GridCellContentClick;
                AddGridColumns();
                FillGrid();

                Button aiPair = new Button();
                aiPair.Text = "AI\u914d\u5bf9";
                aiPair.Width = 82;
                aiPair.Click += delegate { RunAiPairing(); };

                Button applyExcel = new Button();
                applyExcel.Text = "\u8bfb\u53d6\u5f53\u524dExcel\u884c\u6276\u6b63";
                applyExcel.Width = 150;
                applyExcel.Click += delegate { ApplyCurrentExcelQuantity(); };

                Button selectAll = new Button();
                selectAll.Text = "\u5168\u9009";
                selectAll.Width = 70;
                selectAll.Click += delegate { SetChecked(true); };

                Button clearAll = new Button();
                clearAll.Text = "\u5168\u4e0d\u9009";
                clearAll.Width = 80;
                clearAll.Click += delegate { SetChecked(false); };

                Button ok = new Button();
                ok.Text = "\u786e\u8ba4\u5199\u5165";
                ok.Width = 95;
                ok.Click += delegate { ConfirmWrite(); };

                Button cancel = new Button();
                cancel.Text = "\u53d6\u6d88";
                cancel.Width = 75;
                cancel.Click += delegate { DialogResult = DialogResult.Cancel; Close(); };

                FlowLayoutPanel buttons = new FlowLayoutPanel();
                buttons.Dock = DockStyle.Bottom;
                buttons.Height = 44;
                buttons.FlowDirection = FlowDirection.RightToLeft;
                buttons.Padding = new Padding(8);
                buttons.Controls.Add(cancel);
                buttons.Controls.Add(ok);
                buttons.Controls.Add(clearAll);
                buttons.Controls.Add(selectAll);
                buttons.Controls.Add(applyExcel);
                buttons.Controls.Add(aiPair);

                status = new Label();
                status.Dock = DockStyle.Bottom;
                status.Height = 28;
                status.Padding = new Padding(8, 3, 8, 2);
                status.Text = "\u672c\u7a97\u53e3\u53ef\u4e0e\u5b9a\u989d\u8f93\u5165\u8868\u548cExcel\u540c\u65f6\u64cd\u4f5c\uff1bAI\u914d\u5bf9\u65f6\u4f1a\u8bfb\u53d6\u5f53\u524d\u4e24\u8fb9\u6846\u9009\u3002";

                Controls.Add(grid);
                Controls.Add(buttons);
                Controls.Add(status);
                Controls.Add(top);
            }

            private void AddGridColumns()
            {
                DataGridViewCheckBoxColumn checkedColumn = new DataGridViewCheckBoxColumn();
                checkedColumn.Name = "Checked";
                checkedColumn.HeaderText = "\u914d\u5bf9";
                checkedColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                checkedColumn.Width = 48;
                grid.Columns.Add(checkedColumn);

                DataGridViewButtonColumn correctColumn = new DataGridViewButtonColumn();
                correctColumn.Name = "Correct";
                correctColumn.HeaderText = "\u6276\u6b63";
                correctColumn.Text = "\u6276\u6b63";
                correctColumn.UseColumnTextForButtonValue = true;
                correctColumn.AutoSizeMode = DataGridViewAutoSizeColumnMode.None;
                correctColumn.Width = 54;
                grid.Columns.Add(correctColumn);

                grid.Columns.Add("QuantityName", "\u5de5\u7a0b\u91cf\u540d\u79f0");
                grid.Columns.Add("QuantityUnit", "\u5355\u4f4d");
                grid.Columns.Add("ExcelQuantity", "Excel\u5de5\u7a0b\u91cf");
                grid.Columns.Add("QuotaCode", "\u914d\u5bf9\u5b9a\u989d");
                grid.Columns.Add("QuotaName", "\u5b9a\u989d\u540d\u79f0");
                grid.Columns.Add("QuotaUnit", "\u5b9a\u989d\u5355\u4f4d");
                grid.Columns.Add("QuotaQuantity", "\u5b9a\u989d\u5de5\u7a0b\u91cf");
                grid.Columns.Add("Source", "\u6765\u6e90");

                grid.Columns["QuantityName"].FillWeight = 170;
                grid.Columns["QuantityUnit"].FillWeight = 50;
                grid.Columns["ExcelQuantity"].FillWeight = 75;
                grid.Columns["QuotaCode"].FillWeight = 75;
                grid.Columns["QuotaName"].FillWeight = 210;
                grid.Columns["QuotaUnit"].FillWeight = 60;
                grid.Columns["QuotaQuantity"].FillWeight = 85;
                grid.Columns["Source"].FillWeight = 75;

                foreach (DataGridViewColumn column in grid.Columns)
                {
                    if (column.Name != "Checked" && column.Name != "QuantityName")
                    {
                        column.ReadOnly = true;
                    }
                }
            }

            private void FillGrid()
            {
                grid.Rows.Clear();
                foreach (MappingTrainerPairRow pair in pairs)
                {
                    int rowIndex = grid.Rows.Add(
                        pair.Checked,
                        "\u6276\u6b63",
                        pair.Quantity == null ? "" : pair.Quantity.QuantityName,
                        pair.Quantity == null ? "" : pair.Quantity.Unit,
                        pair.Quantity == null ? "" : pair.Quantity.QuantityValue,
                        pair.Target == null ? "" : pair.Target.Code,
                        pair.Target == null ? "" : pair.Target.Name,
                        pair.Target == null ? "" : pair.Target.Unit,
                        pair.Target == null ? "" : pair.Target.QuotaQuantity,
                        BuildSourceText(pair));
                    grid.Rows[rowIndex].Tag = pair;
                    if (pair.Quantity != null)
                    {
                        grid.Rows[rowIndex].Cells["QuantityName"].ToolTipText = pair.Quantity.RawText ?? "";
                    }
                }
            }

            private static string BuildSourceText(MappingTrainerPairRow pair)
            {
                if (pair == null)
                {
                    return "";
                }
                return pair.Source ?? "";
            }

            private void GridCellContentClick(object sender, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0)
                {
                    return;
                }

                if (grid.Columns[e.ColumnIndex].Name == "Correct")
                {
                    CorrectPairTarget(e.RowIndex);
                }
            }

            private bool LoadCurrentSelectionsForPairing()
            {
                DataGridView quotaGrid = GetField<DataGridView>(mainForm, "dataGridViewDE");
                List<MappingTrainerTarget> currentTargets = BuildMappingTrainerTargets(quotaGrid);
                if (currentTargets.Count == 0)
                {
                    status.Text = "\u8bf7\u5148\u5728\u5b9a\u989d\u8f93\u5165\u8868\u4e2d\u6846\u9009\u8981\u6279\u91cf\u914d\u5bf9\u7684\u5b9a\u989d\u884c\u3002";
                    return false;
                }

                AiExcelSelectionContext selection;
                string error;
                if (!TryReadActiveExcelSelectionContext(out selection, out error))
                {
                    status.Text = error;
                    return false;
                }

                List<MappingTrainerQuantity> currentQuantities = BuildMappingTrainerQuantities(selection);
                if (currentQuantities.Count == 0)
                {
                    status.Text = "\u5f53\u524d Excel \u9009\u533a\u6ca1\u6709\u8bc6\u522b\u5230\u53ef\u914d\u5bf9\u7684\u5de5\u7a0b\u91cf\u884c\u3002";
                    return false;
                }

                targets = currentTargets;
                quantities = currentQuantities;
                pairs.Clear();
                foreach (MappingTrainerTarget target in targets)
                {
                    pairs.Add(new MappingTrainerPairRow { Target = target, Quantity = null, Checked = false, Source = "\u672a\u914d\u5bf9" });
                }

                summaryLabel.Text = "\u5df2\u8bfb\u53d6\u5b9a\u989d " + targets.Count.ToString(CultureInfo.InvariantCulture) + " \u884c\uff0cExcel\u5de5\u7a0b\u91cf " + quantities.Count.ToString(CultureInfo.InvariantCulture) + " \u884c\u3002";
                FillGrid();
                return true;
            }

            private void RunAiPairing()
            {
                grid.EndEdit();
                if (!LoadCurrentSelectionsForPairing())
                {
                    return;
                }

                foreach (MappingTrainerPairRow pair in pairs)
                {
                    pair.Quantity = null;
                    pair.Checked = false;
                    pair.Source = "\u672a\u914d\u5bf9";
                    pair.Confidence = 0;
                    pair.Reason = "";
                }

                int localCount = ApplyLocalQuantityPairs();
                DeepSeekExcelMatchSettings settings = LoadDeepSeekExcelMatchSettings();
                int aiCount = 0;
                if (settings.IsAvailable)
                {
                    try
                    {
                        status.Text = "DeepSeek\u6b63\u5728\u6279\u91cf\u914d\u5bf9\u5b9a\u989d\u548cExcel\u5de5\u7a0b\u91cf...";
                        status.Refresh();
                        aiCount = ApplyDeepSeekPairs(settings);
                    }
                    catch (Exception ex)
                    {
                        Log("Mapping trainer DeepSeek pairing failed: " + ex.Message);
                        status.Text = "DeepSeek\u914d\u5bf9\u5931\u8d25\uff1a" + ex.Message + "\uff1b\u5df2\u4fdd\u7559\u672c\u5730\u6570\u91cf\u914d\u5bf9\u7ed3\u679c\u3002";
                    }
                }
                else
                {
                    status.Text = "DeepSeek\u672a\u542f\u7528\u6216\u672a\u914d\u7f6eAPI Key\uff0c\u5df2\u53ea\u6267\u884c\u672c\u5730\u6570\u91cf\u914d\u5bf9\u3002";
                }

                FillGrid();
                if (settings.IsAvailable)
                {
                    status.Text = "\u5df2\u5b8c\u6210\u914d\u5bf9\uff1a\u672c\u5730\u6570\u91cf " + localCount.ToString(CultureInfo.InvariantCulture) + " \u884c\uff0cAI " + aiCount.ToString(CultureInfo.InvariantCulture) + " \u884c\u3002\u8bf7\u68c0\u67e5\u540e\u786e\u8ba4\u5199\u5165\u3002";
                }
            }

            private int ApplyLocalQuantityPairs()
            {
                int count = 0;
                HashSet<string> usedQuantities = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (MappingTrainerPairRow pair in pairs)
                {
                    decimal quotaQuantity;
                    string quotaError;
                    if (pair.Target == null || !TryEvaluateDecimal(pair.Target.QuotaQuantity, out quotaQuantity, out quotaError))
                    {
                        continue;
                    }

                    MappingTrainerQuantity best = null;
                    decimal bestScore = Decimal.MaxValue;
                    foreach (MappingTrainerQuantity quantity in quantities)
                    {
                        if (quantity == null || usedQuantities.Contains(quantity.QuantityId))
                        {
                            continue;
                        }

                        decimal excelQuantity;
                        string excelError;
                        if (!TryEvaluateDecimal(quantity.QuantityValue, out excelQuantity, out excelError))
                        {
                            continue;
                        }

                        decimal score = RelativeDifference(quotaQuantity, excelQuantity);
                        if (score <= 0.03m && score < bestScore)
                        {
                            bestScore = score;
                            best = quantity;
                        }
                    }

                    if (best == null)
                    {
                        continue;
                    }

                    pair.Quantity = best;
                    pair.Checked = !String.IsNullOrWhiteSpace(best.QuantityName);
                    pair.Source = "\u672c\u5730\u6570\u91cf";
                    pair.Confidence = 100;
                    pair.Reason = "\u5b9a\u989d\u5de5\u7a0b\u91cf\u548cExcel\u5de5\u7a0b\u91cf\u63a5\u8fd1";
                    usedQuantities.Add(best.QuantityId);
                    count++;
                }

                return count;
            }

            private int ApplyDeepSeekPairs(DeepSeekExcelMatchSettings settings)
            {
                List<MappingTrainerTarget> unmatchedTargets = pairs
                    .Where(p => p.Quantity == null && p.Target != null)
                    .Select(p => p.Target)
                    .ToList();
                HashSet<string> usedQuantities = new HashSet<string>(
                    pairs.Where(p => p.Quantity != null).Select(p => p.Quantity.QuantityId),
                    StringComparer.OrdinalIgnoreCase);
                List<MappingTrainerQuantity> availableQuantities = quantities
                    .Where(q => q != null && !usedQuantities.Contains(q.QuantityId))
                    .ToList();

                if (unmatchedTargets.Count == 0 || availableQuantities.Count == 0)
                {
                    return 0;
                }

                List<MappingTrainerAiPair> aiPairs = RequestMappingTrainerAiPairs(settings, unmatchedTargets, availableQuantities, aiNameCheckBox.Checked);
                Dictionary<string, MappingTrainerTarget> targetById = unmatchedTargets.ToDictionary(t => t.TargetId, StringComparer.OrdinalIgnoreCase);
                Dictionary<string, MappingTrainerQuantity> quantityById = availableQuantities.ToDictionary(q => q.QuantityId, StringComparer.OrdinalIgnoreCase);
                HashSet<string> usedTargets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                int applied = 0;

                foreach (MappingTrainerAiPair aiPair in aiPairs.OrderByDescending(p => p.Confidence))
                {
                    if (aiPair == null || aiPair.Confidence < 55)
                    {
                        continue;
                    }

                    MappingTrainerTarget target;
                    MappingTrainerQuantity quantity;
                    if (!targetById.TryGetValue(aiPair.TargetId ?? "", out target) ||
                        !quantityById.TryGetValue(aiPair.QuantityId ?? "", out quantity) ||
                        usedTargets.Contains(target.TargetId) ||
                        usedQuantities.Contains(quantity.QuantityId))
                    {
                        continue;
                    }

                    if (aiNameCheckBox.Checked && !String.IsNullOrWhiteSpace(aiPair.QuantityName))
                    {
                        quantity.QuantityName = CleanTrainerText(aiPair.QuantityName);
                    }
                    else if (String.IsNullOrWhiteSpace(quantity.QuantityName) && !String.IsNullOrWhiteSpace(aiPair.QuantityName))
                    {
                        quantity.QuantityName = CleanTrainerText(aiPair.QuantityName);
                    }

                    MappingTrainerPairRow pair = pairs.FirstOrDefault(p => Object.ReferenceEquals(p.Target, target));
                    if (pair == null)
                    {
                        continue;
                    }

                    pair.Quantity = quantity;
                    pair.Checked = aiPair.Confidence >= 65 && !String.IsNullOrWhiteSpace(quantity.QuantityName);
                    pair.Source = "AI\u914d\u5bf9";
                    pair.Confidence = aiPair.Confidence;
                    pair.Reason = aiPair.Reason ?? "";
                    usedTargets.Add(target.TargetId);
                    usedQuantities.Add(quantity.QuantityId);
                    applied++;
                }

                return applied;
            }

            private static List<MappingTrainerAiPair> RequestMappingTrainerAiPairs(DeepSeekExcelMatchSettings settings, List<MappingTrainerTarget> targets, List<MappingTrainerQuantity> quantities, bool normalizeNames)
            {
                JavaScriptSerializer serializer = new JavaScriptSerializer();
                serializer.MaxJsonLength = 1024 * 1024 * 4;
                string requestJson = BuildMappingTrainerAiRequestJson(serializer, settings, targets, quantities, normalizeNames);
                string responseJson = SendDeepSeekExcelMatchRequest(settings, requestJson);
                return ParseMappingTrainerAiResponse(serializer, responseJson);
            }

            private static string BuildMappingTrainerAiRequestJson(JavaScriptSerializer serializer, DeepSeekExcelMatchSettings settings, List<MappingTrainerTarget> targets, List<MappingTrainerQuantity> quantities, bool normalizeNames)
            {
                List<object> targetObjects = new List<object>();
                foreach (MappingTrainerTarget target in targets ?? new List<MappingTrainerTarget>())
                {
                    targetObjects.Add(new Dictionary<string, object>
                    {
                        { "target_id", target.TargetId ?? "" },
                        { "code", target.Code ?? "" },
                        { "name", target.Name ?? "" },
                        { "unit", target.Unit ?? "" },
                        { "quota_quantity", target.QuotaQuantity ?? "" }
                    });
                }

                List<object> quantityObjects = new List<object>();
                foreach (MappingTrainerQuantity quantity in quantities ?? new List<MappingTrainerQuantity>())
                {
                    quantityObjects.Add(new Dictionary<string, object>
                    {
                        { "quantity_id", quantity.QuantityId ?? "" },
                        { "local_quantity_name", quantity.QuantityName ?? "" },
                        { "unit", quantity.Unit ?? "" },
                        { "excel_quantity", quantity.QuantityValue ?? "" },
                        { "cell_address", quantity.CellAddress ?? "" },
                        { "raw_row_text", TruncateForPrompt(quantity.RawText, 280) }
                    });
                }

                Dictionary<string, object> body = new Dictionary<string, object>();
                body["task"] = "Match selected railway quota rows to Excel engineering quantity rows.";
                body["normalize_quantity_names"] = normalizeNames;
                body["rules"] = new string[]
                {
                    "Return strict JSON: {\"results\":[{\"target_id\":\"t0\",\"quantity_id\":\"q0\",\"quantity_name\":\"short Chinese quantity name\",\"confidence\":80,\"reason\":\"short reason\"}]}",
                    "target_id must come from targets and quantity_id must come from quantities.",
                    "Use quota name, quota unit, quota quantity, Excel quantity, and row text together.",
                    "Prefer rows whose quantities are numerically close, but do not ignore strong text conflicts.",
                    "Do not match one quantity row to multiple targets.",
                    "If normalize_quantity_names is true, summarize long quantity names; otherwise keep the local name unless it is empty or noisy.",
                    "If unsure, omit the result or set confidence below 55."
                };
                body["targets"] = targetObjects;
                body["quantities"] = quantityObjects;

                Dictionary<string, object> payload = new Dictionary<string, object>();
                payload["model"] = settings.Model;
                payload["stream"] = false;
                payload["temperature"] = 0.1;
                payload["max_tokens"] = 2200;
                payload["response_format"] = new Dictionary<string, object> { { "type", "json_object" } };
                payload["messages"] = new object[]
                {
                    new Dictionary<string, object>
                    {
                        { "role", "system" },
                        { "content", "You are a conservative railway construction quota and Excel quantity matching assistant. Return JSON only." }
                    },
                    new Dictionary<string, object>
                    {
                        { "role", "user" },
                        { "content", serializer.Serialize(body) }
                    }
                };

                return serializer.Serialize(payload);
            }

            private static List<MappingTrainerAiPair> ParseMappingTrainerAiResponse(JavaScriptSerializer serializer, string responseJson)
            {
                Dictionary<string, object> root = serializer.DeserializeObject(responseJson) as Dictionary<string, object>;
                List<object> choices = GetJsonList(root, "choices");
                Dictionary<string, object> firstChoice = choices == null || choices.Count == 0 ? null : choices[0] as Dictionary<string, object>;
                Dictionary<string, object> message = firstChoice == null ? null : ReadJsonObject(firstChoice, "message");
                string content = message == null ? "" : ReadJsonString(message, "content", "");
                if (String.IsNullOrWhiteSpace(content))
                {
                    return new List<MappingTrainerAiPair>();
                }

                Dictionary<string, object> resultRoot = serializer.DeserializeObject(content) as Dictionary<string, object>;
                List<object> results = GetJsonList(resultRoot, "results");
                if (results == null)
                {
                    return new List<MappingTrainerAiPair>();
                }

                List<MappingTrainerAiPair> parsed = new List<MappingTrainerAiPair>();
                foreach (object item in results)
                {
                    Dictionary<string, object> row = item as Dictionary<string, object>;
                    if (row == null)
                    {
                        continue;
                    }

                    parsed.Add(new MappingTrainerAiPair
                    {
                        TargetId = ReadJsonString(row, "target_id", ""),
                        QuantityId = ReadJsonString(row, "quantity_id", ""),
                        QuantityName = ReadJsonString(row, "quantity_name", ""),
                        Confidence = ReadJsonInt(row, "confidence", 0),
                        Reason = ReadJsonString(row, "reason", "")
                    });
                }

                return parsed;
            }

            private void CorrectPairTarget(int rowIndex)
            {
                if (rowIndex < 0 || rowIndex >= grid.Rows.Count)
                {
                    return;
                }

                MappingTrainerPairRow pair = grid.Rows[rowIndex].Tag as MappingTrainerPairRow;
                MappingTrainerTarget selected = BuildCurrentSelectedTrainerTarget();
                if (pair == null || selected == null)
                {
                    status.Text = "\u8bf7\u5148\u5728\u5b9a\u989d\u8f93\u5165\u8868\u4e2d\u70b9\u9009\u8981\u6276\u6b63\u5230\u7684\u5b9a\u989d\u884c\u3002";
                    return;
                }

                pair.Target = selected;
                pair.Checked = pair.Quantity != null && !String.IsNullOrWhiteSpace(pair.Quantity.QuantityName);
                pair.Source = "\u4eba\u5de5\u6276\u6b63";
                pair.Confidence = 100;
                pair.Reason = "\u4f7f\u7528\u5b9a\u989d\u8f93\u5165\u8868\u5f53\u524d\u9009\u4e2d\u5b9a\u989d";
                FillGrid();
                if (rowIndex < grid.Rows.Count)
                {
                    grid.Rows[rowIndex].Selected = true;
                }
                status.Text = "\u5df2\u6276\u6b63\u914d\u5bf9\u5b9a\u989d\uff1a" + selected.Code + " " + selected.Name;
            }

            private MappingTrainerTarget BuildCurrentSelectedTrainerTarget()
            {
                DataGridView quotaGrid = GetField<DataGridView>(mainForm, "dataGridViewDE");
                DataGridViewRow row = GetCurrentQuotaRow(quotaGrid);
                if (row == null)
                {
                    return null;
                }

                List<MappingTrainerTarget> selected = BuildMappingTrainerTargetsFromRows(new List<DataGridViewRow> { row });
                return selected.Count == 0 ? null : selected[0];
            }

            private static List<MappingTrainerTarget> BuildMappingTrainerTargetsFromRows(List<DataGridViewRow> rows)
            {
                List<MappingTrainerTarget> result = new List<MappingTrainerTarget>();
                int index = 0;
                foreach (DataGridViewRow row in rows ?? new List<DataGridViewRow>())
                {
                    string code = GetRowValue(row, "\u5b9a\u989d\u7f16\u53f7", "\u5b9a\u989d\u7f16\u53f7DE", "\u7f16\u53f7");
                    if (String.IsNullOrWhiteSpace(code))
                    {
                        continue;
                    }

                    result.Add(new MappingTrainerTarget
                    {
                        TargetId = "c" + index.ToString(CultureInfo.InvariantCulture),
                        TargetKind = GuessMappingTargetKind(code),
                        Code = code.Trim(),
                        Name = GetRowValue(row, "\u5de5\u7a0b\u6216\u8d39\u7528\u9879\u76ee\u540d\u79f0", "\u540d\u79f0", "\u9879\u76ee\u540d\u79f0"),
                        Unit = GetRowValue(row, "\u5355\u4f4d", "\u5b9a\u989d\u5355\u4f4d"),
                        QuotaQuantity = GetRowValue(row, "\u5de5\u7a0b\u6570\u91cf\u8f93\u5165", "\u5de5\u7a0b\u6570\u91cf")
                    });
                    index++;
                }

                return result;
            }

            private void ApplyCurrentExcelQuantity()
            {
                grid.EndEdit();
                DataGridViewRow gridRow = grid.CurrentRow;
                MappingTrainerPairRow pair = gridRow == null ? null : gridRow.Tag as MappingTrainerPairRow;
                if (pair == null)
                {
                    status.Text = "\u8bf7\u5148\u9009\u4e2d\u9700\u8981\u6276\u6b63\u7684\u9884\u89c8\u884c\u3002";
                    return;
                }

                ExcelCellAddress cell;
                string error;
                if (!TryGetActiveExcelCell(out cell, out error))
                {
                    status.Text = "\u8bf7\u5148\u5728 WPS/Excel \u91cc\u70b9\u9009\u5bf9\u5e94\u7684\u5de5\u7a0b\u6570\u91cf\u5355\u5143\u683c\u3002";
                    return;
                }

                string name = BuildQuantityNameNearActiveExcelCell(cell);
                MappingTrainerQuantity quantity = new MappingTrainerQuantity();
                quantity.QuantityId = "m" + DateTime.Now.Ticks.ToString(CultureInfo.InvariantCulture);
                quantity.CellAddress = NormalizeCellAddress(cell.CellAddress);
                quantity.QuantityName = CleanTrainerText(name);
                quantity.OriginalName = quantity.QuantityName;
                quantity.Unit = "";
                quantity.QuantityValue = cell.DisplayValue ?? "";
                quantity.RawText = (name + " " + quantity.QuantityValue).Trim();

                pair.Quantity = quantity;
                pair.Checked = !String.IsNullOrWhiteSpace(quantity.QuantityName);
                pair.Source = "\u4eba\u5de5\u6276\u6b63";
                pair.Confidence = 100;
                pair.Reason = "\u4f7f\u7528\u5f53\u524dExcel\u5355\u5143\u683c\u884c";
                FillGrid();
                status.Text = "\u5df2\u8bfb\u53d6\u5f53\u524dExcel\u884c\uff0c\u53ef\u7ee7\u7eed\u7f16\u8f91\u5de5\u7a0b\u91cf\u540d\u79f0\u540e\u5199\u5165\u3002";
            }

            private void SetChecked(bool value)
            {
                grid.EndEdit();
                foreach (DataGridViewRow row in grid.Rows)
                {
                    row.Cells["Checked"].Value = value;
                    MappingTrainerPairRow pair = row.Tag as MappingTrainerPairRow;
                    if (pair != null)
                    {
                        pair.Checked = value;
                    }
                }
            }

            private void ConfirmWrite()
            {
                grid.EndEdit();
                List<MappingTrainerPairRow> accepted = GetAcceptedRows();
                if (accepted.Count == 0)
                {
                    status.Text = "\u6ca1\u6709\u52fe\u9009\u4efb\u4f55\u6709\u6548\u914d\u5bf9\u884c\u3002";
                    return;
                }

                try
                {
                    SaveMappingTrainerPairs(accepted);
                    status.Text = "\u5df2\u5199\u5165 " + accepted.Count.ToString(CultureInfo.InvariantCulture) + " \u6761\u5b9a\u989d-\u5de5\u7a0b\u91cf\u5bf9\u5e94\u6837\u672c\u3002";
                    DialogResult = DialogResult.OK;
                    Close();
                }
                catch (Exception ex)
                {
                    Log("Mapping trainer save failed: " + ex);
                    status.Text = "\u5199\u5165\u5931\u8d25\uff1a" + ex.Message;
                }
            }

            private List<MappingTrainerPairRow> GetAcceptedRows()
            {
                List<MappingTrainerPairRow> accepted = new List<MappingTrainerPairRow>();
                foreach (DataGridViewRow row in grid.Rows)
                {
                    bool isChecked = row.Cells["Checked"].Value is bool && (bool)row.Cells["Checked"].Value;
                    MappingTrainerPairRow pair = row.Tag as MappingTrainerPairRow;
                    if (!isChecked || pair == null || pair.Target == null || pair.Quantity == null)
                    {
                        continue;
                    }

                    pair.Checked = true;
                    pair.Quantity.QuantityName = CleanTrainerText(Convert.ToString(row.Cells["QuantityName"].Value, CultureInfo.CurrentCulture));
                    if (!String.IsNullOrWhiteSpace(pair.Quantity.QuantityName) && !String.IsNullOrWhiteSpace(pair.Target.Code))
                    {
                        accepted.Add(pair);
                    }
                }

                return accepted;
            }
        }
    }
}
