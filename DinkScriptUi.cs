using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Windows.Forms;

namespace DinkCel
{
    internal sealed class ScriptCellWrite
    {
        public int Row, Column;
        public object Value;
    }

    internal sealed class AiConfiguration
    {
        public string Endpoint = "https://api.openai.com/v1/chat/completions";
        public string Model = "";
        public string ProtectedKey = "";
    }

    internal static class AiClient
    {
        public static string Send(AiConfiguration config, string prompt)
        {
            Uri endpoint;
            if (!Uri.TryCreate(config.Endpoint, UriKind.Absolute, out endpoint) ||
                (endpoint.Scheme != "https" && !(endpoint.Scheme == "http" && endpoint.IsLoopback)))
                throw new InvalidOperationException("AI endpoint must use HTTPS or HTTP localhost.");
            var body = new JavaScriptSerializer().Serialize(new { model = config.Model,
                messages = new[] { new { role = "user", content = prompt } } });
            var request = (HttpWebRequest)WebRequest.Create(endpoint);
            request.Method = "POST"; request.ContentType = "application/json"; request.Timeout = 20000;
            request.AllowAutoRedirect = false;
            if (!String.IsNullOrEmpty(config.ProtectedKey))
                request.Headers[HttpRequestHeader.Authorization] = "Bearer " + Encoding.UTF8.GetString(
                    ProtectedData.Unprotect(Convert.FromBase64String(config.ProtectedKey), null,
                        DataProtectionScope.CurrentUser));
            byte[] payload = Encoding.UTF8.GetBytes(body);
            using (Stream stream = request.GetRequestStream()) stream.Write(payload, 0, payload.Length);
            using (var response = request.GetResponse())
            using (var reader = new StreamReader(response.GetResponseStream()))
            {
                var result = new JavaScriptSerializer().DeserializeObject(reader.ReadToEnd()) as Dictionary<string, object>;
                var choices = result["choices"] as object[];
                var choice = choices[0] as Dictionary<string, object>;
                var message = choice["message"] as Dictionary<string, object>;
                return Convert.ToString(message["content"]);
            }
        }
    }

    internal sealed partial class SpreadsheetForm
    {
        private bool scriptEnabled;
        private bool aiEnabled;
        private AiConfiguration aiConfiguration;
        private readonly Dictionary<string, string> aiFormulaCache = new Dictionary<string, string>();
        private readonly HashSet<string> aiFormulaPending = new HashSet<string>();
        private int aiFormulaRequests;
        private readonly HashSet<string> scriptFunctions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private static readonly Regex ScriptRange = new Regex(@"^([A-Z]+)([1-9][0-9]*)(?::([A-Z]+)([1-9][0-9]*))?$", RegexOptions.IgnoreCase);

        private void ConfigureCustomFunctions(FormulaEngine engine)
        {
            scriptFunctions.Clear();
            if (scriptEnabled)
                foreach (string name in DinkScript.CustomFunctions(scriptCode)) scriptFunctions.Add(name);
            Dictionary<string, Dictionary<int, object>> snapshot = scriptFunctions.Count > 0 ?
                ScriptSnapshot() : null;
            engine.HasCustomFunction = name => scriptFunctions.Contains(name);
            engine.CustomFunction = (name, args) =>
                new DinkScript().Invoke(scriptCode, name, args, sheets[activeSheetIndex].Name,
                    (sheet, range) => ScriptRead(snapshot, sheet, range), null, null,
                    prompt => GenerateAiFormula(prompt));
        }

        private string GenerateAiFormula(string prompt)
        {
            if (!aiEnabled) return "#AI_DISABLED!";
            AiConfiguration config = LoadAiSettings();
            string key = config.Endpoint + "\n" + config.Model + "\n" + scriptCode + "\n" + prompt;
            string cached;
            if (aiFormulaCache.TryGetValue(key, out cached)) return cached;
            if (aiFormulaPending.Contains(key) || aiFormulaPending.Count > 0) return "#BUSY!";
            if (aiFormulaRequests >= 10) return "#AI_LIMIT!";
            aiFormulaPending.Add(key);
            aiFormulaRequests++;
            Task.Run(() =>
            {
                string result;
                try { result = GenerateAi(prompt); }
                catch (Exception error) { result = "#AI_ERROR! " + error.Message; }
                if (!IsDisposed && IsHandleCreated)
                    BeginInvoke((Action)(() =>
                    {
                        aiFormulaPending.Remove(key);
                        aiFormulaCache[key] = result;
                        if (scriptEnabled && aiEnabled) Recalculate();
                    }));
            });
            return "#BUSY!";
        }

        private static Rectangle ParseScriptRange(string address)
        {
            Match match = ScriptRange.Match((address ?? "").Trim());
            if (!match.Success) throw new ArgumentException("Invalid range: " + address);
            Func<string, int> column = letters =>
            {
                int value = 0;
                foreach (char c in letters.ToUpperInvariant())
                    value = checked(value * 26 + c - 'A' + 1);
                return value - 1;
            };
            int left = column(match.Groups[1].Value), top = Int32.Parse(match.Groups[2].Value) - 1;
            int right = match.Groups[3].Success ? column(match.Groups[3].Value) : left;
            int bottom = match.Groups[4].Success ? Int32.Parse(match.Groups[4].Value) - 1 : top;
            if (left < 0 || right < left || right >= ColumnCount || top < 0 || bottom < top ||
                bottom >= MaxRowCount || (long)(right - left + 1) * (bottom - top + 1) > 100000)
                throw new ArgumentException("Range is outside the sheet or exceeds 100,000 cells.");
            return Rectangle.FromLTRB(left, top, right + 1, bottom + 1);
        }

        private Dictionary<string, Dictionary<int, object>> ScriptSnapshot()
        {
            var result = new Dictionary<string, Dictionary<int, object>>(StringComparer.OrdinalIgnoreCase);
            foreach (SheetState sheet in sheets)
                result[sheet.Name] = sheet.Cells.ToDictionary(x => x.Key, x => x.Value.Value);
            Dictionary<int, object> active = result[sheets[activeSheetIndex].Name];
            if (scanAllCells)
            {
                active.Clear();
                for (int row = 0; row < RowCount; row++)
                    for (int col = 0; col < ColumnCount; col++)
                    {
                        object value = grid[col, row].Value;
                        if (value != null) active[row * ColumnCount + col] = value;
                    }
            }
            else
                foreach (int key in changedCells)
                {
                    if (key < 0 || key / ColumnCount >= RowCount) continue;
                    object value = grid[key % ColumnCount, key / ColumnCount].Value;
                    if (value == null) active.Remove(key);
                    else active[key] = value;
                }
            return result;
        }

        private static object[][] ScriptRead(Dictionary<string, Dictionary<int, object>> snapshot,
            string sheet, string address)
        {
            Dictionary<int, object> cells;
            if (!snapshot.TryGetValue(sheet, out cells)) throw new ArgumentException("Unknown sheet: " + sheet);
            Rectangle range = ParseScriptRange(address);
            var values = new object[range.Height][];
            for (int row = 0; row < range.Height; row++)
            {
                values[row] = new object[range.Width];
                for (int col = 0; col < range.Width; col++)
                {
                    object raw;
                    cells.TryGetValue((range.Top + row) * ColumnCount + range.Left + col, out raw);
                    double number;
                    string text = Convert.ToString(raw, CultureInfo.InvariantCulture);
                    values[row][col] = Double.TryParse(text, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out number) ? (object)number : raw;
                }
            }
            return values;
        }

        private void OpenScriptEditor()
        {
            using (var form = new Form { Text = "DinkCel Script Editor", Width = 900, Height = 660,
                StartPosition = FormStartPosition.CenterParent, Font = new Font("Segoe UI", 10),
                BackColor = theme.Chrome, ForeColor = theme.Text })
            {
                var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 42, Padding = new Padding(8, 5, 0, 0) };
                var functions = new ComboBox { Width = 180, DropDownStyle = ComboBoxStyle.DropDown,
                    Text = "main" };
                var save = new Button { Text = "Save script", Width = 110 };
                var run = new Button { Text = "Run", Width = 75 };
                var stop = new Button { Text = "Stop", Width = 75, Enabled = false };
                top.Controls.AddRange(new Control[] { functions, save, run, stop });
                var editor = new TextBox { Multiline = true, AcceptsTab = true, ScrollBars = ScrollBars.Both,
                    WordWrap = false, Dock = DockStyle.Fill, Font = new Font("Consolas", 11),
                    Text = String.IsNullOrWhiteSpace(scriptCode) ?
                        "function main() {\r\n  const sheet = DinkCel.getActiveWorkbook().getActiveSheet();\r\n  sheet.getRange('A1').setValue('Hello DinkCel');\r\n  Logger.log('Done');\r\n}\r\n\r\n/** @customfunction */\r\nfunction DOUBLE(value) { return Number(value) * 2; }\r\n" : scriptCode };
                var output = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical,
                    Dock = DockStyle.Bottom, Height = 130, Font = new Font("Consolas", 9),
                    BackColor = theme.Sheet, ForeColor = theme.Text };
                form.Controls.Add(editor); form.Controls.Add(output); form.Controls.Add(top);
                CancellationTokenSource source = null;
                Action saveCode = () =>
                {
                    if (scriptCode != editor.Text)
                    {
                        scriptCode = editor.Text;
                        aiFormulaCache.Clear();
                        MarkDirty();
                    }
                    scriptEnabled = true;
                    Recalculate();
                    output.AppendText("Saved to workbook. Use .dinkcel to keep the script.\r\n");
                };
                save.Click += (s, e) => saveCode();
                form.FormClosing += (s, e) =>
                {
                    if (source != null)
                    {
                        source.Cancel();
                        e.Cancel = true;
                        return;
                    }
                    if (editor.Text == scriptCode) return;
                    DialogResult choice = MessageBox.Show(form, "Save script changes?", "DinkCel Script",
                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
                    if (choice == DialogResult.Cancel) e.Cancel = true;
                    else if (choice == DialogResult.Yes) saveCode();
                };
                editor.TextChanged += (s, e) =>
                {
                    string selected = functions.Text;
                    functions.Items.Clear();
                    foreach (Match match in Regex.Matches(editor.Text, @"\bfunction\s+([A-Za-z_$][\w$]*)\s*\("))
                        functions.Items.Add(match.Groups[1].Value);
                    functions.Text = selected;
                };
                stop.Click += (s, e) => { if (source != null) source.Cancel(); };
                run.Click += async (s, e) =>
                {
                    saveCode();
                    grid.EndEdit();
                    Dictionary<string, Dictionary<int, object>> snapshot = ScriptSnapshot();
                    var writes = new List<ScriptCellWrite>();
                    string activeName = sheets[activeSheetIndex].Name;
                    string code = scriptCode, function = functions.Text.Trim();
                    int activeIndex = activeSheetIndex;
                    source = new CancellationTokenSource();
                    run.Enabled = false; stop.Enabled = true;
                    output.AppendText("Running " + function + "...\r\n");
                    try
                    {
                        object answer = await Task.Run(() => new DinkScript().Invoke(code, function,
                            new object[0], activeName,
                            (sheet, range) => ScriptRead(snapshot, sheet, range),
                            (sheet, range, values) =>
                            {
                                if (!String.Equals(sheet, activeName, StringComparison.OrdinalIgnoreCase))
                                    throw new InvalidOperationException("Writing another sheet is not yet supported.");
                                Rectangle area = ParseScriptRange(range);
                                if (values == null || values.Length != area.Height ||
                                    values.Any(x => x == null || x.Length != area.Width))
                                    throw new ArgumentException("setValues size must match the range.");
                                if (writes.Count + area.Width * area.Height > 100000)
                                    throw new InvalidOperationException("Script changed more than 100,000 cells.");
                                for (int row = 0; row < area.Height; row++)
                                    for (int col = 0; col < area.Width; col++)
                                    {
                                        writes.Add(new ScriptCellWrite { Row = area.Top + row,
                                            Column = area.Left + col, Value = values[row][col] });
                                        snapshot[activeName][(area.Top + row) * ColumnCount + area.Left + col] = values[row][col];
                                    }
                            }, message => BeginInvoke((Action)(() => output.AppendText(message + "\r\n"))),
                            prompt => GenerateAi(prompt), source.Token));
                        if (activeSheetIndex != activeIndex)
                            throw new InvalidOperationException("Active sheet changed while script was running.");
                        ApplyScriptWrites(writes);
                        output.AppendText("Done: " + writes.Count + " cells. " + Convert.ToString(answer) + "\r\n");
                    }
                    catch (Exception error) { output.AppendText("Error: " + error.Message + "\r\n"); }
                    finally { run.Enabled = true; stop.Enabled = false; source.Dispose(); source = null; }
                };
                form.ShowDialog(this);
            }
        }

        private void ApplyScriptWrites(List<ScriptCellWrite> writes)
        {
            if (writes.Count == 0) return;
            if (sheets[activeSheetIndex].Protected || grid.ReadOnly)
                throw new InvalidOperationException("The active sheet is protected.");
            if (writes.Count > 100000 || writes.Any(x => x.Row < 0 || x.Row >= MaxRowCount ||
                x.Column < 0 || x.Column >= ColumnCount))
                throw new InvalidOperationException("Script write is outside allowed cell limits.");
            EnsureRowCapacity(writes.Max(x => x.Row) + 1);
            SheetState before = CaptureSheet();
            loading = true;
            try
            {
                foreach (ScriptCellWrite write in writes)
                {
                    int key = write.Row * ColumnCount + write.Column;
                    grid[write.Column, write.Row].Value = write.Value == null ? null :
                        Convert.ToString(write.Value, CultureInfo.InvariantCulture);
                    changedCells.Add(key);
                    if (write.Value == null) { gridOccupied.Remove(key); formulaKeys.Remove(key); }
                    else
                    {
                        gridOccupied.Add(key);
                        if (Convert.ToString(write.Value).StartsWith("=")) formulaKeys.Add(key);
                        else formulaKeys.Remove(key);
                    }
                }
            }
            catch
            {
                loading = false;
                RestoreSheet(before);
                throw;
            }
            finally { loading = false; }
            Recalculate(); RecordChange(); MarkDirty();
        }

        private static string AiSettingsPath { get { return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "DinkCel", "ai.json"); } }

        private AiConfiguration LoadAiSettings()
        {
            if (aiConfiguration != null) return aiConfiguration;
            try { if (File.Exists(AiSettingsPath))
                return aiConfiguration = new JavaScriptSerializer().Deserialize<AiConfiguration>(File.ReadAllText(AiSettingsPath)); }
            catch { }
            return aiConfiguration = new AiConfiguration();
        }

        private void OpenAiSettings()
        {
            AiConfiguration config = LoadAiSettings();
            using (var form = new Form { Text = "AI model connection", Width = 590, Height = 245,
                StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false })
            {
                var endpoint = new TextBox { Left = 120, Top = 18, Width = 440, Text = config.Endpoint };
                var model = new TextBox { Left = 120, Top = 52, Width = 440, Text = config.Model };
                var key = new TextBox { Left = 120, Top = 86, Width = 440, UseSystemPasswordChar = true };
                var enabled = new CheckBox { Left = 120, Top = 122, Width = 440,
                    Text = "Allow AI requests in this session", Checked = aiEnabled };
                form.Controls.AddRange(new Control[] { new Label { Left = 12, Top = 20, Text = "Endpoint" },
                    new Label { Left = 12, Top = 54, Text = "Model" },
                    new Label { Left = 12, Top = 88, Text = "API key" }, endpoint, model, key, enabled });
                var save = new Button { Text = "Save", Left = 468, Top = 155, Width = 92,
                    DialogResult = DialogResult.OK };
                form.Controls.Add(save); form.AcceptButton = save;
                if (form.ShowDialog(this) != DialogResult.OK) return;
                Uri url;
                if (!Uri.TryCreate(endpoint.Text, UriKind.Absolute, out url) ||
                    (url.Scheme != "https" && !(url.Scheme == "http" && url.IsLoopback)))
                { MessageBox.Show(this, "Use HTTPS or HTTP on localhost."); return; }
                config.Endpoint = endpoint.Text.Trim(); config.Model = model.Text.Trim();
                if (key.Text.Length > 0)
                    config.ProtectedKey = Convert.ToBase64String(ProtectedData.Protect(
                        Encoding.UTF8.GetBytes(key.Text), null, DataProtectionScope.CurrentUser));
                Directory.CreateDirectory(Path.GetDirectoryName(AiSettingsPath));
                File.WriteAllText(AiSettingsPath, new JavaScriptSerializer().Serialize(config));
                aiEnabled = enabled.Checked;
                aiFormulaCache.Clear();
                aiFormulaRequests = 0;
                Recalculate();
            }
        }

        private string GenerateAi(string prompt)
        {
            if (!aiEnabled) throw new InvalidOperationException("Enable AI in Script > AI settings first.");
            AiConfiguration config = LoadAiSettings();
            if (String.IsNullOrWhiteSpace(config.Model)) throw new InvalidOperationException("AI model is missing.");
            bool approved = false;
            Invoke((Action)(() => approved = MessageBox.Show(this,
                "Send this text to " + config.Endpoint + "?\n\n" + prompt,
                "AI request preview", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                MessageBoxDefaultButton.Button2) == DialogResult.Yes));
            if (!approved) throw new OperationCanceledException("AI request declined.");
            return AiClient.Send(config, prompt);
        }
    }
}
