using System;
using System.IO;
using System.Drawing;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Net.Sockets;
using System.Security.Cryptography;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.IO.Compression;
using System.Linq;

namespace DinkCel
{
    internal static class OpenFileTests
    {
        [STAThread]
        private static void Main()
        {
            try { Run(); }
            catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }
        }

        private static void Run()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            string directory = Path.Combine(Path.GetTempPath(),
                "DinkCelOpenTest_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                string csv = Path.Combine(directory, "sample.csv");
                File.WriteAllText(csv, "name,value\r\n\"A,B\",42\r\n", new UTF8Encoding(true));
                using (var form = new SpreadsheetForm(csv))
                {
                    form.Show();
                    Application.DoEvents();
                    var grid = (DataGridView)Field(form, "grid");
                    Equal("name", grid[0, 0].Value);
                    Equal("A,B", grid[0, 1].Value);
                    Equal("42", grid[1, 1].Value);
                    Equal(csv, Field(form, "currentPath"));
                    Rectangle rowHeader = grid.GetCellDisplayRectangle(-1, 0, false);
                    Rectangle dropRow = grid.GetCellDisplayRectangle(-1, 1, false);
                    var leftClick = new MouseEventArgs(MouseButtons.Left, 1, 0, 0, 0);
                    Invoke(form, "GridCellMouseDown", grid,
                        new DataGridViewCellMouseEventArgs(-1, 0,
                            rowHeader.Width / 2, rowHeader.Height / 2, leftClick));
                    SetField(form, "headerStartPoint", new Point(
                        rowHeader.Left + rowHeader.Width / 2,
                        rowHeader.Top + rowHeader.Height / 2));
                    var drop = new MouseEventArgs(MouseButtons.Left, 1,
                        dropRow.Left + dropRow.Width / 2,
                        dropRow.Top + dropRow.Height / 2, 0);
                    Invoke(form, "GridMouseMove", grid, drop);
                    Invoke(form, "GridMouseUp", grid, drop);
                    Application.DoEvents();
                    Equal("name", grid[0, 1].Value);
                    Call(form, "Undo");
                    Equal("name", grid[0, 0].Value);

                    Rectangle columnHeader = grid.GetCellDisplayRectangle(1, -1, false);
                    Rectangle dropColumn = grid.GetCellDisplayRectangle(2, -1, false);
                    Invoke(form, "GridCellMouseDown", grid,
                        new DataGridViewCellMouseEventArgs(1, -1,
                            columnHeader.Width / 2, columnHeader.Height / 2, leftClick));
                    SetField(form, "headerStartPoint", new Point(
                        columnHeader.Left + columnHeader.Width / 2,
                        columnHeader.Top + columnHeader.Height / 2));
                    var dropColumnEvent = new MouseEventArgs(MouseButtons.Left, 1,
                        dropColumn.Left + dropColumn.Width / 2,
                        dropColumn.Top + dropColumn.Height / 2, 0);
                    Invoke(form, "GridMouseMove", grid, dropColumnEvent);
                    Invoke(form, "GridMouseUp", grid, dropColumnEvent);
                    Application.DoEvents();
                    Equal("42", grid[2, 1].Value);
                    Call(form, "Undo");
                    Equal("42", grid[1, 1].Value);

                    grid.ClearSelection();
                    grid[0, 1].Selected = true;
                    grid[1, 1].Selected = true;
                    Invoke(form, "GridKeyDown", grid,
                        new KeyEventArgs(Keys.Delete));
                    Equal(null, grid[0, 1].Value);
                    Equal(null, grid[1, 1].Value);
                    Call(form, "Undo");
                    Equal("A,B", grid[0, 1].Value);
                    Equal("42", grid[1, 1].Value);
                    Equal(false, Field(form, "dirty"));
                    Call(form, "Redo");
                    Equal(null, grid[0, 1].Value);
                    Equal(null, grid[1, 1].Value);
                    Equal(true, Field(form, "dirty"));
                    Call(form, "Undo");
                    Equal("A,B", grid[0, 1].Value);
                    Equal("42", grid[1, 1].Value);
                    Equal(false, Field(form, "dirty"));
                    grid[4, 0].Value = "new";
                    Call(form, "Undo");
                    Equal(null, grid[4, 0].Value);
                    Call(form, "Redo");
                    Equal("new", grid[4, 0].Value);
                    Call(form, "Undo");
                    grid[2, 1].Value = "=B2";
                    grid[0, 1].Style.BackColor = Color.Yellow;
                    grid.Rows[1].Height = 39;
                    Call(form, "MoveHeader", true, 1, 3);
                    Equal("A,B", grid[0, 3].Value);
                    Equal("42", grid[1, 3].Value);
                    Equal("=B4", grid[2, 3].Value);
                    Equal(Color.Yellow, grid[0, 3].Style.BackColor);
                    Equal(39, grid.Rows[3].Height);
                    Call(form, "Undo");
                    Equal("A,B", grid[0, 1].Value);
                    Equal("=B2", grid[2, 1].Value);
                    Equal(39, grid.Rows[1].Height);
                    Call(form, "Redo");
                    Equal("A,B", grid[0, 3].Value);
                    Call(form, "Undo");
                    Call(form, "SelectHeader", 1, true);
                    Equal(26, grid.SelectedCells.Count);
                    Call(form, "InsertRow");
                    Equal(null, grid[0, 1].Value);
                    Equal("A,B", grid[0, 2].Value);
                    Equal("=B3", grid[2, 2].Value);
                    Equal(Color.Yellow, grid[0, 2].Style.BackColor);
                    Call(form, "Undo");
                    Equal("A,B", grid[0, 1].Value);
                    Equal("=B2", grid[2, 1].Value);
                    Call(form, "Redo");
                    Equal("A,B", grid[0, 2].Value);
                    Equal("=B3", grid[2, 2].Value);
                    Call(form, "DeleteRow");
                    Equal("A,B", grid[0, 1].Value);
                    Equal("=B2", grid[2, 1].Value);

                    grid[3, 0].Value = "=$B$2";
                    grid.Columns[1].Width = 151;
                    Call(form, "MoveHeader", false, 1, 3);
                    Equal("42", grid[3, 1].Value);
                    Equal("=$D$2", grid[2, 0].Value);
                    Equal(151, grid.Columns[3].Width);
                    Call(form, "Undo");
                    Equal("42", grid[1, 1].Value);
                    Equal("=$B$2", grid[3, 0].Value);
                    Equal(151, grid.Columns[1].Width);
                    Call(form, "SelectHeader", 1, false);
                    Equal(grid.RowCount, grid.SelectedCells.Count);
                    Call(form, "InsertColumn");
                    Equal(null, grid[1, 1].Value);
                    Equal("42", grid[2, 1].Value);
                    Equal("=$C$2", grid[4, 0].Value);
                    Call(form, "DeleteColumn");
                    Equal("42", grid[1, 1].Value);
                    Equal("=$B$2", grid[3, 0].Value);
                    grid[0, 1].Value = "A,B\r\n\"Z\"";
                    Equal(true, Invoke(form, "SaveDocument"));
                    Equal(csv, Field(form, "currentPath"));
                    Equal(false, Field(form, "dirty"));
                    grid[0, 0].Value = "changed";
                    Call(form, "Undo");
                    Equal("name", grid[0, 0].Value);
                    Equal(false, Field(form, "dirty"));
                    var csvRows = CsvFile.Read(csv, 200, 26);
                    Equal("A,B\r\n\"Z\"", csvRows[1][0]);
                    Equal("=B2", csvRows[1][2]);
                    Equal("=$B$2", csvRows[0][3]);
                    byte[] bytes = File.ReadAllBytes(csv);
                    Equal((byte)0xEF, bytes[0]);
                    Equal((byte)0xBB, bytes[1]);
                    Equal((byte)0xBF, bytes[2]);
                    string edited = Path.Combine(directory, "edited.dinkcel");
                    Equal(true, Invoke(form, "WriteWorkbook", edited));
                    form.Close();
                    using (var reopened = new SpreadsheetForm(edited))
                    {
                        reopened.Show();
                        Application.DoEvents();
                        var reopenedGrid = (DataGridView)Field(reopened, "grid");
                        Equal("A,B\n\"Z\"", reopenedGrid[0, 1].Value);
                        Equal("=$B$2", reopenedGrid[3, 0].Value);
                        Equal(Color.Yellow, reopenedGrid[0, 1].Style.BackColor);
                        reopened.Close();
                    }
                }

                string native = Path.Combine(directory, "old.dinkcel");
                File.WriteAllText(native,
                    "<workbook rows=\"20\" columns=\"10\"><cell row=\"2\" column=\"3\">old</cell></workbook>");
                using (var form = new SpreadsheetForm(native))
                {
                    form.Show();
                    Application.DoEvents();
                    var grid = (DataGridView)Field(form, "grid");
                    Equal("old", grid[2, 1].Value);
                    Equal(native, Field(form, "currentPath"));
                    form.Close();
                }
                string scripted = Path.Combine(directory, "scripted.dinkcel");
                using (var form = new SpreadsheetForm(null))
                {
                    form.Show(); Application.DoEvents();
                    SetField(form, "scriptCode", "/** @customfunction */ function DOUBLE(x) { return x * 2; }");
                    SetField(form, "scriptEnabled", true);
                    var grid = (DataGridView)Field(form, "grid");
                    grid[0, 0].Value = "7";
                    grid[1, 0].Value = "=DOUBLE(A1)";
                    form.GetType().GetMethod("Recalculate", BindingFlags.Instance |
                        BindingFlags.NonPublic, null, Type.EmptyTypes, null).Invoke(form, null);
                    var calculated = (System.Collections.Generic.Dictionary<int, string>)Field(form, "calculated");
                    Equal("14", calculated[1]);
                    grid[0, 0].Value = "8";
                    form.GetType().GetMethod("Recalculate", BindingFlags.Instance |
                        BindingFlags.NonPublic, null, Type.EmptyTypes, null).Invoke(form, null);
                    Equal("16", calculated[1]);
                    var writes = new System.Collections.Generic.List<ScriptCellWrite> {
                        new ScriptCellWrite { Row = 0, Column = 2, Value = "script" },
                        new ScriptCellWrite { Row = 1, Column = 2, Value = 42 }
                    };
                    Call(form, "ApplyScriptWrites", writes);
                    Equal("script", grid[2, 0].Value);
                    Equal("42", grid[2, 1].Value);
                    Call(form, "Undo");
                    Equal(null, grid[2, 0].Value);
                    Equal(null, grid[2, 1].Value);
                    Call(form, "Redo");
                    Equal("script", grid[2, 0].Value);
                    Equal(true, Invoke(form, "WriteWorkbook", scripted));
                    form.Close();
                }
                using (var form = new SpreadsheetForm(scripted))
                {
                    form.Show(); Application.DoEvents();
                    Equal(true, ((string)Field(form, "scriptCode")).Contains("DOUBLE"));
                    Equal(false, Field(form, "scriptEnabled"));
                    form.Close();
                }
                TestAiClient();
                TestDynamicArrayUi(directory);
                TestAdvancedExcel(directory);
                TestNamedObjects(directory);
                Console.WriteLine("Open file: CSV and older .dinkcel passed.");
            }
            finally
            {
                Directory.Delete(directory, true);
            }
        }

        private static void TestAdvancedExcel(string directory)
        {
            double solved;
            Equal(true, GoalSeekSolver.Solve(x => x * x, 25, 2, out solved));
            Equal(true, Math.Abs(solved - 5) < 1e-6);
            Equal("A,B", AdvancedData.SplitLine("\"A,B\",C", ',')[0]);
            Equal("C", AdvancedData.SplitLine("\"A,B\",C", ',')[1]);
            Equal(2, AdvancedData.UniqueRows(new[] { new[] { "A", "1" },
                new[] { "a", "1" }, new[] { "B", "2" } }, false).Count);
            Equal(2, AdvancedData.UniqueRows(new[] { new[] { "A\u001fB", "C" },
                new[] { "A", "B\u001fC" } }, false).Count);
            var pattern = AdvancedData.InferFlashFill(new[] {
                Tuple.Create("John Smith", "", "John"), Tuple.Create("Jane Doe", "", "Jane") });
            Equal("Alan", pattern("Alan Turing", ""));
            var criteria = new[] {
                new System.Collections.Generic.Dictionary<string, string> { { "Amount", ">20" } },
                new System.Collections.Generic.Dictionary<string, string> { { "City", "Hanoi" } } };
            Equal(true, AdvancedData.MatchesCriteria(new[] { "Hue", "30" },
                new[] { "City", "Amount" }, criteria));
            Equal(false, AdvancedData.MatchesCriteria(new[] { "Hue", "10" },
                new[] { "City", "Amount" }, criteria));
            string path = Path.Combine(directory, "scenarios.dinkcel");
            using (var form = new SpreadsheetForm(null))
            {
                form.Show(); Application.DoEvents();
                var grid = (DataGridView)Field(form, "grid");
                grid[0, 0].Value = "2"; grid[1, 0].Value = "=A1*2";
                double value = (double)Invoke(form, "EvaluateWhatIf", 0, 1,
                    new System.Collections.Generic.Dictionary<int, string> { { 0, "50" } });
                Equal(100.0, value);
                var scenario = new ScenarioDefinition { Name = "High", Sheet = "Sheet1" };
                scenario.Values[0] = "50";
                ((System.Collections.Generic.List<ScenarioDefinition>)Field(form, "scenarios")).Add(scenario);
                Equal(true, Invoke(form, "WriteWorkbook", path));
                form.Close();
            }
            using (var form = new SpreadsheetForm(path))
            {
                form.Show(); Application.DoEvents();
                var scenarios = (System.Collections.Generic.List<ScenarioDefinition>)Field(form, "scenarios");
                Equal(1, scenarios.Count); Equal("50", scenarios[0].Values[0]);
                form.Close();
            }
        }

        private static void TestNamedObjects(string directory)
        {
            string native = Path.Combine(directory, "names.dinkcel");
            string xlsx = Path.Combine(directory, "names.xlsx");
            using (var form = new SpreadsheetForm(null))
            {
                form.Show(); Application.DoEvents();
                var names = (System.Collections.Generic.List<NamedRange>)Field(form, "namedRanges");
                names.Add(new NamedRange { Name = "Scale", Formula = "=A1*2", Sheet = "Sheet1" });
                names.Add(new NamedRange { Name = "LoopA", Formula = "=LoopB+1", Sheet = "Sheet1" });
                names.Add(new NamedRange { Name = "LoopB", Formula = "=LoopA+1", Sheet = "Sheet1" });
                var grid = (DataGridView)Field(form, "grid");
                grid[0, 0].Value = "4";
                grid[1, 0].Value = "=Scale";
                grid[2, 0].Value = "=LoopA";
                form.GetType().GetMethod("Recalculate", BindingFlags.Instance |
                    BindingFlags.NonPublic, null, Type.EmptyTypes, null).Invoke(form, null);
                Equal("8", grid[1, 0].FormattedValue);
                Equal("#CYCLE!", grid[2, 0].FormattedValue);
                Call(form, "AddSheet");
                grid = (DataGridView)Field(form, "grid");
                string secondSheet = ((System.Collections.Generic.List<SheetState>)Field(form, "sheets"))[1].Name;
                names.Add(new NamedRange { Name = "Scale", Formula = "=A1*3",
                    ScopeSheet = secondSheet, Sheet = secondSheet });
                names.Add(new NamedRange { Name = "Origin", Sheet = "Sheet1",
                    ScopeSheet = secondSheet, Range = new Rectangle(0, 0, 1, 1) });
                grid[0, 0].Value = "5";
                grid[1, 0].Value = "=Scale";
                grid[2, 0].Value = "=Origin";
                form.GetType().GetMethod("Recalculate", BindingFlags.Instance |
                    BindingFlags.NonPublic, null, Type.EmptyTypes, null).Invoke(form, null);
                Equal("15", grid[1, 0].FormattedValue);
                Equal("4", grid[2, 0].FormattedValue);
                Equal(true, Invoke(form, "WriteWorkbook", native));
                Equal(true, Invoke(form, "WriteXlsx", xlsx));
                form.Close();
            }
            foreach (string path in new[] { native, xlsx })
                using (var form = new SpreadsheetForm(path))
                {
                    form.Show(); Application.DoEvents();
                    var grid = (DataGridView)Field(form, "grid");
                    Equal("8", grid[1, 0].FormattedValue);
                    Equal("#CYCLE!", grid[2, 0].FormattedValue);
                    Call(form, "SwitchSheet", 1);
                    Equal("15", grid[1, 0].FormattedValue);
                    Equal("4", grid[2, 0].FormattedValue);
                    form.Close();
                }
        }

        private static void TestDynamicArrayUi(string directory)
        {
            string native = Path.Combine(directory, "dynamic.dinkcel");
            string xlsx = Path.Combine(directory, "dynamic.xlsx");
            using (var form = new SpreadsheetForm(null))
            {
                form.Show(); Application.DoEvents();
                var grid = (DataGridView)Field(form, "grid");
                grid[4, 0].Value = "=SEQUENCE(2,2)";
                Equal("1", grid[4, 0].FormattedValue);
                Equal("4", grid[5, 1].FormattedValue);
                grid[5, 1].Value = "blocked";
                Equal("#SPILL!", grid[4, 0].FormattedValue);
                grid[5, 1].Value = null;
                Equal("4", grid[5, 1].FormattedValue);
                Invoke(form, "SelectRectangle", new Rectangle(5, 1, 1, 1), 5, 1, false);
                Invoke(form, "CopySelected");
                Invoke(form, "SelectRectangle", new Rectangle(8, 0, 1, 1), 8, 0, false);
                Invoke(form, "PasteSelected");
                Equal("4", grid[8, 0].Value);
                Equal(true, Invoke(form, "WriteWorkbook", native));
                Equal(true, Invoke(form, "WriteXlsx", xlsx));
                form.Close();
            }
            using (var archive = ZipFile.OpenRead(xlsx))
            using (var stream = archive.GetEntry("xl/worksheets/sheet1.xml").Open())
            {
                var sheetXml = System.Xml.Linq.XDocument.Load(stream);
                var formula = sheetXml.Descendants().First(x => x.Name.LocalName == "f");
                Equal("_xlfn.SEQUENCE(2,2)", formula.Value);
                Equal("E1:F2", (string)formula.Attribute("ref"));
            }
            using (var form = new SpreadsheetForm(native))
            {
                form.Show(); Application.DoEvents();
                var grid = (DataGridView)Field(form, "grid");
                Equal("=SEQUENCE(2,2)", grid[4, 0].Value);
                Equal("4", grid[5, 1].FormattedValue);
                form.Close();
            }
            using (var form = new SpreadsheetForm(xlsx))
            {
                form.Show(); Application.DoEvents();
                var grid = (DataGridView)Field(form, "grid");
                Equal("=SEQUENCE(2,2)", grid[4, 0].Value);
                Equal("4", grid[5, 1].FormattedValue);
                form.Close();
            }
        }

        private static void TestAiClient()
        {
            var server = new TcpListener(IPAddress.Loopback, 0);
            try
            {
                server.Start();
                int port = ((IPEndPoint)server.LocalEndpoint).Port;
                string received = null, authorization = null;
                var responseTask = Task.Run(() =>
                {
                    using (var client = server.AcceptTcpClient())
                    using (var stream = client.GetStream())
                    {
                        var reader = new StreamReader(stream, Encoding.UTF8);
                        int length = 0;
                        string line;
                        while (!String.IsNullOrEmpty(line = reader.ReadLine()))
                        {
                            if (line.StartsWith("Content-Length:", StringComparison.OrdinalIgnoreCase))
                                length = Int32.Parse(line.Substring(15).Trim());
                            if (line.StartsWith("Authorization:", StringComparison.OrdinalIgnoreCase))
                                authorization = line.Substring(14).Trim();
                        }
                        var body = new char[length];
                        int count = 0;
                        while (count < body.Length) count += reader.Read(body, count, body.Length - count);
                        received = new string(body);
                        byte[] payload = Encoding.UTF8.GetBytes("{\"choices\":[{\"message\":{\"content\":\"AI answer\"}}]}");
                        byte[] header = Encoding.ASCII.GetBytes("HTTP/1.1 200 OK\r\nContent-Type: application/json\r\nContent-Length: " + payload.Length + "\r\nConnection: close\r\n\r\n");
                        stream.Write(header, 0, header.Length);
                        stream.Write(payload, 0, payload.Length);
                    }
                });
                string secret = "test-private-key";
                var settings = new AiConfiguration { Endpoint = "http://localhost:" + port + "/v1/chat/completions",
                    Model = "test-model", ProtectedKey = Convert.ToBase64String(ProtectedData.Protect(
                        Encoding.UTF8.GetBytes(secret), null, DataProtectionScope.CurrentUser)) };
                string settingsJson = new JavaScriptSerializer().Serialize(settings);
                Equal(false, settingsJson.Contains(secret));
                Equal(settings.ProtectedKey,
                    new JavaScriptSerializer().Deserialize<AiConfiguration>(settingsJson).ProtectedKey);
                Equal("AI answer", AiClient.Send(settings, "only this prompt"));
                responseTask.Wait();
                Equal("Bearer " + secret, authorization);
                Equal(true, received.Contains("only this prompt"));
                Equal(false, received.Contains("unrelated sheet data"));
            }
            finally { server.Stop(); }
        }

        private static object Field(object target, string name)
        {
            return target.GetType().GetField(name,
                BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target);
        }

        private static void SetField(object target, string name, object value)
        {
            target.GetType().GetField(name,
                BindingFlags.Instance | BindingFlags.NonPublic).SetValue(target, value);
        }

        private static void Call(object target, string name, params object[] args)
        {
            Invoke(target, name, args);
        }

        private static object Invoke(object target, string name, params object[] args)
        {
            return target.GetType().GetMethod(name,
                BindingFlags.Instance | BindingFlags.NonPublic).Invoke(target, args);
        }

        private static void Equal(object expected, object actual)
        {
            if (!object.Equals(expected, actual))
                throw new Exception("Expected " + expected + ", got " + actual);
        }
    }
}
