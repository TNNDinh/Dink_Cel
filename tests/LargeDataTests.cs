using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace DinkCel
{
    internal static class LargeDataTests
    {
        [STAThread]
        private static void Main()
        {
            string path = Path.Combine(Path.GetTempPath(), "DinkCel_large_" + Guid.NewGuid().ToString("N") + ".csv");
            string native = Path.ChangeExtension(path, ".dinkcel");
            string xlsx = Path.ChangeExtension(path, ".xlsx");
            try
            {
                using (var writer = new StreamWriter(path, false, new UTF8Encoding(false)))
                {
                    writer.WriteLine("Group,Date,Amount");
                    for (int i = 1; i <= 20000; i++) writer.WriteLine("G" + (i % 20) + ",2024-01-01," + i);
                }
                Application.EnableVisualStyles();
                var watch = Stopwatch.StartNew();
                using (var form = new SpreadsheetForm(path))
                {
                    form.Show(); Application.DoEvents();
                    var grid = (DataGridView)typeof(SpreadsheetForm).GetField("grid", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(form);
                    if (grid.RowCount < 20001 || (string)grid[2, 20000].Value != "20000") throw new Exception("Large CSV import failed.");
                    Console.WriteLine("large import ms=" + watch.ElapsedMilliseconds);
                    watch.Restart();
                    grid[2, 20000].Value = "20001";
                    Console.WriteLine("large edit ms=" + watch.ElapsedMilliseconds);
                    if ((string)grid[2, 20000].Value != "20001") throw new Exception("Large edit failed.");
                    var pivot = new PivotDefinition { SourceRange = new System.Drawing.Rectangle(0, 0, 3, 20001) };
                    pivot.Rows.Add(new PivotAxisField { Column = 0 });
                    pivot.Values.Add(new PivotValueField { Column = 2, Aggregate = "Count" });
                    pivot.Values.Add(new PivotValueField { Column = 2, Aggregate = "Sum" });
                    watch.Restart();
                    PivotResult output = PivotEngine.Build(pivot, (r, c) =>
                        Convert.ToString(grid[c, r].Value) ?? "", 50000, 26);
                    if (output.Rows[output.Rows.Count - 1][1] != "20000" ||
                        output.Rows[output.Rows.Count - 1][2] != "200010001")
                        throw new Exception("Large pivot aggregation failed.");
                    Console.WriteLine("large pivot ms=" + watch.ElapsedMilliseconds);
                    string sourceName = ((System.Collections.Generic.List<SheetState>)typeof(SpreadsheetForm)
                        .GetField("sheets", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(form))[0].Name;
                    typeof(SpreadsheetForm).GetMethod("AddSheet", BindingFlags.NonPublic | BindingFlags.Instance).Invoke(form, null);
                    ((System.Collections.Generic.List<SheetState>)typeof(SpreadsheetForm)
                        .GetField("sheets", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(form))[1].Name = "Pivot1";
                    var live = new PivotDefinition { SourceSheet = sourceName, TargetSheet = "Pivot1",
                        SourceRange = new System.Drawing.Rectangle(0, 0, 3, 20001), GroupColumn = 0, ValueColumn = 2 };
                    live.Rows.Add(new PivotAxisField { Column = 0 });
                    live.Values.Add(new PivotValueField { Column = 2, Aggregate = "Sum" });
                    live.Values.Add(new PivotValueField { Column = 2, Aggregate = "Count" });
                    ((System.Collections.Generic.List<PivotDefinition>)typeof(SpreadsheetForm)
                        .GetField("pivots", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(form)).Add(live);
                    watch.Restart();
                    typeof(SpreadsheetForm).GetMethod("RefreshPivot", BindingFlags.NonPublic | BindingFlags.Instance)
                        .Invoke(form, new object[] { live });
                    if ((string)grid[1, 21].Value != "200010001" || (string)grid[2, 21].Value != "20000")
                        throw new Exception("Large pivot UI refresh failed.");
                    Console.WriteLine("large pivot UI ms=" + watch.ElapsedMilliseconds);
                    if (!(bool)typeof(SpreadsheetForm).GetMethod("WriteWorkbook", BindingFlags.NonPublic | BindingFlags.Instance)
                        .Invoke(form, new object[] { native })) throw new Exception("Large save failed.");
                    watch.Restart();
                    if (!(bool)typeof(SpreadsheetForm).GetMethod("WriteXlsx", BindingFlags.NonPublic | BindingFlags.Instance)
                        .Invoke(form, new object[] { xlsx })) throw new Exception("Large XLSX save failed.");
                    Console.WriteLine("large xlsx save ms=" + watch.ElapsedMilliseconds);
                    form.Close();
                }
                using (var reopened = new SpreadsheetForm(native))
                {
                    reopened.Show(); Application.DoEvents();
                    var grid = (DataGridView)typeof(SpreadsheetForm).GetField("grid", BindingFlags.NonPublic | BindingFlags.Instance).GetValue(reopened);
                    if (grid.RowCount < 20001 || (string)grid[2, 20000].Value != "20001")
                        throw new Exception("Large native reopen failed.");
                    reopened.Close();
                }
                var imported = XlsxFile.Read(xlsx, 50000, 26);
                if (imported.Sheets[0].Cells[20000 * 26 + 2].Text != "20001")
                    throw new Exception("Large XLSX reopen failed.");
            }
            catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }
            finally { if (File.Exists(path)) File.Delete(path); if (File.Exists(native)) File.Delete(native);
                if (File.Exists(xlsx)) File.Delete(xlsx); }
        }
    }
}
