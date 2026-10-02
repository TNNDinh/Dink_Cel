using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace DinkCel
{
    internal static class V3UiTests
    {
        private static object Field(object target, string name)
        {
            return target.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(target);
        }
        private static object Call(object target, string name, params object[] args)
        {
            return target.GetType().GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance).Invoke(target, args);
        }
        private static void Equal(object expected, object actual)
        {
            if (!object.Equals(expected, actual)) throw new Exception("Expected " + expected + ", got " + actual);
        }

        [STAThread]
        private static void Main()
        {
            try { Run(); }
            catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }
        }

        private static void Run()
        {
            EmbeddedDependencies.Install();
            var rename = typeof(SpreadsheetForm).GetMethod("RenameSheetReferences",
                BindingFlags.NonPublic | BindingFlags.Static);
            Equal("='Sales Data'!A1+\"Sheet1!B1\"+'Sales Data'!B2",
                rename.Invoke(null, new object[] { "=Sheet1!A1+\"Sheet1!B1\"+'Sheet1'!B2",
                    "Sheet1", "Sales Data" }));
            Application.EnableVisualStyles();
            string path = Path.Combine(Path.GetTempPath(), "DinkCel_v3ui_" + Guid.NewGuid().ToString("N") + ".dinkcel");
            try
            {
                using (var form = new SpreadsheetForm(null))
                {
                    form.Show(); Application.DoEvents();
                    var grid = (DataGridView)Field(form, "grid");
                    grid[0, 0].Value = "Category"; grid[1, 0].Value = "Amount";
                    grid[0, 1].Value = "Red"; grid[1, 1].Value = "10";
                    grid[0, 2].Value = "Blue"; grid[1, 2].Value = "4";
                    grid[0, 3].Value = "Red"; grid[1, 3].Value = "6";
                    var chart = new ChartDefinition { Title = "Sales", Kind = "Column",
                        Range = new Rectangle(0, 0, 2, 4) };
                    using (var control = (Chart)Call(form, "BuildChart", chart))
                    { Equal(1, control.Series.Count); Equal(3, control.Series[0].Points.Count); }
                    ((List<ChartDefinition>)Field(form, "charts")).Add(chart);
                    ((List<TableDefinition>)Field(form, "tables")).Add(new TableDefinition
                    { Name = "SalesTable", Range = new Rectangle(0, 0, 2, 4) });
                    ((List<NamedRange>)Field(form, "namedRanges")).Add(new NamedRange
                    { Name = "Amounts", Sheet = "Sheet1", Range = new Rectangle(1, 1, 1, 3) });
                    Call(form, "AddSheet");
                    ((List<SheetState>)Field(form, "sheets"))[1].Name = "Pivot1";
                    var pivot = new PivotDefinition { SourceSheet = "Sheet1", TargetSheet = "Pivot1",
                        SourceRange = new Rectangle(0, 0, 2, 4), GroupColumn = 0, ValueColumn = 1 };
                    ((List<PivotDefinition>)Field(form, "pivots")).Add(pivot);
                    Call(form, "RefreshPivot", pivot);
                    Equal("Blue", grid[0, 1].Value);
                    Equal("4", grid[1, 1].Value);
                    Equal("Red", grid[0, 2].Value);
                    Equal("16", grid[1, 2].Value);
                    var validation = new ValidationRule { Range = new Rectangle(2, 0, 1, 1) };
                    validation.Choices.Add("Yes"); validation.Choices.Add("No");
                    ((List<ValidationRule>)Field(form, "validations")).Add(validation);
                    grid[2, 0].Value = "Maybe";
                    Equal(null, grid[2, 0].Value);
                    grid[2, 0].Value = "Yes";
                    Equal("Yes", grid[2, 0].Value);
                    grid[3, 0].Value = "=SUM(Amounts)";
                    Equal("20", grid[3, 0].FormattedValue);
                    Call(form, "SwitchSheet", 0);
                    grid[1, 1].Value = "20";
                    Call(form, "RefreshPivot", pivot);
                    Equal("26", grid[1, 2].Value);
                    Equal("30", grid[3, 0].FormattedValue);
                    Equal(true, Call(form, "WriteWorkbook", path));
                    form.Close();
                }
                using (var reopened = new SpreadsheetForm(path))
                {
                    reopened.Show(); Application.DoEvents();
                    Equal(1, ((List<TableDefinition>)Field(reopened, "tables")).Count);
                    Equal(1, ((List<ChartDefinition>)Field(reopened, "charts")).Count);
                    Equal(1, ((List<NamedRange>)Field(reopened, "namedRanges")).Count);
                    Equal(1, ((List<PivotDefinition>)Field(reopened, "pivots")).Count);
                    Call(reopened, "SwitchSheet", 1);
                    var grid = (DataGridView)Field(reopened, "grid");
                    Equal("26", grid[1, 2].Value);
                    Equal("30", grid[3, 0].FormattedValue);
                    Equal(1, ((List<ValidationRule>)Field(reopened, "validations")).Count);
                    reopened.Close();
                }
                Console.WriteLine("v0.3 UI: chart, pivot refresh, validation and named ranges passed.");
            }
            finally { if (File.Exists(path)) File.Delete(path); }
        }
    }
}
