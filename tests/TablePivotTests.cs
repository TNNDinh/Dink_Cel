using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace DinkCel
{
    internal static class TablePivotTests
    {
        private static void Check(bool condition, string message)
        {
            if (!condition) throw new Exception(message);
        }

        private static object Field(object target, string name)
        {
            return target.GetType().GetField(name, BindingFlags.Instance | BindingFlags.NonPublic).GetValue(target);
        }

        private static object Call(object target, string name, params object[] args)
        {
            return target.GetType().GetMethod(name, BindingFlags.Instance | BindingFlags.NonPublic).Invoke(target, args);
        }

        [STAThread]
        private static void Main()
        {
            try { Run(); Console.WriteLine("v0.8 Table/Pivot: structured references, expand, aggregation, grouping and persistence passed."); }
            catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }
        }

        private static void Run()
        {
            string[,] input = {
                { "Region", "Product", "Date", "Amount", "Units" },
                { "North", "A", "2024-01-01", "10", "2" },
                { "North", "A", "2024-01-15", "20", "3" },
                { "North", "B", "2024-02-01", "5", "4" },
                { "South", "A", "2024-01-01", "7", "1" },
                { "South", "B", "2024-02-02", "8", "5" }
            };
            var pivot = new PivotDefinition { SourceRange = new Rectangle(0, 0, 5, 6) };
            pivot.Rows.Add(new PivotAxisField { Column = 0 });
            pivot.Rows.Add(new PivotAxisField { Column = 1 });
            pivot.Columns.Add(new PivotAxisField { Column = 2, DateGroup = "Month" });
            foreach (string aggregate in new[] { "Sum", "Count", "Average", "Min", "Max" })
                pivot.Values.Add(new PivotValueField { Column = 3, Aggregate = aggregate });
            Func<int, int, string> read = (r, c) => input[r, c];
            PivotResult result = PivotEngine.Build(pivot, read, 100, 26);
            int january = Array.IndexOf(result.Rows[0], "2024-01 · Sum Amount");
            int february = Array.IndexOf(result.Rows[0], "2024-02 · Sum Amount");
            int total = Array.IndexOf(result.Rows[0], "Tổng · Sum Amount");
            Check(january > 0 && february > 0 && total > 0, "month column headers");
            string[] grand = result.Rows.Last();
            Check(grand[0] == "Tổng cộng" && grand[january] == "37" && grand[february] == "13" &&
                grand[total] == "50", "grand totals");
            Check(grand[total + 1] == "5" && grand[total + 2] == "10" &&
                grand[total + 3] == "5" && grand[total + 4] == "20", "count average min max");
            pivot.SortByValue = true;
            result = PivotEngine.Build(pivot, read, 100, 26);
            Check(result.Rows[1][0].Contains("South"), "sort groups by value");
            pivot.SortByValue = false;
            pivot.Columns[0].DateGroup = "Year";
            result = PivotEngine.Build(pivot, read, 100, 26);
            Check(result.Rows[0].Any(v => v != null && v.StartsWith("2024 · Sum")), "year grouping");
            pivot.Columns[0].DateGroup = "Day";
            result = PivotEngine.Build(pivot, read, 100, 26);
            Check(result.Rows[0].Any(v => v != null && v.StartsWith("2024-01-01 · Sum")), "day grouping");
            pivot.Columns[0].DateGroup = "Month";
            pivot.Filters.Add(new PivotFilterField { Column = 0, Value = "North" });
            result = PivotEngine.Build(pivot, read, 100, 26);
            Check(result.Rows.Last()[total] == "35", "pivot filter");
            pivot.Collapsed.Add("North");
            result = PivotEngine.Build(pivot, read, 100, 26);
            Check(result.Rows.Any(r => r[0].Contains("▶ North") && r[total] == "35") &&
                !result.Rows.Any(r => r[0].Trim() == "A"), "pivot collapse");

            Check(FormulaEngine.ShiftReferences("=SUM(Table1[Doanh thu])+A1", 1, 0, 50000, 26) ==
                "=SUM(Table1[Doanh thu])+A2", "structured reference survives fill");
            Check(FormulaEngine.ShiftStructureReferences("=SUM(Table1[A1])+A2", true, 0, true, 50000, 26) ==
                "=SUM(Table1[A1])+A3", "structured reference survives insert");
            Check(FormulaEngine.MoveStructureReferences("=SUM(Table1[A1])+A2", true, 1, 2) ==
                "=SUM(Table1[A1])+A3", "structured reference survives move");
            string native = Path.Combine(Path.GetTempPath(), "DinkCel_table_" + Guid.NewGuid().ToString("N") + ".dinkcel");
            string xlsx = Path.ChangeExtension(native, ".xlsx");
            try
            {
                Application.EnableVisualStyles();
                using (var form = new SpreadsheetForm(null))
                {
                    form.Show(); Application.DoEvents();
                    var grid = (DataGridView)Field(form, "grid");
                    grid[0, 0].Value = "Item"; grid[1, 0].Value = "Doanh thu"; grid[2, 0].Value = "Gấp đôi";
                    grid[0, 1].Value = "A"; grid[1, 1].Value = "10";
                    grid[0, 2].Value = "B"; grid[1, 2].Value = "20";
                    var table = new TableDefinition { Name = "Table1", Range = new Rectangle(0, 0, 3, 3),
                        Style = "TableStyleMedium9", BandedRows = true, BandedColumns = true };
                    table.Filters.Add(new FilterCriterion { Column = 0, Operator = "One Of", Value1 = "A\nB" });
                    ((List<TableDefinition>)Field(form, "tables")).Add(table);
                    grid[4, 0].Value = "=SUM(Table1[Doanh thu])";
                    Check(Convert.ToString(grid[4, 0].FormattedValue) == "30", "SUM structured reference");
                    grid[2, 1].Value = "=[@Doanh thu]*2";
                    Check(Convert.ToString(grid[2, 2].FormattedValue) == "40", "calculated column");
                    grid[1, 3].Value = "30";
                    Check(table.Range.Bottom == 4 && Convert.ToString(grid[2, 3].FormattedValue) == "60" &&
                        Convert.ToString(grid[4, 0].FormattedValue) == "60", "table auto expand");
                    Check((bool)Call(form, "AddTotalRow", table) && table.TotalRow && table.Range.Bottom == 5 &&
                        Convert.ToString(grid[1, 4].Value) == "=SUM(Table1[Doanh thu])", "table total row");
                    Check((bool)Call(form, "WriteWorkbook", native), "write native table");
                    Check((bool)Call(form, "WriteXlsx", xlsx), "write xlsx table");
                    form.Close();
                }
                var loaded = XlsxFile.Read(xlsx, 200, 26).Sheets[0].Tables[0];
                Check(loaded.Style == "TableStyleMedium9" && loaded.BandedColumns &&
                    loaded.CalculatedColumns.ContainsKey(2) && loaded.Filters.Count == 1 &&
                    loaded.Range.Bottom == 5 && loaded.TotalRow && loaded.Filters[0].Operator == "One Of" &&
                    DataTools.Match("B", "Text", loaded.Filters[0].Operator, loaded.Filters[0].Value1, ""),
                    "xlsx table metadata");
                using (var reopened = new SpreadsheetForm(native))
                {
                    reopened.Show(); Application.DoEvents();
                    var table = ((List<TableDefinition>)Field(reopened, "tables"))[0];
                    Check(table.Range.Bottom == 5 && table.TotalRow && table.BandedColumns && table.Filters.Count == 1 &&
                        table.CalculatedColumns.ContainsKey(2), "native table metadata");
                    var grid = (DataGridView)Field(reopened, "grid");
                    Check(Convert.ToString(grid[4, 0].FormattedValue) == "60", "native structured reference");
                    var pivotSheet = new PivotDefinition { SourceSheet = "Sheet1", SourceTable = "Table1",
                        SourceRange = table.Range, TargetSheet = "Pivot1", GroupColumn = 0,
                        ValueColumn = 1, Aggregate = "Sum" };
                    pivotSheet.Rows.Add(new PivotAxisField { Column = 0 });
                    pivotSheet.Values.Add(new PivotValueField { Column = 1, Aggregate = "Sum" });
                    pivotSheet.Filters.Add(new PivotFilterField { Column = 0, Value = "A" });
                    Call(reopened, "AddSheet");
                    ((List<SheetState>)Field(reopened, "sheets"))[1].Name = "Pivot1";
                    ((List<PivotDefinition>)Field(reopened, "pivots")).Add(pivotSheet);
                    Call(reopened, "RefreshPivot", pivotSheet);
                    Check(Convert.ToString(grid[1, 2].Value) == "10", "pivot UI refresh with filter");
                    Check((bool)Call(reopened, "WriteWorkbook", native), "persist advanced pivot");
                    reopened.Close();
                }
                using (var final = new SpreadsheetForm(native))
                {
                    final.Show(); Application.DoEvents();
                    var pivots = (List<PivotDefinition>)Field(final, "pivots");
                    Check(pivots.Count == 1 && pivots[0].SourceTable == "Table1" &&
                        pivots[0].Rows.Count == 1 && pivots[0].Values.Count == 1 &&
                        pivots[0].Filters.Count == 1, "advanced pivot metadata");
                    Call(final, "SwitchSheet", 1);
                    var grid = (DataGridView)Field(final, "grid");
                    Check(Convert.ToString(grid[1, 2].Value) == "10", "persisted pivot result");
                    final.Close();
                }
            }
            finally
            {
                if (File.Exists(native)) File.Delete(native);
                if (File.Exists(xlsx)) File.Delete(xlsx);
            }
            using (var form = new SpreadsheetForm(null))
            {
                form.Show(); Application.DoEvents();
                var grid = (DataGridView)Field(form, "grid");
                grid[0, 0].Value = "Name"; grid[1, 0].Value = "Amount";
                grid[0, 1].Value = "A"; grid[1, 1].Value = "10";
                grid[0, 2].Value = "B"; grid[1, 2].Value = "20";
                var table = new TableDefinition { Name = "Table1", Range = new Rectangle(0, 0, 2, 3) };
                table.Filters.Add(new FilterCriterion { Column = 1, Operator = "Greater", Value1 = "0" });
                ((List<TableDefinition>)Field(form, "tables")).Add(table);
                grid[3, 0].Value = "=SUM(Table1[Amount])";
                grid.CurrentCell = grid[0, 1];
                Call(form, "InsertRow");
                Check(table.Range.Height == 4 && Convert.ToString(grid[0, 2].Value) == "A" &&
                    Convert.ToString(grid[3, 0].Value) == "=SUM(Table1[Amount])", "insert inside table");
                Call(form, "DeleteRow");
                Check(table.Range.Height == 3 && Convert.ToString(grid[0, 1].Value) == "A", "delete inside table");
                Call(form, "Undo");
                Check(((List<TableDefinition>)Field(form, "tables"))[0].Range.Height == 4 &&
                    Convert.ToString(grid[0, 2].Value) == "A", "undo table structure");
                Call(form, "Redo");
                Check(((List<TableDefinition>)Field(form, "tables"))[0].Range.Height == 3 &&
                    Convert.ToString(grid[0, 1].Value) == "A", "redo table structure");
                grid.CurrentCell = grid[0, 199];
                Call(form, "InsertRow");
                Check(grid.RowCount > 200, "insert at initial grid edge");
                form.GetType().GetField("dirty", BindingFlags.Instance | BindingFlags.NonPublic)
                    .SetValue(form, false);
                form.Close();
            }
        }
    }
}
