using System;
using System.Drawing;
using System.IO;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;

namespace DinkCel
{
    internal static class DataToolsTests
    {
        private static void Check(bool condition, string message)
        {
            if (!condition) throw new Exception(message);
        }

        [STAThread]
        private static void Main()
        {
            try { Run(); Console.WriteLine("v0.7.1 data tools: matching, validation and XLSX round trip passed."); }
            catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }
        }

        private static void Run()
        {
            Check(DataTools.Compare("10", "2", "Number") > 0, "numeric sort");
            Check(DataTools.Compare("2024-01-02", "2024-01-10", "Date") < 0, "date sort");
            Check(DataTools.Match("Report Q4", "Text", "Contains", "q4", ""), "contains");
            Check(DataTools.Match("2024-03-10", "Date", "Between", "2024-03-01", "2024-03-31"), "date between");
            Check(!DataTools.Match("abc", "Number", "Greater", "10", ""), "reject nonnumeric filter");
            Check(DataTools.Match("", "Text", "Blank", "", ""), "blank");
            Check(DataTools.FindMatch("Quarter 2024", "Q*2024", false, true, true), "wildcard");
            Check(DataTools.Replace("Cat cat", "cat", "dog", false, false, false) == "dog dog", "replace all occurrences");
            var number = new ValidationRule { Kind = "Whole Number", Operator = "Between", Value1 = "1", Value2 = "10" };
            Check(DataTools.Valid(number, "5") && !DataTools.Valid(number, "5.5") && !DataTools.Valid(number, "11"), "whole number validation");
            var length = new ValidationRule { Kind = "Text Length", Operator = "Greater Or Equal", Value1 = "3" };
            Check(DataTools.Valid(length, "abcd") && !DataTools.Valid(length, "ab"), "length validation");
            var book = new WorkbookSnapshot();
            var sheet = book.Sheets[0];
            sheet.Cells[0] = new CellSnapshot { Text = "Amount" };
            sheet.Filters.Add(new FilterCriterion { Column = 0, Kind = "Number", Operator = "Between", Value1 = "10", Value2 = "20" });
            sheet.Filters.Add(new FilterCriterion { Column = 1, Kind = "Text", Operator = "Contains", Value1 = "ok" });
            sheet.Validations.Add(new ValidationRule { Range = new Rectangle(0, 1, 1, 10), Kind = "Decimal",
                Operator = "Greater", Value1 = "0", AllowBlank = false, InputTitle = "Amount",
                InputMessage = "Positive", ErrorTitle = "Invalid", ErrorMessage = "Enter positive amount" });
            sheet.Validations.Add(new ValidationRule { Range = new Rectangle(1, 1, 1, 10), Kind = "Date",
                Operator = "Between", Value1 = "2024-01-01", Value2 = "2024-12-31" });
            sheet.Rules.Add(new ConditionalRule { Range = new Rectangle(0, 1, 1, 10), Kind = "Between",
                Value1 = "10", Value2 = "20", Color = Color.LightGreen });
            sheet.Rules.Add(new ConditionalRule { Range = new Rectangle(0, 1, 1, 10), Kind = "Color Scale",
                Color = Color.Red, Color2 = Color.White });
            sheet.Rules.Add(new ConditionalRule { Range = new Rectangle(0, 1, 1, 10), Kind = "Text Contains",
                Value1 = "urgent", Color = Color.Yellow });
            sheet.Rules.Add(new ConditionalRule { Range = new Rectangle(0, 1, 1, 10), Kind = "Data Bar", Color = Color.Blue });
            sheet.Rules.Add(new ConditionalRule { Range = new Rectangle(0, 1, 1, 10), Kind = "Icon Set", Color = Color.Green });
            string path = Path.Combine(Path.GetTempPath(), "DinkCel_data_" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                XlsxFile.Write(path, book, 200, 26);
                var loaded = XlsxFile.Read(path, 200, 26).Sheets[0];
                Check(loaded.Filters.Count == 2 && loaded.Filters[0].Operator == "Between", "filter round trip");
                Check(loaded.Validations.Count == 2 && loaded.Validations[0].InputMessage == "Positive" &&
                    loaded.Validations[0].Kind == "Decimal" && loaded.Validations[1].Kind == "Date" &&
                    DataTools.Valid(loaded.Validations[1], "2024-06-15"), "validation round trip");
                Check(loaded.Rules.Count == 5 && loaded.Rules[0].Kind == "Between" &&
                    loaded.Rules[1].Kind == "Color Scale" && loaded.Rules[2].Kind == "Text Contains" &&
                    loaded.Rules[3].Kind == "Data Bar" && loaded.Rules[4].Kind == "Icon Set", "conditional formatting round trip");
            }
            finally { if (File.Exists(path)) File.Delete(path); }

            Application.EnableVisualStyles();
            string dinkcel = Path.ChangeExtension(path, ".dinkcel");
            try
            {
                using (var form = new SpreadsheetForm(null))
                {
                    form.Show(); Application.DoEvents();
                    var grid = (DataGridView)typeof(SpreadsheetForm).GetField("grid", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(form);
                    grid[0, 0].Value = "Amount";
                    grid[0, 1].Value = "10"; grid[0, 1].Tag = new CellExtras { Indent = 2 };
                    grid[0, 2].Value = "2";
                    grid.CurrentCell = grid[0, 1]; grid.ClearSelection(); grid[0, 1].Selected = true;
                    typeof(SpreadsheetForm).GetMethod("SortDataRows", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(form,
                        new object[] { new List<Tuple<int, bool, string>> { Tuple.Create(0, false, "Number") } });
                    Check((string)grid[0, 0].Value == "Amount" && (string)grid[0, 1].Value == "2" &&
                        (string)grid[0, 2].Value == "10", "sort with header");
                    Check(((CellExtras)grid[0, 2].Tag).Indent == 2, "sort preserves extras");
                    var filters = (List<FilterCriterion>)typeof(SpreadsheetForm).GetField("activeFilters", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(form);
                    filters.Add(new FilterCriterion { Column = 0, Kind = "Number", Operator = "Greater", Value1 = "5" });
                    typeof(SpreadsheetForm).GetMethod("ApplyFreezeAndFilter", BindingFlags.Instance | BindingFlags.NonPublic).Invoke(form, null);
                    Check(!grid.Rows[1].Visible && grid.Rows[2].Visible, "typed filter applies");
                    var rules = (List<ConditionalRule>)typeof(SpreadsheetForm).GetField("conditionalRules", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(form);
                    rules.Add(new ConditionalRule { Range = new Rectangle(0, 1, 1, 2), Kind = "Duplicate", Color = Color.Orange });
                    var validations = (List<ValidationRule>)typeof(SpreadsheetForm).GetField("validations", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(form);
                    validations.Add(new ValidationRule { Range = new Rectangle(0, 1, 1, 2), Kind = "Whole Number", Value1 = "0", Value2 = "100" });
                    validations.Add(new ValidationRule { Range = new Rectangle(1, 1, 1, 2), Kind = "Custom Formula", Value1 = "=B2>5" });
                    var accept = typeof(SpreadsheetForm).GetMethod("CanAcceptValue", BindingFlags.Instance | BindingFlags.NonPublic);
                    Check((bool)accept.Invoke(form, new object[] { 1, 1, "6" }) &&
                        !(bool)accept.Invoke(form, new object[] { 2, 1, "4" }), "relative custom validation formula");
                    Check((bool)typeof(SpreadsheetForm).GetMethod("WriteWorkbook", BindingFlags.Instance | BindingFlags.NonPublic)
                        .Invoke(form, new object[] { dinkcel }), "write dinkcel");
                    form.Close();
                }
                using (var reopened = new SpreadsheetForm(dinkcel))
                {
                    reopened.Show(); Application.DoEvents();
                    var filters = (List<FilterCriterion>)typeof(SpreadsheetForm).GetField("activeFilters", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(reopened);
                    var rules = (List<ConditionalRule>)typeof(SpreadsheetForm).GetField("conditionalRules", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(reopened);
                    var validations = (List<ValidationRule>)typeof(SpreadsheetForm).GetField("validations", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(reopened);
                    Check(filters.Count == 1 && filters[0].Value1 == "5", "dinkcel filter round trip");
                    Check(rules.Count == 1 && rules[0].Kind == "Duplicate", "dinkcel conditional round trip");
                    Check(validations.Count == 2 && validations[0].Kind == "Whole Number" &&
                        validations[1].Kind == "Custom Formula", "dinkcel validation round trip");
                    reopened.Close();
                }
            }
            finally { if (File.Exists(dinkcel)) File.Delete(dinkcel); }
        }
    }
}
