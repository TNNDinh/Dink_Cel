using System;
using System.IO;
using System.Reflection;
using System.Drawing;
using System.Collections.Generic;
using System.Windows.Forms;

namespace DinkCel
{
    internal static class MultiSheetTests
    {
        private static object Field(object target, string name)
        {
            return target.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(target);
        }
        private static object Call(object target, string name, params object[] args)
        {
            return target.GetType().GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance).Invoke(target, args);
        }
        private static void SetField(object target, string name, object value)
        {
            target.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).SetValue(target, value);
        }
        private static void Equal(object expected, object actual)
        {
            if (!object.Equals(expected, actual)) throw new Exception("Expected " + expected + ", got " + actual);
        }

        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            string path = Path.Combine(Path.GetTempPath(), "DinkCel_sheets_" + Guid.NewGuid().ToString("N") + ".dinkcel");
            string xlsx = Path.ChangeExtension(path, ".xlsx");
            try
            {
                using (var form = new SpreadsheetForm(null))
                {
                    form.Show(); Application.DoEvents();
                    var grid = (DataGridView)Field(form, "grid");
                    grid[0, 0].Value = "First";
                    grid[0, 1].Value = "Hidden";
                    grid[0, 1].Style.Format = "0.00";
                    grid.ClearSelection();
                    grid[0, 0].Selected = true;
                    grid[1, 0].Selected = true;
                    Call(form, "MergeSelection");
                    Equal(1, ((List<Rectangle>)Field(form, "merges")).Count);
                    using (var image = new Bitmap(400, 120)) grid.DrawToBitmap(image, new Rectangle(0, 0, 400, 120));
                    Call(form, "UnmergeSelection");
                    Equal(0, ((List<Rectangle>)Field(form, "merges")).Count);
                    ((List<ConditionalRule>)Field(form, "conditionalRules")).Add(new ConditionalRule
                    { Range = new Rectangle(0, 1, 1, 1), Threshold = 2, Color = Color.LightGreen });
                    SetField(form, "filterColumn", 0);
                    SetField(form, "filterValue", "First");
                    Call(form, "ApplyFreezeAndFilter");
                    Equal(false, grid.Rows[1].Visible);
                    SetField(form, "filterColumn", -1);
                    Call(form, "ApplyFreezeAndFilter");
                    Equal(true, grid.Rows[1].Visible);
                    grid.CurrentCell = grid[1, 1];
                    Call(form, "FreezeAtCell");
                    Equal(true, grid.Rows[0].Frozen);
                    Equal(true, grid.Columns[0].Frozen);
                    grid.CurrentCell = grid[0, 0];
                    Call(form, "AddSheet");
                    Equal(true, Field(form, "dirty"));
                    grid[0, 0].Value = "Second";
                    grid[1, 0].Value = "=LEN(Sheet1!A1)";
                    Equal("5", grid[1, 0].FormattedValue);
                    Call(form, "SwitchSheet", 0);
                    Equal("First", grid[0, 0].Value);
                    Equal("0.00", grid[0, 1].Style.Format);
                    Equal(1, ((List<ConditionalRule>)Field(form, "conditionalRules")).Count);
                    Equal(true, Field(form, "dirty"));
                    Call(form, "SwitchSheet", 1);
                    Equal("Second", grid[0, 0].Value);
                    Equal("5", grid[1, 0].FormattedValue);
                    Call(form, "Undo");
                    Equal(null, grid[1, 0].Value);
                    Call(form, "Redo");
                    Equal("=LEN(Sheet1!A1)", grid[1, 0].Value);
                    Equal(true, Call(form, "WriteWorkbook", path));
                    Equal(false, Field(form, "dirty"));
                    Equal(true, Call(form, "WriteXlsx", xlsx));
                    form.Close();
                }
                using (var reopened = new SpreadsheetForm(path))
                {
                    reopened.Show(); Application.DoEvents();
                    var grid = (DataGridView)Field(reopened, "grid");
                    Equal("First", grid[0, 0].Value);
                    Equal("0.00", grid[0, 1].Style.Format);
                    Equal(1, ((List<ConditionalRule>)Field(reopened, "conditionalRules")).Count);
                    Call(reopened, "SwitchSheet", 1);
                    Equal("Second", grid[0, 0].Value);
                    Equal("5", grid[1, 0].FormattedValue);
                    reopened.Close();
                }
                using (var reopened = new SpreadsheetForm(xlsx))
                {
                    reopened.Show(); Application.DoEvents();
                    var grid = (DataGridView)Field(reopened, "grid");
                    Equal("First", grid[0, 0].Value);
                    Equal("0.00", grid[0, 1].Style.Format);
                    Equal(1, ((List<ConditionalRule>)Field(reopened, "conditionalRules")).Count);
                    Call(reopened, "SwitchSheet", 1);
                    Equal("Second", grid[0, 0].Value);
                    Equal("5", grid[1, 0].FormattedValue);
                    reopened.Close();
                }
                Console.WriteLine("Multi-sheet UI: .dinkcel and .xlsx passed.");
            }
            finally { if (File.Exists(path)) File.Delete(path); if (File.Exists(xlsx)) File.Delete(xlsx); }
        }
    }
}
