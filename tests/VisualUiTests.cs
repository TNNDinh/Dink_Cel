using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace DinkCel
{
    internal static class VisualUiTests
    {
        private static object Field(object target, string name)
        {
            return target.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(target);
        }
        private static object Call(object target, string name, params object[] args)
        {
            return target.GetType().GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance)
                .Invoke(target, args);
        }
        private static void Check(bool condition, string message)
        {
            if (!condition) throw new Exception(message);
        }

        [STAThread]
        private static int Main()
        {
            try
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                EmbeddedDependencies.Install();
                Check(ThemePalette.All.Length == 6, "Six themes should be available");
                Check(ThemePalette.Find("night").Id == "dark", "Legacy theme should load");
                var book = new WorkbookSnapshot();
                book.Sheets.Add(new SheetSnapshot { Name = "Archive", Hidden = true, TabColor = Color.Coral });
                foreach (string extension in new[] { ".xlsx", ".xls", ".ods" })
                {
                    string interchange = Path.Combine(Path.GetTempPath(),
                        "DinkCel-visual-" + Guid.NewGuid().ToString("N") + extension);
                    try
                    {
                        WorkbookSnapshot loaded;
                        if (extension == ".xlsx")
                        { XlsxFile.Write(interchange, book, 200, 26); loaded = XlsxFile.Read(interchange, 200, 26); }
                        else if (extension == ".xls")
                        { XlsFile.Write(interchange, book, 200, 26); loaded = XlsFile.Read(interchange, 200, 26); }
                        else
                        { OdsFile.Write(interchange, book, 200, 26); loaded = OdsFile.Read(interchange, 200, 26); }
                        Check(loaded.Sheets.Count == 2 && loaded.Sheets[1].Hidden,
                            extension + " should preserve hidden sheets");
                        if (extension == ".xlsx")
                            Check(loaded.Sheets[1].TabColor.ToArgb() == Color.Coral.ToArgb(),
                                "XLSX should preserve tab color");
                    }
                    finally { if (File.Exists(interchange)) File.Delete(interchange); }
                }
                string path = Path.Combine(Path.GetTempPath(), "DinkCel-visual-" + Guid.NewGuid().ToString("N") + ".dinkcel");
                try
                {
                    using (var form = new SpreadsheetForm(null))
                    {
                        form.Show(); Application.DoEvents();
                        Check(((Panel)Field(form, "welcomePanel")).Visible, "Welcome should open for a new session");
                        Call(form, "NewDocument");
                        Check(!((Panel)Field(form, "welcomePanel")).Visible, "New workbook should close welcome");
                        Call(form, "ToggleInspector");
                        Application.DoEvents();
                        var inspector = (Panel)Field(form, "inspector");
                        var visibleGrid = (DataGridView)Field(form, "grid");
                        Check(inspector.Visible, "Inspector should open");
                        if (inspector.Parent.ClientSize.Width >= 1050)
                            Check(visibleGrid.Right <= inspector.Left, "Inspector should not cover the grid on wide screens");
                        Call(form, "ToggleInspector");
                        Call(form, "SelectRectangle", new Rectangle(0, 0, 2, 2), 0, 0, false);
                        Call(form, "UpdateFloatingActions");
                        Check(((FlowLayoutPanel)Field(form, "floatingActions")).Visible,
                            "Mouse range should show compact format actions");
                        visibleGrid.Focus();
                        Call(form, "HandleEditingShortcut", Keys.Down);
                        Check(!((FlowLayoutPanel)Field(form, "floatingActions")).Visible,
                            "Keyboard navigation should hide compact actions");
                        Call(form, "ApplyZoom", 125);
                        var grid = (DataGridView)Field(form, "grid");
                        Check(grid.Columns[0].Width == 150, "Zoom should resize the visible grid");
                        grid[0, 0].Value = "42";
                        Call(form, "DuplicateSheet", 0);
                        Call(form, "MoveSheet", 1, 0);
                        Call(form, "HideSheet", 1);
                        ((System.Collections.Generic.List<SheetState>)Field(form, "sheets"))[0].TabColor = Color.Coral;
                        Check(((FlowLayoutPanel)Field(form, "sheetTabs")).Controls.Count == 2,
                            "Hidden sheet should disappear from tabs");
                        Check((bool)Call(form, "WriteWorkbook", path), "Workbook save should succeed");
                        Check(File.ReadAllText(path).Contains("tabColor"), "DinkCel should save tab color");
                        form.Close();
                    }
                    using (var reopened = new SpreadsheetForm(path))
                    {
                        reopened.Show(); Application.DoEvents();
                        var tabs = (FlowLayoutPanel)Field(reopened, "sheetTabs");
                        Check(tabs.Controls.Count == 2, "Hidden sheet should stay hidden after save");
                        var grid = (DataGridView)Field(reopened, "grid");
                        Check(grid.Columns[0].Width == 120, "Zoom must not change saved column width");
                        Check(Convert.ToString(grid[0, 0].Value) == "42", "Duplicated sheet data should survive save");
                        reopened.Close();
                    }
                }
                finally { if (File.Exists(path)) File.Delete(path); }
                Console.WriteLine("Visual UI: welcome, themes, zoom and sheet tabs passed.");
                return 0;
            }
            catch (Exception error)
            {
                Console.Error.WriteLine(error);
                return 1;
            }
        }
    }
}
