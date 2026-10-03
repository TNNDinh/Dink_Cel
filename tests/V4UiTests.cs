using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Windows.Forms;

namespace DinkCel
{
    internal static class V4UiTests
    {
        private static object Field(object target, string name)
        { return target.GetType().GetField(name, BindingFlags.NonPublic | BindingFlags.Instance).GetValue(target); }
        private static object Call(object target, string name, params object[] args)
        { return target.GetType().GetMethod(name, BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.Static).Invoke(target, args); }
        private static void Equal(object expected, object actual)
        { if (!object.Equals(expected, actual)) throw new Exception("Expected " + expected + ", got " + actual); }
        private static void Check(bool condition, string message)
        { if (!condition) throw new Exception(message); }

        [STAThread]
        private static void Main()
        {
            try { Run(); }
            catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }
        }

        private static void Run()
        {
            EmbeddedDependencies.Install();
            Application.EnableVisualStyles();
            using (var form = new SpreadsheetForm(null))
            {
                form.Show(); Application.DoEvents();
                var grid = (DataGridView)Field(form, "grid");
                grid.Focus();
                grid[0, 0].Value = "1"; grid[0, 1].Value = "2";
                grid[0, 3].Value = "9";
                Call(form, "SelectRectangle", new Rectangle(0, 0, 1, 1), 0, 0, false);
                Call(form, "HandleEditingShortcut", Keys.Control | Keys.Down);
                Equal(1, grid.CurrentCell.RowIndex);
                Call(form, "HandleEditingShortcut", Keys.Control | Keys.Down);
                Equal(3, grid.CurrentCell.RowIndex);
                Call(form, "HandleEditingShortcut", Keys.Control | Keys.Home);
                Equal(0, grid.CurrentCell.RowIndex);
                Call(form, "HandleEditingShortcut", Keys.Control | Keys.Shift | Keys.Down);
                Equal(2, grid.SelectedCells.Count);
                Call(form, "HandleEditingShortcut", Keys.Shift | Keys.Right);
                Equal(4, grid.SelectedCells.Count);
                Call(form, "HandleEditingShortcut", Keys.Control | Keys.End);
                Equal(3, grid.CurrentCell.RowIndex);
                Call(form, "HandleEditingShortcut", Keys.Control | Keys.Space);
                Equal(200, grid.SelectedCells.Count);
                Call(form, "HandleEditingShortcut", Keys.Shift | Keys.Space);
                Equal(26, grid.SelectedCells.Count);
                Call(form, "HandleEditingShortcut", Keys.Control | Keys.A);
                Equal(5200, grid.SelectedCells.Count);
                Call(form, "SelectRectangle", new Rectangle(0, 0, 1, 1), 0, 0, false);
                Equal(true, Call(grid, "ProcessDataGridViewKey", new KeyEventArgs(Keys.Control | Keys.Down)));
                Equal(1, grid.CurrentCell.RowIndex);
                Equal(true, Call(grid, "ProcessDialogKey", Keys.Tab));
                Equal(1, grid.CurrentCell.ColumnIndex);
                Console.WriteLine("v0.4 navigation passed.");

                Equal(true, Call(form, "GoToAddress", "B2:D3,F5"));
                Equal(7, grid.SelectedCells.Count);
                Equal(5, grid.CurrentCell.ColumnIndex);
                Equal(4, grid.CurrentCell.RowIndex);
                Equal(false, Call(form, "GoToAddress", "AA201"));
                Console.WriteLine("v0.4 name box passed.");

                Call(form, "SelectRectangle", new Rectangle(1, 0, 1, 1), 1, 0, false);
                var formulaBar = (TextBox)Field(form, "contentBox");
                formulaBar.Text = "=A1+1";
                Equal(true, Call(form, "CommitFormulaBar"));
                Equal("=A1+1", grid[1, 0].Value);
                Equal("2", grid[1, 0].FormattedValue);
                Console.WriteLine("v0.4 formula bar passed.");
                Call(form, "SelectRectangle", new Rectangle(1, 0, 1, 1), 1, 0, false);
                Call(form, "CopySelected");
                Call(form, "SelectRectangle", new Rectangle(3, 2, 1, 1), 3, 2, false);
                Call(form, "PasteSelected");
                Equal("=C3+1", grid[3, 2].Value);
                Call(form, "Undo");
                Equal(null, grid[3, 2].Value);
                Call(form, "PasteValues");
                Equal("2", grid[3, 2].Value);
                Console.WriteLine("v0.4 clipboard paste passed.");
                grid[1, 0].Style.BackColor = Color.LightBlue;
                Call(form, "SelectRectangle", new Rectangle(1, 0, 1, 1), 1, 0, false);
                Call(form, "CopySelected");
                grid[9, 0].Value = "old";
                Call(form, "SelectRectangle", new Rectangle(9, 0, 1, 1), 9, 0, false);
                Call(form, "PasteFormats");
                Equal("old", grid[9, 0].Value);
                Equal(Color.LightBlue.ToArgb(), grid[9, 0].Style.BackColor.ToArgb());
                Call(form, "PasteFormulas");
                Equal("=I1+1", grid[9, 0].Value);
                Equal(Color.LightBlue.ToArgb(), grid[9, 0].Style.BackColor.ToArgb());
                grid[10, 0].Value = "X"; grid[11, 0].Value = "Y";
                Call(form, "SelectRectangle", new Rectangle(10, 0, 2, 1), 10, 0, false);
                Call(form, "CopySelected");
                Call(form, "SelectRectangle", new Rectangle(12, 0, 1, 1), 12, 0, false);
                Call(form, "PasteTranspose");
                Equal("X", grid[12, 0].Value);
                Equal("Y", grid[12, 1].Value);
                Clipboard.SetText("\"a\tb\"\t42\r\nleft\tright");
                Call(form, "SelectRectangle", new Rectangle(13, 0, 1, 1), 13, 0, false);
                Call(form, "PasteSelected");
                Equal("a\tb", grid[13, 0].Value);
                Equal("right", grid[14, 1].Value);
                grid[2, 0].Value = "3";
                Equal(true, Call(form, "GoToAddress", "A1,C1"));
                Equal(2, grid.SelectedCells.Count);
                Call(form, "CopySelected");
                Call(form, "SelectRectangle", new Rectangle(0, 9, 1, 1), 0, 9, false);
                Call(form, "PasteSelected");
                Equal("1", grid[0, 9].Value);
                Equal(null, grid[1, 9].Value);
                Equal("3", grid[2, 9].Value);
                grid[2, 0].Value = null;
                grid[21, 0].Value = "old";
                Call(form, "SelectRectangle", new Rectangle(20, 0, 1, 1), 20, 0, false);
                Call(form, "CopySelected");
                Call(form, "SelectRectangle", new Rectangle(21, 0, 1, 1), 21, 0, false);
                Call(form, "PasteSelected");
                Check(string.IsNullOrEmpty(Convert.ToString(grid[21, 0].Value)),
                    "Pasting a blank cell should clear the destination");

                grid[4, 0].Value = "=A1";
                Call(form, "SelectRectangle", new Rectangle(4, 0, 1, 3), 4, 0, false);
                Call(form, "FillDown");
                Equal("=A2", grid[4, 1].Value);
                Equal("=A3", grid[4, 2].Value);
                grid[15, 0].Value = "=A1";
                Call(form, "SelectRectangle", new Rectangle(15, 0, 3, 1), 15, 0, false);
                Call(form, "FillRight");
                Equal("=B1", grid[16, 0].Value);
                Equal("=C1", grid[17, 0].Value);
                Call(form, "AutoFillSelection", new Rectangle(0, 0, 1, 2), 5, 0);
                Equal("3", grid[0, 2].Value);
                Equal("6", grid[0, 5].Value);
                grid[5, 0].Value = "2026-10-03";
                Call(form, "AutoFillSelection", new Rectangle(5, 0, 1, 1), 2, 5);
                Equal("2026-10-04", grid[5, 1].Value);
                Equal("2026-10-05", grid[5, 2].Value);
                Console.WriteLine("v0.4 fill handle passed.");

                grid[7, 0].Value = "1"; grid[7, 1].Value = "3";
                Call(form, "SelectRectangle", new Rectangle(7, 0, 1, 4), 7, 0, false);
                Call(form, "FillSeries");
                Equal("3", grid[7, 1].Value);
                Equal("5", grid[7, 2].Value);
                Equal("7", grid[7, 3].Value);
                Console.WriteLine("v0.4 series passed.");

                grid.Focus();
                Call(form, "SelectRectangle", new Rectangle(18, 0, 1, 1), 18, 0, false);
                Equal(true, Call(form, "HandleEditingShortcut", Keys.F2));
                Check(grid.IsCurrentCellInEditMode, "F2 should edit the current cell");
                ((TextBox)grid.EditingControl).Text = "typed";
                Equal(true, Call(form, "HandleEditingShortcut", Keys.Enter));
                Equal("typed", grid[18, 0].Value);
                Equal(1, grid.CurrentCell.RowIndex);
                Call(form, "HandleEditingShortcut", Keys.Shift | Keys.Enter);
                Equal(0, grid.CurrentCell.RowIndex);
                Call(form, "HandleEditingShortcut", Keys.Tab);
                Equal(19, grid.CurrentCell.ColumnIndex);
                Call(form, "HandleEditingShortcut", Keys.Shift | Keys.Tab);
                Equal(18, grid.CurrentCell.ColumnIndex);
                Call(form, "HandleEditingShortcut", Keys.F2);
                ((TextBox)grid.EditingControl).Text = "cancelled";
                Call(form, "HandleEditingShortcut", Keys.Escape);
                Equal("typed", grid[18, 0].Value);
                Console.WriteLine("v0.4 edit keys passed.");

                Call(form, "SelectRectangle", new Rectangle(0, 0, 1, 1), 0, 0, false);
                Call(form, "CutSelected");
                Call(form, "SelectRectangle", new Rectangle(2, 0, 1, 1), 2, 0, false);
                Call(form, "PasteSelected");
                Equal(null, grid[0, 0].Value);
                Equal("1", grid[2, 0].Value);
                Call(form, "Undo");
                Equal("1", grid[0, 0].Value);
                Equal(null, grid[2, 0].Value);

                Call(form, "SelectRectangle", new Rectangle(0, 0, 1, 1), 0, 0, false);
                Call(form, "CutSelected");
                Call(form, "AddSheet");
                Call(form, "PasteSelected");
                Equal("1", grid[0, 0].Value);
                Call(form, "Undo");
                Equal(null, grid[0, 0].Value);
                Call(form, "SwitchSheet", 0);
                Equal("1", grid[0, 0].Value);
                Call(form, "SwitchSheet", 1);
                Call(form, "Redo");
                Equal("1", grid[0, 0].Value);
                Call(form, "SwitchSheet", 0);
                Equal(null, grid[0, 0].Value);
                grid[0, 0].Value = "99";
                Call(form, "SwitchSheet", 1);
                Call(form, "Undo");
                Call(form, "SwitchSheet", 0);
                Equal("99", grid[0, 0].Value);
                Call(form, "SwitchSheet", 1);
                Call(form, "Redo");
                Call(form, "SwitchSheet", 0);
                Equal("99", grid[0, 0].Value);
                Call(form, "SelectRectangle", new Rectangle(1, 0, 1, 1), 1, 0, false);
                formulaBar.Text = "saved from formula bar";
                Call(form, "SaveActiveSheet");
                Equal("saved from formula bar", grid[1, 0].Value);
                ((List<NamedRange>)Field(form, "namedRanges")).Add(new NamedRange
                { Name = "MainValues", Sheet = "Sheet1", Range = new Rectangle(0, 1, 1, 2) });
                Call(form, "SwitchSheet", 1);
                Equal(true, Call(form, "GoToAddress", "MainValues"));
                Equal(2, grid.SelectedCells.Count);
                Equal(1, grid.CurrentCell.RowIndex);

                Call(form, "SwitchSheet", 0);
                grid[0, 0].Value = "10";
                grid[1, 0].Value = "=A1*2";
                grid[2, 0].Value = "=B1+1";
                grid[3, 0].Value = "=XLOOKUP(10,A1:A1,C1:C1)";
                Equal("21", grid[2, 0].FormattedValue);
                Equal("21", grid[3, 0].FormattedValue);
                var retainedEngine = (FormulaEngine)Field(form, "formulaEngine");
                grid[0, 0].Value = "12";
                Equal(true, object.ReferenceEquals(retainedEngine, Field(form, "formulaEngine")));
                Equal("25", grid[2, 0].FormattedValue);
                Equal("#N/A", grid[3, 0].FormattedValue);
                Call(form, "SwitchSheet", 1);
                grid[1, 0].Value = "=Sheet1!C1";
                Equal("25", grid[1, 0].FormattedValue);

                Call(form, "SwitchSheet", 0);
                Call(form, "SelectRectangle", new Rectangle(0, 0, 1, 1), 0, 0, false);
                Call(form, "ApplyFontFamily", "Consolas");
                Call(form, "ToggleFontStyle", FontStyle.Strikeout);
                Call(form, "SetVerticalAlignment", 0);
                Call(form, "ToggleWrap");
                Call(form, "SetBorders", "Tất cả");
                Call(form, "StartFormatPainter");
                Call(form, "SelectRectangle", new Rectangle(4, 0, 1, 1), 4, 0, false);
                Call(form, "ApplyFormatPainter");
                Equal("Consolas", grid[4, 0].Style.Font.Name);
                Equal("thin", ((CellExtras)grid[4, 0].Tag).Left.Style);
                grid[4, 0].Value = "1234.5";
                Call(form, "ApplyNumberFormat", "Percentage");
                Equal("123450.00%", grid[4, 0].FormattedValue);
                Call(form, "ClearFormats");
                Equal(null, grid[4, 0].Tag);
                Equal("1234.5", grid[4, 0].Value);
                Call(form, "ClearAll");
                Equal(null, grid[4, 0].Value);
                Call(form, "SelectRectangle", new Rectangle(0, 0, 1, 1), 0, 0, false);
                Call(form, "SetHidden", true, true);
                Check(!grid.Rows[0].Visible, "Row should hide immediately");
                Call(form, "SelectRectangle", new Rectangle(1, 1, 1, 1), 1, 1, false);
                Call(form, "SetHidden", false, true);
                Check(!grid.Columns[1].Visible, "Column should hide immediately");
                string formattingPath = Path.Combine(Path.GetTempPath(),
                    "DinkCel_format_" + Guid.NewGuid().ToString("N") + ".dinkcel");
                try
                {
                    Equal(true, Call(form, "WriteWorkbook", formattingPath));
                    Check(File.ReadAllText(formattingPath).Contains("hiddenRow"), "Hidden row must be serialized");
                    using (var reopened = new SpreadsheetForm(formattingPath))
                    {
                        reopened.Show(); Application.DoEvents();
                        var savedGrid = (DataGridView)Field(reopened, "grid");
                        Check(!savedGrid.Rows[0].Visible, "Hidden row should survive .dinkcel reload");
                        Check(!savedGrid.Columns[1].Visible, "Hidden column should survive .dinkcel reload");
                        Equal("Consolas", savedGrid[0, 0].Style.Font.Name);
                        Equal(true, ((CellExtras)savedGrid[0, 0].Tag).Wrap);
                        Equal("thin", ((CellExtras)savedGrid[0, 0].Tag).Left.Style);
                    }
                }
                finally { if (File.Exists(formattingPath)) File.Delete(formattingPath); }

                var htmlData = new DataObject();
                htmlData.SetData(DataFormats.Html, "<html><head><style>.xl65{font-family:Consolas;font-weight:bold;" +
                    "color:#112233;background-color:#DDEEFF;text-align:right;border-left:1px solid #FF0000;" +
                    "mso-number-format:0.00%}</style></head><body><table><tr>" +
                    "<td class=xl65>0.25</td><td style='font-style:italic'>Note</td>" +
                    "</tr></table></body></html>");
                htmlData.SetData(DataFormats.UnicodeText, "0.25\tNote");
                Clipboard.SetDataObject(htmlData, true);
                Call(form, "SelectRectangle", new Rectangle(7, 1, 1, 1), 7, 1, false);
                Call(form, "PasteSelected");
                Equal("0.25", grid[7, 1].Value);
                Equal("25.00%", grid[7, 1].FormattedValue);
                Equal("Consolas", grid[7, 1].Style.Font.Name);
                Equal(Color.FromArgb(221, 238, 255).ToArgb(), grid[7, 1].Style.BackColor.ToArgb());
                Equal("thin", ((CellExtras)grid[7, 1].Tag).Left.Style);
                Equal("Note", grid[8, 1].Value);
                Equal(true, (grid[8, 1].Style.Font.Style & FontStyle.Italic) != 0);
                Equal("1 1/4", Call(form, "FormatNumeric", 1.25, "# ?/?"));
                Equal("12.50%", Call(form, "FormatNumeric", 0.125, "0.00%"));
                Equal("12.5", Call(form, "FormatNumeric", 12.5, "General"));
                Equal("#,##0.00", Call(form, "NormalizeNumberFormat", "N2"));
                Equal("0.0%", Call(form, "NormalizeNumberFormat", "P1"));
                Check(Convert.ToString(Call(form, "FormatNumeric", 1234.5, "#,##0.00 ₫")).Contains("₫"),
                    "Currency should display its symbol");
                Check(Convert.ToString(Call(form, "FormatNumeric", -12.5, "#,##0.00;(#,##0.00);–")).Contains("("),
                    "Accounting should mark negative values");
                Equal("12:00:00", Call(form, "FormatNumeric", 0.5, "HH:mm:ss"));
                Check(Convert.ToString(Call(form, "FormatNumeric", 1234.5, "0.00E+00")).Contains("E+"),
                    "Scientific should use an exponent");
                Equal(new DateTime(2026, 10, 3).ToString("dd/MM/yyyy"),
                    Call(form, "FormatNumeric", new DateTime(2026, 10, 3).ToOADate(), "dd/MM/yyyy"));
                grid[9, 1].Value = "Long text for automatic column and row sizing";
                Call(form, "SelectRectangle", new Rectangle(9, 1, 1, 1), 9, 1, false);
                grid.Columns[9].Width = 35;
                Call(form, "AutoFitColumn");
                Check(grid.Columns[9].Width > 35, "AutoFit should widen the selected column");
                grid.Columns[9].Width = 50;
                grid[9, 1].Style.WrapMode = DataGridViewTriState.True;
                Call(form, "AutoFitRow");
                Check(grid.Rows[1].Height > 27, "AutoFit should grow a wrapped row");
                Console.WriteLine("v0.6 formatting: UI, clipboard HTML, .dinkcel and number formats passed.");

            }
            Console.WriteLine("v0.4 UI: navigation, ranges, formula bar, fill, clipboard, undo passed.");
        }
    }
}
