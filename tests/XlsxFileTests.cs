using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace DinkCel
{
    internal static class XlsxFileTests
    {
        private static void Equal(object expected, object actual)
        {
            if (!object.Equals(expected, actual)) throw new Exception("Expected " + expected + ", got " + actual);
        }

        private static void Main()
        {
            try { Run(); }
            catch (Exception error) { Console.Error.WriteLine(error); Environment.Exit(1); }
        }

        private static void Run()
        {
            string path = Path.Combine(Path.GetTempPath(), "DinkCel_xlsx_" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                var book = new WorkbookSnapshot();
                book.Sheets[0].Name = "Sales";
                book.Sheets[0].Cells[0] = new CellSnapshot { Text = "Revenue" };
                book.Sheets[0].Cells[26] = new CellSnapshot { Text = "1234.5", NumberFormat = "#,##0.00",
                    HasFont = true, FontStyle = FontStyle.Bold, FontSize = 14F,
                    ForeColor = Color.DarkBlue, BackColor = Color.LightYellow,
                    Alignment = DataGridViewContentAlignment.MiddleRight };
                book.Sheets[0].Cells[27] = new CellSnapshot { Text = "=A2*2" };
                book.Sheets[0].Merges.Add(new Rectangle(0, 0, 2, 1));
                book.Sheets[0].FreezeRow = 1;
                book.Sheets[0].FilterColumn = 0;
                book.Sheets[0].FilterValue = "12";
                book.Sheets[0].Rules.Add(new ConditionalRule { Range = new Rectangle(0, 1, 1, 1), Threshold = 1000, Color = Color.LightGreen });
                book.Sheets[0].ColumnWidths[0] = 180;
                book.Sheets[0].RowHeights[0] = 35;
                var second = new SheetSnapshot { Name = "Ghi chú" };
                second.Cells[0] = new CellSnapshot { Text = "Xin chào" };
                second.Cells[1] = new CellSnapshot { Text = "Value" };
                second.Cells[26] = new CellSnapshot { Text = "Item" };
                second.Cells[27] = new CellSnapshot { Text = "5" };
                second.Tables.Add(new TableDefinition { Name = "NotesTable", Range = new Rectangle(0, 0, 2, 2) });
                second.Charts.Add(new ChartDefinition { Title = "Notes Chart", Kind = "Column",
                    Range = new Rectangle(0, 0, 2, 2) });
                var dropdown = new ValidationRule { Range = new Rectangle(2, 0, 1, 2) };
                dropdown.Choices.Add("Yes"); dropdown.Choices.Add("No");
                second.Validations.Add(dropdown);
                book.Sheets.Add(second);
                book.NamedRanges.Add(new NamedRange { Name = "Revenue", Sheet = "Sales", Range = new Rectangle(0, 1, 1, 1) });
                XlsxFile.Write(path, book, 200, 26);
                var read = XlsxFile.Read(path, 200, 26);
                Equal(2, read.Sheets.Count);
                Equal("Sales", read.Sheets[0].Name);
                Equal("Ghi chú", read.Sheets[1].Name);
                Equal("Xin chào", read.Sheets[1].Cells[0].Text);
                Equal("NotesTable", read.Sheets[1].Tables[0].Name);
                Equal("Notes Chart", read.Sheets[1].Charts[0].Title);
                Equal(2, read.Sheets[1].Validations[0].Choices.Count);
                Equal("Revenue", read.NamedRanges[0].Name);
                Equal("=A2*2", read.Sheets[0].Cells[27].Text);
                Equal("#,##0.00", read.Sheets[0].Cells[26].NumberFormat);
                Equal(true, read.Sheets[0].Cells[26].HasFont);
                Equal(FontStyle.Bold, read.Sheets[0].Cells[26].FontStyle);
                Equal(Color.LightYellow.ToArgb(), read.Sheets[0].Cells[26].BackColor.ToArgb());
                Equal(Color.DarkBlue.ToArgb(), read.Sheets[0].Cells[26].ForeColor.ToArgb());
                Equal(DataGridViewContentAlignment.MiddleRight, read.Sheets[0].Cells[26].Alignment);
                Equal(1, read.Sheets[0].Merges.Count);
                Equal(1, read.Sheets[0].FreezeRow);
                Equal(0, read.Sheets[0].FilterColumn);
                Equal("12", read.Sheets[0].FilterValue);
                Equal(1, read.Sheets[0].Rules.Count);
                Equal(1000.0, read.Sheets[0].Rules[0].Threshold);
                Equal(180, read.Sheets[0].ColumnWidths[0]);
                book.Sheets[1].Cells[0].Text = "Updated";
                XlsxFile.Write(path, book, 200, 26);
                Equal("Updated", XlsxFile.Read(path, 200, 26).Sheets[1].Cells[0].Text);
                Console.WriteLine("XLSX: multi-sheet round trip passed.");
                if (Environment.GetEnvironmentVariable("DINKCEL_KEEP_XLSX") == "1") Console.WriteLine(path);
                var imported = XlsxFile.Read(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                    "fixtures", "office.xlsx"), 200, 26);
                Equal(2, imported.Sheets.Count);
                Equal("Imported", imported.Sheets[0].Name);
                Equal("=A2*3", imported.Sheets[0].Cells[27].Text);
                Equal("0.00", imported.Sheets[0].Cells[26].NumberFormat);
                Equal(FontStyle.Bold, imported.Sheets[0].Cells[26].FontStyle);
                Equal(Color.Blue.ToArgb(), imported.Sheets[0].Cells[26].ForeColor.ToArgb());
                Equal(Color.Yellow.ToArgb(), imported.Sheets[0].Cells[26].BackColor.ToArgb());
                Equal(1, imported.Sheets[0].FreezeRow);
                Equal(1, imported.Sheets[0].FreezeColumn);
                Equal(1, imported.Sheets[0].Merges.Count);
                Console.WriteLine("XLSX: third-party import passed.");
                var externalV3 = XlsxFile.Read(Path.Combine(AppDomain.CurrentDomain.BaseDirectory,
                    "fixtures", "thirdparty_v3.xlsx"), 200, 26);
                Equal(1, externalV3.Sheets[0].Charts.Count);
                Equal("External Sales", externalV3.Sheets[0].Charts[0].Title);
                Equal(1, externalV3.Sheets[0].Tables.Count);
                Equal("SalesTable", externalV3.Sheets[0].Tables[0].Name);
                Equal(1, externalV3.Sheets[0].Validations.Count);
                Equal(2, externalV3.Sheets[0].Validations[0].Choices.Count);
                Equal("Amounts", externalV3.NamedRanges[0].Name);
                Console.WriteLine("XLSX: third-party v0.3 features passed.");
            }
            finally { if (File.Exists(path) && Environment.GetEnvironmentVariable("DINKCEL_KEEP_XLSX") != "1") File.Delete(path); }
        }
    }
}
