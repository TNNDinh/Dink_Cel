using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

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
            XDocument themeStyles = XDocument.Parse("<styleSheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'>" +
                "<fonts><font><sz val='11'/><name val='Aptos'/><color theme='4' tint='0.5'/></font></fonts>" +
                "<fills><fill><patternFill patternType='none'/></fill><fill><patternFill patternType='gray125'/></fill>" +
                "<fill><patternFill patternType='solid'><fgColor theme='5'/></patternFill></fill></fills>" +
                "<borders><border/></borders><cellXfs><xf fontId='0' fillId='2' borderId='0'>" +
                "<alignment horizontal='center' vertical='bottom'/></xf></cellXfs></styleSheet>");
            XDocument theme = XDocument.Parse("<a:theme xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main'>" +
                "<a:themeElements><a:clrScheme><a:accent1><a:srgbClr val='112233'/></a:accent1>" +
                "<a:accent2><a:srgbClr val='445566'/></a:accent2></a:clrScheme></a:themeElements></a:theme>");
            var themed = XlsxStyles.Read(themeStyles, theme)[0];
            Equal("Aptos", themed.FontName);
            Equal(true, themed.HasFont);
            Equal(Color.FromArgb(136, 144, 153).ToArgb(), themed.ForeColor.ToArgb());
            Equal(Color.FromArgb(68, 85, 102).ToArgb(), themed.BackColor.ToArgb());
            Equal(DataGridViewContentAlignment.BottomCenter, themed.Alignment);
            string path = Path.Combine(Path.GetTempPath(), "DinkCel_xlsx_" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                var book = new WorkbookSnapshot();
                book.StructureProtected = true;
                book.Sheets[0].Name = "Sales";
                book.Sheets[0].Protected = true;
                book.Sheets[0].Cells[0] = new CellSnapshot { Text = "Revenue" };
                book.Sheets[0].Cells[26] = new CellSnapshot { Text = "1234.5", NumberFormat = "#,##0.00",
                    HasFont = true, FontName = "Consolas", FontStyle = FontStyle.Bold | FontStyle.Strikeout, FontSize = 14F,
                    ForeColor = Color.DarkBlue, BackColor = Color.LightYellow,
                    Alignment = DataGridViewContentAlignment.TopRight,
                    Extras = new CellExtras { Wrap = true, Shrink = true, Indent = 2, Rotation = -30,
                        Left = new BorderEdge { Style = "thin", Color = Color.Red },
                        Bottom = new BorderEdge { Style = "double", Color = Color.Blue } } };
                book.Sheets[0].Cells[27] = new CellSnapshot { Text = "=A2*2" };
                book.Sheets[0].Cells[2] = new CellSnapshot { Text = "Website",
                    Extras = new CellExtras { Hyperlink = "https://example.com/report" } };
                book.Sheets[0].Cells[3] = new CellSnapshot { Text = "Other sheet",
                    Extras = new CellExtras { Hyperlink = "#'Ghi chú'!A1" } };
                book.Sheets[0].Cells[4] = new CellSnapshot { Text = "Memo",
                    Extras = new CellExtras { Note = "Check this number" } };
                book.Sheets[0].Merges.Add(new Rectangle(0, 0, 2, 1));
                book.Sheets[0].FreezeRow = 1;
                book.Sheets[0].ShowGridlines = false;
                book.Sheets[0].ShowHeadings = false;
                book.Sheets[0].FormulaView = true;
                book.Sheets[0].ViewMode = "pageBreakPreview";
                book.Sheets[0].FilterColumn = 0;
                book.Sheets[0].FilterValue = "12";
                book.Sheets[0].Rules.Add(new ConditionalRule { Range = new Rectangle(0, 1, 1, 1), Threshold = 1000, Color = Color.LightGreen });
                book.Sheets[0].ColumnWidths[0] = 180;
                book.Sheets[0].RowHeights[0] = 35;
                book.Sheets[0].HiddenRows.Add(4);
                book.Sheets[0].HiddenColumns.Add(3);
                var second = new SheetSnapshot { Name = "Ghi chú" };
                second.SplitX = 1800;
                second.SplitY = 900;
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
                Equal(true, read.StructureProtected);
                Equal(true, read.Sheets[0].Protected);
                Equal("Sales", read.Sheets[0].Name);
                Equal("Ghi chú", read.Sheets[1].Name);
                Equal("Xin chào", read.Sheets[1].Cells[0].Text);
                Equal("NotesTable", read.Sheets[1].Tables[0].Name);
                Equal("Notes Chart", read.Sheets[1].Charts[0].Title);
                Equal(1800.0, read.Sheets[1].SplitX);
                Equal(900.0, read.Sheets[1].SplitY);
                Equal(2, read.Sheets[1].Validations[0].Choices.Count);
                Equal("Revenue", read.NamedRanges[0].Name);
                Equal("=A2*2", read.Sheets[0].Cells[27].Text);
                Equal("https://example.com/report", read.Sheets[0].Cells[2].Extras.Hyperlink);
                Equal("#'Ghi chú'!A1", read.Sheets[0].Cells[3].Extras.Hyperlink);
                Equal("Check this number", read.Sheets[0].Cells[4].Extras.Note);
                Equal("#,##0.00", read.Sheets[0].Cells[26].NumberFormat);
                Equal(true, read.Sheets[0].Cells[26].HasFont);
                Equal(FontStyle.Bold | FontStyle.Strikeout, read.Sheets[0].Cells[26].FontStyle);
                Equal("Consolas", read.Sheets[0].Cells[26].FontName);
                Equal(true, read.Sheets[0].Cells[26].Extras.Wrap);
                Equal(true, read.Sheets[0].Cells[26].Extras.Shrink);
                Equal(2, read.Sheets[0].Cells[26].Extras.Indent);
                Equal(-30, read.Sheets[0].Cells[26].Extras.Rotation);
                Equal("thin", read.Sheets[0].Cells[26].Extras.Left.Style);
                Equal(Color.Red.ToArgb(), read.Sheets[0].Cells[26].Extras.Left.Color.ToArgb());
                Equal("double", read.Sheets[0].Cells[26].Extras.Bottom.Style);
                Equal(true, read.Sheets[0].HiddenRows.Contains(4));
                Equal(true, read.Sheets[0].HiddenColumns.Contains(3));
                Equal(Color.LightYellow.ToArgb(), read.Sheets[0].Cells[26].BackColor.ToArgb());
                Equal(Color.DarkBlue.ToArgb(), read.Sheets[0].Cells[26].ForeColor.ToArgb());
                Equal(DataGridViewContentAlignment.TopRight, read.Sheets[0].Cells[26].Alignment);
                Equal(1, read.Sheets[0].Merges.Count);
                Equal(1, read.Sheets[0].FreezeRow);
                Equal(false, read.Sheets[0].ShowGridlines);
                Equal(false, read.Sheets[0].ShowHeadings);
                Equal(true, read.Sheets[0].FormulaView);
                Equal("pageBreakPreview", read.Sheets[0].ViewMode);
                Equal(0, read.Sheets[0].FilterColumn);
                Equal("12", read.Sheets[0].FilterValue);
                Equal(1, read.Sheets[0].Rules.Count);
                Equal(1000.0, read.Sheets[0].Rules[0].Threshold);
                Equal(180, read.Sheets[0].ColumnWidths[0]);
                book.Sheets[1].Cells[0].Text = "Updated";
                XlsxFile.Write(path, book, 200, 26);
                Equal("Updated", XlsxFile.Read(path, 200, 26).Sheets[1].Cells[0].Text);
                Console.WriteLine("XLSX: multi-sheet round trip passed.");
                using (var archive = ZipFile.Open(path, ZipArchiveMode.Update))
                {
                    using (var stream = new StreamWriter(archive.CreateEntry("xl/vbaProject.bin").Open())) stream.Write("fixture");
                    using (var stream = new StreamWriter(archive.CreateEntry("xl/externalLinks/externalLink1.xml").Open())) stream.Write("fixture");
                    using (var stream = new StreamWriter(archive.CreateEntry("xl/media/image1.png").Open())) stream.Write("fixture");
                    using (var stream = new StreamWriter(archive.CreateEntry("xl/comments1.xml").Open())) stream.Write("fixture");
                }
                var losses = XlsxFile.PotentialLosses(path);
                Equal(true, losses.Contains("VBA/macros"));
                Equal(true, losses.Contains("External connections and linked workbooks"));
                Equal(true, losses.Contains("Images, shapes or drawing placement"));
                Equal(true, losses.Contains("Comment formatting or authors"));
                Equal(true, losses.Any(x => x.Contains("Chart")));
                Equal(true, losses.Contains("Split panes and independent scrolling"));
                Equal("Updated", XlsxFile.Read(path, 200, 26).Sheets[1].Cells[0].Text);
                Console.WriteLine("XLSX: unsupported feature audit passed.");
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
