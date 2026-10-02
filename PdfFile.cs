using System;
using System.Globalization;
using System.IO;
using System.Linq;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace DinkCel
{
    internal static class PdfFile
    {
        private static string ColumnName(int column)
        {
            string text = "";
            do { text = (char)('A' + column % 26) + text; column = column / 26 - 1; } while (column >= 0);
            return text;
        }

        public static void Write(string path, WorkbookSnapshot workbook, int rowCount, int columnCount)
        {
            EmbeddedDependencies.Install();
            var document = new PdfDocument();
            document.Info.Title = "DinkCel workbook";
            var options = new XPdfFontOptions(PdfFontEncoding.Unicode);
            var titleFont = new XFont("Arial", 14, XFontStyle.Bold, options);
            var headerFont = new XFont("Arial", 8, XFontStyle.Bold, options);
            var cellFont = new XFont("Arial", 8, XFontStyle.Regular, options);
            const int rowsPerPage = 26, columnsPerPage = 8;
            foreach (SheetSnapshot sheet in workbook.Sheets)
            {
                int lastRow = sheet.Cells.Count == 0 ? 0 : sheet.Cells.Keys.Max() / columnCount;
                int lastColumn = sheet.Cells.Count == 0 ? 0 : sheet.Cells.Keys.Max() % columnCount;
                foreach (int key in sheet.Cells.Keys) lastColumn = Math.Max(lastColumn, key % columnCount);
                int rowPages = Math.Max(1, lastRow / rowsPerPage + 1);
                int columnPages = Math.Max(1, lastColumn / columnsPerPage + 1);
                var engine = new FormulaEngine((r, c) => Raw(sheet, r, c, columnCount),
                    (name, r, c) =>
                    {
                        SheetSnapshot other = workbook.Sheets.FirstOrDefault(s =>
                            string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
                        return other == null ? null : Raw(other, r, c, columnCount);
                    }, sheet.Name, rowCount, columnCount, name =>
                    {
                        NamedRange named = workbook.NamedRanges.FirstOrDefault(n =>
                            string.Equals(n.Name, name, StringComparison.OrdinalIgnoreCase));
                        return named == null ? null : new FormulaNamedRange
                        { Sheet = named.Sheet, FirstRow = named.Range.Top, FirstColumn = named.Range.Left,
                            LastRow = named.Range.Bottom - 1, LastColumn = named.Range.Right - 1 };
                    });
                for (int cp = 0; cp < columnPages; cp++)
                    for (int rp = 0; rp < rowPages; rp++)
                    {
                        PdfPage page = document.AddPage();
                        page.Size = PdfSharp.PageSize.A4;
                        page.Orientation = PdfSharp.PageOrientation.Landscape;
                        using (XGraphics graphics = XGraphics.FromPdfPage(page))
                        {
                            double left = 32, top = 34, cellHeight = 18;
                            double width = (page.Width.Point - 64 - 38) / columnsPerPage;
                            graphics.DrawString(sheet.Name, titleFont, XBrushes.Black,
                                new XRect(left, 12, page.Width.Point - 64, 24), XStringFormats.CenterLeft);
                            graphics.DrawString("DinkCel • " + (rp + 1) + "/" + rowPages + " • " + (cp + 1) + "/" + columnPages,
                                cellFont, XBrushes.Gray, new XRect(left, page.Height.Point - 22,
                                    page.Width.Point - 64, 12), XStringFormats.CenterRight);
                            DrawCell(graphics, "#", left, top, 38, cellHeight, headerFont, true);
                            for (int c = 0; c < columnsPerPage; c++)
                                if (cp * columnsPerPage + c < columnCount)
                                    DrawCell(graphics, ColumnName(cp * columnsPerPage + c), left + 38 + c * width,
                                        top, width, cellHeight, headerFont, true);
                            for (int r = 0; r < rowsPerPage; r++)
                            {
                                int row = rp * rowsPerPage + r;
                                if (row > lastRow) break;
                                double y = top + (r + 1) * cellHeight;
                                DrawCell(graphics, (row + 1).ToString(), left, y, 38, cellHeight, headerFont, true);
                                for (int c = 0; c < columnsPerPage; c++)
                                {
                                    int column = cp * columnsPerPage + c;
                                    if (column >= columnCount) break;
                                    CellSnapshot cell;
                                    string raw = sheet.Cells.TryGetValue(row * columnCount + column, out cell) ? cell.Text ?? "" : "";
                                    string value = raw.StartsWith("=", StringComparison.Ordinal) ? engine.Display(row, column) : raw;
                                    if (cell != null && !string.IsNullOrEmpty(cell.NumberFormat))
                                    {
                                        double number;
                                        if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
                                            try { value = number.ToString(cell.NumberFormat, CultureInfo.CurrentCulture); }
                                            catch (FormatException) { }
                                    }
                                    DrawCell(graphics, value, left + 38 + c * width, y, width, cellHeight, cellFont, false);
                                }
                            }
                        }
                    }
            }
            string temporary = path + ".tmp";
            try
            {
                document.Save(temporary);
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
            }
            finally { document.Close(); if (File.Exists(temporary)) File.Delete(temporary); }
        }

        private static string Raw(SheetSnapshot sheet, int row, int column, int columns)
        {
            CellSnapshot cell;
            return sheet.Cells.TryGetValue(row * columns + column, out cell) ? cell.Text ?? "" : "";
        }

        private static void DrawCell(XGraphics graphics, string value, double x, double y,
            double width, double height, XFont font, bool header)
        {
            var bounds = new XRect(x, y, width, height);
            if (header) graphics.DrawRectangle(XBrushes.LightGray, bounds);
            graphics.DrawRectangle(XPens.LightGray, bounds);
            value = (value ?? "").Replace('\r', ' ').Replace('\n', ' ');
            if (value.Length > 50) value = value.Substring(0, 49) + "…";
            graphics.DrawString(value, font, XBrushes.Black,
                new XRect(x + 3, y + 1, width - 6, height - 2), XStringFormats.CenterLeft);
        }
    }
}
