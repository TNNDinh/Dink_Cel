using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using PdfSharp.Drawing;
using PdfSharp.Pdf;

namespace DinkCel
{
    internal static class PdfFile
    {
        private static readonly XPdfFontOptions FontOptions = new XPdfFontOptions(PdfFontEncoding.Unicode);

        public static void Write(string path, WorkbookSnapshot workbook, int rowCount, int columnCount)
        { Write(path, workbook, rowCount, columnCount, Rectangle.Empty, null); }

        public static void Write(string path, WorkbookSnapshot workbook, int rowCount, int columnCount,
            Rectangle selection, string selectedSheet)
        {
            EmbeddedDependencies.Install();
            var document = new PdfDocument();
            document.Info.Title = "DinkCel workbook";
            foreach (SheetSnapshot sheet in workbook.Sheets)
            {
                if (selectedSheet != null && !string.Equals(sheet.Name, selectedSheet,
                    StringComparison.OrdinalIgnoreCase)) continue;
                var pages = PrintLayout.Plan(sheet, rowCount, columnCount,
                    selectedSheet == null ? Rectangle.Empty : selection, selectedSheet == null);
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
                    }, null, (currentSheet, tableName, header, currentRow) =>
                    {
                        foreach (SheetSnapshot source in workbook.Sheets)
                        {
                            if (tableName == null && !string.Equals(source.Name, currentSheet,
                                StringComparison.OrdinalIgnoreCase)) continue;
                            foreach (TableDefinition table in source.Tables)
                            {
                                if (tableName != null && !string.Equals(table.Name, tableName,
                                    StringComparison.OrdinalIgnoreCase)) continue;
                                int first = table.Range.Top + (table.HeaderRow ? 1 : 0);
                                int last = table.Range.Bottom - (table.TotalRow ? 1 : 0) - 1;
                                if (currentRow >= 0 && (currentRow < first || currentRow > last)) continue;
                                for (int column = table.Range.Left; column < table.Range.Right; column++)
                                {
                                    string name = table.HeaderRow ? Raw(source, table.Range.Top,
                                        column, columnCount) : "Column" + (column - table.Range.Left + 1);
                                    if (!string.Equals(name, header, StringComparison.OrdinalIgnoreCase)) continue;
                                    return new FormulaNamedRange { Sheet = source.Name,
                                        FirstRow = currentRow >= 0 ? currentRow : first,
                                        LastRow = currentRow >= 0 ? currentRow : last,
                                        FirstColumn = column, LastColumn = column };
                                }
                            }
                        }
                        return null;
                    });
                Func<int, int, string> display = (r, c) =>
                {
                    string raw = Raw(sheet, r, c, columnCount);
                    return raw.StartsWith("=", StringComparison.Ordinal) ? engine.Display(r, c) : raw;
                };
                foreach (PrintPagePlan plan in pages)
                {
                    var page = document.AddPage();
                    page.Size = sheet.Print.Paper == "A3" ? PdfSharp.PageSize.A3 :
                        sheet.Print.Paper == "Letter" ? PdfSharp.PageSize.Letter :
                        sheet.Print.Paper == "Legal" ? PdfSharp.PageSize.Legal : PdfSharp.PageSize.A4;
                    page.Orientation = sheet.Print.Landscape ? PdfSharp.PageOrientation.Landscape :
                        PdfSharp.PageOrientation.Portrait;
                    using (XGraphics graphics = XGraphics.FromPdfPage(page))
                        DrawPage(graphics, page, plan, display, columnCount);
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

        private static double ColumnWidth(SheetSnapshot sheet, int column)
        { int value; return (sheet.ColumnWidths.TryGetValue(column, out value) ? value : 120) * .75; }

        private static double RowHeight(SheetSnapshot sheet, int row)
        { int value; return (sheet.RowHeights.TryGetValue(row, out value) ? value : 27) * .75; }

        private static XBrush Brush(Color color)
        { return new XSolidBrush(XColor.FromArgb(color.R, color.G, color.B)); }

        private static void DrawPage(XGraphics graphics, PdfPage pdfPage, PrintPagePlan plan,
            Func<int, int, string> display, int columns)
        {
            SheetSnapshot sheet = plan.Sheet;
            PrintSettings settings = sheet.Print;
            double left = settings.MarginLeft * .72;
            double top = settings.MarginTop * .72;
            double pageWidth = pdfPage.Width.Point, pageHeight = pdfPage.Height.Point;
            var small = new XFont("Arial", 9, XFontStyle.Regular, FontOptions);
            string header = PrintLayout.HeaderFooter(settings.Header, plan);
            string footer = PrintLayout.HeaderFooter(settings.Footer, plan);
            if (header.Length > 0) graphics.DrawString(header, small, XBrushes.Black,
                new XRect(left, Math.Max(2, top - 24), pageWidth - left, 18), XStringFormats.CenterLeft);
            if (footer.Length > 0) graphics.DrawString(footer, small, XBrushes.Black,
                new XRect(left, pageHeight - settings.MarginBottom * .72 + 5,
                    pageWidth - left - settings.MarginRight * .72, 18), XStringFormats.CenterRight);
            if (plan.Chart != null)
            {
                using (var chart = ChartRendering.Build(plan.Chart, display,
                    Color.White, Color.Black, Color.SteelBlue))
                {
                    chart.Size = new Size(1100, 700);
                    using (var stream = new MemoryStream())
                    {
                        chart.SaveImage(stream,
                            System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Png);
                        stream.Position = 0;
                        using (XImage image = XImage.FromStream(stream))
                            graphics.DrawImage(image, left, top,
                                pageWidth - left - settings.MarginRight * .72,
                                pageHeight - top - settings.MarginBottom * .72 - 30);
                    }
                }
                return;
            }
            double y = top;
            Action<int> drawRow = row =>
            {
                double x = left;
                double height = RowHeight(sheet, row) * plan.Scale;
                for (int column = plan.FirstColumn; column < plan.LastColumn; column++)
                {
                    double width = ColumnWidth(sheet, column) * plan.Scale;
                    var rect = new XRect(x, y, width, height);
                    CellSnapshot cell;
                    sheet.Cells.TryGetValue(row * columns + column, out cell);
                    if (cell != null && !cell.BackColor.IsEmpty)
                        graphics.DrawRectangle(Brush(cell.BackColor), rect);
                    if (settings.Gridlines) graphics.DrawRectangle(XPens.LightGray, rect);
                    string text = display(row, column).Replace('\r', ' ').Replace('\n', ' ');
                    if (cell != null && !string.IsNullOrEmpty(cell.NumberFormat))
                    {
                        double number;
                        if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
                            try { text = number.ToString(cell.NumberFormat, CultureInfo.CurrentCulture); }
                            catch (FormatException) { }
                    }
                    if (text.Length > 60) text = text.Substring(0, 59) + "…";
                    XFontStyle style = cell != null && cell.HasFont &&
                        (cell.FontStyle & FontStyle.Bold) != 0 ? XFontStyle.Bold : XFontStyle.Regular;
                    var font = new XFont("Arial", Math.Max(6, cell != null && cell.HasFont ?
                        cell.FontSize * plan.Scale : 8 * plan.Scale), style, FontOptions);
                    graphics.DrawString(text, font, cell == null || cell.ForeColor.IsEmpty ?
                        XBrushes.Black : Brush(cell.ForeColor),
                        new XRect(x + 2, y + 1, Math.Max(2, width - 4), Math.Max(2, height - 2)),
                        XStringFormats.CenterLeft);
                    x += width;
                }
                y += height;
            };
            if (plan.TitleRows > 0)
                for (int row = 0; row < plan.TitleRows; row++) drawRow(row);
            for (int row = plan.FirstRow; row < plan.LastRow; row++) drawRow(row);
        }
    }
}
