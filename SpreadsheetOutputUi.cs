using System;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace DinkCel
{
    internal sealed partial class SpreadsheetForm
    {
        private void ShowThirdPartyLicenses()
        {
            string content;
            using (Stream source = GetType().Assembly.GetManifestResourceStream("ThirdPartyLicenses.txt"))
            {
                if (source == null) { MessageBox.Show(this, "Không tìm thấy giấy phép thư viện."); return; }
                using (var reader = new StreamReader(source, Encoding.UTF8)) content = reader.ReadToEnd();
            }
            using (var window = new Form { Text = "Giấy phép thư viện - DinkCel",
                Width = 900, Height = 650, StartPosition = FormStartPosition.CenterParent })
            {
                var box = new TextBox { Dock = DockStyle.Fill, ReadOnly = true, Multiline = true,
                    ScrollBars = ScrollBars.Both, WordWrap = false, Text = content,
                    Font = new Font("Consolas", 9F) };
                window.Controls.Add(box);
                window.ShowDialog(this);
            }
        }

        private PrintDocument CreatePrintDocument() { return CreatePrintDocument(Rectangle.Empty); }

        private PrintDocument CreatePrintDocument(Rectangle selection)
        {
            grid.EndEdit(); SaveActiveSheet();
            SheetSnapshot sheet = SnapshotFromState(sheets[activeSheetIndex]);
            var pages = PrintLayout.Plan(sheet, RowCount, ColumnCount, selection, selection.IsEmpty);
            int pageIndex = 0;
            var document = new PrintDocument { DocumentName = sheet.Name + " - DinkCel" };
            Size paper = PrintLayout.PaperSize(sheet.Print);
            document.DefaultPageSettings.PaperSize = new PaperSize(sheet.Print.Paper,
                sheet.Print.Landscape ? paper.Height : paper.Width,
                sheet.Print.Landscape ? paper.Width : paper.Height);
            document.DefaultPageSettings.Landscape = sheet.Print.Landscape;
            document.DefaultPageSettings.Margins = new Margins(sheet.Print.MarginLeft,
                sheet.Print.MarginRight, sheet.Print.MarginTop, sheet.Print.MarginBottom);
            document.BeginPrint += delegate { pageIndex = 0; };
            document.PrintPage += delegate(object sender, PrintPageEventArgs e)
            {
                if (pageIndex >= pages.Count) { e.HasMorePages = false; return; }
                DrawPrintPage(e.Graphics, e.PageBounds, pages[pageIndex++], sheet);
                e.HasMorePages = pageIndex < pages.Count;
            };
            return document;
        }

        private void DrawPrintPage(Graphics graphics, Rectangle pageBounds,
            PrintPagePlan page, SheetSnapshot sheet)
        {
            PrintSettings options = sheet.Print;
            float left = options.MarginLeft, top = options.MarginTop;
            using (var font = new Font("Arial", 9F))
            using (var pen = new Pen(Color.LightGray))
            {
                graphics.DrawString(PrintLayout.HeaderFooter(options.Header, page), font,
                    Brushes.Black, left, Math.Max(4, top - 30));
                graphics.DrawString(PrintLayout.HeaderFooter(options.Footer, page), font,
                    Brushes.Black, left, pageBounds.Height - options.MarginBottom + 12);
                if (page.Chart != null)
                {
                    using (var chart = ChartRendering.Build(page.Chart, (r, c) =>
                        DisplayPrintCell(sheet, r, c), Color.White, Color.Black, Color.SteelBlue))
                    {
                        chart.Size = new Size(1000, 650);
                        using (var stream = new MemoryStream())
                        {
                            chart.SaveImage(stream, System.Windows.Forms.DataVisualization.Charting.ChartImageFormat.Png);
                            stream.Position = 0;
                            using (Image image = Image.FromStream(stream))
                                graphics.DrawImage(image, left, top, pageBounds.Width - options.MarginLeft -
                                    options.MarginRight, pageBounds.Height - options.MarginTop - options.MarginBottom - 45);
                        }
                    }
                    return;
                }
                float y = top;
                Action<int> drawRow = row =>
                {
                    float x = left;
                    float height = (float)(SheetRowHeight(sheet, row) * page.Scale);
                    for (int column = page.FirstColumn; column < page.LastColumn; column++)
                    {
                        float width = (float)(SheetColumnWidth(sheet, column) * page.Scale);
                        var box = new RectangleF(x, y, width, height);
                        CellSnapshot cell;
                        sheet.Cells.TryGetValue(row * ColumnCount + column, out cell);
                        if (cell != null && !cell.BackColor.IsEmpty)
                            using (var fill = new SolidBrush(cell.BackColor)) graphics.FillRectangle(fill, box);
                        if (options.Gridlines) graphics.DrawRectangle(pen, box.X, box.Y, box.Width, box.Height);
                        string value = DisplayPrintCell(sheet, row, column).Replace('\r', ' ').Replace('\n', ' ');
                        if (value.Length > 80) value = value.Substring(0, 79) + "…";
                        using (var brush = new SolidBrush(cell == null || cell.ForeColor.IsEmpty ?
                            Color.Black : cell.ForeColor))
                        using (var cellFont = cell != null && cell.HasFont ? new Font(cell.FontName,
                            Math.Max(6, cell.FontSize * (float)page.Scale), cell.FontStyle) :
                            new Font("Arial", Math.Max(6, 9F * (float)page.Scale)))
                            graphics.DrawString(value, cellFont, brush, box);
                        x += width;
                    }
                    y += height;
                };
                if (page.TitleRows > 0)
                    for (int row = 0; row < page.TitleRows; row++) drawRow(row);
                for (int row = page.FirstRow; row < page.LastRow; row++) drawRow(row);
            }
        }

        private static double SheetColumnWidth(SheetSnapshot sheet, int column)
        { int width; return (sheet.ColumnWidths.TryGetValue(column, out width) ? width : 120) * 100.0 / 96; }

        private static double SheetRowHeight(SheetSnapshot sheet, int row)
        { int height; return (sheet.RowHeights.TryGetValue(row, out height) ? height : 27) * 100.0 / 96; }

        private static string SheetCellText(SheetSnapshot sheet, int row, int column)
        {
            CellSnapshot cell;
            return sheet.Cells.TryGetValue(row * ColumnCount + column, out cell) ? cell.Text ?? "" : "";
        }

        private string DisplayPrintCell(SheetSnapshot sheet, int row, int column)
        {
            string raw = SheetCellText(sheet, row, column);
            if (!raw.StartsWith("=", StringComparison.Ordinal)) return raw;
            string value;
            if (calculated.TryGetValue(row * ColumnCount + column, out value)) return value;
            return formulaEngine == null ? raw : formulaEngine.Display(row, column);
        }

        private void PreviewPrint()
        {
            using (PrintDocument document = CreatePrintDocument())
            using (var preview = new PrintPreviewDialog { Document = document,
                Width = 1000, Height = 700 })
                preview.ShowDialog(this);
        }

        private void PrintWorkbook()
        {
            using (PrintDocument document = CreatePrintDocument())
            using (var dialog = new PrintDialog { Document = document, UseEXDialog = true })
                if (dialog.ShowDialog(this) == DialogResult.OK) document.Print();
        }

        private void ExportPdf() { ExportPdf(Rectangle.Empty); }

        private void ExportPdfSelection() { ExportPdf(SelectionRange(false)); }

        private void ExportPdf(Rectangle selection)
        {
            using (var dialog = new SaveFileDialog { Filter = "PDF (*.pdf)|*.pdf",
                DefaultExt = "pdf", FileName = currentPath == null ? "BangTinh.pdf" :
                    Path.GetFileNameWithoutExtension(currentPath) + ".pdf" })
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    grid.EndEdit(); SaveActiveSheet();
                    var book = new WorkbookSnapshot(); book.Sheets.Clear();
                    book.NamedRanges.AddRange(namedRanges);
                    foreach (SheetState sheet in sheets) book.Sheets.Add(SnapshotFromState(sheet));
                    PdfFile.Write(dialog.FileName, book, RowCount, ColumnCount, selection,
                        selection.IsEmpty ? null : sheets[activeSheetIndex].Name);
                    status.Text = "Đã xuất PDF " + Path.GetFileName(dialog.FileName);
                }
                catch (Exception error)
                {
                    MessageBox.Show(this, "Không xuất được PDF: " + error.Message,
                        "DinkCel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }
}
