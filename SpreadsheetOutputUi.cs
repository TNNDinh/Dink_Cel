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

        private PrintDocument CreatePrintDocument()
        {
            grid.EndEdit();
            int lastRow = 0, lastColumn = 0;
            for (int r = 0; r < RowCount; r++)
                for (int c = 0; c < ColumnCount; c++)
                    if (!string.IsNullOrEmpty(Convert.ToString(grid[c, r].Value)))
                    { lastRow = Math.Max(lastRow, r); lastColumn = Math.Max(lastColumn, c); }
            const int rowsPerPage = 28, columnsPerPage = 8;
            int rowPages = lastRow / rowsPerPage + 1;
            int columnPages = lastColumn / columnsPerPage + 1;
            int pageIndex = 0;
            var document = new PrintDocument();
            document.DocumentName = sheets[activeSheetIndex].Name + " - DinkCel";
            document.DefaultPageSettings.Landscape = true;
            document.PrintPage += delegate(object sender, PrintPageEventArgs e)
            {
                int columnPage = pageIndex / rowPages;
                int rowPage = pageIndex % rowPages;
                Rectangle bounds = e.MarginBounds;
                float top = bounds.Top + 34;
                float cellHeight = Math.Min(22F, (bounds.Height - 60F) / (rowsPerPage + 1));
                float width = (bounds.Width - 38F) / columnsPerPage;
                using (var titleFont = new Font("Arial", 13F, FontStyle.Bold))
                using (var headerFont = new Font("Arial", 8F, FontStyle.Bold))
                using (var cellFont = new Font("Arial", 8F))
                using (var border = new Pen(Color.LightGray))
                {
                    e.Graphics.DrawString(sheets[activeSheetIndex].Name, titleFont, Brushes.Black,
                        bounds.Left, bounds.Top);
                    for (int c = 0; c < columnsPerPage; c++)
                    {
                        int column = columnPage * columnsPerPage + c;
                        if (column >= ColumnCount) break;
                        var rectangle = new RectangleF(bounds.Left + 38 + c * width, top, width, cellHeight);
                        e.Graphics.DrawRectangle(border, rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
                        e.Graphics.DrawString(grid.Columns[column].HeaderText, headerFont, Brushes.Black, rectangle);
                    }
                    for (int r = 0; r < rowsPerPage; r++)
                    {
                        int row = rowPage * rowsPerPage + r;
                        if (row > lastRow) break;
                        float y = top + (r + 1) * cellHeight;
                        e.Graphics.DrawString((row + 1).ToString(), headerFont, Brushes.Black,
                            bounds.Left, y);
                        for (int c = 0; c < columnsPerPage; c++)
                        {
                            int column = columnPage * columnsPerPage + c;
                            if (column >= ColumnCount) break;
                            var rectangle = new RectangleF(bounds.Left + 38 + c * width, y, width, cellHeight);
                            e.Graphics.DrawRectangle(border, rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
                            string text = Convert.ToString(grid[column, row].FormattedValue) ?? "";
                            if (text.Length > 50) text = text.Substring(0, 49) + "…";
                            e.Graphics.DrawString(text.Replace('\n', ' '), cellFont, Brushes.Black, rectangle);
                        }
                    }
                }
                pageIndex++;
                e.HasMorePages = pageIndex < rowPages * columnPages;
            };
            return document;
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

        private void ExportPdf()
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
                    PdfFile.Write(dialog.FileName, book, RowCount, ColumnCount);
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
