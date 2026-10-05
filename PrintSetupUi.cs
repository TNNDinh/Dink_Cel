using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DinkCel
{
    internal sealed partial class SpreadsheetForm
    {
        private static NumericUpDown MarginControl(int hundredths)
        {
            return new NumericUpDown { DecimalPlaces = 2, Minimum = 0, Maximum = 3,
                Increment = .05M, Value = Math.Max(0, Math.Min(300, hundredths)) / 100M };
        }

        private void ConfigurePageSetup()
        {
            var layout = new TableLayoutPanel { AutoScroll = true };
            var landscape = new CheckBox { Text = "Ngang (Landscape)", Checked = printSettings.Landscape, AutoSize = true };
            var paper = DataChoice("A4", "Letter", "Legal", "A3"); paper.Text = printSettings.Paper;
            var left = MarginControl(printSettings.MarginLeft);
            var right = MarginControl(printSettings.MarginRight);
            var top = MarginControl(printSettings.MarginTop);
            var bottom = MarginControl(printSettings.MarginBottom);
            var scale = new NumericUpDown { Minimum = 10, Maximum = 400,
                Value = Math.Max(10, Math.Min(400, printSettings.Scale)) };
            var fit = new CheckBox { Text = "Vừa một trang", Checked = printSettings.FitToOnePage, AutoSize = true };
            var titles = new NumericUpDown { Minimum = 0, Maximum = RowCount,
                Value = Math.Max(0, Math.Min(RowCount, printSettings.TitleRows)) };
            var header = new TextBox { Text = printSettings.Header };
            var footer = new TextBox { Text = printSettings.Footer };
            var gridlines = new CheckBox { Text = "In đường lưới", Checked = printSettings.Gridlines, AutoSize = true };
            var chart = new CheckBox { Text = "In biểu đồ", Checked = printSettings.PrintCharts, AutoSize = true };
            var area = new TextBox { Text = printSettings.PrintArea.IsEmpty ? "" :
                ChartRangeText(printSettings.PrintArea) };
            DataField(layout, "Khổ giấy", paper); DataField(layout, "Hướng", landscape);
            DataField(layout, "Lề trái (inch)", left); DataField(layout, "Lề phải (inch)", right);
            DataField(layout, "Lề trên (inch)", top); DataField(layout, "Lề dưới (inch)", bottom);
            DataField(layout, "Thu phóng (%)", scale); DataField(layout, "Vừa trang", fit);
            DataField(layout, "Lặp hàng đầu", titles); DataField(layout, "Vùng in", area);
            DataField(layout, "Đầu trang", header); DataField(layout, "Chân trang", footer);
            DataField(layout, "Đường lưới", gridlines); DataField(layout, "Biểu đồ", chart);
            using (var dialog = DataDialog("Thiết lập trang in", layout))
            {
                dialog.Width = 520; dialog.Height = 660;
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                Rectangle selected = Rectangle.Empty;
                if (!string.IsNullOrWhiteSpace(area.Text) && !ParseChartRange(area.Text, out selected))
                { MessageBox.Show(this, "Vùng in cần có dạng A1:C10."); return; }
                printSettings.Paper = paper.Text; printSettings.Landscape = landscape.Checked;
                printSettings.MarginLeft = (int)(left.Value * 100);
                printSettings.MarginRight = (int)(right.Value * 100);
                printSettings.MarginTop = (int)(top.Value * 100);
                printSettings.MarginBottom = (int)(bottom.Value * 100);
                printSettings.Scale = (int)scale.Value; printSettings.FitToOnePage = fit.Checked;
                printSettings.TitleRows = (int)titles.Value; printSettings.PrintArea = selected;
                printSettings.Header = header.Text; printSettings.Footer = footer.Text;
                printSettings.Gridlines = gridlines.Checked; printSettings.PrintCharts = chart.Checked;
                RecordChange(); MarkDirty();
            }
        }

        private void SetPrintArea()
        {
            printSettings.PrintArea = SelectionRange(false);
            RecordChange(); MarkDirty();
            status.Text = "Đã đặt vùng in " + ChartRangeText(printSettings.PrintArea);
        }

        private void ClearPrintArea()
        {
            printSettings.PrintArea = Rectangle.Empty;
            RecordChange(); MarkDirty(); status.Text = "Đã bỏ vùng in";
        }

        private void AddPageBreak()
        {
            if (grid.CurrentCell == null || grid.CurrentCell.RowIndex == 0) return;
            int row = grid.CurrentCell.RowIndex;
            if (!printSettings.PageBreakRows.Contains(row)) printSettings.PageBreakRows.Add(row);
            printSettings.PageBreakRows.Sort();
            RecordChange(); MarkDirty(); status.Text = "Đã ngắt trang trước hàng " + (row + 1);
        }

        private void RemovePageBreak()
        {
            if (grid.CurrentCell == null) return;
            if (printSettings.PageBreakRows.Remove(grid.CurrentCell.RowIndex))
            { RecordChange(); MarkDirty(); status.Text = "Đã bỏ ngắt trang"; }
        }

        private void PrintSelection()
        {
            using (var document = CreatePrintDocument(SelectionRange(false)))
            using (var dialog = new System.Windows.Forms.PrintDialog { Document = document, UseEXDialog = true })
                if (dialog.ShowDialog(this) == DialogResult.OK) document.Print();
        }
    }
}
