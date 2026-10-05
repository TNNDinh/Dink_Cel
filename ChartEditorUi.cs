using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace DinkCel
{
    internal sealed partial class SpreadsheetForm
    {
        private static string ChartRangeText(Rectangle range)
        {
            return ((char)('A' + range.Left)).ToString() + (range.Top + 1) + ":" +
                (char)('A' + range.Right - 1) + range.Bottom;
        }

        private static bool ParseChartRange(string text, out Rectangle range)
        {
            range = Rectangle.Empty;
            string[] ends = text.Split(':');
            if (ends.Length != 2) return false;
            int x1, y1, x2, y2;
            if (!ParseCellAddress(ends[0], out x1, out y1) ||
                !ParseCellAddress(ends[1], out x2, out y2)) return false;
            range = Rectangle.FromLTRB(Math.Min(x1, x2), Math.Min(y1, y2),
                Math.Max(x1, x2) + 1, Math.Max(y1, y2) + 1);
            return range.Width >= 1 && range.Height >= 1;
        }

        private bool ConfigureChart(ChartDefinition definition)
        {
            var layout = new TableLayoutPanel { AutoScroll = true };
            var title = new TextBox { Text = definition.Title };
            var kind = DataChoice("Column", "Line", "Pie", "Bar", "Area", "Scatter",
                "Stacked", "100% Stacked", "Combo");
            kind.Text = definition.Kind;
            var rangeBox = new TextBox { Text = ChartRangeText(definition.Range) };
            var axisX = new TextBox { Text = definition.AxisTitleX };
            var axisY = new TextBox { Text = definition.AxisTitleY };
            var legend = new CheckBox { Text = "Hiện chú giải", Checked = definition.Legend, AutoSize = true };
            var labels = new CheckBox { Text = "Hiện nhãn dữ liệu", Checked = definition.DataLabels, AutoSize = true };
            var gridlines = new CheckBox { Text = "Hiện đường lưới", Checked = definition.Gridlines, AutoSize = true };
            var series = new CheckedListBox { Height = 105, CheckOnClick = true };
            var seriesName = new TextBox();
            var colorButton = new Button { Text = "Chọn màu..." };
            var names = new Dictionary<int, string>(definition.SeriesNames);
            var colors = new Dictionary<int, Color>(definition.SeriesColors);
            int[] columns = new int[0];
            Action populate = delegate
            {
                Rectangle selected;
                if (!ParseChartRange(rangeBox.Text, out selected) || selected.Width < 2) return;
                series.Items.Clear();
                columns = Enumerable.Range(selected.Left + 1, selected.Width - 1).ToArray();
                for (int i = 0; i < columns.Length; i++)
                {
                    int column = columns[i];
                    string header = Convert.ToString(grid[column, selected.Top].Value) ?? "";
                    series.Items.Add((char)('A' + column) + " — " + header,
                        definition.SeriesColumns.Count == 0 || definition.SeriesColumns.Contains(column));
                }
                if (series.Items.Count > 0) series.SelectedIndex = 0;
            };
            rangeBox.Leave += delegate { populate(); };
            series.SelectedIndexChanged += delegate
            {
                if (series.SelectedIndex < 0 || series.SelectedIndex >= columns.Length) return;
                int column = columns[series.SelectedIndex];
                string value;
                seriesName.Text = names.TryGetValue(column, out value) ? value : "";
                Color color;
                colorButton.BackColor = colors.TryGetValue(column, out color) ? color : SystemColors.Control;
            };
            seriesName.TextChanged += delegate
            {
                if (series.SelectedIndex >= 0 && series.SelectedIndex < columns.Length)
                    names[columns[series.SelectedIndex]] = seriesName.Text;
            };
            colorButton.Click += delegate
            {
                if (series.SelectedIndex < 0 || series.SelectedIndex >= columns.Length) return;
                using (var dialog = new ColorDialog { FullOpen = true })
                    if (dialog.ShowDialog(this) == DialogResult.OK)
                    { colors[columns[series.SelectedIndex]] = dialog.Color; colorButton.BackColor = dialog.Color; }
            };
            Rectangle place = definition.Placement.IsEmpty ?
                new Rectangle(Math.Min(18, definition.Range.Right + 1), definition.Range.Top, 8, 14) :
                definition.Placement;
            var left = new NumericUpDown { Minimum = 1, Maximum = 26, Value = Math.Max(1, Math.Min(26, place.Left + 1)) };
            var top = new NumericUpDown { Minimum = 1, Maximum = 50000, Value = Math.Max(1, Math.Min(50000, place.Top + 1)) };
            var width = new NumericUpDown { Minimum = 2, Maximum = 26, Value = Math.Max(2, Math.Min(26, place.Width)) };
            var height = new NumericUpDown { Minimum = 3, Maximum = 100, Value = Math.Max(3, Math.Min(100, place.Height)) };
            DataField(layout, "Tiêu đề", title); DataField(layout, "Loại", kind);
            DataField(layout, "Vùng dữ liệu", rangeBox);
            DataField(layout, "Trục ngang", axisX); DataField(layout, "Trục dọc", axisY);
            DataField(layout, "Chú giải", legend); DataField(layout, "Nhãn", labels);
            DataField(layout, "Đường lưới", gridlines);
            DataField(layout, "Các series", series); DataField(layout, "Tên series", seriesName);
            DataField(layout, "Màu series", colorButton);
            DataField(layout, "Cột đặt", left); DataField(layout, "Hàng đặt", top);
            DataField(layout, "Rộng (cột)", width); DataField(layout, "Cao (hàng)", height);
            populate();
            using (var dialog = DataDialog("Chỉnh biểu đồ", layout))
            {
                dialog.Width = 520; dialog.Height = 720;
                if (dialog.ShowDialog(this) != DialogResult.OK) return false;
                Rectangle selected;
                if (!ParseChartRange(rangeBox.Text, out selected) || selected.Width < 2 || selected.Height < 2)
                { MessageBox.Show(this, "Vùng dữ liệu phải có ít nhất hai cột và hai hàng."); return false; }
                if (selected != definition.Range) populate();
                if (series.CheckedIndices.Count == 0)
                { MessageBox.Show(this, "Hãy chọn ít nhất một series."); return false; }
                if ((int)left.Value - 1 + (int)width.Value > ColumnCount)
                { MessageBox.Show(this, "Biểu đồ vượt quá cột Z. Hãy giảm cột đặt hoặc độ rộng."); return false; }
                definition.Title = title.Text.Trim(); definition.Kind = kind.Text;
                definition.Range = selected; definition.AxisTitleX = axisX.Text;
                definition.AxisTitleY = axisY.Text; definition.Legend = legend.Checked;
                definition.DataLabels = labels.Checked; definition.Gridlines = gridlines.Checked;
                definition.Placement = new Rectangle((int)left.Value - 1, (int)top.Value - 1,
                    (int)width.Value, (int)height.Value);
                EnsureRowCapacity(Math.Min(MaxRowCount, definition.Placement.Bottom));
                definition.SeriesColumns.Clear();
                foreach (int index in series.CheckedIndices) definition.SeriesColumns.Add(columns[index]);
                definition.SeriesNames.Clear(); definition.SeriesColors.Clear();
                foreach (int column in definition.SeriesColumns)
                {
                    if (names.ContainsKey(column)) definition.SeriesNames[column] = names[column];
                    if (colors.ContainsKey(column)) definition.SeriesColors[column] = colors[column];
                }
                RecordChange(); MarkDirty();
                return true;
            }
        }
    }
}
