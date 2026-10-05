using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms.DataVisualization.Charting;

namespace DinkCel
{
    internal static class ChartRendering
    {
        public static Chart Build(ChartDefinition definition, Func<int, int, string> read,
            Color background, Color foreground, Color accent)
        {
            var chart = new Chart { Dock = System.Windows.Forms.DockStyle.Fill,
                BackColor = background, ForeColor = foreground };
            var area = new ChartArea("Main") { BackColor = background };
            area.AxisX.Title = definition.AxisTitleX;
            area.AxisY.Title = definition.AxisTitleY;
            area.AxisX.LabelStyle.ForeColor = foreground;
            area.AxisY.LabelStyle.ForeColor = foreground;
            area.AxisX.TitleForeColor = foreground;
            area.AxisY.TitleForeColor = foreground;
            area.AxisX.MajorGrid.Enabled = definition.Gridlines;
            area.AxisY.MajorGrid.Enabled = definition.Gridlines;
            area.AxisX.MajorGrid.LineColor = Color.LightGray;
            area.AxisY.MajorGrid.LineColor = Color.LightGray;
            chart.ChartAreas.Add(area);
            chart.Legends.Add(new Legend("Legend") { Enabled = definition.Legend,
                BackColor = background, ForeColor = foreground });
            chart.Titles.Add(new Title(definition.Title) { ForeColor = foreground });
            Rectangle range = definition.Range;
            int[] columns = definition.SeriesColumns.Count > 0 ?
                definition.SeriesColumns.Where(c => c > range.Left && c < range.Right).Distinct().ToArray() :
                Enumerable.Range(range.Left + 1, Math.Max(0, range.Width - 1)).ToArray();
            if (definition.Kind == "Pie") columns = columns.Take(1).ToArray();
            Color[] palette = { accent, Color.Teal, Color.DarkOrange, Color.MediumPurple,
                Color.SteelBlue, Color.ForestGreen, Color.IndianRed };
            for (int i = 0; i < columns.Length; i++)
            {
                int column = columns[i];
                string name;
                if (!definition.SeriesNames.TryGetValue(column, out name) || string.IsNullOrWhiteSpace(name))
                    name = read(range.Top, column);
                if (string.IsNullOrWhiteSpace(name)) name = "Series " + (i + 1);
                string unique = name;
                for (int suffix = 2; chart.Series.Any(s => s.Name == unique); suffix++)
                    unique = name + " " + suffix;
                var series = new Series(unique) { ChartType = Type(definition.Kind, i),
                    Color = definition.SeriesColors.ContainsKey(column) ? definition.SeriesColors[column] :
                        palette[i % palette.Length], IsValueShownAsLabel = definition.DataLabels,
                    BorderWidth = definition.Kind == "Line" || definition.Kind == "Combo" ? 3 : 1 };
                for (int row = range.Top + 1; row < range.Bottom; row++)
                {
                    double y;
                    if (!TryNumber(read(row, column), out y)) continue;
                    string label = read(row, range.Left) ?? "";
                    if (definition.Kind == "Scatter")
                    {
                        double x;
                        if (TryNumber(label, out x)) series.Points.AddXY(x, y);
                    }
                    else series.Points.AddXY(label, y);
                }
                if (definition.Kind == "Pie")
                    for (int point = 0; point < series.Points.Count; point++)
                        series.Points[point].Color = palette[point % palette.Length];
                chart.Series.Add(series);
            }
            return chart;
        }

        private static bool TryNumber(string raw, out double value)
        {
            return double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out value) ||
                double.TryParse(raw, NumberStyles.Any, CultureInfo.CurrentCulture, out value);
        }

        private static SeriesChartType Type(string kind, int index)
        {
            switch (kind)
            {
                case "Line": return SeriesChartType.Line;
                case "Pie": return SeriesChartType.Pie;
                case "Bar": return SeriesChartType.Bar;
                case "Area": return SeriesChartType.Area;
                case "Scatter": return SeriesChartType.Point;
                case "Stacked": return SeriesChartType.StackedColumn;
                case "100% Stacked": return SeriesChartType.StackedColumn100;
                case "Combo": return index % 2 == 0 ? SeriesChartType.Column : SeriesChartType.Line;
                default: return SeriesChartType.Column;
            }
        }
    }
}
