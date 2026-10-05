using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;

namespace DinkCel
{
    internal sealed class PrintPagePlan
    {
        public SheetSnapshot Sheet;
        public Rectangle Area;
        public int FirstRow, LastRow, FirstColumn, LastColumn;
        public int TitleRows;
        public double Scale;
        public int Number, Total;
        public ChartDefinition Chart;
    }

    internal static class PrintLayout
    {
        public static Size PaperSize(PrintSettings settings)
        {
            Size portrait = settings.Paper == "Letter" ? new Size(850, 1100) :
                settings.Paper == "Legal" ? new Size(850, 1400) :
                settings.Paper == "A3" ? new Size(1169, 1654) : new Size(827, 1169);
            return settings.Landscape ? new Size(portrait.Height, portrait.Width) : portrait;
        }

        public static Rectangle UsedArea(SheetSnapshot sheet, int columns, int rows)
        {
            if (!sheet.Print.PrintArea.IsEmpty) return Rectangle.Intersect(
                new Rectangle(0, 0, columns, rows), sheet.Print.PrintArea);
            if (sheet.Cells.Count == 0) return new Rectangle(0, 0, 1, 1);
            int lastRow = sheet.Cells.Keys.Max() / columns;
            int lastColumn = sheet.Cells.Keys.Max(key => key % columns);
            return new Rectangle(0, 0, lastColumn + 1, lastRow + 1);
        }

        private static double Width(SheetSnapshot sheet, int column)
        {
            int value;
            return (sheet.ColumnWidths.TryGetValue(column, out value) ? value : 120) * 100.0 / 96;
        }

        private static double Height(SheetSnapshot sheet, int row)
        {
            int value;
            return (sheet.RowHeights.TryGetValue(row, out value) ? value : 27) * 100.0 / 96;
        }

        public static List<PrintPagePlan> Plan(SheetSnapshot sheet, int rows, int columns,
            Rectangle selection, bool includeCharts)
        {
            var result = new List<PrintPagePlan>();
            PrintSettings settings = sheet.Print;
            Rectangle area = selection.IsEmpty ? UsedArea(sheet, columns, rows) :
                Rectangle.Intersect(selection, new Rectangle(0, 0, columns, rows));
            Size paper = PaperSize(settings);
            double availableWidth = Math.Max(100, paper.Width - settings.MarginLeft - settings.MarginRight);
            double availableHeight = Math.Max(100, paper.Height - settings.MarginTop - settings.MarginBottom - 50);
            if (!area.IsEmpty)
            {
                int titleRows = Math.Min(settings.TitleRows, area.Top == 0 ?
                    Math.Max(0, area.Height - 1) : rows);
                double titleHeight = Enumerable.Range(0, titleRows).Sum(r => Height(sheet, r));
                double scale = Math.Max(.1, Math.Min(4, settings.Scale / 100.0));
                if (settings.FitToOnePage)
                {
                    double totalWidth = Enumerable.Range(area.Left, area.Width).Sum(c => Width(sheet, c));
                    double totalHeight = Enumerable.Range(area.Top, area.Height).Sum(r => Height(sheet, r)) +
                        (area.Top == 0 ? 0 : titleHeight);
                    scale = Math.Min(1, Math.Min(availableWidth / Math.Max(1, totalWidth),
                        availableHeight / Math.Max(1, totalHeight)));
                }
                var columnSlices = new List<Tuple<int, int>>();
                for (int column = area.Left; column < area.Right;)
                {
                    int start = column; double used = 0;
                    do { used += Width(sheet, column++) * scale; }
                    while (column < area.Right && used + Width(sheet, column) * scale <= availableWidth);
                    columnSlices.Add(Tuple.Create(start, column));
                }
                var rowSlices = new List<Tuple<int, int>>();
                int firstData = area.Top == 0 ? area.Top + titleRows : area.Top;
                if (firstData >= area.Bottom) firstData = area.Top;
                for (int row = firstData; row < area.Bottom;)
                {
                    int start = row;
                    double used = titleHeight * scale;
                    do { used += Height(sheet, row++) * scale; }
                    while (row < area.Bottom &&
                        (settings.FitToOnePage || !settings.PageBreakRows.Contains(row)) &&
                        used + Height(sheet, row) * scale <= availableHeight);
                    rowSlices.Add(Tuple.Create(start, row));
                }
                foreach (var cols in columnSlices)
                    foreach (var lines in rowSlices)
                        result.Add(new PrintPagePlan { Sheet = sheet, Area = area,
                            FirstColumn = cols.Item1, LastColumn = cols.Item2,
                            FirstRow = lines.Item1, LastRow = lines.Item2,
                            TitleRows = titleRows, Scale = scale });
            }
            if (includeCharts && settings.PrintCharts)
                foreach (ChartDefinition chart in sheet.Charts)
                    result.Add(new PrintPagePlan { Sheet = sheet, Chart = chart, Scale = 1 });
            for (int i = 0; i < result.Count; i++)
            { result[i].Number = i + 1; result[i].Total = result.Count; }
            return result;
        }

        public static string HeaderFooter(string text, PrintPagePlan page)
        {
            return (text ?? "").Replace("&P", page.Number.ToString())
                .Replace("&N", page.Total.ToString()).Replace("&F", page.Sheet.Name)
                .Replace("&D", DateTime.Now.ToShortDateString());
        }
    }
}
