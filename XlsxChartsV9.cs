using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace DinkCel
{
    internal static partial class XlsxCharts
    {
        private static XElement RichTitle(string value)
        {
            return new XElement(Chart + "title", new XElement(Chart + "tx",
                new XElement(Chart + "rich", new XElement(Drawing + "bodyPr"),
                    new XElement(Drawing + "lstStyle"),
                    new XElement(Drawing + "p", new XElement(Drawing + "r",
                        new XElement(Drawing + "rPr", new XAttribute("lang", "vi-VN")),
                        new XElement(Drawing + "t", value ?? ""))))), Value("overlay", 0));
        }

        private static string Cell(SheetSnapshot sheet, int row, int col, int columns)
        {
            CellSnapshot cell;
            return sheet.Cells.TryGetValue(row * columns + col, out cell) ? cell.Text ?? "" : "";
        }

        private static XElement Series(ChartDefinition definition, SheetSnapshot sheet,
            int columns, int column, int index, bool scatter)
        {
            Rectangle range = definition.Range;
            string prefix = SheetReference(sheet.Name);
            var series = new XElement(Chart + "ser", Value("idx", index), Value("order", index));
            string name;
            if (definition.SeriesNames.TryGetValue(column, out name) && !string.IsNullOrWhiteSpace(name))
                series.Add(new XElement(Chart + "tx", new XElement(Chart + "v", name)));
            else
                series.Add(new XElement(Chart + "tx", new XElement(Chart + "strRef",
                    new XElement(Chart + "f", prefix + "$" + ColumnName(column) + "$" + (range.Top + 1)))));
            Color color;
            if (definition.SeriesColors.TryGetValue(column, out color))
                series.Add(new XElement(Chart + "spPr", new XElement(Drawing + "solidFill",
                    new XElement(Drawing + "srgbClr", new XAttribute("val",
                        color.R.ToString("X2") + color.G.ToString("X2") + color.B.ToString("X2"))))));
            int count = Math.Max(0, range.Height - 1);
            string first = prefix + "$" + ColumnName(range.Left) + "$" + (range.Top + 2) +
                ":$" + ColumnName(range.Left) + "$" + range.Bottom;
            string second = prefix + "$" + ColumnName(column) + "$" + (range.Top + 2) +
                ":$" + ColumnName(column) + "$" + range.Bottom;
            var categories = new XElement(Chart + (scatter ? "numCache" : "strCache"));
            if (scatter) categories.Add(new XElement(Chart + "formatCode", "General"));
            categories.Add(Value("ptCount", count));
            var values = new XElement(Chart + "numCache",
                new XElement(Chart + "formatCode", "General"), Value("ptCount", count));
            for (int row = range.Top + 1; row < range.Bottom; row++)
            {
                string label = Cell(sheet, row, range.Left, columns);
                double x, y;
                if (!scatter || double.TryParse(label, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out x))
                    categories.Add(new XElement(Chart + "pt", new XAttribute("idx", row - range.Top - 1),
                        new XElement(Chart + "v", label)));
                if (double.TryParse(Cell(sheet, row, column, columns), NumberStyles.Float,
                    CultureInfo.InvariantCulture, out y))
                    values.Add(new XElement(Chart + "pt", new XAttribute("idx", row - range.Top - 1),
                        new XElement(Chart + "v", y.ToString(CultureInfo.InvariantCulture))));
            }
            series.Add(new XElement(Chart + (scatter ? "xVal" : "cat"),
                new XElement(Chart + (scatter ? "numRef" : "strRef"),
                    new XElement(Chart + "f", first), categories)));
            series.Add(new XElement(Chart + (scatter ? "yVal" : "val"),
                new XElement(Chart + "numRef", new XElement(Chart + "f", second), values)));
            return series;
        }

        private static XElement Graph(string kind, IEnumerable<XElement> series, int xAxis, int yAxis,
            bool labels)
        {
            bool pie = kind == "Pie", scatter = kind == "Scatter";
            string name = pie ? "pieChart" : scatter ? "scatterChart" :
                kind == "Line" ? "lineChart" : kind == "Area" ? "areaChart" : "barChart";
            var graph = new XElement(Chart + name);
            if (pie) graph.Add(Value("varyColors", 1));
            else if (scatter) graph.Add(Value("scatterStyle", "marker"));
            else if (name == "barChart")
            {
                graph.Add(Value("barDir", kind == "Bar" ? "bar" : "col"));
                graph.Add(Value("grouping", kind == "Stacked" ? "stacked" :
                    kind == "100% Stacked" ? "percentStacked" : "clustered"));
            }
            else graph.Add(Value("grouping", "standard"));
            foreach (XElement entry in series) graph.Add(entry);
            if (labels) graph.Add(new XElement(Chart + "dLbls", Value("showLegendKey", 0),
                Value("showVal", 1), Value("showCatName", 0)));
            if (!pie) { graph.Add(Value("axId", xAxis)); graph.Add(Value("axId", yAxis)); }
            return graph;
        }

        public static XDocument BuildChart(ChartDefinition definition, SheetSnapshot sheet,
            int chartNumber, int columns)
        {
            Rectangle range = definition.Range;
            int[] seriesColumns = definition.SeriesColumns.Count == 0 ?
                Enumerable.Range(range.Left + 1, Math.Max(0, range.Width - 1)).ToArray() :
                definition.SeriesColumns.Where(c => c > range.Left && c < range.Right).Distinct().ToArray();
            if (definition.Kind == "Pie") seriesColumns = seriesColumns.Take(1).ToArray();
            int xAxis = 10 + chartNumber * 2, yAxis = xAxis + 1;
            var plot = new XElement(Chart + "plotArea", new XElement(Chart + "layout"));
            bool scatter = definition.Kind == "Scatter";
            if (definition.Kind == "Combo" && seriesColumns.Length > 1)
            {
                plot.Add(Graph("Column", new[] { Series(definition, sheet, columns,
                    seriesColumns[0], 0, false) }, xAxis, yAxis, definition.DataLabels));
                plot.Add(Graph("Line", seriesColumns.Skip(1).Select((c, i) => Series(definition,
                    sheet, columns, c, i + 1, false)), xAxis, yAxis, definition.DataLabels));
            }
            else plot.Add(Graph(definition.Kind == "Combo" ? "Column" : definition.Kind,
                seriesColumns.Select((c, i) => Series(definition, sheet, columns, c, i, scatter)),
                xAxis, yAxis, definition.DataLabels));
            if (definition.Kind != "Pie")
            {
                var axisX = new XElement(Chart + (scatter ? "valAx" : "catAx"), Value("axId", xAxis),
                    new XElement(Chart + "scaling", Value("orientation", "minMax")),
                    Value("axPos", "b"));
                var axisY = new XElement(Chart + "valAx", Value("axId", yAxis),
                    new XElement(Chart + "scaling", Value("orientation", "minMax")),
                    Value("axPos", "l"));
                if (!string.IsNullOrEmpty(definition.AxisTitleX)) axisX.Add(RichTitle(definition.AxisTitleX));
                if (definition.Gridlines) axisY.Add(new XElement(Chart + "majorGridlines"));
                if (!string.IsNullOrEmpty(definition.AxisTitleY)) axisY.Add(RichTitle(definition.AxisTitleY));
                axisX.Add(Value("crossAx", yAxis), Value("crosses", "autoZero"));
                axisY.Add(Value("crossAx", xAxis), Value("crosses", "autoZero"));
                plot.Add(axisX, axisY);
            }
            var chart = new XElement(Chart + "chart", RichTitle(definition.Title), plot);
            if (definition.Legend) chart.Add(new XElement(Chart + "legend", Value("legendPos", "r"),
                new XElement(Chart + "layout")));
            chart.Add(Value("plotVisOnly", 1));
            return new XDocument(new XElement(Chart + "chartSpace",
                new XAttribute(XNamespace.Xmlns + "a", Drawing), chart));
        }

        public static ChartDefinition ParseChart(XDocument document, int columns, int rows)
        {
            XElement plot = document.Descendants(Chart + "plotArea").FirstOrDefault();
            if (plot == null) return null;
            XElement graph = plot.Elements().FirstOrDefault(x => x.Name == Chart + "barChart" ||
                x.Name == Chart + "lineChart" || x.Name == Chart + "pieChart" ||
                x.Name == Chart + "areaChart" || x.Name == Chart + "scatterChart");
            if (graph == null) return null;
            XElement firstSeries = graph.Elements(Chart + "ser").FirstOrDefault();
            if (firstSeries == null) return null;
            XElement categoryElement = firstSeries.Element(Chart + "cat") ?? firstSeries.Element(Chart + "xVal");
            XElement valueElement = firstSeries.Element(Chart + "val") ?? firstSeries.Element(Chart + "yVal");
            string category = categoryElement == null ? null :
                (string)categoryElement.Descendants(Chart + "f").FirstOrDefault();
            string value = valueElement == null ? null :
                (string)valueElement.Descendants(Chart + "f").FirstOrDefault();
            Rectangle categoryRange = ParseFormulaRange(category, columns, rows);
            Rectangle valueRange = ParseFormulaRange(value, columns, rows);
            if (categoryRange.IsEmpty || valueRange.IsEmpty) return null;
            string direction = (string)(graph.Element(Chart + "barDir") == null ? null :
                graph.Element(Chart + "barDir").Attribute("val"));
            string grouping = (string)(graph.Element(Chart + "grouping") == null ? null :
                graph.Element(Chart + "grouping").Attribute("val"));
            string kind = graph.Name == Chart + "pieChart" ? "Pie" :
                graph.Name == Chart + "lineChart" ? "Line" :
                graph.Name == Chart + "areaChart" ? "Area" :
                graph.Name == Chart + "scatterChart" ? "Scatter" :
                direction == "bar" ? "Bar" :
                grouping == "stacked" ? "Stacked" :
                grouping == "percentStacked" ?
                    "100% Stacked" : "Column";
            if (plot.Elements(Chart + "barChart").Any() && plot.Elements(Chart + "lineChart").Any())
                kind = "Combo";
            XElement chartElement = document.Descendants(Chart + "chart").FirstOrDefault();
            XElement title = chartElement == null ? null : chartElement.Element(Chart + "title");
            var definition = new ChartDefinition { Kind = kind,
                Title = title == null ? "Chart" : string.Concat(title.Descendants(Drawing + "t").Select(x => x.Value)),
                Range = Rectangle.FromLTRB(categoryRange.Left, Math.Max(0, categoryRange.Top - 1),
                    valueRange.Right, Math.Max(categoryRange.Bottom, valueRange.Bottom)),
                Legend = chartElement != null && chartElement.Element(Chart + "legend") != null,
                DataLabels = plot.Descendants(Chart + "dLbls").Any(),
                Gridlines = plot.Descendants(Chart + "majorGridlines").Any() };
            XElement xAxis = plot.Element(Chart + "catAx") ?? plot.Elements(Chart + "valAx").FirstOrDefault();
            XElement yAxis = plot.Elements(Chart + "valAx").LastOrDefault();
            definition.AxisTitleX = xAxis == null ? "" : string.Concat(xAxis.Elements(Chart + "title")
                .Descendants(Drawing + "t").Select(x => x.Value));
            definition.AxisTitleY = yAxis == null ? "" : string.Concat(yAxis.Elements(Chart + "title")
                .Descendants(Drawing + "t").Select(x => x.Value));
            foreach (XElement entry in plot.Elements().SelectMany(x => x.Elements(Chart + "ser")))
            {
                XElement valueReference = entry.Element(Chart + "val") ?? entry.Element(Chart + "yVal");
                string formula = valueReference == null ? null :
                    (string)valueReference.Descendants(Chart + "f").FirstOrDefault();
                Rectangle valueRef = ParseFormulaRange(formula, columns, rows);
                if (valueRef.IsEmpty || definition.SeriesColumns.Contains(valueRef.Left)) continue;
                int column = valueRef.Left;
                definition.SeriesColumns.Add(column);
                XElement text = entry.Element(Chart + "tx");
                string name = text == null ? null : (string)text.Element(Chart + "v");
                if (name != null) definition.SeriesNames[column] = name;
                XElement rgbElement = entry.Descendants(Drawing + "srgbClr").FirstOrDefault();
                string rgb = rgbElement == null ? null : (string)rgbElement.Attribute("val");
                if (!string.IsNullOrEmpty(rgb) && rgb.Length == 6)
                    definition.SeriesColors[column] = ColorTranslator.FromHtml("#" + rgb);
            }
            if (definition.SeriesColumns.Count > 0)
                definition.Range = Rectangle.FromLTRB(definition.Range.Left, definition.Range.Top,
                    Math.Max(definition.Range.Right, definition.SeriesColumns.Max() + 1),
                    definition.Range.Bottom);
            return definition;
        }
    }
}
