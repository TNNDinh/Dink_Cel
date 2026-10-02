using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace DinkCel
{
    internal static class XlsxCharts
    {
        public static readonly XNamespace Chart = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        public static readonly XNamespace Drawing = "http://schemas.openxmlformats.org/drawingml/2006/main";
        public static readonly XNamespace SpreadsheetDrawing = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        public static readonly XNamespace Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace PackageRel = "http://schemas.openxmlformats.org/package/2006/relationships";

        private static XElement Value(string name, object value)
        {
            return new XElement(Chart + name, new XAttribute("val", value));
        }

        private static string ColumnName(int column)
        {
            string text = "";
            do { text = (char)('A' + column % 26) + text; column = column / 26 - 1; } while (column >= 0);
            return text;
        }

        private static string SheetReference(string name)
        {
            return "'" + name.Replace("'", "''") + "'!";
        }

        public static XDocument BuildChart(ChartDefinition definition, SheetSnapshot sheet,
            int chartNumber, int columns)
        {
            Rectangle range = definition.Range;
            string prefix = SheetReference(sheet.Name);
            var plot = new XElement(Chart + "plotArea", new XElement(Chart + "layout"));
            bool pie = definition.Kind == "Pie", line = definition.Kind == "Line";
            var graph = new XElement(Chart + (pie ? "pieChart" : line ? "lineChart" : "barChart"));
            if (pie) graph.Add(Value("varyColors", 1));
            else if (line) graph.Add(Value("grouping", "standard"));
            else { graph.Add(Value("barDir", "col")); graph.Add(Value("grouping", "clustered")); }
            int seriesIndex = 0;
            for (int column = range.Left + 1; column < range.Right; column++)
            {
                if (pie && seriesIndex > 0) break;
                var series = new XElement(Chart + "ser", Value("idx", seriesIndex), Value("order", seriesIndex));
                string seriesRef = prefix + "$" + ColumnName(column) + "$" + (range.Top + 1);
                series.Add(new XElement(Chart + "tx", new XElement(Chart + "strRef",
                    new XElement(Chart + "f", seriesRef))));
                var categories = new XElement(Chart + "strCache", Value("ptCount", Math.Max(0, range.Height - 1)));
                var values = new XElement(Chart + "numCache", new XElement(Chart + "formatCode", "General"),
                    Value("ptCount", Math.Max(0, range.Height - 1)));
                for (int row = range.Top + 1; row < range.Bottom; row++)
                {
                    CellSnapshot labelCell, valueCell;
                    string label = sheet.Cells.TryGetValue(row * columns + range.Left, out labelCell) ? labelCell.Text : "";
                    string raw = sheet.Cells.TryGetValue(row * columns + column, out valueCell) ? valueCell.Text : "";
                    categories.Add(new XElement(Chart + "pt", new XAttribute("idx", row - range.Top - 1),
                        new XElement(Chart + "v", label ?? "")));
                    double number;
                    if (double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
                        values.Add(new XElement(Chart + "pt", new XAttribute("idx", row - range.Top - 1),
                            new XElement(Chart + "v", number.ToString(CultureInfo.InvariantCulture))));
                }
                string categoryRef = prefix + "$" + ColumnName(range.Left) + "$" + (range.Top + 2) +
                    ":$" + ColumnName(range.Left) + "$" + range.Bottom;
                string valueRef = prefix + "$" + ColumnName(column) + "$" + (range.Top + 2) +
                    ":$" + ColumnName(column) + "$" + range.Bottom;
                series.Add(new XElement(Chart + "cat", new XElement(Chart + "strRef",
                    new XElement(Chart + "f", categoryRef), categories)));
                series.Add(new XElement(Chart + "val", new XElement(Chart + "numRef",
                    new XElement(Chart + "f", valueRef), values)));
                graph.Add(series);
                seriesIndex++;
            }
            if (!pie)
            {
                graph.Add(Value("axId", 10 + chartNumber * 2));
                graph.Add(Value("axId", 11 + chartNumber * 2));
            }
            plot.Add(graph);
            if (!pie)
            {
                int categoryAxis = 10 + chartNumber * 2, valueAxis = 11 + chartNumber * 2;
                plot.Add(new XElement(Chart + "catAx", Value("axId", categoryAxis),
                    new XElement(Chart + "scaling", Value("orientation", "minMax")),
                    Value("axPos", "b"), Value("crossAx", valueAxis), Value("crosses", "autoZero")));
                plot.Add(new XElement(Chart + "valAx", Value("axId", valueAxis),
                    new XElement(Chart + "scaling", Value("orientation", "minMax")),
                    Value("axPos", "l"), Value("crossAx", categoryAxis), Value("crosses", "autoZero")));
            }
            var title = new XElement(Chart + "title",
                new XElement(Chart + "tx", new XElement(Chart + "rich",
                    new XElement(Drawing + "bodyPr"), new XElement(Drawing + "lstStyle"),
                    new XElement(Drawing + "p", new XElement(Drawing + "r",
                        new XElement(Drawing + "rPr", new XAttribute("lang", "vi-VN")),
                        new XElement(Drawing + "t", definition.Title))))),
                Value("overlay", 0));
            var chart = new XElement(Chart + "chart", title, plot,
                new XElement(Chart + "legend", Value("legendPos", "r"), new XElement(Chart + "layout")),
                Value("plotVisOnly", 1));
            return new XDocument(new XElement(Chart + "chartSpace",
                new XAttribute(XNamespace.Xmlns + "a", Drawing), chart));
        }

        private static XElement Marker(string kind, int column, int row)
        {
            return new XElement(SpreadsheetDrawing + kind,
                new XElement(SpreadsheetDrawing + "col", column),
                new XElement(SpreadsheetDrawing + "colOff", 0),
                new XElement(SpreadsheetDrawing + "row", row),
                new XElement(SpreadsheetDrawing + "rowOff", 0));
        }

        public static XDocument BuildDrawing(IList<ChartDefinition> definitions)
        {
            var root = new XElement(SpreadsheetDrawing + "wsDr",
                new XAttribute(XNamespace.Xmlns + "a", Drawing),
                new XAttribute(XNamespace.Xmlns + "c", Chart),
                new XAttribute(XNamespace.Xmlns + "r", Rel));
            for (int i = 0; i < definitions.Count; i++)
            {
                Rectangle range = definitions[i].Range;
                int left = Math.Min(20, range.Right + 1), top = range.Top + i * 15;
                var frame = new XElement(SpreadsheetDrawing + "graphicFrame", new XAttribute("macro", ""),
                    new XElement(SpreadsheetDrawing + "nvGraphicFramePr",
                        new XElement(SpreadsheetDrawing + "cNvPr", new XAttribute("id", i + 2),
                            new XAttribute("name", "DinkCel Chart " + (i + 1))),
                        new XElement(SpreadsheetDrawing + "cNvGraphicFramePr")),
                    new XElement(SpreadsheetDrawing + "xfrm",
                        new XElement(Drawing + "off", new XAttribute("x", 0), new XAttribute("y", 0)),
                        new XElement(Drawing + "ext", new XAttribute("cx", 0), new XAttribute("cy", 0))),
                    new XElement(Drawing + "graphic",
                        new XElement(Drawing + "graphicData",
                            new XAttribute("uri", Chart.NamespaceName),
                            new XElement(Chart + "chart", new XAttribute(Rel + "id", "rIdChart" + (i + 1))))));
                root.Add(new XElement(SpreadsheetDrawing + "twoCellAnchor",
                    Marker("from", left, top), Marker("to", left + 8, top + 14), frame,
                    new XElement(SpreadsheetDrawing + "clientData")));
            }
            return new XDocument(root);
        }

        public static XDocument BuildDrawingRelationships(IList<int> chartNumbers)
        {
            var root = new XElement(PackageRel + "Relationships");
            for (int i = 0; i < chartNumbers.Count; i++)
                root.Add(new XElement(PackageRel + "Relationship",
                    new XAttribute("Id", "rIdChart" + (i + 1)),
                    new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"),
                    new XAttribute("Target", "../charts/chart" + chartNumbers[i] + ".xml")));
            return new XDocument(root);
        }

        public static ChartDefinition ParseChart(XDocument document, int columns, int rows)
        {
            XElement plot = document.Descendants(Chart + "plotArea").FirstOrDefault();
            if (plot == null) return null;
            XElement graph = plot.Element(Chart + "barChart") ?? plot.Element(Chart + "lineChart") ??
                plot.Element(Chart + "pieChart");
            if (graph == null) return null;
            XElement series = graph.Elements(Chart + "ser").FirstOrDefault();
            if (series == null) return null;
            XElement catElement = series.Element(Chart + "cat");
            XElement valElement = series.Element(Chart + "val");
            string category = catElement == null ? null :
                (string)catElement.Descendants(Chart + "f").FirstOrDefault();
            string value = valElement == null ? null :
                (string)valElement.Descendants(Chart + "f").FirstOrDefault();
            Rectangle catRange = ParseFormulaRange(category, columns, rows);
            Rectangle valRange = ParseFormulaRange(value, columns, rows);
            if (catRange.IsEmpty || valRange.IsEmpty) return null;
            int left = catRange.Left, top = Math.Max(0, catRange.Top - 1);
            int right = valRange.Right, bottom = Math.Max(catRange.Bottom, valRange.Bottom);
            string title = string.Concat(document.Descendants(Chart + "title").Descendants(Drawing + "t")
                .Select(x => x.Value));
            return new ChartDefinition { Title = string.IsNullOrEmpty(title) ? "Chart" : title,
                Kind = graph.Name == Chart + "lineChart" ? "Line" : graph.Name == Chart + "pieChart" ? "Pie" : "Column",
                Range = new Rectangle(left, top, right - left, bottom - top) };
        }

        private static Rectangle ParseFormulaRange(string formula, int columns, int rows)
        {
            if (string.IsNullOrEmpty(formula)) return Rectangle.Empty;
            int bang = formula.LastIndexOf('!');
            string[] refs = (bang >= 0 ? formula.Substring(bang + 1) : formula).Replace("$", "").Split(':');
            int x1, y1, x2, y2;
            if (!ParseAddress(refs[0], out x1, out y1)) return Rectangle.Empty;
            if (refs.Length > 1)
            { if (!ParseAddress(refs[1], out x2, out y2)) return Rectangle.Empty; }
            else { x2 = x1; y2 = y1; }
            return x1 < 0 || y1 < 0 || x2 >= columns || y2 >= rows ? Rectangle.Empty :
                new Rectangle(x1, y1, x2 - x1 + 1, y2 - y1 + 1);
        }

        private static bool ParseAddress(string text, out int column, out int row)
        {
            column = row = -1;
            int i = 0;
            while (i < text.Length && char.IsLetter(text[i])) i++;
            if (i == 0 || i == text.Length || !int.TryParse(text.Substring(i), out row)) return false;
            column = 0;
            foreach (char c in text.Substring(0, i).ToUpperInvariant()) column = column * 26 + c - 'A' + 1;
            column--; row--;
            return true;
        }
    }
}
