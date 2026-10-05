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
                Rectangle placement = definitions[i].Placement.IsEmpty ?
                    new Rectangle(Math.Min(20, range.Right + 1), range.Top + i * 15, 8, 14) :
                    definitions[i].Placement;
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
                    Marker("from", placement.Left, placement.Top),
                    Marker("to", placement.Right, placement.Bottom), frame,
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
