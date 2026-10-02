using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DinkCel
{
    internal sealed class XlsxStyles
    {
        private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private readonly Dictionary<string, int> styleIds = new Dictionary<string, int>();
        private readonly List<CellSnapshot> styles = new List<CellSnapshot>();
        private readonly Dictionary<string, int> fonts = new Dictionary<string, int>();
        private readonly Dictionary<string, int> fills = new Dictionary<string, int>();
        private readonly Dictionary<string, int> formats = new Dictionary<string, int>();
        private readonly Dictionary<string, int> differentialColors = new Dictionary<string, int>();

        private static string ColorKey(Color color) { return color.IsEmpty ? "" : color.ToArgb().ToString("X8", CultureInfo.InvariantCulture); }
        private static string FontKey(CellSnapshot cell)
        {
            return cell.HasFont ? cell.FontSize.ToString(CultureInfo.InvariantCulture) + ":" + (int)cell.FontStyle + ":" + ColorKey(cell.ForeColor)
                : "10:0:" + ColorKey(cell.ForeColor);
        }
        private static string Key(CellSnapshot cell)
        {
            return FontKey(cell) + "|" + ColorKey(cell.BackColor) + "|" + cell.Alignment + "|" + cell.NumberFormat;
        }

        public XlsxStyles(WorkbookSnapshot workbook)
        {
            Add(new CellSnapshot());
            foreach (SheetSnapshot sheet in workbook.Sheets)
                foreach (CellSnapshot cell in sheet.Cells.Values) Add(cell);
            foreach (SheetSnapshot sheet in workbook.Sheets)
                foreach (ConditionalRule rule in sheet.Rules)
                {
                    string color = ColorKey(rule.Color);
                    if (!differentialColors.ContainsKey(color)) differentialColors[color] = differentialColors.Count;
                }
            foreach (CellSnapshot cell in styles)
            {
                string font = FontKey(cell);
                if (!fonts.ContainsKey(font)) fonts[font] = fonts.Count;
                string fill = ColorKey(cell.BackColor);
                if (fill.Length > 0 && !fills.ContainsKey(fill)) fills[fill] = fills.Count + 2;
                if (!string.IsNullOrEmpty(cell.NumberFormat) && !formats.ContainsKey(cell.NumberFormat))
                    formats[cell.NumberFormat] = formats.Count + 164;
            }
        }

        private void Add(CellSnapshot cell)
        {
            string key = Key(cell);
            if (!styleIds.ContainsKey(key)) { styleIds[key] = styles.Count; styles.Add(cell); }
        }
        public int Index(CellSnapshot cell) { return styleIds[Key(cell)]; }
        public int DifferentialIndex(Color color) { return differentialColors[ColorKey(color)]; }

        private static XElement ColorElement(Color color)
        {
            return new XElement(S + "color", new XAttribute("rgb", ColorKey(color)));
        }

        public XDocument Document()
        {
            var numFmts = new XElement(S + "numFmts", new XAttribute("count", formats.Count));
            foreach (var item in formats.OrderBy(x => x.Value)) numFmts.Add(new XElement(S + "numFmt", new XAttribute("numFmtId", item.Value), new XAttribute("formatCode", item.Key)));
            var fontElements = new XElement(S + "fonts", new XAttribute("count", fonts.Count));
            foreach (var key in fonts.OrderBy(x => x.Value).Select(x => x.Key))
            {
                CellSnapshot style = styles.First(x => FontKey(x) == key);
                var font = new XElement(S + "font", new XElement(S + "sz", new XAttribute("val", style.HasFont ? style.FontSize : 10)),
                    new XElement(S + "name", new XAttribute("val", "Arial")));
                if ((style.FontStyle & FontStyle.Bold) != 0) font.Add(new XElement(S + "b"));
                if ((style.FontStyle & FontStyle.Italic) != 0) font.Add(new XElement(S + "i"));
                if ((style.FontStyle & FontStyle.Underline) != 0) font.Add(new XElement(S + "u"));
                if (!style.ForeColor.IsEmpty) font.Add(ColorElement(style.ForeColor));
                fontElements.Add(font);
            }
            var fillElements = new XElement(S + "fills", new XAttribute("count", fills.Count + 2),
                new XElement(S + "fill", new XElement(S + "patternFill", new XAttribute("patternType", "none"))),
                new XElement(S + "fill", new XElement(S + "patternFill", new XAttribute("patternType", "gray125"))));
            foreach (var key in fills.OrderBy(x => x.Value).Select(x => x.Key))
            {
                Color color = Color.FromArgb(unchecked((int)uint.Parse(key, NumberStyles.HexNumber, CultureInfo.InvariantCulture)));
                fillElements.Add(new XElement(S + "fill", new XElement(S + "patternFill", new XAttribute("patternType", "solid"),
                    new XElement(S + "fgColor", new XAttribute("rgb", ColorKey(color))),
                    new XElement(S + "bgColor", new XAttribute("indexed", 64)))));
            }
            var xfs = new XElement(S + "cellXfs", new XAttribute("count", styles.Count));
            foreach (CellSnapshot cell in styles)
            {
                int format = string.IsNullOrEmpty(cell.NumberFormat) ? 0 : formats[cell.NumberFormat];
                int fill = cell.BackColor.IsEmpty ? 0 : fills[ColorKey(cell.BackColor)];
                var xf = new XElement(S + "xf", new XAttribute("numFmtId", format),
                    new XAttribute("fontId", fonts[FontKey(cell)]), new XAttribute("fillId", fill),
                    new XAttribute("borderId", 0), new XAttribute("xfId", 0));
                if (format != 0) xf.SetAttributeValue("applyNumberFormat", 1);
                if (fill != 0) xf.SetAttributeValue("applyFill", 1);
                if (cell.Alignment != DataGridViewContentAlignment.NotSet)
                {
                    string alignment = cell.Alignment == DataGridViewContentAlignment.MiddleCenter ? "center" :
                        cell.Alignment == DataGridViewContentAlignment.MiddleRight ? "right" : "left";
                    xf.Add(new XElement(S + "alignment", new XAttribute("horizontal", alignment), new XAttribute("vertical", "center")));
                    xf.SetAttributeValue("applyAlignment", 1);
                }
                xfs.Add(xf);
            }
            var dxfs = new XElement(S + "dxfs", new XAttribute("count", differentialColors.Count));
            foreach (var item in differentialColors.OrderBy(x => x.Value))
                dxfs.Add(new XElement(S + "dxf", new XElement(S + "fill", new XElement(S + "patternFill",
                    new XAttribute("patternType", "solid"), new XElement(S + "fgColor", new XAttribute("rgb", item.Key)),
                    new XElement(S + "bgColor", new XAttribute("indexed", 64))))));
            return new XDocument(new XElement(S + "styleSheet", numFmts, fontElements, fillElements,
                new XElement(S + "borders", new XAttribute("count", 1), new XElement(S + "border")),
                new XElement(S + "cellStyleXfs", new XAttribute("count", 1), new XElement(S + "xf", new XAttribute("numFmtId", 0), new XAttribute("fontId", 0), new XAttribute("fillId", 0), new XAttribute("borderId", 0))),
                xfs, new XElement(S + "cellStyles", new XAttribute("count", 1),
                    new XElement(S + "cellStyle", new XAttribute("name", "Normal"), new XAttribute("xfId", 0), new XAttribute("builtinId", 0))), dxfs));
        }

        public static List<Color> ReadDifferentialColors(XDocument document)
        {
            var colors = new List<Color>();
            XElement dxfs = document.Descendants(S + "dxfs").FirstOrDefault();
            if (dxfs == null) return colors;
            foreach (var dxf in dxfs.Elements(S + "dxf"))
            {
                XElement fg = dxf.Descendants(S + "fgColor").FirstOrDefault();
                string rgb = fg == null ? null : (string)fg.Attribute("rgb");
                colors.Add(string.IsNullOrEmpty(rgb) ? Color.Empty : ParseColor(rgb));
            }
            return colors;
        }

        public static List<CellSnapshot> Read(XDocument document)
        {
            var result = new List<CellSnapshot>();
            var customFormats = document.Descendants(S + "numFmt").ToDictionary(x => (int)x.Attribute("numFmtId"), x => (string)x.Attribute("formatCode"));
            var fontElements = document.Descendants(S + "fonts").Elements(S + "font").ToList();
            var fillElements = document.Descendants(S + "fills").Elements(S + "fill").ToList();
            var xfs = document.Descendants(S + "cellXfs").Elements(S + "xf");
            foreach (var xf in xfs)
            {
                var cell = new CellSnapshot();
                int format = (int?)xf.Attribute("numFmtId") ?? 0;
                if (customFormats.ContainsKey(format)) cell.NumberFormat = customFormats[format];
                else if (format == 2) cell.NumberFormat = "0.00";
                else if (format == 4) cell.NumberFormat = "#,##0.00";
                else if (format == 10) cell.NumberFormat = "0.00%";
                int fontIndex = (int?)xf.Attribute("fontId") ?? 0;
                if (fontIndex >= 0 && fontIndex < fontElements.Count)
                {
                    var font = fontElements[fontIndex];
                    XElement size = font.Element(S + "sz");
                    cell.FontSize = size == null ? 10F : (float)(double)size.Attribute("val");
                    cell.FontStyle = (font.Element(S + "b") == null ? 0 : FontStyle.Bold) |
                        (font.Element(S + "i") == null ? 0 : FontStyle.Italic) |
                        (font.Element(S + "u") == null ? 0 : FontStyle.Underline);
                    cell.HasFont = fontIndex != 0 || cell.FontStyle != FontStyle.Regular;
                    XElement fontColor = font.Element(S + "color");
                    string rgb = fontColor == null ? null : (string)fontColor.Attribute("rgb");
                    if (!string.IsNullOrEmpty(rgb)) cell.ForeColor = ParseColor(rgb);
                }
                int fillIndex = (int?)xf.Attribute("fillId") ?? 0;
                if (fillIndex > 1 && fillIndex < fillElements.Count)
                {
                    XElement fillColor = fillElements[fillIndex].Descendants(S + "fgColor").FirstOrDefault();
                    string rgb = fillColor == null ? null : (string)fillColor.Attribute("rgb");
                    if (!string.IsNullOrEmpty(rgb)) cell.BackColor = ParseColor(rgb);
                }
                XElement alignment = xf.Element(S + "alignment");
                string align = alignment == null ? null : (string)alignment.Attribute("horizontal");
                if (align == "center") cell.Alignment = DataGridViewContentAlignment.MiddleCenter;
                else if (align == "right") cell.Alignment = DataGridViewContentAlignment.MiddleRight;
                else if (align == "left") cell.Alignment = DataGridViewContentAlignment.MiddleLeft;
                result.Add(cell);
            }
            return result;
        }

        private static Color ParseColor(string rgb)
        {
            if (rgb.Length == 8) rgb = rgb.Substring(2);
            return ColorTranslator.FromHtml("#" + rgb);
        }
    }
}
