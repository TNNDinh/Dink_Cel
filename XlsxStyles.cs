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
        private readonly Dictionary<string, int> borders = new Dictionary<string, int>();
        private readonly Dictionary<string, int> formats = new Dictionary<string, int>();
        private readonly Dictionary<string, int> differentialColors = new Dictionary<string, int>();

        private static string ColorKey(Color color) { return color.IsEmpty ? "" : color.ToArgb().ToString("X8", CultureInfo.InvariantCulture); }
        private static string FontKey(CellSnapshot cell)
        {
            return (cell.HasFont ? cell.FontName + ":" + cell.FontSize.ToString(CultureInfo.InvariantCulture) +
                ":" + (int)cell.FontStyle : "Arial:10:0") + ":" + ColorKey(cell.ForeColor);
        }
        private static string EdgeKey(BorderEdge edge)
        { return edge == null || !edge.Exists ? "" : edge.Style + ":" + ColorKey(edge.Color); }
        private static string BorderKey(CellSnapshot cell)
        {
            CellExtras e = cell.Extras;
            return e == null ? "|||" : EdgeKey(e.Left) + "|" + EdgeKey(e.Right) + "|" +
                EdgeKey(e.Top) + "|" + EdgeKey(e.Bottom);
        }
        private static string Key(CellSnapshot cell)
        {
            CellExtras e = cell.Extras;
            return FontKey(cell) + "|" + ColorKey(cell.BackColor) + "|" + cell.Alignment +
                "|" + cell.NumberFormat + "|" + BorderKey(cell) + "|" +
                (e == null ? "" : e.Wrap + ":" + e.Shrink + ":" + e.Indent + ":" + e.Rotation);
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
                string border = BorderKey(cell);
                if (!borders.ContainsKey(border)) borders[border] = borders.Count;
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

        private static XElement WriteBorderEdge(string name, BorderEdge edge)
        {
            var element = new XElement(S + name);
            if (edge == null || !edge.Exists) return element;
            element.SetAttributeValue("style", edge.Style);
            if (!edge.Color.IsEmpty) element.Add(ColorElement(edge.Color));
            return element;
        }

        private static BorderEdge ReadBorderEdge(XElement border, string name, Color[] theme)
        {
            XElement element = border == null ? null : border.Element(S + name);
            if (element == null) return new BorderEdge();
            return new BorderEdge { Style = (string)element.Attribute("style") ?? "",
                Color = ReadColor(element.Element(S + "color"), theme) };
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
                    new XElement(S + "name", new XAttribute("val", style.HasFont ? style.FontName : "Arial")));
                if ((style.FontStyle & FontStyle.Bold) != 0) font.Add(new XElement(S + "b"));
                if ((style.FontStyle & FontStyle.Italic) != 0) font.Add(new XElement(S + "i"));
                if ((style.FontStyle & FontStyle.Underline) != 0) font.Add(new XElement(S + "u"));
                if ((style.FontStyle & FontStyle.Strikeout) != 0) font.Add(new XElement(S + "strike"));
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
            var borderElements = new XElement(S + "borders", new XAttribute("count", borders.Count));
            foreach (var key in borders.OrderBy(x => x.Value).Select(x => x.Key))
            {
                CellExtras e = styles.First(x => BorderKey(x) == key).Extras;
                var border = new XElement(S + "border");
                border.Add(WriteBorderEdge("left", e == null ? null : e.Left));
                border.Add(WriteBorderEdge("right", e == null ? null : e.Right));
                border.Add(WriteBorderEdge("top", e == null ? null : e.Top));
                border.Add(WriteBorderEdge("bottom", e == null ? null : e.Bottom));
                border.Add(new XElement(S + "diagonal"));
                borderElements.Add(border);
            }
            var xfs = new XElement(S + "cellXfs", new XAttribute("count", styles.Count));
            foreach (CellSnapshot cell in styles)
            {
                int format = string.IsNullOrEmpty(cell.NumberFormat) ? 0 : formats[cell.NumberFormat];
                int fill = cell.BackColor.IsEmpty ? 0 : fills[ColorKey(cell.BackColor)];
                var xf = new XElement(S + "xf", new XAttribute("numFmtId", format),
                    new XAttribute("fontId", fonts[FontKey(cell)]), new XAttribute("fillId", fill),
                    new XAttribute("borderId", borders[BorderKey(cell)]), new XAttribute("xfId", 0));
                if (format != 0) xf.SetAttributeValue("applyNumberFormat", 1);
                if (fill != 0) xf.SetAttributeValue("applyFill", 1);
                if (borders[BorderKey(cell)] != 0) xf.SetAttributeValue("applyBorder", 1);
                CellExtras extras = cell.Extras;
                if (cell.Alignment != DataGridViewContentAlignment.NotSet || extras != null)
                {
                    string alignment = cell.Alignment == DataGridViewContentAlignment.TopCenter ||
                        cell.Alignment == DataGridViewContentAlignment.MiddleCenter ||
                        cell.Alignment == DataGridViewContentAlignment.BottomCenter ? "center" :
                        cell.Alignment == DataGridViewContentAlignment.TopRight ||
                        cell.Alignment == DataGridViewContentAlignment.MiddleRight ||
                        cell.Alignment == DataGridViewContentAlignment.BottomRight ? "right" : "left";
                    string vertical = cell.Alignment == DataGridViewContentAlignment.TopLeft ||
                        cell.Alignment == DataGridViewContentAlignment.TopCenter ||
                        cell.Alignment == DataGridViewContentAlignment.TopRight ? "top" :
                        cell.Alignment == DataGridViewContentAlignment.BottomLeft ||
                        cell.Alignment == DataGridViewContentAlignment.BottomCenter ||
                        cell.Alignment == DataGridViewContentAlignment.BottomRight ? "bottom" : "center";
                    var align = new XElement(S + "alignment", new XAttribute("horizontal", alignment),
                        new XAttribute("vertical", vertical));
                    if (extras != null)
                    {
                        if (extras.Wrap) align.SetAttributeValue("wrapText", 1);
                        if (extras.Shrink) align.SetAttributeValue("shrinkToFit", 1);
                        if (extras.Indent != 0) align.SetAttributeValue("indent", extras.Indent);
                        if (extras.Rotation != 0) align.SetAttributeValue("textRotation",
                            extras.Rotation < 0 ? 90 - extras.Rotation : extras.Rotation);
                    }
                    xf.Add(align);
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
                borderElements,
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
        { return Read(document, null); }

        public static List<CellSnapshot> Read(XDocument document, XDocument themeDocument)
        {
            var result = new List<CellSnapshot>();
            Color[] theme = ThemeColors(themeDocument);
            var customFormats = document.Descendants(S + "numFmt").ToDictionary(x => (int)x.Attribute("numFmtId"), x => (string)x.Attribute("formatCode"));
            var fontElements = document.Descendants(S + "fonts").Elements(S + "font").ToList();
            var fillElements = document.Descendants(S + "fills").Elements(S + "fill").ToList();
            var borderElements = document.Descendants(S + "borders").Elements(S + "border").ToList();
            var xfs = document.Descendants(S + "cellXfs").Elements(S + "xf");
            foreach (var xf in xfs)
            {
                var cell = new CellSnapshot();
                int format = (int?)xf.Attribute("numFmtId") ?? 0;
                if (customFormats.ContainsKey(format)) cell.NumberFormat = customFormats[format];
                else if (format == 2) cell.NumberFormat = "0.00";
                else if (format == 4) cell.NumberFormat = "#,##0.00";
                else if (format == 10) cell.NumberFormat = "0.00%";
                else if (format == 9) cell.NumberFormat = "0%";
                else if (format == 1) cell.NumberFormat = "0";
                else if (format == 3) cell.NumberFormat = "#,##0";
                else if (format == 11) cell.NumberFormat = "0.00E+00";
                else if (format == 12) cell.NumberFormat = "# ?/?";
                else if (format == 13) cell.NumberFormat = "# ??/??";
                else if (format >= 14 && format <= 17) cell.NumberFormat = "dd/MM/yyyy";
                else if (format >= 18 && format <= 21) cell.NumberFormat = "HH:mm:ss";
                else if (format == 22) cell.NumberFormat = "dd/MM/yyyy HH:mm";
                else if (format == 44) cell.NumberFormat = "#,##0.00;(#,##0.00);–";
                int fontIndex = (int?)xf.Attribute("fontId") ?? 0;
                if (fontIndex >= 0 && fontIndex < fontElements.Count)
                {
                    var font = fontElements[fontIndex];
                    XElement fontName = font.Element(S + "name");
                    cell.FontName = fontName == null ? "Arial" : (string)fontName.Attribute("val") ?? "Arial";
                    XElement size = font.Element(S + "sz");
                    cell.FontSize = size == null ? 10F : (float)(double)size.Attribute("val");
                    cell.FontStyle = (font.Element(S + "b") == null ? 0 : FontStyle.Bold) |
                        (font.Element(S + "i") == null ? 0 : FontStyle.Italic) |
                        (font.Element(S + "u") == null ? 0 : FontStyle.Underline) |
                        (font.Element(S + "strike") == null ? 0 : FontStyle.Strikeout);
                    cell.HasFont = true;
                    cell.ForeColor = ReadColor(font.Element(S + "color"), theme);
                }
                int fillIndex = (int?)xf.Attribute("fillId") ?? 0;
                if (fillIndex > 1 && fillIndex < fillElements.Count)
                {
                    cell.BackColor = ReadColor(fillElements[fillIndex].Descendants(S + "fgColor").FirstOrDefault(), theme);
                }
                int borderIndex = (int?)xf.Attribute("borderId") ?? 0;
                if (borderIndex > 0 && borderIndex < borderElements.Count)
                {
                    XElement border = borderElements[borderIndex];
                    cell.Extras = new CellExtras { Left = ReadBorderEdge(border, "left", theme),
                        Right = ReadBorderEdge(border, "right", theme), Top = ReadBorderEdge(border, "top", theme),
                        Bottom = ReadBorderEdge(border, "bottom", theme) };
                }
                XElement alignment = xf.Element(S + "alignment");
                string align = alignment == null ? null : (string)alignment.Attribute("horizontal");
                string vertical = alignment == null ? null : (string)alignment.Attribute("vertical");
                if (align != null || vertical != null)
                    cell.Alignment = vertical == "top" ? (align == "center" ? DataGridViewContentAlignment.TopCenter :
                        align == "right" ? DataGridViewContentAlignment.TopRight : DataGridViewContentAlignment.TopLeft) :
                        vertical == "bottom" ? (align == "center" ? DataGridViewContentAlignment.BottomCenter :
                        align == "right" ? DataGridViewContentAlignment.BottomRight : DataGridViewContentAlignment.BottomLeft) :
                        align == "center" ? DataGridViewContentAlignment.MiddleCenter :
                        align == "right" ? DataGridViewContentAlignment.MiddleRight : DataGridViewContentAlignment.MiddleLeft;
                if (alignment != null)
                {
                    int indent = (int?)alignment.Attribute("indent") ?? 0;
                    int rotation = (int?)alignment.Attribute("textRotation") ?? 0;
                    bool wrap = (bool?)alignment.Attribute("wrapText") ?? false;
                    bool shrink = (bool?)alignment.Attribute("shrinkToFit") ?? false;
                    if (indent > 0 || rotation != 0 || wrap || shrink)
                    {
                        if (cell.Extras == null) cell.Extras = new CellExtras();
                        cell.Extras.Indent = Math.Min(indent, 15);
                        cell.Extras.Rotation = rotation > 90 ? 90 - rotation : rotation;
                        cell.Extras.Wrap = wrap;
                        cell.Extras.Shrink = shrink;
                    }
                }
                result.Add(cell);
            }
            return result;
        }

        private static Color ParseColor(string rgb)
        {
            if (rgb.Length == 8) rgb = rgb.Substring(2);
            return ColorTranslator.FromHtml("#" + rgb);
        }

        private static Color[] ThemeColors(XDocument document)
        {
            Color[] colors = { Color.White, Color.Black, Color.FromArgb(238, 236, 225),
                Color.FromArgb(31, 73, 125), Color.FromArgb(79, 129, 189), Color.FromArgb(192, 80, 77),
                Color.FromArgb(155, 187, 89), Color.FromArgb(128, 100, 162),
                Color.FromArgb(75, 172, 198), Color.FromArgb(247, 150, 70) };
            if (document == null) return colors;
            XElement scheme = document.Descendants().FirstOrDefault(x => x.Name.LocalName == "clrScheme");
            if (scheme == null) return colors;
            string[] names = { "lt1", "dk1", "lt2", "dk2", "accent1", "accent2",
                "accent3", "accent4", "accent5", "accent6" };
            for (int i = 0; i < names.Length; i++)
            {
                XElement slot = scheme.Elements().FirstOrDefault(x => x.Name.LocalName == names[i]);
                if (slot == null) continue;
                XElement definition = slot.Elements().FirstOrDefault();
                if (definition == null) continue;
                string rgb = (string)definition.Attribute("lastClr") ?? (string)definition.Attribute("val");
                if (!string.IsNullOrEmpty(rgb) && rgb.Length == 6)
                    colors[i] = ParseColor(rgb);
            }
            return colors;
        }

        private static Color ReadColor(XElement element, Color[] office)
        {
            if (element == null) return Color.Empty;
            string rgb = (string)element.Attribute("rgb");
            if (!string.IsNullOrEmpty(rgb)) return ParseColor(rgb);
            int? indexed = (int?)element.Attribute("indexed");
            if (indexed.HasValue && indexed.Value == 64) return Color.Empty;
            int? theme = (int?)element.Attribute("theme");
            if (!theme.HasValue || theme.Value < 0 || theme.Value >= office.Length) return Color.Empty;
            Color source = office[theme.Value];
            double tint = (double?)element.Attribute("tint") ?? 0;
            Func<int, int> adjust = channel => Math.Max(0, Math.Min(255,
                (int)Math.Round(tint < 0 ? channel * (1 + tint) : channel + (255 - channel) * tint)));
            return Color.FromArgb(adjust(source.R), adjust(source.G), adjust(source.B));
        }
    }
}
