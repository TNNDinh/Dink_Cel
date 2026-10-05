using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DinkCel
{
    internal static class XlsxFile
    {
        private static readonly Dictionary<string, string> FutureFunctions =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                { "FILTER", "_xlfn._xlws.FILTER" }, { "SORT", "_xlfn._xlws.SORT" },
                { "SORTBY", "_xlfn.SORTBY" }, { "UNIQUE", "_xlfn.UNIQUE" },
                { "SEQUENCE", "_xlfn.SEQUENCE" }, { "LET", "_xlfn.LET" },
                { "CHOOSECOLS", "_xlfn.CHOOSECOLS" }, { "CHOOSEROWS", "_xlfn.CHOOSEROWS" },
                { "TAKE", "_xlfn.TAKE" }, { "DROP", "_xlfn.DROP" },
                { "VSTACK", "_xlfn.VSTACK" }, { "HSTACK", "_xlfn.HSTACK" }
            };

        private static string ImportFormula(string formula)
        {
            return Regex.Replace(formula ?? "",
                @"_xlfn\.(?:_xlws\.)?([A-Za-z][A-Za-z0-9_]*)\(",
                match => FutureFunctions.ContainsKey(match.Groups[1].Value) ?
                    match.Groups[1].Value + "(" : match.Value, RegexOptions.IgnoreCase);
        }

        private static string ExportFormula(string formula)
        {
            var result = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < formula.Length;)
            {
                char current = formula[i];
                if (current == '"')
                {
                    result.Append(current); i++;
                    if (quoted && i < formula.Length && formula[i] == '"')
                    { result.Append(formula[i++]); continue; }
                    quoted = !quoted;
                    continue;
                }
                if (!quoted && (Char.IsLetter(current) || current == '_'))
                {
                    int start = i;
                    while (i < formula.Length && (Char.IsLetterOrDigit(formula[i]) ||
                        formula[i] == '_')) i++;
                    string word = formula.Substring(start, i - start);
                    string future;
                    result.Append(i < formula.Length && formula[i] == '(' &&
                        FutureFunctions.TryGetValue(word, out future) ? future : word);
                }
                else { result.Append(current); i++; }
            }
            return result.ToString();
        }
        private static readonly XNamespace S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        private static readonly XNamespace P = "http://schemas.openxmlformats.org/package/2006/relationships";
        private static readonly XNamespace C = "http://schemas.openxmlformats.org/package/2006/content-types";

        private static XDocument ReadXml(ZipArchive zip, string name)
        {
            ZipArchiveEntry entry = zip.GetEntry(name);
            if (entry == null) throw new InvalidDataException("Missing XLSX part: " + name);
            using (Stream stream = entry.Open()) return XDocument.Load(stream);
        }

        private static void WriteXml(ZipArchive zip, string name, XDocument document)
        {
            var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (var stream = entry.Open())
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false)))
                document.Save(writer);
        }

        private static string ColumnName(int index)
        {
            string text = "";
            do { text = (char)('A' + index % 26) + text; index = index / 26 - 1; } while (index >= 0);
            return text;
        }

        private static Color XlsxRuleColor(string rgb)
        {
            return string.IsNullOrEmpty(rgb) || rgb.Length < 6 ? Color.LightGreen :
                ColorTranslator.FromHtml("#" + rgb.Substring(rgb.Length - 6));
        }

        private static string ValidationFormula(string kind, string value)
        {
            if (kind != "Date" && kind != "Time") return value;
            DateTime date;
            if (!DataTools.Temporal(value, out date)) return value;
            return (kind == "Time" ? date.TimeOfDay.TotalDays : date.ToOADate())
                .ToString("0.##########", CultureInfo.InvariantCulture);
        }

        private static int ColumnIndex(string address)
        {
            address = address.Replace("$", "");
            int result = 0;
            foreach (char c in address) { if (!char.IsLetter(c)) break; result = result * 26 + char.ToUpperInvariant(c) - 'A' + 1; }
            return result - 1;
        }

        private static int RowIndex(string address)
        {
            address = address.Replace("$", "");
            int i = 0; while (i < address.Length && char.IsLetter(address[i])) i++;
            int row; return int.TryParse(address.Substring(i), out row) ? row - 1 : -1;
        }

        private static string PartPath(string target)
        {
            target = target.Replace('\\', '/').TrimStart('/');
            if (target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase)) return target;
            var parts = new List<string> { "xl" };
            foreach (string part in target.Split('/'))
            {
                if (part == "..") { if (parts.Count > 0) parts.RemoveAt(parts.Count - 1); }
                else if (part != ".") parts.Add(part);
            }
            return string.Join("/", parts.ToArray());
        }

        private static string ResolvePart(string basePart, string target)
        {
            return new Uri(new Uri("http://dinkcel/" + basePart), target.Replace('\\', '/'))
                .AbsolutePath.TrimStart('/');
        }

        private static string RangeAddress(Rectangle range)
        {
            return ColumnName(range.Left) + (range.Top + 1) + ":" +
                ColumnName(range.Right - 1) + range.Bottom;
        }

        private static string AbsoluteRangeAddress(Rectangle range)
        {
            return "$" + ColumnName(range.Left) + "$" + (range.Top + 1) + ":$" +
                ColumnName(range.Right - 1) + "$" + range.Bottom;
        }

        private static Rectangle ParseArea(string reference, int rows, int columns)
        {
            string[] bounds = reference.Split(':');
            int left = ColumnIndex(bounds[0]), top = RowIndex(bounds[0]);
            int right = bounds.Length == 2 ? ColumnIndex(bounds[1]) : left;
            int bottom = bounds.Length == 2 ? RowIndex(bounds[1]) : top;
            return left < 0 || top < 0 || right >= columns || bottom >= rows || right < left || bottom < top ?
                Rectangle.Empty : new Rectangle(left, top, right - left + 1, bottom - top + 1);
        }

        // The XLSX writer builds a new package. Keep this audit deliberately conservative:
        // an unknown part or construct must be shown to the user before any save.
        public static List<string> PotentialLosses(string path)
        {
            var losses = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
            using (var zip = ZipFile.OpenRead(path))
            {
                foreach (ZipArchiveEntry entry in zip.Entries)
                {
                    string name = entry.FullName.Replace('\\', '/');
                    if (name.EndsWith("/", StringComparison.Ordinal)) continue;
                    if (name.IndexOf("vbaProject", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        name.IndexOf("macrosheets/", StringComparison.OrdinalIgnoreCase) >= 0)
                        losses.Add("VBA/macros");
                    else if (name.IndexOf("externalLinks/", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        name.IndexOf("connections", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        name.IndexOf("queryTables/", StringComparison.OrdinalIgnoreCase) >= 0)
                        losses.Add("External connections and linked workbooks");
                    else if (name.IndexOf("pivot", StringComparison.OrdinalIgnoreCase) >= 0)
                        losses.Add("Excel Pivot Tables and pivot caches");
                    else if (name.IndexOf("media/", StringComparison.OrdinalIgnoreCase) >= 0)
                        losses.Add("Images, shapes or drawing placement");
                    else if (name.StartsWith("xl/drawings/", StringComparison.OrdinalIgnoreCase) &&
                        name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                        !name.Contains("/_rels/"))
                    {
                        if (ReadXml(zip, name).Descendants().Any(e =>
                            e.Name.LocalName == "pic" || e.Name.LocalName == "sp" ||
                            e.Name.LocalName == "grpSp" || e.Name.LocalName == "cxnSp"))
                            losses.Add("Images, shapes or drawing placement");
                    }
                    else if (name.IndexOf("threadedComments", StringComparison.OrdinalIgnoreCase) >= 0 ||
                        name.IndexOf("persons/", StringComparison.OrdinalIgnoreCase) >= 0)
                        losses.Add("Threaded comments");
                    else if (name.IndexOf("comments", StringComparison.OrdinalIgnoreCase) >= 0)
                        losses.Add("Comment formatting or authors");
                    else if (name.IndexOf("charts/", StringComparison.OrdinalIgnoreCase) >= 0)
                        losses.Add("Chart details or unsupported chart types");
                    else if (!KnownPart(name)) losses.Add("Other package part: " + name);
                }
                XDocument workbook = ReadXml(zip, "xl/workbook.xml");
                if (workbook.Descendants(S + "workbookProtection").Any()) losses.Add("Workbook protection");
                if (workbook.Descendants(S + "externalReferences").Any()) losses.Add("External workbook references");
                foreach (ZipArchiveEntry entry in zip.Entries.Where(e =>
                    e.FullName.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase) &&
                    e.FullName.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) &&
                    !e.FullName.Contains("/_rels/")))
                {
                    XDocument sheet = ReadXml(zip, entry.FullName);
                    if (sheet.Descendants(S + "sheetProtection").Any()) losses.Add("Sheet protection");
                    if (sheet.Descendants(S + "pane").Any(p =>
                        (string)p.Attribute("state") == "split")) losses.Add("Split panes and independent scrolling");
                    if (sheet.Descendants(S + "legacyDrawing").Any()) losses.Add("Legacy drawings or notes");
                    if (sheet.Descendants(S + "extLst").Any()) losses.Add("Excel extensions");
                    if (sheet.Descendants(S + "f").Any(f => (string)f.Attribute("t") == "shared" ||
                        (string)f.Attribute("t") == "array" &&
                        !FormulaEngine.HasDynamicArraySyntax("=" + ImportFormula(f.Value))))
                        losses.Add("Shared or unsupported array formulas");
                    if (sheet.Descendants(S + "f").Any(f => (string)f.Attribute("t") == "array" &&
                        FormulaEngine.HasDynamicArraySyntax("=" + ImportFormula(f.Value))))
                        losses.Add("Dynamic array cached values and metadata may change");
                    if (sheet.Root.Elements().Any(e => !KnownSheetElement(e.Name.LocalName)))
                        losses.Add("Other worksheet features: " + entry.FullName);
                }
            }
            return losses.ToList();
        }

        private static bool KnownPart(string name)
        {
            return name == "[Content_Types].xml" || name == "_rels/.rels" ||
                name == "docProps/app.xml" || name == "docProps/core.xml" ||
                name == "xl/workbook.xml" || name == "xl/_rels/workbook.xml.rels" ||
                name == "xl/styles.xml" || name == "xl/sharedStrings.xml" ||
                name == "xl/theme/theme1.xml" ||
                System.Text.RegularExpressions.Regex.IsMatch(name,
                    @"^xl/worksheets/sheet\d+\.xml$|^xl/worksheets/_rels/sheet\d+\.xml\.rels$|^xl/tables/table\d+\.xml$|^xl/drawings/drawing\d+\.xml$|^xl/drawings/_rels/drawing\d+\.xml\.rels$");
        }

        private static bool KnownSheetElement(string name)
        {
            return new[] { "dimension", "sheetPr", "sheetViews", "sheetFormatPr", "cols",
                "sheetData", "autoFilter", "mergeCells", "conditionalFormatting", "dataValidations",
                "hyperlinks", "printOptions", "pageMargins", "pageSetup", "headerFooter",
                "rowBreaks", "tableParts", "drawing" }.Contains(name);
        }

        public static WorkbookSnapshot Read(string path, int rows, int columns)
        {
            using (var zip = ZipFile.OpenRead(path))
            {
                var book = ReadXml(zip, "xl/workbook.xml");
                var rels = ReadXml(zip, "xl/_rels/workbook.xml.rels");
                var paths = rels.Root.Elements(P + "Relationship").ToDictionary(x => (string)x.Attribute("Id"), x => PartPath((string)x.Attribute("Target")));
                var strings = new List<string>();
                if (zip.GetEntry("xl/sharedStrings.xml") != null)
                    foreach (var item in ReadXml(zip, "xl/sharedStrings.xml").Descendants(S + "si"))
                        strings.Add(string.Concat(item.Descendants(S + "t").Select(x => x.Value)));
                var styles = new List<CellSnapshot>();
                var differentialColors = new List<Color>();
                if (zip.GetEntry("xl/styles.xml") != null)
                {
                    XDocument styleDocument = ReadXml(zip, "xl/styles.xml");
                    XDocument themeDocument = zip.GetEntry("xl/theme/theme1.xml") == null ? null :
                        ReadXml(zip, "xl/theme/theme1.xml");
                    styles = XlsxStyles.Read(styleDocument, themeDocument);
                    differentialColors = XlsxStyles.ReadDifferentialColors(styleDocument);
                }
                var result = new WorkbookSnapshot(); result.Sheets.Clear();
                result.StructureProtected = book.Descendants(S + "workbookProtection")
                    .Any(e => (bool?)e.Attribute("lockStructure") == true);
                foreach (var sheetInfo in book.Descendants(S + "sheet"))
                {
                    string id = (string)sheetInfo.Attribute(R + "id");
                    if (!paths.ContainsKey(id)) continue;
                    string sheetPart = paths[id];
                    var document = ReadXml(zip, sheetPart);
                    var sheet = new SheetSnapshot { Name = (string)sheetInfo.Attribute("name") ?? "Sheet",
                        Hidden = string.Equals((string)sheetInfo.Attribute("state"), "hidden", StringComparison.OrdinalIgnoreCase) };
                    sheet.Protected = document.Descendants(S + "sheetProtection")
                        .Any(e => (bool?)e.Attribute("sheet") == true);
                    XElement setup = document.Root.Element(S + "pageSetup");
                    if (setup != null)
                    {
                        sheet.Print.Landscape = (string)setup.Attribute("orientation") != "portrait";
                        int paper = (int?)setup.Attribute("paperSize") ?? 9;
                        sheet.Print.Paper = paper == 1 ? "Letter" : paper == 5 ? "Legal" :
                            paper == 8 ? "A3" : "A4";
                        sheet.Print.Scale = (int?)setup.Attribute("scale") ?? 100;
                        sheet.Print.FitToOnePage = (int?)setup.Attribute("fitToWidth") == 1 &&
                            (int?)setup.Attribute("fitToHeight") == 1;
                    }
                    XElement margins = document.Root.Element(S + "pageMargins");
                    if (margins != null)
                    {
                        sheet.Print.MarginLeft = (int)Math.Round(((double?)margins.Attribute("left") ?? .5) * 100);
                        sheet.Print.MarginRight = (int)Math.Round(((double?)margins.Attribute("right") ?? .5) * 100);
                        sheet.Print.MarginTop = (int)Math.Round(((double?)margins.Attribute("top") ?? .6) * 100);
                        sheet.Print.MarginBottom = (int)Math.Round(((double?)margins.Attribute("bottom") ?? .6) * 100);
                    }
                    XElement printOptions = document.Root.Element(S + "printOptions");
                    if (printOptions != null)
                        sheet.Print.Gridlines = (bool?)printOptions.Attribute("gridLines") ?? false;
                    XElement headerFooter = document.Root.Element(S + "headerFooter");
                    if (headerFooter != null)
                    {
                        sheet.Print.Header = ((string)headerFooter.Element(S + "oddHeader") ?? "")
                            .Replace("&C", "").Replace("&L", "").Replace("&R", "");
                        sheet.Print.Footer = ((string)headerFooter.Element(S + "oddFooter") ?? "")
                            .Replace("&C", "").Replace("&L", "").Replace("&R", "");
                    }
                    foreach (XElement entry in document.Descendants(S + "rowBreaks").Elements(S + "brk"))
                    { int row = (int?)entry.Attribute("id") ?? -1; if (row > 0 && row < rows) sheet.Print.PageBreakRows.Add(row); }
                    string tabRgb = (string)document.Descendants(S + "tabColor").Select(e => e.Attribute("rgb")).FirstOrDefault();
                    if (!string.IsNullOrEmpty(tabRgb) && tabRgb.Length >= 6)
                        sheet.TabColor = ColorTranslator.FromHtml("#" + tabRgb.Substring(tabRgb.Length - 6));
                    var columnStyles = new Dictionary<int, int>();
                    XElement sheetView = document.Descendants(S + "sheetView").FirstOrDefault();
                    if (sheetView != null)
                    {
                        sheet.ShowGridlines = (bool?)sheetView.Attribute("showGridLines") ?? true;
                        sheet.ShowHeadings = (bool?)sheetView.Attribute("showRowColHeaders") ?? true;
                        sheet.FormulaView = (bool?)sheetView.Attribute("showFormulas") ?? false;
                        sheet.ViewMode = (string)sheetView.Attribute("view") ?? "normal";
                    }
                    var pane = document.Descendants(S + "pane").FirstOrDefault();
                    if (pane != null)
                    {
                        if ((string)pane.Attribute("state") == "split")
                        {
                            sheet.SplitX = (double?)pane.Attribute("xSplit") ?? 0;
                            sheet.SplitY = (double?)pane.Attribute("ySplit") ?? 0;
                        }
                        else
                        {
                            sheet.FreezeColumn = Math.Min(columns, (int?)pane.Attribute("xSplit") ?? 0);
                            sheet.FreezeRow = Math.Min(rows, (int?)pane.Attribute("ySplit") ?? 0);
                        }
                    }
                    foreach (var col in document.Descendants(S + "col"))
                    {
                        int from = (int?)col.Attribute("min") ?? 0, to = (int?)col.Attribute("max") ?? 0;
                        double width = (double?)col.Attribute("width") ?? 0;
                        for (int c = from; c <= to && c <= columns; c++) if (c > 0)
                        {
                            if (width > 0) sheet.ColumnWidths[c - 1] = Math.Max(20, (int)(width * 7 + 5));
                            if ((bool?)col.Attribute("hidden") == true) sheet.HiddenColumns.Add(c - 1);
                            if (col.Attribute("style") != null) columnStyles[c - 1] = (int)col.Attribute("style");
                        }
                    }
                    var dynamicSpans = document.Descendants(S + "c").Select(cell =>
                    {
                        XElement formula = cell.Element(S + "f");
                        return formula != null && (string)formula.Attribute("t") == "array" &&
                            FormulaEngine.HasDynamicArraySyntax("=" + ImportFormula(formula.Value)) ?
                            ParseArea((string)formula.Attribute("ref") ?? "", rows, columns) : Rectangle.Empty;
                    }).Where(area => !area.IsEmpty).ToList();
                    foreach (var row in document.Descendants(S + "sheetData").Elements(S + "row"))
                    {
                        int r = ((int?)row.Attribute("r") ?? 0) - 1;
                        double height = (double?)row.Attribute("ht") ?? 0;
                        if (r >= 0 && r < rows && height > 0) sheet.RowHeights[r] = Math.Max(2, (int)(height * 96 / 72));
                        if (r >= 0 && r < rows && (bool?)row.Attribute("hidden") == true)
                            sheet.HiddenRows.Add(r);
                        foreach (var cell in row.Elements(S + "c"))
                        {
                            string address = (string)cell.Attribute("r") ?? "";
                            int col = ColumnIndex(address), line = RowIndex(address);
                            if (col < 0 || col >= columns || line < 0 || line >= rows)
                                throw new InvalidDataException("XLSX contains cells outside the 50,000 x 26 grid: " + address);
                            string type = (string)cell.Attribute("t") ?? "";
                            string value = (string)cell.Element(S + "v") ?? "";
                            if (type == "s") { int index; if (int.TryParse(value, out index) && index >= 0 && index < strings.Count) value = strings[index]; }
                            else if (type == "inlineStr") value = string.Concat(cell.Descendants(S + "t").Select(x => x.Value));
                            else if (type == "b") value = value == "1" ? "TRUE" : "FALSE";
                            var formula = cell.Element(S + "f");
                            if (formula != null && !string.IsNullOrEmpty(formula.Value)) value = "=" + ImportFormula(formula.Value);
                            else if (formula == null && dynamicSpans.Any(area => area.Contains(col, line)))
                                value = "";
                            var snapshot = new CellSnapshot { Text = value };
                            int inheritedStyle;
                            if (!columnStyles.TryGetValue(col, out inheritedStyle)) inheritedStyle = 0;
                            int styleIndex = (int?)cell.Attribute("s") ?? (int?)row.Attribute("s") ?? inheritedStyle;
                            if (styleIndex >= 0 && styleIndex < styles.Count)
                            {
                                CellSnapshot style = styles[styleIndex];
                                snapshot.NumberFormat = style.NumberFormat;
                                snapshot.HasFont = style.HasFont;
                                snapshot.FontName = style.FontName;
                                snapshot.FontStyle = style.FontStyle;
                                snapshot.FontSize = style.FontSize;
                                snapshot.ForeColor = style.ForeColor;
                                snapshot.BackColor = style.BackColor;
                                snapshot.Alignment = style.Alignment;
                                snapshot.Extras = CellExtras.Copy(style.Extras);
                            }
                            sheet.Cells[line * columns + col] = snapshot;
                        }
                    }
                    foreach (var merge in document.Descendants(S + "mergeCell"))
                    {
                        string[] bounds = ((string)merge.Attribute("ref") ?? "").Split(':');
                        if (bounds.Length == 2)
                        {
                            int x = ColumnIndex(bounds[0]), y = RowIndex(bounds[0]);
                            int right = ColumnIndex(bounds[1]), bottom = RowIndex(bounds[1]);
                            if (x >= 0 && y >= 0 && right < columns && bottom < rows)
                                sheet.Merges.Add(new Rectangle(x, y, right - x + 1, bottom - y + 1));
                        }
                    }
                    foreach (var formatting in document.Descendants(S + "conditionalFormatting"))
                    {
                        string sqref = (string)formatting.Attribute("sqref") ?? "";
                        foreach (string area in sqref.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                        {
                            string[] bounds = area.Split(':');
                            int x = ColumnIndex(bounds[0]), y = RowIndex(bounds[0]);
                            int right = bounds.Length == 2 ? ColumnIndex(bounds[1]) : x;
                            int bottom = bounds.Length == 2 ? RowIndex(bounds[1]) : y;
                            if (x < 0 || y < 0 || right >= columns || bottom >= rows) continue;
                            foreach (var rule in formatting.Elements(S + "cfRule"))
                            {
                                string type = (string)rule.Attribute("type") ?? "";
                                string op = (string)rule.Attribute("operator") ?? "";
                                int dxf = (int?)rule.Attribute("dxfId") ?? -1;
                                var formulas = rule.Elements(S + "formula").Select(v => v.Value).ToList();
                                string kind = type == "cellIs" ? op == "greaterThan" ? "Greater" : op == "lessThan" ? "Less" :
                                    op == "between" ? "Between" : op == "equal" ? "Equal" : "" :
                                    type == "duplicateValues" ? "Duplicate" : type == "uniqueValues" ? "Unique" :
                                    type == "containsText" ? "Text Contains" : type == "containsBlanks" ? "Blank" :
                                    type == "expression" ? "Formula" : type == "colorScale" ? "Color Scale" :
                                    type == "dataBar" ? "Data Bar" : type == "iconSet" ? "Icon Set" : "";
                                if (kind.Length == 0) continue;
                                var entry = new ConditionalRule { Range = new Rectangle(x, y, right - x + 1, bottom - y + 1),
                                    Kind = kind, Value1 = formulas.Count > 0 ? formulas[0] : "",
                                    Value2 = formulas.Count > 1 ? formulas[1] : "",
                                    Color = dxf >= 0 && dxf < differentialColors.Count ? differentialColors[dxf] : Color.LightGreen };
                                if (kind == "Formula" && !entry.Value1.StartsWith("=", StringComparison.Ordinal)) entry.Value1 = "=" + entry.Value1;
                                if (kind == "Text Contains") entry.Value1 = (string)rule.Attribute("text") ?? entry.Value1;
                                if (kind == "Color Scale")
                                {
                                    var colors = rule.Descendants(S + "color").Select(v => XlsxRuleColor((string)v.Attribute("rgb"))).ToList();
                                    if (colors.Count > 0) entry.Color2 = colors[0];
                                    if (colors.Count > 1) entry.Color = colors[colors.Count - 1];
                                }
                                if (kind == "Data Bar") entry.Color = XlsxRuleColor((string)rule.Descendants(S + "color").Select(v => v.Attribute("rgb")).FirstOrDefault());
                                double threshold;
                                if (double.TryParse(entry.Value1, NumberStyles.Float, CultureInfo.InvariantCulture, out threshold)) entry.Threshold = threshold;
                                sheet.Rules.Add(entry);
                            }
                        }
                    }
                    XElement filter = document.Descendants(S + "autoFilter").FirstOrDefault();
                    if (filter != null)
                    {
                        foreach (XElement filterColumn in filter.Descendants(S + "filterColumn"))
                        {
                            int column = (int?)filterColumn.Attribute("colId") ?? -1;
                            if (column < 0 || column >= columns) continue;
                            var customs = filterColumn.Descendants(S + "customFilter").ToList();
                            if (customs.Count == 0)
                            {
                                XElement choicesRoot = filterColumn.Element(S + "filters");
                                if (choicesRoot != null)
                                {
                                    var choices = choicesRoot.Elements(S + "filter").Select(v => (string)v.Attribute("val") ?? "").ToList();
                                    if ((bool?)choicesRoot.Attribute("blank") == true) choices.Add("");
                                    if (choices.Count > 0)
                                    {
                                        sheet.Filters.Add(new FilterCriterion { Column = column, Operator = "One Of",
                                            Value1 = string.Join("\n", choices.ToArray()) });
                                        if (sheet.FilterColumn < 0) sheet.FilterColumn = column;
                                    }
                                }
                            }
                            if (customs.Count == 0) continue;
                            string first = (string)customs[0].Attribute("val") ?? "";
                            string op = (string)customs[0].Attribute("operator") ?? "equal";
                            string kind = op == "greaterThan" || op == "lessThan" || op == "greaterThanOrEqual" || op == "lessThanOrEqual" ? "Number" : "Text";
                            string mapped = op == "greaterThan" ? "Greater" : op == "lessThan" ? "Less" :
                                op == "notEqual" && first.Length == 0 ? "Nonblank" :
                                first.Length == 0 ? "Blank" : first.StartsWith("*") && first.EndsWith("*") ? "Contains" :
                                first.EndsWith("*") ? "Begins With" : "Equals";
                            if (customs.Count > 1) mapped = "Between";
                            sheet.Filters.Add(new FilterCriterion { Column = column, Kind = kind, Operator = mapped,
                                Value1 = first.Trim('*'), Value2 = customs.Count > 1 ? ((string)customs[1].Attribute("val") ?? "").Trim('*') : "" });
                            if (sheet.FilterColumn < 0)
                            { sheet.FilterColumn = column; sheet.FilterValue = first.Trim('*'); }
                        }
                    }
                    foreach (var validation in document.Descendants(S + "dataValidation"))
                    {
                        string type = (string)validation.Attribute("type") ?? "";
                        string kind = type == "list" ? "List" : type == "whole" ? "Whole Number" :
                            type == "decimal" ? "Decimal" : type == "date" ? "Date" : type == "time" ? "Time" :
                            type == "textLength" ? "Text Length" : type == "custom" ? "Custom Formula" : "";
                        if (kind.Length == 0) continue;
                        string formula = (string)validation.Element(S + "formula1") ?? "";
                        string[] choices = kind == "List" && formula.StartsWith("\"", StringComparison.Ordinal) && formula.EndsWith("\"", StringComparison.Ordinal) ?
                            formula.Substring(1, formula.Length - 2).Replace("\"\"", "\"").Split(',') : new string[0];
                        if (kind == "List" && choices.Length == 0)
                        {
                            Rectangle source = ParseArea(formula.TrimStart('='), rows, columns);
                            if (!source.IsEmpty)
                                choices = Enumerable.Range(source.Top, source.Height)
                                    .SelectMany(r => Enumerable.Range(source.Left, source.Width)
                                        .Select(c => { CellSnapshot item; return sheet.Cells.TryGetValue(r * columns + c, out item) ? item.Text : ""; }))
                                    .Where(v => v.Length > 0).ToArray();
                        }
                        string sqref = (string)validation.Attribute("sqref") ?? "";
                        foreach (string area in sqref.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries))
                        {
                            Rectangle range = ParseArea(area, rows, columns);
                            if (range.IsEmpty) continue;
                            string operation = (string)validation.Attribute("operator") ?? "between";
                            var rule = new ValidationRule { Range = range, Kind = kind,
                                Operator = operation == "notBetween" ? "Not Between" : operation == "equal" ? "Equals" :
                                    operation == "notEqual" ? "Not Equals" : operation == "greaterThan" ? "Greater" :
                                    operation == "greaterThanOrEqual" ? "Greater Or Equal" : operation == "lessThan" ? "Less" :
                                    operation == "lessThanOrEqual" ? "Less Or Equal" : "Between",
                                Value1 = kind == "Custom Formula" ? "=" + formula.TrimStart('=') : formula,
                                Value2 = (string)validation.Element(S + "formula2") ?? "",
                                AllowBlank = (bool?)validation.Attribute("allowBlank") ?? true,
                                InputTitle = (string)validation.Attribute("promptTitle") ?? "",
                                InputMessage = (string)validation.Attribute("prompt") ?? "",
                                ErrorTitle = (string)validation.Attribute("errorTitle") ?? "",
                                ErrorMessage = (string)validation.Attribute("error") ?? "",
                                ErrorStyle = (string)validation.Attribute("errorStyle") == "warning" ? "Warning" :
                                    (string)validation.Attribute("errorStyle") == "information" ? "Information" : "Stop" };
                            rule.Choices.AddRange(choices);
                            sheet.Validations.Add(rule);
                        }
                    }
                    string relPath = sheetPart.Substring(0, sheetPart.LastIndexOf('/') + 1) +
                        "_rels/" + sheetPart.Substring(sheetPart.LastIndexOf('/') + 1) + ".rels";
                    var linkTargets = new Dictionary<string, string>();
                    if (zip.GetEntry(relPath) != null)
                        foreach (XElement rel in ReadXml(zip, relPath).Root.Elements(P + "Relationship"))
                            if (((string)rel.Attribute("Type") ?? "").EndsWith("/hyperlink", StringComparison.Ordinal))
                                linkTargets[(string)rel.Attribute("Id")] = (string)rel.Attribute("Target") ?? "";
                    foreach (XElement link in document.Descendants(S + "hyperlink"))
                    {
                        string reference = (string)link.Attribute("ref") ?? "";
                        int col = ColumnIndex(reference), row = RowIndex(reference);
                        if (col < 0 || col >= columns || row < 0 || row >= rows) continue;
                        string target = (string)link.Attribute("location");
                        if (target != null) target = "#" + target;
                        else if (!linkTargets.TryGetValue((string)link.Attribute(R + "id") ?? "", out target)) continue;
                        int key = row * columns + col;
                        CellSnapshot cell;
                        if (!sheet.Cells.TryGetValue(key, out cell))
                        { cell = new CellSnapshot(); sheet.Cells[key] = cell; }
                        if (cell.Extras == null) cell.Extras = new CellExtras();
                        cell.Extras.Hyperlink = target;
                    }
                    if (zip.GetEntry(relPath) != null)
                        foreach (XElement rel in ReadXml(zip, relPath).Root.Elements(P + "Relationship"))
                        {
                            if (!((string)rel.Attribute("Type") ?? "").EndsWith("/comments", StringComparison.Ordinal)) continue;
                            string commentsPart = ResolvePart(sheetPart, (string)rel.Attribute("Target") ?? "");
                            if (zip.GetEntry(commentsPart) == null) continue;
                            foreach (XElement note in ReadXml(zip, commentsPart).Descendants(S + "comment"))
                            {
                                string reference = (string)note.Attribute("ref") ?? "";
                                int col = ColumnIndex(reference), row = RowIndex(reference);
                                if (col < 0 || col >= columns || row < 0 || row >= rows) continue;
                                int key = row * columns + col;
                                CellSnapshot cell;
                                if (!sheet.Cells.TryGetValue(key, out cell))
                                { cell = new CellSnapshot(); sheet.Cells[key] = cell; }
                                if (cell.Extras == null) cell.Extras = new CellExtras();
                                cell.Extras.Note = string.Concat(note.Descendants(S + "t").Select(t => t.Value));
                            }
                        }
                    if (zip.GetEntry(relPath) != null)
                    {
                        var sheetRels = ReadXml(zip, relPath);
                        var tablePaths = sheetRels.Root.Elements(P + "Relationship")
                            .Where(x => ((string)x.Attribute("Type") ?? "").EndsWith("/table", StringComparison.Ordinal))
                            .ToDictionary(x => (string)x.Attribute("Id"),
                                x => ResolvePart(sheetPart, (string)x.Attribute("Target")));
                        foreach (var tablePart in document.Descendants(S + "tablePart"))
                        {
                            string tableId = (string)tablePart.Attribute(R + "id");
                            if (!tablePaths.ContainsKey(tableId) || zip.GetEntry(tablePaths[tableId]) == null) continue;
                            XElement tableDefinition = ReadXml(zip, tablePaths[tableId]).Root;
                            Rectangle range = ParseArea((string)tableDefinition.Attribute("ref") ?? "", rows, columns);
                            if (!range.IsEmpty)
                            {
                                XElement style = tableDefinition.Element(S + "tableStyleInfo");
                                var table = new TableDefinition
                                { Name = (string)tableDefinition.Attribute("displayName") ??
                                    (string)tableDefinition.Attribute("name") ?? "Table", Range = range,
                                    HeaderRow = ((int?)tableDefinition.Attribute("headerRowCount") ?? 1) != 0,
                                    TotalRow = ((int?)tableDefinition.Attribute("totalsRowCount") ?? 0) != 0,
                                    Style = (string)(style == null ? null : style.Attribute("name")) ?? "TableStyleMedium2",
                                    BandedRows = (bool?)(style == null ? null : style.Attribute("showRowStripes")) ?? true,
                                    BandedColumns = (bool?)(style == null ? null : style.Attribute("showColumnStripes")) ?? false,
                                    Filter = tableDefinition.Element(S + "autoFilter") != null };
                                int field = range.Left;
                                foreach (XElement column in tableDefinition.Descendants(S + "tableColumn"))
                                {
                                    string formula = (string)column.Element(S + "calculatedColumnFormula") ?? "";
                                    if (formula.Length > 0) table.CalculatedColumns[field] = "=" + formula.TrimStart('=');
                                    field++;
                                }
                                foreach (XElement filterColumn in tableDefinition.Descendants(S + "filterColumn"))
                                {
                                    int column = range.Left + ((int?)filterColumn.Attribute("colId") ?? -1);
                                    if (column < range.Left || column >= range.Right) continue;
                                    var customs = filterColumn.Descendants(S + "customFilter").ToList();
                                    if (customs.Count == 0)
                                    {
                                        XElement choicesRoot = filterColumn.Element(S + "filters");
                                        if (choicesRoot != null)
                                        {
                                            var choices = choicesRoot.Elements(S + "filter").Select(v => (string)v.Attribute("val") ?? "").ToList();
                                            if ((bool?)choicesRoot.Attribute("blank") == true) choices.Add("");
                                            if (choices.Count > 0) table.Filters.Add(new FilterCriterion { Column = column,
                                                Operator = "One Of", Value1 = string.Join("\n", choices.ToArray()) });
                                        }
                                    }
                                    if (customs.Count == 0) continue;
                                    string value = (string)customs[0].Attribute("val") ?? "";
                                    string op = (string)customs[0].Attribute("operator") ?? "equal";
                                    string mapped = customs.Count > 1 ? "Between" : op == "greaterThan" ? "Greater" :
                                        op == "lessThan" ? "Less" : op == "notEqual" && value.Length == 0 ? "Nonblank" :
                                        value.Length == 0 ? "Blank" : value.StartsWith("*") && value.EndsWith("*") ? "Contains" :
                                        value.EndsWith("*") ? "Begins With" : "Equals";
                                    table.Filters.Add(new FilterCriterion { Column = column,
                                        Kind = op == "greaterThan" || op == "lessThan" || op == "greaterThanOrEqual" ||
                                            op == "lessThanOrEqual" ? "Number" : "Text",
                                        Operator = mapped, Value1 = value.Trim('*'),
                                        Value2 = customs.Count > 1 ? (string)customs[1].Attribute("val") ?? "" : "" });
                                }
                                sheet.Tables.Add(table);
                            }
                        }
                        var drawingPaths = sheetRels.Root.Elements(P + "Relationship")
                            .Where(x => ((string)x.Attribute("Type") ?? "").EndsWith("/drawing", StringComparison.Ordinal))
                            .ToDictionary(x => (string)x.Attribute("Id"),
                                x => ResolvePart(sheetPart, (string)x.Attribute("Target")));
                        foreach (var drawing in document.Descendants(S + "drawing"))
                        {
                            string drawingId = (string)drawing.Attribute(R + "id");
                            if (!drawingPaths.ContainsKey(drawingId) || zip.GetEntry(drawingPaths[drawingId]) == null) continue;
                            string drawingPart = drawingPaths[drawingId];
                            XDocument drawingXml = ReadXml(zip, drawingPart);
                            string drawingRelPath = drawingPart.Substring(0, drawingPart.LastIndexOf('/') + 1) +
                                "_rels/" + drawingPart.Substring(drawingPart.LastIndexOf('/') + 1) + ".rels";
                            if (zip.GetEntry(drawingRelPath) == null) continue;
                            XDocument drawingRels = ReadXml(zip, drawingRelPath);
                            var chartPaths = drawingRels.Root.Elements(P + "Relationship")
                                .Where(x => ((string)x.Attribute("Type") ?? "").EndsWith("/chart", StringComparison.Ordinal))
                                .ToDictionary(x => (string)x.Attribute("Id"),
                                    x => ResolvePart(drawingPart, (string)x.Attribute("Target")));
                            foreach (XElement chartRef in drawingXml.Descendants(XlsxCharts.Chart + "chart"))
                            {
                                string chartId = (string)chartRef.Attribute(R + "id");
                                if (!chartPaths.ContainsKey(chartId) || zip.GetEntry(chartPaths[chartId]) == null) continue;
                                ChartDefinition definition = XlsxCharts.ParseChart(
                                    ReadXml(zip, chartPaths[chartId]), columns, rows);
                                if (definition != null)
                                {
                                    XElement anchor = chartRef.Ancestors(XlsxCharts.SpreadsheetDrawing + "twoCellAnchor")
                                        .FirstOrDefault();
                                    if (anchor != null)
                                    {
                                        XElement from = anchor.Element(XlsxCharts.SpreadsheetDrawing + "from");
                                        XElement to = anchor.Element(XlsxCharts.SpreadsheetDrawing + "to");
                                        if (from != null && to != null)
                                        {
                                            int x1 = (int?)from.Element(XlsxCharts.SpreadsheetDrawing + "col") ?? 0;
                                            int y1 = (int?)from.Element(XlsxCharts.SpreadsheetDrawing + "row") ?? 0;
                                            int x2 = (int?)to.Element(XlsxCharts.SpreadsheetDrawing + "col") ?? x1 + 8;
                                            int y2 = (int?)to.Element(XlsxCharts.SpreadsheetDrawing + "row") ?? y1 + 14;
                                            definition.Placement = Rectangle.FromLTRB(x1, y1, x2, y2);
                                        }
                                    }
                                    sheet.Charts.Add(definition);
                                }
                            }
                        }
                    }
                    result.Sheets.Add(sheet);
                }
                if (result.Sheets.Count == 0) result.Sheets.Add(new SheetSnapshot());
                foreach (var defined in book.Descendants(S + "definedName"))
                {
                    string name = (string)defined.Attribute("name") ?? "";
                    if (name == "_xlnm.Print_Area" || name == "_xlnm.Print_Titles")
                    {
                        int index = (int?)defined.Attribute("localSheetId") ?? -1;
                        if (index < 0 || index >= result.Sheets.Count) continue;
                        int bangIndex = defined.Value.LastIndexOf('!');
                        if (bangIndex < 0) continue;
                        string address = defined.Value.Substring(bangIndex + 1);
                        if (name == "_xlnm.Print_Area")
                            result.Sheets[index].Print.PrintArea = ParseArea(address, rows, columns);
                        else
                        {
                            var match = System.Text.RegularExpressions.Regex.Match(address,
                                @"\$?1:\$?([1-9][0-9]*)");
                            int count;
                            if (match.Success && int.TryParse(match.Groups[1].Value, out count))
                                result.Sheets[index].Print.TitleRows = Math.Min(rows, count);
                        }
                        continue;
                    }
                    if (name.Length == 0 || name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)) continue;
                    int localIndex = (int?)defined.Attribute("localSheetId") ?? -1;
                    if (localIndex >= result.Sheets.Count) continue;
                    string scope = localIndex < 0 ? "" : result.Sheets[localIndex].Name;
                    var definedAddress = System.Text.RegularExpressions.Regex.Match(defined.Value,
                        @"^(?:'(?<quoted>(?:[^']|'')+)'|(?<plain>[^!]+))!(?<area>\$?[A-Za-z]{1,3}\$?[1-9][0-9]*(?::\$?[A-Za-z]{1,3}\$?[1-9][0-9]*)?)$");
                    Rectangle range = definedAddress.Success ? ParseArea(definedAddress.Groups["area"].Value, rows, columns) : Rectangle.Empty;
                    if (!range.IsEmpty)
                        result.NamedRanges.Add(new NamedRange { Name = name,
                            Sheet = (definedAddress.Groups["quoted"].Success ? definedAddress.Groups["quoted"].Value :
                                definedAddress.Groups["plain"].Value).Replace("''", "'"),
                            ScopeSheet = scope, Range = range });
                    else
                        result.NamedRanges.Add(new NamedRange { Name = name, Sheet = scope,
                            ScopeSheet = scope, Formula = "=" + ImportFormula(defined.Value) });
                }
                return result;
            }
        }

        public static void Write(string path, WorkbookSnapshot book, int rows, int columns)
        {
            string temporary = path + ".tmp";
            try
            {
                if (File.Exists(temporary)) File.Delete(temporary);
                using (var zip = ZipFile.Open(temporary, ZipArchiveMode.Create))
                {
                    var types = new XElement(C + "Types",
                        new XElement(C + "Default", new XAttribute("Extension", "rels"), new XAttribute("ContentType", "application/vnd.openxmlformats-package.relationships+xml")),
                        new XElement(C + "Default", new XAttribute("Extension", "xml"), new XAttribute("ContentType", "application/xml")),
                        new XElement(C + "Override", new XAttribute("PartName", "/xl/workbook.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")),
                        new XElement(C + "Override", new XAttribute("PartName", "/xl/styles.xml"), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")));
                    var sheetsElement = new XElement(S + "sheets");
                    var relationships = new XElement(P + "Relationships");
                    var styleCatalog = new XlsxStyles(book);
                    int tableNumber = 0;
                    int chartNumber = 0;
                    for (int i = 0; i < book.Sheets.Count; i++)
                    {
                        int number = i + 1;
                        string part = "xl/worksheets/sheet" + number + ".xml";
                        types.Add(new XElement(C + "Override", new XAttribute("PartName", "/" + part), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
                        var sheetEntry = new XElement(S + "sheet", new XAttribute("name", book.Sheets[i].Name),
                            new XAttribute("sheetId", number), new XAttribute(R + "id", "rId" + number));
                        if (book.Sheets[i].Hidden) sheetEntry.SetAttributeValue("state", "hidden");
                        sheetsElement.Add(sheetEntry);
                        relationships.Add(new XElement(P + "Relationship", new XAttribute("Id", "rId" + number), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"), new XAttribute("Target", "worksheets/sheet" + number + ".xml")));
                        var tableIds = new List<string>();
                        var sheetRels = new XElement(P + "Relationships");
                        foreach (TableDefinition table in book.Sheets[i].Tables)
                        {
                            if (table.Range.IsEmpty || table.Range.Right > columns || table.Range.Bottom > rows) continue;
                            int tableId = ++tableNumber;
                            string tablePart = "xl/tables/table" + tableId + ".xml";
                            string relationId = "rIdTable" + tableId;
                            tableIds.Add(relationId);
                            sheetRels.Add(new XElement(P + "Relationship", new XAttribute("Id", relationId),
                                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"),
                                new XAttribute("Target", "../tables/table" + tableId + ".xml")));
                            types.Add(new XElement(C + "Override", new XAttribute("PartName", "/" + tablePart),
                                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml")));
                            WriteXml(zip, tablePart, BuildTable(table, book.Sheets[i], tableId, columns));
                        }
                        string drawingRelation = null;
                        if (book.Sheets[i].Charts.Count > 0)
                        {
                            drawingRelation = "rIdDrawing" + number;
                            string drawingPart = "xl/drawings/drawing" + number + ".xml";
                            sheetRels.Add(new XElement(P + "Relationship", new XAttribute("Id", drawingRelation),
                                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"),
                                new XAttribute("Target", "../drawings/drawing" + number + ".xml")));
                            types.Add(new XElement(C + "Override", new XAttribute("PartName", "/" + drawingPart),
                                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawing+xml")));
                            var chartNumbers = new List<int>();
                            foreach (ChartDefinition chart in book.Sheets[i].Charts)
                            {
                                int chartId = ++chartNumber;
                                chartNumbers.Add(chartId);
                                string chartPart = "xl/charts/chart" + chartId + ".xml";
                                types.Add(new XElement(C + "Override", new XAttribute("PartName", "/" + chartPart),
                                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")));
                                WriteXml(zip, chartPart, XlsxCharts.BuildChart(chart, book.Sheets[i], chartId, columns));
                            }
                            WriteXml(zip, drawingPart, XlsxCharts.BuildDrawing(book.Sheets[i].Charts));
                            WriteXml(zip, "xl/drawings/_rels/drawing" + number + ".xml.rels",
                                XlsxCharts.BuildDrawingRelationships(chartNumbers));
                        }
                        var hyperlinkIds = new Dictionary<int, string>();
                        foreach (var cell in book.Sheets[i].Cells)
                        {
                            string link = cell.Value.Extras == null ? "" : cell.Value.Extras.Hyperlink;
                            if (string.IsNullOrWhiteSpace(link) || link.StartsWith("#", StringComparison.Ordinal)) continue;
                            string relationId = "rIdHyperlink" + cell.Key;
                            hyperlinkIds[cell.Key] = relationId;
                            sheetRels.Add(new XElement(P + "Relationship", new XAttribute("Id", relationId),
                                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"),
                                new XAttribute("Target", link), new XAttribute("TargetMode", "External")));
                        }
                        string commentsRelation = null;
                        string vmlRelation = null;
                        if (book.Sheets[i].Cells.Any(c => c.Value.Extras != null &&
                            !string.IsNullOrEmpty(c.Value.Extras.Note)))
                        {
                            commentsRelation = "rIdComments" + number;
                            vmlRelation = "rIdVmlComments" + number;
                            string commentsPart = "xl/comments/comment" + number + ".xml";
                            string vmlPart = "xl/drawings/commentsDrawing" + number + ".vml";
                            sheetRels.Add(new XElement(P + "Relationship", new XAttribute("Id", commentsRelation),
                                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"),
                                new XAttribute("Target", "../comments/comment" + number + ".xml")));
                            sheetRels.Add(new XElement(P + "Relationship", new XAttribute("Id", vmlRelation),
                                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"),
                                new XAttribute("Target", "../drawings/commentsDrawing" + number + ".vml")));
                            types.Add(new XElement(C + "Override", new XAttribute("PartName", "/" + commentsPart),
                                new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml")));
                            if (!types.Elements(C + "Default").Any(e => (string)e.Attribute("Extension") == "vml"))
                                types.Add(new XElement(C + "Default", new XAttribute("Extension", "vml"),
                                    new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.vmlDrawing")));
                            WriteXml(zip, commentsPart, BuildComments(book.Sheets[i], columns));
                            WriteXml(zip, vmlPart, BuildCommentsVml(book.Sheets[i], columns));
                        }
                        if (sheetRels.HasElements)
                            WriteXml(zip, "xl/worksheets/_rels/sheet" + number + ".xml.rels", new XDocument(sheetRels));
                        WriteXml(zip, part, BuildSheet(book.Sheets[i], rows, columns, styleCatalog,
                            tableIds, drawingRelation, hyperlinkIds, vmlRelation));
                    }
                    relationships.Add(new XElement(P + "Relationship", new XAttribute("Id", "rIdStyles"), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"), new XAttribute("Target", "styles.xml")));
                    WriteXml(zip, "[Content_Types].xml", new XDocument(types));
                    WriteXml(zip, "_rels/.rels", new XDocument(new XElement(P + "Relationships", new XElement(P + "Relationship", new XAttribute("Id", "rId1"), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"), new XAttribute("Target", "xl/workbook.xml")))));
                    var workbookElement = new XElement(S + "workbook", new XAttribute(XNamespace.Xmlns + "r", R));
                    if (book.StructureProtected)
                        workbookElement.Add(new XElement(S + "workbookProtection", new XAttribute("lockStructure", 1)));
                    workbookElement.Add(sheetsElement);
                    if (book.NamedRanges.Count > 0 || book.Sheets.Any(s => !s.Print.PrintArea.IsEmpty ||
                        s.Print.TitleRows > 0))
                    {
                        var definitions = new XElement(S + "definedNames");
                        foreach (NamedRange named in book.NamedRanges)
                        {
                            int scopeIndex = String.IsNullOrEmpty(named.ScopeSheet) ? -1 :
                                book.Sheets.FindIndex(s => String.Equals(s.Name, named.ScopeSheet,
                                    StringComparison.OrdinalIgnoreCase));
                            if (!String.IsNullOrEmpty(named.ScopeSheet) && scopeIndex < 0) continue;
                            string expression = !String.IsNullOrEmpty(named.Formula) ?
                                ExportFormula(named.Formula.TrimStart('=')) :
                                named.Range.IsEmpty || String.IsNullOrEmpty(named.Sheet) ? "" :
                                "'" + named.Sheet.Replace("'", "''") + "'!" + AbsoluteRangeAddress(named.Range);
                            if (expression.Length == 0) continue;
                            var element = new XElement(S + "definedName", new XAttribute("name", named.Name), expression);
                            if (scopeIndex >= 0) element.Add(new XAttribute("localSheetId", scopeIndex));
                            definitions.Add(element);
                        }
                        for (int i = 0; i < book.Sheets.Count; i++)
                        {
                            SheetSnapshot sheet = book.Sheets[i];
                            string quoted = "'" + sheet.Name.Replace("'", "''") + "'!";
                            if (!sheet.Print.PrintArea.IsEmpty)
                                definitions.Add(new XElement(S + "definedName",
                                    new XAttribute("name", "_xlnm.Print_Area"),
                                    new XAttribute("localSheetId", i),
                                    quoted + AbsoluteRangeAddress(sheet.Print.PrintArea)));
                            if (sheet.Print.TitleRows > 0)
                                definitions.Add(new XElement(S + "definedName",
                                    new XAttribute("name", "_xlnm.Print_Titles"),
                                    new XAttribute("localSheetId", i),
                                    quoted + "$1:$" + sheet.Print.TitleRows));
                        }
                        workbookElement.Add(definitions);
                    }
                    workbookElement.Add(new XElement(S + "calcPr", new XAttribute("calcMode", "auto"),
                        new XAttribute("fullCalcOnLoad", 1)));
                    WriteXml(zip, "xl/workbook.xml", new XDocument(workbookElement));
                    WriteXml(zip, "xl/_rels/workbook.xml.rels", new XDocument(relationships));
                    WriteXml(zip, "xl/styles.xml", styleCatalog.Document());
                }
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        private static XDocument BuildComments(SheetSnapshot sheet, int columns)
        {
            var list = new XElement(S + "commentList");
            foreach (var pair in sheet.Cells.OrderBy(c => c.Key))
            {
                string note = pair.Value.Extras == null ? "" : pair.Value.Extras.Note;
                if (string.IsNullOrEmpty(note)) continue;
                list.Add(new XElement(S + "comment",
                    new XAttribute("ref", ColumnName(pair.Key % columns) + (pair.Key / columns + 1)),
                    new XAttribute("authorId", 0),
                    new XElement(S + "text", new XElement(S + "t", note))));
            }
            return new XDocument(new XElement(S + "comments",
                new XElement(S + "authors", new XElement(S + "author", "DinkCel")), list));
        }

        private static XDocument BuildCommentsVml(SheetSnapshot sheet, int columns)
        {
            XNamespace v = "urn:schemas-microsoft-com:vml";
            XNamespace o = "urn:schemas-microsoft-com:office:office";
            XNamespace x = "urn:schemas-microsoft-com:office:excel";
            var root = new XElement("xml", new XAttribute(XNamespace.Xmlns + "v", v),
                new XAttribute(XNamespace.Xmlns + "o", o), new XAttribute(XNamespace.Xmlns + "x", x),
                new XElement(o + "shapelayout", new XAttribute(v + "ext", "edit"),
                    new XElement(o + "idmap", new XAttribute(v + "ext", "edit"), new XAttribute("data", 1))),
                new XElement(v + "shapetype", new XAttribute("id", "_x0000_t202"),
                    new XAttribute("coordsize", "21600,21600"), new XAttribute(o + "spt", 202),
                    new XAttribute("path", "m,l,21600r21600,l21600,xe"),
                    new XElement(v + "stroke", new XAttribute("joinstyle", "miter")),
                    new XElement(v + "path", new XAttribute("gradientshapeok", "t"),
                        new XAttribute(o + "connecttype", "rect"))));
            int shape = 1025;
            foreach (var pair in sheet.Cells.OrderBy(c => c.Key))
            {
                if (pair.Value.Extras == null || string.IsNullOrEmpty(pair.Value.Extras.Note)) continue;
                int row = pair.Key / columns, col = pair.Key % columns;
                root.Add(new XElement(v + "shape", new XAttribute("id", "_x0000_s" + shape++),
                    new XAttribute("type", "#_x0000_t202"),
                    new XAttribute("style", "position:absolute;visibility:hidden"),
                    new XAttribute("fillcolor", "#ffffe1"),
                    new XAttribute(o + "insetmode", "auto"),
                    new XElement(v + "textbox", new XAttribute("style", "mso-direction-alt:auto"),
                        new XElement("div", new XAttribute("style", "text-align:left"))),
                    new XElement(x + "ClientData", new XAttribute("ObjectType", "Note"),
                        new XElement(x + "MoveWithCells"), new XElement(x + "SizeWithCells"),
                        new XElement(x + "Anchor", col + ", 15, " + row + ", 2, " +
                            Math.Min(25, col + 3) + ", 15, " + (row + 4) + ", 2"),
                        new XElement(x + "AutoFill", "False"),
                        new XElement(x + "Row", row), new XElement(x + "Column", col))));
            }
            return new XDocument(root);
        }

        private static XDocument BuildSheet(SheetSnapshot sheet, int rows, int columns, XlsxStyles styles,
            IList<string> tableIds, string drawingRelation, IDictionary<int, string> hyperlinkIds,
            string vmlRelation)
        {
            var root = new XElement(S + "worksheet", new XAttribute(XNamespace.Xmlns + "r", R));
            var dynamicKeys = sheet.Cells.Where(x => FormulaEngine.HasDynamicArraySyntax(x.Value.Text))
                .Select(x => x.Key).ToArray();
            var spillDimensions = new Dictionary<int, int[]>();
            if (dynamicKeys.Length > 0)
            {
                var engine = new FormulaEngine((row, column) =>
                {
                    CellSnapshot cell;
                    return sheet.Cells.TryGetValue(row * columns + column, out cell) ? cell.Text : "";
                }, rows, columns);
                engine.PrepareSpills(dynamicKeys, (row, column) =>
                {
                    CellSnapshot cell;
                    return sheet.Cells.TryGetValue(row * columns + column, out cell) &&
                        !String.IsNullOrEmpty(cell.Text) || sheet.Merges.Any(m => m.Contains(column, row));
                });
                spillDimensions = engine.SpillDimensions();
            }
            if (!sheet.TabColor.IsEmpty || sheet.Print.FitToOnePage)
            {
                var properties = new XElement(S + "sheetPr");
                if (!sheet.TabColor.IsEmpty)
                    properties.Add(new XElement(S + "tabColor",
                        new XAttribute("rgb", "FF" + sheet.TabColor.R.ToString("X2") +
                            sheet.TabColor.G.ToString("X2") + sheet.TabColor.B.ToString("X2"))));
                if (sheet.Print.FitToOnePage)
                    properties.Add(new XElement(S + "pageSetUpPr", new XAttribute("fitToPage", 1)));
                root.Add(properties);
            }
            var view = new XElement(S + "sheetView", new XAttribute("workbookViewId", 0),
                new XAttribute("showGridLines", sheet.ShowGridlines ? 1 : 0),
                new XAttribute("showRowColHeaders", sheet.ShowHeadings ? 1 : 0),
                new XAttribute("showFormulas", sheet.FormulaView ? 1 : 0));
            if (sheet.ViewMode == "pageBreakPreview") view.SetAttributeValue("view", "pageBreakPreview");
            if (sheet.FreezeRow > 0 || sheet.FreezeColumn > 0)
                view.Add(new XElement(S + "pane", new XAttribute("xSplit", sheet.FreezeColumn),
                    new XAttribute("ySplit", sheet.FreezeRow),
                    new XAttribute("topLeftCell", ColumnName(sheet.FreezeColumn) + (sheet.FreezeRow + 1)),
                    new XAttribute("state", "frozen")));
            else if (sheet.SplitX > 0 || sheet.SplitY > 0)
                view.Add(new XElement(S + "pane", new XAttribute("xSplit", sheet.SplitX),
                    new XAttribute("ySplit", sheet.SplitY), new XAttribute("state", "split")));
            root.Add(new XElement(S + "sheetViews", view));
            if (sheet.ColumnWidths.Count > 0 || sheet.HiddenColumns.Count > 0)
            {
                var cols = new XElement(S + "cols");
                foreach (int index in sheet.ColumnWidths.Keys.Union(sheet.HiddenColumns).OrderBy(x => x))
                {
                    int width;
                    if (!sheet.ColumnWidths.TryGetValue(index, out width)) width = 120;
                    var col = new XElement(S + "col", new XAttribute("min", index + 1),
                        new XAttribute("max", index + 1), new XAttribute("width",
                            Math.Max(1, (width - 5) / 7.0).ToString(CultureInfo.InvariantCulture)),
                        new XAttribute("customWidth", 1));
                    if (sheet.HiddenColumns.Contains(index)) col.SetAttributeValue("hidden", 1);
                    cols.Add(col);
                }
                root.Add(cols);
            }
            var data = new XElement(S + "sheetData");
            var rowCells = sheet.Cells.GroupBy(x => x.Key / columns)
                .ToDictionary(group => group.Key, group => group.OrderBy(x => x.Key).ToList());
            for (int r = 0; r < rows; r++)
            {
                List<KeyValuePair<int, CellSnapshot>> entries;
                if (!rowCells.TryGetValue(r, out entries)) entries = new List<KeyValuePair<int, CellSnapshot>>();
                if (entries.Count == 0 && !sheet.RowHeights.ContainsKey(r) && !sheet.HiddenRows.Contains(r)) continue;
                var row = new XElement(S + "row", new XAttribute("r", r + 1));
                if (sheet.HiddenRows.Contains(r)) row.SetAttributeValue("hidden", 1);
                if (sheet.RowHeights.ContainsKey(r)) { row.SetAttributeValue("ht", (sheet.RowHeights[r] * 72.0 / 96).ToString(CultureInfo.InvariantCulture)); row.SetAttributeValue("customHeight", 1); }
                foreach (var pair in entries)
                {
                    var cell = pair.Value;
                    string address = ColumnName(pair.Key % columns) + (r + 1);
                    var element = new XElement(S + "c", new XAttribute("r", address));
                    int styleIndex = styles.Index(cell);
                    if (styleIndex > 0) element.SetAttributeValue("s", styleIndex);
                    if (cell.Text.StartsWith("=", StringComparison.Ordinal))
                    {
                        var formula = new XElement(S + "f", ExportFormula(cell.Text.Substring(1)));
                        int[] size;
                        if (spillDimensions.TryGetValue(pair.Key, out size))
                        {
                            formula.SetAttributeValue("t", "array");
                            formula.SetAttributeValue("ref", address + ":" +
                                ColumnName(pair.Key % columns + size[1] - 1) +
                                (r + size[0]).ToString(CultureInfo.InvariantCulture));
                            formula.SetAttributeValue("aca", 1);
                        }
                        element.Add(formula);
                    }
                    else
                    {
                        double number;
                        if (double.TryParse(cell.Text, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) element.Add(new XElement(S + "v", cell.Text));
                        else { element.SetAttributeValue("t", "inlineStr"); element.Add(new XElement(S + "is", new XElement(S + "t", cell.Text))); }
                    }
                    row.Add(element);
                }
                data.Add(row);
            }
            root.Add(data);
            if (sheet.Protected)
                root.Add(new XElement(S + "sheetProtection", new XAttribute("sheet", 1)));
            var filterRules = new List<FilterCriterion>(sheet.Filters);
            if (sheet.FilterColumn >= 0 && sheet.FilterColumn < columns &&
                !filterRules.Any(f => f.Column == sheet.FilterColumn))
                filterRules.Add(new FilterCriterion { Column = sheet.FilterColumn, Operator = "Contains", Value1 = sheet.FilterValue });
            if (filterRules.Count > 0)
            {
                var filterElement = new XElement(S + "autoFilter", new XAttribute("ref", "A1:" + ColumnName(columns - 1) + rows));
                foreach (FilterCriterion criterion in filterRules.Where(f => f.Column >= 0 && f.Column < columns))
                {
                    if (criterion.Operator == "One Of")
                    {
                        var choices = criterion.Value1.Split('\n').Select(v => v.TrimEnd('\r')).ToList();
                        var selected = new XElement(S + "filters");
                        if (choices.Remove("")) selected.SetAttributeValue("blank", 1);
                        foreach (string choice in choices) selected.Add(new XElement(S + "filter", new XAttribute("val", choice)));
                        filterElement.Add(new XElement(S + "filterColumn", new XAttribute("colId", criterion.Column), selected));
                        continue;
                    }
                    string operation = criterion.Operator == "Greater" ? "greaterThan" : criterion.Operator == "Less" ? "lessThan" :
                        criterion.Operator == "Nonblank" ? "notEqual" : "equal";
                    string value = criterion.Operator == "Contains" ? "*" + criterion.Value1 + "*" :
                        criterion.Operator == "Begins With" ? criterion.Value1 + "*" :
                        criterion.Operator == "Blank" || criterion.Operator == "Nonblank" ? "" : criterion.Value1;
                    var customs = new XElement(S + "customFilters",
                        new XElement(S + "customFilter", new XAttribute("operator", criterion.Operator == "Between" ? "greaterThanOrEqual" : operation), new XAttribute("val", value)));
                    if (criterion.Operator == "Between")
                    {
                        customs.SetAttributeValue("and", 1);
                        customs.Add(new XElement(S + "customFilter", new XAttribute("operator", "lessThanOrEqual"), new XAttribute("val", criterion.Value2)));
                    }
                    filterElement.Add(new XElement(S + "filterColumn", new XAttribute("colId", criterion.Column), customs));
                }
                root.Add(filterElement);
            }
            if (sheet.Merges.Count > 0)
            {
                var merges = new XElement(S + "mergeCells", new XAttribute("count", sheet.Merges.Count));
                foreach (var rect in sheet.Merges) merges.Add(new XElement(S + "mergeCell", new XAttribute("ref", ColumnName(rect.X) + (rect.Y + 1) + ":" + ColumnName(rect.Right - 1) + rect.Bottom)));
                root.Add(merges);
            }
            int priority = 1;
            foreach (ConditionalRule rule in sheet.Rules)
            {
                Rectangle rect = rule.Range;
                if (rect.X < 0 || rect.Y < 0 || rect.Right > columns || rect.Bottom > rows) continue;
                string reference = ColumnName(rect.X) + (rect.Y + 1) + ":" + ColumnName(rect.Right - 1) + rect.Bottom;
                string type = rule.Kind == "Duplicate" ? "duplicateValues" : rule.Kind == "Unique" ? "uniqueValues" :
                    rule.Kind == "Text Contains" ? "containsText" : rule.Kind == "Blank" ? "containsBlanks" :
                    rule.Kind == "Formula" ? "expression" : rule.Kind == "Color Scale" ? "colorScale" :
                    rule.Kind == "Data Bar" ? "dataBar" : rule.Kind == "Icon Set" ? "iconSet" : "cellIs";
                var entry = new XElement(S + "cfRule", new XAttribute("type", type), new XAttribute("priority", priority++));
                if (type == "cellIs")
                {
                    entry.SetAttributeValue("operator", rule.Kind == "Less" ? "lessThan" : rule.Kind == "Equal" ? "equal" :
                        rule.Kind == "Between" ? "between" : "greaterThan");
                    entry.Add(new XElement(S + "formula", string.IsNullOrEmpty(rule.Value1) ?
                        rule.Threshold.ToString(CultureInfo.InvariantCulture) : rule.Value1));
                    if (rule.Kind == "Between") entry.Add(new XElement(S + "formula", rule.Value2));
                }
                else if (type == "expression") entry.Add(new XElement(S + "formula", rule.Value1.TrimStart('=')));
                else if (type == "containsText")
                {
                    entry.SetAttributeValue("text", rule.Value1);
                    entry.Add(new XElement(S + "formula", "NOT(ISERROR(SEARCH(\"" + rule.Value1.Replace("\"", "\"\"") + "\"," + ColumnName(rect.X) + (rect.Y + 1) + ")))") );
                }
                if (type == "colorScale")
                    entry.Add(new XElement(S + "colorScale",
                        new XElement(S + "cfvo", new XAttribute("type", "min")),
                        new XElement(S + "cfvo", new XAttribute("type", "max")),
                        new XElement(S + "color", new XAttribute("rgb", (rule.Color2.IsEmpty ? Color.White : rule.Color2).ToArgb().ToString("X8", CultureInfo.InvariantCulture))),
                        new XElement(S + "color", new XAttribute("rgb", rule.Color.ToArgb().ToString("X8", CultureInfo.InvariantCulture)))));
                if (type == "dataBar") entry.Add(new XElement(S + "dataBar",
                    new XElement(S + "cfvo", new XAttribute("type", "min")),
                    new XElement(S + "cfvo", new XAttribute("type", "max")),
                    new XElement(S + "color", new XAttribute("rgb", rule.Color.ToArgb().ToString("X8", CultureInfo.InvariantCulture)))));
                if (type == "iconSet") entry.Add(new XElement(S + "iconSet", new XAttribute("iconSet", "3TrafficLights1"),
                    new XElement(S + "cfvo", new XAttribute("type", "percent"), new XAttribute("val", 0)),
                    new XElement(S + "cfvo", new XAttribute("type", "percent"), new XAttribute("val", 33)),
                    new XElement(S + "cfvo", new XAttribute("type", "percent"), new XAttribute("val", 67))));
                if (type != "colorScale" && type != "dataBar" && type != "iconSet")
                    entry.SetAttributeValue("dxfId", styles.DifferentialIndex(rule.Color));
                root.Add(new XElement(S + "conditionalFormatting", new XAttribute("sqref", reference), entry));
            }
            if (sheet.Validations.Count > 0)
            {
                var entries = new XElement(S + "dataValidations");
                foreach (ValidationRule rule in sheet.Validations)
                {
                    if (rule.Range.IsEmpty || rule.Range.Right > columns || rule.Range.Bottom > rows) continue;
                    string csv = string.Join(",", rule.Choices.ToArray()).Replace("\"", "\"\"");
                    if (rule.Kind == "List" && csv.Length > 250) continue;
                    string type = rule.Kind == "Whole Number" ? "whole" : rule.Kind == "Decimal" ? "decimal" :
                        rule.Kind == "Date" ? "date" : rule.Kind == "Time" ? "time" :
                        rule.Kind == "Text Length" ? "textLength" : rule.Kind == "Custom Formula" ? "custom" : "list";
                    string operation = rule.Operator == "Not Between" ? "notBetween" : rule.Operator == "Equals" ? "equal" :
                        rule.Operator == "Not Equals" ? "notEqual" : rule.Operator == "Greater" ? "greaterThan" :
                        rule.Operator == "Greater Or Equal" ? "greaterThanOrEqual" : rule.Operator == "Less" ? "lessThan" :
                        rule.Operator == "Less Or Equal" ? "lessThanOrEqual" : "between";
                    var entry = new XElement(S + "dataValidation", new XAttribute("type", type),
                        new XAttribute("allowBlank", rule.AllowBlank ? 1 : 0), new XAttribute("showDropDown", 0),
                        new XAttribute("sqref", RangeAddress(rule.Range)));
                    if (type != "list" && type != "custom") entry.SetAttributeValue("operator", operation);
                    if (!string.IsNullOrEmpty(rule.InputTitle) || !string.IsNullOrEmpty(rule.InputMessage))
                    { entry.SetAttributeValue("showInputMessage", 1); entry.SetAttributeValue("promptTitle", rule.InputTitle); entry.SetAttributeValue("prompt", rule.InputMessage); }
                    if (!string.IsNullOrEmpty(rule.ErrorTitle) || !string.IsNullOrEmpty(rule.ErrorMessage))
                    { entry.SetAttributeValue("showErrorMessage", 1); entry.SetAttributeValue("errorTitle", rule.ErrorTitle); entry.SetAttributeValue("error", rule.ErrorMessage); }
                    entry.SetAttributeValue("errorStyle", rule.ErrorStyle == "Warning" ? "warning" : rule.ErrorStyle == "Information" ? "information" : "stop");
                    entry.Add(new XElement(S + "formula1", type == "list" ? "\"" + csv + "\"" :
                        type == "custom" ? rule.Value1.TrimStart('=') : ValidationFormula(rule.Kind, rule.Value1)));
                    if (rule.Value2.Length > 0 && type != "list" && type != "custom") entry.Add(new XElement(S + "formula2", ValidationFormula(rule.Kind, rule.Value2)));
                    entries.Add(entry);
                }
                entries.SetAttributeValue("count", entries.Elements().Count());
                if (entries.HasElements) root.Add(entries);
            }
            var hyperlinks = new XElement(S + "hyperlinks");
            foreach (var cell in sheet.Cells)
            {
                string link = cell.Value.Extras == null ? "" : cell.Value.Extras.Hyperlink;
                if (string.IsNullOrWhiteSpace(link)) continue;
                var item = new XElement(S + "hyperlink", new XAttribute("ref",
                    ColumnName(cell.Key % columns) + (cell.Key / columns + 1)));
                if (link.StartsWith("#", StringComparison.Ordinal))
                    item.SetAttributeValue("location", link.Substring(1));
                else
                    item.SetAttributeValue(R + "id", hyperlinkIds[cell.Key]);
                hyperlinks.Add(item);
            }
            if (hyperlinks.HasElements) root.Add(hyperlinks);
            PrintSettings print = sheet.Print;
            root.Add(new XElement(S + "printOptions", new XAttribute("gridLines", print.Gridlines ? 1 : 0)));
            root.Add(new XElement(S + "pageMargins",
                new XAttribute("left", (print.MarginLeft / 100.0).ToString(CultureInfo.InvariantCulture)),
                new XAttribute("right", (print.MarginRight / 100.0).ToString(CultureInfo.InvariantCulture)),
                new XAttribute("top", (print.MarginTop / 100.0).ToString(CultureInfo.InvariantCulture)),
                new XAttribute("bottom", (print.MarginBottom / 100.0).ToString(CultureInfo.InvariantCulture)),
                new XAttribute("header", "0.3"), new XAttribute("footer", "0.3")));
            root.Add(new XElement(S + "pageSetup",
                new XAttribute("paperSize", print.Paper == "Letter" ? 1 : print.Paper == "Legal" ? 5 :
                    print.Paper == "A3" ? 8 : 9),
                new XAttribute("orientation", print.Landscape ? "landscape" : "portrait"),
                new XAttribute("scale", Math.Max(10, Math.Min(400, print.Scale))),
                new XAttribute("fitToWidth", print.FitToOnePage ? 1 : 0),
                new XAttribute("fitToHeight", print.FitToOnePage ? 1 : 0)));
            root.Add(new XElement(S + "headerFooter",
                new XElement(S + "oddHeader", "&C" + print.Header),
                new XElement(S + "oddFooter", "&C" + print.Footer)));
            if (print.PageBreakRows.Count > 0)
            {
                var breaks = new XElement(S + "rowBreaks", new XAttribute("count", print.PageBreakRows.Count),
                    new XAttribute("manualBreakCount", print.PageBreakRows.Count));
                foreach (int row in print.PageBreakRows)
                    breaks.Add(new XElement(S + "brk", new XAttribute("id", row),
                        new XAttribute("min", 0), new XAttribute("max", columns - 1),
                        new XAttribute("man", 1)));
                root.Add(breaks);
            }
            if (!string.IsNullOrEmpty(drawingRelation))
                root.Add(new XElement(S + "drawing", new XAttribute(R + "id", drawingRelation)));
            if (!string.IsNullOrEmpty(vmlRelation))
                root.Add(new XElement(S + "legacyDrawing", new XAttribute(R + "id", vmlRelation)));
            if (tableIds.Count > 0)
            {
                var parts = new XElement(S + "tableParts", new XAttribute("count", tableIds.Count));
                foreach (string id in tableIds) parts.Add(new XElement(S + "tablePart", new XAttribute(R + "id", id)));
                root.Add(parts);
            }
            return new XDocument(root);
        }

        private static XDocument BuildTable(TableDefinition table, SheetSnapshot sheet, int id, int columns)
        {
            var autoFilter = new XElement(S + "autoFilter", new XAttribute("ref", RangeAddress(table.Range)));
            foreach (FilterCriterion criterion in table.Filters)
            {
                if (criterion.Column < table.Range.Left || criterion.Column >= table.Range.Right) continue;
                if (criterion.Operator == "One Of")
                {
                    var choices = criterion.Value1.Split('\n').Select(v => v.TrimEnd('\r')).ToList();
                    var selected = new XElement(S + "filters");
                    if (choices.Remove("")) selected.SetAttributeValue("blank", 1);
                    foreach (string choice in choices) selected.Add(new XElement(S + "filter", new XAttribute("val", choice)));
                    autoFilter.Add(new XElement(S + "filterColumn",
                        new XAttribute("colId", criterion.Column - table.Range.Left), selected));
                    continue;
                }
                string op = criterion.Operator == "Greater" ? "greaterThan" : criterion.Operator == "Less" ? "lessThan" :
                    criterion.Operator == "Between" ? "greaterThanOrEqual" : criterion.Operator == "Nonblank" ? "notEqual" : "equal";
                string value = criterion.Operator == "Contains" ? "*" + criterion.Value1 + "*" :
                    criterion.Operator == "Begins With" ? criterion.Value1 + "*" :
                    criterion.Operator == "Blank" || criterion.Operator == "Nonblank" ? "" : criterion.Value1;
                var customs = new XElement(S + "customFilters", new XElement(S + "customFilter",
                    new XAttribute("operator", op), new XAttribute("val", value)));
                if (criterion.Operator == "Between")
                {
                    customs.SetAttributeValue("and", 1);
                    customs.Add(new XElement(S + "customFilter", new XAttribute("operator", "lessThanOrEqual"),
                        new XAttribute("val", criterion.Value2)));
                }
                autoFilter.Add(new XElement(S + "filterColumn",
                    new XAttribute("colId", criterion.Column - table.Range.Left), customs));
            }
            var header = new XElement(S + "tableColumns", new XAttribute("count", table.Range.Width));
            for (int c = table.Range.Left; c < table.Range.Right; c++)
            {
                CellSnapshot cell;
                string name = sheet.Cells.TryGetValue(table.Range.Top * columns + c, out cell) ? cell.Text : "";
                if (string.IsNullOrEmpty(name)) name = "Column" + (c - table.Range.Left + 1);
                var column = new XElement(S + "tableColumn", new XAttribute("id", c - table.Range.Left + 1),
                    new XAttribute("name", name));
                string formula;
                if (table.CalculatedColumns.TryGetValue(c, out formula))
                    column.Add(new XElement(S + "calculatedColumnFormula", formula.TrimStart('=')));
                header.Add(column);
            }
            return new XDocument(new XElement(S + "table", new XAttribute("id", id),
                new XAttribute("name", table.Name), new XAttribute("displayName", table.Name),
                new XAttribute("ref", RangeAddress(table.Range)), new XAttribute("headerRowCount", table.HeaderRow ? 1 : 0),
                new XAttribute("totalsRowCount", table.TotalRow ? 1 : 0),
                table.Filter ? autoFilter : null, header,
                new XElement(S + "tableStyleInfo", new XAttribute("name", table.Style),
                    new XAttribute("showFirstColumn", 0), new XAttribute("showLastColumn", 0),
                    new XAttribute("showRowStripes", table.BandedRows ? 1 : 0),
                    new XAttribute("showColumnStripes", table.BandedColumns ? 1 : 0))));
        }
    }
}
