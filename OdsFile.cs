using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace DinkCel
{
    internal static class OdsFile
    {
        private static readonly XNamespace Office = "urn:oasis:names:tc:opendocument:xmlns:office:1.0";
        private static readonly XNamespace Table = "urn:oasis:names:tc:opendocument:xmlns:table:1.0";
        private static readonly XNamespace Text = "urn:oasis:names:tc:opendocument:xmlns:text:1.0";
        private static readonly XNamespace Manifest = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
        private static readonly XNamespace Formula = "urn:oasis:names:tc:opendocument:xmlns:of:1.2";
        private static readonly Regex ExcelReference = new Regex(@"(?<![A-Za-z0-9_\.])(?:(?<sheet>'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_]*)!)?(?<first>\$?[A-Z]{1,2}\$?\d{1,3})(?::(?<last>\$?[A-Z]{1,2}\$?\d{1,3}))?", RegexOptions.Compiled);
        private static readonly Regex OdfReference = new Regex(@"\[(?<first>(?:'[^']+'|[^\.]*)?\.\$?[A-Z]{1,2}\$?\d{1,3})(?::(?<last>(?:'[^']+'|[^\.]*)?\.\$?[A-Z]{1,2}\$?\d{1,3}))?\]", RegexOptions.Compiled);

        private static string TransformOutsideStrings(string formula, Func<string, string> transform)
        {
            var result = new StringBuilder();
            var segment = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < formula.Length; i++)
            {
                char c = formula[i];
                if (c == '"')
                {
                    if (!quoted) { result.Append(transform(segment.ToString())); segment.Clear(); quoted = true; }
                    else if (i + 1 < formula.Length && formula[i + 1] == '"') { result.Append("\"\""); i++; continue; }
                    else quoted = false;
                    result.Append(c);
                }
                else if (quoted) result.Append(c);
                else segment.Append(c);
            }
            result.Append(transform(segment.ToString()));
            return result.ToString();
        }

        private static string ToOdfFormula(string excel)
        {
            string body = excel.StartsWith("=", StringComparison.Ordinal) ? excel.Substring(1) : excel;
            return "of:=" + TransformOutsideStrings(body, segment => ExcelReference.Replace(segment, match =>
            {
                string sheet = match.Groups["sheet"].Success ? match.Groups["sheet"].Value : "";
                string prefix = sheet.Length == 0 ? "." : sheet + ".";
                return "[" + prefix + match.Groups["first"].Value +
                    (match.Groups["last"].Success ? ":" + prefix + match.Groups["last"].Value : "") + "]";
            }));
        }

        private static string FromOdfFormula(string odf)
        {
            string body = odf.StartsWith("of:=", StringComparison.Ordinal) ? odf.Substring(4) :
                odf.StartsWith("oooc:=", StringComparison.Ordinal) ? odf.Substring(6) : odf.TrimStart('=');
            return "=" + TransformOutsideStrings(body, segment => OdfReference.Replace(segment, match =>
            {
                string first = match.Groups["first"].Value;
                int dot = first.LastIndexOf('.');
                string sheet = first.Substring(0, dot).TrimStart('$');
                string address = first.Substring(dot + 1);
                string prefix = sheet.Length == 0 ? "" : sheet + "!";
                if (!match.Groups["last"].Success) return prefix + address;
                string last = match.Groups["last"].Value;
                return prefix + address + ":" + last.Substring(last.LastIndexOf('.') + 1);
            }));
        }

        private static void WriteXml(ZipArchive zip, string name, XDocument document)
        {
            var entry = zip.CreateEntry(name, CompressionLevel.Optimal);
            using (var stream = entry.Open())
            using (var writer = new StreamWriter(stream, new UTF8Encoding(false))) document.Save(writer);
        }

        public static WorkbookSnapshot Read(string path, int rows, int columns)
        {
            using (var zip = ZipFile.OpenRead(path))
            {
                var entry = zip.GetEntry("content.xml");
                if (entry == null) throw new InvalidDataException("ODS is missing content.xml.");
                XDocument document;
                using (var stream = entry.Open()) document = XDocument.Load(stream);
                var result = new WorkbookSnapshot(); result.Sheets.Clear();
                foreach (XElement table in document.Descendants(Table + "table"))
                {
                    var sheet = new SheetSnapshot { Name = (string)table.Attribute(Table + "name") ?? "Sheet" };
                    int r = 0;
                    foreach (XElement row in table.Elements(Table + "table-row"))
                    {
                        int rowRepeat = Math.Max(1, (int?)row.Attribute(Table + "number-rows-repeated") ?? 1);
                        if (rowRepeat > rows - r && row.Elements(Table + "table-cell").Any(cell =>
                            !string.IsNullOrEmpty((string)cell.Attribute(Table + "formula")) ||
                            ReadCellText(cell).Length > 0))
                            throw new InvalidDataException("ODS contains data outside the 200 x 26 grid.");
                        for (int repeatRow = 0; repeatRow < rowRepeat && r < rows; repeatRow++, r++)
                        {
                            int c = 0;
                            foreach (XElement cell in row.Elements())
                            {
                                if (cell.Name != Table + "table-cell" && cell.Name != Table + "covered-table-cell") continue;
                                int repeat = Math.Max(1, (int?)cell.Attribute(Table + "number-columns-repeated") ?? 1);
                                if (cell.Name == Table + "covered-table-cell") { c += repeat; continue; }
                                string formula = (string)cell.Attribute(Table + "formula");
                                string value = formula == null ? ReadCellText(cell) : FromOdfFormula(formula);
                                int spanColumns = (int?)cell.Attribute(Table + "number-columns-spanned") ?? 1;
                                int spanRows = (int?)cell.Attribute(Table + "number-rows-spanned") ?? 1;
                                if (repeat > columns - c && value.Length > 0)
                                    throw new InvalidDataException("ODS contains data outside the 200 x 26 grid.");
                                int visible = Math.Max(0, Math.Min(repeat, columns - c));
                                for (int k = 0; k < visible; k++, c++)
                                {
                                    string text = formula == null || (repeatRow == 0 && k == 0) ? value :
                                        FormulaEngine.ShiftReferences(value, repeatRow, k, rows, columns);
                                    if (text.Length > 0) sheet.Cells[r * columns + c] = new CellSnapshot { Text = text };
                                    if (spanColumns > 1 || spanRows > 1)
                                        sheet.Merges.Add(new Rectangle(c, r, spanColumns, spanRows));
                                }
                                c += repeat - visible;
                            }
                        }
                    }
                    result.Sheets.Add(sheet);
                }
                foreach (XElement named in document.Descendants(Table + "named-range"))
                {
                    string name = (string)named.Attribute(Table + "name") ?? "";
                    string address = (string)named.Attribute(Table + "cell-range-address") ?? "";
                    string[] ends = address.Split(':');
                    if (name.Length == 0 || ends.Length == 0) continue;
                    int column, row, lastColumn, lastRow;
                    string sheetName;
                    if (!ParseNamedAddress(ends[0], out sheetName, out column, out row)) continue;
                    string otherName;
                    if (ends.Length > 1)
                    { if (!ParseNamedAddress(ends[1], out otherName, out lastColumn, out lastRow)) continue; }
                    else { lastColumn = column; lastRow = row; }
                    if (column >= 0 && row >= 0 && lastColumn < columns && lastRow < rows)
                        result.NamedRanges.Add(new NamedRange { Name = name, Sheet = sheetName,
                            Range = new Rectangle(column, row, lastColumn - column + 1, lastRow - row + 1) });
                }
                if (result.Sheets.Count == 0) result.Sheets.Add(new SheetSnapshot());
                return result;
            }
        }

        private static string ReadCellText(XElement cell)
        {
            string type = (string)cell.Attribute(Office + "value-type") ?? "";
            if (type == "float" || type == "currency" || type == "percentage")
                return (string)cell.Attribute(Office + "value") ?? "";
            if (type == "boolean") return (string)cell.Attribute(Office + "boolean-value") == "true" ? "TRUE" : "FALSE";
            var lines = cell.Elements(Text + "p").Select(x => x.Value).ToArray();
            return lines.Length > 0 ? string.Join("\n", lines) : (string)cell.Attribute(Office + "string-value") ?? "";
        }

        private static string ColumnName(int column)
        {
            string text = "";
            do { text = (char)('A' + column % 26) + text; column = column / 26 - 1; } while (column >= 0);
            return text;
        }

        private static bool ParseNamedAddress(string text, out string sheet, out int column, out int row)
        {
            sheet = ""; column = row = -1;
            int dot = text.LastIndexOf('.');
            if (dot < 0) return false;
            sheet = text.Substring(0, dot).TrimStart('$').Trim('\'').Replace("''", "'");
            string cell = text.Substring(dot + 1).Replace("$", "");
            Match match = Regex.Match(cell, @"^([A-Z]+)([1-9][0-9]*)$", RegexOptions.IgnoreCase);
            if (!match.Success) return false;
            column = 0;
            foreach (char c in match.Groups[1].Value.ToUpperInvariant()) column = column * 26 + c - 'A' + 1;
            column--;
            row = int.Parse(match.Groups[2].Value) - 1;
            return true;
        }

        private static string NamedAddress(NamedRange named, int column, int row)
        {
            string sheet = Regex.IsMatch(named.Sheet, @"^[A-Za-z_][A-Za-z0-9_]*$") ? named.Sheet :
                "'" + named.Sheet.Replace("'", "''") + "'";
            return "$" + sheet + ".$" + ColumnName(column) + "$" + (row + 1);
        }

        public static void Write(string path, WorkbookSnapshot workbook, int rows, int columns)
        {
            string temporary = path + ".tmp";
            try
            {
                if (File.Exists(temporary)) File.Delete(temporary);
                using (var zip = ZipFile.Open(temporary, ZipArchiveMode.Create))
                {
                    var mime = zip.CreateEntry("mimetype", CompressionLevel.NoCompression);
                    using (var stream = mime.Open())
                    {
                        byte[] bytes = Encoding.ASCII.GetBytes("application/vnd.oasis.opendocument.spreadsheet");
                        stream.Write(bytes, 0, bytes.Length);
                    }
                    var spreadsheet = new XElement(Office + "spreadsheet");
                    if (workbook.NamedRanges.Count > 0)
                    {
                        var expressions = new XElement(Table + "named-expressions");
                        foreach (NamedRange named in workbook.NamedRanges)
                            if (!named.Range.IsEmpty)
                            {
                                string first = NamedAddress(named, named.Range.Left, named.Range.Top);
                                string last = NamedAddress(named, named.Range.Right - 1, named.Range.Bottom - 1);
                                expressions.Add(new XElement(Table + "named-range",
                                    new XAttribute(Table + "name", named.Name),
                                    new XAttribute(Table + "cell-range-address", first + ":" + last),
                                    new XAttribute(Table + "base-cell-address", first)));
                            }
                        spreadsheet.Add(expressions);
                    }
                    foreach (SheetSnapshot sheet in workbook.Sheets)
                    {
                        var table = new XElement(Table + "table", new XAttribute(Table + "name", sheet.Name));
                        int maxRow = sheet.Cells.Count == 0 ? 0 : sheet.Cells.Keys.Max() / columns;
                        foreach (Rectangle merge in sheet.Merges) maxRow = Math.Max(maxRow, merge.Bottom - 1);
                        var merged = new Dictionary<int, Rectangle>();
                        foreach (Rectangle merge in sheet.Merges)
                            for (int r = merge.Top; r < merge.Bottom; r++)
                                for (int c = merge.Left; c < merge.Right; c++)
                                    if (r >= 0 && r < rows && c >= 0 && c < columns)
                                        merged[r * columns + c] = merge;
                        for (int r = 0; r <= Math.Min(maxRow, rows - 1); r++)
                        {
                            var line = new XElement(Table + "table-row");
                            int maxColumn = sheet.Cells.Keys.Where(key => key / columns == r).Select(key => key % columns).DefaultIfEmpty(0).Max();
                            foreach (Rectangle merge in sheet.Merges)
                                if (merge.Top <= r && merge.Bottom > r) maxColumn = Math.Max(maxColumn, merge.Right - 1);
                            for (int c = 0; c <= Math.Min(maxColumn, columns - 1); c++)
                            {
                                Rectangle merge;
                                if (merged.TryGetValue(r * columns + c, out merge) && (merge.X != c || merge.Y != r))
                                { line.Add(new XElement(Table + "covered-table-cell")); continue; }
                                CellSnapshot cell;
                                string value = sheet.Cells.TryGetValue(r * columns + c, out cell) ? cell.Text ?? "" : "";
                                var element = new XElement(Table + "table-cell");
                                if (merged.ContainsKey(r * columns + c))
                                {
                                    element.SetAttributeValue(Table + "number-columns-spanned", merge.Width);
                                    element.SetAttributeValue(Table + "number-rows-spanned", merge.Height);
                                }
                                if (value.StartsWith("=", StringComparison.Ordinal))
                                {
                                    element.SetAttributeValue(Table + "formula", ToOdfFormula(value));
                                    element.SetAttributeValue(Office + "value-type", "string");
                                    element.Add(new XElement(Text + "p", value));
                                }
                                else
                                {
                                    double number;
                                    if (double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
                                    {
                                        element.SetAttributeValue(Office + "value-type", "float");
                                        element.SetAttributeValue(Office + "value", value);
                                    }
                                    else element.SetAttributeValue(Office + "value-type", "string");
                                    element.Add(new XElement(Text + "p", value));
                                }
                                line.Add(element);
                            }
                            table.Add(line);
                        }
                        spreadsheet.Add(table);
                    }
                    WriteXml(zip, "content.xml", new XDocument(new XElement(Office + "document-content",
                        new XAttribute(XNamespace.Xmlns + "office", Office),
                        new XAttribute(XNamespace.Xmlns + "table", Table),
                        new XAttribute(XNamespace.Xmlns + "text", Text),
                        new XAttribute(XNamespace.Xmlns + "of", Formula),
                        new XAttribute(Office + "version", "1.2"),
                        new XElement(Office + "body", spreadsheet))));
                    WriteXml(zip, "META-INF/manifest.xml", new XDocument(new XElement(Manifest + "manifest",
                        new XAttribute(Manifest + "version", "1.2"),
                        new XElement(Manifest + "file-entry", new XAttribute(Manifest + "full-path", "/"),
                            new XAttribute(Manifest + "media-type", "application/vnd.oasis.opendocument.spreadsheet")),
                        new XElement(Manifest + "file-entry", new XAttribute(Manifest + "full-path", "content.xml"),
                            new XAttribute(Manifest + "media-type", "text/xml")))));
                }
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }
    }
}
