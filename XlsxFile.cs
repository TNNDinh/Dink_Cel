using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DinkCel
{
    internal static class XlsxFile
    {
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

        private static int ColumnIndex(string address)
        {
            int result = 0;
            foreach (char c in address) { if (!char.IsLetter(c)) break; result = result * 26 + char.ToUpperInvariant(c) - 'A' + 1; }
            return result - 1;
        }

        private static int RowIndex(string address)
        {
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
                    styles = XlsxStyles.Read(styleDocument);
                    differentialColors = XlsxStyles.ReadDifferentialColors(styleDocument);
                }
                var result = new WorkbookSnapshot(); result.Sheets.Clear();
                foreach (var sheetInfo in book.Descendants(S + "sheet"))
                {
                    string id = (string)sheetInfo.Attribute(R + "id");
                    if (!paths.ContainsKey(id)) continue;
                    var document = ReadXml(zip, paths[id]);
                    var sheet = new SheetSnapshot { Name = (string)sheetInfo.Attribute("name") ?? "Sheet" };
                    var pane = document.Descendants(S + "pane").FirstOrDefault();
                    if (pane != null)
                    {
                        sheet.FreezeColumn = Math.Min(columns, (int?)pane.Attribute("xSplit") ?? 0);
                        sheet.FreezeRow = Math.Min(rows, (int?)pane.Attribute("ySplit") ?? 0);
                    }
                    foreach (var col in document.Descendants(S + "col"))
                    {
                        int from = (int?)col.Attribute("min") ?? 0, to = (int?)col.Attribute("max") ?? 0;
                        double width = (double?)col.Attribute("width") ?? 0;
                        for (int c = from; c <= to && c <= columns; c++) if (c > 0) sheet.ColumnWidths[c - 1] = Math.Max(20, (int)(width * 7 + 5));
                    }
                    foreach (var row in document.Descendants(S + "sheetData").Elements(S + "row"))
                    {
                        int r = ((int?)row.Attribute("r") ?? 0) - 1;
                        double height = (double?)row.Attribute("ht") ?? 0;
                        if (r >= 0 && r < rows && height > 0) sheet.RowHeights[r] = Math.Max(2, (int)(height * 96 / 72));
                        foreach (var cell in row.Elements(S + "c"))
                        {
                            string address = (string)cell.Attribute("r") ?? "";
                            int col = ColumnIndex(address), line = RowIndex(address);
                            if (col < 0 || col >= columns || line < 0 || line >= rows)
                                throw new InvalidDataException("XLSX contains cells outside the 200 x 26 grid: " + address);
                            string type = (string)cell.Attribute("t") ?? "";
                            string value = (string)cell.Element(S + "v") ?? "";
                            if (type == "s") { int index; if (int.TryParse(value, out index) && index >= 0 && index < strings.Count) value = strings[index]; }
                            else if (type == "inlineStr") value = string.Concat(cell.Descendants(S + "t").Select(x => x.Value));
                            else if (type == "b") value = value == "1" ? "TRUE" : "FALSE";
                            var formula = cell.Element(S + "f");
                            if (formula != null && !string.IsNullOrEmpty(formula.Value)) value = "=" + formula.Value;
                            var snapshot = new CellSnapshot { Text = value };
                            int styleIndex = (int?)cell.Attribute("s") ?? 0;
                            if (styleIndex >= 0 && styleIndex < styles.Count)
                            {
                                CellSnapshot style = styles[styleIndex];
                                snapshot.NumberFormat = style.NumberFormat;
                                snapshot.HasFont = style.HasFont;
                                snapshot.FontStyle = style.FontStyle;
                                snapshot.FontSize = style.FontSize;
                                snapshot.ForeColor = style.ForeColor;
                                snapshot.BackColor = style.BackColor;
                                snapshot.Alignment = style.Alignment;
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
                                if ((string)rule.Attribute("type") != "cellIs" || (string)rule.Attribute("operator") != "greaterThan") continue;
                                int dxf = (int?)rule.Attribute("dxfId") ?? -1;
                                double threshold;
                                if (dxf < 0 || dxf >= differentialColors.Count || differentialColors[dxf].IsEmpty ||
                                    !double.TryParse((string)rule.Element(S + "formula"), NumberStyles.Float, CultureInfo.InvariantCulture, out threshold)) continue;
                                sheet.Rules.Add(new ConditionalRule { Range = new Rectangle(x, y, right - x + 1, bottom - y + 1), Threshold = threshold, Color = differentialColors[dxf] });
                            }
                        }
                    }
                    XElement filter = document.Descendants(S + "autoFilter").FirstOrDefault();
                    if (filter != null)
                    {
                        XElement filterColumn = filter.Descendants(S + "filterColumn").FirstOrDefault();
                        if (filterColumn != null)
                        {
                            sheet.FilterColumn = (int?)filterColumn.Attribute("colId") ?? -1;
                            XElement custom = filterColumn.Descendants(S + "customFilter").FirstOrDefault();
                            if (custom != null) sheet.FilterValue = ((string)custom.Attribute("val") ?? "").Trim('*');
                        }
                    }
                    result.Sheets.Add(sheet);
                }
                if (result.Sheets.Count == 0) result.Sheets.Add(new SheetSnapshot());
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
                    for (int i = 0; i < book.Sheets.Count; i++)
                    {
                        int number = i + 1;
                        string part = "xl/worksheets/sheet" + number + ".xml";
                        types.Add(new XElement(C + "Override", new XAttribute("PartName", "/" + part), new XAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")));
                        sheetsElement.Add(new XElement(S + "sheet", new XAttribute("name", book.Sheets[i].Name), new XAttribute("sheetId", number), new XAttribute(R + "id", "rId" + number)));
                        relationships.Add(new XElement(P + "Relationship", new XAttribute("Id", "rId" + number), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"), new XAttribute("Target", "worksheets/sheet" + number + ".xml")));
                        WriteXml(zip, part, BuildSheet(book.Sheets[i], rows, columns, styleCatalog));
                    }
                    relationships.Add(new XElement(P + "Relationship", new XAttribute("Id", "rIdStyles"), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"), new XAttribute("Target", "styles.xml")));
                    WriteXml(zip, "[Content_Types].xml", new XDocument(types));
                    WriteXml(zip, "_rels/.rels", new XDocument(new XElement(P + "Relationships", new XElement(P + "Relationship", new XAttribute("Id", "rId1"), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"), new XAttribute("Target", "xl/workbook.xml")))));
                    WriteXml(zip, "xl/workbook.xml", new XDocument(new XElement(S + "workbook", new XAttribute(XNamespace.Xmlns + "r", R), sheetsElement)));
                    WriteXml(zip, "xl/_rels/workbook.xml.rels", new XDocument(relationships));
                    WriteXml(zip, "xl/styles.xml", styleCatalog.Document());
                }
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        private static XDocument BuildSheet(SheetSnapshot sheet, int rows, int columns, XlsxStyles styles)
        {
            var root = new XElement(S + "worksheet");
            if (sheet.FreezeRow > 0 || sheet.FreezeColumn > 0)
                root.Add(new XElement(S + "sheetViews", new XElement(S + "sheetView", new XAttribute("workbookViewId", 0),
                    new XElement(S + "pane", new XAttribute("xSplit", sheet.FreezeColumn), new XAttribute("ySplit", sheet.FreezeRow),
                        new XAttribute("topLeftCell", ColumnName(sheet.FreezeColumn) + (sheet.FreezeRow + 1)), new XAttribute("state", "frozen")))));
            if (sheet.ColumnWidths.Count > 0)
            {
                var cols = new XElement(S + "cols");
                foreach (var pair in sheet.ColumnWidths.OrderBy(x => x.Key)) cols.Add(new XElement(S + "col", new XAttribute("min", pair.Key + 1), new XAttribute("max", pair.Key + 1), new XAttribute("width", Math.Max(1, (pair.Value - 5) / 7.0).ToString(CultureInfo.InvariantCulture)), new XAttribute("customWidth", 1)));
                root.Add(cols);
            }
            var data = new XElement(S + "sheetData");
            for (int r = 0; r < rows; r++)
            {
                var entries = sheet.Cells.Where(x => x.Key / columns == r).OrderBy(x => x.Key).ToList();
                if (entries.Count == 0 && !sheet.RowHeights.ContainsKey(r)) continue;
                var row = new XElement(S + "row", new XAttribute("r", r + 1));
                if (sheet.RowHeights.ContainsKey(r)) { row.SetAttributeValue("ht", (sheet.RowHeights[r] * 72.0 / 96).ToString(CultureInfo.InvariantCulture)); row.SetAttributeValue("customHeight", 1); }
                foreach (var pair in entries)
                {
                    var cell = pair.Value;
                    string address = ColumnName(pair.Key % columns) + (r + 1);
                    var element = new XElement(S + "c", new XAttribute("r", address));
                    int styleIndex = styles.Index(cell);
                    if (styleIndex > 0) element.SetAttributeValue("s", styleIndex);
                    if (cell.Text.StartsWith("=", StringComparison.Ordinal)) element.Add(new XElement(S + "f", cell.Text.Substring(1)));
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
            if (sheet.FilterColumn >= 0 && sheet.FilterColumn < columns)
                root.Add(new XElement(S + "autoFilter", new XAttribute("ref", "A1:" + ColumnName(columns - 1) + rows),
                    new XElement(S + "filterColumn", new XAttribute("colId", sheet.FilterColumn),
                        new XElement(S + "customFilters", new XElement(S + "customFilter",
                            new XAttribute("operator", "equal"), new XAttribute("val", "*" + sheet.FilterValue + "*"))))));
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
                root.Add(new XElement(S + "conditionalFormatting", new XAttribute("sqref", reference),
                    new XElement(S + "cfRule", new XAttribute("type", "cellIs"), new XAttribute("operator", "greaterThan"),
                        new XAttribute("dxfId", styles.DifferentialIndex(rule.Color)), new XAttribute("priority", priority++),
                        new XElement(S + "formula", rule.Threshold.ToString(CultureInfo.InvariantCulture)))));
            }
            return new XDocument(root);
        }
    }
}
