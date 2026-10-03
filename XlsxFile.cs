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
                foreach (var sheetInfo in book.Descendants(S + "sheet"))
                {
                    string id = (string)sheetInfo.Attribute(R + "id");
                    if (!paths.ContainsKey(id)) continue;
                    string sheetPart = paths[id];
                    var document = ReadXml(zip, sheetPart);
                    var sheet = new SheetSnapshot { Name = (string)sheetInfo.Attribute("name") ?? "Sheet",
                        Hidden = string.Equals((string)sheetInfo.Attribute("state"), "hidden", StringComparison.OrdinalIgnoreCase) };
                    string tabRgb = (string)document.Descendants(S + "tabColor").Select(e => e.Attribute("rgb")).FirstOrDefault();
                    if (!string.IsNullOrEmpty(tabRgb) && tabRgb.Length >= 6)
                        sheet.TabColor = ColorTranslator.FromHtml("#" + tabRgb.Substring(tabRgb.Length - 6));
                    var columnStyles = new Dictionary<int, int>();
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
                        for (int c = from; c <= to && c <= columns; c++) if (c > 0)
                        {
                            if (width > 0) sheet.ColumnWidths[c - 1] = Math.Max(20, (int)(width * 7 + 5));
                            if ((bool?)col.Attribute("hidden") == true) sheet.HiddenColumns.Add(c - 1);
                            if (col.Attribute("style") != null) columnStyles[c - 1] = (int)col.Attribute("style");
                        }
                    }
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
                                throw new InvalidDataException("XLSX contains cells outside the 200 x 26 grid: " + address);
                            string type = (string)cell.Attribute("t") ?? "";
                            string value = (string)cell.Element(S + "v") ?? "";
                            if (type == "s") { int index; if (int.TryParse(value, out index) && index >= 0 && index < strings.Count) value = strings[index]; }
                            else if (type == "inlineStr") value = string.Concat(cell.Descendants(S + "t").Select(x => x.Value));
                            else if (type == "b") value = value == "1" ? "TRUE" : "FALSE";
                            var formula = cell.Element(S + "f");
                            if (formula != null && !string.IsNullOrEmpty(formula.Value)) value = "=" + formula.Value;
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
                            if (!range.IsEmpty) sheet.Tables.Add(new TableDefinition
                            { Name = (string)tableDefinition.Attribute("displayName") ??
                                (string)tableDefinition.Attribute("name") ?? "Table", Range = range });
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
                                if (definition != null) sheet.Charts.Add(definition);
                            }
                        }
                    }
                    result.Sheets.Add(sheet);
                }
                if (result.Sheets.Count == 0) result.Sheets.Add(new SheetSnapshot());
                foreach (var defined in book.Descendants(S + "definedName"))
                {
                    string name = (string)defined.Attribute("name") ?? "";
                    if (name.Length == 0 || name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)) continue;
                    int bang = defined.Value.LastIndexOf('!');
                    if (bang < 0) continue;
                    string sheetName = defined.Value.Substring(0, bang).Trim('\'').Replace("''", "'");
                    Rectangle range = ParseArea(defined.Value.Substring(bang + 1), rows, columns);
                    if (!range.IsEmpty) result.NamedRanges.Add(new NamedRange
                    { Name = name, Sheet = sheetName, Range = range });
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
                        if (sheetRels.HasElements)
                            WriteXml(zip, "xl/worksheets/_rels/sheet" + number + ".xml.rels", new XDocument(sheetRels));
                        WriteXml(zip, part, BuildSheet(book.Sheets[i], rows, columns, styleCatalog,
                            tableIds, drawingRelation));
                    }
                    relationships.Add(new XElement(P + "Relationship", new XAttribute("Id", "rIdStyles"), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"), new XAttribute("Target", "styles.xml")));
                    WriteXml(zip, "[Content_Types].xml", new XDocument(types));
                    WriteXml(zip, "_rels/.rels", new XDocument(new XElement(P + "Relationships", new XElement(P + "Relationship", new XAttribute("Id", "rId1"), new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"), new XAttribute("Target", "xl/workbook.xml")))));
                    var workbookElement = new XElement(S + "workbook", new XAttribute(XNamespace.Xmlns + "r", R), sheetsElement);
                    if (book.NamedRanges.Count > 0)
                    {
                        var definitions = new XElement(S + "definedNames");
                        foreach (NamedRange named in book.NamedRanges)
                            if (!named.Range.IsEmpty)
                            {
                                string quotedSheet = "'" + named.Sheet.Replace("'", "''") + "'";
                                definitions.Add(new XElement(S + "definedName", new XAttribute("name", named.Name),
                                    quotedSheet + "!" + AbsoluteRangeAddress(named.Range)));
                            }
                        workbookElement.Add(definitions);
                    }
                    WriteXml(zip, "xl/workbook.xml", new XDocument(workbookElement));
                    WriteXml(zip, "xl/_rels/workbook.xml.rels", new XDocument(relationships));
                    WriteXml(zip, "xl/styles.xml", styleCatalog.Document());
                }
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        private static XDocument BuildSheet(SheetSnapshot sheet, int rows, int columns, XlsxStyles styles,
            IList<string> tableIds, string drawingRelation)
        {
            var root = new XElement(S + "worksheet", new XAttribute(XNamespace.Xmlns + "r", R));
            if (!sheet.TabColor.IsEmpty)
                root.Add(new XElement(S + "sheetPr", new XElement(S + "tabColor",
                    new XAttribute("rgb", "FF" + sheet.TabColor.R.ToString("X2") +
                        sheet.TabColor.G.ToString("X2") + sheet.TabColor.B.ToString("X2")))));
            if (sheet.FreezeRow > 0 || sheet.FreezeColumn > 0)
                root.Add(new XElement(S + "sheetViews", new XElement(S + "sheetView", new XAttribute("workbookViewId", 0),
                    new XElement(S + "pane", new XAttribute("xSplit", sheet.FreezeColumn), new XAttribute("ySplit", sheet.FreezeRow),
                        new XAttribute("topLeftCell", ColumnName(sheet.FreezeColumn) + (sheet.FreezeRow + 1)), new XAttribute("state", "frozen")))));
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
            for (int r = 0; r < rows; r++)
            {
                var entries = sheet.Cells.Where(x => x.Key / columns == r).OrderBy(x => x.Key).ToList();
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
            var filterRules = new List<FilterCriterion>(sheet.Filters);
            if (sheet.FilterColumn >= 0 && sheet.FilterColumn < columns &&
                !filterRules.Any(f => f.Column == sheet.FilterColumn))
                filterRules.Add(new FilterCriterion { Column = sheet.FilterColumn, Operator = "Contains", Value1 = sheet.FilterValue });
            if (filterRules.Count > 0)
            {
                var filterElement = new XElement(S + "autoFilter", new XAttribute("ref", "A1:" + ColumnName(columns - 1) + rows));
                foreach (FilterCriterion criterion in filterRules.Where(f => f.Column >= 0 && f.Column < columns))
                {
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
            if (!string.IsNullOrEmpty(drawingRelation))
                root.Add(new XElement(S + "drawing", new XAttribute(R + "id", drawingRelation)));
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
            var header = new XElement(S + "tableColumns", new XAttribute("count", table.Range.Width));
            for (int c = table.Range.Left; c < table.Range.Right; c++)
            {
                CellSnapshot cell;
                string name = sheet.Cells.TryGetValue(table.Range.Top * columns + c, out cell) ? cell.Text : "";
                if (string.IsNullOrEmpty(name)) name = "Column" + (c - table.Range.Left + 1);
                header.Add(new XElement(S + "tableColumn", new XAttribute("id", c - table.Range.Left + 1),
                    new XAttribute("name", name)));
            }
            return new XDocument(new XElement(S + "table", new XAttribute("id", id),
                new XAttribute("name", table.Name), new XAttribute("displayName", table.Name),
                new XAttribute("ref", RangeAddress(table.Range)), new XAttribute("headerRowCount", 1),
                new XElement(S + "autoFilter", new XAttribute("ref", RangeAddress(table.Range))), header,
                new XElement(S + "tableStyleInfo", new XAttribute("name", "TableStyleMedium2"),
                    new XAttribute("showFirstColumn", 0), new XAttribute("showLastColumn", 0),
                    new XAttribute("showRowStripes", 1), new XAttribute("showColumnStripes", 0))));
        }
    }
}
