using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace DinkCel
{
    internal static class XlsFile
    {
        private static Color PaletteColor(HSSFWorkbook book, short index)
        {
            var entry = book.GetCustomPalette().GetColor(index);
            byte[] rgb = entry == null ? null : entry.RGB;
            return rgb == null || rgb.Length < 3 ? Color.Empty : Color.FromArgb(rgb[0], rgb[1], rgb[2]);
        }

        private static BorderEdge ReadBorder(HSSFWorkbook book, BorderStyle style, short color)
        {
            return new BorderEdge { Style = style == BorderStyle.None ? "" :
                style.ToString().ToLowerInvariant(), Color = PaletteColor(book, color) };
        }

        private static BorderStyle XlsBorder(string style)
        {
            switch ((style ?? "").ToLowerInvariant())
            {
                case "thin": return BorderStyle.Thin;
                case "medium": return BorderStyle.Medium;
                case "thick": return BorderStyle.Thick;
                case "dashed": return BorderStyle.Dashed;
                case "dotted": return BorderStyle.Dotted;
                case "double": return BorderStyle.Double;
                default: return BorderStyle.None;
            }
        }

        private static short XlsColor(HSSFWorkbook book, Color color)
        {
            return color.IsEmpty ? (short)0 : book.GetCustomPalette().FindSimilarColor(
                color.R, color.G, color.B).Indexed;
        }

        private static string StyleKey(CellSnapshot cell)
        {
            var element = new XElement("style", new XAttribute("number", cell.NumberFormat ?? ""),
                new XAttribute("font", cell.HasFont), new XAttribute("family", cell.FontName ?? ""),
                new XAttribute("size", cell.FontSize.ToString(CultureInfo.InvariantCulture)),
                new XAttribute("fontFlags", (int)cell.FontStyle),
                new XAttribute("fore", cell.ForeColor.ToArgb()),
                new XAttribute("back", cell.BackColor.ToArgb()),
                new XAttribute("align", (int)cell.Alignment));
            CellExtras.WriteXml(element, cell.Extras);
            return element.ToString(SaveOptions.DisableFormatting);
        }

        public static WorkbookSnapshot Read(string path, int rows, int columns)
        {
            EmbeddedDependencies.Install();
            using (var input = File.OpenRead(path))
            {
                var xls = new HSSFWorkbook(input);
                var result = new WorkbookSnapshot();
                result.Sheets.Clear();
                for (int i = 0; i < xls.NumberOfSheets; i++)
                {
                    ISheet source = xls.GetSheetAt(i);
                    var sheet = new SheetSnapshot { Name = source.SheetName, Hidden = xls.IsSheetHidden(i) };
                    for (int c = 0; c < columns; c++)
                    {
                        int width = source.GetColumnWidth(c);
                        if (width != 8 * 256) sheet.ColumnWidths[c] = Math.Max(20, (int)(width / 256.0 * 7 + 5));
                        if (source.IsColumnHidden(c)) sheet.HiddenColumns.Add(c);
                    }
                    for (int r = 0; r <= source.LastRowNum; r++)
                    {
                        IRow line = source.GetRow(r);
                        if (line == null) continue;
                        if (r >= rows)
                        {
                            if (line.PhysicalNumberOfCells > 0) throw new InvalidDataException("XLS contains data outside the 50,000 x 26 grid.");
                            continue;
                        }
                        if (line.HeightInPoints > 0 && Math.Abs(line.HeightInPoints - source.DefaultRowHeightInPoints) > 0.1)
                            sheet.RowHeights[r] = (int)Math.Round(line.HeightInPoints * 96.0 / 72.0);
                        if (line.ZeroHeight) sheet.HiddenRows.Add(r);
                        for (int c = 0; c < line.LastCellNum; c++)
                        {
                            ICell cell = line.GetCell(c);
                            if (cell == null) continue;
                            if (c >= columns) throw new InvalidDataException("XLS contains data outside the 50,000 x 26 grid.");
                            string value = "";
                            switch (cell.CellType)
                            {
                                case CellType.Formula: value = "=" + cell.CellFormula; break;
                                case CellType.Numeric: value = cell.NumericCellValue.ToString("R", CultureInfo.InvariantCulture); break;
                                case CellType.Boolean: value = cell.BooleanCellValue ? "TRUE" : "FALSE"; break;
                                case CellType.String: value = cell.StringCellValue; break;
                                case CellType.Error: value = "#ERROR!"; break;
                            }
                            var snapshot = new CellSnapshot { Text = value };
                            if (cell.CellStyle != null)
                            {
                                string format = cell.CellStyle.GetDataFormatString();
                                if (!string.IsNullOrEmpty(format) && format != "General") snapshot.NumberFormat = format;
                                if (cell.CellStyle.Alignment == HorizontalAlignment.Center) snapshot.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
                                else if (cell.CellStyle.Alignment == HorizontalAlignment.Right) snapshot.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleRight;
                                ICellStyle style = cell.CellStyle;
                                IFont font = style.GetFont(xls);
                                if (font != null)
                                {
                                    snapshot.HasFont = true;
                                    snapshot.FontName = font.FontName;
                                    snapshot.FontSize = font.FontHeightInPoints;
                                    snapshot.FontStyle = (font.IsBold ? FontStyle.Bold : 0) |
                                        (font.IsItalic ? FontStyle.Italic : 0) |
                                        (font.IsStrikeout ? FontStyle.Strikeout : 0) |
                                        (font.Underline != FontUnderlineType.None ? FontStyle.Underline : 0);
                                    snapshot.ForeColor = PaletteColor(xls, font.Color);
                                }
                                if (style.FillPattern == FillPattern.SolidForeground)
                                    snapshot.BackColor = PaletteColor(xls, style.FillForegroundColor);
                                string horizontal = style.Alignment == HorizontalAlignment.Center ? "center" :
                                    style.Alignment == HorizontalAlignment.Right ? "right" : "left";
                                snapshot.Alignment = style.VerticalAlignment == VerticalAlignment.Top ?
                                    horizontal == "center" ? System.Windows.Forms.DataGridViewContentAlignment.TopCenter :
                                    horizontal == "right" ? System.Windows.Forms.DataGridViewContentAlignment.TopRight :
                                    System.Windows.Forms.DataGridViewContentAlignment.TopLeft :
                                    style.VerticalAlignment == VerticalAlignment.Bottom ?
                                    horizontal == "center" ? System.Windows.Forms.DataGridViewContentAlignment.BottomCenter :
                                    horizontal == "right" ? System.Windows.Forms.DataGridViewContentAlignment.BottomRight :
                                    System.Windows.Forms.DataGridViewContentAlignment.BottomLeft :
                                    horizontal == "center" ? System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter :
                                    horizontal == "right" ? System.Windows.Forms.DataGridViewContentAlignment.MiddleRight :
                                    System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
                                var extras = new CellExtras { Left = ReadBorder(xls, style.BorderLeft, style.LeftBorderColor),
                                    Right = ReadBorder(xls, style.BorderRight, style.RightBorderColor),
                                    Top = ReadBorder(xls, style.BorderTop, style.TopBorderColor),
                                    Bottom = ReadBorder(xls, style.BorderBottom, style.BottomBorderColor),
                                    Wrap = style.WrapText, Shrink = style.ShrinkToFit,
                                    Indent = style.Indention, Rotation = style.Rotation };
                                if (extras.Left.Exists || extras.Right.Exists || extras.Top.Exists || extras.Bottom.Exists ||
                                    extras.Wrap || extras.Shrink || extras.Indent != 0 || extras.Rotation != 0)
                                    snapshot.Extras = extras;
                            }
                            if (value.Length > 0 || snapshot.NumberFormat.Length > 0 || snapshot.Extras != null ||
                                snapshot.HasFont || !snapshot.BackColor.IsEmpty)
                                sheet.Cells[r * columns + c] = snapshot;
                        }
                    }
                    for (int m = 0; m < source.NumMergedRegions; m++)
                    {
                        CellRangeAddress merge = source.GetMergedRegion(m);
                        if (merge.LastRow < rows && merge.LastColumn < columns)
                            sheet.Merges.Add(new Rectangle(merge.FirstColumn, merge.FirstRow,
                                merge.LastColumn - merge.FirstColumn + 1,
                                merge.LastRow - merge.FirstRow + 1));
                    }
                    result.Sheets.Add(sheet);
                }
                for (int i = 0; i < xls.NumberOfNames; i++)
                {
                    IName name = xls.GetNameAt(i);
                    string formula = name.RefersToFormula ?? "";
                    int bang = formula.LastIndexOf('!');
                    if (bang < 0 || name.NameName.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase)) continue;
                    string sheetName = formula.Substring(0, bang).Trim('\'').Replace("''", "'");
                    Rectangle range = ParseRange(formula.Substring(bang + 1), rows, columns);
                    if (!range.IsEmpty) result.NamedRanges.Add(new NamedRange
                    { Name = name.NameName, Sheet = sheetName, Range = range });
                }
                if (result.Sheets.Count == 0) result.Sheets.Add(new SheetSnapshot());
                return result;
            }
        }

        public static void Write(string path, WorkbookSnapshot source, int rows, int columns)
        {
            EmbeddedDependencies.Install();
            string temporary = path + ".tmp";
            try
            {
                var xls = new HSSFWorkbook();
                var styleCache = new Dictionary<string, ICellStyle>();
                foreach (SheetSnapshot sheet in source.Sheets)
                {
                    ISheet target = xls.CreateSheet(sheet.Name);
                    if (sheet.Hidden) xls.SetSheetHidden(xls.NumberOfSheets - 1, true);
                    foreach (var width in sheet.ColumnWidths)
                        target.SetColumnWidth(width.Key, Math.Min(255 * 256, (int)(Math.Max(20, width.Value) / 7.0 * 256)));
                    foreach (int column in sheet.HiddenColumns) target.SetColumnHidden(column, true);
                    foreach (var height in sheet.RowHeights)
                        target.CreateRow(height.Key).HeightInPoints = (float)(height.Value * 72.0 / 96.0);
                    foreach (int row in sheet.HiddenRows)
                        (target.GetRow(row) ?? target.CreateRow(row)).ZeroHeight = true;
                    if (sheet.FreezeRow > 0 || sheet.FreezeColumn > 0)
                        target.CreateFreezePane(sheet.FreezeColumn, sheet.FreezeRow);
                    foreach (var entry in sheet.Cells)
                    {
                        int r = entry.Key / columns, c = entry.Key % columns;
                        if (r < 0 || r >= rows || c < 0 || c >= columns) continue;
                        IRow line = target.GetRow(r) ?? target.CreateRow(r);
                        ICell cell = line.CreateCell(c);
                        string text = entry.Value.Text ?? "";
                        if (text.StartsWith("=", StringComparison.Ordinal)) cell.SetCellFormula(text.Substring(1));
                        else
                        {
                            double number;
                            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number)) cell.SetCellValue(number);
                            else cell.SetCellValue(text);
                        }
                        if (!string.IsNullOrEmpty(entry.Value.NumberFormat) || entry.Value.HasFont ||
                            !entry.Value.ForeColor.IsEmpty || !entry.Value.BackColor.IsEmpty ||
                            entry.Value.Extras != null ||
                            entry.Value.Alignment != System.Windows.Forms.DataGridViewContentAlignment.NotSet)
                        {
                            string key = StyleKey(entry.Value);
                            ICellStyle style;
                            if (!styleCache.TryGetValue(key, out style))
                            {
                            style = xls.CreateCellStyle();
                            if (!string.IsNullOrEmpty(entry.Value.NumberFormat))
                                style.DataFormat = xls.CreateDataFormat().GetFormat(entry.Value.NumberFormat);
                            if (entry.Value.Alignment == System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter) style.Alignment = HorizontalAlignment.Center;
                            else if (entry.Value.Alignment == System.Windows.Forms.DataGridViewContentAlignment.MiddleRight) style.Alignment = HorizontalAlignment.Right;
                            else if (entry.Value.Alignment == System.Windows.Forms.DataGridViewContentAlignment.TopCenter ||
                                entry.Value.Alignment == System.Windows.Forms.DataGridViewContentAlignment.BottomCenter)
                                style.Alignment = HorizontalAlignment.Center;
                            else if (entry.Value.Alignment == System.Windows.Forms.DataGridViewContentAlignment.TopRight ||
                                entry.Value.Alignment == System.Windows.Forms.DataGridViewContentAlignment.BottomRight)
                                style.Alignment = HorizontalAlignment.Right;
                            if (entry.Value.Alignment == System.Windows.Forms.DataGridViewContentAlignment.TopLeft ||
                                entry.Value.Alignment == System.Windows.Forms.DataGridViewContentAlignment.TopCenter ||
                                entry.Value.Alignment == System.Windows.Forms.DataGridViewContentAlignment.TopRight)
                                style.VerticalAlignment = VerticalAlignment.Top;
                            else if (entry.Value.Alignment == System.Windows.Forms.DataGridViewContentAlignment.BottomLeft ||
                                entry.Value.Alignment == System.Windows.Forms.DataGridViewContentAlignment.BottomCenter ||
                                entry.Value.Alignment == System.Windows.Forms.DataGridViewContentAlignment.BottomRight)
                                style.VerticalAlignment = VerticalAlignment.Bottom;
                            if (entry.Value.HasFont || !entry.Value.ForeColor.IsEmpty)
                            {
                                IFont font = xls.CreateFont();
                                font.FontName = entry.Value.FontName;
                                font.FontHeightInPoints = (short)Math.Round(entry.Value.FontSize);
                                font.IsBold = (entry.Value.FontStyle & FontStyle.Bold) != 0;
                                font.IsItalic = (entry.Value.FontStyle & FontStyle.Italic) != 0;
                                font.IsStrikeout = (entry.Value.FontStyle & FontStyle.Strikeout) != 0;
                                if ((entry.Value.FontStyle & FontStyle.Underline) != 0)
                                    font.Underline = FontUnderlineType.Single;
                                if (!entry.Value.ForeColor.IsEmpty) font.Color = XlsColor(xls, entry.Value.ForeColor);
                                style.SetFont(font);
                            }
                            if (!entry.Value.BackColor.IsEmpty)
                            { style.FillPattern = FillPattern.SolidForeground;
                                style.FillForegroundColor = XlsColor(xls, entry.Value.BackColor); }
                            CellExtras extras = entry.Value.Extras;
                            if (extras != null)
                            {
                                style.WrapText = extras.Wrap;
                                style.ShrinkToFit = extras.Shrink;
                                style.Indention = (short)extras.Indent;
                                style.Rotation = (short)extras.Rotation;
                                style.BorderLeft = XlsBorder(extras.Left.Style);
                                style.BorderRight = XlsBorder(extras.Right.Style);
                                style.BorderTop = XlsBorder(extras.Top.Style);
                                style.BorderBottom = XlsBorder(extras.Bottom.Style);
                                if (extras.Left.Exists) style.LeftBorderColor = XlsColor(xls, extras.Left.Color);
                                if (extras.Right.Exists) style.RightBorderColor = XlsColor(xls, extras.Right.Color);
                                if (extras.Top.Exists) style.TopBorderColor = XlsColor(xls, extras.Top.Color);
                                if (extras.Bottom.Exists) style.BottomBorderColor = XlsColor(xls, extras.Bottom.Color);
                            }
                            styleCache[key] = style;
                            }
                            cell.CellStyle = style;
                        }
                    }
                    foreach (Rectangle merge in sheet.Merges)
                        if (merge.Width > 0 && merge.Height > 0)
                            target.AddMergedRegion(new CellRangeAddress(merge.Top, merge.Bottom - 1,
                                merge.Left, merge.Right - 1));
                }
                foreach (NamedRange named in source.NamedRanges)
                    if (!named.Range.IsEmpty)
                    {
                        IName name = xls.CreateName();
                        name.NameName = named.Name;
                        name.RefersToFormula = "'" + named.Sheet.Replace("'", "''") + "'!$" +
                            ColumnName(named.Range.Left) + "$" + (named.Range.Top + 1) + ":$" +
                            ColumnName(named.Range.Right - 1) + "$" + named.Range.Bottom;
                    }
                xls.ForceFormulaRecalculation = true;
                using (var output = File.Create(temporary)) xls.Write(output);
                if (File.Exists(path)) File.Replace(temporary, path, null);
                else File.Move(temporary, path);
            }
            finally { if (File.Exists(temporary)) File.Delete(temporary); }
        }

        private static string ColumnName(int column)
        {
            string name = "";
            do { name = (char)('A' + column % 26) + name; column = column / 26 - 1; } while (column >= 0);
            return name;
        }

        private static Rectangle ParseRange(string text, int rows, int columns)
        {
            string[] bounds = text.Replace("$", "").Split(':');
            int x1, y1, x2, y2;
            if (!ParseAddress(bounds[0], out x1, out y1)) return Rectangle.Empty;
            if (bounds.Length == 2)
            { if (!ParseAddress(bounds[1], out x2, out y2)) return Rectangle.Empty; }
            else { x2 = x1; y2 = y1; }
            return x1 < 0 || y1 < 0 || x2 >= columns || y2 >= rows ? Rectangle.Empty :
                new Rectangle(x1, y1, x2 - x1 + 1, y2 - y1 + 1);
        }

        private static bool ParseAddress(string text, out int column, out int row)
        {
            column = row = -1;
            Match match = Regex.Match(text, @"^([A-Za-z]+)([1-9][0-9]*)$");
            if (!match.Success) return false;
            column = 0;
            foreach (char c in match.Groups[1].Value.ToUpperInvariant()) column = column * 26 + c - 'A' + 1;
            column--;
            row = int.Parse(match.Groups[2].Value) - 1;
            return true;
        }
    }
}
