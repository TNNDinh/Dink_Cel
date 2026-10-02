using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Text.RegularExpressions;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace DinkCel
{
    internal static class XlsFile
    {
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
                    var sheet = new SheetSnapshot { Name = source.SheetName };
                    for (int c = 0; c < columns; c++)
                    {
                        int width = source.GetColumnWidth(c);
                        if (width != 8 * 256) sheet.ColumnWidths[c] = Math.Max(20, (int)(width / 256.0 * 7 + 5));
                    }
                    for (int r = 0; r <= source.LastRowNum; r++)
                    {
                        IRow line = source.GetRow(r);
                        if (line == null) continue;
                        if (r >= rows)
                        {
                            if (line.PhysicalNumberOfCells > 0) throw new InvalidDataException("XLS contains data outside the 200 x 26 grid.");
                            continue;
                        }
                        if (line.HeightInPoints > 0 && Math.Abs(line.HeightInPoints - source.DefaultRowHeightInPoints) > 0.1)
                            sheet.RowHeights[r] = (int)Math.Round(line.HeightInPoints * 96.0 / 72.0);
                        for (int c = 0; c < line.LastCellNum; c++)
                        {
                            ICell cell = line.GetCell(c);
                            if (cell == null) continue;
                            if (c >= columns) throw new InvalidDataException("XLS contains data outside the 200 x 26 grid.");
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
                            }
                            if (value.Length > 0 || snapshot.NumberFormat.Length > 0)
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
                foreach (SheetSnapshot sheet in source.Sheets)
                {
                    ISheet target = xls.CreateSheet(sheet.Name);
                    foreach (var width in sheet.ColumnWidths)
                        target.SetColumnWidth(width.Key, Math.Min(255 * 256, (int)(Math.Max(20, width.Value) / 7.0 * 256)));
                    foreach (var height in sheet.RowHeights)
                        target.CreateRow(height.Key).HeightInPoints = (float)(height.Value * 72.0 / 96.0);
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
                        if (!string.IsNullOrEmpty(entry.Value.NumberFormat) ||
                            entry.Value.Alignment != System.Windows.Forms.DataGridViewContentAlignment.NotSet)
                        {
                            ICellStyle style = xls.CreateCellStyle();
                            if (!string.IsNullOrEmpty(entry.Value.NumberFormat))
                                style.DataFormat = xls.CreateDataFormat().GetFormat(entry.Value.NumberFormat);
                            if (entry.Value.Alignment == System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter) style.Alignment = HorizontalAlignment.Center;
                            else if (entry.Value.Alignment == System.Windows.Forms.DataGridViewContentAlignment.MiddleRight) style.Alignment = HorizontalAlignment.Right;
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
