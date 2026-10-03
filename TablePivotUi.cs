using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace DinkCel
{
    internal sealed partial class SpreadsheetForm
    {
        private Font tableBoldFont;

        private static int DataStart(TableDefinition table) { return table.Range.Top + (table.HeaderRow ? 1 : 0); }
        private static int DataEnd(TableDefinition table) { return table.Range.Bottom - (table.TotalRow ? 1 : 0); }

        private static Rectangle ShiftTableRange(Rectangle range, bool rows, int index, bool insert)
        {
            int start = rows ? range.Top : range.Left;
            int length = rows ? range.Height : range.Width;
            if (insert)
            {
                if (index <= start) start++;
                else if (index < start + length) length++;
            }
            else
            {
                if (index < start) start--;
                else if (index < start + length) length--;
            }
            return rows ? new Rectangle(range.X, start, range.Width, length) :
                new Rectangle(start, range.Y, length, range.Height);
        }

        private void AdjustTableStructure(bool rows, int index, bool insert)
        {
            foreach (TableDefinition table in tables.ToArray())
            {
                table.Range = ShiftTableRange(table.Range, rows, index, insert);
                if (table.Range.Width < 1 || table.Range.Height < (table.HeaderRow ? 2 : 1))
                { tables.Remove(table); continue; }
                if (rows)
                {
                    foreach (int column in table.CalculatedColumns.Keys.ToArray())
                        table.CalculatedColumns[column] = FormulaEngine.ShiftStructureReferences(
                            table.CalculatedColumns[column], true, index, insert, MaxRowCount, ColumnCount);
                    continue;
                }
                var shifted = new Dictionary<int, string>();
                foreach (var entry in table.CalculatedColumns)
                {
                    if (insert || entry.Key != index)
                        shifted[entry.Key >= index + (insert ? 0 : 1) ?
                            entry.Key + (insert ? 1 : -1) : entry.Key] =
                            FormulaEngine.ShiftStructureReferences(entry.Value, rows, index, insert,
                                MaxRowCount, ColumnCount);
                }
                table.CalculatedColumns.Clear();
                foreach (var entry in shifted) table.CalculatedColumns[entry.Key] = entry.Value;
                foreach (FilterCriterion filter in table.Filters.ToArray())
                {
                    if (!insert && filter.Column == index) { table.Filters.Remove(filter); continue; }
                    if (filter.Column >= index + (insert ? 0 : 1)) filter.Column += insert ? 1 : -1;
                }
            }
            string sheetName = sheets[activeSheetIndex].Name;
            foreach (PivotDefinition pivot in pivots)
                if (string.Equals(pivot.SourceSheet, sheetName, StringComparison.OrdinalIgnoreCase) &&
                    string.IsNullOrEmpty(pivot.SourceTable))
                    pivot.SourceRange = ShiftTableRange(pivot.SourceRange, rows, index, insert);
        }

        private bool MoveCrossesTable(bool rows, int source, int target)
        {
            int low = Math.Min(source, target), high = Math.Max(source, target);
            foreach (TableDefinition table in tables)
            {
                int start = rows ? table.Range.Top : table.Range.Left;
                int end = rows ? table.Range.Bottom : table.Range.Right;
                if (high < start || low >= end) continue;
                if (rows && source >= DataStart(table) && source < DataEnd(table) &&
                    target >= DataStart(table) && target < DataEnd(table))
                    continue;
                return true;
            }
            return false;
        }

        private TableDefinition CurrentTable()
        {
            if (grid.CurrentCell != null)
            {
                TableDefinition current = tables.FirstOrDefault(t => t.Range.Contains(grid.CurrentCell.ColumnIndex,
                    grid.CurrentCell.RowIndex));
                if (current != null) return current;
            }
            string choice = ChooseOption("Chọn Table", tables.Select(t => t.Name).ToList());
            return choice == null ? null : tables.FirstOrDefault(t => t.Name == choice);
        }

        private string TableHeader(TableDefinition table, int column)
        {
            string header = table.HeaderRow ? Convert.ToString(grid[column, table.Range.Top].Value) ?? "" : "";
            return header.Length == 0 ? "Column" + (column - table.Range.Left + 1) : header;
        }

        private FormulaNamedRange ResolveStructuredRange(string currentSheet, string tableName, string header, int currentRow)
        {
            foreach (SheetState sheet in sheets)
            {
                if (tableName == null && !string.Equals(sheet.Name, currentSheet, StringComparison.OrdinalIgnoreCase)) continue;
                IEnumerable<TableDefinition> definitions = sheet == sheets[activeSheetIndex] ? tables : sheet.Tables;
                foreach (TableDefinition table in definitions)
                {
                    if (tableName != null && !string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase)) continue;
                    if (currentRow >= 0 && (currentRow < DataStart(table) || currentRow >= DataEnd(table))) continue;
                    for (int column = table.Range.Left; column < table.Range.Right; column++)
                    {
                        string columnName = sheet == sheets[activeSheetIndex] ? TableHeader(table, column) :
                            StateRaw(sheet, table.Range.Top, column);
                        if (string.IsNullOrEmpty(columnName)) columnName = "Column" + (column - table.Range.Left + 1);
                        if (!string.Equals(header, columnName, StringComparison.OrdinalIgnoreCase)) continue;
                        int first = currentRow >= 0 ? currentRow : DataStart(table);
                        int last = currentRow >= 0 ? currentRow : DataEnd(table) - 1;
                        if (last < first) return null;
                        return new FormulaNamedRange { Sheet = sheet.Name, FirstRow = first,
                            LastRow = last, FirstColumn = column, LastColumn = column };
                    }
                }
            }
            return null;
        }

        private void CreateStyledTable()
        {
            Rectangle range = SelectionRange(true);
            if (range.Width < 1 || range.Height < 2)
            { MessageBox.Show(this, "Table cần hàng tiêu đề và dữ liệu."); return; }
            string name = Prompt("Tên Table", "Table" + (sheets.SelectMany(s => s.Tables).Count() + 1));
            if (name == null) return;
            name = name.Trim();
            if (!ValidName(name) || tables.Any(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase)) ||
                sheets.SelectMany(s => s.Tables).Any(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase)))
            { MessageBox.Show(this, "Tên Table không hợp lệ hoặc đã tồn tại trong workbook."); return; }
            if (tables.Any(t => t.Range.IntersectsWith(range)))
            { MessageBox.Show(this, "Vùng Table đang chồng lên Table khác."); return; }
            var table = new TableDefinition { Name = name, Range = range };
            for (int col = range.Left; col < range.Right; col++)
                if (string.IsNullOrWhiteSpace(Convert.ToString(grid[col, range.Top].Value)))
                    grid[col, range.Top].Value = "Column" + (col - range.Left + 1);
            tables.Add(table);
            RecordChange(); MarkDirty(); grid.Invalidate();
            status.Text = "Đã tạo Table " + name;
        }

        private void ConfigureTable()
        {
            TableDefinition table = CurrentTable();
            if (table == null) return;
            var layout = new TableLayoutPanel { AutoScroll = true };
            var name = new TextBox { Text = table.Name };
            var style = DataChoice("TableStyleLight9", "TableStyleMedium2", "TableStyleMedium9", "TableStyleDark1");
            style.Text = table.Style;
            var header = new CheckBox { Text = "Hiện hàng tiêu đề", Checked = table.HeaderRow, AutoSize = true };
            var total = new CheckBox { Text = "Hiện hàng tổng", Checked = table.TotalRow, AutoSize = true };
            var bandRows = new CheckBox { Text = "Kẻ sọc hàng", Checked = table.BandedRows, AutoSize = true };
            var bandColumns = new CheckBox { Text = "Kẻ sọc cột", Checked = table.BandedColumns, AutoSize = true };
            var auto = new CheckBox { Text = "Tự mở rộng", Checked = table.AutoExpand, AutoSize = true };
            var filter = new CheckBox { Text = "Bộ lọc Table", Checked = table.Filter, AutoSize = true };
            DataField(layout, "Tên", name); DataField(layout, "Kiểu", style);
            DataField(layout, "Tiêu đề", header); DataField(layout, "Tổng", total);
            DataField(layout, "Hàng", bandRows); DataField(layout, "Cột", bandColumns);
            DataField(layout, "Mở rộng", auto); DataField(layout, "Lọc", filter);
            using (var dialog = DataDialog("Thiết lập Table", layout))
            {
                dialog.Height = 440;
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                string nextName = name.Text.Trim();
                if (!ValidName(nextName) || tables.Any(t => !ReferenceEquals(t, table) &&
                    string.Equals(t.Name, nextName, StringComparison.OrdinalIgnoreCase)) ||
                    sheets.Where((s, i) => i != activeSheetIndex).SelectMany(s => s.Tables)
                    .Any(t => string.Equals(t.Name, nextName, StringComparison.OrdinalIgnoreCase)))
                { MessageBox.Show(this, "Tên Table không hợp lệ hoặc đã tồn tại."); return; }
                string oldName = table.Name;
                table.Name = nextName;
                if (total.Checked && !table.TotalRow && !AddTotalRow(table))
                { table.Name = oldName; return; }
                if (!total.Checked && table.TotalRow) RemoveTotalRow(table);
                if (!string.Equals(oldName, nextName, StringComparison.OrdinalIgnoreCase))
                    RenameTableReferences(oldName, nextName);
                table.Style = style.Text; table.HeaderRow = header.Checked;
                table.BandedRows = bandRows.Checked; table.BandedColumns = bandColumns.Checked;
                table.AutoExpand = auto.Checked; table.Filter = filter.Checked;
                if (!table.Filter) table.Filters.Clear();
                Recalculate(); ApplyFreezeAndFilter(); RecordChange(); MarkDirty(); grid.Invalidate();
            }
        }

        private static string RebaseTableName(string formula, string oldName, string newName)
        {
            if (string.IsNullOrEmpty(formula) || !formula.StartsWith("=", StringComparison.Ordinal)) return formula;
            var pattern = new Regex(@"(?<![A-Za-z0-9_])" + Regex.Escape(oldName) + @"(?=\[)", RegexOptions.IgnoreCase);
            var result = new StringBuilder(); var code = new StringBuilder(); bool quoted = false;
            for (int i = 0; i < formula.Length; i++)
            {
                char ch = formula[i];
                if (ch == '"')
                {
                    if (!quoted) { result.Append(pattern.Replace(code.ToString(), newName)); code.Clear(); quoted = true; }
                    else if (i + 1 < formula.Length && formula[i + 1] == '"')
                    { result.Append("\"\""); i++; continue; }
                    else quoted = false;
                    result.Append(ch);
                }
                else if (quoted) result.Append(ch);
                else code.Append(ch);
            }
            result.Append(pattern.Replace(code.ToString(), newName));
            return result.ToString();
        }

        private void RenameTableReferences(string oldName, string newName)
        {
            foreach (PivotDefinition pivot in pivots)
                if (string.Equals(pivot.SourceTable, oldName, StringComparison.OrdinalIgnoreCase)) pivot.SourceTable = newName;
            foreach (TableDefinition definition in tables)
                foreach (int column in definition.CalculatedColumns.Keys.ToArray())
                    definition.CalculatedColumns[column] = RebaseTableName(definition.CalculatedColumns[column], oldName, newName);
            loading = true;
            try
            {
                foreach (int key in formulaKeys.ToArray())
                {
                    DataGridViewCell cell = grid[key % ColumnCount, key / ColumnCount];
                    string raw = Convert.ToString(cell.Value) ?? "";
                    string renamed = RebaseTableName(raw, oldName, newName);
                    if (renamed != raw) cell.Value = renamed;
                }
                foreach (SheetState sheet in sheets.Where((s, i) => i != activeSheetIndex))
                    foreach (int key in sheet.Cells.Keys.ToArray())
                    {
                        CellState cell = sheet.Cells[key]; string raw = Convert.ToString(cell.Value) ?? "";
                        string renamed = RebaseTableName(raw, oldName, newName);
                        if (renamed != raw)
                            sheet.Cells[key] = new CellState { Value = renamed, Style = cell.Style,
                                Extras = CellExtras.Copy(cell.Extras) };
                    }
            }
            finally { loading = false; }
        }

        private bool AddTotalRow(TableDefinition table)
        {
            if (table.Range.Bottom >= MaxRowCount) return false;
            EnsureRowCapacity(table.Range.Bottom + 1);
            for (int col = table.Range.Left; col < table.Range.Right; col++)
                if (!string.IsNullOrEmpty(Convert.ToString(grid[col, table.Range.Bottom].Value)))
                { MessageBox.Show(this, "Hàng dưới Table đã có dữ liệu. Hãy chèn hàng trống trước."); return false; }
            loading = true;
            try
            {
                int row = table.Range.Bottom;
                table.Range = new Rectangle(table.Range.X, table.Range.Y, table.Range.Width, table.Range.Height + 1);
                table.TotalRow = true;
                grid[table.Range.Left, row].Value = "Tổng cộng";
                for (int col = table.Range.Left + 1; col < table.Range.Right; col++)
                {
                    double number;
                    if (DataTools.Number(CellDisplay(Math.Min(row - 1, DataStart(table)), col), out number))
                        grid[col, row].Value = "=SUM(" + table.Name + "[" + TableHeader(table, col) + "])";
                }
            }
            finally { loading = false; }
            return true;
        }

        private void RemoveTotalRow(TableDefinition table)
        {
            int row = table.Range.Bottom - 1;
            loading = true;
            try
            {
                for (int col = table.Range.Left; col < table.Range.Right; col++) ClearCell(col, row);
                table.Range = new Rectangle(table.Range.X, table.Range.Y, table.Range.Width, table.Range.Height - 1);
                table.TotalRow = false;
            }
            finally { loading = false; }
        }

        private void SetCalculatedColumn()
        {
            TableDefinition table = CurrentTable();
            if (table == null) return;
            var headers = Enumerable.Range(table.Range.Left, table.Range.Width).Select(c => TableHeader(table, c)).ToList();
            string selected = ChooseOption("Cột tính", headers);
            if (selected == null) return;
            int column = table.Range.Left + headers.IndexOf(selected);
            string formula = Prompt("Công thức cho cột " + selected, table.CalculatedColumns.ContainsKey(column) ?
                table.CalculatedColumns[column] : "=");
            if (formula == null) return;
            if (!formula.StartsWith("=", StringComparison.Ordinal))
            { MessageBox.Show(this, "Công thức cần bắt đầu bằng =."); return; }
            table.CalculatedColumns[column] = formula;
            loading = true;
            try
            {
                for (int row = DataStart(table); row < DataEnd(table); row++)
                    grid[column, row].Value = FormulaEngine.ShiftReferences(formula,
                        row - DataStart(table), 0, RowCount, ColumnCount);
            }
            finally { loading = false; }
            Recalculate(); RecordChange(); MarkDirty();
        }

        private void AutoExpandTable(int row, int column)
        {
            foreach (TableDefinition table in tables)
            {
                if (!table.AutoExpand) continue;
                if (row == table.Range.Bottom && column >= table.Range.Left && column < table.Range.Right)
                {
                    if (table.TotalRow)
                    {
                        if (table.Range.Bottom >= MaxRowCount) continue;
                        EnsureRowCapacity(table.Range.Bottom + 1);
                        object entered = grid[column, row].Value;
                        int footer = table.Range.Bottom - 1;
                        loading = true;
                        try
                        {
                            for (int col = table.Range.Left; col < table.Range.Right; col++)
                            {
                                CopyCell(col, footer, col, row);
                                ClearCell(col, footer);
                            }
                            grid[column, footer].Value = entered;
                        }
                        finally { loading = false; }
                        table.Range = new Rectangle(table.Range.X, table.Range.Y, table.Range.Width, table.Range.Height + 1);
                        FillNewTableRow(table, footer);
                    }
                    else
                    {
                        table.Range = new Rectangle(table.Range.X, table.Range.Y, table.Range.Width, table.Range.Height + 1);
                        FillNewTableRow(table, row);
                    }
                    formulaEngine = null;
                    grid.Invalidate(); return;
                }
                if (column == table.Range.Right && row >= table.Range.Top && row < table.Range.Bottom &&
                    table.Range.Right < ColumnCount)
                {
                    table.Range = new Rectangle(table.Range.X, table.Range.Y, table.Range.Width + 1, table.Range.Height);
                    formulaEngine = null;
                    grid.Invalidate(); return;
                }
                if (row == DataStart(table) && table.Range.Contains(column, row))
                {
                    string raw = Convert.ToString(grid[column, row].Value) ?? "";
                    if (raw.StartsWith("=", StringComparison.Ordinal) && !table.CalculatedColumns.ContainsKey(column) &&
                        Enumerable.Range(row + 1, Math.Max(0, DataEnd(table) - row - 1))
                            .All(r => string.IsNullOrEmpty(Convert.ToString(grid[column, r].Value))))
                    {
                        table.CalculatedColumns[column] = raw;
                        loading = true;
                        try
                        {
                            for (int r = row + 1; r < DataEnd(table); r++)
                                grid[column, r].Value = FormulaEngine.ShiftReferences(raw, r - row, 0, RowCount, ColumnCount);
                        }
                        finally { loading = false; }
                    }
                }
            }
        }

        private void FillNewTableRow(TableDefinition table, int row)
        {
            loading = true;
            try
            {
                foreach (var entry in table.CalculatedColumns)
                    if (string.IsNullOrEmpty(Convert.ToString(grid[entry.Key, row].Value)))
                        grid[entry.Key, row].Value = FormulaEngine.ShiftReferences(entry.Value,
                            row - DataStart(table), 0, RowCount, ColumnCount);
            }
            finally { loading = false; }
        }

        private void SetTableFilter()
        {
            TableDefinition table = CurrentTable();
            if (table == null) return;
            var layout = new TableLayoutPanel();
            var columns = Enumerable.Range(table.Range.Left, table.Range.Width).Select(c => TableHeader(table, c)).ToArray();
            var col = DataChoice(columns);
            if (grid.CurrentCell != null && table.Range.Contains(grid.CurrentCell.ColumnIndex, grid.CurrentCell.RowIndex))
                col.SelectedIndex = grid.CurrentCell.ColumnIndex - table.Range.Left;
            var kind = DataChoice("Text", "Number", "Date");
            var op = DataChoice("Contains", "Begins With", "Equals", "One Of", "Greater", "Less", "Between", "Blank", "Nonblank");
            var first = new TextBox(); var second = new TextBox();
            DataField(layout, "Cột", col); DataField(layout, "Kiểu", kind);
            DataField(layout, "Điều kiện", op); DataField(layout, "Giá trị 1", first); DataField(layout, "Giá trị 2", second);
            using (var dialog = DataDialog("Lọc Table " + table.Name, layout))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                int column = table.Range.Left + col.SelectedIndex;
                table.Filters.RemoveAll(f => f.Column == column);
                table.Filters.Add(new FilterCriterion { Column = column, Kind = kind.Text,
                    Operator = op.Text, Value1 = op.Text == "One Of" ? string.Join("\n",
                        first.Text.Split(',').Select(v => v.Trim()).ToArray()) : first.Text,
                    Value2 = second.Text });
                table.Filter = true;
                ApplyFreezeAndFilter(); RecordChange(); MarkDirty();
            }
        }

        private void ClearTableFilter()
        {
            TableDefinition table = CurrentTable();
            if (table == null || table.Filters.Count == 0) return;
            table.Filters.Clear(); ApplyFreezeAndFilter(); RecordChange(); MarkDirty();
        }

        private void ApplyTableFormatting(DataGridViewCellFormattingEventArgs e)
        {
            TableDefinition table = tables.FirstOrDefault(t => t.Range.Contains(e.ColumnIndex, e.RowIndex));
            if (table == null) return;
            DataGridViewCell cell = grid[e.ColumnIndex, e.RowIndex];
            bool header = table.HeaderRow && e.RowIndex == table.Range.Top;
            bool total = table.TotalRow && e.RowIndex == table.Range.Bottom - 1;
            Color primary = table.Style == "TableStyleMedium9" ? Color.FromArgb(33, 123, 96) :
                table.Style == "TableStyleDark1" ? Color.FromArgb(43, 53, 75) :
                table.Style == "TableStyleLight9" ? Color.FromArgb(75, 125, 180) :
                Color.FromArgb(38, 105, 181);
            if (cell.Style.BackColor.IsEmpty)
            {
                Color background = header ? primary : total ? Blend(Color.White, primary, .24) : Color.Empty;
                if (!header && !total && table.BandedRows && (e.RowIndex - DataStart(table)) % 2 == 1)
                    background = Blend(Color.White, primary, .10);
                if (!header && !total && table.BandedColumns && (e.ColumnIndex - table.Range.Left) % 2 == 1)
                    background = Blend(background.IsEmpty ? Color.White : background, primary, .08);
                if (!background.IsEmpty) e.CellStyle.BackColor = background;
            }
            if (header && cell.Style.ForeColor.IsEmpty) e.CellStyle.ForeColor = Color.White;
            if (header && table.Filter)
            { e.Value = (Convert.ToString(cell.Value) ?? "") + "  ▾"; e.FormattingApplied = true; }
            if ((header || total) && cell.Style.Font == null)
            {
                if (tableBoldFont == null) tableBoldFont = new Font(grid.Font, FontStyle.Bold);
                e.CellStyle.Font = tableBoldFont;
            }
        }

        private void TableContextMenuNeeded(object sender, DataGridViewCellContextMenuStripNeededEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            TableDefinition table = tables.FirstOrDefault(t => t.HeaderRow && t.Range.Top == e.RowIndex &&
                e.ColumnIndex >= t.Range.Left && e.ColumnIndex < t.Range.Right);
            if (table == null) return;
            var context = new ContextMenuStrip { BackColor = theme.Chrome, ForeColor = theme.Text };
            int column = e.ColumnIndex, row = e.RowIndex;
            context.Items.Add("Lọc " + TableHeader(table, column) + "...", null, delegate
            { grid.CurrentCell = grid[column, row]; SetTableFilter(); });
            context.Items.Add("Bỏ lọc Table", null, delegate
            { grid.CurrentCell = grid[column, row]; ClearTableFilter(); });
            context.Items.Add("Thiết lập Table...", null, delegate
            { grid.CurrentCell = grid[column, row]; ConfigureTable(); });
            e.ContextMenuStrip = context;
        }
    }
}
