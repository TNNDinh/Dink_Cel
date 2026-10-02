using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace DinkCel
{
    internal sealed partial class SpreadsheetForm
    {
        private ComboBox validationEditor;

        private Rectangle SelectionRange(bool useDataWhenSingle)
        {
            if (grid.SelectedCells.Count > 1)
            {
                int left = grid.SelectedCells.Cast<DataGridViewCell>().Min(c => c.ColumnIndex);
                int right = grid.SelectedCells.Cast<DataGridViewCell>().Max(c => c.ColumnIndex);
                int top = grid.SelectedCells.Cast<DataGridViewCell>().Min(c => c.RowIndex);
                int bottom = grid.SelectedCells.Cast<DataGridViewCell>().Max(c => c.RowIndex);
                return new Rectangle(left, top, right - left + 1, bottom - top + 1);
            }
            if (useDataWhenSingle)
            {
                int bottom = 0, right = 0;
                for (int r = 0; r < RowCount; r++)
                    for (int c = 0; c < ColumnCount; c++)
                        if (!string.IsNullOrEmpty(Convert.ToString(grid[c, r].Value)))
                        { bottom = Math.Max(bottom, r); right = Math.Max(right, c); }
                return new Rectangle(0, 0, right + 1, bottom + 1);
            }
            return grid.CurrentCell == null ? Rectangle.Empty :
                new Rectangle(grid.CurrentCell.ColumnIndex, grid.CurrentCell.RowIndex, 1, 1);
        }

        private string ChooseOption(string title, IList<string> values)
        {
            if (values.Count == 0) return null;
            using (var form = new Form { Text = title, Width = 420, Height = 145,
                StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false })
            {
                var combo = new ComboBox { Left = 12, Top = 12, Width = 380,
                    DropDownStyle = ComboBoxStyle.DropDownList };
                foreach (string value in values) combo.Items.Add(value);
                combo.SelectedIndex = 0;
                var ok = new Button { Text = "OK", Left = 312, Top = 48,
                    Width = 80, DialogResult = DialogResult.OK };
                form.Controls.Add(combo); form.Controls.Add(ok); form.AcceptButton = ok;
                return form.ShowDialog(this) == DialogResult.OK ? combo.Text : null;
            }
        }

        private static bool ValidName(string name)
        {
            return !string.IsNullOrEmpty(name) && name.Length <= 31 &&
                Regex.IsMatch(name, @"^[A-Za-z_][A-Za-z0-9_]*$") &&
                !Regex.IsMatch(name, @"^[A-Za-z]{1,2}[1-9][0-9]{0,2}$", RegexOptions.IgnoreCase);
        }

        private void CreateTable()
        {
            Rectangle range = SelectionRange(true);
            if (range.Height < 2 || range.Width < 1)
            { MessageBox.Show(this, "Table cần hàng tiêu đề và ít nhất một hàng dữ liệu."); return; }
            string name = Prompt("Tên Table", "Table" + (tables.Count + 1));
            if (name == null) return;
            name = name.Trim();
            if (!ValidName(name) || tables.Any(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase)))
            { MessageBox.Show(this, "Tên Table không hợp lệ hoặc đã tồn tại."); return; }
            tables.Add(new TableDefinition { Name = name, Range = range });
            loading = true;
            try
            {
                for (int r = range.Top; r < range.Bottom; r++)
                    for (int c = range.Left; c < range.Right; c++)
                    {
                        var cell = grid[c, r];
                        if (r == range.Top)
                        {
                            cell.Style.BackColor = theme.Accent;
                            cell.Style.ForeColor = Color.White;
                            Font original = cell.InheritedStyle.Font ?? grid.Font;
                            cell.Style.Font = new Font(original, original.Style | FontStyle.Bold);
                        }
                        else if ((r - range.Top) % 2 == 0)
                            cell.Style.BackColor = Color.FromArgb(228, 240, 248);
                    }
            }
            finally { loading = false; }
            RecordChange(); MarkDirty(); grid.Invalidate();
            status.Text = "Đã tạo " + name + " (hàng đầu là tiêu đề)";
        }

        private void AddDropdown()
        {
            Rectangle range = SelectionRange(false);
            if (range.IsEmpty) return;
            string list = Prompt("Nhập các lựa chọn, cách nhau bằng dấu phẩy", "Có,Không");
            if (list == null) return;
            var values = list.Split(',').Select(x => x.Trim()).Where(x => x.Length > 0)
                .Distinct(StringComparer.CurrentCultureIgnoreCase).ToList();
            if (values.Count == 0) return;
            validations.RemoveAll(v => v.Range.IntersectsWith(range));
            var rule = new ValidationRule { Range = range };
            rule.Choices.AddRange(values);
            validations.Add(rule);
            RecordChange(); MarkDirty();
            status.Text = "Đã thêm danh sách chọn cho " + range.Width * range.Height + " ô";
        }

        private ValidationRule ValidationFor(int row, int column)
        {
            return validations.LastOrDefault(v => v.Range.Contains(column, row));
        }

        private bool CanAcceptValue(int row, int column, string value)
        {
            ValidationRule rule = ValidationFor(row, column);
            return rule == null || string.IsNullOrEmpty(value) ||
                rule.Choices.Any(option => string.Equals(option, value, StringComparison.CurrentCultureIgnoreCase));
        }

        private bool ValidateCellChange(int row, int column)
        {
            string value = Convert.ToString(grid[column, row].Value) ?? "";
            if (CanAcceptValue(row, column, value)) return true;
            CellState old;
            object previous = lastState != null && lastState.Cells.TryGetValue(row * ColumnCount + column, out old)
                ? old.Value : null;
            loading = true;
            grid[column, row].Value = previous;
            loading = false;
            syncingContent = true;
            contentBox.Text = Convert.ToString(previous) ?? "";
            syncingContent = false;
            status.Text = "Giá trị không có trong danh sách chọn";
            return false;
        }

        private void ShowValidationDropdown(int row, int column)
        {
            ValidationRule rule = ValidationFor(row, column);
            if (rule == null || row < 0 || column < 0) return;
            if (validationEditor != null)
            { grid.Controls.Remove(validationEditor); validationEditor.Dispose(); validationEditor = null; }
            Rectangle bounds = grid.GetCellDisplayRectangle(column, row, true);
            validationEditor = new ComboBox { Bounds = bounds, DropDownStyle = ComboBoxStyle.DropDownList };
            foreach (string option in rule.Choices) validationEditor.Items.Add(option);
            string existing = Convert.ToString(grid[column, row].Value) ?? "";
            int index = rule.Choices.FindIndex(x => string.Equals(x, existing, StringComparison.CurrentCultureIgnoreCase));
            if (index >= 0) validationEditor.SelectedIndex = index;
            ComboBox editor = validationEditor;
            editor.SelectionChangeCommitted += delegate
            {
                grid[column, row].Value = editor.Text;
                if (validationEditor == editor) validationEditor = null;
                grid.Controls.Remove(editor); editor.Dispose(); grid.Focus();
            };
            editor.Leave += delegate
            {
                if (validationEditor != editor) return;
                validationEditor = null; grid.Controls.Remove(editor); editor.Dispose();
            };
            grid.Controls.Add(editor);
            editor.BringToFront(); editor.Focus(); editor.DroppedDown = true;
        }

        private void DefineNamedRange()
        {
            Rectangle range = SelectionRange(false);
            if (range.IsEmpty) return;
            string name = Prompt("Tên vùng", "Vung" + (namedRanges.Count + 1));
            if (name == null) return;
            name = name.Trim();
            if (!ValidName(name) || namedRanges.Any(n => string.Equals(n.Name, name, StringComparison.OrdinalIgnoreCase)))
            { MessageBox.Show(this, "Tên vùng không hợp lệ hoặc đã tồn tại."); return; }
            namedRanges.Add(new NamedRange { Name = name, Sheet = sheets[activeSheetIndex].Name, Range = range });
            Recalculate(); MarkDirty(); status.Text = "Đã đặt tên vùng " + name;
        }

        private void GoToNamedRange()
        {
            string name = ChooseOption("Đi tới vùng", namedRanges.Select(n => n.Name).ToList());
            if (name == null) return;
            NamedRange named = namedRanges.First(n => n.Name == name);
            int index = sheets.FindIndex(s => s.Name == named.Sheet);
            if (index < 0) return;
            SwitchSheet(index);
            grid.ClearSelection();
            for (int r = named.Range.Top; r < named.Range.Bottom; r++)
                for (int c = named.Range.Left; c < named.Range.Right; c++)
                    if (r < RowCount && c < ColumnCount) grid[c, r].Selected = true;
            grid.CurrentCell = grid[named.Range.Left, named.Range.Top];
        }

        private void DeleteNamedRange()
        {
            string name = ChooseOption("Xóa vùng có tên", namedRanges.Select(n => n.Name).ToList());
            if (name == null) return;
            namedRanges.RemoveAll(n => n.Name == name);
            Recalculate(); MarkDirty();
        }

        private void CreateChart()
        {
            Rectangle range = SelectionRange(true);
            if (range.Width < 2 || range.Height < 2)
            { MessageBox.Show(this, "Biểu đồ cần cột nhãn, cột số và ít nhất một hàng dữ liệu."); return; }
            string kind = ChooseOption("Loại biểu đồ", new[] { "Column", "Line", "Pie" });
            if (kind == null) return;
            string title = Prompt("Tiêu đề biểu đồ", "Biểu đồ " + (charts.Count + 1));
            if (title == null) return;
            var definition = new ChartDefinition { Title = title, Kind = kind, Range = range };
            charts.Add(definition);
            RecordChange(); MarkDirty();
            ShowChart(definition);
        }

        private void OpenChart()
        {
            string title = ChooseOption("Xem biểu đồ", charts.Select(c => c.Title).ToList());
            ChartDefinition definition = charts.FirstOrDefault(c => c.Title == title);
            if (definition != null) ShowChart(definition);
        }

        private void ShowChart(ChartDefinition definition)
        {
            using (var form = new Form { Text = definition.Title, Width = 900, Height = 620,
                StartPosition = FormStartPosition.CenterParent })
            {
                var chart = BuildChart(definition);
                var save = new Button { Text = "Lưu PNG...", Dock = DockStyle.Bottom, Height = 36 };
                save.Click += delegate
                {
                    using (var dialog = new SaveFileDialog { Filter = "PNG (*.png)|*.png", DefaultExt = "png" })
                        if (dialog.ShowDialog(form) == DialogResult.OK)
                            chart.SaveImage(dialog.FileName, ChartImageFormat.Png);
                };
                form.Controls.Add(chart); form.Controls.Add(save);
                form.ShowDialog(this);
            }
        }

        private Chart BuildChart(ChartDefinition definition)
        {
            var chart = new Chart { Dock = DockStyle.Fill, BackColor = Color.White };
            chart.ChartAreas.Add(new ChartArea("Main"));
            chart.Legends.Add(new Legend("Legend"));
            chart.Titles.Add(definition.Title);
            Rectangle range = definition.Range;
            for (int c = range.Left + 1; c < range.Right; c++)
            {
                if (definition.Kind == "Pie" && c > range.Left + 1) break;
                string seriesName = Convert.ToString(grid[c, range.Top].Value);
                if (string.IsNullOrEmpty(seriesName)) seriesName = grid.Columns[c].HeaderText;
                var series = new Series(seriesName) { ChartType = definition.Kind == "Line" ?
                    SeriesChartType.Line : definition.Kind == "Pie" ? SeriesChartType.Pie : SeriesChartType.Column };
                for (int r = range.Top + 1; r < range.Bottom; r++)
                {
                    string label = Convert.ToString(grid[range.Left, r].FormattedValue) ?? "";
                    string raw = Convert.ToString(grid[c, r].FormattedValue) ?? "";
                    double value;
                    if (double.TryParse(raw, NumberStyles.Any, CultureInfo.CurrentCulture, out value) ||
                        double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                        series.Points.AddXY(label, value);
                }
                chart.Series.Add(series);
            }
            return chart;
        }

        private void CreatePivot()
        {
            Rectangle range = SelectionRange(true);
            if (range.Width < 2 || range.Height < 2)
            { MessageBox.Show(this, "Pivot Table cần tiêu đề và ít nhất một hàng dữ liệu."); return; }
            var headers = new List<string>();
            for (int c = range.Left; c < range.Right; c++)
            {
                string name = Convert.ToString(grid[c, range.Top].Value);
                headers.Add(string.IsNullOrEmpty(name) ? grid.Columns[c].HeaderText : name);
            }
            string group = ChooseOption("Cột nhóm", headers);
            if (group == null) return;
            string value = ChooseOption("Cột giá trị", headers);
            if (value == null) return;
            string aggregate = ChooseOption("Phép tổng hợp", new[] { "Sum", "Count" });
            if (aggregate == null) return;
            string sourceName = sheets[activeSheetIndex].Name;
            AddSheet();
            int targetNumber = pivots.Count + 1;
            while (sheets.Any(s => s != sheets[activeSheetIndex] && s.Name == "Pivot" + targetNumber)) targetNumber++;
            string targetName = "Pivot" + targetNumber;
            sheets[activeSheetIndex].Name = targetName;
            var pivot = new PivotDefinition { SourceSheet = sourceName, SourceRange = range,
                GroupColumn = range.Left + headers.IndexOf(group),
                ValueColumn = range.Left + headers.IndexOf(value),
                Aggregate = aggregate, TargetSheet = targetName };
            pivots.Add(pivot);
            RefreshPivot(pivot);
            RefreshSheetTabs();
        }

        private void RefreshAllPivots()
        {
            int original = activeSheetIndex;
            foreach (PivotDefinition pivot in pivots.ToArray()) RefreshPivot(pivot);
            if (original < sheets.Count) SwitchSheet(original);
        }

        private void RefreshPivot(PivotDefinition pivot)
        {
            int sourceIndex = sheets.FindIndex(s => s.Name == pivot.SourceSheet);
            int targetIndex = sheets.FindIndex(s => s.Name == pivot.TargetSheet);
            if (sourceIndex < 0 || targetIndex < 0) return;
            SaveActiveSheet();
            SheetState source = sheets[sourceIndex];
            var engine = new FormulaEngine((r, c) => StateRaw(source, r, c),
                (name, r, c) =>
                {
                    SheetState other = sheets.FirstOrDefault(s =>
                        string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
                    return other == null ? null : StateRaw(other, r, c);
                }, source.Name, RowCount, ColumnCount, name =>
                {
                    NamedRange named = namedRanges.FirstOrDefault(n =>
                        string.Equals(n.Name, name, StringComparison.OrdinalIgnoreCase));
                    return named == null ? null : new FormulaNamedRange
                    { Sheet = named.Sheet, FirstRow = named.Range.Top, FirstColumn = named.Range.Left,
                        LastRow = named.Range.Bottom - 1, LastColumn = named.Range.Right - 1 };
                });
            var groups = new SortedDictionary<string, double>(StringComparer.CurrentCultureIgnoreCase);
            for (int r = pivot.SourceRange.Top + 1; r < pivot.SourceRange.Bottom; r++)
            {
                CellState groupCell, valueCell;
                string group = source.Cells.TryGetValue(r * ColumnCount + pivot.GroupColumn, out groupCell) ?
                    Convert.ToString(groupCell.Value) ?? "" : "";
                if (group.StartsWith("=", StringComparison.Ordinal))
                    group = engine.Display(r, pivot.GroupColumn);
                if (group.Length == 0) continue;
                double amount = 0;
                if (pivot.Aggregate == "Count") amount = 1;
                else if (source.Cells.TryGetValue(r * ColumnCount + pivot.ValueColumn, out valueCell))
                {
                    string raw = Convert.ToString(valueCell.Value) ?? "";
                    if (raw.StartsWith("=", StringComparison.Ordinal)) raw = engine.Display(r, pivot.ValueColumn);
                    double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out amount);
                }
                if (!groups.ContainsKey(group)) groups[group] = 0;
                groups[group] += amount;
            }
            if (targetIndex != activeSheetIndex) SwitchSheet(targetIndex);
            loading = true;
            try
            {
                for (int r = 0; r < RowCount; r++)
                    for (int c = 0; c < 2; c++) grid[c, r].Value = null;
                grid[0, 0].Value = "Nhóm";
                grid[1, 0].Value = pivot.Aggregate == "Count" ? "Số lượng" : "Tổng";
                int row = 1;
                foreach (var item in groups)
                {
                    if (row >= RowCount - 1) break;
                    grid[0, row].Value = item.Key;
                    grid[1, row].Value = item.Value.ToString(CultureInfo.InvariantCulture);
                    row++;
                }
                grid[0, row].Value = "Tổng cộng";
                grid[1, row].Value = groups.Values.Sum().ToString(CultureInfo.InvariantCulture);
                grid[0, 0].Style.Font = new Font(grid.Font, FontStyle.Bold);
                grid[1, 0].Style.Font = new Font(grid.Font, FontStyle.Bold);
            }
            finally { loading = false; }
            Recalculate(); RecordChange(); MarkDirty(); SaveActiveSheet();
            status.Text = "Đã làm mới " + pivot.TargetSheet;
        }

        private static string StateRaw(SheetState sheet, int row, int column)
        {
            CellState cell;
            return sheet.Cells.TryGetValue(row * ColumnCount + column, out cell) ?
                Convert.ToString(cell.Value) ?? "" : "";
        }
    }
}
