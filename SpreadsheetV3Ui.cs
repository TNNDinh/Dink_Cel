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
                foreach (int key in gridOccupied)
                    if (!string.IsNullOrEmpty(Convert.ToString(grid[key % ColumnCount, key / ColumnCount].Value)))
                    { bottom = Math.Max(bottom, key / ColumnCount); right = Math.Max(right, key % ColumnCount); }
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
                MaximizeBox = false, MinimizeBox = false,
                BackColor = theme.Chrome, ForeColor = theme.Text, Font = DinkDesign.Ui })
            {
                var combo = new ComboBox { Left = 12, Top = 12, Width = 380,
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    BackColor = theme.Sheet, ForeColor = theme.Text };
                foreach (string value in values) combo.Items.Add(value);
                combo.SelectedIndex = 0;
                var ok = DinkDesign.Button("OK", delegate { });
                ok.SetBounds(312, 48, 80, 28);
                ok.AutoSize = false; ok.DialogResult = DialogResult.OK;
                ok.BackColor = theme.AccentSoft; ok.ForeColor = theme.Accent;
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
            CreateStyledTable();
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
            if (rule == null) return true;
            if (rule.Kind == "Custom Formula")
                return string.IsNullOrEmpty(value) ? rule.AllowBlank :
                    EvaluateRuleFormula(rule.Value1, row, column, value, rule.Range.Top, rule.Range.Left);
            string checkedValue = value != null && value.StartsWith("=", StringComparison.Ordinal) ?
                EvaluateAt(value, row, column, value) : value;
            return DataTools.Valid(rule, checkedValue);
        }

        private bool ValidateCellChange(int row, int column)
        {
            string value = Convert.ToString(grid[column, row].Value) ?? "";
            if (CanAcceptValue(row, column, value)) return true;
            ValidationRule rule = ValidationFor(row, column);
            string alert = string.IsNullOrEmpty(rule.ErrorMessage) ?
                "Giá trị không đạt điều kiện kiểm tra dữ liệu." : rule.ErrorMessage;
            if (rule.ErrorStyle == "Information")
            {
                MessageBox.Show(this, alert, string.IsNullOrEmpty(rule.ErrorTitle) ? "DinkCel" : rule.ErrorTitle,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return true;
            }
            if (rule.ErrorStyle == "Warning" &&
                MessageBox.Show(this, alert + "\nVẫn giữ giá trị này?",
                    string.IsNullOrEmpty(rule.ErrorTitle) ? "DinkCel" : rule.ErrorTitle,
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes) return true;
            CellState old;
            object previous = lastState != null && lastState.Cells.TryGetValue(row * ColumnCount + column, out old)
                ? old.Value : null;
            loading = true;
            grid[column, row].Value = previous;
            loading = false;
            syncingContent = true;
            contentBox.Text = Convert.ToString(previous) ?? "";
            syncingContent = false;
            status.Text = alert;
            return false;
        }

        private void ShowValidationDropdown(int row, int column)
        {
            ValidationRule rule = ValidationFor(row, column);
            if (rule == null || rule.Kind != "List" || row < 0 || column < 0) return;
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
                StartPosition = FormStartPosition.CenterParent, BackColor = theme.Surface,
                ForeColor = theme.Text })
            {
                var chart = BuildChart(definition);
                var save = DinkDesign.Button("Lưu PNG...", delegate { });
                save.AutoSize = false;
                save.Dock = DockStyle.Bottom;
                save.Height = 36;
                save.BackColor = theme.AccentSoft;
                save.ForeColor = theme.Text;
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
            var chart = new Chart { Dock = DockStyle.Fill, BackColor = theme.Sheet,
                ForeColor = theme.Text };
            chart.ChartAreas.Add(new ChartArea("Main"));
            chart.Legends.Add(new Legend("Legend"));
            chart.Titles.Add(definition.Title);
            chart.ChartAreas[0].BackColor = theme.Sheet;
            chart.ChartAreas[0].AxisX.LabelStyle.ForeColor = theme.Muted;
            chart.ChartAreas[0].AxisY.LabelStyle.ForeColor = theme.Muted;
            chart.ChartAreas[0].AxisX.LineColor = theme.Border;
            chart.ChartAreas[0].AxisY.LineColor = theme.Border;
            chart.ChartAreas[0].AxisX.MajorGrid.LineColor = theme.GridLine;
            chart.ChartAreas[0].AxisY.MajorGrid.LineColor = theme.GridLine;
            chart.Legends[0].BackColor = theme.Sheet;
            chart.Legends[0].ForeColor = theme.Text;
            chart.Titles[0].ForeColor = theme.Text;
            Rectangle range = definition.Range;
            for (int c = range.Left + 1; c < range.Right; c++)
            {
                if (definition.Kind == "Pie" && c > range.Left + 1) break;
                string seriesName = Convert.ToString(grid[c, range.Top].Value);
                if (string.IsNullOrEmpty(seriesName)) seriesName = grid.Columns[c].HeaderText;
                var series = new Series(seriesName) { ChartType = definition.Kind == "Line" ?
                    SeriesChartType.Line : definition.Kind == "Pie" ? SeriesChartType.Pie : SeriesChartType.Column };
                Color[] chartColors = { theme.Accent, theme.Logo, theme.Dark ? Color.Turquoise : Color.Teal,
                    theme.Dark ? Color.Orange : Color.DarkOrange, theme.Muted };
                series.Color = chartColors[(c - range.Left - 1) % chartColors.Length];
                for (int r = range.Top + 1; r < range.Bottom; r++)
                {
                    string label = Convert.ToString(grid[range.Left, r].FormattedValue) ?? "";
                    string raw = Convert.ToString(grid[c, r].FormattedValue) ?? "";
                    double value;
                    if (double.TryParse(raw, NumberStyles.Any, CultureInfo.CurrentCulture, out value) ||
                        double.TryParse(raw, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                        series.Points.AddXY(label, value);
                }
                if (definition.Kind == "Pie")
                    for (int point = 0; point < series.Points.Count; point++)
                        series.Points[point].Color = chartColors[point % chartColors.Length];
                chart.Series.Add(series);
            }
            return chart;
        }

        private void CreatePivot()
        {
            ConfigurePivot(null);
        }

        private void RefreshAllPivots()
        {
            int original = activeSheetIndex;
            foreach (PivotDefinition pivot in pivots.ToArray()) RefreshPivot(pivot);
            if (original < sheets.Count) SwitchSheet(original);
        }

        private void RefreshPivot(PivotDefinition pivot)
        {
            RefreshPivotAdvanced(pivot);
        }

        private static string StateRaw(SheetState sheet, int row, int column)
        {
            CellState cell;
            return sheet.Cells.TryGetValue(row * ColumnCount + column, out cell) ?
                Convert.ToString(cell.Value) ?? "" : "";
        }
    }
}
