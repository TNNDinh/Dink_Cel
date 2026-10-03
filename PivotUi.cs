using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

namespace DinkCel
{
    internal sealed partial class SpreadsheetForm
    {
        private readonly Dictionary<string, Dictionary<int, string>> pivotRowKeys =
            new Dictionary<string, Dictionary<int, string>>(StringComparer.OrdinalIgnoreCase);

        private void ConfigurePivot(PivotDefinition existing)
        {
            Rectangle range;
            SheetState source;
            TableDefinition sourceTable = null;
            if (existing == null)
            {
                range = SelectionRange(true);
                source = sheets[activeSheetIndex];
                if (grid.CurrentCell != null)
                    sourceTable = tables.FirstOrDefault(t => t.Range.Contains(grid.CurrentCell.ColumnIndex,
                        grid.CurrentCell.RowIndex) && grid.SelectedCells.Count <= 1);
                if (sourceTable != null) range = sourceTable.Range;
            }
            else
            {
                source = sheets.FirstOrDefault(s => s.Name == existing.SourceSheet);
                if (source == null) return;
                range = existing.SourceRange;
                sourceTable = source.Tables.FirstOrDefault(t => t.Name == existing.SourceTable);
            }
            if (range.Width < 2 || range.Height < 2)
            { MessageBox.Show(this, "Pivot cần ít nhất hai cột, hàng tiêu đề và một hàng dữ liệu."); return; }
            var headers = Enumerable.Range(range.Left, range.Width).Select(c =>
                existing == null ? Convert.ToString(grid[c, range.Top].Value) ?? "" : StateRaw(source, range.Top, c))
                .Select((v, i) => string.IsNullOrEmpty(v) ? "Column" + (i + 1) : v).ToArray();
            var layout = new TableLayoutPanel { AutoScroll = true };
            var rowBoxes = new List<ComboBox>(); var rowGroups = new List<ComboBox>();
            var colBoxes = new List<ComboBox>(); var colGroups = new List<ComboBox>();
            var valueBoxes = new List<ComboBox>(); var aggregates = new List<ComboBox>();
            var filterBoxes = new List<ComboBox>(); var filterValues = new List<TextBox>();
            string[] optional = new[] { "(Không)" }.Concat(headers).ToArray();
            for (int i = 0; i < 3; i++)
            {
                ComboBox field = DataChoice(optional);
                if (i == 0) field.SelectedIndex = 1;
                ComboBox group = DataChoice("None", "Day", "Month", "Year");
                DataField(layout, "Rows " + (i + 1), field); DataField(layout, "Nhóm ngày", group);
                rowBoxes.Add(field); rowGroups.Add(group);
            }
            for (int i = 0; i < 2; i++)
            {
                ComboBox field = DataChoice(optional); ComboBox group = DataChoice("None", "Day", "Month", "Year");
                DataField(layout, "Columns " + (i + 1), field); DataField(layout, "Nhóm ngày", group);
                colBoxes.Add(field); colGroups.Add(group);
            }
            for (int i = 0; i < 3; i++)
            {
                ComboBox field = DataChoice(optional);
                if (i == 0) field.SelectedIndex = headers.Length;
                ComboBox aggregate = DataChoice("Sum", "Count", "Average", "Min", "Max");
                DataField(layout, "Values " + (i + 1), field); DataField(layout, "Tổng hợp", aggregate);
                valueBoxes.Add(field); aggregates.Add(aggregate);
            }
            for (int i = 0; i < 2; i++)
            {
                ComboBox field = DataChoice(optional); var value = new TextBox();
                DataField(layout, "Filter " + (i + 1), field); DataField(layout, "Chỉ hiện", value);
                filterBoxes.Add(field); filterValues.Add(value);
            }
            var grand = new CheckBox { Text = "Grand Total", Checked = true, AutoSize = true };
            var subtotal = new CheckBox { Text = "Subtotal", Checked = true, AutoSize = true };
            var sort = DataChoice("A → Z", "Z → A", "Giá trị nhỏ → lớn", "Giá trị lớn → nhỏ");
            DataField(layout, "Tổng cuối", grand); DataField(layout, "Tổng nhóm", subtotal);
            DataField(layout, "Sắp xếp", sort);
            if (existing != null)
            {
                var rows = existing.Rows.Count > 0 ? existing.Rows :
                    new List<PivotAxisField> { new PivotAxisField { Column = existing.GroupColumn } };
                var values = existing.Values.Count > 0 ? existing.Values :
                    new List<PivotValueField> { new PivotValueField { Column = existing.ValueColumn, Aggregate = existing.Aggregate } };
                for (int i = 0; i < Math.Min(3, rows.Count); i++)
                { rowBoxes[i].SelectedIndex = rows[i].Column - range.Left + 1; rowGroups[i].Text = rows[i].DateGroup; }
                for (int i = 0; i < Math.Min(2, existing.Columns.Count); i++)
                { colBoxes[i].SelectedIndex = existing.Columns[i].Column - range.Left + 1; colGroups[i].Text = existing.Columns[i].DateGroup; }
                for (int i = 0; i < Math.Min(3, values.Count); i++)
                { valueBoxes[i].SelectedIndex = values[i].Column - range.Left + 1; aggregates[i].Text = values[i].Aggregate; }
                for (int i = 0; i < Math.Min(2, existing.Filters.Count); i++)
                { filterBoxes[i].SelectedIndex = existing.Filters[i].Column - range.Left + 1; filterValues[i].Text = existing.Filters[i].Value; }
                grand.Checked = existing.GrandTotal; subtotal.Checked = existing.Subtotal;
                sort.SelectedIndex = existing.SortByValue ? existing.SortDescending ? 3 : 2 : existing.SortDescending ? 1 : 0;
            }
            using (var dialog = DataDialog(existing == null ? "Tạo Pivot Table" : "Thiết lập Pivot Table", layout))
            {
                dialog.Width = 510; dialog.Height = 660;
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                var rows = rowBoxes.Select((b, i) => new { Box = b, Group = rowGroups[i] })
                    .Where(x => x.Box.SelectedIndex > 0).Select(x => new PivotAxisField
                    { Column = range.Left + x.Box.SelectedIndex - 1, DateGroup = x.Group.Text }).ToList();
                var columns = colBoxes.Select((b, i) => new { Box = b, Group = colGroups[i] })
                    .Where(x => x.Box.SelectedIndex > 0).Select(x => new PivotAxisField
                    { Column = range.Left + x.Box.SelectedIndex - 1, DateGroup = x.Group.Text }).ToList();
                var values = valueBoxes.Select((b, i) => new { Box = b, Aggregate = aggregates[i] })
                    .Where(x => x.Box.SelectedIndex > 0).Select(x => new PivotValueField
                    { Column = range.Left + x.Box.SelectedIndex - 1, Aggregate = x.Aggregate.Text }).ToList();
                if (rows.Count == 0 || values.Count == 0)
                { MessageBox.Show(this, "Pivot cần ít nhất một trường Rows và một trường Values."); return; }
                PivotDefinition pivot = existing;
                if (pivot == null)
                {
                    string sourceName = sheets[activeSheetIndex].Name;
                    AddSheet();
                    int number = pivots.Count + 1;
                    while (sheets.Any(s => s.Name == "Pivot" + number)) number++;
                    sheets[activeSheetIndex].Name = "Pivot" + number;
                    pivot = new PivotDefinition { SourceSheet = sourceName,
                        TargetSheet = sheets[activeSheetIndex].Name, SourceRange = range,
                        SourceTable = sourceTable == null ? "" : sourceTable.Name };
                    pivots.Add(pivot); RefreshSheetTabs();
                }
                pivot.Rows.Clear(); pivot.Rows.AddRange(rows);
                pivot.Columns.Clear(); pivot.Columns.AddRange(columns);
                pivot.Values.Clear(); pivot.Values.AddRange(values);
                pivot.Filters.Clear();
                for (int i = 0; i < filterBoxes.Count; i++) if (filterBoxes[i].SelectedIndex > 0)
                    pivot.Filters.Add(new PivotFilterField { Column = range.Left + filterBoxes[i].SelectedIndex - 1,
                        Value = filterValues[i].Text });
                pivot.GroupColumn = rows[0].Column; pivot.ValueColumn = values[0].Column;
                pivot.Aggregate = values[0].Aggregate; pivot.GrandTotal = grand.Checked;
                pivot.Subtotal = subtotal.Checked; pivot.SortDescending = sort.SelectedIndex == 1 || sort.SelectedIndex == 3;
                pivot.SortByValue = sort.SelectedIndex >= 2;
                RefreshPivotAdvanced(pivot);
            }
        }

        private void EditCurrentPivot()
        {
            PivotDefinition pivot = pivots.FirstOrDefault(p => p.TargetSheet == sheets[activeSheetIndex].Name);
            if (pivot == null)
            {
                string choice = ChooseOption("Chọn Pivot", pivots.Select(p => p.TargetSheet).ToList());
                pivot = pivots.FirstOrDefault(p => p.TargetSheet == choice);
            }
            if (pivot != null) ConfigurePivot(pivot);
        }

        private void TogglePivotGroup()
        {
            PivotDefinition pivot = pivots.FirstOrDefault(p => p.TargetSheet == sheets[activeSheetIndex].Name);
            if (pivot == null || grid.CurrentCell == null) return;
            Dictionary<int, string> keys;
            if (!pivotRowKeys.TryGetValue(pivot.TargetSheet, out keys))
            { RefreshPivotAdvanced(pivot); if (!pivotRowKeys.TryGetValue(pivot.TargetSheet, out keys)) return; }
            string key;
            if (!keys.TryGetValue(grid.CurrentCell.RowIndex, out key))
            { status.Text = "Hãy chọn hàng nhóm của Pivot."; return; }
            if (!pivot.Collapsed.Add(key)) pivot.Collapsed.Remove(key);
            RefreshPivotAdvanced(pivot);
        }

        private void RefreshPivotAdvanced(PivotDefinition pivot)
        {
            int sourceIndex = sheets.FindIndex(s => s.Name == pivot.SourceSheet);
            int targetIndex = sheets.FindIndex(s => s.Name == pivot.TargetSheet);
            if (sourceIndex < 0 || targetIndex < 0) return;
            SaveActiveSheet();
            SheetState source = sheets[sourceIndex];
            if (!string.IsNullOrEmpty(pivot.SourceTable))
            {
                TableDefinition table = source.Tables.FirstOrDefault(t => t.Name == pivot.SourceTable);
                if (table != null) pivot.SourceRange = new Rectangle(table.Range.X, table.Range.Y,
                    table.Range.Width, table.Range.Height - (table.TotalRow ? 1 : 0));
            }
            var engine = new FormulaEngine((r, c) => StateRaw(source, r, c),
                (name, r, c) =>
                {
                    SheetState other = sheets.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
                    return other == null ? null : StateRaw(other, r, c);
                }, source.Name, RowCount, ColumnCount, name =>
                {
                    NamedRange named = namedRanges.FirstOrDefault(n => string.Equals(n.Name, name, StringComparison.OrdinalIgnoreCase));
                    return named == null ? null : new FormulaNamedRange { Sheet = named.Sheet,
                        FirstRow = named.Range.Top, LastRow = named.Range.Bottom - 1,
                        FirstColumn = named.Range.Left, LastColumn = named.Range.Right - 1 };
                }, null, ResolveStructuredRange);
            Func<int, int, string> read = (r, c) =>
            {
                string raw = StateRaw(source, r, c);
                return raw.StartsWith("=", StringComparison.Ordinal) ? engine.Display(r, c) : raw;
            };
            PivotResult result;
            try { result = PivotEngine.Build(pivot, read, MaxRowCount, ColumnCount); }
            catch (InvalidOperationException error)
            { MessageBox.Show(this, error.Message, "Pivot Table"); return; }
            if (targetIndex != activeSheetIndex) SwitchSheet(targetIndex);
            EnsureRowCapacity(result.Rows.Count);
            loading = true; grid.SuspendLayout();
            try
            {
                int clearRows = Math.Max(pivot.LastOutputRows, result.Rows.Count);
                int clearColumns = Math.Max(pivot.LastOutputColumns, result.Rows[0].Length);
                foreach (int key in gridOccupied.ToArray())
                    if (key / ColumnCount < clearRows && key % ColumnCount < clearColumns)
                        ClearCell(key % ColumnCount, key / ColumnCount);
                for (int row = 0; row < result.Rows.Count; row++)
                {
                    bool strong = row == 0 || result.Rows[row][0].StartsWith("Tổng ", StringComparison.Ordinal) ||
                        result.Rows[row][0].StartsWith("Tổng cộng", StringComparison.Ordinal);
                    for (int col = 0; col < result.Rows[row].Length; col++)
                    {
                        string value = result.Rows[row][col];
                        if (!string.IsNullOrEmpty(value)) grid[col, row].Value = value;
                        if (strong)
                        {
                            grid[col, row].Style.Font = tableBoldFont ?? (tableBoldFont = new Font(grid.Font, FontStyle.Bold));
                            grid[col, row].Style.BackColor = row == 0 ? theme.AccentSoft : Blend(Color.White, theme.Accent, .16);
                        }
                    }
                }
                pivot.LastOutputRows = result.Rows.Count;
                pivot.LastOutputColumns = result.Rows[0].Length;
                pivotRowKeys[pivot.TargetSheet] = result.ExpandableRows;
            }
            finally { grid.ResumeLayout(); loading = false; }
            Recalculate(); RecordChange(); MarkDirty(); SaveActiveSheet();
            status.Text = "Đã làm mới " + pivot.TargetSheet + " (" + (result.Rows.Count - 1) + " hàng)";
        }
    }
}
