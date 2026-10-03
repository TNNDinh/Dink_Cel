using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace DinkCel
{
    internal static class DataTools
    {
        public static bool Number(string text, out double value)
        {
            return double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value) ||
                double.TryParse(text, NumberStyles.Float, CultureInfo.CurrentCulture, out value) ||
                double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands,
                    CultureInfo.CurrentCulture, out value);
        }

        public static int Compare(string left, string right, string kind)
        {
            double a, b;
            DateTime da, db;
            if (kind == "Number" && Number(left, out a) && Number(right, out b)) return a.CompareTo(b);
            if (kind == "Date" && Temporal(left, out da) && Temporal(right, out db)) return da.CompareTo(db);
            if (kind == "Auto")
            {
                if (Number(left, out a) && Number(right, out b)) return a.CompareTo(b);
                if (Temporal(left, out da) && Temporal(right, out db)) return da.CompareTo(db);
            }
            return StringComparer.CurrentCultureIgnoreCase.Compare(left, right);
        }

        public static bool Temporal(string text, out DateTime date)
        {
            if (DateTime.TryParse(text, CultureInfo.CurrentCulture, DateTimeStyles.None, out date)) return true;
            if (DateTime.TryParse(text, CultureInfo.InvariantCulture, DateTimeStyles.None, out date)) return true;
            double serial;
            if (Number(text, out serial) && serial >= 0 && serial < 2958466)
            {
                try { date = DateTime.FromOADate(serial); return true; }
                catch (ArgumentException) { }
            }
            date = DateTime.MinValue; return false;
        }

        public static bool Match(string actual, string kind, string op, string first, string second)
        {
            actual = actual ?? "";
            if (op == "Blank") return actual.Length == 0;
            if (op == "Nonblank") return actual.Length > 0;
            if (op == "Contains") return actual.IndexOf(first ?? "", StringComparison.CurrentCultureIgnoreCase) >= 0;
            if (op == "Begins With") return actual.StartsWith(first ?? "", StringComparison.CurrentCultureIgnoreCase);
            double n1, n2;
            DateTime d1, d2;
            if (kind == "Number" && (!Number(actual, out n1) || !Number(first, out n2) ||
                op == "Between" && !Number(second, out n2))) return false;
            if (kind == "Date" && (!Temporal(actual, out d1) || !Temporal(first, out d2) ||
                op == "Between" && !Temporal(second, out d2))) return false;
            int comparison = Compare(actual, first ?? "", kind);
            if (op == "Equals") return comparison == 0;
            if (op == "Greater") return comparison > 0;
            if (op == "Less") return comparison < 0;
            if (op == "Between") return comparison >= 0 && Compare(actual, second ?? "", kind) <= 0;
            return true;
        }

        public static bool Valid(ValidationRule rule, string value)
        {
            if (rule == null) return true;
            value = value ?? "";
            if (value.Length == 0) return rule.AllowBlank;
            if (rule.Kind == "List") return rule.Choices.Any(x => string.Equals(x, value, StringComparison.CurrentCultureIgnoreCase));
            if (rule.Kind == "Custom Formula") return true;
            double number;
            if (rule.Kind == "Whole Number" || rule.Kind == "Decimal" || rule.Kind == "Text Length")
            {
                if (rule.Kind == "Text Length") number = value.Length;
                else if (!Number(value, out number)) return false;
                if (rule.Kind == "Whole Number" && number != Math.Truncate(number)) return false;
                double lo, hi;
                if (!Number(rule.Value1, out lo)) return false;
                if (!Number(rule.Value2, out hi)) hi = lo;
                return TestBounds(number, lo, hi, rule.Operator);
            }
            DateTime date, low, high;
            if (rule.Kind == "Date" || rule.Kind == "Time")
            {
                if (!Temporal(value, out date) || !Temporal(rule.Value1, out low)) return false;
                if (!Temporal(rule.Value2, out high)) high = low;
                double actual = rule.Kind == "Time" ? date.TimeOfDay.TotalSeconds : date.Date.ToOADate();
                double a = rule.Kind == "Time" ? low.TimeOfDay.TotalSeconds : low.Date.ToOADate();
                double b = rule.Kind == "Time" ? high.TimeOfDay.TotalSeconds : high.Date.ToOADate();
                return TestBounds(actual, a, b, rule.Operator);
            }
            return true;
        }

        private static bool TestBounds(double value, double low, double high, string op)
        {
            switch (op)
            {
                case "Between": return value >= low && value <= high;
                case "Not Between": return value < low || value > high;
                case "Equals": return value == low;
                case "Not Equals": return value != low;
                case "Greater": return value > low;
                case "Greater Or Equal": return value >= low;
                case "Less": return value < low;
                case "Less Or Equal": return value <= low;
                default: return false;
            }
        }

        public static bool FindMatch(string text, string query, bool matchCase, bool whole, bool wildcard)
        {
            if (string.IsNullOrEmpty(query)) return false;
            if (wildcard)
            {
                string pattern = Regex.Escape(query).Replace(@"\*", ".*").Replace(@"\?", ".");
                return Regex.IsMatch(text ?? "", whole ? "^(?:" + pattern + ")$" : pattern,
                    matchCase ? RegexOptions.None : RegexOptions.IgnoreCase);
            }
            var comparison = matchCase ? StringComparison.CurrentCulture : StringComparison.CurrentCultureIgnoreCase;
            return whole ? string.Equals(text, query, comparison) : (text ?? "").IndexOf(query, comparison) >= 0;
        }

        public static string Replace(string text, string query, string replacement, bool matchCase, bool whole, bool wildcard)
        {
            if (!FindMatch(text, query, matchCase, whole, wildcard)) return text;
            if (whole) return replacement;
            string pattern = wildcard ? Regex.Escape(query).Replace(@"\*", ".*").Replace(@"\?", ".") : Regex.Escape(query);
            return Regex.Replace(text ?? "", pattern, delegate(Match m) { return replacement ?? ""; },
                matchCase ? RegexOptions.None : RegexOptions.IgnoreCase);
        }
    }

    internal sealed partial class SpreadsheetForm
    {
        private static Color ColorFromAttribute(string value)
        {
            return string.IsNullOrEmpty(value) ? Color.Empty : ColorTranslator.FromHtml(value);
        }

        private static ComboBox DataChoice(params string[] values)
        {
            var box = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 180 };
            box.Items.AddRange(values);
            if (box.Items.Count > 0) box.SelectedIndex = 0;
            return box;
        }

        private static void DataField(TableLayoutPanel layout, string label, Control control)
        {
            int row = layout.RowCount++;
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.Controls.Add(new Label { Text = label, AutoSize = true, Anchor = AnchorStyles.Left, Margin = new Padding(4, 8, 4, 4) }, 0, row);
            control.Anchor = AnchorStyles.Left | AnchorStyles.Right;
            layout.Controls.Add(control, 1, row);
        }

        private Form DataDialog(string title, TableLayoutPanel layout)
        {
            var form = new Form { Text = title, Width = 440, Height = Math.Min(670, 120 + layout.RowCount * 38),
                StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog,
                MaximizeBox = false, MinimizeBox = false, BackColor = theme.Surface, ForeColor = theme.Text };
            layout.Dock = DockStyle.Fill; layout.Padding = new Padding(14); layout.ColumnCount = 2;
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 42));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 58));
            var buttons = new FlowLayoutPanel { Dock = DockStyle.Bottom, FlowDirection = FlowDirection.RightToLeft, Height = 46 };
            var ok = new Button { Text = "Áp dụng", DialogResult = DialogResult.OK, Width = 95 };
            var cancel = new Button { Text = "Hủy", DialogResult = DialogResult.Cancel, Width = 80 };
            buttons.Controls.Add(ok); buttons.Controls.Add(cancel);
            form.Controls.Add(layout); form.Controls.Add(buttons); form.AcceptButton = ok; form.CancelButton = cancel;
            return form;
        }

        private void SortRowsAdvanced()
        {
            if (grid.CurrentCell == null) return;
            var layout = new TableLayoutPanel { AutoScroll = true };
            var columns = new List<ComboBox>(); var orders = new List<ComboBox>(); var types = new List<ComboBox>();
            for (int i = 0; i < 3; i++)
            {
                var col = DataChoice(new[] { "(Không)" }.Concat(Enumerable.Range(0, ColumnCount).Select(x => ((char)('A' + x)).ToString())).ToArray());
                if (i == 0) col.SelectedIndex = grid.CurrentCell.ColumnIndex + 1;
                var order = DataChoice("A → Z / nhỏ → lớn", "Z → A / lớn → nhỏ");
                var kind = DataChoice("Auto", "Text", "Number", "Date", "Color");
                DataField(layout, "Cột cấp " + (i + 1), col); DataField(layout, "Thứ tự", order); DataField(layout, "Kiểu dữ liệu", kind);
                columns.Add(col); orders.Add(order); types.Add(kind);
            }
            using (var dialog = DataDialog("Sắp xếp nhiều cấp", layout))
            {
                dialog.Height = 470;
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                var keys = new List<Tuple<int, bool, string>>();
                for (int i = 0; i < columns.Count; i++) if (columns[i].SelectedIndex > 0)
                    keys.Add(Tuple.Create(columns[i].SelectedIndex - 1, orders[i].SelectedIndex == 1, types[i].Text));
                SortDataRows(keys);
            }
        }

        private void SortDataRows(List<Tuple<int, bool, string>> keys)
        {
            if (keys.Count == 0 || grid.CurrentCell == null) return;
            int first = 1, last = RowCount - 1;
            if (grid.SelectedCells.Count > 1)
            { first = Math.Max(1, grid.SelectedCells.Cast<DataGridViewCell>().Min(c => c.RowIndex)); last = grid.SelectedCells.Cast<DataGridViewCell>().Max(c => c.RowIndex); }
            if (tables.Any(t => t.Range.Top == first)) first++;
            while (last > first && Enumerable.Range(0, ColumnCount).All(c => string.IsNullOrEmpty(Convert.ToString(grid[c, last].Value)))) last--;
            if (last <= first) return;
            if (merges.Any(m => m.IntersectsWith(new Rectangle(0, first, ColumnCount, last - first + 1))))
            { MessageBox.Show(this, "Hãy bỏ gộp ô trong vùng trước khi sắp xếp.", "DinkCel"); return; }
            var rows = new List<Tuple<int, CellState[]>>();
            for (int r = first; r <= last; r++)
            {
                var cells = new CellState[ColumnCount];
                for (int c = 0; c < ColumnCount; c++)
                    cells[c] = new CellState { Value = grid[c, r].Value,
                        Style = grid[c, r].HasStyle ? new DataGridViewCellStyle(grid[c, r].Style) : null,
                        Extras = CellExtras.Copy(grid[c, r].Tag as CellExtras) };
                rows.Add(Tuple.Create(r, cells));
            }
            rows.Sort((a, b) =>
            {
                foreach (var key in keys)
                {
                    var ca = a.Item2[key.Item1]; var cb = b.Item2[key.Item1];
                    int cmp = key.Item3 == "Color" ?
                        (ca.Style == null ? 0 : ca.Style.BackColor.ToArgb()).CompareTo(cb.Style == null ? 0 : cb.Style.BackColor.ToArgb()) :
                        DataTools.Compare(Convert.ToString(ca.Value) ?? "", Convert.ToString(cb.Value) ?? "", key.Item3);
                    if (cmp != 0) return key.Item2 ? -cmp : cmp;
                }
                return a.Item1.CompareTo(b.Item1);
            });
            loading = true;
            try
            {
                for (int r = first; r <= last; r++) for (int c = 0; c < ColumnCount; c++)
                {
                    var source = rows[r - first].Item2[c]; var target = grid[c, r];
                    target.Value = source.Value;
                    target.Style = source.Style == null ? new DataGridViewCellStyle() : source.Style;
                    target.Tag = CellExtras.Copy(source.Extras);
                }
            }
            finally { loading = false; }
            Recalculate(); ApplyFreezeAndFilter(); RecordChange(); MarkDirty();
        }

        private void SetAdvancedFilter()
        {
            if (grid.CurrentCell == null) return;
            var layout = new TableLayoutPanel();
            var col = DataChoice(Enumerable.Range(0, ColumnCount).Select(i => ((char)('A' + i)).ToString()).ToArray());
            col.SelectedIndex = grid.CurrentCell.ColumnIndex;
            var kind = DataChoice("Text", "Number", "Date");
            var op = DataChoice("Contains", "Begins With", "Equals", "Greater", "Less", "Between", "Blank", "Nonblank");
            var v1 = new TextBox(); var v2 = new TextBox();
            DataField(layout, "Cột", col); DataField(layout, "Kiểu", kind); DataField(layout, "Điều kiện", op);
            DataField(layout, "Giá trị 1", v1); DataField(layout, "Giá trị 2", v2);
            using (var dialog = DataDialog("Lọc dữ liệu", layout))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                activeFilters.RemoveAll(f => f.Column == col.SelectedIndex);
                activeFilters.Add(new FilterCriterion { Column = col.SelectedIndex, Kind = kind.Text,
                    Operator = op.Text, Value1 = v1.Text, Value2 = v2.Text });
                filterColumn = -1; filterValue = "";
                ApplyFreezeAndFilter(); RecordChange(); MarkDirty();
            }
        }

        private bool FilterPasses(int row)
        {
            if (row == 0 || row < freezeRow) return true;
            if (filterColumn >= 0 && filterColumn < ColumnCount && !activeFilters.Any(f => f.Column == filterColumn) &&
                (Convert.ToString(grid[filterColumn, row].Value) ?? "").IndexOf(filterValue, StringComparison.CurrentCultureIgnoreCase) < 0) return false;
            return activeFilters.All(f => f.Column >= 0 && f.Column < ColumnCount &&
                DataTools.Match(CellDisplay(row, f.Column), f.Kind, f.Operator, f.Value1, f.Value2));
        }

        private string CellDisplay(int row, int column)
        {
            string value;
            return calculated.TryGetValue(row * ColumnCount + column, out value) ? value :
                Convert.ToString(grid[column, row].Value) ?? "";
        }

        private void ConfigureValidation()
        {
            Rectangle range = SelectionRange(false);
            if (range.IsEmpty) return;
            var layout = new TableLayoutPanel { AutoScroll = true };
            var kind = DataChoice("List", "Whole Number", "Decimal", "Date", "Time", "Text Length", "Custom Formula");
            var op = DataChoice("Between", "Not Between", "Equals", "Not Equals", "Greater", "Greater Or Equal", "Less", "Less Or Equal");
            var first = new TextBox(); var second = new TextBox();
            var blank = new CheckBox { Text = "Cho phép ô trống", Checked = true, AutoSize = true };
            var title = new TextBox(); var message = new TextBox(); var errorTitle = new TextBox(); var error = new TextBox();
            var style = DataChoice("Stop", "Warning", "Information");
            DataField(layout, "Loại", kind); DataField(layout, "Điều kiện", op);
            DataField(layout, "Danh sách / giá trị 1", first); DataField(layout, "Giá trị 2", second);
            DataField(layout, "Ô trống", blank); DataField(layout, "Tiêu đề gợi ý", title);
            DataField(layout, "Nội dung gợi ý", message); DataField(layout, "Tiêu đề lỗi", errorTitle);
            DataField(layout, "Thông báo lỗi", error); DataField(layout, "Mức cảnh báo", style);
            ValidationRule existing = validations.LastOrDefault(v => v.Range == range);
            if (existing != null)
            {
                kind.Text = existing.Kind; op.Text = existing.Operator;
                first.Text = existing.Kind == "List" ? string.Join(",", existing.Choices.ToArray()) : existing.Value1;
                second.Text = existing.Value2; blank.Checked = existing.AllowBlank;
                title.Text = existing.InputTitle; message.Text = existing.InputMessage;
                errorTitle.Text = existing.ErrorTitle; error.Text = existing.ErrorMessage; style.Text = existing.ErrorStyle;
            }
            using (var dialog = DataDialog("Kiểm tra dữ liệu", layout))
            {
                dialog.Height = 550;
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                var rule = new ValidationRule { Range = range, Kind = kind.Text, Operator = op.Text,
                    Value1 = first.Text, Value2 = second.Text, AllowBlank = blank.Checked,
                    InputTitle = title.Text, InputMessage = message.Text,
                    ErrorTitle = errorTitle.Text, ErrorMessage = error.Text, ErrorStyle = style.Text };
                if (rule.Kind == "List") rule.Choices.AddRange(first.Text.Split(',').Select(s => s.Trim()).Where(s => s.Length > 0));
                if (rule.Kind == "List" && rule.Choices.Count == 0) { MessageBox.Show(this, "Danh sách trống."); return; }
                if (rule.Kind == "Custom Formula" && !rule.Value1.StartsWith("=", StringComparison.Ordinal))
                { MessageBox.Show(this, "Công thức cần bắt đầu bằng dấu =."); return; }
                validations.RemoveAll(v => v.Range.IntersectsWith(range));
                validations.Add(rule); RecordChange(); MarkDirty();
            }
        }

        private void ClearValidation()
        {
            Rectangle range = SelectionRange(false);
            if (!range.IsEmpty && validations.RemoveAll(v => v.Range.IntersectsWith(range)) > 0)
            { RecordChange(); MarkDirty(); }
        }

        private bool EvaluateRuleFormula(string expression, int row, int column, string candidate, int baseRow, int baseColumn)
        {
            string shifted = FormulaEngine.ShiftReferences(expression, row - baseRow, column - baseColumn, RowCount, ColumnCount);
            string result = EvaluateAt(shifted, row, column, candidate);
            double number;
            return result.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ||
                (DataTools.Number(result, out number) && number != 0);
        }

        private string EvaluateAt(string expression, int row, int column, string candidate)
        {
            var engine = new FormulaEngine((r, c) => r == row && c == column ? candidate :
                Convert.ToString(grid[c, r].Value) ?? "", RowCount, ColumnCount);
            return engine.EvaluateExpression(expression);
        }

        private void ConfigureConditionalFormatting()
        {
            Rectangle range = SelectionRange(false);
            if (range.IsEmpty) return;
            var layout = new TableLayoutPanel();
            var kind = DataChoice("Equal", "Greater", "Less", "Between", "Duplicate", "Unique", "Text Contains", "Blank", "Formula", "Color Scale", "Data Bar", "Icon Set");
            var first = new TextBox(); var second = new TextBox();
            DataField(layout, "Quy tắc", kind); DataField(layout, "Giá trị 1 / công thức", first);
            DataField(layout, "Giá trị 2", second);
            using (var dialog = DataDialog("Định dạng có điều kiện", layout))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK) return;
                using (var picker = new ColorDialog { Color = Color.FromArgb(182, 230, 197) })
                {
                    if (picker.ShowDialog(this) != DialogResult.OK) return;
                    var rule = new ConditionalRule { Range = range, Kind = kind.Text,
                        Value1 = first.Text, Value2 = second.Text, Color = picker.Color };
                    double threshold;
                    if (DataTools.Number(first.Text, out threshold)) rule.Threshold = threshold;
                    if (rule.Kind == "Color Scale")
                    {
                        using (var secondPicker = new ColorDialog { Color = Color.FromArgb(248, 200, 120) })
                        {
                            if (secondPicker.ShowDialog(this) != DialogResult.OK) return;
                            rule.Color2 = secondPicker.Color;
                        }
                    }
                    conditionalRules.Add(rule); conditionalStatistics.Clear(); grid.Invalidate(); RecordChange(); MarkDirty();
                }
            }
        }

        private void ClearConditionalFormatting()
        {
            Rectangle range = SelectionRange(false);
            if (!range.IsEmpty && conditionalRules.RemoveAll(v => v.Range.IntersectsWith(range)) > 0)
            { conditionalStatistics.Clear(); grid.Invalidate(); RecordChange(); MarkDirty(); }
        }

        private readonly Dictionary<ConditionalRule, Tuple<double, double, Dictionary<string, int>>> conditionalStatistics =
            new Dictionary<ConditionalRule, Tuple<double, double, Dictionary<string, int>>>();

        private Tuple<double, double, Dictionary<string, int>> ConditionalStats(ConditionalRule rule)
        {
            Tuple<double, double, Dictionary<string, int>> result;
            if (conditionalStatistics.TryGetValue(rule, out result)) return result;
            double min = double.PositiveInfinity, max = double.NegativeInfinity;
            var counts = new Dictionary<string, int>(StringComparer.CurrentCultureIgnoreCase);
            for (int row = rule.Range.Top; row < Math.Min(RowCount, rule.Range.Bottom); row++)
                for (int col = rule.Range.Left; col < Math.Min(ColumnCount, rule.Range.Right); col++)
                {
                    string raw = CellDisplay(row, col);
                    if (raw.Length > 0) { int n; counts.TryGetValue(raw, out n); counts[raw] = n + 1; }
                    double value;
                    if (DataTools.Number(raw, out value)) { min = Math.Min(min, value); max = Math.Max(max, value); }
                }
            result = Tuple.Create(min, max, counts); conditionalStatistics[rule] = result;
            return result;
        }

        private void ApplyConditionalFormatting(DataGridViewCellFormattingEventArgs e, string raw)
        {
            foreach (ConditionalRule rule in conditionalRules)
            {
                if (!rule.Range.Contains(e.ColumnIndex, e.RowIndex)) continue;
                bool match = false;
                double value, a, b;
                bool numeric = DataTools.Number(raw, out value);
                bool hasA = DataTools.Number(rule.Value1, out a);
                bool hasB = DataTools.Number(rule.Value2, out b);
                var stats = rule.Kind == "Duplicate" || rule.Kind == "Unique" || rule.Kind == "Color Scale" ||
                    rule.Kind == "Data Bar" || rule.Kind == "Icon Set" ? ConditionalStats(rule) : null;
                switch (rule.Kind)
                {
                    case "Greater": match = numeric && value > (hasA ? a : rule.Threshold); break;
                    case "Less": match = numeric && hasA && value < a; break;
                    case "Equal": match = DataTools.Match(raw, "Auto", "Equals", rule.Value1, ""); break;
                    case "Between": match = numeric && hasA && hasB && value >= a && value <= b; break;
                    case "Text Contains": match = raw.IndexOf(rule.Value1, StringComparison.CurrentCultureIgnoreCase) >= 0; break;
                    case "Blank": match = raw.Length == 0; break;
                    case "Formula": match = EvaluateRuleFormula(rule.Value1, e.RowIndex, e.ColumnIndex,
                        Convert.ToString(grid[e.ColumnIndex, e.RowIndex].Value) ?? "", rule.Range.Top, rule.Range.Left); break;
                    case "Duplicate": case "Unique":
                        int count; stats.Item3.TryGetValue(raw, out count);
                        match = raw.Length > 0 && (rule.Kind == "Duplicate" ? count > 1 : count == 1); break;
                    case "Color Scale": case "Data Bar": case "Icon Set":
                        if (numeric && !double.IsInfinity(stats.Item1))
                        {
                            double fraction = stats.Item2 <= stats.Item1 ? .5 :
                                Math.Max(0, Math.Min(1, (value - stats.Item1) / (stats.Item2 - stats.Item1)));
                            if (rule.Kind == "Icon Set") e.Value = (fraction < .33 ? "▼ " : fraction < .67 ? "● " : "▲ ") + raw;
                            else e.CellStyle.BackColor = Blend(rule.Color2.IsEmpty ? Color.White : rule.Color2, rule.Color, fraction);
                            if (rule.Kind != "Color Scale") e.CellStyle.BackColor = Blend(Color.White, rule.Color, .2 + fraction * .45);
                            e.FormattingApplied = true;
                        }
                        break;
                }
                if (match) e.CellStyle.BackColor = rule.Color;
            }
        }

        private static Color Blend(Color from, Color to, double fraction)
        {
            return Color.FromArgb((int)(from.R + (to.R - from.R) * fraction),
                (int)(from.G + (to.G - from.G) * fraction), (int)(from.B + (to.B - from.B) * fraction));
        }

        private bool PaintDataBar(DataGridViewCellPaintingEventArgs e)
        {
            foreach (ConditionalRule rule in conditionalRules)
            {
                if (rule.Kind != "Data Bar" || !rule.Range.Contains(e.ColumnIndex, e.RowIndex)) continue;
                double value;
                if (!DataTools.Number(CellDisplay(e.RowIndex, e.ColumnIndex), out value)) continue;
                var stats = ConditionalStats(rule);
                if (double.IsInfinity(stats.Item1)) continue;
                double fraction = stats.Item2 <= stats.Item1 ? .5 :
                    Math.Max(0, Math.Min(1, (value - stats.Item1) / (stats.Item2 - stats.Item1)));
                e.Paint(e.CellBounds, DataGridViewPaintParts.Background | DataGridViewPaintParts.SelectionBackground |
                    DataGridViewPaintParts.Border | DataGridViewPaintParts.Focus);
                Rectangle bar = Rectangle.Inflate(e.CellBounds, -3, -4);
                bar.Width = Math.Max(1, (int)(bar.Width * fraction));
                using (var brush = new SolidBrush(Color.FromArgb(175, rule.Color)))
                    e.Graphics.FillRectangle(brush, bar);
                TextRenderer.DrawText(e.Graphics, Convert.ToString(e.FormattedValue) ?? "",
                    e.CellStyle.Font ?? grid.Font, Rectangle.Inflate(e.CellBounds, -4, -2),
                    e.State.HasFlag(DataGridViewElementStates.Selected) ? e.CellStyle.SelectionForeColor : e.CellStyle.ForeColor,
                    TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
                e.Handled = true;
                return true;
            }
            return false;
        }

        private sealed class SearchHit
        {
            public int Sheet, Key;
            public string Display;
            public override string ToString() { return Display; }
        }

        private List<SearchHit> SearchWorkbook(string query, bool allSheets, bool formulas,
            bool matchCase, bool whole, bool wildcard)
        {
            SaveActiveSheet();
            var matches = new List<SearchHit>();
            for (int i = 0; i < sheets.Count; i++)
            {
                if (!allSheets && i != activeSheetIndex) continue;
                SheetState sheet = sheets[i];
                int index = i;
                var engine = formulas ? null : new FormulaEngine((r, c) =>
                {
                    CellState state;
                    return sheet.Cells.TryGetValue(r * ColumnCount + c, out state) ? Convert.ToString(state.Value) ?? "" : "";
                }, (name, r, c) =>
                {
                    SheetState other = sheets.FirstOrDefault(s => string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase));
                    CellState state;
                    return other == null ? null : other.Cells.TryGetValue(r * ColumnCount + c, out state) ?
                        Convert.ToString(state.Value) ?? "" : "";
                }, sheet.Name, RowCount, ColumnCount);
                foreach (var pair in sheet.Cells.OrderBy(p => p.Key))
                {
                    int row = pair.Key / ColumnCount, col = pair.Key % ColumnCount;
                    string raw = Convert.ToString(pair.Value.Value) ?? "";
                    string value = !formulas && raw.StartsWith("=", StringComparison.Ordinal) ? engine.Display(row, col) : raw;
                    if (!DataTools.FindMatch(value, query, matchCase, whole, wildcard)) continue;
                    matches.Add(new SearchHit { Sheet = index, Key = pair.Key,
                        Display = sheet.Name + "!" + (char)('A' + col) + (row + 1) + "  " + value });
                }
            }
            return matches;
        }

        private void NavigateHit(SearchHit hit)
        {
            if (hit.Sheet != activeSheetIndex) SwitchSheet(hit.Sheet);
            int row = hit.Key / ColumnCount, col = hit.Key % ColumnCount;
            if (!grid.Rows[row].Visible) { status.Text = "Ô tìm thấy đang bị ẩn bởi bộ lọc"; return; }
            grid.ClearSelection(); grid.CurrentCell = grid[col, row]; grid[col, row].Selected = true;
            grid.FirstDisplayedScrollingRowIndex = row;
            status.Text = "Đã tìm thấy " + hit.Display;
        }

        private bool AcceptOnSheet(SheetState sheet, ValidationRule rule, int row, int col, string candidate)
        {
            if (rule == null) return true;
            if (candidate.Length == 0) return rule.AllowBlank;
            var engine = new FormulaEngine((r, c) =>
            {
                if (r == row && c == col) return candidate;
                CellState state;
                return sheet.Cells.TryGetValue(r * ColumnCount + c, out state) ? Convert.ToString(state.Value) ?? "" : "";
            }, RowCount, ColumnCount);
            if (rule.Kind == "Custom Formula")
            {
                string shifted = FormulaEngine.ShiftReferences(rule.Value1, row - rule.Range.Top,
                    col - rule.Range.Left, RowCount, ColumnCount);
                string answer = engine.EvaluateExpression(shifted);
                double numeric;
                return answer.Equals("TRUE", StringComparison.OrdinalIgnoreCase) ||
                    DataTools.Number(answer, out numeric) && numeric != 0;
            }
            string value = candidate.StartsWith("=", StringComparison.Ordinal) ? engine.EvaluateExpression(candidate) : candidate;
            return DataTools.Valid(rule, value);
        }

        private int ReplaceHits(IEnumerable<SearchHit> hits, string query, string replacement,
            bool formulas, bool matchCase, bool whole, bool wildcard)
        {
            SaveActiveSheet();
            int changed = 0;
            foreach (SearchHit hit in hits)
            {
                SheetState sheet = sheets[hit.Sheet];
                CellState cell;
                if (!sheet.Cells.TryGetValue(hit.Key, out cell)) continue;
                string raw = Convert.ToString(cell.Value) ?? "";
                if (!formulas && raw.StartsWith("=", StringComparison.Ordinal)) continue;
                string value = DataTools.Replace(raw, query, replacement, matchCase, whole, wildcard);
                if (value == raw) continue;
                int row = hit.Key / ColumnCount, col = hit.Key % ColumnCount;
                ValidationRule rule = sheet.Validations.LastOrDefault(v => v.Range.Contains(col, row));
                if (rule != null && hit.Sheet == activeSheetIndex && !CanAcceptValue(row, col, value)) continue;
                if (rule != null && hit.Sheet != activeSheetIndex &&
                    !AcceptOnSheet(sheet, rule, row, col, value)) continue;
                cell.Value = value; changed++;
            }
            if (changed > 0)
            {
                otherSheetsDirty = true;
                RestoreSheet(sheets[activeSheetIndex]);
                RecordChange(); MarkDirty();
            }
            return changed;
        }

        private void OpenFindReplace()
        {
            var form = new Form { Text = "Tìm và thay thế", Width = 570, Height = 520,
                StartPosition = FormStartPosition.CenterParent, BackColor = theme.Surface,
                ForeColor = theme.Text, FormBorderStyle = FormBorderStyle.SizableToolWindow };
            var top = new TableLayoutPanel { Dock = DockStyle.Top, Height = 275, ColumnCount = 2, Padding = new Padding(10) };
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 32));
            top.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 68));
            var query = new TextBox(); var replacement = new TextBox();
            var scope = DataChoice("Trang hiện tại", "Toàn bộ workbook");
            var mode = DataChoice("Giá trị hiển thị", "Công thức / dữ liệu gốc");
            var matchCase = new CheckBox { Text = "Phân biệt hoa thường", AutoSize = true };
            var whole = new CheckBox { Text = "Khớp toàn bộ ô", AutoSize = true };
            var wildcard = new CheckBox { Text = "Ký tự đại diện * ?", AutoSize = true };
            DataField(top, "Tìm", query); DataField(top, "Thay bằng", replacement);
            DataField(top, "Phạm vi", scope); DataField(top, "Tìm trong", mode);
            DataField(top, "Tùy chọn", matchCase); DataField(top, "", whole); DataField(top, "", wildcard);
            var results = new ListBox { Dock = DockStyle.Fill, HorizontalScrollbar = true };
            var bottom = new FlowLayoutPanel { Dock = DockStyle.Bottom, Height = 48, FlowDirection = FlowDirection.LeftToRight };
            var next = new Button { Text = "Tìm tiếp", Width = 95 };
            var all = new Button { Text = "Tìm tất cả", Width = 95 };
            var replace = new Button { Text = "Thay", Width = 80 };
            var replaceAll = new Button { Text = "Thay tất cả", Width = 95 };
            bottom.Controls.Add(next); bottom.Controls.Add(all); bottom.Controls.Add(replace); bottom.Controls.Add(replaceAll);
            form.Controls.Add(results); form.Controls.Add(top); form.Controls.Add(bottom);
            Func<List<SearchHit>> search = () => SearchWorkbook(query.Text, scope.SelectedIndex == 1,
                mode.SelectedIndex == 1, matchCase.Checked, whole.Checked, wildcard.Checked);
            Action<bool> runFind = showAll =>
            {
                if (query.Text.Length == 0) return;
                var matches = search(); results.Items.Clear();
                foreach (SearchHit hit in matches) results.Items.Add(hit);
                if (matches.Count == 0) { status.Text = "Không tìm thấy kết quả"; return; }
                if (showAll) { status.Text = "Tìm thấy " + matches.Count + " ô"; return; }
                int current = grid.CurrentCell == null ? -1 : grid.CurrentCell.RowIndex * ColumnCount + grid.CurrentCell.ColumnIndex;
                SearchHit chosen = matches.FirstOrDefault(h => h.Sheet > activeSheetIndex ||
                    h.Sheet == activeSheetIndex && h.Key > current) ?? matches[0];
                results.SelectedItem = chosen; NavigateHit(chosen);
            };
            next.Click += delegate { runFind(false); };
            all.Click += delegate { runFind(true); };
            results.DoubleClick += delegate { if (results.SelectedItem is SearchHit) NavigateHit((SearchHit)results.SelectedItem); };
            replace.Click += delegate
            {
                if (query.Text.Length == 0) return;
                SearchHit hit = results.SelectedItem as SearchHit;
                if (hit == null) { runFind(false); hit = results.SelectedItem as SearchHit; }
                if (hit == null) return;
                int count = ReplaceHits(new[] { hit }, query.Text, replacement.Text,
                    mode.SelectedIndex == 1, matchCase.Checked, whole.Checked, wildcard.Checked);
                status.Text = "Đã thay " + count + " ô"; runFind(false);
            };
            replaceAll.Click += delegate
            {
                if (query.Text.Length == 0) return;
                int count = ReplaceHits(search(), query.Text, replacement.Text,
                    mode.SelectedIndex == 1, matchCase.Checked, whole.Checked, wildcard.Checked);
                status.Text = "Đã thay " + count + " ô"; results.Items.Clear();
            };
            form.ShowDialog(this); form.Dispose();
        }
    }
}
