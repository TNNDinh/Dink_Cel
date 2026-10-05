using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DinkCel
{
    internal sealed partial class SpreadsheetForm
    {
        private void AddAdvancedExcelMenus()
        {
            ToolStripMenuItem data = menu.Items.OfType<ToolStripMenuItem>()
                .FirstOrDefault(item => item.Text == "Dữ liệu");
            if (data != null)
            {
                data.DropDownItems.Add(new ToolStripSeparator());
                AddMenuItem(data, "Tách văn bản thành cột...", Keys.None, TextToColumns);
                AddMenuItem(data, "Xóa dòng trùng...", Keys.None, RemoveDuplicateRows);
                AddMenuItem(data, "Điền nhanh (Flash Fill)", Keys.Control | Keys.E, FlashFill);
                AddMenuItem(data, "Lọc nâng cao theo vùng điều kiện...", Keys.None, AdvancedCriteriaFilter);
            }
            var whatIf = new ToolStripMenuItem("What-if");
            AddMenuItem(whatIf, "Goal Seek...", Keys.None, GoalSeek);
            AddMenuItem(whatIf, "Data Table từ vùng chọn...", Keys.None, CreateDataTable);
            AddMenuItem(whatIf, "Lưu Scenario...", Keys.None, SaveScenario);
            AddMenuItem(whatIf, "Áp dụng Scenario...", Keys.None, ApplyScenario);
            AddMenuItem(whatIf, "Xóa Scenario...", Keys.None, DeleteScenario);
            menu.Items.Add(whatIf);
        }

        private static Rectangle CellAddress(string address)
        {
            Rectangle result = ParseScriptRange(address);
            if (result.Width != 1 || result.Height != 1)
                throw new ArgumentException("Enter one cell address, such as A1.");
            return result;
        }

        private double EvaluateWhatIf(int targetRow, int targetColumn,
            IDictionary<int, string> overrides)
        {
            var engine = new FormulaEngine((row, column) =>
            {
                string replacement;
                if (overrides.TryGetValue(row * ColumnCount + column, out replacement)) return replacement;
                return row < RowCount ? Convert.ToString(grid[column, row].Value) ?? "" : "";
            }, (sheetName, row, column) =>
            {
                SheetState state = sheets.FirstOrDefault(sheet =>
                    String.Equals(sheet.Name, sheetName, StringComparison.OrdinalIgnoreCase));
                if (state == null) return null;
                CellState cell;
                return state.Cells.TryGetValue(row * ColumnCount + column, out cell) ?
                    Convert.ToString(cell.Value) ?? "" : "";
            }, sheets[activeSheetIndex].Name, MaxRowCount, ColumnCount,
                name =>
                {
                    NamedRange named = FindName(name, "");
                    if (named != null && !String.IsNullOrEmpty(named.Formula)) return null;
                    return named == null ? null : new FormulaNamedRange { Sheet = named.Sheet,
                        FirstRow = named.Range.Top, LastRow = named.Range.Bottom - 1,
                        FirstColumn = named.Range.Left, LastColumn = named.Range.Right - 1 };
                });
            engine.ResolveScopedRange = (name, sheet) =>
            {
                NamedRange named = FindName(name, sheet);
                return named == null || !String.IsNullOrEmpty(named.Formula) ? null :
                    new FormulaNamedRange { Sheet = named.Sheet,
                        FirstRow = named.Range.Top, LastRow = named.Range.Bottom - 1,
                        FirstColumn = named.Range.Left, LastColumn = named.Range.Right - 1 };
            };
            engine.ResolveNamedFormula = (name, sheet) =>
            {
                NamedRange named = FindName(name, sheet);
                return named == null ? null : named.Formula;
            };
            string answer = engine.Display(targetRow, targetColumn);
            double number;
            if (!Double.TryParse(answer, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
                throw new InvalidOperationException("The target formula does not return a number: " + answer);
            return number;
        }

        private void GoalSeek()
        {
            if (grid.CurrentCell == null || grid.ReadOnly) return;
            int targetRow = grid.CurrentCell.RowIndex, targetCol = grid.CurrentCell.ColumnIndex;
            string formula = Convert.ToString(grid[targetCol, targetRow].Value) ?? "";
            if (!formula.StartsWith("=", StringComparison.Ordinal))
            { MessageBox.Show(this, "Select a formula cell first."); return; }
            string goalText = Prompt("Giá trị mục tiêu", "100");
            if (goalText == null) return;
            double goal;
            if (!DataTools.Number(goalText, out goal))
            { MessageBox.Show(this, "The target must be a number."); return; }
            string variableText = Prompt("Ô thay đổi (ví dụ A1)", "A1");
            if (variableText == null) return;
            try
            {
                Rectangle variable = CellAddress(variableText);
                if (variable.X == targetCol && variable.Y == targetRow)
                    throw new InvalidOperationException("The variable must differ from the formula cell.");
                string raw = Convert.ToString(grid[variable.X, variable.Y].Value) ?? "";
                if (raw.StartsWith("=", StringComparison.Ordinal))
                    throw new InvalidOperationException("The variable cell must contain a value, not a formula.");
                double initial;
                if (!DataTools.Number(raw, out initial)) initial = 0;
                int key = variable.Y * ColumnCount + variable.X;
                double answer;
                bool solved = GoalSeekSolver.Solve(value => EvaluateWhatIf(targetRow, targetCol,
                    new Dictionary<int, string> { { key, value.ToString("R", CultureInfo.InvariantCulture) } }),
                    goal, initial, out answer);
                if (!solved) { MessageBox.Show(this, "Goal Seek did not converge."); return; }
                string text = answer.ToString("0.##########", CultureInfo.InvariantCulture);
                if (MessageBox.Show(this, "Set " + variableText.ToUpperInvariant() + " to " + text + "?",
                    "Goal Seek", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
                ApplyScriptWrites(new List<ScriptCellWrite> { new ScriptCellWrite
                { Row = variable.Y, Column = variable.X, Value = text } });
            }
            catch (Exception error) { MessageBox.Show(this, error.Message, "Goal Seek"); }
        }

        private void CreateDataTable()
        {
            Rectangle range = SelectionRange(false);
            if (range.Width < 2 || range.Height < 2 || (long)(range.Width - 1) * (range.Height - 1) > 500)
            { MessageBox.Show(this, "Select a 2D table of at most 500 result cells."); return; }
            string formula = Convert.ToString(grid[range.Left, range.Top].Value) ?? "";
            if (!formula.StartsWith("=", StringComparison.Ordinal))
            { MessageBox.Show(this, "The top-left cell must contain the result formula."); return; }
            string rowInput = Prompt("Input cell for values in the top row (blank = unused)", "A1");
            if (rowInput == null) return;
            string columnInput = Prompt("Input cell for values in the first column (blank = unused)", "");
            if (columnInput == null) return;
            try
            {
                Rectangle rowCell = String.IsNullOrWhiteSpace(rowInput) ? Rectangle.Empty : CellAddress(rowInput);
                Rectangle columnCell = String.IsNullOrWhiteSpace(columnInput) ? Rectangle.Empty : CellAddress(columnInput);
                if (rowCell.IsEmpty && columnCell.IsEmpty)
                    throw new ArgumentException("Choose at least one input cell.");
                var writes = new List<ScriptCellWrite>();
                for (int r = range.Top + 1; r < range.Bottom; r++)
                    for (int c = range.Left + 1; c < range.Right; c++)
                    {
                        var overrides = new Dictionary<int, string>();
                        if (!rowCell.IsEmpty) overrides[rowCell.Y * ColumnCount + rowCell.X] =
                            Convert.ToString(grid[c, range.Top].Value) ?? "";
                        if (!columnCell.IsEmpty) overrides[columnCell.Y * ColumnCount + columnCell.X] =
                            Convert.ToString(grid[range.Left, r].Value) ?? "";
                        double value = EvaluateWhatIf(range.Top, range.Left, overrides);
                        writes.Add(new ScriptCellWrite { Row = r, Column = c,
                            Value = value.ToString("0.##########", CultureInfo.InvariantCulture) });
                    }
                ApplyScriptWrites(writes);
            }
            catch (Exception error) { MessageBox.Show(this, error.Message, "Data Table"); }
        }

        private void TextToColumns()
        {
            Rectangle range = SelectionRange(false);
            if (range.Width != 1 || grid.ReadOnly) { MessageBox.Show(this, "Select one column."); return; }
            string delimiter = Prompt("Ký tự phân cách: , ; hoặc Tab", ",");
            if (delimiter == null) return;
            char separator = delimiter.Equals("Tab", StringComparison.OrdinalIgnoreCase) ? '\t' :
                delimiter.Length == 1 ? delimiter[0] : '\0';
            if (separator == '\0') { MessageBox.Show(this, "Enter one delimiter character."); return; }
            var parts = new List<string[]>();
            int width = 1;
            for (int row = range.Top; row < range.Bottom; row++)
            {
                string[] split = AdvancedData.SplitLine(Convert.ToString(grid[range.Left, row].Value) ?? "", separator);
                parts.Add(split); width = Math.Max(width, split.Length);
            }
            if (range.Left + width > ColumnCount)
            { MessageBox.Show(this, "Result exceeds the 26-column grid."); return; }
            if (width > 1 && Enumerable.Range(range.Top, range.Height).Any(row =>
                Enumerable.Range(range.Left + 1, width - 1).Any(col =>
                    !String.IsNullOrEmpty(Convert.ToString(grid[col, row].Value)))) &&
                MessageBox.Show(this, "Overwrite cells to the right?", "Text to Columns",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            var writes = new List<ScriptCellWrite>();
            for (int r = 0; r < parts.Count; r++)
                for (int c = 0; c < width; c++)
                    writes.Add(new ScriptCellWrite { Row = range.Top + r, Column = range.Left + c,
                        Value = c < parts[r].Length ? parts[r][c] : null });
            ApplyScriptWrites(writes);
        }

        private void RemoveDuplicateRows()
        {
            Rectangle range = SelectionRange(true);
            if (range.IsEmpty || range.Height < 2 || grid.ReadOnly) return;
            bool header = MessageBox.Show(this, "Does the first row contain headers?", "Remove Duplicates",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
            var raw = new string[range.Height][];
            var saved = new CellState[range.Height, range.Width];
            for (int r = 0; r < range.Height; r++)
            {
                raw[r] = new string[range.Width];
                for (int c = 0; c < range.Width; c++)
                {
                    saved[r, c] = CaptureCell(range.Left + c, range.Top + r);
                    raw[r][c] = Convert.ToString(saved[r, c].Value) ?? "";
                }
            }
            List<int> unique = AdvancedData.UniqueRows(raw, header);
            if (unique.Count == range.Height) { status.Text = "No duplicate rows found"; return; }
            loading = true;
            try
            {
                for (int r = 0; r < range.Height; r++)
                    for (int c = 0; c < range.Width; c++)
                    {
                        if (r < unique.Count) RestoreCell(range.Left + c, range.Top + r, saved[unique[r], c]);
                        else ClearCell(range.Left + c, range.Top + r);
                        changedCells.Add((range.Top + r) * ColumnCount + range.Left + c);
                    }
            }
            finally { loading = false; }
            RecordChange(); Recalculate(); MarkDirty();
            status.Text = "Removed " + (range.Height - unique.Count) + " duplicate rows";
        }

        private void FlashFill()
        {
            Rectangle range = SelectionRange(false);
            if (range.Width != 1 || range.Left == 0 || grid.ReadOnly)
            { MessageBox.Show(this, "Select a destination column to the right of source data."); return; }
            var examples = new List<Tuple<string, string, string>>();
            for (int row = range.Top; row < range.Bottom; row++)
            {
                string target = Convert.ToString(grid[range.Left, row].Value) ?? "";
                if (target.Length > 0) examples.Add(Tuple.Create(
                    Convert.ToString(grid[range.Left - 1, row].Value) ?? "",
                    range.Left > 1 ? Convert.ToString(grid[range.Left - 2, row].Value) ?? "" : "", target));
            }
            if (examples.Count == 0) { MessageBox.Show(this, "Enter at least one example in the destination column."); return; }
            Func<string, string, string> pattern = AdvancedData.InferFlashFill(examples);
            if (pattern == null) { MessageBox.Show(this, "Could not infer a pattern from the examples."); return; }
            var writes = new List<ScriptCellWrite>();
            for (int row = range.Top; row < range.Bottom; row++)
                if (String.IsNullOrEmpty(Convert.ToString(grid[range.Left, row].Value)))
                    writes.Add(new ScriptCellWrite { Row = row, Column = range.Left,
                        Value = pattern(Convert.ToString(grid[range.Left - 1, row].Value) ?? "",
                            range.Left > 1 ? Convert.ToString(grid[range.Left - 2, row].Value) ?? "" : "") });
            ApplyScriptWrites(writes);
        }

        private void AdvancedCriteriaFilter()
        {
            Rectangle data = SelectionRange(true);
            if (data.Height < 2) { MessageBox.Show(this, "Select a data range with headers."); return; }
            string address = Prompt("Vùng điều kiện (hàng đầu là tên cột)", "H1:J3");
            if (address == null) return;
            try
            {
                Rectangle criteria = ParseScriptRange(address);
                if (criteria.Height < 2) throw new ArgumentException("Criteria need a header and at least one row.");
                string[] headers = Enumerable.Range(data.Left, data.Width)
                    .Select(col => Convert.ToString(grid[col, data.Top].Value) ?? "").ToArray();
                var alternatives = new List<Dictionary<string, string>>();
                for (int row = criteria.Top + 1; row < criteria.Bottom; row++)
                {
                    var condition = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    for (int col = criteria.Left; col < criteria.Right; col++)
                    {
                        string header = Convert.ToString(grid[col, criteria.Top].Value) ?? "";
                        string test = Convert.ToString(grid[col, row].Value) ?? "";
                        if (header.Length > 0 && test.Length > 0) condition[header] = test;
                    }
                    if (condition.Count > 0) alternatives.Add(condition);
                }
                if (alternatives.Count == 0) throw new ArgumentException("No criteria were entered.");
                for (int row = data.Top + 1; row < data.Bottom; row++)
                {
                    string[] values = Enumerable.Range(data.Left, data.Width)
                        .Select(col => CellDisplay(row, col)).ToArray();
                    manualHiddenRows[row] = !AdvancedData.MatchesCriteria(values, headers, alternatives);
                }
                ApplyFreezeAndFilter(); RecordChange(); MarkDirty();
            }
            catch (Exception error) { MessageBox.Show(this, error.Message, "Advanced Filter"); }
        }

        private void SaveScenario()
        {
            if (grid.SelectedCells.Count == 0) return;
            string name = Prompt("Tên Scenario", "Scenario " + (scenarios.Count + 1));
            if (String.IsNullOrWhiteSpace(name)) return;
            if (scenarios.Any(item => String.Equals(item.Name, name, StringComparison.OrdinalIgnoreCase)))
            { MessageBox.Show(this, "Scenario name already exists."); return; }
            var scenario = new ScenarioDefinition { Name = name.Trim(), Sheet = sheets[activeSheetIndex].Name };
            foreach (DataGridViewCell cell in grid.SelectedCells)
                scenario.Values[cell.RowIndex * ColumnCount + cell.ColumnIndex] =
                    Convert.ToString(cell.Value) ?? "";
            if (scenario.Values.Count > 1000) { MessageBox.Show(this, "A scenario supports up to 1,000 cells."); return; }
            scenarios.Add(scenario); MarkDirty();
        }

        private void ApplyScenario()
        {
            string name = ChooseOption("Scenario", scenarios.Select(s => s.Name).ToList());
            ScenarioDefinition scenario = scenarios.FirstOrDefault(s => s.Name == name);
            if (scenario == null) return;
            int index = sheets.FindIndex(s => s.Name == scenario.Sheet);
            if (index < 0) { MessageBox.Show(this, "The scenario sheet no longer exists."); return; }
            SwitchSheet(index);
            ApplyScriptWrites(scenario.Values.Select(x => new ScriptCellWrite { Row = x.Key / ColumnCount,
                Column = x.Key % ColumnCount, Value = x.Value }).ToList());
        }

        private void DeleteScenario()
        {
            string name = ChooseOption("Delete Scenario", scenarios.Select(s => s.Name).ToList());
            if (name == null) return;
            scenarios.RemoveAll(s => s.Name == name); MarkDirty();
        }

        private void SerializeScenarios(XElement root)
        {
            foreach (ScenarioDefinition scenario in scenarios)
            {
                var element = new XElement("scenario", new XAttribute("name", scenario.Name),
                    new XAttribute("sheet", scenario.Sheet));
                foreach (var value in scenario.Values)
                    element.Add(new XElement("input", new XAttribute("row", value.Key / ColumnCount),
                        new XAttribute("column", value.Key % ColumnCount), value.Value));
                root.Add(element);
            }
        }

        private static void ReadScenarios(XElement root, WorkbookSnapshot workbook)
        {
            foreach (XElement element in root.Elements("scenario"))
            {
                var scenario = new ScenarioDefinition { Name = (string)element.Attribute("name") ?? "",
                    Sheet = (string)element.Attribute("sheet") ?? "" };
                foreach (XElement input in element.Elements("input"))
                {
                    int row = (int?)input.Attribute("row") ?? -1;
                    int column = (int?)input.Attribute("column") ?? -1;
                    if (row >= 0 && row < MaxRowCount && column >= 0 && column < ColumnCount)
                        scenario.Values[row * ColumnCount + column] = input.Value;
                }
                if (scenario.Name.Length > 0) workbook.Scenarios.Add(scenario);
            }
        }
    }
}
