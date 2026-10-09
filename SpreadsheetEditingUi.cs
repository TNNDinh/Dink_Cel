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
    internal enum PasteKind { All, Values, Formulas, Formats, Transpose }

    internal sealed class CrossSheetMove
    {
        public int SourceSheetIndex;
        public readonly Dictionary<int, CellState> OriginalCells = new Dictionary<int, CellState>();
    }

    internal sealed partial class SpreadsheetForm
    {
        private bool formulaBarChanged;
        private bool pendingEditMove;
        private int editMoveRow, editMoveColumn;
        private int selectionAnchorRow, selectionAnchorColumn;
        private bool selectingByKeyboard;
        private bool[,] copiedMask;
        private int copiedTop, copiedLeft, copiedSheetIndex;
        private bool cutPending;
        private string copiedToken;
        private Rectangle fillSelection;
        private Rectangle fillDragSource;
        private readonly Dictionary<string, CrossSheetMove> crossSheetMoves =
            new Dictionary<string, CrossSheetMove>();

        private string MoveKey(long revision)
        { return activeSheetIndex.ToString(CultureInfo.InvariantCulture) + ":" +
            revision.ToString(CultureInfo.InvariantCulture); }

        private void BeforeUndoCrossSheetMove()
        {
            CrossSheetMove move;
            if (!crossSheetMoves.TryGetValue(MoveKey(lastState.RevisionId), out move) ||
                move.SourceSheetIndex >= sheets.Count) return;
            SheetState source = sheets[move.SourceSheetIndex];
            foreach (var entry in move.OriginalCells)
                if (!source.Cells.ContainsKey(entry.Key)) source.Cells[entry.Key] = entry.Value;
            otherSheetsDirty = true;
        }

        private void AfterRedoCrossSheetMove(long revision)
        {
            CrossSheetMove move;
            if (!crossSheetMoves.TryGetValue(MoveKey(revision), out move) ||
                move.SourceSheetIndex >= sheets.Count) return;
            SheetState source = sheets[move.SourceSheetIndex];
            foreach (var entry in move.OriginalCells)
            {
                CellState current;
                if (source.Cells.TryGetValue(entry.Key, out current) &&
                    object.Equals(current.Value, entry.Value.Value)) source.Cells.Remove(entry.Key);
            }
            otherSheetsDirty = true;
        }

        private static string CellAddress(int column, int row)
        {
            return ((char)('A' + column)).ToString() + (row + 1).ToString(CultureInfo.InvariantCulture);
        }

        private static bool ParseCellAddress(string text, out int column, out int row)
        {
            column = row = -1;
            Match match = Regex.Match(text.Trim(), @"^\$?([A-Z])\$?([1-9][0-9]{0,4})$", RegexOptions.IgnoreCase);
            if (!match.Success) return false;
            column = char.ToUpperInvariant(match.Groups[1].Value[0]) - 'A';
            row = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture) - 1;
            return row < MaxRowCount && column < ColumnCount;
        }

        private static Rectangle SelectedRectangle(DataGridView view, bool requireContiguous)
        {
            if (view.CurrentCell == null) return Rectangle.Empty;
            if (view.SelectedCells.Count == 0)
                return new Rectangle(view.CurrentCell.ColumnIndex, view.CurrentCell.RowIndex, 1, 1);
            int left = view.SelectedCells.Cast<DataGridViewCell>().Min(c => c.ColumnIndex);
            int right = view.SelectedCells.Cast<DataGridViewCell>().Max(c => c.ColumnIndex);
            int top = view.SelectedCells.Cast<DataGridViewCell>().Min(c => c.RowIndex);
            int bottom = view.SelectedCells.Cast<DataGridViewCell>().Max(c => c.RowIndex);
            Rectangle range = new Rectangle(left, top, right - left + 1, bottom - top + 1);
            return requireContiguous && view.SelectedCells.Count != range.Width * range.Height ?
                Rectangle.Empty : range;
        }

        private void SelectRectangle(Rectangle range, int activeColumn, int activeRow, bool keepExisting)
        {
            if (range.IsEmpty || range.Left < 0 || range.Top < 0 ||
                range.Right > ColumnCount || range.Bottom > RowCount) return;
            var previous = keepExisting ? grid.SelectedCells.Cast<DataGridViewCell>()
                .Select(cell => cell.RowIndex * ColumnCount + cell.ColumnIndex).ToArray() : null;
            selectingByKeyboard = true;
            selectingHeader = true;
            try
            {
                if (!keepExisting) grid.ClearSelection();
                grid.CurrentCell = grid[activeColumn, activeRow];
                if (previous != null)
                    foreach (int index in previous)
                        grid[index % ColumnCount, index / ColumnCount].Selected = true;
                for (int row = range.Top; row < range.Bottom; row++)
                    for (int column = range.Left; column < range.Right; column++)
                        grid[column, row].Selected = true;
            }
            finally { selectingHeader = false; selectingByKeyboard = false; }
            UpdateSelection();
            grid.Invalidate();
        }

        private void AddressBoxKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                if (grid.CurrentCell != null)
                    addressBox.Text = CellAddress(grid.CurrentCell.ColumnIndex, grid.CurrentCell.RowIndex);
                grid.Focus(); e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                if (GoToAddress(addressBox.Text))
                { grid.Focus(); UpdateSelection(); }
                else status.Text = "Địa chỉ hoặc tên vùng không hợp lệ";
                e.SuppressKeyPress = true;
            }
        }

        private bool GoToAddress(string input)
        {
            string[] parts = input.Split(new[] { ',', ';' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return false;
            var requests = new List<Tuple<int, Rectangle>>();
            foreach (string raw in parts)
            {
                string part = raw.Trim();
                NamedRange named = namedRanges.FirstOrDefault(n =>
                    string.Equals(n.Name, part, StringComparison.OrdinalIgnoreCase));
                if (named != null)
                {
                    int index = sheets.FindIndex(s => string.Equals(s.Name, named.Sheet,
                        StringComparison.OrdinalIgnoreCase));
                    if (index < 0 || named.Range.IsEmpty) return false;
                    requests.Add(Tuple.Create(index, named.Range));
                    continue;
                }
                string[] ends = part.Split(':');
                if (ends.Length > 2) return false;
                int x1, y1, x2, y2;
                if (!ParseCellAddress(ends[0], out x1, out y1)) return false;
                if (ends.Length == 2)
                { if (!ParseCellAddress(ends[1], out x2, out y2)) return false; }
                else { x2 = x1; y2 = y1; }
                requests.Add(Tuple.Create(activeSheetIndex, Rectangle.FromLTRB(
                    Math.Min(x1, x2), Math.Min(y1, y2), Math.Max(x1, x2) + 1, Math.Max(y1, y2) + 1)));
            }
            if (requests.Any(x => x.Item1 != requests[0].Item1)) return false;
            if (requests[0].Item1 != activeSheetIndex) SwitchSheet(requests[0].Item1);
            EnsureRowCapacity(requests.Max(x => x.Item2.Bottom));
            grid.ClearSelection();
            for (int i = 0; i < requests.Count; i++)
            {
                Rectangle range = requests[i].Item2;
                SelectRectangle(range, range.Left, range.Top, i > 0);
            }
            selectionAnchorColumn = requests[0].Item2.Left;
            selectionAnchorRow = requests[0].Item2.Top;
            return true;
        }

        private void FormulaBarKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                formulaBarChanged = false; UpdateSelection(); grid.Focus();
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Enter)
            {
                if (CommitFormulaBar())
                {
                    grid.Focus();
                    MoveActive(e.Shift ? -1 : 1, 0, false, false);
                }
                e.SuppressKeyPress = true;
            }
        }

        private bool CommitFormulaBar()
        {
            if (!formulaBarChanged || grid.CurrentCell == null) return true;
            if (grid.ReadOnly)
            { formulaBarChanged = false; UpdateSelection(); return false; }
            int row = grid.CurrentCell.RowIndex, column = grid.CurrentCell.ColumnIndex;
            string value = contentBox.Text;
            if (!CanAcceptValue(row, column, value))
            {
                formulaBarChanged = false;
                UpdateSelection();
                status.Text = "Giá trị không có trong danh sách chọn";
                return false;
            }
            formulaBarChanged = false;
            if (!string.Equals(Convert.ToString(grid[column, row].Value) ?? "", value,
                StringComparison.Ordinal)) grid[column, row].Value = value;
            return true;
        }

        private void FinishEditMove()
        {
            pendingEditMove = false;
            MoveActive(editMoveRow, editMoveColumn, false, false);
        }

        private bool HasData(int row, int column)
        {
            return !string.IsNullOrEmpty(Convert.ToString(grid[column, row].Value));
        }

        private int LastUsedRow()
        {
            for (int row = RowCount - 1; row >= 0; row--)
                for (int column = 0; column < ColumnCount; column++)
                    if (HasData(row, column)) return row;
            return 0;
        }

        private int LastUsedColumn()
        {
            for (int column = ColumnCount - 1; column >= 0; column--)
                for (int row = 0; row < RowCount; row++)
                    if (HasData(row, column)) return column;
            return 0;
        }

        private int FindCtrlEdge(int row, int column, int rowStep, int columnStep)
        {
            int limit = rowStep == 0 ? ColumnCount : RowCount;
            int position = rowStep == 0 ? column : row;
            int step = rowStep == 0 ? columnStep : rowStep;
            bool occupied = HasData(row, column);
            int next = position + step;
            if (next < 0 || next >= limit) return position;
            bool nextOccupied = rowStep == 0 ? HasData(row, next) : HasData(next, column);
            if (occupied && nextOccupied)
            {
                while (next >= 0 && next < limit &&
                    (rowStep == 0 ? HasData(row, next) : HasData(next, column)))
                { position = next; next += step; }
                return position;
            }
            while (next >= 0 && next < limit)
            {
                if (rowStep == 0 ? HasData(row, next) : HasData(next, column)) return next;
                next += step;
            }
            return step > 0 ? limit - 1 : 0;
        }

        private void MoveActive(int rowStep, int columnStep, bool control, bool extend)
        {
            if (grid.CurrentCell == null) return;
            int row = grid.CurrentCell.RowIndex, column = grid.CurrentCell.ColumnIndex;
            if (!extend)
            { selectionAnchorRow = row; selectionAnchorColumn = column; }
            else if (grid.SelectedCells.Count <= 1)
            { selectionAnchorRow = row; selectionAnchorColumn = column; }
            int nextRow = row, nextColumn = column;
            if (rowStep != 0)
                nextRow = control ? FindCtrlEdge(row, column, rowStep, 0) :
                    Math.Max(0, Math.Min(RowCount - 1, row + rowStep));
            if (columnStep != 0)
                nextColumn = control ? FindCtrlEdge(row, column, 0, columnStep) :
                    Math.Max(0, Math.Min(ColumnCount - 1, column + columnStep));
            Rectangle range = extend ? Rectangle.FromLTRB(
                Math.Min(selectionAnchorColumn, nextColumn), Math.Min(selectionAnchorRow, nextRow),
                Math.Max(selectionAnchorColumn, nextColumn) + 1, Math.Max(selectionAnchorRow, nextRow) + 1) :
                new Rectangle(nextColumn, nextRow, 1, 1);
            SelectRectangle(range, nextColumn, nextRow, false);
            if (!extend)
            { selectionAnchorRow = nextRow; selectionAnchorColumn = nextColumn; }
        }

        private bool HandleEditingShortcut(Keys keyData)
        {
            if (addressBox.Focused || contentBox.Focused || !grid.ContainsFocus) return false;
            lastInputKeyboard = true;
            floatingActions.Visible = false;
            Keys key = keyData & Keys.KeyCode;
            bool control = (keyData & Keys.Control) != 0;
            bool shift = (keyData & Keys.Shift) != 0;
            bool alt = (keyData & Keys.Alt) != 0;
            if (alt) return false;
            if (grid.ReadOnly && (key == Keys.F2 || control &&
                (key == Keys.X || key == Keys.V || key == Keys.D || key == Keys.R))) return true;
            if (key == Keys.Escape)
            {
                if (grid.IsCurrentCellInEditMode) grid.CancelEdit();
                cutPending = false; fillDragging = false;
                grid.Capture = false; grid.Cursor = Cursors.Default;
                UpdateSelection(); grid.Invalidate(); return true;
            }
            if (grid.IsCurrentCellInEditMode)
            {
                if (key == Keys.Enter || key == Keys.Tab)
                {
                    editMoveRow = key == Keys.Enter ? (shift ? -1 : 1) : 0;
                    editMoveColumn = key == Keys.Tab ? (shift ? -1 : 1) : 0;
                    pendingEditMove = true;
                    if (!grid.EndEdit()) pendingEditMove = false;
                    return true;
                }
                return false;
            }
            if (control && key == Keys.C) { CopySelected(); return true; }
            if (control && key == Keys.X) { CutSelected(); return true; }
            if (control && key == Keys.V) { PasteClipboard(shift ? PasteKind.Values : PasteKind.All); return true; }
            if (control && key == Keys.D) { FillDown(); return true; }
            if (control && key == Keys.R) { FillRight(); return true; }
            if (control && key == Keys.A) { SelectRectangle(new Rectangle(0, 0, ColumnCount, RowCount), 0, 0, false); return true; }
            if (control && !shift && (key == Keys.PageUp || key == Keys.PageDown))
            {
                MoveToAdjacentSheet(key == Keys.PageDown ? 1 : -1);
                return true;
            }
            if ((!control && !shift && key == Keys.F5) ||
                (control && !shift && key == Keys.G))
            {
                addressBox.Focus(); addressBox.SelectAll();
                return true;
            }
            if (!control && shift && key == Keys.F11) { AddSheet(); return true; }
            if (key == Keys.Delete && !control && !shift)
            {
                if (!grid.ReadOnly) ClearSelectedCells();
                return true;
            }
            if (control && key == Keys.Space) { SelectHeader(grid.CurrentCell.ColumnIndex, false); return true; }
            if (!control && shift && key == Keys.Space) { SelectHeader(grid.CurrentCell.RowIndex, true); return true; }
            if (control && (key == Keys.Home || key == Keys.End))
            {
                int row = key == Keys.Home ? 0 : LastUsedRow();
                int column = key == Keys.Home ? 0 : LastUsedColumn();
                if (shift)
                    SelectRectangle(Rectangle.FromLTRB(Math.Min(selectionAnchorColumn, column),
                        Math.Min(selectionAnchorRow, row), Math.Max(selectionAnchorColumn, column) + 1,
                        Math.Max(selectionAnchorRow, row) + 1), column, row, false);
                else
                {
                    SelectRectangle(new Rectangle(column, row, 1, 1), column, row, false);
                    selectionAnchorColumn = column; selectionAnchorRow = row;
                }
                return true;
            }
            if (key == Keys.F2 && !control && !shift)
            {
                grid.BeginEdit(true);
                TextBox editor = grid.EditingControl as TextBox;
                if (editor != null) editor.SelectionStart = editor.TextLength;
                return true;
            }
            int dr = 0, dc = 0;
            if (key == Keys.Up) dr = -1;
            else if (key == Keys.Down) dr = 1;
            else if (key == Keys.Left) dc = -1;
            else if (key == Keys.Right) dc = 1;
            if (dr != 0 || dc != 0) { MoveActive(dr, dc, control, shift); return true; }
            if (!control && (key == Keys.Enter || key == Keys.Tab))
            {
                MoveActive(key == Keys.Enter ? (shift ? -1 : 1) : 0,
                    key == Keys.Tab ? (shift ? -1 : 1) : 0, false, false);
                return true;
            }
            return false;
        }

        private static string ClipboardField(string value)
        {
            if (value.IndexOfAny(new[] { '\t', '\r', '\n', '"' }) < 0) return value;
            return "\"" + value.Replace("\"", "\"\"") + "\"";
        }

        private void CopySelection(bool cut)
        {
            if (grid.CurrentCell == null) return;
            grid.EndEdit();
            Rectangle range = SelectedRectangle(grid, false);
            copiedTop = range.Top; copiedLeft = range.Left; copiedSheetIndex = activeSheetIndex;
            copiedCells = new CellState[range.Height, range.Width];
            copiedDisplays = new string[range.Height, range.Width];
            copiedMask = new bool[range.Height, range.Width];
            Dictionary<int, string> spillDisplays = formulaEngine != null && formulaEngine.HasSpills ?
                formulaEngine.SpillDisplays() : null;
            var text = new StringBuilder();
            for (int row = 0; row < range.Height; row++)
            {
                if (row > 0) text.Append("\r\n");
                for (int column = 0; column < range.Width; column++)
                {
                    if (column > 0) text.Append('\t');
                    DataGridViewCell cell = grid[range.Left + column, range.Top + row];
                    bool selected = cell.Selected || (range.Width == 1 && range.Height == 1);
                    copiedMask[row, column] = selected;
                    if (!selected) continue;
                    copiedCells[row, column] = CaptureCell(range.Left + column, range.Top + row);
                    copiedDisplays[row, column] = Convert.ToString(cell.FormattedValue) ?? "";
                    if (cell.Value == null && spillDisplays != null &&
                        spillDisplays.ContainsKey((range.Top + row) * ColumnCount + range.Left + column))
                        copiedCells[row, column].Value = copiedDisplays[row, column];
                    text.Append(ClipboardField(copiedDisplays[row, column]));
                }
            }
            copiedToken = Guid.NewGuid().ToString("N");
            var clipboardData = new DataObject();
            clipboardData.SetData(DataFormats.UnicodeText, text.ToString());
            clipboardData.SetData("DinkCel.Cells", copiedToken);
            copiedClipboardText = text.ToString();
            try
            {
                Clipboard.SetDataObject(clipboardData, true, 4, 80);
                clipboardFallback = false;
            }
            catch (System.Runtime.InteropServices.ExternalException)
            {
                // Another desktop app can temporarily own the clipboard. Keep DinkCel's internal copy usable.
                clipboardFallback = true;
                clipboardFallbackSequence = GetClipboardSequenceNumber();
            }
            cutPending = cut;
            status.Text = cut ? "Đã cắt vùng chọn; chọn ô đích để dán" : "Đã sao chép vùng chọn";
            grid.Invalidate();
        }

        private void CutSelected()
        {
            if (!grid.ReadOnly) CopySelection(true);
        }
        private void PasteValues() { PasteClipboard(PasteKind.Values); }
        private void PasteFormulas() { PasteClipboard(PasteKind.Formulas); }
        private void PasteFormats() { PasteClipboard(PasteKind.Formats); }
        private void PasteTranspose() { PasteClipboard(PasteKind.Transpose); }

        private static List<List<string>> ParseClipboardText(string text)
        {
            var rows = new List<List<string>>();
            var current = new List<string>();
            var field = new StringBuilder();
            bool quoted = false;
            for (int i = 0; i < text.Length; i++)
            {
                char ch = text[i];
                if (ch == '"')
                {
                    if (quoted && i + 1 < text.Length && text[i + 1] == '"')
                    { field.Append('"'); i++; }
                    else quoted = !quoted;
                }
                else if (!quoted && ch == '\t')
                { current.Add(field.ToString()); field.Clear(); }
                else if (!quoted && (ch == '\n' || ch == '\r'))
                {
                    if (ch == '\r' && i + 1 < text.Length && text[i + 1] == '\n') i++;
                    current.Add(field.ToString()); field.Clear(); rows.Add(current);
                    current = new List<string>();
                }
                else field.Append(ch);
            }
            if (current.Count > 0 || field.Length > 0 || rows.Count == 0)
            { current.Add(field.ToString()); rows.Add(current); }
            return rows;
        }

        private void PasteClipboard(PasteKind kind)
        {
            if (grid.CurrentCell == null || grid.ReadOnly) return;
            bool internalCopy = copiedCells != null && copiedMask != null && clipboardFallback &&
                GetClipboardSequenceNumber() == clipboardFallbackSequence;
            bool hasText = false, hasHtml = false;
            try
            {
                internalCopy |= copiedCells != null && copiedMask != null &&
                    Clipboard.ContainsData("DinkCel.Cells") &&
                    string.Equals(Convert.ToString(Clipboard.GetData("DinkCel.Cells")),
                        copiedToken, StringComparison.Ordinal);
                hasText = Clipboard.ContainsText();
                hasHtml = Clipboard.ContainsText(TextDataFormat.Html);
            }
            catch (System.Runtime.InteropServices.ExternalException)
            { if (!internalCopy) return; }
            if (!internalCopy && !hasText && !hasHtml) return;
            string text = hasText ? Clipboard.GetText() : copiedClipboardText;
            List<List<CellState>> externalHtml = null;
            if (!internalCopy && hasHtml)
                externalHtml = ParseExcelHtml(Clipboard.GetText(TextDataFormat.Html));
            if (externalHtml != null && externalHtml.Count == 0) externalHtml = null;
            if (kind == PasteKind.Formats && !internalCopy && externalHtml == null) return;
            int targetRow = grid.CurrentCell.RowIndex, targetColumn = grid.CurrentCell.ColumnIndex;
            int rows, columns;
            List<List<string>> external = null;
            if (internalCopy)
            { rows = copiedCells.GetLength(0); columns = copiedCells.GetLength(1); }
            else
            {
                external = externalHtml == null ? ParseClipboardText(text) :
                    externalHtml.Select(r => r.Select(c => Convert.ToString(c.Value) ?? "").ToList()).ToList();
                rows = external.Count;
                columns = external.Max(r => r.Count);
            }
            int outputRows = kind == PasteKind.Transpose ? columns : rows;
            int outputColumns = kind == PasteKind.Transpose ? rows : columns;
            if (targetRow + outputRows > MaxRowCount || targetColumn + outputColumns > ColumnCount)
            { status.Text = "Vùng dán vượt quá bảng 50.000 × 26"; return; }
            EnsureRowCapacity(targetRow + outputRows);
            for (int r = 0; r < outputRows; r++)
                for (int c = 0; c < outputColumns; c++)
                {
                    int sr = kind == PasteKind.Transpose ? c : r;
                    int sc = kind == PasteKind.Transpose ? r : c;
                    if (internalCopy && !copiedMask[sr, sc]) continue;
                    if (!internalCopy && sc >= external[sr].Count) continue;
                    if (kind == PasteKind.Formats) continue;
                    string raw = internalCopy ? Convert.ToString(copiedCells[sr, sc].Value) ?? "" : external[sr][sc];
                    string value = kind == PasteKind.Values && internalCopy ? copiedDisplays[sr, sc] :
                        internalCopy && raw.StartsWith("=", StringComparison.Ordinal) && !cutPending ?
                            FormulaEngine.ShiftReferences(raw,
                                targetRow + r - copiedTop - sr, targetColumn + c - copiedLeft - sc,
                                RowCount, ColumnCount) : raw;
                    if (!CanAcceptValue(targetRow + r, targetColumn + c, value))
                    { status.Text = "Vùng dán có giá trị không hợp lệ cho dropdown"; return; }
                }
            bool moveCut = cutPending && internalCopy && kind == PasteKind.All;
            CrossSheetMove crossMove = null;
            if (moveCut && copiedSheetIndex != activeSheetIndex)
            {
                crossMove = new CrossSheetMove { SourceSheetIndex = copiedSheetIndex };
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < columns; c++)
                        if (copiedMask[r, c]) crossMove.OriginalCells.Add(
                            (copiedTop + r) * ColumnCount + copiedLeft + c, copiedCells[r, c]);
            }
            loading = true;
            grid.SuspendLayout();
            try
            {
                if (moveCut && copiedSheetIndex == activeSheetIndex)
                    for (int r = 0; r < rows; r++)
                        for (int c = 0; c < columns; c++)
                            if (copiedMask[r, c])
                            {
                                grid[copiedLeft + c, copiedTop + r].Value = null;
                                grid[copiedLeft + c, copiedTop + r].Style = new DataGridViewCellStyle();
                                grid[copiedLeft + c, copiedTop + r].Tag = null;
                            }
                for (int r = 0; r < outputRows; r++)
                    for (int c = 0; c < outputColumns; c++)
                    {
                        int sr = kind == PasteKind.Transpose ? c : r;
                        int sc = kind == PasteKind.Transpose ? r : c;
                        if (internalCopy && !copiedMask[sr, sc]) continue;
                        if (!internalCopy && sc >= external[sr].Count) continue;
                        DataGridViewCell cell = grid[targetColumn + c, targetRow + r];
                        if (kind != PasteKind.Formats)
                        {
                            string raw = internalCopy ? Convert.ToString(copiedCells[sr, sc].Value) ?? "" : external[sr][sc];
                            cell.Value = kind == PasteKind.Values && internalCopy ? copiedDisplays[sr, sc] :
                                internalCopy && raw.StartsWith("=", StringComparison.Ordinal) && !moveCut ?
                                FormulaEngine.ShiftReferences(raw,
                                    targetRow + r - copiedTop - sr, targetColumn + c - copiedLeft - sc,
                                    RowCount, ColumnCount) : raw;
                        }
                        if ((internalCopy || externalHtml != null) &&
                            (kind == PasteKind.All || kind == PasteKind.Formats || kind == PasteKind.Transpose))
                        {
                            CellState source = internalCopy ? copiedCells[sr, sc] : externalHtml[sr][sc];
                            cell.Style = source.Style == null ? new DataGridViewCellStyle() :
                                new DataGridViewCellStyle(source.Style);
                            cell.Tag = CellExtras.Copy(source.Extras);
                        }
                    }
                if (moveCut && copiedSheetIndex != activeSheetIndex && copiedSheetIndex < sheets.Count)
                {
                    SheetState source = sheets[copiedSheetIndex];
                    for (int r = 0; r < rows; r++)
                        for (int c = 0; c < columns; c++)
                            if (copiedMask[r, c]) source.Cells.Remove((copiedTop + r) * ColumnCount + copiedLeft + c);
                    otherSheetsDirty = true;
                }
            }
            finally { grid.ResumeLayout(); loading = false; }
            if (cutPending) cutPending = false;
            grid.Invalidate();
            Recalculate(); RecordChange(); MarkDirty();
            if (crossMove != null) crossSheetMoves[MoveKey(lastState.RevisionId)] = crossMove;
            SelectRectangle(new Rectangle(targetColumn, targetRow, outputColumns, outputRows),
                targetColumn, targetRow, false);
        }

        private static int PositiveModulo(int value, int divisor)
        { int mod = value % divisor; return mod < 0 ? mod + divisor : mod; }

        private static bool TryNumberOrDate(string text, out double value, out bool date, out string pattern)
        {
            date = false; pattern = null;
            if (double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value)) return true;
            string[] patterns = { "yyyy-MM-dd", "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d-M-yyyy" };
            DateTime parsed;
            foreach (string format in patterns)
                if (DateTime.TryParseExact(text, format, CultureInfo.InvariantCulture,
                    DateTimeStyles.None, out parsed))
                { value = parsed.ToOADate(); date = true; pattern = format; return true; }
            value = 0; return false;
        }

        private static string SeriesText(string first, string second, int offset)
        {
            double a, b; bool dateA, dateB; string formatA, formatB;
            if (!TryNumberOrDate(first, out a, out dateA, out formatA)) return null;
            if (second == null) b = a + 1;
            else if (!TryNumberOrDate(second, out b, out dateB, out formatB) || dateA != dateB) return null;
            double value = a + (b - a) * offset;
            if (dateA)
            {
                try { return DateTime.FromOADate(value).ToString(formatA, CultureInfo.InvariantCulture); }
                catch (ArgumentException) { return null; }
            }
            return value.ToString("G15", CultureInfo.InvariantCulture);
        }

        private void FillDown() { FillEdge(true); }
        private void FillRight() { FillEdge(false); }

        private void FillEdge(bool down)
        {
            Rectangle range = SelectedRectangle(grid, true);
            if (range.IsEmpty || (down ? range.Height : range.Width) < 2) return;
            loading = true;
            try
            {
                for (int row = range.Top; row < range.Bottom; row++)
                    for (int column = range.Left; column < range.Right; column++)
                    {
                        if (down ? row == range.Top : column == range.Left) continue;
                        int sourceRow = down ? range.Top : row;
                        int sourceColumn = down ? column : range.Left;
                        CellState source = CaptureCell(sourceColumn, sourceRow);
                        string raw = Convert.ToString(source.Value) ?? "";
                        string value = FormulaEngine.ShiftReferences(raw, row - sourceRow,
                            column - sourceColumn, RowCount, ColumnCount);
                        if (!CanAcceptValue(row, column, value)) continue;
                        grid[column, row].Value = value;
                        grid[column, row].Style = source.Style == null ? new DataGridViewCellStyle() :
                            new DataGridViewCellStyle(source.Style);
                        grid[column, row].Tag = CellExtras.Copy(source.Extras);
                    }
            }
            finally { loading = false; }
            Recalculate(); RecordChange(); MarkDirty(); UpdateSelection();
        }

        private void FillSeries()
        {
            Rectangle range = SelectedRectangle(grid, true);
            if (range.IsEmpty) return;
            if (range.Width == 1 && range.Height == 1)
            {
                string answer = Prompt("Số ô trong chuỗi", "10");
                int count;
                if (!int.TryParse(answer, out count) || count < 2) return;
                range.Height = Math.Min(count, RowCount - range.Top);
            }
            bool vertical = range.Height >= range.Width;
            loading = true;
            try
            {
                for (int row = range.Top; row < range.Bottom; row++)
                    for (int column = range.Left; column < range.Right; column++)
                    {
                        int index = vertical ? row - range.Top : column - range.Left;
                        if (index == 0) continue;
                        int firstRow = vertical ? range.Top : row;
                        int firstColumn = vertical ? column : range.Left;
                        string first = Convert.ToString(grid[firstColumn, firstRow].Value) ?? "";
                        string second = null;
                        if ((vertical ? range.Height : range.Width) > 1)
                            second = Convert.ToString(grid[vertical ? column : range.Left + 1,
                                vertical ? range.Top + 1 : row].Value);
                        if (string.IsNullOrEmpty(second)) second = null;
                        if (index == 1 && second != null) continue;
                        string value = SeriesText(first, second, index);
                        if (value != null && CanAcceptValue(row, column, value))
                            grid[column, row].Value = value;
                    }
            }
            finally { loading = false; }
            Recalculate(); RecordChange(); MarkDirty();
            SelectRectangle(range, range.Left, range.Top, false);
        }

        private void AutoFillSelection(Rectangle source, int targetRow, int targetColumn)
        {
            if (source.IsEmpty || source.Contains(targetColumn, targetRow)) return;
            bool vertical = targetRow < source.Top || targetRow >= source.Bottom;
            if ((targetColumn < source.Left || targetColumn >= source.Right) &&
                Math.Abs(targetColumn - (source.Right - 1)) > Math.Abs(targetRow - (source.Bottom - 1)))
                vertical = false;
            Rectangle output = vertical ? Rectangle.FromLTRB(source.Left, Math.Min(source.Top, targetRow),
                source.Right, Math.Max(source.Bottom - 1, targetRow) + 1) :
                Rectangle.FromLTRB(Math.Min(source.Left, targetColumn), source.Top,
                    Math.Max(source.Right - 1, targetColumn) + 1, source.Bottom);
            var originals = new CellState[source.Height, source.Width];
            for (int row = 0; row < source.Height; row++)
                for (int column = 0; column < source.Width; column++)
                    originals[row, column] = CaptureCell(source.Left + column, source.Top + row);
            loading = true;
            try
            {
                for (int row = output.Top; row < output.Bottom; row++)
                    for (int column = output.Left; column < output.Right; column++)
                    {
                        if (source.Contains(column, row)) continue;
                        int sourceRow = PositiveModulo(row - source.Top, source.Height);
                        int sourceColumn = PositiveModulo(column - source.Left, source.Width);
                        CellState original = originals[sourceRow, sourceColumn];
                        string raw = Convert.ToString(original.Value) ?? "";
                        string value = FormulaEngine.ShiftReferences(raw,
                            row - source.Top - sourceRow, column - source.Left - sourceColumn,
                            RowCount, ColumnCount);
                        if (!raw.StartsWith("=", StringComparison.Ordinal))
                        {
                            int firstRow = vertical ? 0 : sourceRow;
                            int firstColumn = vertical ? sourceColumn : 0;
                            string first = Convert.ToString(originals[firstRow, firstColumn].Value) ?? "";
                            string second = null;
                            if ((vertical ? source.Height : source.Width) > 1)
                                second = Convert.ToString(originals[vertical ? 1 : sourceRow,
                                    vertical ? sourceColumn : 1].Value);
                            string series = SeriesText(first, second,
                                vertical ? row - source.Top : column - source.Left);
                            if (series != null) value = series;
                        }
                        if (!CanAcceptValue(row, column, value)) continue;
                        DataGridViewCell destination = grid[column, row];
                        destination.Value = value;
                        destination.Style = original.Style == null ? new DataGridViewCellStyle() :
                            new DataGridViewCellStyle(original.Style);
                        destination.Tag = CellExtras.Copy(original.Extras);
                    }
            }
            finally { loading = false; }
            Recalculate(); RecordChange(); MarkDirty();
            SelectRectangle(output, source.Left, source.Top, false);
        }
    }
}
