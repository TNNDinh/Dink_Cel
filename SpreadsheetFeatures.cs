using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DinkCel
{
    internal sealed partial class SpreadsheetForm
    {
        private void AddSpreadsheetMenus()
        {
            var sheet = new ToolStripMenuItem("Trang tính");
            AddMenuItem(sheet, "Thêm trang tính", Keys.None, AddSheet);
            AddMenuItem(sheet, "Đổi tên trang tính", Keys.None, RenameSheet);
            AddMenuItem(sheet, "Xóa trang tính", Keys.None, DeleteSheet);
            AddMenuItem(sheet, "Đặt tên vùng...", Keys.None, DefineNamedRange);
            AddMenuItem(sheet, "Đi tới vùng có tên...", Keys.None, GoToNamedRange);
            AddMenuItem(sheet, "Xóa vùng có tên...", Keys.None, DeleteNamedRange);
            menu.Items.Add(sheet);
            var data = new ToolStripMenuItem("Dữ liệu");
            AddMenuItem(data, "Sắp xếp tăng dần", Keys.None, delegate { SortRows(false); });
            AddMenuItem(data, "Sắp xếp giảm dần", Keys.None, delegate { SortRows(true); });
            AddMenuItem(data, "Lọc theo nội dung...", Keys.None, SetFilter);
            AddMenuItem(data, "Bỏ lọc", Keys.None, delegate { filterColumn = -1; filterValue = ""; ApplyFreezeAndFilter(); RecordChange(); MarkDirty(); });
            AddMenuItem(data, "Tìm và thay thế...", Keys.Control | Keys.H, FindReplace);
            AddMenuItem(data, "Tạo Pivot Table...", Keys.None, CreatePivot);
            AddMenuItem(data, "Làm mới Pivot Table", Keys.None, RefreshAllPivots);
            menu.Items.Add(data);
            var insert = new ToolStripMenuItem("Chèn");
            AddMenuItem(insert, "Tạo Table từ vùng chọn...", Keys.None, CreateTable);
            AddMenuItem(insert, "Biểu đồ từ vùng chọn...", Keys.None, CreateChart);
            AddMenuItem(insert, "Xem biểu đồ...", Keys.None, OpenChart);
            AddMenuItem(insert, "Danh sách chọn cho ô...", Keys.None, AddDropdown);
            menu.Items.Add(insert);
            var cells = new ToolStripMenuItem("Ô");
            AddMenuItem(cells, "Định dạng số...", Keys.None, SetNumberFormat);
            AddMenuItem(cells, "Gộp ô đã chọn", Keys.None, MergeSelection);
            AddMenuItem(cells, "Bỏ gộp ô", Keys.None, UnmergeSelection);
            AddMenuItem(cells, "Tự khớp độ rộng cột", Keys.None, AutoFitColumn);
            AddMenuItem(cells, "Tự khớp chiều cao hàng", Keys.None, AutoFitRow);
            AddMenuItem(cells, "Cố định tại ô đã chọn", Keys.None, FreezeAtCell);
            AddMenuItem(cells, "Bỏ cố định", Keys.None, delegate { freezeRow = freezeColumn = 0; ApplyFreezeAndFilter(); RecordChange(); MarkDirty(); });
            AddMenuItem(cells, "Tô màu có điều kiện...", Keys.None, ConditionalColor);
            AddMenuItem(cells, "Dán chỉ giá trị", Keys.None, PasteValues);
            AddMenuItem(cells, "Dán chỉ định dạng", Keys.None, PasteFormats);
            menu.Items.Add(cells);
            var fileMenu = menu.Items[0] as ToolStripMenuItem;
            if (fileMenu != null)
            {
                fileMenu.DropDownItems.Add(new ToolStripSeparator());
                AddMenuItem(fileMenu, "Xem trước khi in...", Keys.None, PreviewPrint);
                AddMenuItem(fileMenu, "In...", Keys.Control | Keys.P, PrintWorkbook);
                AddMenuItem(fileMenu, "Xuất PDF...", Keys.None, ExportPdf);
            }
        }

        private void SaveActiveSheet()
        {
            CommitFormulaBar();
            if (sheets.Count > activeSheetIndex)
                sheets[activeSheetIndex] = CaptureSheet();
        }

        private void StoreHistory()
        {
            if (activeSheetIndex >= sheetHistories.Count) return;
            SheetHistory history = sheetHistories[activeSheetIndex];
            history.Undo.Clear(); history.Undo.AddRange(undoHistory);
            history.Redo.Clear(); history.Redo.AddRange(redoHistory);
            history.Last = lastState;
            history.NextRevision = nextRevision;
            history.SavedRevision = savedRevision;
        }

        private void LoadHistory()
        {
            while (sheetHistories.Count <= activeSheetIndex) sheetHistories.Add(new SheetHistory());
            SheetHistory history = sheetHistories[activeSheetIndex];
            undoHistory.Clear(); undoHistory.AddRange(history.Undo);
            redoHistory.Clear(); redoHistory.AddRange(history.Redo);
            nextRevision = history.NextRevision;
            savedRevision = history.SavedRevision;
            lastState = history.Last ?? CaptureSheet();
            if (history.Last == null) history.Last = lastState;
        }

        private void MarkAllHistoriesSaved()
        {
            StoreHistory();
            foreach (SheetHistory history in sheetHistories)
                history.SavedRevision = history.NextRevision;
        }

        private void RefreshSheetTabs()
        {
            sheetTabs.Controls.Clear();
            sheetTabs.AllowDrop = true;
            for (int i = 0; i < sheets.Count; i++)
            {
                int index = i;
                if (sheets[i].Hidden) continue;
                var button = new DinkButton { AccentUnderline = index == activeSheetIndex };
                button.Text = sheets[i].Name;
                button.AutoSize = true;
                button.Height = 29;
                button.FlatStyle = FlatStyle.Flat;
                button.FlatAppearance.BorderSize = 0;
                button.Font = index == activeSheetIndex ? DinkDesign.UiBold : DinkDesign.Ui;
                button.Padding = new Padding(8, 1, 8, 1);
                button.Margin = new Padding(2, 2, 2, 1);
                button.BackColor = index == activeSheetIndex ? theme.AccentSoft : theme.Chrome;
                button.ForeColor = index == activeSheetIndex ? theme.Accent :
                    sheets[i].TabColor.IsEmpty ? theme.Text : sheets[i].TabColor;
                button.Tag = index;
                button.Click += delegate { SwitchSheet(index); };
                button.ContextMenuStrip = SheetTabContext(index);
                Point dragOrigin = Point.Empty;
                button.MouseDown += delegate(object sender, MouseEventArgs e) { dragOrigin = e.Location; };
                button.MouseMove += delegate(object sender, MouseEventArgs e)
                {
                    if (e.Button == MouseButtons.Left &&
                        Math.Abs(e.X - dragOrigin.X) + Math.Abs(e.Y - dragOrigin.Y) > 8)
                        button.DoDragDrop(index, DragDropEffects.Move);
                };
                sheetTabs.Controls.Add(button);
            }
            var add = DinkDesign.Button("+", delegate { AddSheet(); });
            add.BackColor = theme.Chrome;
            add.ForeColor = theme.Accent;
            sheetTabs.Controls.Add(add);
            var addContext = new ContextMenuStrip { BackColor = theme.Chrome, ForeColor = theme.Text };
            addContext.Renderer = new ToolStripProfessionalRenderer(new DinkMenuColors(theme));
            addContext.Items.Add("Thêm trang tính", null, delegate { AddSheet(); });
            addContext.Items.Add("Hiện trang ẩn...", null, delegate { UnhideSheet(); });
            add.ContextMenuStrip = addContext;
            sheetTabs.DragEnter -= SheetTabDragEnter;
            sheetTabs.DragDrop -= SheetTabDragDrop;
            sheetTabs.DragEnter += SheetTabDragEnter;
            sheetTabs.DragDrop += SheetTabDragDrop;
        }

        private ContextMenuStrip SheetTabContext(int index)
        {
            var context = new ContextMenuStrip { BackColor = theme.Chrome, ForeColor = theme.Text };
            context.Renderer = new ToolStripProfessionalRenderer(new DinkMenuColors(theme));
            context.Items.Add("Đổi tên", null, delegate { SwitchSheet(index); RenameSheet(); });
            context.Items.Add("Nhân bản", null, delegate { DuplicateSheet(index); });
            context.Items.Add("Sao chép", null, delegate { DuplicateSheet(index); });
            context.Items.Add("Chuyển sang trái", null, delegate { MoveSheet(index, index - 1); });
            context.Items.Add("Chuyển sang phải", null, delegate { MoveSheet(index, index + 1); });
            context.Items.Add("Ẩn", null, delegate { HideSheet(index); });
            context.Items.Add("Hiện trang ẩn...", null, delegate { UnhideSheet(); });
            context.Items.Add("Màu tab...", null, delegate { ChooseSheetTabColor(index); });
            context.Items.Add(new ToolStripSeparator());
            context.Items.Add("Xóa", null, delegate { SwitchSheet(index); DeleteSheet(); });
            return context;
        }

        private void SheetTabDragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.Data.GetDataPresent(typeof(int)) ? DragDropEffects.Move : DragDropEffects.None;
        }

        private void SheetTabDragDrop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(typeof(int))) return;
            int source = (int)e.Data.GetData(typeof(int));
            Point point = sheetTabs.PointToClient(new Point(e.X, e.Y));
            Control target = sheetTabs.GetChildAtPoint(point);
            if (target != null && target.Tag is int) MoveSheet(source, (int)target.Tag);
        }

        private void MoveSheet(int source, int target)
        {
            if (source < 0 || source >= sheets.Count || target < 0 || target >= sheets.Count || source == target) return;
            SaveActiveSheet(); StoreHistory();
            SheetState state = sheets[source];
            SheetHistory history = sheetHistories[source];
            sheets.RemoveAt(source); sheetHistories.RemoveAt(source);
            sheets.Insert(target, state); sheetHistories.Insert(target, history);
            if (activeSheetIndex == source) activeSheetIndex = target;
            else if (source < activeSheetIndex && target >= activeSheetIndex) activeSheetIndex--;
            else if (source > activeSheetIndex && target <= activeSheetIndex) activeSheetIndex++;
            RefreshSheetTabs(); MarkDirty(); otherSheetsDirty = true;
        }

        private void DuplicateSheet(int index)
        {
            SaveActiveSheet(); StoreHistory();
            SheetState source = sheets[index];
            var copy = new SheetState { Name = source.Name + " Copy", Background = source.Background,
                TabColor = source.TabColor,
                ThemeId = source.ThemeId, FreezeRow = source.FreezeRow,
                FreezeColumn = source.FreezeColumn, FilterColumn = source.FilterColumn,
                FilterValue = source.FilterValue };
            int number = 2;
            while (sheets.Any(s => string.Equals(s.Name, copy.Name, StringComparison.OrdinalIgnoreCase)))
                copy.Name = source.Name + " Copy " + number++;
            Array.Copy(source.RowHeights, copy.RowHeights, RowCount);
            Array.Copy(source.ColumnWidths, copy.ColumnWidths, ColumnCount);
            Array.Copy(source.HiddenRows, copy.HiddenRows, RowCount);
            Array.Copy(source.HiddenColumns, copy.HiddenColumns, ColumnCount);
            foreach (var pair in source.Cells)
                copy.Cells.Add(pair.Key, new CellState { Value = pair.Value.Value,
                    Style = pair.Value.Style == null ? null : new DataGridViewCellStyle(pair.Value.Style),
                    Extras = CellExtras.Copy(pair.Value.Extras) });
            copy.Merges.AddRange(source.Merges);
            copy.Rules.AddRange(source.Rules);
            copy.Tables.AddRange(source.Tables);
            copy.Charts.AddRange(source.Charts);
            copy.Validations.AddRange(source.Validations);
            sheets.Insert(index + 1, copy);
            sheetHistories.Insert(index + 1, new SheetHistory());
            if (activeSheetIndex > index) activeSheetIndex++;
            SwitchSheet(index + 1);
            MarkDirty(); otherSheetsDirty = true;
        }

        private void HideSheet(int index)
        {
            if (sheets.Count(s => !s.Hidden) <= 1) return;
            if (index == activeSheetIndex)
            {
                int next = sheets.FindIndex(s => !s.Hidden && s != sheets[index]);
                SwitchSheet(next);
            }
            sheets[index].Hidden = true;
            RefreshSheetTabs(); MarkDirty(); otherSheetsDirty = true;
        }

        private void UnhideSheet()
        {
            List<string> options = sheets.Where(s => s.Hidden).Select(s => s.Name).ToList();
            if (options.Count == 0) return;
            string choice = ChooseOption("Hiện trang tính", options);
            SheetState sheet = sheets.FirstOrDefault(s => s.Hidden && s.Name == choice);
            if (sheet == null) return;
            sheet.Hidden = false;
            RefreshSheetTabs(); MarkDirty(); otherSheetsDirty = true;
        }

        private void ChooseSheetTabColor(int index)
        {
            using (var picker = new ColorDialog { Color = sheets[index].TabColor.IsEmpty ? theme.Accent : sheets[index].TabColor,
                FullOpen = true })
            {
                if (picker.ShowDialog(this) != DialogResult.OK) return;
                sheets[index].TabColor = picker.Color;
                RefreshSheetTabs(); MarkDirty(); otherSheetsDirty = true;
            }
        }

        private void SwitchSheet(int index)
        {
            if (index == activeSheetIndex || index < 0 || index >= sheets.Count) return;
            if (!CommitFormulaBar()) return;
            grid.EndEdit();
            bool unsaved = dirty;
            if (unsaved) otherSheetsDirty = true;
            SaveActiveSheet();
            StoreHistory();
            activeSheetIndex = index;
            RestoreSheet(sheets[index]);
            LoadHistory();
            dirty = unsaved;
            UpdateTitle();
            RefreshSheetTabs();
        }

        private void AddSheet()
        {
            SaveActiveSheet();
            StoreHistory();
            int number = sheets.Count + 1;
            while (sheets.Any(s => string.Equals(s.Name, "Sheet" + number, StringComparison.OrdinalIgnoreCase))) number++;
            var state = new SheetState { Name = "Sheet" + number, Background = sheetBackground, ThemeId = theme.Id };
            for (int r = 0; r < RowCount; r++) state.RowHeights[r] = 27;
            for (int c = 0; c < ColumnCount; c++) state.ColumnWidths[c] = 120;
            sheets.Add(state);
            sheetHistories.Add(new SheetHistory());
            activeSheetIndex = sheets.Count - 1;
            RestoreSheet(state);
            LoadHistory();
            RefreshSheetTabs();
            MarkDirty();
            otherSheetsDirty = true;
        }

        private void RenameSheet()
        {
            string name = Prompt("Tên trang tính", sheets[activeSheetIndex].Name);
            if (name == null) return;
            name = name.Trim();
            if (name.Length == 0 || name.Length > 31 || name.IndexOfAny(new[] { '[', ']', ':', '*', '?', '/', '\\' }) >= 0 ||
                sheets.Any(s => s != sheets[activeSheetIndex] && string.Equals(s.Name, name, StringComparison.OrdinalIgnoreCase)))
            { MessageBox.Show(this, "Tên trang tính không hợp lệ hoặc đã tồn tại."); return; }
            string oldName = sheets[activeSheetIndex].Name;
            SaveActiveSheet();
            sheets[activeSheetIndex].Name = name;
            foreach (NamedRange named in namedRanges)
                if (string.Equals(named.Sheet, oldName, StringComparison.OrdinalIgnoreCase)) named.Sheet = name;
            foreach (PivotDefinition pivot in pivots)
            {
                if (string.Equals(pivot.SourceSheet, oldName, StringComparison.OrdinalIgnoreCase)) pivot.SourceSheet = name;
                if (string.Equals(pivot.TargetSheet, oldName, StringComparison.OrdinalIgnoreCase)) pivot.TargetSheet = name;
            }
            foreach (SheetState sheet in sheets)
                foreach (CellState cell in sheet.Cells.Values)
                    if (cell.Value is string)
                        cell.Value = RenameSheetReferences((string)cell.Value, oldName, name);
            loading = true;
            try
            {
                for (int r = 0; r < RowCount; r++)
                    for (int c = 0; c < ColumnCount; c++)
                    {
                        string raw = Convert.ToString(grid[c, r].Value);
                        if (!string.IsNullOrEmpty(raw) && raw.StartsWith("=", StringComparison.Ordinal))
                            grid[c, r].Value = RenameSheetReferences(raw, oldName, name);
                    }
            }
            finally { loading = false; }
            Recalculate();
            ResetHistory();
            RefreshSheetTabs();
            MarkDirty();
            otherSheetsDirty = true;
        }

        private static string RenameSheetReferences(string formula, string oldName, string newName)
        {
            if (string.IsNullOrEmpty(formula) || !formula.StartsWith("=", StringComparison.Ordinal)) return formula;
            string replacement = Regex.IsMatch(newName, @"^[A-Za-z_][A-Za-z0-9_]*$") ?
                newName : "'" + newName.Replace("'", "''") + "'";
            var pattern = new Regex(@"(?<![A-Za-z0-9_])(?:'(?<quoted>(?:[^']|'')+)'|(?<plain>[A-Za-z_][A-Za-z0-9_]*))!", RegexOptions.IgnoreCase);
            var output = new StringBuilder();
            var code = new StringBuilder();
            bool quotedText = false;
            for (int i = 0; i < formula.Length; i++)
            {
                char ch = formula[i];
                if (ch == '"')
                {
                    if (!quotedText)
                    {
                        output.Append(pattern.Replace(code.ToString(), m =>
                            string.Equals((m.Groups["quoted"].Success ? m.Groups["quoted"].Value.Replace("''", "'") : m.Groups["plain"].Value),
                                oldName, StringComparison.OrdinalIgnoreCase) ? replacement + "!" : m.Value));
                        code.Clear(); quotedText = true;
                    }
                    else if (i + 1 < formula.Length && formula[i + 1] == '"')
                    { output.Append("\"\""); i++; continue; }
                    else quotedText = false;
                    output.Append(ch);
                }
                else if (quotedText) output.Append(ch);
                else code.Append(ch);
            }
            output.Append(pattern.Replace(code.ToString(), m =>
                string.Equals((m.Groups["quoted"].Success ? m.Groups["quoted"].Value.Replace("''", "'") : m.Groups["plain"].Value),
                    oldName, StringComparison.OrdinalIgnoreCase) ? replacement + "!" : m.Value));
            return output.ToString();
        }

        private void DeleteSheet()
        {
            if (sheets.Count == 1 || (!sheets[activeSheetIndex].Hidden &&
                sheets.Count(s => !s.Hidden) <= 1)) return;
            crossSheetMoves.Clear();
            string deletedName = sheets[activeSheetIndex].Name;
            sheets.RemoveAt(activeSheetIndex);
            sheetHistories.RemoveAt(activeSheetIndex);
            namedRanges.RemoveAll(n => string.Equals(n.Sheet, deletedName, StringComparison.OrdinalIgnoreCase));
            pivots.RemoveAll(p => string.Equals(p.SourceSheet, deletedName, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(p.TargetSheet, deletedName, StringComparison.OrdinalIgnoreCase));
            activeSheetIndex = Math.Min(activeSheetIndex, sheets.Count - 1);
            RestoreSheet(sheets[activeSheetIndex]);
            LoadHistory();
            RefreshSheetTabs();
            MarkDirty();
            otherSheetsDirty = true;
        }

        private string Prompt(string title, string initial)
        {
            using (var form = new Form { Text = title, Width = 420, Height = 145, StartPosition = FormStartPosition.CenterParent, FormBorderStyle = FormBorderStyle.FixedDialog, MaximizeBox = false, MinimizeBox = false,
                BackColor = theme.Chrome, ForeColor = theme.Text, Font = DinkDesign.Ui })
            {
                var box = new TextBox { Left = 12, Top = 12, Width = 380, Text = initial,
                    BackColor = theme.Sheet, ForeColor = theme.Text };
                var ok = DinkDesign.Button("OK", delegate { });
                ok.SetBounds(312, 48, 80, 28);
                ok.AutoSize = false; ok.DialogResult = DialogResult.OK;
                ok.BackColor = theme.AccentSoft; ok.ForeColor = theme.Accent;
                form.Controls.Add(box); form.Controls.Add(ok); form.AcceptButton = ok;
                return form.ShowDialog(this) == DialogResult.OK ? box.Text : null;
            }
        }

        private static SheetState StateFromSnapshot(SheetSnapshot source)
        {
            var state = new SheetState { Name = source.Name, Hidden = source.Hidden, TabColor = source.TabColor,
                Background = source.Background, ThemeId = source.ThemeId,
                FreezeRow = source.FreezeRow, FreezeColumn = source.FreezeColumn,
                FilterColumn = source.FilterColumn, FilterValue = source.FilterValue };
            state.Merges.AddRange(source.Merges);
            state.Rules.AddRange(source.Rules);
            state.Tables.AddRange(source.Tables);
            state.Charts.AddRange(source.Charts);
            state.Validations.AddRange(source.Validations);
            for (int r = 0; r < RowCount; r++) state.RowHeights[r] = source.RowHeights.ContainsKey(r) ? source.RowHeights[r] : 27;
            for (int c = 0; c < ColumnCount; c++) state.ColumnWidths[c] = source.ColumnWidths.ContainsKey(c) ? source.ColumnWidths[c] : 120;
            foreach (int row in source.HiddenRows) state.HiddenRows[row] = true;
            foreach (int column in source.HiddenColumns) state.HiddenColumns[column] = true;
            foreach (var pair in source.Cells)
            {
                var cell = pair.Value;
                var style = new DataGridViewCellStyle();
                if (cell.HasFont) style.Font = new Font(cell.FontName, cell.FontSize, cell.FontStyle);
                if (!cell.ForeColor.IsEmpty) style.ForeColor = cell.ForeColor;
                if (!cell.BackColor.IsEmpty) style.BackColor = cell.BackColor;
                style.Alignment = cell.Alignment;
                style.Format = cell.NumberFormat;
                if (cell.Extras != null)
                {
                    style.WrapMode = cell.Extras.Wrap ? DataGridViewTriState.True : DataGridViewTriState.NotSet;
                    style.Padding = new Padding(cell.Extras.Indent * 8, 0, 0, 0);
                }
                state.Cells[pair.Key] = new CellState { Value = cell.Text, Style = style,
                    Extras = CellExtras.Copy(cell.Extras) };
            }
            return state;
        }

        private SheetSnapshot SnapshotFromState(SheetState source)
        {
            var result = new SheetSnapshot { Name = source.Name, Hidden = source.Hidden, TabColor = source.TabColor,
                Background = source.Background, ThemeId = source.ThemeId,
                FreezeRow = source.FreezeRow, FreezeColumn = source.FreezeColumn,
                FilterColumn = source.FilterColumn, FilterValue = source.FilterValue };
            result.Merges.AddRange(source.Merges);
            result.Rules.AddRange(source.Rules);
            result.Tables.AddRange(source.Tables);
            result.Charts.AddRange(source.Charts);
            result.Validations.AddRange(source.Validations);
            for (int r = 0; r < RowCount; r++) if (source.RowHeights[r] != 27) result.RowHeights[r] = source.RowHeights[r];
            for (int c = 0; c < ColumnCount; c++) if (source.ColumnWidths[c] != 120) result.ColumnWidths[c] = source.ColumnWidths[c];
            for (int r = 0; r < RowCount; r++) if (source.HiddenRows[r]) result.HiddenRows.Add(r);
            for (int c = 0; c < ColumnCount; c++) if (source.HiddenColumns[c]) result.HiddenColumns.Add(c);
            foreach (var pair in source.Cells)
            {
                var style = pair.Value.Style;
                result.Cells[pair.Key] = new CellSnapshot { Text = Convert.ToString(pair.Value.Value) ?? "",
                    HasFont = style != null && style.Font != null,
                    FontName = style != null && style.Font != null ? style.Font.Name : "Arial",
                    FontStyle = style != null && style.Font != null ? style.Font.Style : FontStyle.Regular,
                    FontSize = style != null && style.Font != null ? style.Font.Size : 10F,
                    ForeColor = style == null ? Color.Empty : style.ForeColor,
                    BackColor = style == null ? Color.Empty : style.BackColor,
                    Alignment = style == null ? DataGridViewContentAlignment.NotSet : style.Alignment,
                    NumberFormat = style == null ? "" : style.Format,
                    Extras = CellExtras.Copy(pair.Value.Extras) };
            }
            return result;
        }

        private static XElement SerializeSheet(SheetSnapshot sheet)
        {
            var root = new XElement("sheet", new XAttribute("name", sheet.Name), new XAttribute("freezeRow", sheet.FreezeRow),
                new XAttribute("freezeColumn", sheet.FreezeColumn), new XAttribute("filterColumn", sheet.FilterColumn),
                new XAttribute("filterValue", sheet.FilterValue));
            if (!sheet.Background.IsEmpty) root.SetAttributeValue("background", ColorTranslator.ToHtml(sheet.Background));
            if (sheet.Hidden) root.SetAttributeValue("hidden", true);
            if (!sheet.TabColor.IsEmpty) root.SetAttributeValue("tabColor", ColorTranslator.ToHtml(sheet.TabColor));
            if (!string.IsNullOrEmpty(sheet.ThemeId)) root.SetAttributeValue("theme", sheet.ThemeId);
            foreach (var pair in sheet.RowHeights) root.Add(new XElement("row", new XAttribute("index", pair.Key), new XAttribute("height", pair.Value)));
            foreach (var pair in sheet.ColumnWidths) root.Add(new XElement("column", new XAttribute("index", pair.Key), new XAttribute("width", pair.Value)));
            foreach (int row in sheet.HiddenRows) root.Add(new XElement("hiddenRow", new XAttribute("index", row)));
            foreach (int column in sheet.HiddenColumns) root.Add(new XElement("hiddenColumn", new XAttribute("index", column)));
            foreach (var merge in sheet.Merges) root.Add(new XElement("merge", new XAttribute("row", merge.Y), new XAttribute("column", merge.X),
                new XAttribute("width", merge.Width), new XAttribute("height", merge.Height)));
            foreach (var rule in sheet.Rules) root.Add(new XElement("conditional", new XAttribute("row", rule.Range.Y),
                new XAttribute("column", rule.Range.X), new XAttribute("width", rule.Range.Width),
                new XAttribute("height", rule.Range.Height), new XAttribute("threshold", rule.Threshold.ToString(CultureInfo.InvariantCulture)),
                new XAttribute("color", ColorTranslator.ToHtml(rule.Color))));
            SerializeSheetMetadata(root, sheet);
            foreach (var pair in sheet.Cells)
            {
                var cell = pair.Value;
                var element = new XElement("cell", new XAttribute("row", pair.Key / ColumnCount + 1),
                    new XAttribute("column", pair.Key % ColumnCount + 1), cell.Text);
                if (cell.HasFont) { element.SetAttributeValue("fontStyle", (int)cell.FontStyle); element.SetAttributeValue("fontSize", cell.FontSize.ToString(CultureInfo.InvariantCulture)); element.SetAttributeValue("fontName", cell.FontName); }
                if (!cell.ForeColor.IsEmpty) element.SetAttributeValue("fore", ColorTranslator.ToHtml(cell.ForeColor));
                if (!cell.BackColor.IsEmpty) element.SetAttributeValue("back", ColorTranslator.ToHtml(cell.BackColor));
                if (cell.Alignment != DataGridViewContentAlignment.NotSet) element.SetAttributeValue("align", cell.Alignment);
                if (!string.IsNullOrEmpty(cell.NumberFormat)) element.SetAttributeValue("numberFormat", cell.NumberFormat);
                CellExtras.WriteXml(element, cell.Extras);
                root.Add(element);
            }
            return root;
        }

        private bool WriteXlsx(string path)
        {
            try
            {
                grid.EndEdit(); SaveActiveSheet();
                var workbook = new WorkbookSnapshot(); workbook.Sheets.Clear();
                workbook.NamedRanges.AddRange(namedRanges);
                workbook.Pivots.AddRange(pivots);
                foreach (SheetState sheet in sheets) workbook.Sheets.Add(SnapshotFromState(sheet));
                XlsxFile.Write(path, workbook, RowCount, ColumnCount);
                currentPath = path; csvDocument = null; dirty = false;
                otherSheetsDirty = false;
                savedRevision = nextRevision;
                MarkAllHistoriesSaved();
                UpdateTitle();
                status.Text = "Saved " + Path.GetFileName(path);
                return true;
            }
            catch (Exception error) { MessageBox.Show(this, error.Message, "DinkCel", MessageBoxButtons.OK, MessageBoxIcon.Error); return false; }
        }

        private bool WriteXls(string path)
        {
            return WriteInterchange(path, delegate(WorkbookSnapshot workbook)
            { XlsFile.Write(path, workbook, RowCount, ColumnCount); });
        }

        private bool WriteOds(string path)
        {
            return WriteInterchange(path, delegate(WorkbookSnapshot workbook)
            { OdsFile.Write(path, workbook, RowCount, ColumnCount); });
        }

        private bool WriteInterchange(string path, Action<WorkbookSnapshot> writer)
        {
            try
            {
                grid.EndEdit(); SaveActiveSheet();
                var workbook = new WorkbookSnapshot(); workbook.Sheets.Clear();
                workbook.NamedRanges.AddRange(namedRanges);
                workbook.Pivots.AddRange(pivots);
                foreach (SheetState sheet in sheets) workbook.Sheets.Add(SnapshotFromState(sheet));
                writer(workbook);
                currentPath = path; csvDocument = null; dirty = false; otherSheetsDirty = false;
                savedRevision = nextRevision; MarkAllHistoriesSaved(); UpdateTitle();
                status.Text = "Đã lưu " + Path.GetFileName(path);
                return true;
            }
            catch (Exception error)
            {
                MessageBox.Show(this, "Không lưu được tệp: " + error.Message, "DinkCel",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private void SortRows(bool descending)
        {
            if (grid.CurrentCell == null) return;
            int column = grid.CurrentCell.ColumnIndex;
            int first = 0, last = RowCount - 1;
            if (grid.SelectedCells.Count > 1) { first = grid.SelectedCells.Cast<DataGridViewCell>().Min(c => c.RowIndex); last = grid.SelectedCells.Cast<DataGridViewCell>().Max(c => c.RowIndex); }
            while (last > first && Enumerable.Range(0, ColumnCount).All(c => string.IsNullOrEmpty(Convert.ToString(grid[c, last].Value)))) last--;
            if (last <= first) return;
            var rows = new List<CellState[]>();
            for (int r = first; r <= last; r++)
            {
                var values = new CellState[ColumnCount];
                for (int c = 0; c < ColumnCount; c++) values[c] = new CellState { Value = grid[c, r].Value, Style = new DataGridViewCellStyle(grid[c, r].Style) };
                rows.Add(values);
            }
            rows.Sort((a, b) =>
            {
                string x = Convert.ToString(a[column].Value) ?? "", y = Convert.ToString(b[column].Value) ?? "";
                double nx, ny;
                int cmp = double.TryParse(x, out nx) && double.TryParse(y, out ny) ? nx.CompareTo(ny) : StringComparer.CurrentCultureIgnoreCase.Compare(x, y);
                return descending ? -cmp : cmp;
            });
            loading = true;
            for (int r = first; r <= last; r++) for (int c = 0; c < ColumnCount; c++)
            { grid[c, r].Value = rows[r - first][c].Value; grid[c, r].Style = rows[r - first][c].Style; }
            loading = false; Recalculate(); RecordChange(); MarkDirty();
        }

        private void SetFilter()
        {
            if (grid.CurrentCell == null) return;
            string value = Prompt("Hiện các hàng chứa", filterValue);
            if (value == null) return;
            filterColumn = grid.CurrentCell.ColumnIndex; filterValue = value;
            ApplyFreezeAndFilter(); RecordChange(); MarkDirty();
        }

        private void ApplyFreezeAndFilter()
        {
            if (filterColumn >= ColumnCount || filterColumn < -1) filterColumn = -1;
            if (grid.CurrentCell != null && grid.CurrentCell.RowIndex > 0 && filterColumn >= 0 &&
                grid.CurrentCell.RowIndex >= freezeRow &&
                (Convert.ToString(grid[filterColumn, grid.CurrentCell.RowIndex].Value) ?? "").IndexOf(filterValue, StringComparison.CurrentCultureIgnoreCase) < 0)
                grid.CurrentCell = grid[0, 0];
            for (int r = RowCount - 1; r >= 0; r--) grid.Rows[r].Frozen = false;
            for (int c = ColumnCount - 1; c >= 0; c--) grid.Columns[c].Frozen = false;
            for (int r = 0; r < RowCount; r++)
            {
                grid.Rows[r].Visible = !manualHiddenRows[r] && (filterColumn < 0 || r == 0 || r < freezeRow ||
                    (Convert.ToString(grid[filterColumn, r].Value) ?? "").IndexOf(filterValue, StringComparison.CurrentCultureIgnoreCase) >= 0);
            }
            for (int r = 0; r < freezeRow && r < RowCount; r++) grid.Rows[r].Frozen = grid.Rows[r].Visible;
            for (int c = 0; c < freezeColumn && c < ColumnCount; c++) grid.Columns[c].Frozen = true;
        }

        private void FindReplace()
        {
            string find = Prompt("Tìm", ""); if (string.IsNullOrEmpty(find)) return;
            string replacement = Prompt("Thay bằng (Hủy để chỉ tìm)", "");
            int start = grid.CurrentCell == null ? 0 : grid.CurrentCell.RowIndex * ColumnCount + grid.CurrentCell.ColumnIndex + 1;
            int matches = 0;
            if (replacement != null) loading = true;
            try
            {
            for (int offset = 0; offset < RowCount * ColumnCount; offset++)
            {
                int key = (start + offset) % (RowCount * ColumnCount);
                var cell = grid[key % ColumnCount, key / ColumnCount];
                string raw = Convert.ToString(cell.Value) ?? "";
                if (raw.IndexOf(find, StringComparison.CurrentCultureIgnoreCase) < 0) continue;
                matches++;
                if (replacement == null) { grid.CurrentCell = cell; cell.Selected = true; grid.FirstDisplayedScrollingRowIndex = cell.RowIndex; break; }
                cell.Value = ReplaceIgnoreCase(raw, find, replacement);
            }
            }
            finally { if (replacement != null) loading = false; }
            if (replacement != null && matches > 0) { Recalculate(); RecordChange(); MarkDirty(); }
            status.Text = "Tìm thấy " + matches + " ô";
        }

        private static string ReplaceIgnoreCase(string text, string find, string replacement)
        {
            int index = 0;
            while ((index = text.IndexOf(find, index, StringComparison.CurrentCultureIgnoreCase)) >= 0)
            { text = text.Substring(0, index) + replacement + text.Substring(index + find.Length); index += replacement.Length; }
            return text;
        }

        private void SetNumberFormat()
        {
            string format = Prompt("Mã định dạng Excel (ví dụ #,##0.00, 0.0%, dd/MM/yyyy)", "#,##0.00"); if (format == null) return;
            format = NormalizeNumberFormat(format);
            try { FormatNumeric(1234.5, format); }
            catch (FormatException) { MessageBox.Show(this, "Định dạng số không hợp lệ."); return; }
            ApplyToSelection(c => c.Style.Format = format);
        }

        private void MergeSelection()
        {
            if (grid.SelectedCells.Count < 2) return;
            int left = grid.SelectedCells.Cast<DataGridViewCell>().Min(c => c.ColumnIndex);
            int right = grid.SelectedCells.Cast<DataGridViewCell>().Max(c => c.ColumnIndex);
            int top = grid.SelectedCells.Cast<DataGridViewCell>().Min(c => c.RowIndex);
            int bottom = grid.SelectedCells.Cast<DataGridViewCell>().Max(c => c.RowIndex);
            var rect = new Rectangle(left, top, right - left + 1, bottom - top + 1);
            if (merges.Any(m => m.IntersectsWith(rect))) return;
            for (int r = top; r <= bottom; r++)
                for (int c = left; c <= right; c++)
                    if ((r != top || c != left) && !string.IsNullOrEmpty(Convert.ToString(grid[c, r].Value)))
                    { MessageBox.Show(this, "Hãy xóa nội dung ở các ô khác trước khi gộp để tránh mất dữ liệu."); return; }
            merges.Add(rect); UpdateMergedReadOnly(); RecordChange(); MarkDirty(); grid.Invalidate();
        }

        private void UpdateMergedReadOnly()
        {
            for (int r = 0; r < RowCount; r++)
                for (int c = 0; c < ColumnCount; c++)
                    grid[c, r].ReadOnly = merges.Any(m => m.Contains(c, r) && (m.X != c || m.Y != r));
        }

        private void PaintMergedCells(object sender, PaintEventArgs e)
        {
            var drawingState = e.Graphics.Save();
            e.Graphics.SetClip(new Rectangle(grid.RowHeadersWidth, grid.ColumnHeadersHeight,
                Math.Max(0, grid.ClientSize.Width - grid.RowHeadersWidth),
                Math.Max(0, grid.ClientSize.Height - grid.ColumnHeadersHeight)));
            foreach (Rectangle merge in merges)
            {
                if (merge.X < 0 || merge.Y < 0 || merge.Right > ColumnCount || merge.Bottom > RowCount) continue;
                Rectangle first = grid.GetCellDisplayRectangle(merge.X, merge.Y, false);
                Rectangle last = grid.GetCellDisplayRectangle(merge.Right - 1, merge.Bottom - 1, false);
                var bounds = Rectangle.FromLTRB(first.Left, first.Top, last.Right, last.Bottom);
                if (bounds.Width <= 0 || bounds.Height <= 0 || !bounds.IntersectsWith(grid.ClientRectangle)) continue;
                var cell = grid[merge.X, merge.Y];
                Color background = cell.Style.BackColor.IsEmpty ? sheetBackground : cell.Style.BackColor;
                using (var brush = new SolidBrush(background)) e.Graphics.FillRectangle(brush, bounds);
                using (var pen = new Pen(theme.Border)) e.Graphics.DrawRectangle(pen, bounds.Left, bounds.Top, bounds.Width - 1, bounds.Height - 1);
                string value = Convert.ToString(cell.FormattedValue) ?? "";
                Color foreground = cell.Style.ForeColor.IsEmpty ? theme.Text : cell.Style.ForeColor;
                TextRenderer.DrawText(e.Graphics, value, cell.Style.Font ?? grid.Font,
                    Rectangle.Inflate(bounds, -5, -3), foreground,
                    TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);
                if (grid.CurrentCell == cell)
                    using (var pen = new Pen(theme.Accent, 2)) e.Graphics.DrawRectangle(pen, bounds.Left, bounds.Top, bounds.Width - 1, bounds.Height - 1);
            }
            e.Graphics.Restore(drawingState);
        }

        private void UnmergeSelection()
        {
            if (grid.CurrentCell == null) return;
            merges.RemoveAll(m => m.Contains(grid.CurrentCell.ColumnIndex, grid.CurrentCell.RowIndex) ||
                grid.SelectedCells.Cast<DataGridViewCell>().Any(c => m.Contains(c.ColumnIndex, c.RowIndex)));
            UpdateMergedReadOnly();
            RecordChange(); MarkDirty(); grid.Invalidate();
        }

        private void AutoFitColumn()
        {
            if (grid.CurrentCell == null) return;
            foreach (int c in grid.SelectedCells.Cast<DataGridViewCell>().Select(cell => cell.ColumnIndex).Distinct())
            {
                int width = 65;
                for (int r = 0; r < RowCount; r++)
                {
                    DataGridViewCell cell = grid[c, r];
                    Font font = cell.InheritedStyle.Font ?? grid.Font;
                    width = Math.Max(width, TextRenderer.MeasureText(Convert.ToString(cell.FormattedValue) ?? "", font).Width +
                        cell.InheritedStyle.Padding.Left + 18);
                }
                grid.Columns[c].Width = Math.Min(800, width);
            }
        }

        private void AutoFitRow()
        {
            if (grid.CurrentCell == null) return;
            foreach (int r in grid.SelectedCells.Cast<DataGridViewCell>().Select(cell => cell.RowIndex).Distinct())
            {
                int height = 27;
                for (int c = 0; c < ColumnCount; c++)
                {
                    DataGridViewCell cell = grid[c, r];
                    Font font = cell.InheritedStyle.Font ?? grid.Font;
                    TextFormatFlags flags = cell.InheritedStyle.WrapMode == DataGridViewTriState.True ?
                        TextFormatFlags.WordBreak : TextFormatFlags.SingleLine;
                    height = Math.Max(height, TextRenderer.MeasureText(Convert.ToString(cell.FormattedValue) ?? "",
                        font, new Size(Math.Max(20, grid.Columns[c].Width - 8), 1000), flags).Height + 8);
                }
                grid.Rows[r].Height = Math.Min(400, height);
            }
        }

        private void FreezeAtCell()
        {
            if (grid.CurrentCell == null) return;
            freezeRow = grid.CurrentCell.RowIndex; freezeColumn = grid.CurrentCell.ColumnIndex;
            ApplyFreezeAndFilter(); RecordChange(); MarkDirty();
        }

        private void ConditionalColor()
        {
            string threshold = Prompt("Tô màu giá trị lớn hơn", "0");
            double limit; if (threshold == null || !double.TryParse(threshold, out limit)) return;
            using (var picker = new ColorDialog { Color = Color.LightGreen })
            {
                if (picker.ShowDialog(this) != DialogResult.OK) return;
                int left = grid.SelectedCells.Count == 0 ? grid.CurrentCell.ColumnIndex : grid.SelectedCells.Cast<DataGridViewCell>().Min(c => c.ColumnIndex);
                int right = grid.SelectedCells.Count == 0 ? left : grid.SelectedCells.Cast<DataGridViewCell>().Max(c => c.ColumnIndex);
                int top = grid.SelectedCells.Count == 0 ? grid.CurrentCell.RowIndex : grid.SelectedCells.Cast<DataGridViewCell>().Min(c => c.RowIndex);
                int bottom = grid.SelectedCells.Count == 0 ? top : grid.SelectedCells.Cast<DataGridViewCell>().Max(c => c.RowIndex);
                conditionalRules.Add(new ConditionalRule { Range = new Rectangle(left, top, right - left + 1, bottom - top + 1), Threshold = limit, Color = picker.Color });
                grid.Invalidate(); RecordChange(); MarkDirty();
            }
        }

        private void PasteSpecial(bool formatOnly)
        {
            PasteClipboard(formatOnly ? PasteKind.Formats : PasteKind.Values);
        }
    }
}
