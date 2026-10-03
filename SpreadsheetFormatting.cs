using System;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Collections.Generic;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DinkCel
{
    internal sealed class BorderEdge
    {
        public string Style = "";
        public Color Color = Color.Empty;
        public BorderEdge Copy() { return new BorderEdge { Style = Style, Color = Color }; }
        public bool Exists { get { return !string.IsNullOrEmpty(Style) && Style != "none"; } }
    }

    internal sealed class CellExtras
    {
        public BorderEdge Left = new BorderEdge();
        public BorderEdge Right = new BorderEdge();
        public BorderEdge Top = new BorderEdge();
        public BorderEdge Bottom = new BorderEdge();
        public bool Wrap;
        public bool Shrink;
        public int Indent;
        public int Rotation;

        public static CellExtras Copy(CellExtras source)
        {
            return source == null ? null : new CellExtras { Left = source.Left.Copy(), Right = source.Right.Copy(),
                Top = source.Top.Copy(), Bottom = source.Bottom.Copy(), Wrap = source.Wrap,
                Shrink = source.Shrink, Indent = source.Indent, Rotation = source.Rotation };
        }

        private static void WriteEdge(XElement element, string name, BorderEdge edge)
        {
            if (!edge.Exists) return;
            element.SetAttributeValue(name, edge.Style);
            if (!edge.Color.IsEmpty) element.SetAttributeValue(name + "Color", ColorTranslator.ToHtml(edge.Color));
        }

        public static void WriteXml(XElement element, CellExtras extras)
        {
            if (extras == null) return;
            WriteEdge(element, "borderLeft", extras.Left);
            WriteEdge(element, "borderRight", extras.Right);
            WriteEdge(element, "borderTop", extras.Top);
            WriteEdge(element, "borderBottom", extras.Bottom);
            if (extras.Wrap) element.SetAttributeValue("wrap", true);
            if (extras.Shrink) element.SetAttributeValue("shrink", true);
            if (extras.Indent != 0) element.SetAttributeValue("indent", extras.Indent);
            if (extras.Rotation != 0) element.SetAttributeValue("rotation", extras.Rotation);
        }

        private static BorderEdge ReadEdge(XElement element, string name)
        {
            string style = (string)element.Attribute(name) ?? "";
            string color = (string)element.Attribute(name + "Color");
            return new BorderEdge { Style = style, Color = color == null ? Color.Empty : ColorTranslator.FromHtml(color) };
        }

        public static CellExtras ReadXml(XElement element)
        {
            var result = new CellExtras { Left = ReadEdge(element, "borderLeft"),
                Right = ReadEdge(element, "borderRight"), Top = ReadEdge(element, "borderTop"),
                Bottom = ReadEdge(element, "borderBottom"), Wrap = (bool?)element.Attribute("wrap") ?? false,
                Shrink = (bool?)element.Attribute("shrink") ?? false,
                Indent = (int?)element.Attribute("indent") ?? 0,
                Rotation = (int?)element.Attribute("rotation") ?? 0 };
            return result.Left.Exists || result.Right.Exists || result.Top.Exists || result.Bottom.Exists ||
                result.Wrap || result.Shrink || result.Indent != 0 || result.Rotation != 0 ? result : null;
        }
    }

    internal sealed partial class SpreadsheetForm
    {
        private readonly ToolStripComboBox fontCombo = new ToolStripComboBox();
        private CellState painterSource;
        private readonly bool[] manualHiddenRows = new bool[RowCount];
        private readonly bool[] manualHiddenColumns = new bool[ColumnCount];

        private void AddFormattingMenus(ToolStripMenuItem formatMenu)
        {
            AddMenuItem(formatMenu, "Gạch ngang", Keys.None, delegate { ToggleFontStyle(FontStyle.Strikeout); });
            AddMenuItem(formatMenu, "Chọn phông chữ...", Keys.None, ChooseFontFamily);
            AddMenuItem(formatMenu, "Sao chép định dạng", Keys.None, StartFormatPainter);
            AddMenuItem(formatMenu, "Dán định dạng đã sao chép", Keys.None, ApplyFormatPainter);
            AddMenuItem(formatMenu, "Xóa định dạng", Keys.None, ClearFormats);
            AddMenuItem(formatMenu, "Xóa toàn bộ", Keys.None, ClearAll);
            var border = new ToolStripMenuItem("Viền ô");
            foreach (string edge in new[] { "Tất cả", "Ngoài", "Trái", "Phải", "Trên", "Dưới", "Không viền" })
            {
                string choice = edge;
                border.DropDownItems.Add(edge, null, delegate { SetBorders(choice); });
            }
            formatMenu.DropDownItems.Add(border);
            AddMenuItem(formatMenu, "Kiểu viền...", Keys.None, SetBorderStyle);
            var vertical = new ToolStripMenuItem("Căn dọc");
            vertical.DropDownItems.Add("Trên", null, delegate { SetVerticalAlignment(0); });
            vertical.DropDownItems.Add("Giữa", null, delegate { SetVerticalAlignment(1); });
            vertical.DropDownItems.Add("Dưới", null, delegate { SetVerticalAlignment(2); });
            formatMenu.DropDownItems.Add(vertical);
            AddMenuItem(formatMenu, "Xuống dòng trong ô", Keys.None, ToggleWrap);
            AddMenuItem(formatMenu, "Thu chữ vừa ô", Keys.None, ToggleShrink);
            AddMenuItem(formatMenu, "Thụt lề...", Keys.None, SetIndent);
            AddMenuItem(formatMenu, "Xoay chữ...", Keys.None, SetRotation);
            var formats = new ToolStripMenuItem("Định dạng số");
            foreach (var pair in new[] { "General", "Number", "Currency", "Accounting", "Percentage", "Date", "Time", "Scientific", "Fraction", "Custom..." })
            {
                string choice = pair;
                formats.DropDownItems.Add(pair, null, delegate { ApplyNumberFormat(choice); });
            }
            formatMenu.DropDownItems.Add(formats);
            var dimensions = new ToolStripMenuItem("Hàng và cột");
            dimensions.DropDownItems.Add("Ẩn hàng", null, delegate { SetHidden(true, true); });
            dimensions.DropDownItems.Add("Hiện hàng", null, delegate { SetHidden(true, false); });
            dimensions.DropDownItems.Add("Ẩn cột", null, delegate { SetHidden(false, true); });
            dimensions.DropDownItems.Add("Hiện cột", null, delegate { SetHidden(false, false); });
            dimensions.DropDownItems.Add("Chiều cao hàng...", null, delegate { SetDimension(true); });
            dimensions.DropDownItems.Add("Độ rộng cột...", null, delegate { SetDimension(false); });
            dimensions.DropDownItems.Add("Tự khớp chiều cao", null, delegate { AutoFitRow(); });
            dimensions.DropDownItems.Add("Tự khớp độ rộng", null, delegate { AutoFitColumn(); });
            formatMenu.DropDownItems.Add(dimensions);
        }

        private void AddFontPicker()
        {
            toolbar.Items.Add(new ToolStripLabel("Phông"));
            fontCombo.AutoSize = false;
            fontCombo.Width = 130;
            fontCombo.DropDownStyle = ComboBoxStyle.DropDownList;
            foreach (FontFamily font in FontFamily.Families) fontCombo.Items.Add(font.Name);
            fontCombo.SelectedIndexChanged += delegate
            {
                if (!syncingToolbar && fontCombo.SelectedItem != null)
                    ApplyFontFamily(fontCombo.SelectedItem.ToString());
            };
            toolbar.Items.Add(fontCombo);
        }

        private void ChooseFontFamily()
        {
            using (var dialog = new FontDialog { MinSize = 6, MaxSize = 72 })
            {
                dialog.Font = grid.CurrentCell == null ? grid.Font : grid.CurrentCell.InheritedStyle.Font ?? grid.Font;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                    ApplyToSelection(delegate(DataGridViewCell cell)
                    { cell.Style.Font = new Font(dialog.Font.FontFamily, dialog.Font.Size, dialog.Font.Style); });
            }
        }

        private void ApplyFontFamily(string family)
        {
            ApplyToSelection(delegate(DataGridViewCell cell)
            {
                Font font = cell.InheritedStyle.Font ?? grid.Font;
                cell.Style.Font = new Font(family, font.Size, font.Style);
            });
        }

        private void StartFormatPainter()
        {
            if (grid.CurrentCell == null) return;
            painterSource = CaptureCell(grid.CurrentCell.ColumnIndex, grid.CurrentCell.RowIndex);
            status.Text = "Đã lấy định dạng. Chọn ô đích rồi dùng Dán định dạng đã sao chép.";
        }

        private void ApplyFormatPainter()
        {
            if (painterSource == null) return;
            ApplyToSelection(delegate(DataGridViewCell cell)
            {
                cell.Style = painterSource.Style == null ? new DataGridViewCellStyle() :
                    new DataGridViewCellStyle(painterSource.Style);
                cell.Tag = CellExtras.Copy(painterSource.Extras);
            });
            painterSource = null;
        }

        private void ClearFormats()
        {
            ApplyToSelection(delegate(DataGridViewCell cell)
            { cell.Style = new DataGridViewCellStyle(); cell.Tag = null; });
        }

        private void ClearAll()
        {
            loading = true;
            try
            {
                ApplyToSelection(delegate(DataGridViewCell cell)
                { cell.Value = null; cell.Style = new DataGridViewCellStyle(); cell.Tag = null; });
            }
            finally { loading = false; }
            Recalculate();
        }

        private static CellExtras Extras(DataGridViewCell cell)
        {
            var extras = cell.Tag as CellExtras;
            if (extras == null) cell.Tag = extras = new CellExtras();
            return extras;
        }

        private static string borderStyle = "thin";
        private static Color borderColor = Color.FromArgb(80, 90, 105);

        private void SetBorderStyle()
        {
            string choice = Prompt("Kiểu viền: thin, medium, thick, dashed, dotted, double", borderStyle);
            if (choice == null) return;
            choice = choice.Trim().ToLowerInvariant();
            if (!new[] { "thin", "medium", "thick", "dashed", "dotted", "double" }.Contains(choice))
            { MessageBox.Show(this, "Kiểu viền không hợp lệ."); return; }
            borderStyle = choice;
            using (var dialog = new ColorDialog { Color = borderColor })
                if (dialog.ShowDialog(this) == DialogResult.OK) borderColor = dialog.Color;
        }

        private void SetBorders(string choice)
        {
            var selected = grid.SelectedCells.Cast<DataGridViewCell>()
                .Select(cell => cell.RowIndex * ColumnCount + cell.ColumnIndex).ToArray();
            ApplyToSelection(delegate(DataGridViewCell cell)
            {
                CellExtras extras = Extras(cell);
                bool all = choice == "Tất cả";
                bool outer = choice == "Ngoài";
                bool clear = choice == "Không viền";
                int row = cell.RowIndex, column = cell.ColumnIndex;
                if (all || clear || choice == "Trái" || outer && !selected.Contains(row * ColumnCount + column - 1)) extras.Left = NewEdge(clear);
                if (all || clear || choice == "Phải" || outer && !selected.Contains(row * ColumnCount + column + 1)) extras.Right = NewEdge(clear);
                if (all || clear || choice == "Trên" || outer && !selected.Contains((row - 1) * ColumnCount + column)) extras.Top = NewEdge(clear);
                if (all || clear || choice == "Dưới" || outer && !selected.Contains((row + 1) * ColumnCount + column)) extras.Bottom = NewEdge(clear);
            });
        }

        private static BorderEdge NewEdge(bool clear)
        { return new BorderEdge { Style = clear ? "" : borderStyle, Color = borderColor }; }

        private void SetVerticalAlignment(int vertical)
        {
            ApplyToSelection(delegate(DataGridViewCell cell)
            {
                DataGridViewContentAlignment old = cell.InheritedStyle.Alignment;
                bool right = old == DataGridViewContentAlignment.TopRight ||
                    old == DataGridViewContentAlignment.MiddleRight || old == DataGridViewContentAlignment.BottomRight;
                bool center = old == DataGridViewContentAlignment.TopCenter ||
                    old == DataGridViewContentAlignment.MiddleCenter || old == DataGridViewContentAlignment.BottomCenter;
                cell.Style.Alignment = vertical == 0 ? (right ? DataGridViewContentAlignment.TopRight :
                    center ? DataGridViewContentAlignment.TopCenter : DataGridViewContentAlignment.TopLeft) :
                    vertical == 1 ? (right ? DataGridViewContentAlignment.MiddleRight :
                    center ? DataGridViewContentAlignment.MiddleCenter : DataGridViewContentAlignment.MiddleLeft) :
                    right ? DataGridViewContentAlignment.BottomRight : center ?
                    DataGridViewContentAlignment.BottomCenter : DataGridViewContentAlignment.BottomLeft;
            });
        }

        private void ToggleWrap()
        {
            bool enable = grid.CurrentCell != null && !Extras(grid.CurrentCell).Wrap;
            ApplyToSelection(delegate(DataGridViewCell cell)
            { Extras(cell).Wrap = enable; cell.Style.WrapMode = enable ? DataGridViewTriState.True : DataGridViewTriState.False; });
        }

        private void ToggleShrink()
        {
            bool enable = grid.CurrentCell != null && !Extras(grid.CurrentCell).Shrink;
            ApplyToSelection(delegate(DataGridViewCell cell) { Extras(cell).Shrink = enable; });
        }

        private void SetIndent()
        {
            string input = Prompt("Thụt lề (0–15)", "1");
            int value;
            if (input == null || !int.TryParse(input, out value) || value < 0 || value > 15) return;
            ApplyToSelection(delegate(DataGridViewCell cell)
            { Extras(cell).Indent = value; cell.Style.Padding = new Padding(value * 8, 0, 0, 0); });
        }

        private void SetRotation()
        {
            string input = Prompt("Góc xoay chữ (-90 đến 90)", "45");
            int value;
            if (input == null || !int.TryParse(input, out value) || value < -90 || value > 90) return;
            ApplyToSelection(delegate(DataGridViewCell cell) { Extras(cell).Rotation = value; });
        }

        private void ApplyNumberFormat(string name)
        {
            if (name == "Custom...") { SetNumberFormat(); return; }
            string format = name == "General" ? "" : name == "Number" ? "#,##0.00" :
                name == "Currency" ? "#,##0.00 ₫" : name == "Accounting" ? "#,##0.00;(#,##0.00);–" :
                name == "Percentage" ? "0.00%" : name == "Date" ? "dd/MM/yyyy" :
                name == "Time" ? "HH:mm:ss" : name == "Scientific" ? "0.00E+00" : "# ?/?";
            ApplyToSelection(delegate(DataGridViewCell cell) { cell.Style.Format = format; });
        }

        private static string FormatNumeric(double number, string format)
        {
            if (format == "General") return number.ToString("0.##########", CultureInfo.InvariantCulture);
            if (Regex.IsMatch(format, @"^[NPC]\d+$", RegexOptions.IgnoreCase))
                return number.ToString(format, CultureInfo.CurrentCulture);
            string clean = Regex.Replace(format, @"\[(?:Red|Blue|Green|Black|White|Yellow|Color\d+|\$[^\]]+)\]", "",
                RegexOptions.IgnoreCase);
            string probe = Regex.Replace(clean, "\"[^\"]*\"", "");
            probe = Regex.Replace(probe, @"\\.", "");
            if (Regex.IsMatch(probe, @"[ydhs]", RegexOptions.IgnoreCase) ||
                Regex.IsMatch(probe, @"m[/\-]d|d[/\-]m", RegexOptions.IgnoreCase))
            {
                try { return DateTime.FromOADate(number).ToString(clean, CultureInfo.CurrentCulture); }
                catch (ArgumentException) { return "#NUM!"; }
            }
            if (clean.Contains("?/"))
            {
                int whole = (int)Math.Truncate(number);
                double fraction = Math.Abs(number - whole);
                int limit = clean.Contains("??/??") ? 99 : 9;
                int numerator = 0, denominator = 1;
                double error = double.MaxValue;
                for (int d = 1; d <= limit; d++)
                {
                    int n = (int)Math.Round(fraction * d);
                    double delta = Math.Abs(fraction - (double)n / d);
                    if (delta < error) { error = delta; numerator = n; denominator = d; }
                }
                if (numerator == 0) return whole.ToString(CultureInfo.CurrentCulture);
                if (numerator == denominator) return (whole + Math.Sign(number)).ToString(CultureInfo.CurrentCulture);
                return (whole == 0 ? (number < 0 ? "-" : "") : whole.ToString(CultureInfo.CurrentCulture) + " ") +
                    numerator + "/" + denominator;
            }
            return number.ToString(clean, CultureInfo.CurrentCulture);
        }

        private static string NormalizeNumberFormat(string format)
        {
            Match legacy = Regex.Match(format.Trim(), @"^(?<kind>[NPC])(?<digits>\d+)$", RegexOptions.IgnoreCase);
            if (!legacy.Success) return format;
            int digits = Math.Min(10, int.Parse(legacy.Groups["digits"].Value));
            string decimals = digits == 0 ? "" : "." + new string('0', digits);
            string kind = legacy.Groups["kind"].Value.ToUpperInvariant();
            return kind == "P" ? "0" + decimals + "%" :
                kind == "C" ? "#,##0" + decimals + " ₫" : "#,##0" + decimals;
        }

        private static string HtmlAttribute(string attributes, string name)
        {
            Match match = Regex.Match(attributes, @"(?:^|\s)" + Regex.Escape(name) +
                @"\s*=\s*(?:""(?<v>[^""]*)""|'(?<v>[^']*)'|(?<v>[^\s>]+))", RegexOptions.IgnoreCase);
            return match.Success ? WebUtility.HtmlDecode(match.Groups["v"].Value) : "";
        }

        private static Color CssColor(string text)
        {
            try
            {
                text = text.Trim();
                if (text.StartsWith("rgb(", StringComparison.OrdinalIgnoreCase))
                {
                    string[] parts = text.Substring(4).TrimEnd(')').Split(',');
                    return Color.FromArgb(int.Parse(parts[0]), int.Parse(parts[1]), int.Parse(parts[2]));
                }
                return text.Length == 0 ? Color.Empty : ColorTranslator.FromHtml(text);
            }
            catch (Exception) { return Color.Empty; }
        }

        private static void ApplyCssBorder(BorderEdge edge, string css)
        {
            if (css.IndexOf("none", StringComparison.OrdinalIgnoreCase) >= 0) return;
            edge.Style = css.IndexOf("double", StringComparison.OrdinalIgnoreCase) >= 0 ? "double" :
                css.IndexOf("dashed", StringComparison.OrdinalIgnoreCase) >= 0 ? "dashed" :
                css.IndexOf("dotted", StringComparison.OrdinalIgnoreCase) >= 0 ? "dotted" :
                css.IndexOf("thick", StringComparison.OrdinalIgnoreCase) >= 0 ||
                css.IndexOf("3px", StringComparison.OrdinalIgnoreCase) >= 0 ? "thick" : "thin";
            Match color = Regex.Match(css, @"#[0-9a-fA-F]{3,8}|rgb\([^)]+\)|\b(?:black|white|red|blue|green|gray|grey)\b",
                RegexOptions.IgnoreCase);
            edge.Color = color.Success ? CssColor(color.Value) : Color.Black;
        }

        private static CellState ParseHtmlStyle(string styleText, string value)
        {
            var style = new DataGridViewCellStyle();
            var extras = new CellExtras();
            string family = "Arial";
            float size = 10F;
            FontStyle flags = FontStyle.Regular;
            bool hasFont = false;
            string horizontal = null, vertical = null;
            foreach (string declaration in styleText.Split(';'))
            {
                int colon = declaration.IndexOf(':');
                if (colon < 0) continue;
                string key = declaration.Substring(0, colon).Trim().ToLowerInvariant();
                string css = declaration.Substring(colon + 1).Trim().Trim('"', '\'');
                if (key == "font-family") { family = css.Split(',')[0].Trim().Trim('"', '\''); hasFont = true; }
                else if (key == "font-size")
                {
                    Match number = Regex.Match(css, @"[0-9]+(?:\.[0-9]+)?");
                    float parsed;
                    if (number.Success && float.TryParse(number.Value, NumberStyles.Float,
                        CultureInfo.InvariantCulture, out parsed)) { size = parsed; hasFont = true; }
                }
                else if (key == "font-weight" && (css == "bold" || css == "700"))
                { flags |= FontStyle.Bold; hasFont = true; }
                else if (key == "font-style" && css == "italic")
                { flags |= FontStyle.Italic; hasFont = true; }
                else if (key == "text-decoration")
                {
                    if (css.Contains("underline")) flags |= FontStyle.Underline;
                    if (css.Contains("line-through")) flags |= FontStyle.Strikeout;
                    hasFont = true;
                }
                else if (key == "color") style.ForeColor = CssColor(css);
                else if (key == "background" || key == "background-color") style.BackColor = CssColor(css);
                else if (key == "text-align") horizontal = css;
                else if (key == "vertical-align") vertical = css;
                else if (key == "white-space" || key == "mso-wrap-style")
                    extras.Wrap = css == "normal" || css == "wrap";
                else if (key == "mso-number-format") style.Format = css.Replace("\\", "");
                else if (key == "mso-rotate")
                { int rotation; if (int.TryParse(css, out rotation)) extras.Rotation = Math.Max(-90, Math.Min(90, rotation)); }
                else if (key == "padding-left")
                { Match number = Regex.Match(css, @"\d+"); if (number.Success) extras.Indent = Math.Min(15, int.Parse(number.Value) / 8); }
                else if (key == "border")
                { ApplyCssBorder(extras.Left, css); ApplyCssBorder(extras.Right, css);
                    ApplyCssBorder(extras.Top, css); ApplyCssBorder(extras.Bottom, css); }
                else if (key == "border-left") ApplyCssBorder(extras.Left, css);
                else if (key == "border-right") ApplyCssBorder(extras.Right, css);
                else if (key == "border-top") ApplyCssBorder(extras.Top, css);
                else if (key == "border-bottom") ApplyCssBorder(extras.Bottom, css);
            }
            if (hasFont)
            {
                try { style.Font = new Font(family, Math.Max(6, Math.Min(72, size)), flags); }
                catch (ArgumentException) { style.Font = new Font("Arial", 10F, flags); }
            }
            if (horizontal != null || vertical != null)
                style.Alignment = vertical == "top" ?
                    horizontal == "center" ? DataGridViewContentAlignment.TopCenter :
                    horizontal == "right" ? DataGridViewContentAlignment.TopRight : DataGridViewContentAlignment.TopLeft :
                    vertical == "bottom" ?
                    horizontal == "center" ? DataGridViewContentAlignment.BottomCenter :
                    horizontal == "right" ? DataGridViewContentAlignment.BottomRight : DataGridViewContentAlignment.BottomLeft :
                    horizontal == "center" ? DataGridViewContentAlignment.MiddleCenter :
                    horizontal == "right" ? DataGridViewContentAlignment.MiddleRight : DataGridViewContentAlignment.MiddleLeft;
            if (extras.Wrap) style.WrapMode = DataGridViewTriState.True;
            if (extras.Indent > 0) style.Padding = new Padding(extras.Indent * 8, 0, 0, 0);
            return new CellState { Value = value, Style = style, Extras = extras };
        }

        internal static List<List<CellState>> ParseExcelHtml(string html)
        {
            var rows = new List<List<CellState>>();
            if (string.IsNullOrEmpty(html)) return rows;
            var classes = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (Match css in Regex.Matches(html, @"\.(?<name>[\w-]+)\s*\{(?<style>[^}]*)\}",
                RegexOptions.IgnoreCase | RegexOptions.Singleline))
                classes[css.Groups["name"].Value] = css.Groups["style"].Value;
            Match table = Regex.Match(html, @"<table\b[^>]*>(?<body>.*?)</table>",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (!table.Success) return rows;
            foreach (Match row in Regex.Matches(table.Groups["body"].Value, @"<tr\b[^>]*>(?<body>.*?)</tr>",
                RegexOptions.IgnoreCase | RegexOptions.Singleline))
            {
                var cells = new List<CellState>();
                foreach (Match cell in Regex.Matches(row.Groups["body"].Value,
                    @"<(?:td|th)\b(?<attrs>[^>]*)>(?<body>.*?)</(?:td|th)>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline))
                {
                    string attrs = cell.Groups["attrs"].Value;
                    string classStyle;
                    string className = HtmlAttribute(attrs, "class").Split(' ')[0];
                    classes.TryGetValue(className, out classStyle);
                    string css = (classStyle ?? "") + ";" + HtmlAttribute(attrs, "style");
                    string body = Regex.Replace(cell.Groups["body"].Value, @"<br\s*/?>", "\n", RegexOptions.IgnoreCase);
                    string value = WebUtility.HtmlDecode(Regex.Replace(body, "<[^>]*>", "")).Replace('\u00a0', ' ');
                    CellState parsed = ParseHtmlStyle(css, value);
                    string bgcolor = HtmlAttribute(attrs, "bgcolor");
                    if (bgcolor.Length > 0) parsed.Style.BackColor = CssColor(bgcolor);
                    cells.Add(parsed);
                }
                if (cells.Count > 0) rows.Add(cells);
            }
            return rows;
        }

        private void SetDimension(bool row)
        {
            string input = Prompt(row ? "Chiều cao hàng (8–400 px)" : "Độ rộng cột (20–800 px)", row ? "27" : "120");
            int value;
            if (input == null || !int.TryParse(input, out value) || value < (row ? 8 : 20) || value > (row ? 400 : 800)) return;
            if (grid.CurrentCell == null) return;
            if (row) grid.Rows[grid.CurrentCell.RowIndex].Height = value;
            else grid.Columns[grid.CurrentCell.ColumnIndex].Width = value;
        }

        private void SetHidden(bool row, bool hidden)
        {
            if (grid.CurrentCell == null) return;
            int[] indices = hidden ? grid.SelectedCells.Cast<DataGridViewCell>()
                .Select(cell => row ? cell.RowIndex : cell.ColumnIndex).Distinct().ToArray() :
                Enumerable.Range(0, row ? RowCount : ColumnCount)
                    .Where(i => row ? manualHiddenRows[i] : manualHiddenColumns[i]).ToArray();
            if (indices.Length == 0) return;
            if (hidden)
            {
                if ((row ? grid.Rows.Cast<DataGridViewRow>().Count(r => r.Visible) :
                    grid.Columns.Cast<DataGridViewColumn>().Count(c => c.Visible)) <= indices.Length) return;
                int alternate = Enumerable.Range(0, row ? RowCount : ColumnCount)
                    .First(i => !indices.Contains(i) && (row ? grid.Rows[i].Visible : grid.Columns[i].Visible));
                grid.CurrentCell = row ? grid[grid.CurrentCell.ColumnIndex, alternate] :
                    grid[alternate, grid.CurrentCell.RowIndex];
            }
            foreach (int index in indices)
                if (row) { manualHiddenRows[index] = hidden; grid.Rows[index].Visible = !hidden; }
                else { manualHiddenColumns[index] = hidden; grid.Columns[index].Visible = !hidden; }
            if (row) ApplyFreezeAndFilter();
            RecordChange(); MarkDirty();
        }

        private void RestoreHidden(SheetState state)
        {
            Array.Copy(state.HiddenRows, manualHiddenRows, RowCount);
            Array.Copy(state.HiddenColumns, manualHiddenColumns, ColumnCount);
            if (!state.HiddenRows.Contains(true) && !state.HiddenColumns.Contains(true))
            {
                for (int row = 0; row < RowCount; row++) if (!grid.Rows[row].Visible) grid.Rows[row].Visible = true;
                for (int column = 0; column < ColumnCount; column++) if (!grid.Columns[column].Visible) grid.Columns[column].Visible = true;
                return;
            }
            for (int row = 0; row < RowCount; row++) grid.Rows[row].Visible = true;
            for (int column = 0; column < ColumnCount; column++) grid.Columns[column].Visible = true;
            int visibleRow = Array.FindIndex(state.HiddenRows, hidden => !hidden);
            int visibleColumn = Array.FindIndex(state.HiddenColumns, hidden => !hidden);
            if (visibleRow < 0 || visibleColumn < 0) return;
            grid.CurrentCell = grid[visibleColumn, visibleRow];
            for (int row = 0; row < RowCount; row++) if (state.HiddenRows[row]) grid.Rows[row].Visible = false;
            for (int column = 0; column < ColumnCount; column++) if (state.HiddenColumns[column]) grid.Columns[column].Visible = false;
        }

        private void PaintCellExtras(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            CellExtras extras = grid[e.ColumnIndex, e.RowIndex].Tag as CellExtras;
            if (extras == null) return;
            Rectangle bounds = e.CellBounds;
            string text = Convert.ToString(grid[e.ColumnIndex, e.RowIndex].FormattedValue) ?? "";
            bool customText = extras.Rotation != 0 || extras.Shrink && !extras.Wrap;
            e.Paint(bounds, customText ? DataGridViewPaintParts.Background | DataGridViewPaintParts.Border |
                DataGridViewPaintParts.Focus | DataGridViewPaintParts.SelectionBackground : DataGridViewPaintParts.All);
            if (customText)
            {
                Font baseFont = e.CellStyle.Font ?? grid.Font;
                float size = baseFont.Size;
                if (extras.Shrink && extras.Rotation == 0)
                {
                    while (size > 6F && TextRenderer.MeasureText(text, baseFont).Width > bounds.Width - 8)
                    { size -= 0.5F; baseFont = new Font(baseFont.FontFamily, size, baseFont.Style); }
                    TextRenderer.DrawText(e.Graphics, text, baseFont, Rectangle.Inflate(bounds, -3, -2),
                        e.State.HasFlag(DataGridViewElementStates.Selected) ? e.CellStyle.SelectionForeColor :
                        e.CellStyle.ForeColor, TextFormatFlags.VerticalCenter | TextFormatFlags.Left |
                        TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
                }
                else
                {
                    var saved = e.Graphics.Save();
                    e.Graphics.SetClip(bounds);
                    e.Graphics.TranslateTransform(bounds.Left + bounds.Width / 2F, bounds.Top + bounds.Height / 2F);
                    e.Graphics.RotateTransform(-extras.Rotation);
                    using (var brush = new SolidBrush(e.State.HasFlag(DataGridViewElementStates.Selected) ?
                        e.CellStyle.SelectionForeColor : e.CellStyle.ForeColor))
                        e.Graphics.DrawString(text, baseFont, brush, -bounds.Width / 2F + 4, -baseFont.Height / 2F);
                    e.Graphics.Restore(saved);
                }
            }
            DrawEdge(e.Graphics, extras.Left, bounds.Left, bounds.Top, bounds.Left, bounds.Bottom);
            DrawEdge(e.Graphics, extras.Right, bounds.Right - 1, bounds.Top, bounds.Right - 1, bounds.Bottom);
            DrawEdge(e.Graphics, extras.Top, bounds.Left, bounds.Top, bounds.Right, bounds.Top);
            DrawEdge(e.Graphics, extras.Bottom, bounds.Left, bounds.Bottom - 1, bounds.Right, bounds.Bottom - 1);
            e.Handled = true;
        }

        private static void DrawEdge(Graphics graphics, BorderEdge edge, int x1, int y1, int x2, int y2)
        {
            if (!edge.Exists) return;
            using (var pen = new Pen(edge.Color.IsEmpty ? Color.Black : edge.Color,
                edge.Style == "thick" ? 3 : edge.Style == "medium" ? 2 : 1))
            {
                if (edge.Style == "dashed") pen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dash;
                if (edge.Style == "dotted") pen.DashStyle = System.Drawing.Drawing2D.DashStyle.Dot;
                graphics.DrawLine(pen, x1, y1, x2, y2);
                if (edge.Style == "double") graphics.DrawLine(pen, x1 + (x1 == x2 ? 2 : 0),
                    y1 + (y1 == y2 ? 2 : 0), x2 + (x1 == x2 ? 2 : 0), y2 + (y1 == y2 ? 2 : 0));
            }
        }
    }
}
